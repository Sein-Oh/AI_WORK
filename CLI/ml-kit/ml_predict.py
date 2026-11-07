# -*- coding: utf-8 -*-
"""
ml_predict.py - Zero-Office Ultra-Fast Standalone ML Inference Engine.
Supports Directory Targeting (v1/, v2/) or Specific Model/Meta File Targeting.
Handles Batch Files (.csv, .xlsx) or Inline Single/Multi Record Inferences.

DEPLOYMENT: this single file + a trained version folder (vX/) is a complete
inference package. It has no dependency on the other ml-kit files.
Requires only NumPy, Pandas, Scikit-learn, Joblib (+ openpyxl for .xlsx).

This file is also the SHARED RUNTIME of ml-kit (section C*): ml_train.py and
ml_report.py import data loading, numeric coercion, the --eval formula engine,
the Agent CLI Protocol and Windows-safe file helpers from here, so training and
inference can never drift apart.

Designed for lightweight 26B-30B LLM agents with 1-turn declarative execution.
"""

from __future__ import annotations   # type hints never touch numpy/pandas (they may be missing: R5)

import sys
import os
import re
import ast
import csv
import io
import glob
import contextlib
import json
import time
import shutil
import threading
import difflib
import functools
import argparse
import warnings
from datetime import datetime
from itertools import product
from typing import Dict, Any, List, Optional, Tuple, Callable

# R5: the core libraries are imported guarded so that a missing one still answers with 1 JSON line
# (missing_library + install_command) instead of a traceback. Every entry point calls check_core_libraries().
try:
    import numpy as np
    import pandas as pd
    import joblib
    CORE_IMPORT_ERROR: Optional[ImportError] = None
except ImportError as _core_exc:
    np = pd = joblib = None  # type: ignore[assignment]
    CORE_IMPORT_ERROR = _core_exc

warnings.filterwarnings("ignore", category=UserWarning)


# ##############################################################################
# PART C. SHARED RUNTIME (imported by ml_train.py / ml_report.py)
# ##############################################################################

# ==============================================================================
# C1. Console & Agent CLI Protocol (Clean Terminal Principle - No Emojis)
# ==============================================================================

VERSION = "2.9.71"                         # must equal metadata.version of every skills/<name>/SKILL.md (tests check it)
TOOL_NAME = "ml-kit"                      # tool folder name, used in the error log header
TOOL_DIR = os.path.dirname(os.path.abspath(__file__))
ERROR_LOG_MAX = 2 * 1024 * 1024           # errors.txt larger than this -> errors.old.txt (1 generation)
CURRENT_ARGV: Optional[List[str]] = None  # full command (script + args) for the error log; None = sys.argv


def round_pred(v: Any, digits: int = 4) -> Any:
    """예측·실측 비교값 반올림: 소수 4자리를 기본으로, 작은 값은 유효숫자 4자리까지 남김 (0.000005가 0이 되지 않게)."""
    a = np.asarray(v, dtype=float)
    with np.errstate(divide="ignore", invalid="ignore"):
        mag = np.floor(np.log10(np.abs(a)))
    dec = np.where(np.isfinite(mag), np.maximum(digits, digits - 1 - mag), digits).astype(int)
    out = np.array([round(float(x), int(d)) if np.isfinite(x) else float(x) for x, d in zip(a.ravel(), dec.ravel())]).reshape(a.shape)
    return float(out) if out.ndim == 0 else out

def setup_console() -> None:
    """Forces UTF-8 on stdout/stderr (Windows consoles and pipes default to CP949)."""
    for stream in (sys.stdout, sys.stderr):
        try:
            if stream is not None and (stream.encoding or "").lower() != "utf-8":
                stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


def json_safe(obj: Any) -> Any:
    """Converts NumPy scalars/arrays and NaN/Inf into strict-JSON-compatible values."""
    if isinstance(obj, dict):
        return {str(k): json_safe(v) for k, v in obj.items()}
    if isinstance(obj, (list, tuple)):
        return [json_safe(v) for v in obj]
    if np is None:     # missing_library answer (R5): plain Python values only
        return None if isinstance(obj, float) and obj != obj else obj
    if isinstance(obj, np.ndarray):
        return [json_safe(v) for v in obj.tolist()]
    if isinstance(obj, (bool, np.bool_)):
        return bool(obj)
    if isinstance(obj, (np.integer,)):
        return int(obj)
    if isinstance(obj, (float, np.floating)):
        f = float(obj)
        return f if np.isfinite(f) else None
    return obj


def dumps_json(obj: Any, indent: Optional[int] = None) -> str:
    return json.dumps(json_safe(obj), ensure_ascii=False, indent=indent)


# stdout holds exactly one JSON line per command: the main thread and the time limit watchdog (run_cli) both answer
# through emit_json_response, and only the first one writes.
_OUT_LOCK = threading.Lock()
_OUT_DONE: List[bool] = [False]


def emit_json_response(payload: Dict[str, Any]) -> bool:
    """Prints standard 1-line JSON response to stdout for LLM consumption. False (nothing written) when this command
    already wrote its line (the watchdog answered first, or the other way round)."""
    setup_console()
    text = dumps_json(payload) + "\n"
    with _OUT_LOCK:
        if _OUT_DONE[0]:
            return False
        sys.stdout.write(text)
        sys.stdout.flush()
        _OUT_DONE[0] = True
        return True


class InputError(ValueError):
    """User/agent input or data schema problem. Retrying with the same arguments will not help.
    extra: keys added to the error JSON (e.g. next_commands); error_type: 'needs_user_decision' when the tool stopped
    for a choice only the user can make (the directive then says to ask and run next_commands[0] only if they agree)."""

    def __init__(self, message: str = "", extra: Optional[Dict[str, Any]] = None, error_type: Optional[str] = None):
        super().__init__(message)
        self.extra = dict(extra or {})
        self.error_type = error_type


# ------------------------------------------------------------------------------
# Time limit: OpenCode's command tool stops a command at 120 s (the LLM often does not pass a longer
# timeout). Long work (ml_export input search, also inside ml_report) checks the clock before each slow step,
# saves what it finished, and stops cleanly with a CONTINUE result: running the same command again continues.
# ------------------------------------------------------------------------------
TIME_LIMIT_ENV = "ML_KIT_TIME_LIMIT"       # seconds for one command; 0 = no limit (human test runners set 0)
DEFAULT_TIME_LIMIT = 100.0                 # below OpenCode's 120 s, with room for the JSON and file writes
TIME_RESERVE = 30.0                        # kept free after the last slow step (table, workbook, Excel check, report)
RUN_START = time.monotonic()               # reset by run_cli for each command


TIME_LIMIT_ARG: List[float] = []            # --time-limit of this command (JsonArgumentParser), over the env value


def time_limit() -> float:
    if TIME_LIMIT_ARG:
        return float("inf") if TIME_LIMIT_ARG[-1] <= 0 else TIME_LIMIT_ARG[-1]
    raw = os.environ.get(TIME_LIMIT_ENV, "").strip()
    if not raw:
        return DEFAULT_TIME_LIMIT
    try:
        v = float(raw)
    except ValueError:
        return DEFAULT_TIME_LIMIT
    return float("inf") if v <= 0 else v


class Unfinished(BaseException):
    """Long work stopped at the time limit with its finished parts saved. BaseException so that no
    'except Exception' on the way (e.g. ml_report's optional workbook) swallows it; run_cli turns it into an ok
    result with continue: true and the same command in next_commands."""

    def __init__(self, progress: str):
        super().__init__(progress)
        self.progress = progress


# A finished step that cannot be saved where it belongs (<data_root>/.resume_train, <version>/export/_fits) is saved in
# a fallback place (JIT: resume_save): <tool data folder>/ml-kit-resume/<hash>, then %TEMP%/ml-kit-resume/<hash> (hash =
# the first 12 characters of the hash of the original folder, so paths stay short). The next run looks in every place
# (resume_find). When no place can be written the step is lost at a stop, so check_time does not stop with CONTINUE
# (the same command would train it again forever) and does not go past the limit either: it stops with STOP_RETRY.
RESUME_FALLBACK_DIR = "ml-kit-resume"      # a name of its own: ML_KIT_DATA_DIR may be a user folder
TEMP_RESUME_DIR = "ml-kit-resume"
# Only what resume_places/resume_save make is ever swept: <hash> folders and the saved step files
# (<key>.joblib, <key>.m<k>.joblib, <key>.m<k>.f<fold>.joblib, <key>.fit.joblib, <key>.imp.joblib, <key>.prep.joblib,
# the start markers <step>.start (step_begin), best_run.json, each also with .tmp).
_RESUME_SUB_RE = re.compile(r"[0-9a-f]{12}")
_RESUME_FILE_RE = re.compile(r"(?:[0-9a-f]{16}(?:\.m\d+(?:\.f\d+)?|\.fit|\.imp|\.prep)?\.(?:joblib|start)|best_run\.json)"
                             r"(?:\.tmp)?")
RESUME_MAX_AGE = 3 * 24 * 3600             # saved steps of runs never continued are removed after this many seconds
SAVE_FAILED: List[str] = []                # finished steps no place could hold in this command
RUN_SAVED: List[str] = []                  # finished steps saved in this command (the watchdog: progress kept)
_RESUME_PLACE: Dict[str, int] = {}         # original folder -> index of the place that worked in this command


def resume_places(primary: str) -> List[str]:
    """[original folder, data folder fallback, %TEMP% fallback] for the finished steps of `primary`."""
    import hashlib
    primary = os.path.abspath(primary)
    h = hashlib.sha1(os.path.normcase(primary).encode("utf-8")).hexdigest()[:12]
    temp = os.environ.get("TEMP") or os.environ.get("TMP") or __import__("tempfile").gettempdir()
    return [primary, os.path.join(get_data_dir(), RESUME_FALLBACK_DIR, h), os.path.join(temp, TEMP_RESUME_DIR, h)]


def _place_label(places: List[str], i: int) -> str:
    h = os.path.basename(places[i])
    return [os.path.basename(places[0]), f"데이터 폴더 {RESUME_FALLBACK_DIR}\\{h}", f"임시 폴더 {TEMP_RESUME_DIR}\\{h}"][i]


def resume_find(primary: str, name: str) -> Optional[str]:
    """The saved step `name` of `primary` in the original folder or a fallback place (None when not saved). When it
    is in several places (a file saved again elsewhere after the original folder became blocked), the newest one."""
    found = []
    for d in resume_places(primary):
        path = os.path.join(d, name)
        try:
            if os.path.isfile(path):
                found.append((os.path.getmtime(path), path))
        except OSError:
            pass
    return max(found)[1] if found else None


def resume_sweep_old() -> None:
    """At the start of a command: removes the saved steps older than RESUME_MAX_AGE in every fallback folder
    (<tool data folder>/ml-kit-resume/<hash>, %TEMP%/ml-kit-resume/<hash>), also those of data files never trained
    again, and the folders left empty. Only the <hash> folders and the file names the tool makes (_RESUME_SUB_RE,
    _RESUME_FILE_RE): anything else there (a user file, ML_KIT_DATA_DIR set to a user folder) is never touched.
    Never raises."""
    temp = os.environ.get("TEMP") or os.environ.get("TMP") or __import__("tempfile").gettempdir()
    now = time.time()
    for root in (os.path.join(get_data_dir(), RESUME_FALLBACK_DIR), os.path.join(temp, TEMP_RESUME_DIR)):
        try:
            subs = [os.path.join(root, n) for n in os.listdir(root) if _RESUME_SUB_RE.fullmatch(n)]
        except OSError:
            continue
        for sub in subs:
            try:
                if not os.path.isdir(sub) or os.path.islink(sub):
                    continue
                for name in os.listdir(sub):
                    p = os.path.join(sub, name)
                    try:
                        if (_RESUME_FILE_RE.fullmatch(name) and os.path.isfile(p) and not os.path.islink(p)
                                and now - os.path.getmtime(p) > RESUME_MAX_AGE):
                            os.remove(p)
                    except OSError:
                        pass
                if not os.listdir(sub):
                    os.rmdir(sub)
            except OSError:
                pass
        try:
            if not os.listdir(root):
                os.rmdir(root)
        except OSError:
            pass


def resume_save(primary: str, name: str, write: Callable[[str], Any], what: str) -> Optional[str]:
    """Saves a finished step (write(path) writes the file) in the first place of resume_places that works and
    checks the file is there. A fallback place is a JIT recovery (note_recovered). Returns the saved path, or None
    when no place worked (SAVE_FAILED: check_time then stops with STOP_RETRY instead of CONTINUE)."""
    places = resume_places(primary)
    key = os.path.normcase(places[0])
    errors: List[str] = []
    for i in range(_RESUME_PLACE.get(key, 0), len(places)):
        path = os.path.join(places[i], name)
        tmp = path + ".tmp"
        err = ""
        try:
            os.makedirs(places[i], exist_ok=True)
            write(tmp)
            retry_fs(lambda: os.replace(tmp, path))
        except Exception as e:
            err = f"{type(e).__name__}: {e}"[:160]
            try:
                os.remove(tmp)
            except OSError:
                pass
        if not err and os.path.isfile(path) and os.path.getsize(path) > 0:
            _RESUME_PLACE[key] = i
            RUN_SAVED.append(name)
            if i > 0:
                note_recovered(f"중간 저장 실패 ({os.path.basename(places[0])}) -> 대체 위치 {_place_label(places, i)}에 저장")
                for d in places[:i]:          # an older copy left in a place that worked before is not read again
                    try:
                        os.remove(os.path.join(d, name))
                    except OSError:
                        pass
            return path
        errors.append(f"{_place_label(places, i)}: {err or '저장한 파일을 찾지 못함'}")
    _RESUME_PLACE[key] = len(places)          # no place works: the next steps of this command do not try again
    reason = f"{what} {name}를 어디에도 저장하지 못했습니다 ({'; '.join(errors)})"
    if not SAVE_FAILED:
        record_recovery_attempt(f"중간 저장 {len(errors)}곳 시도 실패 ({os.path.basename(places[0])}, 데이터 폴더, 임시 폴더)")
        sys.stderr.write(f"[WARN] 중간 저장 실패: {reason}. 시간 제한에 닿으면 이어서 진행할 수 없어 멈춥니다\n")
    SAVE_FAILED.append(reason)
    return None


def resume_clear(primary: str, drop: Callable[[str, str], bool], fallback_only: bool = False) -> None:
    """Removes the saved steps of `primary` for which drop(name, path) is true in every place (or only the fallback
    places), then the places left empty. Never raises."""
    for i, d in enumerate(resume_places(primary)):
        if (fallback_only and i == 0) or not os.path.isdir(d):
            continue
        for name in os.listdir(d):
            p = os.path.join(d, name)
            if i > 0 and not _RESUME_FILE_RE.fullmatch(name):     # a fallback place: only the tool's own files
                continue
            try:
                if drop(name, p):
                    os.remove(p)
            except OSError:
                pass
        for folder in ([d] if i == 0 else [d, os.path.dirname(d)]):
            try:
                if not os.listdir(folder):
                    os.rmdir(folder)
            except OSError:
                pass


SAVE_FAILED_HINT = "확인할 것: 데이터 폴더와 임시 폴더(%TEMP%)의 쓰기 권한, 디스크 남은 공간, 경로 길이(260자). 명령 인자는 고치지 마세요."


def save_failed_error() -> InputError:
    err = InputError(
        "중간 결과를 저장하지 못해 나눠서 진행할 수 없습니다 (시간 제한 전에 멈춤): " + SAVE_FAILED[0]
        + ". 데이터 폴더와 임시 폴더(%TEMP%)의 쓰기 권한, 디스크 남은 공간, 경로 길이(260자)를 확인한 뒤 같은 명령을 다시 실행하십시오.")
    err.hint = SAVE_FAILED_HINT           # error_payload: this hint, not the command's own (column names, --eval)
    return err


def save_failed_warning() -> Optional[str]:
    """The result warning of a command that ended although a finished step could not be saved (None otherwise)."""
    if not SAVE_FAILED:
        return None
    return (f"중간 저장 실패: {SAVE_FAILED[0]}. 이번 실행은 시간 제한 안에 끝나 결과에는 영향이 없습니다 "
            f"(데이터 폴더와 임시 폴더의 쓰기 권한·디스크 공간을 확인하십시오)")


def time_would_pass(next_step_seconds: float) -> bool:
    """True when the next slow step (estimated next_step_seconds) would end past the time limit minus TIME_RESERVE
    (check_time would stop). Lets a caller save its partial work only when it is about to stop."""
    limit = time_limit()
    if limit == float("inf"):
        return False
    return time.monotonic() - RUN_START + max(next_step_seconds, 0.0) > limit - TIME_RESERVE


def check_time(next_step_seconds: float, progress: str) -> None:
    """Raises Unfinished when the next slow step (estimated next_step_seconds) would end past the time limit
    minus TIME_RESERVE. The caller must have saved every finished step (they are reused by the next run); when a
    step could not be saved anywhere (SAVE_FAILED) it raises the STOP_RETRY error instead."""
    if time_would_pass(next_step_seconds):
        if SAVE_FAILED:
            raise save_failed_error()
        raise Unfinished(progress)


# The tool counts the CONTINUE runs of one command (2.9.71, AGENTS.md "사용법 단순성" 5: the agent never counts):
# <data folder>/continue_state.json {command key: {n, progress, t}}. More than CONTINUE_MAX runs, or a run that saved
# nothing new and stopped at the same progress as the run before -> STOP_RETRY instead of CONTINUE. Any other end of
# the command (ok, error) clears its entry. Entries older than CONTINUE_STATE_AGE are dropped.
CONTINUE_MAX = 300
CONTINUE_STATE_AGE = 24 * 3600.0


def _continue_key(cmd: Optional[str] = None) -> str:
    import hashlib
    text = cmd or " ".join(CURRENT_ARGV if CURRENT_ARGV is not None else sys.argv)
    return hashlib.sha1(text.encode("utf-8")).hexdigest()[:16]


def _continue_state(update: Optional[Callable[[Dict[str, Any]], None]] = None) -> Dict[str, Any]:
    """Reads (and with update, rewrites) the CONTINUE counts. Never fails: no counting when the folder is not writable."""
    path = os.path.join(get_data_dir(), "continue_state.json")
    try:
        state = json.loads(read_text_file(path)) if os.path.isfile(path) else {}
        if not isinstance(state, dict):
            state = {}
    except Exception:
        state = {}
    now = time.time()
    state = {k: v for k, v in state.items() if isinstance(v, dict) and now - float(v.get("t", 0)) < CONTINUE_STATE_AGE}
    if update is not None:
        update(state)
        try:
            os.makedirs(os.path.dirname(path), exist_ok=True)
            atomic_write_text(path, json.dumps(state, ensure_ascii=False))
        except Exception:
            pass
    return state


def continue_clear() -> None:
    """The command ended (ok or error): its CONTINUE count starts again next time."""
    key = _continue_key(command_text())
    if key in _continue_state():
        _continue_state(lambda st: st.pop(key, None))


def unfinished_payload(exc: "Unfinished") -> Dict[str, Any]:
    cmd = command_text()
    key = _continue_key(cmd)
    prev = _continue_state().get(key) or {}
    n = int(prev.get("n", 0)) + 1
    stuck = prev.get("progress") == exc.progress and not RUN_SAVED
    if n > CONTINUE_MAX or stuck:
        _continue_state(lambda st: st.pop(key, None))
        why = (f"같은 명령을 {n - 1}번 이어서 실행했지만 끝나지 않았습니다" if n > CONTINUE_MAX else
               "바로 앞 실행과 같은 곳에서 멈추고 새로 저장한 것이 없어 더 이어서 해도 끝나지 않습니다")
        return {"status": "error", "error_type": "unfinished_limit", "progress": exc.progress,
                "message": f"{why} (진행: {exc.progress}). 지금까지 한 일은 저장되어 있습니다.",
                "agent_directive": "STOP_RETRY: 다시 실행하지 말고 message와 progress를 사용자에게 그대로 알리세요."}
    _continue_state(lambda st: st.__setitem__(key, {"n": n, "progress": exc.progress, "t": time.time()}))
    return {
        "status": "ok", "verified": False, "continue": True, "progress": exc.progress, "continue_runs": n,
        "next_commands": [cmd] if cmd else [],
        "agent_directive": ("CONTINUE: 아직 끝나지 않았습니다. 명령 시간 제한 때문에 나눠서 진행하며, 지금까지 한 일은 저장되어 있습니다. "
                            + ("next_commands의 명령을" if cmd else "방금 실행한 명령을")
                            + " 바꾸지 말고 그대로 다시 실행하세요 (사용자에게 묻지 않음, 몇 번째인지 세지 않음: 멈출 때는 도구가 알림). "
                            "끝나면 VERIFIED 결과가 나옵니다."),
    }


# ------------------------------------------------------------------------------
# A step that is longer than the limit by itself (R10): every run does at least one new step without a clock check (so
# running again always moves on). When that step alone does not end before the agent stops the command (120 s), the run
# leaves no JSON and saves nothing, and the same command would start it again forever. So:
#   - step_begin marks such a step before it starts (<step>.start in the resume places, {"tries", "limit"}), and
#     step_end removes the mark once the step is saved. A run that finds the mark of the step it is about to start
#     knows an earlier run was stopped inside it; after STEP_TRIES stopped tries under this limit (a busy PC may stop it
#     once) it does not start it again but stops with STOP_RETRY (step_too_long_error: [추천] the same command with
#     --time-limit STEP_LIMIT_RECOMMENDED, under the 600000 ms command timeout SKILL.md asks for).
#   - run_cli starts a watchdog: at the limit + WATCHDOG_GRACE seconds it writes the one JSON line itself (CONTINUE
#     when this run saved a step, or the stopped step may be tried again; STOP_RETRY otherwise) and ends the process.
# ------------------------------------------------------------------------------
STEP_TRIES = 2
WATCHDOG_GRACE = 15.0                      # after the limit (100 s): still before the agent's 120 s
STEP_LIMIT_RECOMMENDED = 570               # --time-limit for a 600000 ms command timeout (30 s kept for the JSON)
CURRENT_STEP: Dict[str, Any] = {}          # the step this command is doing: label, and its mark (step_begin)


def step_label(label: str) -> None:
    """Names the step this command starts now (the watchdog's progress), without a mark: a step with a clock check
    or one started after another step of this command (the mark of the first step stays until step_end)."""
    CURRENT_STEP["label"] = label


def _read_mark(path: str) -> Dict[str, Any]:
    try:
        with open(path, encoding="utf-8") as f:
            st = json.load(f)
        return st if isinstance(st, dict) else {}
    except Exception:
        return {}


def step_begin(primary: Optional[str], base: str, label: str) -> None:
    """Before a step done without a clock check as the first step of this command (R10). Raises the STOP_RETRY error
    (step_too_long_error) when STEP_TRIES earlier runs under this limit (or a larger one) were stopped inside this
    step; otherwise marks it (<base>.start, tries + 1). No mark without a limit or a resume folder. Never fails
    otherwise: a mark that cannot be written only loses the stop count."""
    CURRENT_STEP.clear()
    step_label(label)
    limit = time_limit()
    if not primary or limit == float("inf"):
        return
    name = base + ".start"
    tries = 0
    at = resume_find(primary, name)
    if at:
        st = _read_mark(at)
        try:
            if float(st.get("limit", 0)) >= limit:      # stopped under this limit (a larger one now: try again)
                tries = int(st.get("tries", 0))
        except (TypeError, ValueError):
            tries = 0
    if tries >= STEP_TRIES:
        raise step_too_long_error(label, limit, tries)
    body = dumps_json({"tries": tries + 1, "limit": limit, "label": label})
    for d in resume_places(primary):
        try:
            os.makedirs(d, exist_ok=True)
            with open(os.path.join(d, name), "w", encoding="utf-8") as f:
                f.write(body)
            CURRENT_STEP.update(primary=primary, name=name, tries=tries + 1)
            return
        except OSError:
            continue


def step_end() -> None:
    """The marked step is saved (or finished): its mark is removed (every place). Never raises."""
    name, primary = CURRENT_STEP.get("name"), CURRENT_STEP.get("primary")
    if name and primary:
        resume_clear(primary, lambda n, p: n == name)
    CURRENT_STEP.pop("name", None)


def step_too_long_error(label: str, limit: float, tries: int) -> InputError:
    """STOP_RETRY: one step does not end within the command time limit."""
    cmd = command_text(drop={"--time-limit": ""})
    head = (f"이 데이터는 한 단계({label})가 명령 시간 제한({limit:.0f}초) 안에 끝나지 않습니다 "
            f"({tries}번 시도했지만 그 단계 도중에 끊김). ")
    if cmd and limit < STEP_LIMIT_RECOMMENDED:
        err = InputError(head + "명령 실행 도구의 시간 제한(timeout)을 600000으로 주고 시간 제한을 늘린 명령으로 1번만 다시 실행하세요. "
                         f"[추천] {cmd} --time-limit {STEP_LIMIT_RECOMMENDED}", extra={"time_limit": limit})
        err.hint = "[추천] 명령은 명령 실행 도구의 시간 제한(timeout) 600000으로 실행하세요."
    else:
        err = InputError(head + "시간 제한을 늘려도 끝나지 않았습니다. 사용자에게 그대로 알리세요 (입력 열이나 행을 줄이거나, "
                         "사람이 명령 창에서 ML_KIT_TIME_LIMIT=0으로 직접 실행).", extra={"time_limit": limit})
        err.hint = "다시 실행하지 마세요."
    return err


# A result this command already saved for good (a version folder recorded in the ledger, a performance table, a report,
# a result file): mark_committed. Running the same command again would do that work again (a second version, a second
# file), so the watchdog no longer answers CONTINUE once such a result is saved: it answers ok + UNVERIFIED (saved, not
# checked) with the saved paths (2.9.69). rerun_safe=True: the same command run again reuses it (ml_report -f reuses a
# fresh best version, ml_train --best remembers the best versions it made), so CONTINUE stays right.
RUN_COMMITTED: List[Dict[str, Any]] = []


def mark_committed(what: str, rerun_safe: bool = False, **keys: Any) -> None:
    """Records a result this command saved for good (what: '학습 버전 v2', keys: version, version_dir, output_file ...)."""
    RUN_COMMITTED.append({"what": what, "rerun_safe": rerun_safe, "keys": keys})


def committed_rerun_safe() -> None:
    """The results saved so far are reused by the same command run again (e.g. ml_train --best saved best_run.json)."""
    for c in RUN_COMMITTED:
        c["rerun_safe"] = True


def committed_payload(items: List[Dict[str, Any]]) -> Dict[str, Any]:
    """ok + UNVERIFIED of a command stopped by the watchdog after it saved its result (mark_committed)."""
    what = ", ".join(c["what"] for c in items)
    keys: Dict[str, Any] = {}
    for c in items:
        keys.update(c["keys"])
    note = (f"{what} 저장됨, 확인하지 못함 (시간 제한({time_limit():.0f}초)에 닿아 저장 뒤 확인과 결과 정리 전에 명령을 끝냄)")
    return {
        "status": "ok", "verified": False, **keys,
        "saved": [{"what": c["what"], **c["keys"]} for c in items],
        "unverified": [note],
        "agent_directive": (f"UNVERIFIED: {note}. 같은 명령을 다시 실행하지 마세요 (다시 하면 같은 일을 또 해서 새 버전이나 "
                            "새 파일이 하나 더 생김). saved의 내용(버전, 경로)과 확인하지 못했다는 것을 사용자에게 그대로 알리세요."),
    }


def _watchdog_payload() -> Tuple[Dict[str, Any], int]:
    """The JSON line (and exit code) of a command still running at the limit + WATCHDOG_GRACE."""
    label = CURRENT_STEP.get("label") or "작업"
    limit = time_limit()
    if any(not c["rerun_safe"] for c in RUN_COMMITTED):
        return committed_payload(RUN_COMMITTED), 0
    if RUN_SAVED or RUN_COMMITTED:
        p = unfinished_payload(Unfinished(f"{label} 도중 시간 제한({limit:.0f}초)에 닿아 멈춤. 그 전 단계까지 저장됨"))
        return p, (0 if p["status"] == "ok" else 1)
    tries = int(CURRENT_STEP.get("tries") or 0)
    if CURRENT_STEP.get("name") and tries < STEP_TRIES:
        p = unfinished_payload(Unfinished(f"{label} 도중 시간 제한({limit:.0f}초)에 닿아 멈춤 "
                                          f"(그 단계 다시 시도 {tries}/{STEP_TRIES})"))
        return p, (0 if p["status"] == "ok" else 1)
    return error_payload(step_too_long_error(label, limit, max(tries, 1))), 1


WATCHDOG_HOLD: List[bool] = []            # the command is writing its result (version commit): wait a little
WATCHDOG_HOLD_SECONDS = 4.0


def _watchdog_fire() -> None:
    if WATCHDOG_HOLD:
        end = time.monotonic() + WATCHDOG_HOLD_SECONDS
        while not _OUT_DONE[0] and time.monotonic() < end:
            time.sleep(0.1)
    try:
        payload, code = _watchdog_payload()
        payload = _with_recovered(payload)
    except BaseException as e:                    # never leave the command without its line
        payload, code = error_payload(RuntimeError(f"시간 제한 감시 오류: {type(e).__name__}")), 1
    with _OUT_LOCK:
        if _OUT_DONE[0]:                          # the command answered first: it ends by itself
            return
    sys.stderr.write(f"\n[STATUS] 시간 제한({time_limit():.0f}초 + {WATCHDOG_GRACE:.0f}초)에 닿아 명령을 끝냄: "
                     f"{payload.get('progress') or payload.get('message')}\n")
    if code:
        write_error_log(payload)
    if emit_json_response(payload):
        try:
            sys.stderr.flush()
        except Exception:
            pass
        os._exit(code)


def start_watchdog() -> Optional[threading.Timer]:
    """run_cli: the time limit watchdog of an ml-kit command (none without a limit, or when the program is not an
    ml-kit command: in-process use such as tests)."""
    limit = time_limit()
    if limit == float("inf") or command_text() is None:
        return None
    t = threading.Timer(max(limit + WATCHDOG_GRACE - (time.monotonic() - RUN_START), 0.0), _watchdog_fire)
    t.daemon = True
    t.start()
    return t


# pip package of an import name (R5 install_command); other names install under their own name
PIP_NAMES = {"sklearn": "scikit-learn", "docx": "python-docx", "dateutil": "python-dateutil", "yaml": "PyYAML",
             "cv2": "opencv-python", "fitz": "pymupdf"}


class MissingLibraryError(InputError):
    """R5: a library the tool needs is not installed. error_payload answers missing_library + install_command."""

    def __init__(self, pip_name: str, feature: str):
        super().__init__(f"{feature}에 필요한 {pip_name}이(가) 설치되어 있지 않습니다. "
                         f"설치 명령: python -m pip install {pip_name}")
        self.pip_name = pip_name
        self.feature = feature


def missing_library_error(exc: ImportError, feature: str = "ml-kit") -> MissingLibraryError:
    """MissingLibraryError of an ImportError ('No module named sklearn.utils' -> scikit-learn)."""
    name = getattr(exc, "name", None)
    if not name:
        m = re.search(r"No module named '?([\w.]+)", str(exc))
        name = m.group(1) if m else "unknown"
    top = name.split(".")[0]
    return MissingLibraryError(PIP_NAMES.get(top, top), feature)


def require_lib(module: str, pip_name: str, feature: str) -> Any:
    """Imports module or raises MissingLibraryError (R5)."""
    import importlib
    try:
        return importlib.import_module(module)
    except ImportError:
        raise MissingLibraryError(pip_name, feature)


def exit_missing_library(exc: ImportError, feature: str = "ml-kit") -> None:
    """A core import of an entry point failed (outside run_cli): the missing_library JSON line, the error log, exit 1."""
    payload = error_payload(missing_library_error(exc, feature))
    sys.stderr.write(f"[ERROR] {payload['message']}\n")
    write_error_log(payload)
    emit_json_response(payload)
    sys.exit(1)


def check_core_libraries(feature: str = "ml-kit") -> None:
    """Entry points call this before anything else: numpy / pandas / joblib missing -> exit_missing_library."""
    if CORE_IMPORT_ERROR is not None:
        exit_missing_library(CORE_IMPORT_ERROR, feature)


class ColumnNameError(InputError):
    """A given column name matches no column (resolve_column_names): the message lists [후보] or no similar column."""


# Exceptions caused by arguments, data schema or environment (not by a bug in the tool)
_STOP_RETRY_ERRORS = (InputError, FileNotFoundError, NotADirectoryError, IsADirectoryError, PermissionError)


def error_payload(exc: BaseException, hint: str = "") -> Dict[str, Any]:
    """Builds the standard error JSON with an agent_directive the LLM can act on deterministically."""
    msg = str(exc) or type(exc).__name__
    if isinstance(exc, KeyboardInterrupt):
        return {
            "status": "error",
            "error_type": "interrupted",
            "message": "사용자 또는 상위 프로세스에 의해 실행이 중단되었습니다. 임시 산출물은 롤백되었습니다.",
            "agent_directive": "STOP_RETRY: 실행이 중단되었습니다. 다시 실행하기 전에 사용자에게 물으세요.",
        }
    if isinstance(exc, ImportError) and not isinstance(exc, MissingLibraryError):
        exc = missing_library_error(exc, "이 기능")
    if isinstance(exc, MissingLibraryError):
        return {
            "status": "error",
            "error_type": "missing_library",
            "message": str(exc),
            "missing_package": exc.pip_name,
            "install_command": f"python -m pip install {exc.pip_name}",
            "agent_directive": (f"STOP_RETRY: {exc.feature}에 필요한 {exc.pip_name}이(가) 설치되어 있지 않습니다. "
                                "install_command를 사용자에게 전달하세요. 직접 설치하지 마세요."),
        }
    if isinstance(exc, InputError) and getattr(exc, "error_type", None) == "needs_user_decision":
        return {
            "status": "error",
            "error_type": "needs_user_decision",
            "message": msg,
            **{k: v for k, v in exc.extra.items() if k not in ("status", "message", "agent_directive", "error_type")},
            "agent_directive": (
                "STOP_RETRY: 사용자가 정할 일이 있어 멈췄습니다. message를 그대로 알리고 사용자에게 물으세요. 스스로 실행하지 말고, "
                "사용자가 그렇게 하라고 하면 next_commands[0].command를 글자 그대로 1번 실행하세요."),
        }
    if isinstance(exc, _STOP_RETRY_ERRORS):
        hint = getattr(exc, "hint", None) or hint
        return {
            "status": "error",
            "error_type": "input",
            "message": msg,
            **{k: v for k, v in getattr(exc, "extra", {}).items()
               if k not in ("status", "message", "agent_directive", "error_type")},
            "agent_directive": (
                "STOP_RETRY: 입력이나 데이터 문제로 멈췄습니다. message에 [추천]이 있으면 [추천] 뒤의 전체 명령을 그대로 1번만 실행하세요. "
                "[후보]가 있으면 다시 실행하지 말고 후보를 사용자에게 보여 주고, 사용자가 고르면 같은 번호의 confirm_commands를 1번 실행하세요 "
                "(confirm_commands가 없으면 고른 값으로 1번 실행). 둘 다 없으면 다시 실행하지 말고 message를 그대로 알리세요."
                + (f" {hint}" if hint else "")
            ),
        }
    return {
        "status": "error",
        "error_type": "internal",
        "message": f"{type(exc).__name__}: {msg}",
        "agent_directive": "ERROR: 도구 내부 오류입니다. 다시 실행하지 말고 message를 그대로 알리세요.",
    }


_RECOVERED: List[str] = []


def note_recovered(msg: str) -> None:
    """Records a JIT recovery that changed the result (e.g. saved under a new name); reported as JSON `recovered`."""
    msg = str(msg)[:200]
    if msg in _RECOVERED:          # the same recovery for every target of one run is reported once
        return
    sys.stderr.write(f"[WARN] {msg}\n")
    _RECOVERED.append(msg)


def _with_recovered(payload: Dict[str, Any]) -> Dict[str, Any]:
    if _RECOVERED and "recovered" not in payload:
        payload = {**payload, "recovered": list(_RECOVERED)}
    chosen = globals().get("MODEL_CHOSEN")
    if chosen and payload.get("status") == "ok" and "model_chosen" not in payload:
        payload = {**payload, "model_chosen": chosen[-1]}      # -m <데이터 파일>: 도구가 고른 버전 (2.9.71)
    return payload


_RECOVERY_ATTEMPTS: List[str] = []


def record_recovery_attempt(msg: str) -> None:
    """Records a JIT recovery that hit its limit (R4); written as the '복구 시도:' line of the error log (R6)."""
    _RECOVERY_ATTEMPTS.append(str(msg)[:200])


# ------------------------------------------------------------------------------
# Error log (R6): failed runs are appended to <data>/logs/errors.txt for later analysis
# by people and development agents. The path is never put in stdout JSON.
# ------------------------------------------------------------------------------

def engenius_home() -> Optional[str]:
    """Fixed folder for tool data when the tool files themselves run from the En-Genius tray's temporary folder
    (%TEMP%\\engenius-skill\\<id>, deleted after each run, so <tool folder>/data would not survive). The deployed skill
    (docs/engenius) runs the tools in C:\\CLI instead and shares OpenCode's data folder."""
    if not os.environ.get("ENGENIUS_OUT") or "engenius-skill" not in TOOL_DIR.replace("\\", "/").lower().split("/"):
        return None
    home = os.environ.get("USERPROFILE") or os.path.expanduser("~")
    return os.path.join(home, "Documents", "En-Genius", "CLI-data", "ml-kit")


def get_data_dir() -> str:
    """Tool data folder: ML_KIT_DATA_DIR > En-Genius fixed folder/data > <tool folder>/data (not created here)."""
    return os.path.abspath(os.environ.get("ML_KIT_DATA_DIR")
                           or os.path.join(engenius_home() or TOOL_DIR, "data"))


def error_log_path() -> str:
    return os.path.join(get_data_dir(), "logs", "errors.txt")


# The shell that runs the commands the tools print (next_commands, [추천] full commands) as given.
# Same order as the OpenCode shell tool: pwsh (PowerShell 7) when on PATH, else Windows PowerShell 5.1.
# ML_KIT_SHELL=pwsh|powershell|cmd|bash forces one (2.9.68). Both PowerShells are exact, cmd and bash best effort.
SHELL_ENV = "ML_KIT_SHELL"
SHELLS = ("pwsh", "powershell", "cmd", "bash")
_FOUND_SHELL: List[str] = []
# Characters PowerShell reads as a single quote (doubled inside a single-quoted string)
_PS_SINGLE_QUOTES = "'\u2018\u2019\u201a\u201b"
_PLAIN_PUNCT = "_-.:\\/,=+@!"
_PLAIN_INT = re.compile(r"-?(?:0|[1-9]\d{0,8})")
_PLAIN_OPTION = re.compile(r"--?[A-Za-z][A-Za-z0-9-]*")


def command_shell() -> str:
    """The shell for command strings: ML_KIT_SHELL > pwsh when on PATH > powershell (Windows PowerShell 5.1)."""
    forced = os.environ.get(SHELL_ENV, "").strip().lower()
    if forced in SHELLS:
        return forced
    if not _FOUND_SHELL:
        _FOUND_SHELL.append("pwsh" if shutil.which("pwsh") else "powershell")
    return _FOUND_SHELL[0]


def _net_space(ch: str) -> bool:
    """.NET char.IsWhiteSpace (what makes Windows PowerShell 5.1 wrap an argument in double quotes)."""
    import unicodedata
    return ch in "\t\n\v\f\r\x85\xa0" or unicodedata.category(ch) in ("Zs", "Zl", "Zp")


def _plain_arg(s: str, shell: str) -> bool:
    """A value every shell passes as is without quotes: letters and digits (any script) and _-.:\\/,=+@! only.
    Quoted anyway: shapes PowerShell may turn into a number (007, 1e3, 0x10, 1.50, 2-8), a leading - that is not an
    option (-a.b splits in two), a leading = : + @ (array, splatting) !, a leading, trailing or doubled comma (array)
    and, for bash, a backslash."""
    if not s or s == "--":
        return False
    for ch in s:
        if not (ch.isalnum() if not ch.isascii() else (ch.isalnum() or ch in _PLAIN_PUNCT)):
            return False
    if s[0] in "=:+,@!" or s[-1] == "," or ",," in s:
        return False
    if shell == "bash":
        return "\\" not in s
    if s[0] == "-" and not (_PLAIN_OPTION.fullmatch(s) or _PLAIN_INT.fullmatch(s)):
        return False
    if (s[0].isdigit() or s[0] == "." and s[1:2].isdigit()) and not _PLAIN_INT.fullmatch(s):
        return False
    return True


def _argv_escape(s: str, wrapped: bool, quote: str = '\\"') -> str:
    """Escapes s so the Windows command-line parser (CommandLineToArgvW, Python argv) gets it back: backslashes before
    a " are doubled and the " becomes quote; in a wrapped value the trailing backslashes are doubled too (before the closing ")."""
    out: List[str] = []
    slashes = 0
    for ch in s:
        if ch == "\\":
            slashes += 1
            continue
        if ch == '"':
            out.append("\\" * (2 * slashes) + quote)
        else:
            out.append("\\" * slashes + ch)
        slashes = 0
    out.append("\\" * (2 * slashes if wrapped else slashes))
    return "".join(out)


def _ps_single(s: str) -> str:
    return "'" + "".join(ch * 2 if ch in _PS_SINGLE_QUOTES else ch for ch in s) + "'"


def quote_arg(value: Any, shell: Optional[str] = None) -> str:
    """Quotes one argument of a command string for the shell that runs it (command_shell) (2.9.68).

    - A value without shell special characters (_plain_arg) stays as is.
    - PowerShell: single quotes '...' (a ' inside is '') keep $, backtick, {}, @, #, ; and ( literal.
      Windows PowerShell 5.1 hands the value to python wrapped in "..." when it has a space but leaves a " inside as is,
      so " -> \\" (backslashes before it doubled), trailing backslashes of a wrapped value doubled and an empty value
      '""' (5.1 drops an empty argument) are done here. PowerShell 7.3+ escapes by itself, so nothing is done for it.
    - cmd: "..." (a " inside is "", % is cut as "^%" so %NAME% is not expanded). bash: '...' (a ' inside is '\\'')."""
    s = str(value)
    shell = shell or command_shell()
    if _plain_arg(s, shell):
        return s
    if shell == "pwsh":
        return _ps_single(s)
    if shell == "bash":
        return "'" + s.replace("'", "'\\''") + "'"
    if shell == "cmd":
        return '"' + _argv_escape(s, True, '""').replace("%", '"^%"') + '"'
    if not s:
        return _ps_single('""')
    return _ps_single(_argv_escape(s, any(_net_space(ch) for ch in s)))


def write_error_log(payload: Dict[str, Any], trace: str = "", argv: Optional[List[str]] = None) -> None:
    """Appends one failed run to errors.txt. Never raises: logging must not break the 1-line JSON contract.
    Only paths and arguments are recorded (no document contents)."""
    try:
        import platform
        path = error_log_path()
        os.makedirs(os.path.dirname(path), exist_ok=True)
        try:  # rotation failure must not lose the entry
            if os.path.exists(path) and os.path.getsize(path) > ERROR_LOG_MAX:
                os.replace(path, path[:-4] + ".old.txt")
        except Exception:
            pass
        if argv is None:
            argv = CURRENT_ARGV if CURRENT_ARGV is not None else [os.path.basename(sys.argv[0] or "")] + sys.argv[1:]
        lines = [f"===== {datetime.now().isoformat(timespec='seconds')} | {TOOL_NAME} {VERSION} | "
                 f"Python {platform.python_version()} | {platform.platform()}",
                 f"명령: {' '.join(quote_arg(a) for a in argv)}",
                 f"결과: {payload.get('error_type')} | {payload.get('message')}"]
        if _RECOVERY_ATTEMPTS:
            lines.append(f"복구 시도: {'; '.join(_RECOVERY_ATTEMPTS)}")
        if trace:
            lines.append("traceback:\n" + trace.rstrip())
        with open(path, "a", encoding="utf-8") as f:
            f.write("\n".join(lines) + "\n\n")
    except Exception:
        pass


def run_cli(execute: Callable[[], Dict[str, Any]], hint: str = "") -> None:
    """Runs a tool body, guaranteeing exactly one JSON line on stdout and exit code 0/1/130.
    Failures (exit 1) are appended to <data>/logs/errors.txt; interruptions (130) are not."""
    global RUN_START
    RUN_START = time.monotonic()
    SAVE_FAILED.clear()
    RUN_SAVED.clear()
    RUN_COMMITTED.clear()
    MODEL_CHOSEN.clear()
    CURRENT_STEP.clear()
    WATCHDOG_HOLD.clear()
    _RESUME_PLACE.clear()
    _OUT_DONE[0] = False
    watchdog = start_watchdog()
    try:
        resume_sweep_old()
        result = execute()
    except Unfinished as u:
        sys.stderr.write(f"[STATUS] 시간 제한으로 여기서 멈춤 (저장됨, 같은 명령으로 이어서): {u.progress}\n")
        payload = _with_recovered(unfinished_payload(u))
        if payload["status"] != "ok":
            write_error_log(payload)
        emit_json_response(payload)
        if payload["status"] != "ok":
            sys.exit(1)
        return
    except KeyboardInterrupt as e:
        sys.stderr.write("\n[ERROR] 실행 중단 (KeyboardInterrupt)\n")
        emit_json_response(_with_recovered(error_payload(e)))
        sys.exit(130)
    except Exception as e:
        import traceback
        trace = traceback.format_exc()
        sys.stderr.write(f"\n[ERROR] {e}\n")
        payload = _with_recovered(error_payload(e, hint))
        continue_clear()
        write_error_log(payload, trace if payload.get("error_type") == "internal" else "")
        emit_json_response(payload)
        sys.exit(1)
    finally:
        if watchdog is not None:
            watchdog.cancel()
        step_end()       # ended by itself (also with an error): only a run stopped from outside leaves its mark
    continue_clear()
    emit_json_response(_with_recovered(result))


# Option names a light LLM invents for an option that exists (2.9.69): an argument error gives [추천] the full command
# with the real option (the first of the list this command has). None is read as the real option (2.9.71: one way to
# write each option, AGENTS.md "사용법 단순성" 1; '--target' for --targets is an argument error with that [추천]).
SAME_MEANING: Dict[str, List[str]] = {
    "--model-dir": ["--model"], "--model-path": ["--model"], "--models": ["--model"], "--dir": ["--model"],
    "--version-dir": ["--model"], "--model-folder": ["--model"],
    "--input-file": ["--file"], "--data": ["--file"], "--data-file": ["--file"], "--path": ["--file"],
    "--input": ["--file", "--inputs"], "--inputs": ["--input"], "--query": ["--input"],
    "--target": ["--targets"], "--target-col": ["--targets"], "--target-cols": ["--targets"],
    "--target-column": ["--targets"], "--y": ["--targets"],
    "--feature": ["--features"], "--columns": ["--features"], "--feature-cols": ["--features"], "--x": ["--features"],
    "--output": ["--save"], "--out": ["--save"], "--output-file": ["--save"], "-o": ["--save"],
    "--header": ["--header-row"], "--unit": ["--unit-row"], "--exclude": ["--exclude-where"],
    "--positive": ["--positive-class"], "--time": ["--time-limit"], "--description": ["--desc"],
}
# Options with nothing to do in ml-kit: dropped in the [추천] command (value: how many arguments follow)
DROP_OPTIONS = {"--json": 0, "--encoding": 1, "--verbose": 0, "--quiet": 0}
# Never given as a lookalike, even as [후보]: replaces files or retrains everything
NO_FUZZY_OPTIONS = {"--overwrite", "--best", "--help"}
_OPTION_TOKEN = re.compile(r"--?[A-Za-z][A-Za-z0-9_-]*")


# 스크립트 -> 그 명령을 설명하는 작업별 스킬 (skills/<이름>/SKILL.md, 2.9.71에 스킬 하나를 넷으로 나눔)
SKILL_OF_SCRIPT = {"ml_train.py": "ml-train", "ml_predict.py": "ml-predict", "ml_report.py": "ml-report",
                   "ml_export.py": "ml-report", "ml_map.py": "ml-map"}


class _ProbeError(Exception):
    """A probe parse (JsonArgumentParser._parses) failed: the [추천] command would not parse either."""


class JsonArgumentParser(argparse.ArgumentParser):
    """argparse that reports usage errors as Agent CLI Protocol JSON instead of bare text. Every ml-kit command takes
    --time-limit SEC (over ML_KIT_TIME_LIMIT, 0 = none): the [추천] command of a step longer than the limit
    (step_too_long_error). Abbreviations are not read (allow_abbrev=False: '--data x' was --data-row); the same-meaning
    names (SAME_MEANING) are not read either: an argument error with [추천] the fixed full command."""

    def __init__(self, *args: Any, **kwargs: Any) -> None:
        kwargs.setdefault("allow_abbrev", False)
        super().__init__(*args, **kwargs)
        self._probing = False
        self.add_argument("--time-limit", dest="time_limit", type=float, default=argparse.SUPPRESS,
                          help=f"명령 1번의 시간 제한(초, 기본 {TIME_LIMIT_ENV} 또는 {DEFAULT_TIME_LIMIT:.0f}, 0 = 없음). "
                               f"한 단계가 시간 제한보다 길다는 STOP_RETRY의 [추천] 명령에만 씀")

    def _options(self) -> set:
        return {o for a in self._actions for o in a.option_strings}

    def _same_meaning(self, flag: str) -> Optional[str]:
        opts = self._options()
        return next((o for o in SAME_MEANING.get(flag, []) if o in opts), None)

    def parse_known_args(self, args=None, namespace=None):  # type: ignore[override]
        given = [str(a) for a in (sys.argv[1:] if args is None else args)]
        if not self._probing:
            self._given = given                       # what _hint fixes
        read = given
        ns, extra = super().parse_known_args(read, namespace)
        if getattr(ns, "time_limit", None) is not None and not self._probing:
            TIME_LIMIT_ARG[:] = [float(ns.time_limit)]
        return ns, extra

    def _parses(self, args: List[str]) -> bool:
        """True when args parse without an argument error (the [추천] command is checked before it is given)."""
        self._probing = True
        try:
            with contextlib.redirect_stderr(io.StringIO()):
                self.parse_args(args)
            return True
        except (_ProbeError, SystemExit):
            return False
        finally:
            self._probing = False

    def _hint(self) -> str:
        """[추천] the full command with same-meaning names replaced and DROP_OPTIONS removed (when it parses),
        else [후보] lookalike option names (never NO_FUZZY_OPTIONS), else nothing."""
        import difflib
        given = list(getattr(self, "_given", None) or [str(a) for a in sys.argv[1:]])
        if "-h" in given or "--help" in given:
            return ""
        opts = self._options()
        fixed: List[str] = []
        unknown: List[str] = []
        changes: List[str] = []
        i = 0
        while i < len(given):
            a = given[i]
            key, eq, val = a.partition("=")
            if _OPTION_TOKEN.fullmatch(key) and key not in opts:
                if key in DROP_OPTIONS:
                    changes.append(f"{key} 뺌")
                    takes_value = not eq and DROP_OPTIONS[key] and i + 1 < len(given) and not given[i + 1].startswith("-")
                    i += 2 if takes_value else 1
                    continue
                real = self._same_meaning(key)
                if real:
                    changes.append(f"{key} -> {real}")
                    fixed.append(real + eq + val)
                    i += 1
                    continue
                unknown.append(key)
            fixed.append(a)
            i += 1
        if changes and not unknown and self._parses(fixed):
            cmd = command_text(argv=[sys.argv[0]] + fixed)
            if cmd:
                return f" 같은 뜻의 옵션으로 고친 명령 ({', '.join(changes)}). [추천] {cmd}"
        hints = []
        for key in unknown:
            near = difflib.get_close_matches(key, sorted(o for o in opts if o.startswith("--")), n=3, cutoff=0.6)
            if near and near[0] in NO_FUZZY_OPTIONS:      # '--overwrit': never guessed (replaces files, retrains all)
                hints.append(f"'{key}' 옵션은 없습니다. 무엇을 하려던 것인지 사용자에게 확인하세요.")
            elif near:
                hints.append(f"'{key}' 옵션은 없습니다. [후보] {', '.join(o for o in near if o not in NO_FUZZY_OPTIONS)} (뜻이 같은지 사용자에게 확인)")
        return (" " + " ".join(hints)) if hints else ""

    def print_help(self, file=None) -> None:  # type: ignore[override]
        """R7: help text goes to stderr, stdout gets exactly 1 JSON line (argparse then exits 0). A help text that cannot
        be built (a bare '%' in an option help) still gives the JSON line."""
        try:
            super().print_help(sys.stderr)
        except Exception as e:
            sys.stderr.write(f"[WARN] 도움말을 만들지 못했습니다 ({type(e).__name__}). SKILL.md를 보세요.\n")
        name = SKILL_OF_SCRIPT.get(os.path.basename(sys.argv[0] or ""), "ml-train")
        skill = os.path.join(TOOL_DIR, "skills", name, "SKILL.md")
        emit_json_response({
            "status": "ok", "tool": "ml-kit", "command": "help", "verified": True, "skill_file": skill,
            "agent_directive": f"VERIFIED: 도움말을 stderr에 출력했습니다. {name} 스킬(SKILL.md)의 명령 템플릿을 그대로 사용하세요. "
                               "도구 코드(.py)는 읽지 마세요. 방법을 모르겠으면 사용자에게 물어보세요.",
        })

    def error(self, message: str) -> None:  # type: ignore[override]
        if self._probing:
            raise _ProbeError(message)
        self.print_usage(sys.stderr)
        try:
            hint = self._hint()
        except Exception:                 # a hint never breaks the one JSON line
            hint = ""
        payload = {
            "status": "error",
            "error_type": "input",
            "message": f"CLI 인자 오류: {message}.{hint}",
            "agent_directive": ("STOP_RETRY: 명령 인자가 잘못되었습니다. message에 [추천]이 있으면 [추천] 뒤의 명령을 그대로 1번만 "
                                "실행하세요. 없으면 다시 실행하지 말고 message를 그대로 알리세요 ([후보]가 있으면 사용자에게 묻기, "
                                "옵션 이름은 SKILL.md 기준)."),
        }
        write_error_log(payload)
        emit_json_response(payload)
        sys.exit(2)


def sanitize_filename(name: str) -> str:
    r"""
    Replaces characters forbidden in Windows filenames (\ / : * ? " < > |).
    Guarantees targets like '유속(m/s)' or '비용(원/kg)' save safely on Windows filesystem.
    """
    return re.sub(r'[\\/*?:"<>|]', "_", str(name))


def normalize_version(v: Optional[str]) -> Optional[str]:
    """'2', 'V2', ' v2 ' -> 'v2'."""
    if v is None:
        return None
    s = str(v).strip().lower()
    if not s:
        return None
    return s if s.startswith("v") else f"v{s}"


def warn_long_path(path: str) -> None:
    """Windows without LongPathsEnabled fails on paths >= 260 chars."""
    if os.name == "nt" and len(os.path.abspath(path)) > 240:
        sys.stderr.write(
            f"[WARN] 경로 길이가 {len(os.path.abspath(path))}자로 Windows 260자 제한에 근접합니다. "
            f"저장 실패 시 데이터 파일을 더 짧은 경로(예: C:\\work\\)로 옮기십시오.\n"
        )


# ==============================================================================
# C2. Encoding-tolerant File Reading (Windows Notepad / PowerShell / Excel outputs)
# ==============================================================================

def read_text_file(path: str) -> str:
    """Reads text written by any common Windows tool (UTF-8 +/- BOM, UTF-16 BOM, CP949/EUC-KR)."""
    with open(path, "rb") as f:
        raw = f.read()
    if raw.startswith(b"\xff\xfe") or raw.startswith(b"\xfe\xff"):
        return raw.decode("utf-16")
    for enc in ("utf-8-sig", "cp949", "euc-kr"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    raise InputError(f"텍스트 파일 인코딩 인식 실패 (UTF-8/UTF-16/CP949 아님): {path}")


def load_json_file(path: str) -> Any:
    try:
        return json.loads(read_text_file(path))
    except json.JSONDecodeError as e:
        raise InputError(f"JSON 파일 파싱 실패 ({path}): {e}")


# ==============================================================================
# C3. Column Name Resolution (unit-suffix aliases & typo suggestions)
# ==============================================================================

_UNIT_SUFFIX_RE = re.compile(r"\s*[\(\[][^\(\)\[\]]*[\)\]]\s*$")


def strip_unit(name: str) -> str:
    """'반응온도(°C)' -> '반응온도', '유속 [m/s]' -> '유속'."""
    return _UNIT_SUFFIX_RE.sub("", str(name)).strip()


def _alias_key(name: str) -> str:
    return re.sub(r"\s+", "", strip_unit(name)).lower()


def build_alias_map(columns: List[str]) -> Dict[str, str]:
    """Maps unit-less alias keys to their unique full column name (ambiguous aliases are dropped)."""
    groups: Dict[str, List[str]] = {}
    for c in columns:
        groups.setdefault(_alias_key(c), []).append(c)
    return {k: v[0] for k, v in groups.items() if len(v) == 1 and k}


def unit_of(name: str) -> str:
    """'압력(bar)' -> 'bar', '반응온도 [°C]' -> '°C', '압력' -> ''."""
    m = _UNIT_SUFFIX_RE.search(str(name))
    return m.group(0).strip()[1:-1].strip() if m else ""


def other_unit(given: str, name: str) -> bool:
    """given and name both carry a unit and the units differ ('압력(MPa)' / '압력(bar)', also '압력(MPa)' /
    '반응압력(bar)'): renaming given to name would read a MPa value as bar, whatever the names are. '°C', '℃' and 'C'
    (the degree sign) and spaces count as the same unit; case does not ('mg' / 'Mg')."""
    def norm(u: str) -> str:
        return re.sub(r"\s+", "", u).replace("\u2103", "C").replace("\u00b0", "")
    ug, un = unit_of(given), unit_of(name)
    return bool(ug and un) and norm(ug) != norm(un)


def unit_hint(given: str, name: str, where: str) -> str:
    """Why a different-unit name is not renamed to the column/model name (no [추천]: the value must be converted)."""
    ug, un = unit_of(given), unit_of(name)
    same = "" if _alias_key(given) == _alias_key(name) else f" '{given}'이(가) '{name}'과(와) 같은 값이면"
    return (f"{where} '{given}'은(는) '{name}'과(와) 단위가 다릅니다 ({ug} / {un}). 이름만 바꾸면 {ug} 값이 {un} 값으로 "
            f"쓰여 결과가 틀립니다.{same} 값을 {un} 단위로 환산해 '{name}'(으)로 주십시오. 환산 값은 사용자에게 확인하십시오")


def unit_marked(cands: List[str], name: str, quote: bool = False) -> List[str]:
    """[후보] names (repr-quoted with quote), each candidate whose unit differs from name (other_unit) marked
    "(단위 다름: 값 환산 필요)": picking it needs the value converted, not only the name changed."""
    return [(repr(c) if quote else c) + ("(단위 다름: 값 환산 필요)" if other_unit(c, name) else "") for c in cands]


def check_same_column(given: List[str], resolved: List[str], option: str,
                      values: Optional[Dict[str, Any]] = None) -> None:
    """Two given names that resolve to one column ('온도' auto-matched to '온도(°C)' next to '온도(°C)'): one would
    silently win. STOP_RETRY with the given names (and values) as [후보]."""
    groups: Dict[str, List[str]] = {}
    for g, r in zip(given, resolved):
        groups.setdefault(r, []).append(g)
    for col, same in groups.items():
        if len(same) > 1:
            cands = [f"{g}={values[g]}" if values is not None and g in values else g for g in same]
            raise InputError(
                f"{option}의 이름 {', '.join(repr(g) for g in same)}은(는) 같은 열 '{col}'을(를) 가리킵니다 (단위 없는 "
                f"이름은 단위 붙은 이름으로 자동 매칭). 값 하나만 주십시오. 사용자에게 어느 값인지 물으십시오: "
                f"[후보] {', '.join(cands)}")


def _same_unit_name(given: str, column: str) -> bool:
    """given is column written otherwise ('수율( 1차)', '생산량 ( t )', '압력 [bar]', '온도(C)' / '온도(℃)' for '수율(1차)',
    '생산량(t)', '압력(bar)', '온도(°C)'): the same name without unit (case and spaces aside) and the same unit when
    spaces are dropped and '°C', '℃' and 'C' are one unit (other_unit). Only for names that both carry a unit; the case
    of a unit counts ('mg' / 'Mg')."""
    return bool(unit_of(given) and unit_of(column)) and _alias_key(given) == _alias_key(column) and \
        not other_unit(given, column)


MESSAGE_LIST_MAX = 40     # names listed in one error message ('[사용 가능한 컬럼 목록]'); the rest are counted


def name_list(names: List[Any], limit: int = MESSAGE_LIST_MAX) -> str:
    """A name list for a message, at most limit names and '외 N개 (전체 M개)' (R2: a wide sheet of 500 columns must not
    fill the LLM context with one error message)."""
    names = list(names)
    if len(names) <= limit:
        return str(names)
    return f"{names[:limit]} 외 {len(names) - limit}개 (전체 {len(names)}개)"


def suggest_columns(name: str, columns: List[str], n: int = 3, bare: bool = True) -> List[str]:
    """Similar columns: the columns of the same name without unit first (all of them: '수율' for '수율(1차)' and
    '수율(최종)', whatever their length), then close spellings, up to n (more only for same-name columns). bare: a name
    given without unit is also compared with the column names without unit ('온두' / '온도(°C)'). Close spellings
    are never a sure match: they are [후보] for the user, never [추천]."""
    key = _alias_key(name)
    cols = [str(c) for c in columns]
    hits = [c for c in cols if key and _alias_key(c) == key]
    n = max(n, len(hits))
    for c in difflib.get_close_matches(str(name), cols, n=n, cutoff=0.5):
        if c not in hits:
            hits.append(c)
    if bare and strip_unit(name) == str(name).strip():
        by_bare: Dict[str, List[str]] = {}
        for c in cols:
            by_bare.setdefault(strip_unit(c), []).append(c)
        for b in difflib.get_close_matches(str(name).strip(), list(by_bare), n=n, cutoff=0.5):
            hits.extend(c for c in by_bare[b] if c not in hits)
    return hits[:n]


def _match_names(names: List[str], columns: List[str]) -> List[Optional[str]]:
    """The column each name is (None when it is none): the exact name, the same name and unit written otherwise
    ('수율( 1차)', '온도(C)' for '온도(°C)': _same_unit_name), or a unique name written WITHOUT a unit
    ('생산효율' -> '생산효율(%)'; a different unit is another column). Only these deterministic equivalences match."""
    col_list = [str(c) for c in columns]
    col_set = set(col_list)
    alias_map = build_alias_map(col_list)
    out: List[Optional[str]] = []
    for n in names:
        same = [] if n in col_set else [c for c in col_list if _same_unit_name(n, c)]
        if n in col_set:
            out.append(n)
        elif len(same) == 1:
            out.append(same[0])
        elif strip_unit(n) == n.strip() and _alias_key(n) in alias_map:
            out.append(alias_map[_alias_key(n)])
        else:
            out.append(None)
    return out


def known_columns(names: List[str], columns: List[str]) -> List[str]:
    """The columns the names match (resolve_column_names without errors or messages): the columns another option
    takes, passed as exclude= when resolving this one (-x next to -y)."""
    return [c for c in _match_names(names, columns) if c is not None]


_NAME_PART = re.compile(r"(\s*[,;|=]\s*)")


def command_replacing(old: str, new: str) -> Optional[str]:
    """The running command with the name `old` replaced by `new` (a full command for [후보] next_commands, 2.9.71):
    an argument equal to old, or one part of a list or a 'k=v' argument ('-x "온도,압력"', '-i "온도=1,압력=2"').
    None when the running program is not an ml-kit command or old is in no argument."""
    argv = list(CURRENT_ARGV if CURRENT_ARGV is not None else sys.argv)
    if not argv:
        return None
    found = False
    for i in range(1, len(argv)):
        a = str(argv[i])
        if a.startswith("-") and i > 1:
            continue
        parts = _NAME_PART.split(a)
        if old in [p.strip() for p in parts]:
            argv[i] = "".join(new if p.strip() == old else p for p in parts)
            found = True
            break
    return command_text(argv=argv) if found else None


def candidate_commands(name: str, cands: List[str]) -> List[Dict[str, Any]]:
    """confirm_commands for [후보] names (run only the one the user picks): the running command with `name` replaced by
    each candidate (all or none). destructive when the command replaces something (--overwrite)."""
    cmds = [command_replacing(name, c) for c in cands]
    if not cmds or any(c is None for c in cmds):
        return []
    risky = "--overwrite" in (CURRENT_ARGV if CURRENT_ARGV is not None else sys.argv)
    return [{"purpose": f"'{name}' 대신 '{c}' (사용자가 고른 경우에만)", "command": cmd, **({"destructive": True} if risky else {})}
            for c, cmd in zip(cands, cmds)]


def resolve_column_names(names: List[str], columns: List[str], kind: str,
                         layout: Optional[Dict[str, Any]] = None, values: bool = False,
                         exclude: Optional[List[str]] = None) -> Tuple[List[str], Dict[str, str]]:
    """
    Resolves requested names against real columns: the exact name, the same name and unit written otherwise
    ('온도(C)' -> '온도(°C)') or a unique unit-less alias ('생산효율' -> '생산효율(%)'), each matched with an [INFO]
    note. Anything else is no sure match and raises InputError with [후보] (the user picks, even one candidate):
    a close spelling may be another quantity ('속도' / '가속도', '수율' / '함수율'), never an automatic [추천].
    layout (df.attrs of load_data_file): when the missing names are all in another top row (two-level or merged
    header, title row), the suggestion is that row (header_layout_hint), with its unit row when the row below it is one.
    values: the names carry values (-i '압력(MPa)=0.3'): one column of the same name with another unit ('압력(bar)')
    is no rename, the value must be converted (unit_hint).
    A column found for another given name is not suggested for a missing one (except the same name in another unit).
    exclude: columns this command already uses elsewhere (ml_map -i next to -y / --find / --vary, -x next to -y,
    the other names of ml_export --keep): matched when named, never listed as a candidate.
    """
    col_list = [str(c) for c in columns]
    resolved: List[str] = []
    applied: Dict[str, str] = {}
    missing: List[str] = []
    for n, hit in zip(names, _match_names(names, col_list)):
        if hit is None:
            missing.append(n)
            continue
        resolved.append(hit)
        if hit != n:
            applied[n] = hit
    if missing:
        hint = header_layout_hint(layout, missing)
        if hint:
            raise InputError(f"지정하신 {kind} 컬럼 {missing}을(를) 머리글에서 찾을 수 없습니다. {hint}"
                             f"\n[지금 머리글 컬럼 목록]: {col_list}")
        hints = []
        picks: List[Dict[str, Any]] = []
        taken = set(resolved)
        banned = set(exclude or [])
        for m in missing:
            # a column another given name already takes is no candidate ('습도' next to a given '속도'), unless it
            # is the same name in another unit ('압력(MPa)' next to a given '압력(bar)': unit_hint)
            free = [c for c in col_list if c not in taken or _alias_key(c) == _alias_key(m)]
            s = suggest_columns(m, [c for c in free if c not in banned])
            near_banned = [c for c in suggest_columns(m, free) if c in banned]
            used = (f" (비슷한 {', '.join(repr(c) for c in near_banned)}은(는) 이 명령에서 이미 다른 이름이나 옵션에 "
                    f"쓰인 열)") if near_banned else ""
            if values and not near_banned and len(s) == 1 and other_unit(m, s[0]):
                hints.append(unit_hint(m, s[0], kind))
            elif s:
                shown = unit_marked(s, m, quote=True) if values else [repr(c) for c in s]
                cmds = candidate_commands(m, s) if len(missing) == 1 and not values else []
                if cmds:      # 2.9.71: numbered, the same number of next_commands is the full command (no splicing)
                    picks = cmds
                    hints.append(f"'{m}' -> [후보] " + " ".join(f"{i}) {c}" for i, c in enumerate(shown, 1))
                                 + f"{used} (사용자가 고른 번호의 confirm_commands를 실행)")
                else:
                    hints.append(f"'{m}' -> [후보] {', '.join(shown)}{used}")
            elif near_banned:
                hints.append(f"'{m}' -> 비슷한 열 {', '.join(repr(c) for c in near_banned)}은(는) 이 명령에서 이미 다른 "
                             f"이름이나 옵션에 쓰인 열입니다")
            else:
                hints.append(f"'{m}' -> 유사 컬럼 없음")
        ask = (" 비슷한 이름은 다른 양일 수 있으니 어느 열인지 사용자에게 물으십시오."
               if any("[후보]" in h for h in hints) else "")
        raise ColumnNameError(
            f"지정하신 {kind} 컬럼 {missing}을(를) 데이터셋에서 찾을 수 없습니다. "
            f"{' / '.join(hints)}.{ask}\n[사용 가능한 컬럼 목록]: {name_list(col_list)}",
            extra={"confirm_commands": picks} if picks else None,
        )
    for short, full in applied.items():
        sys.stderr.write(f"[INFO] 컬럼명 자동 매칭: '{short}' -> '{full}'\n")
    return resolved, applied


# ==============================================================================
# C3-B. Row conditions: ml_train.py --exclude-where, ml_predict.py --where (constrained search)
# ==============================================================================
# The syntax of office-cli excel.py read --where: '<column> <op> <value>' joined by ';' (AND) and '|' (OR, ';' binds
# first), operators >= <= != > < = ~ (~ = contains), a value with ; or | in single or double quotes.
# Numbers compare as numbers ('1,500' is 1500), dates (YYYY-MM-DD) as dates, other values as trimmed text, case-insensitive.

_WHERE_HEAD = re.compile(r"\s*([^;|]+?)\s*(>=|<=|!=|>|<|=|~)[ \t]*", re.S)
_WHERE_FIXES = (("==", "="), ("=>", ">="), ("=<", "<="), ("<>", "!="), ("=!", "!="))   # one meaning only
_WHERE_AND = re.compile(r"\s+(?:and|AND|And|그리고)\s+|\s*&&\s*")
_WHERE_OR = re.compile(r"\s+(?:or|OR|Or|또는)\s+|\s*\|\|\s*")
_WHERE_CONTAINS = re.compile(r"\s+(?:contains|CONTAINS|포함)\s+")
_WHERE_QUOTED = re.compile(r"(\"[^\"]*\"|'[^']*')")
_WHERE_WORD = re.compile(f"(?P<and>{_WHERE_AND.pattern})|(?P<or>{_WHERE_OR.pattern})|(?P<has>{_WHERE_CONTAINS.pattern})")
_WHERE_OP = re.compile(r">=|<=|!=|==|=>|=<|<>|=!|>|<|=|~")
_WHERE_DATE = re.compile(r"\d{4}-\d{2}-\d{2}(?:[ T]\d{2}:\d{2}(?::\d{2})?)?")
WHERE_FORM = ("'<열> <연산자> <값>' (연산자 >= <= > < = != ~, ~는 포함. 그리고는 ;, 또는은 | (;가 먼저 묶임), "
              "값에 ; | 가 있으면 작은따옴표로 감쌈. 예: 생산효율 > 90; 라인 = B | 라인 = C)")

Cond = Tuple[str, str, str]
# ',' between two conditions ('라인 = A, 온도 > 180'): read as ONE condition (라인 equals the text 'A, 온도 > 180') it
# matched no row and training excluded nothing without a word (2.9.69). A ',' after a condition with an operator and
# before '<name with a letter> <operator>' (outside quotes) is that mistake: STOP_RETRY [추천] the command with ';'.
# '1,500' (no name after it) and a quoted value ('비고 = \'A, B > 3\'') are not touched.
_WHERE_COMMA = re.compile(r"(?:>=|<=|!=|>|<|=|~)[^;|]*?(,)\s*(?=[^;|,=<>!~'\"]*[^\W\d_][^;|,=<>!~'\"]*?(?:>=|<=|!=|>|<|=|~))")


def _where_comma_fix(text: str) -> str:
    """text with every ',' that joins two conditions (outside quotes) replaced by '; ' (_WHERE_COMMA)."""
    masked = _WHERE_QUOTED.sub(lambda m: "Q" * len(m.group(0)), text)
    cuts = [m.start(1) for m in _WHERE_COMMA.finditer(masked)]
    for i in reversed(cuts):
        text = text[:i] + ";" + text[i + 1:]
    return text


def _where_parts(text: str) -> Optional[List[List[Cond]]]:
    """'<col> <op> <value>' joined by ; and | -> OR groups [[cond, ...], ...]; ; | inside a quoted value do not split."""
    groups: List[List[Cond]] = [[]]
    pos, n = 0, len(text)
    while pos < n:
        if not text[pos:].strip():
            break
        m = _WHERE_HEAD.match(text, pos)
        if not m or not m.group(1).strip():
            return None
        col, op = m.group(1).strip(), m.group(2)
        pos = m.end()
        q = text[pos] if pos < n and text[pos] in "\"'" else ""
        close = re.compile(re.escape(q) + r"\s*(?=[;|]|$)").search(text, pos + 1) if q else None
        if close:
            val, pos = text[pos + 1:close.start()], close.end()
        else:
            nxt = re.search(r"[;|]", text[pos:])
            stop = pos + nxt.start() if nxt else n
            val, pos = text[pos:stop].strip(), stop
            if len(val) >= 2 and val[0] == val[-1] and val[0] in "\"'":
                val = val[1:-1]
        if op in (">", "<", ">=", "<=", "~") and val == "":
            return None
        groups[-1].append((col, op, val))
        if pos < n:
            sep = text[pos]
            pos += 1
            if sep == "|":
                if not groups[-1]:
                    return None
                groups.append([])
    groups = [g for g in groups if g]
    return groups or None


def _where_value_text(v: str) -> str:
    """A value as written in a condition: in single quotes when it holds ; | an operator sign or a joining word
    ('불량 포함 시료', 'A and B'), so the text reads back as the same value."""
    quote = re.search(r"[;|]", v) or v != v.strip() or _WHERE_WORD.search(f" {v} ") or _WHERE_OP.search(v)
    return f"'{v}'" if quote else v


def _where_words(part: str) -> Tuple[str, bool]:
    """Joining words outside a value -> the notation (' and ' -> '; ', ' or ' -> ' | ', '<열> 포함 <값>' -> '~'),
    and whether such a word was left inside a value. A word is a joint only where it can be one: 'and' / 'or' after
    a condition with an operator and before another '<열> <연산자>', '포함' right after a column name (no operator
    yet in the condition). '비고 = 불량 포함 시료' keeps its value."""
    out: List[str] = []
    pos = 0
    has_op = False            # the current condition (since the last ; | or joint) has an operator
    kept = False
    for m in _WHERE_WORD.finditer(part):
        before = part[pos:m.start()]
        cut = max(before.rfind(";"), before.rfind("|"))
        has_op = bool(_WHERE_OP.search(before[cut + 1:])) if cut >= 0 else has_op or bool(_WHERE_OP.search(before))
        out.append(before)
        rest = part[m.end():]
        nxt = re.search(r"[;|]", rest)
        upto = rest[:nxt.start()] if nxt else rest
        cond = re.split(f"{_WHERE_AND.pattern}|{_WHERE_OR.pattern}", upto)[0]
        if m.group("has"):
            if has_op:
                out.append(m.group(0))
                kept = True
            else:
                out.append(" ~ ")
                has_op = True
        elif has_op and (_WHERE_OP.search(cond) or _WHERE_CONTAINS.search(cond)):
            out.append("; " if m.group("and") else " | ")
            has_op = False
        else:
            out.append(m.group(0))
            kept = True
        pos = m.end()
    out.append(part[pos:])
    return "".join(out), kept


_WHERE_SIZE_OPS = (">", "<", ">=", "<=")


def _where_size_ok(v: str) -> bool:
    n = coerce_numeric(pd.Series([v])).iloc[0]
    return (n is not None and n == n) or bool(_WHERE_DATE.fullmatch(v.strip()))


def _where_size_value(groups: List[List[Cond]], option: str) -> List[List[Cond]]:
    """A size comparison (> < >= <=) needs a number or a date: '9 0' -> STOP_RETRY [추천] '90' (spaces removed, one
    meaning), other text -> STOP_RETRY (a text value compared by size matched no row without a word)."""
    bad = [(c, o, v) for g in groups for c, o, v in g if o in _WHERE_SIZE_OPS and not _where_size_ok(v)]
    if not bad:
        return groups
    c, o, v = bad[0]
    fix = re.sub(r"\s+", "", v)
    if len(bad) == 1 and fix != v and _where_size_ok(fix):
        fixed = [[(c2, o2, fix if (c2, o2, v2) == bad[0] else v2) for c2, o2, v2 in g] for g in groups]
        raise InputError(f"{option}의 '{c} {o} {v}'에서 크기 비교 값 '{v}'이(가) 숫자로 읽히지 않습니다 (가운데 공백). "
                         f"[추천] {where_retry(option, where_text(fixed))}")
    raise InputError(f"{option}의 '{c} {o} {v}'에서 크기 비교(> < >= <=) 값 '{v}'이(가) 숫자나 날짜(YYYY-MM-DD)가 "
                     f"아닙니다. 글자 값은 = != ~로 비교하십시오.")


def where_text(groups: List[List[Cond]]) -> str:
    """The OR groups as one condition text (a value with ; or | in single quotes)."""
    return " | ".join("; ".join(f"{c} {o} {_where_value_text(v)}" for c, o, v in g) for g in groups)


def _outside_quotes(text: str, fn: Callable[[str], str]) -> str:
    return "".join(part if i % 2 else fn(part) for i, part in enumerate(_WHERE_QUOTED.split(text)))


def where_retry(option: str, fixed: str) -> str:
    """[추천] for a fixed condition: the running command with the option replaced (a full command, R8), else the text."""
    cmd = command_text(drop={option: ""})
    return f"{cmd} {option} {quote_arg(fixed)}" if cmd else fixed


def option_retry(fixed: str, option: str, *aliases: str) -> str:
    """[추천] for a fixed option value: the running command with `option` (also written as one of `aliases`) given
    `fixed` (a full command, R8); in-process use (no ml-kit command): the option and the value."""
    cmd = command_text(drop={o: "" for o in (option,) + aliases})
    return f"{cmd} {option} {quote_arg(fixed)}" if cmd else f"{option} {quote_arg(fixed, 'cmd')}"


def parse_where(text: str, option: str = "--where") -> List[List[Cond]]:
    """Condition text -> OR groups [[(column, op, value), ...], ...] (every condition of a group must hold).
    Notation with one meaning (==, =>, <>, AND, OR) -> STOP_RETRY [추천] the command with the fixed text."""
    raw = str(text).strip()
    kept: List[bool] = []

    def fix(part: str) -> str:
        part, k = _where_words(part)
        kept.append(k)
        for a, b in _WHERE_FIXES:
            part = part.replace(a, b)
        return part
    fixed = _outside_quotes(raw, fix)
    comma = _where_comma_fix(fixed)
    groups = _where_parts(comma)
    if groups is None:
        raise InputError(f"{option} 식 '{raw}'을(를) 읽을 수 없습니다. 형식: {WHERE_FORM}")
    if comma != fixed:
        raise InputError(f"{option} 식 '{raw}'에서 조건 사이를 ','로 나눴습니다. ','로 나누면 조건 하나(앞 열의 값이 ',' 뒤까지 "
                         f"전부)로 읽혀 맞는 행이 없습니다. 그리고는 ;, 또는은 |로 나눕니다. "
                         f"[추천] {where_retry(option, where_text(groups))}")
    if fixed != raw:
        # a word or sign left inside a value ('불량 포함 시료') is quoted in the [추천] so it stays the value
        quoted = any(kept) or any(_where_value_text(v) != v for g in groups for _, _, v in g)
        raise InputError(f"{option} 식 '{raw}'에 쓸 수 없는 표기가 있습니다 (같음은 =, 같지 않음은 !=, 그리고는 ;, 또는은 |). "
                         + ("값 안의 낱말(포함, and, or)과 기호는 바꾸지 않았습니다. 값을 작은따옴표로 감싸세요. " if quoted else "")
                         + f"형식: {WHERE_FORM}. [추천] {where_retry(option, where_text(groups))}")
    return _where_size_value(groups, option)


def resolve_where(groups: List[List[Cond]], columns: List[str], option: str = "--where",
                  layout: Optional[Dict[str, Any]] = None) -> List[List[Cond]]:
    """Column names of the conditions resolved like -x / -y (resolve_column_names: exact, a unique name without its
    unit, a similar name -> [후보])."""
    names = list(dict.fromkeys(c for g in groups for c, _, _ in g))
    resolved = resolve_column_names(names, columns, f"조건({option})의 열", layout=layout)[0]
    rmap = dict(zip(names, resolved))
    return [[(rmap[c], o, v) for c, o, v in g] for g in groups]


_CODE_RE = re.compile(r"\s*[-+]?0\d")     # a code with a leading zero ('007', '01'), never '0.5'

_FIXED_CODE_RE = re.compile(r"\d{3,}")    # a fixed-width code ('001' ~ '120'); two digits ('05', '10') read as numbers


def is_code_column(series: pd.Series) -> bool:
    """A category / code column, matched as exact text ('007' is not '7', not a number to interpolate) by
    --where / --exclude-where and ml_map keys. Only when some cell has a leading zero ('007') and
    (a) the same number is written in two ways ('007' and '7', '01' and '001'), or
    (b) every filled cell is a code of the same width of 3 or more digits ('001' ~ '120').
    Otherwise the column is read as numbers, as training reads it: months '01' ~ '12' and '05, 10, 15, 20' are
    number columns ('월 = 1' meets '01'). A number or date column (Excel cells) is never one; missing-value marks
    and blank cells do not count."""
    if pd.api.types.is_numeric_dtype(series) or pd.api.types.is_datetime64_any_dtype(series):
        return False
    filled = [v.strip() if isinstance(v, str) else v for v in series.tolist()
              if v is not None and not (isinstance(v, float) and v != v)]
    filled = [v for v in filled if not (isinstance(v, str) and (v == "" or v in MISSING_MARKS))]
    zero = [isinstance(v, str) and _CODE_RE.match(v) is not None for v in filled]
    if not any(zero):
        return False
    if all(isinstance(v, str) and _FIXED_CODE_RE.fullmatch(v) for v in filled) and len({len(v) for v in filled}) == 1:
        return True
    nums = coerce_numeric(pd.Series(filled, dtype=object)).tolist()
    ways: Dict[float, list] = {}               # number -> [its notations (01, 001, 1), one has a leading zero]
    for v, n, z in zip(filled, nums, zero):
        if n is not None and n == n:
            ways.setdefault(n, [set(), False])
            ways[n][0].add(str(v))
            ways[n][1] = ways[n][1] or z
    return any(len(texts) > 1 and has_zero for texts, has_zero in ways.values())


def _where_cells(series: pd.Series) -> pd.Series:
    """The cells of a condition column as stripped text ('' for a blank)."""
    return series.map(lambda v: "" if v is None or (isinstance(v, float) and v != v) else str(v).strip())


def _compare_as_text(series: pd.Series, text: pd.Series, val: str) -> bool:
    """'= v' / '!= v' compares as text only for codes with a leading zero: a code column (is_code_column), or a code
    value ('= 01') on a column with text values ('01' is not '1'). Otherwise a number value meets the cells read as
    numbers by number ('2' = '2.0', '1,000' = '1000', as training reads them), text cells and missing-value marks
    ('측정불가') by text: one '측정불가' in a number column does not make it a text column."""
    if pd.api.types.is_numeric_dtype(series) or _WHERE_DATE.fullmatch(val.strip()):
        return False
    if is_code_column(series):
        return True
    if not _CODE_RE.match(val):
        return False
    filled = text[(text != "") & ~text.isin(MISSING_MARKS)]
    return bool(coerce_numeric(filled).isna().any())


def _where_cond_mask(series: pd.Series, op: str, val: str) -> np.ndarray:
    """Rows meeting one condition. A blank cell never meets '!=' (an empty cell is not a value that differs)."""
    text = _where_cells(series)
    blank = (text == "").to_numpy(dtype=bool)
    if op == "~":
        return text.str.casefold().str.contains(val.strip().casefold(), regex=False).to_numpy(dtype=bool)
    vnum = coerce_numeric(pd.Series([val])).iloc[0]
    cmp = {"=": np.equal, "!=": np.not_equal, ">": np.greater, "<": np.less, ">=": np.greater_equal,
           "<=": np.less_equal}[op]
    if op in ("=", "!=") and _compare_as_text(series, text, val):
        eq = (text.str.casefold() == val.strip().casefold()).to_numpy(dtype=bool)
        return eq if op == "=" else ~eq & ~blank
    if vnum == vnum and vnum is not None:
        num = coerce_numeric(series).to_numpy(dtype=float)
        ok = ~np.isnan(num)
        out = np.zeros(len(series), dtype=bool)
        out[ok] = cmp(num[ok], float(vnum))
        if op in ("=", "!="):        # a text cell against a number: compared as text
            tv = (text.str.casefold() == val.strip().casefold()).to_numpy(dtype=bool)
            out[~ok] = tv[~ok] if op == "=" else ~tv[~ok]
        return out & ~blank if op == "!=" else out
    if _WHERE_DATE.fullmatch(val.strip()):
        dt = parse_datetime(series)
        vd = pd.Timestamp(val.strip())
        ok = dt.notna().to_numpy(dtype=bool)
        out = np.zeros(len(series), dtype=bool)
        if len(val.strip()) == 10:
            left = dt[ok].dt.normalize()
        else:
            left = dt[ok]
        out[ok] = cmp(left.to_numpy(), np.datetime64(vd))
        if op == "!=":
            out[~ok] = True
        return out & ~blank if op == "!=" else out
    if op not in ("=", "!="):
        return np.zeros(len(series), dtype=bool)   # a size comparison with a text value: no row
    eq = (text.str.casefold() == val.strip().casefold()).to_numpy(dtype=bool)
    return eq if op == "=" else ~eq & ~blank


def where_mask(df: pd.DataFrame, groups: List[List[Cond]]) -> np.ndarray:
    """Rows of df that meet the condition (any OR group whose conditions all hold). Column names must be resolved."""
    out = np.zeros(len(df), dtype=bool)
    for g in groups:
        hit = np.ones(len(df), dtype=bool)
        for col, op, val in g:
            hit &= _where_cond_mask(df[col], op, val)
        out |= hit
    return out


def where_notation_variants(df: pd.DataFrame, groups: List[List[Cond]]) -> List[Tuple[Cond, Dict[str, int]]]:
    """Conditions '= v' / '!= v' compared as text (a code column) whose number is also written another way in
    that column ('월 = 1' and cells '01'): [(condition, {other notation: rows})]. Those rows read as the same number
    in training but do not meet '=' (and do meet '!='): the caller warns and offers where_with_variants."""
    out: List[Tuple[Cond, Dict[str, int]]] = []
    # notations the condition already names ('월 = 1 | 월 = 01', '월 != 1; 월 != 01') need no note
    named = {(c, o, v.strip().casefold()) for g in groups for c, o, v in g}
    for cond in dict.fromkeys(c for g in groups for c in g):
        col, op, val = cond
        if op not in ("=", "!=") or col not in df.columns:
            continue
        series = df[col]
        text = _where_cells(series)
        vnum = coerce_numeric(pd.Series([val])).iloc[0]
        if vnum is None or vnum != vnum or not _compare_as_text(series, text, val):
            continue
        same = (coerce_numeric(series) == float(vnum)) & (text != "") & (text.str.casefold() != val.strip().casefold())
        counts = {str(k): int(n) for k, n in text[same].value_counts().items() if (col, op, str(k).casefold()) not in named}
        if counts:
            out.append((cond, counts))
    return out


def where_with_variants(groups: List[List[Cond]], variants: List[Tuple[Cond, Dict[str, int]]]) -> List[List[Cond]]:
    """The condition that also covers the other notations: '월 = 1' -> '월 = 1 | 월 = 01' (each OR group repeated
    with the other notation), '월 != 1' -> '월 != 1; 월 != 01'."""
    vm = {cond: list(cnt) for cond, cnt in variants}
    out: List[List[Cond]] = []
    for g in groups:
        base = list(g) + [(c, "!=", t) for c, o, v in g if o == "!=" and (c, o, v) in vm for t in vm[(c, o, v)]]
        alts = [[cond] + ([(cond[0], "=", t) for t in vm[cond]] if cond[1] == "=" and cond in vm else []) for cond in base]
        for combo in product(*alts):
            if list(combo) not in out:
                out.append(list(combo))
    return out


# ==============================================================================
# C4. Robust Data Loader & Layout Parser with Auto-Header Sniffing
# ==============================================================================

_NA_PLACEHOLDERS = ["-", "N/A", "n/a", "NA", "null", "NULL", "미측정", "None", "nan", "NaN", ""]


_GROUPED_NUMBER_RE = re.compile(r"[+-]?[1-9]\d{0,2}(,\d{3})+(\.\d+)?")   # never '0,149'


def _strip_thousands(text: str) -> str:
    """'1,500' -> '1500'. 천 단위 묶음이 아닌 쉼표('175,6', '1,2')는 그대로 둡니다."""
    return text.replace(",", "") if "," in text and _GROUPED_NUMBER_RE.fullmatch(text) else text


_DECIMAL_COMMA_TEXT_RE = re.compile(r"[+-]?\d+,\d+")


def unparsed_number_cells(series: pd.Series, rows: Optional[np.ndarray] = None) -> Tuple[List[int], List[str]]:
    """Positions and first texts of the filled cells (of rows, a boolean mask; all rows if None) that a numeric
    model input cannot read (coerce_numeric -> NaN): training fills them with the median."""
    if pd.api.types.is_numeric_dtype(series) or pd.api.types.is_datetime64_any_dtype(series):
        return [], []
    text = series.astype("object").map(lambda v: v.strip() if isinstance(v, str) else v)
    bad = (text.notna() & ~text.isin(_NA_PLACEHOLDERS) & coerce_numeric(series).isna()).to_numpy(dtype=bool)
    if rows is not None:
        bad = bad & np.asarray(rows, dtype=bool)
    pos = np.flatnonzero(bad).tolist()
    return pos, list(dict.fromkeys(str(text.iloc[i]) for i in pos))[:3]


def decimal_comma_text_share(series: pd.Series) -> Tuple[int, int, int, List[str]]:
    """(filled cells, decimal-comma texts '175,6' / '0,149' not read as numbers, cells read as numbers, examples)
    of a column: a column of such texts is numbers saved with a decimal comma, not a category."""
    if pd.api.types.is_numeric_dtype(series) or pd.api.types.is_datetime64_any_dtype(series):
        return 0, 0, 0, []
    text = series.astype("object").map(lambda v: v.strip() if isinstance(v, str) else v)
    filled = (text.notna() & ~text.isin(_NA_PLACEHOLDERS)).to_numpy(dtype=bool)
    num = coerce_numeric(series).notna().to_numpy(dtype=bool) & filled
    dc = np.array([isinstance(v, str) and bool(_DECIMAL_COMMA_TEXT_RE.fullmatch(v)) for v in text], dtype=bool)
    dc = dc & filled & ~num
    ex = list(dict.fromkeys(str(text.iloc[i]) for i in np.flatnonzero(dc)[:20]))[:3]
    return int(filled.sum()), int(dc.sum()), int(num.sum()), ex


_PERCENT_RE = re.compile(r"([+-]?(\d[\d,]*\.?\d*|\.\d+)([eE][-+]?\d+)?)\s*%")


def _strip_percent(text: str) -> str:
    """'64.4%' / '64.4 %' -> '64.4' (an Excel CSV export of percent cells); any other text unchanged."""
    m = _PERCENT_RE.fullmatch(text)
    return m.group(1) if m else text


def percent_cells(series: pd.Series) -> Tuple[int, List[str]]:
    """(cells written as a number with '%', first examples) of a column: coerce_numeric reads them as the number
    without '%' (64.4, not 0.644). Training, prediction and ml_map warn with percent_warning."""
    if pd.api.types.is_numeric_dtype(series) or pd.api.types.is_datetime64_any_dtype(series):
        return 0, []
    hits = [v.strip() for v in series.astype("object")
            if isinstance(v, str) and "%" in v and _PERCENT_RE.fullmatch(v.strip())]
    return len(hits), list(dict.fromkeys(hits))[:3]


def percent_warning(df: pd.DataFrame, columns: List[str], where: str = "") -> List[str]:
    """One warning per column of columns with numbers written with '%' (percent_cells)."""
    out = []
    for c in dict.fromkeys(columns):
        if c not in df.columns:
            continue
        n, ex = percent_cells(df[c])
        if n:
            first = _strip_percent(ex[0])
            out.append(f"{where}'{c}' 열의 {n}칸({', '.join(repr(v) for v in ex)})은 %를 떼고 숫자로 읽었습니다 "
                       f"({ex[0]} -> {first}, 100으로 나누지 않음. .xlsx의 백분율 서식 칸도 화면에 보이는 숫자로 읽음). "
                       f"열의 단위가 %인지 확인하십시오.")
    return out


def percent_mix(df: pd.DataFrame, columns: List[str]) -> List[Tuple[str, str, str]]:
    """(column, a '%' cell, a plain number cell) of each column of columns holding both forms (check_percent_mix)."""
    out = []
    for c in dict.fromkeys(columns):
        if c not in df.columns:
            continue
        s = df[c]
        if pd.api.types.is_numeric_dtype(s) or pd.api.types.is_datetime64_any_dtype(s):
            continue
        obj = s.astype("object")
        is_pct = np.array([isinstance(v, str) and "%" in v and bool(_PERCENT_RE.fullmatch(v.strip())) for v in obj],
                          dtype=bool)
        is_num = coerce_numeric(s).notna().to_numpy(dtype=bool)
        if is_pct.any() and (is_num & ~is_pct).any():
            out.append((c, str(obj.iloc[int(np.flatnonzero(is_pct)[0])]).strip(),
                        str(obj.iloc[int(np.flatnonzero(is_num & ~is_pct)[0])]).strip()))
    return out


def check_percent_mix(df: pd.DataFrame, columns: List[str], where: str = "") -> None:
    """STOP_RETRY when a column of columns holds numbers written with '%' ('11.9%') AND plain numbers ('0.119'):
    coerce_numeric reads 11.9% as 11.9, so the two forms may differ by 100x and no rule tells which scale is meant.
    The user makes the column one form. A column of '%' cells only is read as before (percent_warning)."""
    for c, pct, plain in percent_mix(df, columns)[:1]:
        raise InputError(
            f"{where}'{c}' 열에 %를 붙인 숫자 칸({pct!r})과 % 없는 숫자 칸({plain!r})이 섞여 있습니다. "
            f"%를 붙인 값은 % 없이 읽으므로({pct} -> {_strip_percent(pct)}) 두 형식은 100배 차이가 날 수 있어 "
            f"어느 배율인지 정할 수 없습니다. 사용자에게 알려 이 열을 모두 %를 붙인 값으로, 또는 모두 % 없는 숫자"
            f"(한 가지 배율, 예: 11.9 또는 0.119)로 고친 뒤 다시 실행하십시오 (.xlsx는 열 전체를 같은 셀 서식으로).")


_NA_PLACEHOLDER_SET = frozenset(_NA_PLACEHOLDERS)


def _number_cell(v: Any) -> Any:
    """One cell of coerce_numeric before pd.to_numeric: text stripped, '%' and thousands separators removed, a
    placeholder ('-', 'N/A', ...) made NaN; any other value unchanged."""
    if not isinstance(v, str):
        return v
    t = v.strip()
    if "%" in t:
        t = _strip_percent(t)
    t = _strip_thousands(t)
    return np.nan if t in _NA_PLACEHOLDER_SET else t


def coerce_numeric(series: pd.Series) -> pd.Series:
    """
    Converts a column to float regardless of its storage dtype (object in pandas 2,
    'str' in pandas 3). Strips thousands separators, a trailing '%' ('64.4%' -> 64.4, the number as written:
    the column name usually says (%) already; percent_cells tells the user) and placeholder NaN strings.
    Numeric and datetime columns are returned unchanged.
    """
    if pd.api.types.is_numeric_dtype(series) or pd.api.types.is_datetime64_any_dtype(series):
        return series
    s = series.astype("object").where(series.notna(), None)
    # 천 단위 쉼표만 지움 ('1,500', '1,234,567.5'). '175,6'처럼 쉼표 뒤가 3자리가 아니면 천 단위가 아니므로
    # 지우지 않음 -> 숫자로 못 읽은 값(결측 경고)이 됨 (1756으로 조용히 바뀌지 않게). 칸마다 한 번에 (열이 많은
    # 큰 데이터: 칸마다 정규식과 replace를 따로 돌리면 학습 전 준비만 수십 초)
    s = pd.Series([_number_cell(v) for v in s.to_numpy()], index=s.index, name=s.name, dtype=object)
    return pd.to_numeric(s, errors="coerce").astype(float)


_CSV_ENCODINGS = ["utf-8-sig", "cp949", "euc-kr", "latin1"]


_CSV_DELIMITERS = ",;\t|"


def _delimiter_score(lines: List[str], d: str) -> int:
    """Lines of the most common width (2+ fields) when split by d: a table splits most lines evenly."""
    try:
        widths = [len(r) for r in csv.reader(lines, delimiter=d)]
    except csv.Error:
        return 0
    multi = [w for w in widths if w > 1]
    return max(multi.count(w) for w in set(multi)) if multi else 0


def _sniff_delimiter(sample: str) -> str:
    """csv.Sniffer's delimiter, unless another one splits more lines into the table width: a short ';' file of
    decimal-comma numbers ('163,1;2,19;A;61,66' under a title line) is sniffed as ','."""
    try:
        guess = csv.Sniffer().sniff(sample, delimiters=_CSV_DELIMITERS).delimiter
    except csv.Error:
        guess = ","
    lines = [ln for ln in sample.splitlines() if ln.strip()][:50]
    scores = {d: _delimiter_score(lines, d) for d in _CSV_DELIMITERS}
    best = max(_CSV_DELIMITERS, key=lambda d: scores[d])
    return best if scores[best] > scores.get(guess, 0) else guess


# Number cells of a CSV not separated by commas (fullmatch). Each name is assigned once in this module: a second
# assignment of a shared name (an unanchored pattern for check_comma_numbers) would silently change this check.
_DECIMAL_COMMA_RE = re.compile(r"\s*[-+]?\d+,\d+\s*")                          # '175,6', '0,149', '27,8448'
_CSV_THOUSANDS_RE = re.compile(r"\s*[-+]?[1-9]\d{0,2}(,\d{3})+(\.\d+)?\s*")   # '1,500', '1,346.8' (never '0,149')
_EURO_NUMBER_RE = re.compile(r"\s*[-+]?[1-9]\d{0,2}(\.\d{3})+,\d+\s*")       # '1.346,8': thousands dot, decimal comma
_DOT_THOUSANDS_RE = re.compile(r"\s*[-+]?[1-9]\d{0,2}(\.\d{3})+\s*")          # '1.346': thousands dot or 3 decimals
_DOT_COMMA_NUMBER_RE = re.compile(r"\s*[-+]?[\d.,]*\d[\d.,]*\s*")              # digits with '.' and ',' only


_PLAIN_DOT_RE = re.compile(r"\s*[-+]?\d+\.\d+\s*")                              # '61.5', '0.300', '1.300'


def _comma_columns(df: pd.DataFrame) -> List[Dict[str, Any]]:
    """Number cells of each column of a CSV read as text (_decimal_commas): decimal-comma cells ('175,6', '1,2'),
    thousands-comma cells ('1,500'), European numbers ('1.346,8') and other dotted comma numbers, dot-grouped numbers
    without a comma ('1.300', '2.125'), plain decimal-point numbers that cannot be a thousands group ('61.5', '0.300').
    text: most filled cells are not numbers (a line-list / note column whose '1,2' is no number)."""
    out = []
    for j in df.columns:
        col = df[j]
        vals = col[col.map(lambda v: isinstance(v, str) and v.strip() != "")].astype(str)
        vals = vals[~vals.str.strip().isin(_NA_PLACEHOLDERS)]
        commas = vals[vals.str.contains(",", regex=False)]
        dec = commas[commas.str.fullmatch(_DECIMAL_COMMA_RE)]
        dotted = commas[commas.str.contains(".", regex=False) & commas.str.fullmatch(_DOT_COMMA_NUMBER_RE)]
        nocomma = vals[~vals.str.contains(",", regex=False) & vals.str.contains(".", regex=False)]
        grouped = nocomma[nocomma.str.fullmatch(_DOT_THOUSANDS_RE)]
        plain = nocomma[nocomma.str.fullmatch(_PLAIN_DOT_RE)]
        numberish = int(vals.str.fullmatch(_DOT_COMMA_NUMBER_RE).sum())
        out.append({"dec": dec, "only_dec": dec[~dec.str.fullmatch(_CSV_THOUSANDS_RE)],
                    "thousands": commas[commas.str.fullmatch(_CSV_THOUSANDS_RE)],
                    "euro": dotted[dotted.str.fullmatch(_EURO_NUMBER_RE)], "dotted": dotted,
                    "grouped": grouped, "plain": plain[~plain.str.fullmatch(_DOT_THOUSANDS_RE)],
                    "text": len(vals) > 0 and numberish * 2 < len(vals)})
    return out


def _decimal_commas(df: pd.DataFrame, sep: str, trained: bool = False) -> pd.DataFrame:
    """A CSV not separated by commas (';', tab) whose numbers use a decimal comma ('175,6', European Excel): a comma
    that cannot be a thousands separator ('175,6', '2,19', '0,149', '27,8448') means the '<digits>,<digits>' cells are
    decimal numbers, read with a decimal point (df.attrs["decimal_comma"], layout_warnings). European numbers with a
    thousands dot ('1.346,8') are read as 1346.8 (with a warning) when every number cell with both '.' and ',' has that
    form. Judged per column: a mostly-text column (a line list '1,2' among '1,2,3' / 'A') is no evidence and
    is never changed, and a file with a plain decimal-point column ('61.5') uses '.' as its decimal point, so no
    column is read with decimal commas there (a '175,6' column stays text: training stops on it as before).
    Refused (InputError, a value would change by 1000x either way) in a decimal-comma file: cells that could be
    thousands too ('1,500'), dot-grouped numbers without a comma ('1.300', '2.125': 1300 or 1.3), or European numbers
    mixed with other '.'/',' forms. trained: the model (or the ml_map data file) was read with decimal commas, so this
    file is read the same way whatever its delimiter (a one-column file is sniffed as ','), and refused when it has no
    decimal-comma cell but a '1.300' or '1,500' (read differently from the training file). A comma CSV, or one
    without such cells, is unchanged (coerce_numeric strips thousands separators)."""
    if df.empty or (sep == "," and not trained):
        return df
    cols = _comma_columns(df)
    num = [c for c in cols if not c["text"]]               # text columns: never evidence, never changed
    for c in num:
        c["dc"] = not (c["only_dec"].empty and c["euro"].empty)
    confirmed = any(c["dc"] for c in num)
    # a plain decimal-point number ('61.5') in a column without decimal commas: '.' is this file's decimal point
    dot_file = any(not c["dc"] and not c["plain"].empty and c["grouped"].empty for c in num)
    if not trained and (not confirmed or dot_file):
        return df

    def cat(key: str) -> pd.Series:
        parts = [c[key] for c in num if not c[key].empty]
        return pd.concat(parts) if parts else pd.Series([], dtype=object)

    only_dec, euro, dec, thousands, dotted, grouped = (cat(k) for k in ("only_dec", "euro", "dec", "thousands",
                                                                        "dotted", "grouped"))
    first = (only_dec if not only_dec.empty else euro).iloc[0].strip() if confirmed else None
    why = (f"소수점 쉼표('{first}')" if confirmed else "소수점 쉼표(학습 파일이 소수점 쉼표 형식)")
    tail = (" 파일 하나에는 한 가지 숫자 형식만 쓰십시오: Excel에서 숫자 형식을 소수점 '.'으로 바꿔 저장하거나 "
            ".xlsx로 저장한 뒤 다시 실행하십시오.")
    if not thousands.empty:
        t0 = thousands.iloc[0].strip()
        raise InputError(
            f"CSV(구분 기호 '{sep}')의 숫자에 {why}와 천 단위로도 읽히는 쉼표('{t0}')가 섞여 있어 값을 정할 수 "
            f"없습니다 ('{t0}'이 소수인지 천 단위인지에 따라 1000배 달라짐).{tail}")
    if not euro.empty:
        odd = dotted[~dotted.str.fullmatch(_EURO_NUMBER_RE)]
        if not odd.empty:
            raise InputError(
                f"CSV(구분 기호 '{sep}')의 숫자에 점(.)이 천 단위인 숫자('{euro.iloc[0].strip()}')와 다른 형식의 숫자"
                f"('{odd.iloc[0].strip()}')가 섞여 있어 값을 정할 수 없습니다 (점이 천 단위인지 소수점인지에 따라 1000배 "
                f"달라짐).{tail}")
    if not grouped.empty:
        g0 = grouped.iloc[0].strip()
        raise InputError(
            f"CSV(구분 기호 '{sep}')의 숫자에 {why}와 쉼표 없이 점(.)으로 묶인 숫자('{g0}', {len(grouped)}칸)가 "
            f"섞여 있어 값을 정할 수 없습니다 ('{g0}'의 점이 천 단위인지 소수점인지에 따라 1000배 달라짐).{tail}")
    if not confirmed:
        return df

    def fix(v: Any) -> Any:
        if not isinstance(v, str) or "," not in v:
            return v
        if _EURO_NUMBER_RE.fullmatch(v):
            return v.strip().replace(".", "").replace(",", ".")
        return v.replace(",", ".") if _DECIMAL_COMMA_RE.fullmatch(v) else v

    out = df.copy()
    for j, c in zip(df.columns, cols):
        if not c["text"]:
            out[j] = df[j].map(fix)
    e0 = euro.iloc[0].strip() if not euro.empty else None
    out.attrs["decimal_comma"] = [f"{first} -> {fix(first)}", int(len(dec) + len(euro)), sep,
                                  f"{e0} -> {fix(e0)}" if e0 else None, int(len(euro))]
    return out


def _read_csv_fast(file_path: str, decimal_comma: bool = False) -> pd.DataFrame:
    """
    Reads CSV with the fast C parser: encoding and delimiter are detected from the first
    64KB, then the whole file is parsed as strings (numeric conversion happens later in
    coerce_numeric, identically for training and inference). Rows of different widths
    (a title line above the table) are padded to the widest row.
    """
    try:
        with open(file_path, "rb") as f:
            head = f.read(65536)
    except PermissionError:
        raise PermissionError(f"파일이 다른 프로그램(Excel 등)에서 잠겨 있습니다. 파일을 닫고 다시 실행하십시오: {file_path}")

    last_err: Optional[Exception] = None
    for enc in _CSV_ENCODINGS:
        sample = None
        for cut in range(4):  # the 64KB boundary may split a multi-byte character
            try:
                sample = head[:len(head) - cut].decode(enc)
                break
            except UnicodeDecodeError:
                continue
        if sample is None:
            continue
        sep = _sniff_delimiter("\n".join(sample.splitlines()[:50]))
        df = None
        try:
            df = pd.read_csv(file_path, header=None, sep=sep, encoding=enc, dtype=str, low_memory=False)
        except UnicodeDecodeError as e:
            last_err = e
            continue
        except pd.errors.ParserError:
            # ragged rows: a title line narrower than the table ("Expected 1 fields in line 2, saw 4"). Every row is
            # padded to the widest row (names=0..n-1), so header detection and [추천] --header-row work as for .xlsx.
            try:
                with open(file_path, "r", encoding=enc, newline="") as f:
                    width = max((len(r) for r in csv.reader(f, delimiter=sep)), default=1)
                df = pd.read_csv(file_path, header=None, names=list(range(width)), sep=sep, encoding=enc,
                                 dtype=str, engine="python")
            except Exception as e:
                last_err = e
                continue
        except PermissionError:
            raise PermissionError(f"파일이 다른 프로그램(Excel 등)에서 잠겨 있습니다. 파일을 닫고 다시 실행하십시오: {file_path}")
        except Exception as e:
            last_err = e
            continue
        return _decimal_commas(df, sep, decimal_comma)
    raise InputError(f"CSV 파일 로드 실패 (지원되는 인코딩 모두 시도 완료): {last_err}")


HEADER_SCAN_ROWS = 15     # rows searched for the header (auto detection and the --header-row hint)


def load_data_file(
    file_path: str,
    header_row: Optional[int] = None,
    unit_row: Optional[int] = None,
    data_row: Optional[int] = None,
    decimal_comma: bool = False
) -> pd.DataFrame:
    """
    Loads tabular data from CSV (multi-encoding, auto delimiter) or Excel (.xlsx).
    Handles non-standard header, unit, and data row specifications.
    Includes Auto-Header Sniffing: skips blank/title top rows automatically if header_row is not explicit.
    decimal_comma: the model / data file this file goes with was read with decimal commas (_decimal_commas).
    """
    df_raw = read_raw_table(file_path, decimal_comma)
    total_rows = len(df_raw)
    if total_rows == 0:
        raise InputError("입력 데이터 파일이 비어 있습니다.")

    # Auto-Header Sniffing: If header_row is not specified or 1, skip leading empty rows
    if header_row is not None and header_row > 1:
        header_idx = header_row - 1
    else:
        header_idx = sniff_header_index(df_raw)

    if header_idx >= total_rows:
        raise InputError(f"유효한 헤더 행을 찾을 수 없습니다 (전체 행 수: {total_rows}).")
    return _layout_frame(df_raw, header_idx, unit_row, data_row)


def read_raw_table(file_path: str, decimal_comma: bool = False) -> pd.DataFrame:
    """The sheet / CSV as read, without a header (cells as stored; CSV cells as text)."""
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"입력 데이터 파일을 찾을 수 없습니다: {file_path}")

    ext = os.path.splitext(file_path)[1].lower()

    if ext in [".xlsx", ".xlsm"]:
        try:
            df_raw = pd.read_excel(file_path, header=None)
            df_raw = _excel_percent_cells(file_path, df_raw)
        except ImportError:
            raise InputError("openpyxl 패키지가 설치되어 있지 않아 .xlsx를 읽을 수 없습니다. `python -m pip install openpyxl`로 설치하십시오.")
        except PermissionError:
            raise PermissionError(f"파일이 다른 프로그램(Excel 등)에서 잠겨 있습니다. 파일을 닫고 다시 실행하십시오: {file_path}")
    elif ext == ".xls":
        raise InputError("구형 .xls 형식은 지원하지 않습니다. Excel에서 '.xlsx' 또는 '.csv'로 다른 이름으로 저장 후 지정하십시오.")
    elif ext == ".csv":
        df_raw = _read_csv_fast(file_path, decimal_comma)
    else:
        raise InputError(f"지원하지 않는 데이터 포맷입니다 ({ext}). .xlsx 또는 .csv 파일을 지정하십시오.")
    return df_raw


_FORMAT_LITERAL_RE = re.compile(r'"[^"]*"|\\.|\[[^\]]*\]')   # quoted text, escaped character, [color] / [>=100]


def _percent_format(fmt: Any) -> bool:
    """An Excel number format that shows the value x100 with '%' ('0%', '0.0%', '#,##0.00%'); a '%' in quoted text
    or escaped ('0.0"%"', '0.0\\%') is a literal sign that does not scale the value."""
    return isinstance(fmt, str) and "%" in _FORMAT_LITERAL_RE.sub("", fmt)


def _percent_text(v: float) -> str:
    """0.119 (an Excel percent cell) -> '11.9%': the number Excel shows (15 significant digits, Excel's precision),
    as a CSV export writes it, so coerce_numeric, percent_warning and check_percent_mix treat both files alike."""
    return f"{float(v) * 100:.15g}%"


def _excel_percent_cells(file_path: str, df_raw: pd.DataFrame) -> pd.DataFrame:
    """The number cells of the first sheet (the one pd.read_excel read) with a percent number format, written as the
    text Excel shows ('11.9%' for 0.119): read_excel gives the stored fraction, a CSV export of the same
    sheet gives '11.9%' (read as 11.9), so the same data would differ by 100x between the two file formats.
    Sheets without a percent format in their styles are returned unchanged (one cheap style check, no second read)."""
    from openpyxl import load_workbook
    from openpyxl.styles.numbers import BUILTIN_FORMATS

    wb = load_workbook(file_path, read_only=True, data_only=True, keep_links=False)
    try:
        try:   # style table check: no percent format anywhere -> nothing to do
            fmts = [BUILTIN_FORMATS.get(st.numFmtId) if st.numFmtId < 164 else wb._number_formats[st.numFmtId - 164]
                    for st in wb._cell_styles]
            if not any(_percent_format(f) for f in fmts):
                return df_raw
        except Exception:
            pass
        ws = wb.worksheets[0]
        if getattr(ws, "reset_dimensions", None):
            ws.reset_dimensions()
        hits: Dict[int, Dict[int, str]] = {}
        ncols, nrows = df_raw.shape[1], len(df_raw)
        for r, row in enumerate(ws.rows):
            if r >= nrows:
                break
            for c, cell in enumerate(row[:ncols]):
                v = cell.value
                if (cell.data_type == "n" and isinstance(v, (int, float)) and not isinstance(v, bool)
                        and _percent_format(cell.number_format)):
                    hits.setdefault(c, {})[r] = _percent_text(v)
    finally:
        wb.close()
    if not hits:
        return df_raw
    out = df_raw.copy()
    for c, cells in hits.items():
        col = out.iloc[:, c].astype(object)
        for r, text in cells.items():
            col.iat[r] = text
        out[out.columns[c]] = col
    return out


def sniff_header_index(df_raw: pd.DataFrame) -> int:
    """0-based header row read without --header-row: the first row of 2+ filled cells with some text."""
    for r_i in range(min(HEADER_SCAN_ROWS, len(df_raw))):
        row_items = df_raw.iloc[r_i].dropna().tolist()
        if len(row_items) >= 2 and any(isinstance(v, str) and str(v).strip() for v in row_items):
            return r_i
    return 0


class _SharedList(list):
    """A list in df.attrs (source_rows: one number per data row) that copies of the frame share: pandas deep-copies
    attrs on every Series / frame operation (__finalize__), so a 100k-row list made each column operation of loading,
    training and prediction ~0.1 s slower. Never changed after _layout_frame sets it."""

    def __deepcopy__(self, memo: Dict[int, Any]) -> "_SharedList":
        return self


def _layout_frame(df_raw: pd.DataFrame, header_idx: int, unit_row: Optional[int],
                  data_row: Optional[int]) -> pd.DataFrame:
    """The data frame of load_data_file from the raw table and the 0-based header row."""
    total_rows = len(df_raw)
    raw_headers = df_raw.iloc[header_idx].values
    headers = [_cell_text(c) or f"col_{i}" for i, c in enumerate(raw_headers)]
    auto_named = [not _cell_text(c) for c in raw_headers]   # header cell was blank

    # Handle unit row if provided. A row that is not a unit row (numbers, text data, a blank row, the header itself or
    # a row above it) would rename every column ('반응온도(192.58)') or turn the real unit row into data: refused, with
    # the header + unit pair of the file when there is one, else the header row in use without --unit-row.
    if unit_row is not None:
        j = unit_row - 1
        unit_cells = _row_cells(df_raw, j) if 0 <= j < total_rows else []
        ok, reason = unit_row_verdict(df_raw, j, header_idx, given=True) if unit_cells and j > header_idx else (False, "")
        if not ok:
            pair = _header_unit_pair(df_raw)
            if pair and pair != (header_idx + 1, unit_row):
                hint = f" 이 파일의 머리글은 {pair[0]}행, 단위는 {pair[1]}행으로 보입니다. [추천] {layout_retry(pair[0], pair[1])}"
            elif header_idx < j < total_rows and _header_under_title(df_raw, header_idx, j):
                hint = f" 이 파일의 머리글은 {unit_row}행으로 보입니다. [추천] {layout_retry(unit_row, None)}"
            else:
                hint = f" [추천] {layout_retry(header_idx + 1, None)}"
            if not unit_cells:
                why = "빈 행이라"
            elif j <= header_idx:
                why = f"머리글 행({header_idx + 1}행)보다 아래가 아니라"
            elif _numeric_share(unit_cells) >= 0.5:
                why = f"값이 대부분 숫자라({unit_cells[:4]})"
            else:
                why = f"{reason}({unit_cells[:4]})"
            raise InputError(f"단위 행(--unit-row {unit_row})은 {why} 단위 행이 아닙니다. "
                             f"단위 행 번호를 고치거나 단위 행이 없으면 --unit-row를 빼십시오.{hint}")
        units = df_raw.iloc[unit_row - 1].values
        headers = [_with_unit(h, u) for h, u in zip(headers, units)]
        # a few Hangul words that are no known unit ('점검중' among '도씨, 킬로그램'): read as units, with a warning
        odd_units = [h for h, u in zip(headers, units) if _odd_unit_word(_cell_text(u))]
    else:
        odd_units = []

    # De-duplicate headers (handles accidental duplicate column names in source files). A generated name can
    # itself be taken (a, a, a_1), so the suffix grows until the name is unused.
    used = set(headers)
    seen_headers: Dict[str, int] = {}
    unique_headers: List[str] = []
    for h in headers:
        if h not in seen_headers:
            seen_headers[h] = 0
            unique_headers.append(h)
            continue
        n = seen_headers[h]
        while True:
            n += 1
            cand = f"{h}_{n}"
            if cand not in used:
                break
        seen_headers[h] = n
        used.add(cand)
        unique_headers.append(cand)
    headers = unique_headers

    # Determine data start row
    if data_row is not None:
        start_idx = max(0, data_row - 1)
    else:
        start_idx = max(header_idx, unit_row - 1 if unit_row is not None else header_idx) + 1

    if start_idx >= total_rows:
        raise InputError(f"데이터 시작 행 ({start_idx + 1}) 이후에 유효한 데이터 행이 없습니다.")

    df_data = df_raw.iloc[start_idx:].copy()
    df_data.columns = headers

    # Clean rows that are entirely NaN, and entirely empty columns without a header (col_<i>). A named empty
    # column stays: training then says it has no values, prediction imputes it with a warning.
    df_data = df_data.dropna(how="all", axis=0)
    if df_data.empty:   # only blank rows below the header: say so (the later messages blamed the columns)
        raise InputError(f"머리글({header_idx + 1}행) 아래의 {total_rows - start_idx}행이 모두 빈 행이라 데이터가 0행입니다. "
                         f"데이터가 채워진 파일인지, 머리글·데이터 시작 행(--header-row, --data-row)이 맞는지 확인하십시오.")
    # The header (and unit) row repeated inside the data (page breaks of an exported report) carries no values:
    # dropped, with the row numbers in attrs (layout_warnings), so numeric columns stay numeric.
    repeated = _repeated_header_rows(df_raw, df_data, header_idx, unit_row, start_idx)
    if repeated:
        df_data = df_data.drop(index=repeated)
    empty_auto = [h for h, auto in zip(headers, auto_named) if auto and df_data[h].isna().all()]
    if empty_auto:
        df_data = df_data.drop(columns=empty_auto)
    # 1-based row number in the source file for each data row (exact for .xlsx; CSV blank lines are skipped by the reader)
    source_rows = [int(i) + 1 for i in df_data.index]
    df_data = df_data.reset_index(drop=True)
    df_data.attrs["source_rows"] = _SharedList(source_rows)
    # top rows of the file that could be the real header, so a missing column name can be looked up there
    # ([추천] --header-row N): not the header row in use, and only rows above the data or mostly text rows
    # (a data row with a text value such as '압력' in a category column is not a header)
    # Each candidate: [row, cells, unit row below it or None, cells joined with that unit row ('반응온도(°C)')]
    candidates = []
    for i in range(min(HEADER_SCAN_ROWS, total_rows)):
        if i == header_idx:
            continue
        cells = _row_cells(df_raw, i)
        if cells and (i < start_idx or _numeric_share(cells) < 0.5):
            below = _row_cells(df_raw, i + 1) if i + 1 < total_rows else []
            unit_next = i + 2 if below and _looks_like_unit_row(df_raw, i + 1) else None
            joined = ([_with_unit(_cell_text(h), u) for h, u in zip(df_raw.iloc[i].values, df_raw.iloc[i + 1].values)
                       if _cell_text(h)] if unit_next else [])
            candidates.append([i + 1, cells, unit_next, joined])
    df_data.attrs["header_candidates"] = _SharedList(candidates)
    df_data.attrs["header_row"] = header_idx + 1
    df_data.attrs["unit_row"] = unit_row
    df_data.attrs["data_row"] = data_row
    # read without layout options: is the row right below the header a unit row (predicted as data otherwise)?
    df_data.attrs["unit_row_below"] = (header_idx + 2 if unit_row is None and data_row is None
                                       and header_idx + 1 < total_rows and _looks_like_unit_row(df_raw, header_idx + 1, header_idx)
                                       else None)
    df_data.attrs["repeated_header_rows"] = [int(i) + 1 for i in repeated]
    df_data.attrs["odd_unit_names"] = odd_units
    df_data.attrs["summary_rows"] = _SharedList(_summary_rows(df_data, source_rows))
    df_data.attrs["decimal_comma"] = df_raw.attrs.get("decimal_comma")
    return df_data


# A text cell that names a totals / average row (spaces removed, lower case). Totals words always mark a totals row;
# average words can also be values of a category ('집계' = 평균/최대) and are judged per column (_summary_rows).
TOTAL_WORDS = {"합계", "총계", "소계", "계", "총합", "total", "sum", "subtotal", "grandtotal"}
AVERAGE_WORDS = {"평균", "average", "avg", "mean"}
# Wording of the same rows in daily reports: '일계', '월계', '누계', '총 합계', '전체합계', '합계:', 'Total:',
# '소계(A)', '주간계', '평균값', 'AVG.'. Exact words: TOTAL_WORDS, AVERAGE_WORDS and the period totals below.
# Loose words ('...합계', short '...계' other than ordinary words, 'Total'/'Sum' as a separate word, '평균값'/'평균치')
# are also values of a category ('A계'/'B계', '생산계', 'pH계'): _summary_rows counts a loose totals word only when
# the row's numbers are totals (_loose_by_value), a loose average word when its rows are few between the data
# (LOOSE_MAX_TEXTS).
_TOTAL_PERIOD_WORDS = {"일계", "월계", "주계", "연계", "년계", "누계"}
_GYE_NOT_TOTAL = ("단계", "설계", "경계", "관계", "세계", "시계", "기계", "회계", "체계", "한계", "업계", "생계", "통계",
                  "학계", "가계", "직계", "외계", "하계", "동계", "춘계", "추계", "도계", "력계", "량계", "압계", "류계",
                  "공계", "위계", "중계", "각계", "조계", "어계", "수계", "면계", "전계")
_AVG_SUFFIXES = ("값", "치")     # '평균값', '평균치'; '평균온도' is a name, not an average row
_SUMMARY_PAREN_RE = re.compile(r"\([^)]*\)|\[[^\]]*\]")
_TOTAL_EN = {"total", "sum", "subtotal", "grandtotal"}
LOOSE_MAX_TEXTS = 2       # few loose rows: at most this many texts, <= other filled cells + 1, < half of the rows


@functools.lru_cache(maxsize=8192)
def _word_kind(text: str) -> str:
    """'total', 'avg', 'agg' (AGGREGATE_WORDS), 'total_loose', 'avg_loose' or '' for a text cell of a totals /
    average / statistic row. Spaces, parenthesized parts and a trailing ':' / '.' are ignored ('총 합계', '소계(A)',
    'Total:', 'AVG.'). 'total' / 'avg': an exact word; '_loose': a word that can also be a category value."""
    low = _SUMMARY_PAREN_RE.sub(" ", text.lower())
    key = "".join(low.split()).rstrip(":.：")
    if not key:
        return ""
    if key in TOTAL_WORDS or key in _TOTAL_PERIOD_WORDS:
        return "total"
    if key in AVERAGE_WORDS:
        return "avg"
    if key in AGGREGATE_WORDS:
        return "agg"
    words = [w.strip(":.：") for w in low.split()]
    if (key.endswith("합계") or (len(key) <= 4 and key.endswith("계") and not key.endswith(_GYE_NOT_TOTAL))
            or (len(key) <= 15 and any(w in _TOTAL_EN for w in words))):
        return "total_loose"
    if len(key) == 3 and key.startswith("평균") and key.endswith(_AVG_SUFFIXES):
        return "avg_loose"
    return ""


def _repeated_header_rows(df_raw: pd.DataFrame, df_data: pd.DataFrame, header_idx: int, unit_row: Optional[int],
                          start_idx: int) -> List[Any]:
    """Index labels (0-based raw rows) of df_data rows identical to the header row (2+ names), and of rows identical
    to the unit row right after such a row (the --unit-row, or the row below the header skipped by --data-row)."""
    head = [_cell_text(v) for v in df_raw.iloc[header_idx].tolist()]
    filled = [k for k, t in enumerate(head) if t]
    if len(filled) < 2 or df_data.empty:
        return []
    k0 = filled[0]
    col = df_data.iloc[:, k0].astype("string").str.strip()
    hits = col[(col == head[k0]).fillna(False).to_numpy(dtype=bool)].index
    u = unit_row - 1 if unit_row is not None else (header_idx + 1 if header_idx + 1 < start_idx else None)
    unit = [_cell_text(v) for v in df_raw.iloc[u].tolist()] if u is not None and 0 <= u < len(df_raw) else None
    out: List[Any] = []
    for i in hits:
        if [_cell_text(v) for v in df_raw.iloc[i].tolist()] != head:
            continue
        out.append(i)
        if unit and any(unit) and i + 1 in df_data.index and \
                [_cell_text(v) for v in df_raw.iloc[i + 1].tolist()] == unit:
            out.append(i + 1)
    return out


# Other words of an aggregation-method category ('집계' = 평균/최대/최소): a column holding them besides a totals word
# names how each row was measured, not totals rows
AGGREGATE_WORDS = {"최대", "최소", "최고", "최저", "중앙값", "중간값", "표준편차", "편차", "누적", "max", "min", "maximum",
                   "minimum", "median", "std", "stdev"}
SUMMARY_SHARE = 0.2       # average words in this share of a column's filled cells (3+ cells) make it a category


def exclude_mark(v: Any) -> str:
    """A '제외' cell as text; 0 / FALSE (0-1 or TRUE/FALSE flags) mean "not marked", like an empty cell."""
    if v is None or (not isinstance(v, str) and pd.isna(v)) or isinstance(v, (bool, np.bool_)) and not v:
        return ""
    t = str(v).strip()
    if t.upper() == "FALSE":
        return ""
    try:
        return "" if float(t) == 0 else t
    except ValueError:
        return t


def _summary_key(v: Any) -> str:
    return "".join(v.split()).lower() if isinstance(v, str) else ""


def _summary_kind(v: Any) -> str:
    return _word_kind(v.strip()) if isinstance(v, str) else ""


def _summary_rows(df: pd.DataFrame, source_rows: List[int]) -> List[List[Any]]:
    """[source row, column, text] of rows that look like totals / averages: a cell 합계/소계/Total ... (TOTAL_WORDS) or
    평균/Mean (AVERAGE_WORDS) and a number in the same row. Not dropped (a data value decision): layout_warnings names
    them. Not counted: rows the user marked in the '제외' column (a mark other than '포함'), the '제외' column itself
    (its text is the reason), words of a statistic label column, loose totals words (_word_kind 'total_loose':
    '일합계', '총 합계', '주간계', 'Grand Total') whose row's numbers are not totals (_loose_by_value: a category
    'A계'/'B계', '생산계'/'품질팀', 'pH계' among sensors) and loose average words ('평균값') of a column where they are
    not few between the data (more than the other filled cells + 1, half of the rows or more, or more than
    LOOSE_MAX_TEXTS distinct texts). Exact totals words always count (a subtotal per shift
    or line is a large share of a column), except in a statistic label column: filled in half of the rows or more,
    and totals and other aggregation words (AGGREGATE_WORDS, 2+ cells) make up half of its filled cells or more, or
    every filled cell is such a word or an average word ('집계' = 합계/평균/최대/최소, 'Sum/Avg'). Average words of a category column (next to other
    filled cells in a majority of the column's filled cells or SUMMARY_SHARE of them (3+ cells), in every row, or next
    to aggregation words) count only when the row's number cells equal the mean of the data rows directly above it
    (_summary_by_value: 'A조, B조, C조, 평균'); other average words always count. Reads the cells as one array: a Series
    of df would deep-copy df.attrs (source_rows) on every operation."""
    if df.empty:
        return []
    columns = list(df.columns)
    arr = df.to_numpy(dtype=object)
    kinds = list(df.dtypes)
    k_excl = next((k for k, c in enumerate(columns) if strip_unit(str(c)) == EXCLUDE_COL), None)
    marked = (np.array([bool(t) and t != INCLUDE_MARK for t in map(exclude_mark, arr[:, k_excl])], dtype=bool)
              if k_excl is not None else np.zeros(len(arr), dtype=bool))
    out: Dict[int, List[Any]] = {}
    nums: List[Any] = []      # number matrix of the other columns, read once when a value check is needed
    for k, col in enumerate(columns):
        if k == k_excl or pd.api.types.is_numeric_dtype(kinds[k]) or pd.api.types.is_datetime64_any_dtype(kinds[k]):
            continue
        cells = arr[:, k]
        try:          # number cells are never words: only the others are read one by one
            cand = np.flatnonzero(pd.isna(pd.to_numeric(cells, errors="coerce")))
        except (TypeError, ValueError):
            cand = np.arange(len(cells))
        kind = {int(i): _summary_kind(cells[i]) for i in cand}
        if not any(kind.values()):
            continue
        n_filled = len(cells) - len(cand) + sum(
            1 for i in kind if _summary_key(cells[i]) or not (isinstance(cells[i], str) or pd.isna(cells[i])))
        loose = [i for i, t in kind.items() if t.endswith("_loose")]
        if loose:
            # few loose rows of 1-2 texts between the data ('총 합계' below the data, '주간계'/'월간계'); many rows or
            # 3+ texts ('A계'/'B계'/'C계', '생산계', 'pH계' in every row) are a category
            few = (len(loose) <= max(1, n_filled - len(loose) + 1) and len(loose) * 2 < len(cells)
                   and len({_summary_key(cells[i]) for i in loose}) <= LOOSE_MAX_TEXTS)
            if not nums:
                nums.append(_number_matrix(arr, kinds, k_excl))
            sure = _loose_by_value(nums[0], arr, k, kind, few)
            for i in loose:
                kind[i] = kind[i][:-len("_loose")] if (i in sure if kind[i] == "total_loose" else few) else ""
        hits = [i for i, t in kind.items() if t == "total"]
        avg = [i for i, t in kind.items() if t == "avg"]
        if not hits and not avg:
            continue
        n_stat = sum(1 for t in kind.values() if t in ("total", "agg"))
        # 통계 이름 열은 집계·평균 글자가 합계 글자만큼 있어야 함 (합계/평균/최대/최소 묶음, Sum/Avg 번갈아).
        # 병합된 일자 열의 일별 합계 + 맨 아래 최대/최소 한 묶음은 통계 이름 열이 아님
        n_aggavg = sum(1 for t in kind.values() if t in ("agg", "avg"))
        if n_stat >= 2 and n_filled * 2 >= len(cells) and n_aggavg * 2 >= len(hits) and n_aggavg > 0 and (
                n_stat * 2 >= n_filled or sum(1 for t in kind.values() if t) >= n_filled):
            hits = []                 # a statistic label column: its totals words name how a row was measured
        if avg:
            n_avg = len(avg)
            share = n_avg * 2 > n_filled or (n_avg >= 3 and n_avg >= SUMMARY_SHARE * n_filled)
            aggregate = any(t == "agg" for t in kind.values())
            if aggregate or (n_filled > n_avg and share) or (n_avg == len(cells) and n_avg >= 2):
                is_bound = np.zeros(len(cells), dtype=bool)
                is_bound[[i for i, t in kind.items() if t]] = True
                starts = _block_starts(is_bound)
                avg = [i for i in avg if starts[i][1] - starts[i][0] >= 2]   # 2+ data rows directly above
                if avg and not nums:
                    nums.append(_number_matrix(arr, kinds, k_excl))
                avg = [i for i in avg if _summary_by_value(nums[0], arr, k, i, starts[i])]
        for i in sorted(hits + avg):
            if i in out or marked[i]:
                continue
            others = [v for c, v in enumerate(arr[i]) if c != k and c != k_excl]
            if any(_numeric_share([_cell_text(v)]) == 1 for v in others if _cell_text(v)):
                out[i] = [source_rows[i] if i < len(source_rows) else i + 1, str(col), str(cells[i]).strip()]
    return [out[i] for i in sorted(out)]


def _number_matrix(arr: np.ndarray, kinds: List[Any], k_excl: Optional[int]) -> np.ndarray:
    """Numbers of the cells (NaN: text, blank, date) for the value check of _summary_rows, read like the model inputs
    (coerce_numeric: '2,452' of a CSV saved with a thousands format is 2452)."""
    out = np.full(arr.shape, np.nan)
    for c in range(arr.shape[1]):
        if c == k_excl or pd.api.types.is_datetime64_any_dtype(kinds[c]) or pd.api.types.is_bool_dtype(kinds[c]):
            continue
        try:
            col = pd.Series(arr[:, c])
            v = pd.to_numeric(col, errors="coerce").to_numpy(dtype=float)
            if np.isnan(v).any() and col.map(lambda t: isinstance(t, str) and ("," in t or "%" in t)).any():
                v = coerce_numeric(col).to_numpy(dtype=float)
            out[:, c] = v
        except (TypeError, ValueError):
            continue
    return out


LOOSE_WALK = 1000         # earlier summary rows read back for one loose row ('월합계' = sum of the '일합계' rows above)


def _sum_mean_hits(rows: np.ndarray, sums: np.ndarray, count: np.ndarray, size: np.ndarray,
                   varies: Optional[np.ndarray] = None,
                   nz: Optional[np.ndarray] = None) -> Tuple[np.ndarray, np.ndarray, np.ndarray]:
    """Per row (2-D arrays, one row per candidate): the number cells equal to the sums of a block of rows (column
    sums, number counts, sums of absolute values), the other number cells equal to the mean of the block (a
    totals row that averages 온도/압력/속도 and sums only 생산량), and the cells equal to the value of a column
    that is constant in the block (a setpoint, also an all-zero column: its own mean in any block, also in a
    category row), within the precision of _summary_by_value (half a unit of the last shown decimal or 1e-3
    relative). A sum or mean column counts only with 2+ numbers in the block, not all zero: one row or a zero
    column is its own sum. varies (same shape, None: all True): the column is not constant in the block; a mean cell
    counts only there. nz (same shape, None: no check): nonzero numbers of the column in the block; a sum cell
    counts only with 2+ (a 0/1 flag with one 1 is its own sum, as a category row holds it; such a cell is no mean
    cell either). Constant cells are counted only when varies and nz are given (the block reads (a)/(c) of
    _loose_by_value; the read-back of earlier summary rows keeps them as misses)."""
    filled = ~np.isnan(rows)
    use = filled & (count >= 2) & (size > 0)
    none = np.zeros(len(rows), dtype=int)
    with np.errstate(invalid="ignore"):
        dec = 0.5 * 10.0 ** -_shown_decimals(rows)
        is_sum = use & (np.abs(rows - sums) <= np.maximum(dec, 1e-3 * np.abs(sums)) + 1e-9)
        mean = sums / np.where(count > 0, count, 1)
        near_mean = filled & (count >= 2) & (np.abs(rows - mean) <= np.maximum(dec, 1e-3 * np.abs(mean)) + 1e-9)
    is_mean = use & ~is_sum & near_mean
    if varies is not None:
        is_mean &= varies
    # a constant column (also all zero) is its own mean: no evidence either way, left out of the count
    n_const = (near_mean & ~varies).sum(axis=1) if varies is not None and nz is not None else none
    if nz is not None:
        is_sum &= nz >= 2
    return is_sum.sum(axis=1), is_mean.sum(axis=1), n_const


def _sums_match(rows: np.ndarray, sums: np.ndarray, count: np.ndarray, size: np.ndarray,
                varies: Optional[np.ndarray] = None, nz: Optional[np.ndarray] = None) -> np.ndarray:
    """Per row: at least one number cell is a block sum (_sum_mean_hits) and the sums together with the cells equal
    to the mean of a column that varies in the block (neutral: an averaged column) are at least half of the row's
    number cells. Cells equal to the value of a column constant in the block (a setpoint repeated in the totals row
    '일합계 185 3.5 1080') are left out of the count: they say nothing about sum or category."""
    hit, n_mean, n_const = _sum_mean_hits(rows, sums, count, size, varies, nz)
    return (hit >= 1) & ((hit + n_mean) * 2 >= (~np.isnan(rows)).sum(axis=1) - n_const)


def _shown_decimals(values: np.ndarray) -> np.ndarray:
    """Decimals of numbers as shown (1.5 -> 1, 300.0 -> 0, up to 9); NaN -> 0."""
    out = np.zeros(values.shape)
    v = np.where(np.isnan(values), 0.0, np.abs(values))
    open_ = np.ones(values.shape, dtype=bool)
    for d in range(10):
        x = v * 10.0 ** d
        done = open_ & (np.abs(x - np.round(x)) <= 1e-6 * np.maximum(x, 1.0))
        out[done] = d
        open_ &= ~done
    out[open_] = 9
    return out


def _loose_by_value(num: np.ndarray, arr: np.ndarray, k: int, kind: Dict[int, str], few: bool) -> set:
    """Rows of loose totals words ('일합계', '총 합계', '주간계', 'Grand Total') in column k that are totals by their
    numbers. Confirmed when at least one number cell and, together with the cells equal to the block mean (an
    averaged column, neutral), at least half of the row's number cells are (a) the sums of the data rows
    directly above it (_block_starts), (b) the sums of the earlier summary rows of one text since the last row of its
    own text ('월합계' = the '일합계' rows of the month), or (c) the sums of all data rows above it. When the loose
    rows are few between the data (few: 1-2 texts, not more than the other filled cells + 1, under half of the
    rows), a row whose numbers are far above the data of their columns (above max + max(range, max / 2) of 2+
    numbers >= 0: '총 합계 999' below 300..302) also counts, and a row without such numbers counts. A category
    ('생산계'/'품질팀', 'pH계' among sensors, 'A계'/'B계') holds numbers like its data rows, not their sums.
    One pass of array operations for (a), (c) and the far check; (b) reads back at most LOOSE_WALK summary rows."""
    loose = np.array(sorted(i for i, t in kind.items() if t == "total_loose"), dtype=int)
    if not len(loose):
        return set()
    n = len(arr)
    is_bound = np.zeros(n, dtype=bool)
    is_bound[[i for i, t in kind.items() if t]] = True
    starts = np.array(_block_starts(is_bound), dtype=int)
    other = num.copy()
    other[:, k] = np.nan
    data = np.where(is_bound[:, None], np.nan, other)
    filled = ~np.isnan(data)
    zero = np.zeros((1, data.shape[1]))
    csum = np.vstack([zero, np.cumsum(np.where(filled, data, 0.0), axis=0)])
    cabs = np.vstack([zero, np.cumsum(np.where(filled, np.abs(data), 0.0), axis=0)])
    ccnt = np.vstack([zero, np.cumsum(filled, axis=0)])
    cnz = np.vstack([zero, np.cumsum(filled & (data != 0), axis=0)])
    # does a column vary in a block: sums of squares about one value of the column (exact 0 for a constant block)
    first = data[filled.argmax(axis=0), np.arange(data.shape[1])] if len(data) else np.zeros(data.shape[1])
    dev = np.where(filled, data - np.where(np.isnan(first), 0.0, first)[None, :], 0.0)
    cdev = np.vstack([zero, np.cumsum(dev, axis=0)])
    csq = np.vstack([zero, np.cumsum(dev * dev, axis=0)])
    tiny = 1e-12 * csq[-1]

    def varies(a: np.ndarray, b: np.ndarray) -> np.ndarray:
        cnt = ccnt[b] - ccnt[a]
        with np.errstate(invalid="ignore", divide="ignore"):
            m2 = (csq[b] - csq[a]) - (cdev[b] - cdev[a]) ** 2 / np.where(cnt > 0, cnt, 1)
        return (cnt >= 2) & (m2 > tiny[None, :])

    rows = other[loose]
    s, e = starts[loose, 0], starts[loose, 1]
    sure = (_sums_match(rows, csum[e] - csum[s], ccnt[e] - ccnt[s], cabs[e] - cabs[s], varies(s, e), cnz[e] - cnz[s])
            | _sums_match(rows, csum[loose], ccnt[loose], cabs[loose], varies(np.zeros_like(loose), loose),
                          cnz[loose]))
    if few:
        dcnt = ccnt[-1]
        with np.errstate(all="ignore"):
            dmax = np.max(np.where(filled, data, -np.inf), axis=0)
            dmin = np.min(np.where(filled, data, np.inf), axis=0)
            far = np.where((dcnt >= 2) & (dmin >= 0) & (dmax > 0), dmax + np.maximum(dmax - dmin, dmax / 2), np.nan)
            judged = ~np.isnan(rows) & ~np.isnan(far)[None, :]
            n_far = (judged & (rows > far[None, :])).sum(axis=1)
        # no mean-neutral cells here (unlike _sums_match): without an exact sum, a category row with one large
        # column and a value near a column mean would be taken for a totals row
        sure |= ~judged.any(axis=1) | ((n_far >= 1) & (n_far * 2 >= (~np.isnan(rows)).sum(axis=1)))
    out = {int(i) for i in loose[sure]}
    totals = sorted(i for i, t in kind.items() if t in ("total", "total_loose"))
    keys = [_summary_key(arr[i, k]) for i in totals]
    for p, i in enumerate(totals):
        if i in out or kind[i] != "total_loose" or p == 0 or keys[p - 1] == keys[p]:
            continue
        picked = []
        for q in range(p - 1, max(-1, p - 1 - LOOSE_WALK), -1):
            if keys[q] == keys[p]:
                break
            if keys[q] == keys[p - 1]:
                picked.append(totals[q])
        if len(picked) >= 2:
            blk = other[picked]
            f = ~np.isnan(blk)
            sums = np.where(f, blk, 0.0).sum(axis=0)
            with np.errstate(invalid="ignore"):
                if not (np.abs(other[i] - sums) <= np.maximum(0.5, 1e-3 * np.abs(sums)) + 1e-9).any():
                    continue          # no cell near a sum (the exact check below is slower)
            with np.errstate(invalid="ignore"), warnings.catch_warnings():
                warnings.simplefilter("ignore", RuntimeWarning)
                vary = (np.nanmax(blk, axis=0) != np.nanmin(blk, axis=0))[None, :]
            if _sums_match(other[i][None, :], np.where(f, blk, 0.0).sum(axis=0)[None, :], f.sum(axis=0)[None, :],
                           np.where(f, np.abs(blk), 0.0).sum(axis=0)[None, :], vary)[0]:
                out.add(i)
    return out


def _block_starts(is_bound: np.ndarray) -> List[Tuple[int, int]]:
    """For each row i: (start, end) of the data rows directly above it, in one pass. Summary rows right above i are
    skipped (a 평균 row below a 소계 row), then rows are taken back to the previous summary / aggregation row or the
    first row."""
    out: List[Tuple[int, int]] = []
    end = start = 0           # block of rows < i: rows start..end-1 (end = 1 + last data row above i)
    for i, b in enumerate(is_bound):
        out.append((start, end))
        if not b:
            if end != i:      # a data row after summary rows (or the first row): a new block
                start = i
            end = i + 1
    return out


def _summary_by_value(num: np.ndarray, arr: np.ndarray, k: int, i: int, block_range: Tuple[int, int]) -> bool:
    """Row i (an average word in column k) holds the mean of the data rows directly above it (block_range of
    _block_starts, 2+ rows). Every number cell of row i must match within its displayed precision (half a unit of
    its last decimal) or 1e-3 relative, and one column must vary in the block (a constant lot or date column is its
    own mean)."""
    block = num[block_range[0]:block_range[1]]
    if len(block) < 2:
        return False
    row = num[i].copy()
    row[k] = np.nan
    filled = ~np.isnan(block)
    count = filled.sum(axis=0)
    use = np.flatnonzero(~np.isnan(row) & (count > 0))
    if len(use) == 0:
        return False
    sub = block[:, use]
    want = np.nansum(sub, axis=0) / count[use]
    texts = [str(arr[i, c]).strip() for c in use]
    dec = np.array([len(t.split(".")[1]) if "." in t and "e" not in t.lower() else 0 for t in texts])
    tol = np.maximum(0.5 * 10.0 ** -dec, 1e-3 * np.abs(want)) + 1e-9
    if (np.abs(row[use] - want) > tol).any():
        return False
    return bool(((count[use] >= 2) & (np.nanmax(sub, axis=0) != np.nanmin(sub, axis=0))).any())


def layout_warnings(df: pd.DataFrame, targets: Optional[List[str]] = None, training: bool = True) -> List[str]:
    """Warnings about rows of the file that are not ordinary data rows (df.attrs of load_data_file):
    repeated header rows (dropped) and totals / average rows (kept: the user decides with the '제외' column in
    training and ml_map; prediction only notes them). A totals word in a target column counts only in the first
    column (a class label '계' is not a totals row)."""
    out: List[str] = []
    dec = df.attrs.get("decimal_comma")
    if dec:
        out.append(f"CSV(구분 기호 '{dec[2]}')의 숫자 쉼표를 소수점으로 읽었습니다 ({dec[1]}칸, 예: {dec[0]})")
        if dec[3]:
            out.append(f"CSV(구분 기호 '{dec[2]}')의 숫자 점(.)을 천 단위 구분 기호로 읽었습니다 ({dec[4]}칸, 예: {dec[3]})")
    odd = df.attrs.get("odd_unit_names") or []
    if odd:
        out.append(f"{df.attrs.get('unit_row')}행을 단위로 읽음: {', '.join(odd[:5])}{' 등' if len(odd) > 5 else ''}. "
                   f"상태·결측 표시로 보이는 낱말이 있습니다. 데이터 행이면 --unit-row를 빼세요")
    rep = df.attrs.get("repeated_header_rows") or []
    if rep:
        out.append(f"{_row_list(rep)}행은 머리글(단위) 행과 같아(페이지마다 반복된 머리글) 데이터에서 뺐습니다")
    first = str(df.columns[0]) if len(df.columns) else None
    rows = [r for r in df.attrs.get("summary_rows") or []
            if not targets or r[1] not in targets or r[1] == first]
    if rows:
        where = ", ".join(f"{r[0]}행 '{r[2]}'" for r in rows[:5]) + (f" 외 {len(rows) - 5}개" if len(rows) > 5 else "")
        if training:
            out.append(f"합계·평균 행으로 보이는 행이 있습니다 ({where}). 데이터 행이 아니면 데이터 파일의 '제외' 열에 "
                       f"이유(예: 합계)를 적고 다시 실행하십시오")
        else:
            out.append(f"합계·평균 행으로 보이는 행이 있습니다 ({where}). 이 행의 예측값은 데이터 행의 예측이 아닙니다")
    return out


def _row_list(rows: List[int]) -> str:
    return ", ".join(str(r) for r in rows[:10]) + (f" 외 {len(rows) - 10}개" if len(rows) > 10 else "")


def _cell_text(v: Any) -> str:
    """Header / unit cell as text: blank -> '', a whole-number Excel cell (year 2023 read as 2023.0) -> '2023'."""
    if v is None or (not isinstance(v, str) and pd.isna(v)):
        return ""
    if isinstance(v, (float, np.floating)) and float(v).is_integer():
        return str(int(v))
    return str(v).strip()


def _with_unit(header: str, unit: Any) -> str:
    u = _cell_text(unit)
    return f"{header}({u})" if not _blank_mark(u) else header


def _row_cells(df_raw: pd.DataFrame, i: int) -> List[str]:
    return [t for t in (_cell_text(v) for v in df_raw.iloc[i].tolist()) if t]


UNIT_SCAN_CELLS = 10      # filled cells below a unit that decide whether its column holds data (majority)


# Missing-value marks of measurement sheets: a blank cell for the layout judgment (unit row, first data row), never a
# unit. Data values are not changed by this (coerce_numeric reads them as missing numbers anyway). Matched with case:
# 'NG' is a mark, 'ng' is the nanogram unit.
MISSING_MARKS = set(_NA_PLACEHOLDERS) | {"측정불가", "측정 불가", "측정안됨", "결측", "없음", "해당없음", "NG", "N.G.",
                                         "N.D.", "ND", "n.d.", "#N/A", "#VALUE!", "#DIV/0!"}


def _blank_mark(t: str) -> bool:
    return not t or t in MISSING_MARKS


def _is_data_value(v: Any) -> bool:
    """A number or a date (a date cell, 'YYYY-MM-DD' text): what sits below a unit in a data column."""
    if isinstance(v, (datetime, np.datetime64)):
        return True
    t = _cell_text(v)
    return _numeric_share([t]) == 1 or bool(_DATE_RE.match(t))


def _unit_text(t: str) -> bool:
    """Short text that can be a unit ('°C', 'L/min', 'yyyy-mm-dd'); a note ('비고: 2026년 1월 측정') is not."""
    return len(t) <= 12 and t.count(" ") <= 1


_KOREAN_UNITS = {"개", "회", "초", "분", "시간", "일", "주", "월", "년", "원", "천원", "만원", "명", "건", "점", "매", "장",
                 "대", "도", "배", "번", "개월", "퍼센트", "미터", "그램", "리터", "톤", "와트", "볼트", "암페어"}


# Korean unit words of a unit row the user named (--unit-row): never a state word (_odd_unit_word). Any other short
# Hangul word there is a unit too ('마리', '킬로'); only state words ('점검중', '센서고장', '정지': a data row of a
# measurement sheet) make the row suspect: half or more of its text cells -> not a unit row, fewer -> read with a
# warning. The tool's own guess (no --unit-row) knows only _KOREAN_UNITS.
_KOREAN_UNIT_WORDS = _KOREAN_UNITS | {
    "도씨", "섭씨", "화씨", "켈빈", "킬로그램", "밀리그램", "그램", "톤", "킬로톤", "기압", "바", "밀리바", "파스칼", "킬로파스칼",
    "메가파스칼", "헥토파스칼", "미터", "킬로미터", "센티미터", "밀리미터", "마이크로미터", "마이크론", "나노미터", "인치", "피트",
    "리터", "밀리리터", "시씨", "세제곱미터", "제곱미터", "평", "헥타르", "시", "시간", "분", "초", "밀리초", "회", "회전",
    "개", "세트", "쌍", "벌", "켤레", "박스", "상자", "팩", "롤", "묶음", "포", "통", "병", "캔", "봉", "알", "정", "가지", "명",
    "퍼센트", "프로", "퍼밀", "피피엠", "와트", "킬로와트", "메가와트", "킬로와트시", "볼트", "킬로볼트", "암페어", "밀리암페어",
    "옴", "헤르츠", "킬로헤르츠", "데시벨", "뉴턴", "줄", "킬로줄", "칼로리", "킬로칼로리", "마력", "알피엠", "노트", "룩스",
    "도", "라디안", "주기", "사이클", "건", "차", "호", "층", "동", "칸", "달러", "엔", "유로", "위안",
    "세", "살", "억원", "만원", "백만원", "천원", "원", "킬로", "센티", "미리", "바이트", "킬로바이트", "메가바이트",
    "기가바이트", "테라바이트", "개수", "마리", "점", "퍼센트포인트", "배", "대",
    "킬로그램중", "그램중", "톤중",              # weight units (kgf, gf, tf): before the rule of words ending in "중"
}

# State and status words of a measurement sheet's data cells (a machine state, a missing reading): a row of them is a
# data row, not units. Also any Hangul word ending in '중' ('측정중', '가동중') or holding one of _STATE_PARTS.
_STATE_WORDS = {"정지", "가동", "운전", "대기", "정상", "이상", "불량", "양품", "미가동", "휴무", "휴지", "오류", "에러",
                "누락", "미입력", "미측정", "고장", "점검", "보류", "완료", "합격", "불합격", "판정", "확인", "재측정"}
_STATE_PARTS = ("고장", "정지", "불량", "오류", "에러", "누락", "점검", "측정", "이상", "교체", "수리", "정비")


def _hangul_word(t: str) -> bool:
    t = t.strip()
    return bool(t) and all("\uac00" <= ch <= "\ud7a3" or "\u3131" <= ch <= "\u318e" or ch.isspace() for ch in t)


def _odd_unit_word(t: str) -> bool:
    """A state word in a unit row ('점검중', '센서고장', '정지'; '도씨', '마리', '킬로' are units): the row may be data."""
    t = t.strip()
    if not _hangul_word(t) or t in _KOREAN_UNIT_WORDS:
        return False
    return t in _STATE_WORDS or (len(t) > 1 and t.endswith("중")) or any(w in t for w in _STATE_PARTS)


KOREAN_UNIT_MAX = 6      # a Hangul word of a unit row the user named (--unit-row): at most this many letters, no space


def _unit_shaped(t: str, given: bool = False) -> bool:
    """Text a unit is written as: some letter, digit or sign that is not Hangul ('kg/cm2 G', '°C', '%', 'mm'), or a
    Korean unit word ('개', '시간'). Other Hangul words ('점검중', '센서고장') are not units when the tool judges the
    row by itself; in the unit row the user named (given: --unit-row) a short Hangul word without space is a unit
    ('도씨', '킬로그램', '기압', '회전')."""
    t = t.strip()
    if not t or not _unit_text(t):
        return False
    if t in _KOREAN_UNITS or (given and len(t) <= KOREAN_UNIT_MAX and " " not in t):
        return True
    return any(not ("\uac00" <= ch <= "\ud7a3" or "\u3131" <= ch <= "\u318e" or ch.isspace()) for ch in t)


UNIT_SCAN_ROWS = 50 * UNIT_SCAN_CELLS    # rows below a unit row read for the judgment (bounded on a large sheet)


def _number(v: Any) -> Optional[float]:
    t = _cell_text(v)
    if isinstance(v, (datetime, np.datetime64)) or not t or _numeric_share([t]) != 1:
        return None
    return float(_strip_percent(t.strip()).replace(",", ""))


def _below_block(df_raw: pd.DataFrame, j: int, skip: List[List[str]]) -> np.ndarray:
    """Cells of the rows below row j (up to UNIT_SCAN_ROWS), without the rows equal to a row of `skip` (the header and
    unit row repeated on page breaks): one read for every column of the judgment."""
    block = df_raw.iloc[j + 1:j + 1 + UNIT_SCAN_ROWS].to_numpy(dtype=object)
    if not len(block):
        return block
    keep = np.ones(len(block), dtype=bool)
    for row in skip:
        k0 = next((k for k, t in enumerate(row) if t), None)
        if k0 is None or k0 >= block.shape[1]:
            continue
        for i in np.flatnonzero(np.fromiter((_cell_text(v) == row[k0] for v in block[:, k0]), dtype=bool,
                                            count=len(block))):
            if [_cell_text(x) for x in block[i]] == row:
                keep[i] = False
    return block[keep]


def _column_below(block: np.ndarray, k: int) -> List[Any]:
    """Filled cells of column k of _below_block (missing marks dropped)."""
    if not len(block) or k >= block.shape[1]:
        return []
    return [v for v in block[:, k] if not _blank_mark(_cell_text(v))]


def _unit_cell(v: Any, below: List[Any], given: bool = False) -> str:
    """One filled cell of a candidate unit row against the filled cells of its column below:
    'unit'  text that is not among the values below ('°C', a scale '1000' far outside the numbers below);
    'data'  a number or date in the range of the values below, or text that is a value below (category 'A');
    'blank' a missing-value mark ('측정불가', 'NG', '-') that is not a value below;
    'note'  long text (not a unit, not a value below)."""
    t = _cell_text(v)
    texts = {_cell_text(b) for b in below if not _is_data_value(b)}
    if not _is_data_value(v):
        if t in texts:
            return "data"
        if _blank_mark(t):
            return "blank"
        return "unit" if _unit_shaped(t, given) else "note"      # '점검중', '센서고장' are not units (unless given)
    x = _number(v)
    nums = [n for n in (_number(b) for b in below) if n is not None]
    if x is None or not nums or x == 0:
        return "data"              # a date, a number over a text column, a zero
    lo, hi = min(nums), max(nums)
    w = max(hi - lo, 0.5 * max(abs(lo), abs(hi)), 1e-9)
    return "unit" if (x > hi + 3 * w or x < lo - 3 * w) else "data"     # a scale unit: 1000 over 1..5


def unit_row_verdict(df_raw: pd.DataFrame, j: int, header: Optional[int] = None,
                     given: bool = False) -> Tuple[bool, str]:
    """(row j (0-based) is the unit row of the header row `header` (default j - 1), why not). Judged per cell, then
    by the row (_unit_cell). given: the user named the row (--unit-row), so short Hangul words count as units
    ('도씨', '킬로그램'); without it (the tool's own guess) they are data ('점검중', '센서고장' of a first data row):
    - no cell is data (a number or date in the range below, or a text value found below such as a category 'A'),
      and some text cell is a unit; missing-value marks ('측정불가', 'NG') count as blank, so a first data row of
      '측정불가, 측정불가, A' or '측정불가, 측정불가' is data, not units;
    - every filled cell sits under a filled header cell, and the header row is mostly text (a header under a title row
      of fewer cells is not a unit row); a column with no header and no numbers or dates below ('비고' over blank
      cells) is not judged;
    - most unit text cells sit over data: the majority of the next filled cells (up to UNIT_SCAN_CELLS) are numbers or
      dates, so one text value below ('측정불가', '<0.1', '12%') does not decide;
    - the next row with values in those columns is not text in all of them (a header over its unit row)."""
    header = j - 1 if header is None else header
    if header < 0 or j <= header or j + 1 >= len(df_raw):
        return False, "머리글 행 아래가 아니라"
    head = [_cell_text(v) for v in df_raw.iloc[header].tolist()]
    head_cells = [t for t in head if t]
    if not head_cells or _numeric_share(head_cells) >= 0.5:
        return False, "머리글 행이 글자 행이 아니라"
    row = df_raw.iloc[j].tolist()
    if all(_blank_mark(_cell_text(v)) for v in row):
        return False, "빈 칸과 빈 값 표시(측정불가 등)뿐이라"
    block = _below_block(df_raw, j, [head, [_cell_text(v) for v in row]])
    kinds: Dict[int, str] = {}
    for k, v in enumerate(row):
        if not _cell_text(v):
            continue
        below = _column_below(block, k)
        if k >= len(head) or not head[k]:
            if _blank_mark(_cell_text(v)):
                continue
            if any(_is_data_value(b) for b in below[:UNIT_SCAN_CELLS * 5]):   # a name under a title row
                return False, "이름 없는 머리글 칸 아래에 숫자·날짜가 든 열이 있어(제목 행 아래의 머리글)"
            continue                                    # '비고' over a column without data: not judged
        kinds[k] = _unit_cell(v, below, given)
    if any(c == "data" for c in kinds.values()):
        return False, "아래 데이터와 같은 값(그 열 범위 안의 숫자·날짜, 아래에도 있는 글자 값)이 든 데이터 행이라"
    units = [k for k, c in kinds.items() if c == "unit" and not _is_data_value(row[k])]
    text = [k for k, c in kinds.items() if c in ("unit", "note") and not _is_data_value(row[k])]
    if not units:
        return False, "단위로 볼 짧은 글자가 없어"
    odd = [_cell_text(row[k]) for k in units if _odd_unit_word(_cell_text(row[k]))]
    if odd and len(odd) * 2 >= len(text):
        return False, "상태·결측 표시로 보이는 한글 낱말(점검중, 고장, 정지 등)이 글자 칸의 절반 이상이라"
    good = 0
    for k in units:
        below = _column_below(block, k)[:UNIT_SCAN_CELLS]
        if below and sum(_is_data_value(b) for b in below) * 2 > len(below):
            good += 1
    if good * 2 <= len(text):
        return False, "단위 칸 아래 열의 값이 대부분 숫자·날짜가 아니라"
    for i in range(j + 1, len(df_raw)):
        vals = [df_raw.iat[i, k] for k in units if not _blank_mark(_cell_text(df_raw.iat[i, k]))]
        if vals:
            if len(vals) < 2 or any(_is_data_value(v) for v in vals):
                return True, ""
            return False, "바로 아래 행도 글자 행이라(단위 행 위의 머리글)"
    return False, "아래에 데이터가 없어"


def _looks_like_unit_row(df_raw: pd.DataFrame, j: int, header: Optional[int] = None) -> bool:
    """Row j (0-based) is the unit row of the header row `header` (default j - 1): unit_row_verdict."""
    return unit_row_verdict(df_raw, j, header)[0]


def _header_under_title(df_raw: pd.DataFrame, title: int, j: int) -> bool:
    """Row j (0-based) is the real header under the row `title` read as the header: 2+ text names, more cells than
    that row, and the next filled row is mostly numbers or dates (a title row '시험 기록, 작성자' over '온도, 압력, 수율')."""
    cells, top = _row_cells(df_raw, j), _row_cells(df_raw, title)
    if len(cells) < 2 or _numeric_share(cells) >= 0.5 or len(cells) <= len(top):
        return False
    for i in range(j + 1, len(df_raw)):
        vals = [v for v in df_raw.iloc[i].tolist() if not _blank_mark(_cell_text(v))]
        if vals:
            return sum(_is_data_value(v) for v in vals) * 2 >= len(vals)
    return False


def _header_unit_pair(df_raw: pd.DataFrame) -> Optional[Tuple[int, int]]:
    """(header row, unit row), 1-based: the first row of 2+ text names whose next row is its unit row
    (_looks_like_unit_row: units under the names, data below). A title row over the header is not such a pair."""
    for i in range(max(min(HEADER_SCAN_ROWS, len(df_raw) - 1), 0)):
        h = _row_cells(df_raw, i)
        if len(h) >= 2 and _numeric_share(h) < 0.5 and _looks_like_unit_row(df_raw, i + 1, i):
            return i + 1, i + 2
    return None


LAYOUT_OPTS = {"--header-row": "header_row", "--unit-row": "unit_row", "--data-row": "data_row"}
PATH_OPTS = ("-f", "--file", "-m", "--model", "--save", "--points")


def layout_retry(header_row: int, unit_row: Optional[int]) -> str:
    """The [추천] for a new header (and unit) row: the running command with --header-row/--unit-row replaced and
    --data-row dropped (several options change, so a full command, R8); option text when the running program is not
    an ml-kit command (in-process use)."""
    opts = f"--header-row {header_row}" + (f" --unit-row {unit_row}" if unit_row else "")
    cmd = command_text(drop=LAYOUT_OPTS)
    return f"{cmd} {opts}" if cmd else opts


def command_text(drop: Optional[Dict[str, str]] = None, argv: Optional[List[str]] = None) -> Optional[str]:
    """The running ml-kit command (or argv: script + arguments) as one full command line (python, absolute script and
    path options), without the options in drop. None when the program is not an ml-kit command (in-process use)."""
    if argv is None:
        argv = CURRENT_ARGV if CURRENT_ARGV is not None else list(sys.argv)
    script = os.path.basename(argv[0] or "") if argv else ""
    if not (script.startswith("ml_") and script.endswith(".py")):
        return None
    out: List[str] = []
    i = 1
    while i < len(argv):
        a = str(argv[i])
        key, has_eq = a.split("=", 1)[0], "=" in a
        hit = [o for o in (drop or {}) if o == key or (key.startswith("--") and len(key) > 3 and o.startswith(key))]
        if len(hit) == 1:
            i += 1 if has_eq else 2
            continue
        if key in PATH_OPTS and not has_eq and i + 1 < len(argv):
            out += [a, quote_arg(os.path.abspath(str(argv[i + 1]).rstrip("\\/") or str(argv[i + 1])))]
            i += 2
            continue
        out.append(quote_arg(a))
        i += 1
    return " ".join([f"python {quote_arg(os.path.abspath(argv[0]))}"] + out)


def _numeric_share(cells: List[str]) -> float:
    n = 0
    for c in cells:
        try:
            float(_strip_percent(c.strip()).replace(",", ""))
            n += 1
        except ValueError:
            pass
    return n / len(cells) if cells else 0.0


def names_in_cells(names: List[str], cells: List[str], model_names: bool = False) -> bool:
    """True when every name is one of the cells, exactly or by a unique unit-less alias written without a unit
    ('온도' ~ '온도(°C)') or the same unit written otherwise ('온도(C)' ~ '온도(°C)'), the rule of resolve_column_names.
    model_names (prediction, the align_input_columns rule): a model name with a unit also matches a cell without a
    unit ('반응온도(°C)' ~ '반응온도')."""
    cell_set = set(cells)
    alias = build_alias_map(cells)

    def _hit(n: str) -> bool:
        if n in cell_set or sum(_same_unit_name(n, c) for c in cells) == 1:
            return True
        src = alias.get(_alias_key(n))
        if src is None:
            return False
        return strip_unit(n) == n.strip() or (model_names and strip_unit(src) == src.strip())
    return bool(names) and all(_hit(n) for n in names)


def find_header_layout(candidates: Optional[List[List[Any]]], names: List[str],
                       model_names: bool = False) -> Optional[Tuple[int, Optional[int]]]:
    """(header row, unit row or None), 1-based, of the candidate row (df.attrs["header_candidates"] of load_data_file)
    that holds every name: as written ('온도'), or joined with the unit row below it ('온도(°C)' = '온도' + '°C',
    names of a model trained with --unit-row). The unit row is given whenever the row below is one (text).
    model_names: model names with a unit also match header cells without one (names_in_cells)."""
    for cand in candidates or []:
        row, cells = cand[0], cand[1]
        unit_next = cand[2] if len(cand) > 2 else None
        joined = cand[3] if len(cand) > 3 else []
        if names_in_cells(names, cells, model_names) or (joined and names_in_cells(names, joined)):
            return int(row), (int(unit_next) if unit_next else None)
    return None


def header_layout_hint(layout: Optional[Dict[str, Any]], names: List[str]) -> str:
    """'파일 N행에 이 이름들이 있습니다 ... [추천] ...' for names found in another top row, else ''.
    [추천] --header-row N alone when no unit row is involved; with a unit row below N, or when the run gave
    --unit-row/--data-row (now wrong), the full command with the new --header-row/--unit-row (layout_retry)."""
    layout = layout or {}
    found = find_header_layout(layout.get("header_candidates"), names)
    if not found:
        return ""
    row, unit = found
    where = f"파일 {row}행에 이 이름들이 있습니다 (두 줄 머리글, 병합된 머리글이나 제목 행)"
    if unit:
        where += f", {unit}행은 단위 행입니다"
    if unit is None and layout.get("unit_row") is None and layout.get("data_row") is None:
        return f"{where}. [추천] {option_retry(str(row), '--header-row')}"
    return f"{where}. [추천] {layout_retry(row, unit)}"


# ==============================================================================
# C4-B. Feature Schema (numeric / categorical / datetime), identical at train and inference
# ==============================================================================
#
# feature_schema (stored in vX/model_<target>.json):
#   {"numeric": [...], "categorical": {col: [known categories]},
#    "datetime": {col: [parts]}, "input_columns": [...model input column order...]}
# Datetime columns are expanded OUTSIDE the sklearn pipeline (no custom code is pickled,
# so ml_predict.py stays standalone); categorical columns are one-hot encoded INSIDE the
# pipeline by sklearn's OneHotEncoder.

DATE_PARTS = ["year", "month", "day", "dayofweek", "hour"]
_DATE_RE = re.compile(r"^\s*\d{4}[-/.]\d{1,2}[-/.]\d{1,2}([ T]\d{1,2}:\d{2}(:\d{2}(\.\d+)?)?)?\s*$")
_DATE8_RE = re.compile(r"^(\d{4})(\d{2})(\d{2})$")


def parse_datetime(series: pd.Series) -> pd.Series:
    """
    The one date parser of training (--dtypes col:datetime, auto detection, datetime_stats) and prediction:
    Excel datetimes, 'YYYY-MM-DD[ HH:MM[:SS]]' strings (also / and . separators) and 8-digit YYYYMMDD
    (number or text, e.g. 20240103). Anything else is NaT.
    """
    if pd.api.types.is_datetime64_any_dtype(series):
        return series

    def _one(v):
        if v is None or (isinstance(v, float) and np.isnan(v)):
            return pd.NaT
        if isinstance(v, (datetime, pd.Timestamp, np.datetime64)):
            return pd.Timestamp(v)
        if isinstance(v, (float, np.floating)) and float(v).is_integer():
            v = int(v)
        t = str(v).strip()
        m8 = _DATE8_RE.match(re.sub(r"\.0+$", "", t))     # '20240103.0' (text of a number cell) like 20240103
        if m8:
            try:
                return pd.Timestamp(year=int(m8.group(1)), month=int(m8.group(2)), day=int(m8.group(3)))
            except (ValueError, TypeError):
                return pd.NaT
        if not _DATE_RE.match(t):
            return pd.NaT
        try:
            return pd.Timestamp(t.replace("/", "-").replace(".", "-", 2))
        except (ValueError, TypeError):
            return pd.NaT

    return pd.to_datetime(series.astype("object").map(_one), errors="coerce")


def date_part_column(col: str, part: str) -> str:
    return f"{col}_{part}"


def normalize_category(series: pd.Series) -> pd.Series:
    """Category values as trimmed strings; placeholders become NaN (imputed as '__missing__')."""
    def _one(v):
        if v is None or (isinstance(v, float) and np.isnan(v)):
            return np.nan
        if isinstance(v, (float, np.floating)) and float(v).is_integer():
            v = int(v)
        t = str(v).strip()
        return np.nan if t in _NA_PLACEHOLDERS else t
    return series.astype("object").map(_one).astype("object")


CATEGORY_CANDIDATES_MAX = 10    # [후보] lists at most this many training categories (R8)


def _category_key(v: Any) -> str:
    """The form two category values are matched in: full-width -> half-width (NFKC), trimmed, case-folded."""
    import unicodedata
    return unicodedata.normalize("NFKC", str(v)).strip().casefold()


def match_categories(values: Any, known: List[Any]) -> Tuple[Dict[str, str], List[str]]:
    """Given category values against the training categories: ({value: training category} for the values that differ
    from exactly one training category only by case, surrounding spaces or full-width letters, [values matching none]).
    A value equal to a training category is kept; one matching two training categories ('a' for 'A' and 'a') is unseen."""
    known_s = [str(k) for k in known]
    by_key: Dict[str, List[str]] = {}
    for k in known_s:
        by_key.setdefault(_category_key(k), []).append(k)
    fixes: Dict[str, str] = {}
    unseen: List[str] = []
    for v in dict.fromkeys(str(x) for x in values):
        if v in known_s:
            continue
        hit = by_key.get(_category_key(v), [])
        if len(hit) == 1:
            fixes[v] = hit[0]
        else:
            unseen.append(v)
    return fixes, unseen


def category_candidates(known: List[Any]) -> str:
    """'[후보] A, B, C' of the training categories (at most CATEGORY_CANDIDATES_MAX, the rest counted)."""
    known = [str(k) for k in known]
    text = f"[후보] {', '.join(known[:CATEGORY_CANDIDATES_MAX])}"
    if len(known) > CATEGORY_CANDIDATES_MAX:
        text += f" (학습 범주 {len(known)}개 중 {CATEGORY_CANDIDATES_MAX}개)"
    return text


def _category_recovered(col: str, fixes: Dict[str, str]) -> None:
    for given, k in fixes.items():
        note_recovered(f"'{col}' 값 '{given}'이(가) 학습 범주와 대소문자·공백·전각만 달라 -> '{k}'(으)로 맞춰 예측")


def align_categories(df: pd.DataFrame, schema: Dict[str, Any]) -> Dict[str, List[str]]:
    """Category inputs of df (file or -i) matched to the training categories in place (match_categories, reported with
    note_recovered). Returns {column: values matching no training category} (the caller warns about them)."""
    left: Dict[str, List[str]] = {}
    for col, known in (schema.get("categorical") or {}).items():
        if col not in df.columns:
            continue
        s = normalize_category(df[col])
        fixes, unseen = match_categories(s.dropna(), known)
        if fixes:
            df[col] = s.map(lambda v: fixes.get(v, v))
            _category_recovered(col, fixes)
        if unseen:
            left[col] = unseen
    return left


def fixed_category_value(col: str, value: Any, known: List[Any], option: str = "-i") -> str:
    """One fixed category input (--sweep / --maximize / --find): the training category it means, or STOP_RETRY with the
    training categories as [후보] (an unseen category would be predicted as the baseline without a word)."""
    v = normalize_category(pd.Series([value])).iloc[0]
    if isinstance(v, str):
        fixes, unseen = match_categories([v], known)
        if not unseen:
            _category_recovered(col, fixes)
            return fixes.get(v, v)
    raise InputError(f"{option}의 '{col}={value}'은(는) 학습 때 없던 값이라 예측할 수 없습니다. "
                     f"학습 범주 중 하나로 다시 지정하십시오. {category_candidates(known)}")


def prepare_features(df: pd.DataFrame, schema: Dict[str, Any]) -> pd.DataFrame:
    """Builds the model input frame (column order = schema['input_columns'])."""
    out: Dict[str, pd.Series] = {}
    for col in schema.get("numeric", []):
        out[col] = coerce_numeric(df[col]).astype(float)
    for col, parts in schema.get("datetime", {}).items():
        dt = parse_datetime(df[col])
        for part in parts:
            out[date_part_column(col, part)] = getattr(dt.dt, part).astype(float)
    for col in schema.get("categorical", {}):
        out[col] = normalize_category(df[col])
    frame = pd.DataFrame(out, index=df.index)
    return frame[schema["input_columns"]]


def input_column_owner(schema: Dict[str, Any]) -> Dict[str, str]:
    """Maps each model input column back to its raw column (date parts -> date column)."""
    owner = {c: c for c in schema.get("numeric", [])}
    owner.update({c: c for c in schema.get("categorical", {})})
    for col, parts in schema.get("datetime", {}).items():
        for part in parts:
            owner[date_part_column(col, part)] = col
    return owner


# ==============================================================================
# C5. Bulletproof AST Formula Sandbox (--eval), shared by training and inference
# ==============================================================================

_ALLOWED_AST_NODES = (
    ast.Expression, ast.BinOp, ast.UnaryOp, ast.Constant, ast.Name,
    ast.Load, ast.Call, ast.Add, ast.Sub, ast.Mult, ast.Div,
    ast.FloorDiv, ast.Mod, ast.Pow, ast.USub, ast.UAdd
)

SAFE_MATH: Dict[str, Any] = {
    "log": lambda x: np.log(np.clip(np.abs(x), 1e-9, None)),
    "log1p": lambda x: np.log1p(np.clip(np.abs(x), 0, None)),
    "log10": lambda x: np.log10(np.clip(np.abs(x), 1e-9, None)),
    "sqrt": lambda x: np.sqrt(np.clip(np.abs(x), 0, None)),
    "exp": lambda x: np.exp(np.clip(x, -50, 50)),
    "abs": np.abs,
    "sin": np.sin,
    "cos": np.cos,
    "tan": np.tan,
    "square": np.square,
    "clip": lambda x, a_min, a_max: np.clip(x, a_min, a_max),
    "pi": np.pi,
    "e": np.e,
    "radians": np.radians,
    "degrees": np.degrees,
} if np is not None else {}      # numpy missing: check_core_libraries answers before any formula (R5)

_NUMBER_NAME_RE = re.compile(r"(\d+\.?\d*|\.\d+)([eE][-+]?\d+)?")    # a column name that reads as a number
_NAME_BOUNDARY_L = r"(?<![가-힣a-zA-Z0-9_])"
_NAME_BOUNDARY_R = r"(?![가-힣a-zA-Z0-9_])"
_NUMBER_START_RE = re.compile(r"\.?\d")    # a name starting like a number ('2(mm)', '0.5(mm)', '2')
_NUMBER_BOUNDARY_L = r"(?<![가-힣a-zA-Z0-9_.])"   # ... is never read inside a larger number ('0.5' / '5(mm)')
_NUMBER_BOUNDARY_R = r"(?![가-힣a-zA-Z0-9_.])"    # ('2.5' / '2')


def _validate_eval_ast(formula_ast: ast.AST, allowed_identifiers: set, columns: List[str],
                       targets: Optional[set] = None) -> None:
    """Ensures AST contains only safe arithmetic nodes and declared identifiers. An unknown name is never matched to
    a target (a formula cannot read it): a target as near as the other columns leaves [후보] only."""
    banned = set(targets or [])
    for node in ast.walk(formula_ast):
        if not isinstance(node, _ALLOWED_AST_NODES):
            raise InputError(f"보안 위반: 허용되지 않은 문법 노드 감지: {type(node).__name__}")
        if isinstance(node, ast.Call) and not (isinstance(node.func, ast.Name) and node.func.id in SAFE_MATH):
            raise InputError(f"허용되지 않은 함수 호출입니다. 사용 가능 함수: {sorted(k for k in SAFE_MATH if callable(SAFE_MATH[k]))}")
        if isinstance(node, ast.Name) and node.id not in allowed_identifiers:
            s = suggest_columns(node.id, [c for c in columns if c not in banned])
            near_t = [c for c in suggest_columns(node.id, columns) if c in banned]
            if near_t:
                other = f"비슷한 {', '.join(near_t)}은(는) 타겟이라 수식에 쓸 수 없음"
                hint = (f" -> [후보] {', '.join(s)} ({other}. 이름을 사용자에게 확인)" if s else
                        f" ({other}. 이름을 사용자에게 확인)")
            else:
                hint = f" -> [후보] {', '.join(s)} (어느 열인지 사용자에게 확인)" if s else ""
            raise InputError(
                f"수식의 식별자 '{node.id}'에 해당하는 컬럼이 없습니다"
                + hint
                + f". 컬럼명은 데이터의 전체 이름(단위 포함) 또는 단위를 뺀 고유 이름으로 쓰십시오. [사용 가능한 컬럼 목록]: {name_list(columns)}"
            )


def compile_formula(formula: str, columns: List[str]) -> Tuple[str, Dict[str, str], str]:
    """
    Replaces column names in a formula with internal tokens (prevents collisions with
    function names and lets names contain '(', '%', '/', etc.).
    Full names are matched first (longest first), then unique unit-less aliases.
    Returns (tokenized_formula, token_to_column, canonical_formula_with_full_names).
    """
    col_list = [str(c) for c in columns]
    # a column named like a number ('1', '2'): the number in the formula ('*2') would silently read that column.
    # The name without unit counts too: '2' of '2(mm)' would be its alias. A full name written with its unit
    # ('2(mm)') is that column; a number inside a larger number ('2.5', '0.5') is no name (_NUMBER_BOUNDARY_*)
    rest = formula
    for c in sorted(col_list, key=len, reverse=True):
        if _NUMBER_START_RE.match(c.strip()) and not _NUMBER_NAME_RE.fullmatch(c.strip()):
            rest = re.sub(_NUMBER_BOUNDARY_L + re.escape(c.strip()) + _NAME_BOUNDARY_R, " ", rest)
    for c in col_list:
        for num in dict.fromkeys([c.strip(), strip_unit(c)]):
            if _NUMBER_NAME_RE.fullmatch(num) and \
                    re.search(_NUMBER_BOUNDARY_L + re.escape(num) + _NUMBER_BOUNDARY_R, rest):
                raise InputError(
                    f"데이터에 숫자 이름의 열 '{c}'이(가) 있어 수식 '{formula}'의 숫자 {num}이(가) 숫자인지 그 열"
                    f"인지 알 수 없습니다 (수식은 열 이름을 따옴표나 괄호로 감싸 쓰는 방법이 없음). 사용자에게 알려 데이터의 "
                    f"열 '{c}' 이름을 글자가 든 이름(예: '열{c.strip()}')으로 바꾼 뒤 다시 실행하십시오.")
    token_to_col: Dict[str, str] = {}
    transformed = formula
    idx = 0

    candidates: List[Tuple[str, str]] = [(c, c) for c in sorted(col_list, key=len, reverse=True)]
    alias_map = build_alias_map(col_list)
    col_set = set(col_list)
    aliases = [(strip_unit(full), full) for key, full in alias_map.items()]
    aliases = [(a, full) for a, full in aliases if a and a != full and a not in col_set and a not in SAFE_MATH
               and not _NUMBER_NAME_RE.fullmatch(a)]    # an alias that reads as a number is never a formula name
    candidates += sorted(aliases, key=lambda t: len(t[0]), reverse=True)

    for name, full in candidates:
        pattern = (_NUMBER_BOUNDARY_L if _NUMBER_START_RE.match(name) else _NAME_BOUNDARY_L) + re.escape(name) \
            + (_NUMBER_BOUNDARY_R if _NUMBER_NAME_RE.fullmatch(name) else _NAME_BOUNDARY_R)
        if re.search(pattern, transformed):
            token = next((t for t, c in token_to_col.items() if c == full), None)
            if token is None:
                token = f"__COL_TOKEN_{idx}__"
                idx += 1
                token_to_col[token] = full
            transformed = re.sub(pattern, lambda _m, t=token: t, transformed)

    canonical = transformed
    for token, col in token_to_col.items():
        canonical = canonical.replace(token, col)
    return transformed, token_to_col, canonical


def eval_source_columns(evals: List[str], columns: List[str]) -> List[str]:
    """Data columns that derived formulas read (in first-use order). Names created by earlier formulas are not
    sources. Stored formulas use full column names, so this matches exactly what replay needs."""
    out: List[str] = []
    known = [str(c) for c in columns]
    for e in evals:
        if "=" not in e:
            continue
        new_col, formula = (x.strip() for x in e.split("=", 1))
        _, token_to_col, _ = compile_formula(formula, known)
        for col in token_to_col.values():
            if col not in out and col in columns:
                out.append(col)
        known.append(new_col)
    return out


def model_raw_inputs(meta: Dict[str, Any]) -> List[str]:
    """Columns the user must provide to a model: its own non-derived features plus the source columns of its
    derived features."""
    return list(meta["input_columns_required"])


def apply_eval_expressions(
    df: pd.DataFrame,
    eval_expressions: List[str],
    target_cols: Optional[List[str]] = None,
    strict: bool = True
) -> Tuple[pd.DataFrame, List[str], List[str]]:
    """
    Safely evaluates derived feature formulas 'new = expr' in order (later formulas may
    reference earlier ones).
    - Strictly blocks any reference to target variables (Y) to prevent data leakage.
    - Infinite results become NaN for downstream imputation.
    - strict=False (inference replay): failures are logged as [WARN] and skipped.
    Returns (df, canonical_expressions, created_columns).
    """
    if not eval_expressions:
        return df, [], []

    df = df.copy()
    target_set = set(target_cols or [])
    applied: List[str] = []
    created: List[str] = []

    for expr in eval_expressions:
        try:
            if "=" not in expr:
                raise InputError(f"수식 형식 오류: '=' 기호가 누락되었습니다: '{expr}'")
            new_col, formula = expr.split("=", 1)
            new_col, formula = new_col.strip(), formula.strip()
            if not new_col:
                raise InputError(f"생성할 파생변수명이 지정되지 않았습니다: '{expr}'")

            columns = [str(c) for c in df.columns]
            transformed, token_to_col, canonical = compile_formula(formula, columns)

            leaked = [c for c in token_to_col.values() if c in target_set]
            if leaked:
                raise InputError(
                    f"타겟 누수: 파생변수 '{new_col}' 수식에 타겟 {leaked}이(가) 들어 있습니다. "
                    f"타겟은 입력이나 파생변수 수식에 쓸 수 없습니다."
                )

            try:
                tree = ast.parse(transformed, mode="eval")
            except SyntaxError as e:
                raise InputError(f"--eval 수식 문법 오류 ('{expr}'): {e.msg}")

            eval_env = dict(SAFE_MATH)
            for token, col in token_to_col.items():
                eval_env[token] = coerce_numeric(df[col]).astype(float)
            _validate_eval_ast(tree, set(eval_env.keys()), columns, target_set)

            with np.errstate(all="ignore"):
                res_val = eval(compile(tree, "<eval>", "eval"), {"__builtins__": {}}, eval_env)
            if not isinstance(res_val, (pd.Series, np.ndarray, int, float)):
                raise InputError("수식 결과가 숫자형 배열이 아닙니다.")
            ser = pd.Series(res_val, index=df.index, dtype=float).replace([np.inf, -np.inf], np.nan)
            df[new_col] = ser
            applied.append(f"{new_col} = {canonical}")
            created.append(new_col)
        except Exception as ex:
            if strict:
                if isinstance(ex, InputError):
                    raise
                raise InputError(f"--eval 파생 수식 연산 실패 ('{expr}'): {ex}")
            sys.stderr.write(f"[WARN] 파생변수 연산 생략/실패 ('{expr}'): {ex}\n")

    return df, applied, created


# ==============================================================================
# C6. Windows-safe File Operations
# ==============================================================================

def retry_fs(fn: Callable[[], Any], attempts: int = 6, delay: float = 0.5) -> Any:
    """Retries a filesystem call that may transiently fail on Windows (antivirus/indexer/Explorer handles)."""
    for i in range(attempts):
        try:
            return fn()
        except PermissionError:
            if i == attempts - 1:
                record_recovery_attempt(f"파일 작업 {attempts}회 재시도 실패")
                raise
            time.sleep(delay * (i + 1))


def _make_writable_and_retry(func, path, _exc_info):
    try:
        os.chmod(path, 0o700)
        func(path)
    except Exception:
        pass


def remove_tree(path: str) -> None:
    """Removes a directory tree, clearing Windows read-only flags; retries on transient locks."""
    if not os.path.exists(path):
        return

    def _rm():
        if sys.version_info >= (3, 12):
            shutil.rmtree(path, onexc=_make_writable_and_retry)
        else:
            shutil.rmtree(path, onerror=_make_writable_and_retry)
        if os.path.exists(path):
            raise PermissionError(f"폴더 삭제 실패 (다른 프로그램이 사용 중): {path}")

    retry_fs(_rm)


def atomic_write_text(path: str, text: str) -> None:
    """Writes to a sibling temp file then os.replace()s it, so readers never see a half-written file."""
    tmp = f"{path}.tmp{os.getpid()}"
    with open(tmp, "w", encoding="utf-8", newline="\n") as f:
        f.write(text)
    try:
        retry_fs(lambda: os.replace(tmp, path))
    finally:
        if os.path.exists(tmp):
            try:
                os.remove(tmp)
            except OSError:
                pass


_STAMP_SUFFIX = re.compile(r"_\d{8}_\d{6}(_\d+)?$")


def _timestamped(path: str) -> str:
    """'a.csv' -> 'a_YYYYMMDD_HHMMSS.csv'. An existing time suffix is replaced (never doubled); never returns an existing path."""
    base, ext = os.path.splitext(path)
    base = _STAMP_SUFFIX.sub("", base)
    alt = f"{base}_{datetime.now().strftime('%Y%m%d_%H%M%S')}{ext}"
    n = 2
    while os.path.exists(alt):
        alt = f"{base}_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{n}{ext}"
        n += 1
    return alt


def avoid_overwrite(path: str, overwrite: bool = False, what: str = "결과 파일") -> str:
    """Never replaces an existing result file unless --overwrite was given: returns a timestamped sibling instead."""
    if overwrite or not os.path.exists(path):
        return path
    alt = _timestamped(path)
    note_recovered(f"{what} {os.path.basename(path)} 이미 있음 -> {os.path.basename(alt)}로 저장 (덮어쓰려면 --overwrite)")
    return alt


def save_with_lock_fallback(path: str, writer: Callable[[str], None], what: str = "파일") -> str:
    """
    Calls writer(path). If the file is locked (opened in Excel/Word), saves to a
    timestamped sibling instead and returns the path actually written.
    """
    try:
        writer(path)
        return path
    except PermissionError:
        alt = _timestamped(path)
        writer(alt)
        note_recovered(f"{what} {os.path.basename(path)} 잠김(Excel/Word 등에서 열림) -> {os.path.basename(alt)}로 저장")
        return alt


def write_csv_excel_compatible(df: pd.DataFrame, path: str) -> str:
    """
    Writes a CSV that people and LLM agents read: UTF-8 with BOM (utf-8-sig). Korean Windows Excel opens it by
    double-click, and an agent's file reader does not mistake it for another encoding (a CP949 file was read as
    garbled text by an agent, which then made up the content). Readers in this tool detect the encoding anyway.
    Returns the path actually written (may differ if the target was locked).
    """
    def _write(p: str) -> None:
        df.to_csv(p, index=False, encoding="utf-8-sig")

    return save_with_lock_fallback(path, _write, what="CSV 파일")


# ==============================================================================
# C8. Table Interpolation (ml_map.py, and the measured-data comparison of ml_predict.py)
# ==============================================================================
#
# grid      : every combination of the key values exists (one key: evenly spaced) -> classic multilinear
#             interpolation, end segments extended linearly outside the table.
# scattered : 1 key -> the two neighbouring rows; 2-3 keys -> the enclosing simplex (Delaunay, keys scaled to 0-1);
#             outside the measured area a least-squares trend line (never the two end points).
# Rows need a '__row' column (Excel row number shown as evidence).

EXCLUDE_COL = "제외"
INCLUDE_MARK = "포함"
NEAREST_N = 5
MAX_KEYS = 3              # interpolation up to 3 numeric keys; more keys -> nearest rows only
NUMERIC_SHARE = 0.9       # a column is numeric when this share of its non-empty cells parse as numbers


def _is_numeric_col(s: pd.Series) -> bool:
    """A number column of ml_map (a key to interpolate, a value column). A code column ('007', '7', '12':
    is_code_column) is not: its rows are used only for the exact same code."""
    if pd.api.types.is_datetime64_any_dtype(s) or is_code_column(s):
        return False
    filled = s.notna() & (s.astype(str).str.strip() != "")
    if filled.sum() == 0:
        return False
    return coerce_numeric(s)[filled].notna().mean() >= NUMERIC_SHARE


def _cell(v: Any) -> Any:
    """JSON / Excel friendly cell value."""
    if isinstance(v, (pd.Timestamp, datetime)):
        return v.isoformat(sep=" ", timespec="minutes")
    if isinstance(v, (np.floating, float)):
        return None if not np.isfinite(v) else round(float(v), 6)
    if isinstance(v, np.integer):
        return int(v)
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return None
    return v


def _fmt(v: float) -> str:
    return f"{v:.6g}"


class Table:
    """Rows filtered by categorical keys, numeric keys (n x k) and value columns (n x m)."""

    def __init__(self, df: pd.DataFrame, keys: List[str], values: List[str], warnings: List[str]):
        self.keys, self.values = keys, values
        K = np.column_stack([coerce_numeric(df[c]).to_numpy(dtype=float) for c in keys])
        V = np.column_stack([coerce_numeric(df[c]).to_numpy(dtype=float) for c in values])
        ok = np.isfinite(K).all(axis=1) & np.isfinite(V).all(axis=1)
        if (~ok).sum():
            warnings.append(f"기준 열이나 구할 열에 빈칸·글자가 있는 {int((~ok).sum())}행은 쓰지 않았습니다")
        self.df = df[ok].reset_index(drop=True)
        self.K, self.V = K[ok], V[ok]
        self.rows = self.df["__row"].to_numpy()
        n, k = self.K.shape
        if n < max(2, k + 1):
            raise InputError(f"계산에 쓸 수 있는 행이 {n}개뿐입니다 (기준 열 {k}개에는 {max(2, k + 1)}행 이상 필요).")
        self.lo, self.hi = self.K.min(axis=0), self.K.max(axis=0)
        span = self.hi - self.lo
        self.span = np.where(span > 0, span, 1.0)
        self._merge(warnings)
        self.mode = "grid" if self._is_grid() else "scattered"
        self._tri = None
        self._trend = None

    def _merge(self, warnings: List[str]) -> None:
        """One point per key combination: identical duplicates are merged, differing ones averaged (warned)."""
        frame = pd.DataFrame(self.K, columns=self.keys)
        groups = frame.groupby(self.keys, sort=True).indices
        pk, pv, pr, conflicts = [], [], [], 0
        for _, idx in groups.items():
            idx = np.asarray(idx)
            vals = self.V[idx]
            if len(idx) > 1 and not np.allclose(vals, vals[0], equal_nan=True):
                conflicts += 1
            pk.append(self.K[idx[0]])
            pv.append(vals.mean(axis=0))
            pr.append([int(r) for r in self.rows[idx]])
        self.conflicts = conflicts
        if conflicts:
            warnings.append(f"같은 기준 값에 서로 다른 값이 있는 조합 {conflicts}개는 평균으로 계산했습니다")
        self.PK, self.PV, self.PR = np.array(pk), np.array(pv), pr

    def _is_grid(self) -> bool:
        if self.conflicts:
            return False
        self.axes = [np.unique(self.PK[:, j]) for j in range(self.PK.shape[1])]
        if any(len(a) < 2 for a in self.axes):
            return False
        if len(self.axes) == 1:
            # One key: every table is a "full grid", so only an evenly spaced key (0, 5, 10, ...) counts as a map table.
            # Uneven keys (process records) are interpolated the same way inside, but extrapolated by the trend line.
            d = np.diff(self.axes[0])
            return bool(np.allclose(d, d[0], rtol=1e-6, atol=1e-9 * max(1.0, abs(d[0]))))
        return int(np.prod([len(a) for a in self.axes])) == len(self.PK)

    # -- forward ---------------------------------------------------------------
    def out_of_range(self, x: np.ndarray) -> Dict[str, str]:
        def share(d: float, span: float) -> str:
            r = d / span
            return f"데이터 폭의 {r:.1f}배" if r >= 1 else f"데이터 폭의 {r:.0%}" if r >= 0.01 else "데이터 폭의 1% 미만"
        out = {}
        for j, key in enumerate(self.keys):
            lo, hi, span = self.lo[j], self.hi[j], self.span[j]
            if x[j] < lo:
                out[key] = f"범위 {_fmt(lo)}~{_fmt(hi)}보다 {_fmt(lo - x[j])} 낮음 ({share(lo - x[j], span)})"
            elif x[j] > hi:
                out[key] = f"범위 {_fmt(lo)}~{_fmt(hi)}보다 {_fmt(x[j] - hi)} 높음 ({share(x[j] - hi, span)})"
        return out

    def _basis(self, idx: List[int], w: List[float]) -> List[Dict[str, Any]]:
        out = []
        for i, wi in zip(idx, w):
            if abs(wi) > 1e-12:
                out.append({"row": self.PR[i][0] if len(self.PR[i]) == 1 else self.PR[i], "weight": round(float(wi), 6)})
        return out

    def _grid(self, x: np.ndarray) -> Tuple[np.ndarray, List[Dict[str, Any]]]:
        seg, frac = [], []
        for j, a in enumerate(self.axes):
            i = int(np.clip(np.searchsorted(a, x[j], side="right") - 1, 0, len(a) - 2))
            seg.append(i)
            frac.append((x[j] - a[i]) / (a[i + 1] - a[i]))   # <0 or >1 outside the table: linear extension
        index = {tuple(p): n for n, p in enumerate(self.PK)}
        val = np.zeros(self.PV.shape[1])
        idx, wts = [], []
        for corner in product((0, 1), repeat=len(self.axes)):
            w = 1.0
            key = []
            for j, c in enumerate(corner):
                w *= frac[j] if c else 1.0 - frac[j]
                key.append(self.axes[j][seg[j] + c])
            n = index[tuple(key)]
            val += w * self.PV[n]
            idx.append(n)
            wts.append(w)
        return val, self._basis(idx, wts)

    def _trend_value(self, x: np.ndarray) -> np.ndarray:
        if self._trend is None:
            A = np.column_stack([np.ones(len(self.PK)), (self.PK - self.lo) / self.span])
            self._trend = np.linalg.lstsq(A, self.PV, rcond=None)[0]
        return np.concatenate([[1.0], (x - self.lo) / self.span]) @ self._trend

    def _scattered(self, x: np.ndarray) -> Tuple[Optional[np.ndarray], List[Dict[str, Any]]]:
        if self.PK.shape[1] == 1:
            a = self.PK[:, 0]
            if not (a[0] <= x[0] <= a[-1]):
                return None, []
            i = int(np.clip(np.searchsorted(a, x[0], side="right") - 1, 0, len(a) - 2))
            t = (x[0] - a[i]) / (a[i + 1] - a[i])
            return (1 - t) * self.PV[i] + t * self.PV[i + 1], self._basis([i, i + 1], [1 - t, t])
        if self._tri is None:
            try:
                from scipy.spatial import Delaunay
            except ImportError:
                raise InputError("scipy 패키지가 없어 기준 2~3개 보간을 할 수 없습니다. `python -m pip install scipy`로 설치하십시오.")
            try:
                self._tri = Delaunay((self.PK - self.lo) / self.span)
            except Exception as e:  # degenerate layout (all points on a line, too few points)
                raise InputError(f"실측 점이 한 줄로 놓여 있거나 너무 적어 둘러싼 점을 찾을 수 없습니다: {e}")
        q = (x - self.lo) / self.span
        s = int(self._tri.find_simplex(q))
        if s < 0:
            return None, []
        T = self._tri.transform[s]
        b = T[:-1] @ (q - T[-1])
        w = np.append(b, 1 - b.sum())
        verts = self._tri.simplices[s]
        return w @ self.PV[verts], self._basis(list(verts), list(w))

    def at(self, x: np.ndarray) -> Dict[str, Any]:
        oor = self.out_of_range(x)
        if self.mode == "grid":
            val, basis = self._grid(x)
            method = "격자 보간" + (" (끝 구간 직선 연장)" if oor else "")
            extrap = bool(oor)
        else:
            val, basis = self._scattered(x)
            if val is None:
                val, basis = self._trend_value(x), []
                method, extrap = "전체 추세선 외삽 (실측이 둘러싼 범위 밖)", True
            else:
                method = "위아래 실측 2점 보간" if self.PK.shape[1] == 1 else f"둘러싼 실측 {len(basis)}점 보간"
                extrap = False
        return {"values": {c: round(float(v), 6) for c, v in zip(self.values, val)}, "method": method,
                "basis": basis, "extrapolated": extrap, "out_of_range": oor}


class MultiTable:
    """Several value columns, each interpolated on its own non-empty rows (ml_map.py without -y).
    Rows are dropped for an empty key, but a blank in one value column does not remove that row from the others:
    value columns with the same empty cells share one Table. A value column with too few values for a table is left
    out with a warning (error only when no column is left). Range (lo, hi, span) and rows (df, for nearest_rows) are
    those of the rows used by the kept tables (a row of a dropped column only is not data); the other attributes come
    from the table of the first kept value column."""

    def __init__(self, df: pd.DataFrame, keys: List[str], values: List[str], warnings: List[str]):
        K = np.column_stack([coerce_numeric(df[c]).to_numpy(dtype=float) for c in keys])
        key_ok = np.isfinite(K).all(axis=1)
        if (~key_ok).sum():
            warnings.append(f"기준 열에 빈칸·글자가 있는 {int((~key_ok).sum())}행은 쓰지 않았습니다")
        df = df[key_ok].reset_index(drop=True)
        groups: Dict[bytes, List[str]] = {}
        masks: Dict[bytes, np.ndarray] = {}
        for c in values:
            m = np.isfinite(coerce_numeric(df[c]).to_numpy(dtype=float))
            groups.setdefault(m.tobytes(), []).append(c)
            masks[m.tobytes()] = m
        self.tables: List[Tuple[List[str], "Table"]] = []
        used = np.zeros(len(df), dtype=bool)
        last_error: Optional[InputError] = None

        def _add(msgs: List[str]) -> None:   # the same Table warning from several column groups is said once
            for w in msgs:
                if w not in warnings:
                    warnings.append(w)

        for key, cols in groups.items():
            n_blank = int((~masks[key]).sum())
            msgs: List[str] = []
            if n_blank and len(groups) == 1:
                msgs.append(f"구할 열에 빈칸·글자가 있는 {n_blank}행은 쓰지 않았습니다")
            elif n_blank:
                msgs.append(f"{', '.join(cols)}: 빈칸·글자가 있는 {n_blank}행은 이 열 계산에서만 뺐습니다")
            try:
                tab = Table(df[masks[key]].reset_index(drop=True), keys, cols, msgs)
            except InputError as e:
                if len(groups) == 1:
                    raise InputError(f"구할 열 {', '.join(cols)}: {e}")
                last_error = e
                _add([f"{', '.join(cols)}: 값이 있는 행이 적어 계산하지 않았습니다 ({e})"])
                continue
            _add(msgs)
            self.tables.append((cols, tab))
            used |= masks[key]
        if not self.tables:
            raise last_error or InputError("계산할 수 있는 구할 열이 없습니다.")
        self.values = [c for c in values if any(c in cols for cols, _ in self.tables)]
        self.primary = self.tables[0][1]
        self.mode = "grid" if all(t.mode == "grid" for _, t in self.tables) else "scattered"
        if len(groups) > 1:
            # range and rows of every row a kept table uses (not only the rows of the first value column)
            self.df, self.K = df[used].reset_index(drop=True), K[key_ok][used]
            self.lo, self.hi = self.K.min(axis=0), self.K.max(axis=0)
            span = self.hi - self.lo
            self.span = np.where(span > 0, span, 1.0)
            self.rows = self.df["__row"].to_numpy()

    def __getattr__(self, name: str) -> Any:
        primary = self.__dict__.get("primary")
        if primary is None:
            raise AttributeError(name)
        return getattr(primary, name)

    def at(self, x: np.ndarray) -> Dict[str, Any]:
        if len(self.tables) == 1:
            return self.primary.at(x)
        parts = [(cols, t.at(x)) for cols, t in self.tables]
        vals = {c: v for _, r in parts for c, v in r["values"].items()}
        methods = list(dict.fromkeys(r["method"] for _, r in parts))
        # out-of-range text per key: one text when every column group says the same, else per column group
        oor: Dict[str, str] = {}
        for key in self.keys:
            texts = [(cols, r["out_of_range"][key]) for cols, r in parts if key in r["out_of_range"]]
            if not texts:
                continue
            if len({t for _, t in texts}) == 1 and len(texts) == len(parts):
                oor[key] = texts[0][1]
            else:
                oor[key] = "; ".join(f"{', '.join(cols)}: {t}" for cols, t in texts)
        by_col = {c: r["basis"] for cols, r in parts for c in cols}
        return {"values": {c: vals[c] for c in self.values},
                "method": methods[0] if len(methods) == 1 else "; ".join(f"{', '.join(c)}: {r['method']}" for c, r in parts),
                "basis": parts[0][1]["basis"], "basis_by_column": {c: by_col[c] for c in self.values},
                "extrapolated": any(r["extrapolated"] for _, r in parts), "out_of_range": oor}


def nearest_rows(df: pd.DataFrame, cols: List[str], target: np.ndarray, span: np.ndarray,
                 show: List[str], n: int = NEAREST_N) -> List[Dict[str, Any]]:
    X = np.column_stack([coerce_numeric(df[c]).to_numpy(dtype=float) for c in cols])
    num = {c: coerce_numeric(df[c]) for c in show if _is_numeric_col(df[c])}    # numbers, not text from CSV
    d = np.sqrt(np.nansum(((X - target) / span) ** 2, axis=1))
    d[~np.isfinite(X).all(axis=1)] = np.inf
    out = []
    for i in np.argsort(d, kind="stable")[:n]:
        if not np.isfinite(d[i]):
            break
        rec = {"row": int(df["__row"].iloc[i]), "distance": round(float(d[i]), 4)}
        rec.update({c: _cell(num[c].iloc[i] if c in num else df[c].iloc[i]) for c in show})
        out.append(rec)
    return out


def _crossings(xs: np.ndarray, ys: np.ndarray, y0: float) -> List[float]:
    """x where the piecewise linear curve (xs, ys) crosses y0 (segments with NaN are skipped)."""
    out: List[float] = []
    for i in range(len(xs) - 1):
        a, b = ys[i], ys[i + 1]
        if not (np.isfinite(a) and np.isfinite(b)):
            continue
        if a == y0:
            out.append(float(xs[i]))
        elif (a - y0) * (b - y0) < 0:
            out.append(float(xs[i] + (y0 - a) * (xs[i + 1] - xs[i]) / (b - a)))
    if len(ys) and np.isfinite(ys[-1]) and ys[-1] == y0:
        out.append(float(xs[-1]))
    return sorted(set(round(v, 10) for v in out))


# ##############################################################################
# PART P. INFERENCE ENGINE
# ##############################################################################

# ==============================================================================
# 1. Derived Feature Replay (shared AST sandbox, identical to training)
# ==============================================================================

def apply_inference_evals(df: pd.DataFrame, eval_expressions: List[str]) -> pd.DataFrame:
    """Applies derived feature formulas recorded in model metadata prior to inference."""
    df_out, _, _ = apply_eval_expressions(df, eval_expressions, strict=False)
    return df_out


def check_inline_aliases(df: pd.DataFrame, required: List[str]) -> None:
    """-i: '온도' next to '온도(°C)' when the model has '온도(°C)' but not '온도' (align_input_columns would match
    '온도' to '온도(°C)'): one value would be silently dropped -> STOP_RETRY. A model with both '온도' and '온도(°C)'
    takes both. A name with the same unit written otherwise ('온도(C)', '온도(℃)', '온도( °C)': _same_unit_name, the
    other align_input_columns rule) counts as the model name too ('온도' + '온도(C)')."""
    req = set(required)
    given = [str(c) for c in df.columns]
    target = []
    for c in given:
        if c in req:
            hit = []
        elif strip_unit(c) != c.strip():
            hit = [f for f in required if _same_unit_name(c, f)]
        else:
            hit = [f for f in required if _alias_key(f) == _alias_key(c)]
        target.append(hit[0] if len(hit) == 1 else c)
    values = {str(k): v for k, v in df.iloc[0].to_dict().items()} if len(df) == 1 else None
    check_same_column(given, target, "-i", values)


def check_inline_unknown(df: pd.DataFrame, known: List[str], inputs: List[str]) -> None:
    """-i (k=v, JSON text or .json file): a key no model uses ('압력(MPa)=0.5' next to '압력(bar)=3', '습도=50') would
    be dropped without a word -> STOP_RETRY. known: every name a model takes (inputs, features, derived names) after
    the unit-less renames of align_input_columns; inputs: the model inputs shown. Matched like a missing input: a
    model name of the same name without unit (given or not: '압력(MPa)' for '압력(bar)'), else a close spelling among
    the model names not given (suggest_columns; '습도' is no typo of the given '속도'). A different unit -> unit_hint,
    else [후보] (even one: a close spelling may be another quantity), none -> the model inputs. File input (-f) keeps its extra
    columns (ID, memo)."""
    given = [str(c) for c in df.columns]
    known = list(dict.fromkeys(known))
    unknown = [c for c in given if c not in set(known)]
    if not unknown:
        return
    present = set(given)
    hints = []
    for c in unknown:
        s = [f for f in known if _alias_key(f) == _alias_key(c)] or \
            suggest_columns(c, [f for f in known if f not in present])
        if len(s) == 1 and other_unit(c, s[0]):
            hints.append(unit_hint(c, s[0], "-i의 입력") + (f". '{s[0]}' 값도 -i에 있습니다: 어느 값을 쓸지 사용자에게 물으십시오"
                                                          if s[0] in present else ""))
        elif s:
            hints.append(f"-i의 '{c}'에 가까운 모델 입력: [후보] {', '.join(unit_marked(s, c))}. 어느 것인지 사용자에게 물으십시오")
        else:
            hints.append(f"-i의 '{c}'과(와) 비슷한 모델 입력이 없습니다")
    raise InputError(f"-i의 {', '.join(repr(c) for c in unknown)}은(는) 모델 입력이 아니라 예측에 쓰이지 않습니다. "
                     f"{' / '.join(hints)}. 모델 입력이 아닌 값을 빼고 예측할지 사용자에게 물으십시오.\n[모델 입력 목록]: {name_list(dict.fromkeys(inputs))}")


def align_input_columns(df: pd.DataFrame, required: List[str], derived: List[str]) -> Dict[str, str]:
    """
    Renames input columns written without unit suffix to the trained full name
    ('반응온도' -> '반응온도(°C)') or with the same unit spaced otherwise ('반응온도( °C)'), in place. Only unique,
    unambiguous matches are applied.
    """
    applied: Dict[str, str] = {}
    present = set(str(c) for c in df.columns)
    input_alias = build_alias_map([str(c) for c in df.columns])
    for feat in required:
        if feat in present or feat in derived:
            continue
        key = re.sub(r"\s+", "", strip_unit(feat)).lower()
        src = input_alias.get(key)
        if src is None or src in required or strip_unit(src) != src.strip():
            # the same unit spaced otherwise ('반응압력( bar)' for '반응압력(bar)'): the same name (resolve_column_names)
            same = [c for c in present if c not in required and _same_unit_name(c, feat)]
            src = same[0] if len(same) == 1 else None
        if src is not None:
            df.rename(columns={src: feat}, inplace=True)
            present.discard(src)
            present.add(feat)
            applied[src] = feat
            sys.stderr.write(f"[INFO] 입력 컬럼명 자동 매칭: '{src}' -> '{feat}'\n")
    return applied


# ==============================================================================
# 4. Target Model Resolution (Directory Mode vs Specific File Mode)
# ==============================================================================

# ml_train writes <data>_ml_data/best_version.json after every ledger change: per task the version a report or a
# prediction should use (ml_train task_best_version). `-m <data file>` (or the _ml_data folder) reads it, so the agent
# never chooses a version folder itself (2.9.71, AGENTS.md "사용법 단순성" 5).
BEST_POINTER = "best_version.json"
DATA_EXTS = (".xlsx", ".xls", ".xlsm", ".csv")
MODEL_CHOSEN: List[str] = []


def model_dir_from_data(path: str, targets: Optional[List[str]] = None, option: str = "-m") -> str:
    """-m value -> model path: a version folder or .joblib as given; a data file or its _ml_data folder -> the best
    version folder of best_version.json (one task, or the task holding `targets`). Several tasks with no way to tell
    -> STOP_RETRY [후보] full commands; no pointer -> STOP_RETRY [후보] the version folders as full commands."""
    ap = os.path.abspath(str(path).strip().strip('"'))
    if os.path.isfile(ap) and os.path.splitext(ap)[1].lower() in DATA_EXTS:
        root = os.path.join(os.path.dirname(ap), os.path.splitext(os.path.basename(ap))[0] + "_ml_data")
    elif os.path.isdir(ap) and os.path.basename(ap.rstrip("\\/")).endswith("_ml_data"):
        root = ap
    else:
        return path
    if not os.path.isdir(root):
        raise InputError(f"'{os.path.basename(ap)}'의 학습 결과 폴더가 없습니다 ({root}). 먼저 ml_train.py로 학습하십시오.")
    try:
        tasks = (json.loads(read_text_file(os.path.join(root, BEST_POINTER))) or {}).get("tasks") or {}
    except (OSError, ValueError):
        tasks = {}
    if not tasks:
        vers = sorted((d for d in os.listdir(root) if re.match(r"^v\d+$", d) and os.path.isdir(os.path.join(root, d))),
                      key=lambda d: int(d[1:]))
        if not vers:
            raise InputError(f"'{os.path.basename(root)}'에 학습된 버전 폴더가 없습니다. 먼저 ml_train.py로 학습하십시오.")
        cands = [command_with_value(str(path), os.path.join(root, v)) or os.path.join(root, v) for v in vers[-5:]]
        raise InputError("최고 버전 기록(best_version.json)이 없어 쓸 버전을 정하지 못했습니다. 쓸 버전을 사용자에게 물으십시오. [후보] "
                         + " / ".join(f"{i}) {c}" for i, c in enumerate(cands, 1)))
    keys = {strip_unit(str(t)).strip() for t in (targets or []) if str(t).strip()}
    picks = list(tasks)
    if len(picks) > 1 and keys:
        hold = [t for t in picks if keys & {strip_unit(str(x)).strip() for x in tasks[t].get("targets", [])}]
        picks = hold or picks
    if len(picks) > 1:
        cands = [command_with_value(str(path), os.path.join(root, tasks[t]["version"])) or os.path.join(root, tasks[t]["version"])
                 for t in picks]
        raise InputError("회귀와 분류 모델이 함께 있어 어느 버전으로 예측할지 정하지 못했습니다. 사용자에게 물으십시오. [후보] "
                         + " / ".join(f"{i}) {c} ({t})" for i, (c, t) in enumerate(zip(cands, picks), 1)))
    info = tasks[picks[0]]
    vdir = os.path.join(root, info["version"])
    msg = f"최고 버전 {info['version']} ({picks[0]}, best_version.json)으로 예측"
    if not info.get("complete", True):
        msg += f". 이 버전에 없는 타겟 최고 모델: {', '.join(info.get('missing') or [])} (ml_report.py -f로 한 버전에 묶을 수 있음)"
    sys.stderr.write(f"[INFO] {option} {os.path.basename(ap)} -> {vdir}\n")
    MODEL_CHOSEN.append(msg)
    return vdir


def resolve_model_arg(args: argparse.Namespace) -> argparse.Namespace:
    """args.model: a data file -> its best version folder (model_dir_from_data), before any work."""
    names = []
    for v in (getattr(args, "targets", None), getattr(args, "maximize", None), getattr(args, "minimize", None)):
        if v:
            names += [x for x in str(v).split(",") if x.strip()]
    if getattr(args, "find", None):
        names.append(str(args.find).split("=", 1)[0])
    args.model = model_dir_from_data(args.model, names)
    return args


def resolve_target_models(
    target_path: str,
    filter_targets: Optional[List[str]] = None,
    kind: str = "타겟(-y)",
    values: bool = False
) -> List[Dict[str, Any]]:
    """
    Resolves model and metadata pairs from a directory (v1/, v2/) or specific file.
    Returns a list of dicts:
    [{"target": "생산효율", "joblib_path": "...", "meta_path": "...", "meta": {...}, "pipeline": ...}]
    filter_targets are resolved like column names (resolve_column_names: exact name first, then a unique name
    without its unit; '수율' for '수율(1차)' and '수율(최종)' -> [후보], an exact '수율(1차)' takes that one only).
    kind / values: passed to resolve_column_names (values: a goal value, a different unit needs converting).
    """
    abs_path = os.path.abspath(target_path)
    if not os.path.exists(abs_path):
        raise FileNotFoundError(f"지정한 모델 경로를 찾을 수 없습니다: {abs_path}")

    model_items = []

    if os.path.isdir(abs_path):
        # Directory Mode: discover all model_*.joblib files
        joblib_files = glob.glob(os.path.join(abs_path, "model_*.joblib"))
        if not joblib_files:
            raise FileNotFoundError(f"해당 폴더 내에 'model_*.joblib' 모델 파일이 존재하지 않습니다: {abs_path}")

        for j_path in joblib_files:
            fname = os.path.basename(j_path)
            # Match model_{target}.joblib
            match = re.match(r"^model_(.+)\.joblib$", fname)
            if not match:
                continue
            target_name = match.group(1)

            m_path = os.path.join(abs_path, f"model_{target_name}.json")
            meta_dict = {}
            if os.path.exists(m_path):
                try:
                    with open(m_path, "r", encoding="utf-8") as f:
                        meta_dict = json.load(f)
                except Exception as e:
                    sys.stderr.write(f"[WARN] 메타데이터 로드 실패 ({m_path}): {e}\n")

            display_target = meta_dict.get("target", target_name)
            model_items.append({
                "target": display_target,
                "target_safe": target_name,
                "joblib_path": j_path,
                "meta_path": m_path if os.path.exists(m_path) else None,
                "meta": meta_dict
            })

    else:
        # Specific File Mode
        j_path = abs_path
        if not j_path.endswith(".joblib"):
            raise InputError(f"모델 파일은 '.joblib' 확장자여야 합니다: {j_path}")

        fname = os.path.basename(j_path)
        match = re.match(r"^model_(.+)\.joblib$", fname)
        target_name = match.group(1) if match else os.path.splitext(fname)[0]

        # Resolve Meta Path
        m_path = os.path.splitext(j_path)[0] + ".json"

        meta_dict = {}
        if m_path and os.path.exists(m_path):
            try:
                with open(m_path, "r", encoding="utf-8") as f:
                    meta_dict = json.load(f)
            except Exception as e:
                sys.stderr.write(f"[WARN] 메타데이터 로드 실패 ({m_path}): {e}\n")

        model_items.append({
            "target": meta_dict.get("target", target_name),
            "joblib_path": j_path,
            "meta_path": m_path,
            "meta": meta_dict
        })

    if filter_targets and model_items:
        names = sorted(str(it["target"]) for it in model_items)
        safe = {str(it.get("target_safe")): str(it["target"]) for it in model_items if it.get("target_safe")}
        # a target given by its file-safe name; the same unit spaced otherwise ('생산량( t )') is matched by
        # resolve_column_names like any column name
        given = [safe[f] if f not in names and f in safe else f for f in filter_targets]
        wanted = set(resolve_column_names(given, names, kind, values=values)[0])
        model_items = [it for it in model_items if str(it["target"]) in wanted]
    if not model_items:
        raise InputError(f"조건에 일치하는 대상 모델을 찾을 수 없습니다 (filter={filter_targets})")
    old = [it["target"] for it in model_items
           if not it["meta"].get("feature_schema") or not it["meta"].get("input_columns_required")]
    if old:   # beta: no fallback paths for models made by older versions
        raise InputError(f"타겟 {old}의 모델은 ml-kit 2.9.1보다 먼저 학습했거나 모델 정보 파일(model_<타겟>.json)이 없습니다. "
                         "ml-kit 2.9.1 이상으로 같은 데이터를 다시 학습한 뒤 사용하십시오.")

    # Load Pipelines (warn on scikit-learn version mismatch between training and serving PCs)
    try:
        import sklearn
        sk_now = sklearn.__version__
    except Exception:
        sk_now = None
    for item in model_items:
        sk_train = item["meta"].get("sklearn_version")
        if sk_train and sk_now and sk_train != sk_now:
            sys.stderr.write(
                f"[WARN] scikit-learn 버전 불일치: 학습={sk_train}, 현재={sk_now}. "
                f"폐쇄망 PC 간 동일 버전 wheel 설치를 권장합니다.\n"
            )
        try:
            item["pipeline"] = joblib.load(item["joblib_path"])
        except Exception as e:
            raise InputError(
                f"모델 파일을 읽지 못했습니다 ({item['joblib_path']}): {e}. "
                f"학습 PC와 scikit-learn 버전(학습={sk_train or '미기록'}, 현재={sk_now})이 다르면 동일 버전으로 맞추십시오."
            )

    return model_items


# ==============================================================================
# 5. Core Prediction Engine
# ==============================================================================

def _no_repeated_keys(pairs: List[Tuple[str, Any]]) -> Dict[str, Any]:
    """json object_pairs_hook for -i: a name given twice in one object ({"온도":180,"온도":190}) would silently keep
    the last value (the same rule as 'k=v')."""
    out: Dict[str, Any] = {}
    for k, v in pairs:
        if k in out:
            vals = [f"{k}={json.dumps(x, ensure_ascii=False)}" for kk, x in pairs if kk == k]
            raise InputError(f"입력 JSON의 한 객체에 '{k}'가 두 번 있습니다. 한 객체는 한 행입니다. 사용자에게 어느 값인지 "
                             f"물으십시오: [후보] {', '.join(vals)}")
        out[k] = v
    return out


# 'k1=v1; k2=v2' or 'k1=v1 k2=v2' ('k=v' pairs joined by ';' or spaces instead of ','): read as ONE pair (k1 = the
# text 'v1; k2=v2'), so k2 looked missing and the tool asked to fill it with the median (2.9.69). A ';' or a space
# before '<name>=' inside a value is that mistake: STOP_RETRY [추천] the command with ','.
_KV_SLIP = re.compile(r"(?:;|\s)\s*[^\s,;=]+\s*=")
_KV_SPLIT = re.compile(r"\s*;\s*|\s+(?=[^\s,;=]+\s*=)")


def command_with_value(old: str, new: str) -> Optional[str]:
    """The running command with the option whose value is `old` given `new` instead (a full command for [추천]),
    None when the running program is not an ml-kit command or no option has that value."""
    argv = CURRENT_ARGV if CURRENT_ARGV is not None else list(sys.argv)
    for j in range(2, len(argv)):
        if argv[j] == old and str(argv[j - 1]).startswith("-") and argv[j - 1] != old:
            cmd = command_text(drop={argv[j - 1]: ""})
            return f"{cmd} {argv[j - 1]} {quote_arg(new)}" if cmd else None
    return None


def _kv_separator_slip(input_str: str, row: Dict[str, str]) -> None:
    """'k1=v1; k2=v2' / 'k1=v1 k2=v2' -> STOP_RETRY ([추천] the pairs joined by ',' when the split has one meaning)."""
    slips = {k: v for k, v in row.items() if _KV_SLIP.search(v)}
    if not slips:
        return
    fixed: Dict[str, str] = {}
    clean = True
    for k, v in row.items():
        if k not in slips:
            fixed[k] = v
            continue
        pieces = [p for p in _KV_SPLIT.split(v) if p.strip()]
        first, rest = pieces[0], pieces[1:]
        if ";" not in v and re.search(r"\s", first.strip()):
            clean = False              # '160 압력 x=3': where the next name starts is not clear
        fixed[k] = first.strip()
        for p in rest:
            if "=" not in p:
                clean = False
                continue
            k2, v2 = p.split("=", 1)
            if k2.strip() in fixed or k2.strip() in row:
                clean = False
            fixed[k2.strip()] = v2.strip()
    text = ",".join(f"{k}={v}" for k, v in fixed.items())
    head = (f"입력 '{input_str}'에서 '이름=값' 사이를 ','가 아닌 ';'이나 띄어쓰기로 나눴습니다. 그러면 앞 이름의 값이 뒤까지 "
            f"전부로 읽혀 뒤 입력이 빠집니다. 'k1=v1,k2=v2'처럼 ','로 나누십시오.")
    if clean:
        raise InputError(f"{head} [추천] {command_with_value(input_str, text) or text}")
    raise InputError(head + " 이름에 띄어쓰기가 있어 어디서 나눌지 정하지 못했습니다. 사용자에게 입력을 확인하십시오.")


def parse_inline_input(input_str: str) -> pd.DataFrame:
    """
    Parses inline input into a DataFrame:
    - path to a .json file (recommended on Windows: no shell quote escaping)
    - JSON object/array string
    - 'k1=v1,k2=v2' pairs
    """
    clean_str = input_str.strip()
    if clean_str.lower().endswith(".json") and os.path.isfile(clean_str):
        try:
            obj = json.loads(read_text_file(clean_str), object_pairs_hook=_no_repeated_keys)
        except json.JSONDecodeError as e:
            raise InputError(f"JSON 파일 파싱 실패 ({clean_str}): {e}")
        if isinstance(obj, dict):
            return pd.DataFrame([obj])
        if isinstance(obj, list) and obj:
            return pd.DataFrame(obj)
        raise InputError(f"입력 JSON 파일은 객체 또는 객체 배열이어야 합니다: {clean_str}")
    if clean_str.startswith("{") or clean_str.startswith("["):
        try:
            obj = json.loads(clean_str, object_pairs_hook=_no_repeated_keys)
        except InputError:
            raise
        except Exception as e:
            raise InputError(f"인라인 JSON 파싱 실패: {e}. Windows에서는 'k1=v1,k2=v2' 형식이나 .json 파일 경로를 권장합니다.")
        if isinstance(obj, dict):
            return pd.DataFrame([obj])
        if isinstance(obj, list) and obj:
            return pd.DataFrame(obj)
        raise InputError("인라인 JSON은 객체 또는 객체 배열이어야 합니다.")

    # key=val, key2=val2 format. Values stay text like a cell read from a file: numeric inputs are converted
    # later (prepare_features / coerce_numeric), so a category code such as '007' is not turned into 7.
    # A piece without '=' belongs to the previous value ('라인=A,B' is the text 'A,B'). A NUMERIC input written
    # that way ('압력=500,000', '온도=180,190') is ambiguous: check_comma_numbers stops with [추천] instead.
    row_dict: Dict[str, str] = {}
    last_key: Optional[str] = None
    for pair in clean_str.split(","):
        if "=" in pair:
            k, v = pair.split("=", 1)
            last_key = k.strip()
            if last_key in row_dict:
                raise InputError(
                    f"입력 '{input_str}'에 '{last_key}='가 두 번 있습니다. 'k1=v1,k2=v2' 입력은 한 행입니다. 여러 행은 "
                    f"JSON 배열(-i \"[{{...}},{{...}}]\")이나 .json 파일 경로로 주십시오. 한 행이면 사용자에게 값을 확인하십시오.")
            row_dict[last_key] = v.strip()
        elif pair.strip():
            if last_key is None:
                raise InputError(f"입력 '{pair.strip()}' 앞에 '이름='이 없습니다. 'k1=v1,k2=v2' 형식으로 주십시오: '{input_str}'")
            row_dict[last_key] = f"{row_dict[last_key]},{pair.strip()}"
    if not row_dict:
        raise InputError(f"유효한 인라인 입력 형식이 아닙니다 (JSON 또는 'k1=v1,k2=v2'): '{input_str}'")
    _kv_separator_slip(input_str, row_dict)
    # '온도' next to '온도(°C)' (one column when the model / data has only '온도(°C)') is checked where the names are
    # matched to the model inputs or data columns (check_same_column), not here: a model can have both.
    df = pd.DataFrame([row_dict])
    df.attrs["kv_pairs"] = dict(row_dict)     # 'k=v' text form (check_comma_numbers rebuilds it)
    return df


_KV_THOUSANDS_RE = re.compile(r"[+-]?[1-9]\d{0,2}(,\d{3})+(\.\d+)?")      # 'k=v' input value (fullmatch)


def check_comma_numbers(pairs: Optional[Dict[str, Any]], numeric: Callable[[str], bool], option: str,
                        command_for: Callable[[Dict[str, str]], str]) -> None:
    """'k=v' input (ml_predict -i/--find, ml_map -i/--find): a numeric value with a comma was split at the comma
    and joined back, so '온도=180,190' could be 180190 or two values. Never accepted silently:
    thousands grouping (1,000 / 180,190) -> STOP_RETRY [추천] <full command with the value written without commas>,
    anything else (1,5) -> STOP_RETRY without a suggestion. Text values keep the joined text.
    pairs: the given names and values (None for JSON input); numeric(name) -> the name is a numeric input."""
    if not pairs:
        return
    bad = {k: str(v) for k, v in pairs.items() if isinstance(v, str) and "," in v and numeric(k)}
    if not bad:
        return
    what = ", ".join(f"'{k}={v}'" for k, v in bad.items())
    head = (f"{option}의 숫자 입력 {what}에 쉼표가 있습니다. 쉼표는 입력을 나누는 기호라 숫자 값 안에 쓰면 "
            f"값 하나인지 여러 개인지 알 수 없습니다.")
    if all(_KV_THOUSANDS_RE.fullmatch(v.strip()) for v in bad.values()):
        fixed = {k: (str(v).replace(",", "") if k in bad else str(v)) for k, v in pairs.items()}
        raise InputError(f"{head} 천 단위 구분 기호로 보고 쉼표를 뺀 명령을 추천합니다 (값을 여러 개 주려던 것이면 "
                         f"실행하지 말고 사용자에게 값을 물으십시오). [추천] {command_for(fixed)}")
    raise InputError(f"{head} 숫자 하나를 쉼표 없이 쓰십시오 (소수점은 '.'). 사용자에게 값을 확인하십시오.")


def kv_text(pairs: Dict[str, str]) -> str:
    return ",".join(f"{k}={v}" for k, v in pairs.items())


# ==============================================================================
# 4-B. Measured-data comparison (map check)
# ==============================================================================
#
# Every regression prediction is compared with the training rows themselves (vX/model_dataset.csv, which travels
# with the version folder): <= 3 numeric inputs -> the same table interpolation as ml_map.py; more inputs ->
# inverse-distance weighted mean of the 5 nearest measured rows (inputs scaled to 0-1). Categorical inputs keep
# only rows of the same category. Rows left out of training ('학습제외_<target>' = 1) are not used.
# The comparison never fails the prediction: problems become map_check_notes.

MAP_CHECK_MAX_ROWS = 20000
MAP_WARN_NEAR = 0.1       # IDW comparison only warns when the nearest measured row is this close (inputs scaled to 0-1)
SNAPSHOT_ROW_COL = "원본행"
SNAPSHOT_EXCL_PREFIX = "학습제외_"


def _load_snapshot(version_dir: str, cache: Dict[str, Any]) -> Optional[pd.DataFrame]:
    if version_dir not in cache:
        path = os.path.join(version_dir, "model_dataset.csv")
        cache[version_dir] = load_data_file(path) if os.path.isfile(path) else None
    return cache[version_dir]


def _idw_rows(sub: pd.DataFrame, keys: List[str], target: str) -> Dict[str, np.ndarray]:
    """Numeric arrays for _idw. Built once per category: converting the columns for every input row is what takes minutes."""
    K = np.column_stack([coerce_numeric(sub[c]).to_numpy(dtype=float) for c in keys])
    y = coerce_numeric(sub[target]).to_numpy(dtype=float)
    ok = np.isfinite(K).all(axis=1) & np.isfinite(y)
    K, y, rows = K[ok], y[ok], sub["__row"].to_numpy()[ok]
    lo, hi = K.min(axis=0), K.max(axis=0)
    return {"K": K, "y": y, "rows": rows, "lo": lo, "hi": hi, "span": np.where(hi - lo > 0, hi - lo, 1.0)}


def _idw(data: Dict[str, np.ndarray], x: np.ndarray) -> Dict[str, Any]:
    K, y, rows, lo, hi, span = (data[k] for k in ("K", "y", "rows", "lo", "hi", "span"))
    d = np.sqrt((((K - x) / span) ** 2).sum(axis=1))
    near = np.argsort(d, kind="stable")[:NEAREST_N]
    if d[near[0]] == 0:
        near = near[d[near] == 0]
        w = np.full(len(near), 1.0 / len(near))
    else:
        w = 1.0 / d[near]
        w = w / w.sum()
    return {"value": float(w @ y[near]), "method": f"가까운 실측 {len(near)}행 거리 가중평균",
            "basis": [{"row": int(rows[i]), "weight": round(float(wi), 6)} for i, wi in zip(near, w)],
            "extrapolated": bool(((x < lo) | (x > hi)).any()), "nearest_distance": float(d[near[0]])}


def compare_with_measured(meta: Dict[str, Any], target: str, version_dir: str, df_evaled: pd.DataFrame,
                          cache: Dict[str, Any]) -> Dict[str, Any]:
    """Measured-data comparison values for each input row (NaN where not computable) or a note why not."""
    schema = meta.get("feature_schema") or {}
    derived = {e.split("=", 1)[0].strip() for e in (meta.get("derived_evals") or []) if "=" in e}
    raw = model_raw_inputs(meta)
    cats = [c for c in schema.get("categorical", {}) if c in raw]
    keys = [c for c in raw if c not in cats and c not in schema.get("datetime", {}) and c not in derived]
    if not keys:
        return {"note": f"타겟 '{target}': 숫자 입력이 없어 실측 비교값을 계산하지 않았습니다"}
    if len(df_evaled) > MAP_CHECK_MAX_ROWS:
        return {"note": f"입력이 {len(df_evaled):,}행으로 많아 실측 비교값을 계산하지 않았습니다 (최대 {MAP_CHECK_MAX_ROWS:,}행)"}
    data = _load_snapshot(version_dir, cache)
    if data is None or any(c not in data.columns for c in keys + cats + [target, SNAPSHOT_ROW_COL]):
        return {"note": f"타겟 '{target}': 버전 폴더에 학습 데이터 사본(model_dataset.csv)이 없거나 열이 달라 실측 비교값을 계산하지 않았습니다 "
                        "(2.7.0 이전 버전이면 다시 학습하면 됩니다)"}
    data = data.copy()
    notes: List[str] = []
    data["__row"] = coerce_numeric(data[SNAPSHOT_ROW_COL]).fillna(-1).astype(int)
    excl = f"{SNAPSHOT_EXCL_PREFIX}{target}"
    if excl in data.columns:
        data = data[coerce_numeric(data[excl]).fillna(0) != 1]
    data = data[coerce_numeric(data[target]).notna()]
    for c in cats:
        data[c] = normalize_category(data[c])

    X = np.column_stack([coerce_numeric(df_evaled[c]).to_numpy(dtype=float) if c in df_evaled.columns
                         else np.full(len(df_evaled), np.nan) for c in keys])
    C = pd.DataFrame({c: normalize_category(df_evaled[c]) if c in df_evaled.columns else None for c in cats},
                     index=df_evaled.index) if cats else None
    out = {"values": np.full(len(df_evaled), np.nan), "method": [""] * len(df_evaled),
           "basis": [[] for _ in range(len(df_evaled))], "extrapolated": [False] * len(df_evaled),
           "warnable": [False] * len(df_evaled)}
    tables: Dict[Tuple, Any] = {}
    for i in range(len(df_evaled)):
        if not np.isfinite(X[i]).all():
            out["method"][i] = "계산 불가 (입력 빈칸)"
            continue
        cat_key = tuple(C.iloc[i].tolist()) if C is not None else ()
        if cat_key not in tables:
            sub = data
            for c, v in zip(cats, cat_key):
                if (sub[c] == v).any():
                    sub = sub[sub[c] == v]
                else:
                    notes.append(f"'{c}={v}'은(는) 학습 데이터에 없어 모든 행으로 비교했습니다")
            tab = None
            if len(keys) <= MAX_KEYS:
                try:
                    tab = Table(sub, keys, [target], [])
                except InputError:
                    tab = None
            tables[cat_key] = (sub, tab, None)
        sub, tab, idw_rows = tables[cat_key]
        try:
            if tab is not None:
                r = tab.at(X[i])
                out["values"][i] = r["values"][target]
                out["method"][i], out["basis"][i], out["extrapolated"][i] = r["method"], r["basis"], r["extrapolated"]
                out["warnable"][i] = not r["extrapolated"]
            else:
                if idw_rows is None:
                    idw_rows = _idw_rows(sub, keys, target)
                    tables[cat_key] = (sub, tab, idw_rows)
                r = _idw(idw_rows, X[i])
                out["values"][i] = r["value"]
                out["method"][i], out["basis"][i], out["extrapolated"][i] = r["method"], r["basis"], r["extrapolated"]
                # with many inputs the nearest rows can be far apart: only a close neighbourhood is worth a warning
                out["warnable"][i] = (not r["extrapolated"]) and r["nearest_distance"] <= MAP_WARN_NEAR
        except Exception as e:  # e.g. too few rows of that category
            out["method"][i] = f"계산 불가 ({str(e)[:60]})"
    out["keys"] = keys
    if notes:
        out["note"] = "; ".join(dict.fromkeys(notes))
    return out


LOW_CONFIDENCE_THRESHOLD = 0.6
INLINE_SHOW_MAX = 5        # inline (-i) results with more records are saved to a file (truncated: true)
WARNINGS_SHOW_MAX = 20     # extrapolation warnings in the JSON (the total is extrapolation_warning_count)


def _predict_classification(
    pipe: Any,
    X_frame: Any,
    target_name: str,
    df_result: pd.DataFrame,
    is_inline: bool,
    low_conf_warnings: List[str],
    decision_threshold: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    """
    Adds pred_<t> (label), proba_<t>_<class> (per-class probability) and conf_<t>
    (max probability) columns, and returns the JSON summary for a classification target.
    Binary models trained with --positive-class use the tuned decision threshold.
    """
    classes = [str(c) for c in pipe.classes_]
    proba = pipe.predict_proba(X_frame)
    conf = proba.max(axis=1)
    if decision_threshold and decision_threshold.get("positive_class") in classes and len(classes) == 2:
        pos = classes.index(decision_threshold["positive_class"])
        labels = np.where(proba[:, pos] >= float(decision_threshold["threshold"]), classes[pos], classes[1 - pos]).astype(str)
    else:
        labels = pipe.predict(X_frame).astype(str)

    df_result[f"pred_{target_name}"] = labels
    for i, c in enumerate(classes):
        df_result[f"proba_{target_name}_{c}"] = np.round(proba[:, i], 4)
    df_result[f"conf_{target_name}"] = np.round(conf, 4)

    n_low = int((conf < LOW_CONFIDENCE_THRESHOLD).sum())
    if n_low:
        low_conf_warnings.append(
            f"타겟 '{target_name}': 예측 확신도 {LOW_CONFIDENCE_THRESHOLD} 미만 {n_low}건 (전체 {len(conf)}건). "
            f"해당 행은 conf_{target_name} 컬럼으로 확인하십시오."
        )

    counts = pd.Series(labels).value_counts()
    summary: Dict[str, Any] = {
        "task": "classification",
        "classes": classes,
        "predicted_class_counts": {c: int(counts.get(c, 0)) for c in classes},
        "mean_confidence": round(float(conf.mean()), 4),
        "min_confidence": round(float(conf.min()), 4),
        "low_confidence_count": n_low,
    }
    if decision_threshold:
        summary["decision_threshold"] = {
            "positive_class": decision_threshold.get("positive_class"),
            "threshold": decision_threshold.get("threshold"),
        }
    if is_inline and len(labels) <= INLINE_SHOW_MAX:
        summary["predictions"] = labels.tolist()
        summary["probabilities"] = [
            {c: round(float(p), 4) for c, p in zip(classes, row)} for row in proba
        ]
        summary["confidences"] = [round(float(v), 4) for v in conf]
    return summary


def _fmt4(v: float) -> str:
    """Input values in warnings: 4 significant digits, so 3e-06 is not shown as 0.00."""
    return f"{float(v):.4g}"


_LAYOUT_ARTEFACT_RE = re.compile(r"^.*\(\s*[-+]?(\d[\d,]*\.?\d*|\.\d+)([eE][-+]?\d+)?\s*\)$")


def _typo_hints(missing: List[str], available: List[str], model_names: List[str], is_inline: bool) -> List[str]:
    """A missing model input and the given name that looks like a typo of it, worded the way the user fixes it:
    "'온도'이(가) 없습니다. -i의 [후보] 온두 ...": a close spelling may be another quantity, so the user decides
    (never [추천]). Only given names that are not model inputs count as typos."""
    where = "-i의" if is_inline else "입력 데이터 열"
    known = set(model_names)
    # 'bar(3)' / '°C(192.58)': a unit row read as the header and a data row as units, not a name to rename
    extra = [c for c in available if c not in known and not _LAYOUT_ARTEFACT_RE.match(str(c))]
    out = []
    found: Dict[str, List[str]] = {}
    for m in missing:
        cands = difflib.get_close_matches(m, extra, n=3, cutoff=0.5)
        # a given name without unit next to a model name with one ('온두' / '온도(°C)'): compared without the unit
        bare = [e for e in extra if strip_unit(e) == str(e).strip()]
        cands += [c for c in difflib.get_close_matches(strip_unit(m), bare, n=3, cutoff=0.5) if c not in cands]
        alias = build_alias_map(extra).get(_alias_key(m))
        if alias and alias not in cands:
            cands.insert(0, alias)
        found[m] = cands
    # one given name close to two or more missing inputs ('솔도' for '속도' and '습도'): one [후보] for that name
    wanted: Dict[str, List[str]] = {}
    for m, cands in found.items():
        for c in cands:
            wanted.setdefault(c, []).append(m)
    shared = {c: ms for c, ms in wanted.items() if len(ms) > 1}
    for c, ms in shared.items():
        out.append(f"{where} '{c}'은(는) 여러 모델 입력과 비슷합니다. '{c}'을(를) 어느 이름으로 고칠지 사용자에게 물으십시오: "
                   f"[후보] {', '.join(unit_marked(ms, c))}")
    for m in missing:
        cands = [c for c in found[m] if c not in shared]
        if len(cands) == 1 and other_unit(cands[0], m):
            # '압력(MPa)' for the model's '압력(bar)': a rename would read the MPa value as bar (no [추천])
            out.append(unit_hint(cands[0], m, where))
        elif cands:
            shown = unit_marked(cands, m)
            out.append(f"'{m}'이(가) 없습니다. {where} [후보] {', '.join(shown)} 중 '{m}'(으)로 고칠 이름을 사용자에게 물으십시오")
    return out


def _predict_command(args: argparse.Namespace, layout: str = "", inline: Optional[str] = None) -> str:
    """The same prediction command without --header-row/--unit-row/--data-row (plus `layout` flags when given),
    or with -i <inline> instead of the input. Paths are absolute and quoted (a trailing backslash or a relative
    path must not break the command on Windows)."""
    parts = [f"python {quote_arg(os.path.abspath(__file__))}", "-m", quote_arg(os.path.abspath(args.model))]
    if getattr(args, "targets", None):
        parts += ["-y", quote_arg(args.targets)]
    if inline is not None:
        parts += ["-i", quote_arg(inline)]
    else:
        parts += ["-f", quote_arg(os.path.abspath(args.file))]
    if layout:
        parts.append(layout.strip())
    if getattr(args, "save", None):
        parts += ["--save", quote_arg(os.path.abspath(args.save))]
    if getattr(args, "overwrite", False):
        parts.append("--overwrite")
    return " ".join(parts)


def _inputs_present(needed: List[str], columns: List[str]) -> bool:
    """Every model input is a column, exactly or as a unique unit-less column (the align_input_columns rule)."""
    cols = set(columns)
    alias = build_alias_map(columns)
    for f in needed:
        src = f if f in cols else alias.get(_alias_key(f))
        if src is None or (src != f and strip_unit(src) != src.strip()):
            return False
    return True


def _first_row_text_share(df: pd.DataFrame, numeric: List[str]) -> Optional[float]:
    """Share of text (not a number, not a missing-value mark) among the filled numeric-input cells of the first data
    row that has a filled numeric input (a row of blank inputs is skipped, up to FIRST_ROW_SCAN rows); None when the
    file has no rows, misses a numeric input or no such row is found."""
    if df.empty:
        return None
    alias = build_alias_map([str(c) for c in df.columns])
    cols = []
    for f in numeric:
        src = f if f in df.columns else alias.get(_alias_key(f))
        if src is None:
            return None
        cols.append(src)
    for i in range(min(FIRST_ROW_SCAN, len(df))):
        vals = [v for v in (df[c].iloc[i] for c in cols)
                if pd.notna(v) and not _blank_mark(str(v).strip())]
        if vals:
            return float(coerce_numeric(pd.Series(vals, dtype=object)).isna().mean())
    return None


FIRST_ROW_SCAN = 10


def _first_row_numeric(df: pd.DataFrame, numeric: List[str]) -> bool:
    """The first data row is data: at least half of its filled numeric inputs are numbers (one '측정불가' among
    numbers is data; a unit row read as data, '°C' 'bar', is not). The complement of _first_row_is_text."""
    share = _first_row_text_share(df, numeric)
    return share is not None and share <= 0.5


def _model_inputs(model_items: List[Dict[str, Any]]) -> Tuple[List[str], List[str]]:
    """(every raw model input, the numeric ones) over all target models."""
    needed = list(dict.fromkeys(c for it in model_items for c in model_raw_inputs(it["meta"])))
    numeric = list(dict.fromkeys(c for it in model_items for c in model_raw_inputs(it["meta"])
                                 if c not in it["meta"]["feature_schema"].get("categorical", {})
                                 and c not in it["meta"]["feature_schema"].get("datetime", {})))
    return needed, numeric


def _has_layout(args: argparse.Namespace) -> bool:
    return any(getattr(args, k, None) is not None for k in ("header_row", "unit_row", "data_row"))


def _layout_fix(args: argparse.Namespace, model_items: List[Dict[str, Any]],
                df_input: Optional[pd.DataFrame] = None) -> Tuple[Optional[str], List[int]]:
    """The layout of a prediction file whose layout options (often copied from training) or automatic reading miss
    the model columns, when exactly ONE layout reads every model input: (layout option text, '' = no options, []).
    (None, rows) when several top rows hold every model input (the tool cannot pick), (None, []) when none does.
    - layout options given, and the file read WITHOUT them has every model input and data in the first data row
      -> no options; a unit or notes row right below that header -> the options that skip it (_fit_unit_layout)
    - the model inputs are all in ONE other top row (two-level / merged header, title rows) -> --header-row N
      (with the unit row below it when there is one)
    - read without options, the first data row is a unit or notes row -> the options that skip it"""
    needed, numeric = _model_inputs(model_items)
    has_layout = _has_layout(args)
    read_ok = df_input is not None and _inputs_present(needed, [str(c) for c in df_input.columns])
    if has_layout and not read_ok:      # options that read the file with every model input are kept as given
        try:
            auto = load_data_file(args.file)
        except Exception:
            auto = None
        if auto is not None and _inputs_present(needed, [str(c) for c in auto.columns]):
            layout = _fit_unit_layout(args.file, needed, None, int(auto.attrs.get("header_row", 1)) + 1, numeric)
            if layout:
                return layout, []
            if not auto.attrs.get("unit_row_below") and (_first_row_numeric(auto, numeric) or _row_has_data(auto, 0)):
                return "", []
    if df_input is None:
        return None, []
    if not read_ok:
        cands = df_input.attrs.get("header_candidates") or []
        hits = [c for c in cands if find_header_layout([c], needed, model_names=True) is not None]
        rows = sorted({int(c[0]) for c in hits})
        if len(rows) > 1:
            return None, rows
        if len(rows) == 1 and rows[0] != getattr(args, "header_row", None):
            row, unit = find_header_layout(hits, needed, model_names=True)
            layout = f"--header-row {row}"
            if unit:
                # a model trained without units needs the unit row skipped (--data-row), not joined to the names
                layout = _fit_unit_layout(args.file, needed, row, unit, numeric) or f"{layout} --unit-row {unit}"
            return layout, []
        return None, []
    if not has_layout and not df_input.empty and (df_input.attrs.get("unit_row_below")
                                                  or _first_row_is_text(df_input, numeric)):
        layout = _fit_unit_layout(args.file, needed, None, int(df_input.attrs.get("header_row", 1)) + 1, numeric)
        if layout:
            return layout, []
    return None, []


_LAYOUT_OPT_RE = re.compile(r"--(header-row|unit-row|data-row) (\d+)")


def _apply_layout(args: argparse.Namespace, layout: str) -> None:
    """Sets args.header_row / unit_row / data_row to the option text of _layout_fix (absent = None)."""
    vals: Dict[str, Optional[int]] = {"header_row": None, "unit_row": None, "data_row": None}
    for flag, v in _LAYOUT_OPT_RE.findall(layout):
        vals[flag.replace("-", "_")] = int(v)
    for k, v in vals.items():
        setattr(args, k, v)


def _layout_note(df: pd.DataFrame, had_options: bool) -> str:
    """The recovered line of a layout the tool chose: '레이아웃 옵션이 이 파일과 달라 머리글 N행으로 읽음(, 단위 M행)
    (, 데이터 K행부터)' when layout options were given, else '머리글 N행(, 단위 M행)(, 데이터 K행부터)(으)로 읽음'."""
    parts = [f"머리글 {df.attrs.get('header_row', 1)}행"]
    if df.attrs.get("unit_row"):
        parts.append(f"단위 {df.attrs['unit_row']}행")
    if df.attrs.get("data_row"):
        parts.append(f"데이터 {df.attrs['data_row']}행부터")
    if had_options:
        return f"레이아웃 옵션이 이 파일과 달라 {parts[0]}으로 읽음" + "".join(f", {x}" for x in parts[1:])
    return ", ".join(parts) + ("로 읽음" if parts[-1].endswith("부터") else "으로 읽음")


def _skipped_rows_warning(args: argparse.Namespace, df: pd.DataFrame) -> Optional[str]:
    """The rows between the header (or unit row) and the first data row the tool chose (--data-row), with their
    source row numbers: they are not predicted, so the user hears about them (warnings)."""
    start = df.attrs.get("data_row")
    if not start:
        return None
    first = int(df.attrs.get("unit_row") or df.attrs.get("header_row", 1)) + 1
    rows = list(range(first, int(start)))
    if not rows:
        return None
    try:
        raw = read_raw_table(args.file)
        shown = " / ".join(", ".join(c for c in _row_cells(raw, r - 1))[:60] for r in rows[:3] if r - 1 < len(raw))
    except Exception:
        shown = ""
    return (f"파일 {', '.join(map(str, rows))}행은 데이터가 아닌 단위·메모 행으로 보고 건너뛰었습니다 (예측하지 않음)"
            + (f": {shown}" if shown else ""))


def _layout_candidates_text(rows: List[int]) -> str:
    return (f"모델 입력 열이 모두 들어 있는 행이 파일에 여러 개 있습니다 ({', '.join(map(str, rows))}행). 어느 행이 머리글인지 "
            f"사용자에게 물으십시오. [후보] " + ", ".join(f"--header-row {r}" for r in rows[:10]))


def _layout_hint(args: argparse.Namespace, model_items: List[Dict[str, Any]], df_input: Optional[pd.DataFrame] = None,
                 missing: Optional[List[str]] = None) -> str:
    """Why a file still misses some model columns after _load_prediction_file (no single layout reads every model
    input): [후보] rows when several rows hold every input; the missing names in another top row -> [추천] that
    layout (only some inputs there, so the tool does not switch by itself)."""
    _, rows = _layout_fix(args, model_items, df_input)
    if rows:
        return _layout_candidates_text(rows)
    if df_input is not None and missing:
        needed, numeric = _model_inputs(model_items)
        found = find_header_layout(df_input.attrs.get("header_candidates"), missing, model_names=True)
        if found is not None and found[0] != getattr(args, "header_row", None):
            row, unit = found
            where = f"파일 {row}행에 이 이름들이 있습니다 (두 줄 머리글, 병합된 머리글이나 제목 행)"
            if unit is None and not _has_layout(args):
                return f"{where}. [추천] {option_retry(str(row), '--header-row')}"
            layout = f"--header-row {row}"
            if unit:
                where += f", {unit}행은 단위 행입니다"
                layout = (_fit_unit_layout(args.file, needed, row, unit, numeric) if getattr(args, "file", None)
                          else "") or f"{layout} --unit-row {unit}"
            return f"{where}. [추천] " + _predict_command(args, layout=layout)
    return ""


def _load_prediction_file(args: argparse.Namespace, model_items: List[Dict[str, Any]]) -> pd.DataFrame:
    """Reads the prediction file. A layout that does not fit the file (training layout options on another file, a
    header in another row, a unit row read as data) is recovered in place (JIT, AGENTS.md) when exactly one layout
    reads every model input: recovered '레이아웃 옵션이 이 파일과 달라 머리글 N행으로 읽음'. Several candidate rows
    stop with [후보]; no layout that fits stops with the loader's / the column check's message."""
    decimal_comma = any(bool(it["meta"].get("decimal_comma")) for it in model_items)
    had_options = _has_layout(args)

    def _read() -> pd.DataFrame:
        return load_data_file(file_path=args.file, header_row=args.header_row, unit_row=args.unit_row,
                              data_row=args.data_row, decimal_comma=decimal_comma)
    try:
        df = _read()
    except InputError as e:
        layout, rows = _layout_fix(args, model_items)
        if layout is None:
            if rows:
                raise InputError(f"{str(e).split('[추천]')[0].rstrip()} {_layout_candidates_text(rows)}")
            raise
        _apply_layout(args, layout)
        df = _read()
        note_recovered(_layout_note(df, had_options))
        df.attrs["skipped_rows_warning"] = _skipped_rows_warning(args, df)
        return df
    layout, rows = _layout_fix(args, model_items, df)
    if rows:
        raise InputError(_layout_candidates_text(rows))
    if layout is not None:
        _apply_layout(args, layout)
        df = _read()
        note_recovered(_layout_note(df, had_options))
        df.attrs["skipped_rows_warning"] = _skipped_rows_warning(args, df)
    return df


def _row_has_data(df: pd.DataFrame, i: int) -> bool:
    """Row i of a data frame has a number or a date in some filled cell."""
    return any(not _blank_mark(_cell_text(v)) and _is_data_value(v) for v in df.iloc[i].tolist())


def _notes_row(df_raw: pd.DataFrame, j: int) -> bool:
    """Row j (0-based) is a notes row: filled, no number or date, and some cell longer than a unit
    ('비고: 2026년 1월 측정'). A row of short text ('측정불가', 'A') is not one."""
    cells = [v for v in df_raw.iloc[j].tolist() if not _blank_mark(_cell_text(v))]
    return bool(cells) and not any(_is_data_value(v) for v in cells) and any(not _unit_text(_cell_text(v)) for v in cells)


_NUMBERISH_RE = re.compile(r"^[<>~]?\s*[-+]?\.?\d")     # '<0.1', '12%', '1.5e3': a measured value written as text


def _text_input_row(df: pd.DataFrame, i: int, numeric: List[str]) -> bool:
    """Row i of a data frame holds text in its numeric inputs: most filled cells (missing marks are blank) are text
    that is not a number, a date or a value written as text ('<0.1', '12%'): a unit row the unit rule did not
    accept ('kg/cm2 G'), or a row of labels."""
    alias = build_alias_map([str(c) for c in df.columns])
    vals = []
    for f in numeric:
        src = f if f in df.columns else alias.get(_alias_key(f))
        if src is not None and i < len(df) and not _blank_mark(_cell_text(df[src].iloc[i])):
            vals.append(df[src].iloc[i])
    text = [v for v in vals if not _is_data_value(v) and not _NUMBERISH_RE.match(_cell_text(v))]
    return bool(vals) and len(text) * 2 > len(vals)


def _fit_unit_layout(path: str, needed: List[str], header_row: Optional[int], unit_row: int,
                     numeric: Optional[List[str]] = None) -> str:
    """Layout options for a file whose row `unit_row` (1-based, below the header) is not data: '--unit-row M'
    (a unit row by _looks_like_unit_row, model trained with units '반응온도(°C)'), else '--data-row M+1' (the row
    skipped: a unit row with the unit-less names of a model trained without units, a notes row, or a row of text in the
    numeric inputs `numeric` (_text_input_row) over a first data row of numbers), else '' (the row is data: numbers in
    some cells, '측정불가, 측정불가, A, 61.66', or only missing marks and categories). Every model input must be a
    column. `header_row` (or None) is put in front."""
    head = f"--header-row {header_row} " if header_row else ""
    try:
        raw = read_raw_table(path)
    except Exception:
        return ""
    hdr = header_row - 1 if header_row and header_row > 1 else sniff_header_index(raw)
    j = unit_row - 1
    if not hdr < j < len(raw) - 1:
        return ""
    is_unit = _looks_like_unit_row(raw, j, hdr)
    tries: List[Tuple[Dict[str, Any], str, bool]] = []
    # a row of Korean unit words ('도씨, 킬로그램') is read as units only when the model's names hold them
    # ('온도(도씨)'): the tool's own guess never takes other Hangul words for units ('점검중, 센서고장' is data)
    if is_unit or unit_row_verdict(raw, j, hdr, given=True)[0]:
        tries.append(({"unit_row": unit_row, "data_row": None}, f"--unit-row {unit_row}", False))
    skip = is_unit or _notes_row(raw, j)
    check_first = False
    # Fallback: text in the numeric inputs of the row (a unit row the rule did not accept, 'kg/cm2 G') -> skip it
    # when the next row reads as data. A row with a number or date ('점검중, 센서고장, 45') or with text that is
    # not unit-shaped ('점검중, 센서고장') is data (its marks are imputed and reported as missing inputs).
    if not skip and numeric and not _row_has_data(raw, j) and _unit_shaped_row(raw, j):
        try:
            auto: Optional[pd.DataFrame] = _layout_frame(raw, hdr, None, None)
        except Exception:
            auto = None
        rows = (auto.attrs.get("source_rows") or []) if auto is not None else []
        if unit_row in rows and _text_input_row(auto, rows.index(unit_row), numeric):
            skip = check_first = True
    if skip:
        tries.append(({"unit_row": None, "data_row": unit_row + 1}, f"--data-row {unit_row + 1}", check_first))
    for opts, layout, first in tries:
        try:
            df = _layout_frame(raw, hdr, **opts)
        except Exception:
            continue
        if _inputs_present(needed, [str(c) for c in df.columns]) and \
                (not first or _first_row_numeric(df, numeric or [])):
            return head + layout
    return ""


def _unit_shaped_row(df_raw: pd.DataFrame, j: int) -> bool:
    """Every filled cell of row j (missing marks are blank) is unit-shaped (_unit_shaped)."""
    cells = [_cell_text(v) for v in df_raw.iloc[j].tolist() if not _blank_mark(_cell_text(v))]
    return bool(cells) and all(_unit_shaped(c) for c in cells)


def _first_row_is_text(df: pd.DataFrame, numeric: List[str]) -> bool:
    """Most filled numeric-input cells of the first data row are text that is not a missing-value mark (a unit row,
    a notes row read as data)."""
    share = _first_row_text_share(df, numeric)
    return share is not None and share > 0.5


def _offer_median_fill(args: argparse.Namespace, df_input: pd.DataFrame, model_items: List[Dict[str, Any]]) -> None:
    """-i (one record, every key a model name) without some numeric model inputs: never filled by the tool (the user's
    value is a fact the tool cannot know). Stops with needs_user_decision and next_commands[0] = the same command with
    the missing inputs set to their training medians (values shown); run only when the user agrees. Other cases
    (several records, unknown or misspelled keys, a missing text / date input) go on to the usual messages."""
    if len(df_input) != 1:
        return
    needed: List[str] = []
    known: List[str] = []
    medians: Dict[str, float] = {}
    text_inputs: set = set()
    for it in model_items:
        meta = it["meta"]
        schema = meta.get("feature_schema") or {}
        derived = [e.split("=", 1)[0].strip() for e in meta.get("derived_evals", []) if "=" in e]
        raw = model_raw_inputs(meta)
        needed += raw
        known += raw + list(meta.get("features_required", [])) + derived
        text_inputs |= set(schema.get("categorical", {})) | set(schema.get("datetime", {}))
        for c in raw:
            st = (meta.get("feature_stats") or {}).get(c) or {}
            if c not in medians and st.get("median") is not None:
                medians[c] = float(st["median"])
    needed = list(dict.fromkeys(needed))
    probe = df_input.copy()
    with contextlib.redirect_stderr(io.StringIO()):
        align_input_columns(probe, list(dict.fromkeys(known)), [])
    if any(str(c) not in set(known) for c in probe.columns):
        return                                  # a misspelled / unknown key: check_inline_unknown explains it
    missing = [c for c in needed if c not in probe.columns]
    if not missing or any(c in text_inputs or c not in medians for c in missing):
        return
    fill = {c: float(f"{medians[c]:.6g}") for c in missing}
    pairs = dict(df_input.attrs.get("kv_pairs") or {str(k): v for k, v in df_input.iloc[0].to_dict().items()})
    pairs.update({c: f"{v:g}" for c, v in fill.items()})
    shown = ", ".join(f"{c}={v:g}" for c, v in fill.items())
    raise InputError(
        f"-i에 모델 입력 {', '.join(repr(c) for c in missing)}이(가) 없어 예측하지 않았습니다. 학습 데이터의 중앙값({shown})으로 채워 "
        f"예측할 수 있지만, 실제 조건과 다르면 예측도 달라집니다. 사용자에게 실제 값을 묻거나, 중앙값으로 채울지 확인하십시오.",
        extra={"missing_inputs": missing, "fill_values": fill,
               "next_commands": [{"purpose": f"빠진 입력을 학습 중앙값({shown})으로 채워 예측 (사용자가 원할 때만)",
                                  "command": _predict_command(args, inline=kv_text(pairs)), "needs_user_confirm": True}]},
        error_type="needs_user_decision")


def execute_predict(args: argparse.Namespace) -> Dict[str, Any]:
    """Orchestrates 1-Turn declarative inference."""
    # 1. Resolve Targets & Models
    filter_targets = [t.strip() for t in args.targets.split(",") if t.strip()] if args.targets else None
    model_items = resolve_target_models(
        target_path=args.model,
        filter_targets=filter_targets
    )

    # --save format is checked before any work: only CSV and Excel can be written and read back (R1)
    if args.save:
        out_ext = os.path.splitext(args.save)[1].lower()
        if out_ext not in (".csv", ".xlsx", ".xls"):
            raise InputError(f"결과 파일(--save)은 .csv 또는 .xlsx만 저장합니다: '{os.path.basename(args.save)}'. "
                             f"[추천] {option_retry(os.path.splitext(args.save)[0] + '.csv', '--save')}")

    # 2. Resolve Input Data (Batch File or Inline Input)
    file_warnings: List[str] = []
    if args.input:
        df_input = parse_inline_input(args.input)
        is_inline = True
        source_desc = f"직접 입력 {len(df_input)}건"
    elif args.file:
        df_input = _load_prediction_file(args, model_items)
        is_inline = False
        source_desc = os.path.basename(args.file)
        file_warnings = layout_warnings(df_input, training=False)   # repeated header rows, totals rows
        if df_input.attrs.get("skipped_rows_warning"):
            file_warnings.insert(0, df_input.attrs["skipped_rows_warning"])
    else:
        raise InputError("추론할 입력 데이터(--file 또는 --input)를 지정해야 합니다.")

    df_result = df_input.copy()
    predictions_summary = {}
    drift_warnings: List[str] = []
    low_conf_warnings: List[str] = []
    category_warnings: List[str] = []
    missing_counts: Dict[str, int] = {}
    column_aliases: Dict[str, str] = {}
    map_warnings: List[str] = []
    map_notes: List[str] = []
    snapshot_cache: Dict[str, Any] = {}
    known_names: List[str] = []       # every name a model takes (-i keys beyond these are refused)
    known_inputs: List[str] = []
    numeric_inputs: List[str] = []    # numbers written with '%' in these are read without it (percent_warning)

    if is_inline:
        _offer_median_fill(args, df_input, model_items)

    # 3. Perform Inference per Target
    for item in model_items:
        target_name = item["target"]
        pipe = item["pipeline"]
        meta = item["meta"]

        req_features = meta.get("features_required", [])
        derived_evals = meta.get("derived_evals", [])
        feature_stats = meta.get("feature_stats", {})
        if not req_features:
            raise InputError(
                f"타겟 '{target_name}'의 메타데이터(model_*.json)에 features_required가 없습니다. "
                f"모델과 같은 폴더의 model_<타겟>.json 파일을 확인하십시오."
            )
        derived_names = [e.split("=", 1)[0].strip() for e in derived_evals if "=" in e]

        # Accept unit-less input column names, then replay derived features
        raw_needed = model_raw_inputs(meta)
        known_names += raw_needed + req_features + derived_names
        known_inputs += raw_needed
        numeric_inputs += [c for c in raw_needed if c not in meta["feature_schema"].get("categorical", {})
                           and c not in meta["feature_schema"].get("datetime", {})]
        if is_inline:
            check_inline_aliases(df_input, list(dict.fromkeys(raw_needed + req_features)))
        column_aliases.update(align_input_columns(df_input, list(dict.fromkeys(raw_needed + req_features)), derived_names))
        if is_inline:
            num_raw = {c for c in raw_needed if c not in meta["feature_schema"].get("categorical", {})
                       and c not in meta["feature_schema"].get("datetime", {})}
            check_comma_numbers(df_input.attrs.get("kv_pairs"), lambda k: column_aliases.get(k, k) in num_raw, "-i",
                                lambda fixed: _predict_command(args, inline=kv_text(fixed)))
        missing_raw = [c for c in raw_needed if c not in df_input.columns]
        if missing_raw and any(c not in req_features for c in missing_raw):
            available_cols = [str(c) for c in df_input.columns]
            hint = "" if is_inline else _layout_hint(args, model_items, df_input, missing_raw)
            raise InputError(
                f"타겟 '{target_name}' 추론에 필요한 입력 '{', '.join(missing_raw)}'이(가) 입력 데이터에 없습니다 "
                f"(파생변수 계산에 쓰는 원본 입력 포함). [입력 데이터 컬럼 목록]: {available_cols}"
                + (f" {hint}" if hint else ""))
        # '11.9%' next to '0.119' in one numeric input: no scale is sure (the same check as training)
        check_percent_mix(df_input, [c for c in numeric_inputs if c in df_input.columns], "입력 데이터 ")
        # 'c' / ' C ' / full-width for the training category 'C' (match_categories, the same as --sweep / --find)
        align_categories(df_input, meta["feature_schema"])
        df_evaled = apply_inference_evals(df_input, derived_evals)

        # Check required features
        missing_feats = [f for f in req_features if f not in df_evaled.columns]
        if missing_feats:
            available_cols = [str(c) for c in df_evaled.columns]
            hint = "" if is_inline else _layout_hint(args, model_items, df_input, missing_feats)
            if not hint:
                hint = " / ".join(_typo_hints(missing_feats, available_cols, req_features + raw_needed, is_inline))
            raise InputError(
                f"타겟 '{target_name}' 추론에 필요한 필수 피처 '{', '.join(missing_feats)}'이(가) 입력 데이터에 누락되었습니다. "
                + f"[입력 데이터 컬럼 목록]: {available_cols}"
                + (f"\n{hint}" if hint else "")
            )

        schema = meta["feature_schema"]
        X_model = prepare_features(df_evaled, schema)
        numeric_view = X_model[[c for c in schema["numeric"]]]
        owner = input_column_owner(schema)

        # Missing / unparseable inputs are imputed by the pipeline -> tell the agent explicitly
        for col, n_miss in X_model.isna().sum().items():
            if n_miss:
                raw_col = owner.get(col, col)
                missing_counts[raw_col] = max(missing_counts.get(raw_col, 0), int(n_miss))

        # Categories never seen in training are encoded as "none of the known categories"
        for col, known in schema.get("categorical", {}).items():
            vals = X_model[col].dropna()
            unseen = sorted(set(vals) - set(known))
            if unseen:
                category_warnings.append(
                    f"범주형 피처 '{col}'에 학습 때 없던 값 {unseen[:5]}이(가) {int(vals.isin(unseen).sum())}건 있습니다. "
                    f"해당 행의 예측 신뢰도가 낮을 수 있습니다. [학습 범주]: {known[:10]}"
                )
        for col, rng in (meta.get("datetime_stats") or {}).items():
            dt = parse_datetime(df_evaled[col]).dropna()
            if not dt.empty:
                lo, hi = pd.Timestamp(rng["min"]), pd.Timestamp(rng["max"])
                n_out = int(((dt < lo) | (dt > hi)).sum())
                if n_out:
                    drift_warnings.append(
                        f"날짜 피처 '{col}' 학습 기간({lo.date()} ~ {hi.date()}) 밖의 값 {n_out}건 (외삽)"
                    )

        # Data Drift / Out-of-bounds Check (with quantitative severity ratio)
        if feature_stats:
            for fname in numeric_view.columns:
                if fname in feature_stats:
                    f_min = feature_stats[fname].get("min")
                    f_max = feature_stats[fname].get("max")
                    if f_min is not None and f_max is not None:
                        col_series = numeric_view[fname].dropna()
                        if not col_series.empty:
                            c_min = float(col_series.min())
                            c_max = float(col_series.max())
                            f_range = (f_max - f_min) if f_max > f_min else (abs(f_max) or 1.0)
                            n_out = int(((col_series > f_max) | (col_series < f_min)).sum())

                            if c_max > f_max:
                                over_pct = round(float((c_max - f_max) / f_range * 100), 1)
                                drift_warnings.append(
                                    f"입력 '{fname}' 학습 범위 초과: 학습 최대={_fmt4(f_max)}, 입력 최대={_fmt4(c_max)} (범위의 +{over_pct}%, 범위 밖 {n_out}건)"
                                )
                            if c_min < f_min:
                                under_pct = round(float((f_min - c_min) / f_range * 100), 1)
                                drift_warnings.append(
                                    f"입력 '{fname}' 학습 범위 미달: 학습 최소={_fmt4(f_min)}, 입력 최소={_fmt4(c_min)} (범위의 -{under_pct}%, 범위 밖 {n_out}건)"
                                )

        X_input = X_model

        # Execute Prediction
        task = meta.get("task", "regression")
        estimator = pipe.named_steps.get("regressor") or pipe.named_steps.get("classifier") or pipe
        desc = meta.get("description", "")

        if task == "classification":
            predictions_summary[target_name] = {
                "task": "classification",
                "model": meta.get("champion_model", type(estimator).__name__),
                "description": desc,
                "features_used": req_features,
                **_predict_classification(pipe, X_input, target_name, df_result, is_inline, low_conf_warnings,
                                          meta.get("decision_threshold")),
            }
            continue

        preds = pipe.predict(X_input)

        pred_col_name = f"pred_{target_name}"
        df_result[pred_col_name] = round_pred(preds)

        # Record summary
        predictions_summary[target_name] = {
            "task": "regression",
            "model": meta.get("champion_model", type(estimator).__name__),
            "description": desc,
            "features_used": req_features,
            "mean_pred": round_pred(float(np.mean(preds))),
            "min_pred": round_pred(float(np.min(preds))),
            "max_pred": round_pred(float(np.max(preds)))
        }

        # Approximate prediction interval from the champion's out-of-fold residuals
        interval = meta.get("prediction_interval")
        lo = hi = None
        if interval:
            lo = preds + float(interval["residual_lower"])
            hi = preds + float(interval["residual_upper"])
            df_result[f"pred_lo_{target_name}"] = round_pred(lo)
            df_result[f"pred_hi_{target_name}"] = round_pred(hi)
            predictions_summary[target_name]["interval_level"] = interval["level"]
        if is_inline and len(preds) <= INLINE_SHOW_MAX:
            predictions_summary[target_name]["predictions"] = [round_pred(float(p)) for p in preds]
            if lo is not None:
                predictions_summary[target_name]["intervals"] = [
                    [round_pred(float(l)), round_pred(float(h))] for l, h in zip(lo, hi)
                ]

        # Measured-data comparison (map check): same conditions looked up in the training rows
        try:
            mc = compare_with_measured(meta, target_name, os.path.dirname(item["joblib_path"]), df_evaled, snapshot_cache)
        except Exception as e:  # the comparison must never break the prediction
            mc = {"note": f"타겟 '{target_name}': 실측 비교값 계산 실패 ({type(e).__name__}: {str(e)[:80]})"}
        if mc.get("note"):
            map_notes.append(mc["note"])
        if "values" in mc:
            vals = mc["values"]
            diff = preds - vals
            outside = ((vals < lo) | (vals > hi)) if lo is not None else np.zeros(len(vals), dtype=bool)
            outside = outside & np.isfinite(vals) & np.array(mc["warnable"], dtype=bool)
            df_result[f"map_{target_name}"] = round_pred(vals)
            df_result[f"map_diff_{target_name}"] = round_pred(diff)
            df_result[f"map_방식_{target_name}"] = [m + (" [외삽]" if e else "") for m, e in zip(mc["method"], mc["extrapolated"])]
            predictions_summary[target_name]["map_check_keys"] = mc["keys"]
            predictions_summary[target_name]["map_check_outside_interval"] = int(outside.sum())
            if is_inline and len(preds) <= INLINE_SHOW_MAX:
                predictions_summary[target_name]["map_check"] = [
                    {"value": round_pred(float(v)) if np.isfinite(v) else None,
                     "diff": round_pred(float(d)) if np.isfinite(d) else None,
                     "method": m, "basis": b, "extrapolated": e, "outside_interval": bool(o)}
                    for v, d, m, b, e, o in zip(vals, diff, mc["method"], mc["basis"], mc["extrapolated"], outside)]
            if outside.any():
                map_warnings.append(
                    f"타겟 '{target_name}': 실측 비교값이 90% 예측 범위 밖인 건 {int(outside.sum())}건 "
                    f"(모델 예측과 실측 데이터가 많이 다름, 해당 조건의 가까운 실측 행 확인 필요)")

    if is_inline:
        check_inline_unknown(df_input, known_names, known_inputs)

    # 4. Save Output if Requested or File-based
    inline_truncated = is_inline and len(df_result) > INLINE_SHOW_MAX
    out_path = args.save
    saved_file = None
    if out_path:
        saved_file = os.path.abspath(out_path)
    elif not is_inline and args.file:
        file_dir = os.path.dirname(os.path.abspath(args.file))
        file_stem = os.path.splitext(os.path.basename(args.file))[0]
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        saved_file = os.path.join(file_dir, f"{file_stem}_pred_{ts}.csv")
    elif inline_truncated:
        # more inline records than the JSON shows: the full result goes to a file (next to the .json input,
        # otherwise in the <data>_ml_data folder that holds the version folder)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        src = args.input.strip()
        if src.lower().endswith(".json") and os.path.isfile(src):
            saved_file = os.path.join(os.path.dirname(os.path.abspath(src)),
                                      f"{os.path.splitext(os.path.basename(src))[0]}_pred_{ts}.csv")
        else:
            version_dir = os.path.dirname(os.path.abspath(model_items[0]["joblib_path"]))
            saved_file = os.path.join(os.path.dirname(version_dir), f"inline_pred_{ts}.csv")

    if saved_file:
        warn_long_path(saved_file)
        os.makedirs(os.path.dirname(saved_file), exist_ok=True)
        ext = os.path.splitext(saved_file)[1].lower()
        if ext == ".xls":
            saved_file = os.path.splitext(saved_file)[0] + ".xlsx"
            ext = ".xlsx"
        saved_file = avoid_overwrite(saved_file, getattr(args, "overwrite", False), what="예측 결과 파일")
        if ext == ".xlsx":
            saved_file = save_with_lock_fallback(
                saved_file, lambda p: df_result.to_excel(p, index=False), what="결과 Excel 파일")
        else:
            # CSV Output with Windows Excel CP949 compatibility (utf-8-sig fallback)
            saved_file = write_csv_excel_compatible(df_result, saved_file)
        mark_committed(f"예측 결과 파일 {os.path.basename(saved_file)}", output_file=saved_file)

    # R1: the saved file is read back (rows and columns) before saying VERIFIED
    unverified: List[str] = []
    if saved_file:
        try:
            back = load_data_file(saved_file, header_row=1)
            if len(back) != len(df_result) or list(map(str, back.columns)) != list(map(str, df_result.columns)):
                unverified.append(f"저장한 결과 파일을 다시 읽으니 {len(back)}행 x {back.shape[1]}열 "
                                  f"(기대 {len(df_result)}행 x {df_result.shape[1]}열)")
        except Exception as e:
            unverified.append(f"저장한 결과 파일을 다시 읽지 못함: {type(e).__name__}: {e}")

    # 5. Build Response
    drift_warnings = list(dict.fromkeys(drift_warnings))  # same feature is checked once per target
    for w in drift_warnings[:WARNINGS_SHOW_MAX]:
        sys.stderr.write(f"[WARN] 학습 범위 밖: {w}\n")
    if len(drift_warnings) > WARNINGS_SHOW_MAX:
        sys.stderr.write(f"[WARN] 학습 범위 밖: 외 {len(drift_warnings) - WARNINGS_SHOW_MAX}건\n")

    for w in low_conf_warnings:
        sys.stderr.write(f"[WARN] 낮은 확신도: {w}\n")
    category_warnings = list(dict.fromkeys(category_warnings))
    for w in category_warnings:
        sys.stderr.write(f"[WARN] 학습 때 없던 종류: {w}\n")
    missing_warnings = [
        f"피처 '{c}': 빈 값 또는 숫자로 읽을 수 없는 값 {n}건 -> 학습 데이터 기준값(중앙값/결측 범주)으로 대치됨"
        for c, n in missing_counts.items()
    ]
    for w in missing_warnings:
        sys.stderr.write(f"[WARN] 빈 입력: {w}\n")
    for w in map_warnings:
        sys.stderr.write(f"[WARN] 실측 비교: {w}\n")
    file_warnings += [w for w in percent_warning(df_input, numeric_inputs) if w not in file_warnings]
    for w in file_warnings:
        sys.stderr.write(f"[WARN] {w}\n")

    def _brief(t: str) -> str:
        ps = predictions_summary[t]
        if ps.get("task") == "classification":
            return f"{t}({ps['predicted_class_counts']})"
        return f"{t}({ps['mean_pred']})"

    targets_brief = ", ".join(_brief(t) for t in predictions_summary)
    preview = (
        f"[예측 완료] 입력: {source_desc} | 예측 요약: [{targets_brief}]"
        + (f" | 저장: {os.path.basename(saved_file)}" if saved_file else "")
        + (f" | 외삽 경고 {len(drift_warnings)}건" if drift_warnings else "")
        + (f" | 저확신 경고 {len(low_conf_warnings)}건" if low_conf_warnings else "")
        + (f" | 결측 대치 {sum(missing_counts.values())}건" if missing_counts else "")
        + (f" | 미학습 범주 경고 {len(category_warnings)}건" if category_warnings else "")
        + (f" | 실측 비교 경고 {len(map_warnings)}건" if map_warnings else "")
    )
    agent_directive = (
        ("VERIFIED: " if not unverified else "UNVERIFIED: 결과 파일 확인 실패 (" + "; ".join(unverified) + "). 다시 실행하지 말고 unverified를 그대로 알리세요. ")
        + f"타겟 {len(model_items)}개를 예측했습니다. "
        + ("predictions_summary의 예측값과 회귀는 intervals(90% 예측 범위), 분류는 probabilities(종류별 확률)를 알리세요. "
           if not saved_file else "결과 파일 output_file 경로를 알리세요. ")
        + ("학습 범위 밖 입력이 있어 예측은 참고용입니다 (extrapolation_warnings). " if drift_warnings else "")
        + (f"extrapolation_warnings에는 {len(drift_warnings)}건 중 {WARNINGS_SHOW_MAX}건만 있으니 전체 건수도 알리세요. "
           if len(drift_warnings) > WARNINGS_SHOW_MAX else "")
        + ("분류 확신도가 낮은 행이 있습니다 (low_confidence_warnings, 확률 0.6 미만). " if low_conf_warnings else "")
        + ("빈 값이나 숫자가 아닌 입력을 학습 기준값으로 채웠습니다 (missing_input_warnings). " if missing_counts else "")
        + ("학습 때 없던 종류가 있습니다 (unseen_category_warnings). " if category_warnings else "")
        + ("실측 비교값이 예측 범위 밖입니다 (map_check_warnings). " if map_warnings else "")
        + ("warnings(데이터가 아닌 행, '%' 없이 읽은 숫자)를 빠짐없이 알리세요. " if file_warnings else "")
        + ("예측값마다 실측 비교값(map_check의 value, diff, method)을 '실측 비교값'으로 함께 알리세요. "
           if any("map_check_keys" in p for p in predictions_summary.values()) else "")
        + (f"직접 입력이 {len(df_result)}건으로 {INLINE_SHOW_MAX}건보다 많아 건별 예측은 JSON에 없습니다 (truncated). "
           f"전체 결과는 output_file({os.path.basename(saved_file)})로 알리세요. "
           if inline_truncated and saved_file else
           f"결과 파일: {os.path.basename(saved_file)}. " if saved_file else "")
        + ("위 경고를 함께 알리세요." if (drift_warnings or low_conf_warnings or missing_counts or category_warnings or map_warnings)
           else "다음: 예측 결과를 사용자에게 알리세요.")
    )

    result = {
        "status": "ok",
        "source": source_desc,
        "sample_count": len(df_result),
        "total_rows": len(df_result),
        "targets_predicted": list(predictions_summary.keys()),
        "output_file": saved_file,
        "verified": not unverified,
        **({"unverified": unverified} if unverified else {}),
        "predictions_summary": predictions_summary,
        "drift_warning": bool(drift_warnings),
        "extrapolation_warnings": drift_warnings[:WARNINGS_SHOW_MAX],
        "extrapolation_warning_count": len(drift_warnings),
        "extrapolation_warnings_truncated": len(drift_warnings) > WARNINGS_SHOW_MAX,
        "low_confidence_warning": bool(low_conf_warnings),
        "low_confidence_warnings": low_conf_warnings,
        "missing_input_warning": bool(missing_counts),
        "missing_input_counts": missing_counts,
        "missing_input_warnings": missing_warnings,
        "unseen_category_warning": bool(category_warnings),
        "unseen_category_warnings": category_warnings,
        "map_check_warning": bool(map_warnings),
        "map_check_warnings": map_warnings,
        **({"warnings": file_warnings} if file_warnings else {}),
        "preview": preview,
        "agent_directive": agent_directive
    }
    if is_inline:   # R2: per-record values are in the JSON only up to INLINE_SHOW_MAX records
        result["count"] = len(df_result)
        result["truncated"] = bool(inline_truncated)
        if inline_truncated:
            result["shown"] = 0
            if saved_file:
                result["full_result_file"] = saved_file
    if map_notes:
        result["map_check_notes"] = list(dict.fromkeys(map_notes))
    if column_aliases:
        result["column_aliases"] = column_aliases
    return result


# ==============================================================================
# 5-B. Model-based reverse lookup (--find / --vary)
# ==============================================================================
#
# "Which value of <vary> makes the model predict <target> = y0, the other inputs fixed?"
#   - <vary> is scanned over its training range in FIND_STEPS steps in one batch prediction.
#   - Crossings closer than FIND_MERGE x range are one answer (tree models predict in steps and add
#     near-duplicate crossings); a merged answer is reported as a span. Windows narrower than that and without an
#     answer are step noise and dropped.
#   - Operating windows: <vary> ranges where y0 lies inside the 90% prediction interval, i.e.
#     prediction in [y0 - residual_upper, y0 - residual_lower].
#   - Inputs the user did not fix: numeric -> training median (reported), categorical/date -> STOP_RETRY [후보].
#   - Measured cross-check: the same lookup on the training rows (model_dataset.csv) when the model has
#     <= 3 numeric inputs (table interpolation), plus the 5 nearest measured rows.
#   - No extrapolation: tree models are flat outside the training range.

FIND_STEPS = 2000
FIND_MERGE = 0.03     # answers / windows closer than 3% of the scanned range are merged (tree models step)


def _spans(xs: np.ndarray, ok: np.ndarray, gap: float) -> List[List[float]]:
    """Contiguous runs of ok=True as [from, to]; runs separated by less than gap are joined."""
    out: List[List[float]] = []
    for x, flag in zip(xs, ok):
        if not flag:
            continue
        if out and x - out[-1][1] <= gap:
            out[-1][1] = float(x)
        else:
            out.append([float(x), float(x)])
    return out


def _r6(v: float) -> float:
    """6 decimals, but small values keep 6 significant digits (an input of 3e-06 must not become 0)."""
    return round_pred(float(v), 6)


def _model_curve(pipe: Any, meta: Dict[str, Any], base: Dict[str, Any], vary: str, xs: np.ndarray) -> np.ndarray:
    df = pd.DataFrame({**{k: [v] * len(xs) for k, v in base.items()}, vary: xs})
    df = apply_inference_evals(df, meta.get("derived_evals", []))
    return np.asarray(pipe.predict(prepare_features(df, meta["feature_schema"])), dtype=float)


def _find_command(args: argparse.Namespace, find: Optional[str] = None, inline: Optional[str] = None) -> str:
    """The same --find command with a corrected --find or -i text (absolute, quoted paths)."""
    parts = [f"python {quote_arg(os.path.abspath(__file__))}", "-m", quote_arg(os.path.abspath(args.model))]
    parts += ["--find", quote_arg(find if find is not None else args.find), "--vary", quote_arg(args.vary)]
    i_text = inline if inline is not None else getattr(args, "input", None)
    if i_text:
        parts += ["-i", quote_arg(i_text)]
    return " ".join(parts)


def execute_find(args: argparse.Namespace) -> Dict[str, Any]:
    goal_df = parse_inline_input(args.find)
    if len(goal_df) != 1:
        # a JSON array of several objects: only the first would be used without a word (ml_map refuses it too)
        raise InputError(f"--find은(는) 한 줄('열=값')로 주십시오 (지금 {len(goal_df)}줄).")
    goal = goal_df.iloc[0].to_dict()
    if len(goal) != 1:
        raise InputError("--find에는 목표 하나만 주십시오 (예: --find \"생산효율=50\").")
    check_comma_numbers(goal_df.attrs.get("kv_pairs"), lambda k: True, "--find",
                        lambda fixed: _find_command(args, find=kv_text(fixed)))
    (goal_name, goal_val), = goal.items()
    try:
        y0 = float(goal_val)
    except (TypeError, ValueError):
        raise InputError(f"목표 값 '{goal_val}'이(가) 숫자가 아닙니다.")
    # the goal is resolved like ml_map --find: exact name, then a unique name without its unit ('수율' for '수율(1차)'
    # and '수율(최종)' -> [후보]); a different unit -> unit_hint
    items = resolve_target_models(target_path=args.model, filter_targets=[str(goal_name)],
                                  kind="목표(--find)", values=True)
    item = items[0]
    meta, pipe, target = item["meta"], item["pipeline"], item["target"]
    if other_unit(str(goal_name), str(target)):
        # '생산량(kg)=500' for the target '생산량(t)': 500 would be solved as 500 t (ml_map --find refuses it too)
        raise InputError(unit_hint(str(goal_name), str(target), "목표(--find)"))
    if meta.get("task", "regression") != "regression":
        raise InputError(f"타겟 '{target}'은(는) 분류 모델입니다. 모델 반대 찾기는 숫자 타겟(회귀)만 됩니다.")
    schema = meta["feature_schema"]
    raw_inputs = model_raw_inputs(meta)
    stats = meta.get("feature_stats", {})
    numeric_in = [f for f in raw_inputs if f not in schema.get("categorical", {}) and f not in schema.get("datetime", {})]

    given_df = parse_inline_input(args.input) if args.input else None
    if given_df is not None and len(given_df) != 1:
        raise InputError(f"-i은(는) 한 줄('열=값,열=값')로 주십시오 (지금 {len(given_df)}줄). 고정 입력이 여러 경우면 "
                         f"경우마다 따로 실행하십시오.")
    given = given_df.iloc[0].to_dict() if given_df is not None else {}
    # a --vary typo never suggests an input fixed by -i, an -i typo never the --vary input (a retry would stop)
    vary = resolve_column_names([args.vary.strip()], raw_inputs, "바꿀 입력(--vary)",
                                exclude=known_columns([str(k) for k in given], raw_inputs))[0][0]
    if vary not in numeric_in or vary not in stats:
        raise InputError(f"'{vary}'은(는) 숫자 입력이 아니라 반대 찾기로 바꿀 수 없습니다. [후보] {', '.join(numeric_in)}")
    names, _ = (resolve_column_names([str(k) for k in given], raw_inputs, "고정 입력(-i)", values=True,
                                     exclude=[vary]) if given else ([], {}))
    check_same_column([str(k) for k in given], names, "-i", {str(k): v for k, v in given.items()})
    fixed = {n: given[k] for k, n in zip(given, names)}
    resolved = dict(zip([str(k) for k in given], names))
    if given_df is not None:
        check_comma_numbers(given_df.attrs.get("kv_pairs"), lambda k: resolved.get(k) in numeric_in, "-i",
                            lambda fx: _find_command(args, inline=kv_text(fx)))
    if vary in fixed:
        raise InputError(f"'{vary}'는 --vary로 찾는 입력이라 -i에 함께 줄 수 없습니다.")

    filled: Dict[str, float] = {}
    for f in raw_inputs:
        if f == vary or f in fixed:
            continue
        if f in numeric_in and f in stats:
            filled[f] = float(stats[f]["median"])
        elif f in schema.get("categorical", {}):
            raise InputError(f"모델 입력 '{f}' 값이 필요합니다. -i에 '{f}=<값>'을 넣어 주십시오. "
                             f"{category_candidates(schema['categorical'][f])}")
        else:
            raise InputError(f"모델 입력 '{f}' 값이 필요합니다. -i에 '{f}=<값>'을 넣어 주십시오.")
    warnings: List[str] = []
    for f, v in fixed.items():
        if f in schema.get("categorical", {}):
            fixed[f] = fixed_category_value(f, v, schema["categorical"][f])
        if f not in numeric_in:
            continue
        st = stats.get(f)
        try:
            fv = float(v)
        except (TypeError, ValueError):
            raise InputError(f"고정 입력 '{f}'의 값 '{v}'이(가) 숫자가 아닙니다. -i에 '{f}=<숫자>'로 다시 지정하십시오.")
        fixed[f] = fv
        if st and (fv < st["min"] or fv > st["max"]):
            warnings.append(f"고정 입력 '{f}'={_fmt(fv)}이(가) 학습 범위({_fmt(st['min'])}~{_fmt(st['max'])}) 밖이라 답을 믿기 어렵습니다")
    base = {**fixed, **filled}

    lo, hi = float(stats[vary]["min"]), float(stats[vary]["max"])
    xs = np.linspace(lo, hi, FIND_STEPS + 1)
    ys = _model_curve(pipe, meta, base, vary, xs)
    merge = FIND_MERGE * (hi - lo)

    # answers: crossings, near-duplicates merged into a span
    answers = []
    cross = np.array(_crossings(xs, ys, y0))
    for span in _spans(cross, np.ones(len(cross), dtype=bool), merge):
        mid = (span[0] + span[1]) / 2
        answers.append({vary: _r6(mid), **({"from": _r6(span[0]), "to": _r6(span[1])} if span[1] > span[0] else {})})
    # verification: the curve must bracket the target right around every answer
    step = (hi - lo) / FIND_STEPS
    unverified = []
    for a in answers:
        a_lo, a_hi = a.get("from", a[vary]) - step, a.get("to", a[vary]) + step
        chk = _model_curve(pipe, meta, base, vary, np.linspace(max(lo, a_lo), min(hi, a_hi), 50))
        chk = np.array([chk.min(), chk.max()])
        if not (chk[0] <= y0 <= chk[1]):
            unverified.append(f"{vary}={_fmt(a[vary])} 근처에서 목표값을 지나지 않음")
        a["check"] = [round_pred(float(c)) for c in chk]

    # operating windows: target inside the 90% prediction interval
    iv = meta.get("prediction_interval")
    windows: List[Dict[str, float]] = []
    if iv:
        band = (y0 - float(iv["residual_upper"]), y0 - float(iv["residual_lower"]))
        for a, b in _spans(xs, (ys >= band[0]) & (ys <= band[1]), merge):
            if b - a < merge and not any(a <= ans[vary] <= b for ans in answers):
                continue   # a sliver touching the band without crossing the target: step noise
            windows.append({"from": _r6(a), "to": _r6(b)})
    else:
        warnings.append("예측 오차 정보가 없는 모델이라 운전 구간을 계산하지 않았습니다")

    # measured cross-check on the training rows
    measured: Dict[str, Any] = {}
    data = _load_snapshot(os.path.dirname(item["joblib_path"]), {})
    near: List[Dict[str, Any]] = []
    if data is not None and all(c in data.columns for c in [target, SNAPSHOT_ROW_COL] + [c for c in raw_inputs if c in numeric_in]):
        data = data.copy()
        data["__row"] = coerce_numeric(data[SNAPSHOT_ROW_COL]).fillna(-1).astype(int)
        excl = f"{SNAPSHOT_EXCL_PREFIX}{target}"
        if excl in data.columns:
            data = data[coerce_numeric(data[excl]).fillna(0) != 1]
        for c in schema.get("categorical", {}):
            if c in data.columns and c in base:
                hit = normalize_category(data[c]) == normalize_category(pd.Series([base[c]])).iloc[0]
                data = data[hit] if hit.any() else data
        data = data[coerce_numeric(data[target]).notna()]
        mkeys = [vary] + [f for f in numeric_in if f != vary]
        fixed_num = np.array([float(base[f]) for f in mkeys[1:]])
        if len(mkeys) <= MAX_KEYS:
            try:
                tab = Table(data, mkeys, [target], [])
                def _measured(x: float) -> float:
                    q = np.concatenate([[x], fixed_num])
                    if tab.mode == "grid":
                        return float(tab.at(q)["values"][target])
                    v, _ = tab._scattered(q)            # inside the measured area only (no trend line)
                    return float(v[0]) if v is not None else np.nan
                mys = np.array([_measured(x) for x in xs])
                m_cross = np.array(_crossings(xs, mys, y0))
                m_ans = _spans(m_cross, np.ones(len(m_cross), dtype=bool), merge)
                measured = {"method": "학습 데이터 표 보간 (ml_map 방식)", "count": len(m_ans),
                            "solutions": [_r6((a + b) / 2) for a, b in m_ans],
                            "reachable": {"min": _r6(np.nanmin(mys)), "max": _r6(np.nanmax(mys))}}
                if tab.mode == "scattered" and len(m_ans) > 3:
                    warnings.append(f"실측 기준 답이 {len(m_ans)}개로 많습니다 (측정 잡음). 모델 답을 기준으로 보십시오")
            except InputError as e:
                measured = {"note": f"실측 기준 반대 찾기를 하지 못했습니다: {e}"}
        else:
            measured = {"note": f"모델 숫자 입력이 {len(numeric_in)}개라 실측 기준 반대 찾기는 생략하고 가까운 실측 행만 보여 줍니다"}
        span_y = max(float(np.ptp(coerce_numeric(data[target]))), 1e-12)
        spans = [span_y] + [max(float(stats[f]["max"]) - float(stats[f]["min"]), 1e-12) for f in mkeys[1:]]
        near = nearest_rows(data, [target] + mkeys[1:], np.concatenate([[y0], fixed_num]), np.array(spans),
                            list(dict.fromkeys(mkeys + [target])))
        if measured.get("solutions") and answers:
            far = [a[vary] for a in answers if min(abs(a[vary] - m) for m in measured["solutions"]) > 0.05 * (hi - lo)]
            if far:
                warnings.append(f"모델 답 {[_fmt(v) for v in far]}이(가) 실측 기준 답 {[_fmt(v) for v in measured['solutions']]}과(와) "
                                f"많이 다릅니다. 실제 시험으로 확인하십시오")
    else:
        measured = {"note": "버전 폴더에 학습 데이터 사본(model_dataset.csv)이 없어 실측 기준 확인을 하지 못했습니다"}

    reach = (float(np.min(ys)), float(np.max(ys)))
    res: Dict[str, Any] = {
        "status": "ok", "command": "find", "target": target, "find": {target: y0}, "vary": vary,
        "fixed": {k: (float(v) if isinstance(v, (int, float, np.floating)) else v) for k, v in fixed.items()},
        "filled_defaults": {k: _r6(v) for k, v in filled.items()},
        "model": meta.get("champion_model"), "training_range": {vary: [_r6(lo), _r6(hi)]},
        "count": len(answers), "solutions": answers, "operating_windows": windows,
        "reachable": {"min": _r6(reach[0]), "max": _r6(reach[1]), "within": f"{vary} {_fmt(lo)}~{_fmt(hi)}"},
        "measured": measured, "nearest_rows": near, "verified": not unverified,
    }
    if filled:
        warnings.insert(0, "주지 않은 입력은 학습 데이터 중앙값으로 고정했습니다: "
                        + ", ".join(f"{k}={_fmt(v)}" for k, v in filled.items()))
    if warnings:
        res["warnings"] = warnings
    if unverified:
        res["unverified"] = unverified
    win_txt = ", ".join(f"{_fmt(w['from'])}~{_fmt(w['to'])}" for w in windows) or "없음"
    if answers:
        ans_txt = ", ".join(_fmt(a[vary]) + (f"({_fmt(a['from'])}~{_fmt(a['to'])})" if "from" in a else "") for a in answers)
        res["preview"] = f"모델 기준 {target}={_fmt(y0)}이 되는 {vary}: {ans_txt} | 운전 구간(90% 오차 안): {win_txt}"
    else:
        res["preview"] = (f"모델 기준 {target}={_fmt(y0)}이 되는 {vary}: 찾은 값 없음 "
                          f"({vary} {_fmt(lo)}~{_fmt(hi)}에서 예측 {_fmt(reach[0])}~{_fmt(reach[1])}) | 운전 구간: {win_txt}")
    if measured.get("solutions") is not None:
        res["preview"] += " | 실측 기준: " + (", ".join(_fmt(v) for v in measured["solutions"]) or "없음")
    res["agent_directive"] = (
        ("VERIFIED: 찾은 답 근처를 다시 예측해 목표값을 지나는지 확인했습니다. " if not unverified else
         f"UNVERIFIED: 확인하지 못한 답이 있습니다 ({'; '.join(unverified)}). 다시 실행하지 말고 unverified를 그대로 알리세요. ")
        + ("답을 모두 알리세요(하나만 고르지 않음). " if answers else "찾은 값 없음(count 0): reachable(가능한 예측 범위)을 알리세요. ")
        + "operating_windows(목표가 90% 예측 오차 안에 드는 운전 구간), measured(학습 데이터 기준 답), "
        + "nearest_rows(가까운 실측 5행)를 함께 알리세요. 모델이 예측한 조건이지 보장된 조건이 아니라는 점을 덧붙이세요."
        + (" warnings도 반드시 알리세요." if warnings else ""))
    return res


# ==============================================================================
# 5-C. What-if analysis: --sweep (one input over a range) and --maximize / --minimize (constrained search)
# ==============================================================================
#
# Both use the regression models of the version folder, the fixed inputs of -i and the training medians for the
# numeric inputs not given (reported, like --find). The search stays inside the training range of the inputs it
# varies (no extrapolation: tree models are flat outside it); the sweep may leave it, and every such point is marked.

SWEEP_DEFAULT_POINTS = 21       # --sweep "<input>" or "<input>=a:b" without a step
SWEEP_MAX_POINTS = 500          # more points -> STOP_RETRY with [추천] a larger step
SWEEP_SHOWN = 25                # points in the JSON (evenly picked, ends kept); the table holds all
SWEEP_RANGE_TOL = 0.005         # a point this share of the training span outside it still counts as inside
SEARCH_MAX_VARY = 6
SEARCH_GRID_BUDGET = 20000      # model evaluations of the grid stage
SEARCH_REFINE_ROUNDS = 3        # coordinate search rounds after the grid (more than 3 inputs)
SEARCH_LINE_POINTS = 41
SEARCH_SOLUTIONS = 5
SEARCH_DISTINCT = 0.1           # solutions closer than this (inputs scaled to the training range) are one answer
TREND_FLAT = 0.01               # steps smaller than this share of the predicted span count as flat


def _regression_items(args: argparse.Namespace, names: Optional[List[str]] = None) -> Tuple[List[Dict[str, Any]], List[str]]:
    """The regression target models of -m (filtered by names when given) and the classification targets left out."""
    items = resolve_target_models(target_path=args.model, filter_targets=names)
    reg = [it for it in items if it["meta"].get("task", "regression") == "regression"]
    skipped = [str(it["target"]) for it in items if it not in reg]
    if not reg:
        raise InputError(f"조건별 예측(--sweep)과 조건 찾기(--maximize/--minimize)는 숫자 타겟(회귀) 모델만 씁니다. "
                         f"이 폴더의 타겟 {skipped}은(는) 분류 모델입니다.")
    return reg, skipped


def _model_input_info(items: List[Dict[str, Any]]) -> Tuple[List[str], List[str], Dict[str, List[Any]], Dict[str, Dict[str, Any]]]:
    """(raw inputs, numeric raw inputs, categorical inputs -> known values, numeric stats) over the models."""
    raw: List[str] = []
    cats: Dict[str, List[Any]] = {}
    dates: set = set()
    stats: Dict[str, Dict[str, Any]] = {}
    for it in items:
        meta = it["meta"]
        schema = meta.get("feature_schema") or {}
        raw += model_raw_inputs(meta)
        for c, known in (schema.get("categorical") or {}).items():
            cats.setdefault(c, list(known))
        dates |= set(schema.get("datetime") or {})
        for c, st in (meta.get("feature_stats") or {}).items():
            stats.setdefault(c, st)
    raw = list(dict.fromkeys(raw))
    numeric = [c for c in raw if c not in cats and c not in dates]
    return raw, numeric, cats, stats


def _whatif_command(args: argparse.Namespace, inline: str) -> Optional[str]:
    cmd = command_text(drop={"-i": "input", "--input": "input"})
    return f"{cmd} -i {quote_arg(inline)}" if cmd else None


def _base_inputs(args: argparse.Namespace, items: List[Dict[str, Any]], vary: List[str]) -> Dict[str, Any]:
    """The fixed inputs (-i, one record) and the training medians of the numeric inputs not given (not the varied ones).
    A missing text / date input stops (STOP_RETRY [후보]): no value of it is neutral."""
    raw, numeric, cats, stats = _model_input_info(items)
    given_df = parse_inline_input(args.input) if getattr(args, "input", None) else None
    if given_df is not None and len(given_df) != 1:
        raise InputError(f"-i은(는) 한 줄('열=값,열=값')로 주십시오 (지금 {len(given_df)}줄). 고정 입력이 여러 경우면 경우마다 따로 실행하십시오.")
    given = given_df.iloc[0].to_dict() if given_df is not None else {}
    names, _ = (resolve_column_names([str(k) for k in given], raw, "고정 입력(-i)", values=True, exclude=vary)
                if given else ([], {}))
    check_same_column([str(k) for k in given], names, "-i", {str(k): v for k, v in given.items()})
    fixed = {n: given[k] for k, n in zip(given, names)}
    resolved = dict(zip([str(k) for k in given], names))
    if given_df is not None:
        check_comma_numbers(given_df.attrs.get("kv_pairs"), lambda k: resolved.get(k) in numeric, "-i",
                            lambda fx: _whatif_command(args, kv_text(fx)))
    both = [v for v in vary if v in fixed]
    if both:
        raise InputError(f"{both}은(는) 바꿔 보는 입력이라 -i에 함께 줄 수 없습니다. -i에서 빼십시오.")
    warnings: List[str] = []
    for f in list(fixed):
        if f in cats:
            fixed[f] = fixed_category_value(f, fixed[f], cats[f])
        if f not in numeric:
            continue
        try:
            fixed[f] = float(fixed[f])
        except (TypeError, ValueError):
            raise InputError(f"고정 입력 '{f}'의 값 '{fixed[f]}'이(가) 숫자가 아닙니다. -i에 '{f}=<숫자>'로 다시 지정하십시오.")
        st = stats.get(f)
        if st and (fixed[f] < st["min"] or fixed[f] > st["max"]):
            warnings.append(f"고정 입력 '{f}'={_fmt(fixed[f])}이(가) 학습 범위({_fmt(st['min'])}~{_fmt(st['max'])}) 밖이라 "
                            f"예측을 믿기 어렵습니다 (외삽)")
    filled: Dict[str, float] = {}
    for f in raw:
        if f in vary or f in fixed:
            continue
        if f in numeric and f in stats:
            filled[f] = float(stats[f]["median"])
        elif f in cats:
            raise InputError(f"모델 입력 '{f}' 값이 필요합니다. -i에 '{f}=<값>'을 넣어 주십시오. "
                             f"{category_candidates(cats[f])}")
        else:
            raise InputError(f"모델 입력 '{f}' 값이 필요합니다. -i에 '{f}=<값>'을 넣어 주십시오.")
    if filled:
        warnings.insert(0, "주지 않은 입력은 학습 데이터 중앙값으로 고정했습니다: "
                        + ", ".join(f"{k}={_fmt(v)}" for k, v in filled.items()))
    return {"base": {**fixed, **filled}, "fixed": fixed, "filled": filled, "warnings": warnings,
            "raw": raw, "numeric": numeric, "stats": stats}


def _predict_rows(item: Dict[str, Any], rows: pd.DataFrame) -> np.ndarray:
    meta = item["meta"]
    df = apply_inference_evals(rows.copy(), meta.get("derived_evals", []))
    return np.asarray(item["pipeline"].predict(prepare_features(df, meta["feature_schema"])), dtype=float)


def _vary_names(args: argparse.Namespace, text: str, numeric: List[str], stats: Dict[str, Any], option: str) -> List[str]:
    given = [v.strip() for v in str(text).split(",") if v.strip()]
    if not given:
        raise InputError(f"{option}에 바꿔 볼 숫자 입력을 주십시오. [후보] {', '.join(numeric[:10])}")
    names = resolve_column_names(given, numeric, f"바꿔 볼 입력({option})")[0]
    bad = [n for n in names if n not in stats]
    if bad:
        raise InputError(f"{bad}의 학습 범위가 모델 정보에 없어 바꿔 볼 수 없습니다. [후보] {', '.join([c for c in numeric if c in stats][:10])}")
    if len(set(names)) != len(names):
        raise InputError(f"{option}에 같은 입력이 두 번 있습니다: {given}")
    return names


def _sweep_values(args: argparse.Namespace, name: str, given: str, spec: Optional[str], lo: float, hi: float) -> np.ndarray:
    """'a:b:step' (both ends included), 'a:b' (SWEEP_DEFAULT_POINTS points), nothing (the training range)."""
    if spec is None or not spec.strip():
        return np.linspace(lo, hi, SWEEP_DEFAULT_POINTS)
    parts = [p.strip() for p in spec.split(":")]
    try:
        nums = [float(p) for p in parts]
    except ValueError:
        raise InputError(f"--sweep '{args.sweep}'을(를) 읽을 수 없습니다. 형식: '<입력>=<시작>:<끝>:<간격>' "
                         f"(예: {strip_unit(name)}={_fmt(lo)}:{_fmt(hi)}:{_fmt((hi - lo) / 10 or 1)}), 숫자에 쉼표 없이")
    if len(nums) not in (2, 3) or nums[0] == nums[1] or (len(nums) == 3 and nums[2] <= 0):
        raise InputError(f"--sweep '{args.sweep}': 시작과 끝이 다르고 간격은 0보다 커야 합니다. 형식: '<입력>=<시작>:<끝>:<간격>'")
    a, b = nums[0], nums[1]
    if len(nums) == 2:
        return np.linspace(a, b, SWEEP_DEFAULT_POINTS)
    step = nums[2] if b > a else -nums[2]
    n = int(np.floor((b - a) / step + 1e-9)) + 1
    if n > SWEEP_MAX_POINTS:
        new_step = abs(b - a) / (SWEEP_MAX_POINTS - 1)
        nice = float(f"{new_step:.2g}")
        while nice < new_step:          # rounded up: the new step never makes too many points
            nice = float(f"{nice * 1.1:.2g}")
        new_step = nice
        cmd = command_text(drop={"--sweep": "sweep"})
        fixed = f"{given}={_fmt(a)}:{_fmt(b)}:{_fmt(new_step)}"
        rec = f"{cmd} --sweep {quote_arg(fixed)}" if cmd else fixed
        raise InputError(f"--sweep의 점이 {n}개로 많습니다 (최대 {SWEEP_MAX_POINTS}개). 간격을 넓히십시오. [추천] {rec}")
    xs = a + step * np.arange(n)
    if abs(xs[-1] - b) > 1e-9 * max(1.0, abs(b)):
        xs = np.append(xs, b)       # the end is always a point
    return xs


def _trend(xs: np.ndarray, ys: np.ndarray, vary: str, target: str) -> Dict[str, Any]:
    span = float(np.ptp(ys)) if len(ys) else 0.0
    imax, imin = int(np.argmax(ys)), int(np.argmin(ys))
    out: Dict[str, Any] = {"max": {vary: _r6(xs[imax]), "value": _r6(ys[imax])},
                           "min": {vary: _r6(xs[imin]), "value": _r6(ys[imin])},
                           "start": _r6(ys[0]), "end": _r6(ys[-1])}
    tol = TREND_FLAT * span
    d = np.diff(ys)
    up, down = int((d > tol).sum()), int((d < -tol).sum())
    v, t = strip_unit(vary), strip_unit(target)
    if span <= 1e-12 * max(1.0, float(np.max(np.abs(ys)))):
        out["direction"], text = "flat", f"{v}를 바꿔도 {t} 예측이 거의 변하지 않음"
    elif down == 0:
        out["direction"], text = "increasing", f"{v}가 커질수록 {t}이(가) 커짐 ({_fmt(ys[0])} -> {_fmt(ys[-1])})"
    elif up == 0:
        out["direction"], text = "decreasing", f"{v}가 커질수록 {t}이(가) 작아짐 ({_fmt(ys[0])} -> {_fmt(ys[-1])})"
    elif 0 < imax < len(ys) - 1 and ys[imax] - max(ys[0], ys[-1]) > 5 * tol:
        out["direction"], text = "peak", f"{v} {_fmt(xs[imax])}에서 {t}이(가) 가장 큼 ({_fmt(ys[imax])})"
    elif 0 < imin < len(ys) - 1 and min(ys[0], ys[-1]) - ys[imin] > 5 * tol:
        out["direction"], text = "valley", f"{v} {_fmt(xs[imin])}에서 {t}이(가) 가장 작음 ({_fmt(ys[imin])})"
    else:
        out["direction"], text = "mixed", (f"{v}에 따라 {t}이(가) 오르내림 (최대 {_fmt(ys[imax])} @ {_fmt(xs[imax])}, "
                                           f"최소 {_fmt(ys[imin])} @ {_fmt(xs[imin])})")
    out["text"] = text
    return out


def _sweep_chart(path: str, xs: np.ndarray, vary: str, curves: Dict[str, Dict[str, Any]], lo: float, hi: float) -> Optional[str]:
    """Line chart PNG (one panel per target, 90% band, training range shaded). None when matplotlib is missing."""
    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
        from matplotlib import font_manager
    except ImportError:
        return None
    have = {f.name for f in font_manager.fontManager.ttflist}
    font = next((f for f in ("Malgun Gothic", "NanumGothic", "AppleGothic", "Noto Sans CJK KR") if f in have), None)
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        rc = {"axes.unicode_minus": False}
        if font:
            rc["font.family"] = font
        with matplotlib.rc_context(rc):
            n = len(curves)
            fig, axes = plt.subplots(n, 1, figsize=(6.4, 2.6 * n + 0.4), squeeze=False)
            for ax, (t, c) in zip(axes[:, 0], curves.items()):
                if c.get("lo") is not None:
                    ax.fill_between(xs, c["lo"], c["hi"], color="#93C5FD", alpha=0.45, linewidth=0, label="90% 범위")
                ax.plot(xs, c["y"], color="#1E3A8A", linewidth=1.8, marker="o", markersize=2.5, label="예측")
                ax.axvspan(lo, hi, color="#E5E7EB", alpha=0.5, zorder=0, label="학습 범위")
                ax.set_ylabel(t)
                ax.grid(True, linewidth=0.4, alpha=0.6)
            axes[-1, 0].set_xlabel(vary)
            axes[0, 0].legend(loc="best", fontsize=8)
            fig.tight_layout()
            fig.savefig(path, dpi=120)
            plt.close(fig)
    return path


def _version_dir_of(items: List[Dict[str, Any]]) -> str:
    return os.path.dirname(os.path.abspath(items[0]["joblib_path"]))


def _save_table(args: argparse.Namespace, df: pd.DataFrame, default_name: str, what: str) -> Tuple[str, List[str]]:
    """Writes the result table (--save .xlsx/.csv, else <data folder>/<default_name>.xlsx) and reads it back (R1)."""
    if getattr(args, "save", None):
        path = os.path.abspath(args.save)
        if os.path.splitext(path)[1].lower() not in (".xlsx", ".csv"):
            raise InputError(f"결과 파일(--save)은 .xlsx 또는 .csv만 저장합니다: '{os.path.basename(path)}'. "
                             f"[추천] {option_retry(os.path.splitext(args.save)[0] + '.xlsx', '--save')}")
    else:
        path = os.path.join(os.path.dirname(_version_dir_of_args(args)), default_name + ".xlsx")
    os.makedirs(os.path.dirname(path), exist_ok=True)
    path = avoid_overwrite(path, getattr(args, "overwrite", False), what=what)
    if path.lower().endswith(".csv"):
        path = write_csv_excel_compatible(df, path)
    else:
        path = save_with_lock_fallback(path, lambda p: df.to_excel(p, index=False), what=what)
    mark_committed(f"{what} {os.path.basename(path)}", output_file=path)
    unverified: List[str] = []
    try:
        back = load_data_file(path, header_row=1)
        if len(back) != len(df):
            unverified.append(f"저장한 표를 다시 읽으니 {len(back)}행 (기대 {len(df)}행)")
    except Exception as e:
        unverified.append(f"저장한 표를 다시 읽지 못함: {type(e).__name__}: {e}")
    return path, unverified


def _version_dir_of_args(args: argparse.Namespace) -> str:
    p = os.path.abspath(args.model)
    return p if os.path.isdir(p) else os.path.dirname(p)


def _pick_evenly(n: int, k: int) -> List[int]:
    if n <= k:
        return list(range(n))
    return sorted(set(int(round(i)) for i in np.linspace(0, n - 1, k)))


def execute_sweep(args: argparse.Namespace) -> Dict[str, Any]:
    raw_spec = str(args.sweep)
    m = re.fullmatch(r"\s*([^=]+?)\s*(?:=(.*))?", raw_spec, re.S)
    if not m or not m.group(1).strip():
        raise InputError(f"--sweep '{raw_spec}'을(를) 읽을 수 없습니다. 형식: '<입력>=<시작>:<끝>:<간격>' (예: 반응온도=170:200:5)")
    filt = [t.strip() for t in args.targets.split(",") if t.strip()] if args.targets else None
    items, skipped = _regression_items(args, filt)
    raw, numeric, _, stats = _model_input_info(items)
    vary = _vary_names(args, m.group(1), numeric, stats, "--sweep")
    if len(vary) != 1:
        raise InputError("--sweep에는 입력 하나만 주십시오 (예: --sweep \"반응온도=170:200:5\"). 여러 입력을 함께 바꾸는 최적 조건 찾기는 --maximize/--minimize입니다.")
    vary_name = vary[0]
    info = _base_inputs(args, items, vary)
    st = stats[vary_name]
    lo, hi = float(st["min"]), float(st["max"])
    xs = _sweep_values(args, vary_name, m.group(1).strip(), m.group(2), lo, hi)
    rows = pd.DataFrame({**{k: [v] * len(xs) for k, v in info["base"].items()}, vary_name: xs})
    tol = SWEEP_RANGE_TOL * (hi - lo)       # 170 for a training minimum of 170.03 is not an extrapolation
    in_range = (xs >= lo - tol - 1e-12) & (xs <= hi + tol + 1e-12)
    table = pd.DataFrame({vary_name: [_r6(x) for x in xs]})
    curves: Dict[str, Dict[str, Any]] = {}
    trend: Dict[str, Any] = {}
    for it in items:
        t = str(it["target"])
        ys = _predict_rows(it, rows)
        iv = it["meta"].get("prediction_interval")
        c: Dict[str, Any] = {"y": ys, "lo": None, "hi": None}
        table[f"pred_{t}"] = [round_pred(float(v)) for v in ys]
        if iv:
            c["lo"], c["hi"] = ys + float(iv["residual_lower"]), ys + float(iv["residual_upper"])
            table[f"pred_lo_{t}"] = [round_pred(float(v)) for v in c["lo"]]
            table[f"pred_hi_{t}"] = [round_pred(float(v)) for v in c["hi"]]
        curves[t] = c
        trend[t] = _trend(xs, ys, vary_name, t)
    table["학습범위"] = np.where(in_range, "안", "밖(외삽)")
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out, unverified = _save_table(args, table, f"sweep_{sanitize_filename(strip_unit(vary_name))}_{stamp}", "조건별 예측 표")
    chart = None
    try:
        chart = _sweep_chart(os.path.splitext(out)[0] + ".png", xs, vary_name, curves, lo, hi)
        if chart and not (os.path.isfile(chart) and os.path.getsize(chart) > 0):
            unverified.append("그래프 PNG를 저장했지만 다시 확인하지 못함")
    except Exception as e:                 # the table is the result; the chart is an extra
        unverified.append(f"그래프를 그리지 못함: {type(e).__name__}: {str(e)[:120]}")
        chart = None
    warnings_out = list(info["warnings"])
    n_out = int((~in_range).sum())
    if n_out:
        warnings_out.append(f"{vary_name}의 {n_out}점이 학습 범위({_fmt(lo)}~{_fmt(hi)}) 밖입니다 (표의 '학습범위' 열 '밖(외삽)', "
                            f"points의 in_range false): 그 점의 예측은 참고용입니다. 트리 모델은 범위 밖에서 값이 거의 변하지 않습니다")
    if chart is None and not any("그래프" in u for u in unverified):
        warnings_out.append("matplotlib이 없어 그래프는 만들지 않았습니다 (표는 정상). 설치: python -m pip install matplotlib")
    if skipped:
        warnings_out.append(f"분류 타겟 {skipped}은(는) 조건별 예측에서 뺐습니다 (숫자 타겟만)")
    pick = _pick_evenly(len(xs), SWEEP_SHOWN)
    points = []
    for i in pick:
        p: Dict[str, Any] = {vary_name: _r6(xs[i]), "in_range": bool(in_range[i])}
        for t, c in curves.items():
            p[t] = round_pred(float(c["y"][i]))
            if c["lo"] is not None:
                p[f"{t}_90%"] = [round_pred(float(c["lo"][i])), round_pred(float(c["hi"][i]))]
        points.append(p)
    truncated = len(pick) < len(xs)
    res: Dict[str, Any] = {
        "status": "ok", "command": "sweep", "vary": vary_name, "targets": list(curves),
        "fixed": {k: (float(v) if isinstance(v, (int, float, np.floating)) else v) for k, v in info["fixed"].items()},
        "filled_defaults": {k: _r6(v) for k, v in info["filled"].items()},
        "training_range": {vary_name: [_r6(lo), _r6(hi)]},
        "count": len(xs), "shown": len(points), "truncated": truncated, "points": points,
        "out_of_range_points": n_out, "trend": trend, "output_file": out, "verified": not unverified,
    }
    if chart:
        res["chart_file"] = chart
    if truncated:
        res["full_result_file"] = out
    if warnings_out:
        res["warnings"] = warnings_out
    if unverified:
        res["unverified"] = unverified
    res["preview"] = (f"{vary_name} {_fmt(xs[0])}~{_fmt(xs[-1])} ({len(xs)}점) 조건별 예측 | "
                      + " | ".join(tr["text"] for tr in trend.values()) + f" | 표: {os.path.basename(out)}")
    res["agent_directive"] = (
        ("VERIFIED: 표를 다시 읽어 확인했습니다. " if not unverified else
         f"UNVERIFIED: 확인하지 못한 것이 있습니다 ({'; '.join(unverified)}). 다시 실행하지 말고 unverified를 그대로 알리세요. ")
        + "trend의 타겟별 text(경향)와 points(조건별 예측값" + (", 90% 범위" if any(c["lo"] is not None for c in curves.values()) else "")
        + "), output_file(엑셀 표)" + (", chart_file(그래프)" if chart else "") + " 경로를 알리세요. "
        + (f"points는 전체 {len(xs)}점 중 고르게 고른 {len(points)}점이고 전체는 output_file에 있습니다. " if truncated else "")
        + ("in_range가 false인 점은 학습 범위 밖(외삽)이라 참고용이라고 알리세요. " if n_out else "")
        + ("warnings도 반드시 알리세요. " if warnings_out else "")
        + "모델 예측이지 실측이 아니라는 점을 덧붙이세요.")
    return res


SEARCH_RECHECK_TOL = 1e-6      # relative difference allowed between a reported and a re-predicted value


def _recheck_solutions(items: List[Dict[str, Any]], base: Dict[str, Any], vary: List[str],
                       solutions: List[Dict[str, Any]], conds: Optional[List[List[Cond]]], goal: str) -> List[str]:
    """R1 check of --maximize / --minimize: every answer is predicted again from the inputs as reported (rounded) and
    the fixed inputs, in one new call per model; the reported predictions must match and the --where condition must
    hold. Returns what did not (empty = verified)."""
    if not solutions:
        return []
    rows = pd.DataFrame({**{k: [v] * len(solutions) for k, v in base.items()},
                         **{v: [s["inputs"][v] for s in solutions] for v in vary}})
    frame = rows.copy()
    for it in items:
        frame[str(it["target"])] = _predict_rows(it, rows)
    ok = where_mask(frame, conds) if conds else np.ones(len(solutions), dtype=bool)
    bad: List[str] = []
    for i, s in enumerate(solutions):
        diff = [t for t, v in s["predictions"].items()
                if abs(float(frame[t].iloc[i]) - float(v)) > SEARCH_RECHECK_TOL * max(1.0, abs(float(v))) + 1e-4]
        if diff:
            bad.append(f"답 {i + 1}: 다시 예측한 {', '.join(diff)} 값이 다름")
        elif not ok[i]:
            bad.append(f"답 {i + 1}: 적힌 입력값으로 다시 예측하면 조건을 만족하지 않음")
    return bad


def execute_search(args: argparse.Namespace) -> Dict[str, Any]:
    goal_text = args.maximize if args.maximize else args.minimize
    sense = 1.0 if args.maximize else -1.0
    option = "--maximize" if args.maximize else "--minimize"
    items, skipped = _regression_items(args)
    names = [str(it["target"]) for it in items]
    goal = resolve_column_names([str(goal_text).strip()], names, f"목표({option})")[0][0]
    raw, numeric, _, stats = _model_input_info(items)
    if not args.vary:
        raise InputError(f"{option}에는 바꿔 볼 입력 --vary \"<입력1>,<입력2>\"가 필요합니다. [후보] {', '.join(numeric[:10])}")
    vary = _vary_names(args, args.vary, numeric, stats, "--vary")
    if len(vary) > SEARCH_MAX_VARY:
        raise InputError(f"--vary 입력은 {SEARCH_MAX_VARY}개까지입니다 (지금 {len(vary)}개). 영향이 큰 입력만 고르도록 사용자에게 물으십시오.")
    info = _base_inputs(args, items, vary)
    conds = None
    if args.where:
        conds = resolve_where(parse_where(args.where, "--where"), names + raw, "--where")
    lows = np.array([float(stats[v]["min"]) for v in vary])
    highs = np.array([float(stats[v]["max"]) for v in vary])
    spans = np.maximum(highs - lows, 1e-12)
    evals = {"n": 0}

    def evaluate(points: np.ndarray) -> Tuple[pd.DataFrame, np.ndarray, np.ndarray]:
        rows = pd.DataFrame({**{k: [v] * len(points) for k, v in info["base"].items()},
                             **{v: points[:, j] for j, v in enumerate(vary)}})
        frame = rows.copy()
        for it in items:
            frame[str(it["target"])] = _predict_rows(it, rows)
        evals["n"] += len(points)
        ok = where_mask(frame, conds) if conds else np.ones(len(points), dtype=bool)
        return frame, ok, sense * frame[goal].to_numpy(dtype=float)

    k = len(vary)
    per = max(int(np.floor(SEARCH_GRID_BUDGET ** (1.0 / k))), 3)
    per = min(per, 2001)
    axes = [np.linspace(lows[j], highs[j], per) for j in range(k)]
    grid = np.array(np.meshgrid(*axes, indexing="ij")).reshape(k, -1).T
    frame, ok, score = evaluate(grid)
    frames = [frame.assign(_ok=ok, _score=score)]
    method = f"격자 탐색 (입력마다 {per}점, {len(grid)}개 조합)"
    if k > 1 and ok.any():
        # coordinate search from the best feasible grid point: one input at a time over its whole training range
        best = grid[int(np.argmax(np.where(ok, score, -np.inf)))].copy()
        for _ in range(SEARCH_REFINE_ROUNDS):
            moved = False
            for j in range(k):
                line = np.repeat(best[None, :], SEARCH_LINE_POINTS, axis=0)
                line[:, j] = np.linspace(lows[j], highs[j], SEARCH_LINE_POINTS)
                f2, ok2, s2 = evaluate(line)
                frames.append(f2.assign(_ok=ok2, _score=s2))
                if ok2.any():
                    i = int(np.argmax(np.where(ok2, s2, -np.inf)))
                    cur = np.where(ok2, s2, -np.inf)[int(np.argmin(np.abs(line[:, j] - best[j])))]
                    if s2[i] > cur + 1e-12 and not np.isclose(line[i, j], best[j]):
                        best = line[i].copy()
                        moved = True
            if not moved:
                break
        method += f" + 좌표 탐색 ({SEARCH_REFINE_ROUNDS}회까지, 입력마다 {SEARCH_LINE_POINTS}점)"
    allf = pd.concat(frames, ignore_index=True)
    feas = allf[allf["_ok"]].sort_values("_score", ascending=False, kind="stable")
    chosen: List[int] = []
    pts = feas[vary].to_numpy(dtype=float)
    for i in range(len(feas)):
        if all(np.max(np.abs(pts[i] - pts[c]) / spans) > SEARCH_DISTINCT for c in chosen):
            chosen.append(i)
        if len(chosen) >= SEARCH_SOLUTIONS:
            break
    goal_item = next(it for it in items if str(it["target"]) == goal)
    iv = goal_item["meta"].get("prediction_interval")
    solutions = []
    for c in chosen:
        r = feas.iloc[c]
        sol: Dict[str, Any] = {"inputs": {v: _r6(r[v]) for v in vary},
                               "predictions": {t: round_pred(float(r[t])) for t in names}}
        if iv:
            sol["interval_90"] = {goal: [round_pred(float(r[goal]) + float(iv["residual_lower"])),
                                         round_pred(float(r[goal]) + float(iv["residual_upper"]))]}
        solutions.append(sol)
    unverified = _recheck_solutions(items, info["base"], vary, solutions, conds, goal)
    warnings_out = list(info["warnings"])
    if skipped:
        warnings_out.append(f"분류 타겟 {skipped}은(는) 조건 찾기에서 뺐습니다 (숫자 타겟만)")
    best_free = allf.iloc[int(np.argmax(allf["_score"].to_numpy()))]
    near: List[Dict[str, Any]] = []
    data = _load_snapshot(_version_dir_of(items), {})
    if data is not None and solutions and all(c in data.columns for c in [SNAPSHOT_ROW_COL] + vary):
        data = data.copy()
        data["__row"] = coerce_numeric(data[SNAPSHOT_ROW_COL]).fillna(-1).astype(int)
        keys = vary + [c for c in info["numeric"] if c not in vary and c in data.columns and c in info["base"]]
        target_pt = np.array([solutions[0]["inputs"][v] for v in vary] + [float(info["base"][c]) for c in keys[len(vary):]])
        span_k = np.array([max(float(stats[c]["max"]) - float(stats[c]["min"]), 1e-12) for c in keys])
        show = list(dict.fromkeys(keys + [t for t in names if t in data.columns]))
        near = nearest_rows(data, keys, target_pt, span_k, show)
    count = len(solutions)
    where_txt = where_text(conds) if conds else ""
    word = "최대" if sense > 0 else "최소"
    res: Dict[str, Any] = {
        "status": "ok", "command": "maximize" if sense > 0 else "minimize", "goal": goal, "vary": vary,
        "where": where_txt or None,
        "fixed": {k2: (float(v) if isinstance(v, (int, float, np.floating)) else v) for k2, v in info["fixed"].items()},
        "filled_defaults": {k2: _r6(v) for k2, v in info["filled"].items()},
        "training_range": {v: [_r6(lows[j]), _r6(highs[j])] for j, v in enumerate(vary)},
        "method": method, "evaluations": evals["n"], "feasible_count": int(len(feas)),
        "count": count, "solutions": solutions, "nearest_rows": near, "verified": not unverified,
    }
    if unverified:
        res["unverified"] = unverified
    if not count:
        res["best_without_where"] = {"inputs": {v: _r6(best_free[v]) for v in vary},
                                     "predictions": {t: round_pred(float(best_free[t])) for t in names}}
    if warnings_out:
        res["warnings"] = warnings_out
    cond_txt = f" ({where_txt})" if where_txt else ""
    if count:
        s0 = solutions[0]
        res["preview"] = (f"{goal} {word} 조건{cond_txt}: " + ", ".join(f"{v}={_fmt(x)}" for v, x in s0["inputs"].items())
                          + f" -> {goal} {_fmt(s0['predictions'][goal])} | 답 {count}개 | 학습 범위 안에서만 찾음")
        res["agent_directive"] = (
            ("VERIFIED: 답마다 적힌 입력값(inputs)으로 모델을 다시 예측해 predictions와 조건을 확인했습니다. " if not unverified else
             f"UNVERIFIED: 다시 예측해 확인하지 못한 답이 있습니다 ({'; '.join(unverified)}). 다시 실행하지 말고 unverified를 그대로 알리세요. ")
            + "solutions(답 모두, 첫 번째가 가장 좋음)의 inputs와 predictions, "
            "nearest_rows(가까운 실측 행)를 알리세요. 학습 범위 안에서 모델이 예측한 조건이지 보장된 조건이 아니라는 점을 덧붙이세요."
            + (" warnings도 반드시 알리세요." if warnings_out else ""))
    else:
        res["preview"] = (f"{goal} {word} 조건{cond_txt}: 0건 (학습 범위 안에서 조건을 만족하는 조합 없음) | "
                          f"조건 없이 {word}: {_fmt(best_free[goal])}")
        res["agent_directive"] = (
            f"VERIFIED: 0건. 학습 범위 안에서 조건({where_txt})을 만족하는 조합을 찾지 못했습니다. 사용자에게 그대로 알리고 "
            "best_without_where(조건 없이 가장 좋은 조합)를 함께 알리세요. 조건을 바꿔 다시 실행하지 마세요."
            + (" warnings도 반드시 알리세요." if warnings_out else ""))
    return res


# ==============================================================================
# 6. CLI Entry Point
# ==============================================================================

def main():
    setup_console()
    check_core_libraries("예측")
    parser = JsonArgumentParser(
        description="학습한 모델로 예측 (한 건, 파일, 반대 찾기, 조건별 예측, 조건 찾기)"
    )
    parser.add_argument(
        "--model", "-m", required=True,
        help="학습한 데이터 파일(도구가 최고 버전을 고름), 버전 폴더(예: sample_ml_data/v1) 또는 .joblib 파일"
    )
    parser.add_argument(
        "--targets", "-y",
        help="예측할 특정 타겟 목록 (쉼표 구분, 폴더 모드에서 생략 시 폴더 내 모든 타겟 모델 자동 추론)"
    )
    parser.add_argument(
        "--file", "-f",
        help="배치 추론 대상 데이터 파일 경로 (.xlsx, .csv)"
    )
    parser.add_argument(
        "--input", "-i",
        help="단일/인라인 추론용 입력 데이터 ('k1=v1,k2=v2' 형식, JSON 문자열, 또는 .json 파일 경로)"
    )
    parser.add_argument(
        "--save",
        help="예측 결과 저장 파일 경로 (.csv, .xlsx, 생략 시 원본 파일 위치에 pred_*.csv 자동 생성)"
    )
    parser.add_argument("--header-row", type=int, help="컬럼명 헤더 행 번호 (1-based, 미지정 시 상단 빈 행 자동 건너뜀)")
    parser.add_argument("--unit-row", type=int, help="단위 행 번호 (선택)")
    parser.add_argument("--data-row", type=int, help="데이터 시작 행 번호 (선택)")
    parser.add_argument("--overwrite", action="store_true",
                        help="--save 파일이 이미 있으면 덮어쓰기 (생략 시 시각을 붙인 새 이름으로 저장)")
    parser.add_argument("--find", help="모델 반대 찾기 목표 ('생산효율=50', 숫자 타겟). --vary와 함께")
    parser.add_argument("--vary", help="반대 찾기에서 값을 찾을 숫자 입력 (예: 반응온도). 나머지 입력은 -i로 고정, 생략하면 학습 중앙값. "
                                       "--maximize/--minimize에서는 바꿔 볼 입력 목록 (쉼표 구분, 6개까지)")
    parser.add_argument("--sweep", help="조건별 예측: 입력 하나를 바꿔 가며 예측 ('반응온도=170:200:5' = 시작:끝:간격, "
                                        "'반응온도' = 학습 범위). 나머지 입력은 -i, 생략하면 학습 중앙값. 표(엑셀)와 그래프 저장")
    parser.add_argument("--maximize", help="조건 찾기: 이 숫자 타겟이 가장 커지는 --vary 입력 조합 (학습 범위 안)")
    parser.add_argument("--minimize", help="조건 찾기: 이 숫자 타겟이 가장 작아지는 --vary 입력 조합 (학습 범위 안)")
    parser.add_argument("--where", help="조건 찾기의 제약 (예: '불순물 <= 100; 반응온도 < 190'. ;는 그리고, |는 또는)")

    args = parser.parse_args()

    search = args.maximize or args.minimize
    if args.where and not search:
        parser.error("--where는 --maximize 또는 --minimize와 함께 씁니다.")
    if args.maximize and args.minimize:
        parser.error("--maximize와 --minimize는 하나만 주십시오.")
    if args.sweep:
        if args.find or search or args.file:
            parser.error("조건별 예측(--sweep)에는 --find, --maximize, --minimize, -f를 함께 쓰지 않습니다.")
        if args.vary:
            parser.error("--sweep에는 바꿀 입력을 --sweep 값에 씁니다 (예: --sweep \"반응온도=170:200:5\"). --vary를 빼십시오.")
        run_cli(lambda: execute_sweep(resolve_model_arg(args)), hint="확인할 것: 입력 이름과 --sweep 범위.")
        return
    if search:
        if args.find or args.file or args.targets:
            parser.error("조건 찾기(--maximize/--minimize)에는 --find, -f, -y를 함께 쓰지 않습니다.")
        if not args.vary:
            parser.error("조건 찾기(--maximize/--minimize)에는 바꿔 볼 입력 --vary \"<입력1>,<입력2>\"가 필요합니다.")
        run_cli(lambda: execute_search(resolve_model_arg(args)), hint="확인할 것: 타겟, --vary, --where의 이름.")
        return
    if args.find or args.vary:
        if not (args.find and args.vary):
            parser.error("모델 반대 찾기는 --find와 --vary를 함께 주어야 합니다.")
        if args.file or args.targets or args.save:
            parser.error("반대 찾기(--find)에는 -f, -y, --save를 쓰지 않습니다.")
        run_cli(lambda: execute_find(resolve_model_arg(args)), hint="확인할 것: 타겟 이름과 입력 이름.")
        return
    if not args.file and not args.input:
        parser.error("추론을 수행하려면 --file (배치 파일) 또는 --input (인라인 데이터) 중 하나를 지정해야 합니다.")

    run_cli(lambda: execute_predict(resolve_model_arg(args)), hint="확인할 것: 모델 폴더 경로와 입력 이름.")


if __name__ == "__main__":
    main()
