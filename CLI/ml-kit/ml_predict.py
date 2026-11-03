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

import sys
import os
import re
import ast
import csv
import glob
import json
import time
import shutil
import difflib
import functools
import argparse
import warnings
from datetime import datetime
from itertools import product
from typing import Dict, Any, List, Optional, Tuple, Callable

import numpy as np
import pandas as pd
import joblib

warnings.filterwarnings("ignore", category=UserWarning)


# ##############################################################################
# PART C. SHARED RUNTIME (imported by ml_train.py / ml_report.py)
# ##############################################################################

# ==============================================================================
# C1. Console & Agent CLI Protocol (Clean Terminal Principle - No Emojis)
# ==============================================================================

VERSION = "2.9.33"                         # must equal SKILL.md metadata.version (tests check it)
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


def emit_json_response(payload: Dict[str, Any]) -> None:
    """Prints standard 1-line JSON response to stdout for LLM consumption."""
    setup_console()
    sys.stdout.write(dumps_json(payload) + "\n")
    sys.stdout.flush()


class InputError(ValueError):
    """User/agent input or data schema problem. Retrying with the same arguments will not help."""


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
            "agent_directive": "STOP_RETRY: Execution was interrupted. Ask the user before running again.",
        }
    if isinstance(exc, _STOP_RETRY_ERRORS):
        return {
            "status": "error",
            "error_type": "input",
            "message": msg,
            "agent_directive": (
                "STOP_RETRY: Argument or data schema validation failed. "
                "If the message contains '[추천]', re-run ONCE with the suggested name, option or command "
                "(a [추천] that starts with python is a full command: run it as given, once); "
                "otherwise report the message to the user." + (f" {hint}" if hint else "")
            ),
        }
    return {
        "status": "error",
        "error_type": "internal",
        "message": f"{type(exc).__name__}: {msg}",
        "agent_directive": "ERROR: Unexpected internal failure. Do not retry; report the message to the user.",
    }


_RECOVERED: List[str] = []


def note_recovered(msg: str) -> None:
    """Records a JIT recovery that changed the result (e.g. saved under a new name); reported as JSON `recovered`."""
    msg = str(msg)[:200]
    sys.stderr.write(f"[WARN] {msg}\n")
    _RECOVERED.append(msg)


def _with_recovered(payload: Dict[str, Any]) -> Dict[str, Any]:
    if _RECOVERED and "recovered" not in payload:
        payload = {**payload, "recovered": list(_RECOVERED)}
    return payload


_RECOVERY_ATTEMPTS: List[str] = []


def record_recovery_attempt(msg: str) -> None:
    """Records a JIT recovery that hit its limit (R4); written as the '복구 시도:' line of the error log (R6)."""
    _RECOVERY_ATTEMPTS.append(str(msg)[:200])


# ------------------------------------------------------------------------------
# Error log (R6): failed runs are appended to <data>/logs/errors.txt for later analysis
# by people and development agents. The path is never put in stdout JSON.
# ------------------------------------------------------------------------------

def get_data_dir() -> str:
    """Tool data folder: ML_KIT_DATA_DIR > <tool folder>/data (not created here)."""
    return os.path.abspath(os.environ.get("ML_KIT_DATA_DIR") or os.path.join(TOOL_DIR, "data"))


def error_log_path() -> str:
    return os.path.join(get_data_dir(), "logs", "errors.txt")


def quote_arg(value: Any) -> str:
    """Quotes one command-line argument the way cmd.exe expects (for logs and next_commands)."""
    s = str(value)
    if s and all(ch not in s for ch in ' \t"&|<>^()%,;='):
        return s
    return '"' + s.replace('"', '\\"') + '"'


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
    try:
        result = execute()
    except KeyboardInterrupt as e:
        sys.stderr.write("\n[ERROR] 실행 중단 (KeyboardInterrupt)\n")
        emit_json_response(_with_recovered(error_payload(e)))
        sys.exit(130)
    except Exception as e:
        import traceback
        trace = traceback.format_exc()
        sys.stderr.write(f"\n[ERROR] {e}\n")
        payload = _with_recovered(error_payload(e, hint))
        write_error_log(payload, trace if payload.get("error_type") == "internal" else "")
        emit_json_response(payload)
        sys.exit(1)
    emit_json_response(_with_recovered(result))


class JsonArgumentParser(argparse.ArgumentParser):
    """argparse that reports usage errors as Agent CLI Protocol JSON instead of bare text."""

    def print_help(self, file=None) -> None:  # type: ignore[override]
        """R7: help text goes to stderr, stdout gets exactly 1 JSON line (argparse then exits 0)."""
        super().print_help(sys.stderr)
        skill = os.path.join(os.path.dirname(os.path.abspath(sys.argv[0] or __file__)), "SKILL.md")
        emit_json_response({
            "status": "ok", "tool": "ml-kit", "command": "help", "verified": True, "skill_file": skill,
            "agent_directive": "VERIFIED: 도움말을 stderr에 출력했습니다. SKILL.md의 명령 템플릿을 그대로 사용하세요. "
                               "도구 코드(.py)는 읽지 마세요. 방법을 모르겠으면 사용자에게 물어보세요.",
        })

    def error(self, message: str) -> None:  # type: ignore[override]
        self.print_usage(sys.stderr)
        payload = {
            "status": "error",
            "error_type": "input",
            "message": f"CLI 인자 오류: {message}",
            "agent_directive": "STOP_RETRY: Invalid CLI arguments. Check the option names in SKILL.md and report to the user.",
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
            f"{' / '.join(hints)}.{ask}\n[사용 가능한 컬럼 목록]: {col_list}"
        )
    for short, full in applied.items():
        sys.stderr.write(f"[INFO] 컬럼명 자동 매칭: '{short}' -> '{full}'\n")
    return resolved, applied


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
    hits = [v.strip() for v in series.astype("object") if isinstance(v, str) and _PERCENT_RE.fullmatch(v.strip())]
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
                       f"({ex[0]} -> {first}, 100으로 나누지 않음). 열의 단위가 %인지 확인하십시오.")
    return out


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
    # 지우지 않음 -> 숫자로 못 읽은 값(결측 경고)이 됨 (예전에는 1756으로 조용히 바뀜)
    s = s.map(lambda v: _strip_thousands(_strip_percent(v.strip())) if isinstance(v, str) else v)
    s = s.replace(_NA_PLACEHOLDERS, np.nan)
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
# assignment of a shared name (2.9.19 _THOUSANDS_RE, unanchored for check_comma_numbers) silently changed this check.
_DECIMAL_COMMA_RE = re.compile(r"\s*[-+]?\d+,\d+\s*")                          # '175,6', '0,149', '27,8448'
_CSV_THOUSANDS_RE = re.compile(r"\s*[-+]?[1-9]\d{0,2}(,\d{3})+(\.\d+)?\s*")   # '1,500', '1,346.8' (never '0,149')
_EURO_NUMBER_RE = re.compile(r"\s*[-+]?[1-9]\d{0,2}(\.\d{3})+,\d+\s*")       # '1.346,8': thousands dot, decimal comma
_DOT_THOUSANDS_RE = re.compile(r"\s*[-+]?[1-9]\d{0,2}(\.\d{3})+\s*")          # '1.346': thousands dot or 3 decimals
_DOT_COMMA_NUMBER_RE = re.compile(r"\s*[-+]?[\d.,]*\d[\d.,]*\s*")              # digits with '.' and ',' only


def _decimal_commas(df: pd.DataFrame, sep: str, trained: bool = False) -> pd.DataFrame:
    """A CSV not separated by commas (';', tab) whose numbers use a decimal comma ('175,6', European Excel): a comma
    that cannot be a thousands separator ('175,6', '2,19', '0,149', '27,8448') means every '<digits>,<digits>' cell is
    a decimal number, read with a decimal point (df.attrs["decimal_comma"], layout_warnings). European numbers with a
    thousands dot ('1.346,8') are read as 1346.8 (with a warning) when every number cell with both '.' and ',' has that
    form. Refused (InputError, a value would change by 1000x either way): cells that could be thousands too ('1,500'),
    dot-grouped numbers without a comma ('1.300', '2.125': 1300 or 1.3), or European numbers mixed with other '.'/','
    forms in the same file. trained: the model (or the ml_map data file) was read with decimal commas, so this file is
    read the same way whatever its delimiter (a one-column file is sniffed as ','), and refused when it has no
    decimal-comma cell but a '1.300' or '1,500' (read differently from the training file). A comma CSV, or one
    without such cells, is unchanged (coerce_numeric strips thousands separators)."""
    if df.empty or (sep == "," and not trained):
        return df
    cells = pd.Series(df.to_numpy(dtype=object).ravel())
    vals = cells[cells.map(lambda v: isinstance(v, str) and "," in v)].astype(str)
    dec = vals[vals.str.fullmatch(_DECIMAL_COMMA_RE)]
    thousands = vals[vals.str.fullmatch(_CSV_THOUSANDS_RE)]
    only_dec = dec[~dec.str.fullmatch(_CSV_THOUSANDS_RE)]
    dotted = vals[vals.str.contains(".", regex=False) & vals.str.fullmatch(_DOT_COMMA_NUMBER_RE)]
    euro = dotted[dotted.str.fullmatch(_EURO_NUMBER_RE)]
    confirmed = not (only_dec.empty and euro.empty)
    if not confirmed and not trained:
        return df
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
    text = cells[cells.map(lambda v: isinstance(v, str) and "," not in v and "." in v)].astype(str)
    grouped = text[text.str.fullmatch(_DOT_THOUSANDS_RE)]
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

    out = df.apply(lambda c: c.map(fix))
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
        ok, reason = unit_row_verdict(df_raw, j, header_idx) if unit_cells and j > header_idx else (False, "")
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
    df_data.attrs["summary_rows"] = _SharedList(_summary_rows(df_data, source_rows))
    df_data.attrs["decimal_comma"] = df_raw.attrs.get("decimal_comma")
    return df_data


# A text cell that names a totals / average row (spaces removed, lower case). Totals words always mark a totals row;
# average words can also be values of a category ('집계' = 평균/최대) and are judged per column (_summary_rows).
TOTAL_WORDS = {"합계", "총계", "소계", "계", "총합", "total", "sum", "subtotal", "grandtotal"}
AVERAGE_WORDS = {"평균", "average", "avg", "mean"}
# Wording of the same rows in daily reports (2.9.25): '일계', '월계', '누계', '총 합계', '전체합계', '합계:', 'Total:',
# '소계(A)', '주간계', '평균값', 'AVG.'. Exact words: TOTAL_WORDS, AVERAGE_WORDS and the period totals below.
# Loose words ('...합계', short '...계' other than ordinary words, 'Total'/'Sum' as a separate word, '평균값'/'평균치')
# are also values of a category ('A계'/'B계', '생산계', 'pH계'): _summary_rows counts a loose totals word only when
# the row's numbers are totals (_loose_by_value, 2.9.27), a loose average word when its rows are few between the data
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


def _unit_cell(v: Any, below: List[Any]) -> str:
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
        return "unit" if _unit_text(t) else "note"
    x = _number(v)
    nums = [n for n in (_number(b) for b in below) if n is not None]
    if x is None or not nums or x == 0:
        return "data"              # a date, a number over a text column, a zero
    lo, hi = min(nums), max(nums)
    w = max(hi - lo, 0.5 * max(abs(lo), abs(hi)), 1e-9)
    return "unit" if (x > hi + 3 * w or x < lo - 3 * w) else "data"     # a scale unit: 1000 over 1..5


def unit_row_verdict(df_raw: pd.DataFrame, j: int, header: Optional[int] = None) -> Tuple[bool, str]:
    """(row j (0-based) is the unit row of the header row `header` (default j - 1), why not). Judged per cell, then
    by the row (_unit_cell):
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
        kinds[k] = _unit_cell(v, below)
    if any(c == "data" for c in kinds.values()):
        return False, "아래 데이터와 같은 값(그 열 범위 안의 숫자·날짜, 아래에도 있는 글자 값)이 든 데이터 행이라"
    units = [k for k, c in kinds.items() if c == "unit" and not _is_data_value(row[k])]
    text = [k for k, c in kinds.items() if c in ("unit", "note") and not _is_data_value(row[k])]
    if not units:
        return False, "단위로 볼 짧은 글자가 없어"
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
PATH_OPTS = ("-f", "--file", "-m", "--model", "--meta", "-o", "--out", "--save", "--points", "-d", "--data-dir")


def layout_retry(header_row: int, unit_row: Optional[int]) -> str:
    """The [추천] for a new header (and unit) row: the running command with --header-row/--unit-row replaced and
    --data-row dropped (several options change, so a full command, R8); option text when the running program is not
    an ml-kit command (in-process use)."""
    opts = f"--header-row {header_row}" + (f" --unit-row {unit_row}" if unit_row else "")
    argv = CURRENT_ARGV if CURRENT_ARGV is not None else list(sys.argv)
    script = os.path.basename(argv[0] or "") if argv else ""
    if not (script.startswith("ml_") and script.endswith(".py")):
        return opts
    out: List[str] = []
    i = 1
    while i < len(argv):
        a = str(argv[i])
        key, has_eq = a.split("=", 1)[0], "=" in a
        hit = [o for o in LAYOUT_OPTS if o == key or (key.startswith("--") and len(key) > 3 and o.startswith(key))]
        if len(hit) == 1:
            i += 1 if has_eq else 2
            continue
        if key in PATH_OPTS and not has_eq and i + 1 < len(argv):
            out += [a, quote_arg(os.path.abspath(str(argv[i + 1]).rstrip("\\/") or str(argv[i + 1])))]
            i += 2
            continue
        out.append(quote_arg(a))
        i += 1
    return " ".join([f"python {quote_arg(os.path.abspath(argv[0]))}"] + out + [opts])


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


def find_header_row(candidates: Optional[List[List[Any]]], names: List[str]) -> Optional[int]:
    """1-based row (df.attrs["header_candidates"] of load_data_file) that holds every name (a two-level / merged header)."""
    found = find_header_layout(candidates, names)
    return found[0] if found else None


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
        return f"{where}. [추천] --header-row {row}"
    return f"{where}. [추천] {layout_retry(row, unit)}"


# ==============================================================================
# C4-B. Feature Schema (numeric / categorical / datetime), identical at train and inference
# ==============================================================================
#
# feature_schema (stored in vX/model_<target>.json, models trained by 2.1.0+):
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


def schema_raw_columns(schema: Dict[str, Any]) -> List[str]:
    """User-facing input columns (before date expansion)."""
    return list(schema.get("numeric", [])) + list(schema.get("datetime", {})) + list(schema.get("categorical", {}))


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
}

_NUMBER_NAME_RE = re.compile(r"(\d+\.?\d*|\.\d+)([eE][-+]?\d+)?")    # a column name that reads as a number
_NAME_BOUNDARY_L = r"(?<![가-힣a-zA-Z0-9_])"
_NAME_BOUNDARY_R = r"(?![가-힣a-zA-Z0-9_])"


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
                + f". 컬럼명은 데이터의 전체 이름(단위 포함) 또는 단위를 뺀 고유 이름으로 쓰십시오. [사용 가능한 컬럼 목록]: {columns}"
            )


def compile_formula(formula: str, columns: List[str]) -> Tuple[str, Dict[str, str], str]:
    """
    Replaces column names in a formula with internal tokens (prevents collisions with
    function names and lets names contain '(', '%', '/', etc.).
    Full names are matched first (longest first), then unique unit-less aliases.
    Returns (tokenized_formula, token_to_column, canonical_formula_with_full_names).
    """
    col_list = [str(c) for c in columns]
    # a column named like a number ('1', '2'): the number in the formula ('*2', also '2.5') would silently read that
    # column (the same boundaries as the name replacement below)
    for c in col_list:
        if _NUMBER_NAME_RE.fullmatch(c.strip()) and \
                re.search(_NAME_BOUNDARY_L + re.escape(c) + _NAME_BOUNDARY_R, formula):
            raise InputError(
                f"데이터에 숫자 이름의 열 '{c}'이(가) 있어 수식 '{formula}'의 숫자 {c.strip()}이(가) 숫자인지 그 열인지 알 수 "
                f"없습니다 (수식은 열 이름을 따옴표나 괄호로 감싸 쓰는 방법이 없음). 사용자에게 알려 데이터의 열 '{c}' 이름을 "
                f"글자가 든 이름(예: '열{c.strip()}')으로 바꾼 뒤 다시 실행하십시오.")
    token_to_col: Dict[str, str] = {}
    transformed = formula
    idx = 0

    candidates: List[Tuple[str, str]] = [(c, c) for c in sorted(col_list, key=len, reverse=True)]
    alias_map = build_alias_map(col_list)
    col_set = set(col_list)
    aliases = [(strip_unit(full), full) for key, full in alias_map.items()]
    aliases = [(a, full) for a, full in aliases if a and a != full and a not in col_set and a not in SAFE_MATH]
    candidates += sorted(aliases, key=lambda t: len(t[0]), reverse=True)

    for name, full in candidates:
        pattern = _NAME_BOUNDARY_L + re.escape(name) + _NAME_BOUNDARY_R
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
                    f"[DATA LEAKAGE DETECTED] 파생변수 '{new_col}' 수식에 타겟 변수 {leaked}이(가) 포함되어 있습니다. "
                    f"타겟 변수는 독립변수 및 파생변수 생성에 절대 투입될 수 없습니다."
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


def write_csv_excel_compatible(df: pd.DataFrame, path: str, encoding: Optional[str] = None) -> str:
    """
    Writes CSV that opens correctly by double-click in Korean Windows Excel:
    CP949 by default, utf-8-sig fallback when characters are not representable.
    Returns the path actually written (may differ if the target was locked).
    """
    def _write(p: str) -> None:
        enc = encoding or "cp949"
        try:
            df.to_csv(p, index=False, encoding=enc)
        except (UnicodeEncodeError, LookupError):
            df.to_csv(p, index=False, encoding="utf-8-sig")

    return save_with_lock_fallback(path, _write, what="CSV 파일")


# ==============================================================================
# C8. Table Interpolation (ml_map.py, and the measured-data comparison of ml_predict.py, 2.6.0+)
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
    if pd.api.types.is_datetime64_any_dtype(s):
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
    takes both."""
    req = set(required)
    given = [str(c) for c in df.columns]
    target = []
    for c in given:
        hit = [] if c in req or strip_unit(c) != c.strip() else [f for f in required if _alias_key(f) == _alias_key(c)]
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
                     f"{' / '.join(hints)}. 모델 입력이 아닌 값을 빼고 예측할지 사용자에게 물으십시오.\n[모델 입력 목록]: {list(dict.fromkeys(inputs))}")


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

def resolve_target_models(
    target_path: str,
    meta_path: Optional[str] = None,
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
        m_path = meta_path
        if not m_path:
            cand_m = os.path.join(os.path.dirname(j_path), f"model_{target_name}.json")
            if os.path.exists(cand_m):
                m_path = cand_m
            else:
                cand_m2 = os.path.splitext(j_path)[0] + ".json"
                if os.path.exists(cand_m2):
                    m_path = cand_m2

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
                f"Joblib 모델 로드 실패 ({item['joblib_path']}): {e}. "
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
# 4-B. Measured-data comparison (map check, 2.7.0)
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
    if getattr(args, "meta", None):
        parts += ["--meta", quote_arg(os.path.abspath(args.meta))]
    if getattr(args, "targets", None):
        parts += ["-y", quote_arg(args.targets)]
    if inline is not None:
        parts += ["-i", quote_arg(inline)]
    else:
        parts += ["-f", quote_arg(os.path.abspath(args.file))]
    if layout:
        parts.append(layout.strip())
    if getattr(args, "out", None):
        parts += ["-o", quote_arg(os.path.abspath(args.out))]
    if getattr(args, "encoding", None):
        parts += ["--encoding", quote_arg(args.encoding)]
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


def _layout_hint(args: argparse.Namespace, model_items: List[Dict[str, Any]], df_input: Optional[pd.DataFrame] = None,
                 missing: Optional[List[str]] = None) -> str:
    """Why a file misses the model columns, as ONE suggestion (the tool never switches the layout by itself):
    - layout options given, and the file read WITHOUT them has every model input and numbers in the first data row
      -> [추천] <the command without them>; a unit row read as data -> the command with --unit-row <header + 1>
    - the missing names are in another top row (two-level / merged header) -> [추천] --header-row N"""
    needed = list(dict.fromkeys(c for it in model_items for c in model_raw_inputs(it["meta"])))
    numeric = [c for it in model_items for c in model_raw_inputs(it["meta"])
               if c not in it["meta"]["feature_schema"].get("categorical", {})
               and c not in it["meta"]["feature_schema"].get("datetime", {})]
    numeric = list(dict.fromkeys(numeric))
    has_layout = any(getattr(args, k, None) is not None for k in ("header_row", "unit_row", "data_row"))
    if has_layout and getattr(args, "file", None):
        cmd = None
        try:
            auto = load_data_file(args.file)
        except Exception:
            auto = None
        if auto is not None and _inputs_present(needed, [str(c) for c in auto.columns]):
            # the row below the header decides (_fit_unit_layout): a unit or notes row -> the options that skip it;
            # else the first data row is data (numbers, or a measurement failure row with a number) -> no options
            layout = _fit_unit_layout(args.file, needed, None, int(auto.attrs.get("header_row", 1)) + 1, numeric)
            if layout:
                cmd = _predict_command(args, layout=layout)
            elif not auto.attrs.get("unit_row_below") and (_first_row_numeric(auto, numeric) or _row_has_data(auto, 0)):
                cmd = _predict_command(args)
        if cmd:
            return ("레이아웃 옵션(--header-row/--unit-row/--data-row)이 이 파일과 맞지 않습니다. "
                    f"이 파일의 머리글로 읽으면 모델 입력 열이 모두 있습니다. [추천] {cmd}")
    if df_input is not None and missing:
        found = find_header_layout(df_input.attrs.get("header_candidates"), missing, model_names=True)
        if found is not None and found[0] != getattr(args, "header_row", None):
            row, unit = found
            where = f"파일 {row}행에 이 이름들이 있습니다 (두 줄 머리글, 병합된 머리글이나 제목 행)"
            if unit is None and not has_layout:
                return f"{where}. [추천] --header-row {row}"
            layout = f"--header-row {row}"
            if unit:
                where += f", {unit}행은 단위 행입니다"
                # a model trained without units needs the unit row skipped (--data-row), not joined to the names
                layout = (_fit_unit_layout(args.file, needed, row, unit, numeric) if getattr(args, "file", None)
                          else "") or f"{layout} --unit-row {unit}"
            return f"{where}. [추천] " + _predict_command(args, layout=layout)
    return ""


def _unit_row_hint(args: argparse.Namespace, model_items: List[Dict[str, Any]], df_input: pd.DataFrame) -> str:
    """A file read without layout options whose first data row is a unit row ('°C', 'bar'; _looks_like_unit_row on
    the file, so a first data row of blank or '측정불가' inputs does not hide it) or a notes row: predicting it would
    impute every input. The layout that skips it (_fit_unit_layout) -> '[추천] <the command with it>', else ''."""
    if any(getattr(args, k, None) is not None for k in ("header_row", "unit_row", "data_row")) or df_input.empty:
        return ""
    needed = list(dict.fromkeys(c for it in model_items for c in model_raw_inputs(it["meta"])))
    numeric = list(dict.fromkeys(c for it in model_items for c in model_raw_inputs(it["meta"])
                                 if c not in it["meta"]["feature_schema"].get("categorical", {})
                                 and c not in it["meta"]["feature_schema"].get("datetime", {})))
    if not df_input.attrs.get("unit_row_below") and not _first_row_is_text(df_input, numeric):
        return ""
    unit_row = int(df_input.attrs.get("header_row", 1)) + 1
    layout = _fit_unit_layout(args.file, needed, None, unit_row, numeric)
    if not layout:
        return ""
    first = [str(df_input[c].iloc[0]) for c in df_input.columns[:6]]
    return (f"파일 {unit_row}행(첫 데이터 행)이 데이터가 아니라 단위·메모 행으로 보입니다 ({first}). 그대로 예측하면 "
            f"이 행을 빈 값으로 채워 예측합니다. [추천] {_predict_command(args, layout=layout)}")


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
    check_first = False
    if not is_unit and not _notes_row(raw, j):
        # 2.9.17 fallback: text in the numeric inputs of the row (a unit row the rule did not accept) -> skip it
        # when the next row reads as data
        try:
            auto = _layout_frame(raw, hdr, None, None)
        except Exception:
            return ""
        rows = auto.attrs.get("source_rows") or []
        if not numeric or unit_row not in rows or not _text_input_row(auto, rows.index(unit_row), numeric):
            return ""
        check_first = True
    tries = ([({"unit_row": unit_row, "data_row": None}, f"--unit-row {unit_row}")] if is_unit else []) + \
        [({"unit_row": None, "data_row": unit_row + 1}, f"--data-row {unit_row + 1}")]
    for opts, layout in tries:
        try:
            df = _layout_frame(raw, hdr, **opts)
        except Exception:
            continue
        if _inputs_present(needed, [str(c) for c in df.columns]) and \
                (not check_first or _first_row_numeric(df, numeric or [])):
            return head + layout
    return ""


def _first_row_is_text(df: pd.DataFrame, numeric: List[str]) -> bool:
    """Most filled numeric-input cells of the first data row are text that is not a missing-value mark (a unit row,
    a notes row read as data)."""
    share = _first_row_text_share(df, numeric)
    return share is not None and share > 0.5


def execute_predict(args: argparse.Namespace) -> Dict[str, Any]:
    """Orchestrates 1-Turn declarative inference."""
    # 1. Resolve Targets & Models
    filter_targets = [t.strip() for t in args.targets.split(",") if t.strip()] if args.targets else None
    model_items = resolve_target_models(
        target_path=args.model,
        meta_path=args.meta,
        filter_targets=filter_targets
    )

    # -o format is checked before any work: only CSV and Excel can be written and read back (R1)
    if args.out:
        out_ext = os.path.splitext(args.out)[1].lower()
        if out_ext not in (".csv", ".xlsx", ".xls"):
            raise InputError(f"결과 파일(-o)은 .csv 또는 .xlsx만 저장합니다: '{os.path.basename(args.out)}'. "
                             f"[추천] {os.path.splitext(args.out)[0]}.csv")

    # 2. Resolve Input Data (Batch File or Inline Input)
    file_warnings: List[str] = []
    if args.input:
        df_input = parse_inline_input(args.input)
        is_inline = True
        source_desc = f"Inline Input ({len(df_input)} record(s))"
    elif args.file:
        try:
            df_input = load_data_file(
                file_path=args.file,
                header_row=args.header_row,
                unit_row=args.unit_row,
                data_row=args.data_row,
                decimal_comma=any(bool(it["meta"].get("decimal_comma")) for it in model_items)
            )
        except InputError as e:
            # the layout options (often copied from training) do not fit this file: suggest the command without them
            # (one [추천] only: this checked suggestion replaces the loader's own)
            hint = _layout_hint(args, model_items)
            if hint:
                raise InputError(f"{str(e).split('[추천]')[0].rstrip()} {hint}")
            raise
        is_inline = False
        source_desc = os.path.basename(args.file)
        hint = _unit_row_hint(args, model_items, df_input)
        if hint:
            raise InputError(hint)
        file_warnings = layout_warnings(df_input, training=False)   # repeated header rows, totals rows
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
                f"모델과 같은 폴더의 .json 파일을 확인하거나 --meta로 지정하십시오."
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
                                    f"피처 '{fname}' 상한 초과 외삽(Extrapolation) 감지: 학습 최대={_fmt4(f_max)}, 입력 최대={_fmt4(c_max)} (범위 대비 +{over_pct}% 초과 이탈, 범위 밖 {n_out}건)"
                                )
                            if c_min < f_min:
                                under_pct = round(float((f_min - c_min) / f_range * 100), 1)
                                drift_warnings.append(
                                    f"피처 '{fname}' 하한 미달 외삽(Extrapolation) 감지: 학습 최소={_fmt4(f_min)}, 입력 최소={_fmt4(c_min)} (범위 대비 -{under_pct}% 미달 이탈, 범위 밖 {n_out}건)"
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
    out_path = args.out
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
            saved_file = write_csv_excel_compatible(df_result, saved_file, encoding=getattr(args, "encoding", None))

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
        sys.stderr.write(f"[WARN] Extrapolation detected: {w}\n")
    if len(drift_warnings) > WARNINGS_SHOW_MAX:
        sys.stderr.write(f"[WARN] Extrapolation detected: 외 {len(drift_warnings) - WARNINGS_SHOW_MAX}건\n")

    for w in low_conf_warnings:
        sys.stderr.write(f"[WARN] Low confidence: {w}\n")
    category_warnings = list(dict.fromkeys(category_warnings))
    for w in category_warnings:
        sys.stderr.write(f"[WARN] Unseen category: {w}\n")
    missing_warnings = [
        f"피처 '{c}': 빈 값 또는 숫자로 읽을 수 없는 값 {n}건 -> 학습 데이터 기준값(중앙값/결측 범주)으로 대치됨"
        for c, n in missing_counts.items()
    ]
    for w in missing_warnings:
        sys.stderr.write(f"[WARN] Missing input: {w}\n")
    for w in map_warnings:
        sys.stderr.write(f"[WARN] Map check: {w}\n")
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
        f"[추론 완료] 소스: {source_desc} | 타겟 예측 요약: [{targets_brief}] | "
        f"{'저장: ' + os.path.basename(saved_file) if saved_file else '인라인 반환'}"
        + (f" | 외삽 경고 {len(drift_warnings)}건" if drift_warnings else "")
        + (f" | 저확신 경고 {len(low_conf_warnings)}건" if low_conf_warnings else "")
        + (f" | 결측 대치 {sum(missing_counts.values())}건" if missing_counts else "")
        + (f" | 미학습 범주 경고 {len(category_warnings)}건" if category_warnings else "")
        + (f" | 실측 비교 경고 {len(map_warnings)}건" if map_warnings else "")
    )
    agent_directive = (
        ("VERIFIED: " if not unverified else "UNVERIFIED: 결과 파일 확인 실패 (" + "; ".join(unverified) + "). 다시 실행하지 말고 그대로 알리세요. ")
        + f"Prediction completed for {len(model_items)} target(s)"
        + (" with extrapolation warning (inputs outside training range; predictions may be unreliable)" if drift_warnings else "")
        + (f" (extrapolation_warnings shows {WARNINGS_SHOW_MAX} of {len(drift_warnings)}; tell the user the total count)"
           if len(drift_warnings) > WARNINGS_SHOW_MAX else "")
        + (" with low-confidence warning (class probability < 0.6 for some rows)" if low_conf_warnings else "")
        + (" with missing-input warning (empty or non-numeric inputs were imputed)" if missing_counts else "")
        + (" with unseen-category warning" if category_warnings else "")
        + (" with map-check warning (measured-data comparison outside the prediction interval)" if map_warnings else "")
        + ". "
        + ("Tell the user every item of warnings (rows of the file that are not data rows, numbers read without '%'). "
           if file_warnings else "")
        + ("Report each prediction together with its measured-data comparison (map_check: value, diff, method) in Korean as "
           "'실측 비교값'. " if any("map_check_keys" in p for p in predictions_summary.values()) else "")
        + (f"Inline input has {len(df_result)} records, more than {INLINE_SHOW_MAX}: per-record predictions are not in "
           f"the JSON (truncated). Give the user output_file ({os.path.basename(saved_file)}) for every record. "
           if inline_truncated and saved_file else
           f"Output saved to {os.path.basename(saved_file)}. " if saved_file else "Inline predictions returned. ")
        + "Next step: report inference results to user."
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
# 5-B. Model-based reverse lookup (--find / --vary, 2.8.0)
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
    if getattr(args, "meta", None):
        parts += ["--meta", quote_arg(os.path.abspath(args.meta))]
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
    items = resolve_target_models(target_path=args.model, meta_path=args.meta, filter_targets=[str(goal_name)],
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
                             f"[후보] {', '.join(map(str, schema['categorical'][f][:20]))}")
        else:
            raise InputError(f"모델 입력 '{f}' 값이 필요합니다. -i에 '{f}=<값>'을 넣어 주십시오.")
    warnings: List[str] = []
    for f, v in fixed.items():
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
         f"UNVERIFIED: 확인하지 못한 답이 있습니다 ({'; '.join(unverified)}). 재시도하지 말고 그대로 알리세요. ")
        + ("답을 모두 알리세요(하나만 고르지 않음). " if answers else "찾은 값 없음(count 0): reachable(가능한 예측 범위)을 알리세요. ")
        + "operating_windows(목표가 90% 예측 오차 안에 드는 운전 구간), measured(학습 데이터 기준 답), "
        + "nearest_rows(가까운 실측 5행)를 함께 알리세요. 모델이 예측한 조건이지 보장된 조건이 아니라는 점을 덧붙이세요."
        + (" warnings도 반드시 알리세요." if warnings else ""))
    return res


# ==============================================================================
# 6. CLI Entry Point
# ==============================================================================

def main():
    setup_console()
    parser = JsonArgumentParser(
        description="Zero-Office Ultra-Fast Standalone ML Inference Engine (Folder or Specific Files)"
    )
    parser.add_argument(
        "--model", "-m", required=True,
        help="학습 산출물 버전 폴더 경로(예: sample_v2_ml_data/v1/) 또는 개별 .joblib 파일 경로"
    )
    parser.add_argument(
        "--meta",
        help="개별 .joblib 지정 시 동반할 .json 메타데이터 파일 경로 (생략 시 동일 폴더 자동 탐지)"
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
        "--out", "-o",
        help="예측 결과 저장 파일 경로 (.csv, .xlsx, 생략 시 원본 파일 위치에 pred_*.csv 자동 생성)"
    )
    parser.add_argument("--header-row", type=int, help="컬럼명 헤더 행 번호 (1-based, 미지정 시 상단 빈 행 자동 건너뜀)")
    parser.add_argument("--unit-row", type=int, help="단위 행 번호 (선택)")
    parser.add_argument("--data-row", type=int, help="데이터 시작 행 번호 (선택)")
    parser.add_argument("--encoding", help="결과 CSV 저장 인코딩 (기본값: Windows Excel 최적화 cp949)")
    parser.add_argument("--overwrite", action="store_true",
                        help="--out 파일이 이미 있으면 덮어쓰기 (생략 시 시각을 붙인 새 이름으로 저장)")
    parser.add_argument("--find", help="모델 반대 찾기 목표 ('생산효율=50', 숫자 타겟). --vary와 함께")
    parser.add_argument("--vary", help="반대 찾기에서 값을 찾을 숫자 입력 (예: 반응온도). 나머지 입력은 -i로 고정, 생략하면 학습 중앙값")

    args = parser.parse_args()

    if args.find or args.vary:
        if not (args.find and args.vary):
            parser.error("모델 반대 찾기는 --find와 --vary를 함께 주어야 합니다.")
        if args.file or args.targets or args.out:
            parser.error("반대 찾기(--find)에는 -f, -y, -o를 쓰지 않습니다.")
        run_cli(lambda: execute_find(args), hint="Check the target name and the input names in SKILL.md.")
        return
    if not args.file and not args.input:
        parser.error("추론을 수행하려면 --file (배치 파일) 또는 --input (인라인 데이터) 중 하나를 지정해야 합니다.")

    run_cli(lambda: execute_predict(args), hint="Check model path or feature names.")


if __name__ == "__main__":
    main()
