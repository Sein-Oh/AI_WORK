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

VERSION = "2.9.12"                         # must equal SKILL.md metadata.version (tests check it)
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
                "If the message contains '[추천]', re-run ONCE with the suggested name; "
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


def suggest_columns(name: str, columns: List[str], n: int = 3) -> List[str]:
    alias_hit = build_alias_map(columns).get(_alias_key(name))
    hits = [alias_hit] if alias_hit else []
    for c in difflib.get_close_matches(str(name), [str(c) for c in columns], n=n, cutoff=0.5):
        if c not in hits:
            hits.append(c)
    return hits[:n]


def resolve_column_names(names: List[str], columns: List[str], kind: str) -> Tuple[List[str], Dict[str, str]]:
    """
    Resolves requested names against real columns. Exact match first, then a unique
    unit-less alias ('생산효율' -> '생산효율(%)'). A name with a different unit
    ('반응온도(C)') is NOT auto-matched. Raises InputError with suggestions otherwise.
    """
    col_list = [str(c) for c in columns]
    col_set = set(col_list)
    alias_map = build_alias_map(col_list)
    resolved: List[str] = []
    applied: Dict[str, str] = {}
    missing: List[str] = []
    for n in names:
        if n in col_set:
            resolved.append(n)
        elif strip_unit(n) == n.strip() and _alias_key(n) in alias_map:
            # Only names written WITHOUT a unit are auto-matched; a different unit is a typo
            full = alias_map[_alias_key(n)]
            resolved.append(full)
            applied[n] = full
        else:
            missing.append(n)
    if missing:
        hints = []
        for m in missing:
            s = suggest_columns(m, col_list)
            hints.append(f"'{m}' -> [추천] {s}" if s else f"'{m}' -> 유사 컬럼 없음")
        raise InputError(
            f"지정하신 {kind} 컬럼 {missing}을(를) 데이터셋에서 찾을 수 없습니다. "
            f"{' / '.join(hints)}\n[사용 가능한 컬럼 목록]: {col_list}"
        )
    for short, full in applied.items():
        sys.stderr.write(f"[INFO] 컬럼명 자동 매칭: '{short}' -> '{full}'\n")
    return resolved, applied


# ==============================================================================
# C4. Robust Data Loader & Layout Parser with Auto-Header Sniffing
# ==============================================================================

_NA_PLACEHOLDERS = ["-", "N/A", "n/a", "NA", "null", "NULL", "미측정", "None", "nan", "NaN", ""]


def coerce_numeric(series: pd.Series) -> pd.Series:
    """
    Converts a column to float regardless of its storage dtype (object in pandas 2,
    'str' in pandas 3). Strips thousands separators and placeholder NaN strings.
    Numeric and datetime columns are returned unchanged.
    """
    if pd.api.types.is_numeric_dtype(series) or pd.api.types.is_datetime64_any_dtype(series):
        return series
    s = series.astype("object").where(series.notna(), None)
    s = s.map(lambda v: v.replace(",", "").strip() if isinstance(v, str) else v)
    s = s.replace(_NA_PLACEHOLDERS, np.nan)
    return pd.to_numeric(s, errors="coerce").astype(float)


_CSV_ENCODINGS = ["utf-8-sig", "cp949", "euc-kr", "latin1"]


def _sniff_delimiter(sample: str) -> str:
    try:
        return csv.Sniffer().sniff(sample, delimiters=",;\t|").delimiter
    except csv.Error:
        return ","


def _read_csv_fast(file_path: str) -> pd.DataFrame:
    """
    Reads CSV with the fast C parser: encoding and delimiter are detected from the first
    64KB, then the whole file is parsed as strings (numeric conversion happens later in
    coerce_numeric, identically for training and inference). Falls back to the slower
    python parser only for malformed files.
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
        try:
            return pd.read_csv(file_path, header=None, sep=sep, encoding=enc, dtype=str, low_memory=False)
        except UnicodeDecodeError as e:
            last_err = e
            continue
        except pd.errors.ParserError:
            try:
                return pd.read_csv(file_path, header=None, sep=sep, encoding=enc, dtype=str, engine="python")
            except Exception as e:
                last_err = e
                continue
        except PermissionError:
            raise PermissionError(f"파일이 다른 프로그램(Excel 등)에서 잠겨 있습니다. 파일을 닫고 다시 실행하십시오: {file_path}")
        except Exception as e:
            last_err = e
            continue
    raise InputError(f"CSV 파일 로드 실패 (지원되는 인코딩 모두 시도 완료): {last_err}")


def load_data_file(
    file_path: str,
    header_row: Optional[int] = None,
    unit_row: Optional[int] = None,
    data_row: Optional[int] = None
) -> pd.DataFrame:
    """
    Loads tabular data from CSV (multi-encoding, auto delimiter) or Excel (.xlsx).
    Handles non-standard header, unit, and data row specifications.
    Includes Auto-Header Sniffing: skips blank/title top rows automatically if header_row is not explicit.
    """
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
        df_raw = _read_csv_fast(file_path)
    else:
        raise InputError(f"지원하지 않는 데이터 포맷입니다 ({ext}). .xlsx 또는 .csv 파일을 지정하십시오.")

    total_rows = len(df_raw)
    if total_rows == 0:
        raise InputError("입력 데이터 파일이 비어 있습니다.")

    # Auto-Header Sniffing: If header_row is not specified or 1, skip leading empty rows
    if header_row is not None and header_row > 1:
        header_idx = header_row - 1
    else:
        # Sniff first non-empty row containing at least 2 valid string/non-null cells
        header_idx = 0
        for r_i in range(min(15, total_rows)):
            row_items = df_raw.iloc[r_i].dropna().tolist()
            if len(row_items) >= 2 and any(isinstance(v, str) and str(v).strip() for v in row_items):
                header_idx = r_i
                break

    if header_idx >= total_rows:
        raise InputError(f"유효한 헤더 행을 찾을 수 없습니다 (전체 행 수: {total_rows}).")

    raw_headers = df_raw.iloc[header_idx].values
    headers = [str(c).strip() if pd.notna(c) and str(c).strip() else f"col_{i}" for i, c in enumerate(raw_headers)]
    auto_named = [not (pd.notna(c) and str(c).strip()) for c in raw_headers]   # header cell was blank

    # Handle unit row if provided
    if unit_row is not None and 1 <= unit_row <= total_rows:
        units = df_raw.iloc[unit_row - 1].values
        combined_headers = []
        for h, u in zip(headers, units):
            u_str = str(u).strip() if pd.notna(u) else ""
            if u_str and u_str not in ["-", "None", "nan"]:
                combined_headers.append(f"{h}({u_str})")
            else:
                combined_headers.append(h)
        headers = combined_headers

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
        last_meta_idx = header_idx
        if unit_row is not None and 1 <= unit_row <= total_rows:
            last_meta_idx = max(header_idx, unit_row - 1)
        start_idx = last_meta_idx + 1

    if start_idx >= total_rows:
        raise InputError(f"데이터 시작 행 ({start_idx + 1}) 이후에 유효한 데이터 행이 없습니다.")

    df_data = df_raw.iloc[start_idx:].copy()
    df_data.columns = headers

    # Clean rows that are entirely NaN, and entirely empty columns without a header (col_<i>). A named empty
    # column stays: training then says it has no values, prediction imputes it with a warning.
    df_data = df_data.dropna(how="all", axis=0)
    empty_auto = [h for h, auto in zip(headers, auto_named) if auto and df_data[h].isna().all()]
    if empty_auto:
        df_data = df_data.drop(columns=empty_auto)
    # 1-based row number in the source file for each data row (exact for .xlsx; CSV blank lines are skipped by the reader)
    source_rows = [int(i) + 1 for i in df_data.index]
    df_data = df_data.reset_index(drop=True)
    df_data.attrs["source_rows"] = source_rows
    return df_data


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
        m8 = _DATE8_RE.match(t)
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

_NAME_BOUNDARY_L = r"(?<![가-힣a-zA-Z0-9_])"
_NAME_BOUNDARY_R = r"(?![가-힣a-zA-Z0-9_])"


def _validate_eval_ast(formula_ast: ast.AST, allowed_identifiers: set, columns: List[str]) -> None:
    """Ensures AST contains only safe arithmetic nodes and declared identifiers."""
    for node in ast.walk(formula_ast):
        if not isinstance(node, _ALLOWED_AST_NODES):
            raise InputError(f"보안 위반: 허용되지 않은 문법 노드 감지: {type(node).__name__}")
        if isinstance(node, ast.Call) and not (isinstance(node.func, ast.Name) and node.func.id in SAFE_MATH):
            raise InputError(f"허용되지 않은 함수 호출입니다. 사용 가능 함수: {sorted(k for k in SAFE_MATH if callable(SAFE_MATH[k]))}")
        if isinstance(node, ast.Name) and node.id not in allowed_identifiers:
            s = suggest_columns(node.id, columns)
            raise InputError(
                f"수식의 식별자 '{node.id}'에 해당하는 컬럼이 없습니다"
                + (f" -> [추천] {s}" if s else "")
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
            _validate_eval_ast(tree, set(eval_env.keys()), columns)

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
    value columns with the same empty cells share one Table. Attributes other than at() come from the table of the
    first value column (range, rows for nearest_rows)."""

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
        self.values = values
        self.tables: List[Tuple[List[str], "Table"]] = []
        for key, cols in groups.items():
            n_blank = int((~masks[key]).sum())
            if n_blank and len(groups) == 1:
                warnings.append(f"구할 열에 빈칸·글자가 있는 {n_blank}행은 쓰지 않았습니다")
            elif n_blank:
                warnings.append(f"{', '.join(cols)}: 빈칸·글자가 있는 {n_blank}행은 이 열 계산에서만 뺐습니다")
            self.tables.append((cols, Table(df[masks[key]].reset_index(drop=True), keys, cols, warnings)))
        self.primary = self.tables[0][1]
        self.mode = "grid" if all(t.mode == "grid" for _, t in self.tables) else "scattered"

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
        oor: Dict[str, str] = {}
        for _, r in parts:
            oor.update(r["out_of_range"])
        return {"values": {c: vals[c] for c in self.values},
                "method": methods[0] if len(methods) == 1 else "; ".join(f"{', '.join(c)}: {r['method']}" for c, r in parts),
                "basis": parts[0][1]["basis"], "extrapolated": any(r["extrapolated"] for _, r in parts),
                "out_of_range": oor}


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


def align_input_columns(df: pd.DataFrame, required: List[str], derived: List[str]) -> Dict[str, str]:
    """
    Renames input columns written without unit suffix to the trained full name
    ('반응온도' -> '반응온도(°C)'), in place. Only unique, unambiguous matches are applied.
    """
    applied: Dict[str, str] = {}
    present = set(str(c) for c in df.columns)
    input_alias = build_alias_map([str(c) for c in df.columns])
    for feat in required:
        if feat in present or feat in derived:
            continue
        key = re.sub(r"\s+", "", strip_unit(feat)).lower()
        src = input_alias.get(key)
        if src is not None and src not in required and strip_unit(src) == src.strip():
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
    filter_targets: Optional[List[str]] = None
) -> List[Dict[str, Any]]:
    """
    Resolves model and metadata pairs from a directory (v1/, v2/) or specific file.
    Returns a list of dicts:
    [{"target": "생산효율", "joblib_path": "...", "meta_path": "...", "meta": {...}, "pipeline": ...}]
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
            if filter_targets:
                wanted = set(filter_targets) | {strip_unit(t) for t in filter_targets}
                if not ({display_target, target_name, strip_unit(display_target)} & wanted):
                    continue

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

def parse_inline_input(input_str: str) -> pd.DataFrame:
    """
    Parses inline input into a DataFrame:
    - path to a .json file (recommended on Windows: no shell quote escaping)
    - JSON object/array string
    - 'k1=v1,k2=v2' pairs
    """
    clean_str = input_str.strip()
    if clean_str.lower().endswith(".json") and os.path.isfile(clean_str):
        obj = load_json_file(clean_str)
        if isinstance(obj, dict):
            return pd.DataFrame([obj])
        if isinstance(obj, list) and obj:
            return pd.DataFrame(obj)
        raise InputError(f"입력 JSON 파일은 객체 또는 객체 배열이어야 합니다: {clean_str}")
    if clean_str.startswith("{") or clean_str.startswith("["):
        try:
            obj = json.loads(clean_str)
        except Exception as e:
            raise InputError(f"인라인 JSON 파싱 실패: {e}. Windows에서는 'k1=v1,k2=v2' 형식이나 .json 파일 경로를 권장합니다.")
        if isinstance(obj, dict):
            return pd.DataFrame([obj])
        if isinstance(obj, list) and obj:
            return pd.DataFrame(obj)
        raise InputError("인라인 JSON은 객체 또는 객체 배열이어야 합니다.")

    # key=val, key2=val2 format. Values stay text like a cell read from a file: numeric inputs are converted
    # later (prepare_features / coerce_numeric), so a category code such as '007' is not turned into 7.
    row_dict = {}
    for pair in clean_str.split(","):
        if "=" in pair:
            k, v = pair.split("=", 1)
            row_dict[k.strip()] = v.strip()
    if not row_dict:
        raise InputError(f"유효한 인라인 입력 형식이 아닙니다 (JSON 또는 'k1=v1,k2=v2'): '{input_str}'")
    return pd.DataFrame([row_dict])


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
    if args.input:
        df_input = parse_inline_input(args.input)
        is_inline = True
        source_desc = f"Inline Input ({len(df_input)} record(s))"
    elif args.file:
        df_input = load_data_file(
            file_path=args.file,
            header_row=args.header_row,
            unit_row=args.unit_row,
            data_row=args.data_row
        )
        is_inline = False
        source_desc = os.path.basename(args.file)
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
        column_aliases.update(align_input_columns(df_input, list(dict.fromkeys(raw_needed + req_features)), derived_names))
        missing_raw = [c for c in raw_needed if c not in df_input.columns]
        if missing_raw and any(c not in req_features for c in missing_raw):
            available_cols = [str(c) for c in df_input.columns]
            raise InputError(
                f"타겟 '{target_name}' 추론에 필요한 입력 '{', '.join(missing_raw)}'이(가) 입력 데이터에 없습니다 "
                f"(파생변수 계산에 쓰는 원본 입력 포함). [입력 데이터 컬럼 목록]: {available_cols}")
        df_evaled = apply_inference_evals(df_input, derived_evals)

        # Check required features
        missing_feats = [f for f in req_features if f not in df_evaled.columns]
        if missing_feats:
            available_cols = [str(c) for c in df_evaled.columns]
            hints = [f"'{m}' -> [추천] {suggest_columns(m, available_cols)}" for m in missing_feats
                     if suggest_columns(m, available_cols)]
            raise InputError(
                f"타겟 '{target_name}' 추론에 필요한 필수 피처 '{', '.join(missing_feats)}'이(가) 입력 데이터에 누락되었습니다. "
                + (" / ".join(hints) + " " if hints else "")
                + f"\n[입력 데이터 컬럼 목록]: {available_cols}"
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
                            f_range = max(f_max - f_min, 1e-6)
                            n_out = int(((col_series > f_max) | (col_series < f_min)).sum())

                            if c_max > f_max:
                                over_pct = round(float((c_max - f_max) / f_range * 100), 1)
                                drift_warnings.append(
                                    f"피처 '{fname}' 상한 초과 외삽(Extrapolation) 감지: 학습 최대={f_max:.2f}, 입력 최대={c_max:.2f} (범위 대비 +{over_pct}% 초과 이탈, 범위 밖 {n_out}건)"
                                )
                            if c_min < f_min:
                                under_pct = round(float((f_min - c_min) / f_range * 100), 1)
                                drift_warnings.append(
                                    f"피처 '{fname}' 하한 미달 외삽(Extrapolation) 감지: 학습 최소={f_min:.2f}, 입력 최소={c_min:.2f} (범위 대비 -{under_pct}% 미달 이탈, 범위 밖 {n_out}건)"
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
    for w in drift_warnings:
        sys.stderr.write(f"[WARN] Extrapolation detected: {w}\n")

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
        + (" with low-confidence warning (class probability < 0.6 for some rows)" if low_conf_warnings else "")
        + (" with missing-input warning (empty or non-numeric inputs were imputed)" if missing_counts else "")
        + (" with unseen-category warning" if category_warnings else "")
        + (" with map-check warning (measured-data comparison outside the prediction interval)" if map_warnings else "")
        + ". "
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
        "drift_warnings": drift_warnings,
        "extrapolation_warnings": drift_warnings,
        "low_confidence_warning": bool(low_conf_warnings),
        "low_confidence_warnings": low_conf_warnings,
        "missing_input_warning": bool(missing_counts),
        "missing_input_counts": missing_counts,
        "missing_input_warnings": missing_warnings,
        "unseen_category_warning": bool(category_warnings),
        "unseen_category_warnings": category_warnings,
        "map_check_warning": bool(map_warnings),
        "map_check_warnings": map_warnings,
        "preview": preview,
        "agent_directive": agent_directive
    }
    if is_inline:   # R2: per-record values are in the JSON only up to INLINE_SHOW_MAX records
        result["count"] = len(df_result)
        result["truncated"] = bool(inline_truncated)
        if inline_truncated:
            result["shown"] = 0
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


def _model_curve(pipe: Any, meta: Dict[str, Any], base: Dict[str, Any], vary: str, xs: np.ndarray) -> np.ndarray:
    df = pd.DataFrame({**{k: [v] * len(xs) for k, v in base.items()}, vary: xs})
    df = apply_inference_evals(df, meta.get("derived_evals", []))
    return np.asarray(pipe.predict(prepare_features(df, meta["feature_schema"])), dtype=float)


def execute_find(args: argparse.Namespace) -> Dict[str, Any]:
    goal = parse_inline_input(args.find).iloc[0].to_dict()
    if len(goal) != 1:
        raise InputError("--find에는 목표 하나만 주십시오 (예: --find \"생산효율=50\").")
    (goal_name, goal_val), = goal.items()
    try:
        y0 = float(goal_val)
    except (TypeError, ValueError):
        raise InputError(f"목표 값 '{goal_val}'이(가) 숫자가 아닙니다.")
    items = resolve_target_models(target_path=args.model, meta_path=args.meta, filter_targets=[str(goal_name)])
    item = items[0]
    meta, pipe, target = item["meta"], item["pipeline"], item["target"]
    if meta.get("task", "regression") != "regression":
        raise InputError(f"타겟 '{target}'은(는) 분류 모델입니다. 모델 반대 찾기는 숫자 타겟(회귀)만 됩니다.")
    schema = meta["feature_schema"]
    raw_inputs = model_raw_inputs(meta)
    stats = meta.get("feature_stats", {})
    numeric_in = [f for f in raw_inputs if f not in schema.get("categorical", {}) and f not in schema.get("datetime", {})]

    vary = resolve_column_names([args.vary.strip()], raw_inputs, "바꿀 입력(--vary)")[0][0]
    if vary not in numeric_in or vary not in stats:
        raise InputError(f"'{vary}'은(는) 숫자 입력이 아니라 반대 찾기로 바꿀 수 없습니다. [후보] {', '.join(numeric_in)}")
    given = parse_inline_input(args.input).iloc[0].to_dict() if args.input else {}
    names, _ = resolve_column_names([str(k) for k in given], raw_inputs, "고정 입력(-i)") if given else ([], {})
    fixed = {n: given[k] for k, n in zip(given, names)}
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
        answers.append({vary: round(mid, 6), **({"from": round(span[0], 6), "to": round(span[1], 6)} if span[1] > span[0] else {})})
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
            windows.append({"from": round(a, 6), "to": round(b, 6)})
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
                            "solutions": [round((a + b) / 2, 6) for a, b in m_ans],
                            "reachable": {"min": round(float(np.nanmin(mys)), 6), "max": round(float(np.nanmax(mys)), 6)}}
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
        "filled_defaults": {k: round(v, 6) for k, v in filled.items()},
        "model": meta.get("champion_model"), "training_range": {vary: [round(lo, 6), round(hi, 6)]},
        "count": len(answers), "solutions": answers, "operating_windows": windows,
        "reachable": {"min": round(reach[0], 6), "max": round(reach[1], 6), "within": f"{vary} {_fmt(lo)}~{_fmt(hi)}"},
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
