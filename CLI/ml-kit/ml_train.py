# -*- coding: utf-8 -*-
"""
ml_train.py - Autonomous Regression / Classification Machine Learning Engine with
Collision-Free AST Feature Engineering, 5-Model Competitive Benchmark,
Atomic Version Packaging, and Single Source of Truth (SSOT) Ledger.

--task regression (default) : champion by Out-of-Fold R2
--task classification       : champion by Out-of-Fold Macro-F1 (Stratified K-Fold)

The latest training version is the final version: ml_report.py --file and ml_predict.py --model <data file> use it.

Shared runtime (data loading, --eval engine, CLI protocol) is imported from _ml_runtime.py.

Designed for lightweight 26B-30B LLM agents with 1-turn declarative execution.
100% Standalone: Pure Python, Scikit-learn, Pandas, NumPy, and Joblib.
"""

import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _ml_runtime import (  # noqa: E402  shared runtime of ml-kit
    check_core_libraries, exit_missing_library, limit_threads, model_jobs, ask, numbered, command_with_value,
    setup_console, dumps_json, run_cli, JsonArgumentParser,
    InputError, sanitize_filename, normalize_version, warn_long_path, normalize_labels,
    read_text_file, load_json_file, resolve_column_names, known_columns, coerce_numeric,
    load_data_file, apply_eval_expressions, eval_source_columns,
    parse_datetime, normalize_category, prepare_features, date_part_column, input_column_owner, DATE_PARTS,
    strip_unit, layout_warnings, exclude_mark as _exclude_mark, unparsed_number_cells, decimal_comma_text_share, percent_warning, check_percent_mix,
    retry_fs, remove_tree, atomic_write_text, write_csv_excel_compatible, command_text, quote_arg,
    parse_where, resolve_where, where_text, where_mask, where_notation_variants, where_with_variants, where_retry,
    check_time, time_would_pass, Unfinished, VERSION, resume_find, resume_save, resume_clear, save_failed_warning,
    RESUME_MAX_AGE, option_retry, time_limit, final_version, step_begin, step_end, step_label, WATCHDOG_HOLD, mark_committed,
)

setup_console()

import re
import json
import time
import argparse
import warnings
from datetime import datetime
from typing import Any, Callable, Dict, List, Optional, Tuple

check_core_libraries("모델 학습")
try:  # R5: a missing library answers with 1 JSON line (missing_library + install_command)
    import numpy as np
    import pandas as pd
    import joblib

    import sklearn
    from sklearn.model_selection import GroupKFold, KFold, StratifiedGroupKFold, StratifiedKFold
    from sklearn.pipeline import Pipeline
    from sklearn.impute import SimpleImputer
    from sklearn.preprocessing import StandardScaler
    from sklearn.metrics import (
        r2_score, mean_absolute_error, root_mean_squared_error,
        f1_score, accuracy_score, balanced_accuracy_score, roc_auc_score,
        confusion_matrix, classification_report, make_scorer, get_scorer,
    )
    from sklearn.exceptions import ConvergenceWarning

    # 5 Core Regression Algorithms
    from sklearn.ensemble import HistGradientBoostingRegressor, RandomForestRegressor, ExtraTreesRegressor
    from sklearn.linear_model import RidgeCV, ElasticNetCV
    from sklearn.compose import ColumnTransformer
    from sklearn.preprocessing import OneHotEncoder
    from sklearn.base import clone

    # 5 Core Classification Algorithms (--task classification)
    from sklearn.ensemble import HistGradientBoostingClassifier, RandomForestClassifier, ExtraTreesClassifier
    from sklearn.linear_model import LogisticRegression, LogisticRegressionCV
    from sklearn.neighbors import KNeighborsClassifier
except ImportError as _import_exc:
    exit_missing_library(_import_exc, "모델 학습")

TASK_REGRESSION = "regression"
TASK_CLASSIFICATION = "classification"
PRIMARY_METRIC = {TASK_REGRESSION: "val_r2", TASK_CLASSIFICATION: "val_f1_macro"}
METRIC_LABEL = {TASK_REGRESSION: "검증 R2", TASK_CLASSIFICATION: "검증 Macro-F1"}   # 점수 용어: 학습 / 검증 (제외 행 포함/미포함)

warnings.filterwarnings("ignore", category=UserWarning)
warnings.filterwarnings("ignore", category=ConvergenceWarning)
warnings.filterwarnings("ignore", category=FutureWarning)


# ==============================================================================
# 1. Console Progress & Formatting (Clean Terminal Principle - No Emojis)
# ==============================================================================

def log_progress(percent: int, message: str) -> None:
    """One progress line on stderr, flushed: '[PROGRESS] 45% <message>' (the shared log tag, AGENTS.md; one line per
    step, so stream readers never wait for a carriage-return redraw)."""
    sys.stderr.write(f"[PROGRESS] {percent}% {message}\n")
    sys.stderr.flush()


# ==============================================================================
# 2. Bulletproof AST Formula Sandbox & Feature Engineering (--eval)
# ==============================================================================

_EVAL_NEXT = re.compile(r"\s*[^\W\d][^,;=()<>!]*?\s*=(?!=)")    # '<name> =' (not '==') after a ',' between formulas


def _eval_comma_fix(text: str) -> str:
    """'a = x/y, b = x*2' -> 'a = x/y; b = x*2': a ',' outside parentheses after a formula with '=' and before
    '<name> =' joins two formulas (2.9.69). The ',' of a function call (max(a, b)) is never touched."""
    out, depth, seen_eq = [], 0, False
    for i, ch in enumerate(text):
        if ch in "([":
            depth += 1
        elif ch in ")]":
            depth = max(depth - 1, 0)
        elif ch == ";":
            seen_eq = False
        elif ch == "=" and depth == 0:
            seen_eq = True
        if ch == "," and depth == 0 and seen_eq and _EVAL_NEXT.match(text, i + 1):
            out.append(";")
            seen_eq = False
            continue
        out.append(ch)
    return "".join(out)


def run_safe_eval(df: pd.DataFrame, eval_str: str, target_cols: List[str]) -> Tuple[pd.DataFrame, List[str], List[str]]:
    """
    Safely evaluates semicolon-separated derived feature formulas (--eval).
    Column names may be written in full ('반응온도(°C)') or without the unit suffix
    ('반응온도') when unique. Stored expressions always use full column names so that
    ml_predict.py replays them exactly. Target references are blocked (leakage guard).
    """
    if not eval_str or not eval_str.strip():
        return df, [], []
    fixed = _eval_comma_fix(eval_str.strip())
    if fixed != eval_str.strip():
        raise InputError(f"--eval '{eval_str.strip()}'에서 식 사이를 ','로 나눴습니다. 식은 ';'로 나눕니다 "
                         f"(','는 max(a, b) 같은 함수 안에서만). [추천] {option_retry(fixed, '--eval')}")
    expressions = [e.strip() for e in eval_str.split(";") if e.strip()]
    return apply_eval_expressions(df, expressions, target_cols=target_cols, strict=True)


# ==============================================================================
# 3. Robust Data Loader & Layout Parser with Auto-Header Sniffing
# ==============================================================================

def parse_feature_or_target_list(input_arg: str, param_name: str = "features") -> List[str]:
    """
    Parses comma-separated string OR file path (.txt, .json, .csv).
    If a valid file path is provided, reads feature/target names:
      - .json: list ["A", "B"] or dict {"features": [...], "targets": [...]}
      - .txt / .csv: line-by-line or comma-separated, ignoring empty lines and # comments
    If not a file, splits the string by comma.
    """
    if not input_arg or not str(input_arg).strip():
        return []

    stripped = str(input_arg).strip()

    # 1. File Path Mode
    if os.path.isfile(stripped):
        ext = os.path.splitext(stripped)[1].lower()
        items: List[str] = []

        if ext == ".json":
            data = load_json_file(stripped)
            if isinstance(data, list):
                items = [str(x).strip() for x in data if str(x).strip()]
            elif isinstance(data, dict) and data:
                cand = (
                    data.get(param_name)
                    or data.get("features")
                    or data.get("targets")
                    or data.get("columns")
                    or list(data.values())[0]
                )
                if isinstance(cand, list):
                    items = [str(x).strip() for x in cand if str(x).strip()]
        else:
            # .txt, .csv, generic text files (UTF-8/UTF-16/CP949 written by Notepad, PowerShell, Excel)
            for line in read_text_file(stripped).splitlines():
                line_str = line.strip()
                if not line_str or line_str.startswith("#"):
                    continue
                # Split line by comma in case multiple features are on the same line
                for part in line_str.split(","):
                    p = part.strip()
                    if p and not p.startswith("#"):
                        items.append(p)

        if not items:
            raise InputError(f"지정하신 파일 '{stripped}'에서 유효한 컬럼명을 추출할 수 없습니다.")
        return items

    # 2. In-line Comma-separated string mode
    return [col.strip() for col in stripped.split(",") if col.strip()]


# ==============================================================================
# 4. Fail-Fast Validation & Type Sanitization
# ==============================================================================

# --dtypes type words -> the canonical word stored in the ledger and read by infer_feature_schema
DTYPE_WORDS = {"float": "float", "float64": "float", "num": "float", "int": "int", "int64": "int",
               "str": "str", "object": "str", "category": "str", "cat": "str", "datetime": "datetime", "date": "datetime"}


def resolve_dtypes(dtypes_arg: Optional[str], columns: List[str], layout: Optional[Dict[str, Any]] = None) -> Optional[str]:
    """--dtypes 'col:type,...' with every column name resolved like --features / --targets (resolve_column_names: exact name, a unique
    name without its unit, else [후보]) and every type checked (a misspelt type -> [추천] the corrected --dtypes).
    Returns the canonical 'full column name:type' list (an unknown name or type stops the run, never silently ignored)."""
    if not dtypes_arg or not str(dtypes_arg).strip():
        return None
    if ";" in str(dtypes_arg):
        pieces = [p.strip() for p in re.split(r"[;,]", str(dtypes_arg)) if p.strip()]
        if pieces and all(":" in p for p in pieces):     # '온도:str;압력:str': the pairs are joined by ','
            raise InputError(f"--dtypes '{dtypes_arg}'에서 '<열>:<형식>' 사이를 ';'로 나눴습니다. ','로 나눕니다. "
                             f"[추천] {option_retry(','.join(pieces), '--dtypes')}")
    pairs: List[Tuple[str, str]] = []
    for raw in str(dtypes_arg).split(","):
        raw = raw.strip()
        if not raw:
            continue
        if ":" not in raw:
            raise InputError(f"--dtypes의 '{raw}'에 형식이 없습니다. '<열>:<형식>'으로 쓰십시오 "
                             f"(형식: float, int, str, datetime. 예: --dtypes \"일시:datetime,설비코드:str\").")
        name, kind = (x.strip() for x in raw.split(":", 1))
        pairs.append((name, kind))
    bad = [(n, k) for n, k in pairs if k.lower() not in DTYPE_WORDS]
    if bad:
        import difflib
        fixed, sure = [], True
        for n, k in pairs:
            if k.lower() in DTYPE_WORDS:
                fixed.append(f"{n}:{k}")
                continue
            close = difflib.get_close_matches(k.lower(), list(DTYPE_WORDS), n=1, cutoff=0.6)
            sure = sure and bool(close)
            fixed.append(f"{n}:{DTYPE_WORDS[close[0]]}" if close else f"{n}:{k}")
        rec = f" [추천] {option_retry(','.join(fixed), '--dtypes')}" if sure else ""
        raise InputError(f"--dtypes의 형식 {[k for _, k in bad]}을(를) 알 수 없습니다. 쓸 수 있는 형식: float, int, str, datetime.{rec}")
    names, _ = resolve_column_names([n for n, _ in pairs], columns, "--dtypes 열", layout=layout)
    out = dict(zip(names, (DTYPE_WORDS[k.lower()] for _, k in pairs)))
    return ",".join(f"{c}:{k}" for c, k in out.items())


def leakage_error(overlap: List[str], features: List[str], targets: List[str]) -> InputError:
    """A column given in both --features and --targets: STOP_RETRY with [후보] of the two full commands (the column taken out of --targets,
    or out of --features) and the same commands in next_commands. Which one is meant is the user's choice."""
    names = ", ".join(f"'{t}'" for t in overlap)
    cmd = command_text(drop={"--features": "features", "--features": "features", "--targets": "targets", "--targets": "targets"})
    cands: List[Dict[str, Any]] = []
    feats_wo = [f for f in features if f not in overlap]
    tgts_wo = [t for t in targets if t not in overlap]
    if cmd and not any('"' in n for n in features + targets):
        if tgts_wo:
            cands.append({"purpose": f"{names}을(를) 타겟(--targets)에서 빼고 입력으로만 쓰기",
                          "command": f"{cmd} --features {quote_arg(','.join(features))} --targets {quote_arg(','.join(tgts_wo))}"})
        if feats_wo:
            cands.append({"purpose": f"{names}을(를) 입력(--features)에서 빼고 타겟으로만 예측하기",
                          "command": f"{cmd} --features {quote_arg(','.join(feats_wo))} --targets {quote_arg(','.join(targets))}"})
    msg = (f"타겟 누수: {names}이(가) 입력(--features)과 타겟(--targets)에 모두 있습니다. 맞힐 값을 입력으로 쓰면 학습이 의미가 없습니다. "
           f"어느 쪽에서 뺄지 사용자에게 물으십시오.")
    if cands:   # the user picks: confirm_commands (next_commands would run without asking)
        text, items = numbered([c["purpose"] for c in cands], [c["command"] for c in cands], "{}")
        return InputError(f"{msg} {text}", extra={"confirm_commands": items}, error_type="needs_user_decision")
    return InputError(msg)


def validate_and_sanitize_data(
    df: pd.DataFrame,
    feature_names: List[str],
    target_names: List[str],
    dtypes_arg: Optional[str] = None,
    task: str = "regression"
) -> Tuple[pd.DataFrame, List[str], List[str], Dict[str, str]]:
    """
    Validates explicit feature (X) and target (Y) columns.
    - task='classification': targets keep their labels (normalized to strings), not coerced to float.
    - Exact names first, then unique unit-less aliases ('생산효율' -> '생산효율(%)').
    - Fail-fast with suggestions and the available columns list if spelling fails.
    - Converts every X/Y column to float regardless of pandas version (object / str dtype),
      stripping thousands separators and placeholder strings ('-', 'N/A', '미측정', ...).
    Feature columns are NOT converted here (see infer_feature_schema).
    Returns (df_clean, features, targets, aliases_applied).
    """
    all_available = [str(c) for c in df.columns]

    layout = dict(df.attrs)    # two-level / merged header, title row -> [추천] --header-row N (+ unit row)
    # --targets first; a --features typo never suggests a target (nor a --targets typo a feature): the other option's columns are excluded
    target_names, tgt_alias = resolve_column_names(target_names, all_available, "타겟(--targets)", layout=layout,
                                                   exclude=known_columns(feature_names, all_available))
    feature_names, feat_alias = resolve_column_names(feature_names, all_available, "피처(--features)", layout=layout,
                                                     exclude=target_names)
    feature_names = list(dict.fromkeys(feature_names))
    target_names = list(dict.fromkeys(target_names))

    # Data Leakage Check: Targets in Features
    overlap = [t for t in target_names if t in set(feature_names)]
    if overlap:
        raise leakage_error(overlap, feature_names, target_names)

    df_clean = df.copy()

    # Manual dtypes override if provided (e.g., "colA:float,colB:str")
    if dtypes_arg:
        for pair in dtypes_arg.split(","):
            if ":" in pair:
                c_name, c_type = pair.split(":", 1)
                c_name = c_name.strip()
                c_type = c_type.strip().lower()
                if c_name in df_clean.columns:
                    try:
                        if c_type in ["float", "float64"]:
                            df_clean[c_name] = coerce_numeric(df_clean[c_name]).astype(float)
                        elif c_type in ["int", "int64"]:
                            # a number column (blank cells stay blank and are imputed like any number: never 0)
                            df_clean[c_name] = coerce_numeric(df_clean[c_name]).astype(float)
                        elif c_type in ["str", "object"]:
                            df_clean[c_name] = df_clean[c_name].astype(str)
                        elif c_type in ["datetime", "date"]:
                            df_clean[c_name] = parse_datetime(df_clean[c_name])  # same parser as prediction
                    except Exception as e:
                        sys.stderr.write(f"[WARN] 타입 변환 실패 ({c_name} -> {c_type}): {e}\n")

    # Targets: numeric for regression (pandas 2 'object' / pandas 3 'str' safe), labels for classification.
    # Features keep their raw values here; infer_feature_schema() decides numeric/categorical/datetime.
    numeric_cols = target_names if task == TASK_REGRESSION else []
    for col in numeric_cols:
        df_clean[col] = coerce_numeric(df_clean[col])

    if task == TASK_CLASSIFICATION:
        for col in target_names:
            df_clean[col] = normalize_labels(df_clean[col])

    all_nan = [c for c in numeric_cols if df_clean[c].notna().sum() == 0]
    if all_nan:
        raise InputError(
            f"컬럼 {all_nan}에 숫자로 변환 가능한 값이 하나도 없습니다. 헤더/데이터 시작 행(--header-row, --data-row)을 확인하십시오."
            + (" 타겟이 범주(라벨)라면 --task classification을 지정하십시오." if any(c in target_names for c in all_nan) else "")
        )
    empty_targets = [c for c in target_names if df_clean[c].notna().sum() == 0]
    if empty_targets:
        raise InputError(f"타겟 컬럼 {empty_targets}에 유효한 값이 없습니다.")

    return df_clean, feature_names, target_names, {**feat_alias, **tgt_alias}


# ==============================================================================
# 4-B. Feature Type Inference (numeric / categorical / datetime)
# ==============================================================================

MAX_CATEGORIES = 200


def column_kind(raw: pd.Series) -> str:
    """How a column without --dtypes enters the model: 'numeric', 'datetime' or 'categorical' (infer_feature_schema,
    and the input diagnosis of a run without --features)."""
    non_null = raw.notna() & (raw.astype(str).str.strip() != "")
    n = int(non_null.sum())
    is_dt_obj = raw.map(lambda v: isinstance(v, (datetime, pd.Timestamp, np.datetime64)))
    if pd.api.types.is_datetime64_any_dtype(raw) or (n and is_dt_obj.sum() >= 0.9 * n):
        return "datetime"  # Excel date cells arrive as datetime objects
    if n == 0 or pd.api.types.is_numeric_dtype(raw):
        return "numeric"
    if coerce_numeric(raw).notna().sum() >= 0.9 * n:
        return "numeric"
    if parse_datetime(raw).notna().sum() >= 0.9 * n:
        return "datetime"
    return "categorical"


def infer_feature_schema(
    df: pd.DataFrame,
    feature_names: List[str],
    derived_features: List[str],
    dtypes_arg: Optional[str] = None,
) -> Dict[str, Any]:
    """
    Decides how each feature enters the model, without LLM involvement:
    - numeric     : >= 90% of non-empty values parse as numbers (thousands separators allowed)
    - datetime    : Excel dates or >= 90% 'YYYY-MM-DD[ HH:MM]' strings (YYYYMMDD only with --dtypes col:datetime) -> year/month/day/dayofweek/hour
    - categorical : everything else (one-hot encoded). ID-like columns are rejected.
    --dtypes 'col:str' forces categorical, 'col:datetime' forces datetime, 'col:float' forces numeric.
    """
    forced: Dict[str, str] = {}
    if dtypes_arg:
        for pair in dtypes_arg.split(","):
            if ":" in pair:
                c, t = pair.split(":", 1)
                t = t.strip().lower()
                kind = ("categorical" if t in ("str", "object", "category", "cat") else
                        "datetime" if t in ("datetime", "date") else
                        "numeric" if t in ("float", "float64", "int", "int64", "num") else None)
                if kind:
                    forced[c.strip()] = kind

    numeric: List[str] = []
    categorical: Dict[str, List[str]] = {}
    dates: Dict[str, List[str]] = {}

    for col in feature_names:
        raw = df[col]
        kind = forced.get(col)
        if col in derived_features:
            kind = "numeric"
        if kind is None:
            kind = column_kind(raw)

        if kind == "numeric":
            numeric.append(col)
        elif kind == "datetime":
            dt = parse_datetime(raw)
            if dt.notna().sum() == 0:
                raise InputError(f"컬럼 '{col}'에서 날짜를 읽을 수 없습니다. 'YYYY-MM-DD' 또는 'YYYYMMDD' 형식인지 확인하십시오.")
            parts = [pt for pt in DATE_PARTS if getattr(dt.dt, pt).nunique(dropna=True) > 1]
            dates[col] = parts or ["year"]
        else:
            cats = normalize_category(raw)
            uniq = sorted(cats.dropna().unique().tolist())
            n_valid = int(cats.notna().sum())
            if len(uniq) > MAX_CATEGORIES or (len(uniq) > 50 and len(uniq) > 0.5 * max(n_valid, 1)):
                raise InputError(
                    f"피처 '{col}'은(는) 문자형 값의 종류가 {len(uniq)}개로 너무 많아 ID/일련번호로 보입니다. "
                    f"--features 목록에서 제외하십시오. 숫자 컬럼인데 잘못 판별되었다면 --dtypes '{col}:float'를 지정하십시오."
                )
            categorical[col] = uniq
            sys.stderr.write(f"[INFO] 범주형 피처 감지: '{col}' ({len(uniq)}개 범주, 원-핫 인코딩)\n")

    for col, parts in dates.items():
        sys.stderr.write(f"[INFO] 날짜 피처 감지: '{col}' -> {[date_part_column(col, p) for p in parts]}\n")

    input_columns = numeric + [date_part_column(c, p) for c, ps in dates.items() for p in ps] + list(categorical)
    return {"numeric": numeric, "categorical": categorical, "datetime": dates, "input_columns": input_columns}


def build_pipeline(estimator: Any, schema: Dict[str, Any], step_name: str) -> Pipeline:
    """Leakage-free pipeline: numeric -> median impute + scale, categorical -> one-hot (unknown ignored)."""
    num_cols = [c for c in schema["input_columns"] if c not in schema["categorical"]]
    cat_cols = list(schema["categorical"])
    transformers = []
    if num_cols:
        transformers.append(("num", Pipeline([
            # keep_empty_features: a sparse column can be all-empty in a small training fold (no zero-feature error)
            ("imputer", SimpleImputer(strategy="median", keep_empty_features=True)),
            ("scaler", StandardScaler()),
        ]), num_cols))
    if cat_cols:
        transformers.append(("cat", Pipeline([
            ("imputer", SimpleImputer(strategy="constant", fill_value="__missing__")),
            ("onehot", OneHotEncoder(handle_unknown="ignore", sparse_output=False)),
        ]), cat_cols))
    pre = ColumnTransformer(transformers, remainder="drop", verbose_feature_names_out=False)
    return Pipeline([("preprocess", pre), (step_name, estimator)])


PERMUTATION_MAX_ROWS = 5000
KNN_MAX_ROWS = 50000


PERMUTATION_REPEATS = 5
IMPORTANCE_SAVE_SECONDS = 12.0   # finished importance columns are saved this often: a killed run keeps its progress


def _importance_path(resume: Optional[str]) -> Optional[str]:
    return f"{resume}.imp.joblib" if resume else None


def _permutation_means(pipe: Pipeline, Xs: pd.DataFrame, ys: np.ndarray, seed: int, scoring: Optional[str],
                       resume: Optional[str], target_name: str, champion_name: str,
                       before_stop: Callable[[], None]) -> np.ndarray:
    """
    Mean drop of the score when one input column is shuffled (PERMUTATION_REPEATS times), per column of Xs.
    Each column has its own random stream (seed, column index), so the result does not depend on how the work is
    split between runs. Split execution (R10): before each column the clock is checked with the time of one column
    (first estimated from one prediction of the whole pipeline x PERMUTATION_REPEATS, then the longest measured
    column of this run); the finished columns are saved in <resume>.imp.joblib (before_stop first) every
    IMPORTANCE_SAVE_SECONDS and at a stop, which ends with CONTINUE. Only the first column of a run that did no other
    step (a candidate, the final fit) is done without the clock, so running again always moves on and ends, and a
    run killed from outside keeps the columns saved so far.
    """
    scorer = get_scorer(scoring) if scoring else (lambda est, X_, y_: est.score(X_, y_))
    n_cols = Xs.shape[1]
    path = _importance_path(resume)
    done: Dict[int, float] = {}
    saved_at = path and resume_find(os.path.dirname(path), os.path.basename(path))
    if saved_at:
        try:
            st = joblib.load(saved_at)
            if st.get("champion") == champion_name and st.get("columns") == list(map(str, Xs.columns)):
                done = {int(k): float(v) for k, v in st["done"].items()}
        except Exception:
            sys.stderr.write(f"[WARN] 저장된 순열 중요도 {os.path.basename(saved_at)}를 읽지 못해 다시 계산합니다\n")
    todo = [j for j in range(n_cols) if j not in done]
    if not todo:
        return np.array([done[j] for j in range(n_cols)], dtype=float)
    t0 = time.monotonic()
    base = float(scorer(pipe, Xs, ys))       # the whole pipeline (transform + predict), like one shuffled column
    col_seconds = max(time.monotonic() - t0, 1e-3) * PERMUTATION_REPEATS * 1.2
    if resume:
        sys.stderr.write(f"[INFO] 타겟 '{target_name}': 순열 중요도 남은 열 {len(todo)}/{n_cols}개, "
                         f"예상 {col_seconds * len(todo):.0f}초 (예측 1회 {col_seconds / PERMUTATION_REPEATS / 1.2:.2f}초)\n")
    meta = {"champion": champion_name, "columns": list(map(str, Xs.columns))}

    def save() -> None:
        before_stop()
        if _resume_save(path, dict(meta, done=dict(done))):
            step_end()

    new_cols = 0
    measured = 0.0                             # longest column of this run
    saved_new = time.monotonic()
    other_step = _step_done()                  # this run already moved on (a fold, a candidate, the fit)
    Xw = Xs.copy()
    if resume and not other_step:              # the first column of this run, without a clock check: marked
        step_begin(os.path.dirname(resume), f"{os.path.basename(resume)}.imp",
                   f"학습: 타겟 '{target_name}' 순열 중요도 {len(done) + 1}/{n_cols}열")
    elif resume:
        step_label(f"학습: 타겟 '{target_name}' 순열 중요도 ({len(done)}/{n_cols}열 끝남)")
    for j in todo:
        if resume and (other_step or new_cols >= 1) and time_would_pass(col_seconds):
            save()
            check_time(col_seconds, f"학습: 타겟 '{target_name}' 순열 중요도 {len(done)}/{n_cols}열 끝남, 저장됨 "
                                    f"(나머지 열 전에 멈춤, 약 {1 + TARGETS_AFTER[0] * STEPS_PER_TARGET}단계 남음)")
        t_col = time.monotonic()
        orig = Xs.iloc[:, j]
        rng = np.random.default_rng([seed, j])
        scores = []
        for _ in range(PERMUTATION_REPEATS):
            Xw.isetitem(j, orig.take(rng.permutation(len(orig))).set_axis(Xw.index))
            scores.append(float(scorer(pipe, Xw, ys)))
        Xw.isetitem(j, orig)
        done[j] = base - float(np.mean(scores))
        new_cols += 1
        measured = max(measured, time.monotonic() - t_col)
        col_seconds = measured * 1.2           # measured columns replace the first estimate (a busy PC is slower)
        if resume and len(done) < n_cols and time.monotonic() - saved_new >= IMPORTANCE_SAVE_SECONDS:
            save()                             # a run killed from outside still keeps these columns
            saved_new = time.monotonic()
    return np.array([done[j] for j in range(n_cols)], dtype=float)


def feature_importance_ranks(
    pipe: Pipeline, X: pd.DataFrame, y: np.ndarray, schema: Dict[str, Any], seed: int, scoring: Optional[str],
    notes: Optional[List[str]] = None, resume: Optional[str] = None, target_name: str = "", champion_name: str = "",
    before_stop: Callable[[], None] = lambda: None,
) -> List[Dict[str, Any]]:
    """
    Permutation importance on at most PERMUTATION_MAX_ROWS rows (keeps large data fast),
    aggregated back to user-facing columns (date parts are summed into their date column; a categorical
    column is permuted as a whole because one-hot encoding happens inside the pipeline).
    notes: gets one line when the ranks carry no information (calculation failed or every
    importance is 0), so the training result can give top_features [] with a reason.
    resume: split execution of a long calculation (_permutation_means); before_stop saves what the caller must keep.
    """
    Xs, ys = X, y
    if len(X) > PERMUTATION_MAX_ROWS:
        idx = np.random.default_rng(seed).choice(len(X), PERMUTATION_MAX_ROWS, replace=False)
        Xs, ys = X.iloc[idx], y[idx]
    try:
        raw = np.maximum(_permutation_means(pipe, Xs, ys, seed, scoring, resume, target_name, champion_name,
                                            before_stop), 0)
        if notes is not None and not float(raw.sum()) > 0:
            notes.append("모든 입력의 순열 중요도가 0입니다 (입력을 섞어도 점수가 떨어지지 않음)")
    except InputError:                         # a finished step saved nowhere at a stop (check_time: STOP_RETRY)
        raise
    except Exception as e:
        raw = np.ones(X.shape[1])
        if notes is not None:
            notes.append(f"순열 중요도 계산 실패 ({type(e).__name__})")

    owner = input_column_owner(schema)
    agg: Dict[str, float] = {}
    for col, imp in zip(X.columns, raw):
        agg[owner.get(col, col)] = agg.get(owner.get(col, col), 0.0) + float(imp)
    total = sum(agg.values()) or 1.0
    ranks = [{"feature": f, "importance": round(v, 4), "share": round(v / total, 4), "pct": round(v / total * 100, 1)}
             for f, v in agg.items()]
    return sorted(ranks, key=lambda r: r["importance"], reverse=True)


TOP_FEATURES_MAX = 5      # top_features entries per target in the training result
TOP_FEATURES_PREVIEW = 3  # of which the preview line shows this many


def top_feature_shares(champion: Dict[str, Any]) -> Tuple[List[Dict[str, Any]], Optional[str]]:
    """
    The best model's most important inputs for the training result JSON, from the same permutation
    importances as leaderboard.json and the Word report (original column names, shares of the sum over all
    inputs). Returns ([], reason) when the importances carry no information.
    """
    ranks = champion.get("feature_importances") or []
    note = champion.get("importance_note")
    if note or not ranks:
        return [], note or "중요도 기록이 없습니다"
    top = [{"feature": r["feature"], "importance": round(float(r.get("share", r["pct"] / 100)), 3)}
           for r in ranks[:TOP_FEATURES_MAX]]
    extra = round(sum(t["importance"] for t in top) - 1.0, 3)
    if extra > 0:                 # rounding (0.64 + 0.344 + 0.017): the shares never add up to more than 1
        top[0]["importance"] = round(top[0]["importance"] - extra, 3)
    return top, None


def _model_params(pipe: Pipeline) -> Dict[str, Any]:
    """Hyper-parameters chosen by the CV-tuned linear models (for the leaderboard)."""
    est = pipe.steps[-1][1]
    out = {}
    for attr, key in (("alpha_", "alpha"), ("l1_ratio_", "l1_ratio"), ("C_", "C")):
        val = getattr(est, attr, None)
        if val is None:
            continue
        first = np.ravel(np.asarray(val, dtype=object))[0]
        if isinstance(first, (int, float, np.integer, np.floating)):
            out[key] = round(float(first), 6)
    return out


# ==============================================================================
# 5. 5-Model Competitive Regression Pool
# ==============================================================================

LINEAR_ALPHAS = np.logspace(-3, 3, 13)


TREE_SAMPLE_ROWS = 10000


def _bagging_kwargs(n_rows: int) -> Dict[str, Any]:
    """Large data: each tree sees a bootstrap sample of TREE_SAMPLE_ROWS rows (3~5x faster, same accuracy class)."""
    if n_rows > TREE_SAMPLE_ROWS:
        return {"bootstrap": True, "max_samples": TREE_SAMPLE_ROWS / n_rows}
    return {}


def get_regression_models(seed: int = 42, n_rows: int = 0) -> Dict[str, Any]:
    """
    Returns 5 distinct, mathematically diverse regression models:
    1. HistGradientBoostingRegressor (GBDT, non-linear, robust)
    2. RandomForestRegressor (Bagging ensemble, variance reduction)
    3. ExtraTreesRegressor (Extremely randomized trees, tabular smooth)
    4. Ridge (L2 linear; alpha auto-tuned by efficient LOO-CV)
    5. ElasticNet (L1 + L2 linear; alpha / l1_ratio auto-tuned by inner 3-fold CV)
    """
    return {
        "HistGradientBoosting": HistGradientBoostingRegressor(random_state=seed),
        "RandomForest": RandomForestRegressor(n_estimators=100, random_state=seed, n_jobs=model_jobs(n_rows), **_bagging_kwargs(n_rows)),
        "ExtraTrees": ExtraTreesRegressor(n_estimators=100, random_state=seed, n_jobs=model_jobs(n_rows), **_bagging_kwargs(n_rows)),
        "Ridge": RidgeCV(alphas=LINEAR_ALPHAS),
        "ElasticNet": ElasticNetCV(l1_ratio=[0.1, 0.5, 0.9], cv=3, max_iter=5000, random_state=seed),
    }


def fit_safe_model(estimator: Any, y_train: np.ndarray) -> Any:
    """
    Fresh copy of a candidate model whose inner settings fit a (small) training set, so tiny data never ends in an
    internal error: KNN asks for at most as many neighbours as there are rows, ElasticNetCV's inner CV needs
    2+ rows per inner fold, LogisticRegressionCV's inner stratified CV needs every class in every inner fold
    (smallest class 2 -> 2 inner folds, 1 -> plain LogisticRegression with the default C).
    Settings are unchanged when the training set is large enough.
    """
    est = clone(estimator)
    n = len(y_train)
    if isinstance(est, KNeighborsClassifier):
        if est.n_neighbors > n:
            est.set_params(n_neighbors=max(1, n))
    elif isinstance(est, ElasticNetCV):
        cv = est.cv if isinstance(est.cv, int) else 5
        if n < 2 * cv:
            est.set_params(cv=max(2, min(cv, n // 2)))
    elif isinstance(est, LogisticRegressionCV):
        k = int(pd.Series(y_train).value_counts().min())
        cv = est.cv if isinstance(est.cv, int) else 5
        if k < 2:
            est = LogisticRegression(max_iter=est.max_iter, class_weight=est.class_weight)
        elif k < cv:
            est.set_params(cv=StratifiedKFold(n_splits=max(2, k)))
    return est


PREDICTION_INTERVAL_LEVEL = 0.90


def duplicate_groups(X: pd.DataFrame, y: np.ndarray) -> Tuple[np.ndarray, int]:
    """CV group per identical row (all inputs + target), and the number of extra copies (rows - distinct rows).
    Groups are 0..g-1 in order of first appearance; with no duplicates every row is its own group (n_dup = 0)."""
    frame = X.reset_index(drop=True).astype("object").copy()
    frame["__y__"] = pd.Series(y).astype("object").to_numpy()
    codes = pd.util.hash_pandas_object(frame.astype(str), index=False).to_numpy()
    _, first, groups = np.unique(codes, return_index=True, return_inverse=True)
    order = np.argsort(np.argsort(first))          # renumber by first appearance (stable fold assignment)
    groups = order[groups]
    return groups.astype(int), int(len(y) - len(first))


def duplicate_warning(target_name: str, n_dup: int, n_rows: int, note: Optional[str] = None) -> str:
    how = (f"{note}." if note else "같은 행이 학습과 검증에 나뉘어 점수가 부풀지 않도록 묶어서 검증했습니다.")
    return (f"타겟 '{target_name}': 입력과 타겟이 완전히 같은 중복 행이 {n_dup}개 있습니다 (전체 {n_rows}행). "
            f"{how} 중복이 실수라면 데이터에서 지우십시오.")


# Candidate models of a target are saved one by one, and inside a candidate every cross-validation fold: a single
# slow target, or a single slow candidate, is split at the time limit too (R10; the check before a target cannot
# protect the first target of a run, and a candidate whose folds take longer than the limit together would start
# again from fold 1 in every run and never end). <resume>.m<k>.f<j>.joblib holds fold j of candidate k until the
# candidate is finished, <resume>.m<k>.joblib the finished candidate (its fold files are then removed); the same
# command run again reuses them. The folds are the same splits (same splitter, seed, groups) in every run, and each
# fold's model is built only from its own training rows, so a split run gives the same scores, OOF predictions and
# final model as one run.
RUN_MODEL_SECONDS: List[float] = []
RUN_FOLD_SECONDS: List[float] = []          # every fold trained in this command (all candidates, all targets)


def _step_done() -> bool:
    """True when this command already finished a step (a fold, a candidate, a final fit, a target or the saved
    preparation): only then may a clock check stop it, so every run moves on (R10)."""
    return bool(RUN_FOLD_SECONDS or RUN_MODEL_SECONDS or RUN_FIT_SECONDS or RUN_PREP_SECONDS)


def _candidate_path(resume: Optional[str], m_idx: int) -> Optional[str]:
    """Where the candidate is meant to be saved (<resume_dir>/<key>.m<k>.joblib); resume_save may put it elsewhere."""
    return f"{resume}.m{m_idx}.joblib" if resume else None


def _fold_path(resume: str, m_idx: int, fold: int) -> str:
    return f"{resume}.m{m_idx}.f{fold}.joblib"


def _candidate_load(resume: Optional[str], m_idx: int) -> Optional[Tuple[Any, ...]]:
    want = _candidate_path(resume, m_idx)
    path = want and resume_find(os.path.dirname(want), os.path.basename(want))
    if not path:
        return None
    try:
        res = joblib.load(path)
        if isinstance(res, tuple):
            return res
    except Exception:
        pass
    try:                                       # damaged saved candidate: drop it and train again (JIT: rebuilt)
        os.remove(path)
    except OSError:
        pass
    sys.stderr.write(f"[WARN] 저장된 중간 학습 결과 {os.path.basename(path)}를 읽지 못해 다시 학습합니다\n")
    return None


def _fold_load(resume: str, m_idx: int, fold: int, model_name: str, n_folds: int,
               n_val: int) -> Optional[Tuple[Dict[str, Any], float]]:
    """A fold saved by an earlier run of the same command: (its result, its seconds), or None. A damaged file, or one
    that does not match this candidate (another model, fold count or fold size), is dropped and the fold trained again
    (JIT: rebuilt)."""
    want = _fold_path(resume, m_idx, fold)
    path = resume_find(os.path.dirname(want), os.path.basename(want))
    if not path:
        return None
    try:
        st = joblib.load(path)
        if (isinstance(st, dict) and st.get("model") == model_name and st.get("folds") == n_folds
                and len(st["res"]["pred"]) == n_val):
            return st["res"], float(st["seconds"])
    except Exception:
        pass
    try:
        os.remove(path)
    except OSError:
        pass
    sys.stderr.write(f"[WARN] 저장된 교차검증 fold 결과 {os.path.basename(path)}를 읽지 못해 그 fold를 다시 학습합니다\n")
    return None


def _cv_folds(resume: Optional[str], m_idx: int, model_name: str, model_count: int, splits: List[Any],
              target_name: str, known: List[float], fold_fn: Callable[[np.ndarray, np.ndarray], Dict[str, Any]],
              log_offset: int, log_weight: int) -> Tuple[List[Dict[str, Any]], float]:
    """
    The cross-validation of one candidate, fold by fold: fold_fn(train_idx, val_idx) -> the fold's result (its OOF
    part 'pred', training score, ...). Returns the results in fold order and the candidate's seconds (sum of its
    folds, saved ones included). Split execution (R10): folds saved by an earlier run are reused; before each new
    fold the clock is checked (only after a step of this command: _step_done) with the longest fold of this
    candidate x 1.2, or, before its first fold, the longest fold known (this command's folds, this target's saved
    candidates), and no check when nothing is known; every finished fold but the last is saved
    (<resume>.m<k>.f<j>.joblib), the last one is saved with the whole candidate by the caller.
    """
    n = len(splits)
    results: List[Optional[Dict[str, Any]]] = [None] * n
    own: List[float] = []                      # seconds of this candidate's folds (saved ones included)
    if resume:
        for k, (_, val_idx) in enumerate(splits, 1):
            saved = _fold_load(resume, m_idx, k, model_name, n, len(val_idx))
            if saved is not None:
                results[k - 1] = saved[0]
                own.append(saved[1])
        if own:
            sys.stderr.write(f"[INFO] 타겟 '{target_name}': '{model_name}' 교차검증 fold {len(own)}/{n}개는 "
                             f"이전 실행에서 저장한 결과를 씀 (나눠서 진행)\n")
    for k, (train_idx, val_idx) in enumerate(splits, 1):
        if results[k - 1] is not None:
            continue
        label = f"학습: 타겟 '{target_name}' '{model_name}' 교차검증 fold {k}/{n}"
        if resume and _step_done():
            pool = own or (RUN_FOLD_SECONDS + [t / n for t in known])
            if pool:
                done = sum(r is not None for r in results)
                left = (model_count - m_idx) * n - done + 2 + TARGETS_AFTER[0] * STEPS_PER_TARGET
                check_time(max(pool) * 1.2,
                           f"학습: 타겟 '{target_name}' 후보 모델 {m_idx}/{model_count}개 끝남, '{model_name}' 교차검증 "
                           f"fold {done}/{n}개 끝남, 저장됨 (fold {k} 전에 멈춤, 약 {left}단계 남음)")
            step_label(label)
        elif resume:                           # the first step of this run, without a clock check: marked
            step_begin(os.path.dirname(resume), f"{os.path.basename(resume)}.m{m_idx}.f{k}", label)
        pct = log_offset + int(log_weight * (m_idx * n + k) / (model_count * n))
        log_progress(pct, f"Target '{target_name}' | {model_name} (Fold {k}/{n})...")
        t_fold = time.monotonic()
        res = fold_fn(train_idx, val_idx)
        seconds = time.monotonic() - t_fold
        results[k - 1] = res
        own.append(seconds)
        if resume:
            RUN_FOLD_SECONDS.append(seconds)
            if any(r is None for r in results):    # the last fold is saved with the candidate (_candidate_save)
                _resume_save(_fold_path(resume, m_idx, k),
                             {"model": model_name, "folds": n, "res": res, "seconds": seconds})
                step_end()
    return results, float(sum(own))            # type: ignore[return-value]


def _candidate_save(resume: Optional[str], m_idx: int, res: Tuple[Any, ...], seconds: float) -> None:
    """res + (seconds,): the measured time is kept with the candidate for the clock checks of later runs. Its fold
    files are removed once the candidate is saved (in every place)."""
    if not resume:
        return
    RUN_MODEL_SECONDS.append(seconds)
    if _resume_save(_candidate_path(resume, m_idx), res + (seconds,)):
        step_end()
        prefix = os.path.basename(resume) + f".m{m_idx}.f"
        resume_clear(os.path.dirname(resume), lambda name, p: name.startswith(prefix))


def _fit_check(resume: Optional[str], target_name: str, model_count: int, estimate: float) -> None:
    """Clock check before the final fit of a target (also one whose candidates were all reused): only after a step
    of this command (_step_done), so every run moves on (R10). A step saved nowhere makes check_time stop with
    STOP_RETRY instead, so '저장됨' is true when it stops."""
    label = f"학습: 타겟 '{target_name}' 1위 모델 전체 적합"
    if resume and _step_done():
        check_time(estimate, f"학습: 타겟 '{target_name}' 후보 모델 {model_count}/{model_count}개 끝남, 저장됨 "
                             f"(1위 모델 전체 적합 전에 멈춤, 약 {2 + TARGETS_AFTER[0] * STEPS_PER_TARGET}단계 남음)")
        step_label(label)
    elif resume:                               # the first step of this run, without a clock check: marked
        step_begin(os.path.dirname(resume), f"{os.path.basename(resume)}.fit", label)


def _fit_estimate(champion_cv_seconds: float, n_folds: int) -> float:
    """Seconds of the final fit of the champion on all rows, from its own measured cross-validation (n_folds fits on
    (n_folds - 1) / n_folds of the rows, with their predictions): one fit on all rows is about 1 / (n_folds - 1) of it,
    x 1.6 for models slower than linear in the rows. Not the longest candidate: a slow candidate that did not win
    (RandomForest) made the run stop before a fast final fit."""
    return champion_cv_seconds * 1.6 / max(n_folds - 1, 1)


def _final_fit_and_importance(estimator: Any, champion_name: str, step_name: str, X: pd.DataFrame, y: np.ndarray,
                              schema: Dict[str, Any], seed: int, scoring: Optional[str], resume: Optional[str],
                              target_name: str, model_count: int, fit_estimate: float, log_offset: int,
                              log_weight: int) -> Tuple[Any, List[Dict[str, Any]], List[str]]:
    """
    The last step of a target: the best model fitted on all rows, then its permutation importance. Split execution
    (R10): the fit is checked against the clock like a candidate, and when the importance must stop at the time limit
    the fitted model is saved first (<resume>.fit.joblib) with the finished importance columns, so the next run of the
    same command neither fits again nor repeats the finished columns (the same model as one run).
    """
    fit_path = f"{resume}.fit.joblib" if resume else None
    saved = None
    at = fit_path and resume_find(os.path.dirname(fit_path), os.path.basename(fit_path))
    if at:
        try:
            saved = joblib.load(at)
            if not (isinstance(saved, tuple) and len(saved) == 3 and saved[0] == champion_name):
                saved = None
        except Exception:
            saved = None
            sys.stderr.write(f"[WARN] 저장된 1위 모델 적합 결과 {os.path.basename(at)}를 읽지 못해 다시 적합합니다\n")
    fresh = saved is None
    if fresh:
        _fit_check(resume, target_name, model_count, fit_estimate)
        log_progress(log_offset + log_weight - 4, f"Target '{target_name}' | 1위 챔피언({champion_name}) 전체 데이터 적합 중...")
        t_fit = time.monotonic()
        final_pipeline = build_pipeline(fit_safe_model(estimator, y), schema, step_name)
        final_pipeline.fit(X, y)
        fit_seconds = time.monotonic() - t_fit
        if resume:
            RUN_MODEL_SECONDS.append(fit_seconds)   # a step done in this command
    else:
        final_pipeline, fit_seconds = saved[1], saved[2]
        sys.stderr.write(f"[INFO] 타겟 '{target_name}': 이전 실행에서 저장한 1위 모델 전체 적합을 씀 (나눠서 진행)\n")

    fit_saved: List[bool] = []

    def save_fit() -> None:                    # once: the importance saves its columns again and again
        if fresh and resume and not fit_saved:
            fit_saved.append(_resume_save(fit_path, (champion_name, final_pipeline, fit_seconds)))
            step_end()

    log_progress(log_offset + log_weight - 2, f"Target '{target_name}' | 순열 피처 중요도(Feature Importance) 산출 중...")
    notes: List[str] = []
    ranks = feature_importance_ranks(final_pipeline, X, y, schema, seed, scoring=scoring, notes=notes, resume=resume,
                                     target_name=target_name, champion_name=champion_name, before_stop=save_fit)
    step_end()                                 # the target's last step ended within this run
    return final_pipeline, ranks, notes


def train_and_benchmark_models(
    X: pd.DataFrame,
    y_series: pd.Series,
    target_name: str,
    schema: Dict[str, Any],
    seed: int = 42,
    log_offset: int = 20,
    log_weight: int = 50,
    resume: Optional[str] = None
) -> Tuple[Dict[str, Any], Any, Dict[str, Any]]:
    """
    Performs 5-Fold Leakage-Free Cross-Validation across the 5 models.
    Selects champion model by Out-of-Fold (OOF) R2 score.
    Fits champion model on 100% of data and extracts permutation feature importance.
    Prediction interval: 5%/95% quantiles of the champion's OOF residuals (90% level).
    """
    valid_mask = y_series.notna().values
    row_ids = np.flatnonzero(valid_mask)  # position in the loaded data (reported as 1-based data_row)
    X = X.loc[valid_mask].reset_index(drop=True)
    y = y_series.loc[valid_mask].astype(float).values

    if len(y) < 10:
        raise InputError(f"타겟 '{target_name}'의 유효 학습 샘플 수가 너무 적습니다 ({len(y)}개). 최소 10개 이상 필요합니다.")
    if np.unique(y).size <= 1:
        raise InputError(f"타겟 '{target_name}' 값이 모두 같아({float(f'{y[0]:.12g}')}) 학습할 수 없습니다. "
                         f"값이 변하는 타겟을 지정하십시오.")

    limit_threads(len(y))                       # small table: one thread is faster (MANY_THREADS_ROWS)
    models = get_regression_models(seed=seed, n_rows=int(len(y) * 0.8))
    model_count = len(models)
    # Dynamic K-Fold split: adapts gracefully to small sample sizes (N < 25)
    groups, n_dup = duplicate_groups(X, y)
    n_distinct = int(groups.max()) + 1
    if n_distinct < 10:
        raise InputError(f"타겟 '{target_name}': 입력+타겟이 다른 행이 {n_distinct}개뿐입니다 (전체 {len(y)}행, 나머지 "
                         f"{n_dup}행은 완전히 같은 중복 행). 서로 다른 행이 10개 이상 필요합니다. 중복이 실수라면 데이터를 고치십시오.")
    n_splits = min(5, max(2, len(y) // 3))
    if n_dup:
        # identical rows (inputs + target) stay in one fold: a copy in the training folds would leak the answer
        n_splits = max(2, min(n_splits, n_distinct))
        kf = GroupKFold(n_splits=n_splits, shuffle=True, random_state=seed)
    else:
        kf = KFold(n_splits=n_splits, shuffle=True, random_state=seed)
    splits = list(kf.split(X, y, groups if n_dup else None))

    leaderboard = []
    model_oof_map = {}

    known: List[float] = []                   # measured seconds of this target's candidates (saved ones included)
    for m_idx, (model_name, regressor) in enumerate(models.items()):
        saved = _candidate_load(resume, m_idx)
        if saved is not None and saved[0] == model_name and len(saved) == 4:
            model_oof_map[model_name] = saved[2]
            leaderboard.append(saved[1])
            known.append(float(saved[3]))
            continue

        def reg_fold(train_idx: np.ndarray, val_idx: np.ndarray, regressor: Any = regressor) -> Dict[str, Any]:
            # Leakage-Free Pipeline: fitted ONLY on training fold
            pipe = build_pipeline(fit_safe_model(regressor, y[train_idx]), schema, "regressor")
            pipe.fit(X.iloc[train_idx], y[train_idx])
            return {"train": r2_score(y[train_idx], pipe.predict(X.iloc[train_idx])),
                    "pred": pipe.predict(X.iloc[val_idx]), "params": _model_params(pipe)}

        folds, seconds = _cv_folds(resume, m_idx, model_name, model_count, splits, target_name, known, reg_fold,
                                   log_offset, log_weight)
        oof_preds = np.zeros(len(y))
        for (_, val_idx), f in zip(splits, folds):
            oof_preds[val_idx] = f["pred"]
        train_r2_list = [f["train"] for f in folds]
        params = folds[-1]["params"]

        model_oof_map[model_name] = oof_preds
        val_r2 = float(r2_score(y, oof_preds))
        avg_train_r2 = float(np.mean(train_r2_list))
        entry = {
            "model_name": model_name,
            "val_r2": round(val_r2, 4),
            "train_r2": round(avg_train_r2, 4),
            "overfit_gap": round(avg_train_r2 - val_r2, 4),
            "mae": float(f"{float(mean_absolute_error(y, oof_preds)):.6g}"),
            "rmse": float(f"{float(root_mean_squared_error(y, oof_preds)):.6g}")
        }
        if params:
            entry["params"] = params
        leaderboard.append(entry)
        known.append(seconds)
        _candidate_save(resume, m_idx, (model_name, entry, oof_preds), seconds)

    # Sort leaderboard by val_r2 descending
    leaderboard.sort(key=lambda item: item["val_r2"], reverse=True)
    champion_entry = leaderboard[0]
    champion_name = champion_entry["model_name"]

    # Mark rank and champion flag
    for rank, entry in enumerate(leaderboard, 1):
        entry["rank"] = rank
        entry["is_champion"] = (rank == 1)

    # Residual diagnostics of Champion Model (OOF -> honest error estimate)
    residuals = y - model_oof_map[champion_name]
    residual_skew = round(float(pd.Series(residuals).skew()), 4)
    alpha = (1 - PREDICTION_INTERVAL_LEVEL) / 2
    interval = {
        "level": PREDICTION_INTERVAL_LEVEL,
        "residual_lower": float(f"{float(np.quantile(residuals, alpha)):.6g}"),   # significant digits: tiny targets
        "residual_upper": float(f"{float(np.quantile(residuals, 1 - alpha)):.6g}"),
        "method": "champion OOF residual quantiles (constant width)",
    }

    # Fit Champion Model on 100% of data, then its permutation importance (split at the time limit)
    final_pipeline, feature_ranks, importance_notes = _final_fit_and_importance(
        models[champion_name], champion_name, "regressor", X, y, schema, seed, None, resume, target_name,
        model_count, _fit_estimate(known[list(models).index(champion_name)], len(splits)), log_offset, log_weight)

    champion_summary = {
        "task": TASK_REGRESSION,
        "primary_metric": "val_r2",
        "val_score": champion_entry["val_r2"],
        "champion_model": champion_name,
        "val_r2": champion_entry["val_r2"],
        "train_r2": champion_entry["train_r2"],
        "overfit_gap": champion_entry["overfit_gap"],
        "mae": champion_entry["mae"],
        "rmse": champion_entry["rmse"],
        "residual_skew": residual_skew,
        "prediction_interval": interval,
        "top_features": feature_ranks[:3],
        "weak_features": [f["feature"] for f in feature_ranks if f["pct"] <= 1.0],
        "feature_importances": feature_ranks
    }
    if importance_notes:
        champion_summary["importance_note"] = importance_notes[0]

    if n_dup:
        champion_summary["duplicate_rows"] = n_dup
        champion_summary["duplicate_grouped"] = True
    target_benchmark_res = {
        "target": target_name,
        "task": TASK_REGRESSION,
        "sample_count": len(y),
        "champion": champion_summary,
        "leaderboard": leaderboard
    }

    oof_df = pd.DataFrame({
        "data_row": row_ids + 1,
        "actual": y,
        "predicted": np.round(model_oof_map[champion_name], 6),
        "residual": np.round(residuals, 6),
    })
    return target_benchmark_res, final_pipeline, champion_summary, oof_df


# ==============================================================================
# 5-B. 5-Model Competitive Classification Pool (--task classification)
# ==============================================================================

def get_classification_models(seed: int = 42, balanced: bool = False, n_rows: int = 0) -> Dict[str, Any]:
    """
    Returns 5 distinct classifiers:
    1. HistGradientBoostingClassifier (GBDT, non-linear, robust)
    2. RandomForestClassifier (Bagging ensemble)
    3. ExtraTreesClassifier (Extremely randomized trees)
    4. LogisticRegression (linear baseline; C auto-tuned by inner 3-fold CV on Macro-F1)
    5. KNeighborsClassifier (instance-based; skipped above KNN_MAX_ROWS rows for speed)
    balanced=True applies class_weight='balanced' where supported (KNN has no class weights).
    """
    cw = "balanced" if balanced else None
    models = {
        "HistGradientBoosting": HistGradientBoostingClassifier(random_state=seed, class_weight=cw),
        "RandomForest": RandomForestClassifier(n_estimators=100, random_state=seed, n_jobs=model_jobs(n_rows), class_weight=cw,
                                               **_bagging_kwargs(n_rows)),
        "ExtraTrees": ExtraTreesClassifier(n_estimators=100, random_state=seed, n_jobs=model_jobs(n_rows), class_weight=cw,
                                           **_bagging_kwargs(n_rows)),
        # make_scorer object (not the "f1_macro" string): the string fails for binary labels in scikit-learn 1.9
        # l1_ratios=(0.0,) (L2, 지금 기본과 같음)·use_legacy_attributes=False: scikit-learn 1.10에서 바뀌는 기본값을 미리 정함 (FutureWarning)
        "LogisticRegression": LogisticRegressionCV(Cs=8, cv=3, max_iter=3000, class_weight=cw,
                                                   scoring=make_scorer(f1_score, average="macro"),
                                                   l1_ratios=(0.0,), use_legacy_attributes=False),
        "KNeighbors": KNeighborsClassifier(n_neighbors=5),
    }
    if n_rows > KNN_MAX_ROWS:
        sys.stderr.write(f"[INFO] 데이터 {n_rows}행 > {KNN_MAX_ROWS}행: 속도를 위해 KNeighbors 모델을 경쟁에서 제외\n")
        models.pop("KNeighbors")
    return models


def _roc_auc(y_true: np.ndarray, proba: np.ndarray, classes: List[str]) -> Optional[float]:
    try:
        if len(classes) == 2:
            return float(roc_auc_score(y_true == classes[1], proba[:, 1]))
        return float(roc_auc_score(y_true, proba, multi_class="ovr", average="macro", labels=classes))
    except Exception:
        return None


def best_binary_threshold(y: np.ndarray, proba_pos: np.ndarray, positive: str) -> Tuple[float, Dict[str, float]]:
    """Threshold on P(positive) that maximizes positive-class F1 on OOF probabilities."""
    y_pos = (y == positive)
    best_t, best_f1 = 0.5, -1.0
    for t in np.round(np.arange(0.05, 0.951, 0.01), 2):
        f1 = f1_score(y_pos, proba_pos >= t, zero_division=0)
        if f1 > best_f1 + 1e-12:
            best_t, best_f1 = float(t), float(f1)
    pred = proba_pos >= best_t
    tp = int((pred & y_pos).sum())
    stats = {
        "f1": round(best_f1, 4),
        "precision": round(tp / max(int(pred.sum()), 1), 4),
        "recall": round(tp / max(int(y_pos.sum()), 1), 4),
    }
    return best_t, stats


def train_and_benchmark_classifiers(
    X: pd.DataFrame,
    y_series: pd.Series,
    target_name: str,
    schema: Dict[str, Any],
    seed: int = 42,
    log_offset: int = 20,
    log_weight: int = 50,
    positive_class: Optional[str] = None,
    resume: Optional[str] = None
) -> Tuple[Dict[str, Any], Any, Dict[str, Any]]:
    """
    Stratified K-Fold competition of 5 classifiers. Champion = highest Out-of-Fold Macro-F1
    (fair under class imbalance). Fits champion on 100% of data and extracts permutation
    importance (scoring = f1_macro). Binary + positive_class: decision threshold tuned on
    OOF probabilities to maximize the positive-class F1.
    """
    valid_mask = y_series.notna().values
    row_ids = np.flatnonzero(valid_mask)  # position in the loaded data (reported as 1-based data_row)
    X = X.loc[valid_mask].reset_index(drop=True)
    y = y_series.loc[valid_mask].astype(str).values

    classes = sorted(pd.unique(y).tolist())
    counts = pd.Series(y).value_counts()
    class_counts = {c: int(counts[c]) for c in classes}
    if len(classes) < 2:
        raise InputError(f"타겟 '{target_name}'의 클래스가 {len(classes)}개뿐입니다. 분류에는 2개 이상의 클래스가 필요합니다.")
    rare = [c for c in classes if class_counts[c] < 2]
    if rare:
        raise InputError(
            f"타겟 '{target_name}'에 샘플이 1개뿐인 클래스 {rare}이(가) 있습니다. "
            f"클래스별 최소 2개 이상 필요합니다. 해당 행을 제외하거나 클래스를 병합하십시오. [클래스별 건수]: {class_counts}"
        )
    if len(classes) > 50:
        raise InputError(
            f"타겟 '{target_name}'의 클래스가 {len(classes)}개로 너무 많습니다. 연속형 수치라면 --task regression을 사용하십시오."
        )
    if len(y) < 10:
        raise InputError(f"타겟 '{target_name}'의 유효 학습 샘플 수가 너무 적습니다 ({len(y)}개). 최소 10개 이상 필요합니다.")
    if positive_class is not None:
        if len(classes) != 2:
            raise InputError(f"--positive-class는 이진 분류(클래스 2개)에서만 사용할 수 있습니다. 타겟 '{target_name}' 클래스: {classes}")
        if positive_class not in classes:
            raise InputError(f"--positive-class '{positive_class}'가 타겟 '{target_name}'의 클래스에 없습니다. "
                             f"놓치면 안 되는 클래스를 사용자에게 물으십시오: "
                             + ask(classes, [command_with_value(str(positive_class), str(c)) for c in classes],
                                   "--positive-class {}"))

    min_ratio = min(class_counts.values()) / len(y)
    balanced = min_ratio < 0.20
    if balanced:
        sys.stderr.write(
            f"[WARN] 타겟 '{target_name}' 클래스 불균형 감지 (최소 클래스 비율 {min_ratio:.1%}) -> class_weight='balanced' 적용\n"
        )

    limit_threads(len(y))                       # small table: one thread is faster (MANY_THREADS_ROWS)
    models = get_classification_models(seed=seed, balanced=balanced, n_rows=len(y))
    model_count = len(models)
    groups, n_dup = duplicate_groups(X, y)
    splits = None
    grouped_note = None
    if n_dup:
        # identical rows (inputs + class) stay in one fold; folds need at least one distinct row of every class
        class_groups = pd.Series(groups).groupby(pd.Series(y)).nunique()
        single = [str(c) for c, n in class_groups.items() if n < 2]
        if single:
            grouped_note = (f"클래스 {single}의 행이 모두 같은 값(서로 다른 행 1개)이라 묶어서 검증할 수 없어 "
                            f"일반 층화 교차검증을 했습니다. 이 클래스의 검증 점수는 부풀 수 있습니다")
        else:
            n_splits = max(2, min(5, int(class_groups.min())))
            sgk = StratifiedGroupKFold(n_splits=n_splits, shuffle=True, random_state=seed)
            splits = list(sgk.split(X, y, groups))
            if any(len(np.unique(y[tr])) < len(classes) for tr, _ in splits):
                splits = None
                grouped_note = ("중복 행을 묶으면 어떤 학습 fold에 빠지는 클래스가 있어 일반 층화 교차검증을 했습니다. "
                                "검증 점수는 부풀 수 있습니다")
    if splits is None:
        n_splits = max(2, min(5, min(class_counts.values())))
        skf = StratifiedKFold(n_splits=n_splits, shuffle=True, random_state=seed)
        splits = list(skf.split(X, y))

    leaderboard = []
    model_oof_map: Dict[str, np.ndarray] = {}
    model_proba_map: Dict[str, np.ndarray] = {}

    known: List[float] = []                   # measured seconds of this target's candidates (saved ones included)
    for m_idx, (model_name, clf) in enumerate(models.items()):
        saved = _candidate_load(resume, m_idx)
        if saved is not None and saved[0] == model_name and len(saved) == 5:
            model_oof_map[model_name], model_proba_map[model_name] = saved[2], saved[3]
            leaderboard.append(saved[1])
            known.append(float(saved[4]))
            continue

        def clf_fold(train_idx: np.ndarray, val_idx: np.ndarray, clf: Any = clf) -> Dict[str, Any]:
            pipe = build_pipeline(fit_safe_model(clf, y[train_idx]), schema, "classifier")
            pipe.fit(X.iloc[train_idx], y[train_idx])
            return {"train": f1_score(y[train_idx], pipe.predict(X.iloc[train_idx]), average="macro", zero_division=0),
                    "pred": pipe.predict(X.iloc[val_idx]), "proba": pipe.predict_proba(X.iloc[val_idx]),
                    "cols": [classes.index(c) for c in pipe.classes_], "params": _model_params(pipe)}

        folds, seconds = _cv_folds(resume, m_idx, model_name, model_count, splits, target_name, known, clf_fold,
                                   log_offset, log_weight)
        oof_pred = np.empty(len(y), dtype=object)
        oof_proba = np.zeros((len(y), len(classes)))
        for (_, val_idx), f in zip(splits, folds):
            oof_pred[val_idx] = f["pred"]
            oof_proba[np.ix_(val_idx, f["cols"])] = f["proba"]
        train_f1_list = [f["train"] for f in folds]
        params = folds[-1]["params"]

        oof_pred = oof_pred.astype(str)
        model_oof_map[model_name] = oof_pred
        model_proba_map[model_name] = oof_proba
        val_f1 = float(f1_score(y, oof_pred, average="macro", zero_division=0))
        avg_train_f1 = float(np.mean(train_f1_list))
        auc = _roc_auc(y, oof_proba, classes)
        entry = {
            "model_name": model_name,
            "val_f1_macro": round(val_f1, 4),
            "train_f1_macro": round(avg_train_f1, 4),
            "overfit_gap": round(avg_train_f1 - val_f1, 4),
            "accuracy": round(float(accuracy_score(y, oof_pred)), 4),
            "balanced_accuracy": round(float(balanced_accuracy_score(y, oof_pred)), 4),
            "roc_auc": round(auc, 4) if auc is not None else None,
        }
        if params:
            entry["params"] = params
        leaderboard.append(entry)
        known.append(seconds)
        _candidate_save(resume, m_idx, (model_name, entry, oof_pred, oof_proba), seconds)

    leaderboard.sort(key=lambda item: item["val_f1_macro"], reverse=True)
    for rank, entry in enumerate(leaderboard, 1):
        entry["rank"] = rank
        entry["is_champion"] = (rank == 1)
    champion_entry = leaderboard[0]
    champion_name = champion_entry["model_name"]

    # Binary decision threshold for the chosen positive class (default 0.5 otherwise)
    threshold_info = None
    champion_oof = model_oof_map[champion_name]
    if positive_class is not None:
        pos_idx = classes.index(positive_class)
        thr, stats = best_binary_threshold(y, model_proba_map[champion_name][:, pos_idx], positive_class)
        neg = classes[1 - pos_idx]
        champion_oof = np.where(model_proba_map[champion_name][:, pos_idx] >= thr, positive_class, neg).astype(str)
        threshold_info = {"positive_class": positive_class, "threshold": thr, "oof_positive_metrics": stats}
        sys.stderr.write(f"[INFO] 타겟 '{target_name}' 양성 클래스 '{positive_class}' 판정 기준 확률 = {thr:.2f} (양성 F1 최대화)\n")

    cm = confusion_matrix(y, champion_oof, labels=classes)
    report = classification_report(y, champion_oof, labels=classes, output_dict=True, zero_division=0)
    per_class = {
        c: {
            "precision": round(float(report[c]["precision"]), 4),
            "recall": round(float(report[c]["recall"]), 4),
            "f1": round(float(report[c]["f1-score"]), 4),
            "support": int(report[c]["support"]),
        }
        for c in classes
    }

    final_pipeline, feature_ranks, importance_notes = _final_fit_and_importance(
        models[champion_name], champion_name, "classifier", X, y, schema, seed, "f1_macro", resume, target_name,
        model_count, _fit_estimate(known[list(models).index(champion_name)], len(splits)), log_offset, log_weight)

    champion_summary = {
        "task": TASK_CLASSIFICATION,
        "primary_metric": "val_f1_macro",
        "val_score": champion_entry["val_f1_macro"],
        "champion_model": champion_name,
        "val_f1_macro": champion_entry["val_f1_macro"],
        "train_f1_macro": champion_entry["train_f1_macro"],
        "overfit_gap": champion_entry["overfit_gap"],
        "accuracy": champion_entry["accuracy"],
        "balanced_accuracy": champion_entry["balanced_accuracy"],
        "roc_auc": champion_entry["roc_auc"],
        "classes": classes,
        "class_counts": class_counts,
        "class_weight": "balanced" if balanced else None,
        "decision_threshold": threshold_info,
        "per_class": per_class,
        "confusion_matrix": {"labels": classes, "matrix": cm.tolist()},
        "top_features": feature_ranks[:3],
        "weak_features": [f["feature"] for f in feature_ranks if f["pct"] <= 1.0],
        "feature_importances": feature_ranks
    }
    if importance_notes:
        champion_summary["importance_note"] = importance_notes[0]
    if n_dup:
        champion_summary["duplicate_rows"] = n_dup
        champion_summary["duplicate_grouped"] = grouped_note is None
        if grouped_note:
            champion_summary["duplicate_note"] = grouped_note
    target_benchmark_res = {
        "target": target_name,
        "task": TASK_CLASSIFICATION,
        "sample_count": len(y),
        "champion": champion_summary,
        "leaderboard": leaderboard
    }
    oof_df = pd.DataFrame({"data_row": row_ids + 1, "actual": y, "predicted": champion_oof})
    for i_c, c in enumerate(classes):
        oof_df[f"proba_{c}"] = np.round(model_proba_map[champion_name][:, i_c], 6)
    return target_benchmark_res, final_pipeline, champion_summary, oof_df


# ==============================================================================
# 6. Atomic Transaction Guard & Revision Management
# ==============================================================================

class TransactionGuard:
    """
    Manages atomic version folders:
    - Removes abandoned .tmp_* staging folders (created by this tool) on startup.
    - Never deletes a vN folder it did not create in this run: incomplete folders are reported and skipped.
    - Allocates next version folder (v1, v2, ...). An existing --version is refused unless overwrite=True.
    - Stages all files in .tmp_vX/ and atomically renames it to vX/ (commit).
    - With overwrite, the old vX is parked as .bak_vX until finalize(), so a failure restores it.
    - rollback() undoes everything not yet finalized (staging folder, committed folder, parked backup).
    """
    def __init__(self, data_root: str, target_version: Optional[str] = None, overwrite: bool = False):
        self.data_root = os.path.abspath(data_root)
        os.makedirs(self.data_root, exist_ok=True)
        self.target_version = target_version
        self.overwrite = overwrite
        self.version_str = ""
        self.tmp_dir = ""
        self.final_dir = ""
        self.backup_dir = ""
        self.is_committed = False
        self.is_finalized = False
        self.incomplete_versions: List[str] = []

    def _next_version(self) -> str:
        nums = [int(item[1:]) for item in os.listdir(self.data_root)
                if re.match(r"^v\d+$", item) and os.path.isdir(os.path.join(self.data_root, item))]
        return f"v{max(nums, default=0) + 1}"

    def setup(self) -> str:
        # 1. Startup Garbage Collection of broken temporary folders
        self._garbage_collect()

        # 2. Determine Version
        if self.target_version:
            self.version_str = normalize_version(self.target_version)
            if not re.match(r"^v[0-9A-Za-z_.-]+$", self.version_str):
                raise InputError(f"버전명은 영문/숫자만 사용할 수 있습니다 (예: v3): '{self.target_version}'")
            if os.path.exists(os.path.join(self.data_root, self.version_str)) and not self.overwrite:
                raise InputError(
                    f"버전 폴더 '{self.version_str}'이(가) 이미 있어 덮어쓰지 않았습니다. "
                    f"기존 {self.version_str}를 바꾸는 것은 사용자가 요청한 경우에만 --overwrite로 합니다. "
                    f"[추천] {option_retry(self._next_version(), '--version', '--version')}"
                )
        else:
            # Auto-infer next version (incomplete vN folders still take their number)
            self.version_str = self._next_version()

        self.tmp_dir = os.path.join(self.data_root, f".tmp_{self.version_str}")
        self.final_dir = os.path.join(self.data_root, self.version_str)

        # Clean any preexisting tmp folder
        remove_tree(self.tmp_dir)
        os.makedirs(self.tmp_dir, exist_ok=True)

        return self.version_str

    def _garbage_collect(self) -> None:
        """Removes orphaned .tmp_* staging folders, restores parked .bak_* backups, reports incomplete vN folders."""
        for item in sorted(os.listdir(self.data_root)):
            p = os.path.join(self.data_root, item)
            if not os.path.isdir(p):
                continue
            if item.startswith(".tmp_"):
                try:
                    remove_tree(p)
                    sys.stderr.write(f"[INFO] 고아 임시 폴더 자동 청소: {item}\n")
                except Exception:
                    pass
            elif item.startswith(".bak_"):
                original = os.path.join(self.data_root, item[len(".bak_"):])
                if not os.path.exists(original):
                    try:
                        retry_fs(lambda: os.rename(p, original))
                        sys.stderr.write(f"[WARN] 중단된 덮어쓰기 복원: {item} -> {os.path.basename(original)}\n")
                    except Exception:
                        pass
                else:
                    sys.stderr.write(f"[WARN] 이전 덮어쓰기의 백업 폴더 {item}이(가) 남아 있습니다. 지우지 않았습니다.\n")
            elif re.match(r"^v\d+$", item):
                lead_p = os.path.join(p, "leaderboard.json")
                ds_p = os.path.join(p, "model_dataset.csv")
                if not (os.path.exists(lead_p) and os.path.exists(ds_p)):
                    self.incomplete_versions.append(item)
                    sys.stderr.write(
                        f"[WARN] 불완전한 버전 폴더 {item} (leaderboard.json 또는 model_dataset.csv 없음): "
                        f"지우지 않았습니다. 필요 없으면 사용자가 직접 지우세요. 이 버전 번호는 건너뜁니다.\n"
                    )

    def commit(self) -> str:
        """Atomically renames .tmp_vX to vX (retries transient Windows locks from antivirus/Explorer)."""
        if os.path.exists(self.final_dir):
            if not self.overwrite:
                raise InputError(f"버전 폴더 '{self.version_str}'이(가) 이미 있어 덮어쓰지 않았습니다.")
            # Park the old version; it is deleted only after the new one is recorded (finalize)
            backup = os.path.join(self.data_root, f".bak_{self.version_str}")
            if os.path.exists(backup):
                # A leftover backup may be the only good copy of an earlier version: never delete it here
                raise InputError(f"이전 덮어쓰기의 백업 폴더 '.bak_{self.version_str}'이(가) 남아 있어 진행하지 않았습니다. "
                                 f"사용자가 내용을 확인한 뒤 직접 정리하세요.")
            self.backup_dir = backup
            retry_fs(lambda: os.rename(self.final_dir, self.backup_dir))
        retry_fs(lambda: os.rename(self.tmp_dir, self.final_dir))
        self.is_committed = True
        return self.final_dir

    def finalize(self) -> None:
        """Called once the ledger records the committed version: the parked old version is no longer needed."""
        self.is_finalized = True
        if self.backup_dir:
            try:
                remove_tree(self.backup_dir)
            except Exception as e:
                sys.stderr.write(f"[WARN] 이전 버전 백업 폴더 정리 실패 (다음 학습 때 다시 안내): {e}\n")

    def rollback(self, quiet: bool = False) -> None:
        """Undoes a run that was not finalized: staging folder, committed folder and parked backup.
        quiet: a run stopped at the time limit (Unfinished, R10) is not an abnormal stop (no [WARN])."""
        if self.is_finalized:
            return
        try:
            if self.is_committed and os.path.exists(self.final_dir):
                remove_tree(self.final_dir)  # created by this run, not yet in the ledger
            if self.backup_dir and os.path.exists(self.backup_dir) and not os.path.exists(self.final_dir):
                retry_fs(lambda: os.rename(self.backup_dir, self.final_dir))
                sys.stderr.write(f"[WARN] 비정상 중단: 기존 {self.version_str} 폴더를 되돌렸습니다\n")
            if self.tmp_dir and os.path.exists(self.tmp_dir):
                remove_tree(self.tmp_dir)
            if not quiet:
                sys.stderr.write(f"[WARN] 비정상 중단: 임시 스테이징 폴더 롤백 완료 ({self.tmp_dir})\n")
        except Exception:
            pass


def verify_version_dir(version_dir: str, target_names: List[str], ledger_path: str, version: str) -> List[str]:
    """Re-reads a committed version folder. Returns what could not be confirmed (empty = verified)."""
    problems: List[str] = []
    try:
        load_json_file(os.path.join(version_dir, "leaderboard.json"))
    except Exception as e:
        problems.append(f"leaderboard.json 읽기 실패: {e}")
    ds = os.path.join(version_dir, "model_dataset.csv")
    if not (os.path.exists(ds) and os.path.getsize(ds) > 0):
        problems.append("model_dataset.csv 없음 또는 빈 파일")
    for t in target_names:
        safe = sanitize_filename(t)
        try:
            joblib.load(os.path.join(version_dir, f"model_{safe}.joblib"))
        except Exception as e:
            problems.append(f"model_{safe}.joblib 다시 읽기 실패: {type(e).__name__}")
        try:
            meta = load_json_file(os.path.join(version_dir, f"model_{safe}.json"))
            if meta.get("version") != version or meta.get("target") != t:
                problems.append(f"model_{safe}.json의 version/target 불일치")
        except Exception as e:
            problems.append(f"model_{safe}.json 읽기 실패: {type(e).__name__}")
    records = load_history_ledger(ledger_path)
    if not records or records[-1].get("version") != version:
        problems.append(f"history.jsonl 마지막 기록이 {version}이 아님")
    return problems


# ==============================================================================
# 7. SSOT Ledger (history.jsonl) Management
# ==============================================================================

def append_history_ledger(ledger_path: str, record: Dict[str, Any],
                          on_recorded: Optional[Callable[[], None]] = None) -> None:
    """
    Appends experiment record to history.jsonl and automatically synchronizes
    history.json (standard indented JSON array) as Single Source of Truth (SSOT).
    on_recorded runs right after the jsonl line is written, before the json sync, so a training run
    can mark its version folder final the moment the ledger names it (no rollback after that point).
    """
    data_root = os.path.dirname(ledger_path)
    os.makedirs(data_root, exist_ok=True)

    # 1. Append to history.jsonl (streaming ledger). A version replaced with --overwrite keeps one record: the old
    # record of that version is removed (the ledger is rewritten atomically), so the final version and the
    # is_best / delta comparisons never see the replaced folder's scores.
    ver = record.get("version")
    existing = load_history_ledger(ledger_path) if ver else []
    if any(r.get("version") == ver for r in existing):
        kept = [r for r in existing if r.get("version") != ver]
        atomic_write_text(ledger_path, "".join(dumps_json(r) + "\n" for r in kept + [record]))
        sys.stderr.write(f"[INFO] 원장(history.jsonl)의 {ver} 기록을 이번 기록으로 교체했습니다 (--overwrite)\n")
    else:
        with open(ledger_path, "a", encoding="utf-8", newline="\n") as f:
            f.write(dumps_json(record) + "\n")
    if on_recorded:
        on_recorded()

    # 2. Synchronize history.json (standard formatted JSON array), written atomically
    json_path = os.path.join(data_root, "history.json")
    all_records = load_history_ledger(ledger_path)
    try:
        atomic_write_text(json_path, dumps_json(all_records, indent=2))
    except OSError as e:
        sys.stderr.write(f"[WARN] history.json 동기화 실패 (history.jsonl은 정상 기록됨): {e}\n")


def load_history_ledger(ledger_path: str) -> List[Dict[str, Any]]:
    """Loads all records from history.jsonl."""
    if not os.path.exists(ledger_path):
        return []
    records = []
    with open(ledger_path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line:
                try:
                    records.append(json.loads(line))
                except Exception:
                    sys.stderr.write("[WARN] history.jsonl 손상 라인 1건을 건너뜀\n")
    return records


# ==============================================================================
# 7-B. Run Lock, Bounded Correlation, Next Commands
# ==============================================================================

class RunLock:
    """
    Prevents two trainings on the same dataset at once (they would collide on version numbers
    and the startup garbage collection could delete the other run's staging folder).
    Uses an OS file lock, so a crashed process never leaves a stale lock behind.
    """

    def __init__(self, data_root: str):
        self.path = os.path.join(data_root, ".train.lock")
        self.fh = None

    def acquire(self) -> None:
        os.makedirs(os.path.dirname(self.path), exist_ok=True)
        fh = open(self.path, "a+")
        try:
            if os.name == "nt":
                import msvcrt
                fh.seek(0)
                msvcrt.locking(fh.fileno(), msvcrt.LK_NBLCK, 1)
            else:
                import fcntl
                fcntl.flock(fh.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
        except OSError:
            fh.close()
            raise InputError(
                f"같은 데이터에 대한 다른 학습이 이미 실행 중입니다 ({os.path.dirname(self.path)}). "
                f"동시에 실행할 수 없으니 진행 중인 학습이 끝난 뒤 다시 실행하십시오."
            )
        self.fh = fh

    def release(self) -> None:
        if self.fh is None:
            return
        try:
            if os.name == "nt":
                import msvcrt
                self.fh.seek(0)
                msvcrt.locking(self.fh.fileno(), msvcrt.LK_UNLCK, 1)
            else:
                import fcntl
                fcntl.flock(self.fh.fileno(), fcntl.LOCK_UN)
        except OSError:
            pass
        self.fh.close()
        self.fh = None


CORR_MAX_FEATURES = 30


def bounded_correlation(
    corr_full: pd.DataFrame, importances: List[Dict[str, Any]], schema: Dict[str, Any], targets: List[str]
) -> Tuple[Dict[str, Any], bool]:
    """
    Ledger copy of the correlation matrix: all columns when small, otherwise the
    CORR_MAX_FEATURES most important numeric inputs plus targets (full matrix: vX/correlation.csv).
    """
    inputs = [c for c in corr_full.columns if c not in targets]
    if len(inputs) <= CORR_MAX_FEATURES:
        return corr_full.to_dict(), False
    owner = input_column_owner(schema)
    order = {r["feature"]: i for i, r in enumerate(importances)}
    ranked = sorted(inputs, key=lambda c: order.get(owner.get(c, c), len(order)))
    keep = ranked[:CORR_MAX_FEATURES] + targets
    return corr_full.loc[keep, keep].to_dict(), True


def predict_next_command(model_ref: str, purpose: str) -> Dict[str, Any]:
    """The prediction command for a new file, without layout options: ml_predict.py finds the header of each file
    (a layout that does not fit the new file is recovered in place, so copying the training layout is not needed)."""
    here = os.path.dirname(os.path.abspath(__file__))
    return {"purpose": purpose,
            "command": f"python {quote_arg(os.path.join(here, 'ml_predict.py'))} --model {quote_arg(model_ref)} --file {quote_arg('<새데이터.xlsx>')}",
            "needs_input": True, "placeholder": "<새데이터.xlsx>"}


def predict_one_command(model_ref: str, inputs: List[str]) -> Dict[str, Any]:
    """The one-case prediction command with the model's own input names (2.9.75, 집 PC opencode 2026-11-07: with only the
    --file template and the retraining '--eval "a = b * c"' commands in front of them, Gemma 4 wrote '--eval "온도=182; ..."' and
    solar wrote a file for one condition). <값> is the user's value for each input."""
    here = os.path.dirname(os.path.abspath(__file__))
    pairs = ",".join(f"{c}=<값>" for c in inputs)
    return {"purpose": "새 조건 한 건 예측 (<값>을 사용자가 준 값으로 바꿀 것, 파일을 만들지 않음)",
            "command": f"python {quote_arg(os.path.join(here, 'ml_predict.py'))} --model {quote_arg(model_ref)} --input {quote_arg(pairs)}",
            "needs_input": True, "placeholder": "<값>"}


def build_next_commands(
    args: argparse.Namespace, raw_file: str, task: str, positive_class: Optional[str],
    feature_names: List[str], created_features: List[str], applied_evals: List[str],
    target_names: List[str], champion: Dict[str, Any], schema: Dict[str, Any],
    mean_score: float, staging_dir: str, final_dir: str, exclude_where: Optional[str] = None,
) -> List[Dict[str, Any]]:
    """
    Ready-to-run follow-up commands so a small LLM does not have to compose arguments:
    deployment/prediction when the goal is met, otherwise feature pruning and an interaction feature.
    Long feature lists, and name lists with a double quote (cannot be quoted on the command line), are written to
    <version>/next_features_*.txt (targets: next_targets.txt) and passed by path. A command whose --eval, --dtypes or
    --positive-class text would need a double quote is left out.
    Every command has needs_input: commands with a placeholder (e.g. <새데이터.xlsx>) are True,
    carry the placeholder name, and are listed after the commands that can run as-is.
    """
    here = os.path.dirname(os.path.abspath(__file__))
    train_py = quote_arg(os.path.join(here, "ml_train.py"))
    base = [f"python {train_py}", "--file", quote_arg(raw_file)]
    for flag, val in (("--header-row", args.header_row), ("--unit-row", args.unit_row),
                      ("--data-row", args.data_row), ("--dtypes", args.dtypes)):
        if val is not None:
            base += [flag, quote_arg(val)]
    no_where = list(base)                     # for the '<조건> 행을 빼고 다시 학습' template (exclude_command)
    if exclude_where is not None:
        base += ["--exclude-where", quote_arg(exclude_where)]
    task_flags: List[str] = []
    if task == TASK_CLASSIFICATION:
        task_flags = ["--task", TASK_CLASSIFICATION]
        if positive_class is not None:
            task_flags += ["--positive-class", quote_arg(positive_class)]
    base += task_flags

    def names_arg(names: List[str], fname: str, limit: int = 15) -> str:
        if len(names) <= limit and not any('"' in n for n in names):
            return quote_arg(",".join(names))
        with open(os.path.join(staging_dir, fname), "w", encoding="utf-8") as f:
            f.write("\n".join(names) + "\n")
        return quote_arg(os.path.join(final_dir, fname))

    def features_arg(feats: List[str], tag: str) -> str:
        return names_arg(feats, f"next_features_{tag}.txt")

    def quotable(*texts: Optional[str]) -> bool:
        return not any(t is not None and '"' in str(t) for t in texts)

    targets_arg = ["--targets", names_arg(target_names, "next_targets.txt", limit=10 ** 6)]
    base += targets_arg
    retrain_ok = quotable(args.dtypes, positive_class, exclude_where)

    def exclude_command() -> Optional[Dict[str, Any]]:
        """The same training (same inputs, --features written out) with rows matching <조건> left out (2.9.81 집 PC Gemma A3:
        '생산효율이 80보다 큰 행은 빼고 다시 학습' -> the long command was written into the answer, not run). An earlier
        --exclude-where is kept: '<before>; <조건>' (';' = and)."""
        if not (retrain_ok and quotable(*base_features)) or not base_features:
            return None
        where = f"{exclude_where}; <조건>" if exclude_where else "<조건>"
        parts = no_where + task_flags + targets_arg + ["--features", features_arg(base_features, "all"),
                                                       "--exclude-where", quote_arg(where)]
        if applied_evals and quotable(*applied_evals):
            parts += ["--eval", quote_arg("; ".join(applied_evals))]
        return {"purpose": "조건에 맞는 행을 빼고 다시 학습 (<조건>을 사용자가 말한 조건으로 바꿀 것, 예: 생산효율 > 80)",
                "command": " ".join(parts), "needs_input": True, "placeholder": "<조건>"}

    base_features = [f for f in feature_names if f not in created_features]
    cmds: List[Dict[str, Any]] = []

    if mean_score >= 0.90:  # caller passes the weakest target's score
        cmds.append({
            "purpose": REPORT_PURPOSE[task],
            "command": f"python {quote_arg(os.path.join(here, 'ml_report.py'))} --file {quote_arg(raw_file)}",
            "needs_input": False,
        })
        # 방금 학습한 버전 폴더 (그 버전의 입력 이름과 같은 --input 틀)
        cmds.append(predict_next_command(
            final_dir, "목표 성능 달성: 새 데이터 예측 (<새데이터.xlsx>를 사용자가 준 실제 파일 경로로 바꿀 것)"))
        if base_features and quotable(*base_features):
            cmds.append(predict_one_command(final_dir, base_features))
        cmds += [c for c in (exclude_command(),) if c]
        return cmds

    weak = [w for w in champion.get("weak_features", []) if w in feature_names]
    keep_base = [f for f in base_features if f not in weak]
    keep_evals = [e for e in applied_evals if e.split("=", 1)[0].strip() not in weak]
    if weak and keep_base and retrain_ok and quotable(*keep_evals):
        parts = base + ["--features", features_arg(keep_base, "pruned")]
        if keep_evals:
            parts += ["--eval", quote_arg("; ".join(keep_evals))]
        parts += ["--tag", quote_arg("저기여피처_제외")]
        cmds.append({"purpose": f"기여도 1% 이하 피처 {weak[:5]} 제외 후 재학습", "command": " ".join(parts),
                     "needs_input": False})

    numeric_top = [r["feature"] for r in champion.get("top_features", []) if r["feature"] in schema["numeric"]]
    if len(numeric_top) >= 2:
        a, b = numeric_top[0], numeric_top[1]
        new_name = f"{strip_unit(a)}x{strip_unit(b)}"
        existing = set(feature_names) | {e.split("=", 1)[0].strip() for e in applied_evals}
        if new_name not in existing and retrain_ok and quotable(*applied_evals, a, b):
            evals = applied_evals + [f"{new_name} = {a} * {b}"]
            parts = base + ["--features", features_arg(base_features, "all"), "--eval", quote_arg("; ".join(evals)),
                            "--tag", quote_arg("교차항_추가")]
            cmds.append({"purpose": f"상위 피처 '{a}', '{b}' 교차항 추가 후 재학습", "command": " ".join(parts),
                         "needs_input": False})
    if not cmds:
        # 2.9.76: 다시 학습할 후보가 없으면 첫 명령이 <값> 틀이 되어 "next_commands[0]으로 개선"이 예측을 실행했음 -> 보고서를 앞에
        cmds.append({"purpose": "다시 학습할 후보(약한 입력, 교차항)가 없음: " + REPORT_PURPOSE[task],
                     "command": f"python {quote_arg(os.path.join(here, 'ml_report.py'))} --file {quote_arg(raw_file)}",
                     "needs_input": False, "no_retrain": True})
    if base_features and quotable(*base_features):        # 목표 전이라도 예측은 할 수 있음 (placeholder 명령은 늘 뒤에)
        cmds.append(predict_one_command(final_dir, base_features))
    cmds += [c for c in (exclude_command(),) if c]
    return cmds


# ==============================================================================
# 8. Main Execution Controller
# ==============================================================================

def _same_experiment(record: Dict[str, Any], task: str, targets: List[str]) -> bool:
    """Ledger records are comparable (delta_score, is_best) only for the same task and target set."""
    return record.get("task", TASK_REGRESSION) == task and sorted(record.get("targets", [])) == sorted(targets)


# ==============================================================================
# 7-B. Row Exclusion: '제외' column + automatic outlier detection
# ==============================================================================
#
# The data file is never modified. A column named '제외' (fixed name, unit suffix ignored) is read per row:
#   blank   -> the automatic outlier check may exclude the row
#   '포함'  -> always used (the automatic check never excludes it)
#   other   -> excluded by the user; the text is the reason
# Excluded rows are left out of training and of the primary (cross-validated) score. They are scored
# separately by the final model, which never saw them, so the run also reports an all-rows score.

EXCLUDE_COL = "제외"
INCLUDE_MARK = "포함"
SNAPSHOT_ROW_COL = "원본행"              # model_dataset.csv: Excel row number of each row
SNAPSHOT_EXCL_PREFIX = "학습제외_"       # model_dataset.csv: 1 = left out of training for that target
OUTLIER_Z = 3.5               # robust z-score (median / MAD) above which a value is an outlier
OUTLIER_MAX_FRACTION = 0.10   # more candidates than this is a skewed distribution, not outliers: exclude none
MIN_TRAIN_ROWS = 10


def find_exclude_column(columns: List[Any]) -> Optional[str]:
    for c in columns:
        if strip_unit(str(c)) == EXCLUDE_COL:
            return str(c)
    return None


def _outlier_flags(values: pd.Series) -> Tuple[np.ndarray, Optional[float]]:
    """Robust outlier flags: |0.6745 (x - median) / MAD| > 3.5; IQR fences (3 x IQR) when MAD is 0."""
    x = pd.to_numeric(values, errors="coerce").astype(float)
    ok = x.notna()
    if ok.sum() < MIN_TRAIN_ROWS:
        return np.zeros(len(x), dtype=bool), None
    med = float(x[ok].median())
    mad = float((x[ok] - med).abs().median())
    if mad > 0:
        flag = ok & ((0.6745 * (x - med) / mad).abs() > OUTLIER_Z)
    else:
        q1, q3 = float(x[ok].quantile(0.25)), float(x[ok].quantile(0.75))
        iqr = q3 - q1
        if iqr <= 0:
            return np.zeros(len(x), dtype=bool), med
        flag = ok & ((x < q1 - 3 * iqr) | (x > q3 + 3 * iqr))
    return flag.to_numpy(dtype=bool), med


def plan_exclusions(marks: Optional[pd.Series], y: pd.Series, X: pd.DataFrame, schema: Dict[str, Any],
                    created: List[str], task: str, target_name: str,
                    cond: Optional[np.ndarray] = None, cond_text: str = "") -> Dict[str, Any]:
    """Rows to leave out of training for one target: the user's '제외' marks, the rows meeting --exclude-where
    (cond, unless the row is marked: '포함' keeps it, a written reason wins) plus automatic outliers."""
    n = len(y)
    valid = y.notna().to_numpy()
    text = (marks.map(_exclude_mark).to_numpy() if marks is not None else np.array([""] * n, dtype=object))
    forced = text == INCLUDE_MARK
    manual = (text != "") & ~forced & valid
    cond_all = np.asarray(cond, dtype=bool) if cond is not None else np.zeros(n, dtype=bool)
    by_cond = cond_all & valid & ~forced & ~manual
    cond_forced = cond_all & valid & forced
    left = int((valid & ~manual & ~by_cond).sum())
    if by_cond.any() and left < MIN_TRAIN_ROWS:
        raise InputError(f"--exclude-where 조건({cond_text})에 맞는 {int(by_cond.sum())}행과 '제외' 표시 행을 빼면 타겟 "
                         f"'{target_name}'의 학습 행이 {left}행만 남습니다 (최소 {MIN_TRAIN_ROWS}행). 조건을 사용자와 확인하십시오.")
    candidate = valid & ~manual & ~forced & ~by_cond

    check_cols = ([target_name] if task == TASK_REGRESSION else []) + \
        [c for c in schema["numeric"] if c not in created and c in X.columns]
    auto = np.zeros(n, dtype=bool)
    reasons = [""] * n
    auto_cols: List[List[str]] = [[] for _ in range(n)]
    for col in check_cols:
        series = y if col == target_name else X[col]
        flags, med = _outlier_flags(series.where(candidate))
        for i in np.flatnonzero(flags):
            auto[i] = True
            auto_cols[i].append(strip_unit(col))
            val = pd.to_numeric(pd.Series([series.iloc[i]]), errors="coerce").iloc[0]
            reasons[i] += ("; " if reasons[i] else "") + f"{strip_unit(col)} {val:.6g} (중앙값 {med:.6g})"

    note = None
    if auto.any():
        if auto.sum() > OUTLIER_MAX_FRACTION * max(candidate.sum(), 1):
            note = (f"이상치 후보가 {int(auto.sum())}행({auto.sum() / max(candidate.sum(), 1):.0%})으로 많아 자동 제외하지 않았습니다 "
                    f"(값이 넓게 퍼진 데이터일 수 있음)")
            auto[:] = False
        elif (valid & ~manual & ~by_cond & ~auto).sum() < MIN_TRAIN_ROWS:
            note = f"자동 제외하면 학습 행이 {MIN_TRAIN_ROWS}개 미만이 되어 자동 제외하지 않았습니다"
            auto[:] = False
    if task == TASK_CLASSIFICATION and auto.any():
        y_arr = y.to_numpy()
        for cls in pd.unique(y[valid]):
            in_cls = valid & (y_arr == cls)
            if (in_cls & ~manual & ~auto).sum() < 2:
                auto[in_cls] = False   # never remove a class from training
    return {"manual": manual, "auto": auto, "forced": forced, "mask": manual | auto | by_cond, "cond": by_cond,
            "cond_forced": cond_forced, "text": text, "reasons": reasons, "auto_cols": auto_cols, "note": note,
            "checked_columns": check_cols}


EXCLUDED_PREVIEW_MAX = 20    # excluded rows listed in the training JSON; excluded_rows.csv holds all


def excluded_preview_keys(rows: List[Dict[str, Any]]) -> Dict[str, Any]:
    """JSON keys of the excluded rows: excluded_preview (first EXCLUDED_PREVIEW_MAX: target, row, kind 직접/자동,
    reason, value = the target value), excluded_reasons (rows per reason: the '제외' text, or '자동 이상치: <columns>'),
    excluded_count, excluded_preview_truncated."""
    def _val(v: Any) -> Any:
        return v if isinstance(v, (int, float, str)) and not (isinstance(v, float) and v != v) else str(v)
    preview = [{"target": r["타겟"], "row": int(r["원본 행"]), "kind": "직접" if r["구분"] == "직접 표시" else
                ("조건" if r["구분"] == "조건" else "자동"), "reason": r["이유"], "value": _val(r["실제값"])}
               for r in rows[:EXCLUDED_PREVIEW_MAX]]
    reasons: Dict[str, int] = {}
    for r in rows:
        reasons[r["_key"]] = reasons.get(r["_key"], 0) + 1
    return {"excluded_count": len(rows), "excluded_preview": preview, "excluded_reasons": reasons,
            "excluded_preview_truncated": len(rows) > EXCLUDED_PREVIEW_MAX}


def score_with_excluded(task: str, pipe: Any, info: Dict[str, Any], oof_df: pd.DataFrame,
                        X: pd.DataFrame, y: pd.Series, mask: np.ndarray) -> Tuple[Optional[float], np.ndarray]:
    """All-rows score: OOF predictions of the training rows + final-model predictions of the excluded rows."""
    idx = np.flatnonzero(mask)
    if len(idx) == 0:
        return None, np.array([])
    Xe = X.iloc[idx]
    if task == TASK_REGRESSION:
        pred = pipe.predict(Xe)
        actual = np.concatenate([oof_df["actual"].to_numpy(dtype=float), y.iloc[idx].to_numpy(dtype=float)])
        score = float(r2_score(actual, np.concatenate([oof_df["predicted"].to_numpy(dtype=float), pred])))
    else:
        thr = info.get("decision_threshold")
        if thr:
            classes = list(pipe.classes_)
            pos = thr["positive_class"]
            neg = [c for c in classes if c != pos][0]
            pred = np.where(pipe.predict_proba(Xe)[:, classes.index(pos)] >= thr["threshold"], pos, neg)
        else:
            pred = pipe.predict(Xe)
        pred = np.asarray(pred).astype(str)
        actual = np.concatenate([oof_df["actual"].astype(str).to_numpy(), y.iloc[idx].astype(str).to_numpy()])
        score = float(f1_score(actual, np.concatenate([oof_df["predicted"].astype(str).to_numpy(), pred]),
                               labels=info["classes"], average="macro", zero_division=0))
    return round(score, 4), pred


def _semicolon_names(names: List[str], option: str, columns: List[str]) -> None:
    """--features '온도;압력' (names joined by ';' instead of ','): when a name with ';' is no column but every piece is one
    (exact or a unique name without its unit) -> STOP_RETRY [추천] the command with ',' (2.9.69). A column whose name
    has ';' is used as is."""
    if not any(";" in n and n not in columns for n in names):
        return
    fixed: List[str] = []
    for n in names:
        if ";" in n and n not in columns:
            pieces = [p.strip() for p in n.split(";") if p.strip()]
            if len(known_columns(pieces, columns)) != len(pieces):
                return                   # not every piece is a column: the usual [후보] of resolve_column_names
            fixed += pieces
        else:
            fixed.append(n)
    raise InputError(f"{option} '{','.join(names)}'에서 열 이름 사이를 ';'로 나눴습니다. ','로 나눕니다. "
                     f"[추천] {option_retry(','.join(fixed), option)}")


# ------------------------------------------------------------------------------
# Input diagnosis (2.9.72): a run without --features chooses its inputs from every column except the targets and the '제외'
# column, and every column it leaves out has a reason (excluded_inputs). A run with --features keeps the given inputs and
# warns about ID-like and target-leaking ones (suspect_inputs, warnings): the user decides, nothing is removed silently.
# Dates and text columns with few kinds (라인, 설비) stay inputs.
# ------------------------------------------------------------------------------
AUTO_MISSING_SHARE = 0.5       # more empty cells than this share of the rows: left out
AUTO_TEXT_UNIQUE_SHARE = 0.5   # a text column with more kinds than this share of its filled cells ...
AUTO_TEXT_MIN_KINDS = 10       # ... and at least this many kinds is a notes / memo column
AUTO_LEAK_CORR = 0.98          # |correlation| with a numeric target from which a column counts as the target itself
AUTO_LEAK_NAME_CORR = 0.9      # the same name as a target in another unit ('온도(K)' for '온도(°C)') from this one
AUTO_MIN_PAIRS = 8             # leakage is judged on at least this many rows holding both values
AUTO_ID_MIN_ROWS = 5           # a column needs this many filled cells to be called an ID
AUTO_SHOWN = 50                # names of features_auto / excluded_inputs in the JSON (all in the version folder file)
INPUT_KIND_LABEL = {"marker": "제외 표시 열", "empty": "빈 칸 많음", "constant": "상수", "leak": "타겟 누수",
                    "id": "ID·일련번호", "text": "글자 종류 많음"}
_ID_NAME_RE = re.compile(r"번호|순번|일련|로트|(?<![A-Za-z])(id|no|lot|serial)(?![A-Za-z])", re.I)


def _leak_reason(col: str, num: Optional[pd.Series], labels: pd.Series, targets: List[str],
                 tvals: Dict[str, Any], task: str) -> Optional[str]:
    """Why `col` is a target written another way (None when it is not): nearly the same numbers as a numeric target
    (another unit, a scaled copy, a value computed from it), or labels that match a class target one to one."""
    for t in targets:
        if task == TASK_REGRESSION:
            if num is None:
                continue
            x, y = num.to_numpy(dtype=float), tvals[t]
            ok = np.isfinite(x) & np.isfinite(y)
            if ok.sum() < AUTO_MIN_PAIRS or np.std(x[ok]) == 0 or np.std(y[ok]) == 0:
                continue
            r = float(np.corrcoef(x[ok], y[ok])[0, 1])
            same_name = strip_unit(col) == strip_unit(t)
            if abs(r) >= AUTO_LEAK_CORR or (same_name and abs(r) >= AUTO_LEAK_NAME_CORR):
                return (f"타겟 '{t}'과(와) " + ("이름이 같고 단위만 다르며 " if same_name else "")
                        + f"값이 거의 같음 (상관계수 {r:.3f}): 타겟을 다른 단위로 적었거나 타겟에서 계산한 열로 보여 "
                        "입력으로 쓰면 결과가 의미 없음")
        else:
            y = tvals[t]
            both = labels.notna().to_numpy() & y.notna().to_numpy()
            if both.sum() < AUTO_MIN_PAIRS:
                continue
            pairs = pd.DataFrame({"x": labels[both].astype(str).to_numpy(), "y": y[both].astype(str).to_numpy()})
            if (pairs["y"].nunique() >= 2 and pairs.groupby("x")["y"].nunique().max() == 1
                    and pairs.groupby("y")["x"].nunique().max() == 1):
                return (f"값이 타겟 '{t}'의 종류와 1:1로 맞음: 타겟을 다른 이름·코드로 적은 열로 보여 "
                        "입력으로 쓰면 결과가 의미 없음")
    return None


def _id_reason(col: str, kind: str, num: Optional[pd.Series], vals: pd.Series, forced: bool) -> Optional[str]:
    """Why `col` is an ID / serial number (None when it is not). Its filled cells are all different and it is a run
    of whole numbers counting by 1 (1001, 1002, ...), a code with digits ('S-001') or named like one ('시료번호', 'ID').
    Measured numbers that are all different (175.3, 182.1, ...) are inputs."""
    n = len(vals)
    if kind == "datetime" or n < AUTO_ID_MIN_ROWS or int(vals.nunique()) != n:
        return None
    named = bool(_ID_NAME_RE.search(strip_unit(col)))
    if kind == "numeric" and num is not None:
        v = num.dropna().to_numpy(dtype=float)
        if len(v) < AUTO_ID_MIN_ROWS:
            return None
        whole = bool(np.all(np.isclose(v, np.round(v))))
        d = np.diff(np.sort(v))
        counting = whole and len(d) > 0 and float(np.median(d)) == 1.0 and (v.max() - v.min() + 1) <= 1.5 * len(v)
        if counting or named:
            return "행마다 다른 번호 (ID·일련번호로 보임" + (", 1씩 늘어남" if counting else "") + "): 측정값이 아니라 입력으로 배울 것이 없음"
        return None
    if forced:
        return None
    text = vals.astype(str)
    code = float((text.str.contains(r"\d", regex=True) & ~text.str.contains(r"\s", regex=True)
                  & (text.str.len() <= 30)).mean())      # 'S-001', 'LOT2026-01' (a sentence with a number is a note)
    if code >= 0.9 or named:
        ex = ", ".join(repr(str(x)) for x in vals.astype(str).head(2))
        return f"행마다 다른 코드 (ID·일련번호로 보임, 예: {ex}): 새 데이터의 값은 학습에 없는 값이라 배울 것이 없음"
    return None


def diagnose_inputs(df: pd.DataFrame, columns: List[str], targets: List[str], task: str,
                    dtypes: Optional[str] = None) -> List[Dict[str, str]]:
    """[{column, kind, reason}] of the columns that should not be model inputs (kind: INPUT_KIND_LABEL keys), in
    column order. Rules, the first that holds: more than AUTO_MISSING_SHARE empty, one value (constant), a target
    written another way (_leak_reason), an ID (_id_reason), a text column whose kinds are nearly one per row
    (notes). Columns forced with --dtypes keep their type; a forced text column is never called notes or a code."""
    forced = {p.split(":", 1)[0].strip(): p.split(":", 1)[1].strip()
              for p in str(dtypes or "").split(",") if ":" in p}
    kinds_of = {"str": "categorical", "datetime": "datetime", "float": "numeric", "int": "numeric"}
    tvals: Dict[str, Any] = {}
    for t in targets:
        tvals[t] = (coerce_numeric(df[t]).to_numpy(dtype=float) if task == TASK_REGRESSION
                    else normalize_labels(df[t]))
    n_rows = len(df)
    out: List[Dict[str, str]] = []
    for c in columns:
        raw = df[c]
        if pd.api.types.is_numeric_dtype(raw) and not pd.api.types.is_bool_dtype(raw):
            labels = raw.astype(float)          # a number column: no cell-by-cell work (wide data)
        else:
            labels = normalize_labels(raw)      # blanks and placeholders ('-', 'N/A') are empty
        filled = labels.notna()
        n = int(filled.sum())

        def add(kind: str, reason: str) -> None:
            out.append({"column": c, "kind": kind, "reason": reason})
        if n_rows and n_rows - n > AUTO_MISSING_SHARE * n_rows:
            add("empty", f"빈 칸이 많음 ({n_rows}행 중 {n_rows - n}행 비어 있음, {int(AUTO_MISSING_SHARE * 100)}% 초과)")
            continue
        vals = labels[filled]
        kinds = int(vals.nunique())
        if kinds <= 1:
            add("constant", "모든 행의 값이 같음 (상수" + (f" {vals.iloc[0]!r}" if n else "") + "): 입력으로 배울 것이 없음")
            continue
        kind = kinds_of.get(forced.get(c, ""), None) or (
            "numeric" if pd.api.types.is_numeric_dtype(raw) and not pd.api.types.is_bool_dtype(raw) else column_kind(raw))
        num = coerce_numeric(raw) if kind == "numeric" else None
        leak = _leak_reason(c, num, labels, targets, tvals, task)
        if leak:
            add("leak", leak)
            continue
        ident = _id_reason(c, kind, num, vals, c in forced)
        if ident:
            add("id", ident)
            continue
        if kind == "categorical" and c not in forced and (
                kinds > MAX_CATEGORIES or (kinds >= AUTO_TEXT_MIN_KINDS and kinds > AUTO_TEXT_UNIQUE_SHARE * n)):
            add("text", f"글자 값의 종류가 {kinds}개로 거의 행마다 달라 (비고·메모 같은 열) 입력으로 배울 수 없음")
    return out


def auto_inputs(df_raw: pd.DataFrame, target_names: List[str], task: str,
                dtypes: Optional[str]) -> Tuple[List[str], List[Dict[str, str]]]:
    """A run without --features: (inputs, excluded [{column, kind, reason}]) over every column except the targets."""
    columns = [str(c) for c in df_raw.columns]
    targets, _ = resolve_column_names(target_names, columns, "타겟(--targets)", layout=dict(df_raw.attrs))
    excl_col = find_exclude_column(columns)
    excluded: List[Dict[str, str]] = []
    cands: List[str] = []
    for c in columns:
        if c in targets:
            continue
        if c == excl_col:
            excluded.append({"column": c, "kind": "marker",
                             "reason": "학습에서 뺄 행을 표시하는 열 (이유·'포함' 표시는 입력이 아님)"})
            continue
        cands.append(c)
    diag = diagnose_inputs(df_raw, cands, targets, task, dtypes)
    left = {d["column"] for d in diag}
    excluded += diag
    order = {c: i for i, c in enumerate(columns)}
    excluded.sort(key=lambda d: order.get(d["column"], 0))
    features = [c for c in cands if c not in left]
    for d in excluded:
        sys.stderr.write(f"[INFO] 입력 자동 선택: '{d['column']}' 제외 - {d['reason']}\n")
    if not features:
        raise InputError("입력(--features)을 주지 않아 도구가 입력 열을 고르려 했지만 쓸 수 있는 열이 없습니다 ("
                         + "; ".join(f"'{d['column']}': {INPUT_KIND_LABEL[d['kind']]}" for d in excluded[:10])
                         + "). 입력으로 쓸 열을 사용자에게 묻고 --features로 주십시오.",
                         extra={"excluded_inputs": [{"column": d["column"], "reason": d["reason"]} for d in excluded]})
    sys.stderr.write(f"[INFO] 입력 자동 선택: {len(features)}개 사용 ({', '.join(features[:10])}"
                     f"{' 등' if len(features) > 10 else ''}), {len(excluded)}개 제외\n")
    return features, excluded


def _prepare_group(df_raw: pd.DataFrame, group: Dict[str, Any], task: str) -> Dict[str, Any]:
    """
    One feature recipe (features, --eval formulas, --dtypes, positive class) applied to the loaded data.
    A run has one group.
    """
    feats_in, tgts_in = group["features"], group["targets"]
    # no --features (2.9.72): the tool chooses the inputs from every column except the targets (auto_inputs)
    auto = feats_in is None or (isinstance(feats_in, str) and not feats_in.strip())
    feature_names = ([] if auto else parse_feature_or_target_list(feats_in, "features") if isinstance(feats_in, str)
                     else list(feats_in or []))
    target_names = parse_feature_or_target_list(tgts_in, "targets") if isinstance(tgts_in, str) else list(tgts_in or [])
    if not feature_names and not auto:
        raise InputError("학습에 투입할 독립변수 목록(--features)이 지정되지 않았거나 파일이 비어 있습니다.")
    columns = [str(c) for c in df_raw.columns]
    for names, given, opt in ((feature_names, feats_in, "--features"), (target_names, tgts_in, "--targets")):
        if isinstance(given, str) and names:
            _semicolon_names(names, opt, columns)
    if not target_names:
        raise InputError("예측할 종속변수 목록(--targets)이 지정되지 않았거나 파일이 비어 있습니다.")
    # --dtypes names resolved like --features / --targets, types checked (the canonical list is what the ledger keeps)
    dtypes = resolve_dtypes(group.get("dtypes"), [str(c) for c in df_raw.columns], dict(df_raw.attrs))
    excluded_inputs: List[Dict[str, str]] = []
    if auto:
        feature_names, excluded_inputs = auto_inputs(df_raw, target_names, task, dtypes)
    # --exclude-where: column names resolved like --features / --targets, the canonical text is what the ledger keeps
    where_txt: Optional[str] = None
    where_hit: Optional[np.ndarray] = None
    group_warnings_pre: List[str] = []
    where_fix: Optional[Dict[str, Any]] = None
    if group.get("exclude_where"):
        conds = resolve_where(parse_where(group["exclude_where"], "--exclude-where"), [str(c) for c in df_raw.columns],
                              "--exclude-where", dict(df_raw.attrs))
        where_txt = where_text(conds)
        where_hit = where_mask(df_raw, conds)
        sys.stderr.write(f"[INFO] --exclude-where 조건({where_txt})에 맞는 행: {int(where_hit.sum())}개\n")
        if not where_hit.any():
            group_warnings_pre.append(f"--exclude-where 조건({where_txt})에 맞는 행이 없어 조건으로 뺀 행은 0개입니다")
            sys.stderr.write(f"[WARN] {group_warnings_pre[-1]}\n")
        variants = where_notation_variants(df_raw, conds)
        if variants:
            # a code column compared as text: '월 = 1' leaves the '01' rows (the same number in training) in
            fixed = where_text(where_with_variants(conds, variants))
            cmd = where_retry("--exclude-where", fixed)
            for (c, o, v), cnt in variants:
                notes = ", ".join(f"'{t}' {n}행" for t, n in cnt.items())
                group_warnings_pre.append(
                    f"--exclude-where '{c} {o} {v}': {notes}은(는) 표기가 달라 " + ("빼지 않음" if o == "=" else "함께 뺌")
                    + f" ('{c}' 열은 앞자리 0을 구분하는 코드 열이라 글자 그대로 비교). 같은 값으로 보려면 [추천] {cmd}")
                sys.stderr.write(f"[WARN] {group_warnings_pre[-1]}\n")
            if cmd.startswith("python "):
                where_fix = {"purpose": f"--exclude-where를 표기가 다른 같은 값까지 포함한 조건({fixed})으로 다시 학습 "
                                        "(사용자가 원할 때만)", "command": cmd, "needs_input": False}

    # Fail-Fast Validation & Type Sanitization
    log_progress(10, "컬럼 유효성 검증 및 타입 정제 중...")
    df_clean, feature_names, target_names, column_aliases = validate_and_sanitize_data(
        df=df_raw, feature_names=feature_names, target_names=target_names, dtypes_arg=dtypes, task=task)
    excl_col = find_exclude_column(list(df_raw.columns))
    if excl_col and excl_col in target_names:
        raise InputError(f"'{excl_col}' 열은 학습에서 뺄 행을 표시하는 열이라 타겟(--targets)으로 쓸 수 없습니다.")
    if excl_col and excl_col in feature_names:
        feature_names = [f for f in feature_names if f != excl_col]
        sys.stderr.write(f"[WARN] '{excl_col}' 열은 뺄 행 표시용이라 피처에서 제외했습니다.\n")
        if not feature_names:
            raise InputError(f"'{excl_col}' 열을 빼면 학습할 피처(--features)가 없습니다.")

    # --features given: ID-like and target-leaking inputs are said (the user decides; nothing is removed)
    suspect_inputs: List[Dict[str, str]] = []
    if not auto:
        suspect_inputs = [d for d in diagnose_inputs(df_raw, [f for f in feature_names if f in df_raw.columns],
                                                     target_names, task, dtypes) if d["kind"] in ("id", "leak")]
        for d in suspect_inputs:
            group_warnings_pre.append(f"입력 '{d['column']}': {d['reason']}. 입력으로 두었습니다 (빼려면 사용자에게 묻고 "
                                      f"--features에서 빼고 다시 학습)")
            sys.stderr.write(f"[WARN] {group_warnings_pre[-1]}\n")

    # Feature Engineering (--eval): a string from the CLI or the canonical list stored in the ledger
    applied_evals: List[str] = []
    created_features: List[str] = []
    evals = group.get("eval")
    if evals:
        log_progress(15, "방탄 AST 파생변수 연산 중 (--eval)...")
        if isinstance(evals, str):
            df_clean, applied_evals, created_features = run_safe_eval(df=df_clean, eval_str=evals, target_cols=target_names)
        else:
            df_clean, applied_evals, created_features = apply_eval_expressions(
                df_clean, list(evals), target_cols=target_names, strict=True)
        for c_feat in created_features:
            if c_feat not in feature_names:
                feature_names.append(c_feat)

    # Feature typing (numeric / categorical / datetime) and model input frame
    feature_schema = infer_feature_schema(df_clean, feature_names, created_features, dtypes_arg=dtypes)
    X_all = prepare_features(df_clean, feature_schema)
    empty_numeric = [c for c in feature_schema["numeric"] if X_all[c].notna().sum() == 0]
    if empty_numeric:
        raise InputError(
            f"컬럼 {empty_numeric}에 숫자로 변환 가능한 값이 하나도 없습니다. 헤더/데이터 시작 행(--header-row, --data-row)을 확인하십시오."
        )
    for col in feature_schema["numeric"]:
        df_clean[col] = X_all[col]
    group_warnings = group_warnings_pre + _decimal_comma_category_check(df_raw, feature_schema, feature_names,
                                                                        created_features, dtypes)
    # Columns the user must provide at prediction time: own raw features + sources of derived features
    sources = eval_source_columns(applied_evals, [str(c) for c in df_raw.columns])
    raw_required = [f for f in feature_names if f not in created_features]
    raw_required += [c for c in sources if c not in raw_required and c not in target_names]
    for c in raw_required:
        if c not in df_clean.columns and c in df_raw.columns:
            df_clean[c] = df_raw[c]
    return {"df": df_clean, "features": feature_names, "targets": target_names, "aliases": column_aliases,
            "raw_required": raw_required,
            "evals": applied_evals, "created": created_features, "schema": feature_schema, "X": X_all,
            "dtypes": dtypes, "positive_class": group.get("positive_class"), "warnings": group_warnings,
            "where": where_txt, "where_hit": where_hit, "where_fix": where_fix,
            "auto": auto, "excluded_inputs": excluded_inputs, "suspect_inputs": suspect_inputs}


DECIMAL_COMMA_STOP_SHARE = 0.5     # a selected text feature with this share of '175,6' cells is refused


def _decimal_comma_category_check(df_raw: pd.DataFrame, schema: Dict[str, Any], features: List[str],
                                  created: List[str], dtypes: Optional[str]) -> List[str]:
    """A categorical feature whose cells are numbers saved with a decimal comma ('175,6', '0,149' text) would train
    as a high-cardinality category: refused (STOP_RETRY) when half of its filled cells or more are such texts,
    a warning when the rest are numbers. Columns forced with --dtypes are the user's choice."""
    forced = {p.split(":", 1)[0].strip() for p in str(dtypes or "").split(",") if ":" in p}
    out: List[str] = []
    for c in schema["categorical"]:
        if c in created or c not in features or c in forced or c not in df_raw.columns:
            continue
        n, dc, num, ex = decimal_comma_text_share(df_raw[c])
        if not dc:
            continue
        examples = ", ".join(repr(v) for v in ex)
        if dc >= DECIMAL_COMMA_STOP_SHARE * n:
            raise InputError(
                f"피처 '{c}'의 채운 칸 {n}개 중 {dc}개({examples})가 소수점 쉼표로 쓴 숫자 글자라 숫자로 읽지 못해 "
                f"종류(범주) 값으로 학습하게 됩니다. 엑셀에서 쉼표를 점(.)으로 바꾸거나(예: 175,6 -> 175.6) 숫자 칸으로 "
                f"저장한 뒤 다시 실행하십시오. 숫자가 아닌 코드이면 --dtypes \"{c}:str\"를 붙이십시오.")
        if dc + num >= 0.9 * n:
            out.append(f"피처 '{c}'의 {dc}칸({examples})은 소수점 쉼표로 쓴 숫자 글자라 숫자로 읽지 못해 이 열을 "
                       "종류(범주) 값으로 학습했습니다. 쉼표를 점(.)으로 고치면 숫자 피처로 학습합니다.")
    return out


UNPARSED_WARN_COLUMNS = 5          # columns named in the unparsed-number warning; the rest are counted


def _numeric_input_columns(g: Dict[str, Any]) -> List[str]:
    """Data columns the model reads as numbers: numeric raw features and --eval source columns (not the
    categorical / datetime ones)."""
    schema = g["schema"]
    return [c for c in g["raw_required"] if c not in schema.get("categorical", {}) and c not in schema.get("datetime", {})]


def _unparsed_number_warnings(bad: Dict[str, Dict[str, Any]]) -> List[str]:
    """Warnings of text cells in numeric inputs of the trained rows (median-imputed), one per column (rows of
    every target together), the first UNPARSED_WARN_COLUMNS columns named and the rest counted."""
    cols = [c for c, v in bad.items() if v["rows"]]
    out = []
    for c in cols[:UNPARSED_WARN_COLUMNS]:
        ex = ", ".join(repr(v) for v in bad[c]["examples"][:3])
        out.append(f"숫자 열 '{c}'의 {len(bad[c]['rows'])}칸({ex})은 숫자로 읽지 못해 빈 값으로 보고 중앙값으로 채웠습니다. "
                   "값을 고치거나 그 행을 '제외' 열에 표시하세요.")
    if len(cols) > UNPARSED_WARN_COLUMNS:
        rest = cols[UNPARSED_WARN_COLUMNS:]
        more = f" 외 {len(rest) - 10}개 열" if len(rest) > 10 else ""
        out.append(f"숫자로 읽지 못해 중앙값으로 채운 칸이 있는 숫자 열이 {len(rest)}개 더 있습니다 "
                   f"({', '.join(repr(c) for c in rest[:10])}{more}).")
    return out


def _merge_schemas(schemas: List[Dict[str, Any]]) -> Dict[str, Any]:
    """Union of feature schemas (one schema when every target uses the same features)."""
    numeric: List[str] = []
    categorical: Dict[str, List[str]] = {}
    dates: Dict[str, List[str]] = {}
    inputs: List[str] = []
    for s in schemas:
        numeric += [c for c in s["numeric"] if c not in numeric]
        for c, v in s["categorical"].items():
            categorical.setdefault(c, v)
        for c, v in s["datetime"].items():
            dates.setdefault(c, v)
        inputs += [c for c in s["input_columns"] if c not in inputs]
    return {"numeric": numeric, "categorical": categorical, "datetime": dates, "input_columns": inputs}


def _union_frame(frames: List[pd.DataFrame]) -> pd.DataFrame:
    """Column-wise union (first occurrence of a column name wins); all frames share the loaded row index."""
    if len(frames) == 1:
        return frames[0]
    parts, seen = [], set()
    for fr in frames:
        cols = [c for c in fr.columns if c not in seen]
        seen.update(cols)
        if cols:
            parts.append(fr[cols])
    return pd.concat(parts, axis=1)


def dry_run_train(args: argparse.Namespace) -> Dict[str, Any]:
    """--dry-run: which version folder the command would write and, with --overwrite, which existing folder (models,
    report files) it would replace. Nothing is read beyond the folder list, nothing is trained or written (R8)."""
    raw_file = os.path.abspath(args.file)
    if not os.path.exists(raw_file):
        raise FileNotFoundError(f"입력 파일이 존재하지 않습니다: {raw_file}")
    data_root = os.path.join(os.path.dirname(raw_file), f"{os.path.splitext(os.path.basename(raw_file))[0]}_ml_data")
    nums = [int(i[1:]) for i in (os.listdir(data_root) if os.path.isdir(data_root) else [])
            if re.match(r"^v\d+$", i) and os.path.isdir(os.path.join(data_root, i))]
    version = normalize_version(args.version) if args.version else f"v{max(nums, default=0) + 1}"
    version_dir = os.path.join(data_root, version)
    exists = os.path.isdir(version_dir)
    argv = [a for a in sys.argv if a != "--dry-run"]
    cmd = command_text(argv=argv)
    fields: Dict[str, Any] = {"status": "ok", "verified": True, "dry_run": True, "version": version,
                              "version_dir": version_dir, "replaces": bool(exists and args.overwrite)}
    if exists and not args.overwrite:
        raise InputError(f"버전 폴더 '{version}'이(가) 이미 있어 덮어쓰지 않습니다. 기존 {version}를 바꾸는 것은 사용자가 요청한 경우에만 "
                         f"--overwrite로 합니다. [추천] {option_retry(f'v{max(nums, default=0) + 1}', '--version', '--version')}")
    if exists:
        files = sorted(os.listdir(version_dir))
        fields["replaced_files"] = files[:20]
        fields["replaced_count"] = len(files)
        fields["truncated"] = len(files) > 20
        what = f"기존 {version} 폴더(파일 {len(files)}개: 모델·결과)를 이번 학습 결과로 바꿈"
    else:
        what = f"새 버전 {version}를 만듦 (기존 폴더는 바꾸지 않음)"
    fields["preview"] = what
    if cmd:
        fields["confirm_commands"] = [{"command": cmd, "destructive": bool(exists), "description": what}]
    fields["agent_directive"] = ("VERIFIED: 미리보기만 했습니다 (학습·저장 안 함). preview" + (", replaced_files" if exists else "")
                                 + "를 사용자에게 보여 주고, 사용자가 '예'라고 한 뒤에만 confirm_commands[0]을 1번 실행하세요.")
    return fields


def execute_ml_train(args: argparse.Namespace) -> Dict[str, Any]:
    """Orchestrates 1-Turn declarative regression / classification pipeline."""
    task = (getattr(args, "task", None) or TASK_REGRESSION).lower()
    positive_class = getattr(args, "positive_class", None)
    if positive_class is not None:
        if task != TASK_CLASSIFICATION:
            raise InputError("--positive-class는 --task classification과 함께 사용해야 합니다.")
        positive_class = normalize_labels(pd.Series([positive_class])).iloc[0]
    if getattr(args, "dry_run", False):
        return dry_run_train(args)
    group = {"features": args.features, "targets": args.targets, "eval": args.eval,
             "dtypes": args.dtypes, "positive_class": positive_class,
             "exclude_where": getattr(args, "exclude_where", None)}
    return train_version(args, task, [group])


# ------------------------------------------------------------------------------
# Split execution (R10): the agent's command tool stops a command at 120 s. Every target whose 5-model competition
# finished is saved in <data_root>/.resume_train/<key>.joblib (key = the exact training input: data, schema, task,
# seed, tool and scikit-learn versions; in a fallback place when that folder cannot be written: resume_save).
# Before the next target the clock is checked (check_time); a run that would
# pass the limit stops with CONTINUE, leaving no version folder and no ledger record. The same command run again
# reuses the saved targets and goes on; the version folder and history.jsonl are written once, by the run that ends.
# ------------------------------------------------------------------------------
RESUME_DIR = ".resume_train"

# Seconds of every target trained in this command: the clock check before a new target uses them.
RUN_FIT_SECONDS: List[float] = []


def _resume_key(task: str, target: str, X: pd.DataFrame, y: pd.Series, schema: Dict[str, Any], seed: int,
                extra: Dict[str, Any]) -> str:
    import hashlib
    h = hashlib.sha1()
    for part in (VERSION, sklearn.__version__, task, target, dumps_json(schema), seed, dumps_json(extra)):
        h.update(str(part).encode("utf-8") + b"\x1e")
    h.update(pd.util.hash_pandas_object(X, index=False).values.tobytes())
    h.update(pd.util.hash_pandas_object(y.astype(object), index=False).values.tobytes())
    return h.hexdigest()[:16]                 # short file names: room under MAX_PATH (260) in deep folders


def _resume_load(want: str) -> Optional[Tuple[Any, Any, Any, Any]]:
    path = resume_find(os.path.dirname(want), os.path.basename(want))
    if not path:
        return None
    try:
        res = joblib.load(path)
        if isinstance(res, tuple) and len(res) == 4:
            return res
    except Exception:
        pass
    try:                                       # damaged saved target: drop it and train again (JIT: rebuilt)
        os.remove(path)
    except OSError:
        pass
    sys.stderr.write(f"[WARN] 저장된 중간 학습 결과 {os.path.basename(path)}를 읽지 못해 다시 학습합니다\n")
    return None


def _resume_save(want: str, res: Tuple[Any, ...]) -> bool:
    """Saves a finished step for the next run of the same command (at `want`, or in a fallback place: resume_save)
    and checks the file is there. False when no place worked (check_time then stops with STOP_RETRY)."""
    return resume_save(os.path.dirname(want), os.path.basename(want), lambda tmp: joblib.dump(res, tmp),
                       "중간 학습 결과") is not None


def _resume_cleanup(resume_dir: str, used: List[str]) -> None:
    """Removes the saved targets of this (finished) run and those older than RESUME_MAX_AGE, in the original folder
    and in the fallback places."""
    now = time.time()
    keys = {os.path.basename(p).split(".")[0] for p in used}
    # a finished target and its saved candidates (<key>.joblib, <key>.m<k>.joblib), or an old one
    resume_clear(resume_dir, lambda name, p: name.split(".")[0] in keys or now - os.path.getmtime(p) > RESUME_MAX_AGE)


# The preparation of a run (everything before the first model: reading the file, checks, feature typing,
# correlation, row exclusions, the resume key of each target) is one step of its own: on large data it takes tens of
# seconds, so a run continued after CONTINUE that did it again would have no time left for a model (2.9.67). When it
# took at least PREP_SAVE_SECONDS under a time limit it is saved whole in <resume_dir>/<key>.prep.joblib (key = the
# data file's bytes, the training arguments that shape the data, tool and library versions) and the next run of the
# same command uses it. The saved objects are the same objects (pickle), so a split run trains exactly what one run does.
PREP_SAVE_SECONDS = 5.0
PREP_LABEL = "학습 준비(데이터 읽기, 검증, 상관행렬)"
RUN_PREP_SECONDS: List[float] = []        # a preparation saved in this command (a step: _step_done)
STEPS_PER_TARGET = 5 * 5 + 2              # progress estimate: candidates x folds + final fit + importance
TARGETS_AFTER: List[int] = [0]            # targets after the one being trained (the steps left in progress)


def _prep_key(args: argparse.Namespace, task: str, groups: List[Dict[str, Any]], raw_file: str) -> str:
    import hashlib

    def spec(v: Any) -> Any:                 # a list given as a file (--features, --targets, --eval, ...): its content
        if isinstance(v, str) and len(v) < 1024 and os.path.isfile(v):
            with open(v, "rb") as f:
                return ["file", hashlib.sha1(f.read()).hexdigest()]
        return v

    h = hashlib.sha1()
    for part in (VERSION, sklearn.__version__, pd.__version__, np.__version__, task, args.header_row, args.unit_row,
                 args.data_row, args.seed, dumps_json([{k: spec(v) for k, v in sorted(g.items())} for g in groups])):
        h.update(str(part).encode("utf-8") + b"\x1e")
    with open(raw_file, "rb") as f:
        for chunk in iter(lambda: f.read(4 * 1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()[:16]


def _prep_load(want: str) -> Optional[Dict[str, Any]]:
    path = resume_find(os.path.dirname(want), os.path.basename(want))
    if not path:
        return None
    try:
        res = joblib.load(path)
        if isinstance(res, dict) and res.get("version") == VERSION:
            return res
    except Exception:
        pass
    try:                                       # damaged: prepared again (JIT: rebuilt)
        os.remove(path)
    except OSError:
        pass
    sys.stderr.write(f"[WARN] 저장된 {PREP_LABEL} 결과 {os.path.basename(path)}를 읽지 못해 다시 준비합니다\n")
    return None


def _prepare_run(args: argparse.Namespace, task: str, groups: List[Dict[str, Any]], raw_file: str) -> Dict[str, Any]:
    """The preparation of train_version (see PREP_SAVE_SECONDS): a dict of everything training and packaging use."""
    log_progress(5, f"입력 데이터 로드 중: {os.path.basename(raw_file)}...")
    df_raw = load_data_file(   # the shared loader: identical parsing in ml_predict.py
        file_path=raw_file,
        header_row=args.header_row,
        unit_row=args.unit_row,
        data_row=args.data_row
    )

    prepared = [_prepare_group(df_raw, g, task) for g in groups]
    excl_col = find_exclude_column(list(df_raw.columns))
    source_rows = df_raw.attrs.get("source_rows") or list(range(1, len(df_raw) + 1))
    column_aliases: Dict[str, str] = {}
    target_group: Dict[str, Dict[str, Any]] = {}
    target_names: List[str] = []
    for g in prepared:
        column_aliases.update(g["aliases"])
        for t in g["targets"]:
            if t in target_group:
                raise InputError(f"타겟 '{t}'이(가) 학습 조건 두 곳에 들어 있습니다.")
            target_group[t] = g
            target_names.append(t)
    # Version-wide views: identical to the group's own values when there is one group
    g0 = prepared[0]
    feature_names = list(dict.fromkeys(f for g in prepared for f in g["features"]))
    applied_evals = list(dict.fromkeys(e for g in prepared for e in g["evals"]))
    feature_schema = _merge_schemas([g["schema"] for g in prepared])
    X_all = _union_frame([g["X"] for g in prepared])
    df_clean = _union_frame([g["df"][[c for c in g["features"] + g["targets"] if c in g["df"].columns]]
                             for g in prepared]) if len(prepared) > 1 else g0["df"]

    # Correlation (numeric inputs incl. date parts + numeric targets) for SSOT reporting
    log_progress(20, "변수 간 2차원 상관행렬 계산 중...")
    target_numeric = target_names if task == TASK_REGRESSION else []
    corr_parts = []
    for g in prepared:
        num_inputs = [c for c in g["schema"]["input_columns"] if c not in g["schema"]["categorical"]]
        corr_parts.append(g["X"][num_inputs])
    corr_parts.append(pd.concat([target_group[t]["df"][[t]] for t in target_numeric], axis=1)
                      if target_numeric else pd.DataFrame(index=df_raw.index))
    corr_frame = _union_frame(corr_parts) if len(prepared) > 1 else pd.concat(corr_parts, axis=1)
    corr_full = corr_frame.corr().round(4)

    # rows of the file that are not ordinary data rows: repeated header rows (dropped by the loader),
    # totals / average rows (kept: the '제외' column is the user's decision)
    train_warnings: List[str] = layout_warnings(df_raw, target_names)
    for g in prepared:
        train_warnings.extend(w for w in g["warnings"] if w not in train_warnings)
    # numbers written with '%' ('64.4%' of an Excel CSV export) are read as 64.4: said, never silent; '%' cells
    # next to plain numbers in one column stop (no scale is sure)
    pct_cols = ([c for g in prepared for c in _numeric_input_columns(g)]
                + (target_names if task == TASK_REGRESSION else []))
    # (on the rows training uses: a row the user marked in '제외' never stops the run)
    marked = (df_raw[excl_col].map(_exclude_mark).map(lambda t: bool(t) and t != INCLUDE_MARK).to_numpy(dtype=bool)
              if excl_col else np.zeros(len(df_raw), dtype=bool))
    check_percent_mix(df_raw[~marked], pct_cols)
    train_warnings.extend(percent_warning(df_raw, pct_cols))
    for w in train_warnings:
        sys.stderr.write(f"[WARN] {w}\n")
    # text cells of numeric inputs in the rows a model learns from (median-imputed): not the unit row below the
    # header (target empty), rows marked in '제외' or rows without a target
    unparsed: Dict[str, Dict[str, Any]] = {}
    unit_below = df_raw.attrs.get("unit_row_below")
    unit_pos = [i for i, r in enumerate(source_rows) if r == unit_below] if unit_below else []
    targets: Dict[str, Dict[str, Any]] = {}
    for target_name in target_names:
        g = target_group[target_name]
        extra = {"positive_class": g["positive_class"]} if task == TASK_CLASSIFICATION else {}
        y_full = g["df"][target_name]
        plan = plan_exclusions(g["df"][excl_col] if excl_col else None, y_full, g["X"], g["schema"],
                               g["created"], task, target_name, cond=g["where_hit"], cond_text=g["where"] or "")
        trained = np.array(y_full.notna(), dtype=bool) & ~np.asarray(plan["mask"], dtype=bool)
        trained[unit_pos] = False
        for c in _numeric_input_columns(g):
            if c in df_raw.columns:
                pos, ex = unparsed_number_cells(df_raw[c], trained)
                entry = unparsed.setdefault(c, {"rows": set(), "examples": []})
                entry["rows"].update(pos)
                entry["examples"] += [v for v in ex if v not in entry["examples"]]
        y_train = y_full.mask(plan["mask"])
        targets[target_name] = {"plan": plan, "key": _resume_key(task, target_name, g["X"], y_train, g["schema"],
                                                                 args.seed, extra)}
    return {"version": VERSION, "prepared": prepared, "target_names": target_names, "excl_col": excl_col,
            "source_rows": source_rows, "column_aliases": column_aliases, "feature_names": feature_names,
            "applied_evals": applied_evals, "feature_schema": feature_schema, "X_all": X_all, "df_clean": df_clean,
            "target_numeric": target_numeric, "corr_full": corr_full, "warnings": train_warnings,
            "unparsed": unparsed, "targets": targets,
            "raw_info": {"rows": int(len(df_raw)), "columns": int(df_raw.shape[1]),
                         "decimal_comma": bool(df_raw.attrs.get("decimal_comma"))}}


def train_version(args: argparse.Namespace, task: str, groups: List[Dict[str, Any]]) -> Dict[str, Any]:
    """
    Trains every target of every group into ONE new version folder.
    One group: all --targets targets share the --features features.
    """
    start_time = time.time()
    primary_metric = PRIMARY_METRIC[task]
    metric_label = METRIC_LABEL[task]
    raw_file = os.path.abspath(args.file)
    if not os.path.exists(raw_file):
        raise FileNotFoundError(f"입력 파일이 존재하지 않습니다: {raw_file}")
    warn_long_path(raw_file)

    file_dir = os.path.dirname(raw_file)
    file_stem = os.path.splitext(os.path.basename(raw_file))[0]
    data_root = os.path.join(file_dir, f"{file_stem}_ml_data")
    ledger_path = os.path.join(data_root, "history.jsonl")

    # One training run per dataset at a time (OS lock is released automatically if the process dies)
    run_lock = RunLock(data_root)
    run_lock.acquire()

    # Transaction Guard Setup
    guard = TransactionGuard(data_root=data_root, target_version=args.version,
                             overwrite=bool(getattr(args, "overwrite", False)))
    try:
        current_version = guard.setup()
    except BaseException:
        run_lock.release()
        raise

    try:
        resume_dir = os.path.join(data_root, RESUME_DIR)
        resume_used: List[str] = []
        # Everything before the first model (reading, checks, typing, correlation, exclusions, keys): saved whole when
        # it took long, so a run continued after CONTINUE does not pay it again (R10)
        prep_path = os.path.join(resume_dir, _prep_key(args, task, groups, raw_file) + ".prep.joblib")
        P = _prep_load(prep_path)
        if P is None:
            step_begin(resume_dir, os.path.basename(prep_path)[:-len(".joblib")], PREP_LABEL)
            t_prep = time.monotonic()
            P = _prepare_run(args, task, groups, raw_file)
            prep_seconds = time.monotonic() - t_prep
            if time_limit() != float("inf") and prep_seconds >= PREP_SAVE_SECONDS and _resume_save(prep_path, P):
                RUN_PREP_SECONDS.append(prep_seconds)
            step_end()
            if RUN_PREP_SECONDS:            # a step of this run: nothing else may fit before the limit
                check_time(0.0, f"학습: {PREP_LABEL} 저장됨 (첫 학습 단계 전에 멈춤, 약 "
                                f"{len(P['target_names']) * STEPS_PER_TARGET}단계 남음)")
        else:
            sys.stderr.write(f"[INFO] 이전 실행에서 저장한 {PREP_LABEL} 결과를 씀 (나눠서 진행)\n")
            for w in P["warnings"]:
                sys.stderr.write(f"[WARN] {w}\n")
        resume_used.append(prep_path)
        prepared, target_names = P["prepared"], P["target_names"]
        excl_col, source_rows, column_aliases = P["excl_col"], P["source_rows"], P["column_aliases"]
        feature_names, applied_evals, feature_schema = P["feature_names"], P["applied_evals"], P["feature_schema"]
        X_all, df_clean, target_numeric, corr_full = P["X_all"], P["df_clean"], P["target_numeric"], P["corr_full"]
        unparsed, raw_info = P["unparsed"], P["raw_info"]
        target_group = {t: g for g in prepared for t in g["targets"]}
        g0 = prepared[0]

        # Multi-Target 5-Model Competitive Training (one task per run)
        target_benchmarks = {}
        target_champions = {}
        target_summaries = {}
        all_champion_scores = []
        exclusions: Dict[str, Dict[str, Any]] = {}
        excluded_rows_out: List[Dict[str, Any]] = []
        train_warnings: List[str] = list(P["warnings"])
        benchmark_fn = train_and_benchmark_models if task == TASK_REGRESSION else train_and_benchmark_classifiers

        total_targets = len(target_names)
        reused_targets: List[str] = []
        for t_idx, target_name in enumerate(target_names):
            g = target_group[target_name]
            t_offset = 25 + int(55 * t_idx / total_targets)
            t_weight = int(55 / total_targets)
            TARGETS_AFTER[0] = total_targets - t_idx - 1

            log_progress(t_offset, f"타겟 [{t_idx+1}/{total_targets}] '{target_name}' 5대 모델 경쟁 학습 ({task})...")

            extra = {"positive_class": g["positive_class"]} if task == TASK_CLASSIFICATION else {}
            y_full = g["df"][target_name]
            plan = P["targets"][target_name]["plan"]
            if plan["mask"].any():
                sys.stderr.write(f"[INFO] 타겟 '{target_name}': 학습에서 뺀 행 {int(plan['mask'].sum())}개 "
                                 f"(직접 표시 {int(plan['manual'].sum())}, 조건 {int(plan['cond'].sum())}, "
                                 f"자동 이상치 {int(plan['auto'].sum())})\n")
            if plan["cond_forced"].any():
                rows_kept = [source_rows[i] if i < len(source_rows) else i + 1 for i in np.flatnonzero(plan["cond_forced"])]
                w = (f"타겟 '{target_name}': '{excl_col}' 열에 '{INCLUDE_MARK}'으로 표시한 {len(rows_kept)}행(원본 행 "
                     f"{', '.join(map(str, rows_kept[:10]))}{' 등' if len(rows_kept) > 10 else ''})은 --exclude-where 조건에 맞지만 "
                     f"학습에 넣었습니다 ('{INCLUDE_MARK}' 표시가 우선)")
                if w not in train_warnings:
                    train_warnings.append(w)
                    sys.stderr.write(f"[WARN] {w}\n")
            if plan["note"]:
                sys.stderr.write(f"[WARN] 타겟 '{target_name}': {plan['note']}\n")
            y_train = y_full.mask(plan["mask"])
            r_path = os.path.join(resume_dir, P["targets"][target_name]["key"] + ".joblib")
            cached = _resume_load(r_path)
            if cached is not None:
                t_bench, champion_pipe, champion_info, oof_df = cached
                reused_targets.append(target_name)
                sys.stderr.write(f"[INFO] 타겟 '{target_name}': 이전 실행에서 저장한 학습 결과를 씀 (나눠서 진행)\n")
            else:
                if RUN_FIT_SECONDS:     # at least one new target per run: running again always moves on
                    check_time(max(RUN_FIT_SECONDS) * 1.2,
                               f"학습: 타겟 {t_idx}/{total_targets}개 끝남, 저장됨 ('{target_name}' 학습 전에 멈춤, "
                               f"약 {(total_targets - t_idx) * STEPS_PER_TARGET}단계 남음)")
                t_fit = time.monotonic()
                t_bench, champion_pipe, champion_info, oof_df = benchmark_fn(
                    X=g["X"],
                    y_series=y_train,
                    target_name=target_name,
                    schema=g["schema"],
                    seed=args.seed,
                    log_offset=t_offset,
                    log_weight=t_weight,
                    resume=r_path[:-len(".joblib")],
                    **extra
                )
                RUN_FIT_SECONDS.append(time.monotonic() - t_fit)
                # the next target, or the packaging of a run near the limit, may be stopped: keep this target
                if total_targets > 1 or time_would_pass(0.0):
                    _resume_save(r_path, (t_bench, champion_pipe, champion_info, oof_df))
                step_end()
            resume_used.append(r_path)

            if champion_info.get("duplicate_rows"):
                train_warnings.append(duplicate_warning(target_name, champion_info["duplicate_rows"],
                                                        t_bench["sample_count"], champion_info.get("duplicate_note")))
                sys.stderr.write(f"[WARN] {train_warnings[-1]}\n")
            all_score, excl_pred = score_with_excluded(task, champion_pipe, champion_info, oof_df,
                                                       g["X"], y_full, plan["mask"])
            excl_info: Dict[str, Any] = {
                "column": excl_col,
                "manual": int(plan["manual"].sum()),
                "auto": int(plan["auto"].sum()),
                "forced_include": int(plan["forced"].sum()),
                "auto_note": plan["note"],
                "all_rows_score": all_score,
            }
            if g["where"]:
                excl_info["condition"] = int(plan["cond"].sum())
                excl_info["condition_text"] = g["where"]
            champion_info["excluded"] = excl_info
            if all_score is not None:
                champion_info["all_rows_score"] = all_score
            exclusions[target_name] = {**excl_info, "mask": plan["mask"]}
            for k, i in enumerate(np.flatnonzero(plan["mask"])):
                actual = y_full.iloc[i]
                excluded_rows_out.append({
                    "타겟": target_name,
                    "원본 행": source_rows[i] if i < len(source_rows) else i + 1,
                    "구분": "직접 표시" if plan["manual"][i] else ("조건" if plan["cond"][i] else "자동 이상치"),
                    "이유": (plan["text"][i] if plan["manual"][i]
                           else (f"조건: {g['where']}" if plan["cond"][i] else plan["reasons"][i])),
                    "_key": (plan["text"][i] if plan["manual"][i]
                             else (f"조건: {g['where']}" if plan["cond"][i]
                                   else "자동 이상치: " + ", ".join(plan["auto_cols"][i]))),
                    "실제값": actual if task == TASK_CLASSIFICATION else round(float(actual), 6),
                    "예측값": (str(excl_pred[k]) if task == TASK_CLASSIFICATION else round(float(excl_pred[k]), 6)),
                })

            target_benchmarks[target_name] = t_bench
            target_champions[target_name] = {
                "pipeline": champion_pipe,
                "info": champion_info,
                "oof": oof_df
            }
            all_champion_scores.append(champion_info["val_score"])

            # Target statistics
            y_valid = y_full[~plan["mask"]].dropna()
            if task == TASK_REGRESSION:
                target_summaries[target_name] = {
                    "mean": round(float(y_valid.mean()), 4),
                    "std": round(float(y_valid.std()), 4),
                    "min": round(float(y_valid.min()), 4),
                    "max": round(float(y_valid.max()), 4)
                }
            else:
                target_summaries[target_name] = {
                    "classes": champion_info["classes"],
                    "class_counts": champion_info["class_counts"],
                    "missing_dropped": int(g["df"][target_name].isna().sum())
                }

        mean_val_score = float(np.mean(all_champion_scores))

        # Check Previous History (same task only) for Delta
        log_progress(85, f"이전 이력 대조 및 성능 변화(Delta {metric_label}) 산출 중...")
        # the record of a version replaced with --overwrite is no earlier run to compare with
        history_records = [h for h in load_history_ledger(ledger_path) if h.get("version") != current_version]

        # Resolve project description: explicit arg takes precedence, otherwise inherit from previous version
        project_desc = (getattr(args, "description", None) or "").strip()
        if project_desc and os.path.isfile(project_desc):
            project_desc = read_text_file(project_desc).strip()

        if not project_desc and history_records:
            project_desc = history_records[-1].get("description", "")

        # Scores are comparable only for the same task AND the same target set
        same_task = [h for h in history_records if _same_experiment(h, task, target_names)]

        def _rec_score(h: Dict[str, Any]) -> Optional[float]:
            v = h.get("mean_val_score")
            return float(v) if isinstance(v, (int, float)) else None

        prev_scores = [v for v in (_rec_score(h) for h in same_task) if v is not None]
        last_score = _rec_score(same_task[-1]) if same_task else None
        delta_score: Any = round(mean_val_score - last_score, 4) if last_score is not None else "base"
        is_best = mean_val_score > max(prev_scores, default=-999.0)

        # Generate Actionable Recommendations for 26B-30B LLM.
        # The goal is judged on the WEAKEST target (same rule as the report verdict), not on the mean.
        min_score = min(all_champion_scores)
        primary_target = min(target_names, key=lambda t: target_champions[t]["info"]["val_score"])
        prim_champ = target_champions[primary_target]["info"]
        top_feats = prim_champ.get("top_features", [])
        # A feature is only "weak" if it is weak for EVERY target (multi-target runs share one feature list)
        weak_sets = [set(target_champions[t]["info"].get("weak_features", [])) for t in target_names]
        weak_feats = [f for f in prim_champ.get("weak_features", []) if all(f in ws for ws in weak_sets)]
        below = [t for t in target_names if target_champions[t]["info"]["val_score"] < 0.90]

        recommendations = []
        if min_score >= 0.90:
            recommendations.append(
                f"모든 타겟이 목표 성능({metric_label} >= 0.90)을 달성했습니다. 예측 배포와 보고서 작성이 가능합니다."
                + (" 보고서를 만들면 엑셀만으로 예측하는 성능예측표(엑셀 계산기)도 함께 만들어집니다." if task == TASK_REGRESSION else ""))
        else:
            if len(target_names) > 1:
                recommendations.append(f"목표({metric_label} >= 0.90) 미달 타겟: {below}. 아래 권고는 가장 낮은 '{primary_target}' 기준입니다.")
            weak_classes = [c for c, m in prim_champ.get("per_class", {}).items() if m["recall"] < 0.70]
            if task == TASK_CLASSIFICATION and weak_classes:
                recommendations.append(
                    f"재현율(Recall)이 0.70 미만인 클래스 {weak_classes[:3]}의 구분력이 부족합니다. "
                    f"해당 클래스를 가르는 피처를 추가하거나 --eval로 비율/차이 파생변수를 만드십시오."
                )
            elif len(top_feats) >= 2:
                f1, f2 = top_feats[0]["feature"], top_feats[1]["feature"]
                recommendations.append(f"주요 기여 피처 '{f1}'와(과) '{f2}'의 교차 상호작용('{f1} * {f2}') 생성을 권장합니다.")
        if task == TASK_CLASSIFICATION and prim_champ.get("class_weight"):
            recommendations.append("클래스 불균형으로 class_weight='balanced'가 적용되었습니다. 소수 클래스 데이터 추가 수집을 권장합니다.")
        if weak_feats:
            recommendations.append(f"기여도가 1% 미만인 피처 {weak_feats[:2]}을(를) 배제하여 노이즈를 감소시키십시오.")

        # Task-specific score keys (regression keeps its historical key names)
        score_keys: Dict[str, Any] = {
            "task": task,
            "primary_metric": primary_metric,
            "mean_val_score": round(mean_val_score, 4),
            "delta_score": delta_score,
        }

        # Stage Artifacts in .tmp_vX/
        log_progress(90, f"버전 산출물 패키징 중 ({current_version})...")
        staging_dir = guard.tmp_dir

        # 1. Unified leaderboard.json (Single file covering all targets)
        leaderboard_data = {
            "version": current_version,
            "timestamp": datetime.now().isoformat(),
            **score_keys,
            "is_best": is_best,
            "targets": target_benchmarks
        }
        with open(os.path.join(staging_dir, "leaderboard.json"), "w", encoding="utf-8") as f:
            f.write(dumps_json(leaderboard_data, indent=2))

        def _datetime_stats(g: Dict[str, Any], rows: np.ndarray) -> Dict[str, Any]:
            out = {}
            for col in g["schema"]["datetime"]:
                dt = parse_datetime(g["df"][col][rows]).dropna()
                if not dt.empty:
                    out[col] = {"min": dt.min().isoformat(), "max": dt.max().isoformat()}
            return out

        # Correlation: full matrix as a file, a bounded matrix in the ledger (keeps history.json small)
        corr_full.to_csv(os.path.join(staging_dir, "correlation.csv"), encoding="utf-8-sig")
        corr_matrix, corr_truncated = bounded_correlation(
            corr_full, target_champions[target_names[0]]["info"]["feature_importances"],
            feature_schema, target_numeric)
        target_correlations = {
            t: {c: v for c, v in corr_full[t].drop(labels=target_numeric).items()} for t in target_numeric
        }

        # 2. Target Models and Metadata
        for target_name, champ_data in target_champions.items():
            g = target_group[target_name]
            pipe = champ_data["pipeline"]
            info = champ_data["info"]
            safe_target = sanitize_filename(target_name)

            # Save Model Joblib with safe filename
            joblib_path = os.path.join(staging_dir, f"model_{safe_target}.joblib")
            joblib.dump(pipe, joblib_path, compress=3)

            # Out-of-fold predictions of the champion (report: actual vs predicted, residuals, ROC)
            champ_data["oof"].to_csv(os.path.join(staging_dir, f"oof_{safe_target}.csv"), index=False, encoding="utf-8-sig")

            # Feature statistics for inference validation (numeric extrapolation guard)
            feat_stats = {}
            used_rows = ~exclusions[target_name]["mask"]   # training range = rows the model learned from
            source_numeric = [c for c in g["raw_required"] if c not in g["schema"]["numeric"]
                              and c not in g["schema"].get("categorical", {}) and c not in g["schema"].get("datetime", {})]
            for fname in list(g["schema"]["numeric"]) + source_numeric:
                s = coerce_numeric(g["df"][fname])[used_rows].dropna()
                if not s.empty:
                    # significant digits, not decimals: an input of 3e-06 must not be stored as 0
                    feat_stats[fname] = {k: float(f"{float(v):.12g}") for k, v in (
                        ("mean", s.mean()), ("std", s.std()), ("min", s.min()), ("max", s.max()), ("median", s.median()))}

            # Inference Metadata JSON with safe filename
            meta_json = {
                "version": current_version,
                "target": target_name,
                "target_safe_name": safe_target,
                "task": task,
                "primary_metric": primary_metric,
                "val_score": info["val_score"],
                "description": project_desc,
                "champion_model": info["champion_model"],
            }
            if task == TASK_REGRESSION:
                meta_json.update({
                    "val_r2": info["val_r2"],
                    "mae": info["mae"],
                    "rmse": info["rmse"],
                    "residual_skew": info.get("residual_skew", 0.0),
                    "prediction_interval": info.get("prediction_interval"),
                })
            else:
                meta_json.update({
                    "val_f1_macro": info["val_f1_macro"],
                    "accuracy": info["accuracy"],
                    "balanced_accuracy": info["balanced_accuracy"],
                    "roc_auc": info["roc_auc"],
                    "classes": info["classes"],
                    "class_counts": info["class_counts"],
                    "decision_threshold": info.get("decision_threshold"),
                })
            excl = {k: v for k, v in exclusions[target_name].items() if k != "mask"}
            if excl["manual"] or excl["auto"] or excl["forced_include"] or excl["auto_note"]:
                meta_json["excluded"] = excl
            meta_json.update({
                "features_required": g["features"],
                "input_columns_required": g["raw_required"],
                "feature_count": len(g["features"]),
                "feature_stats": feat_stats,
                "datetime_stats": _datetime_stats(g, used_rows),
                "feature_schema": g["schema"],
                "decimal_comma": raw_info["decimal_comma"],   # prediction CSVs are read the same way
                "derived_evals": g["evals"],
                "target_summary": target_summaries.get(target_name, {}),
                "sklearn_version": sklearn.__version__,
                "python_version": sys.version.split()[0]
            })
            with open(os.path.join(staging_dir, f"model_{safe_target}.json"), "w", encoding="utf-8") as f:
                f.write(dumps_json(meta_json, indent=2))

        # 3. Clean Dataset Snapshot (model_dataset.csv)
        raw_extra = [c for g in prepared for c in g["raw_required"] if c not in feature_names]
        dataset_snapshot_cols = list(dict.fromkeys(c for c in feature_names + raw_extra + target_names if c in df_clean.columns))
        snapshot_df = df_clean[dataset_snapshot_cols].copy()
        # Excel row numbers and per-target training exclusions, so ml_predict.py can compare its predictions
        # with the measured rows the model learned from (evidence rows are shown as Excel row numbers)
        snapshot_df[SNAPSHOT_ROW_COL] = [source_rows[i] if i < len(source_rows) else i + 1 for i in range(len(snapshot_df))]
        for t in target_names:
            snapshot_df[f"{SNAPSHOT_EXCL_PREFIX}{t}"] = np.where(exclusions[t]["mask"], 1, 0)
        snapshot_path = os.path.join(staging_dir, "model_dataset.csv")

        # Windows 한국어 Excel 호환: cp949 기본, 표현 불가 문자 포함 시 utf-8-sig 자동 폴백
        write_csv_excel_compatible(snapshot_df, snapshot_path)

        # 3-B. Rows left out of training (user marks + automatic outliers), for the user to review in Excel
        excluded_rows_file = None
        if excluded_rows_out:
            ex_path = write_csv_excel_compatible(pd.DataFrame(excluded_rows_out).drop(columns="_key"),
                                                 os.path.join(staging_dir, "excluded_rows.csv"))
            excluded_rows_file = os.path.join(guard.final_dir, os.path.basename(ex_path))
        exclusion_summary = {t: {k: v for k, v in e.items() if k != "mask"} for t, e in exclusions.items()
                             if e["manual"] or e["auto"] or e["forced_include"] or e["auto_note"] or e.get("condition")}

        # 4. Append Record to history.jsonl (SSOT)
        duration_sec = round(time.time() - start_time, 2)
        # Missing feature cells (incl. unparseable strings) that SimpleImputer(median) fills
        missing_imputed = int(X_all.isna().sum().sum())
        prescription = recommendations[0] if recommendations else "모델 경쟁 학습 및 평가 완료"

        owner = input_column_owner(feature_schema)
        feature_missing: Dict[str, int] = {}
        for col, n_miss in X_all.isna().sum().items():
            raw_col = owner.get(col, col)
            feature_missing[raw_col] = max(feature_missing.get(raw_col, 0), int(n_miss))
        for w in _unparsed_number_warnings(unparsed):
            train_warnings.append(w)
            sys.stderr.write(f"[WARN] {w}\n")
        data_profile = {
            "source_file": os.path.basename(raw_file),
            "rows_loaded": raw_info["rows"],
            "columns_loaded": raw_info["columns"],
            "rows_used": {t: int(target_benchmarks[t]["sample_count"]) for t in target_names},
            "target_missing_rows": {t: int(target_group[t]["df"][t].isna().sum()) for t in target_names},
            "feature_missing": feature_missing,
            "missing_imputed_cells": missing_imputed,
            "layout": {"header_row": args.header_row, "unit_row": args.unit_row, "data_row": args.data_row},
        }

        run_tag = getattr(args, "tag", None) or ""
        history_record = {
            "version": current_version,
            "timestamp": datetime.now().isoformat(),
            "tag": run_tag,
            "description": project_desc,
            "duration_sec": duration_sec,
            "data_hygiene": {
                "total_rows": len(df_clean),
                "total_features": len(feature_names),
                "missing_imputed": missing_imputed
            },
            "file": os.path.basename(raw_file),
            "data_profile": data_profile,
            "targets": target_names,
            "target_summaries": target_summaries,
            "features_used": feature_names,
            "derived_evals": applied_evals,
            **score_keys,
            "is_best": is_best,
            "prescription": prescription,
            "recommendations": recommendations,
            "targets_benchmark": target_benchmarks,
            "feature_types": {
                "numeric": feature_schema["numeric"],
                "categorical": list(feature_schema["categorical"]),
                "datetime": feature_schema["datetime"],
            },
            "correlation_matrix": corr_matrix,
            "correlation_truncated": corr_truncated,
            "target_correlations": target_correlations
        }
        if exclusion_summary:
            history_record["exclusions"] = exclusion_summary
        auto_groups = [g for g in prepared if g.get("auto")]
        if auto_groups:      # no --features: the inputs the tool chose and why the other columns were left out (2.9.72)
            history_record["input_selection"] = {
                "auto": True, "features": [f for f in g0["features"] if f not in g0["created"]],
                "excluded": [{"column": d["column"], "reason": d["reason"]} for d in g0["excluded_inputs"]]}
            with open(os.path.join(staging_dir, AUTO_FEATURES_FILE), "w", encoding="utf-8") as f:
                f.write("\n".join(history_record["input_selection"]["features"]) + "\n")
        # the training conditions, kept in the record (--history, the report's history)
        if len(prepared) == 1:
            if g0["dtypes"]:
                history_record["dtypes"] = g0["dtypes"]
            if g0["positive_class"] is not None:
                history_record["positive_class"] = g0["positive_class"]
            if g0["where"]:
                history_record["exclude_where"] = g0["where"]
        if True:
            next_commands = build_next_commands(
                args=args, raw_file=raw_file, task=task, positive_class=g0["positive_class"],
                feature_names=g0["features"], created_features=g0["created"], applied_evals=g0["evals"],
                target_names=target_names, champion={**prim_champ, "weak_features": weak_feats}, schema=g0["schema"],
                mean_score=min_score, staging_dir=staging_dir, final_dir=guard.final_dir, exclude_where=g0["where"],
            )
        # 2.9.71: the tool counts the improvement rounds (AGENTS.md "사용법 단순성" 5): when the last IMPROVE_ROUNDS
        # runs of this task and targets (this one included) did not beat the best score before them, it stops offering
        # more retraining and offers the report instead
        improve_stopped = None
        if min_score < 0.90:
            scores = [float(h["mean_val_score"]) for h in history_records
                      if _same_experiment(h, task, target_names) and isinstance(h.get("mean_val_score"), (int, float))]
            scores.append(float(mean_val_score))
            if len(scores) > IMPROVE_ROUNDS and max(scores[-IMPROVE_ROUNDS:]) <= max(scores[:-IMPROVE_ROUNDS]) + 1e-4:
                improve_stopped = (f"최근 {IMPROVE_ROUNDS}번의 학습이 이전 최고 점수({max(scores[:-IMPROVE_ROUNDS]):.4f})를 넘지 못해 "
                                   f"재학습 제안을 멈춤. 지금까지의 최고 모델로 보고서를 만듭니다")
                next_commands = [{"purpose": REPORT_PURPOSE[task], "needs_input": False,
                                  "command": f"python {quote_arg(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'ml_report.py'))} --file {quote_arg(raw_file)}"}]
        if True:              # the command again with the condition covering the other notations
            where_fixes = [g["where_fix"] for g in prepared if g.get("where_fix")]
            next_commands = [c for i, c in enumerate(where_fixes) if c not in where_fixes[:i]] + next_commands
        history_record["next_commands"] = next_commands

        # Commit Transaction: atomic rename .tmp_vX -> vX, then record it in the ledger.
        # The ledger never names a version folder that does not exist (a failed commit leaves no record).
        # A run that trained something and is near the limit stops here, before the version is written: the next run
        # reuses every saved target and only packages (a stop after the commit could not continue, R10)
        if (RUN_FIT_SECONDS or RUN_PREP_SECONDS) and time_would_pass(0.0) and all(
                resume_find(resume_dir, P["targets"][t]["key"] + ".joblib") for t in target_names):
            raise Unfinished(f"학습: 타겟 {total_targets}/{total_targets}개 끝남, 저장됨 (버전 저장 전에 멈춤, 약 1단계 남음)")
        WATCHDOG_HOLD.append(True)             # the version is being written: the watchdog waits a little
        final_version_dir = guard.commit()
        append_history_ledger(ledger_path, history_record, on_recorded=guard.finalize)
        # saved for good: a stop from now on answers ok + UNVERIFIED, never CONTINUE (the same command would train again)
        mark_committed(f"학습 버전 {current_version}", version=current_version, version_dir=final_version_dir)
        _resume_cleanup(resume_dir, resume_used)
        unverified = verify_version_dir(final_version_dir, target_names, ledger_path, current_version)
        for u in unverified:
            sys.stderr.write(f"[WARN] 저장 확인 실패: {u}\n")
        log_progress(100, f"학습 및 산출물 패키징 완료: {current_version} ({os.path.basename(final_version_dir)}) [소요: {duration_sec}초]")

        # The most important inputs per target (users ask right after training)
        top_features: Dict[str, List[Dict[str, Any]]] = {}
        for t in target_names:
            top_features[t], imp_note = top_feature_shares(target_champions[t]["info"])
            if imp_note:
                train_warnings.append(f"타겟 '{t}': 중요 입력(top_features)을 내지 못했습니다 ({imp_note}).")
                sys.stderr.write(f"[WARN] {train_warnings[-1]}\n")
        split_note = save_failed_warning()      # a finished step saved nowhere (this run ended before the limit)
        if split_note and split_note not in train_warnings:
            train_warnings.append(split_note)
        top_brief = " | ".join(
            "중요 입력" + (f"({t})" if len(target_names) > 1 else "") + ": "
            + ", ".join(f"{r['feature']} {r['importance']:.2f}" for r in top_features[t][:TOP_FEATURES_PREVIEW])
            for t in target_names if top_features[t])

        # no --features: the chosen inputs and the left-out columns with their reasons (2.9.72)
        auto_brief = ""
        auto_keys: Dict[str, Any] = {}
        suspects = [d for g in prepared for d in g.get("suspect_inputs") or []]
        if auto_groups:
            chosen = history_record["input_selection"]["features"]
            excl = history_record["input_selection"]["excluded"]
            auto_keys = {"features_auto": chosen[:AUTO_SHOWN], "features_auto_count": len(chosen),
                         "features_auto_file": os.path.join(final_version_dir, AUTO_FEATURES_FILE),
                         "excluded_inputs": excl[:AUTO_SHOWN], "excluded_inputs_count": len(excl)}
            kinds = {d["column"]: INPUT_KIND_LABEL[d["kind"]] for d in g0["excluded_inputs"]}
            auto_brief = (f"입력 자동 선택 {len(chosen)}개" + (f", 뺀 열 {len(excl)}개: " + ", ".join(
                f"{d['column']}({kinds[d['column']]})" for d in excl[:5]) + (" 등" if len(excl) > 5 else "")
                if excl else ", 뺀 열 없음"))
        if suspects:
            auto_keys["suspect_inputs"] = [{"column": d["column"], "reason": d["reason"]} for d in suspects][:AUTO_SHOWN]
            preview_suspect = "의심 입력: " + ", ".join(f"{d['column']}({INPUT_KIND_LABEL[d['kind']]})" for d in suspects[:5])
            auto_brief = f"{auto_brief} | {preview_suspect}" if auto_brief else preview_suspect
        # LLM Directive
        short = "R2" if task == TASK_REGRESSION else "Macro-F1"
        target_brief = ", ".join([f"{t}({target_champions[t]['info']['champion_model']}: 검증 {short}={target_champions[t]['info']['val_score']:.4f})" for t in target_names])
        delta_txt = delta_score if delta_score == "base" else f"{delta_score:+.4f}"
        excl_brief = "; ".join(
            f"{t}: 학습 제외 {e['manual'] + e['auto'] + e.get('condition', 0)}행(직접 {e['manual']}, "
            + (f"조건 {e['condition']}, " if e.get("condition") else "") + f"자동 {e['auto']}), "
            f"검증 {short}(제외 행 포함)={e['all_rows_score'] if e['all_rows_score'] is not None else '-'}"
            for t, e in exclusion_summary.items() if e["manual"] or e["auto"] or e.get("condition"))
        preview = (
            f"[{current_version} 완료 / {task}] 타겟: [{target_brief}] | "
            + (f"{excl_brief} | " if excl_brief else "")
            + (f"{auto_brief} | " if auto_brief else "") +
            f"평균 {metric_label}={mean_val_score:.4f} (변화: {delta_txt}) | "
            + (f"{top_brief} | " if top_brief else "") +
            f"소요 {duration_sec}초 | "
            f"저장: {final_version_dir}"
        )
        agent_directive = (
            (f"UNVERIFIED: {current_version} 학습은 끝났지만 저장 결과를 확인하지 못했습니다 ({'; '.join(unverified)}). "
             f"다시 실행하지 말고 unverified를 그대로 알리세요. " if unverified else "VERIFIED: ")
            + f"{current_version} 학습을 마쳤습니다 ({task}, 평균 {metric_label}={mean_val_score:.4f}). "
            + ("학습에서 뺀 행이 있습니다. preview의 제외 행 수와 '제외 행 포함' 점수, excluded_reasons(이유별 행 수), "
               "excluded_preview(행, 직접/자동, 이유, 값)를 그대로, excluded_rows_file 경로도 알리세요. "
               "그 파일은 열지 말고 다른 이유를 짐작하지 마세요. " if excl_brief else "")
            + ("입력(--features)을 주지 않아 도구가 입력을 골랐습니다. preview의 입력 자동 선택(features_auto)과 뺀 열과 이유"
               "(excluded_inputs의 column, reason)를 알리세요. " if auto_groups else "")
            + ("ID·일련번호나 타겟과 같은 값으로 보이는 입력이 있습니다 (suspect_inputs): 그대로 알리고, 사용자가 빼라고 하면 "
               "--features에서 빼고 다시 학습합니다. " if suspects else "")
            + ("preview의 중요 입력(top_features, 전체 입력 대비 비중)을 알리세요. " if top_brief else "")
            + ("warnings(중복 행, 반복된 머리글 행, 합계 행 등)를 빠짐없이 알리세요. " if train_warnings else "") +
            (f"{improve_stopped} (improve_stopped를 알리세요). " if improve_stopped else "") +
            "다음: 사용자가 원하면 next_commands[0]을 실행하거나 예측으로 넘어갑니다. "
            "needs_input=true인 명령의 자리표시자(<...>)만 사용자가 준 값(파일이면 경로)으로 바꾸고, 모르면 먼저 물으세요."
        )

        shown_types, types_cut = capped_feature_types(history_record["feature_types"])
        result = {
            "status": "ok",
            "version": current_version,
            "tag": run_tag,
            "description": project_desc,
            "duration_sec": duration_sec,
            "data_root": data_root,
            "version_dir": final_version_dir,
            "targets": target_names,
            **score_keys,
            "is_best": is_best,
            "goal_met": bool(min_score >= 0.90),
            **({"improve_stopped": improve_stopped} if improve_stopped else {}),
            "verified": not unverified,
            "recommendations": recommendations,
            "next_commands": next_commands,
            "feature_types": shown_types,
            "truncated": types_cut,
            "top_features": top_features,
            **auto_keys,
            "preview": preview,
            "agent_directive": agent_directive
        }
        if train_warnings:
            result["warnings"] = train_warnings
        if exclusion_summary:
            result["exclusions"] = exclusion_summary
            if excluded_rows_file:
                result["excluded_rows_file"] = excluded_rows_file
        if excluded_rows_out:   # the rows and reasons in the JSON: the LLM never needs to open the CSV
            result.update(excluded_preview_keys(excluded_rows_out))
        if reused_targets:   # trained by earlier runs of the same command (split execution, R10)
            result["resumed_targets"] = reused_targets
        if column_aliases:
            result["column_aliases"] = column_aliases
        if unverified:
            result["unverified"] = unverified
        if guard.incomplete_versions:
            result["incomplete_versions"] = guard.incomplete_versions
        if types_cut:     # every input name is in the ledger record of this version
            result["feature_count"] = len(feature_names)
            result["full_result_file"] = os.path.join(data_root, "history.json")
        return result

    except BaseException as exc:
        # Includes KeyboardInterrupt (Ctrl+C) so the staging folder never leaks. A stop at the time limit (Unfinished)
        # leaves no version folder and no ledger record: the finished targets wait in .resume_train (or a fallback
        # place) for the next run.
        guard.rollback(quiet=isinstance(exc, Unfinished))
        raise
    finally:
        run_lock.release()


TASK_ORDER = [TASK_REGRESSION, TASK_CLASSIFICATION]
AUTO_FEATURES_FILE = "features_auto.txt"   # vX/: the inputs a run without --features chose, one per line (usable as --features)
IMPROVE_ROUNDS = 5   # training runs in a row without a new best before the tool stops offering retraining (2.9.71)
REPORT_PURPOSE = {  # the report and the calculator use the latest training version (2.9.83); regression also makes the calculator
    TASK_REGRESSION: "방금 학습한 버전으로 ML 결과 보고서(Word)와 성능예측표(엑셀 계산기) 생성 (성능예측표 입력은 그 버전의 입력 전부, "
                     "제목은 학습 때 데이터 설명에서 자동으로 정함)",
    TASK_CLASSIFICATION: "방금 학습한 버전으로 ML 결과 보고서(Word) 생성",
}


# ==============================================================================
# 9. CLI Interface & History Mode
# ==============================================================================

HISTORY_SHOWN = 30        # --history records in the JSON (the latest); history.json holds all (R2)
FEATURE_TYPES_SHOWN = 30  # names per feature_types list in the training result; the ledger record holds all


def capped_feature_types(ft: Dict[str, Any]) -> Tuple[Dict[str, Any], bool]:
    """feature_types of the training result with at most FEATURE_TYPES_SHOWN names per kind (R2)."""
    out, cut = {}, False
    for kind, names in ft.items():
        items = list(names.items()) if isinstance(names, dict) else list(names)
        if len(items) > FEATURE_TYPES_SHOWN:
            cut = True
            items = items[:FEATURE_TYPES_SHOWN]
        out[kind] = dict(items) if isinstance(names, dict) else items
    return out, cut


def print_history_summary(data_root: str) -> Dict[str, Any]:
    """
    Prints a human-readable summary table of all experiments to stderr and returns
    a compact JSON summary (stdout stays a single JSON line for the agent).
    """
    ledger_path = os.path.join(data_root, "history.jsonl")
    records = load_history_ledger(ledger_path)
    if not records:     # 0 records is a result, not an error (R2)
        return {
            "status": "ok", "verified": True, "no_history": True, "data_root": data_root,
            "record_count": 0, "count": 0, "shown": 0, "truncated": False, "records": [], "latest_version": None,
            "preview": "학습 기록 0건",
            "agent_directive": ("VERIFIED: 0건. 이 데이터 파일의 학습 기록(history.jsonl)이 없습니다. "
                                "사용자에게 학습 기록이 없다고 그대로 알리세요."),
        }

    proj_desc = records[-1].get("description", "")
    out = sys.stderr

    out.write("\n" + "=" * 125 + "\n")
    out.write(f"[STATUS] 실험 이력 (history.jsonl): {data_root}\n")
    if proj_desc:
        out.write(f"[STATUS] 설명: {proj_desc}\n")
    out.write("=" * 125 + "\n")
    out.write(f"{'Version':<8} | {'Tag / Memo':<16} | {'Task / Score':<16} | {'Delta':<10} | {'Time':<7} | {'Targets':<22} | {'Status'}\n")
    out.write(f"{'-'*8}-+-{'-'*16}-+-{'-'*16}-+-{'-'*10}-+-{'-'*7}-+-{'-'*22}-+-{'-'*15}\n")

    rows = []
    for r in records:
        ver = r.get("version", "-")
        tag = r.get("tag") or "-"
        task = r.get("task", TASK_REGRESSION)
        score = r.get("mean_val_score")
        delta = r.get("delta_score", "base")
        row = {
            "version": ver, "tag": tag, "task": task,
            "primary_metric": r.get("primary_metric", PRIMARY_METRIC.get(task, "val_r2")),
            "mean_val_score": score, "delta_score": delta,
            "is_best": bool(r.get("is_best")), "targets": r.get("targets", []),
        }
        rows.append(row)

        tag_s = tag if len(tag) <= 15 else tag[:13] + ".."
        short = "R2" if task == TASK_REGRESSION else "F1"
        score_s = f"{short} {score:.4f}" if isinstance(score, (int, float)) else "-"
        delta_str = f"{delta:+.4f}" if isinstance(delta, (int, float)) else str(delta)
        dur = f"{r.get('duration_sec', 0.0):.1f}s" if "duration_sec" in r else "-"
        targets_str = ", ".join(r.get("targets", []))
        if len(targets_str) > 20:
            targets_str = targets_str[:18] + ".."
        status = "[BEST 챔피언]" if r.get("is_best") else "[수렴]"
        out.write(f" {ver:<7} | {tag_s:<16} | {score_s:<16} | {delta_str:<10} | {dur:<7} | {targets_str:<22} | {status}\n")

    out.write("=" * 125 + "\n\n")

    final = final_version(data_root)      # the same choice as the report and prediction (a damaged latest folder is skipped)
    shown = rows[-HISTORY_SHOWN:]
    cut = len(shown) < len(rows)
    return {
        "status": "ok",
        "data_root": data_root,
        "description": proj_desc,
        "record_count": len(rows),
        "count": len(rows),
        "shown": len(shown),
        "truncated": cut,
        **({"full_result_file": os.path.join(data_root, "history.json")} if cut else {}),
        "latest_version": final["version"] if final else None,     # the final version the report and prediction use (2.9.83)
        "latest_version_dir": final["version_dir"] if final else None,
        **({"latest_version_note": f"마지막 학습 {', '.join(final['skipped'])}의 폴더가 없거나 손상돼 그 앞의 {final['version']}을(를) 씀"}
           if final and final["skipped"] else {}),
        "records": shown,
        "verified": True,
        "preview": (f"학습 기록 {len(rows)}건, 최종 버전 {final['version'] if final else '없음'} (보고서·예측이 쓰는 버전)"
                    + (f". 마지막 학습 {', '.join(final['skipped'])}의 폴더가 없거나 손상돼 그 앞 버전을 씀" if final and final["skipped"] else "")),
        "agent_directive": ("VERIFIED: 학습 이력입니다. records의 version, mean_val_score, delta_score, tag를 표로 알리세요. "
                            "history.jsonl을 다시 읽지 마세요."
                            + (f" records에는 최근 {len(shown)}건만 있습니다 (전체 {len(rows)}건: full_result_file)." if cut else "")),
    }


def main():
    parser = JsonArgumentParser(
        description="회귀/분류 모델 학습 (5개 모델 경쟁, 버전 관리)"
    )
    parser.add_argument("--file", help="입력 데이터 파일 경로 (.xlsx, .csv)")
    parser.add_argument("--features", help="독립변수(피처) 목록 (쉼표 구분 문자열 또는 파일 경로: .txt, .json, .csv). "
                                                 "생략하면 타겟을 뺀 모든 열에서 도구가 고름 (ID·상수·빈 열·타겟 누수 열 제외)")
    parser.add_argument("--targets", help="종속변수(타겟) 목록 (쉼표 구분 문자열 또는 파일 경로: .txt, .json, .csv)")
    parser.add_argument("--eval", help="방탄 AST 파생변수 생성 수식 (세미콜론 구분). 컬럼명은 전체 이름 또는 단위를 뺀 고유 이름")
    parser.add_argument("--tag", help="해당 회차 실험 의도 및 변경 사항 메모 태그 (예: '베이스라인_학습', '파생변수2종_추가')")
    parser.add_argument(
        "--desc", dest="description",
        help="데이터셋 도메인 및 모델링 목적 설명 (문자열 또는 파일 경로 .txt, .md 지원. 미지정 시 이전 회차 설명 자동 상속)"
    )
    parser.add_argument("--header-row", type=int, help="컬럼명 헤더 행 번호 (1-based, 미지정 시 상단 빈 행 자동 건너뜀)")
    parser.add_argument("--unit-row", type=int, help="단위 행 번호 (선택)")
    parser.add_argument("--data-row", type=int, help="데이터 시작 행 번호 (선택)")
    parser.add_argument("--dtypes", help="데이터 타입 수동 오버라이드 (예: '일시:datetime,설비코드:str')")
    parser.add_argument("--exclude-where", dest="exclude_where",
                        help="학습에서 뺄 행의 조건 (예: '생산효율 > 90; 라인 = B | 라인 = C'. ;는 그리고, |는 또는, "
                             "연산자 >= <= > < = != ~). 데이터 파일은 바꾸지 않음")
    parser.add_argument("--version", help="버전 수동 지정 (미지정 시 자동 계산 v1, v2...)")
    parser.add_argument("--overwrite", action="store_true",
                        help="--version으로 지정한 버전 폴더가 이미 있으면 교체 (사용자 확인 후에만 사용)")
    parser.add_argument("--dry-run", dest="dry_run", action="store_true",
                        help="학습하지 않고 만들 버전 폴더와 (--overwrite면) 바꿀 기존 폴더만 보여 줌 (confirm_commands에 실제 명령)")
    parser.add_argument("--task", choices=[TASK_REGRESSION, TASK_CLASSIFICATION], default=None,
                        help="학습 유형: regression(기본값, 연속 수치 예측) 또는 classification(범주/라벨 예측)")
    parser.add_argument("--positive-class", dest="positive_class",
                        help="이진 분류의 양성 클래스 (예: 불량). 지정 시 양성 F1이 최대가 되도록 판정 기준 확률을 자동 조정")
    parser.add_argument("--seed", type=int, default=42, help="재현용 난수 시드 (기본값: 42)")
    parser.add_argument("--history", action="store_true", help="누적 실험 이력 조회 모드 (표는 stderr, 요약 JSON은 stdout)")

    args = parser.parse_args()

    if args.dry_run and args.history:
        parser.error("--dry-run은 --features/--targets 학습에만 씁니다 (--history는 읽기만 함)")
    if args.history:
        if not args.file:
            parser.error("--history 조회 시 --file 인자가 필요합니다.")
        raw_file = os.path.abspath(args.file)
        file_stem = os.path.splitext(os.path.basename(raw_file))[0]
        data_root = os.path.join(os.path.dirname(raw_file), f"{file_stem}_ml_data")
        run_cli(lambda: print_history_summary(data_root))
        return

    if not args.file or not args.targets:
        parser.error("--file과 --targets는 모델 학습에 필수적인 인자입니다 (--features를 빼면 도구가 입력 열을 고름).")

    run_cli(lambda: execute_ml_train(args), hint="확인할 것: 열 이름과 --eval 식.")


if __name__ == "__main__":
    main()
