# -*- coding: utf-8 -*-
"""
ml_train.py - Autonomous Regression / Classification Machine Learning Engine with
Collision-Free AST Feature Engineering, 5-Model Competitive Benchmark,
Atomic Version Packaging, and Single Source of Truth (SSOT) Ledger.

--task regression (default) : champion by Out-of-Fold R2
--task classification       : champion by Out-of-Fold Macro-F1 (Stratified K-Fold)

--best (2.4.0)              : every target of the history retrained with its best recipe into one version per task

Shared runtime (data loading, --eval engine, CLI protocol) is imported from ml_predict.py.

Designed for lightweight 26B-30B LLM agents with 1-turn declarative execution.
100% Standalone: Pure Python, Scikit-learn, Pandas, NumPy, and Joblib.
"""

import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from ml_predict import (  # noqa: E402  shared runtime lives in ml_predict.py
    setup_console, dumps_json, run_cli, JsonArgumentParser,
    InputError, sanitize_filename, normalize_version, warn_long_path,
    read_text_file, load_json_file, resolve_column_names, coerce_numeric,
    load_data_file as _load_table, apply_eval_expressions, eval_source_columns,
    parse_datetime, normalize_category, prepare_features, date_part_column, input_column_owner, DATE_PARTS,
    strip_unit,
    retry_fs, remove_tree, atomic_write_text, write_csv_excel_compatible,
)

setup_console()

import re
import json
import time
import argparse
import warnings
from datetime import datetime
from typing import Any, Callable, Dict, List, Optional, Tuple

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
    confusion_matrix, classification_report, make_scorer,
)
from sklearn.inspection import permutation_importance
from sklearn.exceptions import ConvergenceWarning

# 5 Core Regression Algorithms
from sklearn.ensemble import HistGradientBoostingRegressor, RandomForestRegressor, ExtraTreesRegressor
from sklearn.linear_model import RidgeCV, ElasticNetCV
from sklearn.compose import ColumnTransformer
from sklearn.preprocessing import OneHotEncoder
from sklearn.base import clone

# 5 Core Classification Algorithms (--task classification)
from sklearn.ensemble import HistGradientBoostingClassifier, RandomForestClassifier, ExtraTreesClassifier
from sklearn.linear_model import LogisticRegressionCV
from sklearn.neighbors import KNeighborsClassifier

TASK_REGRESSION = "regression"
TASK_CLASSIFICATION = "classification"
PRIMARY_METRIC = {TASK_REGRESSION: "val_r2", TASK_CLASSIFICATION: "val_f1_macro"}
METRIC_LABEL = {TASK_REGRESSION: "검증 R2", TASK_CLASSIFICATION: "검증 Macro-F1"}   # 점수 용어: 학습 / 검증 (제외 행 포함/미포함) (2.5.0)

warnings.filterwarnings("ignore", category=UserWarning)
warnings.filterwarnings("ignore", category=ConvergenceWarning)
warnings.filterwarnings("ignore", category=FutureWarning)


# ==============================================================================
# 1. Console Progress & Formatting (Clean Terminal Principle - No Emojis)
# ==============================================================================

def log_progress(percent: int, message: str) -> None:
    """
    Outputs real-time progress to stderr with immediate flush.
    Ensures newline-based streaming in agent environments (OpenCode, CI, Subprocess)
    to prevent idle stream timeouts while maintaining clean formatting.
    """
    bar_len = 20
    filled = int(bar_len * percent / 100)
    bar = "=" * filled + "-" * (bar_len - filled)
    
    # In agent/piped execution (OpenCode, etc.), output with newline so stream readers don't block
    is_interactive = hasattr(sys.stderr, "isatty") and sys.stderr.isatty()
    if is_interactive:
        sys.stderr.write(f"\r[PROGRESS {percent:3d}%] [{bar}] {message}")
        if percent >= 100:
            sys.stderr.write("\n")
    else:
        sys.stderr.write(f"[PROGRESS {percent:3d}%] [{bar}] {message}\n")
    
    sys.stderr.flush()


# ==============================================================================
# 2. Bulletproof AST Formula Sandbox & Feature Engineering (--eval)
# ==============================================================================

def run_safe_eval(df: pd.DataFrame, eval_str: str, target_cols: List[str]) -> Tuple[pd.DataFrame, List[str], List[str]]:
    """
    Safely evaluates semicolon-separated derived feature formulas (--eval).
    Column names may be written in full ('반응온도(°C)') or without the unit suffix
    ('반응온도') when unique. Stored expressions always use full column names so that
    ml_predict.py replays them exactly. Target references are blocked (leakage guard).
    """
    if not eval_str or not eval_str.strip():
        return df, [], []
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


def load_data_file(
    file_path: str,
    header_row: Optional[int] = None,
    unit_row: Optional[int] = None,
    data_row: Optional[int] = None
) -> Tuple[pd.DataFrame, str]:
    """
    Loads tabular data from CSV (multi-encoding) or Excel (.xlsx) via the shared loader
    (identical parsing in ml_predict.py). Returns (dataframe, file_name).
    """
    df_data = _load_table(file_path, header_row=header_row, unit_row=unit_row, data_row=data_row)
    return df_data, os.path.basename(file_path)


# ==============================================================================
# 4. Fail-Fast Validation & Type Sanitization
# ==============================================================================

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

    header_candidates = df.attrs.get("header_candidates")    # two-level / merged header -> [추천] --header-row N
    feature_names, feat_alias = resolve_column_names(feature_names, all_available, "피처(-x)", header_candidates=header_candidates)
    target_names, tgt_alias = resolve_column_names(target_names, all_available, "타겟(-y)", header_candidates=header_candidates)
    feature_names = list(dict.fromkeys(feature_names))
    target_names = list(dict.fromkeys(target_names))

    # Data Leakage Check: Targets in Features
    overlap = [t for t in target_names if t in set(feature_names)]
    if overlap:
        raise InputError(
            f"[DATA LEAKAGE DETECTED] 타겟 변수 '{', '.join(overlap)}'이(가) 피처(-x) 목록에 포함되어 있습니다. "
            f"종속변수는 독립변수로 사용될 수 없습니다."
        )

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
                            df_clean[c_name] = coerce_numeric(df_clean[c_name]).fillna(0).astype(int)
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
            non_null = raw.notna() & (raw.astype(str).str.strip() != "")
            n = int(non_null.sum())
            is_dt_obj = raw.map(lambda v: isinstance(v, (datetime, pd.Timestamp, np.datetime64)))
            if pd.api.types.is_datetime64_any_dtype(raw) or (n and is_dt_obj.sum() >= 0.9 * n):
                kind = "datetime"  # Excel date cells arrive as datetime objects
            elif n == 0 or pd.api.types.is_numeric_dtype(raw):
                kind = "numeric"
            elif coerce_numeric(raw).notna().sum() >= 0.9 * n:
                kind = "numeric"
            elif parse_datetime(raw).notna().sum() >= 0.9 * n:
                kind = "datetime"
            else:
                kind = "categorical"

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
                    f"-x 목록에서 제외하십시오. 숫자 컬럼인데 잘못 판별되었다면 --dtypes '{col}:float'를 지정하십시오."
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
            ("imputer", SimpleImputer(strategy="median")),
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


def feature_importance_ranks(
    pipe: Pipeline, X: pd.DataFrame, y: np.ndarray, schema: Dict[str, Any], seed: int, scoring: Optional[str]
) -> List[Dict[str, Any]]:
    """
    Permutation importance on at most PERMUTATION_MAX_ROWS rows (keeps large data fast),
    aggregated back to user-facing columns (date parts are summed into their date column).
    """
    Xs, ys = X, y
    if len(X) > PERMUTATION_MAX_ROWS:
        idx = np.random.default_rng(seed).choice(len(X), PERMUTATION_MAX_ROWS, replace=False)
        Xs, ys = X.iloc[idx], y[idx]
    try:
        res = permutation_importance(pipe, Xs, ys, n_repeats=5, random_state=seed, scoring=scoring)
        raw = np.maximum(res.importances_mean, 0)
    except Exception:
        raw = np.ones(X.shape[1])

    owner = input_column_owner(schema)
    agg: Dict[str, float] = {}
    for col, imp in zip(X.columns, raw):
        agg[owner.get(col, col)] = agg.get(owner.get(col, col), 0.0) + float(imp)
    total = sum(agg.values()) or 1.0
    ranks = [{"feature": f, "importance": round(v, 4), "pct": round(v / total * 100, 1)} for f, v in agg.items()]
    return sorted(ranks, key=lambda r: r["importance"], reverse=True)


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
        "RandomForest": RandomForestRegressor(n_estimators=100, random_state=seed, n_jobs=-1, **_bagging_kwargs(n_rows)),
        "ExtraTrees": ExtraTreesRegressor(n_estimators=100, random_state=seed, n_jobs=-1, **_bagging_kwargs(n_rows)),
        "Ridge": RidgeCV(alphas=LINEAR_ALPHAS),
        "ElasticNet": ElasticNetCV(l1_ratio=[0.1, 0.5, 0.9], cv=3, max_iter=5000, random_state=seed),
    }


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


def train_and_benchmark_models(
    X: pd.DataFrame,
    y_series: pd.Series,
    target_name: str,
    schema: Dict[str, Any],
    seed: int = 42,
    log_offset: int = 20,
    log_weight: int = 50
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
    # small training folds (heavy duplication): the inner CV of ElasticNetCV must fit in every training fold
    min_train = min(len(tr) for tr, _ in splits)
    if min_train < 6:
        models["ElasticNet"].set_params(cv=max(2, min(3, min_train // 2)))

    leaderboard = []
    model_oof_map = {}

    for m_idx, (model_name, regressor) in enumerate(models.items()):
        oof_preds = np.zeros(len(y))
        train_r2_list = []
        params = {}

        for fold, (train_idx, val_idx) in enumerate(splits, 1):
            pct = log_offset + int(log_weight * (m_idx * n_splits + fold) / (model_count * n_splits))
            log_progress(pct, f"Target '{target_name}' | {model_name} (Fold {fold}/{n_splits})...")

            # Leakage-Free Pipeline: fitted ONLY on training fold
            pipe = build_pipeline(clone(regressor), schema, "regressor")
            pipe.fit(X.iloc[train_idx], y[train_idx])
            train_r2_list.append(r2_score(y[train_idx], pipe.predict(X.iloc[train_idx])))
            oof_preds[val_idx] = pipe.predict(X.iloc[val_idx])
            params = _model_params(pipe)
            del pipe  # Explicit memory cleanup per fold

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

    # Fit Champion Model on 100% of data
    log_progress(log_offset + log_weight - 4, f"Target '{target_name}' | 1위 챔피언({champion_name}) 전체 데이터 적합 중...")
    final_pipeline = build_pipeline(clone(models[champion_name]), schema, "regressor")
    final_pipeline.fit(X, y)

    log_progress(log_offset + log_weight - 2, f"Target '{target_name}' | 순열 피처 중요도(Feature Importance) 산출 중...")
    feature_ranks = feature_importance_ranks(final_pipeline, X, y, schema, seed, scoring=None)

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

def normalize_labels(series: pd.Series) -> pd.Series:
    """
    Converts class labels to stable strings so the same class never splits by storage type:
    1, 1.0, '1' -> '1'; ' 양품 ' -> '양품'. Placeholder strings become missing.
    """
    def _norm(v):
        if v is None or (isinstance(v, float) and np.isnan(v)):
            return None
        if isinstance(v, (bool, np.bool_)):
            return str(bool(v))
        if isinstance(v, (int, np.integer)):
            return str(int(v))
        if isinstance(v, (float, np.floating)):
            return str(int(v)) if float(v).is_integer() else str(float(v))
        t = str(v).strip()
        if t in ("", "-", "N/A", "n/a", "NA", "null", "NULL", "None", "nan", "NaN", "미측정"):
            return None
        try:
            f = float(t.replace(",", ""))
            if f.is_integer() and re.fullmatch(r"[+-]?[\d,]+(\.0+)?", t):
                return str(int(f))
        except ValueError:
            pass
        return t

    return series.astype("object").map(_norm)


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
        "RandomForest": RandomForestClassifier(n_estimators=100, random_state=seed, n_jobs=-1, class_weight=cw,
                                               **_bagging_kwargs(n_rows)),
        "ExtraTrees": ExtraTreesClassifier(n_estimators=100, random_state=seed, n_jobs=-1, class_weight=cw,
                                           **_bagging_kwargs(n_rows)),
        # make_scorer object (not the "f1_macro" string): the string fails for binary labels in scikit-learn 1.9
        "LogisticRegression": LogisticRegressionCV(Cs=8, cv=3, max_iter=3000, class_weight=cw,
                                                   scoring=make_scorer(f1_score, average="macro")),
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
    positive_class: Optional[str] = None
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
            raise InputError(f"--positive-class '{positive_class}'가 타겟 '{target_name}'의 클래스에 없습니다. [추천] {classes}")

    min_ratio = min(class_counts.values()) / len(y)
    balanced = min_ratio < 0.20
    if balanced:
        sys.stderr.write(
            f"[WARN] 타겟 '{target_name}' 클래스 불균형 감지 (최소 클래스 비율 {min_ratio:.1%}) -> class_weight='balanced' 적용\n"
        )

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
    # small training folds: KNN cannot ask for more neighbours than a training fold has
    min_train = min(len(tr) for tr, _ in splits)
    if "KNeighbors" in models and min_train < 5:
        models["KNeighbors"].set_params(n_neighbors=max(1, min_train))

    leaderboard = []
    model_oof_map: Dict[str, np.ndarray] = {}
    model_proba_map: Dict[str, np.ndarray] = {}

    for m_idx, (model_name, clf) in enumerate(models.items()):
        oof_pred = np.empty(len(y), dtype=object)
        oof_proba = np.zeros((len(y), len(classes)))
        train_f1_list = []
        params = {}

        for fold, (train_idx, val_idx) in enumerate(splits, 1):
            pct = log_offset + int(log_weight * (m_idx * n_splits + fold) / (model_count * n_splits))
            log_progress(pct, f"Target '{target_name}' | {model_name} (Fold {fold}/{n_splits})...")

            pipe = build_pipeline(clone(clf), schema, "classifier")
            pipe.fit(X.iloc[train_idx], y[train_idx])
            train_f1_list.append(f1_score(y[train_idx], pipe.predict(X.iloc[train_idx]), average="macro", zero_division=0))
            oof_pred[val_idx] = pipe.predict(X.iloc[val_idx])
            proba = pipe.predict_proba(X.iloc[val_idx])
            col_idx = [classes.index(c) for c in pipe.classes_]
            oof_proba[np.ix_(val_idx, col_idx)] = proba
            params = _model_params(pipe)
            del pipe

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

    log_progress(log_offset + log_weight - 4, f"Target '{target_name}' | 1위 챔피언({champion_name}) 전체 데이터 적합 중...")
    final_pipeline = build_pipeline(clone(models[champion_name]), schema, "classifier")
    final_pipeline.fit(X, y)

    log_progress(log_offset + log_weight - 2, f"Target '{target_name}' | 순열 피처 중요도(Feature Importance) 산출 중...")
    feature_ranks = feature_importance_ranks(final_pipeline, X, y, schema, seed, scoring="f1_macro")

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
                    f"[추천] --version {self._next_version()}"
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

    def rollback(self) -> None:
        """Undoes a run that was not finalized: staging folder, committed folder and parked backup."""
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

    # 1. Append to history.jsonl (streaming ledger)
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


def _q(v: str) -> str:
    return '"' + str(v) + '"'


def layout_flags(args: argparse.Namespace) -> str:
    """' --header-row 4 --unit-row 5' when the data was read with a layout: a new file of the same layout needs it too."""
    return "".join(f" {flag} {int(val)}" for flag, val in (
        ("--header-row", getattr(args, "header_row", None)), ("--unit-row", getattr(args, "unit_row", None)),
        ("--data-row", getattr(args, "data_row", None))) if val is not None)


def predict_next_command(final_dir: str, args: argparse.Namespace, purpose: str) -> Dict[str, Any]:
    here = os.path.dirname(os.path.abspath(__file__))
    return {"purpose": purpose,
            "command": f"python {_q(os.path.join(here, 'ml_predict.py'))} -m {_q(final_dir)} -f \"<새데이터.xlsx>\""
                       + layout_flags(args),
            "needs_input": True, "placeholder": "<새데이터.xlsx>"}


def build_next_commands(
    args: argparse.Namespace, raw_file: str, task: str, positive_class: Optional[str],
    feature_names: List[str], created_features: List[str], applied_evals: List[str],
    target_names: List[str], champion: Dict[str, Any], schema: Dict[str, Any],
    mean_score: float, staging_dir: str, final_dir: str, data_root: str,
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
    train_py = _q(os.path.join(here, "ml_train.py"))
    base = [f"python {train_py}", "-f", _q(raw_file)]
    for flag, val in (("--header-row", args.header_row), ("--unit-row", args.unit_row),
                      ("--data-row", args.data_row), ("--dtypes", args.dtypes)):
        if val is not None:
            base += [flag, _q(val)]
    if task == TASK_CLASSIFICATION:
        base += ["--task", TASK_CLASSIFICATION]
        if positive_class is not None:
            base += ["--positive-class", _q(positive_class)]

    def names_arg(names: List[str], fname: str, limit: int = 15) -> str:
        if len(names) <= limit and not any('"' in n for n in names):
            return _q(",".join(names))
        with open(os.path.join(staging_dir, fname), "w", encoding="utf-8") as f:
            f.write("\n".join(names) + "\n")
        return _q(os.path.join(final_dir, fname))

    def features_arg(feats: List[str], tag: str) -> str:
        return names_arg(feats, f"next_features_{tag}.txt")

    def quotable(*texts: Optional[str]) -> bool:
        return not any(t is not None and '"' in str(t) for t in texts)

    base += ["-y", names_arg(target_names, "next_targets.txt", limit=10 ** 6)]
    retrain_ok = quotable(args.dtypes, positive_class)

    base_features = [f for f in feature_names if f not in created_features]
    cmds: List[Dict[str, Any]] = []

    if mean_score >= 0.90:  # caller passes the weakest target's score
        cmds.append({
            "purpose": REPORT_PURPOSE[task],
            "command": f"python {_q(os.path.join(here, 'ml_report.py'))} -f {_q(raw_file)}",
            "needs_input": False,
        })
        cmds.append(predict_next_command(
            final_dir, args, "목표 성능 달성: 새 데이터 예측 (<새데이터.xlsx>를 사용자가 준 실제 파일 경로로 바꿀 것)"))
        return cmds

    weak = [w for w in champion.get("weak_features", []) if w in feature_names]
    keep_base = [f for f in base_features if f not in weak]
    keep_evals = [e for e in applied_evals if e.split("=", 1)[0].strip() not in weak]
    if weak and keep_base and retrain_ok and quotable(*keep_evals):
        parts = base + ["-x", features_arg(keep_base, "pruned")]
        if keep_evals:
            parts += ["--eval", _q("; ".join(keep_evals))]
        parts += ["--tag", _q("저기여피처_제외")]
        cmds.append({"purpose": f"기여도 1% 이하 피처 {weak[:5]} 제외 후 재학습", "command": " ".join(parts),
                     "needs_input": False})

    numeric_top = [r["feature"] for r in champion.get("top_features", []) if r["feature"] in schema["numeric"]]
    if len(numeric_top) >= 2:
        a, b = numeric_top[0], numeric_top[1]
        new_name = f"{strip_unit(a)}x{strip_unit(b)}"
        existing = set(feature_names) | {e.split("=", 1)[0].strip() for e in applied_evals}
        if new_name not in existing and retrain_ok and quotable(*applied_evals, a, b):
            evals = applied_evals + [f"{new_name} = {a} * {b}"]
            parts = base + ["-x", features_arg(base_features, "all"), "--eval", _q("; ".join(evals)),
                            "--tag", _q("교차항_추가")]
            cmds.append({"purpose": f"상위 피처 '{a}', '{b}' 교차항 추가 후 재학습", "command": " ".join(parts),
                         "needs_input": False})
    return cmds


# ==============================================================================
# 8. Main Execution Controller
# ==============================================================================

def _same_experiment(record: Dict[str, Any], task: str, targets: List[str]) -> bool:
    """Ledger records are comparable (delta_score, is_best) only for the same task and target set."""
    return record.get("task", TASK_REGRESSION) == task and sorted(record.get("targets", [])) == sorted(targets)


def data_fingerprint(path: str) -> Dict[str, Any]:
    """Size, modification time and SHA-256 of the data file (best-version freshness check)."""
    import hashlib
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    st = os.stat(path)
    return {"size": int(st.st_size), "mtime": datetime.fromtimestamp(st.st_mtime).isoformat(timespec="seconds"),
            "sha256": h.hexdigest()}


# ==============================================================================
# 7-B. Row Exclusion: '제외' column + automatic outlier detection (2.5.0)
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
SNAPSHOT_ROW_COL = "원본행"              # model_dataset.csv: Excel row number of each row (2.7.0)
SNAPSHOT_EXCL_PREFIX = "학습제외_"       # model_dataset.csv: 1 = left out of training for that target (2.7.0)
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


def _exclude_mark(v: Any) -> str:
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


def plan_exclusions(marks: Optional[pd.Series], y: pd.Series, X: pd.DataFrame, schema: Dict[str, Any],
                    created: List[str], task: str, target_name: str) -> Dict[str, Any]:
    """Rows to leave out of training for one target: the user's '제외' marks plus automatic outliers."""
    n = len(y)
    valid = y.notna().to_numpy()
    text = (marks.map(_exclude_mark).to_numpy() if marks is not None else np.array([""] * n, dtype=object))
    forced = text == INCLUDE_MARK
    manual = (text != "") & ~forced & valid
    candidate = valid & ~manual & ~forced

    check_cols = ([target_name] if task == TASK_REGRESSION else []) + \
        [c for c in schema["numeric"] if c not in created and c in X.columns]
    auto = np.zeros(n, dtype=bool)
    reasons = [""] * n
    for col in check_cols:
        series = y if col == target_name else X[col]
        flags, med = _outlier_flags(series.where(candidate))
        for i in np.flatnonzero(flags):
            auto[i] = True
            val = pd.to_numeric(pd.Series([series.iloc[i]]), errors="coerce").iloc[0]
            reasons[i] += ("; " if reasons[i] else "") + f"{strip_unit(col)} {val:.6g} (중앙값 {med:.6g})"

    note = None
    if auto.any():
        if auto.sum() > OUTLIER_MAX_FRACTION * max(candidate.sum(), 1):
            note = (f"이상치 후보가 {int(auto.sum())}행({auto.sum() / max(candidate.sum(), 1):.0%})으로 많아 자동 제외하지 않았습니다 "
                    f"(값이 넓게 퍼진 데이터일 수 있음)")
            auto[:] = False
        elif (valid & ~manual & ~auto).sum() < MIN_TRAIN_ROWS:
            note = f"자동 제외하면 학습 행이 {MIN_TRAIN_ROWS}개 미만이 되어 자동 제외하지 않았습니다"
            auto[:] = False
    if task == TASK_CLASSIFICATION and auto.any():
        y_arr = y.to_numpy()
        for cls in pd.unique(y[valid]):
            in_cls = valid & (y_arr == cls)
            if (in_cls & ~manual & ~auto).sum() < 2:
                auto[in_cls] = False   # never remove a class from training
    return {"manual": manual, "auto": auto, "forced": forced, "mask": manual | auto,
            "text": text, "reasons": reasons, "note": note, "checked_columns": check_cols}


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


def _prepare_group(df_raw: pd.DataFrame, group: Dict[str, Any], task: str) -> Dict[str, Any]:
    """
    One feature recipe (features, --eval formulas, --dtypes, positive class) applied to the loaded data.
    A normal run has one group; a best run (--best) has one group per distinct recipe.
    """
    feats_in, tgts_in = group["features"], group["targets"]
    feature_names = parse_feature_or_target_list(feats_in, "features") if isinstance(feats_in, str) else list(feats_in or [])
    target_names = parse_feature_or_target_list(tgts_in, "targets") if isinstance(tgts_in, str) else list(tgts_in or [])
    if not feature_names:
        raise InputError("학습에 투입할 독립변수 목록(-x, --features)이 지정되지 않았거나 파일이 비어 있습니다.")
    if not target_names:
        raise InputError("예측할 종속변수 목록(-y, --targets)이 지정되지 않았거나 파일이 비어 있습니다.")
    dtypes = group.get("dtypes")

    # Fail-Fast Validation & Type Sanitization
    log_progress(10, "컬럼 유효성 검증 및 타입 정제 중...")
    df_clean, feature_names, target_names, column_aliases = validate_and_sanitize_data(
        df=df_raw, feature_names=feature_names, target_names=target_names, dtypes_arg=dtypes, task=task)
    excl_col = find_exclude_column(list(df_raw.columns))
    if excl_col and excl_col in target_names:
        raise InputError(f"'{excl_col}' 열은 학습에서 뺄 행을 표시하는 열이라 타겟(-y)으로 쓸 수 없습니다.")
    if excl_col and excl_col in feature_names:
        feature_names = [f for f in feature_names if f != excl_col]
        sys.stderr.write(f"[WARN] '{excl_col}' 열은 뺄 행 표시용이라 피처에서 제외했습니다.\n")
        if not feature_names:
            raise InputError(f"'{excl_col}' 열을 빼면 학습할 피처(-x)가 없습니다.")

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
    # Columns the user must provide at prediction time: own raw features + sources of derived features (2.9.1)
    sources = eval_source_columns(applied_evals, [str(c) for c in df_raw.columns])
    raw_required = [f for f in feature_names if f not in created_features]
    raw_required += [c for c in sources if c not in raw_required and c not in target_names]
    for c in raw_required:
        if c not in df_clean.columns and c in df_raw.columns:
            df_clean[c] = df_raw[c]
    return {"df": df_clean, "features": feature_names, "targets": target_names, "aliases": column_aliases,
            "raw_required": raw_required,
            "evals": applied_evals, "created": created_features, "schema": feature_schema, "X": X_all,
            "dtypes": dtypes, "positive_class": group.get("positive_class")}


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


def execute_ml_train(args: argparse.Namespace) -> Dict[str, Any]:
    """Orchestrates 1-Turn declarative regression / classification pipeline."""
    task = (getattr(args, "task", None) or TASK_REGRESSION).lower()
    positive_class = getattr(args, "positive_class", None)
    if positive_class is not None:
        if task != TASK_CLASSIFICATION:
            raise InputError("--positive-class는 --task classification과 함께 사용해야 합니다.")
        positive_class = normalize_labels(pd.Series([positive_class])).iloc[0]
    group = {"features": args.features, "targets": args.targets, "eval": args.eval,
             "dtypes": args.dtypes, "positive_class": positive_class}
    return train_version(args, task, [group])


def train_version(args: argparse.Namespace, task: str, groups: List[Dict[str, Any]],
                  best_info: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    """
    Trains every target of every group into ONE new version folder.
    Normal runs: one group (all -y targets share the -x features). Best runs (--best): one group per
    distinct recipe, so each target keeps its own features/--eval; each model_<t>.json carries its own
    features_required, so ml_predict.py needs no change.
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
        log_progress(5, f"입력 데이터 로드 중: {os.path.basename(raw_file)}...")
        df_raw, file_name = load_data_file(
            file_path=raw_file,
            header_row=args.header_row,
            unit_row=args.unit_row,
            data_row=args.data_row
        )

        prepared = [_prepare_group(df_raw, g, task) for g in groups]
        excl_col = find_exclude_column(list(df_raw.columns))
        source_rows = df_raw.attrs.get("source_rows") or list(range(1, len(df_raw) + 1))
        exclusions: Dict[str, Dict[str, Any]] = {}
        excluded_rows_out: List[Dict[str, Any]] = []
        g0 = prepared[0]
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

        # Multi-Target 5-Model Competitive Training (one task per run)
        target_benchmarks = {}
        target_champions = {}
        target_summaries = {}
        all_champion_scores = []
        train_warnings: List[str] = []
        benchmark_fn = train_and_benchmark_models if task == TASK_REGRESSION else train_and_benchmark_classifiers

        total_targets = len(target_names)
        for t_idx, target_name in enumerate(target_names):
            g = target_group[target_name]
            t_offset = 25 + int(55 * t_idx / total_targets)
            t_weight = int(55 / total_targets)

            log_progress(t_offset, f"타겟 [{t_idx+1}/{total_targets}] '{target_name}' 5대 모델 경쟁 학습 ({task})...")

            extra = {"positive_class": g["positive_class"]} if task == TASK_CLASSIFICATION else {}
            y_full = g["df"][target_name]
            plan = plan_exclusions(g["df"][excl_col] if excl_col else None, y_full, g["X"], g["schema"],
                                   g["created"], task, target_name)
            if plan["mask"].any():
                sys.stderr.write(f"[INFO] 타겟 '{target_name}': 학습에서 뺀 행 {int(plan['mask'].sum())}개 "
                                 f"(직접 표시 {int(plan['manual'].sum())}, 자동 이상치 {int(plan['auto'].sum())})\n")
            if plan["note"]:
                sys.stderr.write(f"[WARN] 타겟 '{target_name}': {plan['note']}\n")
            t_bench, champion_pipe, champion_info, oof_df = benchmark_fn(
                X=g["X"],
                y_series=y_full.mask(plan["mask"]),
                target_name=target_name,
                schema=g["schema"],
                seed=args.seed,
                log_offset=t_offset,
                log_weight=t_weight,
                **extra
            )

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
            champion_info["excluded"] = excl_info
            if all_score is not None:
                champion_info["all_rows_score"] = all_score
            exclusions[target_name] = {**excl_info, "mask": plan["mask"]}
            for k, i in enumerate(np.flatnonzero(plan["mask"])):
                actual = y_full.iloc[i]
                excluded_rows_out.append({
                    "타겟": target_name,
                    "원본 행": source_rows[i] if i < len(source_rows) else i + 1,
                    "구분": "직접 표시" if plan["manual"][i] else "자동 이상치",
                    "이유": plan["text"][i] if plan["manual"][i] else plan["reasons"][i],
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
        history_records = load_history_ledger(ledger_path)

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
                "derived_evals": g["evals"],
                "target_summary": target_summaries.get(target_name, {}),
                "sklearn_version": sklearn.__version__,
                "python_version": sys.version.split()[0]
            })
            if best_info:
                meta_json["best_source"] = best_info["best_source"].get(target_name)
            with open(os.path.join(staging_dir, f"model_{safe_target}.json"), "w", encoding="utf-8") as f:
                f.write(dumps_json(meta_json, indent=2))

        # 3. Clean Dataset Snapshot (model_dataset.csv)
        raw_extra = [c for g in prepared for c in g["raw_required"] if c not in feature_names]
        dataset_snapshot_cols = list(dict.fromkeys(c for c in feature_names + raw_extra + target_names if c in df_clean.columns))
        snapshot_df = df_clean[dataset_snapshot_cols].copy()
        # 2.7.0: Excel row numbers and per-target training exclusions, so ml_predict.py can compare its predictions
        # with the measured rows the model learned from (evidence rows are shown as Excel row numbers)
        snapshot_df[SNAPSHOT_ROW_COL] = [source_rows[i] if i < len(source_rows) else i + 1 for i in range(len(snapshot_df))]
        for t in target_names:
            snapshot_df[f"{SNAPSHOT_EXCL_PREFIX}{t}"] = np.where(exclusions[t]["mask"], 1, 0)
        snapshot_path = os.path.join(staging_dir, "model_dataset.csv")

        # Windows 한국어 Excel 호환: cp949 기본, 표현 불가 문자 포함 시 utf-8-sig 자동 폴백
        write_csv_excel_compatible(snapshot_df, snapshot_path, encoding=getattr(args, "encoding", None))

        # 3-B. Rows left out of training (user marks + automatic outliers), for the user to review in Excel
        excluded_rows_file = None
        if excluded_rows_out:
            ex_path = write_csv_excel_compatible(pd.DataFrame(excluded_rows_out), os.path.join(staging_dir, "excluded_rows.csv"))
            excluded_rows_file = os.path.join(guard.final_dir, os.path.basename(ex_path))
        exclusion_summary = {t: {k: v for k, v in e.items() if k != "mask"} for t, e in exclusions.items()
                             if e["manual"] or e["auto"] or e["forced_include"] or e["auto_note"]}

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
        data_profile = {
            "source_file": os.path.basename(raw_file),
            "rows_loaded": int(len(df_raw)),
            "columns_loaded": int(df_raw.shape[1]),
            "rows_used": {t: int(target_benchmarks[t]["sample_count"]) for t in target_names},
            "target_missing_rows": {t: int(target_group[t]["df"][t].isna().sum()) for t in target_names},
            "feature_missing": feature_missing,
            "missing_imputed_cells": missing_imputed,
            "layout": {"header_row": args.header_row, "unit_row": args.unit_row, "data_row": args.data_row},
        }

        run_tag = getattr(args, "tag", None) or ("best" if best_info else "")
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
        # Recipe inputs that ml_train --best needs to replay a run (added in 2.4.0)
        if len(prepared) == 1:
            if g0["dtypes"]:
                history_record["dtypes"] = g0["dtypes"]
            if g0["positive_class"] is not None:
                history_record["positive_class"] = g0["positive_class"]
        try:
            history_record["data_fingerprint"] = data_fingerprint(raw_file)
        except OSError as e:
            sys.stderr.write(f"[WARN] 데이터 파일 지문 계산 실패 (보고서가 최고 버전을 다시 만들 수 있음): {e}\n")
        if best_info:
            history_record["best"] = True
            history_record["best_source"] = best_info["best_source"]
            history_record["target_recipes"] = {
                t: {"features": [f for f in target_group[t]["features"] if f not in target_group[t]["created"]],
                    "derived_evals": target_group[t]["evals"],
                    "dtypes": target_group[t]["dtypes"],
                    "positive_class": target_group[t]["positive_class"]}
                for t in target_names
            }
            next_commands = build_best_next_commands(raw_file, guard.final_dir, args, task)
        else:
            next_commands = build_next_commands(
                args=args, raw_file=raw_file, task=task, positive_class=g0["positive_class"],
                feature_names=g0["features"], created_features=g0["created"], applied_evals=g0["evals"],
                target_names=target_names, champion={**prim_champ, "weak_features": weak_feats}, schema=g0["schema"],
                mean_score=min_score, staging_dir=staging_dir, final_dir=guard.final_dir, data_root=data_root,
            )
        history_record["next_commands"] = next_commands

        # Commit Transaction: atomic rename .tmp_vX -> vX, then record it in the ledger.
        # The ledger never names a version folder that does not exist (a failed commit leaves no record).
        final_version_dir = guard.commit()
        append_history_ledger(ledger_path, history_record, on_recorded=guard.finalize)
        unverified = verify_version_dir(final_version_dir, target_names, ledger_path, current_version)
        for u in unverified:
            sys.stderr.write(f"[WARN] 저장 확인 실패: {u}\n")
        log_progress(100, f"학습 및 산출물 패키징 완료: {current_version} ({os.path.basename(final_version_dir)}) [소요: {duration_sec}초]")

        # LLM Directive
        short = "R2" if task == TASK_REGRESSION else "Macro-F1"
        target_brief = ", ".join([f"{t}({target_champions[t]['info']['champion_model']}: 검증 {short}={target_champions[t]['info']['val_score']:.4f})" for t in target_names])
        delta_txt = delta_score if delta_score == "base" else f"{delta_score:+.4f}"
        excl_brief = "; ".join(
            f"{t}: 학습 제외 {e['manual'] + e['auto']}행(직접 {e['manual']}, 자동 {e['auto']}), "
            f"검증 {short}(제외 행 포함)={e['all_rows_score'] if e['all_rows_score'] is not None else '-'}"
            for t, e in exclusion_summary.items() if e["manual"] or e["auto"])
        preview = (
            f"[{current_version} 완료 / {task}] 타겟: [{target_brief}] | "
            + (f"{excl_brief} | " if excl_brief else "") +
            f"평균 {metric_label}={mean_val_score:.4f} (Delta: {delta_txt}) | "
            f"소요시간: {duration_sec}s | "
            f"산출물 저장: {final_version_dir}"
        )
        agent_directive = (
            (f"UNVERIFIED: {current_version} 학습은 끝났지만 저장 결과를 확인하지 못했습니다 ({'; '.join(unverified)}). "
             f"재시도하지 말고 이 내용을 사용자에게 그대로 알리세요. " if unverified else "VERIFIED: ")
            + f"{current_version} competitive training completed. "
            f"Champion models and leaderboard packaged into '{os.path.basename(final_version_dir)}'. "
            f"Task={task}, Mean {metric_label}={mean_val_score:.4f}. "
            + ("Some rows were left out of training ('제외' column or automatic outliers): tell the user the counts and "
               "the score with excluded rows (검증 ... 제외 행 포함) from preview, and the excluded_rows_file path so they can review it. " if excl_brief else "")
            + ("Tell the user every item of warnings (duplicate rows). " if train_warnings else "") +
            "Next step: run next_commands[0] if the user wants to continue, or proceed to inference. "
            "A command with needs_input=true has a placeholder: replace it with a path the user gave, or ask the user first."
        )

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
            "verified": not unverified,
            "recommendations": recommendations,
            "next_commands": next_commands,
            "feature_types": history_record["feature_types"],
            "preview": preview,
            "agent_directive": agent_directive
        }
        if best_info:
            result["best_targets"] = {
                t: {"task": task, "source_version": src.get("version"), "source_score": src.get("score"),
                    "new_score": target_champions[t]["info"]["val_score"]}
                for t, src in best_info["best_source"].items()
            }
        if train_warnings:
            result["warnings"] = train_warnings
        if exclusion_summary:
            result["exclusions"] = exclusion_summary
            if excluded_rows_file:
                result["excluded_rows_file"] = excluded_rows_file
        if column_aliases:
            result["column_aliases"] = column_aliases
        if unverified:
            result["unverified"] = unverified
        if guard.incomplete_versions:
            result["incomplete_versions"] = guard.incomplete_versions
        return result

    except BaseException:
        # Includes KeyboardInterrupt (Ctrl+C) so the staging folder never leaks
        guard.rollback()
        raise
    finally:
        run_lock.release()


# ==============================================================================
# 8-B. Best Version (--best): every target retrained with its best recipe, one version per task
# ==============================================================================

BEST_TAG = "best"
TASK_ORDER = [TASK_REGRESSION, TASK_CLASSIFICATION]


REPORT_PURPOSE = {  # 2.9.2: a regression report also makes the performance calculator (ml_export.py)
    TASK_REGRESSION: "ML 결과 보고서(Word)와 성능예측표(엑셀 계산기) 생성 (모든 타겟, 입력을 줄이며 다시 학습해 몇 분 걸림. "
                     "성능예측표 제목은 학습 때 데이터 설명에서 자동으로 정함)",
    TASK_CLASSIFICATION: "ML 결과 보고서(Word) 생성 (이 데이터의 모든 타겟을 1개 문서로)",
}


def build_best_next_commands(raw_file: str, final_dir: str, args: argparse.Namespace,
                             task: str = TASK_REGRESSION) -> List[Dict[str, Any]]:
    """args carries the data layout (build_best_version fills it from the latest record of the task)."""
    here = os.path.dirname(os.path.abspath(__file__))
    return [
        {"purpose": REPORT_PURPOSE.get(task, REPORT_PURPOSE[TASK_CLASSIFICATION]),
         "command": f"python {_q(os.path.join(here, 'ml_report.py'))} -f {_q(raw_file)}",
         "needs_input": False},
        predict_next_command(final_dir, args, "새 데이터 예측 (<새데이터.xlsx>를 사용자가 준 실제 파일 경로로 바꿀 것)"),
    ]


def _record_task(rec: Dict[str, Any]) -> str:
    return rec.get("task", TASK_REGRESSION)


def _target_score(rec: Dict[str, Any], target: str) -> Optional[float]:
    champ = ((rec.get("targets_benchmark") or {}).get(target) or {}).get("champion") or {}
    v = champ.get("val_score")
    return float(v) if isinstance(v, (int, float)) else None


def _target_recipe(rec: Dict[str, Any], target: str) -> Dict[str, Any]:
    """Training conditions of one target in one ledger record (features before --eval, formulas, dtypes)."""
    stored = (rec.get("target_recipes") or {}).get(target)
    if stored:
        return {"features": list(stored.get("features") or []), "derived_evals": list(stored.get("derived_evals") or []),
                "dtypes": stored.get("dtypes"), "positive_class": stored.get("positive_class")}
    evals = list(rec.get("derived_evals") or [])
    derived = {e.split("=", 1)[0].strip() for e in evals if "=" in e}
    pos = rec.get("positive_class")
    if pos is None and _record_task(rec) == TASK_CLASSIFICATION:
        champ = ((rec.get("targets_benchmark") or {}).get(target) or {}).get("champion") or {}
        pos = (champ.get("decision_threshold") or {}).get("positive_class")
    return {"features": [f for f in rec.get("features_used", []) if f not in derived], "derived_evals": evals,
            "dtypes": rec.get("dtypes"), "positive_class": pos if _record_task(rec) == TASK_CLASSIFICATION else None}


def _version_intact(data_root: str, rec: Dict[str, Any]) -> bool:
    vdir = os.path.join(data_root, str(rec.get("version", "")))
    if not rec.get("version") or not os.path.isdir(vdir):
        return False
    need = ["leaderboard.json", "model_dataset.csv"] + [f"model_{sanitize_filename(t)}.json" for t in rec.get("targets", [])]
    return all(os.path.isfile(os.path.join(vdir, n)) for n in need)


def _unique_records(records: List[Dict[str, Any]]) -> List[Tuple[int, Dict[str, Any]]]:
    """(ledger index, record) with the last record per version (an --overwrite run replaces the folder)."""
    last: Dict[str, int] = {}
    for i, r in enumerate(records):
        last[str(r.get("version"))] = i
    return [(i, r) for i, r in enumerate(records) if last.get(str(r.get("version"))) == i]


def select_best_recipes(records: List[Dict[str, Any]], data_root: str, task: str) -> Dict[str, Dict[str, Any]]:
    """
    Per target: the record with the highest validation score (ties -> the later record) among versions
    whose folder is intact. A best record's recipe traces back to its original version (best_source).
    Returns {target: {recipe, selected_version, selected_score, source_version, source_score}} in
    first-appearance order.
    """
    chosen: Dict[str, Dict[str, Any]] = {}
    intact: Dict[str, bool] = {}
    for _, rec in _unique_records(records):
        if _record_task(rec) != task:
            continue
        ver = str(rec.get("version"))
        if ver not in intact:
            intact[ver] = _version_intact(data_root, rec)
            if not intact[ver]:
                sys.stderr.write(f"[WARN] 버전 폴더 {ver}이(가) 없거나 불완전해 최고 조건 후보에서 제외했습니다\n")
        if not intact[ver]:
            continue
        for t in rec.get("targets", []):
            s = _target_score(rec, t)
            if s is None:
                continue
            if t not in chosen or s >= chosen[t]["selected_score"]:
                src = (rec.get("best_source") or {}).get(t) if rec.get("best") else None
                chosen[t] = {
                    "recipe": _target_recipe(rec, t),
                    "selected_version": ver, "selected_score": s,
                    "source_version": (src or {}).get("version") or ver,
                    "source_score": (src or {}).get("score") if src and src.get("score") is not None else s,
                }
    return chosen


def _latest_layout(records: List[Dict[str, Any]], task: str) -> Dict[str, Any]:
    for rec in reversed(records):
        if _record_task(rec) == task:
            return (rec.get("data_profile") or {}).get("layout") or {}
    return {}


def build_best_version(args: argparse.Namespace, task: str, records: List[Dict[str, Any]]) -> Dict[str, Any]:
    raw_file = os.path.abspath(args.file)
    data_root = os.path.join(os.path.dirname(raw_file), f"{os.path.splitext(os.path.basename(raw_file))[0]}_ml_data")
    chosen = select_best_recipes(records, data_root, task)
    if not chosen:
        raise InputError(f"{task} 학습 기록 중 버전 폴더가 온전한 기록이 없어 최고 버전을 만들 수 없습니다. "
                         f"먼저 ml_train.py로 학습하십시오.")
    groups: Dict[str, Dict[str, Any]] = {}
    for t, c in chosen.items():
        r = c["recipe"]
        key = dumps_json([r["features"], r["derived_evals"], r["dtypes"], r["positive_class"]])
        g = groups.setdefault(key, {"features": r["features"], "targets": [], "eval": r["derived_evals"],
                                    "dtypes": r["dtypes"], "positive_class": r["positive_class"]})
        g["targets"].append(t)
        sys.stderr.write(f"[INFO] 최고 조건: '{t}' <- {c['selected_version']} (점수 {c['selected_score']:.4f}, "
                         f"원래 출처 {c['source_version']}), 피처 {len(r['features'])}개\n")
    layout = _latest_layout(records, task)
    run_args = argparse.Namespace(**vars(args))
    for k, default in (("seed", 42), ("description", None), ("tag", None), ("encoding", None)):
        if getattr(run_args, k, None) is None:
            setattr(run_args, k, default)
    for k in ("header_row", "unit_row", "data_row"):
        if getattr(run_args, k, None) is None:
            setattr(run_args, k, layout.get(k))
    run_args.version, run_args.overwrite = None, False
    best_source = {t: {"version": c["source_version"], "score": c["source_score"]} for t, c in chosen.items()}
    for t, c in chosen.items():
        if c["selected_version"] != c["source_version"]:
            best_source[t]["selected_version"] = c["selected_version"]
    log_progress(3, f"최고 버전 만들기 ({task}): 타겟 {len(chosen)}개, 학습 조건 {len(groups)}종을 현재 데이터로 다시 학습")
    return train_version(run_args, task, list(groups.values()), best_info={"best_source": best_source})


def _ledger_for(raw_file: str) -> Tuple[str, List[Dict[str, Any]]]:
    data_root = os.path.join(os.path.dirname(raw_file), f"{os.path.splitext(os.path.basename(raw_file))[0]}_ml_data")
    return data_root, load_history_ledger(os.path.join(data_root, "history.jsonl"))


def _tasks_to_build(records: List[Dict[str, Any]], task: Optional[str], ledger_path: str) -> List[str]:
    if not records:
        raise InputError(f"학습 이력이 없습니다 ({ledger_path}). 먼저 ml_train.py로 학습한 뒤 다시 실행하십시오.")
    present = [t for t in TASK_ORDER if any(_record_task(r) == t for r in records)]
    if task:
        if task not in present:
            raise InputError(f"학습 이력에 {task} 기록이 없습니다. [추천] --task {present[0]}")
        return [task]
    return present


def best_summary(version_result: Dict[str, Any]) -> Dict[str, Any]:
    return {"task": version_result["task"], "version": version_result["version"],
            "version_dir": version_result["version_dir"], "targets": version_result["targets"],
            "mean_val_score": version_result.get("mean_val_score"), "made": True,
            **({"warnings": version_result["warnings"]} if version_result.get("warnings") else {})}


def execute_make_best(args: argparse.Namespace) -> Dict[str, Any]:
    """ml_train.py --best: builds one best version per task (always retrains)."""
    raw_file = os.path.abspath(args.file)
    if not os.path.exists(raw_file):
        raise FileNotFoundError(f"입력 파일이 존재하지 않습니다: {raw_file}")
    data_root, records = _ledger_for(raw_file)
    tasks = _tasks_to_build(records, getattr(args, "task", None), os.path.join(data_root, "history.jsonl"))
    built = []
    for task in tasks:
        built.append(build_best_version(args, task, records))
        records = load_history_ledger(os.path.join(data_root, "history.jsonl"))
    return combine_best_results(built)


def combine_best_results(built: List[Dict[str, Any]]) -> Dict[str, Any]:
    """Result JSON of --best: the first version's training keys plus best_* keys covering every task."""
    result = dict(built[0])
    summaries = [best_summary(b) for b in built]
    best_targets: Dict[str, Any] = {}
    for b in built:
        best_targets.update(b.get("best_targets", {}))
    unverified = [u for b in built for u in b.get("unverified", [])]
    result.update({
        "made_best": True,
        "best_version": built[0]["version"],
        "best_version_dir": built[0]["version_dir"],
        "best_versions": summaries,
        "best_targets": best_targets,
        "verified": not unverified,
    })
    if unverified:
        result["unverified"] = unverified
    brief = "; ".join(f"{t}: {v['source_version']}({v['source_score']:.4f}) -> {v['new_score']:.4f}"
                      for t, v in best_targets.items() if isinstance(v.get("source_score"), (int, float)))
    vers = ", ".join(f"{s['version']}({s['task']}, 타겟 {len(s['targets'])}개)" for s in summaries)
    result["preview"] = f"[최고 버전 완료] {vers} | 타겟별 출처 -> 이번 점수: {brief}"
    multi = (f" 작업 유형이 {len(built)}가지라 최고 버전도 {len(built)}개입니다 (best_versions). 보고서도 {len(built)}개가 됩니다."
             if len(built) > 1 else "")
    head = (f"UNVERIFIED: 저장 결과를 확인하지 못했습니다 ({'; '.join(unverified)}). 재시도하지 말고 사용자에게 알리세요. "
            if unverified else "VERIFIED: ")
    result["agent_directive"] = (
        head + f"타겟별 최고 학습 조건으로 모든 타겟을 다시 학습해 {vers}을(를) 만들었습니다.{multi} "
        f"best_targets의 타겟별 출처 버전과 점수를 사용자에게 알리세요. 보고서는 next_commands[0], "
        f"예측은 ml_predict.py -m <best_version_dir>을 사용하세요."
    )
    return result


def best_is_fresh(records: List[Dict[str, Any]], data_root: str, raw_file: str, task: str) -> Optional[Dict[str, Any]]:
    """
    The latest best record of the task is reusable when no other training of that task was recorded after
    it, its version folder is intact, and the data file content is unchanged (SHA-256, else size + mtime).
    """
    idx = next((i for i in range(len(records) - 1, -1, -1)
                if records[i].get("best") and _record_task(records[i]) == task), None)
    if idx is None:
        return None
    best = records[idx]
    if any(_record_task(r) == task and not r.get("best") for r in records[idx + 1:]):
        sys.stderr.write(f"[INFO] {best.get('version')} 이후 새 학습 기록이 있어 최고 버전을 다시 만듭니다 ({task})\n")
        return None
    if not _version_intact(data_root, best):
        return None
    if not _fingerprint_same(best, raw_file):
        sys.stderr.write(f"[INFO] 데이터 파일 내용이 {best.get('version')} 때와 달라 최고 버전을 다시 만듭니다 ({task})\n")
        return None
    return best


def _fingerprint_same(rec: Dict[str, Any], raw_file: str) -> bool:
    fp_old = rec.get("data_fingerprint") or {}
    if not fp_old:
        return False
    try:
        fp_new = data_fingerprint(raw_file)
    except OSError:
        return False
    return (fp_old.get("sha256") == fp_new["sha256"]) if fp_old.get("sha256") else (
        fp_old.get("size") == fp_new["size"] and fp_old.get("mtime") == fp_new["mtime"])


def single_best_version(records: List[Dict[str, Any]], data_root: str, raw_file: str, task: str) -> Optional[Dict[str, Any]]:
    """
    The version that already holds the best of EVERY target of the task (one training run, or the latest best
    version when later runs beat nothing), trained on the current data file content. Retraining its recipes on the
    same data would give the same models, so the report uses it as is.
    """
    chosen = select_best_recipes(records, data_root, task)
    versions = {c["selected_version"] for c in chosen.values()}
    if len(versions) != 1:
        return None
    ver = versions.pop()
    rec = next((r for _, r in reversed(_unique_records(records)) if str(r.get("version")) == ver), None)
    if rec is None or not _fingerprint_same(rec, raw_file):
        return None
    return rec


def ensure_best_versions(args: argparse.Namespace) -> Dict[str, Any]:
    """
    For ml_report.py -f: reuses each task's up-to-date best version, otherwise builds it (in-process).
    Returns {"made_best", "best_versions": [...], "best_targets": {...}, "built": [training results]}.
    """
    raw_file = os.path.abspath(args.file)
    if not os.path.exists(raw_file):
        raise FileNotFoundError(f"입력 파일이 존재하지 않습니다: {raw_file}")
    data_root, records = _ledger_for(raw_file)
    tasks = _tasks_to_build(records, getattr(args, "task", None), os.path.join(data_root, "history.jsonl"))
    versions, best_targets, built = [], {}, []
    for task in tasks:
        rec = best_is_fresh(records, data_root, raw_file, task)
        if rec is not None:
            sys.stderr.write(f"[INFO] 최신 최고 버전 {rec['version']} ({task})을(를) 그대로 사용합니다\n")
        else:
            rec = single_best_version(records, data_root, raw_file, task)
            if rec is not None:
                sys.stderr.write(f"[INFO] {rec['version']}에 모든 타겟의 최고 점수가 있어 다시 학습하지 않고 그대로 사용합니다 ({task})\n")
        if rec is not None:
            versions.append({"task": task, "version": rec["version"],
                             "version_dir": os.path.join(data_root, rec["version"]), "targets": rec.get("targets", []),
                             "mean_val_score": rec.get("mean_val_score"), "made": False})
            sources = rec.get("best_source") if rec.get("best") else None
            for t in rec.get("targets", []):
                src = (sources or {}).get(t) or {"version": rec["version"], "score": _target_score(rec, t)}
                best_targets[t] = {"task": task, "source_version": src.get("version"), "source_score": src.get("score"),
                                   "new_score": _target_score(rec, t)}
            continue
        res = build_best_version(args, task, records)
        built.append(res)
        versions.append(best_summary(res))
        best_targets.update(res.get("best_targets", {}))
        records = load_history_ledger(os.path.join(data_root, "history.jsonl"))
    return {"made_best": bool(built), "best_versions": versions, "best_targets": best_targets, "built": built,
            "data_root": data_root}


# ==============================================================================
# 9. CLI Interface & History Mode
# ==============================================================================

def print_history_summary(data_root: str) -> Dict[str, Any]:
    """
    Prints a human-readable summary table of all experiments to stderr and returns
    a compact JSON summary (stdout stays a single JSON line for the agent).
    """
    ledger_path = os.path.join(data_root, "history.jsonl")
    records = load_history_ledger(ledger_path)
    if not records:
        return {
            "status": "no_history",
            "message": f"실험 원장이 존재하지 않습니다: {ledger_path}",
            "agent_directive": "VERIFIED: No experiments yet. Next step: run baseline training (v1).",
        }

    proj_desc = records[-1].get("description", "")
    out = sys.stderr

    out.write("\n" + "=" * 125 + "\n")
    out.write(f" [ML EXPERIMENT HISTORY LEDGER] Root: {data_root}\n")
    if proj_desc:
        out.write(f" [PROJECT DESCRIPTION] {proj_desc}\n")
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
        if r.get("best"):
            row["best"] = True  # built by --best (every target with its best recipe)
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

    # best_version: best record of the latest experiment (same task AND same targets as the latest run)
    latest = rows[-1]
    same = [x for x in rows if x["task"] == latest["task"] and sorted(x["targets"]) == sorted(latest["targets"])
            and isinstance(x["mean_val_score"], (int, float))]
    best = max(same, key=lambda x: x["mean_val_score"]) if same else None

    # best_by_task: best record per task (across all target sets)
    best_by_task: Dict[str, Dict[str, Any]] = {}
    for x in rows:
        if isinstance(x["mean_val_score"], (int, float)):
            cur = best_by_task.get(x["task"])
            if cur is None or x["mean_val_score"] > cur["mean_val_score"]:
                best_by_task[x["task"]] = {"version": x["version"], "mean_val_score": x["mean_val_score"]}
    return {
        "status": "ok",
        "data_root": data_root,
        "description": proj_desc,
        "record_count": len(rows),
        "best_version": best["version"] if best else None,
        "best_mean_val_score": best["mean_val_score"] if best else None,
        "best_by_task": best_by_task,
        "latest_version": rows[-1]["version"],
        "records": rows,
        "verified": True,
        "agent_directive": "VERIFIED: History summary returned. Do not re-read history.jsonl for the same information.",
    }


def main():
    parser = JsonArgumentParser(
        description="Autonomous Regression / Classification ML Engine with 5-Model Competitive Benchmark & Version Lineage"
    )
    parser.add_argument("--file", "-f", help="입력 데이터 파일 경로 (.xlsx, .csv)")
    parser.add_argument("--features", "-x", help="독립변수(피처) 목록 (쉼표 구분 문자열 또는 파일 경로: .txt, .json, .csv)")
    parser.add_argument("--targets", "-y", help="종속변수(타겟) 목록 (쉼표 구분 문자열 또는 파일 경로: .txt, .json, .csv)")
    parser.add_argument("--eval", help="방탄 AST 파생변수 생성 수식 (세미콜론 구분). 컬럼명은 전체 이름 또는 단위를 뺀 고유 이름")
    parser.add_argument("--tag", help="해당 회차 실험 의도 및 변경 사항 메모 태그 (예: '베이스라인_학습', '파생변수2종_추가')")
    parser.add_argument(
        "--description", "--desc", dest="description",
        help="데이터셋 도메인 및 모델링 목적 설명 (문자열 또는 파일 경로 .txt, .md 지원. 미지정 시 이전 회차 설명 자동 상속)"
    )
    parser.add_argument("--header-row", type=int, help="컬럼명 헤더 행 번호 (1-based, 미지정 시 상단 빈 행 자동 건너뜀)")
    parser.add_argument("--unit-row", type=int, help="단위 행 번호 (선택)")
    parser.add_argument("--data-row", type=int, help="데이터 시작 행 번호 (선택)")
    parser.add_argument("--dtypes", help="데이터 타입 수동 오버라이드 (예: '일시:datetime,설비코드:str')")
    parser.add_argument("--encoding", help="model_dataset.csv 저장 인코딩 (기본값: Windows Excel 최적화 cp949, 특수문자 발생 시 utf-8-sig 자동 폴백)")
    parser.add_argument("--version", "-v", help="버전 수동 지정 (미지정 시 자동 계산 v1, v2...)")
    parser.add_argument("--overwrite", action="store_true",
                        help="--version으로 지정한 버전 폴더가 이미 있으면 교체 (사용자 확인 후에만 사용)")
    parser.add_argument("--task", choices=[TASK_REGRESSION, TASK_CLASSIFICATION], default=None,
                        help="학습 유형: regression(기본값, 연속 수치 예측) 또는 classification(범주/라벨 예측). "
                             "--best와 함께 쓰면 그 작업 유형만")
    parser.add_argument("--positive-class", dest="positive_class",
                        help="이진 분류의 양성 클래스 (예: 불량). 지정 시 양성 F1이 최대가 되도록 판정 기준 확률을 자동 조정")
    parser.add_argument("--seed", type=int, default=42, help="재현용 난수 시드 (기본값: 42)")
    parser.add_argument("--history", action="store_true", help="누적 실험 이력 조회 모드 (표는 stderr, 요약 JSON은 stdout)")
    parser.add_argument("--best", action="store_true",
                        help="이력에서 타겟별 최고 점수의 학습 조건으로 모든 타겟을 현재 데이터로 다시 학습해 새 버전 1개(작업 유형별)를 만듦")

    args = parser.parse_args()

    if args.best and not args.history:
        if not args.file:
            parser.error("--best에는 --file(-f) 인자가 필요합니다.")
        ignored = [flag for flag, val in (("-x", args.features), ("-y", args.targets), ("--eval", args.eval),
                                          ("--version", args.version), ("--positive-class", args.positive_class),
                                          ("--dtypes", args.dtypes)) if val]
        if ignored:
            sys.stderr.write(f"[WARN] --best는 이력의 학습 조건을 쓰므로 {ignored}을(를) 무시합니다\n")
            args.features = args.targets = args.eval = args.version = args.positive_class = args.dtypes = None
        run_cli(lambda: execute_make_best(args), hint="Train at least once with ml_train.py before --best.")
        return

    if args.history:
        if not args.file:
            parser.error("--history 조회 시 --file 인자가 필요합니다.")
        raw_file = os.path.abspath(args.file)
        file_stem = os.path.splitext(os.path.basename(raw_file))[0]
        data_root = os.path.join(os.path.dirname(raw_file), f"{file_stem}_ml_data")
        run_cli(lambda: print_history_summary(data_root))
        return

    if not args.file or not args.targets or not args.features:
        parser.error("--file, --features(-x), --targets(-y)는 모델 학습에 필수적인 인자입니다.")

    run_cli(lambda: execute_ml_train(args), hint="Check column names or formula syntax.")


if __name__ == "__main__":
    main()
