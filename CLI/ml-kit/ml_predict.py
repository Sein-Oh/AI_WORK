# -*- coding: utf-8 -*-
"""
ml_predict.py - Zero-Office Ultra-Fast Standalone ML Inference Engine.
Supports Directory Targeting (v1/, v2/) or Specific Model/Meta File Targeting.
Handles Batch Files (.csv, .xlsx) or Inline Single/Multi Record Inferences.

DEPLOYMENT: this file + _ml_runtime.py (the shared runtime of ml-kit) + a trained version folder (vX/) is a complete
inference package. It imports no other ml-kit file, and no other ml-kit file imports it.
Requires only NumPy, Pandas, Scikit-learn, Joblib (+ openpyxl for .xlsx).

Designed for lightweight 26B-30B LLM agents with 1-turn declarative execution.
"""

from __future__ import annotations   # type hints never touch numpy/pandas (they may be missing: R5)

import sys
import os
import re
import io
import json
import difflib
import argparse
import warnings
import contextlib
from datetime import datetime
from typing import Dict, Any, List, Optional, Tuple

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
try:
    from _ml_runtime import (  # noqa: E402  shared runtime of ml-kit
        Cond, InputError, JsonArgumentParser, ask, command_setting, layout_retry, MAX_KEYS, NEAREST_N, SNAPSHOT_EXCL_PREFIX, SNAPSHOT_ROW_COL, TOOL_DIR,
        Table, _alias_key, _blank_mark, _cell, _cell_text, _crossings, _fmt, _is_data_value, _layout_frame,
        _looks_like_unit_row, _row_cells, _same_unit_name, _unit_shaped, _unit_text, align_categories,
        apply_inference_evals, atomic_write_text, avoid_overwrite, build_alias_map, category_candidates,
        check_comma_numbers, check_core_libraries, limit_threads, check_percent_mix, check_same_column, coerce_numeric,
        command_replacing, command_text, find_header_layout, fixed_category_value, inline_value, input_column_owner,
        known_columns, kv_text, layout_warnings, load_data_file, looks_inline_input, mark_committed, model_raw_inputs,
        name_list, nearest_rows, normalize_category, normalize_labels, note_recovered, np, numbered, option_retry,
        other_unit, parse_datetime, parse_inline_input, parse_where, pd, percent_warning, prepare_features, quote_arg,
        read_raw_table, read_text_file, resolve_column_names, resolve_model_arg, resolve_target_models, resolve_where,
        round_pred, run_cli, sanitize_filename, save_with_lock_fallback, setup_console, sniff_header_index, strip_unit,
        suggest_columns, unit_hint, unit_marked, unit_row_verdict, warn_long_path, where_mask, where_text,
        write_csv_excel_compatible,
    )
except ModuleNotFoundError as _exc:            # copied without the shared runtime: still 1 JSON line (ACI contract)
    if getattr(_exc, "name", "") != "_ml_runtime":
        raise
    sys.stdout.write(json.dumps({"status": "error", "error_type": "missing_file",
                                 "message": "ml_predict.py와 같은 폴더에 _ml_runtime.py가 없습니다. ml-kit 폴더의 _ml_runtime.py를 함께 복사하십시오.",
                                 "agent_directive": "STOP_RETRY: 다시 실행하지 말고 message를 사용자에게 그대로 알리십시오."},
                                ensure_ascii=False) + "\n")
    sys.exit(1)


# ##############################################################################
# PART P. INFERENCE ENGINE
# ##############################################################################

# ==============================================================================
# 1. Derived Feature Replay (shared AST sandbox, identical to training)
# ==============================================================================

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
            hints.append(f"-i의 '{c}'에 가까운 모델 입력: {ask(unit_marked(s, c), [command_replacing(c, x) for x in s], f'-i의 {c} 대신 {{}}')}. "
                         f"어느 것인지 사용자에게 물으십시오")
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
# 5. Core Prediction Engine
# ==============================================================================

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
        cmds = [command_replacing(c, m) for m in ms] if is_inline else []
        out.append(f"{where} '{c}'은(는) 여러 모델 입력과 비슷합니다. '{c}'을(를) 어느 이름으로 고칠지 사용자에게 물으십시오: "
                   f"{ask(unit_marked(ms, c), cmds, f'{c} 대신 {{}}')}")
    for m in missing:
        cands = [c for c in found[m] if c not in shared]
        if len(cands) == 1 and other_unit(cands[0], m):
            # '압력(MPa)' for the model's '압력(bar)': a rename would read the MPa value as bar (no [추천])
            out.append(unit_hint(cands[0], m, where))
        elif cands:
            shown = unit_marked(cands, m)
            cmds = [command_replacing(c, m) for c in cands] if is_inline else []
            out.append(f"'{m}'이(가) 없습니다. {where} {ask(shown, cmds, f'{{}}을(를) {m}(으)로')} 중 '{m}'(으)로 고칠 이름을 "
                       f"사용자에게 물으십시오")
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


def check_inline_numbers(df_input: pd.DataFrame, numeric: List[str]) -> None:
    """-i 숫자 입력의 값이 비었거나 숫자가 아니면 멈춤 (2.9.76): 그대로 두면 중앙값으로 채운 예측이 VERIFIED로 나가 사용자가 준 값처럼
    보고됨 (학습 결과의 '-i "온도=<값>,..."' 틀을 <값> 그대로 실행, '온도=200:280:10' 같은 범위). 자리표시자면 값을 물으라고 함."""
    bad: List[str] = []
    for col in numeric:
        if col not in df_input.columns:
            continue
        raw = df_input[col]
        text = raw.map(lambda v: "" if v is None or (isinstance(v, float) and np.isnan(v)) else str(v).strip())
        wrong = (text != "") & coerce_numeric(raw).isna()     # 빈 값은 원래대로(중앙값 채우기 확인), 글자만
        bad += [f"{col}={text[i]!r}" for i in raw.index[wrong]]
    if not bad:
        return
    placeholder = any("<" in b and ">" in b for b in bad)
    sweep = any(re.search(r"\d\s*:\s*\d", b) for b in bad)
    raise InputError(
        f"-i의 숫자 입력 값이 숫자가 아닙니다: {', '.join(bad[:10])}. 예측하지 않았습니다 (중앙값으로 채우지 않음). "
        + ("<값> 자리에 넣을 값을 사용자에게 물은 뒤 그 값으로 바꿔 실행하십시오." if placeholder else
           "값 여러 개를 차례로 예측하려면 --vary \"<입력>\" --sweep \"<시작>:<끝>:<간격>\"을 쓰십시오 (ml-predict 스킬)." if sweep else
           "사용자에게 숫자 값을 확인하십시오."))


def _check_prediction_file(args: argparse.Namespace, model_items: List[Dict[str, Any]]) -> None:
    """-f that does not exist (2.9.74, 집 PC opencode): '이름=값' text -> [추천] the command with -i; a file the agent
    meant to write for one condition -> say that a few conditions go in -i without a file, and give that command
    (next_commands, needs_input: the model inputs with <값>)."""
    path = str(args.file)
    if os.path.exists(path):
        return
    base = command_text(drop={"-f": "", "--file": "", "-i": "", "--input": ""})
    if looks_inline_input(path):
        cmd = f"{base} -i {quote_arg(inline_value(path))}" if base else None
        raise InputError("-f에는 데이터 파일 경로를 줍니다. '이름=값' 입력은 -i로 줍니다."
                         + (f" [추천] {cmd}" if cmd else ""))
    # a typo of a file in the same folder: [후보] with numbered full commands (2.9.75), the user picks
    folder = os.path.dirname(os.path.abspath(path))
    try:
        listing = os.listdir(folder)
        # 학습 데이터(옆에 <이름>_ml_data가 있음)는 새 데이터 후보가 아님 (2.9.76: 학습 파일 전체를 예측하는 명령을 [후보]로 줬음)
        names = [n for n in listing if os.path.splitext(n)[1].lower() in (".csv", ".xlsx", ".xlsm", ".xls", ".tsv", ".txt")
                 and f"{os.path.splitext(n)[0]}_ml_data" not in listing]
    except OSError:
        names = []
    near = [os.path.join(folder, n) for n in difflib.get_close_matches(os.path.basename(path), names, n=5, cutoff=0.6)]
    if near:
        text, cmds = numbered(near, [command_replacing(str(args.file), c) for c in near], "-f {}")
        raise InputError(f"입력 데이터 파일을 찾을 수 없습니다: {path}. 같은 폴더의 비슷한 파일: {text}. "
                         "어느 파일인지 사용자에게 물으십시오 (새 조건 몇 건이면 파일 없이 -i \"이름=값,...\"로 예측합니다).",
                         extra={"confirm_commands": cmds} if cmds else None)
    needed, _ = _model_inputs(model_items)
    template = ",".join(f"{c}=<값>" for c in needed)
    extra = {}
    if base and needed:
        extra["next_commands"] = [{"purpose": "새 조건 한 건 예측 (<값>을 사용자가 준 값으로 바꿈)",
                                   "command": f"{base} -i {quote_arg(template)}", "needs_input": True, "placeholder": "<값>"}]
    raise InputError(f"입력 데이터 파일을 찾을 수 없습니다: {path}. 새 조건 몇 건은 파일을 만들지 않고 "
                     f"-i \"이름=값,...\"로 예측합니다 (모델 입력: {', '.join(needed)}). "
                     "입력 값을 사용자에게 묻고 next_commands[0]의 <값>을 그 값으로 바꿔 실행하세요.", extra=extra,
                     error_type="needs_user_decision" if extra else None)


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
            f"사용자에게 물으십시오. " + ask([f"{r}행" for r in rows], [layout_retry(r, None) for r in rows], "머리글 {}"))


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


# ------------------------------------------------------------------------------
# Per-prediction explanation (2.9.72, -i with up to INLINE_SHOW_MAX records): the effect of each input on the
# prediction, measured against a baseline record (training median; the most common value for text inputs). The
# symmetric two-point estimate averages "take the input away from the record" (record -> its value replaced by the
# baseline) and "give it to the baseline" (baseline -> the record's value), so effects of two inputs that act together
# are split fairly. One batch prediction per model: 2 + 2 x inputs rows per record. Classification: the effect on the
# probability of the predicted class.
# ------------------------------------------------------------------------------
EXPLAIN_SHOWN = 5        # inputs per record in the JSON (largest effects first)
EXPLAIN_PREVIEW = 2      # raised / lowered inputs named in the preview line


def explain_baseline(meta: Dict[str, Any], version_dir: str, cache: Dict[str, Any]) -> Dict[str, Any]:
    """Baseline value of every raw model input: the training median (numbers), the most common training value
    (text inputs, from model_dataset.csv; the first known kind without it), the median date (model_dataset.csv)."""
    schema = meta.get("feature_schema") or {}
    cats = schema.get("categorical") or {}
    dates = schema.get("datetime") or {}
    stats = meta.get("feature_stats") or {}
    data = None
    base: Dict[str, Any] = {}
    for c in model_raw_inputs(meta):
        if c in cats or c in dates:
            if data is None:
                data = _load_snapshot(version_dir, cache)
            col = data[c] if data is not None and c in data.columns else None
            if c in cats:
                vals = normalize_category(col).dropna() if col is not None else pd.Series(dtype=object)
                if len(vals):
                    base[c] = vals.value_counts().index[0]
                elif cats[c]:
                    base[c] = cats[c][0]
            elif col is not None:
                dt = parse_datetime(col).dropna()
                if len(dt):           # as text, the way a date cell or -i value is read ('YYYY-MM-DD[ HH:MM]')
                    mid = dt.sort_values().iloc[len(dt) // 2]
                    base[c] = mid.strftime("%Y-%m-%d %H:%M" if (mid.hour or mid.minute) else "%Y-%m-%d")
        elif (stats.get(c) or {}).get("median") is not None:
            base[c] = float(stats[c]["median"])
    return base


def _model_scores(item: Dict[str, Any], rows: pd.DataFrame, label: Optional[str] = None) -> np.ndarray:
    """Predictions of rows (raw inputs): regression values, or the probability of `label` (classification)."""
    meta = item["meta"]
    df = apply_inference_evals(rows.copy(), meta.get("derived_evals", []))
    X = prepare_features(df, meta["feature_schema"])
    if label is None:
        return np.asarray(item["pipeline"].predict(X), dtype=float)
    classes = [str(c) for c in item["pipeline"].classes_]
    return np.asarray(item["pipeline"].predict_proba(X)[:, classes.index(label)], dtype=float)


def explain_predictions(item: Dict[str, Any], records: pd.DataFrame, base: Dict[str, Any],
                        labels: Optional[List[str]] = None) -> List[Dict[str, Any]]:
    """Per record: [{input, value, baseline, effect}] sorted by size (all inputs; the caller cuts)."""
    names = [c for c in model_raw_inputs(item["meta"]) if c in base and c in records.columns]
    out: List[Dict[str, Any]] = []
    if not names:
        return out
    schema = item["meta"].get("feature_schema") or {}
    records = records.copy()
    for c in names:                        # -i values are text cells: numbers as numbers (a number baseline fits in)
        if c not in (schema.get("categorical") or {}) and c not in (schema.get("datetime") or {}):
            records[c] = coerce_numeric(records[c])
    records = records.astype(object)
    blocks = []
    for i in range(len(records)):
        x = records.iloc[[i]].reset_index(drop=True)
        b = x.copy()
        for c in names:
            b.loc[0, c] = base[c]
        rows = [x, b]
        for c in names:
            xb = x.copy()
            xb.loc[0, c] = base[c]
            bx = b.copy()
            bx.loc[0, c] = x.loc[0, c]
            rows += [xb, bx]
        blocks.append(pd.concat(rows, ignore_index=True))
    for i, block in enumerate(blocks):
        f = _model_scores(item, block, labels[i] if labels else None)
        fx, fb = f[0], f[1]
        effects = []
        for j, c in enumerate(names):
            e = 0.5 * ((fx - f[2 + 2 * j]) + (f[3 + 2 * j] - fb))
            effects.append({"input": c, "value": _cell(records.iloc[i][c]), "baseline": _cell(base[c]),
                            "effect": round_pred(float(e))})
        effects.sort(key=lambda r: -abs(float(r["effect"])))
        out.append({"record": i + 1, "prediction": round_pred(float(fx)), "baseline_prediction": round_pred(float(fb)),
                    **({"class": labels[i]} if labels else {}), "effects": effects})
    return out


def explain_brief(target: str, rec: Dict[str, Any]) -> str:
    """One short line: '수율: 높인 입력: 온도 +3.2, 압력 +1.1 / 낮춘 입력: 속도 -0.8'."""
    scale = max(abs(float(rec["prediction"])), abs(float(rec["baseline_prediction"])), 1e-12)
    top = max((abs(float(r["effect"])) for r in rec["effects"]), default=0.0)
    eff = [r for r in rec["effects"] if abs(float(r["effect"])) > max(1e-6 * scale, 0.01 * top)]
    up = [r for r in eff if float(r["effect"]) > 0][:EXPLAIN_PREVIEW]
    down = [r for r in eff if float(r["effect"]) < 0][:EXPLAIN_PREVIEW]
    if not up and not down:
        return ""
    def _fx(r: Dict[str, Any]) -> str:
        return f"{strip_unit(r['input'])} {float(r['effect']):+.3g}"
    head = f"{strip_unit(target)}" + (f" '{rec['class']}' 확률" if rec.get("class") is not None else "")
    parts = ([f"높인 입력: {', '.join(_fx(r) for r in up)}"] if up else []) + \
            ([f"낮춘 입력: {', '.join(_fx(r) for r in down)}"] if down else [])
    return f"{head}: " + " / ".join(parts)


# ------------------------------------------------------------------------------
# Retrain alert (2.9.72, -f): when the new file holds the measured values of a target, the prediction error on those
# rows is compared with the validation score of training (model_<t>.json). Clearly worse -> a warning and
# next_commands = ml_train on the new file with the same inputs and targets. Rule (at least ACTUAL_MIN_ROWS rows):
# regression RMSE > ACTUAL_RMSE_RATIO x validation RMSE, or R2 more than ACTUAL_R2_DROP below validation R2 with RMSE
# above ACTUAL_RMSE_SOFT x validation RMSE (a narrow range of new values lowers R2 alone without larger errors);
# classification accuracy more than ACTUAL_ACC_DROP below validation accuracy.
# ------------------------------------------------------------------------------
ACTUAL_MIN_ROWS = 10
ACTUAL_RMSE_RATIO = 1.5
ACTUAL_RMSE_SOFT = 1.2
ACTUAL_R2_DROP = 0.1
ACTUAL_ACC_DROP = 0.1


def _actual_column(df: pd.DataFrame, target: str, inputs: List[str]) -> Optional[str]:
    cols = [str(c) for c in df.columns]
    if target in cols:
        return target
    same = [c for c in cols if c not in inputs and not c.startswith(("pred_", "proba_", "conf_", "map_"))
            and (strip_unit(c) == strip_unit(target) or _same_unit_name(c, target))]
    return same[0] if len(same) == 1 else None


def actual_check(item: Dict[str, Any], df_input: pd.DataFrame, df_result: pd.DataFrame) -> Optional[Dict[str, Any]]:
    """The error of the predictions on the rows of the new file that hold the target's measured value, next to the
    validation score of training. None when the file has no such column or no measured value."""
    meta = item["meta"]
    t = str(item["target"])
    col = _actual_column(df_input, t, model_raw_inputs(meta))
    if col is None or f"pred_{t}" not in df_result.columns:
        return None
    task = meta.get("task", "regression")
    out: Dict[str, Any] = {"column": col}
    if task == "classification":
        actual = normalize_labels(df_input[col]).to_numpy(dtype=object)
        pred = df_result[f"pred_{t}"].astype(str).to_numpy(dtype=object)
        ok = np.array([a is not None for a in actual])
        n = int(ok.sum())
        if not n:
            return None
        a, p = actual[ok].astype(str), pred[ok]
        acc = float(np.mean(a == p))
        f1s = []
        for c in sorted(set(a) | set(p)):
            tp = float(np.sum((a == c) & (p == c)))
            prec = tp / max(float(np.sum(p == c)), 1.0)
            rec = tp / max(float(np.sum(a == c)), 1.0)
            f1s.append(0.0 if prec + rec == 0 else 2 * prec * rec / (prec + rec))
        out.update({"rows": n, "accuracy": round(acc, 4), "f1_macro": round(float(np.mean(f1s)), 4)})
        val_acc = meta.get("accuracy")
        out["validation"] = {"accuracy": val_acc, "f1_macro": meta.get("val_f1_macro")}
        worse = n >= ACTUAL_MIN_ROWS and val_acc is not None and acc < float(val_acc) - ACTUAL_ACC_DROP
        out["text"] = (f"{strip_unit(t)} 정확도 {acc:.3f} (학습 검증 {float(val_acc):.3f})" if val_acc is not None
                       else f"{strip_unit(t)} 정확도 {acc:.3f}")
    else:
        actual = coerce_numeric(df_input[col]).to_numpy(dtype=float)
        pred = coerce_numeric(df_result[f"pred_{t}"]).to_numpy(dtype=float)
        ok = np.isfinite(actual) & np.isfinite(pred)
        n = int(ok.sum())
        if not n:
            return None
        err = pred[ok] - actual[ok]
        rmse = float(np.sqrt(np.mean(err ** 2)))
        mae = float(np.mean(np.abs(err)))
        var = float(np.sum((actual[ok] - actual[ok].mean()) ** 2))
        r2 = 1.0 - float(np.sum(err ** 2)) / var if n >= 2 and var > 0 else None
        out.update({"rows": n, "rmse": round_pred(rmse), "mae": round_pred(mae),
                    "r2": round(r2, 4) if r2 is not None else None})
        v_rmse, v_r2 = meta.get("rmse"), meta.get("val_r2")
        out["validation"] = {"rmse": v_rmse, "mae": meta.get("mae"), "r2": v_r2}
        worse = n >= ACTUAL_MIN_ROWS and v_rmse is not None and (
            rmse > ACTUAL_RMSE_RATIO * float(v_rmse)
            or (r2 is not None and v_r2 is not None and r2 < float(v_r2) - ACTUAL_R2_DROP
                and rmse > ACTUAL_RMSE_SOFT * float(v_rmse)))
        out["text"] = (f"{strip_unit(t)} RMSE {rmse:.4g}" + (f" (학습 검증 {float(v_rmse):.4g})" if v_rmse is not None else "")
                       + (f", R2 {r2:.3f}" + (f" (학습 검증 {float(v_r2):.3f})" if v_r2 is not None else "")
                          if r2 is not None else ""))
    out["worse"] = bool(worse)
    out["judged"] = n >= ACTUAL_MIN_ROWS
    if not out["judged"]:
        out["note"] = f"실제값이 있는 행이 {n}개로 {ACTUAL_MIN_ROWS}개보다 적어 학습 때와 비교해 판단하지 않음"
    return out


def _ledger_record(version_dir: str, version: Optional[str]) -> Dict[str, Any]:
    """The history.jsonl record of the version (its --dtypes, ...); {} when there is none (a copied version folder)."""
    path = os.path.join(os.path.dirname(os.path.abspath(version_dir)), "history.jsonl")
    rec: Dict[str, Any] = {}
    try:
        for line in read_text_file(path).splitlines():
            try:
                r = json.loads(line)
            except ValueError:
                continue
            if isinstance(r, dict) and r.get("version") == version:
                rec = r
    except OSError:
        pass
    return rec


def retrain_commands(args: argparse.Namespace, items: List[Dict[str, Any]], checks: Dict[str, Dict[str, Any]],
                     out_dir: str) -> List[Dict[str, Any]]:
    """ml_train on the new file with the inputs and targets of the models whose actual error is clearly worse (one
    command per training recipe; every target of that recipe with measured values in the file). Long input lists are
    written to <out_dir>/<file>_retrain_features[_<n>].txt and passed by path (like ml_train's next_commands)."""
    groups: Dict[Any, Dict[str, Any]] = {}
    for it in items:
        t = str(it["target"])
        if t not in checks:
            continue
        meta = it["meta"]
        evals = list(meta.get("derived_evals") or [])
        derived = {e.split("=", 1)[0].strip() for e in evals if "=" in e}
        feats = [f for f in meta.get("features_required", []) if f not in derived]
        pos = (meta.get("decision_threshold") or {}).get("positive_class") if meta.get("task") == "classification" else None
        key = (meta.get("task", "regression"), tuple(feats), tuple(evals), pos)
        g = groups.setdefault(key, {"targets": [], "worse": False, "meta": meta, "dir": os.path.dirname(it["joblib_path"])})
        g["targets"].append(checks[t]["column"])
        g["worse"] = g["worse"] or checks[t]["worse"]
    train_py = os.path.join(TOOL_DIR, "ml_train.py")
    stem = os.path.splitext(os.path.basename(args.file))[0]
    cmds: List[Dict[str, Any]] = []
    for k, ((task, feats, evals, pos), g) in enumerate(groups.items()):
        if not g["worse"]:
            continue
        texts = list(evals) + ([pos] if pos is not None else []) + list(g["targets"])
        if any('"' in str(x) for x in texts):
            continue                     # cannot be quoted on the command line
        if len(feats) > 15 or any('"' in f for f in feats):
            fpath = os.path.join(out_dir, f"{stem}_retrain_features{'' if k == 0 else f'_{k + 1}'}.txt")
            atomic_write_text(fpath, "\n".join(feats) + "\n")
            x_arg = quote_arg(fpath)
        else:
            x_arg = quote_arg(",".join(feats))
        parts = [f"python {quote_arg(train_py)}", "-f", quote_arg(os.path.abspath(args.file)), "-x", x_arg,
                 "-y", quote_arg(",".join(g["targets"]))]
        for flag, val in (("--header-row", args.header_row), ("--unit-row", args.unit_row), ("--data-row", args.data_row)):
            if val is not None:
                parts += [flag, quote_arg(val)]
        dtypes = _ledger_record(g["dir"], g["meta"].get("version")).get("dtypes")
        if dtypes and '"' not in str(dtypes):
            parts += ["--dtypes", quote_arg(dtypes)]
        if evals:
            parts += ["--eval", quote_arg("; ".join(evals))]
        if task == "classification":
            parts += ["--task", "classification"]
            if pos is not None:
                parts += ["--positive-class", quote_arg(pos)]
        desc = str(g["meta"].get("description") or "").strip()
        if desc and '"' not in desc and "\n" not in desc and len(desc) <= 300:
            parts += ["--desc", quote_arg(desc)]
        parts += ["--tag", quote_arg("새데이터_재학습")]
        cmds.append({"purpose": f"새 데이터({os.path.basename(args.file)})로 같은 입력·타겟({', '.join(g['targets'])})을 다시 학습 "
                                "(실제값으로 본 오차가 학습 때보다 커짐)",
                     "command": " ".join(parts), "needs_input": False})
    return cmds


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
        _check_prediction_file(args, model_items)
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
            check_inline_numbers(df_input, sorted(num_raw))
            holders = [f"{c}={v}" for c in df_input.columns for v in df_input[c].astype(str) if re.fullmatch(r"\s*<[^<>]*>\s*", v)]
            if holders:                       # 글자·날짜 입력의 <값>도 (2.9.76 2차 검토: 범주 입력은 VERIFIED로 나갔음)
                raise InputError(f"-i에 자리표시자가 그대로 있습니다: {', '.join(holders[:10])}. 예측하지 않았습니다. "
                                 "<값> 자리에 넣을 값을 사용자에게 물은 뒤 그 값으로 바꿔 실행하십시오.")
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

    # 3-B. Explanation of each record (-i, up to INLINE_SHOW_MAX records) and the actual error on a file that holds
    # the measured target values (2.9.72). Neither may break the prediction.
    explain_lines: List[str] = []
    if is_inline and len(df_input) <= INLINE_SHOW_MAX:
        for item in model_items:
            t = str(item["target"])
            try:
                base = explain_baseline(item["meta"], os.path.dirname(item["joblib_path"]), snapshot_cache)
                labels = (df_result[f"pred_{t}"].astype(str).tolist()
                          if item["meta"].get("task") == "classification" else None)
                recs = explain_predictions(item, df_input, base, labels)
            except Exception as e:
                map_notes.append(f"타겟 '{t}': 입력별 영향(explain) 계산 실패 ({type(e).__name__}: {str(e)[:80]})")
                continue
            if not recs:
                continue
            predictions_summary[t]["explain"] = [
                {**{k: v for k, v in r.items() if k != "effects"}, "effects": r["effects"][:EXPLAIN_SHOWN],
                 "effects_count": len(r["effects"])} for r in recs]
            line = explain_brief(t, recs[0])
            if line:
                explain_lines.append(line)
    actual_checks: Dict[str, Dict[str, Any]] = {}
    if not is_inline:
        for item in model_items:
            try:
                chk = actual_check(item, df_input, df_result)
            except Exception as e:
                map_notes.append(f"타겟 '{item['target']}': 실제값 비교(actual_check) 실패 ({type(e).__name__}: {str(e)[:80]})")
                continue
            if chk:
                actual_checks[str(item["target"])] = chk

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

    # retrain alert: the new file's measured values are clearly worse predicted than in training
    retrain_next: List[Dict[str, Any]] = []
    worse = [t for t, c in actual_checks.items() if c["worse"]]
    for t in worse:
        file_warnings.append(f"타겟 '{t}': 새 파일의 실제값 {actual_checks[t]['rows']}행으로 본 오차가 학습 때보다 큽니다 "
                             f"({actual_checks[t]['text']}). 데이터 경향이 바뀌었을 수 있어 새 데이터로 다시 학습하기를 권합니다")
    if worse:
        try:
            retrain_next = retrain_commands(args, model_items, actual_checks,
                                            os.path.dirname(saved_file) if saved_file else os.path.dirname(os.path.abspath(args.file)))
        except Exception as e:
            map_notes.append(f"다시 학습 명령을 만들지 못함 ({type(e).__name__}: {str(e)[:80]})")

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

    def _case_lines() -> List[str]:
        """직접 입력 몇 건(파일 저장 없음): 사용자에게 그대로 전할 건별 예측 글 (2.9.78, 집 PC r5: 지시문이 길어 Gemma가 뒤쪽의
        '실측 비교값 함께'를 놓침 -> 알릴 값을 preview에 다 넣고 지시문은 'preview 그대로' 한 줄)."""
        out: List[str] = []
        for t, ps in predictions_summary.items():
            preds = ps.get("predictions") or []
            for k, pred in enumerate(preds):
                tag = f"{t}" + (f" {k + 1}번" if len(preds) > 1 else "")
                if ps.get("task") == "classification":
                    probs = (ps.get("probabilities") or [{}])[k] if k < len(ps.get("probabilities") or []) else {}
                    out.append(f"{tag} {pred}" + (f" (확률 {probs.get(pred)})" if probs.get(pred) is not None else ""))
                    continue
                text = f"{tag} {_fmt(float(pred))}"
                iv = (ps.get("intervals") or [])[k] if k < len(ps.get("intervals") or []) else None
                extra = [f"90% 범위 {_fmt(float(iv[0]))} ~ {_fmt(float(iv[1]))}"] if iv else []
                mc = (ps.get("map_check") or [])[k] if k < len(ps.get("map_check") or []) else None
                if mc and mc.get("value") is not None:
                    extra.append(f"실측 비교값 {_fmt(float(mc['value']))} (차이 {float(mc['diff']):+.4g})")
                out.append(text + (f" ({', '.join(extra)})" if extra else ""))
        return out

    case_lines = _case_lines() if is_inline and not saved_file and not inline_truncated else []
    targets_brief = ", ".join(_brief(t) for t in predictions_summary)
    preview = (
        (f"[예측 완료] 입력: {source_desc} | 예측: " + "; ".join(case_lines) if case_lines else
         f"[예측 완료] 입력: {source_desc} | 예측 요약: [{targets_brief}]")
        + (f" | 저장: {os.path.basename(saved_file)}" if saved_file else "")
        + (f" | 외삽 경고 {len(drift_warnings)}건 (학습 범위 밖 입력이라 예측은 참고용)" if drift_warnings else "")
        + (f" | 저확신 경고 {len(low_conf_warnings)}건" if low_conf_warnings else "")
        + (f" | 결측 대치 {sum(missing_counts.values())}건" if missing_counts else "")
        + (f" | 미학습 범주 경고 {len(category_warnings)}건" if category_warnings else "")
        + (f" | 실측 비교 경고 {len(map_warnings)}건" if map_warnings else "")
        + (" | 입력별 영향 " + " ; ".join(explain_lines) if explain_lines else "")
        + (" | 실제값 비교: " + "; ".join(c["text"] + (" -> 학습 때보다 나쁨, 재학습 권장" if c["worse"] else
                                                     " (비슷함)" if c["judged"] else " (행이 적어 판단 안 함)")
                                       for c in actual_checks.values()) if actual_checks else "")
    )
    warn_keys = ([k for k, on in (("extrapolation_warnings", drift_warnings), ("low_confidence_warnings", low_conf_warnings),
                                   ("missing_input_warnings", missing_counts), ("unseen_category_warnings", category_warnings),
                                   ("map_check_warnings", map_warnings), ("warnings", file_warnings)) if on])
    short = (                        # 직접 입력 몇 건: 알릴 것은 preview에 다 있음, 지시문은 짧게 (2.9.78)
        ("VERIFIED: " if not unverified else "UNVERIFIED: ")
        + f"타겟 {len(model_items)}개를 예측했습니다. preview의 예측값, 90% 범위, 실측 비교값과 입력별 영향('높인 입력'·'낮춘 입력')을 "
          "글 그대로 사용자에게 알리세요 (화살표 같은 기호로 바꾸지 않음)."
        + (f" 경고({', '.join(warn_keys)})도 함께 알리세요." if warn_keys else "")
        + (f" extrapolation_warnings는 {len(drift_warnings)}건 중 {WARNINGS_SHOW_MAX}건만 있으니 전체 건수도 알리세요."
           if len(drift_warnings) > WARNINGS_SHOW_MAX else "")
        + (f" 확인하지 못한 것: {'; '.join(unverified)}." if unverified else "")
    ) if case_lines and not retrain_next else None
    agent_directive = short or (
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
        + ("preview의 입력별 영향(explain: 학습 기준값에서 바꿀 때 예측을 높인 입력과 낮춘 입력)을 '높인 입력'·'낮춘 입력' 글 그대로 알리세요 "
           "(화살표 같은 기호로 바꾸지 않음). " if explain_lines else "")
        + ("새 파일에 실제값이 있어 예측 오차를 학습 때와 비교했습니다 (actual_check). preview의 실제값 비교를 알리세요. "
           if actual_checks else "")
        + ("오차가 학습 때보다 크게 나빠졌습니다: warnings를 알리고, 새 데이터로 다시 학습하려면 next_commands[0]"
           "(같은 입력·타겟으로 새 파일을 학습)을 실행합니다. " if retrain_next else "")
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
    if actual_checks:
        result["actual_check"] = actual_checks
        result["retrain_recommended"] = bool(worse)
    if retrain_next:
        result["next_commands"] = retrain_next
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
        raise InputError(f"'{vary}'은(는) 숫자 입력이 아니라 반대 찾기로 바꿀 수 없습니다. "
                         + ask(numeric_in, [command_setting("--vary", n) for n in numeric_in], "--vary {}"))
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
                             f"{category_candidates(schema['categorical'][f], f)}")
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
                             f"{category_candidates(cats[f], f)}")
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
        raise InputError(f"{option}에 바꿔 볼 숫자 입력을 주십시오. "
                         + ask(numeric, [command_setting(option, n) for n in numeric], f"{option} {{}}"))
    names = resolve_column_names(given, numeric, f"바꿔 볼 입력({option})")[0]
    bad = [n for n in names if n not in stats]
    if bad:
        ok = [c for c in numeric if c in stats]
        raise InputError(f"{bad}의 학습 범위가 모델 정보에 없어 바꿔 볼 수 없습니다. "
                         + ask(ok, [command_replacing(bad[0], c) if len(bad) == 1 else None for c in ok], f"{bad[0]} 대신 {{}}"))
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


SUGGEST_SHEET = "실험계획"


def _save_table(args: argparse.Namespace, df: pd.DataFrame, default_name: str, what: str,
                sheet: Optional[str] = None) -> Tuple[str, List[str]]:
    """Writes the result table (--save .xlsx/.csv, else <data folder>/<default_name>.xlsx; `sheet`: the sheet name)
    and reads it back (R1: rows, and the sheet name)."""
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
        path = save_with_lock_fallback(path, lambda p: df.to_excel(p, index=False, sheet_name=sheet or "Sheet1"), what=what)
    mark_committed(f"{what} {os.path.basename(path)}", output_file=path)
    unverified: List[str] = []
    try:
        back = load_data_file(path, header_row=1)
        if len(back) != len(df):
            unverified.append(f"저장한 표를 다시 읽으니 {len(back)}행 (기대 {len(df)}행)")
        if sheet and path.lower().endswith(".xlsx"):
            import openpyxl
            wb = openpyxl.load_workbook(path, read_only=True)
            try:
                if sheet not in wb.sheetnames:
                    unverified.append(f"저장한 표에 '{sheet}' 시트가 없음 ({wb.sheetnames})")
            finally:
                wb.close()
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
        raise InputError(f"{option}에는 바꿔 볼 입력 --vary \"<입력1>,<입력2>\"가 필요합니다. "
                         + ask(numeric, [command_setting("--vary", n) for n in numeric], "--vary {}"))
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
# 5-D. Next experiment suggestion: --suggest N (2.9.72)
# ==============================================================================
#
# N input combinations inside the training range of the varied inputs (--vary, default every numeric model input not
# fixed with -i; the other numeric inputs at -i or the training median) chosen from SUGGEST_POOL random candidates:
#   - without --maximize/--minimize: where the 90% prediction interval is widest and measured rows are far (space
#     filling): new experiments there add the most information;
#   - with --maximize/--minimize: a good predicted value plus half the interval (an upper confidence bound), so the
#     points are good and still informative.
# The 90% interval of a point is the champion's out-of-fold residual interval scaled by the residuals of the nearest
# training rows (SUGGEST_NEIGHBORS, ratio clipped to SUGGEST_LOCAL_CLIP) and by the distance to the nearest measured
# row (sqrt((1 + (d / d_typ)^2) / 2), d_typ = the typical distance between training rows, so a point as far as two
# measured rows usually are keeps the trained width). Points are picked one by one; a picked point counts as
# measured for the next ones, so they spread out. Bounded work (pool x reference rows), no time limit split needed.

SUGGEST_MAX = 30
SUGGEST_POOL = 3000
SUGGEST_REF_ROWS = 5000        # training rows (sampled) used for distances and local residuals
SUGGEST_NEIGHBORS = 10
SUGGEST_LOCAL_CLIP = (0.5, 3.0)
SUGGEST_DIST_MAX = 5.0         # distance factor cap
SUGGEST_MIN_GAP = 0.02         # picked points closer than this (inputs scaled to the training range) are one point
SUGGEST_SHOWN = 10             # points in the JSON (the table holds all)


def _suggest_step(span: float, whole: bool) -> float:
    step = 10.0 ** np.floor(np.log10(span / 200.0)) if span > 0 else 1.0
    return max(step, 1.0) if whole else step


def _sq_dist(A: np.ndarray, B: np.ndarray) -> np.ndarray:
    d = (A ** 2).sum(axis=1)[:, None] + (B ** 2).sum(axis=1)[None, :] - 2.0 * A @ B.T
    return np.maximum(d, 0.0)


def execute_suggest(args: argparse.Namespace) -> Dict[str, Any]:
    n_want = int(args.suggest)
    if n_want < 1 or n_want > SUGGEST_MAX:
        fixed = str(min(max(n_want, 1), SUGGEST_MAX))
        raise InputError(f"--suggest는 1~{SUGGEST_MAX}개입니다 (지금 {n_want}). [추천] {option_retry(fixed, '--suggest')}")
    goal_text = args.maximize or args.minimize
    sense = 1.0 if args.maximize else -1.0
    filt = [t.strip() for t in args.targets.split(",") if t.strip()] if args.targets else None
    items, skipped = _regression_items(args, filt)
    names = [str(it["target"]) for it in items]
    goal = None
    if goal_text:
        option = "--maximize" if args.maximize else "--minimize"
        goal = resolve_column_names([str(goal_text).strip()], names, f"목표({option})")[0][0]
    raw, numeric, cats, stats = _model_input_info(items)
    if args.vary:
        vary = _vary_names(args, args.vary, numeric, stats, "--vary")
    else:     # the tool decides: every numeric input with a training range that -i does not fix
        with contextlib.redirect_stderr(io.StringIO()):
            given = _base_inputs(args, items, [])["fixed"] if args.input else {}
        vary = [c for c in numeric if c in stats and c not in given and float(stats[c]["max"]) > float(stats[c]["min"])]
        if not vary:
            raise InputError("바꿔 볼 숫자 입력이 없습니다 (모든 숫자 입력이 -i로 고정되었거나 학습 범위가 한 값). "
                             "-i에서 바꿔 볼 입력을 빼십시오.")
    info = _base_inputs(args, items, vary)
    conds = None
    if args.where:
        conds = resolve_where(parse_where(args.where, "--where"), names + raw, "--where")
    vdir = _version_dir_of(items)
    data = _load_snapshot(vdir, {})
    keys = vary + [c for c in info["numeric"] if c not in vary and c in info["base"] and c in stats]
    lows = np.array([float(stats[c]["min"]) for c in keys])
    highs = np.array([float(stats[c]["max"]) for c in keys])
    spans = np.maximum(highs - lows, 1e-12)
    k = len(vary)
    rng = np.random.default_rng(0)
    warnings_out = list(info["warnings"])

    # reference rows (measured, scaled) and the out-of-fold residuals of each target on them
    ref = None
    rows_no = None
    resid: Dict[str, np.ndarray] = {}
    if data is not None and all(c in data.columns for c in keys):
        R = np.column_stack([coerce_numeric(data[c]).to_numpy(dtype=float) for c in keys])
        ok = np.isfinite(R).all(axis=1)
        idx = np.flatnonzero(ok)
        if len(idx) > SUGGEST_REF_ROWS:
            idx = np.sort(rng.choice(idx, SUGGEST_REF_ROWS, replace=False))
        if len(idx) >= 2:
            ref = (R[idx] - lows) / spans
            rows_no = (coerce_numeric(data[SNAPSHOT_ROW_COL]).fillna(-1).astype(int).to_numpy()[idx]
                       if SNAPSHOT_ROW_COL in data.columns else idx + 1)
            for it in items:
                t = str(it["target"])
                path = os.path.join(vdir, f"oof_{it['meta'].get('target_safe_name', sanitize_filename(t))}.csv")
                if not os.path.isfile(path):
                    continue
                try:
                    oof = load_data_file(path, header_row=1)
                    pos = coerce_numeric(oof["data_row"]).to_numpy(dtype=float) - 1
                    res = np.abs(coerce_numeric(oof["residual"]).to_numpy(dtype=float))
                    full = np.full(len(data), np.nan)
                    good = np.isfinite(pos) & (pos >= 0) & (pos < len(data))
                    full[pos[good].astype(int)] = res[good]
                    resid[t] = full[idx]
                except Exception:
                    pass
    if ref is None:
        warnings_out.append("학습 데이터(model_dataset.csv)를 읽지 못해 실측 행과의 거리 없이 제안끼리만 고르게 퍼뜨렸습니다 "
                            "(90% 범위는 학습 때 폭 그대로)")
    if ref is not None and len(ref) > 1:
        probe = ref[rng.choice(len(ref), min(500, len(ref)), replace=False)]
        dd = _sq_dist(probe, ref)
        dd[dd <= 1e-18] = np.inf                     # itself (and exact duplicates)
        nn = np.sqrt(dd.min(axis=1))
        nn = nn[np.isfinite(nn)]
        d_typ = float(np.median(nn)) if len(nn) else 0.0
    else:
        d_typ = 0.0
    if d_typ <= 0:
        d_typ = 1.0 / max(SUGGEST_POOL, 1) ** (1.0 / max(len(keys), 1))

    # candidates: random inside the training range of the varied inputs, rounded to a readable step
    whole = []
    for c in vary:
        w = False
        if data is not None and c in data.columns:
            v = coerce_numeric(data[c]).dropna().to_numpy(dtype=float)
            w = bool(len(v)) and bool(np.all(np.isclose(v, np.round(v))))
        whole.append(w)
    steps = np.array([_suggest_step(float(spans[j]), whole[j]) for j in range(k)])
    U = rng.random((SUGGEST_POOL, k))
    pts = lows[:k] + U * spans[:k]
    pts = np.clip(np.round(pts / steps) * steps, lows[:k], highs[:k])
    pts = np.unique(pts, axis=0)
    rows = pd.DataFrame({**{c: [v] * len(pts) for c, v in info["base"].items()}, **{v: pts[:, j] for j, v in enumerate(vary)}})
    frame = rows.copy()
    for it in items:
        frame[str(it["target"])] = _predict_rows(it, rows)
    feasible = where_mask(frame, conds) if conds else np.ones(len(pts), dtype=bool)
    if not feasible.any():
        raise InputError(f"학습 범위 안에서 --where 조건({where_text(conds)})을 만족하는 조합을 찾지 못했습니다. "
                         "사용자에게 그대로 알리고 조건을 바꿀지 물으십시오.")
    full = np.column_stack([(pts - lows[:k]) / spans[:k]]
                           + [np.full((len(pts), 1), (float(info["base"][c]) - lows[j]) / spans[j])
                              for j, c in enumerate(keys) if j >= k])

    # distance to the nearest measured row and the local residual ratio of each target
    d_ref = np.full(len(pts), np.inf)
    near_i = np.zeros(len(pts), dtype=int)
    ratio = {str(it["target"]): np.ones(len(pts)) for it in items}
    if ref is not None:
        nb = min(SUGGEST_NEIGHBORS, len(ref))
        for s in range(0, len(pts), 500):
            dd = _sq_dist(full[s:s + 500], ref)
            near_i[s:s + 500] = np.argmin(dd, axis=1)
            d_ref[s:s + 500] = np.sqrt(dd[np.arange(len(dd)), near_i[s:s + 500]])
            part = np.argpartition(dd, nb - 1, axis=1)[:, :nb]
            for t, r in resid.items():
                g = np.nanquantile(r, 0.9) if np.isfinite(r).any() else np.nan
                loc = np.nanquantile(np.where(np.isfinite(r[part]), r[part], np.nan), 0.9, axis=1) \
                    if np.isfinite(r).any() else np.full(len(part), np.nan)
                q = np.where(np.isfinite(loc) & (g > 0), loc / g if g > 0 else 1.0, 1.0)
                ratio[t][s:s + 500] = np.clip(q, *SUGGEST_LOCAL_CLIP)

    def dist_factor(d: np.ndarray) -> np.ndarray:
        return np.minimum(np.sqrt((1.0 + (d / d_typ) ** 2) / 2.0), SUGGEST_DIST_MAX)

    ivs = {str(it["target"]): it["meta"].get("prediction_interval") for it in items}
    goal_half = None
    if goal:
        iv = ivs.get(goal)
        goal_half = (float(iv["residual_upper"]) - float(iv["residual_lower"])) / 2.0 if iv else 0.0
        if not goal_half or goal_half <= 0:
            goal_half = float(np.std(frame[goal].to_numpy(dtype=float))) or 1.0
    d_now = d_ref.copy() if ref is not None else np.full(len(pts), 1.0)
    picked: List[int] = []
    for _ in range(min(n_want, int(feasible.sum()))):
        fac = dist_factor(d_now)
        width = np.mean([ratio[t] for t in ratio], axis=0) * fac
        if goal:
            score = sense * frame[goal].to_numpy(dtype=float) / goal_half + ratio[goal] * fac
        else:
            score = width
        score = np.where(feasible, score, -np.inf)
        if picked:
            gap = np.sqrt(_sq_dist(full, full[picked]).min(axis=1))
            score = np.where(gap < SUGGEST_MIN_GAP, -np.inf, score)
        i = int(np.argmax(score))
        if not np.isfinite(score[i]):
            break
        picked.append(i)
        d_now = np.minimum(d_now, np.sqrt(_sq_dist(full, full[[i]])[:, 0]))

    points: List[Dict[str, Any]] = []
    table_rows: List[Dict[str, Any]] = []
    for n, i in enumerate(picked, 1):
        fac = float(dist_factor(np.array([d_ref[i]]))[0]) if ref is not None else 1.0
        p: Dict[str, Any] = {"no": n, "inputs": {v: _r6(pts[i, j]) for j, v in enumerate(vary)},
                             "predictions": {t: round_pred(float(frame[t].iloc[i])) for t in names},
                             "interval_90": {}, "interval_width": {}}
        tr: Dict[str, Any] = {"번호": n, **{v: _r6(pts[i, j]) for j, v in enumerate(vary)},
                              **{c: (_r6(v) if isinstance(v, (int, float, np.floating)) else v)
                                 for c, v in info["base"].items()}}
        for t in names:
            iv = ivs.get(t)
            y = float(frame[t].iloc[i])
            tr[f"pred_{t}"] = round_pred(y)
            if iv:
                m = float(ratio[t][i]) * fac
                lo, hi = y + float(iv["residual_lower"]) * m, y + float(iv["residual_upper"]) * m
                p["interval_90"][t] = [round_pred(lo), round_pred(hi)]
                p["interval_width"][t] = round_pred(hi - lo)
                tr[f"pred_lo_{t}"], tr[f"pred_hi_{t}"], tr[f"폭_{t}"] = round_pred(lo), round_pred(hi), round_pred(hi - lo)
        if ref is not None:
            p["nearest_row"] = int(rows_no[near_i[i]])
            p["nearest_distance"] = round(float(d_ref[i]), 4)
            tr["가까운실측행"], tr["실측거리"] = p["nearest_row"], p["nearest_distance"]
        points.append(p)
        table_rows.append(tr)
    table = pd.DataFrame(table_rows)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out, unverified = _save_table(args, table, f"suggest_{stamp}", "실험 계획 표", sheet=SUGGEST_SHEET)
    unverified += _recheck_solutions(items, info["base"], vary, points, conds, goal or names[0])
    if skipped:
        warnings_out.append(f"분류 타겟 {skipped}은(는) 실험 제안에서 뺐습니다 (숫자 타겟만)")
    if len(picked) < n_want:
        warnings_out.append(f"서로 다른 조합을 {len(picked)}개만 찾았습니다 (요청 {n_want}개): 학습 범위가 좁거나 조건을 "
                            "만족하는 조합이 적습니다")
    mode = (f"{goal} {'최대' if sense > 0 else '최소'} 쪽 + 불확실성" if goal
            else "불확실한 곳 (예측 90% 범위가 넓고 실측과 먼 곳)")
    shown = points[:SUGGEST_SHOWN]
    res: Dict[str, Any] = {
        "status": "ok", "command": "suggest", "mode": "maximize" if goal and sense > 0 else "minimize" if goal else "explore",
        "goal": goal, "vary": vary, "where": where_text(conds) if conds else None, "targets": names,
        "fixed": {c: (float(v) if isinstance(v, (int, float, np.floating)) else v) for c, v in info["fixed"].items()},
        "filled_defaults": {c: _r6(v) for c, v in info["filled"].items()},
        "training_range": {v: [_r6(lows[j]), _r6(highs[j])] for j, v in enumerate(vary)},
        "interval_method": "학습 OOF 잔차의 90% 범위 x 가까운 실측 행의 잔차 비 x 실측 행과의 거리 보정 (추정)",
        "count": len(points), "shown": len(shown), "truncated": len(shown) < len(points), "points": shown,
        "output_file": out, "verified": not unverified,
    }
    if len(shown) < len(points):
        res["full_result_file"] = out
    if warnings_out:
        res["warnings"] = warnings_out
    if unverified:
        res["unverified"] = unverified

    def _brief(p: Dict[str, Any]) -> str:
        t = goal or names[0]
        s = ", ".join(f"{strip_unit(v)}={_fmt(x)}" for v, x in p["inputs"].items())
        s += f" -> {strip_unit(t)} {_fmt(p['predictions'][t])}"
        if t in p["interval_width"]:
            s += f" (90% 폭 {_fmt(p['interval_width'][t])})"
        return s
    res["preview"] = (f"다음 실험 제안 {len(points)}개 [{mode}]: "
                      + " | ".join(f"{p['no']}) {_brief(p)}" for p in points[:3])
                      + (f" | 외 {len(points) - 3}개" if len(points) > 3 else "") + f" | 표: {os.path.basename(out)}")
    res["agent_directive"] = (
        ("VERIFIED: 표를 다시 읽고 제안 조건을 다시 예측해 확인했습니다. " if not unverified else
         f"UNVERIFIED: 확인하지 못한 것이 있습니다 ({'; '.join(unverified)}). 다시 실행하지 말고 unverified를 그대로 알리세요. ")
        + "preview(제안 조건과 예측, 90% 범위 폭)와 output_file(실험계획 표) 경로를 알리세요. "
        + (f"points는 {len(points)}개 중 {len(shown)}개이고 전체는 output_file에 있습니다. " if len(shown) < len(points) else "")
        + ("warnings(중앙값으로 고정한 입력 등)도 알리세요. " if warnings_out else "")
        + "실험으로 확인할 후보 조건이지 결과가 보장된 조건이 아니라는 점을 덧붙이세요.")
    return res


# ==============================================================================
# 6. CLI Entry Point
# ==============================================================================

def main():
    setup_console()
    check_core_libraries("예측")
    limit_threads(0)             # prediction: one thread is enough even for a large file (MANY_THREADS_ROWS)
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
    parser.add_argument("--suggest", type=int,
                        help=f"다음 실험 제안: 학습 범위 안에서 예측이 가장 불확실하고 실측과 먼 조건 N개 (1~{SUGGEST_MAX}). "
                             "--maximize/--minimize와 함께 쓰면 좋은 예측값과 불확실성을 함께 봄. --vary 생략 시 모든 숫자 입력, "
                             "-i로 고정, 표(실험계획 시트) 저장")

    args = parser.parse_args()

    search = args.maximize or args.minimize
    if args.suggest is not None:
        if args.find or args.sweep or args.file:
            parser.error("실험 제안(--suggest)에는 --find, --sweep, -f를 함께 쓰지 않습니다.")
        if args.maximize and args.minimize:
            parser.error("--maximize와 --minimize는 하나만 주십시오.")
        if args.where and not search:
            parser.error("--where는 --maximize 또는 --minimize와 함께 씁니다.")
        if search and args.targets:
            parser.error("--suggest에 --maximize/--minimize를 주면 -y를 쓰지 않습니다 (목표 타겟은 --maximize/--minimize 값).")
        run_cli(lambda: execute_suggest(resolve_model_arg(args)), hint="확인할 것: --vary, -i, 목표 타겟의 이름.")
        return
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
