# -*- coding: utf-8 -*-
"""
ml-kit 표 보간 / 반대 찾기 (ml_map.py)

Reads a data table (.xlsx / .csv) directly (no Excel) and answers two questions with linear interpolation:
  forward : values of the other columns at given key values      --input "반응온도=232.5,소비전력=60"
  reverse : key value where a column reaches a target value      --find "생산효율=50" --vary 반응온도 [--input fixed keys]
Batch    : --points <file> with key columns -> result workbook.

Data shape is decided automatically from the numeric key columns (1-3):
  grid      : every combination of the key values exists -> classic multilinear interpolation
              (nearest lower/upper value per key, weighted by the distance ratio); outside the table the end
              segments are extended linearly (extrapolation warning).
  scattered : any other table (process records). 1 key: the two neighbouring rows; 2-3 keys: the measured rows
              of the enclosing simplex (Delaunay, keys scaled to 0-1). Outside the measured area a least-squares
              trend line (all rows) is used instead, never the two end points (unstable on noisy data).
Every answer lists the rows it used (Excel row numbers) and their weights, plus the 5 nearest measured rows.
The data file is never modified. Rows with text in the '제외' column (other than '포함') are not used.

Shared runtime (data loading, column matching, CLI protocol) is imported from _ml_runtime.py.
"""
import contextlib
import glob
import io
import json
import re
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _ml_runtime import (  # noqa: E402  shared runtime of ml-kit
    check_core_libraries, exit_missing_library,
    setup_console, run_cli, JsonArgumentParser, InputError, ColumnNameError, resolve_column_names, coerce_numeric,
    percent_warning, check_percent_mix, percent_mix, _strip_percent, name_list,
    load_data_file, layout_warnings, parse_inline_input, strip_unit, save_with_lock_fallback, mark_committed, avoid_overwrite, warn_long_path,
    check_comma_numbers, check_same_column, suggest_columns, build_alias_map, known_columns, numbered, command_replacing, _alias_key, _same_unit_name, kv_text, quote_arg, header_layout_hint, EXCLUDE_COL, INCLUDE_MARK, MAX_KEYS, CATEGORY_CANDIDATES_MAX, Table, MultiTable, nearest_rows, _is_numeric_col, _cell, _fmt, _crossings,
    _trained_data_files, model_dir_from_data, model_raw_inputs, read_text_file, ask, parse_datetime,
)


def _model_find_command(args, goal_col: str, goal_val, vary: str, fixed: dict) -> "Tuple[str, List[str]]":
    """(ml_predict.py --find 전체 명령, 모델 입력이 아니라 뺀 --input 열) of the same question on the model trained from this
    data file; ("", []) when there is no trained model, the goal is no regression target or --vary no model input.
    Names are the model's own (the data file read with its unit row has '반응온도(°C)' for the model's '반응온도')."""
    ap = os.path.abspath(args.file)
    if ap not in _trained_data_files(os.path.dirname(ap)):
        return "", []
    try:
        with contextlib.redirect_stderr(io.StringIO()):
            vdir = model_dir_from_data(ap)
    except Exception:  # noqa: BLE001  no single version to use: no hint (the usual message stays)
        return "", []
    meta = None
    for path in glob.glob(os.path.join(vdir, "model_*.json")):
        try:
            m = json.loads(read_text_file(path))
        except (OSError, ValueError):
            continue
        if _alias_key(str(m.get("target", ""))) == _alias_key(goal_col):
            meta = m
    if meta is None or meta.get("task", "regression") != "regression":
        return "", []
    inputs = {_alias_key(c): c for c in model_raw_inputs(meta)}
    if _alias_key(vary) not in inputs:
        return "", []

    def show(v) -> str:
        try:
            return _fmt(float(v))
        except (TypeError, ValueError):
            return str(v)
    keep = {inputs[_alias_key(c)]: v for c, v in fixed.items() if _alias_key(c) in inputs}
    dropped = [c for c in fixed if _alias_key(c) not in inputs]
    parts = [f"python {quote_arg(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'ml_predict.py'))}",
             "--model", quote_arg(ap), "--find", quote_arg(f"{meta['target']}={show(goal_val)}"),
             "--vary", quote_arg(inputs[_alias_key(vary)])]
    if keep:
        parts += ["--input", quote_arg(kv_text({c: show(v) for c, v in keep.items()}))]
    return " ".join(parts), dropped


def _model_find_hint(args, goal_col: str, goal_val, vary: str, fixed: dict) -> str:
    """--find의 기준 열이 표 보간 상한(MAX_KEYS)을 넘을 때, 이 데이터 파일로 학습한 모델로 같은 질문을 푸는 ml_predict.py --find
    전체 명령 (2.9.80 집 PC Gemma P3: 입력 6개로 ml_map을 돌려 멈추고 사용자에게 범위를 물음). 이름은 모델 입력 이름
    (사용자가 쓴 오타·단위 표기를 옮기지 않음). 예측은 읽기만 함.
    - 목표가 모델의 회귀 타겟이고 --vary와 --input가 모두 모델 입력이면 [추천].
    - --input에 모델 입력이 아닌 열이 있으면 그 열을 뺀 명령은 다른 질문일 수 있어 [후보] (사용자가 정함).
    - 학습한 모델이 없거나, 목표가 타겟이 아니거나, --vary가 모델 입력이 아니면 빈 글."""
    cmd, dropped = _model_find_command(args, goal_col, goal_val, vary, fixed)
    if not cmd:
        return ""
    if not dropped:
        return f" 이 파일로 학습한 모델이 있어 모델로 찾는 명령: [추천] {cmd}"
    return (f" 이 파일로 학습한 모델로 찾을 수 있지만 {', '.join(repr(c) for c in dropped)}은(는) 모델 입력이 아니라 빼야 합니다 "
            f"(고정 조건이 달라짐). 사용자에게 물으십시오: "
            + ask([f"{', '.join(dropped)}을(를) 빼고 모델로 찾기"], [cmd], "{}"))


setup_console()

from datetime import datetime  # noqa: E402
from typing import Any, Dict, List, Optional, Tuple  # noqa: E402

check_core_libraries("표 보간")
try:  # R5: a missing library answers with 1 JSON line (missing_library + install_command)
    import numpy as np  # noqa: E402
    import pandas as pd  # noqa: E402
except ImportError as _import_exc:
    exit_missing_library(_import_exc, "표 보간")

MAX_SOLUTIONS_JSON = 20   # more reverse-lookup answers than this go to a result workbook
SCAN_STEPS = 2000         # reverse lookup sampling steps on scattered data with 2-3 keys


# ==============================================================================
# 1. Table loading
# ==============================================================================

def load_table(args) -> Dict[str, Any]:
    path = os.path.abspath(args.file)
    if not os.path.isfile(path):
        raise FileNotFoundError(f"데이터 파일이 없습니다: {path}")
    warn_long_path(path)
    df = load_data_file(path, header_row=args.header_row, unit_row=args.unit_row, data_row=args.data_row)
    unit_note = []
    below = df.attrs.get("unit_row_below")
    if below and args.header_row is None and args.unit_row is None and args.data_row is None:
        # 머리글 바로 아래 단위 행 (검토 B C1): 데이터로 읽으면 작은 표는 숫자 열이 없어지고 큰 표는 이름에 단위가 없음.
        # 학습·예측처럼 단위 행으로 읽는다 (판단 대행, warnings에 알림)
        df = load_data_file(path, header_row=int(df.attrs.get("header_row") or 1), unit_row=int(below))
        unit_note = [f"머리글 아래 {int(below)}행을 단위 행으로 읽어 열 이름에 단위를 붙였습니다 "
                     f"(예: {', '.join(str(c) for c in list(df.columns)[:3])})"]
    rows = df.attrs.get("source_rows") or list(range(1, len(df) + 1))
    decimal_comma = bool(df.attrs.get("decimal_comma"))     # --points CSVs are read the same way
    df = df.copy()
    df["__row"] = rows
    excl = next((c for c in df.columns if strip_unit(str(c)) == EXCLUDE_COL), None)
    dropped = 0
    if excl is not None:
        text = df[excl].map(lambda v: "" if pd.isna(v) else str(v).strip())
        mask = (text != "") & (text != INCLUDE_MARK)
        dropped = int(mask.sum())
        df = df[~mask]
    notes = unit_note + layout_warnings(df)     # repeated header rows (dropped by the loader), totals rows
    for w in notes:
        sys.stderr.write(f"[WARN] {w}\n")
    return {"path": path, "df": df.reset_index(drop=True), "exclude_col": excl, "excluded": dropped,
            "decimal_comma": decimal_comma,
            "layout_warnings": notes,
            "saved_at": datetime.fromtimestamp(os.path.getmtime(path)).strftime("%Y-%m-%d %H:%M")}


def data_columns(df: pd.DataFrame, exclude_col: Optional[str]) -> List[str]:
    return [str(c) for c in df.columns if c != "__row" and c != exclude_col]


# ==============================================================================
# 3. Argument handling
# ==============================================================================

def _parse_pairs(text: str, what: str) -> Tuple[Dict[str, Any], Optional[Dict[str, str]]]:
    """Names and values of one 'k=v,...' row, plus the 'k=v' text form (None for JSON) for check_comma_numbers."""
    if not text:
        return {}, None
    df = parse_inline_input(text)
    if len(df) != 1:
        raise InputError(f"{what}은(는) 한 줄('열=값,열=값')로 주십시오.")
    return {str(k): v for k, v in df.iloc[0].to_dict().items()}, df.attrs.get("kv_pairs")


def _map_command(args, **change: Optional[str]) -> str:
    """The same ml_map.py command with a corrected --input / --find text (absolute, quoted paths)."""
    parts = [f"python {quote_arg(os.path.abspath(__file__))}", "--file", quote_arg(os.path.abspath(args.file))]
    for flag, key in (("--header-row", "header_row"), ("--unit-row", "unit_row"), ("--data-row", "data_row")):
        if getattr(args, key, None) is not None:
            parts += [flag, str(getattr(args, key))]
    for flag, key in (("--input", "input"), ("--targets", "targets"), ("--find", "find"), ("--vary", "vary")):
        val = change.get(key, getattr(args, key, None))
        if val:
            parts += [flag, quote_arg(val)]
    if getattr(args, "save", None):
        parts += ["--save", quote_arg(args.save)]
    if getattr(args, "overwrite", False):
        parts.append("--overwrite")
    return " ".join(parts)


def _number(v: Any, error: str, option: str, given: str, pct: Optional[List[str]] = None) -> float:
    """A number given in --input / --find: '%' is taken off (10%, not 0.1, as coerce_numeric reads data cells; said in pct),
    nan and inf are no numbers (검토 B A3·D2: 'nan'은 빈 값처럼, 'inf'는 내부 오류로 이어짐)."""
    text = str(v).strip()
    bare = _strip_percent(text)
    try:
        num = float(bare)
    except (TypeError, ValueError):
        raise InputError(error)
    if not np.isfinite(num):
        raise InputError(error)
    if pct is not None and bare != text:
        pct.append(f"{option} '{given}'의 값 '{text}'은(는) %를 떼고 숫자로 읽었습니다 ({text} -> {bare}, "
                   f"100으로 나누지 않음). 열의 단위가 %인지 확인하십시오.")
    return num


def _split_keys(df: pd.DataFrame, pairs: Dict[str, Any], columns: List[str], args=None,
                kv: Optional[Dict[str, str]] = None, reserved: Optional[List[str]] = None,
                pct: Optional[List[str]] = None) -> Tuple[Dict[str, float], Dict[str, str]]:
    """Resolves --input names; numeric columns need numeric values, text columns become row filters. reserved (--targets, the
    --find goal, --vary) are never suggested as a key. A value written with '%' ('10%') is read without it as
    coerce_numeric reads data cells (10, not 0.1); pct collects the percent_warning of such --input values."""
    names, _ = resolve_column_names(list(pairs), columns, "기준(--input)" if args is not None else "조회 파일(--points) 기준",
                                    layout=df.attrs, values=True, exclude=reserved)
    check_same_column(list(pairs), names, "--input" if args is not None else "--points 조회 파일 열",
                      pairs if args is not None else None)
    resolved = dict(zip(pairs, names))
    if kv and args is not None:
        check_comma_numbers(kv, lambda k: k in resolved and _is_numeric_col(df[resolved[k]]), "--input",
                            lambda fixed: _map_command(args, input=kv_text(fixed)))
    numeric, categorical = {}, {}
    for given, col in zip(pairs, names):
        v = pairs[given]
        if _is_numeric_col(df[col]):
            numeric[col] = _number(v, f"기준 열 '{col}'은(는) 숫자 열인데 값 '{v}'이(가) 숫자가 아닙니다.", "--input", given,
                                   pct)
        else:
            categorical[col] = str(v).strip()
    return numeric, categorical


_ASK_POINTS = (" --targets가 없어 이 열이 기준 열인지 구할 열의 실측값·메모인지 알 수 없어 사용자에게 물으십시오: 기준 열이면 조회 파일의 "
               "머리글을 고르신 데이터 열 이름으로 고치고, 실측값·메모면 --targets에 구할 열을 넣어 다시 실행하십시오 (그 열의 실측값 "
               "열은 결과에 그대로 붙습니다).")


def _ask_without_y(e: InputError, ask: bool) -> InputError:
    """Without --targets every numeric column not used as a key is a value column, so a points column near one of them may be
    its measured value: the user says which (the --targets guidance). Only for an unknown column name (ColumnNameError of
    resolve_column_names), never for another error (a key value that is no number, two names of one column)."""
    if not ask or not isinstance(e, ColumnNameError):
        return e
    head, sep, tail = str(e).partition("\n")
    return ColumnNameError(f"{head}{_ASK_POINTS}{sep}{tail}")


def _is_exact_key(name: str, cols: List[str]) -> bool:
    """name is a column of cols exactly: the same name, or the same name and unit spaced otherwise."""
    return name in cols or sum(_same_unit_name(name, c) for c in cols) == 1


def _is_key_name(name: str, cols: List[str]) -> bool:
    """name is a column of cols as --input matches it: an exact name (_is_exact_key) or a unique name written without
    unit."""
    if _is_exact_key(name, cols):
        return True
    return strip_unit(name) == name.strip() and _alias_key(name) in build_alias_map(cols)


def _points_keys(points_df: pd.DataFrame, keycols: List[str], targets: List[str]) -> Tuple[List[Any], List[Any]]:
    """(key columns, measured columns) of a --points file. A column of the --targets target's name ('생산량(실측)' for '생산량')
    is a measured or memo column, never a key, unless it names a non-target column exactly: '수율' next to the
    columns '수율(1차)' and '수율(최종)' with --targets 수율(최종) is no unique name of '수율(1차)' (keycols leave the target
    out). Other columns of a data column's name without unit are keys (a different unit stops in _split_keys)."""
    tkeys = {_alias_key(t) for t in targets}
    bases = {strip_unit(x) for x in keycols}
    usable, measured = [], []
    for c in points_df.columns:
        name = str(c)
        if re.fullmatch(r"col_\d+", name):
            continue
        if _is_exact_key(name, keycols):
            usable.append(c)
        elif _alias_key(name) in tkeys:
            measured.append(c)
        elif _is_key_name(name, keycols):
            usable.append(c)
        elif strip_unit(name) in bases:
            usable.append(c)
    return usable, measured


def _check_near_points(points_df: pd.DataFrame, skip: List[Any], unused: List[str], layout: Dict[str, Any],
                       ask: bool = False, used: Optional[List[str]] = None, targets: Optional[List[str]] = None) -> None:
    """A numeric --points column that is no key but nearly matches a data column not used as a key ('반응 압력(bar)'
    for '반응압력(bar)', '압력(MPa)' for '반응압력(bar)'): the result would look as if the pressure was used. The same
    message as --input ([후보] / unit_hint). Columns like no data column (memo, ID) pass through.
    Near means the same name without unit or a close spelling of the whole name; the comparison of the names without
    unit only builds the [후보] list (as the gate it stopped measured columns: '수율_실측' / '수율(1차)').
    The check comes first for every column: a column holding a --targets target's name ('수율_실측', '목표수율' for --targets
    '수율(최종)') passes through as a measured value or memo only when no unused column is near (a key column
    '냉각수 온도(°C)' with --targets '온도' is near '냉각수온도' and stops). A column and a near column that both hold a --targets
    target's name without unit (targets) are compared without it: '수율_실측' / '함수율' share only '수율' and pass
    ('_실측' / '함'), '함수율%' / '함수율' stay near ('함%' / '함'). unused never holds a --targets target; ask (no --targets): the --targets
    guidance. used (the keys and --targets columns) are never listed."""
    tnames = [_alias_key(t) for t in (targets or []) if _alias_key(t)]

    def still_near(col: str, cand: str) -> bool:
        a, b = re.sub(r"\s+", "", col).lower(), re.sub(r"\s+", "", cand).lower()
        whole = (a, b)
        for t in tnames:
            if t in a and t in b:
                a, b = a.replace(t, ""), b.replace(t, "")
        if (a, b) == whole:
            return True
        # 남은 부분은 기호·괄호를 빼고 비교 ('_상' / '(상)', '1' / '(1단)'은 같은 위치·단의 키 열)
        a, b = re.sub(r"[\W_]", "", a), re.sub(r"[\W_]", "", b)
        if not (a or b):
            return True
        if a and b and (a in b or b in a):
            return True
        return bool(a and b and suggest_columns(a, [b], bare=False))

    near = [str(c) for c in points_df.columns
            if c not in skip and not re.fullmatch(r"col_\d+", str(c)) and _is_numeric_col(points_df[c])
            and any(still_near(str(c), u) for u in suggest_columns(str(c), unused, bare=False))]
    if not near:
        return
    try:
        resolve_column_names(near, unused + [c for c in (used or []) if c not in unused], "조회 파일(--points)",
                             layout=layout, values=True, exclude=used)
    except ColumnNameError as e:
        head, sep, tail = str(e).partition("\n")
        raise _ask_without_y(ColumnNameError(
            f"{head} 이 열은 기준 열로 쓰이지 않아 결과가 그 값을 반영하지 않습니다. 조회 파일의 머리글을 데이터 "
            f"열 이름으로 고친 뒤 같은 명령을 다시 실행하십시오.{sep}{tail}"), ask)


def _date_text(ts: Any) -> str:
    """A date cell as the user writes it: '2026-01-06', with the time only when it is not midnight."""
    if pd.isna(ts):
        return ""
    ts = pd.Timestamp(ts)
    return str(ts.date()) if ts == ts.normalize() else ts.strftime("%Y-%m-%d %H:%M")


def _is_date_col(s: pd.Series) -> bool:
    """Date cells (an Excel date column, read as datetime64 or as Timestamp objects)."""
    if pd.api.types.is_datetime64_any_dtype(s):
        return True
    filled = [v for v in s if not pd.isna(v)]
    return bool(filled) and all(isinstance(v, (pd.Timestamp, datetime)) for v in filled)


def _filter(df: pd.DataFrame, categorical: Dict[str, str]) -> pd.DataFrame:
    for col, given in categorical.items():
        v = given
        if _is_date_col(df[col]):
            # 날짜 열 (검토 B C2): '2026-01-06'과 칸의 '2026-01-06 00:00:00'을 날짜로 비교, 후보도 날짜 글로
            vals = df[col].map(_date_text)
            want = parse_datetime(pd.Series([given], dtype=object)).iloc[0]
            if not pd.isna(want):
                v = _date_text(want)
        else:
            vals = df[col].map(lambda s: "" if pd.isna(s) else str(s).strip())
        if not (vals == v).any():
            known = sorted(set(vals) - {""})[:CATEGORY_CANDIDATES_MAX]
            text, cmds = numbered(known, [command_replacing(given, k) for k in known], f"'{given}' 대신 '{{}}'")
            raise InputError(f"'{col}' 열에 '{given}' 값이 없습니다. {text}", extra={"confirm_commands": cmds} if cmds else None)
        df = df[vals == v]
    return df


def _value_columns(df: pd.DataFrame, args_y: Optional[str], columns: List[str], skip: List[str]) -> List[str]:
    """Value columns, numeric judged on the whole table (df before the --input text filter): a column that is empty
    in the filtered rows is still a value column and MultiTable leaves it out with a warning."""
    if args_y:
        names, _ = resolve_column_names([s.strip() for s in args_y.split(",") if s.strip()], columns, "구할 열(--targets)",
                                        layout=df.attrs)
        bad = [c for c in names if not _is_numeric_col(df[c])]
        if bad:
            raise InputError(f"구할 열 {bad}은(는) 숫자 열이 아닙니다.")
        return [c for c in names if c not in skip]
    vals = [c for c in columns if c not in skip and _is_numeric_col(df[c])]
    if not vals:
        raise InputError("기준 열 말고 계산할 숫자 열이 없습니다.")
    return vals


def _default_save(base_dir: str, stem: str) -> str:
    return os.path.join(base_dir, f"{stem}_보간_{datetime.now():%Y%m%d_%H%M%S}.xlsx")


def _resolve_save(path: Optional[str], data_path: str) -> Optional[str]:
    if not path:
        return None
    path = path if os.path.isabs(path) else os.path.join(os.path.dirname(data_path), path)
    if not path.lower().endswith(".xlsx"):
        path += ".xlsx"
    return path


def _write_book(path: str, sheets: Dict[str, pd.DataFrame], overwrite: bool = False) -> Tuple[str, bool]:
    path = avoid_overwrite(path, overwrite, what="결과 파일")

    def _w(p: str) -> None:
        with pd.ExcelWriter(p) as xw:
            for name, frame in sheets.items():
                frame.to_excel(xw, sheet_name=name, index=False)
    real = save_with_lock_fallback(path, _w, what="결과 파일")
    mark_committed(f"결과 파일 {os.path.basename(real)}", output_file=real)
    try:
        back = pd.read_excel(real, sheet_name=None)
        ok = all(len(back[n]) == len(f) for n, f in sheets.items())
    except Exception:
        ok = False
    return real, ok


def _basis_text(basis: List[Dict[str, Any]]) -> str:
    return " + ".join(f"{b['row']}행 x {b['weight']:.4g}" for b in basis)


def _basis_differs(r: Dict[str, Any]) -> bool:
    by_col = r.get("basis_by_column")
    return bool(by_col) and len({_basis_text(b) for b in by_col.values()}) > 1


def _basis_cells(r: Dict[str, Any], per_column: bool) -> Dict[str, str]:
    """Excel '근거' cells: one column, or one per value column (MultiTable) when any result of the run used
    different rows per column (per_column is decided once per run, so a sheet never mixes the two layouts)."""
    by_col = r.get("basis_by_column")
    if per_column and by_col:
        return {f"근거_{c}": _basis_text(b) for c, b in by_col.items()}
    return {"근거": _basis_text(r["basis"])}


FAR_FIND = 0.3            # --find: an answer outside the data range by more than this share of its span is no answer
TWO_VALUE_GAP = 0.1       # two places where a key takes one value are separate answers when this share of the span apart


def _drop_single_keys(sub: pd.DataFrame, numeric: Dict[str, float], warnings: List[str], option: str = "--input",
                      need_one: bool = True) -> None:
    """A numeric key with one value only in the rows used (검토 B B5: a constant key made the Delaunay step fail and its qhull
    text reached the message): given that very value, it says nothing and is left out (warned); another value cannot be
    computed from this table (stop)."""
    for k in list(numeric):
        vals = np.unique(coerce_numeric(sub[k]).dropna().to_numpy(dtype=float))
        if len(vals) != 1:
            continue
        if not np.isclose(numeric[k], vals[0], rtol=1e-9, atol=1e-12):
            raise InputError(f"기준 열 '{k}'은(는) 표에 {_fmt(vals[0])} 하나뿐이라 {_fmt(numeric[k])}에서의 값을 계산할 수 없습니다. "
                             f"{option}의 '{k}' 값을 사용자에게 확인하십시오.")
        del numeric[k]
        warnings.append(f"기준 열 '{k}'은(는) 표에 {_fmt(vals[0])} 하나뿐이라 기준에서 빼고 계산했습니다")
    if need_one and not numeric:
        raise InputError("숫자 기준 열이 표에 값 하나뿐이라 보간할 수 없습니다 (기준 열마다 서로 다른 값이 2개 이상 필요합니다).")


def _key_curves(sub: pd.DataFrame, key: str, values: List[str]) -> Dict[str, Tuple[np.ndarray, np.ndarray]]:
    """For each value column v: the key read along v (rows merged on equal v, sorted), kept only when it is a smooth curve
    that turns a few times (a hill: 수율 along 온도). Noisy records (the key jumps up and down along v) are not kept."""
    out = {}
    k_all = coerce_numeric(sub[key]).to_numpy(dtype=float)
    for v in values:
        x = coerce_numeric(sub[v]).to_numpy(dtype=float)
        ok = np.isfinite(x) & np.isfinite(k_all)
        if ok.sum() < 3:
            continue
        s = pd.Series(k_all[ok]).groupby(x[ok]).mean()
        xs, ks = s.index.to_numpy(dtype=float), s.to_numpy(dtype=float)
        if len(xs) < 3:
            continue
        d = np.sign(np.diff(ks))
        d = d[d != 0]
        turns = int((d[1:] != d[:-1]).sum())
        if 0 < turns <= max(1, (len(xs) - 2) // 5):
            out[v] = (xs, ks)
    return out


def _two_values(curves: Dict[str, Tuple[np.ndarray, np.ndarray]], k0: float) -> Dict[str, List[float]]:
    """{value column: its values (two or more, far apart) where the key curve takes k0} (검토 B A1)."""
    out = {}
    for v, (xs, ks) in curves.items():
        gap = TWO_VALUE_GAP * (xs[-1] - xs[0])
        groups: List[List[float]] = []
        for c in _crossings(xs, ks, k0):
            if groups and c - groups[-1][-1] <= gap:
                groups[-1].append(c)
            else:
                groups.append([c])
        if len(groups) >= 2:
            out[v] = [float(np.mean(g)) for g in groups]
    return out


def _missing_keys(sub: pd.DataFrame, keys: List[str], candidates: List[str], values: List[str]) -> List[str]:
    """Columns not given that the table needs as keys (검토 B A2: --input "온도=175" on a 온도 x 압력 grid averaged every 압력
    and said VERIFIED): rows of one key combination differ in the value columns, and adding these level columns (a text
    column, or a number column with few values) makes every combination one value. [] when the keys already do, or when no
    such columns do (replicate measurements: averaged with a warning as before)."""
    frame = pd.DataFrame({f"__k{j}": coerce_numeric(sub[k]).to_numpy(dtype=float) for j, k in enumerate(keys)})
    ok = frame.notna().all(axis=1).to_numpy()
    n = int(ok.sum())
    if n < 2:
        return []

    def cell(c: str) -> pd.Series:
        if _is_numeric_col(sub[c]):
            return pd.Series(coerce_numeric(sub[c]).to_numpy(dtype=float))
        return pd.Series(sub[c].map(lambda s: None if pd.isna(s) or str(s).strip() == "" else str(s).strip()).to_numpy(),
                         dtype=object)

    def mixed(extra: List[str]) -> bool:
        vcols = [v for v in values if v not in extra]
        if not vcols:
            return False
        g = frame.copy()
        for j, c in enumerate(extra):
            g[f"__e{j}"] = cell(c)
        for j, v in enumerate(vcols):
            g[f"__v{j}"] = pd.Series(coerce_numeric(sub[v]).to_numpy(dtype=float)).round(9)
        g = g[ok & g[[f"__e{j}" for j in range(len(extra))]].notna().all(axis=1).to_numpy()]
        by = [c for c in g.columns if c.startswith(("__k", "__e"))]
        nun = g.groupby(by, dropna=True)[[f"__v{j}" for j in range(len(vcols))]].nunique()
        return bool((nun > 1).to_numpy().any())

    if not mixed([]):
        return []
    level = []
    for c in candidates:
        s = cell(c)[ok]
        if s.isna().any():
            continue
        distinct = s.nunique()
        if distinct >= 2 and (not _is_numeric_col(sub[c]) or distinct <= n / 2):
            level.append(c)
    if not level or mixed(level):
        return []
    for c in reversed(list(level)):         # the fewest columns that still make one value per combination
        rest = [x for x in level if x != c]
        if rest and not mixed(rest):
            level = rest
    return level


def _column_values_text(sub: pd.DataFrame, c: str, n: int = 10) -> str:
    if _is_numeric_col(sub[c]):
        vals = [_fmt(v) for v in np.unique(coerce_numeric(sub[c]).dropna().to_numpy(dtype=float))]
    else:
        vals = sorted({str(v).strip() for v in sub[c] if not pd.isna(v) and str(v).strip()})
    return f"'{c}' {', '.join(vals[:n])}" + (f" 외 {len(vals) - n}개" if len(vals) > n else "")


def _flat_runs(xs: np.ndarray, ys: np.ndarray, y0: float) -> List[Tuple[float, float]]:
    """(from, to) of runs of two or more neighbouring nodes that all equal y0 (검토 B D3: the whole run is the answer)."""
    out: List[Tuple[float, float]] = []
    start = None
    for i in range(len(xs)):
        on = bool(np.isfinite(ys[i]) and ys[i] == y0)
        if on and start is None:
            start = i
        if (not on or i == len(xs) - 1) and start is not None:
            end = i if on else i - 1
            if end > start:
                out.append((float(xs[start]), float(xs[end])))
            start = None
    return out


# ==============================================================================
# 4. Commands
# ==============================================================================

def _common(tab: Dict[str, Any], result: Dict[str, Any], warnings: List[str],
            input_pct: Optional[List[str]] = None) -> Dict[str, Any]:
    """input_pct: percent_warning of the --input values and the --points key columns (said like the data's)."""
    if tab["excluded"]:
        warnings.insert(0, f"'{tab['exclude_col']}' 열에 표시된 {tab['excluded']}행은 쓰지 않았습니다")
    warnings[:0] = tab.get("layout_warnings") or []
    # numbers written with '%' ('64.4%') in the columns used are read as 64.4 (coerce_numeric): said, never silent
    used = (list(result.get("keys") or []) + list(result.get("columns") or []) + [result.get("vary") or ""]
            + list(result.get("find") or {}) + list(result.get("fixed") or {}))
    pct = [w for w in percent_warning(tab["df"], used) + list(input_pct or []) if w not in warnings]
    if pct:
        warnings.extend(pct)
        if "warnings" not in result.get("agent_directive", ""):
            result["agent_directive"] = result.get("agent_directive", "") + " warnings도 함께 알리세요."
    result.update({"data_file": tab["path"], "data_saved_at": tab["saved_at"]})
    if warnings:
        result["warnings"] = warnings
    return result


def execute_forward(args) -> Dict[str, Any]:
    tab = load_table(args)
    df, columns = tab["df"], data_columns(tab["df"], tab["exclude_col"])
    warnings: List[str] = []
    points_df = None
    # the --targets columns: never a key, never suggested as one (a measured '생산량_실측' is no typo of '생산량'); a --targets typo
    # never suggests an --input key either
    pairs, kv = _parse_pairs(args.input, "--input") if args.input else ({}, None)
    targets = resolve_column_names([s.strip() for s in args.targets.split(",") if s.strip()], columns,
                                   "구할 열(--targets)", layout=df.attrs,
                                   exclude=known_columns(list(pairs), columns))[0] if args.targets else []
    if args.points:
        ppath = os.path.abspath(args.points)
        if not os.path.isfile(ppath):
            raise FileNotFoundError(f"조회 파일이 없습니다: {ppath}")
        points_df = load_data_file(ppath, decimal_comma=tab["decimal_comma"])
        keycols = [c for c in columns if c not in targets]
        usable, measured = _points_keys(points_df, keycols, targets)
        if not usable:
            # the data file read with the wrong header row (title row, two-level header): the same hint as --input
            names = [str(c) for c in points_df.columns if not re.fullmatch(r"col_\d+", str(c))]
            hint = header_layout_hint(df.attrs, names) if names else ""
            raise InputError(f"조회 파일 열 {list(points_df.columns)} 중 데이터에 있는 기준 열이 없습니다. "
                             + (f"데이터 {hint}\n" if hint else "") + f"[사용 가능한 컬럼 목록]: {name_list(columns)}")
        # '12%' next to '0.12' in one key column of the points file: no scale is sure (as in the data)
        check_percent_mix(points_df, usable, "조회 파일 ")
        input_pct = percent_warning(points_df, usable, "조회 파일 ")
        # a number of each key column decides numeric / text (a blank or text first row is the row's own problem: D4)
        sample = {}
        for c in usable:
            nums = coerce_numeric(points_df[c])
            hit = np.flatnonzero(np.isfinite(nums.to_numpy(dtype=float)))
            sample[str(c)] = points_df[c].iloc[int(hit[0]) if len(hit) else 0]
        try:
            numeric, categorical = _split_keys(df, sample, keycols)
        except InputError as e:
            raise _ask_without_y(e, not args.targets)
        if categorical:
            raise InputError(f"조회 파일에는 숫자 기준 열만 넣어 주십시오 (글자 열: {list(categorical)}).")
        key_src = dict(zip(numeric, usable))
        _check_near_points(points_df, usable + measured, [c for c in keycols if c not in numeric], df.attrs,
                           ask=not args.targets, used=list(numeric) + targets, targets=targets)
    else:
        input_pct = []
        numeric, categorical = _split_keys(df, pairs, columns, args, kv, reserved=targets, pct=input_pct)
        key_src = {}
    if not numeric:
        raise InputError("숫자 기준 열이 하나 이상 필요합니다 (예: --input \"반응온도=232.5\").")
    given_keys = list(numeric)
    sub = _filter(df, categorical)
    if points_df is not None:
        # a key with one value in the table: every query must hold that value (then the key is left out, as for --input)
        for k in given_keys:
            vals = np.unique(coerce_numeric(sub[k]).dropna().to_numpy(dtype=float))
            pv = coerce_numeric(points_df[key_src[k]]).to_numpy(dtype=float)
            pv = pv[np.isfinite(pv)]
            if len(vals) == 1 and len(pv) and np.allclose(pv, vals[0], rtol=1e-9, atol=1e-12):
                numeric[k] = float(vals[0])
            elif len(vals) == 1:
                numeric[k] = float(pv[~np.isclose(pv, vals[0], rtol=1e-9, atol=1e-12)][0]) if len(pv) else float(vals[0])
    _drop_single_keys(sub, numeric, warnings, "--input" if points_df is None else "조회 파일")
    keys = list(numeric)
    values = _value_columns(df, args.targets, columns, skip=given_keys + list(categorical))
    # '11.9%' next to '0.119' in a data column used: no scale is sure. Without --targets every numeric column is a value
    # column: one nobody asked for is left out with a warning instead of stopping the lookup
    check_percent_mix(df, keys + (values if args.targets else []))
    if not args.targets:
        mixed = [c for c, _, _ in percent_mix(df, values)]
        if mixed and len(mixed) < len(values):
            values = [c for c in values if c not in mixed]
            warnings.append(f"값 열 {mixed}은(는) %를 붙인 숫자 칸과 % 없는 숫자 칸이 섞여 있어(100배 차이가 날 수 있음) "
                            f"계산하지 않았습니다. 그 열이 필요하면 열을 한 가지 형식으로 고친 뒤 --targets로 지정하십시오")
        check_percent_mix(df, values)       # every value column mixed: nothing left to compute
    show = list(dict.fromkeys(keys + list(categorical) + values))

    if len(keys) > MAX_KEYS:
        x = np.array([numeric[k] for k in keys])
        K = np.column_stack([coerce_numeric(sub[c]).to_numpy(dtype=float) for c in keys])
        span = np.nanmax(K, axis=0) - np.nanmin(K, axis=0)
        near = nearest_rows(sub, keys, x, np.where(span > 0, span, 1.0), show)
        warnings.append(f"기준 열이 {len(keys)}개라 보간하지 않고 가까운 실측 {len(near)}행만 보여 줍니다 (보간은 {MAX_KEYS}개까지)")
        return _common(tab, {"status": "ok", "mode": "nearest_only", "keys": keys, "columns": values,
                             "points": [], "nearest_rows": near, "count": 0, "verified": True,
                             "preview": f"기준 {len(keys)}개: 가까운 실측 {len(near)}행",
                             "agent_directive": "VERIFIED: 기준 열이 많아 보간값 없이 가까운 실측 행만 계산했습니다. "
                                                "nearest_rows와 warnings를 사용자에게 알리세요."}, warnings, input_pct)

    # 검토 B A2: a key column not given (압력 of a 온도 x 압력 grid, 라인 of a table per line): its rows would be averaged
    cands = [c for c in columns if c not in given_keys + list(categorical) + targets]
    missing = _missing_keys(sub, keys, cands, values)
    if missing:
        seen = "; ".join(_column_values_text(sub, c) for c in missing)
        names = ", ".join(repr(c) for c in missing)
        head = (f"기준으로 주지 않은 열 {names}이(가) 있어 계산하지 않았습니다: 같은 {', '.join(keys)} 값에 그 열이 다른 행이 여러 개라 "
                f"평균하면 틀린 값이 됩니다 (표의 값: {seen}).")
        if points_df is not None:
            raise InputError(f"{head} 조회 파일에 {names} 열을 더한 뒤 같은 명령을 다시 실행하십시오.")
        given_text = dict(kv) if kv else {str(k): str(v) for k, v in pairs.items()}
        cmd = _map_command(args, input=kv_text({**given_text, **{c: "<값>" for c in missing}}))
        raise InputError(f"{head} 사용자에게 {names} 값을 물은 뒤 next_commands[0]의 <값>을 그 값으로 바꿔 실행하십시오.",
                         extra={"missing_keys": missing,
                                "next_commands": [{"purpose": f"빠진 기준 {names} 값을 넣어 다시 계산 (<값>을 사용자가 준 값으로 바꿈)",
                                                   "command": cmd, "needs_input": True, "placeholder": "<값>"}]},
                         error_type="needs_user_decision")
    # 검토 B A1: one key that goes up and down along a value column (수율 along 온도): one 수율 has two 온도
    curves = _key_curves(sub, keys[0], values) if len(keys) == 1 else {}
    if curves and points_df is None:
        k, k0 = keys[0], numeric[keys[0]]
        two = _two_values(curves, k0)
        if two:
            found = "; ".join(f"'{v}' {', '.join(_fmt(x) for x in xs_)}" for v, xs_ in two.items())
            cmds = [_map_command(args, input=kv_text(categorical) if categorical else None, targets=None,
                                 find=f"{k}={_fmt(k0)}", vary=v) for v in two]
            head = (f"'{k}'={_fmt(k0)}인 곳이 여러 군데라({found}) 사잇값 하나로 낼 수 없어 계산하지 않았습니다 "
                    f"(이 표는 {', '.join(repr(v) for v in two)}에 따라 '{k}'이(가) 오르내립니다). ")
            if len(two) == 1:
                raise InputError(f"{head}'{k}'이(가) {_fmt(k0)}이 되는 '{next(iter(two))}'를 모두 찾는 명령: [추천] {cmds[0]}")
            text, items = numbered(list(two), cmds, f"'{k}'={_fmt(k0)}이 되는 '{{}}' 찾기")
            raise InputError(f"{head}어느 열을 찾을지 사용자에게 물으십시오: {text}", extra={"confirm_commands": items})

    t = MultiTable(sub, keys, values, warnings)   # each value column on its own non-empty rows
    values = t.values                             # a value column with too few values is left out (warned)
    show = list(dict.fromkeys(keys + list(categorical) + values))
    data_range = {k: [round(float(t.lo[j]), 6), round(float(t.hi[j]), 6)] for j, k in enumerate(keys)}

    if points_df is not None:
        X = np.column_stack([coerce_numeric(points_df[key_src[k]]).to_numpy(dtype=float) for k in keys])
        bad = ~np.isfinite(X).all(axis=1)
        # 검토 B D4: a blank key and a key that is no number are told apart (both are not computed)
        blank = np.column_stack([points_df[key_src[k]].map(lambda s: pd.isna(s) or str(s).strip() == "").to_numpy(dtype=bool)
                                 for k in keys]).any(axis=1) & bad
        # 검토 B A1: a query where the key holds at two places far apart is not computed (one value would be wrong)
        two_at = {i: _two_values(curves, float(x[0])) for i, x in enumerate(X) if not bad[i] and curves}
        two_at = {i: tv for i, tv in two_at.items() if tv}
        # numbers stay numbers in the result sheet (a CSV is read as text)
        pcols = {str(c): (coerce_numeric(points_df[c]) if _is_numeric_col(points_df[c]) else points_df[c])
                 for c in points_df.columns}
        results = {i: t.at(x) for i, x in enumerate(X) if not bad[i] and i not in two_at}
        per_column = any(_basis_differs(r) for r in results.values())
        out_rows, near_rows = [], []
        for i, x in enumerate(X):
            base = {c: _cell(col.iloc[i]) for c, col in pcols.items()}
            if bad[i]:
                out_rows.append({**base, "방식": "계산 불가 (기준 값 빈칸)" if blank[i] else "계산 불가 (기준 값이 숫자가 아님)"})
                continue
            if i in two_at:
                found = "; ".join(f"{v} {', '.join(_fmt(z) for z in zs)}" for v, zs in two_at[i].items())
                out_rows.append({**base, "방식": f"계산 불가 ({keys[0]} 값 하나에 맞는 곳이 여러 군데: {found})"})
                continue
            r = results[i]
            out_rows.append({**base, **{f"보간_{c}": v for c, v in r["values"].items()}, "방식": r["method"],
                             "외삽": "예" if r["extrapolated"] else "", "범위 밖": "; ".join(f"{k}: {v}" for k, v in r["out_of_range"].items()),
                             **_basis_cells(r, per_column)})
            for nr in nearest_rows(t.df, keys, x, t.span, show):
                near_rows.append({"조회 번호": i + 1, **{("원본 행" if k == "row" else "거리" if k == "distance" else k): v
                                                    for k, v in nr.items()}})
        save = _resolve_save(args.save, tab["path"]) or _default_save(os.path.dirname(os.path.abspath(args.points)),
                                                                     os.path.splitext(os.path.basename(args.points))[0])
        real, ok = _write_book(save, {"결과": pd.DataFrame(out_rows), "가까운 실측": pd.DataFrame(near_rows)}, args.overwrite)
        n_ext = sum(1 for r in out_rows if r.get("외삽") == "예")
        n_blank, n_text = int(blank.sum()), int((bad & ~blank).sum())
        failed = int(bad.sum()) + len(two_at)
        if n_blank:
            warnings.append(f"기준 값이 빈칸인 조회 {n_blank}건은 계산하지 않았습니다")
        if n_text:
            warnings.append(f"기준 값이 숫자가 아닌 조회 {n_text}건은 계산하지 않았습니다 (결과 파일 '방식' 열)")
        if two_at:
            warnings.append(f"조회 {len(two_at)}건은 '{keys[0]}' 값 하나에 맞는 곳이 여러 군데라 계산하지 않았습니다. "
                            f"그 값이 되는 곳은 --find \"{keys[0]}=<값>\" --vary <구할 열>로 찾습니다 (결과 파일 '방식' 열)")
        if n_ext:
            warnings.append(f"조회 {n_ext}건은 데이터 범위 밖이라 외삽값입니다 (결과 파일 '외삽' 열)")
        res = {"status": "ok", "mode": t.mode, "keys": keys, "columns": values, "count": len(out_rows),
               "computed_count": len(out_rows) - failed, "failed_count": failed,
               "extrapolated_count": n_ext, "output_file": real, "data_range": data_range, "verified": ok,
               "preview": (f"조회 {len(out_rows)}건 중 {len(out_rows) - failed}건 보간 ({'격자' if t.mode == 'grid' else '흩어진 기록'})"
                           + (f", {failed}건 계산 불가" if failed else "") + f", 외삽 {n_ext}건 -> {real}")}
        res["agent_directive"] = (("VERIFIED: " if ok else "UNVERIFIED: 결과 파일을 다시 읽어 확인하지 못했습니다. "
                                   "다시 실행하지 말고 unverified를 그대로 알리세요. ")
                                  + "preview를 그대로 알리고 output_file 경로와 warnings를 사용자에게 알리세요.")
        if not ok:
            res["unverified"] = ["결과 파일 다시 읽기"]
        return _common(tab, res, warnings, input_pct)

    x = np.array([numeric[k] for k in keys])
    r = t.at(x)
    near = nearest_rows(t.df, keys, x, t.span, show)
    # numeric columns not given as keys: with --targets every other one; without it the value columns themselves (each one
    # moves with the others), never one value column alone (검토 B A2: 'asked' left this list always empty without --targets)
    if args.targets:
        unused = [c for c in columns if c not in given_keys + list(categorical) + values and _is_numeric_col(sub[c])]
    else:
        unused = list(values) if len(values) > 1 else []
    mixed_in = t.mode == "scattered" and bool(unused)
    if mixed_in:
        warnings.append(f"흩어진 기록이라 기준으로 주지 않은 열 {unused[:5]}의 영향이 섞인 값입니다. 참고용으로 쓰십시오")
    point = {"input": {**{k: numeric[k] for k in keys}, **categorical}, **r}
    if r["extrapolated"]:
        warnings.append("데이터 범위 밖이라 외삽값입니다. 참고용으로만 쓰십시오: "
                        + ("; ".join(f"{k} {v}" for k, v in r["out_of_range"].items()) or "실측이 둘러싼 범위 밖"))
    vals_txt = ", ".join(f"{c} {_fmt(v)}" for c, v in r["values"].items())
    res = {"status": "ok", "mode": t.mode, "keys": keys, "columns": values, "points": [point],
           "nearest_rows": near, "data_range": data_range, "count": 1, "verified": True,
           "preview": f"{r['method']}: {vals_txt}" + (" [외삽]" if r["extrapolated"] else "")
           + (" (참고값: 기준으로 주지 않은 열의 영향이 섞임)" if mixed_in else "")}
    if args.save:
        real, ok = _write_book(_resolve_save(args.save, tab["path"]), {
            "결과": pd.DataFrame([{**point["input"], **{f"보간_{c}": v for c, v in r["values"].items()},
                                  "방식": r["method"], "외삽": "예" if r["extrapolated"] else "",
                                  **_basis_cells(r, _basis_differs(r))}]),
            "가까운 실측": pd.DataFrame([{("원본 행" if k == "row" else "거리" if k == "distance" else k): v
                                       for k, v in nr.items()} for nr in near])}, args.overwrite)
        res["output_file"], res["verified"] = real, ok
    res["agent_directive"] = (
        ("VERIFIED: " if res["verified"] else "UNVERIFIED: 결과 파일을 다시 읽어 확인하지 못했습니다. 다시 실행하지 말고 그대로 알리세요. ")
        + "preview의 값과 method, basis(사용한 원본 행과 비율)를 사용자에게 알리고, nearest_rows(가까운 실측 5행)를 표로 보여 주세요. "
        + ("extrapolated=true이므로 warnings의 범위 밖 내용을 반드시 함께 알리세요. " if r["extrapolated"] else "")
        + f"값은 데이터 파일의 마지막 저장 시각({tab['saved_at']}) 기준입니다.")
    return _common(tab, res, warnings, input_pct)


def execute_reverse(args) -> Dict[str, Any]:
    tab = load_table(args)
    df, columns = tab["df"], data_columns(tab["df"], tab["exclude_col"])
    warnings: List[str] = []
    goal, goal_kv = _parse_pairs(args.find, "--find")
    if len(goal) != 1:
        raise InputError("--find에는 목표 하나만 주십시오 (예: --find \"생산효율=50\").")
    check_comma_numbers(goal_kv, lambda k: True, "--find", lambda fixed: _map_command(args, find=kv_text(fixed)))
    (goal_name, goal_val), = goal.items()
    pairs, kv = _parse_pairs(args.input, "--input")
    # a typo never suggests a column another option names: the goal not --input/--vary not the goal or --input
    goal_col = resolve_column_names([goal_name], columns, "목표(--find)", layout=df.attrs, values=True,
                                    exclude=known_columns([args.vary.strip()] + list(pairs), columns))[0][0]
    input_pct: List[str] = []
    y0 = _number(goal_val, f"목표 값 '{goal_val}'이(가) 숫자가 아닙니다.", "--find", str(goal_name), input_pct)
    vary = resolve_column_names([args.vary.strip()], columns, "바꿀 기준(--vary)", layout=df.attrs,
                                exclude=[goal_col] + known_columns(list(pairs), columns))[0][0]
    numeric, categorical = _split_keys(df, pairs, columns, args, kv, reserved=[goal_col, vary], pct=input_pct)
    if vary in numeric or vary in categorical:
        raise InputError(f"'{vary}'는 --vary로 찾는 열이라 --input에 함께 줄 수 없습니다.")
    if goal_col in numeric or goal_col in categorical or goal_col == vary:
        raise InputError(f"목표 열 '{goal_col}'은(는) 기준 열(--vary, --input)과 달라야 합니다.")
    if not _is_numeric_col(df[vary]) or not _is_numeric_col(df[goal_col]):
        raise InputError("--vary 열과 목표 열은 숫자 열이어야 합니다.")
    keys = [vary] + list(numeric)
    check_percent_mix(df, keys + [goal_col])
    if len(keys) > MAX_KEYS:
        raise InputError(f"기준 열은 --vary 포함 {MAX_KEYS}개까지입니다 (지금 {len(keys)}개)."
                         + _model_find_hint(args, goal_col, goal_val, vary, {**numeric, **categorical}))
    sub = _filter(df, categorical)
    vary_vals = np.unique(coerce_numeric(sub[vary]).dropna().to_numpy(dtype=float))
    if len(vary_vals) < 2:
        raise InputError(f"--vary 열 '{vary}'은(는) 표에 값이 {len(vary_vals)}개뿐이라 바꿔 가며 찾을 수 없습니다"
                         + (f" ({_fmt(vary_vals[0])} 하나)." if len(vary_vals) else "."))
    if numeric:
        fixed_num = dict(numeric)
        _drop_single_keys(sub, fixed_num, warnings, need_one=False)       # --vary stays the key
        keys = [vary] + list(fixed_num)
    t = Table(sub, keys, [goal_col], warnings)
    fixed = np.array([numeric[k] for k in keys[1:]])
    others = [c for c in columns if c not in keys + [goal_col] + list(categorical) and _is_numeric_col(sub[c])]
    if others and t.mode == "scattered":
        warnings.append(f"고정하지 않은 열 {others[:5]}의 영향이 섞인 값입니다 (흩어진 기록)")

    def f(v: float) -> float:
        return float(t.at(np.concatenate([[v], fixed]))["values"][goal_col])

    lo, hi = float(t.lo[0]), float(t.hi[0])
    if t.mode == "grid" or len(keys) == 1:
        nodes = t.axes[0] if t.mode == "grid" else t.PK[:, 0]
        ys = np.array([f(v) for v in nodes])
        sols = [(x, False) for x in _crossings(nodes, ys, y0)]
        runs = _flat_runs(nodes, ys, y0)
        reach = (float(np.nanmin(ys)), float(np.nanmax(ys)))
        if t.mode == "grid":
            # end segments extended linearly (same rule as forward extrapolation)
            for (xa, ya, xb, yb, side) in ((nodes[0], ys[0], nodes[1], ys[1], "low"), (nodes[-2], ys[-2], nodes[-1], ys[-1], "high")):
                if yb != ya:
                    xs = xa + (y0 - ya) * (xb - xa) / (yb - ya)
                    if (side == "low" and xs < lo) or (side == "high" and xs > hi):
                        sols.append((float(xs), True))
        else:
            # scattered, one key: trend line outside the measured range
            a_lo, a_hi = t._trend_value(np.array([lo])), t._trend_value(np.array([hi]))
            if a_hi[0] != a_lo[0]:
                xs = lo + (y0 - a_lo[0]) * (hi - lo) / (a_hi[0] - a_lo[0])
                if xs < lo or xs > hi:
                    sols.append((float(xs), True))
    else:
        grid = np.linspace(lo, hi, SCAN_STEPS + 1)
        ys = np.full(len(grid), np.nan)
        for n, v in enumerate(grid):
            val, _ = t._scattered(np.concatenate([[v], fixed]))   # no trend extrapolation inside the scan
            if val is not None:
                ys[n] = float(val[0])
        sols = [(x, False) for x in _crossings(grid, ys, y0)]
        finite = ys[np.isfinite(ys)]
        reach = (float(finite.min()), float(finite.max())) if finite.size else (float("nan"), float("nan"))
        if not finite.size:
            warnings.append("고정한 조건이 실측이 둘러싼 범위 밖이라 찾을 수 없습니다")
        runs = []
    # 검토 B A7: an extrapolated answer far outside the data (k1: 온도 -137.6 for a range 10~100) is no answer
    far = [x for x, ext in sols if ext and (x < lo - FAR_FIND * (hi - lo) or x > hi + FAR_FIND * (hi - lo))]
    if far:
        warnings.append(f"데이터 범위({vary} {_fmt(lo)}~{_fmt(hi)})에서 데이터 폭의 {FAR_FIND:.0%}보다 먼 외삽 답 "
                        f"{', '.join(_fmt(x) for x in far)}은(는) 믿기 어려워 답으로 내지 않았습니다")
    # 검토 B D3: a flat stretch where every node equals the goal is one answer (from~to), not its nodes
    spans: Dict[float, Tuple[float, float]] = {}
    for a, b in runs:
        spans[(a + b) / 2] = (a, b)
    sols = sorted([(x, ext) for x, ext in sols if x not in far and not any(a <= x <= b for a, b in runs)]
                  + [(m, False) for m in spans])

    solutions, unverified = [], []
    tol = 1e-6 * max(1.0, abs(y0)) if (t.mode == "grid" or len(keys) == 1) else 1e-3 * max(1.0, float(np.ptp(t.PV[:, 0])))
    for x, ext in sols:
        r = t.at(np.concatenate([[x], fixed]))
        back = r["values"][goal_col]
        if abs(back - y0) > tol:
            unverified.append(f"{vary}={_fmt(x)}에서 {goal_col}={_fmt(back)} (목표 {_fmt(y0)})")
        span = spans.get(x) if not ext else None
        solutions.append({vary: round(x, 6), **({"from": round(span[0], 6), "to": round(span[1], 6)} if span else {}),
                          "extrapolated": ext or r["extrapolated"], "method": r["method"],
                          "basis": r["basis"], "check": round(back, 6),
                          **({"out_of_range": r["out_of_range"]} if r["out_of_range"] else {})})
    model_cmd = ""
    if t.mode == "scattered" and len(solutions) > 3:
        # 검토 B D5: many answers are measurement noise; the model trained from this file answers the same question
        model_cmd, dropped = _model_find_command(args, goal_col, y0, vary, {**numeric, **categorical})
        if dropped:
            model_cmd = ""
        warnings.append(f"흩어진 기록이라 {goal_col} 값이 들쭉날쭉해 답이 {len(solutions)}개입니다. "
                        + ("이 파일로 학습한 모델로 찾는 명령이 next_commands에 있습니다" if model_cmd else
                           "다른 조건을 --input로 고정하거나, 예측 모델(ml_train.py 학습 후 ml_predict.py)로 보는 것이 맞습니다"))
    if any(s["extrapolated"] for s in solutions):
        warnings.append(f"데이터 범위({_fmt(lo)}~{_fmt(hi)}) 밖에서 찾은 답이 있습니다. 외삽값이라 참고용입니다")

    # nearest measured rows: goal column close to the target, fixed keys close to their values
    near_cols, near_target, near_span = [goal_col] + keys[1:], np.concatenate([[y0], fixed]), np.concatenate([[max(float(np.ptp(t.V[:, 0])), 1e-12)], t.span[1:]])
    show = list(dict.fromkeys(keys + [goal_col] + list(categorical)))
    near = nearest_rows(t.df, near_cols, near_target, near_span, show)

    res: Dict[str, Any] = {"status": "ok", "mode": t.mode, "find": {goal_col: y0}, "vary": vary,
                           "fixed": {**{k: numeric[k] for k in keys[1:]}, **categorical},
                           "count": len(solutions), "reachable": {"min": round(reach[0], 6), "max": round(reach[1], 6),
                                                                 "within": f"{vary} {_fmt(lo)}~{_fmt(hi)}"},
                           "nearest_rows": near, "data_range": {vary: [round(lo, 6), round(hi, 6)]}}
    if len(solutions) > MAX_SOLUTIONS_JSON or args.save:
        save = _resolve_save(args.save, tab["path"]) or _default_save(os.path.dirname(tab["path"]),
                                                                     os.path.splitext(os.path.basename(tab["path"]))[0])
        real, ok = _write_book(save, {"답": pd.DataFrame([{vary: s[vary], f"확인_{goal_col}": s["check"], "외삽": "예" if s["extrapolated"] else "",
                                                        "방식": s["method"], "근거": _basis_text(s["basis"])} for s in solutions]),
                                      "가까운 실측": pd.DataFrame([{("원본 행" if k == "row" else "거리" if k == "distance" else k): v
                                                               for k, v in nr.items()} for nr in near])}, args.overwrite)
        res["output_file"] = real
        if not ok:
            unverified.append("결과 파일 다시 읽기")
        res["truncated"] = len(solutions) > MAX_SOLUTIONS_JSON
        if res["truncated"]:
            solutions = solutions[:MAX_SOLUTIONS_JSON]
            res["shown"] = len(solutions)
    res["solutions"] = solutions
    res["verified"] = not unverified
    if unverified:
        res["unverified"] = unverified
    if res["count"] == 0:
        res["preview"] = (f"{goal_col}={_fmt(y0)}이 되는 {vary}: 찾은 값 없음 "
                          f"({vary} {_fmt(lo)}~{_fmt(hi)}에서 {goal_col}은 {_fmt(reach[0])}~{_fmt(reach[1])})")
        directive = "찾은 값 없음(count 0). preview의 도달 가능한 범위(reachable)와 nearest_rows를 사용자에게 알리세요. "
    elif model_cmd:
        res["preview"] = (f"{goal_col}={_fmt(y0)}이 되는 {vary}: 답 {res['count']}개 (흩어진 기록의 측정 잡음). "
                          f"학습한 모델로 찾기 권장 (next_commands)")
        res["next_commands"] = [{"purpose": "이 파일로 학습한 모델로 같은 목표 찾기 (측정 잡음 없이, 읽기만 함)", "command": model_cmd}]
        directive = ("답이 측정 잡음으로 많습니다. preview를 알리고 next_commands[0].command를 글자 그대로 1번 실행해 "
                     "모델 기준 답을 알리세요. ")
    else:
        res["preview"] = f"{goal_col}={_fmt(y0)}이 되는 {vary}: " + ", ".join(
            (f"{_fmt(s['from'])}~{_fmt(s['to'])}" if "from" in s else _fmt(s[vary])) + (" [외삽]" if s["extrapolated"] else "")
            for s in solutions)
        directive = ("답을 모두 알리세요(하나만 고르지 않음). 답마다 method와 basis(원본 행과 비율)를 함께 알리고, "
                     "nearest_rows(가까운 실측 5행)를 표로 보여 주세요. "
                     + ("from~to가 있는 답은 그 구간 전체가 답입니다. " if any("from" in s for s in solutions) else ""))
    res["agent_directive"] = (("VERIFIED: 찾은 답을 보간식에 다시 넣어 목표값이 나오는지 확인했습니다. " if not unverified else
                               f"UNVERIFIED: 확인하지 못한 항목이 있습니다 ({'; '.join(unverified)}). 다시 실행하지 말고 unverified를 그대로 알리세요. ")
                              + directive + ("warnings도 함께 알리세요. " if warnings else "")
                              + f"값은 데이터 파일의 마지막 저장 시각({tab['saved_at']}) 기준입니다.")
    return _common(tab, res, warnings, input_pct)


# ==============================================================================
# 5. CLI
# ==============================================================================

def main():
    setup_console()
    parser = JsonArgumentParser(description="ml-kit 표 보간 / 반대 찾기 (선형 보간, 파일 직접 읽기)")
    parser.add_argument("--file", required=True, help="데이터 파일 (.xlsx, .csv)")
    parser.add_argument("--input", help="기준 열과 값 ('반응온도=232.5,소비전력=60'). 글자 열(라인=A)은 그 값인 행만 사용")
    parser.add_argument("--targets", help="구할 열 (쉼표 구분, 생략 시 기준 열을 뺀 숫자 열 전부)")
    parser.add_argument("--find", help="반대 찾기 목표 ('생산효율=50')")
    parser.add_argument("--vary", help="반대 찾기에서 값을 찾을 기준 열 (예: 반응온도). 다른 기준은 --input로 고정")
    parser.add_argument("--points", help="여러 점 조회 파일 (.xlsx, .csv, 열 이름 = 기준 열). 결과는 엑셀 파일로 저장")
    parser.add_argument("--save", help="결과 엑셀 저장 경로 (.xlsx). 이름만 주면 데이터 파일 옆")
    parser.add_argument("--overwrite", action="store_true",
                        help="--save 파일이 이미 있으면 덮어쓰기 (생략 시 시각을 붙인 새 이름으로 저장)")
    parser.add_argument("--header-row", type=int, help="데이터 머리글 행 번호 (생략 시 자동)")
    parser.add_argument("--unit-row", type=int, help="데이터 단위 행 번호 (선택)")
    parser.add_argument("--data-row", type=int, help="데이터 시작 행 번호 (선택)")
    args = parser.parse_args()

    if args.find or args.vary:
        if not (args.find and args.vary):
            parser.error("반대 찾기는 --find와 --vary를 함께 주어야 합니다.")
        if args.points or args.targets:
            parser.error("반대 찾기(--find)에는 --points, --targets를 쓰지 않습니다.")
        run_cli(lambda: execute_reverse(args))
        return
    if not args.input and not args.points:
        parser.error("--input(기준 값) 또는 --points(조회 파일) 중 하나를 주십시오.")
    if args.input and args.points:
        parser.error("--input와 --points는 함께 쓸 수 없습니다.")
    run_cli(lambda: execute_forward(args))


if __name__ == "__main__":
    main()
