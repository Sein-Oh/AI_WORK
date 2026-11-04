# -*- coding: utf-8 -*-
"""
ml-kit 표 보간 / 반대 찾기 (ml_map.py, 2.6.0)

Reads a data table (.xlsx / .csv) directly (no Excel) and answers two questions with linear interpolation:
  forward : values of the other columns at given key values      -i "반응온도=232.5,소비전력=60"
  reverse : key value where a column reaches a target value      --find "생산효율=50" --vary 반응온도 [-i fixed keys]
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

Shared runtime (data loading, column matching, CLI protocol) is imported from ml_predict.py.
"""
import re
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from ml_predict import (  # noqa: E402  shared runtime lives in ml_predict.py
    check_core_libraries, exit_missing_library,
    setup_console, run_cli, JsonArgumentParser, InputError, ColumnNameError, resolve_column_names, coerce_numeric,
    percent_warning, check_percent_mix, percent_mix, _strip_percent, name_list,
    load_data_file, layout_warnings, parse_inline_input, strip_unit, save_with_lock_fallback, avoid_overwrite, warn_long_path,
    check_comma_numbers, check_same_column, suggest_columns, build_alias_map, known_columns, _alias_key, _same_unit_name, kv_text, quote_arg, header_layout_hint, EXCLUDE_COL, INCLUDE_MARK, MAX_KEYS, Table, MultiTable, nearest_rows, _is_numeric_col, _cell, _fmt, _crossings,
)

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
    notes = layout_warnings(df)     # repeated header rows (dropped by the loader), totals rows
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
    """The same ml_map.py command with a corrected -i / --find text (absolute, quoted paths)."""
    parts = [f"python {quote_arg(os.path.abspath(__file__))}", "-f", quote_arg(os.path.abspath(args.file))]
    for flag, key in (("--header-row", "header_row"), ("--unit-row", "unit_row"), ("--data-row", "data_row")):
        if getattr(args, key, None) is not None:
            parts += [flag, str(getattr(args, key))]
    for flag, key in (("-i", "input"), ("-y", "targets"), ("--find", "find"), ("--vary", "vary")):
        val = change.get(key, getattr(args, key, None))
        if val:
            parts += [flag, quote_arg(val)]
    if getattr(args, "save", None):
        parts += ["--save", quote_arg(args.save)]
    if getattr(args, "overwrite", False):
        parts.append("--overwrite")
    return " ".join(parts)


def _split_keys(df: pd.DataFrame, pairs: Dict[str, Any], columns: List[str], args=None,
                kv: Optional[Dict[str, str]] = None, reserved: Optional[List[str]] = None,
                pct: Optional[List[str]] = None) -> Tuple[Dict[str, float], Dict[str, str]]:
    """Resolves -i names; numeric columns need numeric values, text columns become row filters. reserved (-y, the
    --find goal, --vary) are never suggested as a key. A value written with '%' ('10%') is read without it as
    coerce_numeric reads data cells (10, not 0.1); pct collects the percent_warning of such -i values."""
    names, _ = resolve_column_names(list(pairs), columns, "기준(-i)" if args is not None else "조회 파일(--points) 기준",
                                    layout=df.attrs, values=True, exclude=reserved)
    check_same_column(list(pairs), names, "-i" if args is not None else "--points 조회 파일 열",
                      pairs if args is not None else None)
    resolved = dict(zip(pairs, names))
    if kv and args is not None:
        check_comma_numbers(kv, lambda k: k in resolved and _is_numeric_col(df[resolved[k]]), "-i",
                            lambda fixed: _map_command(args, input=kv_text(fixed)))
    numeric, categorical = {}, {}
    for given, col in zip(pairs, names):
        v = pairs[given]
        if _is_numeric_col(df[col]):
            text = str(v).strip()
            try:
                numeric[col] = float(_strip_percent(text))
            except (TypeError, ValueError):
                raise InputError(f"기준 열 '{col}'은(는) 숫자 열인데 값 '{v}'이(가) 숫자가 아닙니다.")
            if pct is not None and _strip_percent(text) != text:
                pct.append(f"-i '{given}'의 값 '{text}'은(는) %를 떼고 숫자로 읽었습니다 ({text} -> {_strip_percent(text)}, "
                           f"100으로 나누지 않음). 열의 단위가 %인지 확인하십시오.")
        else:
            categorical[col] = str(v).strip()
    return numeric, categorical


_ASK_POINTS = (" -y가 없어 이 열이 기준 열인지 구할 열의 실측값·메모인지 알 수 없어 사용자에게 물으십시오: 기준 열이면 조회 파일의 "
               "머리글을 고르신 데이터 열 이름으로 고치고, 실측값·메모면 -y에 구할 열을 넣어 다시 실행하십시오 (그 열의 실측값 "
               "열은 결과에 그대로 붙습니다).")


def _ask_without_y(e: InputError, ask: bool) -> InputError:
    """Without -y every numeric column not used as a key is a value column, so a points column near one of them may be
    its measured value: the user says which (the -y guidance). Only for an unknown column name (ColumnNameError of
    resolve_column_names), never for another error (a key value that is no number, two names of one column)."""
    if not ask or not isinstance(e, ColumnNameError):
        return e
    head, sep, tail = str(e).partition("\n")
    return ColumnNameError(f"{head}{_ASK_POINTS}{sep}{tail}")


def _is_exact_key(name: str, cols: List[str]) -> bool:
    """name is a column of cols exactly: the same name, or the same name and unit spaced otherwise."""
    return name in cols or sum(_same_unit_name(name, c) for c in cols) == 1


def _is_key_name(name: str, cols: List[str]) -> bool:
    """name is a column of cols as -i matches it: an exact name (_is_exact_key) or a unique name written without
    unit."""
    if _is_exact_key(name, cols):
        return True
    return strip_unit(name) == name.strip() and _alias_key(name) in build_alias_map(cols)


def _points_keys(points_df: pd.DataFrame, keycols: List[str], targets: List[str]) -> Tuple[List[Any], List[Any]]:
    """(key columns, measured columns) of a --points file. A column of the -y target's name ('생산량(실측)' for '생산량')
    is a measured or memo column, never a key, unless it names a non-target column exactly: '수율' next to the
    columns '수율(1차)' and '수율(최종)' with -y 수율(최종) is no unique name of '수율(1차)' (keycols leave the target
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
    message as -i ([후보] / unit_hint). Columns like no data column (memo, ID) pass through.
    Near means the same name without unit or a close spelling of the whole name; the comparison of the names without
    unit only builds the [후보] list (as the gate it stopped measured columns: '수율_실측' / '수율(1차)').
    The check comes first for every column: a column holding a -y target's name ('수율_실측', '목표수율' for -y
    '수율(최종)') passes through as a measured value or memo only when no unused column is near (a key column
    '냉각수 온도(°C)' with -y '온도' is near '냉각수온도' and stops). A column and a near column that both hold a -y
    target's name without unit (targets) are compared without it: '수율_실측' / '함수율' share only '수율' and pass
    ('_실측' / '함'), '함수율%' / '함수율' stay near ('함%' / '함'). unused never holds a -y target; ask (no -y): the -y
    guidance. used (the keys and -y columns) are never listed."""
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


def _filter(df: pd.DataFrame, categorical: Dict[str, str]) -> pd.DataFrame:
    for col, v in categorical.items():
        vals = df[col].map(lambda s: "" if pd.isna(s) else str(s).strip())
        if not (vals == v).any():
            known = sorted(set(vals) - {""})[:20]
            raise InputError(f"'{col}' 열에 '{v}' 값이 없습니다. [후보] {', '.join(known)}")
        df = df[vals == v]
    return df


def _value_columns(df: pd.DataFrame, args_y: Optional[str], columns: List[str], skip: List[str]) -> List[str]:
    """Value columns, numeric judged on the whole table (df before the -i text filter): a column that is empty
    in the filtered rows is still a value column and MultiTable leaves it out with a warning."""
    if args_y:
        names, _ = resolve_column_names([s.strip() for s in args_y.split(",") if s.strip()], columns, "구할 열(-y)",
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


# ==============================================================================
# 4. Commands
# ==============================================================================

def _common(tab: Dict[str, Any], result: Dict[str, Any], warnings: List[str],
            input_pct: Optional[List[str]] = None) -> Dict[str, Any]:
    """input_pct: percent_warning of the -i values and the --points key columns (said like the data's)."""
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
    # the -y columns: never a key, never suggested as one (a measured '생산량_실측' is no typo of '생산량'); a -y typo
    # never suggests an -i key either
    pairs, kv = _parse_pairs(args.input, "-i") if args.input else ({}, None)
    targets = resolve_column_names([s.strip() for s in args.targets.split(",") if s.strip()], columns,
                                   "구할 열(-y)", layout=df.attrs,
                                   exclude=known_columns(list(pairs), columns))[0] if args.targets else []
    if args.points:
        ppath = os.path.abspath(args.points)
        if not os.path.isfile(ppath):
            raise FileNotFoundError(f"조회 파일이 없습니다: {ppath}")
        points_df = load_data_file(ppath, decimal_comma=tab["decimal_comma"])
        keycols = [c for c in columns if c not in targets]
        usable, measured = _points_keys(points_df, keycols, targets)
        if not usable:
            # the data file read with the wrong header row (title row, two-level header): the same hint as -i
            names = [str(c) for c in points_df.columns if not re.fullmatch(r"col_\d+", str(c))]
            hint = header_layout_hint(df.attrs, names) if names else ""
            raise InputError(f"조회 파일 열 {list(points_df.columns)} 중 데이터에 있는 기준 열이 없습니다. "
                             + (f"데이터 {hint}\n" if hint else "") + f"[사용 가능한 컬럼 목록]: {name_list(columns)}")
        # '12%' next to '0.12' in one key column of the points file: no scale is sure (as in the data)
        check_percent_mix(points_df, usable, "조회 파일 ")
        input_pct = percent_warning(points_df, usable, "조회 파일 ")
        sample = {str(c): points_df[c].iloc[0] for c in usable}
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
        raise InputError("숫자 기준 열이 하나 이상 필요합니다 (예: -i \"반응온도=232.5\").")
    keys = list(numeric)
    sub = _filter(df, categorical)
    values = _value_columns(df, args.targets, columns, skip=keys + list(categorical))
    asked = list(values)
    # '11.9%' next to '0.119' in a data column used: no scale is sure. Without -y every numeric column is a value
    # column: one nobody asked for is left out with a warning instead of stopping the lookup (LOW-3, 2.9.42)
    check_percent_mix(df, keys + (values if args.targets else []))
    if not args.targets:
        mixed = [c for c, _, _ in percent_mix(df, values)]
        if mixed and len(mixed) < len(values):
            values = [c for c in values if c not in mixed]
            warnings.append(f"값 열 {mixed}은(는) %를 붙인 숫자 칸과 % 없는 숫자 칸이 섞여 있어(100배 차이가 날 수 있음) "
                            f"계산하지 않았습니다. 그 열이 필요하면 열을 한 가지 형식으로 고친 뒤 -y로 지정하십시오")
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

    t = MultiTable(sub, keys, values, warnings)   # each value column on its own non-empty rows
    values = t.values                             # a value column with too few values is left out (warned)
    show = list(dict.fromkeys(keys + list(categorical) + values))
    data_range = {k: [round(float(t.lo[j]), 6), round(float(t.hi[j]), 6)] for j, k in enumerate(keys)}

    if points_df is not None:
        X = np.column_stack([coerce_numeric(points_df[key_src[k]]).to_numpy(dtype=float) for k in keys])
        bad = ~np.isfinite(X).all(axis=1)
        # numbers stay numbers in the result sheet (a CSV is read as text)
        pcols = {str(c): (coerce_numeric(points_df[c]) if _is_numeric_col(points_df[c]) else points_df[c])
                 for c in points_df.columns}
        results = {i: t.at(x) for i, x in enumerate(X) if not bad[i]}
        per_column = any(_basis_differs(r) for r in results.values())
        out_rows, near_rows = [], []
        for i, x in enumerate(X):
            base = {c: _cell(col.iloc[i]) for c, col in pcols.items()}
            if bad[i]:
                out_rows.append({**base, "방식": "계산 불가 (기준 값 빈칸)"})
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
        if bad.any():
            warnings.append(f"기준 값이 빈칸인 조회 {int(bad.sum())}건은 계산하지 않았습니다")
        if n_ext:
            warnings.append(f"조회 {n_ext}건은 데이터 범위 밖이라 외삽값입니다 (결과 파일 '외삽' 열)")
        res = {"status": "ok", "mode": t.mode, "keys": keys, "columns": values, "count": len(out_rows),
               "extrapolated_count": n_ext, "output_file": real, "data_range": data_range, "verified": ok,
               "preview": f"조회 {len(out_rows)}건 보간 ({'격자' if t.mode == 'grid' else '흩어진 기록'}), 외삽 {n_ext}건 -> {real}"}
        res["agent_directive"] = (("VERIFIED: " if ok else "UNVERIFIED: 결과 파일을 다시 읽어 확인하지 못했습니다. ")
                                  + "output_file 경로와 외삽 건수(extrapolated_count), warnings를 사용자에게 알리세요.")
        if not ok:
            res["unverified"] = ["결과 파일 다시 읽기"]
        return _common(tab, res, warnings, input_pct)

    x = np.array([numeric[k] for k in keys])
    r = t.at(x)
    near = nearest_rows(t.df, keys, x, t.span, show)
    unused = [c for c in columns if c not in keys + asked + list(categorical) and _is_numeric_col(sub[c])]
    if t.mode == "scattered" and unused:
        warnings.append(f"흩어진 기록이라 기준으로 주지 않은 열 {unused[:5]}의 영향이 섞인 값입니다. 참고용으로 쓰십시오")
    point = {"input": {**{k: numeric[k] for k in keys}, **categorical}, **r}
    if r["extrapolated"]:
        warnings.append("데이터 범위 밖이라 외삽값입니다. 참고용으로만 쓰십시오: "
                        + ("; ".join(f"{k} {v}" for k, v in r["out_of_range"].items()) or "실측이 둘러싼 범위 밖"))
    vals_txt = ", ".join(f"{c} {_fmt(v)}" for c, v in r["values"].items())
    res = {"status": "ok", "mode": t.mode, "keys": keys, "columns": values, "points": [point],
           "nearest_rows": near, "data_range": data_range, "count": 1, "verified": True,
           "preview": f"{r['method']}: {vals_txt}" + (" [외삽]" if r["extrapolated"] else "")}
    if args.save:
        real, ok = _write_book(_resolve_save(args.save, tab["path"]), {
            "결과": pd.DataFrame([{**point["input"], **{f"보간_{c}": v for c, v in r["values"].items()},
                                  "방식": r["method"], "외삽": "예" if r["extrapolated"] else "",
                                  **_basis_cells(r, _basis_differs(r))}]),
            "가까운 실측": pd.DataFrame([{("원본 행" if k == "row" else "거리" if k == "distance" else k): v
                                       for k, v in nr.items()} for nr in near])}, args.overwrite)
        res["output_file"], res["verified"] = real, ok
    res["agent_directive"] = (
        ("VERIFIED: " if res["verified"] else "UNVERIFIED: 결과 파일을 다시 읽어 확인하지 못했습니다. ")
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
    pairs, kv = _parse_pairs(args.input, "-i")
    # a typo never suggests a column another option names: the goal not -i/--vary, --vary not the goal or -i
    goal_col = resolve_column_names([goal_name], columns, "목표(--find)", layout=df.attrs, values=True,
                                    exclude=known_columns([args.vary.strip()] + list(pairs), columns))[0][0]
    try:
        y0 = float(goal_val)
    except (TypeError, ValueError):
        raise InputError(f"목표 값 '{goal_val}'이(가) 숫자가 아닙니다.")
    vary = resolve_column_names([args.vary.strip()], columns, "바꿀 기준(--vary)", layout=df.attrs,
                                exclude=[goal_col] + known_columns(list(pairs), columns))[0][0]
    input_pct: List[str] = []
    numeric, categorical = _split_keys(df, pairs, columns, args, kv, reserved=[goal_col, vary], pct=input_pct)
    if vary in numeric or vary in categorical:
        raise InputError(f"'{vary}'는 --vary로 찾는 열이라 -i에 함께 줄 수 없습니다.")
    if goal_col in numeric or goal_col in categorical or goal_col == vary:
        raise InputError(f"목표 열 '{goal_col}'은(는) 기준 열(--vary, -i)과 달라야 합니다.")
    if not _is_numeric_col(df[vary]) or not _is_numeric_col(df[goal_col]):
        raise InputError("--vary 열과 목표 열은 숫자 열이어야 합니다.")
    keys = [vary] + list(numeric)
    check_percent_mix(df, keys + [goal_col])
    if len(keys) > MAX_KEYS:
        raise InputError(f"기준 열은 --vary 포함 {MAX_KEYS}개까지입니다 (지금 {len(keys)}개).")
    sub = _filter(df, categorical)
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
    sols = sorted(sols)

    solutions, unverified = [], []
    tol = 1e-6 * max(1.0, abs(y0)) if (t.mode == "grid" or len(keys) == 1) else 1e-3 * max(1.0, float(np.ptp(t.PV[:, 0])))
    for x, ext in sols:
        r = t.at(np.concatenate([[x], fixed]))
        back = r["values"][goal_col]
        if abs(back - y0) > tol:
            unverified.append(f"{vary}={_fmt(x)}에서 {goal_col}={_fmt(back)} (목표 {_fmt(y0)})")
        solutions.append({vary: round(x, 6), "extrapolated": ext or r["extrapolated"], "method": r["method"],
                          "basis": r["basis"], "check": round(back, 6),
                          **({"out_of_range": r["out_of_range"]} if r["out_of_range"] else {})})
    if t.mode == "scattered" and len(solutions) > 3:
        warnings.append(f"흩어진 기록이라 {goal_col} 값이 들쭉날쭉해 답이 {len(solutions)}개입니다. "
                        f"다른 조건을 -i로 고정하거나, 예측 모델(ml_train.py 학습 후 ml_predict.py)로 보는 것이 맞습니다")
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
    else:
        res["preview"] = f"{goal_col}={_fmt(y0)}이 되는 {vary}: " + ", ".join(
            _fmt(s[vary]) + (" [외삽]" if s["extrapolated"] else "") for s in solutions)
        directive = ("답을 모두 알리세요(하나만 고르지 않음). 답마다 method와 basis(원본 행과 비율)를 함께 알리고, "
                     "nearest_rows(가까운 실측 5행)를 표로 보여 주세요. ")
    res["agent_directive"] = (("VERIFIED: 찾은 답을 보간식에 다시 넣어 목표값이 나오는지 확인했습니다. " if not unverified else
                               f"UNVERIFIED: 확인하지 못한 항목이 있습니다 ({'; '.join(unverified)}). 재시도하지 말고 그대로 알리세요. ")
                              + directive + ("warnings도 함께 알리세요. " if warnings else "")
                              + f"값은 데이터 파일의 마지막 저장 시각({tab['saved_at']}) 기준입니다.")
    return _common(tab, res, warnings, input_pct)


# ==============================================================================
# 5. CLI
# ==============================================================================

def main():
    setup_console()
    parser = JsonArgumentParser(description="ml-kit 표 보간 / 반대 찾기 (선형 보간, 파일 직접 읽기)")
    parser.add_argument("--file", "-f", required=True, help="데이터 파일 (.xlsx, .csv)")
    parser.add_argument("--input", "-i", help="기준 열과 값 ('반응온도=232.5,소비전력=60'). 글자 열(라인=A)은 그 값인 행만 사용")
    parser.add_argument("--targets", "-y", help="구할 열 (쉼표 구분, 생략 시 기준 열을 뺀 숫자 열 전부)")
    parser.add_argument("--find", help="반대 찾기 목표 ('생산효율=50')")
    parser.add_argument("--vary", help="반대 찾기에서 값을 찾을 기준 열 (예: 반응온도). 다른 기준은 -i로 고정")
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
            parser.error("반대 찾기(--find)에는 --points, -y를 쓰지 않습니다.")
        run_cli(lambda: execute_reverse(args))
        return
    if not args.input and not args.points:
        parser.error("-i(기준 값) 또는 --points(조회 파일) 중 하나를 주십시오.")
    if args.input and args.points:
        parser.error("-i와 --points는 함께 쓸 수 없습니다.")
    run_cli(lambda: execute_forward(args))


if __name__ == "__main__":
    main()
