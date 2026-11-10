# -*- coding: utf-8 -*-
"""
ml-kit 학습 데이터셋 만들기 (ml_dataset.py), 2.9.85

A training run makes versions in <data>_ml_data next to its data file, so a different set of rows is a different
data file. This tool writes that file (the data file is never changed):
  pick     the rows meeting --where (or the source rows --rows) -> a new file with only those rows.
  exclude  every row, with the '제외' column filled for the rows meeting --where (or --rows): training leaves them out
           and still scores them separately (the reason is the text in the cell). A row already marked in '제외'
           (a reason or '포함') keeps its mark.
Output: CSV (UTF-8 with BOM) by default, .xlsx when --save ends with .xlsx. Default name <data>_뽑기<N>.csv /
<data>_제외<N>.csv next to the data file (the first free number). The saved file is read back and compared.
next_commands: the training of the new file with the inputs and targets of the last training of the data file.
"""
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _ml_runtime import (  # noqa: E402  shared runtime of ml-kit
    check_core_libraries,
    setup_console, run_cli, JsonArgumentParser, InputError, load_data_file, quote_arg, avoid_overwrite,
    save_with_lock_fallback, parse_where, resolve_where, where_mask, where_text, where_notation_variants,
    where_with_variants, where_retry,
    exclude_mark, final_version, strip_unit, EXCLUDE_COL, VERSION, saved_as, command_text,
)
check_core_libraries("데이터셋")
setup_console()

import re  # noqa: E402
from typing import Any, Dict, List, Optional, Tuple  # noqa: E402

import numpy as np  # noqa: E402
import pandas as pd  # noqa: E402

INCLUDE_MARK = "포함"
KIND_NAME = {"pick": "뽑기", "exclude": "제외"}
PREVIEW_ROWS = 20          # source row numbers listed in the JSON
OUTPUT_EXTS = (".csv", ".xlsx")
FORMULA_START = ("=", "+", "-", "@")   # a cell text starting so is a formula in Excel (and when the xlsx is saved)
_MADE_NAME = re.compile(r"^(?P<parent>.+)_(?:뽑기|제외)\d+$")   # <원본>_뽑기N / <원본>_제외N made by this tool


def progress(msg: str) -> None:
    sys.stderr.write(f"[PROGRESS] {msg}\n")
    sys.stderr.flush()


# ==============================================================================
# 1. Rows
# ==============================================================================

def parse_rows(text: str, source_rows: List[int]) -> Tuple[np.ndarray, List[int]]:
    """--rows '15,27' (source row numbers, as in Excel and in excluded_rows.csv) -> (row mask, the numbers).
    A number that is not a data row (header, title, blank or outside the file) is an input error."""
    parts = [p.strip() for p in str(text).split(",") if p.strip()]
    if not parts:
        raise InputError("--rows에 행 번호가 없습니다. 형식: --rows \"15,27\" (원본 파일의 행 번호, 쉼표로 구분)")
    bad = [p for p in parts if not re.fullmatch(r"\d+", p)]
    if bad:
        raise InputError(f"--rows의 {bad}은(는) 행 번호가 아닙니다. 형식: --rows \"15,27\" (원본 파일의 행 번호, 쉼표로 구분, "
                         f"범위는 쓰지 않음)")
    nums = list(dict.fromkeys(int(p) for p in parts))
    pos = {r: i for i, r in enumerate(source_rows)}
    missing = [n for n in nums if n not in pos]
    if missing:
        lo, hi = (min(source_rows), max(source_rows)) if source_rows else (0, 0)
        raise InputError(f"--rows의 {missing}행은 데이터 행이 아닙니다 (머리글·제목·빈 행이거나 파일 밖). 데이터 행은 {lo}~{hi}행입니다. "
                         f"행 번호를 사용자에게 확인하십시오.")
    mask = np.zeros(len(source_rows), dtype=bool)
    mask[[pos[n] for n in nums]] = True
    return mask, nums


def select_rows(df: pd.DataFrame, args) -> Tuple[np.ndarray, str, List[str]]:
    """(rows chosen, what chose them in words, warnings) from --where or --rows."""
    warnings: List[str] = []
    if args.where:
        conds = resolve_where(parse_where(args.where, "--where"), [str(c) for c in df.columns], "--where", dict(df.attrs))
        text = where_text(conds)
        mask = where_mask(df, conds)
        variants = where_notation_variants(df, conds)
        if variants:
            # a code column compared as text: '월 = 1' leaves out the '01' rows (the same number in training)
            cmd = where_retry("--where", where_text(where_with_variants(conds, variants)))
            for (c, o, v), cnt in variants:
                notes = ", ".join(f"'{t}' {n}행" for t, n in cnt.items())
                warnings.append(f"--where '{c} {o} {v}': {notes}은(는) 표기가 달라 " + ("고르지 않음" if o == "=" else "함께 고름")
                                + f" ('{c}' 열은 앞자리 0을 구분하는 코드 열이라 글자 그대로 비교). 같은 값으로 보려면 [추천] {cmd}")
        return mask, f"조건 {text}", warnings
    source_rows = list(df.attrs.get("source_rows") or range(1, len(df) + 1))
    unit_below = df.attrs.get("unit_row_below")
    if unit_below and str(unit_below) in [p.strip() for p in str(args.rows).split(",")]:
        # 2.9.88 D3: read without --unit-row the unit row is kept as a data row (training leaves it out): not a row to pick
        raise InputError(f"--rows의 {unit_below}행은 머리글 아래의 단위 행입니다 (데이터 행이 아님). 데이터 행은 "
                         f"{unit_below + 1}~{max(source_rows)}행입니다. 행 번호를 사용자에게 확인하십시오.")
    mask, nums = parse_rows(args.rows, source_rows)
    return mask, f"행 {', '.join(map(str, nums))}", warnings


# ==============================================================================
# 2. Saving and checking
# ==============================================================================

def output_path(data_file: str, kind: str, save: Optional[str], overwrite: bool) -> str:
    """The file to write. --save: relative to the data file's folder (its default place); no extension -> .csv; a folder
    that does not exist is made (2.9.88 B3: was an internal error). An existing file gets a new name; its 'recovered'
    note waits for the save (saved_as, B6)."""
    folder = os.path.dirname(os.path.abspath(data_file))
    stem = os.path.splitext(os.path.basename(data_file))[0]
    if save:
        path = save if os.path.isabs(save) else os.path.join(folder, save)
        ext = os.path.splitext(path)[1].lower()
        if ext not in OUTPUT_EXTS:
            if ext:
                fixed = os.path.abspath(os.path.splitext(path)[0] + ".csv")
                cmd = command_text(drop={"--save": "save"})
                raise InputError(f"--save는 .csv 또는 .xlsx만 됩니다: '{save}'. [추천] "
                                 + (f"{cmd} --save {quote_arg(fixed)}" if cmd else f"--save {quote_arg(fixed)}"))
            path += ".csv"
        if os.path.abspath(path) == os.path.abspath(data_file):
            raise InputError("--save가 원본 데이터 파일과 같습니다. 원본은 바꾸지 않습니다. 다른 이름을 주십시오.")
        return avoid_overwrite(os.path.abspath(path), overwrite, what="데이터셋", defer=True)
    n = 1
    while os.path.exists(os.path.join(folder, f"{stem}_{KIND_NAME[kind]}{n}.csv")):
        n += 1
    return os.path.join(folder, f"{stem}_{KIND_NAME[kind]}{n}.csv")


def write_table(df: pd.DataFrame, path: str) -> str:
    try:
        os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    except OSError as e:
        raise InputError(f"데이터셋을 저장할 폴더를 만들 수 없습니다: {os.path.dirname(os.path.abspath(path))} ({e}). "
                         "--save 경로를 사용자에게 확인하십시오.")

    def _write(p: str) -> None:
        if p.lower().endswith(".xlsx"):
            df.to_excel(p, index=False)
        else:
            df.to_csv(p, index=False, encoding="utf-8-sig")
    return save_with_lock_fallback(path, _write, what="데이터셋")


def _cells(df: pd.DataFrame) -> List[List[str]]:
    """Cells as text for the comparison (a blank, NaN or None is ''; numbers by value)."""
    def one(v: Any) -> str:
        if v is None or (isinstance(v, float) and v != v):
            return ""
        if isinstance(v, (int, float, np.integer, np.floating)) and not isinstance(v, bool):
            return repr(float(v))
        s = str(v).strip()
        try:
            return repr(float(s)) if s and re.fullmatch(r"[-+]?(\d+\.?\d*|\.\d+)([eE][-+]?\d+)?", s) else s
        except ValueError:
            return s
    return [[one(v) for v in row] for row in df.itertuples(index=False, name=None)]


def verify(path: str, want: pd.DataFrame) -> List[str]:
    """Reads the saved file back with the training loader: same columns, rows and cells."""
    try:
        got = load_data_file(path)
    except Exception as e:
        return [f"저장한 파일을 다시 읽지 못함: {type(e).__name__}: {e}"]
    problems = []
    if [str(c) for c in got.columns] != [str(c) for c in want.columns]:
        problems.append(f"열이 다름 (저장 {list(got.columns)[:5]} / 기대 {list(want.columns)[:5]})")
    elif len(got) != len(want):
        problems.append(f"행 수가 다름 (저장 {len(got)} / 기대 {len(want)})")
    elif _cells(got) != _cells(want.reset_index(drop=True)):
        problems.append("다시 읽은 값이 기대값과 다름")
    return problems


# ==============================================================================
# 3. Next training command
# ==============================================================================

def trained_source(data_file: str) -> Tuple[str, Optional[Dict[str, Any]]]:
    """(the data file whose training is copied, its final_version). A file this tool made (<원본>_제외1.csv) that was not
    trained yet uses the training of its source, then of the source's source (2.9.88 C4: '그 행도 빼줘' on 공정_제외1.csv
    gave a <타겟> template)."""
    path = os.path.abspath(data_file)
    for _ in range(10):
        stem, ext = os.path.splitext(path)
        info = final_version(stem + "_ml_data")
        if info:
            return path, info
        m = _MADE_NAME.match(os.path.basename(stem))
        parent = next((os.path.join(os.path.dirname(path), m.group("parent") + e) for e in (ext, ".csv", ".xlsx", ".xls")
                       if m and os.path.isfile(os.path.join(os.path.dirname(path), m.group("parent") + e))), None)
        if not parent:
            break
        path = parent
    return os.path.abspath(data_file), None


def train_command(data_file: str, new_file: str) -> Dict[str, Any]:
    """The training of new_file with the inputs, targets, formulas, options and description (--desc) of the last
    training of data_file (or of the file it was made from, trained_source); without one a template with <타겟>."""
    here = os.path.dirname(os.path.abspath(__file__))
    head = f"python {quote_arg(os.path.join(here, 'ml_train.py'))} --file {quote_arg(new_file)}"
    src, info = trained_source(data_file)
    rec = (info or {}).get("record") or {}
    targets = [str(t) for t in rec.get("targets") or []]
    if not targets:
        return {"purpose": "새 데이터셋으로 학습 (<타겟>을 사용자가 말한 타겟으로 바꿀 것)",
                "command": f"{head} --targets \"<타겟>\"", "needs_input": True, "placeholder": "<타겟>"}
    evals = [str(e) for e in rec.get("derived_evals") or []]
    created = {e.split("=", 1)[0].strip() for e in evals}
    auto = bool((rec.get("input_selection") or {}).get("auto"))
    feats = [str(f) for f in rec.get("features_used") or [] if str(f) not in created]
    desc = str(rec.get("description") or "").strip()
    desc = desc if desc and "\n" not in desc and "\r" not in desc else ""
    texts = [",".join(targets), ",".join(feats), "; ".join(evals), str(rec.get("dtypes") or ""),
             str(rec.get("positive_class") or ""), desc]
    if any('"' in t for t in texts):
        return {"purpose": "새 데이터셋으로 학습 (이름에 큰따옴표가 있어 명령을 만들지 못함: 입력과 타겟을 사용자에게 확인)",
                "command": f"{head} --targets \"<타겟>\"", "needs_input": True, "placeholder": "<타겟>"}
    parts = [head]
    if rec.get("task") == "classification":
        parts += ["--task", "classification"]
        if rec.get("positive_class") is not None:
            parts += ["--positive-class", quote_arg(rec["positive_class"])]
    parts += ["--targets", quote_arg(",".join(targets))]
    if not auto and feats:
        parts += ["--features", quote_arg(",".join(feats))]
    if evals:
        parts += ["--eval", quote_arg("; ".join(evals))]
    if rec.get("dtypes"):
        parts += ["--dtypes", quote_arg(rec["dtypes"])]
    if desc:
        parts += ["--desc", quote_arg(desc)]
    return {"purpose": f"새 데이터셋으로 학습 ({os.path.basename(src)}의 {info['version']} 학습과 같은 타겟·입력)",
            "command": " ".join(parts), "needs_input": False}


# ==============================================================================
# 4. Command
# ==============================================================================

def execute(args) -> Dict[str, Any]:
    if bool(args.where) == bool(args.rows):
        raise InputError("--where(조건) 또는 --rows(행 번호) 중 하나만 주십시오.")
    data_file = os.path.abspath(args.file)
    progress(f"데이터 읽기: {os.path.basename(data_file)}")
    df = load_data_file(data_file, header_row=args.header_row, unit_row=args.unit_row, data_row=args.data_row)
    source_rows = list(df.attrs.get("source_rows") or range(1, len(df) + 1))
    mask, chosen_by, warnings = select_rows(df, args)
    hit = int(mask.sum())
    hit_rows = [source_rows[i] for i in np.flatnonzero(mask)]
    kind = args.action
    excl_col = next((str(c) for c in df.columns if strip_unit(str(c)) == EXCLUDE_COL), None)
    if args.reason and kind == "pick":
        warnings.append("--reason은 exclude에서만 씁니다. pick에서는 쓰지 않았습니다")
    base = {"status": "ok", "action": kind, "data_file": data_file, "rows_total": len(df), "count": hit,
            "chosen_by": chosen_by, "rows": hit_rows[:PREVIEW_ROWS], "rows_truncated": hit > PREVIEW_ROWS}
    if hit == 0:
        return {**base, "verified": True, "next_commands": [], **({"warnings": warnings} if warnings else {}),
                "preview": f"{chosen_by}에 맞는 행이 없습니다 (0행). 파일을 만들지 않았습니다.",
                "agent_directive": (f"VERIFIED: 찾은 행 없음 (count 0). {chosen_by}에 맞는 행이 없어 파일을 만들지 않았습니다. "
                                    "preview를 그대로 알리고 조건을 사용자에게 확인하세요.")}

    kept_marks = 0
    if kind == "pick":
        out = df[mask].reset_index(drop=True)
        what = f"{hit}행을 뽑은"
    else:
        out = df.copy()
        col = excl_col or EXCLUDE_COL
        if excl_col is None:
            out[col] = ""
        marks = out[col].map(exclude_mark)
        already = mask & (marks != "").to_numpy()
        kept_marks = int(already.sum())
        reason = (args.reason or "").strip() or chosen_by
        if reason == INCLUDE_MARK:
            raise InputError(f"--reason '{INCLUDE_MARK}'은(는) 학습에 넣으라는 표시라 뺄 이유로 쓸 수 없습니다. 이유를 다른 말로 주십시오.")
        if reason.startswith(FORMULA_START):
            reason_error(args, reason)
        new = mask & ~already
        out[col] = [reason if new[i] else out[col].iloc[i] for i in range(len(out))]
        if kept_marks:
            warnings.append(f"{kept_marks}행은 '{col}' 열에 이미 표시가 있어 그대로 두었습니다 (이유 또는 '{INCLUDE_MARK}')")
        what = f"{int(new.sum())}행에 제외 표시를 한"
    path = output_path(data_file, kind, args.save, args.overwrite)
    if args.dry_run:
        # the save runs with every option of the preview (2.9.88 B4: --header-row/--unit-row/--data-row/--overwrite were
        # left out, so the file was read differently from the preview)
        layout = [[o, str(v)] for o, v in (("--header-row", args.header_row), ("--unit-row", args.unit_row),
                                           ("--data-row", args.data_row)) if v is not None]
        cmd = " ".join([f"python {quote_arg(os.path.abspath(__file__))} {kind} --file {quote_arg(data_file)}"]
                       + (["--where", quote_arg(args.where)] if args.where else ["--rows", quote_arg(args.rows)])
                       + (["--reason", quote_arg(args.reason)] if args.reason and kind == "exclude" else [])
                       + sum(layout, []) + ["--save", quote_arg(path)] + (["--overwrite"] if args.overwrite else []))
        return {**base, "verified": True, "dry_run": True, "output_file": path, **({"warnings": warnings} if warnings else {}),
                "confirm_commands": [{"purpose": "미리보기대로 저장", "command": cmd}],
                "preview": f"미리보기: {chosen_by} -> {what} 데이터셋을 {os.path.basename(path)}로 저장합니다 (저장하지 않음).",
                "agent_directive": "VERIFIED: 미리보기입니다 (저장하지 않음). preview를 알리고, 사용자가 원하면 confirm_commands[0]을 실행하세요."}
    progress(f"저장: {os.path.basename(path)}")
    real = write_table(out, path)
    saved_as(path)
    problems = verify(real, out)
    nxt = train_command(data_file, real)
    res = {**base, "output_file": real, "rows_saved": len(out), "verified": not problems,
           "next_commands": [nxt],
           "preview": f"{chosen_by} -> {what} 데이터셋 저장: {os.path.basename(real)} ({len(out)}행, 원본은 그대로)"}
    if kind == "exclude":
        res["marked"] = hit - kept_marks
    if warnings:
        res["warnings"] = warnings
    if problems:
        res["unverified"] = problems
    res["agent_directive"] = (
        ("VERIFIED: 저장한 파일을 다시 읽어 확인했습니다. " if not problems else
         "UNVERIFIED: 저장한 파일을 확인하지 못했습니다 (" + "; ".join(problems) + "). 다시 실행하지 말고 unverified를 알리세요. ")
        + "preview" + (", warnings" if warnings else "") + "를 그대로 알리세요. 학습은 사용자가 원할 때 next_commands[0]"
        + ("의 <타겟>을 바꿔" if nxt.get("needs_input") else "을 그대로") + " 실행합니다.")
    return res


def reason_error(args, reason: str) -> None:
    """--reason starting with = + - @: Excel reads it as a formula (an xlsx saved it as one and it read back as 0,
    2.9.88 D6). STOP_RETRY with [추천] the same command with those leading signs left out (the same words)."""
    plain = reason.lstrip("".join(FORMULA_START) + " ")
    msg = (f"--reason '{reason}'은(는) {', '.join(FORMULA_START)}로 시작해 엑셀이 수식으로 읽습니다 (저장하면 이유 대신 계산 값이 됨). "
           "이 기호로 시작하지 않는 이유를 주십시오.")
    # --save goes back as its full path: command_text would make a relative --save absolute from the current folder,
    # while this tool reads it next to the data file
    cmd = command_text(drop={"--reason": "reason", "--save": "save"})
    save = ""
    if args.save:
        folder = os.path.dirname(os.path.abspath(args.file))
        save = f" --save {quote_arg(os.path.abspath(args.save if os.path.isabs(args.save) else os.path.join(folder, args.save)))}"
    if plain and cmd:
        msg += f" [추천] {cmd} --reason {quote_arg(plain)}{save}"
    raise InputError(msg)


def main():
    setup_console()
    parser = JsonArgumentParser(description=f"ml-kit 학습 데이터셋 만들기 (원본은 바꾸지 않음, ml-kit {VERSION})")
    parser.add_argument("action", choices=["pick", "exclude"],
                        help="pick: 맞는 행만 새 파일로, exclude: 모든 행 + 맞는 행의 '제외' 열에 이유를 적은 새 파일")
    parser.add_argument("--file", required=True, help="원본 데이터 파일 (.xlsx, .csv)")
    parser.add_argument("--where", help="고를 행의 조건 (예: '반응온도 >= 180; 라인 = B | 라인 = C'. ;는 그리고, |는 또는, "
                                        "연산자 >= <= > < = != ~)")
    parser.add_argument("--rows", help="고를 원본 행 번호 (쉼표 구분, 엑셀 행 번호·excluded_rows.csv의 원본 행, 예: '15,27')")
    parser.add_argument("--reason", help="exclude: '제외' 열에 적을 이유 (생략 시 조건 또는 행 번호)")
    parser.add_argument("--save", help="저장 경로 (.csv 기본, .xlsx 가능, 이름만 주면 원본 옆. 생략 시 <원본>_뽑기N.csv / <원본>_제외N.csv)")
    parser.add_argument("--overwrite", action="store_true", help="--save 파일이 이미 있으면 덮어쓰기 (생략 시 시각을 붙인 새 이름)")
    parser.add_argument("--dry-run", dest="dry_run", action="store_true", help="저장하지 않고 고른 행과 저장할 파일만 보여 줌")
    parser.add_argument("--header-row", type=int, help="데이터 머리글 행 번호 (생략 시 자동)")
    parser.add_argument("--unit-row", type=int, help="데이터 단위 행 번호 (선택)")
    parser.add_argument("--data-row", type=int, help="데이터 시작 행 번호 (선택)")
    args = parser.parse_args()
    run_cli(lambda: execute(args), hint="확인할 것: 데이터 파일 경로, 조건의 열 이름, 행 번호.")


if __name__ == "__main__":
    main()
