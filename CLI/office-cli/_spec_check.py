"""excel.py check의 내부 모듈 (진입점 아님, 스킬에 적지 않음): 엑셀 표의 값을 등록된 규격서와 한 번에 대조.

excel.py가 불러 씁니다 (excel.py check). 주소·읽기 도우미는 excel.py의 것을 씁니다.
"""
import os
import re
from typing import Any, Dict, List, Optional, Sequence, Tuple

import _office as off
import _runtime as rt
from excel import (TOOL, _read_matrix_closed, _read_matrix_csv, addr, closed_csv, col_letter, col_number,  # noqa: F401
                   is_closed_xlsx, out_value, parse_area, split_sheet)

# ==============================================================================
# 9-1. check: 엑셀 표의 값을 등록된 규격서와 한 번에 대조 (항목 찾기 -> 수치 조건 -> 판정 -> 메모 파일)
# ==============================================================================

CHECK_MAX_ROWS = 500           # 한 번에 대조하는 행 수 상한
CHECK_SHOWN = 20               # JSON에 담는 결과 수
CHECK_SPEC_CHARS = 120         # 결과의 규격 문장 글자 수
CHECK_MEMO_SPEC = 60           # 칸 메모에 넣는 규격 문장 앞부분 글자 수
VERDICTS = ("일치", "불일치", "근거없음", "판단필요")
_UNIT_IN_ITEM = re.compile(r"\s*[\(\[]\s*[^\)\]]{1,8}\s*[\)\]]\s*$")


def _col_arg(text: Optional[str], what: str) -> Optional[int]:
    if not text:
        return None
    t = text.strip().upper().replace("$", "")
    if not re.fullmatch(r"[A-Z]{1,3}", t):
        raise rt.InputError(f"{what}은(는) 열 문자로 적습니다 (예: B). 받은 값: {text}")
    return col_number(t)


def _check_rows(args, cols: List[int]) -> Tuple[str, int, List[List[Any]], Optional[List[Any]], Dict[str, Any]]:
    """대조할 행: (시트 이름, 첫 행 번호, 행마다 [cols의 값], 머리글(첫 행 위 한 줄 또는 None), target_resolved)."""
    sh, cells = split_sheet(args.range)
    r1, _, r2, _ = parse_area(cells, None, None)
    if r2 is None:
        raise rt.InputError("--range는 행 범위가 있는 주소로 적습니다 (예: B2:D50).")
    if r2 - r1 + 1 > CHECK_MAX_ROWS:
        raise rt.InputError(f"한 번에 {CHECK_MAX_ROWS}행까지 대조합니다 (받은 {r2 - r1 + 1}행). --range를 나눠 실행하세요.")
    c_lo, c_hi = min(cols), max(cols)
    top = r1 - 1 if r1 > 1 else r1
    head = getattr(args, "header_row", None)
    if head == 0:
        top = r1                                           # 0 = 머리글 없음 (남은 행 명령: 원래 범위가 1행부터, 1.4.71)
    elif head is not None:
        if not 1 <= head < r1:
            raise rt.InputError(f"--header-row는 --range의 첫 행({r1})보다 위의 행 번호(머리글이 없으면 0)로 적습니다. "
                                f"받은 값: {head}")
        top = head
    address = addr(top, c_lo, r2, c_hi)
    sheet_arg = sh or args.sheet
    closed = is_closed_xlsx(args.target)
    csvp = closed_csv(args.target)
    if closed or csvp:
        ref = (f"'{sheet_arg}'!" if sheet_arg else "") + address
        sheet, rr, cc, matrix, _ = (_read_matrix_closed(closed, sheet_arg, address) if closed
                                    else _read_matrix_csv(csvp, address))
        resolved = {"app": "file", "name": os.path.basename(closed or csvp), "path": closed or csvp, "sheet": sheet}
    else:
        with off.open_target("excel", args.target, readonly=True) as t:
            ws = t.sheet(sheet_arg)
            rng = ws.Range(ws.Cells(top, c_lo), ws.Cells(r2, c_hi))
            matrix = off.range_to_matrix(rng.Value)
            sheet = str(ws.Name)
            resolved = t.resolved()
    width = c_hi - c_lo + 1
    matrix = [list(r) + [None] * (width - len(r)) for r in matrix]
    pick = lambda row: [row[c - c_lo] for c in cols]  # noqa: E731
    header = pick(matrix[0]) if top < r1 and matrix else None
    body = [pick(r) for r in (matrix[r1 - top:] if top < r1 else matrix)]
    body += [[None] * len(cols)] * max(0, (r2 - r1 + 1) - len(body))
    return sheet, r1, body, header, resolved


def _judge_row(docs: List[Dict[str, Any]], item: str, value: Any, sec: str, unit_hint: Any) -> Dict[str, Any]:
    """한 행 판정: {verdict, reason, evidence, spec_text, condition, match_mode}.
    숫자가 든 항목 글을 표기가 다른 곳(층 B, literal False)에서만 찾았으면 reason에 표시하고 match_mode는 exact_mixed."""
    out = _judge_row_core(docs, item, value, sec, unit_hint)
    if out.pop("_literal", True) is False:
        out["reason"] = (out["reason"] + " / " if out["reason"] else "") + LAYER_B_NOTE
        out["match_mode"] = "exact_mixed"
    return out


LAYER_B_NOTE = "표기가 다른 곳에서 찾음 (항목 글이 글자 그대로는 없음: 부호·단위 확인)"


def _judge_row_core(docs: List[Dict[str, Any]], item: str, value: Any, sec: str, unit_hint: Any) -> Dict[str, Any]:
    import reader
    out: Dict[str, Any] = {"verdict": "근거없음", "reason": "", "evidence": "", "spec_text": "", "match_mode": "exact"}
    query = _UNIT_IN_ITEM.sub("", item).strip() or item
    mode, hits = reader.find_item(docs, query, reader.section_key(sec))
    out["match_mode"] = mode
    if not hits:
        if mode == "number":
            # 숫자 안전망이 시간 상한으로 규격서를 끝까지 보지 못함: '근거없음'으로 단정하지 않음
            out.update(verdict="판단필요", reason="숫자 안전망 검색이 시간 상한으로 규격서 일부만 확인함")
            return out
        out["reason"] = "규격서에서 항목을 찾지 못함" + (f" (장절 '{sec}' 안)" if sec else "")
        return out
    best = hits[0]
    out["_literal"] = best.get("literal", True) if mode == "exact" else True
    top = [h for h in hits if h["score"] == best["score"] and h.get("literal", True) == best.get("literal", True)][:5]
    conds: List[Dict[str, Any]] = []
    sentences: List[str] = []
    for h in top:
        cs, sent = reader.near_conditions(h["unit"]["text"], h["terms"])
        sentences.append(sent)
        for c in cs:
            if not any(cond_same(c, x) for x in conds):
                conds.append(c)
    out["evidence"] = reader.evidence_text(best["doc"]["file"], best["unit"])
    out["spec_text"] = re.sub(r"\s+", " ", sentences[0])[:CHECK_SPEC_CHARS]
    qty = reader.quantity_of(value)
    cell_unit = (qty[1] if qty else None) or reader.unit_of(unit_hint) or reader.unit_of(item)
    # 보수적 판정: 같은 항목 절에 합치지 못한 공차(±N)나 % 조건이 있으면 기준 값과 어떻게 합칠지 확실하지 않음
    # ('정격 전류 2 A, 허용 오차 5 %'에서 % 조건을 단위가 달라 버리면 2 A만 남아 2.05 A가 불일치로 단정됨)
    loose = [c for c in conds if c["kind"] == "pm" or (c["base"] == "%" and not (cell_unit and cell_unit[0] == "%"))]
    odd = next((c["ambiguous"] for c in conds if c.get("ambiguous")), "")
    if odd and mode == "exact":
        out.update(verdict="판단필요", reason=f"{odd}: 조건을 그대로 판정에 쓰지 않음", condition=_cond_list(conds))
        return out
    if loose and mode == "exact":
        out.update(verdict="판단필요", reason=f"같은 항목에 합치는 규칙이 확실하지 않은 조건이 있음: {_cond_list(conds)}",
                   condition=_cond_list(conds))
        return out
    if cell_unit and conds:
        same = [c for c in conds if c["base"] == cell_unit[0]]
        if not same:
            out.update(verdict="판단필요", reason=f"단위가 다름 (칸 {cell_unit[2]}, 규격 {', '.join(sorted({c['unit'] for c in conds}))})")
            return out
        conds = same
    if mode != "exact":
        out.update(verdict="판단필요", reason={"partial": "검색어 일부만 맞음(partial)",
                                              "number": "숫자 값만 같은 곳에서 찾음(number)"}.get(mode, "한 글자 다른 낱말로 찾음(fuzzy)"))
        if conds:
            out["condition"] = ", ".join(reader.cond_text(c) for c in conds[:3])
        return out
    if not conds:
        out.update(verdict="판단필요", reason="근거 문장에서 단위가 있는 수치 조건을 찾지 못함")
        return out
    if len(conds) > 1:
        out.update(verdict="판단필요", reason=f"같은 항목에 조건이 {len(conds)}개 (하나로 합치는 규칙이 확실하지 않음): {_cond_list(conds)}",
                   condition=_cond_list(conds))
        return out
    c = conds[0]
    out["condition"] = reader.cond_text(c)
    if qty is None:
        out.update(verdict="판단필요", reason="칸 값이 숫자가 아님")
        return out
    if not cell_unit:
        # 칸 값·값 열 머리글·항목에 단위가 없으면 규격 단위로 추정하지 않음 (30이 mm인지 cm인지 모름)
        out.update(verdict="판단필요", reason=f"칸 단위 없음 (규격은 {c['unit']}): 값 열 머리글에 단위를 적거나(예: 값({c['unit']})) 단위 열을 두세요(머리글 '단위' 또는 --unit-col)")
        return out
    v = qty[0] * cell_unit[1]
    ok_ = reader.cond_holds(c, v)
    alt = [a for a in c.get("alt") or [] if reader.cond_holds(a, v)]
    if not ok_ and alt:
        # 기준 값과 다르지만 같은 문장의 같은 단위 범위 안: 그 범위가 이 항목의 허용 범위일 수 있어 단정하지 않음
        out.update(verdict="판단필요", reason=f"기준 값 {out['condition']}와 다르지만 같은 문장의 범위 {reader.cond_text(alt[0])} 안",
                   condition=f"{out['condition']}, {reader.cond_text(alt[0])}")
        return out
    out.update(verdict="일치" if ok_ else "불일치", reason=out["condition"])
    return out


def _cond_list(conds: List[Dict[str, Any]]) -> str:
    import reader
    return ", ".join(reader.cond_text(c) for c in conds[:3]) + (" 등" if len(conds) > 3 else "")


def cond_same(a: Dict[str, Any], b: Dict[str, Any]) -> bool:
    return (a["kind"] == "pm") == (b["kind"] == "pm") and all(a.get(k) == b.get(k)
                                                              for k in ("base", "lo", "hi", "lo_strict", "hi_strict", "d"))


def _check_command(args, extra: Sequence[str] = ()) -> List[str]:
    argv = ["check", "--target", args.target, "--range", args.range, "--item-col", args.item_col,
            "--value-col", args.value_col, "--spec", args.spec]
    for flag, val in (("--section-col", args.section_col), ("--section", args.section), ("--sheet", args.sheet),
                      ("--unit-col", getattr(args, "unit_col", None)),
                      ("--header-row", str(args.header_row) if getattr(args, "header_row", None) is not None else None)):
        if val:
            argv += [flag, val]
    return argv + list(extra)


UNIT_HEADERS = ("단위", "unit", "units", "단위(unit)")


def cmd_check(args) -> Dict[str, Any]:
    import reader
    item_c = _col_arg(args.item_col, "--item-col")
    value_c = _col_arg(args.value_col, "--value-col")
    sec_c = _col_arg(args.section_col, "--section-col")
    unit_c = _col_arg(getattr(args, "unit_col", None), "--unit-col")
    if item_c is None or value_c is None:
        raise rt.InputError("--item-col(항목 열)과 --value-col(값 열)을 열 문자로 지정하세요 (예: --item-col B --value-col D).")
    cols = [item_c, value_c] + ([sec_c] if sec_c else [])
    # 단위 열: --unit-col, 생략하면 값 열 바로 오른쪽 열의 머리글이 '단위'이면 그 열 (항목·값·단위 3열 표, 회사 PC 2026-11-06 O1)
    auto_unit = unit_c is None and value_c + 1 not in cols
    unit_try = unit_c if unit_c is not None else (value_c + 1 if auto_unit else None)
    if unit_try is not None and unit_try in cols:
        raise rt.InputError("--unit-col은 항목·값·장절 열과 다른 열로 적습니다.")
    sheet, r1, body, header, resolved = _check_rows(args, cols + ([unit_try] if unit_try is not None else []))
    unit_idx = None
    if unit_try is not None:
        head_text = re.sub(r"\s+", "", str(header[-1] if header else "") or "").lower()
        if unit_c is not None or head_text in UNIT_HEADERS:
            unit_idx = len(cols)
            unit_c = unit_try
        else:
            unit_c = None
    docs, info = reader.load_spec_docs(args.spec)       # 등록 안 된 규격서는 먼저 등록 (오래 걸리면 CONTINUE)
    if not docs:
        raise rt.InputError(f"--spec에서 읽을 수 있는 규격 문서가 없습니다: {args.spec} (지원 형식: PDF, HWP, HWPX, DOCX, PPTX, XLSX, CSV, TXT, MD).")
    unit_hint = header[1] if header else None
    sheet_ref = sheet if re.fullmatch(r"[\w가-힣]+", sheet) else f"'{sheet}'"
    results: List[Dict[str, Any]] = []
    late: List[str] = []              # 시간 상한으로 대조하지 못한 칸 (판단필요, 메모 없음)
    late_row = 0                      # 대조하지 못한 첫 행 번호
    for i, row in enumerate(body):
        item, value = row[0], row[1]
        if item is None or str(item).strip() == "":
            continue
        row_unit = row[unit_idx] if unit_idx is not None else None
        hint = row_unit if row_unit is not None and str(row_unit).strip() else unit_hint   # 행의 단위 칸이 머리글보다 먼저
        sec = str(row[2]).strip() if sec_c and row[2] is not None else (args.section or "")
        if late or (results and rt.time_up()):
            # 명령 시간 제한 전에 결과를 내려고 남은 행은 대조하지 않음 (대조는 저장할 중간 결과가 없어 CONTINUE가 아님)
            late.append(f"{sheet_ref}!{addr(r1 + i, value_c)}")
            late_row = late_row or r1 + i
            j = {"verdict": "판단필요", "reason": "대조하지 못함 (명령 시간 제한)", "evidence": "", "spec_text": "",
                 "match_mode": None}
        else:
            j = _judge_row(docs, str(item).strip(), value, sec, hint)
        cell = f"{sheet_ref}!{addr(r1 + i, value_c)}"
        results.append({"cell": cell, "item": str(item).strip()[:60], "value": out_value(value), **j})
    counts = {v: sum(1 for r in results if r["verdict"] == v) for v in VERDICTS}
    not_decided = counts["판단필요"]
    # 메모 파일 (칸, 메모) -> excel.py comment --file로 칸 메모. 판정은 규격 문장을 기계로 읽은 참고 판정이라
    # 메모에 '자동 판정(참고)'과 읽은 조건·근거를 함께 적어 사람이 확인할 수 있게 함
    # 대조하지 못한 행(late)은 메모에 넣지 않음 (판정이 없는 칸에 '자동 판정' 메모가 달리지 않게)
    memo_rows = [["칸", "메모"]] + [[r["cell"], check_memo(r)] for r in results if r["cell"] not in late]
    memo = rt.write_csv_rows(rt.resolve_output_path(None, f"check_memo_{sheet}", ".csv"), memo_rows)
    problems: List[str] = []
    memo_problem = rt.check_saved_file(memo, len(memo_rows), delimiter=",")
    if memo_problem:
        problems.append(f"메모 파일: {memo_problem}")
    fields: Dict[str, Any] = {"target_resolved": resolved, "count": len(results), "counts": counts,
                              "results": [{k: r.get(k, "") for k in ("cell", "item", "value", "verdict", "reason", "condition",
                                                                     "spec_text", "evidence", "match_mode")}
                                          for r in results[:CHECK_SHOWN]],
                              "truncated": len(results) > CHECK_SHOWN, "shown": min(len(results), CHECK_SHOWN),
                              "needs_review": [r["cell"] for r in results if r["verdict"] == "판단필요"][:50],
                              "not_decided": not_decided, "memo_file": memo}
    if unit_c is not None:
        fields["unit_col"] = col_letter(unit_c) + (" (머리글이 '단위'라 자동으로 씀)" if auto_unit else "")
    if args.save:
        header_row = ["칸", "항목", "값", "판정", "이유", "규격 조건", "근거", "규격 문장", "찾은 방식"]
        rows = [[r["cell"], r["item"], r["value"], r["verdict"], r["reason"], r.get("condition", ""), r["evidence"],
                 r["spec_text"], r["match_mode"]] for r in results]
        out = rt.resolve_output_path(args.save, f"check_{sheet}", ".xlsx")
        out, _ = reader.write_xlsx_table(out, header_row, rows, [12, 24, 10, 10, 30, 18, 40, 60, 10], "대조결과")
        fields["output_file"] = out
        if len(results) > CHECK_SHOWN:
            fields["full_result_file"] = out
    elif len(results) > CHECK_SHOWN:
        fields["full_result_file"] = memo
    if info.get("remaining"):
        # 규격서 등록이 남음: 대조가 끝나지 않았으니 메모 미리보기 대신 같은 명령(이어서 등록하고 다시 대조)만 줌 (CONTINUE와 같은 흐름)
        nexts = [rt.next_command("excel.py", _check_command(args), f"규격서 {info['remaining']}개가 아직 등록되지 않아 다시 대조 (같은 명령)")]
        problems.append(f"규격 폴더의 문서 {info['remaining']}개를 아직 등록하지 못해 대조에 쓰지 않음")
        action = ("대조가 끝나지 않았습니다: 사용자에게 지금까지의 결과를 알리고 next_commands[0](같은 명령)을 그대로 다시 실행해 "
                  "남은 규격서를 등록하고 이어서 대조하세요. 메모는 대조를 마친 결과의 next_commands로 합니다.")
    elif os.path.splitext(str(args.target or ""))[1].lower() in (".csv", ".tsv"):
        # CSV에는 메모를 달 수 없음 (메모 명령을 주던 것, 2026-11-09 검토 B13)
        nexts = []
        action = "CSV 파일이라 칸 메모는 달 수 없습니다. 결과(full_result_file이 있으면 그 파일)를 사용자에게 알리세요. 같은 check는 다시 실행하지 마세요."
    else:
        nexts = [rt.next_command("excel.py", ["comment", "--target", args.target, "--action", "add", "--file", memo, "--dry-run"],
                                 "자동 판정(참고)·읽은 조건·근거를 값 칸 메모로 달기 미리보기, 변경 없음 "
                                 "(실제로 다는 명령은 그 결과의 confirm_commands, 사용자 확인 뒤에만)")]
        action = ("사용자가 메모를 원하면 next_commands[0](미리보기, 변경 없음)을 실행해 보여 주고, 사용자가 '예'라고 한 뒤에만 "
                  "그 결과의 confirm_commands[0]을 실행하세요. 같은 check는 다시 실행하지 마세요.")
    if late:
        problems.append(f"명령 시간 제한으로 {len(late)}행({late[0]}부터)은 대조하지 못함 (판단필요로 두고 메모에서 뺌)")
        if not info.get("remaining"):
            # 남은 행만 대조하는 전체 명령 (머리글 행은 원래 범위 위 행으로 고정)
            _, cells = split_sheet(args.range)
            _, c1, r2, c2 = parse_area(cells, None, None)
            rest = ((args.range.rsplit("!", 1)[0] + "!") if "!" in args.range else "") + \
                addr(late_row, c1 or min(cols), r2, c2 or max(cols))
            # 원래 범위가 1행부터면 머리글이 없었음: 0(머리글 없음)을 넘겨 남은 행 범위 위의 데이터 행을 머리글로 읽지 않게
            head_row = getattr(args, "header_row", None)
            head_row = head_row if head_row is not None else (r1 - 1 if r1 > 1 else 0)
            rest_args = type(args)(**{**vars(args), "range": rest, "header_row": head_row})
            nexts.append(rt.next_command("excel.py", _check_command(rest_args),
                                         f"대조하지 못한 {len(late)}행({rest})만 다시 대조 (읽기 전용)"))
            action += (f" 대조하지 못한 {len(late)}행은 사용자에게 알리고, 사용자에게 확인을 받은 뒤 "
                       f"next_commands[{len(nexts) - 1}]을 그대로 실행하세요.")
    if info.get("skipped"):
        problems.append(f"읽지 못한 규격 문서 {len(info['skipped'])}개: " + ", ".join(os.path.basename(s["file"]) for s in info["skipped"][:5]))
    fields["next_commands"] = nexts
    preview = (f"{len(results)}행 대조(자동 판정, 참고용): 일치 {counts['일치']}, 불일치 {counts['불일치']}, "
               f"근거없음 {counts['근거없음']}, 판단필요 {counts['판단필요']}")
    if not results:
        fields.pop("next_commands", None)
        return rt.ok_unverified(TOOL, "check", "0건: 대조할 항목이 없습니다",
                                [f"0건(찾은 항목 없음): {args.range}의 항목 열({args.item_col})이 비어 있음"], **fields)
    # 판정은 자연어 규격 문장을 기계로 해석한 참고 판정 (해석은 휴리스틱): 늘 UNVERIFIED, 사용자 확인 뒤 메모
    problems.insert(0, CHECK_ADVISORY)
    if not_decided:
        problems.insert(1, f"판단필요 {not_decided}행(needs_review)은 도구가 판정하지 못함: 사용자가 근거를 보고 정해야 함")
    payload = rt.ok_unverified(TOOL, "check", preview, problems, **fields)
    # 기본 문구('다시 실행하지 말고 그대로 알리세요')는 메모 미리보기·이어서 대조와 부딪혀 이 명령의 행동으로 바꿈
    payload["agent_directive"] = "UNVERIFIED: 실행했지만 확인하지 못했습니다: " + "; ".join(problems) + ". " + action
    return payload


CHECK_ADVISORY = ("자동 판정은 참고용입니다(규격 문장을 기계로 해석함). 행마다 읽은 조건(condition)과 근거 문장(spec_text)을 "
                  "사용자에게 보여 주고 확인을 받은 뒤 메모를 다세요")


def check_memo(r: Dict[str, Any]) -> str:
    """check 한 행의 칸 메모: '자동 판정(참고): 불일치 - 읽은 조건: 3 A 이하 - 근거: 규격서.txt 3.2절: 소비 전류는 ...'."""
    cond = r.get("condition") or "없음"
    memo = f"자동 판정(참고): {r['verdict']}"
    if r.get("reason") and r["reason"] != r.get("condition"):
        memo += f" ({r['reason']})"
    memo += f" - 읽은 조건: {cond}"
    memo += f" - 근거: {r['evidence']}: {r['spec_text'][:CHECK_MEMO_SPEC]}" if r.get("evidence") else " - 근거: 없음"
    return memo
