"""excel.py 쓰기: write, style, range(merge·copy·clear·insert·delete·sort), sheet, chart, replace와
그 실행이 남기는 되돌리기 지점.
"""
import json
import os
import shlex

import pytest

import _office as off
import _runtime as rt
import excel
import office
from fake_excel import FakeExcelWorkbook, FakeXlApp, SheetNameRefused
from fake_office import make_win32com
from office_helpers import EMPTY_FIND, run_excel, run_main, sheet_names, undo_points, ws_of


def _no_create_imperative(out):
    # 시나리오 10: message나 [후보]에 --new-sheet / '만드세요'가 있으면 경량 LLM이 묻지 않고 새 시트를 만든다
    texts = [out["message"]] + list(out.get("candidates", []))
    assert not any("--new-sheet" in x or "만드세요" in x or "붙이세요" in x for x in texts), texts


class _T:
    def __init__(self, wb):
        self.doc = wb

    def sheet(self, name):
        return self.doc.Worksheets(name)


def _fill(wb):
    for ws in wb.sheets:
        for r, row in enumerate([["월", "값"], ["1월", 3], ["2월", 5]], 1):
            for c, v in enumerate(row, 1):
                ws.set(r, c, v)


def _names(wb):
    return {str(ws.Name): [str(o.Name) for o in excel._chart_objects(ws)] for ws in wb.sheets}


def _charts(wb):
    return {str(ws.Name): [(str(o.Name), excel._chart_title(o)) for o in excel._chart_objects(ws)] for ws in wb.sheets}


def _formula1(wb):
    ws = wb.Worksheets("Sheet1")
    return [fc.kw["Formula1"] for fcs in ws.fcs.values() for fc in fcs.items]


def _col(ws, c, r1=1, r2=5):
    return [ws.values.get((r, c)) for r in range(r1, r2 + 1)]


def _tool_log_text():
    p = rt.error_log_path()
    return open(p, encoding="utf-8").read() if os.path.exists(p) else ""


# ---------------------------------------------------------------------------
# write
# ---------------------------------------------------------------------------


def test_write_sine_header_and_column_formulas(xl, capsys):
    app, wb = xl
    code, out = run_excel(capsys, "write", "--range", "A1:B101", "--header", "x,sin(x)",
                          "--formula", "=(ROW()-2)*4*PI()/99", "=SIN(A2)")
    assert code == 0 and out["range_written"] == "A1:B101" and out["rows"] == 101
    ws = ws_of(wb)
    assert ws.values[(1, 1)] == "x" and ws.values[(1, 2)] == "sin(x)"
    assert ("Formula", "$A$2:$A$101", "=(ROW()-2)*4*PI()/99") in ws.log
    assert ("Formula", "$B$2:$B$101", "=SIN(A2)") in ws.log
    assert app.ScreenUpdating is True and app.Calculation == -4105       # 대량 쓰기 설정 복원
    assert "chart" in out["next_commands"][0]["command"]
    assert out["target_resolved"]["sheet"] == "Sheet1"


def test_write_formula_count_mismatch_is_input_error(xl, capsys):
    code, out = run_excel(capsys, "write", "--range", "A2:C10", "--formula", "=1", "=2")
    assert code == 1 and out["error_type"] == "input"


def test_write_checks_everything_before_writing_the_header(xl, capsys):
    # 회사 PC 오류 로그: 헤더를 먼저 쓰고 수식 검사에서 실패해 A1:B1만 남았음. 추천 범위는 머리글 행부터
    app, wb = xl
    code, out = run_excel(capsys, "write", "--range", "A1", "--header", "x,sin(x)", "--formula", "=(ROW()-2)*0.1", "=SIN(A2)")
    assert code == 1 and out["message"].endswith("[추천] --range A1:B101")
    assert (1, 1) not in ws_of(wb).values and (1, 2) not in ws_of(wb).values


def test_write_numbers_given_as_formulas_points_to_values(xl, capsys):
    code, out = run_excel(capsys, "write", "--range", "C6:C8", "--formula", "73.42", "49.97", "14.7")
    assert code == 1 and "[추천] --values '[[73.42], [49.97], [14.7]]'" in out["message"]


def test_write_over_existing_values_needs_overwrite(xl, capsys):
    app, wb = xl
    ws = ws_of(wb)
    ws.set(2, 2, "keep")
    code, out = run_excel(capsys, "write", "--range", "A1", "--values", '[["a","b"],["c","new"]]')
    assert code == 1 and out["error_type"] == "needs_user_decision" and "B2=keep" in out["message"]
    assert ws.values[(2, 2)] == "keep" and (1, 1) not in ws.values
    code, out = run_excel(capsys, "write", "--range", "A1", "--values", '[["a","b"],["c","new"]]', "--overwrite")
    assert code == 0 and ws.values[(2, 2)] == "new"


def test_write_reads_back_and_reports_mismatch(xl, capsys, monkeypatch):
    # R1: VERIFIED는 쓴 칸을 다시 읽어 같을 때만. Excel이 값을 바꿨으면 UNVERIFIED와 그 칸
    app, wb = xl
    code, out = run_excel(capsys, "write", "--range", "A1", "--values", '[["코드","값"],["001",7]]')
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:") and out["verified"] is True
    ws = ws_of(wb)
    real_set = ws.set
    monkeypatch.setattr(ws, "set", lambda r, c, v: real_set(r, c, 999 if (r, c) == (2, 5) else v))
    code, out = run_excel(capsys, "write", "--range", "D1", "--values", '[["x","y"],[1,2]]')
    assert code == 0 and out["agent_directive"].startswith("UNVERIFIED:") and "E2" in out["unverified"][0]


def test_write_csv_new_sheet_and_append_skips_header(xl, capsys, tmp_path):
    app, wb = xl
    f = tmp_path / "d.csv"
    f.write_text("코드,금액\n001,1200\n002,800\n", encoding="utf-8")
    code, out = run_excel(capsys, "write", "--file", str(f), "--sheet", "검색결과", "--new-sheet")
    assert code == 0 and out["sheet"] == "검색결과" and out["range_written"] == "A1:B3"
    ws = ws_of(wb, "검색결과")
    assert ws.values[(2, 1)] == "001" and ws.values[(2, 2)] == 1200
    f2 = tmp_path / "e.csv"
    f2.write_text("코드,금액\n003,500\n", encoding="utf-8")
    code, out = run_excel(capsys, "write", "--file", str(f2), "--sheet", "검색결과", "--append")
    assert out["range_written"] == "A4:B4" and ws.values[(4, 1)] == "003"


def test_write_missing_sheet_is_not_created_without_flag(xl, capsys):
    app, wb = xl
    code, out = run_excel(capsys, "write", "--sheet", "결과표", "--value", "1")
    assert code == 1 and out["error_type"] == "needs_user_decision"
    assert out["agent_directive"].startswith("STOP_RETRY: 다시 실행하지 마세요.")
    assert "만들지 물어보세요" in out["agent_directive"]
    assert "만들라고 답한 경우에만 같은 명령에 --new-sheet" in out["agent_directive"]   # 플래그는 사용자 동의 조건으로만
    _no_create_imperative(out)
    assert "[추천]" not in out["message"] and "[후보]" not in out["message"] and "candidates" not in out
    assert "있는 시트: Sheet1, 요약" in out["message"] and "사용자에게 물어보세요" in out["message"]
    assert sheet_names(wb) == ["Sheet1", "요약"]                       # 새 시트 없음
    code, out = run_excel(capsys, "write", "--range", "'결과표'!A1", "--value", "1")
    assert code == 1 and sheet_names(wb) == ["Sheet1", "요약"]


def test_write_sheet_differs_only_in_spacing_recommends_existing(xl, capsys):
    app, wb = xl
    code, out = run_excel(capsys, "write", "--sheet", "Sheet 1", "--value", "1")
    assert code == 1 and out["error_type"] == "input" and out["message"].endswith("[추천] Sheet1")
    _no_create_imperative(out)
    assert sheet_names(wb) == ["Sheet1", "요약"]


def test_write_similar_but_different_sheet_asks_user(xl, capsys):
    # Sheet2를 Sheet1의 오타로 단정해 덮어쓰지 않음: [후보]는 실제 기존 시트만, 새 시트 여부는 사용자에게 물음
    app, wb = xl
    code, out = run_excel(capsys, "write", "--sheet", "Sheet2", "--value", "1")
    assert code == 1 and out["error_type"] == "needs_user_decision" and out["message"].endswith("[후보] Sheet1")
    assert out["candidates"] == ["Sheet1"]
    assert out["agent_directive"].startswith("STOP_RETRY: 다시 실행하지 마세요.")
    assert "만들라고 답한 경우에만" in out["agent_directive"]
    _no_create_imperative(out)
    assert sheet_names(wb) == ["Sheet1", "요약"] and (1, 1) not in ws_of(wb, "Sheet1").values


def test_scenario10_blank_workbook_sheet9_asks_user(capsys, monkeypatch):
    # 시나리오 10: Sheet1만 있는 빈 통합문서에 "Sheet9에 써줘" -> 만들지 않고 사용자에게 물음
    app = FakeXlApp()
    wb = FakeExcelWorkbook(app, "Book1", sheets=("Sheet1",))
    wb.FullName = "Book1"                                                # 저장한 적 없는 새 통합문서
    fake, _ = make_win32com()
    monkeypatch.setattr(off, "win32com", fake)
    monkeypatch.setattr(off, "pythoncom", None)
    monkeypatch.setattr(off, "list_open", lambda kind: [])
    monkeypatch.setattr(off, "get_active", lambda kind: {"name": wb.Name, "path": "", "app": app, "doc": wb})
    code, out = run_excel(capsys, "write", "--sheet", "Sheet9", "--value", "테스트")
    assert code == 1 and out["error_type"] == "needs_user_decision"
    assert out.get("candidates", []) == ["Sheet1"] and "[추천]" not in out["message"]
    _no_create_imperative(out)
    assert "Sheet9" in out["agent_directive"] and "물어보세요" in out["agent_directive"]
    assert sheet_names(wb) == ["Sheet1"]


def test_write_new_sheet_flag_rules(xl, capsys):
    app, wb = xl
    code, out = run_excel(capsys, "write", "--sheet", "요약", "--new-sheet", "--value", "1")
    assert code == 1 and "이미 있어" in out["message"] and (1, 1) not in ws_of(wb, "요약").values
    code, out = run_excel(capsys, "write", "--new-sheet", "--value", "1")
    assert code == 1 and "새 시트 이름" in out["message"]
    code, out = run_excel(capsys, "write", "--range", "'새시트'!B2", "--new-sheet", "--value", "7")
    assert code == 0 and out["sheet"] == "새시트" and ws_of(wb, "새시트").values[(2, 2)] == 7


def test_write_requires_content(xl, capsys):
    code, out = run_excel(capsys, "write", "--range", "A1")
    assert code == 1 and "--value" in out["message"]


# ---------------------------------------------------------------------------
# 데이터 손상 방지
# ---------------------------------------------------------------------------


def test_text_that_excel_would_reinterpret_is_kept_as_text(xl, capsys, tmp_path):
    app, wb = xl
    f = tmp_path / "코드.csv"
    f.write_text("코드,품번,금액,큰수,지수,날짜\n01234,1-2,(100),12345678901234567890,1E5,2026-10-28\n", encoding="utf-8")
    code, out = run_excel(capsys, "write", "--file", str(f))
    assert code == 0, out
    ws = ws_of(wb)
    assert [ws.values[(2, c)] for c in range(1, 7)] == ["01234", "1-2", "(100)", "12345678901234567890", "1E5", "2026-10-28"]
    assert {(2, c) for c in range(1, 6)} <= ws.prefixed and (2, 6) not in ws.prefixed     # 연도가 있는 날짜는 날짜로 변환되게 둠
    assert excel.coerce("12345678901234567890") == "12345678901234567890" and excel.coerce("1200") == 1200


def test_replace_result_that_looks_numeric_stays_text(xl, capsys):
    app, wb = xl
    ws = ws_of(wb)
    ws.set(1, 1, "AB-001")
    code, out = run_excel(capsys, "replace", "--find", "AB-", "--replace", "")
    assert code == 0 and ws.values[(1, 1)] == "001" and (1, 1) in ws.prefixed


def test_range_clear_requires_range(xl, capsys):
    app, wb = xl
    ws = ws_of(wb)
    ws.set(1, 1, "중요")
    code, out = run_excel(capsys, "range", "--action", "clear")
    assert code == 1 and "--range" in out["message"] and ws.values[(1, 1)] == "중요"


def test_merge_with_several_values_needs_overwrite(xl, capsys):
    app, wb = xl
    ws = ws_of(wb)
    for c, v in enumerate(["이름", "부서", "직급"], 1):
        ws.set(1, c, v)
    code, out = run_excel(capsys, "range", "--action", "merge", "--range", "A1:C1")
    assert code == 1 and out["error_type"] == "needs_user_decision" and "B1=부서" in out["message"]
    assert not [x for x in ws.log if x[0] == "Merge"]
    code, out = run_excel(capsys, "range", "--action", "merge", "--range", "A1:C1", "--overwrite")
    assert code == 0 and ("Merge", "$A$1:$C$1", False) in ws.log
    assert out["agent_directive"].startswith("VERIFIED:")                 # 병합 상태를 다시 읽음 (R1)


def test_copy_over_existing_values_needs_overwrite(xl, capsys):
    app, wb = xl
    src, dst = ws_of(wb), ws_of(wb, "요약")
    src.set(1, 1, "원본")
    dst.set(2, 2, "기존 값")
    code, out = run_excel(capsys, "range", "--action", "copy", "--range", "A1:B2", "--dest", "요약!A1")
    assert code == 1 and out["error_type"] == "needs_user_decision" and "B2=기존 값" in out["message"]
    assert not [x for x in src.log if x[0] == "Copy"]
    code, out = run_excel(capsys, "range", "--action", "copy", "--range", "A1:B2", "--dest", "요약!D1")
    assert code == 0 and ("Copy", "$A$1:$B$2", "$D$1") in src.log
    assert out["agent_directive"].startswith("VERIFIED:") and dst.values[(1, 4)] == "원본"     # 대상을 다시 읽음 (R1)


def test_hide_row_address_hides_rows_not_columns(xl, capsys):
    import types
    assert excel._axis(types.SimpleNamespace(axis=None, range="5:7")) == "rows"
    assert excel._axis(types.SimpleNamespace(axis=None, range="'시트 1'!$5:$7")) == "rows"
    assert excel._axis(types.SimpleNamespace(axis=None, range="C:E")) == "cols"
    assert excel._axis(types.SimpleNamespace(axis="cols", range="5:7")) == "cols"


def test_append_uses_lowest_row_of_all_columns_and_start_row_header(xl, capsys, tmp_path):
    app, wb = xl
    ws = ws_of(wb)
    # C5부터 머리글, A열은 비어 있고 D열이 더 길게 차 있는 표
    for r, row in enumerate([["품목", "수량"], ["사과", 3], [None, 5]], 5):
        for c, v in enumerate(row, 3):
            if v is not None:
                ws.set(r, c, v)
    f = tmp_path / "추가.csv"
    f.write_text("품목,수량\n배,2\n", encoding="utf-8")
    code, out = run_excel(capsys, "write", "--file", str(f), "--range", "C5", "--append")
    assert code == 0, out
    assert ws.values[(8, 3)] == "배" and ws.values[(8, 4)] == 2 and ws.values[(7, 4)] == 5     # D7을 덮어쓰지 않음
    assert out["range_written"] == "C8:D8"                                                   # 머리글(5행)을 다시 쓰지 않음
    code, out = run_excel(capsys, "write", "--file", str(f), "--range", "C5", "--append", "--header", "품목,수량")
    assert code == 0 and out["range_written"] == "C9:D9" and ws.values[(9, 3)] == "배"          # --header도 같은 머리글이면 생략


# ---------------------------------------------------------------------------
# replace의 시트, 빈 바꾸기 거절, merge 값 판단, 수식 덮어쓰기, --append 열, 상대 수식 복사, autofit
# ---------------------------------------------------------------------------


def test_replace_writes_to_sheet_named_in_range(xl, capsys):
    app, wb = xl
    s1, s2 = wb.Worksheets("Sheet1"), wb.Worksheets("요약")
    s1.set(1, 1, "초안")
    s2.set(1, 1, "초안 보고")
    code, out = run_main(excel, capsys, "replace", "--find", "초안", "--replace", "최종", "--range", "요약!A1:A2")
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:")
    assert s2.values[(1, 1)] == "최종 보고" and s1.values[(1, 1)] == "초안"       # 이전: 지금 시트(Sheet1)에 씀
    code, out = run_main(excel, capsys, "replace", "--find", "초안", "--replace", "x", "--range", "요약!A1", "--all-sheets")
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:") and "--all-sheets" in out["message"]


def test_replace_without_replace_or_format_is_refused(xl, capsys):
    app, wb = xl
    wb.Worksheets("Sheet1").set(1, 1, "가")
    code, out = run_main(excel, capsys, "replace", "--find", "가")
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:") and "--replace" in out["message"]
    code, out = run_main(excel, capsys, "replace", "--find", "가", "--bold")
    assert code == 0 and out["changed_cells"] == 1


def test_merge_single_value_outside_top_left_needs_decision(xl, capsys):
    app, wb = xl
    ws = wb.Worksheets("Sheet1")
    ws.set(1, 2, "남는값")                                   # A1은 비고 B1에만 값
    code, out = run_main(excel, capsys, "range", "--action", "merge", "--range", "A1:B1")
    assert code == 1 and out["error_type"] == "needs_user_decision" and "B1" in out["message"]
    assert "$A$1:$B$1" not in ws.merged
    code, out = run_main(excel, capsys, "range", "--action", "merge", "--range", "A1:B1", "--overwrite")
    assert code == 0 and "$A$1:$B$1" in ws.merged


def test_formula_with_empty_result_is_not_overwritten(xl, capsys):
    app, wb = xl
    ws = wb.Worksheets("Sheet1")
    ws.formulas[(1, 1)] = '=IF(B1="","",B1)'
    ws.values[(1, 1)] = ""                                     # 수식 결과가 빈 글
    code, out = run_main(excel, capsys, "write", "--range", "A1", "--value", "5")
    assert code == 1 and out["error_type"] == "needs_user_decision" and ws.formulas.get((1, 1))
    assert excel._cell_empty(ws.Cells(1, 1)) is False and excel._cell_empty(ws.Cells(9, 9)) is True


def test_append_formula_scans_all_written_columns(xl, capsys):
    app, wb = xl
    ws = wb.Worksheets("Sheet1")
    for r in (1, 2, 3):
        ws.set(r, 1, r)
    for r in (1, 2, 3, 4):
        ws.set(r, 2, r * 10)                                   # B열이 한 행 더 김
    code, out = run_main(excel, capsys, "write", "--append", "--range", "A1:B1", "--formula", "=ROW()", "=ROW()*2")
    assert code == 0, out
    assert ws.values[(4, 2)] == 40 and out["range_written"] == "A5:B5"      # 이전: 4행에 써서 B4를 덮음
    assert ws.formulas[(5, 1)] == "=ROW()" and ws.formulas[(5, 2)] == "=ROW()*2"


def test_copy_relative_formula_is_verified(xl, capsys):
    app, wb = xl
    ws = wb.Worksheets("Sheet1")
    ws.set(1, 1, 3)
    ws.Range("B1").Formula = "=A1*2"
    code, out = run_main(excel, capsys, "range", "--action", "copy", "--range", "A1:B1", "--dest", "D1")
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    assert ws.formulas[(1, 5)] == "=D1*2"                     # 상대 참조가 바뀌어 값은 다르지만 같은 수식


def test_autofit_is_unverified(xl, capsys):
    app, wb = xl
    wb.Worksheets("Sheet1").set(1, 1, "가")
    code, out = run_main(excel, capsys, "range", "--action", "autofit", "--range", "A1")
    assert code == 0 and out["verified"] is False and out["agent_directive"].startswith("UNVERIFIED:")


# ---------------------------------------------------------------------------
# write --new-sheet는 모든 확인이 끝난 뒤에 시트를 만듦 (확인에서 멈추면 시트도 지점도 남지 않음)
# ---------------------------------------------------------------------------


def test_new_sheet_is_not_created_when_input_check_fails(xl, capsys):
    _, wb = xl
    code, out = run_main(excel, capsys, "write", "--range", "A1", "--value", "1")
    assert code == 0, out
    before = undo_points()
    code, out = run_main(excel, capsys, "write", "--new-sheet", "--sheet", "새시트", "--formula", "5")
    assert code == 1 and "[추천] --value 5" in out["message"], out
    assert "새시트" not in sheet_names(wb) and undo_points() == before
    code, out = run_main(excel, capsys, "write", "--new-sheet", "--sheet", "새시트", "--value", "5")   # 안내대로 다시 실행
    assert code == 0 and out["sheet"] == "새시트", out
    assert len(undo_points()) == len(before) + 1


def test_new_sheet_with_bad_range_creates_nothing(xl, capsys):
    _, wb = xl
    code, out = run_main(excel, capsys, "write", "--new-sheet", "--sheet", "새시트", "--range", "A1", "--formula", "=A1", "=B1")
    assert code == 1 and "--range" in out["message"]
    assert "새시트" not in sheet_names(wb) and undo_points() == []


# ---------------------------------------------------------------------------
# --append --formula 상대 참조
# ---------------------------------------------------------------------------


def test_append_formula_rebased_to_target_rows(xl, capsys):
    app, wb = xl
    ws = wb.Worksheets("Sheet1")
    for r, row in enumerate([["월", "값", "두배"], ["1월", 3, 6], ["2월", 5, 10], ["3월", 7, None], ["4월", 9, None]], 1):
        for c, v in enumerate(row, 1):
            if v is not None:
                ws.set(r, c, v)
    code, out = run_main(excel, capsys, "write", "--range", "C2:C3", "--formula", "=B2*2", "--append")
    assert code == 0 and out["range_written"] == "C4:C5", out
    assert ws.Cells(4, 3).Formula == "=B4*2"                      # 원래 행(B2)이 아니라 쓴 행 기준
    assert out["agent_directive"].startswith("VERIFIED")
    code, out = run_main(excel, capsys, "write", "--range", "C2", "--formula", "=$B$2+B2", "--append")
    assert code == 0 and ws.Cells(6, 3).Formula == "=$B$2+B6", out    # $ 참조는 그대로


def test_append_formula_conversion_failure_rejected(xl, capsys, monkeypatch):
    app, wb = xl
    ws = wb.Worksheets("Sheet1")
    for r, row in enumerate([["값", "두배"], [3, 6]], 1):
        for c, v in enumerate(row, 1):
            ws.set(r, c, v)

    def boom(**kw):
        raise Exception("COM")
    monkeypatch.setattr(app, "ConvertFormula", boom, raising=False)
    code, out = run_main(excel, capsys, "write", "--range", "B2", "--formula", "=A2*2", "--append")
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY") and "[추천]" in out["message"], out
    assert ws.Cells(3, 2).Formula in ("", None)


# ---------------------------------------------------------------------------
# replace / style / range / sheet
# ---------------------------------------------------------------------------


def test_replace_keeps_formula_cells(xl, capsys):
    app, wb = xl
    ws = ws_of(wb)
    ws.set(1, 1, "초안 보고")
    ws.set(1, 2, 3.0)
    ws.formulas[(1, 2)] = "=1+2"
    ws.set(1, 3, 3.0)
    code, out = run_excel(capsys, "replace", "--find", "3", "--replace", "9", "--whole", "--bg", "yellow")
    assert code == 0 and out["changed_cells"] == 2
    assert ws.formulas[(1, 2)] == "=1+2" and ws.values[(1, 3)] == 9
    assert out["formula_cells_kept"] == ["Sheet1!B1"]
    code, out = run_excel(capsys, "replace", "--map", '{"초안": "최종"}')
    assert ws.values[(1, 1)] == "최종 보고" and out["agent_directive"].startswith("VERIFIED:")     # 다시 읽어 확인 (R1)


def test_replace_no_match_is_ok_with_stop_hint(xl, capsys):
    code, out = run_excel(capsys, "replace", "--find", "없는말", "--replace", "x")
    assert code == 0 and out["changed_cells"] == 0 and "재시도하지" in out["agent_directive"]


def test_style_theme_zebra_and_highlight(xl, capsys):
    app, wb = xl
    ws = ws_of(wb)
    rows = [["지역", "금액"]] + [[f"지역{i}", float(i * 100)] for i in range(1, 8)] + [["합계", 2800.0]]
    for r, row in enumerate(rows, 1):
        for c, v in enumerate(row, 1):
            ws.set(r, c, v)
    code, out = run_excel(capsys, "style", "--theme", "navy", "--highlight", ">=500:pink")
    assert code == 0
    applied = out["applied"]
    assert applied["summary_row"] is True and applied["number_columns"] == ["B"]
    assert applied["zebra_rows"] == 3
    assert any(k[0].startswith("$A$3:$B$3,$A$5:$B$5") for k in ws.styles)
    fc = ws.fcs["$A$1:$B$9"].items[0]
    assert fc.kw["Type"] == 1 and fc.kw["Operator"] == 7
    assert out["agent_directive"].startswith("VERIFIED:") and out["verified"] is True        # 다시 읽어 확인 (R1)


def test_style_unknown_theme_and_color(xl, capsys):
    code, out = run_excel(capsys, "style", "--theme", "nvy")
    assert code == 1 and "[추천] navy" in out["message"]
    code, out = run_excel(capsys, "style", "--bg", "노랑")
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:")


def test_range_move_uses_cut_and_merge_rules(xl, capsys):
    app, wb = xl
    ws = ws_of(wb)
    ws.set(1, 1, "가")
    code, out = run_excel(capsys, "range", "--action", "move", "--range", "A1:B3", "--dest", "B2")
    assert code == 0 and ("Cut", "$A$1:$B$3", "$B$2") in ws.log
    code, out = run_excel(capsys, "range", "--action", "merge")
    assert code == 1
    run_excel(capsys, "range", "--action", "merge", "--range", "A1:B1")
    assert ("Merge", "$A$1:$B$1", False) in ws.log and app.DisplayAlerts is True


def test_sheet_actions(xl, capsys):
    app, wb = xl
    code, out = run_excel(capsys, "sheet", "--action", "add", "--name", "sheet1")
    assert code == 0 and "이미 있어" in out["preview"] and out["sheets"] == ["Sheet1", "요약"]
    code, out = run_excel(capsys, "sheet", "--action", "rename", "--name", "요약", "--new-name", "Sheet1")
    assert code == 1
    code, out = run_excel(capsys, "sheet", "--action", "delete", "--name", "요약")
    assert code == 0 and out["sheets"] == ["Sheet1"] and out["agent_directive"].startswith("VERIFIED:")   # 목록 다시 읽음 (R1)
    code, out = run_excel(capsys, "sheet", "--action", "delete", "--name", "Sheet1")
    assert code == 1 and "마지막" in out["message"]


# ---------------------------------------------------------------------------
# style은 --highlight 등 입력을 서식을 지우기 전에 확인
# ---------------------------------------------------------------------------


def test_style_bad_highlight_does_not_clear_formats(xl, capsys):
    _, wb = xl
    ws = wb.Worksheets(1)
    code, out = run_main(excel, capsys, "style", "--range", "A1:B2", "--clear", "--highlight", "bad")
    assert code == 1 and "--highlight" in out["message"]
    assert not [x for x in ws.log if x[0] == "ClearFormats"] and undo_points() == []


# ---------------------------------------------------------------------------
# --highlight 따옴표·와일드카드
# ---------------------------------------------------------------------------


def test_highlight_escapes_quote_and_wildcards(xl, capsys):
    app, wb = xl
    _fill(wb)
    code, out = run_main(excel, capsys, "style", "--range", "A1:B3", "--highlight", '포함:a"b*?~:yellow')
    assert code == 0, out
    assert _formula1(wb)[-1] == '=ISNUMBER(SEARCH("a""b~*~?~~",A1))'
    code, out = run_main(excel, capsys, "style", "--range", "A1:B3", "--highlight", '=say "hi":pink')
    assert code == 0, out
    assert _formula1(wb)[-1] == '="say ""hi"""'


# ---------------------------------------------------------------------------
# 숨기기 해제와 전체·서식 지우기는 바뀐 것으로 봄
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("argv", [
    ["range", "--action", "unhide", "--range", "B:C"],
    ["range", "--action", "clear", "--range", "D10:E12"],
])
def test_unhide_and_full_clear_are_treated_as_changes(xl_s2, capsys, argv):
    code, out = run_main(excel, capsys, *argv)
    assert code == 0 and out.get("unchanged") is not True, out
    assert "바뀐 것 없음" not in out["preview"], out


# ---------------------------------------------------------------------------
# range insert / delete
# ---------------------------------------------------------------------------


def test_insert_rows_shifts_down_and_verifies(xl_table, capsys):
    code, out = run_main(excel, capsys, "range", "--action", "insert", "--range", "3:4")
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    assert _col(xl_table, 1, 1, 7) == ["이름", "가", None, None, "나", "다", "라"], _col(xl_table, 1, 1, 7)
    assert out["moved"] == {"from": "A3", "to": "A5"} and out["count"] == 2 and out["axis"] == "rows", out
    assert len(off.list_snapshots("excel")) == 1                                  # 되돌리기 지점


def test_insert_cols(xl_table, capsys):
    code, out = run_main(excel, capsys, "range", "--action", "insert", "--range", "B")
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    assert xl_table.values[(1, 3)] == "부서" and (1, 2) not in xl_table.values and out["moved"]["to"] == "C1", out


def test_delete_with_values_needs_overwrite_and_hides_values(xl_table, capsys):
    code, out = run_main(excel, capsys, "range", "--action", "delete", "--range", "2:3")
    assert code == 1 and out["error_type"] == "needs_user_decision" and "--overwrite" in out["agent_directive"], out
    assert "6개" in out["message"] and "A2" in out["message"], out
    for v in ("영업", "300", "개발"):           # 값은 보이지 않음 (stdout, 오류 로그 모두)
        assert v not in out["message"] and v not in _tool_log_text(), v
    assert _col(xl_table, 1) == ["이름", "가", "나", "다", "라"]                       # 바꾸지 않음
    assert off.list_snapshots("excel") == []                                     # 되돌리기 지점도 남기지 않음


def test_delete_dry_run_then_overwrite(xl_table, capsys):
    code, out = run_main(excel, capsys, "range", "--action", "delete", "--range", "B:B", "--dry-run")
    assert code == 0 and out["dry_run"] is True and out["filled_cells"] == 5, out
    assert out["filled_addresses"] == ["B1", "B2", "B3", "B4", "B5"] and "영업" not in json.dumps(out, ensure_ascii=False), out
    assert "--overwrite" in out["next_commands"][0]["command"] and "--dry-run" not in out["next_commands"][0]["command"], out
    assert xl_table.values[(1, 2)] == "부서" and off.list_snapshots("excel") == []
    code, out = run_main(excel, capsys, "range", "--action", "delete", "--range", "B:B", "--overwrite")
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    assert xl_table.values[(1, 2)] == "금액" and out["moved"] == {"from": "C1", "to": "B1"}, out
    assert len(off.list_snapshots("excel")) == 1


def test_delete_empty_rows_needs_no_overwrite(xl_table, capsys):
    code, out = run_main(excel, capsys, "range", "--action", "delete", "--range", "8:9")
    assert code == 0 and out["filled_cells"] == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    assert _col(xl_table, 1) == ["이름", "가", "나", "다", "라"]


def test_cell_range_needs_axis(xl_table, capsys):
    code, out = run_main(excel, capsys, "range", "--action", "insert", "--range", "A2:C3")
    assert code == 1 and "[후보]" in out["message"] and "--axis rows" in out["message"], out
    code, out = run_main(excel, capsys, "range", "--action", "insert", "--range", "A2:C3", "--axis", "rows")
    assert code == 0 and _col(xl_table, 1, 1, 4) == ["이름", None, None, "가"], out


def test_dry_run_only_for_insert_delete(xl_table, capsys):
    code, out = run_main(excel, capsys, "range", "--action", "clear", "--range", "A1", "--dry-run")
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:"), out


# ---------------------------------------------------------------------------
# range sort
# ---------------------------------------------------------------------------


def test_sort_by_header_name_desc(xl_table, capsys):
    code, out = run_main(excel, capsys, "range", "--action", "sort", "--range", "A1:C5", "--by", "금액", "--desc")
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    assert _col(xl_table, 3) == ["금액", 1000, 300, 200, 100] and _col(xl_table, 1) == ["이름", "라", "가", "다", "나"]
    assert out["header"] is True and out["rows"] == 4 and out["keys"][0]["order"] == "desc", out
    assert len(off.list_snapshots("excel")) == 1


def test_sort_multi_key_desc_one(xl_table, capsys):
    code, out = run_main(excel, capsys, "range", "--action", "sort", "--range", "A1:C5", "--by", "부서,금액", "--desc", "금액")
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    assert _col(xl_table, 2) == ["부서", "개발", "기획", "영업", "영업"] and _col(xl_table, 3) == ["금액", 100, 1000, 300, 200], out


def test_sort_letter_auto_header(xl_table, capsys):
    code, out = run_main(excel, capsys, "range", "--action", "sort", "--range", "A1:C5", "--by", "C")
    assert code == 0 and out["header"] is True and "자동" in out["header_decided_by"], out
    assert _col(xl_table, 3) == ["금액", 100, 200, 300, 1000]


def test_sort_letter_header_no(xl_table, capsys):
    code, out = run_main(excel, capsys, "range", "--action", "sort", "--range", "A2:C5", "--by", "C", "--header", "no")
    assert code == 0 and out["header"] is False, out
    assert _col(xl_table, 3) == ["금액", 100, 200, 300, 1000]


def test_sort_unknown_column_lists_headers(xl_table, capsys):
    code, out = run_main(excel, capsys, "range", "--action", "sort", "--range", "A1:C5", "--by", "매출")
    assert code == 1 and "[후보] 이름, 부서, 금액" in out["message"], out
    assert "이름, 부서" not in _tool_log_text()                                         # 머리글은 문서 내용
    assert off.list_snapshots("excel") == []


def test_sort_refuses_partial_table_and_merged(xl_table, capsys):
    code, out = run_main(excel, capsys, "range", "--action", "sort", "--range", "A1:B5", "--by", "부서")
    assert code == 1 and out["error_type"] == "needs_user_decision" and "A1:C5" in out["message"], out
    assert _col(xl_table, 2) == ["부서", "영업", "개발", "영업", "기획"]
    xl_table.merged.add("$A$1:$C$5")
    code, out = run_main(excel, capsys, "range", "--action", "sort", "--range", "A1:C5", "--by", "부서")
    assert code == 1 and "병합" in out["message"], out


def test_sort_already_sorted_is_unchanged(xl_table, capsys):
    code, out = run_main(excel, capsys, "range", "--action", "sort", "--range", "A1:C5", "--by", "이름")
    assert code == 0 and out.get("unchanged") is True and off.list_snapshots("excel") == [], out
    assert not [x for x in xl_table.log if x[0] == "Sort"]


def test_sort_verify_catches_wrong_order(xl_table, capsys, monkeypatch):
    import fake_excel
    monkeypatch.setattr(fake_excel.FakeRange, "Sort", lambda self, **kw: None)      # Excel이 정렬하지 않음
    code, out = run_main(excel, capsys, "range", "--action", "sort", "--range", "A1:C5", "--by", "금액")
    assert code == 0 and out["agent_directive"].startswith("UNVERIFIED:"), out


# ---------------------------------------------------------------------------
# --new-sheet가 없는 명령의 없는 시트는 '찾을 수 없음' ([후보]/[추천]), --new-sheet 안내 없음
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("argv", [
    ["style", "--sheet", "요약표", "--range", "A1:B2", "--bold"],
    ["replace", "--sheet", "요약표", "--find", "a", "--replace", "b"],
    ["chart", "--sheet", "요약표", "--range", "A1:B3"],
    ["range", "--action", "width", "--sheet", "요약표", "--range", "A1", "--size", "20"],
])
def test_missing_sheet_without_new_sheet_option(xl, capsys, argv):
    code, out = run_main(excel, capsys, *argv)
    assert code == 1 and out["error_type"] != "needs_user_decision", out
    assert "--new-sheet" not in out["agent_directive"] and "--new-sheet" not in out["message"]
    assert "[후보] 요약" in out["message"]


def test_excel_write_still_asks_about_new_sheet(xl, capsys):
    code, out = run_main(excel, capsys, "write", "--sheet", "요약표", "--range", "A1", "--values", '[["a"]]')
    assert code == 1 and out["error_type"] == "needs_user_decision" and "--new-sheet" in out["agent_directive"]


# ---------------------------------------------------------------------------
# 잘못된 시트 이름은 시트를 만들기 전에 막음 ([추천] 고친 이름, 빈 시트가 남지 않음)
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("argv, rec", [
    (["write", "--new-sheet", "--sheet", "1/2분기", "--range", "A1", "--value", "1"], "1_2분기"),
    (["sheet", "--action", "add", "--name", "1/2분기"], "1_2분기"),
    (["sheet", "--action", "add", "--name", "'따옴표'"], "_따옴표_"),
    (["sheet", "--action", "rename", "--name", "요약", "--new-name", "a:b?"], "a_b_"),
    (["sheet", "--action", "copy", "--name", "Sheet1", "--new-name", "x" * 40], "x" * 31),
])
def test_invalid_sheet_name_is_refused_before_any_change(xl, capsys, argv, rec):
    _, wb = xl
    code, out = run_main(excel, capsys, *argv)
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:"), out
    assert f"[추천] {rec}" in out["message"], out
    assert sheet_names(wb) == ["Sheet1", "요약"] and undo_points() == []


def test_recommended_name_avoids_existing_sheet(xl, capsys):
    code, out = run_main(excel, capsys, "sheet", "--action", "add", "--name", "Sheet[1]")
    assert code == 1 and "[추천] Sheet_1_" in out["message"], out
    _, wb = xl
    wb.Worksheets.Add().Name = "Sheet_1_"
    code, out = run_main(excel, capsys, "sheet", "--action", "add", "--name", "Sheet[1]")
    assert code == 1 and "[추천] Sheet_1_ 2" in out["message"], out


def test_default_copy_name_of_long_sheet_fits_31_chars(xl, capsys):
    _, wb = xl
    wb.Worksheets(1).Name = "가" * 31
    code, out = run_main(excel, capsys, "sheet", "--action", "copy", "--name", "가" * 31)
    assert code == 0, out
    assert len(sheet_names(wb)[-1]) <= 31


def test_target_sheet_create_validates_name(xl):
    _, wb = xl
    with pytest.raises(rt.InputError):
        with off.open_target("excel", "active", readonly=False, tool="excel", command="write") as t:
            t.sheet("a*b", create=True)
    assert sheet_names(wb) == ["Sheet1", "요약"] and undo_points() == []


# ---------------------------------------------------------------------------
# 엑셀이 예약한 시트 이름 History(대소문자 무관)는 시트를 만들기 전에 막음
# ---------------------------------------------------------------------------


def test_fake_excel_refuses_reserved_name_like_real_excel(xl_ab):
    _, wb = xl_ab
    with pytest.raises(SheetNameRefused):
        wb.Worksheets.Add().Name = "history"


@pytest.mark.parametrize("argv", [
    ["sheet", "--action", "add", "--name", "History"],
    ["sheet", "--action", "add", "--name", "HISTORY"],
    ["write", "--new-sheet", "--sheet", "history", "--range", "A1", "--value", "1"],
    ["sheet", "--action", "rename", "--name", "a_b", "--new-name", "History"],
    ["sheet", "--action", "copy", "--name", "Sheet1", "--new-name", "History"],
])
def test_reserved_history_refused_before_any_change(xl_ab, capsys, argv):
    _, wb = xl_ab
    code, out = run_main(excel, capsys, *argv)
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:"), out
    assert "예약" in out["message"] and "[추천] " in out["message"], out
    assert out["message"].split("[추천] ")[1].strip().lower() == "history_", out
    assert sheet_names(wb) == ["Sheet1", "a_b"] and undo_points() == []        # 빈 시트·복사본이 남지 않음


def test_reserved_recommendation_avoids_existing_and_works(xl_ab, capsys):
    _, wb = xl_ab
    wb.Worksheets.Add().Name = "History_"
    code, out = run_main(excel, capsys, "sheet", "--action", "add", "--name", "History")
    assert code == 1 and "[추천] History_ 2" in out["message"], out
    code, out = run_main(excel, capsys, "sheet", "--action", "add", "--name", "History_ 2")    # [추천]대로 1회 재시도
    assert code == 0 and sheet_names(wb)[-1] == "History_ 2", out


def test_reserved_name_through_target_sheet_create(xl_ab):
    _, wb = xl_ab
    with pytest.raises(rt.InputError):
        with off.open_target("excel", "active", readonly=False, tool="excel", command="write") as t:
            t.sheet("History", create=True)
    assert sheet_names(wb) == ["Sheet1", "a_b"] and undo_points() == []


def test_names_containing_history_are_allowed():
    off.check_sheet_name("History 2026", [])
    off.check_sheet_name("작업History", [])


# ---------------------------------------------------------------------------
# chart
# ---------------------------------------------------------------------------


def test_chart_numeric_x_becomes_scatter_and_offers_transfer(xl, capsys):
    app, wb = xl
    ws = ws_of(wb)
    ws.set(1, 1, "x")
    ws.set(1, 2, "sin(x)")
    for r in range(2, 12):
        ws.set(r, 1, float(r))
        ws.set(r, 2, 0.5)
    code, out = run_excel(capsys, "chart", "--range", "A1:B11", "--type", "line", "--title", "사인 2주기")
    assert code == 0 and out["chart_type"] == "scatter_lines" and out["chart_name"] == "사인 2주기"
    assert out["agent_directive"].startswith("VERIFIED:")                 # 차트 목록·제목을 다시 읽음 (R1)
    cmd = out["next_commands"][0]["command"]
    assert "transfer" in cmd and '"사인 2주기"' in cmd
    code, out = run_excel(capsys, "chart", "--range", "A1:B11", "--title", "사인 2주기")
    assert out["chart_name"] == "사인 2주기 2"                    # 이름 중복 방지
    code, out = run_excel(capsys, "chart", "--action", "list")
    assert [c["name"] for c in out["charts"]] == ["사인 2주기", "사인 2주기 2"]


@pytest.mark.parametrize("kind,expected", [("scatter", -4169), ("scatter_lines", 75), ("line", 75), ("column", 51)])
def test_chart_type_survives_set_source_data(xl, capsys, kind, expected):
    # 실기 I5: 종류를 SetSourceData 앞에서만 정하면 Excel 2007이 분산형을 74(선+표식)로 바꾸고,
    # 뒤에서만 정하면 첫 열(x)이 계열로 그려짐. 앞뒤로 모두 정해야 함. 'scatter' 상수도 74로 잘못돼 있었음
    app, wb = xl
    ws = ws_of(wb)
    ws.set(1, 1, "x")
    ws.set(1, 2, "y")
    for r in range(2, 6):
        ws.set(r, 1, float(r))
        ws.set(r, 2, float(r * r))
    code, out = run_excel(capsys, "chart", "--range", "A1:B5", "--type", kind, "--title", "c")
    assert code == 0, out
    chart = ws.charts[-1]
    assert chart.type_at_source == expected and chart.Chart.ChartType == expected


def test_info_detail_reports_chart_type_series_and_covered_data(xl, capsys):
    app, wb = xl
    ws = ws_of(wb)
    ws.set(1, 1, "x")
    ws.set(1, 2, "y")
    for r in range(2, 12):
        ws.set(r, 1, float(r))
        ws.set(r, 2, float(r))
    run_excel(capsys, "chart", "--range", "A1:B11", "--type", "scatter", "--title", "옆 차트")            # 데이터 오른쪽
    run_excel(capsys, "chart", "--range", "A1:B11", "--type", "column", "--title", "덮는 차트", "--at", "A3")  # 데이터 위
    code, out = run_excel(capsys, "info", "--detail")
    assert code == 0
    sheet = [s for s in out["sheets"] if s["charts"]][0]
    charts = {c["name"]: c for c in sheet["charts"]}
    assert charts["옆 차트"]["chart_type"] == "scatter" and charts["옆 차트"]["series"] == 1
    assert not charts["옆 차트"].get("covers_data") and charts["덮는 차트"]["covers_data"] is True
    assert out["layout_warnings"] == 1 and "덮는 차트" in sheet["warnings"][0] and "배치 경고 1건" in out["preview"]
    code, plain = run_excel(capsys, "info")
    assert "layout_warnings" not in plain and "chart_type" not in plain["sheets"][0].get("charts", [{}])[0]


def test_chart_default_position_does_not_cover_other_data(xl, capsys):
    # 실기 J2: A1:C3 차트를 바로 오른쪽(D열)에 두어 F1:G5의 다른 데이터를 가렸음
    app, wb = xl
    ws = ws_of(wb)
    for r in range(1, 4):
        for c in range(1, 4):
            ws.set(r, c, "머리" if r == 1 else float(r * c))
    for r in range(1, 6):
        ws.set(r, 6, float(r))
        ws.set(r, 7, float(r * 2))
    code, out = run_excel(capsys, "chart", "--range", "A1:C3", "--type", "column", "--title", "막대")
    assert code == 0
    chart = ws.charts[-1]
    assert chart.Left >= 7 * 48.0 and chart.Top == 0           # G열(7) 오른쪽, 범위의 첫 행
    code, out = run_excel(capsys, "info", "--detail")
    assert out["layout_warnings"] == 0


def test_chart_delete_rules(xl, capsys):
    app, wb = xl
    ws = ws_of(wb)
    ws.set(1, 1, "a")
    ws.set(2, 1, 1.0)
    run_excel(capsys, "chart", "--range", "A1:A2", "--title", "매출")
    code, out = run_excel(capsys, "chart", "--action", "delete")
    assert code == 1
    code, out = run_excel(capsys, "chart", "--action", "delete", "--name", "매출차트")
    assert code == 1 and "[후보] 매출" in out["message"] and "[추천]" not in out["message"]
    code, out = run_excel(capsys, "chart", "--action", "delete", "--name", "매출")
    assert code == 0 and ws.charts == []


def test_argument_error_is_json(capsys):
    code, out = run_excel(capsys, "chart", "--type", "donut")
    assert code == 2 and out["tool"] == "excel"


# ---------------------------------------------------------------------------
# 제목이 같은 차트가 여러 시트에 있으면 시트로 가림 (next_commands에 --sheet, 이름이 제목보다 먼저)
# ---------------------------------------------------------------------------


def test_same_title_on_two_sheets_next_command_targets_new_chart(xl, capsys):
    app, wb = xl
    _fill(wb)
    run_main(excel, capsys, "chart", "--sheet", "Sheet1", "--range", "A1:B3", "--title", "매출")
    code, out = run_main(excel, capsys, "chart", "--sheet", "요약", "--range", "A1:B3", "--title", "매출")
    assert code == 0
    argv = shlex.split(out["next_commands"][0]["command"])
    assert argv[argv.index("--sheet") + 1] == "요약"
    ws, obj = excel.find_chart(_T(wb), argv[argv.index("--chart") + 1], argv[argv.index("--sheet") + 1])
    assert ws.Name == "요약"


def test_same_name_without_sheet_is_ambiguous(xl, capsys):
    app, wb = xl
    _fill(wb)
    run_main(excel, capsys, "chart", "--sheet", "Sheet1", "--range", "A1:B3", "--title", "매출")
    run_main(excel, capsys, "chart", "--sheet", "요약", "--range", "A1:B3", "--title", "매출")
    with pytest.raises(Exception) as ei:
        excel.find_chart(_T(wb), "매출", None)
    assert "[후보]" in str(ei.value) and "Sheet1!매출" in str(ei.value) and "요약!매출" in str(ei.value)
    ws, _ = excel.find_chart(_T(wb), "요약!매출", None)             # 후보 형식을 그대로 받음
    assert ws.Name == "요약"


def test_delete_by_name_does_not_pick_first_sheet(xl, capsys):
    app, wb = xl
    _fill(wb)
    run_main(excel, capsys, "chart", "--sheet", "Sheet1", "--range", "A1:B3", "--title", "매출")
    run_main(excel, capsys, "chart", "--sheet", "요약", "--range", "A1:B3", "--title", "매출")
    code, out = run_main(excel, capsys, "chart", "--action", "delete", "--name", "매출")
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:") and "[후보]" in out["message"]
    assert _names(wb) == {"Sheet1": ["매출"], "요약": ["매출"]}
    code, out = run_main(excel, capsys, "chart", "--action", "delete", "--name", "요약!매출")
    assert code == 0 and _names(wb) == {"Sheet1": ["매출"], "요약": []}, out


def test_exact_name_beats_other_chart_title(xl, capsys):
    app, wb = xl
    _fill(wb)
    # 앞 시트 차트의 제목이 뒤 시트 차트의 이름과 같아도 이름이 같은 차트를 고름 (이전: 앞 시트의 제목 차트)
    run_main(excel, capsys, "chart", "--sheet", "Sheet1", "--range", "A1:B3", "--title", "매출", "--name", "첫차트")
    run_main(excel, capsys, "chart", "--sheet", "요약", "--range", "A1:B3", "--title", "합계", "--name", "매출")
    ws, obj = excel.find_chart(_T(wb), "매출", None)
    assert ws.Name == "요약" and obj.Name == "매출"
    code, out = run_main(excel, capsys, "chart", "--action", "delete", "--name", "매출")
    assert code == 0 and _names(wb) == {"Sheet1": ["첫차트"], "요약": []}, out


# ---------------------------------------------------------------------------
# '시트!이름'을 --sheet와 함께 줘도, 차트 이름에 '!'가 있어도 찾음
# ---------------------------------------------------------------------------


def test_candidate_with_sheet_option_resolves(xl, capsys):
    app, wb = xl
    _fill(wb)
    run_main(excel, capsys, "chart", "--sheet", "Sheet1", "--range", "A1:B3", "--title", "매출", "--name", "Chart 1")
    run_main(excel, capsys, "chart", "--sheet", "Sheet1", "--range", "A1:B3", "--title", "매출", "--name", "Chart 2")
    code, out = run_main(excel, capsys, "chart", "--action", "delete", "--sheet", "Sheet1", "--name", "매출")
    assert code == 1 and "Sheet1!Chart 1" in out["message"], out
    code, out = run_main(excel, capsys, "chart", "--action", "delete", "--sheet", "Sheet1", "--name", "Sheet1!Chart 1")
    assert code == 0 and _names(wb)["Sheet1"] == ["Chart 2"], out


def test_bang_in_chart_name_round_trips(xl, capsys):
    app, wb = xl
    _fill(wb)
    run_main(excel, capsys, "chart", "--sheet", "Sheet1", "--range", "A1:B3", "--title", "매출 증가!")
    run_main(excel, capsys, "chart", "--sheet", "요약", "--range", "A1:B3", "--title", "매출 증가!")
    ws, obj = excel.find_chart(_T(wb), "요약!매출 증가!", None)
    assert ws.Name == "요약" and obj.Name == "매출 증가!"
    ws, obj = excel.find_chart(_T(wb), "Sheet1!매출 증가!", "Sheet1")
    assert ws.Name == "Sheet1"


# ---------------------------------------------------------------------------
# '시트!이름' 해석
# ---------------------------------------------------------------------------


def test_quoted_sheet_with_bang(xl, capsys):
    app, wb = xl
    _fill(wb)
    wb.sheets[0].Name = "A"
    wb.sheets[1].Name = "A!B"
    for s in ("A", "A!B"):
        run_main(excel, capsys, "chart", "--sheet", s, "--range", "A1:B3", "--title", "차트")
    code, out = run_main(excel, capsys, "chart", "--action", "delete", "--name", "차트")
    assert code == 1 and "'A!B'!차트" in out["message"] and "'A'!차트" in out["message"], out
    code, out = run_main(excel, capsys, "chart", "--action", "delete", "--name", "'A!B'!차트")
    assert code == 0 and _charts(wb) == {"A": [("차트", "차트")], "A!B": []}, out
    ws, obj = excel.find_chart(_T(wb), "'A'!차트")
    assert ws.Name == "A"


def test_unquoted_split_falls_back_to_later_bang(xl, capsys):
    app, wb = xl
    _fill(wb)
    wb.sheets[0].Name = "A"
    wb.sheets[1].Name = "A!B"
    run_main(excel, capsys, "chart", "--sheet", "A!B", "--range", "A1:B3", "--title", "차트")
    ws, obj = excel.find_chart(_T(wb), "A!B!차트")         # 'A'에서 '!B!차트'는 없음 -> 'A!B'에서 '차트'
    assert ws.Name == "A!B"


def test_literal_bang_chart_name_on_other_sheet(xl, capsys):
    app, wb = xl
    _fill(wb)
    run_main(excel, capsys, "chart", "--sheet", "요약", "--range", "A1:B3", "--title", "t1", "--name", "Sheet1!x")
    ws, obj = excel.find_chart(_T(wb), "Sheet1!x")         # Sheet1에 차트가 없어도 글자 그대로의 이름으로 찾음
    assert ws.Name == "요약" and obj.Name == "Sheet1!x"
    run_main(excel, capsys, "chart", "--sheet", "Sheet1", "--range", "A1:B3", "--title", "t2", "--name", "x")
    code, out = run_main(excel, capsys, "chart", "--action", "delete", "--name", "Sheet1!x")
    assert code == 1 and "[후보]" in out["message"], out          # 둘 다 맞으면 고르지 않음
    assert _charts(wb) == {"Sheet1": [("x", "t2")], "요약": [("Sheet1!x", "t1")]}
    cands = out["message"].split("[후보]")[1].strip().split(", ")
    assert len(cands) == 2
    for cand in cands:                                       # 후보는 그대로 주면 그 차트 하나로 찾아짐
        excel.find_chart(_T(wb), cand)
    code, out = run_main(excel, capsys, "chart", "--action", "delete", "--name", "'Sheet1'!x")
    assert code == 0 and _charts(wb) == {"Sheet1": [], "요약": [("Sheet1!x", "t1")]}, out


def test_sheet_option_contradicts_prefix(xl, capsys):
    app, wb = xl
    _fill(wb)
    run_main(excel, capsys, "chart", "--sheet", "Sheet1", "--range", "A1:B3", "--title", "매출")
    code, out = run_main(excel, capsys, "chart", "--action", "delete", "--sheet", "요약", "--name", "Sheet1!매출")
    assert code == 1 and "Sheet1" in out["message"] and "요약" in out["message"] and "먼저 만드세요" not in out["message"], out
    code, out = run_main(excel, capsys, "chart", "--action", "delete", "--sheet", "요약", "--name", "'Sheet1'!매출")
    assert code == 1 and "--sheet" in out["message"], out
    assert _charts(wb)["Sheet1"] == [("매출", "매출")]


# ---------------------------------------------------------------------------
# 한 시트의 같은 이름 차트
# ---------------------------------------------------------------------------


def test_duplicate_names_same_sheet(xl, capsys):
    app, wb = xl
    _fill(wb)
    run_main(excel, capsys, "chart", "--sheet", "Sheet1", "--range", "A1:B3", "--title", "매출")
    run_main(excel, capsys, "chart", "--sheet", "Sheet1", "--range", "A1:B3", "--title", "비용")
    excel._chart_objects(wb.sheets[0])[1].Name = "매출"
    code, out = run_main(excel, capsys, "chart", "--action", "delete", "--name", "매출")
    assert code == 1 and "Sheet1!#1" in out["message"] and "Sheet1!#2" in out["message"], out
    assert len(_charts(wb)["Sheet1"]) == 2
    code, out = run_main(excel, capsys, "chart", "--action", "delete", "--name", "Sheet1!#2")
    assert code == 0 and out["agent_directive"].startswith("VERIFIED"), out
    assert _charts(wb)["Sheet1"] == [("매출", "매출")]


# ---------------------------------------------------------------------------
# 차트 후보 이름
# ---------------------------------------------------------------------------


def test_chart_labels_count_all_charts_on_sheet(xl, capsys):
    app, wb = xl
    _fill(wb)
    run_main(excel, capsys, "chart", "--sheet", "Sheet1", "--range", "A1:B3", "--title", "Sales", "--name", "Chart 1")
    run_main(excel, capsys, "chart", "--sheet", "Sheet1", "--range", "A1:B3", "--title", "Other", "--name", "Chart 2")
    run_main(excel, capsys, "chart", "--sheet", "Sheet1", "--range", "A1:B3", "--title", "Sales", "--name", "Chart 3")
    excel._chart_objects(wb.sheets[0])[1].Name = "Chart 1"     # 후보가 아닌 2번 차트가 1번과 이름이 같음
    assert [n for n, _ in _charts(wb)["Sheet1"]] == ["Chart 1", "Chart 1", "Chart 3"]
    code, out = run_main(excel, capsys, "chart", "--action", "delete", "--name", "Sales")
    assert code == 1 and "[후보]" in out["message"], out
    cands = out["message"].split("[후보]")[1].strip().split(", ")
    assert cands == ["Sheet1!#1", "Sheet1!Chart 3"], cands
    found = [excel._chart_title(excel.find_chart(_T(wb), c)[1]) for c in cands]    # 후보마다 차트 하나로 찾아짐
    assert found == ["Sales", "Sales"]


# ---------------------------------------------------------------------------
# 0건 결과
# ---------------------------------------------------------------------------


def test_excel_chart_list_zero(xl, capsys):
    code, out = run_main(excel, capsys, "chart", "--action", "list")
    assert code == 0 and out["count"] == 0 and out["agent_directive"].startswith("VERIFIED: 0건"), out


# ---------------------------------------------------------------------------
# 빈 찾을 글은 모든 글자 사이에 바꿀 글을 끼워 넣음 -> 바꾸기 전에 STOP_RETRY
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("extra", [[], ["--all-sheets"]])
@pytest.mark.parametrize("argv", EMPTY_FIND)
def test_excel_empty_find_is_refused(xl_s2, capsys, argv, extra):
    app, wb = xl_s2
    ws, ws2 = wb.Worksheets(1), wb.Worksheets(2)
    ws.Range("A1").Value = "abc"
    ws.Range("B2").Value = "hello"
    ws2.Range("A1").Value = "xyz"
    code, out = run_main(excel, capsys, "replace", "--range", "A1:C3", *argv, *extra)
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:"), out
    assert "excel.py write" in out["message"], out
    assert ws.Range("A1").Value == "abc" and ws.Range("B2").Value == "hello" and ws2.Range("A1").Value == "xyz"
    assert undo_points("excel") == []


# ---------------------------------------------------------------------------
# 쓰기 전에 멈춘 실행은 되돌리기 지점을 남기지 않음
# ---------------------------------------------------------------------------


def test_refused_write_leaves_no_undo_point(xl, capsys):
    code, out = run_main(excel, capsys, "write", "--range", "A1", "--values", '[["a", 1]]')
    assert code == 0
    snaps = off.list_snapshots("excel")
    assert len(snaps) == 1
    first = snaps[0]["meta_file"]
    code, out = run_main(excel, capsys, "write", "--range", "A1", "--values", '[["b", 2]]')     # 덮어쓰기 확인
    assert code == 1 and out["error_type"] == "needs_user_decision"
    snaps = off.list_snapshots("excel")
    assert len(snaps) == 1 and snaps[0]["meta_file"] == first   # undo는 첫 쓰기 전 상태로 되돌림
    code, out = run_main(excel, capsys, "write", "--sheet", "없는시트", "--range", "A1", "--values", '[["c"]]')
    assert code == 1 and len(off.list_snapshots("excel")) == 1


# ---------------------------------------------------------------------------
# 이번 지점은 문서를 바꾸지 않았을 때만 지움
# ---------------------------------------------------------------------------


def test_point_is_kept_when_document_changed_before_input_error(xl):
    with pytest.raises(rt.InputError):
        with off.open_target("excel", "active", readonly=False, tool="excel", command="write") as t:
            t.sheet("중간", create=True)                    # 문서를 바꿈 (t.changed)
            raise rt.InputError("바꾼 뒤 멈춤")
    assert len(undo_points()) == 1


def test_dirty_document_keeps_point_unless_input_error(xl):
    _, wb = xl
    wb.Saved = False                                        # 원래 저장 안 된 문서: 저장 상태로는 바뀐 것을 알 수 없음
    with pytest.raises(RuntimeError):
        with off.open_target("excel", "active", readonly=False, tool="excel", command="write"):
            raise RuntimeError("COM 오류 (중간까지 바꿨는지 모름)")
    assert len(undo_points()) == 1
    with pytest.raises(rt.InputError):
        with off.open_target("excel", "active", readonly=False, tool="excel", command="write"):
            raise rt.InputError("바꾸기 전 입력 오류")
    assert len(undo_points()) == 1                              # 이번 실행의 지점만 지움


# ---------------------------------------------------------------------------
# 오래된 지점 정리는 명령이 바꾼 뒤에만 (거절된 실행이 진짜 옛 지점을 밀어내지 않음)
# ---------------------------------------------------------------------------


def test_refused_run_does_not_evict_old_point(xl, capsys):
    for i in range(off.SNAPSHOT_KEEP):
        code, out = run_main(excel, capsys, "write", "--range", f"A{i + 1}", "--value", f"v{i}")
        assert code == 0, out
    before = undo_points()
    assert len(before) == off.SNAPSHOT_KEEP
    code, out = run_main(excel, capsys, "write", "--sheet", "없는시트", "--range", "A1", "--value", "c")
    assert code == 1
    assert undo_points() == before
    code, out = run_main(excel, capsys, "write", "--range", "C1", "--value", "새 값")        # 성공한 실행은 정리
    assert code == 0 and len(undo_points()) == off.SNAPSHOT_KEEP and before[0] not in undo_points()


# ---------------------------------------------------------------------------
# 바꾼 것 없이 성공한 실행은 되돌리기 지점을 남기지 않음 (다음 undo가 바뀐 것 없는 지점으로 가지 않게)
# ---------------------------------------------------------------------------


def test_noop_replace_drops_its_point(xl, capsys):
    code, out = run_main(excel, capsys, "write", "--range", "A1", "--value", "초안")
    assert code == 0, out
    before = undo_points()
    assert len(before) == 1
    code, out = run_main(excel, capsys, "replace", "--find", "없는글", "--replace", "x")
    assert code == 0 and out["changed_cells"] == 0, out
    assert undo_points() == before                              # 다음 undo는 write를 되돌림
    code, out = run_main(office, capsys, "undo", "--list")
    assert code == 0 and len(out["snapshots"]) == 1 and out["snapshots"][0]["by"] == "excel write", out


def test_chart_delete_all_with_no_chart_drops_point(xl, capsys):
    code, out = run_main(excel, capsys, "chart", "--action", "delete", "--all")
    assert code == 0 and "0개" in out["preview"], out
    assert undo_points() == []


def test_sheet_add_existing_and_select_drop_point(xl, capsys):
    code, out = run_main(excel, capsys, "sheet", "--action", "add", "--name", "요약")
    assert code == 0 and "이미 있어" in out["preview"], out
    code, out = run_main(excel, capsys, "sheet", "--action", "select", "--name", "요약")
    assert code == 0, out
    assert undo_points() == []
    code, out = run_main(excel, capsys, "sheet", "--action", "add", "--name", "새시트")    # 실제로 바꾸면 지점이 남음
    assert code == 0 and len(undo_points()) == 1, out["message"]


def test_real_change_on_unsaved_document_keeps_point(xl, capsys):
    _, wb = xl
    wb.Saved = False                         # 저장 상태로는 바뀐 것을 알 수 없음: 시트 삭제가 t.changed를 표시해야 함
    code, out = run_main(excel, capsys, "sheet", "--action", "delete", "--name", "요약")
    assert code == 0 and sheet_names(wb) == ["Sheet1"], out
    assert len(undo_points()) == 1


# ---------------------------------------------------------------------------
# 바꾼 것 없이 성공한 명령은 바꾼 것으로 표시하지 않음 (되돌리기 지점을 남기지 않고 0곳 바뀜으로 보고)
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("argv", [
    ["replace", "--find", "초안", "--replace", "초안"],
    ["replace", "--map", '{"초안": "초안"}'],
])
def test_excel_replace_same_text_reports_zero(xl_ab, capsys, argv):
    _, wb = xl_ab
    wb.Worksheets(1).set(1, 1, "초안 보고")
    code, out = run_main(excel, capsys, *argv)
    assert code == 0 and out["changed_cells"] == 0 and out["unchanged_same_text"] == 1, out
    assert "0칸 바뀜" in out["preview"] and "처리" not in out["preview"], out
    assert undo_points() == []


def test_excel_replace_formula_cell_without_format_is_not_counted(xl_ab, capsys):
    _, wb = xl_ab
    ws = wb.Worksheets(1)
    ws.Range("A1").Formula = '="초안"'
    code, out = run_main(excel, capsys, "replace", "--find", "초안", "--replace", "최종")
    assert code == 0 and out["changed_cells"] == 0 and out["formula_cells_kept"], out
    assert undo_points() == []


@pytest.mark.parametrize("argv", [
    ["range", "--action", "unmerge", "--range", "A1:B2"],
    ["range", "--action", "unfreeze"],
    ["range", "--action", "clear", "--range", "A1:B2", "--type", "contents"],
    ["range", "--action", "move", "--range", "A5:B5", "--dest", "A5"],
])
def test_excel_range_noop_drops_point(xl_ab, capsys, argv):
    _, wb = xl_ab
    wb.Worksheets(1).set(5, 1, "값")                      # 사용 범위 A5:A5 (A1:B2 칸은 비어 있음)
    code, out = run_main(excel, capsys, *argv)
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    assert "바뀐 것 없음" in out["preview"] and out.get("unchanged") is True, out
    assert undo_points() == []
    assert wb.Worksheets(1).values[(5, 1)] == "값"


def test_excel_range_real_changes_keep_point(xl_ab, capsys):
    app, wb = xl_ab
    ws = wb.Worksheets(1)
    ws.set(1, 1, "값")
    ws.hidden["cols"].add(2)
    code, out = run_main(excel, capsys, "range", "--action", "unhide", "--range", "B:B")
    assert code == 0 and "바뀐 것 없음" not in out["preview"] and ws.hidden["cols"] == set(), out
    assert len(undo_points()) == 1
    ws.merged.add("$A$1:$B$2")
    code, out = run_main(excel, capsys, "range", "--action", "unmerge", "--range", "A1:B2")
    assert code == 0 and ws.merged == set() and len(undo_points()) == 2, out
    app.ActiveWindow.FreezePanes = True
    code, out = run_main(excel, capsys, "range", "--action", "unfreeze")
    assert code == 0 and app.ActiveWindow.FreezePanes is False and len(undo_points()) == 3, out
    code, out = run_main(excel, capsys, "range", "--action", "clear", "--range", "A1:B2")
    assert code == 0 and (1, 1) not in ws.values and len(undo_points()) == 4, out


# ---------------------------------------------------------------------------
# 닫힌 xlsx의 작은 범위 읽기는 시트 전체를 훑지 않음, Sheet!A1의 시트로 바꾼 뒤에 셈
# ---------------------------------------------------------------------------


def test_closed_bounds_only_when_needed(tmp_path, monkeypatch):
    import openpyxl
    p = tmp_path / "b.xlsx"
    wb = openpyxl.Workbook()
    wb.active.title = "첫"
    wb.active["A1"] = 1
    s2 = wb.create_sheet("둘")
    for r in range(1, 6):
        s2.cell(r, 2, r)
    wb.save(str(p))
    calls = []
    real = excel._closed_bounds

    def counting(ws):
        calls.append(ws.title)
        return real(ws)
    monkeypatch.setattr(excel, "_closed_bounds", counting)
    excel._read_matrix_closed(str(p), None, "A1:B2")
    assert calls == []
    sheet, r0, c0, rows, _ = excel._read_matrix_closed(str(p), None, "'둘'!B:B")
    assert calls == ["둘"] and sheet == "둘" and len(rows) == 5
    excel._read_matrix_closed(str(p), None, None)
    assert calls == ["둘", "첫"]


# ---------------------------------------------------------------------------
# 엑셀 범위 밖 칸 주소
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("ref", ["A0", "A1048577", "XFE1"])
def test_parse_cell_rejects_out_of_range(ref):
    with pytest.raises(rt.InputError):
        excel.parse_cell(ref)
