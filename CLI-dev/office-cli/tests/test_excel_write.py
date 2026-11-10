"""excel.py 쓰기: write, style, range(merge·copy·clear·insert·delete·sort), sheet, chart, replace와
그 실행이 남기는 되돌리기 지점.
"""
import json
import os
import re

import pytest

import _office as off
import _runtime as rt
import excel
import office
from fake_excel import FakeExcelWorkbook, FakeXlApp, SheetNameRefused
from fake_office import make_win32com
from office_helpers import option_value, recommended_argv, shell_sim, split_command, EMPTY_FIND, run_excel, run_main, set_active, sheet_names, stub_com, undo_points, ws_of


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
    # 차트 만들기는 문서를 바꿈: next_commands가 아니라 confirm_commands (1.4.75, 11차 검토)
    assert "next_commands" not in out and "chart" in out["confirm_commands"][0]["command"]
    assert "confirm_commands[0]" in out["agent_directive"] and "destructive" not in out["confirm_commands"][0]
    assert "--title" not in out["confirm_commands"][0]["command"]        # 자리표시 제목 '차트'를 그대로 쓰지 않게 (1.4.48)
    assert out["target_resolved"]["sheet"] == "Sheet1"


def test_write_formula_count_mismatch_is_input_error(xl, capsys):
    code, out = run_excel(capsys, "write", "--range", "A2:C10", "--formula", "=1", "=2")
    assert code == 1 and out["error_type"] == "input"


def test_write_checks_everything_before_writing_the_header(xl, capsys):
    # 회사 PC 오류 로그: 헤더를 먼저 쓰고 수식 검사에서 실패해 A1:B1만 남았음. 추천 범위는 머리글 행부터
    app, wb = xl
    code, out = run_excel(capsys, "write", "--range", "A1", "--header", "x,sin(x)", "--formula", "=(ROW()-2)*0.1", "=SIN(A2)")
    assert code == 1, out
    # 1.4.93 (검토 B4): 옆에 데이터가 없으면 행 수를 지어내지 않고(전에는 [추천] A1:B101) 사용자에게 마지막 행을 묻게 함
    assert "[추천]" not in out["message"] and "마지막 행" in out["message"] and "머리글 행부터" in out["message"], out
    assert (1, 1) not in ws_of(wb).values and (1, 2) not in ws_of(wb).values
    code, out = run_excel(capsys, "write", "--range", "A1:B101", "--header", "x,sin(x)", "--formula", "=(ROW()-2)*0.1", "=SIN(A2)")
    assert code == 0, out


def test_write_numbers_given_as_formulas_points_to_values(xl, capsys):
    code, out = run_excel(capsys, "write", "--range", "C6:C8", "--formula", "73.42", "49.97", "14.7")
    assert code == 1, out
    argv = recommended_argv(out["message"], "excel")             # 1.4.79: [추천]은 전체 명령
    assert option_value(argv, "--values") == "[[73.42], [49.97], [14.7]]" and "--formula" not in argv, argv
    assert "73.42" not in argv and option_value(argv, "--range") == "C6:C8", argv
    code, out = run_excel(capsys, *argv)
    assert code == 0, out


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
    assert "[후보] 중 무엇을 할지 물어보세요" in out["agent_directive"]
    assert "만들라고 답한 경우에만 같은 명령에 --new-sheet" in out["agent_directive"]   # 플래그는 사용자 동의 조건으로만
    # 1.4.83: 고를 것(새 시트 만들기)도 번호와 전체 명령. 확인 뒤에만 실행하는 confirm_commands (next_commands 아님)
    assert "[후보] 1) 새 시트 '결과표' 만들기" in out["message"], out
    assert out["confirm_commands"][-1]["command"].endswith("--new-sheet") and "next_commands" not in out, out
    _no_create_imperative(out)
    assert "[추천]" not in out["message"]
    assert "있는 시트: Sheet1, 요약" in out["message"] and "사용자에게 물어보세요" in out["message"]
    assert sheet_names(wb) == ["Sheet1", "요약"]                       # 새 시트 없음
    code, out = run_excel(capsys, "write", "--range", "'결과표'!A1", "--value", "1")
    assert code == 1 and sheet_names(wb) == ["Sheet1", "요약"]


def test_write_sheet_differs_only_in_spacing_recommends_existing(xl, capsys):
    app, wb = xl
    code, out = run_excel(capsys, "write", "--sheet", "Sheet 1", "--value", "1")
    assert code == 1 and out["error_type"] == "input", out
    argv = recommended_argv(out["message"], "excel")             # 1.4.79: [추천]은 전체 명령
    assert option_value(argv, "--sheet") == "Sheet1" and "Sheet 1" not in argv and option_value(argv, "--value") == "1", argv
    _no_create_imperative(out)
    assert sheet_names(wb) == ["Sheet1", "요약"]


def test_write_similar_but_different_sheet_asks_user(xl, capsys):
    # Sheet2를 Sheet1의 오타로 단정해 덮어쓰지 않음: [후보]는 실제 기존 시트만, 새 시트 여부는 사용자에게 물음
    app, wb = xl
    code, out = run_excel(capsys, "write", "--sheet", "Sheet2", "--value", "1")
    assert code == 1 and out["error_type"] == "needs_user_decision"
    assert out["message"].endswith("[후보] 1) 기존 시트 'Sheet1'에 쓰기 2) 새 시트 'Sheet2' 만들기"), out    # 1.4.83: 번호
    first, new = (split_command(c["command"])[2:] for c in out["confirm_commands"])
    assert first[first.index("--sheet") + 1] == "Sheet1" and "--new-sheet" not in first and new[-1] == "--new-sheet", out
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
    assert out.get("candidates", []) == ["기존 시트 'Sheet1'에 쓰기", "새 시트 'Sheet9' 만들기"] and "[추천]" not in out["message"]
    _no_create_imperative(out)
    assert "Sheet9" in out["message"] and "물어보세요" in out["agent_directive"]
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


def test_file_absolute_path_lookalike_recommends_full_command(xl, capsys, tmp_path, isolated_dirs):
    """1.4.79 회귀: 절대 경로 --file의 파일 이름만 표기가 다르면 같은 폴더의 파일로 고친 전체 명령을 [추천]
    (값 'Data_A.csv'만 주지 않고, 데이터 폴더의 다른 파일을 고르지도 않음)."""
    app, wb = xl
    d = tmp_path / "입력"
    d.mkdir()
    (d / "Data_A.csv").write_text("품목,수량\n볼트,3\n", encoding="utf-8")
    (isolated_dirs["data"] / "data_a.csv").write_text("다른,파일\n", encoding="utf-8")   # 데이터 폴더의 비슷한 이름은 후보 아님
    given = str(d / "data a.csv")
    code, out = run_excel(capsys, "write", "--file", given)
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:"), out
    argv = recommended_argv(out["message"], "excel")
    assert argv == ["write", "--file", str(d / "Data_A.csv")] and given not in argv, argv
    code, out = run_excel(capsys, *argv)                        # [추천]대로 다시 실행하면 그 파일을 씀
    assert code == 0, out
    assert [ws_of(wb).values[(1, 1)], ws_of(wb).values[(2, 2)]] == ["품목", 3]


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
    assert code == 1, out
    argv = recommended_argv(out["message"], "excel")             # 1.4.79: [추천]은 전체 명령
    assert option_value(argv, "--value") == "5" and "--formula" not in argv and option_value(argv, "--sheet") == "새시트", argv
    assert "--new-sheet" in argv
    assert "새시트" not in sheet_names(wb) and undo_points() == before
    code, out = run_main(excel, capsys, *argv)                   # 안내대로 다시 실행
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
    assert code == 1, out
    argv = recommended_argv(out["message"], "excel")             # 1.4.79: [추천]은 전체 명령
    assert argv == ["style", "--theme", "navy"], argv
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
    # 오류 로그의 머리 줄(===== 시각 | 도구 | Python | OS)은 빼고 봄: OS 이름에 '300'이 들어 있는 기계가 있음 (예: ...-v300-x86_64)
    log = "\n".join(ln for ln in _tool_log_text().splitlines() if not ln.startswith("====="))
    for v in ("영업", "300", "개발"):           # 값은 보이지 않음 (stdout, 오류 로그 모두)
        assert v not in out["message"] and v not in log, v
    assert _col(xl_table, 1) == ["이름", "가", "나", "다", "라"]                       # 바꾸지 않음
    assert off.list_snapshots("excel") == []                                     # 되돌리기 지점도 남기지 않음


def test_delete_dry_run_then_overwrite(xl_table, capsys):
    code, out = run_main(excel, capsys, "range", "--action", "delete", "--range", "B:B", "--dry-run")
    assert code == 0 and out["dry_run"] is True and out["filled_cells"] == 5, out
    assert out["filled_addresses"] == ["B1", "B2", "B3", "B4", "B5"] and "영업" not in json.dumps(out, ensure_ascii=False), out
    assert "--overwrite" in out["confirm_commands"][0]["command"] and "--dry-run" not in out["confirm_commands"][0]["command"], out
    assert out["confirm_commands"][0]["destructive"] is True  # 11차 검토 (1.4.75): 덮어쓰기·삭제·발송은 destructive (묻지 말고 해도 늘 물음)
    assert xl_table.values[(1, 2)] == "부서" and off.list_snapshots("excel") == []
    code, out = run_main(excel, capsys, "range", "--action", "delete", "--range", "B:B", "--overwrite")
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    assert xl_table.values[(1, 2)] == "금액" and out["moved"] == {"from": "C1", "to": "B1"}, out
    assert len(off.list_snapshots("excel")) == 1


def test_delete_empty_rows_needs_no_overwrite(xl_table, capsys):
    code, out = run_main(excel, capsys, "range", "--action", "delete", "--range", "8:9")
    assert code == 0 and out["filled_cells"] == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    assert _col(xl_table, 1) == ["이름", "가", "나", "다", "라"]


# 지울 칸을 참조하는 다른 칸의 수식 (#REF!가 됨, 1.4.41)


def _formula(ws, r, c, f, v=0):
    ws.formulas[(r, c)] = f
    ws.values[(r, c)] = v


def test_formula_refs_parses_a1_references():
    refs = excel.formula_refs("=SUM($B$2:B5)+'내 시트'!C3+요약!A:A+Sheet1!4:6+\"A9\"+LOG10(2)+[책.xlsx]Sheet1!A1")
    assert refs == [(None, 2, 2, 5, 2), ("내 시트", 3, 3, 3, 3), ("요약", 1, 1, 1048576, 1), ("Sheet1", 4, 1, 6, 16384)], refs


def test_delete_rows_referenced_by_formula_needs_decision(xl_table, capsys):
    _formula(xl_table, 7, 1, "=C3*2")              # 지울 3행만 참조 -> #REF!
    _formula(xl_table, 8, 1, "=SUM(C2:C5)")        # 범위 일부만 지움 -> 범위가 줄 뿐
    _formula(xl_table, 3, 4, "=C2")                # 지울 행 안의 수식은 함께 사라짐
    code, out = run_main(excel, capsys, "range", "--action", "delete", "--range", "3", "--overwrite", "--dry-run")
    assert code == 0 and out["broken_refs"] == 1 and out["broken_ref_cells"] == ["Sheet1!A7"], out
    assert "#REF!" in out["preview"] and "--overwrite" in out["confirm_commands"][0]["command"], out
    code, out = run_main(excel, capsys, "range", "--action", "delete", "--range", "3")
    assert code == 1 and out["error_type"] == "needs_user_decision", out
    assert "#REF!" in out["message"] and "Sheet1!A7" in out["message"] and "Sheet1!A8" not in out["message"], out
    assert "C3*2" not in out["message"] and "C3*2" not in _tool_log_text()          # 수식 글은 보이지 않음
    assert xl_table.values[(3, 1)] == "나"                                          # 바꾸지 않음
    code, out = run_main(excel, capsys, "range", "--action", "delete", "--range", "3", "--overwrite")
    assert code == 0 and out["broken_refs"] == 1 and "#REF!" in out["preview"], out


def test_delete_empty_rows_referenced_by_formula_needs_decision(xl_table, capsys):
    _formula(xl_table, 2, 4, "=A9")
    code, out = run_main(excel, capsys, "range", "--action", "delete", "--range", "8:9")
    assert code == 1 and out["error_type"] == "needs_user_decision" and "Sheet1!D2" in out["message"], out
    assert "값이 있는 칸" not in out["message"], out
    code, out = run_main(excel, capsys, "range", "--action", "delete", "--range", "8:9", "--dry-run")
    assert code == 0 and out["filled_cells"] == 0 and out["broken_refs"] == 1 and "--overwrite" in out["confirm_commands"][0]["command"]


def test_delete_cols_referenced_from_other_sheet(monkeypatch, capsys):
    app = FakeXlApp()
    wb = FakeExcelWorkbook(app, "보고.xlsx", sheets=("Sheet1", "요약"))
    stub_com(monkeypatch)
    set_active(monkeypatch, app, wb)
    ws, other = wb.Worksheets("Sheet1"), wb.Worksheets("요약")
    _formula(other, 1, 1, "=sheet1!B2")             # 시트 이름은 대소문자 무시
    _formula(other, 2, 1, "=B2")                   # 요약 시트 자기 칸
    _formula(other, 3, 1, "=SUM(Sheet1!A:C)")      # 열 범위 일부
    code, out = run_main(excel, capsys, "range", "--action", "delete", "--range", "B", "--dry-run")
    assert code == 0 and out["broken_ref_cells"] == ["요약!A1"], out


def test_delete_without_references_unchanged(xl_table, capsys):
    code, out = run_main(excel, capsys, "range", "--action", "delete", "--range", "8:9")
    assert code == 0 and out["broken_refs"] == 0 and out["broken_ref_cells"] == [], out


def test_cell_range_needs_axis(xl_table, capsys):
    code, out = run_main(excel, capsys, "range", "--action", "insert", "--range", "A2:C3")
    assert code == 1 and "[후보] 1) --axis rows" in out["message"], out
    assert [c["command"].split()[-2:] for c in out["confirm_commands"]] == [["--axis", "rows"], ["--axis", "cols"]], out
    code, out = run_main(excel, capsys, "range", "--action", "insert", "--range", "A2:C3", "--axis", "rows")
    assert code == 0 and _col(xl_table, 1, 1, 4) == ["이름", None, None, "가"], out


def test_dry_run_only_for_insert_delete_clear_move(xl_table, capsys):
    code, out = run_main(excel, capsys, "range", "--action", "merge", "--range", "A1:B1", "--dry-run")
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:"), out
    code, out = run_main(excel, capsys, "sheet", "--action", "rename", "--name", "Sheet1", "--new-name", "X", "--dry-run")
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
    assert code == 1 and "[후보] 1) 이름 2) 부서 3) 금액" in out["message"], out
    assert [c["command"].split("--by ")[-1] for c in out["confirm_commands"]] == ["이름", "부서", "금액"], out   # 고른 번호의 전체 명령
    assert "이름, 부서" not in _tool_log_text() and "1) 이름" not in _tool_log_text()                                        # 머리글은 문서 내용
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


def _wide_table(ws):
    rows = [("팀", "부서", "등급", "지역", "이름"),
            ("A", "영업", 2, "서울", "다"), ("A", "영업", 1, "부산", "나"), ("A", "영업", 1, "서울", "라"),
            ("B", "개발", 1, "서울", "가"), ("A", "영업", 1, "부산", "마")]
    for i, row in enumerate(rows, start=1):
        for j, v in enumerate(row, start=1):
            ws.values[(i, j)] = v


def test_sort_desc_not_in_by_gives_numbered_commands(xl_table, capsys):
    """--desc가 --by에 없음: [후보] 번호와 같은 번호의 전체 명령 (--desc 값만 바꿈)."""
    code, out = run_main(excel, capsys, "range", "--action", "sort", "--range", "A1:C5", "--by", "부서,금액", "--desc", "급액")
    assert code == 1 and "[후보] 1) 부서 2) 금액" in out["message"], out
    assert [c["command"].split("--desc ")[-1] for c in out["confirm_commands"]] == ["부서", "금액"], out


def test_sort_more_than_three_keys_uses_sort_fields(xl_table, capsys):
    """1.4.46: 기준 4개 이상은 Worksheet.Sort.SortFields (전에는 3개까지만, 4개면 STOP_RETRY)."""
    _wide_table(xl_table)
    code, out = run_main(excel, capsys, "range", "--action", "sort", "--range", "A1:E6", "--by", "팀,부서,등급,지역,이름",
                         "--desc", "지역")
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:") and len(out["keys"]) == 5, out
    assert _col(xl_table, 5, 1, 6) == ["이름", "라", "나", "마", "다", "가"], _col(xl_table, 5, 1, 6)
    applied = [x for x in xl_table.log if x[0] == "SortApply"]
    assert applied and applied[0][2] == excel.XL_YES and len(applied[0][3]) == 5, xl_table.log
    assert xl_table.Sort.SortFields.Count == 0                                     # 시트에 정렬 기준을 남기지 않음
    assert not [x for x in xl_table.log if x[0] == "Sort"]


def test_sort_too_many_keys(xl_table, capsys):
    code, out = run_main(excel, capsys, "range", "--action", "sort", "--range", "A1:C5", "--by", ",".join(["금액"] * 65))
    assert code == 1 and "64개" in out["message"], out


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
    assert "[후보] 1) 요약" in out["message"]


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
    assert sheet_names(wb) == ["Sheet1", "요약"] and undo_points() == []
    _rerun_recommended_name(capsys, wb, argv, out, rec)


def _name_option(argv):
    """시트 이름을 받는 옵션: write는 --sheet, rename·copy는 --new-name, add는 --name."""
    return "--sheet" if argv[0] == "write" else "--new-name" if "--new-name" in argv else "--name"


def _rerun_recommended_name(capsys, wb, argv, out, rec):
    """[추천] 전체 명령은 이름 옵션 값만 고친 이름으로 바꾸고, 그대로 다시 실행하면 그 이름의 시트가 생김 (1.4.79: [추천]은 전체 명령)."""
    opt = _name_option(argv)
    got = recommended_argv(out["message"], "excel")
    bad = argv[argv.index(opt) + 1]
    assert option_value(got, opt) == rec and bad not in got, got
    assert [a for a in got if a != rec] == [a for a in argv if a != bad], got      # 다른 인자는 그대로
    code, again = run_main(excel, capsys, *got)
    assert code == 0 and rec in sheet_names(wb), again


def test_recommended_name_avoids_existing_sheet(xl, capsys):
    code, out = run_main(excel, capsys, "sheet", "--action", "add", "--name", "Sheet[1]")
    assert code == 1, out
    assert option_value(recommended_argv(out["message"], "excel"), "--name") == "Sheet_1_", out     # 1.4.79: 전체 명령
    _, wb = xl
    wb.Worksheets.Add().Name = "Sheet_1_"
    code, out = run_main(excel, capsys, "sheet", "--action", "add", "--name", "Sheet[1]")
    assert code == 1, out
    _rerun_recommended_name(capsys, wb, ["sheet", "--action", "add", "--name", "Sheet[1]"], out, "Sheet_1_ 2")


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
    assert "예약" in out["message"], out
    assert sheet_names(wb) == ["Sheet1", "a_b"] and undo_points() == []        # 빈 시트·복사본이 남지 않음
    opt = _name_option(argv)                                                    # 1.4.79: [추천]은 전체 명령
    rec = option_value(recommended_argv(out["message"], "excel"), opt)
    assert rec.lower() == "history_", out
    _rerun_recommended_name(capsys, wb, argv, out, rec)


def test_reserved_recommendation_avoids_existing_and_works(xl_ab, capsys):
    _, wb = xl_ab
    wb.Worksheets.Add().Name = "History_"
    code, out = run_main(excel, capsys, "sheet", "--action", "add", "--name", "History")
    assert code == 1, out
    argv = recommended_argv(out["message"], "excel")                            # 1.4.79: [추천]은 전체 명령
    assert argv == ["sheet", "--action", "add", "--name", "History_ 2"], argv
    code, out = run_main(excel, capsys, *argv)                                 # [추천]대로 1회 재시도
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
    assert "next_commands" not in out                                     # 옮기기는 문서를 바꿈 (1.4.75)
    cmd = out["confirm_commands"][0]["command"]
    assert "transfer" in cmd and "'사인 2주기'" in cmd
    argv = split_command(cmd)[2:]                                         # 1.4.79: 앱은 --from-app/--to-app (--from/--to는 outlook 전용)
    assert argv[:3] == ["transfer", "--from-app", "excel"] and argv[argv.index("--chart") + 1] == "사인 2주기", argv
    assert argv[argv.index("--to-app") + 1] in ("word", "ppt", "word,ppt") and "--to" not in argv, argv
    code, out = run_excel(capsys, "chart", "--range", "A1:B11", "--title", "사인 2주기")
    assert out["chart_name"] == "사인 2주기 2"                    # 이름 중복 방지
    code, out = run_excel(capsys, "chart", "--action", "list")
    assert [c["name"] for c in out["charts"]] == ["사인 2주기", "사인 2주기 2"]


@pytest.mark.parametrize("kind,expected", [("scatter", -4169), ("scatter_lines", 75), ("line", 75), ("column", 51)])
def test_chart_type_survives_set_source_data(xl, capsys, kind, expected):
    # 실기 I5: 종류를 SetSourceData 앞에서만 정하면 Excel이 분산형을 74(선+표식)로 바꾸고,
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
    assert code == 1 and "[후보] 1) 매출" in out["message"] and out["confirm_commands"][0]["destructive"] is True and "[추천]" not in out["message"]
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
    argv = split_command(out["confirm_commands"][0]["command"])
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
    cands = re.split(r"\s*\d+\) ", out["message"].split("[후보]")[1])[1:]
    assert len(cands) == 2 and len(out["confirm_commands"]) == 2, out
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
    cands = re.split(r"\s*\d+\) ", out["message"].split("[후보]")[1])[1:]
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

    def counting(ws, path=None):
        calls.append(ws.title)
        return real(ws, path)
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


# ---------------------------------------------------------------------------
# Gemma 실행 기록(1.4.38): 칸 하나 쓰기에 워드식 --text, 표 파일 못 찾음
# ---------------------------------------------------------------------------

def test_text_option_on_excel_write_recommends_value(capsys):
    # 철자가 비슷한 --target이 아니라 같은 뜻의 --value를 추천 (그대로 따르면 바로 성공해야 함)
    code, out = run_excel(capsys, "write", "--range", "A2", "--text", "반가워")
    assert code == 2 and out["error_type"] == "input"
    assert "--value 반가워" in out["message"].split("[추천] ", 1)[1] and "--target" not in out["message"]


def test_value_writes_single_cell(capsys, xl):
    _, wb = xl
    code, out = run_excel(capsys, "write", "--range", "A2", "--value", "반가워")
    assert code == 0 and out["status"] == "ok", out
    assert ws_of(wb).Range("A2").Value == "반가워"


def test_missing_table_file_lists_only_table_files(capsys, xl, isolated_dirs):
    for name in ("excel_results.csv", "메모.md", "chart_export.png", "보고서.pdf"):
        (isolated_dirs["data"] / name).write_text("x", encoding="utf-8")
    code, out = run_excel(capsys, "write", "--range", "A2", "--file", "temp_val.csv")
    assert code == 1 and out["status"] == "error"
    msg = out["message"]
    assert "excel_results.csv" in msg and "메모.md" in msg
    assert "chart_export.png" not in msg and "보고서.pdf" not in msg


# ---------------------------------------------------------------------------
# comment: 칸 메모 (목록 / 달기 / 바꾸기·덧붙이기 / 지우기, 1.4.50)
# ---------------------------------------------------------------------------

from fake_excel import FakeComment  # noqa: E402
from office_helpers import error_log_text  # noqa: E402


def _note(ws, r, c, text):
    ws.notes[(r, c)] = FakeComment(ws, r, c, text)


def _notes(ws):
    return {excel.addr(r, c): n.text for (r, c), n in ws.notes.items()}


def test_comment_list_empty_and_range(xl, capsys):
    _, wb = xl
    ws = ws_of(wb)
    code, out = run_excel(capsys, "comment")
    assert code == 0 and out["count"] == 0 and out["agent_directive"].startswith("VERIFIED: 0건"), out
    _note(ws, 2, 2, "첫 메모")
    _note(ws, 9, 4, "멀리 있는 메모")
    code, out = run_excel(capsys, "comment")
    assert code == 0 and out["count"] == 2 and out["truncated"] is False
    assert [(i["cell"], i["text"]) for i in out["items"]] == [("B2", "첫 메모"), ("D9", "멀리 있는 메모")]
    assert out["items"][0]["author"] == "검토자" and out["sheet"] == "Sheet1"
    code, out = run_excel(capsys, "comment", "--range", "A1:C5")
    assert out["count"] == 1 and out["items"][0]["cell"] == "B2"


def test_comment_list_is_capped_with_next_command(xl, capsys, monkeypatch):
    _, wb = xl
    ws = ws_of(wb)
    for r in range(1, 31):
        _note(ws, r, 1, f"메모 {r} " + "가" * 50)
    monkeypatch.setattr(rt, "DEFAULT_MAX_CHARS", 600)
    code, out = run_excel(capsys, "comment")
    assert code == 0 and out["count"] == 30 and out["truncated"] is True and 0 < out["shown"] < 30
    nxt = out["next_commands"][0]["command"]
    assert f"--start {out['shown'] + 1}" in nxt and " comment " in nxt


def test_comment_add_one_verifies_and_leaves_undo_point(xl, capsys):
    _, wb = xl
    ws = ws_of(wb)
    code, out = run_excel(capsys, "comment", "--action", "add", "--range", "B5", "--text", "규격 3.1과 다름\\n확인 필요")
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    assert ws.notes[(5, 2)].text == "규격 3.1과 다름\n확인 필요"
    assert out["applied"] == ["메모 추가: Sheet1!B5"] and out["count"] == 1
    assert len(undo_points()) == 1
    # --action add를 빠뜨려도 --text가 있으면 달기 (목록을 보여 주고 '완료'라고 하지 않음)
    code, out = run_excel(capsys, "comment", "--range", "C5", "--text", "둘째")
    assert code == 0 and ws.notes[(6 - 1, 3)].text == "둘째"


def test_comment_add_from_file_with_sheet(xl, capsys, tmp_path):
    _, wb = xl
    f = tmp_path / "메모.csv"
    f.write_text("칸,메모\nB2,값이 규격보다 큼\n'요약'!C3,\"쉼표, 있는 메모\"\n", encoding="utf-8")
    code, out = run_excel(capsys, "comment", "--action", "add", "--file", str(f))
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    assert _notes(ws_of(wb)) == {"B2": "값이 규격보다 큼"}
    assert _notes(ws_of(wb, "요약")) == {"C3": "쉼표, 있는 메모"}
    assert out["count"] == 2 and len(undo_points()) == 1
    # --sheet: 시트 이름이 없는 칸은 그 시트에
    f2 = tmp_path / "메모2.csv"
    f2.write_text("D4,요약 시트 메모\n", encoding="utf-8")
    code, out = run_excel(capsys, "comment", "--file", str(f2), "--sheet", "요약")
    assert code == 0 and _notes(ws_of(wb, "요약"))["D4"] == "요약 시트 메모"


def test_comment_file_bad_rows_add_nothing(xl, capsys, tmp_path):
    _, wb = xl
    f = tmp_path / "메모.csv"
    f.write_text("칸,메모\nB2,좋음\n가나다,잘못된 칸\nC3,\n", encoding="utf-8")
    code, out = run_excel(capsys, "comment", "--action", "add", "--file", str(f))
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:")
    assert "3번째 줄" in out["message"] and "4번째 줄" in out["message"]
    assert ws_of(wb).notes == {} and undo_points() == []


def test_comment_existing_note_needs_user_decision(xl, capsys, monkeypatch, isolated_dirs):
    from office_helpers import run_fresh
    _, wb = xl
    ws = ws_of(wb)
    _note(ws, 5, 2, "ZQX 기존 메모")
    code, out = run_fresh(excel, capsys, monkeypatch, "comment", "--action", "add", "--range", "B5", "--text", "새 의견")
    assert code == 1 and out["error_type"] == "needs_user_decision", out
    assert "B5" in out["message"] and "ZQX" in out["message"]
    assert "--overwrite" in out["agent_directive"] and "--append" in out["agent_directive"]
    log = error_log_text(isolated_dirs)
    assert "B5" in log and "ZQX" not in log and "새 의견" not in log
    assert ws.notes[(5, 2)].text == "ZQX 기존 메모" and undo_points() == []
    code, out = run_excel(capsys, "comment", "--range", "B5", "--text", "덧붙임", "--append")
    assert code == 0 and ws.notes[(5, 2)].text == "ZQX 기존 메모\n덧붙임" and out["applied"] == ["메모 덧붙이기: Sheet1!B5"]
    code, out = run_excel(capsys, "comment", "--range", "B5", "--text", "바꾼 메모", "--overwrite")
    assert code == 0 and ws.notes[(5, 2)].text == "바꾼 메모" and out["agent_directive"].startswith("VERIFIED:")
    code, out = run_excel(capsys, "comment", "--range", "B5", "--text", "x", "--overwrite", "--append")
    assert code in (1, 2) and out["status"] == "error"


def test_comment_add_dry_run_changes_nothing(xl, capsys):
    _, wb = xl
    ws = ws_of(wb)
    _note(ws, 2, 2, "기존")
    code, out = run_excel(capsys, "comment", "--range", "B2:B3", "--text", "의견", "--dry-run")
    assert code == 0 and out["dry_run"] is True and out["agent_directive"].startswith("VERIFIED: 미리보기")
    assert ws.notes[(2, 2)].text == "기존" and (3, 2) not in ws.notes and undo_points() == []
    assert any("B2" in p and "이미" in p for p in out["planned"]) and any("B3" in p for p in out["planned"])
    cmds = [c["command"] for c in out["confirm_commands"]]
    assert all("--dry-run" not in c for c in cmds)
    assert any("--overwrite" in c for c in cmds) and any("--append" in c for c in cmds)
    flags = {("--overwrite" in c["command"]): c.get("destructive", False) for c in out["confirm_commands"]}
    assert flags == {True: True, False: False}  # 11차 검토 (1.4.75): 덮어쓰기·삭제·발송은 destructive (묻지 말고 해도 늘 물음)
    code, out = run_excel(capsys, "comment", "--range", "C9", "--text", "의견", "--dry-run")
    assert out["confirm_commands"][0]["command"].endswith('--text 의견') or "--dry-run" not in out["confirm_commands"][0]["command"]
    assert (9, 3) not in ws.notes


def test_comment_delete(xl, capsys):
    _, wb = xl
    ws = ws_of(wb)
    code, out = run_excel(capsys, "comment", "--action", "delete")
    assert code == 1 and "--range" in out["message"]
    _note(ws, 2, 2, "가")
    _note(ws, 3, 2, "나")
    _note(ws, 9, 9, "다")
    code, out = run_excel(capsys, "comment", "--action", "delete", "--range", "B1:B5", "--dry-run")
    assert code == 0 and out["dry_run"] is True and len(ws.notes) == 3 and undo_points() == []
    assert out["planned"] == ["메모 삭제: Sheet1!B2", "메모 삭제: Sheet1!B3"]
    assert "--dry-run" not in out["confirm_commands"][0]["command"] and out["confirm_commands"][0]["destructive"] is True
    code, out = run_excel(capsys, "comment", "--action", "delete", "--range", "B1:B5")
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:") and list(ws.notes) == [(9, 9)]
    assert out["applied"] == ["메모 삭제: Sheet1!B2", "메모 삭제: Sheet1!B3"] and len(undo_points()) == 1
    code, out = run_excel(capsys, "comment", "--action", "delete", "--range", "A1:A3")
    assert code == 0 and out["count"] == 0 and out["agent_directive"].startswith("VERIFIED: 0건") and len(undo_points()) == 1


def test_comment_refused_by_excel_is_input_error(xl, capsys):
    _, wb = xl
    ws = ws_of(wb)
    ws.refuse_notes = True
    code, out = run_excel(capsys, "comment", "--range", "B2", "--text", "의견")
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:") and "보호" in out["message"]


# ---------------------------------------------------------------------------
# 문서에 없던 서식 옵션 점검 (1.4.53): --wrap, --italic, range border --weight/--box,
# write --no-autofit, style --no-zebra, export --encoding
# ---------------------------------------------------------------------------


def test_style_wrap_and_italic_are_read_back(xl, capsys):
    app, wb = xl
    ws = ws_of(wb)
    ws.set(1, 1, "긴 글")
    code, out = run_excel(capsys, "style", "--range", "A1:B2", "--wrap", "--italic")
    assert code == 0 and out["verified"] is True, out
    assert ws.props("$A$1:$B$2", "Misc").get("WrapText") is True


def test_style_wrap_not_applied_is_unverified(xl, capsys, monkeypatch):
    app, wb = xl
    real = excel._apply_format

    def no_wrap(rng, spec):
        real(rng, {k: v for k, v in spec.items() if k != "wrap"})
    monkeypatch.setattr(excel, "_apply_format", no_wrap)
    code, out = run_excel(capsys, "style", "--range", "A1:B2", "--wrap")
    assert code == 0 and out["verified"] is False and any("wrap" in u for u in out["unverified"]), out


@pytest.mark.parametrize("box,edges,inner", [(False, (7, 8, 9, 10, 11, 12), ()), (True, (7, 8, 9, 10), (11, 12))])
def test_range_border_weight_and_box_are_read_back(xl, capsys, box, edges, inner):
    app, wb = xl
    ws = ws_of(wb)
    argv = ["range", "--action", "border", "--range", "A1:C3", "--weight", "medium"] + (["--box"] if box else [])
    code, out = run_excel(capsys, *argv)
    assert code == 0 and out["verified"] is True, out
    for b in edges:
        p = ws.props("$A$1:$C$3", f"Border{b}")
        assert p.get("LineStyle") == 1 and p.get("Weight") == excel.XL_WEIGHT["medium"], b
    for b in inner:
        assert ws.props("$A$1:$C$3", f"Border{b}").get("LineStyle") is None, b


def test_range_border_none_is_read_back(xl, capsys):
    app, wb = xl
    ws = ws_of(wb)
    code, out = run_excel(capsys, "range", "--action", "border", "--range", "A1:B2", "--weight", "none")
    assert code == 0 and out["verified"] is True, out
    assert ws.props("$A$1:$B$2", "Border7").get("LineStyle") == -4142


def test_range_border_failure_is_unverified(xl, capsys, monkeypatch):
    monkeypatch.setattr(excel, "_apply_border", lambda rng, weight, box=False: None)
    code, out = run_excel(capsys, "range", "--action", "border", "--range", "A1:B2", "--weight", "thick")
    assert code == 0 and out["verified"] is False and any("테두리" in u for u in out["unverified"]), out


def test_write_table_autofits_unless_no_autofit(xl, capsys, tmp_path):
    app, wb = xl
    ws = ws_of(wb)
    f = tmp_path / "t.csv"
    f.write_text("a,b\n1,2\n", encoding="utf-8")
    code, out = run_excel(capsys, "write", "--range", "A1", "--file", str(f))
    assert code == 0 and any(x[0] == "AutoFit" for x in ws.log), ws.log
    ws.log.clear()
    code, out = run_excel(capsys, "write", "--range", "D1", "--file", str(f), "--no-autofit")
    assert code == 0 and not any(x[0] == "AutoFit" for x in ws.log), ws.log


def test_style_theme_no_zebra(xl, capsys):
    app, wb = xl
    ws = ws_of(wb)
    rows = [["지역", "금액"]] + [[f"지역{i}", float(i)] for i in range(1, 8)]
    for r, row in enumerate(rows, 1):
        for c, v in enumerate(row, 1):
            ws.set(r, c, v)
    code, out = run_excel(capsys, "style", "--theme", "navy", "--no-zebra")
    assert code == 0 and out["applied"]["zebra_rows"] == 0 and out["verified"] is True, out


def test_export_csv_cp949_and_row_count(xl, capsys, tmp_path):
    app, wb = xl
    ws = ws_of(wb)
    for r, row in enumerate([["이름", "온도"], ["가", "25℃"], ["나", "30℃"]], 1):
        for c, v in enumerate(row, 1):
            ws.set(r, c, v)
    out_path = tmp_path / "out.csv"
    code, out = run_excel(capsys, "export", "--save", str(out_path), "--encoding", "cp949")
    assert code == 0 and out["verified"] is True, out
    assert out_path.read_bytes().decode("cp949").splitlines()[1] == "가,25℃"


def test_export_csv_cp949_unencodable_is_input_error(xl, capsys, tmp_path):
    app, wb = xl
    ws = ws_of(wb)
    ws.set(1, 1, "이름")
    ws.set(2, 1, "\U0001F600 값")                     # cp949로 쓸 수 없는 글자
    out_path = tmp_path / "out.csv"
    code, out = run_excel(capsys, "export", "--save", str(out_path), "--encoding", "cp949")
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:"), out
    assert "A2" in out["message"] and recommended_argv(out["message"], "excel")[-2:] == ["--encoding", "utf-8-sig"], out
    assert not out_path.exists()


# ---------------------------------------------------------------------------
# write --cells: 흩어진 칸 여러 개를 한 번에 (1.4.54). 회사 기록: C6..C16에 write 16번, 확인 8번
# ---------------------------------------------------------------------------


def test_write_cells_inline_many_sheets_one_undo_point(xl, capsys):
    _, wb = xl
    code, out = run_excel(capsys, "write", "--cells", "C6=ft; C7=℃; 요약!C11=20; C8==C6&C7")
    assert code == 0 and out["verified"] is True and out["agent_directive"].startswith("VERIFIED:"), out
    ws, ws2 = ws_of(wb), ws_of(wb, "요약")
    assert ws.values[(6, 3)] == "ft" and ws.values[(7, 3)] == "℃" and ws2.values[(11, 3)] == 20
    assert ws.formulas[(8, 3)] == "=C6&C7"                                # =로 시작하는 값은 수식
    assert out["count"] == 4 and len(undo_points()) == 1


def test_write_cells_file_with_header_and_quotes(xl, capsys, tmp_path):
    _, wb = xl
    f = tmp_path / "칸값.csv"
    f.write_text("칸,값\nC6,ft\n'요약'!B2,\"쉼표, 든 값\"\nC9,0012\n", encoding="utf-8")
    code, out = run_excel(capsys, "write", "--cells", str(f))
    assert code == 0 and out["verified"] is True, out
    assert ws_of(wb).values[(6, 3)] == "ft" and ws_of(wb, "요약").values[(2, 2)] == "쉼표, 든 값"
    assert ws_of(wb).values[(9, 3)] == "0012"                            # 앞자리 0은 글자 그대로
    code, out = run_excel(capsys, "write", "--cells", "C10='a;b'; C11=x")
    assert code == 0 and ws_of(wb).values[(10, 3)] == "a;b" and ws_of(wb).values[(11, 3)] == "x"


def test_write_cells_conflicts_one_decision_one_overwrite(xl, capsys):
    _, wb = xl
    ws = ws_of(wb)
    ws.set(6, 3, "old6")
    ws.set(7, 3, "old7")
    ws.set(8, 3, "same")
    cells = "C6=a; C7=b; C8=same; C9=c"
    code, out = run_excel(capsys, "write", "--cells", cells)
    assert code == 1 and out["error_type"] == "needs_user_decision", out
    assert "C6=old6" in out["message"] and "C7=old7" in out["message"] and "C8" not in out["message"]
    assert (9, 3) not in ws.values and undo_points() == []           # 하나도 쓰지 않음
    code, out = run_excel(capsys, "write", "--cells", cells, "--overwrite")
    assert code == 0 and out["verified"] is True, out
    assert [ws.values[(r, 3)] for r in (6, 7, 8, 9)] == ["a", "b", "same", "c"]
    assert out["count"] == 3 and out["unchanged"] == 1 and len(undo_points()) == 1


def test_write_cells_dry_run(xl, capsys):
    _, wb = xl
    ws = ws_of(wb)
    ws.set(6, 3, "old")
    code, out = run_excel(capsys, "write", "--cells", "C6=a; C7=b", "--dry-run")
    assert code == 0 and out["dry_run"] is True and out["agent_directive"].startswith("VERIFIED: 미리보기"), out
    assert ws.values[(6, 3)] == "old" and (7, 3) not in ws.values and undo_points() == []
    assert any("덮어쓰기" in p and "C6" in p for p in out["planned"]) and any("C7" in p for p in out["planned"])
    assert "next_commands" not in out, out              # 1.4.73: 실제로 쓰는 명령은 사용자 확인 뒤에만 (confirm_commands)
    cmd = out["confirm_commands"][0]["command"]
    assert "--dry-run" not in cmd and "--overwrite" in cmd and "--cells" in cmd
    assert out["confirm_commands"][0]["destructive"] is True  # 11차 검토 (1.4.75): 덮어쓰기·삭제·발송은 destructive (묻지 말고 해도 늘 물음)
    code, out = run_excel(capsys, "write", "--cells", "C8=a; C9=b", "--dry-run")
    assert code == 0 and "destructive" not in out["confirm_commands"][0], out           # 빈 칸만: 덮어쓰기 아님
    ws.set(8, 3, "a")
    ws.set(9, 3, "b")
    code, out = run_excel(capsys, "write", "--cells", "C8=a; C9=b", "--dry-run")    # 1.4.83: 모두 같은 값이면 실행할 명령 없음
    assert code == 0 and out["count"] == 0 and "confirm_commands" not in out and "바뀔 것 없음" in out["preview"], out
    assert out["agent_directive"] == rt.DRY_RUN_NOTHING, out          # 1.4.86: 미리보기 0건 문구는 하나


@pytest.mark.parametrize("shell", ["powershell", "pwsh", "cmd", "bash"])
def test_write_cells_confirm_command_runs_as_given_in_each_shell(xl, capsys, monkeypatch, shell):
    """1.4.76: confirm_commands의 따옴표는 실행 셸에 맞춤 (PowerShell은 작은따옴표, 5.1은 안의 "를 \\"로).
    $·백틱·{}·@·#·;·%·" 가 든 --cells를 그 셸이 실행하면(shell_sim) 같은 값이 그대로 도구에 들어간다."""
    _, wb = xl
    monkeypatch.setenv("OFFICE_CLI_SHELL", shell)
    ws_of(wb).set(6, 3, "old")
    cells = 'C6==$A$1*2; C7={{업체명}} @a #b; C8=100% `x; C9=\'a;b\'; C10=부서 = "영업"; C11=C:\\a b\\'
    code, out = run_excel(capsys, "write", "--cells", cells, "--dry-run")
    assert code == 0 and out["dry_run"] is True, out
    argv = shell_sim.run_shell(out["confirm_commands"][0]["command"], shell)
    assert argv[argv.index("--cells") + 1] == cells and "--overwrite" in argv, (shell, out["confirm_commands"])
    code, out = run_excel(capsys, *argv[2:])
    ws = ws_of(wb)
    assert code == 0 and ws.formulas[(6, 3)] == "=$A$1*2" and ws.values[(7, 3)] == "{{업체명}} @a #b", out
    assert ws.values[(8, 3)] == "100% `x" and ws.values[(9, 3)] == "a;b" and ws.values[(11, 3)] == "C:\\a b\\"


def test_write_cells_bad_items_write_nothing(xl, capsys, tmp_path):
    _, wb = xl
    code, out = run_excel(capsys, "write", "--cells", "C6=a; 가나=b; C8")
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:"), out
    assert "가나" in out["message"] and "C8" in out["message"]
    assert ws_of(wb).values == {} and undo_points() == []
    code, out = run_excel(capsys, "write", "--cells", "C6=a; C6=b")
    assert code == 1 and "C6" in out["message"] and ws_of(wb).values == {}
    code, out = run_excel(capsys, "write", "--cells", "C6=a", "--value", "x")
    assert code == 1 and "--cells" in out["message"]


def test_write_cells_missing_sheet_suggests(xl, capsys):
    _, wb = xl
    code, out = run_excel(capsys, "write", "--cells", "C6=a; 요 약!C7=b")
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:") and "요약" in out["message"], out
    assert ws_of(wb).values == {}


def test_write_file_with_cell_value_header_points_to_cells(xl, capsys, tmp_path):
    _, wb = xl
    f = tmp_path / "칸값.csv"
    f.write_text("칸,값\nC6,ft\n", encoding="utf-8")
    code, out = run_excel(capsys, "write", "--file", str(f))
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:"), out
    assert "[추천] python" in out["message"] and "--cells" in out["message"]
    assert ws_of(wb).values == {}


def test_write_cells_reads_back(xl, capsys, monkeypatch):
    _, wb = xl
    ws = ws_of(wb)
    real_set = ws.set
    monkeypatch.setattr(ws, "set", lambda r, c, v: real_set(r, c, "바뀜" if (r, c) == (7, 3) else v))
    code, out = run_excel(capsys, "write", "--cells", "C6=a; C7=b")
    assert code == 0 and out["verified"] is False and "C7" in " ".join(out["unverified"]), out


# ---------------------------------------------------------------------------
# 시트 보호 사전 확인과 엑셀이 받지 않는 수식 (A6, 1.4.55)
# ---------------------------------------------------------------------------


def _protect(ws, unlocked=(), **allow):
    ws.ProtectContents = True
    ws.unlocked = set(unlocked)
    for k, v in allow.items():
        setattr(ws.Protection, k, v)


@pytest.mark.parametrize("argv", [
    ("write", "--range", "C6", "--value", "x"),
    ("write", "--cells", "C6=x; C7=y"),
    ("replace", "--find", "a", "--replace", "b"),
    ("style", "--range", "A1:B2", "--bold"),
    ("range", "--action", "merge", "--range", "A1:B1"),
    ("range", "--action", "width", "--range", "A1", "--size", "20"),
    ("range", "--action", "insert", "--range", "3:4"),
    ("range", "--action", "sort", "--range", "A1:B3", "--by", "A"),
    ("comment", "--action", "add", "--range", "B2", "--text", "의견"),
])
def test_protected_sheet_stops_every_write(xl, capsys, argv):
    _, wb = xl
    ws = ws_of(wb)
    for r, row in enumerate([["a", "x"], ["a", "y"], ["c", "z"]], 1):
        for c, v in enumerate(row, 1):
            ws.set(r, c, v)
    _protect(ws, unlocked=[(9, 4), (10, 4)])
    ws.ProtectDrawingObjects = True
    before = dict(ws.values)
    code, out = run_excel(capsys, *argv)
    assert code == 1 and out["error_type"] == "sheet_protected", out
    assert out["agent_directive"].startswith("STOP_RETRY:") and "시트 보호 해제는 사용자가 직접" in out["message"]
    assert "암호를 풀지 않음" in out["message"]
    assert ws.values == before and undo_points() == [] and ws.notes == {}


def test_protected_sheet_lists_unlocked_cells_as_candidates(xl, capsys):
    _, wb = xl
    ws = ws_of(wb)
    ws.set(1, 1, "제목")
    ws.set(9, 4, None)
    ws.set(5, 3, 1)
    _protect(ws, unlocked=[(5, 3)])
    code, out = run_excel(capsys, "write", "--range", "A1", "--value", "x", "--overwrite")
    assert code == 1 and out["error_type"] == "sheet_protected" and "A1" in out["message"], out
    assert out["candidates"] == ["C5"] and "[후보] C5" in out["message"]


def test_protected_sheet_allows_unlocked_cells_and_allowed_formatting(xl, capsys):
    _, wb = xl
    ws = ws_of(wb)
    _protect(ws, unlocked=[(6, 3), (7, 3)], AllowFormattingCells=True)
    code, out = run_excel(capsys, "write", "--cells", "C6=1; C7=2")
    assert code == 0 and out["verified"] is True, out
    code, out = run_excel(capsys, "style", "--range", "C6:C7", "--bold")
    assert code == 0, out


def test_unprotected_sheet_comment_and_chart_unaffected(xl, capsys):
    _, wb = xl
    ws = ws_of(wb)
    ws.ProtectDrawingObjects = True          # 시트 보호가 꺼져 있으면 개체 보호 표시는 의미 없음
    code, out = run_excel(capsys, "comment", "--action", "add", "--range", "B2", "--text", "의견")
    assert code == 0, out


def test_rejected_formula_names_formula_and_recommends_comma(xl, capsys):
    _, wb = xl
    ws = ws_of(wb)
    ws.reject_formula = lambda f: ";" in f
    code, out = run_excel(capsys, "write", "--range", "C8", "--formula", "=SUM(A1;A2)")
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:"), out
    assert "=SUM(A1;A2)" in out["message"] and "쉼표" in out["message"] and "영어 함수 이름" in out["message"]
    assert "[추천] python" in out["message"] and "=SUM(A1,A2)" in out["message"].split("[추천]")[1]
    assert (8, 3) not in ws.formulas and undo_points() == []


def test_rejected_formula_without_mechanical_fix_has_no_recommendation(xl, capsys):
    _, wb = xl
    ws = ws_of(wb)
    ws.reject_formula = lambda f: "합계" in f
    code, out = run_excel(capsys, "write", "--cells", "C6=1; C7==합계(A1:A2); C8=3")
    assert code == 1 and "=합계(A1:A2)" in out["message"] and "[추천]" not in out["message"], out
    assert "큰따옴표" in out["message"]
    assert ws.values == {} and ws.formulas == {} and undo_points() == []      # 먼저 쓴 C6도 되돌림


def test_rejected_formula_rolls_back_new_sheet(xl, capsys, monkeypatch):
    app, wb = xl
    real_add = wb.Worksheets.Add

    def add(**kw):
        ws = real_add(**kw)
        ws.reject_formula = lambda f: "SUMA" in f
        return ws
    monkeypatch.setattr(wb.Worksheets, "Add", add)
    code, out = run_excel(capsys, "write", "--sheet", "계산", "--new-sheet", "--range", "A1:A3", "--header", "합",
                          "--formula", "=SUMA(B1)")
    assert code == 1 and "=SUMA(B1)" in out["message"], out
    assert sheet_names(wb) == ["Sheet1", "요약"] and undo_points() == []


def test_rejected_formula_log_has_no_formula(xl, capsys):
    _, wb = xl
    ws_of(wb).reject_formula = lambda f: "VLOOKUPX" in f
    code, out = run_excel(capsys, "write", "--range", "C8", "--formula", "=VLOOKUPX(1)")
    assert code == 1
    log = _tool_log_text()
    assert "VLOOKUPX" not in log and "수식" in log


# ---------------------------------------------------------------------------
# 독립 검토 (1.4.63): 칸값·메모 파일의 셋째 열, --cells 따옴표·수식 나누기, --append 재실행
# ---------------------------------------------------------------------------


def test_write_cells_file_third_column_stops_before_writing(xl, capsys, tmp_path):
    # 값의 쉼표가 열을 나눔: C7,1,500 -> 1만 쓰고 VERIFIED였음
    _, wb = xl
    f = tmp_path / "칸값.csv"
    f.write_text("칸,값\nC6,ft\nC7,1,500\n", encoding="utf-8")
    code, out = run_excel(capsys, "write", "--cells", str(f))
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:"), out
    assert "3번째 줄" in out["message"] and "큰따옴표" in out["message"]
    assert '고친 줄: 3번째 줄 C7,"1,500"' in out["message"] and "[추천]" not in out["message"]   # 인자 값으로 읽히지 않게
    assert ws_of(wb).values == {} and undo_points() == []
    assert "1,500" not in _tool_log_text()                               # 고친 줄(값)은 로그에 남기지 않음
    f.write_text("칸,값\nC6,ft\nC7,\"1,500\"\nC8,x,\n", encoding="utf-8")   # 따옴표로 감싸거나 뒤 빈 칸은 그대로
    code, out = run_excel(capsys, "write", "--cells", str(f))
    assert code == 0 and ws_of(wb).values[(7, 3)] == 1500 and ws_of(wb).values[(8, 3)] == "x", out


def test_comment_file_third_column_stops_before_adding(xl, capsys, tmp_path):
    _, wb = xl
    f = tmp_path / "메모.csv"
    f.write_text("칸,메모\nB5,규격 3.1과 다름, 재확인 필요\nB6,좋음\nB7,a,b\n", encoding="utf-8")
    code, out = run_excel(capsys, "comment", "--action", "add", "--file", str(f))
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:"), out
    assert "2, 4번째 줄" in out["message"] and "큰따옴표" in out["message"]
    assert 'B5,"규격 3.1과 다름, 재확인 필요"' in out["message"]
    assert ws_of(wb).notes == {} and undo_points() == []


@pytest.mark.parametrize("cells,want", [
    ("C6='007; C7=5", {(6, 3): "007", (7, 3): 5}),                            # 닫히지 않은 따옴표가 뒤 항목을 삼키지 않음
    ("C6='it's'; C7=1", {(6, 3): "it's", (7, 3): 1}),
    ("C6='a;b'; C7=2", {(6, 3): "a;b", (7, 3): 2}),
    ("'요약'!C3=5; C4=6", {(4, 3): 6}),
])
def test_write_cells_quote_must_close_before_separator(xl, capsys, cells, want):
    _, wb = xl
    code, out = run_excel(capsys, "write", "--cells", cells)
    assert code == 0 and out["verified"] is True, out
    assert ws_of(wb).values == want
    if "007" in cells:
        assert (6, 3) in ws_of(wb).prefixed                               # Excel 글자 표시 '는 그대로 (앞자리 0 유지)


def test_write_cells_formula_keeps_separators_in_quotes_and_parens(xl, capsys):
    _, wb = xl
    ws = ws_of(wb)
    code, out = run_excel(capsys, "write", "--cells", 'C6==IF(A1="a;b",1,2); C7=3')
    assert code == 0 and ws.formulas[(6, 3)] == '=IF(A1="a;b",1,2)' and ws.values[(7, 3)] == 3, out
    ws.reject_formula = lambda f: ";" in f
    code, out = run_excel(capsys, "write", "--cells", "C8==SUM(A1;A2)")
    assert code == 1 and "[추천] python" in out["message"], out                # 수식을 쪼개지 않고 ';' -> ',' 추천까지
    assert "C8==SUM(A1,A2)" in out["message"].split("[추천]")[1]
    assert (8, 3) not in ws.formulas


def test_rejected_formula_rollback_keeps_text_constants(xl, capsys):
    # 되돌릴 때 rng.Formula = prior가 글자 007, 1-2, 긴 숫자 글자, '=abc를 숫자·날짜·수식으로 바꿨음 (1.4.63)
    _, wb = xl
    ws = ws_of(wb)
    ws.typing = True
    old = {(6, 3): "007", (7, 3): "1-2", (8, 3): "12345678901234567890", (9, 3): "=abc"}
    for (r, c), v in old.items():
        ws.set(r, c, "'" + v)
    ws.set(10, 3, 42)
    ws.reject_formula = lambda f: "BAD" in f
    code, out = run_excel(capsys, "write", "--cells", "C6=a; C7=b; C8=c; C9=d; C10=e; C11==BAD(1)", "--overwrite")
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:"), out
    assert {k: ws.values[k] for k in old} == old and ws.values[(10, 3)] == 42
    assert all(k in ws.prefixed for k in old) and ws.formulas == {}
    assert undo_points() == [] and "되돌리지 못한" not in out["message"]


def test_rejected_formula_rollback_mismatch_keeps_undo_point(xl, capsys, monkeypatch):
    # 되돌린 뒤 다시 읽은 값이 다르면 되돌리기 지점을 남기고 undo를 안내
    _, wb = xl
    ws = ws_of(wb)
    ws.set(6, 3, "원래")
    ws.reject_formula = lambda f: "BAD" in f
    real_set = ws.set
    monkeypatch.setattr(ws, "set", lambda r, c, v: real_set(r, c, "망가짐" if v == "'원래" else v))
    code, out = run_excel(capsys, "write", "--cells", "C6=a; C7==BAD(1)", "--overwrite")
    assert code == 1 and "C6" in out["message"] and "office.py undo" in out["message"], out
    assert len(undo_points()) == 1


def test_comment_append_rerun_does_not_duplicate(xl, capsys):
    _, wb = xl
    ws = ws_of(wb)
    _note(ws, 5, 2, "기존 메모")
    code, out = run_excel(capsys, "comment", "--range", "B5", "--text", "덧붙임", "--append")
    assert code == 0 and ws.notes[(5, 2)].text == "기존 메모\n덧붙임" and len(undo_points()) == 1
    code, out = run_excel(capsys, "comment", "--range", "B5", "--text", "덧붙임", "--append")
    assert code == 0 and ws.notes[(5, 2)].text == "기존 메모\n덧붙임", out          # 두 번 붙지 않음
    assert out["unchanged"] == 1 and out["count"] == 0 and out["applied"] == [] and len(undo_points()) == 1
    code, out = run_excel(capsys, "comment", "--range", "B5", "--text", "기존 메모\\n덧붙임")
    assert code == 0 and out["unchanged"] == 1, out                               # 같은 메모는 결정 없이 그대로


@pytest.mark.parametrize("cells,fixed", [
    ("C6=ft,C7=℃", "C6=ft; C7=℃"),                              # 회사 PC 2026-11-06 O2 (Gemma가 쉼표로 나눔)
    ("C6=ft, 요약!C7=20", "C6=ft; 요약!C7=20"),
    ("C5=a; C6=ft,C7=b", "C5=a; C6=ft; C7=b"),
    ("C6='ft',C7='x'", "C6='ft'; C7='x'"),                     # 따옴표가 쉼표 앞에서 닫힘 (값 전체를 묶은 것이 아님)
    ("C6=ft C7=20", "C6=ft; C7=20"),                             # 공백·| 로 나눔 (감사: C6에 'ft C7=20'이 들어갔음)
    ("C6=a|C7=b", "C6=a; C7=b"),
])
def test_write_cells_comma_separated_is_refused_with_fixed_command(xl, capsys, cells, fixed):
    """쉼표로 나눈 칸을 그대로 쓰면 C6에 'ft,C7=℃'가 들어간다. 쓰지 않고 ;로 바꾼 전체 명령을 [추천]."""
    _, wb = xl
    code, out = run_excel(capsys, "write", "--cells", cells)
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:"), out
    assert "[추천] python" in out["message"], out
    cmd = out["message"].split("[추천] ", 1)[1]
    argv = split_command(cmd)
    assert argv[argv.index("--cells") + 1] == fixed, cmd
    assert ws_of(wb).values == {}


@pytest.mark.parametrize("cells", ["C6=1,234", "C6=a,b", "C6==IF(A1=1,B2,0)", "C6='x,C7=1'", "C6=크기 3,4",
                                   "C6=R1=10k,R2=4.7k", "C6=T1=25, T2=30"])   # 값 자체가 '이름=값,이름=값' (1.4.77 검토)
def test_write_cells_comma_inside_value_is_kept(xl, capsys, cells):
    """값 안의 쉼표(숫자, 글, 수식, 따옴표로 묶은 값)는 칸 구분으로 보지 않음."""
    _, wb = xl
    code, out = run_excel(capsys, "write", "--cells", cells)
    assert code == 0 and out["status"] == "ok", out


def test_write_cells_value_with_name_value_pairs_is_written_as_is(xl, capsys):
    """1.4.77 검토 치명: 'C6=R1=10k,R2=4.7k'에 [추천]이 C6=R1=10k; R2=4.7k(다른 칸 R2에 씀)였음. 값 그대로 C6에."""
    _, wb = xl
    code, out = run_excel(capsys, "write", "--cells", "C6=R1=10k,R2=4.7k")
    assert code == 0 and out["verified"] is True, out
    assert ws_of(wb).values.get((6, 3)) == "R1=10k,R2=4.7k" and (2, 18) not in ws_of(wb).values


# ---------------------------------------------------------------------------
# --dry-run: replace, range clear/move, sheet delete, chart delete (바꾸지 않고 미리보기 -> confirm_commands[0] 실행)
# ---------------------------------------------------------------------------


def _confirm_argv(out, argv):
    """confirm_commands[0]이 python "<excel.py>" + (실행한 인자 - --dry-run)인지 확인하고 그 인자를 돌려줌."""
    assert out["dry_run"] is True and out["agent_directive"].startswith("VERIFIED: 미리보기"), out
    assert out["preview"].startswith("미리보기:") and out["preview"].endswith("(바꾸지 않음)"), out
    assert "next_commands" not in out and out["confirm_commands"][0]["destructive"] is True, out
    parts = split_command(out["confirm_commands"][0]["command"])
    assert parts[0] == "python" and parts[1].endswith("excel.py"), parts
    assert parts[2:] == [a for a in argv if a != "--dry-run"], parts
    return parts[2:]


def test_replace_dry_run_changes_nothing_then_confirm(xl_table, capsys):
    argv = ["replace", "--find", "영업", "--replace", "판매", "--dry-run"]
    code, out = run_excel(capsys, *argv)
    assert code == 0 and out["count"] == 2 and "칸 2개" in out["preview"], out
    assert out["planned"] == ["변경: Sheet1!B2 (영업 -> 판매)", "변경: Sheet1!B4 (영업 -> 판매)"], out
    real = _confirm_argv(out, argv)
    assert _col(xl_table, 2) == ["부서", "영업", "개발", "영업", "기획"] and off.list_snapshots("excel") == []
    code, out = run_excel(capsys, *real)
    assert code == 0 and out["changed_cells"] == 2 and _col(xl_table, 2) == ["부서", "판매", "개발", "판매", "기획"], out
    assert len(off.list_snapshots("excel")) == 1


def test_replace_dry_run_no_match_says_nothing_changes(xl_table, capsys):
    code, out = run_excel(capsys, "replace", "--find", "없는글", "--replace", "x", "--dry-run")
    assert code == 0 and out["count"] == 0 and "바뀔 것 없음" in out["preview"], out
    assert "confirm_commands" not in out and out["agent_directive"].startswith("VERIFIED:") and "0건" in out["agent_directive"], out


def test_range_clear_dry_run_then_confirm(xl_table, capsys):
    argv = ["range", "--action", "clear", "--range", "A2:B3", "--type", "contents", "--dry-run"]
    code, out = run_excel(capsys, *argv)
    assert code == 0 and out["count"] == 4 and out["filled_cells"] == 4 and "값이 있는 칸 4개" in out["preview"], out
    real = _confirm_argv(out, argv)
    assert _col(xl_table, 1) == ["이름", "가", "나", "다", "라"] and off.list_snapshots("excel") == []
    code, out = run_excel(capsys, *real)
    assert code == 0 and _col(xl_table, 1) == ["이름", None, None, "다", "라"] and xl_table.values.get((2, 2)) is None, out


def test_range_clear_dry_run_empty_range_is_zero(xl_table, capsys):
    code, out = run_excel(capsys, "range", "--action", "clear", "--range", "H20:J30", "--type", "contents", "--dry-run")
    assert code == 0 and out["count"] == 0 and "바뀔 것 없음" in out["preview"] and "confirm_commands" not in out, out


def test_range_move_dry_run_adds_overwrite_when_dest_has_values(xl_table, capsys):
    argv = ["range", "--action", "move", "--range", "A1:A2", "--dest", "C1", "--dry-run"]
    code, out = run_excel(capsys, *argv)
    assert code == 0 and out["count"] == 2 and out["overwrite_cells"] == 2 and "덮어써집니다" in out["preview"], out
    cmd = split_command(out["confirm_commands"][0]["command"])
    assert cmd[2:] == [a for a in argv if a != "--dry-run"] + ["--overwrite"], cmd
    assert xl_table.values[(1, 3)] == "금액" and off.list_snapshots("excel") == [] and not any(e[0] == "Cut" for e in xl_table.log)
    code, out = run_excel(capsys, *cmd[2:])
    assert code == 0 and ("Cut", "$A$1:$A$2", "$C$1") in xl_table.log, out


def test_range_move_dry_run_to_empty_dest(xl_table, capsys):
    argv = ["range", "--action", "move", "--range", "A1:B2", "--dest", "F1", "--dry-run"]
    code, out = run_excel(capsys, *argv)
    assert code == 0 and out["count"] == 4 and "overwrite_cells" not in out, out
    real = _confirm_argv(out, argv)
    code, out = run_excel(capsys, *real)
    assert code == 0 and ("Cut", "$A$1:$B$2", "$F$1") in xl_table.log, out


def test_sheet_delete_dry_run_then_confirm(xl, capsys):
    app, wb = xl
    _formula(ws_of(wb), 1, 5, "=요약!A1*2")          # 지울 시트를 참조 -> #REF!
    argv = ["sheet", "--action", "delete", "--name", "요약", "--dry-run"]
    code, out = run_excel(capsys, *argv)
    assert code == 0 and out["count"] == 1 and out["broken_refs"] == 1 and out["broken_ref_cells"] == ["Sheet1!E1"], out
    assert "'요약'" in out["preview"] and "#REF!" in out["preview"] and "A1*2" not in out["preview"], out
    real = _confirm_argv(out, argv)
    assert sheet_names(wb) == ["Sheet1", "요약"] and off.list_snapshots("excel") == []
    code, out = run_excel(capsys, *real)
    assert code == 0 and sheet_names(wb) == ["Sheet1"], out


def test_chart_delete_dry_run_then_confirm(xl, capsys):
    app, wb = xl
    ws = ws_of(wb)
    ws.set(1, 1, "a")
    ws.set(2, 1, 1.0)
    run_excel(capsys, "chart", "--range", "A1:A2", "--title", "매출")
    run_excel(capsys, "chart", "--range", "A1:A2", "--title", "비용")
    snaps = len(off.list_snapshots("excel"))
    argv = ["chart", "--action", "delete", "--name", "매출", "--dry-run"]
    code, out = run_excel(capsys, *argv)
    assert code == 0 and out["count"] == 1 and out["planned"] == ["삭제: Sheet1!매출"] and "'매출'" in out["preview"], out
    real = _confirm_argv(out, argv)
    assert _names(wb)["Sheet1"] == ["매출", "비용"] and len(off.list_snapshots("excel")) == snaps
    code, out = run_excel(capsys, *real)
    assert code == 0 and _names(wb)["Sheet1"] == ["비용"], out
    argv = ["chart", "--action", "delete", "--all", "--dry-run"]
    code, out = run_excel(capsys, *argv)
    assert code == 0 and out["count"] == 1 and "1개" in out["preview"], out
    real = _confirm_argv(out, argv)
    assert _names(wb)["Sheet1"] == ["비용"]
    code, out = run_excel(capsys, *real)
    assert code == 0 and _names(wb)["Sheet1"] == [], out
    code, out = run_excel(capsys, "chart", "--action", "delete", "--all", "--dry-run")
    assert code == 0 and out["count"] == 0 and "바뀔 것 없음" in out["preview"] and "confirm_commands" not in out, out
    code, out = run_excel(capsys, "chart", "--action", "list", "--dry-run")      # 읽기만 하는 동작: 받고 warnings (1.4.86)
    assert code == 0 and out["count"] == 0 and any("--dry-run을 무시" in w for w in out["warnings"]), out


def test_style_many_cells_in_one_command(xl, capsys):
    """1.4.91: 행마다 다른 칸을 칠할 때 칸마다 명령을 따로 실행하던 것 -> --range "B2;D5;F7"(칸 여러 개는 ;) 명령 한 번으로,
    칸마다 다시 읽어 확인. ','로 이으면 ';'로 고친 전체 명령 [추천], 떨어진 칸에 --theme은 받지 않음."""
    app, wb = xl
    ws = ws_of(wb)
    for r in range(1, 21):
        for c in range(1, 9):
            ws.set(r, c, float(r * c))
    code, out = run_excel(capsys, "style", "--range", "B2;D5;F7", "--bg", "red")
    assert code == 0 and out["verified"] is True and out["areas"] == 3 and out["range"] == "B2;D5;F7", out["message"] if code else out
    colored = {k[0] for k in ws.styles if k[1] == "Interior"}
    assert any(a.replace("$", "") == "B2,D5,F7" for a in colored), colored        # Excel에는 한 번 (B2,D5,F7)
    many = ";".join(f"{c}{r}" for c in "EFGH" for r in range(1, 21))
    code, out = run_excel(capsys, "style", "--range", many, "--bold")
    assert code == 0 and out["verified"] is True and out["areas"] == 80, out["message"] if code else out          # 255자 넘는 목록은 나눠서 적용
    code, out = run_excel(capsys, "style", "--range", "B2,D5", "--bg", "red")
    assert code == 1 and "[추천]" in out["message"] and '--range "B2;D5"' in out["message"].replace("'", '"'), out
    code, out = run_excel(capsys, "style", "--range", "B2;D5", "--theme", "navy")
    assert code == 1 and "--theme" in out["message"], out


# ---------------------------------------------------------------------------
# 2026-11-09 검토 B4: 칸 하나에 수식 여러 개면 [추천] 범위는 옆 열 데이터의 끝까지 (100행을 지어내지 않음)
# ---------------------------------------------------------------------------


def test_multi_formula_single_cell_recommends_table_rows(xl_table, capsys):
    code, out = run_main(excel, capsys, "write", "--range", "D2", "--formula", "=B2*2", "=C2*3")
    assert code == 1 and "[추천]" in out["message"], out
    rec = recommended_argv(out["message"], "excel")
    assert option_value(rec, "--range") == "D2:E5", rec
    code, out = run_main(excel, capsys, "write", "--range", "H2", "--formula", "=B2*2", "=C2*3")
    assert code == 1 and "[추천]" not in out["message"] and "마지막 행" in out["message"], out


# ---------------------------------------------------------------------------
# 2026-11-09 검토 A7: 통화 서식 칸(pywin32가 decimal.Decimal로 줌)도 숫자로 (정렬, 머리글 판별, --where)
# ---------------------------------------------------------------------------


def test_currency_cells_are_numbers(xl_table, capsys):
    from decimal import Decimal
    for r, v in ((2, "10"), (3, "100"), (4, "9"), (5, "1000")):
        xl_table.values[(r, 3)] = Decimal(v)
    code, out = run_main(excel, capsys, "read", "--where", "금액 >= 10")
    assert code == 0 and out["count"] == 3, out
    code, out = run_main(excel, capsys, "range", "--action", "sort", "--range", "A1:C5", "--by", "C")
    assert code == 0 and "이미 정렬" not in out["preview"], out
    assert [float(xl_table.values[(r, 3)]) for r in range(2, 6)] == [9, 10, 100, 1000], out


# ---------------------------------------------------------------------------
# 2026-11-09 검토 A8: + - @ = 로 시작하는 글은 글자 그대로, 글을 썼는데 글이 아닌 값이 되면 UNVERIFIED
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("text", ["+82-10-1234-5678", "-007", "- 검토 필요", "@담당", "=아님"])
def test_text_starting_with_formula_marks_stays_text(xl, capsys, text):
    _, wb = xl
    ws = ws_of(wb)
    ws.typing = True                         # 실제 Excel처럼 대입한 글을 해석
    code, out = run_excel(capsys, "write", "--range", "A1", "--value", text)
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    assert ws.values[(1, 1)] == text and (1, 1) in ws.prefixed, ws.values


def test_text_changed_by_excel_is_not_verified():
    assert excel._same_cell("+82-10-1234-5678", -6840) is False
    assert excel._same_cell("- 검토", -2146826259) is False
    assert excel._same_cell("2026-01-05", 46027.0) is True and excel._same_cell("true", True) is True
    assert excel._same_cell("abc", 5) is False and excel._same_cell(5, 5.0) is True


# ---------------------------------------------------------------------------
# 2026-11-09 검토 A9: 수식은 글자만이 아니라 값도 확인 (오류 값이면 UNVERIFIED), 동적 배열 함수는 Formula2
# ---------------------------------------------------------------------------


def test_formula_with_error_value_is_unverified(xl_table, capsys):
    code, out = run_main(excel, capsys, "write", "--range", "D2", "--formula", "=합계(B2:B3)")
    assert code == 0 and out["agent_directive"].startswith("UNVERIFIED:") and "#NAME?" in out["agent_directive"], out
    code, out = run_main(excel, capsys, "write", "--cells", "E2==합계(C2:C3)")
    assert code == 0 and out["agent_directive"].startswith("UNVERIFIED:") and "#NAME?" in out["agent_directive"], out
    code, out = run_main(excel, capsys, "write", "--range", "F2", "--formula", "=SUM(C2:C3)")
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out


def test_dynamic_array_formula_uses_formula2(xl_table, capsys):
    code, out = run_main(excel, capsys, "write", "--range", "E2", "--formula", "=UNIQUE(B2:B5)")
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    assert any(e[0] == "Formula2" for e in xl_table.log), xl_table.log[-5:]
    xl_table.no_formula2 = True                       # Office 2019: Formula2 없음 -> Formula
    code, out = run_main(excel, capsys, "write", "--range", "F2", "--formula", "=SORT(B2:B5)")
    assert code == 0 and xl_table.formulas[(2, 6)] == "=SORT(B2:B5)", out


# ---------------------------------------------------------------------------
# 2026-11-09 검토 B8: write가 아닌 쓰기 명령은 경로 오타면 빈 새 파일을 만들지 않음
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("argv", [["style", "--range", "A1:C3", "--bold"], ["replace", "--find", "a", "--replace", "b"],
                                  ["sheet", "--action", "add", "--name", "새"], ["range", "--action", "freeze", "--range", "A2"]])
def test_non_write_commands_do_not_create_missing_file(xl, capsys, tmp_path, argv):
    missing = tmp_path / "보고서_오타.xlsx"
    code, out = run_excel(capsys, argv[0], "--target", str(missing), *argv[1:])
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:") and not missing.exists(), out


# ---------------------------------------------------------------------------
# 2026-11-09 검토 B12: 잘못된 주소는 바꾸기 전에 STOP_RETRY (ERROR가 아님)
# ---------------------------------------------------------------------------


def test_bad_addresses_are_input_errors_before_change(xl, capsys):
    app, wb = xl
    ws = ws_of(wb)
    ws.values[(1, 1)], ws.values[(2, 1)] = "a", 1
    code, out = run_excel(capsys, "range", "--action", "copy", "--range", "A1:A2", "--dest", "요약")
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:"), out
    assert option_value(recommended_argv(out["message"], "excel"), "--dest") == "요약!A1", out
    code, out = run_excel(capsys, "chart", "--range", "A1:A2", "--at", "요약!E1")
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:") and "--at" in out["message"], out
    app.ActiveWindow.FreezePanes = True
    code, out = run_excel(capsys, "range", "--action", "freeze", "--range", "3:3")
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:"), out
    assert option_value(recommended_argv(out["message"], "excel"), "--range") == "A3" and app.ActiveWindow.FreezePanes is True


# ---------------------------------------------------------------------------
# 2026-11-09 검토 B13 (엑셀 쓰기): 뒤집힌 범위, --align·--border 확인, 닫힌 파일 select, --append 시작 열
# ---------------------------------------------------------------------------


def test_b13_excel_write_small_fixes(xl, capsys):
    _, wb = xl
    ws = ws_of(wb)
    code, out = run_excel(capsys, "write", "--range", "C5:A1", "--value", "x")
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:") and "-" not in out["preview"].split("(")[-1], out
    assert ws.values[(1, 1)] == "x" and ws.values[(5, 3)] == "x"
    code, out = run_excel(capsys, "style", "--range", "A1:B2", "--align", "center")
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    code, out = run_excel(capsys, "style", "--range", "A1:B2", "--border", "thin")
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out


def test_append_without_range_starts_at_table_column(xl, capsys):
    _, wb = xl
    ws = ws_of(wb)
    for r, row in enumerate([("이름", "값"), ("a", 1), ("b", 2)], start=3):
        for c, v in enumerate(row, start=2):
            ws.values[(r, c)] = v
    code, out = run_excel(capsys, "write", "--append", "--values", '[["c", 3]]')
    assert code == 0 and ws.values.get((6, 2)) == "c" and ws.values.get((6, 3)) == 3 and (6, 1) not in ws.values, ws.values


def test_select_sheet_in_closed_file_is_refused(xl, capsys, tmp_path, monkeypatch):
    import openpyxl
    from fake_office import FakeExcelApp
    monkeypatch.setattr(off, "_new_instance", lambda kind: FakeExcelApp())
    p = tmp_path / "닫힘.xlsx"
    w = openpyxl.Workbook()
    w.create_sheet("둘")
    w.save(p)
    code, out = run_excel(capsys, "sheet", "--target", str(p), "--action", "select", "--name", "Sheet2")
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:") and "열린 통합문서" in out["message"], out


def test_c11_append_does_not_overwrite_filtered_rows(xl, capsys):
    """2026-11-09 검토 C11: End(xlUp)은 필터로 숨긴 행을 건너뜀 -> --append가 숨은 데이터 행에 쓰던 것."""
    _, wb = xl
    ws = ws_of(wb)
    for r, v in enumerate(["이름", "a", "b", "숨은 c"], start=1):
        ws.values[(r, 1)] = v
    ws.hidden_rows = {4}
    code, out = run_excel(capsys, "write", "--range", "A1", "--append", "--value", "새")
    assert code == 0 and ws.values[(4, 1)] == "숨은 c" and ws.values[(5, 1)] == "새", ws.values
