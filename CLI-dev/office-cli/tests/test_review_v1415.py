"""1.4.15 회귀 테스트: 6차 점검에서 찾은 오류 (P1 예약된 시트 이름 History, P2 transfer --sheet [추천]이 있는 시트와 겹침,
P3 휘발성 수식이 있는 닫힌 파일의 바꾼 것 없는 성공, P4 바꾼 것 없는 성공이 바꾼 것으로 표시되던 명령들).

항목마다 실패 상황을 재현하고 고친 동작을 확인합니다. 가짜 COM(fake_*.py)을 씁니다.
가짜 Excel은 실제 Excel처럼 쓸 수 없는 시트 이름(History 포함)을 Name에서 COM 오류로 거절합니다.
"""
import json
import types

import pytest

import _office as off
import _runtime as rt
import excel
import office
import ppt
import word
from fake_excel import FakeExcelWorkbook, FakeXlApp, SheetNameRefused
from fake_office import FakeExcelApp, FakeWorkbooks, make_win32com
from fake_ppt import FakePresentation
from fake_word import FakeWordDocument


def _run(module, capsys, *argv):
    rt.RECOVERED.clear()
    with pytest.raises(SystemExit) as ei:
        module.main(list(argv))
    out = [ln for ln in capsys.readouterr().out.splitlines() if ln.strip()]
    assert len(out) == 1, out
    return ei.value.code, json.loads(out[0])


def _com(monkeypatch):
    fake, _ = make_win32com()
    monkeypatch.setattr(off, "win32com", fake)
    monkeypatch.setattr(off, "pythoncom", None)
    monkeypatch.setattr(off, "list_open", lambda kind: [])


def _active(monkeypatch, app, doc):
    monkeypatch.setattr(off, "get_active", lambda kind: {"name": doc.Name, "path": doc.FullName, "app": app, "doc": doc})


@pytest.fixture
def xl(monkeypatch):
    app = FakeXlApp()
    wb = FakeExcelWorkbook(app, "보고.xlsx", sheets=("Sheet1", "a_b"))
    _com(monkeypatch)
    _active(monkeypatch, app, wb)
    return app, wb


@pytest.fixture
def pp(monkeypatch):
    prs = FakePresentation(3)
    app = types.SimpleNamespace(ActiveWindow=types.SimpleNamespace(View=types.SimpleNamespace(GotoSlide=lambda n: None),
                                                                   Presentation=prs))
    _com(monkeypatch)
    _active(monkeypatch, app, prs)
    return prs


@pytest.fixture
def wd(monkeypatch):
    doc = FakeWordDocument("추진 배경\r2026년 계획입니다.\r일정")
    _com(monkeypatch)
    _active(monkeypatch, object(), doc)
    return doc


def _points(kind="excel"):
    return sorted(s["meta_file"] for s in off.list_snapshots(kind))


def _sheet_names(wb):
    return [str(wb.Worksheets(i).Name) for i in range(1, wb.Worksheets.Count + 1)]


# ---------------------------------------------------------------------------
# P1: 엑셀이 예약한 시트 이름 History(대소문자 무관)는 시트를 만들기 전에 막음
# ---------------------------------------------------------------------------

def test_p1_fake_excel_refuses_reserved_name_like_real_excel(xl):
    _, wb = xl
    with pytest.raises(SheetNameRefused):
        wb.Worksheets.Add().Name = "history"


@pytest.mark.parametrize("argv", [
    ["sheet", "--action", "add", "--name", "History"],
    ["sheet", "--action", "add", "--name", "HISTORY"],
    ["write", "--new-sheet", "--sheet", "history", "--range", "A1", "--value", "1"],
    ["sheet", "--action", "rename", "--name", "a_b", "--new-name", "History"],
    ["sheet", "--action", "copy", "--name", "Sheet1", "--new-name", "History"],
])
def test_p1_reserved_history_refused_before_any_change(xl, capsys, argv):
    _, wb = xl
    code, out = _run(excel, capsys, *argv)
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:"), out
    assert "예약" in out["message"] and "[추천] " in out["message"], out
    assert out["message"].split("[추천] ")[1].strip().lower() == "history_", out
    assert _sheet_names(wb) == ["Sheet1", "a_b"] and _points() == []        # 빈 시트·복사본이 남지 않음


def test_p1_reserved_recommendation_avoids_existing_and_works(xl, capsys):
    _, wb = xl
    wb.Worksheets.Add().Name = "History_"
    code, out = _run(excel, capsys, "sheet", "--action", "add", "--name", "History")
    assert code == 1 and "[추천] History_ 2" in out["message"], out
    code, out = _run(excel, capsys, "sheet", "--action", "add", "--name", "History_ 2")    # [추천]대로 1회 재시도
    assert code == 0 and _sheet_names(wb)[-1] == "History_ 2", out


def test_p1_reserved_name_through_target_sheet_create(xl):
    _, wb = xl
    with pytest.raises(rt.InputError):
        with off.open_target("excel", "active", readonly=False, tool="excel", command="write") as t:
            t.sheet("History", create=True)
    assert _sheet_names(wb) == ["Sheet1", "a_b"] and _points() == []


def test_p1_names_containing_history_are_allowed():
    off.check_sheet_name("History 2026", [])
    off.check_sheet_name("작업History", [])


# ---------------------------------------------------------------------------
# P2: transfer --sheet의 [추천]은 받는 문서의 시트 이름을 안 뒤에 (따라 한 번 다시 실행하면 성공)
# ---------------------------------------------------------------------------

def test_p2_transfer_recommendation_avoids_existing_sheet(xl):
    _, wb = xl
    args = types.SimpleNamespace(sheet="a/b", to_target="active", overwrite=False)
    with pytest.raises(rt.InputError) as ei:
        office._to_excel(args, [["h"], ["v"]], "Word 표")
    rec = str(ei.value).split("[추천] ")[1].strip()
    assert rec == "a_b 2", rec                               # 있는 시트 a_b와 겹치지 않음
    assert _sheet_names(wb) == ["Sheet1", "a_b"] and _points() == []
    args.sheet = rec
    out = office._to_excel(args, [["h"], ["v"]], "Word 표")   # [추천]대로 1회 재시도하면 성공
    assert out["sheet"] == "a_b 2" and _sheet_names(wb) == ["Sheet1", "a_b", "a_b 2"]


def test_p2_transfer_reserved_name_into_open_document(xl):
    _, wb = xl
    args = types.SimpleNamespace(sheet="History", to_target="active", overwrite=False)
    with pytest.raises(rt.InputError) as ei:
        office._to_excel(args, [["h"], ["v"]], "Word 표")
    assert "[추천] History_" in str(ei.value)
    assert _sheet_names(wb) == ["Sheet1", "a_b"] and _points() == []


def test_p2_new_file_target_is_checked_before_creating(monkeypatch, tmp_path):
    def boom(*a, **k):
        raise AssertionError("새 파일을 만들면 안 됨")
    monkeypatch.setattr(off, "open_target", boom)
    target = tmp_path / "새.xlsx"
    args = types.SimpleNamespace(sheet="History", to_target=str(target), overwrite=False)
    with pytest.raises(rt.InputError) as ei:
        office._to_excel(args, [["h"]], "Word 표")
    assert "[추천] History_" in str(ei.value) and not target.exists()
    assert office._new_file_target(str(target)) and not office._new_file_target("active")


# ---------------------------------------------------------------------------
# P3: 도구가 연 닫힌 파일은 바꾼 것(t.changed)이 있을 때만 저장
#     (휘발성 수식으로 열자마자 저장 안 된 상태여도 바꾼 것 없는 성공은 다시 저장하지 않고 지점도 남기지 않음)
# ---------------------------------------------------------------------------

def _volatile_open(monkeypatch):
    orig_open = FakeWorkbooks.Open

    def vopen(self, p, **kw):
        wb = orig_open(self, p, **kw)
        wb.Saved = False                     # Excel: 휘발성 수식(NOW, RAND)이 열 때 다시 계산되어 저장 안 된 상태

        def save():
            wb.rec("Save")
            wb.Saved = True
            with open(p, "wb") as f:
                f.write(b"resaved")
        wb.Save = save
        return wb
    monkeypatch.setattr(FakeWorkbooks, "Open", vopen)
    monkeypatch.setattr(off, "_new_instance", lambda kind: FakeExcelApp())


def test_p3_closed_volatile_file_noop_is_not_resaved(monkeypatch, tmp_path):
    path = tmp_path / "vol.xlsx"
    path.write_bytes(b"original")
    _com(monkeypatch)
    _volatile_open(monkeypatch)
    with off.open_target("excel", str(path), readonly=False, tool="excel", command="replace"):
        pass                                 # 아무것도 바꾸지 않음
    assert _points() == [] and path.read_bytes() == b"original"


def test_p3_closed_volatile_file_real_change_is_saved(monkeypatch, tmp_path):
    path = tmp_path / "vol.xlsx"
    path.write_bytes(b"original")
    _com(monkeypatch)
    _volatile_open(monkeypatch)
    with off.open_target("excel", str(path), readonly=False, tool="excel", command="write") as t:
        t.set_focus()                        # 바꾼 명령은 set_focus(또는 t.changed = True)로 표시
    assert path.read_bytes() == b"resaved" and len(_points()) == 1


# ---------------------------------------------------------------------------
# P4: 바꾼 것 없이 성공한 명령은 바꾼 것으로 표시하지 않음 (되돌리기 지점을 남기지 않고 0곳 바뀜으로 보고)
# ---------------------------------------------------------------------------

def test_p4_ppt_move_to_same_position_keeps_previous_point(pp, capsys):
    code, out = _run(ppt, capsys, "notes", "--slide", "1", "--text", "실제 변경")
    assert code == 0, out
    before = _points("ppt")
    assert len(before) == 1
    code, out = _run(ppt, capsys, "slide", "--action", "move", "--slide", "2", "--to", "2")
    assert code == 0 and "바뀐 것 없음" in out["preview"], out
    assert _points("ppt") == before                         # 다음 undo는 notes를 되돌림
    code, out = _run(ppt, capsys, "slide", "--action", "move", "--slide", "2", "--to", "3")
    assert code == 0 and len(_points("ppt")) == 2, out      # 실제로 옮기면 지점이 남음


def test_p4_ppt_notes_same_text_is_noop(pp, capsys):
    code, out = _run(ppt, capsys, "notes", "--slide", "1", "--text", "같은 노트")
    assert code == 0
    before = _points("ppt")
    code, out = _run(ppt, capsys, "notes", "--slide", "1", "--text", "같은 노트", "--overwrite")
    assert code == 0 and out.get("unchanged") is True and "바뀐 것 없음" in out["preview"], out
    assert _points("ppt") == before


def test_p4_ppt_replace_same_text_reports_zero(pp, capsys):
    pp.Slides(1).Shapes.AddTextbox(1, 0, 0, 100, 40).TextFrame.TextRange.Text = "A 계획"
    code, out = _run(ppt, capsys, "replace", "--map", '{"A": "A"}')
    assert code == 0 and out["replaced"] == 0 and out["unchanged_same_text"] == {"A": 1}, out
    assert "0곳" in out["preview"] and _points("ppt") == []


def test_p4_ppt_replace_find_only_is_refused(pp, capsys):
    code, out = _run(ppt, capsys, "replace", "--find", "A")
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:"), out
    assert _points("ppt") == []


@pytest.mark.parametrize("argv", [
    ["replace", "--find", "계획", "--replace", "계획"],
    ["replace", "--map", '{"계획": "계획"}'],
])
def test_p4_word_replace_same_text_reports_zero(wd, capsys, argv):
    code, out = _run(word, capsys, *argv)
    assert code == 0 and out["replaced"] == 0 and out["unchanged_same_text"] == {"계획": 1}, out
    assert "0곳" in out["preview"] and "처리" not in out["preview"], out
    assert _points("word") == [] and wd.Saved is True


def test_p4_word_replace_same_text_with_format_still_changes(wd, capsys):
    code, out = _run(word, capsys, "replace", "--find", "계획", "--replace", "계획", "--bold")
    assert code == 0 and out["replaced"] == 1, out


@pytest.mark.parametrize("argv", [
    ["replace", "--find", "초안", "--replace", "초안"],
    ["replace", "--map", '{"초안": "초안"}'],
])
def test_p4_excel_replace_same_text_reports_zero(xl, capsys, argv):
    _, wb = xl
    wb.Worksheets(1).set(1, 1, "초안 보고")
    code, out = _run(excel, capsys, *argv)
    assert code == 0 and out["changed_cells"] == 0 and out["unchanged_same_text"] == 1, out
    assert "0칸 바뀜" in out["preview"] and "처리" not in out["preview"], out
    assert _points() == []


def test_p4_excel_replace_formula_cell_without_format_is_not_counted(xl, capsys):
    _, wb = xl
    ws = wb.Worksheets(1)
    ws.Range("A1").Formula = '="초안"'
    code, out = _run(excel, capsys, "replace", "--find", "초안", "--replace", "최종")
    assert code == 0 and out["changed_cells"] == 0 and out["formula_cells_kept"], out
    assert _points() == []


@pytest.mark.parametrize("argv", [
    ["range", "--action", "unhide", "--range", "B:C"],
    ["range", "--action", "unhide", "--range", "2:3"],
    ["range", "--action", "unmerge", "--range", "A1:B2"],
    ["range", "--action", "unfreeze"],
    ["range", "--action", "clear", "--range", "D10:E12"],
    ["range", "--action", "clear", "--range", "A1:B2", "--type", "contents"],
    ["range", "--action", "move", "--range", "A5:B5", "--dest", "A5"],
])
def test_p4_excel_range_noop_drops_point(xl, capsys, argv):
    _, wb = xl
    wb.Worksheets(1).set(5, 1, "값")                      # 사용 범위 A5:A5 (A1:B2 칸은 비어 있음)
    code, out = _run(excel, capsys, *argv)
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    assert "바뀐 것 없음" in out["preview"] and out.get("unchanged") is True, out
    assert _points() == []
    assert wb.Worksheets(1).values[(5, 1)] == "값"


def test_p4_excel_range_real_changes_keep_point(xl, capsys):
    app, wb = xl
    ws = wb.Worksheets(1)
    ws.set(1, 1, "값")
    ws.hidden["cols"].add(2)
    code, out = _run(excel, capsys, "range", "--action", "unhide", "--range", "B:B")
    assert code == 0 and "바뀐 것 없음" not in out["preview"] and ws.hidden["cols"] == set(), out
    assert len(_points()) == 1
    ws.merged.add("$A$1:$B$2")
    code, out = _run(excel, capsys, "range", "--action", "unmerge", "--range", "A1:B2")
    assert code == 0 and ws.merged == set() and len(_points()) == 2, out
    app.ActiveWindow.FreezePanes = True
    code, out = _run(excel, capsys, "range", "--action", "unfreeze")
    assert code == 0 and app.ActiveWindow.FreezePanes is False and len(_points()) == 3, out
    code, out = _run(excel, capsys, "range", "--action", "clear", "--range", "A1:B2")
    assert code == 0 and (1, 1) not in ws.values and len(_points()) == 4, out
