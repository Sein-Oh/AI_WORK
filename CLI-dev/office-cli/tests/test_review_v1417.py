"""1.4.17 회귀 테스트: 8차 점검에서 찾은 오류 (F1 빈 찾을 글, F2 기본 시트가 여럿인 새 파일로 transfer,
F3 ppt/word 대소문자만 다른 바꾸기에서 이미 바꿀 글인 곳을 바꾼 것으로 셈).

항목마다 실패 상황을 재현하고 고친 동작을 확인합니다. 가짜 COM(fake_*.py)을 씁니다.
"""
import json
import types

import pytest

import _office as off
import _runtime as rt
import excel
import fake_office
import office
import ppt
import word
from fake_excel import FakeExcelWorkbook, FakeXlApp
from fake_office import make_win32com
from fake_ppt import FakePresentation
from fake_word import FakeWordDocument


@pytest.fixture(autouse=True)
def env(monkeypatch, tmp_path):
    monkeypatch.setenv("OFFICE_CLI_DATA_DIR", str(tmp_path / "d"))
    monkeypatch.setenv("OFFICE_CLI_CACHE_DIR", str(tmp_path / "c"))


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


def _points(kind):
    return [s["meta_file"] for s in off.list_snapshots(kind)]


@pytest.fixture
def xl(monkeypatch):
    app = FakeXlApp()
    wb = FakeExcelWorkbook(app, "b.xlsx", sheets=("Sheet1", "S2"))
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
    doc = FakeWordDocument("iPhone and iPhone\r끝")
    _com(monkeypatch)
    _active(monkeypatch, object(), doc)
    return doc


# ---------------------------------------------------------------------------
# F1: 빈 찾을 글은 모든 글자 사이에 바꿀 글을 끼워 넣음 -> 바꾸기 전에 STOP_RETRY
# ---------------------------------------------------------------------------

EMPTY_FIND = [
    ["--find", "", "--replace", "X"],
    ["--map", '{"": "X"}'],
    ["--map", '{"a": "b", "": "X"}'],
]


@pytest.mark.parametrize("extra", [[], ["--all-sheets"]])
@pytest.mark.parametrize("argv", EMPTY_FIND)
def test_f1_excel_empty_find_is_refused(xl, capsys, argv, extra):
    app, wb = xl
    ws, ws2 = wb.Worksheets(1), wb.Worksheets(2)
    ws.Range("A1").Value = "abc"
    ws.Range("B2").Value = "hello"
    ws2.Range("A1").Value = "xyz"
    code, out = _run(excel, capsys, "replace", "--range", "A1:C3", *argv, *extra)
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:"), out
    assert "excel.py write" in out["message"], out
    assert ws.Range("A1").Value == "abc" and ws.Range("B2").Value == "hello" and ws2.Range("A1").Value == "xyz"
    assert _points("excel") == []


@pytest.mark.parametrize("argv", EMPTY_FIND)
def test_f1_ppt_empty_find_is_refused(pp, capsys, argv):
    tb = pp.Slides(1).Shapes.AddTextbox(1, 0, 0, 100, 40)
    tb.TextFrame.TextRange.Text = "abc"
    code, out = _run(ppt, capsys, "replace", *argv)
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:"), out
    assert "ppt.py write" in out["message"], out
    assert tb.TextFrame.TextRange.Text == "abc"
    assert _points("ppt") == []


@pytest.mark.parametrize("argv", EMPTY_FIND)
def test_f1_word_empty_find_is_refused(wd, capsys, argv):
    before = wd.text
    code, out = _run(word, capsys, "replace", *argv)
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:"), out
    assert "word.py write" in out["message"], out
    assert wd.text == before
    assert _points("word") == []


# ---------------------------------------------------------------------------
# F2: 새 파일의 기본 시트가 여럿(Office 2007/2010: Sheet1~3)이어도 기본 시트 이름으로 받음
# ---------------------------------------------------------------------------

def test_f2_transfer_new_file_with_three_default_sheets(monkeypatch, tmp_path):
    _com(monkeypatch)

    def add3(self, Template=None):
        wb = fake_office.FakeWorkbook("", sheets=("Sheet1", "Sheet2", "Sheet3"))
        self.items.append(wb)
        return wb

    monkeypatch.setattr(fake_office.FakeWorkbooks, "Add", add3)
    monkeypatch.setattr(off, "_new_instance", lambda kind: fake_office.FakeExcelApp())
    for sheet in ("Sheet1", "sheet3"):
        target = tmp_path / f"new_{sheet}.xlsx"
        args = types.SimpleNamespace(sheet=sheet, to_target=str(target), overwrite=False, no_focus=True)
        # 가짜 시트는 쓰기 단계의 Range가 없어 거기서 멈춤 (그 전 '이미 있음'으로 막히지 않으면 됨)
        try:
            office._to_excel(args, [["h"], ["v"]], "Word 표")
        except rt.InputError as e:
            pytest.fail(f"새 파일의 기본 시트 이름 {sheet}이(가) 막힘: {e}")
        except AttributeError as e:
            assert "Range" in str(e)


# ---------------------------------------------------------------------------
# F3: 대소문자만 다른 바꾸기에서 이미 바꿀 글인 곳은 바꾼 곳으로 세지 않음
# ---------------------------------------------------------------------------

def test_f3_ppt_case_only_already_replaced_is_noop(pp, capsys):
    tb = pp.Slides(1).Shapes.AddTextbox(1, 0, 0, 100, 40)
    tb.TextFrame.TextRange.Text = "iPhone and iPhone"
    code, out = _run(ppt, capsys, "replace", "--find", "iphone", "--replace", "iPhone")
    assert code == 0 and out["replaced"] == 0 and out["unchanged_same_text"] == {"iphone": 2}, out
    assert out["slides"] == [] and "0곳" in out["preview"], out
    assert tb.TextFrame.TextRange.Text == "iPhone and iPhone"
    assert _points("ppt") == []


def test_f3_ppt_counts_only_changed_occurrences(pp, capsys):
    tb = pp.Slides(1).Shapes.AddTextbox(1, 0, 0, 100, 40)
    tb.TextFrame.TextRange.Text = "iPhone and IPHONE"
    code, out = _run(ppt, capsys, "replace", "--find", "iPhone", "--replace", "iPhone")
    assert code == 0 and out["replaced"] == 1 and out["unchanged_same_text"] == {"iPhone": 1}, out
    assert out["agent_directive"].startswith("VERIFIED:"), out
    assert tb.TextFrame.TextRange.Text == "iPhone and iPhone"


def test_f3_ppt_case_only_mixed(pp, capsys):
    tb = pp.Slides(1).Shapes.AddTextbox(1, 0, 0, 100, 40)
    tb.TextFrame.TextRange.Text = "iphone, iPhone, IPHONE"
    code, out = _run(ppt, capsys, "replace", "--find", "iphone", "--replace", "iPhone")
    assert code == 0 and out["replaced"] == 2 and out["unchanged_same_text"] == {"iphone": 1}, out
    assert tb.TextFrame.TextRange.Text == "iPhone, iPhone, iPhone"


def test_f3_ppt_case_only_readback_checks_exact_case(pp, capsys, monkeypatch):
    tb = pp.Slides(1).Shapes.AddTextbox(1, 0, 0, 100, 40)
    tb.TextFrame.TextRange.Text = "IPHONE"
    real = ppt.replace_in_range

    def lower_only(tr, find, rep, match_case, fmt):   # 바꿨다고 보고하지만 대소문자가 그대로 남은 상황
        tr.Text = str(tr.Text).lower()
        return 1, 0

    monkeypatch.setattr(ppt, "replace_in_range", lower_only)
    code, out = _run(ppt, capsys, "replace", "--find", "iphone", "--replace", "iPhone")
    monkeypatch.setattr(ppt, "replace_in_range", real)
    assert code == 0 and out["agent_directive"].startswith("UNVERIFIED:"), out


def test_f3_word_case_only_already_replaced_is_noop(wd, capsys):
    code, out = _run(word, capsys, "replace", "--find", "iphone", "--replace", "iPhone")
    assert code == 0 and out["replaced"] == 0 and out["unchanged_same_text"] == {"iphone": 2}, out
    assert "0곳" in out["preview"], out
    assert _points("word") == [] and wd.Saved is True


def test_f3_word_counts_only_changed_occurrences(monkeypatch, capsys):
    doc = FakeWordDocument("iphone, iPhone, IPHONE")
    _com(monkeypatch)
    _active(monkeypatch, object(), doc)
    code, out = _run(word, capsys, "replace", "--find", "iphone", "--replace", "iPhone")
    assert code == 0 and out["replaced"] == 2 and out["unchanged_same_text"] == {"iphone": 1}, out
    assert out["agent_directive"].startswith("VERIFIED:"), out
    assert doc.text.rstrip("\r") == "iPhone, iPhone, iPhone"


def test_f3_word_same_text_with_format_still_changes(wd, capsys):
    code, out = _run(word, capsys, "replace", "--find", "iphone", "--replace", "iPhone", "--bold")
    assert code == 0 and out["replaced"] == 2, out
