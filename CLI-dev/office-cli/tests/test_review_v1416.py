# -*- coding: utf-8 -*-
"""1.4.16 (7차 점검): ppt replace 대소문자 통일, 새 파일 기본 시트로 transfer, unhide/clear(전체·서식)는 바뀐 것으로 봄."""
import json
import os
import sys
import types

import pytest

HERE = os.path.dirname(os.path.abspath(__file__))
TOOL = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(HERE))), "CLI", "office-cli")
sys.path[:0] = [TOOL, HERE]
import _office as off  # noqa: E402
import _runtime as rt  # noqa: E402
import excel  # noqa: E402
import office  # noqa: E402
import ppt  # noqa: E402
from fake_excel import FakeExcelWorkbook, FakeXlApp  # noqa: E402
from fake_office import make_win32com  # noqa: E402
from fake_ppt import FakePresentation  # noqa: E402


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


@pytest.fixture
def pp(monkeypatch):
    prs = FakePresentation(3)
    app = types.SimpleNamespace(ActiveWindow=types.SimpleNamespace(View=types.SimpleNamespace(GotoSlide=lambda n: None),
                                                                   Presentation=prs))
    _com(monkeypatch)
    monkeypatch.setattr(off, "get_active", lambda kind: {"name": prs.Name, "path": prs.FullName, "app": app, "doc": prs})
    return prs


@pytest.fixture
def xl(monkeypatch):
    app = FakeXlApp()
    wb = FakeExcelWorkbook(app, "b.xlsx", sheets=("Sheet1", "S2"))
    _com(monkeypatch)
    monkeypatch.setattr(off, "get_active", lambda kind: {"name": wb.Name, "path": wb.FullName, "app": app, "doc": wb})
    return app, wb


def test_ppt_case_normalizing_replace_really_replaces(pp, capsys):
    tb = pp.Slides(1).Shapes.AddTextbox(1, 0, 0, 100, 40)
    tb.TextFrame.TextRange.Text = "IPHONE and iphone"
    code, out = _run(ppt, capsys, "replace", "--find", "iPhone", "--replace", "iPhone")
    assert code == 0 and out["replaced"] == 2, out
    assert tb.TextFrame.TextRange.Text == "iPhone and iPhone"


def test_ppt_same_text_same_case_is_still_noop(pp, capsys):
    tb = pp.Slides(1).Shapes.AddTextbox(1, 0, 0, 100, 40)
    tb.TextFrame.TextRange.Text = "iPhone and iPhone"
    code, out = _run(ppt, capsys, "replace", "--find", "iPhone", "--replace", "iPhone")
    assert code == 0 and out["replaced"] == 0 and out["unchanged_same_text"] == {"iPhone": 2}, out


def test_transfer_into_new_file_with_default_sheet_name(monkeypatch, tmp_path):
    from fake_office import FakeExcelApp
    _com(monkeypatch)
    monkeypatch.setattr(off, "_new_instance", lambda kind: FakeExcelApp())
    target = tmp_path / "new.xlsx"
    args = types.SimpleNamespace(sheet="Sheet1", to_target=str(target), overwrite=False, no_focus=True)
    # 기본 시트 이름과 같아도 '이미 있음'으로 막지 않고 그 빈 시트에 받음 (가짜 시트는 쓰기 단계의 Range가 없어 거기서 멈춤)
    try:
        office._to_excel(args, [["h"], ["v"]], "Word 표")
    except rt.InputError as e:
        pytest.fail(f"새 파일의 기본 시트 이름이 막힘: {e}")
    except AttributeError as e:
        assert "Range" in str(e)


@pytest.mark.parametrize("argv", [
    ["range", "--action", "unhide", "--range", "B:C"],
    ["range", "--action", "clear", "--range", "D10:E12"],
])
def test_unhide_and_full_clear_are_treated_as_changes(xl, capsys, argv):
    code, out = _run(excel, capsys, *argv)
    assert code == 0 and out.get("unchanged") is not True, out
    assert "바뀐 것 없음" not in out["preview"], out
