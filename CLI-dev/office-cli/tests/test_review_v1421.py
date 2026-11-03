"""1.4.21 (12차 점검) 회귀 테스트.

F2: 시트를 주지 않으면 find_chart가 모든 시트에서 이름/제목이 맞는 첫 차트를 골랐고, chart 생성의 next_commands에 --sheet가
    없어 다른 시트의 같은 제목 차트('매출')가 옮겨지거나(VERIFIED) chart --action delete --name이 첫 차트를 지웠음.
F1: 이름이 같은 열린 문서가 여럿일 때 match_open이 transfer에 없는 --target을 안내했음.
"""
import json
import shlex

import pytest

import _office as off
import excel
import office
import test_excel
import test_review_v1420
from fake_ppt import FakePresentation

xl = test_excel.xl                  # 가짜 Excel 통합문서 (Sheet1, 요약)
env = test_review_v1420.env         # 엑셀 차트 + 열린 워드 (transfer)


def _run(mod, capsys, *argv):
    with pytest.raises(SystemExit) as ei:
        mod.main(list(argv))
    out = [ln for ln in capsys.readouterr().out.splitlines() if ln.strip()]
    assert len(out) == 1, out
    return ei.value.code, json.loads(out[0])


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


def test_same_title_on_two_sheets_next_command_targets_new_chart(xl, capsys):
    app, wb = xl
    _fill(wb)
    _run(excel, capsys, "chart", "--sheet", "Sheet1", "--range", "A1:B3", "--title", "매출")
    code, out = _run(excel, capsys, "chart", "--sheet", "요약", "--range", "A1:B3", "--title", "매출")
    assert code == 0
    argv = shlex.split(out["next_commands"][0]["command"])
    assert argv[argv.index("--sheet") + 1] == "요약"
    ws, obj = excel.find_chart(_T(wb), argv[argv.index("--chart") + 1], argv[argv.index("--sheet") + 1])
    assert ws.Name == "요약"


def test_same_name_without_sheet_is_ambiguous(xl, capsys):
    app, wb = xl
    _fill(wb)
    _run(excel, capsys, "chart", "--sheet", "Sheet1", "--range", "A1:B3", "--title", "매출")
    _run(excel, capsys, "chart", "--sheet", "요약", "--range", "A1:B3", "--title", "매출")
    with pytest.raises(Exception) as ei:
        excel.find_chart(_T(wb), "매출", None)
    assert "[후보]" in str(ei.value) and "Sheet1!매출" in str(ei.value) and "요약!매출" in str(ei.value)
    ws, _ = excel.find_chart(_T(wb), "요약!매출", None)             # 후보 형식을 그대로 받음
    assert ws.Name == "요약"


def test_delete_by_name_does_not_pick_first_sheet(xl, capsys):
    app, wb = xl
    _fill(wb)
    _run(excel, capsys, "chart", "--sheet", "Sheet1", "--range", "A1:B3", "--title", "매출")
    _run(excel, capsys, "chart", "--sheet", "요약", "--range", "A1:B3", "--title", "매출")
    code, out = _run(excel, capsys, "chart", "--action", "delete", "--name", "매출")
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:") and "[후보]" in out["message"]
    assert _names(wb) == {"Sheet1": ["매출"], "요약": ["매출"]}
    code, out = _run(excel, capsys, "chart", "--action", "delete", "--name", "요약!매출")
    assert code == 0 and _names(wb) == {"Sheet1": ["매출"], "요약": []}, out


def test_exact_name_beats_other_chart_title(xl, capsys):
    app, wb = xl
    _fill(wb)
    # 앞 시트 차트의 제목이 뒤 시트 차트의 이름과 같아도 이름이 같은 차트를 고름 (이전: 앞 시트의 제목 차트)
    _run(excel, capsys, "chart", "--sheet", "Sheet1", "--range", "A1:B3", "--title", "매출", "--name", "첫차트")
    _run(excel, capsys, "chart", "--sheet", "요약", "--range", "A1:B3", "--title", "합계", "--name", "매출")
    ws, obj = excel.find_chart(_T(wb), "매출", None)
    assert ws.Name == "요약" and obj.Name == "매출"
    code, out = _run(excel, capsys, "chart", "--action", "delete", "--name", "매출")
    assert code == 0 and _names(wb) == {"Sheet1": ["첫차트"], "요약": []}, out


CH = ("transfer", "--from", "excel", "--chart", "차트")
TWO_SAME = [{"name": "a.docx", "path": "C:\\x\\a.docx", "doc": object(), "app": None},
            {"name": "a.docx", "path": "C:\\y\\a.docx", "doc": object(), "app": None}]


def test_ambiguous_open_documents_name_transfer_options(env, capsys, monkeypatch):
    env["prs"] = FakePresentation(1)
    monkeypatch.setattr(off, "list_open", lambda kind: TWO_SAME if kind == "word" else [])
    # 앱이 둘이면 넣기 전에 받는 문서를 모두 찾음(_probe_target). 보내는 쪽은 open_target(option=--from-target)
    for argv, opt in (((*CH, "--to", "word,ppt", "--to-target", "a.docx"), "--to-target"),
                      (("transfer", "--from", "word", "--from-target", "a.docx", "--to", "excel"), "--from-target")):
        code, out = _run(office, capsys, *argv)
        assert code == 1 and "[후보]" in out["message"] and env["inserted"] == [], out
        assert f"{opt}에 전체 경로" in out["message"], out


def test_ambiguous_focus_names_name_option(monkeypatch, capsys):
    monkeypatch.setattr(off, "list_open", lambda kind: TWO_SAME)
    code, out = _run(office, capsys, "focus", "--app", "word", "--name", "a.docx")
    assert code == 1 and "--name에 전체 경로" in out["message"], out
