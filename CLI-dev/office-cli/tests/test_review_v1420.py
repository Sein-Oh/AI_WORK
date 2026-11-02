"""1.4.20 (11차 점검) 회귀 테스트: transfer --to word,ppt에서 한 앱만 열려 있거나 한 앱만 실패할 때.

이전: 워드에 먼저 넣은 뒤 PowerPoint가 열려 있지 않으면 'STOP_RETRY: 실행하지 않았습니다 ... --target'으로 끝나
(transfer에는 --target이 없음) 다시 실행하면 워드에 두 번 들어갔음.
"""
import json
import shlex

import pytest

import _office as off
import _runtime as rt
import excel
import office
import ppt
from fake_excel import FakeExcelWorkbook, FakeXlApp
from fake_office import make_win32com
from fake_ppt import FakePresentation
from fake_word import FakeWordDocument


def _run(capsys, *argv):
    with pytest.raises(SystemExit) as ei:
        office.main(list(argv))
    out = [ln for ln in capsys.readouterr().out.splitlines() if ln.strip()]
    assert len(out) == 1, out
    return ei.value.code, json.loads(out[0])


def _recommended_argv(message):
    """message의 '[추천] python .../office.py <인자>'에서 인자만."""
    cmd = message.split("[추천] ", 1)[1]
    parts = shlex.split(cmd)
    assert parts[0] == "python" and parts[1].endswith("office.py")
    return parts[2:]


@pytest.fixture
def env(monkeypatch, tmp_path):
    """엑셀 차트/범위 하나, 워드 문서 열림, PowerPoint는 holder['prs']가 None이면 열린 문서 없음."""
    fake, _ = make_win32com()
    monkeypatch.setattr(off, "win32com", fake)
    monkeypatch.setattr(off, "pythoncom", None)
    monkeypatch.setattr(off, "list_open", lambda kind: [])
    monkeypatch.setattr(off, "find_apps", lambda kind: [])
    app = FakeXlApp()
    wb = FakeExcelWorkbook(app)
    wd = FakeWordDocument("제목\r본문")
    holder = {"prs": None, "word": wd, "inserted": []}
    pp_app = type("A", (), {})()
    pp_app.ActiveWindow = type("W", (), {})()

    real_get_active = off.get_active

    def active(kind):
        if kind == "excel":
            return {"name": wb.Name, "path": wb.FullName, "app": app, "doc": wb}
        if kind == "word" and holder["word"] is not None:
            return {"name": wd.Name, "path": wd.FullName, "app": object(), "doc": wd}
        if kind == "ppt" and holder["prs"] is not None:
            return {"name": holder["prs"].Name, "path": holder["prs"].FullName, "app": pp_app, "doc": holder["prs"]}
        return real_get_active(kind)                # find_apps가 빈 목록: 실제 "열린 문서 없음" 오류
    monkeypatch.setattr(off, "get_active", active)
    chart = type("C", (), {"Name": "차트"})()
    monkeypatch.setattr(excel, "find_chart", lambda t, name, sheet=None: (None, chart))
    monkeypatch.setattr(excel, "_chart_title", lambda obj: "차트")
    png = tmp_path / "c.png"
    png.write_bytes(office.PNG_SIGNATURE + b"x" * 50)
    monkeypatch.setattr(office, "image_by_export", lambda obj, stem: str(png))
    monkeypatch.setattr(office, "image_by_redraw", lambda obj, stem: str(png))
    monkeypatch.setattr(office, "insert_picture_word",
                        lambda tgt, img, anchor, occurrence=None: holder["inserted"].append("word") or {"app": "Word"})
    monkeypatch.setattr(office, "insert_picture_ppt",
                        lambda tgt, img, slide, title: holder["inserted"].append("ppt") or {"app": "PowerPoint"})
    monkeypatch.setattr(rt, "load_config", lambda: {"transfer_method": {"word": ["export"], "ppt": ["export"]}})
    monkeypatch.setattr(office, "_excel_range_rows", lambda args: ([["항목", "값"], ["A", 1]], "Sheet1"))
    return holder


CHART = ("transfer", "--from", "excel", "--chart", "차트", "--to", "word,ppt")
RANGE = ("transfer", "--from", "excel", "--range", "A1:B2", "--to", "word,ppt")


def _check_target_not_open(out):
    assert out["status"] == "error" and out["error_type"] == "target_not_open"
    assert "아무것도 바꾸지 않았습니다" in out["message"] and "PowerPoint" in out["message"]
    assert "--target " not in out["message"] + out["agent_directive"]          # transfer에 없는 옵션
    assert out["agent_directive"].startswith("STOP_RETRY:") and "실행하지 않았습니다" not in out["agent_directive"]
    argv = _recommended_argv(out["message"])
    assert argv[argv.index("--to") + 1] == "word"
    return argv


def test_chart_word_only_open_changes_nothing_and_recommends_word(env, capsys):
    code, out = _run(capsys, *CHART)
    assert code == 1 and env["inserted"] == []
    argv = _check_target_not_open(out)
    code, out = _run(capsys, *argv)
    assert code == 0 and env["inserted"] == ["word"] and list(out["results"]) == ["word"], out


def test_range_word_only_open_changes_nothing_and_recommends_word(env, capsys):
    before = env["word"].text
    code, out = _run(capsys, *RANGE)
    assert code == 1 and env["word"].text == before
    argv = _check_target_not_open(out)
    code, out = _run(capsys, *argv)
    assert code == 0 and list(out["results"]) == ["word"] and env["word"].text != before, out


def test_no_receiver_open_names_to_target_for_all_apps(env, capsys):
    env["word"] = None
    code, out = _run(capsys, *CHART)
    assert code == 1 and env["inserted"] == [] and out["error_type"] == "needs_user_decision"
    assert "--to-target" in out["agent_directive"] and "모든 앱" in out["agent_directive"]
    assert "--target " not in out["agent_directive"]


def test_single_missing_target_and_source_name_transfer_options(env, capsys):
    code, out = _run(capsys, "transfer", "--from", "excel", "--range", "A1:B2", "--to", "ppt")
    assert code == 1 and '--to-target "' in out["agent_directive"] and "--target " not in out["agent_directive"]
    env["word"] = None
    code, out = _run(capsys, "transfer", "--from", "word", "--to", "excel")
    assert code == 1 and '--from-target "' in out["agent_directive"] and "--target " not in out["agent_directive"]


def test_invalid_slide_is_checked_before_word_changes(env, capsys):
    env["prs"] = FakePresentation(1)
    before = env["word"].text
    code, out = _run(capsys, *RANGE, "--slide", "5")
    assert code == 1 and "슬라이드 5번이 없습니다" in out["message"] and env["word"].text == before
    code, out = _run(capsys, *CHART, "--slide", "5")
    assert code == 1 and env["inserted"] == []


def test_invalid_anchor_is_checked_before_anything(env, capsys):
    env["prs"] = FakePresentation(1)
    code, out = _run(capsys, *RANGE, "--anchor", "없는 문구")
    assert code == 1 and "찾을 수 없습니다" in out["message"] and env["prs"].Slides.Count == 1


def test_range_late_failure_is_unverified_with_lists(env, capsys, monkeypatch):
    env["prs"] = FakePresentation(1)

    def boom(self, rows):
        raise RuntimeError("PPT 표 만들기 실패")
    monkeypatch.setattr(ppt.SlideWriter, "table", boom)
    before = env["word"].text
    code, out = _run(capsys, *RANGE)
    assert code == 0 and out["status"] == "ok" and out["verified"] is False, out
    assert out["succeeded"] == ["word"] and out["failed"] == ["ppt"] and "PPT 표 만들기 실패" in out["failures"]["ppt"]
    assert out["agent_directive"].startswith("UNVERIFIED:") and "한 번 더" in out["agent_directive"]
    assert env["word"].text != before


def test_chart_no_fallback_failure_after_word_is_unverified(env, capsys, monkeypatch):
    env["prs"] = FakePresentation(1)

    def ppt_fail(tgt, img, slide, title):
        raise rt.InputError("PowerPoint 대상 문제", no_fallback=True)
    monkeypatch.setattr(office, "insert_picture_ppt", ppt_fail)
    code, out = _run(capsys, *CHART)
    assert code == 0 and out["verified"] is False and env["inserted"] == ["word"], out
    assert out["succeeded"] == ["word"] and out["failed"] == ["ppt"]
    assert out["agent_directive"].startswith("UNVERIFIED:") and "--to ppt" in out["agent_directive"]


def test_undo_points_after_partial_and_preflight_failure(env, capsys, monkeypatch):
    # 실제 스냅샷 저장: 미리 확인에서 멈추면 지점 없음, 워드만 바뀌면 워드 지점만 남음
    code, out = _run(capsys, *RANGE)
    assert code == 1 and off.list_snapshots() == []
    env["prs"] = FakePresentation(1)

    def boom(self, rows):
        raise RuntimeError("PPT 표 만들기 실패")
    monkeypatch.setattr(ppt.SlideWriter, "table", boom)
    code, out = _run(capsys, *RANGE)
    assert code == 0 and out["failed"] == ["ppt"]
    assert [s["app"] for s in off.list_snapshots()] == ["word"]


def test_chart_next_commands_use_open_receivers(monkeypatch):
    monkeypatch.setattr(off, "list_open", lambda kind: [{"name": "a"}] if kind == "word" else [])
    assert excel._open_receivers() == ["word"]
    monkeypatch.setattr(off, "list_open", lambda kind: [])
    assert excel._open_receivers() == ["word", "ppt"]

    def boom(kind):
        raise RuntimeError("COM 없음")
    monkeypatch.setattr(off, "list_open", boom)
    assert excel._open_receivers() == ["word", "ppt"]
