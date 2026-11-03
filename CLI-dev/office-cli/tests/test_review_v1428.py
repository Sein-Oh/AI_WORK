"""1.4.28 (19차 점검) 회귀 테스트.

M1/M2/L1/L2/L4: 1.4.27의 오류 로그 가리기는 message의 표식([후보], [추천] 등)에만 기대어, 표식이 없는 문서 글이 로그에 남았음
       (excel 덮어쓰기·병합·옮기기의 'A1=<칸 값>', ppt 기존 제목·노트, '(있는 시트: ...)', word comment --file의 '후보:'로 바뀐 후보,
       '사용 가능한 메일 폴더:'처럼 종류에 공백이 있는 목록, 가린 값의 잘린 앞부분, 가리지 않은 traceback).
       오류를 만드는 곳이 문서 글을 뺀 log_message를 주고, 표식 자르기는 대비로 남깁니다.
L3: 1~2자 가린 값(--value 1)이 관계없는 글('Sheet1'의 1)까지 가렸음. 3자 이상만, 짧은 값은 따옴표 안만 가립니다.
L5: undo가 되돌릴 수 없는 작업만 있을 때 needs_user_decision이었고, undo --list는 그런 작업이 20개 넘게 앞에 있으면
    되돌릴 수 있는 지점을 목록에서 뺐음.
"""
import json
import os
import sys

import pytest

import _office as off
import _runtime as rt
import excel
import office
import ppt
import test_excel
import test_ppt
import test_word
import word
from fake_excel import FakeChartObject, FakeExcelWorkbook, FakeXlApp
from fake_office import make_win32com
from fake_word import FakeWordDocument
from test_word_review import PageDoc

pp = test_ppt.pp
wd = test_word.wd
xl = test_excel.xl

# 문서에서 가져온 글을 나타내는 표지 (칸 값, 제목, 노트, 시트·문서·파일 이름, 단락). 오류 로그에 하나도 남으면 안 됨
S = "ZQX"


def _run(mod, capsys, monkeypatch, tool, *argv):
    monkeypatch.setattr(rt, "CURRENT_ARGV", None)
    monkeypatch.setattr(sys, "argv", [f"{tool}.py"] + list(argv))
    with pytest.raises(SystemExit) as ei:
        mod.main(list(argv))
    out = [ln for ln in capsys.readouterr().out.splitlines() if ln.strip()]
    assert len(out) == 1, out
    return ei.value.code, json.loads(out[0])


def _log(d):
    p = os.path.join(d["data"], "logs", "errors.txt")
    return open(p, encoding="utf-8").read() if os.path.exists(p) else ""


def _assert_clean(out, isolated_dirs, marker=S):
    """stdout message에는 문서 글이 그대로 있고(사용자에게 보여 줄 것), 오류 로그에는 없어야 함."""
    assert out["status"] == "error" and marker in out["message"], out
    log = _log(isolated_dirs)
    assert log and marker not in log, log
    return log


# 검토자 재현 (M1, M2, L1, L2, L4) --------------------------------------------------------------------------------

def test_ppt_existing_title_not_logged(pp, capsys, isolated_dirs, monkeypatch):
    pp["prs"].Slides(1).Shapes.Title.TextFrame.TextRange.Text = "인수합병 비밀 일정"
    code, out = _run(ppt, capsys, monkeypatch, "ppt", "write", "--slide", "1", "--title", "새 제목")
    assert code == 1 and out["error_type"] == "needs_user_decision" and "비밀 일정" in out["message"], out
    log = _log(isolated_dirs)
    assert "비밀 일정" not in log and "슬라이드 1에 이미 제목이 있습니다" in log, log


def test_ppt_existing_notes_not_logged(pp, capsys, isolated_dirs, monkeypatch):
    pp["prs"].Slides(1).NotesPage.body.TextFrame.TextRange.Text = "기밀 발표 메모 내용"
    code, out = _run(ppt, capsys, monkeypatch, "ppt", "notes", "--slide", "1", "--text", "새 노트")
    assert code == 1 and "기밀 발표" in out["message"], out
    assert "기밀 발표" not in _log(isolated_dirs)


def test_word_long_find_no_prefix_preview(wd, capsys, isolated_dirs, monkeypatch):
    longv = "고객사 홍길동 계약 해지 예정 문서 " * 20
    code, out = _run(word, capsys, monkeypatch, "word", "replace", "--find", longv, "--replace", "x")
    assert code == 1 and "홍길동" not in out["message"] and "1번째 찾을 문구" in out["message"], out
    assert "홍길동" not in _log(isolated_dirs)


def test_short_value_does_not_overmask():
    m = rt.log_message("시트 'Sheet1'의 A1:B10 범위가 비어 있습니다", ["write", "--value", "1"], "excel")
    assert "Sheet1" in m and "A1:B10" in m, m
    # 따옴표 안의 짧은 값은 가림
    assert rt.log_message("값 '1'을 넣지 못했습니다", ["write", "--value", "1"], "excel") == "값 '<1자 생략>'을 넣지 못했습니다"


def test_excel_overwrite_cell_values_not_logged(xl, capsys, isolated_dirs, monkeypatch):
    app, wb = xl
    wb.Worksheets("Sheet1").Range("A1").Value = "홍길동 연봉 9800만원"
    code, out = _run(excel, capsys, monkeypatch, "excel", "write", "--range", "A1", "--value", "x")
    assert code == 1 and "9800" in out["message"], out
    log = _log(isolated_dirs)
    assert "9800" not in log and "A1" in log, log


def test_excel_merge_values_not_logged(xl, capsys, isolated_dirs, monkeypatch):
    app, wb = xl
    ws = wb.Worksheets("Sheet1")
    ws.Range("A1").Value = "a"
    ws.Range("B1").Value = "기밀 매출 12억"
    code, out = _run(excel, capsys, monkeypatch, "excel", "range", "--action", "merge", "--range", "A1:C1")
    assert code == 1 and "12억" in out["message"], out
    assert "12억" not in _log(isolated_dirs)


def test_word_comment_csv_candidates_not_logged(wd, capsys, isolated_dirs, monkeypatch, tmp_path):
    wd["doc"] = FakeWordDocument("계획 홍길동 해고 예정\r계획 김철수 해고 예정\r일정")
    f = tmp_path / "memo.csv"
    f.write_text("기준 글,메모 내용\n계획,확인\n일정,확인2\n", encoding="utf-8")
    code, out = _run(word, capsys, monkeypatch, "word", "comment", "--action", "add", "--file", str(f))
    assert code == 1 and "홍길동" in out["message"], out
    log = _log(isolated_dirs)
    assert "홍길동" not in log and "김철수" not in log and "줄: 1" in log, log


# 표식 대비 (c), 앞부분 (d), traceback (e) -------------------------------------------------------------------------

def test_cut_markers_kinds_with_spaces_and_unbracketed():
    for msg in ("메일 폴더 'x'을(를) 찾을 수 없습니다. 사용 가능한 메일 폴더: ZQX받은편지함",
                "문서 'a'을(를) 찾을 수 없습니다. 사용 가능한 열린 Word 문서: ZQX.docx",
                "1번째 줄: 고를 수 없습니다. 후보: 1: ZQX단락",
                "표기만 다릅니다. 추천: ZQX시트"):
        m = rt.log_message(msg, [], "word")
        assert S not in m and "생략>" in m, m


def test_truncated_prefix_of_hidden_value_masked():
    v = "고객사 홍길동 계약 해지 예정 문서"
    m = rt.log_message(f"찾을 문구가 깁니다: '{v[:12]}...'", ["replace", "--find", v], "word")
    assert "홍길동" not in m and "<12자 생략>" in m, m
    # 10자보다 짧은 앞부분은 가리지 않음 (관계없는 글을 가리지 않게)
    assert "고객사" in rt.log_message("고객사 목록", ["replace", "--find", v], "word")


def test_traceback_masked_without_marker_cut(isolated_dirs):
    trace = ('Traceback (most recent call last):\n  File "word.py", line 1, in f\nKeyError: \'비밀계획서 초안\'\n'
             "  [후보] 줄은 코드 글이라 자르지 않음\n")
    rt.write_error_log("word", ["write", "--text", "비밀계획서 초안"], {"error_type": "internal", "message": "KeyError"}, trace)
    log = _log(isolated_dirs)
    assert "비밀계획서" not in log and "KeyError: '<8자 생략>'" in log and "[후보] 줄은 코드 글" in log, log


def test_internal_error_traceback_masked_through_run_cli(isolated_dirs, monkeypatch, capsys):
    monkeypatch.setattr(rt, "CURRENT_ARGV", ["write", "--text", "비밀계획서 초안"])

    def boom():
        raise KeyError("비밀계획서 초안")
    with pytest.raises(SystemExit):
        rt.run_cli("word", "write", boom)
    out = json.loads(capsys.readouterr().out)
    assert "비밀계획서" in out["message"]                                  # stdout은 그대로
    log = _log(isolated_dirs)
    assert "traceback:" in log and "비밀계획서" not in log, log


def test_result_dict_log_message_key_popped(isolated_dirs, monkeypatch, capsys):
    monkeypatch.setattr(rt, "CURRENT_ARGV", ["send"])
    with pytest.raises(SystemExit):
        rt.run_cli("outlook", "send", lambda: {"status": "error", "tool": "outlook", "command": "send", "error_type": "input",
                                               "message": "수신자 확인 실패: ZQX홍길동", "agent_directive": "STOP_RETRY: x",
                                               rt.LOG_MESSAGE_KEY: "수신자 1개 확인 실패"})
    out = json.loads(capsys.readouterr().out)
    assert rt.LOG_MESSAGE_KEY not in out and S in out["message"], out
    log = _log(isolated_dirs)
    assert S not in log and "수신자 1개 확인 실패" in log, log


def test_error_classes_carry_log_message():
    e = rt.ambiguous("'x'이(가) 2곳에 있습니다.", ["ZQX a", "ZQX b"])
    assert S in str(e) and S not in e.log_message and "[후보 2개 생략]" in e.log_message
    e = rt.UserDecisionError("시트가 없습니다.", "묻기", candidates=["ZQX"])
    assert S in str(e) and S not in e.log_message
    for cands in (["ZQX가", "ZQX나"], ["zqx 시트"], ["ZQX"]):
        e = rt.not_found("시트", "ZQX 시트" if cands == ["zqx 시트"] else "없음", cands)
        assert S.lower() in str(e).lower()
        assert S.lower() not in e.log_message.lower().replace("'zqx 시트'", ""), e.log_message
    assert e.log_message.endswith("[사용 가능한 이름 1개 생략]")


# 가드: 문서 글이 오류 로그에 남지 않음 (needs_user_decision / not_found / ambiguous 경로마다) ---------------------

@pytest.fixture
def xs(monkeypatch):
    """시트 이름과 칸 값에 표지가 든 가짜 Excel."""
    app = FakeXlApp()
    wb = FakeExcelWorkbook(app, "보고.xlsx", sheets=("ZQX인사", "ZQX급여"))
    fake, _ = make_win32com()
    monkeypatch.setattr(off, "win32com", fake)
    monkeypatch.setattr(off, "pythoncom", None)
    monkeypatch.setattr(off, "list_open", lambda kind: [])
    monkeypatch.setattr(off, "get_active", lambda kind: {"name": wb.Name, "path": wb.FullName, "app": app, "doc": wb})
    monkeypatch.setattr(off, "save_snapshot", lambda *a, **k: None)
    ws = wb.Worksheets("ZQX인사")
    ws.Range("A1").Value = "ZQX홍길동"
    ws.Range("B1").Value = "ZQX연봉"
    ws.Range("D1").Value = "ZQX주소"
    return app, wb


@pytest.mark.parametrize("argv", [
    ["write", "--range", "A1", "--value", "x"],                                # 덮어쓰기 칸 값
    ["range", "--action", "merge", "--range", "A1:C1"],                         # 병합으로 지워질 값
    ["range", "--action", "move", "--range", "A1", "--dest", "D1"],             # 옮길 자리의 값
    ["write", "--sheet", "없는시트", "--range", "A1", "--value", "x"],             # (있는 시트: ...)
    ["write", "--sheet", "zqx 인사", "--range", "A1", "--value", "x"],            # 표기만 다른 시트 [추천]
    ["sheet", "--action", "select", "--name", "없는시트"],                         # 사용 가능한 시트:
    ["sheet", "--action", "select", "--name", "인사"],                             # 비슷한 시트 [후보]
])
def test_guard_excel(xs, capsys, isolated_dirs, monkeypatch, argv):
    code, out = _run(excel, capsys, monkeypatch, "excel", *argv)
    assert code == 1, out
    _assert_clean(out, isolated_dirs)


def test_guard_excel_chart_names(xs, capsys, isolated_dirs, monkeypatch):
    app, wb = xs
    ws = wb.Worksheets("ZQX인사")
    ws.charts.append(FakeChartObject(ws, "ZQX매출차트"))
    code, out = _run(excel, capsys, monkeypatch, "excel", "chart", "--action", "delete", "--name", "없는차트")
    assert code == 1, out
    _assert_clean(out, isolated_dirs)


def test_guard_ppt_title_and_notes(pp, capsys, isolated_dirs, monkeypatch):
    s = pp["prs"].Slides(1)
    s.Shapes.Title.TextFrame.TextRange.Text = "ZQX 제목"
    s.NotesPage.body.TextFrame.TextRange.Text = "ZQX 노트"
    code, out = _run(ppt, capsys, monkeypatch, "ppt", "write", "--slide", "1", "--title", "새 제목")
    _assert_clean(out, isolated_dirs)
    code, out = _run(ppt, capsys, monkeypatch, "ppt", "notes", "--slide", "1", "--text", "새 노트")
    _assert_clean(out, isolated_dirs)


@pytest.mark.parametrize("doc,argv", [
    ("ZQX 보고서 초안 작성\r일정", ["write", "--anchor", "ZQY 보고서 초안", "--text", "본문"]),     # 비슷한 문구
    ("ZQX 보고서\r일정", ["write", "--anchor", "zqx보고서", "--text", "본문"]),                    # [추천] 앵커 (공백만 다름)
    ("계획 ZQX1\r계획 ZQX2\r일정", ["write", "--anchor", "계획", "--text", "본문"]),                 # [후보] 단락
])
def test_guard_word_anchor(wd, capsys, isolated_dirs, monkeypatch, doc, argv):
    wd["doc"] = FakeWordDocument(doc)
    code, out = _run(word, capsys, monkeypatch, "word", *argv)
    assert code == 1, out
    _assert_clean(out, isolated_dirs)


def test_guard_word_comment_csv(wd, capsys, isolated_dirs, monkeypatch, tmp_path):
    wd["doc"] = FakeWordDocument("계획 ZQX1\r계획 ZQX2\r일정")
    f = tmp_path / "memo.csv"
    f.write_text("기준 글,메모 내용,몇 번째\nZQX없는글,확인\n일정,확인2,ZQX\n", encoding="utf-8")
    code, out = _run(word, capsys, monkeypatch, "word", "comment", "--action", "add", "--file", str(f))
    _assert_clean(out, isolated_dirs)                     # 3번째 열이 숫자가 아님: CSV 칸 값
    f.write_text("기준 글,메모 내용\n계획,확인\nZQX없는글,확인2\n", encoding="utf-8")
    code, out = _run(word, capsys, monkeypatch, "word", "comment", "--action", "add", "--file", str(f))
    _assert_clean(out, isolated_dirs)                     # 기준 글과 후보 단락


@pytest.mark.parametrize("paras,section", [
    ([("1. ZQX 개요", 1), ("본문", 10), ("2. ZQX 결론", 1), ("본문", 10)], "부록"),                  # 사용 가능한 제목:
    ([("1. ZQX 개요", 1), ("본문", 10), ("2. ZQX 결론", 1), ("본문", 10)], "개요 계획"),             # 비슷한 제목 [후보]
    ([("ZQX 검증 방법", 2), ("a", 10), ("ZQX 검증 방법 2", 2), ("b", 10)], "검증 방법"),              # 제목 여러 개 [후보]
])
def test_guard_word_section(wd, capsys, isolated_dirs, monkeypatch, paras, section):
    wd["doc"] = PageDoc(paras)
    code, out = _run(word, capsys, monkeypatch, "word", "read", "--section", section)
    assert code == 1, out
    _assert_clean(out, isolated_dirs)


def test_guard_open_document_names(wd, capsys, isolated_dirs, monkeypatch):
    docs = [{"name": "ZQX인사평가.docx", "path": "C:\\docs\\ZQX인사평가.docx", "app": object(), "doc": object()},
            {"name": "ZQX급여.docx", "path": "C:\\docs\\ZQX급여.docx", "app": object(), "doc": object()}]
    monkeypatch.setattr(off, "list_open", lambda kind: docs)
    code, out = _run(word, capsys, monkeypatch, "word", "read", "--target", "없는문서.docx")
    assert code == 1, out
    _assert_clean(out, isolated_dirs)


def test_guard_input_file_names(wd, capsys, isolated_dirs, monkeypatch):
    for n in ("ZQX인사.md", "ZQX급여.md"):
        (isolated_dirs["data"] / n).write_text("x", encoding="utf-8")
    code, out = _run(word, capsys, monkeypatch, "word", "write", "--file", "없는파일.md")
    assert code == 1, out
    _assert_clean(out, isolated_dirs)


def test_guard_office_open_document_names(capsys, isolated_dirs, monkeypatch):
    docs = [{"name": "ZQX인사.xlsx", "path": "", "app": object(), "doc": object()}]
    monkeypatch.setattr(off, "list_open", lambda kind: docs)
    code, out = _run(office, capsys, monkeypatch, "office", "focus", "--app", "excel", "--name", "없는문서")
    assert code == 1, out
    _assert_clean(out, isolated_dirs)


# L5: undo ---------------------------------------------------------------------------------------------------------

def _barrier(i):
    return {"app": "word", "name": f"n{i}.docx", "barrier": True, "reason": "새로 만든 파일", "created_at": "20261103_101500",
            "by": "word write", "meta_file": f"{i}.json"}


def test_undo_plain_barrier_only_is_zero(capsys, monkeypatch):
    monkeypatch.setattr(off, "list_snapshots", lambda kind=None: [_barrier(0)])
    code, out = _run(office, capsys, monkeypatch, "office", "undo")
    assert code == 0 and out["status"] == "ok" and out["count"] == 0 and out["restored"] is False, out
    assert out["agent_directive"].startswith("VERIFIED: 0건"), out


def test_undo_plain_barrier_then_undoable_still_asks(capsys, monkeypatch):
    snaps = [_barrier(0), {"app": "word", "name": "old.docx", "created_at": "20261103_100000", "by": "word write",
                           "meta_file": "o.json", "original_path": ""}]
    monkeypatch.setattr(off, "list_snapshots", lambda kind=None: snaps)
    code, out = _run(office, capsys, monkeypatch, "office", "undo")
    assert code == 1 and out["error_type"] == "needs_user_decision", out


def test_undo_list_many_barriers_shows_undoable(capsys, monkeypatch):
    snaps = [_barrier(i) for i in range(25)]
    snaps.append({"app": "word", "name": "old.docx", "created_at": "20261103_100000", "by": "word write",
                  "meta_file": "o.json", "original_path": ""})
    monkeypatch.setattr(off, "list_snapshots", lambda kind=None: snaps)
    code, out = _run(office, capsys, monkeypatch, "office", "undo", "--list")
    assert code == 0 and out["count"] == 1 and out["shown"] == 1 and out["barriers"] == 25, out
    undoable = [it for it in out["snapshots"] if "command" in it]
    assert len(undoable) == 1 and undoable[0]["pick"] == 26 and "--pick 26" in undoable[0]["command"], out["snapshots"]
    assert out["truncated"] is True and len(out["snapshots"]) == 21, out       # 앞 20개 되돌릴 수 없는 작업 + 지점 1개
    assert out["shown"] == sum(1 for it in out["snapshots"] if "command" in it)


def test_undo_list_more_than_20_undoable(capsys, monkeypatch):
    snaps = [{"app": "word", "name": f"d{i}.docx", "created_at": "20261103_100000", "by": "word write",
              "meta_file": f"{i}.json", "original_path": ""} for i in range(25)]
    snaps.insert(3, _barrier(99))
    monkeypatch.setattr(off, "list_snapshots", lambda kind=None: snaps)
    code, out = _run(office, capsys, monkeypatch, "office", "undo", "--list")
    assert out["count"] == 25 and out["shown"] == 20 and out["truncated"] is True and out["barriers"] == 1, out
    assert [it["pick"] for it in out["snapshots"]] == list(range(1, 22)), out["snapshots"]
