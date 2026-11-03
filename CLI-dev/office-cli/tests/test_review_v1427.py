"""1.4.27 (18차 점검) 회귀 테스트.

F1/F5: 제목 굵게 표시 규칙이 수준마다 달랐음 (ppt 3단계 이하만 ** __를 빼고 1·2단계 제목과 --title은 '**핵심** 요약' 그대로).
       word·ppt 모든 수준의 제목과 슬라이드 제목이 같은 규칙(** __를 뺀 글)을 씁니다.
F2: 단어 안의 '__'(report__2024__final.xlsx)를 굵게 표시로 읽어 'report2024final.xlsx'가 들어갔음 (CommonMark는 단어 안 __를 굵게로 보지 않음).
F3: '2026. 11. 3. 회의'를 번호 목록으로 읽어 연도가 사라졌고, 3부터 시작하는 번호 목록을 1부터 다시 매겼음.
F4: 오류 로그 message에 가린 인자 값(--anchor, --highlight)과 문서에서 가져온 후보·비슷한 문구가 그대로 남았음.
F6: undo --list가 되돌릴 수 없는 작업만 있을 때 '지점 0개'인데 count 1, 항목의 command를 실행하라는 안내였음.
F7: reader clear가 지우지 못한 캐시도 applied에 넣고, clear_done CSV를 지우기 전에 썼음 (지웠는지 열 없음).
F8: ppt slide --action list(0장), outlook list/search(0통)에 count와 '0건' 안내가 없었음.
"""
import json
import os
import sys

import pytest

import _office as off
import _runtime as rt
import excel
import office
import outlook
import ppt
import reader
import test_excel
import test_ppt
import test_word
import word
from fake_ppt import FakePresentation

xl = test_excel.xl
pp = test_ppt.pp
wd = test_word.wd


def _run(mod, capsys, *argv):
    with pytest.raises(SystemExit) as ei:
        mod.main(list(argv))
    out = [ln for ln in capsys.readouterr().out.splitlines() if ln.strip()]
    assert len(out) == 1, out
    return ei.value.code, json.loads(out[0])


def _texts(pp):
    prs = pp["prs"]
    s = prs.Slides(prs.Slides.Count)
    return [sh.TextFrame.TextRange.Text for sh in test_ppt.live(s) if hasattr(sh, "TextFrame")]


def _ppt_md(capsys, tmp_path, md, *extra):
    p = tmp_path / "s.md"
    p.write_text(md, encoding="utf-8")
    return test_ppt.run(capsys, "write", "--file", str(p), *extra)


def _word_md(capsys, tmp_path, md):
    p = tmp_path / "w.md"
    p.write_text(md, encoding="utf-8")
    return test_word.run(capsys, "write", "--file", str(p))


def _log_text(isolated_dirs):
    with open(os.path.join(isolated_dirs["data"], "logs", "errors.txt"), encoding="utf-8") as f:
        return f.read()


# F2 단어 안 __ ---------------------------------------------------------------------------------------------------

@pytest.mark.parametrize("text,plain", [
    ("report__2024__final.xlsx", "report__2024__final.xlsx"),
    ("파일__x__끝", "파일__x__끝"),
    ("__init__ 메서드", "init 메서드"),          # CommonMark: 양쪽이 단어 경계면 굵게
    ("(__a__)", "(a)"),
    ("x**2**y", "x2y"),                          # ** 는 단어 안에서도 굵게 (그대로)
    ("가 **나** 다 __라__", "가 나 다 라"),
])
def test_split_bold_intraword_underscore(text, plain):
    assert rt.split_bold(text)[0] == plain


def test_word_text_intraword_underscore_literal(wd, capsys):
    code, out = test_word.run(capsys, "write", "--text", "파일 report__2024__final.xlsx 참조")
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    assert "파일 report__2024__final.xlsx 참조" in test_word.paras(wd["doc"]), wd["doc"].text


def test_ppt_text_intraword_underscore_literal(pp, capsys):
    code, out = test_ppt.run(capsys, "write", "--slide", "new", "--text", "report__2024__final.xlsx 참조")
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    assert "report__2024__final.xlsx 참조" in _texts(pp)


# F1/F5 제목 규칙 -------------------------------------------------------------------------------------------------

@pytest.mark.parametrize("md,title,body0", [
    ("## **핵심** 요약\n\n본문\n", "핵심 요약", "본문"),
    ("# __init__ 메서드\n\n본문\n", "init 메서드", "본문"),
    ("## 클래스\n\n### __init__ 메서드\n\n본문\n", "클래스", "init 메서드"),
    ("## 클래스\n\n#### **핵심** 요약\n\n본문\n", "클래스", "핵심 요약"),
    ("## report__2024__final 파일\n\n본문\n", "report__2024__final 파일", "본문"),
])
def test_ppt_headings_same_rule_every_level(pp, capsys, tmp_path, md, title, body0):
    code, out = _ppt_md(capsys, tmp_path, md)
    texts = _texts(pp)
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    assert texts[0] == title and texts[1].split("\r")[0] == body0, texts


def test_ppt_md_title_moved_to_body_is_plain(pp, capsys, tmp_path):
    code, out = _ppt_md(capsys, tmp_path, "## **핵심** 요약\n\n본문\n", "--title", "새 제목")
    texts = _texts(pp)
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    assert texts == ["새 제목", "핵심 요약\r본문"], texts


def test_ppt_title_option_plain(pp, capsys):
    code, out = test_ppt.run(capsys, "write", "--slide", "new", "--title", "**핵심** 요약", "--text", "본문")
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    assert _texts(pp)[0] == "핵심 요약", _texts(pp)


@pytest.mark.parametrize("level", ["#", "##", "###", "####"])
def test_word_headings_same_rule_every_level(wd, capsys, tmp_path, level):
    code, out = _word_md(capsys, tmp_path, f"{level} __init__ 메서드\n\n본문 **강조**\n")
    paras = test_word.paras(wd["doc"])
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    assert "init 메서드" in paras and "본문 강조" in paras, paras


# F3 연도와 1이 아닌 번호 ----------------------------------------------------------------------------------------

def test_parse_blocks_year_is_not_list_and_start_number_kept():
    assert rt.parse_markdown_blocks("2026. 11. 3. 회의 결과") == [("paragraph", "2026. 11. 3. 회의 결과")]
    assert not rt.looks_like_markdown("2026. 11. 3. 회의 결과를 보고합니다.")
    assert rt.parse_markdown_blocks("3. 셋\n4. 넷") == [("paragraph", "3. 셋"), ("paragraph", "4. 넷")]
    assert rt.parse_markdown_blocks("1. 하나\n2. 둘") == [("numbered", ["하나", "둘"])]
    assert rt.parse_markdown_blocks("1. 하나\n\n2. 둘") == [("numbered", ["하나", "둘"])]   # 빈 줄로 띄운 같은 목록


def test_word_date_line_kept(wd, capsys):
    code, out = test_word.run(capsys, "write", "--text", "2026. 11. 3. 회의 결과를 보고합니다.")
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    assert "2026. 11. 3. 회의 결과를 보고합니다." in test_word.paras(wd["doc"]), wd["doc"].text


def test_word_md_list_from_3_keeps_numbers(wd, capsys, tmp_path):
    code, out = _word_md(capsys, tmp_path, "3. 셋\n4. 넷\n")
    paras = test_word.paras(wd["doc"])
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    assert "3. 셋" in paras and "4. 넷" in paras, paras


def test_ppt_date_line_kept(pp, capsys):
    code, out = test_ppt.run(capsys, "write", "--slide", "new", "--text", "2026. 11. 3. 회의 결과")
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    assert "2026. 11. 3. 회의 결과" in _texts(pp)


def test_ppt_text_list_from_3_keeps_numbers(pp, capsys):
    code, out = test_ppt.run(capsys, "write", "--slide", "new", "--text", "3. 셋\\n4. 넷\\n- 끝\\n1. 하나\\n2. 둘")
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    assert ppt._text_lines("3. 셋\\n4. 넷\\n- 끝\\n1. 하나\\n2. 둘") == [
        ("3. 셋", ""), ("4. 넷", ""), ("끝", "bullet"), ("하나", "numbered"), ("둘", "numbered")]
    assert "3. 셋\r4. 넷\r끝\r하나\r둘" in _texts(pp)


def test_ppt_md_list_from_3_keeps_numbers(pp, capsys, tmp_path):
    code, out = _ppt_md(capsys, tmp_path, "## 안건\n\n3. 셋\n4. 넷\n")
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    assert _texts(pp) == ["안건", "3. 셋\r4. 넷"], _texts(pp)


# F4 오류 로그 message 가리기 -------------------------------------------------------------------------------------

def test_log_masks_anchor_and_cuts_similar_text(wd, capsys, isolated_dirs, monkeypatch):
    argv = ["write", "--anchor", "2026년 비밀계획", "--text", "본문"]
    monkeypatch.setattr(sys, "argv", ["word.py"] + argv)
    monkeypatch.setattr(rt, "CURRENT_ARGV", None)
    code, out = _run(word, capsys, *argv)
    assert code == 1 and "2026년 비밀계획" in out["message"] and "비슷한 문구: '2026년 계획입'" in out["message"], out
    log = _log_text(isolated_dirs)
    assert "비밀계획" not in log and "계획입" not in log and "<10자 생략>" in log, log
    assert "문서에서 가져온 후보·문구" in log, log


def test_log_masks_excel_highlight_rule(xl, capsys, isolated_dirs, monkeypatch):
    argv = ["style", "--range", "A1:B2", "--highlight", "기밀완료"]
    monkeypatch.setattr(sys, "argv", ["excel.py"] + argv)
    monkeypatch.setattr(rt, "CURRENT_ARGV", None)
    code, out = _run(excel, capsys, *argv)
    assert code != 0 and "입력: 기밀완료" in out["message"], out
    log = _log_text(isolated_dirs)
    assert "기밀완료" not in log and "입력: <4자 생략>" in log, log


def test_log_cuts_candidates(isolated_dirs):
    msg = "제목 '결론'을(를) 찾을 수 없습니다. 비슷한 이름이 있지만 같은 것인지 도구가 정할 수 없습니다. [후보] 1.1 인수 계획, 1.2 비밀 일정"
    rt.write_error_log("word", ["read", "--section", "결론"], {"error_type": "input", "message": msg})
    log = _log_text(isolated_dirs)
    assert "인수 계획" not in log and "비밀 일정" not in log and "[후보]" not in log, log
    assert "제목 '<2자 생략>'을(를) 찾을 수 없습니다." in log, log
    msg2 = "제목 'x'을(를) 찾을 수 없습니다. 사용 가능한 제목: 1. 개요, 2. 비밀"
    assert "비밀" not in rt.log_message(msg2, [], "word")
    assert rt.log_message("시트 'a' [추천] 기밀시트", [], "excel").endswith("자 생략>")


def test_log_message_masks_equals_and_short_forms():
    m = rt.log_message("검색어 '견적서' / 절 '기밀'", ["search", "-q견적서", "--section=기밀"], "reader")
    assert "견적서" not in m and "기밀" not in m, m


def test_log_never_raises(isolated_dirs):
    class Bad:
        def __str__(self):
            raise RuntimeError("x")
    assert rt.log_message(Bad(), ["--text", "a"], "word").startswith("<message 생략")
    rt.write_error_log("word", ["write", "--text", "a"], {"error_type": "input", "message": Bad()})
    assert "message 생략" in _log_text(isolated_dirs)


# F6 undo --list 되돌릴 수 없는 작업만 -------------------------------------------------------------------------

def test_undo_list_barrier_only_is_zero(capsys, monkeypatch):
    monkeypatch.setattr(off, "list_snapshots", lambda kind=None: [
        {"app": "word", "name": "새문서.docx", "barrier": True, "reason": "새로 만든 파일이라 되돌릴 이전 상태가 없음",
         "created_at": "20261103_101500", "by": "word write", "meta_file": "x.json"}])
    code, out = _run(office, capsys, "undo", "--list")
    assert code == 0 and out["count"] == 0 and out["shown"] == 0 and out["barriers"] == 1, out
    assert out["agent_directive"].startswith("VERIFIED: 0건") and "command" not in out["agent_directive"], out


def test_undo_list_count_matches_preview(capsys, monkeypatch):
    snaps = [{"app": "excel", "name": "a.xlsx", "created_at": "20261103_101500", "by": "excel write", "meta_file": "1.json"},
             {"app": "word", "name": "b.docx", "barrier": True, "reason": "새 파일", "created_at": "20261103_101400",
              "by": "word write", "meta_file": "2.json"}]
    monkeypatch.setattr(off, "list_snapshots", lambda kind=None: snaps)
    code, out = _run(office, capsys, "undo", "--list")
    assert code == 0 and out["count"] == 1 and out["shown"] == 1 and out["preview"] == "되돌릴 수 있는 지점 1개", out


# F7 reader clear applied / done CSV ----------------------------------------------------------------------------

def test_reader_clear_applied_only_removed_and_done_csv(tmp_path, capsys, monkeypatch):
    docs = tmp_path / "docs"
    docs.mkdir()
    n = reader.MAX_LIST_ITEMS + 2
    for i in range(n):
        (docs / f"d{i:02d}.txt").write_text(f"내용 {i}", encoding="utf-8")
    code, out = _run(reader, capsys, "index", "--folder", str(docs))
    assert code == 0, out
    stuck = {reader._doc_cache_path(str(docs / "d00.txt")), reader._doc_cache_path(str(docs / "d01.txt"))}
    real_remove = os.remove

    def rm(p, *a, **k):
        if p in stuck:
            raise OSError("cannot remove")
        return real_remove(p, *a, **k)
    monkeypatch.setattr(reader.os, "remove", rm)
    code, out = _run(reader, capsys, "clear", "--all")
    assert code == 0 and out["agent_directive"].startswith("UNVERIFIED:"), out
    assert not any(a.endswith(("d00.txt", "d01.txt")) for a in out["applied"]), out["applied"]
    assert len(out["applied"]) == reader.MAX_LIST_ITEMS, out["applied"]
    rows = [r for r in rt.read_csv_rows(out["full_result_file"]) if any(r)]
    assert rows[0] == ["file", "path", "missing", "removed"], rows[0]
    flags = {r[0]: r[3] for r in rows[1:]}
    assert len(flags) == n and flags["d00.txt"] in (False, "False") and flags["d02.txt"] in (True, "True"), flags


# F8 0건 ----------------------------------------------------------------------------------------------------------

def test_ppt_slide_list_zero(pp, capsys):
    pp["prs"] = FakePresentation(0)
    code, out = test_ppt.run(capsys, "slide", "--action", "list")
    assert code == 0 and out["count"] == 0 and out["slides"] == [] and out["agent_directive"].startswith("VERIFIED: 0건"), out


def test_ppt_slide_list_count(pp, capsys):
    code, out = test_ppt.run(capsys, "slide", "--action", "list")
    assert code == 0 and out["count"] == len(out["slides"]) > 0, out


class _Mail:
    Class = 43
    EntryID = "id1"
    Subject = "견적 요청"
    SenderName = "홍길동"
    SenderEmailAddress = "hong@example.com"
    ReceivedTime = "2026-11-03 10:00"
    UnRead = True
    Body = "본문"


class _Items:
    def __init__(self, mails):
        self.mails = list(mails)
        self.filters = []

    def Restrict(self, flt):
        self.filters.append(flt)
        return self

    def Sort(self, *a):
        pass

    def __iter__(self):
        return iter(self.mails)


class _Folder:
    Name = "받은 편지함"

    def __init__(self, mails):
        self.Items = _Items(mails)


def _outlook(monkeypatch, mails):
    folder = _Folder(mails)
    monkeypatch.setattr(outlook, "get_outlook", lambda: object())
    monkeypatch.setattr(outlook, "get_namespace", lambda o: object())
    monkeypatch.setattr(outlook, "resolve_folder", lambda ns, f: folder)
    return folder


@pytest.mark.parametrize("argv", [["list"], ["list", "--unread"], ["search", "--query", "견적"]])
def test_outlook_zero(capsys, monkeypatch, argv):
    _outlook(monkeypatch, [])
    code, out = _run(outlook, capsys, *argv)
    assert code == 0 and out["count"] == 0 and out["mails"] == [], out
    assert out["agent_directive"].startswith("VERIFIED: 0건"), out
    assert "--folder" not in out["agent_directive"], out


@pytest.mark.parametrize("argv", [["list"], ["search", "--query", "견적"]])
def test_outlook_count(capsys, monkeypatch, argv):
    _outlook(monkeypatch, [_Mail()])
    code, out = _run(outlook, capsys, *argv)
    assert code == 0 and out["count"] == 1 and out["mails"][0]["subject"] == "견적 요청", out
