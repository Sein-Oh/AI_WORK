# -*- coding: utf-8 -*-
"""1.4.18 (9차 점검): ppt replace의 UTF-16 위치(이모지 뒤 글), Word 찾기의 ^ 글자."""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
TOOL = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(HERE))), "CLI", "office-cli")
sys.path[:0] = [TOOL, HERE]
import ppt  # noqa: E402
import word  # noqa: E402


def _u16(s):
    return len(s.encode("utf-16-le")) // 2


class _Obj:
    pass


class U16TextRange:
    """PowerPoint처럼 After/Start를 UTF-16 단위로 다루는 가짜 TextRange (Text는 파이썬 문자열)."""

    def __init__(self, text):
        self.Text = text

    def Replace(self, FindWhat, ReplaceWhat, After=0, MatchCase=0, WholeWords=0):
        units, start = 0, len(self.Text)
        for i, ch in enumerate(self.Text):
            if units >= After:
                start = i
                break
            units += _u16(ch)
        hay = self.Text if MatchCase else self.Text.lower()
        idx = hay.find(FindWhat if MatchCase else FindWhat.lower(), start)
        if idx < 0:
            return None
        self.Text = self.Text[:idx] + ReplaceWhat + self.Text[idx + len(FindWhat):]
        found = _Obj()
        found.Start = _u16(self.Text[:idx]) + 1
        found.Font = _Obj()
        return found


def test_ppt_replace_after_emoji_replaces_all():
    for text, find, rep, want in (("\U0001F4CA\U0001F4CA\U0001F4CA 2025 2025", "2025", "2026", "\U0001F4CA\U0001F4CA\U0001F4CA 2026 2026"),
                                  ("\U0001F4CC\U0001F4CC ab ab", "ab", "X", "\U0001F4CC\U0001F4CC X X"),
                                  ("\U0001F4CC 핵심 진행중, 진행중", "진행중", "완료", "\U0001F4CC 핵심 완료, 완료")):
        tr = U16TextRange(text)
        assert ppt.replace_in_range(tr, find, rep, False, {}) == (2, 0), text
        assert tr.Text == want


def test_word_find_escapes_caret():
    seen = {}

    class F:
        def ClearFormatting(self):
            pass

        def Execute(self, FindText="", **kw):
            seen["text"] = FindText
            return False

    rng = _Obj()
    rng.Find = F()
    word._find(rng, "감사합니다^^ 10^6")
    assert seen["text"] == "감사합니다^^^^ 10^^6"


# 1.4.19 (10차 점검): Word Find 255자 한도는 ^를 ^^로 바꾼 길이로 세고, 긴 바꿀 글은 Word Find 없이 다시 셈 ------------

import json  # noqa: E402

import pytest  # noqa: E402

import _office as off  # noqa: E402
from fake_office import make_win32com  # noqa: E402
from fake_word import FakeWordDocument  # noqa: E402


@pytest.fixture
def wd(monkeypatch, tmp_path):
    monkeypatch.setenv("OFFICE_CLI_DATA_DIR", str(tmp_path / "d"))
    holder = {"doc": FakeWordDocument("제목\r본문 ^ 10^6 기호\r일정")}
    fake, _ = make_win32com()
    monkeypatch.setattr(off, "win32com", fake)
    monkeypatch.setattr(off, "pythoncom", None)
    monkeypatch.setattr(off, "list_open", lambda kind: [])
    monkeypatch.setattr(off, "get_active", lambda kind: {"name": holder["doc"].Name, "path": holder["doc"].FullName,
                                                         "app": object(), "doc": holder["doc"]})
    return holder


def _run_word(capsys, *argv):
    with pytest.raises(SystemExit) as ei:
        word.main(list(argv))
    out = [ln for ln in capsys.readouterr().out.splitlines() if ln.strip()]
    assert len(out) == 1, out
    return ei.value.code, json.loads(out[0])


def test_word_long_replacement_containing_find_is_verified(wd, capsys):
    rep = "제목: " + "가" * 300
    code, out = _run_word(capsys, "replace", "--find", "제목", "--replace", rep)
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    assert wd["doc"].text.startswith(rep)


def test_word_find_limit_counts_escaped_carets(wd, capsys):
    code, out = _run_word(capsys, "replace", "--find", "^" * 200, "--replace", "x")
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:"), out
    assert "255" in out["message"]
