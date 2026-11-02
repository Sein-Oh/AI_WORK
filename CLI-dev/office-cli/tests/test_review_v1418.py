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
