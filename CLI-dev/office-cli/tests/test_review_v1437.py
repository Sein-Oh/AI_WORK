"""1.4.37 회귀 테스트.

- SKILL.md에 적은 명령·옵션·선택지가 argparse에 모두 있는지 (앞으로 문서와 코드가 어긋나지 않게)
- B1 outlook list/search --days / --since (받은 날짜 범위)
- B3 excel range --action insert|delete (행·열 넣기·지우기)
- B4 excel range --action sort (정렬)
- B5 excel read --where (조건에 맞는 행만)
"""
import argparse
import datetime as dt
import json
import os
import re

import pytest

import _runtime as rt

import excel
import office
import outlook
import ppt
import reader
import word

TOOL_DIR = os.path.dirname(os.path.abspath(excel.__file__))
SCRIPTS = {"excel.py": excel, "word.py": word, "ppt.py": ppt, "outlook.py": outlook, "office.py": office,
           "reader.py": reader}


# ---------------------------------------------------------------------------------------------------------------
# SKILL.md <-> argparse
# ---------------------------------------------------------------------------------------------------------------

def _subparsers(mod):
    parser = mod.build_cli()
    for act in parser._actions:
        if isinstance(act, argparse._SubParsersAction):
            return dict(act.choices)
    raise AssertionError(mod)


PARSERS = {name: _subparsers(mod) for name, mod in SCRIPTS.items()}
_SPAN = re.compile(r"`([^`]+)`")
_QUOTED = re.compile(r"\"[^\"]*\"|'[^']*'")
_OPT = re.compile(r"(?<![\w-])(--?[a-zA-Z][\w-]*)(?:\s+([^\s\[\]()]+))?")


def _skill_refs():
    """SKILL.md의 백틱 구간에서 (줄 번호, 스크립트, 명령, 옵션, 옵션 뒤 값) 목록.
    표의 한 행 안에서 `excel.py range ...` 다음의 `--dry-run`이나 `sheet --action ...`은 앞의 스크립트·명령으로 봅니다."""
    path = os.path.join(TOOL_DIR, "SKILL.md")
    text = open(path, encoding="utf-8").read()
    refs = []
    for no, line in enumerate(text.splitlines(), start=1):
        script = sub = None
        for span in _SPAN.findall(line):
            words = span.split()
            if not words:
                continue
            if words[0].endswith(".py") and words[0] in PARSERS:
                script = words[0]
                sub = words[1] if len(words) > 1 and words[1] in PARSERS[script] else None
                rest = " ".join(words[2:] if sub else words[1:])
            elif script and words[0] in PARSERS[script]:
                sub = words[0]
                rest = " ".join(words[1:])
            elif words[0].startswith("-"):
                rest = span
                if not line.startswith("|"):      # 표 밖(규칙 글)의 옵션은 앞 명령에 딸린 것이 아님
                    script = sub = None
            else:
                continue
            for opt, val in _OPT.findall(_QUOTED.sub(" Q ", rest)):
                refs.append((no, script, sub, opt, val))
    return refs


def _option_in(sub_parser, opt):
    return sub_parser._option_string_actions.get(opt)


def test_skill_commands_and_options_exist_in_argparse():
    refs = _skill_refs()
    assert len(refs) > 80, refs[:5]          # 정규식이 표를 제대로 읽는지
    problems = []
    for no, script, sub, opt, val in refs:
        if script and sub:
            act = _option_in(PARSERS[script][sub], opt)
            if act is None:
                problems.append(f"{no}행: {script} {sub}에 {opt} 없음")
                continue
            if act.choices and val and "|" in val:
                bad = [v for v in val.split("|") if v not in act.choices]
                if bad:
                    problems.append(f"{no}행: {script} {sub} {opt}의 선택지에 {bad} 없음")
        else:       # 스크립트·명령을 알 수 없는 옵션 (규칙 절의 `--target` 등): 어느 명령에든 있어야 함
            if not any(_option_in(p, opt) for subs in PARSERS.values() for p in subs.values()):
                problems.append(f"{no}행: {opt}은(는) 어느 명령에도 없음")
    assert not problems, problems


def test_skill_mentions_new_and_existing_features():
    """리뷰에서 SKILL.md에 빠졌던 기존 기능과 1.4.37 새 기능 (경량 LLM은 SKILL.md에 없는 기능을 '지원 안 함'으로 답함)."""
    text = open(os.path.join(TOOL_DIR, "SKILL.md"), encoding="utf-8").read()
    for needle in ("outlook.py folders", "list [--unread]", "--where subject", "reply --all", "new --cc", "--body-file",
                   "slide --action add|duplicate|delete|move", "sheet --action add|rename|copy|delete", "--new-name",
                   "range --action freeze", "export --save", "--number-format", "--all-sheets",
                   "--type column|bar|line|pie|scatter|area", "--toc", "--callout", "--highlight yellow"):
        assert needle in text, needle


# ---------------------------------------------------------------------------------------------------------------
# B1 outlook --days / --since
# ---------------------------------------------------------------------------------------------------------------

def _run(mod, capsys, *argv):
    with pytest.raises(SystemExit) as ei:
        mod.main(list(argv))
    lines = [ln for ln in capsys.readouterr().out.splitlines() if ln.strip()]
    assert len(lines) == 1, lines
    return ei.value.code, json.loads(lines[0])


class _Mail:
    Class = 43
    SenderName = "홍길동"
    SenderEmailAddress = "hong@example.com"
    Body = "본문"

    def __init__(self, eid, subject, received, unread=False):
        self.EntryID = eid
        self.Subject = subject
        self.ReceivedTime = received
        self.UnRead = unread


class _Items:
    """Restrict는 조건 글을 기록하고, 안 읽은 메일 조건만 실제로 거름 (날짜는 도구가 받은 시각으로 다시 거르는지 확인).
    refuse: 날짜 조건을 거부하는 Outlook."""

    def __init__(self, mails, log, refuse=False):
        self.mails = list(mails)
        self.log = log
        self.refuse = refuse

    def Restrict(self, flt):
        self.log.append(flt)
        if self.refuse and ("ReceivedTime" in flt or "datereceived" in flt):
            raise Exception("Outlook: 조건을 해석할 수 없음")
        keep = [m for m in self.mails if m.UnRead] if flt == "[UnRead] = True" else self.mails
        return _Items(keep, self.log, self.refuse)

    def Sort(self, *a):
        self.mails.sort(key=lambda m: m.ReceivedTime, reverse=True)

    def __iter__(self):
        return iter(self.mails)


TODAY = dt.date(2026, 11, 4)


def _mails():
    return [_Mail("c", "지난주 견적", dt.datetime(2026, 10, 28, 8, 0), unread=True),
            _Mail("a", "오늘 견적", dt.datetime(2026, 11, 4, 9, 0), unread=True),
            _Mail("d", "지난달 견적", dt.datetime(2026, 10, 1, 8, 0)),
            _Mail("b", "어제 회의", dt.datetime(2026, 11, 3, 23, 59))]


@pytest.fixture
def ol(monkeypatch):
    log = []
    state = {"refuse": False, "mails": _mails()}

    class _Folder:
        Name = "받은 편지함"

        @property
        def Items(self):
            return _Items(state["mails"], log, state["refuse"])

    monkeypatch.setattr(outlook, "_today", lambda: TODAY)
    monkeypatch.setattr(rt, "RECOVERED", [])
    monkeypatch.setattr(outlook, "get_outlook", lambda: object())
    monkeypatch.setattr(outlook, "get_namespace", lambda o: object())
    monkeypatch.setattr(outlook, "resolve_folder", lambda ns, f: _Folder())
    return {"log": log, "state": state}


def test_list_days_includes_today(ol, capsys):
    code, out = _run(outlook, capsys, "list", "--days", "2")
    assert code == 0 and [m["id"] for m in out["mails"]] == ["a", "b"], out
    assert out["period"] == {"since": "2026-11-03", "until": "2026-11-04", "days": 2}, out
    assert "2026-11-03 ~ 2026-11-04" in out["preview"], out
    assert "[ReceivedTime] >= '2026-11-03 00:00'" in ol["log"], ol["log"]       # 지역 설정과 무관한 YYYY-MM-DD


def test_list_days_one_is_today_only(ol, capsys):
    code, out = _run(outlook, capsys, "list", "--days", "1")
    assert code == 0 and [m["id"] for m in out["mails"]] == ["a"], out


def test_list_since_with_unread(ol, capsys):
    code, out = _run(outlook, capsys, "list", "--since", "2026-10-28", "--unread")
    assert code == 0 and [m["id"] for m in out["mails"]] == ["a", "c"], out
    assert out["period"]["days"] == 8 and ol["log"][0] == "[UnRead] = True", (out, ol["log"])


def test_search_days_with_query_uses_one_dasl_filter(ol, capsys):
    code, out = _run(outlook, capsys, "search", "--query", "견적", "--days", "7")
    assert code == 0 and [m["id"] for m in out["mails"]] == ["a", "b"], out      # 가짜 Restrict는 글로 거르지 않음 (c는 10-28)
    flt = ol["log"][-1]
    assert flt.startswith("@SQL=") and flt.count("@SQL=") == 1 and "견적" in flt, flt
    assert '"urn:schemas:httpmail:datereceived" >= ' in flt and ") AND (" in flt, flt
    utc = dt.datetime(2026, 10, 29).astimezone(dt.timezone.utc).strftime("%Y-%m-%d %H:%M")
    assert f"'{utc}'" in flt, flt                                                 # DASL 날짜 비교는 UTC
    assert out["period"]["since"] == "2026-10-29" and "2026-10-29" in out["preview"], out


def test_search_date_only_is_allowed(ol, capsys):
    code, out = _run(outlook, capsys, "search", "--since", "2026-11-03")
    assert code == 0 and [m["id"] for m in out["mails"]] == ["a", "b"], out


def test_zero_result_states_period(ol, capsys):
    ol["state"]["mails"] = [m for m in _mails() if m.EntryID in ("c", "d")]
    for argv in (["list", "--days", "3"], ["search", "--query", "견적", "--days", "3"]):
        code, out = _run(outlook, capsys, *argv)
        assert code == 0 and out["count"] == 0 and out["mails"] == [], out
        assert out["agent_directive"].startswith("VERIFIED: 0건") and "2026-11-02 ~ 2026-11-04" in out["agent_directive"], out


def test_restrict_refused_falls_back_and_filters(ol, capsys):
    ol["state"]["refuse"] = True
    code, out = _run(outlook, capsys, "list", "--days", "2")
    assert code == 0 and [m["id"] for m in out["mails"]] == ["a", "b"], out
    assert out["agent_directive"].startswith("VERIFIED:") and out["recovered"], out
    rt.RECOVERED.clear()
    code, out = _run(outlook, capsys, "search", "--query", "견적", "--days", "7")
    assert code == 0 and [m["id"] for m in out["mails"]] == ["a", "b"] and out["recovered"], out
    assert "datereceived" not in ol["log"][-1] and "견적" in ol["log"][-1], ol["log"]
    rt.RECOVERED.clear()


@pytest.mark.parametrize("argv, rec", [
    (["--since", "2026/10/01"], "[추천] 2026-10-01"),
    (["--since", "20261001"], "[추천] 2026-10-01"),
    (["--since", "2026년 10월 1일"], "[추천] 2026-10-01"),
    (["--since", "2026-1-5"], "[추천] 2026-01-05"),
    (["--days", "0"], "[추천] 1"),
    (["--days", "-7"], "[추천] 7"),
])
def test_bad_date_recommends(ol, capsys, argv, rec):
    for cmd in ("list", "search"):
        code, out = _run(outlook, capsys, cmd, *argv)
        assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:") and out["message"].endswith(rec), out
    assert not ol["log"], ol["log"]          # Outlook에 묻기 전에 멈춤


@pytest.mark.parametrize("argv", [["--since", "2026-02-30"], ["--since", "지난주"], ["--since", "2026-12-01"],
                                  ["--days", "3", "--since", "2026-11-01"]])
def test_bad_date_without_guess(ol, capsys, argv):
    code, out = _run(outlook, capsys, "list", *argv)
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:") and "[추천]" not in out["message"], out
