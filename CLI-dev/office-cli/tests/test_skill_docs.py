"""SKILL.md가 코드와 어긋나지 않음: 버전 표기, 명령·옵션이 argparse에 있음, 기능 목록."""
import argparse
import os
import re

import _runtime as rt
import excel
import office
import outlook
import ppt
import reader
import word
from office_helpers import DEV, TOOL_DIR


SCRIPTS = {"excel.py": excel, "word.py": word, "ppt.py": ppt, "outlook.py": outlook, "office.py": office,
           "reader.py": reader}


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


# ---------------------------------------------------------------------------
# 도구 버전과 SKILL.md 표기가 같음
# ---------------------------------------------------------------------------


def test_version_matches_skill_md():
    import re
    for skill in (os.path.join(TOOL_DIR, "SKILL.md"), os.path.join(DEV, "tools", "office-cli-test", "SKILL.md")):
        with open(skill, encoding="utf-8") as f:
            m = re.search(r'^\s+version:\s*"([^"]+)"', f.read(), re.M)
        assert m and m.group(1) == rt.VERSION, skill


# ---------------------------------------------------------------------------
# SKILL.md <-> argparse
# ---------------------------------------------------------------------------


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
    """SKILL.md에 기존 기능과 새 기능이 모두 있음 (경량 LLM은 SKILL.md에 없는 기능을 '지원 안 함'으로 답함)."""
    text = open(os.path.join(TOOL_DIR, "SKILL.md"), encoding="utf-8").read()
    for needle in ("outlook.py folders", "list [--unread]", "--where subject", "reply --all", "new --cc", "--body-file",
                   "slide --action add|duplicate|delete|move", "sheet --action add|rename|copy|delete", "--new-name",
                   "range --action freeze", "export --save", "--number-format", "--all-sheets",
                   "--type column|bar|line|pie|scatter|area", "--toc", "--callout", "--highlight yellow",
                   "--days 7", "--since 2026-11-01", "--action insert|delete", "--dry-run", "--action sort", "--by", "--where \""):
        assert needle in text, needle
