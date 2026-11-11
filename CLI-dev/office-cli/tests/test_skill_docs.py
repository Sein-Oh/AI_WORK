"""작업별 스킬(CLI/office-cli/skills/<이름>/SKILL.md)이 코드와 어긋나지 않음: 버전 표기, OpenCode frontmatter 규칙,
명령·옵션이 argparse에 있음, 명령 템플릿이 전체 명령, 모든 명령이 어느 스킬에든 있음, 기능 목록, 크기."""
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
SKILL_NAMES = ("office-excel", "office-doc", "office-mail", "office-spec", "office-data")
SKILLS_DIR = os.path.join(TOOL_DIR, "skills")
SKILL_FILES = {name: os.path.join(SKILLS_DIR, name, "SKILL.md") for name in SKILL_NAMES}
TEST_SKILL = os.path.join(DEV, "tools", "office-cli-test", "SKILL.md")
MAX_SKILL_BYTES = 8192                # 경량 LLM 컨텍스트: 스킬 하나 8KB 이하 (6KB -> 7.2KB -> 8KB, 사용자 결정 2026-11-10, AGENTS.md "사용법 단순성" 7)


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


_FULL = re.compile(r'^python\s+"C:\\CLI\\office-cli\\(\w+\.py)"(?=\s|$)')


def _read(path):
    with open(path, encoding="utf-8") as f:
        return f.read()


def _skill_refs(path):
    """SKILL.md의 백틱 구간에서 (줄 번호, 스크립트, 명령, 옵션, 옵션 뒤 값) 목록.
    명령 템플릿은 `python "C:\\CLI\\office-cli\\<파일>.py" <명령> ...` 전체 명령.
    표의 한 행 안에서 전체 명령 다음의 `--dry-run`이나 `sheet --action ...`은 앞의 스크립트·명령으로 봅니다."""
    refs = []
    for no, line in enumerate(_read(path).splitlines(), start=1):
        script = sub = None
        for span in _SPAN.findall(line):
            m = _FULL.match(span)
            if m:
                span = m.group(1) + span[m.end():]
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


def _frontmatter(path):
    text = _read(path)
    m = re.match(r"---\n(.*?)\n---\n", text, re.S)
    assert m, path
    return m.group(1)


# ---------------------------------------------------------------------------
# 배치: 작업별 스킬 5개, 예전 통합 SKILL.md 없음
# ---------------------------------------------------------------------------


def test_skill_layout():
    assert sorted(os.listdir(SKILLS_DIR)) == sorted(SKILL_NAMES)
    for name, path in SKILL_FILES.items():
        assert os.listdir(os.path.dirname(path)) == ["SKILL.md"], name          # 스킬 폴더에는 SKILL.md만
    assert not os.path.exists(os.path.join(TOOL_DIR, "SKILL.md"))               # 남으면 옛 통합 스킬이 다시 쓰임


def test_skill_size_limit():
    # 줄 끝을 LF로 세어 잼: Windows에서 core.autocrlf=true로 받으면 CRLF가 되어 줄 수만큼 커짐 (집 PC 2026-11-06, 6123 -> 6173)
    for name, path in SKILL_FILES.items():
        with open(path, "rb") as f:
            size = len(f.read().replace(b"\r\n", b"\n"))
        assert size <= MAX_SKILL_BYTES, (name, size)


# ---------------------------------------------------------------------------
# 도구 버전과 SKILL.md 표기가 같음, OpenCode frontmatter 규칙
# ---------------------------------------------------------------------------


def test_version_matches_skill_md():
    for skill in list(SKILL_FILES.values()) + [TEST_SKILL]:
        m = re.search(r'^\s+version:\s*"([^"]+)"', _read(skill), re.M)
        assert m and m.group(1) == rt.VERSION, skill


def test_frontmatter_follows_opencode_rules():
    """폴더명 = name, name 형식, 키는 name/description/license/compatibility/metadata만, description 1024자 이하."""
    for skill in list(SKILL_FILES.values()) + [TEST_SKILL]:
        fm = _frontmatter(skill)
        keys = re.findall(r"^([A-Za-z_][\w-]*):", fm, re.M)
        assert set(keys) <= {"name", "description", "license", "compatibility", "metadata"}, (skill, keys)
        assert {"name", "description"} <= set(keys), skill
        name = re.search(r"^name:\s*(\S+)\s*$", fm, re.M).group(1)
        assert re.fullmatch(r"[a-z0-9]+(-[a-z0-9]+)*", name) and name == os.path.basename(os.path.dirname(skill)), skill
        desc = re.search(r"^description:\s*(.+)$", fm, re.M).group(1).strip()
        assert 0 < len(desc) <= 1024, skill


def test_runtime_help_points_to_task_skill():
    for tool, name in rt.SKILL_OF_TOOL.items():
        assert name in SKILL_NAMES and os.path.isfile(rt.skill_file(tool)), tool
    for (tool, cmd), name in rt.SKILL_OF_COMMAND.items():     # 도구 파일과 다른 스킬에 있는 명령 (excel compare -> office-data)
        assert name in SKILL_NAMES and cmd in PARSERS[f"{tool}.py"], (tool, cmd)
        assert f'{tool}.py" {cmd}' in _read(SKILL_FILES[name]), (tool, cmd)


# ---------------------------------------------------------------------------
# SKILL.md <-> argparse
# ---------------------------------------------------------------------------


def test_skill_commands_and_options_exist_in_argparse():
    problems = []
    for name, path in SKILL_FILES.items():
        refs = _skill_refs(path)
        assert len(refs) > 15, (name, refs[:5])          # 정규식이 표를 제대로 읽는지
        for no, script, sub, opt, val in refs:
            if script and sub:
                act = _option_in(PARSERS[script][sub], opt)
                if act is None:
                    problems.append(f"{name} {no}행: {script} {sub}에 {opt} 없음")
                    continue
                if act.choices and val and val.startswith("<") and "/" in val:      # 1.4.83: 고르기는 <a/b> (| 는 PowerShell 파이프)
                    bad = [v for v in val.strip("<>").split("/") if v not in act.choices]
                    if bad:
                        problems.append(f"{name} {no}행: {script} {sub} {opt}의 선택지에 {bad} 없음")
            else:       # 스크립트·명령을 알 수 없는 옵션 (규칙 절의 `--target` 등): 어느 명령에든 있어야 함
                if not any(_option_in(p, opt) for subs in PARSERS.values() for p in subs.values()):
                    problems.append(f"{name} {no}행: {opt}은(는) 어느 명령에도 없음")
    assert not problems, problems


def test_command_templates_are_full_commands():
    """경량 LLM이 경로를 조합하지 않게: 스크립트로 시작하는 백틱 구간은 모두 `python "C:\\CLI\\office-cli\\<파일>.py"`로 시작."""
    bare = []
    for name, path in SKILL_FILES.items():
        for no, line in enumerate(_read(path).splitlines(), start=1):
            for span in _SPAN.findall(line):
                first = span.split()[0] if span.split() else ""
                if first in PARSERS or (first == "python" and not _FULL.match(span)):
                    bare.append(f"{name} {no}행: {span[:60]}")
    assert not bare, bare


def test_every_command_is_in_some_skill():
    """예전 통합 SKILL.md의 명령이 작업별 스킬로 나뉘며 빠지지 않음: 모든 스크립트의 모든 하위 명령이 어느 스킬의 전체 명령에 있음."""
    seen = set()
    for path in SKILL_FILES.values():
        for _no, script, sub, _opt, _val in _skill_refs(path):
            seen.add((script, sub))
        for m in re.finditer(r'python "C:\\CLI\\office-cli\\(\w+\.py)" (\w+)', _read(path)):
            seen.add((m.group(1), m.group(2)))
    missing = [f"{s} {c}" for s, subs in PARSERS.items() for c in subs if (s, c) not in seen and (s, c) not in README_ONLY]
    assert not missing, missing


# 경량 LLM에게는 알리지 않는 명령 (README에만): 범위를 고르기 어려워 잘못 쓰기 쉬움 (사용자 결정 2026-11-10)
README_ONLY = {("word.py", "format")}


def test_readme_only_commands_are_not_in_skills_but_in_readme():
    readme = _read(os.path.join(os.path.dirname(os.path.dirname(__file__)), "README.md"))
    for script, sub in README_ONLY:
        assert all(f"{script}\" {sub}" not in _read(p) for p in SKILL_FILES.values()), (script, sub)
        assert f"| `{sub}` |" in readme, (script, sub)


def test_skill_mentions_new_and_existing_features():
    """스킬에 기존 기능과 새 기능이 모두 있음 (경량 LLM은 스킬에 없는 기능을 '지원 안 함'으로 답함). 기능마다 맡은 스킬에."""
    need = {
        "office-excel": ("sheet --action <add/rename/copy/delete>", "--new-name", "range --action <freeze", "export --save",
                         "--number-format", "--all-sheets", "--type <column/bar/line/pie/scatter/area>", "--action <insert/delete>",
                         "--dry-run", "--action sort", "--by", "--where \"", "--formulas", "--cells", "comment --action add",
                         "office.py\" undo", "transfer --from-app excel --chart"),
        "office-doc": ("slide --action <add/duplicate/delete/move>", "--toc", "--callout", "--highlight yellow", "fill --target", "--slots",
                       "outline", "revisions", "notes --slide", "--page-number", "transfer --from-app word --what text",
                       "office.py\" undo", "focus --app"),
        "office-mail": ("outlook.py\" folders", "list [--unread]", "--where subject", "--all", "new --to", "--cc", "--body-file",
                        "--days 7", "--since 2026-11-01", "thread --id", "calendar --days", "attachments --id", "send --id"),
        "office-spec": ("reader.py\" index --folder", "--section", "search --query", "diff --file", "--file2", "list [--name",
                        "clear --name", "--missing", "synonyms.txt", "--regex", "excel.py\" check", "comment --action add",
                        "not_searched", "match_mode"),
        "office-data": ("excel.py\" compare --target", "--target2", "--sheet2", "--key", "--save", "replaced_guess", "moved", "--level", "비교결과",
                        "confirm_commands"),
    }
    for name, needles in need.items():
        text = _read(SKILL_FILES[name])
        for needle in needles:
            assert needle in text, (name, needle)


def test_no_pipe_alternatives_in_templates():
    """1.4.83 검토: 템플릿의 'freeze|unfreeze'를 그대로 베끼면 따옴표 밖의 |가 PowerShell 파이프가 됨 -> 고르기는 <a/b>."""
    bad = []
    for name, path in SKILL_FILES.items():
        for no, line in enumerate(open(path, encoding="utf-8").read().splitlines(), 1):
            for span in re.findall(r"`([^`]+)`", line):
                if re.search(r"(--[a-z-]+\s+)[a-z_]+\|[a-z_]", _QUOTED.sub("", span)):
                    bad.append(f"{name} {no}: {span[:80]}")
    assert not bad, bad
