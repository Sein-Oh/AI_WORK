# -*- coding: utf-8 -*-
"""작업별 스킬 4개(AIKIT/.opencode/skills/<이름>/SKILL.md, 2.9.71, 2.9.72: 입력 자동 선택, --suggest)와 코드가 어긋나지 않는지: 버전, OpenCode 머리말, 백틱 명령의 --옵션,
크기(8KB), 공통 규칙이 4개에서 같음, 모든 스크립트가 어느 스킬에든 있음, 꼭 적어야 할 기능."""
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from mlkit_helpers import KIT  # noqa: E402
import _ml_runtime as rt  # noqa: E402

SKILLS = ("ml-train", "ml-predict", "ml-report", "ml-map", "ml-dataset")
MAX_SKILL_BYTES = 8192                # 경량 LLM 컨텍스트: 스킬 하나 8KB 이하 (6KB -> 7.2KB -> 8KB, 사용자 결정 2026-11-10, AGENTS.md "사용법 단순성" 7)


def _path(name):
    return os.path.join(os.path.dirname(KIT), ".opencode", "skills", name, "SKILL.md")      # C:\AIKIT\.opencode\skills


def _read(name):
    return open(_path(name), encoding="utf-8").read()


def test_old_single_skill_is_gone():
    assert not os.path.exists(os.path.join(KIT, "SKILL.md"))
    assert sorted(n for n in os.listdir(os.path.join(os.path.dirname(KIT), ".opencode", "skills")) if n.startswith("ml-")) == sorted(SKILLS)


def test_version_and_frontmatter():
    for name in SKILLS:
        text = _read(name)
        head = text.split("---", 2)[1]
        keys = re.findall(r"^([a-z]+):", head, re.M)
        assert keys == ["name", "description", "license", "compatibility", "metadata"], (name, keys)
        assert re.search(rf"^name: {name}$", head, re.M), name
        m = re.search(r'^\s*version:\s*"?([0-9.]+)"?\s*$', head, re.M)
        assert m and m.group(1) == rt.VERSION, name


def test_skill_size_limit():
    for name in SKILLS:
        with open(_path(name), "rb") as f:
            size = len(f.read().replace(b"\r\n", b"\n"))
        assert size <= MAX_SKILL_BYTES, (name, size)


def test_common_rules_are_the_same_in_every_skill():
    """공통 규칙이 스킬마다 다르면 경량 LLM이 어느 쪽을 따를지 흔들림: 4개에서 글자까지 같음."""
    blocks = {name: re.search(r"## 규칙 \(ml-kit 공통\)\n(.*?)\n## ", _read(name), re.S).group(1) for name in SKILLS}
    assert len(set(blocks.values())) == 1, {k: len(v) for k, v in blocks.items()}


def _argparse_options(script):
    src = open(os.path.join(KIT, script), encoding="utf-8").read()
    opts = set()
    for call in re.findall(r"add_argument\(([^)]*)", src):
        opts.update(re.findall(r"[\"'](--[a-z][a-z0-9-]*)[\"']", call))
    if "add_sheet_argument(parser)" in src:          # 공통 런타임이 붙이는 --sheet (2.9.88)
        opts.add("--sheet")
    return opts


def _backtick_spans(text):
    spans = re.findall(r"```[a-z]*\n(.*?)```", text, flags=re.S)
    rest = re.sub(r"```.*?```", "", text, flags=re.S)
    return spans + re.findall(r"`([^`\n]+)`", rest)


def test_skill_options_exist_in_argparse():
    scripts = sorted(f for f in os.listdir(KIT) if f.startswith("ml_") and f.endswith(".py"))
    per = {s: _argparse_options(s) for s in scripts}
    union = set().union(*per.values())
    assert "--history" in per["ml_train.py"] and "--save" in per["ml_predict.py"] and "--out" not in per["ml_predict.py"]
    bad, checked = [], 0
    for name in SKILLS:
        for span in _backtick_spans(_read(name)):
            named = [s for s in scripts if s in span]
            if not named and re.search(r"\w+\.py\b", span):
                continue                                                   # another tool (office-cli excel.py)
            for line in span.splitlines():
                line_scripts = [s for s in scripts if s in line] or named
                allowed = set().union(*(per[s] for s in line_scripts)) if line_scripts else union
                for o in re.findall(r"(?<![\w-])(--[a-z][a-z0-9-]*)", line):
                    checked += 1
                    if o not in allowed:
                        bad.append(f"{name}: {o} ({', '.join(line_scripts) or 'any ml-kit script'}): {line.strip()[:80]}")
    assert checked >= 20
    assert not bad, bad


def test_every_script_is_in_its_skill():
    want = {"ml-train": ["ml_train.py"], "ml-predict": ["ml_predict.py"], "ml-report": ["ml_report.py", "ml_export.py"],
            "ml-map": ["ml_map.py"], "ml-dataset": ["ml_dataset.py"]}
    for name, scripts in want.items():
        text = _read(name)
        for s in scripts:
            assert f'python "C:\\AIKIT\\ml-kit\\{s}"' in text, (name, s)
    assert set(rt.SKILL_OF_SCRIPT.values()) == set(SKILLS)


def test_skill_mentions_features():
    """경량 LLM은 스킬에 없는 기능을 '지원 안 함'으로 답함: 기능마다 맡은 스킬에."""
    need = {
        "ml-train": ("--task classification", "--positive-class", "--dtypes", "--eval \"", "ml-dataset", "features.txt",
                     "--history", "--dry-run", "goal_met", "improve_stopped", "delta_score",
                     "입력 열(피처)은 묻지 않는다", "--features \"<피처1>,<피처2>\""),
        "ml-predict": ("--model \"<데이터파일>\"", "model_chosen", "--file \"<새데이터파일>\"", "--save", "--find", "--vary", "--sweep",
                       "--maximize", "--minimize", "--where", "--suggest <개수>", "실험계획", "실제값 비교"),
        "ml-report": ("ml_report.py\" --file", "--title", "--slots", "ml_export.py\" --model", "최종 버전"),
        "ml-map": ("ml_map.py\" --file", "--find", "--vary", "--points", "extrapolated"),
        "ml-dataset": ("ml_dataset.py\" pick --file", "ml_dataset.py\" exclude --file", "--where \"<조건>\"", "--rows", "--reason",
                       "next_commands[0]", "--save"),
    }
    for name, needles in need.items():
        text = _read(name)
        for needle in needles:
            assert needle in text, (name, needle)


def test_removed_options_are_not_in_skills():
    """2.9.83: 최고 버전(--best)과 성능예측표 --keep, 2.9.84: 성능예측표 --inputs를 없앰. 스킬에 남으면 경량 LLM이 없는 옵션을 씀."""
    for name in SKILLS:
        text = _read(name)
        for gone in ("--best", "--keep", "best_version", "최고 버전", "--inputs", "--exclude-where"):
            assert gone not in text, (name, gone)


def test_train_templates_have_no_x():
    """2.9.72: --features는 사용자가 입력 열을 말했을 때만 (표의 한 줄). 학습 명령 틀에는 --features가 없어 도구가 입력을 고름 (한 가지 형태)."""
    text = _read("ml-train")
    templates = [b for b in re.findall(r"```powershell\n(.*?)```", text, flags=re.S) if "ml_train.py" in b]
    assert templates and all(" --features " not in b for b in templates), templates
