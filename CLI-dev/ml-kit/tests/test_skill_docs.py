# -*- coding: utf-8 -*-
"""SKILL.md와 코드가 어긋나지 않는지: 버전, 백틱 명령의 --옵션, 꼭 적어야 할 기능."""
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from mlkit_helpers import KIT  # noqa: E402
import ml_predict as mp  # noqa: E402


def test_version_matches_skill_metadata():
    text = open(os.path.join(KIT, "SKILL.md"), encoding="utf-8").read()
    m = re.search(r'^\s*version:\s*"?([0-9.]+)"?\s*$', text, re.M)
    assert m and m.group(1) == mp.VERSION
    import ml_report
    assert ml_report.REPORT_VERSION == mp.VERSION


def _argparse_options(script):
    src = open(os.path.join(KIT, script), encoding="utf-8").read()
    opts = set()
    for call in re.findall(r"add_argument\(([^)]*)", src):
        opts.update(re.findall(r"[\"'](--[a-z][a-z0-9-]*)[\"']", call))
    return opts


def _backtick_spans(text):
    spans = re.findall(r"```[a-z]*\n(.*?)```", text, flags=re.S)
    rest = re.sub(r"```.*?```", "", text, flags=re.S)
    return spans + re.findall(r"`([^`\n]+)`", rest)


def test_skill_options_exist_in_argparse():
    text = open(os.path.join(KIT, "SKILL.md"), encoding="utf-8").read()
    scripts = sorted(f for f in os.listdir(KIT) if f.startswith("ml_") and f.endswith(".py"))
    per = {s: _argparse_options(s) for s in scripts}
    union = set().union(*per.values())
    assert "--history" in per["ml_train.py"] and "--out" in per["ml_predict.py"]
    bad, checked = [], 0
    for span in _backtick_spans(text):
        opts = re.findall(r"(?<![\w-])(--[a-z][a-z0-9-]*)", span)
        if not opts:
            continue
        named = [s for s in scripts if s in span]
        if not named and re.search(r"\w+\.py\b", span):
            continue                                                   # another tool (office-cli excel.py)
        for line in span.splitlines():
            line_scripts = [s for s in scripts if s in line] or named
            allowed = set().union(*(per[s] for s in line_scripts)) if line_scripts else union
            for o in re.findall(r"(?<![\w-])(--[a-z][a-z0-9-]*)", line):
                checked += 1
                if o not in allowed:
                    bad.append(f"{o} ({', '.join(line_scripts) or 'any ml-kit script'}): {line.strip()[:80]}")
    assert checked >= 20
    assert not bad, bad


def test_skill_mentions_batch_features():
    text = open(os.path.join(KIT, "SKILL.md"), encoding="utf-8").read()
    assert re.search(r"ml_train\.py\" -f \"[^\"]+\" --history", text)
    assert re.search(r"ml_predict\.py\" -m \"[^\"]+\" -f \"[^\"]+\" --out \"[^\"]+\.xlsx>?\"", text)
    assert "--eval \"" in text and "top_features" in text
