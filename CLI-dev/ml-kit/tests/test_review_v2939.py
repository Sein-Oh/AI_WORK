# -*- coding: utf-8 -*-
"""
2.9.39: important inputs in the training result, and SKILL.md option drift.

1. ml_train.py stdout JSON has top_features: {target: [{"feature", "importance"}, ...]} (up to 5, sorted, shares of
   the best model's permutation importances over all inputs) for regression, classification and multi-target.
   A categorical column is one entry under its own name (no one-hot names). The preview names them (중요 입력).
2. When the importances carry no information the target gets [] and a line in warnings.
3. Every --option in a backtick command of SKILL.md exists in that script's argparse.
"""
import os
import random
import re
import sys

import numpy as np
import pytest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from test_layout_matrix_v2917 import _write  # noqa: E402
from test_cli_smoke import run_raw  # noqa: E402

KIT = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(HERE))), "CLI", "ml-kit")
sys.path.insert(0, KIT)
FEATURES = ["온도", "압력", "라인", "잡음"]


def _ok(proc, res):
    assert proc.returncode == 0 and res["status"] == "ok", (res, proc.stderr[-2000:])
    return res


@pytest.fixture(scope="module")
def data(tmp_path_factory):
    tmp = tmp_path_factory.mktemp("v2939")
    r = random.Random(3)
    rows = [["온도", "압력", "라인", "잡음", "수율", "강도", "등급"]]
    for _ in range(90):
        t, p, n = r.uniform(150, 200), r.uniform(1, 5), r.uniform(0, 1)
        line = r.choice(["A", "B", "C"])
        bump = {"A": 0.0, "B": 25.0, "C": 50.0}[line]
        y = 2.0 * t + 4.0 * p + bump + r.uniform(-1, 1)
        rows.append([round(t, 2), round(p, 3), line, round(n, 3), round(y, 3),
                     round(3.0 * p + 0.1 * t + r.uniform(-0.2, 0.2), 3),
                     "상" if t + (30 if line == "C" else 0) > 185 else "하"])
    return _write(str(tmp / "공정.xlsx"), rows, [])


def _train(path, *extra):
    return _ok(*run_raw("ml_train.py", "-f", path, "-x", ",".join(FEATURES), *extra))


def _check_top(top, names=FEATURES):
    assert isinstance(top, list) and 1 <= len(top) <= 5, top
    shares = [e["importance"] for e in top]
    assert shares == sorted(shares, reverse=True), top
    assert all(0.0 <= s <= 1.0 for s in shares) and sum(shares) <= 1.0 + 1e-6, top
    assert all(set(e) == {"feature", "importance"} for e in top), top
    assert all(e["feature"] in names for e in top), top                 # original names, no one-hot '라인_A'
    assert len({e["feature"] for e in top}) == len(top), top


def test_regression_top_features(data):
    res = _train(data, "-y", "수율")
    top = res["top_features"]["수율"]
    _check_top(top)
    assert len(top) == 4                                               # every input (4 <= 5)
    assert abs(sum(e["importance"] for e in top) - 1.0) < 0.01         # all inputs listed: shares sum to 1
    assert top[0]["feature"] == "온도"
    assert "라인" in [e["feature"] for e in top[:3]]                    # categorical column aggregated under its name
    assert top[-1]["feature"] == "잡음"
    assert "중요 입력: 온도 " in res["preview"]
    assert "중요 입력" in res["agent_directive"]


def test_classification_top_features(data):
    res = _train(data, "-y", "등급", "--task", "classification")
    top = res["top_features"]["등급"]
    _check_top(top)
    assert top[0]["feature"] == "온도"
    assert "중요 입력: " in res["preview"]


def test_multi_target_top_features(data):
    res = _train(data, "-y", "수율,강도")
    assert set(res["top_features"]) == {"수율", "강도"}
    for t in ("수율", "강도"):
        _check_top(res["top_features"][t])
    assert res["top_features"]["강도"][0]["feature"] == "압력"
    assert "중요 입력(수율): " in res["preview"] and "중요 입력(강도): " in res["preview"]


def test_top_features_empty_with_reason():
    import ml_train
    top, note = ml_train.top_feature_shares({"feature_importances": [{"feature": "a", "importance": 0.0, "share": 0.0,
                                                                       "pct": 0.0}],
                                             "importance_note": "모든 입력의 순열 중요도가 0입니다"})
    assert top == [] and "0" in note
    assert ml_train.top_feature_shares({})[0] == []


def test_importance_failure_is_noted():
    import pandas as pd
    import ml_train

    class Broken:
        def fit(self, *a, **k):
            return self

        def predict(self, X):
            raise RuntimeError("boom")

        def score(self, X, y):
            raise RuntimeError("boom")

    X = pd.DataFrame({"a": [1.0, 2.0, 3.0], "b": [3.0, 1.0, 2.0]})
    notes = []
    ranks = ml_train.feature_importance_ranks(Broken(), X, np.array([1.0, 2.0, 3.0]),
                                              {"numeric": ["a", "b"], "categorical": {}, "datetime": {}}, 0, None,
                                              notes=notes)
    assert len(ranks) == 2 and notes and "실패" in notes[0]
    assert abs(sum(r["share"] for r in ranks) - 1.0) < 1e-6


# ---- SKILL.md option drift --------------------------------------------------------------------------------------

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
