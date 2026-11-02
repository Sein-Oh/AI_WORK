# -*- coding: utf-8 -*-
"""
2.9.12 (검토 수정): performance calculator precision for small targets, "no measured data" threshold by grid
resolution, at most 8 inputs, --keep/--inputs per target, per-target rebuild commands, case-insensitive sheet
names, --overwrite for ml_export/ml_map (and through ml_report), export warnings in the report result,
report -t targets passed to the calculator, ml_map value columns each on their own rows.
"""
import itertools
import json
import os
import sys

import numpy as np
import openpyxl
import pandas as pd
import pytest
from scipy.interpolate import RegularGridInterpolator

HERE = os.path.dirname(os.path.abspath(__file__))
KIT = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(HERE))), "CLI", "ml-kit")
sys.path.insert(0, HERE)
sys.path.insert(0, KIT)
from test_cli_smoke import run_tool, run_fail  # noqa: E402
import ml_export as mx  # noqa: E402
import ml_report as mr  # noqa: E402


# E1 ---------------------------------------------------------------------------------------------------------

def test_value_format_follows_magnitude():
    assert mx.value_format(np.array([120.0, 250.0])) == "0.00"
    assert mx.value_format(np.array([0.004, 0.005])) == "0.000000"
    assert mx.value_format(np.array([2e-6, 8e-6])) == "0.000E+00"
    assert mx._sig(0.0000123456) == 1.235e-05


def test_tiny_target_calculator_keeps_precision(tmp_path):
    r = np.random.default_rng(3)
    n = 200
    df = pd.DataFrame({"a": r.uniform(0, 1, n), "b": r.uniform(0, 1, n)})
    df["y"] = 2e-6 + 4e-6 * df.a + 2e-6 * df.b + r.normal(0, 2e-8, n)
    p = tmp_path / "tiny.csv"
    df.to_csv(p, index=False)
    v = run_tool("ml_train.py", "-f", str(p), "-x", "a,b", "-y", "y")["version_dir"]
    meta = json.load(open(os.path.join(v, "model_y.json"), encoding="utf-8"))
    assert meta["prediction_interval"]["residual_upper"] != 0                     # was rounded to 6 decimals
    res = run_tool("ml_export.py", "-m", v, "--inputs", "a,b", "--save", str(tmp_path / "tiny.xlsx"))
    t = res["targets"]["y"]
    assert t["rmse"] > 0 and t["agreement"]["agreement_pct"] > 99
    wb = openpyxl.load_workbook(res["output_file"])
    grid = np.array([row[2] for row in wb["표_y"].iter_rows(min_row=2, max_col=3, values_only=True)], dtype=float)
    assert grid.min() > 1e-6 and len(np.unique(grid)) > 100                     # not rounded to 0.000002 steps
    calc = wb["계산기"]
    cells = [c for row in calc.iter_rows() for c in row if isinstance(c.value, str) and c.value.startswith("=IF(COUNT(")]
    assert any(c.number_format == "0.000E+00" for c in cells)
    assert any('TEXT(' in str(c.value) and '"0.000E+00"' in str(c.value) for row in calc.iter_rows() for c in row)


# E2 ---------------------------------------------------------------------------------------------------------

class _SumModel:
    def predict(self, X):
        return np.asarray(X, dtype=float).sum(axis=1)


@pytest.mark.parametrize("k", [6, 8])
def test_uncovered_threshold_does_not_flag_training_rows(k):
    r = np.random.default_rng(k)
    cols = [f"x{i}" for i in range(k)]
    data = pd.DataFrame(r.uniform(0, 1, (400, k)), columns=cols)
    data.iloc[:, 0] *= 0.5                                   # no measured data for x0 > 0.5 ...
    data.iloc[:3, :] = 0.0
    data.iloc[:3, 0] = 1.0                                   # ... except one corner, so the range of x0 is 0~1
    cand = {"inputs": cols, "evals": [], "pipe": _SumModel(), "interval": None,
            "schema": {"numeric": cols, "categorical": {}, "datetime": {}, "input_columns": cols}}
    tab = mx.build_grid(cand, data)
    f = RegularGridInterpolator(tab["axes"], tab["dist"].reshape([tab["n"]] * k))
    lo, span = tab["lo"], tab["hi"] - tab["lo"]
    at_rows = f(np.clip(data.to_numpy(), lo, tab["hi"]))
    assert (at_rows > tab["uncovered_dist"]).mean() < 0.05             # 2.9.11: 75-100% of training rows flagged
    q = r.uniform(0, 1, (500, k)) * span + lo
    q[:, 0] = 0.9 + 0.1 * r.uniform(0, 1, 500)                # x0 0.9~1: inside the range, no measured rows near
    assert (f(q) > tab["uncovered_dist"]).mean() > 0.9


# E3 ---------------------------------------------------------------------------------------------------------

@pytest.fixture(scope="module")
def nine(tmp_path_factory):
    work = tmp_path_factory.mktemp("nine")
    r = np.random.default_rng(9)
    cols = [f"x{i}" for i in range(9)]
    df = pd.DataFrame(r.uniform(0, 1, (120, 9)), columns=cols)
    df["y"] = sum((i + 1) * df[c] for i, c in enumerate(cols)) + r.normal(0, 0.05, 120)
    p = work / "nine.csv"
    df.to_csv(p, index=False)
    v = run_tool("ml_train.py", "-f", str(p), "-x", ",".join(cols), "-y", "y")["version_dir"]
    return {"work": work, "v": v, "cols": cols}


def test_more_than_eight_inputs_recommends_a_set(nine):
    bad = run_fail("ml_export.py", "-m", nine["v"], "--inputs", ",".join(nine["cols"]))
    assert bad["agent_directive"].startswith("STOP_RETRY") and "[추천]" in bad["message"]
    rec = bad["message"].split("[추천]")[1].strip().split(",")
    assert 1 <= len(rec) <= 8 and set(rec) <= set(nine["cols"])


def test_automatic_search_stays_within_eight(nine):
    keep = ",".join(nine["cols"][2:])                        # 7 kept: only two candidates to try
    res = run_tool("ml_export.py", "-m", nine["v"], "--keep", keep, "--save", str(nine["work"] / "n.xlsx"))
    t = res["targets"]["y"]
    assert len(t["inputs"]) <= 8
    big = [c for c in t["candidates"] if c["count"] == 9][0]
    assert "command" not in big and big["table_rows"] is None and "note" in big
    assert any("8개" in w for w in res["warnings"])


def test_keep_over_eight_is_refused(nine):
    bad = run_fail("ml_export.py", "-m", nine["v"], "--keep", ",".join(nine["cols"]))
    assert bad["agent_directive"].startswith("STOP_RETRY") and "8개" in bad["message"]


# E4 ---------------------------------------------------------------------------------------------------------

def test_map_blank_in_other_value_column_keeps_row(tmp_path):
    A = np.arange(0, 101, 10, dtype=float)
    df = pd.DataFrame({"A": A, "B": A * 2, "C": A ** 2 / 10})
    df.loc[5, "B"] = np.nan                                  # the A=50 row has a blank B only
    p = tmp_path / "m.csv"
    df.to_csv(p, index=False)
    res = run_tool("ml_map.py", "-f", str(p), "-i", "A=50")
    assert res["points"][0]["values"]["C"] == pytest.approx(250.0)      # 2.9.11: 260 (row dropped)
    assert res["points"][0]["values"]["B"] == pytest.approx(100.0)


# E5 / E8 ----------------------------------------------------------------------------------------------------

@pytest.fixture(scope="module")
def two(tmp_path_factory):
    """Best version where ya uses a,b and yb uses a,c."""
    work = tmp_path_factory.mktemp("two")
    r = np.random.default_rng(2)
    n = 150
    df = pd.DataFrame({"a": r.uniform(0, 1, n), "b": r.uniform(0, 1, n), "c": r.uniform(0, 1, n)})
    df["ya"] = 3 * df.a + 2 * df.b + r.normal(0, 0.02, n)
    df["yb"] = 2 * df.a + 4 * df.c + r.normal(0, 0.02, n)
    p = work / "two.csv"
    df.to_csv(p, index=False)
    run_tool("ml_train.py", "-f", str(p), "-x", "a,b", "-y", "ya")
    run_tool("ml_train.py", "-f", str(p), "-x", "a,c", "-y", "yb")
    best = run_tool("ml_train.py", "-f", str(p), "--best")
    return {"work": work, "v": best["best_version_dir"]}


def test_keep_and_inputs_apply_to_targets_that_have_them(two):
    res = run_tool("ml_export.py", "-m", two["v"], "--keep", "b", "--save", str(two["work"] / "k.xlsx"))
    assert "b" in res["targets"]["ya"]["inputs"] and res["targets"]["ya"]["keep"] == ["b"]
    assert res["targets"]["yb"]["keep"] == []
    bad = run_fail("ml_export.py", "-m", two["v"], "--keep", "zzz")
    assert bad["agent_directive"].startswith("STOP_RETRY")
    # rebuild commands keep the other target's inputs (per-target --inputs), no single -y
    cmds = [c["command"] for c in res["next_commands"]]
    assert cmds and all(" -y " not in c and "ya=" in c and "yb=" in c for c in cmds)
    assert res["truncated"] is False and res["next_commands_count"] == len(cmds)
    alt = [c for c in res["targets"]["ya"]["candidates"] if "command" in c][0]
    spec = alt["command"].split("--inputs ")[1].split(" --title")[0].strip('"')
    again = run_tool("ml_export.py", "-m", two["v"], "--inputs", spec, "--save", str(two["work"] / "k2.xlsx"))
    assert again["targets"]["ya"]["inputs"] == alt["inputs"]
    assert again["targets"]["yb"]["inputs"] == res["targets"]["yb"]["inputs"]


def test_plain_inputs_list_per_target(two):
    res = run_tool("ml_export.py", "-m", two["v"], "--inputs", "a,c", "--save", str(two["work"] / "p.xlsx"))
    assert res["targets"]["ya"]["inputs"] == ["a"] and res["targets"]["yb"]["inputs"] == ["a", "c"]


# E6 ---------------------------------------------------------------------------------------------------------

def test_export_and_map_overwrite(two, tmp_path):
    save = str(two["work"] / "ow.xlsx")
    first = run_tool("ml_export.py", "-m", two["v"], "--inputs", "a", "--save", save)
    second = run_tool("ml_export.py", "-m", two["v"], "--inputs", "a", "--save", save)
    assert second["output_file"] != first["output_file"] and second.get("recovered")
    third = run_tool("ml_export.py", "-m", two["v"], "--inputs", "a", "--save", save, "--overwrite")
    assert third["output_file"] == first["output_file"] and not third.get("recovered")
    p = tmp_path / "g.csv"
    pd.DataFrame({"A": [0.0, 10, 20], "C": [1.0, 2, 3]}).to_csv(p, index=False)
    m1 = run_tool("ml_map.py", "-f", str(p), "-i", "A=5", "--save", "out.xlsx")
    m2 = run_tool("ml_map.py", "-f", str(p), "-i", "A=5", "--save", "out.xlsx", "--overwrite")
    assert m1["output_file"] == m2["output_file"]


def test_report_passes_overwrite_and_targets(monkeypatch, tmp_path):
    import ml_export
    seen = {}

    def fake(args):
        seen.update(vars(args))
        return {"output_file": "x.xlsx", "targets": {}, "verified": True, "warnings": ["w1"]}
    monkeypatch.delenv("ML_KIT_SKIP_EXPORT", raising=False)
    monkeypatch.setattr(ml_export, "execute_export", fake)
    ex = mr.make_export(str(tmp_path), "regression", None, ["생산효율"], overwrite=True)
    assert seen["overwrite"] is True and seen["targets"] == "생산효율"
    result = {"preview": "p", "agent_directive": "VERIFIED: r."}
    mr.attach_export(result, ex)                                          # E7
    assert result["export_warnings"] == ["w1"] and "export_warnings" in result["agent_directive"]


# E9 ---------------------------------------------------------------------------------------------------------

def test_sheet_names_unique_without_case():
    used = {"계산기"}
    a = mx._sheet_name("표_", "Temp", used)
    b = mx._sheet_name("표_", "TEMP", used)
    assert a.lower() != b.lower()


def test_multitable_single_group_is_plain_table():
    import ml_predict as mp
    df = pd.DataFrame({"A": [0.0, 1, 2, 3], "C": [0.0, 1, 4, 9], "__row": [2, 3, 4, 5]})
    w1, w2 = [], []
    a = mp.MultiTable(df, ["A"], ["C"], w1).at(np.array([1.5]))
    b = mp.Table(df, ["A"], ["C"], w2).at(np.array([1.5]))
    assert a == b and w1 == w2 == []
    assert list(itertools.chain(a["values"])) == ["C"]


# 예측값 반올림: 작은 값은 유효숫자 4자리 --------------------------------------------------------------------

def test_round_pred_keeps_small_values():
    import ml_predict as mp
    assert mp.round_pred(5.123456e-6) == pytest.approx(5.123e-6)
    assert mp.round_pred(123.456789) == 123.4568
    assert mp.round_pred(0.0) == 0.0
    arr = mp.round_pred(np.array([1.23456e-7, 2.5, np.nan]))
    assert arr[0] == pytest.approx(1.235e-7) and arr[1] == 2.5 and np.isnan(arr[2])
