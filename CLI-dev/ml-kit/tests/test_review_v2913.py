# -*- coding: utf-8 -*-
"""
2.9.13 (2차 점검 수정): small input values in feature_stats and --find, duplicate rows in CV, constant target,
thousands separator in -i, training layout options on a new file, two-level header hint, typo hint direction,
'20240103.0' dates, full_result_file, small numbers in the Word report, ml_map per-column basis / sparse column /
warnings / range, ml_export [추천] with every target and names that cannot be written in --inputs.
"""
import json
import os
import shlex
import sys

import numpy as np
import openpyxl
import pandas as pd
import pytest

HERE = os.path.dirname(os.path.abspath(__file__))
KIT = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(HERE))), "CLI", "ml-kit")
sys.path.insert(0, HERE)
sys.path.insert(0, KIT)
from test_cli_smoke import run_tool, run_fail  # noqa: E402
import ml_predict as mp  # noqa: E402
import ml_report as mr  # noqa: E402
import ml_train as mt  # noqa: E402
import ml_export as mx  # noqa: E402


# 1. small inputs -------------------------------------------------------------------------------------------

@pytest.fixture(scope="module")
def tiny(tmp_path_factory):
    work = tmp_path_factory.mktemp("tiny")
    r = np.random.default_rng(1)
    n = 150
    df = pd.DataFrame({"도핑농도": r.uniform(1e-6, 9e-6, n), "온도": r.uniform(0, 10, n)})
    df["저항"] = 1e6 * df["도핑농도"] + 0.5 * df["온도"] + r.normal(0, 0.05, n)
    p = work / "tiny.csv"
    df.to_csv(p, index=False)
    v = run_tool("ml_train.py", "-f", str(p), "-x", "도핑농도,온도", "-y", "저항")["version_dir"]
    return {"work": work, "v": v}


def test_small_inputs_keep_training_range(tiny):
    meta = json.load(open(os.path.join(tiny["v"], "model_저항.json"), encoding="utf-8"))
    st = meta["feature_stats"]["도핑농도"]
    assert 1e-6 <= st["min"] < st["median"] < st["max"] <= 9e-6           # 2.9.12: all 0.0
    res = run_tool("ml_predict.py", "-m", tiny["v"], "-i", "도핑농도=5e-6,온도=5")
    assert res["drift_warning"] is False                                      # 2.9.12: +500% warning on every predict
    out = run_tool("ml_predict.py", "-m", tiny["v"], "-i", "도핑농도=2e-5,온도=5")
    assert out["drift_warning"] and "0.00" not in out["extrapolation_warnings"][0]
    f = run_tool("ml_predict.py", "-m", tiny["v"], "--find", "저항=6", "--vary", "도핑농도")
    lo, hi = f["training_range"]["도핑농도"]
    assert 0 < lo < hi and f["count"] >= 1                                    # 2.9.12: [0, 0] and count 0
    assert 1e-6 < f["solutions"][0]["도핑농도"] < 9e-6
    assert f["filled_defaults"]["온도"] > 0


# 2. duplicate rows ------------------------------------------------------------------------------------------

def test_duplicate_groups_unit():
    X = pd.DataFrame({"a": [1.0, 2.0, 1.0, 3.0], "b": ["x", "y", "x", "z"]})
    g, n = mt.duplicate_groups(X, np.array([5.0, 6.0, 5.0, 7.0]))
    assert n == 1 and list(g) == [0, 1, 0, 2]
    g, n = mt.duplicate_groups(X, np.array([5.0, 6.0, 9.0, 7.0]))        # same inputs, other target: not a copy
    assert n == 0 and list(g) == [0, 1, 2, 3]


def test_duplicate_rows_do_not_leak_across_folds(tmp_path):
    r = np.random.default_rng(5)
    base = pd.DataFrame({"x": r.uniform(0, 10, 30), "z": r.uniform(0, 10, 30), "y": r.normal(0, 1, 30)})   # pure noise
    df = pd.concat([base] * 4, ignore_index=True)
    p = tmp_path / "dup.csv"
    df.to_csv(p, index=False)
    res = run_tool("ml_train.py", "-f", str(p), "-x", "x,z", "-y", "y")
    assert res["mean_val_score"] < 0.5 and res["goal_met"] is False         # 2.9.12: R2 1.0, goal_met
    assert any("중복" in w and "90" in w for w in res["warnings"])


def test_duplicate_rows_classification(tmp_path):
    r = np.random.default_rng(6)
    base = pd.DataFrame({"x": r.uniform(0, 10, 40), "z": r.uniform(0, 10, 40)})
    base["c"] = np.where(r.uniform(0, 1, 40) > 0.5, "A", "B")                  # unrelated to x, z
    df = pd.concat([base] * 3, ignore_index=True)
    p = tmp_path / "dupc.csv"
    df.to_csv(p, index=False)
    res = run_tool("ml_train.py", "-f", str(p), "-x", "x,z", "-y", "c", "--task", "classification")
    assert res["mean_val_score"] < 0.8 and any("중복" in w for w in res["warnings"])


# 3. constant target -----------------------------------------------------------------------------------------

def test_constant_regression_target_is_refused(tmp_path):
    p = tmp_path / "const.csv"
    pd.DataFrame({"x": np.arange(20, dtype=float), "y": [5.0] * 20}).to_csv(p, index=False)
    bad = run_fail("ml_train.py", "-f", str(p), "-x", "x", "-y", "y")
    assert bad["agent_directive"].startswith("STOP_RETRY") and "모두 같아(5.0)" in bad["message"]


# 4. thousands separator in -i -------------------------------------------------------------------------------

def test_inline_thousands_separator():
    df = mp.parse_inline_input("온도=3,압력=500,000,라인=A")
    assert df.iloc[0].to_dict() == {"온도": "3", "압력": "500,000", "라인": "A"}
    assert float(mp.coerce_numeric(df["압력"]).iloc[0]) == 500000.0
    with pytest.raises(mp.InputError):
        mp.parse_inline_input("500,온도=3")


# 5. / 6. layout options and two-level header ----------------------------------------------------------------

@pytest.fixture(scope="module")
def layout(tmp_path_factory):
    work = tmp_path_factory.mktemp("layout")
    r = np.random.default_rng(7)
    n = 60
    a, b = r.uniform(0, 10, n), r.uniform(0, 10, n)
    y = 2 * a + b + r.normal(0, 0.1, n)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["공정 기록"])
    ws.append(["a", "b", "y"])
    ws.append(["mm", "kg", "%"])
    for row in zip(a, b, y):
        ws.append([float(v) for v in row])
    train = work / "train.xlsx"
    wb.save(train)
    res = run_tool("ml_train.py", "-f", str(train), "-x", "a,b", "-y", "y", "--header-row", "2", "--unit-row", "3")
    new = work / "new.csv"
    pd.DataFrame({"a(mm)": [1.0, 2.0], "b(kg)": [3.0, 4.0]}).to_csv(new, index=False)
    return {"work": work, "res": res, "new": new}


def test_training_layout_on_plain_file_suggests_command_without_it(layout):
    cmd = [c for c in layout["res"]["next_commands"] if "ml_predict.py" in c["command"]][0]["command"]
    cmd = cmd.replace("<새데이터.xlsx>", str(layout["new"]))
    argv = shlex.split(cmd)[2:]                                                 # drop 'python <script>'
    bad = run_fail("ml_predict.py", *argv)
    assert bad["agent_directive"].startswith("STOP_RETRY") and "[추천] python" in bad["message"]
    rec = bad["message"].split("[추천] ")[1]
    assert "--header-row" not in rec and "--unit-row" not in rec
    ok = run_tool("ml_predict.py", *shlex.split(rec)[2:])                       # the suggestion works as given
    assert ok["sample_count"] == 2
    # a longer plain file: the layout reads a data row as the header -> model columns missing, same suggestion
    longer = layout["work"] / "longer.csv"
    pd.DataFrame({"a": np.arange(8.0), "b": np.arange(8.0)}).to_csv(longer, index=False)
    bad = run_fail("ml_predict.py", *[str(longer) if x == str(layout["new"]) else x for x in argv])
    rec = bad["message"].split("[추천] ")[1]
    assert rec.startswith("python") and "--header-row" not in rec
    assert run_tool("ml_predict.py", *shlex.split(rec)[2:])["sample_count"] == 8


def test_two_level_header_hint(tmp_path, layout):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["조건", None, "결과"])
    ws.merge_cells("A1:B1")
    ws.append(["온도(°C)", "압력(%)", "수율(%)"])
    r = np.random.default_rng(8)
    for _ in range(30):
        t, pr = float(r.uniform(100, 200)), float(r.uniform(10, 90))
        ws.append([t, pr, 0.2 * t + 0.1 * pr])
    p = tmp_path / "merged.xlsx"
    wb.save(p)
    bad = run_fail("ml_train.py", "-f", str(p), "-x", "온도,압력", "-y", "수율")
    assert "[추천] --header-row 2" in bad["message"]                              # 2.9.12: no hint
    v = run_tool("ml_train.py", "-f", str(p), "-x", "온도,압력", "-y", "수율", "--header-row", "2")["version_dir"]
    bad = run_fail("ml_predict.py", "-m", v, "-f", str(p))
    assert "[추천] --header-row 2" in bad["message"]


# 7. typo in an inline key -----------------------------------------------------------------------------------

def test_inline_typo_points_at_the_user_key(tiny):
    bad = run_fail("ml_predict.py", "-m", tiny["v"], "-i", "도핑농두=5e-6,온도=5")
    assert "-i의 '도핑농두' -> [추천] '도핑농도'" in bad["message"]            # 2.9.12: "'도핑농도' -> [추천] ['도핑농두']"


# 8. dates written as '20240103.0' ---------------------------------------------------------------------------

def test_parse_datetime_text_with_point_zero():
    out = mp.parse_datetime(pd.Series(["20240103.0", 20240103.0, "20240103", "2024.01.03"], dtype=object))
    assert all(d == pd.Timestamp("2024-01-03") for d in out)


# 9. full_result_file ----------------------------------------------------------------------------------------

def test_truncated_inline_has_full_result_file(tiny, tmp_path):
    rows = [{"도핑농도": 2e-6 + i * 1e-6, "온도": 5.0} for i in range(7)]
    p = tmp_path / "in.json"
    p.write_text(json.dumps(rows), encoding="utf-8")
    res = run_tool("ml_predict.py", "-m", tiny["v"], "-i", str(p))
    assert res["truncated"] is True and res["full_result_file"] == res["output_file"]


# 10. report numbers -----------------------------------------------------------------------------------------

def test_report_fmt_keeps_small_values():
    assert mr.fmt(0.123456) == "0.1235" and mr.fmt(1234.5) == "1,234.5000"
    assert mr.fmt(3.2e-6) == "3.200e-06"                                       # 2.9.12: 0.0000
    assert mr.fmt(0.004567) == "0.004567" and mr.fmt(0.0) == "0.0000"           # 2.9.15: 4 significant digits
    t = {"task": "regression", "oof": pd.DataFrame({"actual": [1e-6, 3e-6, 5e-6]}),
         "champ": {"champion_model": "Ridge", "val_r2": 0.95, "mae": 2.5e-7, "overfit_gap": 0.01}}
    text = mr.text_performance(t)
    assert "MAE는 2.500e-07" in text and "0.0000" not in text


# 11. - 13. ml_map MultiTable --------------------------------------------------------------------------------

def _frame():
    A = np.arange(0, 101, 10, dtype=float)
    df = pd.DataFrame({"A": A, "B": A * 2, "C": A ** 2 / 10, "D": np.nan, "__row": np.arange(2, 13)})
    df.loc[5, "B"] = np.nan                         # A=50: B blank -> B and C use different rows
    df.loc[3, "D"] = 7.0                            # D: one value only
    return df


def test_multitable_basis_per_column_and_sparse_column():
    w = []
    t = mp.MultiTable(_frame(), ["A"], ["B", "C", "D"], w)                   # 2.9.12: InputError for D
    assert t.values == ["B", "C"] and any(w_.startswith("D:") for w_ in w)
    r = t.at(np.array([50.0]))
    rows_b = {b["row"] for b in r["basis_by_column"]["B"]}
    rows_c = {b["row"] for b in r["basis_by_column"]["C"]}
    assert rows_c == {7} and rows_b == {6, 8}                                 # 2.9.12: parts[0] basis for every column
    assert len(w) == len(set(w))
    # range and nearest rows come from every key-valid row, not only the rows of B
    assert len(t.df) == 11 and t.lo[0] == 0 and t.hi[0] == 100


def test_multitable_dedupes_table_warnings():
    df = pd.DataFrame({"A": [0.0, 0, 10, 20, 30], "B": [1.0, 2, 3, 4, np.nan], "C": [1.0, 3, 5, np.nan, 7],
                       "__row": [2, 3, 4, 5, 6]})
    w = []
    mp.MultiTable(df, ["A"], ["B", "C"], w)
    assert sum("평균" in x for x in w) == 1                                    # 2.9.12: once per column group


def test_map_excel_basis_per_column(tmp_path):
    df = _frame().drop(columns=["__row", "D"])
    p = tmp_path / "m.csv"
    df.to_csv(p, index=False)
    res = run_tool("ml_map.py", "-f", str(p), "-i", "A=50", "--save", "out.xlsx")
    assert res["columns"] == ["B", "C"]
    sheet = pd.read_excel(res["output_file"], sheet_name="결과")
    assert "근거_B" in sheet.columns and "근거_C" in sheet.columns
    assert "7행" in sheet["근거_C"].iloc[0] and "6행" in sheet["근거_B"].iloc[0]


# 14. ml_export [추천] with every target -------------------------------------------------------------------

@pytest.fixture(scope="module")
def nine2(tmp_path_factory):
    work = tmp_path_factory.mktemp("nine2")
    r = np.random.default_rng(9)
    cols = [f"x{i}" for i in range(9)]
    df = pd.DataFrame(r.uniform(0, 1, (120, 9)), columns=cols)
    df["ya"] = sum((i + 1) * df[c] for i, c in enumerate(cols)) + r.normal(0, 0.05, 120)
    df["yb"] = 3 * df.x0 + 2 * df.x1 + r.normal(0, 0.05, 120)
    p = work / "nine2.csv"
    df.to_csv(p, index=False)
    v = run_tool("ml_train.py", "-f", str(p), "-x", ",".join(cols), "-y", "ya,yb")["version_dir"]
    return {"work": work, "v": v, "cols": cols}


def test_inputs_over_eight_recommends_every_target(nine2):
    spec = f"ya={','.join(nine2['cols'])};yb=x0,x1"
    bad = run_fail("ml_export.py", "-m", nine2["v"], "--inputs", spec)
    rec = bad["message"].split("[추천] ")[1].strip()
    parts = dict(p.split("=", 1) for p in rec.split(";"))
    assert parts["yb"] == "x0,x1"                                              # 2.9.12: [추천] ya=... only
    assert 1 <= len(parts["ya"].split(",")) <= 8


def test_inputs_spec_safe():
    assert mx.spec_safe({"y": ["a", "b"]})
    assert not mx.spec_safe({"y": ["a,b", "c"]}) and not mx.spec_safe({"y=1": ["a"]}) and not mx.spec_safe({"y": ["a;b"]})
