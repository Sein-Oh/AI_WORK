# -*- coding: utf-8 -*-
"""
2.9.14 (3차 점검 수정): grouped CV with a single-row class, heavy duplication in regression, --header-row hint,
--unit-row in the layout suggestion, commas in numeric -i values, absolute paths in suggested commands, capped
extrapolation warnings, STOP_RETRY wording, report reusing a version that holds every best, ml_map value columns
emptied by the text filter / unused columns / MultiTable range / --points layout and numbers, report number format,
single-target spec_safe, duplicate rows in the Word report.
"""
import argparse
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
import ml_export as mx  # noqa: E402


def _csv(path, df):
    df.to_csv(path, index=False, encoding="utf-8-sig")
    return str(path)


# 1. classification: a class whose rows are all identical ---------------------------------------------------

def test_class_with_one_distinct_row_falls_back(tmp_path):
    r = np.random.default_rng(0)
    df = pd.DataFrame({"a": r.normal(size=30), "b": r.normal(size=30), "cls": "B"})
    df = pd.concat([df, pd.DataFrame({"a": [5.0, 5.0], "b": [5.0, 5.0], "cls": ["A", "A"]})], ignore_index=True)
    res = run_tool("ml_train.py", "-f", _csv(tmp_path / "c1.csv", df), "-x", "a,b", "-y", "cls",
                   "--task", "classification")                                 # 2.9.13: ERROR shape mismatch
    assert any("일반 층화 교차검증" in w for w in res["warnings"])
    # multiclass: the single-row class is no longer silently never predicted (recall 0)
    df3 = pd.DataFrame({"a": r.normal(size=40), "b": r.normal(size=40), "cls": np.where(np.arange(40) % 2, "B", "C")})
    df3 = pd.concat([df3, pd.DataFrame({"a": [9.0] * 3, "b": [9.0] * 3, "cls": ["A"] * 3})], ignore_index=True)
    res3 = run_tool("ml_train.py", "-f", _csv(tmp_path / "c3.csv", df3), "-x", "a,b", "-y", "cls",
                    "--task", "classification")
    assert any("일반 층화 교차검증" in w for w in res3["warnings"])
    meta = json.load(open(os.path.join(res3["version_dir"], "model_cls.json"), encoding="utf-8"))
    lb = json.load(open(os.path.join(res3["version_dir"], "leaderboard.json"), encoding="utf-8"))
    assert lb["targets"]["cls"]["champion"]["per_class"]["A"]["recall"] > 0
    assert meta["task"] == "classification"


# 2. regression: heavy duplication ---------------------------------------------------------------------------

def test_regression_too_few_distinct_rows_stops(tmp_path):
    df = pd.DataFrame({"a": [1.0] * 8 + [2, 3, 4, 5], "b": [1.0] * 8 + [2, 1, 5, 3], "y": [1.0] * 8 + [2, 3, 4, 2]})
    bad = run_fail("ml_train.py", "-f", _csv(tmp_path / "d2.csv", df), "-x", "a,b", "-y", "y")
    assert bad["agent_directive"].startswith("STOP_RETRY") and "입력+타겟이 다른 행이 5개뿐" in bad["message"]


def test_regression_ten_distinct_rows_with_many_copies(tmp_path):
    r = np.random.default_rng(1)
    base = pd.DataFrame({"a": r.uniform(0, 1, 10), "b": r.uniform(0, 1, 10)})
    base["y"] = base.a + base.b
    df = pd.concat([base] + [base.iloc[[0]]] * 30, ignore_index=True)
    res = run_tool("ml_train.py", "-f", _csv(tmp_path / "d10.csv", df), "-x", "a,b", "-y", "y")   # never ERROR
    assert res["warnings"]


# 3. --header-row hint -------------------------------------------------------------------------------------

@pytest.fixture(scope="module")
def heads(tmp_path_factory):
    work = tmp_path_factory.mktemp("heads")
    r = np.random.default_rng(2)
    a, b, y = r.uniform(0, 1, 30), r.uniform(0, 1, 30), r.uniform(0, 1, 30)
    u = _csv(work / "u.csv", pd.DataFrame({"온도(°C)": a, "압력": b, "y": y}))
    u2 = _csv(work / "u2.csv", pd.DataFrame({"온도(C)": a, "온도(K)": b, "y": y}))
    u3 = _csv(work / "u3.csv", pd.DataFrame({"온도": a, "설비": ["B", "압력"] * 15, "y": y}))
    return {"u": u, "u2": u2, "u3": u3}


def test_header_hint_not_the_current_row(heads):
    # 2.9.33: '온도(C)' is '온도(°C)' (the same unit) and matches; a typo '온도x' is a [후보] (was [추천] for '온도(C)')
    bad = run_fail("ml_train.py", "-f", heads["u"], "-x", "온도x,압력", "-y", "y")
    assert "--header-row" not in bad["message"]                                 # 2.9.13: [추천] --header-row 1
    assert "'온도x' -> [후보] '온도(°C)'" in bad["message"]
    bad = run_fail("ml_train.py", "-f", heads["u2"], "-x", "온도", "-y", "y")
    assert "--header-row" not in bad["message"] and "[후보]" in bad["message"]
    bad = run_fail("ml_train.py", "-f", heads["u3"], "-x", "온도,압력", "-y", "y")   # '압력' is a data value
    assert "--header-row" not in bad["message"]                                 # 2.9.13: [추천] --header-row 3
    bad = run_fail("ml_map.py", "-f", heads["u"], "-i", "온도x=0.5")
    assert "--header-row" not in bad["message"] and "[후보] '온도(°C)'" in bad["message"]       # 2.9.33: 온도(C), [추천]


def test_find_header_row_rules():
    cands = [[1, ["제목"]], [3, ["온도(°C)", "압력"]]]
    assert mp.find_header_row(cands, ["온도", "압력"]) == 3
    assert mp.find_header_row(cands, ["온도(K)"]) is None                       # a different unit is not the same name
    assert mp.find_header_row(cands, ["온도(C)"]) == 3                          # 2.9.33: '°C' / 'C' is one unit (was None)
    assert mp.find_header_row([[2, ["온도(C)", "온도(K)"]]], ["온도"]) is None   # ambiguous alias


# 4. layout suggestion with a unit row ---------------------------------------------------------------------

@pytest.fixture(scope="module")
def unit_layout(tmp_path_factory):
    work = tmp_path_factory.mktemp("unitlay")
    r = np.random.default_rng(3)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["제목"])
    ws.append([None])
    ws.append(["온도", "압력", "효율"])
    ws.append(["°C", "bar", "%"])
    for _ in range(40):
        t, p = float(r.uniform(150, 200)), float(r.uniform(1, 5))
        ws.append([t, p, 0.3 * t + 2 * p + float(r.normal(0, 0.2))])
    path = work / "lay.xlsx"
    wb.save(path)
    v = run_tool("ml_train.py", "-f", str(path), "-x", "온도,압력", "-y", "효율",
                 "--header-row", "3", "--unit-row", "4")["version_dir"]
    return {"path": str(path), "v": v}


def test_layout_hint_keeps_the_unit_row(unit_layout):
    bad = run_fail("ml_predict.py", "-m", unit_layout["v"], "-f", unit_layout["path"], "--header-row", "2", "--unit-row", "4")
    rec = bad["message"].split("[추천] ")[1].strip()
    assert rec.startswith("python") and "--unit-row 4" in rec                   # 2.9.13: no --unit-row (unit row as data)
    ok = run_tool("ml_predict.py", *shlex.split(rec)[2:])
    assert ok["sample_count"] == 40 and not ok["missing_input_warning"]


def test_predict_command_absolute_quoted_paths():
    ns = argparse.Namespace(model="rel_dir", meta=None, targets=None, file="새 파일.xlsx", out=None, encoding=None,
                            overwrite=False)
    cmd = mp._predict_command(ns)
    assert mp.quote_arg(os.path.abspath("rel_dir")) in cmd and mp.quote_arg(os.path.abspath("새 파일.xlsx")) in cmd


# 5. commas in numeric -i values ----------------------------------------------------------------------------

@pytest.fixture(scope="module")
def small(tmp_path_factory):
    work = tmp_path_factory.mktemp("comma")
    r = np.random.default_rng(4)
    n = 60
    df = pd.DataFrame({"온도": r.uniform(170, 190, n), "압력": r.uniform(1000, 3500, n), "라인": r.choice(["A", "B"], n)})
    df["y"] = df.온도 * 0.5 + df.압력 * 0.01 + (df.라인 == "A") * 3 + r.normal(0, 0.3, n)
    data = _csv(work / "small.csv", df)
    return {"data": data, "v": run_tool("ml_train.py", "-f", data, "-x", "온도,압력,라인", "-y", "y")["version_dir"]}


def test_inline_comma_number_stops_with_command(small):
    bad = run_fail("ml_predict.py", "-m", small["v"], "-i", "온도=180,190,압력=3000,라인=A")
    assert bad["agent_directive"].startswith("STOP_RETRY")                      # 2.9.13: 180190, VERIFIED
    rec = bad["message"].split("[추천] ")[1].strip()
    assert rec.startswith("python") and '"온도=180190,압력=3000,라인=A"' in rec and "천 단위" in bad["message"]
    assert run_tool("ml_predict.py", *shlex.split(rec)[2:])["sample_count"] == 1
    bad = run_fail("ml_predict.py", "-m", small["v"], "-i", "온도=1,5,압력=3000,라인=A")
    assert "[추천]" not in bad["message"]
    # text values keep the joined text (not a number): unseen category, no stop
    ok = run_tool("ml_predict.py", "-m", small["v"], "-i", "온도=180,압력=3000,라인=A,B")
    assert ok["unseen_category_warning"]


def test_find_comma_number_stops_with_command(small):
    bad = run_fail("ml_predict.py", "-m", small["v"], "--find", "y=1,000", "--vary", "온도", "-i", "압력=3000,라인=A")
    assert '--find "y=1000"' in bad["message"].split("[추천] ")[1]
    bad = run_fail("ml_predict.py", "-m", small["v"], "--find", "y=120", "--vary", "온도", "-i", "압력=3,000,라인=A")
    rec = bad["message"].split("[추천] ")[1].strip()
    assert '"압력=3000,라인=A"' in rec
    assert run_tool("ml_predict.py", *shlex.split(rec)[2:])["command"] == "find"


def test_map_comma_number_stops_with_command(tmp_path):
    A = np.arange(0, 2001, 500, dtype=float)
    p = _csv(tmp_path / "g.csv", pd.DataFrame({"A": A, "Y": A * 2}))
    bad = run_fail("ml_map.py", "-f", p, "-i", "A=1,000", "-y", "Y")            # 2.9.13: '숫자가 아닙니다' without a fix
    rec = bad["message"].split("[추천] ")[1].strip()
    assert '"A=1000"' in rec
    assert run_tool("ml_map.py", *shlex.split(rec)[2:])["points"][0]["values"]["Y"] == 2000
    bad = run_fail("ml_map.py", "-f", p, "--find", "Y=1,000", "--vary", "A")
    assert '--find "Y=1000"' in bad["message"]


# 7. / 8. extrapolation warnings cap, directive wording -----------------------------------------------------

def test_extrapolation_warnings_capped(tmp_path):
    r = np.random.default_rng(5)
    cols = [f"x{i}" for i in range(25)]
    df = pd.DataFrame(r.uniform(0, 1, (40, 25)), columns=cols)
    df["y"] = df.sum(axis=1)
    v = run_tool("ml_train.py", "-f", _csv(tmp_path / "w.csv", df), "-x", ",".join(cols), "-y", "y")["version_dir"]
    res = run_tool("ml_predict.py", "-m", v, "-i", ",".join(f"{c}=5" for c in cols))
    assert "drift_warnings" not in res                                         # 2.9.13: two copies of every warning
    assert len(res["extrapolation_warnings"]) == 20 and res["extrapolation_warning_count"] == 25
    assert res["extrapolation_warnings_truncated"] is True and "25" in res["agent_directive"]


def test_stop_retry_directive_mentions_full_command():
    d = mp.error_payload(mp.InputError("x [추천] python a.py"))["agent_directive"]
    assert "full command" in d and "run it as given, once" in d


# 9. / 17. report reuses a version that already holds every best; duplicate rows in the report ---------------

def test_report_reuses_single_best_version(tmp_path):
    pytest.importorskip("docx")
    pytest.importorskip("matplotlib")
    import docx
    r = np.random.default_rng(6)
    n = 50
    df = pd.DataFrame({"a": r.uniform(0, 10, n), "b": r.uniform(0, 5, n)})
    df["y1"] = 2 * df.a + r.normal(0, 0.3, n)
    df["y2"] = df.a + 3 * df.b + r.normal(0, 0.3, n)
    df = pd.concat([df, df.iloc[:4]], ignore_index=True)                       # 4 duplicate rows
    data = _csv(tmp_path / "multi.csv", df)
    tr = run_tool("ml_train.py", "-f", data, "-x", "a,b", "-y", "y1,y2")
    rep = run_tool("ml_report.py", "-f", data)
    assert rep["made_best"] is False and rep["best_version"] == tr["version"] == "v1"   # 2.9.13: rebuilt v2
    assert not os.path.isdir(os.path.join(tr["data_root"], "v2"))
    assert rep["best_targets"]["y1"]["source_version"] == "v1"
    text = "\n".join(p.text for p in docx.Document(rep["report_file"]).paragraphs)
    assert "중복 행 4개" in text                                                 # 2.9.13: not in the report


# 10. - 14. ml_map --------------------------------------------------------------------------------------------

@pytest.fixture(scope="module")
def lines(tmp_path_factory):
    work = tmp_path_factory.mktemp("lines")
    df = pd.DataFrame({"라인": ["L1"] * 4 + ["L2"] * 4, "A": [10.0, 20, 30, 40] * 2,
                       "Y1": [20.0, 40, 60, 80] * 2, "Y2": [np.nan] * 4 + [11.0, 21, 31, 41]})
    return _csv(work / "map3.csv", df)


def test_map_value_column_empty_after_text_filter(lines):
    res = run_tool("ml_map.py", "-f", lines, "-i", "라인=L1,A=25", "-y", "Y1,Y2")  # 2.9.13: '숫자 열이 아닙니다'
    assert res["columns"] == ["Y1"] and any(w.startswith("Y2:") for w in res["warnings"])
    res = run_tool("ml_map.py", "-f", lines, "-i", "라인=L1,A=25")                # 2.9.13: Y2 dropped silently
    assert any(w.startswith("Y2:") for w in res["warnings"])
    bad = run_fail("ml_map.py", "-f", lines, "-i", "라인=L1,A=25", "-y", "Y2")
    assert "숫자 열이 아닙니다" not in bad["message"] and "Y2" in bad["message"]


def test_map_dropped_column_is_not_unused_input(tmp_path):
    df = pd.DataFrame({"A": [0.0, 7, 15, 31, 46, 60], "Y1": [1.0, 3, 2, 6, 4, 9], "Y2": [np.nan] * 5 + [1.0]})
    res = run_tool("ml_map.py", "-f", _csv(tmp_path / "s.csv", df), "-i", "A=20")
    assert res["columns"] == ["Y1"]
    assert not any("기준으로 주지 않은 열" in w for w in res.get("warnings", []))   # 2.9.13: ['Y2'] as an input


def test_multitable_range_from_kept_tables_and_out_of_range_per_group():
    A = np.arange(0, 101, 10, dtype=float)
    df = pd.DataFrame({"A": A, "Y1": np.where(A <= 60, A * 2, np.nan), "Y2": np.where(A <= 40, A + 1, np.nan),
                       "Y3": np.where(A == 100, 5.0, np.nan), "__row": np.arange(2, 13)})
    w = []
    t = mp.MultiTable(df, ["A"], ["Y1", "Y2", "Y3"], w)
    assert t.values == ["Y1", "Y2"]
    assert t.hi[0] == 60 and len(t.df) == 7                                   # 2.9.13: 100 (row of Y3 only)
    oor = t.at(np.array([80.0]))["out_of_range"]["A"]
    assert "Y1:" in oor and "Y2:" in oor                                      # 2.9.13: last group wins


@pytest.fixture(scope="module")
def grid_gap(tmp_path_factory):
    work = tmp_path_factory.mktemp("gridgap")
    rows = [(a, b) for a in (10, 20, 30, 40) for b in (1, 2, 3)]
    df = pd.DataFrame(rows, columns=["A", "B"]).astype(float)
    df["Y1"] = df.A * df.B
    df["Y2"] = df.A + df.B
    df.loc[(df.A == 20) & (df.B == 1), "Y2"] = np.nan
    data = _csv(work / "map1.csv", df)
    pts = _csv(work / "pts.csv", pd.DataFrame({"A": [15, 30, 35], "B": [1.5, 2.0, 2.5]}))
    return {"data": data, "pts": pts, "work": work}


def test_points_basis_layout_once_and_numbers(grid_gap):
    res = run_tool("ml_map.py", "-f", grid_gap["data"], "--points", grid_gap["pts"], "--save", "p.xlsx")
    sheet = pd.read_excel(res["output_file"], sheet_name="결과")
    assert "근거" not in sheet.columns and "근거_Y1" in sheet.columns            # 2.9.13: 근거 and 근거_<열> mixed
    assert sheet["근거_Y1"].notna().all()
    assert pd.api.types.is_numeric_dtype(sheet["A"]) and pd.api.types.is_numeric_dtype(sheet["B"])   # 2.9.13: text


# 15. / 16. report numbers, single-target spec_safe ----------------------------------------------------------

def test_report_fmt_no_padding():
    assert mr.fmt(0.0066) == "0.0066"                                          # 2.9.13: 0.006600
    assert mr.fmt(0.004567) == "0.004567" and mr.fmt(4e-5) == "0.00004" and mr.fmt(3.2e-6) == "3.200e-06"  # 2.9.15
    assert mr.fmt(1234.5) == "1,234.5000" and mr.fmt(0.0) == "0.0000"


def test_spec_safe_single_target_ignores_target_name():
    assert mx.spec_safe({"y=1,a": ["a", "b"]}, single=True)                    # 2.9.13: False
    assert not mx.spec_safe({"y": ["a,b"]}, single=True) and not mx.spec_safe({"y": ["a=b"]}, single=True)
    assert not mx.spec_safe({"y=1": ["a"]}) and not mx.spec_safe({"y": ["a;b"]})
