# -*- coding: utf-8 -*-
"""
2.9.11 (검토 수정): one date parser for training and prediction (YYYYMMDD), inline -i values kept as text,
named empty columns kept, data layout in predict next_commands, inline results over 5 records saved to a file,
unique header names, -o extension checked first, double quotes in next_commands names, --history verified.
"""
import argparse
import json
import os
import sys

import numpy as np
import pandas as pd
import pytest

HERE = os.path.dirname(os.path.abspath(__file__))
KIT = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(HERE))), "CLI", "ml-kit")
sys.path.insert(0, HERE)
sys.path.insert(0, KIT)
from test_cli_smoke import run_tool, run_fail  # noqa: E402
import ml_predict as mp  # noqa: E402
import ml_train as mt  # noqa: E402


@pytest.fixture(scope="module")
def coded(tmp_path_factory):
    """Numeric a, b + category code with leading zeros ('007' must not become 7)."""
    work = tmp_path_factory.mktemp("v2911")
    r = np.random.default_rng(11)
    n = 240
    codes = np.array(["007", "012", "101"])[r.integers(0, 3, n)]
    df = pd.DataFrame({"a": r.uniform(0, 10, n), "b": r.uniform(0, 5, n), "코드": codes})
    df["y"] = 3 * df.a + 2 * df.b + np.select([df.코드 == "007", df.코드 == "012"], [20.0, 0.0], -20.0) + r.normal(0, 0.2, n)
    p = work / "coded.csv"
    df.to_csv(p, index=False)
    res = run_tool("ml_train.py", "-f", str(p), "-x", "a,b,코드", "-y", "y", "--dtypes", "코드:str")
    return {"work": work, "file": str(p), "v": res["version_dir"], "train": res}


# T1 ---------------------------------------------------------------------------------------------------------

def test_parse_datetime_reads_yyyymmdd():
    s = pd.Series([20240103, "20240215", 20240301.0, "2024-04-05", "20241340", None, "abc"], dtype=object)
    got = mp.parse_datetime(s)
    assert list(got[:4]) == [pd.Timestamp("2024-01-03"), pd.Timestamp("2024-02-15"),
                             pd.Timestamp("2024-03-01"), pd.Timestamp("2024-04-05")]
    assert got[4:].isna().all()


def test_dtypes_datetime_yyyymmdd_predicts_without_imputation(tmp_path):
    r = np.random.default_rng(1)
    n = 200
    days = pd.date_range("2024-01-01", periods=n, freq="D")
    df = pd.DataFrame({"일자": days.strftime("%Y%m%d"), "a": r.uniform(0, 1, n)})
    df["y"] = 0.05 * np.arange(n) + 2 * df.a + r.normal(0, 0.05, n)
    p = tmp_path / "d.csv"
    df.to_csv(p, index=False)
    v = run_tool("ml_train.py", "-f", str(p), "-x", "일자,a", "-y", "y", "--dtypes", "일자:datetime")["version_dir"]
    meta = json.load(open(os.path.join(v, "model_y.json"), encoding="utf-8"))
    assert "일자" in meta["feature_schema"]["datetime"] and meta["datetime_stats"]["일자"]["min"].startswith("2024-01-01")
    res = run_tool("ml_predict.py", "-m", v, "-f", str(p), "-o", str(tmp_path / "pred.csv"))
    assert res["missing_input_warning"] is False                    # before 2.9.11 every date was imputed
    assert not any("날짜" in w for w in res["extrapolation_warnings"])


# T2 ---------------------------------------------------------------------------------------------------------

def test_inline_pairs_stay_text():
    df = mp.parse_inline_input("코드=007,a=1.5")
    assert df.loc[0, "코드"] == "007" and df.loc[0, "a"] == "1.5"


def test_inline_category_code_with_leading_zero(coded):
    res = run_tool("ml_predict.py", "-m", coded["v"], "-i", "a=5,b=2,코드=007")
    assert res["unseen_category_warning"] is False and res["missing_input_warning"] is False
    js = coded["work"] / "one.json"
    js.write_text(json.dumps({"a": 5, "b": 2, "코드": "007"}), encoding="utf-8")
    same = run_tool("ml_predict.py", "-m", coded["v"], "-i", str(js))
    assert res["predictions_summary"]["y"]["predictions"] == same["predictions_summary"]["y"]["predictions"]


def test_find_with_text_inline_values(coded):
    res = run_tool("ml_predict.py", "-m", coded["v"], "--find", "y=30", "--vary", "a", "-i", "b=2,코드=007")
    assert res["fixed"] == {"b": 2.0, "코드": "007"} and res["count"] >= 1


# T3 ---------------------------------------------------------------------------------------------------------

def test_named_empty_column_is_kept_and_unnamed_dropped(tmp_path):
    p = tmp_path / "e.csv"
    p.write_text("온도1,온도2,,y\n,1,,2\n,2,,4\n,3,,6\n", encoding="utf-8")
    df = mp.load_data_file(str(p))
    assert list(df.columns) == ["온도1", "온도2", "y"]
    bad = run_fail("ml_train.py", "-f", str(p), "-x", "온도1", "-y", "y")
    assert bad["agent_directive"].startswith("STOP_RETRY")
    assert "값이 하나도 없습니다" in bad["message"] and "[추천]" not in bad["message"]


def test_prediction_with_named_empty_column_imputes(coded, tmp_path):
    df = pd.read_csv(coded["file"], dtype=str).head(10)
    df["b"] = ""
    p = tmp_path / "blank_b.csv"
    df.to_csv(p, index=False)
    res = run_tool("ml_predict.py", "-m", coded["v"], "-f", str(p), "-o", str(tmp_path / "out.csv"))
    assert res["missing_input_warning"] is True and res["missing_input_counts"]["b"] == 10


# T4 ---------------------------------------------------------------------------------------------------------

def test_predict_next_commands_carry_layout(tmp_path):
    args = argparse.Namespace(header_row=3, unit_row=4, data_row=None, dtypes=None)
    cmds = mt.build_next_commands(
        args=args, raw_file=str(tmp_path / "d.xlsx"), task=mt.TASK_REGRESSION, positive_class=None,
        feature_names=["a"], created_features=[], applied_evals=[], target_names=["y"], champion={}, schema={"numeric": ["a"]},
        mean_score=0.99, staging_dir=str(tmp_path), final_dir=str(tmp_path / "v1"), data_root=str(tmp_path))
    pred = [c for c in cmds if "ml_predict.py" in c["command"]][0]
    assert pred["command"].endswith("--header-row 3 --unit-row 4")
    best = mt.build_best_next_commands(str(tmp_path / "d.xlsx"), str(tmp_path / "v2"), args)
    assert best[1]["command"].endswith("--header-row 3 --unit-row 4")


# T5 ---------------------------------------------------------------------------------------------------------

def test_inline_over_five_records_saved_to_file(coded):
    recs = [{"a": float(i), "b": 1.0, "코드": "012"} for i in range(8)]
    js = coded["work"] / "many.json"
    js.write_text(json.dumps(recs), encoding="utf-8")
    res = run_tool("ml_predict.py", "-m", coded["v"], "-i", str(js))
    assert res["truncated"] is True and res["count"] == 8 and res["output_file"]
    assert os.path.dirname(res["output_file"]) == str(coded["work"]) and res["verified"] is True
    assert len(pd.read_csv(res["output_file"], encoding="cp949")) == 8
    one = run_tool("ml_predict.py", "-m", coded["v"], "-i", "a=1,b=1,코드=012")
    assert one["truncated"] is False and one["output_file"] is None


def test_inline_json_string_over_five_saved_in_data_root(coded):
    recs = json.dumps([{"a": float(i), "b": 1.0, "코드": "012"} for i in range(6)])
    res = run_tool("ml_predict.py", "-m", coded["v"], "-i", recs)
    assert res["truncated"] is True
    assert os.path.dirname(res["output_file"]) == os.path.dirname(coded["v"].rstrip("\\/"))


# T6 ---------------------------------------------------------------------------------------------------------

def test_duplicate_headers_become_unique(tmp_path):
    p = tmp_path / "dup.csv"
    p.write_text("a,a,a_1,a\n1,2,3,4\n", encoding="utf-8")
    cols = list(mp.load_data_file(str(p)).columns)
    assert len(set(cols)) == 4 and cols[0] == "a" and cols[2] == "a_1"


# T7 ---------------------------------------------------------------------------------------------------------

def test_out_extension_checked_first(coded, tmp_path):
    bad = run_fail("ml_predict.py", "-m", coded["v"], "-f", coded["file"], "-o", str(tmp_path / "r.txt"))
    assert bad["agent_directive"].startswith("STOP_RETRY") and "[추천]" in bad["message"]
    assert bad["message"].rstrip().endswith("r.csv") and not (tmp_path / "r.txt").exists()


# T8 ---------------------------------------------------------------------------------------------------------

def test_quoted_names_go_through_files(tmp_path):
    args = argparse.Namespace(header_row=None, unit_row=None, data_row=None, dtypes=None)
    cmds = mt.build_next_commands(
        args=args, raw_file=str(tmp_path / "d.csv"), task=mt.TASK_REGRESSION, positive_class=None,
        feature_names=['온도"A', "b", "c"], created_features=[], applied_evals=[], target_names=['수율"%'],
        champion={"weak_features": ["c"], "top_features": []}, schema={"numeric": ['온도"A', "b", "c"]},
        mean_score=0.5, staging_dir=str(tmp_path), final_dir=str(tmp_path), data_root=str(tmp_path))
    cmd = cmds[0]["command"]
    assert "'" not in cmd and "next_features_pruned.txt" in cmd and "next_targets.txt" in cmd
    assert (tmp_path / "next_features_pruned.txt").read_text(encoding="utf-8").split() == ['온도"A', "b"]
    assert mt.parse_feature_or_target_list(str(tmp_path / "next_targets.txt"), "targets") == ['수율"%']


# T9 ---------------------------------------------------------------------------------------------------------

def test_history_has_verified(coded):
    res = run_tool("ml_train.py", "-f", coded["file"], "--history")
    assert res["verified"] is True and res["agent_directive"].startswith("VERIFIED")


def test_report_predict_command_layout_from_history():
    import ml_report as mr
    recs = [{"task": "regression", "data_profile": {"layout": {"header_row": 2, "unit_row": 3, "data_row": None}}}]
    ns = mr._layout_args(argparse.Namespace(header_row=None, unit_row=None, data_row=None), recs, "regression")
    assert mt.layout_flags(ns) == " --header-row 2 --unit-row 3"
    ns = mr._layout_args(argparse.Namespace(header_row=5, unit_row=None, data_row=None), recs, "regression")
    assert mt.layout_flags(ns) == " --header-row 5 --unit-row 3"
