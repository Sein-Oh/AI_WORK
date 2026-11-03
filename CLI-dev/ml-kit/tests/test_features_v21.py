# -*- coding: utf-8 -*-
"""2.1.0 train/predict features: categorical & datetime inputs, tuned linear models, fast CSV,
missing-input warnings, prediction intervals, binary threshold, bounded correlation,
next_commands, run lock, xlsx layouts and multi-target runs."""
import json
import os
import shlex
import subprocess
import sys

import numpy as np
import pandas as pd
import pytest

HERE = os.path.dirname(os.path.abspath(__file__))
DEV = os.path.dirname(HERE)  # CLI-dev\ml-kit (tests, tools)
KIT = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(HERE))), "CLI", "ml-kit")  # 실행 파일은 CLI\ml-kit
sys.path.insert(0, HERE)
from make_sample import make_sample, make_class_sample  # noqa: E402
from test_cli_smoke import run_tool, run_fail, run_raw  # noqa: E402


def _meta(work, stem, version, target):
    return json.loads((work / f"{stem}_ml_data" / version / f"model_{target}.json").read_text(encoding="utf-8"))


@pytest.fixture(scope="module")
def mixed(tmp_path_factory):
    """Numeric + categorical (라인) + datetime (일시) features, regression target."""
    work = tmp_path_factory.mktemp("mixed")
    rng = np.random.default_rng(3)
    n = 240
    dates = pd.date_range("2025-01-01", periods=n, freq="6h")
    line = rng.choice(["A라인", "B라인", "C라인"], n)
    temp = rng.uniform(200, 280, n)
    y = 0.2 * temp + pd.Series(line).map({"A라인": 0, "B라인": 6, "C라인": 12}).values \
        + 3 * (dates.hour.values >= 12) + rng.normal(0, 0.8, n)
    df = pd.DataFrame({"일시": dates.strftime("%Y-%m-%d %H:%M"), "라인": line,
                       "반응온도(°C)": temp.round(2), "수율(%)": y.round(3)})
    df.loc[5, "라인"] = None
    path = work / "mixed.csv"
    df.to_csv(path, index=False, encoding="utf-8-sig", sep=";")  # semicolon CSV: delimiter sniffing
    payload = run_tool("ml_train.py", "-f", str(path), "-x", "일시,라인,반응온도", "-y", "수율")
    return work, path, payload


def test_categorical_and_datetime_features_detected(mixed):
    work, _, payload = mixed
    ft = payload["feature_types"]
    assert ft["categorical"] == ["라인"]
    assert "일시" in ft["datetime"] and "hour" in ft["datetime"]["일시"]
    assert payload["mean_val_score"] > 0.9  # category + hour effects are learnable
    meta = _meta(work, "mixed", "v1", "수율(%)")
    assert meta["feature_schema"]["categorical"]["라인"] == ["A라인", "B라인", "C라인"]
    assert (work / "mixed_ml_data" / "v1" / "correlation.csv").is_file()


def test_predict_with_categories_dates_and_warnings(mixed):
    work, _, _ = mixed
    vdir = str(work / "mixed_ml_data" / "v1")
    ok = run_tool("ml_predict.py", "-m", vdir, "-i", "일시=2025-02-01 15:00,라인=C라인,반응온도=250")
    ps = ok["predictions_summary"]["수율(%)"]
    assert ok["unseen_category_warning"] is False and ok["missing_input_warning"] is False
    lo, hi = ps["intervals"][0]
    assert lo < ps["predictions"][0] < hi and ps["interval_level"] == 0.9

    warn = run_tool("ml_predict.py", "-m", vdir, "-i", "일시=2027-01-01 10:00,라인=Z라인,반응온도=")
    assert warn["unseen_category_warning"] is True and "Z라인" in warn["unseen_category_warnings"][0]
    assert warn["missing_input_counts"] == {"반응온도(°C)": 1}
    assert any("학습 기간" in w for w in warn["extrapolation_warnings"])


def test_predict_standalone_with_new_schema(mixed, tmp_path):
    """New-style pipelines (ColumnTransformer) must still unpickle with ml_predict.py alone."""
    import shutil
    work, _, _ = mixed
    shutil.copy(os.path.join(KIT, "ml_predict.py"), tmp_path / "ml_predict.py")
    shutil.copytree(work / "mixed_ml_data" / "v1", tmp_path / "v1")
    proc = subprocess.run([sys.executable, "-I", str(tmp_path / "ml_predict.py"), "-m", str(tmp_path / "v1"),
                           "-i", "일시=2025-02-01 15:00,라인=B라인,반응온도=240"],
                          capture_output=True, text=True, encoding="utf-8", cwd=str(tmp_path))
    assert proc.returncode == 0, proc.stdout + proc.stderr


def test_id_like_text_column_rejected(tmp_path):
    df = pd.DataFrame({"lot": [f"LOT{i:04d}" for i in range(120)], "x": np.arange(120.0), "y": np.arange(120.0) * 2})
    p = tmp_path / "ids.csv"
    df.to_csv(p, index=False)
    payload = run_fail("ml_train.py", "-f", str(p), "-x", "lot,x", "-y", "y")
    assert "ID" in payload["message"]


def test_tuned_linear_models_on_wide_linear_data(tmp_path):
    rng = np.random.default_rng(0)
    X = rng.normal(size=(200, 60))
    df = pd.DataFrame(X, columns=[f"x{i}" for i in range(60)])
    df["t"] = 2 * df["x0"] - df["x1"] + rng.normal(0, 1, 200)
    p = tmp_path / "wide.csv"
    df.to_csv(p, index=False, sep="\t")  # tab-separated
    payload = run_tool("ml_train.py", "-f", str(p), "-x", ",".join(f"x{i}" for i in range(60)), "-y", "t")
    lb = json.loads((tmp_path / "wide_ml_data" / "v1" / "leaderboard.json").read_text(encoding="utf-8"))
    board = {e["model_name"]: e for e in lb["targets"]["t"]["leaderboard"]}
    assert board["ElasticNet"]["val_r2"] > 0.75 and "alpha" in board["ElasticNet"]["params"]
    assert board["Ridge"]["val_r2"] > 0.6
    hist = json.loads((tmp_path / "wide_ml_data" / "history.json").read_text(encoding="utf-8"))[-1]
    assert hist["correlation_truncated"] is True and len(hist["correlation_matrix"]) == 31
    assert len(hist["target_correlations"]["t"]) == 60
    assert payload["next_commands"]


def test_next_command_runs(tmp_path):
    data = make_sample(str(tmp_path / "nc.csv"))
    payload = run_tool("ml_train.py", "-f", data, "-x", "반응온도,소비전력,압축압력,촉매비율,수분함량", "-y", "생산효율")
    assert payload["next_commands"], payload
    cmd = payload["next_commands"][0]["command"].replace("<새데이터.xlsx>", data)
    assert cmd.startswith("python ")
    args = shlex.split(cmd)[1:]
    proc = subprocess.run([sys.executable] + args, capture_output=True, text=True, encoding="utf-8", timeout=600)
    out = json.loads(proc.stdout.strip().splitlines()[-1])
    assert out["status"] == "ok", proc.stderr[-2000:]


def test_next_commands_flag_placeholders_and_put_runnable_first(tmp_path):
    # 실기 M8: 목표 달성 시 next_commands[0]이 <새데이터.xlsx> 자리표시자 명령이라 SKILL대로 그대로 실행하면 실패했음
    data = make_sample(str(tmp_path / "goal.csv"))
    payload = run_tool("ml_train.py", "-f", data, "-x", "반응온도,소비전력,압축압력,촉매비율,수분함량", "-y", "생산효율")
    assert payload["mean_val_score"] >= 0.9, payload
    cmds = payload["next_commands"]
    assert all(isinstance(c["needs_input"], bool) for c in cmds)
    assert cmds[0]["needs_input"] is False and "<" not in cmds[0]["command"] and "ml_report.py" in cmds[0]["command"]
    flagged = [c for c in cmds if c["needs_input"]]
    assert flagged and all(c["placeholder"] in c["command"] for c in flagged)
    assert cmds.index(flagged[0]) > 0                                   # 자리표시자 명령은 뒤에
    assert "needs_input" in payload["agent_directive"]
    assert "성능예측표" in cmds[0]["purpose"] and not any("ml_export.py" in c["command"] for c in cmds)   # 보고서와 세트
    assert "성능예측표" in payload["recommendations"][0]

    cls = make_class_sample(str(tmp_path / "cls.csv"))                  # 목표 미달: 재학습 명령은 모두 바로 실행 가능
    payload = run_tool("ml_train.py", "-f", cls, "-x", "반응온도,촉매비율,수분함량", "-y", "품질등급", "--task", "classification")
    if payload["mean_val_score"] < 0.9:
        assert payload["next_commands"] and all(c["needs_input"] is False for c in payload["next_commands"])
    assert not any("성능예측표" in c["purpose"] for c in payload["next_commands"])   # 분류는 성능예측표 없음


def test_positive_class_threshold(tmp_path):
    data = make_class_sample(str(tmp_path / "bin.csv"), n=200, imbalanced=True)
    payload = run_tool("ml_train.py", "-f", data, "-x", "반응온도,촉매비율,수분함량", "-y", "품질등급",
                       "--task", "classification", "--positive-class", "불량")
    meta = _meta(tmp_path, "bin", "v1", "품질등급")
    thr = meta["decision_threshold"]
    assert thr["positive_class"] == "불량" and 0.05 <= thr["threshold"] <= 0.95
    pred = run_tool("ml_predict.py", "-m", str(tmp_path / "bin_ml_data" / "v1"), "-i", "반응온도=275,촉매비율=1.4,수분함량=12")
    assert pred["predictions_summary"]["품질등급"]["decision_threshold"]["positive_class"] == "불량"
    assert "--positive-class" in " ".join(c["command"] for c in payload["next_commands"]) or payload["mean_val_score"] >= 0.9


def test_positive_class_validation(tmp_path):
    data = make_class_sample(str(tmp_path / "pv.csv"))
    bad = run_fail("ml_train.py", "-f", data, "-x", "반응온도,촉매비율", "-y", "품질등급",
                   "--task", "classification", "--positive-class", "A등급")
    assert "이진 분류" in bad["message"]
    bad2 = run_fail("ml_train.py", "-f", data, "-x", "반응온도,촉매비율", "-y", "불량여부",
                    "--task", "classification", "--positive-class", "예")
    assert "[후보]" in bad2["message"] and "[추천]" not in bad2["message"]     # 2.9.32: [추천] ['...'] -> [후보]


def test_run_lock_blocks_concurrent_training(tmp_path):
    sys.path.insert(0, KIT)
    import ml_train
    data = make_sample(str(tmp_path / "lk.csv"))
    lock = ml_train.RunLock(str(tmp_path / "lk_ml_data"))
    lock.acquire()
    try:
        payload = run_fail("ml_train.py", "-f", data, "-x", "반응온도,소비전력", "-y", "생산효율")
        assert "실행 중" in payload["message"]
    finally:
        lock.release()
    run_tool("ml_train.py", "-f", data, "-x", "반응온도,소비전력", "-y", "생산효율")


def test_xlsx_layout_multi_target_and_excel_dates(tmp_path):
    import openpyxl
    from datetime import datetime, timedelta
    rng = np.random.default_rng(5)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["공정 데이터 보고서"])
    ws.append([])
    ws.append(["측정일", "반응온도", "촉매비율", "생산효율", "인장강도"])
    ws.append(["", "°C", "wt%", "%", "MPa"])
    for i in range(80):
        t, c = rng.uniform(200, 280), rng.uniform(0.2, 1.5)
        ws.append([datetime(2025, 1, 1) + timedelta(days=i), round(t, 1), round(c, 2),
                   round(0.2 * t + 8 * c + rng.normal(0, 1), 2), round(3 * c + 0.01 * t + rng.normal(0, 0.2), 2)])
    p = tmp_path / "공정.xlsx"
    wb.save(p)
    payload = run_tool("ml_train.py", "-f", str(p), "--header-row", "3", "--unit-row", "4",
                       "-x", "측정일,반응온도,촉매비율", "-y", "생산효율,인장강도")
    assert payload["targets"] == ["생산효율(%)", "인장강도(MPa)"]
    assert "측정일" in payload["feature_types"]["datetime"]
    out = tmp_path / "pred.xlsx"
    pred = run_tool("ml_predict.py", "-m", payload["version_dir"], "-f", str(p),
                    "--header-row", "3", "--unit-row", "4", "-o", str(out))
    assert set(pred["targets_predicted"]) == {"생산효율(%)", "인장강도(MPa)"}
    df = pd.read_excel(out)
    assert {"pred_생산효율(%)", "pred_lo_생산효율(%)", "pred_hi_인장강도(MPa)"} <= set(df.columns)


def test_multi_target_pruning_keeps_features_needed_by_any_target(tmp_path):
    """next_commands must not drop a feature that is weak for one target but important for another."""
    rng = np.random.default_rng(1)
    n = 160
    df = pd.DataFrame({"a": rng.normal(size=n), "b": rng.normal(size=n), "noise": rng.normal(size=n)})
    df["ya"] = 3 * df["a"] + rng.normal(0, 0.3, n)   # only a matters
    df["yb"] = 3 * df["b"] + rng.normal(0, 2.5, n)   # only b matters, weaker target
    p = tmp_path / "mt.csv"
    df.to_csv(p, index=False)
    payload = run_tool("ml_train.py", "-f", str(p), "-x", "a,b,noise", "-y", "ya,yb")
    prune = [c for c in payload["next_commands"] if "저기여피처_제외" in c["command"]]
    for c in prune:
        x_arg = c["command"].split(" -x ")[1].split('"')[1]
        assert "a" in x_arg.split(",") and "b" in x_arg.split(","), c["command"]
