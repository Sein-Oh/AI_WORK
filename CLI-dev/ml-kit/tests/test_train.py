# -*- coding: utf-8 -*-
"""학습(ml_train.py): 피처 유형, 분류, 작은 데이터와 중복 행의 교차검증, 학습 제외 행, --eval, next_commands, --history,
최종 버전(마지막 학습), 중요 입력(top_features).
"""
import argparse
import json
import os
import random
import re
import shutil
import subprocess
import sys
import time

import numpy as np
import pandas as pd
import pytest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from make_sample import make_class_sample, make_sample  # noqa: E402
from mlkit_helpers import FEATURES, TARGET, recommendation, recommended_command, split_command, HAS_DOCX, HAS_MPL, KIT, run_fail, run_raw, run_tool, write_df, write_rows  # noqa: E402
import _ml_runtime as rt  # noqa: E402
import ml_train as mt  # noqa: E402

try:
    import docx
except ImportError:          # 보고서 검사(needs_report)는 건너뜀
    docx = None

needs_report = pytest.mark.skipif(not (HAS_DOCX and HAS_MPL), reason="python-docx/matplotlib not installed")


# ---- 피처 유형: 범주, 날짜, 넓은 선형 데이터, 여러 타겟 --------------------------------------------------


def _meta(work, stem, version, target):
    return json.loads((work / f"{stem}_ml_data" / version / f"model_{target}.json").read_text(encoding="utf-8"))


@pytest.fixture(scope="module")
def mixed(tmp_path_factory):
    """숫자 + 범주(라인) + 날짜(일시) 피처, 회귀 타겟. 세미콜론 CSV(구분자 판별)."""
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
    payload = run_tool("ml_train.py", "--file", str(path), "--features", "일시,라인,반응온도", "--targets", "수율")
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
    ok = run_tool("ml_predict.py", "--model", vdir, "--input", "일시=2025-02-01 15:00,라인=C라인,반응온도=250")
    ps = ok["predictions_summary"]["수율(%)"]
    assert ok["unseen_category_warning"] is False and ok["missing_input_warning"] is False
    lo, hi = ps["intervals"][0]
    assert lo < ps["predictions"][0] < hi and ps["interval_level"] == 0.9

    warn = run_tool("ml_predict.py", "--model", vdir, "--input", "일시=2027-01-01 10:00,라인=Z라인,반응온도=")
    assert warn["unseen_category_warning"] is True and "Z라인" in warn["unseen_category_warnings"][0]
    assert warn["missing_input_counts"] == {"반응온도(°C)": 1}
    assert any("학습 기간" in w for w in warn["extrapolation_warnings"])


def test_predict_standalone_with_new_schema(mixed, tmp_path):
    """새 형식 파이프라인(ColumnTransformer)도 ml_predict.py만으로 불러 예측한다."""
    import shutil
    work, _, _ = mixed
    for f in ("ml_predict.py", "_ml_runtime.py"):
        shutil.copy(os.path.join(KIT, f), tmp_path / f)
    shutil.copytree(work / "mixed_ml_data" / "v1", tmp_path / "v1")
    proc = subprocess.run([sys.executable, "-I", str(tmp_path / "ml_predict.py"), "--model", str(tmp_path / "v1"),
                           "--input", "일시=2025-02-01 15:00,라인=B라인,반응온도=240"],
                          capture_output=True, text=True, encoding="utf-8", cwd=str(tmp_path))
    assert proc.returncode == 0, proc.stdout + proc.stderr


def test_id_like_text_column_rejected(tmp_path):
    df = pd.DataFrame({"lot": [f"LOT{i:04d}" for i in range(120)], "x": np.arange(120.0), "y": np.arange(120.0) * 2})
    p = tmp_path / "ids.csv"
    df.to_csv(p, index=False)
    payload = run_fail("ml_train.py", "--file", str(p), "--features", "lot,x", "--targets", "y")
    assert "ID" in payload["message"]


def test_tuned_linear_models_on_wide_linear_data(tmp_path):
    rng = np.random.default_rng(0)
    X = rng.normal(size=(200, 60))
    df = pd.DataFrame(X, columns=[f"x{i}" for i in range(60)])
    df["t"] = 2 * df["x0"] - df["x1"] + rng.normal(0, 1, 200)
    p = tmp_path / "wide.csv"
    df.to_csv(p, index=False, sep="\t")  # tab-separated
    payload = run_tool("ml_train.py", "--file", str(p), "--features", ",".join(f"x{i}" for i in range(60)), "--targets", "t")
    lb = json.loads((tmp_path / "wide_ml_data" / "v1" / "leaderboard.json").read_text(encoding="utf-8"))
    board = {e["model_name"]: e for e in lb["targets"]["t"]["leaderboard"]}
    assert board["ElasticNet"]["val_r2"] > 0.75 and "alpha" in board["ElasticNet"]["params"]
    assert board["Ridge"]["val_r2"] > 0.6
    hist = json.loads((tmp_path / "wide_ml_data" / "history.json").read_text(encoding="utf-8"))[-1]
    assert hist["correlation_truncated"] is True and len(hist["correlation_matrix"]) == 31
    assert len(hist["target_correlations"]["t"]) == 60
    assert payload["next_commands"]


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
    payload = run_tool("ml_train.py", "--file", str(p), "--header-row", "3", "--unit-row", "4",
                       "--features", "측정일,반응온도,촉매비율", "--targets", "생산효율,인장강도")
    assert payload["targets"] == ["생산효율(%)", "인장강도(MPa)"]
    assert "측정일" in payload["feature_types"]["datetime"]
    out = tmp_path / "pred.xlsx"
    pred = run_tool("ml_predict.py", "--model", payload["version_dir"], "--file", str(p),
                    "--header-row", "3", "--unit-row", "4", "--save", str(out))
    assert set(pred["targets_predicted"]) == {"생산효율(%)", "인장강도(MPa)"}
    df = pd.read_excel(out)
    assert {"pred_생산효율(%)", "pred_lo_생산효율(%)", "pred_hi_인장강도(MPa)"} <= set(df.columns)


def test_multi_target_pruning_keeps_features_needed_by_any_target(tmp_path):
    """next_commands의 피처 줄이기는 한 타겟에 약해도 다른 타겟에 중요한 피처를 빼지 않는다."""
    rng = np.random.default_rng(1)
    n = 160
    df = pd.DataFrame({"a": rng.normal(size=n), "b": rng.normal(size=n), "noise": rng.normal(size=n)})
    df["ya"] = 3 * df["a"] + rng.normal(0, 0.3, n)   # only a matters
    df["yb"] = 3 * df["b"] + rng.normal(0, 2.5, n)   # only b matters, weaker target
    p = tmp_path / "mt.csv"
    df.to_csv(p, index=False)
    payload = run_tool("ml_train.py", "--file", str(p), "--features", "a,b,noise", "--targets", "ya,yb")
    prune = [c for c in payload["next_commands"] if "저기여피처_제외" in c["command"]]
    for c in prune:
        x_arg = c["command"].split(" --features ")[1].split('"')[1]
        assert "a" in x_arg.split(",") and "b" in x_arg.split(","), c["command"]


def test_positive_class_threshold(tmp_path):
    data = make_class_sample(str(tmp_path / "bin.csv"), n=200, imbalanced=True)
    payload = run_tool("ml_train.py", "--file", data, "--features", "반응온도,촉매비율,수분함량", "--targets", "품질등급",
                       "--task", "classification", "--positive-class", "불량")
    meta = _meta(tmp_path, "bin", "v1", "품질등급")
    thr = meta["decision_threshold"]
    assert thr["positive_class"] == "불량" and 0.05 <= thr["threshold"] <= 0.95
    pred = run_tool("ml_predict.py", "--model", str(tmp_path / "bin_ml_data" / "v1"), "--input", "반응온도=275,촉매비율=1.4,수분함량=12")
    assert pred["predictions_summary"]["품질등급"]["decision_threshold"]["positive_class"] == "불량"
    assert "--positive-class" in " ".join(c["command"] for c in payload["next_commands"]) or payload["mean_val_score"] >= 0.9


def test_positive_class_validation(tmp_path):
    data = make_class_sample(str(tmp_path / "pv.csv"))
    bad = run_fail("ml_train.py", "--file", data, "--features", "반응온도,촉매비율", "--targets", "품질등급",
                   "--task", "classification", "--positive-class", "A등급")
    assert "이진 분류" in bad["message"]
    bad2 = run_fail("ml_train.py", "--file", data, "--features", "반응온도,촉매비율", "--targets", "불량여부",
                    "--task", "classification", "--positive-class", "예")
    assert "[후보]" in bad2["message"] and "[추천]" not in bad2["message"]     # 2.9.32: [추천] ['...'] -> [후보]


def test_run_lock_blocks_concurrent_training(tmp_path):
    sys.path.insert(0, KIT)
    import ml_train
    data = make_sample(str(tmp_path / "lk.csv"))
    lock = ml_train.RunLock(str(tmp_path / "lk_ml_data"))
    lock.acquire()
    try:
        payload = run_fail("ml_train.py", "--file", data, "--features", "반응온도,소비전력", "--targets", "생산효율")
        assert "실행 중" in payload["message"]
    finally:
        lock.release()
    run_tool("ml_train.py", "--file", data, "--features", "반응온도,소비전력", "--targets", "생산효율")


# ---- 분류 (--task classification) ------------------------------------------------------------


FEATS = "반응온도,촉매비율,수분함량"


@pytest.fixture(scope="module")
def cls_data(tmp_path_factory):
    work = tmp_path_factory.mktemp("cls")
    return work, make_class_sample(str(work / "proc.csv"))


@pytest.fixture(scope="module")
def multiclass(cls_data):
    work, data = cls_data
    payload = run_tool("ml_train.py", "--file", data, "--features", FEATS, "--targets", "품질등급",
                       "--task", "classification", "--desc", "등급 분류", "--tag", "cls_base")
    return work, data, payload


def test_multiclass_train_contract(multiclass):
    work, _, payload = multiclass
    assert payload["task"] == "classification"
    assert payload["primary_metric"] == "val_f1_macro"
    assert 0 < payload["mean_val_score"] <= 1
    assert "mean_val_r2" not in payload
    meta = json.loads((work / "proc_ml_data" / "v1" / "model_품질등급.json").read_text(encoding="utf-8"))
    assert meta["task"] == "classification" and meta["classes"] == ["A등급", "B등급", "C등급"]
    lb = json.loads((work / "proc_ml_data" / "v1" / "leaderboard.json").read_text(encoding="utf-8"))
    champ = lb["targets"]["품질등급"]["champion"]
    assert len(lb["targets"]["품질등급"]["leaderboard"]) == 5
    assert champ["confusion_matrix"]["labels"] == ["A등급", "B등급", "C등급"]
    assert set(champ["per_class"]) == {"A등급", "B등급", "C등급"}
    oof_head = (work / "proc_ml_data" / "v1" / "oof_품질등급.csv").read_text(encoding="utf-8-sig").splitlines()[0]
    assert oof_head == "data_row,actual,predicted,proba_A등급,proba_B등급,proba_C등급"


def test_multiclass_predict_inline_probabilities(multiclass):
    work, _, _ = multiclass
    payload = run_tool("ml_predict.py", "--model", str(work / "proc_ml_data" / "v1"),
                       "--input", "반응온도=250,촉매비율=1.0,수분함량=18")
    ps = payload["predictions_summary"]["품질등급"]
    assert ps["task"] == "classification"
    assert ps["predictions"][0] in ps["classes"]
    assert abs(sum(ps["probabilities"][0].values()) - 1.0) < 1e-3
    assert "low_confidence_warning" in payload


def test_multiclass_predict_batch_columns(multiclass):
    import pandas as pd
    work, data, _ = multiclass
    out = work / "pred.csv"
    run_tool("ml_predict.py", "--model", str(work / "proc_ml_data" / "v1"), "--file", data, "--save", str(out))
    assert open(out, "rb").read(3) == b"\xef\xbb\xbf"                     # 2.9.53: UTF-8 BOM, not CP949
    df = pd.read_csv(out, encoding="utf-8-sig")
    for col in ["pred_품질등급", "conf_품질등급", "proba_품질등급_A등급", "proba_품질등급_C등급"]:
        assert col in df.columns
    assert (df["pred_품질등급"] == df["품질등급"]).mean() > 0.6


def test_binary_integer_labels(cls_data):
    work, data = cls_data
    payload = run_tool("ml_train.py", "--file", data, "--features", FEATS, "--targets", "불량여부", "--task", "classification")
    meta = json.loads((work / "proc_ml_data" / payload["version"] / "model_불량여부.json").read_text(encoding="utf-8"))
    assert meta["classes"] == ["0", "1"]  # integer labels normalized to stable strings
    lb = json.loads((work / "proc_ml_data" / payload["version"] / "leaderboard.json").read_text(encoding="utf-8"))
    assert lb["targets"]["불량여부"]["champion"]["roc_auc"] is not None


def test_imbalanced_uses_balanced_weights(tmp_path):
    data = make_class_sample(str(tmp_path / "imb.csv"), n=200, imbalanced=True)
    proc, payload = run_raw("ml_train.py", "--file", data, "--features", FEATS, "--targets", "품질등급", "--task", "classification")
    assert payload["status"] == "ok"
    assert "class_weight='balanced'" in proc.stderr
    assert any("불균형" in r for r in payload["recommendations"])


def test_string_target_without_task_hints_classification(cls_data):
    _, data = cls_data
    payload = run_fail("ml_train.py", "--file", data, "--features", FEATS, "--targets", "품질등급")
    assert "--task classification" in payload["message"]


def test_single_sample_class_rejected(tmp_path):
    import pandas as pd
    p = tmp_path / "rare.csv"
    make_class_sample(str(p))
    df = pd.read_csv(p, encoding="utf-8-sig")
    df.loc[0, "품질등급"] = "D등급"
    df.to_csv(p, index=False, encoding="utf-8-sig")
    payload = run_fail("ml_train.py", "--file", str(p), "--features", FEATS, "--targets", "품질등급", "--task", "classification")
    assert "D등급" in payload["message"]


def test_history_mixed_tasks_compare_within_task(multiclass):
    work, data, _ = multiclass
    run_tool("ml_train.py", "--file", data, "--features", "반응온도,촉매비율", "--targets", "수분함량", "--version", "v9")  # regression on same ledger
    hist = run_tool("ml_train.py", "--file", data, "--history")
    tasks = {r["version"]: r["task"] for r in hist["records"]}
    assert tasks["v1"] == "classification" and tasks["v9"] == "regression"
    # 2.9.83: 최종 버전 = 작업 유형과 상관없이 마지막 학습 (작업 유형마다의 최고 버전 best_by_task는 없앰)
    assert hist["latest_version"] == "v9" and "best_by_task" not in hist and "best_version" not in hist


@pytest.mark.report
def test_report_supports_classification(multiclass):
    """분류 버전도 통합 보고서를 만든다 (--model, --targets)."""
    pytest.importorskip("docx")
    pytest.importorskip("matplotlib")
    work, _, _ = multiclass
    payload = run_tool("ml_report.py", "--model", str(work / "proc_ml_data" / "v1"), "--targets", "품질등급")
    assert payload["task"] == "classification" and payload["targets"] == ["품질등급"]


def test_pre_2_0_model_metadata_is_regression(tmp_path):
    """'task' 키가 없는 모델 메타(2.0.0 전 형식)는 회귀로 예측한다."""
    from make_sample import make_sample
    data = make_sample(str(tmp_path / "old.csv"))
    run_tool("ml_train.py", "--file", data, "--features", "반응온도,소비전력,압축압력,촉매비율,수분함량", "--targets", "생산효율")
    meta_p = tmp_path / "old_ml_data" / "v1" / "model_생산효율(%).json"
    meta = json.loads(meta_p.read_text(encoding="utf-8"))
    for k in ("task", "primary_metric", "val_score"):
        meta.pop(k)
    meta_p.write_text(json.dumps(meta, ensure_ascii=False), encoding="utf-8")
    payload = run_tool("ml_predict.py", "--model", str(tmp_path / "old_ml_data" / "v1"),
                       "--input", "반응온도=240,소비전력=65,압축압력=3,촉매비율=0.8,수분함량=20")
    assert payload["predictions_summary"]["생산효율(%)"]["task"] == "regression"


def test_delta_compares_same_targets_only(tmp_path):
    """같은 원장의 다른 타겟은 비교 기준(delta_score)이 되지 않는다."""
    data = make_class_sample(str(tmp_path / "dt.csv"))
    run_tool("ml_train.py", "--file", data, "--features", FEATS, "--targets", "불량여부", "--task", "classification")
    second = run_tool("ml_train.py", "--file", data, "--features", FEATS, "--targets", "품질등급", "--task", "classification")
    assert second["delta_score"] == "base" and second["is_best"] is True
    third = run_tool("ml_train.py", "--file", data, "--features", FEATS, "--targets", "품질등급", "--task", "classification", "--seed", "7")
    assert isinstance(third["delta_score"], float)
    hist = run_tool("ml_train.py", "--file", data, "--history")
    assert hist["latest_version"] == "v3"                                     # 2.9.83: 마지막 학습
    assert [r["delta_score"] for r in ledger(data)] == ["base", "base", third["delta_score"]]   # v2가 v1(다른 타겟)과 비교되지 않음


# ---- 작은 데이터, 중복 행, 교차검증 ----------------------------------------------------------------------------


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
    res = run_tool("ml_train.py", "--file", str(p), "--features", "x,z", "--targets", "y")
    assert res["mean_val_score"] < 0.5 and res["goal_met"] is False         # 2.9.12: R2 1.0, goal_met
    assert any("중복" in w and "90" in w for w in res["warnings"])


def test_duplicate_rows_classification(tmp_path):
    r = np.random.default_rng(6)
    base = pd.DataFrame({"x": r.uniform(0, 10, 40), "z": r.uniform(0, 10, 40)})
    base["c"] = np.where(r.uniform(0, 1, 40) > 0.5, "A", "B")                  # unrelated to x, z
    df = pd.concat([base] * 3, ignore_index=True)
    p = tmp_path / "dupc.csv"
    df.to_csv(p, index=False)
    res = run_tool("ml_train.py", "--file", str(p), "--features", "x,z", "--targets", "c", "--task", "classification")
    assert res["mean_val_score"] < 0.8 and any("중복" in w for w in res["warnings"])


def test_constant_regression_target_is_refused(tmp_path):
    p = tmp_path / "const.csv"
    pd.DataFrame({"x": np.arange(20, dtype=float), "y": [5.0] * 20}).to_csv(p, index=False)
    bad = run_fail("ml_train.py", "--file", str(p), "--features", "x", "--targets", "y")
    assert bad["agent_directive"].startswith("STOP_RETRY") and "모두 같아(5.0)" in bad["message"]


def test_class_with_one_distinct_row_falls_back(tmp_path):
    r = np.random.default_rng(0)
    df = pd.DataFrame({"a": r.normal(size=30), "b": r.normal(size=30), "cls": "B"})
    df = pd.concat([df, pd.DataFrame({"a": [5.0, 5.0], "b": [5.0, 5.0], "cls": ["A", "A"]})], ignore_index=True)
    res = run_tool("ml_train.py", "--file", write_df(tmp_path / "c1.csv", df), "--features", "a,b", "--targets", "cls",
                   "--task", "classification")                                 # 2.9.13: ERROR shape mismatch
    assert any("일반 층화 교차검증" in w for w in res["warnings"])
    # multiclass: the single-row class is no longer silently never predicted (recall 0)
    df3 = pd.DataFrame({"a": r.normal(size=40), "b": r.normal(size=40), "cls": np.where(np.arange(40) % 2, "B", "C")})
    df3 = pd.concat([df3, pd.DataFrame({"a": [9.0] * 3, "b": [9.0] * 3, "cls": ["A"] * 3})], ignore_index=True)
    res3 = run_tool("ml_train.py", "--file", write_df(tmp_path / "c3.csv", df3), "--features", "a,b", "--targets", "cls",
                    "--task", "classification")
    assert any("일반 층화 교차검증" in w for w in res3["warnings"])
    meta = json.load(open(os.path.join(res3["version_dir"], "model_cls.json"), encoding="utf-8"))
    lb = json.load(open(os.path.join(res3["version_dir"], "leaderboard.json"), encoding="utf-8"))
    assert lb["targets"]["cls"]["champion"]["per_class"]["A"]["recall"] > 0
    assert meta["task"] == "classification"


def test_regression_too_few_distinct_rows_stops(tmp_path):
    df = pd.DataFrame({"a": [1.0] * 8 + [2, 3, 4, 5], "b": [1.0] * 8 + [2, 1, 5, 3], "y": [1.0] * 8 + [2, 3, 4, 2]})
    bad = run_fail("ml_train.py", "--file", write_df(tmp_path / "d2.csv", df), "--features", "a,b", "--targets", "y")
    assert bad["agent_directive"].startswith("STOP_RETRY") and "입력+타겟이 다른 행이 5개뿐" in bad["message"]


def test_regression_ten_distinct_rows_with_many_copies(tmp_path):
    r = np.random.default_rng(1)
    base = pd.DataFrame({"a": r.uniform(0, 1, 10), "b": r.uniform(0, 1, 10)})
    base["y"] = base.a + base.b
    df = pd.concat([base] + [base.iloc[[0]]] * 30, ignore_index=True)
    res = run_tool("ml_train.py", "--file", write_df(tmp_path / "d10.csv", df), "--features", "a,b", "--targets", "y")   # never ERROR
    assert res["warnings"]


def _schema(cols, categorical=None):
    cat = categorical or {}
    return {"numeric": [c for c in cols if c not in cat], "categorical": cat, "datetime": {}, "input_columns": list(cols)}


@pytest.fixture(autouse=False)
def quiet(monkeypatch):
    monkeypatch.setattr(mt, "log_progress", lambda p, m: None)


def _class_data(ncls, per, dup, extra=0):
    counts = [per] * ncls
    counts[0] += extra
    y = sum([[f"c{i}"] * c for i, c in enumerate(counts)], [])
    n = len(y)
    if dup:   # identical (input, class) pairs
        y = [y[min((i // 2) * 2, n - 1)] for i in range(n)]
        X = pd.DataFrame({"a": np.repeat(np.arange((n + 1) // 2), 2)[:n].astype(float), "b": np.zeros(n)})
    else:
        X = pd.DataFrame({"a": np.arange(n, dtype=float), "b": np.random.default_rng(n).normal(size=n)})
    return X, pd.Series(y)


@pytest.mark.parametrize("ncls,per,dup", [(3, 4, True), (4, 3, False), (4, 3, True), (5, 2, False), (5, 2, True),
                                          (5, 3, False), (5, 4, True), (2, 5, True)])
def test_tiny_classification_shapes(quiet, ncls, per, dup):
    X, y = _class_data(ncls, per, dup)
    res, pipe, champ, _ = mt.train_and_benchmark_classifiers(X, y, "c", _schema(["a", "b"]))   # 2.9.14: ValueError n_splits=3
    assert len(res["leaderboard"]) == 5 and pipe.predict(X.iloc[:2]).shape == (2,)


@pytest.mark.parametrize("n,big", [(10, 0), (12, 3), (10, 10), (10, 20)])
def test_tiny_regression_shapes(quiet, n, big):
    idx = np.array(list(range(n)) + [0] * big)
    X = pd.DataFrame({"a": idx.astype(float), "b": (idx % 3).astype(float)})
    res, pipe, _, _ = mt.train_and_benchmark_models(X, pd.Series(idx * 2.0 + 1), "y", _schema(["a", "b"]))
    assert len(res["leaderboard"]) == 5


def test_sparse_column_empty_in_a_training_fold(quiet):
    a = np.full(10, np.nan)
    a[:2] = [1.0, 2.0]
    X = pd.DataFrame({"a": a})
    mt.train_and_benchmark_models(X, pd.Series(np.arange(10, dtype=float)), "y", _schema(["a"]))   # 2.9.14: 0 features


def test_fit_safe_model_settings():
    lr = mt.get_classification_models()["LogisticRegression"]
    assert mt.fit_safe_model(lr, np.array(["a"] * 5 + ["b"] * 5)).cv == 3                       # large enough: unchanged
    two = mt.fit_safe_model(lr, np.array(["a"] * 5 + ["b"] * 2))
    assert two.cv.get_n_splits() == 2
    one = mt.fit_safe_model(lr, np.array(["a"] * 5 + ["b"]))
    assert type(one).__name__ == "LogisticRegression"
    knn = mt.get_classification_models()["KNeighbors"]
    assert mt.fit_safe_model(knn, np.array(["a", "b", "a"])).n_neighbors == 3
    en = mt.get_regression_models()["ElasticNet"]
    assert mt.fit_safe_model(en, np.arange(5.0)).cv == 2 and mt.fit_safe_model(en, np.arange(9.0)).cv == 3


def test_tiny_multiclass_cli(tmp_path):
    df = pd.DataFrame({"a": range(10), "b": [0, 1, 2] * 3 + [0], "c": list("ppqqrrsstt")})
    res = run_tool("ml_train.py", "--file", write_df(tmp_path / "c5x2.csv", df), "--features", "a,b", "--targets", "c", "--task", "classification")
    assert res["agent_directive"].startswith("VERIFIED")


# ---- 학습 제외 행 ('제외' 열과 자동 이상치, 데이터 파일은 바꾸지 않음) --------------------------------


N = 200


def _excl_frame(marks=None, spikes=True):
    r = np.random.default_rng(0)
    t = r.uniform(170, 195, N)
    p = r.uniform(2.5, 3.5, N)
    c = r.uniform(2, 3, N)
    y = 0.5 * t + 8 * p - 3 * c + r.normal(0, 1, N)
    if spikes:
        y[5] += 60          # target outlier -> automatic
        y[17] -= 50         # target outlier, but marked '포함'
        t[30] = 400         # feature outlier -> automatic
    df = pd.DataFrame({"반응온도": t, "압축압력": p, "촉매비율": c, "생산효율": y,
                       "등급": np.where(y > np.median(y), "양품", "불량")})
    if marks is not None:
        df["제외"] = marks
    return df


def _excl_marks():
    m = [""] * N
    m[50] = "센서오류"
    m[17] = "포함"
    return m


def test_marks_auto_outliers_and_all_rows_score(tmp_path):
    f = tmp_path / "d.xlsx"
    _excl_frame(_excl_marks()).to_excel(f, index=False)
    before = f.read_bytes()
    res = run_tool("ml_train.py", "--file", str(f), "--features", "반응온도,압축압력,촉매비율,제외", "--targets", "생산효율")
    ex = res["exclusions"]["생산효율"]
    assert (ex["manual"], ex["auto"], ex["forced_include"]) == (1, 2, 1)
    assert isinstance(ex["all_rows_score"], float) and ex["all_rows_score"] < res["mean_val_score"]
    assert "제외" not in res["feature_types"]["numeric"]                 # never a feature
    assert "검증 R2(제외 행 포함)" in res["preview"] and "검증 R2=" in res["preview"]
    assert f.read_bytes() == before                                       # data file untouched

    raw = open(res["excluded_rows_file"], "rb").read()
    assert raw.startswith(b"\xef\xbb\xbf")                                # 2.9.53: UTF-8 BOM (Excel, LLM read it)
    rows = pd.read_csv(res["excluded_rows_file"], encoding="utf-8-sig")
    # 2.9.53: the JSON itself carries the rows and the reasons (the LLM never opens the CSV)
    prev = {p["row"]: p for p in res["excluded_preview"]}
    assert set(prev) == {7, 32, 52} and res["excluded_count"] == 3
    assert prev[52]["kind"] == "직접" and prev[52]["reason"] == "센서오류" and prev[52]["target"] == "생산효율"
    assert prev[7]["kind"] == "자동" and "생산효율" in prev[7]["reason"] and isinstance(prev[7]["value"], float)
    assert res["excluded_reasons"] == {"센서오류": 1, "자동 이상치: 생산효율": 1, "자동 이상치: 반응온도": 1}
    assert "excluded_preview" in res["agent_directive"]
    by_row = {int(r["원본 행"]): r for _, r in rows.iterrows()}
    assert set(by_row) == {7, 32, 52}                                     # Excel rows (header is row 1)
    assert by_row[52]["구분"] == "직접 표시" and by_row[52]["이유"] == "센서오류"
    assert by_row[7]["구분"] == "자동 이상치" and "생산효율" in by_row[7]["이유"]
    assert "반응온도 400" in by_row[32]["이유"]

    meta = json.load(open(os.path.join(res["version_dir"], "model_생산효율.json"), encoding="utf-8"))
    assert meta["feature_stats"]["반응온도"]["max"] < 200                 # training range ignores the outlier
    assert meta["excluded"]["auto"] == 2
    assert meta["target_summary"]["max"] < 150
    hist = [json.loads(ln) for ln in open(os.path.join(res["data_root"], "history.jsonl"), encoding="utf-8")]
    assert hist[-1]["exclusions"]["생산효율"]["manual"] == 1
    assert hist[-1]["data_profile"]["rows_used"]["생산효율"] == N - 3


# ---- --exclude-where: 조건으로 학습에서 뺄 행 (2.9.56, office-cli excel.py read --where와 같은 식) ----------------

WHERE = "생산효율 > 112; 라인 = B | 라인 = C"


def _where_frame():
    df = _excl_frame(_excl_marks())
    df["라인"] = ["A", "B"] * (N // 2)
    df.loc[17, "라인"] = "C"            # matches the condition but is marked '포함' -> kept, with a warning
    df.loc[50, "라인"] = "C"            # matches the condition and is marked '센서오류' -> the user's reason wins
    df = df.rename(columns={"생산효율": "생산효율(%)"})
    return df


def _where_expected(df):
    y = df["생산효율(%)"]
    hit = ((y > 112) & (df["라인"] == "B")) | (df["라인"] == "C")
    return hit & (df["제외"] == "")


@pytest.fixture(scope="module")
def where_run(tmp_path_factory):
    work = tmp_path_factory.mktemp("where")
    f = work / "d.xlsx"
    df = _where_frame()
    df.to_excel(f, index=False)
    res = run_tool("ml_train.py", "--file", str(f), "--features", "반응온도,압축압력,촉매비율", "--targets", "생산효율",
                   "--exclude-where", WHERE)
    return work, f, df, res


def test_exclude_where_leaves_out_matching_rows(where_run):
    _, f, df, res = where_run
    want = _where_expected(df)
    assert 0 < want.sum() < 60
    ex = res["exclusions"]["생산효율(%)"]
    assert ex["condition"] == int(want.sum()) and ex["manual"] == 1, ex
    text = "조건: 생산효율(%) > 112; 라인 = B | 라인 = C"                    # column names resolved (unit added)
    assert res["excluded_reasons"][text] == int(want.sum())
    rows = {p["row"] for p in res["excluded_preview"] if p["kind"] == "조건"}
    assert rows <= {i + 2 for i in np.flatnonzero(want.to_numpy())} and rows
    assert any("포함" in w and "--exclude-where" in w for w in res["warnings"]), res.get("warnings")
    hist = [json.loads(ln) for ln in open(os.path.join(res["data_root"], "history.jsonl"), encoding="utf-8")]
    assert hist[-1]["exclude_where"] == "생산효율(%) > 112; 라인 = B | 라인 = C"
    snap = pd.read_csv(os.path.join(res["version_dir"], "model_dataset.csv"), encoding="utf-8-sig")
    assert int(snap["학습제외_생산효율(%)"].sum()) == ex["manual"] + ex["auto"] + ex["condition"]


def test_exclude_where_in_retrain_commands(tmp_path):
    args = argparse.Namespace(header_row=None, unit_row=None, data_row=None, dtypes=None)
    cmds = mt.build_next_commands(
        args=args, raw_file=str(tmp_path / "d.xlsx"), task=mt.TASK_REGRESSION, positive_class=None,
        feature_names=["a", "b", "c"], created_features=[], applied_evals=[], target_names=["y"],
        champion={"weak_features": ["c"], "top_features": [{"feature": "a"}, {"feature": "b"}]},
        schema={"numeric": ["a", "b", "c"]}, mean_score=0.5, staging_dir=str(tmp_path), final_dir=str(tmp_path / "v1"),
        exclude_where="a > 1; 라인 = 'B|C'")
    retrain = [c["command"] for c in cmds if "ml_train.py" in c["command"] and not c["needs_input"]]   # 틀은 '<앞 조건>; <조건>'
    assert len(retrain) == 2 and all(split_command(c)[split_command(c).index("--exclude-where") + 1] == "a > 1; 라인 = 'B|C'"
                                     for c in retrain), retrain


def test_exclude_where_errors(tmp_path):
    f = tmp_path / "d.csv"
    _where_frame().to_csv(f, index=False)
    base = ["--file", str(f), "--features", "반응온도,압축압력,촉매비율", "--targets", "생산효율"]
    bad = run_fail("ml_train.py", *base, "--exclude-where", "생산효율 > 0")             # every row
    assert bad["agent_directive"].startswith("STOP_RETRY") and "1행만 남습니다" in bad["message"], bad   # the '포함' row
    bad = run_fail("ml_train.py", *base, "--exclude-where", "라인x = B")               # unknown column -> [후보]
    assert "[후보]" in bad["message"] and "라인" in bad["message"], bad
    bad = run_fail("ml_train.py", *base, "--exclude-where", "라인 == B")               # fixable notation -> [추천] command
    rec = bad["message"].split("[추천] ")[1]
    assert rec.startswith("python ") and "--exclude-where '라인 = B'" in rec, rec
    bad = run_fail("ml_train.py", *base, "--exclude-where", "생산효율 >")              # unreadable
    assert "--exclude-where" in bad["message"] and "형식" in bad["message"], bad


def test_where_mask_rules():
    df = pd.DataFrame({"a": ["1,500", "20", "", "x"], "b": ["Line A", "line a", "B", "A;B"],
                       "d": ["2026-01-02", "2026-02-01", None, "2025-12-31"]})
    g = rt.parse_where("a >= 20; b = line a | b ~ ';' | d < 2026-01-01")
    assert rt.where_text(g) == "a >= 20; b = line a | b ~ ';' | d < 2026-01-01"
    assert list(rt.where_mask(df, g)) == [True, True, False, True]
    assert list(rt.where_mask(df, rt.parse_where("d < 2026-01-01"))) == [False, False, False, True]
    assert list(rt.where_mask(df, rt.parse_where("a > 100"))) == [True, False, False, False]   # '1,500' is 1500
    assert list(rt.where_mask(df, rt.parse_where("b != B"))) == [True, True, False, True]


def test_where_words_inside_value_are_kept():
    """2.9.59: 값 안의 '포함'·'and'·'or'까지 바꿔 틀린 [추천](비고 ~ ...)을 냈다."""
    g = rt.parse_where("비고 = 불량 포함 시료")
    assert g == [[("비고", "=", "불량 포함 시료")]]
    assert rt.parse_where("비고 = A and B | 라인 = C") == [[("비고", "=", "A and B")], [("라인", "=", "C")]]
    with pytest.raises(rt.InputError) as e:                     # 진짜 표기 오류는 [추천], 값은 작은따옴표로
        rt.parse_where("비고 == 불량 포함 시료 and 라인 = A")
    msg = str(e.value)
    assert "작은따옴표" in msg and msg.split("[추천] ")[1] == "비고 = '불량 포함 시료'; 라인 = A", msg
    assert rt.parse_where(msg.split("[추천] ")[1]) == [[("비고", "=", "불량 포함 시료"), ("라인", "=", "A")]]
    with pytest.raises(rt.InputError) as e:                     # 열 다음의 '포함'은 연산자
        rt.parse_where("비고 포함 불량 and 생산효율 > 3")
    assert str(e.value).split("[추천] ")[1] == "비고 ~ 불량; 생산효율 > 3"


def test_where_not_equal_keeps_blank_and_size_value_checked():
    df = pd.DataFrame({"c": [1.0, np.nan, 3.0], "t": ["A", "", "B"], "d": ["2026-01-01", None, "2026-01-02"]})
    assert list(rt.where_mask(df, rt.parse_where("c != 1"))) == [False, False, True]     # 빈 칸은 빼지 않음
    assert list(rt.where_mask(df, rt.parse_where("t != A"))) == [False, False, True]
    assert list(rt.where_mask(df, rt.parse_where("d != 2026-01-01"))) == [False, False, True]
    with pytest.raises(rt.InputError) as e:
        rt.parse_where("c > 9 0")
    assert str(e.value).split("[추천] ")[1] == "c > 90"
    with pytest.raises(rt.InputError) as e:
        rt.parse_where("c >= 높음")
    assert "[추천]" not in str(e.value) and "숫자" in str(e.value)


def test_where_code_text_column_compares_as_text():
    df = pd.DataFrame({"k": ["01", "1", "02"], "n": [1, 2, 1]})
    assert list(rt.where_mask(df, rt.parse_where("k = 01"))) == [True, False, False]     # 글자 열: 글자 비교
    assert list(rt.where_mask(df, rt.parse_where("k != 01"))) == [False, True, True]
    assert list(rt.where_mask(df, rt.parse_where("n = 01"))) == [True, False, True]      # 숫자 열: 숫자 비교


def _marked_pressure_frame(n=60):
    """숫자 열(압력 1.0/2.0/3.0)에 결측 표시 '측정불가' 한 칸."""
    r = np.random.default_rng(5)
    p = r.choice([1.0, 2.0, 3.0], n)
    t = r.uniform(150, 200, n).round(1)
    df = pd.DataFrame({"온도": t, "압력": p.astype(object), "생산량": (2 * t + 10 * p + r.normal(0, 1, n)).round(2)})
    df.loc[4, "압력"] = "측정불가"
    return df


def test_where_number_column_with_missing_mark_compares_numbers(tmp_path):
    """2.9.61: 숫자 열에 '측정불가' 한 칸만 있어도 글자 열로 보아 '압력 = 2'가 '2.0'과 안 맞고 '!= 2'가 전부 맞았다.
    행 수는 2.9.59와 같다 (결측 표시 칸은 '!='에 맞음, 빈 칸은 맞지 않음)."""
    df = _marked_pressure_frame()
    f = write_df(tmp_path / "w.csv", df)
    raw = rt.load_data_file(f)
    num = pd.to_numeric(df["압력"], errors="coerce")
    eq = int((num == 2).sum())
    assert 0 < eq < len(df)
    for cond, want in (("압력 = 2", eq), ("압력 != 2", len(df) - eq),
                       ("압력 != 2; 온도 > 180", int(((num != 2) & (df["온도"] > 180)).sum())),
                       ("압력 = 2.0", eq), ("압력 = 02", eq), ("압력 = 측정불가", 1)):
        assert int(rt.where_mask(raw, rt.parse_where(cond)).sum()) == want, cond
    res = run_tool("ml_train.py", "--file", f, "--features", "온도,압력", "--targets", "생산량", "--exclude-where", "압력 = 2")
    assert res["exclusions"]["생산량"]["condition"] == eq, res["exclusions"]


def test_where_thousands_and_code_columns():
    df = pd.DataFrame({"a": ["1,000", "1000", "측정불가", "", "2,000.5"], "c": ["007", "7", "12", "007", None]})
    assert list(rt.where_mask(df, rt.parse_where("a = 1000"))) == [True, True, False, False, False]
    assert list(rt.where_mask(df, rt.parse_where("a = 1,000"))) == [True, True, False, False, False]
    assert list(rt.where_mask(df, rt.parse_where("a != 1000"))) == [False, False, True, False, True]
    assert list(rt.where_mask(df, rt.parse_where("c = 7"))) == [False, True, False, False, False]     # 코드 열: 글자
    assert list(rt.where_mask(df, rt.parse_where("c = 007"))) == [True, False, False, True, False]
    assert list(rt.where_mask(df, rt.parse_where("c != 7"))) == [True, False, True, True, False]
    mix = pd.DataFrame({"c": ["01", "001", "02", "002"]})      # 5차 검토: '001'이 '01'까지 빼지 않게
    assert list(rt.where_mask(mix, rt.parse_where("c = 001"))) == [False, True, False, False]
    assert list(rt.where_mask(mix, rt.parse_where("c = 01"))) == [True, False, False, False]


@pytest.mark.parametrize("cells,code", [
    (["05", "10", "15", "20", "05"], False),          # 3차 검토: 두 자리 숫자 열 ('05' 하나)
    (["5", "10", "15", "20"], False),
    (["01", "02", "03", "10", "11"], False),          # 4차 검토: 앞자리 0 칸이 많아도 숫자 열 (월 01~12)
    (["007", "7", "12", "13", "14", "15"], True),     # (a) 같은 숫자의 0 없는 표기
    (["01", "001", "02", "002"], True),               # 5차 검토: (a) 같은 숫자를 0 개수만 다르게 적음
    (["01", "011", "012", "02", "021"], False),       # 같은 수의 다른 표기 없음 (자릿수만 섞임)
    (["001", "050", "120", "100"], True),             # (b) 3자리 이상 고정폭 코드
    (["01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12"], False),
    (["05", "10", "측정불가", "", None, "20"], False),  # 빈 값 표시는 세지 않음
])
def test_code_column_judgment(cells, code):
    assert rt.is_code_column(pd.Series(cells, dtype=object)) is code


def test_where_on_code_column_warns_other_notation(tmp_path):
    """3차 검토(2.9.63): 코드 열에서 '월 = 1'이 '01' 행을 말없이 빼지 않았다(학습은 둘 다 숫자 1). warnings와 그 표기를 포함한
    조건의 [추천] 전체 명령(next_commands[0])을 준다."""
    r = np.random.default_rng(5)
    months = ["01", "02", "03", "1", "2", "3", "10", "11", "12"] * 6
    x = r.uniform(0, 10, len(months)).round(3)
    df = pd.DataFrame({"월": months, "온도": x, "수율": (3 * x + r.normal(0, 0.1, len(months))).round(3)})
    f = write_df(tmp_path / "mo.csv", df)
    raw = rt.load_data_file(f)
    g = rt.parse_where("월 = 1")
    assert rt.is_code_column(raw["월"]) and int(rt.where_mask(raw, g).sum()) == 6
    variants = rt.where_notation_variants(raw, g)
    assert variants == [(("월", "=", "1"), {"01": 6})]
    assert rt.where_text(rt.where_with_variants(g, variants)) == "월 = 1 | 월 = 01"
    neg = rt.parse_where("월 != 1; 온도 > 3")
    assert rt.where_text(rt.where_with_variants(neg, rt.where_notation_variants(raw, neg))) == "월 != 1; 온도 > 3; 월 != 01"
    res = run_tool("ml_train.py", "--file", f, "--features", "월,온도", "--targets", "수율", "--exclude-where", "월 = 1")
    assert res["exclusions"]["수율"]["condition"] == 6
    w = [x for x in res.get("warnings", []) if "'01' 6행" in x]
    assert w and "빼지 않음" in w[0] and "[추천] python " in w[0], res.get("warnings")
    cmd = res["next_commands"][0]["command"]
    assert "--exclude-where '월 = 1 | 월 = 01'" in cmd and cmd.count("--exclude-where") == 1, cmd
    proc = subprocess.run([sys.executable] + split_command(cmd)[1:], capture_output=True, text=True, encoding="utf-8")
    again = json.loads(proc.stdout)
    assert again["exclusions"]["수율"]["condition"] == 12 and not any("'01'" in x for x in again.get("warnings", []))


def test_where_on_zero_padded_month_column_compares_numbers(tmp_path):
    """4차 검토: 월 '01'~'12' 열을 앞자리 0 칸이 절반 이상이라 코드 열로 보아 '월 = 1'이 0행이었다 (학습은 숫자로 읽음).
    학습에서 숫자로 읽는 열은 = / != 도 숫자 비교."""
    months = [f"{i:02d}" for i in range(1, 13)] * 2
    df = pd.DataFrame({"월": months, "y": np.arange(24.0)})
    raw = rt.load_data_file(write_df(tmp_path / "mon.csv", df))
    assert not rt.is_code_column(raw["월"])
    for cond, want in (("월 = 1", 2), ("월 = 01", 2), ("월 != 1", 22), ("월 > 6", 12)):
        g = rt.parse_where(cond)
        assert int(rt.where_mask(raw, g).sum()) == want, cond
        assert rt.where_notation_variants(raw, g) == []


def test_where_on_two_digit_number_column_compares_numbers(tmp_path):
    """'05, 10, 15, 20' 열은 숫자 열: '온도 = 5'가 '05' 행에 맞는다 (경고 없음)."""
    df = pd.DataFrame({"온도": ["05", "10", "15", "20"] * 5, "y": np.arange(20.0)})
    f = write_df(tmp_path / "n.csv", df)
    raw = rt.load_data_file(f)
    g = rt.parse_where("온도 = 5")
    assert not rt.is_code_column(raw["온도"]) and int(rt.where_mask(raw, g).sum()) == 5
    assert rt.where_notation_variants(raw, g) == []


def test_no_exclude_column_still_finds_outliers_and_clean_data_has_no_key(tmp_path):
    f = tmp_path / "d.csv"
    _excl_frame().to_csv(f, index=False)
    res = run_tool("ml_train.py", "--file", str(f), "--features", "반응온도,압축압력,촉매비율", "--targets", "생산효율")
    assert res["exclusions"]["생산효율"]["auto"] == 3 and res["exclusions"]["생산효율"]["column"] is None

    g = tmp_path / "clean.csv"
    _excl_frame(spikes=False).to_csv(g, index=False)
    res = run_tool("ml_train.py", "--file", str(g), "--features", "반응온도,압축압력,촉매비율", "--targets", "생산효율")
    assert "exclusions" not in res and "excluded_rows_file" not in res


def test_skewed_distribution_is_not_trimmed():
    """후보가 10%를 넘으면 넓은 분포이지 이상치가 아니다: 아무 행도 빼지 않는다."""
    r = np.random.default_rng(1)
    y = pd.Series(np.concatenate([r.normal(0, 1, 160), r.normal(30, 1, 40)]))
    X = pd.DataFrame({"a": r.normal(0, 1, 200)})
    plan = mt.plan_exclusions(None, y, X, {"numeric": ["a"]}, [], mt.TASK_REGRESSION, "y")
    assert not plan["mask"].any() and "자동 제외하지 않았습니다" in plan["note"]


def test_classification_keeps_every_class(tmp_path):
    """자동 이상치는 어떤 클래스도 학습에서 없애지 않는다."""
    y = pd.Series(["A"] * 98 + ["B", "B"])
    x = np.r_[np.random.default_rng(2).normal(0, 1, 98), 50.0, 60.0]   # both 'B' rows look like outliers
    plan = mt.plan_exclusions(None, y, pd.DataFrame({"a": x}), {"numeric": ["a"]}, [], mt.TASK_CLASSIFICATION, "y")
    assert not plan["mask"][98:].any()

    f = tmp_path / "c.xlsx"
    _excl_frame(_excl_marks()).to_excel(f, index=False)
    res = run_tool("ml_train.py", "--file", str(f), "--features", "반응온도,압축압력,촉매비율", "--targets", "등급",
                   "--task", "classification", "--positive-class", "불량")
    ex = res["exclusions"]["등급"]
    assert ex["manual"] == 1 and ex["auto"] == 1                          # target is a label: features only
    assert 0 <= ex["all_rows_score"] <= 1 and "검증 Macro-F1(제외 행 포함)" in res["preview"]


def test_exclude_column_as_target_is_input_error(tmp_path):
    f = tmp_path / "d.csv"
    _excl_frame(_excl_marks()).to_csv(f, index=False)
    payload = run_fail("ml_train.py", "--file", str(f), "--features", "반응온도", "--targets", "제외")
    assert payload["agent_directive"].startswith("STOP_RETRY") and "제외" in payload["message"]


def test_exclude_column_zero_and_false_are_not_marks():
    import ml_train as mt
    assert [mt._exclude_mark(v) for v in (0, 0.0, "0", "FALSE", False, None, float("nan"), "", " 센서오류 ", 1, "포함")] == \
        ["", "", "", "", "", "", "", "", "센서오류", "1", "포함"]


@pytest.mark.report
def test_report_shows_terms_and_excluded_rows(tmp_path):
    from docx import Document
    f = tmp_path / "d.xlsx"
    _excl_frame(_excl_marks()).to_excel(f, index=False)
    res = run_tool("ml_train.py", "--file", str(f), "--features", "반응온도,압축압력,촉매비율", "--targets", "생산효율")
    rep = run_tool("ml_report.py", "--model", res["version_dir"])
    doc = Document(rep["report_file"])
    text = "\n".join(p.text for p in doc.paragraphs)
    cells = " ".join(c.text for tb in doc.tables for row in tb.rows for c in row.cells)
    assert "검증 R² (제외 행 미포함)" in cells and "검증 R² (제외 행 포함)" in cells
    assert "학습 제외 3행 (직접 표시 1, 자동 이상치 2)" in text and "excluded_rows.csv" in text
    terms = next(p.text for p in doc.paragraphs if p.text.startswith("판정 기준:")).split("\n")   # one term per line
    assert terms[1] == "점수 용어" and terms[2].startswith("- 학습 R²: 모델이 배운 행")
    assert terms[3].startswith("- 검증 R²: 모델이 배우지 않은 행") and terms[5].startswith("- 검증 R² (제외 행 포함)")
    assert "제외 행 포함 전체" not in text + cells and "교차검증 R²는" not in text


# ---- --eval 파생 변수 ----------------------------------------------------------------------------------------


def test_compile_formula_alias_and_canonical():
    cols = ["반응온도(°C)", "소비전력(kW)", "log(x)"]
    tok, mapping, canon = rt.compile_formula("log(소비전력) / 반응온도(°C)", cols)
    assert canon == "log(소비전력(kW)) / 반응온도(°C)"
    assert set(mapping.values()) == {"소비전력(kW)", "반응온도(°C)"}


def test_eval_blocks_attribute_access():
    df = pd.DataFrame({"a": [1.0, 2.0]})
    with pytest.raises(rt.InputError):
        rt.apply_eval_expressions(df, ["b = a.__class__"])


# 파생 변수의 원본 열이 --features에 없어도(--eval "온도상한 = clip(온도, -1e12, 250)") 모델이 그 열을 입력으로 기록해
# 예측, --find, 성능예측표가 그 열을 받는다 (input_columns_required, model_dataset.csv, feature_stats).


@pytest.fixture(scope="module")
def capped(tmp_path_factory):
    d = tmp_path_factory.mktemp("cap")
    r = np.random.default_rng(9)
    n = 300
    df = pd.DataFrame({"온도(°C)": r.uniform(200, 300, n), "압력": r.uniform(1, 3, n)})
    df["효율"] = 0.3 * np.minimum(df["온도(°C)"], 250) + 5 * df["압력"] + r.normal(0, 0.2, n)
    path = d / "c.csv"
    df.to_csv(path, index=False)
    res = run_tool("ml_train.py", "--file", str(path), "--features", "압력", "--eval", "온도상한 = clip(온도, -1e12, 250)", "--targets", "효율")
    return d, res["version_dir"]


def test_source_column_is_recorded(capped):
    _, vdir = capped
    meta = json.load(open(os.path.join(vdir, "model_효율.json"), encoding="utf-8"))
    assert meta["input_columns_required"] == ["압력", "온도(°C)"] and "온도(°C)" in meta["feature_stats"]
    snap = pd.read_csv(os.path.join(vdir, "model_dataset.csv"), encoding="utf-8-sig")
    assert "온도(°C)" in snap.columns and "온도상한" in snap.columns


def test_prediction_uses_source_and_caps(capped):
    _, vdir = capped
    hi = run_tool("ml_predict.py", "--model", vdir, "--input", "온도=280,압력=2")["predictions_summary"]["효율"]["predictions"][0]
    cap = run_tool("ml_predict.py", "--model", vdir, "--input", "온도=250,압력=2")["predictions_summary"]["효율"]["predictions"][0]
    assert hi == pytest.approx(cap, abs=1e-6)                                   # above 250 is treated as 250
    miss = run_fail("ml_predict.py", "--model", vdir, "--input", "압력=2")
    assert "'온도(°C)'" in miss["message"] and miss["missing_inputs"] == ["온도(°C)"]    # the source input, not '온도상한'
    assert miss["error_type"] == "needs_user_decision"                                # 2.9.54: '원본 입력' only


@pytest.mark.report
def test_find_and_export_work_on_the_source(capped, tmp_path):
    _, vdir = capped
    f = run_tool("ml_predict.py", "--model", vdir, "--find", "효율=80", "--vary", "온도", "--input", "압력=2")
    assert f["count"] >= 1 and abs(f["solutions"][0]["온도(°C)"] - (80 - 10) / 0.3) < 2
    e = run_tool("ml_export.py", "--model", vdir, "--save", str(tmp_path / "c.xlsx"))
    assert e["targets"]["효율"]["inputs"] == ["온도(°C)", "압력"] and e["targets"]["효율"]["check"]["max_diff"] < 1e-9


# ---- next_commands --------------------------------------------------------------------------------------


@pytest.fixture(scope="module")
def sample_model(tmp_path_factory):
    """make_sample 데이터의 입력 5개 -> 생산효율 학습 (next_commands 검사 둘이 같이 씀)."""
    data = make_sample(str(tmp_path_factory.mktemp("nc") / "nc.csv"))
    payload = run_tool("ml_train.py", "--file", data, "--features", "반응온도,소비전력,압축압력,촉매비율,수분함량", "--targets", "생산효율")
    return data, payload


def test_next_command_runs(sample_model):
    data, payload = sample_model
    assert payload["next_commands"], payload
    cmd = payload["next_commands"][0]["command"].replace("<새데이터.xlsx>", data)
    assert cmd.startswith("python ")
    args = split_command(cmd)[1:]
    proc = subprocess.run([sys.executable] + args, capture_output=True, text=True, encoding="utf-8", timeout=600)
    out = json.loads(proc.stdout.strip().splitlines()[-1])
    assert out["status"] == "ok", proc.stderr[-2000:]


def test_next_commands_flag_placeholders_and_put_runnable_first(sample_model, tmp_path):
    # 실기 M8: 목표 달성 시 next_commands[0]이 <새데이터.xlsx> 자리표시자 명령이라 SKILL대로 그대로 실행하면 실패했음
    _, payload = sample_model
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
    payload = run_tool("ml_train.py", "--file", cls, "--features", "반응온도,촉매비율,수분함량", "--targets", "품질등급", "--task", "classification")
    if payload["mean_val_score"] < 0.9:                                  # 재학습 명령은 바로 실행 가능, 한 건 예측 틀만 맨 뒤 (2.9.75)
        cmds = payload["next_commands"]
        ready = [c for c in cmds if c["needs_input"] is False]                     # 틀(한 건 예측, 행 빼고 다시 학습)은 맨 뒤
        assert cmds and cmds[:len(ready)] == ready and len(cmds) - len(ready) == 2, cmds
        assert " --input " in cmds[len(ready)]["command"] and "<값>" in cmds[len(ready)]["command"], cmds
        assert cmds[-1]["placeholder"] == "<조건>", cmds
    assert not any("성능예측표" in c["purpose"] for c in payload["next_commands"])   # 분류는 성능예측표 없음


def test_next_commands_give_one_case_prediction_with_input_names(sample_model):
    """2.9.75 집 PC opencode: 학습 결과에 --file 틀과 재학습 '--eval "a = b * c"'만 있어 Gemma 4가 한 건 예측을 '--eval "온도=182; ..."'로,
    solar는 한 건을 파일로 만들려 했음 -> 모델 입력 이름의 '--input "이름=<값>,..."' 틀(맨 뒤). <값>을 채우면 그대로 실행됨."""
    data, payload = sample_model
    one = [c for c in payload["next_commands"] if " --input " in c["command"]]
    assert len(one) == 1 and one[0]["placeholder"] == "<값>", payload["next_commands"]
    assert all(c["needs_input"] for c in payload["next_commands"][payload["next_commands"].index(one[0]):])   # 틀은 뒤에
    cmd = split_command(one[0]["command"])
    pairs = cmd[cmd.index("--input") + 1]
    assert [p.split("(")[0] for p in pairs.split(",")] == ["반응온도", "소비전력", "압축압력", "촉매비율", "수분함량"], pairs
    assert "반응온도(°C)=<값>" in pairs                                       # 모델의 실제 열 이름 (단위 포함)
    filled = [a.replace("<값>", "1") for a in cmd[2:]]
    assert run_tool("ml_predict.py", *filled)["status"] == "ok"


def test_predict_next_commands_have_no_layout(tmp_path):
    args = argparse.Namespace(header_row=3, unit_row=4, data_row=None, dtypes=None)
    cmds = mt.build_next_commands(
        args=args, raw_file=str(tmp_path / "d.xlsx"), task=mt.TASK_REGRESSION, positive_class=None,
        feature_names=["a"], created_features=[], applied_evals=[], target_names=["y"], champion={}, schema={"numeric": ["a"]},
        mean_score=0.99, staging_dir=str(tmp_path), final_dir=str(tmp_path / "v1"))
    pred = [c for c in cmds if "ml_predict.py" in c["command"]][0]
    assert pred["command"].endswith("--file '<새데이터.xlsx>'")                     # 2.9.53: --header-row 3 --unit-row 4


def test_quoted_names_go_through_files(tmp_path):
    args = argparse.Namespace(header_row=None, unit_row=None, data_row=None, dtypes=None)
    cmds = mt.build_next_commands(
        args=args, raw_file=str(tmp_path / "d.csv"), task=mt.TASK_REGRESSION, positive_class=None,
        feature_names=['온도"A', "b", "c"], created_features=[], applied_evals=[], target_names=['수율"%'],
        champion={"weak_features": ["c"], "top_features": []}, schema={"numeric": ['온도"A', "b", "c"]},
        mean_score=0.5, staging_dir=str(tmp_path), final_dir=str(tmp_path))
    cmd = cmds[0]["command"]
    assert "'" not in cmd and "next_features_pruned.txt" in cmd and "next_targets.txt" in cmd
    assert (tmp_path / "next_features_pruned.txt").read_text(encoding="utf-8").split() == ['온도"A', "b"]
    assert mt.parse_feature_or_target_list(str(tmp_path / "next_targets.txt"), "targets") == ['수율"%']


# ---- --history ----------------------------------------------------------------------------------------------


def test_history_has_verified(coded):
    res = run_tool("ml_train.py", "--file", coded["file"], "--history")
    assert res["verified"] is True and res["agent_directive"].startswith("VERIFIED")


def test_history_latest_version_is_the_last_run(tmp_path):
    """2.9.83: --history의 latest_version은 마지막 학습 (점수가 더 낮고 타겟이 줄었어도). 최고 버전 키는 없음."""
    r = np.random.default_rng(7)
    n = 80
    df = pd.DataFrame({"a": r.uniform(0, 1, n), "b": r.uniform(0, 1, n), "z": r.uniform(0, 1, n)})
    df["A"] = 5 * df.a + r.normal(0, 0.05, n)
    df["B"] = 3 * df.b + r.normal(0, 0.05, n)
    data = write_df(tmp_path / "h.csv", df)
    run_tool("ml_train.py", "--file", data, "--features", "a,b", "--targets", "A,B")
    v2 = run_tool("ml_train.py", "--file", data, "--features", "z", "--targets", "A")                  # later, worse, one target
    hist = run_tool("ml_train.py", "--file", data, "--history")
    assert hist["latest_version"] == "v2" and hist["latest_version_dir"] == v2["version_dir"]     # 2.9.82: best_version v1
    for gone in ("best_version", "best_version_dir", "best_by_task", "best_version_note"):
        assert gone not in hist, gone
    assert not os.path.exists(os.path.join(v2["data_root"], "best_version.json"))


# ---- 최종 버전 (2.9.83): 마지막 학습으로 보고서, --best는 없음 -------------------------------------------


Y1, Y2, Y3, LABEL = "수율(%)", "강도(MPa)", "두께(mm)", "등급"


def make_multi(path, n=60, seed=3):
    """입력 3개, 원인이 서로 다른 숫자 타겟 3개, 클래스 라벨 1개."""
    rng = np.random.default_rng(seed)
    a = rng.uniform(0, 10, n)
    b = rng.uniform(0, 5, n)
    c = rng.uniform(1, 3, n)
    df = pd.DataFrame({
        "온도(°C)": a.round(3), "압력(bar)": b.round(3), "속도(m/s)": c.round(3),
        Y1: (3 * a + rng.normal(0, 1.0, n)).round(3),          # depends on 온도 only
        Y2: (2 * b * c + rng.normal(0, 0.5, n)).round(3),       # interaction 압력 x 속도
        Y3: (a + b + rng.normal(0, 0.3, n)).round(3),
        LABEL: np.where(a > 5, "A", "B"),
    })
    df.to_csv(path, index=False, encoding="utf-8-sig")
    return path


def ledger(data):
    root = os.path.splitext(data)[0] + "_ml_data"
    with open(os.path.join(root, "history.jsonl"), encoding="utf-8") as f:
        return [json.loads(x) for x in f if x.strip()]


def meta(version_dir, target):
    from _ml_runtime import sanitize_filename  # noqa: E402
    with open(os.path.join(version_dir, f"model_{sanitize_filename(target)}.json"), encoding="utf-8") as f:
        return json.load(f)


def test_best_option_is_gone(tmp_path):
    """2.9.83: --best(타겟마다 이력의 최고 조건으로 다시 학습)를 없앰. 주면 인자 오류이고 학습하지 않는다."""
    data = make_multi(str(tmp_path / "b.csv"))
    run_tool("ml_train.py", "--file", data, "--features", "온도", "--targets", "수율")
    proc, err = run_raw("ml_train.py", "--file", data, "--best")
    assert proc.returncode == 2 and err["status"] == "error", err
    assert [r["version"] for r in ledger(data)] == ["v1"]
    assert sorted(os.listdir(os.path.splitext(data)[0] + "_ml_data")) == [".train.lock", "history.json", "history.jsonl", "v1"]


@needs_report
@pytest.mark.report
def test_report_file_uses_the_last_run_of_any_task(tmp_path):
    """회귀 다음에 분류를 학습하면 보고서 --file은 분류(v2) 보고서 1개. 이력에만 있는 회귀 타겟은 warnings로 알림."""
    data = make_multi(str(tmp_path / "mix.csv"))
    run_tool("ml_train.py", "--file", data, "--features", "온도,압력", "--targets", "수율")
    run_tool("ml_train.py", "--file", data, "--features", "온도,속도", "--targets", "등급", "--task", "classification")
    rep = run_tool("ml_report.py", "--file", data)
    assert rep["version"] == "v2" and rep["task"] == "classification" and rep["targets"] == [LABEL], rep
    assert os.path.isfile(rep["report_file"]) and "report_count" not in rep
    assert any(Y1 in w and "v2" in w for w in rep["warnings"]), rep.get("warnings")
    assert "warnings" in rep["agent_directive"]
    assert [r["version"] for r in ledger(data)] == ["v1", "v2"]                                  # 다시 학습하지 않음


@pytest.mark.report
def test_no_history_report_stops(tmp_path):
    err = run_fail("ml_report.py", "--file", make_multi(str(tmp_path / "empty.csv")))
    assert err["agent_directive"].startswith("STOP_RETRY") and "학습 결과 폴더가 없습니다" in err["message"]
    assert "ml_train.py" in err["message"]


@needs_report
@pytest.mark.report
def test_report_m_path_unchanged(tmp_path):
    data = make_multi(str(tmp_path / "single.csv"))
    tr = run_tool("ml_train.py", "--file", data, "--features", "온도,압력,속도", "--targets", "수율,두께")
    run_tool("ml_train.py", "--file", data, "--features", "온도", "--targets", "수율")              # later run: --model ignores it
    rep = run_tool("ml_report.py", "--model", tr["version_dir"])
    assert rep["version"] == tr["version"] == "v1" and set(rep["targets"]) == {Y1, Y3}
    for key in ("best_version", "made_best", "best_reports", "best_targets", "best_version_dir", "warnings"):
        assert key not in rep, key
    rec = ledger(data)[0]
    assert "best" not in rec and "best_source" not in rec


# ---- 중요 입력 top_features ----------------------------------------------------------------------------


# 결과 JSON의 top_features: 타겟마다 최고 모델의 순열 중요도를 비중(합 1)으로 큰 순서 5개까지. 글자 열은 원래 이름 하나.
# 중요도에 정보가 없으면 [] + warnings.


TOP_X = ["온도", "압력", "라인", "잡음"]


def _ok_pr(proc, res):
    assert proc.returncode == 0 and res["status"] == "ok", (res, proc.stderr[-2000:])
    return res


@pytest.fixture(scope="module")
def top_data(tmp_path_factory):
    tmp = tmp_path_factory.mktemp("top")
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
    return write_rows(str(tmp / "공정.xlsx"), rows, [])


def _train_top(path, *extra):
    return _ok_pr(*run_raw("ml_train.py", "--file", path, "--features", ",".join(TOP_X), *extra))


def _check_top(top, names=TOP_X):
    assert isinstance(top, list) and 1 <= len(top) <= 5, top
    shares = [e["importance"] for e in top]
    assert shares == sorted(shares, reverse=True), top
    assert all(0.0 <= s <= 1.0 for s in shares) and sum(shares) <= 1.0 + 1e-6, top
    assert all(set(e) == {"feature", "importance"} for e in top), top
    assert all(e["feature"] in names for e in top), top                 # original names, no one-hot '라인_A'
    assert len({e["feature"] for e in top}) == len(top), top


def test_regression_top_features(top_data):
    res = _train_top(top_data, "--targets", "수율")
    top = res["top_features"]["수율"]
    _check_top(top)
    assert len(top) == 4                                               # every input (4 <= 5)
    assert abs(sum(e["importance"] for e in top) - 1.0) < 0.01         # all inputs listed: shares sum to 1
    assert top[0]["feature"] == "온도"
    assert "라인" in [e["feature"] for e in top[:3]]                    # categorical column aggregated under its name
    assert top[-1]["feature"] == "잡음"
    assert "중요 입력: 온도 " in res["preview"]
    assert "중요 입력" in res["agent_directive"]


def test_classification_top_features(top_data):
    res = _train_top(top_data, "--targets", "등급", "--task", "classification")
    top = res["top_features"]["등급"]
    _check_top(top)
    assert top[0]["feature"] == "온도"
    assert "중요 입력: " in res["preview"]


def test_multi_target_top_features(top_data):
    res = _train_top(top_data, "--targets", "수율,강도")
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


# ---- --dtypes 이름과 형식 확인 (저장소 지침 반영, 2.9.47) --------------------------------------------------------


def test_resolve_dtypes_names_and_types():
    cols = ["일시", "설비코드(번호)", "온도(°C)", "수율"]
    assert mt.resolve_dtypes("설비코드:str, 온도:float", cols) == "설비코드(번호):str,온도(°C):float"
    assert mt.resolve_dtypes("일시:date,수율:int64", cols) == "일시:datetime,수율:int"
    assert mt.resolve_dtypes(None, cols) is None and mt.resolve_dtypes(" ", cols) is None
    with pytest.raises(rt.InputError) as e:
        mt.resolve_dtypes("설비코드:strr,일시:datetime", cols)
    assert '[추천] --dtypes 설비코드:str,일시:datetime' in str(e.value)        # 함수 직접 호출: 옵션과 값 (명령으로 실행하면 전체 명령)
    with pytest.raises(rt.InputError) as e:
        mt.resolve_dtypes("설비코드:텍스트", cols)
    assert "[추천]" not in str(e.value) and "float, int, str, datetime" in str(e.value)
    with pytest.raises(rt.InputError) as e:
        mt.resolve_dtypes("설비코드", cols)
    assert "형식이 없습니다" in str(e.value)
    with pytest.raises(rt.ColumnNameError) as e:
        mt.resolve_dtypes("설비코두:str", cols)                    # a close spelling is asked, never auto-picked
    assert "[후보]" in str(e.value) and "설비코드(번호)" in str(e.value)


def test_train_dtypes_unitless_name_applies(tmp_path):
    rows = [["설비코드(번호)", "온도", "수율"]] + [[1000 + i % 4, 150 + i, 0.3 * (150 + i) + (i % 4)] for i in range(40)]
    data = write_rows(str(tmp_path / "d.csv"), rows, [])
    res = run_tool("ml_train.py", "--file", data, "--features", "설비코드(번호),온도", "--targets", "수율", "--dtypes", "설비코드:str")
    assert "설비코드(번호)" in res["feature_types"]["categorical"], res["feature_types"]    # 2.9.46: ignored (numeric)
    err = run_fail("ml_train.py", "--file", data, "--features", "설비코드(번호),온도", "--targets", "수율", "--dtypes", "설비코드:strr")
    assert err["agent_directive"].startswith("STOP_RETRY") and recommended_command(err["message"])[1][-2:] == ["--dtypes", "설비코드:str"], err


# ---- JSON 목록 상한 + truncated (R2, 2.9.49) ---------------------------------------------------------------------


def test_history_records_capped(tmp_path):
    root = tmp_path / "h_ml_data"
    root.mkdir()
    recs = [{"version": f"v{i}", "task": "regression", "mean_val_score": 0.5 + i / 100, "targets": ["y"]}
            for i in range(1, 36)]
    (root / "history.jsonl").write_text("".join(json.dumps(r) + "\n" for r in recs), encoding="utf-8")
    res = mt.print_history_summary(str(root))                                     # 2.9.48: all 35 records
    assert res["record_count"] == 35 and res["count"] == 35 and res["shown"] == mt.HISTORY_SHOWN
    assert res["truncated"] is True and res["full_result_file"].endswith("history.json")
    assert [r["version"] for r in res["records"]][-1] == "v35" and res["records"][0]["version"] == "v6"
    assert res["latest_version"] is None          # 기록만 있고 버전 폴더가 없음: 보고서·예측이 쓸 최종 버전 없음 (2.9.83)
    (root / "history.jsonl").write_text("".join(json.dumps(r) + "\n" for r in recs[:3]), encoding="utf-8")
    res = mt.print_history_summary(str(root))
    assert res["truncated"] is False and res["shown"] == 3 and "full_result_file" not in res


def test_feature_types_capped():
    ft = {"numeric": [f"x{i}" for i in range(40)], "categorical": ["c"], "datetime": {f"d{i}": ["year"] for i in range(35)}}
    out, cut = mt.capped_feature_types(ft)
    assert cut and len(out["numeric"]) == mt.FEATURE_TYPES_SHOWN and out["categorical"] == ["c"]
    assert len(out["datetime"]) == mt.FEATURE_TYPES_SHOWN and isinstance(out["datetime"], dict)
    assert mt.capped_feature_types({"numeric": ["a"], "categorical": [], "datetime": {}}) == (
        {"numeric": ["a"], "categorical": [], "datetime": {}}, False)


def test_message_column_list_capped():
    cols = [f"열{i}" for i in range(100)]
    assert rt.name_list(cols[:3]) == "['열0', '열1', '열2']"
    text = rt.name_list(cols)
    assert "열39" in text and "'열40'" not in text and "외 60개 (전체 100개)" in text
    with pytest.raises(rt.ColumnNameError) as e:                                # 2.9.48: all 100 names
        rt.resolve_column_names(["없는열zz"], cols, "피처(--features)")
    assert "외 60개" in str(e.value)


# ---- 학습도 나눠서 진행 (R10, 2.9.57): 끝낸 타겟은 .resume_train에 저장, 같은 명령을 다시 실행하면 이어서 ----------------

SPLIT_X = "반응온도(°C),소비전력(kW),압축압력(bar)"
SPLIT_Y = "생산효율(%),수분함량(%),촉매비율(wt%)"


def test_training_stops_at_time_limit_and_the_same_command_continues(tmp_path, monkeypatch):
    ref = make_sample(str(tmp_path / "ref.csv"))
    want = run_tool("ml_train.py", "--file", ref, "--features", SPLIT_X, "--targets", SPLIT_Y)         # no limit: one run
    data = make_sample(str(tmp_path / "s.csv"))
    root = tmp_path / "s_ml_data"
    monkeypatch.setenv("ML_KIT_TIME_LIMIT", "30.5")         # 0.5 s after the reserve: one new target per run
    runs = []
    for _ in range(80):
        proc, out = run_raw("ml_train.py", "--file", data, "--features", SPLIT_X, "--targets", SPLIT_Y, "--desc", "나눠 학습")
        assert proc.returncode == 0 and out["status"] == "ok", proc.stderr
        runs.append(out)
        if not out.get("continue"):
            break
        # nothing is written twice: no version folder, no ledger record until the run that ends
        assert not (root / "history.jsonl").exists() and not (root / "v1").exists()
        assert "[WARN] 비정상 중단" not in proc.stderr
    progress = [r.get("progress") or "" for r in runs]
    assert len(runs) >= 3 and all(r.get("continue") for r in runs[:-1]), progress
    first, last = runs[0], runs[-1]
    assert first["agent_directive"].startswith("CONTINUE:") and first["verified"] is False
    # every target is split by candidate model (2.9.60) and the run stops before the next one
    assert "후보 모델" in first["progress"] and "5개 끝남" in first["progress"], progress
    assert any("'수분함량(%)'" in t for t in progress) and any("'촉매비율(wt%)'" in t for t in progress), progress
    cmd = first["next_commands"][0]
    assert cmd.startswith("python ") and "ml_train.py" in cmd and "--desc" in cmd
    assert last["version"] == "v1" and last["agent_directive"].startswith("VERIFIED")
    assert last["resumed_targets"] == ["생산효율(%)", "수분함량(%)"]
    assert last["mean_val_score"] == pytest.approx(want["mean_val_score"], abs=1e-9)   # the same models as one run
    _assert_same_version(want["version_dir"], last["version_dir"])       # 2.9.66: OOF and predictions too
    hist = [json.loads(ln) for ln in open(root / "history.jsonl", encoding="utf-8")]
    assert [h["version"] for h in hist] == ["v1"]
    assert not (root / ".resume_train").exists()                                       # cleaned after the commit
    assert sorted(os.listdir(root)) == [".train.lock", "history.json", "history.jsonl", "v1"]


def _assert_same_version(a, b):
    """Two version folders trained from the same data and command (one run / split runs, 2.9.66): the same
    leaderboard, OOF files (byte for byte) and final model predictions."""
    import joblib
    lb = [json.load(open(os.path.join(d, "leaderboard.json"), encoding="utf-8")) for d in (a, b)]
    for x in lb:
        x.pop("timestamp", None)
    assert lb[0] == lb[1]
    names = sorted(n for n in os.listdir(a) if n.startswith(("oof_", "model_")) and n.endswith((".csv", ".joblib")))
    assert names and names == sorted(n for n in os.listdir(b) if n.startswith(("oof_", "model_"))
                                     and n.endswith((".csv", ".joblib")))
    frame = pd.read_csv(os.path.join(a, "model_dataset.csv"), encoding="utf-8-sig")
    for n in names:
        if n.endswith(".csv"):
            assert open(os.path.join(a, n), "rb").read() == open(os.path.join(b, n), "rb").read(), n
            continue
        pa, pb = joblib.load(os.path.join(a, n)), joblib.load(os.path.join(b, n))
        X = frame[list(pa.feature_names_in_)]
        assert np.array_equal(pa.predict(X), pb.predict(X)), n


def _split_runs(monkeypatch, *argv, limit=40, seconds="30.5"):
    monkeypatch.setenv("ML_KIT_TIME_LIMIT", seconds)        # 0.5 s after the reserve: one new step per run
    runs = []
    for _ in range(limit):
        proc, out = run_raw("ml_train.py", *argv)
        assert proc.returncode == 0 and out["status"] == "ok", proc.stderr
        runs.append(out)
        if not out.get("continue"):
            break
    assert all(r.get("continue") for r in runs[:-1]) and not runs[-1].get("continue"), [r.get("progress") for r in runs]
    return runs


def test_single_target_is_split_by_candidate_model(tmp_path, monkeypatch):
    """2.9.59: 타겟이 1개면 나눌 곳이 없어 오래 걸리면 JSON 없이 끊겼다. 이제 후보 모델마다 저장하고 멈춘다."""
    ref = make_sample(str(tmp_path / "ref.csv"))
    want = run_tool("ml_train.py", "--file", ref, "--features", SPLIT_X, "--targets", "생산효율(%)")
    data = make_sample(str(tmp_path / "one.csv"))
    root = tmp_path / "one_ml_data"
    runs = _split_runs(monkeypatch, "--file", data, "--features", SPLIT_X, "--targets", "생산효율(%)", limit=40)
    assert len(runs) >= 3, [r.get("progress") for r in runs]
    assert "후보 모델" in runs[0]["progress"] and "5개 끝남" in runs[0]["progress"], runs[0]["progress"]
    # the final fit is estimated by the champion's own time (2.9.65): it may go on in the run of the last candidate
    assert all("후보 모델" in r["progress"] or "1위 모델" in r["progress"] or "순열 중요도" in r["progress"]
               for r in runs[:-1]), [r.get("progress") for r in runs]
    last = runs[-1]
    assert last["version"] == "v1" and last["agent_directive"].startswith("VERIFIED")
    assert last["mean_val_score"] == pytest.approx(want["mean_val_score"], abs=1e-9)   # the same models as one run
    assert not (root / ".resume_train").exists()


def _block(path):
    """A file where the folder should be: makedirs and every write into it fail (same on Windows and Linux)."""
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text("not a folder", encoding="utf-8")


def _fallbacks(primary):
    return rt.resume_places(str(primary))[1:]


@pytest.mark.parametrize("targets", ["생산효율(%)", "생산효율(%),수분함량(%)"])
def test_blocked_resume_folder_continues_in_fallback(tmp_path, monkeypatch, targets):
    """3차 검토(2.9.63): 2.9.62는 중간 저장이 실패하면 나누지 않고 끝까지 진행해 큰 데이터에서 120초를 넘겨 JSON 없이 끊겼다.
    이제 대체 위치(데이터 폴더 resume\\<해시>)에 저장하고(recovered) CONTINUE로 나눠 끝까지, 끝나면 대체 위치도 지운다."""
    data = make_sample(str(tmp_path / "f.csv"))
    root = tmp_path / "f_ml_data"
    _block(root / ".resume_train")
    fb = _fallbacks(root / ".resume_train")[0]
    # 느린 PC는 실행마다 한 단계만 진행: 타겟 2개면 약 55단계 (회사 PC 2026-11-06: 40회 상한에 걸림)
    runs = _split_runs(monkeypatch, "--file", data, "--features", SPLIT_X, "--targets", targets, limit=120)
    assert len(runs) >= 2, [r.get("progress") for r in runs]
    first, last = runs[0], runs[-1]
    assert first["agent_directive"].startswith("CONTINUE:"), first
    assert any("대체 위치 데이터 폴더" in r for r in first["recovered"]), first.get("recovered")
    assert os.path.basename(fb) in first["recovered"][0]
    assert last["version"] == "v1" and last["agent_directive"].startswith("VERIFIED"), last
    assert not os.path.exists(fb)                                  # cleaned after the commit
    assert (root / ".resume_train").is_file()                      # the user's file is not touched


def test_resume_saved_nowhere_stops_with_stop_retry(tmp_path, monkeypatch):
    """원래 위치·데이터 폴더·임시 폴더가 모두 막히면 제한을 넘겨 계속하지 않고, 멈출 지점에서 STOP_RETRY 1줄 JSON."""
    data = make_sample(str(tmp_path / "n.csv"))
    root = tmp_path / "n_ml_data"
    _block(root / ".resume_train")
    temp = tmp_path / "temp"
    monkeypatch.setenv("TEMP", str(temp))
    monkeypatch.setenv("ML_KIT_DATA_DIR", str(tmp_path / "data"))
    _block(tmp_path / "data" / rt.RESUME_FALLBACK_DIR)
    _block(temp / rt.TEMP_RESUME_DIR)
    monkeypatch.setenv("ML_KIT_TIME_LIMIT", "30.5")
    t0 = time.monotonic()
    proc, out = run_raw("ml_train.py", "--file", data, "--features", SPLIT_X, "--targets", "생산효율(%),수분함량(%)")
    assert proc.returncode == 1 and out["status"] == "error", out
    assert out["agent_directive"].startswith("STOP_RETRY") and "저장하지 못해" in out["message"], out
    assert "쓰기 권한" in out["message"] and "[WARN] 중간 저장 실패" in proc.stderr
    # 4차 검토: the directive points at the save problem, not the training hint (column names, --eval)
    assert "--eval" not in out["agent_directive"] and "쓰기 권한" in out["agent_directive"], out["agent_directive"]
    assert time.monotonic() - t0 < 60
    assert not (root / "history.jsonl").exists()


def test_resume_save_checks_the_file(tmp_path, monkeypatch):
    monkeypatch.setattr(rt, "SAVE_FAILED", [])
    monkeypatch.setattr(rt, "_RESUME_PLACE", {})
    monkeypatch.setattr(rt, "_RECOVERED", [])
    monkeypatch.setenv("ML_KIT_DATA_DIR", str(tmp_path / "data"))
    monkeypatch.setenv("TEMP", str(tmp_path / "temp"))
    assert mt._resume_save(str(tmp_path / "r" / "a.joblib"), (1, 2, 3, 4)) is True and not rt.SAVE_FAILED
    assert mt._resume_load(str(tmp_path / "r" / "a.joblib")) == (1, 2, 3, 4) and not rt._RECOVERED
    _block(tmp_path / "x")                                            # original blocked -> data folder
    assert mt._resume_save(str(tmp_path / "x" / "a.joblib"), (1, 2, 3, 4)) is True and rt._RECOVERED
    data_fb, temp_fb = _fallbacks(tmp_path / "x")
    assert os.path.isfile(os.path.join(data_fb, "a.joblib"))
    assert mt._resume_load(str(tmp_path / "x" / "a.joblib")) == (1, 2, 3, 4)   # found in the fallback
    assert len(os.path.basename(data_fb)) == 12
    _block(tmp_path / "y")                                            # original and data folder blocked -> %TEMP%
    monkeypatch.setattr(rt, "_RESUME_PLACE", {})
    _block(tmp_path / "data2" / rt.RESUME_FALLBACK_DIR)
    monkeypatch.setenv("ML_KIT_DATA_DIR", str(tmp_path / "data2"))
    assert mt._resume_save(str(tmp_path / "y" / "a.joblib"), (1, 2, 3, 4)) is True
    assert os.path.isfile(os.path.join(_fallbacks(tmp_path / "y")[1], "a.joblib"))
    _block(tmp_path / "temp2" / rt.TEMP_RESUME_DIR)                   # every place blocked
    monkeypatch.setenv("TEMP", str(tmp_path / "temp2"))
    assert mt._resume_save(str(tmp_path / "z" / "q" / "a.joblib"), (1,)) is True    # a plain new folder works
    _block(tmp_path / "w")
    assert mt._resume_save(str(tmp_path / "w" / "a.joblib"), (1, 2, 3, 4)) is False
    assert rt.SAVE_FAILED and rt.save_failed_warning().startswith("중간 저장 실패")
    monkeypatch.setenv("ML_KIT_TIME_LIMIT", "30.001")
    with pytest.raises(rt.InputError) as e:                           # stops where it would stop, as STOP_RETRY
        rt.check_time(1000.0, "멈춤")
    assert "저장하지 못해" in str(e.value)
    monkeypatch.setenv("ML_KIT_TIME_LIMIT", "0")
    rt.check_time(1000.0, "제한 없음")                                 # no limit: goes on to the end


def test_resume_find_newest_and_old_fallbacks_swept(tmp_path, monkeypatch):
    """4차 검토: (a) 같은 이름이 여러 위치에 있으면 가장 최근 것을 읽고, 대체 위치에 저장하면 앞 위치의 옛 파일을 지운다
    (run.json 같은 고정 이름). (b) 실행 시작 때 대체 폴더 전체에서 보관 기한(3일)이 지난 항목을 지운다."""
    monkeypatch.setattr(rt, "SAVE_FAILED", [])
    monkeypatch.setattr(rt, "_RESUME_PLACE", {})
    monkeypatch.setattr(rt, "_RECOVERED", [])
    monkeypatch.setenv("ML_KIT_DATA_DIR", str(tmp_path / "data"))
    monkeypatch.setenv("TEMP", str(tmp_path / "temp"))
    primary = str(tmp_path / "rt")
    places = rt.resume_places(primary)
    for i, d in enumerate(places[:2]):
        os.makedirs(d)
        with open(os.path.join(d, "run.json"), "w", encoding="utf-8") as f:
            f.write(str(i))
    old = time.time() - 100
    os.utime(os.path.join(places[0], "run.json"), (old, old))
    assert rt.resume_find(primary, "run.json") == os.path.join(places[1], "run.json")
    os.utime(os.path.join(places[1], "run.json"), (old - 100, old - 100))
    assert rt.resume_find(primary, "run.json") == os.path.join(places[0], "run.json")
    rt._RESUME_PLACE[os.path.normcase(places[0])] = 1         # the original folder failed earlier in this command
    path = rt.resume_save(primary, "run.json", lambda tmp: open(tmp, "w", encoding="utf-8").write("new"), "기록")
    assert path == os.path.join(places[1], "run.json")
    assert not os.path.exists(os.path.join(places[0], "run.json"))     # the older copy is not read again
    assert rt.resume_find(primary, "run.json") == path
    # (b) sweep: an old saved step of another data file in both fallback roots goes, a recent one stays
    stale = [os.path.join(tmp_path / "data", rt.RESUME_FALLBACK_DIR, "aaaaaaaaaaaa", "0123456789abcdef.m0.joblib"),
             os.path.join(tmp_path / "temp", rt.TEMP_RESUME_DIR, "bbbbbbbbbbbb", "0123456789abcdef.joblib")]
    for f in stale:
        os.makedirs(os.path.dirname(f))
        open(f, "wb").write(b"1")
        t = time.time() - rt.RESUME_MAX_AGE - 60
        os.utime(f, (t, t))
    rt.resume_sweep_old()
    assert not any(os.path.exists(os.path.dirname(f)) for f in stale)
    assert not os.path.exists(tmp_path / "temp" / rt.TEMP_RESUME_DIR)      # left empty -> removed
    assert os.path.isfile(path)                                              # recent: kept


def test_resume_sweep_never_touches_user_files(tmp_path, monkeypatch):
    """5차 검토: ML_KIT_DATA_DIR를 사용자 폴더로 두면 resume_sweep_old가 그 아래 resume 폴더의 모든 파일·폴더를 3일 지나면 지웠다
    (docs/resume/2025/cv.docx). 이제 대체 폴더 이름은 ml-kit-resume, 지우는 것은 12자리 16진수 하위 폴더 안의 도구 파일 이름뿐."""
    docs = tmp_path / "docs"
    monkeypatch.setenv("ML_KIT_DATA_DIR", str(docs))
    monkeypatch.setenv("TEMP", str(tmp_path / "temp"))
    old = time.time() - rt.RESUME_MAX_AGE - 3600
    user = [docs / "resume" / "2025" / "cv.docx", docs / "resume" / "0123456789ab" / "0123456789abcdef.joblib",
            docs / rt.RESUME_FALLBACK_DIR / "2025" / "cv.docx", docs / rt.RESUME_FALLBACK_DIR / "notes.txt",
            docs / rt.RESUME_FALLBACK_DIR / "0123456789ab" / "cv.docx",
            tmp_path / "temp" / rt.TEMP_RESUME_DIR / "0123456789ab" / "model.joblib"]
    tool = [docs / rt.RESUME_FALLBACK_DIR / "0123456789ab" / n
            for n in ("0123456789abcdef.joblib", "0123456789abcdef.m3.joblib", "0123456789abcdef.fit.joblib",
                      "0123456789abcdef.imp.joblib.tmp")]
    for f in user + tool:
        os.makedirs(f.parent, exist_ok=True)
        f.write_bytes(b"1")
        os.utime(f, (old, old))
    rt.resume_sweep_old()
    assert all(f.is_file() for f in user), [str(f) for f in user if not f.is_file()]
    assert not any(f.exists() for f in tool)
    primary = str(tmp_path / "rt")                            # resume_clear in a fallback place: only the tool's names
    fb = rt.resume_places(primary)[1]
    os.makedirs(fb)
    for n in ("cv.docx", "0123456789abcdef.joblib"):
        open(os.path.join(fb, n), "wb").write(b"1")
    rt.resume_clear(primary, lambda name, p: True, fallback_only=True)
    assert os.listdir(fb) == ["cv.docx"]


def test_full_fit_after_reused_candidates_is_checked(tmp_path, monkeypatch):
    """후보를 모두 재사용한 타겟의 1위 전체 적합·순열 중요도 앞에도 시간 확인 (이 실행에서 다른 단계를 했을 때만)."""
    r = np.random.default_rng(3)
    X = pd.DataFrame({"a": r.uniform(0, 1, 60), "b": r.uniform(0, 1, 60)})
    y = pd.Series(3 * X.a + X.b + r.normal(0, 0.05, 60))
    schema = {"numeric": ["a", "b"], "categorical": {}, "datetime": {}, "input_columns": ["a", "b"]}
    resume = str(tmp_path / "rt" / "k")
    monkeypatch.setattr(rt, "SAVE_FAILED", [])
    monkeypatch.setattr(mt, "RUN_MODEL_SECONDS", [])
    monkeypatch.setattr(mt, "RUN_FOLD_SECONDS", [])
    monkeypatch.setattr(mt, "RUN_FIT_SECONDS", [])
    monkeypatch.setenv("ML_KIT_TIME_LIMIT", "0")
    mt.train_and_benchmark_models(X=X, y_series=y, target_name="y", schema=schema, resume=resume)   # saves 5 candidates
    assert sorted(os.listdir(tmp_path / "rt")) == [f"k.m{i}.joblib" for i in range(5)]   # fit saved only at a stop
    monkeypatch.setattr(mt, "RUN_MODEL_SECONDS", [])
    monkeypatch.setattr(mt, "RUN_FOLD_SECONDS", [])
    monkeypatch.setenv("ML_KIT_TIME_LIMIT", "30.5")
    with pytest.raises(rt.Unfinished) as e:                   # first step (the fit) goes on, then stops (4차 검토)
        mt.train_and_benchmark_models(X=X, y_series=y, target_name="y", schema=schema, resume=resume)
    assert "순열 중요도 0/2열" in e.value.progress, e.value.progress
    monkeypatch.setattr(mt, "RUN_MODEL_SECONDS", [])
    monkeypatch.setattr(mt, "RUN_FOLD_SECONDS", [])
    monkeypatch.setattr(mt, "RUN_FIT_SECONDS", [5.0])        # another target was trained in this command
    with pytest.raises(rt.Unfinished) as e:                   # the saved fit is reused: no stop before it
        mt.train_and_benchmark_models(X=X, y_series=y, target_name="y", schema=schema, resume=resume)
    assert "순열 중요도 0/2열" in e.value.progress, e.value.progress
    os.remove(rt.resume_find(str(tmp_path / "rt"), "k.fit.joblib"))
    with pytest.raises(rt.Unfinished) as e:
        mt.train_and_benchmark_models(X=X, y_series=y, target_name="y", schema=schema, resume=resume)
    assert "1위 모델 전체 적합 전에 멈춤" in e.value.progress and "5/5개" in e.value.progress


def _wide_frame(n=60, p=24, seed=11):
    r = np.random.default_rng(seed)
    X = pd.DataFrame({f"x{i}": r.uniform(0, 1, n) for i in range(p)})
    y = pd.Series(5 * X.x0 + 3 * X.x1 - 2 * X.x2 + X.x3 + r.normal(0, 0.05, n))
    cols = list(X.columns)
    return X, y, {"numeric": cols, "categorical": {}, "datetime": {}, "input_columns": cols}


@pytest.mark.parametrize("task", ["regression", "classification"])
def test_wide_target_last_step_is_split_and_matches_one_run(tmp_path, monkeypatch, task):
    """4차 검토: 열이 많으면 마지막 단계(1위 모델 전체 적합 + 순열 중요도)가 120초를 넘어 JSON 없이 끊겼고, 다시 실행해도
    같은 단계를 반복했다. 이제 전체 적합을 저장하고 순열 중요도를 열 묶음으로 나눠 저장하며 CONTINUE: 실행마다 최소 한 단계
    (중요도는 1열 이상)를 하므로 유한 횟수에 끝나고, 결과(리더보드·모델·중요도)는 한 번에 학습한 것과 같다."""
    X, y, schema = _wide_frame()
    if task == "classification":
        y = pd.Series(np.where(y > y.median(), "A", "B"))
        fn = mt.train_and_benchmark_classifiers
    else:
        fn = mt.train_and_benchmark_models
    monkeypatch.setattr(rt, "SAVE_FAILED", [])
    monkeypatch.setenv("ML_KIT_DATA_DIR", str(tmp_path / "data"))
    monkeypatch.setenv("ML_KIT_TIME_LIMIT", "0")
    want = fn(X=X, y_series=y, target_name="y", schema=schema)
    monkeypatch.setenv("ML_KIT_TIME_LIMIT", "30.5")          # every check after a step stops (in-process clock)
    resume = str(tmp_path / "rt" / "k")
    progress = []
    for _ in range(60):
        monkeypatch.setattr(mt, "RUN_MODEL_SECONDS", [])      # a new run of the same command
        monkeypatch.setattr(mt, "RUN_FOLD_SECONDS", [])
        monkeypatch.setattr(mt, "RUN_FIT_SECONDS", [])
        try:
            got = fn(X=X, y_series=y, target_name="y", schema=schema, resume=resume)
            break
        except rt.Unfinished as e:
            progress.append(e.progress)
    else:
        pytest.fail(f"끝나지 않음: {progress[-3:]}")
    imp_stops = [t for t in progress if "순열 중요도" in t]
    assert len(progress) <= 5 * 5 + 1 + 24, progress                    # at least one fold or column per run
    assert imp_stops and "0/24열" in imp_stops[0], progress
    assert len({t for t in imp_stops}) == len(imp_stops), imp_stops                 # every run moves on
    assert got[0]["leaderboard"] == want[0]["leaderboard"]
    assert got[2]["feature_importances"] == want[2]["feature_importances"]          # even the importances
    assert np.array_equal(got[1].predict(X), want[1].predict(X))


def test_wide_data_command_split_runs_end_like_one_run(tmp_path, monkeypatch):
    """4차 검토 재현(명령 단위): 열이 많은 데이터를 작은 ML_KIT_TIME_LIMIT로 학습하면 실행마다 제한 안에 1줄 JSON(CONTINUE)을 내고
    유한 횟수에 끝나며, 리더보드·중요도는 한 번에 학습한 것과 같다."""
    X, y, _ = _wide_frame(n=60, p=8)
    df = X.round(5).assign(y=y.round(5))
    cols = ",".join(X.columns)
    ref = write_df(tmp_path / "ref.csv", df)
    want = run_tool("ml_train.py", "--file", ref, "--features", cols, "--targets", "y")
    data = write_df(tmp_path / "w.csv", df)
    t0 = time.monotonic()
    runs = _split_runs(monkeypatch, "--file", data, "--features", cols, "--targets", "y", limit=5 * 5 + 1 + 8 + 1,
                       seconds="30.01")                     # 0.01 s after the reserve: every step stops
    progress = [r.get("progress") or "" for r in runs]
    assert any("순열 중요도" in t for t in progress), progress
    assert time.monotonic() - t0 < 60 * len(runs)
    last = runs[-1]
    assert last["agent_directive"].startswith("VERIFIED") and last["top_features"] == want["top_features"], last
    board = [json.load(open(os.path.join(r["version_dir"], "leaderboard.json"), encoding="utf-8")) for r in (want, last)]
    for b in board:
        b.pop("timestamp", None)
    assert board[0] == board[1]
    assert not (tmp_path / "w_ml_data" / ".resume_train").exists()


# ---- 후보 모델 안에서도 교차검증 fold 단위로 나눠 진행 (R10, 2.9.66) -------------------------------------------------
# 후보 하나의 교차검증(fold 5개)이 120초를 넘으면 그 후보는 매번 처음부터 다시 해 끝나지 않았다(무응답 반복).
# 이제 fold마다 <키>.m<번호>.f<fold>.joblib에 저장하고, 다음 fold 앞에서 시계를 본다.

FOLD_CASES = [("regression", False), ("regression", True), ("classification", False), ("classification", True)]
_FOLD_FILE = re.compile(r"k\.m\d+\.f\d+\.joblib")


def _new_run(monkeypatch, fold=(), model=(), fit=()):
    """A new run of the same command in this process: the steps done in this command start empty."""
    monkeypatch.setattr(mt, "RUN_FOLD_SECONDS", list(fold))
    monkeypatch.setattr(mt, "RUN_MODEL_SECONDS", list(model))
    monkeypatch.setattr(mt, "RUN_FIT_SECONDS", list(fit))
    monkeypatch.setattr(mt, "RUN_PREP_SECONDS", [])
    monkeypatch.setattr(rt, "CURRENT_STEP", {})


@pytest.fixture(autouse=True)
def _fresh_command_steps(monkeypatch):
    """Every test starts like a new command: steps of in-process trainings of other tests do not count."""
    _new_run(monkeypatch)


def _fold_data(task, dup):
    r = np.random.default_rng(5)
    n = 70
    X = pd.DataFrame({"a": r.uniform(0, 1, n).round(2), "b": r.uniform(0, 1, n).round(2)})
    y = pd.Series(4 * X.a - 2 * X.b + r.normal(0, 0.05, n))
    if task == "classification":
        y = pd.Series(np.where(y > y.median(), "A", "B"))
    if dup:                                     # identical rows (inputs + target): one fold per copy group (2.9.13)
        X = pd.concat([X, X.iloc[:12]], ignore_index=True)
        y = pd.concat([y, y.iloc[:12]], ignore_index=True)
    fn = mt.train_and_benchmark_classifiers if task == "classification" else mt.train_and_benchmark_models
    return X, y, {"numeric": ["a", "b"], "categorical": {}, "datetime": {}, "input_columns": ["a", "b"]}, fn


def _assert_same_as_one_run(got, want, X):
    assert got[0]["leaderboard"] == want[0]["leaderboard"]
    assert got[3].equals(want[3])                                            # OOF predictions, to the last bit
    assert got[2]["feature_importances"] == want[2]["feature_importances"]
    assert got[2]["top_features"] == want[2]["top_features"]
    assert np.array_equal(got[1].predict(X), want[1].predict(X))            # the final model


@pytest.mark.parametrize("task,dup", FOLD_CASES)
def test_slow_candidate_is_split_by_fold_and_matches_one_run(tmp_path, monkeypatch, task, dup):
    """2.9.66: 실행마다 fold 1개만 끝낼 시간이어도 fold를 저장하며 CONTINUE로 나아가 유한 횟수에 끝나고, 결과(리더보드·OOF·
    최종 모델 예측·중요도)는 한 번에 학습한 것과 같다. 회귀·분류(층화), 중복 행 묶음 fold 모두."""
    X, y, schema, fn = _fold_data(task, dup)
    monkeypatch.setattr(rt, "SAVE_FAILED", [])
    monkeypatch.setenv("ML_KIT_DATA_DIR", str(tmp_path / "data"))
    monkeypatch.setenv("ML_KIT_TIME_LIMIT", "0")
    want = fn(X=X, y_series=y, target_name="y", schema=schema)
    assert bool(want[2].get("duplicate_rows")) == dup
    monkeypatch.setenv("ML_KIT_TIME_LIMIT", "30.5")          # every check after a step stops: one fold per run
    resume = str(tmp_path / "rt" / "k")
    progress, fold_files = [], set()
    for _ in range(60):
        _new_run(monkeypatch)
        try:
            got = fn(X=X, y_series=y, target_name="y", schema=schema, resume=resume)
            break
        except rt.Unfinished as e:
            progress.append(e.progress)
            fold_files |= {n for n in os.listdir(tmp_path / "rt") if _FOLD_FILE.fullmatch(n)}
    else:
        pytest.fail(f"끝나지 않음: {progress[-3:]}")
    fold_stops = [t for t in progress if "교차검증 fold" in t]
    assert fold_stops and "'HistGradientBoosting' 교차검증 fold 1/" in fold_stops[0], progress
    assert len(set(progress)) == len(progress), progress                    # every run moves on
    assert len(progress) <= 5 * 5 + 1 + 2, progress                          # folds + final fit + 2 importance columns
    assert "k.m0.f1.joblib" in fold_files and "k.m4.f1.joblib" in fold_files, fold_files
    assert not any(_FOLD_FILE.fullmatch(n) for n in os.listdir(tmp_path / "rt"))   # merged into the candidates
    _assert_same_as_one_run(got, want, X)


@pytest.mark.parametrize("task", ["regression", "classification"])
def test_killed_run_keeps_saved_folds_and_damaged_fold_is_trained_again(tmp_path, monkeypatch, capsys, task):
    """2.9.66: fold를 저장한 뒤 바깥에서 끊긴 실행(에이전트의 120초 강제 종료)도 끝낸 fold가 남아 다음 실행이 이어서 하고,
    손상된 fold 파일은 버리고 그 fold만 다시 학습한다 (JIT: 다시 생성). 결과는 한 번에 학습한 것과 같다."""
    X, y, schema, fn = _fold_data(task, False)
    monkeypatch.setattr(rt, "SAVE_FAILED", [])
    monkeypatch.setenv("ML_KIT_DATA_DIR", str(tmp_path / "data"))
    monkeypatch.setenv("ML_KIT_TIME_LIMIT", "0")
    want = fn(X=X, y_series=y, target_name="y", schema=schema)
    real, calls = mt.fit_safe_model, []

    def killed_at_8th_fit(est, y_train):                       # candidate 2, fold 3
        calls.append(1)
        if len(calls) == 8:
            raise KeyboardInterrupt("killed")
        return real(est, y_train)

    monkeypatch.setattr(mt, "fit_safe_model", killed_at_8th_fit)
    resume = str(tmp_path / "rt" / "k")
    _new_run(monkeypatch)
    with pytest.raises(KeyboardInterrupt):
        fn(X=X, y_series=y, target_name="y", schema=schema, resume=resume)
    assert sorted(os.listdir(tmp_path / "rt")) == ["k.m0.joblib", "k.m1.f1.joblib", "k.m1.f2.joblib"]
    (tmp_path / "rt" / "k.m1.f2.joblib").write_bytes(b"broken")
    calls.clear()
    monkeypatch.setattr(mt, "fit_safe_model", lambda est, y_train: calls.append(1) or real(est, y_train))
    capsys.readouterr()
    _new_run(monkeypatch)
    got = fn(X=X, y_series=y, target_name="y", schema=schema, resume=resume)
    err = capsys.readouterr().err
    assert "[WARN] 저장된 교차검증 fold 결과 k.m1.f2.joblib를 읽지 못해" in err, err
    assert "교차검증 fold 1/5개는 이전 실행에서 저장한 결과를 씀" in err, err
    assert len(calls) == 4 + 3 * 5 + 1                                     # folds 2-5, 3 candidates, the final fit
    assert sorted(os.listdir(tmp_path / "rt")) == [f"k.m{i}.joblib" for i in range(5)]
    _assert_same_as_one_run(got, want, X)


def test_fold_clock_check_estimates(tmp_path, monkeypatch):
    """2.9.66: fold 앞 시간 확인은 이 후보의 fold 실측(저장된 것 포함) 최대값 x 1.2, 첫 fold는 이 실행의 다른 후보 fold 시간으로
    추정하고, 추정할 값이 없으면 진행한다. 이 실행에서 아무 단계도 안 했으면 멈추지 않는다 (R10: 실행마다 최소 1단계)."""
    import joblib
    X, y, schema, fn = _fold_data("regression", False)
    monkeypatch.setattr(rt, "SAVE_FAILED", [])
    monkeypatch.setenv("ML_KIT_DATA_DIR", str(tmp_path / "data"))
    monkeypatch.setenv("ML_KIT_TIME_LIMIT", "100")
    monkeypatch.setattr(rt, "RUN_START", time.monotonic())
    resume = str(tmp_path / "rt" / "k")
    _new_run(monkeypatch, fold=[1000.0])                       # a slow fold of another candidate in this command
    with pytest.raises(rt.Unfinished) as e:
        fn(X=X, y_series=y, target_name="y", schema=schema, resume=resume)
    assert "후보 모델 0/5개 끝남, 'HistGradientBoosting' 교차검증 fold 0/5개 끝남" in e.value.progress, e.value.progress
    assert "(fold 1 전에 멈춤, 약 27단계 남음)" in e.value.progress
    os.makedirs(tmp_path / "rt", exist_ok=True)               # this candidate's own saved fold was slow: it decides
    joblib.dump({"model": "HistGradientBoosting", "folds": 5, "seconds": 1000.0,
                 "res": {"train": 0.0, "pred": np.zeros(14), "params": {}}}, tmp_path / "rt" / "k.m0.f1.joblib")
    _new_run(monkeypatch, fold=[0.01])
    with pytest.raises(rt.Unfinished) as e:
        fn(X=X, y_series=y, target_name="y", schema=schema, resume=resume)
    assert "fold 1/5개 끝남, 저장됨 (fold 2 전에 멈춤, 약 26단계 남음)" in e.value.progress, e.value.progress
    os.remove(tmp_path / "rt" / "k.m0.f1.joblib")
    _new_run(monkeypatch, fit=[5.0])                           # a step, but no fold time known: goes on to the end
    res = fn(X=X, y_series=y, target_name="y", schema=schema, resume=resume)
    assert len(res[0]["leaderboard"]) == 5
    shutil.rmtree(tmp_path / "rt")
    real = mt.fit_safe_model                                   # no step in this run: the first fold always goes
    monkeypatch.setattr(mt, "fit_safe_model", lambda est, y_train: time.sleep(0.2) or real(est, y_train))
    monkeypatch.setenv("ML_KIT_TIME_LIMIT", "30.1")
    monkeypatch.setattr(rt, "RUN_START", time.monotonic())
    _new_run(monkeypatch)
    with pytest.raises(rt.Unfinished) as e:
        fn(X=X, y_series=y, target_name="y", schema=schema, resume=resume)
    assert "'HistGradientBoosting' 교차검증 fold 1/5개 끝남, 저장됨 (fold 2 전에 멈춤, 약 26단계 남음)" in e.value.progress, e.value.progress


def test_last_step_estimate_uses_saved_candidate_times(tmp_path, monkeypatch):
    """4차 검토 (d): 1위 모델 전체 적합 전 시간 확인은 저장된 후보 실측 시간(재사용 포함)으로 추정한다. 5차 검토: 가장 느린 후보가 아니라
    1위 모델 자신의 교차검증 시간으로 추정한다(느린 RandomForest가 지면 빠른 전체 적합 앞에서 멈춰 실행 횟수가 늘었음).
    (b) 순열 중요도는 예측 1회 시간으로 추정해 예산 안이면 한 번에 끝낸다."""
    import joblib
    X, y, schema = _wide_frame(p=4)
    resume = str(tmp_path / "rt" / "k")
    monkeypatch.setattr(rt, "SAVE_FAILED", [])
    monkeypatch.setattr(mt, "RUN_MODEL_SECONDS", [])
    monkeypatch.setattr(mt, "RUN_FOLD_SECONDS", [])
    monkeypatch.setattr(mt, "RUN_FIT_SECONDS", [])
    monkeypatch.setenv("ML_KIT_TIME_LIMIT", "0")
    first = mt.train_and_benchmark_models(X=X, y_series=y, target_name="y", schema=schema, resume=resume)
    champion = first[0]["champion"]["champion_model"]
    cands = {joblib.load(tmp_path / "rt" / f"k.m{i}.joblib")[0]: tmp_path / "rt" / f"k.m{i}.joblib" for i in range(5)}
    slow = cands[champion]
    other = next(p for n, p in cands.items() if n != champion)
    saved = joblib.load(slow)
    assert isinstance(saved[-1], float)                       # the measured time is saved with the candidate
    o_saved = joblib.load(other)
    joblib.dump(o_saved[:-1] + (1000.0,), other)              # a slow candidate that lost: not the final fit's time
    monkeypatch.setenv("ML_KIT_TIME_LIMIT", "100")
    monkeypatch.setattr(rt, "RUN_START", time.monotonic())
    monkeypatch.setattr(mt, "RUN_MODEL_SECONDS", [0.01])      # a fast step done in this run
    monkeypatch.setattr(mt, "RUN_FOLD_SECONDS", [])
    res = mt.train_and_benchmark_models(X=X, y_series=y, target_name="y", schema=schema, resume=resume)
    assert len(res[2]["feature_importances"]) == 4
    joblib.dump(saved[:-1] + (1000.0,), slow)                 # the champion was slow in an earlier run
    monkeypatch.setattr(rt, "RUN_START", time.monotonic())
    with pytest.raises(rt.Unfinished) as e:
        mt.train_and_benchmark_models(X=X, y_series=y, target_name="y", schema=schema, resume=resume)
    assert "1위 모델 전체 적합 전에 멈춤" in e.value.progress
    joblib.dump(saved[:-1] + (0.01,), slow)
    res = mt.train_and_benchmark_models(X=X, y_series=y, target_name="y", schema=schema, resume=resume)
    assert len(res[2]["feature_importances"]) == 4             # fit and the whole importance in this run


class _SlowScore:
    """A fitted pipeline whose score takes `delay` seconds (a slow prediction: 600 columns, KNN) and that can be
    'killed' (KeyboardInterrupt from outside, like the 120 s limit of the agent) at the kill-th call."""

    def __init__(self, pipe, delay=0.0, kill=None):
        self.pipe, self.delay, self.kill, self.calls = pipe, delay, kill, 0

    def score(self, X, y):
        self.calls += 1
        if self.kill and self.calls == self.kill:
            raise KeyboardInterrupt("killed")
        time.sleep(self.delay)
        return self.pipe.score(X, y)


def _imp_setup(tmp_path, monkeypatch, p=40):
    X, y, schema = _wide_frame(p=p)
    from sklearn.linear_model import Ridge
    pipe = mt.build_pipeline(Ridge(), schema, "regressor").fit(X, y)
    monkeypatch.setattr(rt, "SAVE_FAILED", [])
    monkeypatch.setattr(rt, "_RESUME_PLACE", {})
    monkeypatch.setattr(mt, "RUN_MODEL_SECONDS", [])
    monkeypatch.setattr(mt, "RUN_FOLD_SECONDS", [])
    monkeypatch.setattr(mt, "RUN_FIT_SECONDS", [])
    monkeypatch.setenv("ML_KIT_DATA_DIR", str(tmp_path / "data"))
    monkeypatch.setenv("TEMP", str(tmp_path / "temp"))
    monkeypatch.setenv("ML_KIT_TIME_LIMIT", "0")
    want = mt._permutation_means(pipe, X, y.values, 42, None, None, "y", "M", lambda: None)
    return X, y.values, pipe, want, str(tmp_path / "rt" / "k")


def test_slow_wide_importance_stops_by_the_clock_from_the_second_column(tmp_path, monkeypatch):
    """5차 검토 (fh.py, 600열 HistGB/KNN): 다른 단계를 안 한 실행은 처음 열의 1/5을 시계 없이 계산해 120초를 넘겨 JSON 없이
    끊기기를 반복했다. 이제 시계 없이 하는 열은 1개뿐이고 그 뒤는 실측한 열 시간으로 끊는다: 실행마다 예산 안에서 멈추며
    앞으로 나아가고, 나눠 계산한 값은 한 번에 계산한 것과 같다."""
    X, y, pipe, want, resume = _imp_setup(tmp_path, monkeypatch)
    slow = _SlowScore(pipe, delay=0.05)                       # one column ~0.25 s; 40 columns ~10 s
    monkeypatch.setenv("ML_KIT_TIME_LIMIT", str(rt.TIME_RESERVE + 1.0))    # 1 s of work per run
    done_before, runs = 0, 0
    while True:
        runs += 1
        assert runs <= 40, "끝나지 않음"
        monkeypatch.setattr(rt, "RUN_START", time.monotonic())
        try:
            got = mt._permutation_means(slow, X, y, 42, None, resume, "y", "M", lambda: None)
            break
        except rt.Unfinished as e:
            assert time.monotonic() - rt.RUN_START < 1.0 + 0.5, e.progress          # within the budget (+ a column)
            k = int(e.progress.split("순열 중요도 ")[1].split("/")[0])
            assert k > done_before, e.progress                                          # every run moves on
            done_before = k
    assert 5 <= runs <= 25, runs
    assert np.allclose(got, want)


def test_killed_importance_keeps_columns_saved_on_the_way(tmp_path, monkeypatch):
    """5차 검토: 끝낸 열을 멈출 때만 저장해 명령이 밖에서 강제 종료되면(120초) 진행이 사라졌다. 이제 IMPORTANCE_SAVE_SECONDS마다
    저장(먼저 before_stop: 1위 모델 전체 적합)하므로 강제 종료돼도 다음 실행은 저장된 열부터 이어 가고 값은 한 번에 한 것과 같다."""
    import joblib
    X, y, pipe, want, resume = _imp_setup(tmp_path, monkeypatch, p=12)
    monkeypatch.setattr(mt, "IMPORTANCE_SAVE_SECONDS", 0.0)   # save after every column
    stops = []
    killed = _SlowScore(pipe, kill=1 + 5 * 4 + 2)             # base score, 4 columns, killed in the 5th
    with pytest.raises(KeyboardInterrupt):
        mt._permutation_means(killed, X, y, 42, None, resume, "y", "M", lambda: stops.append(1))
    saved = joblib.load(rt.resume_find(os.path.dirname(resume), "k.imp.joblib"))
    assert sorted(saved["done"]) == [0, 1, 2, 3] and stops, saved
    again = _SlowScore(pipe)
    got = mt._permutation_means(again, X, y, 42, None, resume, "y", "M", lambda: None)
    assert again.calls == 1 + 5 * 8                           # only the 8 columns left
    assert np.allclose(got, want)


def test_damaged_saved_target_is_trained_again(tmp_path, monkeypatch):
    data = make_sample(str(tmp_path / "d.csv"))
    root = tmp_path / "d_ml_data"
    monkeypatch.setenv("ML_KIT_TIME_LIMIT", "30.5")
    proc, out = run_raw("ml_train.py", "--file", data, "--features", SPLIT_X, "--targets", "생산효율(%),수분함량(%)")
    assert out.get("continue"), out
    saved = os.listdir(root / ".resume_train")
    assert saved and all(".m" in n for n in saved), saved          # the first target's candidates (2.9.60)
    for n in saved:
        (root / ".resume_train" / n).write_bytes(b"broken")
    monkeypatch.setenv("ML_KIT_TIME_LIMIT", "0")
    proc, out = run_raw("ml_train.py", "--file", data, "--features", SPLIT_X, "--targets", "생산효율(%),수분함량(%)")
    assert out["status"] == "ok" and out["version"] == "v1" and "resumed_targets" not in out, out
    assert "다시 학습" in proc.stderr


# ---- 학습 전 준비 저장, 시작 표시, 시간 제한 감시 (R10, 2.9.67) ----------------------------------------------------------
# 6차 검토(10000행 x 1000열): 실행마다 학습 전 준비(읽기·검증·상관행렬·타겟 준비, 50~94초)를 다시 하고 그 실행의 첫 단계
# (fold 1개, 시계 없음)를 더하면 120초를 넘어 JSON 없이 끊기고 진행도 없어 같은 명령을 끝없이 반복했다.

def _train_args(data, x, y, **kw):
    base = dict(file=data, features=x, targets=y, eval=None, dtypes=None, exclude_where=None, task=None,
                positive_class=None, header_row=None, unit_row=None, data_row=None, seed=42, version=None,
                overwrite=False, description=None, tag=None)
    base.update(kw)
    return argparse.Namespace(**base)


def test_preparation_is_saved_and_reused_by_the_continued_run(tmp_path, monkeypatch):
    """준비가 오래 걸리면(PREP_SAVE_SECONDS 이상) 통째로 저장하고, 제한에 닿았으면 첫 학습 단계 전에 CONTINUE로 멈춘다.
    같은 명령을 다시 실행하면 준비를 다시 하지 않고 저장한 것을 쓰며, 결과는 한 번에 학습한 것과 같다."""
    ref = make_sample(str(tmp_path / "ref.csv"))
    want = run_tool("ml_train.py", "--file", ref, "--features", SPLIT_X, "--targets", "생산효율(%)")
    data = make_sample(str(tmp_path / "p.csv"))
    root = tmp_path / "p_ml_data"
    monkeypatch.setattr(rt, "SAVE_FAILED", [])
    monkeypatch.setattr(rt, "_RESUME_PLACE", {})
    monkeypatch.setattr(mt, "PREP_SAVE_SECONDS", 0.0)
    monkeypatch.setenv("ML_KIT_TIME_LIMIT", "100")
    monkeypatch.setattr(rt, "RUN_START", time.monotonic() - 80)       # the preparation ends past the limit - reserve
    with pytest.raises(rt.Unfinished) as e:
        mt.execute_ml_train(_train_args(data, SPLIT_X, "생산효율(%)"))
    assert "학습 준비" in e.value.progress and "첫 학습 단계 전에 멈춤" in e.value.progress, e.value.progress
    assert "약 27단계 남음" in e.value.progress, e.value.progress
    saved = os.listdir(root / ".resume_train")
    assert len(saved) == 1 and saved[0].endswith(".prep.joblib"), saved          # no start mark left
    assert not (root / "v1").exists() and not (root / "history.jsonl").exists()

    def no_second_preparation(*a, **k):
        raise AssertionError("준비를 다시 함")
    monkeypatch.setattr(mt, "_prepare_run", no_second_preparation)
    _new_run(monkeypatch)
    monkeypatch.setenv("ML_KIT_TIME_LIMIT", "0")
    monkeypatch.setattr(rt, "RUN_START", time.monotonic())
    got = mt.execute_ml_train(_train_args(data, SPLIT_X, "생산효율(%)"))
    assert got["version"] == "v1" and got["mean_val_score"] == want["mean_val_score"]
    assert got["top_features"] == want["top_features"]
    _assert_same_version(want["version_dir"], got["version_dir"])
    assert not (root / ".resume_train").exists()                                       # cleaned after the commit


def test_preparation_key_follows_data_and_arguments(tmp_path):
    data = make_sample(str(tmp_path / "k.csv"))
    xfile = tmp_path / "x.txt"
    xfile.write_text("반응온도(°C)\n소비전력(kW)\n", encoding="utf-8")
    a = _train_args(data, str(xfile), "생산효율(%)")
    group = {"features": str(xfile), "targets": "생산효율(%)", "eval": None, "dtypes": None, "positive_class": None,
             "exclude_where": None}
    k1 = mt._prep_key(a, "regression", [group], data)
    assert k1 == mt._prep_key(a, "regression", [group], data) and re.fullmatch(r"[0-9a-f]{16}", k1)
    assert rt._RESUME_FILE_RE.fullmatch(k1 + ".prep.joblib") and rt._RESUME_FILE_RE.fullmatch(k1 + ".prep.start")
    assert rt._RESUME_FILE_RE.fullmatch(k1 + ".m2.f3.start") and rt._RESUME_FILE_RE.fullmatch(k1 + ".fit.start.tmp")
    assert k1 != mt._prep_key(a, "classification", [group], data)
    assert k1 != mt._prep_key(_train_args(data, str(xfile), "생산효율(%)", header_row=1), "regression", [group], data)
    xfile.write_text("반응온도(°C)\n", encoding="utf-8")                               # the list file changed
    k2 = mt._prep_key(a, "regression", [group], data)
    assert k2 != k1
    with open(data, "a", encoding="utf-8-sig") as f:                                  # the data changed
        f.write("1,2,3,4,5,6\n")
    assert mt._prep_key(a, "regression", [group], data) != k2


def test_step_mark_stops_a_step_longer_than_the_limit(tmp_path, monkeypatch):
    """시계 없이 하는 단계(실행의 첫 단계)는 시작 표시를 남기고 끝나면 지운다. 표시만 남은 단계(이전 실행이 그 단계 도중 끊김)를
    다시 만나면 2번째까지는 다시 시도하고, 2번 끊긴 단계는 시작하지 않고 STOP_RETRY: [추천] 시간 제한을 늘린 같은 명령.
    시간 제한을 늘리면 다시 시도하고, 늘린 제한으로도 끊겼으면 [추천] 없이 멈춘다."""
    monkeypatch.setenv("ML_KIT_DATA_DIR", str(tmp_path / "data"))
    monkeypatch.setenv("TEMP", str(tmp_path / "temp"))
    monkeypatch.setattr(rt, "CURRENT_STEP", {})
    monkeypatch.setattr(rt, "_RESUME_PLACE", {})
    monkeypatch.setattr(rt, "CURRENT_ARGV", [os.path.join(KIT, "ml_train.py"), "--file", "big.csv", "--features", "a,b", "--targets", "y"])
    monkeypatch.setattr(rt, "TIME_LIMIT_ARG", [])
    monkeypatch.setenv("ML_KIT_TIME_LIMIT", "100")
    rdir = str(tmp_path / "rt")
    mark = tmp_path / "rt" / "0123456789abcdef.m1.f2.start"
    rt.step_begin(rdir, "0123456789abcdef.m1.f2", "학습: 타겟 'y' 'RandomForest' 교차검증 fold 2/5")
    assert json.loads(mark.read_text(encoding="utf-8"))["tries"] == 1
    rt.step_end()                                             # the step ended: no mark
    assert not mark.exists()
    for tries in (1, 2):                                      # two runs stopped inside the step (no step_end)
        rt.step_begin(rdir, "0123456789abcdef.m1.f2", "학습: 타겟 'y' 'RandomForest' 교차검증 fold 2/5")
        assert json.loads(mark.read_text(encoding="utf-8"))["tries"] == tries
    with pytest.raises(rt.InputError) as e:
        rt.step_begin(rdir, "0123456789abcdef.m1.f2", "학습: 타겟 'y' 'RandomForest' 교차검증 fold 2/5")
    msg = str(e.value)
    assert "한 단계(학습: 타겟 'y' 'RandomForest' 교차검증 fold 2/5)가 명령 시간 제한(100초) 안에 끝나지 않습니다" in msg, msg
    rec = msg.split("[추천] ")[1]
    assert rec.startswith("python ") and rec.endswith(" --time-limit 570") and '--features a,b --targets y' in rec, rec
    payload = rt.error_payload(e.value)
    assert payload["agent_directive"].startswith("STOP_RETRY:") and "timeout) 600000" in payload["agent_directive"]
    monkeypatch.setattr(rt, "TIME_LIMIT_ARG", [570.0])      # the [추천] command: a larger limit tries again
    rt.step_begin(rdir, "0123456789abcdef.m1.f2", "fold")
    rt.step_begin(rdir, "0123456789abcdef.m1.f2", "fold")
    with pytest.raises(rt.InputError) as e:
        rt.step_begin(rdir, "0123456789abcdef.m1.f2", "fold")
    assert "[추천]" not in str(e.value) and "시간 제한을 늘려도 끝나지 않았습니다" in str(e.value), str(e.value)
    monkeypatch.setattr(rt, "TIME_LIMIT_ARG", [0.0])        # no limit: no mark, never stopped
    rt.step_begin(rdir, "0123456789abcdef.m1.f2", "fold")
    assert rt.CURRENT_STEP.get("name") is None


def test_time_limit_argument_overrides_the_environment(monkeypatch):
    monkeypatch.setattr(rt, "TIME_LIMIT_ARG", [])
    monkeypatch.setenv("ML_KIT_TIME_LIMIT", "100")
    p = rt.JsonArgumentParser()
    p.add_argument("--file")
    ns = p.parse_args(["--file", "a.csv"])
    assert not hasattr(ns, "time_limit") and rt.time_limit() == 100.0
    p.parse_args(["--file", "a.csv", "--time-limit", "570"])
    assert rt.time_limit() == 570.0
    p.parse_args(["--time-limit", "0"])
    assert rt.time_limit() == float("inf")


_WATCHDOG_SCRIPT = r'''
import os, sys, time
sys.path.insert(0, sys.argv[1])
import ml_predict as mp
import _ml_runtime as rt
rt.WATCHDOG_GRACE = 0.5
p = rt.JsonArgumentParser()
p.add_argument("--mode")
p.add_argument("--rt")
a = p.parse_args(sys.argv[2:])

def body():
    if a.mode == "saved":
        rt.resume_save(a.rt, "0123456789abcdef.m0.f1.joblib", lambda t: open(t, "w").write("x"), "fold")
        rt.step_label("학습: 타겟 'y' 'RandomForest' 교차검증 fold 2/5")
    elif a.mode == "marked":
        rt.step_begin(a.rt, "0123456789abcdef.m0.f1", "학습: 타겟 'y' 'RandomForest' 교차검증 fold 1/5")
    elif a.mode == "fast":
        return {"status": "ok", "verified": True, "agent_directive": "VERIFIED: 끝"}
    time.sleep(30)                       # a slow step
    print("never", flush=True)
    return {"status": "ok", "verified": True, "agent_directive": "VERIFIED: 늦음"}

rt.run_cli(body)
'''


def test_watchdog_writes_one_json_line_at_the_limit(tmp_path, monkeypatch):
    """명령이 시간 제한 + 여유(WATCHDOG_GRACE) 뒤에도 끝나지 않으면 감시 스레드가 stdout에 JSON 1줄을 내고 끝낸다: 이 실행이 저장한
    단계가 있으면 CONTINUE, 표시한 단계를 1번 끊었으면 CONTINUE(다시 시도), 2번째로 끊기거나 아무것도 저장하지 않았으면
    STOP_RETRY(오류 로그에 기록). 제한 안에 끝난 명령은 감시가 끼어들지 않는다 (두 줄 출력 없음)."""
    script = tmp_path / "ml_slow.py"                         # an ml-kit command name: the watchdog runs
    script.write_text(_WATCHDOG_SCRIPT, encoding="utf-8")
    data = tmp_path / "data"
    monkeypatch.setenv("ML_KIT_DATA_DIR", str(data))
    monkeypatch.setenv("TEMP", str(tmp_path / "temp"))
    monkeypatch.setenv("ML_KIT_TIME_LIMIT", "100")
    rt = str(tmp_path / "rt")

    def run(mode, limit="1"):
        t0 = time.monotonic()
        proc = subprocess.run([sys.executable, str(script), KIT, "--mode", mode, "--rt", rt, "--time-limit", limit],
                              capture_output=True, text=True, encoding="utf-8", timeout=120)
        lines = [ln for ln in proc.stdout.splitlines() if ln.strip()]
        assert len(lines) == 1, proc.stdout + proc.stderr
        assert time.monotonic() - t0 < 20, proc.stderr                                # not the 30 s step
        return proc, json.loads(lines[0])

    proc, out = run("saved")
    assert proc.returncode == 0 and out["continue"] is True and out["agent_directive"].startswith("CONTINUE:"), out
    assert "'RandomForest' 교차검증 fold 2/5 도중 시간 제한(1초)에 닿아 멈춤" in out["progress"], out
    assert out["next_commands"][0].endswith("--time-limit 1") and "[STATUS]" in proc.stderr
    proc, out = run("marked")
    assert proc.returncode == 0 and out["continue"] is True and "다시 시도 1/2" in out["progress"], out
    proc, out = run("marked")
    assert proc.returncode == 1 and out["agent_directive"].startswith("STOP_RETRY:"), out
    assert "[추천]" in out["message"] and out["message"].rstrip().endswith("--time-limit 570"), out["message"]
    proc, out = run("marked", "2")                          # stopped twice under a smaller limit: tried again
    assert proc.returncode == 0 and "다시 시도 1/2" in out["progress"], out
    log = (data / "logs" / "errors.txt").read_text(encoding="utf-8")
    assert log.count("===== ") == 1 and "끝나지 않습니다" in log
    proc, out = run("none")
    assert proc.returncode == 1 and out["agent_directive"].startswith("STOP_RETRY:") and "[추천]" in out["message"]
    proc, out = run("fast")
    assert proc.returncode == 0 and out["agent_directive"] == "VERIFIED: 끝", out


_FIRE_AFTER_COMMIT = r'''
import os, sys, threading, time
sys.path.insert(0, sys.argv[1])
import ml_predict as mp
import _ml_runtime as rt
import ml_train as mt
_verify = mt.verify_version_dir

def slow_verify(*a, **k):                # the limit + grace timer elapses while the saved version is checked
    t = threading.Thread(target=rt._watchdog_fire)
    t.start()
    time.sleep(rt.WATCHDOG_HOLD_SECONDS + 4)
    return _verify(*a, **k)

mt.verify_version_dir = slow_verify
sys.argv = [sys.argv[0]] + sys.argv[2:]
mt.main()
'''


def test_watchdog_after_commit_answers_saved_not_continue(tmp_path, monkeypatch):
    """2.9.69 회귀: 버전을 저장(commit, 원장 기록)한 뒤 확인 도중 감시가 끝내면 CONTINUE가 아니라 ok + UNVERIFIED
    ('v1 저장됨, 확인하지 못함', version·version_dir). 예전에는 같은 명령을 다시 하라고 해서 처음부터 다시 학습해 v2가 또 생겼고,
    저장한 단계가 없으면 v1이 저장됐는데도 '한 단계가 너무 길다' 오류를 냈다."""
    data = make_sample(str(tmp_path / "w.csv"))
    script = tmp_path / "ml_firewd.py"
    script.write_text(_FIRE_AFTER_COMMIT, encoding="utf-8")
    monkeypatch.setenv("ML_KIT_DATA_DIR", str(tmp_path / "data"))
    monkeypatch.setenv("ML_KIT_TIME_LIMIT", "1000")
    proc = subprocess.run([sys.executable, str(script), KIT, "--file", data, "--features", "반응온도(°C),소비전력(kW)",
                           "--targets", "생산효율(%)"], capture_output=True, text=True, encoding="utf-8", timeout=600)
    lines = [ln for ln in proc.stdout.splitlines() if ln.strip()]
    assert len(lines) == 1, proc.stdout + proc.stderr
    out = json.loads(lines[0])
    assert proc.returncode == 0 and out["status"] == "ok" and not out.get("continue"), out
    assert out["agent_directive"].startswith("UNVERIFIED:") and "학습 버전 v1 저장됨, 확인하지 못함" in out["agent_directive"]
    assert out["version"] == "v1" and out["version_dir"].endswith("v1") and os.path.isdir(out["version_dir"])
    assert "next_commands" not in out and out["verified"] is False
    ledger = (tmp_path / "w_ml_data" / "history.jsonl").read_text(encoding="utf-8")
    assert len([ln for ln in ledger.splitlines() if ln.strip()]) == 1


def test_watchdog_payload_after_commit(monkeypatch):
    """저장한 결과가 같은 명령으로 다시 쓰이면(rerun_safe) CONTINUE 그대로,
    아니면 ok + UNVERIFIED. 아무것도 저장하지 않았으면 예전처럼 STOP_RETRY."""
    monkeypatch.setattr(rt, "RUN_SAVED", [])
    monkeypatch.setattr(rt, "RUN_COMMITTED", [])
    monkeypatch.setattr(rt, "CURRENT_STEP", {"label": "보고서 작성"})
    rt.mark_committed("학습 버전 v3", rerun_safe=True, version="v3")
    out, code = rt._watchdog_payload()
    assert code == 0 and out["continue"] is True, out
    rt.mark_committed("보고서 r.docx", report_file="r.docx")
    out, code = rt._watchdog_payload()
    assert code == 0 and out["agent_directive"].startswith("UNVERIFIED:") and out["report_file"] == "r.docx", out
    assert [s["what"] for s in out["saved"]] == ["학습 버전 v3", "보고서 r.docx"] and "continue" not in out
    rt.RUN_COMMITTED.clear()
    out, code = rt._watchdog_payload()
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:"), out


def test_step_mark_is_removed_when_the_run_fails_by_itself(tmp_path, monkeypatch):
    """입력 오류로 끝난 실행은 시작 표시를 남기지 않는다 (표시는 밖에서 끊긴 실행만): 같은 오류가 반복돼도 '시간 제한' 오류가 되지 않음."""
    data = make_sample(str(tmp_path / "e.csv"))
    monkeypatch.setenv("ML_KIT_TIME_LIMIT", "100")
    for _ in range(3):
        out = run_fail("ml_train.py", "--file", data, "--features", "없는열", "--targets", "생산효율(%)")
        assert "시간 제한" not in out["message"] and "없는열" in out["message"], out
    assert not (tmp_path / "e_ml_data" / ".resume_train").exists()


# ---- 구분자 실수 (2.9.69): ',' 대신 ';' 또는 그 반대 -> [추천] 고친 전체 명령 ------------------------------------


def _sep_data(tmp_path):
    rng = np.random.default_rng(0)
    n = 60
    t, p = rng.uniform(150, 200, n), rng.uniform(900, 1200, n)
    path = tmp_path / "sep.csv"
    pd.DataFrame({"라인": rng.choice(["A", "B"], n), "온도": t.round(1), "압력": p.round(0),
                  "수율": (0.3 * t + 0.01 * p + rng.normal(0, 1, n)).round(2)}).to_csv(path, index=False)
    return str(path)


def test_exclude_where_with_comma_between_conditions_stops_with_fixed_command(tmp_path):
    """회귀: --exclude-where "라인 = A, 온도 > 180"은 조건 하나(라인 = 'A, 온도 > 180')로 읽혀 0행을 빼고 VERIFIED로 학습했다.
    이제 STOP_RETRY + [추천] ';'로 고친 전체 명령. 그 명령은 그대로 실행되어 조건대로 행을 뺀다."""
    data = _sep_data(tmp_path)
    proc, out = run_raw("ml_train.py", "--file", data, "--features", "온도,압력", "--targets", "수율", "--exclude-where", "라인 = A, 온도 > 180")
    assert proc.returncode == 1 and out["agent_directive"].startswith("STOP_RETRY:"), out
    rec = recommended_command(out["message"])
    assert rec[0] == "ml_train.py" and rec[1][rec[1].index("--exclude-where") + 1] == "라인 = A; 온도 > 180", rec
    res = run_tool(rec[0], *rec[1])
    assert res["exclusions"]["수율"]["condition"] > 0, res


def test_where_comma_detection_leaves_real_values():
    """','가 조건 사이일 때만 잡는다: 천 단위 숫자, 따옴표 값, 이름 없는 텍스트 값은 그대로."""
    assert rt.parse_where("온도 > 1,500") == [[("온도", ">", "1,500")]]
    assert rt.parse_where("비고 = 'A, B > 3'") == [[("비고", "=", "A, B > 3")]]
    assert rt.parse_where("비고 = A, B") == [[("비고", "=", "A, B")]]
    for text, fixed in (("온도 <= 190, 압력 >= 1000", "온도 <= 190; 압력 >= 1000"),
                        ("온도 > 1,500, 2(mm) < 3 | 라인 = B", "온도 > 1,500; 2(mm) < 3 | 라인 = B")):
        with pytest.raises(rt.InputError) as e:
            rt.parse_where(text, "--where")
        assert "','로 나눴습니다" in str(e.value) and recommendation(str(e.value)) == fixed, str(e.value)


def test_eval_dtypes_features_separator_slips_give_fixed_command(tmp_path):
    """--eval 식 사이 ',', --dtypes 쌍 사이 ';', --features 이름 사이 ';' -> STOP_RETRY [추천] 고친 전체 명령 (예전: 문법 오류,
    '형식을 알 수 없음', [후보])."""
    data = _sep_data(tmp_path)
    base = ["--file", data, "--features", "온도,압력", "--targets", "수율"]
    for extra, opt, want in ((["--eval", "a = 온도/압력, b = max(온도, 압력)"], "--eval", "a = 온도/압력; b = max(온도, 압력)"),
                             (["--dtypes", "온도:float;압력:float"], "--dtypes", "온도:float,압력:float")):
        proc, out = run_raw("ml_train.py", *base, *extra)
        assert proc.returncode == 1 and "[추천] python" in out["message"], out
        args = recommended_command(out["message"])[1]
        assert args[args.index(opt) + 1] == want, args
    proc, out = run_raw("ml_train.py", "--file", data, "--features", "온도;압력", "--targets", "수율")
    args = recommended_command(out["message"])[1]
    assert args[args.index("--features") + 1] == "온도,압력" and args.count("--features") == 1, args
    assert mt._eval_comma_fix("a = where(온도 > 1, 1, 0)") == "a = where(온도 > 1, 1, 0)"
    assert mt._eval_comma_fix("a = x == y, b = 1") == "a = x == y; b = 1"


def test_dry_run_shows_version_and_replaced_folder_without_training(tmp_path):
    """2.9.70: --dry-run은 학습하지 않고 만들 버전과 (--overwrite면) 바꿀 기존 폴더의 파일을 보여 주고 confirm_commands에 실제 명령."""
    data = make_sample(str(tmp_path / "d.csv"))
    root = tmp_path / "d_ml_data"
    res = run_tool("ml_train.py", "--file", data, "--features", FEATURES, "--targets", TARGET, "--dry-run")
    assert res["dry_run"] is True and res["version"] == "v1" and res["replaces"] is False and not root.exists(), res
    cmd = res["confirm_commands"][0]
    assert cmd["destructive"] is False and "--dry-run" not in split_command(cmd["command"])
    (root / "v1").mkdir(parents=True)
    (root / "v1" / "model_a.joblib").write_text("x", encoding="utf-8")
    msg = run_fail("ml_train.py", "--file", data, "--features", FEATURES, "--targets", TARGET, "--version", "v1", "--dry-run")["message"]
    rec = split_command(msg.split("[추천] ", 1)[1])
    assert rec[rec.index("--version") + 1] == "v2" and rec[0] == "python", msg
    res = run_tool("ml_train.py", "--file", data, "--features", FEATURES, "--targets", TARGET, "--version", "v1", "--overwrite", "--dry-run")
    assert res["replaces"] is True and res["replaced_files"] == ["model_a.joblib"] and res["confirm_commands"][0]["destructive"] is True
    argv = split_command(res["confirm_commands"][0]["command"])
    assert "--overwrite" in argv and "--dry-run" not in argv
    assert sorted(os.listdir(root)) == ["v1"] and os.listdir(root / "v1") == ["model_a.joblib"]      # 아무것도 바꾸지 않음


def test_tool_stops_offering_retraining_after_rounds_without_gain(tmp_path):
    """2.9.71: 개선 반복 횟수는 도구가 센다. 같은 타겟 학습이 IMPROVE_ROUNDS(5)번 연속 이전 최고를 넘지 못하면
    재학습 next_commands 대신 보고서 명령 하나와 improve_stopped (LLM이 5번을 세지 않음)."""
    data = make_sample(str(tmp_path / "d.csv"))
    x = "반응온도(°C),소비전력(kW),압축압력(bar)"
    for i in range(mt.IMPROVE_ROUNDS):
        res = run_tool("ml_train.py", "--file", data, "--features", x, "--targets", TARGET)
        assert res["goal_met"] is False and "improve_stopped" not in res, (i, res.get("improve_stopped"))
    res = run_tool("ml_train.py", "--file", data, "--features", x, "--targets", TARGET)
    assert "improve_stopped" in res and len(res["next_commands"]) == 1, res
    argv = split_command(res["next_commands"][0]["command"])
    assert os.path.basename(argv[1]) == "ml_report.py" and argv[argv.index("--file") + 1] == data, argv


# ---- 2.9.72: --features 생략 시 입력 자동 선택 (뺀 열과 이유), --features를 준 때의 의심 입력 경고 -------------------------------


def _auto_rows(n=60, seed=0, noise=8.0):
    """시료번호(ID), 일자, 숫자 입력 2개, 라인(범주), 설비(상수), 비고(메모), 측정자(빈 칸), 수율(%) 타겟,
    수율(분율)(타겟을 다른 단위로 적은 열), 제외(표시 열)."""
    r = np.random.default_rng(seed)
    t, p = r.uniform(150, 200, n), r.uniform(1, 5, n)
    line = r.choice(["A", "B"], n)
    y = 0.3 * t + 5 * p + np.where(line == "A", 3, 0) + r.normal(0, noise, n)
    return pd.DataFrame({
        "시료번호": np.arange(1001, 1001 + n), "일자": pd.date_range("2026-01-01", periods=n).strftime("%Y-%m-%d"),
        "반응온도(°C)": t.round(1), "압력(bar)": p.round(2), "라인": line, "설비": ["E1"] * n,
        "비고": [f"메모 {i} 확인" for i in range(n)], "측정자": [None] * (n - 5) + ["김"] * 5,
        "수율(%)": y.round(2), "수율(분율)": (y / 100).round(5), "제외": [""] * n})


@pytest.fixture(scope="module")
def auto_run(tmp_path_factory):
    tmp = tmp_path_factory.mktemp("auto_x")
    f = write_df(os.path.join(str(tmp), "auto.csv"), _auto_rows())
    res = run_tool("ml_train.py", "--file", f, "--targets", "수율(%)")
    return {"f": f, "res": res}


def test_auto_inputs_excluded_with_reasons(auto_run):
    res = auto_run["res"]
    assert res["features_auto"] == ["일자", "반응온도(°C)", "압력(bar)", "라인"]
    assert res["features_auto_count"] == 4
    why = {d["column"]: d["reason"] for d in res["excluded_inputs"]}
    assert list(why) == ["시료번호", "설비", "비고", "측정자", "수율(분율)", "제외"]
    assert "ID" in why["시료번호"] and "상수" in why["설비"] and "종류" in why["비고"]
    assert "빈 칸" in why["측정자"] and "타겟 '수율(%)'" in why["수율(분율)"] and "단위" in why["수율(분율)"]
    assert "표시" in why["제외"]
    assert res["excluded_inputs_count"] == 6
    assert "입력 자동 선택 4개" in res["preview"] and "시료번호(ID·일련번호)" in res["preview"]
    assert "features_auto" in res["agent_directive"] and "excluded_inputs" in res["agent_directive"]
    with open(res["features_auto_file"], encoding="utf-8") as fh:
        assert fh.read().split() == res["features_auto"]
    rec = mt.load_history_ledger(os.path.join(res["data_root"], "history.jsonl"))[-1]
    assert rec["features_used"] == res["features_auto"]
    assert rec["input_selection"]["auto"] is True and len(rec["input_selection"]["excluded"]) == 6


def test_auto_inputs_next_commands_carry_x(auto_run):
    """다시 학습 명령은 고른 입력을 --features로 적음: 같은 명령을 다시 실행하면 같은 입력 (재현)."""
    res = auto_run["res"]
    assert res["goal_met"] is False
    train_cmds = [c["command"] for c in res["next_commands"] if "ml_train.py" in c["command"] and not c["needs_input"]]
    assert train_cmds
    for cmd in train_cmds:
        args = split_command(cmd)
        assert "--features" in args, cmd
    args = split_command(train_cmds[-1])
    x = args[args.index("--features") + 1]
    names = open(x, encoding="utf-8").read().split() if os.path.isfile(x) else x.split(",")
    assert set(names) <= set(res["features_auto"]) and "시료번호" not in names
    again = run_tool(os.path.basename(args[1]), *args[2:])
    assert "features_auto" not in again
    rec = mt.load_history_ledger(os.path.join(again["data_root"], "history.jsonl"))[-1]
    assert [f for f in rec["features_used"] if f in res["features_auto"]] == [f for f in res["features_auto"] if f in names]


@needs_report
@pytest.mark.report
def test_auto_inputs_report_uses_the_last_run(auto_run):
    """입력을 자동으로 고른 학습이 있는 이력에서도 ml_report --file은 다시 학습하지 않고 마지막 학습으로 보고서를 만든다."""
    path = os.path.join(auto_run["res"]["data_root"], "history.jsonl")
    before = mt.load_history_ledger(path)
    rep = run_tool("ml_report.py", "--file", auto_run["f"])
    assert rep["version"] == before[-1]["version"] and rep["agent_directive"].startswith("VERIFIED:"), rep
    assert len(mt.load_history_ledger(path)) == len(before)


def test_explicit_x_warns_id_and_leak(tmp_path):
    """--features를 주면 그대로 학습하고 ID·타겟 누수로 보이는 입력은 경고만 (빼지 않음)."""
    f = write_df(os.path.join(str(tmp_path), "given.csv"), _auto_rows(noise=0.5))
    res = run_tool("ml_train.py", "--file", f, "--targets", "수율(%)", "--features", "시료번호,반응온도,압력,수율(분율)")
    sus = {d["column"]: d["reason"] for d in res["suspect_inputs"]}
    assert set(sus) == {"시료번호", "수율(분율)"} and "ID" in sus["시료번호"] and "타겟" in sus["수율(분율)"]
    assert any("시료번호" in w and "입력으로 두었습니다" in w for w in res["warnings"])
    assert "의심 입력" in res["preview"] and "suspect_inputs" in res["agent_directive"]
    assert "features_auto" not in res
    rec = mt.load_history_ledger(os.path.join(res["data_root"], "history.jsonl"))[-1]
    assert "시료번호" in rec["features_used"] and "수율(분율)" in rec["features_used"]


def test_auto_inputs_classification_leak_and_codes(tmp_path):
    """분류: 타겟 종류와 1:1로 맞는 코드 열은 누수, 'S-001' 코드 열은 ID, 실측값이 모두 다른 숫자 열은 입력."""
    r = np.random.default_rng(3)
    n = 50
    t = r.uniform(150, 200, n).round(2)
    grade = np.where(t > 175, "상", "하")
    df = pd.DataFrame({"코드": [f"S-{i:03d}" for i in range(n)], "반응온도": t, "압력": r.uniform(1, 5, n).round(2),
                       "등급코드": np.where(grade == "상", 1, 2), "등급": grade})
    f = write_df(os.path.join(str(tmp_path), "cls.csv"), df)
    res = run_tool("ml_train.py", "--file", f, "--targets", "등급", "--task", "classification")
    why = {d["column"]: d["reason"] for d in res["excluded_inputs"]}
    assert set(why) == {"코드", "등급코드"} and "ID" in why["코드"] and "1:1" in why["등급코드"]
    assert res["features_auto"] == ["반응온도", "압력"]


def test_auto_inputs_none_left_stops(tmp_path):
    n = 20
    df = pd.DataFrame({"번호": range(1, n + 1), "상수": [1] * n, "y": np.linspace(0, 1, n)})
    f = write_df(os.path.join(str(tmp_path), "none.csv"), df)
    res = run_fail("ml_train.py", "--file", f, "--targets", "y")
    assert res["agent_directive"].startswith("STOP_RETRY:") and "--features" in res["message"]
    assert {d["column"] for d in res["excluded_inputs"]} == {"번호", "상수"}


def test_diagnose_keeps_measured_unique_numbers_and_dates():
    """실측값처럼 모두 다른 소수, 날짜, 종류가 적은 글자 열은 입력으로 남음 (판정 함수)."""
    n = 40
    r = np.random.default_rng(5)
    df = pd.DataFrame({"두께": r.uniform(1, 2, n).round(4), "일자": pd.date_range("2026-01-01", periods=n),
                       "라인": r.choice(["A", "B", "C"], n), "속도": np.arange(40, 40 + n) * 5,
                       "y": r.normal(0, 1, n)})
    assert mt.diagnose_inputs(df, ["두께", "일자", "라인", "속도"], ["y"], "regression") == []


def test_next_commands_offer_the_exclude_where_template(tmp_path):
    """2.9.81 집 PC Gemma A3: '생산효율이 80보다 큰 행은 빼고 다시 학습해줘'에 긴 학습 명령을 답변에 적기만 하고 실행하지 않음 ->
    학습 결과에 같은 학습의 --exclude-where '<조건>' 틀(needs_input, 맨 뒤). 앞의 조건이 있으면 '<앞 조건>; <조건>'으로 지킴."""
    data = make_multi(str(tmp_path / "ex.csv"))
    res = run_tool("ml_train.py", "--file", data, "--features", "온도,압력,속도", "--targets", "수율")
    tpl = [c for c in res["next_commands"] if c.get("placeholder") == "<조건>"]
    assert len(tpl) == 1 and tpl[0]["needs_input"] is True and res["next_commands"][-1] is tpl[0], res["next_commands"]
    argv = split_command(tpl[0]["command"].replace("<조건>", "수율 > 50"))
    assert argv[argv.index("--features") + 1] == "온도(°C),압력(bar),속도(m/s)" and argv[argv.index("--exclude-where") + 1] == "수율 > 50", argv
    again = run_tool("ml_train.py", *argv[2:])
    tpl = [c for c in again["next_commands"] if c.get("placeholder") == "<조건>"][0]
    argv = split_command(tpl["command"])
    assert argv[argv.index("--exclude-where") + 1] == "수율(%) > 50; <조건>", argv     # 앞 조건은 도구가 맞춘 열 이름으로


def test_history_latest_version_skips_a_damaged_folder(tmp_path):
    """2.9.83: --history의 latest_version은 보고서·예측과 같은 최종 버전 (마지막 학습 폴더가 손상되면 그 앞 버전, latest_version_note)."""
    data = make_multi(str(tmp_path / "h.csv"))
    run_tool("ml_train.py", "--file", data, "--features", "온도,압력", "--targets", "수율")
    v2 = run_tool("ml_train.py", "--file", data, "--features", "온도,압력,속도", "--targets", "수율")
    os.remove(os.path.join(v2["version_dir"], "leaderboard.json"))
    h = run_tool("ml_train.py", "--file", data, "--history")
    assert h["latest_version"] == "v1" and "v2" in h["latest_version_note"] and "v2" in h["preview"], h
    p = run_tool("ml_predict.py", "--model", data, "--input", "온도=200,압력=3")
    assert "v1" in p["model_chosen"] and "v2" in p["model_chosen"] and "손상" in p["model_chosen"], p
