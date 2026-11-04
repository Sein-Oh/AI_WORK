# -*- coding: utf-8 -*-
"""학습(ml_train.py): 피처 유형, 분류, 작은 데이터와 중복 행의 교차검증, 학습 제외 행, --eval, next_commands, --history,
--best, 중요 입력(top_features).
"""
import argparse
import json
import os
import random
import shlex
import subprocess
import sys

import numpy as np
import pandas as pd
import pytest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from make_sample import make_class_sample, make_sample  # noqa: E402
from mlkit_helpers import HAS_DOCX, HAS_MPL, KIT, run_fail, run_raw, run_tool, write_df, write_rows  # noqa: E402
import ml_predict as mp  # noqa: E402
import ml_train as mt  # noqa: E402

try:
    import docx
except ImportError:          # --best 보고서 검사(needs_report)는 건너뜀
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
    """새 형식 파이프라인(ColumnTransformer)도 ml_predict.py만으로 불러 예측한다."""
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
    """next_commands의 피처 줄이기는 한 타겟에 약해도 다른 타겟에 중요한 피처를 빼지 않는다."""
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


# ---- 분류 (--task classification) ------------------------------------------------------------


FEATS = "반응온도,촉매비율,수분함량"


@pytest.fixture(scope="module")
def cls_data(tmp_path_factory):
    work = tmp_path_factory.mktemp("cls")
    return work, make_class_sample(str(work / "proc.csv"))


@pytest.fixture(scope="module")
def multiclass(cls_data):
    work, data = cls_data
    payload = run_tool("ml_train.py", "-f", data, "-x", FEATS, "-y", "품질등급",
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
    payload = run_tool("ml_predict.py", "-m", str(work / "proc_ml_data" / "v1"),
                       "-i", "반응온도=250,촉매비율=1.0,수분함량=18")
    ps = payload["predictions_summary"]["품질등급"]
    assert ps["task"] == "classification"
    assert ps["predictions"][0] in ps["classes"]
    assert abs(sum(ps["probabilities"][0].values()) - 1.0) < 1e-3
    assert "low_confidence_warning" in payload


def test_multiclass_predict_batch_columns(multiclass):
    import pandas as pd
    work, data, _ = multiclass
    out = work / "pred.csv"
    run_tool("ml_predict.py", "-m", str(work / "proc_ml_data" / "v1"), "-f", data, "-o", str(out))
    df = pd.read_csv(out, encoding="cp949")
    for col in ["pred_품질등급", "conf_품질등급", "proba_품질등급_A등급", "proba_품질등급_C등급"]:
        assert col in df.columns
    assert (df["pred_품질등급"] == df["품질등급"]).mean() > 0.6


def test_binary_integer_labels(cls_data):
    work, data = cls_data
    payload = run_tool("ml_train.py", "-f", data, "-x", FEATS, "-y", "불량여부", "--task", "classification")
    meta = json.loads((work / "proc_ml_data" / payload["version"] / "model_불량여부.json").read_text(encoding="utf-8"))
    assert meta["classes"] == ["0", "1"]  # integer labels normalized to stable strings
    lb = json.loads((work / "proc_ml_data" / payload["version"] / "leaderboard.json").read_text(encoding="utf-8"))
    assert lb["targets"]["불량여부"]["champion"]["roc_auc"] is not None


def test_imbalanced_uses_balanced_weights(tmp_path):
    data = make_class_sample(str(tmp_path / "imb.csv"), n=200, imbalanced=True)
    proc, payload = run_raw("ml_train.py", "-f", data, "-x", FEATS, "-y", "품질등급", "--task", "classification")
    assert payload["status"] == "ok"
    assert "class_weight='balanced'" in proc.stderr
    assert any("불균형" in r for r in payload["recommendations"])


def test_string_target_without_task_hints_classification(cls_data):
    _, data = cls_data
    payload = run_fail("ml_train.py", "-f", data, "-x", FEATS, "-y", "품질등급")
    assert "--task classification" in payload["message"]


def test_single_sample_class_rejected(tmp_path):
    import pandas as pd
    p = tmp_path / "rare.csv"
    make_class_sample(str(p))
    df = pd.read_csv(p, encoding="utf-8-sig")
    df.loc[0, "품질등급"] = "D등급"
    df.to_csv(p, index=False, encoding="utf-8-sig")
    payload = run_fail("ml_train.py", "-f", str(p), "-x", FEATS, "-y", "품질등급", "--task", "classification")
    assert "D등급" in payload["message"]


def test_history_mixed_tasks_compare_within_task(multiclass):
    work, data, _ = multiclass
    run_tool("ml_train.py", "-f", data, "-x", "반응온도,촉매비율", "-y", "수분함량", "-v", "v9")  # regression on same ledger
    hist = run_tool("ml_train.py", "-f", data, "--history")
    tasks = {r["version"]: r["task"] for r in hist["records"]}
    assert tasks["v1"] == "classification" and tasks["v9"] == "regression"
    assert set(hist["best_by_task"]) == {"classification", "regression"}


def test_report_supports_classification(multiclass):
    """분류 버전도 통합 보고서를 만든다 (-d, -t, -v)."""
    pytest.importorskip("docx")
    pytest.importorskip("matplotlib")
    work, _, _ = multiclass
    payload = run_tool("ml_report.py", "-d", str(work / "proc_ml_data"), "-t", "품질등급", "-v", "v1")
    assert payload["task"] == "classification" and payload["targets"] == ["품질등급"]


def test_pre_2_0_model_metadata_is_regression(tmp_path):
    """'task' 키가 없는 모델 메타(2.0.0 전 형식)는 회귀로 예측한다."""
    from make_sample import make_sample
    data = make_sample(str(tmp_path / "old.csv"))
    run_tool("ml_train.py", "-f", data, "-x", "반응온도,소비전력,압축압력,촉매비율,수분함량", "-y", "생산효율")
    meta_p = tmp_path / "old_ml_data" / "v1" / "model_생산효율(%).json"
    meta = json.loads(meta_p.read_text(encoding="utf-8"))
    for k in ("task", "primary_metric", "val_score"):
        meta.pop(k)
    meta_p.write_text(json.dumps(meta, ensure_ascii=False), encoding="utf-8")
    payload = run_tool("ml_predict.py", "-m", str(tmp_path / "old_ml_data" / "v1"),
                       "-i", "반응온도=240,소비전력=65,압축압력=3,촉매비율=0.8,수분함량=20")
    assert payload["predictions_summary"]["생산효율(%)"]["task"] == "regression"


def test_delta_compares_same_targets_only(tmp_path):
    """같은 원장의 다른 타겟은 비교 기준(delta_score)이 되지 않는다."""
    data = make_class_sample(str(tmp_path / "dt.csv"))
    run_tool("ml_train.py", "-f", data, "-x", FEATS, "-y", "불량여부", "--task", "classification")
    second = run_tool("ml_train.py", "-f", data, "-x", FEATS, "-y", "품질등급", "--task", "classification")
    assert second["delta_score"] == "base" and second["is_best"] is True
    third = run_tool("ml_train.py", "-f", data, "-x", FEATS, "-y", "품질등급", "--task", "classification", "--seed", "7")
    assert isinstance(third["delta_score"], float)
    hist = run_tool("ml_train.py", "-f", data, "--history")
    assert hist["best_version"] in ("v2", "v3")  # never v1 (other target)


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


def test_constant_regression_target_is_refused(tmp_path):
    p = tmp_path / "const.csv"
    pd.DataFrame({"x": np.arange(20, dtype=float), "y": [5.0] * 20}).to_csv(p, index=False)
    bad = run_fail("ml_train.py", "-f", str(p), "-x", "x", "-y", "y")
    assert bad["agent_directive"].startswith("STOP_RETRY") and "모두 같아(5.0)" in bad["message"]


def test_class_with_one_distinct_row_falls_back(tmp_path):
    r = np.random.default_rng(0)
    df = pd.DataFrame({"a": r.normal(size=30), "b": r.normal(size=30), "cls": "B"})
    df = pd.concat([df, pd.DataFrame({"a": [5.0, 5.0], "b": [5.0, 5.0], "cls": ["A", "A"]})], ignore_index=True)
    res = run_tool("ml_train.py", "-f", write_df(tmp_path / "c1.csv", df), "-x", "a,b", "-y", "cls",
                   "--task", "classification")                                 # 2.9.13: ERROR shape mismatch
    assert any("일반 층화 교차검증" in w for w in res["warnings"])
    # multiclass: the single-row class is no longer silently never predicted (recall 0)
    df3 = pd.DataFrame({"a": r.normal(size=40), "b": r.normal(size=40), "cls": np.where(np.arange(40) % 2, "B", "C")})
    df3 = pd.concat([df3, pd.DataFrame({"a": [9.0] * 3, "b": [9.0] * 3, "cls": ["A"] * 3})], ignore_index=True)
    res3 = run_tool("ml_train.py", "-f", write_df(tmp_path / "c3.csv", df3), "-x", "a,b", "-y", "cls",
                    "--task", "classification")
    assert any("일반 층화 교차검증" in w for w in res3["warnings"])
    meta = json.load(open(os.path.join(res3["version_dir"], "model_cls.json"), encoding="utf-8"))
    lb = json.load(open(os.path.join(res3["version_dir"], "leaderboard.json"), encoding="utf-8"))
    assert lb["targets"]["cls"]["champion"]["per_class"]["A"]["recall"] > 0
    assert meta["task"] == "classification"


def test_regression_too_few_distinct_rows_stops(tmp_path):
    df = pd.DataFrame({"a": [1.0] * 8 + [2, 3, 4, 5], "b": [1.0] * 8 + [2, 1, 5, 3], "y": [1.0] * 8 + [2, 3, 4, 2]})
    bad = run_fail("ml_train.py", "-f", write_df(tmp_path / "d2.csv", df), "-x", "a,b", "-y", "y")
    assert bad["agent_directive"].startswith("STOP_RETRY") and "입력+타겟이 다른 행이 5개뿐" in bad["message"]


def test_regression_ten_distinct_rows_with_many_copies(tmp_path):
    r = np.random.default_rng(1)
    base = pd.DataFrame({"a": r.uniform(0, 1, 10), "b": r.uniform(0, 1, 10)})
    base["y"] = base.a + base.b
    df = pd.concat([base] + [base.iloc[[0]]] * 30, ignore_index=True)
    res = run_tool("ml_train.py", "-f", write_df(tmp_path / "d10.csv", df), "-x", "a,b", "-y", "y")   # never ERROR
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
    res = run_tool("ml_train.py", "-f", write_df(tmp_path / "c5x2.csv", df), "-x", "a,b", "-y", "c", "--task", "classification")
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
    res = run_tool("ml_train.py", "-f", str(f), "-x", "반응온도,압축압력,촉매비율,제외", "-y", "생산효율")
    ex = res["exclusions"]["생산효율"]
    assert (ex["manual"], ex["auto"], ex["forced_include"]) == (1, 2, 1)
    assert isinstance(ex["all_rows_score"], float) and ex["all_rows_score"] < res["mean_val_score"]
    assert "제외" not in res["feature_types"]["numeric"]                 # never a feature
    assert "검증 R2(제외 행 포함)" in res["preview"] and "검증 R2=" in res["preview"]
    assert f.read_bytes() == before                                       # data file untouched

    rows = pd.read_csv(res["excluded_rows_file"], encoding="cp949")
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


def test_no_exclude_column_still_finds_outliers_and_clean_data_has_no_key(tmp_path):
    f = tmp_path / "d.csv"
    _excl_frame().to_csv(f, index=False)
    res = run_tool("ml_train.py", "-f", str(f), "-x", "반응온도,압축압력,촉매비율", "-y", "생산효율")
    assert res["exclusions"]["생산효율"]["auto"] == 3 and res["exclusions"]["생산효율"]["column"] is None

    g = tmp_path / "clean.csv"
    _excl_frame(spikes=False).to_csv(g, index=False)
    res = run_tool("ml_train.py", "-f", str(g), "-x", "반응온도,압축압력,촉매비율", "-y", "생산효율")
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
    res = run_tool("ml_train.py", "-f", str(f), "-x", "반응온도,압축압력,촉매비율", "-y", "등급",
                   "--task", "classification", "--positive-class", "불량")
    ex = res["exclusions"]["등급"]
    assert ex["manual"] == 1 and ex["auto"] == 1                          # target is a label: features only
    assert 0 <= ex["all_rows_score"] <= 1 and "검증 Macro-F1(제외 행 포함)" in res["preview"]


def test_exclude_column_as_target_is_input_error(tmp_path):
    f = tmp_path / "d.csv"
    _excl_frame(_excl_marks()).to_csv(f, index=False)
    payload = run_fail("ml_train.py", "-f", str(f), "-x", "반응온도", "-y", "제외")
    assert payload["agent_directive"].startswith("STOP_RETRY") and "제외" in payload["message"]


def test_exclude_column_zero_and_false_are_not_marks():
    import ml_train as mt
    assert [mt._exclude_mark(v) for v in (0, 0.0, "0", "FALSE", False, None, float("nan"), "", " 센서오류 ", 1, "포함")] == \
        ["", "", "", "", "", "", "", "", "센서오류", "1", "포함"]


def test_report_shows_terms_and_excluded_rows(tmp_path):
    from docx import Document
    f = tmp_path / "d.xlsx"
    _excl_frame(_excl_marks()).to_excel(f, index=False)
    res = run_tool("ml_train.py", "-f", str(f), "-x", "반응온도,압축압력,촉매비율", "-y", "생산효율")
    rep = run_tool("ml_report.py", "-m", res["version_dir"])
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
    tok, mapping, canon = mp.compile_formula("log(소비전력) / 반응온도(°C)", cols)
    assert canon == "log(소비전력(kW)) / 반응온도(°C)"
    assert set(mapping.values()) == {"소비전력(kW)", "반응온도(°C)"}


def test_eval_blocks_attribute_access():
    df = pd.DataFrame({"a": [1.0, 2.0]})
    with pytest.raises(mp.InputError):
        mp.apply_eval_expressions(df, ["b = a.__class__"])


# 파생 변수의 원본 열이 -x에 없어도(--eval "온도상한 = clip(온도, -1e12, 250)") 모델이 그 열을 입력으로 기록해
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
    res = run_tool("ml_train.py", "-f", str(path), "-x", "압력", "--eval", "온도상한 = clip(온도, -1e12, 250)", "-y", "효율")
    return d, res["version_dir"]


def test_source_column_is_recorded(capped):
    _, vdir = capped
    meta = json.load(open(os.path.join(vdir, "model_효율.json"), encoding="utf-8"))
    assert meta["input_columns_required"] == ["압력", "온도(°C)"] and "온도(°C)" in meta["feature_stats"]
    snap = pd.read_csv(os.path.join(vdir, "model_dataset.csv"), encoding="cp949")
    assert "온도(°C)" in snap.columns and "온도상한" in snap.columns


def test_prediction_uses_source_and_caps(capped):
    _, vdir = capped
    hi = run_tool("ml_predict.py", "-m", vdir, "-i", "온도=280,압력=2")["predictions_summary"]["효율"]["predictions"][0]
    cap = run_tool("ml_predict.py", "-m", vdir, "-i", "온도=250,압력=2")["predictions_summary"]["효율"]["predictions"][0]
    assert hi == pytest.approx(cap, abs=1e-6)                                   # above 250 is treated as 250
    miss = run_fail("ml_predict.py", "-m", vdir, "-i", "압력=2")
    assert "'온도(°C)'" in miss["message"] and "원본 입력" in miss["message"]


def test_find_and_export_work_on_the_source(capped, tmp_path):
    _, vdir = capped
    f = run_tool("ml_predict.py", "-m", vdir, "--find", "효율=80", "--vary", "온도", "-i", "압력=2")
    assert f["count"] >= 1 and abs(f["solutions"][0]["온도(°C)"] - (80 - 10) / 0.3) < 2
    e = run_tool("ml_export.py", "-m", vdir, "--inputs", "온도,압력", "--save", str(tmp_path / "c.xlsx"))
    assert e["targets"]["효율"]["inputs"] == ["온도(°C)", "압력"]


# ---- next_commands --------------------------------------------------------------------------------------


@pytest.fixture(scope="module")
def sample_model(tmp_path_factory):
    """make_sample 데이터의 입력 5개 -> 생산효율 학습 (next_commands 검사 둘이 같이 씀)."""
    data = make_sample(str(tmp_path_factory.mktemp("nc") / "nc.csv"))
    payload = run_tool("ml_train.py", "-f", data, "-x", "반응온도,소비전력,압축압력,촉매비율,수분함량", "-y", "생산효율")
    return data, payload


def test_next_command_runs(sample_model):
    data, payload = sample_model
    assert payload["next_commands"], payload
    cmd = payload["next_commands"][0]["command"].replace("<새데이터.xlsx>", data)
    assert cmd.startswith("python ")
    args = shlex.split(cmd)[1:]
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
    payload = run_tool("ml_train.py", "-f", cls, "-x", "반응온도,촉매비율,수분함량", "-y", "품질등급", "--task", "classification")
    if payload["mean_val_score"] < 0.9:
        assert payload["next_commands"] and all(c["needs_input"] is False for c in payload["next_commands"])
    assert not any("성능예측표" in c["purpose"] for c in payload["next_commands"])   # 분류는 성능예측표 없음


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


# ---- --history ----------------------------------------------------------------------------------------------


def test_history_has_verified(coded):
    res = run_tool("ml_train.py", "-f", coded["file"], "--history")
    assert res["verified"] is True and res["agent_directive"].startswith("VERIFIED")


def test_history_best_version_covers_every_target(tmp_path):
    r = np.random.default_rng(7)
    n = 80
    df = pd.DataFrame({"a": r.uniform(0, 1, n), "b": r.uniform(0, 1, n), "z": r.uniform(0, 1, n)})
    df["A"] = 5 * df.a + r.normal(0, 0.05, n)
    df["B"] = 3 * df.b + r.normal(0, 0.05, n)
    data = write_df(tmp_path / "h.csv", df)
    v1 = run_tool("ml_train.py", "-f", data, "-x", "a,b", "-y", "A,B")
    run_tool("ml_train.py", "-f", data, "-x", "z", "-y", "A")                       # later, worse, one target
    hist = run_tool("ml_train.py", "-f", data, "--history")
    assert hist["best_version"] == "v1" and hist["best_version_dir"] == v1["version_dir"]   # 2.9.14: v2
    assert hist["best_by_task"]["regression"]["complete"] is True and "best_version_note" not in hist
    rep = run_tool("ml_report.py", "-d", os.path.dirname(v1["version_dir"]), "-o", str(tmp_path / "r.docx"))
    assert rep["version"] == "v1"                                                  # 2.9.14: v2


def _fake_version(root, ver, targets):
    d = root / ver
    d.mkdir(parents=True)
    for name in ["leaderboard.json", "model_dataset.csv"] + [f"model_{t}.json" for t in targets]:
        (d / name).write_text("{}", encoding="utf-8")


def test_task_best_version_split_bests(tmp_path):
    def rec(ver, scores):
        return {"version": ver, "task": "regression", "targets": list(scores),
                "targets_benchmark": {t: {"champion": {"val_score": s}} for t, s in scores.items()},
                "mean_val_score": sum(scores.values()) / len(scores)}
    records = [rec("v1", {"A": 0.9, "B": 0.8}), rec("v2", {"A": 0.95})]
    for r_ in records:
        _fake_version(tmp_path, r_["version"], r_["targets"])
    info = mt.task_best_version(records, str(tmp_path), "regression")
    assert info["version"] == "v1" and info["complete"] is False and info["missing"] == ["A"]
    records.append(rec("v3", {"A": 0.99, "B": 0.99}))
    _fake_version(tmp_path, "v3", ["A", "B"])
    assert mt.task_best_version(records, str(tmp_path), "regression")["complete"] is True


# ---- --best: 타겟마다 이력의 최고 조건으로 다시 학습해 작업 종류마다 버전 하나, ml_report.py -f ----


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
    from ml_predict import sanitize_filename  # noqa: E402
    with open(os.path.join(version_dir, f"model_{sanitize_filename(target)}.json"), encoding="utf-8") as f:
        return json.load(f)


@pytest.fixture(scope="module")
def three_runs(tmp_path_factory):
    work = tmp_path_factory.mktemp("best")
    data = make_multi(str(work / "proc.csv"))
    v1 = run_tool("ml_train.py", "-f", data, "-x", "온도", "-y", "수율", "--desc", "다중 타겟")
    v2 = run_tool("ml_train.py", "-f", data, "-x", "압력,속도", "-y", "수율")      # worse recipe, later run
    v3 = run_tool("ml_train.py", "-f", data, "-x", "압력,속도", "-y", "강도", "--eval", "압속 = 압력 * 속도")
    v4 = run_tool("ml_train.py", "-f", data, "-x", "온도,압력", "-y", "두께")
    assert v1["mean_val_score"] > v2["mean_val_score"]
    best = run_tool("ml_train.py", "-f", data, "--best")
    return work, data, [v1, v2, v3, v4], best


@needs_report
def test_best_builds_one_version_with_every_target(three_runs):
    _, data, runs, best = three_runs
    assert best["version"] == best["best_version"] == "v5" and best["made_best"] is True
    assert best["targets"] == [Y1, Y2, Y3] and best["tag"] == "best"
    assert best["agent_directive"].startswith("VERIFIED:") and best["verified"] is True
    assert len(best["best_versions"]) == 1 and best["best_versions"][0]["version_dir"] == best["version_dir"]
    bt = best["best_targets"]
    assert bt[Y1]["source_version"] == "v1"           # best is not the latest run of that target
    assert bt[Y1]["source_score"] == runs[0]["mean_val_score"]
    assert bt[Y2]["source_version"] == "v3" and bt[Y3]["source_version"] == "v4"
    for t in (Y1, Y2, Y3):
        assert isinstance(bt[t]["new_score"], float) and bt[t]["task"] == "regression"
    assert "ml_report.py" in best["next_commands"][0]["command"] and " -f " in best["next_commands"][0]["command"]
    assert "성능예측표" in best["next_commands"][0]["purpose"]                         # 회귀: 보고서와 성능예측표 세트

    vdir = best["version_dir"]
    assert meta(vdir, Y1)["features_required"] == ["온도(°C)"]
    m2 = meta(vdir, Y2)
    assert m2["features_required"] == ["압력(bar)", "속도(m/s)", "압속"]
    assert m2["derived_evals"] and m2["derived_evals"][0].startswith("압속 = ")
    assert m2["best_source"]["version"] == "v3"
    assert meta(vdir, Y3)["features_required"] == ["온도(°C)", "압력(bar)"]

    rec = ledger(data)[-1]
    assert rec["version"] == "v5" and rec["best"] is True and rec["tag"] == "best"
    assert rec["best_source"][Y1] == {"version": "v1", "score": runs[0]["mean_val_score"]}
    assert rec["target_recipes"][Y2]["features"] == ["압력(bar)", "속도(m/s)"]
    assert rec["data_fingerprint"]["sha256"]
    assert set(rec["features_used"]) == {"온도(°C)", "압력(bar)", "속도(m/s)", "압속"}


@needs_report
def test_best_version_predicts_every_target(three_runs):
    _, _, _, best = three_runs
    pred = run_tool("ml_predict.py", "-m", best["version_dir"], "-i", "온도=5,압력=2,속도=2")
    assert set(pred["targets_predicted"]) == {Y1, Y2, Y3}
    assert pred["predictions_summary"][Y1]["features_used"] == ["온도(°C)"]
    for t in (Y1, Y2, Y3):
        assert len(pred["predictions_summary"][t]["predictions"]) == 1


@needs_report
def test_history_lists_best_record(three_runs):
    _, data, _, _ = three_runs
    hist = run_tool("ml_train.py", "-f", data, "--history")
    assert hist["records"][-1]["best"] is True and hist["latest_version"] == "v5"
    assert "best" not in hist["records"][0]


@needs_report
def test_report_f_reuses_then_rebuilds(three_runs):
    work, data, _, best = three_runs
    rep = run_tool("ml_report.py", "-f", data)
    assert rep["made_best"] is False and rep["best_version"] == "v5" and rep["version"] == "v5"
    assert rep["report_count"] == 1 and len(rep["best_reports"]) == 1
    assert rep["best_version_dir"] == best["version_dir"] and best["version_dir"] in rep["agent_directive"]
    assert set(rep["targets"]) == {Y1, Y2, Y3} and rep["best_targets"][Y1]["source_version"] == "v1"
    d = docx.Document(rep["report_file"])
    text = "\n".join(p.text for p in d.paragraphs)
    heads = [p.text for p in d.paragraphs if p.style.name.startswith("Heading")]
    for i, t in enumerate((Y1, Y2, Y3), 1):
        assert f"4.{i}. {t}" in heads
    assert f"학습 조건 출처: v1 (당시 검증 R² {best['best_targets'][Y1]['source_score']:.4f}), 이번 검증 R²" in text

    again = run_tool("ml_report.py", "-f", data)
    assert again["made_best"] is False and again["best_version"] == "v5"

    # v6: a later training that beats nothing -> v5 still holds every best and is used as is (2.9.14)
    run_tool("ml_train.py", "-f", data, "-x", "온도,속도", "-y", "두께")
    third = run_tool("ml_report.py", "-f", data)
    assert third["made_best"] is False and third["best_version"] == "v5"

    # Changed data file content -> rebuilt even without a new training run
    df = pd.read_csv(data, encoding="utf-8-sig")
    pd.concat([df, df.iloc[:3]]).to_csv(data, index=False, encoding="utf-8-sig")
    fourth = run_tool("ml_report.py", "-f", data)
    assert fourth["made_best"] is True and fourth["best_version"] == "v7"
    assert set(fourth["targets"]) == {Y1, Y2, Y3}
    rec = ledger(data)[-1]
    assert rec["version"] == "v7" and rec["best"] is True


@needs_report
def test_mixed_tasks_two_best_versions(tmp_path):
    data = make_multi(str(tmp_path / "mix.csv"))
    run_tool("ml_train.py", "-f", data, "-x", "온도,압력", "-y", "수율")
    run_tool("ml_train.py", "-f", data, "-x", "온도,속도", "-y", "등급", "--task", "classification")
    rep = run_tool("ml_report.py", "-f", data)
    # each task has one training run that holds its best: used as is, no retraining (2.9.14)
    assert rep["made_best"] is False and rep["report_count"] == 2
    assert [b["version"] for b in rep["best_versions"]] == ["v1", "v2"]
    assert [r["task"] for r in rep["best_reports"]] == ["regression", "classification"]
    assert [b["task"] for b in rep["best_versions"]] == ["regression", "classification"]
    assert rep["best_targets"][LABEL]["task"] == "classification"
    assert len({r["report_file"] for r in rep["best_reports"]}) == 2 and all(os.path.isfile(r["report_file"]) for r in rep["best_reports"])
    assert "2개" in rep["agent_directive"]
    assert rep["version"] == rep["best_reports"][0]["version"] and rep["report_file"] == rep["best_reports"][0]["report_file"]

    only = run_tool("ml_report.py", "-f", data, "--task", "classification")
    assert only["report_count"] == 1 and only["task"] == "classification" and only["made_best"] is False

    tr = run_tool("ml_train.py", "-f", data, "--best", "--task", "regression")
    assert [b["task"] for b in tr["best_versions"]] == ["regression"] and tr["targets"] == [Y1]

    # bests in two different runs -> the report builds one best version first
    run_tool("ml_train.py", "-f", data, "-x", "온도,압력", "-y", "두께")
    built = run_tool("ml_report.py", "-f", data, "--task", "regression")
    assert built["made_best"] is True and set(built["targets"]) == {Y1, Y3}


@needs_report
def test_no_history_stops(tmp_path):
    data = make_multi(str(tmp_path / "empty.csv"))
    for script, args in (("ml_train.py", ("-f", data, "--best")), ("ml_report.py", ("-f", data))):
        err = run_fail(script, *args)
        assert err["agent_directive"].startswith("STOP_RETRY") and "학습 이력이 없습니다" in err["message"]


@needs_report
def test_report_m_path_unchanged(tmp_path):
    data = make_multi(str(tmp_path / "single.csv"))
    tr = run_tool("ml_train.py", "-f", data, "-x", "온도,압력,속도", "-y", "수율,두께")
    rep = run_tool("ml_report.py", "-m", tr["version_dir"])
    assert rep["version"] == rep["best_version"] == tr["version"]
    for key in ("made_best", "best_reports", "best_targets", "best_version_dir"):
        assert key not in rep
    rec = ledger(data)[-1]
    assert "best" not in rec and rec["data_fingerprint"]["size"] > 0


def test_best_task_without_history_is_a_candidate():
    with pytest.raises(mp.InputError) as e:
        mt._tasks_to_build([{"task": "regression"}], "classification", "h.jsonl")
    assert "[추천]" not in str(e.value) and "[후보] --task regression" in str(e.value), str(e.value)


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
    return _ok_pr(*run_raw("ml_train.py", "-f", path, "-x", ",".join(TOP_X), *extra))


def _check_top(top, names=TOP_X):
    assert isinstance(top, list) and 1 <= len(top) <= 5, top
    shares = [e["importance"] for e in top]
    assert shares == sorted(shares, reverse=True), top
    assert all(0.0 <= s <= 1.0 for s in shares) and sum(shares) <= 1.0 + 1e-6, top
    assert all(set(e) == {"feature", "importance"} for e in top), top
    assert all(e["feature"] in names for e in top), top                 # original names, no one-hot '라인_A'
    assert len({e["feature"] for e in top}) == len(top), top


def test_regression_top_features(top_data):
    res = _train_top(top_data, "-y", "수율")
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
    res = _train_top(top_data, "-y", "등급", "--task", "classification")
    top = res["top_features"]["등급"]
    _check_top(top)
    assert top[0]["feature"] == "온도"
    assert "중요 입력: " in res["preview"]


def test_multi_target_top_features(top_data):
    res = _train_top(top_data, "-y", "수율,강도")
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
