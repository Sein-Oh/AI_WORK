# -*- coding: utf-8 -*-
"""
2.9.15 (4차 점검 수정): tiny / duplicated data never ends in an internal error (inner CV of LogisticRegressionCV,
KNN neighbours, ElasticNetCV, an all-empty column in a training fold), --unit-row in the --header-row suggestion and a
numeric unit row refused, a unit row read as the first data row in prediction, model names with units found in a
header + unit row pair, whole-number header cells, --history / ml_report -d best version, report number format,
ml_map unused-column warning on the filtered rows, report UNVERIFIED when the calculator was not verified.
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


def _csv(path, df):
    df.to_csv(path, index=False, encoding="utf-8-sig")
    return str(path)


def _schema(cols, categorical=None):
    cat = categorical or {}
    return {"numeric": [c for c in cols if c not in cat], "categorical": cat, "datetime": {}, "input_columns": list(cols)}


def _rerun(message):
    """Runs the full-command [추천] of a message (python <script> args...)."""
    rec = message.split("[추천] ")[1].split("\n")[0].strip()
    assert rec.startswith("python"), rec
    parts = shlex.split(rec)
    return os.path.basename(parts[1]), parts[2:]


# 1. tiny / duplicated data: no internal error ----------------------------------------------------------------

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
    res = run_tool("ml_train.py", "-f", _csv(tmp_path / "c5x2.csv", df), "-x", "a,b", "-y", "c", "--task", "classification")
    assert res["agent_directive"].startswith("VERIFIED")


# 2. / 4. header + unit row suggestions ------------------------------------------------------------------------

@pytest.fixture(scope="module")
def unit_book(tmp_path_factory):
    work = tmp_path_factory.mktemp("units")
    r = np.random.default_rng(5)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["반응 기록"])
    ws.append([None])
    ws.append(["반응온도", "압력", "효율"])
    ws.append(["°C", "bar", "%"])
    for _ in range(40):
        t, p = float(r.uniform(150, 200)), float(r.uniform(1, 5))
        ws.append([round(t, 2), round(p, 3), round(0.3 * t + 2 * p + float(r.normal(0, 0.2)), 3)])
    path = work / "u.xlsx"
    wb.save(path)
    v = run_tool("ml_train.py", "-f", str(path), "-x", "반응온도,압력", "-y", "효율",
                 "--header-row", "3", "--unit-row", "4")["version_dir"]
    return {"work": work, "path": str(path), "v": v}


def test_numeric_unit_row_refused(unit_book):
    bad = run_fail("ml_train.py", "-f", unit_book["path"], "-x", "반응온도,압력", "-y", "효율",
                   "--header-row", "4", "--unit-row", "5")
    assert "단위 행(--unit-row 5)" in bad["message"] and bad["agent_directive"].startswith("STOP_RETRY")
    script, args = _rerun(bad["message"])
    assert script == "ml_train.py" and args[-4:] == ["--header-row", "3", "--unit-row", "4"]
    res = run_tool(script, *args)
    meta = json.load(open(os.path.join(res["version_dir"], "model_효율(%).json"), encoding="utf-8"))
    assert meta["features_required"] == ["반응온도(°C)", "압력(bar)"]              # 2.9.14: '반응온도(192.58)'


def test_header_hint_carries_the_unit_row(unit_book):
    # --header-row 2 --unit-row 3 (both one row too high): the names are in row 3, the units in row 4
    bad = run_fail("ml_train.py", "-f", unit_book["path"], "-x", "반응온도,압력", "-y", "효율",
                   "--header-row", "2", "--unit-row", "3")
    _, args = _rerun(bad["message"])
    assert args[-4:] == ["--header-row", "3", "--unit-row", "4"]                  # 2.9.14: [추천] --header-row 3
    assert "--data-row" not in args


def test_title_row_hint_carries_the_unit_row(tmp_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["새 조건", "작성 2026-11-02"])
    ws.append(["온도", "압력", "효율"])
    ws.append(["°C", "bar", "%"])
    r = np.random.default_rng(6)
    for _ in range(30):
        t, p = float(r.uniform(150, 200)), float(r.uniform(1, 5))
        ws.append([t, p, 0.3 * t + 2 * p])
    p = tmp_path / "t.xlsx"
    wb.save(p)
    bad = run_fail("ml_train.py", "-f", str(p), "-x", "온도,압력", "-y", "효율")
    _, args = _rerun(bad["message"])
    assert args[-4:] == ["--header-row", "2", "--unit-row", "3"]                  # 2.9.14: --header-row 2 only


def test_predict_finds_model_names_with_units_in_a_header_unit_pair(unit_book, tmp_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["새 조건", "작성 2026-11-02"])
    ws.append(["반응온도", "압력"])
    ws.append(["°C", "bar"])
    ws.append([180, 3])
    ws.append([185, 3.2])
    p = tmp_path / "t2c.xlsx"
    wb.save(p)
    bad = run_fail("ml_predict.py", "-m", unit_book["v"], "-f", str(p))
    script, args = _rerun(bad["message"])                                          # 2.9.14: no layout hint
    assert args[-4:] == ["--header-row", "2", "--unit-row", "3"]
    ok = run_tool(script, *args)
    assert ok["sample_count"] == 2 and not ok["missing_input_warning"]


def test_typo_hints_skip_layout_artefacts():
    hints = mp._typo_hints(["압력(bar)"], ["bar(3)", "°C(180)"], ["압력(bar)", "반응온도(°C)"], False)
    assert hints == []                                                             # 2.9.14: 'bar(3)' -> '압력(bar)'
    assert mp._typo_hints(["압력(bar)"], ["압력(bra)"], ["압력(bar)"], False)


# 3. unit row read as the first data row in prediction -------------------------------------------------------

def test_predict_plain_template_on_a_unit_row_file(unit_book, tmp_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["반응온도", "압력"])
    ws.append(["°C", "bar"])
    ws.append([180, 3])
    ws.append([185, 3.2])
    p = tmp_path / "u2.xlsx"
    wb.save(p)
    bad = run_fail("ml_predict.py", "-m", unit_book["v"], "-f", str(p))           # 2.9.14: VERIFIED, unit row imputed
    script, args = _rerun(bad["message"])
    assert args[-2:] == ["--unit-row", "2"]
    ok = run_tool(script, *args)
    assert ok["sample_count"] == 2 and not ok["missing_input_warning"]


def test_placeholder_first_row_is_not_a_unit_row(unit_book, tmp_path):
    df = pd.DataFrame({"반응온도": ["-", 180, 185], "압력": ["-", 3, 3.2]})
    ok = run_tool("ml_predict.py", "-m", unit_book["v"], "-f", _csv(tmp_path / "ph.csv", df))
    assert ok["sample_count"] == 3


# 5. whole-number header cells -------------------------------------------------------------------------------

def test_year_headers_are_integers(tmp_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["구분", 2023, 2024])
    for i in range(5):
        ws.append([f"L{i}", i * 1.5, i * 2.5])
    p = tmp_path / "y.xlsx"
    wb.save(p)
    assert list(mp.load_data_file(str(p)).columns) == ["구분", "2023", "2024"]   # 2.9.14: '2023.0'


# 6. --history best_version and ml_report -d -----------------------------------------------------------------

def test_history_best_version_covers_every_target(tmp_path):
    r = np.random.default_rng(7)
    n = 80
    df = pd.DataFrame({"a": r.uniform(0, 1, n), "b": r.uniform(0, 1, n), "z": r.uniform(0, 1, n)})
    df["A"] = 5 * df.a + r.normal(0, 0.05, n)
    df["B"] = 3 * df.b + r.normal(0, 0.05, n)
    data = _csv(tmp_path / "h.csv", df)
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


# 7. report numbers ------------------------------------------------------------------------------------------

def test_report_fmt_significant_digits():
    assert mr.fmt(0.00012346) == "0.0001235"                                       # 2.9.14: 0.0001
    assert mr.fmt(0.0066) == "0.0066" and mr.fmt(0.00099) == "0.00099" and mr.fmt(0.001) == "0.0010"
    assert mr.fmt(0.12345678) == "0.1235" and mr.fmt(1234.56789) == "1,234.5679"
    assert mr.fmt(0.00012346, 3) == "0.000123" and mr.fmt(2.5e-7) == "2.500e-07"
    # no jump around 1e-3: neighbours keep the same number of significant digits
    assert mr.fmt(0.0009876) == "0.0009876" and mr.fmt(0.001234) == "0.001234"


# 8. ml_map unused-column warning ----------------------------------------------------------------------------

def test_map_unused_columns_use_the_filtered_rows(tmp_path):
    df = pd.DataFrame({"라인": ["X"] * 6 + ["Y"] * 6, "A": [10, 10, 20, 20, 30, 30] * 2, "B": [1, 2] * 6,
                       "Y1": [20, 30, 30, 40, 40, 50, 120, 130, 130, 140, 140, 150],
                       "Y3": [None] * 6 + [9, 8, 19, 18, 29, 28]})
    res = run_tool("ml_map.py", "-f", _csv(tmp_path / "g.csv", df), "-i", "라인=X,A=15", "-y", "Y1")
    unused = [w for w in res.get("warnings", []) if "기준으로 주지 않은 열" in w]
    assert unused and "'B'" in unused[0] and "Y3" not in unused[0]                # 2.9.14: Y3 (empty for X) listed


# 9. report + unverified calculator -> UNVERIFIED ------------------------------------------------------------

def test_report_unverified_when_calculator_unverified():
    result = {"preview": "p", "verified": True, "agent_directive": "VERIFIED: 보고서 저장."}
    ex = {"result": {"output_file": "x.xlsx", "targets": {}, "verified": False,
                     "unverified": ["엑셀이 없는 환경이라 엑셀에서 수식 계산은 확인하지 못함"]}}
    mr.attach_export(result, ex)
    assert result["verified"] is False and result["export_verified"] is False      # 2.9.14: VERIFIED
    assert result["agent_directive"].startswith("UNVERIFIED:")
    assert result["unverified"] == ["성능예측표: 엑셀이 없는 환경이라 엑셀에서 수식 계산은 확인하지 못함"]
    ok = {"preview": "p", "verified": True, "agent_directive": "VERIFIED: r."}
    mr.attach_export(ok, {"result": {"output_file": "x.xlsx", "targets": {}, "verified": True}})
    assert ok["verified"] is True and ok["agent_directive"].startswith("VERIFIED:")
