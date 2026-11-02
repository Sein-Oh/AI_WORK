# -*- coding: utf-8 -*-
"""2.4.0 make_best: every target retrained with its best recipe from the history into ONE version per task,
and ml_report.py -f building one report from it (reused while the history and the data file are unchanged)."""
import json
import os
import sys

import numpy as np
import pandas as pd
import pytest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from test_cli_smoke import run_tool, run_fail  # noqa: E402

docx = pytest.importorskip("docx")
pytest.importorskip("matplotlib")

Y1, Y2, Y3, LABEL = "수율(%)", "강도(MPa)", "두께(mm)", "등급"


def make_multi(path, n=60, seed=3):
    """3 inputs, 3 numeric targets with different drivers, 1 class label."""
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


sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(HERE))), "CLI", "ml-kit"))


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


def test_best_version_predicts_every_target(three_runs):
    _, _, _, best = three_runs
    pred = run_tool("ml_predict.py", "-m", best["version_dir"], "-i", "온도=5,압력=2,속도=2")
    assert set(pred["targets_predicted"]) == {Y1, Y2, Y3}
    assert pred["predictions_summary"][Y1]["features_used"] == ["온도(°C)"]
    for t in (Y1, Y2, Y3):
        assert len(pred["predictions_summary"][t]["predictions"]) == 1


def test_history_lists_best_record(three_runs):
    _, data, _, _ = three_runs
    hist = run_tool("ml_train.py", "-f", data, "--history")
    assert hist["records"][-1]["best"] is True and hist["latest_version"] == "v5"
    assert "best" not in hist["records"][0]


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


def test_no_history_stops(tmp_path):
    data = make_multi(str(tmp_path / "empty.csv"))
    for script, args in (("ml_train.py", ("-f", data, "--best")), ("ml_report.py", ("-f", data))):
        err = run_fail(script, *args)
        assert err["agent_directive"].startswith("STOP_RETRY") and "학습 이력이 없습니다" in err["message"]


def test_report_m_path_unchanged(tmp_path):
    data = make_multi(str(tmp_path / "single.csv"))
    tr = run_tool("ml_train.py", "-f", data, "-x", "온도,압력,속도", "-y", "수율,두께")
    rep = run_tool("ml_report.py", "-m", tr["version_dir"])
    assert rep["version"] == rep["best_version"] == tr["version"]
    for key in ("made_best", "best_reports", "best_targets", "best_version_dir"):
        assert key not in rep
    rec = ledger(data)[-1]
    assert "best" not in rec and rec["data_fingerprint"]["size"] > 0
