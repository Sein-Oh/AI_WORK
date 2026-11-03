# -*- coding: utf-8 -*-
"""
End-to-end CLI contract tests: train -> predict -> report.

Locks the Agent-Computer Interface: each tool must print exactly one JSON line
on stdout, exit 0 with "status": "ok" on success, and exit non-zero with
"status": "error" + an agent_directive on failure. Run before and after every
optimization so behavior changes are caught.
"""
import json
import os
import subprocess
import sys

import pytest

HERE = os.path.dirname(os.path.abspath(__file__))
DEV = os.path.dirname(HERE)  # CLI-dev\ml-kit (tests, tools)
KIT = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(HERE))), "CLI", "ml-kit")  # 실행 파일은 CLI\ml-kit
sys.path.insert(0, HERE)
from make_sample import make_sample  # noqa: E402

FEATURES = "반응온도(°C),소비전력(kW),압축압력(bar),촉매비율(wt%),수분함량(%)"
TARGET = "생산효율(%)"
INLINE = "반응온도(°C)=240,소비전력(kW)=65,압축압력(bar)=3,촉매비율(wt%)=0.8,수분함량(%)=20"


def run_raw(script, *args):
    proc = subprocess.run(
        [sys.executable, os.path.join(KIT, script), *args],
        capture_output=True, text=True, encoding="utf-8", timeout=600,
    )
    lines = [ln for ln in proc.stdout.splitlines() if ln.strip()]
    assert len(lines) == 1, f"{script} must print exactly 1 JSON line, got {len(lines)}:\n{proc.stdout}\n{proc.stderr}"
    return proc, json.loads(lines[0])


def run_tool(script, *args):
    proc, payload = run_raw(script, *args)
    assert proc.returncode == 0, f"{script} failed:\nSTDOUT:\n{proc.stdout}\nSTDERR:\n{proc.stderr}"
    assert payload.get("status") == "ok", payload
    return payload


def run_fail(script, *args):
    proc, payload = run_raw(script, *args)
    assert proc.returncode != 0
    assert payload["status"] == "error"
    assert payload["agent_directive"].startswith(("STOP_RETRY", "ERROR"))
    return payload


@pytest.fixture(scope="module")
def trained(tmp_path_factory):
    work = tmp_path_factory.mktemp("mlkit")
    data = make_sample(str(work / "sample.csv"))
    payload = run_tool(
        "ml_train.py", "-f", data, "-x", FEATURES, "-y", TARGET,
        "--desc", "smoke test", "--tag", "baseline",
        "--eval", "온도전력비 = 반응온도(°C) / (소비전력(kW) + 1e-6)",
    )
    return work, data, payload


# ---------------------------------------------------------------- train

def test_train_creates_version(trained):
    work, _, payload = trained
    data_root = work / "sample_ml_data"
    assert (data_root / "history.jsonl").is_file()
    assert (data_root / "history.json").is_file()
    assert (data_root / "v1").is_dir()
    oof = (data_root / "v1" / "oof_생산효율(%).csv")
    assert oof.is_file() and oof.read_text(encoding="utf-8-sig").splitlines()[0] == "data_row,actual,predicted,residual"
    rec = json.loads((data_root / "history.jsonl").read_text(encoding="utf-8").splitlines()[0])
    assert rec["data_profile"]["rows_loaded"] == 120 and "feature_missing" in rec["data_profile"]
    assert not any(p.name.startswith(".tmp_") for p in data_root.iterdir())
    assert payload["agent_directive"].startswith("VERIFIED")
    assert payload["goal_met"] == (payload["mean_val_score"] >= 0.9 and all(
        b["champion"]["val_score"] >= 0.9 for b in json.loads(
            (data_root / "v1" / "leaderboard.json").read_text(encoding="utf-8"))["targets"].values()))
    meta = json.loads((data_root / "v1" / "model_생산효율(%).json").read_text(encoding="utf-8"))
    assert "sklearn_version" in meta


def test_train_unitless_names_and_eval_alias(trained):
    """README example B style: names without unit suffix are resolved to full column names."""
    work, data, _ = trained
    payload = run_tool(
        "ml_train.py", "-f", data, "-x", "반응온도,소비전력,압축압력,촉매비율,수분함량", "-y", "생산효율",
        "--eval", "에너지집중도 = 소비전력 / (반응온도 + 1e-6); 로그비 = log1p(에너지집중도)",
        "--tag", "alias",
    )
    assert payload["targets"] == [TARGET]
    assert payload["column_aliases"]["생산효율"] == TARGET
    assert payload["description"] == "smoke test"  # inherited from v1
    meta = json.loads((work / "sample_ml_data" / payload["version"] / "model_생산효율(%).json").read_text(encoding="utf-8"))
    # stored formulas use full names so inference replays them exactly
    assert meta["derived_evals"][0] == "에너지집중도 = 소비전력(kW) / (반응온도(°C) + 1e-6)"
    assert "로그비" in meta["features_required"]


def test_train_column_typo_suggests(trained):
    _, data, _ = trained
    # 2.9.33: '반응온도(C)' is the same column (°C / C) and matches; a typo is a [후보] (never an automatic [추천])
    payload = run_fail("ml_train.py", "-f", data, "-x", "반응온도x,소비전력(kW)", "-y", TARGET)
    assert payload["agent_directive"].startswith("STOP_RETRY")
    assert "[후보]" in payload["message"] and "[추천]" not in payload["message"] and "반응온도(°C)" in payload["message"]


def test_train_eval_leakage_blocked(trained):
    _, data, _ = trained
    payload = run_fail("ml_train.py", "-f", data, "-x", FEATURES, "-y", TARGET, "--eval", "누수 = 생산효율 * 2")
    assert "LEAKAGE" in payload["message"]


def test_train_eval_rejects_code(trained):
    _, data, _ = trained
    payload = run_fail("ml_train.py", "-f", data, "-x", FEATURES, "-y", TARGET,
                       "--eval", "x = __import__('os').getcwd()")
    assert payload["agent_directive"].startswith("STOP_RETRY")


def test_train_strings_with_commas_and_placeholders(tmp_path):
    """Thousands separators / '-' placeholders must be parsed (pandas 3 'str' dtype regression)."""
    import pandas as pd
    p = tmp_path / "comma.csv"
    make_sample(str(p))
    df = pd.read_csv(p, encoding="utf-8-sig")
    df["소비전력(kW)"] = df["소비전력(kW)"].map(lambda v: f"{v * 100:,.1f}")
    df["수분함량(%)"] = df["수분함량(%)"].astype(object)
    df.loc[3, "수분함량(%)"] = "-"
    df["상수열"] = 1.0  # constant column -> NaN correlation must not break JSON
    df.to_csv(p, index=False, encoding="cp949")
    payload = run_tool("ml_train.py", "-f", str(p), "-x", FEATURES + ",상수열", "-y", TARGET)
    hist = (tmp_path / "comma_ml_data" / "history.jsonl").read_text(encoding="utf-8")
    assert "NaN" not in hist
    assert payload["mean_val_score"] > 0.5 and "mean_val_r2" not in payload


def test_train_history_mode_is_json(trained):
    _, data, _ = trained
    payload = run_tool("ml_train.py", "-f", data, "--history")
    assert payload["record_count"] >= 1 and payload["records"][0]["version"] == "v1"


def test_bad_cli_args_emit_json():
    proc, payload = run_raw("ml_train.py", "--no-such-option")
    assert proc.returncode == 2 and payload["status"] == "error"


# ---------------------------------------------------------------- predict

def test_predict_inline(trained):
    work, _, _ = trained
    payload = run_tool("ml_predict.py", "-m", str(work / "sample_ml_data" / "v1"), "-i", INLINE)
    assert payload["agent_directive"].startswith("VERIFIED")
    assert payload["drift_warning"] is False
    assert len(payload["predictions_summary"][TARGET]["predictions"]) == 1


def test_predict_inline_json_file_and_unitless_keys(trained):
    work, _, _ = trained
    f = work / "input.json"
    # PowerShell 5 'Out-File' style: UTF-16 with BOM, keys without units
    f.write_text(json.dumps({"반응온도": 240, "소비전력": 65, "압축압력": 3, "촉매비율": 0.8, "수분함량": 20},
                            ensure_ascii=False), encoding="utf-16")
    payload = run_tool("ml_predict.py", "-m", str(work / "sample_ml_data" / "v1"), "-i", str(f))
    assert payload["column_aliases"]["반응온도"] == "반응온도(°C)"


def test_predict_extrapolation_flag(trained):
    work, _, _ = trained
    payload = run_tool("ml_predict.py", "-m", str(work / "sample_ml_data" / "v1"),
                       "-i", INLINE.replace("반응온도(°C)=240", "반응온도(°C)=900"))
    assert payload["drift_warning"] is True
    assert payload["extrapolation_warnings"]


def test_predict_batch(trained):
    work, data, _ = trained
    out = work / "pred.csv"
    payload = run_tool("ml_predict.py", "-m", str(work / "sample_ml_data" / "v1"), "-f", data, "-o", str(out))
    assert out.is_file() and payload["total_rows"] == 120


def test_predict_missing_feature(trained):
    work, _, _ = trained
    payload = run_fail("ml_predict.py", "-m", str(work / "sample_ml_data" / "v1"), "-i", "반응온도(°C)=240")
    assert "누락" in payload["message"]


def test_predict_standalone_deployment(trained, tmp_path):
    """ml_predict.py + version folder alone must be a complete inference package."""
    import shutil
    work, _, _ = trained
    shutil.copy(os.path.join(KIT, "ml_predict.py"), tmp_path / "ml_predict.py")
    shutil.copytree(work / "sample_ml_data" / "v1", tmp_path / "v1")
    proc = subprocess.run(
        [sys.executable, "-I", str(tmp_path / "ml_predict.py"), "-m", str(tmp_path / "v1"), "-i", INLINE],
        capture_output=True, text=True, encoding="utf-8", timeout=300, cwd=str(tmp_path),
    )
    assert proc.returncode == 0, proc.stdout + proc.stderr
    payload = json.loads(proc.stdout.strip().splitlines()[-1])
    assert payload["status"] == "ok" and payload["predictions_summary"][TARGET]["predictions"]


# ---------------------------------------------------------------- report

try:
    import docx  # noqa: F401
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

needs_docx = pytest.mark.skipif(not HAS_DOCX, reason="python-docx not installed in this environment")


@pytest.mark.skipif(HAS_DOCX, reason="only meaningful when python-docx is missing")
def test_report_without_docx_returns_json(trained):
    work, _, _ = trained
    payload = run_fail("ml_report.py", "-d", str(work / "sample_ml_data"))
    assert payload["agent_directive"].startswith("STOP_RETRY") and "python-docx" in payload["message"]


@needs_docx
def test_report(trained):
    work, _, _ = trained
    out = work / "report.docx"
    payload = run_tool("ml_report.py", "-d", str(work / "sample_ml_data"), "-t", TARGET, "-o", str(out))
    assert out.is_file() and payload["best_version"]


@needs_docx
def test_report_version_without_v_and_jsonl_only(trained):
    """history.json missing -> must fall back to history.jsonl (was reading the wrong file)."""
    work, _, _ = trained
    root = work / "sample_ml_data"
    (root / "history.json").rename(root / "history.json.bak")
    try:
        payload = run_tool("ml_report.py", "-d", str(root), "-t", "생산효율", "-v", "1",
                           "-o", str(work / "r_v1.docx"))
        assert payload["best_version"] == "v1"
    finally:
        (root / "history.json.bak").rename(root / "history.json")


@needs_docx
def test_report_unknown_version_fails(trained):
    work, _, _ = trained
    payload = run_fail("ml_report.py", "-d", str(work / "sample_ml_data"), "-v", "v99")
    assert "v99" in payload["message"]
