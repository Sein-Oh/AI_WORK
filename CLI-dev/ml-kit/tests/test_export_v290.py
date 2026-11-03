# -*- coding: utf-8 -*-
"""
2.9.0 ml_export.py: input selection by retraining (interacting inputs kept, --keep respected), Excel workbook with
formula-only interpolation, protection (only yellow input cells editable, no password, nothing hidden).
The workbook formulas are recalculated with LibreOffice (when installed) and compared with the Python result.
"""
import os
import shutil
import subprocess
import sys

import numpy as np
import openpyxl
import pandas as pd
import pytest
from scipy.interpolate import RegularGridInterpolator

HERE = os.path.dirname(os.path.abspath(__file__))
KIT = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(HERE))), "CLI", "ml-kit")  # 실행 파일은 CLI\ml-kit
sys.path.insert(0, HERE)
sys.path.insert(0, KIT)
from test_cli_smoke import run_raw, run_tool, run_fail  # noqa: E402

SOFFICE = shutil.which("soffice") or shutil.which("libreoffice")


@pytest.fixture(scope="module")
def version(tmp_path_factory):
    """유량 x 농도 interact (유량 alone looks unimportant), 잡음 is unrelated."""
    d = tmp_path_factory.mktemp("exp")
    r = np.random.default_rng(21)
    n = 500
    df = pd.DataFrame({k: r.uniform(0, 1, n) for k in ["온도", "유량", "농도", "잡음"]})
    df["효율"] = 10 * np.sin(3 * df.온도) + 6 * df.유량 * df.농도 + r.normal(0, 0.2, n)
    df["등급"] = np.where(df.효율 > df.효율.median(), "상", "하")
    path = d / "s.csv"
    df.to_csv(path, index=False)
    res = run_tool("ml_train.py", "-f", str(path), "-x", "온도,유량,농도,잡음", "-y", "효율", "--desc", "시험용 공정 효율 데이터. 두 번째 문장.")
    return d, path, res["version_dir"]


def _grid(wb, sheet, k):
    ws = wb[sheet]
    A = np.array([r for r in ws.iter_rows(min_row=2, max_col=k + 1, values_only=True)], dtype=float)
    axes = [np.unique(A[:, j]) for j in range(k)]
    return RegularGridInterpolator(axes, A[:, k].reshape([len(a) for a in axes]))


def _recalc(path, tmp):
    out = tmp / "calc"
    out.mkdir(exist_ok=True)
    shutil.copy(path, out / "t.xlsx")
    subprocess.run([SOFFICE, "--headless", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", str(out / "o"), str(out / "t.xlsx")],
                   capture_output=True, timeout=600)
    return openpyxl.load_workbook(out / "o" / "t.xlsx", data_only=True)


def test_search_keeps_interacting_input_and_builds_locked_workbook(version, tmp_path):
    _, _, vdir = version
    r = run_tool("ml_export.py", "-m", vdir)
    t = r["targets"]["효율"]
    assert t["inputs"] == ["온도", "유량", "농도"] and t["dropped"] == ["잡음"]          # 유량 kept despite low importance
    assert [c["count"] for c in t["candidates"]] == [4, 3, 2, 1] and "0.005" in t["reason"]
    assert os.path.isfile(os.path.join(r["export_model_dir"], "model_효율.joblib"))
    assert r["title"] == "시험용 공정 효율 데이터." and r["next_commands"][0]["command"].count("--inputs") == 1
    wb = openpyxl.load_workbook(r["output_file"])
    assert wb.sheetnames == ["계산기", "일괄계산", "사용법", "표_효율", "정확도", "정보"]
    assert all(wb[n].protection.sheet for n in wb.sheetnames) and wb.security.lockStructure
    assert wb["표_효율"].sheet_state == "veryHidden"                                    # grid not shown (2.9.3)
    assert all(wb[n].sheet_state == "visible" for n in wb.sheetnames if not n.startswith("표_"))
    calc = wb["계산기"]
    unlocked = [c.coordinate for row in calc.iter_rows() for c in row if not c.protection.locked]
    assert unlocked == ["C6", "C7", "C8"] and calc["B6"].value == "온도"
    batch = wb["일괄계산"]
    assert not batch["A3"].protection.locked and batch["E3"].protection.locked
    assert wb["표_효율"].max_row - 1 == t["table_rows"]
    a = t["agreement"]
    assert a["points"] == 2000 and a["agreement_pct"] > 90 and a["max_diff_pct"] >= a["p95_diff_pct"]
    assert any("일치도" in str(c.value) for row in wb["정확도"].iter_rows() for c in row)
    assert any(str(c.value).startswith("모델 예측과 평균") for row in wb["정보"].iter_rows() for c in row if c.value)
    assert any("엑셀" in u for u in r.get("unverified", [])) or r["verified"]          # Excel check only on Windows


def test_status_cells_have_visible_colors(version):
    # conditional formats: Excel reads bgColor for a solid fill, so fgColor alone shows no color
    _, _, vdir = version
    wb = openpyxl.load_workbook(run_tool("ml_export.py", "-m", vdir)["output_file"])
    for sheet in ("계산기", "일괄계산"):
        rules = [(rule.formula[0], rule.dxf) for cf in wb[sheet].conditional_formatting for rule in cf.rules]
        ok = [d for f, d in rules if f.endswith('="정상"')]
        warn = [d for f, d in rules if "범위 밖" in f and "참고용" in f]
        assert ok and warn and len(ok) == len(warn)
        assert all(d.fill.bgColor.rgb.endswith("DCFCE7") and d.font.color.rgb.endswith("166534") for d in ok)
        assert all(d.fill.bgColor.rgb.endswith("FCE4D6") and d.font.color.rgb.endswith("C2410C") for d in warn)


def test_keep_and_inputs_and_errors(version):
    _, _, vdir = version
    k = run_tool("ml_export.py", "-m", vdir, "--keep", "잡음", "--title", "짧은 제목")
    assert "잡음" in k["targets"]["효율"]["inputs"] and k["title"] == "짧은 제목"
    i = run_tool("ml_export.py", "-m", vdir, "--inputs", "온도,유량")
    assert i["targets"]["효율"]["inputs"] == ["온도", "유량"] and len(i["targets"]["효율"]["candidates"]) == 1
    assert "[후보] '온도'" in run_fail("ml_export.py", "-m", vdir, "--keep", "온도x")["message"]     # 2.9.33: was [추천]
    assert "--keep" in run_fail("ml_export.py", "-m", vdir, "--keep", "잡음", "--inputs", "온도")["message"]
    proc, payload = run_raw("ml_export.py", "-h")
    assert proc.returncode == 0 and payload["command"] == "help"


def test_classification_only_version_is_stopped(version):
    _, path, _ = version
    cls = run_tool("ml_train.py", "-f", str(path), "-x", "온도,유량", "-y", "등급", "--task", "classification")
    assert "숫자 타겟" in run_fail("ml_export.py", "-m", cls["version_dir"])["message"]


@pytest.mark.skipif(SOFFICE is None, reason="LibreOffice not installed")
def test_formulas_recalculate_to_the_python_values(version, tmp_path):
    _, _, vdir = version
    r = run_tool("ml_export.py", "-m", vdir, "--inputs", "온도,유량,농도", "--save", str(tmp_path / "f.xlsx"))
    src = openpyxl.load_workbook(r["output_file"])
    interp = _grid(src, "표_효율", 3)
    # calculator at default inputs (medians) + 3 batch rows, one of them outside the training range
    src["일괄계산"]["A3"], src["일괄계산"]["B3"], src["일괄계산"]["C3"] = 0.21, 0.73, 0.4
    src["일괄계산"]["A4"], src["일괄계산"]["B4"], src["일괄계산"]["C4"] = 0.9, 0.1, 0.95
    src["일괄계산"]["A5"], src["일괄계산"]["B5"], src["일괄계산"]["C5"] = 1.05, 0.5, 0.5          # out of range
    mod = tmp_path / "m.xlsx"
    src.save(mod)
    out = _recalc(mod, tmp_path)
    calc = out["계산기"]
    x = [calc[f"C{r}"].value for r in (6, 7, 8)]
    assert calc["C11"].value == pytest.approx(float(interp([x])[0]), abs=1e-9)                # result row = 5 + inputs + 3
    assert calc["F11"].value == "정상" and calc["F6"].value == "정상"
    b = out["일괄계산"]
    for row, xs in ((3, (0.21, 0.73, 0.4)), (4, (0.9, 0.1, 0.95))):
        assert b[f"D{row}"].value == pytest.approx(float(interp([xs])[0]), abs=1e-9)
        assert b[f"E{row}"].value < b[f"D{row}"].value < b[f"F{row}"].value
    assert b["G5"].value == "참고용 (학습 범위 밖)" and b["D6"].value in (None, "")


@pytest.mark.skipif(SOFFICE is None, reason="LibreOffice not installed")
def test_derived_features_stay_inside_the_model(tmp_path):
    r = np.random.default_rng(4)
    n = 300
    df = pd.DataFrame({"a": r.uniform(1, 2, n), "b": r.uniform(1, 2, n)})
    df["y"] = 3 * df.a * df.b + r.normal(0, 0.05, n)
    p = tmp_path / "e.csv"
    df.to_csv(p, index=False)
    v = run_tool("ml_train.py", "-f", str(p), "-x", "a,b", "-y", "y", "--eval", "ab = a * b")["version_dir"]
    res = run_tool("ml_export.py", "-m", v, "--inputs", "a,b", "--save", str(tmp_path / "e.xlsx"))
    t = res["targets"]["y"]
    assert t["inputs"] == ["a", "b"]                                                     # raw inputs only on the sheet
    wb = openpyxl.load_workbook(res["output_file"])
    assert [wb["계산기"][f"B{r}"].value for r in (6, 7)] == ["a", "b"]
    out = _recalc(res["output_file"], tmp_path)
    x = [out["계산기"]["C6"].value, out["계산기"]["C7"].value]
    assert out["계산기"]["C10"].value == pytest.approx(3 * x[0] * x[1], abs=0.15)


def test_several_inputs_and_outputs(tmp_path):
    """Each output keeps its own inputs (training order kept); the calculator asks for the union of inputs and
    an output is computed as soon as ITS inputs are filled."""
    r = np.random.default_rng(31)
    n = 300
    df = pd.DataFrame({"온도": r.uniform(0, 1, n), "압력": r.uniform(0, 1, n), "유량": r.uniform(0, 1, n), "분진": r.uniform(0, 1, n)})
    df["효율"] = 10 * np.sin(3 * df.온도) + r.normal(0, 0.1, n)
    df["전력"] = 5 * df.압력 * df.유량 + 2 * df.유량 + r.normal(0, 0.05, n)
    p = tmp_path / "m.csv"
    df.to_csv(p, index=False)
    v = run_tool("ml_train.py", "-f", str(p), "-x", "온도,압력,유량,분진", "-y", "효율,전력")["version_dir"]
    res = run_tool("ml_export.py", "-m", v, "--save", str(tmp_path / "m.xlsx"))
    assert list(res["targets"]) == ["효율", "전력"]                                       # -y order
    assert res["targets"]["효율"]["inputs"] == ["온도"] and res["targets"]["전력"]["inputs"] == ["압력", "유량"]
    wb = openpyxl.load_workbook(res["output_file"])
    calc = wb["계산기"]
    assert [calc[f"B{r}"].value for r in (6, 7, 8)] == ["온도", "압력", "유량"]          # union, 분진 dropped everywhere
    assert [calc[f"B{r}"].value for r in (11, 12)] == ["효율", "전력"]
    assert "COUNT($C$6)=1" in calc["C11"].value and "COUNT($C$7,$C$8)=2" in calc["C12"].value
    assert wb.sheetnames[3:5] == ["표_효율", "표_전력"]


def test_workbook_has_no_full_paths(version):
    """The folder path before <data>_ml_data (user name, company folders) is personal information."""
    d, _, vdir = version
    res = run_tool("ml_export.py", "-m", vdir, "--inputs", "온도,유량,농도")
    wb = openpyxl.load_workbook(res["output_file"])
    parent = os.path.dirname(os.path.dirname(vdir))
    texts = [str(c.value) for ws in wb.worksheets for row in ws.iter_rows() for c in row if isinstance(c.value, str)]
    assert not any(parent in t for t in texts)
    assert any(t == os.path.join("s_ml_data", "v1") for t in texts)


def test_agreement_basis_near_zero():
    """An output passing near zero is compared relative to its range (a % of a tiny value means nothing)."""
    import ml_export as ex
    model = np.linspace(-1, 1, 101)
    a = ex.agreement(model + 0.02, model, {"residual_lower": -0.1, "residual_upper": 0.1})
    assert a["basis"] == "출력 범위 대비" and a["agreement_pct"] == pytest.approx(99.0)
    assert a["within_interval_pct"] == 100.0
    b = ex.agreement(np.full(10, 101.0), np.full(10, 100.0) + np.arange(10) * 0.0, None)
    assert b["basis"] == "예측값 대비" and b["agreement_pct"] == pytest.approx(99.0) and b["within_interval_pct"] is None
