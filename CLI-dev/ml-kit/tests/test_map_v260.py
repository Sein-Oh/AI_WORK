# -*- coding: utf-8 -*-
"""
2.6.0 ml_map.py: table interpolation (grid / scattered), extrapolation warnings, nearest measured rows,
reverse lookup (all crossings, extrapolated answers, none found), batch points file, '제외' rows, CLI contract.
Expected grid values are the ones checked by hand with Excel FORECAST formulas.
"""
import os
import sys

import numpy as np
import openpyxl
import pandas as pd
import pytest

HERE = os.path.dirname(os.path.abspath(__file__))
KIT = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(HERE))), "CLI", "ml-kit")  # 실행 파일은 CLI\ml-kit
sys.path.insert(0, HERE)
sys.path.insert(0, KIT)
from test_cli_smoke import run_raw, run_tool, run_fail  # noqa: E402


@pytest.fixture
def grid1(tmp_path):
    a = np.arange(0, 101, 5)
    p = tmp_path / "g1.xlsx"
    pd.DataFrame({"A": a, "B": a * a / 10, "C": 100 - a}).to_excel(p, index=False)
    return str(p)


@pytest.fixture
def grid2(tmp_path):
    """2-key grid: 온도 0..100 step 10 x 압력 1..5, 효율 = 2*온도 + 10*압력 + 온도*압력 (bilinear exact)."""
    rows = [{"온도": t, "압력": p, "효율": 2 * t + 10 * p + t * p} for t in range(0, 101, 10) for p in range(1, 6)]
    path = tmp_path / "g2.csv"
    pd.DataFrame(rows).sample(frac=1, random_state=0).to_csv(path, index=False)   # unsorted on purpose
    return str(path)


@pytest.fixture
def process(tmp_path):
    """Scattered process records in a report-style sheet: title row 1, header row 3, unit row 4, data from row 5."""
    rng = np.random.default_rng(1)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "공정 데이터 보고서"
    ws.append([])
    ws.append(["라인", "반응온도", "소비전력", "생산효율", "제외"])
    ws.append(["", "°C", "kW", "%", ""])
    for i in range(120):
        t, p = rng.uniform(200, 280), rng.uniform(40, 90)
        ws.append(["ABC"[i % 3], t, p, 0.2 * t + 0.1 * p, "센서오류" if i == 7 else None])   # exact plane: 0.2 t + 0.1 p
    path = tmp_path / "공정 데이터.xlsx"
    wb.save(path)
    return str(path)


def P(path, *args):
    return ["-f", path, *args]


# -- forward -----------------------------------------------------------------------------------------------

def test_grid_1d_matches_excel_forecast(grid1):
    r = run_tool("ml_map.py", *P(grid1, "-i", "A=12.5"))
    pt = r["points"][0]
    assert r["mode"] == "grid" and pt["values"] == {"B": 16.25, "C": 87.5} and not pt["extrapolated"]
    assert pt["basis"] == [{"row": 4, "weight": 0.5}, {"row": 5, "weight": 0.5}]            # A=10 is Excel row 4
    assert run_tool("ml_map.py", *P(grid1, "-i", "A=10"))["points"][0]["values"]["B"] == 10.0
    up = run_tool("ml_map.py", *P(grid1, "-i", "A=103"))
    assert up["points"][0]["values"] == {"B": 1058.5, "C": -3.0} and up["points"][0]["extrapolated"]
    assert "3 높음 (데이터 폭의 3%)" in up["points"][0]["out_of_range"]["A"] and up["warnings"]
    assert up["agent_directive"].startswith("VERIFIED:") and "extrapolated=true" in up["agent_directive"]
    down = run_tool("ml_map.py", *P(grid1, "-i", "A=-2", "-y", "B"))
    assert down["points"][0]["values"] == {"B": -1.0} and down["columns"] == ["B"]


def test_grid_2d_bilinear(grid2):
    r = run_tool("ml_map.py", *P(grid2, "-i", "온도=25,압력=2.5", "-y", "효율"))
    pt = r["points"][0]
    assert r["mode"] == "grid" and pt["method"] == "격자 보간"
    assert pt["values"]["효율"] == pytest.approx(2 * 25 + 10 * 2.5 + 25 * 2.5)            # bilinear is exact on t*p
    assert len(pt["basis"]) == 4 and sum(b["weight"] for b in pt["basis"]) == pytest.approx(1.0)
    ext = run_tool("ml_map.py", *P(grid2, "-i", "온도=110,압력=3"))["points"][0]
    assert ext["extrapolated"] and ext["values"]["효율"] == pytest.approx(2 * 110 + 30 + 330)


def test_scattered_two_keys_enclosing_rows_and_far_point(process):
    args = ("--header-row", "3", "--unit-row", "4", "-y", "생산효율")
    r = run_tool("ml_map.py", *P(process, "-i", "반응온도=240,소비전력=60", *args))
    pt = r["points"][0]
    assert r["mode"] == "scattered" and pt["method"].startswith("둘러싼 실측 3점")
    assert pt["values"]["생산효율(%)"] == pytest.approx(0.2 * 240 + 0.1 * 60, abs=1e-6)      # linear data: exact
    assert len(pt["basis"]) == 3 and all(b["row"] >= 5 for b in pt["basis"])
    assert len(r["nearest_rows"]) == 5 and r["nearest_rows"][0]["distance"] <= r["nearest_rows"][-1]["distance"]
    assert "센서오류" not in str(r) and any("제외" in w for w in r["warnings"])            # marked row not used

    far = run_tool("ml_map.py", *P(process, "-i", "반응온도=30,소비전력=200", *args))
    fp = far["points"][0]
    assert fp["extrapolated"] and fp["method"].startswith("전체 추세선 외삽")
    assert abs(fp["values"]["생산효율(%)"]) < 100                                           # never the exploding end-point slope
    assert "낮음 (데이터 폭의 2." in fp["out_of_range"]["반응온도(°C)"] and "높음" in fp["out_of_range"]["소비전력(kW)"]


def test_one_uneven_key_is_scattered_and_category_filter(process):
    r = run_tool("ml_map.py", *P(process, "--header-row", "3", "--unit-row", "4", "-i", "반응온도=240,라인=B", "-y", "생산효율"))
    assert r["mode"] == "scattered" and r["points"][0]["method"] == "위아래 실측 2점 보간"
    assert r["points"][0]["input"]["라인"] == "B" and all(n["라인"] == "B" for n in r["nearest_rows"])
    bad = run_fail("ml_map.py", *P(process, "--header-row", "3", "--unit-row", "4", "-i", "반응온도=240,라인=D"))
    assert "[후보]" in bad["message"]


# -- reverse -----------------------------------------------------------------------------------------------

def test_reverse_grid_exact_and_extrapolated(grid1):
    r = run_tool("ml_map.py", *P(grid1, "--find", "B=90", "--vary", "A"))
    assert r["count"] == 1 and r["solutions"][0]["A"] == pytest.approx(30) and r["agent_directive"].startswith("VERIFIED:")
    ext = run_tool("ml_map.py", *P(grid1, "--find", "B=1100", "--vary", "A"))
    assert ext["solutions"][0]["A"] == pytest.approx(105.128205, abs=1e-5) and ext["solutions"][0]["extrapolated"]
    dec = run_tool("ml_map.py", *P(grid1, "--find", "C=42.5", "--vary", "A"))                 # decreasing column
    assert dec["solutions"][0]["A"] == pytest.approx(57.5)


def test_reverse_two_answers_and_none(tmp_path):
    a = np.arange(0, 101, 5)
    p = tmp_path / "hill.csv"
    pd.DataFrame({"A": a, "Y": 100 - (a - 50) ** 2 / 25}).to_csv(p, index=False)         # rises then falls, max 100
    r = run_tool("ml_map.py", "-f", str(p), "--find", "Y=64", "--vary", "A")
    assert [round(s["A"], 6) for s in r["solutions"] if not s["extrapolated"]] == [20.0, 80.0]
    none = run_tool("ml_map.py", "-f", str(p), "--find", "Y=150", "--vary", "A")
    assert none["count"] == 0 and none["reachable"]["max"] == pytest.approx(100) and "찾은 값 없음" in none["preview"]


def test_reverse_scattered_with_fixed_key(process):
    r = run_tool("ml_map.py", *P(process, "--header-row", "3", "--unit-row", "4",
                                 "--find", "생산효율=54", "--vary", "반응온도", "-i", "소비전력=60"))
    assert r["mode"] == "scattered" and r["count"] >= 1
    assert r["solutions"][0]["반응온도(°C)"] == pytest.approx((54 - 6) / 0.2, abs=0.05)       # 240 on the exact plane
    assert r["fixed"] == {"소비전력(kW)": 60.0} and len(r["nearest_rows"]) == 5


# -- batch, errors, contract ----------------------------------------------------------------------------------

def test_points_file_saves_workbook(grid1, tmp_path):
    q = tmp_path / "조회.xlsx"
    pd.DataFrame({"메모": ["a", "b", "c", "d"], "A": [12.5, 103, None, 10]}).to_excel(q, index=False)   # 메모: not a data column
    r = run_tool("ml_map.py", *P(grid1, "--points", str(q)))
    assert r["count"] == 4 and r["extrapolated_count"] == 1 and r["verified"]
    book = pd.read_excel(r["output_file"], sheet_name=None)
    res = book["결과"]
    assert list(res["보간_B"].round(4)[[0, 1, 3]]) == [16.25, 1058.5, 10.0]
    assert res["방식"][2].startswith("계산 불가") and res["외삽"][1] == "예"
    assert len(book["가까운 실측"]) == 15 and "원본 행" in book["가까운 실측"].columns
    again = run_tool("ml_map.py", *P(grid1, "--points", str(q), "--save", r["output_file"]))
    assert again["output_file"] != r["output_file"] and again.get("recovered")             # never overwrites


def test_errors_and_contract(grid1):
    assert run_tool("ml_map.py", *P(grid1, "-i", "a=10"))["keys"] == ["A"]                     # case only: matched
    miss = run_fail("ml_map.py", *P(grid1, "-i", "Ax=10"))
    assert "[후보] 'A'" in miss["message"] and "[추천]" not in miss["message"]    # 2.9.31: short ASCII names of other length
    assert run_fail("ml_map.py", *P(grid1, "-i", "A=abc"))["error_type"] == "input"
    assert run_fail("ml_map.py", *P(grid1, "--find", "B=90", "--vary", "A", "-i", "A=3"))["error_type"] == "input"
    proc, payload = run_raw("ml_map.py", *P(grid1, "--find", "B=90"))                        # --vary missing
    assert proc.returncode == 2 and payload["status"] == "error"
    proc, payload = run_raw("ml_map.py", "-h")
    assert proc.returncode == 0 and payload["command"] == "help" and "usage:" in proc.stderr
