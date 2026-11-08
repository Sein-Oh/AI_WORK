# -*- coding: utf-8 -*-
"""표 보간(ml_map.py): 격자·흩어진 점 보간, 외삽 경고, 가까운 실측 행, 역산(모든 교차점), 조회 파일(--points),
값 열마다 다른 행(MultiTable), '제외' 행. 격자 기대값은 Excel FORECAST 수식으로 손 계산한 값이다.

--points 열 이름 검사는 test_names.py.
"""
import itertools
import os
import sys

import numpy as np
import openpyxl
import pandas as pd
import pytest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from mlkit_helpers import split_command, recommendation, recommended_command, run_fail, run_raw, run_tool, write_book, write_csv, write_df  # noqa: E402
import _ml_runtime as rt  # noqa: E402


# ---- 정방향 보간 ----------------------------------------------------------------------------------------------------


@pytest.fixture
def grid1(tmp_path):
    a = np.arange(0, 101, 5)
    p = tmp_path / "g1.xlsx"
    pd.DataFrame({"A": a, "B": a * a / 10, "C": 100 - a}).to_excel(p, index=False)
    return str(p)


@pytest.fixture
def grid2(tmp_path):
    """키 2개 격자: 온도 0..100(10 간격) x 압력 1..5, 효율 = 2*온도 + 10*압력 + 온도*압력 (쌍선형 보간이 정확)."""
    rows = [{"온도": t, "압력": p, "효율": 2 * t + 10 * p + t * p} for t in range(0, 101, 10) for p in range(1, 6)]
    path = tmp_path / "g2.csv"
    pd.DataFrame(rows).sample(frac=1, random_state=0).to_csv(path, index=False)   # unsorted on purpose
    return str(path)


@pytest.fixture
def process(tmp_path):
    """보고서 모양 시트의 흩어진 공정 기록: 1행 제목, 3행 머리글, 4행 단위, 5행부터 데이터."""
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


def test_code_column_is_matched_exactly_not_interpolated(tmp_path):
    """2.9.61: 코드 '007'과 '7'을 같은 숫자 7로 합쳐 보간했다(VERIFIED). 앞자리 0 코드가 있는 열은 범주처럼 같은 코드의
    행만 쓴다. 학습에 없는 코드('07')는 [후보]."""
    r = np.random.default_rng(4)
    t = r.uniform(150, 200, 90).round(2)
    code = np.array(["007", "7", "12"])[np.arange(90) % 3]
    y = 2 * t + np.select([code == "007", code == "7"], [100.0, 0.0], 50.0)
    p = tmp_path / "c.csv"
    pd.DataFrame({"온도": t, "코드": code, "수율": y.round(2)}).to_csv(p, index=False)
    for c, off in (("7", 0.0), ("007", 100.0)):
        res = run_tool("ml_map.py", "-f", str(p), "-i", f"온도=170,코드={c}", "-y", "수율")
        assert res["keys"] == ["온도"] and res["points"][0]["input"]["코드"] == c, res
        assert res["points"][0]["values"]["수율"] == pytest.approx(340 + off, abs=1e-6), res
        assert all(str(n["코드"]) in (c, c + ".0") for n in res["nearest_rows"]), res["nearest_rows"]
    bad = run_fail("ml_map.py", "-f", str(p), "-i", "온도=170,코드=07", "-y", "수율")
    assert "[후보] 1) 007 2) 12 3) 7" in bad["message"] and "[추천]" not in bad["message"], bad


def test_two_digit_numbers_with_a_leading_zero_are_interpolated(tmp_path):
    """3차 검토(2.9.63): '05, 10, 15, 20'의 '05' 하나로 숫자 열을 코드 열로 보아 보간 대신 틀린 [후보]를 냈다."""
    p = tmp_path / "rt.csv"
    p.write_text("온도,압력,수율\n05,1,50\n10,1,60\n15,1,70\n20,1,80\n05,2,55\n10,2,65\n15,2,75\n20,2,85\n",
                 encoding="utf-8")
    res = run_tool("ml_map.py", "-f", str(p), "-i", "온도=12,압력=1", "-y", "수율")
    assert res["keys"] == ["온도", "압력"] and res["points"][0]["values"]["수율"] == pytest.approx(64.0), res


# ---- 역산 --find --vary --------------------------------------------------------------------------------


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


# ---- 조회 파일 --points, 오류와 계약 --------------------------------------------------------------------


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
    assert "[후보] 1) 'A'" in miss["message"] and "[추천]" not in miss["message"]    # 2.9.31: short ASCII names of other length
    assert run_fail("ml_map.py", *P(grid1, "-i", "A=abc"))["error_type"] == "input"
    assert run_fail("ml_map.py", *P(grid1, "--find", "B=90", "--vary", "A", "-i", "A=3"))["error_type"] == "input"
    proc, payload = run_raw("ml_map.py", *P(grid1, "--find", "B=90"))                        # --vary missing
    assert proc.returncode == 2 and payload["status"] == "error"
    proc, payload = run_raw("ml_map.py", "-h")
    assert proc.returncode == 0 and payload["command"] == "help" and "usage:" in proc.stderr


@pytest.fixture(scope="module")
def grid_gap(tmp_path_factory):
    work = tmp_path_factory.mktemp("gridgap")
    rows = [(a, b) for a in (10, 20, 30, 40) for b in (1, 2, 3)]
    df = pd.DataFrame(rows, columns=["A", "B"]).astype(float)
    df["Y1"] = df.A * df.B
    df["Y2"] = df.A + df.B
    df.loc[(df.A == 20) & (df.B == 1), "Y2"] = np.nan
    data = write_df(work / "map1.csv", df)
    pts = write_df(work / "pts.csv", pd.DataFrame({"A": [15, 30, 35], "B": [1.5, 2.0, 2.5]}))
    return {"data": data, "pts": pts, "work": work}


def test_points_basis_layout_once_and_numbers(grid_gap):
    res = run_tool("ml_map.py", "-f", grid_gap["data"], "--points", grid_gap["pts"], "--save", "p.xlsx")
    sheet = pd.read_excel(res["output_file"], sheet_name="결과")
    assert "근거" not in sheet.columns and "근거_Y1" in sheet.columns            # 2.9.13: 근거 and 근거_<열> mixed
    assert sheet["근거_Y1"].notna().all()
    assert pd.api.types.is_numeric_dtype(sheet["A"]) and pd.api.types.is_numeric_dtype(sheet["B"])   # 2.9.13: text


def test_map_points_header_hint(tmp_path):
    rows = [["맵 기록", "2026-11-02"], ["A", "B"]] + [[float(i), float(i * i)] for i in range(10)]
    data = write_book(tmp_path / "m.xlsx", rows)
    pts = tmp_path / "pts.csv"
    pd.DataFrame({"A": [2.5, 3.5]}).to_csv(pts, index=False, encoding="utf-8-sig")
    bad = run_fail("ml_map.py", "-f", data, "--points", str(pts), "-y", "B")
    _, rec = recommended_command(bad["message"])                  # 2.9.15: no hint, 2.9.71: 전체 명령
    assert rec[rec.index("--header-row") + 1] == "2", rec
    ok = run_tool("ml_map.py", "-f", data, "--points", str(pts), "-y", "B", "--header-row", "2")
    assert ok["count"] == 2


def test_map_points_keep_data_decimal_comma(tmp_path):
    data = [["온도", "수율"]] + [[f"{170 + i},5", f"{80 + i},25"] for i in range(10)]
    path = write_csv(tmp_path / "grid.csv", data)
    pts = write_csv(tmp_path / "pts.csv", [["온도"], ["175"], ["170.500"]])                # 170.5 or 170500
    _, res = run_raw("ml_map.py", "-f", path, "--points", pts)
    assert res["status"] == "error" and res["agent_directive"].startswith("STOP_RETRY:"), res     # 2.9.21: ok
    pts = write_csv(tmp_path / "pts2.csv", [["온도"], ["175"], ["176"]])
    _, res = run_raw("ml_map.py", "-f", path, "--points", pts)
    assert res["status"] == "ok", res


def test_map_comma_number_stops_with_command(tmp_path):
    A = np.arange(0, 2001, 500, dtype=float)
    p = write_df(tmp_path / "g.csv", pd.DataFrame({"A": A, "Y": A * 2}))
    bad = run_fail("ml_map.py", "-f", p, "-i", "A=1,000", "-y", "Y")            # 2.9.13: '숫자가 아닙니다' without a fix
    rec = bad["message"].split("[추천] ")[1].strip()
    assert " -i A=1000" in rec
    assert run_tool("ml_map.py", *split_command(rec)[2:])["points"][0]["values"]["Y"] == 2000
    bad = run_fail("ml_map.py", "-f", p, "--find", "Y=1,000", "--vary", "A")
    assert '--find Y=1000' in bad["message"]


# ---- 값 열마다 따로 (MultiTable): 빈 칸, 근거 행, 범위, 쓰지 않은 열 경고 ----------------


def test_map_blank_in_other_value_column_keeps_row(tmp_path):
    A = np.arange(0, 101, 10, dtype=float)
    df = pd.DataFrame({"A": A, "B": A * 2, "C": A ** 2 / 10})
    df.loc[5, "B"] = np.nan                                  # the A=50 row has a blank B only
    p = tmp_path / "m.csv"
    df.to_csv(p, index=False)
    res = run_tool("ml_map.py", "-f", str(p), "-i", "A=50")
    assert res["points"][0]["values"]["C"] == pytest.approx(250.0)      # 2.9.11: 260 (row dropped)
    assert res["points"][0]["values"]["B"] == pytest.approx(100.0)


def test_multitable_single_group_is_plain_table():
    import _ml_runtime as rt
    df = pd.DataFrame({"A": [0.0, 1, 2, 3], "C": [0.0, 1, 4, 9], "__row": [2, 3, 4, 5]})
    w1, w2 = [], []
    a = rt.MultiTable(df, ["A"], ["C"], w1).at(np.array([1.5]))
    b = rt.Table(df, ["A"], ["C"], w2).at(np.array([1.5]))
    assert a == b and w1 == w2 == []
    assert list(itertools.chain(a["values"])) == ["C"]


def _frame():
    A = np.arange(0, 101, 10, dtype=float)
    df = pd.DataFrame({"A": A, "B": A * 2, "C": A ** 2 / 10, "D": np.nan, "__row": np.arange(2, 13)})
    df.loc[5, "B"] = np.nan                         # A=50: B blank -> B and C use different rows
    df.loc[3, "D"] = 7.0                            # D: one value only
    return df


def test_multitable_basis_per_column_and_sparse_column():
    w = []
    t = rt.MultiTable(_frame(), ["A"], ["B", "C", "D"], w)                   # 2.9.12: InputError for D
    assert t.values == ["B", "C"] and any(w_.startswith("D:") for w_ in w)
    r = t.at(np.array([50.0]))
    rows_b = {b["row"] for b in r["basis_by_column"]["B"]}
    rows_c = {b["row"] for b in r["basis_by_column"]["C"]}
    assert rows_c == {7} and rows_b == {6, 8}                                 # 2.9.12: parts[0] basis for every column
    assert len(w) == len(set(w))
    # range and nearest rows come from every key-valid row, not only the rows of B
    assert len(t.df) == 11 and t.lo[0] == 0 and t.hi[0] == 100


def test_multitable_dedupes_table_warnings():
    df = pd.DataFrame({"A": [0.0, 0, 10, 20, 30], "B": [1.0, 2, 3, 4, np.nan], "C": [1.0, 3, 5, np.nan, 7],
                       "__row": [2, 3, 4, 5, 6]})
    w = []
    rt.MultiTable(df, ["A"], ["B", "C"], w)
    assert sum("평균" in x for x in w) == 1                                    # 2.9.12: once per column group


def test_map_excel_basis_per_column(tmp_path):
    df = _frame().drop(columns=["__row", "D"])
    p = tmp_path / "m.csv"
    df.to_csv(p, index=False)
    res = run_tool("ml_map.py", "-f", str(p), "-i", "A=50", "--save", "out.xlsx")
    assert res["columns"] == ["B", "C"]
    sheet = pd.read_excel(res["output_file"], sheet_name="결과")
    assert "근거_B" in sheet.columns and "근거_C" in sheet.columns
    assert "7행" in sheet["근거_C"].iloc[0] and "6행" in sheet["근거_B"].iloc[0]


@pytest.fixture(scope="module")
def lines(tmp_path_factory):
    work = tmp_path_factory.mktemp("lines")
    df = pd.DataFrame({"라인": ["L1"] * 4 + ["L2"] * 4, "A": [10.0, 20, 30, 40] * 2,
                       "Y1": [20.0, 40, 60, 80] * 2, "Y2": [np.nan] * 4 + [11.0, 21, 31, 41]})
    return write_df(work / "map3.csv", df)


def test_map_value_column_empty_after_text_filter(lines):
    res = run_tool("ml_map.py", "-f", lines, "-i", "라인=L1,A=25", "-y", "Y1,Y2")  # 2.9.13: '숫자 열이 아닙니다'
    assert res["columns"] == ["Y1"] and any(w.startswith("Y2:") for w in res["warnings"])
    res = run_tool("ml_map.py", "-f", lines, "-i", "라인=L1,A=25")                # 2.9.13: Y2 dropped silently
    assert any(w.startswith("Y2:") for w in res["warnings"])
    bad = run_fail("ml_map.py", "-f", lines, "-i", "라인=L1,A=25", "-y", "Y2")
    assert "숫자 열이 아닙니다" not in bad["message"] and "Y2" in bad["message"]


def test_map_dropped_column_is_not_unused_input(tmp_path):
    df = pd.DataFrame({"A": [0.0, 7, 15, 31, 46, 60], "Y1": [1.0, 3, 2, 6, 4, 9], "Y2": [np.nan] * 5 + [1.0]})
    res = run_tool("ml_map.py", "-f", write_df(tmp_path / "s.csv", df), "-i", "A=20")
    assert res["columns"] == ["Y1"]
    assert not any("기준으로 주지 않은 열" in w for w in res.get("warnings", []))   # 2.9.13: ['Y2'] as an input


def test_multitable_range_from_kept_tables_and_out_of_range_per_group():
    A = np.arange(0, 101, 10, dtype=float)
    df = pd.DataFrame({"A": A, "Y1": np.where(A <= 60, A * 2, np.nan), "Y2": np.where(A <= 40, A + 1, np.nan),
                       "Y3": np.where(A == 100, 5.0, np.nan), "__row": np.arange(2, 13)})
    w = []
    t = rt.MultiTable(df, ["A"], ["Y1", "Y2", "Y3"], w)
    assert t.values == ["Y1", "Y2"]
    assert t.hi[0] == 60 and len(t.df) == 7                                   # 2.9.13: 100 (row of Y3 only)
    oor = t.at(np.array([80.0]))["out_of_range"]["A"]
    assert "Y1:" in oor and "Y2:" in oor                                      # 2.9.13: last group wins


def test_map_unused_columns_use_the_filtered_rows(tmp_path):
    df = pd.DataFrame({"라인": ["X"] * 6 + ["Y"] * 6, "A": [10, 10, 20, 20, 30, 30] * 2, "B": [1, 2] * 6,
                       "Y1": [20, 30, 30, 40, 40, 50, 120, 130, 130, 140, 140, 150],
                       "Y3": [None] * 6 + [9, 8, 19, 18, 29, 28]})
    res = run_tool("ml_map.py", "-f", write_df(tmp_path / "g.csv", df), "-i", "라인=X,A=15", "-y", "Y1")
    unused = [w for w in res.get("warnings", []) if "기준으로 주지 않은 열" in w]
    assert unused and "'B'" in unused[0] and "Y3" not in unused[0]                # 2.9.14: Y3 (empty for X) listed


# ---- 묻지 않은 값 열의 섞인 % (24차 확인 검사 LOW-3, 2.9.42) ------------------------------------------------------


@pytest.fixture
def pct_mixed_grid(tmp_path):
    rows = [[t, w, round(0.2 * t - w, 2), (f"{w}%" if t == 150 else w / 100)] for t in (150, 170, 190) for w in (5, 10, 15)]
    df = pd.DataFrame(rows, columns=["온도", "함수율", "수율", "메모율"])
    return write_df(tmp_path / "g.csv", df)


def test_map_without_y_skips_unasked_mixed_percent_column(pct_mixed_grid):
    res = run_tool("ml_map.py", "-f", pct_mixed_grid, "-i", "온도=160,함수율=10")
    assert res["columns"] == ["수율"], res
    assert res["points"][0]["values"]["수율"] == pytest.approx(22.0)
    assert any("['메모율']" in w and "섞여" in w for w in res["warnings"]), res


def test_map_asked_mixed_percent_column_still_stops(pct_mixed_grid):
    res = run_fail("ml_map.py", "-f", pct_mixed_grid, "-i", "온도=160,함수율=10", "-y", "메모율")
    assert res["agent_directive"].startswith("STOP_RETRY") and "'메모율'" in res["message"], res
    res = run_tool("ml_map.py", "-f", pct_mixed_grid, "-i", "온도=160,함수율=10", "-y", "수율")
    assert not any("메모율" in w for w in res.get("warnings", [])), res


def test_map_mixed_percent_key_column_still_stops(pct_mixed_grid):
    res = run_fail("ml_map.py", "-f", pct_mixed_grid, "-i", "온도=160,메모율=0.1")
    assert "'메모율'" in res["message"], res


def test_inline_input_joined_by_semicolon_or_space_gives_fixed_command(grid2):
    """회귀 (2.9.69): -i "온도=12; 압력=3"은 '온도' 하나(값 '12; 압력=3')로 읽혀 압력이 빠진 것으로 보였다 (ml_predict는
    중앙값으로 채울지 물음). 이제 STOP_RETRY + [추천] ','로 고친 전체 명령. 텍스트 값의 띄어쓰기, 천 단위 숫자는 그대로."""
    for given in ("온도=12; 압력=3", "온도=12 압력=3"):
        proc, out = run_raw("ml_map.py", *P(grid2, "-i", given))
        assert proc.returncode == 1 and out["agent_directive"].startswith("STOP_RETRY:"), out
        args = split_command(recommendation(out["message"]))
        assert args[args.index("-i") + 1] == "온도=12,압력=3", args
    assert rt.parse_inline_input("비고=a b,온도=1,000").to_dict("records") == [{"비고": "a b", "온도": "1,000"}]
    assert rt.parse_inline_input("온도 = 12, 압력 = 3").to_dict("records") == [{"온도": "12", "압력": "3"}]
    with pytest.raises(rt.InputError) as e:                       # where a spaced name starts is not clear: no [추천]
        rt.parse_inline_input("반응 온도=1 반응 압력=2")
    assert "[추천]" not in str(e.value) and "','" in str(e.value)
