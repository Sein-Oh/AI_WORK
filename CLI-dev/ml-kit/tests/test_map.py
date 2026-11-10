# -*- coding: utf-8 -*-
"""표 보간(ml_map.py): 격자·흩어진 점 보간, 외삽 경고, 가까운 실측 행, 역산(모든 교차점), 조회 파일(--points),
값 열마다 다른 행(MultiTable), '제외' 행. 격자 기대값은 Excel FORECAST 수식으로 손 계산한 값이다.

--points 열 이름 검사는 test_names.py.
"""
import itertools
import os
import sys

import json

import numpy as np
import openpyxl
import pandas as pd
import pytest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from mlkit_helpers import split_command, recommendation, recommended_command, run_fail, run_raw, run_tool, write_book, write_csv, write_df  # noqa: E402
import _ml_runtime as rt  # noqa: E402


def test_find_with_too_many_inputs_recommends_the_trained_model(tmp_path):
    """2.9.80 집 PC Gemma P3: 입력 6개로 ml_map --find를 돌려 '기준 열은 3개까지'만 받고 사용자에게 범위를 물음 ->
    이 파일로 학습한 모델이 있으면 같은 질문의 ml_predict.py --find 명령을 [추천]. 학습 전에는 [추천] 없음."""
    r = np.random.default_rng(3)
    df = pd.DataFrame({k: r.uniform(0, 10, 80) for k in ("a", "b", "c", "d")})
    df["y"] = 3 * df.a + df.b - df.c + 0.5 * df.d
    f = str(tmp_path / "four.csv")
    df.to_csv(f, index=False)
    argv = ["--file", f, "--find", "y=15", "--vary", "a", "--input", "b=5,c=5,d=5"]
    res = run_fail("ml_map.py", *argv)
    assert "3개까지" in res["message"] and "[추천]" not in res["message"], res
    assert run_tool("ml_train.py", "--file", f, "--features", "a,b,c,d", "--targets", "y")["status"] == "ok"
    res = run_fail("ml_map.py", *argv)
    name, fixed = recommended_command(res["message"])
    assert name == "ml_predict.py" and fixed[fixed.index("--find") + 1] == "y=15" and fixed[fixed.index("--model") + 1] == f, fixed
    assert run_tool("ml_predict.py", *fixed)["status"] == "ok"



def test_model_find_hint_checks_the_model(tmp_path):
    """--find 상한을 넘을 때의 모델 안내는 모델 정보로 판단: 이름은 ml_map이 맞춘 데이터 열 이름(단위 표기 그대로 옮기지 않음),
    --input에 모델 입력이 아닌 열이 있으면 그 열을 뺀 명령을 [후보](고정 조건이 달라짐), 목표가 모델 타겟이 아니면 안내 없음."""
    r = np.random.default_rng(4)
    df = pd.DataFrame({k: r.uniform(0, 10, 80) for k in ("a(°C)", "b", "c", "e")})
    df["y"] = 3 * df["a(°C)"] + df.b - df.c
    df["z"] = df.b * 2
    f = str(tmp_path / "five.csv")
    df.to_csv(f, index=False)
    assert run_tool("ml_train.py", "--file", f, "--features", "a(°C),b,c", "--targets", "y")["status"] == "ok"
    res = run_fail("ml_map.py", "--file", f, "--find", "y=15", "--vary", "a", "--input", "b=5,c=5,e=5")
    assert "[추천]" not in res["message"] and "[후보] 1) " in res["message"] and res["error_type"] == "needs_user_decision", res
    cmd = split_command(res["confirm_commands"][0]["command"])
    assert cmd[cmd.index("--vary") + 1] == "a(°C)" and cmd[cmd.index("--input") + 1] == "b=5,c=5", cmd     # e는 모델 입력이 아님
    assert run_tool("ml_predict.py", *cmd[2:])["status"] == "ok"
    res = run_fail("ml_map.py", "--file", f, "--find", "z=8", "--vary", "a", "--input", "b=5,c=5,e=5")       # z는 모델 타겟이 아님
    assert "ml_predict.py" not in json.dumps(res, ensure_ascii=False), res

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
    return ["--file", path, *args]


def test_grid_1d_matches_excel_forecast(grid1):
    r = run_tool("ml_map.py", *P(grid1, "--input", "A=12.5"))
    pt = r["points"][0]
    assert r["mode"] == "grid" and pt["values"] == {"B": 16.25, "C": 87.5} and not pt["extrapolated"]
    assert pt["basis"] == [{"row": 4, "weight": 0.5}, {"row": 5, "weight": 0.5}]            # A=10 is Excel row 4
    assert run_tool("ml_map.py", *P(grid1, "--input", "A=10"))["points"][0]["values"]["B"] == 10.0
    up = run_tool("ml_map.py", *P(grid1, "--input", "A=103"))
    assert up["points"][0]["values"] == {"B": 1058.5, "C": -3.0} and up["points"][0]["extrapolated"]
    assert "3 높음 (데이터 폭의 3%)" in up["points"][0]["out_of_range"]["A"] and up["warnings"]
    assert up["agent_directive"].startswith("VERIFIED:") and "extrapolated=true" in up["agent_directive"]
    down = run_tool("ml_map.py", *P(grid1, "--input", "A=-2", "--targets", "B"))
    assert down["points"][0]["values"] == {"B": -1.0} and down["columns"] == ["B"]


def test_grid_2d_bilinear(grid2):
    r = run_tool("ml_map.py", *P(grid2, "--input", "온도=25,압력=2.5", "--targets", "효율"))
    pt = r["points"][0]
    assert r["mode"] == "grid" and pt["method"] == "격자 보간"
    assert pt["values"]["효율"] == pytest.approx(2 * 25 + 10 * 2.5 + 25 * 2.5)            # bilinear is exact on t*p
    assert len(pt["basis"]) == 4 and sum(b["weight"] for b in pt["basis"]) == pytest.approx(1.0)
    ext = run_tool("ml_map.py", *P(grid2, "--input", "온도=110,압력=3"))["points"][0]
    assert ext["extrapolated"] and ext["values"]["효율"] == pytest.approx(2 * 110 + 30 + 330)


def test_scattered_two_keys_enclosing_rows_and_far_point(process):
    args = ("--header-row", "3", "--unit-row", "4", "--targets", "생산효율")
    r = run_tool("ml_map.py", *P(process, "--input", "반응온도=240,소비전력=60", *args))
    pt = r["points"][0]
    assert r["mode"] == "scattered" and pt["method"].startswith("둘러싼 실측 3점")
    assert pt["values"]["생산효율(%)"] == pytest.approx(0.2 * 240 + 0.1 * 60, abs=1e-6)      # linear data: exact
    assert len(pt["basis"]) == 3 and all(b["row"] >= 5 for b in pt["basis"])
    assert len(r["nearest_rows"]) == 5 and r["nearest_rows"][0]["distance"] <= r["nearest_rows"][-1]["distance"]
    assert "센서오류" not in str(r) and any("제외" in w for w in r["warnings"])            # marked row not used

    far = run_tool("ml_map.py", *P(process, "--input", "반응온도=30,소비전력=200", *args))
    fp = far["points"][0]
    assert fp["extrapolated"] and fp["method"].startswith("전체 추세선 외삽")
    assert abs(fp["values"]["생산효율(%)"]) < 100                                           # never the exploding end-point slope
    assert "낮음 (데이터 폭의 2." in fp["out_of_range"]["반응온도(°C)"] and "높음" in fp["out_of_range"]["소비전력(kW)"]


def test_one_uneven_key_is_scattered_and_category_filter(process):
    r = run_tool("ml_map.py", *P(process, "--header-row", "3", "--unit-row", "4", "--input", "반응온도=240,라인=B", "--targets", "생산효율"))
    assert r["mode"] == "scattered" and r["points"][0]["method"] == "위아래 실측 2점 보간"
    assert r["points"][0]["input"]["라인"] == "B" and all(n["라인"] == "B" for n in r["nearest_rows"])
    bad = run_fail("ml_map.py", *P(process, "--header-row", "3", "--unit-row", "4", "--input", "반응온도=240,라인=D"))
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
        res = run_tool("ml_map.py", "--file", str(p), "--input", f"온도=170,코드={c}", "--targets", "수율")
        assert res["keys"] == ["온도"] and res["points"][0]["input"]["코드"] == c, res
        assert res["points"][0]["values"]["수율"] == pytest.approx(340 + off, abs=1e-6), res
        assert all(str(n["코드"]) in (c, c + ".0") for n in res["nearest_rows"]), res["nearest_rows"]
    bad = run_fail("ml_map.py", "--file", str(p), "--input", "온도=170,코드=07", "--targets", "수율")
    assert "[후보] 1) 007 2) 12 3) 7" in bad["message"] and "[추천]" not in bad["message"], bad


def test_two_digit_numbers_with_a_leading_zero_are_interpolated(tmp_path):
    """3차 검토(2.9.63): '05, 10, 15, 20'의 '05' 하나로 숫자 열을 코드 열로 보아 보간 대신 틀린 [후보]를 냈다."""
    p = tmp_path / "rt.csv"
    p.write_text("온도,압력,수율\n05,1,50\n10,1,60\n15,1,70\n20,1,80\n05,2,55\n10,2,65\n15,2,75\n20,2,85\n",
                 encoding="utf-8")
    res = run_tool("ml_map.py", "--file", str(p), "--input", "온도=12,압력=1", "--targets", "수율")
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
    r = run_tool("ml_map.py", "--file", str(p), "--find", "Y=64", "--vary", "A")
    assert [round(s["A"], 6) for s in r["solutions"] if not s["extrapolated"]] == [20.0, 80.0]
    none = run_tool("ml_map.py", "--file", str(p), "--find", "Y=150", "--vary", "A")
    assert none["count"] == 0 and none["reachable"]["max"] == pytest.approx(100) and "찾은 값 없음" in none["preview"]


def test_reverse_scattered_with_fixed_key(process):
    r = run_tool("ml_map.py", *P(process, "--header-row", "3", "--unit-row", "4",
                                 "--find", "생산효율=54", "--vary", "반응온도", "--input", "소비전력=60"))
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
    assert run_tool("ml_map.py", *P(grid1, "--input", "a=10"))["keys"] == ["A"]                     # case only: matched
    miss = run_fail("ml_map.py", *P(grid1, "--input", "Ax=10"))
    assert "[후보] 1) 'A'" in miss["message"] and "[추천]" not in miss["message"]    # 2.9.31: short ASCII names of other length
    assert run_fail("ml_map.py", *P(grid1, "--input", "A=abc"))["error_type"] == "input"
    assert run_fail("ml_map.py", *P(grid1, "--find", "B=90", "--vary", "A", "--input", "A=3"))["error_type"] == "input"
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
    res = run_tool("ml_map.py", "--file", grid_gap["data"], "--points", grid_gap["pts"], "--save", "p.xlsx")
    sheet = pd.read_excel(res["output_file"], sheet_name="결과")
    assert "근거" not in sheet.columns and "근거_Y1" in sheet.columns            # 2.9.13: 근거 and 근거_<열> mixed
    assert sheet["근거_Y1"].notna().all()
    assert pd.api.types.is_numeric_dtype(sheet["A"]) and pd.api.types.is_numeric_dtype(sheet["B"])   # 2.9.13: text


def test_map_points_header_hint(tmp_path):
    rows = [["맵 기록", "2026-11-02"], ["A", "B"]] + [[float(i), float(i * i)] for i in range(10)]
    data = write_book(tmp_path / "m.xlsx", rows)
    pts = tmp_path / "pts.csv"
    pd.DataFrame({"A": [2.5, 3.5]}).to_csv(pts, index=False, encoding="utf-8-sig")
    bad = run_fail("ml_map.py", "--file", data, "--points", str(pts), "--targets", "B")
    _, rec = recommended_command(bad["message"])                  # 2.9.15: no hint, 2.9.71: 전체 명령
    assert rec[rec.index("--header-row") + 1] == "2", rec
    ok = run_tool("ml_map.py", "--file", data, "--points", str(pts), "--targets", "B", "--header-row", "2")
    assert ok["count"] == 2


def test_map_points_keep_data_decimal_comma(tmp_path):
    data = [["온도", "수율"]] + [[f"{170 + i},5", f"{80 + i},25"] for i in range(10)]
    path = write_csv(tmp_path / "grid.csv", data)
    pts = write_csv(tmp_path / "pts.csv", [["온도"], ["175"], ["170.500"]])                # 170.5 or 170500
    _, res = run_raw("ml_map.py", "--file", path, "--points", pts)
    assert res["status"] == "error" and res["agent_directive"].startswith("STOP_RETRY:"), res     # 2.9.21: ok
    pts = write_csv(tmp_path / "pts2.csv", [["온도"], ["175"], ["176"]])
    _, res = run_raw("ml_map.py", "--file", path, "--points", pts)
    assert res["status"] == "ok", res


def test_map_comma_number_stops_with_command(tmp_path):
    A = np.arange(0, 2001, 500, dtype=float)
    p = write_df(tmp_path / "g.csv", pd.DataFrame({"A": A, "Y": A * 2}))
    bad = run_fail("ml_map.py", "--file", p, "--input", "A=1,000", "--targets", "Y")            # 2.9.13: '숫자가 아닙니다' without a fix
    rec = bad["message"].split("[추천] ")[1].strip()
    assert " --input A=1000" in rec
    assert run_tool("ml_map.py", *split_command(rec)[2:])["points"][0]["values"]["Y"] == 2000
    bad = run_fail("ml_map.py", "--file", p, "--find", "Y=1,000", "--vary", "A")
    assert '--find Y=1000' in bad["message"]


# ---- 값 열마다 따로 (MultiTable): 빈 칸, 근거 행, 범위, 쓰지 않은 열 경고 ----------------


def test_map_blank_in_other_value_column_keeps_row(tmp_path):
    A = np.arange(0, 101, 10, dtype=float)
    df = pd.DataFrame({"A": A, "B": A * 2, "C": A ** 2 / 10})
    df.loc[5, "B"] = np.nan                                  # the A=50 row has a blank B only
    p = tmp_path / "m.csv"
    df.to_csv(p, index=False)
    res = run_tool("ml_map.py", "--file", str(p), "--input", "A=50")
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
    res = run_tool("ml_map.py", "--file", str(p), "--input", "A=50", "--save", "out.xlsx")
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
    res = run_tool("ml_map.py", "--file", lines, "--input", "라인=L1,A=25", "--targets", "Y1,Y2")  # 2.9.13: '숫자 열이 아닙니다'
    assert res["columns"] == ["Y1"] and any(w.startswith("Y2:") for w in res["warnings"])
    res = run_tool("ml_map.py", "--file", lines, "--input", "라인=L1,A=25")                # 2.9.13: Y2 dropped silently
    assert any(w.startswith("Y2:") for w in res["warnings"])
    bad = run_fail("ml_map.py", "--file", lines, "--input", "라인=L1,A=25", "--targets", "Y2")
    assert "숫자 열이 아닙니다" not in bad["message"] and "Y2" in bad["message"]


def test_map_dropped_column_is_not_unused_input(tmp_path):
    df = pd.DataFrame({"A": [0.0, 7, 15, 31, 46, 60], "Y1": [1.0, 3, 2, 6, 4, 9], "Y2": [np.nan] * 5 + [1.0]})
    res = run_tool("ml_map.py", "--file", write_df(tmp_path / "s.csv", df), "--input", "A=20")
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
    res = run_fail("ml_map.py", "--file", write_df(tmp_path / "g.csv", df), "--input", "라인=X,A=15", "--targets", "Y1")
    # 검토 B A2: A x B 격자에서 B를 주지 않으면 평균하지 않고 멈춤 (전에는 경고만 하고 VERIFIED)
    assert "'B'" in res["message"] and "Y3" not in res["message"], res                # 2.9.14: Y3 (empty for X) listed


# ---- 묻지 않은 값 열의 섞인 % (24차 확인 검사 LOW-3, 2.9.42) ------------------------------------------------------


@pytest.fixture
def pct_mixed_grid(tmp_path):
    rows = [[t, w, round(0.2 * t - w, 2), (f"{w}%" if t == 150 else w / 100)] for t in (150, 170, 190) for w in (5, 10, 15)]
    df = pd.DataFrame(rows, columns=["온도", "함수율", "수율", "메모율"])
    return write_df(tmp_path / "g.csv", df)


def test_map_without_y_skips_unasked_mixed_percent_column(pct_mixed_grid):
    res = run_tool("ml_map.py", "--file", pct_mixed_grid, "--input", "온도=160,함수율=10")
    assert res["columns"] == ["수율"], res
    assert res["points"][0]["values"]["수율"] == pytest.approx(22.0)
    assert any("['메모율']" in w and "섞여" in w for w in res["warnings"]), res


def test_map_asked_mixed_percent_column_still_stops(pct_mixed_grid):
    res = run_fail("ml_map.py", "--file", pct_mixed_grid, "--input", "온도=160,함수율=10", "--targets", "메모율")
    assert res["agent_directive"].startswith("STOP_RETRY") and "'메모율'" in res["message"], res
    res = run_tool("ml_map.py", "--file", pct_mixed_grid, "--input", "온도=160,함수율=10", "--targets", "수율")
    assert not any("메모율" in w for w in res.get("warnings", [])), res


def test_map_mixed_percent_key_column_still_stops(pct_mixed_grid):
    res = run_fail("ml_map.py", "--file", pct_mixed_grid, "--input", "온도=160,메모율=0.1")
    assert "'메모율'" in res["message"], res


def test_inline_input_joined_by_semicolon_or_space_gives_fixed_command(grid2):
    """회귀 (2.9.69): --input "온도=12; 압력=3"은 '온도' 하나(값 '12; 압력=3')로 읽혀 압력이 빠진 것으로 보였다 (ml_predict는
    중앙값으로 채울지 물음). 이제 STOP_RETRY + [추천] ','로 고친 전체 명령. 텍스트 값의 띄어쓰기, 천 단위 숫자는 그대로."""
    for given in ("온도=12; 압력=3", "온도=12 압력=3"):
        proc, out = run_raw("ml_map.py", *P(grid2, "--input", given))
        assert proc.returncode == 1 and out["agent_directive"].startswith("STOP_RETRY:"), out
        args = split_command(recommendation(out["message"]))
        assert args[args.index("--input") + 1] == "온도=12,압력=3", args
    assert rt.parse_inline_input("비고=a b,온도=1,000").to_dict("records") == [{"비고": "a b", "온도": "1,000"}]
    assert rt.parse_inline_input("온도 = 12, 압력 = 3").to_dict("records") == [{"온도": "12", "압력": "3"}]
    with pytest.raises(rt.InputError) as e:                       # where a spaced name starts is not clear: no [추천]
        rt.parse_inline_input("반응 온도=1 반응 압력=2")
    assert "[추천]" not in str(e.value) and "','" in str(e.value)


# ---- 검토 B (2026-11-10, 2.9.87): 틀린 값을 VERIFIED로 내던 경우, 단위 행, 날짜 열, 평평한 구간 ----------------------


def _xlsx(path, df):
    df.to_excel(path, index=False)
    return str(path)


def test_forward_with_two_valued_key_asks_for_find(tmp_path):
    """A1: 산 모양 열(수율)을 기준으로 주면 수율 순으로 이웃한 두 행(온도 30과 75)을 보간해 온도 50.5를 VERIFIED로 냈다.
    수율 80이 되는 온도는 두 곳이므로 같은 질문의 --find 전체 명령을 [추천]. 오르기만 하는 열은 그대로 사잇값."""
    T = np.arange(0, 101, 5)
    f = _xlsx(tmp_path / "nm.xlsx", pd.DataFrame({"온도": T, "수율": 100 - (T - 50) ** 2 / 25 + np.linspace(0, 0.2, len(T))}))
    res = run_fail("ml_map.py", "--file", f, "--input", "수율=80")
    assert res["agent_directive"].startswith("STOP_RETRY") and "50.5" not in json.dumps(res, ensure_ascii=False), res
    name, cmd = recommended_command(res["message"])
    assert name == "ml_map.py" and cmd[cmd.index("--find") + 1] == "수율=80" and cmd[cmd.index("--vary") + 1] == "온도", cmd
    ok = run_tool("ml_map.py", *cmd)
    assert ok["count"] == 2, ok
    res = run_fail("ml_map.py", "--file", f, "--input", "수율=75", "--targets", "온도")
    assert "[추천]" in res["message"], res
    mono = _xlsx(tmp_path / "mono.xlsx", pd.DataFrame({"온도": T, "수율": T ** 1.5 / 10}))
    ok = run_tool("ml_map.py", "--file", mono, "--input", "수율=50")
    assert ok["points"][0]["values"]["온도"] == pytest.approx(500 ** (2 / 3), abs=0.3), ok


def test_forward_without_a_level_key_stops(tmp_path):
    """A2: 격자에서 압력을 주지 않으면 압력 1~3을 평균한 값(31)과 '압력 2'를 VERIFIED로 냈다. 글자 열(라인)을 주지 않아도
    A·B 평균. 빠진 기준 열을 알리고 값을 묻는다 (next_commands의 <값>)."""
    rows = [{"온도(°C)": a, "압력(bar)": b, "출력(kW)": 0.1 * a + 5 * b + 0.01 * a * b} for a in (100, 150, 200, 250) for b in (1, 2, 3)]
    g2 = _xlsx(tmp_path / "g2.xlsx", pd.DataFrame(rows))
    res = run_fail("ml_map.py", "--file", g2, "--input", "온도=175")
    assert "압력(bar)" in res["message"] and "31" not in res["message"], res
    assert res["error_type"] == "needs_user_decision" and res["next_commands"][0]["needs_input"], res
    cmd = split_command(res["next_commands"][0]["command"])
    assert cmd[cmd.index("--input") + 1] == "온도=175,압력(bar)=<값>", cmd
    assert "next_commands" in res["agent_directive"], res
    res = run_fail("ml_map.py", "--file", g2, "--input", "온도=175", "--targets", "출력")
    assert "압력(bar)" in res["message"], res
    g = pd.DataFrame(rows)
    gb = g.assign(라인="B")
    gb["출력(kW)"] *= 2
    cat = _xlsx(tmp_path / "g2_cat.xlsx", pd.concat([g.assign(라인="A"), gb]))
    res = run_fail("ml_map.py", "--file", cat, "--input", "온도=175,압력=2.5")
    assert "'라인'" in res["message"] and "A, B" in res["message"] and "51.75" not in res["message"], res
    ok = run_tool("ml_map.py", "--file", cat, "--input", "온도=175,압력=2.5,라인=A")
    assert ok["points"][0]["values"]["출력(kW)"] == pytest.approx(34.375), ok
    pts = tmp_path / "pts.csv"
    pd.DataFrame({"온도": [175, 120]}).to_csv(pts, index=False)
    res = run_fail("ml_map.py", "--file", g2, "--points", str(pts))
    assert "압력(bar)" in res["message"], res


def test_single_value_key_is_dropped_or_refused_without_qhull_text(tmp_path):
    """B5: 값이 하나뿐인 기준 열을 함께 주면 qhull 오류 글 10여 줄이 message에 그대로 나왔다."""
    f = _xlsx(tmp_path / "const.xlsx", pd.DataFrame({"온도": [10, 10, 10, 10], "압력": [1, 2, 3, 4], "값": [1, 2, 3, 4]}))
    ok = run_tool("ml_map.py", "--file", f, "--input", "온도=10,압력=2.5")
    assert ok["keys"] == ["압력"] and ok["points"][0]["values"]["값"] == pytest.approx(2.5), ok
    assert any("'온도'" in w and "10 하나뿐" in w for w in ok["warnings"]), ok
    bad = run_fail("ml_map.py", "--file", f, "--input", "온도=12,압력=2.5")
    assert "10 하나뿐" in bad["message"] and "qhull" not in bad["message"].lower(), bad
    line = _xlsx(tmp_path / "line.xlsx", pd.DataFrame({"온도": [1, 2, 3, 4], "압력": [1, 2, 3, 4], "값": [1, 2, 3, 4]}))
    bad = run_fail("ml_map.py", "--file", line, "--input", "온도=2.5,압력=2.5")
    assert "qhull" not in bad["message"].lower() and "QH" not in bad["message"], bad


def test_unit_row_below_header_is_read_as_units(tmp_path):
    """C1: 머리글 아래 단위 행을 데이터로 읽어 작은 표는 '숫자 기준 열이 하나 이상 필요합니다'로 멈추고, 큰 표는 이름에
    단위가 붙지 않아 '유량(L/min)'을 받지 못했다. 단위 행으로 읽는다 (warnings에 알림)."""
    def pump(name, k):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["유량", "양정", "효율"])
        ws.append(["L/min", "m", "%"])
        for i in range(k):
            q = 10 + 5 * i
            ws.append([q, 50 - 0.5 * q, 30 + q])
        wb.save(tmp_path / name)
        return str(tmp_path / name)
    p2, p3 = pump("pump2.xlsx", 6), pump("pump3.xlsx", 12)
    ok = run_tool("ml_map.py", "--file", p2, "--input", "유량=35")
    assert ok["keys"] == ["유량(L/min)"] and ok["points"][0]["values"]["효율(%)"] == pytest.approx(65), ok
    assert any("단위 행" in w for w in ok["warnings"]), ok
    rev = run_tool("ml_map.py", "--file", p2, "--find", "효율=60", "--vary", "유량")
    assert rev["solutions"][0]["유량(L/min)"] == pytest.approx(30), rev
    assert run_tool("ml_map.py", "--file", p3, "--input", "유량(L/min)=37")["points"][0]["values"]["효율(%)"] == pytest.approx(67)


def test_date_column_filter_reads_dates(tmp_path):
    """C2: 날짜 열은 '2026-01-06 00:00:00'으로 적어야 걸렀다."""
    d = pd.to_datetime(["2026-01-05"] * 3 + ["2026-01-06"] * 3)
    f = _xlsx(tmp_path / "dt.xlsx", pd.DataFrame({"날짜": d, "온도": [20, 25, 30] * 2, "값": [1, 2, 3, 11, 12, 13]}))
    ok = run_tool("ml_map.py", "--file", f, "--input", "날짜=2026-01-06,온도=27.5", "--targets", "값")
    assert ok["points"][0]["values"]["값"] == pytest.approx(12.5), ok
    bad = run_fail("ml_map.py", "--file", f, "--input", "날짜=2026-01-09,온도=27.5", "--targets", "값")
    assert "[후보] 1) 2026-01-05 2) 2026-01-06" in bad["message"] and "00:00" not in bad["message"], bad
    cmd = split_command(bad["confirm_commands"][1]["command"])
    assert run_tool("ml_map.py", *cmd[2:])["points"][0]["values"]["값"] == pytest.approx(12.5)


def test_flat_stretch_is_one_span(tmp_path):
    """D3: 평평한 구간(10~30 전체가 답)을 마디 값 10, 20, 30으로 나열했다."""
    f = _xlsx(tmp_path / "flat.xlsx", pd.DataFrame({"x": [0, 10, 20, 30, 40], "y": [1, 5, 5, 5, 9]}))
    res = run_tool("ml_map.py", "--file", f, "--find", "y=5", "--vary", "x")
    assert res["count"] == 1 and res["solutions"][0]["from"] == 10 and res["solutions"][0]["to"] == 30, res
    assert "10~30" in res["preview"], res


def test_points_preview_counts_failed_rows(tmp_path):
    """D4: 계산하지 못한 조회도 '보간' 건수에 넣고, 글자 값도 '빈칸'이라고 했다."""
    a = np.arange(0, 101, 10)
    f = write_df(tmp_path / "g.csv", pd.DataFrame({"A": a, "B": a * 2}))
    pts = tmp_path / "q.csv"
    pd.DataFrame({"메모": ["a", "b", "c", "d"], "A": ["15", "", "abc", "25"]}).to_csv(pts, index=False)
    res = run_tool("ml_map.py", "--file", f, "--points", str(pts))
    assert "조회 4건 중 2건 보간" in res["preview"] and "2건 계산 불가" in res["preview"], res
    assert any("빈칸인 조회 1건" in w for w in res["warnings"]) and any("숫자가 아닌 조회 1건" in w for w in res["warnings"]), res
    sheet = pd.read_excel(res["output_file"], sheet_name="결과")
    assert sheet["방식"][2] == "계산 불가 (기준 값이 숫자가 아님)" and sheet["방식"][1] == "계산 불가 (기준 값 빈칸)", sheet


def test_reverse_far_extrapolation_is_no_answer(tmp_path):
    """A7: 데이터 폭의 1.6배 밖 답(온도 -137.6)을 VERIFIED 답으로 냈다. 가까운 외삽(폭의 5%)은 그대로 답."""
    f = _xlsx(tmp_path / "k1.xlsx", pd.DataFrame({"온도": [10, 20, 35, 50, 70, 100], "값": [5, 9, 12, 9, 5, 1]}))
    res = run_tool("ml_map.py", "--file", f, "--find", "값=20", "--vary", "온도")
    assert res["count"] == 0 and "찾은 값 없음" in res["preview"] and "-137" not in res["preview"], res
    assert any("-137" in w for w in res["warnings"]), res
    a = np.arange(0, 101, 5)
    g = _xlsx(tmp_path / "g1.xlsx", pd.DataFrame({"A": a, "B": a * a / 10}))
    assert run_tool("ml_map.py", "--file", g, "--find", "B=5000", "--vary", "A")["count"] == 0
    near = run_tool("ml_map.py", "--file", g, "--find", "B=1100", "--vary", "A")
    assert near["count"] == 1 and near["solutions"][0]["extrapolated"], near


def test_reverse_with_many_noisy_answers_gives_model_command(tmp_path):
    """D5: 흩어진 기록에서 답이 많으면 글로만 '예측 모델로 보라'고 했다. 학습한 모델이 있으면 그 명령을 next_commands로."""
    r = np.random.default_rng(2)
    x = r.uniform(0, 100, 120)
    f = write_df(tmp_path / "noisy.csv", pd.DataFrame({"x": x.round(3), "y": (x + r.normal(0, 8, 120)).round(3)}))
    assert run_tool("ml_train.py", "--file", f, "--features", "x", "--targets", "y")["status"] == "ok"
    res = run_tool("ml_map.py", "--file", f, "--find", "y=50", "--vary", "x")
    assert res["count"] > 3 and "모델로 찾기" in res["preview"], res
    nxt = split_command(res["next_commands"][0]["command"])
    assert os.path.basename(nxt[1]) == "ml_predict.py" and nxt[nxt.index("--find") + 1] == "y=50", nxt
    assert run_tool("ml_predict.py", *nxt[2:])["count"] >= 1
    assert "next_commands" in res["agent_directive"], res


def test_nan_and_percent_values(tmp_path):
    """D2·A3(ml_map): --find 목표의 '%'는 떼고 읽고(경고), 'nan'·'inf'는 숫자가 아님."""
    a = np.arange(0, 101, 10)
    f = write_df(tmp_path / "g.csv", pd.DataFrame({"A": a, "B": a * 2}))
    res = run_tool("ml_map.py", "--file", f, "--find", "B=60%", "--vary", "A")
    assert res["solutions"][0]["A"] == pytest.approx(30) and any("%" in w for w in res["warnings"]), res
    for bad in ("B=nan", "B=inf"):
        assert run_fail("ml_map.py", "--file", f, "--find", bad, "--vary", "A")["error_type"] == "input"
    assert run_fail("ml_map.py", "--file", f, "--input", "A=nan")["error_type"] == "input"
