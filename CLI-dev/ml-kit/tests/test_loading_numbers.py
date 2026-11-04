# -*- coding: utf-8 -*-
"""데이터 읽기 3: 숫자 칸. 소수점 쉼표('175,6'), 천 단위('1,500', 유럽식 '1.346,8'), '%'가 붙은 숫자, 숫자로 읽지 못한 칸.

규칙: 확실하지 않으면 추측하지 않는다. 소수점 쉼표 CSV 속 점으로 묶은 정수('1.300')와 섞인 형식은 STOP_RETRY,
'%' 칸과 일반 숫자가 섞인 열도 STOP_RETRY. 학습한 모델은 소수점 쉼표 판단을 예측 파일에도 그대로 쓴다.
"""
import os
import random
import sys
from datetime import datetime

import numpy as np
import pandas as pd
import pytest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from mlkit_helpers import (H, U, data_file, data_path, family_rows, ok_res, run_raw, stop_msg,  # noqa: E402
                           write_csv, write_csv_header, write_rows)
import ml_predict as mp  # noqa: E402

N = 40
X, Y = "반응온도,압력,라인", "수율"


# ---- coerce_numeric: 천 단위만 떼고 %를 뗀다 ----------------------------------------------------


def test_coerce_numeric_object_and_str_dtypes():
    for dtype in ("object", "string"):
        s = pd.Series(["1,234.5", "-", "N/A", " 7 ", None], dtype=dtype)
        out = mp.coerce_numeric(s)
        assert out.iloc[0] == 1234.5 and out.iloc[3] == 7.0
        assert out.iloc[[1, 2, 4]].isna().all()


def test_coerce_numeric_only_strips_thousands_groups():
    s = pd.Series(["1,500", "1,234,567.5", "175,6", "1,2", " 2,000 ", "-1,000", "12"], dtype="object")
    got = mp.coerce_numeric(s).tolist()
    assert got[0] == 1500 and got[1] == 1234567.5 and got[4] == 2000 and got[5] == -1000 and got[6] == 12
    assert pd.isna(got[2]) and pd.isna(got[3])


def test_zero_group_comma_is_not_thousands():
    s = pd.Series(["0,149", "00,150", "-0,149", "1,500"], dtype="object")
    got = mp.coerce_numeric(s).tolist()
    assert pd.isna(got[0]) and pd.isna(got[1]) and pd.isna(got[2]) and got[3] == 1500


def test_coerce_numeric_strips_percent():
    s = pd.Series(["64.4%", " 1,234 % ", "5", "%", "abc", "-", "(%)"])
    got = mp.coerce_numeric(s).tolist()
    assert got[:3] == [64.4, 1234.0, 5.0] and all(pd.isna(v) for v in got[3:]), got
    n, ex = mp.percent_cells(s)
    assert n == 2 and ex == ["64.4%", "1,234 %"]


# ---- 소수점 쉼표 CSV --------------------------------------------------------------------------------------------


def _comma(v):
    return str(v).replace(".", ",") if isinstance(v, float) else v


def test_decimal_comma_csv(tmp_path):
    data = family_rows("A", 20, 1)
    path = write_csv(str(tmp_path / "c1.csv"), [["시험 기록"], H, U] + [[_comma(v) for v in r] for r in data], ";")
    df = mp.load_data_file(path, header_row=2, unit_row=3)
    t = mp.coerce_numeric(df["반응온도(°C)"])
    assert t.tolist() == [r[0] for r in data]                                     # 2.9.18: 1756.0 for '175,6'
    assert any("소수점" in w for w in mp.layout_warnings(df))
    _, res = run_raw("ml_train.py", "-f", path, "-x", X, "-y", Y, "--header-row", "2", "--unit-row", "3")
    assert res["status"] == "ok" and any("소수점" in w for w in res.get("warnings", [])), res
    # a short prediction file: csv.Sniffer alone picks ',' ('163,1;2,19;A;61,66'), the table width picks ';'
    short = write_csv(str(tmp_path / "c1p.csv"), [["시험 기록"], H, U] + [[_comma(v) for v in r] for r in data[:6]], ";")
    df = mp.load_data_file(short, header_row=2, unit_row=3)
    assert list(df.columns) == ["반응온도(°C)", "압력(bar)", "라인", "수율(%)"]
    assert mp.coerce_numeric(df["압력(bar)"]).tolist() == [r[1] for r in data[:6]]


def test_decimal_comma_mixed_with_thousands_is_refused(tmp_path):
    rows = [H, ["175,6", "1,500", "A", "61,66"], ["180,1", "2,19", "B", "70,2"]]
    path = write_csv(str(tmp_path / "c2.csv"), rows, ";")
    with pytest.raises(mp.InputError) as e:
        mp.load_data_file(path)
    assert "1,500" in str(e.value)


def test_thousands_comma_csv_unchanged(tmp_path):
    path = write_csv(str(tmp_path / "c3.csv"), [["유량", "수율"], ["1,500", "61.5"], ["2,250", "70.25"]], ",")
    df = mp.load_data_file(path)
    assert mp.coerce_numeric(df["유량"]).tolist() == [1500.0, 2250.0] and not mp.layout_warnings(df)
    path = write_csv(str(tmp_path / "c4.csv"), [["유량", "수율"], ["1,500", "61.5"], ["2,250", "70.25"]], ";")
    assert mp.coerce_numeric(mp.load_data_file(path)["유량"]).tolist() == [1500.0, 2250.0]


def _dc(x, d):
    return f"{x:.{d}f}".replace(".", ",")


def _series(seed=1):
    r = np.random.default_rng(seed)
    t, p, th = r.normal(175, 3, N), r.normal(2.2, 0.1, N), r.uniform(0.1, 0.3, N)
    return t, p, th, 0.5 * t + 20 * p + 50 * th + r.normal(0, 0.5, N)


def _num(df, col):
    return mp.coerce_numeric(df[col]).tolist()


def test_four_decimal_comma_numbers(tmp_path):
    t, _, th, y = _series()
    th = th * 100                                           # '27,8448'
    d7 = write_csv_header(tmp_path / "d7.csv", ["온도", "두께", "수율"], [[int(t[i]), _dc(th[i], 4), int(y[i])] for i in range(N)])
    df = mp.load_data_file(d7)
    assert _num(df, "두께") == [round(v, 4) for v in th]    # 2.9.19: 278448.0, VERIFIED
    assert any("소수점" in w for w in mp.layout_warnings(df))
    d6 = write_csv_header(tmp_path / "d6.csv", ["온도", "두께", "수율"],
              [[_dc(t[i], 1), _dc(th[i], 4), _dc(y[i], 1)] for i in range(N)])
    df = mp.load_data_file(d6)                              # 2.9.19: STOP_RETRY (ambiguous)
    assert _num(df, "온도") == [round(v, 1) for v in t] and _num(df, "두께") == [round(v, 4) for v in th]
    _, res = run_raw("ml_train.py", "-f", d7, "-x", "온도,두께", "-y", "수율")
    assert res["status"] == "ok" and any("소수점" in w for w in res.get("warnings", [])), res


def test_leading_zero_comma_is_decimal(tmp_path):
    t, p, th, y = _series()
    d2 = write_csv_header(tmp_path / "d2.csv", ["온도", "두께", "수율"], [[int(t[i]), _dc(th[i], 3), int(y[i])] for i in range(N)])
    df = mp.load_data_file(d2)
    assert _num(df, "두께") == [round(v, 3) for v in th]                    # 2.9.19: 149.0, VERIFIED
    assert any("소수점" in w for w in mp.layout_warnings(df))
    d1 = write_csv_header(tmp_path / "d1.csv", ["온도", "압력", "두께", "수율"],
              [[_dc(t[i], 1), _dc(p[i], 2), _dc(th[i], 3), _dc(y[i], 1)] for i in range(N)])
    df = mp.load_data_file(d1)                                              # 2.9.19: STOP_RETRY (ambiguous)
    assert _num(df, "두께") == [round(v, 3) for v in th] and _num(df, "온도") == [round(v, 1) for v in t]


def _euro(x):
    return f"{x:,.1f}".replace(",", "X").replace(".", ",").replace("X", ".")


def test_european_thousands_dot(tmp_path):
    t, _, _, y = _series()
    flow = np.random.default_rng(2).uniform(1000, 1500, N)
    rows = [[_dc(t[i], 1), _euro(flow[i]), _dc(y[i], 1)] for i in range(N)]
    df = mp.load_data_file(write_csv_header(tmp_path / "d3.csv", ["온도", "유량", "수율"], rows))
    assert _num(df, "유량") == [round(v, 1) for v in flow]                 # 2.9.19: 1.3468 for '1.346,8'
    warns = mp.layout_warnings(df)
    ex = _euro(flow[0])
    assert any("천 단위" in w and f"{ex} -> {ex.replace('.', '').replace(',', '.')}" in w for w in warns), warns
    # only European numbers, no plain decimal comma: still read as 1346.8
    rows2 = [[int(t[i]), _euro(flow[i]), int(y[i])] for i in range(N)]
    df = mp.load_data_file(write_csv_header(tmp_path / "d3b.csv", ["온도", "유량", "수율"], rows2))
    assert _num(df, "유량") == [round(v, 1) for v in flow]
    # mixed with a US number: refused, never 1.3468
    rows3 = [list(r) for r in rows]
    rows3[5][1] = "1,346.8"
    with pytest.raises(mp.InputError):
        mp.load_data_file(write_csv_header(tmp_path / "d3x0.csv", ["온도", "유량", "수율"], rows3))
    # a dot-grouped integer without a comma: 1346 or 1.346 -> refused (2.9.21 read 1346; 2.9.22부터 거부)
    rows3[5][1] = "1.346"
    with pytest.raises(mp.InputError):
        mp.load_data_file(write_csv_header(tmp_path / "d3x1.csv", ["온도", "유량", "수율"], rows3))
    _, res = run_raw("ml_train.py", "-f", write_csv_header(tmp_path / "d3t.csv", ["온도", "유량", "수율"], rows),
                     "-x", "온도,유량", "-y", "수율")
    assert res["status"] == "ok" and any("천 단위" in w for w in res.get("warnings", [])), res


def test_thousands_and_decimal_still_refused(tmp_path):
    rows = [["175,6", "1,500", "61,66"], ["180,1", "2,19", "70,2"]]
    with pytest.raises(mp.InputError):
        mp.load_data_file(write_csv_header(tmp_path / "c.csv", ["온도", "유량", "수율"], rows))


def _e2_rows():
    rows = [["온도", "유량", "수율"]]
    for i in range(12):
        rows.append([f"{175.5 + i:.1f}".replace(".", ","), f"1.{300 + 7 * i:03d}", f"{90.1 + i:.1f}".replace(".", ",")])
    return rows


def test_dot_grouped_integers_refused(tmp_path):
    # 2.9.22: never guessed (2.9.20: 1.3, 1.307, ...; 2.9.21: 1300 with a warning)
    with pytest.raises(mp.InputError, match="1.300"):
        mp.load_data_file(write_csv(tmp_path / "e2.csv", _e2_rows()))


def test_dot_grouped_integers_ambiguous_refused(tmp_path):
    # a plain decimal-point number in the file: '1.300' could be 1.3 too -> refused
    rows = _e2_rows()
    rows[3][2] = "1.5"
    with pytest.raises(mp.InputError):
        mp.load_data_file(write_csv(tmp_path / "e2a.csv", rows))
    # '0.300' cannot be a thousands group
    rows = _e2_rows()
    rows[4][1] = "0.300"
    with pytest.raises(mp.InputError):
        mp.load_data_file(write_csv(tmp_path / "e2b.csv", rows))
    # another dotted cell in the same column
    rows = _e2_rows()
    rows[2][1] = "L1.300x"
    with pytest.raises(mp.InputError):
        mp.load_data_file(write_csv(tmp_path / "e2c.csv", rows))
    _, res = run_raw("ml_train.py", "-f", write_csv(tmp_path / "e2t.csv", _e2_rows()[:1] + [
        r if i != 3 else [r[0], r[1], "1.5"] for i, r in enumerate(_e2_rows()[1:])]), "-x", "온도,유량", "-y", "수율")
    assert res["status"] == "error" and res["agent_directive"].startswith("STOP_RETRY:"), res


def test_dot_grouped_integers_train_refused(tmp_path):
    _, res = run_raw("ml_train.py", "-f", write_csv(tmp_path / "e2t.csv", _e2_rows()), "-x", "온도,유량", "-y", "수율")
    assert res["status"] == "error" and res["agent_directive"].startswith("STOP_RETRY:"), res   # 2.9.21: ok


def test_tab_line_list_with_three_decimals_refused(tmp_path):
    # '1,2' (line list) confirms decimal commas; '2.125' is 2.125 or 2125 -> refused (2.9.21: 2125)
    rows = [["라인", "두께", "온도", "수율"]]
    for i in range(30):
        rows.append([["1,2", "3,4", "1"][i % 3], ["2.125", "1.375", "1.500", "1.125"][i % 4], 170 + i % 10, 80 + i % 7])
    with pytest.raises(mp.InputError, match="2.125"):
        mp.load_data_file(write_csv(tmp_path / "d1.csv", rows, sep="\t"))


def _euro_rows(n=24, seed=8):
    r = np.random.default_rng(seed)
    rows = [["온도", "유량", "수율"]]
    for _ in range(n):
        t, q = float(r.uniform(170, 180)), float(r.uniform(1100, 1500))
        rows.append([f"{t:.1f}".replace(".", ","), _euro(q), f"{0.05 * t + 0.01 * q + r.normal(0, 0.1):.2f}".replace(".", ",")])
    return rows


@pytest.fixture(scope="module")
def euro_model(tmp_path_factory):
    d = tmp_path_factory.mktemp("euro")
    _, res = run_raw("ml_train.py", "-f", write_csv(d / "train.csv", _euro_rows()), "-x", "온도,유량", "-y", "수율")
    assert res["status"] == "ok", res
    return d, res["version_dir"]


def test_model_keeps_decimal_comma_for_prediction(euro_model):
    d, vdir = euro_model
    import json
    with open(os.path.join(vdir, "model_수율.json"), encoding="utf-8") as f:
        assert json.load(f)["decimal_comma"] is True
    # no decimal comma in the prediction file: '1.300' would be 1.3 here and 1300 in training -> refused
    _, res = run_raw("ml_predict.py", "-m", vdir, "-f", write_csv(d / "pred.csv", [["온도", "유량"], [175, "1.300"],
                                                                            [180, "1.350"], [178, "1.320"]]))
    assert res["status"] == "error" and res["agent_directive"].startswith("STOP_RETRY:"), res   # 2.9.21: 1.3, ok
    assert "1.300" in res["message"]
    # the same numbers written the training way: read as 1300
    _, res = run_raw("ml_predict.py", "-m", vdir, "-f", write_csv(d / "pred_ok.csv", [["온도", "유량"], ["175", "1.300,0"],
                                                                               ["180", "1.350,0"]]))
    assert res["status"] == "ok", res


def test_kv_input_thousands_still_recommended():
    with pytest.raises(mp.InputError) as e:
        mp.check_comma_numbers({"온도": "1,500"}, lambda k: True, "-i", lambda p: "python x -i " + mp.kv_text(p))
    assert "[추천] python x -i 온도=1500" in str(e.value)
    with pytest.raises(mp.InputError) as e:
        mp.check_comma_numbers({"온도": "1,5"}, lambda k: True, "-i", lambda p: "x")
    assert "[추천]" not in str(e.value)


# ---- 숫자로 읽지 못한 칸 경고 (학습에 쓰는 숫자 열만) ------------------------------------------------------


def test_training_warns_unparsed_number_cells(tmp_path):
    r = random.Random(1)
    rows = [["온도", "압력", "수율"]]
    for i in range(40):
        t = 150 + i
        rows.append([("175,6" if i in (3, 9, 20) else t), round(r.uniform(2, 4), 2), 0.5 * t + r.uniform(-1, 1)])
    f = write_rows(str(tmp_path / "a.xlsx"), rows, [])
    _, res = run_raw("ml_train.py", "-f", f, "-x", "온도,압력", "-y", "수율")
    assert res["status"] == "ok", res
    assert any("숫자 열 '온도'" in w and "175,6" in w for w in res.get("warnings", [])), res.get("warnings")


def _number_warnings(res):
    return [w for w in res.get("warnings", []) if "숫자로 읽지 못해" in w or "소수점 쉼표" in w]


def test_skill_sample_unit_row_no_false_warning(tmp_path):
    import make_beta_samples
    make_beta_samples.main(str(tmp_path))
    data = tmp_path / "공정_다중출력.xlsx"      # header row 3, unit row 4 (auto-detected: no layout options)
    _, res = run_raw("ml_train.py", "-f", str(data), "-x", "반응온도,압축압력,촉매비율,수분함량,유량,분진",
                     "-y", "생산효율,불순물")
    assert res["status"] == "ok", res
    assert _number_warnings(res) == [], res.get("warnings")


def _train_rows(tmp_path, name, rows, *args):
    f = write_rows(str(tmp_path / name), rows, [])
    _, res = run_raw("ml_train.py", "-f", f, *args)
    return res


def test_category_code_column_not_warned(tmp_path):
    r = random.Random(3)
    rows = [["금형", "온도", "수율"]]
    for i in range(60):
        t = 150 + i % 30
        m = r.choice(["101", "102", "103", "104", "105", "106", "107", "A1", "A2", "B1"])
        rows.append([m, t, 0.5 * t + (5 if m.startswith("A") else 0) + r.uniform(-1, 1)])
    res = _train_rows(tmp_path, "u1.xlsx", rows, "-x", "금형,온도", "-y", "수율")
    assert res["status"] == "ok" and _number_warnings(res) == [], res


def test_excluded_and_target_missing_rows_not_warned(tmp_path):
    r = random.Random(3)
    rows = [["온도", "압력", "수율", "제외"]]
    for i in range(40):
        t, p, y, ex = 150 + i, round(r.uniform(2, 4), 2), 0.5 * (150 + i) + r.uniform(-1, 1), None
        if i in (5, 15):
            t, ex = "설비정지", "정지"
        if i == 25:
            t, y = "점검", None
        rows.append([t, p, y, ex])
    res = _train_rows(tmp_path, "u2.xlsx", rows, "-x", "온도,압력", "-y", "수율")
    assert res["status"] == "ok" and _number_warnings(res) == [], res


def test_forced_datetime_column_not_warned(tmp_path):
    r = random.Random(3)
    rows = [["일자", "온도", "수율"]]
    for i in range(40):
        d = datetime(2026, 1, 1 + i % 28)
        rows.append([(d.strftime("%Y-%m-%d") if i in (4, 8) else int(d.strftime("%Y%m%d"))), 150 + i,
                     0.5 * (150 + i) + r.uniform(-1, 1)])
    res = _train_rows(tmp_path, "u5.xlsx", rows, "-x", "일자,온도", "-y", "수율", "--dtypes", "일자:datetime")
    assert res["status"] == "ok" and _number_warnings(res) == [], res


def test_eval_source_column_warned(tmp_path):
    r = random.Random(3)
    rows = [["T1", "T2", "수율"]]
    for i in range(40):
        a, b = 150 + i, 100 + r.uniform(0, 10)
        rows.append([("175,6" if i in (3, 9, 20) else a), b, 0.5 * (a - b) + r.uniform(-1, 1)])
    res = _train_rows(tmp_path, "u3.xlsx", rows, "-x", "T2", "--eval", "dT=T1-T2", "-y", "수율")
    assert res["status"] == "ok", res
    assert any("'T1'의 3칸" in w and "175,6" in w for w in res["warnings"]), res["warnings"]


def test_one_warning_per_column_and_further_columns_counted(tmp_path):
    r = random.Random(5)
    cols = [f"x{k}" for k in range(7)]
    rows = [cols + ["y1", "y2"]]
    for i in range(40):
        vals = [round(r.uniform(0, 10), 2) for _ in cols]
        y = sum(vals)
        for k in range(7):
            if i == 3 + k:
                vals[k] = f"{k},5"
        rows.append(vals + [y, 2 * y])
    res = _train_rows(tmp_path, "many.xlsx", rows, "-x", ",".join(cols), "-y", "y1,y2")
    assert res["status"] == "ok", res
    named = [w for w in res["warnings"] if w.startswith("숫자 열")]
    assert len(named) == 5 and len(set(named)) == 5, res["warnings"]
    assert all("의 1칸" in w for w in named), named
    assert any("2개 더" in w and "'x5'" in w and "'x6'" in w for w in res["warnings"]), res["warnings"]


def _textdc_rows(share):
    r = random.Random(4)
    rows = [["온도", "압력", "두께", "수율"]]
    for i in range(40):
        t, th = 150 + i, round(r.uniform(0.1, 0.3), 3)
        cell = f"{th:.3f}".replace(".", ",") if (i % 10) < share * 10 else th
        rows.append([t, round(r.uniform(2, 4), 2), cell, 0.5 * t + 100 * th + r.uniform(-1, 1)])
    return rows


def test_decimal_comma_text_feature_refused(tmp_path):
    res = _train_rows(tmp_path, "c.xlsx", _textdc_rows(1.0), "-x", "온도,압력,두께", "-y", "수율")
    assert res["status"] == "error" and res["agent_directive"].startswith("STOP_RETRY"), res
    assert "'두께'" in res["message"] and "0," in res["message"] and "[추천]" not in res["message"], res["message"]


def test_decimal_comma_text_minority_warned(tmp_path):
    res = _train_rows(tmp_path, "c4.xlsx", _textdc_rows(0.4), "-x", "온도,압력,두께", "-y", "수율")
    assert res["status"] == "ok", res
    assert any("피처 '두께'" in w and "소수점 쉼표" in w for w in res["warnings"]), res["warnings"]


def test_decimal_comma_text_forced_category_allowed(tmp_path):
    res = _train_rows(tmp_path, "c5.xlsx", _textdc_rows(1.0), "-x", "온도,압력,두께", "-y", "수율",
                      "--dtypes", "두께:str")
    assert res["status"] == "ok", res


# ---- '%'가 붙은 숫자 --------------------------------------------------------------------------------------------


def test_percent_training_and_prediction_match_plain(name_files):
    proc, res = run_raw("ml_train.py", "-f", name_files["pct"], "-x", "온도,압력,습도(%)", "-y", "수율(%)", "--seed", "1")
    res = ok_res(res, proc)
    w = " ".join(res.get("warnings", []))
    assert "'습도(%)' 열의 60칸" in w and "'수율(%)' 열의 60칸" in w and "100으로 나누지 않음" in w, w
    proc, plain = run_raw("ml_train.py", "-f", name_files["plain"], "-x", "온도,압력,습도(%)", "-y", "수율(%)", "--seed", "1")
    plain = ok_res(plain, proc)
    assert res["mean_val_score"] == plain["mean_val_score"]
    m1 = os.path.join(str(name_files["tmp"]), "pct_ml_data", "v1")
    m2 = os.path.join(str(name_files["tmp"]), "plain_ml_data", "v1")
    proc, a = run_raw("ml_predict.py", "-m", m1, "-i", "온도=180,압력=3,습도=55.5%")
    a = ok_res(a, proc)
    assert any("%를 떼고" in x for x in a.get("warnings", [])), a
    proc, b = run_raw("ml_predict.py", "-m", m2, "-i", "온도=180,압력=3,습도=55.5")
    b = ok_res(b, proc)
    assert a["predictions_summary"] == b["predictions_summary"]


def test_percent_map_warns(name_files):
    proc, res = run_raw("ml_map.py", "-f", name_files["pct"], "-i", "온도=180,압력=3", "-y", "수율(%)")
    res = ok_res(res, proc)
    assert any("'수율(%)' 열의" in w for w in res.get("warnings", [])), res


def test_inline_percent_value_reads_like_data(percent_files):
    proc, a = run_raw("ml_map.py", "-f", percent_files["grid_pct"], "-i", "온도=160,함수율=10%", "-y", "수율")
    a = ok_res(a, proc)
    proc, b = run_raw("ml_map.py", "-f", percent_files["grid"], "-i", "온도=160,함수율=10", "-y", "수율")
    b = ok_res(b, proc)
    assert a["points"][0]["values"] == b["points"][0]["values"]
    assert any("-i '함수율'" in w and "10% -> 10" in w for w in a["warnings"]), a
    assert "warnings" in a["agent_directive"]


def test_reverse_inline_percent_value(percent_files):
    proc, res = run_raw("ml_map.py", "-f", percent_files["grid"], "--find", "수율=20", "--vary", "온도", "-i", "함수율=10%")
    res = ok_res(res, proc)
    assert res["fixed"] == {"함수율": 10.0} and any("-i '함수율'" in w for w in res["warnings"]), res


def test_points_percent_value_reads_like_data(percent_files):
    pa = data_file(percent_files, "p_pct.csv", ["온도", "함수율"], [[160, "12%"], [170, "8%"]])
    pb = data_file(percent_files, "p_plain.csv", ["온도", "함수율"], [[160, 12], [170, 8]])
    proc, a = run_raw("ml_map.py", "-f", percent_files["grid_pct"], "--points", pa, "-y", "수율", "--save",
                      data_path(percent_files, "o_pa.xlsx"), "--overwrite")
    a = ok_res(a, proc)
    proc, b = run_raw("ml_map.py", "-f", percent_files["grid"], "--points", pb, "-y", "수율", "--save",
                      data_path(percent_files, "o_pb.xlsx"), "--overwrite")
    ok_res(b, proc)
    ra = pd.read_excel(a["output_file"], sheet_name="결과")
    rb = pd.read_excel(b["output_file"], sheet_name="결과")
    assert list(ra["보간_수율"]) == list(rb["보간_수율"])
    assert any(w.startswith("조회 파일 '함수율'") for w in a["warnings"]), a


def test_check_percent_mix_unit():
    df = pd.DataFrame({"a": ["11.9%", "0.119", None], "b": ["11.9%", "12%", "-"], "c": [1.0, 2.0, 3.0],
                       "d": ["A", "12%", "B"]})
    with pytest.raises(mp.InputError) as e:
        mp.check_percent_mix(df, ["a"])
    assert "'a'" in str(e.value) and "'11.9%'" in str(e.value) and "'0.119'" in str(e.value)
    mp.check_percent_mix(df, ["b", "c", "d", "없는열"])


def test_train_mixed_feature_stops(percent_files):
    proc, res = run_raw("ml_train.py", "-f", percent_files["mix"], "-x", "온도,함수율", "-y", "수율")
    msg = stop_msg(res, proc)
    assert "'함수율'" in msg and "%" in msg, msg
    assert not os.path.isdir(os.path.join(str(percent_files["tmp"]), "mix_ml_data", "v1"))


def test_train_mixed_target_stops(percent_files):
    proc, res = run_raw("ml_train.py", "-f", percent_files["mix_y"], "-x", "온도,함수율", "-y", "수율")
    msg = stop_msg(res, proc)
    assert "'수율'" in msg, msg


def test_predict_mixed_input_file_stops(percent_files):
    proc, res = run_raw("ml_train.py", "-f", percent_files["moist"], "-x", "온도,함수율", "-y", "수율", "--seed", "1")
    ok_res(res, proc)
    model = os.path.join(str(percent_files["tmp"]), "moist_ml_data", "v1")
    new = data_file(percent_files, "new_mix.csv", ["온도", "함수율"], [[170, "12%"], [180, 0.08]])
    proc, res = run_raw("ml_predict.py", "-m", model, "-f", new)
    msg = stop_msg(res, proc)
    assert "'함수율'" in msg, msg
    same = data_file(percent_files, "new_pct.csv", ["온도", "함수율"], [[170, "12%"], [180, "8%"]])
    proc, res = run_raw("ml_predict.py", "-m", model, "-f", same)
    ok_res(res, proc)


def test_map_mixed_data_and_points_stop(percent_files):
    proc, res = run_raw("ml_map.py", "-f", percent_files["mix"], "-i", "온도=170", "-y", "함수율")
    assert "'함수율'" in stop_msg(res, proc)
    pts = data_file(percent_files, "p_mix.csv", ["온도", "함수율"], [[160, "12%"], [170, 8]])
    proc, res = run_raw("ml_map.py", "-f", percent_files["grid"], "--points", pts, "-y", "수율", "--save", data_path(percent_files, "o_x"))
    assert "'함수율'" in stop_msg(res, proc)


def test_train_mixed_percent_in_excluded_row(near_name_files):
    proc, res = run_raw("ml_train.py", "-f", near_name_files["excl"], "-x", "온도,함수율", "-y", "수율")
    ok_res(res, proc)


def test_train_mixed_percent_in_included_row_stops(near_name_files):
    rows = [["온도", "함수율", "수율", "제외"]] + [[150 + i, f"{i}%" if i == 3 else i, 40 - i, "포함" if i == 3 else ""]
                                                  for i in range(1, 30)]
    path = data_file(near_name_files, "excl_in.csv", rows[0], rows[1:])
    proc, res = run_raw("ml_train.py", "-f", path, "-x", "온도,함수율", "-y", "수율")
    assert "'함수율'" in stop_msg(res, proc)
