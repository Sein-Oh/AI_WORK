# -*- coding: utf-8 -*-
"""
2.9.34 (23차 점검) regressions.

1. R1 (2.9.33 regression): ml_map --points passed a column through as a measured / memo column whenever its name
   held the -y name ('냉각수 온도(°C)' with -y 온도, '함수율%' with -y 수율): a real key column was dropped and the
   result was VERIFIED but wrong. The near check comes first; only a column near no unused column passes through.
2. R2 (2.9.33 regression): ml_map -i '함수율=10%' and a --points value '12%' were refused (ml_predict reads them).
3. H1: a numeric column with '%' cells AND plain number cells ('11.9%', '0.119') trained silently. STOP_RETRY in
   training (feature, target), prediction (input file) and ml_map (data, --points).
4. H2: the unit-less alias of a number-named column ('2' of '2(mm)') silently took the number 2 of an --eval formula.
"""
import os
import random
import sys

import pandas as pd
import pytest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from test_layout_matrix_v2917 import mp, _write  # noqa: E402
from test_cli_smoke import run_raw  # noqa: E402

import ml_map  # noqa: E402  (ml-kit folder is on sys.path through test_layout_matrix_v2917)


def _stop(res, proc):
    assert proc.returncode == 1 and res["agent_directive"].startswith("STOP_RETRY:"), res
    return res["message"]


def _ok(res, proc):
    assert proc.returncode == 0 and res["status"] == "ok", res
    return res


@pytest.fixture(scope="module")
def data(tmp_path_factory):
    tmp = tmp_path_factory.mktemp("v2934")
    r = random.Random(11)
    cool = [["냉각수온도", "유량", "온도"]]                          # map3 shape: -y 온도 next to 냉각수온도
    moist = [["온도", "함수율", "수율"]]
    moist_u = [["반응온도(°C)", "함수율(%)", "수율(최종)"]]
    d = [["온도(°C)", "압력(bar)", "수율(1차)", "수율(최종)", "생산량"]]
    grid_pct = [["온도", "함수율", "수율"]]                         # mapp shape: all '%' cells
    grid = [["온도", "함수율", "수율"]]
    mix = [["온도", "함수율", "수율"]]                              # pctmix shape: '11.9%' next to '0.077'
    mix_y = [["온도", "함수율", "수율"]]
    sieve = [["온도", "0.5(mm)", "2(mm)", "수율"]]
    for i in range(40):
        t, w = round(r.uniform(150, 200), 1), round(r.uniform(5, 15), 1)
        c, f = round(r.uniform(15, 30), 2), round(r.uniform(1, 5), 3)
        cool.append([c, f, round(60 + 0.8 * c + 2 * f + r.uniform(-1, 1), 2)])
        y = round(0.3 * t - w + r.uniform(-1, 1), 2)
        moist.append([t, w, y])
        moist_u.append([t, w, y])
        p = round(r.uniform(1, 5), 3)
        y1 = round(0.3 * t + 5 * p + r.uniform(-1, 1), 2)
        d.append([t, p, y1, round(y1 + 2, 2), round(y1 * 10, 1)])
        mix.append([t, f"{w}%" if i % 2 else round(w / 100, 4), y])
        mix_y.append([t, w, f"{y}%" if i == 3 else y])
        a, b = round(r.uniform(10, 30), 2), round(r.uniform(1, 30), 2)
        sieve.append([t, a, b, round(t + a + b + r.uniform(-1, 1), 2)])
    for t in (150, 170, 190):
        for w in (5, 10, 15):
            grid_pct.append([t, f"{w}%", 0.2 * t - w])
            grid.append([t, w, 0.2 * t - w])
    out = {"tmp": tmp}
    for name, rows in (("cool", cool), ("moist", moist), ("moist_u", moist_u), ("d", d), ("grid_pct", grid_pct),
                       ("grid", grid), ("mix", mix), ("mix_y", mix_y), ("sieve", sieve)):
        out[name] = _write(os.path.join(str(tmp), f"{name}.csv"), rows, [])
    return out


def _file(data, name, head, rows):
    return _write(os.path.join(str(data["tmp"]), name), [head] + rows, [])


def _save(data, name):
    return os.path.join(str(data["tmp"]), name)


# ---------------------------------------------------------------------------------------------------------------
# 1. R1: a key column holding the -y name is near a data column: [후보], never dropped
# ---------------------------------------------------------------------------------------------------------------

def test_points_key_column_with_target_name_stops(data):
    pts = _file(data, "p_cool.csv", ["냉각수 온도(°C)", "유량"], [[20, 3], [28, 3]])
    proc, res = run_raw("ml_map.py", "-f", data["cool"], "--points", pts, "-y", "온도", "--save", _save(data, "o1"))
    msg = _stop(res, proc)
    assert "[후보] '냉각수온도'" in msg and "[추천]" not in msg, msg


@pytest.mark.parametrize("src,col,target,cand", [
    ("moist", "함수율%", "수율", "함수율"),
    ("moist_u", "함수율%", "수율(최종)", "함수율(%)"),
])
def test_points_moisture_column_with_target_name_stops(data, src, col, target, cand):
    pts = _file(data, f"p_{src}.csv", ["온도" if src == "moist" else "반응온도(°C)", col, "메모"],
                [[160, 6, "a"], [180, 14, "b"]])
    proc, res = run_raw("ml_map.py", "-f", data[src], "--points", pts, "-y", target, "--save", _save(data, "o2"))
    msg = _stop(res, proc)
    assert f"[후보] '{cand}'" in msg and "[추천]" not in msg, msg


@pytest.mark.parametrize("col,target", [("수율_실측", "수율(최종)"), ("목표수율", "수율(최종)"),
                                        ("생산량_실측", "생산량"), ("수율_실측", "생산량")])
def test_points_measured_column_still_passes(data, col, target):
    pts = _file(data, f"p_m_{col}_{target}.csv".replace("(", "").replace(")", ""), ["온도", "압력", col],
                [[170, 3, 60], [180, 2.5, 61]])
    out = _save(data, f"o_m_{col}.xlsx")
    proc, res = run_raw("ml_map.py", "-f", data["d"], "--points", pts, "-y", target, "--save", out, "--overwrite")
    res = _ok(res, proc)
    assert res["keys"] == ["온도(°C)", "압력(bar)"] and res["columns"] == [target], res
    assert col in pd.read_excel(res["output_file"], sheet_name="결과").columns


def test_check_near_points_both_directions():
    pts = pd.DataFrame({"냉각수 온도(°C)": [20.0], "유량": [3.0]})
    with pytest.raises(mp.ColumnNameError):
        ml_map._check_near_points(pts, ["유량"], ["냉각수온도"], {}, used=["유량", "온도"])
    ok = pd.DataFrame({"온도": [170.0], "수율_실측": [60.0]})
    ml_map._check_near_points(ok, ["온도"], ["압력(bar)", "수율(1차)"], {}, used=["온도(°C)", "수율(최종)"])


# ---------------------------------------------------------------------------------------------------------------
# 2. R2: '%' values of -i and --points
# ---------------------------------------------------------------------------------------------------------------

def test_inline_percent_value_reads_like_data(data):
    proc, a = run_raw("ml_map.py", "-f", data["grid_pct"], "-i", "온도=160,함수율=10%", "-y", "수율")
    a = _ok(a, proc)
    proc, b = run_raw("ml_map.py", "-f", data["grid"], "-i", "온도=160,함수율=10", "-y", "수율")
    b = _ok(b, proc)
    assert a["points"][0]["values"] == b["points"][0]["values"]
    assert any("-i '함수율'" in w and "10% -> 10" in w for w in a["warnings"]), a
    assert "warnings" in a["agent_directive"]


def test_reverse_inline_percent_value(data):
    proc, res = run_raw("ml_map.py", "-f", data["grid"], "--find", "수율=20", "--vary", "온도", "-i", "함수율=10%")
    res = _ok(res, proc)
    assert res["fixed"] == {"함수율": 10.0} and any("-i '함수율'" in w for w in res["warnings"]), res


def test_points_percent_value_reads_like_data(data):
    pa = _file(data, "p_pct.csv", ["온도", "함수율"], [[160, "12%"], [170, "8%"]])
    pb = _file(data, "p_plain.csv", ["온도", "함수율"], [[160, 12], [170, 8]])
    proc, a = run_raw("ml_map.py", "-f", data["grid_pct"], "--points", pa, "-y", "수율", "--save",
                      _save(data, "o_pa.xlsx"), "--overwrite")
    a = _ok(a, proc)
    proc, b = run_raw("ml_map.py", "-f", data["grid"], "--points", pb, "-y", "수율", "--save",
                      _save(data, "o_pb.xlsx"), "--overwrite")
    _ok(b, proc)
    ra = pd.read_excel(a["output_file"], sheet_name="결과")
    rb = pd.read_excel(b["output_file"], sheet_name="결과")
    assert list(ra["보간_수율"]) == list(rb["보간_수율"])
    assert any(w.startswith("조회 파일 '함수율'") for w in a["warnings"]), a


# ---------------------------------------------------------------------------------------------------------------
# 3. H1: '%' cells next to plain numbers stop
# ---------------------------------------------------------------------------------------------------------------

def test_check_percent_mix_unit():
    df = pd.DataFrame({"a": ["11.9%", "0.119", None], "b": ["11.9%", "12%", "-"], "c": [1.0, 2.0, 3.0],
                       "d": ["A", "12%", "B"]})
    with pytest.raises(mp.InputError) as e:
        mp.check_percent_mix(df, ["a"])
    assert "'a'" in str(e.value) and "'11.9%'" in str(e.value) and "'0.119'" in str(e.value)
    mp.check_percent_mix(df, ["b", "c", "d", "없는열"])


def test_train_mixed_feature_stops(data):
    proc, res = run_raw("ml_train.py", "-f", data["mix"], "-x", "온도,함수율", "-y", "수율")
    msg = _stop(res, proc)
    assert "'함수율'" in msg and "%" in msg, msg
    assert not os.path.isdir(os.path.join(str(data["tmp"]), "mix_ml_data", "v1"))


def test_train_mixed_target_stops(data):
    proc, res = run_raw("ml_train.py", "-f", data["mix_y"], "-x", "온도,함수율", "-y", "수율")
    msg = _stop(res, proc)
    assert "'수율'" in msg, msg


def test_predict_mixed_input_file_stops(data):
    proc, res = run_raw("ml_train.py", "-f", data["moist"], "-x", "온도,함수율", "-y", "수율", "--seed", "1")
    _ok(res, proc)
    model = os.path.join(str(data["tmp"]), "moist_ml_data", "v1")
    new = _file(data, "new_mix.csv", ["온도", "함수율"], [[170, "12%"], [180, 0.08]])
    proc, res = run_raw("ml_predict.py", "-m", model, "-f", new)
    msg = _stop(res, proc)
    assert "'함수율'" in msg, msg
    same = _file(data, "new_pct.csv", ["온도", "함수율"], [[170, "12%"], [180, "8%"]])
    proc, res = run_raw("ml_predict.py", "-m", model, "-f", same)
    _ok(res, proc)


def test_map_mixed_data_and_points_stop(data):
    proc, res = run_raw("ml_map.py", "-f", data["mix"], "-i", "온도=170", "-y", "함수율")
    assert "'함수율'" in _stop(res, proc)
    pts = _file(data, "p_mix.csv", ["온도", "함수율"], [[160, "12%"], [170, 8]])
    proc, res = run_raw("ml_map.py", "-f", data["grid"], "--points", pts, "-y", "수율", "--save", _save(data, "o_x"))
    assert "'함수율'" in _stop(res, proc)


# ---------------------------------------------------------------------------------------------------------------
# 4. H2: an alias that reads as a number
# ---------------------------------------------------------------------------------------------------------------

@pytest.mark.parametrize("formula,col", [("온도*2", "2(mm)"), ("온도*0.5", "0.5(mm)")])   # '온도+2.5': 2.9.35 compiles
def test_number_alias_in_formula_stops(formula, col):
    with pytest.raises(mp.InputError) as e:
        mp.compile_formula(formula, ["온도", "0.5(mm)", "2(mm)", "수율"])
    assert f"'{col}'" in str(e.value)


def test_number_alias_never_a_formula_name():
    tok, cols, canon = mp.compile_formula("온도*3", ["온도", "2(mm)", "수율"])
    assert list(cols.values()) == ["온도"] and canon == "온도*3"


def test_train_number_alias_formula_stops(data):
    proc, res = run_raw("ml_train.py", "-f", data["sieve"], "-x", "온도,0.5(mm)", "-y", "수율", "--eval", "a=온도*2")
    assert "'2(mm)'" in _stop(res, proc)
