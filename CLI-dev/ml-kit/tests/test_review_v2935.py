# -*- coding: utf-8 -*-
"""
2.9.35 (24차 확인 검사) regressions.

1. NEW-1 (2.9.34 regression): ml_map --points stopped a measured / memo column ('수율_실측', '목표수율' with -y 수율)
   with [후보] '함수율' when the data had an unused 함수율 column: the two names share only the -y name 수율.
   Both names are compared without the -y name; only names still near ('함수율%' / '함수율') stop.
2. NEW-2 (2.9.34 regression): compile_formula refused a formula naming a number-named column by its full name
   ('0.5(mm)*2(mm)', '온도+2(mm)'). LOW-1: a number inside a larger number ('2.5', '0.5') is no column '2(mm)' /
   '5(mm)'.
3. LOW-2: training stopped on a '%' cell next to plain numbers in a row the user marked in '제외'.
"""
import os
import random
import sys

import pytest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from test_layout_matrix_v2917 import mp, _write  # noqa: E402
from test_cli_smoke import run_raw  # noqa: E402


def _stop(res, proc):
    assert proc.returncode == 1 and res["agent_directive"].startswith("STOP_RETRY:"), res
    return res["message"]


def _ok(res, proc):
    assert proc.returncode == 0 and res["status"] == "ok", res
    return res


@pytest.fixture(scope="module")
def data(tmp_path_factory):
    tmp = tmp_path_factory.mktemp("v2935")
    r = random.Random(7)
    grid = [["온도", "함수율", "수율", "생산량"]]                 # mapp shape plus 생산량
    cool = [["시간", "냉각수온도", "온도"]]
    sieve = [["온도", "0.5(mm)", "2(mm)", "수율"]]
    excl = [["온도", "함수율", "수율", "제외"]]
    for t in (150, 170, 190):
        for w in (5, 10, 15):
            grid.append([t, w, 0.2 * t - w, 2 * t + w])
            cool.append([t, w + 15, 0.2 * t - w])
    for i in range(40):
        t, w = round(r.uniform(150, 200), 1), round(r.uniform(5, 15), 1)
        a, b = round(r.uniform(10, 30), 2), round(r.uniform(1, 30), 2)
        sieve.append([t, a, b, round(t + a + b + r.uniform(-1, 1), 2)])
        y = round(0.3 * t - w + r.uniform(-1, 1), 2)
        excl.append([t, f"{w}%" if i == 4 else w, y, "측정 오류" if i == 4 else ""])
    out = {"tmp": tmp}
    for name, rows in (("grid", grid), ("cool", cool), ("sieve", sieve), ("excl", excl)):
        out[name] = _write(os.path.join(str(tmp), f"{name}.csv"), rows, [])
    return out


def _file(data, name, head, rows):
    return _write(os.path.join(str(data["tmp"]), name), [head] + rows, [])


def _save(data, name):
    return os.path.join(str(data["tmp"]), name)


# ---------------------------------------------------------------------------------------------------------------
# 1. NEW-1: names that share only the -y name are not near
# ---------------------------------------------------------------------------------------------------------------

@pytest.mark.parametrize("col,target", [("수율_실측", "수율"), ("목표수율", "수율"), ("수율_실측", "수율,생산량"),
                                        ("목표수율", "수율,생산량"), ("생산량_실측", "수율,생산량")])
def test_points_measured_column_passes_next_to_moisture(data, col, target):
    pts = _file(data, f"p_{col}.csv", ["온도", col], [[160, 30], [180, 20]])
    proc, res = run_raw("ml_map.py", "-f", data["grid"], "--points", pts, "-y", target,
                        "--save", _save(data, f"o_{col}_{len(target)}"))
    res = _ok(res, proc)
    assert res["agent_directive"].startswith("VERIFIED:"), res


@pytest.mark.parametrize("col", ["함수율%", "함수 율(%)", "함수율_수율", "함수율 수율", "수율실측함수율"])
def test_points_moisture_like_column_stops(data, col):
    pts = _file(data, f"p_s{len(col)}.csv", ["온도", col], [[160, 8], [180, 12]])
    proc, res = run_raw("ml_map.py", "-f", data["grid"], "--points", pts, "-y", "수율",
                        "--save", _save(data, "o_s"))
    msg = _stop(res, proc)
    assert "[후보] '함수율'" in msg and "[추천]" not in msg, msg


def test_points_cooling_key_still_stops(data):
    pts = _file(data, "p_cool.csv", ["시간", "냉각수 온도(°C)"], [[160, 25]])
    proc, res = run_raw("ml_map.py", "-f", data["cool"], "--points", pts, "-y", "온도", "--save", _save(data, "o_c"))
    assert "[후보] '냉각수온도'" in _stop(res, proc)


# ---------------------------------------------------------------------------------------------------------------
# 2. NEW-2 / LOW-1: number-named columns in --eval
# ---------------------------------------------------------------------------------------------------------------

SIEVE = ["온도", "0.5(mm)", "2(mm)", "수율"]


@pytest.mark.parametrize("formula,cols", [("0.5(mm)*2(mm)", ["0.5(mm)", "2(mm)"]), ("온도+2(mm)", ["온도", "2(mm)"]),
                                          ("2(mm)*2.5", ["2(mm)"])])
def test_full_number_name_compiles(formula, cols):
    tok, t2c, canon = mp.compile_formula(formula, SIEVE)
    assert sorted(t2c.values()) == sorted(cols) and canon == formula, (tok, t2c)


@pytest.mark.parametrize("formula,columns", [("온도*2.5", SIEVE), ("온도*0.5", ["온도", "5(mm)"]),
                                             ("온도*2.5", ["온도", "2"])])
def test_number_inside_number_is_no_column(formula, columns):
    tok, t2c, canon = mp.compile_formula(formula, columns)
    assert list(t2c.values()) == ["온도"] and canon == formula, (tok, t2c)


@pytest.mark.parametrize("formula,columns,col", [("온도*2", SIEVE, "2(mm)"), ("온도*2", ["온도", "2"], "2"),
                                                 ("온도*0.5", SIEVE, "0.5(mm)")])
def test_bare_number_of_column_still_stops(formula, columns, col):
    with pytest.raises(mp.InputError) as e:
        mp.compile_formula(formula, columns)
    assert f"'{col}'" in str(e.value)


def test_train_full_number_names(data):
    proc, res = run_raw("ml_train.py", "-f", data["sieve"], "-x", "온도", "-y", "수율",
                        "--eval", "합=0.5(mm)*2(mm);차=온도+2(mm)")
    _ok(res, proc)


# ---------------------------------------------------------------------------------------------------------------
# 3. LOW-2: a '%' cell in a row marked in '제외' does not stop training
# ---------------------------------------------------------------------------------------------------------------

def test_train_mixed_percent_in_excluded_row(data):
    proc, res = run_raw("ml_train.py", "-f", data["excl"], "-x", "온도,함수율", "-y", "수율")
    _ok(res, proc)


def test_train_mixed_percent_in_included_row_stops(data):
    rows = [["온도", "함수율", "수율", "제외"]] + [[150 + i, f"{i}%" if i == 3 else i, 40 - i, "포함" if i == 3 else ""]
                                                  for i in range(1, 30)]
    path = _file(data, "excl_in.csv", rows[0], rows[1:])
    proc, res = run_raw("ml_train.py", "-f", path, "-x", "온도,함수율", "-y", "수율")
    assert "'함수율'" in _stop(res, proc)
