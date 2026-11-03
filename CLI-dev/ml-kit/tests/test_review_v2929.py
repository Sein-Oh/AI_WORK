# -*- coding: utf-8 -*-
"""
2.9.29 (18차 점검) regressions.

1. A name with another unit got a rename [추천] when the unit-less base names differed ('압력(MPa)' for the model
   input '반응압력(bar)'): the retry read the MPa value as bar and returned VERIFIED. Two names that both carry a
   unit and differ in it never get [추천] (ml_predict -i / -f columns, ml_map -i / --find / --points, --find -i).
2. [후보] of resolve_column_names (ml_map -i, --find -i) did not mark the candidates with another unit.
3. ml_predict --find '생산량(kg)=500' for the target '생산량(t)' solved for 500 t.
4. (2.9.28) A constant column equals the mean of any block: constant setpoint columns plus one chance 0/1 sum made
   category rows ('A계') warned as totals rows. A mean cell counts only in a column that varies in the block.
5. ml_predict --find used only the first object of a JSON array for -i and the goal.
"""
import os
import random
import sys

import pytest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
sys.path.insert(0, os.path.join(os.path.dirname(HERE), "tools"))
from test_layout_matrix_v2917 import mp, _write  # noqa: E402
from test_cli_smoke import run_raw  # noqa: E402


def _data(cols, target, n=60, seed=1):
    """Rows of a process file: 온도(°C), a pressure column of the given name, 속도, and the target."""
    r = random.Random(seed)
    rows = [cols + [target]]
    for _ in range(n):
        t, p, s = round(r.uniform(170, 190), 1), round(r.uniform(2, 4), 2), round(r.uniform(40, 50), 1)
        v = {"온도(°C)": t, "속도": s}
        for c in cols:
            if "압력" in c:
                v[c] = p if c not in v else v[c]
        if "압력차(bar)" in cols:
            v["압력차(bar)"] = round(p * 0.3 + r.uniform(0, 0.2), 2)
        rows.append([v[c] for c in cols] + [round(2 * t + 30 * p + 0.5 * s + r.uniform(-2, 2), 1)])
    return rows


def _train(tmp, name, cols, target="생산량"):
    f = _write(os.path.join(str(tmp), name), _data(cols, target), [])
    _, res = run_raw("ml_train.py", "-f", f, "-x", ",".join(cols), "-y", mp.strip_unit(target))
    assert res["status"] == "ok", res
    return os.path.join(str(tmp), os.path.splitext(name)[0] + "_ml_data", "v1"), f


@pytest.fixture(scope="module")
def models(tmp_path_factory):
    tmp = tmp_path_factory.mktemp("v2929")
    out = {"tmp": tmp}
    out["p4"], out["p4_f"] = _train(tmp, "p4.csv", ["온도(°C)", "반응압력(bar)", "속도"])
    out["p3"], out["p3_f"] = _train(tmp, "p3.csv", ["온도(°C)", "압력(bar)", "압력차(bar)", "속도"])
    out["p2"], _ = _train(tmp, "p2.csv", ["온도(°C)", "압력(bar)", "속도"], target="생산량(t)")
    out["p1"], _ = _train(tmp, "p1.csv", ["온도(°C)", "압력(bar)", "속도"])
    return out


def _stop(res, proc):
    assert proc.returncode == 1 and res["agent_directive"].startswith("STOP_RETRY:"), res
    return res["message"]


# ---------------------------------------------------------------------------------------------------------------
# 1. other unit, different base name: no [추천]
# ---------------------------------------------------------------------------------------------------------------

def test_other_unit_with_other_base_name_predict(models):
    proc, res = run_raw("ml_predict.py", "-m", models["p4"], "-i", "압력(MPa)=0.3,온도=180,속도=45")
    msg = _stop(res, proc)
    assert "[추천]" not in msg and "단위가 다릅니다" in msg and "bar 단위로 환산" in msg, msg
    # file input column: same rule
    f = _write(os.path.join(str(models["tmp"]), "in_mpa.csv"), [["압력(MPa)", "온도(°C)", "속도"], [0.3, 180, 45]], [])
    proc, res = run_raw("ml_predict.py", "-m", models["p4"], "-f", f)
    msg = _stop(res, proc)
    assert "[추천]" not in msg and "단위가 다릅니다" in msg, msg


def test_other_unit_with_other_base_name_ml_map(models):
    proc, res = run_raw("ml_map.py", "-f", models["p4_f"], "-i", "압력(MPa)=0.3", "-y", "생산량")
    msg = _stop(res, proc)
    assert "[추천]" not in msg and "단위가 다릅니다" in msg, msg
    proc, res = run_raw("ml_map.py", "-f", models["p4_f"], "--find", "생산량=530", "--vary", "온도",
                        "-i", "압력(MPa)=0.3")
    msg = _stop(res, proc)
    assert "[추천]" not in msg and "단위가 다릅니다" in msg, msg
    # --points: a points column with the base name of a data column and another unit (other base names are
    # passed through as extra columns of the points file)
    pts = _write(os.path.join(str(models["tmp"]), "pts.csv"), [["압력(MPa)", "온도"], [0.3, 180]], [])
    proc, res = run_raw("ml_map.py", "-f", models["p3_f"], "--points", pts, "-y", "생산량")
    msg = _stop(res, proc)
    assert "[추천]" not in msg and "'압력(bar)'(단위 다름: 값 환산 필요)" in msg, msg


def test_other_unit_with_other_base_name_find(models):
    proc, res = run_raw("ml_predict.py", "-m", models["p4"], "--find", "생산량=530", "--vary", "온도",
                        "-i", "압력(MPa)=0.3")
    msg = _stop(res, proc)
    assert "[추천]" not in msg and "단위가 다릅니다" in msg, msg


def test_same_unit_typo_still_suggested(models):
    proc, res = run_raw("ml_predict.py", "-m", models["p4"], "-i", "압력(bar)=3,온도=180,속도=45")
    assert "[추천] '반응압력(bar)'" in _stop(res, proc)
    proc, res = run_raw("ml_map.py", "-f", models["p4_f"], "-i", "반응압룍(bar)=3", "-y", "생산량")
    assert "[추천] '반응압력(bar)'" in _stop(res, proc)
    # 2.9.31: the same unit spaced otherwise is the same name (resolve_column_names), no [추천] (was the test above)
    proc, res = run_raw("ml_map.py", "-f", models["p4_f"], "-i", "반응압력( bar)=3", "-y", "생산량")
    assert proc.returncode == 0 and res["keys"] == ["반응압력(bar)"], res


def test_other_unit_rule():
    assert mp.other_unit("압력(MPa)", "반응압력(bar)") and mp.other_unit("반응압력(bar)", "압력(MPa)")
    assert not mp.other_unit("압력(bar)", "반응압력(bar)") and not mp.other_unit("압력(MPa)", "반응압력")
    assert not mp.other_unit("온도(C)", "반응온도(℃)")


# ---------------------------------------------------------------------------------------------------------------
# 2. [후보] marks
# ---------------------------------------------------------------------------------------------------------------

def test_candidates_marked(models):
    proc, res = run_raw("ml_map.py", "-f", models["p3_f"], "-i", "압력(MPa)=0.3", "-y", "생산량")
    msg = _stop(res, proc)
    assert "[후보] '압력(bar)'(단위 다름: 값 환산 필요), '압력차(bar)'(단위 다름: 값 환산 필요)" in msg, msg
    proc, res = run_raw("ml_predict.py", "-m", models["p3"], "--find", "생산량=530", "--vary", "온도",
                        "-i", "압력(MPa)=0.3")
    msg = _stop(res, proc)
    assert "'압력(bar)'(단위 다름: 값 환산 필요)" in msg and "'압력차(bar)'(단위 다름: 값 환산 필요)" in msg, msg
    # names without values (-y) are not marked; a same-unit candidate is not marked
    with pytest.raises(mp.InputError) as e:
        mp.resolve_column_names(["압력(MPa)"], ["압력(bar)", "압력차(bar)"], "구할 열(-y)")
    assert "단위 다름" not in str(e.value)
    with pytest.raises(mp.InputError) as e:
        mp.resolve_column_names(["압력(bar)"], ["압력차(bar)", "압력차2(bar)"], "기준(-i)", values=True)
    assert "[후보]" in str(e.value) and "단위 다름" not in str(e.value)


# ---------------------------------------------------------------------------------------------------------------
# 3. --find goal unit
# ---------------------------------------------------------------------------------------------------------------

def test_find_goal_other_unit_refused(models):
    proc, res = run_raw("ml_predict.py", "-m", models["p2"], "--find", "생산량(kg)=500", "--vary", "온도",
                        "-i", "압력=3")
    msg = _stop(res, proc)
    assert "목표(--find)" in msg and "단위가 다릅니다" in msg and "t 단위로 환산" in msg, msg
    for goal in ("생산량(t)=500", "생산량=500", "생산량( t )=500"):
        proc, res = run_raw("ml_predict.py", "-m", models["p2"], "--find", goal, "--vary", "온도", "-i", "압력=3")
        assert res["status"] == "ok", (goal, res)


# ---------------------------------------------------------------------------------------------------------------
# 4. constant columns are no mean evidence
# ---------------------------------------------------------------------------------------------------------------

def _const_rows(seed=3):
    r = random.Random(seed)
    flags = [0, 1, 1, 0, 0, 0, 1, 1, 0, 1, 0, 0, 1, 0, 0, 1]
    rows, k = [["구분", "설정온도", "설정압력", "불량", "생산량"]], 0
    for _ in range(8):
        for _ in range(2):
            rows.append(["공통", 180, 3.5, flags[k], 100 + r.randint(0, 30)])
            k += 1
        rows.append(["A계", 180, 3.5, r.choice([0, 1]), 100 + r.randint(0, 30)])
    return rows


@pytest.mark.parametrize("ext", ["csv", "xlsx"])
def test_constant_columns_category_quiet(tmp_path, ext):
    f = _write(str(tmp_path / f"c.{ext}"), _const_rows(), [])
    assert mp.load_data_file(f).attrs["summary_rows"] == []


def test_totals_with_varying_means_still_warned(tmp_path):
    r = random.Random(7)
    d = [[f"10/{i + 1}", round(r.uniform(175, 185), 1), round(r.uniform(2, 3), 2), round(r.uniform(40, 50), 1),
          r.randint(280, 320)] for i in range(30)]
    avg = [round(sum(x[c] for x in d) / len(d), 2) for c in (1, 2, 3)]
    rows = [["일자", "온도", "압력", "속도", "생산량"]] + d + [["총 합계"] + avg + [sum(x[4] for x in d)]]
    got = mp.load_data_file(_write(str(tmp_path / "t.xlsx"), rows, [])).attrs["summary_rows"]
    assert [g[2] for g in got] == ["총 합계"], got


def test_sums_match_varies():
    import numpy as np
    sums = np.array([[300.0, 30.0, 3000.0]])
    count = np.array([[10.0, 10.0, 10.0]])
    row = np.array([[30.0, 3.0, 3000.0]])
    assert mp._sums_match(row, sums, count, sums)[0]
    assert mp._sums_match(row, sums, count, sums, np.array([[True, True, True]]))[0]
    # the mean columns are constant in the block: only one sum of three cells
    assert not mp._sums_match(row, sums, count, sums, np.array([[False, False, True]]))[0]


# ---------------------------------------------------------------------------------------------------------------
# 5. --find JSON arrays
# ---------------------------------------------------------------------------------------------------------------

def test_find_multi_row_input_refused(models):
    proc, res = run_raw("ml_predict.py", "-m", models["p1"], "--find", "생산량=530", "--vary", "온도",
                        "-i", '[{"압력":2.2},{"압력":3.8}]')
    assert "한 줄" in _stop(res, proc)
    proc, res = run_raw("ml_predict.py", "-m", models["p1"], "--find", '[{"생산량":530},{"생산량":520}]',
                        "--vary", "온도")
    assert "한 줄" in _stop(res, proc)
    proc, res = run_raw("ml_predict.py", "-m", models["p1"], "--find", '[{"생산량":530}]', "--vary", "온도",
                        "-i", '[{"압력":2.2}]')
    assert res["status"] == "ok" and res["fixed"] == {"압력(bar)": 2.2}, res
