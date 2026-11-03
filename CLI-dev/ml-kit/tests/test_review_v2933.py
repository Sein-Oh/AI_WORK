# -*- coding: utf-8 -*-
"""
2.9.33 (22차 점검) regressions.

0. Design: fuzzy name suggestions never give [추천] (an agent retries it by itself, so '속도' -> '가속도' or
   '수율' -> '함수율(%)' put the value into another quantity). They are [후보] (ask the user), even for one candidate.
   Only deterministic same-column equivalences match, automatically with [INFO]: case, spacing, bracket shape and
   '°C' / '℃' / 'C'. sure_typo is gone.
1. R1 (2.9.32 regression): ml_map --points refused measured / memo columns ('수율 실측', '수율_실측', '목표수율') when
   the -y target has a sibling ('수율(1차)' next to -y '수율(최종)'), also with -y 생산량.
2. R2 (2.9.32 regression): the "-y가 없어 ..." guidance was added to unrelated --points errors.
3. H2: a column named like a number ('2') silently took the number 2 of an --eval formula.
4. M1: numbers written with '%' ('64.4%') are read as 64.4 with a warning, in training, prediction and ml_map.
5. L1: ml_export --inputs of constant inputs: [후보], no [추천]. L2: an excluded column is never a [후보].
6. --best --task of a task without history: [후보], no [추천].
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

import ml_export  # noqa: E402  (ml-kit folder is on sys.path through test_layout_matrix_v2917)
import ml_train  # noqa: E402


def _stop(res, proc):
    assert proc.returncode == 1 and res["agent_directive"].startswith("STOP_RETRY:"), res
    return res["message"]


def _ok(res, proc):
    assert proc.returncode == 0 and res["status"] == "ok", res
    return res


def _resolve_error(names, cols, **kw):
    with pytest.raises(mp.InputError) as e:
        mp.resolve_column_names(names, cols, "피처(-x)", **kw)
    return str(e.value)


@pytest.fixture(scope="module")
def data(tmp_path_factory):
    tmp = tmp_path_factory.mktemp("v2933")
    r = random.Random(7)
    d = [["온도(°C)", "압력(bar)", "속도", "습도(%)", "수율(1차)", "수율(최종)", "생산량"]]
    pct = [["온도(°C)", "압력(bar)", "습도(%)", "수율(%)"]]
    plain = [["온도(°C)", "압력(bar)", "습도(%)", "수율(%)"]]
    acc = [["가속도", "압력(bar)", "함수율(%)", "진동"]]
    num = [["온도", "2", "수율"]]
    const = [["온도", "압력", "라인", "수율"]]
    for _ in range(60):
        t, p, s, h = (round(r.uniform(150, 200), 2), round(r.uniform(1, 5), 3), round(r.uniform(10, 30), 2),
                      round(r.uniform(30, 70), 1))
        y1 = round(0.3 * t + 5 * p + r.uniform(-1, 1), 2)
        y2 = round(y1 + s / 5 + r.uniform(-0.5, 0.5), 2)
        d.append([t, p, s, h, y1, y2, round(y2 * 10, 1)])
        yp = round(0.2 * t + 0.3 * h + r.uniform(-1, 1), 2)
        pct.append([t, p, f"{h}%", f"{yp}%"])
        plain.append([t, p, h, yp])
        a, w = round(r.uniform(1, 4), 2), round(r.uniform(5, 12), 2)
        acc.append([a, p, w, round(5 * a + w + r.uniform(-1, 1), 2)])
        num.append([t, round(r.uniform(0, 9), 2), y1])
        const.append([t, p, 5, y1])
    out = {"tmp": tmp}
    for name, rows in (("d", d), ("pct", pct), ("plain", plain), ("acc", acc), ("num", num), ("const", const)):
        out[name] = _write(os.path.join(str(tmp), f"{name}.csv"), rows, [])
    return out


def _points(data, name, head, rows):
    return _write(os.path.join(str(data["tmp"]), name), [head] + rows, [])


# ---------------------------------------------------------------------------------------------------------------
# 0. fuzzy names give [후보] only, deterministic equivalences match
# ---------------------------------------------------------------------------------------------------------------

def test_sure_typo_is_gone():
    assert not hasattr(mp, "sure_typo") and not hasattr(mp, "one_or_ask")


@pytest.mark.parametrize("given,cols", [
    ("속도", ["가속도", "압력(bar)"]), ("수율", ["함수율(%)", "압력(bar)"]), ("속두", ["속도", "압력"]),
    ("압룍", ["압력(bar)", "온도"]), ("온도x", ["온도(°C)", "압력"]), ("tmep", ["temp", "press"]),
    ("압력", ["반응압력(bar)", "온도"]), ("생산량(kg)", ["생산량", "온도"]), ("온돜", ["온도", "압력"]),
])
def test_fuzzy_name_is_a_candidate(given, cols):
    msg = _resolve_error([given], cols)
    assert "[추천]" not in msg and f"[후보] '{cols[0]}'" in msg and "사용자에게 물으십시오" in msg, msg


@pytest.mark.parametrize("given,col", [
    ("온도(C)", "온도(°C)"), ("온도(℃)", "온도(°C)"), ("TEMP(°C)", "temp(℃)"), ("압력 [bar]", "압력(bar)"),
    ("수율( 1차)", "수율(1차)"), ("반응 온도(C)", "반응온도(°C)"), ("생산효율", "생산효율(%)"),
])
def test_same_column_written_otherwise_matches(given, col, capsys):
    got, applied = mp.resolve_column_names([given], [col, "압력(MPa)"], "피처(-x)")
    assert got == [col] and applied == {given: col}
    assert "[INFO] 컬럼명 자동 매칭" in capsys.readouterr().err


def test_equivalence_is_not_unit_case_or_ambiguous():
    msg = _resolve_error(["질량(mg)"], ["질량(Mg)"])                  # unit case counts: another unit
    assert "[추천]" not in msg and "[후보]" in msg, msg
    msg = _resolve_error(["온도(C)"], ["온도(°C)", "온도(℃)"])         # two equivalent columns: the user picks
    assert "[추천]" not in msg and "[후보] '온도(°C)', '온도(℃)'" in msg, msg


def test_train_accel_and_moisture_are_candidates(data):
    """H1: '속도' -> [추천] '가속도' and '수율' -> [추천] '함수율(%)' (2.9.32) trained another quantity on the retry."""
    proc, res = run_raw("ml_train.py", "-f", data["acc"], "-x", "속도,압력", "-y", "진동")
    msg = _stop(res, proc)
    assert "[추천]" not in msg and "[후보] '가속도'" in msg, msg
    proc, res = run_raw("ml_train.py", "-f", data["acc"], "-x", "가속도,압력", "-y", "수율")
    msg = _stop(res, proc)
    assert "[추천]" not in msg and "[후보] '함수율(%)'" in msg, msg


def test_degree_spelling_matches_in_train(data):
    proc, res = run_raw("ml_train.py", "-f", data["d"], "-x", "온도(C),압력(bar)", "-y", "수율(1차)")
    _ok(res, proc)


def test_map_inline_typo_is_a_candidate(data):
    proc, res = run_raw("ml_map.py", "-f", data["d"], "-i", "온도=180,압룍=3", "-y", "생산량")
    msg = _stop(res, proc)
    assert "[추천]" not in msg and "[후보] '압력(bar)'" in msg, msg


# ---------------------------------------------------------------------------------------------------------------
# 1. R1: measured / memo columns of the -y target pass through
# ---------------------------------------------------------------------------------------------------------------

@pytest.mark.parametrize("col", ["수율 실측", "수율_실측", "목표수율", "수율목표"])
@pytest.mark.parametrize("target", ["수율(최종)", "생산량"])
def test_points_measured_column_passes(data, col, target):
    pts = _points(data, f"p_{col}_{target}.csv".replace(" ", "_").replace("(", "").replace(")", ""),
                  ["온도", "압력", col], [[170, 3, 60], [180, 2.5, 61]])
    proc, res = run_raw("ml_map.py", "-f", data["d"], "--points", pts, "-y", target)
    res = _ok(res, proc)
    assert res["keys"] == ["온도(°C)", "압력(bar)"] and res["columns"] == [target], res


def test_points_near_key_name_still_stops(data):
    pts = _points(data, "p_near.csv", ["온도", "압력", "솎도"], [[170, 3, 20]])
    proc, res = run_raw("ml_map.py", "-f", data["d"], "--points", pts, "-y", "생산량")
    msg = _stop(res, proc)
    assert "[추천]" not in msg and "[후보] '속도'" in msg, msg


# ---------------------------------------------------------------------------------------------------------------
# 2. R2: the -y guidance only on the unknown-column error
# ---------------------------------------------------------------------------------------------------------------

def test_points_without_y_guidance_only_for_names(data):
    bad = _points(data, "p_bad.csv", ["온도", "압력"], [["1:30", 3]])
    proc, res = run_raw("ml_map.py", "-f", data["d"], "--points", bad)
    msg = _stop(res, proc)
    assert "숫자가 아닙니다" in msg and "-y가 없어" not in msg, msg
    dup = _points(data, "p_dup.csv", ["온도", "온도(°C)", "압력"], [[170, 170, 3]])
    proc, res = run_raw("ml_map.py", "-f", data["d"], "--points", dup)
    msg = _stop(res, proc)
    assert "같은 열" in msg and "-y가 없어" not in msg, msg
    name = _points(data, "p_name.csv", ["온도", "압력", "수율"], [[170, 3, 60]])
    proc, res = run_raw("ml_map.py", "-f", data["d"], "--points", name)
    msg = _stop(res, proc)
    assert "[후보]" in msg and "-y가 없어" in msg, msg


# ---------------------------------------------------------------------------------------------------------------
# 3. H2: a column named like a number
# ---------------------------------------------------------------------------------------------------------------

def test_number_named_column_in_formula_stops(data):
    with pytest.raises(mp.InputError) as e:
        mp.compile_formula("온도*2", ["온도", "2", "수율"])
    assert "숫자 이름의 열 '2'" in str(e.value)
    with pytest.raises(mp.InputError):
        mp.compile_formula("온도*3 + 2.5", ["온도", "2"])             # 2.9.32: '2.5' became <column 2>.5
    # no collision: a formula without that number, or another number that holds its digit
    assert mp.compile_formula("온도*3 + 12", ["온도", "2"])[1] == {"__COL_TOKEN_0__": "온도"}
    proc, res = run_raw("ml_train.py", "-f", data["num"], "-x", "온도", "-y", "수율", "--eval", "x2=온도*2")
    msg = _stop(res, proc)
    assert "숫자 이름의 열" in msg, msg


# ---------------------------------------------------------------------------------------------------------------
# 4. M1: numbers written with '%'
# ---------------------------------------------------------------------------------------------------------------

def test_coerce_numeric_strips_percent():
    s = pd.Series(["64.4%", " 1,234 % ", "5", "%", "abc", "-", "(%)"])
    got = mp.coerce_numeric(s).tolist()
    assert got[:3] == [64.4, 1234.0, 5.0] and all(pd.isna(v) for v in got[3:]), got
    n, ex = mp.percent_cells(s)
    assert n == 2 and ex == ["64.4%", "1,234 %"]


def test_percent_training_and_prediction_match_plain(data):
    proc, res = run_raw("ml_train.py", "-f", data["pct"], "-x", "온도,압력,습도(%)", "-y", "수율(%)", "--seed", "1")
    res = _ok(res, proc)
    w = " ".join(res.get("warnings", []))
    assert "'습도(%)' 열의 60칸" in w and "'수율(%)' 열의 60칸" in w and "100으로 나누지 않음" in w, w
    proc, plain = run_raw("ml_train.py", "-f", data["plain"], "-x", "온도,압력,습도(%)", "-y", "수율(%)", "--seed", "1")
    plain = _ok(plain, proc)
    assert res["mean_val_score"] == plain["mean_val_score"]
    m1 = os.path.join(str(data["tmp"]), "pct_ml_data", "v1")
    m2 = os.path.join(str(data["tmp"]), "plain_ml_data", "v1")
    proc, a = run_raw("ml_predict.py", "-m", m1, "-i", "온도=180,압력=3,습도=55.5%")
    a = _ok(a, proc)
    assert any("%를 떼고" in x for x in a.get("warnings", [])), a
    proc, b = run_raw("ml_predict.py", "-m", m2, "-i", "온도=180,압력=3,습도=55.5")
    b = _ok(b, proc)
    assert a["predictions_summary"] == b["predictions_summary"]


def test_percent_map_warns(data):
    proc, res = run_raw("ml_map.py", "-f", data["pct"], "-i", "온도=180,압력=3", "-y", "수율(%)")
    res = _ok(res, proc)
    assert any("'수율(%)' 열의" in w for w in res.get("warnings", [])), res


# ---------------------------------------------------------------------------------------------------------------
# 5. L1 / L2
# ---------------------------------------------------------------------------------------------------------------

def test_export_constant_inputs_candidates(data):
    proc, res = run_raw("ml_train.py", "-f", data["const"], "-x", "온도,압력,라인", "-y", "수율")
    _ok(res, proc)
    proc, res = run_raw("ml_export.py", "-m", os.path.join(str(data["tmp"]), "const_ml_data", "v1"), "--inputs", "라인")
    msg = _stop(res, proc)
    assert "[추천]" not in msg and "[후보] 온도, 압력" in msg, msg


def test_excluded_column_never_listed():
    msg = _resolve_error(["수율"], ["온도", "수율(1차)", "수율(최종)"], exclude=["수율(최종)"])
    head = msg.split("(비슷한")[0]
    assert "[후보] '수율(1차)'" in head and "수율(최종)" not in head and "[추천]" not in msg, msg
    assert "이 명령에서 이미 다른 이름이나 옵션에 쓰인 열" in msg, msg
    with pytest.raises(mp.InputError) as e:
        ml_export._check_known(["압력", "압룍"], {"y": ["압력", "온도"]}, "꼭 포함할 입력(--keep)")
    msg = str(e.value)
    assert "[추천]" not in msg and "[후보] '압력'" not in msg and "이미 다른 이름이나 옵션에 쓰인 열" in msg, msg


def test_best_task_without_history_is_a_candidate():
    with pytest.raises(mp.InputError) as e:
        ml_train._tasks_to_build([{"task": "regression"}], "classification", "h.jsonl")
    assert "[추천]" not in str(e.value) and "[후보] --task regression" in str(e.value), str(e.value)
