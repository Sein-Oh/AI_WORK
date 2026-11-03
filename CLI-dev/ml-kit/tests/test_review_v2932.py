# -*- coding: utf-8 -*-
"""
2.9.32 (21차 점검) regressions. An agent retries a [추천] by itself, so a [추천] that names another quantity (or one
whose retry stops) is a defect; when unsure the tools give [후보].

1. G1 (2.9.31 regression): exclude= made an ambiguous unit-less name unique. Data columns '수율(1차)', '수율(최종)' with
   -y 수율(최종): a --points column '수율' became the '수율(1차)' key (VERIFIED), '-i 수율=..' got [추천] '수율(1차)'.
2. sure_typo (proposal B): one common syllable for another ('속두'/'속도', '강도'/'경도'), one ASCII letter for another
   ('TempA'/'TempB'), short ASCII names, a name at the front of another ('온도'/'온도차'), a given name longer than the
   column ('설정온도'/'온도'), another unit or qualifier ('수율(1차)'/'수율(최종)') -> [후보]. Rare-syllable typos ('압룍'),
   insertions ('온도x') and swaps ('tmep') keep [추천].
3. Names used in another role are never suggested: ml_train -x/-y (M1), --find -i/--vary (M2), ml_map --vary/--find
   (M3), --eval and the targets (M4), ml_map -y/-i, ml_export --keep (extra sites).
4. --positive-class: [후보] 불량, 양품, not "[추천] ['불량', '양품']" (M5).
5. L1: ml_map --points without -y adds the -y guidance to [후보] too. L2: '온두' is compared with '온도(°C)' without unit.
"""
import os
import random
import sys

import openpyxl
import pandas as pd
import pytest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from test_layout_matrix_v2917 import mp, _write  # noqa: E402
from test_cli_smoke import run_raw  # noqa: E402

import ml_export  # noqa: E402  (ml-kit folder is on sys.path through test_layout_matrix_v2917)


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
    tmp = tmp_path_factory.mktemp("v2932")
    r = random.Random(3)
    two = [["온도", "압력", "수율(1차)", "수율(최종)"]]
    for _ in range(60):
        t, p = round(r.uniform(170, 190), 1), round(r.uniform(2, 3), 2)
        two.append([t, p, round(0.3 * t + 2 * p + r.uniform(-1, 1), 2), round(0.2 * t + p + r.uniform(-1, 1), 2)])
    proc = [["온도(°C)", "압력(bar)", "속도", "습도", "생산량"]]
    for _ in range(60):
        t, p, s, h = (round(r.uniform(170, 190), 1), round(r.uniform(2, 4), 2), round(r.uniform(40, 50), 1),
                      round(r.uniform(30, 70), 1))
        proc.append([t, p, s, h, round(2 * t + 30 * p + 0.5 * s + r.uniform(-2, 2), 1)])
    cls = [["온도", "압력", "판정"]]
    for i in range(40):
        t, p = round(r.uniform(170, 190), 1), round(r.uniform(2, 4), 2)
        cls.append([t, p, "불량" if t + 10 * p > 215 or i % 7 == 0 else "양품"])
    out = {"tmp": tmp, "two": _write(os.path.join(str(tmp), "two.csv"), two, []),
           "p": _write(os.path.join(str(tmp), "p.csv"), proc, []),
           "c": _write(os.path.join(str(tmp), "c.csv"), cls, [])}
    _, res = run_raw("ml_train.py", "-f", out["p"], "-x", "온도(°C),압력(bar),속도,습도", "-y", "생산량")
    assert res["status"] == "ok", res
    out["pm"] = os.path.join(str(tmp), "p_ml_data", "v1")
    return out


def _points(data, name, head, rows):
    return _write(os.path.join(str(data["tmp"]), name), [head] + rows, [])


# ---------------------------------------------------------------------------------------------------------------
# 1. G1: an excluded column of the same name
# ---------------------------------------------------------------------------------------------------------------

def test_points_unitless_name_of_the_target_is_no_key(data):
    pts = _points(data, "pts.csv", ["온도", "압력", "수율"], [[180, 2.5, 40.1], [185, 2.6, 41]])
    out = str(data["tmp"] / "pts_o.xlsx")
    proc, res = run_raw("ml_map.py", "-f", data["two"], "--points", pts, "-y", "수율(최종)", "--save", out, "--overwrite")
    res = _ok(res, proc)
    assert res["keys"] == ["온도", "압력"] and res["columns"] == ["수율(최종)"], res     # 2.9.31: 수율(1차) a key
    assert "수율" in [c.value for c in openpyxl.load_workbook(res["output_file"])["결과"][1]]


def test_inline_unitless_name_next_to_the_target_asks(data):
    proc, res = run_raw("ml_map.py", "-f", data["two"], "-i", "온도=180,수율=40", "-y", "수율(최종)")
    msg = _stop(res, proc)
    assert "[추천]" not in msg and "[후보] '수율(1차)', '수율(최종)'" in msg, msg
    msg = _resolve_error(["수율"], ["온도", "수율(1차)", "수율(최종)"], exclude=["수율(최종)"])
    assert "[추천]" not in msg and "[후보] '수율(1차)', '수율(최종)'" in msg, msg


# ---------------------------------------------------------------------------------------------------------------
# 2. sure_typo
# ---------------------------------------------------------------------------------------------------------------

ASK = [("속두", "속도"), ("습두", "습도"), ("온두", "온도"), ("도핑농두", "도핑농도"), ("강도", "경도"), ("점도", "정도"),
       ("생산랑", "생산량"), ("TempA", "TempB"), ("온도A", "온도B"), ("Heat", "Head"), ("T1", "X1"), ("O2", "T2"),
       ("Ti", "Tim"), ("Ti", "Time"), ("온도", "온도차"), ("설정온도", "온도"), ("반응온도", "온도(°C)"),
       ("수율(1차)", "수율(최종)"), ("생산량(kg)", "생산량(t)"), ("압력(MPa)", "반응압력(bar)")]
SURE = [("압룍", "압력"), ("속돜", "속도"), ("온도x", "온도"), ("tmep", "temp"), ("압력", "반응압력"),
        ("압력(bar)", "반응압력(bar)"), ("반응온도(C)", "반응온도(°C)"), ("Prssure", "Pressure"), ("수율( 1차)", "수율")]


@pytest.mark.parametrize("a,b", ASK)
def test_not_a_sure_typo(a, b):
    assert not mp.sure_typo(a, b)
    msg = _resolve_error([a], [b, "zzz"])
    assert "[추천]" not in msg and "[후보]" in msg, msg


@pytest.mark.parametrize("a,b", SURE)
def test_sure_typo(a, b):
    assert mp.sure_typo(a, b)
    assert f"[추천] '{b}'" in _resolve_error([a], [b, "zzz"])


def test_common_syllable_typo_cli_is_a_candidate(data):
    proc, res = run_raw("ml_predict.py", "-m", data["pm"], "-i", "온도=180,압력=3,속도=45,습두=50")
    msg = _stop(res, proc)
    assert "[추천]" not in msg and "[후보] 습두" in msg, msg
    proc, res = run_raw("ml_predict.py", "-m", data["pm"], "-i", "온도=180,압력=3,속도=45,습돜=50")
    assert "[추천] '습도'" in _stop(res, proc)


# ---------------------------------------------------------------------------------------------------------------
# 3. names used in another role
# ---------------------------------------------------------------------------------------------------------------

def test_train_x_typo_never_suggests_the_target(data):
    proc, res = run_raw("ml_train.py", "-f", data["p"], "-x", "온도,압력,생산랑", "-y", "생산량")
    msg = _stop(res, proc)
    assert "[추천]" not in msg and "'생산량'" in msg, msg
    proc, res = run_raw("ml_train.py", "-f", data["two"], "-x", "온도,압력,수율", "-y", "수율(최종)")
    msg = _stop(res, proc)
    assert "[추천]" not in msg and "[후보] '수율(1차)', '수율(최종)'" in msg, msg
    # the reverse: a -y typo never suggests a feature
    proc, res = run_raw("ml_train.py", "-f", data["p"], "-x", "온도,압력", "-y", "압룍")
    msg = _stop(res, proc)
    assert "[추천]" not in msg, msg


def test_find_inputs_never_suggest_each_other(data):
    proc, res = run_raw("ml_predict.py", "-m", data["pm"], "--find", "생산량=480", "--vary", "속도", "-i", "속돜=45")
    msg = _stop(res, proc)
    assert "[추천]" not in msg, msg                                   # M2: the --vary input
    proc, res = run_raw("ml_predict.py", "-m", data["pm"], "--find", "생산량=480", "--vary", "속돜", "-i", "속도=45")
    msg = _stop(res, proc)
    assert "[추천]" not in msg, msg                                   # an input fixed by -i


def test_map_find_vary_never_suggest_each_other(data):
    proc, res = run_raw("ml_map.py", "-f", data["p"], "--find", "생산량=480", "--vary", "생산랑")
    msg = _stop(res, proc)
    assert "[추천]" not in msg and "생산량" in msg, msg               # M3: the --find goal
    proc, res = run_raw("ml_map.py", "-f", data["p"], "--find", "속돜=45", "--vary", "속도")
    assert "[추천]" not in _stop(res, proc)                            # the goal is no --vary column
    proc, res = run_raw("ml_map.py", "-f", data["p"], "-i", "속도=45", "-y", "속돜")
    assert "[추천]" not in _stop(res, proc)                            # -y is no -i key


def test_eval_never_suggests_a_target():
    df = pd.DataFrame({"온도": [1.0, 2.0], "압력": [1.0, 2.0], "생산량": [3.0, 4.0]})
    with pytest.raises(mp.InputError) as e:
        mp.apply_eval_expressions(df, ["d = 생산랑 * 2"], target_cols=["생산량"])
    assert "[추천]" not in str(e.value) and "타겟" in str(e.value), str(e.value)
    with pytest.raises(mp.InputError) as e:
        mp.apply_eval_expressions(df, ["d = 온돜 * 2"], target_cols=["생산량"])
    assert "[추천] 온도" in str(e.value), str(e.value)


def test_train_eval_typo_of_the_target(data):
    proc, res = run_raw("ml_train.py", "-f", data["p"], "-x", "온도,압력", "-y", "생산량", "--eval", "d = 생산랑 * 2")
    msg = _stop(res, proc)
    assert "[추천]" not in msg, msg


def test_export_keep_typo_never_suggests_a_given_input():
    with pytest.raises(mp.InputError) as e:
        ml_export._check_known(["압력", "압룍"], {"t": ["압력(bar)", "온도"]}, "입력(--keep)")
    assert "[추천]" not in str(e.value), str(e.value)
    with pytest.raises(mp.InputError) as e:
        ml_export._check_known(["압룍"], {"t": ["압력(bar)", "온도"]}, "입력(--keep)")
    assert "[추천] '압력(bar)'" in str(e.value), str(e.value)


def test_known_columns():
    assert mp.known_columns(["온도", "압력( bar)", "없음"], ["온도(°C)", "압력(bar)"]) == ["온도(°C)", "압력(bar)"]


# ---------------------------------------------------------------------------------------------------------------
# 4. --positive-class
# ---------------------------------------------------------------------------------------------------------------

def test_positive_class_candidates(data):
    proc, res = run_raw("ml_train.py", "-f", data["c"], "-x", "온도,압력", "-y", "판정", "--task", "classification",
                        "--positive-class", "불랑")
    msg = _stop(res, proc)
    assert "[추천]" not in msg and "[후보] 불량, 양품" in msg and "['" not in msg, msg


# ---------------------------------------------------------------------------------------------------------------
# 5. L1, L2
# ---------------------------------------------------------------------------------------------------------------

def test_points_without_y_candidate_has_the_y_guidance(data):
    pts = _points(data, "h2.csv", ["온도(°C)", "압력(bar)", "습도2"], [[180, 3, 1], [185, 3.5, 2]])
    out = str(data["tmp"] / "h2_o.xlsx")
    proc, res = run_raw("ml_map.py", "-f", data["p"], "--points", pts, "--save", out, "--overwrite")
    msg = _stop(res, proc)
    assert "[추천]" not in msg and "[후보] '습도'" in msg and "-y에 구할 열" in msg, msg


def test_unitless_name_against_unit_columns():
    msg = _resolve_error(["온두"], ["온도(°C)", "압력(bar)"])
    assert "[추천]" not in msg and "[후보] '온도(°C)'" in msg, msg           # 2.9.31: 유사 컬럼 없음
    assert "[추천] '온도(°C)'" in _resolve_error(["온돜"], ["온도(°C)", "압력(bar)"])
    hint = mp._typo_hints(["온도(°C)"], ["온두"], ["온도(°C)"], True)
    assert len(hint) == 1 and "[후보] 온두" in hint[0], hint
