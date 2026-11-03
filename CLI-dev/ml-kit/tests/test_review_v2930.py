# -*- coding: utf-8 -*-
"""
2.9.30 (19차 점검) regressions.

1. Target names were matched with the unit stripped and the first hit taken: ml_predict --find '수율=65' with the
   targets '수율(1차)' and '수율(최종)' solved one of them without a word; -y '수율(1차)' also predicted '수율(최종)'
   (ml_report --target the same). Targets are resolved like columns: exact name, then a unique unit-less name,
   several -> [후보].
2. ml_map --points: a points column that nearly matches an unused data column ('반응 압력(bar)', '압력(MPa)' for
   '반응압력(bar)') passed through as a memo column; the result looked as if the pressure was used.
3. ml_predict -i: keys no model uses ('압력(MPa)=0.5' next to '압력(bar)=3', '습도=50') were dropped without a word.
4. Totals rows (2.9.29 trade-off): a shift report with constant setpoints inside each day and only 생산량 summed
   ('일합계') was not warned. Cells equal to the value of a column constant in the block leave the count, and a sum
   cell needs 2+ nonzero numbers in its column.
"""
import os
import random
import sys

import numpy as np
import pytest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
sys.path.insert(0, os.path.join(os.path.dirname(HERE), "tools"))
from test_layout_matrix_v2917 import mp, _write  # noqa: E402
from test_cli_smoke import run_raw  # noqa: E402

import ml_report  # noqa: E402  (ml-kit folder is on sys.path through test_layout_matrix_v2917)


def _stop(res, proc):
    assert proc.returncode == 1 and res["agent_directive"].startswith("STOP_RETRY:"), res
    return res["message"]


def _ok(res, proc):
    assert proc.returncode == 0 and res["status"] == "ok", res
    return res


def _process(n=60, seed=1):
    r = random.Random(seed)
    rows = [["온도(°C)", "압력(bar)", "속도", "생산량"]]
    for _ in range(n):
        t, p, s = round(r.uniform(170, 190), 1), round(r.uniform(2, 4), 2), round(r.uniform(40, 50), 1)
        rows.append([t, p, s, round(2 * t + 30 * p + 0.5 * s + r.uniform(-2, 2), 1)])
    return rows


@pytest.fixture(scope="module")
def models(tmp_path_factory):
    tmp = tmp_path_factory.mktemp("v2930")
    out = {"tmp": tmp}
    r = random.Random(5)
    rows = [["온도", "압력", "수율(1차)", "수율(최종)"]]
    for _ in range(80):
        t, p = round(r.uniform(170, 190), 1), round(r.uniform(2, 3), 2)
        rows.append([t, p, round(0.3 * t + 2 * p + r.uniform(-1, 1), 2), round(0.2 * t + p + r.uniform(-1, 1), 2)])
    f = _write(os.path.join(str(tmp), "two.csv"), rows, [])
    _, res = run_raw("ml_train.py", "-f", f, "-x", "온도,압력", "-y", "수율(1차),수율(최종)")
    assert res["status"] == "ok", res
    out["two"] = os.path.join(str(tmp), "two_ml_data", "v1")
    f = _write(os.path.join(str(tmp), "p1.csv"), _process(), [])
    _, res = run_raw("ml_train.py", "-f", f, "-x", "온도(°C),압력(bar),속도", "-y", "생산량")
    assert res["status"] == "ok", res
    out["p1"], out["p1_f"] = os.path.join(str(tmp), "p1_ml_data", "v1"), f
    return out


# ---------------------------------------------------------------------------------------------------------------
# 1. target names: no silent pick
# ---------------------------------------------------------------------------------------------------------------

def test_find_ambiguous_goal_refused(models):
    proc, res = run_raw("ml_predict.py", "-m", models["two"], "--find", "수율=65", "--vary", "온도", "-i", "압력=2.5")
    msg = _stop(res, proc)
    assert "[후보]" in msg and "수율(1차)" in msg and "수율(최종)" in msg, msg
    proc, res = run_raw("ml_predict.py", "-m", models["two"], "--find", "수율(1차)=65", "--vary", "온도",
                        "-i", "압력=2.5")
    assert _ok(res, proc)["target"] == "수율(1차)"


def test_predict_targets_filter(models):
    proc, res = run_raw("ml_predict.py", "-m", models["two"], "-y", "수율", "-i", "온도=180,압력=2.5")
    msg = _stop(res, proc)
    assert "[후보]" in msg and "수율(1차)" in msg and "수율(최종)" in msg, msg
    # an exact name takes that target only (it also took '수율(최종)': same name without unit)
    proc, res = run_raw("ml_predict.py", "-m", models["two"], "-y", "수율(최종)", "-i", "온도=180,압력=2.5")
    assert _ok(res, proc)["targets_predicted"] == ["수율(최종)"]


def test_report_target_filter(models):
    with pytest.raises(mp.InputError) as e:
        ml_report.load_version(models["two"], ["수율"])
    assert "[후보]" in str(e.value) and "수율(1차)" in str(e.value) and "수율(최종)" in str(e.value)
    ctx = ml_report.load_version(models["two"], ["수율(1차)"])
    assert [t["name"] for t in ctx["targets"]] == ["수율(1차)"]


def test_same_name_candidates_whatever_the_suffix():
    # difflib alone misses a long unit suffix ('수율' / '수율(최종공정)' ratio 0.4)
    got = mp.suggest_columns("수율", ["수율(1차)", "수율(최종공정)", "온도"])
    assert set(got[:2]) == {"수율(1차)", "수율(최종공정)"}, got
    with pytest.raises(mp.InputError) as e:
        mp.resolve_column_names(["수율"], ["수율(1차)", "수율(최종공정)", "온도"], "목표(--find)", values=True)
    assert "[후보]" in str(e.value) and "수율(최종공정)" in str(e.value)


# ---------------------------------------------------------------------------------------------------------------
# 2. ml_map --points near-name columns
# ---------------------------------------------------------------------------------------------------------------

@pytest.fixture(scope="module")
def map_data(tmp_path_factory):
    tmp = tmp_path_factory.mktemp("v2930map")
    rows = [["온도", "반응압력(bar)", "생산량"]]
    for t in (170, 180, 190):
        for p in (2.0, 2.5, 3.0, 3.5, 4.0):
            rows.append([t, p, round(2 * t + 30 * p, 1)])
    return tmp, _write(os.path.join(str(tmp), "map.csv"), rows, [])


def _points(tmp, name, head, rows):
    return _write(os.path.join(str(tmp), name), [head] + rows, [])


def test_points_near_name_refused(map_data):
    tmp, data = map_data
    p = _points(tmp, "sp.csv", ["온도", "반응압룍(bar)"], [[180, 2.0], [180, 4.0]])
    proc, res = run_raw("ml_map.py", "-f", data, "--points", p, "-y", "생산량", "--save", str(tmp / "o1.xlsx"))
    msg = _stop(res, proc)
    assert "[후보] '반응압력(bar)'" in msg and "반응압룍(bar)" in msg and "[추천]" not in msg, msg     # 2.9.33: [추천]
    # 2.9.31: '반응 압력(bar)' (the same name and unit spaced otherwise) is the key itself
    p = _points(tmp, "sp1.csv", ["온도", "반응 압력(bar)"], [[180, 2.0], [180, 4.0]])
    proc, res = run_raw("ml_map.py", "-f", data, "--points", p, "-y", "생산량", "--save", str(tmp / "o1b.xlsx"))
    assert _ok(res, proc)["keys"] == ["온도", "반응압력(bar)"], res
    p = _points(tmp, "mpa.csv", ["온도", "압력(MPa)"], [[180, 0.3], [180, 0.4]])
    proc, res = run_raw("ml_map.py", "-f", data, "--points", p, "-y", "생산량", "--save", str(tmp / "o2.xlsx"))
    msg = _stop(res, proc)
    assert "[추천]" not in msg and "단위가 다릅니다" in msg and "bar 단위로 환산" in msg, msg


def test_points_memo_columns_pass(map_data):
    tmp, data = map_data
    p = _points(tmp, "memo.csv", ["온도", "반응압력(bar)", "번호", "비고"], [[180, 2.2, 1, "시험"], [185, 3.1, 2, "x"]])
    proc, res = run_raw("ml_map.py", "-f", data, "--points", p, "-y", "생산량", "--save", str(tmp / "o3.xlsx"))
    res = _ok(res, proc)
    assert res["keys"] == ["온도", "반응압력(bar)"] and res["count"] == 2, res


# ---------------------------------------------------------------------------------------------------------------
# 3. -i keys no model uses
# ---------------------------------------------------------------------------------------------------------------

def test_inline_other_unit_extra_key_refused(models):
    proc, res = run_raw("ml_predict.py", "-m", models["p1"], "-i", "압력(bar)=3,압력(MPa)=0.5,온도=180,속도=45")
    msg = _stop(res, proc)
    assert "압력(MPa)" in msg and "단위가 다릅니다" in msg and "[추천]" not in msg, msg


def test_inline_unknown_key_refused(models):
    proc, res = run_raw("ml_predict.py", "-m", models["p1"], "-i", "압력=3,온도=180,속도=45,습도=50")
    msg = _stop(res, proc)
    assert "습도" in msg and "[추천]" not in msg and "[모델 입력 목록]" in msg and "속도" in msg, msg
    proc, res = run_raw("ml_predict.py", "-m", models["p1"], "-i", '{"압력": 3, "온도": 180, "속도": 45, "id": 7}')
    msg = _stop(res, proc)
    assert "'id'" in msg, msg


def test_inline_typo_key_suggested(models):
    # a typo of a model input that is not given: [후보] (2.9.33, was [추천] '속도'; a close name may be another word)
    proc, res = run_raw("ml_predict.py", "-m", models["p1"], "-i", "압력=3,온도=180,속돜=45")
    msg = _stop(res, proc)
    assert "[후보] 속돜 중 '속도'" in msg and "[추천]" not in msg, msg
    # 2.9.32: '속두' (one common syllable for another) is [후보]
    proc, res = run_raw("ml_predict.py", "-m", models["p1"], "-i", "압력=3,온도=180,속두=45")
    msg = _stop(res, proc)
    assert "[추천]" not in msg and "[후보] 속두" in msg, msg


def test_inline_exact_keys_ok_and_file_extra_columns_kept(models):
    proc, res = run_raw("ml_predict.py", "-m", models["p1"], "-i", "압력=3,온도=180,속도=45")
    _ok(res, proc)
    f = _write(os.path.join(str(models["tmp"]), "in_extra.csv"),
               [["번호", "온도(°C)", "압력(bar)", "속도", "습도"], [1, 180, 3, 45, 50], [2, 182, 3.1, 46, 51]], [])
    proc, res = run_raw("ml_predict.py", "-m", models["p1"], "-f", f, "-o", str(models["tmp"] / "in_extra_out.csv"))
    _ok(res, proc)


def test_find_and_map_inline_unknown_key_refused(models):
    proc, res = run_raw("ml_predict.py", "-m", models["p1"], "--find", "생산량=470", "--vary", "온도",
                        "-i", "압력=3,속도=45,습도=50")
    msg = _stop(res, proc)
    assert "습도" in msg and "[추천]" not in msg, msg     # '속도' is given: no typo target
    proc, res = run_raw("ml_map.py", "-f", models["p1_f"], "-i", "압력=3,속도=45,습도=50")
    msg = _stop(res, proc)
    assert "습도" in msg and "[추천]" not in msg, msg


# ---------------------------------------------------------------------------------------------------------------
# 4. totals rows: constant setpoints inside the block
# ---------------------------------------------------------------------------------------------------------------

def _shift_rows(days=10, seed=11):
    r = random.Random(seed)
    rows = [["일자", "구분", "설정온도", "설정압력", "생산량"]]
    for d in range(days):
        t, p = r.choice([175, 180, 185]), r.choice([3.0, 3.5])
        a, b = r.randint(400, 600), r.randint(400, 600)
        rows += [[f"11/{d + 1}", "주간", t, p, a], [f"11/{d + 1}", "야간", t, p, b], [f"11/{d + 1}", "일합계", t, p, a + b]]
    return rows


@pytest.mark.parametrize("ext", ["csv", "xlsx"])
def test_shift_report_with_setpoints_warned(tmp_path, ext):
    f = _write(str(tmp_path / f"s.{ext}"), _shift_rows(), [])
    got = mp.load_data_file(f).attrs["summary_rows"]
    assert [g[2] for g in got] == ["일합계"] * 10, got


FP_CONST = [  # r18 fp_const.csv: category rows 'A계' among '공통', constant setpoints, a 0/1 flag
    ("공통", 0, 107), ("공통", 1, 118), ("A계", 0, 111), ("공통", 1, 129), ("공통", 0, 119), ("A계", 1, 120),
    ("공통", 0, 118), ("공통", 0, 102), ("A계", 0, 129), ("공통", 1, 126), ("공통", 1, 115), ("A계", 1, 117),
    ("공통", 0, 107), ("공통", 1, 106), ("A계", 1, 117), ("공통", 0, 126), ("공통", 0, 117), ("A계", 1, 112),
    ("공통", 1, 120), ("공통", 0, 127), ("A계", 0, 107), ("공통", 0, 120), ("공통", 1, 104), ("A계", 1, 123)]


def test_constant_setpoints_category_still_quiet(tmp_path):
    rows = [["구분", "설정온도", "설정압력", "불량", "생산량"]] + [[g, 180, 3.5, b, q] for g, b, q in FP_CONST]
    assert mp.load_data_file(_write(str(tmp_path / "fp.csv"), rows, [])).attrs["summary_rows"] == []


def test_sums_match_constant_and_nonzero():
    sums = np.array([[370.0, 7.0, 1.0, 1080.0]])
    count = np.full((1, 4), 2.0)
    row = np.array([[185.0, 3.5, 1.0, 1080.0]])
    varies = np.array([[False, False, True, True]])
    # constant 185 / 3.5 leave the count: 1 sum (1080) of 2 cells; the flag 1 has one nonzero number (its own sum)
    assert mp._sums_match(row, sums, count, sums, varies, np.array([[2.0, 2.0, 1.0, 2.0]]))[0]
    assert not mp._sums_match(np.array([[185.0, 3.5, 1.0, 999.0]]), sums, count, sums, varies,
                              np.array([[2.0, 2.0, 1.0, 2.0]]))[0]
    # without nz the 2.9.29 rule (constant cells count as misses)
    sums0 = np.array([[370.0, 7.0, 4.0, 1080.0]])
    row0 = np.array([[185.0, 3.5, 0.0, 1080.0]])
    assert mp._sums_match(row0, sums0, count, sums0, varies, np.full((1, 4), 2.0))[0]
    assert not mp._sums_match(row0, sums0, count, sums0, varies)[0]
