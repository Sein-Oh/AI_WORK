# -*- coding: utf-8 -*-
"""
2.9.31 (20차 점검) regressions.

1. ml_map --points (2.9.30 regression): the -y target was a typo target of the near-name check, so a measured column
   '생산량_실측' got "[추천] '생산량'" and the retry failed; '생산량(실측)' was read as a key with the unit '실측'
   ("기준(-i)" [추천]). A points column of the target's name is a measured column (passed through); without -y the
   check gives [후보], never [추천].
2. [추천] between other words or numbers (습도/속도, 온도/농도, T1/T2, 온도1/온도2, 반응온도/반응농도, Temp/Time) and short
   ASCII names (CO2/O2, pH/H, Ax/A): an agent retries a [추천] by itself. sure_typo decides; otherwise [후보].
3. _typo_hints: one given name close to two missing inputs ('솔도' for '속도' and '습도') got two [추천].
4. --eval unknown name: "[추천] ['수율(1차)', '수율(최종)']" (a Python list).
5. Totals rows: the known false warning (a small-integer sum + a constant setpoint, category row 'A계') is pinned.
6. ml_report --target '수율( 1차)' gave [후보] while -y matched: same-unit spacing is matched in resolve_column_names.
7. ml_export: two targets of one name without unit ('수율(1차)', '수율(최종)') were both labelled '수율'.
"""
import os
import random
import sys

import openpyxl
import pandas as pd
import pytest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
sys.path.insert(0, os.path.join(os.path.dirname(HERE), "tools"))
from test_layout_matrix_v2917 import mp, _write  # noqa: E402
from test_cli_smoke import run_raw  # noqa: E402

import ml_report  # noqa: E402  (ml-kit folder is on sys.path through test_layout_matrix_v2917)
import ml_export  # noqa: E402


def _stop(res, proc):
    assert proc.returncode == 1 and res["agent_directive"].startswith("STOP_RETRY:"), res
    return res["message"]


def _ok(res, proc):
    assert proc.returncode == 0 and res["status"] == "ok", res
    return res


def _process(n=60, seed=1):
    r = random.Random(seed)
    rows = [["온도(°C)", "압력(bar)", "속도", "습도", "생산량"]]
    for _ in range(n):
        t, p, s, h = (round(r.uniform(170, 190), 1), round(r.uniform(2, 4), 2), round(r.uniform(40, 50), 1),
                      round(r.uniform(30, 70), 1))
        rows.append([t, p, s, h, round(2 * t + 30 * p + 0.5 * s + r.uniform(-2, 2), 1)])
    return rows


@pytest.fixture(scope="module")
def data(tmp_path_factory):
    tmp = tmp_path_factory.mktemp("v2931")
    out = {"tmp": tmp, "p": _write(os.path.join(str(tmp), "p.csv"), _process(), [])}
    r = random.Random(5)
    rows = [["온도", "압력", "수율(1차)", "수율(최종)"]]
    for _ in range(80):
        t, p = round(r.uniform(170, 190), 1), round(r.uniform(2, 3), 2)
        rows.append([t, p, round(0.3 * t + 2 * p + r.uniform(-1, 1), 2), round(0.2 * t + p + r.uniform(-1, 1), 2)])
    f = _write(os.path.join(str(tmp), "two.csv"), rows, [])
    _, res = run_raw("ml_train.py", "-f", f, "-x", "온도,압력", "-y", "수율(1차),수율(최종)")
    assert res["status"] == "ok", res
    out["two"] = os.path.join(str(tmp), "two_ml_data", "v1")
    _, res = run_raw("ml_train.py", "-f", out["p"], "-x", "온도(°C),압력(bar),속도,습도", "-y", "생산량")
    assert res["status"] == "ok", res
    out["pm"] = os.path.join(str(tmp), "p_ml_data", "v1")
    return out


# ---------------------------------------------------------------------------------------------------------------
# 1. ml_map --points: the target's measured column
# ---------------------------------------------------------------------------------------------------------------

def _points(data, name, head, rows):
    return _write(os.path.join(str(data["tmp"]), name), [head] + rows, [])


def _map_points(data, pts, *extra):
    out = str(data["tmp"] / (os.path.basename(pts) + str(len(extra)) + "_o.xlsx"))
    return run_raw("ml_map.py", "-f", data["p"], "--points", pts, "--save", out, "--overwrite", *extra)


@pytest.mark.parametrize("col", ["생산량_실측", "생산량(실측)", "생산량"])
def test_points_target_measured_column_passes_with_y(data, col):
    pts = _points(data, f"m_{len(col)}_{col[-2:]}.csv", ["온도(°C)", "압력(bar)", col], [[180, 3, 470], [185, 3.5, 490]])
    res = _ok(*reversed(_map_points(data, pts, "-y", "생산량")))
    assert res["keys"] == ["온도(°C)", "압력(bar)"] and res["columns"] == ["생산량"] and res["count"] == 2, res
    got = openpyxl.load_workbook(res["output_file"])["결과"]
    assert col in [c.value for c in got[1]]                      # the measured column is kept in the result sheet


@pytest.mark.parametrize("col", ["생산량_실측", "생산량(실측)"])
def test_points_without_y_asks(data, col):
    pts = _points(data, f"n_{col[-3:-1]}.csv", ["온도(°C)", "압력(bar)", col], [[180, 3, 470], [185, 3.5, 490]])
    proc, res = _map_points(data, pts)
    msg = _stop(res, proc)
    assert "[추천]" not in msg and "[후보]" in msg and "-y" in msg and "기준(-i)" not in msg, msg


def test_points_other_word_is_no_recommendation(data):
    # '습도2' next to the data column '습도': another number, [후보]; a sure typo keeps [추천] and its retry works
    # (2.9.32: '속두' is one common syllable for another, [후보]; the rare-syllable typo '속돜' keeps [추천])
    pts = _points(data, "h2.csv", ["온도(°C)", "압력(bar)", "습도2"], [[180, 3, 1], [185, 3.5, 2]])
    proc, res = _map_points(data, pts, "-y", "생산량")
    msg = _stop(res, proc)
    assert "[추천]" not in msg and "[후보] '습도'" in msg, msg
    pts = _points(data, "sp.csv", ["온도(°C)", "압력 (bar)", "속두"], [[180, 3, 45], [185, 3.5, 46]])
    proc, res = _map_points(data, pts, "-y", "생산량")
    msg = _stop(res, proc)
    assert "[추천]" not in msg and "[후보] '속도'" in msg, msg
    pts = _points(data, "sp1.csv", ["온도(°C)", "압력 (bar)", "속돜"], [[180, 3, 45], [185, 3.5, 46]])
    proc, res = _map_points(data, pts, "-y", "생산량")
    assert "[추천] '속도'" in _stop(res, proc)
    pts = _points(data, "sp2.csv", ["온도(°C)", "압력 (bar)", "속도"], [[180, 3, 45], [185, 3.5, 46]])
    _ok(*reversed(_map_points(data, pts, "-y", "생산량")))


def test_inline_keys_never_suggest_the_value_column(data):
    # -i next to -y / --find / --vary: the column those take is no key ([추천] '생산량' would fail on the retry)
    proc, res = run_raw("ml_map.py", "-f", data["p"], "-i", "온도=180,생산랑=470", "-y", "생산량")
    assert "[추천] '생산량'" not in _stop(res, proc)
    proc, res = run_raw("ml_map.py", "-f", data["p"], "--find", "생산량=470", "--vary", "온도", "-i", "생산랑=3")
    assert "[추천] '생산량'" not in _stop(res, proc)
    proc, res = run_raw("ml_map.py", "-f", data["p"], "-i", "온도=180,압룍(bar)=3", "-y", "생산량")
    assert "[추천] '압력(bar)'" in _stop(res, proc)


# ---------------------------------------------------------------------------------------------------------------
# 2. sure_typo
# ---------------------------------------------------------------------------------------------------------------

OTHER_WORDS = [("습도", "속도"), ("온도", "농도"), ("T1", "T2"), ("온도1", "온도2"), ("반응온도", "반응농도"),
               ("Temp", "Time"), ("CO2", "O2"), ("pH", "H"), ("Ax", "A")]
# 2.9.32: '속두'/'속도' and '도핑농두'/'도핑농도' (one common syllable for another) and '생산량(kg)'/'생산량(t)' (another
# unit) moved to [후보] (test_review_v2932.py); the rare-syllable typos '속돜' and '도핑농돜' keep [추천]
TYPOS = [("속돜", "속도"), ("압력(bar)", "반응압력(bar)"), ("반응 온도", "반응온도"), ("반응온도(C)", "반응온도(°C)"),
         ("온도x", "온도"), ("도핑농돜", "도핑농도"), ("tmep", "temp"), ("생산량(kg)", "생산량")]


@pytest.mark.parametrize("a,b", OTHER_WORDS)
def test_other_word_is_a_candidate(a, b):
    assert not mp.sure_typo(a, b)
    with pytest.raises(mp.InputError) as e:
        mp.resolve_column_names([a], [b, "생산량"], "피처(-x)")
    assert "[추천]" not in str(e.value) and f"[후보] '{b}'" in str(e.value), str(e.value)


@pytest.mark.parametrize("a,b", TYPOS)
def test_typo_still_recommended(a, b):
    assert mp.sure_typo(a, b)


def test_typo_resolve_messages():
    for a, b in (("속돜", "속도"), ("압력(bar)", "반응압력(bar)"), ("반응온도(C)", "반응온도(°C)"), ("온도x", "온도")):
        with pytest.raises(mp.InputError) as e:
            mp.resolve_column_names([a], [b, "생산량"], "피처(-x)")
        assert f"[추천] '{b}'" in str(e.value), str(e.value)


def test_train_and_predict_other_word(data):
    proc, res = run_raw("ml_train.py", "-f", data["p"], "-x", "온도,압력,습두", "-y", "생산량")
    msg = _stop(res, proc)                                    # 2.9.32: one common syllable for another, [후보]
    assert "[추천]" not in msg and "[후보] '습도'" in msg, msg
    proc, res = run_raw("ml_train.py", "-f", data["p"], "-x", "온도,압력,습돜", "-y", "생산량")
    assert "[추천] '습도'" in _stop(res, proc)                               # a rare syllable: a typo
    proc, res = run_raw("ml_predict.py", "-m", data["two"], "-i", "온도=180,압력=2.5,농도=3")
    msg = _stop(res, proc)
    assert "[추천]" not in msg and "농도" in msg, msg


# ---------------------------------------------------------------------------------------------------------------
# 3. one given name close to two missing inputs
# ---------------------------------------------------------------------------------------------------------------

def test_one_name_for_two_inputs(data):
    f = _write(os.path.join(str(data["tmp"]), "sol.csv"), [["온도(°C)", "압력(bar)", "솔도"], [180, 3, 45]], [])
    proc, res = run_raw("ml_predict.py", "-m", data["pm"], "-f", f, "-o", str(data["tmp"] / "sol_out.csv"))
    msg = _stop(res, proc)
    assert "[추천]" not in msg and msg.count("[후보]") == 1 and "[후보] 속도, 습도" in msg, msg
    hints = mp._typo_hints(["속도", "습도"], ["온도(°C)", "솔도"], ["온도(°C)", "속도", "습도"], False)
    assert len(hints) == 1 and "[후보] 속도, 습도" in hints[0], hints
    assert mp._typo_hints(["속도"], ["속돜"], ["속도"], True) == \
        ["-i의 '속돜' -> [추천] '속도' ('속돜'을(를) '속도'(으)로 고쳐 다시 실행)"]
    assert "[후보] 속두" in mp._typo_hints(["속도"], ["속두"], ["속도"], True)[0]     # 2.9.32: [추천] -> [후보]
    assert "[추천]" not in mp._typo_hints(["속도"], ["습도"], ["속도"], True)[0]


# ---------------------------------------------------------------------------------------------------------------
# 4. --eval unknown names
# ---------------------------------------------------------------------------------------------------------------

def _formula_error(expr, cols):
    with pytest.raises(mp.InputError) as e:
        mp.apply_eval_expressions(pd.DataFrame({c: [1.0, 2.0] for c in cols}), [expr])
    return str(e.value)


def test_eval_unknown_name_wording():
    msg = _formula_error("d = 수율 * 2", ["수율(1차)", "수율(최종)", "온도"])
    assert "[후보] 수율(1차), 수율(최종)" in msg and "['" not in msg.split("[사용 가능한")[0], msg
    assert "-> [후보] 온도." in _formula_error("d = 온두 * 2", ["온도", "압력"])     # 2.9.32: [추천] -> [후보]
    assert "-> [추천] 온도." in _formula_error("d = 온돜 * 2", ["온도", "압력"])
    msg = _formula_error("d = T3 - T1", ["T1", "T2"])
    assert "[추천]" not in msg and "[후보] T2" in msg, msg


# ---------------------------------------------------------------------------------------------------------------
# 5. totals rows: known false warning pinned (2.9.30 rule kept)
# ---------------------------------------------------------------------------------------------------------------

CAT = [("일반", 1, 100), ("일반", 1, 104), ("A계", 2, 101), ("일반", 0, 99), ("일반", 2, 103), ("일반", 1, 98),
       ("A계", 1, 97), ("일반", 1, 102), ("일반", 0, 96), ("A계", 0, 105), ("일반", 2, 100), ("일반", 1, 101),
       ("일반", 0, 99), ("일반", 1, 103)]


def test_known_false_warning_small_integer_sum(tmp_path):
    rows = [["라인", "설정온도", "불량수", "생산량"]] + [[g, 180, b, q] for g, b, q in CAT]
    got = mp.load_data_file(_write(str(tmp_path / "cat.csv"), rows, [])).attrs["summary_rows"]
    assert got == [[4, "라인", "A계"]], got        # 불량수 2 = 1 + 1 and a constant setpoint: warned (documented)


# ---------------------------------------------------------------------------------------------------------------
# 6. same-unit spacing everywhere
# ---------------------------------------------------------------------------------------------------------------

def test_spaced_unit_is_exact(data):
    ctx = ml_report.load_version(data["two"], ["수율( 1차)"])
    assert [t["name"] for t in ctx["targets"]] == ["수율(1차)"]
    proc, res = run_raw("ml_predict.py", "-m", data["two"], "-y", "수율( 1차)", "-i", "온도=180,압력=2.5")
    assert _ok(res, proc)["targets_predicted"] == ["수율(1차)"]
    assert mp.resolve_column_names(["압력 ( bar )"], ["압력(bar)", "온도"], "피처(-x)")[0] == ["압력(bar)"]
    proc, res = run_raw("ml_predict.py", "-m", data["pm"], "-i", "온도( °C)=180,압력 [bar]=3,속도=45,습도=50")
    assert _ok(res, proc)["status"] == "ok"                                 # model inputs '온도(°C)', '압력(bar)'
    with pytest.raises(mp.InputError) as e:                                  # another spelling: still [추천]
        mp.resolve_column_names(["온도(C)"], ["온도(°C)"], "피처(-x)")
    assert "[추천] '온도(°C)'" in str(e.value)


# ---------------------------------------------------------------------------------------------------------------
# 7. ml_export labels of two targets of one name
# ---------------------------------------------------------------------------------------------------------------

def test_export_labels_two_targets(data):
    assert ml_export.target_labels(["수율(1차)", "수율(최종)", "온도(°C)"]) == \
        {"수율(1차)": "수율(1차)", "수율(최종)": "수율(최종)", "온도(°C)": "온도"}
    out = str(data["tmp"] / "exp.xlsx")
    proc, res = run_raw("ml_export.py", "-m", data["two"], "--inputs", "온도,압력", "--save", out)
    res = _ok(res, proc)
    wb = openpyxl.load_workbook(res["output_file"])
    heads = [c.value for c in wb["일괄계산"][2] if c.value]
    assert "예측 수율(1차)" in heads and "예측 수율(최종)" in heads and "예측 수율" not in heads, heads
    note = "\n".join(str(c.value) for row in wb["계산기"].iter_rows() for c in row if isinstance(c.value, str))
    assert "- 수율(1차):" in note and "- 수율(최종):" in note and "- 수율:" not in note, note
    assert {"표_수율(1차)", "표_수율(최종)"} <= set(wb.sheetnames), wb.sheetnames
