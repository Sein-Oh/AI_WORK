# -*- coding: utf-8 -*-
"""
2.9.22 (11차 점검) regressions.

1. PERF-R1: the 2.9.21 value check walked back over consecutive summary rows for every candidate (quadratic): a
   100k-row long-format statistic file ('집계' = 합계/평균/최대/최소) took 111s to load (2.9.20: 0.7s).
2. SUM-R1: 2.9.21 warned totals words only when the row equalled the sum of the rows above it: a bottom block
   (합계/평균/최대/최소 under the data) with a ratio column, a grand 합계 under 소계 groups and 소계 rows of a column
   with 평균 as a category were trained silently. Totals words are warned again except in a statistic label column.
3. SUM-J2: the value check read '2,452' (CSV saved with a #,##0 format) as text: shift 평균 rows were not warned.
4. DOT-R1/DOT-R2: a decimal-comma CSV read dot-grouped numbers ('1.300', '2.125') as thousands (2.9.21); a model
   trained on a decimal-comma file read a prediction file without decimal commas the other way ('1.300' = 1.3).
   Dot-grouped numbers in a decimal-comma CSV are refused; the model keeps the decision for prediction files and
   ml_map keeps it for --points files.
"""
import csv
import os
import sys
import time

import numpy as np
import pytest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from test_layout_matrix_v2917 import mp, _write  # noqa: E402
from test_cli_smoke import run_raw  # noqa: E402


def _summary(path):
    df = mp.load_data_file(path)
    return [(r[0], r[2]) for r in df.attrs["summary_rows"]]


def _csv(path, rows, sep=";"):
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        csv.writer(f, delimiter=sep).writerows(rows)
    return str(path)


# 1. PERF-R1 ---------------------------------------------------------------------------------------------------------

def _stat_file(path, n):
    r = np.random.default_rng(3)
    words = ["합계", "평균", "최대", "최소"]
    rows = [["집계"] + [f"x{c}" for c in range(6)]]
    for i in range(n):
        rows.append([words[i % 4]] + [f"{v:.3f}" for v in r.normal(0, 2, 6)])
    return _csv(path, rows, sep=",")


def test_statistic_column_load_linear(tmp_path):
    # 2.9.21: 30k rows 11s, 100k rows 111s. A generous bound for a loaded machine that still catches quadratic growth.
    small, big = _stat_file(tmp_path / "s.csv", 5000), _stat_file(tmp_path / "b.csv", 40000)
    mp.load_data_file(small)
    t0 = time.perf_counter()
    assert _summary(small) == []
    t_small = time.perf_counter() - t0
    t0 = time.perf_counter()
    assert _summary(big) == []
    t_big = time.perf_counter() - t0
    # 8x rows: linear ~8x (2.9.22: 0.03s -> 0.13s), quadratic ~64x (2.9.21: 0.3s -> about 18s)
    assert t_big < max(16 * t_small, 2.0), (t_small, t_big)


# 2. SUM-R1 ----------------------------------------------------------------------------------------------------------

def _lines(n_rep=5, seed=4, ratio=True):
    r = np.random.default_rng(seed)
    rows = []
    for _ in range(n_rep):
        for line in ["L1", "L2", "L3", "L4"]:
            p, b = int(r.integers(500, 2000)), int(r.integers(5, 30))
            rows.append([line, round(float(r.uniform(172, 184)), 1), p, b] + ([round(100 * b / p, 2)] if ratio else []))
    return rows


def _bottom(data, ratio=True):
    cols = list(zip(*data))
    p, b = sum(cols[2]), sum(cols[3])
    n = len(data)
    tot = ["합계", None, p, b] + ([round(100 * b / p, 2)] if ratio else [])
    avg = ["평균", round(sum(cols[1]) / n, 1), round(p / n, 1), round(b / n, 1)] + ([round(100 * b / p, 2)] if ratio else [])
    mx = ["최대"] + [max(c) for c in cols[1:]]
    mn = ["최소"] + [min(c) for c in cols[1:]]
    return [tot, avg, mx, mn]


def test_bottom_block_total_warned(tmp_path):
    head = [["라인", "온도", "생산량", "불량수", "불량률"]]
    data = _lines()
    got = _summary(_write(str(tmp_path / "b5.xlsx"), head + data + _bottom(data), []))
    assert (len(data) + 2, "합계") in got                                     # 2.9.21: [] (ratio column != sum)
    data = _lines(ratio=False)
    got = _summary(_write(str(tmp_path / "b1n.xlsx"), [head[0][:4]] + data + _bottom(data, False), []))
    assert (len(data) + 2, "합계") in got


def test_subtotal_groups_and_grand_total_warned(tmp_path):
    rows = [["라인", "온도", "생산량", "불량수"]]
    data = _lines(ratio=False)
    for line in ["L1", "L2", "L3", "L4"]:
        part = [d for d in data if d[0] == line]
        rows += part + [["소계", None, sum(d[2] for d in part), sum(d[3] for d in part)]]
    rows += [["합계", None, sum(d[2] for d in data), sum(d[3] for d in data)]] + _bottom(data, False)[1:]
    got = _summary(_write(str(tmp_path / "b2.xlsx"), rows, []))
    want = [i + 1 for i, r in enumerate(rows) if r[0] in ("소계", "합계")]
    assert [g[0] for g in got if g[1] in ("소계", "합계")] == want                # 2.9.21: grand 합계 not named


def test_subtotal_in_average_category_column_warned(tmp_path):
    r = np.random.default_rng(6)
    rows = [["측정방식", "온도", "생산량"]]
    for _ in range(5):
        block = [[w, round(float(r.uniform(172, 184)), 1), int(r.integers(900, 1100))] for w in ["개별", "평균"] * 2]
        rows += block + [["소계", None, sum(b[2] for b in block)]]
    got = _summary(_write(str(tmp_path / "b3.xlsx"), rows, []))
    assert got == [(i + 1, "소계") for i, r in enumerate(rows) if r[0] == "소계"]    # 2.9.21: []


def test_statistic_label_columns_stay_quiet(tmp_path):
    r = np.random.default_rng(9)
    for name, words in [("sb", ["합계", "평균", "최대", "최소"]), ("sc", ["Sum", "Avg"]), ("sa", ["평균"])]:
        rows = [["집계", "온도", "수율"]] + [[w, round(float(r.uniform(172, 179)), 1), round(float(r.uniform(88, 92)), 1)]
                                         for _ in range(6) for w in words]
        assert _summary(_write(str(tmp_path / f"{name}.xlsx"), rows, [])) == [], name


def test_lone_totals_in_note_column_warned(tmp_path):
    # a mostly blank note column holding only 소계/합계: not a statistic label column
    rows = [["비고", "생산량", "불량수"]]
    data = [[None, 1000 + 7 * i, 10 + i] for i in range(6)]
    rows += data[:3] + [["소계", sum(d[1] for d in data[:3]), sum(d[2] for d in data[:3])]]
    rows += data[3:] + [["소계", sum(d[1] for d in data[3:]), sum(d[2] for d in data[3:])]]
    rows += [["합계", sum(d[1] for d in data), sum(d[2] for d in data)]]
    assert [g[1] for g in _summary(_write(str(tmp_path / "note.xlsx"), rows, []))] == ["소계", "소계", "합계"]


# 3. SUM-J2 ----------------------------------------------------------------------------------------------------------

def test_shift_average_with_thousands_commas_warned(tmp_path):
    r = np.random.default_rng(5)
    rows = [["날짜", "구분", "생산량", "불량수"]]
    for d in range(8):
        block = [[int(r.integers(2000, 3000)), int(r.integers(5000, 9000))] for _ in range(3)]
        for g, b in zip(["A조", "B조", "C조"], block):
            rows.append([f"2026-10-{d + 1:02d}", g, f"{b[0]:,}", f"{b[1]:,}"])
        rows.append([f"2026-10-{d + 1:02d}", "평균", f"{round(sum(b[0] for b in block) / 3):,}",
                     f"{round(sum(b[1] for b in block) / 3):,}"])
    got = _summary(_csv(tmp_path / "t4.csv", rows, sep=","))
    assert got == [(i + 1, "평균") for i, r in enumerate(rows) if r[1] == "평균"]   # 2.9.21: []


# 4. DOT-R1 / DOT-R2 -------------------------------------------------------------------------------------------------

def test_tab_line_list_with_three_decimals_refused(tmp_path):
    # '1,2' (line list) confirms decimal commas; '2.125' is 2.125 or 2125 -> refused (2.9.21: 2125)
    rows = [["라인", "두께", "온도", "수율"]]
    for i in range(30):
        rows.append([["1,2", "3,4", "1"][i % 3], ["2.125", "1.375", "1.500", "1.125"][i % 4], 170 + i % 10, 80 + i % 7])
    with pytest.raises(mp.InputError, match="2.125"):
        mp.load_data_file(_csv(tmp_path / "d1.csv", rows, sep="\t"))


def _euro(x):
    return f"{x:,.1f}".replace(",", "X").replace(".", ",").replace("X", ".")


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
    _, res = run_raw("ml_train.py", "-f", _csv(d / "train.csv", _euro_rows()), "-x", "온도,유량", "-y", "수율")
    assert res["status"] == "ok", res
    return d, res["version_dir"]


def test_model_keeps_decimal_comma_for_prediction(euro_model):
    d, vdir = euro_model
    import json
    with open(os.path.join(vdir, "model_수율.json"), encoding="utf-8") as f:
        assert json.load(f)["decimal_comma"] is True
    # no decimal comma in the prediction file: '1.300' would be 1.3 here and 1300 in training -> refused
    _, res = run_raw("ml_predict.py", "-m", vdir, "-f", _csv(d / "pred.csv", [["온도", "유량"], [175, "1.300"],
                                                                            [180, "1.350"], [178, "1.320"]]))
    assert res["status"] == "error" and res["agent_directive"].startswith("STOP_RETRY:"), res   # 2.9.21: 1.3, ok
    assert "1.300" in res["message"]
    # the same numbers written the training way: read as 1300
    _, res = run_raw("ml_predict.py", "-m", vdir, "-f", _csv(d / "pred_ok.csv", [["온도", "유량"], ["175", "1.300,0"],
                                                                               ["180", "1.350,0"]]))
    assert res["status"] == "ok", res


def test_map_points_keep_data_decimal_comma(tmp_path):
    data = [["온도", "수율"]] + [[f"{170 + i},5", f"{80 + i},25"] for i in range(10)]
    path = _csv(tmp_path / "grid.csv", data)
    pts = _csv(tmp_path / "pts.csv", [["온도"], ["175"], ["170.500"]])                # 170.5 or 170500
    _, res = run_raw("ml_map.py", "-f", path, "--points", pts)
    assert res["status"] == "error" and res["agent_directive"].startswith("STOP_RETRY:"), res     # 2.9.21: ok
    pts = _csv(tmp_path / "pts2.csv", [["온도"], ["175"], ["176"]])
    _, res = run_raw("ml_map.py", "-f", path, "--points", pts)
    assert res["status"] == "ok", res
