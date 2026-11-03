# -*- coding: utf-8 -*-
"""
2.9.21 (10차 점검) regressions.

1. Average rows decided by values: a shift report 'A조, B조, C조, 평균' per day (평균 = 25% of the column) was read as a
   category and trained silently (SJ; with 소계 too only the 소계 rows were warned, SI). A 평균 row whose numbers equal
   the mean of the rows above it, or a 합계 row equal to their sum, is a summary row whatever its share.
2. False warnings: a long-format statistic label column ('집계' = 합계/평균/최대/최소 SB, 'Sum/Avg' SC) warned every
   totals row, and a column of 평균 in every row (SA) warned every row.
3. A ';' CSV with decimal commas ('175,5') read dot-grouped integers ('1.300') as 1.3 silently (E2): 2.9.21 read 1300
   with a warning when consistent; 2.9.22 refuses them always (test_review_v2922).
"""
import csv
import os
import sys

import numpy as np
import pytest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from test_layout_matrix_v2917 import mp, _write  # noqa: E402
from test_cli_smoke import run_raw  # noqa: E402


def _shift_rows(days=8, sub=False, avg=True, seed=5):
    r = np.random.default_rng(seed)
    rows = [["일자", "구분", "온도", "수율"]]
    for d in range(days):
        block = []
        for g in ["A조", "B조", "C조"]:
            t, y = round(float(r.uniform(170, 180)), 1), round(float(r.uniform(87, 94)), 1)
            block.append([t, y])
            rows.append([f"11/{d + 1}", g, t, y])
        if sub:
            rows.append([f"11/{d + 1}", "소계", round(sum(b[0] for b in block), 1), round(sum(b[1] for b in block), 1)])
        if avg:
            rows.append([f"11/{d + 1}", "평균", round(sum(b[0] for b in block) / 3, 1),
                         round(sum(b[1] for b in block) / 3, 1)])
    return rows


def _summary(path):
    df = mp.load_data_file(path)
    return [(r[0], r[2]) for r in df.attrs["summary_rows"]], mp.layout_warnings(df)


# 1. average / totals rows confirmed by values ---------------------------------------------------------------------

def test_shift_average_rows_warned(tmp_path):
    rows = _shift_rows()
    got, warns = _summary(_write(str(tmp_path / "sj.xlsx"), rows, []))
    assert got == [(i + 1, "평균") for i, r in enumerate(rows) if r[1] == "평균"]       # 2.9.20: none
    assert any("합계·평균" in w for w in warns)


def test_shift_subtotal_and_average_rows_warned(tmp_path):
    rows = _shift_rows(sub=True)
    got, _ = _summary(_write(str(tmp_path / "si.xlsx"), rows, []))
    assert got == [(i + 1, r[1]) for i, r in enumerate(rows) if r[1] in ("소계", "평균")]  # 2.9.20: 소계 only


def test_average_category_values_not_block_means(tmp_path):
    # '평균' is a measurement kind here (25%): its values are not the mean of the rows above -> no warning
    r = np.random.default_rng(7)
    rows = [["구분", "온도", "수율"]]
    for i in range(24):
        rows.append(["평균" if i % 4 == 3 else "측정", round(float(r.uniform(150, 200)), 1),
                     round(float(r.uniform(60, 95)), 1)])
    got, _ = _summary(_write(str(tmp_path / "cat.xlsx"), rows, []))
    assert got == []


# 2. statistic label columns ---------------------------------------------------------------------------------------

def _stat_rows(words, seed=9, label="집계", day=True):
    r = np.random.default_rng(seed)
    rows = [(["일자"] if day else []) + [label, "온도", "수율"]]
    for d in range(6):
        for w in words:
            rows.append(([f"11/{d + 1}"] if day else []) + [w, round(float(r.uniform(172, 179)), 1),
                                                             round(float(r.uniform(88, 92)), 1)])
    return rows


def test_long_format_statistics_not_warned(tmp_path):
    sb = _stat_rows(["합계", "평균", "최대", "최소"])
    assert _summary(_write(str(tmp_path / "sb.xlsx"), sb, []))[0] == []                # 2.9.20: every 합계 row
    sc = _stat_rows(["Sum", "Avg"], label="통계", day=False)
    assert _summary(_write(str(tmp_path / "sc.xlsx"), sc, []))[0] == []                # 2.9.20: every Sum row
    sa = _stat_rows(["평균"], label="통계", day=False)
    assert _summary(_write(str(tmp_path / "sa.xlsx"), sa, []))[0] == []                # 2.9.20: every row


def test_statistic_column_total_not_value_checked(tmp_path):
    # a 집계 column (statistic words in half of its cells): 2.9.22 does not check totals by value any more (the 2.9.21
    # sum check missed ratio columns); its 합계 rows are statistic labels, not warned
    rows = [["집계", "온도", "수율"], ["최대", 10.0, 5.0], ["측정", 11.5, 6.0], ["측정", 12.0, 7.5], ["측정", 13.0, 8.0],
            ["합계", 36.5, 21.5], ["최소", 9.0, 4.0]]
    got, _ = _summary(_write(str(tmp_path / "sbt.xlsx"), rows, []))
    assert got == []                                                                    # 2.9.21: [(6, "합계")]


def test_shift_report_trains_with_warning(tmp_path):
    path = _write(str(tmp_path / "sj_train.xlsx"), _shift_rows(days=10), [])
    _, res = run_raw("ml_train.py", "-f", path, "-x", "온도,구분", "-y", "수율")
    assert res["status"] == "ok" and any("합계·평균" in w for w in res.get("warnings", [])), res


# 3. dot-grouped integers in a decimal-comma CSV --------------------------------------------------------------------

def _csv(path, rows):
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        csv.writer(f, delimiter=";").writerows(rows)
    return str(path)


def _e2_rows():
    rows = [["온도", "유량", "수율"]]
    for i in range(12):
        rows.append([f"{175.5 + i:.1f}".replace(".", ","), f"1.{300 + 7 * i:03d}", f"{90.1 + i:.1f}".replace(".", ",")])
    return rows


def test_dot_grouped_integers_refused(tmp_path):
    # 2.9.22: never guessed (2.9.20: 1.3, 1.307, ...; 2.9.21: 1300 with a warning)
    with pytest.raises(mp.InputError, match="1.300"):
        mp.load_data_file(_csv(tmp_path / "e2.csv", _e2_rows()))


def test_dot_grouped_integers_ambiguous_refused(tmp_path):
    # a plain decimal-point number in the file: '1.300' could be 1.3 too -> refused
    rows = _e2_rows()
    rows[3][2] = "1.5"
    with pytest.raises(mp.InputError):
        mp.load_data_file(_csv(tmp_path / "e2a.csv", rows))
    # '0.300' cannot be a thousands group
    rows = _e2_rows()
    rows[4][1] = "0.300"
    with pytest.raises(mp.InputError):
        mp.load_data_file(_csv(tmp_path / "e2b.csv", rows))
    # another dotted cell in the same column
    rows = _e2_rows()
    rows[2][1] = "L1.300x"
    with pytest.raises(mp.InputError):
        mp.load_data_file(_csv(tmp_path / "e2c.csv", rows))
    _, res = run_raw("ml_train.py", "-f", _csv(tmp_path / "e2t.csv", _e2_rows()[:1] + [
        r if i != 3 else [r[0], r[1], "1.5"] for i, r in enumerate(_e2_rows()[1:])]), "-x", "온도,유량", "-y", "수율")
    assert res["status"] == "error" and res["agent_directive"].startswith("STOP_RETRY:"), res


def test_dot_grouped_integers_train_refused(tmp_path):
    _, res = run_raw("ml_train.py", "-f", _csv(tmp_path / "e2t.csv", _e2_rows()), "-x", "온도,유량", "-y", "수율")
    assert res["status"] == "error" and res["agent_directive"].startswith("STOP_RETRY:"), res   # 2.9.21: ok
