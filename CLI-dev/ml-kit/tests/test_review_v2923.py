# -*- coding: utf-8 -*-
"""
2.9.23 (12차 점검) regressions.

1. SUM-M1/M2: 2.9.22 treated a column as a statistic label column (totals words not warned) when it held 2+ totals
   words in at least half of its filled cells, without any aggregate or average word. A merged 일자/라인 column
   (only the top cell filled) with 합계/소계 rows met that, so the totals rows were trained silently.
2. DC-L2: coerce_numeric stripped every comma, so a text cell '175,6' (decimal comma typed in an .xlsx) became 1756.
   Only real thousands groups ('1,500', '1,234,567.5') are stripped now; '175,6' is an unreadable value.
"""
import os
import sys
from datetime import datetime

import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from test_layout_matrix_v2917 import mp, _write  # noqa: E402


def _summary(path):
    df = mp.load_data_file(path)
    return [(r[0], r[2]) for r in df.attrs["summary_rows"]]


def test_merged_date_column_totals_warned(tmp_path):
    rows = [["일자", "교대", "온도", "생산량"]]
    for d in range(1, 11):
        rows += [[datetime(2026, 10, d), "주간", 150 + d, 300 + 2 * d], [None, "야간", 160 + d, 320 + 2 * d],
                 ["합계", None, 310 + 2 * d, 620 + 4 * d]]
    got = _summary(_write(str(tmp_path / "d2.xlsx"), rows, []))
    assert len(got) == 10 and all(w == "합계" for _, w in got), got


def test_merged_line_column_subtotals_warned(tmp_path):
    rows = [["라인", "온도", "압력", "생산량"]]
    for k in range(1, 6):
        rows += [[f"L{k}", 160 + k, 3.0 + k / 10, 110 + k], [None, 170 + k, 3.2 + k / 10, 120 + k],
                 ["소계", 330 + 2 * k, 6.2 + k / 5, 230 + 2 * k]]
    got = _summary(_write(str(tmp_path / "m2.xlsx"), rows, []))
    assert len(got) == 5 and all(w == "소계" for _, w in got), got


def test_statistic_label_column_still_quiet(tmp_path):
    rows = [["집계", "x1", "x2"]]
    for i in range(12):
        rows.append([["합계", "평균", "최대", "최소"][i % 4], i, i * 2])
    assert _summary(_write(str(tmp_path / "s.xlsx"), rows, [])) == []


def test_coerce_numeric_only_strips_thousands_groups():
    s = pd.Series(["1,500", "1,234,567.5", "175,6", "1,2", " 2,000 ", "-1,000", "12"], dtype="object")
    got = mp.coerce_numeric(s).tolist()
    assert got[0] == 1500 and got[1] == 1234567.5 and got[4] == 2000 and got[5] == -1000 and got[6] == 12
    assert pd.isna(got[2]) and pd.isna(got[3])


# 2.9.24 (13차 점검) --------------------------------------------------------------------------------------------------
import random  # noqa: E402

from test_cli_smoke import run_raw  # noqa: E402


def test_merged_date_totals_with_bottom_max_min_warned(tmp_path):
    rows = [["일자", "교대", "온도", "생산량"]]
    for d in range(1, 11):
        rows += [[datetime(2026, 10, d), "주간", 150 + d, 300 + 2 * d], [None, "야간", 160 + d, 320 + 2 * d],
                 ["합계", None, 310 + 2 * d, 620 + 4 * d]]
    rows += [["평균", None, 160, 320], ["최대", None, 170, 340], ["최소", None, 151, 302]]
    got = _summary(_write(str(tmp_path / "d4.xlsx"), rows, []))
    assert sum(1 for _, w in got if w == "합계") == 10, got


def test_zero_group_comma_is_not_thousands():
    s = pd.Series(["0,149", "00,150", "-0,149", "1,500"], dtype="object")
    got = mp.coerce_numeric(s).tolist()
    assert pd.isna(got[0]) and pd.isna(got[1]) and pd.isna(got[2]) and got[3] == 1500


def test_training_warns_unparsed_number_cells(tmp_path):
    r = random.Random(1)
    rows = [["온도", "압력", "수율"]]
    for i in range(40):
        t = 150 + i
        rows.append([("175,6" if i in (3, 9, 20) else t), round(r.uniform(2, 4), 2), 0.5 * t + r.uniform(-1, 1)])
    f = _write(str(tmp_path / "a.xlsx"), rows, [])
    _, res = run_raw("ml_train.py", "-f", f, "-x", "온도,압력", "-y", "수율")
    assert res["status"] == "ok", res
    assert any("숫자 열 '온도'" in w and "175,6" in w for w in res.get("warnings", [])), res.get("warnings")
