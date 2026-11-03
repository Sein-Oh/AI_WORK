# -*- coding: utf-8 -*-
"""
2.9.27 (16차 점검) regressions.

1. 2.9.26 counted loose totals words ('일합계', '총 합계', '주간계', 'Grand Total') only when they were at most half of
   the filled cells of the column. A daily report with a merged date column has more '일합계' rows (one per day plus
   the last '월합계' / '총 합계') than dates, so none of its totals rows was warned; a blank label column with one
   '총 합계' was not warned either. Loose totals words are now judged by the row's numbers: the sums of the data rows
   directly above, of the earlier summary rows of one text ('월합계' = the '일합계' rows) or of all data rows above,
   or (few loose rows between the data) numbers far above the data. Categories ('생산계'/'품질팀', 'pH계' among
   sensors, 'A계'/'B계'/'공통', '레벨계') hold numbers like the data rows: quiet (2.9.26 warned them when they were
   under half of the column).
2. -i 'k=v': a repeated name was silently overwritten by the later value, and '온도' next to '온도(°C)' (the same
   column) silently used one of the two values. Both stop (STOP_RETRY).
"""
import datetime
import os
import random
import sys

import numpy as np
import pandas as pd
import pytest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
sys.path.insert(0, os.path.join(os.path.dirname(HERE), "tools"))
from test_layout_matrix_v2917 import mp, _write  # noqa: E402
from test_cli_smoke import run_raw  # noqa: E402


def _summary(tmp_path, name, rows, merges=()):
    return mp.load_data_file(_write(str(tmp_path / name), rows, list(merges))).attrs["summary_rows"]


def _daily(word, last=None, shifts=("주간", "야간"), days=10, datecell=lambda d: f"10/{d}", merge=True):
    """Merged-date daily report: shift rows, a '<word>' row with the day's sums, an optional last row (9999)."""
    rows, merges = [["일자", "교대", "온도", "생산량"]], []
    for d in range(1, days + 1):
        r0, tot = len(rows) + 1, [0, 0]
        for j, s in enumerate(shifts):
            a = [150 + d + j, 300 + d + j]
            tot = [tot[0] + a[0], tot[1] + a[1]]
            rows.append([datecell(d) if j == 0 else None, s, *a])
        if merge and len(shifts) > 1:
            merges.append(f"A{r0}:A{r0 + len(shifts) - 1}")
        rows.append([word, None, *tot])
    if last:
        rows.append([last, None, 9999, 99999])
    return rows, merges


@pytest.mark.parametrize("word,last,kw", [
    ("일합계", "월합계", {}),                                                          # A1
    ("일합계", "총 합계", {}),                                                         # A2
    ("일합계", "월합계", {"shifts": ("A조", "B조", "C조")}),                            # A3
    ("일합계", "월합계", {"datecell": lambda d: datetime.datetime(2026, 10, d)}),        # A4
    ("주간계", "월간계", {}),                                                          # A8
    ("일합계", "전체합계", {}),
])
def test_merged_date_daily_report_loose_totals_warned(tmp_path, word, last, kw):
    rows, merges = _daily(word, last, **kw)
    got = _summary(tmp_path, "a.xlsx", rows, merges)
    assert len(got) == 11 and {r[2] for r in got} == {word, last}, got


def test_monthly_total_is_sum_of_daily_totals(tmp_path):
    # '월합계' holding the sums of the '일합계' rows (not of the data rows above it): still a totals row
    rows, merges = _daily("일합계")
    day = [r for r in rows if r[0] == "일합계"]
    rows.append(["월합계", None, sum(r[2] for r in day), sum(r[3] for r in day)])
    rows.append(["10/31", "주간", 160, 320])
    got = _summary(tmp_path, "m.xlsx", rows, merges)
    assert len(got) == 11 and got[-1][2] == "월합계", got


def test_weekly_totals_in_shift_column_warned(tmp_path):
    # B1: '주간계' (a number far above the data) and '월간계' (= the two '주간계') in the shift column
    rows = [["일자", "근무조", "온도", "생산량"]]
    for d in range(1, 15):
        rows += [[f"10/{d}", s, 170 + d, 300 + d] for s in ("A조", "B조", "C조")]
        if d % 7 == 0:
            rows.append([None, "주간계", None, 2100])
    rows.append([None, "월간계", None, 4200])
    got = _summary(tmp_path, "b.xlsx", rows)
    assert [r[2] for r in got] == ["주간계", "주간계", "월간계"], got


def test_note_column_and_blank_label_column_total_warned(tmp_path):
    rows = [["No", "비고", "온도", "생산량"]]                      # D1: one note and '총 합계' in the note column
    rows += [[i + 1, "점검" if i == 3 else None, 170 + i, 300 + i] for i in range(12)]
    rows.append([None, "총 합계", None, 3666])
    assert [r[2] for r in _summary(tmp_path, "d1.xlsx", rows)] == ["총 합계"]
    rows = [["라벨", "온도", "생산량"]] + [[None, 170 + i, 300 + i] for i in range(12)] + [["총 합계", None, 3666]]
    assert [r[2] for r in _summary(tmp_path, "d2.xlsx", rows)] == ["총 합계"]       # D2: only the total filled


@pytest.mark.parametrize("n", [4, 5, 6])
@pytest.mark.parametrize("blank_dates", [False, True])
def test_tiny_file_total_warned(tmp_path, n, blank_dates):
    # E4..E6 and E4b..E6b: a few data rows and a '총 합계' whose number is not their sum (999)
    rows = [["일자", "온도", "생산량"]]
    rows += [[None if blank_dates else f"10/{i + 1}", 170 + i, 300 + i] for i in range(n - 1)]
    rows.append(["총 합계", None, 999])
    assert [r[2] for r in _summary(tmp_path, "e.xlsx", rows)] == ["총 합계"]


def _category(values, seed=1):
    r = random.Random(seed)
    rows = [["구분", "온도", "압력", "수율"]]
    for v in values:
        rows.append([v, round(r.uniform(175, 185), 1), round(r.uniform(2, 3), 2), round(r.uniform(50, 60), 2)])
    return rows


@pytest.mark.parametrize("values", [
    ["A계"] * 10 + ["B계"] * 10,                         # C1
    ["생산계"] * 9 + ["품질팀"] * 11,                    # C2
    ["pH계", "온도센서", "압력센서", "DO센서"] * 5,       # C3
    ["A계"] * 4 + ["B계"] * 4 + ["공통"] * 12,          # C4
    ["레벨계", "온도센서", "압력센서"] * 7,               # C5
    ["생산계", "품질팀", "품질팀", "품질팀"] * 10,
])
def test_loose_category_values_quiet(tmp_path, values):
    assert _summary(tmp_path, "c.xlsx", _category(values)) == []


@pytest.mark.parametrize("name,values", [
    ("라인", ["A계", "B계"]), ("계통", ["1계", "2계"]), ("담당", ["생산계", "품질계", "설비계"]),
    ("계기", ["레벨계", "pH계", "유량계"]),
])
def test_line_dept_gauge_categories_quiet(tmp_path, name, values):
    # L1/L2/L3/L6 (r15)
    rng = np.random.default_rng(0)
    t, p = rng.normal(180, 5, 40).round(1), rng.normal(3, 0.3, 40).round(2)
    df = pd.DataFrame({name: rng.choice(values, 40), "온도": t, "압력": p, "수율": (0.3 * t + 5 * p).round(2)})
    path = str(tmp_path / "l.xlsx")
    df.to_excel(path, index=False)
    assert mp.load_data_file(path).attrs["summary_rows"] == []


@pytest.mark.parametrize("word", ["일계", "총 합계", "일합계", "전체합계", "합계:", "Total:", "소계(A)", "주간계",
                                  "Grand Total", "Sub Total", "합계(일)"])
def test_daily_totals_with_sums_warned(tmp_path, word):
    # r14 words.py / W_일계 / W_총합계: day rows + their sum, date column not merged
    rows = [["일자", "교대", "온도", "생산량"]]
    for d in range(1, 8):
        a, b = [150 + d, 300 + d], [160 + d, 310 + d]
        rows += [[f"10/{d}", "주간", *a], [None, "야간", *b], [word, None, a[0] + b[0], a[1] + b[1]]]
    got = _summary(tmp_path, "w.xlsx", rows)
    assert len(got) == 7 and {r[2] for r in got} == {word}, got


def test_single_shift_daily_total_copy_quiet(tmp_path):
    # one data row per day: '일합계' equals that row (one row is its own sum), not judged as a totals row
    rows, merges = _daily("일합계", shifts=("주간",))
    assert _summary(tmp_path, "s.xlsx", rows, merges) == []


def test_merged_date_daily_report_train_warned(tmp_path):
    rows, merges = _daily("일합계", "월합계")
    f = _write(str(tmp_path / "t.xlsx"), rows, merges)
    _, res = run_raw("ml_train.py", "-f", f, "-x", "온도", "-y", "생산량")
    assert res["status"] == "ok", res
    assert any("합계·평균 행" in w and "외 6개" in w for w in res.get("warnings", [])), res.get("warnings")


def test_kv_repeated_name_refused():
    with pytest.raises(mp.InputError) as e:
        mp.parse_inline_input("온도=195,압력=3.8;온도=165,압력=2.1")
    assert "두 번" in str(e.value) and "JSON 배열" in str(e.value)
    with pytest.raises(mp.InputError):
        mp.parse_inline_input("온도=195, 온도=165")


def test_kv_alias_and_full_name_refused():
    with pytest.raises(mp.InputError) as e:
        mp.parse_inline_input("온도=180,온도(°C)=190")
    assert "[후보] 온도=180, 온도(°C)=190" in str(e.value)
    # two different units are two columns; a plain row still parses
    assert list(mp.parse_inline_input("온도(°C)=180,온도(K)=453").columns) == ["온도(°C)", "온도(K)"]
    assert mp.parse_inline_input("라인=A,B,온도=180").iloc[0].to_dict() == {"라인": "A,B", "온도": "180"}


def test_kv_repeated_name_cli_stop_retry(tmp_path):
    rows = [["온도", "압력", "수율"]] + [[170 + i, 2 + i / 10, 50 + i] for i in range(12)]
    f = _write(str(tmp_path / "k.xlsx"), rows, [])
    proc, res = run_raw("ml_map.py", "-f", f, "-i", "온도=175,온도=176", "-y", "수율")
    assert proc.returncode == 1 and res["status"] == "error", res
    assert res["agent_directive"].startswith("STOP_RETRY:"), res
