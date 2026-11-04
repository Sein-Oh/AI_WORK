# -*- coding: utf-8 -*-
"""데이터 읽기 2: 합계·소계·평균 행(summary_rows) 경고.

합계 글자(합계/소계/계, 일계/월계/누계, Total/Sum)가 있는 행, 그리고 넓힌 글자(일합계, 총 합계, 주간계, 평균값)는 행의 숫자로
판단한다: 위 데이터 행들의 합, 위 합계 행들의 합, 블록 평균(변하는 열만). 범주 값(A계, 생산계, pH계)과 통계 이름 열
(집계 = 합계/평균/최대/최소)은 경고하지 않는다. '제외' 열에 표시한 행은 경고하지 않는다.
알려진 거짓 경고 하나는 그대로 둔다 (test_known_false_warning_small_integer_sum): 합계 행을 놓쳐 학습에 들어가는 것이 더 나쁘다.
"""
import os
import random
import sys
import time
from datetime import datetime

import numpy as np
import pandas as pd
import pytest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from mlkit_helpers import H, U, family_rows, run_raw, write_csv, write_rows  # noqa: E402
import ml_predict as mp  # noqa: E402

N = 40
X, Y = "반응온도,압력,라인", "수율"


# ---- 읽기 도구 ------------------------------------------------------------------------------------------------------


def _summary_warn(path):
    df = mp.load_data_file(path)
    return [(r[0], r[2]) for r in df.attrs["summary_rows"]], mp.layout_warnings(df)


def _summary_pairs(path):
    df = mp.load_data_file(path)
    return [(r[0], r[2]) for r in df.attrs["summary_rows"]]


def _summary(tmp_path, name, rows, merges=()):
    return mp.load_data_file(write_rows(str(tmp_path / name), rows, list(merges))).attrs["summary_rows"]


def _pairs(df):
    return [(r[0], r[2]) for r in df.attrs["summary_rows"]]


# ---- '제외' 열 표시와 집계 범주 --------------------------------------------------------------------------------


def _totals_book(path, mark):
    data = family_rows("A", 20, 1)
    tot = ["합계", round(sum(r[0] for r in data), 1), round(sum(r[1] for r in data), 2), None,
           round(sum(r[3] for r in data), 2)]
    rows = [["구분"] + H + ["제외"], [None] + U + [None]] + [[i + 1] + r + [None] for i, r in enumerate(data)]
    rows.append(tot + [mark])
    return write_rows(path, rows, [])


@pytest.mark.parametrize("mark,warn", [("합계", False), ("합계 행", False), (None, True), ("포함", True)])
def test_totals_row_marked_in_exclude_column(tmp_path, mark, warn):
    path = _totals_book(str(tmp_path / "a1.xlsx"), mark)
    df = mp.load_data_file(path, unit_row=2)
    assert bool(df.attrs["summary_rows"]) is warn, df.attrs["summary_rows"]    # 2.9.18: warned for every mark
    assert all(r[1] != "제외" for r in df.attrs["summary_rows"])


def test_totals_row_marked_train_and_map(tmp_path):
    path = _totals_book(str(tmp_path / "a1.xlsx"), "합계")
    _, res = run_raw("ml_train.py", "-f", path, "-x", X, "-y", Y, "--unit-row", "2")
    assert res["status"] == "ok", res
    assert not any("합계" in w for w in res.get("warnings", [])), res.get("warnings")
    _, res = run_raw("ml_map.py", "-f", path, "--unit-row", "2", "-i", "반응온도(°C)=180", "-y", "수율(%)")
    assert res["status"] == "ok", res
    assert not any("합계·평균" in w for w in res.get("warnings", [])), res.get("warnings")


def test_summary_rows_keep_exclude_marks():
    data = family_rows("A", 6, 3)
    df = pd.DataFrame([["소계"] + r + ["합계"] for r in data[:2]] + [["A"] + r + [None] for r in data[2:]],
                      columns=["구분"] + H + ["제외"])
    assert mp._summary_rows(df, list(range(2, 8))) == []
    df["제외"] = [None] * 6
    assert [r[0] for r in mp._summary_rows(df, list(range(2, 8)))] == [2, 3]


def test_aggregation_category_is_not_a_totals_row():
    data = family_rows("A", 20, 3)
    agg = pd.DataFrame([r + ["평균" if i % 2 else "최대"] for i, r in enumerate(data)], columns=H + ["집계"])
    assert mp._summary_rows(agg, list(range(3, 23))) == []                         # 2.9.18: 10 rows warned
    agg["집계"] = ["평균" if i < 3 else "최대" for i in range(20)]                  # a few, next to '최대'
    assert mp._summary_rows(agg, list(range(3, 23))) == []
    rows = []
    for g in ("A", "B"):                                                          # subtotals per group still warn
        sub = [r for r in data if r[2] == g]
        rows += [[g] + r for r in sub] + [["소계", 170.0, 3.0, g, 70.0]]
    sub = pd.DataFrame(rows, columns=["구분"] + H)
    assert [r[2] for r in mp._summary_rows(sub, list(range(3, 3 + len(rows))))] == ["소계", "소계"]


def test_average_category_still_not_warned():
    data = family_rows("A", 20, 3)
    agg = pd.DataFrame([r + ["평균" if i % 2 else "최대"] for i, r in enumerate(data)], columns=H + ["집계"])
    assert mp._summary_rows(agg, list(range(3, 23))) == []
    agg["집계"] = ["평균" if i % 4 == 0 else "측정" for i in range(20)]        # 25%: a category of the column
    assert mp._summary_rows(agg, list(range(3, 23))) == []
    agg["집계"] = ["평균" if i < 3 else "최대" for i in range(20)]
    assert mp._summary_rows(agg, list(range(3, 23))) == []


# ---- 소계·합계 행 (블록의 합) ----------------------------------------------------------------------------------


def test_shift_subtotal_rows_warned(tmp_path):
    r = np.random.default_rng(5)
    rows = [["일자", "구분", "반응온도", "압력", "생산량"]]
    for d in range(12):
        day = [[f"2026-10-{d + 1:02d}", s, round(float(r.uniform(150, 200)), 1), round(float(r.uniform(1, 5)), 2),
                round(float(r.uniform(300, 450)), 1)] for s in ("주간", "야간", "심야")]
        rows += day + [[f"2026-10-{d + 1:02d}", "소계"] + [round(sum(x[c] for x in day), 2) for c in (2, 3, 4)]]
    path = write_rows(str(tmp_path / "s1.xlsx"), rows, [])
    df = mp.load_data_file(path)
    assert _pairs(df) == [(5 + 4 * d, "소계") for d in range(12)]           # 2.9.19: [] (25% of the column)
    _, res = run_raw("ml_train.py", "-f", path, "-x", "반응온도,압력", "-y", "생산량")
    assert res["status"] == "ok" and any("합계·평균" in w for w in res.get("warnings", [])), res


def test_two_lines_and_total_rows_warned(tmp_path):
    r = np.random.default_rng(5)
    rows = [["구분", "반응온도", "압력", "수율"]]
    for _ in range(15):
        a = [["1라인", round(float(r.uniform(150, 200)), 1), round(float(r.uniform(1, 5)), 2),
              round(float(r.uniform(80, 95)), 1)] for _ in range(2)]
        rows += a + [["합계"] + [a[0][c] + a[1][c] for c in (1, 2, 3)]]
    df = mp.load_data_file(write_rows(str(tmp_path / "s3.xlsx"), rows, []))
    assert _pairs(df) == [(4 + 3 * k, "합계") for k in range(15)]           # 2.9.19: [] (33%)


def test_lone_total_in_blank_column_warned(tmp_path):
    d = family_rows("A", N, 1)
    tot = ["합계", round(sum(x[0] for x in d), 1), round(sum(x[1] for x in d), 2), None, round(sum(x[3] for x in d), 2)]
    df = mp.load_data_file(write_rows(str(tmp_path / "s2.xlsx"), [["비고"] + H] + [[None] + x for x in d] + [tot], []))
    assert _pairs(df) == [(N + 2, "합계")]                                  # 2.9.19: [] (1 of 1 filled cells)
    avg = ["평균"] + tot[1:]
    df = mp.load_data_file(write_rows(str(tmp_path / "s2b.xlsx"), [["비고"] + H] + [[None] + x for x in d] + [avg], []))
    assert _pairs(df) == [(N + 2, "평균")]


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
    got = _summary_pairs(write_rows(str(tmp_path / "b5.xlsx"), head + data + _bottom(data), []))
    assert (len(data) + 2, "합계") in got                                     # 2.9.21: [] (ratio column != sum)
    data = _lines(ratio=False)
    got = _summary_pairs(write_rows(str(tmp_path / "b1n.xlsx"), [head[0][:4]] + data + _bottom(data, False), []))
    assert (len(data) + 2, "합계") in got


def test_subtotal_groups_and_grand_total_warned(tmp_path):
    rows = [["라인", "온도", "생산량", "불량수"]]
    data = _lines(ratio=False)
    for line in ["L1", "L2", "L3", "L4"]:
        part = [d for d in data if d[0] == line]
        rows += part + [["소계", None, sum(d[2] for d in part), sum(d[3] for d in part)]]
    rows += [["합계", None, sum(d[2] for d in data), sum(d[3] for d in data)]] + _bottom(data, False)[1:]
    got = _summary_pairs(write_rows(str(tmp_path / "b2.xlsx"), rows, []))
    want = [i + 1 for i, r in enumerate(rows) if r[0] in ("소계", "합계")]
    assert [g[0] for g in got if g[1] in ("소계", "합계")] == want                # 2.9.21: grand 합계 not named


def test_subtotal_in_average_category_column_warned(tmp_path):
    r = np.random.default_rng(6)
    rows = [["측정방식", "온도", "생산량"]]
    for _ in range(5):
        block = [[w, round(float(r.uniform(172, 184)), 1), int(r.integers(900, 1100))] for w in ["개별", "평균"] * 2]
        rows += block + [["소계", None, sum(b[2] for b in block)]]
    got = _summary_pairs(write_rows(str(tmp_path / "b3.xlsx"), rows, []))
    assert got == [(i + 1, "소계") for i, r in enumerate(rows) if r[0] == "소계"]    # 2.9.21: []


def test_lone_totals_in_note_column_warned(tmp_path):
    # a mostly blank note column holding only 소계/합계: not a statistic label column
    rows = [["비고", "생산량", "불량수"]]
    data = [[None, 1000 + 7 * i, 10 + i] for i in range(6)]
    rows += data[:3] + [["소계", sum(d[1] for d in data[:3]), sum(d[2] for d in data[:3])]]
    rows += data[3:] + [["소계", sum(d[1] for d in data[3:]), sum(d[2] for d in data[3:])]]
    rows += [["합계", sum(d[1] for d in data), sum(d[2] for d in data)]]
    assert [g[1] for g in _summary_pairs(write_rows(str(tmp_path / "note.xlsx"), rows, []))] == ["소계", "소계", "합계"]


def test_merged_date_column_totals_warned(tmp_path):
    rows = [["일자", "교대", "온도", "생산량"]]
    for d in range(1, 11):
        rows += [[datetime(2026, 10, d), "주간", 150 + d, 300 + 2 * d], [None, "야간", 160 + d, 320 + 2 * d],
                 ["합계", None, 310 + 2 * d, 620 + 4 * d]]
    got = _summary_pairs(write_rows(str(tmp_path / "d2.xlsx"), rows, []))
    assert len(got) == 10 and all(w == "합계" for _, w in got), got


def test_merged_line_column_subtotals_warned(tmp_path):
    rows = [["라인", "온도", "압력", "생산량"]]
    for k in range(1, 6):
        rows += [[f"L{k}", 160 + k, 3.0 + k / 10, 110 + k], [None, 170 + k, 3.2 + k / 10, 120 + k],
                 ["소계", 330 + 2 * k, 6.2 + k / 5, 230 + 2 * k]]
    got = _summary_pairs(write_rows(str(tmp_path / "m2.xlsx"), rows, []))
    assert len(got) == 5 and all(w == "소계" for _, w in got), got


def test_merged_date_totals_with_bottom_max_min_warned(tmp_path):
    rows = [["일자", "교대", "온도", "생산량"]]
    for d in range(1, 11):
        rows += [[datetime(2026, 10, d), "주간", 150 + d, 300 + 2 * d], [None, "야간", 160 + d, 320 + 2 * d],
                 ["합계", None, 310 + 2 * d, 620 + 4 * d]]
    rows += [["평균", None, 160, 320], ["최대", None, 170, 340], ["최소", None, 151, 302]]
    got = _summary_pairs(write_rows(str(tmp_path / "d4.xlsx"), rows, []))
    assert sum(1 for _, w in got if w == "합계") == 10, got


# ---- 평균 행 (블록의 평균) --------------------------------------------------------------------------------------


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


def test_shift_average_rows_warned(tmp_path):
    rows = _shift_rows()
    got, warns = _summary_warn(write_rows(str(tmp_path / "sj.xlsx"), rows, []))
    assert got == [(i + 1, "평균") for i, r in enumerate(rows) if r[1] == "평균"]       # 2.9.20: none
    assert any("합계·평균" in w for w in warns)


def test_shift_subtotal_and_average_rows_warned(tmp_path):
    rows = _shift_rows(sub=True)
    got, _ = _summary_warn(write_rows(str(tmp_path / "si.xlsx"), rows, []))
    assert got == [(i + 1, r[1]) for i, r in enumerate(rows) if r[1] in ("소계", "평균")]  # 2.9.20: 소계 only


def test_average_category_values_not_block_means(tmp_path):
    # '평균' is a measurement kind here (25%): its values are not the mean of the rows above -> no warning
    r = np.random.default_rng(7)
    rows = [["구분", "온도", "수율"]]
    for i in range(24):
        rows.append(["평균" if i % 4 == 3 else "측정", round(float(r.uniform(150, 200)), 1),
                     round(float(r.uniform(60, 95)), 1)])
    got, _ = _summary_warn(write_rows(str(tmp_path / "cat.xlsx"), rows, []))
    assert got == []


def test_shift_report_trains_with_warning(tmp_path):
    path = write_rows(str(tmp_path / "sj_train.xlsx"), _shift_rows(days=10), [])
    _, res = run_raw("ml_train.py", "-f", path, "-x", "온도,구분", "-y", "수율")
    assert res["status"] == "ok" and any("합계·평균" in w for w in res.get("warnings", [])), res


def test_shift_average_with_thousands_commas_warned(tmp_path):
    r = np.random.default_rng(5)
    rows = [["날짜", "구분", "생산량", "불량수"]]
    for d in range(8):
        block = [[int(r.integers(2000, 3000)), int(r.integers(5000, 9000))] for _ in range(3)]
        for g, b in zip(["A조", "B조", "C조"], block):
            rows.append([f"2026-10-{d + 1:02d}", g, f"{b[0]:,}", f"{b[1]:,}"])
        rows.append([f"2026-10-{d + 1:02d}", "평균", f"{round(sum(b[0] for b in block) / 3):,}",
                     f"{round(sum(b[1] for b in block) / 3):,}"])
    got = _summary_pairs(write_csv(tmp_path / "t4.csv", rows, sep=","))
    assert got == [(i + 1, "평균") for i, r in enumerate(rows) if r[1] == "평균"]   # 2.9.21: []


# ---- 통계 이름 열 (집계 = 합계/평균/최대/최소) ------------------------------------------------------------


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
    assert _summary_warn(write_rows(str(tmp_path / "sb.xlsx"), sb, []))[0] == []                # 2.9.20: every 합계 row
    sc = _stat_rows(["Sum", "Avg"], label="통계", day=False)
    assert _summary_warn(write_rows(str(tmp_path / "sc.xlsx"), sc, []))[0] == []                # 2.9.20: every Sum row
    sa = _stat_rows(["평균"], label="통계", day=False)
    assert _summary_warn(write_rows(str(tmp_path / "sa.xlsx"), sa, []))[0] == []                # 2.9.20: every row


def test_statistic_column_total_not_value_checked(tmp_path):
    # a 집계 column (statistic words in half of its cells): 2.9.22 does not check totals by value any more (the 2.9.21
    # sum check missed ratio columns); its 합계 rows are statistic labels, not warned
    rows = [["집계", "온도", "수율"], ["최대", 10.0, 5.0], ["측정", 11.5, 6.0], ["측정", 12.0, 7.5], ["측정", 13.0, 8.0],
            ["합계", 36.5, 21.5], ["최소", 9.0, 4.0]]
    got, _ = _summary_warn(write_rows(str(tmp_path / "sbt.xlsx"), rows, []))
    assert got == []                                                                    # 2.9.21: [(6, "합계")]


def test_statistic_label_columns_stay_quiet(tmp_path):
    r = np.random.default_rng(9)
    for name, words in [("sb", ["합계", "평균", "최대", "최소"]), ("sc", ["Sum", "Avg"]), ("sa", ["평균"])]:
        rows = [["집계", "온도", "수율"]] + [[w, round(float(r.uniform(172, 179)), 1), round(float(r.uniform(88, 92)), 1)]
                                         for _ in range(6) for w in words]
        assert _summary_pairs(write_rows(str(tmp_path / f"{name}.xlsx"), rows, [])) == [], name


def test_statistic_label_column_still_quiet(tmp_path):
    rows = [["집계", "x1", "x2"]]
    for i in range(12):
        rows.append([["합계", "평균", "최대", "최소"][i % 4], i, i * 2])
    assert _summary_pairs(write_rows(str(tmp_path / "s.xlsx"), rows, [])) == []


def _stat_file(path, n):
    r = np.random.default_rng(3)
    words = ["합계", "평균", "최대", "최소"]
    rows = [["집계"] + [f"x{c}" for c in range(6)]]
    for i in range(n):
        rows.append([words[i % 4]] + [f"{v:.3f}" for v in r.normal(0, 2, 6)])
    return write_csv(path, rows, sep=",")


def test_statistic_column_load_linear(tmp_path):
    # 2.9.21: 30k rows 11s, 100k rows 111s. A generous bound for a loaded machine that still catches quadratic growth.
    small, big = _stat_file(tmp_path / "s.csv", 5000), _stat_file(tmp_path / "b.csv", 40000)
    mp.load_data_file(small)
    t0 = time.perf_counter()
    assert _summary_pairs(small) == []
    t_small = time.perf_counter() - t0
    t0 = time.perf_counter()
    assert _summary_pairs(big) == []
    t_big = time.perf_counter() - t0
    # 8x rows: linear ~8x (2.9.22: 0.03s -> 0.13s), quadratic ~64x (2.9.21: 0.3s -> about 18s)
    assert t_big < max(16 * t_small, 2.0), (t_small, t_big)


# ---- 합계 글자: 정확한 글자와 넓힌 글자 ------------------------------------------------------------------------


# 일일 보고서 글자 9개(일계, 총 합계, 전체합계, 합계:, Total:, 소계(A), 주간계, 일합계, Grand Total)는
# test_daily_totals_with_sums_warned가 같은 입력으로 더 강하게 확인한다. 여기는 그 목록에 없는 글자만.


@pytest.mark.parametrize("word", ["월계", "누계"])
def test_daily_report_totals_wording_warned(tmp_path, word):
    rows = [["일자", "교대", "온도", "생산량"]]
    for d in range(1, 8):
        a, b = [150 + d, 300 + d], [160 + d, 310 + d]
        rows += [[f"10/{d}", "주간", *a], [None, "야간", *b], [word, None, a[0] + b[0], a[1] + b[1]]]
    got = mp.load_data_file(write_rows(str(tmp_path / "w.xlsx"), rows, [])).attrs["summary_rows"]
    assert len(got) == 7, (word, got)


@pytest.mark.parametrize("word", ["평균값", "AVG.", "AVG"])
def test_average_wording_warned(tmp_path, word):
    rows = [["일자", "교대", "온도", "생산량"]]
    for d in range(1, 8):
        rows += [[f"10/{d}", "주간", 150 + d, 300 + d], [None, "야간", 160 + d, 310 + d]]
    rows.append([word, None, 160, 310])
    got = mp.load_data_file(write_rows(str(tmp_path / "a.xlsx"), rows, [])).attrs["summary_rows"]
    assert [r[2] for r in got] == [word], got


def test_words_ending_with_gye_stay_categories(tmp_path):
    r = random.Random(2)
    rows = [["부서", "공정", "설비", "온도", "수율"]]
    for i in range(40):
        rows.append([r.choice(["설계", "계측", "생산", "품질"]), r.choice(["1단계", "2단계", "3단계"]),
                     r.choice(["온도계", "압력계", "유량계", "하계", "동계"]), 150 + i, 70 + i / 2])
    got = mp.load_data_file(write_rows(str(tmp_path / "c.xlsx"), rows, [])).attrs["summary_rows"]
    assert got == [], got


def test_word_kind_units():
    k = mp._word_kind
    assert [k(w) for w in ("일계", "년계", "Sub Total", "sum", "평균(A조)", "Mean:", "최대")] == \
        ["total", "total", "total", "total", "avg", "avg", "agg"]
    assert [k(w) for w in ("설계", "계측", "summary", "통계", "2단계", "A조")] == [""] * 6


def test_word_kind_split():
    k = mp._word_kind
    assert [k(w) for w in ("합계", "소계(A)", "Total:", "Sub Total", "일계", "누계", "평균", "AVG.")] == \
        ["total", "total", "total", "total", "total", "total", "avg", "avg"]
    assert [k(w) for w in ("총 합계", "A계", "주간계", "Grand Line Total", "Sum A")] == ["total_loose"] * 5
    assert [k(w) for w in ("평균값", "평균치")] == ["avg_loose"] * 2
    assert [k(w) for w in ("SUM-01", "Lot-Sum", "sum_a", "Total_Loss", "평균온도", "평균압력값", "설계")] == [""] * 7


def _category_rows(name, values, n=40, seed=1):
    r = random.Random(seed)
    rows = [[name, "온도", "압력", "수율"]]
    for i in range(n):
        t, p = round(r.uniform(175, 185), 1), round(r.uniform(2.5, 3.5), 2)
        rows.append([r.choice(values), t, p, round(0.3 * t + 2 * p + r.uniform(-1, 1), 2)])
    return rows


@pytest.mark.parametrize("name,values", [
    ("라인", ["A계", "B계"]),                       # L1
    ("계통", ["1계", "2계"]),                       # L2
    ("담당", ["생산계", "품질계", "설비계"]),         # L3
    ("계기", ["레벨계", "pH계", "유량계"]),           # L6
    ("설비", ["고온계", "저항계", "토크계", "소음계", "두께계"]),
    ("라인", ["A합계", "B합계"]),
    ("구분", ["Sum A", "Sum B", "Sum C"]),
])
def test_loose_words_of_category_column_quiet(tmp_path, name, values):
    assert _summary(tmp_path, "c.xlsx", _category_rows(name, values)) == []


def test_loose_words_as_half_of_full_column_quiet(tmp_path):
    # two loose texts and two ordinary words in turn: half of a fully filled column is a category
    rows = _category_rows("계기", ["x"])
    for i, r in enumerate(rows[1:]):
        r[0] = ["레벨계", "온도계", "pH계", "유량계"][i % 4]
    assert _summary(tmp_path, "h.xlsx", rows) == []


def test_category_column_train_has_no_totals_warning(tmp_path):
    f = write_rows(str(tmp_path / "l1.xlsx"), _category_rows("라인", ["A계", "B계"]), [])
    _, res = run_raw("ml_train.py", "-f", f, "-x", "라인,온도,압력", "-y", "수율")
    assert res["status"] == "ok", res
    assert not any("합계·평균 행" in w for w in res.get("warnings", [])), res["warnings"]


def _daily_rows(word, days=20):
    rows = [["일자", "교대", "온도", "생산량"]]
    for d in range(1, days + 1):
        a, b = [150 + d, 300 + d], [160 + d, 310 + d]
        rows += [[f"10/{d}", "주간", *a], [None, "야간", *b], [word, None, a[0] + b[0], a[1] + b[1]]]
    return rows


@pytest.mark.parametrize("word", ["일계", "총 합계", "소계(A)", "주간계", "Grand Total", "합계", "소계"])
def test_daily_report_totals_still_warned(tmp_path, word):
    got = _summary(tmp_path, "w.xlsx", _daily_rows(word))
    assert len(got) == 20 and {r[2] for r in got} == {word}, got


def test_daily_report_totals_train_warned(tmp_path):
    f = write_rows(str(tmp_path / "w.xlsx"), _daily_rows("총 합계"), [])
    _, res = run_raw("ml_train.py", "-f", f, "-x", "온도", "-y", "생산량")
    assert res["status"] == "ok", res
    assert any("합계·평균 행" in w for w in res.get("warnings", [])), res.get("warnings")


@pytest.mark.parametrize("word", ["합계", "소계"])
def test_exact_totals_words_unchanged_in_every_other_row(tmp_path, word):
    # exact words always count, even as half of the column (subtotal per line)
    rows = [["구분", "온도", "생산량"]]
    for i in range(10):
        rows += [[f"L{i}", 150 + i, 300 + i], [word, 150 + i, 300 + i]]
    got = _summary(tmp_path, "s.xlsx", rows)
    assert len(got) == 10, got


def test_loose_words_with_many_texts_are_a_category(tmp_path):
    # three different loose totals texts in one column with numbers like the data rows: a category, not summary rows
    rows = [["일자", "온도", "생산량"]]
    for d, w in enumerate(["A합계", "B합계", "C합계"] * 4):
        rows += [[f"10/{d + 1}", 150 + d, 300 + d], [f"10/{d + 1}b", 151 + d, 301 + d], [w, 152 + d, 302 + d]]
    assert _summary(tmp_path, "m.xlsx", rows) == []
    # 2.9.27: the same texts holding the sums of the two rows above are totals rows (judged by the numbers)
    for r in rows[3::3]:
        r[1], r[2] = 2 * r[1] - 3, 2 * r[2] - 3
    assert len(_summary(tmp_path, "m2.xlsx", rows)) == 12


# ---- 넓힌 합계 글자는 숫자로 판단 (일자 병합 일일 보고서) --------------------------------------------------


def _daily_merged(word, last=None, shifts=("주간", "야간"), days=10, datecell=lambda d: f"10/{d}", merge=True):
    """일자 병합 일일 보고서: 교대 행, 그날 합의 '<word>' 행, 끝에 선택 행 하나(9999)."""
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
    ("일합계", "월합계", {"datecell": lambda d: datetime(2026, 10, d)}),        # A4
    ("주간계", "월간계", {}),                                                          # A8
    ("일합계", "전체합계", {}),
])
def test_merged_date_daily_report_loose_totals_warned(tmp_path, word, last, kw):
    rows, merges = _daily_merged(word, last, **kw)
    got = _summary(tmp_path, "a.xlsx", rows, merges)
    assert len(got) == 11 and {r[2] for r in got} == {word, last}, got


def test_monthly_total_is_sum_of_daily_totals(tmp_path):
    # '월합계' holding the sums of the '일합계' rows (not of the data rows above it): still a totals row
    rows, merges = _daily_merged("일합계")
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
    rows, merges = _daily_merged("일합계", shifts=("주간",))
    assert _summary(tmp_path, "s.xlsx", rows, merges) == []


def test_merged_date_daily_report_train_warned(tmp_path):
    rows, merges = _daily_merged("일합계", "월합계")
    f = write_rows(str(tmp_path / "t.xlsx"), rows, merges)
    _, res = run_raw("ml_train.py", "-f", f, "-x", "온도", "-y", "생산량")
    assert res["status"] == "ok", res
    assert any("합계·평균 행" in w and "외 6개" in w for w in res.get("warnings", [])), res.get("warnings")


# ---- 평균 열 + 합 열의 합계 행, 변하지 않는 열 ------------------------------------------------------------


def _avg(rows, c):
    return round(sum(r[c] for r in rows) / len(rows), 2)


def _sum(rows, c):
    return round(sum(r[c] for r in rows), 2)


def _process_days(seed=7, n=30):
    r = random.Random(seed)
    return [[f"10/{i + 1}", round(r.uniform(175, 185), 1), round(r.uniform(2, 3), 2), round(r.uniform(40, 50), 1),
             r.randint(280, 320)] for i in range(n)]


H_DAYS = ["일자", "온도", "압력", "속도", "생산량"]


@pytest.mark.parametrize("word", ["총 합계", "합계", "전체합계", "Grand Total", "월합계"])
def test_totals_row_with_averaged_columns_warned(tmp_path, word):
    d = _process_days()
    rows = [H_DAYS] + d + [[word, _avg(d, 1), _avg(d, 2), _avg(d, 3), _sum(d, 4)]]
    got = _summary(tmp_path, "a.xlsx", rows)
    assert [r[0] for r in got] == [32] and got[0][2] == word, got


def test_totals_row_ratio_column_warned(tmp_path):
    r = random.Random(3)
    d = []
    for i in range(30):
        p, b = r.randint(280, 320), r.randint(0, 9)
        d.append([f"10/{i + 1}", round(r.uniform(175, 185), 1), p, b, round(b / p * 100, 2)])
    P, B = sum(x[2] for x in d), sum(x[3] for x in d)
    rows = [["일자", "온도", "생산량", "불량수", "불량률"]] + d + [["총 합계", _avg(d, 1), P, B, round(B / P * 100, 2)]]
    assert [x[0] for x in _summary(tmp_path, "r.xlsx", rows)] == [32]


@pytest.mark.parametrize("word,last", [("일합계", "월합계"), ("일계", "월합계"), ("주간계", "월간계"), ("합계", "총 합계")])
def test_daily_report_averaged_columns_all_warned(tmp_path, word, last):
    # merged-date daily report, 3 shifts: 온도/압력/속도 averaged, 생산량 summed, per day and in the last row
    r = random.Random(5)
    rows, merges, every = [["일자", "교대", "온도", "압력", "속도", "생산량"]], [], []
    for day in range(1, 11):
        r0, blk = len(rows) + 1, []
        for j, s in enumerate(["A조", "B조", "C조"]):
            x = [f"10/{day}" if j == 0 else None, s, round(r.uniform(175, 185), 1), round(r.uniform(2, 3), 2),
                 round(r.uniform(40, 50), 1), r.randint(280, 320)]
            rows.append(x)
            blk.append(x)
            every.append(x)
        merges.append(f"A{r0}:A{r0 + 2}")
        rows.append([word, None, _avg(blk, 2), _avg(blk, 3), _avg(blk, 4), _sum(blk, 5)])
    rows.append([last, None, _avg(every, 2), _avg(every, 3), _avg(every, 4), _sum(every, 5)])
    got = _summary(tmp_path, "d.xlsx", rows, merges)
    assert len(got) == 11 and got[-1][0] == 42 and got[-1][2] == last, got


def test_category_rows_still_quiet(tmp_path):
    # categories hold numbers like the data rows (2.9.27 c2 probes): a value near a column mean alone confirms nothing
    r = random.Random(2)
    labels = ["A계"] * 3 + ["B조"] * 17
    r.shuffle(labels)
    rows = [["구분", "온도편차", "생산량"]] + [[lab, round(r.uniform(-2, 2), 2), r.randint(280, 320)] for lab in labels]
    assert _summary(tmp_path, "c.xlsx", rows) == []
    r = random.Random(6)
    labels = ["1계"] * 14 + ["2계"] * 6
    r.shuffle(labels)
    rows = [["계통", "생산량", "원료투입", "전력"]]
    for lab in labels:
        s = 3 if lab == "2계" else 1
        rows.append([lab, r.randint(280, 320) * s, r.randint(300, 340) * s, r.randint(50, 60) * s])
    assert _summary(tmp_path, "g.xlsx", rows) == []
    # a category row whose temperature equals the mean and quantity is no sum: quiet
    d = _process_days(seed=9, n=20)
    rows = [H_DAYS] + d[:10] + [["A계", _avg(d[:10], 1), _avg(d[:10], 2), _avg(d[:10], 3), 301]] + d[10:]
    assert _summary(tmp_path, "m.xlsx", rows) == []


def test_mean_neutral_needs_a_sum(tmp_path):
    a = mp._sums_match
    import numpy as np
    sums = np.array([[300.0, 30.0, 3000.0]])
    count = np.array([[10.0, 10.0, 10.0]])
    # 30.0 (mean of col 0), 3.0 (mean of col 1), 3000 (sum) -> confirmed; all means, no sum -> not
    assert a(np.array([[30.0, 3.0, 3000.0]]), sums, count, sums)[0]
    assert not a(np.array([[30.0, 3.0, 300.0]]), sums, count, sums)[0]
    # one sum among two unrelated cells -> not (1 * 2 < 3)
    assert not a(np.array([[17.0, 9.0, 3000.0]]), sums, count, sums)[0]


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
    f = write_rows(str(tmp_path / f"c.{ext}"), _const_rows(), [])
    assert mp.load_data_file(f).attrs["summary_rows"] == []


def test_totals_with_varying_means_still_warned(tmp_path):
    r = random.Random(7)
    d = [[f"10/{i + 1}", round(r.uniform(175, 185), 1), round(r.uniform(2, 3), 2), round(r.uniform(40, 50), 1),
          r.randint(280, 320)] for i in range(30)]
    avg = [round(sum(x[c] for x in d) / len(d), 2) for c in (1, 2, 3)]
    rows = [["일자", "온도", "압력", "속도", "생산량"]] + d + [["총 합계"] + avg + [sum(x[4] for x in d)]]
    got = mp.load_data_file(write_rows(str(tmp_path / "t.xlsx"), rows, [])).attrs["summary_rows"]
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


def _setpoint_shift_rows(days=10, seed=11):
    r = random.Random(seed)
    rows = [["일자", "구분", "설정온도", "설정압력", "생산량"]]
    for d in range(days):
        t, p = r.choice([175, 180, 185]), r.choice([3.0, 3.5])
        a, b = r.randint(400, 600), r.randint(400, 600)
        rows += [[f"11/{d + 1}", "주간", t, p, a], [f"11/{d + 1}", "야간", t, p, b], [f"11/{d + 1}", "일합계", t, p, a + b]]
    return rows


@pytest.mark.parametrize("ext", ["csv", "xlsx"])
def test_shift_report_with_setpoints_warned(tmp_path, ext):
    f = write_rows(str(tmp_path / f"s.{ext}"), _setpoint_shift_rows(), [])
    got = mp.load_data_file(f).attrs["summary_rows"]
    assert [g[2] for g in got] == ["일합계"] * 10, got


FP_CONST = [  # r18 fp_const.csv: category rows 'A계' among '공통', constant setpoints, a 0/1 flag
    ("공통", 0, 107), ("공통", 1, 118), ("A계", 0, 111), ("공통", 1, 129), ("공통", 0, 119), ("A계", 1, 120),
    ("공통", 0, 118), ("공통", 0, 102), ("A계", 0, 129), ("공통", 1, 126), ("공통", 1, 115), ("A계", 1, 117),
    ("공통", 0, 107), ("공통", 1, 106), ("A계", 1, 117), ("공통", 0, 126), ("공통", 0, 117), ("A계", 1, 112),
    ("공통", 1, 120), ("공통", 0, 127), ("A계", 0, 107), ("공통", 0, 120), ("공통", 1, 104), ("A계", 1, 123)]


def test_constant_setpoints_category_still_quiet(tmp_path):
    rows = [["구분", "설정온도", "설정압력", "불량", "생산량"]] + [[g, 180, 3.5, b, q] for g, b, q in FP_CONST]
    assert mp.load_data_file(write_rows(str(tmp_path / "fp.csv"), rows, [])).attrs["summary_rows"] == []


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


# ---- 알려진 거짓 경고 (그대로 둠) ------------------------------------------------------------------------------


CAT = [("일반", 1, 100), ("일반", 1, 104), ("A계", 2, 101), ("일반", 0, 99), ("일반", 2, 103), ("일반", 1, 98),
       ("A계", 1, 97), ("일반", 1, 102), ("일반", 0, 96), ("A계", 0, 105), ("일반", 2, 100), ("일반", 1, 101),
       ("일반", 0, 99), ("일반", 1, 103)]


def test_known_false_warning_small_integer_sum(tmp_path):
    rows = [["라인", "설정온도", "불량수", "생산량"]] + [[g, 180, b, q] for g, b, q in CAT]
    got = mp.load_data_file(write_rows(str(tmp_path / "cat.csv"), rows, [])).attrs["summary_rows"]
    assert got == [[4, "라인", "A계"]], got        # 불량수 2 = 1 + 1 and a constant setpoint: warned (documented)
