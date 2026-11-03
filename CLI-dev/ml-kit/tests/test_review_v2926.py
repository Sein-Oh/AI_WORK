# -*- coding: utf-8 -*-
"""
2.9.26 (15차 점검) regressions.

1. 2.9.25 read every short word ending with '계' as a totals word, so category columns ('A계'/'B계', '1계'/'2계',
   '생산계'/'품질계'/'설비계', '레벨계'/'pH계'/'유량계') gave a totals warning on every row (train, predict, ml_map
   all asked to mark every row in '제외'). Loose words ('...계', '...합계', 'Total'/'Sum' as a separate word,
   '평균값') count only as summary rows within the data of a column: at most half of its filled cells and at most
   2 distinct texts. Exact words (합계/소계/총계/계, 일계/월계/누계 ..., Total/Sum) keep the 2.9.25 behaviour.
2. 'Total'/'Sum' count only as a separate word ('SUM-01', 'Lot-Sum', 'sum_a', 'Total_Loss' are names), and
   '평균' followed by a name ('평균온도') is not an average word (only '평균값' / '평균치').
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


def _category_rows(name, values, n=40, seed=1):
    r = random.Random(seed)
    rows = [[name, "온도", "압력", "수율"]]
    for i in range(n):
        t, p = round(r.uniform(175, 185), 1), round(r.uniform(2.5, 3.5), 2)
        rows.append([r.choice(values), t, p, round(0.3 * t + 2 * p + r.uniform(-1, 1), 2)])
    return rows


def _summary(tmp_path, name, rows):
    return mp.load_data_file(_write(str(tmp_path / name), rows, [])).attrs["summary_rows"]


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
    f = _write(str(tmp_path / "l1.xlsx"), _category_rows("라인", ["A계", "B계"]), [])
    _, res = run_raw("ml_train.py", "-f", f, "-x", "라인,온도,압력", "-y", "수율")
    assert res["status"] == "ok", res
    assert not any("합계·평균 행" in w for w in res.get("warnings", [])), res["warnings"]


def _daily(word, days=20):
    rows = [["일자", "교대", "온도", "생산량"]]
    for d in range(1, days + 1):
        a, b = [150 + d, 300 + d], [160 + d, 310 + d]
        rows += [[f"10/{d}", "주간", *a], [None, "야간", *b], [word, None, a[0] + b[0], a[1] + b[1]]]
    return rows


@pytest.mark.parametrize("word", ["일계", "총 합계", "소계(A)", "주간계", "Grand Total", "합계", "소계"])
def test_daily_report_totals_still_warned(tmp_path, word):
    got = _summary(tmp_path, "w.xlsx", _daily(word))
    assert len(got) == 20 and {r[2] for r in got} == {word}, got


def test_daily_report_totals_train_warned(tmp_path):
    f = _write(str(tmp_path / "w.xlsx"), _daily("총 합계"), [])
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


def test_word_kind_split():
    k = mp._word_kind
    assert [k(w) for w in ("합계", "소계(A)", "Total:", "Sub Total", "일계", "누계", "평균", "AVG.")] == \
        ["total", "total", "total", "total", "total", "total", "avg", "avg"]
    assert [k(w) for w in ("총 합계", "A계", "주간계", "Grand Line Total", "Sum A")] == ["total_loose"] * 5
    assert [k(w) for w in ("평균값", "평균치")] == ["avg_loose"] * 2
    assert [k(w) for w in ("SUM-01", "Lot-Sum", "sum_a", "Total_Loss", "평균온도", "평균압력값", "설계")] == [""] * 7
