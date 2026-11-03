# -*- coding: utf-8 -*-
"""
2.9.20 (9차 점검) regressions.

1. ml_predict.py assigned _THOUSANDS_RE twice (anchored for decimal-comma CSV, unanchored for 'k=v' input): the CSV
   check used the second one with a prefix match, so '27,8448' counted as thousands (D7: 278448, VERIFIED; D6: false
   STOP_RETRY). The names are distinct now, and no module-level name may be assigned twice.
2. Totals rows: 2.9.19 skipped a column whose totals words were a majority or 20%+ of its cells, which silenced real
   subtotal rows (S1 shift report + 소계 25%, S3 two lines + 합계 33%) and a lone '합계' in a blank column (S2).
   Only average words (평균/mean) can be a category now; '집계' = 평균/최대 still gives no warning.
3. Decimal-comma CSV: '0,149' cannot be thousands (D2: 149 VERIFIED; D1: false STOP_RETRY); European '1.346,8' read
   as 1346.8 with a warning (2.9.19: 1.3468), refused when mixed with other forms.
4. 'ng' (nanogram) was a missing mark: --unit-row over 'ng, ng' refused, 'ng' unit dropped from the name.
5. Loading 100k rows: pandas deep-copied df.attrs['source_rows'] on every Series operation.
"""
import ast
import copy
import csv
import os
import sys

import numpy as np
import pandas as pd
import pytest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import test_layout_matrix_v2917 as base  # noqa: E402
from test_layout_matrix_v2917 import H, mp, _write  # noqa: E402
from test_cli_smoke import KIT, run_raw  # noqa: E402

N = 40


def _csv(path, header, rows, delim=";"):
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=delim)
        w.writerow(header)
        w.writerows(rows)
    return str(path)


def _dc(x, d):
    return f"{x:.{d}f}".replace(".", ",")


def _series(seed=1):
    r = np.random.default_rng(seed)
    t, p, th = r.normal(175, 3, N), r.normal(2.2, 0.1, N), r.uniform(0.1, 0.3, N)
    return t, p, th, 0.5 * t + 20 * p + 50 * th + r.normal(0, 0.5, N)


def _num(df, col):
    return mp.coerce_numeric(df[col]).tolist()


# 1. one assignment per module-level name -------------------------------------------------------------------------

def test_no_module_level_name_assigned_twice():
    tree = ast.parse(open(os.path.join(KIT, "ml_predict.py"), encoding="utf-8").read())
    seen, dup = set(), []
    for node in tree.body:
        targets = node.targets if isinstance(node, ast.Assign) else \
            [node.target] if isinstance(node, (ast.AnnAssign, ast.AugAssign)) else []
        for t in targets:
            for n in ast.walk(t):
                if isinstance(n, ast.Name):
                    (dup.append(n.id) if n.id in seen else seen.add(n.id))
    assert not dup, dup


def test_four_decimal_comma_numbers(tmp_path):
    t, _, th, y = _series()
    th = th * 100                                           # '27,8448'
    d7 = _csv(tmp_path / "d7.csv", ["온도", "두께", "수율"], [[int(t[i]), _dc(th[i], 4), int(y[i])] for i in range(N)])
    df = mp.load_data_file(d7)
    assert _num(df, "두께") == [round(v, 4) for v in th]    # 2.9.19: 278448.0, VERIFIED
    assert any("소수점" in w for w in mp.layout_warnings(df))
    d6 = _csv(tmp_path / "d6.csv", ["온도", "두께", "수율"],
              [[_dc(t[i], 1), _dc(th[i], 4), _dc(y[i], 1)] for i in range(N)])
    df = mp.load_data_file(d6)                              # 2.9.19: STOP_RETRY (ambiguous)
    assert _num(df, "온도") == [round(v, 1) for v in t] and _num(df, "두께") == [round(v, 4) for v in th]
    _, res = run_raw("ml_train.py", "-f", d7, "-x", "온도,두께", "-y", "수율")
    assert res["status"] == "ok" and any("소수점" in w for w in res.get("warnings", [])), res


def test_kv_input_thousands_still_recommended():
    with pytest.raises(mp.InputError) as e:
        mp.check_comma_numbers({"온도": "1,500"}, lambda k: True, "-i", lambda p: "python x -i " + mp.kv_text(p))
    assert "[추천] python x -i 온도=1500" in str(e.value)
    with pytest.raises(mp.InputError) as e:
        mp.check_comma_numbers({"온도": "1,5"}, lambda k: True, "-i", lambda p: "x")
    assert "[추천]" not in str(e.value)


# 2. totals rows ---------------------------------------------------------------------------------------------------

def _rows(df):
    return [(r[0], r[2]) for r in df.attrs["summary_rows"]]


def test_shift_subtotal_rows_warned(tmp_path):
    r = np.random.default_rng(5)
    rows = [["일자", "구분", "반응온도", "압력", "생산량"]]
    for d in range(12):
        day = [[f"2026-10-{d + 1:02d}", s, round(float(r.uniform(150, 200)), 1), round(float(r.uniform(1, 5)), 2),
                round(float(r.uniform(300, 450)), 1)] for s in ("주간", "야간", "심야")]
        rows += day + [[f"2026-10-{d + 1:02d}", "소계"] + [round(sum(x[c] for x in day), 2) for c in (2, 3, 4)]]
    path = _write(str(tmp_path / "s1.xlsx"), rows, [])
    df = mp.load_data_file(path)
    assert _rows(df) == [(5 + 4 * d, "소계") for d in range(12)]           # 2.9.19: [] (25% of the column)
    _, res = run_raw("ml_train.py", "-f", path, "-x", "반응온도,압력", "-y", "생산량")
    assert res["status"] == "ok" and any("합계·평균" in w for w in res.get("warnings", [])), res


def test_two_lines_and_total_rows_warned(tmp_path):
    r = np.random.default_rng(5)
    rows = [["구분", "반응온도", "압력", "수율"]]
    for _ in range(15):
        a = [["1라인", round(float(r.uniform(150, 200)), 1), round(float(r.uniform(1, 5)), 2),
              round(float(r.uniform(80, 95)), 1)] for _ in range(2)]
        rows += a + [["합계"] + [a[0][c] + a[1][c] for c in (1, 2, 3)]]
    df = mp.load_data_file(_write(str(tmp_path / "s3.xlsx"), rows, []))
    assert _rows(df) == [(4 + 3 * k, "합계") for k in range(15)]           # 2.9.19: [] (33%)


def test_lone_total_in_blank_column_warned(tmp_path):
    d = base._data("A", N, 1)
    tot = ["합계", round(sum(x[0] for x in d), 1), round(sum(x[1] for x in d), 2), None, round(sum(x[3] for x in d), 2)]
    df = mp.load_data_file(_write(str(tmp_path / "s2.xlsx"), [["비고"] + H] + [[None] + x for x in d] + [tot], []))
    assert _rows(df) == [(N + 2, "합계")]                                  # 2.9.19: [] (1 of 1 filled cells)
    avg = ["평균"] + tot[1:]
    df = mp.load_data_file(_write(str(tmp_path / "s2b.xlsx"), [["비고"] + H] + [[None] + x for x in d] + [avg], []))
    assert _rows(df) == [(N + 2, "평균")]


def test_average_category_still_not_warned():
    data = base._data("A", 20, 3)
    agg = pd.DataFrame([r + ["평균" if i % 2 else "최대"] for i, r in enumerate(data)], columns=H + ["집계"])
    assert mp._summary_rows(agg, list(range(3, 23))) == []
    agg["집계"] = ["평균" if i % 4 == 0 else "측정" for i in range(20)]        # 25%: a category of the column
    assert mp._summary_rows(agg, list(range(3, 23))) == []
    agg["집계"] = ["평균" if i < 3 else "최대" for i in range(20)]
    assert mp._summary_rows(agg, list(range(3, 23))) == []


def test_summary_rows_keep_exclude_marks():
    data = base._data("A", 6, 3)
    df = pd.DataFrame([["소계"] + r + ["합계"] for r in data[:2]] + [["A"] + r + [None] for r in data[2:]],
                      columns=["구분"] + H + ["제외"])
    assert mp._summary_rows(df, list(range(2, 8))) == []
    df["제외"] = [None] * 6
    assert [r[0] for r in mp._summary_rows(df, list(range(2, 8)))] == [2, 3]


# 3. decimal comma ------------------------------------------------------------------------------------------------

def test_leading_zero_comma_is_decimal(tmp_path):
    t, p, th, y = _series()
    d2 = _csv(tmp_path / "d2.csv", ["온도", "두께", "수율"], [[int(t[i]), _dc(th[i], 3), int(y[i])] for i in range(N)])
    df = mp.load_data_file(d2)
    assert _num(df, "두께") == [round(v, 3) for v in th]                    # 2.9.19: 149.0, VERIFIED
    assert any("소수점" in w for w in mp.layout_warnings(df))
    d1 = _csv(tmp_path / "d1.csv", ["온도", "압력", "두께", "수율"],
              [[_dc(t[i], 1), _dc(p[i], 2), _dc(th[i], 3), _dc(y[i], 1)] for i in range(N)])
    df = mp.load_data_file(d1)                                              # 2.9.19: STOP_RETRY (ambiguous)
    assert _num(df, "두께") == [round(v, 3) for v in th] and _num(df, "온도") == [round(v, 1) for v in t]


def _euro(x):
    return f"{x:,.1f}".replace(",", "X").replace(".", ",").replace("X", ".")


def test_european_thousands_dot(tmp_path):
    t, _, _, y = _series()
    flow = np.random.default_rng(2).uniform(1000, 1500, N)
    rows = [[_dc(t[i], 1), _euro(flow[i]), _dc(y[i], 1)] for i in range(N)]
    df = mp.load_data_file(_csv(tmp_path / "d3.csv", ["온도", "유량", "수율"], rows))
    assert _num(df, "유량") == [round(v, 1) for v in flow]                 # 2.9.19: 1.3468 for '1.346,8'
    warns = mp.layout_warnings(df)
    ex = _euro(flow[0])
    assert any("천 단위" in w and f"{ex} -> {ex.replace('.', '').replace(',', '.')}" in w for w in warns), warns
    # only European numbers, no plain decimal comma: still read as 1346.8
    rows2 = [[int(t[i]), _euro(flow[i]), int(y[i])] for i in range(N)]
    df = mp.load_data_file(_csv(tmp_path / "d3b.csv", ["온도", "유량", "수율"], rows2))
    assert _num(df, "유량") == [round(v, 1) for v in flow]
    # mixed with a US number: refused, never 1.3468
    rows3 = [list(r) for r in rows]
    rows3[5][1] = "1,346.8"
    with pytest.raises(mp.InputError):
        mp.load_data_file(_csv(tmp_path / "d3x0.csv", ["온도", "유량", "수율"], rows3))
    # a dot-grouped integer without a comma: 1346 or 1.346 -> refused (2.9.21 read 1346; 2.9.22 test_review_v2922)
    rows3[5][1] = "1.346"
    with pytest.raises(mp.InputError):
        mp.load_data_file(_csv(tmp_path / "d3x1.csv", ["온도", "유량", "수율"], rows3))
    _, res = run_raw("ml_train.py", "-f", _csv(tmp_path / "d3t.csv", ["온도", "유량", "수율"], rows),
                     "-x", "온도,유량", "-y", "수율")
    assert res["status"] == "ok" and any("천 단위" in w for w in res.get("warnings", [])), res


def test_thousands_and_decimal_still_refused(tmp_path):
    rows = [["175,6", "1,500", "61,66"], ["180,1", "2,19", "70,2"]]
    with pytest.raises(mp.InputError):
        mp.load_data_file(_csv(tmp_path / "c.csv", ["온도", "유량", "수율"], rows))


# 4. 'ng' is a unit -----------------------------------------------------------------------------------------------

def _ng_book(path, units):
    r = np.random.default_rng(3)
    rows = [["농도A", "농도B", "수율"], units]
    for _ in range(20):
        a, b = round(float(r.uniform(150, 160)), 1), round(float(r.uniform(1, 3)), 2)
        rows.append([a, b, round(0.3 * a + 5 * b + float(r.normal(0, 0.5)), 2)])
    return _write(str(path), rows, [])


def test_nanogram_unit_row(tmp_path):
    only = _ng_book(tmp_path / "p2b.xlsx", ["ng", "ng", None])
    df = mp.load_data_file(only, unit_row=2)                               # 2.9.19: refused (missing marks only)
    assert list(df.columns) == ["농도A(ng)", "농도B(ng)", "수율"] and len(df) == 20
    assert mp.load_data_file(only).attrs["unit_row_below"] == 2
    mixed = _ng_book(tmp_path / "p2.xlsx", ["ng", "°C", "%"])
    assert list(mp.load_data_file(mixed, unit_row=2).columns) == ["농도A(ng)", "농도B(°C)", "수율(%)"]
    _, res = run_raw("ml_train.py", "-f", only, "-x", "농도A,농도B", "-y", "수율", "--unit-row", "2")
    assert res["status"] == "ok", res
    # 'NG' (upper case) is still a missing mark
    raw = pd.DataFrame([["농도A", "농도B", "수율"], ["NG", "NG", None], [150.1, 2.0, 60.0], [151.0, 2.5, 61.0]],
                       dtype=object)
    assert not mp._looks_like_unit_row(raw, 1)


# 5. attrs are not deep-copied per operation ----------------------------------------------------------------------

def test_source_rows_not_copied_per_operation(tmp_path):
    path = _csv(tmp_path / "big.csv", ["a", "b"], [[str(i), str(i % 7)] for i in range(2000)], ",")
    df = mp.load_data_file(path)
    rows = df.attrs["source_rows"]
    assert rows == list(range(2, 2002))
    assert df["a"].attrs["source_rows"] is rows                           # 2.9.19: a deep copy per Series
    assert copy.deepcopy(df.attrs)["source_rows"] is rows
    assert mp.coerce_numeric(df["a"]).iloc[-1] == 1999.0
