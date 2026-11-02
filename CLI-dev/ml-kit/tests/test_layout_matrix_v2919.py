# -*- coding: utf-8 -*-
"""
2.9.19 (8차 점검): more layouts for the header / unit row matrix (test_layout_matrix_v2917.py, v2918), totals rows and
decimal-comma CSV files.

Layouts run end to end like 2.9.17/2.9.18 ((a) training with the correct options, (b) training without options,
(c) prediction of a new file with models trained with and without units, template and next_commands):
  N11  a unit row with a note '비고' under a blank header cell (2.9.18 predicted the unit row as data)
  L3s  a unit row with a scale number ('°C', 1000, '', '%'), refused by 2.9.18 as a data row
Prediction files without a target column whose first row holds missing marks (B1 '측정불가, 측정불가, A', B1b without the
category, B1c 'NG, NG, A'): every row predicted, no [추천] that drops a row (2.9.18: [추천] --data-row 3).
Training file whose first data row is '측정불가, 측정불가, A, <blank target>' with --unit-row 2 (B2): STOP_RETRY whose
[추천] has no --unit-row (2.9.18: names '반응온도(측정불가)', VERIFIED).
Totals rows: a row marked in the '제외' column (A1, A1b) and a category column '집계' = 평균/최대 (A2) give no totals
warning; an unmarked totals row and repeated subtotal rows still do (training and ml_map).
Decimal-comma CSV (';', '175,6'): read as decimals with a warning (2.9.18: 1756, VERIFIED); mixed with '1,500' ->
STOP_RETRY; a comma CSV with thousands separators is unchanged.
"""
import csv
import os
import shlex
import sys

import pandas as pd
import pytest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import test_layout_matrix_v2917 as base  # noqa: E402
from test_layout_matrix_v2917 import H, U, mp, _write, _run_with_one_retry, _job, _predict_command  # noqa: E402
from test_cli_smoke import run_raw  # noqa: E402

N_TRAIN, N_PRED = 20, 5
X, Y = "반응온도,압력,라인", "수율"
SCALE_U = ["°C", 1000, None, "%"]

# id: (top rows, correct options, units of the names)
LAYOUTS = {
    "N11_blank_hdr_note": ([H + [None], U + ["비고"]], ["--unit-row", "2"], U),
    "L3s_scale_unit": ([H, SCALE_U], ["--unit-row", "2"], SCALE_U),
}
# models for (c): trained with units (header + unit row), without units (header only)
MODEL_LAYOUTS = {"units": ([H, U], ["--unit-row", "2"]), "plain": ([H], [])}


def _make(folder, lid, n, seed, top=None):
    """(file, 1-based source rows of the data rows)"""
    top = LAYOUTS[lid][0] if top is None else top
    data = base._data("A", n, seed)
    path = _write(os.path.join(str(folder), lid + ".xlsx"), [list(r) for r in top] + data, [])
    return path, list(range(len(top) + 1, len(top) + n + 1))


def _names(lid, with_units):
    units = LAYOUTS[lid][2]
    cols = base.FAMILIES["A"]["cols"]
    return [mp._with_unit(c, u) for c, u in zip(cols, units)] if with_units else list(cols)


def _train(path, opts):
    return _run_with_one_retry("ml_train.py", ["-f", path, "-x", X, "-y", Y, *opts])


PRED_CASES = [(lid, kind, how) for lid in LAYOUTS for kind in ("units", "plain") for how in ("template", "next")
              if not (kind == "plain" and how == "next")]


def _predict_args(trained, path, how):
    if how == "template":
        return ["-m", trained["version_dir"], "-f", path]
    return shlex.split(_predict_command(trained).replace("<새데이터.xlsx>", path))[2:]


def _predict(folder, lid, how, model_res):
    path = _make(folder, lid, N_PRED, 2)[0]
    if isinstance(model_res, BaseException):
        raise model_res
    return _run_with_one_retry("ml_predict.py", _predict_args(model_res[1], path, how))


# prediction files without a target column, first row of missing marks (B1, B1b, B1c)
NOTARGET = {
    "B1_meas_cat": ["측정불가", "측정불가", "A"],
    "B1b_meas_nocat": ["측정불가", "측정불가", None],
    "B1c_ng_cat": ["NG", "NG", "A"],
}


def _make_notarget(folder, bid):
    data = [r[:3] for r in base._data("A", 6, 2)]
    data[0] = list(NOTARGET[bid])
    return _write(os.path.join(str(folder), bid + ".xlsx"), [H[:3]] + data, [])


def _predict_notarget(folder, bid, model_res):
    if isinstance(model_res, BaseException):
        raise model_res
    path = _make_notarget(folder, bid)
    return run_raw("ml_predict.py", "-m", model_res[1]["version_dir"], "-f", path)[1]


@pytest.fixture(scope="module")
def matrix(tmp_path_factory):
    from concurrent.futures import ThreadPoolExecutor
    mktemp = tmp_path_factory.mktemp
    files, out = {}, {}
    for lid in LAYOUTS:
        files[lid] = _make(mktemp(lid.split("_")[0] + "t"), lid, N_TRAIN, 1)
    for kind, (top, _) in MODEL_LAYOUTS.items():
        files[kind] = _make(mktemp("m" + kind), "M_" + kind, N_TRAIN, 1, top)
    with ThreadPoolExecutor(max_workers=4) as pool:
        futs = {("a", lid): pool.submit(_job, _train, files[lid][0], LAYOUTS[lid][1]) for lid in LAYOUTS}
        futs.update({("m", kind): pool.submit(_job, _train, files[kind][0], opts)
                     for kind, (_, opts) in MODEL_LAYOUTS.items()})
        out.update({k: f.result() for k, f in futs.items()})
        futs = {("b", lid): pool.submit(_job, _train, files[lid][0], []) for lid in LAYOUTS}
        for lid, kind, how in PRED_CASES:
            futs[("c", lid, kind, how)] = pool.submit(_job, _predict, mktemp(lid.split("_")[0] + "p"), lid, how,
                                                      out[("m", kind)])
        for bid in NOTARGET:
            for kind in MODEL_LAYOUTS:
                futs[("n", bid, kind)] = pool.submit(_job, _predict_notarget, mktemp(bid.split("_")[0]), bid,
                                                     out[("m", kind)])
        out.update({k: f.result() for k, f in futs.items()})
    out["files"] = files
    return out


def _expect(lid, with_units):
    names = _names(lid, with_units)
    return sorted(names), sorted(n for n, c in zip(names, base.FAMILIES["A"]["cols"]) if c in ("반응온도", "압력"))


@pytest.mark.parametrize("lid", list(LAYOUTS))
def test_train_with_correct_options(matrix, lid):
    first, final = base._result(matrix, ("a", lid))
    assert first is final, first.get("message")     # 2.9.18: both refused as "not a unit row"
    names, numeric, used = base._trained_layout(final)
    assert (names, numeric) == _expect(lid, True)
    assert used == matrix["files"][lid][1]


@pytest.mark.parametrize("lid", list(LAYOUTS))
def test_train_without_options(matrix, lid):
    first, final = base._result(matrix, ("b", lid))
    names, numeric, used = base._trained_layout(final)
    assert (names, numeric) in (_expect(lid, True), _expect(lid, False)), first.get("message")
    assert used == matrix["files"][lid][1]


@pytest.mark.parametrize("lid,kind,how", PRED_CASES)
def test_predict_same_layout(matrix, lid, kind, how):
    first, final = base._result(matrix, ("c", lid, kind, how))
    assert final["status"] == "ok" and final["agent_directive"].startswith("VERIFIED"), (first, final)
    assert final["sample_count"] == N_PRED, (first.get("message"), final.get("preview"))   # 2.9.18: 6 (unit row)
    assert len(mp.load_data_file(final["output_file"])) == N_PRED


@pytest.mark.parametrize("bid", list(NOTARGET))
@pytest.mark.parametrize("kind", list(MODEL_LAYOUTS))
def test_predict_first_row_missing_marks_without_target(matrix, bid, kind):
    res = base._result(matrix, ("n", bid, kind))
    # 2.9.18: '[추천] --data-row 3' (a real data row dropped) for B1 / B1c, B1b
    assert res["status"] == "ok" and res["agent_directive"].startswith("VERIFIED"), res
    assert res["sample_count"] == 6


def test_train_unit_row_option_on_missing_marks_row(tmp_path):
    data = base._data("A", 20, 1)
    data[0] = ["측정불가", "측정불가", "A", None]
    path = _write(str(tmp_path / "b2.xlsx"), [H] + data, [])
    first, final = _train(path, ["--unit-row", "2"])
    assert first["status"] == "error" and first["agent_directive"].startswith("STOP_RETRY"), first
    assert "--unit-row" not in first["message"].split("[추천]")[1], first["message"]    # 2.9.18: VERIFIED
    names, numeric, used = base._trained_layout(final)
    assert names == sorted(H) and used == list(range(3, 22))


# unit-row rules (the matrix above runs them end to end) --------------------------------------------------------

def _raw(rows):
    return pd.DataFrame(rows, dtype=object)


def test_unit_row_rules_per_cell():
    body = [[180.0, 3.91, "B", 73.39], [152.8, 2.1, "A", 59.83], [157.5, 2.73, "B", 60.5], [171.1, 3.53, "B", 68.68]]
    assert not mp._looks_like_unit_row(_raw([H, ["측정불가", "측정불가", "A", None]] + body), 1)   # 'A' is a value below
    assert not mp._looks_like_unit_row(_raw([H, ["NG", "NG", "A", None]] + body), 1)
    assert not mp._looks_like_unit_row(_raw([H, ["측정불가", "측정불가", None, None]] + body), 1)  # missing values only
    assert not mp._looks_like_unit_row(_raw([H[:2], ["측정불가", "측정불가"]] + [r[:2] for r in body]), 1)
    assert mp._looks_like_unit_row(_raw([H, ["°C", 1000, None, "%"]] + body), 1)                  # a scale unit
    assert not mp._looks_like_unit_row(_raw([H, [175.0, "bar", None, "%"]] + body), 1)            # a number in range
    assert mp._looks_like_unit_row(_raw([H + [None], U + ["비고"]] + [r + [None] for r in body]), 1)
    # a header under a title row of fewer cells is still not a unit row of the title
    assert not mp._looks_like_unit_row(_raw([["시험 기록", "작성자", None, None], H] + body), 1)
    assert not mp._looks_like_unit_row(_raw([["반응 공정 데이터", None, None, None], H] + body), 1)
    # the unit row repeated on page breaks is not "a value below"
    assert mp._looks_like_unit_row(_raw([H, U] + body[:2] + [H, U] + body[2:]), 1)


def test_unit_row_note_column_with_data_is_still_refused():
    body = [[180.0, 3.91, "B", 73.39, 1.0], [152.8, 2.1, "A", 59.83, 2.0], [157.5, 2.73, "B", 60.5, 3.0]]
    assert not mp._looks_like_unit_row(_raw([H + [None], U + ["비고"]] + body), 1)


# totals rows ------------------------------------------------------------------------------------------------------

def _totals_book(path, mark):
    data = base._data("A", 20, 1)
    tot = ["합계", round(sum(r[0] for r in data), 1), round(sum(r[1] for r in data), 2), None,
           round(sum(r[3] for r in data), 2)]
    rows = [["구분"] + H + ["제외"], [None] + U + [None]] + [[i + 1] + r + [None] for i, r in enumerate(data)]
    rows.append(tot + [mark])
    return _write(path, rows, [])


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


def test_aggregation_category_is_not_a_totals_row():
    data = base._data("A", 20, 3)
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


# decimal-comma CSV -----------------------------------------------------------------------------------------------

def _csv(path, rows, delim):
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        csv.writer(f, delimiter=delim).writerows(rows)
    return path


def _comma(v):
    return str(v).replace(".", ",") if isinstance(v, float) else v


def test_decimal_comma_csv(tmp_path):
    data = base._data("A", 20, 1)
    path = _csv(str(tmp_path / "c1.csv"), [["시험 기록"], H, U] + [[_comma(v) for v in r] for r in data], ";")
    df = mp.load_data_file(path, header_row=2, unit_row=3)
    t = mp.coerce_numeric(df["반응온도(°C)"])
    assert t.tolist() == [r[0] for r in data]                                     # 2.9.18: 1756.0 for '175,6'
    assert any("소수점" in w for w in mp.layout_warnings(df))
    _, res = run_raw("ml_train.py", "-f", path, "-x", X, "-y", Y, "--header-row", "2", "--unit-row", "3")
    assert res["status"] == "ok" and any("소수점" in w for w in res.get("warnings", [])), res
    # a short prediction file: csv.Sniffer alone picks ',' ('163,1;2,19;A;61,66'), the table width picks ';'
    short = _csv(str(tmp_path / "c1p.csv"), [["시험 기록"], H, U] + [[_comma(v) for v in r] for r in data[:6]], ";")
    df = mp.load_data_file(short, header_row=2, unit_row=3)
    assert list(df.columns) == ["반응온도(°C)", "압력(bar)", "라인", "수율(%)"]
    assert mp.coerce_numeric(df["압력(bar)"]).tolist() == [r[1] for r in data[:6]]


def test_decimal_comma_mixed_with_thousands_is_refused(tmp_path):
    rows = [H, ["175,6", "1,500", "A", "61,66"], ["180,1", "2,19", "B", "70,2"]]
    path = _csv(str(tmp_path / "c2.csv"), rows, ";")
    with pytest.raises(mp.InputError) as e:
        mp.load_data_file(path)
    assert "1,500" in str(e.value)


def test_thousands_comma_csv_unchanged(tmp_path):
    path = _csv(str(tmp_path / "c3.csv"), [["유량", "수율"], ["1,500", "61.5"], ["2,250", "70.25"]], ",")
    df = mp.load_data_file(path)
    assert mp.coerce_numeric(df["유량"]).tolist() == [1500.0, 2250.0] and not mp.layout_warnings(df)
    path = _csv(str(tmp_path / "c4.csv"), [["유량", "수율"], ["1,500", "61.5"], ["2,250", "70.25"]], ";")
    assert mp.coerce_numeric(mp.load_data_file(path)["유량"]).tolist() == [1500.0, 2250.0]
