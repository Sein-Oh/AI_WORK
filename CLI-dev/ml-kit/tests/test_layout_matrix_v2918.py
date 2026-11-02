# -*- coding: utf-8 -*-
"""
2.9.18 (7차 점검): more layouts for the header / unit row matrix of test_layout_matrix_v2917.py.

Layouts: a first data row whose numeric inputs are all '측정불가' (header only, header + unit row), a first data row
blank except the target under a unit row, a CSV whose title line is narrower than the table (ragged, with and
without a unit row), a title row of 2 cells over the header (prediction with a model trained with units), the header
(and unit) row repeated mid-sheet (page breaks of an exported report), a totals row at the end.
For every layout, as in 2.9.17:
  (a) training with the correct options is VERIFIED with the right names and training rows;
  (b) training with no options is VERIFIED with the right names and rows, or one [추천] gives that;
  (c) prediction of a new file of the same layout (models trained with and without units, template and the model's
      next_commands command) is VERIFIED with every data row predicted, or one [추천] gives that.
Repeated header rows are dropped with a warning; a totals row is kept (a data value decision) with a warning that
names it ('제외' 열 in training).
"""
import os
import shlex
import sys

import pandas as pd
import pytest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import test_layout_matrix_v2917 as base  # noqa: E402
from test_layout_matrix_v2917 import H, U, mp, _write, _run_with_one_retry, _job, _predict_command  # noqa: E402

N_TRAIN, N_PRED = 20, 5
X, Y = "반응온도,압력,라인", "수율"


def _meas_first(d):
    d[0][0], d[0][1] = "측정불가", "측정불가"
    return d


def _blank_first(d):
    d[0][0], d[0][1], d[0][2] = None, None, None
    return d


def _total(d):
    nums = [r for r in d if all(isinstance(r[k], float) for k in (0, 1, 3))]
    return d + [[round(sum(r[0] for r in nums), 1), round(sum(r[1] for r in nums), 2), "합계",
                 round(sum(r[3] for r in nums), 2)]]


# id: (top rows, data edit, correct options, units of the names or None, extension, repeat the top rows after data row k)
LAYOUTS = {
    "M1_meas_first": ([H], _meas_first, [], None, ".xlsx", None),
    "M1c_meas_first_unit": ([H, U], _meas_first, ["--unit-row", "2"], U, ".xlsx", None),
    "N9_blank_first_unit": ([H, U], _blank_first, ["--unit-row", "2"], U, ".xlsx", None),
    "T5_ragged_csv": ([["반응 공정 데이터"], H], None, ["--header-row", "2"], None, ".csv", None),
    "T5b_ragged_csv_unit": ([["시험 기록", "작성자"], H, U], None, ["--header-row", "2", "--unit-row", "3"], U,
                            ".csv", None),
    "T3_title2_hdr": ([["시험 기록", "작성 2026-11-02"], H], None, ["--header-row", "2"], None, ".xlsx", None),
    "R1_repeat_hdr_unit": ([H, U], None, ["--unit-row", "2"], U, ".xlsx", 2),
    "S1_total_row": ([H, U], _total, ["--unit-row", "2"], U, ".xlsx", None),
}
MODELS = ("M1c_meas_first_unit", "M1_meas_first")      # trained with units, without units (round a)


def _make(folder, lid, n, seed):
    """(file, 1-based source rows of the data rows, of them the rows with a numeric target)"""
    top, edit, _, _, ext, repeat = LAYOUTS[lid]
    data = base._data("A", n, seed)
    if edit:
        data = edit(data)
    rows, src = [list(r) for r in top], []
    for k, d in enumerate(data):
        if repeat is not None and k == repeat:
            rows += [list(r) for r in top]
        rows.append(d)
        src.append(len(rows))
    path = _write(os.path.join(str(folder), lid + ext), rows, [])
    valid = [rw for rw, d in zip(src, data) if mp.coerce_numeric(pd.Series([d[3]])).notna().all()]
    return path, src, valid


def _names(lid, with_units):
    units = LAYOUTS[lid][3]
    cols = base.FAMILIES["A"]["cols"]
    return [mp._with_unit(c, u) for c, u in zip(cols, units)] if with_units and units else list(cols)


def _train(path, opts):
    return _run_with_one_retry("ml_train.py", ["-f", path, "-x", X, "-y", Y, *opts])


PRED_CASES = [(lid, kind, how) for lid in LAYOUTS for kind in ("units", "plain") for how in ("template", "next")
              if not (kind == "plain" and how == "next")]


def _predict(folder, lid, how, model_res):
    path = _make(folder, lid, N_PRED, 2)[0]
    if isinstance(model_res, BaseException):
        raise model_res
    trained = model_res[1]
    if how == "template":
        args = ["-m", trained["version_dir"], "-f", path]
    else:
        args = shlex.split(_predict_command(trained).replace("<새데이터.xlsx>", path))[2:]
    return _run_with_one_retry("ml_predict.py", args)


@pytest.fixture(scope="module")
def matrix(tmp_path_factory):
    from concurrent.futures import ThreadPoolExecutor
    mktemp = tmp_path_factory.mktemp
    files, out = {}, {}
    for lid in LAYOUTS:
        files[lid] = _make(mktemp(lid.split("_")[0] + "t"), lid, N_TRAIN, 1)
    with ThreadPoolExecutor(max_workers=4) as pool:
        futs = {("a", lid): pool.submit(_job, _train, files[lid][0], LAYOUTS[lid][2]) for lid in LAYOUTS}
        out.update({k: f.result() for k, f in futs.items()})
        futs = {("b", lid): pool.submit(_job, _train, files[lid][0], []) for lid in LAYOUTS}
        for lid, kind, how in PRED_CASES:
            model = out[("a", MODELS[0 if kind == "units" else 1])]
            futs[("c", lid, kind, how)] = pool.submit(_job, _predict, mktemp(lid.split("_")[0] + "p"), lid, how, model)
        out.update({k: f.result() for k, f in futs.items()})
    out["files"] = files
    return out


def _expect(lid, with_units):
    names = _names(lid, with_units)
    return sorted(names), sorted(n for n, c in zip(names, base.FAMILIES["A"]["cols"]) if c in ("반응온도", "압력"))


@pytest.mark.parametrize("lid", list(LAYOUTS))
def test_train_with_correct_options(matrix, lid):
    first, final = base._result(matrix, ("a", lid))
    assert first is final, first.get("message")
    names, numeric, used = base._trained_layout(final)
    assert (names, numeric) == _expect(lid, True)
    data_rows = matrix["files"][lid][2]
    if lid == "S1_total_row":       # the totals row is kept (its values are the user's decision), with a warning
        assert used in (data_rows, data_rows + [data_rows[-1] + 1])     # it may be an automatic outlier
        assert any("합계" in w and "'제외' 열" in w for w in final.get("warnings", [])), final.get("warnings")
    else:
        assert used == data_rows
    if lid == "R1_repeat_hdr_unit":
        assert any("5, 6행" in w and "머리글" in w for w in final.get("warnings", [])), final.get("warnings")


@pytest.mark.parametrize("lid", [lid for lid in LAYOUTS if lid != "S1_total_row"])
def test_train_without_options(matrix, lid):
    first, final = base._result(matrix, ("b", lid))
    names, numeric, used = base._trained_layout(final)
    assert (names, numeric) in (_expect(lid, True), _expect(lid, False)), first.get("message")
    assert used == matrix["files"][lid][2]


@pytest.mark.parametrize("lid,kind,how", PRED_CASES)
def test_predict_same_layout(matrix, lid, kind, how):
    first, final = base._result(matrix, ("c", lid, kind, how))
    assert final["status"] == "ok" and final["agent_directive"].startswith("VERIFIED"), (first, final)
    rows = N_PRED + (1 if lid == "S1_total_row" else 0)
    assert final["sample_count"] == rows, (first.get("message"), final.get("preview"))
    assert len(mp.load_data_file(final["output_file"])) == rows
    if lid == "R1_repeat_hdr_unit":
        assert any("5, 6행" in w for w in final.get("warnings", [])), final.get("warnings")
    if lid == "S1_total_row":
        assert any("합계" in w for w in final.get("warnings", [])), final.get("warnings")


# rules (the matrix above runs them end to end) --------------------------------------------------------------

def test_unit_row_has_no_numbers():
    raw = pd.DataFrame([H, ["측정불가", "측정불가", "A", 61.66], [180.0, 3.91, "B", 73.39], [152.8, 2.1, "A", 59.83],
                        [157.5, 2.73, "B", 60.5]], dtype=object)
    assert not mp._looks_like_unit_row(raw, 1)     # 2.9.17: accepted, names '반응온도(측정불가)'
    raw = pd.DataFrame([H, U, [None, None, None, 61.66], [180.0, 3.91, "B", 73.39], [152.8, 2.1, "A", 59.83]],
                       dtype=object)
    assert mp._looks_like_unit_row(raw, 1)


def test_ragged_csv(tmp_path):
    p = tmp_path / "t.csv"
    p.write_text("반응 공정 데이터\n반응온도,압력,라인,수율\n180,3,A,70\n181,3.1,B,71\n", encoding="cp949")
    df = mp.load_data_file(str(p))                  # 2.9.17: "Expected 1 fields in line 2, saw 4"
    assert list(df.columns) == ["반응온도", "압력", "라인", "수율"] and len(df) == 2


def test_model_names_with_units_match_header_without_units():
    cells = ["반응온도", "압력", "라인", "수율"]
    assert mp.names_in_cells(["반응온도(°C)", "압력(bar)", "라인"], cells, model_names=True)
    assert not mp.names_in_cells(["반응온도(°C)"], cells)              # user names (-x): a unit is part of the name
    assert not mp.names_in_cells(["온도(C)"], ["온도(°C)"], model_names=True)


def test_repeated_header_and_totals_rows(tmp_path):
    rows = [H, U] + [[170.0 + i, 2.0 + i / 10, "A", 60.0 + i] for i in range(4)] + [H, U] + \
        [[175.0 + i, 2.5, "B", 65.0 + i] for i in range(3)] + [[1220.0, 17.1, "합계", 435.0]]
    path = _write(str(tmp_path / "r.xlsx"), rows, [])
    df = mp.load_data_file(path, unit_row=2)
    assert len(df) == 8 and df.attrs["repeated_header_rows"] == [7, 8]
    assert mp.coerce_numeric(df["반응온도(°C)"]).notna().all()
    assert df.attrs["summary_rows"] == [[12, "라인", "합계"]]
    w = mp.layout_warnings(df, ["수율(%)"])
    assert "7, 8행" in w[0] and "'제외' 열" in w[1]
    cls = pd.DataFrame({"x": [1.0, 2.0], "등급": ["계", "상"]})
    cls.attrs["summary_rows"] = [[2, "등급", "계"]]
    assert mp.layout_warnings(cls, ["등급"]) == []                     # a class label '계' of the target
    assert len(mp.layout_warnings(cls, ["등급"], training=False)) == 0 and len(mp.layout_warnings(cls, [], False)) == 1
