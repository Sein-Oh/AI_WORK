# -*- coding: utf-8 -*-
"""
2.9.17 (6차 점검): layout matrix for the header / unit row logic, which regressed in 2.9.14, 2.9.15 and 2.9.16.

Each layout is a small real .xlsx/.csv (openpyxl) of a shape seen in the field: header only, header + units, title rows,
a merged two-level header, notes and blank rows, partial or unusual units, text values in the first data row,
categorical-first data with a text target, a date unit over dates. For every layout:
  (a) training with the correct options is VERIFIED with the right column names and training rows;
  (b) training with no options is VERIFIED with the right names and rows, or STOP_RETRY whose single [추천], followed once,
      gives exactly that;
  (c) prediction of a new file of the same layout, with a model trained with units and one trained without units,
      using both the plain template and the model's next_commands prediction command, is VERIFIED with every data row
      predicted and no unit / header / notes row predicted as data, or STOP_RETRY whose single [추천] gives exactly that.
"""
import os
import shlex
import sys
from datetime import datetime, timedelta

import numpy as np
import openpyxl
import pandas as pd
import pytest

HERE = os.path.dirname(os.path.abspath(__file__))
KIT = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(HERE))), "CLI", "ml-kit")
sys.path.insert(0, HERE)
sys.path.insert(0, KIT)
from test_cli_smoke import run_raw  # noqa: E402
import ml_predict as mp  # noqa: E402

N_TRAIN, N_PRED = 20, 5

FAMILIES = {
    # regression: two numeric inputs, a category, a numeric target
    "A": {"cols": ["반응온도", "압력", "라인", "수율"], "units": ["°C", "bar", None, "%"],
          "x": "반응온도,압력,라인", "y": "수율", "task": []},
    # categorical first, text target (classification)
    "C": {"cols": ["라인", "반응온도", "압력", "등급"], "units": [None, "°C", "bar", None],
          "x": "라인,반응온도,압력", "y": "등급", "task": ["--task", "classification"]},
    # a date input (unit 'yyyy-mm-dd' over date cells), '-' as a blank unit
    "D": {"cols": ["일자", "반응온도", "압력", "라인", "수율"], "units": ["yyyy-mm-dd", "°C", "bar", "-", "%"],
          "x": "일자,반응온도,압력,라인", "y": "수율", "task": []},
}


def _data(fam, n, seed):
    r = np.random.default_rng(seed)
    out = []
    for i in range(n):
        t, p, line = round(float(r.uniform(150, 200)), 1), round(float(r.uniform(1, 5)), 2), str(r.choice(["A", "B"]))
        y = round(0.3 * t + 5 * p + (3 if line == "A" else 0) + float(r.normal(0, 0.5)), 2)
        if fam == "A":
            out.append([t, p, line, y])
        elif fam == "C":
            out.append([line, t, p, "상" if t + 10 * p > 205 else "하"])
        else:
            out.append([datetime(2026, 1, 1) + timedelta(days=i + 40 * seed), t, p, line, y])
    return out


def _set(data, row, col, value):
    data[row][col] = value
    return data


H, U = FAMILIES["A"]["cols"], FAMILIES["A"]["units"]
# id: (family, top rows, data edit, merges, correct options, unit row used for the names or None, file extension)
LAYOUTS = {
    "L1_hdr": ("A", [H], None, [], [], None, ".csv"),
    "L2_hdr_unit": ("A", [H, U], None, [], ["--unit-row", "2"], U, ".xlsx"),
    "L3_title_hdr_unit": ("A", [["반응 공정 데이터 2026"], [None], H, U], None, [],
                          ["--header-row", "3", "--unit-row", "4"], U, ".xlsx"),
    "L4_merged": ("A", [["공정 조건", None, None, "결과"], H, U], None, ["A1:C1"],
                  ["--header-row", "2", "--unit-row", "3"], U, ".xlsx"),
    "L5_unit_NG_first": ("A", [H, U], lambda d: _set(_set(d, 0, 0, "측정불가"), 1, 1, "NG"), [],
                         ["--unit-row", "2"], U, ".xlsx"),
    "L6_cat_first": ("C", [FAMILIES["C"]["cols"]], None, [], [], None, ".xlsx"),
    "L6b_cat_first_unit": ("C", [FAMILIES["C"]["cols"], FAMILIES["C"]["units"]], None, [],
                           ["--unit-row", "2"], FAMILIES["C"]["units"], ".xlsx"),
    "L7_blank_before": ("A", [H, U, [None], [None]], None, [], ["--unit-row", "2"], U, ".xlsx"),
    "L8_notes": ("A", [H, ["비고: 2026년 1월 측정", None, None, None]], None, [], ["--data-row", "3"], None, ".xlsx"),
    "L9_partial_unit": ("A", [H, ["°C", None, None, None]], None, [], ["--unit-row", "2"],
                        ["°C", None, None, None], ".xlsx"),
    "L10_degC": ("A", [H, ["℃", "bar", "-", "%"]], None, [], ["--unit-row", "2"], ["℃", "bar", "-", "%"], ".xlsx"),
    "L11_unit_blank_first": ("A", [H, U], lambda d: _set(d, 0, 0, None), [], ["--unit-row", "2"], U, ".xlsx"),
    "L12_pct_text": ("A", [H, U], lambda d: _set(d, 0, 3, "12%"), [], ["--unit-row", "2"], U, ".xlsx"),
    "L13_title_hdr": ("A", [["반응 공정 데이터"], H], None, [], ["--header-row", "2"], None, ".xlsx"),
    "L14_date_unit": ("D", [FAMILIES["D"]["cols"], FAMILIES["D"]["units"]], None, [], ["--unit-row", "2"],
                      FAMILIES["D"]["units"], ".xlsx"),
    "L15_date_hdr": ("D", [FAMILIES["D"]["cols"]], None, [], [], None, ".xlsx"),
}
# models for (c): per family, one trained with units and one without
MODELS = {"A": ("L2_hdr_unit", "L1_hdr"), "C": ("L6b_cat_first_unit", "L6_cat_first"),
          "D": ("L14_date_unit", "L15_date_hdr")}


def _write(path, rows, merges):
    if path.endswith(".csv"):
        import csv
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            csv.writer(f).writerows(rows)
        return path
    wb = openpyxl.Workbook()
    ws = wb.active
    for row in rows:
        ws.append(row)
    for m in merges:
        ws.merge_cells(m)
    wb.save(path)
    return path


def _make(folder, lid, n, seed):
    """(file, 1-based source rows of the data, of them the rows with a valid target)"""
    fam, top, edit, merges, _, _, ext = LAYOUTS[lid]
    data = _data(fam, n, seed)
    if edit:
        data = edit(data)
    path = _write(os.path.join(str(folder), lid + ext), [list(r) for r in top] + data, merges)
    rows = list(range(len(top) + 1, len(top) + n + 1))
    y = FAMILIES[fam]["cols"].index(FAMILIES[fam]["y"])
    valid = [rw for rw, d in zip(rows, data) if FAMILIES[fam]["task"] or mp.coerce_numeric(pd.Series([d[y]])).notna().all()]
    return path, rows, valid


def _names(lid, with_units):
    fam, units = LAYOUTS[lid][0], LAYOUTS[lid][5]
    cols = FAMILIES[fam]["cols"]
    if with_units and units:
        return [mp._with_unit(c, u) for c, u in zip(cols, units)]
    return list(cols)


def _recommendation(message):
    assert message.count("[추천]") == 1, message
    assert "[후보]" not in message, message
    return message.split("[추천] ")[1].split("\n")[0].strip()


def _follow(script, args, message):
    """The command after following the single [추천] once: a full command, or option values put into the command."""
    rec = _recommendation(message)
    if rec.startswith("python"):
        parts = shlex.split(rec)
        return os.path.basename(parts[1]), parts[2:]
    toks = shlex.split(rec)
    assert toks and toks[0].startswith("--"), f"[추천] is not a command or an option: {rec}"
    new = list(args)
    for i in range(0, len(toks), 2):
        if toks[i] in new:
            new[new.index(toks[i]) + 1] = toks[i + 1]
        else:
            new += toks[i:i + 2]
    return script, new


def _run_with_one_retry(script, args):
    """(first payload, final payload): the [추천] of a STOP_RETRY is followed once."""
    _, first = run_raw(script, *args)
    if first["status"] == "ok":
        return first, first
    assert first["agent_directive"].startswith("STOP_RETRY"), first
    if "[추천]" not in first["message"]:
        return first, first          # nothing to follow (2.9.34: L12's mixed '%' target asks the user)
    script2, args2 = _follow(script, args, first["message"])
    _, final = run_raw(script2, *args2)
    return first, final


def _job(fn, *a):
    try:
        return fn(*a)
    except BaseException as e:   # reported by the test of this case
        return e


def _train(path, lid, opts):
    fam = FAMILIES[LAYOUTS[lid][0]]
    return _run_with_one_retry("ml_train.py", ["-f", path, "-x", fam["x"], "-y", fam["y"], *fam["task"], *opts])


def _predict_command(train_res):
    cmds = [c["command"] for c in train_res.get("next_commands", []) if "ml_predict.py" in c["command"]]
    assert len(cmds) == 1, train_res.get("next_commands")
    return cmds[0]


PRED_CASES = [(lid, kind, how) for lid in LAYOUTS for kind in ("units", "plain") for how in ("template", "next")
              if not (kind == "plain" and how == "next")]    # the unit-less models were trained without options


def _predict(folder, lid, kind, how, model_res):
    path = _make(folder, lid, N_PRED, 2)[0]                  # its own copy: outputs are named by the second
    if isinstance(model_res, BaseException):
        raise model_res
    trained = model_res[1]
    if how == "template":
        args = ["-m", trained["version_dir"], "-f", path]
    else:
        args = shlex.split(_predict_command(trained).replace("<새데이터.xlsx>", path))[2:]
    return _run_with_one_retry("ml_predict.py", args)


def run_matrix(mktemp):
    """Every run of the matrix, 4 at a time (separate processes, ~1 minute): training with the correct options,
    then training without options (same file, next version) and the predictions with the models of the first round.
    `mktemp(name)` makes a new folder. Each result is (first payload, final payload) or the exception it raised."""
    from concurrent.futures import ThreadPoolExecutor
    files, out = {}, {}
    for lid in LAYOUTS:
        folder = mktemp(lid.split("_")[0] + "t")
        files[lid] = _make(folder, lid, N_TRAIN, 1)
    with ThreadPoolExecutor(max_workers=4) as pool:
        futs = {("a", lid): pool.submit(_job, _train, files[lid][0], lid, LAYOUTS[lid][4]) for lid in LAYOUTS}
        out.update({k: f.result() for k, f in futs.items()})
        futs = {("b", lid): pool.submit(_job, _train, files[lid][0], lid, []) for lid in LAYOUTS}
        for lid, kind, how in PRED_CASES:
            model_lid = MODELS[LAYOUTS[lid][0]][0 if kind == "units" else 1]
            folder = mktemp(lid.split("_")[0] + "p")
            futs[("c", lid, kind, how)] = pool.submit(_job, _predict, folder, lid, kind, how, out[("a", model_lid)])
        out.update({k: f.result() for k, f in futs.items()})
    out["files"] = files
    return out


@pytest.fixture(scope="module")
def matrix(tmp_path_factory):
    return run_matrix(tmp_path_factory.mktemp)


def _result(matrix, key):
    res = matrix[key]
    if isinstance(res, BaseException):
        raise res
    return res


def _trained_layout(res):
    """(feature + target names, numeric features, 1-based source rows trained on)"""
    assert res["status"] == "ok" and res["agent_directive"].startswith("VERIFIED"), res
    ft = res["feature_types"]
    names = sorted(list(ft["numeric"]) + list(ft["categorical"]) + list(ft["datetime"]) + list(res["targets"]))
    ds = mp.load_data_file(os.path.join(res["version_dir"], "model_dataset.csv"))
    tgt = ds[res["targets"][0]]
    used = ds.loc[tgt.notna() & (tgt.astype(str).str.strip() != ""), "원본행"]
    return names, sorted(ft["numeric"]), sorted(int(v) for v in used)


def _expect(lid, with_units):
    fam = FAMILIES[LAYOUTS[lid][0]]
    names = _names(lid, with_units)
    numeric = sorted(n for n, c in zip(names, fam["cols"]) if c in ("반응온도", "압력"))
    return sorted(names), numeric


def _mixed_percent_stops(final, lid):
    """2.9.34: L12's target holds '12%' next to plain numbers: no scale is sure, STOP_RETRY (was read as 12)."""
    if lid != "L12_pct_text":
        return False
    assert final["agent_directive"].startswith("STOP_RETRY:") and "%를 붙인 숫자 칸('12%')" in final["message"], final
    return True


@pytest.mark.parametrize("lid", list(LAYOUTS))
def test_train_with_correct_options(matrix, lid):
    first, final = _result(matrix, ("a", lid))
    if _mixed_percent_stops(final, lid):
        return
    assert first is final, first.get("message")             # no [추천] detour with the right options
    names, numeric, used = _trained_layout(final)
    assert (names, numeric) == _expect(lid, True)
    assert used == matrix["files"][lid][2]


@pytest.mark.parametrize("lid", list(LAYOUTS))
def test_train_without_options(matrix, lid):
    first, final = _result(matrix, ("b", lid))
    if _mixed_percent_stops(final, lid):
        return
    names, numeric, used = _trained_layout(final)
    assert (names, numeric) in (_expect(lid, True), _expect(lid, False)), first.get("message")
    assert used == matrix["files"][lid][2]


@pytest.mark.parametrize("lid,kind,how", PRED_CASES)
def test_predict_same_layout(matrix, lid, kind, how):
    first, final = _result(matrix, ("c", lid, kind, how))
    assert final["status"] == "ok" and final["agent_directive"].startswith("VERIFIED"), (first, final)
    rows = len(range(len(LAYOUTS[lid][1]) + 1, len(LAYOUTS[lid][1]) + N_PRED + 1))
    assert final["sample_count"] == rows, (first.get("message"), final.get("preview"))
    assert len(mp.load_data_file(final["output_file"])) == rows


# unit-row rules (the matrix above runs them end to end) ----------------------------------------------------

def test_unit_row_majority_below():
    raw = pd.DataFrame([["반응온도", "압력", "수율"], ["°C", "bar", "%"], ["측정불가", 2.1, 50.0], [181.0, "NG", 51.0],
                        [182.0, 2.3, 52.0], [183.0, 2.4, 53.0]], dtype=object)
    assert mp._looks_like_unit_row(raw, 1)                   # 2.9.16: one text value under a unit refused the row
    raw = pd.DataFrame([["일자", "온도"], ["yyyy-mm-dd", "°C"], [datetime(2026, 1, 1), 180.0],
                        [datetime(2026, 1, 2), 181.0]], dtype=object)
    assert mp._looks_like_unit_row(raw, 1)                   # dates are data
    raw = pd.DataFrame([["라인", "온도", "등급"], ["B", 182.8, "중"], ["A", 193.7, "상"], ["A", 190.1, "중"]], dtype=object)
    assert not mp._looks_like_unit_row(raw, 1)               # categories over categories
    raw = pd.DataFrame([["온도", "압력"], ["비고: 2026년 1월 측정", None], [180.0, 3.0]], dtype=object)
    assert not mp._looks_like_unit_row(raw, 1)               # a notes row
    raw = pd.DataFrame([["공정", "조건", "결과"], ["온도", "압력", "수율"], ["°C", "bar", "%"], [180.0, 3.0, 50.0],
                        [181.0, 3.1, 51.0]], dtype=object)
    assert not mp._looks_like_unit_row(raw, 1)               # a header over a unit row
    assert mp._looks_like_unit_row(raw, 2)


def test_title_row_is_not_a_header_over_units():
    raw = pd.DataFrame([["시험 기록", "작성자 홍길동", None], ["온도", "압력", "수율"], [175.2, 1.6, 55.7],
                        [171.8, 2.2, 56.2]], dtype=object)
    assert mp._header_unit_pair(raw) is None                 # 2.9.16: (1, 2), two [추천] hops
    raw = pd.DataFrame([["새 조건", "작성 2026-11-02", None], ["반응온도", "압력", "유량"], ["°C", "bar", "L/min"],
                        [180.0, 3.0, 10.0]], dtype=object)
    assert mp._header_unit_pair(raw) == (2, 3)


def test_given_unit_row_blank_or_not_below_header(tmp_path):
    path = _write(str(tmp_path / "t.xlsx"), [["반응 공정 데이터"], [None], H, U] + _data("A", 12, 3), [])
    for bad in ("2", "3", "1"):                              # blank, the header row, above the header
        with pytest.raises(mp.InputError) as e:
            mp.load_data_file(path, unit_row=int(bad))
        assert f"--unit-row {bad}" in str(e.value) and "--header-row 3 --unit-row 4" in str(e.value)


def test_given_unit_row_is_the_header_under_a_title(tmp_path):
    rows = [["시험 기록", "작성자 홍길동", None], ["온도", "압력", "수율"]] + [[170.0 + i, 1.0 + i / 10, 50.0 + i] for i in range(12)]
    path = _write(str(tmp_path / "t2.xlsx"), rows, [])
    with pytest.raises(mp.InputError) as e:
        mp.load_data_file(path, unit_row=2)                  # 2.9.16: [추천] header 1 + unit 2, then a second hop
    assert str(e.value).count("[추천]") == 1 and str(e.value).rstrip().endswith("[추천] --header-row 2")
