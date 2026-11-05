# -*- coding: utf-8 -*-
"""데이터 읽기 1: 머리글 행, 단위 행, 데이터 시작 행, 제목 행, 열 이름(빈 열, 중복, 연도 숫자), 날짜, 원본 행 번호.

레이아웃이 틀리면 도구가 [추천](옵션 값 또는 전체 명령)을 하나 주고, 그 [추천]을 그대로 한 번 실행하면 성공해야 한다.
끝까지 실행하는 레이아웃 검사는 test_layout_matrix.py.
"""
import argparse
import copy
import json
import os
import sys
from datetime import datetime

import numpy as np
import openpyxl
import pandas as pd
import pytest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from mlkit_helpers import (split_command, H, U, family_rows, recommendation, recommended_command, run_fail, run_raw,  # noqa: E402
                           run_tool, write_book, write_csv_header, write_df, write_rows)
import ml_predict as mp  # noqa: E402


# ---- 단위 행 판정 규칙 --------------------------------------------------------------------------------------------


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


def test_looks_like_unit_row():
    raw = pd.DataFrame([["라인", "온도", "등급"], ["B", 182.8, "중"], ["A", 193.7, "상"]], dtype=object)
    assert not mp._looks_like_unit_row(raw, 1)                     # categories over categories
    raw = pd.DataFrame([["라인", "온도", "등급"], ["-", "°C", None], ["A", 193.7, "상"]], dtype=object)
    assert mp._looks_like_unit_row(raw, 1)                         # units over numbers, '-' is a blank mark
    raw = pd.DataFrame([["제목", None], ["온도", "압력"], ["°C", "bar"], [180, 3]], dtype=object)
    assert not mp._looks_like_unit_row(raw, 1)                     # a header over a unit row
    assert mp._looks_like_unit_row(raw, 2)
    raw = pd.DataFrame([["온도", "압력"], [180, 3], [185, 3.2]], dtype=object)
    assert not mp._looks_like_unit_row(raw, 1)                     # numbers


def test_unit_row_has_no_numbers():
    raw = pd.DataFrame([H, ["측정불가", "측정불가", "A", 61.66], [180.0, 3.91, "B", 73.39], [152.8, 2.1, "A", 59.83],
                        [157.5, 2.73, "B", 60.5]], dtype=object)
    assert not mp._looks_like_unit_row(raw, 1)     # 2.9.17: accepted, names '반응온도(측정불가)'
    raw = pd.DataFrame([H, U, [None, None, None, 61.66], [180.0, 3.91, "B", 73.39], [152.8, 2.1, "A", 59.83]],
                       dtype=object)
    assert mp._looks_like_unit_row(raw, 1)


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


def _ng_book(path, units):
    r = np.random.default_rng(3)
    rows = [["농도A", "농도B", "수율"], units]
    for _ in range(20):
        a, b = round(float(r.uniform(150, 160)), 1), round(float(r.uniform(1, 3)), 2)
        rows.append([a, b, round(0.3 * a + 5 * b + float(r.normal(0, 0.5)), 2)])
    return write_rows(str(path), rows, [])


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


# ---- 제목 행과 머리글 + 단위 행 쌍 ----------------------------------------------------------------------------


def test_title_row_is_not_a_header_over_units():
    raw = pd.DataFrame([["시험 기록", "작성자 홍길동", None], ["온도", "압력", "수율"], [175.2, 1.6, 55.7],
                        [171.8, 2.2, 56.2]], dtype=object)
    assert mp._header_unit_pair(raw) is None                 # 2.9.16: (1, 2), two [추천] hops
    raw = pd.DataFrame([["새 조건", "작성 2026-11-02", None], ["반응온도", "압력", "유량"], ["°C", "bar", "L/min"],
                        [180.0, 3.0, 10.0]], dtype=object)
    assert mp._header_unit_pair(raw) == (2, 3)


def test_given_unit_row_blank_or_not_below_header(tmp_path):
    path = write_rows(str(tmp_path / "t.xlsx"), [["반응 공정 데이터"], [None], H, U] + family_rows("A", 12, 3), [])
    for bad in ("2", "3", "1"):                              # blank, the header row, above the header
        with pytest.raises(mp.InputError) as e:
            mp.load_data_file(path, unit_row=int(bad))
        assert f"--unit-row {bad}" in str(e.value) and "--header-row 3 --unit-row 4" in str(e.value)


def test_given_unit_row_is_the_header_under_a_title(tmp_path):
    rows = [["시험 기록", "작성자 홍길동", None], ["온도", "압력", "수율"]] + [[170.0 + i, 1.0 + i / 10, 50.0 + i] for i in range(12)]
    path = write_rows(str(tmp_path / "t2.xlsx"), rows, [])
    with pytest.raises(mp.InputError) as e:
        mp.load_data_file(path, unit_row=2)                  # 2.9.16: [추천] header 1 + unit 2, then a second hop
    assert str(e.value).count("[추천]") == 1 and str(e.value).rstrip().endswith("[추천] --header-row 2")


def test_header_unit_pair_skips_text_data_rows(tmp_path):
    # a first data row with blank measurements, then a full one: categories over categories, not units
    rows = [["라인", "설비", "온도", "압력", "유량"], ["A", "M1", None, None, None], ["B", "M2", 181.0, 3.0, 5.0]]
    raw = pd.read_excel(write_book(tmp_path / "p.xlsx", rows), header=None)
    assert mp._header_unit_pair(raw) is None                       # 2.9.15: (1, 2), a data row as the unit row
    rows = [["제목"], [None], ["반응온도", "압력", "제외"], ["°C", "bar", None], [192.5, 3.0, None]]
    raw = pd.read_excel(write_book(tmp_path / "b.xlsx", rows), header=None)
    assert mp._header_unit_pair(raw) == (3, 4)                     # beta sample layout


def test_ragged_csv(tmp_path):
    p = tmp_path / "t.csv"
    p.write_text("반응 공정 데이터\n반응온도,압력,라인,수율\n180,3,A,70\n181,3.1,B,71\n", encoding="cp949")
    df = mp.load_data_file(str(p))                  # 2.9.17: "Expected 1 fields in line 2, saw 4"
    assert list(df.columns) == ["반응온도", "압력", "라인", "수율"] and len(df) == 2


def test_model_names_with_units_match_header_without_units():
    cells = ["반응온도", "압력", "라인", "수율"]
    assert mp.names_in_cells(["반응온도(°C)", "압력(bar)", "라인"], cells, model_names=True)
    assert not mp.names_in_cells(["반응온도(°C)"], cells)              # user names (-x): a unit is part of the name
    assert not mp.names_in_cells(["온도(K)"], ["온도(°C)"], model_names=True)
    assert mp.names_in_cells(["온도(C)"], ["온도(°C)"], model_names=True)  # 2.9.33: '°C' / 'C' is one unit (was False)


def test_repeated_header_and_totals_rows(tmp_path):
    rows = [H, U] + [[170.0 + i, 2.0 + i / 10, "A", 60.0 + i] for i in range(4)] + [H, U] + \
        [[175.0 + i, 2.5, "B", 65.0 + i] for i in range(3)] + [[1220.0, 17.1, "합계", 435.0]]
    path = write_rows(str(tmp_path / "r.xlsx"), rows, [])
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


# ---- 학습 레이아웃 옵션과 [추천] --------------------------------------------------------------------------------


@pytest.fixture(scope="module")
def layout(tmp_path_factory):
    work = tmp_path_factory.mktemp("layout")
    r = np.random.default_rng(7)
    n = 60
    a, b = r.uniform(0, 10, n), r.uniform(0, 10, n)
    y = 2 * a + b + r.normal(0, 0.1, n)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["공정 기록"])
    ws.append(["a", "b", "y"])
    ws.append(["mm", "kg", "%"])
    for row in zip(a, b, y):
        ws.append([float(v) for v in row])
    train = work / "train.xlsx"
    wb.save(train)
    res = run_tool("ml_train.py", "-f", str(train), "-x", "a,b", "-y", "y", "--header-row", "2", "--unit-row", "3")
    new = work / "new.csv"
    pd.DataFrame({"a(mm)": [1.0, 2.0], "b(kg)": [3.0, 4.0]}).to_csv(new, index=False)
    return {"work": work, "res": res, "new": new}


def test_training_layout_on_plain_file_is_recovered(layout):
    cmd = [c for c in layout["res"]["next_commands"] if "ml_predict.py" in c["command"]][0]["command"]
    assert "--header-row" not in cmd and "--unit-row" not in cmd                # 2.9.53: the training layout options
    argv = split_command(cmd.replace("<새데이터.xlsx>", str(layout["new"])))[2:]   # drop 'python <script>'
    assert run_tool("ml_predict.py", *argv)["sample_count"] == 2
    # the training options copied by hand: read with the file's own header row (JIT), said in recovered
    ok = run_tool("ml_predict.py", *argv, "--header-row", "2", "--unit-row", "3")   # 2.9.53: STOP_RETRY [추천]
    assert ok["sample_count"] == 2 and ok["recovered"] == ["레이아웃 옵션이 이 파일과 달라 머리글 1행으로 읽음"], ok
    assert ok["agent_directive"].startswith("VERIFIED")
    # a longer plain file: the layout reads a data row as the header -> model columns missing, same recovery
    longer = layout["work"] / "longer.csv"
    pd.DataFrame({"a": np.arange(8.0), "b": np.arange(8.0)}).to_csv(longer, index=False)
    ok = run_tool("ml_predict.py", *[str(longer) if x == str(layout["new"]) else x for x in argv],
                  "--header-row", "2", "--unit-row", "3")
    assert ok["sample_count"] == 8 and "머리글 1행" in ok["recovered"][0]


def test_two_level_header_hint(tmp_path, layout):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["조건", None, "결과"])
    ws.merge_cells("A1:B1")
    ws.append(["온도(°C)", "압력(%)", "수율(%)"])
    r = np.random.default_rng(8)
    for _ in range(30):
        t, pr = float(r.uniform(100, 200)), float(r.uniform(10, 90))
        ws.append([t, pr, 0.2 * t + 0.1 * pr])
    p = tmp_path / "merged.xlsx"
    wb.save(p)
    bad = run_fail("ml_train.py", "-f", str(p), "-x", "온도,압력", "-y", "수율")
    assert "[추천] --header-row 2" in bad["message"]                              # 2.9.12: no hint
    v = run_tool("ml_train.py", "-f", str(p), "-x", "온도,압력", "-y", "수율", "--header-row", "2")["version_dir"]
    ok = run_tool("ml_predict.py", "-m", v, "-f", str(p))                       # 2.9.53: [추천] --header-row 2
    assert ok["sample_count"] == 30 and ok["recovered"] == ["머리글 2행으로 읽음"], ok


def test_several_header_rows_with_every_input_stop_with_candidates(tmp_path):
    rows = [["온도", "압력", "수율"]]
    r = np.random.default_rng(9)
    for _ in range(30):
        t, pr = float(r.uniform(100, 200)), float(r.uniform(10, 90))
        rows.append([t, pr, 0.2 * t + 0.1 * pr])
    v = run_tool("ml_train.py", "-f", write_book(tmp_path / "t.xlsx", rows), "-x", "온도,압력", "-y", "수율")["version_dir"]
    p = write_book(tmp_path / "two.xlsx", [["시험 기록", "작성자 홍"], ["온도", "압력"], ["비고"], ["온도", "압력"],
                                       [150, 30], [160, 31]])
    bad = run_fail("ml_predict.py", "-m", v, "-f", p)
    assert "[후보] --header-row" in bad["message"] and "recovered" not in bad, bad      # the tool does not guess


@pytest.fixture(scope="module")
def heads(tmp_path_factory):
    work = tmp_path_factory.mktemp("heads")
    r = np.random.default_rng(2)
    a, b, y = r.uniform(0, 1, 30), r.uniform(0, 1, 30), r.uniform(0, 1, 30)
    u = write_df(work / "u.csv", pd.DataFrame({"온도(°C)": a, "압력": b, "y": y}))
    u2 = write_df(work / "u2.csv", pd.DataFrame({"온도(C)": a, "온도(K)": b, "y": y}))
    u3 = write_df(work / "u3.csv", pd.DataFrame({"온도": a, "설비": ["B", "압력"] * 15, "y": y}))
    return {"u": u, "u2": u2, "u3": u3}


def test_header_hint_not_the_current_row(heads):
    # 2.9.33: '온도(C)' is '온도(°C)' (the same unit) and matches; a typo '온도x' is a [후보] (was [추천] for '온도(C)')
    bad = run_fail("ml_train.py", "-f", heads["u"], "-x", "온도x,압력", "-y", "y")
    assert "--header-row" not in bad["message"]                                 # 2.9.13: [추천] --header-row 1
    assert "'온도x' -> [후보] '온도(°C)'" in bad["message"]
    bad = run_fail("ml_train.py", "-f", heads["u2"], "-x", "온도", "-y", "y")
    assert "--header-row" not in bad["message"] and "[후보]" in bad["message"]
    bad = run_fail("ml_train.py", "-f", heads["u3"], "-x", "온도,압력", "-y", "y")   # '압력' is a data value
    assert "--header-row" not in bad["message"]                                 # 2.9.13: [추천] --header-row 3
    bad = run_fail("ml_map.py", "-f", heads["u"], "-i", "온도x=0.5")
    assert "--header-row" not in bad["message"] and "[후보] '온도(°C)'" in bad["message"]       # 2.9.33: 온도(C), [추천]


def test_find_header_layout_rules():
    cands = [[1, ["제목"]], [3, ["온도(°C)", "압력"]]]
    assert mp.find_header_layout(cands, ["온도", "압력"]) == (3, None)
    assert mp.find_header_layout(cands, ["온도(K)"]) is None                    # a different unit is not the same name
    assert mp.find_header_layout(cands, ["온도(C)"]) == (3, None)               # '°C' / 'C' is one unit
    assert mp.find_header_layout([[2, ["온도(C)", "온도(K)"]]], ["온도"]) is None   # ambiguous alias


@pytest.fixture(scope="module")
def unit_layout(tmp_path_factory):
    work = tmp_path_factory.mktemp("unitlay")
    r = np.random.default_rng(3)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["제목"])
    ws.append([None])
    ws.append(["온도", "압력", "효율"])
    ws.append(["°C", "bar", "%"])
    for _ in range(40):
        t, p = float(r.uniform(150, 200)), float(r.uniform(1, 5))
        ws.append([t, p, 0.3 * t + 2 * p + float(r.normal(0, 0.2))])
    path = work / "lay.xlsx"
    wb.save(path)
    v = run_tool("ml_train.py", "-f", str(path), "-x", "온도,압력", "-y", "효율",
                 "--header-row", "3", "--unit-row", "4")["version_dir"]
    return {"path": str(path), "v": v}


def test_layout_recovery_keeps_the_unit_row(unit_layout):
    ok = run_tool("ml_predict.py", "-m", unit_layout["v"], "-f", unit_layout["path"], "--header-row", "2", "--unit-row", "4")
    assert ok["sample_count"] == 40 and not ok["missing_input_warning"]          # 2.9.53: STOP_RETRY [추천]
    assert ok["recovered"] == ["레이아웃 옵션이 이 파일과 달라 머리글 3행으로 읽음, 단위 4행"], ok
    same = run_tool("ml_predict.py", "-m", unit_layout["v"], "-f", unit_layout["path"], "--header-row", "3", "--unit-row", "4")
    assert "recovered" not in same                                               # options that fit are kept


def test_predict_command_absolute_quoted_paths():
    ns = argparse.Namespace(model="rel_dir", meta=None, targets=None, file="새 파일.xlsx", out=None, encoding=None,
                            overwrite=False)
    cmd = mp._predict_command(ns)
    assert mp.quote_arg(os.path.abspath("rel_dir")) in cmd and mp.quote_arg(os.path.abspath("새 파일.xlsx")) in cmd


@pytest.fixture(scope="module")
def unit_book(tmp_path_factory):
    work = tmp_path_factory.mktemp("units")
    r = np.random.default_rng(5)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["반응 기록"])
    ws.append([None])
    ws.append(["반응온도", "압력", "효율"])
    ws.append(["°C", "bar", "%"])
    for _ in range(40):
        t, p = float(r.uniform(150, 200)), float(r.uniform(1, 5))
        ws.append([round(t, 2), round(p, 3), round(0.3 * t + 2 * p + float(r.normal(0, 0.2)), 3)])
    path = work / "u.xlsx"
    wb.save(path)
    v = run_tool("ml_train.py", "-f", str(path), "-x", "반응온도,압력", "-y", "효율",
                 "--header-row", "3", "--unit-row", "4")["version_dir"]
    return {"work": work, "path": str(path), "v": v}


def test_numeric_unit_row_refused(unit_book):
    bad = run_fail("ml_train.py", "-f", unit_book["path"], "-x", "반응온도,압력", "-y", "효율",
                   "--header-row", "4", "--unit-row", "5")
    assert "단위 행(--unit-row 5)" in bad["message"] and bad["agent_directive"].startswith("STOP_RETRY")
    script, args = recommended_command(bad["message"])
    assert script == "ml_train.py" and args[-4:] == ["--header-row", "3", "--unit-row", "4"]
    res = run_tool(script, *args)
    meta = json.load(open(os.path.join(res["version_dir"], "model_효율(%).json"), encoding="utf-8"))
    assert meta["features_required"] == ["반응온도(°C)", "압력(bar)"]              # 2.9.14: '반응온도(192.58)'


def test_header_hint_carries_the_unit_row(unit_book):
    # --header-row 2 --unit-row 3 (both one row too high): the names are in row 3, the units in row 4
    bad = run_fail("ml_train.py", "-f", unit_book["path"], "-x", "반응온도,압력", "-y", "효율",
                   "--header-row", "2", "--unit-row", "3")
    _, args = recommended_command(bad["message"])
    assert args[-4:] == ["--header-row", "3", "--unit-row", "4"]                  # 2.9.14: [추천] --header-row 3
    assert "--data-row" not in args


def test_title_row_hint_carries_the_unit_row(tmp_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["새 조건", "작성 2026-11-02"])
    ws.append(["온도", "압력", "효율"])
    ws.append(["°C", "bar", "%"])
    r = np.random.default_rng(6)
    for _ in range(30):
        t, p = float(r.uniform(150, 200)), float(r.uniform(1, 5))
        ws.append([t, p, 0.3 * t + 2 * p])
    p = tmp_path / "t.xlsx"
    wb.save(p)
    bad = run_fail("ml_train.py", "-f", str(p), "-x", "온도,압력", "-y", "효율")
    _, args = recommended_command(bad["message"])
    assert args[-4:] == ["--header-row", "2", "--unit-row", "3"]                  # 2.9.14: --header-row 2 only


def test_predict_finds_model_names_with_units_in_a_header_unit_pair(unit_book, tmp_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["새 조건", "작성 2026-11-02"])
    ws.append(["반응온도", "압력"])
    ws.append(["°C", "bar"])
    ws.append([180, 3])
    ws.append([185, 3.2])
    p = tmp_path / "t2c.xlsx"
    wb.save(p)
    ok = run_tool("ml_predict.py", "-m", unit_book["v"], "-f", str(p))           # 2.9.53: [추천] --header-row 2 --unit-row 3
    assert ok["sample_count"] == 2 and not ok["missing_input_warning"]
    assert ok["recovered"] == ["머리글 2행, 단위 3행으로 읽음"], ok


def test_typo_hints_skip_layout_artefacts():
    hints = mp._typo_hints(["압력(bar)"], ["bar(3)", "°C(180)"], ["압력(bar)", "반응온도(°C)"], False)
    assert hints == []                                                             # 2.9.14: 'bar(3)' -> '압력(bar)'
    assert mp._typo_hints(["압력(bar)"], ["압력(bra)"], ["압력(bar)"], False)


def test_predict_plain_template_on_a_unit_row_file(unit_book, tmp_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["반응온도", "압력"])
    ws.append(["°C", "bar"])
    ws.append([180, 3])
    ws.append([185, 3.2])
    p = tmp_path / "u2.xlsx"
    wb.save(p)
    ok = run_tool("ml_predict.py", "-m", unit_book["v"], "-f", str(p))           # 2.9.14: unit row imputed; 2.9.53: [추천]
    assert ok["sample_count"] == 2 and not ok["missing_input_warning"]
    assert ok["recovered"] == ["머리글 1행, 단위 2행으로 읽음"], ok


def test_placeholder_first_row_is_not_a_unit_row(unit_book, tmp_path):
    df = pd.DataFrame({"반응온도": ["-", 180, 185], "압력": ["-", 3, 3.2]})
    ok = run_tool("ml_predict.py", "-m", unit_book["v"], "-f", write_df(tmp_path / "ph.csv", df))
    assert ok["sample_count"] == 3


# ---- 글자 첫 데이터 행은 단위 행이 아님, 단위 없이 학습한 모델에 단위 행 파일 --------------------------


def _two_level_rows(n=40):
    r = np.random.default_rng(11)
    rows = [["조건", None, None, None, "결과"], ["라인", "설비", "작업자", "온도", "등급"]]
    for _ in range(n):
        t = float(r.uniform(170, 200))
        rows.append([str(r.choice(["A", "B"])), str(r.choice(["M1", "M2", "M3"])), str(r.choice(["kim", "lee"])),
                     round(t, 2), "상" if t > 188 else ("중" if t > 178 else "하")])
    return rows


@pytest.fixture(scope="module")
def two_level(tmp_path_factory):
    return write_book(tmp_path_factory.mktemp("two") / "two.xlsx", _two_level_rows())


def test_text_first_row_header_hint_has_no_unit_row(two_level):
    bad = run_fail("ml_train.py", "-f", two_level, "-x", "라인,설비,작업자,온도", "-y", "등급", "--task", "classification")
    assert recommendation(bad["message"]) == "--header-row 2"     # 2.9.15: --header-row 2 --unit-row 3
    assert "단위 행" not in bad["message"]
    res = run_tool("ml_train.py", "-f", two_level, "-x", "라인,설비,작업자,온도", "-y", "등급", "--task", "classification",
                   "--header-row", "2")
    assert res["targets"] == ["등급"]


def test_text_first_row_given_as_unit_row_is_refused(two_level):
    bad = run_fail("ml_train.py", "-f", two_level, "-x", "라인,설비,작업자,온도", "-y", "등급", "--task", "classification",
                   "--header-row", "2", "--unit-row", "3")             # 2.9.15: VERIFIED, '온도(182.8)', '등급(중)'
    assert "단위 행(--unit-row 3)" in bad["message"] and bad["agent_directive"].startswith("STOP_RETRY")
    script, args = recommended_command(bad["message"])
    assert "--unit-row" not in args and args[-2:] == ["--header-row", "2"]
    res = run_tool(script, *args)
    assert res["targets"] == ["등급"] and sorted(res["feature_types"]["categorical"]) == ["라인", "설비", "작업자"]


@pytest.fixture(scope="module")
def plain_model(tmp_path_factory):
    work = tmp_path_factory.mktemp("plain")
    r = np.random.default_rng(12)
    rows = [["반응온도", "압력", "효율"]]
    for _ in range(40):
        t, p = float(r.uniform(150, 200)), float(r.uniform(1, 5))
        rows.append([round(t, 2), round(p, 3), round(0.3 * t + 2 * p + float(r.normal(0, 0.2)), 3)])
    path = write_book(work / "plain.xlsx", rows)
    return run_tool("ml_train.py", "-f", path, "-x", "반응온도,압력", "-y", "효율")["version_dir"]


def test_unit_row_file_for_a_unitless_model(plain_model, tmp_path):
    p = write_book(tmp_path / "u2.xlsx", [["반응온도", "압력"], ["°C", "bar"], [180, 3], [185, 3.2]])
    ok = run_tool("ml_predict.py", "-m", plain_model, "-f", p)      # 2.9.15: unit row imputed; 2.9.53: [추천] --data-row 3
    assert ok["sample_count"] == 2 and not ok["missing_input_warning"]
    assert ok["recovered"] == ["머리글 1행, 데이터 3행부터로 읽음"], ok


def test_title_and_unit_row_file_for_a_unitless_model(plain_model, tmp_path):
    p = write_book(tmp_path / "t2c.xlsx", [["새 조건", "작성 2026-11-02"], ["반응온도", "압력"], ["°C", "bar"],
                                      [180, 3], [185, 3.2]])
    ok = run_tool("ml_predict.py", "-m", plain_model, "-f", p)       # 2.9.53: [추천] --header-row 2 --data-row 4
    assert ok["sample_count"] == 2 and not ok["missing_input_warning"]
    assert ok["recovered"] == ["머리글 2행, 데이터 4행부터로 읽음"], ok


def test_first_data_row_with_text_marks_is_kept(p1_model, tmp_path):
    """2.9.59: 첫 데이터 행에 '점검중'·'센서고장'이 섞였다고 단위·메모 행으로 보고 그 데이터 행을 버렸다."""
    body = [[180.5, 3.1, 45.0], [185.0, 3.3, 47.5], [176.0, 2.6, 41.0]]
    p = write_book(tmp_path / "marks.xlsx", [["온도(°C)", "압력(bar)", "속도"], ["점검중", "센서고장", 45.0]] + body)
    ok = run_tool("ml_predict.py", "-m", p1_model["v"], "-f", p)
    assert ok["sample_count"] == 4 and "recovered" not in ok, ok          # 숫자가 든 행은 데이터
    assert ok["missing_input_warning"] is True, ok
    p = write_book(tmp_path / "marks2.xlsx", [["온도(°C)", "압력(bar)", "속도"], ["점검중", "센서고장", "점검중"]] + body)
    ok = run_tool("ml_predict.py", "-m", p1_model["v"], "-f", p)
    assert ok["sample_count"] == 4 and "recovered" not in ok, ok          # 단위 모양이 아닌 글자 행도 데이터


@pytest.mark.parametrize("units", [["도씨", "킬로그램", "시간", "킬로그램"], ["℃", "기압", "분", "킬로그램"],
                                   [None, "킬로그램", None, "킬로그램"], ["섭씨", "기압", "회전", "평"]])
def test_korean_unit_words_in_named_unit_row(tmp_path, units):
    """2.9.61: 목록에 없는 한글 단위('도씨', '킬로그램', '기압')를 메모로 보아 --unit-row 2가 STOP_RETRY(틀린 [추천]),
    예측은 단위 행을 데이터로 예측했다(VERIFIED). 사용자가 준 단위 행의 짧은 한글 낱말은 단위다."""
    r = np.random.default_rng(8)
    n = 40
    rows = [["온도", "압력", "시간", "생산량"], units]
    for _ in range(n):
        t, p, h = r.uniform(150, 200), r.uniform(1, 5), r.uniform(1, 5)
        rows.append([round(t, 2), round(p, 3), round(h, 2), round(2 * t + 10 * p + 3 * h, 2)])
    f = str(tmp_path / "k.csv")
    write_rows(f, rows, [])
    raw = mp.read_raw_table(f)
    assert mp.unit_row_verdict(raw, 1, 0, given=True)[0]
    assert not mp.unit_row_verdict(raw, 1, 0)[0]          # the tool's own guess keeps Hangul words as data
    res = run_tool("ml_train.py", "-f", f, "-x", "온도,압력,시간", "-y", "생산량", "--unit-row", "2")
    ok = run_tool("ml_predict.py", "-m", res["version_dir"], "-f", f)
    assert ok["sample_count"] == n, ok
    assert ok["recovered"] == ["머리글 1행, 단위 2행으로 읽음"], ok


def _unit_rows_file(tmp_path, units, name="k.csv"):
    r = np.random.default_rng(9)
    rows = [["온도", "압력", "시간", "생산량"], units]
    for _ in range(40):
        t, p, h = r.uniform(150, 200), r.uniform(1, 5), r.uniform(1, 5)
        rows.append([round(t, 2), round(p, 3), round(h, 2), round(2 * t + 10 * p + 3 * h, 2)])
    f = str(tmp_path / name)
    write_rows(f, rows, [])
    return f


def test_data_row_given_as_unit_row_is_refused(tmp_path):
    """3차 검토(2.9.63): 데이터 행('점검중, 센서고장, 정지')을 --unit-row로 주면 받아 '온도(점검중)'이 되었다.
    단위가 아닌 한글 낱말이 글자 칸의 절반 이상이면 단위 행이 아니다 (STOP_RETRY [추천] --unit-row를 뺀 명령)."""
    f = _unit_rows_file(tmp_path, ["점검중", "센서고장", "정지", None])
    raw = mp.read_raw_table(f)
    ok, why = mp.unit_row_verdict(raw, 1, 0, given=True)
    assert not ok and "한글 낱말" in why
    bad = run_fail("ml_train.py", "-f", f, "-x", "온도,압력,시간", "-y", "생산량", "--unit-row", "2")
    assert bad["agent_directive"].startswith("STOP_RETRY") and "단위 행이 아닙니다" in bad["message"], bad
    rec = bad["message"].split("[추천] ", 1)[1]
    assert rec.startswith("python ") and "--unit-row" not in rec, rec


@pytest.mark.parametrize("units", [
    ["세", "백만원", "cm", None], ["억원", "백만원", "km", None], ["킬로", "km", "킬로", None],
    ["기가바이트", "Mbps", "초", None], ["개수", "EA", "센티", None], ["점", "퍼센트포인트", "마리", None],
    ["킬로", None, None, None], ["명", "건", "대", "배"],
    ["킬로그램중", "mm", "도씨", None], ["톤중", "그램중", None, None],            # 5차 검토: 중량 단위는 "중"으로 끝나도 상태 낱말이 아님
])
def test_named_unit_row_with_korean_units_is_read(tmp_path, units):
    """4차 검토(2.9.63 회귀): --unit-row로 준 정상 한글 단위 행(세/백만원/cm, 억원, 킬로/km, 기가바이트/Mbps, 개수/EA, 초/센티,
    점/퍼센트포인트, 킬로 한 칸)을 낱말 목록에 없다고 거부했다. 아래 열이 숫자이고 짧은 낱말이면 단위다 (경고도 없음)."""
    f = _unit_rows_file(tmp_path, units)
    raw = mp.read_raw_table(f)
    ok, why = mp.unit_row_verdict(raw, 1, 0, given=True)
    assert ok, why
    df = mp.load_data_file(f, unit_row=2)
    assert len(df) == 40 and not [w for w in mp.layout_warnings(df) if "단위로 읽음" in w]
    named = [c for c in df.columns if "(" in str(c)]
    assert len(named) == sum(u is not None for u in units), list(df.columns)


def test_some_odd_words_in_unit_row_are_read_with_a_warning(tmp_path):
    """단위 행의 일부 칸만 단위가 아닌 한글 낱말이면 받아들이되 warnings로 알린다."""
    f = _unit_rows_file(tmp_path, ["도씨", "점검중", "시간", "킬로그램"])
    res = run_tool("ml_train.py", "-f", f, "-x", "온도,압력,시간", "-y", "생산량", "--unit-row", "2")
    w = [x for x in res.get("warnings", []) if x.startswith("2행을 단위로 읽음")]
    assert w and "압력(점검중)" in w[0] and "--unit-row를 빼세요" in w[0], res.get("warnings")
    clean = run_tool("ml_train.py", "-f", _unit_rows_file(tmp_path, ["도씨", "기압", "시간", "킬로그램"], "c.csv"),
                     "-x", "온도,압력,시간", "-y", "생산량", "--unit-row", "2")
    assert not any("단위로 읽음" in x for x in clean.get("warnings", [])), clean.get("warnings")


def test_skipped_unit_row_is_in_warnings(p1_model, tmp_path):
    body = [[180.5, 3.1, 45.0], [185.0, 3.3, 47.5]]
    p = write_book(tmp_path / "u.xlsx", [["온도", "압력", "속도"], ["kg/cm2 G", "kg/cm2 G", "m/s"]] + body)
    ok = run_tool("ml_predict.py", "-m", p1_model["v"], "-f", p)
    assert ok["sample_count"] == 2, ok
    assert ok["recovered"] == ["머리글 1행, 데이터 3행부터로 읽음"], ok             # 옵션을 주지 않았으니 '옵션이 달라'가 아님
    assert any("2행" in w and "건너" in w for w in ok.get("warnings", [])), ok


# ---- 열 이름: 빈 열, 중복 머리글, 연도 숫자 머리글 --------------------------------------------------------


def test_named_empty_column_is_kept_and_unnamed_dropped(tmp_path):
    p = tmp_path / "e.csv"
    p.write_text("온도1,온도2,,y\n,1,,2\n,2,,4\n,3,,6\n", encoding="utf-8")
    df = mp.load_data_file(str(p))
    assert list(df.columns) == ["온도1", "온도2", "y"]
    bad = run_fail("ml_train.py", "-f", str(p), "-x", "온도1", "-y", "y")
    assert bad["agent_directive"].startswith("STOP_RETRY")
    assert "값이 하나도 없습니다" in bad["message"] and "[추천]" not in bad["message"]


def test_duplicate_headers_become_unique(tmp_path):
    p = tmp_path / "dup.csv"
    p.write_text("a,a,a_1,a\n1,2,3,4\n", encoding="utf-8")
    cols = list(mp.load_data_file(str(p)).columns)
    assert len(set(cols)) == 4 and cols[0] == "a" and cols[2] == "a_1"


def test_year_headers_are_integers(tmp_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["구분", 2023, 2024])
    for i in range(5):
        ws.append([f"L{i}", i * 1.5, i * 2.5])
    p = tmp_path / "y.xlsx"
    wb.save(p)
    assert list(mp.load_data_file(str(p)).columns) == ["구분", "2023", "2024"]   # 2.9.14: '2023.0'


# ---- 날짜 ------------------------------------------------------------------------------------------------------------


def test_parse_datetime_reads_yyyymmdd():
    s = pd.Series([20240103, "20240215", 20240301.0, "2024-04-05", "20241340", None, "abc"], dtype=object)
    got = mp.parse_datetime(s)
    assert list(got[:4]) == [pd.Timestamp("2024-01-03"), pd.Timestamp("2024-02-15"),
                             pd.Timestamp("2024-03-01"), pd.Timestamp("2024-04-05")]
    assert got[4:].isna().all()


def test_parse_datetime_text_with_point_zero():
    out = mp.parse_datetime(pd.Series(["20240103.0", 20240103.0, "20240103", "2024.01.03"], dtype=object))
    assert all(d == pd.Timestamp("2024-01-03") for d in out)


def test_dtypes_datetime_yyyymmdd_predicts_without_imputation(tmp_path):
    r = np.random.default_rng(1)
    n = 200
    days = pd.date_range("2024-01-01", periods=n, freq="D")
    df = pd.DataFrame({"일자": days.strftime("%Y%m%d"), "a": r.uniform(0, 1, n)})
    df["y"] = 0.05 * np.arange(n) + 2 * df.a + r.normal(0, 0.05, n)
    p = tmp_path / "d.csv"
    df.to_csv(p, index=False)
    v = run_tool("ml_train.py", "-f", str(p), "-x", "일자,a", "-y", "y", "--dtypes", "일자:datetime")["version_dir"]
    meta = json.load(open(os.path.join(v, "model_y.json"), encoding="utf-8"))
    assert "일자" in meta["feature_schema"]["datetime"] and meta["datetime_stats"]["일자"]["min"].startswith("2024-01-01")
    res = run_tool("ml_predict.py", "-m", v, "-f", str(p), "--save", str(tmp_path / "pred.csv"))
    assert res["missing_input_warning"] is False                    # before 2.9.11 every date was imputed
    assert not any("날짜" in w for w in res["extrapolation_warnings"])


# ---- 원본 행 번호 --------------------------------------------------------------------------------------------------


def test_loader_keeps_source_row_numbers(tmp_path):
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "공정 기록"
    ws.append([])
    ws.append(["a", "b"])           # header on row 3
    ws.append([1, 2])               # row 4
    ws.append([None, None])         # blank row 5 is dropped
    ws.append([3, 4])               # row 6
    path = tmp_path / "s.xlsx"
    wb.save(path)
    df = mp.load_data_file(str(path))
    assert list(df["a"]) == [1, 3] and df.attrs["source_rows"] == [4, 6]


def test_source_rows_not_copied_per_operation(tmp_path):
    path = write_csv_header(tmp_path / "big.csv", ["a", "b"], [[str(i), str(i % 7)] for i in range(2000)], ",")
    df = mp.load_data_file(path)
    rows = df.attrs["source_rows"]
    assert rows == list(range(2, 2002))
    assert df["a"].attrs["source_rows"] is rows                           # 2.9.19: a deep copy per Series
    assert copy.deepcopy(df.attrs)["source_rows"] is rows
    assert mp.coerce_numeric(df["a"]).iloc[-1] == 1999.0
