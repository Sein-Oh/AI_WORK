# -*- coding: utf-8 -*-
"""
2.9.16 (5차 점검 수정): a first data row of text values (categories, a text target) is not a unit row (header
suggestion, header + unit pair, --unit-row check), a unit row in a new file for a model trained without units
(--data-row suggestion), and the header-row hint of ml_map --points.
"""
import os
import shlex
import sys

import numpy as np
import openpyxl
import pandas as pd
import pytest

HERE = os.path.dirname(os.path.abspath(__file__))
KIT = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(HERE))), "CLI", "ml-kit")
sys.path.insert(0, HERE)
sys.path.insert(0, KIT)
from test_cli_smoke import run_tool, run_fail  # noqa: E402
import ml_predict as mp  # noqa: E402


def _book(path, rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    for row in rows:
        ws.append(row)
    wb.save(path)
    return str(path)


def _recommendation(message):
    return message.split("[추천] ")[1].split("\n")[0].strip()


def _rerun(message):
    rec = _recommendation(message)
    assert rec.startswith("python"), rec
    parts = shlex.split(rec)
    return os.path.basename(parts[1]), parts[2:]


# 1. a text first data row is not a unit row -----------------------------------------------------------------

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
    return _book(tmp_path_factory.mktemp("two") / "two.xlsx", _two_level_rows())


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


def test_text_first_row_header_hint_has_no_unit_row(two_level):
    bad = run_fail("ml_train.py", "-f", two_level, "-x", "라인,설비,작업자,온도", "-y", "등급", "--task", "classification")
    assert _recommendation(bad["message"]) == "--header-row 2"     # 2.9.15: --header-row 2 --unit-row 3
    assert "단위 행" not in bad["message"]
    res = run_tool("ml_train.py", "-f", two_level, "-x", "라인,설비,작업자,온도", "-y", "등급", "--task", "classification",
                   "--header-row", "2")
    assert res["targets"] == ["등급"]


def test_text_first_row_given_as_unit_row_is_refused(two_level):
    bad = run_fail("ml_train.py", "-f", two_level, "-x", "라인,설비,작업자,온도", "-y", "등급", "--task", "classification",
                   "--header-row", "2", "--unit-row", "3")             # 2.9.15: VERIFIED, '온도(182.8)', '등급(중)'
    assert "단위 행(--unit-row 3)" in bad["message"] and bad["agent_directive"].startswith("STOP_RETRY")
    script, args = _rerun(bad["message"])
    assert "--unit-row" not in args and args[-2:] == ["--header-row", "2"]
    res = run_tool(script, *args)
    assert res["targets"] == ["등급"] and sorted(res["feature_types"]["categorical"]) == ["라인", "설비", "작업자"]


def test_header_unit_pair_skips_text_data_rows(tmp_path):
    # a first data row with blank measurements, then a full one: categories over categories, not units
    rows = [["라인", "설비", "온도", "압력", "유량"], ["A", "M1", None, None, None], ["B", "M2", 181.0, 3.0, 5.0]]
    raw = pd.read_excel(_book(tmp_path / "p.xlsx", rows), header=None)
    assert mp._header_unit_pair(raw) is None                       # 2.9.15: (1, 2), a data row as the unit row
    rows = [["제목"], [None], ["반응온도", "압력", "제외"], ["°C", "bar", None], [192.5, 3.0, None]]
    raw = pd.read_excel(_book(tmp_path / "b.xlsx", rows), header=None)
    assert mp._header_unit_pair(raw) == (3, 4)                     # beta sample layout


# 2. unit row in a new file for a model trained without units ------------------------------------------------

@pytest.fixture(scope="module")
def plain_model(tmp_path_factory):
    work = tmp_path_factory.mktemp("plain")
    r = np.random.default_rng(12)
    rows = [["반응온도", "압력", "효율"]]
    for _ in range(40):
        t, p = float(r.uniform(150, 200)), float(r.uniform(1, 5))
        rows.append([round(t, 2), round(p, 3), round(0.3 * t + 2 * p + float(r.normal(0, 0.2)), 3)])
    path = _book(work / "plain.xlsx", rows)
    return run_tool("ml_train.py", "-f", path, "-x", "반응온도,압력", "-y", "효율")["version_dir"]


def test_unit_row_file_for_a_unitless_model(plain_model, tmp_path):
    p = _book(tmp_path / "u2.xlsx", [["반응온도", "압력"], ["°C", "bar"], [180, 3], [185, 3.2]])
    bad = run_fail("ml_predict.py", "-m", plain_model, "-f", p)      # 2.9.15: VERIFIED, unit row imputed
    script, args = _rerun(bad["message"])
    assert args[-2:] == ["--data-row", "3"] and "--unit-row" not in args
    ok = run_tool(script, *args)
    assert ok["sample_count"] == 2 and not ok["missing_input_warning"]


def test_title_and_unit_row_file_for_a_unitless_model(plain_model, tmp_path):
    p = _book(tmp_path / "t2c.xlsx", [["새 조건", "작성 2026-11-02"], ["반응온도", "압력"], ["°C", "bar"],
                                      [180, 3], [185, 3.2]])
    bad = run_fail("ml_predict.py", "-m", plain_model, "-f", p)
    script, args = _rerun(bad["message"])                           # 2.9.15: --unit-row 3 -> '반응온도(°C)' missing
    assert args[-4:] == ["--header-row", "2", "--data-row", "4"]
    ok = run_tool(script, *args)
    assert ok["sample_count"] == 2 and not ok["missing_input_warning"]


# 3. ml_map --points: header row hint --------------------------------------------------------------------------

def test_map_points_header_hint(tmp_path):
    rows = [["맵 기록", "2026-11-02"], ["A", "B"]] + [[float(i), float(i * i)] for i in range(10)]
    data = _book(tmp_path / "m.xlsx", rows)
    pts = tmp_path / "pts.csv"
    pd.DataFrame({"A": [2.5, 3.5]}).to_csv(pts, index=False, encoding="utf-8-sig")
    bad = run_fail("ml_map.py", "-f", data, "--points", str(pts), "-y", "B")
    assert _recommendation(bad["message"]) == "--header-row 2"     # 2.9.15: no hint
    ok = run_tool("ml_map.py", "-f", data, "--points", str(pts), "-y", "B", "--header-row", "2")
    assert ok["count"] == 2
