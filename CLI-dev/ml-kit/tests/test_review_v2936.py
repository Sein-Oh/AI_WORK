# -*- coding: utf-8 -*-
"""
2.9.36 (25차 확인 검사) regression.

2.9.35 compared a --points column and a near data column without the -y name, but the leftovers kept their
punctuation: '_1' / '(1)', '_상' / '(상)', '1' / '(1단)' were no longer near, so a key column for another sensor
position or stage ('온도_1' for data '온도(1)', -y '온도(3)') was dropped silently and the result came out VERIFIED.
The leftovers are now compared without punctuation, and one holding the other counts as near.
"""
import os
import random
import sys

import pytest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from test_layout_matrix_v2917 import _write  # noqa: E402
from test_cli_smoke import run_raw  # noqa: E402


@pytest.fixture(scope="module")
def tmp(tmp_path_factory):
    return str(tmp_path_factory.mktemp("v2936"))


def _data(tmp, name, cols):
    r = random.Random(3)
    rows = [cols]
    for _ in range(30):
        a, b, c = (round(r.uniform(100, 200), 1) for _ in range(3))
        rows.append([round(r.uniform(10, 30), 1), a, b, c])
    return _write(os.path.join(tmp, name), rows, [])


@pytest.mark.parametrize("pcol,dcols,target", [
    ("온도_1", ["시간", "온도(1)", "온도(2)", "온도(3)"], "온도(3)"),
    ("온도#1", ["시간", "온도(1)", "온도(2)", "온도(3)"], "온도(3)"),
    ("유량-1", ["시간", "유량(1)", "유량(2)", "유량(3)"], "유량(3)"),
    ("온도_상", ["시간", "온도(상)", "온도(중)", "온도(하)"], "온도(하)"),
    ("압력1", ["시간", "압력(1단)", "압력(3단)", "압력(2단)"], "압력(2단)"),
    ("수율_1", ["시간", "수율(1차)", "수율(2차)", "수율(최종)"], "수율(최종)"),
])
def test_points_position_key_column_stops(tmp, pcol, dcols, target):
    d = _data(tmp, f"d_{len(pcol)}_{pcol[-1]}.csv", dcols)
    p = _write(os.path.join(tmp, f"p_{len(pcol)}_{pcol[-1]}.csv"), [["시간", pcol], [20, 150], [20, 170]], [])
    proc, res = run_raw("ml_map.py", "-f", d, "--points", p, "-y", target,
                        "--save", os.path.join(tmp, "o.xlsx"), "--overwrite")
    assert proc.returncode == 1 and res["agent_directive"].startswith("STOP_RETRY:"), res
    assert "[후보]" in res["message"] and "[추천]" not in res["message"], res


@pytest.mark.parametrize("pcol", ["수율_실측", "목표수율"])
def test_points_measured_column_still_passes(tmp, pcol):
    d = _data(tmp, "dm.csv", ["온도", "함수율", "수율", "생산량"])
    p = _write(os.path.join(tmp, f"pm_{len(pcol)}.csv"), [["온도", pcol], [15, 30], [25, 40]], [])
    proc, res = run_raw("ml_map.py", "-f", d, "--points", p, "-y", "수율",
                        "--save", os.path.join(tmp, "om.xlsx"), "--overwrite")
    assert proc.returncode == 0 and res["agent_directive"].startswith("VERIFIED:"), res
