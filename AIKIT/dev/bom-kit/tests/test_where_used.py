# -*- coding: utf-8 -*-
"""where-used 기본 스킬의 명령 하나(bom.py where-used)로 make_mbom_samples의 MBOM에서 품목이 쓰인 상위와
제품 1대당 수량이 맞는지 본다."""
import os

import pytest

from bk_helpers import AIKIT, bom, command_args, make_samples

SKILL = os.path.join(AIKIT, ".opencode", "skills", "where-used")


def script(*args):
    return bom("where-used", *args)


@pytest.fixture
def bom_csv(env):
    """이름은 그대로 두지만 이제 BOM 엑셀 파일 (도구가 pandas-kit으로 읽음)."""
    return make_samples(env["dir"] / "mbom") / "MBOM_DR200.xlsx"


def test_where_used_finds_every_place_and_the_quantity_per_product(bom_csv):
    code, d = script("--file", bom_csv, "--find", "000125")             # 볼트 M3x12: 2레벨 6곳 + 3레벨 4곳
    assert code == 0 and d["agent_directive"].startswith("VERIFIED:") and d["count"] == 10
    assert d["totals"] == ["000125 볼트 M3x12 (아연): 10곳, 제품 1대당 31 EA"]
    assert "[쓰임] 000125: 100100 하우징 조립 아래 4 EA (경로 DR-200/100100/000125, 엑셀 행 16)" in d["lines"]
    assert "[쓰임] 000125: 300210 유성기어 스테이지 아래 3 EA (경로 DR-200/100300/300210#2/000125, 엑셀 행 66)" in d["lines"]
    assert os.path.isfile(d["output_file"]) and d["output_file"].endswith("_쓰인곳.csv")


def test_quantity_is_multiplied_along_the_path(bom_csv):
    # 봉강은 유성기어 Z12(3개) 아래에 0.012 KG씩: 경로를 따라 곱해야 0.036 KG
    code, d = script("--file", bom_csv, "--find", "RM-1001")
    assert code == 0 and d["count"] == 11
    assert d["totals"] == ["RM-1001 봉강 SCM440 D12: 11곳, 제품 1대당 0.56 KG"]
    assert "[쓰임] RM-1001: 300211 유성기어 Z12 아래 0.036 KG (경로 DR-200/100300/300210#1/300211/RM-1001, 엑셀 행 55)" in d["lines"]


def test_a_word_in_the_name_finds_every_part_and_counts_them_apart(bom_csv):
    code, d = script("--file", bom_csv, "--find", "나사")
    assert code == 0 and d["count"] == 9
    assert d["totals"] == ["110099 하우징 나사 커버 (폐기): 1곳, 제품 1대당 2 EA",      # 숨긴 행도 BOM에 있으면 셈
                           "000123 나사 M3x8 (아연): 4곳, 제품 1대당 16 EA",
                           "000124 나사 M4x12 (아연): 2곳, 제품 1대당 6 EA",
                           "500130 척 고정 나사 (왼나사): 1곳, 제품 1대당 1 EA",
                           "900020 나사 고정제: 1곳, 제품 1대당 0.5 ML"]              # 단위가 다르면 따로


def test_nothing_found_and_a_missing_column_stop_cleanly(bom_csv):
    code, d = script("--file", bom_csv, "--find", "없는품번")
    assert code == 0 and d["count"] == 0 and "찾지 못했습니다" in d["preview"] and "0건" in d["agent_directive"]
    code, d = script("--file", bom_csv, "--find", "000125", "--part", "없는열")
    assert code == 1 and d["error_type"] == "needs_user_decision" and "'없는열' 열이 없습니다" in d["message"]
    pick = [c["command"] for c in d["confirm_commands"] if "Part Number" in c["command"]][0]
    code, d = bom(*command_args(pick))                                   # 고른 열로 다시
    assert code == 0 and d["count"] == 10
    code, d = script("--file", bom_csv.parent / "없음.xlsx", "--find", "볼트")
    assert code == 1 and d["message"].startswith("BOM을(를) 읽다가 멈췄습니다") and "python" not in d["message"]


def test_skill_md_is_one_command_and_short():
    with open(os.path.join(SKILL, "SKILL.md"), "rb") as f:
        raw = f.read().replace(b"\r\n", b"\n")
    assert len(raw) <= 8192
    text = raw.decode("utf-8")
    assert "name: where-used" in text and 'version: "0.2.0"' in text
    assert text.count(r'python "C:\AIKIT\bom-kit\bom.py" where-used --file') == 1 and "pandas_kit.py" not in text
    assert "mbom-diff" in text                                           # 두 판 비교는 그쪽이라고 알려 줌
