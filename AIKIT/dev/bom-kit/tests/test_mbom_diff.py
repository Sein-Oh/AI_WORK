# -*- coding: utf-8 -*-
"""mbom-diff 기본 스킬의 명령 하나(bom.py diff)로 make_mbom_samples의 정답(MOD_CHANGES)과 맞춰 본다.
LLM은 이 명령 하나만 실행한다 (읽기·경로·비교·요약은 bom-kit이 pandas-kit 명령을 차례로 실행해서 함)."""
import csv
import io
import os

import openpyxl
import pytest

from bk_helpers import AIKIT, bom, command_args, make_samples

import make_mbom_samples as samples  # noqa: E402  (bk_helpers가 경로를 더함)

SKILL = os.path.join(AIKIT, ".opencode", "skills", "mbom-diff")


@pytest.fixture
def files(env):
    return make_samples(env["dir"] / "mbom")


def compare(old, new, *more):
    """(결과 JSON, 변경내용.csv의 행들)."""
    code, out = bom("diff", "--old", old, "--new", new, *more)
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    with open(out["changes_file"], encoding="utf-8-sig") as f:
        rows = list(csv.DictReader(io.StringIO(f.read())))
    return out, rows


def norm(v):
    if v is None:
        return None
    try:
        return round(float(v), 6)
    except (TypeError, ValueError):
        return str(v)


def test_v2_mod_changes_are_found_exactly_and_traps_are_not_changes(env, files):
    out, rows = compare(files / "MBOM_DR200_v2.xlsx", files / "MBOM_DR200_v2_mod.xlsx")
    got = sorted((r["구분"], r["경로"], r["열"] or None, norm(r["옛 값"] or None), norm(r["새 값"] or None)) for r in rows)
    want = sorted((k, p, c, norm(a), norm(b)) for k, p, c, a, b in samples.MOD_CHANGES)
    assert got == want            # 행 밀림·순서만 바뀜·'18' -> 18·제목 블록은 나오지 않음
    for r in rows:                # 원본 엑셀 행을 잃지 않음 (join이 엑셀 행을 그대로 둠)
        assert bool(r["oldp 행"]) == (r["구분"] != "추가") and bool(r["newp 행"]) == (r["구분"] != "삭제"), r
    purposes = [c["purpose"] for c in out["confirm_commands"]]    # 두 원본에 칠하는 office-cli 명령, 행과 칸을 나눠 셈
    assert len(purposes) == 2 and "삭제된 3행, 바뀐 칸 8곳" in purposes[0] and "추가된 3행, 바뀐 칸 8곳" in purposes[1], purposes
    assert "office-cli" in out["confirm_commands"][0]["command"] or "excel.py" in out["confirm_commands"][0]["command"]
    assert out["counts"] == {"대체 후보": 1, "추가": 2, "삭제": 2, "바뀐 칸": 8}
    text = out["preview"]
    assert text.startswith("변경 대체 후보 1건, 추가 2건, 삭제 2건, 바뀐 칸 8곳 (7개 품목).") and "[이동]" not in text
    for line in ("[삭제] 110040: DR-200/100100 아래 (옛 엑셀 행 17)", "[추가] 110040: DR-200 아래 (새 엑셀 행 109)",
                 "[대체 후보] 600160 -> 600161: DR-200/100600 아래 같은 자리",
                 "[추가] 600180: DR-200/100600 아래", "[삭제] 900020: DR-200 아래",
                 "[바뀜] 300213 (DR-200/100300/300210#1 아래): Qty 3 -> 4",
                 "[바뀜] RM-1001 (DR-200/100300/300210#2/300212 아래): Qty 0.045 -> 0.05",
                 "[바뀜] 000123 (DR-200/100500 아래): Qty 2 -> 3",
                 "읽은 표: 옛 판 MBOM_DR200_v2.xlsx MBOM B6:R115"):
        assert line in text, (line, text)
    assert "대체 후보" in out["agent_directive"] and "confirm_commands" in out["agent_directive"]
    # 중간 경고는 사실만: pandas-kit 명령·옵션을 권하는 문장은 빠짐 (경량 모델이 여러 단계를 다시 잇지 않게)
    assert any(w.startswith("옛 판: 숨긴 행") for w in out["warnings"])
    assert not any("python " in w or "--" in w for w in out["warnings"]), out["warnings"]


def test_v1_to_v2_and_same_file(env, files):
    out, _ = compare(files / "MBOM_DR200.xlsx", files / "MBOM_DR200_v2.xlsx")
    assert out["counts"] == {"대체 후보": 0, "추가": 0, "삭제": 1, "바뀐 칸": 2}
    assert "[삭제] 110050: DR-200/100100 아래" in out["preview"]
    out, rows = compare(files / "MBOM_DR200_v2.xlsx", files / "MBOM_DR200_v2.xlsx")
    assert rows == [] and out["count"] == 0 and out["preview"].startswith("두 판이 같습니다 (찾은 변경 없음).")
    assert "0건" in out["agent_directive"]


def test_questions_come_back_as_full_bom_commands(env, files):
    """시트·표 범위를 pandas-kit이 물으면 그 값을 넣은 bom.py 전체 명령으로 (pandas-kit 명령을 LLM에게 주지 않음)."""
    old, new = files / "MBOM_DR200_v2.xlsx", files / "MBOM_DR200_v2_mod.xlsx"
    code, out = bom("diff", "--old", old, "--new", new, "--sheet", "BOM")
    assert code == 1 and out["error_type"] == "needs_user_decision" and out["message"].startswith("옛 판을(를) 읽다가 멈췄습니다")
    assert "[후보] 1) MBOM 2) Routing 3) 변경이력" in out["message"]
    fixed = command_args(out["confirm_commands"][0]["command"])
    assert fixed[0] == "diff" and fixed[fixed.index("--sheet") + 1] == "MBOM" and "--old-sheet" not in fixed
    code, out = bom(*fixed)                                         # 고른 명령으로 끝까지 됨
    assert code == 0 and out["counts"]["바뀐 칸"] == 8, out
    wb = openpyxl.load_workbook(new)                                # 표 가운데 빈 행 둘: 어디까지가 표인지 묻기
    for r in (40, 80):
        for c in wb["MBOM"][r]:
            c.value = None
    split = files / "split.xlsx"
    wb.save(split)
    code, out = bom("diff", "--old", old, "--new", split)
    assert code == 1 and out["message"].startswith("새 판을(를) 읽다가 멈췄습니다") and "[후보] 1) 빈 행까지 한 표로 (B6:R115" in out["message"]
    first = command_args(out["confirm_commands"][0]["command"])
    assert first[0] == "diff" and first[first.index("--new-range") + 1] == "B6:R115"
    code, out = bom(*first)
    assert code == 0 and out["counts"]["삭제"] >= 2, out


def test_missing_column_is_asked_with_full_commands(env, files):
    code, out = bom("diff", "--old", files / "MBOM_DR200.xlsx", "--new", files / "MBOM_DR200_v2.xlsx", "--level", "Lv")
    assert code == 1 and out["message"].startswith("'Lv' 열이 없습니다. Level 열을(를) 정해야 합니다. [후보] 1) Level")
    assert command_args(out["confirm_commands"][0]["command"])[-2:] == ["--level", "Level"]
    code, out = bom("diff", "--bogus")
    assert code == 2 and out["status"] == "error"


def test_paths_number_repeated_parts_and_warn_level_jumps():
    import bom as bk                                               # BOM 업무 규칙 함수 (bom-kit 폴더, bk_helpers가 경로를 더함)
    rows = [dict(zip(("엑셀 행", "Level", "Part Number"), r.split(","))) for r in
            ("2,0,TOP", "3,1,S", "4,2,P", "5,1,S", "6,2,P", "7,3,X", "8,1,Q", "9,3,Z", "10,x,BAD")]
    t = bk.tree(rows, "Level", "Part Number")
    assert [(i["r"]["엑셀 행"], i["parent"], i["path"]) for i in t["items"]] == [
        ("2", "", "TOP"), ("3", "TOP", "TOP/S#1"), ("4", "S", "TOP/S#1/P"), ("5", "TOP", "TOP/S#2"), ("6", "S", "TOP/S#2/P"),
        ("7", "P", "TOP/S#2/P/X"), ("8", "TOP", "TOP/Q"), ("9", "Q", "TOP/Q/Z")]
    warns = bk.tree_warnings(t, "옛 판")
    assert any("건너뛴 행 1개 (엑셀 행 9)" in w for w in warns) and any("품번이 빈 행 1개" in w and "엑셀 행 10" in w for w in warns)
