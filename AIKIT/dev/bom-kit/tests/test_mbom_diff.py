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
    assert out["counts"] == {"상위 삭제": 0, "대체 후보": 1, "추가": 2, "삭제": 2, "바뀐 칸": 8}
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
    assert out["counts"] == {"상위 삭제": 0, "대체 후보": 0, "추가": 0, "삭제": 1, "바뀐 칸": 2}
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


def _bom_xlsx(path, rows, blank_after=()):
    """머리글 1행 B열부터, blank_after의 행 번호(데이터 몇 번째) 뒤에 빈 행."""
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "MBOM"
    ws.append([None, "Level", "Part Number", "Qty"])
    for i, r in enumerate(rows, 1):
        ws.append([None, *r])
        if i in blank_after:
            ws.append([])
    wb.save(path)
    return path


def test_a_removed_parent_is_one_line_and_its_children_are_not_painted(env, tmp_path):
    """상위만 지워져 하위가 다른 상위 아래에 그대로 붙은 것은 한 줄 (집 PC 2026-11-14 sample1·2: 상위 2개 삭제가
    '추가 38, 삭제 39'로 나왔음, 사용자 결정). 하위 행은 세지도 칠하지도 않고, 그 안의 진짜 삭제는 따로."""
    old = _bom_xlsx(tmp_path / "o.xlsx", [(0, "TOP", 1), (1, "A", 1), (2, "A1", 2), (1, "G", 1), (2, "G1", 1), (3, "R", 2),
                                          (2, "G2", 3), (2, "G3", 1)])
    new = _bom_xlsx(tmp_path / "n.xlsx", [(0, "TOP", 1), (1, "A", 1), (2, "A1", 2), (2, "G1", 1), (3, "R", 2), (2, "G2", 3)])
    code, d = bom("diff", "--old", old, "--new", new)
    assert code == 0 and d["verified"], d
    assert d["counts"] == {"상위 삭제": 1, "대체 후보": 0, "추가": 0, "삭제": 1, "바뀐 칸": 0}
    assert d["lines"] == ["[상위 삭제] G: TOP 아래 (옛 엑셀 행 5). 그 하위 3행은 새 판에서 TOP/A 아래에 그대로 있음 "
                          "(옛 엑셀 행 6~8, 새 엑셀 행 5~7)",
                          "[삭제] G3: TOP/G 아래 (옛 엑셀 행 9)"]
    from bk_helpers import command_args
    paint = {c["purpose"].split("의", 1)[0]: command_args(c["command"]) for c in d["confirm_commands"]}
    with open(paint["원본 o.xlsx"][paint["원본 o.xlsx"].index("--file") + 1], encoding="utf-8-sig") as f:
        old_cells = f.read().split()
    assert old_cells == ["B5:D5", "B9:D9"] and not any(k.startswith("원본 n.xlsx") for k in paint)   # 새 판은 칠할 것 없음


def test_both_versions_with_blank_rows_ask_the_range_once(env, tmp_path):
    """두 판 모두 표 안에 빈 행이 있으면 범위를 한 번에 묻고, 1번은 두 판 모두 한 표로 (집 PC 2026-11-14: 두 번 물었음)."""
    from bk_helpers import command_args
    rows = [(0, "TOP", 1), (1, "A", 1), (2, "A1", 2), (1, "B", 1), (2, "B1", 2), (2, "B2", 1), (2, "B3", 1)]
    old = _bom_xlsx(tmp_path / "o.xlsx", rows, blank_after=(3,))
    new = _bom_xlsx(tmp_path / "n.xlsx", rows[:4] + [(2, "B1", 5)] + rows[5:], blank_after=(3,))
    code, d = bom("diff", "--old", old, "--new", new)
    assert code == 1 and d["error_type"] == "needs_user_decision"
    assert d["message"].startswith("옛 판과 새 판을 읽다가 멈췄습니다: 두 판 모두 빈 행으로 나뉜 덩어리가 있어")
    assert "[후보] 1) 두 판 모두 빈 행까지 한 표로 (옛 판 B1:D9" in d["message"] and "2) 옛 판 B1:D4만" in d["message"]
    first = command_args(d["confirm_commands"][0]["command"])
    assert first[-4:] == ["--old-range", "B1:D9", "--new-range", "B1:D9"]
    code, d = bom(*first)
    assert code == 0 and d["lines"] == ["[바뀜] B1 (TOP/B 아래): Qty 2 -> 5 (옛 엑셀 행 7, 새 엑셀 행 7)"]
    plain = _bom_xlsx(tmp_path / "p.xlsx", rows)                                 # 새 판은 빈 행 없음: 옛 판만 물음
    code, d = bom("diff", "--old", old, "--new", plain)
    assert code == 1 and d["message"].startswith("옛 판을(를) 읽다가 멈췄습니다") and "두 판 모두" not in d["message"]
