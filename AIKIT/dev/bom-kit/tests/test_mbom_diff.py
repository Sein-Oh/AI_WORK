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
    wb = openpyxl.load_workbook(new)                                # 표 가운데 빈 행 둘: 묻지 않고 한 표로 (0.2.1)
    for r in (40, 80):
        for c in wb["MBOM"][r]:
            c.value = None
    split = files / "split.xlsx"
    wb.save(split)
    code, out = bom("diff", "--old", old, "--new", split)
    assert code == 0 and out["counts"]["삭제"] >= 2, out
    assert "새 판: 표 안에 빈 행이 있어 빈 행까지 한 표로 읽었습니다 (B6:R115, 빈 행 빼고 107행)." in out["warnings"]


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


def test_blank_rows_inside_the_table_are_read_as_one_table_without_asking(env, tmp_path):
    """표 안에 빈 행이 있으면 묻지 않고 빈 행까지 한 표로 읽고 알림 (사용자 결정 2026-11-14: 전에는 옛 판·새 판마다 [후보]로 물었음)."""
    rows = [(0, "TOP", 1), (1, "A", 1), (2, "A1", 2), (1, "B", 1), (2, "B1", 2), (2, "B2", 1), (2, "B3", 1)]
    old = _bom_xlsx(tmp_path / "o.xlsx", rows, blank_after=(3,))
    new = _bom_xlsx(tmp_path / "n.xlsx", rows[:4] + [(2, "B1", 5)] + rows[5:], blank_after=(3,))
    code, d = bom("diff", "--old", old, "--new", new)
    assert code == 0 and d["verified"], d
    assert d["lines"] == ["[바뀜] B1 (TOP/B 아래): Qty 2 -> 5 (옛 엑셀 행 7, 새 엑셀 행 7)"]
    for who in ("옛 판", "새 판"):
        assert f"{who}: 표 안에 빈 행이 있어 빈 행까지 한 표로 읽었습니다 (B1:D9, 빈 행 빼고 7행)." in d["warnings"]
    assert not any("범위 안의 빈 행" in w for w in d["warnings"])                       # 같은 알림을 두 번 하지 않음
    code, w = bom("where-used", "--file", old, "--find", "B3")                          # where-used도 같음
    assert code == 0 and w["count"] == 1 and any("빈 행까지 한 표로" in x for x in w["warnings"])


def test_the_comparison_is_saved_as_an_excel_file(env, tmp_path):
    """비교 결과 엑셀 파일 (사용자 결정 2026-11-14): 요약 시트(preview와 같은 줄, 알림)와 변경내용 시트(한 줄에 한 변경).
    품번은 숫자처럼 보여도 글자, 엑셀 행 번호는 숫자. bom-kit은 openpyxl 없이 씀 (시험에서만 openpyxl로 확인)."""
    rows = [(0, "TOP", 1), (1, "100300", 1), (2, "000125", 4), (1, "B", 1), (2, "B1", 2)]
    old = _bom_xlsx(tmp_path / "o.xlsx", rows)
    new = _bom_xlsx(tmp_path / "n.xlsx", [rows[0], rows[1], (2, "000125", 6), rows[3]])
    code, d = bom("diff", "--old", old, "--new", new)
    assert code == 0 and d["verified"] and d["preview"].splitlines()[-1] == f"비교 결과 파일: {d['output_file']}"
    assert os.path.dirname(d["output_file"]) == str(tmp_path)                                   # 대상 파일 옆
    assert os.path.basename(d["output_file"]).startswith("o_n_비교결과_") and d["output_file"].endswith(".xlsx")
    wb = openpyxl.load_workbook(d["output_file"])
    assert wb.sheetnames == ["요약", "변경내용"]
    summary = [r[0] for r in wb["요약"].iter_rows(min_row=2, values_only=True)]
    assert summary[0] == d["preview"].splitlines()[0] and summary[1:1 + len(d["lines"])] == d["lines"]
    table = list(wb["변경내용"].iter_rows(values_only=True))
    assert table[0] == ("구분", "품번", "위치", "열", "옛 값", "새 값", "옛 엑셀 행", "새 엑셀 행", "비고")
    assert ("바뀜", "000125", "TOP/100300", "Qty", "4", "6", 4, 4, None) in table
    assert ("삭제", "B1", "TOP/B", None, None, None, 6, None, None) in table
