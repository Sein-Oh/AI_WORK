# -*- coding: utf-8 -*-
"""pandas-kit 첫 판: select, show, join, diff, only, stack, undo, new, export와 공통 계약(1줄 JSON, VERIFIED, 0건, truncated,
recovered, [추천]/[후보], 오류 로그). 도구는 서브프로세스로 실행한다 (stdout 1줄 확인)."""
import json
import os
import pickle
import re
import shlex
import sys
import time

import openpyxl
import pandas as pd
import pytest

from pk_helpers import KIT, command_args, make_bom, ok, recommended, run

sys.path.insert(0, KIT)
import _pd_runtime as rt  # noqa: E402
import pandas_kit as pk  # noqa: E402

REPO_AIKIT = os.path.dirname(KIT)
SKILL = os.path.join(REPO_AIKIT, ".opencode", "skills", "pandas-kit", "SKILL.md")


def _two_boms(env):
    old = make_bom(env["dir"] / "BOM_v1.xlsx")

    def change(ws):
        ws["F7"] = "B"          # 0200 Maker A -> B
        ws["E8"] = 5            # 0200/0105 Qty 3 -> 5
        ws.append([None, 1, "TOP", "0300", 1, "A", "새것"])
        ws.delete_rows(6, 1)    # 0102/0105 지움
    new = make_bom(env["dir"] / "BOM_v2.xlsx", extra=change)
    return old, new


# ---------------------------------------------------------------------------
# 계약: 도움말, 인자 오류, 오류 로그
# ---------------------------------------------------------------------------

def test_help_is_one_json_line_and_points_to_skill(env):
    code, out = run(env, "select", "-h")
    assert code == 0 and out["command"] == "help" and out["skill_file"].endswith(os.path.join("pandas-kit", "SKILL.md"))
    assert out["agent_directive"].startswith("VERIFIED:")


def test_same_meaning_option_gives_full_command_and_errors_are_logged(env):
    path = make_bom(env["dir"] / "b.xlsx")
    ok(env, "select", "--file", path)
    code, out = run(env, "show", "--table", "df1")
    assert code == 2 and out["agent_directive"].startswith("STOP_RETRY:")
    assert recommended(out["message"]) == ["show", "--from", "df1"]
    code, out = run(env, "select", "--from", "df1", "--where", "Makr=A")
    assert code == 1 and out["error_type"] == "needs_user_decision" and "[후보] 1) Maker" in out["message"]
    assert command_args(out["confirm_commands"][0]["command"]) == ["select", "--from", "df1", "--where", "Maker=A"]
    log = (env["data"] / "logs" / "errors.txt").read_text(encoding="utf-8")
    assert log.count("=====") == 2 and "pandas-kit " + rt.VERSION in log
    assert "볼트" not in log                     # 칸 값은 남기지 않음


def test_version_matches_skill_and_skill_size():
    with open(SKILL, "rb") as f:
        raw = f.read().replace(b"\r\n", b"\n")
    assert len(raw) <= 8192
    text = raw.decode("utf-8")
    assert f'version: "{rt.VERSION}"' in text and "name: pandas-kit" in text


def test_skill_options_exist_in_the_tool():
    """스킬에 적힌 명령과 옵션이 도구에 있다 (지어낸 옵션을 적지 않음)."""
    import re
    parser = pk.build_cli()
    sub = next(a for a in parser._actions if a.__class__.__name__ == "_SubParsersAction")
    text = open(SKILL, encoding="utf-8").read()
    found = 0
    for line in text.splitlines():
        m = re.search(r'pandas_kit\.py"\s+(\w[\w-]*)(.*)', line)
        if not m:
            continue
        cmd, rest = m.group(1), m.group(2)
        assert cmd in sub.choices, (cmd, line)
        opts = {o for a in sub.choices[cmd]._actions for o in a.option_strings}
        for o in re.findall(r"(--[a-z][a-z-]*)", rest.split("`")[0]):
            assert o in opts, (cmd, o, line)
        found += 1
    assert found >= 10


def test_tool_imports_only_its_runtime():
    src = open(os.path.join(KIT, "pandas_kit.py"), encoding="utf-8").read()
    assert "import _pd_runtime" in src and "office-cli" not in src.split("OFFICE_EXCEL")[0].split("import _pd_runtime")[1][:2000]
    assert not any(n in src for n in ("import excel", "import _runtime", "import _ml_runtime"))


# ---------------------------------------------------------------------------
# select: 표 찾기, 엑셀 위치, 값은 바꾸지 않음
# ---------------------------------------------------------------------------

def test_select_file_finds_table_under_title_and_keeps_codes(env):
    path = make_bom(env["dir"] / "BOM.xlsx")
    out = ok(env, "select", "--file", path, "--name", "df1")
    assert out["agent_directive"].startswith("VERIFIED:") and out["count"] == 4 and out["table"] == "df1"
    assert out["anchor"]["range"] == "B4:G8" and out["anchor"]["header_row"] == 4 and "B4:G8 (머리글 4행)" in out["preview"]
    cols = {c["name"]: c for c in out["columns"]}
    assert cols["Part Number"]["letter"] == "D" and cols["Part Number"]["kind"] == "글자" and cols["Qty"]["kind"] == "숫자"
    work = out["work"]
    assert re.fullmatch(r"BOM_\d{10}", os.path.basename(work)) and os.path.dirname(work) == str(env["dir"])   # <파일 이름>_YYMMDDHHMM
    shown = ok(env, "show")
    assert shown["rows"][0] == {"엑셀 행": 5, "Level": 1, "Parent": "TOP", "Part Number": "0102", "Qty": 2, "Maker": "A", "Desc": "볼트"}
    assert shown["truncated"] is False and shown["count"] == 4


def test_work_folder_made_again_in_the_same_minute_gets_2(env):
    # 사용자 결정 2026-11-13: <파일 이름>_YYMMDDHHMM, 같은 분에 또 만들면 _2, _3
    from datetime import datetime, timedelta
    path = make_bom(env["dir"] / "BOM.xlsx")
    now = datetime.now()
    taken = {f"BOM_{t:%y%m%d%H%M}" for t in (now, now + timedelta(minutes=1))}   # 분이 바뀌는 순간에도 같은 시험
    for name in taken:
        os.makedirs(env["dir"] / name)
    work = os.path.basename(ok(env, "select", "--file", path)["work"])
    assert work.endswith("_2") and work[:-2] in taken


def test_select_csv_cp949_leading_zero_and_mixed_columns(env):
    p = env["dir"] / "list.csv"
    p.write_bytes("품번,수량,비고\n0102,3,a\n0105,10,\n0200,2,x\n".encode("cp949"))
    out = ok(env, "select", "--file", str(p))
    kinds = {c["name"]: c["kind"] for c in out["columns"]}
    assert kinds == {"품번": "글자", "수량": "숫자", "비고": "글자"}
    rows = ok(env, "show")["rows"]
    assert [r["품번"] for r in rows] == ["0102", "0105", "0200"] and rows[0]["엑셀 행"] == 2 and rows[1]["비고"] is None


def test_where_and_or_contains_and_letters(env):
    path = make_bom(env["dir"] / "BOM.xlsx")
    ok(env, "select", "--file", path, "--name", "df1")
    out = ok(env, "select", "--from", "df1", "--where", "Maker=A|B;Qty>=2", "--name", "a")
    assert out["count"] == 2 and "df1에서" in out["lineage"]
    assert ok(env, "select", "--from", "df1", "--where", "Desc~너", "--columns", "D,Qty")["count"] == 2
    cols = [c["name"] for c in ok(env, "show")["columns"]]
    assert cols == ["Part Number", "Qty"]
    zero = ok(env, "select", "--from", "df1", "--where", "Level=9")
    assert zero["count"] == 0 and "0건" in zero["agent_directive"] and "next_commands" not in zero
    code, out = run(env, "select", "--from", "df1", "--where", "Maker=A, Qty>1")
    assert code == 1 and recommended(out["message"])[-1] == "Maker=A;Qty>1"
    assert ok(env, *recommended(out["message"]))["count"] == 1          # Maker A이고 Qty>1: 0102(2)뿐

def test_numbers_written_as_text_and_numeric_conditions_on_text_columns(env):
    # 집 PC 2026-11-13 (MBOM): 숫자 사이 글자 "18" 한 칸이 열 전체를 글자로 만들고 'Scrap %>3'이 글자 순서로 비교돼 "18"이 빠짐
    def extra(ws):
        ws["H4"] = "Scrap %"
        ws["I4"] = "Note"
        for r, (scrap, note) in enumerate([(3, "1"), ("18", "2"), (4, "N/A"), (None, "30"), (6, "5"), (2, "7"),
                                           (5, "8"), (1, "9"), (2, "10"), (3, "11"), (1, "12")], start=5):
            ws.cell(row=r, column=8, value=scrap)
            ws.cell(row=r, column=9, value=note)
    rows = [(1, "TOP", f"{i:04d}", i, "A", "x") for i in range(1, 12)]
    out = ok(env, "select", "--file", make_bom(env["dir"] / "BOM.xlsx", rows=rows, extra=extra), "--name", "df1")
    kinds = {c["name"]: c["kind"] for c in out["columns"]}
    assert kinds["Scrap %"] == "숫자" and kinds["Note"] == "글자"
    assert any("Scrap %(H) 열의 글자로 저장된 숫자 1칸(엑셀 행 6)" in w for w in out["warnings"])
    assert any("Note(I) 열은 숫자 사이에 숫자가 아닌 칸 1개(엑셀 행 7)" in w for w in out["warnings"])
    assert not any("Part Number" in w for w in out["warnings"])          # 앞자리 0 코드 열은 경고하지 않음
    assert ok(env, "select", "--from", "df1", "--where", "Scrap %>3")["count"] == 4      # 18, 4, 6, 5
    big = ok(env, "select", "--from", "df1", "--where", "Note>=10")      # 글자 열이지만 숫자로 비교: 30, 10, 11, 12
    assert big["count"] == 4 and any("숫자가 아닌 칸 1개" in w for w in big["warnings"])


def test_not_equal_with_choices_means_none_of_them_and_empty_value_needs_equal(env):
    # 집 PC 2026-11-13: 'Item Type!=RAW|CONS'가 "RAW 아님 또는 CONS 아님"이라 모든 행이 남았고, 'Qty>='가 모든 행을 냄
    path = make_bom(env["dir"] / "BOM.xlsx")
    ok(env, "select", "--file", path, "--name", "df1")
    assert ok(env, "select", "--from", "df1", "--where", "Maker!=A|B")["count"] == 1      # C만
    assert ok(env, "select", "--from", "df1", "--where", "Desc!~볼|판")["count"] == 2     # 너트 둘
    assert ok(env, "select", "--from", "df1", "--where", "Maker=A|B")["count"] == 3
    for bad in ("Qty>=", "Qty<", "Desc~"):
        code, out = run(env, "select", "--from", "df1", "--where", bad)
        assert code == 1 and "비교할 값이 없습니다" in out["message"], bad
    assert ok(env, "select", "--from", "df1", "--where", "Desc!=")["count"] == 4


def test_explicit_range_and_unknown_sheet(env):
    path = make_bom(env["dir"] / "BOM.xlsx")
    out = ok(env, "select", "--file", path, "--range", "C4:E6")
    assert out["count"] == 2 and [c["name"] for c in out["columns"]] == ["Parent", "Part Number", "Qty"]
    code, out = run(env, "select", "--file", path, "--sheet", "BOMM")
    assert code == 1 and "[후보] 1) BOM" in out["message"]
    assert command_args(out["confirm_commands"][0]["command"])[-2:] == ["--sheet", "BOM"]


def test_hidden_rows_merged_cells_and_uncomputed_formulas_are_warned(env):
    def extra(ws):
        ws.row_dimensions[6].hidden = True
        ws.merge_cells("G7:G8")
        ws["H4"] = "합계"
        ws["H5"] = "=E5*2"
    path = make_bom(env["dir"] / "BOM.xlsx", extra=extra)
    out = ok(env, "select", "--file", path)
    text = " ".join(out["warnings"])
    assert "숨긴 행" in text and "병합한 칸 1곳" in text and "수식 칸 1개" in text


def test_empty_row_ends_table_and_rest_is_warned(env):
    def extra(ws):
        ws.append([])
        ws.append([None, 9, "X", "9999", 1, "Z", "다른 표"])
    out = ok(env, "select", "--file", make_bom(env["dir"] / "BOM.xlsx", extra=extra))
    assert out["count"] == 4 and any("빈 행 아래" in w for w in out["warnings"])


def test_title_block_above_is_not_taken_as_the_table(env):
    # 집 PC 2026-11-13 (MBOM): 제목 블록(공장·Rev·작성일, 한 행에 두 칸)을 표로 잡고 진짜 표는 "빈 행 아래"로만 알렸음
    def title(ws):
        ws["B1"], ws["D1"] = "공장", "P100"
        ws["B2"], ws["D2"] = "Rev", "C"
    import openpyxl as ox
    wb = ox.Workbook()
    ws = wb.active
    title(ws)
    ws.append([])
    ws.append([None, "Level", "Part Number", "Qty"])
    for i in range(5):
        ws.append([None, 1, f"{i:04d}", i])
    ws.append([])
    ws.append([None, "메모"])
    path = env["dir"] / "titled.xlsx"
    wb.save(path)
    out = ok(env, "select", "--file", str(path))
    assert out["anchor"]["range"] == "B4:D9" and out["count"] == 5
    assert any("표 위의 다른 내용 2행(엑셀 행 1~2)" in w for w in out["warnings"])
    assert any("빈 행 아래" in w for w in out["warnings"])
    narrow = ok(env, "select", "--file", str(path), "--range", "B1:D2")      # 범위를 주면 그 범위가 표
    assert narrow["count"] == 1 and narrow["anchor"]["header_row"] == 1


def test_open_unsaved_workbook_is_not_read(env, monkeypatch, capsys):
    path = make_bom(env["dir"] / "BOM.xlsx")
    monkeypatch.setattr(rt, "open_workbooks", lambda: [{"path": path, "name": "BOM.xlsx", "saved": False}])
    with pytest.raises(SystemExit) as e:
        pk.main(["select", "--file", path])
    out = json.loads(capsys.readouterr().out)
    assert e.value.code == 1 and "저장한 뒤" in out["message"]
    # --file 없이: 열린 통합문서가 여럿이면 [후보]와 같은 번호의 전체 명령
    other = make_bom(env["dir"] / "B2.xlsx")
    monkeypatch.setattr(rt, "open_workbooks", lambda: [{"path": path, "name": "BOM.xlsx", "saved": True},
                                                      {"path": other, "name": "B2.xlsx", "saved": True}])
    with pytest.raises(SystemExit):
        pk.main(["select"])
    out = json.loads(capsys.readouterr().out)
    assert "[후보] 1) BOM.xlsx 2) B2.xlsx" in out["message"] and len(out["confirm_commands"]) == 2


# ---------------------------------------------------------------------------
# 두 표: diff, join, only, stack
# ---------------------------------------------------------------------------

def test_diff_adds_deletes_changed_cells_with_rows_and_paint_files(env):
    old, new = _two_boms(env)
    ok(env, "select", "--file", old, "--name", "df1")
    ok(env, "select", "--file", new, "--name", "df2")
    out = ok(env, "diff", "--old", "df1", "--new", "df2", "--on", "Parent,Part Number")
    assert out["summary"] == "추가 1행, 삭제 1행, 바뀐 칸 2곳" and out["count"] == 4
    rows = ok(env, "show")["rows"]
    changed = {(r["Part Number"], r["열"]): (r["옛 값"], r["새 값"], r["df1 행"], r["df2 행"]) for r in rows if r["구분"] == "바뀜"}
    assert changed == {("0200", "Maker"): ("A", "B", 7, 6), ("0105", "Qty"): (3, 5, 8, 7)}
    assert [(r["구분"], r["df1 행"], r["df2 행"]) for r in rows if r["구분"] != "바뀜"] == [("삭제", 6, None), ("추가", None, 8)]
    old_cells, new_cells = [open(p, encoding="utf-8-sig").read().split() for p in out["paint_files"]]
    assert old_cells == ["B6:G6", "F7", "E8"] and new_cells == ["B8:G8", "F6", "E7"]
    cmd = command_args(out["confirm_commands"][0]["command"])
    assert cmd[:3] == ["style", "--target", old] and "--file" in cmd and "next_commands" in out
    # 집 PC 2026-11-13: split()은 따옴표를 남겨 Windows 경로('C:\...\excel.py')에서 실패함 -> 따옴표를 푸는 shlex (env는 bash 따옴표)
    assert shlex.split(out["confirm_commands"][0]["command"])[1].endswith(os.path.join("office-cli", "excel.py"))


def test_diff_same_number_in_int_and_decimal_columns_is_not_a_change(env):
    # 집 PC 2026-11-13: 한 판에만 소수 수량(0.5)이 있으면 Qty 열이 정수/소수로 달라 2와 2.0을 바뀐 칸으로 셌음
    old = make_bom(env["dir"] / "v1.xlsx")
    new = make_bom(env["dir"] / "v2.xlsx", rows=[(1, "TOP", "0102", 2, "A", "볼트"), (2, "0102", "0105", 4, "B", "너트"),
                                                 (1, "TOP", "0200", 1, "A", "판"), (2, "0200", "0105", 3.5, "C", "너트")])
    ok(env, "select", "--file", old, "--name", "df1")
    ok(env, "select", "--file", new, "--name", "df2")
    out = ok(env, "diff", "--old", "df1", "--new", "df2", "--on", "Parent,Part Number")
    assert out["summary"] == "추가 0행, 삭제 0행, 바뀐 칸 1곳"
    row = ok(env, "show")["rows"][0]
    assert (row["열"], row["옛 값"], row["새 값"]) == ("Qty", 3, 3.5)

def test_diff_refuses_duplicate_keys(env):
    old, new = _two_boms(env)
    ok(env, "select", "--file", old, "--name", "df1")
    ok(env, "select", "--file", new, "--name", "df2")
    code, out = run(env, "diff", "--old", "df1", "--new", "df2", "--on", "Part Number")
    assert code == 1 and "같은 행이 1개" in out["message"]


def test_join_only_stack_keep_both_row_numbers(env):
    old, new = _two_boms(env)
    ok(env, "select", "--file", old, "--name", "df1")
    ok(env, "select", "--file", new, "--name", "df2")
    j = ok(env, "join", "--left", "df1", "--right", "df2", "--on", "Parent,Part Number")
    assert j["count"] == 3 and j["table"] == "df3"
    # 집 PC 2026-11-13: "df1와 df2를"처럼 받침에 맞지 않는 조사가 preview에 나왔음
    assert "df1, df2을(를) Parent,Part Number 기준으로 짝 맞춤" in j["preview"] and "df1와" not in j["preview"]
    row = ok(env, "show")["rows"][0]
    assert row["df1 행"] == 5 and row["df2 행"] == 5 and "Qty_df2" in row
    o = ok(env, "only", "--left", "df2", "--right", "df1", "--on", "Parent,Part Number", "--name", "새것")
    assert o["count"] == 1 and ok(env, "show", "--from", "새것")["rows"][0]["Part Number"] == "0300"
    assert "(Parent,Part Number 기준으로 df1과(와) 비교)" in o["preview"]
    s = ok(env, "stack", "--tables", "df1,df2")
    shown = ok(env, "show", "--from", s["table"])
    assert s["count"] == 8 and shown["rows"][0]["표"] == "df1" and shown["rows"][-1]["표"] == "df2"
    code, out = run(env, "stack", "--tables", "df1;df2")
    assert code == 1 and recommended(out["message"])[-1] == "df1,df2"


def test_join_with_different_key_names_and_back_by_excel_row(env):
    # 집 PC 2026-11-13 (MBOM): Part Number <-> Routing Item처럼 이름이 다른 키, 스킬이 계산한 csv를 엑셀 행으로 되돌리기
    ok(env, "select", "--file", make_bom(env["dir"] / "BOM.xlsx"), "--name", "df1")
    rt_path = env["dir"] / "routing.csv"
    rt_path.write_text("Item,Maker,Time\n0102,A,5\n0200,A,7\n0105,C,9\n", encoding="utf-8")
    ok(env, "select", "--file", str(rt_path), "--name", "rt")
    j = ok(env, "join", "--left", "df1", "--right", "rt", "--on", "Part Number=Item,Maker")
    assert j["count"] == 3 and "Item" not in [c["name"] for c in j["columns"]]       # 0102/A, 0200/A, 0105/C
    code, out = run(env, "join", "--left", "df1", "--right", "rt", "--on", "Part Number")
    assert code == 1 and "[후보]" in out["message"] and "Item" in out["message"]
    fixed = [c["command"] for c in out["confirm_commands"] if "Part Number=Item" in c["command"]]
    assert fixed and ok(env, *command_args(fixed[0]))["count"] == 4
    exp = ok(env, "export", "--from", "df1", "--save", "df1.csv")
    calc = pd.read_csv(exp["output_file"], encoding="utf-8-sig", dtype=str)[["엑셀 행", "Part Number"]]
    calc["계산"] = ["a", "b", "c", "d"]
    calc_path = env["dir"] / "skill.csv"
    calc.to_csv(calc_path, index=False, encoding="utf-8-sig")
    sk = ok(env, "select", "--file", str(calc_path), "--name", "sk")
    assert "엑셀 행번호" in [c["name"] for c in sk["columns"]] and any("엑셀 행=엑셀 행번호" in w for w in sk["warnings"])
    back = ok(env, "join", "--left", "df1", "--right", "sk", "--on", "엑셀 행=엑셀 행번호", "--name", "back")
    rows = ok(env, "show", "--from", "back")["rows"]
    assert back["count"] == 4 and [(r["df1 행"], r["계산"]) for r in rows] == [(5, "a"), (6, "b"), (7, "c"), (8, "d")]


def test_group_counts_sums_and_warns_mixed_units(env):
    def extra(ws):
        ws["H4"] = "UoM"
        for r, u in enumerate(["EA", "KG", "EA", "EA"], start=5):
            ws.cell(row=r, column=8, value=u)
    ok(env, "select", "--file", make_bom(env["dir"] / "BOM.xlsx", extra=extra), "--name", "df1")
    g = ok(env, "group", "--by", "Maker", "--sum", "Qty", "--name", "g")
    rows = ok(env, "show", "--from", "g")["rows"]
    assert g["count"] == 3 and rows[0] == {"Maker": "A", "행 수": 2, "Qty 합계": 3, "엑셀 행 목록": "5, 7"}
    assert not g.get("warnings")                                        # A·C는 EA만, B는 KG만
    mixed = ok(env, "group", "--from", "df1", "--by", "Level", "--sum", "Qty")
    warn = next(w for w in mixed["warnings"] if "단위가 다른 행" in w)
    assert "--by Level,UoM" in warn and "pandas_kit.py" in warn
    code, err = run(env, "group", "--from", "df1", "--by", "Maker", "--sum", "Desc")
    assert code == 1 and "[후보] 1) Level 2) Qty" in err["message"]
    assert command_args(err["confirm_commands"][1]["command"])[-2:] == ["--sum", "Qty"]
    assert ok(env, "group", "--from", "df1", "--by", "Maker")["columns"][1]["name"] == "행 수"


def test_sort_keeps_excel_rows_and_puts_empty_cells_last(env):
    rows = [(1, "TOP", "0102", 2, "A", "x"), (2, "0102", "0105", None, "B", "x"), (1, "TOP", "0200", 9, "A", "x"),
            (2, "0200", "0105", 2, "C", "x")]
    ok(env, "select", "--file", make_bom(env["dir"] / "BOM.xlsx", rows=rows), "--name", "df1")
    up = ok(env, "sort", "--by", "Qty", "--name", "up")
    assert "작은 것부터" in up["preview"]
    assert [r["엑셀 행"] for r in ok(env, "show", "--from", "up")["rows"]] == [5, 8, 7, 6]     # 같은 값은 표 순서, 빈 칸은 끝
    ok(env, "sort", "--from", "df1", "--by", "Maker,Qty", "--desc", "--name", "down")
    assert [r["엑셀 행"] for r in ok(env, "show", "--from", "down")["rows"]] == [8, 6, 7, 5]


# ---------------------------------------------------------------------------
# 상태: 단계 기록, 스냅샷, undo, new, --work, 원본 변경
# ---------------------------------------------------------------------------

def test_undo_and_redefined_name_warns_dependents(env):
    path = make_bom(env["dir"] / "BOM.xlsx")
    ok(env, "select", "--file", path, "--name", "df1")
    ok(env, "select", "--from", "df1", "--where", "Level=2", "--name", "df2")
    again = ok(env, "select", "--from", "df1", "--where", "Level=1", "--name", "df1")
    assert any("df2은(는) 옛 df1 기준" in w for w in again["warnings"])
    u = ok(env, "undo")
    assert u["removed"] == "df1" and u["current"] == "df2" and "그 전 표를 다시 가리킵니다" in u["preview"]
    assert ok(env, "show", "--from", "df1")["count"] == 4
    ok(env, "undo")
    ok(env, "undo")
    code, out = run(env, "undo")
    assert code == 1 and "되돌릴 단계가 없습니다" in out["message"]


def test_broken_or_other_pandas_snapshot_is_rebuilt_from_record(env):
    path = make_bom(env["dir"] / "BOM.xlsx")
    work = ok(env, "select", "--file", path, "--name", "df1")["work"]
    ok(env, "select", "--from", "df1", "--where", "Maker=A", "--name", "df2")
    with open(os.path.join(work, "tables", "s1.pkl"), "wb") as f:
        f.write(b"broken")
    with open(os.path.join(work, "tables", "s2.pkl"), "wb") as f:
        pickle.dump({"pandas": "0.0.1", "df": pd.DataFrame({"x": [1]}), "meta": {}}, f)
    out = ok(env, "show", "--from", "df2")
    assert out["count"] == 2 and len(out["recovered"]) == 2 and "작업 기록으로 다시 만듦" in out["recovered"][0]


def test_new_then_work_option_returns_to_old_work(env):
    path = make_bom(env["dir"] / "BOM.xlsx")
    work = ok(env, "select", "--file", path)["work"]
    assert ok(env, "new")["previous_work"] == work
    code, out = run(env, "show")
    assert code == 1 and "지금 작업 폴더가 없습니다" in out["message"]
    assert ok(env, "show", "--work", work)["table"] == "df1"
    assert ok(env, "show")["table"] == "df1"              # --work로 돌아간 폴더가 지금 작업 폴더
    time.sleep(1.1)
    w2 = ok(env, "new") and ok(env, "select", "--file", path)["work"]
    assert w2 != work and os.path.dirname(w2) == os.path.dirname(work)


def test_changed_source_is_warned_with_command_to_read_again(env):
    path = make_bom(env["dir"] / "BOM.xlsx")
    ok(env, "select", "--file", path, "--range", "B4:G8", "--name", "df1")
    time.sleep(1.1)
    make_bom(path, rows=[(1, "TOP", "9999", 1, "A", "x")])
    out = ok(env, "show")
    w = " ".join(out["warnings"])
    assert "읽은 뒤 바뀌었습니다" in w and "select --file" in w and "--range B4:G8" in w and "--name df1" in w


def test_show_truncates_at_20_rows_with_full_result_file(env):
    rows = [(1, "TOP", f"{i:04d}", i, "A", "x") for i in range(30)]
    ok(env, "select", "--file", make_bom(env["dir"] / "BOM.xlsx", rows=rows))
    out = ok(env, "show")
    assert out["truncated"] is True and out["shown"] == 20 and out["count"] == 30
    full = pd.read_csv(out["full_result_file"], encoding="utf-8-sig", dtype=str)
    assert len(full) == 30 and full["Part Number"].iloc[0] == "0000" and "엑셀 행" in full.columns
    nxt = out["next_commands"][0]["command"]
    assert command_args(nxt)[-2:] == ["--start", "21"]
    more = ok(env, *command_args(nxt))
    assert more["start"] == 21 and more["shown"] == 10 and more["rows"][0]["엑셀 행"] == 25 and "next_commands" not in more
    assert "21~30번째 행" in more["preview"]
    code, err = run(env, "show", "--start", "99")
    assert code == 1 and recommended(err["message"])[-2:] == ["--start", "1"]


# ---------------------------------------------------------------------------
# export: 새 파일, 덮어쓰기는 --overwrite, 원본에는 쓰지 않음
# ---------------------------------------------------------------------------

def test_export_new_file_timestamp_overwrite_and_never_the_source(env):
    path = make_bom(env["dir"] / "BOM.xlsx")
    ok(env, "select", "--file", path, "--where", "Level=2", "--name", "하위")
    out = ok(env, "export")
    assert out["agent_directive"].startswith("VERIFIED:") and out["output_file"].endswith("하위.xlsx")
    ws = openpyxl.load_workbook(out["output_file"]).active
    assert [c.value for c in ws[1]][:3] == ["엑셀 행", "Level", "Parent"] and ws["D2"].value == "0105"
    again = ok(env, "export")
    assert again["output_file"] != out["output_file"] and "--overwrite" in again["recovered"][0]
    assert ok(env, "export", "--overwrite")["output_file"] == out["output_file"]
    csv_out = ok(env, "export", "--save", "하위.csv")
    assert pd.read_csv(csv_out["output_file"], encoding="utf-8-sig", dtype=str)["Part Number"].tolist() == ["0105", "0105"]
    code, err = run(env, "export", "--save", path, "--overwrite")
    assert code == 1 and "원본 파일" in err["message"]
    code, err = run(env, "export", "--save", "x.pdf")
    assert code == 1 and recommended(err["message"])[-2:] == ["--save", "x.xlsx"]


# ---------------------------------------------------------------------------
# chart: 표를 그림(.png)으로 (표는 바꾸지 않음, 단계로 남기지 않음)
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("kind, x", [("bar", "Part Number"), ("line", "Part Number"), ("scatter", "Level"), ("pie", "Maker"),
                                     ("hist", None), ("box", None)])
def test_chart_kinds_save_png_and_verify(env, kind, x):
    import matplotlib.image as mpimg
    ok(env, "select", "--file", make_bom(env["dir"] / "BOM.xlsx"), "--name", "df1")
    out = ok(env, "chart", "--from", "df1", "--kind", kind, *(["--x", x] if x else []), "--y", "Qty")
    assert out["agent_directive"].startswith("VERIFIED:") and out["output_file"].endswith(f"df1_{kind}.png")
    assert mpimg.imread(out["output_file"]).shape[0] > 0 and out["kind"] == kind and out["y"] == ["Qty"]
    assert ok(env, "undo")["removed"] == "df1"                       # 차트는 단계가 아님: 되돌리면 표가 지워짐


def test_chart_sums_repeated_x_limits_bars_and_checks_columns(env):
    rows = [(1, "TOP", f"{i:04d}", 1, "A" if i % 2 else "B", "x") for i in range(40)]
    ok(env, "select", "--file", make_bom(env["dir"] / "BOM.xlsx", rows=rows), "--name", "df1")
    out = ok(env, "chart", "--x", "Maker", "--y", "Qty")
    assert out["count"] == 2 and any("합계로" in w for w in out["warnings"])
    many = ok(env, "chart", "--x", "Part Number", "--y", "Qty,Level")
    assert many["count"] == 30 and many["truncated"] is True and many["output_file"] != out["output_file"]
    assert "--overwrite" in many["recovered"][0]
    code, err = run(env, "chart", "--y", "Desc")
    assert code == 1 and "[후보] 1) Level 2) Qty" in err["message"] and len(err["confirm_commands"]) == 2
    code, err = run(env, "chart", "--kind", "pie", "--x", "Maker", "--y", "Qty,Level")
    assert code == 1 and recommended(err["message"])[-2:] == ["--y", "Qty"]
    code, err = run(env, "chart", "--kind", "pie", "--y", "Qty")
    assert code == 1 and "--x가 필요합니다" in err["message"]
    code, err = run(env, "chart", "--y", "Qty", "--save", "그림.jpg")
    assert code == 1 and recommended(err["message"])[-2:] == ["--save", "그림.png"]
    named = ok(env, "chart", "--y", "Qty", "--save", "수량", "--title", "부품별 수량")
    assert named["output_file"] == os.path.join(named["work"], "수량.png")

def test_chart_warns_when_units_are_mixed(env):
    # 집 PC 2026-11-13 (MBOM): 업체별 Qty 막대가 EA·KG 수량을 그대로 더함
    def extra(ws):
        ws["H4"] = "UoM"
        for r, u in enumerate(["EA", "KG", "EA", "EA"], start=5):
            ws.cell(row=r, column=8, value=u)
    ok(env, "select", "--file", make_bom(env["dir"] / "BOM.xlsx", extra=extra), "--name", "df1")
    out = ok(env, "chart", "--x", "Maker", "--y", "Qty")
    warn = next(w for w in out["warnings"] if "'UoM' 열의 단위가 여러 가지(EA, KG)" in w)
    assert "더해 그렸습니다" in warn and "select --from df1 --where UoM=EA" in warn and "pandas_kit.py" in warn
    ok(env, "select", "--from", "df1", "--where", "UoM=EA", "--name", "ea")
    assert not any("단위" in w for w in ok(env, "chart", "--from", "ea", "--x", "Maker", "--y", "Qty").get("warnings", []))


def test_chart_of_empty_table_is_zero(env):
    ok(env, "select", "--file", make_bom(env["dir"] / "BOM.xlsx"), "--where", "Level=9", "--name", "없음")
    out = ok(env, "chart", "--y", "Qty")
    assert out["count"] == 0 and "0건" in out["agent_directive"] and "output_file" not in out



# ---------------------------------------------------------------------------
# 시험 자료: tools/make_mbom_samples.py (클라우드에서도 같은 MBOM 두 판을 만들어 씀)
# ---------------------------------------------------------------------------

def test_mbom_samples_are_made_and_read(env):
    import subprocess
    script = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "tools", "make_mbom_samples.py")
    out_dir = env["dir"] / "mbom"
    p = subprocess.run([sys.executable, script, "--out", str(out_dir)], capture_output=True, text=True, encoding="utf-8")
    assert p.returncode == 0, p.stderr
    v1 = ok(env, "select", "--file", str(out_dir / "MBOM_DR200.xlsx"), "--name", "v1")
    assert v1["anchor"]["range"] == "B6:R116" and v1["count"] == 110          # 제목 블록·메모는 표가 아님
    assert any("표 위의 다른 내용" in w for w in v1["warnings"]) and any("숨긴 행" in w for w in v1["warnings"])
    assert any("Qty(H) 열의 글자로 저장된 숫자 1칸" in w for w in v1["warnings"])
    v2 = ok(env, "select", "--file", str(out_dir / "MBOM_DR200_v2.xlsx"), "--name", "v2")
    assert v2["count"] == 109
    bolts = ok(env, "select", "--from", "v1", "--where", "Part Number=000125")
    assert bolts["count"] == 10


# ---------------------------------------------------------------------------
# MBOM 두 판으로 써 보며 찾은 '결과가 틀리는데 알리지 않음' (2026-11-14, 개선점_MBOM.md E)
# ---------------------------------------------------------------------------

@pytest.fixture
def mbom(env):
    import subprocess
    script = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "tools", "make_mbom_samples.py")
    out_dir = env["dir"] / "mbom"
    p = subprocess.run([sys.executable, script, "--out", str(out_dir)], capture_output=True, text=True, encoding="utf-8")
    assert p.returncode == 0, p.stderr
    v1 = ok(env, "select", "--file", str(out_dir / "MBOM_DR200.xlsx"), "--name", "v1")
    ok(env, "select", "--file", str(out_dir / "MBOM_DR200_v2.xlsx"), "--name", "v2")
    ok(env, "select", "--file", str(out_dir / "MBOM_DR200.xlsx"), "--sheet", "Routing", "--name", "rt")
    return {"dir": out_dir, "v1": v1}


def test_mbom_title_rows_counted_and_number_written_differently_is_warned(env, mbom):
    assert any("표 위의 다른 내용 4행(엑셀 행 1~4)" in w for w in mbom["v1"]["warnings"])     # 한 칸짜리 제목 줄도 셈
    out = ok(env, "select", "--from", "v1", "--where", "Part Number=125")
    assert out["count"] == 0 and any("적힌 값: 000125" in w and "Part Number=000125" in w for w in out["warnings"])
    out = ok(env, "select", "--from", "v1", "--where", "Op No=10")
    assert out["count"] == 0 and any("칸 36개" in w and "0010" in w for w in out["warnings"])
    assert "warnings" not in ok(env, "select", "--from", "v1", "--where", "Part Number=000125")   # 맞게 쓰면 조용함


def test_mbom_date_condition_must_be_a_full_date(env, mbom):
    for bad in ("2026-13", "2026-12", "내년", "2026-02-30"):
        code, out = run(env, "select", "--from", "v1", "--where", f"Eff From>{bad}")
        assert code == 1 and "연-월-일" in out["message"], (bad, out)
    assert ok(env, "select", "--from", "v1", "--where", "Eff From>2026-12-31")["count"] == 1


def test_mbom_join_says_unmatched_rows_and_ranks_key_candidates_by_shared_values(env, mbom):
    code, out = run(env, "join", "--left", "v1", "--right", "rt", "--on", "Part Number")
    assert code == 1 and "[후보] 1) Item (겹치는 값 6개)" in out["message"]
    assert command_args(out["confirm_commands"][0]["command"])[-2:] == ["--on", "Part Number=Item"]
    j = ok(env, "join", "--left", "v1", "--right", "rt", "--on", "Part Number=Item")
    assert j["count"] == 14 and any("짝이 없어 빠진 행: v1 104행(전체 110행), rt 0행(전체 14행)" in w for w in j["warnings"])


def test_mbom_only_and_diff_with_repeated_keys(env, mbom):
    o = ok(env, "only", "--left", "v1", "--right", "v2", "--on", "Part Number")
    assert o["count"] == 1 and any("v1에 34개, v2에 34개" in w for w in o["warnings"])
    code, out = run(env, "diff", "--old", "v1", "--new", "v2", "--on", "Part Number")
    assert code == 1 and "000125 (엑셀 행 16, 30, 35, 38, 58 등)" in out["message"] and "상위 품번" not in out["message"]


def test_mbom_export_keeps_dates_as_dates(env, mbom):
    out = ok(env, "export", "--from", "v1")
    ws = openpyxl.load_workbook(out["output_file"]).active
    head = [c.value for c in ws[1]]
    cell = ws.cell(row=2, column=head.index("Eff From") + 1)
    assert cell.number_format == "yyyy-mm-dd" and cell.value.year == 2026
