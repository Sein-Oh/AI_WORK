# -*- coding: utf-8 -*-
"""pandas-kit 0.3.0: pandas에서 자주 쓰는 메서드를 명령으로 (사용자 결정 2026-11-15). 명령마다 정상, 0건, 틀린 열 [후보],
구분자 [추천], 스냅샷을 지운 뒤 기록으로 다시 만들기. 도구는 서브프로세스로 실행한다 (stdout 1줄 확인)."""
import glob
import os

import openpyxl
import pytest

from pk_helpers import command_args, ok, recommended, run

HEAD = ["Part Number", "Desc", "Maker", "Qty", "Unit Price", "UoM", "Note"]
ROWS = [("0102", "볼트 M3", "A", 2, 1.5, "EA", "x"),            # 엑셀 행 2
        ("0105", "너트 M3", "B", 4, 0.5, "EA", None),
        ("0200", "판", "A", 1, 10, "EA", "12"),
        ("0105", "너트 M3", "C", 3, 0.5, "EA", None),
        ("RM-1", "강판", "A", 0.5, None, "KG", "abc"),
        ("0300", "와셔", "B", 0, 2, "EA", None)]              # 엑셀 행 7


def make_parts(path, rows=ROWS, head=HEAD):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Parts"
    ws.append(head)
    for r in rows:
        ws.append(list(r))
    for cell in ws["A"]:
        cell.number_format = "@"
    wb.save(path)
    return str(path)


@pytest.fixture
def t(env):
    """표 t (6행 7열)와 0행 표 none."""
    out = ok(env, "select", "--file", make_parts(env["dir"] / "parts.xlsx"), "--name", "t")
    ok(env, "select", "--from", "t", "--where", "Maker=Z", "--name", "none")
    return {"env": env, "work": out["work"]}


def rows_of(env, name, **kw):
    out = ok(env, "show", "--from", name, *[x for k, v in kw.items() for x in (f"--{k}", v)])
    return out["rows"]


def drop_snapshots(work):
    for p in glob.glob(os.path.join(work, "tables", "*.pkl")):
        os.remove(p)


# ---------------------------------------------------------------------------
# 1차: describe, unique, group 집계, calc, --where ^= $=, top
# ---------------------------------------------------------------------------

def test_describe_gives_one_row_per_column(t):
    env = t["env"]
    out = ok(env, "describe", "--from", "t", "--name", "d")
    assert out["agent_directive"].startswith("VERIFIED:") and out["count"] == 7
    assert "열 7개 요약 (숫자 열 2개, 그 밖의 열 5개" in out["preview"]
    rows = {r["열"]: r for r in rows_of(env, "d")}
    q = rows["Qty"]
    assert (q["종류"], q["개수"], q["빈 칸"], q["평균"], q["최소"], q["50%"], q["최대"]) == ("숫자", 6, 0, 1.75, 0, 1.5, 4)
    assert q["표준편차"] == pytest.approx(1.541104, abs=1e-6) and q["종류 수"] is None
    assert (rows["Unit Price"]["개수"], rows["Unit Price"]["빈 칸"]) == (5, 1)
    m = rows["Maker"]
    assert (m["종류"], m["개수"], m["종류 수"], m["가장 많은 값"], m["가장 많은 값 개수"], m["평균"]) == ("글자", 6, 3, "A", 3, None)
    assert rows["Part Number"]["가장 많은 값"] == "0105"                 # 앞자리 0 그대로
    assert not out.get("warnings")                                       # 코드 열(앞자리 0)은 '숫자로 읽히는 글자 열'로 알리지 않음
    out = ok(env, "describe", "--from", "t", "--columns", "Qty,Maker")
    assert out["count"] == 2
    out = ok(env, "describe", "--from", "none")                          # 0행 표: 열마다 개수 0
    assert out["count"] == 7 and all(r["개수"] == 0 for r in rows_of(env, out["table"]))


def test_describe_says_when_a_text_column_reads_as_numbers(env):
    rows = [("A", "12 kg"), ("B", "7"), ("C", "8"), ("D", "9"), ("E", "10"), ("F", "11"), ("G", "12"), ("H", "13"), ("I", "14"), ("J", "15")]
    ok(env, "select", "--file", make_parts(env["dir"] / "w.xlsx", rows, ["Name", "Weight"]), "--name", "w")
    out = ok(env, "describe", "--from", "w")
    assert any("글자 열인데 숫자로 읽히는 칸이 대부분인 열: Weight" in w for w in out["warnings"])


def test_unique_counts_values_of_each_column(t):
    env = t["env"]
    out = ok(env, "unique", "--from", "t", "--columns", "Maker", "--name", "u")
    assert out["count"] == 3 and "서로 다른 값 (Maker 3가지)" in out["preview"]
    assert [(r["열"], r["값"], r["개수"], r["엑셀 행 목록"]) for r in rows_of(env, "u")] == [
        ("Maker", "A", 3, "2, 4, 6"), ("Maker", "B", 2, "3, 7"), ("Maker", "C", 1, "5")]
    out = ok(env, "unique", "--from", "t", "--columns", "Unit Price,Note", "--name", "u2")
    got = [(r["열"], r["값"], r["개수"]) for r in rows_of(env, "u2")]
    assert ("Unit Price", "0.5", 2) == got[0] and ("Unit Price", "10", 1) in got and ("Unit Price", "(빈 칸)", 1) in got
    assert ("Note", "(빈 칸)", 3) in got and "Unit Price 5가지, Note 4가지" in out["preview"]
    out = ok(env, "unique", "--from", "none", "--columns", "Maker")
    assert out["count"] == 0 and "0건" in out["agent_directive"]


def test_unique_wrong_column_and_separator(t):
    env = t["env"]
    code, out = run(env, "unique", "--from", "t", "--columns", "Makr")
    assert code == 1 and "[후보] 1) Maker" in out["message"]
    assert command_args(out["confirm_commands"][0]["command"]) == ["unique", "--from", "t", "--columns", "Maker"]
    code, out = run(env, "unique", "--from", "t", "--columns", "Maker;UoM")
    assert code == 1 and recommended(out["message"]) == ["unique", "--from", "t", "--columns", "Maker,UoM"]
    code, out = run(env, "unique", "--columns", "Maker")                # --from 없음: 표가 둘이라 [후보]
    assert code == 1 and "[후보]" in out["message"] and "--from" in out["confirm_commands"][0]["command"]


def test_group_mean_min_max_count_nunique(t):
    env = t["env"]
    out = ok(env, "group", "--from", "t", "--by", "Maker", "--sum", "Qty", "--mean", "Qty", "--min", "Qty", "--max", "Unit Price",
             "--count", "Note", "--nunique", "Part Number", "--name", "g")
    assert "(행 수, Qty 합계, Qty 평균, Qty 최소, Unit Price 최대, Note 개수, Part Number 종류 수)" in out["preview"]
    a = rows_of(env, "g")[0]
    assert (a["Maker"], a["행 수"], a["Qty 합계"], a["Qty 평균"], a["Qty 최소"], a["Unit Price 최대"], a["Note 개수"],
            a["Part Number 종류 수"], a["엑셀 행 목록"]) == ("A", 3, 3.5, 1.166667, 0.5, 10, 3, 3, "2, 4, 6")
    b = rows_of(env, "g")[1]
    assert (b["Qty 합계"], b["Qty 평균"], b["Note 개수"], b["Part Number 종류 수"]) == (4, 2, 0, 2)
    assert any("'UoM' 열의 단위가 다른 행" in w for w in out["warnings"])          # --sum과 같은 단위 섞임 경고
    out = ok(env, "group", "--from", "t", "--by", "Maker", "--mean", "Qty")     # --sum 없이 평균만이어도 같은 경고
    assert any("'UoM' 열의 단위가 다른 행" in w for w in out["warnings"])
    code, out = run(env, "group", "--from", "t", "--by", "Maker", "--mean", "Desc")
    assert code == 1 and "[후보] 1) Qty 2) Unit Price" in out["message"]
    assert command_args(out["confirm_commands"][0]["command"])[-2:] == ["--mean", "Qty"]
    code, out = run(env, "group", "--from", "t", "--by", "Maker", "--count", "Note;Desc")
    assert code == 1 and recommended(out["message"])[-2:] == ["--count", "Note,Desc"]


def test_calc_adds_columns_and_keeps_rows(t):
    env = t["env"]
    out = ok(env, "calc", "--from", "t", "--expr", "금액=Qty*[Unit Price];반=(Qty+1)/2;음=-Qty", "--name", "c")
    assert out["count"] == 6 and "계산한 열을 더함 (금액=Qty*[Unit Price], 반=(Qty+1)/2, 음=-Qty)" in out["preview"]
    assert [c["name"] for c in out["columns"]][-3:] == ["금액", "반", "음"]
    rows = rows_of(env, "c")
    assert [r["금액"] for r in rows] == [3, 2, 10, 1.5, None, 0] and [r["엑셀 행"] for r in rows] == [2, 3, 4, 5, 6, 7]
    assert [r["반"] for r in rows] == [1.5, 2.5, 1, 2, 0.75, 0.5] and rows[0]["음"] == -2
    assert any("'금액' 열: 식에 쓴 열이 빈 칸인 1행" in w for w in out["warnings"])
    assert rows[0]["Part Number"] == "0102"                                # 있는 열은 그대로
    out = ok(env, "calc", "--from", "c", "--expr", "배=금액*2+10", "--name", "c2")      # 정수가 되는 열, 만든 열을 다시 씀
    assert [r["배"] for r in rows_of(env, "c2")] == [16, 14, 30, 13, None, 10]
    out = ok(env, "calc", "--from", "t", "--expr", "a=Qty*2;b=a+1")                     # 앞 식의 새 열을 뒤 식에서
    assert [r["b"] for r in rows_of(env, out["table"])] == [5, 9, 3, 7, 2, 1]
    out = ok(env, "calc", "--from", "none", "--expr", "금액=Qty*2")
    assert out["count"] == 0 and "0건" in out["agent_directive"]


def test_calc_division_by_zero_and_cells_that_are_not_numbers(t):
    env = t["env"]
    out = ok(env, "calc", "--from", "t", "--expr", "역=1/Qty", "--name", "z")
    assert [r["역"] for r in rows_of(env, "z")][-1] is None
    assert any("'역' 열: 0으로 나눈 1행은 빈 값으로 두었습니다 (엑셀 행 7)" in w for w in out["warnings"])
    out = ok(env, "calc", "--from", "t", "--expr", "n=Note+1", "--name", "n")
    assert [r["n"] for r in rows_of(env, "n")] == [None, None, 13, None, None, None]
    assert any("Note 열의 숫자가 아닌 칸 2행은 계산하지 못해 빈 값으로 두었습니다 (엑셀 행 2, 6)" in w for w in out["warnings"])
    code, out = run(env, "calc", "--from", "t", "--expr", "n=Desc*2")
    assert code == 1 and "숫자로 읽히는 칸이 없어" in out["message"]


def test_calc_expression_mistakes(t):
    env = t["env"]
    code, out = run(env, "calc", "--from", "t", "--expr", "금액=Qty*Unit Price")           # 감싸지 않은 열 이름
    assert code == 1 and recommended(out["message"]) == ["calc", "--from", "t", "--expr", "금액=Qty*[Unit Price]"]
    code, out = run(env, "calc", "--from", "t", "--expr", "금액=Qty*Prce")                 # 틀린 열 이름
    assert code == 1 and "[후보] 1) Unit Price" in out["message"]
    assert command_args(out["confirm_commands"][0]["command"])[-1] == "금액=Qty*[Unit Price]"
    code, out = run(env, "calc", "--from", "t", "--expr", "a=Qty*2,b=Qty+1")              # 식 사이는 ';'
    assert code == 1 and recommended(out["message"])[-1] == "a=Qty*2;b=Qty+1"
    for bad, word in (("Qty=Qty*2", "이미 있는 열"), ("Qty*2", "모양이 아닙니다"), ("a=Qty**2", "읽을 수 없습니다"),
                      ("a=(Qty+1", "닫는 괄호"), ("a=Qty 2", "연산자가 없음"), ("a=__import__('os')", "읽을 수 없습니다"),
                      ("a=sum(Qty)", "함수는 쓸 수 없음"), ("a=Qty%2", "없")):
        code, out = run(env, "calc", "--from", "t", "--expr", bad)
        assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:") and word in out["message"], (bad, out["message"])


def test_where_starts_with_and_ends_with(t):
    env = t["env"]
    assert ok(env, "select", "--from", "t", "--where", "Part Number^=01")["count"] == 3
    assert ok(env, "select", "--from", "t", "--where", "Desc$=M3")["count"] == 3
    assert ok(env, "select", "--from", "t", "--where", "Part Number^=01|RM;Maker=A")["count"] == 2
    assert ok(env, "select", "--from", "t", "--where", "Qty^=0")["count"] == 2                # 숫자 열은 보이는 글자로 (0.5, 0)
    out = ok(env, "select", "--from", "t", "--where", "Desc^=없음")
    assert out["count"] == 0 and "0건" in out["agent_directive"]
    code, out = run(env, "select", "--from", "t", "--where", "Part Number^=")
    assert code == 1 and "비교할 값이 없습니다" in out["message"]


def test_top_takes_the_largest_rows(t):
    env = t["env"]
    out = ok(env, "top", "--from", "t", "--by", "Qty", "--n", "2", "--name", "big")
    assert out["count"] == 2 and "Qty 큰 순 앞 2행" in out["preview"]
    assert [(r["엑셀 행"], r["Qty"]) for r in rows_of(env, "big")] == [(3, 4), (5, 3)]
    out = ok(env, "top", "--from", "t", "--by", "Qty", "--n", "1", "--asc")
    assert rows_of(env, out["table"])[0]["엑셀 행"] == 7
    out = ok(env, "top", "--from", "t", "--by", "Unit Price")                               # --n 생략은 10, 빈 칸 행은 뺌
    assert out["count"] == 5 and any("'Unit Price' 열이 빈 칸인 1행은 순위에서 뺐습니다" in w for w in out["warnings"])
    code, out = run(env, "top", "--from", "t", "--by", "Maker")
    assert code == 1 and "[후보] 1) Qty 2) Unit Price" in out["message"]
    code, out = run(env, "top", "--from", "t", "--by", "Qty", "--n", "0")
    assert code == 1 and recommended(out["message"])[-2:] == ["--n", "10"]
    assert ok(env, "top", "--from", "none", "--by", "Qty")["count"] == 0


def test_first_round_tables_are_rebuilt_from_the_record(t):
    """스냅샷을 모두 지워도 기록(steps.json)으로 같은 표를 다시 만든다 (recovered)."""
    env = t["env"]
    ok(env, "describe", "--from", "t", "--name", "d")
    ok(env, "unique", "--from", "t", "--columns", "Maker,UoM", "--name", "u")
    ok(env, "group", "--from", "t", "--by", "Maker", "--mean", "Qty", "--nunique", "Part Number", "--name", "g")
    ok(env, "calc", "--from", "t", "--expr", "금액=Qty*[Unit Price];역=1/Qty", "--name", "c")
    ok(env, "top", "--from", "c", "--by", "금액", "--n", "3", "--name", "p")
    ok(env, "select", "--from", "t", "--where", "Part Number^=01", "--name", "w")
    before = {n: rows_of(env, n) for n in ("d", "u", "g", "c", "p", "w")}
    drop_snapshots(t["work"])
    for n, rows in before.items():
        out = ok(env, "show", "--from", n)
        assert out["rows"] == rows and any("작업 기록으로 다시 만듦" in r for r in out["recovered"]), n
