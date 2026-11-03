"""1.4.24 (15차 점검) 회귀 테스트.

F1: 제목으로 찾은 후보의 이름이 후보가 아닌 같은 시트 차트와 같으면 [후보] '시트!이름'이 다시 둘로 찾아졌음.
F2: FileHeader가 없는 OLE 파일(이름만 .hwp인 .doc/.xls)이 내부 오류(ERROR:)로 끝났음.
F3: style --highlight 값의 '"'가 Formula1을 깨뜨리고, 포함 규칙에서 ~ ? *가 와일드카드였음.
F4: write --append --formula가 행을 옮기고도 원래 행 기준 수식을 그대로 써 상대 참조가 원래 행을 가리켰음.
F5: 마크다운 코드 블록(```), '-' 1~2개 정렬 줄, '\\|' 칸 글자를 지원하지 않았음.
F6: ppt write --cards가 남은 공간을 보지 않아 긴 --text 아래에서 슬라이드 밖으로 나가고도 VERIFIED였음.
F7: reader parse_pages가 범위를 줄이기 전에 다 돌아 --units 1-999999999가 멈췄음.
"""
import json
import os
import struct
import subprocess
import sys

import pytest

import _runtime as rt
import excel
import ppt
import reader
import test_excel
import test_ppt
from test_review_v1423 import _T, _charts, _fill, _run

xl = test_excel.xl
pp = test_ppt.pp


# F1 차트 후보 이름 --------------------------------------------------------------------------------------------

def test_chart_labels_count_all_charts_on_sheet(xl, capsys):
    app, wb = xl
    _fill(wb)
    _run(excel, capsys, "chart", "--sheet", "Sheet1", "--range", "A1:B3", "--title", "Sales", "--name", "Chart 1")
    _run(excel, capsys, "chart", "--sheet", "Sheet1", "--range", "A1:B3", "--title", "Other", "--name", "Chart 2")
    _run(excel, capsys, "chart", "--sheet", "Sheet1", "--range", "A1:B3", "--title", "Sales", "--name", "Chart 3")
    excel._chart_objects(wb.sheets[0])[1].Name = "Chart 1"     # 후보가 아닌 2번 차트가 1번과 이름이 같음
    assert [n for n, _ in _charts(wb)["Sheet1"]] == ["Chart 1", "Chart 1", "Chart 3"]
    code, out = _run(excel, capsys, "chart", "--action", "delete", "--name", "Sales")
    assert code == 1 and "[후보]" in out["message"], out
    cands = out["message"].split("[후보]")[1].strip().split(", ")
    assert cands == ["Sheet1!#1", "Sheet1!Chart 3"], cands
    found = [excel._chart_title(excel.find_chart(_T(wb), c)[1]) for c in cands]    # 후보마다 차트 하나로 찾아짐
    assert found == ["Sales", "Sales"]


# F2 FileHeader 없는 OLE ---------------------------------------------------------------------------------------

def _ole_without_fileheader(path):
    """Root Entry만 있는 최소 OLE 파일 (이름만 .hwp인 다른 OLE 문서 흉내)."""
    h = bytearray(512)
    h[0:8] = bytes.fromhex("D0CF11E0A1B11AE1")
    struct.pack_into("<HHHHH", h, 24, 0x3E, 3, 0xFFFE, 9, 6)
    struct.pack_into("<IIIIIIIII", h, 40, 0, 1, 1, 0, 4096, 0xFFFFFFFE, 0, 0xFFFFFFFE, 0)
    struct.pack_into("<109I", h, 76, *([0] + [0xFFFFFFFF] * 108))
    fats = struct.pack("<128I", *([0xFFFFFFFD, 0xFFFFFFFE] + [0xFFFFFFFF] * 126))
    d = bytearray(512)
    name = "Root Entry".encode("utf-16le") + b"\0\0"
    d[0:len(name)] = name
    struct.pack_into("<HBB", d, 64, len(name), 5, 1)
    struct.pack_into("<III", d, 68, 0xFFFFFFFF, 0xFFFFFFFF, 0xFFFFFFFF)
    struct.pack_into("<II", d, 116, 0xFFFFFFFE, 0)
    for i in range(1, 4):
        struct.pack_into("<III", d, i * 128 + 68, 0xFFFFFFFF, 0xFFFFFFFF, 0xFFFFFFFF)
    with open(path, "wb") as f:
        f.write(bytes(h) + fats + bytes(d))


def test_ole_without_fileheader_is_input_error(tmp_path, capsys):
    pytest.importorskip("olefile")
    p = tmp_path / "notreally.hwp"
    _ole_without_fileheader(str(p))
    with pytest.raises(rt.InputError) as ei:
        reader.parse_hwp(str(p))
    assert "HWP 5.0 문서가 아닙니다" in str(ei.value)
    code, out = _run(reader, capsys, "read", "--file", str(p))
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY") and out["error_type"] != "internal", out


# F3 --highlight 따옴표·와일드카드 --------------------------------------------------------------------------------

def _formula1(wb):
    ws = wb.Worksheets("Sheet1")
    return [fc.kw["Formula1"] for fcs in ws.fcs.values() for fc in fcs.items]


def test_highlight_escapes_quote_and_wildcards(xl, capsys):
    app, wb = xl
    _fill(wb)
    code, out = _run(excel, capsys, "style", "--range", "A1:B3", "--highlight", '포함:a"b*?~:yellow')
    assert code == 0, out
    assert _formula1(wb)[-1] == '=ISNUMBER(SEARCH("a""b~*~?~~",A1))'
    code, out = _run(excel, capsys, "style", "--range", "A1:B3", "--highlight", '=say "hi":pink')
    assert code == 0, out
    assert _formula1(wb)[-1] == '="say ""hi"""'


# F4 --append --formula 상대 참조 -------------------------------------------------------------------------------

def test_append_formula_rebased_to_target_rows(xl, capsys):
    app, wb = xl
    ws = wb.Worksheets("Sheet1")
    for r, row in enumerate([["월", "값", "두배"], ["1월", 3, 6], ["2월", 5, 10], ["3월", 7, None], ["4월", 9, None]], 1):
        for c, v in enumerate(row, 1):
            if v is not None:
                ws.set(r, c, v)
    code, out = _run(excel, capsys, "write", "--range", "C2:C3", "--formula", "=B2*2", "--append")
    assert code == 0 and out["range_written"] == "C4:C5", out
    assert ws.Cells(4, 3).Formula == "=B4*2"                      # 원래 행(B2)이 아니라 쓴 행 기준
    assert out["agent_directive"].startswith("VERIFIED")
    code, out = _run(excel, capsys, "write", "--range", "C2", "--formula", "=$B$2+B2", "--append")
    assert code == 0 and ws.Cells(6, 3).Formula == "=$B$2+B6", out    # $ 참조는 그대로


def test_append_formula_conversion_failure_rejected(xl, capsys, monkeypatch):
    app, wb = xl
    ws = wb.Worksheets("Sheet1")
    for r, row in enumerate([["값", "두배"], [3, 6]], 1):
        for c, v in enumerate(row, 1):
            ws.set(r, c, v)

    def boom(**kw):
        raise Exception("COM")
    monkeypatch.setattr(app, "ConvertFormula", boom, raising=False)
    code, out = _run(excel, capsys, "write", "--range", "B2", "--formula", "=A2*2", "--append")
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY") and "[추천]" in out["message"], out
    assert ws.Cells(3, 2).Formula in ("", None)


# F5 마크다운 코드 블록·표 -------------------------------------------------------------------------------------

def test_markdown_fenced_code_block_verbatim():
    blocks = rt.parse_markdown_blocks("앞\n```python\n# 주석\n- x\n    들여쓴 줄\n\n1. y\n```\n# 제목")
    assert blocks == [("paragraph", "앞"), ("code", "# 주석"), ("code", "- x"),
                      ("code", "    들여쓴 줄"), ("code", "1. y"), ("heading", (1, "제목"))]   # 1.4.25: 코드 줄은 "code"
    assert rt.parse_markdown_blocks("~~~\n## a\n~~~") == [("code", "## a")]
    assert rt.parse_markdown_blocks("```\n# 끝나지 않음") == [("code", "# 끝나지 않음")]


def test_markdown_table_short_alignment_and_escaped_pipe():
    rows = rt.parse_markdown_blocks("| 이름 | 식 |\n|:-|-:|\n| or | a\\|b |\n| - | 값 |")[0][1]
    assert rows == [["이름", "식"], ["or", "a|b"], ["-", "값"]]


# F6 카드 공간 ----------------------------------------------------------------------------------------------

def test_cards_go_to_new_slide_when_no_space(pp, capsys):
    text = "\\n".join(f"- 줄 {i}" for i in range(1, 16))
    code, out = test_ppt.run(capsys, "write", "--slide", "new", "--title", "핵심", "--text", text,
                             "--cards", "안정성:무중단|확장성:모듈")
    assert code == 0, out
    prs = pp["prs"]
    H, W = float(prs.PageSetup.SlideHeight), float(prs.PageSetup.SlideWidth)
    assert out["slides_written"] == [4, 5] and any("새 슬라이드 5" in w for w in out["warnings"]), out
    for n in (4, 5):
        for sh in test_ppt.live(prs.Slides(n)):
            assert float(sh.Top) + float(sh.Height) <= H + 0.5 and float(sh.Left) + float(sh.Width) <= W + 0.5, sh.Name
    assert any("안정성" in sh.TextFrame.TextRange.Text for sh in test_ppt.live(prs.Slides(5)))
    assert out["agent_directive"].startswith("VERIFIED")


def test_verify_inside_flags_off_slide_shape(pp):
    class Sh:
        Name, Left, Top, Width, Height = "카드", 36.0, 518.0, 300.0, 80.0
    prs = pp["prs"]
    problems = ppt._verify_inside(prs, [(4, Sh())])
    assert problems and "슬라이드 밖" in problems[0]


# F7 parse_pages 큰 범위 ----------------------------------------------------------------------------------------

def test_parse_pages_huge_range_is_fast():
    root = os.path.dirname(reader.__file__)
    code = ("import sys; sys.path.insert(0, %r); import reader; "
            "print(reader.parse_pages('1-999999999', 5), reader.parse_pages('999999990-999999999,2', 5))" % root)
    res = subprocess.run([sys.executable, "-c", code], capture_output=True, text=True, timeout=20)
    assert res.stdout.strip() == "[1, 2, 3, 4, 5] [2]", res.stderr
    with pytest.raises(rt.InputError):
        reader.parse_pages("999999990-999999999", 5)
    assert json.dumps(reader.parse_pages("0-2", 5)) == "[1, 2]"


# 1.4.25 (16차 점검): 코드 블록 줄은 ** __를 굵게로 읽지 않고 그대로, 인라인 코드 ```x```는 펜스가 아님 ------------------

import sys as _sys  # noqa: E402

import _runtime as _rt  # noqa: E402
import test_ppt as _tp  # noqa: E402
import test_word as _tw  # noqa: E402
import ppt as _ppt  # noqa: E402
import word as _word  # noqa: E402

wd = _tw.wd
pp = _tp.pp
_MD = "# 예제\n\n```python\ndef __init__(self):\n    return x**2 + y**2\n```\n"


def test_code_block_kind_and_inline_fence():
    blocks = _rt.parse_markdown_blocks(_MD)
    assert ("code", "def __init__(self):") in blocks and ("code", "    return x**2 + y**2") in blocks
    blocks = _rt.parse_markdown_blocks("```x``` 는 코드\n\n# 제목\n- 항목")
    assert ("heading", (1, "제목")) in blocks and any(k == "bullets" for k, _ in blocks)


def test_word_code_block_verbatim(wd, capsys, tmp_path):
    p = tmp_path / "c.md"
    p.write_text(_MD, encoding="utf-8")
    code, out = _tw.run(capsys, "write", "--file", str(p))
    txt = _tw.paras(wd["doc"])
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    assert "def __init__(self):" in txt and "    return x**2 + y**2" in txt, txt


def test_word_file_starting_with_code_block_verified(wd, capsys, tmp_path):
    p = tmp_path / "c2.md"
    p.write_text("```\nprint(a**2)\n```\n", encoding="utf-8")
    code, out = _tw.run(capsys, "write", "--file", str(p))
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out


def test_ppt_code_block_verbatim(pp, capsys, tmp_path):
    p = tmp_path / "c.md"
    p.write_text(_MD, encoding="utf-8")
    code, out = _tp.run(capsys, "write", "--file", str(p))
    prs = pp["prs"]
    s = prs.Slides(prs.Slides.Count)
    texts = [sh.TextFrame.TextRange.Text for sh in _tp.live(s) if hasattr(sh, "TextFrame")]
    assert code == 0 and any("__init__" in t and "x**2" in t for t in texts), texts
