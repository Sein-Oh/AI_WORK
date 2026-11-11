"""마크다운 해석(_markdown.parse_markdown_blocks 등): 제목·목록·표·코드 블록·굵게·구분선·머리말."""
import pytest

import _markdown as md
import excel
from office_helpers import CODE_MD


# ---------------------------------------------------------------------------
# parse_markdown_blocks 기본: 제목·목록·표·굵게
# ---------------------------------------------------------------------------


def test_markdown_blocks_and_bold():
    blocks = md.parse_markdown_blocks("# 제목\n본문 **굵게**\n이어짐\n\n- a\n- b\n1. x\n2) y\n| 공정 | 수율 |\n|---|---|\n| 가열 | 98 |\n---\n## 끝")
    assert blocks == [("heading", (1, "제목")), ("paragraph", "본문 **굵게** 이어짐"), ("bullets", ["a", "b"]),
                      ("numbered", ["x", "y"]), ("table", [["공정", "수율"], ["가열", "98"]]), ("break", None),
                      ("heading", (2, "끝"))]
    assert md.split_bold("가 **나** 다 __라__") == ("가 나 다 라", [(2, 3), (6, 7)])
    assert md.looks_like_markdown("- 항목") and not md.looks_like_markdown("그냥 문장")


# ---------------------------------------------------------------------------
# 마크다운 코드 블록·표
# ---------------------------------------------------------------------------


def test_markdown_fenced_code_block_verbatim():
    blocks = md.parse_markdown_blocks("앞\n```python\n# 주석\n- x\n    들여쓴 줄\n\n1. y\n```\n# 제목")
    assert blocks == [("paragraph", "앞"), ("code", "# 주석"), ("code", "- x"),
                      ("code", "    들여쓴 줄"), ("code", "1. y"), ("heading", (1, "제목"))]   # 1.4.25: 코드 줄은 "code"
    assert md.parse_markdown_blocks("~~~\n## a\n~~~") == [("code", "## a")]
    assert md.parse_markdown_blocks("```\n# 끝나지 않음") == [("code", "# 끝나지 않음")]


def test_markdown_table_short_alignment_and_escaped_pipe():
    rows = md.parse_markdown_blocks("| 이름 | 식 |\n|:-|-:|\n| or | a\\|b |\n| - | 값 |")[0][1]
    assert rows == [["이름", "식"], ["or", "a|b"], ["-", "값"]]


# ---------------------------------------------------------------------------
# 코드 블록 줄은 ** __를 굵게로 읽지 않고 그대로, 인라인 코드 ```x```는 펜스가 아님
# ---------------------------------------------------------------------------


def test_code_block_kind_and_inline_fence():
    blocks = md.parse_markdown_blocks(CODE_MD)
    assert ("code", "def __init__(self):") in blocks and ("code", "    return x**2 + y**2") in blocks
    blocks = md.parse_markdown_blocks("```x``` 는 코드\n\n# 제목\n- 항목")
    assert ("heading", (1, "제목")) in blocks and any(k == "bullets" for k, _ in blocks)


# ---------------------------------------------------------------------------
# excel 마크다운 표
# ---------------------------------------------------------------------------


def test_excel_load_table_file_first_markdown_table(tmp_path):
    p = tmp_path / "t.md"
    p.write_text("# 보고\n\n| 항목 | 값 |\n|--|--|\n| a | 1 |\n| b\\|c | 2 |\n\n본문\n\n"
                 "| 다른 | 표 |\n|---|---|\n| x | y |\n\n```\n| code | line |\n```\n", encoding="utf-8")
    assert excel.load_table_file(str(p)) == [["항목", "값"], ["a", 1], ["b|c", 2]]


def test_excel_load_table_file_pipe_only_in_code_is_lines(tmp_path):
    p = tmp_path / "c.md"
    p.write_text("```\n| a | b |\n```\n", encoding="utf-8")
    assert excel.load_table_file(str(p)) == [["```"], ["| a | b |"], ["```"]]


# ---------------------------------------------------------------------------
# 단어 안 __
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("text,plain", [
    ("report__2024__final.xlsx", "report__2024__final.xlsx"),
    ("파일__x__끝", "파일__x__끝"),
    ("__init__ 메서드", "init 메서드"),          # CommonMark: 양쪽이 단어 경계면 굵게
    ("(__a__)", "(a)"),
    ("x**2**y", "x2y"),                          # ** 는 단어 안에서도 굵게 (그대로)
    ("가 **나** 다 __라__", "가 나 다 라"),
])
def test_split_bold_intraword_underscore(text, plain):
    assert md.split_bold(text)[0] == plain


# ---------------------------------------------------------------------------
# 연도와 1이 아닌 번호
# ---------------------------------------------------------------------------


def test_parse_blocks_year_is_not_list_and_start_number_kept():
    assert md.parse_markdown_blocks("2026. 11. 3. 회의 결과") == [("paragraph", "2026. 11. 3. 회의 결과")]
    assert not md.looks_like_markdown("2026. 11. 3. 회의 결과를 보고합니다.")
    assert md.parse_markdown_blocks("3. 셋\n4. 넷") == [("paragraph", "3. 셋"), ("paragraph", "4. 넷")]
    assert md.parse_markdown_blocks("1. 하나\n2. 둘") == [("numbered", ["하나", "둘"])]
    assert md.parse_markdown_blocks("1. 하나\n\n2. 둘") == [("numbered", ["하나", "둘"])]   # 빈 줄로 띄운 같은 목록


# ---------------------------------------------------------------------------
# '---'는 구분선, 맨 앞 YAML 머리말은 넣지 않음
# ---------------------------------------------------------------------------


def test_front_matter_lines():
    assert md.front_matter_lines("---\ntitle: 보고서\ndate: 2026-11-03\n---\n# 제목") == 4
    assert md.front_matter_lines("\ufeff---\ntags:\n  - a\n...\n본문") == 4
    assert md.front_matter_lines("---\n본문 문단입니다\n---\n") == 0          # key: 모양이 아님 -> 구분선
    assert md.front_matter_lines("---\ntitle: x\n본문") == 0                  # 닫지 않음
    assert md.front_matter_lines("본문\n---\ntitle: x\n---") == 0             # 첫 줄이 아님
    blocks = md.parse_markdown_blocks("---\ntitle: x\n---\n# 제목\n\n---\n\n끝")
    assert blocks == [("front_matter", 3), ("heading", (1, "제목")), ("break", None), ("paragraph", "끝")]
    assert md.looks_like_markdown("---\ntitle: x\n---\n본문")


def test_table_cells_strip_bold_marks_and_report_whole_bold_cells():
    """검토 3장 끝 (1.5.1): 표 칸의 **굵게**가 워드·PPT 표에 별표 그대로 들어가던 것."""
    import _markdown as md
    rows, bold = md.table_cells([["항목", "값"], ["a", "1"], ["**합계**", "**3**"], ["부분 **굵게** 글", "x"]])
    assert rows[2] == ["합계", "3"] and rows[3][0] == "부분 굵게 글" and bold == [(2, 0), (2, 1)], (rows, bold)
