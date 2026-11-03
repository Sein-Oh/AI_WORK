"""1.4.26 (17차 점검) 회귀 테스트.

H1: ppt _verify_slides가 기대 글에서만 ** __를 빼 코드 줄(def __init__, x**2)이나 제목(## __init__ 설명)이 첫 글이면 UNVERIFIED였음 (1.4.25 회귀).
H2: word _verify_inserted가 '---'로 시작하는 .md에서 '--'를 찾다 UNVERIFIED, 첫 표 행의 '\\|'를 잘못 다뤘음.
H3: ppt write --text의 펜스(```) 줄이 글자 그대로 들어갔음 (.md와 다름).
H4: ppt 3단계 이하 제목(### **x**)에 ** 가 그대로 보였음.
H5: excel load_table_file(.md)이 모든 '|' 줄(두 번째 표, 코드 블록)을 모으고 '-' 2개 정렬 줄을 남기고 '\\|'에서 칸을 나눴음.
B1: 0건 결과(office windows, undo --list, excel·word·ppt info --all, excel chart --action list)에 count와 '0건' 안내가 없었음.
B2: reader clear --dry-run이 20개를 넘으면 잘랐는데 full_result_file이 없었음.
B3: 오류 로그에서 --section(문서 제목 글)과 excel --highlight(셀 값)를 가리지 않았음.
"""
import json
import os

import pytest

import _office as off
import _runtime as rt
import excel
import office
import reader
import test_excel
import test_ppt
import test_word

xl = test_excel.xl
pp = test_ppt.pp
wd = test_word.wd


def _run(mod, capsys, *argv):
    with pytest.raises(SystemExit) as ei:
        mod.main(list(argv))
    out = [ln for ln in capsys.readouterr().out.splitlines() if ln.strip()]
    assert len(out) == 1, out
    return ei.value.code, json.loads(out[0])


def _ppt_md(capsys, tmp_path, md):
    p = tmp_path / "s.md"
    p.write_text(md, encoding="utf-8")
    return test_ppt.run(capsys, "write", "--file", str(p))


def _texts(pp):
    prs = pp["prs"]
    s = prs.Slides(prs.Slides.Count)
    return [sh.TextFrame.TextRange.Text for sh in test_ppt.live(s) if hasattr(sh, "TextFrame")]


# H1 ppt 다시 읽기 -------------------------------------------------------------------------------------------

@pytest.mark.parametrize("md", [
    "## 코드\n\n```python\ndef __init__(self):\n    pass\n```\n",
    "## 수식\n\n```\nx**2 + y**2\n```\n",
    "## __init__ 설명\n\n본문\n",
])
def test_ppt_code_and_title_with_markers_verified(pp, capsys, tmp_path, md):
    code, out = _ppt_md(capsys, tmp_path, md)
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out


def test_ppt_title_option_with_dunder_verified(pp, capsys):
    code, out = test_ppt.run(capsys, "write", "--slide", "new", "--title", "__main__ 실행", "--text", "본문")
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out


# H3 ppt --text 펜스 --------------------------------------------------------------------------------------------

def test_ppt_text_fence_verbatim(pp, capsys):
    code, out = test_ppt.run(capsys, "write", "--slide", "new", "--text", "설명\\n```\\nx**2 = __a__\\n```\\n- 끝")
    texts = "\n".join(_texts(pp))
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    assert "x**2 = __a__" in texts and "```" not in texts, texts


# H4 ppt 3단계 제목 굵게 표시 --------------------------------------------------------------------------------------

def test_ppt_level3_heading_bold_markers_removed(pp, capsys, tmp_path):
    code, out = _ppt_md(capsys, tmp_path, "## 슬라이드\n\n### **핵심** 요약\n\n본문\n")
    texts = "\n".join(_texts(pp))
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    assert "핵심 요약" in texts and "**" not in texts, texts


# H2 word 다시 읽기 ------------------------------------------------------------------------------------------

@pytest.mark.parametrize("md", [
    "---\n\n# 제목\n본문\n",
    "***\n\n본문\n",
    "---\ntitle: 보고서\n---\n\n본문\n",
    "| a\\|b | c |\n|---|---|\n| 1 | 2 |\n",
])
def test_word_md_first_block_verified(wd, capsys, tmp_path, md):
    p = tmp_path / "w.md"
    p.write_text(md, encoding="utf-8")
    code, out = test_word.run(capsys, "write", "--file", str(p))
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out


# H5 excel 마크다운 표 -------------------------------------------------------------------------------------------

def test_excel_load_table_file_first_markdown_table(tmp_path):
    p = tmp_path / "t.md"
    p.write_text("# 보고\n\n| 항목 | 값 |\n|--|--|\n| a | 1 |\n| b\\|c | 2 |\n\n본문\n\n"
                 "| 다른 | 표 |\n|---|---|\n| x | y |\n\n```\n| code | line |\n```\n", encoding="utf-8")
    assert excel.load_table_file(str(p)) == [["항목", "값"], ["a", 1], ["b|c", 2]]


def test_excel_load_table_file_pipe_only_in_code_is_lines(tmp_path):
    p = tmp_path / "c.md"
    p.write_text("```\n| a | b |\n```\n", encoding="utf-8")
    assert excel.load_table_file(str(p)) == [["```"], ["| a | b |"], ["```"]]


# B1 0건 결과 ---------------------------------------------------------------------------------------------

def test_office_windows_zero(monkeypatch, capsys):
    monkeypatch.setattr(off, "open_documents_summary", lambda: [])
    monkeypatch.setattr(off, "running_office_caps", lambda: {})
    code, out = _run(office, capsys, "windows")
    assert code == 0 and out["count"] == 0 and out["agent_directive"].startswith("VERIFIED: 0건"), out


def test_undo_list_zero(capsys):
    code, out = _run(office, capsys, "undo", "--list")
    assert code == 0 and out["count"] == 0 and out["agent_directive"].startswith("VERIFIED: 0건"), out


def test_excel_info_all_zero(xl, capsys):
    code, out = _run(excel, capsys, "info", "--all")
    assert code == 0 and out["count"] == 0 and out["agent_directive"].startswith("VERIFIED: 0건"), out


def test_excel_chart_list_zero(xl, capsys):
    code, out = _run(excel, capsys, "chart", "--action", "list")
    assert code == 0 and out["count"] == 0 and out["agent_directive"].startswith("VERIFIED: 0건"), out


# B2 reader clear 미리보기 전체 목록 ---------------------------------------------------------------------------------

def test_reader_clear_dry_run_truncated_has_full_result_file(tmp_path, capsys):
    docs = tmp_path / "docs"
    docs.mkdir()
    for i in range(reader.MAX_LIST_ITEMS + 3):
        (docs / f"문서{i:02d}.txt").write_text(f"내용 {i}", encoding="utf-8")
        code, out = _run(reader, capsys, "index", "--file", str(docs / f"문서{i:02d}.txt"))
        assert code == 0, out
    code, out = _run(reader, capsys, "clear", "--all", "--dry-run")
    assert code == 0 and out["truncated"] is True and out["count"] == reader.MAX_LIST_ITEMS + 3, out
    path = out["full_result_file"]
    assert os.path.isfile(path)
    rows = [r for r in rt.read_csv_rows(path) if any(r)]
    assert len(rows) == reader.MAX_LIST_ITEMS + 4 and rows[0] == ["file", "path", "missing"], rows[:2]
    assert len(reader.load_index()) == reader.MAX_LIST_ITEMS + 3          # 미리보기는 지우지 않음


# B3 오류 로그 가리기 ---------------------------------------------------------------------------------------------

def test_redact_section_and_excel_highlight():
    red = rt.redact_argv(["read", "--file", "C:/a/보고.docx", "--section", "인수 계획"], "word")
    assert red == ["read", "--file", "C:/a/보고.docx", "--section", "<5자 생략>"], red
    red = rt.redact_argv(["search", "-q", "x", "--section=기밀"], "reader")
    assert red[-1] == "--section=<2자 생략>", red
    red = rt.redact_argv(["style", "--target", "보고.xlsx", "--sheet", "Sheet1", "--highlight", "=완료:lightgreen"], "excel")
    assert red == ["style", "--target", "보고.xlsx", "--sheet", "Sheet1", "--highlight", "<14자 생략>"], red
    assert rt.redact_argv(["format", "--highlight", "yellow"], "word") == ["format", "--highlight", "yellow"]


@pytest.mark.parametrize("mod,key", [("word", "documents"), ("ppt", "presentations")])
def test_word_ppt_info_all_zero(monkeypatch, capsys, mod, key):
    import ppt
    import word
    monkeypatch.setattr(off, "list_open", lambda kind: [])
    code, out = _run({"word": word, "ppt": ppt}[mod], capsys, "info", "--all")
    assert code == 0 and out["count"] == 0 and out[key] == [] and out["agent_directive"].startswith("VERIFIED: 0건"), out
