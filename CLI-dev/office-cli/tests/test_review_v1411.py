"""1.4.11 회귀 테스트: 2차 점검에서 찾은 오류 (A1~A6 reader·outlook, B1~B8 공통·PPT·엑셀).

항목마다 실패 상황을 재현하고 고친 동작을 확인합니다. 가짜 COM(fake_*.py)과 실제 파일(openpyxl, pymupdf 등)을 씁니다.
"""
import json
import os
import zipfile

import pytest

import _office as off
import _runtime as rt
import excel
import outlook
import ppt
import reader
import word
from fake_office import make_win32com
from fake_ppt import FakePresentation, FakeShape


def _run(module, capsys, *argv):
    with pytest.raises(SystemExit) as ei:
        module.main(list(argv))
    out = [ln for ln in capsys.readouterr().out.splitlines() if ln.strip()]
    assert len(out) == 1, out
    return ei.value.code, json.loads(out[0])


def _com(monkeypatch):
    fake, _ = make_win32com()
    monkeypatch.setattr(off, "win32com", fake)
    monkeypatch.setattr(off, "pythoncom", None)
    monkeypatch.setattr(off, "list_open", lambda kind: [])


# ---------------------------------------------------------------------------
# A1: 머리글·쪽 번호만 있는 쪽(빈 쪽·간지)은 '글자 없는 쪽'이 아님. 그림만 있는 쪽만, 조건 안의 쪽만
# ---------------------------------------------------------------------------

def _pixmap(pymupdf):
    pix = pymupdf.Pixmap(pymupdf.csRGB, pymupdf.IRect(0, 0, 20, 20), False)
    pix.clear_with(200)
    return pix


def _pdf(path, pages):
    """pages: 쪽마다 'text'(본문), 'header'(머리글·쪽 번호만), 'image'(그림만)."""
    import pymupdf
    doc = pymupdf.open()
    for i, kind in enumerate(pages, 1):
        page = doc.new_page()
        if kind in ("text", "header"):
            page.insert_text((72, 30), "SPEC DOCUMENT", fontname="helv", fontsize=9)
            page.insert_text((280, 820), f"- {i} -", fontname="helv", fontsize=9)
        if kind == "text":
            page.insert_text((72, 200), f"body text on page {i}", fontname="helv", fontsize=11)
        if kind == "image":
            page.insert_image(pymupdf.Rect(40, 40, 560, 800), pixmap=_pixmap(pymupdf))   # 쪽을 덮는 그림 (1.4.12)
    doc.save(str(path))
    doc.close()


def test_a1_header_only_page_is_not_flagged_and_search_stays_verified(tmp_path, capsys):
    p = tmp_path / "간지.pdf"
    _pdf(p, ["text", "header", "text", "text"])
    code, out = _run(reader, capsys, "read", "--file", str(p))
    assert code == 0 and "pages_without_text" not in out["meta"] and "pages_without_text" not in out
    assert out["agent_directive"].startswith("VERIFIED:") and "그림만" not in out["agent_directive"]
    code, out = _run(reader, capsys, "search", "--query", "warranty", "--file", str(p))
    assert code == 0 and out["total_matches"] == 0 and out["agent_directive"].startswith("VERIFIED: 0건"), out


def test_a1_scanned_page_counts_only_inside_pages_filter(tmp_path, capsys):
    p = tmp_path / "스캔.pdf"
    _pdf(p, ["text", "image", "text"])
    code, out = _run(reader, capsys, "search", "--query", "warranty", "--file", str(p), "--pages", "3")
    assert code == 0 and out["agent_directive"].startswith("VERIFIED: 0건"), out
    code, out = _run(reader, capsys, "search", "--query", "warranty", "--file", str(p), "--pages", "2-3")
    assert code == 0 and out["agent_directive"].startswith("UNVERIFIED:") and any("p.2" in g for g in out["not_searched"])
    code, out = _run(reader, capsys, "read", "--file", str(p), "--pages", "3")
    assert code == 0 and "pages_without_text" not in out and "notes" not in out
    code, out = _run(reader, capsys, "read", "--file", str(p), "--pages", "1-2")
    assert out["pages_without_text"] == [2] and "notes" in out and "truncated_reason" not in out    # A6: 잘림이 아님


def test_a1_section_filter_uses_section_of_preceding_text():
    doc = {"units": [{"page": 1, "section": ["1. 개요"]}, {"page": 3, "section": ["2. 시험"]}]}
    assert reader._pages_in_filter(doc, [2, 4], None, "") == [2, 4]
    assert reader._pages_in_filter(doc, [2, 4], {4}, "") == [4]
    assert reader._pages_in_filter(doc, [2, 4], None, "개요") == [2]
    assert reader._pages_in_filter(doc, [2, 4], None, "시험") == [4]


# ---------------------------------------------------------------------------
# A2: 이전 형식 캐시 - 상한을 넘은 것, stale 이중 집계, 원본 없는 것(clear --missing), directive 안내
# ---------------------------------------------------------------------------

def _old_caches(tmp_path, names, remove=()):
    index = {}
    for k, n in enumerate(names):
        src = tmp_path / n
        src.write_text(f"{n} 검증 문장", encoding="utf-8")
        st = os.stat(src)
        path = os.path.abspath(str(src))
        stamp = f"2026100{k}_000000"
        old = {"file": path, "type": "txt", "mtime": st.st_mtime, "size": st.st_size, "parser_version": 1,
               "parsed_at": stamp, "meta": {}, "units": [{"loc": "1", "heading": "", "text": "옛 검증 문장"}]}
        with open(reader._doc_cache_path(path), "w", encoding="utf-8") as f:
            json.dump(old, f, ensure_ascii=False)
        index[path] = {"type": "txt", "units": 1, "parsed_at": stamp}
        if n in remove:
            os.remove(src)
    with open(reader._index_path(), "w", encoding="utf-8") as f:
        json.dump(index, f, ensure_ascii=False)


def test_a2_old_caches_over_refresh_limit_have_their_own_reason(tmp_path, capsys, monkeypatch):
    monkeypatch.setattr(reader, "STALE_REFRESH_MAX", 1)
    _old_caches(tmp_path, ["a.txt", "b.txt"])
    code, out = _run(reader, capsys, "search", "--query", "검증")
    assert code == 0 and out["total_matches"] == 1                          # 1개는 다시 읽어 검색
    reasons = [s["reason"] for s in out["skipped"]]
    assert len(reasons) == 1 and "상한 1개" in reasons[0] and "다시 읽지 못함" not in reasons[0]
    assert "stale_cache_count" not in out                                   # 검색하지 않은 문서는 stale이 아님
    assert not any("옛 내용으로 검색" in g for g in out["not_searched"])
    assert out["agent_directive"].startswith("UNVERIFIED:") and "index --all" in out["agent_directive"]
    assert any(c["command"].endswith("index --all") for c in out["next_commands"])


def test_a2_old_cache_without_original_points_to_clear_missing(tmp_path, capsys):
    _old_caches(tmp_path, ["gone.txt"], remove=("gone.txt",))
    code, out = _run(reader, capsys, "search", "--query", "검증")
    assert code == 0 and "원본이 없어 다시 읽을 수 없음" in out["skipped"][0]["reason"]
    assert "stale_cache_count" not in out
    assert "clear --missing" in out["agent_directive"] and "index --all" not in out["agent_directive"]
    assert [c for c in out["next_commands"] if "clear --missing" in c["command"]]
    assert not [c for c in out["next_commands"] if "index --all" in c["command"]]


# ---------------------------------------------------------------------------
# A3: xlsx 셀 글자 상한에서 잘린 저장은 UNVERIFIED, 긴 글 저장 안내는 .md
# ---------------------------------------------------------------------------

def test_a3_xlsx_cell_cut_is_unverified(tmp_path, capsys, monkeypatch):
    monkeypatch.setattr(reader, "XLSX_CELL_MAX", 10)
    p = tmp_path / "긴글.txt"
    p.write_text("가" * 40 + "\n\n짧은 단락", encoding="utf-8")
    code, out = _run(reader, capsys, "read", "--file", str(p), "--save", str(tmp_path / "o.xlsx"))
    assert code == 0 and out["agent_directive"].startswith("UNVERIFIED:") and "셀 글자 상한 10자 초과: 1행" in out["agent_directive"]
    assert out["truncated"] is True and isinstance(out["truncated_reason"], list)
    code, out = _run(reader, capsys, "search", "--query", "가가", "--file", str(p), "--save", str(tmp_path / "s.xlsx"), "--full")
    assert code == 0 and out["agent_directive"].startswith("UNVERIFIED:") and "셀 글자 상한" in out["save_problem"]
    code, out = _run(reader, capsys, "read", "--file", str(p), "--save", str(tmp_path / "o.md"))
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:")


def test_a3_long_unit_next_command_saves_md(tmp_path, capsys):
    p = tmp_path / "긴글.txt"
    p.write_text("가" * 9000, encoding="utf-8")
    code, out = _run(reader, capsys, "read", "--file", str(p))
    assert code == 0 and out["units"][0]["text_truncated"] is True
    cmds = [c["command"] for c in out["next_commands"]]
    assert len(cmds) == 1 and "--save" in cmds[0] and ".md" in cmds[0] and ".md" in out["truncated_reason"][0]


# ---------------------------------------------------------------------------
# A6: truncated_reason은 늘 목록, 시트 행 상한만 잘렸으면 --save로 전체를 저장한다고 하지 않음
# ---------------------------------------------------------------------------

def test_a6_row_cap_only_does_not_offer_full_save(tmp_path, capsys, monkeypatch):
    import openpyxl
    monkeypatch.setattr(reader, "MAX_SHEET_ROWS", 5)
    p = tmp_path / "많은행.xlsx"
    wb = openpyxl.Workbook()
    for i in range(10):
        wb.active.append([f"항목 {i}", "값"])
    wb.save(str(p))
    code, out = _run(reader, capsys, "read", "--file", str(p))
    assert code == 0 and out["truncated"] is True and isinstance(out["truncated_reason"], list)
    assert "next_commands" not in out and "--save로 전체" not in out["agent_directive"]
    assert "5행" in out["agent_directive"] and "읽기 상한" in out["agent_directive"]
    code, out = _run(reader, capsys, "read", "--file", str(p), "--save", str(tmp_path / "o.csv"))
    assert code == 0 and isinstance(out["truncated_reason"], list) and out["agent_directive"].startswith("UNVERIFIED:")


# ---------------------------------------------------------------------------
# A4, A5: outlook current --save 확인, 폴더 이름에 ; 가 든 첨부 여러 개
# ---------------------------------------------------------------------------

class _Item:
    Subject = "회의 안내"
    Body = "본문"
    To = "a"
    CC = ""
    SenderName = "김"
    SenderEmailAddress = ""
    EntryID = "1"
    UnRead = False
    Attachments = type("A", (), {"Count": 0})()


def test_a4_current_save_problem_is_unverified(capsys, monkeypatch):
    monkeypatch.setattr(outlook, "get_outlook", lambda: object())
    monkeypatch.setattr(outlook, "current_item", lambda app: (_Item(), "열린 창"))
    code, out = _run(outlook, capsys, "current", "--save")
    assert code == 0 and out["agent_directive"].startswith("VERIFIED: 메일을 읽고 저장한"), out
    monkeypatch.setattr(rt, "check_saved_file", lambda path, rows=None, delimiter=None: "저장한 파일을 다시 열지 못함 (테스트)")
    code, out = _run(outlook, capsys, "current", "--save")
    assert code == 0 and out["agent_directive"].startswith("UNVERIFIED:") and out["save_problem"]


def test_a5_attach_paths_in_folder_with_semicolon(tmp_path):
    folder = tmp_path / "자료;2026"
    folder.mkdir()
    a, b = folder / "a.pdf", tmp_path / "b.pdf"
    a.write_bytes(b"1")
    b.write_bytes(b"2")
    assert outlook._attach_paths(f"{a};{b}") == [str(a), str(b)]
    assert outlook._attach_paths(f"{b};{a}") == [str(b), str(a)]
    with pytest.raises(rt.InputError):
        outlook._attach_paths(f"{a};{tmp_path / '없는.pdf'}")


# ---------------------------------------------------------------------------
# B1, B3: PPT 제목 확인은 제목 개체틀만, .md 첫 제목이 --title과 같으면 본문에 다시 넣지 않음
# ---------------------------------------------------------------------------

class _View:
    def GotoSlide(self, n):
        pass


@pytest.fixture
def pp(monkeypatch):
    holder = {"prs": FakePresentation(3)}
    win = type("W", (), {})()
    win.View = _View()
    app = type("A", (), {})()
    app.ActiveWindow = win
    _com(monkeypatch)
    monkeypatch.setattr(off, "save_snapshot", lambda *a, **k: None)

    def active(kind):
        win.Presentation = holder["prs"]
        return {"name": holder["prs"].Name, "path": holder["prs"].FullName, "app": app, "doc": holder["prs"]}
    monkeypatch.setattr(off, "get_active", active)
    return holder


def test_b1_slide_without_title_placeholder_does_not_ask(pp, capsys):
    prs = pp["prs"]
    s = prs.Slides.Add(4, 12)                                   # 빈 레이아웃: 제목 개체틀 없음
    s.Shapes.add(FakeShape("본문", 36, 100, 800, 300, text="본문 첫 줄"))
    code, out = _run(ppt, capsys, "write", "--slide", "4", "--title", "새 제목")
    assert code == 0 and out["status"] == "ok", out                # 이전: needs_user_decision (본문 첫 줄을 제목으로 봄)
    assert "새 제목" in ppt.slide_units(prs.Slides(4), 4)["text"] + ppt.slide_units(prs.Slides(4), 4)["title"]


def test_b3_same_heading_as_title_is_not_duplicated(pp, capsys, tmp_path):
    md = tmp_path / "내용.md"
    md.write_text("## 개요\n- 가\n", encoding="utf-8")
    code, out = _run(ppt, capsys, "write", "--slide", "new", "--title", "개요 ", "--file", str(md))
    assert code == 0, out
    u = ppt.slide_units(pp["prs"].Slides(4), 4)
    assert u["title"] == "개요" and "개요" not in u["text"] and "가" in u["text"]
    md.write_text("## 첫 절\n- 가\n", encoding="utf-8")
    code, out = _run(ppt, capsys, "write", "--slide", "new", "--title", "발표", "--file", str(md))
    assert "첫 절" in ppt.slide_units(pp["prs"].Slides(5), 5)["text"]


# ---------------------------------------------------------------------------
# B2: 오류 로그 가리기 - 짧은 이름(-q), 줄여 쓴 이름(--que), --map. 파서는 줄임을 받지 않음
# ---------------------------------------------------------------------------

def test_b2_redact_short_abbrev_and_map():
    red = rt.redact_argv(["search", "-q", "비밀", "-q비밀이", "--que", "기밀", "--bod=본문", "--map", '{"a":"b"}',
                          "--file", "C:\\a.txt", "-n", "5"], "reader")
    text = " ".join(red)
    assert "비밀" not in text and "기밀" not in text and "본문" not in text and '"a"' not in text
    assert "C:\\a.txt" in red and red[-2:] == ["-n", "5"]


def test_b2_short_aliases_of_content_options_are_all_mapped():
    import argparse
    for module in (excel, word, ppt, outlook, reader):
        stack = [module.build_cli()]
        while stack:
            p = stack.pop()
            for a in p._actions:
                if isinstance(a, argparse._SubParsersAction):
                    stack += list(a.choices.values())
                    continue
                longs = [o for o in a.option_strings if o.startswith("--")]
                if any(o in rt.REDACT_OPTIONS for o in longs):
                    for short in (o for o in a.option_strings if not o.startswith("--")):
                        assert rt.REDACT_SHORT.get(short) in longs, (module.__name__, short)
            assert p.allow_abbrev is False, module.__name__


def test_b2_abbreviated_option_is_rejected_and_log_is_redacted(capsys):
    code, out = _run(reader, capsys, "search", "--que", "인사평가 기밀", "--file", "a.txt")
    assert code == 2 and out["status"] == "error"
    log = open(rt.error_log_path(), encoding="utf-8").read()
    assert "인사평가" not in log and "--que" in log


# ---------------------------------------------------------------------------
# B4: excel/word/ppt read --save 저장 파일을 다시 열어 확인 (빈 행은 세지 않음)
# ---------------------------------------------------------------------------

def test_b4_excel_read_save_counts_non_blank_rows(tmp_path, capsys, monkeypatch):
    import openpyxl
    p = tmp_path / "표.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"], ws["A2"], ws["A4"] = "머리", 1, 3                  # 3행은 빈 행
    wb.save(str(p))
    code, out = _run(excel, capsys, "read", "--target", str(p), "--save", str(tmp_path / "o.csv"))
    assert code == 0 and out["agent_directive"].startswith("VERIFIED: 저장한 CSV를 다시 열어"), out
    monkeypatch.setattr(rt, "check_saved_file", lambda path, rows=None, delimiter=None: "저장한 파일을 다시 열지 못함 (테스트)")
    code, out = _run(excel, capsys, "read", "--target", str(p), "--save", str(tmp_path / "o2.csv"))
    assert code == 0 and out["agent_directive"].startswith("UNVERIFIED:")


def test_b4_word_and_ppt_read_save_is_reopened(tmp_path, capsys, monkeypatch):
    import docx
    import pptx
    d = tmp_path / "문서.docx"
    doc = docx.Document()
    doc.add_paragraph("첫 단락")
    doc.save(str(d))
    s = tmp_path / "발표.pptx"
    prs = pptx.Presentation()
    prs.slides.add_slide(prs.slide_layouts[0]).shapes.title.text = "제목"
    prs.save(str(s))
    for module, target in ((word, d), (ppt, s)):
        code, out = _run(module, capsys, "read", "--target", str(target), "--save", str(tmp_path / f"{target.stem}.md"))
        assert code == 0 and out["agent_directive"].startswith("VERIFIED: 저장한 파일을 다시 열어"), out
    monkeypatch.setattr(rt, "check_saved_file", lambda path, rows=None, delimiter=None: "저장한 파일을 다시 열지 못함 (테스트)")
    for module, target in ((word, d), (ppt, s)):
        code, out = _run(module, capsys, "read", "--target", str(target), "--save", str(tmp_path / f"{target.stem}2.md"))
        assert code == 0 and out["agent_directive"].startswith("UNVERIFIED:"), out


# ---------------------------------------------------------------------------
# B5: pptm은 python-pptx로 다시 열어 확인
# ---------------------------------------------------------------------------

def test_b5_pptm_is_reopened(tmp_path):
    import pptx
    src = tmp_path / "a.pptx"
    pptx.Presentation().save(str(src))
    dst = tmp_path / "a.pptm"
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w") as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename == "[Content_Types].xml":
                data = data.replace(b"application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
                                    b"application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml")
            zout.writestr(info, data)
    assert rt.check_saved_file(str(dst)) is None
    (tmp_path / "b.pptm").write_bytes(b"broken")
    assert "다시 열지 못함" in rt.check_saved_file(str(tmp_path / "b.pptm"))


# ---------------------------------------------------------------------------
# B6, B7: 환경 변수·~ 가 든 전체 경로는 새 파일로 만들 수 있음, 전체 경로가 없으면 열린 문서를 [후보]로 주지 않음
# ---------------------------------------------------------------------------

class _Created(Exception):
    pass


def test_b6_expanded_absolute_path_can_be_created(monkeypatch, tmp_path):
    _com(monkeypatch)
    monkeypatch.setenv("OCLI_TEST_HOME", str(tmp_path))
    monkeypatch.setenv("HOME", str(tmp_path))
    monkeypatch.setattr(off, "match_open", lambda kind, target, docs=None: None)

    def new_instance(kind):
        raise _Created()
    monkeypatch.setattr(off, "_new_instance", new_instance)
    for target in ("$OCLI_TEST_HOME/새.xlsx", "~/새.xlsx"):
        with pytest.raises(_Created):                     # '찾을 수 없음'이 아니라 새 파일 만들기로 진행
            with off.open_target("excel", target, readonly=False, create=True):
                pass


def test_b7_missing_absolute_path_is_not_ambiguous(monkeypatch, tmp_path):
    _com(monkeypatch)
    monkeypatch.setattr(off, "list_open", lambda kind: [{"name": "보고.xlsx", "path": "C:\\docs\\보고.xlsx"},
                                                        {"name": "예산.xlsx", "path": "C:\\docs\\예산.xlsx"}])
    monkeypatch.setattr(off, "match_open", lambda kind, target, docs=None: None)
    with pytest.raises(rt.InputError) as ei:
        with off.open_target("excel", str(tmp_path / "없음.xlsx"), readonly=True):
            pass
    assert "[후보]" not in str(ei.value) and ei.value.candidates != ["보고.xlsx", "예산.xlsx"]
    with pytest.raises(rt.InputError) as ei:                 # 파일 이름만이면 그대로 [후보]
        with off.open_target("excel", "없음.xlsx", readonly=True):
            pass
    assert sorted(ei.value.candidates) == ["보고.xlsx", "예산.xlsx"]


# ---------------------------------------------------------------------------
# B8: 닫힌 xlsx의 작은 범위 읽기는 시트 전체를 훑지 않음, Sheet!A1의 시트로 바꾼 뒤에 셈
# ---------------------------------------------------------------------------

def test_b8_closed_bounds_only_when_needed(tmp_path, monkeypatch):
    import openpyxl
    p = tmp_path / "b.xlsx"
    wb = openpyxl.Workbook()
    wb.active.title = "첫"
    wb.active["A1"] = 1
    s2 = wb.create_sheet("둘")
    for r in range(1, 6):
        s2.cell(r, 2, r)
    wb.save(str(p))
    calls = []
    real = excel._closed_bounds

    def counting(ws):
        calls.append(ws.title)
        return real(ws)
    monkeypatch.setattr(excel, "_closed_bounds", counting)
    excel._read_matrix_closed(str(p), None, "A1:B2")
    assert calls == []
    sheet, r0, c0, rows, _ = excel._read_matrix_closed(str(p), None, "'둘'!B:B")
    assert calls == ["둘"] and sheet == "둘" and len(rows) == 5
    excel._read_matrix_closed(str(p), None, None)
    assert calls == ["둘", "첫"]
