"""1.4.12 회귀 테스트: 3차 점검에서 찾은 오류 (R1~R6 reader, D1~D4 엑셀·워드·PPT·공통, L1~L2 저장 확인·닫힌 xlsx).

항목마다 실패 상황을 재현하고 고친 동작을 확인합니다. 가짜 COM(fake_*.py)과 실제 파일(openpyxl, pymupdf 등)을 씁니다.
"""
import json
import os
import subprocess
import sys
import time

import pytest

import _office as off
import _runtime as rt
import excel
import ppt
import reader
import word
from fake_excel import FakeExcelWorkbook, FakeXlApp
from fake_office import make_win32com
from fake_ppt import FakePresentation

TOOL_DIR = os.path.dirname(os.path.abspath(reader.__file__))


def _run(module, capsys, *argv):
    rt.RECOVERED.clear()
    with pytest.raises(SystemExit) as ei:
        module.main(list(argv))
    out = [ln for ln in capsys.readouterr().out.splitlines() if ln.strip()]
    assert len(out) == 1, out
    return ei.value.code, json.loads(out[0])


def _index_json():
    with open(os.path.join(os.environ["OFFICE_CLI_CACHE_DIR"], "index.json"), encoding="utf-8") as f:
        return json.load(f)


@pytest.fixture
def xl(monkeypatch):
    app = FakeXlApp()
    wb = FakeExcelWorkbook(app, "보고.xlsx", sheets=("Sheet1", "요약"))
    fake, _ = make_win32com()
    monkeypatch.setattr(off, "win32com", fake)
    monkeypatch.setattr(off, "pythoncom", None)
    monkeypatch.setattr(off, "list_open", lambda kind: [])
    monkeypatch.setattr(off, "get_active", lambda kind: {"name": wb.Name, "path": wb.FullName, "app": app, "doc": wb})
    return app, wb


# ---------------------------------------------------------------------------
# R1: 동시에 index를 실행해도 색인 항목이 사라지지 않음 (잠금 + 끝에 다시 읽어 채움), index --all이 빠진 캐시를 찾음
# ---------------------------------------------------------------------------

def test_r1_concurrent_index_runs_keep_all_entries(tmp_path):
    folders = []
    for name in ("a", "b"):
        d = tmp_path / name
        d.mkdir()
        for i in range(30):
            (d / f"{name}{i}.txt").write_text(f"{name} 문서 {i} 내용", encoding="utf-8")
        folders.append(str(d))
    env = dict(os.environ, PYTHONIOENCODING="utf-8")
    procs = [subprocess.Popen([sys.executable, os.path.join(TOOL_DIR, "reader.py"), "index", "--folder", f],
                              stdout=subprocess.PIPE, stderr=subprocess.DEVNULL, env=env) for f in folders]
    outs = [json.loads(p.communicate(timeout=300)[0].decode("utf-8").strip().splitlines()[-1]) for p in procs]
    assert all(o["status"] == "ok" for o in outs), outs
    index = _index_json()
    assert len(index) == 60, len(index)
    assert not os.path.exists(os.path.join(os.environ["OFFICE_CLI_CACHE_DIR"], "index.lock"))


def test_r1_index_repairs_entries_lost_to_another_run(tmp_path, capsys, monkeypatch):
    for i in range(3):
        (tmp_path / f"d{i}.txt").write_text(f"내용 {i}", encoding="utf-8")
    real = reader._index_files

    def racing(files, max_files):
        res = real(files, max_files)
        # 다른 실행이 옛 색인으로 덮어쓴 상황: 방금 넣은 항목 하나가 사라짐
        index = reader.load_index()
        index.pop(sorted(index)[0])
        reader._save_index(index)
        return res
    monkeypatch.setattr(reader, "_index_files", racing)
    code, out = _run(reader, capsys, "index", "--folder", str(tmp_path))
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    assert any("다시 넣음" in r for r in out["recovered"])
    assert len(_index_json()) == 3


def test_r1_index_all_readds_cache_files_missing_from_index(tmp_path, capsys):
    keep = tmp_path / "keep.txt"
    keep.write_text("남길 문서", encoding="utf-8")
    gone = tmp_path / "gone.txt"
    gone.write_text("지울 문서", encoding="utf-8")
    other = tmp_path / "other.txt"
    other.write_text("다른 문서", encoding="utf-8")
    for p in (keep, gone, other):
        _run(reader, capsys, "read", "--file", str(p))
    index = _index_json()
    for p in (keep, gone):                                   # 색인에서만 빠진 캐시 (중단·동시 실행)
        index.pop(next(k for k in index if os.path.basename(k) == p.name))
    with open(os.path.join(os.environ["OFFICE_CLI_CACHE_DIR"], "index.json"), "w", encoding="utf-8") as f:
        json.dump(index, f)
    gone.unlink()                                            # 원본이 없는 캐시는 다시 넣지 않음
    code, out = _run(reader, capsys, "index", "--all")
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    assert out["readded_to_index"] == 1 and any("1개 -> 색인에 다시 넣음" in r for r in out["recovered"])
    assert sorted(os.path.basename(k) for k in _index_json()) == ["keep.txt", "other.txt"]


def test_r1_index_lock_waits_and_clears_stale_lock(monkeypatch):
    lock = os.path.join(os.environ["OFFICE_CLI_CACHE_DIR"], "index.lock")
    with open(lock, "w") as f:
        f.write("1")
    old = time.time() - reader.INDEX_LOCK_STALE - 5
    os.utime(lock, (old, old))                               # 중단된 실행이 남긴 잠금
    with reader._index_lock():
        assert os.path.exists(lock)
    assert not os.path.exists(lock)
    with open(lock, "w") as f:                               # 살아 있는 잠금: 상한까지 기다린 뒤 잠금 없이 진행 (무한 대기 없음)
        f.write("1")
    monkeypatch.setattr(reader, "INDEX_LOCK_WAIT", 0.2)
    start = time.monotonic()
    with reader._index_lock():
        pass
    assert time.monotonic() - start < 5 and os.path.exists(lock)


# ---------------------------------------------------------------------------
# R2: 글자 없는 쪽 = 머리글·쪽 번호를 뺀 뒤 글이 없고, 그림이 쪽의 큰 부분을 덮는 쪽
# ---------------------------------------------------------------------------

def _pix(pymupdf, w=200, h=280):
    pix = pymupdf.Pixmap(pymupdf.csRGB, pymupdf.IRect(0, 0, w, h), False)
    pix.clear_with(200)
    return pix


def test_r2_scanned_page_with_bates_footer_is_flagged(tmp_path, capsys):
    import pymupdf
    doc = pymupdf.open()
    for i in range(1, 5):
        p = doc.new_page()
        p.insert_text((280, 820), f"- {i} -", fontsize=9)                       # 스캔 쪽에도 찍힌 쪽 번호
        if i == 3:
            p.insert_image(pymupdf.Rect(50, 60, 550, 780), pixmap=_pix(pymupdf))
        else:
            p.insert_text((72, 200), f"Body text page {i} torque {i * 10} Nm.", fontsize=11)
    path = tmp_path / "bates.pdf"
    doc.save(str(path))
    code, out = _run(reader, capsys, "read", "--file", str(path))
    assert code == 0 and out["meta"]["pages_without_text"] == [3]


def test_r2_blank_page_with_recurring_logo_is_not_flagged(tmp_path, capsys):
    import pymupdf
    doc = pymupdf.open()
    for i in range(1, 5):
        p = doc.new_page()
        p.insert_image(pymupdf.Rect(500, 20, 560, 40), pixmap=_pix(pymupdf, 20, 10))   # 모든 쪽의 작은 로고
        if i != 3:
            p.insert_text((72, 200), f"Body text page {i} torque {i * 10} Nm.", fontsize=11)
    path = tmp_path / "logo.pdf"
    doc.save(str(path))
    code, out = _run(reader, capsys, "read", "--file", str(path))
    assert code == 0 and "pages_without_text" not in out["meta"]
    code, out = _run(reader, capsys, "search", "--query", "torque")
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out


def test_r2_ocr_text_layer_over_image_is_not_flagged(tmp_path, capsys):
    import pymupdf
    doc = pymupdf.open()
    for i in range(1, 3):
        p = doc.new_page()
        p.insert_image(p.rect, pixmap=_pix(pymupdf))
        p.insert_text((72, 200), f"OCR layer page {i}", fontsize=11, render_mode=3)     # 보이지 않는 OCR 글자 층
    path = tmp_path / "ocr.pdf"
    doc.save(str(path))
    code, out = _run(reader, capsys, "read", "--file", str(path))
    assert code == 0 and "pages_without_text" not in out["meta"] and not out["meta"].get("ocr_needed")
    assert reader.PARSER_VERSION == 7


# ---------------------------------------------------------------------------
# R3: 캐시도 원본도 없는 문서는 index --all 반복 안내가 아니라 clear --missing
# ---------------------------------------------------------------------------

def test_r3_missing_cache_and_original_is_orphan(tmp_path, capsys):
    a = tmp_path / "a.txt"
    a.write_text("검색어 있음", encoding="utf-8")
    b = tmp_path / "b.txt"
    b.write_text("다른 글", encoding="utf-8")
    for p in (a, b):
        _run(reader, capsys, "read", "--file", str(p))
    os.remove(reader._doc_cache_path(str(b)))
    b.unlink()
    code, out = _run(reader, capsys, "search", "--query", "검색어")
    assert code == 0 and out["missing_file_count"] == 1
    assert "clear --missing" in out["skipped"][0]["reason"]
    cmds = [c["command"] for c in out["next_commands"]]
    assert any("clear --missing --dry-run" in c for c in cmds) and not any("index --all" in c for c in cmds)
    assert "clear --missing" in out["agent_directive"] and "index --all" not in out["agent_directive"]


# ---------------------------------------------------------------------------
# R4: search --save --full은 단락 전체 저장, 아니면 잘린 행을 알림. JSON --full의 잘림 표시
# ---------------------------------------------------------------------------

def test_r4_search_save_full_and_clamped_context(tmp_path, capsys):
    import openpyxl
    p = tmp_path / "long.txt"
    p.write_text("시작 검색어 " + "가" * 5000, encoding="utf-8")
    code, out = _run(reader, capsys, "search", "--query", "검색어", "--file", str(p), "--save", str(tmp_path / "f.xlsx"), "--full")
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    ws = openpyxl.load_workbook(out["output_file"]).active
    assert len(ws.cell(2, 7).value) > 5000 and "context_clamped" not in out
    code, out = _run(reader, capsys, "search", "--query", "검색어", "--save", str(tmp_path / "c.csv"), "--full")
    rows = rt.read_csv_rows(out["output_file"], ",")
    assert code == 0 and len(rows[1][6]) > 5000
    code, out = _run(reader, capsys, "search", "--query", "검색어", "--save", str(tmp_path / "s.xlsx"))
    assert code == 0 and out["context_clamped"] == 1 and "3000자" in out["agent_directive"]
    code, out = _run(reader, capsys, "search", "--query", "검색어", "--full")
    m = out["matches"][0]
    assert m["text_truncated"] and m["text_chars"] > 5000 and len(m["text"]) == 2000
    assert out["truncated"] is True and "--full" in out["next_commands"][-1]["command"]


# ---------------------------------------------------------------------------
# R5: --top, --context 상한
# ---------------------------------------------------------------------------

def test_r5_top_and_context_are_clamped(tmp_path, capsys):
    p = tmp_path / "many.txt"
    p.write_text("\n\n".join(f"항목 {i} 검색어 " + "나" * 2000 for i in range(80)), encoding="utf-8")
    code, out = _run(reader, capsys, "search", "--query", "검색어", "--file", str(p), "--top", "500", "--context", "100000")
    assert code == 0 and out["shown"] == 50 and out["truncated"] is True
    assert len(out["warnings"]) == 2 and all(len(m["snippet"]) < 1200 for m in out["matches"])
    assert any("--save" in c["command"] for c in out["next_commands"])


# ---------------------------------------------------------------------------
# R6: clear --file/--folder의 ~
# ---------------------------------------------------------------------------

def test_r6_clear_expands_home(tmp_path, capsys, monkeypatch):
    monkeypatch.setenv("HOME", str(tmp_path))
    monkeypatch.setenv("USERPROFILE", str(tmp_path))
    p = tmp_path / "docs" / "a.txt"
    p.parent.mkdir()
    p.write_text("내용", encoding="utf-8")
    _run(reader, capsys, "read", "--file", str(p))
    code, out = _run(reader, capsys, "clear", "--file", "~/docs/a.txt", "--dry-run")
    assert code == 0 and out["count"] == 1
    code, out = _run(reader, capsys, "clear", "--folder", "~/docs")
    assert code == 0 and out["count"] == 1 and out["agent_directive"].startswith("VERIFIED:")


# ---------------------------------------------------------------------------
# D1: --new-sheet가 없는 명령의 없는 시트는 '찾을 수 없음' ([후보]/[추천]), --new-sheet 안내 없음
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("argv", [
    ["style", "--sheet", "요약표", "--range", "A1:B2", "--bold"],
    ["replace", "--sheet", "요약표", "--find", "a", "--replace", "b"],
    ["chart", "--sheet", "요약표", "--range", "A1:B3"],
    ["range", "--action", "width", "--sheet", "요약표", "--range", "A1", "--size", "20"],
])
def test_d1_missing_sheet_without_new_sheet_option(xl, capsys, argv):
    code, out = _run(excel, capsys, *argv)
    assert code == 1 and out["error_type"] != "needs_user_decision", out
    assert "--new-sheet" not in out["agent_directive"] and "--new-sheet" not in out["message"]
    assert "[후보] 요약" in out["message"]


def test_d1_excel_write_still_asks_about_new_sheet(xl, capsys):
    code, out = _run(excel, capsys, "write", "--sheet", "요약표", "--range", "A1", "--values", '[["a"]]')
    assert code == 1 and out["error_type"] == "needs_user_decision" and "--new-sheet" in out["agent_directive"]


# ---------------------------------------------------------------------------
# D2: 쓰기 전에 멈춘 실행은 되돌리기 지점을 남기지 않음
# ---------------------------------------------------------------------------

def test_d2_refused_write_leaves_no_undo_point(xl, capsys):
    code, out = _run(excel, capsys, "write", "--range", "A1", "--values", '[["a", 1]]')
    assert code == 0
    snaps = off.list_snapshots("excel")
    assert len(snaps) == 1
    first = snaps[0]["meta_file"]
    code, out = _run(excel, capsys, "write", "--range", "A1", "--values", '[["b", 2]]')     # 덮어쓰기 확인
    assert code == 1 and out["error_type"] == "needs_user_decision"
    snaps = off.list_snapshots("excel")
    assert len(snaps) == 1 and snaps[0]["meta_file"] == first   # undo는 첫 쓰기 전 상태로 되돌림
    code, out = _run(excel, capsys, "write", "--sheet", "없는시트", "--range", "A1", "--values", '[["c"]]')
    assert code == 1 and len(off.list_snapshots("excel")) == 1


# ---------------------------------------------------------------------------
# D3: 출력 크기 상한 (긴 칸, 찾은 값, 슬라이드 글, 워드 단락, 노트)
# ---------------------------------------------------------------------------

def test_d3_clamp_rows_caps_cells_and_first_row():
    rows = [["가" * 30000 for _ in range(40)], [1, 2]]
    shown, rows_cut, cells_cut = rt.clamp_rows(rows)
    assert sum(len(str(v)) for r in shown for v in r) <= rt.DEFAULT_MAX_CHARS + 100
    assert rows_cut and cells_cut > 0
    shown, rows_cut, cells_cut = rt.clamp_rows([["짧은 값", 1], ["둘", 2]])
    assert shown == [["짧은 값", 1], ["둘", 2]] and not rows_cut and cells_cut == 0


def test_d3_excel_read_and_search_cap_long_values(tmp_path, capsys):
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"], ws["B1"], ws["A2"] = "머리", "찾기 " + "다" * 3000, 1
    p = tmp_path / "long.xlsx"
    wb.save(str(p))
    code, out = _run(excel, capsys, "read", "--target", str(p))
    assert code == 0 and out["truncated"] is True and len(out["values"][0][1]) == rt.MAX_CELL_CHARS
    assert out["next_commands"][-1]["command"].endswith("--save")
    code, out = _run(excel, capsys, "read", "--target", str(p), "--search", "찾기")
    hit = out["matches"][0]
    assert code == 0 and hit["value_truncated"] and len(hit["value"]) == rt.MAX_CELL_CHARS and out["truncated"] is True
    assert "--save" in out["next_commands"][0]["command"]
    code, out = _run(excel, capsys, "read", "--target", str(p), "--range", "A1:A2")
    assert code == 0 and out["truncated"] is False and "next_commands" not in out


def test_d3_ppt_and_word_long_text_marked_truncated(tmp_path, capsys):
    import docx
    import pptx
    s = tmp_path / "deck.pptx"
    prs = pptx.Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "제목"
    slide.placeholders[1].text = "라" * 20000
    prs.save(str(s))
    code, out = _run(ppt, capsys, "read", "--target", str(s))
    sl = out["slides"][0]
    assert code == 0 and out["truncated"] is True and sl["text_truncated"] and len(sl["text"]) == rt.DEFAULT_MAX_CHARS
    assert out["next_commands"][-1]["command"].endswith("--save")
    d = tmp_path / "doc.docx"
    doc = docx.Document()
    doc.add_paragraph("마" * 20000)
    doc.save(str(d))
    code, out = _run(word, capsys, "read", "--target", str(d))
    para = out["paragraphs"][0]
    assert code == 0 and out["truncated"] is True and para["text_truncated"] and para["text_chars"] == 20000
    assert out["next_commands"][0]["command"].endswith("--save")


def test_d3_ppt_notes_cut_sets_truncated(monkeypatch, capsys):
    from test_ppt import FakePptApp
    prs = FakePresentation(2)
    app = FakePptApp(prs)
    fake, _ = make_win32com()
    monkeypatch.setattr(off, "win32com", fake)
    monkeypatch.setattr(off, "pythoncom", None)
    monkeypatch.setattr(off, "list_open", lambda kind: [])
    monkeypatch.setattr(off, "get_active", lambda kind: {"name": prs.Name, "path": prs.FullName, "app": app, "doc": prs})
    prs.Slides(1).NotesPage.body.TextFrame.TextRange.Text = "바" * 9000
    code, out = _run(ppt, capsys, "notes")
    assert code == 0 and out["truncated"] is True and out["notes"][0]["text_truncated"] and "truncated_reason" in out


# ---------------------------------------------------------------------------
# D4: 엑셀 범위 밖 칸 주소
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("ref", ["A0", "A1048577", "XFE1"])
def test_d4_parse_cell_rejects_out_of_range(ref):
    with pytest.raises(rt.InputError):
        excel.parse_cell(ref)


def test_d4_parse_area_rejects_row_zero(tmp_path, capsys):
    with pytest.raises(rt.InputError):
        excel.parse_area("0:3", None, None)
    assert excel.parse_cell("XFD1048576") == (1048576, 16384)
    import openpyxl
    wb = openpyxl.Workbook()
    wb.active["A1"] = 1
    p = tmp_path / "x.xlsx"
    wb.save(str(p))
    code, out = _run(excel, capsys, "read", "--target", str(p), "--range", "A0:B2")
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:")


# ---------------------------------------------------------------------------
# L1: 도구가 쓴 CSV는 ','로 다시 읽음 ('| ' 같은 칸 때문에 구분자를 잘못 추측하지 않게)
# ---------------------------------------------------------------------------

def test_l1_csv_with_pipe_cells_is_verified(tmp_path, capsys):
    import openpyxl
    wb = openpyxl.Workbook()
    wb.active["A1"] = " | "                       # 구분자 추측이 '|'를 골라 빈 행으로 읽던 칸
    src = tmp_path / "pipe.xlsx"
    wb.save(str(src))
    code, out = _run(excel, capsys, "read", "--target", str(src), "--save", str(tmp_path / "o.csv"))
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    p = rt.write_csv_rows(str(tmp_path / "w.csv"), [[" | "]])
    assert rt.check_saved_file(p, 1, delimiter=",") is None


# ---------------------------------------------------------------------------
# L2: 닫힌 xlsx의 명시 범위는 끝쪽 빈 행까지 (열린 문서와 같은 rows)
# ---------------------------------------------------------------------------

def test_l2_closed_xlsx_explicit_range_keeps_trailing_empty_rows(tmp_path, capsys):
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["이름", "값", "비고"])
    ws.append(["가", 1, ""])
    p = tmp_path / "x.xlsx"
    wb.save(str(p))
    code, out = _run(excel, capsys, "read", "--target", str(p), "--range", "A1:C5")
    assert code == 0 and out["rows"] == 5 and out["range"] == "A1:C5"
    code, out = _run(excel, capsys, "read", "--target", str(p))
    assert code == 0 and out["rows"] == 2
