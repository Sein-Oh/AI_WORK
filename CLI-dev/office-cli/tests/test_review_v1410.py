"""1.4.10 회귀 테스트: 전체 검토에서 찾은 오류 (W1~W14, R1~R9).

항목마다 실패 상황을 재현하고 고친 동작을 확인합니다. 가짜 COM(fake_*.py)과 실제 파일(openpyxl, pymupdf)을 씁니다.
"""
import json
import os
import zipfile

import pytest

import _office as off
import _runtime as rt
import excel
import office
import outlook
import ppt
import reader
import word
from fake_excel import FakeExcelWorkbook, FakeXlApp
from fake_office import FakeWorkbook, make_win32com
from fake_ppt import FakePresentation, FakeShape
from fake_word import FakeWordDocument


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
# 공통 런타임: W6 ok(verified=False), W14 ok_file 다시 열기, R4 오류 로그 가리기
# ---------------------------------------------------------------------------

def test_ok_refuses_verified_false():
    with pytest.raises(ValueError):
        rt.ok("t", "c", "p", "VERIFIED: x", verified=False)


def test_ok_file_reopens_by_format(tmp_path):
    import openpyxl
    good = tmp_path / "a.xlsx"
    wb = openpyxl.Workbook()
    wb.active["A1"] = 1
    wb.save(str(good))
    assert rt.ok_file("t", "c", "p", str(good), "VERIFIED: ok")["agent_directive"].startswith("VERIFIED:")
    bad = tmp_path / "b.xlsx"
    bad.write_bytes(b"not a zip file")
    out = rt.ok_file("t", "c", "p", str(bad), "VERIFIED: ok")
    assert out["agent_directive"].startswith("UNVERIFIED:") and out["verified"] is False and "다시 열지 못함" in out["unverified"][0]
    for name in ("b.docx", "b.pptx"):
        (tmp_path / name).write_bytes(b"broken")
        assert rt.ok_file("t", "c", "p", str(tmp_path / name), "VERIFIED: ok")["verified"] is False
    csv = tmp_path / "r.csv"
    csv.write_text("a,b\n1,2\n", encoding="utf-8")
    assert rt.ok_file("t", "c", "p", str(csv), "VERIFIED: ok", expected_rows=2)["verified"] is True
    out = rt.ok_file("t", "c", "p", str(csv), "VERIFIED: ok", expected_rows=5)
    assert out["verified"] is False and "행 수" in out["unverified"][0]
    md = tmp_path / "n.md"
    md.write_text("# 제목", encoding="utf-8")
    assert rt.ok_file("t", "c", "p", str(md), "VERIFIED: ok")["verified"] is True
    other = tmp_path / "x.bin"
    other.write_bytes(b"1234")
    assert rt.ok_file("t", "c", "p", str(other), "VERIFIED: ok")["verified"] is False


def test_error_log_redacts_content_options(isolated_dirs):
    argv = ["new", "--to", "kim@company.com", "--subject", "인사 평가", "--body", "비밀 본문입니다",
            "--attach", "C:\\docs\\a.pdf", "--formula", "=A1*2", "=B1", "--text=메모 내용"]
    rt.write_error_log("outlook", argv, {"error_type": "input", "message": "m"})
    text = open(rt.error_log_path(), encoding="utf-8").read()
    for secret in ("kim@company.com", "인사 평가", "비밀 본문입니다", "=A1*2", "=B1", "메모 내용"):
        assert secret not in text
    assert "--to <15자 생략>" in text and "--body <8자 생략>" in text and "--text=<5자 생략>" in text
    assert "C:\\docs\\a.pdf" in text                                     # 경로는 남김
    assert rt.redact_argv(["transfer", "--from", "excel", "--to", "word"], "office") == \
        ["transfer", "--from", "excel", "--to", "word"]                      # office.py의 --from/--to는 앱 이름


# ---------------------------------------------------------------------------
# excel
# ---------------------------------------------------------------------------

@pytest.fixture
def xl(monkeypatch):
    app = FakeXlApp()
    wb = FakeExcelWorkbook(app, "보고.xlsx", sheets=("Sheet1", "요약"))
    _com(monkeypatch)
    monkeypatch.setattr(off, "get_active", lambda kind: {"name": wb.Name, "path": wb.FullName, "app": app, "doc": wb})
    return app, wb


def test_w1_replace_writes_to_sheet_named_in_range(xl, capsys):
    app, wb = xl
    s1, s2 = wb.Worksheets("Sheet1"), wb.Worksheets("요약")
    s1.set(1, 1, "초안")
    s2.set(1, 1, "초안 보고")
    code, out = _run(excel, capsys, "replace", "--find", "초안", "--replace", "최종", "--range", "요약!A1:A2")
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:")
    assert s2.values[(1, 1)] == "최종 보고" and s1.values[(1, 1)] == "초안"       # 이전: 지금 시트(Sheet1)에 씀
    code, out = _run(excel, capsys, "replace", "--find", "초안", "--replace", "x", "--range", "요약!A1", "--all-sheets")
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:") and "--all-sheets" in out["message"]


def test_w13_replace_without_replace_or_format_is_refused(xl, capsys):
    app, wb = xl
    wb.Worksheets("Sheet1").set(1, 1, "가")
    code, out = _run(excel, capsys, "replace", "--find", "가")
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:") and "--replace" in out["message"]
    code, out = _run(excel, capsys, "replace", "--find", "가", "--bold")
    assert code == 0 and out["changed_cells"] == 1


def test_w3_merge_single_value_outside_top_left_needs_decision(xl, capsys):
    app, wb = xl
    ws = wb.Worksheets("Sheet1")
    ws.set(1, 2, "남는값")                                   # A1은 비고 B1에만 값
    code, out = _run(excel, capsys, "range", "--action", "merge", "--range", "A1:B1")
    assert code == 1 and out["error_type"] == "needs_user_decision" and "B1" in out["message"]
    assert "$A$1:$B$1" not in ws.merged
    code, out = _run(excel, capsys, "range", "--action", "merge", "--range", "A1:B1", "--overwrite")
    assert code == 0 and "$A$1:$B$1" in ws.merged


def test_w4_formula_with_empty_result_is_not_overwritten(xl, capsys):
    app, wb = xl
    ws = wb.Worksheets("Sheet1")
    ws.formulas[(1, 1)] = '=IF(B1="","",B1)'
    ws.values[(1, 1)] = ""                                     # 수식 결과가 빈 글
    code, out = _run(excel, capsys, "write", "--range", "A1", "--value", "5")
    assert code == 1 and out["error_type"] == "needs_user_decision" and ws.formulas.get((1, 1))
    assert excel._cell_empty(ws.Cells(1, 1)) is False and excel._cell_empty(ws.Cells(9, 9)) is True


def test_w5_append_formula_scans_all_written_columns(xl, capsys):
    app, wb = xl
    ws = wb.Worksheets("Sheet1")
    for r in (1, 2, 3):
        ws.set(r, 1, r)
    for r in (1, 2, 3, 4):
        ws.set(r, 2, r * 10)                                   # B열이 한 행 더 김
    code, out = _run(excel, capsys, "write", "--append", "--range", "A1:B1", "--formula", "=ROW()", "=ROW()*2")
    assert code == 0, out
    assert ws.values[(4, 2)] == 40 and out["range_written"] == "A5:B5"      # 이전: 4행에 써서 B4를 덮음
    assert ws.formulas[(5, 1)] == "=ROW()" and ws.formulas[(5, 2)] == "=ROW()*2"


def test_w12_copy_relative_formula_is_verified(xl, capsys):
    app, wb = xl
    ws = wb.Worksheets("Sheet1")
    ws.set(1, 1, 3)
    ws.Range("B1").Formula = "=A1*2"
    code, out = _run(excel, capsys, "range", "--action", "copy", "--range", "A1:B1", "--dest", "D1")
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    assert ws.formulas[(1, 5)] == "=D1*2"                     # 상대 참조가 바뀌어 값은 다르지만 같은 수식


def test_w6_autofit_is_unverified(xl, capsys):
    app, wb = xl
    wb.Worksheets("Sheet1").set(1, 1, "가")
    code, out = _run(excel, capsys, "range", "--action", "autofit", "--range", "A1")
    assert code == 0 and out["verified"] is False and out["agent_directive"].startswith("UNVERIFIED:")


def _strip_dimension(src, dst):
    import re
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w") as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.startswith("xl/worksheets/sheet"):
                data = re.sub(rb"<dimension[^>]*/>", b"", data)
            zout.writestr(item, data)


def test_w10_closed_xlsx_reads_match_open_document(tmp_path, capsys, monkeypatch):
    import openpyxl
    monkeypatch.setattr(off, "win32com", None)
    src, nodim = tmp_path / "s.xlsx", tmp_path / "nodim.xlsx"
    wb = openpyxl.Workbook()
    for r in range(3, 6):
        for c in range(3, 6):
            wb.active.cell(r, c, r * 10 + c)                   # C3:E5
    wb.save(str(src))
    _strip_dimension(src, nodim)
    for path in (src, nodim):
        code, out = _run(excel, capsys, "read", "--target", str(path))
        assert code == 0 and out["range"] == "C3:E5" and out["rows"] == 3 and out["values"][0] == [33, 34, 35]
        code, out = _run(excel, capsys, "read", "--target", str(path), "--range", "E5:C3")       # 거꾸로 준 범위
        assert out["range"] == "C3:E5" and out["values"][-1] == [53, 54, 55]
        code, out = _run(excel, capsys, "read", "--target", str(path), "--range", "D:D")
        assert out["cols"] == 1 and out["values"][2:] == [[34], [44], [54]]
        code, out = _run(excel, capsys, "read", "--target", str(path), "--range", "4:4")
        assert out["values"] == [[None, None, 43, 44, 45]] or out["values"] == [["", "", 43, 44, 45]]


def test_w10_csv_reversed_range(tmp_path, capsys, monkeypatch):
    monkeypatch.setattr(off, "win32com", None)
    p = tmp_path / "t.csv"
    p.write_text("a,b\n1,2\n3,4\n", encoding="utf-8")
    code, out = _run(excel, capsys, "read", "--target", str(p), "--range", "B3:A2")
    assert code == 0 and out["values"] == [[1, 2], [3, 4]]


# ---------------------------------------------------------------------------
# word
# ---------------------------------------------------------------------------

@pytest.fixture
def wd(monkeypatch):
    holder = {"doc": FakeWordDocument("office 문서")}
    _com(monkeypatch)
    monkeypatch.setattr(off, "get_active", lambda kind: {"name": holder["doc"].Name, "path": holder["doc"].FullName,
                                                          "app": object(), "doc": holder["doc"]})
    monkeypatch.setattr(off, "save_snapshot", lambda *a, **k: None)
    return holder


def test_w7_case_only_replace_is_verified(wd, capsys):
    code, out = _run(word, capsys, "replace", "--find", "office", "--replace", "Office")
    assert code == 0 and wd["doc"].text.startswith("Office") and out["agent_directive"].startswith("VERIFIED:"), out


def test_w6_word_replace_limit_is_unverified(wd, capsys):
    wd["doc"] = FakeWordDocument("가 " * 5002)
    code, out = _run(word, capsys, "replace", "--find", "가", "--replace", "나")
    assert code == 0 and out["limit_reached"] is True and out["agent_directive"].startswith("UNVERIFIED:")


def test_word_compress_ranges():
    assert word._compress([7, 1, 2, 3, 9, 10]) == "1-3,7,9-10" and word._compress([]) == ""


def test_w13_word_replace_without_replace_or_format_is_refused(wd, capsys):
    code, out = _run(word, capsys, "replace", "--find", "office")
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:") and "--replace" in out["message"]
    code, out = _run(word, capsys, "replace", "--find", "office", "--replace", "")
    assert code == 0 and out["replaced"] == 1


# ---------------------------------------------------------------------------
# ppt
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


def _title(prs, n):
    return ppt.slide_units(prs.Slides(n), n)["title"]


def test_w2_title_with_markdown_file_goes_on_first_slide(pp, capsys, tmp_path):
    md = tmp_path / "내용.md"
    md.write_text("## 첫 절\n- 가\n\n## 둘째 절\n- 나\n", encoding="utf-8")
    code, out = _run(ppt, capsys, "write", "--slide", "new", "--title", "발표 제목", "--file", str(md))
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    prs = pp["prs"]
    assert _title(prs, 4) == "발표 제목" and "첫 절" in ppt.slide_units(prs.Slides(4), 4)["text"]
    assert _title(prs, 5) == "둘째 절"


def test_w9_existing_title_needs_overwrite(pp, capsys):
    prs = pp["prs"]
    prs.Slides(1).Shapes.Title.TextFrame.TextRange.Text = "기존 제목"
    code, out = _run(ppt, capsys, "write", "--slide", "1", "--title", "새 제목")
    assert code == 1 and out["error_type"] == "needs_user_decision" and "--overwrite" in out["agent_directive"]
    assert _title(prs, 1) == "기존 제목"
    code, out = _run(ppt, capsys, "write", "--slide", "1", "--title", "새 제목", "--overwrite")
    assert code == 0 and _title(prs, 1) == "새 제목"
    code, out = _run(ppt, capsys, "write", "--slide", "2", "--title", "빈 슬라이드 제목")      # 빈 제목은 묻지 않음
    assert code == 0 and _title(prs, 2) == "빈 슬라이드 제목"
    code, out = _run(ppt, capsys, "write", "--slide", "1", "--title", "새 제목", "--text", "- 본문")   # 같은 제목
    assert code == 0


def test_w11_truncated_read_continues_with_remaining_selected_slides(pp, capsys):
    prs = FakePresentation(9)
    for n in (1, 5, 9):
        prs.Slides(n).Shapes.add(FakeShape("본문", 36, 100, 800, 300, text="가" * 5000))
    pp["prs"] = prs
    code, out = _run(ppt, capsys, "read", "--slides", "1,5,9")
    assert code == 0 and out["truncated"] is True and [s["no"] for s in out["slides"]] == [1]
    assert out["next_commands"][0]["command"].endswith('--slides "5,9"')          # 이전: 2-9


# ---------------------------------------------------------------------------
# office: W6 transfer 일부 실패, R5 undo --pick, W8 파일 이름만 준 대상
# ---------------------------------------------------------------------------

def test_w6_transfer_partial_failure_is_unverified(monkeypatch, capsys, tmp_path):
    app = FakeXlApp()
    wb = FakeExcelWorkbook(app)
    _com(monkeypatch)
    monkeypatch.setattr(off, "get_active", lambda kind: {"name": wb.Name, "path": wb.FullName, "app": app, "doc": wb})
    chart = type("C", (), {"Name": "차트"})()
    monkeypatch.setattr(excel, "find_chart", lambda t, name, sheet=None: (None, chart))
    monkeypatch.setattr(excel, "_chart_title", lambda obj: "차트")
    png = tmp_path / "c.png"
    png.write_bytes(office.PNG_SIGNATURE + b"x" * 50)

    def boom(*a, **k):
        raise RuntimeError("PPT 실패")
    monkeypatch.setattr(office, "image_by_export", lambda obj, stem: str(png))
    monkeypatch.setattr(office, "image_by_redraw", lambda obj, stem: str(png))
    monkeypatch.setattr(office, "relay_chart", boom)
    monkeypatch.setattr(office, "insert_picture_word", lambda tgt, img, anchor, occurrence=None: {"app": "Word"})
    monkeypatch.setattr(office, "insert_picture_ppt", boom)
    monkeypatch.setattr(rt, "load_config", lambda: {})
    code, out = _run(office, capsys, "transfer", "--from", "excel", "--chart", "차트", "--to", "word,ppt")
    assert code == 0 and "word" in out["results"] and "ppt" in out["failures"]
    assert out["verified"] is False and out["agent_directive"].startswith("UNVERIFIED:")


def test_r5_undo_pick_removes_newer_points_of_same_document(tmp_path, monkeypatch, capsys):
    _com(monkeypatch)
    monkeypatch.setattr(off, "match_open", lambda kind, target, docs=None: None)
    orig = tmp_path / "a.xlsx"
    other = tmp_path / "b.xlsx"
    orig.write_bytes(b"current")
    other.write_bytes(b"other")
    off.save_snapshot("excel", FakeWorkbook(str(orig)), "excel", "write")         # 앞 지점 (고를 것)
    off.save_snapshot("excel", FakeWorkbook(str(other)), "excel", "write")        # 다른 문서
    off.save_snapshot("excel", FakeWorkbook(str(orig)), "excel", "style")         # 같은 문서의 새 지점
    snaps = off.list_snapshots()
    assert [os.path.basename(s["original_path"]) for s in snaps] == ["a.xlsx", "b.xlsx", "a.xlsx"]
    # 앞 지점은 한 시간 전, 파일은 도구가 새 지점 뒤에 저장: 사용자 저장이 아니므로 묻지 않음
    old = snaps[2]
    old["created_ts"] = float(old["created_ts"]) - 3600
    meta_file = old.pop("meta_file")
    with open(meta_file, "w", encoding="utf-8") as f:
        json.dump(old, f, ensure_ascii=False)
    code, out = _run(office, capsys, "undo", "--pick", "3")
    assert code == 0 and orig.read_bytes() == b"snapshot-of-a.xlsx" and out["removed_newer_points"] == 1, out
    left = off.list_snapshots()
    assert [os.path.basename(s["original_path"]) for s in left] == ["b.xlsx"]       # 다음 undo가 앞으로 가지 않음


def test_w8_bare_name_is_not_created_in_cwd(monkeypatch, tmp_path):
    _com(monkeypatch)
    monkeypatch.chdir(tmp_path)
    monkeypatch.setattr(off, "list_open", lambda kind: [{"name": "보고.xlsx", "path": "C:\\docs\\보고.xlsx"}])
    monkeypatch.setattr(off, "match_open", lambda kind, target, docs=None: None)
    with pytest.raises(rt.InputError) as ei:
        with off.open_target("excel", "새파일.xlsx", readonly=False, create=True):
            pass
    assert ei.value.candidates == ["보고.xlsx"] and "전체 경로" in str(ei.value)
    assert not os.path.exists(tmp_path / "새파일.xlsx")
    monkeypatch.setattr(off, "list_open", lambda kind: [])
    with pytest.raises(rt.InputError) as ei:
        with off.open_target("excel", "새파일.xlsx", readonly=False, create=True):
            pass
    assert "[후보]" not in str(ei.value) and "찾을 수 없습니다" in str(ei.value)


# ---------------------------------------------------------------------------
# outlook: R6 저장 이름, R7 본문 줄바꿈, R8 저장 확인, R9 첨부 구분
# ---------------------------------------------------------------------------

class _Mail:
    def __init__(self):
        self.To = ""
        self.Subject = ""
        self.HTMLBody = "<html><body><p>서명</p></body></html>"
        self.Attachments = type("A", (), {"added": [], "Add": lambda s, p: s.added.append(p), "Count": 0})()
        self.GetInspector = object()
        self.displayed = False

    def Display(self):
        self.displayed = True


def _outlook(monkeypatch, capsys, mail, *argv):
    monkeypatch.setattr(outlook, "get_outlook", lambda: type("App", (), {"CreateItem": lambda s, k: mail})())
    return _run(outlook, capsys, *argv)


def test_r6_r8_saved_mail_name_is_sanitized_and_verified(isolated_dirs):
    item = type("M", (), {"Subject": 'RE: 견적/회신? "최종"', "Body": "본문입니다", "To": "a", "CC": "",
                          "SenderName": "김", "SenderEmailAddress": "", "EntryID": "1", "UnRead": False,
                          "Attachments": type("A", (), {"Count": 0})()})()
    args = type("Args", (), {"save": True, "save_attach": False})()
    res = outlook._mail_detail(item, args, "id")
    name = os.path.basename(res["output_file"])
    assert not any(ch in name for ch in ':/?"') and name.startswith("mail_RE_ 견적_회신_")
    assert os.path.isfile(res["output_file"]) and "save_problem" not in res
    assert outlook._safe_name("...", 10) == "메일"


def test_r7_body_file_keeps_windows_paths(monkeypatch, capsys, tmp_path):
    body = tmp_path / "본문.txt"
    body.write_text("자료 위치: C:\\new\\data\\n폴더", encoding="utf-8")
    mail = _Mail()
    code, out = _outlook(monkeypatch, capsys, mail, "new", "--body-file", str(body))
    assert code == 0 and "C:\\new\\data\\n폴더" in mail.HTMLBody
    mail = _Mail()
    code, out = _outlook(monkeypatch, capsys, mail, "new", "--body", "경로 C:\\new\\data.xlsx\\n다음 줄")
    assert "C:\\new\\data.xlsx<br>다음 줄" in mail.HTMLBody


def test_r9_attach_name_with_comma_and_semicolon_list(monkeypatch, capsys):
    data = rt.get_data_dir()
    for n in ("보고서,최종.pdf", "a.pdf", "b.pdf"):
        open(os.path.join(data, n), "w").close()
    mail = _Mail()
    code, out = _outlook(monkeypatch, capsys, mail, "new", "--attach", "보고서,최종.pdf")
    assert code == 0 and out["attached"] == ["보고서,최종.pdf"]
    mail = _Mail()
    code, out = _outlook(monkeypatch, capsys, mail, "new", "--attach", "a.pdf;b.pdf")
    assert code == 0 and out["attached"] == ["a.pdf", "b.pdf"]


# ---------------------------------------------------------------------------
# reader: R1 긴 단위, R2 글자 없는 쪽, R3 상한, R8 저장 확인, R9 캐시 저장 실패
# ---------------------------------------------------------------------------

def test_r1_long_unit_is_marked_truncated_with_save_command(tmp_path, capsys):
    p = tmp_path / "긴글.txt"
    p.write_text("가" * 9000 + "\n\n끝 단락", encoding="utf-8")
    code, out = _run(reader, capsys, "read", "--file", str(p), "--units", "1")
    assert code == 0 and out["truncated"] is True and out["units"][0]["text_truncated"] is True
    assert out["units"][0]["text_chars"] == 9000 and len(out["units"][0]["text"]) == rt.DEFAULT_MAX_CHARS
    assert any("--save" in c["command"] and ".md" in c["command"] for c in out["next_commands"]) and "truncated_reason" in out


def test_r2_partially_scanned_pdf_reports_pages_without_text(tmp_path, capsys):
    import pymupdf
    p = tmp_path / "부분스캔.pdf"
    doc = pymupdf.open()
    doc.new_page().insert_text((72, 72), "first page text", fontname="helv", fontsize=11)
    doc.new_page().insert_image(pymupdf.Rect(72, 72, 200, 200), pixmap=_pixmap(pymupdf))   # 스캔 쪽(그림만, 1.4.11)
    doc.new_page().insert_text((72, 72), "third page text", fontname="helv", fontsize=11)
    doc.save(str(p))
    doc.close()
    code, out = _run(reader, capsys, "read", "--file", str(p))
    assert code == 0 and out["meta"]["pages_without_text"] == [2] and out["pages_without_text"] == [2]
    assert "그림만 있는 쪽" in out["agent_directive"]
    code, out = _run(reader, capsys, "search", "--query", "warranty", "--file", str(p))
    assert code == 0 and out["total_matches"] == 0 and out["agent_directive"].startswith("UNVERIFIED:")
    assert any("그림만 있는 쪽" in g for g in out["not_searched"])


def _pixmap(pymupdf):
    pix = pymupdf.Pixmap(pymupdf.csRGB, pymupdf.IRect(0, 0, 20, 20), False)
    pix.clear_with(200)
    return pix


def test_r3_sheet_row_limit_and_save_hit_limit_are_reported(tmp_path, capsys, monkeypatch):
    import openpyxl
    monkeypatch.setattr(reader, "MAX_SHEET_ROWS", 5)
    monkeypatch.setattr(reader, "MAX_SAVE_HITS", 2)
    p = tmp_path / "많은행.xlsx"
    wb = openpyxl.Workbook()
    for i in range(10):
        wb.active.append([f"항목 {i}", "공통"])
    wb.save(str(p))
    code, out = _run(reader, capsys, "read", "--file", str(p))
    assert code == 0 and out["truncated"] is True and any("5행" in r for r in out["truncated_reason"])
    code, out = _run(reader, capsys, "read", "--file", str(p), "--save", str(tmp_path / "o.csv"))
    assert code == 0 and out["truncated"] is True and out["agent_directive"].startswith("UNVERIFIED:")
    code, out = _run(reader, capsys, "search", "--query", "공통", "--file", str(p), "--save", str(tmp_path / "hits.csv"))
    assert code == 0 and out["saved"] == 2 and out["save_truncated"] is True
    assert out["agent_directive"].startswith("UNVERIFIED:") and "앞 2건만 저장" in out["agent_directive"]


def test_r8_reader_save_is_reopened(tmp_path, capsys, monkeypatch):
    p = tmp_path / "a.txt"
    p.write_text("첫 단락\n\n둘째 단락", encoding="utf-8")
    code, out = _run(reader, capsys, "read", "--file", str(p), "--save", str(tmp_path / "o.csv"))
    assert code == 0 and out["agent_directive"].startswith("VERIFIED: 저장한 파일을 다시 열어")
    monkeypatch.setattr(rt, "check_saved_file", lambda path, rows=None: "저장한 파일을 다시 열지 못함 (테스트)")
    code, out = _run(reader, capsys, "read", "--file", str(p), "--save", str(tmp_path / "o2.md"))
    assert code == 0 and out["agent_directive"].startswith("UNVERIFIED:")
    code, out = _run(reader, capsys, "search", "--query", "단락", "--file", str(p), "--save", str(tmp_path / "s.md"))
    assert code == 0 and out["agent_directive"].startswith("UNVERIFIED:") and out["save_problem"]


def test_r9_index_reports_cache_write_failure(tmp_path, capsys, monkeypatch):
    p = tmp_path / "a.txt"
    p.write_text("내용", encoding="utf-8")
    real = rt.safe_write_text

    def failing(path, text, encoding="utf-8"):
        if path.endswith(".json") and "index" not in os.path.basename(path):
            raise OSError("디스크 가득 참")
        return real(path, text, encoding)
    monkeypatch.setattr(rt, "safe_write_text", failing)
    code, out = _run(reader, capsys, "index", "--file", str(p))
    assert code == 0 and out["agent_directive"].startswith("UNVERIFIED:") and out["cache_failed"] == ["a.txt"]
    assert "캐시에 저장 확인" not in out["agent_directive"]


def test_r9_normalize_doc_compat_filler_is_removed():
    assert not hasattr(reader, "normalize_doc") and reader.PARSER_VERSION == 6
