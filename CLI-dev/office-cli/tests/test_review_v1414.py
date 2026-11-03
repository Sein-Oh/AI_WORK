"""1.4.14 회귀 테스트: 5차 점검에서 찾은 오류 (N1 바꾼 것 없는 성공의 되돌리기 지점, N2 잘못된 시트 이름,
N3 스캔 PDF 읽기·저장, N4 디코딩되지 않는 색인, N5 search 저장 안내, N6 서식만 있는 먼 칸이 있는 닫힌 xlsx).

항목마다 실패 상황을 재현하고 고친 동작을 확인합니다. 가짜 COM(fake_*.py)과 실제 파일(openpyxl, pymupdf 등)을 씁니다.
"""
import json
import types

import pytest

import _office as off
import _runtime as rt
import excel
import office
import ppt
import reader
from fake_excel import FakeExcelWorkbook, FakeXlApp
from fake_office import FakeExcelApp, make_win32com
from fake_ppt import FakePresentation


def _run(module, capsys, *argv):
    rt.RECOVERED.clear()
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


@pytest.fixture
def xl(monkeypatch):
    app = FakeXlApp()
    wb = FakeExcelWorkbook(app, "보고.xlsx", sheets=("Sheet1", "요약"))
    _com(monkeypatch)
    monkeypatch.setattr(off, "get_active", lambda kind: {"name": wb.Name, "path": wb.FullName, "app": app, "doc": wb})
    return app, wb


def _points(kind="excel"):
    return sorted(s["meta_file"] for s in off.list_snapshots(kind))


def _sheet_names(wb):
    return [str(wb.Worksheets(i).Name) for i in range(1, wb.Worksheets.Count + 1)]


# ---------------------------------------------------------------------------
# N1: 바꾼 것 없이 성공한 실행은 되돌리기 지점을 남기지 않음 (다음 undo가 바뀐 것 없는 지점으로 가지 않게)
# ---------------------------------------------------------------------------

def test_n1_noop_replace_drops_its_point(xl, capsys):
    code, out = _run(excel, capsys, "write", "--range", "A1", "--value", "초안")
    assert code == 0, out
    before = _points()
    assert len(before) == 1
    code, out = _run(excel, capsys, "replace", "--find", "없는글", "--replace", "x")
    assert code == 0 and out["changed_cells"] == 0, out
    assert _points() == before                              # 다음 undo는 write를 되돌림
    code, out = _run(office, capsys, "undo", "--list")
    assert code == 0 and len(out["snapshots"]) == 1 and out["snapshots"][0]["by"] == "excel write", out


def test_n1_chart_delete_all_with_no_chart_drops_point(xl, capsys):
    code, out = _run(excel, capsys, "chart", "--action", "delete", "--all")
    assert code == 0 and "0개" in out["preview"], out
    assert _points() == []


def test_n1_sheet_add_existing_and_select_drop_point(xl, capsys):
    code, out = _run(excel, capsys, "sheet", "--action", "add", "--name", "요약")
    assert code == 0 and "이미 있어" in out["preview"], out
    code, out = _run(excel, capsys, "sheet", "--action", "select", "--name", "요약")
    assert code == 0, out
    assert _points() == []
    code, out = _run(excel, capsys, "sheet", "--action", "add", "--name", "새시트")    # 실제로 바꾸면 지점이 남음
    assert code == 0 and len(_points()) == 1, out["message"]


def test_n1_real_change_on_unsaved_document_keeps_point(xl, capsys):
    _, wb = xl
    wb.Saved = False                         # 저장 상태로는 바뀐 것을 알 수 없음: 시트 삭제가 t.changed를 표시해야 함
    code, out = _run(excel, capsys, "sheet", "--action", "delete", "--name", "요약")
    assert code == 0 and _sheet_names(wb) == ["Sheet1"], out
    assert len(_points()) == 1


def test_n1_closed_file_noop_drops_point(monkeypatch, tmp_path):
    path = tmp_path / "닫힌.xlsx"
    path.write_bytes(b"original")
    _com(monkeypatch)
    monkeypatch.setattr(off, "_new_instance", lambda kind: FakeExcelApp())
    with off.open_target("excel", str(path), readonly=False, tool="excel", command="replace"):
        pass                                 # 아무것도 바꾸지 않음 -> 저장도 하지 않음
    assert _points() == [] and path.read_bytes() == b"original"


def test_n1_ppt_clear_of_empty_slide_drops_point(monkeypatch, capsys):
    prs = FakePresentation(2)
    app = types.SimpleNamespace(ActiveWindow=types.SimpleNamespace(View=types.SimpleNamespace(GotoSlide=lambda n: None),
                                                                   Presentation=prs))
    _com(monkeypatch)
    monkeypatch.setattr(off, "save_snapshot", lambda *a, **k: None)      # 지점은 장벽(barrier)으로 남음
    monkeypatch.setattr(off, "get_active", lambda kind: {"name": prs.Name, "path": prs.FullName, "app": app, "doc": prs})
    code, out = _run(ppt, capsys, "slide", "--action", "clear", "--slide", "1")
    assert code == 0, out
    first = _points("ppt")
    assert len(first) == 1
    code, out = _run(ppt, capsys, "slide", "--action", "clear", "--slide", "1")   # 이미 빈 슬라이드
    assert code == 0 and "0개" in out["preview"], out
    assert _points("ppt") == first


# ---------------------------------------------------------------------------
# N2: 잘못된 시트 이름은 시트를 만들기 전에 막음 ([추천] 고친 이름, 빈 시트가 남지 않음)
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("argv, rec", [
    (["write", "--new-sheet", "--sheet", "1/2분기", "--range", "A1", "--value", "1"], "1_2분기"),
    (["sheet", "--action", "add", "--name", "1/2분기"], "1_2분기"),
    (["sheet", "--action", "add", "--name", "'따옴표'"], "_따옴표_"),
    (["sheet", "--action", "rename", "--name", "요약", "--new-name", "a:b?"], "a_b_"),
    (["sheet", "--action", "copy", "--name", "Sheet1", "--new-name", "x" * 40], "x" * 31),
])
def test_n2_invalid_sheet_name_is_refused_before_any_change(xl, capsys, argv, rec):
    _, wb = xl
    code, out = _run(excel, capsys, *argv)
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:"), out
    assert f"[추천] {rec}" in out["message"], out
    assert _sheet_names(wb) == ["Sheet1", "요약"] and _points() == []


def test_n2_recommended_name_avoids_existing_sheet(xl, capsys):
    code, out = _run(excel, capsys, "sheet", "--action", "add", "--name", "Sheet[1]")
    assert code == 1 and "[추천] Sheet_1_" in out["message"], out
    _, wb = xl
    wb.Worksheets.Add().Name = "Sheet_1_"
    code, out = _run(excel, capsys, "sheet", "--action", "add", "--name", "Sheet[1]")
    assert code == 1 and "[추천] Sheet_1_ 2" in out["message"], out


def test_n2_default_copy_name_of_long_sheet_fits_31_chars(xl, capsys):
    _, wb = xl
    wb.Worksheets(1).Name = "가" * 31
    code, out = _run(excel, capsys, "sheet", "--action", "copy", "--name", "가" * 31)
    assert code == 0, out
    assert len(_sheet_names(wb)[-1]) <= 31


def test_n2_transfer_bad_sheet_name_checked_before_opening(monkeypatch, tmp_path):
    # 새로 만들 파일은 열기(만들기) 전에 확인. 열린 문서는 시트 이름을 안 뒤에 확인 (1.4.15, test_review_v1415.py)
    def boom(*a, **k):
        raise AssertionError("문서를 열면 안 됨")
    monkeypatch.setattr(off, "open_target", boom)
    args = types.SimpleNamespace(sheet="표/목록", to_target=str(tmp_path / "새.xlsx"), overwrite=False)
    with pytest.raises(rt.InputError) as ei:
        office._to_excel(args, [["a"]], "Word 표")
    assert "[추천] 표_목록" in str(ei.value)


def test_n2_target_sheet_create_validates_name(xl):
    _, wb = xl
    with pytest.raises(rt.InputError):
        with off.open_target("excel", "active", readonly=False, tool="excel", command="write") as t:
            t.sheet("a*b", create=True)
    assert _sheet_names(wb) == ["Sheet1", "요약"] and _points() == []


# ---------------------------------------------------------------------------
# N3: 스캔 PDF 저장, 스캔 쪽만 고른 --pages는 UNVERIFIED (0개 저장이면 next_commands 없음)
# ---------------------------------------------------------------------------

def _pix(pymupdf):
    pix = pymupdf.Pixmap(pymupdf.csRGB, pymupdf.IRect(0, 0, 40, 56), False)
    pix.clear_with(200)
    return pix


def _scan_pdf(tmp_path, stamp_page=None, pages=2, scan_all=True):
    import pymupdf
    doc = pymupdf.open()
    for i in range(1, pages + 1):
        p = doc.new_page()
        if scan_all or i == stamp_page:
            p.insert_image(pymupdf.Rect(30, 30, 565, 812), pixmap=_pix(pymupdf))
            if i == stamp_page:
                p.insert_text((400, 400), "RECEIVED", fontsize=9)
        else:
            p.insert_text((72, 200), f"Body text page {i} torque {i * 10} Nm.", fontsize=11)
    path = tmp_path / ("scan.pdf" if scan_all else "stamp.pdf")
    doc.save(str(path))
    return path


def test_n3_save_of_all_scanned_pdf_is_unverified_without_next_commands(tmp_path, capsys):
    path = _scan_pdf(tmp_path)
    out_md = tmp_path / "스캔.md"
    code, out = _run(reader, capsys, "read", "--file", str(path), "--save", str(out_md))
    assert code == 0 and out["count"] == 0, out
    assert out["agent_directive"].startswith("UNVERIFIED:") and "OCR" in out["agent_directive"], out
    assert "next_commands" not in out


def test_n3_pages_selecting_only_scanned_page_are_unverified(tmp_path, capsys):
    path = _scan_pdf(tmp_path, stamp_page=3, pages=4, scan_all=False)
    code, out = _run(reader, capsys, "read", "--file", str(path), "--pages", "3")
    assert code == 0 and out["count"] == 1 and out["pages_without_text"] == [3], out
    assert out["agent_directive"].startswith("UNVERIFIED:") and "p.3" in out["agent_directive"], out
    code, out = _run(reader, capsys, "read", "--file", str(path), "--pages", "3", "--save", str(tmp_path / "p3.md"))
    assert code == 0 and out["agent_directive"].startswith("UNVERIFIED:"), out
    assert "next_commands" not in out
    code, out = _run(reader, capsys, "read", "--file", str(path), "--pages", "2-3")       # 본문 쪽이 섞이면 주의만
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out


# ---------------------------------------------------------------------------
# N4: 디코딩되지 않는 색인은 손상으로 보고 다시 만듦 (모든 reader 명령이 실패하던 문제)
# ---------------------------------------------------------------------------

def test_n4_undecodable_index_is_rebuilt(tmp_path, capsys):
    folder = tmp_path / "docs"
    folder.mkdir()
    for i in range(2):
        (folder / f"d{i}.txt").write_text(f"문서 {i} 토크", encoding="utf-8")
    code, out = _run(reader, capsys, "index", "--folder", str(folder))
    assert code == 0, out
    with open(reader._index_path(), "wb") as f:
        f.write(b'{"a": \xff\xfe\x80\x80')
    code, out = _run(reader, capsys, "search", "--query", "토크")
    assert code == 0 and out["total_matches"] == 2, out
    assert any("색인 파일 손상" in r for r in out.get("recovered", [])), out
    with open(reader._index_path(), "wb") as f:
        f.write(b'{"a": \xff\xfe\x80\x80')
    code, out = _run(reader, capsys, "clear", "--all")
    assert code == 0, out


# ---------------------------------------------------------------------------
# N5: search 저장 안내는 next_commands를 가리킴 (directive에 근거.xlsx를 박지 않음)
# ---------------------------------------------------------------------------

def test_n5_search_directive_points_to_next_commands(tmp_path, capsys):
    f = tmp_path / "huge.txt"
    f.write_text("앞 단락 토크\n\n" + "토크 " + "가" * 33000 + "\n", encoding="utf-8")
    code, out = _run(reader, capsys, "search", "--query", "토크", "--file", str(f), "--full", "--top", "1")
    assert code == 0 and out["truncated"], out
    saves = [c["command"] for c in out["next_commands"] if "--save" in c["command"]]
    assert saves and "근거.md" in saves[0], saves
    assert "근거.xlsx" not in out["agent_directive"] and "next_commands" in out["agent_directive"], out


# ---------------------------------------------------------------------------
# N6: 닫힌 xlsx에 서식만 있는 먼 칸(Z300000)이 있어도 큰 범위 읽기는 실제 데이터 끝까지
# ---------------------------------------------------------------------------

def test_n6_stray_formatted_cell_does_not_block_clip(tmp_path, capsys):
    import openpyxl
    from openpyxl.styles import PatternFill
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "데이터"
    for r in range(1, 11):
        ws.append([f"a{r}", r, r * 2])
    ws["Z300000"].fill = PatternFill("solid", fgColor="FFFF00")          # 값 없이 서식만
    p = tmp_path / "stray.xlsx"
    wb.save(p)
    code, out = _run(excel, capsys, "read", "--target", str(p), "--range", "A1:XFD1048576")
    assert code == 0 and out["range"] == "A1:C10" and out["rows"] == 10, out
    assert any("A1:C10" in r for r in out.get("recovered", [])), out
