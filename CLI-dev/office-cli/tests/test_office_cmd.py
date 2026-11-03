"""office.py 테스트: 차트 전달 방식 대체 순서(DRM 대응), matplotlib 다시 그리기, undo 복원, 워드/PPT -> 엑셀."""
import json
import os

import pytest

import _office as off
import _runtime as rt
import excel
import office
from fake_excel import FakeExcelWorkbook, FakeXlApp
from fake_office import FakeExcelApp, FakeWorkbook, make_win32com


def run(capsys, *argv):
    with pytest.raises(SystemExit) as ei:
        office.main(list(argv))
    out = [ln for ln in capsys.readouterr().out.splitlines() if ln.strip()]
    assert len(out) == 1, out
    return ei.value.code, json.loads(out[0])


# ---------------------------------------------------------------------------
# 차트 그림
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("kind", [4, 51, 57, 5, 1, 75, -4169])
def test_render_chart_writes_plain_png(tmp_path, kind):
    data = {"type": kind, "title": "사인 2주기", "width": 300, "height": 180,
            "series": [{"name": "sin(x)", "values": [0, 1, 0, -1, 0], "x": [0, 1.57, 3.14, 4.71, 6.28]},
                       {"name": "b", "values": [1, 2, 3, 2, 1], "x": ["가", "나", "다", "라", "마"]}]}
    path = office.render_chart(data, str(tmp_path / "c.png"))
    assert office.is_plain_png(path) and os.path.getsize(path) > 1000


def test_is_plain_png_detects_encrypted(tmp_path):
    enc = tmp_path / "drm.png"
    enc.write_bytes(b"DRMONE" + os.urandom(100))
    assert not office.is_plain_png(str(enc))


def test_method_order_uses_probe_result(monkeypatch):
    monkeypatch.setattr(rt, "load_config", lambda: {"transfer_method": {"ppt": ["relay"]}})
    assert office.method_order("ppt", None) == ["relay", "export", "redraw"]
    assert office.method_order("word", None) == ["export", "redraw", "relay"]
    assert office.method_order("ppt", "redraw") == ["redraw"]


@pytest.fixture
def chart_env(monkeypatch, tmp_path):
    """엑셀 차트가 하나 있는 상태. 삽입 함수와 그림 방식은 기록용 대역."""
    app = FakeXlApp()
    wb = FakeExcelWorkbook(app)
    fake, _ = make_win32com()
    monkeypatch.setattr(off, "win32com", fake)
    monkeypatch.setattr(off, "pythoncom", None)
    monkeypatch.setattr(off, "list_open", lambda kind: [])
    monkeypatch.setattr(off, "get_active", lambda kind: {"name": wb.Name, "path": wb.FullName, "app": app, "doc": wb})
    chart = type("C", (), {"Name": "사인 2주기"})()
    monkeypatch.setattr(excel, "find_chart", lambda t, name, sheet=None: (None, chart))
    monkeypatch.setattr(excel, "_chart_title", lambda obj: "사인 2주기")
    log = {"insert": [], "calls": []}

    def png(name):
        p = tmp_path / name
        p.write_bytes(office.PNG_SIGNATURE + b"x" * 50)
        return str(p)

    def export(obj, stem):
        log["calls"].append("export")
        raise RuntimeError("엑셀이 저장한 PNG가 보안 프로그램으로 암호화되어 다른 앱이 읽을 수 없습니다")

    def redraw(obj, stem):
        log["calls"].append("redraw")
        return png("r.png")

    monkeypatch.setattr(office, "image_by_export", export)
    monkeypatch.setattr(office, "image_by_redraw", redraw)
    monkeypatch.setattr(office, "insert_picture_word", lambda tgt, img, anchor, occurrence=None: log["insert"].append(("word", img)) or {"app": "Word"})
    monkeypatch.setattr(office, "insert_picture_ppt", lambda tgt, img, slide, title: log["insert"].append(("ppt", img)) or {"app": "PowerPoint"})
    monkeypatch.setattr(rt, "load_config", lambda: {})
    return log


def test_chart_transfer_falls_back_from_encrypted_export_to_redraw(chart_env, capsys):
    code, out = run(capsys, "transfer", "--from", "excel", "--chart", "사인 2주기", "--to", "word,ppt")
    assert code == 0, out
    assert out["results"]["word"]["method"] == "redraw" and out["results"]["ppt"]["method"] == "redraw"
    assert chart_env["calls"].count("redraw") == 1                     # 한 번 그린 그림을 두 앱에 재사용
    assert [k for k, _ in chart_env["insert"]] == ["word", "ppt"]


def test_chart_transfer_all_methods_fail_suggests_probe(chart_env, capsys, monkeypatch):
    def boom(*a, **k):
        raise RuntimeError("실패")
    monkeypatch.setattr(office, "image_by_redraw", boom)
    monkeypatch.setattr(office, "relay_chart", boom)
    code, out = run(capsys, "transfer", "--from", "excel", "--chart", "사인 2주기", "--to", "ppt")
    assert code == 1 and "--probe" in out["message"]


def test_transfer_argument_rules(capsys):
    code, out = run(capsys, "transfer", "--from", "excel", "--to", "word")
    assert code == 1 and "--chart" in out["message"]
    code, out = run(capsys, "transfer", "--from", "word", "--to", "ppt")
    assert code == 1 and "read --save" in out["message"]
    code, out = run(capsys, "transfer", "--from", "엑셀", "--to", "엑셀", "--range", "A1")
    assert code == 1
    code, out = run(capsys, "transfer", "--from", "hwp", "--to", "word")
    assert code == 1 and "[추천]" in out["message"] or "사용 가능한" in out["message"]


# ---------------------------------------------------------------------------
# 워드/PPT -> 엑셀
# ---------------------------------------------------------------------------

def test_word_text_to_new_excel_sheet_with_unique_name(monkeypatch, capsys):
    app = FakeXlApp()
    wb = FakeExcelWorkbook(app, sheets=("Sheet1", "Word 내용"))
    fake, _ = make_win32com()
    monkeypatch.setattr(off, "win32com", fake)
    monkeypatch.setattr(off, "pythoncom", None)
    monkeypatch.setattr(off, "list_open", lambda kind: [])
    monkeypatch.setattr(off, "get_active", lambda kind: {"name": wb.Name, "path": wb.FullName, "app": app, "doc": wb})
    monkeypatch.setattr(off, "save_snapshot", lambda *a, **k: None)
    rows = [["번호", "구분", "상위 제목", "내용"], [1, "제목1", "배경", "배경"], [2, "본문", "배경", "내용입니다"]]
    monkeypatch.setattr(office, "_word_rows", lambda args: (rows, "계획.docx"))
    code, out = run(capsys, "transfer", "--from", "word", "--to", "excel")
    assert code == 0 and out["sheet"] == "Word 내용 2" and out["rows"] == 3
    ws = wb.Worksheets("Word 내용 2")
    assert ws.values[(3, 4)] == "내용입니다"


def test_word_text_to_excel_stays_text(monkeypatch, capsys):
    app = FakeXlApp()
    wb = FakeExcelWorkbook(app, sheets=("Sheet1",))
    fake, _ = make_win32com()
    monkeypatch.setattr(off, "win32com", fake)
    monkeypatch.setattr(off, "pythoncom", None)
    monkeypatch.setattr(off, "list_open", lambda kind: [])
    monkeypatch.setattr(off, "get_active", lambda kind: {"name": wb.Name, "path": wb.FullName, "app": app, "doc": wb})
    monkeypatch.setattr(off, "save_snapshot", lambda *a, **k: None)
    rows = [["코드", "기간", "손익", "식"], ["001", "1-2", "(100)", "=A1+1"]]
    monkeypatch.setattr(office, "_word_rows", lambda args: (rows, "표.docx"))
    code, out = run(capsys, "transfer", "--from", "word", "--to", "excel")
    ws = wb.Worksheets(out["sheet"])
    assert [ws.values[(2, c)] for c in range(1, 5)] == ["001", "1-2", "(100)", "=A1+1"]
    assert {(2, c) for c in range(1, 5)} <= ws.prefixed                    # written as text, not number/date/formula
    assert out["agent_directive"].startswith("VERIFIED:") and out["verified"] is True    # 넣은 칸을 다시 읽음 (R1)


def test_transfer_to_existing_excel_sheet_does_not_overwrite(monkeypatch, capsys):
    app = FakeXlApp()
    wb = FakeExcelWorkbook(app, sheets=("Sheet1", "결과"))
    fake, _ = make_win32com()
    monkeypatch.setattr(off, "win32com", fake)
    monkeypatch.setattr(off, "pythoncom", None)
    monkeypatch.setattr(off, "list_open", lambda kind: [])
    monkeypatch.setattr(off, "get_active", lambda kind: {"name": wb.Name, "path": wb.FullName, "app": app, "doc": wb})
    monkeypatch.setattr(off, "save_snapshot", lambda *a, **k: None)
    wb.Worksheets("결과").values[(1, 1)] = "기존"
    monkeypatch.setattr(office, "_word_rows", lambda args: ([["a", "b"], [1, 2]], "계획.docx"))
    code, out = run(capsys, "transfer", "--from", "word", "--to", "excel", "--sheet", "결과")
    assert code == 1 and out["message"].endswith("[추천] 결과 2") and "--overwrite" in out["message"]
    assert wb.Worksheets("결과").values[(1, 1)] == "기존"
    code, out = run(capsys, "transfer", "--from", "word", "--to", "excel", "--sheet", "결과", "--overwrite")
    assert code == 0 and out["sheet"] == "결과" and wb.Worksheets("결과").values[(1, 1)] == "a"
    code, out = run(capsys, "transfer", "--from", "word", "--to", "excel", "--sheet", "새 결과")
    assert code == 0 and out["sheet"] == "새 결과"                       # --sheet는 새 시트 이름: 없으면 만듦


def test_chart_transfer_stops_on_ambiguous_anchor_without_trying_other_methods(chart_env, capsys, monkeypatch):
    calls = []

    def ambiguous_anchor(tgt, img, anchor, occurrence=None):
        calls.append(img)
        raise rt.ambiguous("'계획'이(가) 단락 2개에 있어 어느 곳인지 고를 수 없습니다.", ["1: 계획 A", "2: 계획 B"])
    monkeypatch.setattr(office, "insert_picture_word", ambiguous_anchor)
    code, out = run(capsys, "transfer", "--from", "excel", "--chart", "사인 2주기", "--to", "word", "--anchor", "계획")
    assert code == 1 and out["error_type"] == "ambiguous" and out["candidates"] == ["1: 계획 A", "2: 계획 B"]
    assert len(calls) == 1                                                # 다른 방식으로 넘어가지 않음


# ---------------------------------------------------------------------------
# undo
# ---------------------------------------------------------------------------

def test_undo_closed_file_restores_backs_up_and_consumes(tmp_path, monkeypatch, capsys):
    fake, _ = make_win32com()
    monkeypatch.setattr(off, "win32com", fake)
    monkeypatch.setattr(off, "pythoncom", None)
    monkeypatch.setattr(off, "match_open", lambda kind, target, docs=None, option="--target": None)
    orig = tmp_path / "a.xlsx"
    orig.write_bytes(b"current")
    off.save_snapshot("excel", FakeWorkbook(str(orig)), "excel", "write")      # 스냅샷 내용: snapshot-of-a.xlsx
    code, out = run(capsys, "undo", "--list")
    assert out["snapshots"][0]["by"] == "excel write"
    assert out["snapshots"][0]["command"].endswith("undo --pick 1")
    code, out = run(capsys, "undo", "--list", "--app", "excel")
    assert out["snapshots"][0]["command"].endswith("undo --pick 1 --app excel")     # 같은 필터로 고르기
    code, out = run(capsys, "undo")
    assert code == 0 and out["restored_as"] == "file_replaced" and out["agent_directive"].startswith("VERIFIED:")   # 스냅샷과 비교 (R1)
    assert orig.read_bytes() == b"snapshot-of-a.xlsx"
    assert open(out["pre_undo_backup"], "rb").read() == b"current"
    code, out = run(capsys, "undo")
    assert code == 0 and out["restored"] is False                         # 소비되어 더 없음


def test_undo_pick_out_of_range(tmp_path, monkeypatch, capsys):
    orig = tmp_path / "a.xlsx"
    orig.write_bytes(b"x")
    off.save_snapshot("excel", FakeWorkbook(str(orig)))
    code, out = run(capsys, "undo", "--pick", "5")
    assert code == 1 and "--list" in out["message"]


# --- 저장한 적 없는 문서 undo (시나리오 12: 새 워드 문서에 차트를 넣은 뒤 "방금 한 거 되돌려줘") ---

class _UntitledWordDoc:
    def __init__(self, name, xml):
        self.Name = name
        self.FullName = name                       # 저장 안 된 문서는 FullName == Name (절대 경로 아님)
        self.Saved = False
        self.WordOpenXML = xml
        self.inserted = []
        self.activated = 0

    def Range(self, *a):
        doc = self

        class _R:
            def InsertXML(self, xml):
                doc.inserted.append(xml)
                doc.WordOpenXML = xml
        return _R()

    def Activate(self):
        self.activated += 1


class _Docs:
    def __init__(self, items):
        self.items = list(items)
        self.added = []

    @property
    def Count(self):
        return len(self.items)

    def __call__(self, i):
        return self.items[i - 1]

    def Add(self, Template=None):
        d = _UntitledWordDoc("문서2", "<new/>")
        self.added.append(Template)
        self.items.append(d)
        return d


@pytest.fixture
def untitled_env(monkeypatch):
    fake, _ = make_win32com()
    monkeypatch.setattr(off, "win32com", fake)
    monkeypatch.setattr(off, "pythoncom", None)
    apps = []
    monkeypatch.setattr(off, "find_apps", lambda kind: list(apps))
    return apps


def test_undo_unsaved_word_restores_in_same_window(untitled_env, capsys):
    doc = _UntitledWordDoc("문서1", "<before/>")
    app = type("WordApp", (), {})()
    app.Documents = _Docs([doc])
    untitled_env.append(app)
    off.save_snapshot("word", doc, "office", "transfer")                 # flat_xml 스냅샷
    doc.WordOpenXML = "<with-chart/>"                                    # 차트가 들어간 상태
    code, out = run(capsys, "undo")
    assert code == 0 and out["restored_as"] == "restored_in_place" and out["name"] == "문서1"
    assert doc.inserted == ["<before/>"] and doc.activated == 1          # 같은 문서의 내용을 바꿈
    assert app.Documents.Count == 1 and app.Documents.added == []          # 새 창을 열지 않음
    assert open(out["pre_undo_backup"], encoding="utf-8").read() == "<with-chart/>"
    assert "같은 창" in out["preview"] and "같은 창" in out["agent_directive"]
    assert out["agent_directive"].startswith("VERIFIED:")


def _close_removes(wb, coll):
    def close(SaveChanges=None):
        wb.rec("Close", SaveChanges)
        coll.items.remove(wb)
    wb.Close = close


def test_undo_unsaved_excel_replaces_and_closes_modified(untitled_env, capsys):
    app = FakeExcelApp()
    wb = FakeWorkbook("")                                                # Book1, 저장 안 됨
    app.Workbooks.items.append(wb)
    _close_removes(wb, app.Workbooks)
    untitled_env.append(app)
    off.save_snapshot("excel", wb, "excel", "write")
    code, out = run(capsys, "undo")
    assert code == 0 and out["restored_as"] == "replaced_untitled_document"
    assert app.Workbooks.Count == 1 and app.Workbooks(1) is not wb        # 바뀐 통합문서는 닫히고 하나만 남음
    assert ("Close", False) in wb.calls                                  # 저장하지 않고 닫음
    assert app.Workbooks(1).template.endswith(".xlsx") and os.path.exists(out["pre_undo_backup"])
    assert "저장하지 않고 닫았습니다" in out["preview"]


def test_undo_unsaved_ppt_replaces_and_closes_modified(untitled_env, capsys):
    from fake_office import FakePptApp, FakePresentation
    app = FakePptApp()
    prs = FakePresentation("프레젠테이션1")
    prs.Saved = 0
    app.Presentations.items.append(prs)
    untitled_env.append(app)
    off.save_snapshot("ppt", prs, "ppt", "write")
    code, out = run(capsys, "undo")
    assert code == 0 and out["restored_as"] == "replaced_untitled_document"
    assert ("Close",) in prs.calls and prs.Saved == -1                   # 묻지 않고 닫힘 (백업은 따로 보관)
    assert app.Presentations.open_args[1:] == (0, -1, -1) and os.path.exists(out["pre_undo_backup"])


def test_undo_unsaved_doc_not_open_opens_new_untitled(untitled_env, capsys):
    app = FakeExcelApp()
    wb = FakeWorkbook("")
    off.save_snapshot("excel", wb, "excel", "write")
    other = FakeWorkbook("")
    other.Name = "Book2"
    app.Workbooks.items.append(other)                                    # Book1은 이미 닫힘
    untitled_env.append(app)
    code, out = run(capsys, "undo")
    assert code == 0 and out["restored_as"] == "new_untitled_document" and out["pre_undo_backup"] == ""
    assert app.Workbooks.Count == 2 and "열려 있지 않아" in out["preview"]
    assert other.calls == []                                             # 다른 문서는 건드리지 않음


def test_undo_unsaved_same_name_twice_does_not_guess(untitled_env, capsys):
    app1, app2 = FakeExcelApp(), FakeExcelApp()
    a, b = FakeWorkbook(""), FakeWorkbook("")                            # 두 Excel 창의 Book1
    app1.Workbooks.items.append(a)
    app2.Workbooks.items.append(b)
    untitled_env.extend([app1, app2])
    off.save_snapshot("excel", a, "excel", "write")
    code, out = run(capsys, "undo")
    assert code == 0 and out["restored_as"] == "new_untitled_document"
    assert "2개" in out["preview"] and "그대로 열려 있습니다" in out["preview"]
    assert not any(c[0] == "Close" for c in a.calls + b.calls)           # 어느 것도 닫지 않음


def test_windows_lists_documents(monkeypatch, capsys):
    monkeypatch.setattr(off, "open_documents_summary",
                        lambda: [{"app": "Excel", "name": "a.xlsx", "path": "", "saved": True, "active": True}])
    monkeypatch.setattr(off, "running_office_caps", lambda: {})
    code, out = run(capsys, "windows")
    assert code == 0 and out["preview"] == "Excel 1개" and out["office"] == {}


def test_windows_reports_office_version_and_missing_pdf(monkeypatch, capsys):
    monkeypatch.setattr(off, "open_documents_summary",
                        lambda: [{"app": "Word", "name": "a.docx", "path": "", "saved": True, "active": True}])
    monkeypatch.setattr(off, "running_office_caps",
                        lambda: {"Word": {"version": "12.0", "major": 12, "pdf_export": False}})
    code, out = run(capsys, "windows")
    assert code == 0 and out["office"]["Word"]["pdf_export"] is False
    assert "Office 12.0" in out["preview"] and "PDF 내보내기 불가: Word" in out["preview"]


# --- 점검 반영: 되돌릴 지점이 없는 직전 작업, 지점 뒤에 다시 저장된 파일 ---

def _undo_env(tmp_path, monkeypatch):
    fake, _ = make_win32com()
    monkeypatch.setattr(off, "win32com", fake)
    monkeypatch.setattr(off, "pythoncom", None)
    monkeypatch.setattr(off, "match_open", lambda kind, target, docs=None, option="--target": None)
    orig = tmp_path / "a.xlsx"
    orig.write_bytes(b"current")
    return orig


def test_undo_stops_at_write_without_snapshot(tmp_path, monkeypatch, capsys):
    orig = _undo_env(tmp_path, monkeypatch)
    off.save_snapshot("excel", FakeWorkbook(str(orig)), "excel", "write")          # 월요일의 지점
    off.save_barrier("excel", FakeWorkbook(str(orig)), "스냅샷 실패", "excel", "write")  # 금요일: 지점 없음
    code, out = run(capsys, "undo")
    assert code == 1 and out["error_type"] == "needs_user_decision" and orig.read_bytes() == b"current"
    code, out = run(capsys, "undo", "--list")
    assert out["snapshots"][0]["undo_available"] is False and "undo_available" not in out["snapshots"][1]
    code, out = run(capsys, "undo", "--pick", "1")
    assert code == 1 and orig.read_bytes() == b"current"
    code, out = run(capsys, "undo", "--pick", "2")                                 # 사용자가 고른 앞 지점
    assert code == 0 and orig.read_bytes() == b"snapshot-of-a.xlsx"


def test_undo_refuses_when_file_saved_after_snapshot(tmp_path, monkeypatch, capsys):
    orig = _undo_env(tmp_path, monkeypatch)
    off.save_snapshot("excel", FakeWorkbook(str(orig)), "excel", "write")
    later = os.path.getmtime(orig) + 3600
    os.utime(orig, (later, later))                                               # 한 시간 뒤 사용자가 저장
    meta = off.list_snapshots()[0]
    assert meta["created_ts"] < later
    code, out = run(capsys, "undo")
    assert code == 1 and out["error_type"] == "needs_user_decision" and "--overwrite" in out["agent_directive"]
    assert orig.read_bytes() == b"current"
    code, out = run(capsys, "undo", "--overwrite")
    assert code == 0 and orig.read_bytes() == b"snapshot-of-a.xlsx"
