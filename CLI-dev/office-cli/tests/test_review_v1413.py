"""1.4.13 회귀 테스트: 4차 점검에서 찾은 오류 (U1~U4 되돌리기 지점, I1~I5 reader, O1 outlook, X1~X3 엑셀·공통).

항목마다 실패 상황을 재현하고 고친 동작을 확인합니다. 가짜 COM(fake_*.py)과 실제 파일(openpyxl, pymupdf 등)을 씁니다.
"""
import datetime
import json
import os
import time
import types

import pytest

import _office as off
import _runtime as rt
import excel
import outlook
import reader
from fake_excel import FakeExcelWorkbook, FakeXlApp
from fake_office import FakeExcelApp, make_win32com


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


def _points():
    return sorted(s["meta_file"] for s in off.list_snapshots("excel"))


def _sheet_names(wb):
    return [str(wb.Worksheets(i).Name) for i in range(1, wb.Worksheets.Count + 1)]


# ---------------------------------------------------------------------------
# U1: write --new-sheet는 모든 확인이 끝난 뒤에 시트를 만듦 (확인에서 멈추면 시트도 지점도 남지 않음)
# ---------------------------------------------------------------------------

def test_u1_new_sheet_is_not_created_when_input_check_fails(xl, capsys):
    _, wb = xl
    code, out = _run(excel, capsys, "write", "--range", "A1", "--value", "1")
    assert code == 0, out
    before = _points()
    code, out = _run(excel, capsys, "write", "--new-sheet", "--sheet", "새시트", "--formula", "5")
    assert code == 1 and "[추천] --value 5" in out["message"], out
    assert "새시트" not in _sheet_names(wb) and _points() == before
    code, out = _run(excel, capsys, "write", "--new-sheet", "--sheet", "새시트", "--value", "5")   # 안내대로 다시 실행
    assert code == 0 and out["sheet"] == "새시트", out
    assert len(_points()) == len(before) + 1


def test_u1_new_sheet_with_bad_range_creates_nothing(xl, capsys):
    _, wb = xl
    code, out = _run(excel, capsys, "write", "--new-sheet", "--sheet", "새시트", "--range", "A1", "--formula", "=A1", "=B1")
    assert code == 1 and "--range" in out["message"]
    assert "새시트" not in _sheet_names(wb) and _points() == []


# ---------------------------------------------------------------------------
# U2: style은 --highlight 등 입력을 서식을 지우기 전에 확인
# ---------------------------------------------------------------------------

def test_u2_style_bad_highlight_does_not_clear_formats(xl, capsys):
    _, wb = xl
    ws = wb.Worksheets(1)
    code, out = _run(excel, capsys, "style", "--range", "A1:B2", "--clear", "--highlight", "bad")
    assert code == 1 and "--highlight" in out["message"]
    assert not [x for x in ws.log if x[0] == "ClearFormats"] and _points() == []


# ---------------------------------------------------------------------------
# U3: 이번 지점은 문서를 바꾸지 않았을 때만 지움
# ---------------------------------------------------------------------------

def test_u3_point_is_kept_when_document_changed_before_input_error(xl):
    with pytest.raises(rt.InputError):
        with off.open_target("excel", "active", readonly=False, tool="excel", command="write") as t:
            t.sheet("중간", create=True)                    # 문서를 바꿈 (t.changed)
            raise rt.InputError("바꾼 뒤 멈춤")
    assert len(_points()) == 1


def test_u3_dirty_document_keeps_point_unless_input_error(xl):
    _, wb = xl
    wb.Saved = False                                        # 원래 저장 안 된 문서: 저장 상태로는 바뀐 것을 알 수 없음
    with pytest.raises(RuntimeError):
        with off.open_target("excel", "active", readonly=False, tool="excel", command="write"):
            raise RuntimeError("COM 오류 (중간까지 바꿨는지 모름)")
    assert len(_points()) == 1
    with pytest.raises(rt.InputError):
        with off.open_target("excel", "active", readonly=False, tool="excel", command="write"):
            raise rt.InputError("바꾸기 전 입력 오류")
    assert len(_points()) == 1                              # 이번 실행의 지점만 지움


def _file_session(monkeypatch, tmp_path, save=None):
    path = tmp_path / "닫힌.xlsx"
    path.write_bytes(b"original")
    _com(monkeypatch)
    real_open = off._open_file

    def opened(kind, app, p, readonly):
        doc = real_open(kind, app, p, readonly)
        if save is not None:
            doc.Save = lambda: save(doc, p)
        return doc
    monkeypatch.setattr(off, "_open_file", opened)
    monkeypatch.setattr(off, "_new_instance", lambda kind: FakeExcelApp())
    return path


def test_u3_closed_file_save_failure_drops_point(monkeypatch, tmp_path):
    def fail(doc, p):
        raise OSError("디스크 가득 참")
    path = _file_session(monkeypatch, tmp_path, fail)
    with pytest.raises(OSError):
        with off.open_target("excel", str(path), readonly=False, tool="excel", command="write") as t:
            t.doc.Saved = False
            t.set_focus()
    assert _points() == [] and path.read_bytes() == b"original"


def test_u3_closed_file_error_after_change_drops_point_and_success_keeps_it(monkeypatch, tmp_path):
    def write(doc, p):
        with open(p, "wb") as f:
            f.write(b"changed by the tool")
        doc.Saved = True
    path = _file_session(monkeypatch, tmp_path, write)
    with pytest.raises(RuntimeError):
        with off.open_target("excel", str(path), readonly=False, tool="excel", command="write") as t:
            t.doc.Saved = False
            t.set_focus()
            raise RuntimeError("바꾼 뒤 COM 오류: 저장하지 않고 닫음")
    assert _points() == []                                  # 디스크 파일은 그대로라 지점은 바뀐 것 없는 상태
    with off.open_target("excel", str(path), readonly=False, tool="excel", command="write") as t:
        t.doc.Saved = False
        t.set_focus()
    assert len(_points()) == 1 and path.read_bytes() == b"changed by the tool"


# ---------------------------------------------------------------------------
# U4: 오래된 지점 정리는 명령이 바꾼 뒤에만 (거절된 실행이 진짜 옛 지점을 밀어내지 않음)
# ---------------------------------------------------------------------------

def test_u4_refused_run_does_not_evict_old_point(xl, capsys):
    for i in range(off.SNAPSHOT_KEEP):
        code, out = _run(excel, capsys, "write", "--range", f"A{i + 1}", "--value", f"v{i}")
        assert code == 0, out
    before = _points()
    assert len(before) == off.SNAPSHOT_KEEP
    code, out = _run(excel, capsys, "write", "--sheet", "없는시트", "--range", "A1", "--value", "c")
    assert code == 1
    assert _points() == before
    code, out = _run(excel, capsys, "write", "--range", "C1", "--value", "새 값")        # 성공한 실행은 정리
    assert code == 0 and len(_points()) == off.SNAPSHOT_KEEP and before[0] not in _points()


# ---------------------------------------------------------------------------
# I1: 색인을 잠깐 읽지 못해도 빈 색인으로 덮어쓰지 않음, 손상된 색인은 캐시 문서로 다시 만듦
# ---------------------------------------------------------------------------

def _docs(tmp_path, n=3):
    folder = tmp_path / "docs"
    folder.mkdir()
    for i in range(n):
        (folder / f"d{i}.txt").write_text(f"문서 {i} 내용 토크", encoding="utf-8")
    return folder


def _index_json():
    with open(reader._index_path(), encoding="utf-8") as f:
        return json.load(f)


def test_i1_transient_index_lock_is_retried_not_emptied(tmp_path, capsys, monkeypatch):
    folder = _docs(tmp_path)
    code, out = _run(reader, capsys, "index", "--folder", str(folder))
    assert code == 0 and len(_index_json()) == 3
    real = rt.read_text_file
    state = {"n": 0}

    def flaky(p):
        if p.endswith("index.json") and state["n"] == 0:
            state["n"] += 1
            raise PermissionError(13, "sharing violation", p)
        return real(p)
    monkeypatch.setattr(rt, "read_text_file", flaky)
    monkeypatch.setattr(time, "sleep", lambda s: None)
    new = folder / "new.txt"
    new.write_text("새 문서", encoding="utf-8")
    reader.load_document_status(str(new))
    assert state["n"] == 1 and len(_index_json()) == 4


def test_i1_unreadable_index_stops_without_overwriting(tmp_path, capsys, monkeypatch):
    folder = _docs(tmp_path)
    _run(reader, capsys, "index", "--folder", str(folder))
    before = _index_json()
    real = rt.read_text_file

    def locked(p):
        if p.endswith("index.json"):
            raise PermissionError(13, "sharing violation", p)
        return real(p)
    monkeypatch.setattr(rt, "read_text_file", locked)
    monkeypatch.setattr(time, "sleep", lambda s: None)
    new = folder / "new.txt"
    new.write_text("새 문서", encoding="utf-8")
    doc, _ = reader.load_document_status(str(new))           # 읽기는 계속, 색인은 그대로
    assert doc["units"] and _index_json() == before
    code, out = _run(reader, capsys, "index", "--all")
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:") and "색인" in out["message"]
    assert _index_json() == before


def test_i1_corrupt_index_is_rebuilt_from_cache(tmp_path, capsys):
    folder = _docs(tmp_path)
    _run(reader, capsys, "index", "--folder", str(folder))
    with open(reader._index_path(), "w", encoding="utf-8") as f:
        f.write("{깨진 색인")
    code, out = _run(reader, capsys, "search", "--query", "토크")
    assert code == 0 and out["total_matches"] == 3, out
    assert any("색인 파일 손상" in r for r in out.get("recovered", [])), out
    assert len(_index_json()) == 3


# ---------------------------------------------------------------------------
# I2/I3: 짧은 글(도장, Bates 번호)만 남은 스캔 쪽도 표시, 모든 쪽이 스캔이면 OCR 필요로 UNVERIFIED
# ---------------------------------------------------------------------------

def _pix(pymupdf):
    pix = pymupdf.Pixmap(pymupdf.csRGB, pymupdf.IRect(0, 0, 40, 56), False)
    pix.clear_with(200)
    return pix


def test_i2_scanned_page_with_stamp_text_is_flagged(tmp_path, capsys):
    import pymupdf
    doc = pymupdf.open()
    for i in range(1, 5):
        p = doc.new_page()
        if i == 3:
            p.insert_image(pymupdf.Rect(30, 30, 565, 812), pixmap=_pix(pymupdf))
            p.insert_text((400, 400), "RECEIVED", fontsize=9)               # 스캔 위에 찍은 도장 글자
        else:
            p.insert_text((72, 200), f"Body text page {i} torque {i * 10} Nm.", fontsize=11)
    path = tmp_path / "stamp.pdf"
    doc.save(str(path))
    code, out = _run(reader, capsys, "read", "--file", str(path))
    assert code == 0 and out["meta"]["pages_without_text"] == [3] and not out["meta"].get("ocr_needed"), out
    assert reader.PARSER_VERSION == 8


def test_i2_i3_single_page_scan_with_bates_number_needs_ocr(tmp_path, capsys):
    import pymupdf
    doc = pymupdf.open()
    p = doc.new_page()
    p.insert_image(p.rect, pixmap=_pix(pymupdf))
    p.insert_text((450, 820), "ABC0000001", fontsize=9)       # 한 쪽짜리라 머리글로 뺄 수 없는 Bates 번호
    path = tmp_path / "scan1.pdf"
    doc.save(str(path))
    code, out = _run(reader, capsys, "read", "--file", str(path))
    assert code == 0 and out["meta"].get("ocr_needed") and out["meta"]["pages_without_text"] == [1], out
    assert out["agent_directive"].startswith("UNVERIFIED:") and "스캔" in out["agent_directive"]


def test_i3_image_only_pdf_read_says_ocr_needed(tmp_path, capsys):
    import pymupdf
    doc = pymupdf.open()
    for _ in range(2):
        p = doc.new_page()
        p.insert_image(p.rect, pixmap=_pix(pymupdf))
    path = tmp_path / "scan2.pdf"
    doc.save(str(path))
    code, out = _run(reader, capsys, "read", "--file", str(path))
    assert code == 0 and out["count"] == 0 and out["meta"].get("ocr_needed"), out
    assert out["agent_directive"].startswith("UNVERIFIED:") and "OCR" in out["agent_directive"]
    assert "읽을 글이 없습니다" not in out["agent_directive"]


def test_i2_ocr_layer_with_little_text_is_not_flagged(tmp_path, capsys):
    import pymupdf
    doc = pymupdf.open()
    for i in range(1, 3):
        p = doc.new_page()
        p.insert_image(p.rect, pixmap=_pix(pymupdf))
        p.insert_text((72, 200), f"OCR {i}", fontsize=11, render_mode=3)      # 짧아도 보이지 않는 OCR 글자 층
    path = tmp_path / "ocr.pdf"
    doc.save(str(path))
    code, out = _run(reader, capsys, "read", "--file", str(path))
    assert code == 0 and "pages_without_text" not in out["meta"] and not out["meta"].get("ocr_needed"), out


# ---------------------------------------------------------------------------
# I4: search --full 저장 안내는 엑셀 칸 상한보다 긴 단락이 있으면 .md
# ---------------------------------------------------------------------------

def test_i4_full_search_with_huge_unit_suggests_md(tmp_path, capsys):
    f = tmp_path / "huge.txt"
    f.write_text("앞 단락 토크\n\n" + "토크 " + "가" * 33000 + "\n", encoding="utf-8")
    code, out = _run(reader, capsys, "search", "--query", "토크", "--file", str(f), "--full")
    assert code == 0 and out["truncated"], out
    saves = [c["command"] for c in out["next_commands"] if "--save" in c["command"]]
    assert saves and "근거.md" in saves[0] and "xlsx" not in saves[0], saves


# ---------------------------------------------------------------------------
# I5: 오래된 잠금은 고유 이름으로 옮긴 뒤 지움 (다른 실행의 새 잠금은 지우지 않음)
# ---------------------------------------------------------------------------

def test_i5_stale_lock_removal_keeps_fresh_lock(tmp_path):
    lock = tmp_path / "index.lock"
    lock.write_text("1")
    old = time.time() - reader.INDEX_LOCK_STALE - 10
    os.utime(lock, (old, old))
    assert reader._remove_stale_lock(str(lock)) is True and not lock.exists()
    lock.write_text("2")                                     # 그 사이 다른 실행이 새로 만든 잠금
    assert reader._remove_stale_lock(str(lock)) is False and lock.read_text() == "2"
    assert not [p.name for p in tmp_path.iterdir() if ".stale." in p.name]
    assert reader.INDEX_LOCK_WAIT > sum(0.5 * (i + 1) for i in range(6))   # 잠근 쪽의 retry_fs 재시도보다 길게


# ---------------------------------------------------------------------------
# O1: --save-attach에서 저장하지 못한 첨부가 있으면 VERIFIED가 아님
# ---------------------------------------------------------------------------

class _Att:
    def __init__(self, name, fail=False, vanish=False):
        self.FileName = name
        self.fail = fail
        self.vanish = vanish
        self.PropertyAccessor = self

    def GetProperty(self, tag):
        return False

    def SaveAsFile(self, path):
        if self.fail:
            raise OSError("잠김")
        if not self.vanish:                                  # 백신이 바로 지운 경우: 오류 없이 파일이 없음
            open(path, "w").close()


class _Atts:
    def __init__(self, items):
        self.items = items

    @property
    def Count(self):
        return len(self.items)

    def Item(self, i):
        return self.items[i - 1]


def _mail(atts):
    return types.SimpleNamespace(Subject="견적", Body="본문", SenderName="a", SenderEmailAddress="a@x", To="", CC="",
                                 ReceivedTime=datetime.datetime(2026, 1, 1), EntryID="X", UnRead=False,
                                 Attachments=_Atts(atts))


@pytest.mark.parametrize("att", [_Att("b.pdf", fail=True), _Att("c.pdf", vanish=True)])
def test_o1_failed_attachment_save_is_unverified(monkeypatch, att):
    item = _mail([_Att("a.xlsx"), att])
    monkeypatch.setattr(outlook, "get_outlook", lambda: None)
    monkeypatch.setattr(outlook, "get_namespace", lambda app: types.SimpleNamespace(GetItemFromID=lambda i: item))
    out = outlook.cmd_read(types.SimpleNamespace(id="X", save=None, save_attach=True))
    assert out["agent_directive"].startswith("UNVERIFIED:") and att.FileName in out["agent_directive"], out
    assert len(out["saved_attachments"]) == 1


def test_o1_all_attachments_saved_is_verified(monkeypatch):
    item = _mail([_Att("a.xlsx")])
    monkeypatch.setattr(outlook, "get_outlook", lambda: None)
    monkeypatch.setattr(outlook, "get_namespace", lambda app: types.SimpleNamespace(GetItemFromID=lambda i: item))
    out = outlook.cmd_read(types.SimpleNamespace(id="X", save=None, save_attach=True))
    assert out["agent_directive"].startswith("VERIFIED:") and len(out["saved_attachments"]) == 1


def test_outlook_new_checks_attachment_before_creating_mail(monkeypatch, tmp_path):
    created = []
    monkeypatch.setattr(outlook, "get_outlook", lambda: types.SimpleNamespace(CreateItem=lambda k: created.append(k)))
    args = types.SimpleNamespace(to="a@x", cc=None, subject="s", body="b", body_file=None,
                                 attach=str(tmp_path / "없는파일.pdf"), send=False)
    with pytest.raises(rt.InputError):
        outlook.cmd_new(args)
    assert created == []


# ---------------------------------------------------------------------------
# X1: 닫힌 xlsx의 큰 범위는 데이터 끝까지만 채워 읽음 (상한을 넘으면 [추천] 범위)
# ---------------------------------------------------------------------------

def _xlsx(path, rows, cols):
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "데이터"
    for r in range(1, rows + 1):
        ws.append([f"r{r}c{c}" for c in range(1, cols + 1)])
    wb.save(path)


def test_x1_huge_explicit_range_is_clipped_to_data(tmp_path, capsys, monkeypatch):
    p = tmp_path / "작은.xlsx"
    _xlsx(p, 3, 2)
    monkeypatch.setattr(excel, "CLOSED_MAX_CELLS", 1000, raising=False)
    code, out = _run(excel, capsys, "read", "--target", str(p), "--range", "A1:Z2000")
    assert code == 0 and out["rows"] == 3 and out["cols"] == 2, out
    assert any("데이터가 있는" in r for r in out.get("recovered", [])), out


def test_x1_whole_sheet_range_does_not_run_out_of_memory(tmp_path, capsys):
    p = tmp_path / "전체.xlsx"
    _xlsx(p, 3, 2)
    code, out = _run(excel, capsys, "read", "--target", str(p), "--range", "A1:XFD1048576")
    assert code == 0 and out["rows"] == 3 and out["range"] == "A1:B3", out


def test_x1_range_too_big_even_after_clip_recommends_narrower(tmp_path, capsys, monkeypatch):
    p = tmp_path / "큰.xlsx"
    _xlsx(p, 6, 5)
    monkeypatch.setattr(excel, "CLOSED_MAX_CELLS", 10, raising=False)
    code, out = _run(excel, capsys, "read", "--target", str(p), "--range", "A1:Z100")
    assert code == 1 and "[추천] --range A1:E2" in out["message"], out


# ---------------------------------------------------------------------------
# X2: 기본 저장 이름이 같은 초에 겹치면 번호를 붙인 새 이름 (덮어쓰지 않음, recovered로 알림)
# ---------------------------------------------------------------------------

def test_x2_default_save_name_in_same_second_is_unique(monkeypatch):
    monkeypatch.setattr(rt, "timestamp", lambda: "20261102_101500")
    rt.RECOVERED.clear()
    first = rt.resolve_output_path(None, "표", ".csv")
    with open(first, "w", encoding="utf-8") as f:
        f.write("a\n")
    second = rt.resolve_output_path(None, "표", ".csv")
    assert first != second and not os.path.exists(second)
    assert os.path.basename(second) == "표_20261102_101500_2.csv"
    assert any("같은 초" in r for r in rt.RECOVERED)
    rt.RECOVERED.clear()


# ---------------------------------------------------------------------------
# X3: 첫 행이 표시 상한을 넘을 때 '긴 값'과 '뺀 열'을 따로 셈
# ---------------------------------------------------------------------------

def test_x3_preview_separates_long_values_and_dropped_columns(tmp_path, capsys):
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["가" * 600 for _ in range(40)])
    ws.append([1] * 40)
    p = tmp_path / "넓은.xlsx"
    wb.save(p)
    code, out = _run(excel, capsys, "read", "--target", str(p))
    assert code == 0 and out["truncated"], out
    assert "긴 값 15칸 잘림" in out["preview"] and "뒤쪽 25열 생략" in out["preview"], out["preview"]
    assert "25열" in out["truncated_reason"]
