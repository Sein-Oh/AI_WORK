"""office.py와 _office.py: transfer(차트·범위 -> 워드/PPT/엑셀, 전달 방식 대체 순서), undo와 스냅샷,
windows, 문서 세션(사용자 창 미간섭, 읽기 전용, 파일 열기·저장 규칙).
"""
import contextlib
import json
import os
import time
import types

import pytest

import _office as off
import _runtime as rt
import excel
import fake_office
import office
import ppt
from fake_excel import FakeExcelWorkbook, FakeXlApp
from fake_office import FakeExcelApp, FakePptApp, FakeWordDoc, FakeWorkbook, FakeWorkbooks, make_win32com
from fake_ppt import FakePresentation
from fake_word import FakeWordDocument
from office_helpers import option_value, split_command, error_log_text, formula_xlsx, run_as_tool, run_main, run_office, SECRET, sheet_names, stub_com, undo_points


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


def _close_removes(wb, coll):
    def close(SaveChanges=None):
        wb.rec("Close", SaveChanges)
        coll.items.remove(wb)
    wb.Close = close


def _undo_env(tmp_path, monkeypatch):
    fake, _ = make_win32com()
    monkeypatch.setattr(off, "win32com", fake)
    monkeypatch.setattr(off, "pythoncom", None)
    monkeypatch.setattr(off, "match_open", lambda kind, target, docs=None, option="--target": None)
    orig = tmp_path / "a.xlsx"
    orig.write_bytes(b"current")
    return orig


@pytest.fixture
def fake_com(monkeypatch):
    """win32com을 가짜로 바꾸고, 열린 문서/활성 문서는 테스트가 지정합니다."""
    state = {"open": {"excel": [], "word": [], "ppt": []}, "active": {}}
    monkeypatch.setattr(off, "pythoncom", None)
    monkeypatch.setattr(off, "list_open", lambda kind: state["open"][kind])

    def get_active(kind):
        if kind not in state["active"]:
            raise rt.InputError("열려 있는 문서가 없습니다.")
        return state["active"][kind]
    monkeypatch.setattr(off, "get_active", get_active)

    def install(**kw):
        fake, calls = make_win32com(**kw)
        monkeypatch.setattr(off, "win32com", fake)
        state["calls"] = calls
        return calls
    state["install"] = install
    install()
    return state


class _Doc2007:
    """Word 2007 문서: SaveAs2가 없음 (pywin32 동적 디스패치는 AttributeError)."""
    def __init__(self):
        self.calls = []

    def SaveAs(self, path, **kw):
        self.calls.append(("SaveAs", path, kw))


class _Doc2010(_Doc2007):
    def SaveAs2(self, path, **kw):
        self.calls.append(("SaveAs2", path, kw))


def _xlsx_created(path, created):
    import datetime as dt
    import openpyxl
    wb = openpyxl.Workbook()
    wb.properties.created = dt.datetime(*created)
    wb.save(path)


class _WordApp:
    """Documents.Open 인자를 기록하고, fail이 있으면 그 오류를 냅니다."""

    def __init__(self, fail=None):
        self.kw = None
        self.fail = fail
        self.Documents = self

    def Open(self, path, **kw):
        self.kw = kw
        if self.fail:
            raise self.fail
        return "doc"


class _Created(Exception):
    pass


def _file_session(monkeypatch, tmp_path, save=None):
    path = tmp_path / "닫힌.xlsx"
    path.write_bytes(b"original")
    stub_com(monkeypatch)
    real_open = off._open_file

    def opened(kind, app, p, readonly):
        doc = real_open(kind, app, p, readonly)
        if save is not None:
            doc.Save = lambda: save(doc, p)
        return doc
    monkeypatch.setattr(off, "_open_file", opened)
    monkeypatch.setattr(off, "_new_instance", lambda kind: FakeExcelApp())
    return path


def _volatile_open(monkeypatch):
    orig_open = FakeWorkbooks.Open

    def vopen(self, p, **kw):
        wb = orig_open(self, p, **kw)
        wb.Saved = False                     # Excel: 휘발성 수식(NOW, RAND)이 열 때 다시 계산되어 저장 안 된 상태

        def save():
            wb.rec("Save")
            wb.Saved = True
            with open(p, "wb") as f:
                f.write(b"resaved")
        wb.Save = save
        return wb
    monkeypatch.setattr(FakeWorkbooks, "Open", vopen)
    monkeypatch.setattr(off, "_new_instance", lambda kind: FakeExcelApp())


@pytest.mark.parametrize("value", [r"C:\work\보고서 1.xlsx", r"C:\work\a.xlsx", r"\\server\공유\b.docx", 'say "hi"',
                                   "A=1,B=2", "", r"C:\a\b"])
def test_split_command_reads_quote_arg_like_windows(value):
    """테스트 도우미 split_command는 도구의 quote_arg로 만든 명령을 Windows처럼 나눈다 (shlex posix는 백슬래시를 지움)."""
    cmd = f"python {rt.quote_arg(r'C:\CLI\office-cli\excel.py')} read --target {rt.quote_arg(value)} --sheet S1"
    assert split_command(cmd) == ["python", r"C:\CLI\office-cli\excel.py", "read", "--target", value, "--sheet", "S1"]


def _recommended_argv(message):
    """message의 '[추천] python .../office.py <인자>'에서 인자만."""
    cmd = message.split("[추천] ", 1)[1]
    parts = split_command(cmd)
    assert parts[0] == "python" and parts[1].endswith("office.py")
    return parts[2:]


@pytest.fixture
def env(monkeypatch, tmp_path):
    """엑셀 차트/범위 하나, 워드 문서 열림, PowerPoint는 holder['prs']가 None이면 열린 문서 없음."""
    fake, _ = make_win32com()
    monkeypatch.setattr(off, "win32com", fake)
    monkeypatch.setattr(off, "pythoncom", None)
    monkeypatch.setattr(off, "list_open", lambda kind: [])
    monkeypatch.setattr(off, "find_apps", lambda kind: [])
    app = FakeXlApp()
    wb = FakeExcelWorkbook(app)
    wd = FakeWordDocument("제목\r본문")
    holder = {"prs": None, "word": wd, "inserted": []}
    pp_app = type("A", (), {})()
    pp_app.ActiveWindow = type("W", (), {})()

    real_get_active = off.get_active

    def active(kind):
        if kind == "excel":
            return {"name": wb.Name, "path": wb.FullName, "app": app, "doc": wb}
        if kind == "word" and holder["word"] is not None:
            return {"name": wd.Name, "path": wd.FullName, "app": object(), "doc": wd}
        if kind == "ppt" and holder["prs"] is not None:
            return {"name": holder["prs"].Name, "path": holder["prs"].FullName, "app": pp_app, "doc": holder["prs"]}
        return real_get_active(kind)                # find_apps가 빈 목록: 실제 "열린 문서 없음" 오류
    monkeypatch.setattr(off, "get_active", active)
    chart = type("C", (), {"Name": "차트"})()
    monkeypatch.setattr(excel, "find_chart", lambda t, name, sheet=None: (None, chart))
    monkeypatch.setattr(excel, "_chart_title", lambda obj: "차트")
    png = tmp_path / "c.png"
    png.write_bytes(office.PNG_SIGNATURE + b"x" * 50)
    monkeypatch.setattr(office, "image_by_export", lambda obj, stem: str(png))
    monkeypatch.setattr(office, "image_by_redraw", lambda obj, stem: str(png))
    monkeypatch.setattr(office, "insert_picture_word",
                        lambda tgt, img, anchor, occurrence=None: holder["inserted"].append("word") or {"app": "Word"})
    monkeypatch.setattr(office, "insert_picture_ppt",
                        lambda tgt, img, slide, title: holder["inserted"].append("ppt") or {"app": "PowerPoint"})
    monkeypatch.setattr(rt, "load_config", lambda: {"transfer_method": {"word": ["export"], "ppt": ["export"]}})
    monkeypatch.setattr(office, "_excel_range_rows", lambda args: ([["항목", "값"], ["A", 1]], "Sheet1", None))
    return holder


CHART = ("transfer", "--from-app", "excel", "--chart", "차트", "--to-app", "word,ppt")


RANGE = ("transfer", "--from-app", "excel", "--range", "A1:B2", "--to-app", "word,ppt")


def _check_target_not_open(out):
    assert out["status"] == "error" and out["error_type"] == "target_not_open"
    assert "아무것도 바꾸지 않았습니다" in out["message"] and "PowerPoint" in out["message"]
    assert "--target " not in out["message"] + out["agent_directive"]          # transfer에 없는 옵션
    assert out["agent_directive"].startswith("STOP_RETRY:") and "실행하지 않았습니다" not in out["agent_directive"]
    argv = _recommended_argv(out["message"])
    assert argv[argv.index("--to-app") + 1] == "word"
    return argv


CH = ("transfer", "--from-app", "excel", "--chart", "차트")


TWO_SAME = [{"name": "a.docx", "path": "C:\\x\\a.docx", "doc": object(), "app": None},
            {"name": "a.docx", "path": "C:\\y\\a.docx", "doc": object(), "app": None}]


def _barrier(i):
    return {"app": "word", "name": f"n{i}.docx", "barrier": True, "reason": "새로 만든 파일", "created_at": "20261103_101500",
            "by": "word write", "meta_file": f"{i}.json"}


_real_insert_word = office.insert_picture_word


_real_relay = office.relay_chart


@pytest.fixture
def relay_spy(monkeypatch):
    calls = []
    monkeypatch.setattr(office, "relay_chart", lambda *a, **k: calls.append(a) or (_ for _ in ()).throw(RuntimeError("relay")))
    return calls


class _Shapes:
    def __init__(self):
        self.items = []

    @property
    def Count(self):
        return len(self.items)

    def __call__(self, i):
        return self.items[i - 1]

    def Paste(self):
        pic = type("Pic", (), {"Width": 400.0, "Height": 300.0, "Left": 0.0, "Top": 0.0, "LockAspectRatio": 0})()
        self.items.append(pic)
        return type("SR", (), {"Count": 1, "Item": lambda s, i: pic})()


class _Slides:
    def __init__(self, n):
        self.slides = [type("Slide", (), {"Shapes": _Shapes()})() for _ in range(n)]

    @property
    def Count(self):
        return len(self.slides)

    def __call__(self, i):
        return self.slides[i - 1]

    def Add(self, i, layout):
        s = type("Slide", (), {"Shapes": _Shapes()})()
        self.slides.insert(i - 1, s)
        return s


def _prs(n=2):
    return type("P", (), {"Slides": _Slides(n), "PageSetup": type("PS", (), {"SlideWidth": 960, "SlideHeight": 540})()})()


class _T:
    def __init__(self, kind, doc):
        self.kind, self.doc = kind, doc

    def set_focus(self, **k):
        pass

    def resolved(self):
        return {"app": off.APP_LABEL[self.kind]}


class _WordDoc:
    def __init__(self):
        self.InlineShapes = type("IS", (), {"Count": 1, "Item": lambda s, i: type("I", (), {
            "Range": type("R", (), {"CopyAsPicture": lambda s: None})()})()})()

    def Range(self, a, b):
        doc = self
        return type("R", (), {"InsertAfter": lambda s, t: None, "PasteSpecial": lambda s, **k: None,
                              "InlineShapes": doc.InlineShapes})()

    def Close(self, **k):
        pass


@pytest.fixture
def relay_env(chart_env, monkeypatch):
    """relay를 진짜로 부르는 상태. 클립보드 복사·임시 워드·문서 열기를 기록."""
    log = {"copied": [], "dispatch": [], "opened": []}
    chart = type("C", (), {"Name": "c", "CopyPicture": lambda self, *a: log["copied"].append(a)})()
    import excel
    import word
    monkeypatch.setattr(excel, "find_chart", lambda t, name, sheet=None: (None, chart))
    monkeypatch.setattr(word, "_insert_position", lambda doc, anchor, pos, occurrence=None: (0, "end"))
    monkeypatch.setattr(office, "relay_chart", _real_relay)
    monkeypatch.setattr(office, "method_order", lambda kind: ["relay"])     # --probe가 relay만 된다고 저장한 PC
    monkeypatch.setattr(office.time, "sleep", lambda s: None)
    prs, wdoc = _prs(2), _WordDoc()
    log["prs"] = prs
    monkeypatch.setattr(off, "get_active", lambda kind: {"name": "x", "path": "", "app": None,
                                                         "doc": prs if kind == "ppt" else wdoc})
    tapp = type("A", (), {"Documents": type("Ds", (), {"Add": lambda s, **k: _WordDoc()})(), "Quit": lambda s, **k: None})()
    monkeypatch.setattr(off.win32com.client, "DispatchEx", lambda n: log["dispatch"].append(n) or tapp, raising=False)

    @contextlib.contextmanager
    def ot(kind, target, **k):
        if kind != "excel":                               # 보내는 엑셀 문서는 빼고 받는 문서만
            log["opened"].append((kind, target))
        yield _T(kind, prs if kind == "ppt" else wdoc)
    monkeypatch.setattr(off, "open_target", ot)
    return log


class _Prs:
    def __init__(self, full, saved=-1, n=2):
        self.FullName = full
        self.Name = os.path.basename(full)
        self.Saved = saved
        self.Slides = _Slides(n)
        self.closed = 0
        self.PageSetup = types.SimpleNamespace(SlideWidth=960, SlideHeight=540)
        self.ReadOnly = 0

    def SaveCopyAs(self, p):
        with open(p, "wb") as f:
            f.write(b"snap")

    def Save(self):
        pass

    def Close(self):
        self.closed += 1


class _Pres:
    def __init__(self):
        self.items = []

    @property
    def Count(self):
        return len(self.items)

    def Open(self, path, *a):
        p = _Prs(path)
        self.items.append(p)
        return p


@pytest.fixture
def relay_real(monkeypatch):
    """open_target은 진짜, COM은 가짜. log: 클립보드 복사, 임시 워드 생성."""
    log = {"copied": 0, "dispatch": 0}
    app = types.SimpleNamespace(Presentations=_Pres(), DisplayAlerts=2, Quit=lambda: None)
    monkeypatch.setattr(off, "pythoncom", None)
    monkeypatch.setattr(office.time, "sleep", lambda s: None)

    def tmp_word(paste_ok):
        class D:
            InlineShapes = types.SimpleNamespace(Count=1 if paste_ok else 0, Item=lambda i: types.SimpleNamespace(
                Range=types.SimpleNamespace(CopyAsPicture=lambda: None)))

            def Range(self, a, b):
                return types.SimpleNamespace(PasteSpecial=lambda **k: None)

            def Close(self, **k):
                pass
        return types.SimpleNamespace(Documents=types.SimpleNamespace(Add=lambda **k: D()), Quit=lambda **k: None,
                                     Visible=True)

    def install(active=None, paste_ok=True):
        def dispatch_ex(name):
            log["dispatch"] += 1
            return tmp_word(paste_ok)
        monkeypatch.setattr(off, "win32com", types.SimpleNamespace(client=types.SimpleNamespace(
            DispatchEx=dispatch_ex, Dispatch=lambda n: app, GetActiveObject=lambda n: app)))

        def get_active(kind):
            if active is None:
                raise rt.UserDecisionError("열린 문서 없음", "x")
            return {"name": active.Name, "path": active.FullName, "app": app, "doc": active}
        monkeypatch.setattr(off, "get_active", get_active)
        monkeypatch.setattr(off, "list_open", lambda kind: [])
    log["install"] = install
    log["app"] = app

    def copy(*a):
        log["copied"] += 1
    log["chart"] = types.SimpleNamespace(CopyPicture=copy)
    return log


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
    assert office.method_order("ppt") == ["relay", "export", "redraw"]
    assert office.method_order("word") == ["export", "redraw", "relay"]


def test_transfer_method_option_removed(capsys):
    """사용자 결정(1.4.52): --method 없음. 방식은 자동 순서(--probe가 저장한 순서)만."""
    code, out = run_office(capsys, "transfer", "--from-app", "excel", "--chart", "c", "--to-app", "word", "--method", "relay")
    assert code == 2 and "--method" in out["message"], out


def test_chart_transfer_falls_back_from_encrypted_export_to_redraw(chart_env, capsys):
    code, out = run_office(capsys, "transfer", "--from-app", "excel", "--chart", "사인 2주기", "--to-app", "word,ppt")
    assert code == 0, out
    assert out["results"]["word"]["method"] == "redraw" and out["results"]["ppt"]["method"] == "redraw"
    assert chart_env["calls"].count("redraw") == 1                     # 한 번 그린 그림을 두 앱에 재사용
    assert [k for k, _ in chart_env["insert"]] == ["word", "ppt"]


def test_chart_transfer_all_methods_fail_suggests_probe(chart_env, capsys, monkeypatch):
    def boom(*a, **k):
        raise RuntimeError("실패")
    monkeypatch.setattr(office, "image_by_redraw", boom)
    monkeypatch.setattr(office, "relay_chart", boom)
    code, out = run_office(capsys, "transfer", "--from-app", "excel", "--chart", "사인 2주기", "--to-app", "ppt")
    assert code == 1 and "--probe" in out["message"]


def test_transfer_argument_rules(capsys):
    code, out = run_office(capsys, "transfer", "--from-app", "excel", "--to-app", "word")
    assert code == 1 and "--chart" in out["message"]
    code, out = run_office(capsys, "transfer", "--from-app", "word", "--to-app", "ppt")
    assert code == 1 and "read --save" in out["message"]
    code, out = run_office(capsys, "transfer", "--from-app", "엑셀", "--to-app", "엑셀", "--range", "A1")
    assert code == 1
    code, out = run_office(capsys, "transfer", "--from-app", "hwp", "--to-app", "word")
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
    code, out = run_office(capsys, "transfer", "--from-app", "word", "--to-app", "excel")
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
    code, out = run_office(capsys, "transfer", "--from-app", "word", "--to-app", "excel")
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
    code, out = run_office(capsys, "transfer", "--from-app", "word", "--to-app", "excel", "--sheet", "결과")
    assert code == 1 and "--overwrite" in out["message"], out
    argv = _recommended_argv(out["message"])                    # 1.4.79: [추천]은 전체 명령 (--sheet만 새 이름)
    assert argv == ["transfer", "--from-app", "word", "--to-app", "excel", "--sheet", "결과 2"], argv
    assert wb.Worksheets("결과").values[(1, 1)] == "기존"
    code, out = run_office(capsys, *argv)                       # [추천]대로 다시 실행하면 새 시트로 받음
    assert code == 0 and out["sheet"] == "결과 2" and wb.Worksheets("결과").values[(1, 1)] == "기존", out
    code, out = run_office(capsys, "transfer", "--from-app", "word", "--to-app", "excel", "--sheet", "결과", "--overwrite")
    assert code == 0 and out["sheet"] == "결과" and wb.Worksheets("결과").values[(1, 1)] == "a"
    code, out = run_office(capsys, "transfer", "--from-app", "word", "--to-app", "excel", "--sheet", "새 결과")
    assert code == 0 and out["sheet"] == "새 결과"                       # --sheet는 새 시트 이름: 없으면 만듦


def test_chart_transfer_stops_on_ambiguous_anchor_without_trying_other_methods(chart_env, capsys, monkeypatch):
    calls = []

    def ambiguous_anchor(tgt, img, anchor, occurrence=None):
        calls.append(img)
        raise rt.ambiguous("'계획'이(가) 단락 2개에 있어 어느 곳인지 고를 수 없습니다.", ["1: 계획 A", "2: 계획 B"])
    monkeypatch.setattr(office, "insert_picture_word", ambiguous_anchor)
    code, out = run_office(capsys, "transfer", "--from-app", "excel", "--chart", "사인 2주기", "--to-app", "word", "--anchor", "계획")
    assert code == 1 and out["error_type"] == "ambiguous" and out["candidates"] == ["1: 계획 A", "2: 계획 B"]
    assert len(calls) == 1                                                # 다른 방식으로 넘어가지 않음


# ---------------------------------------------------------------------------
# transfer 일부 실패는 UNVERIFIED, undo --pick은 같은 문서의 더 새 지점을 정리, 파일 이름만 준 대상은 현재 폴더에 만들지 않음
# ---------------------------------------------------------------------------


def test_transfer_partial_failure_is_unverified(monkeypatch, capsys, tmp_path):
    app = FakeXlApp()
    wb = FakeExcelWorkbook(app)
    stub_com(monkeypatch)
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
    code, out = run_main(office, capsys, "transfer", "--from-app", "excel", "--chart", "차트", "--to-app", "word,ppt")
    assert code == 0 and "word" in out["results"] and "ppt" in out["failures"]
    assert out["verified"] is False and out["agent_directive"].startswith("UNVERIFIED:")


def test_undo_pick_removes_newer_points_of_same_document(tmp_path, monkeypatch, capsys):
    stub_com(monkeypatch)
    monkeypatch.setattr(off, "match_open", lambda kind, target, docs=None, option="--target": None)
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
    code, out = run_main(office, capsys, "undo", "--pick", "3")
    assert code == 0 and orig.read_bytes() == b"snapshot-of-a.xlsx" and out["removed_newer_points"] == 1, out
    left = off.list_snapshots()
    assert [os.path.basename(s["original_path"]) for s in left] == ["b.xlsx"]       # 다음 undo가 앞으로 가지 않음


def test_bare_name_is_not_created_in_cwd(monkeypatch, tmp_path):
    stub_com(monkeypatch)
    monkeypatch.chdir(tmp_path)
    monkeypatch.setattr(off, "list_open", lambda kind: [{"name": "보고.xlsx", "path": "C:\\docs\\보고.xlsx"}])
    monkeypatch.setattr(off, "match_open", lambda kind, target, docs=None, option="--target": None)
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
# 잘못된 시트 이름은 시트를 만들기 전에 막음 ([추천] 고친 이름, 빈 시트가 남지 않음)
# ---------------------------------------------------------------------------


def test_transfer_bad_sheet_name_checked_before_opening(monkeypatch, tmp_path):
    # 새로 만들 파일은 열기(만들기) 전에 확인. 열린 문서는 시트 이름을 안 뒤에 확인 (1.4.15)
    def boom(*a, **k):
        raise AssertionError("문서를 열면 안 됨")
    monkeypatch.setattr(off, "open_target", boom)
    args = types.SimpleNamespace(sheet="표/목록", to_target=str(tmp_path / "새.xlsx"), overwrite=False)
    with pytest.raises(rt.InputError) as ei:
        office._to_excel(args, [["a"]], "Word 표")
    assert "[추천] 표_목록" in str(ei.value)


# ---------------------------------------------------------------------------
# transfer --sheet의 [추천]은 받는 문서의 시트 이름을 안 뒤에 (따라 한 번 다시 실행하면 성공)
# ---------------------------------------------------------------------------


def test_transfer_recommendation_avoids_existing_sheet(xl_ab):
    _, wb = xl_ab
    args = types.SimpleNamespace(sheet="a/b", to_target="active", overwrite=False)
    with pytest.raises(rt.InputError) as ei:
        office._to_excel(args, [["h"], ["v"]], "Word 표")
    rec = str(ei.value).split("[추천] ")[1].strip()
    assert rec == "a_b 2", rec                               # 있는 시트 a_b와 겹치지 않음
    assert sheet_names(wb) == ["Sheet1", "a_b"] and undo_points() == []
    args.sheet = rec
    out = office._to_excel(args, [["h"], ["v"]], "Word 표")   # [추천]대로 1회 재시도하면 성공
    assert out["sheet"] == "a_b 2" and sheet_names(wb) == ["Sheet1", "a_b", "a_b 2"]


def test_transfer_reserved_name_into_open_document(xl_ab):
    _, wb = xl_ab
    args = types.SimpleNamespace(sheet="History", to_target="active", overwrite=False)
    with pytest.raises(rt.InputError) as ei:
        office._to_excel(args, [["h"], ["v"]], "Word 표")
    assert "[추천] History_" in str(ei.value)
    assert sheet_names(wb) == ["Sheet1", "a_b"] and undo_points() == []


def test_new_file_target_is_checked_before_creating(monkeypatch, tmp_path):
    def boom(*a, **k):
        raise AssertionError("새 파일을 만들면 안 됨")
    monkeypatch.setattr(off, "open_target", boom)
    target = tmp_path / "새.xlsx"
    args = types.SimpleNamespace(sheet="History", to_target=str(target), overwrite=False)
    with pytest.raises(rt.InputError) as ei:
        office._to_excel(args, [["h"]], "Word 표")
    assert "[추천] History_" in str(ei.value) and not target.exists()
    assert office._new_file_target(str(target)) and not office._new_file_target("active")


# ---------------------------------------------------------------------------
# 새 파일의 기본 시트 이름으로 transfer
# ---------------------------------------------------------------------------


def test_transfer_into_new_file_with_default_sheet_name(monkeypatch, tmp_path):
    from fake_office import FakeExcelApp
    stub_com(monkeypatch)
    monkeypatch.setattr(off, "_new_instance", lambda kind: FakeExcelApp())
    target = tmp_path / "new.xlsx"
    args = types.SimpleNamespace(sheet="Sheet1", to_target=str(target), overwrite=False, no_focus=True)
    # 기본 시트 이름과 같아도 '이미 있음'으로 막지 않고 그 빈 시트에 받음 (가짜 시트는 쓰기 단계의 Range가 없어 거기서 멈춤)
    try:
        office._to_excel(args, [["h"], ["v"]], "Word 표")
    except rt.InputError as e:
        pytest.fail(f"새 파일의 기본 시트 이름이 막힘: {e}")
    except AttributeError as e:
        assert "Range" in str(e)


# ---------------------------------------------------------------------------
# 새 파일의 기본 시트가 여럿(Office 2007/2010: Sheet1~3)이어도 기본 시트 이름으로 받음
# ---------------------------------------------------------------------------


def test_transfer_new_file_with_three_default_sheets(monkeypatch, tmp_path):
    stub_com(monkeypatch)

    def add3(self, Template=None):
        wb = fake_office.FakeWorkbook("", sheets=("Sheet1", "Sheet2", "Sheet3"))
        self.items.append(wb)
        return wb

    monkeypatch.setattr(fake_office.FakeWorkbooks, "Add", add3)
    monkeypatch.setattr(off, "_new_instance", lambda kind: fake_office.FakeExcelApp())
    for sheet in ("Sheet1", "sheet3"):
        target = tmp_path / f"new_{sheet}.xlsx"
        args = types.SimpleNamespace(sheet=sheet, to_target=str(target), overwrite=False, no_focus=True)
        # 가짜 시트는 쓰기 단계의 Range가 없어 거기서 멈춤 (그 전 '이미 있음'으로 막히지 않으면 됨)
        try:
            office._to_excel(args, [["h"], ["v"]], "Word 표")
        except rt.InputError as e:
            pytest.fail(f"새 파일의 기본 시트 이름 {sheet}이(가) 막힘: {e}")
        except AttributeError as e:
            assert "Range" in str(e)


# ---------------------------------------------------------------------------
# transfer --to-app word,ppt: 한 앱만 열려 있거나 한 앱만 실패할 때 (먼저 넣은 쪽을 다시 넣지 않게)
# ---------------------------------------------------------------------------


def test_chart_word_only_open_changes_nothing_and_recommends_word(env, capsys):
    code, out = run_office(capsys, *CHART)
    assert code == 1 and env["inserted"] == []
    argv = _check_target_not_open(out)
    code, out = run_office(capsys, *argv)
    assert code == 0 and env["inserted"] == ["word"] and list(out["results"]) == ["word"], out


def test_range_word_only_open_changes_nothing_and_recommends_word(env, capsys):
    before = env["word"].text
    code, out = run_office(capsys, *RANGE)
    assert code == 1 and env["word"].text == before
    argv = _check_target_not_open(out)
    code, out = run_office(capsys, *argv)
    assert code == 0 and list(out["results"]) == ["word"] and env["word"].text != before, out


def test_app_typo_in_to_list_recommends_full_command(env, capsys):
    """1.4.79 회귀: --to-app 목록(쉼표) 안의 앱 이름 하나만 틀려도 [추천]은 그 항목만 고친 전체 명령 (값 'word'만 주지 않음)."""
    code, out = run_office(capsys, "transfer", "--from-app", "excel", "--range", "A1:B2", "--to-app", "wrod,ppt")
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:"), out
    argv = _recommended_argv(out["message"])
    assert argv == list(RANGE) and "wrod,ppt" not in argv, argv
    env["prs"] = FakePresentation(1)
    code, out = run_office(capsys, *argv)                       # [추천]대로 다시 실행하면 두 앱에 넣음
    assert code == 0 and sorted(out["results"]) == ["ppt", "word"], out


@pytest.mark.parametrize("old", [
    ("transfer", "--from", "excel", "--range", "A1:B2", "--to", "word,ppt"),
    ("transfer", "--from-app", "excel", "--range", "A1:B2", "--to", "word,ppt"),          # 하나만 옛 이름
    ("transfer", "--from", "excel", "--range", "A1:B2", "--to-app", "word,ppt"),
])
def test_transfer_old_app_option_names_recommend_new_names(env, capsys, old):
    """1.4.79: transfer의 앱은 --from-app/--to-app (--to/--from은 outlook의 받는 사람·보낸 사람 뜻만, 사용법 단순성 4번).
    옛 이름은 받지 않고(별칭 없음) 인자 오류 + 고친 전체 명령 [추천]. 그대로 다시 실행하면 넣어짐."""
    code, out = run_office(capsys, *old)
    assert code == 2 and out["agent_directive"].startswith("STOP_RETRY:") and env["inserted"] == [], out
    argv = _recommended_argv(out["message"])
    assert argv == list(RANGE) and "--from" not in argv and "--to" not in argv, argv
    env["prs"] = FakePresentation(1)
    code, out = run_office(capsys, *argv)
    assert code == 0 and sorted(out["results"]) == ["ppt", "word"], out


def test_transfer_new_app_option_names_in_help_and_errors(capsys):
    """1.4.79: 도움말·안내 글도 새 이름 (--from-app/--to-app). 옛 이름 --to/--from은 transfer 파서에 없음."""
    p = office.build_cli()
    sub = next(a for a in p._actions if a.__class__.__name__ == "_SubParsersAction").choices["transfer"]
    names = {o for a in sub._actions for o in a.option_strings}
    assert {"--from-app", "--to-app", "--from-target", "--to-target"} <= names and not {"--from", "--to"} & names
    code, out = run_office(capsys, "transfer", "--chart", "c")
    assert code == 1 and "--from-app" in out["message"] and "--to-app" in out["message"], out


def test_no_receiver_open_names_to_target_for_all_apps(env, capsys):
    env["word"] = None
    code, out = run_office(capsys, *CHART)
    assert code == 1 and env["inserted"] == [] and out["error_type"] == "needs_user_decision"
    assert "--to-target" in out["agent_directive"] and "모든 앱" in out["agent_directive"]
    assert "--target " not in out["agent_directive"]


def test_single_missing_target_and_source_name_transfer_options(env, capsys):
    code, out = run_office(capsys, "transfer", "--from-app", "excel", "--range", "A1:B2", "--to-app", "ppt")
    assert code == 1 and '--to-target "' in out["agent_directive"] and "--target " not in out["agent_directive"]
    env["word"] = None
    code, out = run_office(capsys, "transfer", "--from-app", "word", "--to-app", "excel")
    assert code == 1 and '--from-target "' in out["agent_directive"] and "--target " not in out["agent_directive"]


def test_invalid_slide_is_checked_before_word_changes(env, capsys):
    env["prs"] = FakePresentation(1)
    before = env["word"].text
    code, out = run_office(capsys, *RANGE, "--slide", "5")
    assert code == 1 and "슬라이드 5번이 없습니다" in out["message"] and env["word"].text == before
    code, out = run_office(capsys, *CHART, "--slide", "5")
    assert code == 1 and env["inserted"] == []


def test_invalid_anchor_is_checked_before_anything(env, capsys):
    env["prs"] = FakePresentation(1)
    code, out = run_office(capsys, *RANGE, "--anchor", "없는 문구")
    assert code == 1 and "찾을 수 없습니다" in out["message"] and env["prs"].Slides.Count == 1


def test_range_late_failure_is_unverified_with_lists(env, capsys, monkeypatch):
    env["prs"] = FakePresentation(1)

    def boom(self, rows):
        raise RuntimeError("PPT 표 만들기 실패")
    monkeypatch.setattr(ppt.SlideWriter, "table", boom)
    before = env["word"].text
    code, out = run_office(capsys, *RANGE)
    assert code == 0 and out["status"] == "ok" and out["verified"] is False, out
    assert out["succeeded"] == ["word"] and out["failed"] == ["ppt"] and "PPT 표 만들기 실패" in out["failures"]["ppt"]
    assert out["agent_directive"].startswith("UNVERIFIED:") and "한 번 더" in out["agent_directive"]
    assert env["word"].text != before


def test_chart_no_fallback_failure_after_word_is_unverified(env, capsys, monkeypatch):
    env["prs"] = FakePresentation(1)

    def ppt_fail(tgt, img, slide, title):
        raise rt.InputError("PowerPoint 대상 문제", no_fallback=True)
    monkeypatch.setattr(office, "insert_picture_ppt", ppt_fail)
    code, out = run_office(capsys, *CHART)
    assert code == 0 and out["verified"] is False and env["inserted"] == ["word"], out
    assert out["succeeded"] == ["word"] and out["failed"] == ["ppt"]
    assert out["agent_directive"].startswith("UNVERIFIED:") and "--to-app ppt" in out["agent_directive"]


def test_undo_points_after_partial_and_preflight_failure(env, capsys, monkeypatch):
    # 실제 스냅샷 저장: 미리 확인에서 멈추면 지점 없음, 워드만 바뀌면 워드 지점만 남음
    code, out = run_office(capsys, *RANGE)
    assert code == 1 and off.list_snapshots() == []
    env["prs"] = FakePresentation(1)

    def boom(self, rows):
        raise RuntimeError("PPT 표 만들기 실패")
    monkeypatch.setattr(ppt.SlideWriter, "table", boom)
    code, out = run_office(capsys, *RANGE)
    assert code == 0 and out["failed"] == ["ppt"]
    assert [s["app"] for s in off.list_snapshots()] == ["word"]


def test_chart_next_commands_use_open_receivers(monkeypatch):
    monkeypatch.setattr(off, "list_open", lambda kind: [{"name": "a"}] if kind == "word" else [])
    assert excel._open_receivers() == ["word"]
    monkeypatch.setattr(off, "list_open", lambda kind: [])
    assert excel._open_receivers() == ["word", "ppt"]

    def boom(kind):
        raise RuntimeError("COM 없음")
    monkeypatch.setattr(off, "list_open", boom)
    assert excel._open_receivers() == ["word", "ppt"]


# ---------------------------------------------------------------------------
# 이름이 같은 열린 문서가 여럿이면 transfer에 있는 옵션으로 [후보] 안내
# ---------------------------------------------------------------------------


def test_ambiguous_open_documents_name_transfer_options(env, capsys, monkeypatch):
    env["prs"] = FakePresentation(1)
    monkeypatch.setattr(off, "list_open", lambda kind: TWO_SAME if kind == "word" else [])
    # 앱이 둘이면 넣기 전에 받는 문서를 모두 찾음(_probe_target). 보내는 쪽은 open_target(option=--from-target)
    for argv, opt in (((*CH, "--to-app", "word,ppt", "--to-target", "a.docx"), "--to-target"),
                      (("transfer", "--from-app", "word", "--from-target", "a.docx", "--to-app", "excel"), "--from-target")):
        code, out = run_main(office, capsys, *argv)
        assert code == 1 and "[후보]" in out["message"] and env["inserted"] == [], out
        assert f"{opt}에 전체 경로" in out["message"], out


def test_ambiguous_focus_names_name_option(monkeypatch, capsys):
    monkeypatch.setattr(off, "list_open", lambda kind: TWO_SAME)
    code, out = run_main(office, capsys, "focus", "--app", "word", "--name", "a.docx")
    assert code == 1 and "--name에 전체 경로" in out["message"], out


# ---------------------------------------------------------------------------
# 대상 문제는 다음 방식으로 넘어가지 않음
# ---------------------------------------------------------------------------


def test_chart_target_recommend_stops_unchanged(chart_env, capsys, monkeypatch, isolated_dirs, relay_spy):
    monkeypatch.setattr(off, "list_open", lambda kind: [{"name": "ZQX_보고서.docx", "path": "", "app": None, "doc": None}])
    monkeypatch.setattr(office, "insert_picture_word", _real_insert_word)
    monkeypatch.setattr(rt, "CURRENT_ARGV", None)
    code, out = run_office(capsys, "transfer", "--from-app", "excel", "--chart", "사인 2주기", "--to-app", "word",
                           "--to-target", "zqx 보고서.docx")
    assert code == 1, out
    argv = _recommended_argv(out["message"])                    # 1.4.79: [추천]은 전체 명령 (--to-target만 열린 문서 이름)
    assert option_value(argv, "--to-target") == "ZQX_보고서.docx" and "zqx 보고서.docx" not in argv, argv
    assert option_value(argv, "--chart") == "사인 2주기" and option_value(argv, "--to-app") == "word", argv
    assert "relay" not in out["message"] and "--probe" not in out["message"] and "시도한 방식" not in out["message"]
    assert relay_spy == [] and chart_env["calls"] == ["export", "redraw"]       # 그림 문제(export)만 다음 방식으로
    log = error_log_text(isolated_dirs)
    assert log and SECRET not in log and "[추천 생략]" in log, log


def test_chart_target_bad_extension_stops(chart_env, capsys, monkeypatch, tmp_path, relay_spy):
    monkeypatch.setattr(office, "insert_picture_word", _real_insert_word)
    code, out = run_office(capsys, "transfer", "--from-app", "excel", "--chart", "사인 2주기", "--to-app", "word",
                           "--to-target", str(tmp_path / "새 문서.txt"))
    assert code == 1 and "확장자" in out["message"] and "relay" not in out["message"], out
    assert relay_spy == []


def test_relay_resolves_target_before_clipboard(chart_env, capsys, monkeypatch):
    copied = []
    chart = type("C", (), {"Name": "사인 2주기", "CopyPicture": lambda self, *a: copied.append(a)})()
    import excel
    monkeypatch.setattr(excel, "find_chart", lambda t, name, sheet=None: (None, chart))
    monkeypatch.setattr(office, "relay_chart", _real_relay)
    monkeypatch.setattr(office, "method_order", lambda kind: ["relay"])
    monkeypatch.setattr(off, "list_open", lambda kind: [{"name": "ZQX_보고서.docx", "path": "", "app": None, "doc": None}])
    code, out = run_office(capsys, "transfer", "--from-app", "excel", "--chart", "사인 2주기", "--to-app", "word",
                           "--to-target", "zqx 보고서.docx")
    assert code == 1, out
    argv = _recommended_argv(out["message"])                    # 1.4.79: [추천]은 전체 명령 (--to-target만 열린 문서 이름)
    assert option_value(argv, "--to-target") == "ZQX_보고서.docx" and "zqx 보고서.docx" not in argv, argv
    assert option_value(argv, "--chart") == "사인 2주기" and option_value(argv, "--to-app") == "word", argv
    assert copied == []                                   # 사용자의 클립보드를 덮어쓰지 않음


def test_resolve_target_errors_are_no_fallback(monkeypatch, tmp_path):
    monkeypatch.setattr(off, "list_open", lambda kind: [{"name": "ZQX가.docx", "path": "", "app": None, "doc": None},
                                                       {"name": "ZQX나.docx", "path": "", "app": None, "doc": None}])
    for target in ("zqx 가.docx", "가나다.docx", "없는.docx", str(tmp_path / "a.txt")):
        with pytest.raises(rt.InputError) as ei:
            off.resolve_target("word", target, readonly=False, create=True, option="--to-target")
        assert ei.value.no_fallback, (target, str(ei.value))


def test_slide_errors_are_no_fallback():
    prs = type("P", (), {"Slides": type("S", (), {"Count": 2})()})()
    for value in ("9", "x"):
        with pytest.raises(rt.InputError) as ei:
            ppt._resolve_slide(prs, value, default_new=False)
        assert ei.value.no_fallback


# ---------------------------------------------------------------------------
# 차트 다시 그리기의 $, 오류 로그의 내부 오류 글
# ---------------------------------------------------------------------------


def test_render_chart_does_not_parse_mathtext(tmp_path, monkeypatch):
    import matplotlib.text as mtext
    seen = []
    real_draw = mtext.Text.draw

    def draw(self, renderer):
        if "$" in self.get_text():
            seen.append(self.get_parse_math())
        return real_draw(self, renderer)
    monkeypatch.setattr(mtext.Text, "draw", draw)
    bad = "단가 $\\frac 와 $"
    for kind in (51, 4, 5):
        data = {"type": kind, "title": bad, "width": 300, "height": 180,
                "series": [{"name": bad, "values": [1, 2], "x": [bad, "$5 ~ $10"]},
                           {"name": "$x^2$", "values": [2, 3], "x": [bad, "b"]}]}
        path = office.render_chart(data, str(tmp_path / f"{kind}.png"))
        assert os.path.getsize(path) > 0
    assert seen and not any(seen), seen                   # 제목, 계열 이름(범례), 눈금 글 모두 글자 그대로
    import matplotlib
    assert matplotlib.rcParams["text.parse_math"] is True   # 전역 설정은 그대로


def test_redraw_with_dollar_title_succeeds(chart_env, capsys, monkeypatch, relay_spy):
    monkeypatch.setattr(office, "image_by_redraw", lambda obj, stem: office.render_chart(
        {"type": 51, "title": "단가 $\\frac 와 $", "width": 300, "height": 180,
         "series": [{"name": "s", "values": [1, 2], "x": ["a", "b"]}]},
        rt.resolve_output_path(None, stem + "_redraw", ".png")))
    code, out = run_office(capsys, "transfer", "--from-app", "excel", "--chart", "사인 2주기", "--to-app", "ppt")
    assert code == 0 and out["results"]["ppt"]["method"] == "redraw" and relay_spy == [], out


def test_guard_chart_internal_error_text_not_logged(chart_env, capsys, monkeypatch, isolated_dirs):
    def boom(kind):
        def f(*a, **k):
            raise kind("ZQX 매출 차트 제목")
        return f
    monkeypatch.setattr(office, "image_by_export", boom(RuntimeError))
    monkeypatch.setattr(office, "image_by_redraw", boom(ValueError))
    monkeypatch.setattr(office, "relay_chart", boom(KeyError))
    monkeypatch.setattr(rt, "CURRENT_ARGV", None)
    code, out = run_office(capsys, "transfer", "--from-app", "excel", "--chart", "사인 2주기", "--to-app", "word")
    assert code == 1 and SECRET in out["message"], out
    log = error_log_text(isolated_dirs)
    assert log and SECRET not in log, log
    assert "export: RuntimeError" in log and "redraw: ValueError" in log and "relay: KeyError" in log, log


# ---------------------------------------------------------------------------
# relay는 --slide를 클립보드·임시 워드보다 먼저 확인
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("target", [None, "file"])
def test_relay_ppt_bad_slide_before_clipboard(relay_env, capsys, tmp_path, target):
    argv = ["transfer", "--from-app", "excel", "--chart", "c", "--to-app", "ppt", "--slide", "9"]
    if target == "file":
        p = tmp_path / "deck.pptx"
        p.write_bytes(b"x")
        argv += ["--to-target", str(p)]                   # 열려 있지 않은 파일: 열어서 슬라이드를 먼저 확인
    code, out = run_office(capsys, *argv)
    assert code == 1 and "슬라이드 9번이 없습니다" in out["message"], out
    assert relay_env["copied"] == [] and relay_env["dispatch"] == []      # 클립보드·임시 워드 그대로


def test_relay_ppt_ok_after_slide_check(relay_env, capsys):
    code, out = run_office(capsys, "transfer", "--from-app", "excel", "--chart", "c", "--to-app", "ppt", "--slide", "2")
    assert code == 0 and out["results"]["ppt"]["method"] == "relay", out
    assert len(relay_env["copied"]) == 1 and relay_env["dispatch"] == ["Word.Application"]
    assert relay_env["prs"].Slides(2).Shapes.Count == 1 and relay_env["opened"] == [("ppt", "active")]


# ---------------------------------------------------------------------------
# --to-target ""는 활성 문서
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("kind", ["word", "ppt"])
def test_relay_empty_target_is_active(relay_env, capsys, kind):
    slide = ["--slide", "1"] if kind == "ppt" else []
    code, out = run_office(capsys, "transfer", "--from-app", "excel", "--chart", "c", "--to-app", kind, "--to-target", "",
                           *slide)
    assert code == 0 and out["results"][kind]["method"] == "relay", out["message"]
    assert "KeyError" not in json.dumps(out, ensure_ascii=False)
    assert relay_env["opened"] == [(kind, "")] and len(relay_env["copied"]) == 1


def test_relay_unit_empty_target_returns_result(relay_env):
    done = office.relay_chart(type("C", (), {"CopyPicture": lambda s, *a: None})(), "", None, None, None, "t")
    assert list(done) == ["word"]


# ---------------------------------------------------------------------------
# relay는 워드 단계 뒤에 PPT 쓰기 세션을 엶 (실제 open_target, 가짜 COM)
# ---------------------------------------------------------------------------


def test_failed_temp_word_keeps_no_point_for_unsaved_active_deck(relay_real, tmp_path):
    deck = _Prs(str(tmp_path / "user.pptx"), saved=0)
    (tmp_path / "user.pptx").write_bytes(b"x")
    relay_real["install"](active=deck, paste_ok=False)
    with pytest.raises(RuntimeError):
        office.relay_chart(relay_real["chart"], None, "active", None, "1", "t")
    assert off.list_snapshots("ppt") == []                # 바꾼 것 없는 지점이 남지 않음
    assert deck.Slides(1).Shapes.Count == 0


def test_undo_still_picks_earlier_real_point(relay_real, tmp_path):
    deck = _Prs(str(tmp_path / "user.pptx"), saved=0)
    (tmp_path / "user.pptx").write_bytes(b"x")
    relay_real["install"](active=deck, paste_ok=False)
    with off.open_target("ppt", "active", readonly=False, tool="ppt", command="write") as t:
        t.changed = True                                  # 앞서 한 진짜 쓰기
    before = [s["meta_file"] for s in off.list_snapshots(None)]
    assert len(before) == 1
    time.sleep(1.05)                                      # 지점 이름은 초 단위
    with pytest.raises(RuntimeError):
        office.relay_chart(relay_real["chart"], None, "active", None, "1", "t")
    assert [s["meta_file"] for s in off.list_snapshots(None)] == before


@pytest.mark.parametrize("where", ["active", "file", "new"])
def test_bad_slide_stops_before_clipboard_and_point(relay_real, tmp_path, where):
    deck = _Prs(str(tmp_path / "user.pptx"), saved=0)
    (tmp_path / "user.pptx").write_bytes(b"x")
    relay_real["install"](active=deck)
    target = {"active": "active", "file": str(tmp_path / "user.pptx"), "new": str(tmp_path / "새.pptx")}[where]
    with pytest.raises(rt.InputError) as ei:
        office.relay_chart(relay_real["chart"], None, target, None, "9", "t")
    assert "슬라이드 9번이 없습니다" in str(ei.value) and ei.value.no_fallback
    assert relay_real["copied"] == 0 and relay_real["dispatch"] == 0    # 클립보드·임시 워드 그대로
    assert off.list_snapshots(None) == []
    assert not (tmp_path / "새.pptx").exists()               # 새 파일도 만들지 않음


def test_success_active_and_file(relay_real, tmp_path):
    deck = _Prs(str(tmp_path / "user.pptx"))
    (tmp_path / "user.pptx").write_bytes(b"x")
    relay_real["install"](active=deck)
    done = office.relay_chart(relay_real["chart"], None, "active", None, "2", "t")
    assert done["ppt"]["source"] == "active" and deck.Slides(2).Shapes.Count == 1 and deck.closed == 0
    p = tmp_path / "file.pptx"
    p.write_bytes(b"x")
    done = office.relay_chart(relay_real["chart"], None, str(p), None, "1", "t")
    opened = relay_real["app"].Presentations.items[-1]
    assert done["ppt"]["source"] == "file" and opened.closed == 1 and opened.Slides(1).Shapes.Count == 1


# ---------------------------------------------------------------------------
# 계산된 값이 없는 수식 칸
# ---------------------------------------------------------------------------


def test_transfer_range_rows_report_uncomputed(tmp_path):
    p = formula_xlsx(tmp_path / "f.xlsx")
    args = types.SimpleNamespace(from_target=p, sheet=None, range="A1:B4")
    rows, sheet, info = office._excel_range_rows(args)
    assert sheet == "합계" and info["cells"] == ["합계!B4"]
    payload = rt.with_uncomputed(rt.ok("office", "transfer", "x", "VERIFIED: 확인", verified=True, rows=4), info)
    assert payload["agent_directive"].startswith("UNVERIFIED:") and payload["rows"] == 4 and payload["uncomputed_count"] == 1
    # 범위 전체가 계산된 값 없는 수식이면 '데이터 없음' 오류에 이유를 붙임
    args = types.SimpleNamespace(from_target=p, sheet=None, range="B4:B4")
    with pytest.raises(rt.InputError) as ei:
        office._excel_range_rows(args)
    assert "계산된 값" in str(ei.value) and "합계!B4" in str(ei.value)
    assert "합계!B4" not in ei.value.log_message and "수식 칸 1개" in ei.value.log_message    # 로그에는 칸 주소 없음


# ---------------------------------------------------------------------------
# undo: 닫힌 파일 복원, --pick 범위
# ---------------------------------------------------------------------------


def test_undo_closed_file_restores_backs_up_and_consumes(tmp_path, monkeypatch, capsys):
    fake, _ = make_win32com()
    monkeypatch.setattr(off, "win32com", fake)
    monkeypatch.setattr(off, "pythoncom", None)
    monkeypatch.setattr(off, "match_open", lambda kind, target, docs=None, option="--target": None)
    orig = tmp_path / "a.xlsx"
    orig.write_bytes(b"current")
    off.save_snapshot("excel", FakeWorkbook(str(orig)), "excel", "write")      # 스냅샷 내용: snapshot-of-a.xlsx
    code, out = run_office(capsys, "undo", "--list")
    assert out["snapshots"][0]["by"] == "excel write"
    assert out["snapshots"][0]["command"].endswith("undo --pick 1")
    code, out = run_office(capsys, "undo", "--list", "--app", "excel")
    assert out["snapshots"][0]["command"].endswith("undo --pick 1 --app excel")     # 같은 필터로 고르기
    code, out = run_office(capsys, "undo")
    assert code == 0 and out["restored_as"] == "file_replaced" and out["agent_directive"].startswith("VERIFIED:")   # 스냅샷과 비교 (R1)
    assert orig.read_bytes() == b"snapshot-of-a.xlsx"
    assert open(out["pre_undo_backup"], "rb").read() == b"current"
    code, out = run_office(capsys, "undo")
    assert code == 0 and out["restored"] is False                         # 소비되어 더 없음


def test_undo_pick_out_of_range(tmp_path, monkeypatch, capsys):
    orig = tmp_path / "a.xlsx"
    orig.write_bytes(b"x")
    off.save_snapshot("excel", FakeWorkbook(str(orig)))
    code, out = run_office(capsys, "undo", "--pick", "5")
    assert code == 1 and "--list" in out["message"]


# ---------------------------------------------------------------------------
# 저장한 적 없는 문서 undo (시나리오 12: 새 워드 문서에 차트를 넣은 뒤 "방금 한 거 되돌려줘")
# ---------------------------------------------------------------------------


def test_undo_unsaved_word_restores_in_same_window(untitled_env, capsys):
    doc = _UntitledWordDoc("문서1", "<before/>")
    app = type("WordApp", (), {})()
    app.Documents = _Docs([doc])
    untitled_env.append(app)
    off.save_snapshot("word", doc, "office", "transfer")                 # flat_xml 스냅샷
    doc.WordOpenXML = "<with-chart/>"                                    # 차트가 들어간 상태
    code, out = run_office(capsys, "undo")
    assert code == 0 and out["restored_as"] == "restored_in_place" and out["name"] == "문서1"
    assert doc.inserted == ["<before/>"] and doc.activated == 1          # 같은 문서의 내용을 바꿈
    assert app.Documents.Count == 1 and app.Documents.added == []          # 새 창을 열지 않음
    assert open(out["pre_undo_backup"], encoding="utf-8").read() == "<with-chart/>"
    assert "같은 창" in out["preview"] and "같은 창" in out["agent_directive"]
    assert out["agent_directive"].startswith("VERIFIED:")


def test_undo_unsaved_excel_replaces_and_closes_modified(untitled_env, capsys):
    app = FakeExcelApp()
    wb = FakeWorkbook("")                                                # Book1, 저장 안 됨
    app.Workbooks.items.append(wb)
    _close_removes(wb, app.Workbooks)
    untitled_env.append(app)
    off.save_snapshot("excel", wb, "excel", "write")
    code, out = run_office(capsys, "undo")
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
    code, out = run_office(capsys, "undo")
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
    code, out = run_office(capsys, "undo")
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
    code, out = run_office(capsys, "undo")
    assert code == 0 and out["restored_as"] == "new_untitled_document"
    assert "2개" in out["preview"] and "그대로 열려 있습니다" in out["preview"]
    assert not any(c[0] == "Close" for c in a.calls + b.calls)           # 어느 것도 닫지 않음


def test_windows_lists_documents(monkeypatch, capsys):
    monkeypatch.setattr(off, "open_documents_summary",
                        lambda: [{"app": "Excel", "name": "a.xlsx", "path": "", "saved": True, "active": True}])
    monkeypatch.setattr(off, "running_office_caps", lambda: {})
    code, out = run_office(capsys, "windows")
    assert code == 0 and out["preview"] == "Excel 1개" and out["office"] == {}


def test_windows_reports_office_version_and_missing_pdf(monkeypatch, capsys):
    monkeypatch.setattr(off, "open_documents_summary",
                        lambda: [{"app": "Word", "name": "a.docx", "path": "", "saved": True, "active": True}])
    monkeypatch.setattr(off, "running_office_caps",
                        lambda: {"Word": {"version": "12.0", "major": 12, "pdf_export": False}})
    code, out = run_office(capsys, "windows")
    assert code == 0 and out["office"]["Word"]["pdf_export"] is False
    assert "Office 12.0" in out["preview"] and "PDF 내보내기 불가: Word" in out["preview"]


# ---------------------------------------------------------------------------
# 되돌릴 지점이 없는 직전 작업, 지점 뒤에 다시 저장된 파일
# ---------------------------------------------------------------------------


def test_undo_stops_at_write_without_snapshot(tmp_path, monkeypatch, capsys):
    orig = _undo_env(tmp_path, monkeypatch)
    off.save_snapshot("excel", FakeWorkbook(str(orig)), "excel", "write")          # 월요일의 지점
    off.save_barrier("excel", FakeWorkbook(str(orig)), "스냅샷 실패", "excel", "write")  # 금요일: 지점 없음
    code, out = run_office(capsys, "undo")
    assert code == 1 and out["error_type"] == "needs_user_decision" and orig.read_bytes() == b"current"
    code, out = run_office(capsys, "undo", "--list")
    assert out["snapshots"][0]["undo_available"] is False and "undo_available" not in out["snapshots"][1]
    code, out = run_office(capsys, "undo", "--pick", "1")
    assert code == 1 and orig.read_bytes() == b"current"
    code, out = run_office(capsys, "undo", "--pick", "2")                                 # 사용자가 고른 앞 지점
    assert code == 0 and orig.read_bytes() == b"snapshot-of-a.xlsx"


def test_undo_refuses_when_file_saved_after_snapshot(tmp_path, monkeypatch, capsys):
    orig = _undo_env(tmp_path, monkeypatch)
    off.save_snapshot("excel", FakeWorkbook(str(orig)), "excel", "write")
    later = os.path.getmtime(orig) + 3600
    os.utime(orig, (later, later))                                               # 한 시간 뒤 사용자가 저장
    meta = off.list_snapshots()[0]
    assert meta["created_ts"] < later
    code, out = run_office(capsys, "undo")
    assert code == 1 and out["error_type"] == "needs_user_decision" and "--overwrite" in out["agent_directive"]
    assert orig.read_bytes() == b"current"
    code, out = run_office(capsys, "undo", "--overwrite")
    assert code == 0 and orig.read_bytes() == b"snapshot-of-a.xlsx"


# ---------------------------------------------------------------------------
# 이번 지점은 문서를 바꾸지 않았을 때만 지움
# ---------------------------------------------------------------------------


def test_closed_file_save_failure_drops_point(monkeypatch, tmp_path):
    def fail(doc, p):
        raise OSError("디스크 가득 참")
    path = _file_session(monkeypatch, tmp_path, fail)
    with pytest.raises(OSError):
        with off.open_target("excel", str(path), readonly=False, tool="excel", command="write") as t:
            t.doc.Saved = False
            t.set_focus()
    assert undo_points() == [] and path.read_bytes() == b"original"


def test_closed_file_error_after_change_drops_point_and_success_keeps_it(monkeypatch, tmp_path):
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
    assert undo_points() == []                                  # 디스크 파일은 그대로라 지점은 바뀐 것 없는 상태
    with off.open_target("excel", str(path), readonly=False, tool="excel", command="write") as t:
        t.doc.Saved = False
        t.set_focus()
    assert len(undo_points()) == 1 and path.read_bytes() == b"changed by the tool"


# ---------------------------------------------------------------------------
# 바꾼 것 없이 성공한 실행은 되돌리기 지점을 남기지 않음 (다음 undo가 바뀐 것 없는 지점으로 가지 않게)
# ---------------------------------------------------------------------------


def test_closed_file_noop_drops_point(monkeypatch, tmp_path):
    path = tmp_path / "닫힌.xlsx"
    path.write_bytes(b"original")
    stub_com(monkeypatch)
    monkeypatch.setattr(off, "_new_instance", lambda kind: FakeExcelApp())
    with off.open_target("excel", str(path), readonly=False, tool="excel", command="replace"):
        pass                                 # 아무것도 바꾸지 않음 -> 저장도 하지 않음
    assert undo_points() == [] and path.read_bytes() == b"original"


# ---------------------------------------------------------------------------
# undo --pick 0
# ---------------------------------------------------------------------------


def test_undo_pick_zero_rejected(tmp_path, capsys):
    orig = tmp_path / "a.xlsx"
    orig.write_bytes(b"current")
    off.save_snapshot("excel", FakeWorkbook(str(orig)))
    code, out = run_main(office, capsys, "undo", "--pick", "0")
    assert code == 1 and "1 이상" in out["message"], out
    assert orig.read_bytes() == b"current"
    code, out = run_main(office, capsys, "undo", "--pick", "-1")
    assert code == 1 and orig.read_bytes() == b"current"


# ---------------------------------------------------------------------------
# undo --list 되돌릴 수 없는 작업만
# ---------------------------------------------------------------------------


def test_undo_list_barrier_only_is_zero(capsys, monkeypatch):
    monkeypatch.setattr(off, "list_snapshots", lambda kind=None: [
        {"app": "word", "name": "새문서.docx", "barrier": True, "reason": "새로 만든 파일이라 되돌릴 이전 상태가 없음",
         "created_at": "20261103_101500", "by": "word write", "meta_file": "x.json"}])
    code, out = run_main(office, capsys, "undo", "--list")
    assert code == 0 and out["count"] == 0 and out["shown"] == 0 and out["barriers"] == 1, out
    assert out["agent_directive"].startswith("VERIFIED: 0건") and "command" not in out["agent_directive"], out


def test_undo_list_count_matches_preview(capsys, monkeypatch):
    snaps = [{"app": "excel", "name": "a.xlsx", "created_at": "20261103_101500", "by": "excel write", "meta_file": "1.json"},
             {"app": "word", "name": "b.docx", "barrier": True, "reason": "새 파일", "created_at": "20261103_101400",
              "by": "word write", "meta_file": "2.json"}]
    monkeypatch.setattr(off, "list_snapshots", lambda kind=None: snaps)
    code, out = run_main(office, capsys, "undo", "--list")
    assert code == 0 and out["count"] == 1 and out["shown"] == 1 and out["preview"] == "되돌릴 수 있는 지점 1개", out


# ---------------------------------------------------------------------------
# undo --list / 되돌릴 지점 없는 작업(장벽)
# ---------------------------------------------------------------------------


def test_undo_plain_barrier_only_is_zero(capsys, monkeypatch):
    monkeypatch.setattr(off, "list_snapshots", lambda kind=None: [_barrier(0)])
    code, out = run_as_tool(office, capsys, monkeypatch, "office", "undo")
    assert code == 0 and out["status"] == "ok" and out["count"] == 0 and out["restored"] is False, out
    assert out["agent_directive"].startswith("VERIFIED: 0건"), out


def test_undo_plain_barrier_then_undoable_still_asks(capsys, monkeypatch):
    snaps = [_barrier(0), {"app": "word", "name": "old.docx", "created_at": "20261103_100000", "by": "word write",
                           "meta_file": "o.json", "original_path": ""}]
    monkeypatch.setattr(off, "list_snapshots", lambda kind=None: snaps)
    code, out = run_as_tool(office, capsys, monkeypatch, "office", "undo")
    assert code == 1 and out["error_type"] == "needs_user_decision", out


def test_undo_list_many_barriers_shows_undoable(capsys, monkeypatch):
    snaps = [_barrier(i) for i in range(25)]
    snaps.append({"app": "word", "name": "old.docx", "created_at": "20261103_100000", "by": "word write",
                  "meta_file": "o.json", "original_path": ""})
    monkeypatch.setattr(off, "list_snapshots", lambda kind=None: snaps)
    code, out = run_as_tool(office, capsys, monkeypatch, "office", "undo", "--list")
    assert code == 0 and out["count"] == 1 and out["shown"] == 1 and out["barriers"] == 25, out
    undoable = [it for it in out["snapshots"] if "command" in it]
    assert len(undoable) == 1 and undoable[0]["pick"] == 26 and "--pick 26" in undoable[0]["command"], out["snapshots"]
    assert out["truncated"] is True and len(out["snapshots"]) == 21, out       # 앞 20개 되돌릴 수 없는 작업 + 지점 1개
    assert out["shown"] == sum(1 for it in out["snapshots"] if "command" in it)


def test_undo_list_more_than_20_undoable(capsys, monkeypatch):
    snaps = [{"app": "word", "name": f"d{i}.docx", "created_at": "20261103_100000", "by": "word write",
              "meta_file": f"{i}.json", "original_path": ""} for i in range(25)]
    snaps.insert(3, _barrier(99))
    monkeypatch.setattr(off, "list_snapshots", lambda kind=None: snaps)
    code, out = run_as_tool(office, capsys, monkeypatch, "office", "undo", "--list")
    assert out["count"] == 25 and out["shown"] == 20 and out["truncated"] is True and out["barriers"] == 1, out
    assert [it["pick"] for it in out["snapshots"]] == list(range(1, 22)), out["snapshots"]


# ---------------------------------------------------------------------------
# 0건 결과
# ---------------------------------------------------------------------------


def test_office_windows_zero(monkeypatch, capsys):
    monkeypatch.setattr(off, "open_documents_summary", lambda: [])
    monkeypatch.setattr(off, "running_office_caps", lambda: {})
    code, out = run_main(office, capsys, "windows")
    assert code == 0 and out["count"] == 0 and out["agent_directive"].startswith("VERIFIED: 0건"), out


def test_undo_list_zero(capsys):
    code, out = run_main(office, capsys, "undo", "--list")
    assert code == 0 and out["count"] == 0 and out["agent_directive"].startswith("VERIFIED: 0건"), out


# ---------------------------------------------------------------------------
# _office 세션: 사용자 창 미간섭, 읽기 전용, 화면 이동 규칙, 스냅샷, 대량 쓰기 설정 복원
# ---------------------------------------------------------------------------


def test_excel_file_read_uses_dedicated_instance_readonly(tmp_path, fake_com):
    path = tmp_path / "a.xlsx"
    path.write_bytes(b"x")
    app = FakeExcelApp()
    calls = fake_com["install"](dispatch_ex=app)
    with off.open_target("excel", str(path)) as t:
        assert t.source == "file" and t.readonly
        t.doc.Saved = False
    assert ("DispatchEx", "Excel.Application") in calls
    assert app.Visible is False and app.Workbooks.open_kwargs["ReadOnly"] is True
    wb = app.Workbooks.items[0]
    assert ("Save",) not in wb.calls and ("Close", False) in wb.calls
    assert app.quit_called


def test_word_save_as_falls_back_to_saveas_on_word2007():
    old, new = _Doc2007(), _Doc2010()
    off.word_save_as(old, "C:/x/a.docx", 16)
    off.word_save_as(new, "C:/x/a.docx", 16)
    off.word_save_as(new, "C:/x/b.docx")
    assert old.calls == [("SaveAs", "C:/x/a.docx", {"FileFormat": 16})]
    assert new.calls == [("SaveAs2", "C:/x/a.docx", {"FileFormat": 16}), ("SaveAs2", "C:/x/b.docx", {})]


@pytest.mark.parametrize("version,addin,expected", [
    ("16.0", False, True), ("14.0", False, True), ("12.0", False, False), ("12.0", True, True), ("11.0", True, False)])
def test_office_caps_pdf_export(monkeypatch, version, addin, expected):
    monkeypatch.setattr(off, "_office12_pdf_addin", lambda: addin)
    app = type("App", (), {"Version": version})()
    caps = off.office_caps(app)
    assert caps["version"] == version and caps["major"] == int(float(version)) and caps["pdf_export"] is expected


def test_require_pdf_export(monkeypatch):
    monkeypatch.setattr(off, "_office12_pdf_addin", lambda: False)
    with pytest.raises(rt.InputError) as ei:
        off.require_pdf_export(type("App", (), {"Version": "12.0"})(), ".docx")
    assert "PDF 내보내기 기능이 없습니다" in str(ei.value) and ".docx" in str(ei.value) and "2007" in str(ei.value)
    off.require_pdf_export(type("App", (), {"Version": "16.0"})(), ".docx")      # 가능: 예외 없음
    off.require_pdf_export(object(), ".docx")                                    # 버전 모름: 시도하도록 통과


def test_missing_file_readonly_suggests_open_documents(fake_com):
    fake_com["open"]["excel"] = [{"name": "10월보고.xlsx", "path": "C:/x/10월보고.xlsx", "app": None, "doc": None}]
    with pytest.raises(rt.InputError) as ei:
        with off.open_target("excel", "10월 보고.xlsx"):
            pass
    assert "[추천] 10월보고.xlsx" in str(ei.value)


def test_write_create_new_file_with_matching_format(tmp_path, fake_com):
    app = FakeExcelApp()
    fake_com["install"](dispatch_ex=app)
    target = tmp_path / "new.xlsm"
    with off.open_target("excel", str(target), readonly=False, create=True) as t:
        assert t.source == "new"
    wb = app.Workbooks.items[0]
    assert ("SaveAs", str(target), 52) in wb.calls


def test_write_without_create_does_not_make_files(tmp_path, fake_com):
    with pytest.raises(rt.InputError):
        with off.open_target("excel", str(tmp_path / "none.xlsx"), readonly=False):
            pass


def test_open_document_matched_by_name_is_not_closed(fake_com):
    wb = FakeWorkbook("C:/x/보고.xlsx")
    fake_com["open"]["excel"] = [{"name": "보고.xlsx", "path": "C:/x/보고.xlsx", "app": FakeExcelApp(), "doc": wb}]
    with off.open_target("excel", "보고.XLSX", readonly=False, snapshot=False) as t:
        assert t.source == "open" and t.doc is wb
        t.doc.Saved = False
    assert wb.calls == []          # 저장도 닫기도 하지 않음 (사용자 문서)


def test_same_name_documents_are_not_guessed(fake_com):
    # 다른 폴더의 같은 파일명 두 개: 첫 번째를 고르지 않고 [후보](전체 경로)
    pa, pb = os.path.abspath(os.path.join("x", "보고.xlsx")), os.path.abspath(os.path.join("y", "보고.xlsx"))
    a, b = FakeWorkbook(pa), FakeWorkbook(pb)
    fake_com["open"]["excel"] = [{"name": "보고.xlsx", "path": pa, "app": FakeExcelApp(), "doc": a},
                                 {"name": "보고.xlsx", "path": pb, "app": FakeExcelApp(), "doc": b}]
    with pytest.raises(rt.InputError) as ei:
        with off.open_target("excel", "보고.xlsx", readonly=False, snapshot=False):
            pass
    assert ei.value.candidates == [pa, pb] and "[후보]" in str(ei.value)
    payload = rt.error_payload(ei.value)
    assert payload["error_type"] == "ambiguous" and payload["agent_directive"].startswith("STOP_RETRY:")
    assert "재시도하지 말고" in payload["agent_directive"] and payload["candidates"] == [pa, pb]
    assert a.calls == [] and b.calls == []
    assert off.match_open("excel", pb, fake_com["open"]["excel"])["doc"] is b     # 전체 경로는 하나로 정해짐


def test_sheet_lookup_rules(fake_com):
    wb = FakeWorkbook("C:/x/a.xlsx", sheets=("매출", "요약"))
    fake_com["active"]["excel"] = {"name": "a.xlsx", "path": "C:/x/a.xlsx", "app": FakeExcelApp(), "doc": wb}
    with off.open_target("excel") as t:
        assert t.sheet("요약 ").Name == "요약"
        assert t.sheet().Name == "매출"
        with pytest.raises(rt.InputError) as ei:
            t.sheet("매출2", create=True)            # 읽기 세션은 만들지 않음
        assert "[후보] 매출" in str(ei.value) and "[추천]" not in str(ei.value)
    with off.open_target("excel", readonly=False, snapshot=False) as t:
        assert t.sheet("신규", create=True).Name == "신규"
        with pytest.raises(rt.InputError) as ei:
            t.sheet("신규2", allow_create=True)       # 쓰기 세션이어도 create 없이는 만들지 않음 (excel write)
        assert ei.value.candidates == ["신규"] and "--new-sheet" not in str(ei.value)   # 새 시트 여부는 사용자에게 (시나리오 10)
        assert isinstance(ei.value, rt.UserDecisionError) and "--new-sheet" in ei.value.directive
        with pytest.raises(rt.InputError) as ei:
            t.sheet("신규2")                          # --new-sheet가 없는 명령: 찾을 수 없음 [후보], 새 시트 안내 없음 (1.4.12)
        assert ei.value.candidates == ["신규"] and not isinstance(ei.value, rt.UserDecisionError)
        assert "--new-sheet" not in str(ei.value) and "--new-sheet" not in str(ei.value.directive or "")


def test_read_session_never_moves_view_write_moves_to_changed_sheet(fake_com):
    wb = FakeWorkbook("C:/x/a.xlsx", sheets=("A", "B"))
    fake_com["active"]["excel"] = {"name": "a.xlsx", "path": "C:/x/a.xlsx", "app": FakeExcelApp(), "doc": wb}
    with off.open_target("excel") as t:
        t.set_focus(sheet=t.sheet("B"))
    assert wb.Worksheets("B").activated == 0
    with off.open_target("excel", readonly=False, snapshot=False, focus=False) as t:
        t.set_focus(sheet=t.sheet("B"))
    assert wb.Worksheets("B").activated == 0
    with off.open_target("excel", readonly=False, snapshot=False) as t:
        t.set_focus(sheet=t.sheet("B"))
        assert t.resolved()["sheet"] == "B"
    assert wb.Worksheets("B").activated == 1


def test_exception_in_session_does_not_save_file(tmp_path, fake_com):
    path = tmp_path / "a.xlsx"
    path.write_bytes(b"x")
    app = FakeExcelApp()
    fake_com["install"](dispatch_ex=app)
    with pytest.raises(RuntimeError):
        with off.open_target("excel", str(path), readonly=False, snapshot=False) as t:
            t.doc.Saved = False
            raise RuntimeError("도중 실패")
    wb = app.Workbooks.items[0]
    assert ("Save",) not in wb.calls and ("Close", False) in wb.calls and app.quit_called


def test_ppt_file_mode_leaves_users_instance_alone(tmp_path, fake_com):
    path = tmp_path / "deck.pptx"
    path.write_bytes(b"x")
    app = FakePptApp()
    fake_com["install"](dispatch=app, active=app)
    with off.open_target("ppt", str(path)) as t:
        pass
    assert app.touched == [] and app.DisplayAlerts == 2 and not app.quit_called
    assert app.Presentations.open_args[1:] == (-1, 0, 0)
    assert t is not None


def test_write_session_takes_snapshot_first(tmp_path, fake_com):
    wb = FakeWorkbook(str(tmp_path / "보고.xlsm"))
    (tmp_path / "보고.xlsm").write_bytes(b"x")
    fake_com["active"]["excel"] = {"name": wb.Name, "path": wb.FullName, "app": FakeExcelApp(), "doc": wb}
    with off.open_target("excel", readonly=False, tool="excel", command="write") as t:
        t.changed = True                     # 바꾼 실행만 지점을 남김 (1.4.14: 바꾼 것 없는 성공은 지점을 지움)
    snaps = off.list_snapshots("excel")
    assert len(snaps) == 1 and snaps[0]["snapshot_file"].endswith(".xlsm")
    assert snaps[0]["by"] == "excel write" and snaps[0]["original_path"] == wb.FullName


def test_word_snapshot_unsaved_uses_flat_xml_without_touching_original(tmp_path):
    orig = tmp_path / "a.docx"
    orig.write_bytes(b"disk")
    off.save_snapshot("word", FakeWordDoc(str(orig), saved=False))
    snap = off.list_snapshots("word")[0]
    assert snap["snapshot_format"] == "flat_xml" and snap["snapshot_file"].endswith(".xml")
    off.save_snapshot("word", FakeWordDoc(str(orig), saved=True))
    assert off.list_snapshots("word")[0]["snapshot_file"].endswith(".docx")


def test_snapshots_of_a_replaced_document_are_dropped(tmp_path):
    # 실기 C7: 같은 경로에 파일을 새로 만들어도 이전 파일의 스냅샷이 undo 대상에 남았음
    path = str(tmp_path / "A.xlsx")
    _xlsx_created(path, (2026, 1, 1, 9, 0, 0))
    off.save_snapshot("excel", FakeWorkbook(path))
    snaps = off.list_snapshots("excel")
    assert len(snaps) == 1 and snaps[0]["doc_created"] == off.doc_identity(path) and "2026-01-01" in snaps[0]["doc_created"]

    import openpyxl                                         # 같은 문서를 다시 저장: 만든 시각 유지 -> 스냅샷 유지
    wb = openpyxl.load_workbook(path)
    wb.active["A1"] = "수정"
    wb.save(path)
    assert len(off.list_snapshots("excel")) == 1

    _xlsx_created(path, (2026, 5, 5, 10, 0, 0))                     # 같은 경로에 새 문서 -> 옛 스냅샷 정리
    assert off.list_snapshots("excel") == []
    assert not [f for f in os.listdir(off._snapshot_dir()) if f.startswith("excel_")]


def test_snapshot_identity_unknown_keeps_snapshot(tmp_path):
    path = str(tmp_path / "old.xls")                        # OOXML이 아니면 비교하지 않음
    with open(path, "wb") as f:
        f.write(b"x")
    off.save_snapshot("excel", FakeWorkbook(path))
    snaps = off.list_snapshots("excel")
    assert len(snaps) == 1 and snaps[0]["doc_created"] is None


def test_snapshot_rotation_and_size_limit(tmp_path, monkeypatch):
    path = tmp_path / "a.xlsx"
    path.write_bytes(b"x")
    for _ in range(off.SNAPSHOT_KEEP + 3):
        off.save_snapshot("excel", FakeWorkbook(str(path)))
    assert len(off.list_snapshots("excel")) == off.SNAPSHOT_KEEP
    monkeypatch.setattr(off, "SNAPSHOT_MAX_BYTES", 0)
    assert off.save_snapshot("excel", FakeWorkbook(str(path))) is None


def test_excel_batch_restores_settings():
    app = FakeExcelApp()
    with off.excel_batch(app):
        assert app.ScreenUpdating is False and app.Calculation == -4135 and app.EnableEvents is False
    assert app.ScreenUpdating is True and app.Calculation == -4105 and app.EnableEvents is True
    with pytest.raises(ValueError):
        with off.excel_batch(app):
            raise ValueError()
    assert app.ScreenUpdating is True


def test_range_to_matrix():
    assert off.range_to_matrix(3) == [[3]]
    assert off.range_to_matrix((("a", "b"),)) == [["a", "b"]]
    assert off.range_to_matrix((("a",), ("b",))) == [["a"], ["b"]]


def test_require_com_message(monkeypatch):
    monkeypatch.setattr(off, "win32com", None)
    with pytest.raises(rt.InputError) as ei:
        off.require_com()
    assert "reader.py" in str(ei.value)


def test_write_session_without_snapshot_leaves_barrier(tmp_path, fake_com, monkeypatch):
    wb = FakeWorkbook(str(tmp_path / "보고.xlsx"))
    (tmp_path / "보고.xlsx").write_bytes(b"x")
    fake_com["active"]["excel"] = {"name": wb.Name, "path": wb.FullName, "app": FakeExcelApp(), "doc": wb}
    monkeypatch.setattr(off, "SNAPSHOT_MAX_BYTES", 0)                    # 스냅샷 실패
    with off.open_target("excel", readonly=False, tool="excel", command="write") as t:
        t.changed = True
    snaps = off.list_snapshots("excel")
    assert len(snaps) == 1 and snaps[0]["barrier"] is True and "스냅샷" in snaps[0]["reason"]


def test_write_target_typo_of_open_document_does_not_create_new_file(tmp_path, fake_com, monkeypatch):
    monkeypatch.chdir(tmp_path)
    monkeypatch.setattr(off, "list_open", lambda kind: [{"name": "주간보고.docx", "path": "C:/x/주간보고.docx"}])
    monkeypatch.setattr(off, "match_open", lambda kind, target, docs=None, option="--target": None)
    with pytest.raises(rt.InputError) as ei:
        with off.open_target("word", "주간보고서.docx", readonly=False, create=True):
            pass
    assert "[후보] 주간보고.docx" in str(ei.value) and not (tmp_path / "주간보고서.docx").exists()


# ---------------------------------------------------------------------------
# 실환경 견고성 (암호 파일, 읽기 전용으로 열린 파일, Office 바쁨)
# ---------------------------------------------------------------------------


def test_background_open_passes_dummy_password_and_disables_macros(tmp_path, monkeypatch):
    app = FakeExcelApp()
    wb = off._open_file("excel", app, str(tmp_path / "a.xlsx"), readonly=True)
    kw = app.Workbooks.open_kwargs
    assert kw["Password"] == off._NO_PASSWORD and kw["IgnoreReadOnlyRecommended"] is True and wb is not None


def test_password_error_becomes_input_error(tmp_path):
    class _App:
        class Workbooks:
            @staticmethod
            def Open(path, **kw):
                raise Exception("입력한 암호가 잘못되었습니다.")
    with pytest.raises(rt.InputError) as ei:
        off._open_file("excel", _App(), str(tmp_path / "secret.xlsx"), readonly=True)
    assert "열기 암호" in str(ei.value)


def test_word_open_skips_password_for_plain_docx(tmp_path):
    # Word 2007: ReadOnly + PasswordDocument 조합은 암호 없는 docx도 "명령을 정상적으로 종료할 수 없습니다"로 실패
    path = tmp_path / "a.docx"
    path.write_bytes(b"PK\x03\x04 zip")
    app = _WordApp()
    assert off._open_file("word", app, str(path), readonly=False) == "doc"
    assert "PasswordDocument" not in app.kw and app.kw["ReadOnly"] is False


def test_word_open_keeps_password_when_unknown_and_reports_office_error(tmp_path):
    path = tmp_path / "drm.docx"
    path.write_bytes(b"DRM-HEADER")                      # 판별할 수 없는 파일: 암호 창을 막되 오류는 그대로
    assert off._word_encrypted(str(path)) is None
    app = _WordApp()
    off._open_file("word", app, str(path), readonly=True)
    assert app.kw["PasswordDocument"] == off._NO_PASSWORD
    with pytest.raises(Exception) as ei:
        off._open_file("word", _WordApp(fail=RuntimeError("다른 오류")), str(path), readonly=True)
    assert not isinstance(ei.value, rt.InputError)


def test_word_encrypted_file_error_without_password_word_becomes_input_error(tmp_path, monkeypatch):
    # Word 2007의 틀린 암호 오류 문구에는 '암호'가 없음
    monkeypatch.setattr(off, "_word_encrypted", lambda p: True)
    with pytest.raises(rt.InputError) as ei:
        off._open_file("word", _WordApp(fail=RuntimeError("명령을 정상적으로 종료할 수 없습니다.")),
                       str(tmp_path / "s.docx"), readonly=True)
    assert "열기 암호" in str(ei.value)


def test_write_to_file_opened_read_only_stops_before_writing(tmp_path, fake_com, monkeypatch):
    path = tmp_path / "공유.xlsx"
    path.write_bytes(b"x")
    real_open = off._open_file

    def locked(kind, app, p, readonly):
        doc = real_open(kind, app, p, readonly)
        doc.ReadOnly = True                         # 다른 사람이 열고 있어 읽기 전용으로 열림
        return doc
    monkeypatch.setattr(off, "_open_file", locked)
    monkeypatch.setattr(off, "_new_instance", lambda kind: FakeExcelApp())
    with pytest.raises(rt.InputError) as ei:
        with off.open_target("excel", str(path), readonly=False):
            raise AssertionError("쓰기 본문까지 오면 안 됨")
    assert "읽기 전용" in str(ei.value)


def test_busy_message_filter_retries_then_gives_up(monkeypatch):
    monkeypatch.setattr(rt, "RECOVERED", [])
    monkeypatch.setattr(rt, "RECOVERY_ATTEMPTS", [])
    f = off.BusyMessageFilter()
    assert f.RetryRejectedCall(0, 100, 2) == f.RETRY_DELAY_MS
    assert rt.RECOVERED == [f.NOTE]
    assert f.RetryRejectedCall(0, f.BUSY_RETRY_MS + 1, 2) == -1
    assert rt.RECOVERED == [] and rt.RECOVERY_ATTEMPTS


def test_no_open_document_asks_for_the_path(monkeypatch):
    monkeypatch.setattr(off, "find_apps", lambda kind: [])
    with pytest.raises(rt.UserDecisionError) as ei:
        off.get_active("excel")
    assert "--target" in ei.value.directive and "추측하지" in ei.value.directive


# ---------------------------------------------------------------------------
# 관리자 권한 실행 (1.4.40): 일반 권한으로 연 문서가 보이지 않으므로 '문서 없음' 대신 권한 안내
# ---------------------------------------------------------------------------


def _elevated(monkeypatch, value=True):
    monkeypatch.setattr(rt, "is_elevated", lambda: value)


def _fake_token_windll(kind_value, ok=1):
    """OpenProcessToken/GetTokenInformation(TokenElevationType)을 흉내 내는 가짜 ctypes.windll."""
    import ctypes

    class _Advapi:
        def OpenProcessToken(self, proc, access, out):
            return 1

        def GetTokenInformation(self, token, cls, buf, size, outsize):
            assert cls == 18
            ctypes.cast(buf, ctypes.POINTER(ctypes.c_int)).contents.value = kind_value
            return ok

    class _Kernel:
        def GetCurrentProcess(self):
            return 0

        def CloseHandle(self, h):
            return 1

    class _Windll:
        advapi32 = _Advapi()
        kernel32 = _Kernel()
    return _Windll()


def test_is_elevated_only_for_uac_split_token(monkeypatch):
    """UAC가 꺼진 PC(EnableLUA=0)는 모든 프로그램이 같은 권한이라 Office 문서가 보이므로 권한 안내를 하지 않는다 (1.4.74).
    권한 안내는 UAC 분리 토큰을 올린 관리자 창(TokenElevationTypeFull=2)일 때만."""
    import ctypes
    monkeypatch.setattr(rt.os, "name", "nt")
    for kind, want in ((1, False), (2, True), (3, False)):
        monkeypatch.setattr(ctypes, "windll", _fake_token_windll(kind), raising=False)
        assert rt.is_elevated() is want, kind
    monkeypatch.setattr(ctypes, "windll", _fake_token_windll(2, ok=0), raising=False)
    assert rt.is_elevated() is False


def test_is_elevated_false_off_windows():
    if os.name != "nt":
        assert rt.is_elevated() is False and rt.elevation_hint() == ""


@pytest.mark.parametrize("tool,argv", [(excel, ("read",)), (excel, ("write", "--range", "A1", "--value", "1")),
                                       (ppt, ("read",))])
def test_no_open_document_when_elevated_explains_admin(monkeypatch, capsys, tool, argv):
    stub_com(monkeypatch)
    monkeypatch.setattr(off, "find_apps", lambda kind: [])
    _elevated(monkeypatch)
    code, out = run_main(tool, capsys, *argv)
    assert code == 1 and out["error_type"] == "elevated_process", out
    assert "관리자 권한" in out["message"] and "일반 PowerShell" in out["message"]
    assert "문서가 없습니다" not in out["message"]                 # 없다고 단정하지 않음
    assert out["agent_directive"].startswith("STOP_RETRY:") and "다시 실행하지" in out["agent_directive"]
    assert "--target" not in out["agent_directive"]                # 경로를 줘도 해결되지 않음


def test_no_open_document_not_elevated_unchanged(monkeypatch, capsys):
    stub_com(monkeypatch)
    monkeypatch.setattr(off, "find_apps", lambda kind: [])
    _elevated(monkeypatch, False)
    code, out = run_main(excel, capsys, "read")
    assert code == 1 and out["error_type"] == "needs_user_decision"
    assert "열려 있는 Excel 문서가 없습니다" in out["message"] and "관리자" not in out["message"]


@pytest.mark.parametrize("tool,argv", [(excel, ("info", "--all")), (office, ("windows",)),
                                       (ppt, ("info", "--all"))])
def test_empty_open_list_when_elevated_is_unverified(monkeypatch, capsys, tool, argv):
    stub_com(monkeypatch)
    monkeypatch.setattr(off, "find_apps", lambda kind: [])
    _elevated(monkeypatch)
    code, out = run_main(tool, capsys, *argv)
    assert code == 0 and out["verified"] is False and out["count"] == 0, out
    assert out["agent_directive"].startswith("UNVERIFIED:") and "관리자 권한" in out["agent_directive"]
    _elevated(monkeypatch, False)
    code, out = run_main(tool, capsys, *argv)
    assert code == 0 and out["agent_directive"].startswith("VERIFIED: 0건")


def test_open_document_name_when_elevated_mentions_admin(monkeypatch, capsys):
    stub_com(monkeypatch)
    monkeypatch.setattr(off, "find_apps", lambda kind: [])
    _elevated(monkeypatch)
    code, out = run_main(excel, capsys, "read", "--target", "보고.xlsx")
    assert code == 1 and "관리자 권한" in out["message"], out


def test_transfer_no_target_open_when_elevated(env, capsys, monkeypatch):
    env["word"] = None
    monkeypatch.setattr(off, "get_active", lambda kind: (_ for _ in ()).throw(off.no_open_doc(kind))
                        if kind != "excel" else {"name": "a.xlsx", "path": "", "app": FakeXlApp(), "doc": None})
    _elevated(monkeypatch)
    code, out = run_office(capsys, *CHART)
    assert code == 1 and out["error_type"] == "elevated_process" and "관리자 권한" in out["message"], out


# ---------------------------------------------------------------------------
# 환경 변수·~ 가 든 전체 경로는 새 파일로 만들 수 있음, 전체 경로가 없으면 열린 문서를 [후보]로 주지 않음
# ---------------------------------------------------------------------------


def test_expanded_absolute_path_can_be_created(monkeypatch, tmp_path):
    stub_com(monkeypatch)
    monkeypatch.setenv("OCLI_TEST_HOME", str(tmp_path))
    monkeypatch.setenv("HOME", str(tmp_path))
    monkeypatch.setattr(off, "match_open", lambda kind, target, docs=None, option="--target": None)

    def new_instance(kind):
        raise _Created()
    monkeypatch.setattr(off, "_new_instance", new_instance)
    for target in ("$OCLI_TEST_HOME/새.xlsx", "~/새.xlsx"):
        with pytest.raises(_Created):                     # '찾을 수 없음'이 아니라 새 파일 만들기로 진행
            with off.open_target("excel", target, readonly=False, create=True):
                pass


def test_missing_absolute_path_is_not_ambiguous(monkeypatch, tmp_path):
    stub_com(monkeypatch)
    monkeypatch.setattr(off, "list_open", lambda kind: [{"name": "보고.xlsx", "path": "C:\\docs\\보고.xlsx"},
                                                        {"name": "예산.xlsx", "path": "C:\\docs\\예산.xlsx"}])
    monkeypatch.setattr(off, "match_open", lambda kind, target, docs=None, option="--target": None)
    with pytest.raises(rt.InputError) as ei:
        with off.open_target("excel", str(tmp_path / "없음.xlsx"), readonly=True):
            pass
    assert "[후보]" not in str(ei.value) and ei.value.candidates != ["보고.xlsx", "예산.xlsx"]
    with pytest.raises(rt.InputError) as ei:                 # 파일 이름만이면 그대로 [후보]
        with off.open_target("excel", "없음.xlsx", readonly=True):
            pass
    assert sorted(ei.value.candidates) == ["보고.xlsx", "예산.xlsx"]


# ---------------------------------------------------------------------------
# 도구가 연 닫힌 파일은 바꾼 것(t.changed)이 있을 때만 저장 (휘발성 수식으로 열자마자 저장 안 된 상태여도)
# ---------------------------------------------------------------------------


def test_closed_volatile_file_noop_is_not_resaved(monkeypatch, tmp_path):
    path = tmp_path / "vol.xlsx"
    path.write_bytes(b"original")
    stub_com(monkeypatch)
    _volatile_open(monkeypatch)
    with off.open_target("excel", str(path), readonly=False, tool="excel", command="replace"):
        pass                                 # 아무것도 바꾸지 않음
    assert undo_points() == [] and path.read_bytes() == b"original"


def test_closed_volatile_file_real_change_is_saved(monkeypatch, tmp_path):
    path = tmp_path / "vol.xlsx"
    path.write_bytes(b"original")
    stub_com(monkeypatch)
    _volatile_open(monkeypatch)
    with off.open_target("excel", str(path), readonly=False, tool="excel", command="write") as t:
        t.set_focus()                        # 바꾼 명령은 set_focus(또는 t.changed = True)로 표시
    assert path.read_bytes() == b"resaved" and len(undo_points()) == 1


# ---------------------------------------------------------------------------
# 기본 파일 이름
# ---------------------------------------------------------------------------


def test_chart_export_with_slash_in_name():
    class _Chart:
        def Export(self, path, fmt):
            with open(path, "wb") as f:
                f.write(office.PNG_SIGNATURE + b"0" * 10)

    class _Obj:
        Chart = _Chart()
    stem = rt.safe_filename("chart_1/4분기 매출", 40)
    path = office.image_by_export(_Obj(), stem)
    assert os.path.dirname(path) == rt.get_work_dir() and "/4" not in os.path.basename(path)
    path = office.image_by_export(_Obj(), "chart_1/4분기 매출")     # 정리하지 않은 이름도 resolve_output_path가 정리
    assert os.path.dirname(path) == rt.get_work_dir()


# ---------------------------------------------------------------------------
# 닫힌 파일 열기 감시 (B4, 1.4.58): DRM 인증 창 등으로 멈추면 그 인스턴스만 끝내고 drm_or_dialog
# ---------------------------------------------------------------------------


@pytest.fixture
def hung_open(fake_com, monkeypatch, tmp_path):
    import threading
    killed = []
    gone = threading.Event()
    monkeypatch.setenv("OFFICE_CLI_OPEN_TIMEOUT", "0.2")
    monkeypatch.setattr(off, "_new_instance", lambda kind: FakeExcelApp())
    monkeypatch.setattr(off, "_instance_pid", lambda kind, app, before: 4242)

    def kill(pid):
        killed.append(pid)
        gone.set()
        return True
    monkeypatch.setattr(off, "_kill_pid", kill)

    def stuck(kind, app, path, readonly):
        if not gone.wait(5):                       # 인증 창이 떠서 돌아오지 않는 Open (프로세스가 끝나면 COM 오류)
            raise AssertionError("감시가 프로세스를 끝내지 않음")
        raise Exception("com_error: RPC 서버를 사용할 수 없습니다")
    monkeypatch.setattr(off, "_open_file", stuck)
    path = tmp_path / "보안 문서.xls"
    path.write_bytes(b"x")
    return path, killed


def test_open_watchdog_kills_only_its_instance(hung_open, capsys):
    path, killed = hung_open
    t0 = time.monotonic()
    code, out = run_main(excel, capsys, "read", "--target", str(path), "--range", "A1:B2")
    assert time.monotonic() - t0 < 4
    assert code == 1 and out["error_type"] == "drm_or_dialog" and killed == [4242], out
    assert out["agent_directive"].startswith("STOP_RETRY:")
    assert "문서 보안 인증 창 등으로 열리지 않음" in out["message"] and "직접 열어 두면 열린 문서로 읽음" in out["message"]
    cmd = out["next_commands"][0]["command"]
    assert "excel.py" in cmd and "read" in cmd and "보안 문서.xls" in cmd and str(tmp_path_of(path)) not in cmd


def tmp_path_of(path):
    return os.path.dirname(str(path))


def test_open_watchdog_quiet_when_open_is_fast(fake_com, monkeypatch, tmp_path):
    killed = []
    monkeypatch.setenv("OFFICE_CLI_OPEN_TIMEOUT", "5")
    monkeypatch.setattr(off, "_new_instance", lambda kind: FakeExcelApp())
    monkeypatch.setattr(off, "_instance_pid", lambda kind, app, before: 4242)
    monkeypatch.setattr(off, "_kill_pid", lambda pid: killed.append(pid) or True)
    path = tmp_path / "a.xls"
    path.write_bytes(b"x")
    with off.open_target("excel", str(path), readonly=True) as t:
        assert t.doc is not None
    assert killed == []


@pytest.fixture
def hung_doc_open(fake_com, monkeypatch, tmp_path):
    """Word·PowerPoint 열기가 멈춤. 감시가 끝내면(kill) 또는 끝내지 않고 도구를 끝내면(_abandon) 멈춘 Open이 돌아옴."""
    import threading
    from fake_office import FakePptApp
    rec = {"killed": [], "abandoned": []}
    gone = threading.Event()
    monkeypatch.setenv("OFFICE_CLI_OPEN_TIMEOUT", "0.2")
    monkeypatch.setattr(off, "_new_instance", lambda kind: FakePptApp() if kind == "ppt" else FakeExcelApp())
    monkeypatch.setattr(off, "_office_pids", lambda kind: set())
    monkeypatch.setattr(off, "_kill_pid", lambda pid: rec["killed"].append(pid) or gone.set() or True)
    monkeypatch.setattr(off, "_abandon", lambda err, kind, tool: rec["abandoned"].append(str(err)) or gone.set())

    def stuck(kind, app, path, readonly):
        if not gone.wait(5):
            raise AssertionError("감시가 아무것도 하지 않음")
        raise Exception("com_error: RPC 서버를 사용할 수 없습니다")
    monkeypatch.setattr(off, "_open_file", stuck)
    return tmp_path, rec


@pytest.mark.parametrize("kind,ext,running,pid,user_windows,kill", [
    ("ppt", ".pptx", False, 4242, True, False),     # 감시 중 사용자가 같은 PowerPoint 프로세스에서 발표 파일을 엶 (1.4.63)
    ("ppt", ".pptx", False, 4242, None, False),     # 창을 확인하지 못함
    ("ppt", ".pptx", False, 4242, False, True),     # 도구가 띄웠고 사용자 문서 없음
    ("ppt", ".pptx", True, None, False, False),     # 사용자가 이미 띄운 PowerPoint: 끝내지 않고 오류만 (전에는 감시 없음)
    ("word", ".docx", False, 4242, True, False),
    ("word", ".docx", False, 4242, False, True),
    ("word", ".docx", False, None, False, False),   # 띄운 프로세스를 가려내지 못함
])
def test_open_watchdog_kills_only_without_user_documents(hung_doc_open, monkeypatch, kind, ext, running, pid,
                                                         user_windows, kill):
    tmp_path, rec = hung_doc_open
    monkeypatch.setattr(off, "_ppt_running", lambda: running)
    monkeypatch.setattr(off, "_instance_pid", lambda k, app, before: pid)
    monkeypatch.setattr(off, "_user_windows", lambda k, p: user_windows)
    path = tmp_path / f"보안 문서{ext}"
    path.write_bytes(b"x")
    with pytest.raises(Exception):
        with off.open_target(kind, str(path), readonly=True, tool=kind):
            pass
    if kill:
        assert rec["killed"] == [4242] and rec["abandoned"] == []
    else:
        assert rec["killed"] == [] and len(rec["abandoned"]) == 1
        msg = rec["abandoned"][0]
        assert "끝내지 않았습니다" in msg and "OFFICE_CLI_OPEN_TIMEOUT" in msg


class _HiddenWordApp:
    """도구가 띄운 숨은 Word. user_docs: 그 인스턴스에서 사용자가 연 문서 수. busy: Documents.Count가 실패하는 횟수."""

    def __init__(self, user_docs, busy=0):
        self.docs = [FakeWordDoc(f"C:\\u\\사용자{i}.docx", saved=False) for i in range(user_docs)]
        self.quit_calls = []
        self.busy = busy
        app = self

        class _Docs:
            @property
            def Count(self):
                if app.busy:
                    app.busy -= 1
                    raise RuntimeError("Call was rejected by callee")      # COM 바쁨
                return len(app.docs)

            def Item(self, k):
                return app.docs[k - 1]
        self.Documents = _Docs()

    def Quit(self, **k):
        self.quit_calls.append(k)


@pytest.mark.parametrize("user_docs,close_fails,busy,quits,note", [
    (0, False, 0, True, ""),
    (1, False, 0, False, "다른 문서 1개"),
    (0, True, 0, True, ""),            # 자기 문서 닫기가 실패해 남아도 그 문서뿐이면 끔 (1.4.66)
    (1, True, 0, False, "다른 문서 1개"),
    (0, False, 2, True, ""),           # 문서 수 읽기가 COM 바쁨으로 실패하면 짧게 재시도 (1.4.66)
    (0, False, 99, False, "읽지 못해"),
])
def test_word_hidden_instance_quits_only_when_no_documents_left(fake_com, monkeypatch, tmp_path, capsys,
                                                                user_docs, close_fails, busy, quits, note):
    # 회귀: 정상 종료 때 Quit(SaveChanges=0)을 무조건 불러 사용자가 그 숨은 Word에서 연 문서의 변경을 버림
    app = _HiddenWordApp(user_docs)
    monkeypatch.setenv("OFFICE_CLI_OPEN_TIMEOUT", "0")
    monkeypatch.setattr(off, "_new_instance", lambda kind: app)
    monkeypatch.setattr(off, "QUIT_COUNT_DELAY", 0)
    path = tmp_path / "a.docx"

    def opened(kind, a, p, readonly):
        d = FakeWordDoc(str(path))

        def close(SaveChanges=None):
            if close_fails:
                raise RuntimeError("닫기 실패")
            a.docs.remove(d)
        a.docs.append(d)
        d.Close = close
        return d
    monkeypatch.setattr(off, "_open_file", opened)
    path.write_bytes(b"PK")
    with off.open_target("word", str(path), readonly=True):
        app.busy = busy
    assert bool(app.quit_calls) is quits
    err = capsys.readouterr().err
    assert ("[INFO]" in err and "끝내지 않고" in err) is (not quits)
    assert note in err


def test_open_watchdog_leaves_running_ppt_alerts(hung_doc_open, monkeypatch):
    # 회귀: 사용자가 띄운 PowerPoint의 DisplayAlerts를 바꾼 채 _abandon(os._exit) -> 되돌리지 않았음
    tmp_path, rec = hung_doc_open
    sets = []

    class _Watched(FakePptApp):
        def __setattr__(self, key, value):
            if key == "DisplayAlerts" and hasattr(self, "DisplayAlerts"):
                sets.append(value)
            super().__setattr__(key, value)
    app = _Watched()
    monkeypatch.setattr(off, "_new_instance", lambda kind: app)
    monkeypatch.setattr(off, "_ppt_running", lambda: True)
    monkeypatch.setattr(off, "_instance_pid", lambda k, a, before: None)
    path = tmp_path / "보안.pptx"
    path.write_bytes(b"x")
    with pytest.raises(Exception):
        with off.open_target("ppt", str(path), readonly=True, tool="ppt"):
            pass
    assert sets == [] and app.DisplayAlerts == 2 and not app.quit_called      # 바꾸지 않음 (abandon이면 되돌릴 수 없음)
    assert "POWERPNT.EXE" in rec["abandoned"][0] and "이미 떠 있던" in rec["abandoned"][0]


def test_open_watchdog_abandon_emits_one_json_line(monkeypatch, capsys, isolated_dirs):
    exits = []
    monkeypatch.setattr(os, "_exit", lambda code: exits.append(code))
    monkeypatch.setattr(rt, "CURRENT_ARGV", ["read", "--target", "a.pptx"])
    err = off._drm_or_dialog_error("ppt", "C:\\x\\a.pptx", "ppt", 30, killed=False)
    off._abandon(err, "ppt", "ppt")
    out = capsys.readouterr().out.strip().splitlines()
    assert exits == [1] and len(out) == 1
    payload = json.loads(out[0])
    assert payload["error_type"] == "drm_or_dialog" and payload["agent_directive"].startswith("STOP_RETRY:")
    assert "열기가 오래 걸린 경우일 수도 있습니다" in payload["message"]


def test_open_timeout_env(monkeypatch):
    monkeypatch.setenv("OFFICE_CLI_OPEN_TIMEOUT", "")
    assert off.open_timeout() == 30.0
    monkeypatch.setenv("OFFICE_CLI_OPEN_TIMEOUT", "0")
    assert off.open_timeout() == 0.0
