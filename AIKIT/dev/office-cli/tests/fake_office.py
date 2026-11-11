"""COM 없이 _office 세션/스냅샷 로직과 도구 로직을 검증하기 위한 가짜 Office 객체."""
import types


class Recorder:
    def __init__(self):
        self.calls = []

    def rec(self, *item):
        self.calls.append(item)


# ---------------------------------------------------------------------------
# Excel
# ---------------------------------------------------------------------------

class FakeRange:
    def __init__(self, ws, address):
        self.ws = ws
        self.address = address

    def __getattr__(self, name):
        raise AttributeError(name)


class FakeSheet:
    def __init__(self, wb, name):
        self.wb = wb
        self._name = name
        self.UsedRange = object()
        self.activated = 0

    @property
    def Name(self):
        return self._name

    @Name.setter
    def Name(self, v):
        from fake_excel import check_fake_sheet_name
        check_fake_sheet_name(v)          # 실제 Excel처럼 쓸 수 없는 이름(History 등)은 COM 오류
        self._name = v

    def Activate(self):
        self.activated += 1
        self.wb.active_sheet = self


class FakeSheets:
    def __init__(self, wb, names):
        self.wb = wb
        self.items = [FakeSheet(wb, n) for n in names]

    @property
    def Count(self):
        return len(self.items)

    def __call__(self, key):
        if isinstance(key, int):
            return self.items[key - 1]
        for s in self.items:
            if s.Name == key:
                return s
        raise Exception("no sheet")

    def __iter__(self):
        return iter(list(self.items))

    def Add(self, After=None):
        s = FakeSheet(self.wb, f"Sheet{len(self.items) + 1}")
        self.items.append(s)
        self.wb.Saved = False
        return s


class FakeWorkbook(Recorder):
    def __init__(self, full_name="", sheets=("Sheet1",)):
        super().__init__()
        self.FullName = full_name
        self.Name = full_name.replace("\\", "/").split("/")[-1] if full_name else "Book1"
        self.Saved = True
        self.Worksheets = FakeSheets(self, sheets)
        self.active_sheet = self.Worksheets(1)

    @property
    def ActiveSheet(self):
        return self.active_sheet

    def Save(self):
        self.rec("Save")
        self.Saved = True

    def SaveAs(self, path, FileFormat=None):
        self.rec("SaveAs", path, FileFormat)
        self.FullName = path
        self.Name = path.replace("\\", "/").split("/")[-1]

    def SaveCopyAs(self, path):
        self.rec("SaveCopyAs", path)
        with open(path, "wb") as f:
            f.write(b"snapshot-of-" + self.Name.encode())

    def Close(self, SaveChanges=None):
        self.rec("Close", SaveChanges)


class FakeWorkbooks:
    def __init__(self, app):
        self.app = app
        self.items = []
        self.open_kwargs = None

    @property
    def Count(self):
        return len(self.items)

    def __call__(self, i):
        return self.items[i - 1]

    def Open(self, path, **kw):
        self.open_kwargs = kw
        wb = FakeWorkbook(path, sheets=("Sheet1", "Sheet2"))
        self.items.append(wb)
        return wb

    def Add(self, Template=None):
        wb = FakeWorkbook("")
        wb.template = Template
        self.items.append(wb)
        return wb


class FakeExcelApp(Recorder):
    def __init__(self):
        super().__init__()
        self.Visible = True
        self.DisplayAlerts = True
        self.AskToUpdateLinks = True
        self.ScreenUpdating = True
        self.EnableEvents = True
        self.Calculation = -4105
        self.Workbooks = FakeWorkbooks(self)
        self.quit_called = False

    def Quit(self):
        self.quit_called = True


# ---------------------------------------------------------------------------
# Word
# ---------------------------------------------------------------------------

class FakeWordDoc(Recorder):
    def __init__(self, full_name, saved=True):
        super().__init__()
        self.FullName = full_name
        self.Name = full_name.replace("\\", "/").split("/")[-1]
        self.Saved = saved
        self.WordOpenXML = "<?xml version='1.0'?><pkg:package/>"

    def Save(self):
        self.rec("Save")

    def SaveAs(self, *a, **k):
        raise AssertionError("원본 문서 경로를 바꾸면 안 됨")

    SaveAs2 = SaveAs

    def Activate(self):
        self.rec("Activate")

    def Close(self, SaveChanges=None):
        self.rec("Close", SaveChanges)


# ---------------------------------------------------------------------------
# PowerPoint
# ---------------------------------------------------------------------------

class FakePresentation(Recorder):
    def __init__(self, full_name):
        super().__init__()
        self.FullName = full_name
        self.Name = full_name.replace("\\", "/").split("/")[-1]
        self.Saved = -1

    def Save(self):
        self.rec("Save")

    def Close(self):
        self.rec("Close")

    def SaveCopyAs(self, path):
        with open(path, "wb") as f:
            f.write(b"ppt-snapshot")


class FakePresentations:
    def __init__(self):
        self.open_args = None
        self.items = []

    @property
    def Count(self):
        return len(self.items)

    def __call__(self, i):
        return self.items[i - 1]

    def Open(self, path, *args):
        self.open_args = (path,) + args
        p = FakePresentation(path)
        self.items.append(p)
        return p


class FakePptApp:
    def __init__(self):
        object.__setattr__(self, "touched", [])
        self.DisplayAlerts = 2
        self.Presentations = FakePresentations()
        self.quit_called = False

    def __setattr__(self, key, value):
        if key in ("Visible", "WindowState"):
            self.touched.append(key)
        object.__setattr__(self, key, value)

    def Quit(self):
        self.quit_called = True


# ---------------------------------------------------------------------------
# win32com 대역
# ---------------------------------------------------------------------------

def make_win32com(dispatch_ex=None, dispatch=None, active=None):
    calls = []

    def _dispatch_ex(prog_id):
        calls.append(("DispatchEx", prog_id))
        return dispatch_ex

    def _dispatch(obj):
        calls.append(("Dispatch", obj if isinstance(obj, str) else type(obj).__name__))
        return dispatch if isinstance(obj, str) else obj

    def _get_active(prog_id):
        calls.append(("GetActiveObject", prog_id))
        if active is None:
            raise Exception("not running")
        return active

    client = types.SimpleNamespace(DispatchEx=_dispatch_ex, Dispatch=_dispatch, GetActiveObject=_get_active)
    return types.SimpleNamespace(client=client), calls
