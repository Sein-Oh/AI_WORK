"""실기 테스트 러너(tools/run_win_tests.py)의 판단 로직 (Windows·pywin32 없이 가짜 모듈로).

- 문서가 남아 있는 프로세스는 Quit도 WM_CLOSE도 하지 않음
- A5는 자기가 띄운 PID만 끝냄
- 닫힌 테스트 문서는 다시 열고(JIT), 못 열면 항목을 바로 ERROR로 끝냄 (com_retry가 오래 돌지 않음)
- 에이전트 모드(--agent)는 ACI 계약을 지키고, 인자를 pywin32 확인보다 먼저 해석함
"""
import datetime
import importlib.util
import json
import os
import sys
import time
import types

import pytest

from office_helpers import TOOL_DIR, TOOLS


if TOOLS not in sys.path:
    sys.path.insert(0, TOOLS)

import win_test_rules as rules  # noqa: E402


class com_error(Exception):
    """pywintypes.com_error 흉내."""


BADINDEX = com_error(-2147352567, "예외가 발생했습니다.", (0, None, None, None, 0, -2147352565), None)


BUSY = com_error(-2147418111, "호출이 거부되었습니다.", None, None)          # RPC_E_CALL_REJECTED


NOT_RUNNING = com_error(-2147221021, "작업을 사용할 수 없습니다.", None, None)  # MK_E_UNAVAILABLE


T0 = datetime.datetime(2026, 10, 28, 13, 0, tzinfo=datetime.timezone.utc)


BEFORE, AFTER = T0 - datetime.timedelta(hours=1), T0 + datetime.timedelta(minutes=5)


def W(hwnd, pid, title, visible=True, cls="XLMAIN"):
    return {"hwnd": hwnd, "pid": pid, "cls": cls, "title": title, "visible": visible}


class Doc:
    def __init__(self, coll, name, path):
        self.coll, self.Name, self.Path = coll, name, path

    def Close(self, *a):
        self.coll.items.remove(self)


class Coll:
    def __init__(self):
        self.items = []

    @property
    def Count(self):
        return len(self.items)

    def __call__(self, i):
        return self.items[i - 1]


class App:
    def __init__(self, pid, kind, docs):
        self.pid, self.kind, self.quit = pid, kind, False
        self.coll = Coll()
        self.coll.items = [Doc(self.coll, n, p) for n, p in docs]
        self.Hwnd = self.HWND = pid * 10
        self.ActiveWindow = types.SimpleNamespace(Hwnd=pid * 10)
        setattr(self, {"excel": "Workbooks", "word": "Documents", "ppt": "Presentations"}[kind], self.coll)

    def Quit(self):
        self.quit = True


class Desk:
    """가짜 데스크톱: 프로세스, 창, 보낸 WM_CLOSE."""

    def __init__(self, apps, extra_windows, starts):
        self.apps, self.extra, self.starts, self.posted = apps, extra_windows, starts, []

    def windows(self):
        found = []
        for a in self.apps:
            if a.kind == "word" or a.quit:          # Quit으로 끝난 앱은 창이 사라짐 (Word 창은 정리 대상 아님)
                continue
            names = [d.Name for d in a.coll.items]
            title = (f"{names[0]} - Excel" if a.kind == "excel" else f"{names[0]} - PowerPoint") if names else \
                ("Excel" if a.kind == "excel" else "PowerPoint")
            found.append(W(a.pid * 10, a.pid, title, cls="XLMAIN" if a.kind == "excel" else "PPTFrameClass"))
        return found + self.extra

    def modules(self):
        desk = self
        office = types.ModuleType("_office")
        office.FRAME_CLASS = {"excel": ("XLMAIN",), "word": ("OpusApp",), "ppt": ("PPTFrameClass",)}
        office.find_apps = lambda kind: [a for a in desk.apps if a.kind == kind]
        info = {}

        def enum(cb, _):
            info.clear()
            for w in desk.windows():
                info[w["hwnd"]] = w
                cb(w["hwnd"], None)

        def window(h):
            return info.get(h) or next(w for w in desk.windows() if w["hwnd"] == h)
        gui = types.SimpleNamespace(
            EnumWindows=enum, GetClassName=lambda h: window(h)["cls"], GetWindowText=lambda h: window(h)["title"],
            IsWindowVisible=lambda h: window(h)["visible"], PostMessage=lambda h, *a: desk.posted.append(h))
        proc = types.SimpleNamespace(
            GetWindowThreadProcessId=lambda h: (0, h // 10 if h % 10 == 0 else window(h)["pid"]),
            GetProcessTimes=lambda h: {"CreationTime": desk.starts[h]})
        api = types.SimpleNamespace(OpenProcess=lambda access, inherit, pid: pid, CloseHandle=lambda h: None)
        con = types.SimpleNamespace(WM_CLOSE=0x10)
        return {"_office": office, "win32gui": gui, "win32process": proc, "win32api": api, "win32con": con}


def run_cleanup(monkeypatch, capsys, desk, test_dir, only_name="", allowed=None, protect=()):
    src = open(os.path.join(TOOLS, "run_win_tests.py"), encoding="utf-8").read()
    script = src.split('CLEANUP_SCRIPT = r"""', 1)[1].split('"""', 1)[0]
    for name, mod in desk.modules().items():
        monkeypatch.setitem(sys.modules, name, mod)
    monkeypatch.setattr(time, "sleep", lambda s: None)
    opts = {"test_dir": test_dir, "keep_untitled": [], "only_name": only_name, "run_start": T0.timestamp(),
            "allowed_pids": allowed, "protect_names": list(protect)}
    monkeypatch.setattr(sys, "argv", ["-c", TOOL_DIR, json.dumps(opts), TOOLS])
    exec(compile(script, "CLEANUP_SCRIPT", "exec"), {"__name__": "__main__"})
    return json.loads(capsys.readouterr().out.strip().splitlines()[-1])


TEST_DIR = os.path.normcase(os.path.join("C:", os.sep, "AIKIT", "dev", "work", "office"))


def a5_desk(hidden_user_frame=True):
    user_xl = App(10, "excel", [("A.xlsx", TEST_DIR)])
    second = App(20, "excel", [("C_second_instance.xlsx", TEST_DIR)])
    second.Quit = lambda: None                    # excel.exe /x로 띄운 인스턴스는 Quit을 받고도 남음
    word = App(30, "word", [("A.docx", TEST_DIR)])
    ppt = App(40, "ppt", [("A.pptx", TEST_DIR)])
    extra = [W(101, 10, "", visible=False)] if hidden_user_frame else []    # Office 16의 숨은 XLMAIN
    return Desk([user_xl, second, word, ppt], extra, {10: AFTER, 20: AFTER, 30: AFTER, 40: AFTER})


class Clock:
    def __init__(self):
        self.now = 1000.0

    def time(self):
        return self.now

    def sleep(self, s):
        self.now += s


@pytest.fixture
def runner(monkeypatch, tmp_path):
    win32com = types.ModuleType("win32com")
    client = types.ModuleType("win32com.client")
    client.GetActiveObject = lambda prog: (_ for _ in ()).throw(NOT_RUNNING)
    win32com.client = client
    monkeypatch.setitem(sys.modules, "pythoncom", types.ModuleType("pythoncom"))
    monkeypatch.setitem(sys.modules, "win32com", win32com)
    monkeypatch.setitem(sys.modules, "win32com.client", client)
    spec = importlib.util.spec_from_file_location("run_win_tests_t", os.path.join(TOOLS, "run_win_tests.py"))
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    clock = Clock()
    monkeypatch.setattr(mod, "time", clock)
    args = types.SimpleNamespace(dir=str(tmp_path), auto=True, only=None, agent=False, skip_outlook=True,
                                 self_email=None, keep_open=False)
    mod.CTX = mod.Ctx(args)
    mod.CTX.docs_ready = True
    mod.LOG_LINES.clear()
    mod.CALLS.clear()
    return mod


def _agent_main(runner, monkeypatch, capsys, tmp_path, argv, groups=None, **patch):
    """Runs runner.main() in agent mode with every Windows step stubbed; returns (exit code, stdout lines)."""
    monkeypatch.setattr(runner, "AGENT", True)
    monkeypatch.setattr(sys, "argv", ["run_win_tests.py", "--agent", "--dir", str(tmp_path)] + argv)
    pc = runner.pythoncom
    monkeypatch.setattr(pc, "CoInitialize", lambda: None, raising=False)
    monkeypatch.setattr(pc, "CoUninitialize", lambda: None, raising=False)
    stubs = {"process_count": lambda img: 0, "untitled_documents": lambda: [], "setup_files": lambda ctx: None,
             "open_user_windows": lambda ctx: None, "ensure_test_docs": lambda ctx: None,
             "save_shots": lambda ctx: (str(tmp_path / "shots"), 0), "close_test_documents": lambda ctx: [],
             "environment_lines": lambda ctx: ["(환경)"]}
    stubs.update(patch)
    for name, fn in stubs.items():
        monkeypatch.setattr(runner, name, fn)
    if groups is not None:
        monkeypatch.setattr(runner, "GROUPS", groups)
    runner.UNVERIFIED_SEEN.clear()
    with pytest.raises(SystemExit) as ex:
        runner.main()
    return ex.value.code, capsys.readouterr().out.splitlines()


def _passing_group(ctx):
    ctx.results.append({"id": "A1", "title": "t", "status": "PASS"})


def _runner_without_pywin32(monkeypatch, argv):
    monkeypatch.setitem(sys.modules, "pythoncom", None)          # import pythoncom -> ImportError
    monkeypatch.setitem(sys.modules, "win32com", None)
    monkeypatch.setitem(sys.modules, "win32com.client", None)
    monkeypatch.setattr(sys, "argv", ["run_win_tests.py"] + argv)
    if TOOLS not in sys.path:
        monkeypatch.syspath_prepend(TOOLS)
    spec = importlib.util.spec_from_file_location("run_win_tests_l3", os.path.join(TOOLS, "run_win_tests.py"))
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


# ---------------------------------------------------------------------------
# win_test_rules (순수 규칙)
# ---------------------------------------------------------------------------


def test_missing_doc_error_detection():
    assert rules.is_missing_doc_error(BADINDEX)
    assert rules.is_missing_doc_error(com_error(-2147352565, "x", None, None))
    assert not rules.is_missing_doc_error(NOT_RUNNING)        # 앱 시작 직후 ROT 등록 지연일 수 있어 재시도 대상
    assert rules.is_doc_gone_error(NOT_RUNNING) and rules.is_doc_gone_error(BADINDEX)
    assert not rules.is_missing_doc_error(BUSY) and not rules.is_doc_gone_error(BUSY)
    assert not rules.is_missing_doc_error(ValueError("x"))


def test_parse_tasklist_pids():
    out = ('"EXCEL.EXE","1234","Console","1","123,456 K"\r\n'
           '"EXCEL.EXE","5678","Console","1","9,000 K"\r\n')
    assert rules.parse_tasklist_pids(out, "EXCEL.EXE") == [1234, 5678]
    assert rules.parse_tasklist_pids("정보: 지정된 조건에 맞는 작업이 실행되고 있지 않습니다.", "EXCEL.EXE") == []


def test_title_has_doc():
    assert rules.title_has_doc("A.xlsx - Excel", "A.xlsx")
    assert rules.title_has_doc("C:\\t\\A.xlsx [읽기 전용] - Excel", "A.xlsx")
    assert not rules.title_has_doc("Data.xlsx - Excel", "A.xlsx")
    assert not rules.title_has_doc("Excel", "A.xlsx")


def test_hidden_frame_of_document_process_is_never_closed():
    """회사 PC 재현: 사용자 Excel(A.xlsx, 테스트 중 시작됨)에 숨은 XLMAIN 창이 있음."""
    windows = [W(1, 100, "A.xlsx - Excel"), W(2, 100, "", visible=False), W(3, 200, "Excel")]
    starts = {100: AFTER, 200: AFTER}
    protected = rules.protected_pids(windows)
    assert protected == {100}
    targets, skipped = rules.select_close_targets(windows, protected, None, T0, starts)
    assert [w["hwnd"] for w in targets] == [3]
    assert [(w["hwnd"], why) for w, why in skipped] == [(2, "문서가 열려 있는 프로세스 (보호)")]


def test_allowed_pids_limit_targets():
    windows = [W(2, 100, "", visible=False), W(3, 200, "Excel"), W(4, 300, "", visible=False)]
    targets, skipped = rules.select_close_targets(windows, set(), {200}, T0, {100: AFTER, 200: AFTER, 300: AFTER})
    assert [w["pid"] for w in targets] == [200]
    assert {w["pid"] for w, _ in skipped} == {100, 300}


def test_pre_existing_or_unknown_start_is_kept():
    windows = [W(2, 100, "Excel"), W(3, 200, "Excel")]
    targets, _ = rules.select_close_targets(windows, set(), None, T0, {100: BEFORE})
    assert targets == []


def test_doc_pids_from_com_are_protected():
    windows = [W(2, 100, "", visible=False)]
    protected = rules.protected_pids(windows, doc_pids=[100])
    assert rules.select_close_targets(windows, protected, None, T0, {100: AFTER})[0] == []
    assert rules.protected_pids([W(5, 400, "a.xlsx", visible=False)], protect_names=["A.xlsx"]) == {400}


def test_may_quit_app():
    assert rules.may_quit_app(100, 1, set())[0] is False
    assert rules.may_quit_app(None, 0, set(), allowed_pids={200})[0] is False
    assert rules.may_quit_app(None, 0, set())[0] is True
    assert rules.may_quit_app(200, 0, set(), {200})[0] is True
    assert rules.may_quit_app(100, 0, {100}, None, T0, AFTER)[0] is False
    assert rules.may_quit_app(100, 0, set(), None, T0, BEFORE)[0] is False
    assert rules.may_quit_app(100, 0, set(), None, T0, AFTER)[0] is True


# ---------------------------------------------------------------------------
# 정리 스크립트(CLEANUP_SCRIPT)를 가짜 Office/pywin32로 실행
# ---------------------------------------------------------------------------


def test_a5_cleanup_keeps_user_excel_with_a_xlsx(monkeypatch, capsys):
    desk = a5_desk()
    user_xl, second = desk.apps[0], desk.apps[1]
    out = run_cleanup(monkeypatch, capsys, desk, TEST_DIR, only_name="C_second_instance.xlsx", allowed=[20],
                      protect=["A.xlsx", "A.docx", "A.pptx"])
    assert out["closed"][0] == "C_second_instance.xlsx"
    assert [d.Name for d in user_xl.coll.items] == ["A.xlsx"] and not user_xl.quit
    assert 100 not in desk.posted and 101 not in desk.posted          # 사용자 Excel 창에는 WM_CLOSE 없음
    assert desk.posted == [200]                                       # 두 번째 Excel(Quit 후에도 남은 빈 창)만
    assert any("보호" in a for a in out["actions"])


def test_cleanup_without_allowed_pids_still_protects_document_process(monkeypatch, capsys):
    """allowed_pids가 없어도(예전 호출 방식) 문서가 있는 프로세스의 숨은 창은 닫지 않음."""
    desk = a5_desk()
    run_cleanup(monkeypatch, capsys, desk, TEST_DIR, only_name="C_second_instance.xlsx")
    assert 101 not in desk.posted and 100 not in desk.posted
    assert not desk.apps[0].quit


def test_end_of_run_cleanup(monkeypatch, capsys):
    runner_xl = App(10, "excel", [("A.xlsx", TEST_DIR), ("B.xlsx", TEST_DIR)])
    user_xl = App(50, "excel", [])                                     # 실행 전부터 있던 빈 Excel
    mine = App(60, "excel", [("내 파일.xlsx", "D:\\work")])            # 실행 중 사용자가 연 문서
    word = App(30, "word", [("A.docx", TEST_DIR)])
    desk = Desk([runner_xl, user_xl, mine, word], [], {10: AFTER, 50: BEFORE, 60: AFTER, 30: AFTER})
    out = run_cleanup(monkeypatch, capsys, desk, TEST_DIR)
    assert set(out["closed"]) >= {"A.xlsx", "B.xlsx", "A.docx"}
    assert runner_xl.quit and word.quit
    assert not user_xl.quit and 500 not in desk.posted
    assert not mine.quit and 600 not in desk.posted and mine.coll.Count == 1


# ---------------------------------------------------------------------------
# 러너 모듈 (가짜 pythoncom/win32com으로 import)
# ---------------------------------------------------------------------------


def test_com_retry_fails_fast_on_missing_doc(runner):
    calls = []

    def missing():
        calls.append(1)
        raise BADINDEX
    with pytest.raises(com_error):
        runner.com_retry(missing)
    assert len(calls) == 1

    busy = []

    def flaky():
        busy.append(1)
        if len(busy) < 3:
            raise BUSY
        return "ok"
    assert runner.com_retry(flaky) == "ok" and len(busy) == 3


def test_closed_test_doc_is_reopened_once(runner, monkeypatch):
    state = {"open": False, "lookups": 0}

    def lookup(kind, name):
        state["lookups"] += 1
        if not state["open"]:
            raise BADINDEX
        return f"<{name}>"
    monkeypatch.setattr(runner, "_lookup_doc", lookup)
    opened = []
    monkeypatch.setattr(runner.os, "startfile", lambda p: (opened.append(p), state.update(open=True)), raising=False)
    assert runner.xl_wb("A.xlsx") == "<A.xlsx>"
    assert opened == [runner.CTX.path("A.xlsx")]
    assert any("A.xlsx가 닫혀 있어 다시 열었습니다" in r for r in runner.CTX.recoveries)
    assert any(line.startswith("[WARN] A.xlsx가 닫혀 있어 다시 열었습니다") for line in runner.LOG_LINES)


def test_unreopenable_doc_fails_items_fast_with_one_reason(runner, monkeypatch):
    lookups = []

    def lookup(kind, name):
        lookups.append(name)
        raise BADINDEX
    monkeypatch.setattr(runner, "_lookup_doc", lookup)
    monkeypatch.setattr(runner.os, "startfile", lambda p: None, raising=False)
    ctx = runner.CTX
    runner.run_test(ctx, "B2", "활성 시트 유지", lambda: runner.activate_xl("A.xlsx", "Sheet1"))
    n = len(lookups)
    assert n < 100                                       # 예전: 중첩 com_retry로 수백 번 (사용자가 Ctrl+C)
    runner.run_test(ctx, "C1", "Excel undo", lambda: runner.a1_value())
    assert len(lookups) == n                             # 이번 그룹에서는 다시 시도하지 않고 바로 ERROR
    assert [r["status"] for r in ctx.results] == ["ERROR", "ERROR"]
    assert all(r["detail"] == "테스트 문서 A.xlsx가 닫혀 다시 열지 못함" for r in ctx.results)
    assert "A.xlsx" in ctx.unrecoverable
    runner.ensure_test_docs(ctx)                         # 다음 그룹 시작 때는 다시 시도
    assert len(lookups) > n


def test_wait_for_does_not_reopen(runner, monkeypatch):
    monkeypatch.setattr(runner, "_lookup_doc", lambda kind, name: (_ for _ in ()).throw(BADINDEX))
    opened = []
    monkeypatch.setattr(runner.os, "startfile", lambda p: opened.append(p), raising=False)
    assert runner.wait_for(lambda: runner.xl_wb("A.xlsx"), timeout=3) is None
    assert opened == []                                  # undo가 다시 여는 중일 수 있어 wait_for 안에서는 열지 않음


def test_non_test_doc_is_not_reopened(runner, monkeypatch):
    monkeypatch.setattr(runner, "_lookup_doc", lambda kind, name: (_ for _ in ()).throw(BADINDEX))
    monkeypatch.setattr(runner.os, "startfile", lambda p: pytest.fail("열면 안 됨"), raising=False)
    with pytest.raises(com_error):
        runner.xl_wb("C_second_instance.xlsx")


def test_close_test_documents_passes_allowed_pids_and_logs_actions(runner, monkeypatch):
    seen = {}

    def fake_run(args, **kw):
        seen["opts"] = json.loads(args[4])
        out = json.dumps({"closed": ["C_second_instance.xlsx"], "actions": ["WM_CLOSE XLMAIN pid 20 hwnd 200"]})
        return types.SimpleNamespace(returncode=0, stdout=out.encode("utf-8"), stderr=b"")
    monkeypatch.setattr(runner.subprocess, "run", fake_run)
    closed = runner.close_test_documents(runner.CTX, only_name="C_second_instance.xlsx", allowed_pids={20}, label="A5")
    assert closed == ["C_second_instance.xlsx"]
    assert seen["opts"]["allowed_pids"] == [20] and "A.xlsx" in seen["opts"]["protect_names"]
    assert any("[A5]" in c and "WM_CLOSE XLMAIN pid 20" in c for c in runner.CTX.cleanup_log)


def test_error_log_has_recoveries_cleanup_and_tool_log(runner, monkeypatch, tmp_path):
    data = tmp_path / "_data"
    (data / "logs").mkdir(parents=True)
    log_file = data / "logs" / "errors.txt"
    log_file.write_text("===== 이전 실행\n\n", encoding="utf-8")
    monkeypatch.setattr(runner, "RUN_ENV", {"OFFICE_CLI_DATA_DIR": str(data)})
    runner.TOOL_LOG_OFFSET[0] = log_file.stat().st_size
    with open(log_file, "a", encoding="utf-8") as f:
        f.write("===== 2026-10-28T13:49:53 | office-cli 1.2.1 | Python 3.13.13 | Windows\n명령: excel.py write\n\n")
    monkeypatch.setattr(runner, "environment_lines", lambda ctx: ["(환경)"])
    ctx = runner.CTX
    ctx.recoveries.append("13:10:00 A.xlsx가 닫혀 있어 다시 열었습니다")
    ctx.cleanup_log.append("13:09:00 [A5]   WM_CLOSE XLMAIN pid 20 hwnd 200")
    text = open(runner.write_error_log(ctx), encoding="utf-8").read()
    assert "러너 자동 복구와 창/프로세스 정리" in text and "A.xlsx가 닫혀 있어 다시 열었습니다" in text
    assert "WM_CLOSE XLMAIN pid 20" in text
    assert "도구 오류 로그" in text and "명령: excel.py write" in text
    assert "이전 실행" not in text


# ---------------------------------------------------------------------------
# 에이전트 모드 계약 (--agent): stdout JSON 1줄, exit 0/1/2/130, verified, unverified_count
# ---------------------------------------------------------------------------


def test_agent_help_and_bad_args_are_json(runner, monkeypatch, capsys, tmp_path):
    code, out = _agent_main(runner, monkeypatch, capsys, tmp_path, ["-h"])
    assert code == 0 and len(out) == 1 and json.loads(out[0])["command"] == "help" and json.loads(out[0])["verified"] is True
    for bad in (["--only", "A,Z"], ["--nope"], ["--skip-out"]):                 # unknown group, unknown option, abbreviation
        code, out = _agent_main(runner, monkeypatch, capsys, tmp_path, bad)
        payload = json.loads(out[0])
        assert code == 2 and len(out) == 1 and payload["status"] == "error" and payload["agent_directive"].startswith("STOP_RETRY")


def test_agent_summary_reports_unverified_count(runner, monkeypatch, capsys, tmp_path):
    def group(ctx):
        _passing_group(ctx)
        runner.UNVERIFIED_SEEN.append({"command": "python excel.py write", "unverified": ["A1 값"]})
    code, out = _agent_main(runner, monkeypatch, capsys, tmp_path, ["--only", "A"], groups=[("A", group)])
    payload = json.loads(out[0])
    assert code == 0 and len(out) == 1 and payload["unverified_count"] == 1 and payload["failed"] == []
    assert payload["agent_directive"].startswith("UNVERIFIED:") and "통과'로 보고하지" in payload["agent_directive"]
    assert payload["verified"] is False and os.path.isfile(payload["report_file"])


def test_agent_summary_clean_run_is_verified(runner, monkeypatch, capsys, tmp_path):
    code, out = _agent_main(runner, monkeypatch, capsys, tmp_path, ["--only", "a"], groups=[("A", _passing_group)])
    payload = json.loads(out[0])
    assert code == 0 and payload["unverified_count"] == 0 and payload["verified"] is True
    assert payload["agent_directive"].startswith("VERIFIED:") and payload["summary"] == {"PASS": 1}


def test_agent_finish_errors_and_interrupts_keep_contract(runner, monkeypatch, capsys, tmp_path):
    def broken(ctx):
        raise RuntimeError("PowerPoint 그림 저장 실패")
    code, out = _agent_main(runner, monkeypatch, capsys, tmp_path, [], groups=[("A", _passing_group)], save_shots=broken)
    payload = json.loads(out[0])
    assert code == 1 and len(out) == 1 and payload["error_type"] == "finish" and payload["agent_directive"].startswith("ERROR:")

    def stop(ctx):
        raise KeyboardInterrupt
    code, out = _agent_main(runner, monkeypatch, capsys, tmp_path, [], groups=[("A", _passing_group)], close_test_documents=stop)
    assert code == 130 and len(out) == 1 and json.loads(out[0])["error_type"] == "interrupted"
    code, out = _agent_main(runner, monkeypatch, capsys, tmp_path, [], groups=[("A", stop)])
    assert code == 130 and len(out) == 1 and json.loads(out[0])["error_type"] == "interrupted"


def test_unknown_group_rejected_without_agent(runner, monkeypatch, capsys):
    monkeypatch.setattr(runner, "AGENT", False)
    monkeypatch.setattr(sys, "argv", ["run_win_tests.py", "--only", "Z"])
    with pytest.raises(SystemExit) as ex:
        runner.main()
    assert ex.value.code == 2 and "Z" in capsys.readouterr().err


# ---------------------------------------------------------------------------
# run_win_tests.py는 인자를 먼저 해석
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("argv,code,status", [(["--agent", "-h"], 0, "ok"), (["--agent", "--nope"], 2, "error"),
                                              (["--agent", "--only", "Z"], 2, "error"), (["--agent"], 1, "error")])
def test_runner_parses_args_before_pywin32_check(monkeypatch, capsys, argv, code, status):
    mod = _runner_without_pywin32(monkeypatch, argv)
    with pytest.raises(SystemExit) as ei:
        mod.main()
    out = [ln for ln in capsys.readouterr().out.splitlines() if ln.strip()]
    assert ei.value.code == code and len(out) == 1, out
    payload = json.loads(out[0])
    assert payload["status"] == status
    if code == 1:
        assert payload["error_type"] == "dependency" and "python -m pip install pywin32" in payload["agent_directive"]
