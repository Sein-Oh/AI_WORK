# -*- coding: utf-8 -*-
"""tools/run_win_tests.py needs Windows/Office, but its AI-readable log writer must keep working.
Windows-only modules are stubbed; one real (failing) tool call is recorded and the log is checked."""
import importlib.util
import os
import sys
import types

HERE = os.path.dirname(os.path.abspath(__file__))
DEV = os.path.dirname(HERE)  # CLI-dev\ml-kit (tests, tools)
KIT = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(HERE))), "CLI", "ml-kit")  # 실행 파일은 CLI\ml-kit


def _load_runner():
    for m in ("pythoncom", "pywintypes", "win32com", "win32com.client"):
        sys.modules.setdefault(m, types.ModuleType(m))
    sys.modules["win32com"].client = sys.modules["win32com.client"]
    spec = importlib.util.spec_from_file_location("run_win_tests", os.path.join(DEV, "tools", "run_win_tests.py"))
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def test_testlog_contains_criteria_calls_and_traceback(tmp_path):
    rwt = _load_runner()
    ctx = rwt.Ctx(type("Args", (), {"dir": str(tmp_path)})())

    def m13():
        d, c, o, e = rwt.run_tool("ml_train.py", "--없는옵션")
        rwt.record(ctx, "M13", "잘못된 옵션", False, f"exit={c}", d, c, e)

    def m6():
        raise RuntimeError("Word COM 실패 (테스트)")

    rwt.run_test(ctx, "M13", "오류 계약", m13)
    rwt.run_test(ctx, "M6", "보고서를 Word로", m6)
    path = rwt.write_testlog(ctx, str(tmp_path / "logs"), "테스트PC", "20260101_000000", fatal=None)
    text = open(path, encoding="utf-8").read()
    assert os.path.basename(path) == "testlog_테스트PC_20260101_000000.txt"
    assert "결과: ERROR=1, FAIL=1" in text
    assert rwt.EXPECT["M13"] in text and rwt.EXPECT["M6"] in text          # pass criteria
    assert "$ python ml_train.py --없는옵션" in text and "exit=2" in text  # recorded call
    assert '"status": "error"' in text                                     # full stdout JSON
    assert "RuntimeError: Word COM 실패" in text and "Traceback" in text   # exception trace
    assert "[환경]" in text and "scikit-learn" in text and "[전체 호출 기록]" in text


def test_only_rejects_unknown_ids(monkeypatch, capsys):
    import pytest
    rwt = _load_runner()
    assert rwt.TEST_IDS == tuple(f"M{i}" for i in range(1, 15))
    assert rwt.parse_only(None) == (None, [])
    assert rwt.parse_only("m1, M4") == ({"M1", "M4"}, [])
    assert rwt.parse_only("M1,M15,X") == (None, ["X", "M15"])
    assert rwt.parse_only(",") == (None, ["(비어 있음)"])
    monkeypatch.setattr(sys, "argv", ["run_win_tests.py", "--only", "M1,M99"])
    with pytest.raises(SystemExit) as ex:
        rwt.main()
    assert ex.value.code == 2 and "M99" in capsys.readouterr().err


def test_decode_console_reads_locale_text():
    import locale
    rwt = _load_runner()
    assert rwt.decode_console(b"Active code page: 65001") == "Active code page: 65001"
    text = "활성 코드 페이지: 949"
    try:                                     # the console code page of this PC (Korean Windows: cp949)
        raw = text.encode("oem" if os.name == "nt" else locale.getpreferredencoding(False))
    except (LookupError, UnicodeEncodeError):
        return
    assert rwt.decode_console(raw) == text
