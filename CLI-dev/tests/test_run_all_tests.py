# -*- coding: utf-8 -*-
"""run_all_tests.py: a silent hung child is stopped at the timeout, CLI-dev/tests is a pytest target, PyYAML is checked."""
import importlib.util
import os
import sys
import time

HERE = os.path.dirname(os.path.abspath(__file__))
spec = importlib.util.spec_from_file_location("run_all_tests_t", os.path.join(os.path.dirname(HERE), "run_all_tests.py"))
rat = importlib.util.module_from_spec(spec)
spec.loader.exec_module(rat)


def test_silent_hung_child_is_killed_at_timeout(capsys):
    t0 = time.time()
    code, out, sec, stop = rat.run_capture([sys.executable, "-c", "import time; print('start', flush=True); time.sleep(60)"],
                                           HERE, timeout=1, poll=0.2)
    assert time.time() - t0 < 20 and not stop and code is not None and code != 0
    assert out.splitlines()[0] == "start" and "시간 초과" in out


def test_normal_child_output_and_exit_code(capsys):
    code, out, sec, stop = rat.run_capture([sys.executable, "-c", "print('a'); print('b'); raise SystemExit(3)"], HERE, timeout=60)
    assert (code, out, stop) == (3, "a\nb", False)


def test_pytest_targets_and_packages():
    assert "tests" in rat.PYTEST_TARGETS and os.path.isdir(os.path.join(rat.HERE, "tests"))
    assert rat.IMPORTS["PyYAML"] == "yaml" and "PyYAML" in rat.PACKAGES
    assert set(rat.PACKAGES) == set(rat.IMPORTS)


def test_unencodable_line_on_cp949_stdout_does_not_stop_reader(monkeypatch):
    """With stdout redirected as cp949, a line with a character cp949 lacks must not stop the reader thread
    (the child would block on a full pipe until the timeout): it is shown with '?' and every line is kept."""
    import io
    buf = io.BytesIO()
    monkeypatch.setattr(sys, "stdout", io.TextIOWrapper(buf, encoding="cp949", errors="strict", line_buffering=True))
    child = "import sys\nprint('a \\u2022 b')\nfor i in range(20000): print('줄', i)\nprint('끝')\n"
    t0 = time.time()
    code, out, sec, stop = rat.run_capture([sys.executable, "-c", child], HERE, timeout=60, poll=0.2)
    sys.stdout.flush()
    shown = buf.getvalue().decode("cp949")
    assert code == 0 and not stop and time.time() - t0 < 30
    lines = out.splitlines()
    assert lines[0] == "a \u2022 b" and lines[-1] == "끝" and len(lines) == 20002
    assert shown.splitlines()[0] == "a ? b" and "끝" in shown


def test_echo_and_say_do_not_raise_on_cp949_stdout(monkeypatch):
    """The diag text and the closing summary (printed inside finally) go through echo(): a character cp949 lacks
    (e.g. in a file name) becomes '?' instead of raising UnicodeEncodeError."""
    import io
    buf = io.BytesIO()
    monkeypatch.setattr(sys, "stdout", io.TextIOWrapper(buf, encoding="cp949", errors="strict", line_buffering=True))
    rat.echo("요약\nppt: a \u2022 b")
    rat.say("끝. C:\\temp\\r\u00e9sum\u00e9\u2022.txt")
    sys.stdout.flush()
    shown = buf.getvalue().decode("cp949")
    assert "ppt: a ? b" in shown and "[STATUS] 끝." in shown
