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
