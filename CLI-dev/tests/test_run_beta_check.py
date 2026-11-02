# -*- coding: utf-8 -*-
"""run_beta_check.py --agent: one stdout JSON line for help, argument errors, interrupts, runner errors and normal ends."""
import importlib.util
import json
import os
import subprocess
import sys

import pytest

HERE = os.path.dirname(os.path.abspath(__file__))
SCRIPT = os.path.join(os.path.dirname(HERE), "run_beta_check.py")


@pytest.fixture
def rb(monkeypatch, tmp_path):
    spec = importlib.util.spec_from_file_location("run_beta_check_t", SCRIPT)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    monkeypatch.setattr(mod, "HERE", str(tmp_path))
    mod.RESULTS.clear()
    return mod


def _cli(*args):
    r = subprocess.run([sys.executable, SCRIPT, *args], capture_output=True, timeout=60,
                       env=dict(os.environ, PYTHONIOENCODING="cp949"))
    return r.returncode, r.stdout.decode("utf-8").splitlines()


def test_agent_help_and_argument_errors_are_one_json_line():
    code, out = _cli("--agent", "-h")
    assert code == 0 and len(out) == 1 and json.loads(out[0])["command"] == "help"
    for bad in (["--agent", "--only", "B,Z"], ["--agent", "--ag"], ["--agent", "--nope"]):
        code, out = _cli(*bad)
        payload = json.loads(out[0])
        assert code == 2 and len(out) == 1 and payload["status"] == "error" and payload["agent_directive"].startswith("STOP_RETRY")


def test_unknown_group_rejected_without_agent():
    r = subprocess.run([sys.executable, SCRIPT, "--only", "Q"], capture_output=True, timeout=60)
    assert r.returncode == 2 and b"Q" in r.stderr


def _run(rb, capsys, tmp_path, *extra):
    code = rb.main(["--agent", "--dir", str(tmp_path / "work"), "--label", "a/b:c", *extra])
    out = capsys.readouterr().out.splitlines()
    assert len(out) == 1
    return code, json.loads(out[0])


def test_interrupt_exits_130_with_json(rb, capsys, tmp_path, monkeypatch):
    def stop(work, groups):
        raise KeyboardInterrupt
    monkeypatch.setattr(rb, "run_groups", stop)
    code, payload = _run(rb, capsys, tmp_path)
    assert code == 130 and payload["error_type"] == "interrupted" and payload["agent_directive"].startswith("STOP_RETRY")


def test_runner_error_and_finish_report(rb, capsys, tmp_path, monkeypatch):
    def boom(work, groups):
        rb.check("H", "H-x", "ok item", True)
        raise RuntimeError("runner bug")
    monkeypatch.setattr(rb, "run_groups", boom)
    code, payload = _run(rb, capsys, tmp_path)
    assert code == 0 and payload["status"] == "ok" and payload["passed"] == 1 and payload["failed"] == 1
    assert payload["verified"] is True and payload["agent_directive"].startswith("VERIFIED")
    assert os.path.basename(payload["report_file"]).startswith("beta_check_abc_")
    assert "RUNNER" in payload["failed_items"][0]


def test_unwritable_report_is_unverified(rb, capsys, tmp_path, monkeypatch):
    monkeypatch.setattr(rb, "run_groups", lambda work, groups: None)
    monkeypatch.setattr(rb, "write_report", lambda *a: (_ for _ in ()).throw(OSError("disk full")))
    code, payload = _run(rb, capsys, tmp_path)
    assert code == 0 and payload["verified"] is False and payload["agent_directive"].startswith("UNVERIFIED")


def test_sample_failure_and_tool_timeout_are_recorded(rb, capsys, tmp_path, monkeypatch):
    import make_beta_samples
    monkeypatch.setattr(make_beta_samples, "main", lambda work: (_ for _ in ()).throw(PermissionError("locked")))
    code, payload = _run(rb, capsys, tmp_path, "--only", "B")
    assert code == 0 and payload["failed"] == 1 and "SAMPLES" in payload["failed_items"][0]

    def timeout(*a, **kw):
        raise subprocess.TimeoutExpired("ml_report.py", 3600)
    monkeypatch.setattr(rb.subprocess, "run", timeout)
    assert rb.run("ml_report.py", cwd=str(tmp_path))[0] == "TIMEOUT"
    rb.RESULTS.clear()
    rb.group_f(str(tmp_path), {})
    assert [r["status"] for r in rb.RESULTS] == ["FAIL"]
