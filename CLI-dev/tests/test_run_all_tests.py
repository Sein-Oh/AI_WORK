# -*- coding: utf-8 -*-
"""run_all_tests.py: a silent hung child is stopped at the timeout, CLI-dev/tests is a pytest target, PyYAML is checked,
stages run in a fixed order into one result file with the summary on top and the tool error logs at the end."""
import importlib.util
import os
import sys
import time

import pytest

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


# ---------------------------------------------------------------------------
# One command, one result file: stage order, summary table, errors.txt at the end (fakes for Windows-only stages)
# ---------------------------------------------------------------------------
@pytest.fixture
def iso(monkeypatch, tmp_path):
    """Result file in tmp_path, tool folders (error logs) in tmp_path/CLI, no data-dir overrides."""
    cli = tmp_path / "CLI"
    for tool in rat.TOOLS:
        (cli / tool).mkdir(parents=True)
    monkeypatch.setattr(rat, "HERE", str(tmp_path))
    monkeypatch.setattr(rat, "CLI_DIR", str(cli))
    monkeypatch.delenv("ML_KIT_DATA_DIR", raising=False)
    monkeypatch.delenv("OFFICE_CLI_DATA_DIR", raising=False)
    return tmp_path


def _fake_runners(monkeypatch, order, behaviour=None):
    behaviour = behaviour or {}

    def make(key):
        def run(stage, args, ctx):
            order.append(key)
            if key in behaviour:
                behaviour[key](stage)
            else:
                stage.passed += 1
        return run
    monkeypatch.setattr(rat, "RUNNERS", {k: make(k) for k in rat.STAGES})


def _result_files(folder):
    return sorted(p for p in os.listdir(folder) if p.startswith("result_") and p.endswith(".md"))


def test_default_stages_run_in_order_into_one_file(iso, monkeypatch, capsys):
    order = []
    _fake_runners(monkeypatch, order)
    assert rat.main(["--label", "회사PC"]) == 0
    assert order == ["diag", "pytest", "beta", "office"] == list(rat.DEFAULT_STAGES)
    files = _result_files(iso)
    assert len(files) == 1 and files[0].startswith("result_회사PC_")
    assert [p for p in os.listdir(iso) if os.path.isfile(iso / p)] == files     # nothing else written next to the runner


def test_only_keeps_fixed_order_and_rejects_unknown(iso, monkeypatch, capsys):
    order = []
    _fake_runners(monkeypatch, order)
    rat.main(["--only", "office, ML,diag", "--label", "a"])
    assert order == ["diag", "ml", "office"]
    with pytest.raises(SystemExit) as e:
        rat.main(["--only", "diag,nope"])
    assert e.value.code == 2
    with pytest.raises(SystemExit):
        rat.main(["--auto"])                                    # replaced by --ask (office questions are off by default)


def _fail_diag(stage):
    stage.passed += 3
    stage.fail("python-docx 기본 템플릿: zip 형식=False")


def _skip_office(stage):
    stage.env_skip("Windows가 아님")


def _beta(stage):
    stage.passed += 26
    stage.skipped += 1
    stage.skips.append("F7 엑셀로 수식 재계산 확인")
    stage.unverified.append("excel.py read -> 수식 값 확인하지 못함")
    stage.details.append(("점검 결과", "BETA-DETAIL"))


def test_summary_table_and_lists_at_top(iso, monkeypatch, capsys):
    _fake_runners(monkeypatch, [], {"diag": _fail_diag, "beta": _beta, "office": _skip_office})
    assert rat.main(["--label", "x"]) == 1                      # a failed stage -> exit 1
    text = (iso / _result_files(iso)[0]).read_text(encoding="utf-8")
    head = text[:text.index("## 1. ")]
    assert "| 1. 환경 진단 | FAIL | 3 | 1 | 0 |" in head
    assert "| 2. 오프라인 테스트 (pytest) | PASS | 1 | 0 | 0 |" in head
    assert "| 3. ml-kit 베타 점검 (run_beta_check) | PASS | 26 | 0 | 1 |" in head
    assert "| 4. office-cli 실기 러너 (실제 Office) | ENV-SKIP | 0 | 0 | 0 |" in head
    assert "### 실패 1건" in head and "- [diag] python-docx 기본 템플릿: zip 형식=False" in head
    assert "- [beta] F7 엑셀로 수식 재계산 확인" in head and "- [office] 단계 전체: Windows가 아님" in head
    assert "### UNVERIFIED (도구가 결과를 확인하지 못한 호출) 1건" in head and "- [beta] excel.py read" in head
    assert "BETA-DETAIL" not in head and "BETA-DETAIL" in text


def test_tool_error_logs_appended_at_the_end(iso, monkeypatch, capsys):
    logs = iso / "CLI" / "ml-kit" / "data" / "logs"
    logs.mkdir(parents=True)
    (logs / "errors.txt").write_text("===== 2026-11-04 | ml-kit 2.9.39\n명령: x\n\n===== 2026-11-04 | ml-kit 2.9.39\n명령: y\n",
                                     encoding="utf-8")
    _fake_runners(monkeypatch, [], {"beta": _beta})
    rat.main(["--only", "diag,beta", "--label", "x"])
    text = (iso / _result_files(iso)[0]).read_text(encoding="utf-8")
    assert "| 도구 오류 로그 | - | - | - | - | - | ml-kit 2건, office-cli 0건 |" in text
    assert text.index("BETA-DETAIL") < text.index("### ml-kit errors.txt") < text.index("명령: y")
    assert text.rstrip().endswith("```")                        # the log is the last thing in the file


def test_interrupt_and_stage_error_keep_the_file(iso, monkeypatch, capsys):
    def boom(stage):
        raise RuntimeError("runner bug")

    def stop(stage):
        stage.passed += 1
        raise KeyboardInterrupt
    order = []
    _fake_runners(monkeypatch, order, {"pytest": boom, "beta": stop})
    assert rat.main(["--label", "x"]) == 130
    assert order == ["diag", "pytest", "beta"]                 # an error goes on to the next stage, Ctrl+C stops
    text = (iso / _result_files(iso)[0]).read_text(encoding="utf-8")
    assert "| 2. 오프라인 테스트 (pytest) | FAIL | 0 | 1 | 0 |" in text and "runner bug" in text
    assert "| 3. ml-kit 베타 점검 (run_beta_check) | 중단 | 1 | 0 | 0 |" in text
    assert "| 4. office-cli 실기 러너 (실제 Office) | 대기 | - | - | - |" in text
    assert "사용자가 중단했습니다" in text and "## 도구 오류 로그" in text


def test_windows_stages_are_env_skip_without_windows(monkeypatch):
    monkeypatch.setattr(rat, "windows_ready", lambda: (False, "Windows가 아님"))
    for key, fn in (("office", rat.step_office), ("ml", rat.step_ml)):
        st = rat.Stage(key)
        fn(st, None, {"label": "x", "stamp": "s"})
        st.finish()
        assert st.status == "ENV-SKIP" and st.skips == ["단계 전체: Windows가 아님"] and not st.failed


def test_parse_office_runner_files():
    errors = ("## 2. 요약\n결과: ENV-SKIP=2, FAIL=1, MANUAL-SKIP=1, PASS=69\n문제 항목: D3(FAIL)\n"
              "건너뜀: D7(ENV-SKIP: PDF 내보내기 기능 없음, 추가 기능), F2(MANUAL-SKIP: --auto)\n")
    report = ("[PASS] A1 x\nUNVERIFIED 결과 1건:\n    excel.py read --range A1\n        unverified: ['수식 값을 확인하지 못함']\n")
    st = rat.Stage("office")
    rat.parse_office_logs(st, errors, report)
    assert (st.passed, st.failed, st.skipped) == (69, 1, 3) and st.failures == ["D3(FAIL)"]
    assert st.skips == ["D7(ENV-SKIP: PDF 내보내기 기능 없음, 추가 기능)", "F2(MANUAL-SKIP: --auto)"]
    assert st.unverified == ["excel.py read --range A1 -> ['수식 값을 확인하지 못함']"]
    clean = rat.Stage("office")
    rat.parse_office_logs(clean, "결과: PASS=3\n문제 항목: 없음\n", "UNVERIFIED 결과 0건 (없음)")
    assert (clean.passed, clean.failed, clean.failures, clean.unverified) == (3, 0, [], [])


def test_parse_pytest_summary():
    out = ("..F.s\nFAILED tests/test_a.py::test_x - assert 1 == 2\nSKIPPED [1] tests/test_b.py:3: Windows only\n"
           "= 1 failed, 3 passed, 1 skipped in 0.5s =")
    assert rat.parse_pytest(out) == (3, 1, 1, ["FAILED tests/test_a.py::test_x - assert 1 == 2"],
                                     ["SKIPPED [1] tests/test_b.py:3: Windows only"])
    assert rat.parse_pytest("12 passed in 3.0s")[:3] == (12, 0, 0)
    assert rat.parse_pytest("1 passed, 2 errors in 1s")[:3] == (1, 2, 0)


FAKE_BETA = r'''import json, os, sys
a = sys.argv
work = a[a.index("--dir") + 1]
os.makedirs(work, exist_ok=True)
rep = os.path.join(os.path.dirname(os.path.abspath(__file__)), "beta_check_x_1.md")
with open(rep, "w", encoding="utf-8") as f:
    f.write("| 묶음 | 항목 | 내용 | 결과 |\n|---|---|---|---|\n| H | H-a | 도움말 | PASS |\n| F | F4 | 일치도 | FAIL |\n"
            "| F | F7 | 재계산 | SKIP |\n\n## 실패·건너뜀 상세\n\n- F4 일치도 [FAIL]: 90%\n- F7 재계산 [SKIP]: 엑셀이 없는 환경\n")
open(rep[:-3] + ".xlsx", "wb").write(b"x")
logs = os.path.join(os.environ["ML_KIT_DATA_DIR"], "logs")
os.makedirs(logs)
open(os.path.join(logs, "errors.txt"), "w", encoding="utf-8").write("===== 일부러 낸 오류")
print(json.dumps({"status": "ok", "report_file": rep}))
'''


def test_beta_stage_moves_report_and_isolates_tool_log(iso, monkeypatch, capsys):
    """step_beta runs run_beta_check.py --agent, puts its report into the result, moves its own files into the
    test folder (one result file next to the runner) and points ML_KIT_DATA_DIR at the test folder."""
    fake = iso / "fake_beta.py"
    fake.write_text(FAKE_BETA, encoding="utf-8")
    work = iso / "work"
    monkeypatch.setattr(rat, "BETA_SCRIPT", str(fake))
    monkeypatch.setattr(rat, "beta_work_dir", lambda stamp: str(work))
    st = rat.Stage("beta")
    rat.step_beta(st, type("A", (), {"timeout": 60})(), {"label": "x", "stamp": "s"})
    st.finish()
    assert (st.status, st.passed, st.failed, st.skipped) == ("FAIL", 1, 1, 1)
    assert st.failures == ["F4 일치도: 90%"] and st.skips == ["F7 재계산: 엑셀이 없는 환경"]
    assert sorted(os.listdir(work)) == ["_data", "beta_check_x_1.md", "beta_check_x_1.xlsx"]
    assert not (iso / "beta_check_x_1.md").exists() and str(work) in st.note
    titles = [t for t, _ in st.details]
    assert titles[0] == "점검 결과 (beta_check 보고서)" and any("일부러 낸 오류" in t for t in titles)
    assert not (iso / "CLI" / "ml-kit" / "data").exists()       # the real tool log is not touched
