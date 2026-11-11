# -*- coding: utf-8 -*-
"""run_all_tests.py: a silent hung child is stopped at the timeout, CLI-dev/tests is a pytest target, PyYAML is not needed,
stages run in a fixed order into one result file with the summary on top and the tool error logs at the end."""
import importlib.util
import json
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
    # 상한 5초 (1초면 바쁜 회사 PC에서 파이썬이 'start'를 찍기 전에 끊김, 2026-11-10)
    code, out, sec, stop = rat.run_capture([sys.executable, "-c", "import time; print('start', flush=True); time.sleep(60)"],
                                           HERE, timeout=5, poll=0.2)
    # 회사 PC 2026-11-10: taskkill이 늦어 18초 만에 종료 코드 없이(None) 돌아옴 -> 직접 한 번 더 끝냄 (최대 약 25초 더)
    assert time.time() - t0 < 45 and not stop and code is not None and code != 0
    assert out.splitlines()[0] == "start" and "시간 초과" in out


def test_child_that_cannot_be_stopped_is_not_reported_as_no_code(monkeypatch):
    # 끝내지 못한 프로세스는 종료 코드 None 대신 -1과 [ERROR] 줄 (성공처럼 보이거나 판정이 깨지지 않게)
    monkeypatch.setattr(rat, "kill_tree", lambda proc: None)
    procs = []
    try:
        code, out, sec, stop = rat.run_capture([sys.executable, "-c", "import time; time.sleep(60)"], HERE,
                                               timeout=1, poll=0.2, procs=procs)
        assert code == -1 and "끝내지 못했습니다" in out, out
    finally:
        for p in procs:
            p.kill()
            p.wait(timeout=10)


def test_normal_child_output_and_exit_code(capsys):
    code, out, sec, stop = rat.run_capture([sys.executable, "-c", "print('a'); print('b'); raise SystemExit(3)"], HERE, timeout=60)
    assert (code, out, stop) == (3, "a\nb", False)


def test_pytest_targets_and_packages():
    assert "tests" in rat.PYTEST_TARGETS and os.path.isdir(os.path.join(rat.HERE, "tests"))
    assert "PyYAML" not in rat.PACKAGES                         # 시나리오는 agent_scenarios.py (회사 반입이 .yaml을 막음)
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


def test_live_failures_merge_with_summary():
    summary = ["FAILED tests/test_a.py::test_x - assert 1 == 2"]
    live = rat.parse_live("..F[LIVE] FAILED tests/test_a.py::test_x - assert 1 == 2\n"
                          "[LIVE] ERROR tests/test_b.py::test_y[a b] (setup) - ValueError: boom\n"
                          "[LIVE] ERROR tests/test_b.py::test_y[a b] (setup) - ValueError: boom\n")
    assert live == ["FAILED tests/test_a.py::test_x - assert 1 == 2", "ERROR tests/test_b.py::test_y[a b] (setup) - ValueError: boom"]
    assert rat.merge_live(summary, live) == summary + [live[1]]
    assert rat.merge_live([], live) == live                     # 시간 초과로 요약이 없을 때


SLOW_TESTS = '''import time
def test_fails_first():
    assert 1 == 2, "보이는 원인"
def test_ok():
    pass
def test_hangs():
    time.sleep(120)
'''


def test_pytest_stage_keeps_failure_names_when_stopped_by_time_limit(iso, monkeypatch, capsys):
    """회사 PC: ml-kit pytest가 시간 상한에 끊겨 실패 F 11개의 이름이 남지 않았다. 이제 플러그인이 실패를 생기는 즉시 [LIVE] 줄로
    내보내므로 끊겨도 실패 이름과 원인이 결과에 남고, 시간 초과도 따로 FAIL로 적힌다."""
    folder = iso / "tests"
    folder.mkdir()
    (folder / "test_slow.py").write_text(SLOW_TESTS, encoding="utf-8")
    (iso / "pytest_live_report.py").write_text(open(os.path.join(os.path.dirname(HERE), "pytest_live_report.py"),
                                                    encoding="utf-8").read(), encoding="utf-8")
    monkeypatch.setattr(rat, "PYTEST_TARGETS", ("tests",))
    st = rat.Stage("pytest")
    rat.step_pytest(st, type("A", (), {"pytest_timeout": 20})(), {})
    assert any("test_slow.py::test_fails_first" in f and "보이는 원인" in f for f in st.failures), st.failures
    assert any("시간 상한 20초로 중단" in f for f in st.failures), st.failures
    assert st.failed == 2


FLAKY_TESTS = '''import os
MARK = os.path.join(os.path.dirname(__file__), "ran_once.txt")
def test_fails_only_the_first_time():
    first = not os.path.exists(MARK)
    open(MARK, "w").close()
    assert not first, "처음만 실패"
def test_always_fails():
    assert 1 == 2, "늘 실패"
def test_ok():
    pass
'''


def test_pytest_stage_reruns_failed_tests_once_alone(iso, monkeypatch, capsys):
    """회사 PC 2026-11-11: PC가 느려진 시간대의 시간 초과 실패가 실행마다 달랐다. 실패한 시험만 혼자 한 번 다시 돌려, 통과하면
    참고(WARN)로, 다시 실패하면 실패로 센다. --no-rerun이면 다시 돌리지 않는다."""
    folder = iso / "tests"
    folder.mkdir()
    (folder / "test_flaky.py").write_text(FLAKY_TESTS, encoding="utf-8")
    (iso / "pytest_live_report.py").write_text(open(os.path.join(os.path.dirname(HERE), "pytest_live_report.py"),
                                                    encoding="utf-8").read(), encoding="utf-8")
    monkeypatch.setattr(rat, "PYTEST_TARGETS", ("tests",))
    st = rat.Stage("pytest")
    rat.step_pytest(st, type("A", (), {"pytest_timeout": 60})(), {})
    assert st.failed == 1 and st.passed == 2, (st.failed, st.passed, st.failures)
    assert len(st.failures) == 1 and "test_always_fails" in st.failures[0], st.failures
    assert any("test_fails_only_the_first_time" in w and "다시 돌려 통과" in w and "처음 원인: AssertionError" in w for w in st.warnings), st.warnings
    assert "실패한 시험만 다시 돌림: 1건 통과, 1건 다시 실패" in st.details[-1][1]
    os.remove(folder / "ran_once.txt")
    st = rat.Stage("pytest")
    rat.step_pytest(st, type("A", (), {"pytest_timeout": 60, "no_rerun": True})(), {})
    assert st.failed == 2 and st.passed == 1 and not any("다시 돌려" in w for w in st.warnings)


def test_diag_template_damage_is_warn_when_tools_do_not_need_it(monkeypatch):
    """python-pptx 기본 템플릿은 어느 도구도 쓰지 않음 -> 참고(WARN). python-docx 기본 템플릿 손상은 ml-kit 보고서가
    폴더 템플릿으로 복구해 빈 문서를 만들 수 있으면 참고, 못 만들면 FAIL."""
    monkeypatch.setattr(rat, "diag_docx", lambda lines: lines.append("docx.Document(): 스타일 4개 (정상 164), Caption 있음=False"))
    monkeypatch.setattr(rat, "diag_pptx", lambda lines: lines.append("pptx.Presentation() 실패: PackageNotFoundError"))
    monkeypatch.setattr(rat, "diag_new_files", lambda lines: lines.append("새로 쓴 xlsx 다시 열기: 정상"))
    monkeypatch.setattr(rat, "diag_system", lambda lines: None)
    monkeypatch.setattr(rat, "diag_skills", lambda stage, lines: None)
    for report_ok, failed in ((True, 0), (False, 1)):
        monkeypatch.setattr(rat, "diag_report_docx", lambda lines, ok=report_ok: ok)
        st = rat.Stage("diag")
        rat.step_diag(st, None, {})
        assert st.failed == failed, st.failures
        assert any("python-pptx" in w and "쓰지 않음" in w for w in st.warnings), st.warnings
        assert any("python-docx" in w for w in st.warnings) == report_ok


def test_report_doc_check_runs_the_real_tool_function():
    lines = []
    assert rat.diag_report_docx(lines) is True, lines
    assert "ml_report.new_report_document): 정상" in lines[0]


def test_warnings_listed_at_top_without_failing(iso, monkeypatch, capsys):
    _fake_runners(monkeypatch, [], {"diag": lambda st: (setattr(st, "passed", 1), st.warnings.append("참고할 일"))})
    assert rat.main(["--only", "diag", "--label", "w"]) == 0
    text = (iso / _result_files(iso)[0]).read_text(encoding="utf-8")
    assert "### 참고 (WARN, 실패로 세지 않음) 1건" in text and "- [diag] 참고할 일" in text


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


# ---------------------------------------------------------------------------
# gemma: agent_scenarios.py 요청을 opencode run --format json으로 한 세션에 보내고 채점
# ---------------------------------------------------------------------------
FAKE_OPENCODE = r'''#!/usr/bin/env python3
"""가짜 opencode: run --format json [--session ID] <요청>. 받은 인자를 calls.jsonl에 남기고 OpenCode 2.x 형식 사건을 씁니다.
S0 요청이면 make_beta_samples.py를 shell 도구로 실행한 것처럼 기록합니다."""
import json, os, sys
argv = sys.argv[1:]
with open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "calls.jsonl"), "a", encoding="utf-8") as f:
    f.write(json.dumps(argv, ensure_ascii=False) + "\n")
if os.environ.get("FAKE_OC_DEAD"):
    sys.stderr.write("provider not found\n")
    sys.exit(1)
sid = argv[argv.index("--session") + 1] if "--session" in argv else "ses_fake1"
msg = argv[-1]
model = argv[argv.index("--model") + 1] if "--model" in argv else ""
if model and model in os.environ.get("FAKE_OC_LIMITED", "").split(","):     # 이 모델은 사용량 제한
    sys.stdout.write(json.dumps({"type": "error", "timestamp": 1, "sessionID": sid,
                                 "error": {"name": "APIError", "data": {"message": "429 Too Many Requests: rate limit exceeded"}}}) + "\n")
    sys.exit(1)
def ev(kind, part, ts):
    part = dict(part, sessionID=sid, messageID="msg_" + str(ts))
    sys.stdout.write(json.dumps({"type": kind, "timestamp": ts, "sessionID": sid, "part": part}, ensure_ascii=False) + "\n")
if msg.startswith("[S0]"):
    out = json.dumps({"status": "ok", "agent_directive": "VERIFIED: 샘플 3개"}, ensure_ascii=False)
    ev("tool_use", {"type": "tool", "tool": "shell", "state": {"status": "completed",
       "input": {"command": 'python "C:\\CLI-dev\\ml-kit\\tools\\make_beta_samples.py" "C:\\CLI-dev\\work\\gemma"'}, "output": out}}, 1000)
    ev("text", {"type": "text", "text": "샘플 3개를 만들었습니다."}, 2000)
else:
    ev("text", {"type": "text", "text": "모르겠습니다."}, 3000)
'''


@pytest.fixture
def fake_oc(tmp_path):
    for tool, name in (("ml-kit", "make_beta_samples.py"), ("office-cli", "make_office_samples.py")):   # 러너가 직접 실행하는 준비 스크립트
        script = tmp_path / tool / "tools" / name
        script.parent.mkdir(parents=True, exist_ok=True)
        script.write_text("import sys, os\nos.makedirs(sys.argv[1], exist_ok=True)\nopen(os.path.join(sys.argv[1], '%s.done'), 'w').close()\n"
                          "print('{\"status\": \"ok\"}')\n" % name, encoding="utf-8")
    exe = tmp_path / "bin" / "opencode"
    exe.parent.mkdir()
    exe.write_text(FAKE_OPENCODE, encoding="utf-8")
    exe.chmod(0o755)
    return exe


def _calls(exe):
    with open(exe.parent / "calls.jsonl", encoding="utf-8") as f:
        return [json.loads(x) for x in f]


@pytest.mark.skipif(os.name == "nt", reason="가짜 opencode는 #! 스크립트 (Linux)")
def test_gemma_stage_sends_all_scenarios_in_one_session_and_grades(iso, fake_oc, capsys):
    work = iso / "work" / "gemma"
    work.mkdir(parents=True)
    (work / "old.txt").write_text("x", encoding="utf-8")
    code = rat.main(["--only", "gemma", "--label", "t", "--opencode", str(fake_oc), "--gemma-dir", str(work)])
    assert code == 1                                            # 가짜는 S0만 맞게 답함
    _, by_id, order = rat._load_grader().load_scenarios()
    manual = [s for s in order if by_id[s].get("manual") or by_id[s].get("paused")]   # 수동 전용(O3)·시험에서 뺀 요청(R)
    setup = [s for s in order if by_id[s].get("setup")]             # 시험 자료 준비 (S0, S1): 러너가 직접 실행
    calls = _calls(fake_oc)
    assert manual and setup == ["S0", "S1"] and len(calls) == len(order) - len(manual) - len(setup)
    assert not any(c[-1].startswith(f"[{s}]") for c in calls for s in manual + setup)
    assert (work / "make_beta_samples.py.done").exists() and (work / "make_office_samples.py.done").exists()
    assert calls[0][:3] == ["run", "--format", "json"] and "--session" not in calls[0]
    assert all(c[c.index("--session") + 1] == "ses_fake1" for c in calls[1:])      # 같은 세션에 이어서
    assert calls[0][-1].startswith("[A1] ") and str(work) in calls[0][-1] and "{DIR}" not in calls[0][-1]
    assert any(p.name.startswith("gemma_old_") for p in work.parent.iterdir())     # 옛 폴더는 지우지 않고 옮김
    runs = next(p for p in (iso / "work").iterdir() if p.name.startswith("gemma_runs_"))   # 시험 흔적은 모두 CLI-dev\work
    assert {f"{s}.jsonl" for s in order if s not in manual + setup} <= {p.name for p in runs.iterdir()}
    assert any(p.name.startswith("grade_t_") and p.suffix == ".md" for p in runs.iterdir())
    text = (iso / _result_files(iso)[0]).read_text(encoding="utf-8")
    assert "Gemma 종합 테스트" in text and "ses_fake1" in text and "채점 결과" in text
    assert "[gemma] S0" not in text and "[gemma] A1" in text                     # S0은 세지 않음, A1 실패로 목록에
    for s in manual:                                                             # 수동 전용은 실패가 아니라 건너뜀
        why = "시험에서 뺌" if by_id[s].get("paused") else "수동 전용"
        assert text.count(f"[gemma] {s}: ") == 1 and f"[gemma] {s}: {why}" in text


@pytest.mark.skipif(os.name == "nt", reason="가짜 opencode는 #! 스크립트 (Linux)")
def test_gemma_stage_stops_when_opencode_gives_no_session(iso, fake_oc, monkeypatch, capsys):
    monkeypatch.setenv("FAKE_OC_DEAD", "1")
    work = iso / "work" / "gemma"
    assert rat.main(["--only", "gemma", "--opencode", str(fake_oc), "--gemma-dir", str(work)]) == 1
    assert len(_calls(fake_oc)) == 1                            # 첫 요청에서 멈춤 (나머지 20개를 기다리지 않음)
    text = (iso / _result_files(iso)[0]).read_text(encoding="utf-8")
    assert "opencode가 응답하지 않음" in text and "provider not found" in text


@pytest.mark.skipif(os.name == "nt", reason="가짜 opencode는 #! 스크립트 (Linux)")
def test_gemma_only_sends_chosen_requests_with_model(iso, fake_oc, capsys):
    """--gemma-only: 고른 요청만 시나리오 순서대로, --gemma-model: opencode --model, 요청별 모델을 models.json과 보고서에."""
    work = iso / "work" / "gemma"
    rat.main(["--only", "gemma", "--label", "t", "--opencode", str(fake_oc), "--gemma-dir", str(work),
              "--gemma-only", "p1,A1", "--gemma-model", "or/gemma-4"])
    calls = _calls(fake_oc)
    assert [c[-1][:4] for c in calls] == ["[A1]", "[P1]"]                          # 준 순서가 아니라 시나리오 순서
    assert (work / "make_beta_samples.py.done").exists()                           # 준비는 고르지 않아도 러너가 직접
    assert all(c[c.index("--model") + 1] == "or/gemma-4" for c in calls)
    runs = next(p for p in (iso / "work").iterdir() if p.name.startswith("gemma_runs_"))
    assert json.loads((runs / "models.json").read_text(encoding="utf-8")) == {"A1": "or/gemma-4", "P1": "or/gemma-4"}
    grade = next(p for p in runs.iterdir() if p.name.startswith("grade_t_") and p.suffix == ".md").read_text(encoding="utf-8")
    assert "요청별 모델: A1=or/gemma-4, P1=or/gemma-4" in grade
    text = (iso / _result_files(iso)[0]).read_text(encoding="utf-8")
    assert "[gemma] M1" not in text and "[gemma] S0" not in text                   # 고르지 않은 요청, 준비 요청은 세지 않음


@pytest.mark.skipif(os.name == "nt", reason="가짜 opencode는 #! 스크립트 (Linux)")
def test_gemma_model_falls_back_on_rate_limit(iso, fake_oc, monkeypatch, capsys):
    """사용량 제한이면 예비 모델로 같은 요청을 다시 보내고, 그 뒤 요청도 예비 모델로. 제한 기록은 이름을 바꿔 남김."""
    monkeypatch.setenv("FAKE_OC_LIMITED", "main")
    work = iso / "work" / "gemma"
    rat.main(["--only", "gemma", "--label", "t", "--opencode", str(fake_oc), "--gemma-dir", str(work),
              "--gemma-only", "A1,P1", "--gemma-model", "main,spare"])
    calls = _calls(fake_oc)
    assert [(c[-1][:4], c[c.index("--model") + 1]) for c in calls] == [("[A1]", "main"), ("[A1]", "spare"), ("[P1]", "spare")]
    runs = next(p for p in (iso / "work").iterdir() if p.name.startswith("gemma_runs_"))
    assert json.loads((runs / "models.json").read_text(encoding="utf-8")) == {"A1": "spare", "P1": "spare"}
    assert (runs / "A1.limited1.jsonl").exists() and (runs / "A1.jsonl").exists()
    grade = next(p for p in runs.iterdir() if p.name.startswith("grade_t_") and p.suffix == ".md").read_text(encoding="utf-8")
    assert "429" not in grade.split("## 요청별 기록")[1].split("### P1")[0]             # 예비 모델의 답으로 채점


@pytest.mark.skipif(os.name == "nt", reason="가짜 opencode는 #! 스크립트 (Linux)")
def test_gemma_rate_limit_without_spare_is_unverified_and_unknown_id_fails(iso, fake_oc, monkeypatch, capsys):
    monkeypatch.setenv("FAKE_OC_LIMITED", "main")
    work = iso / "work" / "gemma"
    rat.main(["--only", "gemma", "--label", "t", "--opencode", str(fake_oc), "--gemma-dir", str(work),
              "--gemma-only", "A1", "--gemma-model", "main"])
    assert len(_calls(fake_oc)) == 1
    text = (iso / _result_files(iso)[0]).read_text(encoding="utf-8")
    assert "사용량 제한에 걸렸고 예비 모델이 없음" in text
    for f in _result_files(iso):
        (iso / f).unlink()
    assert rat.main(["--only", "gemma", "--opencode", str(fake_oc), "--gemma-dir", str(work), "--gemma-only", "A1,Z9"]) == 1
    assert len(_calls(fake_oc)) == 1                                               # 모르는 ID면 아무것도 보내지 않음
    assert "모르는 요청 ID: Z9" in (iso / _result_files(iso)[0]).read_text(encoding="utf-8")


def test_all_test_folders_default_under_cli_dev_work(iso):
    """반입 때 C:\\CLI-dev를 통째로 지우면 시험 흔적도 사라지도록 기본 시험 폴더는 모두 CLI-dev\\work 아래."""
    work = str(iso / "work")
    assert rat.work_root() == work and rat.beta_work_dir("s").startswith(work)
    assert rat._load_grader().load_scenarios()[0]["dir"] == "C:\\CLI-dev\\work\\gemma"
    for rel in ("run_beta_check.py", os.path.join("office-cli", "tools", "run_win_tests.py"),
                os.path.join("ml-kit", "tools", "run_win_tests.py"), "run_all_tests.py"):
        with open(os.path.join(os.path.dirname(HERE), rel), encoding="utf-8") as f:
            text = f.read()
        assert not any(old in text for old in ("ml_beta", "ml_test", "office_test\"", "gemma_test")), rel


def test_gemma_is_opt_in_and_env_skip_without_opencode(iso, monkeypatch, capsys):
    order = []
    _fake_runners(monkeypatch, order)
    rat.main(["--label", "a"])
    assert "gemma" not in order
    order.clear()
    rat.main(["--label", "a", "--gemma"])
    assert order == ["diag", "pytest", "beta", "office", "gemma"]
    monkeypatch.setattr(rat.shutil, "which", lambda name: None)
    assert rat.find_opencode() is None and rat.find_opencode(str(iso / "없음.exe")) is None
    stage = rat.Stage("gemma")
    rat.step_gemma(stage, rat.argparse.Namespace(opencode=None, gemma_dir=None, gemma_timeout=5), {"stamp": "s", "label": "a"})
    assert stage.status == "ENV-SKIP" and "opencode 없음" in stage.note


def test_result_file_masks_user_folder_names():
    """2026-11-07 집 PC: 결과 파일을 올리기 전 사람이 가리다 잘린 출력('C:\\\\Users\\\\Joo...')을 놓쳤음. 러너가 쓸 때 가림."""
    import importlib.util
    dev = os.path.dirname(HERE)
    for path in (os.path.join(dev, "run_all_tests.py"), os.path.join(dev, "run_simplify_check.py")):
        spec = importlib.util.spec_from_file_location("mask_" + os.path.basename(path)[:-3], path)
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
        m = mod.mask_private
        assert m('"C:\\\\Users\\\\Joo...') == '"C:\\\\Users\\\\<사용자>'
        assert m("C:\\Users\\kim\\AppData") == "C:\\Users\\<사용자>\\AppData"
        assert m("C:/Users/kim/Desktop") == "C:/Users/<사용자>/Desktop"
        assert m("C:\\Users\\<사용자>\\.config") == "C:\\Users\\<사용자>\\.config"
        assert m("Users 수 3명, kim") == "Users 수 3명, kim"             # 경로 밖은 그대로
        # 2026-11-07 집 PC: Office 변경 내용·메모의 작성자(Office 사용자 이름)가 결과 파일에 남음
        assert m("{'no': 1, 'author': 'kim', 'date': 'x'}") == "{'no': 1, 'author': '<사용자>', 'date': 'x'}"
        assert m('[{"cell": "C1", "author": "Kim Lee", "text": "a"}]') == '[{"cell": "C1", "author": "<사용자>", "text": "a"}]'
        assert m('{\\"author\\": \\"kim\\"}') == '{\\"author\\": \\"<사용자>\\"}'   # JSON 안의 JSON
        assert m("author 열 3개") == "author 열 3개"


def test_simplify_dry_run_sheet_delete_uses_name():
    """2026-11-07 집 PC D2: sheet delete에 --sheet를 넘겨 'sheet delete에는 --name이 필요합니다'로 실패.
    D 항목은 Office가 필요해 리눅스에서 SKIP되므로, 실행하지 않고 dry()에 넘기는 인자 목록을 소스에서 확인한다."""
    import ast
    path = os.path.join(os.path.dirname(HERE), "run_simplify_check.py")
    with open(path, encoding="utf-8") as f:
        tree = ast.parse(f.read())
    cases = [ast.literal_eval(n.args[5]) for n in ast.walk(tree)
             if isinstance(n, ast.Call) and getattr(n.func, "id", None) == "dry" and len(n.args) >= 6]
    sheet = [a for a in cases if a and a[0] == "sheet"]
    assert sheet, "dry() 안의 sheet 점검을 찾지 못함"
    for argv in sheet:
        assert "--name" in argv and "--sheet" not in argv, argv


def test_pytest_workers_default_is_capped_by_cores_and_memory(monkeypatch):
    """회사 PC 2026-11-10: 코어 수 - 1개로 돌리자 메모리가 모자라 pytest가 89분, 시간 초과 18건. 기본은 최대 3개, 메모리 4GB마다 1개."""
    args = type("A", (), {"pytest_workers": 0})()
    monkeypatch.setattr(rat.os, "cpu_count", lambda: 16)
    monkeypatch.setattr(rat, "total_memory_gb", lambda: 64.0)
    assert rat.pytest_workers(args) == 3
    monkeypatch.setattr(rat, "total_memory_gb", lambda: 8.0)
    assert rat.pytest_workers(args) == 2
    monkeypatch.setattr(rat, "total_memory_gb", lambda: 3.0)
    assert rat.pytest_workers(args) == 1
    monkeypatch.setattr(rat.os, "cpu_count", lambda: 2)
    monkeypatch.setattr(rat, "total_memory_gb", lambda: None)
    assert rat.pytest_workers(args) == 1
    args.pytest_workers = 6
    assert rat.pytest_workers(args) == 6
    assert rat.total_memory_gb() is None or rat.total_memory_gb() > 0


def test_slow_python_start_windows_are_reported():
    # 회사 PC 2026-11-10 낮: 파일 정리 작업 시간대에 프로세스를 띄우는 시험이 수십 배 느려짐 -> 그 시간대를 결과에 남김
    s = [("12:00", 0.4), ("12:01", 0.5), ("12:02", 9.0), ("12:03", 30.0), ("12:04", 0.6), ("12:05", 3.0)]
    assert rat.slow_windows(s) == ["12:02~12:03 파이썬 시작 최대 30.0초 (평소 0.4초)", "12:05~12:05 파이썬 시작 최대 3.0초 (평소 0.4초)"]
    assert rat.slow_windows([("12:00", 0.4), ("12:01", 1.9)]) == [] and rat.slow_windows([]) == []
    # 재부팅 직후처럼 첫 측정부터 느리면: 가장 빠른 측정이 평소 (회사 PC 2026-11-10 15:07 2.4초, 15:29 이후 1초)
    late = [("15:07", 2.36), ("15:12", 16.1), ("15:29", 1.05), ("15:41", 0.59)]
    assert rat.slow_windows(late) == ["15:12~15:12 파이썬 시작 최대 16.1초 (평소 0.59초)"]
    assert rat.probe_python() is not None


def test_long_files_start_first_across_tools(tmp_path):
    # 오래 걸리는 파일을 도구와 관계없이 먼저 (끝에 긴 파일 셋이 함께 몰리던 것)
    ml, oc = tmp_path / "ml-kit" / "tests", tmp_path / "office-cli" / "tests"
    for d, names in ((ml, ["test_train.py", "test_skill_docs.py"]), (oc, ["test_reader.py", "test_word.py"])):
        d.mkdir(parents=True)
        for n in names:
            (d / n).write_text("x" * 100, encoding="utf-8")
    jobs = [("ml-kit", str(tmp_path / "ml-kit"), str(ml / "test_skill_docs.py")),
            ("office-cli", str(tmp_path / "office-cli"), str(oc / "test_word.py")),
            ("office-cli", str(tmp_path / "office-cli"), str(oc / "test_reader.py")),
            ("ml-kit", str(tmp_path / "ml-kit"), str(ml / "test_train.py"))]
    jobs.sort(key=lambda j: -rat.job_seconds(*j))
    assert [os.path.basename(j[2]) for j in jobs[:2]] == ["test_train.py", "test_reader.py"]
