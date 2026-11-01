"""_runtime.py 계약 테스트: stdout 1줄 JSON, 종료 코드, agent_directive, 경로/입출력 도우미."""
import datetime
import json
import os
import subprocess
import sys
import textwrap

import pytest

import _runtime as rt

DEV = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))  # CLI-dev\office-cli (tests, tools)
TOOL_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))), "CLI", "office-cli")  # 실행 파일은 CLI\office-cli


def run_script(tmp_path, body, *args):
    """_runtime을 쓰는 임시 진입점을 하위 프로세스로 실행합니다. (종료 코드, JSON, stderr)"""
    script = tmp_path / "tool.py"
    script.write_text(
        "import sys\n"
        f"sys.path.insert(0, {TOOL_DIR!r})\n"
        "import _runtime as rt\n" + textwrap.dedent(body),
        encoding="utf-8",
    )
    proc = subprocess.run([sys.executable, str(script), *args], capture_output=True, timeout=60,
                          env=dict(os.environ, OFFICE_CLI_DATA_DIR=str(tmp_path / "data"),
                                   OFFICE_CLI_CACHE_DIR=str(tmp_path / "cache")))
    out = proc.stdout.decode("utf-8")
    lines = [ln for ln in out.splitlines() if ln.strip()]
    assert len(lines) == 1, f"stdout은 정확히 1줄이어야 함: {out!r} / stderr: {proc.stderr.decode('utf-8', 'replace')}"
    return proc.returncode, json.loads(lines[0]), proc.stderr.decode("utf-8", "replace")


TOOL_BODY = """
parser, sub = rt.build_parser("demo", "demo tool")
p = sub.add_parser("go")
p.add_argument("--mode", required=True)
args = parser.parse_args()

def execute():
    if args.mode == "ok":
        rt.log("PROGRESS", "작업 중")
        return rt.ok("demo", "go", "완료했습니다", value="한글")
    if args.mode == "input":
        raise rt.not_found("시트", "Sheet3", ["Sheet1", "Sheet2", "시트3"])
    if args.mode == "busy":
        class com_error(Exception):
            pass
        raise com_error(-2146777998, "예외", (0, "Microsoft Excel", "편집 중", None, 0, -2146777998), None)
    if args.mode == "interrupt":
        raise KeyboardInterrupt()
    raise RuntimeError("버그")

rt.run_cli("demo", "go", execute)
"""


def test_ok_response_is_one_json_line_exit_0(tmp_path):
    code, payload, err = run_script(tmp_path, TOOL_BODY, "go", "--mode", "ok")
    assert code == 0
    assert payload["status"] == "ok" and payload["value"] == "한글"
    assert payload["agent_directive"].startswith("VERIFIED:")
    assert "[PROGRESS] 작업 중" in err


def test_input_error_is_stop_retry_with_suggestion(tmp_path):
    code, payload, _ = run_script(tmp_path, TOOL_BODY, "go", "--mode", "input")
    assert code == 1
    assert payload["status"] == "error" and payload["error_type"] == "ambiguous"
    assert payload["agent_directive"].startswith("STOP_RETRY:")
    assert "[후보]" in payload["message"] and "[추천]" not in payload["message"]   # Sheet3 -> Sheet1/Sheet2는 추측하지 않음


def test_com_busy_error_is_explained(tmp_path):
    code, payload, _ = run_script(tmp_path, TOOL_BODY, "go", "--mode", "busy")
    assert code == 1
    assert payload["error_type"] == "office_busy"
    assert payload["hresult"] == "0x800AC472"
    assert "Esc" in payload["message"]


def test_internal_error_is_error_directive(tmp_path):
    code, payload, _ = run_script(tmp_path, TOOL_BODY, "go", "--mode", "boom")
    assert code == 1
    assert payload["error_type"] == "internal"
    assert payload["agent_directive"].startswith("ERROR:")


def test_interrupt_exit_130(tmp_path):
    code, payload, _ = run_script(tmp_path, TOOL_BODY, "go", "--mode", "interrupt")
    assert code == 130 and payload["error_type"] == "interrupted"


def test_argument_error_is_json_exit_2(tmp_path):
    code, payload, _ = run_script(tmp_path, TOOL_BODY, "go")
    assert code == 2
    assert payload["status"] == "error" and payload["tool"] == "demo"
    assert payload["agent_directive"].startswith("STOP_RETRY:")


def test_missing_subcommand_is_json_exit_2(tmp_path):
    code, payload, _ = run_script(tmp_path, TOOL_BODY)
    assert code == 2 and payload["status"] == "error"


# ---------------------------------------------------------------------------
# 도우미 함수
# ---------------------------------------------------------------------------

def test_suggest_and_find_name():
    assert rt.find_name(" sheet1 ", ["Sheet1", "Sheet2"]) == "Sheet1"
    assert rt.suggest("시트 2", ["시트1", "시트2", "요약"])[0] == "시트2"
    assert rt.suggest("매출보고", ["10월 매출보고서", "요약"])[0] == "10월 매출보고서"
    msg = str(rt.not_found("시트", "없는", ["A", "B"]))
    assert "사용 가능한 시트: A, B" in msg


def test_json_safe():
    data = {"nan": float("nan"), "d": datetime.date(2026, 10, 27), "t": (1, 2), "b": b"x"}
    assert rt.json_safe(data) == {"nan": None, "d": "2026-10-27", "t": [1, 2], "b": "x"}


def test_read_text_file_encodings(tmp_path):
    p = tmp_path / "a.txt"
    p.write_bytes("한글 테스트".encode("cp949"))
    assert rt.read_text_file(str(p)) == "한글 테스트"
    p.write_bytes("한글 테스트".encode("utf-16"))
    assert rt.read_text_file(str(p)) == "한글 테스트"
    p.write_bytes("\ufeff한글".encode("utf-8"))
    assert rt.read_text_file(str(p)) == "한글"


def test_read_csv_rows_detects_delimiter(tmp_path):
    p = tmp_path / "a.csv"
    p.write_text("이름;값\n가;1\n나;2\n", encoding="utf-8")
    assert rt.read_csv_rows(str(p)) == [["이름", "값"], ["가", "1"], ["나", "2"]]
    p.write_text("이름\t값\n가\t1\n", encoding="utf-8")
    assert rt.read_csv_rows(str(p))[1] == ["가", "1"]


def test_write_csv_roundtrip_utf8_sig(tmp_path):
    path = rt.write_csv_rows(str(tmp_path / "o.csv"), [["x", "sin(x)"], [0, None]])
    assert open(path, "rb").read().startswith(b"\xef\xbb\xbf")
    assert rt.read_csv_rows(path) == [["x", "sin(x)"], ["0", ""]]


def test_save_with_lock_fallback_uses_new_name(tmp_path):
    target = str(tmp_path / "locked.csv")
    calls = []

    def writer(p):
        calls.append(p)
        if p == target:
            raise PermissionError("locked")
        open(p, "w").close()

    saved = rt.save_with_lock_fallback(target, writer)
    assert saved != target and os.path.exists(saved) and calls[0] == target


def test_resolve_output_path_rules():
    auto = rt.resolve_output_path(None, "search", ".md")
    assert os.path.dirname(auto) == rt.get_work_dir() and auto.endswith(".md")
    named = rt.resolve_output_path("report.docx", "x", ".pdf")
    assert named == os.path.join(rt.get_data_dir(), "report.pdf")
    noext = rt.resolve_output_path("result", "x", "csv")
    assert noext.endswith("result.csv")


def test_resolve_input_path_suggests(tmp_path):
    data = rt.get_data_dir()
    open(os.path.join(data, "매출_10월.csv"), "w").close()
    assert rt.resolve_input_path("매출_10월.csv") == os.path.join(data, "매출_10월.csv")
    with pytest.raises(rt.InputError) as ei:
        rt.resolve_input_path("매출_10월.xlsx")
    assert "[후보] 매출_10월.csv" in str(ei.value) and "[추천]" not in str(ei.value)   # 다른 형식 파일은 추측하지 않음


def test_cleanup_work_dir_removes_old_files():
    work = rt.get_work_dir()
    old = os.path.join(work, "old.md")
    new = os.path.join(work, "new.md")
    for p in (old, new):
        open(p, "w").close()
    past = datetime.datetime.now().timestamp() - 30 * 86400
    os.utime(old, (past, past))
    assert rt.cleanup_work_dir() == 1
    assert not os.path.exists(old) and os.path.exists(new)


def test_config_roundtrip(tmp_path, monkeypatch):
    monkeypatch.setattr(rt, "CONFIG_FILE", str(tmp_path / "config.json"))
    monkeypatch.setattr(rt, "_config_cache", None)
    assert rt.load_config() == {}
    rt.set_config("transfer_method", {"ppt": "relay"})
    monkeypatch.setattr(rt, "_config_cache", None)
    assert rt.load_config()["transfer_method"] == {"ppt": "relay"}


def test_clamp_rows_and_text():
    rows = [[i, "x" * 10] for i in range(500)]
    clipped, cut = rt.clamp_rows(rows, max_rows=50)
    assert len(clipped) == 50 and cut
    clipped, cut = rt.clamp_rows(rows[:5])
    assert len(clipped) == 5 and not cut
    text, cut = rt.clamp_text("가" * 10, max_chars=4)
    assert text == "가가가가" and cut


def test_next_command_quotes_paths():
    cmd = rt.next_command("office.py", ["transfer", "--chart", "사인 2주기", "--to", "word,ppt"], "차트 전달")
    assert cmd["purpose"] == "차트 전달"
    assert '"사인 2주기"' in cmd["command"] and '"word,ppt"' in cmd["command"]
    assert "office.py" in cmd["command"]


def test_markdown_blocks_and_bold():
    blocks = rt.parse_markdown_blocks("# 제목\n본문 **굵게**\n이어짐\n\n- a\n- b\n1. x\n2) y\n| 공정 | 수율 |\n|---|---|\n| 가열 | 98 |\n---\n## 끝")
    assert blocks == [("heading", (1, "제목")), ("paragraph", "본문 **굵게** 이어짐"), ("bullets", ["a", "b"]),
                      ("numbered", ["x", "y"]), ("table", [["공정", "수율"], ["가열", "98"]]), ("break", None),
                      ("heading", (2, "끝"))]
    assert rt.split_bold("가 **나** 다 __라__") == ("가 나 다 라", [(2, 3), (6, 7)])
    assert rt.looks_like_markdown("- 항목") and not rt.looks_like_markdown("그냥 문장")


# ---------------------------------------------------------------------------
# R6 오류 로그 (<데이터 폴더>/logs/errors.txt)
# ---------------------------------------------------------------------------

def _error_log(tmp_path):
    p = tmp_path / "data" / "logs" / "errors.txt"
    return p.read_text(encoding="utf-8") if p.exists() else ""


def test_error_log_input_error_without_traceback(tmp_path):
    code, _, _ = run_script(tmp_path, TOOL_BODY, "go", "--mode", "input")
    assert code == 1
    text = _error_log(tmp_path)
    assert text.startswith(f"===== ") and f"| office-cli {rt.VERSION} | Python " in text
    assert "명령: demo.py go --mode input" in text
    assert "결과: ambiguous | 시트 'Sheet3'" in text
    assert "traceback:" not in text
    assert text.endswith("\n\n")


def test_error_log_internal_error_has_traceback(tmp_path):
    code, payload, _ = run_script(tmp_path, TOOL_BODY, "go", "--mode", "crash")
    assert code == 1 and payload["agent_directive"].startswith("ERROR:")
    text = _error_log(tmp_path)
    assert "결과: internal | RuntimeError: 버그" in text
    assert "traceback:" in text and "RuntimeError" in text.split("traceback:")[1]
    assert "errors.txt" not in json.dumps(payload, ensure_ascii=False)     # 로그 경로는 stdout JSON에 넣지 않음


def test_error_log_argument_error_exit_2(tmp_path):
    code, _, _ = run_script(tmp_path, TOOL_BODY, "go", "--bad")
    assert code == 2
    text = _error_log(tmp_path)
    assert "명령: demo.py go --bad" in text and "결과: input | CLI 인자 오류" in text


def test_error_log_skips_success_and_interrupt(tmp_path):
    assert run_script(tmp_path, TOOL_BODY, "go", "--mode", "ok")[0] == 0
    assert run_script(tmp_path, TOOL_BODY, "go", "--mode", "interrupt")[0] == 130
    assert not (tmp_path / "data" / "logs" / "errors.txt").exists()


def test_error_log_rotates_over_limit(monkeypatch):
    path = rt.error_log_path()
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        f.write("old\n")
    monkeypatch.setattr(rt, "ERROR_LOG_MAX", 2)
    rt.write_error_log("demo", ["go"], {"error_type": "input", "message": "새 항목"})
    assert open(path[:-4] + ".old.txt", encoding="utf-8").read() == "old\n"
    new = open(path, encoding="utf-8").read()
    assert new.startswith("===== ") and "결과: input | 새 항목" in new


def test_error_log_rotation_failure_still_writes(monkeypatch):
    path = rt.error_log_path()
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        f.write("old\n")
    monkeypatch.setattr(rt, "ERROR_LOG_MAX", 2)

    def deny(*a, **k):
        raise PermissionError("locked")
    monkeypatch.setattr(rt.os, "replace", deny)
    rt.write_error_log("demo", ["go"], {"error_type": "input", "message": "회전 실패"})
    assert "결과: input | 회전 실패" in open(path, encoding="utf-8").read()


def test_error_log_recovery_attempts_line(monkeypatch):
    monkeypatch.setattr(rt, "RECOVERY_ATTEMPTS", [])
    monkeypatch.setattr(rt.time, "sleep", lambda s: None)

    def locked():
        raise PermissionError(13, "locked", "report.xlsx")
    with pytest.raises(PermissionError):
        rt.retry_fs(locked, attempts=3)
    rt.write_error_log("demo", ["go"], {"error_type": "input", "message": "잠김"})
    text = open(rt.error_log_path(), encoding="utf-8").read()
    assert "복구 시도: 파일 작업 report.xlsx 3회 재시도 실패" in text


def test_error_log_failure_does_not_break_output(tmp_path):
    body = "rt.error_log_path = lambda: (_ for _ in ()).throw(OSError('disk'))\n" + TOOL_BODY
    code, payload, _ = run_script(tmp_path, body, "go", "--mode", "crash")
    assert code == 1 and payload["status"] == "error"
    code, payload, _ = run_script(tmp_path, body, "go", "--bad")
    assert code == 2 and payload["status"] == "error"


def test_error_log_records_in_process_argv(capsys):
    import reader
    with pytest.raises(SystemExit) as ei:
        reader.main(["read", "--file", "없는파일_로그확인.docx"])
    assert ei.value.code == 1
    capsys.readouterr()
    text = open(rt.error_log_path(), encoding="utf-8").read()
    assert "명령: reader.py read --file 없는파일_로그확인.docx" in text


def test_version_matches_skill_md():
    import re
    for skill in (os.path.join(TOOL_DIR, "SKILL.md"), os.path.join(DEV, "tools", "office-cli-test", "SKILL.md")):
        with open(skill, encoding="utf-8") as f:
            m = re.search(r'^\s+version:\s*"([^"]+)"', f.read(), re.M)
        assert m and m.group(1) == rt.VERSION, skill


# ------------------------------------------------------------------ 공통 런타임 R3/R4/R5/R7/R8

def test_help_is_stderr_text_and_one_json_line(tmp_path):
    for argv in (["-h"], ["go", "-h"]):                       # 최상위와 서브명령 모두
        code, payload, err = run_script(tmp_path, TOOL_BODY, *argv)
        assert code == 0 and payload["command"] == "help" and payload["verified"] is True
        assert payload["agent_directive"].startswith("VERIFIED:") and "usage:" in err
        assert payload["skill_file"].endswith("SKILL.md") and "읽지 마세요" in payload["agent_directive"]


def test_value_error_is_internal_error_with_traceback(tmp_path):
    body = TOOL_BODY.replace('raise RuntimeError("버그")', 'raise ValueError("코드 버그")')
    code, payload, _ = run_script(tmp_path, body, "go", "--mode", "x")
    assert code == 1 and payload["error_type"] == "internal" and payload["agent_directive"].startswith("ERROR:")
    assert "traceback:" in _error_log(tmp_path)


def test_not_found_recommends_only_same_name():
    e = rt.not_found("시트", "sheet 1", ["Sheet_1", "Sheet2"])          # 표기만 다름 -> 바로 재시도해도 같은 뜻
    assert "[추천] Sheet_1" in str(e) and not e.candidates
    e = rt.not_found("시트", "2024", ["2024_1분기", "2024_2분기"])      # 부분 일치 -> 사용자가 고름
    assert "[추천]" not in str(e) and e.candidates == ["2024_1분기", "2024_2분기"]
    e = rt.not_found("첨부 파일", "견적서.pdf", ["견적서_구버전.pdf"])
    assert "[추천]" not in str(e) and "[후보] 견적서_구버전.pdf" in str(e)
    e = rt.not_found("색상", "yelow", ["yellow", "green"], vocab=True)  # 도구가 정한 옵션 값의 오타
    assert "[추천] yellow" in str(e)
    e = rt.not_found("시트", "없음", [])
    assert "[추천]" not in str(e) and "[후보]" not in str(e)


def test_resolve_output_path_never_overwrites(tmp_path, monkeypatch):
    monkeypatch.setattr(rt, "RECOVERED", [])
    target = tmp_path / "보고서.csv"
    target.write_text("기존", encoding="utf-8")
    out = rt.resolve_output_path(str(target), "x", ".csv")
    assert out != str(target) and os.path.basename(out).startswith("보고서_") and target.read_text(encoding="utf-8") == "기존"
    assert rt.RECOVERED and "보고서.csv 이미 있음" in rt.RECOVERED[0]
    assert rt.ok("t", "c", "p")["recovered"] == rt.RECOVERED
    assert rt.resolve_output_path(str(target), "x", ".csv", overwrite=True) == str(target)


def test_timestamped_does_not_stack_stamps(tmp_path):
    p = str(tmp_path / "r_20261028_101500.xlsx")
    import re
    assert re.fullmatch(r"r_\d{8}_\d{6}\.xlsx", os.path.basename(rt._timestamped(p)))


def test_retry_fs_retries_only_transient_errors(monkeypatch):
    monkeypatch.setattr(rt.time, "sleep", lambda s: None)
    calls = []

    def missing():
        calls.append(1)
        raise FileNotFoundError("없음")
    with pytest.raises(FileNotFoundError):
        rt.retry_fs(missing)
    assert len(calls) == 1                                    # 영구 오류는 기다리지 않음

    state = {"n": 0}

    def locked_twice():
        state["n"] += 1
        if state["n"] < 3:
            raise PermissionError("잠김")
        return "ok"
    assert rt.retry_fs(locked_twice) == "ok" and state["n"] == 3


def test_lock_fallback_reports_recovered(tmp_path, monkeypatch):
    monkeypatch.setattr(rt, "RECOVERED", [])
    monkeypatch.setattr(rt.time, "sleep", lambda s: None)
    target = str(tmp_path / "표.csv")

    def writer(p):
        if p == target:
            raise PermissionError("Excel에서 열림")
        open(p, "w").close()
    out = rt.save_with_lock_fallback(target, writer, "CSV 파일")
    assert out != target and os.path.exists(out)
    assert rt.RECOVERED == [f"표.csv 잠김(CSV 파일) -> {os.path.basename(out)}로 저장"]


def test_read_text_file_normalizes_crlf(tmp_path):
    p = tmp_path / "w.txt"
    p.write_bytes("가\r\n나\r\n".encode("cp949"))
    assert rt.read_text_file(str(p)) == "가\n나\n"


def test_unescape_newlines_keeps_windows_paths():
    assert rt.unescape_newlines(r"## 제목\n- 항목") == "## 제목\n- 항목"
    assert rt.unescape_newlines(r"경로 C:\new\data 확인") == r"경로 C:\new\data 확인"          # \n inside a path
    assert rt.unescape_newlines(r"- C:\new\data\n- 다음") == "- C:\\new\\data\n- 다음"
    assert rt.unescape_newlines(r"저장: C:\a\b.xlsx\n다음") == "저장: C:\\a\\b.xlsx\n다음"


@pytest.mark.parametrize("argv, want", [
    (["excel.py", "list"], "[추천] info"),
    (["excel.py", "raed"], "[추천] read"),
    (["office.py", "undo", "--target", "excel"], "[추천] --app"),
    (["excel.py", "read", "--target", "C:\\a", "b.xlsx"], "큰따옴표"),
])
def test_argument_errors_from_real_logs_get_a_hint(argv, want):
    tool_dir = os.path.dirname(rt.__file__)
    proc = subprocess.run([sys.executable, os.path.join(tool_dir, argv[0]), *argv[1:]], capture_output=True, text=True,
                          encoding="utf-8")
    out = [ln for ln in proc.stdout.splitlines() if ln.strip()]
    assert proc.returncode == 2 and len(out) == 1
    payload = json.loads(out[0])
    assert want in payload["message"] and "[추천]" in payload["agent_directive"]


def test_ok_does_not_add_verified_for_the_caller():
    """R1: ok()는 VERIFIED:를 대신 붙이지 않음. 확인하지 못한 결과는 ok_unverified."""
    assert rt.ok("t", "c", "p")["agent_directive"].startswith("VERIFIED:") and rt.ok("t", "c", "p")["verified"] is True
    with pytest.raises(ValueError):
        rt.ok("t", "c", "p", "완료")
