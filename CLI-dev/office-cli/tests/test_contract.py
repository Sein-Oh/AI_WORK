"""ACI 계약과 공통 런타임(_runtime.py): stdout 1줄 JSON, exit 코드, agent_directive, 공통 도우미,
출력 상한, 오류 로그와 가림(문서 글·메일 계정을 남기지 않음), 데이터 폴더.
"""
import datetime
import json
import os
import shutil
import subprocess
import sys
import textwrap
import zipfile

import pytest

import _office as off
import _runtime as rt
import excel
import office
import outlook
import ppt
import reader
import word
from fake_excel import FakeChartObject, FakeExcelWorkbook, FakeXlApp
from fake_office import make_win32com
from fake_word import FakeWordDocument
from office_helpers import assert_not_logged, error_log_text, recommended_argv, run_as_tool, run_main, SECRET, split_command, TOOL_DIR
from sample_docs import new_docx, new_pptx


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


def _error_log(tmp_path):
    p = tmp_path / "data" / "logs" / "errors.txt"
    return p.read_text(encoding="utf-8") if p.exists() else ""


def _log_text(isolated_dirs):
    with open(os.path.join(isolated_dirs["data"], "logs", "errors.txt"), encoding="utf-8") as f:
        return f.read()


@pytest.fixture
def xs(monkeypatch):
    """시트 이름과 칸 값에 표지가 든 가짜 Excel."""
    app = FakeXlApp()
    wb = FakeExcelWorkbook(app, "보고.xlsx", sheets=("ZQX인사", "ZQX급여"))
    fake, _ = make_win32com()
    monkeypatch.setattr(off, "win32com", fake)
    monkeypatch.setattr(off, "pythoncom", None)
    monkeypatch.setattr(off, "list_open", lambda kind: [])
    monkeypatch.setattr(off, "get_active", lambda kind: {"name": wb.Name, "path": wb.FullName, "app": app, "doc": wb})
    monkeypatch.setattr(off, "save_snapshot", lambda *a, **k: None)
    ws = wb.Worksheets("ZQX인사")
    ws.Range("A1").Value = "ZQX홍길동"
    ws.Range("B1").Value = "ZQX연봉"
    ws.Range("D1").Value = "ZQX주소"
    return app, wb


PROBE = ("import sys, json\n"
         "sys.path.insert(0, sys.argv[1])\n"
         "import _runtime as rt\n"
         "print(json.dumps({'config': rt.CONFIG_FILE, 'data': rt.get_data_dir(), 'cache': rt.get_cache_dir()}))\n")


def _env(tmp_path, **extra):
    env = {k: v for k, v in os.environ.items() if not k.startswith("OFFICE_CLI_") and k != "ENGENIUS_OUT"}
    env["USERPROFILE"] = str(tmp_path / "home")
    env.update(extra)
    return env


def _unpacked(tmp_path):
    """트레이처럼 도구를 %TEMP%\\engenius-skill\\<id>\\에 푼 자리."""
    d = tmp_path / "engenius-skill" / "r1"
    if not d.exists():
        shutil.copytree(TOOL_DIR, d, ignore=shutil.ignore_patterns("data", "cache", "config.json", "__pycache__"))
    return str(d)


def _probe(tmp_path, tool_dir=TOOL_DIR, **extra):
    out = subprocess.run([sys.executable, "-c", PROBE, tool_dir], capture_output=True, timeout=60,
                         env=_env(tmp_path, **extra))
    return json.loads(out.stdout.decode("utf-8").strip().splitlines()[-1])


# ---------------------------------------------------------------------------
# stdout 1줄 JSON, exit 코드, agent_directive
# ---------------------------------------------------------------------------


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
# 공통 도우미: 이름 추천, JSON 변환, 파일 읽기·쓰기, 경로, 설정, 출력 상한
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
    clipped, cut, cells, dropped = rt.clamp_rows(rows, max_rows=50)
    assert len(clipped) == 50 and cut and cells == 0 and dropped == 0
    clipped, cut, cells, dropped = rt.clamp_rows(rows[:5])
    assert len(clipped) == 5 and not cut and cells == 0 and dropped == 0
    text, cut = rt.clamp_text("가" * 10, max_chars=4)
    assert text == "가가가가" and cut


def test_next_command_quotes_paths():
    cmd = rt.next_command("office.py", ["transfer", "--chart", "사인 2주기", "--to-app", "word,ppt"], "차트 전달")
    assert cmd["purpose"] == "차트 전달"
    assert "'사인 2주기'" in cmd["command"] and " --to-app word,ppt" in cmd["command"]   # PowerShell 작은따옴표 (1.4.76)
    assert split_command(cmd["command"])[2:] == ["transfer", "--chart", "사인 2주기", "--to-app", "word,ppt"]
    assert "office.py" in cmd["command"]


# ---------------------------------------------------------------------------
# 공통 런타임: help, 내부 오류, 같은 이름만 추천, 덮어쓰지 않는 저장 경로, 재시도, 잠김 복구, 줄바꿈, 인자 오류 힌트
# ---------------------------------------------------------------------------


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


def test_remove_file_waits_for_transient_lock(tmp_path, monkeypatch):
    """1.4.81: 스냅샷·캐시·임시 파일 지우기는 백신이 잠깐 잡아도 다시 시도 (전에는 os.remove 한 번, 실패하면 파일이 남음)."""
    monkeypatch.setattr(rt.time, "sleep", lambda s: None)
    f = tmp_path / "snap.tmp"
    f.write_text("x", encoding="utf-8")
    real, state = os.remove, {"n": 0}

    def flaky(path):
        state["n"] += 1
        if state["n"] < 3:
            raise PermissionError("잠김")
        real(path)
    monkeypatch.setattr(rt.os, "remove", flaky)
    assert rt.remove_file(str(f)) is True and not f.exists() and state["n"] == 3
    assert rt.remove_file(str(f)) is False                     # 없으면 그냥 넘어감 (오류 없음)
    monkeypatch.setattr(rt.os, "remove", lambda p: (_ for _ in ()).throw(PermissionError("잠김")))
    assert rt.remove_file(str(tmp_path / "x")) is False         # 한도를 넘겨도 오류를 올리지 않음


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


# want가 목록이면 [추천] 전체 명령의 인자가 그것과 같아야 함 (1.4.79: [추천]은 전체 명령)
@pytest.mark.parametrize("argv, want", [
    (["excel.py", "list"], ["info"]),
    (["excel.py", "raed"], ["read"]),
    (["office.py", "undo", "--target", "excel"], "undo --app excel"),
    (["excel.py", "read", "--target", "C:\\a", "b.xlsx"], "큰따옴표"),
    (["excel.py", "write", "--cell", "K1", "--value", "3"], "write --range K1 --value 3"),          # 1.4.39
])
def test_argument_errors_from_real_logs_get_a_hint(argv, want):
    tool_dir = os.path.dirname(rt.__file__)
    proc = subprocess.run([sys.executable, os.path.join(tool_dir, argv[0]), *argv[1:]], capture_output=True, text=True,
                          encoding="utf-8")
    out = [ln for ln in proc.stdout.splitlines() if ln.strip()]
    assert proc.returncode == 2 and len(out) == 1
    payload = json.loads(out[0])
    assert "[추천]" in payload["agent_directive"]
    if isinstance(want, list):
        assert recommended_argv(payload["message"], argv[0][:-3]) == want, payload
    else:
        assert want in payload["message"], payload


@pytest.mark.parametrize("argv, fixed", [
    (["excel.py", "check", "--file", "a.xlsx", "--range", "B3:C7", "--item-col", "B", "--value-col", "C", "--spec", "s.txt"],
     ["check", "--target", "a.xlsx", "--range", "B3:C7", "--item-col", "B", "--value-col", "C", "--spec", "s.txt"]),   # Gemma O1
    (["word.py", "write", "--target", "a.docx", "--text", "다음", "--append"],
     ["write", "--target", "a.docx", "--text", "다음", "--pos", "end"]),                                                 # Gemma O4
    (["reader.py", "read", "--target", "a.pdf"], ["read", "--file", "a.pdf"]),              # 필수 --file 누락 오류에서도
    (["excel.py", "read", "--target", "a.xlsx", "--text", "전압"], ["read", "--target", "a.xlsx", "--search", "전압"]),
    (["excel.py", "read", "--target", "a.xlsx", "--encoding", "utf-8", "--json"], ["read", "--target", "a.xlsx"]),
    (["outlook.py", "search", "--query", "견적", "--max", "5"], ["search", "--query", "견적", "--limit", "5"]),
])
def test_same_meaning_option_gets_full_fixed_command(argv, fixed):
    """Gemma 기록 감사 (1.4.77): 같은 뜻의 옵션은 옵션 이름이 아니라 고친 전체 명령을 [추천] (그대로 1번 실행하면 됨)."""
    tool_dir = os.path.dirname(rt.__file__)
    proc = subprocess.run([sys.executable, os.path.join(tool_dir, argv[0]), *argv[1:]], capture_output=True, text=True,
                          encoding="utf-8")
    payload = json.loads([ln for ln in proc.stdout.splitlines() if ln.strip()][-1])
    assert proc.returncode == 2 and "[추천] python" in payload["message"], payload
    got = split_command(payload["message"].split("[추천] ", 1)[1])
    assert got[2:] == fixed, got


@pytest.mark.parametrize("argv, never", [
    (["reader.py", "clear", "--cell", "A1"], "--all"),          # 철자가 비슷한 --all: 캐시 전부 삭제였음
    (["outlook.py", "reply", "--id", "1", "--cell", "x"], "--all"),   # 전체 답장
    (["outlook.py", "new", "--to", "a", "--at", "b", "--body", "x"], "[추천]"),   # --at -> --to(받는 사람)였음
    (["office.py", "undo", "--append"], "[추천]"),               # --append -> --app이었음
    (["excel.py", "info", "--target", "a.xlsx", "--text", "x"], "[추천]"),       # --text -> --target이었음
])
def test_lookalike_option_is_never_recommended(argv, never):
    tool_dir = os.path.dirname(rt.__file__)
    proc = subprocess.run([sys.executable, os.path.join(tool_dir, argv[0]), *argv[1:]], capture_output=True, text=True,
                          encoding="utf-8")
    payload = json.loads([ln for ln in proc.stdout.splitlines() if ln.strip()][-1])
    assert proc.returncode == 2 and never not in payload["message"].split("CLI 인자 오류:", 1)[1].split("옵션:", 1)[0], payload
    if never == "--all":
        assert "[추천]" not in payload["message"] and "[후보] --all" not in payload["message"]


@pytest.mark.parametrize("argv, want", [
    (["outlook.py", "sent", "--id", "3"], "send"),
    (["outlook.py", "sned", "--id", "3"], "send"),
    (["outlook.py", "end", "--id", "3"], "send"),
    (["reader.py", "clea", "--all"], "clear"),
    (["office.py", "redo"], "undo"),
])
def test_destructive_command_typo_is_candidate_not_recommendation(argv, want):
    """11차 검토: 오타가 되돌리기 어려운 명령(발송·캐시 지우기·되돌리기)과 비슷하면 [추천](묻지 않고 실행)이 아니라 [후보](사용자에게 물음)."""
    tool_dir = os.path.dirname(rt.__file__)
    proc = subprocess.run([sys.executable, os.path.join(tool_dir, argv[0]), *argv[1:]], capture_output=True, text=True,
                          encoding="utf-8")
    payload = json.loads([ln for ln in proc.stdout.splitlines() if ln.strip()][-1])
    assert proc.returncode == 2 and "[추천]" not in payload["message"], payload
    assert f"[후보] 1) {want}" in payload["message"], payload
    cmd = payload["confirm_commands"][0]                    # 고른 번호의 전체 명령, 되돌리기 어려운 명령 표시
    assert cmd["command"].startswith("python ") and f" {want}" in cmd["command"] and cmd["destructive"] is True, payload


def test_destructive_flag_marks_only_overwrite_delete_send():
    assert rt.next_command("outlook.py", ["send", "--id", "3"], "보내기", destructive=True)["destructive"] is True
    assert "destructive" not in rt.next_command("excel.py", ["read"], "읽기")


def test_unsafe_next_commands_checker():
    """conftest의 next_commands 점검(모든 테스트에 자동 적용)이 문서를 바꾸는 명령을 잡는지."""
    from office_helpers import unsafe_next_commands   # conftest 이름은 ml-kit tests와 겹쳐 함께 돌 때 다른 모듈을 집음
    nxt = lambda *cmds: {"status": "ok", "next_commands": [{"command": f'python "C:\\CLI\\{c}"'} for c in cmds]}
    assert unsafe_next_commands(nxt('word.py" write --file "a.md"', 'office.py" transfer --from-app excel'))
    assert unsafe_next_commands(nxt('excel.py" chart --target active --range A1:B3'))
    assert not unsafe_next_commands(nxt('reader.py" read --file "a.pdf"', 'excel.py" comment --action add --dry-run',
                                        'excel.py" chart --action list', 'office.py" undo --list'))
    assert not unsafe_next_commands({"continue": True, **nxt('excel.py" write --cells "A1=1"')})


@pytest.mark.parametrize("argv, want", [
    (["office.py", "write", "--target", "보고서.docx", "--text", "요약"], 'word.py write --text'),
    (["office.py", "transfer", "--to-target", "보고서.docx", "--text", "요약"], "글은 받지 않습니다"),
])
def test_other_file_command_is_named_not_recommended(argv, want):
    """1.4.39 (회사 PC 오류 로그): 다른 파일의 명령이 맞는 경우 [추천](값만 바꿔 1회 재시도)이 아니라 맞는 명령을 안내.
    [추천]을 따라 옵션 이름만 바꾸면 또 틀림 (예: office.py transfer --text -> --theme)."""
    tool_dir = os.path.dirname(rt.__file__)
    proc = subprocess.run([sys.executable, os.path.join(tool_dir, argv[0]), *argv[1:]], capture_output=True, text=True,
                          encoding="utf-8")
    payload = json.loads([ln for ln in proc.stdout.splitlines() if ln.strip()][0])
    assert proc.returncode == 2 and want in payload["message"]
    assert "[추천]" not in payload["message"] and "[후보]" not in payload["message"]


def test_ok_does_not_add_verified_for_the_caller():
    """R1: ok()는 VERIFIED:를 대신 붙이지 않음. 확인하지 못한 결과는 ok_unverified."""
    assert rt.ok("t", "c", "p")["agent_directive"].startswith("VERIFIED:") and rt.ok("t", "c", "p")["verified"] is True
    with pytest.raises(ValueError):
        rt.ok("t", "c", "p", "완료")


# ---------------------------------------------------------------------------
# ok(verified=False)는 VERIFIED가 아님, ok_file은 형식대로 다시 열어 확인, 오류 로그는 내용 옵션 값을 가림
# ---------------------------------------------------------------------------


def test_ok_refuses_verified_false():
    with pytest.raises(ValueError):
        rt.ok("t", "c", "p", "VERIFIED: x", verified=False)


def test_ok_file_reopens_by_format(tmp_path):
    import openpyxl
    good = tmp_path / "a.xlsx"
    wb = openpyxl.Workbook()
    wb.active["A1"] = 1
    wb.save(str(good))
    assert rt.ok_file("t", "c", "p", str(good), "VERIFIED: ok")["agent_directive"].startswith("VERIFIED:")
    bad = tmp_path / "b.xlsx"
    bad.write_bytes(b"not a zip file")
    out = rt.ok_file("t", "c", "p", str(bad), "VERIFIED: ok")
    assert out["agent_directive"].startswith("UNVERIFIED:") and out["verified"] is False and "다시 열지 못함" in out["unverified"][0]
    for name in ("b.docx", "b.pptx"):
        (tmp_path / name).write_bytes(b"broken")
        assert rt.ok_file("t", "c", "p", str(tmp_path / name), "VERIFIED: ok")["verified"] is False
    csv = tmp_path / "r.csv"
    csv.write_text("a,b\n1,2\n", encoding="utf-8")
    assert rt.ok_file("t", "c", "p", str(csv), "VERIFIED: ok", expected_rows=2)["verified"] is True
    out = rt.ok_file("t", "c", "p", str(csv), "VERIFIED: ok", expected_rows=5)
    assert out["verified"] is False and "행 수" in out["unverified"][0]
    md = tmp_path / "n.md"
    md.write_text("# 제목", encoding="utf-8")
    assert rt.ok_file("t", "c", "p", str(md), "VERIFIED: ok")["verified"] is True
    other = tmp_path / "x.bin"
    other.write_bytes(b"1234")
    assert rt.ok_file("t", "c", "p", str(other), "VERIFIED: ok")["verified"] is False


def test_error_log_redacts_content_options(isolated_dirs):
    argv = ["new", "--to", "kim@company.com", "--subject", "인사 평가", "--body", "비밀 본문입니다",
            "--attach", "C:\\docs\\a.pdf", "--formula", "=A1*2", "=B1", "--text=메모 내용"]
    rt.write_error_log("outlook", argv, {"error_type": "input", "message": "m"})
    text = open(rt.error_log_path(), encoding="utf-8").read()
    for secret in ("kim@company.com", "인사 평가", "비밀 본문입니다", "=A1*2", "=B1", "메모 내용"):
        assert secret not in text
    assert "--to <15자 생략>" in text and "--body <8자 생략>" in text and "--text=<5자 생략>" in text
    assert "C:\\docs\\a.pdf" in text                                     # 경로는 남김
    assert rt.redact_argv(["transfer", "--from", "excel", "--to", "word"], "office") == \
        ["transfer", "--from", "excel", "--to", "word"]        # office.py에 잘못 쓴 옛 이름 --from/--to도 앱 이름이라 가리지 않음 (1.4.79)


# ---------------------------------------------------------------------------
# excel/word/ppt read --save 저장 파일을 다시 열어 확인 (빈 행은 세지 않음)
# ---------------------------------------------------------------------------


def test_word_and_ppt_read_save_is_reopened(tmp_path, capsys, monkeypatch):
    d = tmp_path / "문서.docx"
    doc = new_docx()
    doc.add_paragraph("첫 단락")
    doc.save(str(d))
    s = tmp_path / "발표.pptx"
    prs = new_pptx()
    prs.slides.add_slide(prs.slide_layouts[0]).shapes.title.text = "제목"
    prs.save(str(s))
    for module, target in ((word, d), (ppt, s)):
        code, out = run_main(module, capsys, "read", "--target", str(target), "--save", str(tmp_path / f"{target.stem}.md"))
        assert code == 0 and out["agent_directive"].startswith("VERIFIED: 저장한 파일을 다시 열어"), out
    monkeypatch.setattr(rt, "check_saved_file", lambda path, rows=None, delimiter=None: "저장한 파일을 다시 열지 못함 (테스트)")
    for module, target in ((word, d), (ppt, s)):
        code, out = run_main(module, capsys, "read", "--target", str(target), "--save", str(tmp_path / f"{target.stem}2.md"))
        assert code == 0 and out["agent_directive"].startswith("UNVERIFIED:"), out


# ---------------------------------------------------------------------------
# pptm은 python-pptx로 다시 열어 확인
# ---------------------------------------------------------------------------


def test_pptm_is_reopened(tmp_path):
    src = tmp_path / "a.pptx"
    new_pptx().save(str(src))
    dst = tmp_path / "a.pptm"
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w") as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename == "[Content_Types].xml":
                data = data.replace(b"application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
                                    b"application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml")
            zout.writestr(info, data)
    assert rt.check_saved_file(str(dst)) is None
    (tmp_path / "b.pptm").write_bytes(b"broken")
    assert "다시 열지 못함" in rt.check_saved_file(str(tmp_path / "b.pptm"))


# ---------------------------------------------------------------------------
# 출력 크기 상한 (긴 칸, 찾은 값, 슬라이드 글, 워드 단락, 노트)
# ---------------------------------------------------------------------------


def test_clamp_rows_caps_cells_and_first_row():
    rows = [["가" * 30000 for _ in range(40)], [1, 2]]
    shown, rows_cut, cells_cut, cols_dropped = rt.clamp_rows(rows)
    assert sum(len(str(v)) for r in shown for v in r) <= rt.DEFAULT_MAX_CHARS + 100
    assert rows_cut and cells_cut > 0 and cols_dropped > 0
    assert cells_cut + cols_dropped == 40                 # 긴 값(보인 칸)과 뺀 열을 따로 셈
    shown, rows_cut, cells_cut, cols_dropped = rt.clamp_rows([["짧은 값", 1], ["둘", 2]])
    assert shown == [["짧은 값", 1], ["둘", 2]] and not rows_cut and cells_cut == 0 and cols_dropped == 0


def test_ppt_and_word_long_text_marked_truncated(tmp_path, capsys):
    s = tmp_path / "deck.pptx"
    prs = new_pptx()
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "제목"
    slide.placeholders[1].text = "라" * 20000
    prs.save(str(s))
    code, out = run_main(ppt, capsys, "read", "--target", str(s))
    sl = out["slides"][0]
    assert code == 0 and out["truncated"] is True and sl["text_truncated"] and len(sl["text"]) == rt.DEFAULT_MAX_CHARS
    assert out["next_commands"][-1]["command"].endswith("--save")
    d = tmp_path / "doc.docx"
    doc = new_docx()
    doc.add_paragraph("마" * 20000)
    doc.save(str(d))
    code, out = run_main(word, capsys, "read", "--target", str(d))
    para = out["paragraphs"][0]
    assert code == 0 and out["truncated"] is True and para["text_truncated"] and para["text_chars"] == 20000
    assert out["next_commands"][0]["command"].endswith("--save")


# ---------------------------------------------------------------------------
# 기본 저장 이름이 같은 초에 겹치면 번호를 붙인 새 이름 (덮어쓰지 않음, recovered로 알림)
# ---------------------------------------------------------------------------


def test_default_save_name_in_same_second_is_unique(monkeypatch):
    monkeypatch.setattr(rt, "timestamp", lambda: "20261102_101500")
    rt.RECOVERED.clear()
    first = rt.resolve_output_path(None, "표", ".csv")
    with open(first, "w", encoding="utf-8") as f:
        f.write("a\n")
    second = rt.resolve_output_path(None, "표", ".csv")
    assert first != second and not os.path.exists(second)
    assert os.path.basename(second) == "표_20261102_101500_2.csv"
    assert any("같은 초" in r for r in rt.RECOVERED)
    rt.RECOVERED.clear()


# ---------------------------------------------------------------------------
# 기본 파일 이름
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("stem", ["chart_1/4분기 매출_redraw", 'a?b*c:d"e<f>g|h\\i', "CON", "끝점. "])
def test_default_stem_is_sanitized(stem):
    path = rt.resolve_output_path(None, stem, ".png")
    assert os.path.dirname(path) == rt.get_work_dir()
    base = os.path.basename(path)
    assert not any(ch in base for ch in '/\\:*?"<>|') and not base.upper().startswith("CON_")
    with open(path, "wb") as f:                # 실제로 만들 수 있는 경로
        f.write(b"x")


# ---------------------------------------------------------------------------
# 오류 로그 (<데이터 폴더>/logs/errors.txt)
# ---------------------------------------------------------------------------


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


# ---------------------------------------------------------------------------
# 오류 로그 가리기 - 짧은 이름(-q), 줄여 쓴 이름(--que), --map. 파서는 줄임을 받지 않음
# ---------------------------------------------------------------------------


def test_redact_short_abbrev_and_map():
    red = rt.redact_argv(["search", "-q", "비밀", "-q비밀이", "--que", "기밀", "--bod=본문", "--map", '{"a":"b"}',
                          "--file", "C:\\a.txt", "-n", "5"], "reader")
    text = " ".join(red)
    assert "비밀" not in text and "기밀" not in text and "본문" not in text and '"a"' not in text
    assert "C:\\a.txt" in red and red[-2:] == ["-n", "5"]


def test_short_aliases_of_content_options_are_all_mapped():
    import argparse
    for module in (excel, word, ppt, outlook, reader):
        stack = [module.build_cli()]
        while stack:
            p = stack.pop()
            for a in p._actions:
                if isinstance(a, argparse._SubParsersAction):
                    stack += list(a.choices.values())
                    continue
                longs = [o for o in a.option_strings if o.startswith("--")]
                if any(o in rt.REDACT_OPTIONS for o in longs):
                    for short in (o for o in a.option_strings if not o.startswith("--")):
                        assert rt.REDACT_SHORT.get(short) in longs, (module.__name__, short)
            assert p.allow_abbrev is False, module.__name__


def test_abbreviated_option_is_rejected_and_log_is_redacted(capsys):
    code, out = run_main(reader, capsys, "search", "--que", "인사평가 기밀", "--file", "a.txt")
    assert code == 2 and out["status"] == "error"
    log = open(rt.error_log_path(), encoding="utf-8").read()
    assert "인사평가" not in log and "--que" in log


# ---------------------------------------------------------------------------
# 오류 로그 가리기
# ---------------------------------------------------------------------------


def test_redact_section_and_excel_highlight():
    red = rt.redact_argv(["read", "--file", "C:/a/보고.docx", "--section", "인수 계획"], "word")
    assert red == ["read", "--file", "C:/a/보고.docx", "--section", "<5자 생략>"], red
    red = rt.redact_argv(["search", "-q", "x", "--section=기밀"], "reader")
    assert red[-1] == "--section=<2자 생략>", red
    red = rt.redact_argv(["style", "--target", "보고.xlsx", "--sheet", "Sheet1", "--highlight", "=완료:lightgreen"], "excel")
    assert red == ["style", "--target", "보고.xlsx", "--sheet", "Sheet1", "--highlight", "<14자 생략>"], red
    assert rt.redact_argv(["format", "--highlight", "yellow"], "word") == ["format", "--highlight", "yellow"]


@pytest.mark.parametrize("mod,key", [("word", "documents"), ("ppt", "presentations")])
def test_word_ppt_info_all_zero(monkeypatch, capsys, mod, key):
    import ppt
    import word
    monkeypatch.setattr(off, "list_open", lambda kind: [])
    code, out = run_main({"word": word, "ppt": ppt}[mod], capsys, "info", "--all")
    assert code == 0 and out["count"] == 0 and out[key] == [] and out["agent_directive"].startswith("VERIFIED: 0건"), out


# ---------------------------------------------------------------------------
# 오류 로그 message 가리기
# ---------------------------------------------------------------------------


def test_log_masks_anchor_and_cuts_similar_text(wd, capsys, isolated_dirs, monkeypatch):
    argv = ["write", "--anchor", "2026년 비밀계획", "--text", "본문"]
    monkeypatch.setattr(sys, "argv", ["word.py"] + argv)
    monkeypatch.setattr(rt, "CURRENT_ARGV", None)
    code, out = run_main(word, capsys, *argv)
    assert code == 1 and "2026년 비밀계획" in out["message"] and "비슷한 문구: '2026년 계획입'" in out["message"], out
    log = _log_text(isolated_dirs)
    assert "비밀계획" not in log and "계획입" not in log and "<10자 생략>" in log, log
    assert "비슷한 문구 <문서 내용 생략>" in log, log            # 1.4.28: 오류가 문서 글을 뺀 log_message를 줌


def test_log_masks_excel_highlight_rule(xl, capsys, isolated_dirs, monkeypatch):
    argv = ["style", "--range", "A1:B2", "--highlight", "기밀완료"]
    monkeypatch.setattr(sys, "argv", ["excel.py"] + argv)
    monkeypatch.setattr(rt, "CURRENT_ARGV", None)
    code, out = run_main(excel, capsys, *argv)
    assert code != 0 and "입력: 기밀완료" in out["message"], out
    log = _log_text(isolated_dirs)
    assert "기밀완료" not in log and "입력: <4자 생략>" in log, log


def test_log_cuts_candidates(isolated_dirs):
    msg = "제목 '결론'을(를) 찾을 수 없습니다. 비슷한 이름이 있지만 같은 것인지 도구가 정할 수 없습니다. [후보] 1.1 인수 계획, 1.2 비밀 일정"
    rt.write_error_log("word", ["read", "--section", "결론"], {"error_type": "input", "message": msg})
    log = _log_text(isolated_dirs)
    assert "인수 계획" not in log and "비밀 일정" not in log and "[후보]" not in log, log
    assert "제목 '<2자 생략>'을(를) 찾을 수 없습니다." in log, log
    msg2 = "제목 'x'을(를) 찾을 수 없습니다. 사용 가능한 제목: 1. 개요, 2. 비밀"
    assert "비밀" not in rt.log_message(msg2, [], "word")
    assert rt.log_message("시트 'a' [추천] 기밀시트", [], "excel").endswith("자 생략>")


def test_log_message_masks_equals_and_short_forms():
    m = rt.log_message("검색어 '견적서' / 절 '기밀'", ["search", "-q견적서", "--section=기밀"], "reader")
    assert "견적서" not in m and "기밀" not in m, m


def test_log_never_raises(isolated_dirs):
    class Bad:
        def __str__(self):
            raise RuntimeError("x")
    assert rt.log_message(Bad(), ["--text", "a"], "word").startswith("<message 생략")
    rt.write_error_log("word", ["write", "--text", "a"], {"error_type": "input", "message": Bad()})
    assert "message 생략" in _log_text(isolated_dirs)


# ---------------------------------------------------------------------------
# 오류 로그에 문서 글을 남기지 않음 (기존 제목·노트, 찾을 글 미리보기, 덮어쓸 칸 값, 댓글 CSV 후보)
# ---------------------------------------------------------------------------


def test_ppt_existing_title_not_logged(pp, capsys, isolated_dirs, monkeypatch):
    pp["prs"].Slides(1).Shapes.Title.TextFrame.TextRange.Text = "인수합병 비밀 일정"
    code, out = run_as_tool(ppt, capsys, monkeypatch, "ppt", "write", "--slide", "1", "--title", "새 제목")
    assert code == 1 and out["error_type"] == "needs_user_decision" and "비밀 일정" in out["message"], out
    log = error_log_text(isolated_dirs)
    assert "비밀 일정" not in log and "슬라이드 1에 이미 제목이 있습니다" in log, log


def test_ppt_existing_notes_not_logged(pp, capsys, isolated_dirs, monkeypatch):
    pp["prs"].Slides(1).NotesPage.body.TextFrame.TextRange.Text = "기밀 발표 메모 내용"
    code, out = run_as_tool(ppt, capsys, monkeypatch, "ppt", "notes", "--slide", "1", "--text", "새 노트")
    assert code == 1 and "기밀 발표" in out["message"], out
    assert "기밀 발표" not in error_log_text(isolated_dirs)


def test_word_long_find_no_prefix_preview(wd, capsys, isolated_dirs, monkeypatch):
    longv = "고객사 홍길동 계약 해지 예정 문서 " * 20
    code, out = run_as_tool(word, capsys, monkeypatch, "word", "replace", "--find", longv, "--replace", "x")
    assert code == 1 and "홍길동" not in out["message"] and "1번째 찾을 문구" in out["message"], out
    assert "홍길동" not in error_log_text(isolated_dirs)


def test_short_value_does_not_overmask():
    m = rt.log_message("시트 'Sheet1'의 A1:B10 범위가 비어 있습니다", ["write", "--value", "1"], "excel")
    assert "Sheet1" in m and "A1:B10" in m, m
    # 따옴표 안의 짧은 값은 가림
    assert rt.log_message("값 '1'을 넣지 못했습니다", ["write", "--value", "1"], "excel") == "값 '<1자 생략>'을 넣지 못했습니다"


def test_excel_overwrite_cell_values_not_logged(xl, capsys, isolated_dirs, monkeypatch):
    app, wb = xl
    wb.Worksheets("Sheet1").Range("A1").Value = "홍길동 연봉 9800만원"
    code, out = run_as_tool(excel, capsys, monkeypatch, "excel", "write", "--range", "A1", "--value", "x")
    assert code == 1 and "9800" in out["message"], out
    log = error_log_text(isolated_dirs)
    assert "9800" not in log and "A1" in log, log


def test_excel_merge_values_not_logged(xl, capsys, isolated_dirs, monkeypatch):
    app, wb = xl
    ws = wb.Worksheets("Sheet1")
    ws.Range("A1").Value = "a"
    ws.Range("B1").Value = "기밀 매출 12억"
    code, out = run_as_tool(excel, capsys, monkeypatch, "excel", "range", "--action", "merge", "--range", "A1:C1")
    assert code == 1 and "12억" in out["message"], out
    assert "12억" not in error_log_text(isolated_dirs)


def test_word_comment_csv_candidates_not_logged(wd, capsys, isolated_dirs, monkeypatch, tmp_path):
    wd["doc"] = FakeWordDocument("계획 홍길동 해고 예정\r계획 김철수 해고 예정\r일정")
    f = tmp_path / "memo.csv"
    f.write_text("기준 글,메모 내용\n계획,확인\n일정,확인2\n", encoding="utf-8")
    code, out = run_as_tool(word, capsys, monkeypatch, "word", "comment", "--action", "add", "--file", str(f))
    assert code == 1 and "홍길동" in out["message"], out
    log = error_log_text(isolated_dirs)
    assert "홍길동" not in log and "김철수" not in log and "줄: 1" in log, log


# ---------------------------------------------------------------------------
# 가림 규칙: 잘림 표식 종류, 숨긴 값의 잘린 앞부분, traceback, 오류의 log_message
# ---------------------------------------------------------------------------


def test_cut_markers_kinds_with_spaces_and_unbracketed():
    for msg in ("메일 폴더 'x'을(를) 찾을 수 없습니다. 사용 가능한 메일 폴더: ZQX받은편지함",
                "문서 'a'을(를) 찾을 수 없습니다. 사용 가능한 열린 Word 문서: ZQX.docx",
                "1번째 줄: 고를 수 없습니다. 후보: 1: ZQX단락",
                "표기만 다릅니다. 추천: ZQX시트"):
        m = rt.log_message(msg, [], "word")
        assert SECRET not in m and "생략>" in m, m


def test_truncated_prefix_of_hidden_value_masked():
    v = "고객사 홍길동 계약 해지 예정 문서"
    m = rt.log_message(f"찾을 문구가 깁니다: '{v[:12]}...'", ["replace", "--find", v], "word")
    assert "홍길동" not in m and "<12자 생략>" in m, m
    # 10자보다 짧은 앞부분은 가리지 않음 (관계없는 글을 가리지 않게)
    assert "고객사" in rt.log_message("고객사 목록", ["replace", "--find", v], "word")


def test_traceback_masked_without_marker_cut(isolated_dirs):
    trace = ('Traceback (most recent call last):\n  File "word.py", line 1, in f\nKeyError: \'비밀계획서 초안\'\n'
             "  [후보] 줄은 코드 글이라 자르지 않음\n")
    rt.write_error_log("word", ["write", "--text", "비밀계획서 초안"], {"error_type": "internal", "message": "KeyError"}, trace)
    log = error_log_text(isolated_dirs)
    assert "비밀계획서" not in log and "KeyError: '<8자 생략>'" in log and "[후보] 줄은 코드 글" in log, log


def test_internal_error_traceback_masked_through_run_cli(isolated_dirs, monkeypatch, capsys):
    monkeypatch.setattr(rt, "CURRENT_ARGV", ["write", "--text", "비밀계획서 초안"])

    def boom():
        raise KeyError("비밀계획서 초안")
    with pytest.raises(SystemExit):
        rt.run_cli("word", "write", boom)
    out = json.loads(capsys.readouterr().out)
    assert "비밀계획서" in out["message"]                                  # stdout은 그대로
    log = error_log_text(isolated_dirs)
    assert "traceback:" in log and "비밀계획서" not in log, log


def test_result_dict_log_message_key_popped(isolated_dirs, monkeypatch, capsys):
    monkeypatch.setattr(rt, "CURRENT_ARGV", ["send"])
    with pytest.raises(SystemExit):
        rt.run_cli("outlook", "send", lambda: {"status": "error", "tool": "outlook", "command": "send", "error_type": "input",
                                               "message": "수신자 확인 실패: ZQX홍길동", "agent_directive": "STOP_RETRY: x",
                                               rt.LOG_MESSAGE_KEY: "수신자 1개 확인 실패"})
    out = json.loads(capsys.readouterr().out)
    assert rt.LOG_MESSAGE_KEY not in out and SECRET in out["message"], out
    log = error_log_text(isolated_dirs)
    assert SECRET not in log and "수신자 1개 확인 실패" in log, log


def test_error_classes_carry_log_message():
    e = rt.ambiguous("'x'이(가) 2곳에 있습니다.", ["ZQX a", "ZQX b"])
    assert SECRET in str(e) and SECRET not in e.log_message and "[후보 2개 생략]" in e.log_message
    e = rt.UserDecisionError("시트가 없습니다.", "묻기", candidates=["ZQX"])
    assert SECRET in str(e) and SECRET not in e.log_message
    for cands in (["ZQX가", "ZQX나"], ["zqx 시트"], ["ZQX"]):
        e = rt.not_found("시트", "ZQX 시트" if cands == ["zqx 시트"] else "없음", cands)
        assert SECRET.lower() in str(e).lower()
        assert SECRET.lower() not in e.log_message.lower().replace("'zqx 시트'", ""), e.log_message
    assert e.log_message.endswith("[사용 가능한 이름 1개 생략]")


# ---------------------------------------------------------------------------
# 가드: 문서 글이 오류 로그에 남지 않음 (needs_user_decision / not_found / ambiguous 경로마다)
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("argv", [
    ["write", "--range", "A1", "--value", "x"],                                # 덮어쓰기 칸 값
    ["range", "--action", "merge", "--range", "A1:C1"],                         # 병합으로 지워질 값
    ["range", "--action", "move", "--range", "A1", "--dest", "D1"],             # 옮길 자리의 값
    ["write", "--sheet", "없는시트", "--range", "A1", "--value", "x"],             # (있는 시트: ...)
    ["write", "--sheet", "zqx 인사", "--range", "A1", "--value", "x"],            # 표기만 다른 시트 [추천]
    ["sheet", "--action", "select", "--name", "없는시트"],                         # 사용 가능한 시트:
    ["sheet", "--action", "select", "--name", "인사"],                             # 비슷한 시트 [후보]
])
def test_guard_excel(xs, capsys, isolated_dirs, monkeypatch, argv):
    code, out = run_as_tool(excel, capsys, monkeypatch, "excel", *argv)
    assert code == 1, out
    assert_not_logged(out, isolated_dirs)


def test_guard_excel_chart_names(xs, capsys, isolated_dirs, monkeypatch):
    app, wb = xs
    ws = wb.Worksheets("ZQX인사")
    ws.charts.append(FakeChartObject(ws, "ZQX매출차트"))
    code, out = run_as_tool(excel, capsys, monkeypatch, "excel", "chart", "--action", "delete", "--name", "없는차트")
    assert code == 1, out
    assert_not_logged(out, isolated_dirs)


def test_guard_ppt_title_and_notes(pp, capsys, isolated_dirs, monkeypatch):
    s = pp["prs"].Slides(1)
    s.Shapes.Title.TextFrame.TextRange.Text = "ZQX 제목"
    s.NotesPage.body.TextFrame.TextRange.Text = "ZQX 노트"
    code, out = run_as_tool(ppt, capsys, monkeypatch, "ppt", "write", "--slide", "1", "--title", "새 제목")
    assert_not_logged(out, isolated_dirs)
    code, out = run_as_tool(ppt, capsys, monkeypatch, "ppt", "notes", "--slide", "1", "--text", "새 노트")
    assert_not_logged(out, isolated_dirs)


@pytest.mark.parametrize("doc,argv", [
    ("ZQX 보고서 초안 작성\r일정", ["write", "--anchor", "ZQY 보고서 초안", "--text", "본문"]),     # 비슷한 문구
    ("ZQX 보고서\r일정", ["write", "--anchor", "zqx보고서", "--text", "본문"]),                    # [추천] 앵커 (공백만 다름)
    ("계획 ZQX1\r계획 ZQX2\r일정", ["write", "--anchor", "계획", "--text", "본문"]),                 # [후보] 단락
])
def test_guard_word_anchor(wd, capsys, isolated_dirs, monkeypatch, doc, argv):
    wd["doc"] = FakeWordDocument(doc)
    code, out = run_as_tool(word, capsys, monkeypatch, "word", *argv)
    assert code == 1, out
    assert_not_logged(out, isolated_dirs)


def test_guard_word_comment_csv(wd, capsys, isolated_dirs, monkeypatch, tmp_path):
    wd["doc"] = FakeWordDocument("계획 ZQX1\r계획 ZQX2\r일정")
    f = tmp_path / "memo.csv"
    f.write_text("기준 글,메모 내용,몇 번째\nZQX없는글,확인\n일정,확인2,ZQX\n", encoding="utf-8")
    code, out = run_as_tool(word, capsys, monkeypatch, "word", "comment", "--action", "add", "--file", str(f))
    assert_not_logged(out, isolated_dirs)                     # 3번째 열이 숫자가 아님: CSV 칸 값
    f.write_text("기준 글,메모 내용\n계획,확인\nZQX없는글,확인2\n", encoding="utf-8")
    code, out = run_as_tool(word, capsys, monkeypatch, "word", "comment", "--action", "add", "--file", str(f))
    assert_not_logged(out, isolated_dirs)                     # 기준 글과 후보 단락


def test_guard_open_document_names(wd, capsys, isolated_dirs, monkeypatch):
    docs = [{"name": "ZQX인사평가.docx", "path": "C:\\docs\\ZQX인사평가.docx", "app": object(), "doc": object()},
            {"name": "ZQX급여.docx", "path": "C:\\docs\\ZQX급여.docx", "app": object(), "doc": object()}]
    monkeypatch.setattr(off, "list_open", lambda kind: docs)
    code, out = run_as_tool(word, capsys, monkeypatch, "word", "read", "--target", "없는문서.docx")
    assert code == 1, out
    assert_not_logged(out, isolated_dirs)


def test_guard_input_file_names(wd, capsys, isolated_dirs, monkeypatch):
    for n in ("ZQX인사.md", "ZQX급여.md"):
        (isolated_dirs["data"] / n).write_text("x", encoding="utf-8")
    code, out = run_as_tool(word, capsys, monkeypatch, "word", "write", "--file", "없는파일.md")
    assert code == 1, out
    assert_not_logged(out, isolated_dirs)


def test_guard_office_open_document_names(capsys, isolated_dirs, monkeypatch):
    docs = [{"name": "ZQX인사.xlsx", "path": "", "app": object(), "doc": object()}]
    monkeypatch.setattr(off, "list_open", lambda kind: docs)
    code, out = run_as_tool(office, capsys, monkeypatch, "office", "focus", "--app", "excel", "--name", "없는문서")
    assert code == 1, out
    assert_not_logged(out, isolated_dirs)


# ---------------------------------------------------------------------------
# outlook 폴더 [후보]는 끝 이름으로 비교, 메일 계정 이름은 로그에 남기지 않음
# ---------------------------------------------------------------------------


def test_not_found_display_and_log_name():
    e = rt.not_found("메일 폴더", "프로젝트", ["ZQX프로젝트A", "기타"], display=lambda n: "ROOT\\" + n, log_name="<계정>\\프로젝트")
    assert e.candidates == ["ROOT\\ZQX프로젝트A"] and "'ROOT\\프로젝트'" in str(e)
    assert "ROOT" not in e.log_message and SECRET not in e.log_message and "<계정>" in e.log_message


# ---------------------------------------------------------------------------
# 경로(extra)의 메일 계정 이름도 로그에 남기지 않음
# ---------------------------------------------------------------------------


def test_not_found_log_extra():
    e = rt.not_found("메일 폴더", "없음", [], extra=" (경로: 계정@x.com\\없음)", log_extra=" (경로: <메일 계정>\\없음)")
    assert "계정@x.com" in str(e) and "계정@x.com" not in e.log_message and "<메일 계정>" in e.log_message
    e = rt.not_found("메일 폴더", "없음", ["ZQX가"], extra=" (경로: 계정@x.com\\없음)", log_extra=" (경로: <계정>)")
    assert "계정@x.com" not in e.log_message and "ZQX" not in e.log_message


# ---------------------------------------------------------------------------
# En-Genius 임시 실행의 데이터 폴더
# ---------------------------------------------------------------------------


def test_engenius_temp_run_uses_fixed_folder(tmp_path):
    got = _probe(tmp_path, _unpacked(tmp_path), ENGENIUS_OUT=str(tmp_path / "out"))
    home = os.path.join(str(tmp_path / "home"), "Documents", "En-Genius", "CLI-data", "office-cli")
    assert got == {"config": os.path.join(home, "config.json"), "data": os.path.join(home, "data"),
                   "cache": os.path.join(home, "cache")}


def test_engenius_run_from_cli_folder_shares_tool_folder(tmp_path):
    got = _probe(tmp_path, ENGENIUS_OUT=str(tmp_path / "out"))
    assert got["config"] == os.path.join(TOOL_DIR, "config.json")
    assert got["data"] == os.path.join(TOOL_DIR, "data") and got["cache"] == os.path.join(TOOL_DIR, "cache")


def test_normal_run_keeps_tool_folder(tmp_path):
    got = _probe(tmp_path)
    assert got["config"] == os.path.join(TOOL_DIR, "config.json")
    assert got["data"] == os.path.join(TOOL_DIR, "data") and got["cache"] == os.path.join(TOOL_DIR, "cache")


def test_env_settings_win_over_engenius(tmp_path):
    got = _probe(tmp_path, _unpacked(tmp_path), ENGENIUS_OUT=str(tmp_path / "out"), OFFICE_CLI_CONFIG=str(tmp_path / "c.json"),
                 OFFICE_CLI_DATA_DIR=str(tmp_path / "d"), OFFICE_CLI_CACHE_DIR=str(tmp_path / "k"))
    assert got == {"config": str(tmp_path / "c.json"), "data": str(tmp_path / "d"), "cache": str(tmp_path / "k")}


def test_engenius_error_log_lands_in_fixed_folder(tmp_path):
    proc = subprocess.run([sys.executable, os.path.join(_unpacked(tmp_path), "excel.py"), "--bogus"], capture_output=True,
                          timeout=60, env=_env(tmp_path, ENGENIUS_OUT=str(tmp_path / "out")))
    assert proc.returncode == 2
    log = tmp_path / "home" / "Documents" / "En-Genius" / "CLI-data" / "office-cli" / "data" / "logs" / "errors.txt"
    assert log.is_file() and "--bogus" in log.read_text(encoding="utf-8")


# ---------------------------------------------------------------------------
# R5 require_lib (1.4.47): 라이브러리가 없으면 missing_library + install_command, 직접 설치하지 않게
# ---------------------------------------------------------------------------


def _check_missing(out, package):
    assert out["status"] == "error" and out["error_type"] == "missing_library", out
    assert out["missing_package"] == package and out["install_command"] == f"python -m pip install {package}", out
    assert out["agent_directive"].startswith("STOP_RETRY:") and "직접 설치하지 마세요" in out["agent_directive"], out


def test_require_lib_ok_and_missing(monkeypatch):
    assert rt.require_lib("json", "json", "시험") is json
    monkeypatch.setitem(sys.modules, "openpyxl", None)
    with pytest.raises(rt.MissingLibraryError) as ei:
        rt.require_lib("openpyxl", "openpyxl", "XLSX 읽기")
    assert isinstance(ei.value, rt.InputError)


@pytest.mark.parametrize("module,package,make", [
    ("docx", "python-docx", "docx"), ("openpyxl", "openpyxl", "xlsx"), ("pptx", "python-pptx", "pptx")])
def test_reader_missing_library(tmp_path, capsys, monkeypatch, module, package, make):
    import openpyxl
    from sample_docs import new_docx, new_pptx
    path = tmp_path / f"a.{make}"
    if make == "docx":
        new_docx().save(str(path))
    elif make == "pptx":
        new_pptx().save(str(path))
    else:
        openpyxl.Workbook().save(str(path))
    monkeypatch.setitem(sys.modules, module, None)
    code, out = run_main(reader, capsys, "read", "--file", str(path))
    assert code == 1
    _check_missing(out, package)


def test_excel_closed_read_missing_openpyxl(tmp_path, capsys, monkeypatch):
    import openpyxl
    path = tmp_path / "a.xlsx"
    openpyxl.Workbook().save(str(path))
    monkeypatch.setitem(sys.modules, "openpyxl", None)
    code, out = run_main(excel, capsys, "read", "--target", str(path))
    assert code == 1
    _check_missing(out, "openpyxl")


def _all_subparsers():
    import argparse
    import excel, office, outlook, ppt, reader, word
    for mod in (excel, word, ppt, outlook, office, reader):
        top = mod.build_cli()
        yield mod.__name__, top
        for act in top._actions:
            if isinstance(act, argparse._SubParsersAction):
                for name, sp in act.choices.items():
                    yield f"{mod.__name__} {name}", sp


def test_every_help_text_builds():
    """1.4.78: 옵션 설명의 '%'(예: '0.0%')가 argparse 서식으로 읽혀 `excel style -h`·`replace -h`가 JSON 없이 traceback만 냈음."""
    bad = []
    for name, p in _all_subparsers():
        try:
            p.format_help()
        except Exception as e:
            bad.append(f"{name}: {e}")
    assert not bad, bad


def test_continue_is_counted_by_the_tool(monkeypatch, tmp_path):
    """1.4.79: CONTINUE 횟수는 도구가 셈. 바로 앞과 같은 progress면 STOP_RETRY(unfinished_limit), 끝난 명령은 기록을 지움."""
    monkeypatch.setenv("OFFICE_CLI_DATA_DIR", str(tmp_path))
    monkeypatch.setattr(rt, "CURRENT_ARGV", ["index", "--folder", "x"])
    first = rt.unfinished_payload(rt.Unfinished("문서 3/10개 등록"), "reader", "index")
    second = rt.unfinished_payload(rt.Unfinished("문서 6/10개 등록"), "reader", "index")
    assert first["continue"] and second["continue_runs"] == 2, second
    stuck = rt.unfinished_payload(rt.Unfinished("문서 6/10개 등록"), "reader", "index")
    assert stuck["status"] == "error" and stuck["error_type"] == "unfinished_limit" and stuck["agent_directive"].startswith("STOP_RETRY:")
    again = rt.unfinished_payload(rt.Unfinished("문서 6/10개 등록"), "reader", "index")     # 기록을 지웠으므로 처음부터
    assert again["continue_runs"] == 1
    rt.continue_clear("reader")
    assert rt._continue_state() == {}
    monkeypatch.setattr(rt, "CONTINUE_MAX", 2)
    for _ in range(2):
        assert rt.unfinished_payload(rt.Unfinished(f"단계 {_}"), "reader", "index")["status"] == "ok"
    assert rt.unfinished_payload(rt.Unfinished("단계 9"), "reader", "index")["error_type"] == "unfinished_limit"


def test_action_typo_is_candidate_and_delete_is_destructive():
    """1.4.83 검토: 'sheet --action del'을 철자만 보고 [추천] delete로 주어 시트 삭제를 묻지 않고 실행하게 했음 -> [후보], destructive.
    'range --action erase'도 merge를 [추천]하지 않음. 읽기 명령에 가장 가까운 오타('raed')만 [추천]."""
    tool_dir = os.path.dirname(rt.__file__)

    def run(*argv):
        proc = subprocess.run([sys.executable, os.path.join(tool_dir, "excel.py"), *argv], capture_output=True, text=True,
                              encoding="utf-8")
        return json.loads([ln for ln in proc.stdout.splitlines() if ln.strip()][-1])
    out = run("sheet", "--action", "del", "--name", "Sheet1")
    assert "[추천]" not in out["message"] and "[후보] 1) " in out["message"], out
    dele = [c for c in out["confirm_commands"] if " delete" in c["command"]]
    assert dele and dele[0]["destructive"] is True, out
    out = run("range", "--action", "erase", "--range", "A1:B2")
    assert "[추천]" not in out["message"] and "[후보]" in out["message"], out
    out = run("raed")
    assert "[추천]" in out["message"] and " read" in out["message"], out


def test_invalid_choice_values_read_with_or_without_quotes():
    """1.4.84 집 PC: Python 3.13.13 argparse는 'choose from list, add'(따옴표 없음)라 값을 못 읽어 빈 '[후보] '를 냈음."""
    want = ["list", "add", "rename", "copy", "delete", "select"]
    new = "argument --action/-a: invalid choice: 'del' (choose from list, add, rename, copy, delete, select)"
    old = "argument --action/-a: invalid choice: 'del' (choose from 'list', 'add', 'rename', 'copy', 'delete', 'select')"
    assert rt.choices_in_message(new) == want
    assert rt.choices_in_message(old) == want
    assert rt.choices_in_message("argument --x: expected one argument") == []


def test_retry_replacing_changes_only_the_named_option(monkeypatch):
    """1.4.83 검토: '--id 2 --index 2'에서 --index만 바꿔야 함 (전에는 --id도 바뀌어 다른 메일의 첨부를 저장), 폴더 이름이 같은 경로는 그대로."""
    monkeypatch.setattr(rt, "CURRENT_TOOL", "outlook")
    monkeypatch.setattr(rt, "CURRENT_ARGV", ["attachments", "--id", "2", "--index", "2", "--save"])
    got = split_command(rt.retry_replacing("2", "1", "--index"))[2:]
    assert got == ["attachments", "--id", "2", "--index", "1", "--save"], got
    monkeypatch.setattr(rt, "CURRENT_TOOL", "excel")
    monkeypatch.setattr(rt, "CURRENT_ARGV", ["read", "--target", "C:\\보고\\요약\\요약.xlsx", "--sheet", "요약"])
    got = split_command(rt.retry_replacing("요약", "요약표"))[2:]
    assert got == ["read", "--target", "C:\\보고\\요약\\요약.xlsx", "--sheet", "요약표"], got    # 값 전체가 같은 인자만
    monkeypatch.setattr(rt, "CURRENT_ARGV", ["chart", "--sheet", "매출", "--name", "매출"])
    got = split_command(rt.retry_replacing("매출", "매출!#1", "--name"))[2:]
    assert got == ["chart", "--sheet", "매출", "--name", "매출!#1"], got
    # 2차 검토: 옵션 묶음 안에서는 모두 (--by 목록과 --desc 둘 다), 목록 안의 시트 주소, --range의 시트가 --sheet보다 앞섬
    monkeypatch.setattr(rt, "CURRENT_ARGV", ["range", "--action", "sort", "--by", "금액,부서", "--desc", "금액"])
    got = split_command(rt.retry_replacing("금액", "B", ("--by", "--desc")))[2:]
    assert got == ["range", "--action", "sort", "--by", "B,부서", "--desc", "B"], got
    monkeypatch.setattr(rt, "CURRENT_ARGV", ["write", "--cells", "C6=1; 요약 !C7=2"])
    got = split_command(rt.retry_replacing("요약", "요약표", "--sheet"))[2:]
    assert got == ["write", "--cells", "C6=1; 요약표!C7=2"], got
    monkeypatch.setattr(rt, "CURRENT_ARGV", ["write", "--range", "매출2!A1", "--value", "1", "--sheet", "기타"])
    got = split_command(rt.retry_replacing("매출2", "매출", "--sheet"))[2:]
    assert got == ["write", "--range", "매출!A1", "--value", "1", "--sheet", "기타"], got
    monkeypatch.setattr(rt, "CURRENT_TOOL", "outlook")
    monkeypatch.setattr(rt, "CURRENT_ARGV", ["search", "--query", "보관", "--folder", "받은 편지함\\보관"])
    got = split_command(rt.retry_replacing("보관", "보관함", "--folder"))[2:]
    assert got == ["search", "--query", "보관", "--folder", "받은 편지함\\보관함"], got
