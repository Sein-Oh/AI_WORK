"""office-cli 테스트 공용 도우미: 진입점 실행(stdout 1줄 JSON), 가짜 COM 연결, 되돌리기 지점, 오류 로그, 샘플 파일.

픽스처는 conftest.py, 가짜 Office 객체는 fake_*.py에 있습니다.
"""
import json
import os
import re
import sys
import zipfile

import pytest

import _office as off
import _runtime as rt
import excel
import office
import outlook
import ppt
import word
from fake_office import make_win32com

TESTS_DIR = os.path.dirname(os.path.abspath(__file__))
DEV = os.path.dirname(TESTS_DIR)                                                   # CLI-dev\office-cli (tests, tools)
TOOL_DIR = os.path.join(os.path.dirname(os.path.dirname(DEV)), "CLI", "office-cli")    # 실행 파일은 CLI\office-cli
TOOLS = os.path.join(DEV, "tools")
SHARED_TESTS = os.path.join(os.path.dirname(DEV), "tests")                         # CLI-dev\tests (shell_sim)
if SHARED_TESTS not in sys.path:
    sys.path.insert(0, SHARED_TESTS)
import shell_sim  # noqa: E402


# ---------------------------------------------------------------------------
# 진입점 실행: stdout이 정확히 1줄 JSON인지 확인하고 (exit 코드, JSON)을 돌려줌
# ---------------------------------------------------------------------------

def run_main(mod, capsys, *argv):
    """mod.main(argv)를 실행. 앞 실행의 recovered가 섞이지 않게 비우고 시작합니다."""
    rt.RECOVERED.clear()
    with pytest.raises(SystemExit) as ei:
        mod.main(list(argv))
    out = [ln for ln in capsys.readouterr().out.splitlines() if ln.strip()]
    assert len(out) == 1, out
    return ei.value.code, json.loads(out[0])


def split_command(cmd):
    """도구가 낸 명령 문자열을 실제 셸처럼 인자 목록으로 나눈다 (CLI-dev/tests/shell_sim.py: 도구의 quote_arg와 같은 셸,
    OFFICE_CLI_SHELL 강제 > pwsh 있음 > Windows PowerShell 5.1). 셸 토크나이저 -> 네이티브 명령줄 조립 -> Windows argv 해석.
    shlex.split(posix=True)는 C:\\work\\a.xlsx 의 백슬래시를 지워 Windows 경로가 깨지므로 테스트는 이 함수를 쓴다.
    """
    return shell_sim.split_command(cmd, "OFFICE_CLI_SHELL")


def recommended_argv(message, tool):
    """message의 '[추천] python "<도구 경로>" <인자>'를 나눠 인자만 돌려준다 (1.4.79: [추천]은 그대로 실행할 전체 명령).
    python으로 시작하고 tool(.py)을 실행하는 명령인지도 확인한다. 명령은 message 끝까지."""
    assert "[추천] " in message, message
    parts = split_command(message.split("[추천] ", 1)[1])
    assert len(parts) >= 2 and parts[0] == "python" and os.path.basename(parts[1]) == f"{tool}.py", (parts, message)
    return parts[2:]


def option_value(argv, option):
    """argv에서 option 바로 뒤의 값 (option이 꼭 한 번만 있어야 함)."""
    assert argv.count(option) == 1, (option, argv)
    return argv[argv.index(option) + 1]


def run_excel(capsys, *argv):
    return run_main(excel, capsys, *argv)


def run_word(capsys, *argv):
    return run_main(word, capsys, *argv)


def run_ppt(capsys, *argv):
    return run_main(ppt, capsys, *argv)


def run_office(capsys, *argv):
    return run_main(office, capsys, *argv)


def run_outlook(capsys, *argv):
    return run_main(outlook, capsys, *argv)


def run_fresh(mod, capsys, monkeypatch, *argv):
    """오류 로그가 이번 인자를 기록하도록 CURRENT_ARGV를 비우고 실행."""
    monkeypatch.setattr(rt, "CURRENT_ARGV", None)
    return run_main(mod, capsys, *argv)


def run_as_tool(mod, capsys, monkeypatch, tool, *argv):
    """명령줄에서 '<tool>.py ...'로 부른 것처럼 sys.argv까지 맞추고 실행 (오류 로그의 명령 줄 확인용)."""
    monkeypatch.setattr(sys, "argv", [f"{tool}.py"] + list(argv))
    return run_fresh(mod, capsys, monkeypatch, *argv)


# ---------------------------------------------------------------------------
# 가짜 COM
# ---------------------------------------------------------------------------

def stub_com(monkeypatch):
    """win32com을 가짜로, 열린 문서 목록은 비움."""
    fake, _ = make_win32com()
    monkeypatch.setattr(off, "win32com", fake)
    monkeypatch.setattr(off, "pythoncom", None)
    monkeypatch.setattr(off, "list_open", lambda kind: [])


def set_active(monkeypatch, app, doc):
    """활성 문서를 doc으로."""
    monkeypatch.setattr(off, "get_active", lambda kind: {"name": doc.Name, "path": doc.FullName, "app": app, "doc": doc})


class Bag:
    """속성만 담는 빈 객체."""


def ws_of(wb, name="Sheet1"):
    return wb.Worksheets(name)


def sheet_names(wb):
    return [str(wb.Worksheets(i).Name) for i in range(1, wb.Worksheets.Count + 1)]


def undo_points(kind="excel"):
    """남아 있는 되돌리기 지점(메타 파일 이름, 정렬)."""
    return sorted(s["meta_file"] for s in off.list_snapshots(kind))


def error_log_text(dirs):
    """오류 로그(errors.txt) 내용. 없으면 빈 글."""
    p = os.path.join(dirs["data"], "logs", "errors.txt")
    return open(p, encoding="utf-8").read() if os.path.exists(p) else ""


SECRET = "ZQX"      # 문서 글 표시: stdout에는 보여도 오류 로그에는 남으면 안 되는 글


def assert_not_logged(out, isolated_dirs, marker=SECRET):
    """stdout message에는 문서 글이 그대로 있고(사용자에게 보여 줄 것), 오류 로그에는 없어야 함."""
    assert out["status"] == "error" and marker in out["message"], out
    log = error_log_text(isolated_dirs)
    assert log and marker not in log, log
    return log


# ---------------------------------------------------------------------------
# 입력 값과 샘플 파일
# ---------------------------------------------------------------------------

EMPTY_FIND = [                    # 빈 찾을 글: --find "" 와 --map의 빈 키
    ["--find", "", "--replace", "X"],
    ["--map", '{"": "X"}'],
    ["--map", '{"a": "b", "": "X"}'],
]

CODE_MD = "# 예제\n\n```python\ndef __init__(self):\n    return x**2 + y**2\n```\n"   # 코드 블록 안의 __, **

TABLE = [("이름", "부서", "금액"), ("가", "영업", 300), ("나", "개발", 100), ("다", "영업", 200), ("라", "기획", 1000)]


def set_cached_value(path, old, new):
    """xlsx의 시트 XML에서 old를 new로 (openpyxl은 수식의 계산된 값을 쓰지 않음: 엑셀에서 저장한 파일 흉내).
    엑셀은 저장할 때 workbook.xml의 fullCalcOnLoad 표시를 지우므로 그것도 지웁니다."""
    with zipfile.ZipFile(path) as z:
        parts = {n: z.read(n) for n in z.namelist()}
    for n in parts:
        if n.startswith("xl/worksheets/"):
            parts[n] = parts[n].replace(old.encode(), new.encode())
    parts["xl/workbook.xml"] = parts["xl/workbook.xml"].replace(b' fullCalcOnLoad="1"', b"")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for n, data in parts.items():
            z.writestr(n, data)


def formula_xlsx(path, n=1, cached=False):
    """'합계' 시트: 값 두 행 + =SUM 수식 n행. cached=False면 계산된 값이 없음 (openpyxl 저장)."""
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "합계"
    ws.append(["항목", "값"])
    ws.append(["a", 10])
    ws.append(["b", 20])
    for i in range(n):
        ws.append([f"합계{i + 1}", "=SUM(B2:B3)"])
    wb.save(path)
    if cached:
        set_cached_value(str(path), "<v></v>", "<v>30</v>")
    return str(path)


# next_commands = 묻지 않고 실행해도 되는 것(읽기·미리보기)만 (1.4.75). conftest가 모든 테스트에 자동 적용.
_SAFE_NEXT = {"attachments", "calendar", "check", "diff", "folders", "index", "info", "list", "outline", "read",
              "search", "thread", "windows", "help"}
_NEXT_RE = re.compile(r'(\w+)\.py[\'"]?\s+(\S+)')


def unsafe_next_commands(payload):
    """payload의 next_commands 중 문서·메일을 바꿀 수 있는 명령 목록 (예외: 같은 명령을 다시 실행하는 CONTINUE·drm_or_dialog)."""
    if payload.get("continue") or payload.get("error_type") == "drm_or_dialog":
        return []
    bad = []
    for c in payload.get("next_commands") or []:
        cmd = c.get("command", "")
        m = _NEXT_RE.search(cmd)
        sub = m.group(2) if m else ""
        listing = re.search(r"--action\s+[\"']?list\b", cmd) or (sub == "undo" and "--list" in cmd) \
            or (sub in ("notes", "revisions", "comment") and not re.search(r"--(text|file|action)\b", cmd))
        if not (sub in _SAFE_NEXT or "--dry-run" in cmd or listing):
            bad.append(cmd)
    return bad
