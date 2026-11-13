# -*- coding: utf-8 -*-
"""pandas-kit 시험 도우미: 도구(AIKIT/pandas-kit)를 서브프로세스로 실행해 stdout 1줄 JSON을 읽고, 시험용 BOM 파일을 만든다."""
import json
import os
import shlex
import subprocess
import sys

import pytest

HERE = os.path.dirname(os.path.abspath(__file__))
KIT = os.path.abspath(os.path.join(HERE, "..", "..", "..", "pandas-kit"))
SCRIPT = os.path.join(KIT, "pandas_kit.py")
sys.path.insert(0, KIT)


@pytest.fixture
def env(tmp_path, monkeypatch):
    """도구 데이터 폴더를 시험 폴더로, 명령 따옴표는 bash 모양 (시험에서 shlex로 나눠 다시 실행)."""
    data = tmp_path / "data"
    monkeypatch.setenv("PANDAS_KIT_DATA_DIR", str(data))
    monkeypatch.setenv("PANDAS_KIT_SHELL", "bash")
    return {"dir": tmp_path, "data": data}


def run(env, *args, cwd=None):
    """(exit code, JSON). stdout은 정확히 1줄이어야 한다."""
    proc = subprocess.run([sys.executable, SCRIPT, *[str(a) for a in args]], capture_output=True, text=True,
                          encoding="utf-8", cwd=str(cwd or env["dir"]), env=os.environ.copy())
    lines = [l for l in proc.stdout.splitlines() if l.strip()]
    assert len(lines) == 1, (proc.stdout, proc.stderr)
    return proc.returncode, json.loads(lines[0])


def ok(env, *args):
    code, out = run(env, *args)
    assert code == 0 and out["status"] == "ok", out
    return out


def recommended(message):
    """message의 '[추천] python <script> args...' -> args (bash 따옴표)."""
    assert "[추천] " in message, message
    parts = shlex.split(message.split("[추천] ", 1)[1])
    assert parts[0] == "python" and parts[1].endswith(".py"), parts
    return parts[2:]


def command_args(command):
    parts = shlex.split(command)
    assert parts[0] == "python", parts
    return parts[2:]


def make_bom(path, rows=None, title=True, extra=None):
    """B열부터 표 (위에 한 칸짜리 제목, 빈 행 둘), 품번은 앞자리 0이 있는 글자."""
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "BOM"
    if title:
        ws["A1"] = "BOM 목록"
        ws.append([])
        ws.append([])
    ws.append([None, "Level", "Parent", "Part Number", "Qty", "Maker", "Desc"])
    for r in rows or [(1, "TOP", "0102", 2, "A", "볼트"), (2, "0102", "0105", 4, "B", "너트"),
                      (1, "TOP", "0200", 1, "A", "판"), (2, "0200", "0105", 3, "C", "너트")]:
        ws.append([None, *r])
    if extra:
        extra(ws)
    wb.save(path)
    return str(path)
