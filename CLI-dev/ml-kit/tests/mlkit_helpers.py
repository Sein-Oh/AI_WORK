# -*- coding: utf-8 -*-
"""ml-kit 테스트 공용 도구: 경로, CLI 실행(stdout 1줄 JSON 확인), 표 파일 만들기, 레이아웃 데이터.

테스트 파일은 이 모듈을 `from mlkit_helpers import ...`로 불러온다. 불러오면 도구 폴더(../../CLI/ml-kit)와
개발 도구 폴더(tools)가 sys.path에 들어가므로 `import ml_predict` 등을 바로 쓸 수 있다.
"""
import csv
import json
import os
import subprocess
import sys
from datetime import datetime, timedelta

import numpy as np
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
DEV = os.path.dirname(HERE)  # CLI-dev\ml-kit (tests, tools)
KIT = os.path.join(os.path.dirname(os.path.dirname(DEV)), "CLI", "ml-kit")  # 실행 파일은 CLI\ml-kit
for _p in (os.path.join(DEV, "tools"), KIT, HERE, os.path.join(os.path.dirname(DEV), "tests")):  # tests: shell_sim
    if _p not in sys.path:
        sys.path.insert(0, _p)

import shell_sim  # noqa: E402  (CLI-dev/tests: 셸 시뮬레이터)

# make_sample()로 만든 데이터의 열 (test_contract, test_safety, test_report_export)
FEATURES = "반응온도(°C),소비전력(kW),압축압력(bar),촉매비율(wt%),수분함량(%)"
TARGET = "생산효율(%)"
INLINE = "반응온도(°C)=240,소비전력(kW)=65,압축압력(bar)=3,촉매비율(wt%)=0.8,수분함량(%)=20"

try:
    import docx  # noqa: F401
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False
try:
    import matplotlib  # noqa: F401
    HAS_MPL = True
except ImportError:
    HAS_MPL = False


def split_command(cmd):
    """도구가 낸 명령 문자열을 실제 셸처럼 인자 목록으로 나눈다 (CLI-dev/tests/shell_sim.py: 도구의 quote_arg와 같은 셸,
    ML_KIT_SHELL 강제 > pwsh 있음 > Windows PowerShell 5.1). 셸 토크나이저 -> 네이티브 명령줄 조립 -> Windows argv 해석.
    shlex.split(posix=True)는 C:\\work\\a.xlsx 의 백슬래시를 지워 Windows 경로가 깨지므로 테스트는 이 함수를 쓴다.
    """
    return shell_sim.split_command(cmd, "ML_KIT_SHELL")


# ---- CLI 실행 ----------------------------------------------------------------------------------------------

def run_bounded(cmd, timeout):
    """subprocess.run(capture_output, timeout)과 같지만 시간 초과 때 자식의 자식까지 끄고 출력 읽기를 15초만 더 기다린다.
    subprocess.run은 끈 뒤 출력 파이프가 닫힐 때까지 끝없이 기다려, 자식이 띄운 프로세스가 파이프를 잡고 있으면 시험 하나가
    47분 걸렸다 (회사 PC 2026-11-11: 600초 상한인데 2840초)."""
    p = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, encoding="utf-8")
    try:
        out, err = p.communicate(timeout=timeout)
    except subprocess.TimeoutExpired:
        if os.name == "nt":
            subprocess.run(["taskkill", "/PID", str(p.pid), "/T", "/F"], capture_output=True)
        p.kill()
        try:
            p.communicate(timeout=15)
        except subprocess.TimeoutExpired:
            pass
        raise
    return subprocess.CompletedProcess(cmd, p.returncode, out, err)


def run_raw(script, *args):
    """도구를 실행하고 (proc, stdout JSON)을 돌려준다. stdout이 정확히 1줄이 아니면 실패."""
    proc = run_bounded([sys.executable, os.path.join(KIT, script), *args], 600)
    lines = [ln for ln in proc.stdout.splitlines() if ln.strip()]
    assert len(lines) == 1, f"{script} must print exactly 1 JSON line, got {len(lines)}:\n{proc.stdout}\n{proc.stderr}"
    return proc, json.loads(lines[0])


def run_tool(script, *args):
    """성공(exit 0, status ok)을 기대하는 실행."""
    proc, payload = run_raw(script, *args)
    assert proc.returncode == 0, f"{script} failed:\nSTDOUT:\n{proc.stdout}\nSTDERR:\n{proc.stderr}"
    assert payload.get("status") == "ok", payload
    return payload


def run_fail(script, *args):
    """실패(exit != 0, status error, STOP_RETRY/ERROR)를 기대하는 실행."""
    proc, payload = run_raw(script, *args)
    assert proc.returncode != 0
    assert payload["status"] == "error"
    assert payload["agent_directive"].startswith(("STOP_RETRY", "ERROR"))
    return payload


def stop_msg(res, proc):
    """exit 1 + STOP_RETRY를 확인하고 message를 돌려준다."""
    assert proc.returncode == 1 and res["agent_directive"].startswith("STOP_RETRY:"), res
    return res["message"]


def ok_res(res, proc):
    """exit 0 + status ok를 확인하고 결과를 돌려준다."""
    assert proc.returncode == 0 and res["status"] == "ok", res
    return res


def recommendation(message):
    """message의 첫 [추천] 값(한 줄)."""
    return message.split("[추천] ")[1].split("\n")[0].strip()


def recommended_command(message):
    """전체 명령인 [추천](python <script> args...)을 (스크립트 이름, 인자 목록)으로."""
    rec = recommendation(message)
    assert rec.startswith("python"), rec
    parts = split_command(rec)
    return os.path.basename(parts[1]), parts[2:]


# ---- 표 파일 만들기 -----------------------------------------------------------------------------------------

def write_rows(path, rows, merges):
    """행 목록을 .csv(utf-8-sig) 또는 .xlsx(병합 범위 포함)로 쓴다."""
    if path.endswith(".csv"):
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            csv.writer(f).writerows(rows)
        return path
    wb = openpyxl.Workbook()
    ws = wb.active
    for row in rows:
        ws.append(row)
    for m in merges:
        ws.merge_cells(m)
    wb.save(path)
    return path


def write_csv(path, rows, sep=";"):
    """구분자를 고른 CSV(utf-8-sig). 기본은 소수점 쉼표 파일의 ';'."""
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        csv.writer(f, delimiter=sep).writerows(rows)
    return str(path)


def write_csv_header(path, header, rows, delim=";"):
    """머리글 한 줄 + 데이터 행의 CSV(utf-8-sig)."""
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=delim)
        w.writerow(header)
        w.writerows(rows)
    return str(path)


def data_file(data, name, head, rows):
    """fixture 폴더(data["tmp"])에 머리글 + 행의 파일을 쓴다 (조회 파일, 새 입력 파일)."""
    return write_rows(os.path.join(str(data["tmp"]), name), [head] + rows, [])


def data_path(data, name):
    return os.path.join(str(data["tmp"]), name)


def write_df(path, df):
    df.to_csv(path, index=False, encoding="utf-8-sig")
    return str(path)


def write_book(path, rows):
    """행 목록을 .xlsx로 쓴다 (병합 없음)."""
    wb = openpyxl.Workbook()
    ws = wb.active
    for row in rows:
        ws.append(row)
    wb.save(path)
    return str(path)


# ---- 레이아웃 데이터 (머리글·단위 행 매트릭스와 합계 행 검사) ------------------------------------------------

FAMILIES = {
    # regression: two numeric inputs, a category, a numeric target
    "A": {"cols": ["반응온도", "압력", "라인", "수율"], "units": ["°C", "bar", None, "%"],
          "x": "반응온도,압력,라인", "y": "수율", "task": []},
    # categorical first, text target (classification)
    "C": {"cols": ["라인", "반응온도", "압력", "등급"], "units": [None, "°C", "bar", None],
          "x": "라인,반응온도,압력", "y": "등급", "task": ["--task", "classification"]},
    # a date input (unit 'yyyy-mm-dd' over date cells), '-' as a blank unit
    "D": {"cols": ["일자", "반응온도", "압력", "라인", "수율"], "units": ["yyyy-mm-dd", "°C", "bar", "-", "%"],
          "x": "일자,반응온도,압력,라인", "y": "수율", "task": []},
}
H, U = FAMILIES["A"]["cols"], FAMILIES["A"]["units"]


def family_rows(fam, n, seed):
    """FAMILIES[fam] 열 순서의 데이터 행 n개 (seed마다 같은 값)."""
    r = np.random.default_rng(seed)
    out = []
    for i in range(n):
        t, p, line = round(float(r.uniform(150, 200)), 1), round(float(r.uniform(1, 5)), 2), str(r.choice(["A", "B"]))
        y = round(0.3 * t + 5 * p + (3 if line == "A" else 0) + float(r.normal(0, 0.5)), 2)
        if fam == "A":
            out.append([t, p, line, y])
        elif fam == "C":
            out.append([line, t, p, "상" if t + 10 * p > 205 else "하"])
        else:
            out.append([datetime(2026, 1, 1) + timedelta(days=i + 40 * seed), t, p, line, y])
    return out


# ---- 검토 A(2026-11-10)의 기본 표: 온도 150~200, 압력 1~5, 수율 = 30 + 0.2*온도 + 3*압력 + 잡음 ----------------------------

def base_rows(n=30, seed=0):
    """[온도, 압력, 수율] 행 n개 (그대로 학습하면 검증 R2 0.9 안팎)."""
    r = np.random.default_rng(seed)
    out = []
    for _ in range(n):
        t, p = round(float(r.uniform(150, 200)), 1), round(float(r.uniform(1, 5)), 2)
        out.append([t, p, round(30 + 0.2 * t + 3 * p + float(r.normal(0, 1)), 2)])
    return out
