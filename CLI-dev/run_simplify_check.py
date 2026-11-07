# -*- coding: utf-8 -*-
"""사용법 단순화(ml-kit 2.9.71, office-cli 1.4.79) 실기 점검. 사람이 실행 (Gemma 없이, 집 PC용).

    python run_simplify_check.py [--label 집PC] [--dir 시험폴더]

틀린 명령을 일부러 실행해 도구가 준 [추천] 전체 명령이나 [후보] 번호의 confirm_commands를, 경량 LLM이 하듯
PowerShell 5.1에 글자 그대로 넣어 실행하고 성공하는지 본다 (Windows 따옴표 처리 확인). 리눅스에서는 bash로 같은 일을 한다.
닫힌 Office 파일에 쓰는 항목은 Windows + Office에서만 실행하고, 아니면 SKIP.
마지막에 OpenCode 설정 폴더의 스킬 배치(ml-kit 스킬 4개, 옛 skills\\ml-kit 없음)를 본다 (setup.bat을 먼저 실행).

결과: 이 파일 옆에 result_simplify_<라벨>_<시각>.md 한 개. 사람용 점검이라 stdout JSON 계약 대상이 아니다.
"""
import argparse
import json
import os
import platform
import shutil
import subprocess
import sys
import time
from datetime import datetime

HERE = os.path.dirname(os.path.abspath(__file__))
CLI = os.path.join(os.path.dirname(HERE), "CLI")
ML = os.path.join(CLI, "ml-kit")
OC = os.path.join(CLI, "office-cli")
WIN = os.name == "nt"
PS_PREFIX = ("$env:PYTHONIOENCODING='utf-8'; $env:PYTHONUTF8='1'; "
             "[Console]::OutputEncoding=[Text.Encoding]::UTF8; ")
RESULTS = []        # (id, title, status, detail)


def log(msg):
    sys.stderr.write(msg + "\n")
    sys.stderr.flush()


def _last_json(text):
    for line in reversed([ln for ln in (text or "").splitlines() if ln.strip()]):
        try:
            return json.loads(line)
        except ValueError:
            continue
    return None


def run_tool(env, script, *args):
    """도구를 인자 목록으로 직접 실행 (틀린 첫 명령). (exit, JSON)"""
    p = subprocess.run([sys.executable, script, *args], capture_output=True, env=env, timeout=900)
    return p.returncode, _last_json(p.stdout.decode("utf-8", "replace"))


def run_shell(env, command):
    """도구가 준 명령 문자열을 LLM처럼 셸에 그대로 넣어 실행. (exit, JSON)"""
    if WIN:
        argv = ["powershell.exe", "-NoProfile", "-NonInteractive", "-Command", PS_PREFIX + command]
    else:
        argv = ["bash", "-c", command]
    p = subprocess.run(argv, capture_output=True, env=env, timeout=900)
    return p.returncode, _last_json(p.stdout.decode("utf-8", "replace"))


def recommended(out):
    msg = str((out or {}).get("message", ""))
    return msg.split("[추천] ", 1)[1].strip().splitlines()[0] if "[추천] " in msg else ""


def record(tid, title, ok, detail=""):
    status = "PASS" if ok is True else ("SKIP" if ok is None else "FAIL")
    RESULTS.append((tid, title, status, detail))
    log(f"[{status}] {tid} {title}" + (f" - {detail}" if detail and status != "PASS" else ""))


def check(tid, title, fn):
    try:
        fn()
    except Exception as e:      # 한 항목의 예외가 나머지를 막지 않음
        record(tid, title, False, f"{type(e).__name__}: {e}")


def ok_(out):
    return bool(out) and out.get("status") == "ok"


def follow(env, tid, title, code, out, kind="추천", need_code=None):
    """첫 실행이 멈추고 준 명령을 셸로 실행해 성공하는지."""
    if kind == "추천":
        cmd = recommended(out)
        if not cmd.startswith("python "):
            return record(tid, title, False, f"[추천]이 전체 명령이 아님: {cmd[:120]!r} / {str((out or {}).get('message', ''))[:200]}")
    else:
        cmds = (out or {}).get("confirm_commands") or []
        if "[후보] 1)" not in str((out or {}).get("message", "")) or not cmds:
            return record(tid, title, False, f"번호 [후보]나 confirm_commands 없음: {str((out or {}).get('message', ''))[:200]}")
        cmd = cmds[0]["command"]
    if need_code is not None and code != need_code:
        return record(tid, title, False, f"첫 실행 exit {code} (기대 {need_code})")
    c2, out2 = run_shell(env, cmd)
    record(tid, title, ok_(out2) and c2 == 0,
           "" if ok_(out2) else f"고친 명령 exit {c2}: {str((out2 or {}).get('message', ''))[:200]} | 명령: {cmd[:200]}")
    return out2


def make_data(work):
    sys.path.insert(0, os.path.join(HERE, "ml-kit", "tests"))
    from make_sample import make_sample
    data = make_sample(os.path.join(work, "공정 데이터.csv"))
    import openpyxl
    wb = openpyxl.Workbook()              # 두 줄 머리글(병합된 위 칸): 도구가 [추천] --header-row 2 전체 명령을 줌
    ws = wb.active
    ws.append(["조건", None, "결과"])
    ws.merge_cells("A1:B1")
    ws.append(["온도(°C)", "압력(%)", "수율(%)"])
    for i in range(40):
        ws.append([170 + i, 10 + i, 0.2 * (170 + i) + 0.1 * (10 + i)])
    titled = os.path.join(work, "두줄 머리글.xlsx")
    wb.save(titled)
    wb = openpyxl.Workbook()
    wb.active.title = "요약"
    wb.active.append(["항목", "금액", "부서"])
    for r in (["가", 100, "영업"], ["나", 250, "영업"], ["다", 300, "개발"]):
        wb.active.append(r)
    wb.create_sheet("데이터").append(["x"])
    book = os.path.join(work, "보고 $자료 (1).xlsx")       # 띄어쓰기·$·괄호가 든 경로 (PowerShell 따옴표)
    wb.save(book)
    return data, titled, book


def ml_checks(env, data, titled):
    T, P, E = (os.path.join(ML, f) for f in ("ml_train.py", "ml_predict.py", "ml_export.py"))
    X = "반응온도(°C),소비전력(kW),압축압력(bar)"
    Y = "생산효율(%)"

    def m1():
        code, out = run_tool(env, T, "-f", data, "-x", "반응온도,소비전력,압축압렬", "-y", "생산효율")
        follow(env, "M1", "열 이름 오타 -> [후보] 1) + confirm_commands[0]을 셸로 실행 -> 학습", code, out, kind="후보", need_code=1)
    check("M1", "열 이름 [후보]", m1)

    def m2():
        code, out = run_tool(env, T, "-f", data, "-x", X, "--target", Y)
        follow(env, "M2", "예전 줄임말 --target -> 인자 오류 + [추천] --targets 전체 명령 -> 학습", code, out, need_code=2)
    check("M2", "줄임말 [추천]", m2)

    def m3():
        code, out = run_tool(env, T, "-f", data, "-x", X, "-y", Y, "--description", "공정 시험 데이터")
        follow(env, "M3", "--description -> [추천] --desc 전체 명령 -> 학습", code, out, need_code=2)
    check("M3", "별칭 [추천]", m3)

    def m4():
        code, out = run_tool(env, T, "-f", data, "-x", X, "-y", Y, "--dtypes", "반응온도(°C):floatt")
        follow(env, "M4", "--dtypes 형식 오타 -> [추천] 전체 명령 -> 학습", code, out, need_code=1)
    check("M4", "--dtypes [추천]", m4)

    def m5():
        code, out = run_tool(env, T, "-f", titled, "-x", "온도,압력", "-y", "수율")
        follow(env, "M5", "두 줄 머리글 엑셀 -> [추천] --header-row 2 전체 명령 -> 학습", code, out, need_code=1)
    check("M5", "머리글 [추천]", m5)

    def m6():
        code, out = run_tool(env, P, "-m", data, "-i", "반응온도=240,소비전력=65,압축압력=3")
        record("M6", "예측 -m <데이터 파일> -> 최고 버전 자동 (model_chosen)", ok_(out) and "model_chosen" in out,
               str((out or {}).get("model_chosen") or (out or {}).get("message", ""))[:200])
    check("M6", "-m 데이터 파일", m6)

    def m7():
        code, out = run_tool(env, P, "-m", data, "-f", data, "--save", os.path.join(os.path.dirname(data), "예측 결과.txt"))
        follow(env, "M7", "--save .txt -> [추천] .csv 전체 명령 -> 예측 저장", code, out, need_code=1)
    check("M7", "--save [추천]", m7)

    def m8():
        code, out = run_tool(env, T, "-f", data, "-x", X, "-y", Y, "--version", "v1", "--overwrite", "--dry-run")
        cmds = (out or {}).get("confirm_commands") or []
        good = ok_(out) and out.get("replaces") is True and cmds and cmds[0].get("destructive") is True
        record("M8", "--dry-run: v1 교체 미리보기 (replaced_files, destructive)", bool(good),
               "" if good else str(out)[:200])
        if good:
            c2, out2 = run_shell(env, cmds[0]["command"])
            record("M8b", "미리보기의 confirm_commands[0]을 셸로 실행 -> v1 교체", ok_(out2) and out2.get("version") == "v1",
                   "" if ok_(out2) else str((out2 or {}).get("message", ""))[:200])
    check("M8", "--dry-run", m8)

    def m9():
        code, out = run_tool(env, E, "-m", data, "--keep", "반응온도x", "--title", "점검")
        follow(env, "M9", "성능예측표 -m <데이터 파일> + --keep 오타 -> [후보] 1) -> 셸로 실행", code, out, kind="후보", need_code=1)
    check("M9", "ml_export", m9)


def office_checks(env, book, work):
    X, W, PP, O = (os.path.join(OC, f) for f in ("excel.py", "word.py", "ppt.py", "office.py"))

    def o1():
        code, out = run_tool(env, X, "read", "--target", book, "--sheet", "요약표")
        follow(env, "O1", "닫힌 엑셀 없는 시트 -> [후보] 1) 요약 + confirm_commands[0]을 셸로 실행 -> 읽기", code, out, kind="후보",
               need_code=1)
    check("O1", "시트 [후보]", o1)

    def o2():
        code, out = run_tool(env, X, "read", "--target", book, "--sheet", "요약", "--where", "금액 >= 200, 부서 = 영업")
        follow(env, "O2", "--where 조건 사이 쉼표 -> [추천] ; 전체 명령 -> 읽기 (1건)", code, out, need_code=1)
    check("O2", "--where [추천]", o2)

    def o3():
        code, out = run_tool(env, O, "transfer", "--from", "excel", "--range", "A1:B2", "--to", "word")
        cmd = recommended(out)
        good = code == 2 and cmd.startswith("python ") and "--from-app" in cmd and "--to-app" in cmd
        record("O3", "transfer 옛 --from/--to -> 인자 오류 + [추천] --from-app/--to-app (실행은 안 함)", good, cmd[:200])
    check("O3", "transfer 이름", o3)

    have_office = WIN

    def o4():
        if not have_office:
            return record("O4", "닫힌 엑셀 --cells 쉼표 -> [추천] -> 셸로 실행 (Excel 필요)", None, "Windows + Office에서만")
        shutil.copy(book, os.path.join(work, "쓰기.xlsx"))
        code, out = run_tool(env, X, "write", "--target", os.path.join(work, "쓰기.xlsx"), "--sheet", "데이터", "--cells", "B1=1, B2=2")
        follow(env, "O4", "닫힌 엑셀 --cells 쉼표 -> [추천] ; 전체 명령 -> 셸로 실행 (Excel)", code, out, need_code=1)
    check("O4", "--cells [추천]", o4)

    def o5():
        if not have_office:
            return record("O5", "닫힌 워드 --slots 쉼표 -> [추천] -> 셸로 실행 (Word 필요)", None, "Windows + Office에서만")
        import docx
        path = os.path.join(work, "양식 $1.docx")
        d = docx.Document()
        d.add_paragraph("업체명: {{업체명}}")
        d.add_paragraph("날짜: {{날짜}}")
        d.save(path)
        code, out = run_tool(env, W, "fill", "--target", path, "--slots", "업체명=한빛, 날짜=2026-11-09")
        follow(env, "O5", "닫힌 워드 --slots 쉼표 -> [추천] ; 전체 명령 -> 셸로 실행 (Word)", code, out, need_code=1)
    check("O5", "--slots [추천]", o5)

    def o6():
        if not have_office:
            return record("O6", "닫힌 PPT --cards '|' -> [추천] -> 셸로 실행 (PowerPoint 필요)", None, "Windows + Office에서만")
        import pptx
        path = os.path.join(work, "발표.pptx")
        try:
            pptx.Presentation().save(path)
        except Exception as e:
            return record("O6", "닫힌 PPT --cards", None, f"python-pptx 기본 템플릿 없음 ({type(e).__name__})")
        code, out = run_tool(env, PP, "write", "--target", path, "--slide", "new", "--title", "핵심", "--cards", "속도:빠름|품질:높음")
        follow(env, "O6", "닫힌 PPT --cards '|' -> [추천] ; 전체 명령 -> 셸로 실행 (PowerPoint)", code, out, need_code=1)
    check("O6", "--cards [추천]", o6)


def setup_checks():
    cfg = os.path.join(os.path.expanduser("~"), ".config", "opencode", "skills")
    if not os.path.isdir(cfg):
        return record("K1", "OpenCode 설정 폴더의 스킬 배치", None, f"{cfg} 없음 (setup.bat을 먼저 실행)")
    want = ["ml-train", "ml-predict", "ml-report", "ml-map"]
    missing = [n for n in want if not os.path.isfile(os.path.join(cfg, n, "SKILL.md"))]
    old = os.path.exists(os.path.join(cfg, "ml-kit"))
    record("K1", "설정 폴더에 ml-kit 스킬 4개, 옛 skills\\ml-kit 없음 (setup.bat 뒤)", not missing and not old,
           f"없는 스킬 {missing}, 옛 폴더 남음={old}")


def write_report(label, work, started):
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    path = os.path.join(HERE, f"result_simplify_{label}_{stamp}.md")
    n = {s: sum(1 for r in RESULTS if r[2] == s) for s in ("PASS", "FAIL", "SKIP")}
    lines = [f"# 사용법 단순화 실기 점검: {label}", "",
             f"- 시각: {datetime.now().isoformat(timespec='seconds')}, 걸린 시간 {time.time() - started:.0f}초",
             f"- PC: {platform.platform()}, Python {platform.python_version()}, 셸: {'PowerShell (Windows)' if WIN else 'bash'}",
             f"- 결과: PASS {n['PASS']}, FAIL {n['FAIL']}, SKIP {n['SKIP']}", "",
             "| 항목 | 내용 | 결과 | 상세 |", "|---|---|---|---|"]
    for tid, title, status, detail in RESULTS:
        lines.append(f"| {tid} | {title} | {status} | {detail.replace('|', '/')[:300]} |")
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")
    return path, n


def main():
    ap = argparse.ArgumentParser(description="사용법 단순화 실기 점검 (사람용)")
    ap.add_argument("--label", default="PC", help="결과 파일 이름에 붙일 PC 라벨 (예: 집PC)")
    ap.add_argument("--dir", help="시험 폴더 (기본: 이 파일 옆 work\\simplify)")
    args = ap.parse_args()
    started = time.time()
    work = os.path.abspath(args.dir or os.path.join(HERE, "work", "simplify"))
    if os.path.isdir(work):
        shutil.rmtree(work, ignore_errors=True)
    os.makedirs(work, exist_ok=True)
    env = dict(os.environ, PYTHONIOENCODING="utf-8", PYTHONUTF8="1",
               ML_KIT_DATA_DIR=os.path.join(work, "_mldata"), OFFICE_CLI_DATA_DIR=os.path.join(work, "_ocdata"),
               ML_KIT_TIME_LIMIT="0", OFFICE_CLI_TIME_LIMIT="0")
    if not WIN:
        env.update(ML_KIT_SHELL="bash", OFFICE_CLI_SHELL="bash")
    log(f"[INFO] 시험 폴더: {work}")
    data, titled, book = make_data(work)
    log("[INFO] ml-kit 점검 (학습이 여러 번 있어 몇 분 걸림)")
    ml_checks(env, data, titled)
    log("[INFO] office-cli 점검")
    office_checks(env, book, work)
    setup_checks()
    path, n = write_report(args.label, work, started)
    log(f"[STATUS] PASS {n['PASS']}, FAIL {n['FAIL']}, SKIP {n['SKIP']} -> {path}")
    return 0 if n["FAIL"] == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
