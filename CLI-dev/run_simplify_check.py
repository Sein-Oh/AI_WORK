# -*- coding: utf-8 -*-
"""사용법 단순화(ml-kit 2.9.71~2.9.73, office-cli 1.4.79~1.4.81) 실기 점검. 사람이 실행 (Gemma 없이, 집 PC용).

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
import re
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
# 결과 파일은 GitHub에 올리므로 경로 안의 사용자 폴더 이름을 가림 (run_all_tests.mask_private와 같은 규칙)
_USER_SEGMENT = re.compile(r"(?i)(\bUsers(?:\\\\|\\|/)+)(?!<사용자>)[^\\/\s\"'<>|:*?]+")


# Office 변경 내용·메모의 작성자(Office 사용자 이름): 'author': '홍길동', "author": "홍길동"
_AUTHOR_VALUE = re.compile(r"""((['"])author\\?\2\s*:\s*\\?['"])(?!<사용자>)[^'"\\]+""")


def mask_private(text):
    text = _USER_SEGMENT.sub(lambda m: m.group(1) + "<사용자>", text)
    return _AUTHOR_VALUE.sub(lambda m: m.group(1) + "<사용자>", text)


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
    record("M9", "성능예측표 -m <데이터 파일> + --keep 오타", None,
           "보고서·성능예측표는 개발 중이라 시험에서 뺌 (2026-11-08)")   # m9는 다시 넣을 때 쓰도록 둠


def ml_new_checks(env, data, work):
    """ml-kit 2.9.72: 입력 자동 선택, 다음 실험 제안, 입력별 영향, 실제값 비교와 재학습 알림."""
    import pandas as pd
    T, P = (os.path.join(ML, f) for f in ("ml_train.py", "ml_predict.py"))
    df = pd.read_csv(data, encoding="utf-8-sig")
    df.insert(0, "시료번호", [f"S-{i:03d}" for i in range(len(df))])
    df["상수"] = 1
    df["효율복사"] = df["생산효율(%)"] * 1.001
    auto = os.path.join(work, "자동 선택 데이터.csv")
    df.to_csv(auto, index=False, encoding="utf-8-sig")
    shifted = df.sample(30, random_state=1).copy()
    shifted["생산효율(%)"] += 15
    shift = os.path.join(work, "새 데이터 (공정 변경).csv")
    shifted.to_csv(shift, index=False, encoding="utf-8-sig")

    def n1():
        code, out = run_tool(env, T, "-f", auto, "-y", "생산효율(%)")
        ex = {e.get("column") for e in (out or {}).get("excluded_inputs", [])}
        good = ok_(out) and {"시료번호", "상수", "효율복사"} <= ex and "반응온도(°C)" in (out or {}).get("features_auto", [])
        record("N1", "-x 없이 학습 -> 도구가 입력을 고름 (ID·상수·타겟 복사 열은 이유와 함께 뺌)", good,
               f"제외={sorted(ex)}, 고른 입력={(out or {}).get('features_auto')}")
        cmds = [c for c in (out or {}).get("next_commands", []) if not c.get("needs_input")]
        if good and cmds:
            c2, out2 = run_shell(env, cmds[0]["command"])
            record("N1b", "학습 결과의 next_commands[0](-x가 적힌 재학습)을 셸로 실행", ok_(out2),
                   "" if ok_(out2) else str((out2 or {}).get("message", ""))[:200])
    check("N1", "입력 자동 선택", n1)

    def n2():
        code, out = run_tool(env, P, "-m", auto, "--suggest", "3")
        f = (out or {}).get("output_file", "")
        record("N2", "--suggest 3 -> 실험 조건 3개와 실험계획 엑셀", ok_(out) and len(out.get("points", [])) == 3 and os.path.isfile(f),
               str((out or {}).get("preview") or (out or {}).get("message", ""))[:200])
        code, out = run_tool(env, P, "-m", auto, "--suggest", "3", "--maximize", "생산효율")
        record("N2b", "--suggest 3 --maximize -> 좋은 쪽 실험 조건", ok_(out) and len(out.get("points", [])) == 3,
               str((out or {}).get("message", ""))[:200])
        code, out = run_tool(env, P, "-m", auto, "--suggest", "99")
        follow(env, "N2c", "--suggest 99(상한 넘음) -> [추천] 전체 명령 -> 실행", code, out, need_code=2) if code == 2 else \
            follow(env, "N2c", "--suggest 99(상한 넘음) -> [추천] 전체 명령 -> 실행", code, out)
    check("N2", "다음 실험 제안", n2)

    def n3():
        code, out = run_tool(env, P, "-m", auto, "-i", "반응온도=240,소비전력=65,압축압력=3,촉매비율=0.8,수분함량=20")
        ex = ((out or {}).get("predictions_summary") or {}).get("생산효율(%)", {}).get("explain")
        record("N3", "한 건 예측 -> 입력별 영향(explain)과 preview의 '높인 입력/낮춘 입력'", ok_(out) and bool(ex) and "입력" in out.get("preview", ""),
               str((out or {}).get("preview", ""))[:200])
    check("N3", "입력별 영향", n3)

    def n4():
        code, out = run_tool(env, P, "-m", auto, "-f", shift)
        cmds = (out or {}).get("next_commands") or []
        good = ok_(out) and out.get("retrain_recommended") is True and cmds and "ml_train.py" in cmds[0]["command"]
        record("N4", "실측이 있는 새 데이터(값이 크게 다름) -> actual_check와 재학습 알림", bool(good),
               json.dumps((out or {}).get("actual_check"), ensure_ascii=False)[:200])
        if good:
            c2, out2 = run_shell(env, cmds[0]["command"])
            record("N4b", "재학습 명령을 셸로 실행 -> 새 데이터로 학습", ok_(out2), "" if ok_(out2) else str((out2 or {}).get("message", ""))[:200])
        code, out = run_tool(env, P, "-m", auto, "-f", auto)
        record("N4c", "학습과 같은 데이터 -> 재학습 알림 없음", ok_(out) and not out.get("retrain_recommended"),
               json.dumps((out or {}).get("actual_check"), ensure_ascii=False)[:200])
    check("N4", "재학습 알림", n4)


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


def new_checks(env, data, book, work):
    """ml-kit 2.9.73 · office-cli 1.4.81: 남은 [후보] 번호, --query 구분자, 빈 결과, --dry-run 미리보기."""
    T = os.path.join(ML, "ml_train.py")
    X, W, PP, R = (os.path.join(OC, f) for f in ("excel.py", "word.py", "ppt.py", "reader.py"))

    def c1():
        code, out = run_tool(env, T, "-f", data, "-x", "반응온도(°C),소비전력(kW)", "-y", "생산효율(%)", "--data-ro", "1")
        follow(env, "C1", "ml-kit 옵션 오타 --data-ro -> [후보] 1) --data-row + confirm_commands[0] -> 학습", code, out, kind="후보",
               need_code=2)
    check("C1", "옵션 [후보] 번호", c1)

    def c2():
        txt = os.path.join(work, "규격 메모.txt")
        with open(txt, "w", encoding="utf-8") as f:
            f.write("정격 전압은 220 V이다.\n절연 저항은 100 MΩ 이상.\n")
        code, out = run_tool(env, R, "search", "--query", "정격 AND 전압", "--file", txt)
        out2 = follow(env, "C2", "reader --query 옛 AND -> [추천] ; 전체 명령 -> 검색 (1건 이상)", code, out, need_code=1)
        if out2 is not None and not out2.get("count"):
            record("C2b", "고친 --query로 찾은 건수", False, f"count={out2.get('count')}")
    check("C2", "--query [추천]", c2)

    def c3():
        import docx
        path = os.path.join(work, "장절 두 번.docx")
        d = docx.Document()
        for h, body in (("검증 방법", "장비 검증"), ("검증 방법", "SW 검증")):
            d.add_heading(h, level=1)
            d.add_paragraph(body)
        d.save(path)
        code, out = run_tool(env, W, "read", "--target", path, "--section", "검증 방법")
        follow(env, "C3", "닫힌 워드 같은 제목 둘 -> [후보] 1) 2) + confirm_commands[0](--occurrence 1) -> 읽기", code, out,
               kind="후보", need_code=1)
    check("C3", "--occurrence [후보]", c3)

    def e1():
        code, out = run_tool(env, X, "read", "--target", book, "--sheet", "요약", "--range", "H10:I12")
        good = code == 0 and out.get("count") == 0 and "찾은 항목 없음" in str(out.get("agent_directive"))
        record("E1", "닫힌 엑셀 빈 범위 읽기 -> count 0 + '찾은 항목 없음'", good, str(out)[:200])
    check("E1", "빈 결과", e1)

    def dry(tid, title, script, path_src, name, argv, unchanged):
        """--dry-run: 파일이 그대로인지 보고, confirm_commands[0]을 셸로 실행해 실제로 바뀌는지 (Office 필요)."""
        if not WIN:
            return record(tid, title, None, "Windows + Office에서만")
        path = os.path.join(work, name)
        shutil.copy(path_src, path) if path_src else None
        before = unchanged(path)
        code, out = run_tool(env, script, *[a.replace("{p}", path) for a in argv], "--dry-run")
        cmds = (out or {}).get("confirm_commands") or []
        if not (ok_(out) and out.get("dry_run") and cmds and "--dry-run" not in cmds[0]["command"]):
            return record(tid, title, False, f"미리보기 응답: {str(out)[:200]}")
        if unchanged(path) != before:
            return record(tid, title, False, "미리보기인데 파일이 바뀜")
        c2, out2 = run_shell(env, cmds[0]["command"])
        record(tid, title, ok_(out2) and unchanged(path) != before,
               "" if ok_(out2) else f"확인 명령 exit {c2}: {str((out2 or {}).get('message', ''))[:200]}")

    def xl_state(path):
        import openpyxl
        wb = openpyxl.load_workbook(path)
        return [(ws.title, [[c.value for c in r] for r in ws.iter_rows()]) for ws in wb.worksheets]

    def d1():
        dry("D1", "닫힌 엑셀 replace --dry-run -> 그대로 -> confirm_commands[0] -> 바뀜 (Excel)", X, book, "바꾸기.xlsx",
            ["replace", "--target", "{p}", "--sheet", "요약", "--find", "영업", "--replace", "영업1팀"], xl_state)
    check("D1", "excel replace --dry-run", d1)

    def d2():
        dry("D2", "닫힌 엑셀 sheet delete --dry-run -> 그대로 -> confirm_commands[0] -> 시트 지움 (Excel)", X, book, "시트지우기.xlsx",
            ["sheet", "--target", "{p}", "--action", "delete", "--name", "데이터"], xl_state)
    check("D2", "excel sheet delete --dry-run", d2)

    def d3():
        if not WIN:
            return record("D3", "닫힌 워드 replace --dry-run (Word)", None, "Windows + Office에서만")
        import docx
        src = os.path.join(work, "_원본.docx")
        d = docx.Document()
        d.add_paragraph("시험 온도는 25 도이다.")
        d.save(src)
        dry("D3", "닫힌 워드 replace --dry-run -> 그대로 -> confirm_commands[0] -> 바뀜 (Word)", W, src, "바꾸기.docx",
            ["replace", "--target", "{p}", "--find", "25 도", "--replace", "23 도"],
            lambda p: [x.text for x in docx.Document(p).paragraphs])
    check("D3", "word replace --dry-run", d3)

    def d4():
        if not WIN:
            return record("D4", "닫힌 PPT slide delete --dry-run (PowerPoint)", None, "Windows + Office에서만")
        import pptx
        src = os.path.join(work, "_원본.pptx")
        try:
            prs = pptx.Presentation()
            for t in ("하나", "둘"):
                prs.slides.add_slide(prs.slide_layouts[5]).shapes.title.text = t
            prs.save(src)
        except Exception as e:
            return record("D4", "닫힌 PPT slide delete --dry-run", None, f"python-pptx 기본 템플릿 없음 ({type(e).__name__})")
        dry("D4", "닫힌 PPT slide delete --dry-run -> 그대로 -> confirm_commands[0] -> 지움 (PowerPoint)", PP, src, "지우기.pptx",
            ["slide", "--target", "{p}", "--action", "delete", "--slide", "2"], lambda p: len(pptx.Presentation(p).slides))
    check("D4", "ppt slide delete --dry-run", d4)


def _bom(path, rows):
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "BOM"
    ws.append(["제품 BOM"])
    ws.append([])
    ws.append(["Level", "Part No", "Description", "Qty", "Ref Des", "Maker"])
    for r in rows:
        ws.append(r)
    wb.save(path)
    return path


def bom_checks(env, work):
    """office-cli 1.4.80 excel compare (BOM 두 판 비교)."""
    X = os.path.join(OC, "excel.py")
    old = _bom(os.path.join(work, "BOM Rev.A.xlsx"), [
        [0, "ASSY-100", "메인보드", 1, "", ""], [1, "R-0100", "저항 10k", 3, "R1-R3", "A사"],
        [1, "C-0200", "콘덴서 1uF", 2, "C1,C2", "B사"], [1, "U-0300", "MCU", 1, "U1", "C사"],
        [0, "ASSY-200", "전원보드", 1, "", ""], [1, "R-0100", "저항 10k", 1, "R10", "A사"], [1, "D-0500", "다이오드", 1, "D1", "F사"]])
    new = _bom(os.path.join(work, "BOM Rev.B $(신규).xlsx"), [
        [0, "ASSY-100", "메인보드", 1, "", ""], [1, "R-0100", "저항 10k", 4, "R1-R4", "A사"],
        [1, "C-0200", "콘덴서 1uF", 2, "C2, C1", "D사"], [1, "U-0301", "MCU 개선", 1, "U1", "C사"],
        [0, "ASSY-200", "전원보드", 1, "", ""], [1, "R-0100", "저항 10k", 1, "R10", "A사"], [1, "L-0400", "인덕터", 1, "L1", "E사"]])

    def b1():
        code, out = run_tool(env, X, "compare", "--target", old, "--target2", new)
        c = (out or {}).get("counts") or {}
        want = {"추가": 1, "삭제": 1, "수량 변경": 1, "위치 변경": 1, "속성 변경": 1, "대체 추정": 1}
        record("B1", "BOM 다단 비교: 상위품번+품번 키, 수량·위치·메이커 변경, 위치 순서만 다름은 그대로, 품번 변경 추정", ok_(out) and c == want,
               f"counts={c}, key={(out or {}).get('key_reason')}")
        cmds = (out or {}).get("next_commands") or []
        if ok_(out) and cmds and not WIN:
            record("B2", "비교 결과의 next_commands[0](바뀐 칸 메모 미리보기)을 셸로 실행 (Excel 필요)", None, "Windows + Office에서만")
        elif ok_(out) and cmds:
            c2, out2 = run_shell(env, cmds[0]["command"])
            record("B2", "비교 결과의 next_commands[0](바뀐 칸 메모 미리보기)을 셸로 실행", ok_(out2) and out2.get("dry_run") is True,
                   "" if ok_(out2) else str((out2 or {}).get("message", ""))[:200])
        out_file = (out or {}).get("output_file")
        if out_file:
            import openpyxl
            names = openpyxl.load_workbook(out_file, read_only=True).sheetnames
            record("B3", "결과 엑셀 시트 (요약, 추가, 삭제, 변경, 대체 추정)", all(n in names for n in ("요약", "추가", "삭제", "변경")), str(names))
    check("B1", "BOM 비교", b1)

    def b4():
        code, out = run_tool(env, X, "compare", "--target", old, "--target2", old.replace("Rev.A", "Rev.A"))
        record("B4", "같은 파일 두 번 -> 입력 오류", code != 0 and (out or {}).get("status") == "error", str((out or {}).get("message", ""))[:150])
    check("B4", "같은 파일", b4)

    def b5():
        same = _bom(os.path.join(work, "BOM 사본.xlsx"), [[0, "ASSY-100", "메인보드", 1, "", ""], [1, "R-0100", "저항 10k", 3, "R3, R1-R2", "A사"]])
        base = _bom(os.path.join(work, "BOM 원본.xlsx"), [[0, "ASSY-100", "메인보드", 1, "", ""], [1, "R-0100", "저항 10k", 3.0, "R1-R3", "A사"]])
        code, out = run_tool(env, X, "compare", "--target", base, "--target2", same)
        record("B5", "표기만 다른 판(3 = 3.0, R1-R3 = R3, R1-R2) -> 바뀐 것 없음 (count 0)", ok_(out) and out.get("count") == 0,
               str((out or {}).get("preview", ""))[:200])
    check("B5", "바뀐 것 없음", b5)


def setup_checks():
    cfg = os.path.join(os.path.expanduser("~"), ".config", "opencode", "skills")
    if not os.path.isdir(cfg):
        return record("K1", "OpenCode 설정 폴더의 스킬 배치", None, f"{cfg} 없음 (setup.bat을 먼저 실행)")
    want = ["ml-train", "ml-predict", "ml-report", "ml-map", "office-data"]
    missing = [n for n in want if not os.path.isfile(os.path.join(cfg, n, "SKILL.md"))]
    old = os.path.exists(os.path.join(cfg, "ml-kit"))
    record("K1", "설정 폴더에 ml-kit 스킬 4개와 office-data, 옛 skills\\ml-kit 없음 (setup.py 뒤)", not missing and not old,
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
        f.write(mask_private("\n".join(lines) + "\n"))
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
    log("[INFO] ml-kit 2.9.72 새 기능 점검")
    ml_new_checks(env, data, work)
    log("[INFO] office-cli 점검")
    office_checks(env, book, work)
    log("[INFO] BOM 비교 점검")
    bom_checks(env, work)
    log("[INFO] ml-kit 2.9.73 · office-cli 1.4.81 점검 ([후보] 번호, --query, 빈 결과, --dry-run)")
    new_checks(env, data, book, work)
    setup_checks()
    path, n = write_report(args.label, work, started)
    log(f"[STATUS] PASS {n['PASS']}, FAIL {n['FAIL']}, SKIP {n['SKIP']} -> {path}")
    return 0 if n["FAIL"] == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
