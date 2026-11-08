# -*- coding: utf-8 -*-
"""ml-kit 베타 기능 도구 점검 (사람이 실행하거나 OpenCode에서 Gemma가 실행)

    python run_beta_check.py [--dir 시험폴더] [--label 회사PC] [--only B,F] [--agent]

샘플 데이터(정답을 아는 데이터)를 새로 만들어 옆 폴더 ..\\CLI\\ml-kit의 도구로 실행하고, 결과를 자동 판정한다.
LLM 판단 없이 도구 자체가 맞는지만 본다 (Gemma의 요청 처리 검증은 docs/opencode/자동테스트 메시지 팩과 grade_opencode.py).

묶음:
  H  -h 도움말 계약 (stdout JSON 1줄)
  B  이상치 자동 제외와 '제외' 열 (2.5.0)
  C  표 보간·반대 찾기 ml_map (2.6.0)
  D  예측 + 실측 비교값 (2.7.0)
  E  모델 반대 찾기와 운전 구간 (2.8.0)
  F  보고서 + 성능예측표 세트, 일치도, 엑셀 잠금·숨김·재계산 (2.9.x). 기본 실행에서 뺌 (--only F, PAUSED)
  G  파생변수 원본 열 (2.9.1)

결과: 이 파일 옆에 beta_check_<라벨>_<시각>.md (요약 + 실패 상세)와 .xlsx (항목별 표).
--agent: 진행 로그는 stderr, stdout에는 결과 JSON 1줄 (status, passed, failed, report_file, verified, agent_directive).
  -h와 인자 오류도 JSON 1줄 (인자 오류 exit 2), Ctrl+C는 exit 130. 점검을 끝내면 실패 항목이 있어도 exit 0.
"""
import argparse
import hashlib
import json
import os
import platform
import re
import subprocess
import sys
import time
import traceback
from datetime import datetime

HERE = os.path.dirname(os.path.abspath(__file__))
KIT = os.path.join(os.path.dirname(HERE), "CLI", "ml-kit")
sys.path.insert(0, os.path.join(HERE, "ml-kit", "tools"))

RESULTS = []          # {"group", "id", "title", "status" (PASS/FAIL/SKIP), "detail"}
AGENT = False
GROUPS = ("H", "B", "C", "D", "E", "F", "G")
# 기본 실행에서 빼는 묶음: --only로 이름을 주면 돌림. F(보고서·성능예측표)는 다른 로직이 마무리될 때까지 시험에서 뺌 (2026-11-08)
PAUSED = {"F": "보고서·성능예측표는 개발 중이라 기본 시험에서 뺌 (--only F로 실행)"}


def emit(payload):
    """Agent mode: the one stdout JSON line."""
    sys.stdout.write(json.dumps(payload, ensure_ascii=False) + "\n")
    sys.stdout.flush()


class Parser(argparse.ArgumentParser):
    """argparse that keeps the agent contract when --agent is given: -h and argument errors become one stdout JSON line."""

    def __init__(self, *a, **kw):
        super().__init__(*a, allow_abbrev=False, **kw)

    def print_help(self, file=None):
        if not AGENT:
            return super().print_help(file)
        text = self.format_help()
        sys.stderr.write(text)
        emit({"status": "ok", "command": "help", "verified": True, "usage": text,
              "agent_directive": "VERIFIED: 도움말입니다. 사용자에게 필요한 옵션만 알리세요."})

    def error(self, message):
        if not AGENT:
            super().error(message)
        sys.stderr.write(f"[ERROR] {message}\n")
        emit({"status": "error", "error_type": "argument", "message": message,
              "agent_directive": "STOP_RETRY: 인자가 잘못되었습니다. 재시도하지 말고 message를 사용자에게 알리세요."})
        sys.exit(2)


def log(msg):
    print(msg, file=sys.stderr if AGENT else sys.stdout, flush=True)


def run(script, *args, cwd, timeout=1800):
    """Runs one ml-kit tool; returns (exit code, last-line JSON or None, stderr tail)."""
    env = dict(os.environ, PYTHONIOENCODING="utf-8")
    t0 = time.time()
    try:
        proc = subprocess.run([sys.executable, os.path.join(KIT, script), *[str(a) for a in args]], cwd=cwd,
                              capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=timeout, env=env)
    except subprocess.TimeoutExpired:
        log(f"[ERROR] {script} 시간 초과 ({timeout}초)")
        return "TIMEOUT", None, f"{timeout}초 안에 끝나지 않음", 0
    except OSError as e:
        log(f"[ERROR] {script} 실행 실패: {e}")
        return "OSERROR", None, f"{type(e).__name__}: {e}", 0
    lines = [ln for ln in proc.stdout.splitlines() if ln.strip()]
    try:
        payload = json.loads(lines[-1]) if lines else None
    except json.JSONDecodeError:
        payload = None
    log(f"[INFO] {script} {' '.join(str(a) for a in args)[:120]} -> exit {proc.returncode} ({time.time() - t0:.0f}초)")
    return proc.returncode, payload, proc.stderr[-1500:], len(lines)


def check(group, cid, title, ok, detail=""):
    status = "SKIP" if ok is None else "PASS" if ok else "FAIL"
    RESULTS.append({"group": group, "id": cid, "title": title, "status": status, "detail": str(detail)[:600]})
    log(f"[{'OK' if status == 'PASS' else status}] {cid} {title}" + (f" | {str(detail)[:160]}" if status != "PASS" else ""))


def safe(group, cid, title, fn):
    try:
        ok, detail = fn()
    except Exception as e:  # one broken check must not stop the others
        ok, detail = False, f"{type(e).__name__}: {e}"
    check(group, cid, title, ok, detail)


def sha(path):
    return hashlib.sha256(open(path, "rb").read()).hexdigest()


# ------------------------------------------------------------------------------------------------
def group_h(work):
    for script in ("ml_train.py", "ml_predict.py", "ml_report.py", "ml_map.py", "ml_export.py"):
        def f(script=script):
            code, p, _, n = run(script, "-h", cwd=work)
            return (code == 0 and n == 1 and p and p.get("command") == "help"), f"exit {code}, stdout {n}줄, {p}"
        safe("H", f"H-{script}", f"{script} -h: stdout JSON 1줄", f)


def group_b(work, st):
    data = os.path.join(work, "공정_다중출력.xlsx")
    before = sha(data)
    code, p, err, _ = run("ml_train.py", "-f", data, "--header-row", 3, "--unit-row", 4,
                          "-x", "반응온도,압축압력,촉매비율,수분함량,유량,분진", "-y", "생산효율,불순물,소비전력",
                          "--desc", "반응 공정 시험 기록으로 생산효율, 불순물, 소비전력을 예측한다", cwd=work)
    st["train"] = p or {}
    safe("B", "B1", "학습 성공 (출력 3개, 보고서형 엑셀)", lambda: (code == 0 and bool(p) and p["status"] == "ok", f"exit {code} {err[-300:]}"))
    def b2():
        e = p["exclusions"]["생산효율(%)"]
        return (e["manual"] == 1 and e["auto"] >= 1 and e["forced_include"] == 1), e
    safe("B", "B2", "생산효율: 직접 제외 1, 자동 제외 1 이상, '포함' 1", b2)
    def b3():
        import pandas as pd
        rows = pd.read_csv(p["excluded_rows_file"], encoding="utf-8-sig")
        manual = rows[rows["구분"] == "직접 표시"]
        return (int(manual["원본 행"].iloc[0]) == 105 and manual["이유"].iloc[0] == "센서오류"), manual.to_dict("records")
    safe("B", "B3", "excluded_rows.csv: '센서오류'가 엑셀 105행으로 기록", b3)
    safe("B", "B4", "데이터 파일을 고치지 않음", lambda: (sha(data) == before, "해시 다름"))
    safe("B", "B5", "preview에 '검증 R2(제외 행 포함)' 용어", lambda: ("검증 R2(제외 행 포함)" in p["preview"], p["preview"][:200]))


def group_c(work):
    m = os.path.join(work, "맵.xlsx")
    def c1():
        _, p, _, _ = run("ml_map.py", "-f", m, "-i", "A=12.5", cwd=work)
        v = p["points"][0]["values"]
        return (p["mode"] == "grid" and abs(v["B"] - 16.25) < 1e-9 and abs(v["C"] - 87.5) < 1e-9), v
    safe("C", "C1", "격자 보간 A=12.5 -> B 16.25, C 87.5", c1)
    def c2():
        _, p, _, _ = run("ml_map.py", "-f", m, "-i", "A=103", cwd=work)
        pt = p["points"][0]
        return (pt["extrapolated"] and abs(pt["values"]["B"] - 1058.5) < 1e-6 and p.get("warnings")), pt
    safe("C", "C2", "범위 밖 A=103 -> 외삽 1058.5 + 경고", c2)
    def c3():
        _, p, _, _ = run("ml_map.py", "-f", m, "--find", "D=64", "--vary", "A", cwd=work)
        xs = [round(s["A"], 6) for s in p["solutions"] if not s["extrapolated"]]
        return (xs == [20.0, 80.0] and p["agent_directive"].startswith("VERIFIED")), p["preview"]
    safe("C", "C3", "반대 찾기 D=64 -> A 20, 80 (답 2개)", c3)
    def c4():
        _, p, _, _ = run("ml_map.py", "-f", os.path.join(work, "공정_다중출력.xlsx"), "--header-row", 3, "--unit-row", 4,
                         "-i", "반응온도=182,압축압력=3", "-y", "불순물", cwd=work)
        return (p["mode"] == "scattered" and len(p["nearest_rows"]) == 5 and p["points"][0]["basis"]), p["preview"]
    safe("C", "C4", "흩어진 기록 기준 2개 -> 둘러싼 실측 보간 + 가까운 실측 5행", c4)


def group_d(work, st):
    vdir = st.get("train", {}).get("version_dir")
    if not vdir:
        check("D", "D0", "학습 버전 없음 (B1 실패)", False)
        return
    def d1():
        _, p, _, _ = run("ml_predict.py", "-m", vdir, "-i", "반응온도=182,압축압력=3,촉매비율=1,수분함량=20,유량=10,분진=0.5", cwd=work)
        mc = p["predictions_summary"]["불순물(ppm)"]["map_check"][0]
        return (mc["value"] is not None and all(b["row"] >= 5 for b in mc["basis"])), mc
    safe("D", "D1", "한 건 예측에 실측 비교값 (근거는 엑셀 행 번호)", d1)
    def d2():
        _, p, _, _ = run("ml_predict.py", "-m", vdir, "-f", os.path.join(work, "새조건.xlsx"), cwd=work)
        import pandas as pd
        out = pd.read_csv(p["output_file"], encoding="utf-8-sig")
        need = {"map_불순물(ppm)", "map_diff_불순물(ppm)", "map_방식_불순물(ppm)"}
        return (need <= set(out.columns) and p["drift_warning"]), sorted(out.columns)[:12]
    safe("D", "D2", "파일 예측에 map_ 열 3종 + 범위 밖 경고", d2)


def group_e(work, st):
    vdir = st.get("train", {}).get("version_dir")
    if not vdir:
        return
    def e1():
        _, p, _, _ = run("ml_predict.py", "-m", vdir, "--find", "불순물=100", "--vary", "반응온도",
                         "-i", "압축압력=3,촉매비율=1,수분함량=20,유량=10,분진=0.5", cwd=work)
        return (p["count"] >= 1 and p["operating_windows"] and p["agent_directive"].startswith("VERIFIED")), p["preview"]
    safe("E", "E1", "모델 반대 찾기: 답 + 운전 구간 + 확인", e1)
    def e2():
        _, p, _, _ = run("ml_predict.py", "-m", vdir, "--find", "불순물=9999", "--vary", "반응온도",
                         "-i", "압축압력=3,촉매비율=1,수분함량=20,유량=10,분진=0.5", cwd=work)
        return (p["count"] == 0 and "찾은 값 없음" in p["preview"]), p["preview"]
    safe("E", "E2", "닿지 않는 목표 -> 찾은 값 없음 + 가능한 범위", e2)
    def e3():
        _, p, _, _ = run("ml_predict.py", "-m", vdir, "--find", "불순물=100", "--vary", "반응온도", cwd=work)
        return (len(p["filled_defaults"]) == 5 and "중앙값" in p["warnings"][0]), p.get("filled_defaults")
    safe("E", "E3", "고정 입력 생략 -> 학습 중앙값으로 채우고 알림", e3)


def group_f(work, st):
    data = os.path.join(work, "공정_다중출력.xlsx")
    code, p, err, _ = run("ml_report.py", "-f", data, "--title", "반응 공정 성능 예측", cwd=work, timeout=3600)
    st["report"] = p or {}
    safe("F", "F1", "보고서 + 성능예측표 생성", lambda: (code == 0 and bool(p) and bool(p.get("report_file") and p.get("export_file")),
                                                     f"exit {code} {err[-300:]}"))
    if not (p or {}).get("export_file"):
        return
    root = os.path.dirname(p["report_file"])
    same = lambda a, b: os.path.normcase(os.path.realpath(a)) == os.path.normcase(os.path.realpath(b))
    safe("F", "F2", "보고서·성능예측표·history가 같은 폴더", lambda: (
        same(os.path.dirname(p["export_file"]), root) and os.path.isfile(os.path.join(root, "history.jsonl")),
        f"보고서 {p['report_file']} / 성능예측표 {p['export_file']} / history 있음 {os.path.isfile(os.path.join(root, 'history.jsonl'))}"))
    expect = {"생산효율(%)": {"반응온도(°C)", "촉매비율(wt%)", "수분함량(%)"}, "불순물(ppm)": {"반응온도(°C)", "압축압력(bar)"},
              "소비전력(kW)": {"압축압력(bar)", "유량(L/min)"}}
    safe("F", "F3", "출력별로 실제 영향 입력만 남김 (분진 제외)", lambda: (
        all(set(p["export"][t]["inputs"]) == s for t, s in expect.items()), {t: v["inputs"] for t, v in p["export"].items()}))
    safe("F", "F4", "모델 대비 일치도 95% 이상", lambda: (
        all(v["agreement"]["agreement_pct"] >= 95 for v in p["export"].values()),
        {t: v["agreement"]["agreement_pct"] for t, v in p["export"].items()}))
    def f5():
        from docx import Document
        text = "\n".join(x.text for x in Document(p["report_file"]).paragraphs)
        return ("성능예측표 (엑셀 계산기)" in text and "점수 용어" in text and "% 일치" in text), "보고서 문구 확인"
    safe("F", "F5", "보고서에 성능예측표 일치도 줄과 점수 용어", f5)
    def f6():
        import openpyxl
        wb = openpyxl.load_workbook(p["export_file"])
        hidden = [n for n in wb.sheetnames if n.startswith("표_")]
        ok = (all(wb[n].protection.sheet for n in wb.sheetnames) and wb.security.lockStructure
              and hidden and all(wb[n].sheet_state == "veryHidden" for n in hidden)
              and all(wb[n].sheet_state == "visible" for n in wb.sheetnames if n not in hidden))
        calc = wb["계산기"]
        unlocked = [c.coordinate for row in calc.iter_rows() for c in row if not c.protection.locked]
        texts = " ".join(str(c.value) for ws in wb.worksheets for row in ws.iter_rows() for c in row if isinstance(c.value, str))
        no_path = os.path.dirname(root) not in texts
        return (ok and unlocked == [f"C{r}" for r in range(6, 6 + 5)] and no_path), f"잠금/숨김 {ok}, 입력칸 {unlocked}, 경로 없음 {no_path}"
    safe("F", "F6", "엑셀: 모든 시트 잠금, 표_ 완전 숨김, 입력 칸 5개만 열림, 전체 경로 없음", f6)
    if p.get("export_verified"):
        check("F", "F7", "엑셀로 수식 재계산 확인 (export_verified)", True)
    else:
        un = " ".join(p.get("export_unverified") or [])
        check("F", "F7", "엑셀로 수식 재계산 확인 (export_verified)", None if "엑셀이 없는 환경" in un else False, un)


def group_g(work):
    import numpy as np
    import pandas as pd
    r = np.random.default_rng(9)
    df = pd.DataFrame({"온도(°C)": r.uniform(200, 300, 300), "압력": r.uniform(1, 3, 300)})
    df["효율"] = 0.3 * np.minimum(df["온도(°C)"], 250) + 5 * df["압력"] + r.normal(0, 0.2, 300)
    path = os.path.join(work, "상한.csv")
    df.to_csv(path, index=False, encoding="utf-8-sig")
    _, p, _, _ = run("ml_train.py", "-f", path, "-x", "압력", "--eval", "온도상한 = clip(온도, -1e12, 250)", "-y", "효율", cwd=work)
    def g1():
        a = run("ml_predict.py", "-m", p["version_dir"], "-i", "온도=280,압력=2", cwd=work)[1]
        b = run("ml_predict.py", "-m", p["version_dir"], "-i", "온도=250,압력=2", cwd=work)[1]
        pa, pb = a["predictions_summary"]["효율"]["predictions"][0], b["predictions_summary"]["효율"]["predictions"][0]
        return abs(pa - pb) < 1e-6, f"280 -> {pa}, 250 -> {pb}"
    safe("G", "G1", "원본 열을 -x에서 뺀 상한 파생변수: 예측 성공, 250 이상은 250과 같음", g1)


# ------------------------------------------------------------------------------------------------
def safe_label(label):
    """Keeps only characters that are safe in a Windows file name (same rule as run_all_tests.py)."""
    return "".join(c for c in str(label) if c.isalnum() or c in "-_") or "PC"


def write_report(label, work, started):
    label = safe_label(label)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    base = os.path.join(HERE, f"beta_check_{label}_{stamp}")
    npass = sum(r["status"] == "PASS" for r in RESULTS)
    nfail = sum(r["status"] == "FAIL" for r in RESULTS)
    nskip = sum(r["status"] == "SKIP" for r in RESULTS)
    ver = "?"
    try:
        ver = re.search(r'VERSION = "([^"]+)"', open(os.path.join(KIT, "_ml_runtime.py"), encoding="utf-8").read()).group(1)
    except Exception:
        pass
    lines = [f"# ml-kit 베타 도구 점검 ({label})", "",
             f"- 일시: {started:%Y-%m-%d %H:%M} ~ {datetime.now():%H:%M}", f"- ml-kit {ver}, Python {platform.python_version()}, {platform.platform()}",
             f"- 시험 폴더: {os.path.basename(work)}", f"- 결과: 통과 {npass}, 실패 {nfail}, 건너뜀 {nskip}", "",
             "| 묶음 | 항목 | 내용 | 결과 |", "|---|---|---|---|"]
    lines += [f"| {r['group']} | {r['id']} | {r['title']} | {r['status']} |" for r in RESULTS]
    fails = [r for r in RESULTS if r["status"] != "PASS"]
    if fails:
        lines += ["", "## 실패·건너뜀 상세", ""] + [f"- {r['id']} {r['title']} [{r['status']}]: {r['detail']}" for r in fails]
    with open(base + ".md", "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")
    try:
        import pandas as pd
        pd.DataFrame(RESULTS).to_excel(base + ".xlsx", index=False)
    except Exception:
        pass
    return base + ".md", npass, nfail, nskip


def run_groups(work, groups):
    import make_beta_samples
    try:
        make_beta_samples.main(work)
    except Exception as e:
        check("-", "SAMPLES", "샘플 데이터 만들기 (make_beta_samples)", False, f"{type(e).__name__}: {e}")
        return
    st = {}
    if "H" in groups:
        group_h(work)
    if groups & {"B", "D", "E", "F"}:          # D·E·F는 B가 학습한 모델과 이력을 씀
        group_b(work, st)
    if "C" in groups:
        group_c(work)
    if "D" in groups:
        group_d(work, st)
    if "E" in groups:
        group_e(work, st)
    if "F" in groups:
        group_f(work, st)
    if "G" in groups:
        group_g(work)


def main(argv=None):
    global AGENT
    argv = sys.argv[1:] if argv is None else argv
    AGENT = "--agent" in argv                       # known before parsing so -h and argument errors keep the contract
    if AGENT:
        for stream in (sys.stdout, sys.stderr):
            try:
                stream.reconfigure(encoding="utf-8", errors="replace")
            except Exception:
                pass
    ap = Parser(description="ml-kit 베타 기능 도구 점검")
    ap.add_argument("--dir", help="시험 폴더 (생략 시 이 폴더 옆 work\\beta_<시각>)")
    ap.add_argument("--label", default="PC", help="결과 파일 이름에 붙일 라벨")
    ap.add_argument("--only", help="실행할 묶음 (쉼표: H,B,C,D,E,F,G)")
    ap.add_argument("--agent", action="store_true", help="에이전트용: 로그는 stderr, stdout은 결과 JSON 1줄")
    args = ap.parse_args(argv)
    os.environ["ML_KIT_TIME_LIMIT"] = "0"     # 사람이 돌리는 점검: ml-kit 명령 시간 제한(에이전트용 100초) 끔
    groups = {g for g in (args.only or ",".join(GROUPS)).upper().replace(" ", "").split(",") if g}
    if not args.only:
        groups -= set(PAUSED)
    bad = sorted(groups - set(GROUPS))
    if bad or not groups:
        ap.error(f"알 수 없는 묶음: {', '.join(bad) or '(비어 있음)'} (가능: {','.join(GROUPS)})")
    started = datetime.now()
    interrupted = False
    work = os.path.abspath(args.dir or os.path.join(HERE, "work", f"beta_{started:%Y%m%d_%H%M%S}"))
    try:
        os.makedirs(work, exist_ok=True)
        run_groups(work, groups)
        for g in ([] if args.only else sorted(PAUSED)):
            check(g, g, f"{g} 묶음", None, PAUSED[g])
    except KeyboardInterrupt:
        interrupted = True
        check("-", "STOP", "사용자 중단 (Ctrl+C)", False)
    except Exception as e:                          # runner bug: still write the report and the one JSON line
        check("-", "RUNNER", "점검 러너 오류", False, f"{type(e).__name__}: {e}\n{traceback.format_exc()[-1500:]}")
    report, npass, nfail, nskip = None, 0, 0, 0
    try:
        report, npass, nfail, nskip = write_report(args.label, work, started)
        log(f"[STATUS] 통과 {npass}, 실패 {nfail}, 건너뜀 {nskip} -> {report}")
    except Exception as e:
        nfail = sum(r["status"] == "FAIL" for r in RESULTS)
        log(f"[ERROR] 결과 파일을 쓰지 못함: {type(e).__name__}: {e}")
    verified = bool(report) and os.path.isfile(report)
    if AGENT:
        fails = [f"{r['id']} {r['title']}" for r in RESULTS if r["status"] == "FAIL"]
        if interrupted:
            emit({"status": "error", "error_type": "interrupted", "message": "사용자가 중단함", "report_file": report,
                  "passed": npass, "failed": nfail, "skipped": nskip,
                  "agent_directive": "STOP_RETRY: 사용자가 중단했습니다. 재시도하지 마세요."})
            return 130
        payload = {"status": "ok", "passed": npass, "failed": nfail, "skipped": nskip, "failed_items": fails[:20],
                   "report_file": report, "verified": verified}
        if verified:
            payload["agent_directive"] = ("VERIFIED: 점검을 끝냈습니다. passed, failed, skipped 수와 failed_items를 표로 알리고, "
                                          "report_file 경로를 그대로 알리세요. 실패 원인을 추측하지 마세요.")
        else:
            payload["agent_directive"] = ("UNVERIFIED: 점검은 끝났지만 결과 파일을 쓰지 못했습니다. 재시도하지 말고 "
                                          "passed, failed, skipped 수와 failed_items를 그대로 알리세요.")
        emit(payload)
        return 0                                    # the check itself finished: failures are in the JSON
    if interrupted:
        return 130
    return 0 if not nfail else 1


if __name__ == "__main__":
    sys.exit(main())
