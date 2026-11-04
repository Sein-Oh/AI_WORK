"""ml-kit Windows 실기 테스트 러너

사용법 (CLI-dev/ml-kit 폴더에서 실행. 시험하는 도구는 옆 폴더 CLI/ml-kit):
    python tools/run_win_tests.py                 # 전체
    python tools/run_win_tests.py --only M1,M3    # 일부만
    python tools/run_win_tests.py --dir "D:\\ml 테스트"

오프라인 pytest(tests/)가 다루지 못하는, 실제 Windows와 Office에서만 확인되는 항목을 시험합니다.
    - 진짜 Excel이 만든 xlsx(제목/단위 행, 날짜 셀)와 Excel이 저장한 CP949 CSV 입력
    - 입력/출력 파일이 Excel/Word에 열려 잠겨 있을 때
    - 결과 CSV를 Excel 2007+에서 열 때 한글, 보고서 docx를 Word에서 열 때
    - next_commands 명령을 cmd.exe / PowerShell에 그대로 넣었을 때 (따옴표, %)
    - 단독 배포 폴더, 학습 강제 종료 뒤 복구, 동시 학습 잠금, 2만 행 처리 시간
    - 모든 호출의 stdout 1줄 JSON / stderr 진행 로그 계약 (에이전트 서브프로세스 조건)
Office는 DispatchEx 전용 인스턴스로만 쓰고 끝나면 닫습니다(사용자 창은 건드리지 않음).

결과 파일 (실행 결과와 상관없이 항상 생성, 준비 단계에서 실패해도 생성):
    <테스트 폴더>/report_<시각>.txt               사람이 보는 요약
    tools/test_logs/testlog_<라벨>_<시각>.txt     AI 판독용 상세 로그 (환경, 통과 기준, 호출별 전체 JSON/stderr,
                                                  실패 시 관련 폴더 상태, 계약 위반, 전체 호출 기록)
    AI에게 판단을 맡기려면 testlog 파일을 그대로 전달하거나 GitHub 웹에서 docs/test_logs/에 업로드하세요
    (git 명령으로 올리지 않음, .gitignore 처리).
    python tools/run_win_tests.py --label 집PC     # 로그 파일 이름에 붙일 라벨
"""
import argparse
import datetime as dt
import json
import locale
import os
import platform
import shutil
import subprocess
import sys
import time
import traceback

DEV = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))      # CLI-dev\ml-kit (tests, tools)
ROOT = os.path.join(os.path.dirname(os.path.dirname(DEV)), "CLI", "ml-kit")  # 실행 파일: CLI-dev 옆의 CLI\ml-kit
sys.path.insert(0, os.path.join(DEV, "tests"))
sys.path.insert(0, ROOT)
PY = sys.executable
CONTRACT = []       # (호출, 문제) - stdout/stderr 계약 위반 기록
CALLS = []          # 모든 도구 호출 기록 (AI 판독용 로그)
CURRENT = {"test": "SETUP"}
LOG_STDERR_CHARS = 4000

# 각 항목의 통과 기준 (로그에 그대로 기록해 실패 원인을 판단할 수 있게 함)
EXPECT = {
    "M1": "status=ok, feature_types에 categorical '라인'과 datetime '일시', mean_val_score >= 0.8, targets에 '생산효율(%)'",
    "M2": "입력 xlsx가 Excel에 열려(잠겨) 있어도 status=ok, description이 M1의 '공정 효율 예측 (실기 테스트)'로 상속",
    "M3": "Excel이 저장한 CSV의 실제 인코딩이 cp949이고 status=ok, mean_val_score >= 0.8",
    "M4": "예측 status=ok, 결과 CSV를 Excel로 열었을 때 머리글에 'pred_생산효율(%)'와 '반응온도(°C)'가 깨지지 않고 보임",
    "M5": "결과 CSV가 Excel에 열려 있으면 (--overwrite를 줘도) status=ok, output_file이 요청 경로와 다른(시각이 붙은) 새 파일, recovered에 '잠김'",
    "M6": "보고서 status=ok, Word로 열림, 표 5개 이상, 문서 첫 문단에 '보고서' 포함",
    "M7": "보고서가 Word에 열려 있으면 (--overwrite를 줘도) status=ok, report_file이 기존 파일과 다른 새 파일, recovered에 '잠김'",
    "M8": "next_commands의 needs_input이 자리표시자('<') 유무와 일치, needs_input=false인 첫 명령을 cmd.exe와 PowerShell에 그대로 넣어 둘 다 status=ok",
    "M9": "ml_predict.py + v1 폴더만 있는 배포 폴더에서 단건 예측 status=ok, predictions와 intervals 존재",
    "M10": "2만 행 학습 300초 미만, 예측 60초 미만, 둘 다 status=ok",
    "M11": "같은 데이터 동시 학습: 두 번째는 status=error + STOP_RETRY + exit 1, 첫 번째는 status=ok",
    "M12": "학습 강제 종료 뒤 재학습 status=ok, .tmp_* 임시 폴더가 남지 않고 새 버전 생성",
    "M13": "--history status=ok, 컬럼 오타는 exit 1 + message에 '[추천]', 없는 옵션은 exit 2 + status=error",
    "M14": "모든 호출의 stdout이 정확히 1줄 JSON, stderr에 \\r 덮어쓰기 없음",
}

try:
    import numpy as np
    import pythoncom
    import pywintypes
    import win32com.client
except ImportError as e:
    print(f"[ERROR] 필요한 패키지가 없습니다: {e}")
    sys.exit(1)

from make_sample import make_class_sample, make_sample  # noqa: E402


class Ctx:
    def __init__(self, args):
        self.dir = os.path.abspath(args.dir)
        self.results = []
        self.state = {}

    def path(self, *parts):
        return os.path.join(self.dir, *parts)


def log(msg):
    print(msg, flush=True)


def decode_console(raw):
    """콘솔 명령(chcp 등) 출력: Windows는 OEM/ANSI 코드 페이지(한국어 949), 그 밖은 로캘 인코딩."""
    encodings = (("oem", "mbcs") if os.name == "nt" else ()) + (locale.getpreferredencoding(False), "utf-8")
    for enc in encodings:
        try:
            return raw.decode(enc)
        except (LookupError, UnicodeDecodeError):
            continue
    return raw.decode("utf-8", "replace")


TEST_IDS = tuple(EXPECT)                 # M1~M14 (항목별 통과 기준과 같은 목록)


def parse_only(text):
    """--only 값을 항목 집합으로. 없는 항목이 있으면 (None, 없는 항목 목록)."""
    if not text:
        return None, []
    only = {x.strip().upper() for x in text.split(",") if x.strip()}
    bad = sorted(only - set(TEST_IDS), key=lambda x: (len(x), x))
    return (None, bad or ["(비어 있음)"]) if bad or not only else (only, [])


def run_tool(script, *args, cwd=None, timeout=900, label=None):
    """도구를 에이전트처럼 서브프로세스로 실행하고 (JSON, 종료코드, stdout, stderr)를 돌려줍니다.
    stdout은 정확히 1줄 JSON이어야 하고, stderr 진행 로그는 줄 단위(\\r 덮어쓰기 없음)여야 합니다."""
    cmd = [PY, os.path.join(ROOT, script) if not os.path.isabs(script) else script] + [str(a) for a in args]
    shown = f"python {os.path.basename(script)} {' '.join(str(a) for a in args)}"
    log(f"    $ {shown}"[:300])
    env = dict(os.environ, PYTHONIOENCODING="utf-8")
    t0 = time.time()
    try:
        proc = subprocess.run(cmd, cwd=cwd or ROOT, capture_output=True, timeout=timeout, env=env)
    except subprocess.TimeoutExpired as e:
        CALLS.append({"test": CURRENT["test"], "label": label or os.path.basename(script), "command": shown,
                      "cwd": cwd or ROOT, "exit_code": "TIMEOUT", "sec": round(time.time() - t0, 1),
                      "stdout": (e.stdout or b"").decode("utf-8", "replace")[-2000:],
                      "stderr_tail": (e.stderr or b"").decode("utf-8", "replace")[-LOG_STDERR_CHARS:], "contract": []})
        raise
    return _parse(proc, label or os.path.basename(script), command=shown, cwd=cwd or ROOT, sec=time.time() - t0)


def _parse(proc, label, command=None, cwd=None, sec=None):
    out = proc.stdout.decode("utf-8", errors="replace")
    err = proc.stderr.decode("utf-8", errors="replace")
    lines = [ln for ln in out.splitlines() if ln.strip()]
    data = None
    issues = []
    if len(lines) != 1:
        issues.append(f"stdout {len(lines)}줄")
    if lines:
        try:
            data = json.loads(lines[-1])
        except Exception:
            issues.append("stdout 마지막 줄이 JSON이 아님")
    if "\r" in err.replace("\r\n", "\n"):
        issues.append("stderr에 \\r 덮어쓰기 (서브프로세스에서는 줄바꿈이어야 함)")
    CONTRACT.extend((label, x) for x in issues)
    CALLS.append({"test": CURRENT["test"], "label": label, "command": command or label, "cwd": cwd,
                  "exit_code": proc.returncode, "sec": None if sec is None else round(sec, 1),
                  "json": data, "stdout": "" if data is not None else out[-2000:],
                  "stderr_tail": err[-LOG_STDERR_CHARS:], "contract": issues})
    return data, proc.returncode, out, err[-1500:]


def run_shell(command, shell_kind, label):
    """next_commands 문자열을 그대로 셸에 넣어 실행합니다."""
    env = dict(os.environ, PYTHONIOENCODING="utf-8")
    t0 = time.time()
    if shell_kind == "cmd":
        proc = subprocess.run(command, shell=True, capture_output=True, timeout=900, env=env)
    else:
        ps = "[Console]::OutputEncoding=[Text.Encoding]::UTF8; " + command
        proc = subprocess.run(["powershell", "-NoProfile", "-Command", ps], capture_output=True, timeout=900, env=env)
    return _parse(proc, label, command=f"[{shell_kind}] {command}", sec=time.time() - t0)


def record(ctx, tid, title, ok, detail="", data=None, code=None, err=""):
    status = "PASS" if ok else "FAIL"
    entry = {"id": tid, "title": title, "status": status, "detail": detail, "exit_code": code,
             "expect": EXPECT.get(tid, ""), "data": data,
             "json": "" if data is None else json.dumps(data, ensure_ascii=False)[:400]}
    if not ok and err:
        entry["stderr_tail"] = err.strip()[-600:]
    ctx.results.append(entry)
    log(f"  => [{status}] {tid} {title} {('- ' + detail) if detail else ''}")


def run_test(ctx, tid, title, fn):
    log(f"\n[{tid}] {title}")
    CURRENT["test"] = tid
    t0 = time.time()
    try:
        fn()
    except Exception as e:
        ctx.results.append({"id": tid, "title": title, "status": "ERROR", "detail": f"{type(e).__name__}: {e}",
                            "expect": EXPECT.get(tid, ""), "trace": traceback.format_exc()[-4000:]})
        log(f"  => [ERROR] {tid} {e}")
    if ctx.results and ctx.results[-1]["id"] == tid:
        ctx.results[-1]["sec"] = round(time.time() - t0, 1)


def ok_(d):
    return (d or {}).get("status") == "ok"


# ---------------------------------------------------------------------------
# Office 전용 인스턴스
# ---------------------------------------------------------------------------
class Office:
    def __init__(self, prog):
        self.app = win32com.client.DispatchEx(prog)
        self.app.Visible = False
        self.app.DisplayAlerts = False if prog.startswith("Excel") else 0

    def quit(self):
        try:
            self.app.Quit()
        except Exception:
            pass


def process_xl_rows(n=150, seed=1):
    rng = np.random.default_rng(seed)
    base = dt.datetime(2026, 1, 1, 8, 0)
    rows = []
    for i in range(n):
        line = "ABC"[i % 3]
        t, p, pr, c, w = rng.uniform(200, 280), rng.uniform(40, 90), rng.uniform(1, 5), rng.uniform(0.2, 1.5), rng.uniform(10, 30)
        eff = 0.2 * t + 8 * c - 0.5 * w + 0.1 * p + {"A": 0, "B": 3, "C": -3}[line] + rng.normal(0, 1.0)
        when = pywintypes.Time(base + dt.timedelta(hours=6 * i))
        rows.append((when, line, round(t, 2), round(p, 2), round(pr, 3), round(c, 3), round(w, 2), round(eff, 3)))
    return rows


HEADERS = ("일시", "라인", "반응온도", "소비전력", "압축압력", "촉매비율", "수분함량", "생산효율")
UNITS = ("", "", "°C", "kW", "bar", "wt%", "%", "%")


DIR_MARKER = ".created_by_ml-kit-tests"


def prepare_test_dir(path):
    """테스트 폴더를 새로 준비한다. 러너가 만든 폴더(표식 파일)나 빈 폴더만 지우고 쓴다.
    사용자의 다른 폴더를 --dir로 잘못 줘도 지우지 않는다."""
    from ml_predict import remove_tree
    if os.path.isdir(path) and os.listdir(path):
        if not os.path.exists(os.path.join(path, DIR_MARKER)):
            raise RuntimeError(
                f"--dir 폴더가 비어 있지 않고 테스트 러너가 만든 폴더가 아니어서 지우지 않았습니다: {path} "
                f"(빈 폴더나 없는 경로를 --dir로 주거나, 필요 없으면 직접 지운 뒤 다시 실행하세요)")
        remove_tree(path)
    os.makedirs(path, exist_ok=True)
    with open(os.path.join(path, DIR_MARKER), "w", encoding="utf-8") as f:
        f.write("ml-kit tools/run_win_tests.py 테스트 폴더. 다음 실행 때 통째로 지워집니다.\n")


def setup(ctx):
    prepare_test_dir(ctx.dir)
    # 러너가 일부러 일으키는 오류가 실제 사용 오류 로그(<도구>/data/logs/errors.txt)와 섞이지 않게 테스트 폴더로 분리
    os.environ["ML_KIT_DATA_DIR"] = os.path.join(ctx.dir, "_data")
    rows = process_xl_rows()
    xl = Office("Excel.Application")
    ctx.state.setdefault("office", {})["Excel"] = str(getattr(xl.app, "Version", "?"))
    try:
        # 1) 실제 Excel이 만든 보고서형 xlsx: 1행 제목, 3행 머리글, 4행 단위, 5행부터 데이터(날짜 셀 포함)
        wb = xl.app.Workbooks.Add()
        ws = wb.Worksheets(1)
        ws.Range("A1").Value = "공정 데이터 보고서 (2026년 1분기)"
        ws.Range("A3:H3").Value = (HEADERS,)
        ws.Range("A4:H4").Value = (UNITS,)
        ws.Range(f"A5:H{4 + len(rows)}").Value = tuple(rows)
        ws.Range(f"A5:A{4 + len(rows)}").NumberFormat = "yyyy-mm-dd hh:mm"
        xlsx = ctx.path("공정 데이터.xlsx")
        wb.SaveAs(xlsx, FileFormat=51)
        wb.Close(SaveChanges=False)
        # 2) Excel '다른 이름으로 저장 > CSV(쉼표로 분리)' = 한국어 Windows에서는 CP949
        wb = xl.app.Workbooks.Add()
        ws = wb.Worksheets(1)
        ws.Range("A1:F1").Value = (("반응온도(°C)", "소비전력(kW)", "압축압력(bar)", "촉매비율(wt%)", "수분함량(%)", "생산효율(%)"),)
        ws.Range(f"A2:F{1 + len(rows)}").Value = tuple(r[2:] for r in rows)
        csv = ctx.path("엑셀저장 CP949.csv")
        wb.SaveAs(csv, FileFormat=6)       # xlCSV
        wb.Close(SaveChanges=False)
    finally:
        xl.quit()
    make_class_sample(ctx.path("품질 등급.csv"))
    make_sample(ctx.path("대용량 2만행.csv"), n=20000, seed=3)
    ctx.state.update(xlsx=xlsx, csv=csv)


# ---------------------------------------------------------------------------
# 테스트
# ---------------------------------------------------------------------------
def test_all(ctx, only):
    s = ctx.state
    xlsx_layout = ["--header-row", "3", "--unit-row", "4"]

    def m1():
        d, c, o, e = run_tool("ml_train.py", "-f", s["xlsx"], *xlsx_layout,
                              "-x", "일시,라인,반응온도,소비전력,압축압력,촉매비율,수분함량", "-y", "생산효율",
                              "--desc", "공정 효율 예측 (실기 테스트)", "--tag", "베이스라인")
        ft = (d or {}).get("feature_types") or {}
        s["m1"] = d
        ok = ok_(d) and "라인" in ft.get("categorical", []) and "일시" in (ft.get("datetime") or {}) \
            and (d or {}).get("mean_val_score", 0) >= 0.8 and "생산효율(%)" in (d or {}).get("targets", [])
        record(ctx, "M1", "Excel이 만든 xlsx (제목/단위 행, 날짜 셀, 범주, 한글 경로+공백)", ok,
               f"score={(d or {}).get('mean_val_score')}, types={ft}", d, c, e)

    def m2():
        xl = Office("Excel.Application")
        try:
            wb = xl.app.Workbooks.Open(s["xlsx"], ReadOnly=False)      # Excel이 파일을 쓰기 잠금
            d, c, o, e = run_tool("ml_train.py", "-f", s["xlsx"], *xlsx_layout,
                                  "-x", "반응온도,소비전력,촉매비율,수분함량,라인", "-y", "생산효율", "--tag", "엑셀에 열린 상태")
            wb.Close(SaveChanges=False)
        finally:
            xl.quit()
        ok = ok_(d) and d.get("description") == "공정 효율 예측 (실기 테스트)"
        record(ctx, "M2", "입력 xlsx가 Excel에 열려 있어도 학습 + 설명 자동 상속", ok,
               f"version={(d or {}).get('version')}, description={(d or {}).get('description')!r}", d, c, e)

    def m3():
        raw = open(s["csv"], "rb").read()
        try:
            raw.decode("utf-8")
            enc = "utf-8"
        except UnicodeDecodeError:
            enc = "cp949"
        d, c, o, e = run_tool("ml_train.py", "-f", s["csv"], "-x", "반응온도,소비전력,촉매비율,수분함량", "-y", "생산효율",
                              "--desc", "CP949 CSV 입력")
        ok = ok_(d) and enc == "cp949" and (d or {}).get("mean_val_score", 0) >= 0.8
        record(ctx, "M3", "Excel이 저장한 CSV(CP949) 입력", ok,
               f"파일 인코딩={enc}, score={(d or {}).get('mean_val_score')}", d, c, e)

    def m4():
        vd = (s.get("m1") or {}).get("version_dir")
        d, c, o, e = run_tool("ml_predict.py", "-m", vd, "-f", s["xlsx"], *xlsx_layout)
        out = (d or {}).get("output_file")
        headers = None
        if out and os.path.exists(out):
            xl = Office("Excel.Application")
            try:
                wb = xl.app.Workbooks.Open(out, ReadOnly=True)     # 더블클릭과 같은 기본 CSV 열기
                ws = wb.Worksheets(1)
                n = ws.UsedRange.Columns.Count
                headers = [ws.Cells(1, i).Value for i in range(1, n + 1)]
                wb.Close(SaveChanges=False)
            finally:
                xl.quit()
        s["pred_csv"] = out
        ok = ok_(d) and headers is not None and "pred_생산효율(%)" in headers and "반응온도(°C)" in headers
        record(ctx, "M4", "xlsx 배치 예측 -> 결과 CSV를 Excel로 열면 한글 정상", ok,
               f"Excel이 읽은 머리글 일부={headers[:3] if headers else None} ... {headers[-3:] if headers else ''}", d, c, e)

    def m5():
        out = s.get("pred_csv")
        vd = (s.get("m1") or {}).get("version_dir")
        xl = Office("Excel.Application")
        try:
            wb = xl.app.Workbooks.Open(out, ReadOnly=False)            # 결과 CSV를 Excel이 잠금
            d, c, o, e = run_tool("ml_predict.py", "-m", vd, "-f", s["xlsx"], *xlsx_layout, "--out", out, "--overwrite")
            wb.Close(SaveChanges=False)
        finally:
            xl.quit()
        got = (d or {}).get("output_file") or ""
        ok = (ok_(d) and os.path.exists(got) and os.path.normcase(got) != os.path.normcase(out)
              and any("잠김" in x for x in (d or {}).get("recovered", [])))
        record(ctx, "M5", "결과 CSV가 Excel에 열려 있으면 새 이름으로 저장", ok,
               f"요청={os.path.basename(out)}, 실제={os.path.basename(got)}", d, c, e)

    def m6():
        root = (s.get("m1") or {}).get("data_root")
        d, c, o, e = run_tool("ml_report.py", "-d", root)
        rep = (d or {}).get("report_file")
        info = None
        if rep and os.path.exists(rep):
            wd = Office("Word.Application")
            s.setdefault("office", {})["Word"] = str(getattr(wd.app, "Version", "?"))
            try:
                doc = wd.app.Documents.Open(rep, ReadOnly=True, AddToRecentFiles=False, Visible=False)
                info = {"tables": int(doc.Tables.Count), "pages": int(doc.ComputeStatistics(2)),
                        "first": doc.Paragraphs(1).Range.Text.strip()[:30]}
                doc.Close(SaveChanges=0)
            finally:
                wd.quit()
        s["report"] = rep
        ok = ok_(d) and info is not None and info["tables"] >= 5 and "보고서" in info["first"]
        record(ctx, "M6", "Word 보고서를 실제 Word로 열기", ok, f"Word 판독={info}", d, c, e)

    def m7():
        rep = s.get("report")
        root = (s.get("m1") or {}).get("data_root")
        wd = Office("Word.Application")
        try:
            doc = wd.app.Documents.Open(rep, ReadOnly=False, AddToRecentFiles=False, Visible=False)   # Word가 잠금
            d, c, o, e = run_tool("ml_report.py", "-d", root, "--overwrite")
            doc.Close(SaveChanges=0)
        finally:
            wd.quit()
        got = (d or {}).get("report_file") or ""
        ok = (ok_(d) and os.path.exists(got) and os.path.normcase(got) != os.path.normcase(rep)
              and any("잠김" in x for x in (d or {}).get("recovered", [])))
        record(ctx, "M7", "보고서가 Word에 열려 있으면 새 이름으로 저장", ok,
               f"기존={os.path.basename(rep)}, 새 파일={os.path.basename(got)}", d, c, e)

    def m8():
        d, c, o, e = run_tool("ml_train.py", "-f", ctx.path("품질 등급.csv"), "-x", "반응온도,촉매비율,수분함량",
                              "-y", "품질등급", "--task", "classification", "--desc", "등급 분류")
        nexts = (d or {}).get("next_commands", [])
        flagged_ok = all(isinstance(x.get("needs_input"), bool) and (x["needs_input"] == ("<" in x["command"]))
                         for x in nexts)
        cmds = [x["command"] for x in nexts if not x.get("needs_input")]
        if not cmds:
            record(ctx, "M8", "next_commands를 cmd.exe / PowerShell에 그대로 실행", False, "실행할 next_commands 없음", d, c, e)
            return
        cmd = cmds[0]
        results = {}
        for kind in ("cmd", "powershell"):
            rd, rc, ro, re_ = run_shell(cmd, kind, f"next_commands via {kind}")
            results[kind] = ((rd or {}).get("status"), (rd or {}).get("version"), rc, re_[-200:] if not ok_(rd) else "")
        ok = ok_(d) and flagged_ok and all(v[0] == "ok" for v in results.values())
        record(ctx, "M8", "next_commands를 cmd.exe / PowerShell에 그대로 실행 (따옴표, %, 괄호) + needs_input 표시", ok,
               f"needs_input 표시 정확={flagged_ok}, 명령={cmd[-120:]!r}, 결과={results}", d, c, e)

    def m9():
        vd = (s.get("m1") or {}).get("version_dir")
        dep = ctx.path("배포 폴더")
        os.makedirs(dep, exist_ok=True)
        shutil.copy2(os.path.join(ROOT, "ml_predict.py"), dep)
        shutil.copytree(vd, os.path.join(dep, "v1"), dirs_exist_ok=True)
        d, c, o, e = run_tool(os.path.join(dep, "ml_predict.py"), "-m", "v1",
                              "-i", "일시=2026-01-20 10:00,라인=B,반응온도=240,소비전력=65,압축압력=3,촉매비율=0.8,수분함량=20",
                              cwd=dep, label="ml_predict.py (배포 폴더)")
        summary = ((d or {}).get("predictions_summary") or {}).get("생산효율(%)", {})
        ok = ok_(d) and summary.get("predictions") and summary.get("intervals")
        record(ctx, "M9", "단독 배포 폴더(ml_predict.py + v1만)에서 예측", ok,
               f"예측={summary.get('predictions')}, 범위={summary.get('intervals')}", d, c, e)

    def m10():
        big = ctx.path("대용량 2만행.csv")
        xargs = ["-x", "반응온도,소비전력,압축압력,촉매비율,수분함량", "-y", "생산효율"]
        t0 = time.time()
        d, c, o, e = run_tool("ml_train.py", "-f", big, *xargs, "--desc", "2만 행 성능")
        t_train = time.time() - t0
        t0 = time.time()
        d2, c2, o2, e2 = run_tool("ml_predict.py", "-m", (d or {}).get("version_dir"), "-f", big)
        t_pred = time.time() - t0
        s["big"] = d
        ok = ok_(d) and ok_(d2) and t_train < 300 and t_pred < 60
        record(ctx, "M10", "2만 행 학습/예측 시간 (기준: 학습 300초, 예측 60초 미만)", ok,
               f"학습 {t_train:.1f}초 (score={(d or {}).get('mean_val_score')}), 예측 {t_pred:.1f}초", d2, c2, e2)

    def m11():
        big = ctx.path("대용량 2만행.csv")
        root = ((s.get("big") or {}).get("data_root")) or ctx.path("대용량 2만행_ml_data")
        xargs = ["-x", "반응온도,소비전력,압축압력,촉매비율,수분함량", "-y", "생산효율"]
        env = dict(os.environ, PYTHONIOENCODING="utf-8")
        a = subprocess.Popen([PY, os.path.join(ROOT, "ml_train.py"), "-f", big, *xargs, "--tag", "동시A"],
                             cwd=ROOT, stdout=subprocess.PIPE, stderr=subprocess.PIPE, env=env)
        time.sleep(4)
        d, c, o, e = run_tool("ml_train.py", "-f", big, *xargs, "--tag", "동시B")
        out_a, err_a = a.communicate(timeout=900)
        da = _parse(type("P", (), {"stdout": out_a, "stderr": err_a, "returncode": a.returncode})(), "ml_train.py (동시A)")[0]
        blocked = (d or {}).get("status") == "error" and (d or {}).get("agent_directive", "").startswith("STOP_RETRY")
        ok = blocked and ok_(da) and c == 1
        record(ctx, "M11", "같은 데이터 동시 학습 -> 두 번째는 STOP_RETRY, 첫 번째는 완료", ok,
               f"B={(d or {}).get('status')} exit={c} msg={(d or {}).get('message', '')[:80]!r}, A={(da or {}).get('status')} {(da or {}).get('version')}",
               d, c, e)
        s["big_root"] = root

    def m12():
        big = ctx.path("대용량 2만행.csv")
        root = s.get("big_root") or ctx.path("대용량 2만행_ml_data")
        xargs = ["-x", "반응온도,소비전력,압축압력,촉매비율,수분함량", "-y", "생산효율"]
        before = sorted(x for x in os.listdir(root) if x.startswith("v")) if os.path.isdir(root) else []
        env = dict(os.environ, PYTHONIOENCODING="utf-8")
        p = subprocess.Popen([PY, os.path.join(ROOT, "ml_train.py"), "-f", big, *xargs, "--tag", "강제종료"],
                             cwd=ROOT, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, env=env)
        time.sleep(6)
        p.kill()                       # 작업 관리자 '작업 끝내기'와 같은 강제 종료 (자신이 띄운 학습 프로세스)
        p.communicate()
        leftovers = [x for x in os.listdir(root) if x.startswith(".tmp_")]
        d, c, o, e = run_tool("ml_train.py", "-f", big, *xargs, "--tag", "종료 후 재학습")
        after_tmp = [x for x in os.listdir(root) if x.startswith(".tmp_")]
        after = sorted(x for x in os.listdir(root) if x.startswith("v"))
        ok = ok_(d) and not after_tmp and (d or {}).get("version") not in before
        record(ctx, "M12", "학습 중 강제 종료 -> 다음 학습이 잠금/임시 폴더 정리 후 정상", ok,
               f"종료 직후 임시 폴더={leftovers}, 재학습={(d or {}).get('version')}, 재학습 후 임시 폴더={after_tmp}, 버전={after}",
               d, c, e)

    def m13():
        root = (s.get("m1") or {}).get("data_root")
        d, c, o, e = run_tool("ml_train.py", "-f", s["xlsx"], "--history")
        bad = run_tool("ml_train.py", "-f", s["xlsx"], "-x", "반응온두", "-y", "생산효율", *xlsx_layout)
        arg = run_tool("ml_train.py", "--없는옵션")
        ok = ok_(d) and (bad[0] or {}).get("status") == "error" and "[추천]" in (bad[0] or {}).get("message", "") \
            and bad[1] == 1 and arg[1] == 2 and (arg[0] or {}).get("status") == "error"
        record(ctx, "M13", "이력 조회 / 오타 추천(exit 1) / 잘못된 옵션(exit 2)도 1줄 JSON", ok,
               f"history={(d or {}).get('status')}, 오타 exit={bad[1]}, 옵션 exit={arg[1]}, root={os.path.basename(root or '')}",
               d, c, e)

    def m14():
        bad = [f"{label}: {problem}" for label, problem in CONTRACT]
        record(ctx, "M14", "모든 호출의 stdout 1줄 JSON / stderr 줄 단위 진행 로그", not bad,
               "위반 없음" if not bad else "; ".join(bad[:10]))

    tests = [("M1", "보고서형 xlsx 학습", m1), ("M2", "열린 xlsx 학습", m2), ("M3", "CP949 CSV", m3),
             ("M4", "결과 CSV를 Excel로", m4), ("M5", "결과 CSV 잠금", m5), ("M6", "보고서를 Word로", m6),
             ("M7", "보고서 잠금", m7), ("M8", "next_commands 셸 실행", m8), ("M9", "단독 배포", m9),
             ("M10", "2만 행 성능", m10), ("M11", "동시 학습 잠금", m11), ("M12", "강제 종료 복구", m12),
             ("M13", "오류 계약", m13), ("M14", "출력 계약", m14)]
    for tid, title, fn in tests:
        if only and tid not in only and tid != "M14":
            continue
        run_test(ctx, tid, title, fn)


def write_report(ctx):
    stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    path = ctx.path(f"report_{stamp}.txt")
    counts = {}
    for r in ctx.results:
        counts[r["status"]] = counts.get(r["status"], 0) + 1
    import sklearn
    lines = [f"ml-kit Windows 테스트 보고서 ({stamp})",
             f"Python {sys.version.split()[0]} / scikit-learn {sklearn.__version__} / 테스트 폴더 {ctx.dir}",
             "요약: " + ", ".join(f"{k}={v}" for k, v in sorted(counts.items())), ""]
    for r in ctx.results:
        lines.append(f"[{r['status']}] {r['id']} {r['title']}")
        for key, label in (("detail", "내용"), ("exit_code", "종료코드"), ("json", "JSON"), ("stderr_tail", "stderr(끝부분)"),
                           ("trace", "trace")):
            if r.get(key) not in (None, ""):
                lines.append(f"    {label}: " + str(r[key]).replace("\n", "\n        "))
        lines.append("")
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    return path, counts


def _safe(fn, default="?"):
    try:
        return fn()
    except Exception:
        return default


def env_info(ctx):
    from importlib import metadata
    pkgs = {}
    for name in ("numpy", "pandas", "scikit-learn", "joblib", "openpyxl", "python-docx", "matplotlib", "pywin32"):
        pkgs[name] = _safe(lambda n=name: metadata.version(n), "미설치")
    skill_ver = _safe(lambda: next(ln.split(":", 1)[1].strip().strip('"') for ln in
                                   open(os.path.join(ROOT, "SKILL.md"), encoding="utf-8") if ln.strip().startswith("version:")))
    git = _safe(lambda: subprocess.run(["git", "-C", ROOT, "log", "-1", "--format=%h %ad %s", "--date=short"],
                                       capture_output=True, text=True, encoding="utf-8").stdout.strip() or "?")
    dirty = _safe(lambda: len([x for x in subprocess.run(["git", "-C", ROOT, "status", "--porcelain", "."],
                                                         capture_output=True, text=True, encoding="utf-8").stdout.splitlines() if x]), "?")
    codepage = _safe(lambda: decode_console(subprocess.run("chcp", shell=True, capture_output=True).stdout).strip())
    return [
        ("OS", f"{platform.platform()} ({platform.version()})"),
        ("Python", f"{sys.version.split()[0]} ({PY})"),
        ("패키지", ", ".join(f"{k} {v}" for k, v in pkgs.items())),
        ("Office", ", ".join(f"{k} {v}" for k, v in (ctx.state.get("office") or {}).items()) or "확인 전"),
        ("ml-kit", f"SKILL.md version {skill_ver}, git {git}, 커밋 안 된 변경 {dirty}개"),
        ("인코딩", f"locale={locale.getpreferredencoding(False)}, stdout={sys.stdout.encoding}, {codepage}"),
        ("테스트 폴더", ctx.dir),
    ]


def _paths_in(data):
    paths = []
    for key in ("data_root", "version_dir", "output_file", "report_file"):
        v = (data or {}).get(key)
        if isinstance(v, str) and v:
            paths.append(v)
    return paths


def _folder_state(path, max_items=40):
    """실패한 호출이 가리킨 폴더의 실제 상태 (파일 이름과 크기)."""
    folder = path if os.path.isdir(path) else os.path.dirname(path)
    if not os.path.isdir(folder):
        return [f"{folder}: 폴더 없음"]
    items = []
    for name in sorted(os.listdir(folder))[:max_items]:
        full = os.path.join(folder, name)
        items.append(f"{name}/" if os.path.isdir(full) else f"{name} ({_safe(lambda: os.path.getsize(full), '?')} B)")
    return [f"{folder}: " + ", ".join(items)]


def write_testlog(ctx, log_dir, label, stamp, fatal=None):
    """AI가 이 파일 하나만 읽고 실패 원인을 판단할 수 있도록 모든 근거를 남깁니다."""
    os.makedirs(log_dir, exist_ok=True)
    safe_label = "".join(ch if ch.isalnum() or ch in "-_" else "_" for ch in (label or "PC"))
    path = os.path.join(log_dir, f"testlog_{safe_label}_{stamp}.txt")
    counts = {}
    for r in ctx.results:
        counts[r["status"]] = counts.get(r["status"], 0) + 1
    bad = [r for r in ctx.results if r["status"] != "PASS"]
    L = ["=" * 78, f"ml-kit 실기 테스트 로그 (AI 판독용)  라벨={label}  생성={stamp}", "=" * 78,
         "결과: " + (", ".join(f"{k}={v}" for k, v in sorted(counts.items())) or "실행된 항목 없음")
         + (f"  /  치명적 오류: {fatal.splitlines()[0]}" if fatal else ""), ""]
    L.append("[판정 표]")
    for r in ctx.results:
        L.append(f"  {r['status']:<5} {r['id']:<6} {r.get('sec', '-')!s:>6}s  {r['title']}")
    L += ["", "[환경]"] + [f"  {k}: {v}" for k, v in env_info(ctx)] + [""]
    if fatal:
        L += ["[치명적 오류 - 테스트를 끝까지 실행하지 못함]", "  " + fatal.replace("\n", "\n  "), ""]

    L.append(f"[실패 상세] {len(bad)}건" + ("" if bad else " - 실패 없음"))
    for r in bad:
        L += ["", "-" * 78, f"[{r['status']}] {r['id']} {r['title']}",
              f"  통과 기준: {r.get('expect') or '-'}", f"  관측 결과: {r.get('detail') or '-'}"]
        if r.get("trace"):
            L.append("  예외 traceback:\n    " + r["trace"].strip().replace("\n", "\n    "))
        calls = [c for c in CALLS if c["test"] == r["id"]]
        for i, c in enumerate(calls, 1):
            L += [f"  호출 {i}: $ {c['command']}",
                  f"    cwd={c.get('cwd')}  exit={c['exit_code']}  {c.get('sec')}s"
                  + (f"  계약위반={c['contract']}" if c.get("contract") else "")]
            if c.get("json") is not None:
                L.append("    stdout JSON: " + json.dumps(c["json"], ensure_ascii=False))
            elif c.get("stdout"):
                L.append("    stdout(끝부분): " + c["stdout"].strip().replace("\n", "\n      "))
            if c.get("stderr_tail", "").strip():
                L.append("    stderr(끝부분):\n      " + c["stderr_tail"].strip().replace("\n", "\n      "))
            for pth in _paths_in(c.get("json")):
                L += ["    폴더 상태: " + x for x in _folder_state(pth)]
        if not calls and r.get("data") is not None:
            L.append("  JSON: " + json.dumps(r["data"], ensure_ascii=False))
    L += ["", "[출력 계약 위반]"] + ([f"  {lab}: {prob}" for lab, prob in CONTRACT] or ["  없음"])
    L += ["", "[전체 호출 기록] (항목, 호출, 종료코드, 시간, status / agent_directive 앞부분)"]
    for c in CALLS:
        j = c.get("json") or {}
        L.append(f"  {c['test']:<6} {c['label'][:40]:<40} exit={c['exit_code']!s:<7} {c.get('sec')!s:>6}s  "
                 f"{j.get('status', '-')} / {str(j.get('agent_directive', ''))[:70]}")
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(L) + "\n")
    return path


def main():
    parser = argparse.ArgumentParser(description="ml-kit Windows 실기 테스트 러너")
    parser.add_argument("--dir", default=os.path.join(os.path.dirname(DEV), "work", "ml 테스트"), help="테스트 폴더 (없는 경로나 빈 폴더. 러너가 만든 폴더만 실행 때마다 비우고 새로 만듦, 한글/공백 경로 권장)")
    parser.add_argument("--only", help="실행할 항목 (예: M1,M4). M1이 만든 모델을 쓰는 항목은 M1도 함께 지정")
    parser.add_argument("--label", default="PC", help="로그 파일 이름에 붙일 PC 라벨 (예: 집PC, 회사PC)")
    parser.add_argument("--log-dir", default=os.path.join(DEV, "tools", "test_logs"),
                        help="AI 판독용 로그 폴더 (기본: tools/test_logs, GitHub 웹으로 docs/test_logs/에 업로드)")
    args = parser.parse_args()
    only, bad = parse_only(args.only)
    if bad:
        parser.error(f"알 수 없는 항목: {', '.join(bad)} (가능: {','.join(TEST_IDS)})")
    pythoncom.CoInitialize()
    ctx = Ctx(args)
    stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    log("=" * 70)
    log(f" ml-kit Windows 테스트 / 테스트 폴더: {ctx.dir}")
    log("=" * 70)
    fatal = None
    try:
        log("[준비] Excel로 테스트 데이터(xlsx, CP949 CSV) 생성")
        try:
            setup(ctx)
        except Exception as e:
            ctx.results.append({"id": "SETUP", "title": "테스트 데이터 준비 (Excel로 xlsx/CSV 생성)", "status": "ERROR",
                                "detail": f"{type(e).__name__}: {e}", "expect": "Excel 실행, xlsx/CSV 저장 성공",
                                "trace": traceback.format_exc()[-4000:]})
            raise
        test_all(ctx, only)
    except BaseException:
        fatal = traceback.format_exc()[-4000:]
        log(f"[ERROR] 테스트 중단: {fatal.splitlines()[-1]}")
    finally:
        os.makedirs(ctx.dir, exist_ok=True)
        testlog = write_testlog(ctx, args.log_dir, args.label, stamp, fatal)
    path, counts = write_report(ctx)
    log("\n" + "=" * 70)
    log(" 결과 요약: " + ", ".join(f"{k}={v}" for k, v in sorted(counts.items())))
    for r in ctx.results:
        if r["status"] != "PASS":
            log(f"   [{r['status']}] {r['id']} {r['title']} {r.get('detail', '')}")
    log(f" 보고서: {path}")
    log(f" AI 판독용 로그: {testlog}")
    log("   (이 파일을 Claude에게 전달하거나 GitHub 웹에서 docs/test_logs/에 업로드하세요)")
    log("=" * 70)


if __name__ == "__main__":
    main()
