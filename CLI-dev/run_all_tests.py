"""CLI 도구 실기 점검 (Windows 회사 PC용, 사람이 실행)

사용법 (C:\\CLI-dev에서. 시험할 도구는 옆 폴더 C:\\CLI의 ml-kit, office-cli):
    python run_all_tests.py --label 회사PC             # 자동 점검 전부 (질문 없음, 자리 비워도 됨)
    python run_all_tests.py --label 회사PC --outlook   # office 단계에 Outlook(G그룹)도
    python run_all_tests.py --label 회사PC --ask       # office 단계의 눈 확인 질문(y/n)에 답하면서
    python run_all_tests.py --label 회사PC --gemma     # 끝에 Gemma 종합 테스트(opencode run)도
    python run_all_tests.py --only diag                # 일부 단계만 (diag, pytest, beta, ml, office, gemma)

단계 (이 순서로 실행. --only로 고르면 고른 단계만 같은 순서로):
    diag    환경 진단: Python, 라이브러리, 문서 보안(DRM) 영향 (python-docx/pptx 기본 템플릿, 새로 쓴 파일)
    pytest  각 도구의 오프라인 테스트와 CLI-dev/tests (python -m pytest)
    beta    ml-kit 베타 점검 (run_beta_check.py, F7 엑셀 재계산 포함). 시험 폴더의 성능예측표를 눈 확인에서 엽니다
    ml      ml-kit 실기 러너 (ml-kit/tools/run_win_tests.py, 10~20분). 기본 단계가 아니며 --only에 넣을 때만
    office  office-cli 실기 러너 (office-cli/tools/run_win_tests.py, 실제 Office, 10~30분)
    gemma   Gemma 종합 테스트: agent_scenarios.yaml의 요청을 `opencode run --format json`으로 한 세션에 차례로 보내고
            grade_opencode.py로 채점 (20~60분). 기본 단계가 아니며 --gemma 또는 --only에 넣을 때만.
            opencode는 --opencode, PATH, C:\\CLI\\opencode.exe 순서로 찾고, 없으면 ENV-SKIP.
            관리자 권한 창이면 열린 문서를 못 보므로 일반 PowerShell에서 실행하세요.
    기본은 diag, pytest, beta, office. Windows + pywin32가 아니면 ml, office는 ENV-SKIP.
    끝에 각 도구의 오류 로그(<도구>/data/logs/errors.txt, 실제 사용 중 생긴 오류)를 늘 붙입니다.

결과는 이 파일 옆의 result_<라벨>_<시각>.md 한 개입니다.
    - 맨 위에 단계별 PASS/FAIL/SKIP 표와 실패, 건너뜀(ENV-SKIP), UNVERIFIED 목록이 있고, 그 아래 단계별 상세, 맨 끝에 도구 오류 로그.
    - 단계가 끝날 때마다 파일을 다시 쓰므로, 중간에 멈추거나 Ctrl+C로 끊어도 그때까지의 내용이 남습니다.
    - 이 파일 하나를 그대로 올리세요 (편집하지 마세요).

이 스크립트는 사람이 실행하는 테스트 도구입니다. 에이전트용 출력 계약(stdout 1줄 JSON) 대상이 아닙니다.
"""
import argparse
import datetime as dt
import glob
import json
import locale
import os
import platform
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import traceback
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))                 # CLI-dev (테스트, 러너)
CLI_DIR = os.path.join(os.path.dirname(HERE), "CLI")               # 실행 파일 (CLI-dev 옆의 CLI)
TOOLS = ("ml-kit", "office-cli")
STAGES = ("diag", "pytest", "beta", "ml", "office", "gemma")      # 실행 순서
DEFAULT_STAGES = ("diag", "pytest", "beta", "office")
TITLES = {"diag": "환경 진단", "pytest": "오프라인 테스트 (pytest)", "beta": "ml-kit 베타 점검 (run_beta_check)",
          "ml": "ml-kit 실기 러너", "office": "office-cli 실기 러너 (실제 Office)",
          "gemma": "Gemma 종합 테스트 (opencode run)"}
PACKAGES = ("numpy", "pandas", "scikit-learn", "scipy", "joblib", "openpyxl", "python-docx", "python-pptx",
            "matplotlib", "pymupdf", "olefile", "pywin32", "pytest", "PyYAML")
IMPORTS = {"numpy": "numpy", "pandas": "pandas", "scikit-learn": "sklearn", "scipy": "scipy", "joblib": "joblib",
           "openpyxl": "openpyxl", "python-docx": "docx", "python-pptx": "pptx", "matplotlib": "matplotlib",
           "pymupdf": "fitz", "olefile": "olefile", "pywin32": "win32com.client", "pytest": "pytest", "PyYAML": "yaml"}
LIST_CAP = 30                                                      # 요약 목록 상한 (넘으면 "외 N건", 전체는 상세에)


# ---------------------------------------------------------------------------
# 결과 파일
# ---------------------------------------------------------------------------
class Stage:
    """한 단계의 결과. status: 대기 / PASS / FAIL / ENV-SKIP / 중단."""

    def __init__(self, key):
        self.key, self.title = key, TITLES[key]
        self.status = "대기"
        self.passed = self.failed = self.skipped = 0
        self.seconds = None
        self.note = ""
        self.failures, self.skips, self.unverified = [], [], []
        self.details = []                                        # (제목, 본문)

    def env_skip(self, why):
        self.status, self.note = "ENV-SKIP", why
        self.skips.append(f"단계 전체: {why}")

    def fail(self, what):
        self.failed += 1
        self.failures.append(what)

    def finish(self):
        if self.status == "대기":
            self.status = "FAIL" if self.failed else "PASS"


def _fence(text):
    return "```text\n" + (text.rstrip() or "(없음)") + "\n```"


def _cell(text):
    return str(text).replace("|", "/").replace("\n", " ")


def _capped(items):
    shown = [f"- {x}" for x in items[:LIST_CAP]]
    if len(items) > LIST_CAP:
        shown.append(f"- ... 외 {len(items) - LIST_CAP}건 (아래 상세)")
    return shown or ["- 없음"]


def tool_versions():
    out = []
    for tool in TOOLS:
        ver = "?"
        skill = os.path.join(CLI_DIR, tool, "SKILL.md")
        if not os.path.isfile(skill):                  # office-cli: 작업별 스킬 여럿 (skills\<이름>\SKILL.md, 버전은 모두 같음)
            found = sorted(glob.glob(os.path.join(CLI_DIR, tool, "skills", "*", "SKILL.md")))
            skill = found[0] if found else skill
        try:
            for line in open(skill, encoding="utf-8"):
                if line.strip().startswith("version:"):
                    ver = line.split(":", 1)[1].strip().strip('"')
                    break
        except OSError:
            ver = "없음"
        out.append(f"{tool} {ver}")
    return ", ".join(out)


class Report:
    """결과 파일 하나 (markdown). 단계가 끝날 때마다 통째로 다시 씁니다 (요약 표가 맨 위에 있어야 하므로)."""

    def __init__(self, path, label, stages):
        self.path, self.label = path, label
        self.started = dt.datetime.now()
        self.ended = None
        self.stages = [Stage(s) for s in stages]
        self.logs = []                                           # 도구 오류 로그 (제목, 본문)
        self.log_note = "아직 모으지 않음"
        self.interrupted = False
        self.write()

    def render(self):
        L = [f"# CLI 실기 점검 결과: {self.label}", "",
             f"- 시작 {self.started:%Y-%m-%d %H:%M:%S}, "
             + (f"끝 {self.ended:%H:%M:%S}" if self.ended else "진행 중 (끝나지 않았으면 마지막으로 기록된 단계까지)"),
             f"- 도구: {tool_versions()} / Python {sys.version.split()[0]} / {platform.platform()}",
             "- 이 파일 하나를 그대로 올리세요. 편집하지 마세요."]
        if self.interrupted:
            L.append("- 사용자가 중단했습니다 (Ctrl+C). 아래는 그때까지의 결과입니다.")
        L += ["", "## 요약", "", "| 단계 | 결과 | PASS | FAIL | SKIP | 시간(초) | 비고 |", "|---|---|---|---|---|---|---|"]
        for i, s in enumerate(self.stages, 1):
            nums = ("-", "-", "-") if s.status == "대기" else (s.passed, s.failed, s.skipped)
            L.append(f"| {i}. {s.title} | {s.status} | {nums[0]} | {nums[1]} | {nums[2]} | "
                     f"{'-' if s.seconds is None else s.seconds} | {_cell(s.note)} |")
        L.append(f"| 도구 오류 로그 | - | - | - | - | - | {_cell(self.log_note)} |")
        fails = [f"[{s.key}] {x}" for s in self.stages for x in s.failures]
        skips = [f"[{s.key}] {x}" for s in self.stages for x in s.skips]
        unver = [f"[{s.key}] {x}" for s in self.stages for x in s.unverified]
        L += ["", f"### 실패 {len(fails)}건", ""] + _capped(fails)
        L += ["", f"### 건너뜀 (ENV-SKIP, SKIP) {len(skips)}건", ""] + _capped(skips)
        L += ["", f"### UNVERIFIED (도구가 결과를 확인하지 못한 호출) {len(unver)}건", ""] + _capped(unver)
        for i, s in enumerate(self.stages, 1):
            L += ["", f"## {i}. {s.title}: {s.status}"]
            if s.note:
                L += ["", s.note]
            for title, text in s.details:
                L += ["", f"### {title}", "", _fence(text)]
        L += ["", "## 도구 오류 로그 (<도구>/data/logs/errors.txt)", "", self.log_note]
        for title, text in self.logs:
            L += ["", f"### {title}", "", _fence(text)]
        return "\n".join(L) + "\n"

    def write(self):
        try:
            with open(self.path, "w", encoding="utf-8") as f:
                f.write(self.render())
        except Exception as e:                                   # 결과 파일 문제로 점검을 멈추지 않음
            echo(f"[WARN] 결과 파일을 쓰지 못했습니다: {type(e).__name__}: {e}")

    def summary_lines(self):
        lines = [f"{s.key:<7} {s.status:<8} PASS {s.passed} / FAIL {s.failed} / SKIP {s.skipped}"
                 + (f"  {s.note}" if s.note else "") for s in self.stages]
        lines.append(f"오류 로그: {self.log_note}")
        lines += [f"  [FAIL] {s.key}: {x}" for s in self.stages for x in s.failures[:10]]
        return lines


def echo(line):
    """Shows text on the console; a character it cannot encode (e.g. cp949 redirected stdout) becomes '?'.
    A display error must not stop the output reader (the child would block on a full pipe until the timeout)
    or raise inside a finally block (the result file would miss its summary)."""
    try:
        print(line, flush=True)
    except UnicodeEncodeError:
        enc = getattr(sys.stdout, "encoding", None) or "utf-8"
        try:
            print(line.encode(enc, errors="replace").decode(enc, errors="replace"), flush=True)
        except Exception:
            pass
    except Exception:
        pass


def say(msg):
    echo(f"[STATUS] {msg}")


def child_env(extra=None):
    return dict(os.environ, PYTHONIOENCODING="utf-8", PYTHONUTF8="1", **(extra or {}))


def kill_tree(proc):
    """하위 프로세스와 그 자식까지 끝냅니다 (Windows: taskkill /T)."""
    if proc.poll() is not None:
        return
    try:
        if os.name == "nt":
            subprocess.run(["taskkill", "/F", "/T", "/PID", str(proc.pid)], capture_output=True, timeout=30)
        else:
            proc.kill()
    except Exception:
        try:
            proc.kill()
        except Exception:
            pass
    try:
        proc.wait(timeout=10)
    except Exception:
        pass


def run_capture(args, cwd, timeout, poll=1.0, env_extra=None):
    """하위 프로세스를 실행해 (종료코드, 출력 전체, 걸린 초, Ctrl+C 여부)를 돌려줍니다. 출력은 화면에도 그대로 보여 줍니다.
    출력은 별도 스레드가 읽으므로, 아무것도 출력하지 않고 멈춘 프로세스도 timeout초가 지나면 끝냅니다."""
    start = time.time()
    lines = []
    interrupted = False
    proc = subprocess.Popen(args, cwd=cwd, env=child_env(env_extra), stdout=subprocess.PIPE, stderr=subprocess.STDOUT)

    def pump():
        try:
            for raw in proc.stdout:
                line = raw.decode("utf-8", errors="replace").rstrip("\r\n")
                lines.append(line)
                echo(line)
        except Exception:
            pass

    reader = threading.Thread(target=pump, daemon=True)
    reader.start()
    try:
        while True:                                  # short waits: Ctrl+C is not seen during one long wait on Windows
            try:
                proc.wait(timeout=poll)
                break
            except subprocess.TimeoutExpired:
                if time.time() - start > timeout:
                    kill_tree(proc)
                    reader.join(5)
                    lines.append(f"[ERROR] 시간 초과 ({timeout}초)로 중단했습니다.")
                    break
    except KeyboardInterrupt:
        interrupted = True
        kill_tree(proc)
        lines.append("[ERROR] 사용자가 중단했습니다 (Ctrl+C).")
    reader.join(5)
    return proc.returncode, "\n".join(lines), round(time.time() - start, 1), interrupted


# ---------------------------------------------------------------------------
# 1. 환경 진단
# ---------------------------------------------------------------------------
def _version(dist):
    try:
        from importlib import metadata
        return metadata.version(dist)
    except Exception:
        return "없음"


def _zip_info(path):
    if not os.path.exists(path):
        return "파일 없음"
    with open(path, "rb") as f:
        head = f.read(16)
    return f"크기 {os.path.getsize(path)}바이트, zip 형식={zipfile.is_zipfile(path)}, 앞 16바이트={head.hex(' ')}"


def diag_docx(lines):
    try:
        import docx
    except Exception as e:
        lines.append(f"python-docx import 실패: {e}")
        return
    tdir = os.path.join(os.path.dirname(docx.__file__), "templates")
    default = os.path.join(tdir, "default.docx")
    lines.append(f"python-docx 기본 템플릿 {default}: {_zip_info(default)} (정상: 38116바이트, zip=True)")
    lines.append(f"python-docx 폴더 템플릿 default-docx-template 있음={os.path.isdir(os.path.join(tdir, 'default-docx-template'))}")
    try:
        d = docx.Document()
        names = [s.name for s in d.styles]
        lines.append(f"docx.Document(): 스타일 {len(names)}개 (정상 164), Caption 있음={'Caption' in names}, "
                     f"본문 단락 {len(d.paragraphs)}개 (정상 0), 앞 단락={[p.text[:60] for p in d.paragraphs][:3]}")
    except Exception as e:
        lines.append(f"docx.Document() 실패: {type(e).__name__}: {e}")


def diag_pptx(lines):
    try:
        import pptx
    except Exception as e:
        lines.append(f"python-pptx import 실패: {e}")
        return
    default = os.path.join(os.path.dirname(pptx.__file__), "templates", "default.pptx")
    lines.append(f"python-pptx 기본 템플릿 {default}: {_zip_info(default)} (정상: 34030바이트, zip=True, 50 4b 03 04로 시작)")
    try:
        prs = pptx.Presentation()
        lines.append(f"pptx.Presentation(): 정상 (레이아웃 {len(prs.slide_layouts)}개)")
    except Exception as e:
        lines.append(f"pptx.Presentation() 실패: {type(e).__name__}: {e}")


def diag_new_files(lines):
    """도구가 새로 쓴 Office 파일도 문서 보안(DRM)이 바꾸는지: 쓰고 잠시 뒤 다시 읽어 zip 형식인지 봅니다."""
    work = tempfile.mkdtemp(prefix="cli_diag_")
    made = []
    try:
        import openpyxl
        p = os.path.join(work, "진단.xlsx")
        wb = openpyxl.Workbook()
        wb.active["A1"] = "진단"
        wb.save(p)
        made.append(p)
    except Exception as e:
        lines.append(f"xlsx 쓰기 실패: {type(e).__name__}: {e}")
    try:
        import docx
        tdir = os.path.join(os.path.dirname(docx.__file__), "templates", "default-docx-template")
        try:
            d = docx.Document(tdir)
        except Exception as e:                       # 회사 PC: 폴더 템플릿에서 docProps/thumbnail.jpeg가 빠져 있었음
            lines.append(f"python-docx 폴더 템플릿 열기 실패 ({type(e).__name__}: {e}) -> default.docx로 진단 계속")
            d = docx.Document()
        d.add_paragraph("진단")
        p = os.path.join(work, "진단.docx")
        d.save(p)
        made.append(p)
    except Exception as e:
        lines.append(f"docx 쓰기 실패: {type(e).__name__}: {e}")
    p = os.path.join(work, "진단.txt")
    with open(p, "w", encoding="utf-8") as f:
        f.write("진단")
    made.append(p)
    time.sleep(3)                                  # 보안 프로그램이 파일을 바꿀 시간
    for p in made:
        lines.append(f"새로 쓴 {os.path.basename(p)}: {_zip_info(p) if not p.endswith('.txt') else f'크기 {os.path.getsize(p)}바이트 (정상 6)'}")
    try:
        import openpyxl
        xlsx = [p for p in made if p.endswith(".xlsx")]
        if xlsx:
            openpyxl.load_workbook(xlsx[0], read_only=True).close()
        lines.append("새로 쓴 xlsx 다시 열기: 정상")
    except Exception as e:
        lines.append(f"새로 쓴 xlsx 다시 열기 실패: {type(e).__name__}: {e}")
    shutil.rmtree(work, ignore_errors=True)


def diag_system(lines):
    """Office 버전, 문서 보안(DRM)으로 보이는 프로세스, 설치된 패키지 전체 (Windows에서만 일부)."""
    if os.name == "nt":
        try:
            import winreg
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Office\ClickToRun\Configuration")
            for name in ("VersionToReport", "ProductReleaseIds", "Platform"):
                try:
                    lines.append(f"Office {name}: {winreg.QueryValueEx(key, name)[0]}")
                except OSError:
                    pass
        except OSError:
            lines.append("Office: ClickToRun 정보 없음 (MSI 설치이거나 Office 없음)")
        try:
            out = subprocess.run(["tasklist", "/FO", "CSV", "/NH"], capture_output=True, timeout=60).stdout
            names = sorted({row.split('","')[0].strip('"') for row in out.decode("mbcs", errors="replace").splitlines() if row})
            keys = ("drm", "fasoo", "softcamp", "markany", "nasca", "document", "secu", "guard", "protect", "dlp", "fsc", "scsk")
            hits = [n for n in names if any(k in n.lower() for k in keys)]
            lines.append(f"보안/DRM으로 보이는 프로세스 (이름에 drm, fasoo, softcamp, markany, secu, guard 등): {hits or '없음'}")
        except Exception as e:
            lines.append(f"프로세스 목록 실패: {e}")
    try:
        out = subprocess.run([sys.executable, "-m", "pip", "list", "--format=freeze", "--disable-pip-version-check"],
                             capture_output=True, timeout=120, env=child_env()).stdout.decode("utf-8", errors="replace")
        pkgs = [x for x in out.splitlines() if x.strip()]
        lines.append(f"설치된 패키지 {len(pkgs)}개: " + ", ".join(pkgs))
    except Exception as e:
        lines.append(f"pip list 실패: {e}")


def _broken(line):
    return "zip 형식=False" in line or "실패" in line or "Caption 있음=False" in line


def step_diag(stage, args, ctx):
    lines = [
        f"OS: {platform.platform()} ({platform.version()})",
        f"Python: {sys.version.split()[0]} ({sys.executable})",
        f"인코딩: 파일시스템 {sys.getfilesystemencoding()}, 기본 {locale.getpreferredencoding(False)}",
        f"CLI 폴더: {CLI_DIR} (도구: {tool_versions()}), 테스트 폴더: {HERE}",
        "",
        "[라이브러리 버전 / import]",
    ]
    for dist in PACKAGES:
        try:
            __import__(IMPORTS[dist])
            lines.append(f"  {dist} {_version(dist)} - import 정상")
            stage.passed += 1
        except Exception as e:
            lines.append(f"  {dist} {_version(dist)} - import 실패: {type(e).__name__}: {e}")
            if dist == "pywin32" and os.name != "nt":
                stage.skipped += 1
                stage.skips.append("pywin32: Windows 전용")
            else:
                stage.fail(f"{dist} import 실패: {type(e).__name__}: {e}"[:200])
    lines += ["", "[문서 보안(DRM) 영향 점검]"]
    for fn, what in ((diag_docx, "python-docx 기본 템플릿"), (diag_pptx, "python-pptx 기본 템플릿"),
                     (diag_new_files, "새로 쓴 파일 다시 읽기 (문서 보안)"), (diag_system, None)):
        n0 = len(lines)
        try:
            fn(lines)
        except Exception as e:
            lines.append(f"{fn.__name__} 오류: {type(e).__name__}: {e}")
            if what is None:
                continue
        if what is None:                                        # Office 버전, 프로세스, pip list: 참고 정보
            continue
        broken = [x for x in lines[n0:] if _broken(x) or " 오류: " in x]
        if broken:
            stage.fail(f"{what}: {broken[0][:200]}")
        else:
            stage.passed += 1
    lines += ["", "[OpenCode 스킬, 전역 지침 (apply_skills.py --check)]"]
    try:
        diag_skills(stage, lines)
    except Exception as e:
        lines.append(f"diag_skills 오류: {type(e).__name__}: {e}")
    text = "\n".join(lines)
    echo(text)
    stage.details.append(("진단 내용", text))


def diag_skills(stage, lines, config_dir=None):
    """OpenCode 설정 폴더의 SKILL.md, AGENTS.md가 C:\\CLI, C:\\CLI-dev와 같은지. 설정 폴더가 없으면(클라우드 등) 참고만.
    다르거나 옛 스킬 폴더(skills\\office-cli)가 남아 있으면 Gemma가 옛 지침으로 동작하므로 FAIL (고치는 명령: python apply_skills.py)."""
    import importlib.util
    spec = importlib.util.spec_from_file_location("apply_skills", os.path.join(HERE, "apply_skills.py"))
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    mod.HERE, mod.CLI_DIR = HERE, CLI_DIR
    cfg = config_dir or mod.default_config_dir()
    if not os.path.isdir(cfg):
        lines.append(f"  OpenCode 설정 폴더 없음 ({cfg}): 확인 안 함")
        return
    rows = mod.compare(cfg)
    lines += [f"  {mod.label(*r)}" for r in rows]
    bad = [r for r in rows if r[3] != "같음"]
    if bad:
        stage.fail("OpenCode 스킬/지침이 C:\\CLI와 다름 (" + ", ".join(f"{r[0]} {r[3]}" for r in bad)
                   + "). C:\\CLI-dev에서 python apply_skills.py 실행 후 OpenCode 다시 시작")
    else:
        stage.passed += 1


# ---------------------------------------------------------------------------
# 2. 오프라인 테스트
# ---------------------------------------------------------------------------
PYTEST_TARGETS = TOOLS + ("tests",)               # 도구별 오프라인 테스트 + CLI-dev/tests (채점기·점검 러너 등 개발 도구)
_PYTEST_COUNT = re.compile(r"(\d+) (passed|failed|skipped|errors?|xfailed|xpassed|deselected)\b")


def parse_pytest(out):
    """pytest -q -rfEs 출력에서 (통과, 실패, 건너뜀, 실패 줄, 건너뜀 줄)."""
    counts = {}
    for line in reversed(out.splitlines()):
        found = _PYTEST_COUNT.findall(line)
        if found and (" in " in line or line.strip().startswith("=")):
            for n, kind in found:
                counts[kind.rstrip("s") if kind.startswith("error") else kind] = int(n)
            break
    fails = [x.strip() for x in out.splitlines() if x.startswith(("FAILED ", "ERROR "))]
    skips = [x.strip() for x in out.splitlines() if x.startswith("SKIPPED ")]
    return (counts.get("passed", 0) + counts.get("xfailed", 0),
            counts.get("failed", 0) + counts.get("error", 0) + counts.get("xpassed", 0),
            counts.get("skipped", 0), fails, skips)


def step_pytest(stage, args, ctx):
    for tool in PYTEST_TARGETS:
        folder = os.path.join(HERE, tool)
        if not os.path.isdir(folder):
            continue
        say(f"{tool} 오프라인 테스트 시작 (몇 분 걸릴 수 있습니다)")
        code, out, sec, stop = run_capture([sys.executable, "-m", "pytest", "-q", "-p", "no:cacheprovider", "-rfEs"],
                                           folder, args.timeout)
        npass, nfail, nskip, fails, skips = parse_pytest(out)
        stage.passed += npass
        stage.skipped += nskip
        stage.failures += [f"{tool}: {x[:200]}" for x in fails]
        stage.failed += max(nfail, len(fails))
        if code not in (0, None) and not nfail and not fails:
            stage.fail(f"{tool}: pytest exit {code} (실패 항목 없이 끝남, 상세 참고)")
        stage.skips += [f"{tool}: {x[:200]}" for x in skips]
        stage.details.append((f"pytest {tool} (exit {code}, {sec}초)", f"cwd={folder}\n{out}"))
        if stop:
            raise KeyboardInterrupt


# ---------------------------------------------------------------------------
# 3. ml-kit 베타 점검 (run_beta_check.py)
# ---------------------------------------------------------------------------
BETA_SCRIPT = os.path.join(HERE, "run_beta_check.py")
GRADER = os.path.join(HERE, "grade_opencode.py")
_BETA_DETAIL = re.compile(r"^- (\S+) (.*) \[(FAIL|SKIP)\]: (.*)$")
_BETA_ROW = re.compile(r"^\| (\S+) \| (\S+) \| (.*) \| (PASS|FAIL|SKIP) \|$")


def _last_json(out):
    for line in reversed(out.splitlines()):
        line = line.strip()
        if line.startswith("{"):
            try:
                return json.loads(line)
            except ValueError:
                continue
    return None


def parse_beta_report(text):
    """beta_check_*.md에서 (통과, 실패 목록, 건너뜀 목록). 실패·건너뜀은 상세 줄(판정 근거)을 씁니다."""
    rows = [m.groups() for m in map(_BETA_ROW.match, text.splitlines()) if m]
    detail = {m.group(1): m.group(4) for m in map(_BETA_DETAIL.match, text.splitlines()) if m}
    npass = sum(r[3] == "PASS" for r in rows)
    fails = [f"{r[1]} {r[2]}" + (f": {detail[r[1]][:150]}" if detail.get(r[1]) else "") for r in rows if r[3] == "FAIL"]
    skips = [f"{r[1]} {r[2]}" + (f": {detail[r[1]][:150]}" if detail.get(r[1]) else "") for r in rows if r[3] == "SKIP"]
    return npass, fails, skips


def work_root():
    """모든 시험 폴더의 부모: 이 파일 옆 work (C:\\CLI-dev\\work). CLI-dev를 지우면 시험 흔적도 함께 사라집니다."""
    return os.path.join(HERE, "work")


def beta_work_dir(stamp):
    return os.path.join(work_root(), f"beta_{stamp}")


def step_beta(stage, args, ctx):
    work = beta_work_dir(ctx["stamp"])
    data_dir = os.path.join(work, "_data")                       # 점검이 일부러 낸 오류가 실제 오류 로그에 섞이지 않게
    cmd = [sys.executable, BETA_SCRIPT, "--agent", "--label", ctx["label"], "--dir", work]
    say("ml-kit 베타 점검 시작 (5~15분, 질문 없음). 엑셀 창이 잠깐 열릴 수 있습니다.")
    code, out, sec, stop = run_capture(cmd, HERE, args.timeout, env_extra={"ML_KIT_DATA_DIR": data_dir})
    payload = _last_json(out) or {}
    report = payload.get("report_file")
    stage.details.append((f"run_beta_check.py 화면 출력 (exit {code}, {sec}초)", f"$ {' '.join(cmd)}\n{out[-20000:]}"))
    if report and os.path.isfile(report):
        with open(report, encoding="utf-8", errors="replace") as f:
            text = f.read()
        npass, fails, skips = parse_beta_report(text)
        stage.passed += npass
        stage.failed += len(fails)
        stage.failures += fails
        stage.skipped += len(skips)
        stage.skips += skips
        stage.details.insert(0, ("점검 결과 (beta_check 보고서)", text))
        for p in (report, report[:-3] + ".xlsx"):                # 결과 파일은 이 파일 하나: 점검 보고서는 시험 폴더로 옮김
            try:
                if os.path.isfile(p) and os.path.isdir(work):
                    shutil.move(p, os.path.join(work, os.path.basename(p)))
            except Exception:
                pass
    else:
        stage.fail(f"점검 보고서 없음 (exit {code}, JSON {str(payload)[:150] or '없음'})")
    tables = sorted(glob.glob(os.path.join(work, "*_ml_data", "성능예측표_*.xlsx")))
    stage.note = f"시험 폴더 {work}" + (f", 성능예측표 {tables[-1]}" if tables else "")
    tool_log = os.path.join(data_dir, "logs", "errors.txt")
    if os.path.isfile(tool_log):
        _attach(stage, "점검 중 생긴 ml-kit 오류 로그 (일부러 낸 오류 포함)", tool_log)
    if stop:
        raise KeyboardInterrupt


# ---------------------------------------------------------------------------
# 4. 실기 러너 (Windows + 실제 Office)
# ---------------------------------------------------------------------------
def windows_ready():
    """실기 러너를 돌릴 수 있는지: (가능 여부, 안 되는 이유)."""
    if os.name != "nt":
        return False, "Windows가 아님 (실제 Office가 필요한 단계)"
    try:
        __import__("win32com.client")
    except Exception:
        return False, "pywin32 없음 (python -m pip install pywin32)"
    return True, ""


def _newer(pattern, since):
    files = [p for p in glob.glob(pattern) if os.path.getmtime(p) >= since - 1]
    return sorted(files, key=os.path.getmtime)


def _attach(stage, name, path):
    try:
        with open(path, encoding="utf-8", errors="replace") as f:
            stage.details.append((name, f"파일: {path}\n{f.read()}"))
    except Exception as e:
        stage.details.append((name, f"파일 {path} 읽기 실패: {e}"))


def parse_counts(text, prefix="결과:"):
    """'결과: ENV-SKIP=2, PASS=69' 줄을 {상태: 수}로."""
    for line in text.splitlines():
        if line.strip().startswith(prefix):
            return {k: int(v) for k, v in re.findall(r"([A-Z][A-Z-]*)=(\d+)", line)}
    return {}


def _add_counts(stage, counts):
    stage.passed += counts.get("PASS", 0)
    stage.skipped += counts.get("ENV-SKIP", 0) + counts.get("MANUAL-SKIP", 0) + counts.get("SKIP", 0)
    return counts.get("FAIL", 0) + counts.get("ERROR", 0)


def parse_office_logs(stage, errors_text, report_text):
    """office-cli 러너의 errors_*.txt(요약)와 report_*.txt(UNVERIFIED 목록)를 단계 결과에 넣습니다."""
    nbad = _add_counts(stage, parse_counts(errors_text))
    problems, skipped = [], []
    for line in errors_text.splitlines():
        s = line.strip()
        if s.startswith("문제 항목:") and "없음" != s.split(":", 1)[1].strip():
            problems = [x.strip() for x in s.split(":", 1)[1].split(", ") if x.strip()]
        elif s.startswith("건너뜀:"):
            skipped = [x.strip() for x in re.split(r"\), (?=[A-Z]+\d*\()", s.split(":", 1)[1]) if x.strip()]
    stage.failures += problems
    stage.failed += max(nbad, len(problems))
    stage.skips += [x if x.endswith(")") else x + ")" for x in skipped]
    lines = report_text.splitlines()
    for i, line in enumerate(lines):
        if line.startswith("UNVERIFIED 결과"):
            cur = None
            for x in lines[i + 1:]:
                if x.startswith("        unverified:") and cur:
                    stage.unverified.append(f"{cur[:150]} -> {x.split(':', 1)[1].strip()[:200]}")
                    cur = None
                elif x.startswith("    ") and x.strip():
                    cur = x.strip()
            break


def step_office(stage, args, ctx):
    ok, why = windows_ready()
    if not ok:
        stage.env_skip(why)
        return
    folder = os.path.join(HERE, "office-cli")
    start = time.time()
    cmd = [sys.executable, os.path.join("tools", "run_win_tests.py")]
    if not args.outlook:
        cmd.append("--skip-outlook")
    if not args.ask:
        cmd.append("--auto")
    if args.office_dir:
        cmd += ["--dir", args.office_dir]
    say("office-cli 실기 러너 시작 (10~30분). Office 창을 만지지 마세요."
        + (" 질문이 나오면 답하세요." if args.ask else " 질문 없이 진행합니다."))
    if args.ask:                                                # 질문에 답할 수 있게 화면을 그대로 씀
        t0, stop = time.time(), False
        try:
            code = subprocess.call(cmd, cwd=folder, env=child_env())
        except KeyboardInterrupt:
            code, stop = 130, True
        out, sec = "(화면 출력은 기록하지 않음: --ask)", round(time.time() - t0, 1)
    else:
        code, out, sec, stop = run_capture(cmd, folder, args.timeout)
    stage.details.append((f"러너 화면 출력 (exit {code}, {sec}초)", f"$ {' '.join(cmd)}\n{out[-20000:]}"))
    test_dir = args.office_dir or os.path.join(work_root(), "office")
    errs = _newer(os.path.join(test_dir, "errors_*.txt"), start)
    reps = _newer(os.path.join(test_dir, "report_*.txt"), start)
    if errs:
        read = lambda p: open(p, encoding="utf-8", errors="replace").read()
        parse_office_logs(stage, read(errs[-1]), read(reps[-1]) if reps else "")
        for p in errs[-1:] + reps[-1:]:
            _attach(stage, os.path.basename(p), p)
    else:
        stage.fail(f"러너 결과 파일 없음 (exit {code}, {test_dir}\\errors_*.txt)")
    if stop:
        raise KeyboardInterrupt


def step_ml(stage, args, ctx):
    ok, why = windows_ready()
    if not ok:
        stage.env_skip(why)
        return
    folder = os.path.join(HERE, "ml-kit")
    start = time.time()
    cmd = [sys.executable, os.path.join("tools", "run_win_tests.py"), "--label", ctx["label"]]
    if args.ml_dir:
        cmd += ["--dir", args.ml_dir]
    say("ml-kit 실기 러너 시작 (10~20분, 질문 없음). Office 창을 만지지 마세요.")
    code, out, sec, stop = run_capture(cmd, folder, args.timeout)
    stage.details.append((f"러너 화면 출력 (exit {code}, {sec}초)", f"$ {' '.join(cmd)}\n{out[-20000:]}"))
    logs = _newer(os.path.join(folder, "tools", "test_logs", "testlog_*.txt"), start)
    if logs:
        text = open(logs[-1], encoding="utf-8", errors="replace").read()
        _add_counts(stage, parse_counts(text))
        for m in re.finditer(r"^  (FAIL|ERROR)\s+(\S+)\s+\S+s\s+(.*)$", text, re.M):
            stage.fail(f"{m.group(2)} {m.group(3)} ({m.group(1)})")
        _attach(stage, os.path.basename(logs[-1]), logs[-1])
    else:
        stage.fail(f"러너 로그 없음 (exit {code})")
    test_dir = args.ml_dir or os.path.join(work_root(), "ml 테스트")
    for p in _newer(os.path.join(test_dir, "report_*.txt"), start):
        _attach(stage, os.path.basename(p), p)
    if stop:
        raise KeyboardInterrupt


# ---------------------------------------------------------------------------
# 도구 오류 로그 (실제 사용 중 생긴 오류. 실기 러너가 일부러 낸 오류는 각 단계 상세에)
# ---------------------------------------------------------------------------
def attach_tool_error_logs(rep):
    found = []
    for tool, env in (("ml-kit", "ML_KIT_DATA_DIR"), ("office-cli", "OFFICE_CLI_DATA_DIR")):
        data = os.environ.get(env) or os.path.join(CLI_DIR, tool, "data")
        n = 0
        for name in ("errors.old.txt", "errors.txt"):
            p = os.path.join(data, "logs", name)
            if not os.path.exists(p):
                continue
            try:
                with open(p, encoding="utf-8", errors="replace") as f:
                    text = f.read()
            except Exception as e:
                text = f"읽기 실패: {e}"
            n += sum(1 for x in text.splitlines() if x.startswith("===== "))
            rep.logs.append((f"{tool} {name}", f"파일: {p}\n{text}"))
        found.append(f"{tool} {n}건")
    rep.log_note = ", ".join(found) + (" (파일 없음)" if not rep.logs else "")


# ---------------------------------------------------------------------------
# 6. Gemma 종합 테스트 (opencode run)
# ---------------------------------------------------------------------------
def find_opencode(given=None):
    """opencode 실행 파일: --opencode, PATH, C:\\CLI\\opencode.exe(이 폴더 옆 CLI) 순서. 없으면 None."""
    if given:
        return given if os.path.isfile(given) else None
    found = shutil.which("opencode")
    if found:
        return found
    for name in ("opencode.exe", "opencode"):
        p = os.path.join(CLI_DIR, name)
        if os.path.isfile(p):
            return p
    return None


def is_elevated():
    """Windows 관리자 권한으로 실행 중인지 (관리자 창의 opencode는 일반 권한으로 연 Office 문서를 COM으로 못 봄)."""
    if os.name != "nt":
        return False
    try:
        import ctypes
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def run_to_files(args, cwd, timeout, out_path, err_path, poll=1.0):
    """stdout은 out_path에 바이트 그대로(opencode가 쓴 UTF-8), stderr는 err_path에 저장합니다.
    (종료코드, 걸린 초, Ctrl+C 여부). timeout초가 지나면 끝냅니다 (run_capture와 같은 규칙)."""
    start = time.time()
    interrupted = False
    with open(out_path, "wb") as out, open(err_path, "wb") as err:
        proc = subprocess.Popen(args, cwd=cwd, env=child_env(), stdin=subprocess.DEVNULL, stdout=out, stderr=err)
        try:
            while True:
                try:
                    proc.wait(timeout=poll)
                    break
                except subprocess.TimeoutExpired:
                    if time.time() - start > timeout:
                        kill_tree(proc)
                        err.write(f"\n[ERROR] 시간 초과 ({timeout}초)로 중단했습니다.\n".encode("utf-8"))
                        break
        except KeyboardInterrupt:
            interrupted = True
            kill_tree(proc)
    return proc.returncode, round(time.time() - start, 1), interrupted


def _load_grader():
    import importlib.util
    spec = importlib.util.spec_from_file_location("grade_opencode", GRADER)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def _read(path, limit=4000):
    try:
        with open(path, encoding="utf-8", errors="replace") as f:
            return f.read()[-limit:]
    except OSError:
        return ""


def step_gemma(stage, args, ctx):
    exe = find_opencode(args.opencode)
    if not exe:
        stage.env_skip("opencode 없음 (--opencode <경로>, PATH, 또는 C:\\CLI\\opencode.exe)")
        return
    try:
        go = _load_grader()
        spec, by_id, order = go.load_scenarios()
    except Exception as e:                                       # PyYAML 없음 등
        stage.env_skip(f"시나리오를 읽지 못함 ({type(e).__name__}: {e})")
        return
    test_dir = args.gemma_dir or os.path.join(work_root(), "gemma")
    if os.path.exists(test_dir):                                 # 이전 모델이 남으면 버전 번호가 달라짐: 지우지 않고 이름만 바꿈
        old = f"{test_dir}_old_{ctx['stamp']}"
        try:
            os.rename(test_dir, old)
            say(f"이전 시험 폴더를 {old}(으)로 옮겼습니다.")
        except OSError as e:
            stage.fail(f"이전 시험 폴더를 옮기지 못함 ({test_dir}: {e}). 열린 파일을 닫고 다시 실행하세요.")
            return
    runs = os.path.join(work_root(), f"gemma_runs_{ctx['stamp']}")
    os.makedirs(runs, exist_ok=True)
    if is_elevated():
        stage.unverified.append("관리자 권한 창에서 실행함: opencode가 열린 Office 문서를 못 볼 수 있음 (일반 PowerShell 권장)")
    sent, session, log = {}, None, []
    say(f"Gemma 종합 테스트 시작: 요청 {len(order)}개를 opencode run으로 한 세션에 보냅니다 (요청마다 최대 {args.gemma_timeout}초).")
    stop = False
    for i, sid in enumerate(order, 1):
        msg = f"[{sid}] " + str(by_id[sid]["message"]).replace("{DIR}", test_dir)
        sent[sid] = msg
        with open(os.path.join(runs, go.RUN_MESSAGES), "w", encoding="utf-8") as f:
            json.dump(sent, f, ensure_ascii=False, indent=1)
        cmd = [exe, "run", "--format", "json"] + (["--session", session] if session else []) + [msg]
        out_path, err_path = os.path.join(runs, f"{sid}.jsonl"), os.path.join(runs, f"{sid}.stderr.txt")
        say(f"[{i}/{len(order)}] {sid} 보내는 중")
        code, sec, stop = run_to_files(cmd, CLI_DIR, args.gemma_timeout, out_path, err_path)
        _, got, errors = go.read_run(out_path, sid, msg)
        session = session or got
        log.append(f"{sid}: exit {code}, {sec}초, 세션 {got or '없음'}" + (f", 오류 {'; '.join(errors)}" if errors else ""))
        if stop:
            break
        if not session:                                          # 첫 요청에서 세션이 안 생기면 나머지도 같음
            stage.fail(f"opencode가 응답하지 않음 ({sid}, exit {code}). 모델 설정(opencode.jsonc의 model)과 접속을 확인하세요.")
            stage.details.append((f"{sid} opencode 오류 출력", _read(err_path) or "(없음)"))
            break
    stage.details.append(("opencode run 실행 기록", "\n".join(log)))
    if session:
        meta, rows, detail = go.grade_runs(runs)
        report, npass, n = go.write_report(ctx["label"], meta, rows, detail, out_dir=runs)
        for r in rows:
            if r["결과"] in ("통과", "통과(참고)"):
                stage.passed += 1
            elif r["결과"] == "실패(참고)":
                stage.skipped += 1
                stage.skips.append(f"{r['id']}: 자유 문장 답이라 사람이 확인 (아래 '최종 답변')")
            elif r["결과"] == "없음":
                stage.skipped += 1
                stage.skips.append(f"{r['id']}: 보내지 못함 (중단 또는 앞 요청 문제)")
            else:
                stage.fail(f"{r['id']}: {r['실패'][:200]}")
        stage.details.insert(0, ("채점 결과 (grade_opencode)", _read(report, limit=60000)))
    stage.note = f"세션 {session or '없음'}, 기록 폴더 {runs}, 시험 폴더 {test_dir}"
    if stop:
        raise KeyboardInterrupt


# ---------------------------------------------------------------------------
RUNNERS = {"diag": step_diag, "pytest": step_pytest, "beta": step_beta, "ml": step_ml, "office": step_office,
           "gemma": step_gemma}


def main(argv=None):
    p = argparse.ArgumentParser(description="CLI 도구 실기 점검 (결과는 파일 1개: result_<라벨>_<시각>.md)")
    p.add_argument("--label", default="PC", help="결과 파일 이름에 붙일 PC 라벨 (예: 회사PC)")
    p.add_argument("--only", help="실행할 단계 (쉼표로): diag,pytest,beta,ml,office,gemma (기본: diag,pytest,beta,office)")
    p.add_argument("--gemma", action="store_true", help="기본 단계 끝에 Gemma 종합 테스트(opencode run)를 더함")
    p.add_argument("--opencode", help=r"opencode 실행 파일 (기본: PATH의 opencode, 없으면 C:\CLI\opencode.exe)")
    p.add_argument("--gemma-dir", help=r"gemma 단계 시험 폴더 (기본 C:\CLI-dev\work\gemma)")
    p.add_argument("--gemma-timeout", type=int, default=900, help="gemma 단계 요청 하나의 시간 상한(초, 기본 900)")
    p.add_argument("--ask", action="store_true", help="office 단계의 눈 확인 질문(y/n)에 답하면서 진행 (기본: 질문 없음)")
    p.add_argument("--outlook", action="store_true", help="office 단계에서 Outlook(G그룹)도 시험 (기본: 건너뜀)")
    p.add_argument("--ml-dir", help=r"ml 단계 시험 폴더 (기본 C:\CLI-dev\work\ml 테스트)")
    p.add_argument("--office-dir", help=r"office 단계 시험 폴더 (기본 C:\CLI-dev\work\office)")
    p.add_argument("--timeout", type=int, default=3600, help="단계별 시간 상한(초, 기본 3600)")
    args = p.parse_args(argv)
    chosen = [s.strip().lower() for s in (args.only or ",".join(DEFAULT_STAGES)).split(",") if s.strip()]
    if args.gemma and "gemma" not in chosen:
        chosen.append("gemma")
    bad = [s for s in chosen if s not in STAGES]
    if bad or not chosen:
        p.error(f"알 수 없는 단계: {', '.join(bad) or '(비어 있음)'} (가능: {', '.join(STAGES)})")
    stages = [s for s in STAGES if s in chosen]                  # 늘 같은 순서

    label = "".join(c for c in args.label if c.isalnum() or c in "-_") or "PC"
    stamp = f"{dt.datetime.now():%Y%m%d_%H%M%S}"
    out = os.path.join(HERE, f"result_{label}_{stamp}.md")
    rep = Report(out, label, stages)
    ctx = {"label": label, "stamp": stamp}
    say(f"결과 파일: {out}")
    say("단계: " + " -> ".join(stages) + " -> 도구 오류 로그")
    if "gemma" in stages:
        say("gemma 단계는 opencode로 Gemma에게 요청을 보냅니다. 관리자 권한이 아닌 일반 PowerShell에서 실행하세요.")
    if any(s in stages for s in ("beta", "ml", "office")):
        say("점검 중 Excel/Word/PowerPoint 창이 열리고 닫힙니다. 열린 Office 문서는 모두 저장하고 닫은 뒤 시작하세요.")
    code = 0
    try:
        for stage in rep.stages:
            t0 = time.time()
            say(f"{stage.key} 단계 시작: {stage.title}")
            try:
                RUNNERS[stage.key](stage, args, ctx)
            except KeyboardInterrupt:
                stage.status = "중단"
                raise
            except Exception:
                stage.fail(f"단계 오류 (러너 문제): {traceback.format_exc().strip().splitlines()[-1][:200]}")
                stage.details.append(("단계 오류", traceback.format_exc()))
                say(f"{stage.key} 단계에서 오류가 났습니다. 다음 단계로 넘어갑니다.")
            finally:
                stage.seconds = round(time.time() - t0)
                stage.finish()
                rep.write()
            say(f"{stage.key} 단계 끝: {stage.status} ({stage.seconds}초)")
    except KeyboardInterrupt:
        rep.interrupted = True
        code = 130
        say("중단했습니다.")
    finally:
        try:
            attach_tool_error_logs(rep)
        except Exception as e:
            rep.log_note = f"모으기 실패: {type(e).__name__}: {e}"
        rep.ended = dt.datetime.now()
        rep.write()
        echo("\n" + "=" * 70 + "\n[STATUS] 요약\n" + "\n".join(rep.summary_lines()) + "\n" + "=" * 70)
        say(f"끝. 이 파일 하나를 올리세요: {out}")
    if code:
        return code
    return 1 if any(s.status == "FAIL" for s in rep.stages) else 0


if __name__ == "__main__":
    sys.exit(main())
