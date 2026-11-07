r"""C:\CLI 설치 (사람이 실행: setup.bat 두 번 클릭, 또는 python C:\CLI\setup.py)

배포는 C:\CLI 폴더 하나입니다. 이 스크립트가 그 폴더를 사용자 PC에 연결합니다.
    python setup.py                 # 설치 (바뀐 것만 복사, PATH 등록, 동작 확인)
    python setup.py --check         # 아무것도 바꾸지 않고 비교·확인만 (다르거나 실패가 있으면 exit 1)
    python setup.py --no-path       # 시스템 PATH는 건드리지 않음
    python setup.py --config-dir D:\oc   # OpenCode 설정 폴더가 기본 위치가 아닐 때

하는 일 (C:\CLI = 이 파일이 있는 폴더):
1. OpenCode 설정 폴더(%USERPROFILE%\.config\opencode)로 복사 (바뀐 것만):
       ml-kit\skills\<이름>\SKILL.md       -> skills\<이름>\SKILL.md  (ml-train, ml-predict, ml-report, ml-map)
       office-cli\skills\<이름>\SKILL.md   -> skills\<이름>\SKILL.md  (office-excel, office-doc, office-mail, office-spec, office-data)
       opencode\AGENTS.md                  -> AGENTS.md (전역 지침)
       opencode\opencode.jsonc             -> opencode.jsonc (있을 때만. 사내 LLM 주소가 들어 있어 사용자가 넣어 둠)
   덮어쓰는 파일은 data\setup_backup_<시각>\config\에 먼저 보관합니다.
   옛 통합 스킬 폴더 skills\office-cli, skills\ml-kit은 보관 후 지웁니다 (링크·정션이면 링크만 지움).
2. 시스템 PATH(HKLM)에 <C:\CLI>\python, <C:\CLI>\python\Scripts가 없으면 뒤에 추가합니다.
   원래 값은 data\setup_backup_<시각>\system_path.txt에 보관하고, 쓴 뒤 다시 읽어 확인하고, 환경 변경을 알립니다.
3. 동작 확인: python\python.exe로 라이브러리 import, 도구 -h가 stdout 1줄 JSON인지, opencode.exe, rg(ripgrep).

개발용: --test-skill <SKILL.md>  (실기 테스트 보조 스킬, 예: C:\CLI-dev\office-cli\tools\office-cli-test\SKILL.md)
이 스크립트는 에이전트용 출력 계약(stdout 1줄 JSON) 대상이 아닙니다. exit 0 = 모두 OK, 1 = 실패(또는 --check에서 다름) 있음.
"""
import argparse
import datetime as dt
import json
import os
import re
import shutil
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))                  # C:\CLI (배포 폴더)
OFFICE_SKILLS = ("office-excel", "office-doc", "office-mail", "office-spec", "office-data")    # office-cli\skills\<이름>\SKILL.md
ML_SKILLS = ("ml-train", "ml-predict", "ml-report", "ml-map")                     # ml-kit\skills\<이름>\SKILL.md
OBSOLETE = (("옛 office-cli 스킬", os.path.join("skills", "office-cli")),           # 설정 폴더에서 지울 옛 스킬 폴더
            ("옛 ml-kit 스킬", os.path.join("skills", "ml-kit")))
ENV_KEY = r"SYSTEM\CurrentControlSet\Control\Session Manager\Environment"     # 시스템 환경 변수 (HKLM)
# 도구가 쓰는 라이브러리 (AGENTS.md 사용자 환경 표): (배포 이름, import 이름)
LIBS = (("numpy", "numpy"), ("pandas", "pandas"), ("scikit-learn", "sklearn"), ("scipy", "scipy"), ("joblib", "joblib"),
        ("openpyxl", "openpyxl"), ("python-docx", "docx"), ("python-pptx", "pptx"), ("matplotlib", "matplotlib"),
        ("pymupdf", "pymupdf"), ("olefile", "olefile"), ("pywin32", "win32com.client"))
TOOL_HELP = (("ml-kit", os.path.join("ml-kit", "ml_predict.py")), ("office-cli", os.path.join("office-cli", "office.py")))
LONG_PATH = 2047                                                     # 이보다 길면 일부 프로그램이 뒤쪽 PATH를 못 봄


def default_config_dir():
    return os.path.join(os.path.expanduser("~"), ".config", "opencode")


def data_dir():
    return os.path.join(HERE, "data")


def python_exe():
    """배포에 든 WinPython (없으면 지금 실행 중인 python)."""
    exe = os.path.join(HERE, "python", "python.exe")
    return exe if os.path.isfile(exe) else sys.executable


def path_entries():
    return [os.path.join(HERE, "python"), os.path.join(HERE, "python", "Scripts")]


class Report:
    """[OK]/[WARN]/[FAIL] 한 줄씩 출력하고 개수를 셉니다. diff는 --check에서 '다름'으로 센 것."""

    def __init__(self):
        self.ok = self.warn = 0
        self.fails, self.diffs = [], []

    def line(self, tag, msg):
        print(f"[{tag}] {msg}", flush=True)
        if tag == "OK":
            self.ok += 1
        elif tag == "WARN":
            self.warn += 1
        elif tag == "FAIL":
            self.fails.append(msg)

    def diff(self, msg):
        self.line("WARN", msg)
        self.diffs.append(msg)


# ---------------------------------------------------------------------------
# 1. 스킬·전역 지침·opencode.jsonc
# ---------------------------------------------------------------------------
def targets(test_skills=()):
    """(이름, 원본, 설정 폴더 기준 대상 상대 경로) 목록."""
    items = [(f"{name} 스킬", os.path.join(HERE, "ml-kit", "skills", name, "SKILL.md"), os.path.join("skills", name, "SKILL.md"))
             for name in ML_SKILLS]
    items += [(f"{name} 스킬", os.path.join(HERE, "office-cli", "skills", name, "SKILL.md"), os.path.join("skills", name, "SKILL.md"))
              for name in OFFICE_SKILLS]
    items.append(("전역 지침", os.path.join(HERE, "opencode", "AGENTS.md"), "AGENTS.md"))
    for src in test_skills:
        src = os.path.abspath(src)
        if os.path.isdir(src):
            src = os.path.join(src, "SKILL.md")
        name = os.path.basename(os.path.dirname(src))
        items.append((f"{name} 스킬(개발용)", src, os.path.join("skills", name, "SKILL.md")))
    return items


def jsonc_source():
    return os.path.join(HERE, "opencode", "opencode.jsonc")


def _read(path):
    try:
        with open(path, "rb") as f:
            return f.read()
    except OSError:
        return None


def version_of(data):
    """SKILL.md 머리말의 version 값 (없으면 빈 글)."""
    m = re.search(rb'^\s*version:\s*"?([^"\r\n]+)"?', data or b"", re.M)
    return m.group(1).decode("utf-8", "replace").strip() if m else ""


def compare(config_dir, test_skills=()):
    """[(이름, 원본, 대상, 상태)]. 상태: 같음 / 다름 / 새 파일(대상 없음) / 원본 없음 / 지울 것(옛 스킬 폴더, 원본은 None)."""
    out = []
    for name, src, rel in targets(test_skills):
        dst = os.path.join(config_dir, rel)
        a, b = _read(src), _read(dst)
        state = "원본 없음" if a is None else "새 파일" if b is None else "같음" if a == b else "다름"
        out.append((name, src, dst, state))
    for name, rel in OBSOLETE:
        dst = os.path.join(config_dir, rel)
        if os.path.lexists(dst):                       # 대상이 없어진 링크도 지울 것
            out.append((name, None, dst, "지울 것"))
    return out


def jsonc_state(config_dir):
    """opencode.jsonc: (원본, 대상, 상태). 상태: 같음 / 다름 / 새 파일 / 설정 폴더 것 사용(원본 없음, 대상 있음) / 없음(둘 다 없음)."""
    src, dst = jsonc_source(), os.path.join(config_dir, "opencode.jsonc")
    a, b = _read(src), _read(dst)
    if a is None:
        return src, dst, "설정 폴더 것 사용" if b is not None else "없음"
    return src, dst, "새 파일" if b is None else "같음" if a == b else "다름"


def _is_link(path):
    """심볼릭 링크나 Windows 정션(mklink /J)인지. 링크면 링크만 지우고 가리키는 폴더는 건드리지 않습니다."""
    isjunction = getattr(os.path, "isjunction", None)          # Python 3.12+
    return os.path.islink(path) or bool(isjunction and isjunction(path))


def _remove_link(path):
    try:
        os.unlink(path)
    except (IsADirectoryError, PermissionError):
        os.rmdir(path)                                  # Windows 디렉터리 링크·정션은 rmdir로 링크만 지움 (대상 내용은 그대로)


def label(name, src, dst, state):
    if src is None:
        return f"{name}: {state}  {dst}"
    a, b = version_of(_read(src)), version_of(_read(dst))
    ver = f" ({b or '?'} -> {a})" if state == "다름" and a else (f" ({a})" if a else "")
    return f"{name}: {state}{ver}  {dst}"


def _copy(src, dst, config_dir, backup, state):
    """다르면 보관 후 덮어쓰고, 다시 읽어 같은지 확인합니다."""
    if state == "다름":
        keep = os.path.join(backup, "config", os.path.relpath(dst, config_dir))
        os.makedirs(os.path.dirname(keep), exist_ok=True)
        shutil.copy2(dst, keep)
    os.makedirs(os.path.dirname(dst), exist_ok=True)
    tmp = dst + ".tmp"
    shutil.copyfile(src, tmp)
    os.replace(tmp, dst)
    return _read(dst) == _read(src)


def _remove_obsolete(name, dst, config_dir, backup, rep):
    keep = os.path.join(backup, "config", os.path.relpath(dst, config_dir))
    err, done = "", " -> 보관 후 지움"
    try:
        if _is_link(dst):
            # 링크·정션: 링크만 지움 (가리키는 폴더는 사용자의 것일 수 있어 지우지도 보관하지도 않음)
            target = os.path.realpath(dst)
            _remove_link(dst)
            done = f" -> 링크만 지움 (가리키던 폴더 {target}는 그대로)"
        elif os.path.isdir(dst):
            shutil.copytree(dst, keep, dirs_exist_ok=True)
            shutil.rmtree(dst)
        else:
            os.makedirs(os.path.dirname(keep), exist_ok=True)
            shutil.copy2(dst, keep)
            os.remove(dst)
    except PermissionError as e:                        # OpenCode가 파일을 잡고 있으면: 닫고 다시 실행
        err = f" ({type(e).__name__}: {e}. OpenCode를 닫고 다시 실행)"
    except OSError as e:
        err = f" ({type(e).__name__}: {e}. 이 폴더를 직접 지우거나 옮긴 뒤 다시 실행)"
    if os.path.lexists(dst):
        rep.line("FAIL", f"{name}: {dst} -> 지우지 못함{err}")
    else:
        rep.line("OK", f"{name}: {dst}{done}")


def step_files(rep, config_dir, backup, check=False, test_skills=()):
    print(f"[INFO] OpenCode 설정 폴더: {config_dir}", flush=True)
    if not check and not os.path.isdir(config_dir):
        rep.line("WARN", "설정 폴더가 없어 새로 만듭니다 (OpenCode를 한 번도 실행하지 않았거나 다른 위치일 수 있음: --config-dir)")
    for name, src, dst, state in compare(config_dir, test_skills):
        if state == "원본 없음":
            rep.line("FAIL", f"{name}: 원본이 없습니다: {src}")
        elif state == "같음":
            rep.line("OK", label(name, src, dst, state))
        elif check:
            rep.diff(label(name, src, dst, state))
        elif state == "지울 것":
            _remove_obsolete(name, dst, config_dir, backup, rep)
        else:
            ok = _copy(src, dst, config_dir, backup, state)
            rep.line("OK" if ok else "FAIL", label(name, src, dst, state) + (" -> 복사함" if ok else " -> 복사 후 내용이 다름"))
    src, dst, state = jsonc_state(config_dir)
    if state == "같음":
        rep.line("OK", f"opencode.jsonc: 같음  {dst}")
    elif state == "설정 폴더 것 사용":
        rep.line("WARN", f"opencode.jsonc: {src}가 없어 설정 폴더의 것을 그대로 사용합니다  {dst}")
    elif state == "없음":
        rep.line("WARN", f"opencode.jsonc: {src}도, 설정 폴더({dst})에도 없습니다. "
                         "사내 LLM 주소가 든 opencode.jsonc를 C:\\CLI\\opencode\\에 넣고 다시 실행하세요")
    elif check:
        rep.diff(f"opencode.jsonc: {state}  {dst}")
    else:
        ok = _copy(src, dst, config_dir, backup, state)
        rep.line("OK" if ok else "FAIL", f"opencode.jsonc: {state}  {dst}" + (" -> 복사함" if ok else " -> 복사 후 내용이 다름"))


# ---------------------------------------------------------------------------
# 2. 시스템 PATH (HKLM, winreg. setx는 1024자에서 잘라 쓰므로 쓰지 않음)
# ---------------------------------------------------------------------------
def _winreg():
    try:
        import winreg
        return winreg
    except ImportError:                                 # Windows가 아님
        return None


def _norm(entry):
    """중복 판정용: 따옴표·앞뒤 공백·끝 \\ 무시, 대소문자 무시, 환경 변수(%SystemRoot% 등) 풀어서."""
    e = os.path.expandvars(entry.strip().strip('"'))
    return e.replace("/", "\\").rstrip("\\").lower()


def merge_path(value, wanted):
    """(새 값, 추가한 항목). 없을 때만 뒤에 붙이고, 원래 값의 순서·내용은 글자 그대로 둡니다."""
    have = {_norm(e) for e in value.split(";") if e.strip()}
    add = []
    for w in wanted:
        if _norm(w) not in have:
            add.append(w)
            have.add(_norm(w))
    if not add:
        return value, []
    sep = "" if not value or value.endswith(";") else ";"
    return value + sep + ";".join(add), add


def read_system_path(reg):
    """(값, 형식). 값이 없으면 ("", REG_EXPAND_SZ)."""
    with reg.OpenKey(reg.HKEY_LOCAL_MACHINE, ENV_KEY, 0, reg.KEY_READ) as key:
        try:
            value, kind = reg.QueryValueEx(key, "Path")
        except FileNotFoundError:
            return "", reg.REG_EXPAND_SZ
    if kind not in (reg.REG_SZ, reg.REG_EXPAND_SZ):
        kind = reg.REG_EXPAND_SZ
    return str(value), kind


def missing_path_entries(reg=None):
    """PATH 확인만 (쓰지 않음): 시스템 PATH에 없는 항목 목록. Windows가 아니면 None."""
    reg = reg or _winreg()
    if reg is None:
        return None
    return merge_path(read_system_path(reg)[0], path_entries())[1]


def broadcast_env_change():
    """열린 프로그램(탐색기 등)에 환경 변수가 바뀌었다고 알립니다 (WM_SETTINGCHANGE "Environment")."""
    import ctypes
    from ctypes import wintypes
    fn = ctypes.windll.user32.SendMessageTimeoutW
    fn.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPCWSTR, wintypes.UINT, wintypes.UINT,
                   ctypes.POINTER(ctypes.c_size_t)]
    fn.restype = wintypes.LPARAM
    result = ctypes.c_size_t(0)
    return bool(fn(0xFFFF, 0x001A, 0, "Environment", 0x0002, 5000, ctypes.byref(result)))   # HWND_BROADCAST, SMTO_ABORTIFHUNG


def step_path(rep, backup, check=False, reg=None):
    reg = reg or _winreg()
    if reg is None:
        rep.line("WARN", "Windows가 아니라 시스템 PATH는 확인하지 않았습니다")
        return
    wanted = path_entries()
    try:
        value, kind = read_system_path(reg)
    except OSError as e:
        rep.line("FAIL", f"시스템 PATH를 읽지 못함 ({type(e).__name__}: {e})")
        return
    new, add = merge_path(value, wanted)
    if not add:
        rep.line("OK", "시스템 PATH: 이미 있음 (" + ", ".join(wanted) + ")")
        return
    if check:
        rep.diff("시스템 PATH에 없음: " + ", ".join(add))
        return
    try:                                                # 쓰기 전에 원래 값을 보관 (보관하지 못하면 쓰지 않음)
        os.makedirs(backup, exist_ok=True)
        keep = os.path.join(backup, "system_path.txt")
        with open(keep, "w", encoding="utf-8") as f:
            f.write(f"# HKLM\\{ENV_KEY} Path ({'REG_EXPAND_SZ' if kind == reg.REG_EXPAND_SZ else 'REG_SZ'}), 설치 전 값\n{value}\n")
    except OSError as e:
        rep.line("FAIL", f"시스템 PATH 원래 값을 보관하지 못해 바꾸지 않았습니다 ({type(e).__name__}: {e})")
        return
    try:
        with reg.OpenKey(reg.HKEY_LOCAL_MACHINE, ENV_KEY, 0, reg.KEY_READ | reg.KEY_SET_VALUE) as key:
            reg.SetValueEx(key, "Path", 0, kind, new)
    except PermissionError:
        rep.line("FAIL", "시스템 PATH를 쓰지 못함 (권한 없음). setup.bat을 마우스 오른쪽 버튼 -> '관리자 권한으로 실행'으로 다시 실행하세요")
        return
    except OSError as e:
        rep.line("FAIL", f"시스템 PATH를 쓰지 못함 ({type(e).__name__}: {e})")
        return
    after, kind2 = read_system_path(reg)
    if after != new or kind2 != kind:
        rep.line("FAIL", f"시스템 PATH를 쓴 뒤 다시 읽은 값이 다릅니다 (원래 값: {keep})")
        return
    rep.line("OK", "시스템 PATH에 추가: " + ", ".join(add) + f" (원래 값 보관: {keep})")
    if len(new) > LONG_PATH:
        rep.line("WARN", f"시스템 PATH가 {len(new)}자입니다. 일부 프로그램은 {LONG_PATH}자 뒤의 항목을 못 볼 수 있습니다")
    try:
        sent = broadcast_env_change()
    except Exception as e:                              # 알림 실패는 PATH 값에 영향 없음
        sent = False
        print(f"[INFO] 환경 변경 알림 오류: {type(e).__name__}: {e}", flush=True)
    if not sent:
        rep.line("WARN", "환경 변경 알림을 보내지 못했습니다. 새 PATH가 안 보이면 로그아웃 후 다시 로그인하세요")


# ---------------------------------------------------------------------------
# 3. 동작 확인
# ---------------------------------------------------------------------------
def _run(cmd, timeout):
    env = dict(os.environ, PYTHONIOENCODING="utf-8")
    p = subprocess.run(cmd, capture_output=True, timeout=timeout, env=env, cwd=HERE)
    return p.returncode, p.stdout.decode("utf-8", "replace"), p.stderr.decode("utf-8", "replace")


def step_verify(rep):
    py = python_exe()
    if py == sys.executable and not os.path.isfile(os.path.join(HERE, "python", "python.exe")):
        rep.line("WARN" if os.name == "nt" else "INFO", f"{os.path.join(HERE, 'python', 'python.exe')}가 없어 지금 python으로 확인합니다: {py}")
    code = ("import importlib, json\nout = {}\n"
            f"for name in {[m for _d, m in LIBS]!r}:\n"
            "    try:\n        importlib.import_module(name); out[name] = ''\n"
            "    except Exception as e:\n        out[name] = f'{type(e).__name__}: {e}'[:160]\n"
            "print(json.dumps(out))")
    try:
        rc, out, err = _run([py, "-c", code], 300)
        got = json.loads(out.strip().splitlines()[-1])
    except Exception as e:
        rep.line("FAIL", f"라이브러리 확인을 실행하지 못함 ({type(e).__name__}: {e})")
        got = None
    if got is not None:
        bad = [(d, got.get(m, "결과 없음")) for d, m in LIBS if got.get(m, "결과 없음")]
        for dist, msg in bad:
            if dist == "pywin32" and os.name != "nt":
                rep.line("WARN", "pywin32: Windows 전용이라 확인하지 않음")
            else:
                rep.line("FAIL", f"라이브러리 {dist} import 실패: {msg}")
        if len(bad) < len(LIBS):
            rep.line("OK", "라이브러리 import: " + ", ".join(d for d, m in LIBS if not got.get(m, "결과 없음")))
    for tool, rel in TOOL_HELP:
        script = os.path.join(HERE, rel)
        try:
            rc, out, err = _run([py, script, "-h"], 180)
        except Exception as e:
            rep.line("FAIL", f"{tool}: {rel} -h 실행 실패 ({type(e).__name__}: {e})")
            continue
        lines = [x for x in out.splitlines() if x.strip()]
        try:
            ok = rc == 0 and len(lines) == 1 and json.loads(lines[0]).get("status") == "ok"
        except ValueError:
            ok = False
        if ok:
            rep.line("OK", f"{tool}: {rel} -h -> stdout 1줄 JSON")
        else:
            tail = (err.strip().splitlines() or [""])[-1][:200]
            rep.line("FAIL", f"{tool}: {rel} -h -> exit {rc}, stdout {len(lines)}줄 (1줄 JSON이어야 함) {tail}")
    exe = os.path.join(HERE, "opencode.exe")
    rep.line("OK" if os.path.isfile(exe) else "FAIL", f"opencode.exe: {'있음' if os.path.isfile(exe) else '없음'}  {exe}")
    rg_places = [os.path.join(HERE, "python", "rg.exe"), os.path.join(HERE, "python", "Scripts", "rg.exe"), os.path.join(HERE, "rg.exe")]
    rg = shutil.which("rg") or next((x for x in rg_places if os.path.isfile(x)), None)
    if rg:
        rep.line("OK", f"rg(ripgrep): {rg}")
    else:
        rep.line("WARN", "rg(ripgrep)가 없습니다. OpenCode가 스킬을 못 불러올 수 있습니다: rg.exe를 C:\\CLI\\python에 두세요 (docs/opencode/README.md 1장)")
    if os.name == "nt" and os.path.normcase(HERE.rstrip("\\")) != os.path.normcase("C:\\CLI"):
        rep.line("WARN", f"이 폴더가 C:\\CLI가 아닙니다 ({HERE}). 스킬과 전역 지침의 명령은 C:\\CLI 기준입니다")


# ---------------------------------------------------------------------------
def main(argv=None):
    try:
        sys.stdout.reconfigure(errors="replace")
    except (AttributeError, ValueError):
        pass
    p = argparse.ArgumentParser(description="C:\\CLI 설치: OpenCode 스킬·전역 지침·opencode.jsonc 복사, 시스템 PATH 등록, 동작 확인")
    p.add_argument("--check", action="store_true", help="아무것도 바꾸지 않고 비교·확인만 (다르거나 실패가 있으면 exit 1)")
    p.add_argument("--config-dir", default=default_config_dir(), help=r"OpenCode 설정 폴더 (기본 %%USERPROFILE%%\.config\opencode)")
    p.add_argument("--no-path", action="store_true", help="시스템 PATH를 확인·등록하지 않음")
    p.add_argument("--test-skill", action="append", default=[], metavar="SKILL.md",
                   help="개발용: 보조 스킬도 함께 복사 (예: C:\\CLI-dev\\office-cli\\tools\\office-cli-test\\SKILL.md)")
    args = p.parse_args(argv)
    cfg = os.path.abspath(args.config_dir)
    backup = os.path.join(data_dir(), f"setup_backup_{dt.datetime.now():%Y%m%d_%H%M%S}")
    rep = Report()
    print(f"[INFO] 설치 폴더: {HERE}" + ("  (--check: 바꾸지 않음)" if args.check else ""), flush=True)
    print("[INFO] 1. OpenCode 스킬, 전역 지침, opencode.jsonc", flush=True)
    step_files(rep, cfg, backup, args.check, args.test_skill)
    print("[INFO] 2. 시스템 PATH", flush=True)
    if args.no_path:
        print("[INFO] --no-path: 건너뜀", flush=True)
    else:
        step_path(rep, backup, args.check)
    print("[INFO] 3. 동작 확인", flush=True)
    step_verify(rep)
    if os.path.isdir(backup):
        print(f"[INFO] 이전 파일 보관: {backup}", flush=True)
    bad = len(rep.fails) + len(rep.diffs)
    print(f"[STATUS] 결과: OK {rep.ok}, WARN {rep.warn}, FAIL {len(rep.fails)}"
          + (f", 다름 {len(rep.diffs)}" if args.check else ""), flush=True)
    for msg in rep.fails:
        print(f"[FAIL] {msg}", flush=True)
    if args.check and rep.diffs:
        print("[STATUS] 다른 것이 있습니다. 적용: C:\\CLI\\setup.bat (또는 python C:\\CLI\\setup.py)", flush=True)
    if not args.check and not rep.fails:
        print("[STATUS] 끝. 새 터미널(PowerShell)을 열고, OpenCode(TUI, Desktop)를 모두 닫았다가 다시 시작하세요.", flush=True)
    return 1 if bad else 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("[WARN] 사용자가 중단했습니다", flush=True)
        sys.exit(130)
