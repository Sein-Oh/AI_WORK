r"""C:\AIKIT 설치 (사람이 실행: setup.bat 두 번 클릭, 또는 python C:\AIKIT\setup.py)

배포는 C:\AIKIT 폴더 하나입니다. 이 스크립트가 그 폴더를 사용자 PC에 연결합니다. 관리자 권한은 시스템 PATH를 쓸 때만 필요합니다.
    python setup.py                 # 설치 (환경 변수, 옛 설정 정리, PATH 등록, 동작 확인)
    python setup.py --check         # 아무것도 바꾸지 않고 확인만 (다르거나 실패가 있으면 exit 1)
    python setup.py --no-path       # 시스템 PATH는 건드리지 않음

하는 일 (C:\AIKIT = 이 파일이 있는 폴더):
1. OpenCode 설정 폴더를 C:\AIKIT으로: 사용자 환경 변수 OPENCODE_CONFIG_DIR(HKCU\Environment)를 이 폴더로 설정합니다.
   OpenCode는 설정 파일 없이 폴더 모양만으로 AGENTS.md(전역 지침)와 skills\<이름>\SKILL.md(기본 스킬 11개와 사용자 스킬)를 읽고,
   opencode.json(셸·권한 기본 설정, 배포 파일)과 opencode.jsonc(사내 LLM 주소, 사용자가 넣음)를 합쳐 읽습니다 (OpenCode 2.2.5 확인).
   복사하지 않으므로 C:\AIKIT을 바꾸면 바로 반영됩니다. 원래 값은 data\setup_backup_<시각>\user_env.txt에 보관하고, 쓴 뒤 다시 읽어 확인합니다.
2. 옛 설정 정리 (원본은 data\setup_backup_<시각>\에 보관):
   - 옛 전역 설정 %USERPROFILE%\.config\opencode (2026-11-12 전 setup이 복사하던 곳, old_config\): 우리 스킬 복사본과 AGENTS.md는 지우고,
     사용자가 만든 스킬은 C:\AIKIT\skills로, opencode.jsonc는 C:\AIKIT에 없을 때 옮깁니다.
   - 이전 구성 (2026-11-12 낮, old_layout\): C:\AIKIT\opencode 폴더(AGENTS.md, opencode.json은 지우고 opencode.jsonc·사용자 스킬은 맨 위로),
     도구 폴더 안의 옛 skills 폴더(C:\AIKIT\ml-kit\skills, C:\AIKIT\office-cli\skills)는 지웁니다.
3. 시스템 PATH(HKLM)에 <C:\AIKIT>\python, <C:\AIKIT>\python\Scripts가 없으면 뒤에 추가합니다.
   원래 값은 data\setup_backup_<시각>\system_path.txt에 보관하고, 쓴 뒤 다시 읽어 확인하고, 환경 변경을 알립니다.
   OpenCode 백그라운드 서비스를 끔 (서비스는 처음 뜰 때의 환경 변수·PATH로 계속 돌아, 끄지 않으면 재부팅 전까지 옛 설정을 씀).
4. 동작 확인: python\python.exe로 라이브러리 import, 도구 -h가 stdout 1줄 JSON인지, opencode.exe, rg(ripgrep),
   opencode debug paths의 config가 C:\AIKIT인지. opencode.json + opencode.jsonc를 합친 설정의 "shell", "instructions",
   permission, "enabled_providers", "autoupdate"를 WARN/INFO로 알림.

개발용: --test-skill <SKILL.md>  (실기 테스트 보조 스킬을 C:\AIKIT\skills에 둠, 예: C:\AIKIT\dev\office-cli\tools\office-cli-test\SKILL.md)
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

HERE = os.path.dirname(os.path.abspath(__file__))                  # C:\AIKIT (배포 폴더)
OFFICE_SKILLS = ("office-excel", "office-doc", "office-mail", "office-spec", "office-data", "skill-maker")    # skills\<이름>\SKILL.md (office-cli)
ML_SKILLS = ("ml-train", "ml-predict", "ml-report", "ml-map", "ml-dataset")                     # skills\<이름>\SKILL.md (ml-kit)
# 옛 전역 설정 폴더에서 지울 우리 스킬 이름 (지금 스킬 + 옛 통합 스킬 폴더 + 실기 보조 스킬)
OLD_SKILLS = ML_SKILLS + OFFICE_SKILLS + ("office-cli", "ml-kit", "office-cli-test")
ENV_KEY = r"SYSTEM\CurrentControlSet\Control\Session Manager\Environment"     # 시스템 환경 변수 (HKLM)
USER_ENV_KEY = "Environment"                                                   # 사용자 환경 변수 (HKCU)
CONFIG_ENV = "OPENCODE_CONFIG_DIR"                                             # OpenCode 설정 폴더를 바꾸는 환경 변수
# opencode.json(우리 설정)에만 두는 키: opencode.jsonc(사용자가 넣는 사내 LLM 설정)에도 있으면 어느 쪽이 쓰일지 헷갈림
OUR_KEYS = ("shell", "permission", "autoupdate")
# 도구가 쓰는 라이브러리 (AGENTS.md 사용자 환경 표): (배포 이름, import 이름)
LIBS = (("numpy", "numpy"), ("pandas", "pandas"), ("scikit-learn", "sklearn"), ("scipy", "scipy"), ("joblib", "joblib"),
        ("openpyxl", "openpyxl"), ("python-docx", "docx"), ("python-pptx", "pptx"), ("matplotlib", "matplotlib"),
        ("pymupdf", "pymupdf"), ("olefile", "olefile"), ("pywin32", "win32com.client"))
TOOL_HELP = (("ml-kit", os.path.join("ml-kit", "ml_predict.py")), ("office-cli", os.path.join("office-cli", "office.py")))
LONG_PATH = 2047                                                     # 이보다 길면 일부 프로그램이 뒤쪽 PATH를 못 봄


def config_dir():
    """OpenCode 설정 폴더 = C:\\AIKIT (OPENCODE_CONFIG_DIR가 가리키는 곳): AGENTS.md, skills\\, opencode.json, opencode.jsonc가 맨 위."""
    return HERE


def old_config_dir():
    """옛 전역 설정 폴더 (OpenCode 기본 위치, 2026-11-12 전 setup이 스킬·지침을 복사하던 곳)."""
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
# 1. OpenCode 설정 폴더 (C:\AIKIT)
# ---------------------------------------------------------------------------
def sources():
    """(이름, 파일) 목록: OpenCode가 C:\\AIKIT에서 바로 읽는 파일 (복사하지 않음)."""
    items = [(f"{name} 스킬", os.path.join(HERE, "skills", name, "SKILL.md")) for name in ML_SKILLS + OFFICE_SKILLS]
    items.append(("전역 지침", os.path.join(config_dir(), "AGENTS.md")))
    items.append(("opencode.json(우리 설정)", os.path.join(config_dir(), "opencode.json")))
    return items


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


def step_sources(rep):
    """OpenCode가 읽을 파일(기본 스킬 11개, 전역 지침, opencode.json)이 C:\\AIKIT에 있는지."""
    missing = [(n, p) for n, p in sources() if not os.path.isfile(p)]
    for name, path in missing:
        rep.line("FAIL", f"{name}: 파일이 없습니다: {path} (저장소의 AIKIT 폴더를 통째로 다시 복사하세요)")
    if not missing:
        vers = sorted({version_of(_read(p)) for n, p in sources() if n.endswith("스킬")} - {""})
        mine = sorted(set(os.listdir(os.path.join(HERE, "skills"))) - set(ML_SKILLS + OFFICE_SKILLS))
        rep.line("OK", f"기본 스킬 {len(ML_SKILLS) + len(OFFICE_SKILLS)}개 ({', '.join(vers)}), 전역 지침, opencode.json: 있음  {HERE}"
                       + (f" (사용자 스킬 {len(mine)}개: {', '.join(mine[:10])})" if mine else ""))


def _winreg():
    try:
        import winreg
        return winreg
    except ImportError:                                 # Windows가 아님
        return None


def read_user_env(reg):
    """HKCU\\Environment의 OPENCODE_CONFIG_DIR (없으면 None)."""
    with reg.OpenKey(reg.HKEY_CURRENT_USER, USER_ENV_KEY, 0, reg.KEY_READ) as key:
        try:
            return reg.QueryValueEx(key, CONFIG_ENV)[0]
        except FileNotFoundError:
            return None


def step_env(rep, backup, check=False, reg=None):
    """사용자 환경 변수 OPENCODE_CONFIG_DIR = C:\\AIKIT (관리자 권한 필요 없음). 이 프로세스에도 넣어 동작 확인에 씀."""
    want = config_dir()
    os.environ[CONFIG_ENV] = want
    reg = reg or _winreg()
    if reg is None:
        rep.line("WARN", f"Windows가 아니라 사용자 환경 변수 {CONFIG_ENV}는 확인하지 않았습니다")
        return
    try:
        cur = read_user_env(reg)
    except OSError as e:
        rep.line("FAIL", f"사용자 환경 변수를 읽지 못함 ({type(e).__name__}: {e})")
        return
    if cur is not None and _norm(cur) == _norm(want):
        rep.line("OK", f"사용자 환경 변수 {CONFIG_ENV} = {cur}")
        return
    if check:
        rep.diff(f"사용자 환경 변수 {CONFIG_ENV}: {cur or '없음'} (설치하면 {want})")
        return
    try:
        os.makedirs(backup, exist_ok=True)
        keep = os.path.join(backup, "user_env.txt")
        with open(keep, "w", encoding="utf-8") as f:
            f.write(f"# HKCU\\{USER_ENV_KEY} {CONFIG_ENV}, 설치 전 값 (없으면 빈 줄)\n{cur or ''}\n")
    except OSError as e:
        rep.line("FAIL", f"사용자 환경 변수 원래 값을 보관하지 못해 바꾸지 않았습니다 ({type(e).__name__}: {e})")
        return
    try:
        with reg.OpenKey(reg.HKEY_CURRENT_USER, USER_ENV_KEY, 0, reg.KEY_READ | reg.KEY_SET_VALUE) as key:
            reg.SetValueEx(key, CONFIG_ENV, 0, reg.REG_SZ, want)
    except OSError as e:
        rep.line("FAIL", f"사용자 환경 변수 {CONFIG_ENV}를 쓰지 못함 ({type(e).__name__}: {e})")
        return
    if _norm(read_user_env(reg) or "") != _norm(want):
        rep.line("FAIL", f"사용자 환경 변수 {CONFIG_ENV}를 쓴 뒤 다시 읽은 값이 다릅니다")
        return
    rep.line("OK", f"사용자 환경 변수 {CONFIG_ENV} = {want} (원래 값: {cur or '없음'}, 보관: {keep})")
    try:
        sent = broadcast_env_change()
    except Exception as e:                              # 알림 실패는 값에 영향 없음
        sent = False
        print(f"[INFO] 환경 변경 알림 오류: {type(e).__name__}: {e}", flush=True)
    if not sent:
        rep.line("WARN", "환경 변경 알림을 보내지 못했습니다. OpenCode에 새 설정이 안 보이면 로그아웃 후 다시 로그인하세요")


def _is_link(path):
    """심볼릭 링크나 Windows 정션(mklink /J)인지. 링크면 링크만 지우고 가리키는 폴더는 건드리지 않습니다."""
    isjunction = getattr(os.path, "isjunction", None)          # Python 3.12+
    return os.path.islink(path) or bool(isjunction and isjunction(path))


def _remove_link(path):
    try:
        os.unlink(path)
    except (IsADirectoryError, PermissionError):
        os.rmdir(path)                                  # Windows 디렉터리 링크·정션은 rmdir로 링크만 지움 (대상 내용은 그대로)


def old_global_items(cfg):
    """옛 전역 설정 폴더에서 정리할 것: [(종류, 경로)]. 종류: 우리 스킬 / 전역 지침 / 사용자 스킬 / opencode.jsonc."""
    out = []
    skills = os.path.join(cfg, "skills")
    if os.path.isdir(skills):
        for name in sorted(os.listdir(skills)):
            path = os.path.join(skills, name)
            if name in OLD_SKILLS:
                out.append(("우리 스킬", path))
            elif os.path.isfile(os.path.join(path, "SKILL.md")) and not _is_link(path):
                out.append(("사용자 스킬", path))
    for kind, name in (("전역 지침", "AGENTS.md"), ("opencode.jsonc", "opencode.jsonc")):
        if os.path.isfile(os.path.join(cfg, name)):
            out.append((kind, os.path.join(cfg, name)))
    return out


def old_layout_items():
    """이전 구성(2026-11-12 낮)에서 정리할 것: [(종류, 경로)]. C:\\AIKIT\\opencode 폴더(설정 폴더였음)와 도구 폴더 안의 옛 skills 폴더
    (덮어쓰기 배포로 남음. OpenCode는 C:\\AIKIT\\skills만 읽어 헷갈리기만 함)."""
    out = []
    for tool in ("ml-kit", "office-cli"):
        path = os.path.join(HERE, tool, "skills")
        if os.path.isdir(path):
            out.append(("우리 스킬", path))
    oc = os.path.join(HERE, "opencode")
    if not os.path.isdir(oc) or _is_link(oc):
        return out
    skills = os.path.join(oc, "skills")
    if os.path.isdir(skills):
        for name in sorted(os.listdir(skills)):
            path = os.path.join(skills, name)
            if name in OLD_SKILLS:
                out.append(("우리 스킬", path))
            elif os.path.isfile(os.path.join(path, "SKILL.md")) and not _is_link(path):
                out.append(("사용자 스킬", path))
    for kind, name in (("전역 지침", "AGENTS.md"), ("우리 설정", "opencode.json"), ("OpenCode 실행 파일", "service.json"),
                       ("opencode.jsonc", "opencode.jsonc")):
        if os.path.isfile(os.path.join(oc, name)):
            out.append((kind, os.path.join(oc, name)))
    return out


def _keep(path, root, backup, sub="old_config"):
    """옛 것을 data\\setup_backup_<시각>\\<sub>\\ 아래 root 기준 같은 상대 경로로 보관 (링크는 보관하지 않음)."""
    if _is_link(path):
        return
    keep = os.path.join(backup, sub, os.path.relpath(path, root))
    os.makedirs(os.path.dirname(keep), exist_ok=True)
    if os.path.isdir(path):
        shutil.copytree(path, keep, dirs_exist_ok=True)
    else:
        shutil.copy2(path, keep)


def _remove(path):
    if _is_link(path):
        _remove_link(path)
    elif os.path.isdir(path):
        shutil.rmtree(path)
    else:
        os.remove(path)


def _tidy(rep, kind, path, root, backup, label="옛 전역 설정", sub="old_config"):
    """옛 설정 하나를 정리: 보관한 뒤 지우거나(우리 것), C:\\AIKIT으로 옮긴 뒤 지움(사용자 스킬은 skills\\, opencode.jsonc는 맨 위)."""
    dst, note = None, "보관 후 지움"
    if kind == "사용자 스킬":
        dst = os.path.join(config_dir(), "skills", os.path.basename(path))
    elif kind == "opencode.jsonc":
        dst = os.path.join(config_dir(), "opencode.jsonc")
    if dst and os.path.exists(dst):
        same = kind == "opencode.jsonc" and _read(dst) == _read(path)
        note = "보관 후 지움 (C:\\AIKIT 쪽 것과 같음)" if same else f"보관 후 지움 (이미 있는 {dst}를 씀. 옛 것은 보관 폴더에)"
        dst = None
    err = ""
    try:
        _keep(path, root, backup, sub)
        if dst:
            if os.path.isdir(path):
                shutil.copytree(path, dst)
            else:
                os.makedirs(os.path.dirname(dst), exist_ok=True)
                shutil.copy2(path, dst)
            note = f"{dst}로 옮김"
        _remove(path)
    except PermissionError as e:                        # OpenCode가 파일을 잡고 있으면: 닫고 다시 실행
        err = f" ({type(e).__name__}: {e}. OpenCode를 모두 닫고 다시 실행)"
    except OSError as e:
        err = f" ({type(e).__name__}: {e}. 직접 지우거나 옮긴 뒤 다시 실행)"
    if os.path.lexists(path):
        rep.line("FAIL", f"{label} {kind}: {path} -> 정리하지 못함{err}")
    else:
        rep.line("OK", f"{label} {kind}: {path} -> {note}")


def step_old_global(rep, backup, check=False, cfg=None):
    """옛 전역 설정 폴더(%USERPROFILE%\\.config\\opencode)에 남은 우리 스킬·지침 정리 (OPENCODE_CONFIG_DIR로 이제 읽히지 않음)."""
    cfg = os.path.abspath(cfg or old_config_dir())
    if _norm(cfg) == _norm(config_dir()):
        return
    items = old_global_items(cfg)
    if not items:
        rep.line("OK", f"옛 전역 설정 폴더에 정리할 것 없음  {cfg}")
        return
    for kind, path in items:
        if check:
            rep.diff(f"옛 전역 설정 {kind}: {path} (설치하면 정리)")
        else:
            _tidy(rep, kind, path, cfg, backup)


def step_old_layout(rep, backup, check=False):
    """이전 구성(C:\\AIKIT\\opencode가 설정 폴더, 스킬은 도구 폴더 안)에서 옮김. 다 비면 C:\\AIKIT\\opencode 폴더도 지움."""
    items = old_layout_items()
    if not items:
        return
    for kind, path in items:
        if check:
            rep.diff(f"이전 구성 {kind}: {path} (설치하면 정리)")
        else:
            _tidy(rep, kind, path, HERE, backup, label="이전 구성", sub="old_layout")
    oc = os.path.join(HERE, "opencode")
    if not check and os.path.isdir(oc):
        for d in (os.path.join(oc, "skills"), oc):
            try:
                if os.path.isdir(d) and not os.listdir(d):
                    os.rmdir(d)
            except OSError:
                pass


def step_test_skills(rep, backup, test_skills):
    """개발용 보조 스킬을 C:\\AIKIT\\skills\\<이름>\\SKILL.md로 복사 (다르면 보관 후 덮어씀)."""
    for src in test_skills:
        src = os.path.abspath(src)
        if os.path.isdir(src):
            src = os.path.join(src, "SKILL.md")
        name = os.path.basename(os.path.dirname(src))
        dst = os.path.join(config_dir(), "skills", name, "SKILL.md")
        a, b = _read(src), _read(dst)
        if a is None:
            rep.line("FAIL", f"{name} 스킬(개발용): 원본이 없습니다: {src}")
            continue
        if a == b:
            rep.line("OK", f"{name} 스킬(개발용): 같음  {dst}")
            continue
        if b is not None:
            keep = os.path.join(backup, "test_skills", name, "SKILL.md")
            os.makedirs(os.path.dirname(keep), exist_ok=True)
            shutil.copy2(dst, keep)
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        shutil.copyfile(src, dst)
        rep.line("OK" if _read(dst) == a else "FAIL", f"{name} 스킬(개발용): {dst} -> 복사함")


# ---------------------------------------------------------------------------
# 3. 시스템 PATH (HKLM, winreg. setx는 1024자에서 잘라 쓰므로 쓰지 않음)
# ---------------------------------------------------------------------------
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
# 4. 동작 확인
# ---------------------------------------------------------------------------
def _run(cmd, timeout):
    env = dict(os.environ, PYTHONIOENCODING="utf-8")
    p = subprocess.run(cmd, capture_output=True, timeout=timeout, env=env, cwd=HERE)
    return p.returncode, p.stdout.decode("utf-8", "replace"), p.stderr.decode("utf-8", "replace")


def step_service(rep, run=None):
    """OpenCode 백그라운드 서비스를 끔: TUI와 Web UI(web.bat)는 이미 떠 있는 서비스를 그대로 다시 쓰므로, 설치 전에 뜬 서비스는
    옛 PATH(python을 못 찾음)와 옛 스킬·지침으로 계속 돈다. OpenCode 창을 닫거나 web.bat을 다시 실행해도 바뀌지 않고 재부팅해야
    풀렸다 (회사 PC 팀원 배포 2026-11-11). 다음에 OpenCode를 실행하면 새 설정으로 다시 뜬다. 서비스가 없거나 opencode.exe가 없으면 그냥 지나감."""
    exe = os.path.join(HERE, "opencode.exe")
    if not os.path.isfile(exe):
        return
    run = run or _run
    try:
        rc, out, err = run([exe, "service", "stop"], 60)
    except Exception as e:                                    # 서비스를 못 꺼도 설치는 실패가 아님
        rep.line("WARN", f"OpenCode 백그라운드 서비스를 끄지 못했습니다 ({type(e).__name__}). 재부팅하거나 'C:\\AIKIT\\web.bat stop'을 실행하세요")
        return
    if rc == 0:
        rep.line("OK", "OpenCode 백그라운드 서비스를 껐습니다 (다음에 OpenCode를 실행하면 새 PATH·스킬·지침으로 다시 뜸)")
    else:
        rep.line("INFO", "OpenCode 백그라운드 서비스: 떠 있지 않거나 끄지 못함 (" + (err or out).strip()[:120] + ")")


def step_opencode_paths(rep, run=None):
    """OpenCode가 OPENCODE_CONFIG_DIR를 따르는지 실제로 확인: 'opencode debug paths'의 config 줄이 C:\\AIKIT인지
    (OpenCode 2.2.5 집 PC 2026-11-12 확인). opencode.exe가 없거나 debug paths가 없는 버전이면 INFO."""
    exe = os.path.join(HERE, "opencode.exe")
    if not os.path.isfile(exe):
        return
    run = run or _run
    try:
        rc, out, err = run([exe, "debug", "paths"], 60)
    except Exception as e:
        rep.line("INFO", f"opencode debug paths를 실행하지 못해 설정 폴더를 확인하지 않았습니다 ({type(e).__name__})")
        return
    m = re.search(r"^\s*config\s+(.+?)\s*$", out, re.M)
    if rc != 0 or not m:
        rep.line("INFO", "opencode debug paths에 config 줄이 없어 설정 폴더를 확인하지 않았습니다 (OpenCode 버전 확인)")
    elif _norm(m.group(1)) == _norm(config_dir()):
        rep.line("OK", f"OpenCode 설정 폴더: {m.group(1)} (opencode debug paths)")
    else:
        rep.line("FAIL", f"OpenCode 설정 폴더가 {m.group(1)}입니다 ({config_dir()}이어야 함). 새 PowerShell 창에서 다시 확인하고, "
                         f"그래도 다르면 로그아웃 후 다시 로그인하세요")


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
        rep.line("WARN", "rg(ripgrep)가 없습니다. OpenCode가 스킬을 못 불러올 수 있습니다: rg.exe를 C:\\AIKIT\\python에 두세요 (docs/opencode/README.md 1장)")
    if os.name == "nt" and os.path.normcase(HERE.rstrip("\\")) != os.path.normcase("C:\\AIKIT"):
        rep.line("WARN", f"이 폴더가 C:\\AIKIT가 아닙니다 ({HERE}). 스킬과 전역 지침의 명령은 C:\\AIKIT 기준입니다")


def parse_jsonc(text):
    """opencode.jsonc -> dict: 문자열 밖의 // 줄 주석, /* */ 주석, 닫는 괄호 앞 쉼표를 빼고 JSON으로 읽음. 못 읽으면 ValueError."""
    out, i, n, in_str = [], 0, len(text), False
    while i < n:
        c = text[i]
        if in_str:
            out.append(c)
            if c == "\\" and i + 1 < n:
                out.append(text[i + 1])
                i += 1
            elif c == '"':
                in_str = False
        elif c == '"':
            in_str = True
            out.append(c)
        elif text.startswith("//", i):
            while i < n and text[i] not in "\r\n":
                i += 1
            continue
        elif text.startswith("/*", i):
            end = text.find("*/", i + 2)
            i = n if end < 0 else end + 2
            continue
        else:
            out.append(c)
        i += 1
    text, out, in_str = "".join(out), [], False
    for i, c in enumerate(text):                      # 문자열 밖의, 닫는 괄호 앞 쉼표만 뺌 (문자열 안의 ", ]"는 그대로)
        if in_str:
            in_str = not (c == '"' and text[i - 1] != "\\")
        elif c == '"':
            in_str = True
        elif c == "," and text[i + 1:].lstrip()[:1] in ("}", "]"):
            continue
        out.append(c)
    return json.loads("".join(out))


def load_config(config_dir):
    """설정 폴더의 opencode.json(우리 설정)과 opencode.jsonc(사내 LLM 설정)를 OpenCode처럼 합친 dict.
    (설정, 두 파일에 함께 있는 우리 키, 못 읽은 파일 목록, opencode.jsonc가 있는지). 객체 값(permission 등)은 키별로 합침."""
    merged, overlap, bad = {}, [], []
    seen = {}
    for name in ("opencode.json", "opencode.jsonc"):
        path = os.path.join(config_dir, name)
        if not os.path.isfile(path):
            continue
        try:
            cfg = parse_jsonc(_read(path).decode("utf-8-sig"))
        except (ValueError, UnicodeDecodeError) as e:
            bad.append(f"{name} ({type(e).__name__})")
            continue
        if not isinstance(cfg, dict):
            continue
        for key, val in cfg.items():
            if key in OUR_KEYS and key in seen and seen[key] != name:
                overlap.append(key)
            seen.setdefault(key, name)
            if isinstance(val, dict) and isinstance(merged.get(key), dict):
                merged[key] = {**merged[key], **val}
            else:
                merged[key] = val
    return merged, overlap, bad, os.path.isfile(os.path.join(config_dir, "opencode.jsonc"))


def step_opencode_config(rep, config_dir):
    """합친 설정에서 도구 결과를 흐리는 설정을 알림 (집 PC opencode 시험 2026-11-07에서 모두 결과를 흐림). WARN만, 고치지 않음.
    셸이 PowerShell이 아님(스킬·전역 지침의 명령과 따옴표는 PowerShell 기준), instructions의 다른 지침(옛 지침이 섞이면 지금 없는 명령을 씀),
    작업 폴더 밖 파일 권한(external_directory가 allow가 아니면 opencode run이 작업 폴더 밖 쓰기를 묻지 않고 거절해 모델이 멈춤),
    우리 키(skills, shell, permission, autoupdate)를 opencode.jsonc에도 적음(opencode.json과 어느 쪽이 쓰일지 헷갈림)."""
    cfg, overlap, bad, has_jsonc = load_config(config_dir)
    for name in bad:
        rep.line("WARN", f"{name}을 읽지 못해 설정을 확인하지 않았습니다  {config_dir}")
    if not has_jsonc:
        rep.line("WARN", f"opencode.jsonc가 없습니다. 사내 LLM 주소가 든 opencode.jsonc를 {config_dir}\\에 넣고 다시 실행하세요")
    if cfg.get("skills"):
        rep.line("WARN", "설정에 \"skills\"(스킬 경로)가 있습니다. 스킬은 C:\\AIKIT\\skills에 두면 설정 없이 읽히니 opencode.jsonc에서 지우세요 "
                         "(예전 경로를 가리키면 옛 스킬이 섞임)")
    if overlap:
        rep.line("WARN", f"opencode.json과 opencode.jsonc에 같은 키가 있습니다 ({', '.join(sorted(set(overlap)))}). 이 키는 opencode.json(배포 파일)에만 두고, "
                         "opencode.jsonc에는 provider, model, enabled_providers만 두세요")
    shell = cfg.get("shell")
    if shell is not None and not re.search(r"(?i)powershell|pwsh", json.dumps(shell)):
        rep.line("WARN", f"설정의 \"shell\"이 PowerShell이 아닙니다 ({json.dumps(shell, ensure_ascii=False)[:80]}). "
                         "스킬과 전역 지침의 명령은 PowerShell 기준입니다: \"shell\": \"powershell\"로 바꾸세요")
    extra = cfg.get("instructions")
    if extra:
        names = ", ".join(str(x) for x in (extra if isinstance(extra, list) else [extra]))[:200]
        rep.line("WARN", f"설정의 \"instructions\"에 다른 지침이 있습니다 ({names}). 전역 지침(AGENTS.md)과 다른 옛 지침이 "
                         "섞이면 지금 없는 명령을 쓸 수 있습니다: 필요 없으면 지우세요")
    perm = cfg.get("permission")
    ext = perm if isinstance(perm, str) else (perm or {}).get("external_directory") if isinstance(perm, dict) else None
    if isinstance(ext, dict) and ext and all(v == "allow" for v in ext.values()):    # 경로 패턴별 허용 {"*": "allow"}
        ext = "allow"
    if ext is None:
        rep.line("INFO", "설정에 permission.external_directory가 없습니다. opencode 버전에 따라 작업 폴더 밖 쓰기를 묻지 않고 "
                         "거절할 수 있습니다 (집 PC opencode 1.19.5 run). 그러면 \"permission\": {\"external_directory\": \"allow\"}")
    elif ext != "allow":
        rep.line("WARN", f"설정의 permission.external_directory가 {json.dumps(ext, ensure_ascii=False)[:60]}입니다. "
                         "작업 폴더 밖(예: 바탕 화면의 파일) 쓰기가 막혀 도구 결과 저장이 실패할 수 있습니다: \"allow\"를 검토하세요")


def step_opencode_providers(rep, config_dir):
    """모델 목록 제한: 폐쇄망에서 쓸 수 없는 기본 모델(opencode Zen, models.dev 목록의 인터넷 모델)이
    `opencode models`와 모델 선택에 나오지 않게 opencode.jsonc의 "enabled_providers"(허용 목록, 나머지는 모두 숨김)를 권함.
    opencode 2.x 소스 기준: enabled_providers = 모두 막고 적은 것만 허용, disabled_providers = 적은 것만 막음 (막기가 이김).
    WARN/INFO만, 고치지 않음 (opencode.jsonc는 사용자가 넣는 파일)."""
    cfg, _overlap, _bad, has_jsonc = load_config(config_dir)
    if not has_jsonc:
        return                                   # 없는 것은 step_opencode_config가 이미 알림
    own = [str(k) for k in (cfg.get("provider") or {})] if isinstance(cfg.get("provider"), dict) else []
    model = cfg.get("model") if isinstance(cfg.get("model"), str) else ""
    used = model.split("/", 1)[0] if "/" in model else ""
    ids = own or ([used] if used else [])
    enabled = cfg.get("enabled_providers")
    if enabled is None:
        line = json.dumps({"enabled_providers": ids or ["<사내 LLM provider 이름>"]}, ensure_ascii=False)[1:-1]
        rep.line("WARN", "opencode.jsonc에 \"enabled_providers\"가 없어 폐쇄망에서 쓸 수 없는 기본 모델(opencode 등)도 "
                         f"`opencode models`와 모델 선택에 나옵니다. 사내 LLM만 보이게 하려면 opencode.jsonc에 {line} 를 넣으세요 "
                         "(적은 provider만 남고 나머지는 숨김)")
    elif isinstance(enabled, list):
        names = [str(x) for x in enabled]
        if not names:
            rep.line("WARN", "opencode.jsonc의 \"enabled_providers\"가 비어 있어 모든 모델이 숨겨집니다. 사내 LLM provider 이름을 넣으세요")
        if own and any(n not in own for n in names):
            rep.line("WARN", f"opencode.jsonc의 \"enabled_providers\" {names} 중 \"provider\"에 없는 이름이 있습니다 (provider: {own}). "
                             "이름이 틀리면 사내 LLM도 숨겨집니다")
        if used and used not in names:
            rep.line("WARN", f"opencode.jsonc의 \"model\"({model})의 provider '{used}'가 \"enabled_providers\"에 없어 그 모델을 쓸 수 없습니다")
    perm = cfg.get("permission")
    web = [k for k in ("webfetch", "websearch") if not (perm == "deny" or (isinstance(perm, dict) and perm.get(k) == "deny"))]
    if web:
        rep.line("INFO", "설정에서 웹 도구(" + ", ".join(web) + ")를 막지 않았습니다. 폐쇄망에서는 모델이 인터넷 검색을 시도하다 시간만 쓰니 "
                         "opencode.json의 \"permission\"에 " + ", ".join(f'"{k}": "deny"' for k in web) + " 를 넣는 것을 권합니다")
    if cfg.get("autoupdate") is not False:
        rep.line("INFO", "설정에 \"autoupdate\": false가 없습니다. 폐쇄망에서는 업데이트를 받을 수 없으니 넣어 두면 확인을 시도하지 않습니다")


# ---------------------------------------------------------------------------
def main(argv=None):
    try:
        sys.stdout.reconfigure(errors="replace")
    except (AttributeError, ValueError):
        pass
    p = argparse.ArgumentParser(description="C:\\AIKIT 설치: OpenCode 설정 폴더를 C:\\AIKIT으로, 옛 설정 정리, 시스템 PATH 등록, 동작 확인")
    p.add_argument("--check", action="store_true", help="아무것도 바꾸지 않고 확인만 (다르거나 실패가 있으면 exit 1)")
    p.add_argument("--no-path", action="store_true", help="시스템 PATH를 확인·등록하지 않음")
    p.add_argument("--test-skill", action="append", default=[], metavar="SKILL.md",
                   help="개발용: 보조 스킬을 C:\\AIKIT\\skills에 둠 (예: C:\\AIKIT\\dev\\office-cli\\tools\\office-cli-test\\SKILL.md)")
    args = p.parse_args(argv)
    backup = os.path.join(data_dir(), f"setup_backup_{dt.datetime.now():%Y%m%d_%H%M%S}")
    rep = Report()
    print(f"[INFO] 설치 폴더: {HERE}" + ("  (--check: 바꾸지 않음)" if args.check else ""), flush=True)
    print(f"[INFO] 1. OpenCode 설정 폴더: {config_dir()}", flush=True)
    step_sources(rep)
    step_env(rep, backup, args.check)
    if args.test_skill and not args.check:
        step_test_skills(rep, backup, args.test_skill)
    print(f"[INFO] 2. 옛 설정 정리: {old_config_dir()}, 이전 구성", flush=True)
    step_old_global(rep, backup, args.check)
    step_old_layout(rep, backup, args.check)
    if not args.check:      # 스킬로 안 되는 일의 파이썬 스크립트를 두는 곳 (전역 지침 8장)
        os.makedirs(os.path.join(data_dir(), "scripts"), exist_ok=True)
        # Web UI의 작업 폴더 (web.bat이 서비스를 여기서 띄움). Web UI는 세션을 프로젝트 폴더 안에 만들어, 폴더가 없는 새 PC에서는
        # 새 세션이 만들어지지 않았음 (팀원 배포 2026-11-11). 처음 한 번 Web UI에서 이 폴더를 프로젝트로 연다
        os.makedirs(os.path.join(HERE, "work"), exist_ok=True)
    print("[INFO] 3. 시스템 PATH", flush=True)
    if args.no_path:
        print("[INFO] --no-path: 건너뜀", flush=True)
    else:
        step_path(rep, backup, args.check)
    if not args.check:
        step_service(rep)
    print("[INFO] 4. 동작 확인", flush=True)
    step_verify(rep)
    step_opencode_paths(rep)
    step_opencode_config(rep, config_dir())
    step_opencode_providers(rep, config_dir())
    if os.path.isdir(backup):
        print(f"[INFO] 이전 파일 보관: {backup}", flush=True)
    bad = len(rep.fails) + len(rep.diffs)
    print(f"[STATUS] 결과: OK {rep.ok}, WARN {rep.warn}, FAIL {len(rep.fails)}"
          + (f", 다름 {len(rep.diffs)}" if args.check else ""), flush=True)
    for msg in rep.fails:
        print(f"[FAIL] {msg}", flush=True)
    if args.check and rep.diffs:
        print("[STATUS] 다른 것이 있습니다. 적용: C:\\AIKIT\\setup.bat (또는 python C:\\AIKIT\\setup.py)", flush=True)
    if not args.check and not rep.fails:
        print(f"[INFO] Web UI(web.bat)를 처음 쓰면 화면에서 프로젝트 열기로 {os.path.join(HERE, 'work')} 를 고른 뒤 새 세션을 만드세요", flush=True)
        print("[STATUS] 끝. 새 터미널(PowerShell)을 열고, OpenCode(TUI, Desktop, Web UI)를 모두 닫았다가 다시 시작하세요. "
              "새 스킬·지침이 안 보이거나 python을 못 찾으면 'C:\\AIKIT\\web.bat stop' 뒤 다시 실행하세요.", flush=True)
    return 1 if bad else 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("[WARN] 사용자가 중단했습니다", flush=True)
        sys.exit(130)
