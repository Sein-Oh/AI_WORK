r"""C:\AIKIT 설치 (사람이 실행: setup.bat 두 번 클릭, 또는 python C:\AIKIT\setup.py)

배포는 C:\AIKIT 폴더 하나입니다. 우리 도구·지침·스킬은 C:\AIKIT을 OpenCode 프로젝트로 열 때만 씁니다 (사용자 결정 2026-11-13).
관리자 권한은 시스템 PATH를 쓸 때만 필요합니다.
    python setup.py                 # 설치 (전역 설정에 우리 설정 합치기, PATH 등록, 동작 확인)
    python setup.py --check         # 아무것도 바꾸지 않고 확인만 (다르거나 실패가 있으면 exit 1)
    python setup.py --no-path       # 시스템 PATH는 건드리지 않음

하는 일 (C:\AIKIT = 이 파일이 있는 폴더):
1. C:\AIKIT 프로젝트 파일 확인: AGENTS.md(프로젝트 지침), .opencode\skills\<이름>\SKILL.md(기본 스킬 13개와 사용자 스킬),
   opencode.json(셸·권한·폐쇄망 잠금). OpenCode는 C:\AIKIT을 프로젝트로 연 대화(web.bat, C:\AIKIT에서 실행한 opencode)에서만 이것을 읽습니다.
2. OpenCode 전역 설정(%USERPROFILE%\.config\opencode): 우리 opencode.json의 셸·권한·폐쇄망 잠금을 전역 opencode.json에 합칩니다
   (모든 프로젝트에 적용, 사용자 결정 2026-11-13). 원래 파일은 data\setup_backup_<시각>에 보관하고, 쓴 뒤 다시 읽어 확인합니다.
   사내 LLM 설정(opencode.jsonc)은 사용자가 이 전역 폴더에 둡니다 (C:\AIKIT\opencode.jsonc에 있으면 이 폴더로 옮기고 원래 파일은 보관). 예전 설치가 남긴 사용자 환경 변수 OPENCODE_CONFIG_DIR=C:\AIKIT는
   지웁니다 (남아 있으면 C:\AIKIT이 전역 설정 폴더가 되어 모든 대화에 우리 지침·스킬이 들어감).
3. 시스템 PATH(HKLM)에 <C:\AIKIT>\python, <C:\AIKIT>\python\Scripts가 없으면 뒤에 추가합니다.
   원래 값은 data\setup_backup_<시각>\system_path.txt에 보관하고, 쓴 뒤 다시 읽어 확인하고, 환경 변경을 알립니다.
   OpenCode 백그라운드 서비스를 끔 (서비스는 처음 뜰 때의 환경 변수·PATH로 계속 돌아, 끄지 않으면 재부팅 전까지 옛 설정을 씀).
4. 동작 확인: python\python.exe로 라이브러리 import, 도구 -h가 stdout 1줄 JSON인지, opencode.exe, rg(ripgrep, 없어도 됨),
   opencode debug paths의 config가 전역 설정 폴더인지. 전역 opencode.json + opencode.jsonc를 합친 설정의 "shell", "instructions",
   permission, "enabled_providers"(적었을 때만 이름 확인), 폐쇄망 잠금(OFFLINE_DENY, OFFLINE_SET)을 WARN/INFO로 알림.

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
OFFICE_SKILLS = ("excel-kit", "excel-format-kit", "word-kit", "ppt-kit", "outlook-kit", "search-kit", "stats-kit")    # .opencode\skills\<이름>\SKILL.md (office-cli)
ML_SKILLS = ("ml-train", "ml-predict", "ml-report", "ml-map", "ml-dataset")                     # .opencode\skills\<이름>\SKILL.md (ml-kit)
PANDAS_SKILLS = ("pandas-kit",)                                                                 # .opencode\skills\<이름>\SKILL.md (pandas-kit)
DEFAULT_SKILLS = ML_SKILLS + OFFICE_SKILLS + PANDAS_SKILLS                                       # 기본 스킬 13개
ENV_KEY = r"SYSTEM\CurrentControlSet\Control\Session Manager\Environment"     # 시스템 환경 변수 (HKLM)
USER_ENV_KEY = "Environment"                                                   # 사용자 환경 변수 (HKCU)
CONFIG_ENV = "OPENCODE_CONFIG_DIR"                                             # OpenCode 전역 설정 폴더를 바꾸는 환경 변수 (예전 설치가 C:\AIKIT으로 넣음, 지금은 지움)
# opencode.json(우리 설정)에만 두는 키: opencode.jsonc(사용자가 넣는 사내 LLM 설정)에도 있으면 어느 쪽이 쓰일지 헷갈림
OUR_KEYS = ("shell", "permission", "autoupdate", "share", "lsp", "formatter")
# 폐쇄망에서 되지 않는 기능 잠금 (배포 opencode.json에 있음, 집 PC에서 인터넷을 막고 확인 2026-11-12).
# 막는 도구: 웹 가져오기·검색·브라우저, 파일 찾기 grep·glob(rg.exe가 필요한데 배포하지 않음, 막으면 모델이 셸로 찾음)
OFFLINE_DENY = ("webfetch", "websearch", "browser", "grep", "glob")
# 끄는 기능: 업데이트 확인, 세션 공유(opencode.ai로 보냄), 언어 서버·서식 도구 자동 내려받기
OFFLINE_SET = (("autoupdate", False), ("share", "disabled"), ("lsp", False), ("formatter", False))
# 도구가 쓰는 라이브러리 (AGENTS.md 사용자 환경 표): (배포 이름, import 이름)
LIBS = (("numpy", "numpy"), ("pandas", "pandas"), ("scikit-learn", "sklearn"), ("scipy", "scipy"), ("joblib", "joblib"),
        ("openpyxl", "openpyxl"), ("python-docx", "docx"), ("python-pptx", "pptx"), ("matplotlib", "matplotlib"),
        ("pymupdf", "pymupdf"), ("olefile", "olefile"), ("pywin32", "win32com.client"))
TOOL_HELP = (("ml-kit", os.path.join("ml-kit", "ml_predict.py")), ("office-cli", os.path.join("office-cli", "office.py")),
             ("pandas-kit", os.path.join("pandas-kit", "pandas_kit.py")))
LONG_PATH = 2047                                                     # 이보다 길면 일부 프로그램이 뒤쪽 PATH를 못 봄


def config_dir():
    """C:\\AIKIT = OpenCode 프로젝트 폴더: AGENTS.md(프로젝트 지침), .opencode\\skills\\, opencode.json이 맨 위."""
    return HERE


def skills_dir():
    """프로젝트 스킬 폴더 C:\\AIKIT\\.opencode\\skills (OpenCode가 프로젝트에서 스킬을 찾는 곳, 사용자가 만든 스킬도 여기)."""
    return os.path.join(HERE, ".opencode", "skills")


def global_config_dir():
    """OpenCode 전역 설정 폴더 (모든 프로젝트): %XDG_CONFIG_HOME% 또는 %USERPROFILE%\\.config 아래 opencode.
    사내 LLM 설정(opencode.jsonc)과 setup이 합치는 셸·권한·폐쇄망 잠금(opencode.json)이 여기."""
    base = os.environ.get("XDG_CONFIG_HOME") or os.path.join(os.path.expanduser("~"), ".config")
    return os.path.join(base, "opencode")


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
# 1. C:\AIKIT 프로젝트 파일
# ---------------------------------------------------------------------------
def sources():
    """(이름, 파일) 목록: C:\\AIKIT을 프로젝트로 열면 OpenCode가 그 자리에서 읽는 파일 (복사하지 않음)."""
    items = [(f"{name} 스킬", os.path.join(skills_dir(), name, "SKILL.md")) for name in DEFAULT_SKILLS]
    items.append(("프로젝트 지침", os.path.join(config_dir(), "AGENTS.md")))
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
    """C:\\AIKIT 프로젝트 파일(기본 스킬 13개, 프로젝트 지침, opencode.json)이 있는지."""
    missing = [(n, p) for n, p in sources() if not os.path.isfile(p)]
    for name, path in missing:
        rep.line("FAIL", f"{name}: 파일이 없습니다: {path} (저장소의 AIKIT 폴더를 통째로 다시 복사하세요. .opencode 폴더도 함께)")
    if not missing:
        vers = sorted({version_of(_read(p)) for n, p in sources() if n.endswith("스킬")} - {""})
        mine = sorted(set(os.listdir(skills_dir())) - set(DEFAULT_SKILLS))
        rep.line("OK", f"기본 스킬 {len(DEFAULT_SKILLS)}개 ({', '.join(vers)}), 프로젝트 지침, opencode.json: 있음  {HERE}"
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
    """사용자 환경 변수 OPENCODE_CONFIG_DIR가 C:\\AIKIT이면 지움 (예전 설치가 넣은 값: 남아 있으면 C:\\AIKIT이 전역 설정 폴더가 되어
    다른 프로젝트 대화에도 우리 지침·스킬이 들어감). 다른 폴더를 가리키면 사용자 설정이라 그대로 두고 알림. 이 프로세스에서도 뺌."""
    if os.environ.get(CONFIG_ENV) and _norm(os.environ[CONFIG_ENV]) == _norm(config_dir()):
        del os.environ[CONFIG_ENV]
    reg = reg or _winreg()
    if reg is None:
        rep.line("WARN", f"Windows가 아니라 사용자 환경 변수 {CONFIG_ENV}는 확인하지 않았습니다")
        return
    try:
        cur = read_user_env(reg)
    except OSError as e:
        rep.line("FAIL", f"사용자 환경 변수를 읽지 못함 ({type(e).__name__}: {e})")
        return
    if not cur:
        rep.line("OK", f"사용자 환경 변수 {CONFIG_ENV}: 없음 (전역 설정은 {global_config_dir()})")
        return
    if _norm(cur) != _norm(config_dir()):
        rep.line("WARN", f"사용자 환경 변수 {CONFIG_ENV}가 {cur}입니다. 사용자가 정한 값이라 그대로 두지만, OpenCode가 그 폴더를 전역 설정으로 "
                         f"읽어 이 setup이 합친 {global_config_dir()}의 설정은 쓰이지 않습니다")
        return
    if check:
        rep.diff(f"사용자 환경 변수 {CONFIG_ENV} = {cur} (예전 설치 값, 설치하면 지움)")
        return
    try:
        os.makedirs(backup, exist_ok=True)
        keep = os.path.join(backup, "user_env.txt")
        with open(keep, "w", encoding="utf-8") as f:
            f.write(f"# HKCU\\{USER_ENV_KEY} {CONFIG_ENV}, 설치 전 값\n{cur}\n")
        with reg.OpenKey(reg.HKEY_CURRENT_USER, USER_ENV_KEY, 0, reg.KEY_READ | reg.KEY_SET_VALUE) as key:
            reg.DeleteValue(key, CONFIG_ENV)
    except OSError as e:
        rep.line("FAIL", f"사용자 환경 변수 {CONFIG_ENV}를 지우지 못함 ({type(e).__name__}: {e})")
        return
    if read_user_env(reg):
        rep.line("FAIL", f"사용자 환경 변수 {CONFIG_ENV}를 지운 뒤에도 남아 있습니다")
        return
    rep.line("OK", f"사용자 환경 변수 {CONFIG_ENV}={cur}를 지웠습니다 (보관: {keep}). 이제 C:\\AIKIT은 프로젝트로 열 때만 쓰입니다")
    try:
        sent = broadcast_env_change()
    except Exception as e:                              # 알림 실패는 값에 영향 없음
        sent = False
        print(f"[INFO] 환경 변경 알림 오류: {type(e).__name__}: {e}", flush=True)
    if not sent:
        rep.line("WARN", "환경 변경 알림을 보내지 못했습니다. OpenCode에 새 설정이 안 보이면 로그아웃 후 다시 로그인하세요")


def step_global_config(rep, backup, check=False):
    """우리 opencode.json의 셸·권한·폐쇄망 잠금(OUR_KEYS)을 전역 opencode.json에 합침 (모든 프로젝트에 적용, 사용자 결정 2026-11-13).
    전역 파일의 다른 키는 그대로, permission은 키별로 합침. 못 읽는 파일은 바꾸지 않음. 쓰기 전 보관, 쓴 뒤 다시 읽어 확인."""
    gdir = global_config_dir()
    target = os.path.join(gdir, "opencode.json")
    try:
        ours = parse_jsonc(_read(os.path.join(config_dir(), "opencode.json")).decode("utf-8-sig"))
    except (AttributeError, ValueError, UnicodeDecodeError) as e:
        rep.line("FAIL", f"C:\\AIKIT\\opencode.json(우리 설정)을 읽지 못함 ({type(e).__name__})")
        return
    cur = {}
    if os.path.isfile(target):
        try:
            cur = parse_jsonc(_read(target).decode("utf-8-sig"))
        except (ValueError, UnicodeDecodeError) as e:
            rep.line("FAIL", f"전역 설정 {target}을 읽지 못해 바꾸지 않았습니다 ({type(e).__name__}). 파일을 고친 뒤 다시 실행하세요")
            return
        if not isinstance(cur, dict):
            rep.line("FAIL", f"전역 설정 {target}이 JSON 객체가 아니라 바꾸지 않았습니다")
            return
    merged = dict(cur)
    for key in OUR_KEYS:
        if key not in ours:
            continue
        if isinstance(ours[key], dict) and isinstance(cur.get(key), dict):
            merged[key] = {**cur[key], **ours[key]}
        else:
            merged[key] = ours[key]
    if merged == cur:
        rep.line("OK", f"전역 설정에 우리 셸·권한·폐쇄망 잠금이 있음  {target}")
        return
    changed = sorted(k for k in OUR_KEYS if merged.get(k) != cur.get(k))
    if check:
        rep.diff(f"전역 설정 {target}에 없거나 다른 우리 설정: {', '.join(changed)} (설치하면 합침)")
        return
    try:
        if os.path.isfile(target):
            os.makedirs(backup, exist_ok=True)
            shutil.copy2(target, os.path.join(backup, "global_opencode.json"))
        os.makedirs(gdir, exist_ok=True)
        with open(target, "w", encoding="utf-8") as f:
            f.write(json.dumps(merged, ensure_ascii=False, indent=2) + "\n")
        again = parse_jsonc(_read(target).decode("utf-8-sig"))
    except (OSError, ValueError, UnicodeDecodeError) as e:
        rep.line("FAIL", f"전역 설정 {target}을 쓰지 못함 ({type(e).__name__}: {e})")
        return
    if again != merged:
        rep.line("FAIL", f"전역 설정 {target}을 쓴 뒤 다시 읽은 값이 다릅니다")
        return
    rep.line("OK", f"전역 설정에 우리 설정을 합쳤습니다 ({', '.join(changed)})  {target}"
                   + (f" (원래 파일 보관: {os.path.join(backup, 'global_opencode.json')})" if cur else ""))


def step_move_llm_config(rep, backup, check=False):
    """C:\\AIKIT\\opencode.jsonc(예전 자리, 이제 C:\\AIKIT 프로젝트에서만 읽힘)를 전역 설정 폴더로 옮김 (사용자 결정 2026-11-13).
    전역에 없으면 복사 -> 같은지 다시 읽어 확인 -> 원래 파일은 보관 폴더로 옮김. 전역에 같은 내용이 있으면 원래 파일만 보관 폴더로.
    전역에 다른 내용이 있으면 어느 쪽이 맞는지 모르니 둘 다 그대로 두고 WARN."""
    src = os.path.join(HERE, "opencode.jsonc")
    if not os.path.isfile(src):
        return
    dst = os.path.join(global_config_dir(), "opencode.jsonc")
    data, have = _read(src), _read(dst) if os.path.isfile(dst) else None
    if data is None:
        rep.line("FAIL", f"{src}을 읽지 못해 옮기지 않았습니다")
        return
    if have is not None and have.replace(b"\r\n", b"\n") != data.replace(b"\r\n", b"\n"):
        rep.line("WARN", f"{src}와 {dst}의 내용이 달라 옮기지 않았습니다. OpenCode는 다른 프로젝트에서 {dst}를 씁니다: "
                         f"맞는 쪽을 {dst}에 두고 {src}는 지우세요")
        return
    if check:
        rep.diff(f"{src} (사내 LLM 설정, 설치하면 {dst}로 옮김)")
        return
    try:
        if have is None:
            os.makedirs(global_config_dir(), exist_ok=True)
            shutil.copy2(src, dst)
            if _read(dst) != data:
                rep.line("FAIL", f"{dst}에 복사한 내용이 원래 파일과 다릅니다. {src}는 그대로 두었습니다")
                return
        os.makedirs(backup, exist_ok=True)
        shutil.move(src, os.path.join(backup, "opencode.jsonc"))
    except OSError as e:
        rep.line("FAIL", f"{src}를 {dst}로 옮기지 못함 ({type(e).__name__}: {e})")
        return
    rep.line("OK", f"사내 LLM 설정을 옮겼습니다: {src} -> {dst}" + (" (같은 내용이 이미 있었음)" if have is not None else "")
                   + f" (원래 파일 보관: {os.path.join(backup, 'opencode.jsonc')})")


def step_global_leftovers(rep):
    """전역 설정 폴더에 남은 예전 복사본(우리 기본 스킬, AGENTS.md)을 알림. 전역 폴더는 모든 프로젝트에 적용되어 C:\\AIKIT 밖 대화에도
    우리 스킬·지침이 들어감. 사용자가 만든 것일 수 있어 지우지 않고 WARN (삭제는 사람이 확인)."""
    gdir = global_config_dir()
    ours = [n for n in DEFAULT_SKILLS
            if os.path.isfile(os.path.join(gdir, "skills", n, "SKILL.md")) or os.path.isfile(os.path.join(gdir, "skill", n, "SKILL.md"))]
    if ours:
        rep.line("WARN", f"전역 설정 폴더에 예전 우리 스킬이 남아 있습니다 ({', '.join(ours)}). 모든 프로젝트에 보이니 "
                         f"{os.path.join(gdir, 'skills')}에서 이 폴더들을 지우세요 (C:\\AIKIT\\.opencode\\skills의 것을 씀)")
    if os.path.isfile(os.path.join(gdir, "AGENTS.md")):
        rep.line("WARN", f"{os.path.join(gdir, 'AGENTS.md')}가 있습니다. 모든 프로젝트에 적용되는 지침이라, 예전에 복사한 우리 지침이면 지우세요 "
                         "(우리 지침은 C:\\AIKIT\\AGENTS.md)")


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
    """OpenCode 전역 설정 폴더 확인: 'opencode debug paths'의 config 줄이 전역 설정 폴더인지 (C:\\AIKIT이면 예전 환경 변수가 이 창에 남은 것).
    opencode.exe가 없거나 debug paths가 없는 버전이면 INFO."""
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
        rep.line("FAIL", f"OpenCode 전역 설정 폴더가 아직 {m.group(1)}입니다 (예전 환경 변수 {CONFIG_ENV}). 새 PowerShell 창에서 다시 확인하고, "
                         "그래도 같으면 로그아웃 후 다시 로그인하세요")
    elif _norm(m.group(1)) == _norm(global_config_dir()):
        rep.line("OK", f"OpenCode 전역 설정 폴더: {m.group(1)} (opencode debug paths)")
    else:
        rep.line("INFO", f"OpenCode 전역 설정 폴더가 {m.group(1)}입니다 (이 setup은 {global_config_dir()}에 합쳤음)")


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
        # rg.exe는 배포하지 않음 (2026-11-12 사용자 결정). OpenCode 2.2.5는 rg 없이 스킬을 불러오고, 파일 찾기 도구(grep, glob)는
        # opencode.json에서 꺼 둠. rg가 필요한 것은 Web UI의 @ 파일 자동 완성뿐 (집 PC 확인).
        rep.line("INFO", "rg(ripgrep) 없음: 정상입니다 (배포에 넣지 않음). 스킬과 도구는 그대로 되고, Web UI의 @ 파일 자동 완성만 안 됩니다")
    if os.name == "nt" and os.path.normcase(HERE.rstrip("\\")) != os.path.normcase("C:\\AIKIT"):
        rep.line("WARN", f"이 폴더가 C:\\AIKIT가 아닙니다 ({HERE}). 스킬과 프로젝트 지침의 명령은 C:\\AIKIT 기준입니다")


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
    """전역 설정 폴더의 opencode.json(setup이 합친 우리 설정)과 opencode.jsonc(사내 LLM 설정)를 OpenCode처럼 합친 dict.
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
    우리 키(shell, permission, autoupdate)를 opencode.jsonc에도 적음(opencode.json과 어느 쪽이 쓰일지 헷갈림)."""
    cfg, overlap, bad, has_jsonc = load_config(config_dir)
    for name in bad:
        rep.line("WARN", f"{name}을 읽지 못해 설정을 확인하지 않았습니다  {config_dir}")
    if not has_jsonc:
        rep.line("WARN", f"opencode.jsonc가 없습니다. 사내 LLM 주소가 든 opencode.jsonc를 전역 설정 폴더 {config_dir}\\에 넣고 다시 실행하세요")
    if overlap:
        rep.line("WARN", f"전역 opencode.json과 opencode.jsonc에 같은 키가 있습니다 ({', '.join(sorted(set(overlap)))}). 이 키는 opencode.json(setup이 합침)에만 두고, "
                         "opencode.jsonc에는 provider, model과 모델 목록 설정(enabled_providers 또는 disabled_providers)만 두세요")
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
    """모델 목록 설정 확인: "enabled_providers"를 적었을 때만 이름이 provider·model과 맞는지 봄 (틀리면 사내 LLM도 숨겨짐).
    없으면 알리지 않음: 기본 모델을 숨기는 방법(enabled_providers, disabled_providers)과 숨길지는 PC마다 사용자가 정함
    (사용자 결정 2026-11-13: 회사 PC는 disabled_providers로 opencode만 숨기고, 집 PC는 보이게 둠).
    opencode 2.x 소스 기준: enabled_providers = 모두 막고 적은 것만 허용, disabled_providers = 적은 것만 막음 (막기가 이김).
    WARN/INFO만, 고치지 않음 (opencode.jsonc는 사용자가 넣는 파일)."""
    cfg, _overlap, _bad, has_jsonc = load_config(config_dir)
    if not has_jsonc:
        return                                   # 없는 것은 step_opencode_config가 이미 알림
    own = [str(k) for k in (cfg.get("provider") or {})] if isinstance(cfg.get("provider"), dict) else []
    model = cfg.get("model") if isinstance(cfg.get("model"), str) else ""
    used = model.split("/", 1)[0] if "/" in model else ""
    enabled = cfg.get("enabled_providers")
    if isinstance(enabled, list):
        names = [str(x) for x in enabled]
        if not names:
            rep.line("WARN", "opencode.jsonc의 \"enabled_providers\"가 비어 있어 모든 모델이 숨겨집니다. 사내 LLM provider 이름을 넣으세요")
        if own and any(n not in own for n in names):
            rep.line("WARN", f"opencode.jsonc의 \"enabled_providers\" {names} 중 \"provider\"에 없는 이름이 있습니다 (provider: {own}). "
                             "이름이 틀리면 사내 LLM도 숨겨집니다")
        if used and used not in names:
            rep.line("WARN", f"opencode.jsonc의 \"model\"({model})의 provider '{used}'가 \"enabled_providers\"에 없어 그 모델을 쓸 수 없습니다")
    perm = cfg.get("permission")
    deny = [k for k in OFFLINE_DENY if not (perm == "deny" or (isinstance(perm, dict) and perm.get(k) == "deny"))]
    unset = [f'"{k}": {json.dumps(v)}' for k, v in OFFLINE_SET if cfg.get(k) != v]
    if deny or unset:
        want = unset + (['"permission": {' + ", ".join(f'"{k}": "deny"' for k in deny) + "}"] if deny else [])
        rep.line("INFO", "폐쇄망에서 되지 않는 기능을 잠그지 않았습니다 (모델이 인터넷·rg를 쓰려다 시간만 씀). 배포 opencode.json처럼 "
                         + ", ".join(want) + " 를 넣는 것을 권합니다")


# ---------------------------------------------------------------------------
def main(argv=None):
    try:
        sys.stdout.reconfigure(errors="replace")
    except (AttributeError, ValueError):
        pass
    p = argparse.ArgumentParser(description="C:\\AIKIT 설치: 프로젝트 파일 확인, 전역 설정에 우리 설정 합치기, 시스템 PATH 등록, 동작 확인")
    p.add_argument("--check", action="store_true", help="아무것도 바꾸지 않고 확인만 (다르거나 실패가 있으면 exit 1)")
    p.add_argument("--no-path", action="store_true", help="시스템 PATH를 확인·등록하지 않음")
    args = p.parse_args(argv)
    backup = os.path.join(data_dir(), f"setup_backup_{dt.datetime.now():%Y%m%d_%H%M%S}")
    rep = Report()
    print(f"[INFO] 설치 폴더: {HERE}" + ("  (--check: 바꾸지 않음)" if args.check else ""), flush=True)
    print(f"[INFO] 1. C:\\AIKIT 프로젝트 파일: {config_dir()}", flush=True)
    step_sources(rep)
    if not args.check:      # 스킬로 안 되는 일의 파이썬 스크립트를 두는 곳 (전역 지침 8장)
        os.makedirs(os.path.join(data_dir(), "scripts"), exist_ok=True)
    print(f"[INFO] 2. OpenCode 전역 설정: {global_config_dir()}", flush=True)
    step_env(rep, backup, args.check)
    step_global_config(rep, backup, args.check)
    step_move_llm_config(rep, backup, args.check)
    step_global_leftovers(rep)
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
    step_opencode_config(rep, global_config_dir())
    step_opencode_providers(rep, global_config_dir())
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
        print(f"[INFO] 우리 도구는 {HERE}를 프로젝트로 열 때만 보입니다: Web UI는 web.bat (처음이면 화면에서 프로젝트 열기로 {HERE}), "
              f"TUI는 {HERE}에서 opencode를 실행하세요", flush=True)
        print("[STATUS] 끝. 새 터미널(PowerShell)을 열고, OpenCode(TUI, Desktop, Web UI)를 모두 닫았다가 다시 시작하세요. "
              "새 스킬·지침이 안 보이거나 python을 못 찾으면 'C:\\AIKIT\\web.bat stop' 뒤 다시 실행하세요.", flush=True)
    return 1 if bad else 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("[WARN] 사용자가 중단했습니다", flush=True)
        sys.exit(130)
