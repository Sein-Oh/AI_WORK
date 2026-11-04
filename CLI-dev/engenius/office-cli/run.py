"""En-Genius office-cli 스킬 실행 입구. run_skill_script(name="office-cli", path="run.py", args=[...])로 불린다.

도구는 이 PC의 C:\\CLI\\office-cli에 있다 (OpenCode와 같은 폴더, 같은 데이터). 이 스킬에는 지시문과 이 파일만 있다.
args[0] = 도구 명령 한 줄 (SKILL.md의 명령, 결과의 next_commands[i].command, [추천]의 전체 명령을 그대로).
  - 맨 앞의 `python`과 스크립트 경로(C:\\CLI\\...\\)는 있어도 되고 없어도 된다.
  - `@input1.csv`, `@input2.md`처럼 쓴 자리는 args[1], args[2]의 내용을 UTF-8 파일로 만들어 그 경로로 바꾼다
    (표, 여러 줄 글, 큰따옴표가 든 값을 넘기는 방법. En-Genius 대화에는 PC에 파일을 쓰는 도구가 없다).
args가 이미 나뉘어 있으면(`["excel.py", "read"]`, `@input` 자리 없음) 그대로 쓴다.

파이썬: 지금 파이썬(트레이의 「스킬 실행 파이썬」)에 도구 라이브러리가 있으면 그것으로, 없으면 C:\\WPy64-*의 WinPython으로 실행한다.
도구 결과(stdout 1줄 JSON, 종료 코드)는 도구의 것 그대로다. 이 입구의 오류도 1줄 JSON이다 (exit 2, STOP_RETRY).
위치 바꾸기(시험용): 환경변수 ENGENIUS_CLI_DIR(C:\\CLI 대신), ENGENIUS_PYTHON(WinPython python.exe).
빌드: CLI-dev/engenius/build.py가 SKILL_VERSION, SKILL_HASH를 채워 office-cli.zip에 넣는다.
지시문 확인: zip의 지시문은 C:\\CLI\\office-cli\\SKILL.md(OpenCode용)로 만든다. 그 본문(버전 줄 제외)이 zip을 만들 때와 다르면
명령 안내가 바뀐 것이므로 실행하지 않고 zip을 다시 가져오라고 안내한다. 도구만 고친 업데이트는 C:\\CLI만 바꾸면 된다.
"""
import glob
import hashlib
import importlib.util
import json
import os
import re
import runpy
import subprocess
import sys

TOOL = "office-cli"
SKILL_VERSION = ""                                   # 빌드가 채운다 (zip을 만들 때의 도구 버전, 표시용)
SKILL_HASH = ""                                      # 빌드가 채운다 (그때 SKILL.md 본문의 skill_hash)
SCRIPTS = ("excel.py", "word.py", "ppt.py", "outlook.py", "reader.py", "office.py")
VERSION_FILE = "_runtime.py"
NEEDS = ("openpyxl", "docx", "pptx")                 # 지금 파이썬으로 실행해도 되는지 보는 라이브러리
OTHER_TOOL = "ml-kit"
OTHER_SCRIPTS = ("ml_train.py", "ml_predict.py", "ml_report.py", "ml_map.py", "ml_export.py")
EXAMPLE = 'args=["word.py write --file @input1.md", "## 제목\\n- 항목"]'

HERE = os.path.dirname(os.path.abspath(__file__))
CLI_DIR = os.environ.get("ENGENIUS_CLI_DIR") or r"C:\CLI"
TOOL_DIR = os.path.join(CLI_DIR, TOOL)
INPUT = re.compile(r"@input(\d+)((?:\.[A-Za-z0-9]{1,8})?)")
MAX_INPUTS = 20


def emit(obj, code):
    sys.stdout.buffer.write((json.dumps(obj, ensure_ascii=False) + "\n").encode(sys.stdout.encoding or "utf-8", "replace"))
    sys.stdout.buffer.flush()
    sys.exit(code)


def stop(message, error_type="input"):
    emit({"status": "error", "error_type": error_type, "message": message,
          "agent_directive": "STOP_RETRY: 명령 형식이나 설치 문제입니다. message의 안내대로 고칠 수 있으면 1번만 다시 실행하고, "
                             "아니면 message를 사용자에게 그대로 알리세요."}, 2)


def split_command(line):
    """Windows 명령 한 줄 -> 인자 목록. 큰따옴표로 묶은 부분은 한 인자, 그 안의 \\"는 큰따옴표 글자
    (도구의 quote_arg가 쓰는 규칙). 따옴표 안에서 \\" 뒤가 공백이나 끝이면 역슬래시 + 닫는 따옴표로 본다
    ("C:\\폴더\\" 같은 폴더 경로). 작은따옴표로 시작해 작은따옴표로 끝나는 인자는 작은따옴표를 벗긴다
    (PowerShell 습관: '=$A$1')."""
    args, cur, quoted, has, i = [], [], False, False, 0
    while i < len(line):
        ch = line[i]
        if ch == "\\" and i + 1 < len(line) and line[i + 1] == '"':
            nxt = line[i + 2] if i + 2 < len(line) else ""
            if quoted and (nxt == "" or nxt.isspace()):
                cur.append("\\")
                quoted = False
                i += 2
                continue
            cur.append('"')
            has = True
            i += 2
            continue
        if ch == '"':
            quoted, has = not quoted, True
        elif ch.isspace() and not quoted:
            if has or cur:
                args.append("".join(cur))
            cur, has = [], False
        else:
            cur.append(ch)
        i += 1
    if quoted:
        return None
    if has or cur:
        args.append("".join(cur))
    return [a[1:-1] if len(a) >= 2 and a[0] == "'" and a[-1] == "'" else a for a in args]


def tool_version():
    """C:\\CLI 도구의 버전 (도구를 불러오지 않고 VERSION 줄만 읽는다). 없으면 None."""
    try:
        with open(os.path.join(TOOL_DIR, VERSION_FILE), encoding="utf-8") as f:
            m = re.search(r'^VERSION\s*=\s*"([^"]+)"', f.read(), re.M)
        return m.group(1) if m else None
    except OSError:
        return None


def skill_hash(text):
    """SKILL.md 본문 지문: 줄바꿈·BOM을 맞추고 버전 줄(version: "...")을 뺀 글의 sha256 앞 16자.
    빌드(build.py)도 이 함수를 불러 쓴다 (같은 계산)."""
    text = text.replace("\r\n", "\n").lstrip("\ufeff")
    text = re.sub(r'(?m)^[ \t]*version:[ \t]*"[^"]*"[ \t]*\n?', "", text)
    return hashlib.sha256(text.strip().encode("utf-8")).hexdigest()[:16]


def installed_skill_hash():
    """C:\\CLI 도구 폴더 SKILL.md의 지문. 없으면 None (확인하지 않음)."""
    try:
        with open(os.path.join(TOOL_DIR, "SKILL.md"), encoding="utf-8") as f:
            return skill_hash(f.read())
    except OSError:
        return None


def check_install():
    if not all(os.path.isfile(os.path.join(TOOL_DIR, s)) for s in SCRIPTS):
        stop(f"이 PC에 {TOOL} 도구가 없습니다 ({TOOL_DIR}). 배포 폴더의 CLI를 C:\\CLI로 복사한 뒤 다시 실행하세요. "
             "사용자에게 그대로 알리세요.", "setup")
    found = installed_skill_hash()
    if SKILL_HASH and found and found != SKILL_HASH:
        stop(f"C:\\CLI의 {TOOL}({tool_version() or '버전 모름'})는 명령 안내가 바뀌었는데 En-Genius의 스킬 zip은 예전 것"
             f"({SKILL_VERSION or '버전 모름'})입니다. 같은 배포본의 {TOOL}.zip을 En-Genius 스킬 탭에서 다시 가져온 뒤(덮어쓰기) "
             "실행하세요. 사용자에게 그대로 알리세요.", "setup")


def winpython():
    """도구 라이브러리가 있는 파이썬: ENGENIUS_PYTHON > C:\\WPy64-*의 WinPython (최신 이름 먼저)."""
    given = os.environ.get("ENGENIUS_PYTHON")
    if given:
        return given if os.path.isfile(given) else None
    found = sorted(glob.glob(r"C:\WPy64-*\python\python.exe") + glob.glob(r"C:\WPy64-*\python-*\python.exe"),
                   reverse=True)
    return found[0] if found else None


def has_libs():
    try:
        return all(importlib.util.find_spec(m) is not None for m in NEEDS)
    except (ImportError, ValueError):
        return False


def usage():
    print(__doc__, file=sys.stderr)
    current = installed_skill_hash()
    emit({"status": "ok", "tool": TOOL, "skill_version": SKILL_VERSION or None, "tool_dir": TOOL_DIR,
          "tool_version": tool_version(), "skill_current": None if not (SKILL_HASH and current) else current == SKILL_HASH,
          "scripts": list(SCRIPTS),
          "python": sys.executable if has_libs() else winpython(),
          "usage": f'run_skill_script(name="{TOOL}", path="run.py", args=["<명령 한 줄>"])', "example": EXAMPLE,
          "agent_directive": "VERIFIED: 사용법입니다. SKILL.md의 명령을 args 첫 칸에 그대로 넣어 실행하세요."}, 0)


def build_argv(raw):
    used = sorted({int(m.group(1)) for m in INPUT.finditer(raw[0])})
    if len(raw) > 1 and not used:
        return [str(a) for a in raw]                       # 이미 나뉜 인자
    argv = split_command(raw[0])
    if argv is None:
        stop("명령의 큰따옴표가 닫히지 않았습니다. 따옴표 짝을 맞춰 다시 실행하세요.")
    contents = raw[1:]
    if used and (used[0] < 1 or used[-1] > len(contents) or len(used) > MAX_INPUTS):
        stop(f"명령에 @input{used[-1]} 자리가 있지만 args에 그 내용이 없습니다 (내용 {len(contents)}칸). "
             "@input1은 args 두 번째 칸, @input2는 세 번째 칸의 내용입니다.")
    if len(contents) > len(used):
        stop(f"args에 쓰이지 않은 칸이 {len(contents) - len(used)}개 있습니다. 명령은 args 첫 칸에 한 줄로 넣고, "
             "그 뒤 칸은 명령의 @input1, @input2 자리 내용만 넣으세요.")
    if not used:
        return argv
    folder = os.path.join(HERE, "_inputs")
    os.makedirs(folder, exist_ok=True)
    paths = {}
    for n in used:
        ext = next((m.group(2) for m in INPUT.finditer(raw[0]) if int(m.group(1)) == n and m.group(2)), ".txt")
        path = os.path.join(folder, f"input{n}{ext}")
        with open(path, "w", encoding="utf-8", newline="") as f:
            f.write(str(contents[n - 1]))
        paths[n] = path
    return [INPUT.sub(lambda m: paths[int(m.group(1))], a) for a in argv]


def main(raw):
    if not raw or raw[0].strip() in ("", "-h", "--help"):
        usage()
    argv = build_argv(raw)
    if argv and os.path.splitext(os.path.basename(argv[0]).lower())[0] in ("python", "python3", "py"):
        argv = argv[1:]
    if not argv:
        stop(f'실행할 스크립트가 없습니다. 예: "{SCRIPTS[0]} -h"')
    script = os.path.basename(argv[0].replace("\\", "/"))
    if script not in SCRIPTS:
        if script in OTHER_SCRIPTS:
            stop(f"'{script}'는 {OTHER_TOOL} 스킬의 명령입니다. 같은 명령을 "
                 f'run_skill_script(name="{OTHER_TOOL}", path="run.py", args=["<같은 명령>"])로 실행하세요.')
        stop(f"'{argv[0]}'는 {TOOL} 스킬의 명령이 아닙니다. 쓸 수 있는 스크립트: {', '.join(SCRIPTS)}")
    check_install()
    path = os.path.join(TOOL_DIR, script)
    if has_libs():
        sys.argv = [path] + argv[1:]
        sys.path[0] = TOOL_DIR
        runpy.run_path(path, run_name="__main__")
        return
    py = winpython()
    if not py:
        stop("도구 라이브러리(pandas, openpyxl 등)가 있는 파이썬을 찾지 못했습니다. WinPython을 C:\\WPy64-...에 두거나, "
             "En-Genius® MCP 설정의 「스킬 실행 파이썬」을 직접 지정으로 WinPython의 python.exe로 바꾸세요. 사용자에게 그대로 알리세요.",
             "setup")
    env = dict(os.environ, PYTHONUTF8="1", PYTHONIOENCODING="utf-8")
    for k in ("PYTHONHOME", "PYTHONPATH", "VIRTUAL_ENV"):
        env.pop(k, None)
    try:
        code = subprocess.call([py, path] + argv[1:], cwd=TOOL_DIR, env=env)
    except OSError as e:
        stop(f"WinPython을 실행하지 못했습니다 ({py}): {e}. 사용자에게 그대로 알리세요.", "setup")
    sys.exit(code)


if __name__ == "__main__":
    main(sys.argv[1:])
