"""En-Genius office-cli 스킬 실행 입구. run_skill_script(name="office-cli", path="run.py", args=[...])로 불린다.

args[0] = 도구 명령 한 줄 (SKILL.md의 명령, 결과의 next_commands[i].command, [추천]의 전체 명령을 그대로).
  - 맨 앞의 `python`과 스크립트 경로(C:\\CLI\\...\\, 임시 폴더 경로)는 있어도 되고 없어도 된다.
  - `@input1.csv`, `@input2.md`처럼 쓴 자리는 args[1], args[2]의 내용을 UTF-8 파일로 만들어 그 경로로 바꾼다
    (표, 여러 줄 글, 큰따옴표가 든 값을 넘기는 방법. En-Genius 대화에는 PC에 파일을 쓰는 도구가 없다).
args가 이미 나뉘어 있으면(`["ml_train.py", "-f", "a.xlsx"]`, `@input` 자리 없음) 그대로 쓴다.

도구는 이 폴더의 스크립트를 같은 프로세스에서 실행하고, 결과(stdout 1줄 JSON, 종료 코드)는 도구의 것 그대로다.
이 입구의 오류도 1줄 JSON이다 (exit 2, STOP_RETRY). 빌드: CLI-dev/engenius/build.py가 office-cli.zip에 넣는다.
"""
import json
import os
import re
import runpy
import sys

TOOL = "office-cli"
HERE = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = ("excel.py", "word.py", "ppt.py", "outlook.py", "reader.py", "office.py")
OTHER_TOOL = "ml-kit"
OTHER_SCRIPTS = ("ml_train.py", "ml_predict.py", "ml_report.py", "ml_map.py", "ml_export.py")
INPUT = re.compile(r"@input(\d+)((?:\.[A-Za-z0-9]{1,8})?)")
MAX_INPUTS = 20


def emit(obj, code):
    sys.stdout.buffer.write((json.dumps(obj, ensure_ascii=False) + "\n").encode(sys.stdout.encoding or "utf-8", "replace"))
    sys.stdout.buffer.flush()
    sys.exit(code)


def stop(message):
    emit({"status": "error", "error_type": "input", "message": message,
          "agent_directive": "STOP_RETRY: 명령 형식 문제입니다. message의 안내대로 고칠 수 있으면 1번만 다시 실행하고, "
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


def usage():
    print(__doc__, file=sys.stderr)
    emit({"status": "ok", "tool": TOOL, "scripts": SCRIPTS,
          "usage": f'run_skill_script(name="{TOOL}", path="run.py", args=["<명령 한 줄>"])',
          "example": 'args=["word.py write --file @input1.md", "## 제목\\n- 항목"]',
          "agent_directive": "VERIFIED: 사용법입니다. SKILL.md의 명령을 args 첫 칸에 그대로 넣어 실행하세요."}, 0)


def main(raw):
    if not raw or raw[0].strip() in ("", "-h", "--help"):
        usage()
    used = sorted({int(m.group(1)) for m in INPUT.finditer(raw[0])})
    if len(raw) > 1 and not used:
        argv = [str(a) for a in raw]                       # 이미 나뉜 인자
    else:
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
        if used:
            folder = os.path.join(HERE, "_inputs")
            os.makedirs(folder, exist_ok=True)
            paths = {}
            for n in used:
                ext = next((m.group(2) for m in INPUT.finditer(raw[0]) if int(m.group(1)) == n and m.group(2)), ".txt")
                path = os.path.join(folder, f"input{n}{ext}")
                with open(path, "w", encoding="utf-8", newline="") as f:
                    f.write(str(contents[n - 1]))
                paths[n] = path
            argv = [INPUT.sub(lambda m: paths[int(m.group(1))], a) for a in argv]
    if argv and os.path.splitext(os.path.basename(argv[0]).lower())[0] in ("python", "python3", "py"):
        argv = argv[1:]
    if not argv:
        stop("실행할 스크립트가 없습니다. 예: " + (f'"{SCRIPTS[0]} -h"' if SCRIPTS else '"-h"'))
    script = os.path.basename(argv[0].replace("\\", "/"))
    if script not in SCRIPTS:
        if script in OTHER_SCRIPTS:
            stop(f"'{script}'는 {OTHER_TOOL} 스킬의 명령입니다. 같은 명령을 "
                 f'run_skill_script(name="{OTHER_TOOL}", path="run.py", args=["<같은 명령>"])로 실행하세요.')
        stop(f"'{argv[0]}'는 {TOOL} 스킬의 명령이 아닙니다. 쓸 수 있는 스크립트: {', '.join(SCRIPTS)}")
    path = os.path.join(HERE, script)
    sys.argv = [path] + argv[1:]
    sys.path[0] = HERE
    runpy.run_path(path, run_name="__main__")


if __name__ == "__main__":
    main(sys.argv[1:])
