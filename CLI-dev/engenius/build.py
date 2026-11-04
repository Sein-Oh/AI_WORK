r"""En-Genius 스킬 가져오기용 zip(docs/engenius/<tool>.zip)과 에이전트 XML(docs/engenius/오피스ML.xml)을 만듭니다.

zip 구성: <tool>/SKILL.md (En-Genius용으로 바꾼 지시문) + <tool>/run.py (CLI-dev/engenius/<tool>/run.py, 도구마다 따로).
도구 코드는 넣지 않는다: run.py가 사용자 PC의 C:\CLI\<tool>을 실행한다 (OpenCode와 같은 폴더). run.py의 SKILL_HASH에
지시문 원본(CLI/<tool>/SKILL.md, 버전 줄 제외)의 지문을 채워, C:\CLI의 SKILL.md가 바뀌었으면(명령 안내 변경) zip을 다시 가져오라고 안내하게 한다.
도구만 고친 업데이트는 zip을 바꾸지 않아도 된다.
SKILL.md는 CLI/<tool>/SKILL.md 본문 앞에 En-Genius 실행 방법을 붙이고, OpenCode·PowerShell 전용 문장만 바꿉니다.
CLI/<tool>/SKILL.md가 바뀌어 바꿀 문장을 못 찾으면 멈춥니다 (그 문장에 맞게 REPLACE를 고칩니다).
같은 입력이면 같은 zip이 나옵니다 (시각 고정). tests/test_engenius.py가 저장소의 zip이 최신인지 확인합니다.

사용: python CLI-dev/engenius/build.py   (저장소 루트에서)
"""
import importlib.util
import io
import os
import re
import sys
import zipfile
from xml.sax.saxutils import escape

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
OUT = os.path.join(ROOT, "docs", "engenius")
TOOLS = ("ml-kit", "office-cli")
STAMP = (2026, 1, 1, 0, 0, 0)

DESCRIPTION = {
    "ml-kit": "엑셀/CSV 표 데이터로 예측 모델 학습, 새 조건 예측, 목표값이 되는 조건 찾기, 표 보간, Word 보고서와 엑셀 성능예측표 "
              "만들기. 사용자 PC에서 실행(run.py).",
    "office-cli": "열린 Excel/Word/PowerPoint 읽기·쓰기(표, 차트, 서식, 슬라이드), 앱 사이 차트·표 옮기기, Outlook 메일, "
                  "PDF/HWP/Office 문서 검색(근거 포함), 되돌리기. 사용자 PC에서 실행(run.py).",
}

HEAD = """# 실행 방법 (En-Genius, 반드시 따른다) - {tool} {version}
0. 도구는 사용자 PC의 `C:\\CLI\\{tool}`에 있고, 이 스킬의 run.py가 그것을 실행한다. 설치 문제(error_type "setup")는 사용자에게 그대로 알린다.
1. 이 문서의 명령은 모두 run_skill_script로 실행한다. path는 항상 "run.py"이고, args 첫 칸에 명령 한 줄을 그대로 넣는다.
   예: run_skill_script(name="{tool}", path="run.py", args=["{example}"]{timeout_example})
   - 명령 앞의 `python`과 `C:\\CLI\\...\\` 같은 스크립트 경로는 있어도 되고 없어도 된다.
   - 결과의 `next_commands[i].command`와 `[추천]` 뒤의 전체 명령도 글자 그대로 args 첫 칸에 넣는다. 명령을 직접 고쳐 쓰지 않는다.
   - 다른 스크립트(.py)를 직접 실행하지 않는다. 항상 run.py로 실행한다.
2. timeout: {timeout_rule}
3. 실행할 때마다 사용자 승인 창이 뜬다. 사용자가 거절하면 다시 부르지 말고 멈춘다.
4. 표, 여러 줄 글, 큰따옴표가 든 값은 명령에 `@input1.csv`, `@input2.md`처럼 쓰고, 그 내용을 args 두 번째 칸부터 순서대로 넣는다.
   run.py가 내용을 UTF-8 파일로 만들어 그 자리에 경로를 넣는다. 한 칸은 2000자까지다. 더 길면 나눠서 여러 번 실행한다.
   예: run_skill_script(name="office-cli", path="run.py", args=["word.py write --file @input1.md", "## 제목\\n- 항목 1\\n- 항목 2"])
5. 파일 경로는 사용자 PC의 전체 경로로 쓴다 (예: "C:\\Users\\<사용자>\\Desktop\\공정.xlsx"). 모르면 사용자에게 묻는다.
6. 결과는 stdout의 마지막 1줄 JSON이다. `[PROGRESS]`, `[INFO]`, `[WARN]` 줄은 무시한다.
7. office-cli 명령(excel.py, word.py, ppt.py, outlook.py, reader.py, office.py)은 office-cli 스킬로,
   ml-kit 명령(ml_train.py, ml_predict.py, ml_report.py, ml_map.py, ml_export.py)은 ml-kit 스킬로 실행한다.
8. 인터넷이 막힌 PC다. 설치를 시도하지 않는다. 결과에 `install_command`가 있으면 사용자에게 전달만 한다.
9. 도구로 할 수 있는 일을 파이썬 코드를 직접 짜서 하지 않는다.

"""
EXAMPLE = {
    "ml-kit": ('ml_train.py -f \\"C:\\\\데이터\\\\공정.xlsx\\" -x \\"온도,압력\\" -y \\"수율\\"', ", timeout=600",
               "학습·보고서·성능예측표(ml_train.py, ml_report.py, ml_export.py)는 600을 준다. 예측과 표 보간은 생략한다(120초)."),
    "office-cli": ("excel.py read --range A1:D20", "",
                   "보통은 생략한다(120초). 큰 폴더 등록(reader.py index)과 PDF 내보내기는 600을 준다."),
}

# CLI/<tool>/SKILL.md의 OpenCode·PowerShell 전용 문장 -> En-Genius 문장. 각 문장은 정확히 1번 나와야 한다.
REPLACE = {
    "ml-kit": [
        ("2. PowerShell에서는 명령 앞에 전역 지침의 인코딩 줄(`$env:PYTHONIOENCODING='utf-8'; ...`)을 붙인다.\n", ""),
        ("6. 학습은 몇 분 걸린다. 명령 실행 도구의 시간 제한은 600000(10분)으로 준다.",
         "6. 학습은 몇 분 걸린다. run_skill_script의 timeout은 600으로 준다."),
        ("도구 위치는 `C:\\CLI\\ml-kit` (office-cli는 `C:\\CLI\\office-cli`)이다. 아래 경로를 그대로 쓴다.",
         "아래 명령의 `C:\\CLI\\...` 경로는 그대로 두어도 된다 (run.py가 처리한다)."),
        ("파일 쓰기 도구로 한 줄에 피처 하나씩 쓴 `features.txt`(UTF-8)를 데이터 파일 옆에 만들고 `-x \"<그 경로>\"`",
         "`-x @input1.txt`로 쓰고, args 두 번째 칸에 피처 이름을 한 줄에 하나씩 넣는다"),
        ("파일 쓰기 도구로 `slots.json`(UTF-8)을 만들고 명령 끝에 `--slots \"<그 경로>\"`를 붙인다.",
         "명령 끝에 `--slots @input1.json`을 붙이고, args 두 번째 칸에 JSON 내용을 넣는다."),
    ],
    "office-cli": [
        ("1. 도구 위치는 `C:\\CLI\\office-cli`이다. 모든 명령은 `python \"C:\\CLI\\office-cli\\<파일>.py\" <명령> ...` 형식이다 "
         "(예: `python \"C:\\CLI\\office-cli\\excel.py\" read`).",
         "1. 모든 명령은 맨 위 '실행 방법'대로 run.py로 실행한다 (예: args=[\"excel.py read\"])."),
        ("2. PowerShell에서는 명령 앞에 전역 지침의 인코딩 줄(`$env:PYTHONIOENCODING='utf-8'; ...`)을 붙인다.\n", ""),
        ("파일 쓰기 도구로 UTF-8 파일(.csv, .md)을 만든 뒤 `--file \"<그 경로>\"`로 넘긴다.",
         "`--file @input1.csv`(글은 `@input1.md`)로 쓰고 내용은 args 두 번째 칸에 넣는다 (실행 방법 4)."),
        ("7. 값에 `$`가 있으면(엑셀 수식 `=$A$1` 등) 작은따옴표 `'...'`로 감싼다.",
         "7. 값에 공백이 있으면 큰따옴표 `\"...\"`로 감싼다. `$`는 그대로 써도 된다 (엑셀 수식 `=$A$1`)."),
    ],
}


def _version(tool):
    with open(os.path.join(ROOT, "CLI", tool, "SKILL.md"), encoding="utf-8") as f:
        return re.search(r'version:\s*"([^"]+)"', f.read()).group(1)


def _renumber(body):
    """'꼭 지킬 규칙' 목록에서 지운 항목 뒤 번호를 다시 매긴다."""
    m = re.search(r"## 꼭 지킬 규칙\n", body)
    if not m:
        return body
    start = m.end()
    end = body.index("\n\n", start)
    items = re.split(r"\n(?=\d+\. )", body[start:end])
    return body[:start] + "\n".join(re.sub(r"^\d+\. ", f"{i}. ", it) for i, it in enumerate(items, 1)) + body[end:]


def skill_md(tool):
    with open(os.path.join(ROOT, "CLI", tool, "SKILL.md"), encoding="utf-8") as f:
        text = f.read()
    body = text.split("---", 2)[2].lstrip("\n")
    body = body.replace("(시간 제한 600000)", "(timeout 600)")
    for old, new in REPLACE[tool]:
        n = body.count(old)
        if n != 1:
            raise SystemExit(f"[ERROR] {tool}/SKILL.md에서 바꿀 문장이 {n}번 나옵니다 (1번이어야 함): {old[:60]}")
        body = body.replace(old, new)
    body = _renumber(body)
    if "파일 쓰기 도구" in body or "PowerShell" in body:
        raise SystemExit(f"[ERROR] {tool}/SKILL.md에 En-Genius에 없는 도구 안내가 남았습니다 (REPLACE에 추가)")
    example, timeout_example, timeout_rule = EXAMPLE[tool]
    head = HEAD.format(tool=tool, version=_version(tool), example=example, timeout_example=timeout_example,
                       timeout_rule=timeout_rule)
    desc = DESCRIPTION[tool]
    return f'---\nname: {tool}\ndescription: "{desc}"\n---\n\n{head}{body}'


def run_py(tool):
    """CLI-dev/engenius/<tool>/run.py에 지시문 버전을 채운 것 (도구마다 따로 둔 실행 입구)."""
    with open(os.path.join(HERE, tool, "run.py"), encoding="utf-8") as f:
        text = f.read()
    spec = importlib.util.spec_from_file_location(f"engenius_run_{tool}", os.path.join(HERE, tool, "run.py"))
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)                                    # 지문 계산을 run.py와 똑같이
    with open(os.path.join(ROOT, "CLI", tool, "SKILL.md"), encoding="utf-8") as f:
        digest = mod.skill_hash(f.read())
    for line, value in (('SKILL_VERSION = ""', _version(tool)), ('SKILL_HASH = ""', digest)):
        if text.count(line) != 1:
            raise SystemExit(f"[ERROR] {tool}/run.py에 '{line}' 줄이 1번 있어야 합니다")
        text = text.replace(line, f'{line[:-2]}"{value}"')
    return text


def zip_bytes(tool):
    files = [("SKILL.md", skill_md(tool).encode("utf-8"))]
    files.append(("run.py", run_py(tool).encode("utf-8")))
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in files:
            info = zipfile.ZipInfo(f"{tool}/{name}", date_time=STAMP)
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = 0o644 << 16
            z.writestr(info, data)
    return buf.getvalue()


AGENT_NAME = "오피스ML"
AGENT_XML = os.path.join(OUT, f"{AGENT_NAME}.xml")
AGENT_DESC = "office-cli, ml-kit 스킬로 Excel/Word/PPT/Outlook 작업과 표 데이터 예측 모델 작업을 하는 도우미 (사용자 PC에서 실행)"


def agent_xml():
    """En-Genius 에이전트 XML(format 1, docs/agent-xml.md): 이름, 설명, 역할 프롬프트(에이전트_프롬프트.md), 도구 제한.
    도구 서버는 고르지 않고(빈 Servers = 내장 Office·Outlook 도구 없음), 도구 제한은 run_skill_script 하나.
    스킬은 넣지 않는다: XML의 Skill은 지시문 사본만이라 run.py가 빠진다 -> zip으로 가져온 스킬을 설정 창에서 붙인다.
    LLM(Provider)도 넣지 않는다: 회사마다 이름이 달라 가져온 뒤 설정 창에서 고른다."""
    with open(os.path.join(OUT, "에이전트_프롬프트.md"), encoding="utf-8") as f:
        prompt = f.read().split("\n", 2)[2].strip()        # 첫 줄(붙여 넣기 안내 제목)과 빈 줄을 뺀 본문
    return ('<?xml version="1.0" encoding="UTF-8"?>\n'
            '<EnGeniusAgent format="1" exported-at="2026-11-04T00:00:00+09:00">\n'
            f"  <Name>{escape(AGENT_NAME)}</Name>\n"
            f"  <Description>{escape(AGENT_DESC)}</Description>\n"
            f"  <Prompt>{escape(prompt)}</Prompt>\n"
            "  <Servers/>\n"
            "  <Tools>\n"
            '    <Tool name="run_skill_script"/>\n'
            "  </Tools>\n"
            "</EnGeniusAgent>\n")


def main():
    os.makedirs(OUT, exist_ok=True)
    for tool in TOOLS:
        path = os.path.join(OUT, f"{tool}.zip")
        data = zip_bytes(tool)
        with open(path, "wb") as f:
            f.write(data)
        print(f"[INFO] {path} ({len(data):,} bytes)", file=sys.stderr)
    with open(AGENT_XML, "w", encoding="utf-8", newline="\n") as f:
        f.write(agent_xml())
    print(f"[INFO] {AGENT_XML}", file=sys.stderr)


if __name__ == "__main__":
    main()
