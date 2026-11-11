"""skill.py - 사용자가 만든 OpenCode 스킬 등록·점검 (Office 불필요)

명령:
    save    스킬 등록: --name <영문 이름> --description "<언제 쓰는 스킬인지>" --file <본문.md> [--project <폴더>]
            머리말(name, description)은 도구가 붙이고, 형식·크기·문자를 확인한 뒤 저장하고 다시 읽어 확인
    list    스킬 목록 (기본 / 사용자 / --project를 주면 그 폴더의 프로젝트 스킬)
    check   스킬 하나를 점검: --name <이름> [--project <폴더>]
    delete  사용자가 만든 스킬 지우기: --name <이름> [--project <폴더>] --dry-run (기본 스킬은 지우지 않음)

스킬 위치 (OpenCode 2.2.5에서 확인, 2026-11-12): 폴더 하나 = 스킬 하나 (SKILL.md와 그 스킬의 scripts\\)
    기본·사용자 스킬 C:\\AIKIT\\skills\\<이름>\\SKILL.md (OpenCode 설정 폴더 = OPENCODE_CONFIG_DIR = C:\\AIKIT, 어느 작업 폴더에서나 보임).
                    기본 스킬 12개는 도구와 함께 배포되고(BASE_SKILLS), 나머지 이름은 사용자가 만든 스킬.
    프로젝트 스킬    <폴더>\\.opencode\\skills\\<이름>\\SKILL.md (--project. 그 폴더에서 연 OpenCode 세션에서만 보임)
경량 LLM이 머리말 형식(--- 줄, name, description)을 틀리지 않게 본문만 받아 도구가 조립한다.
"""
import os
import re
import shutil
import sys
from typing import Any, Dict, List, Optional, Sequence, Tuple

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _runtime as rt  # noqa: E402

TOOL = "skill"
SKILL_MAX_BYTES = 8192                     # 스킬 하나의 크기 한도 (LF 기준, 저장소 규칙과 같음)
DESC_MIN, DESC_MAX = 15, 300               # description 글자 수 (짧으면 모델이 언제 쓸지 모름, 길면 목록이 길어짐)
NAME_RE = re.compile(r"^[a-z0-9]+(-[a-z0-9]+)*$")
NAME_MAX = 40
LIST_MAX = 60
# 도구와 함께 배포하는 기본 스킬: 같은 이름으로 만들거나 지울 수 없음 (C:\AIKIT\skills에 사용자 스킬과 나란히 있음)
BASE_SKILLS = ("ml-train", "ml-dataset", "ml-predict", "ml-report", "ml-map",
               "excel-kit", "word-kit", "ppt-kit", "outlook-kit", "search-kit", "stats-kit", "skill-kit", "office-cli-test")
# 1.6.7에서 이름을 바꾼 기본 스킬의 옛 이름: setup.py가 C:\AIKIT\skills에 남은 것을 정리하므로 이 이름으로 만들지 않음
RENAMED_SKILLS = ("office-excel", "office-doc", "office-mail", "office-spec", "office-data", "skill-maker")
_PATH_IN_TEXT = re.compile(r"[A-Za-z]:\\[^\s\"'`<>|]+\.(?:py|bat)")
_FRONT = re.compile(r"\A\s*---\s*\n(.*?)\n---\s*\n", re.S)


KIT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))     # C:\\AIKIT (도구 폴더들의 부모)
PROJECT_SUB = os.path.join(".opencode", "skills")


def config_dir() -> str:
    """OpenCode 설정 폴더: OPENCODE_CONFIG_DIR (setup.bat이 사용자 환경 변수로 둠), 없으면 도구 폴더의 부모(C:\\AIKIT)."""
    return os.environ.get("OPENCODE_CONFIG_DIR") or KIT_DIR


def project_dir(project: Optional[str]) -> Optional[str]:
    """--project 폴더 (있는 폴더만). 없으면 STOP_RETRY (만들지 않음: 오타 경로에 .opencode가 생기지 않게)."""
    if not project:
        return None
    path = os.path.abspath(os.path.expandvars(str(project).strip().strip('"')))
    if not os.path.isdir(path):
        raise rt.InputError(f"--project 폴더가 없습니다: {path}. 사용자에게 프로젝트 폴더의 전체 경로를 확인하세요.")
    return path


def skills_dir(project: Optional[str] = None) -> str:
    """스킬을 저장·찾는 폴더: --project가 있으면 <폴더>\\.opencode\\skills, 없으면 <설정 폴더>\\skills (기본 스킬과 같은 곳)."""
    root = project_dir(project)
    return os.path.join(root, PROJECT_SUB) if root else os.path.join(config_dir(), "skills")


def skill_path(name: str, project: Optional[str] = None) -> str:
    return os.path.join(skills_dir(project), name, "SKILL.md")


def _fix_name(raw: str) -> str:
    """영문 대문자·공백·밑줄을 스킬 이름 규칙으로 고친 값 (고칠 수 없으면 '')."""
    s = re.sub(r"[\s_]+", "-", str(raw).strip().lower())
    s = re.sub(r"[^a-z0-9-]", "", s)
    s = re.sub(r"-{2,}", "-", s).strip("-")
    return s if NAME_RE.match(s or "") else ""


def check_name(raw: str) -> str:
    name = str(raw or "").strip()
    if NAME_RE.match(name) and len(name) <= NAME_MAX:
        return name
    fixed = _fix_name(name)
    if fixed and fixed != name and len(fixed) <= NAME_MAX:
        raise rt.InputError(f"스킬 이름은 영문 소문자, 숫자, '-'만 씁니다 (받은 값: {name}). 같은 뜻으로 고친 이름: {fixed}. "
                            f"[추천] {rt.retry_command({'--name': fixed})}")
    raise rt.InputError(f"스킬 이름 '{name}'을(를) 쓸 수 없습니다. 영문 소문자, 숫자, '-'만 쓰고 {NAME_MAX}자 이하로 정합니다 "
                        "(예: weekly-report, bom-check). 한글 이름은 쓸 수 없으니 사용자에게 영문 이름을 정해 달라고 하세요.")


def _bad_chars(text: str) -> List[str]:
    """CP949로 표현되지 않는 문자 (이모지, 장식 기호): 폐쇄망 PC 콘솔에서 깨지고 경량 모델이 흔들림."""
    out: List[str] = []
    for ch in text:
        try:
            ch.encode("cp949")
        except UnicodeEncodeError:
            if ch not in out:
                out.append(ch)
    return out


def _parse(text: str) -> Tuple[Dict[str, str], str]:
    """SKILL.md -> (머리말 {name, description, ...}, 본문). 머리말이 없으면 ({}, 전체)."""
    m = _FRONT.match(text)
    if not m:
        return {}, text
    meta: Dict[str, str] = {}
    for line in m.group(1).splitlines():
        if ":" in line and not line.startswith((" ", "\t")):
            k, v = line.split(":", 1)
            meta[k.strip()] = v.strip().strip('"').strip("'")
    return meta, text[m.end():]


def _problems(name: str, description: str, body: str, full: str) -> Tuple[List[str], List[str]]:
    """(고쳐야 하는 문제, 주의). 문제가 있으면 저장하지 않음."""
    problems: List[str] = []
    warnings: List[str] = []
    desc = " ".join(description.split())
    if len(desc) < DESC_MIN:
        problems.append(f"description이 너무 짧습니다 ({len(desc)}자, {DESC_MIN}자 이상). 무엇을 하는 스킬이고 어떤 요청에 쓰는지 한두 문장으로 적습니다 "
                        "(예: \"주간 실적 엑셀을 부서별로 집계해 보고서 표로 만든다. 주간 보고 정리 요청에 사용.\")")
    if len(desc) > DESC_MAX:
        problems.append(f"description이 너무 깁니다 ({len(desc)}자, {DESC_MAX}자 이하). 한두 문장으로 줄입니다")
    if not body.strip():
        problems.append("본문이 비어 있습니다. 언제 쓰는지, 순서, 실행할 명령을 적습니다")
    size = len(full.replace("\r\n", "\n").encode("utf-8"))
    if size > SKILL_MAX_BYTES:
        problems.append(f"스킬이 너무 큽니다 ({size}바이트, {SKILL_MAX_BYTES}바이트 이하). 설명을 줄이거나 작업을 두 스킬로 나눕니다")
    bad = _bad_chars(full)
    if bad:
        problems.append("이 PC에서 깨지는 문자가 있습니다 (이모지·장식 기호): " + " ".join(f"U+{ord(c):04X}" for c in bad[:10])
                        + ". 글자로 바꿉니다 (화살표는 ->, 표시는 [OK])")
    if body.strip() and "`" not in body and not re.search(r"(?m)^\s*(\d+\.|-)\s", body):
        warnings.append("본문에 순서(1. 2.)나 명령(`...`)이 없습니다. 모델이 따라 할 순서와 실행할 명령을 적으면 더 잘 됩니다")
    if os.name == "nt":
        missing = [p for p in dict.fromkeys(_PATH_IN_TEXT.findall(body)) if not os.path.isfile(p)]
        if missing:
            warnings.append("본문의 스크립트 파일이 없습니다: " + ", ".join(missing[:5]) + " (경로를 확인하거나 파일을 그 자리에 두세요)")
    if re.search(r"(?i)pip install|https?://", body):
        warnings.append("본문에 설치(pip install)나 인터넷 주소가 있습니다. 이 PC는 폐쇄망이라 쓸 수 없습니다")
    return problems, warnings


def _compose(name: str, description: str, body: str) -> str:
    desc = " ".join(description.split()).replace('"', "'")
    return f"---\nname: {name}\ndescription: {desc}\n---\n\n{body.strip()}\n"


def cmd_save(args) -> Dict[str, Any]:
    name = check_name(args.name)
    if name in BASE_SKILLS or name in RENAMED_SKILLS:
        raise rt.InputError(f"'{name}'은(는) 도구와 함께 배포하는 기본 스킬 이름이라 쓸 수 없습니다. 다른 이름을 사용자에게 정해 달라고 하세요 "
                            f"(예: my-{name}).")
    src = rt.resolve_input_path(args.file, "스킬 본문 파일", exts=(".md", ".txt"))
    raw = rt.read_text_file(src)
    meta, body = _parse(raw)
    warnings: List[str] = []
    if meta:
        warnings.append("본문 파일의 머리말(--- 사이의 name, description)은 쓰지 않고 --name, --description으로 새로 붙였습니다")
    description = args.description or ""
    full = _compose(name, description, body)
    problems, more = _problems(name, description, body, full)
    warnings += more
    if problems:
        raise rt.InputError("스킬을 저장하지 않았습니다. 고칠 것: " + " / ".join(f"{i}) {p}" for i, p in enumerate(problems, 1)),
                            directive="저장하지 않았습니다. message의 고칠 것을 본문 파일이나 --description에서 고친 뒤 같은 명령을 1번 다시 실행하세요. "
                                      "사용자가 정해야 하는 것(이름, 설명)은 사용자에게 물으세요.")
    dst = skill_path(name, args.project)
    exists = os.path.isfile(dst)
    size = len(full.encode("utf-8"))
    fields: Dict[str, Any] = {"name": name, "skill_file": dst, "bytes": size, "limit_bytes": SKILL_MAX_BYTES, "updated": exists,
                              "kind": "프로젝트" if args.project else "사용자"}
    if warnings:
        fields["warnings"] = warnings
    if exists and not args.overwrite and not args.dry_run:
        raise rt.UserDecisionError(
            f"스킬 '{name}'이(가) 이미 있습니다 ({dst}).", "저장하지 않았습니다. 사용자에게 있던 스킬을 새 내용으로 바꿀지 묻고, "
            "바꾼다고 하면 confirm_commands[0]을 1번 실행하세요. 아니면 다른 이름을 정해 달라고 하세요.",
            candidates=["있던 스킬을 새 내용으로 바꾸기"], commands=[rt.retry_command({"--overwrite": True})], destructive=True)
    if args.dry_run:
        what = "바꿉니다" if exists else "새로 등록합니다"
        return rt.dry_run_result(TOOL, "save", f"미리보기: 스킬 '{name}'을(를) {what} ({size}바이트, 저장하지 않음)", 1,
                                 [f"{dst} ({size}바이트)"], rt.dry_run_command(TOOL, f"스킬 '{name}' 저장", destructive=exists, overwrite=exists),
                                 **fields)
    os.makedirs(os.path.dirname(dst), exist_ok=True)
    rt.safe_write_text(dst, full)
    back = rt.read_text_file(dst).replace("\r\n", "\n")
    got_meta, _ = _parse(back)
    where = (f"{project_dir(args.project)} 폴더에서 연 OpenCode 세션에서만 쓰입니다" if args.project
             else "어느 작업 폴더에서 연 OpenCode 세션에서나 쓰입니다")
    preview = (f"스킬 '{name}'을(를) {'바꿨습니다' if exists else '등록했습니다'} ({size}바이트). {where}. OpenCode를 모두 닫았다가 다시 시작하면 "
               f"쓸 수 있습니다 (안 보이면 {os.path.join(KIT_DIR, 'web.bat')} stop 뒤 다시 실행)")
    if back != full or got_meta.get("name") != name:
        return rt.ok_unverified(TOOL, "save", preview, ["저장한 스킬 파일을 다시 읽은 내용이 쓴 내용과 다름"], **fields)
    return rt.ok(TOOL, "save", preview,
                 "VERIFIED: 스킬 파일을 저장하고 다시 읽어 확인했습니다. 사용자에게 preview와 skill_file을 알리고, OpenCode를 다시 시작한 뒤 "
                 "새 대화에서 description에 적은 요청을 말하면 이 스킬이 쓰인다고 안내하세요. warnings가 있으면 함께 알리세요. 같은 명령을 다시 실행하지 마세요.",
                 verified=True, **fields)


def _folders(project: Optional[str] = None) -> List[Tuple[str, str]]:
    """(종류, 스킬 폴더): 프로젝트(--project일 때), 설정 폴더의 skills 순서. 같은 이름이면 앞의 것을 찾음.
    설정 폴더의 종류 "스킬"은 이름으로 기본/사용자를 나눔 (_kind)."""
    out = [("프로젝트", skills_dir(project))] if project else []
    return out + [("스킬", skills_dir())]


def _kind(folder_kind: str, name: str) -> str:
    return folder_kind if folder_kind == "프로젝트" else "기본" if name in BASE_SKILLS else "사용자"


def _installed(project: Optional[str] = None) -> List[Dict[str, Any]]:
    out: List[Dict[str, Any]] = []
    for kind, root in _folders(project):
        if not os.path.isdir(root):
            continue
        for name in sorted(os.listdir(root)):
            path = os.path.join(root, name, "SKILL.md")
            if not os.path.isfile(path):
                continue
            try:
                text = rt.read_text_file(path)
            except Exception:
                text = ""
            meta, _ = _parse(text)
            out.append({"name": name, "description": rt.clamp_text(meta.get("description", ""), 120)[0],
                        "kind": _kind(kind, name), "bytes": len(text.replace("\r\n", "\n").encode("utf-8")), "skill_file": path})
    return out


def cmd_list(args) -> Dict[str, Any]:
    items = _installed(args.project)
    mine = [x for x in items if x["kind"] != "기본"]
    dirs = {kind: root for kind, root in _folders(args.project)}
    if not items:
        return rt.ok(TOOL, "list", f"0건: 스킬이 없습니다 ({', '.join(dirs.values())})",
                     f"VERIFIED: 0건. 찾은 스킬 없음. 기본 스킬 폴더({skills_dir()})가 비었으면 C:\\AIKIT 폴더를 다시 복사하라고 알리세요.",
                     verified=True, skills=[], count=0, user_count=0, skills_dir=skills_dir(args.project))
    shown = items[:LIST_MAX]
    preview = f"스킬 {len(items)}개 (사용자·프로젝트가 만든 것 {len(mine)}개: " + (", ".join(x["name"] for x in mine[:10]) or "없음") + ")"
    return rt.ok(TOOL, "list", preview, "VERIFIED: 스킬 폴더를 읽었습니다. 이름, 종류(기본/사용자/프로젝트), description을 사용자에게 알리세요.",
                 verified=True, skills=shown, count=len(items), user_count=len(mine), truncated=len(items) > len(shown),
                 skills_dir=skills_dir(args.project))


def _find(name: str, project: Optional[str] = None) -> str:
    """스킬 파일: --project 폴더, 사용자 스킬, 기본 스킬 순서로 찾음."""
    for _kind, root in _folders(project):
        path = os.path.join(root, name, "SKILL.md")
        if os.path.isfile(path):
            return path
    names = sorted({x["name"] for x in _installed(project)})
    raise rt.not_found("스킬", name, names, extra=f" ({skills_dir(project)})")


def cmd_check(args) -> Dict[str, Any]:
    name = str(args.name or "").strip()
    path = _find(name, args.project)
    text = rt.read_text_file(path).replace("\r\n", "\n")
    meta, body = _parse(text)
    problems: List[str] = []
    if not meta:
        problems.append("파일 맨 위에 머리말(--- 줄 사이의 name, description)이 없습니다")
    elif meta.get("name") != name:
        problems.append(f"머리말의 name({meta.get('name', '')})이 폴더 이름({name})과 다릅니다")
    if not NAME_RE.match(name):
        problems.append("폴더 이름이 스킬 이름 규칙(영문 소문자, 숫자, '-')에 맞지 않습니다")
    more, warnings = _problems(name, meta.get("description", ""), body, text)
    problems += more
    fields: Dict[str, Any] = {"name": name, "skill_file": path, "bytes": len(text.encode("utf-8")), "limit_bytes": SKILL_MAX_BYTES,
                              "description": rt.clamp_text(meta.get("description", ""), 300)[0], "problems": problems, "count": len(problems)}
    if warnings:
        fields["warnings"] = warnings
    if problems:
        return rt.ok(TOOL, "check", f"스킬 '{name}': 고칠 것 {len(problems)}개 - " + " / ".join(problems)[:300],
                     "VERIFIED: 스킬 파일을 읽어 점검했습니다. problems를 사용자에게 그대로 알리세요. 고치려면 본문을 고친 파일로 "
                     "skill.py save --name <이름> --description \"<설명>\" --file <본문> --overwrite 를 사용자가 원할 때 실행합니다.",
                     verified=True, **fields)
    return rt.ok(TOOL, "check", f"스킬 '{name}': 문제 없음 ({fields['bytes']}바이트)",
                 "VERIFIED: 0건. 고칠 것 없음. 스킬 형식·크기·문자를 확인했습니다. 스킬이 안 보이면 OpenCode를 다시 시작하라고 알리세요.",
                 verified=True, **fields)


def cmd_delete(args) -> Dict[str, Any]:
    name = str(args.name or "").strip()
    path = _find(name, args.project)
    if name in BASE_SKILLS:
        raise rt.InputError(f"'{name}'은(는) 도구와 함께 배포하는 기본 스킬이라 지우지 않습니다.", error_type="not_supported")
    folder = os.path.dirname(path)
    if args.dry_run:
        return rt.dry_run_result(TOOL, "delete", f"미리보기: 스킬 '{name}'을(를) 지웁니다 ({folder}, 바꾸지 않음)", 1, [folder],
                                 rt.dry_run_command(TOOL, f"스킬 '{name}' 지우기"), name=name, skill_file=path)
    keep = os.path.join(rt.get_data_dir(), "deleted_skills", f"{name}_{rt.timestamp()}")
    os.makedirs(os.path.dirname(keep), exist_ok=True)
    shutil.copytree(folder, keep)
    shutil.rmtree(folder)
    if os.path.exists(folder):
        return rt.ok_unverified(TOOL, "delete", f"스킬 '{name}'을(를) 지웠습니다", ["지운 뒤에도 스킬 폴더가 남아 있음"], name=name, backup=keep)
    return rt.ok(TOOL, "delete", f"스킬 '{name}'을(를) 지웠습니다 (지운 내용은 {keep}에 보관). OpenCode를 다시 시작하면 목록에서 빠집니다",
                 "VERIFIED: 스킬 폴더를 지우고 없는 것을 확인했습니다. 사용자에게 preview를 알리고 다시 실행하지 마세요.",
                 verified=True, name=name, backup=keep)


def build_cli():
    parser, sub = rt.build_parser(TOOL, "사용자 스킬 등록·점검")
    p = sub.add_parser("save", help="스킬 등록 (머리말은 도구가 붙임, 형식·크기 확인 뒤 저장)")
    p.add_argument("--name", required=True, help="스킬 이름: 영문 소문자, 숫자, '-' (예: weekly-report)")
    p.add_argument("--description", required=True, help="무엇을 하는 스킬이고 어떤 요청에 쓰는지 한두 문장")
    p.add_argument("--file", required=True, help="스킬 본문 파일 (.md, UTF-8): 언제 쓰는지, 순서, 실행할 명령")
    p.add_argument("--overwrite", action="store_true", help="같은 이름의 스킬이 있으면 바꾸기 (사용자가 원할 때만)")
    p.add_argument("--dry-run", action="store_true", help="저장하지 않고 확인만")
    p.add_argument("--project", help="프로젝트 폴더: 그 폴더의 .opencode\\skills에 저장 (그 폴더에서 연 세션에서만 쓰임). 없으면 사용자 스킬")
    p = sub.add_parser("list", help="스킬 목록")
    p.add_argument("--project", help="이 프로젝트 폴더의 스킬도 함께")
    p = sub.add_parser("check", help="스킬 점검")
    p.add_argument("--name", required=True, help="스킬 이름")
    p.add_argument("--project", help="프로젝트 폴더의 스킬을 먼저 찾음")
    p = sub.add_parser("delete", help="사용자가 만든 스킬 지우기")
    p.add_argument("--name", required=True, help="스킬 이름")
    p.add_argument("--project", help="프로젝트 폴더의 스킬을 지움")
    p.add_argument("--dry-run", action="store_true", help="지우지 않고 확인만")
    return parser


HANDLERS = {"save": cmd_save, "list": cmd_list, "check": cmd_check, "delete": cmd_delete}


def main(argv: Optional[Sequence[str]] = None) -> None:
    args = build_cli().parse_args(argv)
    rt.run_cli(TOOL, args.command, lambda: HANDLERS[args.command](args))


if __name__ == "__main__":
    main()
