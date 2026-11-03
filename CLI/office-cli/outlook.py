"""outlook.py - 아웃룩 메일

조회 (화면은 그대로, 백그라운드):
    list     폴더의 최근 메일 목록 (--unread, --limit)
    search   제목/본문/보낸사람 검색
    read     메일 하나 읽기 (--id), 첨부 저장 (--save-attach)
    current  지금 보고 있거나 선택한 메일 읽기
    folders  폴더 구조와 안 읽은 메일 수

작성 (기본: 작성 창만 띄우고 사용자가 직접 [보내기]):
    new      새 메일
    reply    답장 (--all: 전체 답장). 원본 대화 이력 유지
    forward  전달

--send 는 사용자가 "보내줘", "바로 발송"처럼 발송을 분명히 요청한 경우에만 씁니다.
발송 전 수신자 주소를 확인하고, 확인되지 않으면 보내지 않고 작성 창으로 바꿉니다.

예:
    python outlook.py current
    python outlook.py search --query "견적" --from "김팀장" --limit 5
    python outlook.py read --id "<id>" --save-attach
    python outlook.py reply --body "확인했습니다."
    python outlook.py new --to "lead@company.com" --subject "주간 보고" --body "첨부합니다." --attach "보고서.pdf"
"""
import os
import re
import sys
from typing import Any, Callable, Dict, List, Optional, Sequence, Tuple

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _office as off  # noqa: E402
import _runtime as rt  # noqa: E402

TOOL = "outlook"
OL_MAIL = 43
DEFAULT_FOLDERS = {"inbox": 6, "받은편지함": 6, "sent": 5, "보낸편지함": 5, "drafts": 16, "임시보관함": 16,
                   "deleted": 3, "지운편지함": 3, "outbox": 4, "보낼편지함": 4, "junk": 23, "정크메일": 23}
BLOCKED_ATTACH = {".exe", ".bat", ".cmd", ".com", ".js", ".vbs", ".vbe", ".wsf", ".ps1", ".scr", ".msi", ".lnk",
                  ".hta", ".jar", ".reg", ".dll"}
SAVE_FAILED = "저장 실패"     # skipped_attachments에서 정책상 건너뛴 것과 저장에 실패한 것을 구분하는 표시
MAX_LIMIT = 200          # list/search --limit 상한 (넘으면 줄이고 warnings로 알림)
PR_ATTACHMENT_HIDDEN = "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B"


# ==============================================================================
# 1. 연결과 폴더
# ==============================================================================

def get_outlook() -> Any:
    off.require_com()
    off.init_com()
    try:
        return off.win32com.client.GetActiveObject("Outlook.Application")
    except Exception:
        rt.log("PROGRESS", "Outlook이 실행 중이 아니어서 백그라운드로 연결합니다.")
        try:
            return off.win32com.client.Dispatch("Outlook.Application")
        except Exception as e:
            raise rt.InputError(f"Outlook에 연결할 수 없습니다 (설치/프로필 확인 필요): {e}")


def get_namespace(app: Any) -> Any:
    ns = app.GetNamespace("MAPI")
    try:
        ns.Logon("", "", False, False)
    except Exception:
        pass
    return ns


def resolve_folder(ns: Any, spec: Optional[str]) -> Any:
    """'inbox'/'받은편지함' 같은 기본 폴더, 또는 '계정\\폴더\\하위폴더' 경로 ('>'로 구분해도 됨)."""
    s = (spec or "inbox").strip()
    if s.lower() in DEFAULT_FOLDERS:
        return ns.GetDefaultFolder(DEFAULT_FOLDERS[s.lower()])
    parts = [p.strip() for p in re.split(r"\\|>", s) if p.strip()]
    stores = {str(f.Name): f for f in ns.Folders}
    found_store = rt.find_name(parts[0], stores)
    alias = re.sub(r"\s+", "", parts[0]).lower()
    if found_store:
        cur = stores[found_store]
        parts = parts[1:]
    elif alias in {re.sub(r"\s+", "", k).lower(): k for k in DEFAULT_FOLDERS}:
        # 'inbox\프로젝트', '받은편지함\프로젝트': 첫 부분이 기본 폴더 별칭이면 그 폴더에서 시작
        key = {re.sub(r"\s+", "", k).lower(): k for k in DEFAULT_FOLDERS}[alias]
        cur = ns.GetDefaultFolder(DEFAULT_FOLDERS[key])
        parts = parts[1:]
    else:
        cur = ns.GetDefaultFolder(6).Parent          # 기본 계정의 최상위
    walked: List[str] = [str(cur.Name)]
    for part in parts:
        subs = {str(f.Name): f for f in cur.Folders}
        name = rt.find_name(part, subs)
        if name is None:
            # 후보도 전체 경로로 (일부만 주면 LLM이 상위 폴더 전체를 검색함)
            raise rt.not_found("메일 폴더", "\\".join(walked + [part]), ["\\".join(walked + [n]) for n in subs],
                               extra=f" (경로: {s})")
        cur = subs[name]
        walked.append(name)
    return cur


def _text(v: Any) -> str:
    return "" if v is None else str(v)


def summary(item: Any) -> Dict[str, Any]:
    sender = _text(getattr(item, "SenderName", ""))
    addr = _text(getattr(item, "SenderEmailAddress", ""))
    try:
        received = rt.json_safe(item.ReceivedTime)
    except Exception:
        received = ""
    try:
        attach = int(item.Attachments.Count)
    except Exception:
        attach = 0
    return {"id": _text(getattr(item, "EntryID", "")), "subject": _text(getattr(item, "Subject", "")),
            "from": f"{sender} <{addr}>" if addr and "@" in addr else sender, "received": received,
            "unread": bool(getattr(item, "UnRead", False)), "attachments": attach}


def _collect(items: Any, limit: int, keep: Optional[Callable[[Any], bool]] = None) -> List[Dict[str, Any]]:
    out: List[Dict[str, Any]] = []
    for item in items:
        if len(out) >= limit:
            break
        try:
            if int(item.Class) == OL_MAIL and (keep is None or keep(item)):
                out.append(summary(item))
        except Exception:
            continue
    return out


def _check_limit(limit: int, warnings: List[str]) -> int:
    if limit < 1:
        raise rt.InputError(f"--limit은 1 이상이어야 합니다 (받은 값: {limit}).")
    if limit > MAX_LIMIT:
        msg = f"--limit {limit}은(는) 너무 많아 {MAX_LIMIT}통으로 줄였습니다"
        rt.log("WARN", msg)
        warnings.append(msg)
        return MAX_LIMIT
    return limit


# ==============================================================================
# 2. 조회
# ==============================================================================

def cmd_list(args) -> Dict[str, Any]:
    ns = get_namespace(get_outlook())
    folder = resolve_folder(ns, args.folder)
    items = folder.Items
    if args.unread:
        items = items.Restrict("[UnRead] = True")
    items.Sort("[ReceivedTime]", True)                 # 거른 뒤 정렬해야 순서가 보장됨
    warnings: List[str] = []
    mails = _collect(items, _check_limit(args.limit, warnings))
    extra = {"warnings": warnings} if warnings else {}
    what = "안 읽은 메일" if args.unread else "메일"
    if not mails:      # R2 0건
        return rt.ok(TOOL, "list", f"0건: '{folder.Name}' 폴더에 {what}이 없습니다",
                     f"VERIFIED: 0건. '{folder.Name}' 폴더에 {what}이 없습니다. 사용자에게 그대로 알리고 조건을 바꿔 다시 실행하지 마세요.",
                     folder=str(folder.Name), mails=[], count=0, **extra)
    return rt.ok(TOOL, "list", f"'{folder.Name}' 폴더 최근 메일 {len(mails)}통" + (" (안 읽은 메일)" if args.unread else ""),
                 folder=str(folder.Name), mails=mails, count=len(mails), **extra)


def _dasl_like(fields: Sequence[str], value: str) -> str:
    # 작은따옴표만 겹쳐 씁니다. % _는 그대로 둠: DASL 와일드카드지만 자기 자신과도 맞으므로 빠지는 메일이 없음
    # (예전에는 _를 공백으로 바꾸고 %를 지워 '견적서_v2', '10% 할인'이 절대 맞지 않고 0건 VERIFIED가 나왔음)
    v = value.replace("'", "''")
    return "(" + " OR ".join(f'"{f}" LIKE \'%{v}%\'' for f in fields) + ")"


def _literal_keep(query: Optional[str], sender: Optional[str], where: str) -> Optional[Callable[[Any], bool]]:
    """검색어에 % _가 있으면 DASL이 더 넓게 찾으므로, 글자 그대로 들어 있는 메일만 남깁니다 (대소문자 무시).
    읽지 못한 칸은 맞는 것으로 봅니다 (빠뜨리는 것보다 하나 더 보여 주는 편이 안전)."""
    def wild(s: Optional[str]) -> bool:
        return bool(s) and ("%" in s or "_" in s)
    if not (wild(query) or wild(sender)):
        return None

    def has(item: Any, attrs: Sequence[str], needle: str) -> bool:
        n = needle.casefold()
        for a in attrs:
            try:
                if n in _text(getattr(item, a)).casefold():
                    return True
            except Exception:
                return True
        return False

    q_attrs = {"subject": ["Subject"], "body": ["Body"], "all": ["Subject", "Body"]}[where]

    def keep(item: Any) -> bool:
        if wild(query) and not has(item, q_attrs, query or ""):
            return False
        if wild(sender) and not has(item, ["SenderName", "SenderEmailAddress"], sender or ""):
            return False
        return True
    return keep


def build_search_filter(query: Optional[str], sender: Optional[str], where: str) -> str:
    """DASL 필터: @SQL= 접두어는 한 번만, 조건마다 괄호로 묶어 AND로 연결."""
    clauses = []
    if query:
        fields = {"subject": ["urn:schemas:httpmail:subject"],
                  "body": ["urn:schemas:httpmail:textdescription"],
                  "all": ["urn:schemas:httpmail:subject", "urn:schemas:httpmail:textdescription"]}[where]
        clauses.append(_dasl_like(fields, query))
    if sender:
        clauses.append(_dasl_like(["urn:schemas:httpmail:fromname", "urn:schemas:httpmail:fromemail"], sender))
    if not clauses:
        raise rt.InputError("--query 또는 --from 중 하나는 지정하세요.")
    return "@SQL=" + " AND ".join(clauses)


def cmd_search(args) -> Dict[str, Any]:
    ns = get_namespace(get_outlook())
    folder = resolve_folder(ns, args.folder)
    flt = build_search_filter(args.query, args.sender, args.where)
    items = folder.Items.Restrict(flt)
    items.Sort("[ReceivedTime]", True)
    warnings: List[str] = []
    mails = _collect(items, _check_limit(args.limit, warnings), _literal_keep(args.query, args.sender, args.where))
    desc = " / ".join(x for x in (args.query and f"'{args.query}'", args.sender and f"보낸사람 '{args.sender}'") if x)
    result: Dict[str, Any] = {"folder": str(folder.Name), "mails": mails, "matched": len(mails), "count": len(mails)}
    if warnings:
        result["warnings"] = warnings
    if not mails:      # R2 0건
        return rt.ok(TOOL, "search", f"0건: {desc}: '{folder.Name}'에서 일치하는 메일 없음",
                     f"VERIFIED: 0건. '{folder.Name}' 폴더에 {desc}에 맞는 메일이 없습니다. "
                     "사용자에게 그대로 알리고 검색어나 조건을 바꿔 다시 실행하지 마세요.", **result)
    result["next_commands"] = [rt.next_command("outlook.py", ["read", "--id", mails[0]["id"]], "첫 번째 메일 본문 읽기")]
    return rt.ok(TOOL, "search", f"{desc}: '{folder.Name}'에서 {len(mails)}통", **result)


def _safe_name(text: str, limit: int, default: str = "메일") -> str:
    """파일·폴더 이름에 쓸 수 없는 글자(\\ / : * ? \" < > |, 제어 문자)를 _로, 끝의 점·공백은 뺍니다.
    (RE: 같은 제목이 NTFS 대체 스트림 이름이 되거나 저장에 실패하던 문제)"""
    return rt.safe_filename(text, limit, default)


def _save_attachments(item: Any) -> Tuple[List[str], List[str]]:
    saved: List[str] = []
    skipped: List[str] = []
    subject = _safe_name(_text(getattr(item, "Subject", "메일")), 40)
    folder = os.path.join(rt.get_work_dir(), "mail_attachments", f"{rt.timestamp()}_{subject}")
    for i in range(1, int(item.Attachments.Count) + 1):
        att = item.Attachments.Item(i)
        name = _text(getattr(att, "FileName", "")) or f"첨부{i}"
        try:
            if att.PropertyAccessor.GetProperty(PR_ATTACHMENT_HIDDEN):
                skipped.append(f"{name} (본문 삽입 이미지)")
                continue
        except Exception:
            pass
        if os.path.splitext(name)[1].lower() in BLOCKED_ATTACH:
            skipped.append(f"{name} (실행 파일 형식이라 저장 안 함)")
            continue
        os.makedirs(folder, exist_ok=True)
        base, ext = os.path.splitext(re.sub(r'[\\/:*?"<>|]+', "_", name))
        dest = os.path.join(folder, base + ext)
        k = 2
        while os.path.exists(dest):                       # 같은 이름 첨부는 번호를 붙여 보존
            dest = os.path.join(folder, f"{base} ({k}){ext}")
            k += 1
        try:
            att.SaveAsFile(dest)
        except Exception as e:
            skipped.append(f"{name} ({SAVE_FAILED}: {e})")
            continue
        if os.path.isfile(dest):                          # R1: 저장한 파일이 실제로 생겼는지 다시 확인
            saved.append(dest)
        else:
            skipped.append(f"{name} ({SAVE_FAILED}: 저장 후 파일이 없음)")
    return saved, skipped


def _mail_detail(item: Any, args, source: str) -> Dict[str, Any]:
    info = summary(item)
    info.update(to=_text(getattr(item, "To", "")), cc=_text(getattr(item, "CC", "")), source=source)
    body = _text(getattr(item, "Body", "")).replace("\r\n", "\n").strip()
    clipped, cut = rt.clamp_text(body, rt.DEFAULT_MAX_CHARS)
    result: Dict[str, Any] = {"mail": info, "body": clipped, "body_truncated": cut}
    if getattr(args, "save", None):
        out = rt.resolve_output_path(None if args.save is True else args.save, "mail_" + _safe_name(info["subject"], 30), ".md")
        md = (f"# {info['subject']}\n\n- 보낸사람: {info['from']}\n- 받는사람: {info['to']}\n- 날짜: {info['received']}\n\n{body}\n")
        result["output_file"] = rt.write_text(out, md)
        problem = rt.check_saved_file(result["output_file"])
        if problem is None and info["subject"] not in rt.read_text_file(result["output_file"]):
            problem = "저장한 파일에 메일 제목이 없음"
        if problem:
            result["save_problem"] = problem
        result["next_commands"] = [rt.next_command("word.py", ["write", "--file", result["output_file"]], "메일 내용을 워드에 넣기")]
    if getattr(args, "save_attach", False) and info["attachments"]:
        saved, skipped = _save_attachments(item)
        result["saved_attachments"] = saved
        if skipped:
            result["skipped_attachments"] = skipped
    return result


def cmd_read(args) -> Dict[str, Any]:
    ns = get_namespace(get_outlook())
    try:
        item = ns.GetItemFromID(args.id)
    except Exception:
        raise rt.InputError("해당 id의 메일을 찾을 수 없습니다. list/search 결과의 id를 그대로 쓰세요.")
    result = _mail_detail(item, args, "id")
    return _detail_response("read", f"'{result['mail']['subject']}' ({result['mail']['from']})", result)


def _attachment_problems(result: Dict[str, Any]) -> List[str]:
    """--save-attach 결과 확인 (R1): 저장에 실패한 첨부와, 저장했다는데 다시 보니 없는 파일."""
    failed = [s for s in result.get("skipped_attachments", []) if SAVE_FAILED in s]
    gone = [os.path.basename(p) for p in result.get("saved_attachments", []) if not os.path.isfile(p)]
    problems = []
    if failed:
        problems.append(f"첨부 {len(failed)}개를 저장하지 못함: {', '.join(failed[:5])}")
    if gone:
        problems.append(f"저장한 첨부를 다시 확인하니 없음: {', '.join(gone[:5])}")
    return problems


def _detail_response(command: str, preview: str, result: Dict[str, Any]) -> Dict[str, Any]:
    """read/current 공통 응답: --save 파일이나 --save-attach 첨부를 다시 확인하지 못했으면 UNVERIFIED (R1)."""
    if "saved_attachments" in result:
        preview += f", 첨부 {len(result['saved_attachments'])}개 저장"
    problems = ([result["save_problem"]] if result.get("save_problem") else []) + _attachment_problems(result)
    if problems:
        return rt.ok_unverified(TOOL, command, preview, problems, **result)
    if "output_file" in result:
        return rt.ok(TOOL, command, preview + f", 저장: {result['output_file']}",
                     "VERIFIED: 메일을 읽고 저장한 파일을 다시 열어 확인했습니다.", verified=True, **result)
    if result.get("saved_attachments"):
        return rt.ok(TOOL, command, preview, "VERIFIED: 메일을 읽고 저장한 첨부 파일이 있는지 다시 확인했습니다.",
                     verified=True, **result)
    return rt.ok(TOOL, command, preview, **result)


def current_item(app: Any) -> Tuple[Optional[Any], str]:
    try:
        insp = app.ActiveInspector()
        if insp is not None and insp.CurrentItem is not None:
            return insp.CurrentItem, "열린 창"
    except Exception:
        pass
    try:
        sel = app.ActiveExplorer().Selection
        if sel.Count > 0:
            return sel.Item(1), "목록에서 선택"
    except Exception:
        pass
    return None, ""


def cmd_current(args) -> Dict[str, Any]:
    item, source = current_item(get_outlook())
    if item is None:
        raise rt.InputError("Outlook에서 열려 있거나 선택된 메일이 없습니다. 메일을 선택한 뒤 다시 요청하세요.")
    result = _mail_detail(item, args, source)
    return _detail_response("current", f"{source} 메일: '{result['mail']['subject']}' ({result['mail']['from']})", result)


def _folder_tree(folder: Any, depth: int) -> Dict[str, Any]:
    node: Dict[str, Any] = {"name": str(folder.Name)}
    try:
        node["unread"] = int(folder.UnReadItemCount)
    except Exception:
        pass
    if depth > 0:
        try:
            subs = [_folder_tree(f, depth - 1) for f in folder.Folders]
            if subs:
                node["folders"] = subs
        except Exception:
            pass
    return node


def cmd_folders(args) -> Dict[str, Any]:
    ns = get_namespace(get_outlook())
    stores = [_folder_tree(store, args.depth) for store in ns.Folders]
    return rt.ok(TOOL, "folders", f"계정/저장소 {len(stores)}개의 폴더 구조", accounts=stores)


# ==============================================================================
# 3. 작성
# ==============================================================================

def _html_prefix(text: str) -> str:
    # 줄바꿈 변환은 _body_arg(--body)에서만: 여기서 다시 바꾸면 --body-file의 C:\new\data 같은 경로가 깨짐
    safe = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace("\n", "<br>")
    return ("<div style='font-family: \"Malgun Gothic\", \"Segoe UI\", sans-serif; font-size: 10pt;'>"
            f"{safe}</div><br>")


def _prepend_html(html: str, prefix: str) -> str:
    """<body> 태그 바로 뒤에 넣습니다 (서명과 원본 대화 이력 위)."""
    if not prefix:
        return html
    m = re.search(r"<body[^>]*>", html or "", re.IGNORECASE)
    return html[:m.end()] + prefix + html[m.end():] if m else prefix + (html or "")


def _attach_paths(spec: Optional[str]) -> List[str]:
    """--attach 값을 파일 경로 목록으로. 여러 개는 ; 로 구분 (쉼표도 됨).
    값 전체가 있는 파일이면 하나로 봅니다 (쉼표가 든 파일 이름 '보고서,최종.pdf')."""
    raw = (spec or "").strip().strip('"')
    if not raw:
        return []
    try:
        return [rt.resolve_input_path(raw, "첨부 파일")]
    except rt.InputError:
        if ";" not in raw and "," not in raw:
            raise
    sep = ";" if ";" in raw else ","
    parts = raw.split(sep)
    paths: List[str] = []
    i = 0
    while i < len(parts):
        if not parts[i].strip():
            i += 1
            continue
        # 있는 파일이 아니면 뒤 조각을 구분자로 이어 붙여 봄 (폴더·파일 이름에 ; 나 , 가 든 경로)
        for j in range(i, len(parts)):
            try:
                paths.append(rt.resolve_input_path(sep.join(parts[i:j + 1]).strip().strip('"'), "첨부 파일"))
            except rt.InputError:
                continue
            i = j + 1
            break
        else:
            rt.resolve_input_path(parts[i].strip().strip('"'), "첨부 파일")    # 찾지 못한 조각으로 오류 ([추천] 포함)
    return paths


def _attach(mail: Any, paths: List[str]) -> List[str]:
    names = []
    for path in paths:
        mail.Attachments.Add(path)
        names.append(os.path.basename(path))
    return names


def _finalize(mail: Any, send: bool, command: str, attached: List[str]) -> Dict[str, Any]:
    extra = {"to": _text(getattr(mail, "To", "")), "subject": _text(getattr(mail, "Subject", "")),
             "attached": attached, "sent": False}
    if send:
        unresolved: List[str] = []
        try:
            recips = mail.Recipients
            if recips.Count == 0:
                unresolved.append("(수신자 없음)")
            elif not recips.ResolveAll():
                unresolved = [str(recips.Item(i).Name) for i in range(1, recips.Count + 1) if not recips.Item(i).Resolved]
        except Exception as e:
            unresolved.append(f"(수신자 확인 실패: {e})")
        if unresolved:
            mail.Display()
            extra["unresolved_recipients"] = unresolved
            return {"status": "error", "tool": TOOL, "command": command, "error_type": "input",
                    "message": f"수신자 주소를 확인할 수 없어 보내지 않고 작성 창을 띄웠습니다: {', '.join(unresolved)}",
                    "agent_directive": "STOP_RETRY: 발송하지 않았습니다. 사용자에게 수신자 주소 확인을 요청하세요.", **extra}
        mail.Send()
        extra["sent"] = True
        # R1: 발송은 Outlook이 처리하므로 실제로 나갔는지는 여기서 확인할 수 없음 (보낸 편지함 확인은 사용자 몫)
        return rt.ok_unverified(TOOL, command, f"메일 발송을 Outlook에 맡겼습니다: '{extra['subject']}' -> {extra['to']}",
                                ["실제 발송 여부(보낸 편지함)는 확인하지 않음. 사용자에게 보낸 편지함을 확인하라고 알리세요"], **extra)
    mail.Display()
    return rt.ok(TOOL, command, f"메일 작성 창을 띄웠습니다: '{extra['subject']}' -> {extra['to'] or '(받는 사람 미입력)'}",
                 "VERIFIED: 작성 창의 받는 사람과 제목을 다시 읽었습니다. 발송은 사용자가 직접 [보내기]를 눌러야 합니다. 대신 발송하지 마세요.",
                 verified=True, **extra)


def _load_signature(mail: Any) -> None:
    try:
        _ = mail.GetInspector                # Inspector를 만들면 기본 서명이 본문에 채워짐 (창은 안 뜸)
    except Exception:
        pass


def _target_item(app: Any, args) -> Any:
    if args.id:
        try:
            return get_namespace(app).GetItemFromID(args.id)
        except Exception:
            raise rt.InputError("해당 id의 메일을 찾을 수 없습니다. list/search 결과의 id를 그대로 쓰세요.")
    item, _ = current_item(app)
    if item is None:
        raise rt.InputError("답장/전달할 메일이 없습니다. Outlook에서 메일을 선택하거나 --id를 지정하세요.")
    return item


def _body_arg(args) -> str:
    if args.body_file:
        return rt.read_text_file(rt.resolve_input_path(args.body_file, "본문 파일"))
    return rt.unescape_newlines(args.body or "")      # Windows 경로 안의 \n(C:\new\data)은 그대로


def cmd_new(args) -> Dict[str, Any]:
    # 본문 파일·첨부 파일 확인은 메일을 만들기 전에 (확인에서 멈추면 아무것도 만들지 않음)
    body = _body_arg(args)
    attach = _attach_paths(args.attach)
    app = get_outlook()
    mail = app.CreateItem(0)
    _load_signature(mail)
    if args.to:
        mail.To = args.to
    if args.cc:
        mail.CC = args.cc
    if args.subject:
        mail.Subject = args.subject
    if body:
        mail.HTMLBody = _prepend_html(mail.HTMLBody, _html_prefix(body))   # Body 대입은 서명을 지우므로 HTML로
    return _finalize(mail, args.send, "new", _attach(mail, attach))


def cmd_reply(args) -> Dict[str, Any]:
    body = _body_arg(args)
    attach = _attach_paths(args.attach)
    app = get_outlook()
    original = _target_item(app, args)
    mail = original.ReplyAll() if args.all else original.Reply()
    _load_signature(mail)
    if body:
        mail.HTMLBody = _prepend_html(mail.HTMLBody, _html_prefix(body))
    return _finalize(mail, args.send, "reply", _attach(mail, attach))


def cmd_forward(args) -> Dict[str, Any]:
    body = _body_arg(args)
    attach = _attach_paths(args.attach)
    app = get_outlook()
    original = _target_item(app, args)
    mail = original.Forward()
    _load_signature(mail)
    if args.to:
        mail.To = args.to
    if body:
        mail.HTMLBody = _prepend_html(mail.HTMLBody, _html_prefix(body))
    return _finalize(mail, args.send, "forward", _attach(mail, attach))


# ==============================================================================
# 4. 진입점
# ==============================================================================

def build_cli():
    parser, sub = rt.build_parser(TOOL, "아웃룩 메일")
    folder_help = "폴더: inbox(기본), sent, drafts, deleted, junk 또는 '계정\\폴더\\하위폴더'"

    p = sub.add_parser("list", help="최근 메일 목록")
    p.add_argument("--folder", help=folder_help)
    p.add_argument("--unread", action="store_true", help="안 읽은 메일만")
    p.add_argument("--limit", "-n", type=int, default=10, help=f"최대 몇 통 (기본 10, 최대 {MAX_LIMIT})")

    p = sub.add_parser("search", help="메일 검색")
    p.add_argument("--query", "-q", help="찾을 말")
    p.add_argument("--from", dest="sender", help="보낸사람 이름/주소 일부")
    p.add_argument("--where", choices=["all", "subject", "body"], default="all", help="검색 범위 (기본 제목+본문)")
    p.add_argument("--folder", help=folder_help)
    p.add_argument("--limit", "-n", type=int, default=10, help=f"최대 몇 통 (기본 10, 최대 {MAX_LIMIT})")

    for name, hlp in (("read", "메일 읽기 (--id)"), ("current", "지금 보고 있거나 선택한 메일 읽기")):
        p = sub.add_parser(name, help=hlp)
        if name == "read":
            p.add_argument("--id", required=True, help="list/search 결과의 id")
        p.add_argument("--save-attach", action="store_true", help="첨부를 작업 폴더에 저장 (실행 파일 형식 제외)")
        p.add_argument("--save", nargs="?", const=True, help="본문을 md 파일로 저장")

    p = sub.add_parser("folders", help="폴더 구조")
    p.add_argument("--depth", type=int, default=2, help="하위 폴더 깊이 (기본 2)")

    def compose(p, to=True):
        if to:
            p.add_argument("--to", help="받는 사람 (여러 명은 ; 로 구분)")
        p.add_argument("--body", help="본문 (줄바꿈은 \\n)")
        p.add_argument("--body-file", help="본문 파일 (.txt/.md)")
        p.add_argument("--attach", help="첨부 파일 (여러 개는 ; 로 구분, 쉼표도 됨. 파일명만 쓰면 data 폴더)")
        p.add_argument("--send", action="store_true",
                       help="사용자가 발송을 명시적으로 요청한 경우에만. 수신자 확인 후 발송 (기본: 작성 창만 표시)")

    p = sub.add_parser("new", help="새 메일 작성")
    compose(p)
    p.add_argument("--cc", help="참조")
    p.add_argument("--subject", help="제목")

    p = sub.add_parser("reply", help="답장")
    compose(p, to=False)
    p.add_argument("--id", help="답장할 메일 id (생략 시 지금 보고 있거나 선택한 메일)")
    p.add_argument("--all", action="store_true", help="전체 답장")

    p = sub.add_parser("forward", help="전달")
    compose(p)
    p.add_argument("--id", help="전달할 메일 id (생략 시 지금 보고 있거나 선택한 메일)")
    return parser


HANDLERS = {"list": cmd_list, "search": cmd_search, "read": cmd_read, "current": cmd_current, "folders": cmd_folders,
            "new": cmd_new, "reply": cmd_reply, "forward": cmd_forward}


def main(argv: Optional[Sequence[str]] = None) -> None:
    args = build_cli().parse_args(argv)
    rt.run_cli(TOOL, args.command, lambda: HANDLERS[args.command](args))


if __name__ == "__main__":
    main()
