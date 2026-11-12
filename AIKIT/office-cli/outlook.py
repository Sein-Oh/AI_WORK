"""outlook.py - 아웃룩 메일

조회 (화면은 그대로, 백그라운드):
    list     폴더의 최근 메일 목록 (--unread, --limit, 받은 날짜 --days N / --since YYYY-MM-DD)
    search   제목/본문/보낸사람 검색 (--days / --since 함께 가능)
    read     메일 하나 읽기 (--id: list/search 결과의 짧은 번호 또는 긴 id). 첨부가 있으면 이름과 내려받기 명령을 함께 줌
    thread   같은 대화의 메일을 시간 순으로 (--id)
    calendar 일정 (--days 7: 오늘부터 7일, 읽기만)
    event    일정 만들기·지우기 (참석자 없는 일정만)
    contacts 연락처·주소록에서 메일 주소 찾기 (읽기만)
    current  지금 보고 있거나 선택한 메일 읽기
    folders  폴더 구조와 안 읽은 메일 수
    attachments  첨부 목록, --save 로 내려받기 (기본: %USERPROFILE%\\Downloads\\메일 첨부\\<시각>_<제목>)

작성 (초안을 임시 보관함에 저장하고 작성 창만 띄움. 보내지 않음):
    new      새 메일
    reply    답장 (--all: 전체 답장). 원본 대화 이력 유지. 보내지 않은 초안은 원본으로 쓰지 않음
    forward  전달
    send     new/reply/forward가 저장한 초안(draft_id) 보내기. 결과의 confirm_commands[0]을 사용자가 확인한 뒤에만.
             이 도구가 만든 초안만, 한 번만 나감 (사용자가 쓰던 초안, 이미 보냈거나 지운 초안, 받은 메일은 보내지 않음).
             발송 전 수신자 주소를 확인하고, 확인되지 않으면 보내지 않고 작성 창을 띄웁니다.

예:
    python outlook.py current
    python outlook.py search --query "견적" --from "김팀장" --limit 5
    python outlook.py attachments --id "<id>" --save
    python outlook.py reply --body "확인했습니다."
    python outlook.py new --to "lead@company.com" --subject "주간 보고" --body "첨부합니다." --attach "보고서.pdf"
"""
import datetime as _dt
import json
import os
import re
import sys
import time
from typing import Any, Callable, Dict, List, Optional, Sequence, Tuple

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _office as off  # noqa: E402
import _runtime as rt  # noqa: E402

TOOL = "outlook"
OL_MAIL = 43
OL_FOLDER_DRAFTS = 16      # 임시 보관함 (send는 여기 있는 초안만 보냄)
DEFAULT_FOLDERS = {"inbox": 6, "받은편지함": 6, "sent": 5, "보낸편지함": 5, "drafts": 16, "임시보관함": 16,
                   "deleted": 3, "지운편지함": 3, "outbox": 4, "보낼편지함": 4, "junk": 23, "정크메일": 23}
BLOCKED_ATTACH = {".exe", ".bat", ".cmd", ".com", ".js", ".vbs", ".vbe", ".wsf", ".ps1", ".scr", ".msi", ".lnk",
                  ".hta", ".jar", ".reg", ".dll"}
SAVE_FAILED = "저장 실패"     # attachments --save의 skipped에서 정책상 건너뛴 것과 저장에 실패한 것을 구분하는 표시
READABLE_DOCS = {".pdf", ".hwp", ".hwpx", ".docx", ".docm", ".pptx", ".pptm", ".xlsx", ".xlsm", ".csv", ".txt", ".md"}   # reader.py가 읽는 첨부 (reader.SUPPORTED)
MAX_LIMIT = 200          # list/search --limit 상한 (넘으면 줄이고 warnings로 알림)
MAX_NAMES_SHOW = 20      # read/current의 attachment_names 상한
MAX_ATTACH_SHOW = 50     # attachments 목록·saved·skipped 상한 (넘으면 truncated, 전체는 저장 폴더)
ATTACH_ROOT_NAME = "메일 첨부"   # 기본 저장 위치: <다운로드>\메일 첨부\<시각>_<제목 40자> (자동 정리하지 않음)
DOWNLOADS_GUID = "{374DE290-123F-4565-9164-39C4925E467B}"     # 사용자 셸 폴더 레지스트리의 '다운로드'
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
            hint = rt.elevation_hint()     # 관리자 권한이면 일반 권한으로 띄운 Outlook에 붙지 못함
            raise rt.InputError(f"Outlook에 연결할 수 없습니다 (설치/프로필 확인 필요): {e}" + (f" {hint}" if hint else ""))


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
    given = list(parts)
    stores = {str(f.Name): f for f in ns.Folders}
    found_store = rt.find_name(parts[0], stores)
    alias = re.sub(r"\s+", "", parts[0]).lower()
    aliases = {re.sub(r"\s+", "", k).lower(): k for k in DEFAULT_FOLDERS}
    from_alias = not found_store and alias in aliases
    if found_store:
        cur = stores[found_store]
        parts = parts[1:]
    elif from_alias:
        # 'inbox\프로젝트', '받은편지함\프로젝트': 첫 부분이 기본 폴더 별칭이면 그 폴더에서 시작
        cur = ns.GetDefaultFolder(DEFAULT_FOLDERS[aliases[alias]])
        parts = parts[1:]
    else:
        cur = ns.GetDefaultFolder(6).Parent          # 기본 계정의 최상위
    walked: List[str] = [str(cur.Name)]
    # 최상위(메일 계정 이름, 보통 메일 주소)는 오류 로그에 남기지 않음 (기본 폴더에서 시작했으면 계정 이름이 아님)
    log_root = walked[0] if from_alias else "<메일 계정>"
    for part in parts:
        subs = {str(f.Name): f for f in cur.Folders}
        name = rt.find_name(part, subs)
        if name is None:
            # 비교는 끝 이름끼리 (전체 경로끼리 비교하면 같은 상위 경로 때문에 관계없는 폴더가 모두 [후보]가 됨).
            # 후보는 전체 경로로 보여 줌 (일부만 주면 LLM이 상위 폴더 전체를 검색함)
            prefix = "\\".join(walked)
            # 로그 쪽 경로: 계정 이름으로 시작한 --folder는 첫 부분을 '<메일 계정>'으로 가림 (log_name과 같은 규칙)
            log_spec = "\\".join(["<메일 계정>"] + given[1:]) if found_store else s
            raise rt.not_found("메일 폴더", part, list(subs), extra=f" (경로: {s})",
                               display=lambda n: prefix + "\\" + n,
                               log_name="\\".join([log_root] + walked[1:] + [part]),
                               log_extra=f" (경로: {log_spec})", option="--folder")
        cur = subs[name]
        walked.append(name)
    return cur


def _text(v: Any) -> str:
    return "" if v is None else str(v)


def _field(item: Any, name: str, failed: Optional[List[str]]) -> Any:
    """항목 속성 하나 (읽지 못하면 ''이고 failed에 이름을 남김: Outlook 보안 창에서 '거부'를 누르면 주소 속성이 예외)."""
    try:
        return getattr(item, name, "")
    except Exception:
        if failed is not None:
            failed.append(name)
        return ""


def summary(item: Any, failed: Optional[List[str]] = None) -> Dict[str, Any]:
    sender = _text(_field(item, "SenderName", failed))
    addr = _text(_field(item, "SenderEmailAddress", failed))
    got = _received(item)   # pywin32는 현지 벽시계 시각에 UTC 표시를 붙여 '+00:00'으로 나왔음: 현지 시각 그대로 (2026-11-09 검토 C14)
    received = got.strftime("%Y-%m-%d %H:%M:%S") if got is not None else ""
    try:
        attach = int(item.Attachments.Count)
    except Exception:
        attach = 0
    return {"entry_id": _text(_field(item, "EntryID", failed)), "subject": _text(_field(item, "Subject", failed)),
            "from": f"{sender} <{addr}>" if addr and "@" in addr else sender, "received": received,
            "unread": bool(_field(item, "UnRead", failed)), "attachments": attach}


def _collect(items: Any, limit: int, keep: Optional[Callable[[Any], bool]] = None,
             since: Optional[_dt.datetime] = None, stats: Optional[Dict[str, Any]] = None,
             until: Optional[_dt.datetime] = None) -> List[Dict[str, Any]]:
    """since: 받은 시각이 이보다 이른 메일은 뺍니다 (Restrict 날짜 조건을 다시 확인. 받은 시각 내림차순이라 처음 만나면 멈춤).
    stats: 읽지 못한 항목 수('unreadable')와 일부 칸을 읽지 못한 메일 수('partial', 칸 이름은 'fields'). 조용히 건너뛰면
    모든 항목이 실패해도 'VERIFIED: 0건'이 되어 (2026-11-09 검토 A10) 호출하는 쪽이 UNVERIFIED나 오류로 냅니다."""
    out: List[Dict[str, Any]] = []
    st = stats if stats is not None else {}
    st.setdefault("unreadable", 0)
    st.setdefault("partial", 0)
    st.setdefault("fields", [])
    for item in items:
        if len(out) >= limit:
            break
        try:
            if since is not None or until is not None:
                got = _received(item)
                if since is not None and got is not None and got < since:
                    break
                if until is not None and got is not None and got >= until:      # --until 다음 날 이후 메일 (내림차순이라 앞쪽)
                    continue
            if int(item.Class) == OL_MAIL and (keep is None or keep(item)):
                failed: List[str] = []
                out.append(summary(item, failed))
                if failed:
                    st["partial"] += 1
                    st["fields"] = sorted(set(st["fields"]) | set(failed))
        except Exception:
            st["unreadable"] += 1
            continue
    return out


def collect_problems(stats: Dict[str, Any], what: str) -> List[str]:
    """_collect의 stats -> UNVERIFIED 이유 목록. 하나도 읽지 못했으면 오류 (Outlook 보안 창 거부 등)."""
    out: List[str] = []
    if stats.get("unreadable"):
        out.append(f"{what} {stats['unreadable']}개를 읽지 못함 (빠진 메일이 있을 수 있음, Outlook 보안 확인 창이 떴으면 '허용' 후 다시 실행)")
    if stats.get("partial"):
        out.append(f"메일 {stats['partial']}통의 일부 칸({', '.join(stats['fields'])})을 읽지 못함")
    return out


# ------------------------------------------------------------------------------
# 받은 날짜 범위 (--days / --since)
# ------------------------------------------------------------------------------

def _received(item: Any) -> Optional[_dt.datetime]:
    """메일의 받은 시각 (현지 시각, tzinfo 없음). 읽지 못하면 None (그 메일은 남김: 빠뜨리는 것보다 안전).
    pywin32는 현지 시각에 tzinfo를 붙여 주므로 tzinfo만 떼어 냅니다."""
    try:
        v = item.ReceivedTime
    except Exception:
        return None
    if isinstance(v, _dt.datetime):
        return v.replace(tzinfo=None)
    try:
        return _dt.datetime.fromisoformat(str(v).strip()).replace(tzinfo=None)
    except ValueError:
        return None


# 고쳐서 [추천]할 수 있는 날짜 표기: 2026/11/1, 2026.11.01, 20261101, 2026년 11월 1일
_DATE_PARTS = re.compile(r"^\s*(\d{4})\s*[-./년]?\s*(\d{1,2})\s*[-./월]?\s*(\d{1,2})\s*일?\s*$")


def _today() -> _dt.date:
    return _dt.date.today()


def _parse_day(raw: str, option: str) -> _dt.date:
    """YYYY-MM-DD (2026.11.1, 2026/11/01 같은 표기는 고친 명령을 [추천])."""
    raw = str(raw).strip()
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", raw):
        try:
            return _dt.date.fromisoformat(raw)
        except ValueError:
            pass
    m = _DATE_PARTS.match(raw)
    fixed = None
    if m:
        try:
            fixed = _dt.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            fixed = None
    # 고칠 수 있는 표기만 [추천] (없는 날짜·알 수 없는 글은 도구가 날짜를 지어내지 않음)
    tail = f" [추천] {rt.retry_command({option: fixed.isoformat()})}" if fixed else f" (예: {_today().isoformat()})"
    raise rt.InputError(f"{option} 날짜 '{raw}'을(를) 읽을 수 없습니다. YYYY-MM-DD 형식으로 쓰세요.{tail}")


def date_period(days: Optional[int], since: Optional[str], until: Optional[str] = None) -> Optional[Dict[str, Any]]:
    """--days N(오늘 포함 최근 N일) 또는 --since YYYY-MM-DD, 끝은 --until YYYY-MM-DD(그날 포함, 없으면 오늘)
    -> {'since': 날짜, 'until': 날짜, 'days': 일수}. 셋 다 없으면 None. --until만 주면 그날까지 전부 (since 없음)."""
    if days is None and since is None and until is None:
        return None
    today = _today()
    if until is not None:
        end = _parse_day(until, "--until")
        if days is not None:
            raise rt.InputError("--until은 --since와 함께 씁니다 (--days는 오늘까지의 최근 N일). 기간이면 --since YYYY-MM-DD --until YYYY-MM-DD.")
        if end > today:
            end = today
        if since is None:
            return {"since": None, "until": end.isoformat(), "days": None}
        start = _parse_day(since, "--since")
        if start > end:
            raise rt.InputError(f"--since {start.isoformat()}이(가) --until {end.isoformat()}보다 뒤입니다. 두 날짜를 바꿔 쓰세요. "
                                f"[추천] {rt.retry_command({'--since': end.isoformat(), '--until': start.isoformat()})}")
        return {"since": start.isoformat(), "until": end.isoformat(), "days": (end - start).days + 1}
    if days is not None and since is not None:
        raise rt.InputError("--days와 --since는 함께 쓸 수 없습니다. 최근 N일이면 --days N, 특정 날짜부터면 --since YYYY-MM-DD "
                            "하나만 쓰세요.")
    if days is not None:
        if days < 1:
            rec = 1 if days == 0 else -days
            raise rt.InputError(f"--days는 1 이상이어야 합니다 (1 = 오늘, 7 = 오늘 포함 최근 7일). 받은 값: {days}. [추천] {rt.retry_command({'--days': rec})}")
        start = today - _dt.timedelta(days=days - 1)
    else:
        raw = str(since).strip()
        start = None
        if re.fullmatch(r"\d{4}-\d{2}-\d{2}", raw):
            try:
                start = _dt.date.fromisoformat(raw)
            except ValueError:
                start = None
        if start is None:
            m = _DATE_PARTS.match(raw)
            fixed = None
            if m:
                try:
                    fixed = _dt.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
                except ValueError:
                    fixed = None
            # 고칠 수 있는 표기만 [추천] (없는 날짜·알 수 없는 글은 도구가 날짜를 지어내지 않음)
            tail = f" [추천] {rt.retry_command({'--since': fixed.isoformat()})}" if fixed else f" (예: {today.isoformat()})"
            raise rt.InputError(f"--since 날짜 '{raw}'을(를) 읽을 수 없습니다. YYYY-MM-DD 형식으로 쓰세요.{tail}")
        if start > today:
            raise rt.InputError(f"--since {start.isoformat()}은(는) 오늘({today.isoformat()})보다 뒤입니다. "
                                "오늘이나 그 전 날짜를 쓰세요.")
    return {"since": start.isoformat(), "until": today.isoformat(), "days": (today - start).days + 1}


def _period_start(period: Dict[str, Any]) -> Optional[_dt.datetime]:
    if not period.get("since"):
        return None
    return _dt.datetime.combine(_dt.date.fromisoformat(period["since"]), _dt.time())


def _period_end(period: Optional[Dict[str, Any]]) -> Optional[_dt.datetime]:
    """--until이 오늘보다 앞이면 그다음 날 0시 (그 시각 전에 받은 메일만), 아니면 None (끝 없음)."""
    if not period or period["until"] >= _today().isoformat():
        return None
    return _dt.datetime.combine(_dt.date.fromisoformat(period["until"]), _dt.time()) + _dt.timedelta(days=1)


def _period_text(period: Dict[str, Any]) -> str:
    if not period.get("since"):
        return f"{period['until']}까지"
    return f"{period['since']} ~ {period['until']} ({period['days']}일)"


def jet_date_filter(start: _dt.datetime) -> str:
    """Restrict용 Jet 조건 (현지 시각). 날짜는 YYYY-MM-DD HH:MM: 지역 설정의 월·일 순서와 관계없이 같은 날로 읽히는 형식
    (MM/DD/YYYY는 일/월 순서 지역에서 뒤바뀔 수 있음)."""
    return f"[ReceivedTime] >= '{start.strftime('%Y-%m-%d %H:%M')}'"


def dasl_date_clause(start: _dt.datetime) -> str:
    """@SQL(DASL) 조건: DASL 날짜 비교는 UTC라 현지 자정을 UTC로 바꿔 넣습니다.
    Outlook이 현지 시각으로 읽더라도 한국처럼 UTC보다 빠른 지역에서는 더 이른 시각이 되어 빠지는 메일이 없고,
    더 들어온 메일은 _collect가 받은 시각으로 다시 거릅니다."""
    utc = start.astimezone(_dt.timezone.utc)
    return f"\"urn:schemas:httpmail:datereceived\" >= '{utc.strftime('%Y-%m-%d %H:%M')}'"


def _restrict_dated(items: Any, flt: str, fallback: Optional[str], what: str) -> Any:
    """날짜 조건이 든 Restrict. Outlook이 날짜 조건을 거부하면 날짜 조건 없는 필터(fallback, None이면 거르지 않음)로
    다시 거르고 _collect가 받은 시각을 메일마다 확인합니다 (JIT 복구, recovered)."""
    try:
        return items.Restrict(flt)
    except Exception as e:
        rt.note_recovered(f"Outlook이 {what} 날짜 조건을 거부해({type(e).__name__}) 받은 시각을 메일마다 확인했습니다")
        return items.Restrict(fallback) if fallback else items


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

def _all_unreadable(mails: List[Dict[str, Any]], stats: Dict[str, Any], folder: Any) -> None:
    """항목은 있는데 하나도 읽지 못했으면 오류 (0건으로 답하지 않음)."""
    if not mails and stats.get("unreadable"):
        raise rt.InputError(f"'{folder.Name}' 폴더의 항목 {stats['unreadable']}개를 하나도 읽지 못했습니다. Outlook에 보안 확인 창"
                            "(프로그램이 메일에 접근하려 함)이 떠 있으면 '허용'을 누른 뒤 같은 명령을 1번만 다시 실행하세요. "
                            "창이 없으면 Outlook이 응답하는지 확인하세요.")


def _with_attachments(item: Any) -> bool:
    """본문 삽입 그림(서명 로고 등)이 아닌 첨부가 하나라도 있는지."""
    return any(not a["hidden_inline"] for a in _attachment_infos(item))


def _both(a: Optional[Callable[[Any], bool]], b: Optional[Callable[[Any], bool]]) -> Optional[Callable[[Any], bool]]:
    if a is None or b is None:
        return a or b
    return lambda item: a(item) and b(item)


def cmd_list(args) -> Dict[str, Any]:
    period = date_period(args.days, args.since, args.until)        # 날짜 오류는 Outlook에 연결하기 전에
    start = _period_start(period) if period else None
    end = _period_end(period)
    ns = get_namespace(get_outlook())
    folder = resolve_folder(ns, args.folder)
    items = folder.Items
    if args.unread:
        items = items.Restrict("[UnRead] = True")
    if start is not None:
        items = _restrict_dated(items, jet_date_filter(start), None, "list")
    if end is not None:
        items = _restrict_dated(items, f"[ReceivedTime] < '{end.strftime('%Y-%m-%d %H:%M')}'", None, "list")
    items.Sort("[ReceivedTime]", True)                 # 거른 뒤 정렬해야 순서가 보장됨
    warnings: List[str] = []
    stats: Dict[str, Any] = {}
    mails = _collect(items, _check_limit(args.limit, warnings), _with_attachments if args.with_attachments else None,
                     since=start, stats=stats, until=end)
    _all_unreadable(mails, stats, folder)
    problems = collect_problems(stats, "항목")
    id_warn = remember_listing(mails, "list")          # 짧은 번호 (이어서 늘어남, --id 3)
    compact_listing(mails)
    if id_warn:
        warnings.append(id_warn)
    extra: Dict[str, Any] = {"warnings": warnings} if warnings else {}
    what = ("안 읽은 메일" if args.unread else "메일") if not args.with_attachments else (
        "첨부 있는 안 읽은 메일" if args.unread else "첨부 있는 메일")
    when = ""
    if period:
        extra["period"] = period
        when = f" (받은 날짜 {_period_text(period)})"
    if problems:
        return rt.ok_unverified(TOOL, "list", f"'{folder.Name}' 폴더 메일 {len(mails)}통{when} (일부 항목을 읽지 못함)", problems,
                                folder=str(folder.Name), mails=mails, count=len(mails), **extra)
    if not mails:      # R2 0건
        return rt.ok(TOOL, "list", f"0건: '{folder.Name}' 폴더에 {what}이 없습니다{when}",
                     f"VERIFIED: 0건. '{folder.Name}' 폴더에 {what}이 없습니다{when}. 사용자에게 그대로 알리고 조건을 바꿔 "
                     "다시 실행하지 마세요.", folder=str(folder.Name), mails=[], count=0, **extra)
    return rt.ok(TOOL, "list", f"'{folder.Name}' 폴더 최근 메일 {len(mails)}통" + (f" ({what})" if what != "메일" else "") + when,
                 folder=str(folder.Name), mails=mails, count=len(mails), **extra)


def _dasl_like(fields: Sequence[str], value: str) -> str:
    # 작은따옴표만 겹쳐 씁니다. % _는 그대로 둠: DASL 와일드카드지만 자기 자신과도 맞으므로 빠지는 메일이 없음
    # (_를 공백으로 바꾸거나 %를 지우면 '견적서_v2', '10% 할인'이 절대 맞지 않아 0건 VERIFIED가 됨)
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


PR_SENDER_SMTP = "http://schemas.microsoft.com/mapi/proptag/0x5D01001F"   # PR_SENDER_SMTP_ADDRESS_W


def build_search_filter(query: Optional[str], sender: Optional[str], where: str,
                        since: Optional[_dt.datetime] = None, until: Optional[_dt.datetime] = None) -> str:
    """DASL 필터: @SQL= 접두어는 한 번만, 조건마다 괄호로 묶어 AND로 연결. since: 받은 시각 하한 (현지 시각)."""
    clauses = []
    if query:
        fields = {"subject": ["urn:schemas:httpmail:subject"],
                  "body": ["urn:schemas:httpmail:textdescription"],
                  "all": ["urn:schemas:httpmail:subject", "urn:schemas:httpmail:textdescription"]}[where]
        clauses.append(_dasl_like(fields, query))
    if sender:
        # 사내 Exchange 보낸 사람의 fromemail은 X.500 주소라 SMTP 주소로 못 찾음: 보낸 사람 SMTP 주소 속성도 대조 (2026-11-09 검토 C16, 실기 확인)
        clauses.append(_dasl_like(["urn:schemas:httpmail:fromname", "urn:schemas:httpmail:fromemail", PR_SENDER_SMTP], sender))
    if since is not None:
        clauses.append("(" + dasl_date_clause(since) + ")")
    if until is not None:
        clauses.append("(" + dasl_date_clause(until).replace(" >= ", " < ") + ")")
    if not clauses:
        raise rt.InputError("--query, --from, --days, --since, --until 중 하나는 지정하세요.")
    return "@SQL=" + " AND ".join(clauses)


def cmd_search(args) -> Dict[str, Any]:
    period = date_period(args.days, args.since, args.until)        # 날짜·인자 오류는 Outlook에 연결하기 전에
    start = _period_start(period) if period else None
    end = _period_end(period)
    if not (args.query or args.sender or period) and args.with_attachments:
        # 첨부 조건만: Outlook의 '첨부 있음'으로 거르고, 본문 삽입 그림만 있는 메일은 _with_attachments가 다시 뺌
        flt = '@SQL=("urn:schemas:httpmail:hasattachment" = 1)'
    else:
        flt = build_search_filter(args.query, args.sender, args.where, start, end)
    ns = get_namespace(get_outlook())
    folder = resolve_folder(ns, args.folder)
    if start is not None or end is not None:
        plain = build_search_filter(args.query, args.sender, args.where) if (args.query or args.sender) else None
        items = _restrict_dated(folder.Items, flt, plain, "search")
    else:
        items = folder.Items.Restrict(flt)
    items.Sort("[ReceivedTime]", True)
    warnings: List[str] = []
    stats: Dict[str, Any] = {}
    keep = _both(_literal_keep(args.query, args.sender, args.where), _with_attachments if args.with_attachments else None)
    mails = _collect(items, _check_limit(args.limit, warnings), keep, since=start, stats=stats, until=end)
    _all_unreadable(mails, stats, folder)
    problems = collect_problems(stats, "찾은 항목")
    id_warn = remember_listing(mails, "search")
    compact_listing(mails)
    if id_warn:
        warnings.append(id_warn)
    desc = " / ".join(x for x in (args.query and f"'{args.query}'", args.sender and f"보낸사람 '{args.sender}'",
                                  period and f"받은 날짜 {_period_text(period)}",
                                  args.with_attachments and "첨부 있는 메일") if x)
    result: Dict[str, Any] = {"folder": str(folder.Name), "mails": mails, "matched": len(mails), "count": len(mails)}
    if period:
        result["period"] = period
    if warnings:
        result["warnings"] = warnings
    if problems:
        if mails:
            result["next_commands"] = [rt.next_command("outlook.py", ["read", "--id", str(mails[0]["id"])], "첫 번째 메일 본문 읽기")]
        return rt.ok_unverified(TOOL, "search", f"{desc}: '{folder.Name}'에서 {len(mails)}통 (일부 항목을 읽지 못함)", problems, **result)
    if not mails:      # R2 0건
        return rt.ok(TOOL, "search", f"0건: {desc}: '{folder.Name}'에서 일치하는 메일 없음",
                     f"VERIFIED: 0건. '{folder.Name}' 폴더에 {desc}에 맞는 메일이 없습니다. "
                     "사용자에게 그대로 알리고 검색어나 조건을 바꿔 다시 실행하지 마세요.", **result)
    result["next_commands"] = [rt.next_command("outlook.py", ["read", "--id", str(mails[0]["id"])], "첫 번째 메일 본문 읽기")]
    return rt.ok(TOOL, "search", f"{desc}: '{folder.Name}'에서 {len(mails)}통", **result)


def _safe_name(text: str, limit: int, default: str = "메일") -> str:
    """파일·폴더 이름에 쓸 수 없는 글자(\\ / : * ? \" < > |, 제어 문자)를 _로, 끝의 점·공백은 뺍니다.
    (RE: 같은 제목이 NTFS 대체 스트림 이름이 되거나 저장에 실패하던 문제)"""
    return rt.safe_filename(text, limit, default)


def _blocked_name(name: str) -> bool:
    """실행 파일 형식 첨부인지. 이름 끝의 공백·점을 지운 뒤 확장자로 봄 (Windows는 저장할 때 끝의 공백·점을 지워
    'evil.exe '와 'a.js.'가 실행 파일이 되는데 검사를 통과하던 것, 2026-11-09 검토 B13)."""
    clean = name.rstrip(" .\t")
    return os.path.splitext(clean)[1].lower() in BLOCKED_ATTACH


def _attachment_infos(item: Any) -> List[Dict[str, Any]]:
    """메일의 첨부 목록: index(1부터), name, size(Outlook이 알려 주면 바이트), hidden_inline(본문 삽입 이미지), blocked(실행 파일 형식)."""
    infos: List[Dict[str, Any]] = []
    try:
        atts = item.Attachments
        total = int(atts.Count)
    except Exception:
        return infos
    for i in range(1, total + 1):
        try:
            att = atts.Item(i)
        except Exception:
            continue
        name = _text(getattr(att, "FileName", "")) or _text(getattr(att, "DisplayName", "")) or f"첨부{i}"
        try:
            size: Optional[int] = int(att.Size)
        except Exception:
            size = None
        try:
            hidden = bool(att.PropertyAccessor.GetProperty(PR_ATTACHMENT_HIDDEN))
        except Exception:
            hidden = False
        infos.append({"index": i, "name": name, "size": size, "hidden_inline": hidden,
                      "blocked": _blocked_name(name), "_att": att})
    return infos


def _public(info: Dict[str, Any]) -> Dict[str, Any]:
    return {k: v for k, v in info.items() if not k.startswith("_")}


def _mail_id_args(item: Any) -> List[str]:
    """next_commands에 넣을 --id (보고 있는 메일도 id로 고정해 다음 명령이 다른 메일을 잡지 않게)."""
    mid = _text(getattr(item, "EntryID", ""))
    return ["--id", mid] if mid else []


def _mail_detail(item: Any, args, source: str) -> Dict[str, Any]:
    info = summary(item)
    info.update(to=_text(getattr(item, "To", "")), cc=_text(getattr(item, "CC", "")), source=source)
    body = _text(getattr(item, "Body", "")).replace("\r\n", "\n").strip()
    clipped, cut = rt.clamp_text(body, rt.DEFAULT_MAX_CHARS)
    result: Dict[str, Any] = {"mail": info, "body": clipped, "body_truncated": cut}
    nexts: List[Dict[str, str]] = []
    if getattr(args, "save", None):
        out = rt.resolve_output_path(None if args.save is True else args.save, "mail_" + _safe_name(info["subject"], 30), ".md")
        md = (f"# {info['subject']}\n\n- 보낸사람: {info['from']}\n- 받는사람: {info['to']}\n- 날짜: {info['received']}\n\n{body}\n")
        result["output_file"] = rt.write_text(out, md)
        problem = rt.check_saved_file(result["output_file"])
        if problem is None and info["subject"] not in rt.read_text_file(result["output_file"]):
            problem = "저장한 파일에 메일 제목이 없음"
        if problem:
            result["save_problem"] = problem
        # 워드에 넣기는 열린 문서를 바꿈: confirm_commands (1.4.75)
        result["confirm_commands"] = [rt.next_command("word.py", ["write", "--file", result["output_file"]],
                                                      "메일 내용을 열린 워드 문서에 넣기 (사용자가 원할 때만)")]
    if info["attachments"]:
        atts = _attachment_infos(item)
        names = [a["name"] for a in atts if not a["hidden_inline"]]
        inline = len(atts) - len(names)
        if inline:
            result["inline_images"] = inline
        if names:
            result["attachment_names"] = names[:MAX_NAMES_SHOW]
            result["attachment_names_truncated"] = len(names) > MAX_NAMES_SHOW
            # 내려받기는 사용자가 원할 때만 (1.4.83: next_commands는 묻지 않고 실행하는 명령이라 confirm_commands로)
            result.setdefault("confirm_commands", []).append(rt.next_command(
                "outlook.py", ["attachments"] + _mail_id_args(item) + ["--save"], "첨부 내려받기 (사용자가 원할 때만)"))
    if nexts:
        result["next_commands"] = nexts
    return result


def cmd_read(args) -> Dict[str, Any]:
    item = get_item(get_namespace(get_outlook()), args.id)
    result = _mail_detail(item, args, "id")
    return _detail_response("read", f"'{result['mail']['subject']}' ({result['mail']['from']})", result)


def _detail_response(command: str, preview: str, result: Dict[str, Any]) -> Dict[str, Any]:
    """read/current 공통 응답: --save 파일을 다시 확인하지 못했으면 UNVERIFIED (R1)."""
    if result.get("attachment_names"):
        n = len(result["attachment_names"])
        preview += f", 첨부 {n}개" + (" 이상" if result["attachment_names_truncated"] else "")
    if result.get("save_problem"):
        return rt.ok_unverified(TOOL, command, preview, [result["save_problem"]], **result)
    if "output_file" in result:
        return rt.ok(TOOL, command, preview + f", 저장: {result['output_file']}",
                     "VERIFIED: 메일을 읽고 저장한 파일을 다시 열어 확인했습니다. " + rt.INTO_OTHER_DOC, verified=True, **result)
    return rt.ok(TOOL, command, preview, **result)


# ------------------------------------------------------------------------------
# 첨부 내려받기 (attachments)
# ------------------------------------------------------------------------------

def _shell_downloads() -> Optional[str]:
    """Windows가 아는 '다운로드' 폴더 (사용자가 D:\\ 등으로 옮겼을 수 있음). 읽지 못하면 None."""
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                            r"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders") as key:
            value = str(winreg.QueryValueEx(key, DOWNLOADS_GUID)[0])
        path = os.path.expandvars(value)
        return path if os.path.isabs(path) else None
    except Exception:
        return None


def downloads_dir() -> str:
    """첨부 기본 저장 위치의 뿌리: 사용자의 다운로드 폴더.
    도구 data 폴더(7일 뒤 정리되는 _work, En-Genius 고정 폴더)와 관계없음."""
    home = os.environ.get("USERPROFILE") or os.path.expanduser("~")
    return _shell_downloads() or os.path.join(home, "Downloads")


def _save_folder(spec: Any, subject: str) -> str:
    if spec is True:
        return os.path.join(downloads_dir(), ATTACH_ROOT_NAME, f"{rt.timestamp()}_{_safe_name(subject, 40)}")
    raw = os.path.expandvars(os.path.expanduser(str(spec).strip().strip('"')))
    if not raw:
        raise rt.InputError("--save 폴더가 비어 있습니다. 폴더를 빼고 --save만 쓰면 다운로드 폴더에 저장합니다.")
    if not os.path.isabs(raw):
        raise rt.InputError(f"--save 폴더는 전체 경로로 지정하세요 (받은 값: {raw}). 사용자에게 전체 경로를 물어보세요. "
                            "폴더를 빼고 --save만 쓰면 다운로드 폴더에 저장합니다.")
    raw = os.path.abspath(raw)
    if os.path.isfile(raw):
        raise rt.InputError(f"--save에는 폴더를 지정하세요. 같은 이름의 파일이 있습니다: {raw}")
    return raw


def _select(infos: List[Dict[str, Any]], names: Optional[Sequence[str]],
            indexes: Optional[Sequence[str]]) -> Optional[List[Dict[str, Any]]]:
    """--name(이름 일부, 대소문자 무시, ; 로 여러 개)과 --index(1,3 또는 2-4)로 고른 첨부. 둘 다 없으면 None (전부).
    맞는 첨부가 없는 값이 하나라도 있으면 [후보](첨부 이름)로 멈춥니다 (추측해서 다른 첨부를 저장하지 않음)."""
    parts = [p.strip() for n in (names or []) for p in str(n).split(";") if p.strip()]
    idx = [p for spec in (indexes or []) for p in re.split(r"[,;\s]+", str(spec)) if p]
    if not parts and not idx:
        if names or indexes:
            raise rt.InputError("--name/--index 값이 비어 있습니다. 첨부 이름 일부나 번호(1,3)를 주세요.")
        return None
    chosen: Dict[int, Dict[str, Any]] = {}
    cands = [a["name"] for a in infos]
    if not infos:
        raise rt.InputError("이 메일에는 첨부가 없어 --name/--index로 고를 것이 없습니다.")
    for part in parts:
        hits = [a for a in infos if part.casefold() in a["name"].casefold()]
        if not hits:
            raise rt.ambiguous(f"--name '{part}'에 맞는 첨부가 없습니다. 사용자에게 첨부를 고르게 하세요.", cands,
                               log_message="--name 값에 맞는 첨부가 없습니다.", replace=part, option="--name")
        chosen.update((a["index"], a) for a in hits)
    for token in idx:
        m = re.fullmatch(r"(\d+)(?:-(\d+))?", token)
        if not m:
            raise rt.InputError(f"--index는 첨부 번호(1부터)를 쉼표로 구분해 주세요 (예: 1,3 또는 2-4). 받은 값: {token}")
        lo, hi = int(m.group(1)), int(m.group(2) or m.group(1))
        hits = [a for a in infos if lo <= a["index"] <= hi]
        if not hits:
            raise rt.ambiguous(f"--index {token}에 해당하는 첨부가 없습니다 (첨부 {len(infos)}개, 번호는 1부터). "
                               "사용자에게 첨부를 고르게 하세요.", cands,
                               log_message=f"--index {token}에 해당하는 첨부가 없습니다 (첨부 {len(infos)}개).",
                               replace=token, values=[str(a["index"]) for a in infos], option="--index")
        chosen.update((a["index"], a) for a in hits)
    return [chosen[k] for k in sorted(chosen)]


def _unique_dest(folder: str, name: str, index: int) -> str:
    base, ext = os.path.splitext(name)
    base = rt.safe_filename(base, 120, f"첨부{index}")
    ext = rt.safe_filename(ext, 20, "") if ext else ""
    dest = os.path.join(folder, base + ext)
    k = 2
    while os.path.exists(dest):                    # 덮어쓰지 않음: ' (2)'를 붙이고 recovered로 알림 (R4)
        dest = os.path.join(folder, f"{base} ({k}){ext}")
        k += 1
    if k > 2:
        rt.note_recovered(f"{base + ext} 이미 있음 -> 덮어쓰지 않고 {os.path.basename(dest)}로 저장")
    return dest


def _save_selected(infos: List[Dict[str, Any]], picked: Optional[List[Dict[str, Any]]], folder: str
                   ) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]], List[str]]:
    """(저장한 첨부, 건너뛴 첨부, 확인하지 못한 내용).
    실행 파일 형식은 늘 건너뛰고, 본문 삽입 이미지는 --name/--index로 직접 고른 경우에만 저장합니다."""
    saved: List[Dict[str, Any]] = []
    skipped: List[Dict[str, Any]] = []
    problems: List[str] = []
    for a in (infos if picked is None else picked):
        if a["blocked"]:
            skipped.append({"name": a["name"], "reason": "실행 파일 형식이라 저장 안 함"})
            continue
        if a["hidden_inline"] and picked is None:
            skipped.append({"name": a["name"], "reason": "본문 삽입 이미지 (저장하려면 --name이나 --index로 고름)"})
            continue
        rt.retry_fs(lambda: os.makedirs(folder, exist_ok=True))
        dest = _unique_dest(folder, a["name"], a["index"])
        try:
            a["_att"].SaveAsFile(dest)
        except Exception as e:
            skipped.append({"name": a["name"], "reason": f"{SAVE_FAILED}: {str(e)[:120]}"})
            problems.append(f"{a['name']} {SAVE_FAILED}")
            continue
        # R1: 저장한 파일이 실제로 있는지, Outlook이 크기를 알려 준 첨부는 빈 파일이 아닌지 다시 확인.
        # (Outlook의 Size에는 메일 속성 몫이 더해져 파일 크기보다 조금 큼. 같은 크기를 요구하면 정상 저장도 실패로 봄)
        try:
            got = os.path.getsize(dest)
        except OSError:
            problems.append(f"{a['name']}: 저장 후 다시 보니 파일이 없음 ({dest})")
            continue
        if a["size"] and got == 0:
            problems.append(f"{a['name']}: 저장한 파일이 비어 있음 (Outlook 크기 {a['size']}바이트)")
        saved.append({"name": a["name"], "path": dest, "size": got})
    return saved, skipped, problems


def cmd_attachments(args) -> Dict[str, Any]:
    listing = args.save is None
    if listing and (args.name or args.index):
        raise rt.InputError("--name/--index는 --save와 함께 씁니다. 목록만 보려면 둘 다 빼세요. "
                            f"[추천] {rt.retry_command({'--save': True})}")
    item = _target_item(get_outlook(), args, "첨부를 볼 메일이 없습니다. Outlook에서 메일을 선택하거나 --id를 지정하세요.",
                        original=False)
    subject = _text(getattr(item, "Subject", ""))
    infos = _attachment_infos(item)
    base: Dict[str, Any] = {"mail": {"entry_id": _text(getattr(item, "EntryID", "")), **_original_info(item)}}
    if listing:
        shown = infos[:MAX_ATTACH_SHOW]
        base.update(attachments=[_public(a) for a in shown], count=len(infos), truncated=len(infos) > len(shown))
        if not infos:      # R2 0건
            return rt.ok(TOOL, "attachments", f"0건: '{subject}' 메일에 첨부 없음",
                         "VERIFIED: 0건. 이 메일에는 첨부 없음. 사용자에게 그대로 알리고 다시 실행하지 마세요.", **base)
        # 내려받기는 사용자가 고른 것만 (1.4.83): [0] 모두, [1..] 첨부 하나씩 (번호 = attachments의 index, 값을 끼워 넣지 않음)
        base["confirm_commands"] = [rt.next_command("outlook.py", ["attachments"] + _mail_id_args(item) + ["--save"],
                                                    "첨부 모두 내려받기 (다운로드 폴더)")] + [
            rt.next_command("outlook.py", ["attachments"] + _mail_id_args(item) + ["--save", "--index", str(a["index"])],
                            f"'{a['name']}'만 내려받기") for a in shown]
        inline = sum(1 for a in infos if a["hidden_inline"])
        return rt.ok(TOOL, "attachments", f"'{subject}' 메일 첨부 {len(infos)}개"
                     + (f" (본문 삽입 이미지 {inline}개 포함)" if inline else ""),
                     "VERIFIED: 첨부 목록을 읽었습니다. 사용자에게 이름을 알리고, 내려받기를 원하면 confirm_commands에서 사용자가 "
                     "고른 것(모두는 confirm_commands[0], 하나만이면 purpose에 그 이름이 있는 명령)을 그대로 1번 실행하세요.", **base)
    folder = _save_folder(args.save, subject)
    picked = _select(infos, args.name, args.index)
    saved, skipped, problems = _save_selected(infos, picked, folder)
    base.update(saved=[s["path"] for s in saved][:MAX_ATTACH_SHOW], count=len(saved),
                folder=folder if saved else None, skipped=skipped[:MAX_ATTACH_SHOW],
                truncated=len(saved) > MAX_ATTACH_SHOW or len(skipped) > MAX_ATTACH_SHOW)
    names = ", ".join(s["name"] for s in saved[:10]) + (" ..." if len(saved) > 10 else "")
    if problems:
        return rt.ok_unverified(TOOL, "attachments", f"첨부 {len(saved)}개 저장: {folder}", problems, **base)
    if not saved:      # R2 0건
        why = "첨부 없음" if not infos else "저장할 첨부 없음 (건너뛴 첨부: " + ", ".join(
            f"{s['name']} - {s['reason']}" for s in skipped[:5]) + ")"
        return rt.ok(TOOL, "attachments", f"0건: '{subject}' 메일에 {why}",
                     f"VERIFIED: 0건. {why}. 사용자에게 그대로 알리고 다시 실행하지 마세요.", **base)
    tail = f" 건너뛴 첨부 {len(skipped)}개는 skipped의 이유와 함께 알리세요." if skipped else ""
    docs = [x["path"] for x in saved if os.path.splitext(x["path"])[1].lower() in READABLE_DOCS]
    if docs:
        # 저장한 문서를 이어서 읽거나 등록하는 명령 (경량 LLM이 경로를 조합하지 않게)
        nexts = [rt.next_command("reader.py", ["read", "--file", d], f"저장한 문서 읽기: {os.path.basename(d)}") for d in docs[:3]]
        if len(docs) > 1:
            nexts.append(rt.next_command("reader.py", ["index", "--folder", folder], f"저장한 문서 {len(docs)}개를 등록해 검색할 수 있게"))
        base["next_commands"] = nexts
        tail += " 문서 내용을 보려면 next_commands의 reader.py 명령을 쓰세요."
    return rt.ok(TOOL, "attachments", f"'{subject}' 메일({base['mail']['from']})의 첨부 {len(saved)}개 저장: {folder} ({names})",
                 f"VERIFIED: 첨부 {len(saved)}개를 저장하고 파일이 있는지 다시 확인했습니다. "
                 f"사용자에게 저장 폴더(folder)와 파일 이름을 그대로 알리세요.{tail}", verified=True, **base)


OL_CLASS_EXPLORER, OL_CLASS_INSPECTOR = 34, 35


def current_item(app: Any) -> Tuple[Optional[Any], str]:
    """사용자가 지금 보는 메일. 맨 앞 창(ActiveWindow)이 목록(Explorer)이면 선택한 메일, 메일 창(Inspector)이면 그 메일.
    ActiveInspector()부터 보면 뒤에 열어 둔 메일 창을 골랐음 (2026-11-09 검토 C15, 실기 확인)."""
    try:
        win = app.ActiveWindow()
        cls = int(win.Class)
        if cls == OL_CLASS_EXPLORER:
            sel = win.Selection
            if sel.Count > 0:
                return sel.Item(1), "목록에서 선택"
        elif cls == OL_CLASS_INSPECTOR and win.CurrentItem is not None:
            return win.CurrentItem, "열린 창"
    except Exception:
        pass
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


FOLDERS_MAX = 150   # folders 결과에 담는 폴더 수 상한 (상한 없이 수백 개가 나오던 것, 2026-11-09 검토 5장 4번)
DEPTH_MAX = 5


def _folder_tree(folder: Any, depth: int, budget: Dict[str, int]) -> Dict[str, Any]:
    node: Dict[str, Any] = {"name": str(folder.Name)}
    budget["n"] += 1
    try:
        node["unread"] = int(folder.UnReadItemCount)
    except Exception:
        pass
    if depth > 0:
        try:
            subs = []
            for f in folder.Folders:
                if budget["n"] >= FOLDERS_MAX:
                    budget["cut"] = 1
                    break
                subs.append(_folder_tree(f, depth - 1, budget))
            if subs:
                node["folders"] = subs
        except Exception:
            pass
    return node


def cmd_folders(args) -> Dict[str, Any]:
    ns = get_namespace(get_outlook())
    depth = max(0, min(args.depth, DEPTH_MAX))
    budget = {"n": 0, "cut": 0}
    stores = []
    for store in ns.Folders:
        if budget["n"] >= FOLDERS_MAX:
            budget["cut"] = 1
            break
        stores.append(_folder_tree(store, depth, budget))
    extra: Dict[str, Any] = {"folder_count": budget["n"], "truncated": bool(budget["cut"])}
    warnings = [f"--depth {args.depth}은 {DEPTH_MAX}로 줄임"] if args.depth > DEPTH_MAX else []
    if budget["cut"]:
        warnings.append(f"폴더가 많아 {FOLDERS_MAX}개까지만 담음: 찾는 폴더가 없으면 --depth를 줄이거나 search·list의 --folder에 상위 폴더 경로를 주세요")
    if warnings:
        extra["warnings"] = warnings
    return rt.ok(TOOL, "folders", f"계정/저장소 {len(stores)}개의 폴더 구조 (폴더 {budget['n']}개" + (", 일부만" if budget["cut"] else "") + ")",
                 "VERIFIED: 폴더 구조를 읽었습니다." + (" 일부만 담았습니다(truncated, warnings)." if budget["cut"] else ""),
                 accounts=stores, **extra)


CONTACT_SCAN_MAX = 5000      # 연락처 폴더에서 훑는 항목 수 상한
CONTACT_SHOW = 20            # JSON에 담는 결과 수
OL_FOLDER_CONTACTS = 10


def _exchange_smtp(entry: Any) -> str:
    """주소록 항목의 메일 주소: Exchange 사용자면 PrimarySmtpAddress, 아니면 Address."""
    for get in ("GetExchangeUser", "GetExchangeDistributionList"):
        try:
            u = getattr(entry, get)()
            if u is not None and _text(getattr(u, "PrimarySmtpAddress", "")):
                return _text(u.PrimarySmtpAddress)
        except Exception:
            pass
    return _text(getattr(entry, "Address", ""))


def cmd_contacts(args) -> Dict[str, Any]:
    """이름(일부)이나 메일 주소 일부로 내 연락처와 주소록(Outlook 이름 확인)을 찾음 (읽기만, 1.6.7 기반 기능).
    받는 사람을 이름만 말했을 때 메일 주소로 바꾸는 데 씀. 여럿이면 모두 보여 주고 고르는 것은 사용자."""
    key = str(args.name or "").strip()
    if not key:
        raise rt.InputError("찾을 이름을 --name으로 주세요 (예: --name 홍길동).")
    ns = get_namespace(get_outlook())
    fold = key.casefold()
    found: List[Dict[str, Any]] = []
    seen: set = set()
    scanned = 0
    try:
        items = ns.GetDefaultFolder(OL_FOLDER_CONTACTS).Items
        for it in items:
            scanned += 1
            if scanned > CONTACT_SCAN_MAX:
                break
            name = _text(getattr(it, "FullName", ""))
            mail = _text(getattr(it, "Email1Address", ""))
            company = _text(getattr(it, "CompanyName", ""))
            if fold in name.casefold() or (mail and fold in mail.casefold()):
                k = (mail or name).casefold()
                if k not in seen:
                    seen.add(k)
                    found.append({"name": name, "email": mail, "company": company, "source": "연락처"})
    except Exception as e:
        rt.log("WARN", f"연락처 폴더를 읽지 못함 ({type(e).__name__})")
    try:
        rcp = ns.CreateRecipient(key)
        if rcp.Resolve():
            entry = rcp.AddressEntry
            mail = _exchange_smtp(entry)
            k = (mail or _text(entry.Name)).casefold()
            if k not in seen:
                seen.add(k)
                found.append({"name": _text(entry.Name), "email": mail, "company": "", "source": "주소록"})
    except Exception as e:
        rt.log("WARN", f"주소록 이름 확인 실패 ({type(e).__name__})")
    fields: Dict[str, Any] = {"query": key, "contacts": found[:CONTACT_SHOW], "count": len(found),
                              "truncated": len(found) > CONTACT_SHOW}
    warnings = [f"연락처가 많아 앞 {CONTACT_SCAN_MAX}개까지만 봄"] if scanned > CONTACT_SCAN_MAX else []
    if warnings:
        fields["warnings"] = warnings
    if not found:
        return rt.ok(TOOL, "contacts", f"0건: '{key}'와(과) 맞는 연락처·주소록 항목이 없습니다",
                     f"VERIFIED: 0건. '{key}'을(를) 찾지 못했습니다 (찾은 항목 없음). 사용자에게 메일 주소를 물으세요.", **fields)
    shown = ", ".join(f"{c['name']} <{c['email']}>" for c in found[:5])
    directive = ("VERIFIED: 하나를 찾았습니다." if len(found) == 1 else
                 f"VERIFIED: {len(found)}명을 찾았습니다. 누구인지 정하지 말고 목록을 보여 주고 사용자가 고르게 하세요.")
    return rt.ok(TOOL, "contacts", f"'{key}': {shown}" + (" 외" if len(found) > 5 else ""), directive, **fields)


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


def _attach_file(spec: str) -> str:
    """--attach 조각 하나를 파일 경로로. 전체 경로가 아니면 data 폴더에서 찾습니다(파일 이름만 쓴 경우).
    전체 경로인데 없으면 같은 폴더의 비슷한 파일을 [후보](전체 경로)로 멈춥니다. 메일은 아직 만들지 않은 때입니다."""
    raw = os.path.expandvars(os.path.expanduser(str(spec).strip().strip('"')))
    if not os.path.isabs(raw):
        return rt.resolve_input_path(raw, "첨부 파일")
    if os.path.isfile(raw):
        return os.path.abspath(raw)
    if os.path.isdir(raw):
        raise rt.InputError(f"--attach에는 파일을 지정하세요. 폴더는 첨부할 수 없습니다: {raw}")
    folder, name = os.path.split(os.path.abspath(raw))
    if not os.path.isdir(folder):
        raise rt.InputError(f"첨부 파일의 폴더가 없습니다: {folder}. 아무것도 만들지 않았습니다. "
                            "사용자에게 파일의 전체 경로를 확인하세요.")
    try:
        files = [n for n in os.listdir(folder) if os.path.isfile(os.path.join(folder, n))]
    except OSError:
        files = []
    close = rt.suggest(name, files, 5)
    if close:
        raise rt.ambiguous(f"첨부 파일 '{name}'이(가) 없습니다 (폴더: {folder}). 아무것도 만들지 않았습니다. "
                           "같은 폴더의 비슷한 파일 중에서 사용자에게 고르게 하세요.", [os.path.join(folder, c) for c in close],
                           replace=raw)
    raise rt.InputError(f"첨부 파일이 없습니다: {raw}. 아무것도 만들지 않았습니다. 사용자에게 파일의 전체 경로를 확인하세요.")


def _attach_paths(spec: Optional[str]) -> List[str]:
    """--attach 값을 파일 경로 목록으로. 여러 개는 ; 로 구분 (쉼표도 됨).
    값 전체가 있는 파일이면 하나로 봅니다 (쉼표가 든 파일 이름 '보고서,최종.pdf')."""
    raw = (spec or "").strip().strip('"')
    if not raw:
        return []
    try:
        return [_attach_file(raw)]
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
                paths.append(_attach_file(sep.join(parts[i:j + 1])))
            except rt.InputError:
                continue
            i = j + 1
            break
        else:
            _attach_file(parts[i])    # 찾지 못한 조각으로 오류 ([추천]/[후보] 포함)
    return paths


def _attach(mail: Any, paths: List[str]) -> Tuple[List[str], List[str]]:
    """파일을 붙이고 (붙인 이름, 확인하지 못한 내용). R1: Attachments.Count가 붙인 수만큼 늘었는지 다시 셉니다
    (전달 메일은 원본 첨부가 이미 있으므로 늘어난 수로 비교)."""
    names: List[str] = []
    problems: List[str] = []
    if not paths:
        return names, problems
    try:
        before: Optional[int] = int(mail.Attachments.Count)
    except Exception:
        before = None
    for path in paths:
        try:
            mail.Attachments.Add(path)
            names.append(os.path.basename(path))
        except Exception as e:
            problems.append(f"{os.path.basename(path)} 첨부 실패 ({str(e)[:120]})")
    try:
        after: Optional[int] = int(mail.Attachments.Count)
    except Exception:
        after = None
    if before is None or after is None:
        problems.append("첨부 개수를 다시 읽지 못함")
    elif after - before != len(names):
        problems.append(f"첨부 {len(names)}개를 넣었는데 메일의 첨부는 {after - before}개 늘어남")
    return names, problems


def _original_info(item: Any) -> Dict[str, Any]:
    """답장·전달·첨부 대상으로 고른 메일 (사용자에게 어느 메일인지 알리게 결과에 넣음)."""
    s = summary(item)
    return {"subject": s["subject"], "from": s["from"], "received": s["received"]}


def _original_note(original: Optional[Dict[str, Any]]) -> str:
    if not original:
        return ""
    return f" (원본: '{original['subject']}', 보낸사람 {original['from'] or '(없음)'}, {original['received']})"


def _unresolved(mail: Any) -> List[str]:
    """보내기 전 수신자 확인: 확인하지 못한 수신자 이름 (없으면 빈 목록)."""
    try:
        recips = mail.Recipients
        if recips.Count == 0:
            return ["(수신자 없음)"]
        if not recips.ResolveAll():
            return [str(recips.Item(i).Name) for i in range(1, recips.Count + 1) if not recips.Item(i).Resolved]
    except Exception as e:
        return [f"(수신자 확인 실패: {e})"]
    return []


def _save_draft(mail: Any) -> str:
    """작성한 메일을 임시 보관함에 저장하고 EntryID를 돌려줍니다 (저장하지 못하면 빈 글자)."""
    try:
        mail.Save()
        return _text(getattr(mail, "EntryID", ""))
    except Exception as e:
        rt.log("WARN", f"초안을 임시 보관함에 저장하지 못함: {type(e).__name__}")
        return ""


def _draft_id(eid: str, subject: str, warnings: List[str]) -> str:
    """초안에 짧은 번호를 붙임 (붙이지 못하면 EntryID 그대로)."""
    m: Dict[str, Any] = {"entry_id": eid, "subject": subject, "from": "", "received": ""}
    warn = remember_listing([m], "draft")
    if warn:
        warnings.append(warn)
    return str(m["id"])


def _recipients(mail: Any) -> Dict[str, str]:
    """받는 사람(to)과, 있으면 참조(cc)·숨은 참조(bcc). 발송 전 확인 화면에 참조가 빠지면 사용자가 모르는 사람에게 나감 (2026-11-09 검토 A11)."""
    out = {"to": _text(_field(mail, "To", None))}
    for key, prop in (("cc", "CC"), ("bcc", "BCC")):
        v = _text(_field(mail, prop, None))
        if v.strip():
            out[key] = v
    return out


def _recipients_text(r: Dict[str, str]) -> str:
    text = r.get("to") or "(받는 사람 미입력)"
    if r.get("cc"):
        text += f", 참조 {r['cc']}"
    if r.get("bcc"):
        text += f", 숨은 참조 {r['bcc']}"
    return text


def _finalize(mail: Any, command: str, attach: Tuple[List[str], List[str]],
              original: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    """작성한 메일을 초안으로 저장하고 작성 창을 띄움. 보내지 않고, 이 초안(같은 항목)을 보내는 send 명령을 confirm_commands에
    (사용자 확인 뒤에만, 1.4.73). 초안은 발송 원장(DRAFTS_FILE)에 '이 도구가 만든 초안'으로 기록.
    new/reply/forward에 --send가 없음 (1.4.72): 초안 창이 열린 채 같은 메일을 또 만들어 보내거나(중복 발송),
    --id 없는 두 번째 명령이 열린 초안을 원본으로 잡아 초안에 답장·전달하는 일(엉뚱한 대상)이 구조적으로 생기지 않게."""
    attached, attach_problems = attach
    extra: Dict[str, Any] = {**_recipients(mail), "subject": _text(getattr(mail, "Subject", "")),
                             "attached": attached, "sent": False}
    note = _original_note(original)
    if original is not None:
        extra["original"] = original
    eid = _save_draft(mail)
    try:
        mail.Display()
    except Exception as e:
        # Outlook에 대화 상자가 떠 있으면 작성 창을 못 띄움. 초안은 이미 저장됐으므로 일반 오류로 끝내면 다시 실행할 때마다
        # 초안이 늘어남 (2026-11-09 검토 C17): 초안 정보와 함께 UNVERIFIED
        # 대화 상자를 닫으면 Outlook이 미뤄 둔 작성 창을 띄우기도 함. 저장하지 못했으면 임시 보관함에 없음 (회사 PC 2026-11-10 M4)
        if eid:
            why = [f"임시 보관함에 초안으로 저장했지만 작성 창을 띄우지 못함 ({type(e).__name__}, Outlook에 열린 대화 상자가 있는지 확인)",
                   "이 명령을 다시 실행하지 말고(초안이 하나 더 생김) 사용자에게 Outlook 대화 상자를 닫고 임시 보관함에서 초안을 열어 확인하라고 알리세요"]
        else:
            why = [f"초안을 임시 보관함에 저장하지 못했고 작성 창도 띄우지 못함 ({type(e).__name__}, Outlook에 열린 대화 상자가 있는지 확인)",
                   "이 명령을 다시 실행하지 말고(창이 하나 더 뜰 수 있음) 사용자에게 Outlook 대화 상자를 닫으면 작성 창이 뜰 수 있으니 "
                   "그 창에서 확인해 직접 보내거나 저장하라고 알리세요. 창이 뜨지 않으면 대화 상자를 닫은 뒤 다시 요청하라고 알리세요"]
        return rt.ok_unverified(TOOL, command, f"메일 작성 창을 띄우지 못했습니다: '{extra['subject']}'{note}", why, **extra)
    if attach_problems or not eid:
        # 첨부를 확인하지 못했거나 초안을 저장하지 못했으면 보내는 명령을 주지 않음 (첨부가 빠진 메일은 되돌릴 수 없음)
        why = attach_problems + ([] if eid else ["초안을 임시 보관함에 저장하지 못해 보내는 명령을 만들지 못함"])
        return rt.ok_unverified(TOOL, command, f"메일 작성 창을 띄웠습니다: '{extra['subject']}'{note}",
                                why + ["보내려면 사용자가 작성 창을 확인하고 직접 [보내기]를 누르라고 알리세요"], **extra)
    problem = remember_draft(eid)
    if problem:
        # 이 도구가 만든 초안으로 기록하지 못하면 send가 거부하므로 보내는 명령을 주지 않음 (사용자가 작성 창에서 보냄)
        return rt.ok_unverified(TOOL, command, f"메일 작성 창을 띄웠습니다(초안 저장): '{extra['subject']}'{note}",
                                [f"초안을 발송 원장에 기록하지 못해 보내는 명령을 만들지 못함 ({problem})",
                                 "보내려면 사용자가 작성 창을 확인하고 직접 [보내기]를 누르라고 알리세요"], **extra)
    warnings: List[str] = []
    extra["draft_id"] = _draft_id(eid, extra["subject"], warnings)
    if warnings:
        extra["warnings"] = warnings
    # 발송은 next_commands가 아니라 confirm_commands로 (1.4.73): next_commands는 '그대로 실행'하는 다음 단계라
    # 경량 LLM이 기계적으로 실행해도 메일이 나가지 않게, 사용자 '예' 뒤에만 실행하는 명령은 키를 나눔
    extra["confirm_commands"] = [rt.next_command("outlook.py", ["send", "--id", extra["draft_id"]],
                                                 "사용자가 받는 사람·참조·제목을 보고 보내라고 한 뒤에만 이 초안을 보내기",
                                                 destructive=True)]
    tell = " original(원본 메일의 제목·보낸사람)," if original is not None else ""
    return rt.ok(TOOL, command,
                 f"메일 작성 창을 띄웠습니다(초안 저장): '{extra['subject']}' -> {_recipients_text(extra)}{note}",
                 f"VERIFIED: 초안을 저장하고 작성 창의 받는 사람·참조와 제목을 다시 읽었습니다. 아직 보내지 않았습니다. "
                 f"받는 사람·참조(cc, bcc)·제목,{tell} "
                 "첨부를 사용자에게 보여 주고 보낼지 물으세요. 사용자가 '예'라고 한 뒤에만 confirm_commands[0](send)을 1번 실행하세요. "
                 "확인 전에는 대신 발송하지 마세요. 사용자가 작성 창에서 직접 보냈으면 실행하지 마세요. "
                 "이 명령을 다시 실행하지 마세요(초안이 하나 더 생김).",
                 verified=True, **extra)


SEND_REFUSED_DIRECTIVE = "발송하지 않았습니다. 같은 명령을 다시 실행하거나 메일을 새로 만들지 말고 message를 사용자에게 그대로 알리세요."


def _send_refused(message: str) -> rt.InputError:
    """send 거부: next_commands 없이 SEND_REFUSED_DIRECTIVE (목록을 다시 보고 다른 id로 보내는 길을 주지 않음, 1.4.73)."""
    return rt.InputError(message, directive=SEND_REFUSED_DIRECTIVE)


def _in_drafts_folder(ns: Any, mail: Any) -> bool:
    """메일이 임시 보관함(기본 계정 또는 그 메일 저장소의 임시 보관함)에 있는지. 확인하지 못하면 False."""
    try:
        here = _text(mail.Parent.EntryID)
    except Exception:
        return False
    if not here:
        return False
    for get in (lambda: ns.GetDefaultFolder(OL_FOLDER_DRAFTS), lambda: mail.Parent.Store.GetDefaultFolder(OL_FOLDER_DRAFTS)):
        try:
            if _text(get().EntryID) == here:
                return True
        except Exception:
            continue
    return False


def cmd_send(args) -> Dict[str, Any]:
    """new/reply/forward가 저장한 초안(같은 항목)을 보냄. 거부(보내지 않음, 1.4.73):
    이 도구가 만든 초안이 아님(발송 원장에 없음: 사용자가 쓰다 둔 초안), 이미 send로 보냄(원장), 임시 보관함에 없음(지운 초안),
    이미 보냈거나 보낼 편지함에 넘긴 메일(Sent·Submitted), 받은 메일. 원장 기록은 잠금 아래 Send 전에: 동시에 실행한 두 send 중
    하나만 나감."""
    ns = get_namespace(get_outlook())
    try:
        eid = entry_id(args.id)
    except rt.InputError:
        raise _send_refused(f"--id {args.id}: 기억한 초안 번호가 아니라 보내지 않았습니다. 사용자에게 작성 창(임시 보관함)의 "
                            "초안을 확인하고 직접 보내라고 알리세요.") from None
    try:
        mail = ns.GetItemFromID(eid)
    except Exception:
        raise _send_refused("초안을 찾을 수 없어 보내지 않았습니다(이미 보냈거나 지웠을 수 있음). "
                            "사용자에게 보낸 편지함과 임시 보관함을 확인하라고 알리세요.") from None
    if (getattr(mail, "Sent", None) is not False or getattr(mail, "Submitted", None) is not False
            or getattr(mail, "Class", OL_MAIL) != OL_MAIL):
        raise _send_refused("이 id는 보내지 않은 초안이 아니라 보내지 않았습니다(이미 보냈거나 보낼 편지함에 넘긴 메일, 또는 받은 메일). "
                            "사용자에게 보낸 편지함과 보낼 편지함을 확인하라고 알리세요.")
    if not _in_drafts_folder(ns, mail):
        raise _send_refused("이 초안은 임시 보관함에 없어(지웠거나 다른 폴더로 옮김) 보내지 않았습니다. "
                            "사용자에게 그 메일을 직접 확인하라고 알리세요.")
    _draft_record(eid, claim=False)                  # 이 도구가 만든 초안인지, 이미 보냈는지 (작성 창을 띄우기 전에)
    extra: Dict[str, Any] = {**_recipients(mail), "subject": _text(getattr(mail, "Subject", "")),
                             "draft_id": str(args.id), "sent": False}
    unresolved = _unresolved(mail)
    if unresolved:
        mail.Display()
        extra["unresolved_recipients"] = unresolved
        return {"status": "error", "tool": TOOL, "command": "send", "error_type": "input",
                "message": f"수신자 주소를 확인할 수 없어 보내지 않고 작성 창을 띄웠습니다: {', '.join(unresolved)}",
                rt.LOG_MESSAGE_KEY: f"수신자 주소 {len(unresolved)}개를 확인할 수 없어 보내지 않고 작성 창을 띄웠습니다 (수신자 생략).",
                "agent_directive": "STOP_RETRY: 발송하지 않았습니다. 사용자에게 수신자 주소 확인을 요청하세요.", **extra}
    _draft_record(eid, claim=True)                   # 잠금 아래 다시 확인하고 '보냄'을 먼저 기록 (동시 실행·두 번 실행 방지)
    try:
        mail.Send()
    except Exception as e:
        _draft_mark_failed(eid)
        raise _send_refused(f"Outlook이 발송을 거부했습니다({type(e).__name__}). 이 초안은 send로 다시 보내지 않습니다. "
                            "사용자에게 작성 창과 보낸 편지함을 확인하고 필요하면 직접 [보내기]를 누르라고 알리세요.") from None
    extra["sent"] = True
    # R1: 발송은 Outlook이 처리하므로 실제로 나갔는지는 여기서 확인할 수 없음 (보낸 편지함 확인은 사용자 몫)
    return rt.ok_unverified(TOOL, "send", f"메일 발송을 Outlook에 맡겼습니다: '{extra['subject']}' -> {_recipients_text(extra)}",
                            ["실제 발송 여부(보낸 편지함)는 확인하지 않음. 사용자에게 보낸 편지함을 확인하라고 알리세요"], **extra)


def _load_signature(mail: Any) -> None:
    try:
        _ = mail.GetInspector                # Inspector를 만들면 기본 서명이 본문에 채워짐 (창은 안 뜸)
    except Exception:
        pass


def _target_item(app: Any, args, missing: str = "답장/전달할 메일이 없습니다. Outlook에서 메일을 선택하거나 --id를 지정하세요.",
                 original: bool = True) -> Any:
    """답장·전달·첨부 대상 메일 (--id, 없으면 열린 창·선택한 메일).
    original=True(답장·전달): 보내지 않은 메일(작성 중인 초안, Sent == False)은 원본으로 쓰지 않음. 앞서 띄운 초안 창이
    열려 있으면 그 초안에 답장·전달하게 되므로 (1.4.72)"""
    item = get_item(get_namespace(app), args.id) if args.id else current_item(app)[0]
    if item is None:
        raise rt.InputError(missing)
    if original and getattr(item, "Sent", True) is False:
        raise rt.InputError("지금 열려 있거나 고른 메일은 보내지 않은 초안이라 원본으로 쓸 수 없습니다. 원본 메일을 --id로 지정하세요: "
                            "next_commands[0]으로 목록을 보고 사용자가 말한 메일의 id를 쓰세요.",
                            directive=RELIST_DIRECTIVE, fields={"next_commands": _relist_commands()})
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
    return _finalize(mail, "new", _attach(mail, attach))


def cmd_reply(args) -> Dict[str, Any]:
    body = _body_arg(args)
    attach = _attach_paths(args.attach)
    app = get_outlook()
    original = _target_item(app, args)
    mail = original.ReplyAll() if args.all else original.Reply()
    _load_signature(mail)
    if body:
        mail.HTMLBody = _prepend_html(mail.HTMLBody, _html_prefix(body))
    return _finalize(mail, "reply", _attach(mail, attach), _original_info(original))


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
    return _finalize(mail, "forward", _attach(mail, attach), _original_info(original))


# ------------------------------------------------------------------------------
# 짧은 번호 (list/search/thread 결과의 id): 목록마다 1부터 다시 매기지 않고 계속 늘어나는 번호를 줍니다.
# 번호 -> EntryID·제목·보낸사람을 최근 ID_KEEP개(ID_KEEP_DAYS일)까지 데이터 폴더에 기억해, 옛 목록의 번호도
# 같은 메일을 가리키거나 '없는 번호' 오류가 됩니다 (다른 메일을 가리키지 않음. 긴 EntryID도 그대로 받음).
# 번호는 줄어들지 않습니다: 다음 번호를 작은 파일(NEXT_FILE)에도 따로 두고 둘 중 큰 값을 쓰며, 1부터 다시 매기지 않습니다.
# 기억 파일을 끝내 읽지 못하거나 깨졌으면 번호를 새로 매기지 않고 그 결과는 EntryID만 씁니다 (깨진 파일은 백업).
# 불러오기~저장은 잠금 파일(LOCK_FILE)의 운영체제 잠금으로 묶어 동시에 실행한 두 목록이 같은 번호를 받지 않게 합니다.
# ------------------------------------------------------------------------------

LISTING_FILE = "outlook_ids.json"
NEXT_FILE = "outlook_ids.next"   # 다음 번호만 (기억 파일이 깨져도 번호가 줄지 않게)
LOCK_FILE = "outlook_ids.lock"
ID_KEEP = 500                # 기억하는 번호 수 (최근 것부터)
ID_KEEP_DAYS = 7             # 이보다 오래 전에 받은 번호는 버림
LOCK_WAIT = 5.0              # 잠금을 기다리는 시간 상한(초)
LOCK_HOLD = 10.0             # 잠금을 쥔 동안 파일 작업 재시도에 쓰는 시간 상한(초): 이보다 오래 쥐지 않음
_SHORT_ID = re.compile(r"\d{1,9}")


class _IdsUnreadable(Exception):
    """기억 파일이 있는데 읽지 못함(잠김) 또는 깨짐. 번호를 새로 매기지 않음."""


def _listing_path() -> str:
    return os.path.join(rt.get_data_dir(), LISTING_FILE)


def _side_path(name: str) -> str:
    return os.path.join(rt.get_data_dir(), name)


def _read_next_file(deadline: Optional[float] = None) -> int:
    try:
        return max(1, int(rt.retry_fs(lambda: rt.read_text_file(_side_path(NEXT_FILE)), deadline=deadline).strip()))
    except Exception:
        return 1


def _load_ids(deadline: Optional[float] = None) -> Dict[str, Any]:
    """기억 파일 -> {"next", "items"}. 파일이 없으면 빈 목록(번호는 NEXT_FILE에서 이어서).
    있는데 읽지 못하거나 깨졌으면 _IdsUnreadable (빈 목록으로 대신하지 않음: 번호가 1부터 다시 매겨짐).
    deadline: 잠금을 쥔 실행이 재시도를 그 시각 안에 끝내게 (rt.retry_fs)."""
    path = _listing_path()
    floor = _read_next_file(deadline)
    try:
        text = rt.retry_fs(lambda: rt.read_text_file(path), deadline=deadline)
    except FileNotFoundError:
        return {"next": floor, "items": {}}
    except Exception as e:
        raise _IdsUnreadable(f"읽지 못함 ({type(e).__name__})") from e
    try:
        data = json.loads(text)
        items = data["items"]
        if not isinstance(items, dict):
            raise ValueError("items")
        nxt = int(data.get("next") or 1)
    except Exception as e:
        raise _IdsUnreadable(f"깨짐 ({type(e).__name__})") from e
    top = max((int(k) for k in items if str(k).isdigit()), default=0)
    return {"next": max(nxt, floor, top + 1), "items": items}


def _backup_broken(reason: str, deadline: Optional[float] = None) -> None:
    """깨진 기억 파일을 옆에 옮겨 두고(지우지 않음), 그 안에 보이는 가장 큰 번호로 NEXT_FILE을 올립니다.
    읽지 못한 경우(잠김)는 건드리지 않습니다."""
    if not reason.startswith("깨짐"):
        return
    path = _listing_path()
    try:
        with open(path, "rb") as f:
            raw = f.read().decode("utf-8", "replace")
        nums = [int(x) for x in re.findall(r'"next"\s*:\s*(\d{1,9})', raw)]
        nums += [int(x) + 1 for x in re.findall(r'"(\d{1,9})"\s*:', raw)]
        nxt = max([_read_next_file(deadline)] + nums)
        rt.safe_write_text(_side_path(NEXT_FILE), str(nxt), deadline=deadline)
        bak = os.path.join(rt.get_data_dir(), f"outlook_ids.broken_{rt.timestamp()}.json")
        rt.retry_fs(lambda: os.replace(path, bak), deadline=deadline)
        rt.log("WARN", f"메일 번호 기억 파일이 깨져 {os.path.basename(bak)}로 옮겼습니다. 다음 목록부터 {nxt}번부터 매깁니다.")
    except Exception as e:
        rt.log("WARN", f"깨진 메일 번호 기억 파일을 옮기지 못했습니다: {type(e).__name__}")


def _try_os_lock(fd: int) -> bool:
    """열린 잠금 파일에 운영체제 잠금을 겁니다(기다리지 않음). Windows는 msvcrt.locking(첫 1바이트), 그 밖은 fcntl.flock.
    잠금은 프로세스가 끝나면(강제 종료 포함) 운영체제가 풉니다."""
    if os.name == "nt":
        import msvcrt
        try:
            os.lseek(fd, 0, os.SEEK_SET)
            msvcrt.locking(fd, msvcrt.LK_NBLCK, 1)
            return True
        except OSError:
            return False
    import fcntl
    try:
        fcntl.flock(fd, fcntl.LOCK_EX | fcntl.LOCK_NB)
        return True
    except OSError:
        return False


def _os_unlock(fd: int) -> None:
    try:
        if os.name == "nt":
            import msvcrt
            os.lseek(fd, 0, os.SEEK_SET)
            msvcrt.locking(fd, msvcrt.LK_UNLCK, 1)
        else:
            import fcntl
            fcntl.flock(fd, fcntl.LOCK_UN)
    except OSError:
        pass


class _IdsLock:
    """outlook_ids.lock을 열어 운영체제 잠금을 겁니다. LOCK_WAIT초 안에 못 잡으면 acquired=False이고 why에 까닭.
    잡으면 deadline(지금 + LOCK_HOLD초)까지만 파일 작업을 재시도합니다.
    잠금 파일은 지우지 않고 남겨 둡니다(파일이 있다고 잠긴 것이 아님). 잠금을 쥔 실행이 끝나면 운영체제가 잠금을 풀어
    남은 잠금을 치우는 판단(pid·나이)이 필요 없습니다. 예전 방식(O_EXCL + 남은 잠금 치우기)은 치우는 사이 경합으로
    두 실행이 함께 잠금을 쥐어 번호가 겹쳤음 (4차 검토)."""

    def __init__(self) -> None:
        self.path = _side_path(LOCK_FILE)
        self.acquired = False
        self.deadline: Optional[float] = None
        self.why = ""
        self._fd: Optional[int] = None

    def __enter__(self) -> "_IdsLock":
        deadline = time.monotonic() + LOCK_WAIT
        opened_err = ""
        while True:
            if self._fd is None:
                try:
                    os.makedirs(os.path.dirname(self.path), exist_ok=True)
                    self._fd = os.open(self.path, os.O_RDWR | os.O_CREAT | getattr(os, "O_BINARY", 0), 0o666)
                except OSError as e:                       # 백신·탐색기가 잡고 있으면 짧게 기다려 다시
                    opened_err = type(e).__name__
            if self._fd is not None and _try_os_lock(self._fd):
                self.acquired = True
                self.deadline = time.monotonic() + LOCK_HOLD
                return self
            if time.monotonic() >= deadline:
                self.why = (f"잠금 파일 {LOCK_FILE}을 열지 못함 ({opened_err})" if self._fd is None
                            else f"다른 실행이 번호 파일을 쓰는 중 ({LOCK_WAIT:.0f}초 기다림)")
                self._close()
                return self
            time.sleep(0.05)

    def _close(self) -> None:
        if self._fd is not None:
            try:
                os.close(self._fd)
            except OSError:
                pass
            self._fd = None

    def __exit__(self, *exc: Any) -> None:
        if self.acquired and self._fd is not None:
            _os_unlock(self._fd)
        self.acquired = False
        self._close()


def _prune_ids(items: Dict[str, Any], now: _dt.datetime) -> Dict[str, Any]:
    """ID_KEEP_DAYS일보다 오래된 번호를 버리고 최근 ID_KEEP개만 남깁니다."""
    cut = (now - _dt.timedelta(days=ID_KEEP_DAYS)).isoformat(timespec="seconds")
    kept = {k: v for k, v in items.items() if isinstance(v, dict) and str(v.get("t", "")) >= cut}
    order = sorted(kept, key=lambda k: (str(kept[k].get("t", "")), int(k) if k.isdigit() else 0))
    return {k: kept[k] for k in order[-ID_KEEP:]}


def remember_listing(mails: List[Dict[str, Any]], what: str) -> Optional[str]:
    """목록의 메일에 짧은 번호(id)를 붙이고 번호 -> EntryID를 기억합니다. 이미 번호가 있는 메일은 같은 번호를 씁니다.
    번호를 붙이지 못하면(잠금 시간 초과, 기억 파일을 읽지 못함·깨짐, 저장 실패) id = EntryID로 두고 경고 문구를 돌려줍니다."""
    def long_ids(why: str) -> str:
        msg = f"짧은 번호를 붙이지 못해 id에 긴 EntryID를 넣었습니다({why}). id를 글자 그대로 쓰세요."
        rt.log("WARN", msg)
        for m in mails:
            m["id"] = m["entry_id"]
        return msg
    if not mails:
        return None
    with _IdsLock() as lock:
        if not lock.acquired:
            return long_ids(lock.why)
        try:
            data = _load_ids(lock.deadline)
        except _IdsUnreadable as e:
            _backup_broken(str(e), lock.deadline)
            return long_ids(f"번호 기억 파일 {e}")
        now = _dt.datetime.now()
        items = _prune_ids(data["items"], now)
        nxt = data["next"]
        by_eid = {v.get("entry_id"): k for k, v in items.items()}
        stamp = now.isoformat(timespec="seconds")
        numbers: List[int] = []
        for m in mails:
            key = by_eid.get(m["entry_id"])
            if key is None:
                key = str(nxt)
                nxt += 1
                by_eid[m["entry_id"]] = key
            items[key] = {"entry_id": m["entry_id"], "subject": m.get("subject", ""), "from": m.get("from", ""),
                          "received": m.get("received", ""), "t": stamp, "what": what[:60]}
            numbers.append(int(key))
        try:
            rt.safe_write_text(_side_path(NEXT_FILE), str(nxt), deadline=lock.deadline)   # 다음 번호 먼저 (기억 파일이 깨져도 줄지 않게)
            rt.safe_write_text(_listing_path(), json.dumps({"next": nxt, "items": _prune_ids(items, now)}, ensure_ascii=False),
                               deadline=lock.deadline)
        except Exception as e:
            return long_ids(f"저장 실패 {type(e).__name__}")
    for n, m in zip(numbers, mails):
        m["id"] = n
    return None


def compact_listing(mails: List[Dict[str, Any]]) -> None:
    """목록 결과에서 짧은 번호(id)가 붙은 메일의 긴 EntryID(140자쯤)를 뺌 (2026-11-09 검토 5장 4번: 목록 출력 크기).
    번호를 붙이지 못해 id가 EntryID면 그대로."""
    for m in mails:
        if m.get("id") != m.get("entry_id"):
            m.pop("entry_id", None)


# ------------------------------------------------------------------------------
# 발송 원장 (1.4.73): new/reply/forward가 만든 초안의 EntryID와 send로 보낸 시각. send는 원장에 있고 아직 보내지 않은
# 초안만 보냄 (사용자가 쓰다 둔 초안·이미 보낸 초안 거부). 번호 기억 파일과 같은 잠금(_IdsLock)으로 묶음.
# 읽지 못하거나 깨졌으면 send는 보내지 않음(fail-closed). 새 초안을 기록할 때 깨진 원장은 옆에 옮기고 새로 시작.
# ------------------------------------------------------------------------------

DRAFTS_FILE = "outlook_drafts.json"


def _load_drafts(deadline: Optional[float] = None) -> Dict[str, Any]:
    path = _side_path(DRAFTS_FILE)
    try:
        text = rt.retry_fs(lambda: rt.read_text_file(path), deadline=deadline)
    except FileNotFoundError:
        return {}
    except Exception as e:
        raise _IdsUnreadable(f"읽지 못함 ({type(e).__name__})") from e
    try:
        data = json.loads(text)
        if not isinstance(data, dict) or not all(isinstance(v, dict) for v in data.values()):
            raise ValueError("drafts")
    except Exception as e:
        raise _IdsUnreadable(f"깨짐 ({type(e).__name__})") from e
    return data


def _prune_drafts(data: Dict[str, Any], now: _dt.datetime) -> Dict[str, Any]:
    """ID_KEEP_DAYS일보다 오래된 초안을 버리고 최근 ID_KEEP개만 (버린 초안은 send가 거부)."""
    cut = (now - _dt.timedelta(days=ID_KEEP_DAYS)).isoformat(timespec="seconds")
    kept = sorted(((k, v) for k, v in data.items() if str(v.get("t", "")) >= cut), key=lambda kv: str(kv[1].get("t", "")))
    return dict(kept[-ID_KEEP:])


def remember_draft(eid: str) -> Optional[str]:
    """이 도구가 만든 초안으로 원장에 기록. 기록하지 못하면 까닭 (그 초안에는 send 명령을 주지 않음)."""
    with _IdsLock() as lock:
        if not lock.acquired:
            return lock.why
        try:
            data = _load_drafts(lock.deadline)
        except _IdsUnreadable as e:
            if not str(e).startswith("깨짐"):
                return f"발송 원장 {e}"
            bak = _side_path(f"outlook_drafts.broken_{rt.timestamp()}.json")
            try:
                rt.retry_fs(lambda: os.replace(_side_path(DRAFTS_FILE), bak), deadline=lock.deadline)
            except Exception as e2:
                return f"깨진 발송 원장을 옮기지 못함 ({type(e2).__name__})"
            rt.log("WARN", f"발송 원장이 깨져 {os.path.basename(bak)}로 옮기고 새로 시작했습니다 (이전 초안은 send로 보내지 않음)")
            data = {}
        now = _dt.datetime.now()
        data = _prune_drafts(data, now)
        data[eid] = {"t": now.isoformat(timespec="seconds"), "sent": ""}
        try:
            rt.safe_write_text(_side_path(DRAFTS_FILE), json.dumps(data, ensure_ascii=False), deadline=lock.deadline)
        except Exception as e:
            return f"발송 원장 저장 실패 ({type(e).__name__})"
    return None


def _draft_record(eid: str, claim: bool) -> None:
    """send 전 원장 확인: 이 도구가 만든 초안이고 아직 보내지 않았는지. claim=True면 잠금 아래 '보냄'을 먼저 기록
    (기록하지 못하면 보내지 않음). 아니면 _send_refused."""
    with _IdsLock() as lock:
        if not lock.acquired:
            raise _send_refused(f"발송 원장을 확인하지 못해 보내지 않았습니다({lock.why}). 사용자에게 알리세요.")
        try:
            data = _load_drafts(lock.deadline)
        except _IdsUnreadable as e:
            raise _send_refused(f"발송 원장을 읽지 못해({e}) 보내지 않았습니다. 사용자에게 작성 창의 초안을 확인하고 "
                                "직접 보내라고 알리세요.") from None
        now = _dt.datetime.now()
        data = _prune_drafts(data, now)
        rec = data.get(eid)
        if rec is None:
            raise _send_refused("이 메일은 이 도구(new·reply·forward)가 최근에 만든 초안이 아니라 보내지 않았습니다. "
                                "사용자가 쓰던 초안은 사용자가 작성 창에서 직접 보내라고 알리세요.")
        if rec.get("failed"):
            # 발송이 실패한 초안 (1.4.75): 이미 보냄과 구분해 알림. fail-closed는 그대로 (send로 다시 보내지 않음)
            raise _send_refused(f"발송이 실패한 초안입니다({rec['failed']}). 다시 보내지 않았습니다. "
                                "사용자에게 작성 창(임시 보관함)에서 초안을 확인하고 직접 [보내기]를 누르라고 알리세요.")
        if rec.get("sent"):
            raise _send_refused(f"이 초안은 이미 send로 보냈습니다({rec['sent']}). 다시 보내지 않았습니다. "
                                "사용자에게 보낸 편지함과 보낼 편지함을 확인하라고 알리세요.")
        if not claim:
            return
        rec["sent"] = now.isoformat(timespec="seconds")
        try:
            rt.safe_write_text(_side_path(DRAFTS_FILE), json.dumps(data, ensure_ascii=False), deadline=lock.deadline)
        except Exception as e:
            raise _send_refused(f"발송 기록을 남기지 못해({type(e).__name__}) 보내지 않았습니다. 사용자에게 알리세요.") from None


def _draft_mark_failed(eid: str) -> None:
    """Send 예외 뒤 원장에 실패 시각을 기록 (1.4.75): 다음 send 거부 문구를 '이미 보냄'과 구분. 기록하지 못해도
    '보냄' 기록이 남아 있어 다시 보내지 않음 (fail-closed). 오류는 삼킴."""
    try:
        with _IdsLock() as lock:
            if not lock.acquired:
                return
            data = _load_drafts(lock.deadline)
            rec = data.get(eid)
            if rec is None:
                return
            rec["failed"] = _dt.datetime.now().isoformat(timespec="seconds")
            rt.safe_write_text(_side_path(DRAFTS_FILE), json.dumps(data, ensure_ascii=False), deadline=lock.deadline)
    except Exception:
        pass


def _relist_commands() -> List[Dict[str, str]]:
    return [rt.next_command("outlook.py", ["list"], "메일 목록을 다시 보고 그 결과의 id 쓰기")]


RELIST_DIRECTIVE = ("같은 id로 재시도하지 마세요. next_commands[0]으로 목록을 다시 보고, 사용자가 말한 메일의 id로 다시 실행하세요. "
                    "어느 메일인지 확실하지 않으면 사용자에게 물으세요.")


def entry_id(raw: Any) -> str:
    """--id 값 -> EntryID. 짧은 번호(숫자 9자리까지)는 기억한 목록에서 찾고, 그 밖은 EntryID 그대로."""
    text = str(raw or "").strip()
    if not _SHORT_ID.fullmatch(text):
        return text
    n = int(text)
    try:
        items = _prune_ids(_load_ids()["items"], _dt.datetime.now())
    except _IdsUnreadable as e:
        raise rt.InputError(f"--id {n}: 메일 번호 기억 파일을 쓸 수 없습니다({e}). 목록을 다시 보세요. 새 목록의 id를 그대로 쓰세요.",
                            directive=RELIST_DIRECTIVE, fields={"next_commands": _relist_commands()})
    if not items:
        raise rt.InputError(f"--id {n}은(는) 목록 번호인데 기억한 메일 목록이 없습니다. outlook.py list나 search를 먼저 실행하고 "
                            "그 결과의 id를 쓰세요.", directive=RELIST_DIRECTIVE, fields={"next_commands": _relist_commands()})
    hit = items.get(str(n))
    if not hit or not hit.get("entry_id"):
        raise rt.InputError(f"--id {n}: 기억한 메일 목록에 없는 번호입니다(오래되었거나 잘못 옮긴 번호). 목록을 다시 보고 "
                            "그 결과의 id를 그대로 쓰세요.", directive=RELIST_DIRECTIVE, fields={"next_commands": _relist_commands()})
    return str(hit["entry_id"])


def get_item(ns: Any, raw: Any) -> Any:
    eid = entry_id(raw)
    try:
        return ns.GetItemFromID(eid)
    except Exception:
        raise rt.InputError("해당 id의 메일을 찾을 수 없습니다(지워졌거나 옮겨졌을 수 있음). 목록을 다시 보고 그 결과의 id를 쓰세요.",
                            directive=RELIST_DIRECTIVE, fields={"next_commands": _relist_commands()})


# ------------------------------------------------------------------------------
# 일정 (calendar, 읽기 전용)
# ------------------------------------------------------------------------------

OL_FOLDER_CALENDAR = 9
CALENDAR_MAX = 50
CALENDAR_SCAN_MAX = 5000      # Outlook이 기간 조건을 거부했을 때 일정을 하나씩 비교하는 상한


def _when(v: Any) -> str:
    try:
        return v.replace(tzinfo=None).strftime("%Y-%m-%d %H:%M")
    except Exception:
        return _text(v)


def cmd_calendar(args) -> Dict[str, Any]:
    if args.days < 1 or args.days > 92:
        raise rt.InputError(f"--days는 1~92입니다 (1 = 오늘, 7 = 오늘부터 7일). 받은 값: {args.days}. [추천] {rt.retry_command({'--days': 7})}")
    start = _dt.datetime.combine(_today(), _dt.time())
    end = start + _dt.timedelta(days=args.days)
    ns = get_namespace(get_outlook())
    folder = ns.GetDefaultFolder(OL_FOLDER_CALENDAR)
    items = folder.Items
    items.Sort("[Start]")
    items.IncludeRecurrences = True                     # 반복 일정을 날마다 펼침 (Sort 뒤, Restrict 전)
    flt = f"[Start] < '{end.strftime('%Y-%m-%d %H:%M')}' AND [End] > '{start.strftime('%Y-%m-%d %H:%M')}'"
    events: List[Dict[str, Any]] = []
    more = False
    bad = 0                                             # 읽지 못한 일정 수 (있으면 0건이어도 VERIFIED로 단정하지 않음)
    problems: List[str] = []
    scanned = 0
    try:
        source, filtered = items.Restrict(flt), True
    except Exception as e:
        # 기간 조건 거부(JIT 복구): 시작 시각 순으로 하나씩 비교 (반복 일정이 끝없이 펼쳐질 수 있어 상한)
        rt.note_recovered(f"Outlook이 일정 기간 조건을 거부해({type(e).__name__}) 일정마다 시각을 비교했습니다")
        source, filtered = items, False
    for it in source:
        if len(events) >= CALENDAR_MAX:
            more = True
            break
        scanned += 1
        if not filtered and scanned > CALENDAR_SCAN_MAX:
            problems.append(f"기간 조건 없이 일정 {CALENDAR_SCAN_MAX}개까지만 비교함 (빠진 일정이 있을 수 있음)")
            break
        try:
            s, e = it.Start, it.End
            s_ = s.replace(tzinfo=None) if hasattr(s, "replace") else None
            if s_ is not None and s_ >= end:
                break
            if not filtered:
                e_ = e.replace(tzinfo=None) if hasattr(e, "replace") else None
                if e_ is not None and e_ <= start:
                    continue
            events.append({"subject": _text(getattr(it, "Subject", "")), "start": _when(s), "end": _when(e),
                           "all_day": bool(getattr(it, "AllDayEvent", False)), "location": _text(getattr(it, "Location", "")),
                           "organizer": _text(getattr(it, "Organizer", ""))})
        except Exception:
            bad += 1
            continue
    if bad:
        problems.append(f"일정 {bad}개를 읽지 못함 (빠진 일정이 있을 수 있음)")
    period = {"since": start.date().isoformat(), "until": (end - _dt.timedelta(days=1)).date().isoformat(), "days": args.days}
    fields: Dict[str, Any] = {"period": period, "events": events, "count": len(events), "truncated": more}
    when = f"{period['since']} ~ {period['until']} ({args.days}일)"
    if problems:
        fields["unread_events"] = bad
        return rt.ok_unverified(TOOL, "calendar", f"{when} 일정 {len(events)}개 (일부 확인 못 함)", problems, **fields)
    if not events:
        return rt.ok(TOOL, "calendar", f"0건: {when} 일정 없음",
                     f"VERIFIED: 0건. {when}에 일정이 없습니다. 사용자에게 그대로 알리고 다시 실행하지 마세요.", verified=True, **fields)
    preview = f"{when} 일정 {len(events)}개" + (f" (앞 {CALENDAR_MAX}개만)" if more else "")
    return rt.ok(TOOL, "calendar", preview, "VERIFIED: 일정을 읽었습니다(읽기만, 바꾸지 않음). 날짜·시각·제목·장소를 사용자에게 알리세요."
                 + (" 더 있으면 --days를 줄여 나눠 보세요." if more else ""), verified=True, **fields)


# ------------------------------------------------------------------------------
# 일정 만들기 (event): 내 일정에만 저장. 참석자·회의 요청은 없음 (보내는 것이 없음)
# ------------------------------------------------------------------------------

OL_APPOINTMENT = 1
EVENT_MAX_MINUTES = 24 * 60 * 14
_EVENT_TIME = re.compile(r"^(\d{4})[-./](\d{1,2})[-./](\d{1,2})(?:[ T](\d{1,2}):(\d{2}))?$")


def _event_start(raw: str) -> Tuple[_dt.datetime, bool]:
    """'YYYY-MM-DD HH:MM' -> (시작, False), 'YYYY-MM-DD' -> (그날 0시, True=종일). 못 읽으면 예시와 함께 오류."""
    m = _EVENT_TIME.match(str(raw).strip())
    try:
        if m:
            day = _dt.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
            if m.group(4) is None:
                return _dt.datetime.combine(day, _dt.time()), True
            return _dt.datetime.combine(day, _dt.time(int(m.group(4)), int(m.group(5)))), False
    except ValueError:
        pass
    raise rt.InputError(f"--start '{raw}'을(를) 읽을 수 없습니다. 'YYYY-MM-DD HH:MM'(예: {_today().isoformat()} 14:00) 또는 "
                        "종일 일정은 'YYYY-MM-DD'로 쓰세요.")


def _com_time(local: _dt.datetime) -> _dt.datetime:
    """Outlook에 넣을 현지 시각. pywin32는 tzinfo 없는 시각을 현지 시각으로 보고 UTC로 바꿔 넣어, 15:00 일정이 06:00(한국 -9시간)에
    저장됐다 (회사 PC 2026-11-11 N7). 읽을 때 pywin32가 현지 시각에 UTC 표시를 붙여 주는 것과 같게, 현지 시각에 UTC 표시를 붙여 넣으면
    바꾸지 않고 그대로 들어간다. 저장 뒤 다시 읽어 시각을 비교한다 (R1)."""
    return local.replace(tzinfo=_dt.timezone.utc)


def _same_event(folder: Any, subject: str, start: _dt.datetime) -> Optional[Any]:
    """같은 제목·같은 시작의 일정 (다시 실행해도 일정이 둘이 되지 않게)."""
    end = start + _dt.timedelta(minutes=1)
    try:
        items = folder.Items.Restrict(f"[Start] >= '{start.strftime('%Y-%m-%d %H:%M')}' AND [Start] < '{end.strftime('%Y-%m-%d %H:%M')}'")
        for it in items:
            if _text(getattr(it, "Subject", "")) == subject:
                return it
    except Exception:
        return None
    return None


def _overlaps(folder: Any, start: _dt.datetime, end: _dt.datetime, skip: str) -> List[str]:
    """겹치는 다른 일정 (제목 시작~끝), 최대 5개. 읽지 못하면 빈 목록."""
    out: List[str] = []
    try:
        items = folder.Items
        items.Sort("[Start]")
        items.IncludeRecurrences = True
        for it in items.Restrict(f"[Start] < '{end.strftime('%Y-%m-%d %H:%M')}' AND [End] > '{start.strftime('%Y-%m-%d %H:%M')}'"):
            if _text(getattr(it, "EntryID", "")) == skip:
                continue
            out.append(f"{_text(getattr(it, 'Subject', ''))} {_when(it.Start)}~{_when(it.End)[-5:]}")
            if len(out) >= 5:
                break
    except Exception:
        return []
    return out


def _day_events(folder: Any, day: _dt.date) -> List[Any]:
    """그날 시작하는 일정 (반복 일정은 펼치지 않음: 지우면 반복 전체가 지워지는 원본만 보임). 읽지 못하면 빈 목록."""
    lo = _dt.datetime.combine(day, _dt.time())
    hi = lo + _dt.timedelta(days=1)
    try:
        return list(folder.Items.Restrict(f"[Start] >= '{lo.strftime('%Y-%m-%d %H:%M')}' AND [Start] < '{hi.strftime('%Y-%m-%d %H:%M')}'"))
    except Exception:
        return []


def _event_delete(args, subject: str, start: _dt.datetime, all_day: bool) -> Dict[str, Any]:
    """event --action delete: 제목과 시작(시각, 또는 날짜만이면 그날)이 맞는 내 일정 하나를 지움. --dry-run이면 미리보기.
    반복 일정·회의(참석자가 있는 일정)는 지우지 않음: 반복 전체가 지워지거나 취소 알림이 필요해 Outlook에서 직접."""
    warnings = [f"{o}는 일정 지우기에 쓰지 않아 무시했습니다" for o, v in (("--minutes", args.minutes), ("--location", args.location),
                                                                  ("--body", args.body), ("--body-file", args.body_file)) if v]
    folder = get_namespace(get_outlook()).GetDefaultFolder(OL_FOLDER_CALENDAR)
    day = _day_events(folder, start.date())

    def at(it: Any) -> Optional[_dt.datetime]:
        try:
            return it.Start.replace(tzinfo=None)
        except Exception:
            return None
    hits = [it for it in day if _text(getattr(it, "Subject", "")) == subject and (all_day or at(it) == start)]
    shown = start.strftime("%Y-%m-%d") if all_day else _when(start)
    if not hits:
        # 그날의 다른 일정을 [후보]로 (같은 제목의 다른 시각, 다른 제목): 고른 번호의 전체 명령 (사용자가 고른 뒤에만)
        others = [(it, at(it)) for it in day if at(it) is not None][:10]
        if not others:
            raise rt.InputError(f"{start.strftime('%Y-%m-%d')}에 시작하는 일정이 없어 지우지 않았습니다 ('{subject}' {shown}). "
                                "[추천] " + rt.next_command("outlook.py", ["calendar", "--days", "7"], "")["command"])
        names = [f"{_text(getattr(it, 'Subject', ''))} {_when(t)}" for it, t in others]
        cmds = [rt.next_command("outlook.py", ["event", "--action", "delete", "--subject", _text(getattr(it, "Subject", "")),
                                               "--start", _when(t), "--dry-run"], "")["command"] for it, t in others]
        raise rt.UserDecisionError(
            f"'{subject}' {shown} 일정이 없어 지우지 않았습니다. 그날 시작하는 일정:", "지우지 않았습니다. [후보]를 사용자에게 보여 주고, "
            "사용자가 고른 번호의 confirm_commands를 1번 실행하세요 (미리보기, 바꾸지 않음).", candidates=names, commands=cmds,
            log_message=f"지울 일정 없음 ({start.strftime('%Y-%m-%d')})")
    if len(hits) > 1:
        names = [f"{subject} {_when(at(it))}" for it in hits]
        times = {_when(at(it)) for it in hits}
        if len(times) < len(hits):              # 같은 제목·같은 시각이 둘 이상: 도구로는 고를 수 없음
            raise rt.InputError(f"'{subject}' {shown} 일정이 {len(hits)}개라 어느 것을 지울지 정할 수 없습니다. Outlook 일정에서 직접 지우세요.",
                                error_type="not_supported", log_message="같은 제목·시각 일정 여러 개")
        cmds = [rt.next_command("outlook.py", ["event", "--action", "delete", "--subject", subject, "--start", t, "--dry-run"], "")["command"]
                for t in [_when(at(it)) for it in hits]]
        raise rt.UserDecisionError(f"그날 '{subject}' 일정이 {len(hits)}개입니다. 지울 것을 고르세요:",
                                   "지우지 않았습니다. [후보]를 보여 주고 사용자가 고른 번호의 confirm_commands를 1번 실행하세요 (미리보기).",
                                   candidates=names, commands=cmds, log_message=f"같은 제목 일정 {len(hits)}개")
    it = hits[0]
    when = f"{_when(at(it))} ~ {_when(it.End)[-5:]}"
    if bool(getattr(it, "IsRecurring", False)) or int(getattr(it, "MeetingStatus", 0) or 0) != 0:
        what = "반복 일정" if bool(getattr(it, "IsRecurring", False)) else "회의(참석자가 있는 일정)"
        raise rt.InputError(f"'{subject}' {when}은(는) {what}이라 도구로 지우지 않습니다 (반복 전체가 지워지거나 취소 알림이 필요함). "
                            "Outlook 일정에서 직접 지우세요.", error_type="not_supported", log_message=f"{what} 지우기")
    desc = f"'{subject}' {when}" + (f" ({_text(getattr(it, 'Location', ''))})" if _text(getattr(it, "Location", "")) else "")
    extra: Dict[str, Any] = {"subject": subject, "start": _when(at(it)), "end": _when(it.End)}
    if warnings:
        extra["warnings"] = warnings
    if args.dry_run:
        return rt.dry_run_result(TOOL, "event", f"미리보기: 일정 {desc}을(를) 지웁니다 (바꾸지 않음)", 1, [f"지우기: {desc}"],
                                 rt.dry_run_command(TOOL, f"일정 {desc} 지우기"), **extra)
    t0 = at(it)
    it.Delete()
    left = [x for x in _day_events(folder, start.date()) if _text(getattr(x, "Subject", "")) == subject and at(x) == t0]
    preview = f"일정을 지웠습니다: {desc} (Outlook 지운 편지함으로 옮겨짐)"
    if left:
        return rt.ok_unverified(TOOL, "event", preview, ["지운 뒤 같은 일정이 아직 보임: 다시 실행하지 말고 Outlook 일정에서 확인하라고 알리세요"],
                                deleted=True, **extra)
    return rt.ok(TOOL, "event", preview, "VERIFIED: 일정을 지우고 다시 찾아 없는 것을 확인했습니다. 같은 명령을 다시 실행하지 마세요. "
                 "되돌리려면 사용자가 Outlook 지운 편지함에서 일정을 다시 옮기면 됩니다.", verified=True, deleted=True, **extra)


def cmd_event(args) -> Dict[str, Any]:
    subject = (args.subject or "").strip()
    if not subject:
        raise rt.InputError("--subject(일정 제목)를 주세요.")
    start, all_day = _event_start(args.start or "")
    if args.action == "delete":
        return _event_delete(args, subject, start, all_day)
    if args.dry_run:
        raise rt.dry_run_unsupported(TOOL, "event (일정 만들기)")
    minutes = 24 * 60 if all_day and args.minutes is None else (args.minutes if args.minutes is not None else 60)
    if minutes < 1 or minutes > EVENT_MAX_MINUTES:
        raise rt.InputError(f"--minutes는 1~{EVENT_MAX_MINUTES}입니다 (받은 값: {minutes}). [추천] {rt.retry_command({'--minutes': 60})}")
    end = start + _dt.timedelta(minutes=minutes)
    body = _body_arg(args)                              # 메모 파일 확인은 Outlook에 연결하기 전에
    warnings: List[str] = []
    if start.date() < _today():
        warnings.append(f"지난 날짜({start.date().isoformat()})의 일정입니다")
    ns = get_namespace(get_outlook())
    folder = ns.GetDefaultFolder(OL_FOLDER_CALENDAR)
    when = f"{start.strftime('%Y-%m-%d')} 종일" if all_day and minutes == 24 * 60 else f"{_when(start)} ~ {_when(end)[-5:]}"
    old = _same_event(folder, subject, start)
    if old is not None:
        return rt.ok(TOOL, "event", f"이미 있는 일정: '{subject}' {when} (새로 만들지 않음)",
                     "VERIFIED: 같은 제목·시작 시각의 일정이 이미 있어 새로 만들지 않았습니다. 사용자에게 그대로 알리고 다시 실행하지 마세요.",
                     verified=True, created=False, subject=subject, start=_when(old.Start), end=_when(old.End))
    app = get_outlook()
    item = app.CreateItem(OL_APPOINTMENT)
    item.Subject = subject
    item.Start = _com_time(start)
    item.Duration = minutes
    if all_day and minutes % (24 * 60) == 0:
        item.AllDayEvent = True
    if args.location:
        item.Location = args.location
    if body:
        item.Body = body
    item.ReminderSet = True
    item.Save()
    eid = _text(getattr(item, "EntryID", ""))
    # R1: 저장한 일정을 다시 찾아 제목·시작·끝 확인
    problems: List[str] = []
    try:
        got = ns.GetItemFromID(eid)
        g_start, g_end = got.Start.replace(tzinfo=None), got.End.replace(tzinfo=None)
        if _text(got.Subject) != subject:
            problems.append("다시 읽은 제목이 다름")
        if g_start != start or g_end != end:
            problems.append(f"다시 읽은 시각이 다름 ({_when(g_start)} ~ {_when(g_end)})")
    except Exception as e:
        problems.append(f"저장한 일정을 다시 읽지 못함 ({type(e).__name__})")
    fields: Dict[str, Any] = {"created": True, "subject": subject, "start": _when(start), "end": _when(end), "all_day": bool(all_day),
                              "location": args.location or ""}
    busy = _overlaps(folder, start, end, eid)
    if busy:
        warnings.append("겹치는 일정: " + ", ".join(busy))
    if warnings:
        fields["warnings"] = warnings
    preview = f"일정을 만들었습니다: '{subject}' {when}" + (f" ({args.location})" if args.location else "")
    if problems:
        return rt.ok_unverified(TOOL, "event", preview, problems + ["다시 실행하지 말고(일정이 둘이 됨) Outlook 일정에서 확인하라고 알리세요"], **fields)
    return rt.ok(TOOL, "event", preview, "VERIFIED: 일정을 저장하고 다시 읽어 확인했습니다. 같은 명령을 다시 실행하지 마세요. "
                 "warnings(겹치는 일정 등)가 있으면 함께 알리세요.", verified=True, **fields)


# ------------------------------------------------------------------------------
# 대화 묶음 (thread, 읽기 전용): 같은 대화 제목의 메일을 받은 편지함과 보낸 편지함에서 시간 순으로
# ------------------------------------------------------------------------------

THREAD_MAX = 20
THREAD_BODY_CHARS = 400
OL_FOLDER_INBOX, OL_FOLDER_SENT = 6, 5


THREAD_SCAN_MAX = 2000        # 대화 제목 조건을 거부했을 때 폴더마다 비교하는 메일 수 상한


def _thread_items(ns: Any, topic: str, limited: Optional[List[int]] = None) -> List[Any]:
    """limited: 상한에 닿아 다 비교하지 못한 폴더 번호를 담음 (그 폴더의 오래된 메일은 빠질 수 있음)."""
    limited = limited if limited is not None else []
    found: List[Any] = []
    flt = "[ConversationTopic] = '" + topic.replace("'", "''") + "'"
    for fid in (OL_FOLDER_INBOX, OL_FOLDER_SENT):
        try:
            items = ns.GetDefaultFolder(fid).Items
        except Exception:
            continue
        try:
            hits = list(items.Restrict(flt))
        except Exception:                               # 조건을 거부하면 메일마다 대화 제목 비교 (최근 2,000통까지)
            rt.note_recovered("Outlook이 대화 제목 조건을 거부해 메일마다 대화 제목을 비교했습니다")
            hits = []
            for k, it in enumerate(items):
                if k >= THREAD_SCAN_MAX:
                    limited.append(fid)
                    break
                try:
                    if _text(it.ConversationTopic) == topic:
                        hits.append(it)
                except Exception:
                    continue
        found += [it for it in hits if getattr(it, "Class", OL_MAIL) == OL_MAIL]
    return found


def cmd_thread(args) -> Dict[str, Any]:
    ns = get_namespace(get_outlook())
    item = get_item(ns, args.id)
    topic = _text(getattr(item, "ConversationTopic", "")) or re.sub(r"^\s*((re|fw|fwd|회신|전달)\s*:\s*)+", "",
                                                                    _text(getattr(item, "Subject", "")), flags=re.I)
    limited: List[int] = []
    mails = _thread_items(ns, topic, limited)
    seen, uniq = set(), []
    for m in mails:
        k = _text(getattr(m, "EntryID", ""))
        if k not in seen:
            seen.add(k)
            uniq.append(m)
    if _text(getattr(item, "EntryID", "")) not in seen:
        uniq.append(item)
    uniq.sort(key=lambda m: _received(m) or _dt.datetime.min)
    out = []
    for m in uniq[-THREAD_MAX:]:
        s = summary(m)
        body, cut = rt.clamp_text(_text(getattr(m, "Body", "")).replace("\r\n", "\n").strip(), THREAD_BODY_CHARS)
        s.update(body=body, body_truncated=cut)
        out.append(s)
    id_warn = remember_listing(out, f"thread {topic[:40]}")
    compact_listing(out)
    fields = {"topic": topic, "mails": out, "count": len(uniq), "shown": len(out), "truncated": len(uniq) > len(out)}
    if id_warn:
        fields["warnings"] = [id_warn]
    if len(out) > 0:
        fields["next_commands"] = [rt.next_command("outlook.py", ["read", "--id", str(out[-1]["id"])], "가장 최근 메일 본문 전체 읽기")]
    if limited:
        names = ", ".join("받은 편지함" if f == OL_FOLDER_INBOX else "보낸 편지함" for f in limited)
        return rt.ok_unverified(TOOL, "thread", f"대화 '{topic}' 메일 {len(uniq)}통 (일부 폴더는 최근 {THREAD_SCAN_MAX}통만 비교)",
                                [f"{names}에서 최근 {THREAD_SCAN_MAX}통까지만 비교해 더 오래된 같은 대화 메일은 빠졌을 수 있음"], **fields)
    return rt.ok(TOOL, "thread", f"대화 '{topic}' 메일 {len(uniq)}통 (오래된 순)" + (f", 최근 {len(out)}통만" if len(uniq) > len(out) else ""),
                 "VERIFIED: 같은 대화의 메일을 읽었습니다(읽기만). 시간 순으로 누가 무엇을 보냈는지 요약해 알리세요.", verified=True, **fields)


# ==============================================================================
# 4. 진입점
# ==============================================================================

def build_cli():
    parser, sub = rt.build_parser(TOOL, "아웃룩 메일")
    folder_help = "폴더: inbox(기본), sent, drafts, deleted, junk 또는 '계정\\폴더\\하위폴더'"

    def dated(p):
        p.add_argument("--days", type=int, help="오늘 포함 최근 N일에 받은 메일만 (1 = 오늘)")
        p.add_argument("--since", help="이 날짜(YYYY-MM-DD)부터 받은 메일만 (--days와 함께 쓰지 않음)")
        p.add_argument("--until", help="이 날짜(YYYY-MM-DD, 그날 포함)까지 받은 메일만 (--since와 함께 쓰면 기간)")
        p.add_argument("--with-attachments", action="store_true", help="첨부가 있는 메일만 (본문 삽입 그림은 첨부로 세지 않음)")

    p = sub.add_parser("list", help="최근 메일 목록")
    p.add_argument("--folder", help=folder_help)
    p.add_argument("--unread", action="store_true", help="안 읽은 메일만")
    p.add_argument("--limit", type=int, default=10, help=f"최대 몇 통 (기본 10, 최대 {MAX_LIMIT})")
    dated(p)

    p = sub.add_parser("search", help="메일 검색")
    p.add_argument("--query", help="찾을 말")
    p.add_argument("--from", dest="sender", help="보낸사람 이름/주소 일부")
    p.add_argument("--where", choices=["all", "subject", "body"], default="all", help="검색 범위 (기본 제목+본문)")
    p.add_argument("--folder", help=folder_help)
    p.add_argument("--limit", type=int, default=10, help=f"최대 몇 통 (기본 10, 최대 {MAX_LIMIT})")
    dated(p)

    for name, hlp in (("read", "메일 읽기 (--id)"), ("current", "지금 보고 있거나 선택한 메일 읽기")):
        p = sub.add_parser(name, help=hlp)
        if name == "read":
            p.add_argument("--id", required=True, help="list/search 결과의 id (짧은 번호 3 또는 긴 id)")
        p.add_argument("--save", nargs="?", const=True, help="본문을 md 파일로 저장 (첨부 내려받기는 attachments --save)")

    p = sub.add_parser("attachments", help="첨부 목록, --save로 내려받기")
    p.add_argument("--id", help="메일 id (생략 시 지금 보고 있거나 선택한 메일)")
    p.add_argument("--save", nargs="?", const=True,
                   help=f"첨부 저장. 폴더 생략 시 다운로드\\{ATTACH_ROOT_NAME}\\<시각>_<제목>, 폴더는 전체 경로 (없으면 만듦)")
    p.add_argument("--name", action="append", help="저장할 첨부 이름 일부 (대소문자 무시, 여러 개는 ; 또는 옵션 반복)")
    p.add_argument("--index", action="append", help="저장할 첨부 번호 (목록의 index, 예: 1,3 또는 2-4)")

    p = sub.add_parser("thread", help="같은 대화의 메일을 시간 순으로 (읽기만)")
    p.add_argument("--id", required=True, help="대화 속 메일 하나의 id (list/search 결과의 짧은 번호 또는 긴 id)")

    p = sub.add_parser("calendar", help="일정 (읽기만)")
    p.add_argument("--days", type=int, default=7, help="오늘부터 N일 (기본 7, 1 = 오늘, 최대 92)")

    p = sub.add_parser("event", help="내 일정 만들기·지우기 (참석자·회의 요청 없음, 보내는 것 없음)")
    p.add_argument("--action", choices=["add", "delete"], default="add",
                   help="add(기본): 만들기, delete: 제목·시작이 맞는 일정 지우기 (--dry-run으로 먼저 확인)")
    p.add_argument("--subject", required=True, help="일정 제목")
    p.add_argument("--start", required=True, help="시작 'YYYY-MM-DD HH:MM' (종일 일정은 'YYYY-MM-DD', 지우기에서 날짜만 주면 그날)")
    p.add_argument("--dry-run", action="store_true", help="지우기 미리보기 (바꾸지 않음)")
    p.add_argument("--minutes", type=int, help="길이(분, 기본 60, 종일은 하루)")
    p.add_argument("--location", help="장소")
    p.add_argument("--body", help="메모 (줄바꿈은 \\n)")
    p.add_argument("--body-file", help="메모 파일 (.txt/.md)")

    p = sub.add_parser("contacts", help="연락처·주소록에서 이름(일부)이나 메일 주소로 찾기 (읽기만)")
    p.add_argument("--name", required=True, help="이름 일부 또는 메일 주소 일부 (예: 홍길동)")

    p = sub.add_parser("folders", help="폴더 구조")
    p.add_argument("--depth", type=int, default=2, help="하위 폴더 깊이 (기본 2, 최대 5, 폴더 150개까지)")

    def compose(p, to=True):
        if to:
            p.add_argument("--to", help="받는 사람 (여러 명은 ; 로 구분)")
        p.add_argument("--body", help="본문 (줄바꿈은 \\n)")
        p.add_argument("--body-file", help="본문 파일 (.txt/.md)")
        p.add_argument("--attach", help="첨부 파일 전체 경로 (여러 개는 ; 로 구분, 쉼표도 됨. 파일명만 쓰면 data 폴더)")

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

    p = sub.add_parser("send", help="new/reply/forward가 저장한 초안 보내기 (사용자가 확인한 뒤에만)")
    p.add_argument("--id", required=True, help="new/reply/forward 결과의 draft_id (confirm_commands[0]을 그대로)")
    return parser


HANDLERS = {"list": cmd_list, "search": cmd_search, "read": cmd_read, "current": cmd_current, "folders": cmd_folders,
            "attachments": cmd_attachments, "new": cmd_new, "reply": cmd_reply, "forward": cmd_forward, "send": cmd_send,
            "thread": cmd_thread, "calendar": cmd_calendar, "event": cmd_event, "contacts": cmd_contacts}


def main(argv: Optional[Sequence[str]] = None) -> None:
    args = build_cli().parse_args(argv)
    rt.run_cli(TOOL, args.command, lambda: HANDLERS[args.command](args))


if __name__ == "__main__":
    main()
