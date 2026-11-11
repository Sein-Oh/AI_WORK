# -*- coding: utf-8 -*-
"""EML -> Outlook 가져오기 시험 (사람이 실행하는 시험 스크립트, 에이전트 출력 계약 대상 아님)

목적: 일반망에서 받아 온 .eml을 Outlook에 창 없이(백그라운드) 넣을 수 있는지, 넣은 메일이 "받은 메일"로 보이는지 확인합니다.
세 가지 방법을 같은 eml 몇 통에 각각 시험해 결과를 비교합니다.

  M1 그대로    Namespace.OpenSharedItem(eml) -> Move(시험 폴더)
  M2 속성      M1 + 저장 전에 메일 상태(PR_MESSAGE_FLAGS: 읽음, 보내지 않음 표시 없음)와 보낸·받은 시각(eml의 Date)을 넣음
  M3 게시물    시험 폴더에 게시물(IPM.Post)을 만들고 메일(IPM.Note)로 바꾼 뒤, eml에서 읽은 제목·본문·보낸 사람·받는 사람·
               시각·첨부를 넣음 (Outlook에서 "보낸 적 없는 메일"로 남지 않게 알려진 방법)

사용법 (Windows, Outlook 2016 이상, Outlook 프로필이 있는 PC):
    python eml_import_trial.py --folder "<eml 폴더>"              :: eml 5통으로 시험 (기본)
    python eml_import_trial.py --folder "<eml 폴더>" --limit 20   :: 20통으로

- 받은편지함 아래에 시험 폴더 "eml 시험"(그 아래 M1_그대로, M2_속성, M3_게시물)을 만들고 그 안에만 넣습니다.
  기존 메일과 원본 eml은 바꾸지 않습니다. 시험이 끝나면 Outlook에서 "eml 시험" 폴더를 지우면 됩니다.
- 결과는 이 파일 옆의 eml_import_trial_<시각>.md 하나입니다. 메일 내용(제목, 주소, 본문)은 적지 않고
  같은지 여부(OK/FAIL)와 걸린 시간만 적으므로 그대로 올려도 됩니다.
- 결과를 본 뒤 Outlook에서 세 폴더의 메일을 직접 열어 보고, 아래를 눈으로 확인해 주세요.
  1. "보내지 않은 메시지" 같은 안내 띠가 보이는지   2. 목록의 날짜가 eml의 날짜인지   3. 보낸 사람·첨부·한글 제목이 맞는지
"""
import argparse
import email
import os
import re
import sys
import tempfile
import time
from datetime import datetime, timezone
from email import policy
from email.header import decode_header, make_header
from email.utils import getaddresses, parsedate_to_datetime
from html import unescape

HERE = os.path.dirname(os.path.abspath(__file__))
TRIAL_FOLDER = "eml 시험"
METHODS = (("M1", "M1_그대로"), ("M2", "M2_속성"), ("M3", "M3_게시물"))

# MAPI 속성 (PropertyAccessor 태그)
PR_MESSAGE_FLAGS = "http://schemas.microsoft.com/mapi/proptag/0x0E070003"
PR_MESSAGE_DELIVERY_TIME = "http://schemas.microsoft.com/mapi/proptag/0x0E060040"
PR_CLIENT_SUBMIT_TIME = "http://schemas.microsoft.com/mapi/proptag/0x00390040"
PR_SENDER_NAME = "http://schemas.microsoft.com/mapi/proptag/0x0C1A001F"
PR_SENDER_EMAIL = "http://schemas.microsoft.com/mapi/proptag/0x0C1F001F"
PR_SENDER_ADDRTYPE = "http://schemas.microsoft.com/mapi/proptag/0x0C1E001F"
PR_SENT_REP_NAME = "http://schemas.microsoft.com/mapi/proptag/0x0042001F"
PR_SENT_REP_EMAIL = "http://schemas.microsoft.com/mapi/proptag/0x0065001F"
PR_SENT_REP_ADDRTYPE = "http://schemas.microsoft.com/mapi/proptag/0x0064001F"
PR_DISPLAY_TO = "http://schemas.microsoft.com/mapi/proptag/0x0E04001F"
PR_DISPLAY_CC = "http://schemas.microsoft.com/mapi/proptag/0x0E03001F"
PR_INTERNET_MESSAGE_ID = "http://schemas.microsoft.com/mapi/proptag/0x1035001F"
MSGFLAG_READ = 1


# ---------------------------------------------------------------------------
# 1. eml 읽기 (표준 라이브러리, Outlook 없이)
# ---------------------------------------------------------------------------
def _decode_bytes(raw, declared=None):
    for enc in (declared, "utf-8", "cp949", "euc-kr"):
        if not enc:
            continue
        try:
            return raw.decode(enc)
        except (UnicodeDecodeError, LookupError):
            continue
    return raw.decode("latin-1")


def decode_header_value(value):
    """RFC2047 인코딩과 인코딩 표시 없는 8비트(CP949 등) 헤더 (office-cli reader 1.0의 parse_eml과 같은 규칙)."""
    if value is None:
        return ""
    s = str(value)
    try:
        raw = s.encode("ascii", "surrogateescape")
        if any(b > 127 for b in raw):
            return _decode_bytes(raw).strip()
    except UnicodeEncodeError:
        pass
    try:
        return str(make_header(decode_header(s))).strip()
    except Exception:
        return s.strip()


def strip_html(html):
    html = re.sub(r"(?is)<(script|style).*?</\1>", " ", html)
    html = re.sub(r"(?i)<br\s*/?>|</p>|</div>|</tr>|</li>", "\n", html)
    text = unescape(re.sub(r"<[^>]+>", " ", html))
    return "\n".join(ln.strip() for ln in text.splitlines() if ln.strip())


def parse_eml(path):
    """eml 한 통 -> dict (제목, 보낸 사람 이름·주소, 받는 사람·참조 표시, 날짜(UTC), 본문 text/html, 첨부 [(이름, bytes)], Message-ID)."""
    with open(path, "rb") as f:
        msg = email.message_from_binary_file(f, policy=policy.compat32)
    raw = {}
    for k, v in msg.raw_items():
        raw.setdefault(k.lower(), v)
    frm = decode_header_value(raw.get("from"))
    name, addr = (getaddresses([frm]) or [("", "")])[0]
    date = None
    if msg.get("Date"):
        try:
            date = parsedate_to_datetime(str(msg.get("Date")))
            if date.tzinfo is None:
                date = date.replace(tzinfo=timezone.utc)
            date = date.astimezone(timezone.utc).replace(tzinfo=None)
        except Exception:
            date = None
    plain, html, atts = [], [], []

    def walk(part):
        ctype = part.get_content_type()
        if ctype == "message/rfc822":                       # 전달된 메일: 첨부로 (본문에 섞지 않음)
            atts.append(("첨부된 메일.eml", part.as_bytes() if hasattr(part, "as_bytes") else b""))
            return
        if part.is_multipart():
            for sub in part.get_payload():
                walk(sub)
            return
        filename = part.get_filename()
        disp = (part.get("Content-Disposition") or "").lower()
        payload = part.get_payload(decode=True)
        if filename or disp.startswith("attachment"):
            atts.append((decode_header_value(filename) if filename else "이름 없는 첨부", payload or b""))
            return
        if payload is None:
            return
        text = _decode_bytes(payload, part.get_content_charset())
        if ctype == "text/plain":
            plain.append(text)
        elif ctype == "text/html":
            html.append(text)

    walk(msg)
    return {"subject": decode_header_value(raw.get("subject")), "from_name": name or addr, "from_addr": addr,
            "to": decode_header_value(raw.get("to")), "cc": decode_header_value(raw.get("cc")), "date": date,
            "text": "\n\n".join(plain), "html": "\n".join(html), "attachments": atts,
            "message_id": str(msg.get("Message-ID") or "").strip()}


# ---------------------------------------------------------------------------
# 2. Outlook에 넣기 (방법 셋)
# ---------------------------------------------------------------------------
def subfolder(parent, name):
    for i in range(1, parent.Folders.Count + 1):
        if parent.Folders.Item(i).Name == name:
            return parent.Folders.Item(i)
    return parent.Folders.Add(name)


def _set(pa, tag, value, errors):
    try:
        pa.SetProperty(tag, value)
    except Exception as e:                                   # noqa: BLE001  기록만 하고 계속 (어느 속성이 막히는지 보는 시험)
        errors.append(f"{tag.rsplit('/', 1)[-1]}: {type(e).__name__}")


def import_m1(ns, path, folder, info, errors):
    item = ns.OpenSharedItem(path)
    return item.Move(folder)


def import_m2(ns, path, folder, info, errors):
    item = ns.OpenSharedItem(path)
    pa = item.PropertyAccessor
    _set(pa, PR_MESSAGE_FLAGS, MSGFLAG_READ, errors)
    if info["date"] is not None:
        _set(pa, PR_MESSAGE_DELIVERY_TIME, info["date"], errors)
        _set(pa, PR_CLIENT_SUBMIT_TIME, info["date"], errors)
    return item.Move(folder)


def import_m3(ns, path, folder, info, errors, tmp):
    post = folder.Items.Add("IPM.Post")
    post.MessageClass = "IPM.Note"
    post.Subject = info["subject"]
    if info["html"]:
        post.HTMLBody = info["html"]
    else:
        post.Body = info["text"]
    pa = post.PropertyAccessor
    for tag, value in ((PR_SENDER_NAME, info["from_name"]), (PR_SENDER_EMAIL, info["from_addr"]), (PR_SENDER_ADDRTYPE, "SMTP"),
                       (PR_SENT_REP_NAME, info["from_name"]), (PR_SENT_REP_EMAIL, info["from_addr"]),
                       (PR_SENT_REP_ADDRTYPE, "SMTP"), (PR_DISPLAY_TO, info["to"]), (PR_DISPLAY_CC, info["cc"]),
                       (PR_INTERNET_MESSAGE_ID, info["message_id"])):
        if value:
            _set(pa, tag, value, errors)
    _set(pa, PR_MESSAGE_FLAGS, MSGFLAG_READ, errors)
    if info["date"] is not None:
        _set(pa, PR_MESSAGE_DELIVERY_TIME, info["date"], errors)
        _set(pa, PR_CLIENT_SUBMIT_TIME, info["date"], errors)
    for k, (name, data) in enumerate(info["attachments"]):
        safe = re.sub(r'[\\/:*?"<>|]', "_", name) or f"첨부{k + 1}"
        p = os.path.join(tmp, f"{k + 1}_{safe}")
        with open(p, "wb") as f:
            f.write(data)
        post.Attachments.Add(p)
    post.Save()
    return post


# ---------------------------------------------------------------------------
# 3. 넣은 메일 확인 (내용은 적지 않고 같은지만)
# ---------------------------------------------------------------------------
def _get(obj, attr, default=None):
    try:
        return getattr(obj, attr)
    except Exception:                                        # noqa: BLE001
        return default


def _as_utc_naive(t):
    """Outlook 시각(현지 시각, pywintypes) -> UTC naive datetime."""
    try:
        local = datetime(t.year, t.month, t.day, t.hour, t.minute, t.second)
        return local.astimezone(timezone.utc).replace(tzinfo=None)
    except Exception:                                        # noqa: BLE001
        return None


def check(item, info):
    sent = _get(item, "Sent")
    received = _as_utc_naive(_get(item, "ReceivedTime"))
    date_ok = (info["date"] is not None and received is not None and abs((received - info["date"]).total_seconds()) <= 120)
    sender = (_get(item, "SenderEmailAddress") or "").lower()
    return {
        "보낸 메일로 보임(Sent)": "OK" if sent else "FAIL (초안처럼 보일 수 있음)",
        "받은 시각 = eml 날짜": "OK" if date_ok else ("eml에 날짜 없음" if info["date"] is None else "FAIL"),
        "제목 같음": "OK" if (_get(item, "Subject") or "") == info["subject"] else "FAIL",
        "보낸 사람 주소 같음": "OK" if info["from_addr"] and sender == info["from_addr"].lower() else "FAIL",
        "첨부 수 같음": "OK" if (_get(item, "Attachments") is not None and item.Attachments.Count == len(info["attachments"])) else "FAIL",
        "읽음 표시": "읽음" if _get(item, "UnRead") is False else "안 읽음",
        "MessageClass": str(_get(item, "MessageClass") or "-"),
    }


# ---------------------------------------------------------------------------
# 4. 실행과 결과 파일
# ---------------------------------------------------------------------------
def main(argv=None):
    p = argparse.ArgumentParser(description="EML -> Outlook 가져오기 시험 (받은편지함\\eml 시험 폴더에만 넣음)")
    p.add_argument("--folder", required=True, help="eml 파일이 든 폴더")
    p.add_argument("--limit", type=int, default=5, help="시험할 eml 수 (기본 5)")
    args = p.parse_args(argv)
    files = sorted(os.path.join(args.folder, f) for f in os.listdir(args.folder) if f.lower().endswith(".eml"))[:args.limit]
    if not files:
        print(f"[ERROR] {args.folder}에 .eml 파일이 없습니다.")
        return 1
    try:
        import win32com.client as w32
    except ImportError:
        print("[ERROR] pywin32가 없습니다 (Windows + Outlook에서 실행하세요).")
        return 1
    print(f"[INFO] eml {len(files)}통으로 시험합니다. 받은편지함\\{TRIAL_FOLDER} 아래 폴더 3개에만 넣습니다.")
    t0 = time.time()
    ns = w32.Dispatch("Outlook.Application").GetNamespace("MAPI")
    root = subfolder(ns.GetDefaultFolder(6), TRIAL_FOLDER)               # 6 = olFolderInbox
    folders = {m: subfolder(root, name) for m, name in METHODS}
    print(f"[INFO] Outlook 연결 {time.time() - t0:.1f}초")
    rows, seconds, errors_all = [], {m: [] for m, _ in METHODS}, []
    with tempfile.TemporaryDirectory() as tmp:
        for n, path in enumerate(files, 1):
            try:
                info = parse_eml(path)
            except Exception as e:                           # noqa: BLE001
                rows.append((n, "eml 읽기", {"오류": f"{type(e).__name__}"}))
                continue
            for m, _ in METHODS:
                errors = []
                start = time.time()
                try:
                    if m == "M1":
                        item = import_m1(ns, path, folders[m], info, errors)
                    elif m == "M2":
                        item = import_m2(ns, path, folders[m], info, errors)
                    else:
                        item = import_m3(ns, path, folders[m], info, errors, tmp)
                    seconds[m].append(time.time() - start)
                    res = check(item, info)
                except Exception as e:                       # noqa: BLE001
                    res = {"오류": f"{type(e).__name__}: {str(e)[:120]}"}
                if errors:
                    res["설정 못 한 속성"] = ", ".join(errors)
                    errors_all += [f"{m} {x}" for x in errors]
                rows.append((n, m, res))
                print(f"[PROGRESS] {n}/{len(files)} {m}: " + ", ".join(f"{k}={v}" for k, v in res.items() if k != "MessageClass"))
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out = os.path.join(HERE, f"eml_import_trial_{stamp}.md")
    keys = ["보낸 메일로 보임(Sent)", "받은 시각 = eml 날짜", "제목 같음", "보낸 사람 주소 같음", "첨부 수 같음", "읽음 표시", "MessageClass"]
    lines = [f"# EML -> Outlook 가져오기 시험 ({stamp})", "",
             f"- eml {len(files)}통, Python {sys.version.split()[0]}, 시험 폴더: 받은편지함\\{TRIAL_FOLDER}", "",
             "## 방법별 평균 시간 (한 통)", "", "| 방법 | 평균 초 | 성공 수 |", "|---|---|---|"]
    for m, name in METHODS:
        s = seconds[m]
        lines.append(f"| {name} | {sum(s) / len(s):.2f} | {len(s)} |" if s else f"| {name} | - | 0 |")
    lines += ["", "## 메일마다 (내용은 적지 않음)", "", "| 메일 | 방법 | " + " | ".join(keys) + " | 비고 |",
              "|" + "---|" * (len(keys) + 3)]
    for n, m, res in rows:
        note = "; ".join(f"{k}: {v}" for k, v in res.items() if k not in keys)
        lines.append(f"| {n} | {m} | " + " | ".join(str(res.get(k, "-")) for k in keys) + f" | {note} |")
    lines += ["", "## 눈으로 확인할 것 (Outlook에서 세 폴더의 메일을 열어 보고 적어 주세요)", "",
              "| 확인 | M1_그대로 | M2_속성 | M3_게시물 |", "|---|---|---|---|",
              "| '보내지 않은 메시지' 같은 안내 띠가 없다 | | | |", "| 목록의 날짜가 원래 메일 날짜다 | | | |",
              "| 보낸 사람이 원래 사람으로 보인다 | | | |", "| 한글 제목·본문이 깨지지 않는다 | | | |",
              "| 첨부를 열 수 있다 | | | |", "| Outlook 검색에 나온다 | | | |", "",
              "시험이 끝나면 Outlook에서 받은편지함\\eml 시험 폴더를 지우면 됩니다."]
    with open(out, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")
    print(f"[STATUS] 끝 ({time.time() - t0:.0f}초). 결과 파일: {out}")
    print("[STATUS] Outlook에서 받은편지함\\eml 시험 아래 세 폴더의 메일을 열어 보고 결과 파일 끝의 표를 채워 주세요.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
