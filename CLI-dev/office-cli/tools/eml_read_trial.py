# -*- coding: utf-8 -*-
"""EML 읽기 시험 (사람이 실행하는 시험 스크립트, 에이전트 출력 계약 대상 아님)

목적: Outlook 없이 .eml 폴더를 바로 읽어 검색용 캐시(office-cli reader)로 쓸 수 있는지 확인합니다.
메일마다 아래를 봅니다.
  1. 헤더: 제목·보낸 사람·받는 사람·날짜를 읽었는지, 한글이 깨지지 않았는지 (RFC2047, 인코딩 표시 없는 CP949 헤더)
  2. 본문: 글(text/plain) 또는 HTML에서 뽑은 글이 있는지, 깨진 글자가 있는지
  3. 첨부: 이름을 읽었는지, PDF·워드·엑셀·한글 등 첨부 안의 글을 office-cli reader로 읽을 수 있는지
  4. 전달된 메일(메일 안의 메일)이 본문에 섞이지 않는지
  5. (--query) 검색어가 어느 메일의 어디(헤더, 본문, 첨부)에서 나오는지

사용법 (C:\\CLI-dev\\office-cli\\tools, 옆 폴더 C:\\CLI\\office-cli의 reader를 씀):
    python eml_read_trial.py --folder "<eml 폴더>"                       :: 폴더의 eml 전부 (하위 폴더 제외, 최대 --limit 50)
    python eml_read_trial.py --folder "<eml 폴더>" --query "과제;일정"      :: 검색도 (; 그리고, | 또는)
    python eml_read_trial.py --folder "<eml 폴더>" --show                 :: 화면에 제목·보낸 사람·본문 앞부분도 보여 줌 (결과 파일에는 안 씀)

결과는 이 파일 옆의 eml_read_trial_<시각>.md 하나입니다. 메일 내용(제목, 주소, 본문)은 적지 않고 OK/FAIL, 글자 수, 형식,
걸린 시간만 적으므로 그대로 올려도 됩니다. Outlook과 원본 eml은 건드리지 않습니다.
"""
import argparse
import os
import re
import sys
import tempfile
import time
from datetime import datetime

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import eml_import_trial as eml  # noqa: E402  같은 eml 읽기 코드 (Outlook 가져오기 시험과 같음)

OFFICE_CLI = os.path.normpath(os.path.join(HERE, "..", "..", "..", "CLI", "office-cli"))   # C:\CLI-dev\..\CLI\office-cli
BROKEN_RE = re.compile("[\ufffd]|[\u00c0-\u00ff]{2,}")     # 대체 문자, 연속된 라틴 확장 글자(UTF-8을 다른 인코딩으로 읽은 흔적)


def broken(text):
    return bool(text) and bool(BROKEN_RE.search(text))


def body_text(info):
    return info["text"].strip() or (eml.strip_html(info["html"]) if info["html"] else "")


def attachment_texts(info, tmp, reader):
    """첨부마다 (형식, 읽은 글자 수 또는 이유). reader가 읽는 형식만 읽음."""
    out = []
    for k, (name, data) in enumerate(info["attachments"]):
        ext = os.path.splitext(name)[1].lower()
        if reader is None:
            out.append((ext or "-", "reader 없음"))
            continue
        if ext not in reader.SUPPORTED:
            out.append((ext or "-", "읽지 않는 형식"))
            continue
        path = os.path.join(tmp, f"{k + 1}{ext}")
        with open(path, "wb") as f:
            f.write(data)
        try:
            units, _ = reader.PARSERS[reader.file_type(path)](path)
            out.append((ext, sum(len(u.get("text", "")) for u in units)))
        except Exception as e:                               # noqa: BLE001  시험 기록
            out.append((ext, f"실패 {type(e).__name__}"))
    return out


def match(query, parts):
    """--query (;=그리고, |=또는): 맞은 곳 이름 목록 (헤더/본문/첨부). 모든 '그리고' 조각이 어딘가에 있어야 맞음."""
    if not query:
        return []
    terms = [[t.strip().lower() for t in g.split("|") if t.strip()] for g in query.split(";") if g.strip()]
    where = set()
    for alts in terms:
        hit = [name for name, text in parts.items() if any(t in text.lower() for t in alts)]
        if not hit:
            return []
        where.update(hit)
    return sorted(where)


def main(argv=None):
    p = argparse.ArgumentParser(description="EML 읽기 시험 (Outlook 없이, 결과 파일에는 메일 내용을 쓰지 않음)")
    p.add_argument("--folder", required=True, help="eml 파일이 든 폴더")
    p.add_argument("--limit", type=int, default=50, help="읽을 eml 수 상한 (기본 50)")
    p.add_argument("--query", help="검색어 (;는 그리고, |는 또는)")
    p.add_argument("--show", action="store_true", help="화면에 제목·보낸 사람·본문 앞부분 보이기 (결과 파일에는 쓰지 않음)")
    args = p.parse_args(argv)
    files = sorted(os.path.join(args.folder, f) for f in os.listdir(args.folder) if f.lower().endswith(".eml"))[:args.limit]
    if not files:
        print(f"[ERROR] {args.folder}에 .eml 파일이 없습니다.")
        return 1
    reader = None
    if os.path.isdir(OFFICE_CLI):
        sys.path.insert(0, OFFICE_CLI)
        try:
            import reader as reader_mod                     # C:\CLI\office-cli\reader.py (첨부 안의 글 읽기)
            reader = reader_mod
        except Exception as e:                               # noqa: BLE001
            print(f"[WARN] office-cli reader를 불러오지 못해 첨부 안의 글은 읽지 않습니다: {type(e).__name__}: {e}")
    else:
        print(f"[WARN] {OFFICE_CLI}가 없어 첨부 안의 글은 읽지 않습니다.")
    print(f"[INFO] eml {len(files)}통을 읽습니다.")
    rows, total, hits = [], time.time(), 0
    with tempfile.TemporaryDirectory() as tmp:
        for n, path in enumerate(files, 1):
            start = time.time()
            try:
                info = eml.parse_eml(path)
            except Exception as e:                           # noqa: BLE001
                rows.append({"n": n, "오류": f"{type(e).__name__}: {str(e)[:100]}"})
                print(f"[PROGRESS] {n}/{len(files)} 읽기 실패: {type(e).__name__}")
                continue
            body = body_text(info)
            atts = attachment_texts(info, tmp, reader)
            sec = time.time() - start
            header = " ".join([info["subject"], info["from_name"], info["from_addr"], info["to"], info["cc"]])
            row = {
                "n": n, "크기KB": round(os.path.getsize(path) / 1024),
                "제목": "OK" if info["subject"] and not broken(info["subject"]) else ("없음" if not info["subject"] else "깨짐"),
                "보낸 사람": "OK" if info["from_addr"] and not broken(info["from_name"]) else ("없음" if not info["from_addr"] else "깨짐"),
                "받는 사람": "OK" if info["to"] and not broken(info["to"]) else ("없음" if not info["to"] else "깨짐"),
                "날짜": "OK" if info["date"] else "없음",
                "본문": ("없음" if not body else "깨짐" if broken(body) else
                         f"OK {len(body)}자 ({'글' if info['text'].strip() else 'HTML'})"),
                "첨부": ", ".join(f"{ext}:{v}" + ("자" if isinstance(v, int) else "") for ext, v in atts) or "없음",
                "전달 메일 분리": "해당 없음" if not any(nm == "첨부된 메일.eml" for nm, _ in info["attachments"]) else "OK",
                "초": round(sec, 2),
            }
            if args.query:
                parts = {"헤더": header, "본문": body}
                where = match(args.query, parts)
                row["검색"] = ", ".join(where) if where else "-"
                hits += bool(where)
            rows.append(row)
            print(f"[PROGRESS] {n}/{len(files)} " + ", ".join(f"{k}={v}" for k, v in row.items() if k not in ("n",)))
            if args.show:
                print(f"    제목: {info['subject']}\n    보낸 사람: {info['from_name']} <{info['from_addr']}>\n"
                      f"    날짜(UTC): {info['date']}\n    본문 앞부분: {body[:200]!r}")
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out = os.path.join(HERE, f"eml_read_trial_{stamp}.md")
    keys = ["크기KB", "제목", "보낸 사람", "받는 사람", "날짜", "본문", "첨부", "전달 메일 분리", "초"] + (["검색"] if args.query else [])
    lines = [f"# EML 읽기 시험 ({stamp})", "",
             f"- eml {len(files)}통, 전체 {time.time() - total:.1f}초, Python {sys.version.split()[0]}, "
             f"첨부 읽기: {'office-cli reader' if reader else '안 함'}",
             (f"- 검색어 지정함 (내용은 적지 않음): 맞은 메일 {hits}통" if args.query else "- 검색 없음"), "",
             "| 메일 | " + " | ".join(keys) + " | 오류 |", "|" + "---|" * (len(keys) + 2)]
    for r in rows:
        lines.append(f"| {r['n']} | " + " | ".join(str(r.get(k, "-")) for k in keys) + f" | {r.get('오류', '')} |")
    lines += ["", "## 눈으로 확인할 것 (--show로 화면에 나온 내용과 Outlook·메일 프로그램에서 본 원래 메일 비교)", "",
              "| 확인 | 결과 |", "|---|---|", "| 한글 제목·보낸 사람이 원래와 같다 | |", "| 본문 앞부분이 원래 본문과 같다 (서명·광고 글만 나오지 않는다) | |",
              "| 첨부 글자 수가 0인 첨부가 실제로 글이 없는 파일(그림, 스캔)이다 | |", "| 검색 결과가 예상한 메일과 같다 | |"]
    with open(out, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")
    print(f"[STATUS] 끝. 결과 파일: {out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
