# -*- coding: utf-8 -*-
"""tools/eml_import_trial.py (EML -> Outlook 가져오기 시험)의 eml 읽기 부분. Outlook에 넣는 부분은 Windows 실기에서만."""
import importlib.util
import os
from datetime import datetime

HERE = os.path.dirname(os.path.abspath(__file__))
_spec = importlib.util.spec_from_file_location("eml_import_trial", os.path.join(HERE, "..", "tools", "eml_import_trial.py"))
trial = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(trial)


def _eml(tmp_path, name, data):
    p = tmp_path / name
    p.write_bytes(data)
    return str(p)


def test_korean_headers_html_body_and_attachments(tmp_path):
    raw = ("From: =?UTF-8?B?7ZmN6ri464+Z?= <hong@example.com>\r\n"
           "To: kim@example.com\r\nSubject: =?UTF-8?B?7Iuc7ZeYIOqysOqzvA==?=\r\n"
           "Date: Mon, 09 Nov 2026 10:30:00 +0900\r\nMessage-ID: <abc@example.com>\r\nMIME-Version: 1.0\r\n"
           "Content-Type: multipart/mixed; boundary=XX\r\n\r\n"
           "--XX\r\nContent-Type: text/html; charset=utf-8\r\n\r\n<p>본문 <b>입니다</b></p>\r\n"
           "--XX\r\nContent-Type: application/pdf\r\nContent-Disposition: attachment; filename=\"=?UTF-8?B?6rec6rKp7IScLnBkZg==?=\"\r\n"
           "Content-Transfer-Encoding: base64\r\n\r\nJVBERi0x\r\n"
           "--XX\r\nContent-Type: message/rfc822\r\n\r\nSubject: inner\r\n\r\nfwd body\r\n--XX--\r\n").encode("utf-8")
    info = trial.parse_eml(_eml(tmp_path, "a.eml", raw))
    assert info["subject"] == "시험 결과" and info["from_name"] == "홍길동" and info["from_addr"] == "hong@example.com"
    assert info["date"] == datetime(2026, 11, 9, 1, 30)                       # UTC
    assert "<b>입니다</b>" in info["html"] and info["message_id"] == "<abc@example.com>"
    names = [n for n, _ in info["attachments"]]
    assert names == ["규격서.pdf", "첨부된 메일.eml"] and info["attachments"][0][1] == b"%PDF-1"
    assert "fwd body" not in info["html"] + info["text"]                       # 전달된 메일은 본문에 섞지 않음


def test_raw_cp949_header_and_plain_body(tmp_path):
    raw = "Subject: 회의 안내\r\nFrom: 담당자 <a@b.kr>\r\nContent-Type: text/plain; charset=euc-kr\r\n\r\n내일 10시\r\n".encode("cp949")
    info = trial.parse_eml(_eml(tmp_path, "b.eml", raw))
    assert info["subject"] == "회의 안내" and info["from_addr"] == "a@b.kr"
    assert info["text"].strip() == "내일 10시" and info["date"] is None and info["attachments"] == []


def test_read_trial_writes_a_result_without_mail_contents(tmp_path, monkeypatch):
    """eml_read_trial.py: 폴더의 eml을 읽고 첨부(.txt) 안의 글까지 office-cli reader로 읽음, 검색은 어디서 맞았는지만.
    결과 파일에는 제목·주소·본문을 쓰지 않음."""
    _spec2 = importlib.util.spec_from_file_location("eml_read_trial", os.path.join(HERE, "..", "tools", "eml_read_trial.py"))
    rd = importlib.util.module_from_spec(_spec2)
    _spec2.loader.exec_module(rd)
    folder = tmp_path / "mails"
    folder.mkdir()
    (folder / "1.eml").write_bytes(("Subject: =?UTF-8?B?7Iuc7ZeYIOqysOqzvA==?=\r\nFrom: a@b.kr\r\nTo: c@d.kr\r\n"
                                    "Date: Mon, 09 Nov 2026 10:30:00 +0900\r\nContent-Type: multipart/mixed; boundary=B\r\n\r\n"
                                    "--B\r\nContent-Type: text/plain; charset=utf-8\r\n\r\n비밀 과제 일정\r\n"
                                    "--B\r\nContent-Type: text/plain; charset=utf-8\r\nContent-Disposition: attachment; filename=note.txt\r\n\r\n"
                                    "첨부 안의 글\r\n--B--\r\n").encode("utf-8"))
    (folder / "2.eml").write_bytes("Subject: 회의\r\nFrom: x@y.kr\r\n\r\n내용 없음\r\n".encode("cp949"))
    monkeypatch.setattr(rd, "HERE", str(tmp_path))
    assert rd.main(["--folder", str(folder), "--query", "과제;일정"]) == 0
    out = next(tmp_path.glob("eml_read_trial_*.md")).read_text(encoding="utf-8")
    assert "| 1 | " in out and "OK 8자 (글)" in out and ".txt:7자" in out and "| 본문 |" in out
    assert "맞은 메일 1통" in out and "받는 사람" in out
    for secret in ("비밀", "시험 결과", "a@b.kr", "회의"):
        assert secret not in out, secret
