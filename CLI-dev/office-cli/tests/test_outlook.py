"""outlook.py 테스트: 검색 필터, 발송 정책(기본 표시/명시 발송/수신자 확인), 첨부 저장 규칙, 폴더 해석."""
import json
import os

import pytest

import _runtime as rt
import outlook


class Recip:
    def __init__(self, name, ok):
        self.Name = name
        self.Resolved = ok


class Recips:
    def __init__(self, items):
        self.items = items

    @property
    def Count(self):
        return len(self.items)

    def Item(self, i):
        return self.items[i - 1]

    def ResolveAll(self):
        return all(r.Resolved for r in self.items)


class Attachments:
    def __init__(self, items=()):
        self.items = list(items)
        self.added = []

    @property
    def Count(self):
        return len(self.items)

    def Item(self, i):
        return self.items[i - 1]

    def Add(self, path):
        self.added.append(path)
        self.items.append(Att(os.path.basename(path)))      # Outlook처럼 Count가 늘어남 (1.4.34부터 다시 셈)


class Mail:
    def __init__(self, recips=(), html="<html><body><p>서명</p></body></html>"):
        self.To = ""
        self.Subject = ""
        self.HTMLBody = html
        self.Recipients = Recips(list(recips))
        self.Attachments = Attachments()
        self.sent = self.displayed = False
        self.GetInspector = object()

    def Send(self):
        self.sent = True

    def Display(self):
        self.displayed = True


class App:
    def __init__(self, mail):
        self.mail = mail

    def CreateItem(self, kind):
        return self.mail


def run(capsys, monkeypatch, mail, *argv):
    monkeypatch.setattr(outlook, "get_outlook", lambda: App(mail))
    with pytest.raises(SystemExit) as ei:
        outlook.main(list(argv))
    lines = [ln for ln in capsys.readouterr().out.splitlines() if ln.strip()]
    assert len(lines) == 1
    return ei.value.code, json.loads(lines[0])


def test_default_is_display_only(capsys, monkeypatch):
    mail = Mail([Recip("a@x.com", True)])
    code, out = run(capsys, monkeypatch, mail, "new", "--to", "a@x.com", "--subject", "보고", "--body", "첫줄\\n둘째줄")
    assert code == 0 and mail.displayed and not mail.sent and out["sent"] is False
    assert "대신 발송하지 마세요" in out["agent_directive"]
    body_pos = mail.HTMLBody.index("첫줄<br>둘째줄")
    assert mail.HTMLBody.index("<body>") < body_pos < mail.HTMLBody.index("서명")     # 서명 위, body 안


def test_explicit_send_after_recipient_check(capsys, monkeypatch):
    mail = Mail([Recip("a@x.com", True)])
    code, out = run(capsys, monkeypatch, mail, "new", "--to", "a@x.com", "--subject", "s", "--send")
    assert code == 0 and mail.sent and out["sent"] is True


def test_unresolved_recipient_is_not_sent(capsys, monkeypatch):
    mail = Mail([Recip("오타주소", False)])
    code, out = run(capsys, monkeypatch, mail, "new", "--to", "오타주소", "--send")
    assert code == 1 and not mail.sent and mail.displayed
    assert out["unresolved_recipients"] == ["오타주소"] and out["agent_directive"].startswith("STOP_RETRY")


def test_attach_resolves_data_dir_and_missing_is_suggested(capsys, monkeypatch):
    data = rt.get_data_dir()
    open(os.path.join(data, "보고서_최종.pdf"), "w").close()
    mail = Mail()
    code, out = run(capsys, monkeypatch, mail, "new", "--attach", "보고서_최종.pdf")
    assert code == 0 and mail.Attachments.added == [os.path.join(data, "보고서_최종.pdf")]
    code, out = run(capsys, monkeypatch, Mail(), "new", "--attach", "보고서_최종.docx")
    assert code == 1 and "[후보] 보고서_최종.pdf" in out["message"] and "[추천]" not in out["message"]   # 다른 파일을 첨부해 발송하지 않게


def test_search_filter_has_single_prefix_and_grouping():
    f = outlook.build_search_filter("견적's", "김팀장", "all")
    assert f.startswith("@SQL=(") and f.count("@SQL=") == 1
    assert ") AND (" in f and "견적''s" in f
    assert outlook.build_search_filter("a%b_c", None, "subject") == '@SQL=("urn:schemas:httpmail:subject" LIKE \'%a%b_c%\')'
    with pytest.raises(rt.InputError):
        outlook.build_search_filter(None, None, "all")


class Att:
    def __init__(self, name, hidden=False, fail=False):
        self.FileName = name
        self.hidden = hidden
        self.fail = fail
        self.PropertyAccessor = self

    def GetProperty(self, tag):
        return self.hidden

    def SaveAsFile(self, path):
        if self.fail:
            raise OSError("잠김")
        open(path, "w").close()


def test_save_attachments_rules(tmp_path):
    item = Mail()
    item.Subject = "견적/회신"
    item.Attachments = Attachments([Att("a.xlsx"), Att("a.xlsx"), Att("logo.png", hidden=True),
                                    Att("setup.exe"), Att("b.pdf", fail=True)])
    infos = outlook._attachment_infos(item)
    saved, skipped, problems = outlook._save_selected(infos, None, str(tmp_path / "out"))
    assert [os.path.basename(s["path"]) for s in saved] == ["a.xlsx", "a (2).xlsx"]          # 같은 이름은 번호 붙여 보존
    reasons = [f"{s['name']} {s['reason']}" for s in skipped]
    assert any("본문 삽입" in s for s in reasons) and any("setup.exe" in s for s in reasons)
    assert any("b.pdf" in s and "실패" in s for s in reasons) and problems == ["b.pdf 저장 실패"]


class Folder:
    def __init__(self, name, subs=()):
        self.Name = name
        self.Folders = list(subs)
        self.Parent = None


class NS:
    def __init__(self):
        proj = Folder("프로젝트", [Folder("2026")])
        self.inbox = Folder("받은 편지함", [proj])
        self.root = Folder("me@company.com", [self.inbox])
        self.inbox.Parent = self.root
        self.Folders = [self.root]

    def GetDefaultFolder(self, n):
        return self.inbox


def test_resolve_folder_paths_and_suggestion():
    ns = NS()
    assert outlook.resolve_folder(ns, None) is ns.inbox
    assert outlook.resolve_folder(ns, "me@company.com\\받은 편지함\\프로젝트\\2026").Name == "2026"
    assert outlook.resolve_folder(ns, "받은 편지함 > 프로젝트").Name == "프로젝트"
    with pytest.raises(rt.InputError) as ei:
        outlook.resolve_folder(ns, "받은 편지함\\프로젝")
    assert "[후보] 받은 편지함\\프로젝트" in str(ei.value) and "[추천]" not in str(ei.value)   # 후보도 전체 경로
    assert outlook.resolve_folder(ns, "inbox\\프로젝트").Name == "프로젝트"                   # 기본 폴더 별칭으로 시작
