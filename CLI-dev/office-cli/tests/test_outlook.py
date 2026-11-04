"""outlook.py: list/search(필터·기간), read/current, attachments, new/reply/forward(발송 정책·첨부), 폴더 해석."""
import datetime
import datetime as dt
import json
import os
import re
import types

import pytest

import _runtime as rt
import outlook
from office_helpers import error_log_text, run_fresh, run_main, run_outlook, SECRET


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


class _DraftMail:
    def __init__(self):
        self.To = ""
        self.Subject = ""
        self.HTMLBody = "<html><body><p>서명</p></body></html>"
        self.Attachments = type("A", (), {"added": [], "Add": lambda s, p: s.added.append(p), "Count": 0})()
        self.GetInspector = object()
        self.displayed = False

    def Display(self):
        self.displayed = True


def _outlook_create(monkeypatch, capsys, mail, *argv):
    monkeypatch.setattr(outlook, "get_outlook", lambda: type("App", (), {"CreateItem": lambda s, k: mail})())
    return run_main(outlook, capsys, *argv)


class _Item:
    Subject = "회의 안내"
    Body = "본문"
    To = "a"
    CC = ""
    SenderName = "김"
    SenderEmailAddress = ""
    EntryID = "1"
    UnRead = False
    Attachments = type("A", (), {"Count": 0})()


class _Att:
    def __init__(self, name, fail=False, vanish=False):
        self.FileName = name
        self.fail = fail
        self.vanish = vanish
        self.PropertyAccessor = self

    def GetProperty(self, tag):
        return False

    def SaveAsFile(self, path):
        if self.fail:
            raise OSError("잠김")
        if not self.vanish:                                  # 백신이 바로 지운 경우: 오류 없이 파일이 없음
            open(path, "w").close()


class _Atts:
    def __init__(self, items):
        self.items = items

    @property
    def Count(self):
        return len(self.items)

    def Item(self, i):
        return self.items[i - 1]


def _mail(atts):
    return types.SimpleNamespace(Subject="견적", Body="본문", SenderName="a", SenderEmailAddress="a@x", To="", CC="",
                                 ReceivedTime=datetime.datetime(2026, 1, 1), EntryID="X", UnRead=False,
                                 Attachments=_Atts(atts))


class _NoAtts:
    Count = 0


class _SearchItem:
    Class = 43

    def __init__(self, i, subject, body="", sender="김팀장"):
        self.EntryID = f"id{i}"
        self.Subject = subject
        self.Body = body
        self.SenderName = sender
        self.SenderEmailAddress = "kim@x.com"
        self.ReceivedTime = "2026-11-01"
        self.UnRead = False
        self.Attachments = _NoAtts()


class _ListItems(list):
    def Restrict(self, flt):
        self.filter = flt
        return self                    # DASL 와일드카드처럼 넓게 돌려줌 (걸러내기는 도구가)

    def Sort(self, *a):
        pass


class _ListFolder:
    Name = "받은 편지함"

    def __init__(self, items):
        self.Items = items


def _outlook_items(monkeypatch, items):
    folder = _ListFolder(_ListItems(items))
    monkeypatch.setattr(outlook, "get_outlook", lambda: object())
    monkeypatch.setattr(outlook, "get_namespace", lambda app: object())
    monkeypatch.setattr(outlook, "resolve_folder", lambda ns, name: folder)
    return folder


class _OneMail:
    Class = 43
    EntryID = "id1"
    Subject = "견적 요청"
    SenderName = "홍길동"
    SenderEmailAddress = "hong@example.com"
    ReceivedTime = "2026-11-03 10:00"
    UnRead = True
    Body = "본문"


class _MailItems:
    def __init__(self, mails):
        self.mails = list(mails)
        self.filters = []

    def Restrict(self, flt):
        self.filters.append(flt)
        return self

    def Sort(self, *a):
        pass

    def __iter__(self):
        return iter(self.mails)


class _MailFolder:
    Name = "받은 편지함"

    def __init__(self, mails):
        self.Items = _MailItems(mails)


def _outlook_mails(monkeypatch, mails):
    folder = _MailFolder(mails)
    monkeypatch.setattr(outlook, "get_outlook", lambda: object())
    monkeypatch.setattr(outlook, "get_namespace", lambda o: object())
    monkeypatch.setattr(outlook, "resolve_folder", lambda ns, f: folder)
    return folder


ROOT_MAIL = "ZQMhong@corp.com"           # 메일 계정 이름 (보통 메일 주소)


class _F:
    def __init__(self, name, subs=()):
        self.Name = name
        self.Folders = list(subs)
        self.Parent = None
        for f in self.Folders:
            f.Parent = self


class _NS:
    def __init__(self):
        self.root = _F(ROOT_MAIL, [_F("ZQX받은편지함"), _F("ZQX프로젝트A", [_F("ZQX회의록")]), _F("ZQX보관")])
        self.Folders = [self.root]

    def GetDefaultFolder(self, n):
        return self.root.Folders[0]

    def Logon(self, *a):
        pass


@pytest.fixture
def mail(monkeypatch):
    app = type("A", (), {"GetNamespace": lambda self, k: _NS()})()
    monkeypatch.setattr(outlook, "get_outlook", lambda: app)


def _result_lines(d):
    return [ln for ln in error_log_text(d).splitlines() if ln.startswith("결과:")]


class FakeAttachment:
    def __init__(self, name, size=100, hidden=False, fail=False, vanish=False):
        self.FileName = name
        self.Size = size
        self.hidden = hidden
        self.fail = fail
        self.vanish = vanish
        self.PropertyAccessor = self
        self.saved_to = []

    def GetProperty(self, tag):
        assert tag == outlook.PR_ATTACHMENT_HIDDEN
        return self.hidden

    def SaveAsFile(self, path):
        if self.fail:
            raise OSError("잠김")
        self.saved_to.append(path)
        if not self.vanish:                                  # vanish: 백신이 바로 지운 경우 (오류 없이 파일이 없음)
            with open(path, "wb") as f:
                f.write(b"data-" + self.FileName.encode("utf-8"))


class FakeAttachments:
    def __init__(self, items=(), grow=True):
        self.items = list(items)
        self.added = []
        self.grow = grow

    @property
    def Count(self):
        return len(self.items)

    def Item(self, i):
        return self.items[i - 1]

    def Add(self, path):
        self.added.append(path)
        if self.grow:                                        # grow=False: Outlook이 조용히 붙이지 못한 경우
            self.items.append(FakeAttachment(os.path.basename(path)))


class FakeRecipients:
    Count = 1

    def ResolveAll(self):
        return True


class FakeDraft:
    def __init__(self, atts=(), grow=True):
        self.To = ""
        self.Subject = ""
        self.HTMLBody = "<html><body><p>서명</p></body></html>"
        self.Attachments = FakeAttachments(atts, grow)
        self.Recipients = FakeRecipients()
        self.GetInspector = object()
        self.displayed = self.sent = False

    def Display(self):
        self.displayed = True

    def Send(self):
        self.sent = True


class FakeMail:
    def __init__(self, atts=(), subject="RE: 견적/회신", entry="M1", grow=True):
        self.EntryID = entry
        self.Subject = subject
        self.Body = "본문입니다"
        self.SenderName = "김팀장"
        self.SenderEmailAddress = "lead@corp.com"
        self.To = "me@corp.com"
        self.CC = ""
        self.UnRead = False
        self.Class = outlook.OL_MAIL
        self.ReceivedTime = datetime.datetime(2026, 11, 4, 9, 0)
        self.Attachments = FakeAttachments(atts)
        self.grow = grow
        self.drafts = []

    def _draft(self, atts=()):
        d = FakeDraft(atts, self.grow)
        self.drafts.append(d)
        return d

    def Reply(self):
        return self._draft()

    def ReplyAll(self):
        return self._draft()

    def Forward(self):                                       # 전달은 원본 첨부를 가지고 시작
        return self._draft([FakeAttachment(a.FileName) for a in self.Attachments.items])


class FakeSelection:
    def __init__(self, item):
        self.item = item
        self.Count = 1 if item is not None else 0

    def Item(self, i):
        return self.item


class FakeOutlookApp:
    def __init__(self, mail=None, grow=True):
        self.mail = mail
        self.created = []
        self.grow = grow

    def ActiveInspector(self):
        return None

    def ActiveExplorer(self):
        return type("E", (), {"Selection": FakeSelection(self.mail)})()

    def CreateItem(self, kind):
        d = FakeDraft(grow=self.grow)
        self.created.append(d)
        return d


@pytest.fixture
def ol_env(monkeypatch, tmp_path):
    """Outlook 대역과 사용자 폴더(USERPROFILE). 다운로드 폴더 레지스트리는 읽지 않음."""
    home = tmp_path / "home"
    home.mkdir()
    monkeypatch.setenv("USERPROFILE", str(home))
    monkeypatch.setattr(outlook, "_shell_downloads", lambda: None)
    monkeypatch.setattr(rt, "RECOVERED", [])
    monkeypatch.setattr(rt, "CURRENT_ARGV", None)
    state = {"home": home, "app": FakeOutlookApp(), "calls": []}

    def get_outlook():
        state["calls"].append("get_outlook")
        return state["app"]
    monkeypatch.setattr(outlook, "get_outlook", get_outlook)

    def get_namespace(app):
        def find(i):
            if state["app"].mail is not None and i == state["app"].mail.EntryID:
                return state["app"].mail
            raise RuntimeError("없음")
        return type("NS", (), {"GetItemFromID": staticmethod(find)})()
    monkeypatch.setattr(outlook, "get_namespace", get_namespace)
    return state


def _attach_mail(ol_env, *atts, **kw):
    ol_env["app"].mail = FakeMail(atts, **kw)
    return ol_env["app"].mail


def _default_root(ol_env):
    return os.path.join(str(ol_env["home"]), "Downloads", "메일 첨부")


class _DatedMail:
    Class = 43
    SenderName = "홍길동"
    SenderEmailAddress = "hong@example.com"
    Body = "본문"

    def __init__(self, eid, subject, received, unread=False):
        self.EntryID = eid
        self.Subject = subject
        self.ReceivedTime = received
        self.UnRead = unread


class _DatedItems:
    """Restrict는 조건 글을 기록하고, 안 읽은 메일 조건만 실제로 거름 (날짜는 도구가 받은 시각으로 다시 거르는지 확인).
    refuse: 날짜 조건을 거부하는 Outlook."""

    def __init__(self, mails, log, refuse=False):
        self.mails = list(mails)
        self.log = log
        self.refuse = refuse

    def Restrict(self, flt):
        self.log.append(flt)
        if self.refuse and ("ReceivedTime" in flt or "datereceived" in flt):
            raise Exception("Outlook: 조건을 해석할 수 없음")
        keep = [m for m in self.mails if m.UnRead] if flt == "[UnRead] = True" else self.mails
        return _DatedItems(keep, self.log, self.refuse)

    def Sort(self, *a):
        self.mails.sort(key=lambda m: m.ReceivedTime, reverse=True)

    def __iter__(self):
        return iter(self.mails)


TODAY = dt.date(2026, 11, 4)


def _mails():
    return [_DatedMail("c", "지난주 견적", dt.datetime(2026, 10, 28, 8, 0), unread=True),
            _DatedMail("a", "오늘 견적", dt.datetime(2026, 11, 4, 9, 0), unread=True),
            _DatedMail("d", "지난달 견적", dt.datetime(2026, 10, 1, 8, 0)),
            _DatedMail("b", "어제 회의", dt.datetime(2026, 11, 3, 23, 59))]


@pytest.fixture
def ol(monkeypatch):
    log = []
    state = {"refuse": False, "mails": _mails()}

    class _Folder:
        Name = "받은 편지함"

        @property
        def Items(self):
            return _DatedItems(state["mails"], log, state["refuse"])

    monkeypatch.setattr(outlook, "_today", lambda: TODAY)
    monkeypatch.setattr(rt, "RECOVERED", [])
    monkeypatch.setattr(outlook, "get_outlook", lambda: object())
    monkeypatch.setattr(outlook, "get_namespace", lambda o: object())
    monkeypatch.setattr(outlook, "resolve_folder", lambda ns, f: _Folder())
    return {"log": log, "state": state}


# ---------------------------------------------------------------------------
# 검색 필터, 발송 정책(기본 표시, 명시 발송, 수신자 확인), 첨부 저장 규칙, 폴더 해석
# ---------------------------------------------------------------------------


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


def test_resolve_folder_paths_and_suggestion():
    ns = NS()
    assert outlook.resolve_folder(ns, None) is ns.inbox
    assert outlook.resolve_folder(ns, "me@company.com\\받은 편지함\\프로젝트\\2026").Name == "2026"
    assert outlook.resolve_folder(ns, "받은 편지함 > 프로젝트").Name == "프로젝트"
    with pytest.raises(rt.InputError) as ei:
        outlook.resolve_folder(ns, "받은 편지함\\프로젝")
    assert "[후보] 받은 편지함\\프로젝트" in str(ei.value) and "[추천]" not in str(ei.value)   # 후보도 전체 경로
    assert outlook.resolve_folder(ns, "inbox\\프로젝트").Name == "프로젝트"                   # 기본 폴더 별칭으로 시작


# ---------------------------------------------------------------------------
# outlook 검색
# ---------------------------------------------------------------------------


def test_dasl_keeps_underscore_and_percent():
    f = outlook.build_search_filter("견적서_v2", None, "subject")
    assert "LIKE '%견적서_v2%'" in f
    f = outlook.build_search_filter("10% 할인", "o'neil", "all")
    assert "LIKE '%10% 할인%'" in f and "o''neil" in f


def test_search_wildcard_query_keeps_only_literal_matches(monkeypatch, capsys):
    folder = _outlook_items(monkeypatch, [_SearchItem(1, "견적서_v2 송부"), _SearchItem(2, "견적서-v2 송부"), _SearchItem(3, "회의", body="10% 할인 안내")])
    code, out = run_main(outlook, capsys, "search", "--query", "견적서_v2")
    assert code == 0 and "견적서_v2" in folder.Items.filter
    assert [m["subject"] for m in out["mails"]] == ["견적서_v2 송부"]      # '_' 와일드카드로 더 찾은 메일은 뺌
    code, out = run_main(outlook, capsys, "search", "--query", "10% 할인", "--where", "body")
    assert code == 0 and [m["entry_id"] for m in out["mails"]] == ["id3"]


def test_limit_upper_bound_and_zero(monkeypatch, capsys):
    _outlook_items(monkeypatch, [_SearchItem(i, f"메일 {i}") for i in range(250)])
    code, out = run_main(outlook, capsys, "list", "--limit", "1000")
    assert code == 0 and len(out["mails"]) == outlook.MAX_LIMIT and out["warnings"]
    code, out = run_main(outlook, capsys, "search", "--query", "메일", "--limit", "1000")
    assert code == 0 and len(out["mails"]) == outlook.MAX_LIMIT and out["warnings"]
    code, out = run_main(outlook, capsys, "list", "--limit", "0")
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY")


# ---------------------------------------------------------------------------
# 0건
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("argv", [["list"], ["list", "--unread"], ["search", "--query", "견적"]])
def test_outlook_zero(capsys, monkeypatch, argv):
    _outlook_mails(monkeypatch, [])
    code, out = run_main(outlook, capsys, *argv)
    assert code == 0 and out["count"] == 0 and out["mails"] == [], out
    assert out["agent_directive"].startswith("VERIFIED: 0건"), out
    assert "--folder" not in out["agent_directive"], out


@pytest.mark.parametrize("argv", [["list"], ["search", "--query", "견적"]])
def test_outlook_count(capsys, monkeypatch, argv):
    _outlook_mails(monkeypatch, [_OneMail()])
    code, out = run_main(outlook, capsys, *argv)
    assert code == 0 and out["count"] == 1 and out["mails"][0]["subject"] == "견적 요청", out


# ---------------------------------------------------------------------------
# outlook --days / --since
# ---------------------------------------------------------------------------


def test_list_days_includes_today(ol, capsys):
    code, out = run_main(outlook, capsys, "list", "--days", "2")
    assert code == 0 and [m["entry_id"] for m in out["mails"]] == ["a", "b"], out
    assert out["period"] == {"since": "2026-11-03", "until": "2026-11-04", "days": 2}, out
    assert "2026-11-03 ~ 2026-11-04" in out["preview"], out
    assert "[ReceivedTime] >= '2026-11-03 00:00'" in ol["log"], ol["log"]       # 지역 설정과 무관한 YYYY-MM-DD


def test_list_days_one_is_today_only(ol, capsys):
    code, out = run_main(outlook, capsys, "list", "--days", "1")
    assert code == 0 and [m["entry_id"] for m in out["mails"]] == ["a"], out


def test_list_since_with_unread(ol, capsys):
    code, out = run_main(outlook, capsys, "list", "--since", "2026-10-28", "--unread")
    assert code == 0 and [m["entry_id"] for m in out["mails"]] == ["a", "c"], out
    assert out["period"]["days"] == 8 and ol["log"][0] == "[UnRead] = True", (out, ol["log"])


def test_search_days_with_query_uses_one_dasl_filter(ol, capsys):
    code, out = run_main(outlook, capsys, "search", "--query", "견적", "--days", "7")
    assert code == 0 and [m["entry_id"] for m in out["mails"]] == ["a", "b"], out      # 가짜 Restrict는 글로 거르지 않음 (c는 10-28)
    flt = ol["log"][-1]
    assert flt.startswith("@SQL=") and flt.count("@SQL=") == 1 and "견적" in flt, flt
    assert '"urn:schemas:httpmail:datereceived" >= ' in flt and ") AND (" in flt, flt
    utc = dt.datetime(2026, 10, 29).astimezone(dt.timezone.utc).strftime("%Y-%m-%d %H:%M")
    assert f"'{utc}'" in flt, flt                                                 # DASL 날짜 비교는 UTC
    assert out["period"]["since"] == "2026-10-29" and "2026-10-29" in out["preview"], out


def test_search_date_only_is_allowed(ol, capsys):
    code, out = run_main(outlook, capsys, "search", "--since", "2026-11-03")
    assert code == 0 and [m["entry_id"] for m in out["mails"]] == ["a", "b"], out


def test_zero_result_states_period(ol, capsys):
    ol["state"]["mails"] = [m for m in _mails() if m.EntryID in ("c", "d")]
    for argv in (["list", "--days", "3"], ["search", "--query", "견적", "--days", "3"]):
        code, out = run_main(outlook, capsys, *argv)
        assert code == 0 and out["count"] == 0 and out["mails"] == [], out
        assert out["agent_directive"].startswith("VERIFIED: 0건") and "2026-11-02 ~ 2026-11-04" in out["agent_directive"], out


def test_restrict_refused_falls_back_and_filters(ol, capsys):
    ol["state"]["refuse"] = True
    code, out = run_main(outlook, capsys, "list", "--days", "2")
    assert code == 0 and [m["entry_id"] for m in out["mails"]] == ["a", "b"], out
    assert out["agent_directive"].startswith("VERIFIED:") and out["recovered"], out
    rt.RECOVERED.clear()
    code, out = run_main(outlook, capsys, "search", "--query", "견적", "--days", "7")
    assert code == 0 and [m["entry_id"] for m in out["mails"]] == ["a", "b"] and out["recovered"], out
    assert "datereceived" not in ol["log"][-1] and "견적" in ol["log"][-1], ol["log"]
    rt.RECOVERED.clear()


@pytest.mark.parametrize("argv, rec", [
    (["--since", "2026/10/01"], "[추천] 2026-10-01"),
    (["--since", "20261001"], "[추천] 2026-10-01"),
    (["--since", "2026년 10월 1일"], "[추천] 2026-10-01"),
    (["--since", "2026-1-5"], "[추천] 2026-01-05"),
    (["--days", "0"], "[추천] 1"),
    (["--days", "-7"], "[추천] 7"),
])
def test_bad_date_recommends(ol, capsys, argv, rec):
    for cmd in ("list", "search"):
        code, out = run_main(outlook, capsys, cmd, *argv)
        assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:") and out["message"].endswith(rec), out
    assert not ol["log"], ol["log"]          # Outlook에 묻기 전에 멈춤


@pytest.mark.parametrize("argv", [["--since", "2026-02-30"], ["--since", "지난주"], ["--since", "2026-12-01"],
                                  ["--days", "3", "--since", "2026-11-01"]])
def test_bad_date_without_guess(ol, capsys, argv):
    code, out = run_main(outlook, capsys, "list", *argv)
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:") and "[추천]" not in out["message"], out


# ---------------------------------------------------------------------------
# outlook current --save 확인, 폴더 이름에 ; 가 든 첨부 여러 개
# ---------------------------------------------------------------------------


def test_current_save_problem_is_unverified(capsys, monkeypatch):
    monkeypatch.setattr(outlook, "get_outlook", lambda: object())
    monkeypatch.setattr(outlook, "current_item", lambda app: (_Item(), "열린 창"))
    code, out = run_main(outlook, capsys, "current", "--save")
    assert code == 0 and out["agent_directive"].startswith("VERIFIED: 메일을 읽고 저장한"), out
    monkeypatch.setattr(rt, "check_saved_file", lambda path, rows=None, delimiter=None: "저장한 파일을 다시 열지 못함 (테스트)")
    code, out = run_main(outlook, capsys, "current", "--save")
    assert code == 0 and out["agent_directive"].startswith("UNVERIFIED:") and out["save_problem"]


def test_attach_paths_in_folder_with_semicolon(tmp_path):
    folder = tmp_path / "자료;2026"
    folder.mkdir()
    a, b = folder / "a.pdf", tmp_path / "b.pdf"
    a.write_bytes(b"1")
    b.write_bytes(b"2")
    assert outlook._attach_paths(f"{a};{b}") == [str(a), str(b)]
    assert outlook._attach_paths(f"{b};{a}") == [str(b), str(a)]
    with pytest.raises(rt.InputError):
        outlook._attach_paths(f"{a};{tmp_path / '없는.pdf'}")


# ---------------------------------------------------------------------------
# outlook 폴더 [후보]는 끝 이름으로 비교, 메일 계정 이름은 로그에 남기지 않음
# ---------------------------------------------------------------------------


def test_folder_candidates_match_leaf_names(mail, capsys, monkeypatch, isolated_dirs):
    code, out = run_fresh(outlook, capsys, monkeypatch, "list", "--folder", "프로젝트")
    assert code == 1 and out["candidates"] == [f"{ROOT_MAIL}\\ZQX프로젝트A"], out
    code, out = run_fresh(outlook, capsys, monkeypatch, "list", "--folder", "회의")    # 비슷한 끝 이름 없음: [후보] 아님
    assert code == 1 and "candidates" not in out and "사용 가능한 메일 폴더" in out["message"], out
    assert f"{ROOT_MAIL}\\ZQX보관" in out["message"]
    code, out = run_fresh(outlook, capsys, monkeypatch, "list", "--folder", "zqx프로젝트a\\회의")
    assert code == 1 and out["candidates"] == [f"{ROOT_MAIL}\\ZQX프로젝트A\\ZQX회의록"], out


def test_folder_same_name_recommends_full_path(mail, capsys, monkeypatch):
    code, out = run_fresh(outlook, capsys, monkeypatch, "list", "--folder", "zqx_보관")
    assert code == 1 and out["message"].endswith(f"[추천] {ROOT_MAIL}\\ZQX보관"), out


@pytest.mark.parametrize("folder", ["프로젝트", "회의", "zqx_보관", f"{ROOT_MAIL}\\없음"])
def test_guard_folder_log_has_no_mailbox_root(mail, capsys, monkeypatch, isolated_dirs, folder):
    code, out = run_fresh(outlook, capsys, monkeypatch, "list", "--folder", folder)
    assert code == 1 and ROOT_MAIL in out["message"], out
    log = error_log_text(isolated_dirs)
    assert log and ROOT_MAIL not in log.replace(f"--folder {folder}", "") and SECRET not in log, log
    assert "<메일 계정>" in log


# ---------------------------------------------------------------------------
# read / current
# ---------------------------------------------------------------------------


def test_read_shows_attachment_names_and_download_command(ol_env, capsys):
    _attach_mail(ol_env, FakeAttachment("견적서.pdf"), FakeAttachment("image001.png", hidden=True))
    code, out = run_outlook(capsys, "read", "--id", "M1")
    assert code == 0 and out["attachment_names"] == ["견적서.pdf"] and out["attachment_names_truncated"] is False
    assert out["inline_images"] == 1
    assert [c["purpose"] for c in out["next_commands"]] == ["첨부 내려받기"]
    assert out["next_commands"][0]["command"].endswith("attachments --id M1 --save")
    assert "saved_attachments" not in out and not os.path.exists(_default_root(ol_env))   # 읽기만으로 저장하지 않음


def test_read_save_puts_word_command_first(ol_env, capsys):
    _attach_mail(ol_env, FakeAttachment("a.pdf"))
    code, out = run_outlook(capsys, "read", "--id", "M1", "--save")
    assert code == 0 and [c["purpose"] for c in out["next_commands"]] == ["메일 내용을 워드에 넣기", "첨부 내려받기"]


def test_current_caps_attachment_names(ol_env, capsys):
    _attach_mail(ol_env, *[FakeAttachment(f"f{i:02d}.pdf") for i in range(25)], entry="SEL")
    code, out = run_outlook(capsys, "current")
    assert code == 0 and len(out["attachment_names"]) == 20 and out["attachment_names_truncated"] is True
    assert "attachments --id SEL --save" in out["next_commands"][0]["command"]


def test_read_without_attachments_has_no_download(ol_env, capsys):
    _attach_mail(ol_env, FakeAttachment("image001.png", hidden=True))
    code, out = run_outlook(capsys, "read", "--id", "M1")
    assert code == 0 and "attachment_names" not in out and "next_commands" not in out


def test_save_attach_option_is_gone(ol_env, capsys):
    _attach_mail(ol_env, FakeAttachment("a.pdf"))
    code, out = run_outlook(capsys, "read", "--id", "M1", "--save-attach")
    assert code == 2 and out["agent_directive"].startswith("STOP_RETRY:")


# ---------------------------------------------------------------------------
# 메일 저장 이름 정리·저장 확인, 본문 파일의 Windows 경로, 첨부 이름의 쉼표·세미콜론
# ---------------------------------------------------------------------------


def test_saved_mail_name_is_sanitized_and_verified(isolated_dirs):
    item = type("M", (), {"Subject": 'RE: 견적/회신? "최종"', "Body": "본문입니다", "To": "a", "CC": "",
                          "SenderName": "김", "SenderEmailAddress": "", "EntryID": "1", "UnRead": False,
                          "Attachments": type("A", (), {"Count": 0})()})()
    args = type("Args", (), {"save": True})()
    res = outlook._mail_detail(item, args, "id")
    name = os.path.basename(res["output_file"])
    assert not any(ch in name for ch in ':/?"') and name.startswith("mail_RE_ 견적_회신_")
    assert os.path.isfile(res["output_file"]) and "save_problem" not in res
    assert outlook._safe_name("...", 10) == "메일"


def test_body_file_keeps_windows_paths(monkeypatch, capsys, tmp_path):
    body = tmp_path / "본문.txt"
    body.write_text("자료 위치: C:\\new\\data\\n폴더", encoding="utf-8")
    mail = _DraftMail()
    code, out = _outlook_create(monkeypatch, capsys, mail, "new", "--body-file", str(body))
    assert code == 0 and "C:\\new\\data\\n폴더" in mail.HTMLBody
    mail = _DraftMail()
    code, out = _outlook_create(monkeypatch, capsys, mail, "new", "--body", "경로 C:\\new\\data.xlsx\\n다음 줄")
    assert "C:\\new\\data.xlsx<br>다음 줄" in mail.HTMLBody


def test_attach_name_with_comma_and_semicolon_list(monkeypatch, capsys):
    data = rt.get_data_dir()
    for n in ("보고서,최종.pdf", "a.pdf", "b.pdf"):
        open(os.path.join(data, n), "w").close()
    mail = _DraftMail()
    code, out = _outlook_create(monkeypatch, capsys, mail, "new", "--attach", "보고서,최종.pdf")
    assert code == 0 and out["attached"] == ["보고서,최종.pdf"]
    mail = _DraftMail()
    code, out = _outlook_create(monkeypatch, capsys, mail, "new", "--attach", "a.pdf;b.pdf")
    assert code == 0 and out["attached"] == ["a.pdf", "b.pdf"]


# ---------------------------------------------------------------------------
# 첨부 저장(1.4.34부터 attachments --save)에서 저장하지 못한 첨부가 있으면 VERIFIED가 아님
# ---------------------------------------------------------------------------


# 1.4.34: read --save-attach 대신 attachments --save (폴더는 전체 경로)
@pytest.mark.parametrize("att", [_Att("b.pdf", fail=True), _Att("c.pdf", vanish=True)])
def test_failed_attachment_save_is_unverified(monkeypatch, att, tmp_path):
    item = _mail([_Att("a.xlsx"), att])
    monkeypatch.setattr(outlook, "get_outlook", lambda: None)
    monkeypatch.setattr(outlook, "get_namespace", lambda app: types.SimpleNamespace(GetItemFromID=lambda i: item))
    out = outlook.cmd_attachments(types.SimpleNamespace(id="X", save=str(tmp_path / "첨부"), name=None, index=None))
    assert out["agent_directive"].startswith("UNVERIFIED:") and att.FileName in out["agent_directive"], out
    assert len(out["saved"]) == 1


def test_all_attachments_saved_is_verified(monkeypatch, tmp_path):
    item = _mail([_Att("a.xlsx")])
    monkeypatch.setattr(outlook, "get_outlook", lambda: None)
    monkeypatch.setattr(outlook, "get_namespace", lambda app: types.SimpleNamespace(GetItemFromID=lambda i: item))
    out = outlook.cmd_attachments(types.SimpleNamespace(id="X", save=str(tmp_path / "첨부"), name=None, index=None))
    assert out["agent_directive"].startswith("VERIFIED:") and len(out["saved"]) == 1


def test_outlook_new_checks_attachment_before_creating_mail(monkeypatch, tmp_path):
    created = []
    monkeypatch.setattr(outlook, "get_outlook", lambda: types.SimpleNamespace(CreateItem=lambda k: created.append(k)))
    args = types.SimpleNamespace(to="a@x", cc=None, subject="s", body="b", body_file=None,
                                 attach=str(tmp_path / "없는파일.pdf"), send=False)
    with pytest.raises(rt.InputError):
        outlook.cmd_new(args)
    assert created == []


# ---------------------------------------------------------------------------
# attachments: 목록
# ---------------------------------------------------------------------------


def test_list_attachments_with_index_size_inline(ol_env, capsys):
    _attach_mail(ol_env, FakeAttachment("견적서.pdf", 2048), FakeAttachment("image001.png", 500, hidden=True), FakeAttachment("setup.exe", 10))
    code, out = run_outlook(capsys, "attachments", "--id", "M1")
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:") and out["count"] == 3
    assert out["attachments"][0] == {"index": 1, "name": "견적서.pdf", "size": 2048, "hidden_inline": False,
                                     "blocked": False}
    assert out["attachments"][1]["hidden_inline"] is True and out["attachments"][2]["blocked"] is True
    assert out["truncated"] is False and "saved" not in out
    assert out["next_commands"][0]["command"].endswith("attachments --id M1 --save")
    assert not os.path.exists(_default_root(ol_env))                     # 목록만: 아무것도 저장하지 않음


def test_list_uses_selected_mail_without_id(ol_env, capsys):
    _attach_mail(ol_env, FakeAttachment("a.pdf"), entry="SEL")
    code, out = run_outlook(capsys, "attachments")
    assert code == 0 and out["count"] == 1 and out["mail"]["entry_id"] == "SEL"
    assert "attachments --id SEL --save" in out["next_commands"][0]["command"]    # 다음 명령은 id로 고정


def test_zero_attachments_is_explicit(ol_env, capsys):
    _attach_mail(ol_env)
    code, out = run_outlook(capsys, "attachments", "--id", "M1")
    assert code == 0 and out["count"] == 0 and out["attachments"] == []
    assert out["agent_directive"].startswith("VERIFIED: 0건.") and "첨부 없음" in out["agent_directive"]
    code, out = run_outlook(capsys, "attachments", "--id", "M1", "--save")
    assert code == 0 and out["count"] == 0 and out["saved"] == [] and out["folder"] is None
    assert "첨부 없음" in out["agent_directive"] and not os.path.exists(_default_root(ol_env))


def test_no_mail_selected_is_stop_retry(ol_env, capsys):
    code, out = run_outlook(capsys, "attachments")
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:") and "--id" in out["message"]


def test_name_without_save_is_refused(ol_env, capsys):
    _attach_mail(ol_env, FakeAttachment("a.pdf"))
    code, out = run_outlook(capsys, "attachments", "--id", "M1", "--name", "a")
    assert code == 1 and "--save" in out["message"]


# ---------------------------------------------------------------------------
# attachments --save
# ---------------------------------------------------------------------------


def test_save_all_to_downloads_skips_blocked_and_inline(ol_env, capsys):
    m = _attach_mail(ol_env, FakeAttachment("견적서.pdf", 2048), FakeAttachment("image001.png", 500, hidden=True), FakeAttachment("setup.exe", 10),
                     FakeAttachment("단가표.xlsx", 4096))
    code, out = run_outlook(capsys, "attachments", "--id", "M1", "--save")
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    folder = out["folder"]
    assert os.path.dirname(folder) == _default_root(ol_env)
    assert os.path.basename(folder).endswith("_RE_ 견적_회신")          # rt.safe_filename(제목)
    assert [os.path.basename(p) for p in out["saved"]] == ["견적서.pdf", "단가표.xlsx"] and out["count"] == 2
    assert all(os.path.isfile(p) and os.path.dirname(p) == folder for p in out["saved"])
    reasons = {s["name"]: s["reason"] for s in out["skipped"]}
    assert "실행 파일" in reasons["setup.exe"] and "본문 삽입" in reasons["image001.png"]
    assert folder in out["preview"] and "folder" in out["agent_directive"]
    assert m.Attachments.items[2].saved_to == []                       # 실행 파일은 저장 시도도 하지 않음
    assert "recovered" not in out


def test_save_to_user_folder_creates_it(ol_env, capsys, tmp_path):
    _attach_mail(ol_env, FakeAttachment("a.pdf"))
    target = tmp_path / "내 폴더" / "첨부"
    code, out = run_outlook(capsys, "attachments", "--id", "M1", "--save", str(target))
    assert code == 0 and out["folder"] == str(target) and os.path.isfile(target / "a.pdf")


def test_relative_folder_is_refused(ol_env, capsys, tmp_path, monkeypatch):
    m = _attach_mail(ol_env, FakeAttachment("a.pdf"))
    monkeypatch.chdir(tmp_path)
    code, out = run_outlook(capsys, "attachments", "--id", "M1", "--save", "첨부")
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:") and "전체 경로" in out["message"]
    assert m.Attachments.items[0].saved_to == [] and not (tmp_path / "첨부").exists()


@pytest.mark.parametrize("argv, expected", [
    (["--name", "견적"], ["견적서.pdf"]),
    (["--name", "PDF"], ["견적서.pdf", "보고서.pdf"]),                     # 대소문자 무시
    (["--name", "단가;보고"], ["단가표.xlsx", "보고서.pdf"]),
    (["--name", "단가", "--name", "견적"], ["견적서.pdf", "단가표.xlsx"]),   # 옵션 반복, 저장은 첨부 순서
    (["--index", "1,3"], ["견적서.pdf", "보고서.pdf"]),
    (["--index", "2-3"], ["단가표.xlsx", "보고서.pdf"]),
    (["--index", "2", "--name", "견적"], ["견적서.pdf", "단가표.xlsx"]),
])
def test_select_by_name_or_index(ol_env, capsys, tmp_path, argv, expected):
    _attach_mail(ol_env, FakeAttachment("견적서.pdf"), FakeAttachment("단가표.xlsx"), FakeAttachment("보고서.pdf"))
    code, out = run_outlook(capsys, "attachments", "--id", "M1", "--save", str(tmp_path / "out"), *argv)
    assert code == 0 and [os.path.basename(p) for p in out["saved"]] == expected, out
    assert sorted(os.listdir(tmp_path / "out")) == sorted(expected)


@pytest.mark.parametrize("argv", [["--name", "계약"], ["--name", "견적;계약"], ["--index", "9"], ["--index", "0"]])
def test_selector_without_match_gives_candidates(ol_env, capsys, tmp_path, argv):
    m = _attach_mail(ol_env, FakeAttachment("견적서.pdf"), FakeAttachment("단가표.xlsx"))
    code, out = run_outlook(capsys, "attachments", "--id", "M1", "--save", str(tmp_path / "out"), *argv)
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:"), out
    assert "[후보] 견적서.pdf, 단가표.xlsx" in out["message"]
    assert all(a.saved_to == [] for a in m.Attachments.items) and not (tmp_path / "out").exists()


def test_bad_index_text_is_stop_retry(ol_env, capsys, tmp_path):
    _attach_mail(ol_env, FakeAttachment("a.pdf"))
    code, out = run_outlook(capsys, "attachments", "--id", "M1", "--save", str(tmp_path / "o"), "--index", "첫째")
    assert code == 1 and "--index" in out["message"]


def test_inline_saved_only_when_selected_blocked_never(ol_env, capsys, tmp_path):
    _attach_mail(ol_env, FakeAttachment("image001.png", hidden=True), FakeAttachment("run.bat"), FakeAttachment("a.pdf"))
    code, out = run_outlook(capsys, "attachments", "--id", "M1", "--save", str(tmp_path / "o"), "--name", "image001;run.bat")
    assert code == 0 and [os.path.basename(p) for p in out["saved"]] == ["image001.png"]
    assert [s["name"] for s in out["skipped"]] == ["run.bat"] and "실행 파일" in out["skipped"][0]["reason"]


def test_only_blocked_and_inline_is_zero_with_reasons(ol_env, capsys):
    _attach_mail(ol_env, FakeAttachment("image001.png", hidden=True), FakeAttachment("run.bat"))
    code, out = run_outlook(capsys, "attachments", "--id", "M1", "--save")
    assert code == 0 and out["count"] == 0 and out["folder"] is None
    assert out["agent_directive"].startswith("VERIFIED: 0건.") and "run.bat" in out["agent_directive"]
    assert not os.path.exists(_default_root(ol_env))


def test_existing_file_is_not_overwritten(ol_env, capsys, tmp_path):
    _attach_mail(ol_env, FakeAttachment("a.pdf"), FakeAttachment("a.pdf"))
    target = tmp_path / "o"
    target.mkdir()
    (target / "a.pdf").write_bytes(b"old")
    code, out = run_outlook(capsys, "attachments", "--id", "M1", "--save", str(target))
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:")
    assert [os.path.basename(p) for p in out["saved"]] == ["a (2).pdf", "a (3).pdf"]
    assert (target / "a.pdf").read_bytes() == b"old"
    assert any("a (2).pdf" in r for r in out["recovered"]) and any("a (3).pdf" in r for r in out["recovered"])


def test_unsafe_attachment_name_is_sanitized(ol_env, capsys, tmp_path):
    _attach_mail(ol_env, FakeAttachment('보고:최종?.pdf'))
    code, out = run_outlook(capsys, "attachments", "--id", "M1", "--save", str(tmp_path / "o"))
    assert code == 0 and os.path.basename(out["saved"][0]) == "보고_최종_.pdf"


@pytest.mark.parametrize("att, word", [(FakeAttachment("b.pdf", vanish=True), "파일이 없음"), (FakeAttachment("b.pdf", fail=True), "저장 실패")])
def test_missing_after_save_is_unverified(ol_env, capsys, tmp_path, att, word):
    _attach_mail(ol_env, FakeAttachment("a.pdf"), att)
    code, out = run_outlook(capsys, "attachments", "--id", "M1", "--save", str(tmp_path / "o"))
    assert code == 0 and out["agent_directive"].startswith("UNVERIFIED:"), out
    assert "b.pdf" in out["agent_directive"] and word in out["agent_directive"]
    assert [os.path.basename(p) for p in out["saved"]] == ["a.pdf"]


def test_empty_file_for_sized_attachment_is_unverified(ol_env, capsys, tmp_path, monkeypatch):
    att = FakeAttachment("a.pdf", size=5000)
    att.SaveAsFile = lambda path: open(path, "wb").close()
    _attach_mail(ol_env, att)
    code, out = run_outlook(capsys, "attachments", "--id", "M1", "--save", str(tmp_path / "o"))
    assert code == 0 and out["agent_directive"].startswith("UNVERIFIED:") and "비어 있음" in out["agent_directive"]


def test_engenius_data_dir_does_not_move_downloads_default(ol_env, capsys, monkeypatch):
    # 도구가 트레이 임시 폴더에서 돌 때처럼 data 폴더가 En-Genius 고정 폴더인 경우 (1.4.35부터 그 경우만)
    home = os.path.join(str(ol_env["home"]), "Documents", "En-Genius", "CLI-data", "office-cli")
    monkeypatch.setattr(rt, "_tool_home", lambda: home)
    monkeypatch.delenv("OFFICE_CLI_DATA_DIR", raising=False)
    assert rt.get_data_dir() == os.path.join(home, "data")
    _attach_mail(ol_env, FakeAttachment("a.pdf"))
    code, out = run_outlook(capsys, "attachments", "--id", "M1", "--save")
    assert code == 0 and os.path.dirname(out["folder"]) == _default_root(ol_env)
    assert not out["folder"].startswith(os.path.join(str(ol_env["home"]), "Documents"))
    assert "_work" not in out["folder"]


def test_download_alias_recommends_attachments(ol_env, capsys):
    code, out = run_outlook(capsys, "download", "--id", "M1")
    assert code == 2 and out["message"].endswith("[추천] attachments")


# ---------------------------------------------------------------------------
# new / reply / forward --attach
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("extra", [[], ["--all"]])
def test_reply_attaches_files(ol_env, capsys, tmp_path, extra):
    m = _attach_mail(ol_env, FakeAttachment("원본.pdf"))
    f1, f2 = tmp_path / "결과 보고.pdf", tmp_path / "표.xlsx"
    f1.write_bytes(b"1")
    f2.write_bytes(b"2")
    code, out = run_outlook(capsys, "reply", "--id", "M1", "--body", "확인했습니다.", "--attach", f"{f1};{f2}", *extra)
    d = m.drafts[0]
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    assert d.Attachments.added == [str(f1), str(f2)] and d.Attachments.Count == 2 and d.displayed and not d.sent
    assert out["attached"] == ["결과 보고.pdf", "표.xlsx"]


def test_forward_keeps_original_and_adds(ol_env, capsys, tmp_path):
    m = _attach_mail(ol_env, FakeAttachment("원본1.pdf"), FakeAttachment("원본2.pdf"))
    f = tmp_path / "추가.docx"
    f.write_bytes(b"1")
    code, out = run_outlook(capsys, "forward", "--id", "M1", "--to", "b@corp.com", "--attach", str(f))
    d = m.drafts[0]
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    assert [a.FileName for a in d.Attachments.items] == ["원본1.pdf", "원본2.pdf", "추가.docx"]
    assert out["attached"] == ["추가.docx"]


def test_new_attaches_several(ol_env, capsys, tmp_path):
    files = [tmp_path / n for n in ("a.pdf", "b.pdf")]
    for f in files:
        f.write_bytes(b"1")
    code, out = run_outlook(capsys, "new", "--to", "a@corp.com", "--subject", "s", "--attach", ";".join(map(str, files)))
    d = ol_env["app"].created[0]
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:") and d.Attachments.Count == 2


@pytest.mark.parametrize("cmd", [["new", "--to", "a@corp.com"], ["reply", "--id", "M1"], ["forward", "--id", "M1"]])
def test_missing_attach_stops_before_draft(ol_env, capsys, tmp_path, cmd):
    m = _attach_mail(ol_env, FakeAttachment("원본.pdf"))
    (tmp_path / "보고서_최종.pdf").write_bytes(b"1")
    (tmp_path / "관계없음.txt").write_bytes(b"1")
    code, out = run_outlook(capsys, *cmd, "--attach", str(tmp_path / "보고서.pdf"))
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:"), out
    assert f"[후보] {tmp_path / '보고서_최종.pdf'}" in out["message"] and "관계없음" not in out["message"]
    assert ol_env["calls"] == [] and m.drafts == [] and ol_env["app"].created == []      # Outlook에 연결하기 전에 멈춤


def test_missing_attach_in_list_stops_and_names_piece(ol_env, capsys, tmp_path):
    ok = tmp_path / "a.pdf"
    ok.write_bytes(b"1")
    code, out = run_outlook(capsys, "new", "--to", "a@corp.com", "--attach", f"{ok};{tmp_path / 'zzz.pdf'}")
    assert code == 1 and "zzz.pdf" in out["message"] and ol_env["calls"] == []


def test_missing_attach_folder_stops(ol_env, capsys, tmp_path):
    code, out = run_outlook(capsys, "new", "--to", "a@corp.com", "--attach", str(tmp_path / "없는 폴더" / "a.pdf"))
    assert code == 1 and "폴더가 없습니다" in out["message"] and ol_env["calls"] == []


def test_attach_count_not_grown_is_unverified_and_not_sent(ol_env, capsys, tmp_path):
    f = tmp_path / "a.pdf"
    f.write_bytes(b"1")
    m = _attach_mail(ol_env, grow=False)
    code, out = run_outlook(capsys, "reply", "--id", "M1", "--attach", str(f))
    assert code == 0 and out["agent_directive"].startswith("UNVERIFIED:") and m.drafts[0].displayed
    code, out = run_outlook(capsys, "reply", "--id", "M1", "--attach", str(f), "--send")
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:")
    assert m.drafts[1].displayed and not m.drafts[1].sent                     # 첨부가 빠진 메일을 보내지 않음


# ---------------------------------------------------------------------------
# 경로(extra)의 메일 계정 이름도 로그에 남기지 않음
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("root", [ROOT_MAIL, ROOT_MAIL.lower()])
def test_folder_path_extra_masks_account_in_log(mail, capsys, monkeypatch, isolated_dirs, root):
    code, out = run_fresh(outlook, capsys, monkeypatch, "list", "--folder", f"{root}\\없음")
    assert code == 1 and f"(경로: {root}\\없음)" in out["message"], out      # 사용자에게는 준 경로 그대로
    res = "\n".join(_result_lines(isolated_dirs))
    assert res and ROOT_MAIL.lower() not in res.lower(), res
    assert "(경로: <메일 계정>\\없음)" in res, res


def test_folder_path_extra_keeps_alias_path_in_log(mail, capsys, monkeypatch, isolated_dirs):
    # 기본 폴더 별칭으로 시작한 경로는 계정 이름이 아니므로 그대로 남김
    code, out = run_fresh(outlook, capsys, monkeypatch, "list", "--folder", "inbox\\없음")
    assert code == 1, out
    res = "\n".join(_result_lines(isolated_dirs))
    assert "(경로: inbox\\없음)" in res and ROOT_MAIL.lower() not in res.lower(), res


# ---------------------------------------------------------------------------
# 짧은 번호, 첨부 저장 뒤 읽기 명령, calendar, thread (C2, 1.4.60)
# ---------------------------------------------------------------------------


class _ThreadMail(_DatedMail):
    def __init__(self, eid, subject, received, topic, body="본문"):
        super().__init__(eid, subject, received)
        self.ConversationTopic = topic
        self.Body = body
        self.Attachments = type("A", (), {"Count": 0})()


class _Items:
    def __init__(self, items, log=None, refuse=False):
        self.items = list(items)
        self.log = log if log is not None else []
        self.refuse = refuse
        self.IncludeRecurrences = False

    def Restrict(self, flt):
        self.log.append(flt)
        if self.refuse:
            raise Exception("조건 거부")
        m = re.search(r"\[ConversationTopic\] = '(.*)'$", flt)
        if m:
            return _Items([i for i in self.items if i.ConversationTopic == m.group(1).replace("''", "'")], self.log)
        return _Items(self.items, self.log)

    def Sort(self, *a):
        self.log.append(("Sort",) + a)

    def __iter__(self):
        return iter(self.items)


@pytest.fixture
def ol_short(monkeypatch):
    mails = _mails()
    by_id = {m.EntryID: m for m in mails}

    class _Folder:
        Name = "받은 편지함"

        @property
        def Items(self):
            return _DatedItems(mails, [])

    class _NS:
        def GetItemFromID(self, eid):
            if eid not in by_id:
                raise RuntimeError("없음")
            return by_id[eid]

    monkeypatch.setattr(outlook, "_today", lambda: TODAY)
    monkeypatch.setattr(rt, "RECOVERED", [])
    monkeypatch.setattr(outlook, "get_outlook", lambda: object())
    monkeypatch.setattr(outlook, "get_namespace", lambda o: _NS())
    monkeypatch.setattr(outlook, "resolve_folder", lambda ns, f: _Folder())
    return by_id


def test_list_gives_short_ids_and_read_accepts_them(ol_short, capsys):
    code, out = run_outlook(capsys, "list", "--limit", "3")
    assert code == 0 and [m["id"] for m in out["mails"]] == [1, 2, 3], out
    second = out["mails"][1]
    code, out2 = run_outlook(capsys, "read", "--id", "2")
    assert code == 0 and out2["mail"]["entry_id"] == second["entry_id"] and out2["mail"]["subject"] == second["subject"]
    code, out3 = run_outlook(capsys, "read", "--id", second["entry_id"])          # 긴 id도 그대로
    assert code == 0 and out3["mail"]["subject"] == second["subject"]
    code, bad = run_outlook(capsys, "read", "--id", "9")
    assert code == 1 and "없는 번호" in bad["message"] and bad["agent_directive"].startswith("STOP_RETRY:")
    assert " list" in bad["next_commands"][0]["command"]


def _listing(prefix, n):
    return [{"entry_id": f"{prefix}{i}", "subject": f"{prefix} 제목 {i}", "from": "홍길동", "received": "2026-11-04T09:00:00"}
            for i in range(1, n + 1)]


def test_short_ids_never_point_to_another_mail_after_new_listing():
    # 회귀: 목록마다 1부터 다시 매겨 search 뒤 thread를 하면 같은 '--id 3'이 다른 메일이 됨 (reply --send가 엉뚱한 메일로)
    search = _listing("S", 5)
    outlook.remember_listing(search, "search")
    assert [m["id"] for m in search] == [1, 2, 3, 4, 5]
    thread = _listing("T", 4) + [{"entry_id": "S3", "subject": "S 제목 3"}]
    outlook.remember_listing(thread, "thread")
    assert [m["id"] for m in thread] == [6, 7, 8, 9, 3]               # 새 번호는 이어서, 이미 본 메일은 같은 번호
    assert outlook.entry_id("3") == "S3" and outlook.entry_id("7") == "T2" and outlook.entry_id("1") == "S1"


def test_short_id_dropped_is_missing_not_other_mail(monkeypatch):
    monkeypatch.setattr(outlook, "ID_KEEP", 4)
    outlook.remember_listing(_listing("A", 3), "list")                   # 1~3
    outlook.remember_listing(_listing("B", 3), "list")                   # 4~6, 기억은 최근 4개(3~6)
    with pytest.raises(rt.InputError) as e:
        outlook.entry_id("1")
    assert "없는 번호" in str(e.value) and e.value.fields["next_commands"]
    assert outlook.entry_id("6") == "B3"
    c = _listing("C", 3)
    outlook.remember_listing(c, "list")                                   # 1부터 다시 매기지 않음
    assert [m["id"] for m in c] == [7, 8, 9] and outlook.entry_id("7") == "C1"


def _assert_old_ids_safe(old):
    """옛 번호(old: 번호 -> entry_id)는 같은 메일이거나 '없는 번호' 오류 (다른 메일을 가리키면 안 됨)."""
    for n, eid in old.items():
        try:
            assert outlook.entry_id(str(n)) == eid, n
        except rt.InputError:
            pass


@pytest.fixture
def no_wait(monkeypatch):
    monkeypatch.setattr(rt.time, "sleep", lambda s: None)


def test_short_ids_corrupt_file_never_restarts_numbering(no_wait):
    # 회귀: 기억 파일이 깨지면 {"next": 1}로 덮어써 번호를 1부터 다시 매김 -> 옛 --id 3이 다른 메일
    a = _listing("A", 5)
    outlook.remember_listing(a, "list")
    old = {m["id"]: m["entry_id"] for m in a}
    path = outlook._listing_path()
    with open(path, "w", encoding="utf-8") as f:
        f.write('{"next": 6, "items": {"3": {"entry_id": "A3"')          # 쓰다 끊긴 파일
    with pytest.raises(rt.InputError) as e:
        outlook.entry_id("3")                                          # 깨진 파일로는 번호를 풀지 않음
    assert "목록을 다시 보세요" in str(e.value) and e.value.fields["next_commands"]
    b = _listing("B", 5)
    warn = outlook.remember_listing(b, "search")
    assert warn and [m["id"] for m in b] == [m["entry_id"] for m in b]   # 이번 결과는 긴 EntryID만
    data_dir = os.path.dirname(path)
    assert any(n.startswith("outlook_ids.broken_") for n in os.listdir(data_dir))   # 깨진 파일은 백업
    c = _listing("C", 5)
    outlook.remember_listing(c, "list")
    assert min(m["id"] for m in c) > max(old)                          # 번호는 이어서 (1부터 다시 아님)
    _assert_old_ids_safe(old)


def test_short_ids_corrupt_file_and_next_file_keeps_largest(no_wait):
    outlook.remember_listing(_listing("A", 4), "list")
    path = outlook._listing_path()
    with open(path, "w", encoding="utf-8") as f:
        f.write("{bad")
    with open(os.path.join(os.path.dirname(path), outlook.NEXT_FILE), "w", encoding="utf-8") as f:
        f.write("40")
    outlook.remember_listing(_listing("B", 2), "list")
    c = _listing("C", 2)
    outlook.remember_listing(c, "list")
    assert [m["id"] for m in c] == [40, 41]


def test_short_ids_read_failure_keeps_file_and_numbering(no_wait, monkeypatch):
    a = _listing("A", 5)
    outlook.remember_listing(a, "list")
    old = {m["id"]: m["entry_id"] for m in a}
    path = outlook._listing_path()
    before = open(path, encoding="utf-8").read()
    orig = rt.read_text_file

    def locked(p):
        if os.path.abspath(p) == os.path.abspath(path):
            raise PermissionError(13, "sharing violation", p)
        return orig(p)
    monkeypatch.setattr(rt, "read_text_file", locked)
    b = _listing("B", 5)
    assert outlook.remember_listing(b, "list")
    assert [m["id"] for m in b] == [m["entry_id"] for m in b]
    with pytest.raises(rt.InputError) as e:
        outlook.entry_id("3")
    assert "목록을 다시 보세요" in str(e.value)
    assert open(path, encoding="utf-8").read() == before                # 읽지 못한 파일은 건드리지 않음
    monkeypatch.setattr(rt, "read_text_file", orig)
    _assert_old_ids_safe(old)
    assert outlook.entry_id("3") == "A3"


def test_short_ids_concurrent_listings_get_distinct_numbers(monkeypatch):
    # 회귀: 동시에 실행한 두 목록이 같은 번호를 받음 (잠금 없이 불러오기~저장)
    import threading
    import time as _time
    orig = outlook._load_ids

    def slow(deadline=None):
        d = orig(deadline)
        _time.sleep(0.2)
        return d
    monkeypatch.setattr(outlook, "_load_ids", slow)
    lists = {t: _listing(t, 5) for t in ("X", "Y")}
    ts = [threading.Thread(target=outlook.remember_listing, args=(lists[t], t)) for t in lists]
    for t in ts:
        t.start()
    for t in ts:
        t.join()
    ids = [m["id"] for t in lists for m in lists[t]]
    assert sorted(ids) == list(range(1, 11)), ids
    for t in lists:
        for m in lists[t]:
            assert outlook.entry_id(str(m["id"])) == m["entry_id"]


def test_short_ids_lock_timeout_and_stale_lock(monkeypatch):
    import time as _time
    data_dir = os.path.dirname(outlook._listing_path())
    os.makedirs(data_dir, exist_ok=True)
    lock = os.path.join(data_dir, outlook.LOCK_FILE)
    open(lock, "w").close()                                            # 만든 실행을 알 수 없는 잠금
    monkeypatch.setattr(outlook, "LOCK_WAIT", 0.2)
    a = _listing("A", 2)
    w = outlook.remember_listing(a, "list")
    assert "outlook_ids.lock이 남아 있음" in w and "30초가 지나면" in w and "다른 실행" not in w, w
    assert [m["id"] for m in a] == ["A1", "A2"]
    old = _time.time() - 3600
    os.utime(lock, (old, old))                                         # 끝나지 못한 실행이 남긴 잠금
    b = _listing("B", 2)
    assert outlook.remember_listing(b, "list") is None and [m["id"] for m in b] == [1, 2]
    assert not os.path.exists(lock)


def test_short_ids_lock_of_dead_process_is_cleared_at_once(monkeypatch):
    # 3차 검토: 잠금을 쥔 실행이 강제 종료되면 30초 동안 번호를 못 붙였음 -> 잠금의 pid가 끝났으면 바로 지움
    import subprocess
    import sys as _sys
    p = subprocess.Popen([_sys.executable, "-c", "pass"])
    p.wait()
    data_dir = os.path.dirname(outlook._listing_path())
    os.makedirs(data_dir, exist_ok=True)
    lock = os.path.join(data_dir, outlook.LOCK_FILE)
    with open(lock, "w") as f:
        f.write(f"{p.pid} dead")
    monkeypatch.setattr(outlook, "LOCK_WAIT", 0.5)
    a = _listing("A", 2)
    assert outlook.remember_listing(a, "list") is None and [m["id"] for m in a] == [1, 2]
    assert not os.path.exists(lock)


def test_short_ids_lock_of_live_process_is_kept(monkeypatch):
    # 살아 있는 실행의 새 잠금은 지우지 않고, 경고에 그 실행을 적음. 풀 때 남의 잠금은 지우지 않음
    data_dir = os.path.dirname(outlook._listing_path())
    os.makedirs(data_dir, exist_ok=True)
    lock = os.path.join(data_dir, outlook.LOCK_FILE)
    with open(lock, "w") as f:
        f.write(f"{os.getpid()} other")
    monkeypatch.setattr(outlook, "LOCK_WAIT", 0.2)
    a = _listing("A", 2)
    w = outlook.remember_listing(a, "list")
    assert f"다른 실행(pid {os.getpid()})이 번호 파일을 쓰는 중" in w, w
    assert open(lock).read() == f"{os.getpid()} other"
    lk = outlook._IdsLock()
    lk.acquired = True                                                 # 남의 잠금을 자기 것으로 착각해도 지우지 않음
    lk.__exit__(None, None, None)
    assert os.path.exists(lock)


def test_short_ids_lock_hold_time_is_bounded(monkeypatch):
    # 3차 검토: 잠금을 쥔 실행이 retry_fs로 30초 넘게 기다려 다른 실행이 쓰는 중인 잠금을 지울 수 있었음
    import time as _time
    outlook.remember_listing(_listing("A", 2), "list")
    path = outlook._listing_path()
    orig = rt.read_text_file

    def locked(p):
        if os.path.abspath(p) == os.path.abspath(path):
            raise PermissionError(13, "sharing violation", p)
        return orig(p)
    monkeypatch.setattr(rt, "read_text_file", locked)
    monkeypatch.setattr(outlook, "LOCK_HOLD", 0.3)
    t0 = _time.monotonic()
    b = _listing("B", 2)
    assert "읽지 못함" in outlook.remember_listing(b, "list")
    assert _time.monotonic() - t0 < outlook.LOCK_STALE / 10 and outlook.LOCK_HOLD < outlook.LOCK_STALE
    assert outlook.LOCK_STALE > 10.0 + 5.0                             # 기본 LOCK_HOLD(10초)보다 넉넉히 김


def test_retry_fs_deadline_stops_waiting():
    import time as _time
    calls = []

    def fail():
        calls.append(1)
        raise PermissionError(13, "locked", "x.json")
    t0 = _time.monotonic()
    with pytest.raises(PermissionError):
        rt.retry_fs(fail, deadline=_time.monotonic() + 0.2)
    assert _time.monotonic() - t0 < 1.0 and len(calls) >= 1


def test_short_id_older_than_keep_days_is_missing(monkeypatch):
    outlook.remember_listing(_listing("A", 2), "list")
    path = outlook._listing_path()
    data = json.loads(open(path, encoding="utf-8").read())
    data["items"]["1"]["t"] = "2026-01-01T00:00:00"
    open(path, "w", encoding="utf-8").write(json.dumps(data))
    with pytest.raises(rt.InputError):
        outlook.entry_id("1")
    assert outlook.entry_id("2") == "A2"


@pytest.mark.parametrize("cmd", [["reply", "--send"], ["forward", "--to", "b@corp.com", "--send"], ["reply"]])
def test_reply_forward_report_original_mail(ol_env, capsys, cmd):
    m = _attach_mail(ol_env, subject="견적 회신 요청")
    code, out = run_outlook(capsys, cmd[0], "--id", "M1", *cmd[1:])
    assert code == 0 and out["original"]["subject"] == "견적 회신 요청" and "lead@corp.com" in out["original"]["from"], out
    assert "견적 회신 요청" in out["preview"] and "original" in out["agent_directive"]
    assert m.drafts


def test_search_next_command_uses_short_id(ol_short, capsys):
    code, out = run_outlook(capsys, "search", "--query", "견적")
    assert code == 0 and out["next_commands"][0]["command"].endswith("read --id 1"), out


def test_short_id_without_listing_is_input_error(ol_short, capsys):
    code, out = run_outlook(capsys, "read", "--id", "1")
    assert code == 1 and "list나 search를 먼저" in out["message"], out


def test_attachments_save_gives_reader_commands(ol_env, capsys, tmp_path):
    _attach_mail(ol_env, FakeAttachment("규격.pdf"), FakeAttachment("시험표.xlsx"), FakeAttachment("사진.png"))
    code, out = run_outlook(capsys, "attachments", "--id", "M1", "--save")
    assert code == 0 and out["count"] == 3, out
    cmds = [c["command"] for c in out["next_commands"]]
    assert any("reader.py" in c and " read --file " in c and "규격.pdf" in c for c in cmds)
    assert any("reader.py" in c and " read --file " in c and "시험표.xlsx" in c for c in cmds)
    assert not any("사진.png" in c for c in cmds)
    assert any(" index --folder " in c for c in cmds)


def test_calendar_lists_events_read_only(monkeypatch, capsys):
    log = []

    class _Ev:
        def __init__(self, subject, start, hours=1, loc="3층 회의실"):
            self.Subject, self.Start, self.End = subject, start, start + dt.timedelta(hours=hours)
            self.AllDayEvent, self.Location, self.Organizer = False, loc, "김팀장"

    evs = [_Ev("주간 회의", dt.datetime(2026, 11, 4, 10)), _Ev("설계 검토", dt.datetime(2026, 11, 6, 14))]

    class _NS:
        def GetDefaultFolder(self, n):
            assert n == outlook.OL_FOLDER_CALENDAR
            return type("F", (), {"Items": _Items(evs, log)})()

    monkeypatch.setattr(outlook, "_today", lambda: TODAY)
    monkeypatch.setattr(outlook, "get_outlook", lambda: object())
    monkeypatch.setattr(outlook, "get_namespace", lambda o: _NS())
    code, out = run_outlook(capsys, "calendar", "--days", "7")
    assert code == 0 and out["count"] == 2 and out["agent_directive"].startswith("VERIFIED:"), out
    assert out["events"][0] == {"subject": "주간 회의", "start": "2026-11-04 10:00", "end": "2026-11-04 11:00", "all_day": False,
                                "location": "3층 회의실", "organizer": "김팀장"}
    assert out["period"] == {"since": "2026-11-04", "until": "2026-11-10", "days": 7}
    assert any("[Start] < '2026-11-11 00:00'" in x for x in log if isinstance(x, str))
    code, out = run_outlook(capsys, "calendar", "--days", "0")
    assert code == 1 and "[추천] 7" in out["message"]


def _calendar_env(monkeypatch, evs, refuse=False):
    class _NS:
        def GetDefaultFolder(self, n):
            return type("F", (), {"Items": _Items(evs, refuse=refuse)})()
    monkeypatch.setattr(rt, "RECOVERED", [])
    monkeypatch.setattr(outlook, "_today", lambda: TODAY)
    monkeypatch.setattr(outlook, "get_outlook", lambda: object())
    monkeypatch.setattr(outlook, "get_namespace", lambda o: _NS())


class _CalEv:
    def __init__(self, subject, start, hours=1):
        self.Subject, self.Start, self.End = subject, start, start + dt.timedelta(hours=hours)
        self.AllDayEvent, self.Location, self.Organizer = False, "", ""


class _BrokenEv:
    @property
    def Start(self):
        raise RuntimeError("COM 오류")


def test_calendar_restrict_refused_compares_each_event(monkeypatch, capsys):
    # 회귀: Outlook이 기간 조건을 거부하면 오류로 끝남 -> 일정마다 시각 비교 (recovered)
    evs = [_CalEv("지난 회의", dt.datetime(2026, 11, 1, 9)), _CalEv("주간 회의", dt.datetime(2026, 11, 4, 10)),
           _CalEv("다음달", dt.datetime(2026, 12, 20, 10))]
    _calendar_env(monkeypatch, evs, refuse=True)
    code, out = run_outlook(capsys, "calendar", "--days", "7")
    assert code == 0 and [e["subject"] for e in out["events"]] == ["주간 회의"] and out["recovered"], out


def test_calendar_unreadable_events_are_not_verified_zero(monkeypatch, capsys):
    # 회귀: 일정 항목 예외를 삼키고 0건 VERIFIED
    _calendar_env(monkeypatch, [_BrokenEv(), _BrokenEv()])
    code, out = run_outlook(capsys, "calendar", "--days", "7")
    assert code == 0 and out["count"] == 0 and out["agent_directive"].startswith("UNVERIFIED:"), out
    assert out["unread_events"] == 2


def test_thread_fallback_scan_limit_is_unverified(monkeypatch, capsys):
    monkeypatch.setattr(outlook, "THREAD_SCAN_MAX", 2)
    inbox = [_ThreadMail(f"x{i}", "다른 일", dt.datetime(2026, 11, 2, 9), "다른 일") for i in range(3)]
    inbox.append(_ThreadMail("i1", "견적 요청", dt.datetime(2026, 11, 1, 9), "견적 요청", "첫 메일"))
    by_id = {m.EntryID: m for m in inbox}

    class _NS:
        def GetItemFromID(self, eid):
            return by_id[eid]

        def GetDefaultFolder(self, n):
            return type("F", (), {"Items": _Items(inbox if n == outlook.OL_FOLDER_INBOX else [], refuse=True)})()
    monkeypatch.setattr(rt, "RECOVERED", [])
    monkeypatch.setattr(outlook, "get_outlook", lambda: object())
    monkeypatch.setattr(outlook, "get_namespace", lambda o: _NS())
    code, out = run_outlook(capsys, "thread", "--id", "i1")
    assert code == 0 and out["agent_directive"].startswith("UNVERIFIED:") and "받은 편지함" in out["agent_directive"], out


def test_calendar_empty_is_zero(monkeypatch, capsys):
    class _NS:
        def GetDefaultFolder(self, n):
            return type("F", (), {"Items": _Items([])})()
    monkeypatch.setattr(outlook, "_today", lambda: TODAY)
    monkeypatch.setattr(outlook, "get_outlook", lambda: object())
    monkeypatch.setattr(outlook, "get_namespace", lambda o: _NS())
    code, out = run_outlook(capsys, "calendar")
    assert code == 0 and out["count"] == 0 and out["agent_directive"].startswith("VERIFIED: 0건"), out


@pytest.mark.parametrize("refuse", [False, True])
def test_thread_collects_conversation_in_time_order(monkeypatch, capsys, refuse):
    inbox = [_ThreadMail("i1", "견적 요청", dt.datetime(2026, 11, 1, 9), "견적 요청", "첫 메일"),
             _ThreadMail("i2", "RE: 견적 요청", dt.datetime(2026, 11, 3, 9), "견적 요청", "셋째"),
             _ThreadMail("x", "다른 일", dt.datetime(2026, 11, 2, 9), "다른 일")]
    sent = [_ThreadMail("s1", "RE: 견적 요청", dt.datetime(2026, 11, 2, 9), "견적 요청", "둘째(보냄)")]
    by_id = {m.EntryID: m for m in inbox + sent}

    class _NS:
        def GetItemFromID(self, eid):
            return by_id[eid]

        def GetDefaultFolder(self, n):
            return type("F", (), {"Items": _Items(inbox if n == outlook.OL_FOLDER_INBOX else sent, refuse=refuse)})()

    monkeypatch.setattr(rt, "RECOVERED", [])
    monkeypatch.setattr(outlook, "get_outlook", lambda: object())
    monkeypatch.setattr(outlook, "get_namespace", lambda o: _NS())
    code, out = run_outlook(capsys, "thread", "--id", "i2")
    assert code == 0 and [m["entry_id"] for m in out["mails"]] == ["i1", "s1", "i2"], out
    assert [m["id"] for m in out["mails"]] == [1, 2, 3] and out["mails"][1]["body"] == "둘째(보냄)"
    assert bool(out.get("recovered")) is refuse
    code, out2 = run_outlook(capsys, "read", "--id", "2")                     # thread 목록 번호로 읽기
    assert code == 0 and out2["mail"]["entry_id"] == "s1"
