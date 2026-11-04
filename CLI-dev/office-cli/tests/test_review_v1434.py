"""1.4.34: 메일 첨부 내려받기 (outlook.py attachments), read/current의 첨부 이름과 내려받기 명령, 답장·전달·새 메일 --attach.

- attachments: 목록(index, name, size, hidden_inline), --save [전체 경로 폴더] (기본: 다운로드\\메일 첨부\\<시각>_<제목>),
  --name/--index로 고르기(맞는 것이 없으면 [후보]), 실행 파일 형식은 저장 안 함, 본문 삽입 이미지는 고른 경우에만,
  같은 이름은 ' (2)'와 recovered, 상대 경로 폴더는 STOP_RETRY, 저장 뒤 파일이 없으면 UNVERIFIED.
- read/current: --save-attach 없음. 첨부가 있으면 attachment_names와 next_commands '첨부 내려받기'.
- new/reply/forward --attach: 없는 파일은 메일을 만들기 전에 STOP_RETRY(같은 폴더의 비슷한 파일 [후보]),
  붙인 뒤 Attachments.Count가 늘었는지 확인(R1). 전달은 원본 첨부를 그대로 두고 더함.
"""
import datetime
import json
import os

import pytest

import _runtime as rt
import outlook


# ---------------------------------------------------------------------------
# Outlook COM 대역
# ---------------------------------------------------------------------------

class Att:
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


class Atts:
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
            self.items.append(Att(os.path.basename(path)))


class Recips:
    Count = 1

    def ResolveAll(self):
        return True


class Draft:
    def __init__(self, atts=(), grow=True):
        self.To = ""
        self.Subject = ""
        self.HTMLBody = "<html><body><p>서명</p></body></html>"
        self.Attachments = Atts(atts, grow)
        self.Recipients = Recips()
        self.GetInspector = object()
        self.displayed = self.sent = False

    def Display(self):
        self.displayed = True

    def Send(self):
        self.sent = True


class Mail:
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
        self.Attachments = Atts(atts)
        self.grow = grow
        self.drafts = []

    def _draft(self, atts=()):
        d = Draft(atts, self.grow)
        self.drafts.append(d)
        return d

    def Reply(self):
        return self._draft()

    def ReplyAll(self):
        return self._draft()

    def Forward(self):                                       # 전달은 원본 첨부를 가지고 시작
        return self._draft([Att(a.FileName) for a in self.Attachments.items])


class Selection:
    def __init__(self, item):
        self.item = item
        self.Count = 1 if item is not None else 0

    def Item(self, i):
        return self.item


class App:
    def __init__(self, mail=None, grow=True):
        self.mail = mail
        self.created = []
        self.grow = grow

    def ActiveInspector(self):
        return None

    def ActiveExplorer(self):
        return type("E", (), {"Selection": Selection(self.mail)})()

    def CreateItem(self, kind):
        d = Draft(grow=self.grow)
        self.created.append(d)
        return d


@pytest.fixture
def env(monkeypatch, tmp_path):
    """Outlook 대역과 사용자 폴더(USERPROFILE). 다운로드 폴더 레지스트리는 읽지 않음."""
    home = tmp_path / "home"
    home.mkdir()
    monkeypatch.setenv("USERPROFILE", str(home))
    monkeypatch.setattr(outlook, "_shell_downloads", lambda: None)
    monkeypatch.setattr(rt, "RECOVERED", [])
    monkeypatch.setattr(rt, "CURRENT_ARGV", None)
    state = {"home": home, "app": App(), "calls": []}

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


def run(capsys, *argv):
    with pytest.raises(SystemExit) as ei:
        outlook.main(list(argv))
    lines = [ln for ln in capsys.readouterr().out.splitlines() if ln.strip()]
    assert len(lines) == 1, lines
    return ei.value.code, json.loads(lines[0])


def _mail(env, *atts, **kw):
    env["app"].mail = Mail(atts, **kw)
    return env["app"].mail


def _default_root(env):
    return os.path.join(str(env["home"]), "Downloads", "메일 첨부")


# ---------------------------------------------------------------------------
# attachments: 목록
# ---------------------------------------------------------------------------

def test_list_attachments_with_index_size_inline(env, capsys):
    _mail(env, Att("견적서.pdf", 2048), Att("image001.png", 500, hidden=True), Att("setup.exe", 10))
    code, out = run(capsys, "attachments", "--id", "M1")
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:") and out["count"] == 3
    assert out["attachments"][0] == {"index": 1, "name": "견적서.pdf", "size": 2048, "hidden_inline": False,
                                     "blocked": False}
    assert out["attachments"][1]["hidden_inline"] is True and out["attachments"][2]["blocked"] is True
    assert out["truncated"] is False and "saved" not in out
    assert out["next_commands"][0]["command"].endswith("attachments --id M1 --save")
    assert not os.path.exists(_default_root(env))                     # 목록만: 아무것도 저장하지 않음


def test_list_uses_selected_mail_without_id(env, capsys):
    _mail(env, Att("a.pdf"), entry="SEL")
    code, out = run(capsys, "attachments")
    assert code == 0 and out["count"] == 1 and out["mail"]["id"] == "SEL"
    assert "attachments --id SEL --save" in out["next_commands"][0]["command"]    # 다음 명령은 id로 고정


def test_zero_attachments_is_explicit(env, capsys):
    _mail(env)
    code, out = run(capsys, "attachments", "--id", "M1")
    assert code == 0 and out["count"] == 0 and out["attachments"] == []
    assert out["agent_directive"].startswith("VERIFIED: 0건.") and "첨부 없음" in out["agent_directive"]
    code, out = run(capsys, "attachments", "--id", "M1", "--save")
    assert code == 0 and out["count"] == 0 and out["saved"] == [] and out["folder"] is None
    assert "첨부 없음" in out["agent_directive"] and not os.path.exists(_default_root(env))


def test_no_mail_selected_is_stop_retry(env, capsys):
    code, out = run(capsys, "attachments")
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:") and "--id" in out["message"]


def test_name_without_save_is_refused(env, capsys):
    _mail(env, Att("a.pdf"))
    code, out = run(capsys, "attachments", "--id", "M1", "--name", "a")
    assert code == 1 and "--save" in out["message"]


# ---------------------------------------------------------------------------
# attachments --save
# ---------------------------------------------------------------------------

def test_save_all_to_downloads_skips_blocked_and_inline(env, capsys):
    m = _mail(env, Att("견적서.pdf", 2048), Att("image001.png", 500, hidden=True), Att("setup.exe", 10),
              Att("단가표.xlsx", 4096))
    code, out = run(capsys, "attachments", "--id", "M1", "--save")
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    folder = out["folder"]
    assert os.path.dirname(folder) == _default_root(env)
    assert os.path.basename(folder).endswith("_RE_ 견적_회신")          # rt.safe_filename(제목)
    assert [os.path.basename(p) for p in out["saved"]] == ["견적서.pdf", "단가표.xlsx"] and out["count"] == 2
    assert all(os.path.isfile(p) and os.path.dirname(p) == folder for p in out["saved"])
    reasons = {s["name"]: s["reason"] for s in out["skipped"]}
    assert "실행 파일" in reasons["setup.exe"] and "본문 삽입" in reasons["image001.png"]
    assert folder in out["preview"] and "folder" in out["agent_directive"]
    assert m.Attachments.items[2].saved_to == []                       # 실행 파일은 저장 시도도 하지 않음
    assert "recovered" not in out


def test_save_to_user_folder_creates_it(env, capsys, tmp_path):
    _mail(env, Att("a.pdf"))
    target = tmp_path / "내 폴더" / "첨부"
    code, out = run(capsys, "attachments", "--id", "M1", "--save", str(target))
    assert code == 0 and out["folder"] == str(target) and os.path.isfile(target / "a.pdf")


def test_relative_folder_is_refused(env, capsys, tmp_path, monkeypatch):
    m = _mail(env, Att("a.pdf"))
    monkeypatch.chdir(tmp_path)
    code, out = run(capsys, "attachments", "--id", "M1", "--save", "첨부")
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
def test_select_by_name_or_index(env, capsys, tmp_path, argv, expected):
    _mail(env, Att("견적서.pdf"), Att("단가표.xlsx"), Att("보고서.pdf"))
    code, out = run(capsys, "attachments", "--id", "M1", "--save", str(tmp_path / "out"), *argv)
    assert code == 0 and [os.path.basename(p) for p in out["saved"]] == expected, out
    assert sorted(os.listdir(tmp_path / "out")) == sorted(expected)


@pytest.mark.parametrize("argv", [["--name", "계약"], ["--name", "견적;계약"], ["--index", "9"], ["--index", "0"]])
def test_selector_without_match_gives_candidates(env, capsys, tmp_path, argv):
    m = _mail(env, Att("견적서.pdf"), Att("단가표.xlsx"))
    code, out = run(capsys, "attachments", "--id", "M1", "--save", str(tmp_path / "out"), *argv)
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:"), out
    assert "[후보] 견적서.pdf, 단가표.xlsx" in out["message"]
    assert all(a.saved_to == [] for a in m.Attachments.items) and not (tmp_path / "out").exists()


def test_bad_index_text_is_stop_retry(env, capsys, tmp_path):
    _mail(env, Att("a.pdf"))
    code, out = run(capsys, "attachments", "--id", "M1", "--save", str(tmp_path / "o"), "--index", "첫째")
    assert code == 1 and "--index" in out["message"]


def test_inline_saved_only_when_selected_blocked_never(env, capsys, tmp_path):
    _mail(env, Att("image001.png", hidden=True), Att("run.bat"), Att("a.pdf"))
    code, out = run(capsys, "attachments", "--id", "M1", "--save", str(tmp_path / "o"), "--name", "image001;run.bat")
    assert code == 0 and [os.path.basename(p) for p in out["saved"]] == ["image001.png"]
    assert [s["name"] for s in out["skipped"]] == ["run.bat"] and "실행 파일" in out["skipped"][0]["reason"]


def test_only_blocked_and_inline_is_zero_with_reasons(env, capsys):
    _mail(env, Att("image001.png", hidden=True), Att("run.bat"))
    code, out = run(capsys, "attachments", "--id", "M1", "--save")
    assert code == 0 and out["count"] == 0 and out["folder"] is None
    assert out["agent_directive"].startswith("VERIFIED: 0건.") and "run.bat" in out["agent_directive"]
    assert not os.path.exists(_default_root(env))


def test_existing_file_is_not_overwritten(env, capsys, tmp_path):
    _mail(env, Att("a.pdf"), Att("a.pdf"))
    target = tmp_path / "o"
    target.mkdir()
    (target / "a.pdf").write_bytes(b"old")
    code, out = run(capsys, "attachments", "--id", "M1", "--save", str(target))
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:")
    assert [os.path.basename(p) for p in out["saved"]] == ["a (2).pdf", "a (3).pdf"]
    assert (target / "a.pdf").read_bytes() == b"old"
    assert any("a (2).pdf" in r for r in out["recovered"]) and any("a (3).pdf" in r for r in out["recovered"])


def test_unsafe_attachment_name_is_sanitized(env, capsys, tmp_path):
    _mail(env, Att('보고:최종?.pdf'))
    code, out = run(capsys, "attachments", "--id", "M1", "--save", str(tmp_path / "o"))
    assert code == 0 and os.path.basename(out["saved"][0]) == "보고_최종_.pdf"


@pytest.mark.parametrize("att, word", [(Att("b.pdf", vanish=True), "파일이 없음"), (Att("b.pdf", fail=True), "저장 실패")])
def test_missing_after_save_is_unverified(env, capsys, tmp_path, att, word):
    _mail(env, Att("a.pdf"), att)
    code, out = run(capsys, "attachments", "--id", "M1", "--save", str(tmp_path / "o"))
    assert code == 0 and out["agent_directive"].startswith("UNVERIFIED:"), out
    assert "b.pdf" in out["agent_directive"] and word in out["agent_directive"]
    assert [os.path.basename(p) for p in out["saved"]] == ["a.pdf"]


def test_empty_file_for_sized_attachment_is_unverified(env, capsys, tmp_path, monkeypatch):
    att = Att("a.pdf", size=5000)
    att.SaveAsFile = lambda path: open(path, "wb").close()
    _mail(env, att)
    code, out = run(capsys, "attachments", "--id", "M1", "--save", str(tmp_path / "o"))
    assert code == 0 and out["agent_directive"].startswith("UNVERIFIED:") and "비어 있음" in out["agent_directive"]


def test_engenius_data_dir_does_not_move_downloads_default(env, capsys, monkeypatch):
    monkeypatch.setenv("ENGENIUS_OUT", str(env["home"] / "out"))
    monkeypatch.delenv("OFFICE_CLI_DATA_DIR", raising=False)
    engenius_data = os.path.join(str(env["home"]), "Documents", "En-Genius", "CLI-data", "office-cli", "data")
    assert rt.get_data_dir() == engenius_data
    _mail(env, Att("a.pdf"))
    code, out = run(capsys, "attachments", "--id", "M1", "--save")
    assert code == 0 and os.path.dirname(out["folder"]) == _default_root(env)
    assert not out["folder"].startswith(os.path.join(str(env["home"]), "Documents"))
    assert "_work" not in out["folder"]


def test_download_alias_recommends_attachments(env, capsys):
    code, out = run(capsys, "download", "--id", "M1")
    assert code == 2 and out["message"].endswith("[추천] attachments")


# ---------------------------------------------------------------------------
# read / current
# ---------------------------------------------------------------------------

def test_read_shows_attachment_names_and_download_command(env, capsys):
    _mail(env, Att("견적서.pdf"), Att("image001.png", hidden=True))
    code, out = run(capsys, "read", "--id", "M1")
    assert code == 0 and out["attachment_names"] == ["견적서.pdf"] and out["attachment_names_truncated"] is False
    assert out["inline_images"] == 1
    assert [c["purpose"] for c in out["next_commands"]] == ["첨부 내려받기"]
    assert out["next_commands"][0]["command"].endswith("attachments --id M1 --save")
    assert "saved_attachments" not in out and not os.path.exists(_default_root(env))   # 읽기만으로 저장하지 않음


def test_read_save_puts_word_command_first(env, capsys):
    _mail(env, Att("a.pdf"))
    code, out = run(capsys, "read", "--id", "M1", "--save")
    assert code == 0 and [c["purpose"] for c in out["next_commands"]] == ["메일 내용을 워드에 넣기", "첨부 내려받기"]


def test_current_caps_attachment_names(env, capsys):
    _mail(env, *[Att(f"f{i:02d}.pdf") for i in range(25)], entry="SEL")
    code, out = run(capsys, "current")
    assert code == 0 and len(out["attachment_names"]) == 20 and out["attachment_names_truncated"] is True
    assert "attachments --id SEL --save" in out["next_commands"][0]["command"]


def test_read_without_attachments_has_no_download(env, capsys):
    _mail(env, Att("image001.png", hidden=True))
    code, out = run(capsys, "read", "--id", "M1")
    assert code == 0 and "attachment_names" not in out and "next_commands" not in out


def test_save_attach_option_is_gone(env, capsys):
    _mail(env, Att("a.pdf"))
    code, out = run(capsys, "read", "--id", "M1", "--save-attach")
    assert code == 2 and out["agent_directive"].startswith("STOP_RETRY:")


# ---------------------------------------------------------------------------
# new / reply / forward --attach
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("extra", [[], ["--all"]])
def test_reply_attaches_files(env, capsys, tmp_path, extra):
    m = _mail(env, Att("원본.pdf"))
    f1, f2 = tmp_path / "결과 보고.pdf", tmp_path / "표.xlsx"
    f1.write_bytes(b"1")
    f2.write_bytes(b"2")
    code, out = run(capsys, "reply", "--id", "M1", "--body", "확인했습니다.", "--attach", f"{f1};{f2}", *extra)
    d = m.drafts[0]
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    assert d.Attachments.added == [str(f1), str(f2)] and d.Attachments.Count == 2 and d.displayed and not d.sent
    assert out["attached"] == ["결과 보고.pdf", "표.xlsx"]


def test_forward_keeps_original_and_adds(env, capsys, tmp_path):
    m = _mail(env, Att("원본1.pdf"), Att("원본2.pdf"))
    f = tmp_path / "추가.docx"
    f.write_bytes(b"1")
    code, out = run(capsys, "forward", "--id", "M1", "--to", "b@corp.com", "--attach", str(f))
    d = m.drafts[0]
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    assert [a.FileName for a in d.Attachments.items] == ["원본1.pdf", "원본2.pdf", "추가.docx"]
    assert out["attached"] == ["추가.docx"]


def test_new_attaches_several(env, capsys, tmp_path):
    files = [tmp_path / n for n in ("a.pdf", "b.pdf")]
    for f in files:
        f.write_bytes(b"1")
    code, out = run(capsys, "new", "--to", "a@corp.com", "--subject", "s", "--attach", ";".join(map(str, files)))
    d = env["app"].created[0]
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:") and d.Attachments.Count == 2


@pytest.mark.parametrize("cmd", [["new", "--to", "a@corp.com"], ["reply", "--id", "M1"], ["forward", "--id", "M1"]])
def test_missing_attach_stops_before_draft(env, capsys, tmp_path, cmd):
    m = _mail(env, Att("원본.pdf"))
    (tmp_path / "보고서_최종.pdf").write_bytes(b"1")
    (tmp_path / "관계없음.txt").write_bytes(b"1")
    code, out = run(capsys, *cmd, "--attach", str(tmp_path / "보고서.pdf"))
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:"), out
    assert f"[후보] {tmp_path / '보고서_최종.pdf'}" in out["message"] and "관계없음" not in out["message"]
    assert env["calls"] == [] and m.drafts == [] and env["app"].created == []      # Outlook에 연결하기 전에 멈춤


def test_missing_attach_in_list_stops_and_names_piece(env, capsys, tmp_path):
    ok = tmp_path / "a.pdf"
    ok.write_bytes(b"1")
    code, out = run(capsys, "new", "--to", "a@corp.com", "--attach", f"{ok};{tmp_path / 'zzz.pdf'}")
    assert code == 1 and "zzz.pdf" in out["message"] and env["calls"] == []


def test_missing_attach_folder_stops(env, capsys, tmp_path):
    code, out = run(capsys, "new", "--to", "a@corp.com", "--attach", str(tmp_path / "없는 폴더" / "a.pdf"))
    assert code == 1 and "폴더가 없습니다" in out["message"] and env["calls"] == []


def test_attach_count_not_grown_is_unverified_and_not_sent(env, capsys, tmp_path):
    f = tmp_path / "a.pdf"
    f.write_bytes(b"1")
    m = _mail(env, grow=False)
    code, out = run(capsys, "reply", "--id", "M1", "--attach", str(f))
    assert code == 0 and out["agent_directive"].startswith("UNVERIFIED:") and m.drafts[0].displayed
    code, out = run(capsys, "reply", "--id", "M1", "--attach", str(f), "--send")
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:")
    assert m.drafts[1].displayed and not m.drafts[1].sent                     # 첨부가 빠진 메일을 보내지 않음
