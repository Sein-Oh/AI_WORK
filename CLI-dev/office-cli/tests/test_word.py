"""word.py: write(마크다운·앵커), replace, read 범위, outline, comment, revisions."""
import json
import struct

import pytest

import _office as off
import fake_word as fw
import word
from fake_office import make_win32com
from fake_word import FakeComment, FakeRange, FakeWordDocument
from office_helpers import assert_not_logged, Bag, option_value, recommended_argv, shell_sim, CODE_MD, EMPTY_FIND, run_as_tool, run_main, run_word, set_active, split_command, stub_com, undo_points


def doc_paras(doc):
    return doc.text.split("\r")


class _Shapes:
    def __init__(self, starts):
        self.starts = starts
        self.Count = len(starts)

    def Item(self, i):
        return type("S", (), {"Range": type("R", (), {"Start": self.starts[i - 1]})()})()


class _Para:
    """COM 단락: Range.Text, Range.Start, Range.InlineShapes(그림 위치), OutlineLevel."""
    def __init__(self, start, text, pictures=(), level=10):
        self.Range = type("R", (), {"Text": text, "Start": start, "InlineShapes": _Shapes([start + p for p in pictures])})()
        self.OutlineLevel = level


def _doc_with_table():
    doc = FakeWordDocument("머리말\r시험 온도\r25도\r표 뒤 문단")
    doc.table_spans = [(doc.text.index("시험 온도"), doc.text.index("표 뒤 문단"))]   # '시험 온도', '25도' 두 칸
    return doc


class _PageShapes:
    Count = 0


class _PageRange:
    def __init__(self, doc, start, end):
        self.doc = doc
        self.Start = start
        self.End = end
        self.InlineShapes = _PageShapes()

    @property
    def Text(self):
        return self.doc.full[self.Start:self.End]

    def Information(self, kind):
        assert kind == word.WD_PAGE_NUMBER
        return self.doc.page_at(self.Start)


class _PagePara:
    def __init__(self, rng, level):
        self.Range = rng
        self.OutlineLevel = level


class PageDoc:
    """단락(글, 개요 수준)과 쪽 시작 위치로 만든 가짜 문서. 개요 수준 10은 본문."""

    def __init__(self, paras, page_starts=None, name="규격.docx"):
        self.full = ""
        self.Paragraphs = []
        for text, level in paras:
            s = len(self.full)
            self.full += text + "\r"
            self.Paragraphs.append(_PagePara(_PageRange(self, s, len(self.full)), level))
        self.page_starts = page_starts or [0]
        self.Name = name
        self.FullName = "C:\\docs\\" + name
        self.Content = _PageRange(self, 0, len(self.full))

    def start_of(self, text):
        return self.full.index(text)

    def page_at(self, pos):
        return max(i + 1 for i, st in enumerate(self.page_starts) if st <= pos)

    def ComputeStatistics(self, kind):
        assert kind == 2
        return len(self.page_starts)

    def GoTo(self, What=None, Which=None, Count=1):
        assert (What, Which) == (1, 1)
        return _PageRange(self, self.page_starts[min(Count, len(self.page_starts)) - 1], len(self.full))

    def Range(self, s, e):
        return _PageRange(self, s, e)


SPEC = [("1. 개요", 1), ("개요 본문", 10), ("1.1 범위", 2), ("범위 본문", 10), ("1.1.1 세부", 3), ("세부 본문", 10),
        ("1.2 용어", 2), ("용어 본문", 10), ("2. 시험", 1), ("시험 본문", 10)]


@pytest.fixture
def wd_pages(monkeypatch):
    holder = {"doc": PageDoc(SPEC)}
    fake, _ = make_win32com()
    monkeypatch.setattr(off, "win32com", fake)
    monkeypatch.setattr(off, "pythoncom", None)
    monkeypatch.setattr(off, "list_open", lambda kind: [])
    monkeypatch.setattr(off, "get_active", lambda kind: {"name": holder["doc"].Name, "path": holder["doc"].FullName,
                                                          "app": object(), "doc": holder["doc"]})
    monkeypatch.setattr(off, "save_snapshot", lambda *a, **k: None)
    return holder


def texts(out):
    return [p["text"] for p in out["paragraphs"]]


@pytest.fixture
def cd(wd_pages):
    wd_pages["doc"] = FakeWordDocument("시험 온도는 25도로 한다.\r습도 조건\r시험 시간과 시험 온도 기록")
    return wd_pages


class _Rev:
    def __init__(self, kind, text):
        self.Type = kind
        self.Author = "홍길동"
        self.Date = "2026-10-27"
        self.Range = type("R", (), {"Text": text, "Information": lambda self, k: 4})()


class _Revs:
    def __init__(self, items):
        self.items = items
        self.Count = len(items)

    def __call__(self, i):
        return self.items[i - 1]


@pytest.fixture
def wd_office(monkeypatch):
    holder = {"doc": FakeWordDocument("office 문서")}
    stub_com(monkeypatch)
    monkeypatch.setattr(off, "get_active", lambda kind: {"name": holder["doc"].Name, "path": holder["doc"].FullName,
                                                          "app": object(), "doc": holder["doc"]})
    monkeypatch.setattr(off, "save_snapshot", lambda *a, **k: None)
    return holder


@pytest.fixture
def wd_doc(monkeypatch):
    doc = FakeWordDocument("추진 배경\r2026년 계획입니다.\r일정")
    stub_com(monkeypatch)
    set_active(monkeypatch, object(), doc)
    return doc


@pytest.fixture
def wd_iphone(monkeypatch):
    doc = FakeWordDocument("iPhone and iPhone\r끝")
    stub_com(monkeypatch)
    set_active(monkeypatch, object(), doc)
    return doc


@pytest.fixture
def wd_caret(monkeypatch, tmp_path):
    monkeypatch.setenv("OFFICE_CLI_DATA_DIR", str(tmp_path / "d"))
    holder = {"doc": FakeWordDocument("제목\r본문 ^ 10^6 기호\r일정")}
    fake, _ = make_win32com()
    monkeypatch.setattr(off, "win32com", fake)
    monkeypatch.setattr(off, "pythoncom", None)
    monkeypatch.setattr(off, "list_open", lambda kind: [])
    monkeypatch.setattr(off, "get_active", lambda kind: {"name": holder["doc"].Name, "path": holder["doc"].FullName,
                                                         "app": object(), "doc": holder["doc"]})
    return holder


def _run_word(capsys, *argv):
    with pytest.raises(SystemExit) as ei:
        word.main(list(argv))
    out = [ln for ln in capsys.readouterr().out.splitlines() if ln.strip()]
    assert len(out) == 1, out
    return ei.value.code, json.loads(out[0])


def _word_md(capsys, tmp_path, md):
    p = tmp_path / "w.md"
    p.write_text(md, encoding="utf-8")
    return run_word(capsys, "write", "--file", str(p))


EMOJI = chr(0x1F680)                   # 저장소 파일에 이모지를 직접 쓰지 않음


def _u16_units(v):
    b = v.encode("utf-16-le", "surrogatepass")
    return "".join(chr(x) for x in struct.unpack("<%dH" % (len(b) // 2), b))


def _back(s):
    return "".join(s).encode("utf-16-le", "surrogatepass").decode("utf-16-le")


class _U16Doc(fw.FakeWordDocument):
    """Word처럼 글을 UTF-16 단위로 담는 가짜 문서 (BMP 밖 글자는 2칸)."""
    def replace(self, s, e, v):
        super().replace(s, e, _u16_units(v))


@pytest.fixture
def u16_doc(monkeypatch):
    def insert_after(self, s):
        self.doc.replace(self.End, self.End, s)
        self.End += len(_u16_units(s))
    monkeypatch.setattr(fw.FakeRange, "InsertAfter", insert_after)
    return _U16Doc("")


def _bold_texts(d):
    return [_back(d.text[tag[1]:tag[2]]) for tag, k, v in d.log
            if isinstance(tag, tuple) and tag[0] == "font" and k == "Bold" and v]


# ---------------------------------------------------------------------------
# write / replace / read 기본 동작
# ---------------------------------------------------------------------------


def test_com_units_inline_picture_becomes_marker():
    # Word는 본문 그림을 '\x01'로 돌려줌. 그림 위치의 그 글자만 [그림]으로 바꾸고 본문의 '/'는 그대로
    doc = type("D", (), {"Paragraphs": [
        _Para(0, "첫 문단\r"),
        _Para(5, "\x01\r", pictures=[0]),                  # 그림만 있는 단락
        _Para(7, "그래프 \x01 참고\r", pictures=[4]),        # 글 사이 그림
        _Para(16, "A/B 비율\r"),                            # 그림 없는 '/'
        _Para(23, "A/B \x01\r", pictures=[4]),              # '/'와 그림이 같은 단락
    ]})()
    texts = [u["text"] for u in word._com_units(doc)]
    assert texts == ["첫 문단", "[그림]", "그래프 [그림] 참고", "A/B 비율", "A/B [그림]"]


def test_write_plain_text_at_end(wd, capsys):
    code, out = run_word(capsys, "write", "--text", "새 문단입니다")
    assert code == 0, out
    assert doc_paras(wd["doc"])[:4] == ["추진 배경", "2026년 계획입니다.", "일정", "새 문단입니다"]
    assert out["agent_directive"].startswith("VERIFIED:") and out["verified"] is True        # 넣은 범위를 다시 읽음 (R1)


def test_write_table_first_markdown_is_verified_and_missing_text_is_not(wd, capsys, monkeypatch):
    code, out = run_word(capsys, "write", "--text", "| 공정 | 수율 |\n|---|---|\n| 가열 | 98 |")
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:")
    monkeypatch.setattr(word, "_verify_inserted", lambda *a: ["넣은 범위에서 'x'을(를) 찾지 못함"])
    code, out = run_word(capsys, "write", "--text", "새 문단")
    assert code == 0 and out["agent_directive"].startswith("UNVERIFIED:") and out["unverified"]


def test_write_after_anchor_goes_after_whole_paragraph(wd, capsys):
    code, out = run_word(capsys, "write", "--text", "추가 문단", "--anchor", "계획", "--pos", "after")
    assert code == 0
    assert doc_paras(wd["doc"])[:4] == ["추진 배경", "2026년 계획입니다.", "추가 문단", "일정"]   # 단락 중간을 쪼개지 않음


def test_anchor_in_several_paragraphs_is_not_guessed(wd, capsys):
    wd["doc"] = FakeWordDocument("계획 개요\r본문\r세부 계획\r끝")
    before = wd["doc"].text
    code, out = run_word(capsys, "write", "--text", "새 문단", "--anchor", "계획", "--pos", "replace")
    assert code == 1 and out["error_type"] == "ambiguous" and out["agent_directive"].startswith("STOP_RETRY:")
    assert out["candidates"] == ["1: 계획 개요", "2: 세부 계획"] and "--occurrence" in out["message"]
    assert [c["command"].split()[-2:] for c in out["confirm_commands"]] == [["--occurrence", "1"], ["--occurrence", "2"]]
    assert wd["doc"].text == before                                       # 어느 단락도 지우지 않음
    code, out = run_word(capsys, "write", "--text", "새 문단", "--anchor", "계획", "--pos", "replace", "--occurrence", "2")
    assert code == 0 and doc_paras(wd["doc"])[:4] == ["계획 개요", "본문", "새 문단", "끝"]
    code, out = run_word(capsys, "write", "--text", "x", "--anchor", "개요", "--occurrence", "5")
    assert code == 1 and "1~1" in out["message"]


def test_anchor_same_paragraph_twice_is_one_candidate(wd, capsys):
    wd["doc"] = FakeWordDocument("계획과 계획\r끝")
    code, out = run_word(capsys, "write", "--text", "추가", "--anchor", "계획")
    assert code == 0 and doc_paras(wd["doc"])[:3] == ["계획과 계획", "추가", "끝"]


def test_anchor_not_found_recommends_similar_text(wd, capsys):
    code, out = run_word(capsys, "write", "--text", "x", "--anchor", "2026 계획")
    assert code == 1 and out["error_type"] == "input", out
    argv = recommended_argv(out["message"], "word")             # 1.4.79: [추천]은 전체 명령
    assert argv == ["write", "--text", "x", "--anchor", "2026년 계획"], argv
    code, out = run_word(capsys, *argv)                         # 추천대로 다시 하면 찾음
    assert code == 0


def test_anchor_position_word_recommends_end_or_start(wd, capsys):
    """1.4.89 집 PC Gemma O4: '문서 끝에 덧붙여'를 --anchor "end"로 써서 '찾을 수 없음'으로 멈추고, 마지막 단락 글을 앵커로 다시 씀.
    문서에 없는 위치 낱말은 --anchor를 뺀 명령(끝) 또는 --pos start(처음)를 [추천]. 문서에 그 글이 있으면 원래대로 그 단락."""
    code, out = run_word(capsys, "write", "--text", "x", "--anchor", "end")
    assert code == 1 and "[추천]" in out["message"], out
    argv = recommended_argv(out["message"], "word")
    assert argv == ["write", "--text", "x"], argv
    code, out = run_word(capsys, *argv)
    assert code == 0 and [p for p in doc_paras(wd["doc"]) if p][-1] == "x", doc_paras(wd["doc"])
    code, out = run_word(capsys, "write", "--text", "y", "--anchor", "처음")
    assert recommended_argv(out["message"], "word") == ["write", "--text", "y", "--pos", "start"], out
    code, out = run_word(capsys, "write", "--text", "z", "--anchor", "end", "--pos", "replace")   # 뜻이 다른 위치는 [추천] 안 함
    assert code == 1 and "[추천]" not in out["message"], out
    wd["doc"] = FakeWordDocument("계획\r끝\r일정")
    code, out = run_word(capsys, "write", "--text", "w", "--anchor", "끝")          # 문서에 있는 글이면 그 단락 뒤
    assert code == 0 and doc_paras(wd["doc"])[:3] == ["계획", "끝", "w"], doc_paras(wd["doc"])


def test_anchor_not_found_does_not_recommend_different_text(wd, capsys):
    """글자가 바뀐 문구('2027' -> '2026')는 뜻이 다를 수 있어 [추천]하지 않음. replace는 끼어든 글자도 [추천] 안 함."""
    before = wd["doc"].text
    code, out = run_word(capsys, "write", "--text", "x", "--anchor", "2027년 계획")
    assert code == 1 and "[추천]" not in out["message"] and "비슷한 문구: '2026년 계획'" in out["message"]
    code, out = run_word(capsys, "write", "--text", "x", "--anchor", "2026 계획", "--pos", "replace")
    assert code == 1 and "[추천]" not in out["message"] and "비슷한 문구" in out["message"]
    assert wd["doc"].text == before


def test_safe_anchor_fix_rules():
    assert word._safe_anchor_fix("2026년계획", "2026년 계획", "replace")       # 공백만 다름
    assert word._safe_anchor_fix("2026 계획", "2026년 계획", "after")          # 끼어든 글자 1자
    assert not word._safe_anchor_fix("2026 계획", "2026년 계획", "replace")
    for anchor, similar in [("Q2 매출 보고", "Q1 매출 보고"), ("결론 분석", "결과 분석"),
                            ("Sales 2025", "Sales 202"), ("3. 결과요약", "2. 결과")]:
        assert not word._safe_anchor_fix(anchor, similar, "after"), (anchor, similar)


def test_write_before_and_replace_anchor(wd, capsys):
    run_word(capsys, "write", "--text", "머리말", "--anchor", "추진", "--pos", "before")
    assert doc_paras(wd["doc"])[0] == "머리말"
    run_word(capsys, "write", "--text", "새 계획", "--anchor", "2026년", "--pos", "replace")
    assert "2026년 계획입니다." not in wd["doc"].text and "새 계획" in doc_paras(wd["doc"])


def test_write_markdown_blocks_styles_and_table(wd, capsys):
    md = "## 결론\n본문 **강조** 끝\n\n- 가\n- 나\n\n| 공정 | 수율 |\n|---|---|\n| 가열 | 98 |"
    code, out = run_word(capsys, "write", "--text", md)
    assert code == 0
    doc = wd["doc"]
    assert ("style", "결론", word.heading_style(2)) in doc.log
    assert ("style", "가", word.WD_STYLE_LIST_BULLET) in doc.log
    assert "본문 강조 끝" in doc_paras(doc)
    assert any(e[0] == ("font", doc.text.index("강조"), doc.text.index("강조") + 2) and e[1] == "Bold" and e[2] is True
               for e in doc.log)
    assert doc.tables == [["공정\t수율", "가열\t98"]]
    assert out["inserted"] == ["제목2", "본문", "글머리 목록 2개", "표 2x2"]


def test_write_escaped_newlines_and_anchor_not_found(wd, capsys):
    code, out = run_word(capsys, "write", "--text", "첫줄\\n둘째줄")
    assert "첫줄" in doc_paras(wd["doc"]) and "둘째줄" in doc_paras(wd["doc"])
    code, out = run_word(capsys, "write", "--text", "x", "--anchor", "없는문구")
    assert code == 1 and "read" in out["message"]


def test_write_needs_content(wd, capsys):
    code, out = run_word(capsys, "write")
    assert code == 1 and out["error_type"] == "input"


def test_replace_no_infinite_loop_and_counts(wd, capsys):
    wd["doc"] = FakeWordDocument("가 가 가")
    code, out = run_word(capsys, "replace", "--find", "가", "--replace", "가가")
    assert code == 0 and out["replaced"] == 3 and wd["doc"].text.startswith("가가 가가 가가")
    assert out["agent_directive"].startswith("VERIFIED:")              # 바꾼 글이 찾는 글을 품어도 확인 (R1)


def test_replace_empty_deletes_and_formatting(wd, capsys):
    wd["doc"] = FakeWordDocument("초안 보고서 초안")
    code, out = run_word(capsys, "replace", "--find", "초안 ", "--replace", "")
    assert out["replaced"] == 1 and wd["doc"].text.startswith("보고서 초안")
    code, out = run_word(capsys, "replace", "--find", "보고서", "--bold", "--color", "blue")
    assert out["replaced"] == 1 and out["agent_directive"].startswith("VERIFIED:")
    assert any(len(e) == 3 and e[1] == "Bold" and e[2] is True for e in wd["doc"].log)


def test_replace_map_and_limits(wd, capsys):
    wd["doc"] = FakeWordDocument("Draft 진행중")
    code, out = run_word(capsys, "replace", "--map", '{"Draft": "Final", "진행중": "완료"}')
    assert out["replaced"] == 2 and wd["doc"].text.startswith("Final 완료") and out["agent_directive"].startswith("VERIFIED:")
    code, out = run_word(capsys, "replace", "--find", "x" * 300, "--replace", "y")
    assert code == 1 and "255" in out["message"]
    code, out = run_word(capsys, "replace", "--find", "없음", "--replace", "y")
    assert code == 0 and out["replaced"] == 0


def test_read_closed_docx_with_python_docx(tmp_path, capsys, monkeypatch):
    from sample_docs import new_docx
    p = tmp_path / "계획.docx"
    d = new_docx()
    d.add_heading("배경", level=1)
    d.add_paragraph("본문입니다")
    d.save(p)
    fake, calls = make_win32com()
    monkeypatch.setattr(off, "win32com", fake)
    monkeypatch.setattr(off, "list_open", lambda kind: [])
    code, out = run_word(capsys, "read", "--target", str(p))
    assert code == 0 and [x["text"] for x in out["paragraphs"]] == ["배경", "본문입니다"]
    assert out["paragraphs"][1]["heading"] == "배경"
    assert not any(c[0] == "DispatchEx" for c in calls)
    code, out = run_word(capsys, "read", "--target", str(p), "--save", str(tmp_path / "o.md"))   # 1.4.48: 읽은 문서에서 옮기기
    assert "next_commands" not in out                                     # 넣기·옮기기는 confirm_commands (1.4.75)
    transfer = [c["command"] for c in out["confirm_commands"] if "transfer" in c["command"]][0]
    assert "--from-target" in transfer and str(p) in transfer, transfer
    argv = split_command(transfer)[2:]                    # 1.4.79: 앱은 --from-app/--to-app (--from/--to는 outlook 전용)
    assert argv[:3] == ["transfer", "--from-app", "word"] and argv[argv.index("--to-app") + 1] == "excel", argv
    assert "--from" not in argv and "--to" not in argv, argv


def test_highlight_color_suggestion(wd, capsys):
    code, out = run_word(capsys, "replace", "--find", "계획", "--highlight", "yelow")
    assert code == 1, out
    argv = recommended_argv(out["message"], "word")             # 1.4.79: [추천]은 전체 명령
    assert argv == ["replace", "--find", "계획", "--highlight", "yellow"], argv


# ---------------------------------------------------------------------------
# 표 안 앵커, 마지막 단락 앵커
# ---------------------------------------------------------------------------


def test_anchor_in_table_cell_inserts_after_table_not_next_cell(wd, capsys):
    wd["doc"] = _doc_with_table()
    code, out = run_word(capsys, "write", "--text", "비고 문단", "--anchor", "시험 온도", "--pos", "after")
    assert code == 0, out
    assert doc_paras(wd["doc"])[:5] == ["머리말", "시험 온도", "25도", "비고 문단", "표 뒤 문단"]
    assert "표 뒤" in out["preview"]


def test_anchor_in_table_cell_before_or_replace_is_refused(wd, capsys):
    wd["doc"] = _doc_with_table()
    before = wd["doc"].text
    for pos in ("before", "replace"):
        code, out = run_word(capsys, "write", "--text", "x", "--anchor", "시험 온도", "--pos", pos)
        assert code == 1 and "표 안" in out["message"] and wd["doc"].text == before


def test_start_of_document_that_begins_with_table_is_refused(wd, capsys):
    doc = FakeWordDocument("칸1\r칸2\r본문")
    doc.table_spans = [(0, doc.text.index("본문"))]
    wd["doc"] = doc
    code, out = run_word(capsys, "write", "--text", "제목", "--pos", "start")
    assert code == 1 and "표로 시작" in out["message"] and doc.text.startswith("칸1")


def test_anchor_on_last_paragraph_makes_new_paragraph(wd, capsys):
    code, out = run_word(capsys, "write", "--text", "끝에 붙일 문단", "--anchor", "일정", "--pos", "after")
    assert code == 0, out
    p = doc_paras(wd["doc"])
    assert "일정" in p and "끝에 붙일 문단" in p and p.index("끝에 붙일 문단") == p.index("일정") + 1


def test_replace_unverified_when_text_remains(wd, capsys, monkeypatch):
    wd["doc"] = FakeWordDocument("초안 보고서")
    monkeypatch.setattr(word, "_count_text", lambda *a, **k: 1)
    code, out = run_word(capsys, "replace", "--find", "초안", "--replace", "최종")
    assert code == 0 and out["agent_directive"].startswith("UNVERIFIED:") and "남아" in out["unverified"][0]


# ---------------------------------------------------------------------------
# 코드 블록 줄은 ** __를 굵게로 읽지 않고 그대로, 인라인 코드 ```x```는 펜스가 아님
# ---------------------------------------------------------------------------


def test_word_code_block_verbatim(wd, capsys, tmp_path):
    p = tmp_path / "c.md"
    p.write_text(CODE_MD, encoding="utf-8")
    code, out = run_word(capsys, "write", "--file", str(p))
    txt = doc_paras(wd["doc"])
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    assert "def __init__(self):" in txt and "    return x**2 + y**2" in txt, txt


def test_word_file_starting_with_code_block_verified(wd, capsys, tmp_path):
    p = tmp_path / "c2.md"
    p.write_text("```\nprint(a**2)\n```\n", encoding="utf-8")
    code, out = run_word(capsys, "write", "--file", str(p))
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out


# ---------------------------------------------------------------------------
# 단어 안 __
# ---------------------------------------------------------------------------


def test_word_text_intraword_underscore_literal(wd, capsys):
    code, out = run_word(capsys, "write", "--text", "파일 report__2024__final.xlsx 참조")
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    assert "파일 report__2024__final.xlsx 참조" in doc_paras(wd["doc"]), wd["doc"].text


# ---------------------------------------------------------------------------
# 제목 규칙
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("level", ["#", "##", "###", "####"])
def test_word_headings_same_rule_every_level(wd, capsys, tmp_path, level):
    code, out = _word_md(capsys, tmp_path, f"{level} __init__ 메서드\n\n본문 **강조**\n")
    paras = doc_paras(wd["doc"])
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    assert "init 메서드" in paras and "본문 강조" in paras, paras


# ---------------------------------------------------------------------------
# 연도와 1이 아닌 번호
# ---------------------------------------------------------------------------


def test_word_date_line_kept(wd, capsys):
    code, out = run_word(capsys, "write", "--text", "2026. 11. 3. 회의 결과를 보고합니다.")
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    assert "2026. 11. 3. 회의 결과를 보고합니다." in doc_paras(wd["doc"]), wd["doc"].text


def test_word_md_list_from_3_keeps_numbers(wd, capsys, tmp_path):
    code, out = _word_md(capsys, tmp_path, "3. 셋\n4. 넷\n")
    paras = doc_paras(wd["doc"])
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    assert "3. 셋" in paras and "4. 넷" in paras, paras


# ---------------------------------------------------------------------------
# 굵게 위치는 UTF-16 단위
# ---------------------------------------------------------------------------


def test_bold_after_emoji(u16_doc):
    w = word.Writer(u16_doc, 0, None)
    w.markdown(f"{EMOJI} **출시** 완료 {EMOJI}{EMOJI} **두번째**")
    assert _bold_texts(u16_doc) == ["출시", "두번째"]
    assert _back(u16_doc.text).startswith(f"{EMOJI} 출시 완료")


def test_bold_without_emoji_unchanged(u16_doc):
    w = word.Writer(u16_doc, 0, None)
    w.text("가 **나** 다")
    assert _bold_texts(u16_doc) == ["나"]
    assert word._u16(EMOJI + "a") == 3 and word._u16("가나") == 2


# ---------------------------------------------------------------------------
# '---'는 구분선, 맨 앞 YAML 머리말은 넣지 않음
# ---------------------------------------------------------------------------


def test_word_rule_is_border_not_page_break(wd, capsys):
    code, out = run_word(capsys, "write", "--text", "# 제목\n\n앞 문단\n\n---\n\n뒤 문단")
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    doc = wd["doc"]
    assert "\x0c" not in doc.text and "구분선" in out["inserted"] and "warnings" not in out
    paras = doc.text.split("\r")
    i = paras.index("앞 문단")
    assert paras[i + 1] == "" and paras[i + 2] == "뒤 문단"
    borders = [e for e in doc.log if len(e) == 3 and isinstance(e[0], tuple) and e[0][0] == "border"]
    assert (("border", doc.text.index("뒤 문단") - 1, word.WD_BORDER_BOTTOM), "LineStyle", 1) in borders


def test_word_front_matter_skipped_with_warning(wd, capsys):
    code, out = run_word(capsys, "write", "--text", "---\ntitle: 비밀 제목\nauthor: kim\n---\n본문 첫 줄")
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    assert "title" not in wd["doc"].text and "\x0c" not in wd["doc"].text and "본문 첫 줄" in wd["doc"].text
    assert out["warnings"] and "YAML 머리말(4줄" in out["warnings"][0]
    assert word._first_inserted("---\ntitle: x\n---\n---\n본문", "auto") == "본문"


# ---------------------------------------------------------------------------
# replace: 대소문자만 바꾸기, 상한, 바꿀 글·서식 없는 바꾸기 거절, 단락 범위 압축
# ---------------------------------------------------------------------------


def test_case_only_replace_is_verified(wd_office, capsys):
    code, out = run_main(word, capsys, "replace", "--find", "office", "--replace", "Office")
    assert code == 0 and wd_office["doc"].text.startswith("Office") and out["agent_directive"].startswith("VERIFIED:"), out


def test_word_replace_limit_is_unverified(wd_office, capsys):
    wd_office["doc"] = FakeWordDocument("가 " * 5002)
    code, out = run_main(word, capsys, "replace", "--find", "가", "--replace", "나")
    assert code == 0 and out["limit_reached"] is True and out["agent_directive"].startswith("UNVERIFIED:")


def test_word_compress_ranges():
    assert word._compress([7, 1, 2, 3, 9, 10]) == "1-3,7,9-10" and word._compress([]) == ""


def test_word_replace_without_replace_or_format_is_refused(wd_office, capsys):
    code, out = run_main(word, capsys, "replace", "--find", "office")
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:") and "--replace" in out["message"]
    code, out = run_main(word, capsys, "replace", "--find", "office", "--replace", "")
    assert code == 0 and out["replaced"] == 1


# ---------------------------------------------------------------------------
# 바꾼 것 없이 성공한 명령은 바꾼 것으로 표시하지 않음 (되돌리기 지점을 남기지 않고 0곳 바뀜으로 보고)
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("argv", [
    ["replace", "--find", "계획", "--replace", "계획"],
    ["replace", "--map", '{"계획": "계획"}'],
])
def test_word_replace_same_text_reports_zero(wd_doc, capsys, argv):
    code, out = run_main(word, capsys, *argv)
    assert code == 0 and out["replaced"] == 0 and out["unchanged_same_text"] == {"계획": 1}, out
    assert "0곳" in out["preview"] and "처리" not in out["preview"], out
    assert undo_points("word") == [] and wd_doc.Saved is True


def test_word_replace_same_text_with_format_still_changes(wd_doc, capsys):
    code, out = run_main(word, capsys, "replace", "--find", "계획", "--replace", "계획", "--bold")
    assert code == 0 and out["replaced"] == 1, out


# ---------------------------------------------------------------------------
# 빈 찾을 글은 모든 글자 사이에 바꿀 글을 끼워 넣음 -> 바꾸기 전에 STOP_RETRY
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("argv", EMPTY_FIND)
def test_word_empty_find_is_refused(wd_iphone, capsys, argv):
    before = wd_iphone.text
    code, out = run_main(word, capsys, "replace", *argv)
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:"), out
    assert "word.py write" in out["message"], out
    assert wd_iphone.text == before
    assert undo_points("word") == []


# ---------------------------------------------------------------------------
# 대소문자만 다른 바꾸기에서 이미 바꿀 글인 곳은 바꾼 곳으로 세지 않음
# ---------------------------------------------------------------------------


def test_word_case_only_already_replaced_is_noop(wd_iphone, capsys):
    code, out = run_main(word, capsys, "replace", "--find", "iphone", "--replace", "iPhone")
    assert code == 0 and out["replaced"] == 0 and out["unchanged_same_text"] == {"iphone": 2}, out
    assert "0곳" in out["preview"], out
    assert undo_points("word") == [] and wd_iphone.Saved is True


def test_word_counts_only_changed_occurrences(monkeypatch, capsys):
    doc = FakeWordDocument("iphone, iPhone, IPHONE")
    stub_com(monkeypatch)
    set_active(monkeypatch, object(), doc)
    code, out = run_main(word, capsys, "replace", "--find", "iphone", "--replace", "iPhone")
    assert code == 0 and out["replaced"] == 2 and out["unchanged_same_text"] == {"iphone": 1}, out
    assert out["agent_directive"].startswith("VERIFIED:"), out
    assert doc.text.rstrip("\r") == "iPhone, iPhone, iPhone"


def test_word_same_text_with_format_still_changes(wd_iphone, capsys):
    code, out = run_main(word, capsys, "replace", "--find", "iphone", "--replace", "iPhone", "--bold")
    assert code == 0 and out["replaced"] == 2, out


# ---------------------------------------------------------------------------
# Word 찾기의 ^ 글자
# ---------------------------------------------------------------------------


def test_word_find_escapes_caret():
    seen = {}

    class F:
        def ClearFormatting(self):
            pass

        def Execute(self, FindText="", **kw):
            seen["text"] = FindText
            return False

    rng = Bag()
    rng.Find = F()
    word._find(rng, "감사합니다^^ 10^6")
    assert seen["text"] == "감사합니다^^^^ 10^^6"


# ---------------------------------------------------------------------------
# Word Find 255자 한도는 ^를 ^^로 바꾼 길이로 세고, 긴 바꿀 글은 Word Find 없이 다시 셈
# ---------------------------------------------------------------------------


def test_word_long_replacement_containing_find_is_verified(wd_caret, capsys):
    rep = "제목: " + "가" * 300
    code, out = _run_word(capsys, "replace", "--find", "제목", "--replace", rep)
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    assert wd_caret["doc"].text.startswith(rep)


def test_word_find_limit_counts_escaped_carets(wd_caret, capsys):
    code, out = _run_word(capsys, "replace", "--find", "^" * 200, "--replace", "x")
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:"), out
    assert "255" in out["message"]


# ---------------------------------------------------------------------------
# read 범위: --pages / --paragraphs / --section
# ---------------------------------------------------------------------------


def test_read_pages_includes_paragraph_continued_from_previous_page(wd_pages, capsys):
    doc = PageDoc([("머리말", 1), ("1쪽 본문", 10), ("다음 쪽까지 이어지는 긴 단락", 10), ("2쪽 본문", 10), ("3쪽 본문", 10)])
    doc.page_starts = [0, doc.start_of("다음 쪽까지") + 4, doc.start_of("3쪽 본문")]
    wd_pages["doc"] = doc
    code, out = run_word(capsys, "read", "--pages", "2")
    assert code == 0, out
    assert texts(out) == ["다음 쪽까지 이어지는 긴 단락", "2쪽 본문"]
    assert [p["page"] for p in out["paragraphs"]] == [2, 2]
    assert [p["no"] for p in out["paragraphs"]] == [3, 4]          # 번호는 문서 전체 단락 번호 그대로
    assert out["page_total"] == 3 and out["count"] == 2


def test_read_pages_out_of_range_is_input_error(wd_pages, capsys):
    code, out = run_word(capsys, "read", "--pages", "9")
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:")
    assert "전체 1~1" in out["message"]


def test_read_paragraphs_open_ended_range(wd_pages, capsys):
    code, out = run_word(capsys, "read", "--paragraphs", "8-")
    assert code == 0, out
    assert texts(out) == ["용어 본문", "2. 시험", "시험 본문"]
    assert [p["no"] for p in out["paragraphs"]] == [8, 9, 10]


def test_read_paragraphs_bad_format(wd_pages, capsys):
    code, out = run_word(capsys, "read", "--paragraphs", "a-3")
    assert code == 1 and "--paragraphs 형식" in out["message"]


def test_read_section_by_number_includes_subsections_only(wd_pages, capsys):
    code, out = run_word(capsys, "read", "--section", "1.1")         # '1.1.1 세부'와 헷갈리지 않음
    assert code == 0, out
    assert texts(out) == ["1.1 범위", "범위 본문", "1.1.1 세부", "세부 본문"]
    assert "단락 3-6" in out["selection"]


def test_read_section_matches_numbering_with_trailing_dot(wd_pages, capsys):
    wd_pages["doc"] = PageDoc([("1. SCOPE", 1), ("a", 10), ("1.1. Scope", 2), ("b", 10), ("1.1.1. Data", 3), ("c", 10)])
    code, out = run_word(capsys, "read", "--section", "1.1")
    assert code == 0, out
    assert texts(out) == ["1.1. Scope", "b", "1.1.1. Data", "c"]


def test_read_section_ambiguous_then_occurrence(wd_pages, capsys):
    wd_pages["doc"] = PageDoc([("1. 장비", 1), ("검증 방법", 2), ("장비 검증", 10), ("2. 소프트웨어", 1), ("검증 방법", 2), ("SW 검증", 10)])
    code, out = run_word(capsys, "read", "--section", "검증 방법")
    assert code == 1 and "[후보] 1) " in out["message"] and "--occurrence" in out["message"]
    assert out["confirm_commands"][1]["command"].endswith("--occurrence 2"), out
    code, out = run_word(capsys, "read", "--section", "검증 방법", "--occurrence", "2")
    assert code == 0 and texts(out) == ["검증 방법", "SW 검증"]


def test_read_section_not_found_and_no_headings(wd_pages, capsys):
    code, out = run_word(capsys, "read", "--section", "부록")
    assert code == 1 and "제목 '부록'" in out["message"]
    wd_pages["doc"] = PageDoc([("본문만 있음", 10)])
    code, out = run_word(capsys, "read", "--section", "1")
    assert code == 1 and "제목" in out["message"] and "--pages" in out["message"]


def test_read_only_one_range_option(wd_pages, capsys):
    code, out = run_word(capsys, "read", "--pages", "1", "--section", "1.1")
    assert code == 1 and "하나만" in out["message"]


def test_read_section_truncated_gives_paragraph_range(wd_pages, capsys, monkeypatch):
    monkeypatch.setattr(word.rt, "DEFAULT_MAX_CHARS", 30)
    code, out = run_word(capsys, "read", "--section", "1")
    assert code == 0 and out["truncated"] is True
    nxt = out["next_commands"][0]["command"]
    shown = len(out["paragraphs"])
    assert f"--paragraphs '{shown + 1}-8'" in nxt                      # 1장 끝(단락 8)까지만 이어 읽기
    assert "--section 1 --save" in out["next_commands"][1]["command"]


# ---------------------------------------------------------------------------
# outline: 제목 수준·단락 번호·쪽
# ---------------------------------------------------------------------------


def test_outline_lists_levels_paragraph_numbers_and_pages(wd_pages, capsys):
    doc = PageDoc(SPEC)
    doc.page_starts = [0, doc.start_of("2. 시험")]
    wd_pages["doc"] = doc
    code, out = run_word(capsys, "outline")
    assert code == 0, out
    heads = out["headings"]
    assert [(h["paragraph"], h["level"], h["text"]) for h in heads][:3] == [(1, 1, "1. 개요"), (3, 2, "1.1 범위"), (5, 3, "1.1.1 세부")]
    assert heads[-1]["page"] == 2 and heads[0]["page"] == 1
    assert out["count"] == 5 and "--section" in out["agent_directive"]


def test_outline_without_headings_is_explicit_zero(wd_pages, capsys):
    wd_pages["doc"] = PageDoc([("본문", 10)])
    code, out = run_word(capsys, "outline")
    assert code == 0 and out["count"] == 0 and "0건" in out["agent_directive"]


# ---------------------------------------------------------------------------
# comment: 목록, 정확한 앵커에 추가, CSV 일괄 추가
# ---------------------------------------------------------------------------


def test_comment_list_and_zero(cd, capsys):
    code, out = run_word(capsys, "comment")
    assert code == 0 and out["count"] == 0 and "0건" in out["agent_directive"]
    doc = cd["doc"]
    doc.Comments.items.append(FakeComment(FakeRange(doc, 0, 5), "규격 4.2항 확인\r값 다름", author="김검토"))
    code, out = run_word(capsys, "comment")
    assert code == 0 and out["count"] == 1
    item = out["items"][0]
    assert item["author"] == "김검토" and item["text"] == "규격 4.2항 확인\n값 다름" and item["target_text"] == "시험 온도"


def test_comment_add_on_exact_anchor_text_and_verified(cd, capsys):
    code, out = run_word(capsys, "comment", "--action", "add", "--anchor", "습도", "--text", "규격에 습도 기준 없음")
    assert code == 0, out
    assert out["verified"] is True and out["agent_directive"].startswith("VERIFIED:")
    c = cd["doc"].Comments(1)
    assert c.Scope.Text == "습도" and c.Range.Text == "규격에 습도 기준 없음"


def test_comment_add_ambiguous_anchor_adds_nothing(cd, capsys):
    code, out = run_word(capsys, "comment", "--action", "add", "--anchor", "시험 온도", "--text", "확인")
    assert code == 1 and "[후보]" in out["message"]
    assert cd["doc"].Comments.Count == 0
    code, out = run_word(capsys, "comment", "--action", "add", "--anchor", "시험 온도", "--text", "확인", "--occurrence", "2")
    assert code == 0 and cd["doc"].Comments(1).Scope.Start == cd["doc"].text.index("시험 온도 기록")


def test_comment_add_from_csv_all_or_nothing(cd, capsys, tmp_path):
    bad = tmp_path / "메모.csv"
    bad.write_text("기준 글,메모 내용\n습도,기준 없음\n없는 문구,확인\n", encoding="utf-8")
    code, out = run_word(capsys, "comment", "--action", "add", "--file", str(bad))
    assert code == 1 and "하나도 달지 않았습니다" in out["message"] and "2번째 줄" in out["message"]
    assert "[추천]" not in out["message"] and "[후보]" not in out["message"]   # 여러 줄 오류에 재시도 표시를 섞지 않음
    assert cd["doc"].Comments.Count == 0
    good = tmp_path / "메모2.csv"
    good.write_text("습도,기준 없음\n시험 온도,4.2항과 다름,2\n", encoding="utf-8")
    code, out = run_word(capsys, "comment", "--action", "add", "--file", str(good))
    assert code == 0 and out["count"] == 2 and out["verified"] is True
    assert sorted(c.Range.Text for c in cd["doc"].Comments.items) == ["4.2항과 다름", "기준 없음"]


def test_comment_add_unverified_when_reread_differs(cd, capsys, monkeypatch):
    real_add = cd["doc"].Comments.Add
    monkeypatch.setattr(cd["doc"].Comments, "Add", lambda Range=None, Text="": real_add(Range=Range, Text=Text[:2]))
    code, out = run_word(capsys, "comment", "--action", "add", "--anchor", "습도", "--text", "규격에 습도 기준 없음")
    assert code == 0 and out["verified"] is False and out["agent_directive"].startswith("UNVERIFIED:")


def test_comment_add_requires_anchor_and_text(cd, capsys):
    code, out = run_word(capsys, "comment", "--action", "add", "--text", "메모")
    assert code == 1 and "--anchor" in out["message"]
    code, out = run_word(capsys, "comment", "--action", "add", "--anchor", "습도")
    assert code == 1 and "메모 내용" in out["message"]


# ---------------------------------------------------------------------------
# revisions: 변경 추적 목록
# ---------------------------------------------------------------------------


def test_revisions_list_types_and_zero(wd_pages, capsys):
    doc = wd_pages["doc"]
    doc.Revisions = _Revs([])
    doc.TrackRevisions = False
    code, out = run_word(capsys, "revisions")
    assert code == 0 and out["count"] == 0 and "0건" in out["agent_directive"]
    doc.Revisions = _Revs([_Rev(1, "추가된 문장"), _Rev(2, "지운 문장"), _Rev(99, "?")])
    doc.TrackRevisions = True
    code, out = run_word(capsys, "revisions")
    assert code == 0 and out["count"] == 3 and out["track_changes"] is True
    assert [i["type"] for i in out["items"]] == ["삽입", "삭제", "기타(99)"]
    assert out["items"][0]["page"] == 4 and out["items"][0]["text"] == "추가된 문장"


def test_read_pages_does_not_pull_previous_paragraph_across_blank_break(wd_pages, capsys):
    # 쪽 나누기만 있는 빈 단락 다음 쪽: 빈 단락 앞의 단락은 앞 쪽에서 끝났으므로 넣지 않음
    doc = PageDoc([("1쪽 마지막 문장", 10), ("\x0c", 10), ("2쪽 첫 문장", 10)])
    doc.page_starts = [0, doc.start_of("2쪽 첫 문장")]
    wd_pages["doc"] = doc
    code, out = run_word(capsys, "read", "--pages", "2")
    assert code == 0 and texts(out) == ["2쪽 첫 문장"]


def test_read_paragraphs_noncontiguous_truncated_continues_with_paragraphs(wd_pages, capsys, monkeypatch):
    monkeypatch.setattr(word.rt, "DEFAULT_MAX_CHARS", 20)
    code, out = run_word(capsys, "read", "--paragraphs", "1-3,9")
    assert code == 0 and out["truncated"] is True
    shown = [p["no"] for p in out["paragraphs"]]
    rest = [n for n in (1, 2, 3, 9) if n not in shown]
    assert f"--paragraphs '{word._compress(rest)}'" in out["next_commands"][0]["command"]


def test_comment_text_without_action_add_still_adds(cd, capsys):
    code, out = run_word(capsys, "comment", "--anchor", "습도", "--text", "확인 필요")
    assert code == 0 and out["command"] == "comment" and cd["doc"].Comments.Count == 1


# ---------------------------------------------------------------------------
# word 다시 읽기
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("md", [
    "---\n\n# 제목\n본문\n",
    "***\n\n본문\n",
    "---\ntitle: 보고서\n---\n\n본문\n",
    "| a\\|b | c |\n|---|---|\n| 1 | 2 |\n",
])
def test_word_md_first_block_verified(wd, capsys, tmp_path, md):
    p = tmp_path / "w.md"
    p.write_text(md, encoding="utf-8")
    code, out = run_word(capsys, "write", "--file", str(p))
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out


# ---------------------------------------------------------------------------
# 가드: 문서 글이 오류 로그에 남지 않음 (needs_user_decision / not_found / ambiguous 경로마다)
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("paras,section", [
    ([("1. ZQX 개요", 1), ("본문", 10), ("2. ZQX 결론", 1), ("본문", 10)], "부록"),                  # 사용 가능한 제목:
    ([("1. ZQX 개요", 1), ("본문", 10), ("2. ZQX 결론", 1), ("본문", 10)], "개요 계획"),             # 비슷한 제목 [후보]
    ([("ZQX 검증 방법", 2), ("a", 10), ("ZQX 검증 방법 2", 2), ("b", 10)], "검증 방법"),              # 제목 여러 개 [후보]
])
def test_guard_word_section(wd, capsys, isolated_dirs, monkeypatch, paras, section):
    wd["doc"] = PageDoc(paras)
    code, out = run_as_tool(word, capsys, monkeypatch, "word", "read", "--section", section)
    assert code == 1, out
    assert_not_logged(out, isolated_dirs)


# ---------------------------------------------------------------------------
# word.py fill: 자리({{이름}})와 표 칸(표2!R3C4)을 값 파일로 한 번에 (C4, 1.4.61)
# ---------------------------------------------------------------------------


@pytest.fixture
def form_doc(wd, monkeypatch):
    from fake_word import FakeGridTable
    doc = FakeWordDocument("시험 성적서\r업체: {{업체명}}\r시험일: {{시험일}}\r담당: {{업체명}} 품질팀")
    doc.grid_tables = [FakeGridTable([["항목", "값"], ["전압", ""]]),
                       FakeGridTable([["항목", "기준", "측정"], ["전압", "5 V", ""], ["전류", "3 A", "2.1 A"], ["무게", None, ""]])]
    wd["doc"] = doc
    points = []
    monkeypatch.setattr(off, "save_snapshot", lambda *a, **k: points.append(a) or None)
    return doc, points


def _vals(tmp_path, text, name="값.csv"):
    f = tmp_path / name
    f.write_text(text, encoding="utf-8")
    return str(f)


def test_fill_places_and_cells(form_doc, capsys, tmp_path):
    doc, points = form_doc
    f = _vals(tmp_path, "자리,값\n{{업체명}},한빛전자\n{{시험일}},2026-11-04\n표2!R2C3,5.1 V\n표1!R2C2,\"5,1\"\n")
    code, out = run_word(capsys, "fill", "--file", f)
    assert code == 0 and out["verified"] is True and out["agent_directive"].startswith("VERIFIED:"), out
    assert "업체: 한빛전자" in doc.text and "담당: 한빛전자 품질팀" in doc.text and "{{" not in doc.text
    assert doc.grid_tables[1].grid[1][2] == "5.1 V" and doc.grid_tables[0].grid[1][1] == "5,1"
    assert out["places"] == {"{{업체명}}": 2, "{{시험일}}": 1} and out["count"] == 4
    assert len(points) == 1                                              # 되돌리기 지점 하나


def test_fill_nonempty_cell_needs_decision_then_overwrite(form_doc, capsys, tmp_path):
    doc, _ = form_doc
    f = _vals(tmp_path, "{{업체명}},한빛전자\n표2!R3C3,2.5 A\n표2!R2C3,5.1 V\n")
    code, out = run_word(capsys, "fill", "--file", f)
    assert code == 1 and out["error_type"] == "needs_user_decision" and "표2!R3C3 [행 '전류', 열 '측정']='2.1 A'" in out["message"], out
    assert "{{업체명}}" in doc.text and doc.grid_tables[1].grid[1][2] == ""             # 아무것도 바꾸지 않음
    code, out = run_word(capsys, "fill", "--file", f, "--overwrite")
    assert code == 0 and doc.grid_tables[1].grid[2][2] == "2.5 A" and out["overwritten"] == 1, out


def test_fill_dry_run_changes_nothing(form_doc, capsys, tmp_path):
    doc, points = form_doc
    before = doc.text
    f = _vals(tmp_path, "{{업체명}},한빛전자\n표2!R3C3,2.5 A\n")
    code, out = run_word(capsys, "fill", "--file", f, "--dry-run")
    assert code == 0 and out["dry_run"] is True and doc.text == before and points == [], out
    assert any("덮어쓰기: 표2!R3C3" in p for p in out["planned"]) and any("{{업체명}} 2곳" in p for p in out["planned"])
    cmd = out["confirm_commands"][0]["command"]
    assert "--dry-run" not in cmd and "--overwrite" in cmd and " fill " in cmd
    assert out["confirm_commands"][0]["destructive"] is True  # 11차 검토 (1.4.75): 덮어쓰기·삭제·발송은 destructive (묻지 말고 해도 늘 물음)


@pytest.mark.parametrize("shell", ["powershell", "pwsh", "cmd", "bash"])
def test_fill_confirm_command_path_survives_each_shell(form_doc, capsys, tmp_path, monkeypatch, shell):
    """1.4.76: 값 파일 경로에 공백·$·'·()·%·끝 백슬래시 폴더가 있어도 confirm_commands를 그 셸로 실행하면 같은 경로."""
    doc, _ = form_doc
    monkeypatch.setenv("OFFICE_CLI_SHELL", shell)
    folder = tmp_path / "값 $x (1) it's 100%"
    folder.mkdir()
    f = _vals(folder, "{{업체명}},한빛 $전자\n표2!R3C3,2.5 A\n")
    code, out = run_word(capsys, "fill", "--file", f, "--dry-run")
    assert code == 0 and out["dry_run"] is True, out
    argv = shell_sim.run_shell(out["confirm_commands"][0]["command"], shell)
    assert argv[argv.index("--file") + 1] == f, (shell, out["confirm_commands"])
    code, out = run_word(capsys, *argv[2:])
    assert code == 0 and "업체: 한빛 $전자" in doc.text and doc.grid_tables[1].grid[2][2] == "2.5 A", out


@pytest.mark.parametrize("slots", [
    "업체명=한빛전자; 시험일=2026-11-04; 표2!R2C3=5.1 V",                # 자리 이름만 (회사 PC 2026-11-06 O5)
    "{{업체명}}=한빛전자\\n{{시험일}}=2026-11-04\\n표2!R2C3=5.1 V",   # {{ }}와 줄바꿈 구분
    '{"업체명": "한빛전자", "시험일": "2026-11-04", "표2!R2C3": "5.1 V"}',    # JSON 객체
])
def test_fill_slots_without_values_file(form_doc, capsys, slots):
    """1.4.77: 값 파일 없이 --slots로 바로 채움 (Gemma가 값 파일이 없어 --slots를 지어내고 실패했음)."""
    doc, points = form_doc
    code, out = run_word(capsys, "fill", "--slots", slots)
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    assert "업체: 한빛전자" in doc.text and "시험일: 2026-11-04" in doc.text and doc.grid_tables[1].grid[1][2] == "5.1 V"
    assert len(points) == 1


def test_fill_slots_dry_run_confirm_keeps_slots(form_doc, capsys):
    doc, _ = form_doc
    before = doc.text
    code, out = run_word(capsys, "fill", "--slots", "업체명=한빛 $전자; 시험일='2026;11'", "--dry-run")
    assert code == 0 and doc.text == before, out
    argv = split_command(out["confirm_commands"][0]["command"])
    assert argv[argv.index("--slots") + 1] == "업체명=한빛 $전자; 시험일='2026;11'" and "--dry-run" not in argv
    code, out = run_word(capsys, *argv[2:])
    assert code == 0 and "업체: 한빛 $전자" in doc.text and "시험일: 2026;11" in doc.text, out


def test_fill_slots_comma_separated_is_refused_with_full_command(form_doc, capsys):
    doc, _ = form_doc
    before = doc.text
    code, out = run_word(capsys, "fill", "--slots", "업체명=한빛전자, 시험일=2026-11-04")
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:") and doc.text == before, out
    cmd = out["message"].split("[추천] ", 1)[1]
    assert cmd.startswith("python ") and split_command(cmd)[split_command(cmd).index("--slots") + 1] == "업체명=한빛전자; 시험일=2026-11-04"
    code, out = run_word(capsys, "fill", "--slots", "업체명=한빛전자, 서울; 시험일=2026-11-04")   # 값 안의 쉼표는 그대로
    assert code == 0 and "업체: 한빛전자, 서울" in doc.text, out


@pytest.mark.parametrize("slots", ["업체명=한빛전자 시험일=2026-11-04", "업체명=한빛전자|시험일=2026-11-04"])
def test_fill_slots_space_or_bar_separated_is_refused(form_doc, capsys, slots):
    doc, _ = form_doc
    code, out = run_word(capsys, "fill", "--slots", slots)
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:"), out
    cmd = split_command(out["message"].split("[추천] ", 1)[1])
    assert cmd[cmd.index("--slots") + 1] == "업체명=한빛전자; 시험일=2026-11-04", cmd


def test_fill_slots_value_with_name_value_pairs_is_kept(form_doc, capsys):
    """1.4.77 검토: '업체명=V1=5, V2=12'는 자리 업체명의 값 'V1=5, V2=12' (자리를 쉼표로 나눈 것이 아님)."""
    doc, _ = form_doc
    code, out = run_word(capsys, "fill", "--slots", "업체명=V1=5, V2=12; 시험일=2026-11-04")
    assert code == 0 and "업체: V1=5, V2=12" in doc.text, out


@pytest.mark.parametrize("slots", ['{"업체명": {"a": 1}}', '{"업체명": true}', '{"업체명": [1, 2]}'])
def test_fill_slots_json_non_scalar_is_refused(form_doc, capsys, slots):
    doc, _ = form_doc
    before = doc.text
    code, out = run_word(capsys, "fill", "--slots", slots)
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:") and doc.text == before, out


@pytest.mark.parametrize("argv,needle", [
    ((), "채울 값이 없습니다"),
    (("--file", "C:\\work\\성적서양식.docx"), "--target"),                    # 양식 문서를 값 파일로 (회사 PC 2026-11-06 O5)
    (("--file", "a.csv", "--slots", "업체명=a"), "하나만"),
])
def test_fill_value_source_errors_explain_slots(form_doc, capsys, argv, needle):
    code, out = run_word(capsys, "fill", *argv)
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:") and needle in out["message"], out
    assert "--slots" in out["message"] and "한빛" not in out["message"]       # 예시 값을 실제 값으로 옮겨 쓰지 않게


def test_fill_bad_rows_and_missing_places_change_nothing(form_doc, capsys, tmp_path):
    doc, _ = form_doc
    before = doc.text
    code, out = run_word(capsys, "fill", "--file", _vals(tmp_path, "{{업체명}},a\n업체명,b\n표X,c\n"))
    assert code == 1 and "2번째 줄" in out["message"] and "3번째 줄" in out["message"] and doc.text == before, out
    code, out = run_word(capsys, "fill", "--file", _vals(tmp_path, "{{업체명}},a\n{{담당자}},b\n표9!R1C1,c\n표2!R4C2,d\n"))
    assert code == 1 and "{{담당자}}" in out["message"] and "{{시험일}}" in out["message"], out     # 문서에 있는 자리 안내
    assert "표 9번이 없습니다" in out["message"] and "표2!R4C2" in out["message"] and doc.text == before


def test_fill_third_column_stops_before_filling(form_doc, capsys, tmp_path):
    # 값의 쉼표가 열을 나눔: {{업체명}},서울시 강남구, 101호 -> 앞부분만 채우고 VERIFIED였음 (1.4.63)
    doc, points = form_doc
    before = doc.text
    code, out = run_word(capsys, "fill", "--file", _vals(tmp_path, "자리,값\n{{업체명}},서울시 강남구, 101호\n표2!R2C3,1,500\n"))
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:"), out
    assert "2, 3번째 줄" in out["message"] and "큰따옴표" in out["message"] and '{{업체명}},"서울시 강남구, 101호"' in out["message"]
    assert doc.text == before and doc.grid_tables[1].grid[1][2] == "" and undo_points("word") == []


def test_fill_preview_shows_row_and_column_heads(form_doc, capsys, tmp_path):
    # 병합 표에서 표2!R3C4가 다른 칸일 수 있음: 미리보기·결과에 그 행 첫 칸과 그 열 첫 행 글을 함께 (1.4.63)
    doc, _ = form_doc
    f = _vals(tmp_path, "표2!R2C3,5.1 V\n표2!R3C3,2.5 A\n")
    code, out = run_word(capsys, "fill", "--file", f, "--dry-run")
    assert code == 0, out
    assert any("표2!R2C3" in p and "행 '전압'" in p and "열 '측정'" in p for p in out["planned"]), out["planned"]
    assert any("표2!R3C3" in p and "행 '전류'" in p and "열 '측정'" in p for p in out["planned"])
    code, out = run_word(capsys, "fill", "--file", f)
    assert code == 1 and out["error_type"] == "needs_user_decision" and "행 '전류'" in out["message"], out
    code, out = run_word(capsys, "fill", "--file", f, "--overwrite")
    assert code == 0 and any("행 '전압'" in p and "열 '측정'" in p for p in out["applied"]), out


def test_fill_protected_document_changes_nothing(form_doc, capsys, tmp_path):
    doc, points = form_doc
    doc.ProtectionType = 2                                   # wdAllowOnlyFormFields
    before = doc.text
    code, out = run_word(capsys, "fill", "--file", _vals(tmp_path, "{{업체명}},한빛전자\n표2!R2C3,5.1 V\n"))
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:") and "보호" in out["message"], out
    assert doc.text == before and doc.grid_tables[1].grid[1][2] == "" and undo_points("word") == []


def test_fill_failure_midway_says_how_many_changed(form_doc, capsys, tmp_path, monkeypatch):
    doc, points = form_doc
    from fake_word import FakeCellRange
    real = FakeCellRange.Text

    def refuse(s, v):
        if v == "막힘":
            raise RuntimeError("com_error: 이 명령을 사용할 수 없습니다")
        real.fset(s, v)
    monkeypatch.setattr(FakeCellRange, "Text", property(real.fget, refuse))
    code, out = run_word(capsys, "fill", "--file", _vals(tmp_path, "{{업체명}},한빛전자\n표1!R2C2,x\n표2!R2C3,막힘\n"))
    assert code == 1 and out["status"] == "error", out
    assert "앞의 3곳" in out["message"] and "office.py undo" in out["message"]
    assert len(points) == 1


def test_fill_same_value_cell_is_unchanged(form_doc, capsys, tmp_path):
    doc, points = form_doc
    code, out = run_word(capsys, "fill", "--file", _vals(tmp_path, "표2!R3C3,2.1 A\n"))
    assert code == 0 and out["unchanged"] == 1 and out["count"] == 0, out


def test_fill_reads_back(form_doc, capsys, tmp_path, monkeypatch):
    doc, _ = form_doc
    from fake_word import FakeCellRange
    monkeypatch.setattr(FakeCellRange, "Text", property(lambda s: s.grid[s.r][s.c] + "\r\x07",
                                                        lambda s, v: s.grid[s.r].__setitem__(s.c, v[:2])))
    code, out = run_word(capsys, "fill", "--file", _vals(tmp_path, "표2!R2C3,5.1 V\n"))
    assert code == 0 and out["verified"] is False and "표2!R2C3" in out["unverified"][0], out


@pytest.mark.parametrize("doc_place, slot", [("{{업체명}}", "{{ 업체명 }}"), ("{{ 업체명 }}", "업체명"), ("{{업체 명}}", "{{업체명}}")])
def test_fill_place_differing_only_in_spaces_is_filled_and_reported(wd, monkeypatch, capsys, doc_place, slot):
    """괄호 안 띄어쓰기만 다른 자리는 문서의 자리 하나와만 맞으면 그 자리로 채우고 recovered로 알림 (검토 1.4.77에서 남긴 것)."""
    doc = FakeWordDocument(f"업체: {doc_place}\r시험일: {{{{시험일}}}}")
    wd["doc"] = doc
    monkeypatch.setattr(off, "save_snapshot", lambda *a, **k: None)
    code, out = run_word(capsys, "fill", "--slots", f"{slot}=한빛전자; 시험일=2026-11-09")
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out.get("message")
    assert "업체: 한빛전자" in doc.text and doc_place in " ".join(out["recovered"]), out


def test_fill_place_spaces_with_two_matches_is_not_guessed(wd, monkeypatch, capsys):
    doc = FakeWordDocument("업체: {{ 업체명}}\r담당: {{업체명 }}")
    wd["doc"] = doc
    monkeypatch.setattr(off, "save_snapshot", lambda *a, **k: None)
    code, out = run_word(capsys, "fill", "--slots", "업체명=한빛전자")
    assert code == 1 and out["status"] == "error" and "{{ 업체명}}" in doc.text and "{{업체명 }}" in doc.text, out


def test_fill_slots_newline_inside_value_is_kept(form_doc, capsys):
    """'주소=서울\\n강남구': 다음 줄이 '자리='가 아니면 값 안의 줄바꿈 (예전: '강남구' 항목이 잘못되었다며 멈춤)."""
    doc, _ = form_doc
    code, out = run_word(capsys, "fill", "--slots", "업체명=한빛\\n전자; 시험일=2026-11-04")
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    assert "업체: 한빛\v전자" in doc.text and "시험일: 2026-11-04" in doc.text


def test_read_empty_document_is_explicit_zero(wd_pages, capsys):
    """1.4.81 빈 결과 명시: 필터 없이 읽은 빈 문서도 count 0 (전에는 빈 paragraphs만)."""
    wd_pages["doc"] = PageDoc([("", 10), (" ", 10)])     # 빈 단락만 (실제 빈 문서도 빈 단락 하나)
    code, out = run_word(capsys, "read")
    assert code == 0 and out["count"] == 0 and out["paragraphs"] == [], (out.get("message"), out)
    assert out["agent_directive"].startswith("VERIFIED: 0건.") and "찾은 항목 없음" in out["agent_directive"], out


# ---------------------------------------------------------------------------
# replace / write --dry-run: 문서를 바꾸지 않고 바뀔 곳만 보여 주고, confirm_commands[0] = 같은 명령에서 --dry-run만 뺀 것
# ---------------------------------------------------------------------------


def _confirm_argv(out):
    return split_command(out["confirm_commands"][0]["command"])[2:]


def test_replace_dry_run_changes_nothing_then_confirm_replaces(wd_doc, capsys):
    before = wd_doc.text
    argv = ["replace", "--find", "계획", "--replace", "목표", "--bold"]
    code, out = run_main(word, capsys, *argv, "--dry-run")
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:") and out["dry_run"] is True, out
    assert out["count"] == 1 and out["by_text"] == {"계획": 1}, out
    assert out["preview"].startswith("미리보기:") and out["preview"].endswith("(바꾸지 않음)"), out
    assert out["planned"] == ["'계획': 2026년 계획입니다. -> 2026년 목표입니다. (굵게)"] and out["truncated"] is False, out
    assert out["confirm_commands"][0]["destructive"] is True and _confirm_argv(out) == argv, out
    assert wd_doc.text == before and wd_doc.Saved is True and undo_points("word") == []
    code, out = run_main(word, capsys, *_confirm_argv(out))
    assert code == 0 and out["replaced"] == 1 and "2026년 목표입니다." in wd_doc.text, out
    assert len(undo_points("word")) == 1                  # 실제 실행은 되돌리기 지점을 남김 (미리보기 검사가 헛돌지 않음)


def test_replace_dry_run_zero_and_same_text(wd_doc, capsys):
    code, out = run_main(word, capsys, "replace", "--find", "없는글", "--replace", "x", "--dry-run")
    assert code == 0 and out["count"] == 0 and "바뀔 것 없음" in out["preview"] and "confirm_commands" not in out, out
    assert "count 0" in out["agent_directive"] and out["preview"].endswith("(바꾸지 않음)"), out
    code, out = run_main(word, capsys, "replace", "--find", "계획", "--replace", "계획", "--dry-run")
    assert code == 0 and out["count"] == 0 and out["unchanged_same_text"] == {"계획": 1}, out
    assert undo_points("word") == [] and wd_doc.Saved is True


def test_replace_dry_run_planned_is_capped(wd, capsys):
    wd["doc"] = FakeWordDocument("가 " * 60)
    before = wd["doc"].text
    code, out = run_main(word, capsys, "replace", "--map", '{"가": "나"}', "--dry-run")
    assert code == 0 and out["count"] == 60 and len(out["planned"]) == 50 and out["truncated"] is True, out
    assert wd["doc"].text == before
    assert _confirm_argv(out) == ["replace", "--map", '{"가": "나"}']


def test_write_replace_dry_run_shows_removed_paragraph_then_confirm(wd_doc, capsys):
    before = wd_doc.text
    argv = ["write", "--text", "새 계획", "--anchor", "2026년", "--pos", "replace"]
    code, out = run_main(word, capsys, *argv, "--dry-run")
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:") and out["dry_run"] is True and out["count"] == 1, out
    assert out["planned"] == ["지우기: 단락 '2026년 계획입니다.'", "넣기: 글 '새 계획'"], out
    assert out["preview"].startswith("미리보기:") and out["preview"].endswith("(바꾸지 않음)") and "지워질 단락" in out["preview"], out
    assert out["confirm_commands"][0]["destructive"] is True and _confirm_argv(out) == argv, out
    assert wd_doc.text == before and wd_doc.Saved is True and undo_points("word") == []
    code, out = run_main(word, capsys, *_confirm_argv(out))
    assert code == 0 and "2026년 계획입니다." not in wd_doc.text and "새 계획" in doc_paras(wd_doc), out


@pytest.mark.parametrize("extra", [[], ["--anchor", "일정", "--pos", "after"], ["--anchor", "추진", "--pos", "before"]])
def test_write_dry_run_other_positions_change_nothing(wd_doc, capsys, extra):
    before = wd_doc.text
    argv = ["write", "--text", "추가", *extra]
    code, out = run_main(word, capsys, *argv, "--dry-run")
    assert code == 0 and out["dry_run"] is True and out["planned"] == ["넣기: 글 '추가'"], out
    assert "destructive" not in out["confirm_commands"][0] and _confirm_argv(out) == argv, out
    assert wd_doc.text == before and undo_points("word") == []


def test_write_dry_run_anchor_errors_same_as_real_run(wd, capsys):
    wd["doc"] = _doc_with_table()
    code, out = run_main(word, capsys, "write", "--text", "x", "--anchor", "시험 온도", "--pos", "replace", "--dry-run")
    assert code == 1 and "표 안에" in out["message"], out
    code, out = run_main(word, capsys, "write", "--text", "x", "--anchor", "없는문구", "--pos", "replace", "--dry-run")
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:"), out
