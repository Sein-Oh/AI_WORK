"""ppt.py: 슬라이드 배치, 마크다운 -> 슬라이드, replace, 슬라이드 관리, 노트, 화면 이동 규칙."""
import types

import pytest

import _office as off
import ppt
from fake_office import make_win32com
from fake_ppt import FakePptWindowApp, FakePresentation, FakeShape
from office_helpers import Bag, CODE_MD, EMPTY_FIND, recommended_argv, run_main, run_ppt, set_active, stub_com, undo_points, split_command


def live(slide):
    return [s for s in slide.Shapes.items if not s.deleted]


def bullet(shape, i):
    return shape.TextFrame.TextRange.paras[i].ParagraphFormat.Bullet


class com_error(Exception):
    """pywintypes.com_error 흉내 (args: hresult, 설명, excepinfo, argerr)."""


E_FAIL = (-2147352567, "예외가 발생했습니다.", (0, None, None, None, 0, -2147467259), None)


def _pdf_export(pp, behavior, version="16.0"):
    calls = []

    def export(path, fmt, **kw):
        calls.append((path, fmt, kw))
        if "PrintRange" not in kw:      # 실제 pywin32: 생략하면 인자 변환 단계에서 TypeError
            raise TypeError("The Python instance can not be converted to a COM object")
        if behavior in ("ok", "empty"):
            with open(path, "wb") as f:
                f.write(b"%PDF" if behavior == "ok" else b"")
        else:
            raise behavior
    pp["prs"].ExportAsFixedFormat = export
    pp["app"].Version = version
    return calls


def _title(prs, n):
    return ppt.slide_units(prs.Slides(n), n)["title"]


@pytest.fixture
def pp_prs(monkeypatch):
    prs = FakePresentation(3)
    app = types.SimpleNamespace(ActiveWindow=types.SimpleNamespace(View=types.SimpleNamespace(GotoSlide=lambda n: None),
                                                                   Presentation=prs))
    stub_com(monkeypatch)
    set_active(monkeypatch, app, prs)
    return prs


def _u16(s):
    return len(s.encode("utf-16-le")) // 2


class U16TextRange:
    """PowerPoint처럼 After/Start를 UTF-16 단위로 다루는 가짜 TextRange (Text는 파이썬 문자열)."""

    def __init__(self, text):
        self.Text = text

    def Replace(self, FindWhat, ReplaceWhat, After=0, MatchCase=0, WholeWords=0):
        units, start = 0, len(self.Text)
        for i, ch in enumerate(self.Text):
            if units >= After:
                start = i
                break
            units += _u16(ch)
        hay = self.Text if MatchCase else self.Text.lower()
        idx = hay.find(FindWhat if MatchCase else FindWhat.lower(), start)
        if idx < 0:
            return None
        self.Text = self.Text[:idx] + ReplaceWhat + self.Text[idx + len(FindWhat):]
        found = Bag()
        found.Start = _u16(self.Text[:idx]) + 1
        found.Font = Bag()
        return found


def _texts(pp):
    prs = pp["prs"]
    s = prs.Slides(prs.Slides.Count)
    return [sh.TextFrame.TextRange.Text for sh in live(s) if hasattr(sh, "TextFrame")]


def _ppt_md(capsys, tmp_path, md, *extra):
    p = tmp_path / "s.md"
    p.write_text(md, encoding="utf-8")
    return run_ppt(capsys, "write", "--file", str(p), *extra)


# ---------------------------------------------------------------------------
# write / replace / slide / read 기본 동작과 화면 이동
# ---------------------------------------------------------------------------


def test_write_new_slide_title_placeholder_and_bullets(pp, capsys):
    code, out = run_ppt(capsys, "write", "--slide", "new", "--title", "추진 전략", "--text", "- 안정성\\n- 확장성\\n설명")
    assert code == 0, out
    prs = pp["prs"]
    assert prs.Slides.Count == 4 and out["slides_written"] == [4]
    s = prs.Slides(4)
    assert s.Shapes.Title.TextFrame.TextRange.Text == "추진 전략"
    body = live(s)[1]
    assert body.TextFrame.TextRange.Text == "안정성\r확장성\r설명"
    assert body.Top >= 20 + 60                                     # 제목 아래에 쌓임
    assert pp["app"].ActiveWindow.View.goto == [4]                 # 고친 슬라이드로만 이동
    assert out["agent_directive"].startswith("VERIFIED:") and out["verified"] is True   # 다시 읽어 확인 (R1)


def test_write_bullets_set_type_explicitly(pp, capsys):
    # Visible만 켜면 글상자에 따라 번호 목록(Type=2)이 되므로 종류를 명시해야 함 (실기 F3)
    code, out = run_ppt(capsys, "write", "--slide", "new", "--title", "목록", "--text", "- 안정성\\n1. 확장성\\n설명")
    assert code == 0, out
    body = live(pp["prs"].Slides(4))[1]
    assert body.TextFrame.TextRange.Text == "안정성\r확장성\r설명"
    assert bullet(body, 1).Visible == -1 and bullet(body, 1).Type == ppt.PP_BULLET_UNNUMBERED
    assert bullet(body, 2).Visible == -1 and bullet(body, 2).Type == ppt.PP_BULLET_NUMBERED
    assert bullet(body, 3).Visible == 0 and "Type" not in vars(bullet(body, 3))


def test_info_detail_reports_lists_outside_overlap_and_overflow(pp, capsys):
    run_ppt(capsys, "write", "--slide", "new", "--title", "배치", "--text", "- 가\\n1. 나\\n다")
    s = pp["prs"].Slides(4)
    body = live(s)[1]
    tr = body.TextFrame.TextRange
    tr.Paragraphs().Count = 3
    tr.BoundHeight = body.Height + 50                                       # 글자가 상자보다 큼
    s.Shapes.add(FakeShape("그림 밖", left=900, top=100, width=200, height=100, kind=13))   # 슬라이드 너비 960
    s.Shapes.add(FakeShape("겹친 상자", left=body.Left, top=body.Top, width=body.Width, height=body.Height, text="겹침"))
    code, out = run_ppt(capsys, "info", "--detail", "--slides", "4")
    assert code == 0 and [x["no"] for x in out["slides"]] == [4] and out["slides_total"] == 4
    det = {d["name"]: d for d in out["slides"][0]["shape_detail"]}
    assert det[body.Name]["lists"] == ["bullet", "numbered", "none"] and det[body.Name]["text_overflow"] is True
    assert det["그림 밖"]["kind"] == "picture" and det["그림 밖"]["outside"] is True
    warnings = " | ".join(out["slides"][0]["warnings"])
    assert "슬라이드 밖" in warnings and "넘칩니다" in warnings and "겹칩니다" in warnings
    assert out["layout_warnings"] == len(out["slides"][0]["warnings"])


def test_markdown_bullets_and_numbered_lists_keep_their_kind():
    sections = ppt._md_sections("## 제목\n\n- 가\n- 나\n\n1. 첫째\n2. 둘째\n")
    assert sections[0]["lines"] == [("가", "bullet"), ("나", "bullet"), ("첫째", "numbered"), ("둘째", "numbered")]


def test_write_markdown_file_makes_one_slide_per_heading(pp, capsys, tmp_path):
    md = tmp_path / "결과.md"
    md.write_text("# 검색 결과\n\n## 보고서.pdf > p.3\n\n**마일스톤** 일정 정리\n\n## 계획.docx > 단락 2\n\n- 3분기 적용\n\n"
                  "| 항목 | 값 |\n|---|---|\n| A | 1 |\n", encoding="utf-8")
    code, out = run_ppt(capsys, "write", "--file", str(md))
    assert code == 0
    prs = pp["prs"]
    assert out["slides_written"] == [4, 5, 6] and prs.Slides.Count == 6
    assert prs.Slides(5).Shapes.Title.TextFrame.TextRange.Text == "보고서.pdf > p.3"
    assert live(prs.Slides(5))[1].TextFrame.TextRange.Text == "마일스톤 일정 정리"      # ** 제거
    table = [s for s in live(prs.Slides(6)) if s.HasTable][0].Table
    assert table.Cell(2, 1).Shape.TextFrame.TextRange.Text == "A"
    assert out["agent_directive"].startswith("VERIFIED:")


def test_write_long_csv_table_splits_slides_with_header(pp, capsys, tmp_path):
    f = tmp_path / "t.csv"
    f.write_text("이름,값\n" + "".join(f"항목{i},{i}\n" for i in range(1, 21)), encoding="utf-8")
    code, out = run_ppt(capsys, "write", "--file", str(f), "--title", "결과")
    assert out["slides_written"] == [4, 5]
    t2 = [s for s in live(pp["prs"].Slides(5)) if s.HasTable][0].Table
    assert t2.Cell(1, 1).Shape.TextFrame.TextRange.Text == "이름" and t2.Rows.Count == 6
    assert out["agent_directive"].startswith("VERIFIED:")


def test_write_image_fits_remaining_space_keeping_ratio(pp, capsys, tmp_path):
    img = tmp_path / "chart.png"
    img.write_bytes(b"\x89PNG")
    code, out = run_ppt(capsys, "write", "--slide", "2", "--image", str(img))
    assert code == 0
    pic = [s for s in live(pp["prs"].Slides(2)) if s.Type == 13][0]
    assert pic.Width / pic.Height == pytest.approx(1600 / 900)
    assert pic.Top + pic.Height <= 540 - ppt.MARGIN + 0.01 and pic.Left + pic.Width <= 960


def test_write_bad_slide_number_and_no_content(pp, capsys):
    code, out = run_ppt(capsys, "write", "--slide", "9", "--text", "x")
    assert code == 1 and "new" in out["message"]
    code, out = run_ppt(capsys, "write")
    assert code == 1


def test_cards(pp, capsys):
    code, out = run_ppt(capsys, "write", "--slide", "new", "--title", "핵심", "--cards", "안정성:무중단; 확장성:모듈")
    cards = [s for s in live(pp["prs"].Slides(4)) if s.Name.startswith("Shape")]
    assert len(cards) == 2 and cards[0].TextFrame.TextRange.Text == "안정성\r무중단"


def test_replace_in_shapes_tables_no_loop(pp, capsys):
    prs = pp["prs"]
    s = prs.Slides(1)
    s.Shapes.add(FakeShape("tb", text="초안 초안 진행중"))
    tbl = s.Shapes.AddTable(1, 1, 0, 0, 10, 10)
    tbl.Table.Cell(1, 1).Shape.TextFrame.TextRange.Text = "초안"
    code, out = run_ppt(capsys, "replace", "--find", "초안", "--replace", "초안본", "--bold")
    assert code == 0 and out["replaced"] == 3 and out["slides"] == [1]
    assert out["agent_directive"].startswith("VERIFIED:")                  # 바꾼 글이 찾는 글을 품음 -> 바꾼 글 확인
    assert live(s)[1].TextFrame.TextRange.Text == "초안본 초안본 진행중"
    code, out = run_ppt(capsys, "replace", "--find", "진행중", "--replace", "")
    assert out["replaced"] == 1 and live(s)[1].TextFrame.TextRange.Text == "초안본 초안본 "
    assert out["agent_directive"].startswith("VERIFIED:")                  # 다시 읽어 '진행중'이 없음 (R1)


def test_read_does_not_move_view(pp, capsys):
    s = pp["prs"].Slides(2)
    s.Shapes.Title.TextFrame.TextRange.Text = "둘째"
    s.Shapes.add(FakeShape("tb", text="본문 내용"))
    code, out = run_ppt(capsys, "read", "--slides", "2")
    assert code == 0 and out["slides"][0]["title"] == "둘째" and out["slides"][0]["text"] == "본문 내용"
    code, out = run_ppt(capsys, "info")
    assert out["slides_total"] == 3
    assert pp["app"].ActiveWindow.View.goto == []


def test_slide_actions(pp, capsys):
    prs = pp["prs"]
    first = prs.Slides(1)
    code, out = run_ppt(capsys, "slide", "--action", "move", "--slide", "1", "--dest", "3")
    assert code == 0 and prs.Slides(3) is first and out["agent_directive"].startswith("VERIFIED:")
    code, out = run_ppt(capsys, "slide", "--action", "duplicate", "--slide", "3")
    assert out["slides_total"] == 4 and out["agent_directive"].startswith("VERIFIED:")     # 장 수를 다시 읽음 (R1)
    code, out = run_ppt(capsys, "slide", "--action", "clear")
    assert code == 1 and "--all-slides" in out["message"]
    prs.Slides(2).Shapes.add(FakeShape("x", text="a"))
    code, out = run_ppt(capsys, "slide", "--action", "clear", "--slide", "2")
    assert code == 0 and prs.Slides(2).Shapes.Count == 0 and out["agent_directive"].startswith("VERIFIED:")
    code, out = run_ppt(capsys, "slide", "--action", "delete", "--slide", "9")
    assert code == 1


def test_read_closed_pptx_uses_python_pptx(tmp_path, capsys, monkeypatch):
    from sample_docs import new_pptx
    p = tmp_path / "d.pptx"
    prs = new_pptx()
    s = prs.slides.add_slide(prs.slide_layouts[5])
    s.shapes.title.text = "제목"
    prs.save(p)
    fake, calls = make_win32com()
    monkeypatch.setattr(off, "win32com", fake)
    monkeypatch.setattr(off, "list_open", lambda kind: [])
    code, out = run_ppt(capsys, "read", "--target", str(p))
    assert code == 0 and out["slides"][0]["title"] == "제목"
    assert not calls
    code, out = run_ppt(capsys, "read", "--target", str(p), "--save", str(tmp_path / "o.md"))   # 1.4.48: 읽은 문서에서 옮기기
    assert "next_commands" not in out                                     # 넣기·옮기기는 confirm_commands (1.4.75)
    transfer = [c["command"] for c in out["confirm_commands"] if "transfer" in c["command"]][0]
    assert "--from-target" in transfer and str(p) in transfer, transfer
    argv = split_command(transfer)[2:]                    # 1.4.79: 앱은 --from-app/--to-app (--from/--to는 outlook 전용)
    assert argv[:3] == ["transfer", "--from-app", "ppt"] and argv[argv.index("--to-app") + 1] == "excel", argv
    assert "--from" not in argv and "--to" not in argv, argv


def test_export_pdf_passes_print_range_none(pp, capsys):
    calls = _pdf_export(pp, "ok")
    code, out = run_ppt(capsys, "export", "--save", "deck.pdf")
    assert code == 0 and out["output_file"].endswith("deck.pdf")
    assert calls[0][1] == 2 and calls[0][2] == {"PrintRange": None}
    assert out["agent_directive"].startswith("VERIFIED:") and out["verified"] is True     # 저장한 파일 확인 (R1)


def test_export_empty_file_is_unverified(pp, capsys):
    _pdf_export(pp, "empty")
    code, out = run_ppt(capsys, "export", "--save", "deck.pdf")
    assert code == 0 and out["agent_directive"].startswith("UNVERIFIED:") and "0바이트" in out["unverified"][0]


def test_export_pdf_on_office_without_pdf_is_input_error(pp, capsys):
    # PDF 내보내기가 없는 Office(16 미만)는 내보내기를 시도하지 않고 바로 안내
    calls = _pdf_export(pp, com_error(*E_FAIL), version="11.0")
    code, out = run_ppt(capsys, "export", "--save", "deck.pdf")
    assert code == 1 and out["error_type"] == "input" and calls == []
    assert out["agent_directive"].startswith("STOP_RETRY:") and ".pptx" in out["message"] and "2007" not in out["message"]


def test_export_pdf_e_fail_on_newer_office_is_not_relabelled(pp, capsys):
    _pdf_export(pp, com_error(*E_FAIL), version="16.0")
    code, out = run_ppt(capsys, "export", "--save", "deck.pdf")
    assert code == 1 and out["error_type"] == "office" and "PDF 내보내기 기능이 없습니다" not in out["message"]


# ---------------------------------------------------------------------------
# write: .md 첫 제목은 첫 슬라이드로, 기존 제목은 --overwrite, 잘린 read는 남은 슬라이드로 이어 읽기
# ---------------------------------------------------------------------------


def test_title_with_markdown_file_goes_on_first_slide(pp, capsys, tmp_path):
    md = tmp_path / "내용.md"
    md.write_text("## 첫 절\n- 가\n\n## 둘째 절\n- 나\n", encoding="utf-8")
    code, out = run_main(ppt, capsys, "write", "--slide", "new", "--title", "발표 제목", "--file", str(md))
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    prs = pp["prs"]
    assert _title(prs, 4) == "발표 제목" and "첫 절" in ppt.slide_units(prs.Slides(4), 4)["text"]
    assert _title(prs, 5) == "둘째 절"


def test_existing_title_needs_overwrite(pp, capsys):
    prs = pp["prs"]
    prs.Slides(1).Shapes.Title.TextFrame.TextRange.Text = "기존 제목"
    code, out = run_main(ppt, capsys, "write", "--slide", "1", "--title", "새 제목")
    assert code == 1 and out["error_type"] == "needs_user_decision" and "--overwrite" in out["agent_directive"]
    assert _title(prs, 1) == "기존 제목"
    code, out = run_main(ppt, capsys, "write", "--slide", "1", "--title", "새 제목", "--overwrite")
    assert code == 0 and _title(prs, 1) == "새 제목"
    code, out = run_main(ppt, capsys, "write", "--slide", "2", "--title", "빈 슬라이드 제목")      # 빈 제목은 묻지 않음
    assert code == 0 and _title(prs, 2) == "빈 슬라이드 제목"
    code, out = run_main(ppt, capsys, "write", "--slide", "1", "--title", "새 제목", "--text", "- 본문")   # 같은 제목
    assert code == 0


def test_truncated_read_continues_with_remaining_selected_slides(pp, capsys):
    prs = FakePresentation(9)
    for n in (1, 5, 9):
        prs.Slides(n).Shapes.add(FakeShape("본문", 36, 100, 800, 300, text="가" * 5000))
    pp["prs"] = prs
    code, out = run_main(ppt, capsys, "read", "--slides", "1,5,9")
    assert code == 0 and out["truncated"] is True and [s["no"] for s in out["slides"]] == [1]
    assert out["next_commands"][0]["command"].endswith("--slides '5,9'")          # 이전: 2-9


# ---------------------------------------------------------------------------
# PPT 제목 확인은 제목 개체틀만, .md 첫 제목이 --title과 같으면 본문에 다시 넣지 않음
# ---------------------------------------------------------------------------


def test_slide_without_title_placeholder_does_not_ask(pp, capsys):
    prs = pp["prs"]
    s = prs.Slides.Add(4, 12)                                   # 빈 레이아웃: 제목 개체틀 없음
    s.Shapes.add(FakeShape("본문", 36, 100, 800, 300, text="본문 첫 줄"))
    code, out = run_main(ppt, capsys, "write", "--slide", "4", "--title", "새 제목")
    assert code == 0 and out["status"] == "ok", out                # 이전: needs_user_decision (본문 첫 줄을 제목으로 봄)
    assert "새 제목" in ppt.slide_units(prs.Slides(4), 4)["text"] + ppt.slide_units(prs.Slides(4), 4)["title"]


def test_same_heading_as_title_is_not_duplicated(pp, capsys, tmp_path):
    md = tmp_path / "내용.md"
    md.write_text("## 개요\n- 가\n", encoding="utf-8")
    code, out = run_main(ppt, capsys, "write", "--slide", "new", "--title", "개요 ", "--file", str(md))
    assert code == 0, out
    u = ppt.slide_units(pp["prs"].Slides(4), 4)
    assert u["title"] == "개요" and "개요" not in u["text"] and "가" in u["text"]
    md.write_text("## 첫 절\n- 가\n", encoding="utf-8")
    code, out = run_main(ppt, capsys, "write", "--slide", "new", "--title", "발표", "--file", str(md))
    assert "첫 절" in ppt.slide_units(pp["prs"].Slides(5), 5)["text"]


# ---------------------------------------------------------------------------
# 카드 공간
# ---------------------------------------------------------------------------


def test_cards_go_to_new_slide_when_no_space(pp, capsys):
    text = "\\n".join(f"- 줄 {i}" for i in range(1, 16))
    code, out = run_ppt(capsys, "write", "--slide", "new", "--title", "핵심", "--text", text,
                        "--cards", "안정성:무중단; 확장성:모듈")
    assert code == 0, out
    prs = pp["prs"]
    H, W = float(prs.PageSetup.SlideHeight), float(prs.PageSetup.SlideWidth)
    assert out["slides_written"] == [4, 5] and any("새 슬라이드 5" in w for w in out["warnings"]), out
    for n in (4, 5):
        for sh in live(prs.Slides(n)):
            assert float(sh.Top) + float(sh.Height) <= H + 0.5 and float(sh.Left) + float(sh.Width) <= W + 0.5, sh.Name
    assert any("안정성" in sh.TextFrame.TextRange.Text for sh in live(prs.Slides(5)))
    assert out["agent_directive"].startswith("VERIFIED")


def test_verify_inside_flags_off_slide_shape(pp):
    class Sh:
        Name, Left, Top, Width, Height = "카드", 36.0, 518.0, 300.0, 80.0
    prs = pp["prs"]
    problems = ppt._verify_inside(prs, [(4, Sh())])
    assert problems and "슬라이드 밖" in problems[0]


# ---------------------------------------------------------------------------
# 코드 블록 줄은 ** __를 굵게로 읽지 않고 그대로, 인라인 코드 ```x```는 펜스가 아님
# ---------------------------------------------------------------------------


def test_ppt_code_block_verbatim(pp, capsys, tmp_path):
    p = tmp_path / "c.md"
    p.write_text(CODE_MD, encoding="utf-8")
    code, out = run_ppt(capsys, "write", "--file", str(p))
    prs = pp["prs"]
    s = prs.Slides(prs.Slides.Count)
    texts = [sh.TextFrame.TextRange.Text for sh in live(s) if hasattr(sh, "TextFrame")]
    assert code == 0 and any("__init__" in t and "x**2" in t for t in texts), texts


# ---------------------------------------------------------------------------
# ppt --text 펜스
# ---------------------------------------------------------------------------


def test_ppt_text_fence_verbatim(pp, capsys):
    code, out = run_ppt(capsys, "write", "--slide", "new", "--text", "설명\\n```\\nx**2 = __a__\\n```\\n- 끝")
    texts = "\n".join(_texts(pp))
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    assert "x**2 = __a__" in texts and "```" not in texts, texts


# ---------------------------------------------------------------------------
# ppt 3단계 제목 굵게 표시
# ---------------------------------------------------------------------------


def test_ppt_level3_heading_bold_markers_removed(pp, capsys, tmp_path):
    code, out = _ppt_md(capsys, tmp_path, "## 슬라이드\n\n### **핵심** 요약\n\n본문\n")
    texts = "\n".join(_texts(pp))
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    assert "핵심 요약" in texts and "**" not in texts, texts


# ---------------------------------------------------------------------------
# 단어 안 __
# ---------------------------------------------------------------------------


def test_ppt_text_intraword_underscore_literal(pp, capsys):
    code, out = run_ppt(capsys, "write", "--slide", "new", "--text", "report__2024__final.xlsx 참조")
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    assert "report__2024__final.xlsx 참조" in _texts(pp)


# ---------------------------------------------------------------------------
# 제목 규칙
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("md,title,body0", [
    ("## **핵심** 요약\n\n본문\n", "핵심 요약", "본문"),
    ("# __init__ 메서드\n\n본문\n", "init 메서드", "본문"),
    ("## 클래스\n\n### __init__ 메서드\n\n본문\n", "클래스", "init 메서드"),
    ("## 클래스\n\n#### **핵심** 요약\n\n본문\n", "클래스", "핵심 요약"),
    ("## report__2024__final 파일\n\n본문\n", "report__2024__final 파일", "본문"),
])
def test_ppt_headings_same_rule_every_level(pp, capsys, tmp_path, md, title, body0):
    code, out = _ppt_md(capsys, tmp_path, md)
    texts = _texts(pp)
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    assert texts[0] == title and texts[1].split("\r")[0] == body0, texts


def test_ppt_md_title_moved_to_body_is_plain(pp, capsys, tmp_path):
    code, out = _ppt_md(capsys, tmp_path, "## **핵심** 요약\n\n본문\n", "--title", "새 제목")
    texts = _texts(pp)
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    assert texts == ["새 제목", "핵심 요약\r본문"], texts


def test_ppt_title_option_plain(pp, capsys):
    code, out = run_ppt(capsys, "write", "--slide", "new", "--title", "**핵심** 요약", "--text", "본문")
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    assert _texts(pp)[0] == "핵심 요약", _texts(pp)


# ---------------------------------------------------------------------------
# 연도와 1이 아닌 번호
# ---------------------------------------------------------------------------


def test_ppt_date_line_kept(pp, capsys):
    code, out = run_ppt(capsys, "write", "--slide", "new", "--text", "2026. 11. 3. 회의 결과")
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    assert "2026. 11. 3. 회의 결과" in _texts(pp)


def test_ppt_text_list_from_3_keeps_numbers(pp, capsys):
    code, out = run_ppt(capsys, "write", "--slide", "new", "--text", "3. 셋\\n4. 넷\\n- 끝\\n1. 하나\\n2. 둘")
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    assert ppt._text_lines("3. 셋\\n4. 넷\\n- 끝\\n1. 하나\\n2. 둘") == [
        ("3. 셋", ""), ("4. 넷", ""), ("끝", "bullet"), ("하나", "numbered"), ("둘", "numbered")]
    assert "3. 셋\r4. 넷\r끝\r하나\r둘" in _texts(pp)


def test_ppt_md_list_from_3_keeps_numbers(pp, capsys, tmp_path):
    code, out = _ppt_md(capsys, tmp_path, "## 안건\n\n3. 셋\n4. 넷\n")
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    assert _texts(pp) == ["안건", "3. 셋\r4. 넷"], _texts(pp)


# ---------------------------------------------------------------------------
# '---'는 구분선, 맨 앞 YAML 머리말은 넣지 않음
# ---------------------------------------------------------------------------


def test_ppt_md_sections_skip_rule_and_front_matter():
    secs = ppt._md_sections("---\ntitle: 비밀\n---\n# 첫\n본문\n\n---\n\n둘째 줄\n## 다음\n내용")
    assert [s["title"] for s in secs] == ["첫", "다음"]
    assert secs[0]["lines"] == [("본문", ""), ("둘째 줄", "")]


# ---------------------------------------------------------------------------
# 바꾼 것 없이 성공한 실행은 되돌리기 지점을 남기지 않음 (다음 undo가 바뀐 것 없는 지점으로 가지 않게)
# ---------------------------------------------------------------------------


def test_ppt_clear_of_empty_slide_drops_point(monkeypatch, capsys):
    prs = FakePresentation(2)
    app = types.SimpleNamespace(ActiveWindow=types.SimpleNamespace(View=types.SimpleNamespace(GotoSlide=lambda n: None),
                                                                   Presentation=prs))
    stub_com(monkeypatch)
    monkeypatch.setattr(off, "save_snapshot", lambda *a, **k: None)      # 지점은 장벽(barrier)으로 남음
    monkeypatch.setattr(off, "get_active", lambda kind: {"name": prs.Name, "path": prs.FullName, "app": app, "doc": prs})
    code, out = run_main(ppt, capsys, "slide", "--action", "clear", "--slide", "1")
    assert code == 0, out
    first = undo_points("ppt")
    assert len(first) == 1
    code, out = run_main(ppt, capsys, "slide", "--action", "clear", "--slide", "1")   # 이미 빈 슬라이드
    assert code == 0 and "0개" in out["preview"], out
    assert undo_points("ppt") == first


# ---------------------------------------------------------------------------
# 바꾼 것 없이 성공한 명령은 바꾼 것으로 표시하지 않음 (되돌리기 지점을 남기지 않고 0곳 바뀜으로 보고)
# ---------------------------------------------------------------------------


def test_slide_move_old_to_option_recommends_dest(pp, capsys):
    """1.4.79: 옮길 위치는 --dest (excel range --dest와 같은 이름, --to는 outlook 받는 사람 뜻만). 옛 --to는 받지 않고 고친 전체 명령 [추천]."""
    prs = pp["prs"]
    first = prs.Slides(1)
    code, out = run_ppt(capsys, "slide", "--action", "move", "--slide", "1", "--to", "3")
    assert code == 2 and out["agent_directive"].startswith("STOP_RETRY:") and prs.Slides(1) is first, out
    argv = recommended_argv(out["message"], "ppt")
    assert argv == ["slide", "--action", "move", "--slide", "1", "--dest", "3"], argv
    code, out = run_ppt(capsys, *argv)                          # [추천]대로 다시 실행하면 옮겨짐
    assert code == 0 and prs.Slides(3) is first, out
    code, out = run_ppt(capsys, "slide", "--action", "move", "--slide", "1")
    assert code == 1 and "--dest" in out["message"], out        # 위치가 없다는 안내도 새 이름


def test_ppt_move_to_same_position_keeps_previous_point(pp_prs, capsys):
    code, out = run_main(ppt, capsys, "notes", "--slide", "1", "--text", "실제 변경")
    assert code == 0, out
    before = undo_points("ppt")
    assert len(before) == 1
    code, out = run_main(ppt, capsys, "slide", "--action", "move", "--slide", "2", "--dest", "2")
    assert code == 0 and "바뀐 것 없음" in out["preview"], out
    assert undo_points("ppt") == before                         # 다음 undo는 notes를 되돌림
    code, out = run_main(ppt, capsys, "slide", "--action", "move", "--slide", "2", "--dest", "3")
    assert code == 0 and len(undo_points("ppt")) == 2, out      # 실제로 옮기면 지점이 남음


def test_ppt_notes_same_text_is_noop(pp_prs, capsys):
    code, out = run_main(ppt, capsys, "notes", "--slide", "1", "--text", "같은 노트")
    assert code == 0
    before = undo_points("ppt")
    code, out = run_main(ppt, capsys, "notes", "--slide", "1", "--text", "같은 노트", "--overwrite")
    assert code == 0 and out.get("unchanged") is True and "바뀐 것 없음" in out["preview"], out
    assert undo_points("ppt") == before


def test_ppt_replace_same_text_reports_zero(pp_prs, capsys):
    pp_prs.Slides(1).Shapes.AddTextbox(1, 0, 0, 100, 40).TextFrame.TextRange.Text = "A 계획"
    code, out = run_main(ppt, capsys, "replace", "--map", '{"A": "A"}')
    assert code == 0 and out["replaced"] == 0 and out["unchanged_same_text"] == {"A": 1}, out
    assert "0곳" in out["preview"] and undo_points("ppt") == []


def test_ppt_replace_find_only_is_refused(pp_prs, capsys):
    code, out = run_main(ppt, capsys, "replace", "--find", "A")
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:"), out
    assert undo_points("ppt") == []


# ---------------------------------------------------------------------------
# replace: 대소문자만 다른 바꾸기
# ---------------------------------------------------------------------------


def test_ppt_case_normalizing_replace_really_replaces(pp_prs, capsys):
    tb = pp_prs.Slides(1).Shapes.AddTextbox(1, 0, 0, 100, 40)
    tb.TextFrame.TextRange.Text = "IPHONE and iphone"
    code, out = run_main(ppt, capsys, "replace", "--find", "iPhone", "--replace", "iPhone")
    assert code == 0 and out["replaced"] == 2, out
    assert tb.TextFrame.TextRange.Text == "iPhone and iPhone"


def test_ppt_same_text_same_case_is_still_noop(pp_prs, capsys):
    tb = pp_prs.Slides(1).Shapes.AddTextbox(1, 0, 0, 100, 40)
    tb.TextFrame.TextRange.Text = "iPhone and iPhone"
    code, out = run_main(ppt, capsys, "replace", "--find", "iPhone", "--replace", "iPhone")
    assert code == 0 and out["replaced"] == 0 and out["unchanged_same_text"] == {"iPhone": 2}, out


# ---------------------------------------------------------------------------
# 빈 찾을 글은 모든 글자 사이에 바꿀 글을 끼워 넣음 -> 바꾸기 전에 STOP_RETRY
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("argv", EMPTY_FIND)
def test_ppt_empty_find_is_refused(pp_prs, capsys, argv):
    tb = pp_prs.Slides(1).Shapes.AddTextbox(1, 0, 0, 100, 40)
    tb.TextFrame.TextRange.Text = "abc"
    code, out = run_main(ppt, capsys, "replace", *argv)
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:"), out
    assert "ppt.py write" in out["message"], out
    assert tb.TextFrame.TextRange.Text == "abc"
    assert undo_points("ppt") == []


# ---------------------------------------------------------------------------
# 대소문자만 다른 바꾸기에서 이미 바꿀 글인 곳은 바꾼 곳으로 세지 않음
# ---------------------------------------------------------------------------


def test_ppt_case_only_already_replaced_is_noop(pp_prs, capsys):
    tb = pp_prs.Slides(1).Shapes.AddTextbox(1, 0, 0, 100, 40)
    tb.TextFrame.TextRange.Text = "iPhone and iPhone"
    code, out = run_main(ppt, capsys, "replace", "--find", "iphone", "--replace", "iPhone")
    assert code == 0 and out["replaced"] == 0 and out["unchanged_same_text"] == {"iphone": 2}, out
    assert out["slides"] == [] and "0곳" in out["preview"], out
    assert tb.TextFrame.TextRange.Text == "iPhone and iPhone"
    assert undo_points("ppt") == []


def test_ppt_counts_only_changed_occurrences(pp_prs, capsys):
    tb = pp_prs.Slides(1).Shapes.AddTextbox(1, 0, 0, 100, 40)
    tb.TextFrame.TextRange.Text = "iPhone and IPHONE"
    code, out = run_main(ppt, capsys, "replace", "--find", "iPhone", "--replace", "iPhone")
    assert code == 0 and out["replaced"] == 1 and out["unchanged_same_text"] == {"iPhone": 1}, out
    assert out["agent_directive"].startswith("VERIFIED:"), out
    assert tb.TextFrame.TextRange.Text == "iPhone and iPhone"


def test_ppt_case_only_mixed(pp_prs, capsys):
    tb = pp_prs.Slides(1).Shapes.AddTextbox(1, 0, 0, 100, 40)
    tb.TextFrame.TextRange.Text = "iphone, iPhone, IPHONE"
    code, out = run_main(ppt, capsys, "replace", "--find", "iphone", "--replace", "iPhone")
    assert code == 0 and out["replaced"] == 2 and out["unchanged_same_text"] == {"iphone": 1}, out
    assert tb.TextFrame.TextRange.Text == "iPhone, iPhone, iPhone"


def test_ppt_case_only_readback_checks_exact_case(pp_prs, capsys, monkeypatch):
    tb = pp_prs.Slides(1).Shapes.AddTextbox(1, 0, 0, 100, 40)
    tb.TextFrame.TextRange.Text = "IPHONE"
    real = ppt.replace_in_range

    def lower_only(tr, find, rep, match_case, fmt):   # 바꿨다고 보고하지만 대소문자가 그대로 남은 상황
        tr.Text = str(tr.Text).lower()
        return 1, 0

    monkeypatch.setattr(ppt, "replace_in_range", lower_only)
    code, out = run_main(ppt, capsys, "replace", "--find", "iphone", "--replace", "iPhone")
    monkeypatch.setattr(ppt, "replace_in_range", real)
    assert code == 0 and out["agent_directive"].startswith("UNVERIFIED:"), out


# ---------------------------------------------------------------------------
# replace 위치는 UTF-16 단위 (이모지 뒤 글)
# ---------------------------------------------------------------------------


def test_ppt_replace_after_emoji_replaces_all():
    for text, find, rep, want in (("\U0001F4CA\U0001F4CA\U0001F4CA 2025 2025", "2025", "2026", "\U0001F4CA\U0001F4CA\U0001F4CA 2026 2026"),
                                  ("\U0001F4CC\U0001F4CC ab ab", "ab", "X", "\U0001F4CC\U0001F4CC X X"),
                                  ("\U0001F4CC 핵심 진행중, 진행중", "진행중", "완료", "\U0001F4CC 핵심 완료, 완료")):
        tr = U16TextRange(text)
        assert ppt.replace_in_range(tr, find, rep, False, {}) == (2, 0), text
        assert tr.Text == want


# ---------------------------------------------------------------------------
# notes: 발표자 노트
# ---------------------------------------------------------------------------


def test_notes_read_only_slides_with_notes_and_zero(pp, capsys):
    code, out = run_ppt(capsys, "notes")
    assert code == 0 and out["count"] == 0 and "0건" in out["agent_directive"]
    pp["prs"].Slides(2).NotesPage.body.TextFrame.TextRange.Text = "첫 줄\r둘째 줄"
    code, out = run_ppt(capsys, "notes")
    assert code == 0 and out["count"] == 1
    assert out["notes"][0]["no"] == 2 and out["notes"][0]["notes"] == "첫 줄\n둘째 줄"


def test_notes_write_replace_and_append_verified(pp, capsys):
    code, out = run_ppt(capsys, "notes", "--slide", "3", "--text", "발표 요점\\n시간 2분")
    assert code == 0, out
    body = pp["prs"].Slides(3).NotesPage.body.TextFrame.TextRange
    assert body.Text == "발표 요점\r시간 2분"
    assert out["verified"] is True and out["agent_directive"].startswith("VERIFIED:")
    code, out = run_ppt(capsys, "notes", "--slide", "3", "--text", "질문 대비", "--append")
    assert code == 0 and body.Text == "발표 요점\r시간 2분\r질문 대비"
    assert out["before"] == "발표 요점\n시간 2분"
    assert pp["app"].ActiveWindow.View.goto[-1] == 3          # 고친 슬라이드로 화면 이동


def test_notes_write_from_file_and_errors(pp, capsys, tmp_path):
    f = tmp_path / "노트.md"
    f.write_text("파일에서 읽은 노트\n두 번째 줄\n", encoding="utf-8")
    code, out = run_ppt(capsys, "notes", "--slide", "1", "--file", str(f))
    assert code == 0 and pp["prs"].Slides(1).NotesPage.body.TextFrame.TextRange.Text == "파일에서 읽은 노트\r두 번째 줄"
    code, out = run_ppt(capsys, "notes", "--text", "슬라이드 없음")
    assert code == 1 and "--slide" in out["message"]
    code, out = run_ppt(capsys, "notes", "--slide", "9", "--text", "x")
    assert code == 1 and "1~3" in out["message"]


def test_notes_write_unverified_when_reread_differs(pp, capsys, monkeypatch):
    monkeypatch.setattr(ppt, "_notes_text", lambda slide: "")
    code, out = run_ppt(capsys, "notes", "--slide", "1", "--text", "새 노트")
    assert code == 0 and out["verified"] is False and out["agent_directive"].startswith("UNVERIFIED:")


def test_notes_existing_needs_append_or_overwrite(pp, capsys):
    body = pp["prs"].Slides(1).NotesPage.body.TextFrame.TextRange
    body.Text = "기존 노트"
    code, out = run_ppt(capsys, "notes", "--slide", "1", "--text", "새 노트")
    assert code == 1 and out["error_type"] == "needs_user_decision" and body.Text == "기존 노트"
    assert "--append" in out["agent_directive"] and "--overwrite" in out["agent_directive"]
    code, out = run_ppt(capsys, "notes", "--slide", "1", "--text", "새 노트", "--overwrite")
    assert code == 0 and body.Text == "새 노트"


def test_notes_read_with_slide_and_missing_body(pp, capsys):
    pp["prs"].Slides(1).NotesPage.body.TextFrame.TextRange.Text = "1번 노트"
    pp["prs"].Slides(2).NotesPage.body.TextFrame.TextRange.Text = "2번 노트"
    code, out = run_ppt(capsys, "notes", "--slide", "2")                # 읽기에서 --slide도 그 슬라이드만
    assert code == 0 and [n["no"] for n in out["notes"]] == [2]
    code, out = run_ppt(capsys, "notes", "--slides", "x")
    assert code == 1 and "--slides 형식" in out["message"]
    pp["prs"].Slides(3).NotesPage.Shapes.Placeholders.items.pop()   # 노트 본문 개체틀 없음
    code, out = run_ppt(capsys, "notes", "--slide", "3", "--text", "x")
    assert code == 1 and "본문 개체틀" in out["message"]


def test_notes_from_crlf_file_has_no_blank_lines(pp, capsys, tmp_path):
    f = tmp_path / "노트_crlf.md"
    f.write_bytes("첫 줄\r\n둘째 줄\r\n".encode("utf-8"))        # Windows 메모장 형식
    code, out = run_ppt(capsys, "notes", "--slide", "2", "--file", str(f))
    assert code == 0 and pp["prs"].Slides(2).NotesPage.body.TextFrame.TextRange.Text == "첫 줄\r둘째 줄"


def test_write_reports_unverified_when_text_is_missing(pp, capsys, monkeypatch):
    import ppt
    monkeypatch.setattr(ppt, "slide_units", lambda slide, idx: {"no": idx, "title": "", "text": "", "notes": ""})
    code, out = run_ppt(capsys, "write", "--slide", "new", "--title", "제목")
    assert code == 0 and out["agent_directive"].startswith("UNVERIFIED:") and "제목" in out["unverified"][0]


# ---------------------------------------------------------------------------
# 출력 크기 상한 (긴 칸, 찾은 값, 슬라이드 글, 워드 단락, 노트)
# ---------------------------------------------------------------------------


def test_ppt_notes_cut_sets_truncated(monkeypatch, capsys):
    prs = FakePresentation(2)
    app = FakePptWindowApp(prs)
    fake, _ = make_win32com()
    monkeypatch.setattr(off, "win32com", fake)
    monkeypatch.setattr(off, "pythoncom", None)
    monkeypatch.setattr(off, "list_open", lambda kind: [])
    monkeypatch.setattr(off, "get_active", lambda kind: {"name": prs.Name, "path": prs.FullName, "app": app, "doc": prs})
    prs.Slides(1).NotesPage.body.TextFrame.TextRange.Text = "바" * 9000
    code, out = run_main(ppt, capsys, "notes")
    assert code == 0 and out["truncated"] is True and out["notes"][0]["text_truncated"] and "truncated_reason" in out


# ---------------------------------------------------------------------------
# ppt 다시 읽기
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("md", [
    "## 코드\n\n```python\ndef __init__(self):\n    pass\n```\n",
    "## 수식\n\n```\nx**2 + y**2\n```\n",
    "## __init__ 설명\n\n본문\n",
])
def test_ppt_code_and_title_with_markers_verified(pp, capsys, tmp_path, md):
    code, out = _ppt_md(capsys, tmp_path, md)
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out


def test_ppt_title_option_with_dunder_verified(pp, capsys):
    code, out = run_ppt(capsys, "write", "--slide", "new", "--title", "__main__ 실행", "--text", "본문")
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out


# ---------------------------------------------------------------------------
# 0건
# ---------------------------------------------------------------------------


def test_ppt_slide_list_zero(pp, capsys):
    pp["prs"] = FakePresentation(0)
    code, out = run_ppt(capsys, "slide", "--action", "list")
    assert code == 0 and out["count"] == 0 and out["slides"] == [] and out["agent_directive"].startswith("VERIFIED: 0건"), out


def test_ppt_slide_list_count(pp, capsys):
    code, out = run_ppt(capsys, "slide", "--action", "list")
    assert code == 0 and out["count"] == len(out["slides"]) > 0, out


def test_cards_with_old_bar_separator_get_full_command(pp, capsys):
    """1.4.79: 카드 사이는 ; (저장소 구분자 규칙). '|'만 쓰면 쓰지 않고 ;로 고친 전체 명령을 [추천]."""
    before = pp["prs"].Slides.Count
    code, out = run_ppt(capsys, "write", "--slide", "new", "--title", "핵심", "--cards", "안정성:무중단|확장성:모듈")
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:") and pp["prs"].Slides.Count == before, out
    rec = out["message"].split("[추천] ", 1)[1]
    argv = split_command(rec)
    assert argv[0] == "python" and argv[argv.index("--cards") + 1] == "안정성:무중단; 확장성:모듈", argv
    code, out = run_ppt(capsys, *argv[2:])
    assert code == 0, out


def test_read_without_text_is_explicit_zero(pp, capsys):
    """1.4.81 빈 결과 명시: 글이 없는 슬라이드만 있으면 count 0과 '찾은 항목 없음' (빈 배열만 주지 않음)."""
    code, out = run_ppt(capsys, "read")
    assert code == 0 and out["count"] == 0 and out["agent_directive"].startswith("VERIFIED: 0건."), out
    assert "찾은 항목 없음" in out["agent_directive"] and out["preview"].startswith("0건:"), out


# ---------------------------------------------------------------------------
# --dry-run: replace, slide delete·clear는 바꾸지 않고 미리 보기 (읽기 세션, 되돌리기 지점 없음)
# confirm_commands[0]은 --dry-run만 뺀 같은 명령
# ---------------------------------------------------------------------------


def _confirm_argv(out):
    cmd = out["confirm_commands"][0]
    assert cmd.get("destructive") is True, cmd
    argv = split_command(cmd["command"])
    assert argv[0] == "python" and argv[1].endswith("ppt.py"), argv
    assert not any("next_commands" in k for k in out), out
    return argv[2:]


def _assert_preview(out, count):
    assert out["status"] == "ok" and out["dry_run"] is True and out["count"] == count, out
    assert out["agent_directive"].startswith("VERIFIED: 미리보기만 했고"), out
    assert out["preview"].startswith("미리보기:") and out["preview"].endswith("(바꾸지 않음)"), out


def test_ppt_replace_dry_run_counts_without_changing(pp_prs, capsys):
    tb = pp_prs.Slides(2).Shapes.AddTextbox(1, 0, 0, 100, 40)
    tb.TextFrame.TextRange.Text = "초안 초안 진행중"
    argv = ["replace", "--target", "active", "--find", "초안", "--replace", "확정", "--bold"]
    code, out = run_main(ppt, capsys, *argv, "--dry-run")
    assert code == 0, out
    _assert_preview(out, 2)
    assert out["slides"] == [2] and out["planned"] == ["슬라이드 2: '초안' -> '확정' 2곳"] and out["truncated"] is False, out
    assert tb.TextFrame.TextRange.Text == "초안 초안 진행중" and undo_points("ppt") == []
    assert _confirm_argv(out) == argv
    code, out = run_main(ppt, capsys, *argv)               # confirm 명령대로 실행하면 바뀜
    assert code == 0 and out["replaced"] == 2 and tb.TextFrame.TextRange.Text == "확정 확정 진행중", out
    assert len(undo_points("ppt")) == 1


def test_ppt_replace_dry_run_zero(pp_prs, capsys):
    tb = pp_prs.Slides(1).Shapes.AddTextbox(1, 0, 0, 100, 40)
    tb.TextFrame.TextRange.Text = "iPhone"
    code, out = run_main(ppt, capsys, "replace", "--find", "없는글", "--replace", "x", "--dry-run")
    assert code == 0, out
    _assert_preview(out, 0)
    assert "바뀔 것 없음" in out["preview"] and "confirm_commands" not in out and out["planned"] == [], out
    code, out = run_main(ppt, capsys, "replace", "--find", "iphone", "--replace", "iPhone", "--dry-run")
    assert code == 0 and out["count"] == 0 and out["unchanged_same_text"] == {"iphone": 1}, out
    assert "바뀔 것 없음" in out["preview"] and undo_points("ppt") == []


def test_ppt_replace_dry_run_map_matches_real_count(pp_prs, capsys):
    tb = pp_prs.Slides(1).Shapes.AddTextbox(1, 0, 0, 100, 40)
    tb.TextFrame.TextRange.Text = "iphone, iPhone, IPHONE, aa"
    argv = ["replace", "--target", "active", "--map", '{"iphone": "iPhone", "a": "b"}']
    code, out = run_main(ppt, capsys, *argv, "--dry-run")
    assert code == 0 and out["count"] == 4 and out["by_text"] == {"iphone": 2, "a": 2}, out
    assert _confirm_argv(out) == argv and tb.TextFrame.TextRange.Text == "iphone, iPhone, IPHONE, aa"
    code, real = run_main(ppt, capsys, *argv)
    assert code == 0 and real["replaced"] == out["count"] and real["by_text"] == out["by_text"], real


def test_ppt_replace_dry_run_planned_capped(pp_prs, capsys):
    pp_prs.Slides(1).Shapes.AddTextbox(1, 0, 0, 100, 40).TextFrame.TextRange.Text = " ".join(f"w{i:02d}" for i in range(60))
    mapping = "{" + ", ".join(f'"w{i:02d}": "v"' for i in range(60)) + "}"
    code, out = run_main(ppt, capsys, "replace", "--map", mapping, "--dry-run")
    assert code == 0 and out["count"] == 60 and len(out["planned"]) == 50 and out["truncated"] is True, out


def test_ppt_slide_delete_dry_run(pp_prs, capsys):
    pp_prs.Slides(2).Shapes.Title.TextFrame.TextRange.Text = "둘째 장"
    second = pp_prs.Slides(2)
    argv = ["slide", "--target", "active", "--action", "delete", "--slide", "2"]
    code, out = run_main(ppt, capsys, *argv, "--dry-run")
    assert code == 0, out
    _assert_preview(out, 1)
    assert "둘째 장" in out["preview"] and out["planned"][0].startswith("삭제: 슬라이드 2('둘째 장')"), out
    assert pp_prs.Slides.Count == 3 and pp_prs.Slides(2) is second and undo_points("ppt") == []
    assert _confirm_argv(out) == argv
    code, out = run_main(ppt, capsys, *argv)
    assert code == 0 and pp_prs.Slides.Count == 2 and len(undo_points("ppt")) == 1, out
    code, out = run_main(ppt, capsys, "slide", "--action", "delete", "--slide", "9", "--dry-run")
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:"), out


def test_ppt_slide_clear_dry_run_one_and_all(pp_prs, capsys):
    pp_prs.Slides(1).Shapes.add(FakeShape("a", text="가"))
    pp_prs.Slides(3).Shapes.add(FakeShape("b", text="나"))
    before = [int(pp_prs.Slides(n).Shapes.Count) for n in (1, 2, 3)]
    argv = ["slide", "--target", "active", "--action", "clear", "--all-slides"]
    code, out = run_main(ppt, capsys, *argv, "--dry-run")
    assert code == 0, out
    _assert_preview(out, sum(before))
    assert len(out["planned"]) == sum(1 for b in before if b), out
    assert [int(pp_prs.Slides(n).Shapes.Count) for n in (1, 2, 3)] == before and undo_points("ppt") == []
    assert _confirm_argv(out) == argv
    code, out = run_main(ppt, capsys, "slide", "--action", "clear", "--dry-run")
    assert code == 1 and "--all-slides" in out["message"], out
    code, out = run_main(ppt, capsys, *argv)
    assert code == 0 and all(int(pp_prs.Slides(n).Shapes.Count) == 0 for n in (1, 2, 3)), out
    code, out = run_main(ppt, capsys, "slide", "--action", "clear", "--slide", "2", "--dry-run")
    assert code == 0, out
    _assert_preview(out, 0)
    assert "바뀔 것 없음" in out["preview"] and "confirm_commands" not in out, out


def test_ppt_slide_list_ignores_dry_run_with_warning(pp_prs, capsys):
    code, out = run_main(ppt, capsys, "slide", "--action", "list", "--dry-run")
    assert code == 0 and out["count"] == 3 and "--dry-run" in out["warnings"][0], out
