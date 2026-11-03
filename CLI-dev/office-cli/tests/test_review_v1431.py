"""1.4.31 (22차 점검) 회귀 테스트.

R1 (1.4.30 회귀): relay --to ppt가 쓰기 세션(되돌리기 지점)을 임시 워드 단계보다 먼저 열어, 워드 단계가 실패하면 저장 안 된
    활성 PPT에 바꾼 것 없는 지점이 남고 다음 undo가 그 지점을 먼저 골랐음.
H1: 닫힌 xlsx를 data_only로 읽어 계산된 값이 없는 수식 칸(openpyxl이 쓰고 엑셀에서 연 적 없는 파일)이 빈 칸인데 VERIFIED.
H2: docx 글상자(w:txbxContent) 글을 읽지 않아 search가 VERIFIED 0건 (word read도 같은 경로).
H3: transfer --from word --what table(닫힌 docx)이 python-docx Document.tables(본문 바로 아래 표만)로 표를 세어 콘텐츠 컨트롤(w:sdt)
    안의 표가 있으면 read 결과와 다른 표를 옮겼음.
L1: word 굵게 위치를 파이썬 글자로 세어 앞에 이모지가 있으면 굵게가 밀렸음 (Word 위치는 UTF-16 단위).
L2: excel read --search에 count가 없고 0건 안내가 R2 모양이 아니었음.
L3: 마크다운 '---'가 페이지 나누기, YAML 머리말이 페이지 나누기와 글로 들어갔음.
"""
import json
import os
import struct
import time
import types

import pytest

import _office as off
import _runtime as rt
import excel
import office
import ppt
import reader
import word
import fake_word as fw
import test_review_v1430 as v1430
import test_word

wd = test_word.wd                      # fixture (가짜 활성 워드 문서)
EMOJI = chr(0x1F680)                   # 저장소 파일에 이모지를 직접 쓰지 않음


def _main(mod, capsys, *argv):
    with pytest.raises(SystemExit) as ei:
        mod.main(list(argv))
    out = [ln for ln in capsys.readouterr().out.splitlines() if ln.strip()]
    assert len(out) == 1, out
    return ei.value.code, json.loads(out[0])


# R1: relay는 워드 단계 뒤에 PPT 쓰기 세션을 엶 (실제 open_target, 가짜 COM) ---------------------------------------------

class _Prs:
    def __init__(self, full, saved=-1, n=2):
        self.FullName = full
        self.Name = os.path.basename(full)
        self.Saved = saved
        self.Slides = v1430._Slides(n)
        self.closed = 0
        self.PageSetup = types.SimpleNamespace(SlideWidth=960, SlideHeight=540)
        self.ReadOnly = 0

    def SaveCopyAs(self, p):
        with open(p, "wb") as f:
            f.write(b"snap")

    def Save(self):
        pass

    def Close(self):
        self.closed += 1


class _Pres:
    def __init__(self):
        self.items = []

    @property
    def Count(self):
        return len(self.items)

    def Open(self, path, *a):
        p = _Prs(path)
        self.items.append(p)
        return p


@pytest.fixture
def relay_real(monkeypatch):
    """open_target은 진짜, COM은 가짜. log: 클립보드 복사, 임시 워드 생성."""
    log = {"copied": 0, "dispatch": 0}
    app = types.SimpleNamespace(Presentations=_Pres(), DisplayAlerts=2, Quit=lambda: None)
    monkeypatch.setattr(off, "pythoncom", None)
    monkeypatch.setattr(office.time, "sleep", lambda s: None)

    def tmp_word(paste_ok):
        class D:
            InlineShapes = types.SimpleNamespace(Count=1 if paste_ok else 0, Item=lambda i: types.SimpleNamespace(
                Range=types.SimpleNamespace(CopyAsPicture=lambda: None)))

            def Range(self, a, b):
                return types.SimpleNamespace(PasteSpecial=lambda **k: None)

            def Close(self, **k):
                pass
        return types.SimpleNamespace(Documents=types.SimpleNamespace(Add=lambda **k: D()), Quit=lambda **k: None,
                                     Visible=True)

    def install(active=None, paste_ok=True):
        def dispatch_ex(name):
            log["dispatch"] += 1
            return tmp_word(paste_ok)
        monkeypatch.setattr(off, "win32com", types.SimpleNamespace(client=types.SimpleNamespace(
            DispatchEx=dispatch_ex, Dispatch=lambda n: app, GetActiveObject=lambda n: app)))

        def get_active(kind):
            if active is None:
                raise rt.UserDecisionError("열린 문서 없음", "x")
            return {"name": active.Name, "path": active.FullName, "app": app, "doc": active}
        monkeypatch.setattr(off, "get_active", get_active)
        monkeypatch.setattr(off, "list_open", lambda kind: [])
    log["install"] = install
    log["app"] = app

    def copy(*a):
        log["copied"] += 1
    log["chart"] = types.SimpleNamespace(CopyPicture=copy)
    return log


def test_r1_failed_temp_word_keeps_no_point_for_unsaved_active_deck(relay_real, tmp_path):
    deck = _Prs(str(tmp_path / "user.pptx"), saved=0)
    (tmp_path / "user.pptx").write_bytes(b"x")
    relay_real["install"](active=deck, paste_ok=False)
    with pytest.raises(RuntimeError):
        office.relay_chart(relay_real["chart"], None, "active", None, "1", "t")
    assert off.list_snapshots("ppt") == []                # 바꾼 것 없는 지점이 남지 않음
    assert deck.Slides(1).Shapes.Count == 0


def test_r1_undo_still_picks_earlier_real_point(relay_real, tmp_path):
    deck = _Prs(str(tmp_path / "user.pptx"), saved=0)
    (tmp_path / "user.pptx").write_bytes(b"x")
    relay_real["install"](active=deck, paste_ok=False)
    with off.open_target("ppt", "active", readonly=False, tool="ppt", command="write") as t:
        t.changed = True                                  # 앞서 한 진짜 쓰기
    before = [s["meta_file"] for s in off.list_snapshots(None)]
    assert len(before) == 1
    time.sleep(1.05)                                      # 지점 이름은 초 단위
    with pytest.raises(RuntimeError):
        office.relay_chart(relay_real["chart"], None, "active", None, "1", "t")
    assert [s["meta_file"] for s in off.list_snapshots(None)] == before


@pytest.mark.parametrize("where", ["active", "file", "new"])
def test_r1_bad_slide_stops_before_clipboard_and_point(relay_real, tmp_path, where):
    deck = _Prs(str(tmp_path / "user.pptx"), saved=0)
    (tmp_path / "user.pptx").write_bytes(b"x")
    relay_real["install"](active=deck)
    target = {"active": "active", "file": str(tmp_path / "user.pptx"), "new": str(tmp_path / "새.pptx")}[where]
    with pytest.raises(rt.InputError) as ei:
        office.relay_chart(relay_real["chart"], None, target, None, "9", "t")
    assert "슬라이드 9번이 없습니다" in str(ei.value) and ei.value.no_fallback
    assert relay_real["copied"] == 0 and relay_real["dispatch"] == 0    # 클립보드·임시 워드 그대로
    assert off.list_snapshots(None) == []
    assert not (tmp_path / "새.pptx").exists()               # 새 파일도 만들지 않음


def test_r1_success_active_and_file(relay_real, tmp_path):
    deck = _Prs(str(tmp_path / "user.pptx"))
    (tmp_path / "user.pptx").write_bytes(b"x")
    relay_real["install"](active=deck)
    done = office.relay_chart(relay_real["chart"], None, "active", None, "2", "t")
    assert done["ppt"]["source"] == "active" and deck.Slides(2).Shapes.Count == 1 and deck.closed == 0
    p = tmp_path / "file.pptx"
    p.write_bytes(b"x")
    done = office.relay_chart(relay_real["chart"], None, str(p), None, "1", "t")
    opened = relay_real["app"].Presentations.items[-1]
    assert done["ppt"]["source"] == "file" and opened.closed == 1 and opened.Slides(1).Shapes.Count == 1


# H1: 계산된 값이 없는 수식 칸 ---------------------------------------------------------------------------------------

def _formula_xlsx(path, n=1, cached=False):
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "합계"
    ws.append(["항목", "값"])
    ws.append(["a", 10])
    ws.append(["b", 20])
    for i in range(n):
        ws.append([f"합계{i + 1}", "=SUM(B2:B3)"])
    wb.save(path)
    if cached:
        import test_reader
        test_reader.set_cached_value(str(path), "<v></v>", "<v>30</v>")
    return str(path)


def test_h1_detect_helper(tmp_path):
    p = _formula_xlsx(tmp_path / "f.xlsx", n=25)
    info = rt.xlsx_uncomputed_formulas(p)
    assert info["count"] == 25 and info["cells"][0] == "합계!B4" and len(info["cells"]) == rt.UNCOMPUTED_SHOW
    assert rt.xlsx_uncomputed_formulas(p, area=(1, 1, 3, 2))["count"] == 0        # 범위 밖
    assert rt.xlsx_uncomputed_formulas(p, sheet="없음")["count"] == 0
    fields = rt.uncomputed_fields(info)
    assert fields["uncomputed_count"] == 25 and fields["uncomputed_truncated"] is True
    assert rt.xlsx_uncomputed_formulas(_formula_xlsx(tmp_path / "c.xlsx", cached=True))["count"] == 0


def test_h1_excel_read_closed_is_unverified(tmp_path, capsys):
    p = _formula_xlsx(tmp_path / "f.xlsx")
    code, out = _main(excel, capsys, "read", "--target", p)
    assert code == 0 and out["agent_directive"].startswith("UNVERIFIED:"), out
    assert out["uncomputed_formulas"] == ["합계!B4"] and out["uncomputed_count"] == 1
    assert "계산된 값" in out["agent_directive"] and "엑셀에서 파일을 열어 저장" in out["agent_directive"]
    assert out["values"][3] == ["합계1", ""]
    # 수식 칸이 없는 범위, 계산된 값이 있는 파일은 그대로 VERIFIED
    code, out = _main(excel, capsys, "read", "--target", p, "--range", "A1:B3")
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:") and "uncomputed_formulas" not in out
    code, out = _main(excel, capsys, "read", "--target", _formula_xlsx(tmp_path / "c.xlsx", cached=True))
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:") and out["values"][3] == ["합계1", 30]


def test_h1_excel_read_search_save_and_export(tmp_path, capsys):
    p = _formula_xlsx(tmp_path / "f.xlsx")
    code, out = _main(excel, capsys, "read", "--target", p, "--search", "30")
    assert code == 0 and out["count"] == 0 and out["agent_directive"].startswith("UNVERIFIED:"), out   # 0건이라 단정하지 않음
    code, out = _main(excel, capsys, "read", "--target", p, "--save", str(tmp_path / "o.csv"))
    assert code == 0 and out["agent_directive"].startswith("UNVERIFIED:") and out["uncomputed_count"] == 1
    code, out = _main(excel, capsys, "export", "--target", p, "--save", str(tmp_path / "e.csv"))
    assert code == 0 and out["agent_directive"].startswith("UNVERIFIED:") and out["uncomputed_count"] == 1


def test_h1_transfer_range_rows_report_uncomputed(tmp_path):
    p = _formula_xlsx(tmp_path / "f.xlsx")
    args = types.SimpleNamespace(from_target=p, sheet=None, range="A1:B4")
    rows, sheet, info = office._excel_range_rows(args)
    assert sheet == "합계" and info["cells"] == ["합계!B4"]
    payload = rt.with_uncomputed(rt.ok("office", "transfer", "x", "VERIFIED: 확인", verified=True, rows=4), info)
    assert payload["agent_directive"].startswith("UNVERIFIED:") and payload["rows"] == 4 and payload["uncomputed_count"] == 1
    # 범위 전체가 계산된 값 없는 수식이면 '데이터 없음' 오류에 이유를 붙임
    args = types.SimpleNamespace(from_target=p, sheet=None, range="B4:B4")
    with pytest.raises(rt.InputError) as ei:
        office._excel_range_rows(args)
    assert "계산된 값" in str(ei.value) and "합계!B4" in str(ei.value)
    assert "합계!B4" not in ei.value.log_message and "수식 칸 1개" in ei.value.log_message    # 로그에는 칸 주소 없음


def test_h1_reader_read_and_search(tmp_path, capsys):
    p = _formula_xlsx(tmp_path / "f.xlsx")
    code, out = _main(reader, capsys, "read", "--file", p)
    assert code == 0 and out["agent_directive"].startswith("UNVERIFIED:") and out["uncomputed_formulas"] == ["합계!B4"], out
    code, out = _main(reader, capsys, "search", "--query", "없는글자ZQX", "--file", p)
    assert code == 0 and out["agent_directive"].startswith("UNVERIFIED:"), out
    assert any("계산된 값" in g for g in out["not_searched"])
    c = _formula_xlsx(tmp_path / "c.xlsx", cached=True)
    code, out = _main(reader, capsys, "search", "--query", "없는글자ZQX", "--file", c)
    assert code == 0 and out["agent_directive"].startswith("VERIFIED: 0건."), out


# H2: 글상자 글 -------------------------------------------------------------------------------------------------------

_NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
       'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
       'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
       'xmlns:v="urn:schemas-microsoft-com:vml"')


def _txbx(text):
    return f'<w:txbxContent><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:txbxContent>'


def _textbox_docx(path, body=True):
    import docx
    from docx.oxml import parse_xml
    d = docx.Document()
    d.add_paragraph("본문 첫 단락")
    tail = d.add_paragraph("본문 끝 단락")
    if body:
        # Word 2010 이후: 같은 글상자를 DrawingML(mc:Choice)과 VML(mc:Fallback) 두 벌로
        alt = parse_xml(f'<w:p {_NS}><w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wps:txbx>'
                        f'{_txbx("두 벌 저장 BOXTWO")}</wps:txbx></w:drawing></mc:Choice><mc:Fallback><w:pict><v:shape>'
                        f'<v:textbox>{_txbx("두 벌 저장 BOXTWO")}</v:textbox></v:shape></w:pict></mc:Fallback>'
                        f'</mc:AlternateContent></w:r></w:p>')
        vml = parse_xml(f'<w:p {_NS}><w:r><w:t>앞 글</w:t></w:r><w:r><w:pict><v:shape><v:textbox>'
                        f'{_txbx("옛 글상자 BOXOLD")}</v:textbox></v:shape></w:pict></w:r></w:p>')
        tail._p.addprevious(alt)
        tail._p.addprevious(vml)
    d.save(path)
    return str(path)


def test_h2_parse_docx_reads_textboxes_once(tmp_path):
    units, meta = reader.parse_docx(_textbox_docx(tmp_path / "t.docx"))
    texts = [(u["loc"], u["text"]) for u in units]
    assert texts == [("단락 1", "본문 첫 단락"), ("글상자 1", "두 벌 저장 BOXTWO"), ("단락 2", "앞 글"),
                     ("글상자 2", "옛 글상자 BOXOLD"), ("단락 3", "본문 끝 단락")], texts
    assert meta["textboxes"] == 2 and meta["paragraphs"] == 3
    # 글상자가 없는 문서는 단위 번호·위치 그대로
    units, meta = reader.parse_docx(_textbox_docx(tmp_path / "n.docx", body=False))
    assert [u["loc"] for u in units] == ["단락 1", "단락 2"] and "textboxes" not in meta


def test_h2_textbox_in_table_cell_not_doubled(tmp_path):
    import docx
    from docx.oxml import parse_xml
    d = docx.Document()
    cell = d.add_table(rows=2, cols=1).cell(0, 0)
    cell.paragraphs[0].text = "머리"
    cell._tc.append(parse_xml(f'<w:p {_NS}><w:r><w:pict><v:shape><v:textbox>{_txbx("칸 글상자 BOXCELL")}'
                              f'</v:textbox></v:shape></w:pict></w:r></w:p>'))
    d.tables[0].cell(1, 0).text = "값"
    p = str(tmp_path / "c.docx")
    d.save(p)
    units, _ = reader.parse_docx(p)
    joined = "\n".join(u["text"] for u in units)
    assert joined.count("BOXCELL") == 1 and units[-1]["loc"] == "글상자 1", [(u["loc"], u["text"]) for u in units]


def test_h2_search_and_word_read_include_textbox(tmp_path, capsys):
    p = _textbox_docx(tmp_path / "t.docx")
    code, out = _main(reader, capsys, "search", "--query", "BOXOLD", "--file", p)
    assert code == 0 and out["total_matches"] == 1 and out["matches"][0]["loc"] == "글상자 2", out
    code, out = _main(reader, capsys, "search", "--query", "BOXTWO", "--file", p)
    assert out["total_matches"] == 1                       # 두 벌 저장도 한 번
    code, out = _main(reader, capsys, "read", "--file", p)
    assert any("BOXTWO" in u["text"] for u in out["units"])
    code, out = _main(word, capsys, "read", "--target", p)
    assert code == 0 and sum("BOXTWO" in u["text"] for u in out["paragraphs"]) == 1, out
    assert any("BOXOLD" in u["text"] for u in out["paragraphs"])


# H3: 콘텐츠 컨트롤 안의 표 번호 ---------------------------------------------------------------------------------------

def _sdt_docx(path):
    import docx
    from docx.oxml import parse_xml
    d = docx.Document()
    d.add_paragraph("서식 안내")
    t1 = d.add_table(rows=2, cols=1)
    t1.cell(0, 0).text, t1.cell(1, 0).text = "SDT표머리", "sdt값"
    d.add_paragraph("본문")
    t2 = d.add_table(rows=2, cols=1)
    t2.cell(0, 0).text, t2.cell(1, 0).text = "일반표머리", "일반값"
    sdt = parse_xml(f'<w:sdt {_NS}><w:sdtContent/></w:sdt>')
    t1._tbl.addprevious(sdt)
    sdt[0].append(t1._tbl)                                  # 첫 표를 콘텐츠 컨트롤 안으로
    d.save(path)
    return str(path)


def test_h3_word_table_numbers_match_read(tmp_path):
    p = _sdt_docx(tmp_path / "s.docx")
    units, _ = reader.parse_docx(p)
    assert [u["loc"] for u in units if u["kind"] == "row"][0].startswith("표 1") and "SDT표머리" in units[1]["text"]
    for n, head in ((1, "SDT표머리"), (2, "일반표머리")):
        rows, name = office._word_rows(types.SimpleNamespace(what="table", table=n, from_target=p))
        assert rows[0] == [head] and name == "s.docx", rows
    with pytest.raises(rt.InputError) as ei:
        office._word_rows(types.SimpleNamespace(what="table", table=3, from_target=p))
    assert "전체 2개" in str(ei.value)


# L1: 굵게 위치는 UTF-16 단위 -------------------------------------------------------------------------------------------

def _u16(v):
    b = v.encode("utf-16-le", "surrogatepass")
    return "".join(chr(x) for x in struct.unpack("<%dH" % (len(b) // 2), b))


def _back(s):
    return "".join(s).encode("utf-16-le", "surrogatepass").decode("utf-16-le")


class _U16Doc(fw.FakeWordDocument):
    """Word처럼 글을 UTF-16 단위로 담는 가짜 문서 (BMP 밖 글자는 2칸)."""
    def replace(self, s, e, v):
        super().replace(s, e, _u16(v))


@pytest.fixture
def u16_doc(monkeypatch):
    def insert_after(self, s):
        self.doc.replace(self.End, self.End, s)
        self.End += len(_u16(s))
    monkeypatch.setattr(fw.FakeRange, "InsertAfter", insert_after)
    return _U16Doc("")


def _bold_texts(d):
    return [_back(d.text[tag[1]:tag[2]]) for tag, k, v in d.log
            if isinstance(tag, tuple) and tag[0] == "font" and k == "Bold" and v]


def test_l1_bold_after_emoji(u16_doc):
    w = word.Writer(u16_doc, 0, None)
    w.markdown(f"{EMOJI} **출시** 완료 {EMOJI}{EMOJI} **두번째**")
    assert _bold_texts(u16_doc) == ["출시", "두번째"]
    assert _back(u16_doc.text).startswith(f"{EMOJI} 출시 완료")


def test_l1_bold_without_emoji_unchanged(u16_doc):
    w = word.Writer(u16_doc, 0, None)
    w.text("가 **나** 다")
    assert _bold_texts(u16_doc) == ["나"]
    assert word._u16(EMOJI + "a") == 3 and word._u16("가나") == 2


# L2: excel read --search count와 0건 ----------------------------------------------------------------------------------

def test_l2_excel_search_count_and_zero(tmp_path, capsys):
    p = _formula_xlsx(tmp_path / "c.xlsx", cached=True)
    code, out = _main(excel, capsys, "read", "--target", p, "--search", "없는값")
    assert code == 0 and out["count"] == 0 and out["shown"] == 0
    assert out["agent_directive"].startswith("VERIFIED: 0건.") and "조건을 바꿔 다시 실행하지 마세요" in out["agent_directive"]
    code, out = _main(excel, capsys, "read", "--target", p, "--search", "a")
    assert code == 0 and out["count"] == out["total_matches"] == 1 and out["shown"] == 1 and out["truncated"] is False


# L3: '---'는 구분선, 맨 앞 YAML 머리말은 넣지 않음 --------------------------------------------------------------------

def test_l3_front_matter_lines():
    assert rt.front_matter_lines("---\ntitle: 보고서\ndate: 2026-11-03\n---\n# 제목") == 4
    assert rt.front_matter_lines("\ufeff---\ntags:\n  - a\n...\n본문") == 4
    assert rt.front_matter_lines("---\n본문 문단입니다\n---\n") == 0          # key: 모양이 아님 -> 구분선
    assert rt.front_matter_lines("---\ntitle: x\n본문") == 0                  # 닫지 않음
    assert rt.front_matter_lines("본문\n---\ntitle: x\n---") == 0             # 첫 줄이 아님
    blocks = rt.parse_markdown_blocks("---\ntitle: x\n---\n# 제목\n\n---\n\n끝")
    assert blocks == [("front_matter", 3), ("heading", (1, "제목")), ("break", None), ("paragraph", "끝")]
    assert rt.looks_like_markdown("---\ntitle: x\n---\n본문")


def test_l3_word_rule_is_border_not_page_break(wd, capsys):
    code, out = test_word.run(capsys, "write", "--text", "# 제목\n\n앞 문단\n\n---\n\n뒤 문단")
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    doc = wd["doc"]
    assert "\x0c" not in doc.text and "구분선" in out["inserted"] and "warnings" not in out
    paras = doc.text.split("\r")
    i = paras.index("앞 문단")
    assert paras[i + 1] == "" and paras[i + 2] == "뒤 문단"
    borders = [e for e in doc.log if len(e) == 3 and isinstance(e[0], tuple) and e[0][0] == "border"]
    assert (("border", doc.text.index("뒤 문단") - 1, word.WD_BORDER_BOTTOM), "LineStyle", 1) in borders


def test_l3_word_front_matter_skipped_with_warning(wd, capsys):
    code, out = test_word.run(capsys, "write", "--text", "---\ntitle: 비밀 제목\nauthor: kim\n---\n본문 첫 줄")
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    assert "title" not in wd["doc"].text and "\x0c" not in wd["doc"].text and "본문 첫 줄" in wd["doc"].text
    assert out["warnings"] and "YAML 머리말(4줄" in out["warnings"][0]
    assert word._first_inserted("---\ntitle: x\n---\n---\n본문", "auto") == "본문"


def test_l3_ppt_md_sections_skip_rule_and_front_matter():
    secs = ppt._md_sections("---\ntitle: 비밀\n---\n# 첫\n본문\n\n---\n\n둘째 줄\n## 다음\n내용")
    assert [s["title"] for s in secs] == ["첫", "다음"]
    assert secs[0]["lines"] == [("본문", ""), ("둘째 줄", "")]
