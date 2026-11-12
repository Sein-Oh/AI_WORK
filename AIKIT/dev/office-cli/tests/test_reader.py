"""reader.py: 형식별 파서(docx/pptx/xlsx/pdf/hwpx/hwp/csv/txt), 장절·번호·쪽, 스캔 쪽, 캐시와 색인,
search/list/clear, CLI 계약. 샘플 파일은 테스트 안에서 만듭니다 (바이너리 고정 파일 없음).
"""
import json
import os
import re
import struct
import subprocess
import sys
import time
import types
import zipfile

import pytest

import _runtime as rt
import _xlsx_formula as xf
import office
import reader
import word
from office_helpers import recommended_argv, split_command, error_log_text, formula_xlsx, run_fresh, run_main, SECRET, set_cached_value, TOOL_DIR
from sample_docs import new_docx, new_pptx


def make_docx(path):
    d = new_docx()
    d.add_heading("추진 배경", level=1)
    d.add_paragraph("2026년 공정 최적화 마일스톤을 정리합니다.")
    t = d.add_table(rows=2, cols=3)
    a = t.cell(0, 0).merge(t.cell(0, 1))
    a.text = "병합 헤더"
    t.cell(0, 2).text = "값"
    t.cell(1, 0).text = "가"
    t.cell(1, 1).text = "나"
    t.cell(1, 2).text = "100"
    d.add_heading("일정", level=1)
    d.add_paragraph("3분기 양산 적용")
    d.save(path)


def make_pptx(path):
    from pptx.util import Inches
    prs = new_pptx()
    for i, title in enumerate(["첫 슬라이드", "두 번째 슬라이드", "세 번째 슬라이드"], 1):
        s = prs.slides.add_slide(prs.slide_layouts[5])  # 제목만
        s.shapes.title.text = title
        box = s.shapes.add_textbox(Inches(1), Inches(2), Inches(5), Inches(1))
        tf = box.text_frame
        p = tf.paragraphs[0]
        r1 = p.add_run()
        r1.text = "Hel"
        r2 = p.add_run()
        r2.text = f"lo 본문{i}"
        if i == 2:
            rows = s.shapes.add_table(2, 2, Inches(1), Inches(3), Inches(4), Inches(1)).table
            rows.cell(0, 0).text = "항목"
            rows.cell(0, 1).text = "케이블카"
            s.notes_slide.notes_text_frame.text = "발표 노트"
    # 슬라이드 순서 바꾸기: 3번째를 맨 앞으로 (파일 번호 순서와 실제 순서가 달라짐)
    sld_ids = prs.slides._sldIdLst
    last = sld_ids[-1]
    sld_ids.remove(last)
    sld_ids.insert(0, last)
    prs.save(path)


def make_xlsx(path):
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "매출"
    ws.append(["지역", "금액"])
    ws.append(["서울", 1200])
    ws.append([None, None])
    ws.append(["부산", 800.0])
    ws2 = wb.create_sheet("요약")
    ws2.append(["합계", "=SUM(매출!B2:B4)"])
    wb.save(path)
    set_cached_value(path, "<v></v>", "<v>2000</v>")   # 엑셀에서 저장한 파일처럼 수식 칸에 계산된 값


def make_pdf(path):
    pymupdf = reader._import_pymupdf()           # 메시지는 그때의 stderr로 (1.4.90)
    doc = pymupdf.open()
    for text in ("Phu Quoc cable car 7.9km", "Safari tour schedule"):
        page = doc.new_page()
        page.insert_text((72, 72), text)
    doc.save(path)
    doc.close()


def make_hwpx(path):
    ns = 'xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph"'
    sec0 = (f'<?xml version="1.0" encoding="UTF-8"?><hs:sec xmlns:hs="x" {ns}>'
            '<hp:p><hp:run><hp:t>결재 </hp:t><hp:t>일자</hp:t></hp:run>'
            '<hp:run><hp:tbl><hp:tr><hp:tc><hp:subList><hp:p><hp:run><hp:t>표 안 문단</hp:t></hp:run></hp:p>'
            '</hp:subList></hp:tc></hp:tr></hp:tbl></hp:run></hp:p>'
            '<hp:p><hp:run><hp:t>둘째<hp:tab/>문단</hp:t></hp:run></hp:p></hs:sec>')
    sec10 = f'<?xml version="1.0" encoding="UTF-8"?><hs:sec xmlns:hs="x" {ns}><hp:p><hp:run><hp:t>열번째 구역</hp:t></hp:run></hp:p></hs:sec>'
    sec2 = f'<?xml version="1.0" encoding="UTF-8"?><hs:sec xmlns:hs="x" {ns}><hp:p><hp:run><hp:t>두번째 구역</hp:t></hp:run></hp:p></hs:sec>'
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("mimetype", "application/hwp+zip")
        z.writestr("Contents/section0.xml", sec0)
        z.writestr("Contents/section10.xml", sec10)
        z.writestr("Contents/section2.xml", sec2)


class Samples(dict):
    """샘플 파일 경로. 만들지 못한 형식(예: DRM으로 python-pptx 기본 템플릿 손상)을 꺼내는 테스트만 건너뜁니다."""

    def __init__(self, files, broken):
        super().__init__(files)
        self.broken = broken

    def __getitem__(self, key):
        if key in self.broken:
            pytest.skip(self.broken[key])
        return super().__getitem__(key)


def build_samples(tmp_path):
    files = {
        "docx": tmp_path / "계획서.docx", "pptx": tmp_path / "발표.pptx", "xlsx": tmp_path / "매출.xlsx",
        "pdf": tmp_path / "여행.pdf", "hwpx": tmp_path / "양식.hwpx",
    }
    makers = {"docx": make_docx, "pptx": make_pptx, "xlsx": make_xlsx, "pdf": make_pdf, "hwpx": make_hwpx}
    broken = {}
    for kind, make in makers.items():
        try:
            make(str(files[kind]))
        except pytest.skip.Exception as e:      # 한 형식을 못 만들어도 다른 형식 테스트는 그대로 실행
            broken[kind] = str(e.msg) if getattr(e, "msg", None) else str(e)
    return Samples({k: str(v) for k, v in files.items() if k not in broken}, broken)


@pytest.fixture
def samples(tmp_path):
    return build_samples(tmp_path)


def texts(units):
    return [u["text"] for u in units]


def run_reader_env(tmp_env, *args):
    proc = subprocess.run([sys.executable, os.path.join(TOOL_DIR, "reader.py"), *args],
                          capture_output=True, timeout=120, env=tmp_env)
    lines = [ln for ln in proc.stdout.decode("utf-8").splitlines() if ln.strip()]
    assert len(lines) == 1, proc.stdout.decode("utf-8") + proc.stderr.decode("utf-8", "replace")
    return proc.returncode, json.loads(lines[0])


@pytest.fixture
def env(tmp_path):
    return dict(os.environ, OFFICE_CLI_DATA_DIR=str(tmp_path / "data"), OFFICE_CLI_CACHE_DIR=str(tmp_path / "cache"))


W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'


UNIT_KEYS = {"id", "kind", "loc", "page", "page_source", "section", "number", "heading", "text"}


def run_reader(*args):
    """reader.py를 하위 프로세스로 실행 (캐시/데이터 폴더는 conftest가 정한 환경변수)."""
    proc = subprocess.run([sys.executable, os.path.join(TOOL_DIR, "reader.py"), *args],
                          capture_output=True, timeout=120, env=dict(os.environ))
    lines = [ln for ln in proc.stdout.decode("utf-8").splitlines() if ln.strip()]
    assert len(lines) == 1, proc.stdout.decode("utf-8") + proc.stderr.decode("utf-8", "replace")
    return proc.returncode, json.loads(lines[0])


def by_text(units, start):
    return [u for u in units if u["text"].startswith(start)][0]


def make_text_pdf(path, pages, font="korea"):
    """pages: [[(글, 크기, 굵게)], ...]. 표는 ("TABLE", rows) 항목으로 선을 그려 넣습니다."""
    pymupdf = reader._import_pymupdf()           # 메시지는 그때의 stderr로 (1.4.90)
    doc = pymupdf.open()
    for items in pages:
        page = doc.new_page()
        y = 72
        for item in items:
            if item[0] == "TABLE":
                rows = item[1]
                x0, y0, cw, rh = 72, y, 150, 22
                for i in range(len(rows) + 1):
                    page.draw_line((x0, y0 + i * rh), (x0 + len(rows[0]) * cw, y0 + i * rh))
                for j in range(len(rows[0]) + 1):
                    page.draw_line((x0 + j * cw, y0), (x0 + j * cw, y0 + len(rows) * rh))
                for i, r in enumerate(rows):
                    for j, c in enumerate(r):
                        page.insert_text((x0 + j * cw + 4, y0 + i * rh + 15), c, fontname=font, fontsize=10)
                y = y0 + len(rows) * rh + 20
                continue
            text, size, bold = item
            fontname = ("hebo" if bold else "helv") if font == "helv" else font
            page.insert_text((72, y), text, fontname=fontname, fontsize=size)
            y += size + 8
    doc.save(str(path))
    doc.close()


def korean_pdf(path):
    make_text_pdf(path, [
        [("III. 개발 내용", 16, False), ("1. 개발 내용 종합", 14, False), ("가. 검증기술", 12, False),
         ("O 프로그램 전체 연구 목표", 10, False), ("검증 절차를 수행한다.", 10, False),
         ("TABLE", [["항목", "기준", "시험방법"], ["내전압", "1.5 kV 1분", "KS C 1234"], ["절연저항", "100 MΩ 이상", "메거"]])],
        [("나. 시험평가", 12, False), ("시험평가 본문입니다.", 10, False), ("2. 기대 효과", 14, False),
         ("기대 효과 본문입니다.", 10, False), ("IV. 활용 계획", 16, False), ("1. 사업화", 14, False)],
    ])


def english_pdf(path):
    make_text_pdf(path, [
        [("1. SCOPE", 16, True), ("This spec covers data.", 10, False), ("1.1. Scope", 13, True),
         ("1.1.1. Data", 11, True), ("Data body text.", 10, False)],
        [("1.2. Purpose", 13, True), ("Purpose body.", 10, False), ("2 REQUIREMENTS", 16, True),
         ("2.1 General", 13, True), ("Requirement body.", 10, False)],
    ], font="helv")


def hash_pdf(path):
    make_text_pdf(path, [
        [("1.0# Introduction", 14, False), ("intro text here", 10, False), ("1.1# Purpose", 12, False),
         ("1.1.1# Information", 11, False), ("info text", 10, False)],
        [("2.0# Design", 14, False), ("design text", 10, False)],
    ], font="helv")


def _numbering_part(d):
    """문서의 번호 매기기 파트. 없으면 만듭니다: 회사 PC처럼 python-docx 기본 템플릿이 보안 프로그램으로 바뀌어
    numbering.xml이 없으면 python-docx가 NotImplementedError를 냄 (2026-10-29 실기)."""
    try:
        return d.part.numbering_part
    except (KeyError, NotImplementedError):
        from docx.opc.constants import CONTENT_TYPE as CT, RELATIONSHIP_TYPE as RT
        from docx.opc.packuri import PackURI
        from docx.oxml import parse_xml
        from docx.parts.numbering import NumberingPart
        part = NumberingPart(PackURI("/word/numbering.xml"), CT.WML_NUMBERING, parse_xml(f"<w:numbering {W_NS}/>"),
                             d.part.package)
        d.part.relate_to(part, RT.NUMBERING)
        return d.part.numbering_part


def add_numbering(d):
    """numbering.xml에 3.로 시작하는 다단계 번호(3. / 3.1. / 3.1.1 / 가))를 넣고 numId 90을 돌려줍니다."""
    from docx.oxml import parse_xml
    numbering = _numbering_part(d).element
    numbering.insert(0, parse_xml(
        f'<w:abstractNum {W_NS} w:abstractNumId="90">'
        '<w:lvl w:ilvl="0"><w:start w:val="3"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>'
        '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2."/></w:lvl>'
        '<w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2.%3"/></w:lvl>'
        '<w:lvl w:ilvl="3"><w:start w:val="1"/><w:numFmt w:val="ganada"/><w:lvlText w:val="%4)"/></w:lvl>'
        '</w:abstractNum>'))
    numbering.append(parse_xml(f'<w:num {W_NS} w:numId="90"><w:abstractNumId w:val="90"/></w:num>'))


def numbered(d, text, ilvl):
    from docx.oxml import parse_xml
    p = d.add_paragraph(text)
    p._p.get_or_add_pPr().append(parse_xml(f'<w:numPr {W_NS}><w:ilvl w:val="{ilvl}"/><w:numId w:val="90"/></w:numPr>'))
    return p


def spec_docx(path, saved_layout=False):
    """제목 스타일 + 자동 번호 + 일반 번호 단락 + 강제 쪽 나눔 + 요구사항 표."""
    from docx.enum.text import WD_BREAK
    from docx.oxml import parse_xml
    d = new_docx()
    add_numbering(d)
    d.add_heading("시험 규격", level=1)
    d.add_paragraph("이 규격은 시험 방법을 정한다.")
    numbered(d, "일반 사항", 0)
    numbered(d, "적용 범위", 1)
    numbered(d, "세부 조건", 2)
    numbered(d, "온도 조건", 3)
    d.add_paragraph("시험 온도는 25도로 한다.")
    d.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
    d.add_heading("요구사항", level=1)
    d.add_paragraph("가. 검증기술")
    d.add_paragraph("O 프로그램 전체 연구 목표")
    t = d.add_table(rows=3, cols=3)
    for i, row in enumerate([["항목", "기준", "시험방법"], ["내전압", "1.5 kV 1분", "KS C\n1234"],
                             ["절연저항", "100 MΩ 이상", "메거"]]):
        for j, v in enumerate(row):
            t.cell(i, j).text = v
    d.add_paragraph("마지막 쪽 문단")
    if saved_layout:
        # Word가 저장한 쪽 경계 흉내: 3쪽 시작 (2쪽은 쪽 나눔 뒤에 Word가 기록)
        for p in d.paragraphs:
            if p.text in ("요구사항", "마지막 쪽 문단"):
                p.runs[0]._r.insert(0, parse_xml(f'<w:lastRenderedPageBreak {W_NS}/>'))
    d.save(str(path))


@pytest.fixture
def docs(tmp_path):
    folder = tmp_path / "규격"
    (folder / "sub").mkdir(parents=True)
    korean_pdf(folder / "계획서.pdf")
    english_pdf(folder / "spec.pdf")
    spec_docx(folder / "sub" / "규격.docx")
    (folder / "메일.msg").write_bytes(b"\xd0\xcf\x11\xe0")          # 지원하지 않는 형식 (1.5.0부터 .eml은 읽음)
    (folder / "그림.png").write_bytes(b"\x89PNG")
    return folder


def _old_cache(src, text="옛날 단락 내용"):
    src.write_text(text, encoding="utf-8")
    st = os.stat(src)
    path = os.path.abspath(str(src))
    old = {"file": path, "type": "pdf", "mtime": st.st_mtime, "size": st.st_size, "parser_version": 1,
           "parsed_at": "20261001_000000", "meta": {"pages": 3},
           "units": [{"loc": "p.3", "heading": "", "text": "옛날 캐시의 검증 문장"},
                     {"loc": "슬라이드 2", "heading": "제목", "text": "검증 슬라이드"}]}
    with open(reader._doc_cache_path(path), "w", encoding="utf-8") as f:
        json.dump(old, f, ensure_ascii=False)
    with open(reader._index_path(), "w", encoding="utf-8") as f:
        json.dump({path: {"type": "txt", "units": 2, "parsed_at": "20261001_000000"}}, f, ensure_ascii=False)
    return path


def make_pdf_at(path, pages):
    """pages: [[(x, y, 글, 크기)], ...] 위치를 정해 글을 넣습니다 (A4 595x842)."""
    pymupdf = reader._import_pymupdf()           # 메시지는 그때의 stderr로 (1.4.90)
    doc = pymupdf.open()
    for items in pages:
        page = doc.new_page()
        for x, y, text, size in items:
            page.insert_text((x, y), text, fontname="korea", fontsize=size)
    doc.save(str(path))
    doc.close()


def _cached(tmp_path, *names):
    paths = []
    for n in names:
        f = tmp_path / n
        f.write_text(f"{n} 내용", encoding="utf-8")
        run_reader("index", "--file", str(f))
        paths.append(f)
    return paths


def _pixmap(pymupdf):
    pix = pymupdf.Pixmap(pymupdf.csRGB, pymupdf.IRect(0, 0, 20, 20), False)
    pix.clear_with(200)
    return pix


def _pdf(path, pages):
    """pages: 쪽마다 'text'(본문), 'header'(머리글·쪽 번호만), 'image'(그림만)."""
    pymupdf = reader._import_pymupdf()           # 메시지는 그때의 stderr로 (1.4.90)
    doc = pymupdf.open()
    for i, kind in enumerate(pages, 1):
        page = doc.new_page()
        if kind in ("text", "header"):
            page.insert_text((72, 30), "SPEC DOCUMENT", fontname="helv", fontsize=9)
            page.insert_text((280, 820), f"- {i} -", fontname="helv", fontsize=9)
        if kind == "text":
            page.insert_text((72, 200), f"body text on page {i}", fontname="helv", fontsize=11)
        if kind == "image":
            page.insert_image(pymupdf.Rect(40, 40, 560, 800), pixmap=_pixmap(pymupdf))   # 쪽을 덮는 그림 (1.4.12)
    doc.save(str(path))
    doc.close()


def _old_caches(tmp_path, names, remove=()):
    index = {}
    for k, n in enumerate(names):
        src = tmp_path / n
        src.write_text(f"{n} 검증 문장", encoding="utf-8")
        st = os.stat(src)
        path = os.path.abspath(str(src))
        stamp = f"2026100{k}_000000"
        old = {"file": path, "type": "txt", "mtime": st.st_mtime, "size": st.st_size, "parser_version": 1,
               "parsed_at": stamp, "meta": {}, "units": [{"loc": "1", "heading": "", "text": "옛 검증 문장"}]}
        with open(reader._doc_cache_path(path), "w", encoding="utf-8") as f:
            json.dump(old, f, ensure_ascii=False)
        index[path] = {"type": "txt", "units": 1, "parsed_at": stamp}
        if n in remove:
            os.remove(src)
    with open(reader._index_path(), "w", encoding="utf-8") as f:
        json.dump(index, f, ensure_ascii=False)


def _pix_large(pymupdf, w=200, h=280):
    pix = pymupdf.Pixmap(pymupdf.csRGB, pymupdf.IRect(0, 0, w, h), False)
    pix.clear_with(200)
    return pix


def _docs(tmp_path, n=3):
    folder = tmp_path / "docs"
    folder.mkdir()
    for i in range(n):
        (folder / f"d{i}.txt").write_text(f"문서 {i} 내용 토크", encoding="utf-8")
    return folder


def _index_json():
    with open(reader._index_path(), encoding="utf-8") as f:
        return json.load(f)


def _pix(pymupdf):
    pix = pymupdf.Pixmap(pymupdf.csRGB, pymupdf.IRect(0, 0, 40, 56), False)
    pix.clear_with(200)
    return pix


def _scan_pdf(tmp_path, stamp_page=None, pages=2, scan_all=True):
    pymupdf = reader._import_pymupdf()           # 메시지는 그때의 stderr로 (1.4.90)
    doc = pymupdf.open()
    for i in range(1, pages + 1):
        p = doc.new_page()
        if scan_all or i == stamp_page:
            p.insert_image(pymupdf.Rect(30, 30, 565, 812), pixmap=_pix(pymupdf))
            if i == stamp_page:
                p.insert_text((400, 400), "RECEIVED", fontsize=9)
        else:
            p.insert_text((72, 200), f"Body text page {i} torque {i * 10} Nm.", fontsize=11)
    path = tmp_path / ("scan.pdf" if scan_all else "stamp.pdf")
    doc.save(str(path))
    return path


class _Stream:
    def __init__(self, data):
        self.data = data

    def read(self):
        return self.data


def _fake_ole(monkeypatch, open_error=None, section=b"not zlib"):
    import olefile

    class Ole:
        def __init__(self, path):
            if open_error:
                raise open_error

        def exists(self, name):
            return True

        def openstream(self, name):
            if name == "FileHeader":
                return _Stream(b"\0" * 36 + (1).to_bytes(4, "little"))     # 압축된 본문
            return _Stream(section)

        def listdir(self):
            return [["BodyText", "Section0"]]

        def close(self):
            pass
    monkeypatch.setattr(olefile, "isOleFile", lambda p: True)
    monkeypatch.setattr(olefile, "OleFileIO", Ole)


def _ole_without_fileheader(path):
    """Root Entry만 있는 최소 OLE 파일 (이름만 .hwp인 다른 OLE 문서 흉내)."""
    h = bytearray(512)
    h[0:8] = bytes.fromhex("D0CF11E0A1B11AE1")
    struct.pack_into("<HHHHH", h, 24, 0x3E, 3, 0xFFFE, 9, 6)
    struct.pack_into("<IIIIIIIII", h, 40, 0, 1, 1, 0, 4096, 0xFFFFFFFE, 0, 0xFFFFFFFE, 0)
    struct.pack_into("<109I", h, 76, *([0] + [0xFFFFFFFF] * 108))
    fats = struct.pack("<128I", *([0xFFFFFFFD, 0xFFFFFFFE] + [0xFFFFFFFF] * 126))
    d = bytearray(512)
    name = "Root Entry".encode("utf-16le") + b"\0\0"
    d[0:len(name)] = name
    struct.pack_into("<HBB", d, 64, len(name), 5, 1)
    struct.pack_into("<III", d, 68, 0xFFFFFFFF, 0xFFFFFFFF, 0xFFFFFFFF)
    struct.pack_into("<II", d, 116, 0xFFFFFFFE, 0)
    for i in range(1, 4):
        struct.pack_into("<III", d, i * 128 + 68, 0xFFFFFFFF, 0xFFFFFFFF, 0xFFFFFFFF)
    with open(path, "wb") as f:
        f.write(bytes(h) + fats + bytes(d))


@pytest.fixture
def one_file(tmp_path):
    p = tmp_path / "a.txt"
    p.write_text("hello", encoding="utf-8")
    return p


_NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
       'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
       'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
       'xmlns:v="urn:schemas-microsoft-com:vml"')


def _txbx(text):
    return f'<w:txbxContent><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:txbxContent>'


def _textbox_docx(path, body=True):
    import docx
    from docx.oxml import parse_xml
    d = new_docx()
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


def _sdt_docx(path):
    import docx
    from docx.oxml import parse_xml
    d = new_docx()
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


# ---------------------------------------------------------------------------
# 샘플 파일 생성
# ---------------------------------------------------------------------------


def test_samples_skip_only_broken_kind(tmp_path):
    s = Samples({"docx": "a.docx"}, {"pptx": "python-pptx 기본 템플릿이 손상됨 (DRM 추정)"})
    assert s["docx"] == "a.docx"
    with pytest.raises(pytest.skip.Exception, match="DRM"):
        s["pptx"]


@pytest.mark.parametrize("docx_default", ["encrypted", "stub"])
def test_drm_damaged_templates_still_make_samples(tmp_path, monkeypatch, docx_default):
    """DRM이 라이브러리 기본 템플릿을 망가뜨린 PC 흉내 (2026-11-05 실기: default.pptx 없음, default.docx는 스타일 4개뿐인
    zip): docx는 폴더 템플릿으로, pptx는 코드로 만든 최소 프레젠테이션으로 만들어 건너뛰지 않고 시험한다."""
    import docx
    import sample_docs
    enc = tmp_path / "default.enc"
    enc.write_bytes(b"DRMONE\x00\x01" + os.urandom(64))
    real_default, real_folder = sample_docs.docx_template_paths()
    bad_docx = enc
    if docx_default == "stub":                     # zip이지만 제목 스타일이 없고 빈 단락이 있는 문서
        bad_docx = tmp_path / "stub.docx"
        d = sample_docs.docx_from_folder(real_folder)
        for st in [s for s in d.styles if s.name.startswith("Heading")]:
            st.element.getparent().remove(st.element)
        d.add_paragraph("")
        d.save(str(bad_docx))
        monkeypatch.setattr(docx, "Document", lambda docx=None, _real=docx.Document: _real(docx or str(bad_docx)))
    monkeypatch.setattr(sample_docs, "docx_template_paths", lambda: (str(bad_docx), real_folder))
    monkeypatch.setattr(sample_docs, "pptx_template_path", lambda: str(enc))
    out = tmp_path / "out"
    out.mkdir()
    s = build_samples(out)
    assert set(s.keys()) == {"docx", "pptx", "xlsx", "pdf", "hwpx"}
    units, _ = reader.parse_docx(s["docx"])
    assert units[0]["text"] == "추진 배경"
    units, _ = reader.parse_pptx(s["pptx"])
    assert [u["heading"] for u in units] == ["세 번째 슬라이드", "첫 슬라이드", "두 번째 슬라이드"]
    assert "항목 | 케이블카" in units[2]["text"] and "[노트] 발표 노트" in units[2]["text"]


# ---------------------------------------------------------------------------
# 파서
# ---------------------------------------------------------------------------


def test_docx_order_heading_and_merged_cells(samples):
    units, meta = reader.parse_docx(samples["docx"])
    assert texts(units)[0] == "추진 배경"
    rows = [u for u in units if u["loc"].startswith("표 1")]
    assert [u["kind"] for u in rows] == ["row", "row"]           # 1.3.0: 표는 행 단위 (1.4.92: 머리글 행도 단위)
    assert rows[0]["text"] == "병합 헤더 | 값"
    assert rows[1]["text"] == "병합 헤더: 가 | 병합 헤더: 나 | 값: 100"   # 가로 병합 머리글은 이어 씀
    assert rows[1]["heading"] == "추진 배경"
    assert [u for u in units if u["text"] == "3분기 양산 적용"][0]["heading"] == "일정"
    assert meta["tables"] == 1


def test_pptx_real_slide_order_runs_table_notes(samples):
    units, meta = reader.parse_pptx(samples["pptx"])
    assert [u["heading"] for u in units] == ["세 번째 슬라이드", "첫 슬라이드", "두 번째 슬라이드"]
    assert "Hello 본문1" in units[1]["text"]          # 서식 경계에서 단어가 쪼개지지 않음
    assert "항목 | 케이블카" in units[2]["text"]
    assert "[노트] 발표 노트" in units[2]["text"]


def test_xlsx_rows_skip_empty_and_numbers(samples):
    units, meta = reader.parse_xlsx(samples["xlsx"])
    assert meta["sheets"] == ["매출", "요약"]
    assert [u["loc"] for u in units[:3]] == ["매출!1행", "매출!2행", "매출!4행"]
    assert units[2]["text"] == "부산\t800"
    assert units[3]["text"] == "합계\t2000" and "uncomputed_formulas" not in meta   # 계산된 값이 있는 수식 칸


def test_pdf_pages(samples):
    units, meta = reader.parse_pdf(samples["pdf"])
    assert meta["pages"] == 2 and units[0]["loc"] == "p.1"
    assert "7.9km" in units[0]["text"]


def test_hwpx_paragraphs_nested_tables_and_section_order(samples):
    units, _ = reader.parse_hwpx(samples["hwpx"])
    assert texts(units) == ["결재 일자", "표 안 문단", "둘째\t문단", "두번째 구역", "열번째 구역"]


def test_hwp_para_text_control_characters():
    def wchar(c):
        return c.to_bytes(2, "little")
    ext_ctrl = wchar(11) + b"\x00" * 14        # 표/개체 (확장 컨트롤, 8 WCHAR)
    sec_ctrl = wchar(2) + b"dces" + b"\x00" * 10  # 구역 정의 (확장 컨트롤)
    rec = (sec_ctrl + "가나".encode("utf-16-le") + ext_ctrl + "다".encode("utf-16-le")
           + wchar(30) + "라".encode("utf-16-le") + wchar(9) + b"\x00" * 14 + "마".encode("utf-16-le") + wchar(13))
    assert reader._hwp_para_text(rec) == "가나다 라\t마"


def test_encrypted_office_file_gives_drm_hint(tmp_path):
    fake = tmp_path / "보안.docx"
    fake.write_bytes(b"DRMONE\x00\x01" + os.urandom(64))
    with pytest.raises(rt.InputError) as ei:
        reader.parse_docx(str(fake))
    assert "DRM" in str(ei.value)


def test_old_office_format_is_redirected():
    with pytest.raises(rt.InputError) as ei:
        reader.file_type("a.xls")
    assert "excel.py" in str(ei.value)


def test_matcher_and_or_and_snippet():
    m = reader.build_matcher("케이블카;사파리|단가", False)
    assert m("케이블카와 사파리") == ["케이블카", "사파리"]
    assert m("케이블카만") is None
    assert m("견적 단가") == ["단가"]
    snip = reader.make_snippet("가" * 200 + "단가표" + "나" * 200, ["단가"], 10)
    assert snip.startswith("...") and "**단가**" in snip


def test_cache_reused_and_invalidated(samples, monkeypatch):
    calls = []
    orig = reader.PARSERS["docx"]
    monkeypatch.setitem(reader.PARSERS, "docx", lambda p: (calls.append(p), orig(p))[1])
    reader.load_document(samples["docx"])
    reader.load_document(samples["docx"])
    assert len(calls) == 1
    os.utime(samples["docx"], (1, 1))
    reader.load_document(samples["docx"])
    assert len(calls) == 2
    assert os.path.abspath(samples["docx"]) in reader.load_index()


# ---------------------------------------------------------------------------
# 번호 분류와 제목 판정
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("line,family,number,depth", [
    ("III. 개발 내용", "roman", "III", 1),
    ("\u2162. 개발 내용", "roman", "III", 1),                 # 전각 로마 숫자
    ("1. SCOPE", "arabic", "1", 1),
    ("1.1. Scope", "arabic", "1.1", 2),
    ("1.1 Scope", "arabic", "1.1", 2),
    ("1.1.1. Data", "arabic", "1.1.1", 3),
    ("1.0# Introduction", "arabic", "1.0", 1),
    ("1.1# Purpose", "arabic", "1.1", 2),
    ("가. 검증기술", "ganada", "가", 1),
    ("(1) 시험 조건", "paren_arabic", "(1)", 1),
    ("1) 시험 조건", "rparen_arabic", "1)", 1),
    ("(가) 세부", "paren_ganada", "(가)", 1),
    ("가) 세부", "rparen_ganada", "가)", 1),
    ("\u2460 첫째 항", "circled", "(1)", 1),                   # 원문자 입력 -> 출력은 (1)
    ("제1장 총칙", "ko_jang", "제1장", 1),
    ("제2절 적용", "ko_jeol", "제2절", 1),
    ("제12조(목적) 이 규정은 시험 절차를 정한다.", "ko_jo", "제12조", 1),
    ("Chapter 1 Overview", "chapter", "Chapter 1", 1),
    ("Annex A Test data", "annex", "Annex A", 1),
    ("부록 1 시험 양식", "annex_ko", "부록 1", 1),
    ("Q3. 만족도는?", "question", "Q3", 1),
    ("A. Purpose", "alpha_upper", "A", 1),
])
def test_classify_number_families(line, family, number, depth):
    info = reader.classify_number(line)
    assert info and info["family"] == family and info["number"] == number and info["depth"] == depth


@pytest.mark.parametrize("line", ["O 프로그램 전체 연구 목표", "- 하이픈 항목", "\u25cb 동그라미 항목", "\u25a1 네모 항목",
                                  "\u00b7 가운뎃점 항목", "* 별표 항목"])
def test_bullets_have_no_number(line):
    info = reader.classify_number(line)
    assert info["family"] == "bullet" and info["number"] is None
    assert reader.heading_title(info, line) is None


@pytest.mark.parametrize("line", ["2026. 10. 28. 회의", "3.5kg 이하", "본문 1. 안의 번호", "2026년 계획"])
def test_numbers_inside_text_or_dates_are_not_numbers(line):
    info = reader.classify_number(line)
    assert info is None or info["plain"] or info["number"] is None


@pytest.mark.parametrize("line,is_heading", [
    ("1. 개발 내용 종합", True),
    ("1. 시스템은 입력을 검증한다.", False),            # 문장 끝
    ("가. 성능을 확인하였음.", False),
    ("가. 시험 결과를 별도로 정리함", False),           # 개조식 문장 끝
    ("가. 소음", True),                                # 짧은 명사는 제목
    ("1. 개요", True),
    ("1. " + "가" * 90, False),                         # 너무 김
    ("2 REQUIREMENTS", False),                          # 구분 기호 없는 숫자는 굵게/큰 글씨일 때만
])
def test_heading_guard(line, is_heading):
    assert bool(reader.heading_title(reader.classify_number(line), line)) is is_heading


def test_plain_number_heading_needs_strong_signal():
    line = "2 REQUIREMENTS"
    assert reader.heading_title(reader.classify_number(line), line, strong=True) == line


def test_article_heading_keeps_number_and_title_only():
    line = "제12조(목적) 이 규정은 시험 절차를 정한다."
    assert reader.heading_title(reader.classify_number(line), line) == "제12조(목적)"


def test_outline_first_appearance_levels():
    b = reader.UnitBuilder()
    for t in ["III. 개발 내용", "1. 개발 내용 종합", "가. 검증기술", "O 프로그램 전체 연구 목표", "나. 시험", "2. 기대 효과",
              "IV. 활용", "1. 사업화"]:
        b.add_text("x", t)
    sec = {u["text"]: u["section"] for u in b.units}
    assert sec["O 프로그램 전체 연구 목표"] == ["III. 개발 내용", "1. 개발 내용 종합", "가. 검증기술"]
    assert sec["나. 시험"] == ["III. 개발 내용", "1. 개발 내용 종합", "나. 시험"]
    assert sec["2. 기대 효과"] == ["III. 개발 내용", "2. 기대 효과"]
    assert sec["1. 사업화"] == ["IV. 활용", "1. 사업화"]
    bullet = [u for u in b.units if u["text"].startswith("O ")][0]
    assert bullet["kind"] == "paragraph" and bullet["number"] is None


# ---------------------------------------------------------------------------
# 표 행
# ---------------------------------------------------------------------------


def test_table_rows_header_labels_merged_and_multiline():
    rows = reader.table_rows([["항목", "기준", "", "시험방법"],
                              ["내전압", "1.5 kV", "1분", "KS C\n1234"],
                              ["", "", "", ""],
                              ["절연저항", "100 MΩ 이상", None, "메거"]])
    assert [r["row"] for r in rows] == [1, 2, 3] and rows[0]["text"] == "항목 | 기준 | 시험방법"   # 1.4.92: 머리글 행도 단위
    assert rows[1]["text"] == "항목: 내전압 | 기준: 1.5 kV | 기준: 1분 | 시험방법: KS C 1234"
    assert rows[2]["text"] == "항목: 절연저항 | 기준: 100 MΩ 이상 | 시험방법: 메거"


def test_table_rows_without_header_and_row_number():
    rows = reader.table_rows([["1", "내전압 1.5 kV 이상이어야 하며 1분간 견딜 것"], ["REQ-002", "절연저항 100 MΩ 이상"]])
    assert rows[0]["text"] == "1 | 내전압 1.5 kV 이상이어야 하며 1분간 견딜 것" and rows[0]["number"] == "1"
    assert rows[1]["number"] == "REQ-002"


# ---------------------------------------------------------------------------
# 형식별 파서: 공통 키, 장절, 번호, 쪽
# ---------------------------------------------------------------------------


def test_pdf_korean_sections_pages_table_rows(tmp_path):
    f = tmp_path / "계획서.pdf"
    korean_pdf(f)
    units, meta = reader.parse_pdf(str(f))
    assert all(set(u) == UNIT_KEYS for u in units)
    assert [u["id"] for u in units] == list(range(1, len(units) + 1))
    bullet = by_text(units, "O 프로그램")
    assert bullet["section"] == ["III. 개발 내용", "1. 개발 내용 종합", "가. 검증기술"] and bullet["number"] is None
    assert bullet["page"] == 1 and bullet["page_source"] == "pdf" and bullet["loc"] == "p.1"
    row = [u for u in units if u["kind"] == "row" and "내전압" in u["text"]][0]
    assert row["text"] == "항목: 내전압 | 기준: 1.5 kV 1분 | 시험방법: KS C 1234"
    assert row["section"][-1] == "가. 검증기술" and row["page"] == 1
    assert not [u for u in units if u["kind"] == "paragraph" and "KS C 1234" in u["text"]]   # 표 글은 단락에 중복 없음
    h = by_text(units, "나. 시험평가")
    assert h["kind"] == "heading" and h["page"] == 2 and h["number"] == "나"
    assert by_text(units, "기대 효과 본문")["section"] == ["III. 개발 내용", "2. 기대 효과"]
    assert by_text(units, "1. 사업화")["section"] == ["IV. 활용 계획", "1. 사업화"]
    assert meta["pages"] == 2


def test_pdf_english_dotted_with_font_signal(tmp_path):
    f = tmp_path / "spec.pdf"
    english_pdf(f)
    units, _ = reader.parse_pdf(str(f))
    assert by_text(units, "Data body")["section"] == ["1. SCOPE", "1.1. Scope", "1.1.1. Data"]
    purpose = by_text(units, "1.2. Purpose")
    assert purpose["section"] == ["1. SCOPE", "1.2. Purpose"] and purpose["number"] == "1.2" and purpose["page"] == 2
    assert by_text(units, "Requirement body")["section"] == ["2 REQUIREMENTS", "2.1 General"]   # '2 ' 굵게 큰 글씨


def test_pdf_hash_numbering(tmp_path):
    f = tmp_path / "hash.pdf"
    hash_pdf(f)
    units, _ = reader.parse_pdf(str(f))
    info = by_text(units, "info text")
    assert info["section"] == ["1.0# Introduction", "1.1# Purpose", "1.1.1# Information"]
    assert by_text(units, "1.1.1# Information")["number"] == "1.1.1"
    design = by_text(units, "design text")
    assert design["section"] == ["2.0# Design"] and design["page"] == 2


def test_docx_styles_auto_numbering_page_breaks_and_table_rows(tmp_path):
    f = tmp_path / "규격.docx"
    spec_docx(f)
    units, meta = reader.parse_docx(str(f))
    assert all(set(u) == UNIT_KEYS for u in units)
    assert meta["page_source"] == "page_break"
    sub = by_text(units, "3.1.1 세부 조건")                    # Word 자동 번호를 글 앞에 붙임
    assert sub["number"] == "3.1.1"
    assert sub["section"] == ["시험 규격", "3. 일반 사항", "3.1. 적용 범위", "3.1.1 세부 조건"]
    assert by_text(units, "가) 온도 조건")["number"] == "가)"
    body = by_text(units, "시험 온도는")
    assert body["page"] == 1 and body["section"][-1] == "가) 온도 조건"
    h = by_text(units, "요구사항")
    assert h["kind"] == "heading" and h["page"] == 2 and h["section"] == ["요구사항"]
    assert by_text(units, "O 프로그램")["section"] == ["요구사항", "가. 검증기술"]
    row = [u for u in units if u["kind"] == "row" and "내전압" in u["text"]][0]
    assert row["text"] == "항목: 내전압 | 기준: 1.5 kV 1분 | 시험방법: KS C 1234"
    assert row["loc"] == "표 1 2행" and row["page"] == 2 and row["section"] == ["요구사항", "가. 검증기술"]


def test_docx_saved_layout_pages(tmp_path):
    f = tmp_path / "저장됨.docx"
    spec_docx(f, saved_layout=True)
    units, meta = reader.parse_docx(str(f))
    assert meta["page_source"] == "word_saved_layout"
    last = by_text(units, "마지막 쪽 문단")
    assert last["page"] == 3 and last["page_source"] == "word_saved_layout"
    assert by_text(units, "시험 온도는")["page"] == 1


def test_docx_without_page_info_has_null_pages(tmp_path):
    d = new_docx()
    d.add_paragraph("1. 개요")
    d.add_paragraph("본문입니다.")
    d.save(str(tmp_path / "a.docx"))
    units, meta = reader.parse_docx(str(tmp_path / "a.docx"))
    assert all(u["page"] is None and u["page_source"] is None for u in units)
    assert units[1]["section"] == ["1. 개요"]


def test_pptx_slide_section_and_page(tmp_path):
    prs = new_pptx()
    for title in ["1. 개요", "2. 일정"]:
        s = prs.slides.add_slide(prs.slide_layouts[5])
        s.shapes.title.text = title
    prs.save(str(tmp_path / "발표.pptx"))
    units, _ = reader.parse_pptx(str(tmp_path / "발표.pptx"))
    assert [(u["kind"], u["page"], u["page_source"], u["section"], u["number"]) for u in units] == [
        ("slide", 1, "slide", ["1. 개요"], "1"), ("slide", 2, "slide", ["2. 일정"], "2")]


def test_xlsx_rows_section_is_sheet(tmp_path):
    import openpyxl
    wb = openpyxl.Workbook()
    wb.active.title = "매출"
    wb.active.append(["1.1", "서울"])
    wb.save(str(tmp_path / "a.xlsx"))
    units, _ = reader.parse_xlsx(str(tmp_path / "a.xlsx"))
    u = units[0]
    assert (u["kind"], u["section"], u["page"], u["number"], u["loc"]) == ("row", ["매출"], None, None, "매출!1행")


def test_hwpx_outline_levels_and_numbers(tmp_path):
    hp = 'xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph"'
    hh = 'xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head"'
    header = (f'<?xml version="1.0" encoding="UTF-8"?><hh:head {hh}><hh:refList>'
              '<hh:numberings><hh:numbering id="1" start="0">'
              '<hh:paraHead start="1" level="1" numFormat="DIGIT">^1.</hh:paraHead>'
              '<hh:paraHead start="1" level="2" numFormat="HANGUL_SYLLABLE">^2.</hh:paraHead>'
              '</hh:numbering></hh:numberings><hh:paraProperties>'
              '<hh:paraPr id="1"><hh:heading type="OUTLINE" idRef="1" level="0"/></hh:paraPr>'
              '<hh:paraPr id="2"><hh:heading type="OUTLINE" idRef="1" level="1"/></hh:paraPr>'
              '<hh:paraPr id="0"><hh:heading type="NONE" idRef="0" level="0"/></hh:paraPr>'
              '</hh:paraProperties></hh:refList></hh:head>')
    sec = (f'<?xml version="1.0" encoding="UTF-8"?><hs:sec xmlns:hs="x" {hp}>'
           '<hp:p paraPrIDRef="1"><hp:run><hp:t>개요</hp:t></hp:run></hp:p>'
           '<hp:p paraPrIDRef="2"><hp:run><hp:t>목적</hp:t></hp:run></hp:p>'
           '<hp:p paraPrIDRef="0"><hp:run><hp:t>본문 내용</hp:t></hp:run></hp:p>'
           '<hp:p paraPrIDRef="1" pageBreak="1"><hp:run><hp:t>범위</hp:t></hp:run></hp:p></hs:sec>')
    path = tmp_path / "양식.hwpx"
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("Contents/header.xml", header)
        z.writestr("Contents/section0.xml", sec)
    units, _ = reader.parse_hwpx(str(path))
    assert [u["text"] for u in units] == ["1. 개요", "가. 목적", "본문 내용", "2. 범위"]
    assert units[2]["section"] == ["1. 개요", "가. 목적"]
    assert units[3]["section"] == ["2. 범위"] and units[3]["number"] == "2" and units[3]["page"] == 2


# ---------------------------------------------------------------------------
# PDF 머리글/바닥글, 2단 편집
# ---------------------------------------------------------------------------


def test_pdf_repeated_header_and_page_number_do_not_break_sections(tmp_path):
    p = tmp_path / "hdr.pdf"
    head = (72, 40, "제2장 시스템 요구사항", 9)
    make_pdf_at(p, [
        [head, (72, 120, "2.1 전기 요구사항", 14), (72, 150, "2.1.1 정격", 12), (72, 180, "정격 전압은 220 V로 한다.", 10),
         (290, 815, "- 1 -", 9)],
        [head, (72, 120, "정격 전류는 10 A 이하로 한다.", 10), (72, 150, "10.5 kV 이하 기기는 별도로 시험한다.", 10),
         (290, 815, "- 2 -", 9)],
        [head, (72, 120, "2.2 기계 요구사항", 14), (72, 150, "기계 본문입니다.", 10), (290, 815, "- 3 -", 9)],
    ])
    units, meta = reader.parse_pdf(str(p))
    assert meta["header_footer_removed"] == 6
    assert not [u for u in units if u["text"].startswith("- ")]
    # 1.4.92: 반복 머리글은 장절을 끊지 않고, 문서 번호처럼 검색되게 한 번만 따로 단위로 (2026-11-09 검토 A13)
    hdr = [u for u in units if "시스템 요구사항" in u["text"]]
    assert len(hdr) == 1 and hdr[0]["section"] == ["머리글·바닥글"] and hdr[0]["loc"] == "p.1 머리글·바닥글", hdr
    cur = by_text(units, "정격 전류는")
    assert cur["section"] == ["2.1 전기 요구사항", "2.1.1 정격"] and cur["page"] == 2      # 머리글이 장절을 끊지 않음
    kv = by_text(units, "10.5 kV")
    assert kv["kind"] != "heading" and kv["section"] == ["2.1 전기 요구사항", "2.1.1 정격"]   # 수치 줄은 제목 아님
    assert by_text(units, "기계 본문")["section"] == ["2.2 기계 요구사항"]


def test_pdf_two_columns_read_left_column_first(tmp_path):
    p = tmp_path / "col.pdf"
    make_pdf_at(p, [[
        (72, 80, "보고서 제목 전체 폭으로 길게 쓴 줄입니다 가나다라마바사아자차카타파하", 12),
        (60, 130, "1. 왼쪽 장", 13), (60, 160, "왼쪽 본문 첫째 문장이다.", 10), (60, 190, "왼쪽 본문 둘째 문장이다.", 10),
        (320, 130, "3. 오른쪽 장", 13), (320, 160, "오른쪽 본문 첫째 문장이다.", 10), (320, 190, "오른쪽 본문 둘째 문장이다.", 10),
    ]])
    units, _ = reader.parse_pdf(str(p))
    texts = [u["text"] for u in units]
    assert texts.index("1. 왼쪽 장") < texts.index("왼쪽 본문 둘째 문장이다.") < texts.index("3. 오른쪽 장")
    assert by_text(units, "왼쪽 본문 첫째")["section"] == ["1. 왼쪽 장"]
    assert by_text(units, "오른쪽 본문 둘째")["section"] == ["3. 오른쪽 장"]


@pytest.mark.parametrize("line", ["10.5 kV 이하", "1.5 kV 1분", "3 mm 이상", "2 개월", "1.2 설계 기준 ........ 12",
                                  "(3) Verification Review; and", "가) 시험 장비,", "1) 설계 검토 및",
                                  "1. 사업 개요 ······ 3", "2.1 범위\t5"])
def test_measurements_and_toc_lines_are_not_headings(line):
    assert reader.heading_title(reader.classify_number(line), line, strong=True) is None


@pytest.mark.parametrize("line", ["1. A형 시험", "2.1 개요", "3. 분석 방법", "1 일반 사항"])
def test_real_headings_still_detected(line):
    assert reader.heading_title(reader.classify_number(line), line, strong=True) == line


# ---------------------------------------------------------------------------
# DOCX 변경 추적·콘텐츠 컨트롤·여러 쪽 표·빈 번호 단락
# ---------------------------------------------------------------------------


def test_docx_tracked_insertions_content_controls_and_table_row_pages(tmp_path):
    from docx.oxml import parse_xml
    d = new_docx()
    p = d.add_paragraph("납기는 ")
    p._p.append(parse_xml(f'<w:ins {W_NS} w:id="1" w:author="a"><w:r><w:t>30일 이내</w:t></w:r></w:ins>'))
    p._p.append(parse_xml(f'<w:del {W_NS} w:id="2" w:author="a"><w:r><w:delText>60일</w:delText></w:r></w:del>'))
    body = d.element.body
    body.insert(len(body) - 1, parse_xml(
        f'<w:sdt {W_NS}><w:sdtContent><w:p><w:r><w:t>콘텐츠 컨트롤 안의 보증 기간 조항</w:t></w:r></w:p></w:sdtContent></w:sdt>'))
    t = d.add_table(rows=4, cols=2)
    for i, row in enumerate([["번호", "요구사항"], ["R-1", "첫째"], ["R-2", "둘째"], ["R-3", "셋째"]]):
        for j, v in enumerate(row):
            t.cell(i, j).text = v
    # R-3 행 앞에서 쪽이 바뀜 (Word가 저장한 쪽 경계)
    t.cell(3, 0).paragraphs[0].runs[0]._r.insert(0, parse_xml(f'<w:lastRenderedPageBreak {W_NS}/>'))
    path = tmp_path / "trk.docx"
    d.save(str(path))
    units, _ = reader.parse_docx(str(path))
    assert by_text(units, "납기는")["text"] == "납기는 30일 이내"                   # 지운 글(60일)은 없음
    assert by_text(units, "콘텐츠 컨트롤 안의")
    rows = {u["number"]: u for u in units if u["kind"] == "row"}
    assert rows["R-1"]["page"] == 1 and rows["R-3"]["page"] == 2


def test_docx_empty_numbered_paragraph_keeps_counting(tmp_path):
    d = new_docx()
    add_numbering(d)
    numbered(d, "일반 사항", 0)          # 3.
    numbered(d, "", 0)                   # 4. (빈 줄이어도 Word는 번호를 매김)
    numbered(d, "시험 방법", 0)          # 5.
    path = tmp_path / "num.docx"
    d.save(str(path))
    units, _ = reader.parse_docx(str(path))
    assert by_text(units, "5. 시험 방법")


# ---------------------------------------------------------------------------
# 긴 단위 잘림, 일부만 스캔된 PDF, 시트 행 상한, 저장 확인, 캐시 저장 실패, doc 호환 채움 글자
# ---------------------------------------------------------------------------


def test_long_unit_is_marked_truncated_with_save_command(tmp_path, capsys):
    p = tmp_path / "긴글.txt"
    p.write_text("가" * 9000 + "\n\n끝 단락", encoding="utf-8")
    code, out = run_main(reader, capsys, "read", "--file", str(p), "--units", "1")
    assert code == 0 and out["truncated"] is True and out["units"][0]["text_truncated"] is True
    assert out["units"][0]["text_chars"] == 9000 and len(out["units"][0]["text"]) == rt.DEFAULT_MAX_CHARS
    assert any("--save" in c["command"] and ".md" in c["command"] for c in out["next_commands"]) and "truncated_reason" in out


def test_partially_scanned_pdf_reports_pages_without_text(tmp_path, capsys):
    pymupdf = reader._import_pymupdf()           # 메시지는 그때의 stderr로 (1.4.90)
    p = tmp_path / "부분스캔.pdf"
    doc = pymupdf.open()
    doc.new_page().insert_text((72, 72), "first page text", fontname="helv", fontsize=11)
    doc.new_page().insert_image(pymupdf.Rect(40, 40, 560, 800), pixmap=_pixmap(pymupdf))   # 스캔 쪽(쪽을 덮는 그림만, 1.4.12)
    doc.new_page().insert_text((72, 72), "third page text", fontname="helv", fontsize=11)
    doc.save(str(p))
    doc.close()
    code, out = run_main(reader, capsys, "read", "--file", str(p))
    assert code == 0 and out["meta"]["pages_without_text"] == [2] and out["pages_without_text"] == [2]
    assert "그림만 있는 쪽" in out["agent_directive"]
    code, out = run_main(reader, capsys, "search", "--query", "warranty", "--file", str(p))
    assert code == 0 and out["total_matches"] == 0 and out["agent_directive"].startswith("UNVERIFIED:")
    assert any("그림만 있는 쪽" in g for g in out["not_searched"])


def test_sheet_row_limit_and_save_hit_limit_are_reported(tmp_path, capsys, monkeypatch):
    import openpyxl
    monkeypatch.setattr(reader, "MAX_SHEET_ROWS", 5)
    monkeypatch.setattr(reader, "MAX_SAVE_HITS", 2)
    p = tmp_path / "많은행.xlsx"
    wb = openpyxl.Workbook()
    for i in range(10):
        wb.active.append([f"항목 {i}", "공통"])
    wb.save(str(p))
    code, out = run_main(reader, capsys, "read", "--file", str(p))
    assert code == 0 and out["truncated"] is True and any("5행" in r for r in out["truncated_reason"])
    code, out = run_main(reader, capsys, "read", "--file", str(p), "--save", str(tmp_path / "o.csv"))
    assert code == 0 and out["truncated"] is True and out["agent_directive"].startswith("UNVERIFIED:")
    code, out = run_main(reader, capsys, "search", "--query", "공통", "--file", str(p), "--save", str(tmp_path / "hits.csv"))
    assert code == 0 and out["saved"] == 2 and out["save_truncated"] is True
    assert out["agent_directive"].startswith("UNVERIFIED:") and "앞 2건만 저장" in out["agent_directive"]


def test_reader_save_is_reopened(tmp_path, capsys, monkeypatch):
    p = tmp_path / "a.txt"
    p.write_text("첫 단락\n\n둘째 단락", encoding="utf-8")
    code, out = run_main(reader, capsys, "read", "--file", str(p), "--save", str(tmp_path / "o.csv"))
    assert code == 0 and out["agent_directive"].startswith("VERIFIED: 저장한 파일을 다시 열어")
    monkeypatch.setattr(rt, "check_saved_file", lambda path, rows=None, delimiter=None: "저장한 파일을 다시 열지 못함 (테스트)")
    code, out = run_main(reader, capsys, "read", "--file", str(p), "--save", str(tmp_path / "o2.md"))
    assert code == 0 and out["agent_directive"].startswith("UNVERIFIED:")
    code, out = run_main(reader, capsys, "search", "--query", "단락", "--file", str(p), "--save", str(tmp_path / "s.md"))
    assert code == 0 and out["agent_directive"].startswith("UNVERIFIED:") and out["save_problem"]


def test_index_reports_cache_write_failure(tmp_path, capsys, monkeypatch):
    p = tmp_path / "a.txt"
    p.write_text("내용", encoding="utf-8")
    real = rt.safe_write_text

    def failing(path, text, encoding="utf-8"):
        if path.endswith(".json") and "index" not in os.path.basename(path):
            raise OSError("디스크 가득 참")
        return real(path, text, encoding)
    monkeypatch.setattr(rt, "safe_write_text", failing)
    code, out = run_main(reader, capsys, "index", "--file", str(p))
    assert code == 0 and out["agent_directive"].startswith("UNVERIFIED:") and out["cache_failed"] == ["a.txt"]
    assert "캐시에 저장 확인" not in out["agent_directive"]


def test_normalize_doc_compat_filler_is_removed():
    assert not hasattr(reader, "normalize_doc") and reader.PARSER_VERSION >= 6


# ---------------------------------------------------------------------------
# 머리글·쪽 번호만 있는 쪽(빈 쪽·간지)은 '글자 없는 쪽'이 아님. 그림만 있는 쪽만, 조건 안의 쪽만
# ---------------------------------------------------------------------------


def test_header_only_page_is_not_flagged_and_search_stays_verified(tmp_path, capsys):
    p = tmp_path / "간지.pdf"
    _pdf(p, ["text", "header", "text", "text"])
    code, out = run_main(reader, capsys, "read", "--file", str(p))
    assert code == 0 and "pages_without_text" not in out["meta"] and "pages_without_text" not in out
    assert out["agent_directive"].startswith("VERIFIED:") and "그림만" not in out["agent_directive"]
    code, out = run_main(reader, capsys, "search", "--query", "warranty", "--file", str(p))
    assert code == 0 and out["total_matches"] == 0 and out["agent_directive"].startswith("VERIFIED: 0건"), out


def test_scanned_page_counts_only_inside_pages_filter(tmp_path, capsys):
    p = tmp_path / "스캔.pdf"
    _pdf(p, ["text", "image", "text"])
    code, out = run_main(reader, capsys, "search", "--query", "warranty", "--file", str(p), "--pages", "3")
    assert code == 0 and out["agent_directive"].startswith("VERIFIED: 0건"), out
    code, out = run_main(reader, capsys, "search", "--query", "warranty", "--file", str(p), "--pages", "2-3")
    assert code == 0 and out["agent_directive"].startswith("UNVERIFIED:") and any("p.2" in g for g in out["not_searched"])
    code, out = run_main(reader, capsys, "read", "--file", str(p), "--pages", "3")
    assert code == 0 and "pages_without_text" not in out and "notes" not in out
    code, out = run_main(reader, capsys, "read", "--file", str(p), "--pages", "1-2")
    assert out["pages_without_text"] == [2] and "notes" in out and "truncated_reason" not in out    # A6: 잘림이 아님


def test_section_filter_uses_section_of_preceding_text():
    doc = {"units": [{"page": 1, "section": ["1. 개요"]}, {"page": 3, "section": ["2. 시험"]}]}
    assert reader._pages_in_filter(doc, [2, 4], None, "") == [2, 4]
    assert reader._pages_in_filter(doc, [2, 4], {4}, "") == [4]
    assert reader._pages_in_filter(doc, [2, 4], None, "개요") == [2]
    assert reader._pages_in_filter(doc, [2, 4], None, "시험") == [4]


# ---------------------------------------------------------------------------
# 글자 없는 쪽 = 머리글·쪽 번호를 뺀 뒤 글이 없고, 그림이 쪽의 큰 부분을 덮는 쪽
# ---------------------------------------------------------------------------


def test_scanned_page_with_bates_footer_is_flagged(tmp_path, capsys):
    pymupdf = reader._import_pymupdf()           # 메시지는 그때의 stderr로 (1.4.90)
    doc = pymupdf.open()
    for i in range(1, 5):
        p = doc.new_page()
        p.insert_text((280, 820), f"- {i} -", fontsize=9)                       # 스캔 쪽에도 찍힌 쪽 번호
        if i == 3:
            p.insert_image(pymupdf.Rect(50, 60, 550, 780), pixmap=_pix_large(pymupdf))
        else:
            p.insert_text((72, 200), f"Body text page {i} torque {i * 10} Nm.", fontsize=11)
    path = tmp_path / "bates.pdf"
    doc.save(str(path))
    code, out = run_main(reader, capsys, "read", "--file", str(path))
    assert code == 0 and out["meta"]["pages_without_text"] == [3]


def test_blank_page_with_recurring_logo_is_not_flagged(tmp_path, capsys):
    pymupdf = reader._import_pymupdf()           # 메시지는 그때의 stderr로 (1.4.90)
    doc = pymupdf.open()
    for i in range(1, 5):
        p = doc.new_page()
        p.insert_image(pymupdf.Rect(500, 20, 560, 40), pixmap=_pix_large(pymupdf, 20, 10))   # 모든 쪽의 작은 로고
        if i != 3:
            p.insert_text((72, 200), f"Body text page {i} torque {i * 10} Nm.", fontsize=11)
    path = tmp_path / "logo.pdf"
    doc.save(str(path))
    code, out = run_main(reader, capsys, "read", "--file", str(path))
    assert code == 0 and "pages_without_text" not in out["meta"]
    code, out = run_main(reader, capsys, "search", "--query", "torque")
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out


def test_ocr_text_layer_over_image_is_not_flagged(tmp_path, capsys):
    pymupdf = reader._import_pymupdf()           # 메시지는 그때의 stderr로 (1.4.90)
    doc = pymupdf.open()
    for i in range(1, 3):
        p = doc.new_page()
        p.insert_image(p.rect, pixmap=_pix_large(pymupdf))
        p.insert_text((72, 200), f"OCR layer page {i}", fontsize=11, render_mode=3)     # 보이지 않는 OCR 글자 층
    path = tmp_path / "ocr.pdf"
    doc.save(str(path))
    code, out = run_main(reader, capsys, "read", "--file", str(path))
    assert code == 0 and "pages_without_text" not in out["meta"] and not out["meta"].get("ocr_needed")
    assert reader.PARSER_VERSION >= 7


# ---------------------------------------------------------------------------
# 짧은 글(도장, Bates 번호)만 남은 스캔 쪽도 표시, 모든 쪽이 스캔이면 OCR 필요로 UNVERIFIED
# ---------------------------------------------------------------------------


def test_scanned_page_with_stamp_text_is_flagged(tmp_path, capsys):
    pymupdf = reader._import_pymupdf()           # 메시지는 그때의 stderr로 (1.4.90)
    doc = pymupdf.open()
    for i in range(1, 5):
        p = doc.new_page()
        if i == 3:
            p.insert_image(pymupdf.Rect(30, 30, 565, 812), pixmap=_pix(pymupdf))
            p.insert_text((400, 400), "RECEIVED", fontsize=9)               # 스캔 위에 찍은 도장 글자
        else:
            p.insert_text((72, 200), f"Body text page {i} torque {i * 10} Nm.", fontsize=11)
    path = tmp_path / "stamp.pdf"
    doc.save(str(path))
    code, out = run_main(reader, capsys, "read", "--file", str(path))
    assert code == 0 and out["meta"]["pages_without_text"] == [3] and not out["meta"].get("ocr_needed"), out
    assert reader.PARSER_VERSION >= 8


def test_single_page_scan_with_bates_number_needs_ocr(tmp_path, capsys):
    pymupdf = reader._import_pymupdf()           # 메시지는 그때의 stderr로 (1.4.90)
    doc = pymupdf.open()
    p = doc.new_page()
    p.insert_image(p.rect, pixmap=_pix(pymupdf))
    p.insert_text((450, 820), "ABC0000001", fontsize=9)       # 한 쪽짜리라 머리글로 뺄 수 없는 Bates 번호
    path = tmp_path / "scan1.pdf"
    doc.save(str(path))
    code, out = run_main(reader, capsys, "read", "--file", str(path))
    assert code == 0 and out["meta"].get("ocr_needed") and out["meta"]["pages_without_text"] == [1], out
    assert out["agent_directive"].startswith("UNVERIFIED:") and "스캔" in out["agent_directive"]


def test_image_only_pdf_read_says_ocr_needed(tmp_path, capsys):
    pymupdf = reader._import_pymupdf()           # 메시지는 그때의 stderr로 (1.4.90)
    doc = pymupdf.open()
    for _ in range(2):
        p = doc.new_page()
        p.insert_image(p.rect, pixmap=_pix(pymupdf))
    path = tmp_path / "scan2.pdf"
    doc.save(str(path))
    code, out = run_main(reader, capsys, "read", "--file", str(path))
    assert code == 0 and out["count"] == 0 and out["meta"].get("ocr_needed"), out
    assert out["agent_directive"].startswith("UNVERIFIED:") and "OCR" in out["agent_directive"]
    assert "읽을 글이 없습니다" not in out["agent_directive"]


def test_ocr_layer_with_little_text_is_not_flagged(tmp_path, capsys):
    pymupdf = reader._import_pymupdf()           # 메시지는 그때의 stderr로 (1.4.90)
    doc = pymupdf.open()
    for i in range(1, 3):
        p = doc.new_page()
        p.insert_image(p.rect, pixmap=_pix(pymupdf))
        p.insert_text((72, 200), f"OCR {i}", fontsize=11, render_mode=3)      # 짧아도 보이지 않는 OCR 글자 층
    path = tmp_path / "ocr.pdf"
    doc.save(str(path))
    code, out = run_main(reader, capsys, "read", "--file", str(path))
    assert code == 0 and "pages_without_text" not in out["meta"] and not out["meta"].get("ocr_needed"), out


# ---------------------------------------------------------------------------
# 스캔 PDF 저장, 스캔 쪽만 고른 --pages는 UNVERIFIED (0개 저장이면 next_commands 없음)
# ---------------------------------------------------------------------------


def test_save_of_all_scanned_pdf_is_unverified_without_next_commands(tmp_path, capsys):
    path = _scan_pdf(tmp_path)
    out_md = tmp_path / "스캔.md"
    code, out = run_main(reader, capsys, "read", "--file", str(path), "--save", str(out_md))
    assert code == 0 and out["count"] == 0, out
    assert out["agent_directive"].startswith("UNVERIFIED:") and "OCR" in out["agent_directive"], out
    assert "next_commands" not in out


def test_pages_selecting_only_scanned_page_are_unverified(tmp_path, capsys):
    path = _scan_pdf(tmp_path, stamp_page=3, pages=4, scan_all=False)
    code, out = run_main(reader, capsys, "read", "--file", str(path), "--pages", "3")
    assert code == 0 and out["count"] == 1 and out["pages_without_text"] == [3], out
    assert out["agent_directive"].startswith("UNVERIFIED:") and "p.3" in out["agent_directive"], out
    code, out = run_main(reader, capsys, "read", "--file", str(path), "--pages", "3", "--save", str(tmp_path / "p3.md"))
    assert code == 0 and out["agent_directive"].startswith("UNVERIFIED:"), out
    assert "next_commands" not in out
    code, out = run_main(reader, capsys, "read", "--file", str(path), "--pages", "2-3")       # 본문 쪽이 섞이면 주의만
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out


# ---------------------------------------------------------------------------
# 손상된 한글 파일
# ---------------------------------------------------------------------------


def test_damaged_hwpx_is_input_error(tmp_path):
    p = tmp_path / "깨짐.hwpx"
    with zipfile.ZipFile(p, "w") as z:
        z.writestr("Contents/section0.xml", "<hs:sec><hp:p>닫히지 않음")
    with pytest.raises(rt.InputError) as ei:
        reader.load_document_status(str(p))
    assert "손상" in str(ei.value)


def test_damaged_hwpx_cli_not_internal(tmp_path, capsys):
    p = tmp_path / "깨짐.hwpx"
    with zipfile.ZipFile(p, "w") as z:
        z.writestr("Contents/section0.xml", "<bad")
    code, out = run_main(reader, capsys, "read", "--file", str(p))
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY") and out["error_type"] != "internal", out
    log = rt.error_log_path()
    assert not os.path.exists(log) or "traceback" not in open(log, encoding="utf-8").read()


@pytest.mark.parametrize("open_error", [None, OSError("incorrect OLE sector index")])
def test_damaged_hwp_is_input_error(tmp_path, monkeypatch, open_error):
    p = tmp_path / "깨짐.hwp"
    p.write_bytes(b"x")
    _fake_ole(monkeypatch, open_error)
    with pytest.raises(rt.InputError) as ei:
        reader.load_document_status(str(p))
    assert "손상" in str(ei.value)


# ---------------------------------------------------------------------------
# FileHeader 없는 OLE
# ---------------------------------------------------------------------------


def test_ole_without_fileheader_is_input_error(tmp_path, capsys):
    pytest.importorskip("olefile")
    p = tmp_path / "notreally.hwp"
    _ole_without_fileheader(str(p))
    with pytest.raises(rt.InputError) as ei:
        reader.parse_hwp(str(p))
    assert "HWP 5.0 문서가 아닙니다" in str(ei.value)
    code, out = run_main(reader, capsys, "read", "--file", str(p))
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY") and out["error_type"] != "internal", out


# ---------------------------------------------------------------------------
# 계산된 값이 없는 수식 칸
# ---------------------------------------------------------------------------


def test_check_failure_is_not_reused_from_cache(tmp_path, capsys, monkeypatch):
    """24차 LOW-5 (1.4.42): 수식 칸 확인 실패(MemoryError 등)를 캐시에서 그대로 다시 쓰지 않음."""
    c = formula_xlsx(tmp_path / "c.xlsx", cached=True)
    real = xf.xlsx_uncomputed_formulas
    monkeypatch.setattr(xf, "xlsx_uncomputed_formulas",
                        lambda *a, **k: {"count": 0, "cells": [], "sheets": {}, "unknown": True, "error": "MemoryError"})
    code, out = run_main(reader, capsys, "read", "--file", c)
    assert code == 0 and out["agent_directive"].startswith("UNVERIFIED:"), out
    monkeypatch.setattr(xf, "xlsx_uncomputed_formulas", real)          # 일시적 실패가 지나감 (파일은 그대로)
    code, out = run_main(reader, capsys, "read", "--file", c)
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    cached = reader.load_cached_only(os.path.abspath(c))
    assert "uncomputed_formulas" not in cached["meta"]                  # 캐시 파일도 고쳐 씀


def test_check_failure_rechecked_in_index_search(tmp_path, capsys, monkeypatch):
    c = formula_xlsx(tmp_path / "c.xlsx", cached=True)
    real = xf.xlsx_uncomputed_formulas
    monkeypatch.setattr(xf, "xlsx_uncomputed_formulas",
                        lambda *a, **k: {"count": 0, "cells": [], "sheets": {}, "unknown": True, "error": "MemoryError"})
    code, out = run_main(reader, capsys, "index", "--file", c)
    assert code == 0, out
    monkeypatch.setattr(xf, "xlsx_uncomputed_formulas", real)
    code, out = run_main(reader, capsys, "search", "--query", "없는글자ZQX")       # 색인(캐시)만 검색
    assert code == 0 and out["agent_directive"].startswith("VERIFIED: 0건."), out


def test_recheck_uncomputed_is_capped_per_run(tmp_path, capsys, monkeypatch):
    # 회귀: 확인 실패가 남은 큰 엑셀이 많으면 검색 한 번에 모두 다시 확인해 시간 제한을 넘음
    files = [formula_xlsx(tmp_path / f"c{i}.xlsx", cached=True) for i in range(reader.RECHECK_MAX + 2)]
    calls = []

    def failing(*a, **k):
        calls.append(1)
        return {"count": 0, "cells": [], "sheets": {}, "unknown": True, "error": "MemoryError"}
    monkeypatch.setattr(xf, "xlsx_uncomputed_formulas", failing)
    for f in files:
        run_main(reader, capsys, "index", "--file", f)
    calls.clear()
    code, out = run_main(reader, capsys, "search", "--query", "없는글자ZQX")
    assert code == 0 and len(calls) == reader.RECHECK_MAX, calls
    assert out["agent_directive"].startswith("UNVERIFIED:"), out          # 확인하지 못한 문서는 그대로 UNVERIFIED
    calls.clear()
    monkeypatch.setattr(rt, "time_limit", lambda: rt.TIME_RESERVE)       # 시간 여유가 없으면 다시 확인하지 않음
    run_main(reader, capsys, "search", "--query", "없는글자ZQX")
    assert calls == []


def test_reader_read_and_search(tmp_path, capsys):
    p = formula_xlsx(tmp_path / "f.xlsx")
    code, out = run_main(reader, capsys, "read", "--file", p)
    assert code == 0 and out["agent_directive"].startswith("UNVERIFIED:") and out["uncomputed_formulas"] == ["합계!B4"], out
    code, out = run_main(reader, capsys, "search", "--query", "없는글자ZQX", "--file", p)
    assert code == 0 and out["agent_directive"].startswith("UNVERIFIED:"), out
    assert any("계산된 값" in g for g in out["not_searched"])
    c = formula_xlsx(tmp_path / "c.xlsx", cached=True)
    code, out = run_main(reader, capsys, "search", "--query", "없는글자ZQX", "--file", c)
    assert code == 0 and out["agent_directive"].startswith("VERIFIED: 0건."), out


# ---------------------------------------------------------------------------
# 글상자 글
# ---------------------------------------------------------------------------


def test_parse_docx_reads_textboxes_once(tmp_path):
    units, meta = reader.parse_docx(_textbox_docx(tmp_path / "t.docx"))
    texts = [(u["loc"], u["text"]) for u in units]
    assert texts == [("단락 1", "본문 첫 단락"), ("글상자 1", "두 벌 저장 BOXTWO"), ("단락 2", "앞 글"),
                     ("글상자 2", "옛 글상자 BOXOLD"), ("단락 3", "본문 끝 단락")], texts
    assert meta["textboxes"] == 2 and meta["paragraphs"] == 3
    # 글상자가 없는 문서는 단위 번호·위치 그대로
    units, meta = reader.parse_docx(_textbox_docx(tmp_path / "n.docx", body=False))
    assert [u["loc"] for u in units] == ["단락 1", "단락 2"] and "textboxes" not in meta


def test_textbox_in_table_cell_not_doubled(tmp_path):
    import docx
    from docx.oxml import parse_xml
    d = new_docx()
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


def test_search_and_word_read_include_textbox(tmp_path, capsys):
    p = _textbox_docx(tmp_path / "t.docx")
    code, out = run_main(reader, capsys, "search", "--query", "BOXOLD", "--file", p)
    assert code == 0 and out["total_matches"] == 1 and out["matches"][0]["loc"] == "글상자 2", out
    code, out = run_main(reader, capsys, "search", "--query", "BOXTWO", "--file", p)
    assert out["total_matches"] == 1                       # 두 벌 저장도 한 번
    code, out = run_main(reader, capsys, "read", "--file", p)
    assert any("BOXTWO" in u["text"] for u in out["units"])
    code, out = run_main(word, capsys, "read", "--target", p)
    assert code == 0 and sum("BOXTWO" in u["text"] for u in out["paragraphs"]) == 1, out
    assert any("BOXOLD" in u["text"] for u in out["paragraphs"])


# ---------------------------------------------------------------------------
# 콘텐츠 컨트롤 안의 표 번호
# ---------------------------------------------------------------------------


def test_word_table_numbers_match_read(tmp_path):
    p = _sdt_docx(tmp_path / "s.docx")
    units, _ = reader.parse_docx(p)
    assert [u["loc"] for u in units if u["kind"] == "row"][0].startswith("표 1") and "SDT표머리" in units[1]["text"]
    for n, head in ((1, "SDT표머리"), (2, "일반표머리")):
        rows, name = office._word_rows(types.SimpleNamespace(what="table", table=n, from_target=p))
        assert rows[0] == [head] and name == "s.docx", rows
    with pytest.raises(rt.InputError) as ei:
        office._word_rows(types.SimpleNamespace(what="table", table=3, from_target=p))
    assert "전체 2개" in str(ei.value)


# ---------------------------------------------------------------------------
# CLI 계약
# ---------------------------------------------------------------------------


def test_cli_read_pages_and_summary(samples, env):
    code, out = run_reader_env(env, "read", "--file", samples["pptx"], "--pages", "2")
    assert code == 0 and out["status"] == "ok"
    assert out["units_total"] == 3 and [u["no"] for u in out["units"]] == [2]
    code, out = run_reader_env(env, "read", "--file", samples["pdf"], "--format", "summary")
    assert code == 0 and "units" not in out and out["meta"]["pages"] == 2


def test_cli_read_save_gives_write_commands(samples, env):
    code, out = run_reader_env(env, "read", "--file", samples["docx"], "--save")
    assert code == 0 and os.path.exists(out["output_file"]) and out["output_file"].endswith(".md")
    # 열린 문서에 넣기는 confirm_commands (1.4.75, 11차 검토: next_commands는 묻지 않고 실행해도 되는 것만)
    assert "word.py" in out["confirm_commands"][0]["command"] and "next_commands" not in out
    assert "confirm_commands" in out["agent_directive"] and "next_commands" not in out["agent_directive"]
    text = open(out["output_file"], encoding="utf-8").read()
    assert "## 표 1 2행 - 추진 배경" in text


def test_cli_read_truncation_offers_next_page(tmp_path, env):
    big = tmp_path / "긴문서.txt"
    big.write_text("\n\n".join(f"{i}번 단락 " + "가" * 500 for i in range(1, 60)), encoding="utf-8")
    code, out = run_reader_env(env, "read", "--file", str(big))
    assert out["truncated"] is True
    assert "--pages" in out["next_commands"][0]["command"]


def test_cli_search_index_folder_save_and_no_match(samples, env, tmp_path):
    for k in ("docx", "pptx", "pdf"):
        run_reader_env(env, "read", "--file", samples[k], "--format", "summary")
    code, out = run_reader_env(env, "search", "--query", "케이블카")
    assert code == 0 and out["files_searched"] == 3 and out["total_matches"] == 1
    assert out["matches"][0]["file"] == "발표.pptx" and out["matches"][0]["loc"] == "슬라이드 3"
    code, out = run_reader_env(env, "search", "--query", "cable|양산")
    assert out["total_matches"] == 2

    code, out = run_reader_env(env, "search", "--query", "마일스톤", "--save")
    assert out["matched"] == 1 and os.path.exists(out["output_file"])
    assert "word.py" in out["confirm_commands"][0]["command"]
    assert not any(" write " in c["command"] for c in out.get("next_commands", []))

    code, out = run_reader_env(env, "search", "--query", "양산", "--file", str(tmp_path), "--save", "결과.csv")
    assert out["total_matches"] >= 1 and out["output_file"].endswith("결과.csv")
    assert out["from_cache"] >= 3                       # read로 등록한 문서는 다시 읽지 않음
    assert "excel.py" in out["confirm_commands"][0]["command"]

    code, out = run_reader_env(env, "search", "--query", "존재하지않는단어")
    assert code == 0 and out["total_matches"] == 0 and out["agent_directive"].startswith("VERIFIED:")


def test_cli_search_without_index_is_stop_retry(env):
    code, out = run_reader_env(env, "search", "--query", "무엇")
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:")


def test_cli_list(samples, env):
    run_reader_env(env, "read", "--file", samples["hwpx"], "--format", "summary")
    code, out = run_reader_env(env, "list")
    assert code == 0 and out["total"] == 1 and out["documents"][0]["file"] == "양식.hwpx"


def test_cli_missing_file_suggests(samples, env, tmp_path):
    data = tmp_path / "data"
    data.mkdir(exist_ok=True)
    (data / "보고서_최종.pdf").write_bytes(b"x")
    code, out = run_reader_env(env, "read", "--file", "보고서_최종.docx")
    assert code == 1 and "[후보] 1) 보고서_최종.pdf" in out["message"] and "[추천]" not in out["message"]


def test_cli_argument_error(env):
    code, out = run_reader_env(env, "read")
    assert code == 2 and out["status"] == "error"


# ---------------------------------------------------------------------------
# xlsx 셀 글자 상한에서 잘린 저장은 UNVERIFIED, 긴 글 저장 안내는 .md
# ---------------------------------------------------------------------------


def test_xlsx_cell_cut_is_unverified(tmp_path, capsys, monkeypatch):
    monkeypatch.setattr(reader, "XLSX_CELL_MAX", 10)
    p = tmp_path / "긴글.txt"
    p.write_text("가" * 40 + "\n\n짧은 단락", encoding="utf-8")
    code, out = run_main(reader, capsys, "read", "--file", str(p), "--save", str(tmp_path / "o.xlsx"))
    assert code == 0 and out["agent_directive"].startswith("UNVERIFIED:") and "셀 글자 상한 10자 초과: 1행" in out["agent_directive"]
    assert out["truncated"] is True and isinstance(out["truncated_reason"], list)
    code, out = run_main(reader, capsys, "search", "--query", "가가", "--file", str(p), "--save", str(tmp_path / "s.xlsx"), "--full")
    assert code == 0 and out["agent_directive"].startswith("UNVERIFIED:") and "셀 글자 상한" in out["save_problem"]
    code, out = run_main(reader, capsys, "read", "--file", str(p), "--save", str(tmp_path / "o.md"))
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:")


def test_long_unit_next_command_saves_md(tmp_path, capsys):
    p = tmp_path / "긴글.txt"
    p.write_text("가" * 9000, encoding="utf-8")
    code, out = run_main(reader, capsys, "read", "--file", str(p))
    assert code == 0 and out["units"][0]["text_truncated"] is True
    cmds = [c["command"] for c in out["next_commands"]]
    assert len(cmds) == 1 and "--save" in cmds[0] and ".md" in cmds[0] and ".md" in out["truncated_reason"][0]


# ---------------------------------------------------------------------------
# truncated_reason은 늘 목록, 시트 행 상한만 잘렸으면 --save로 전체를 저장한다고 하지 않음
# ---------------------------------------------------------------------------


def test_row_cap_only_does_not_offer_full_save(tmp_path, capsys, monkeypatch):
    import openpyxl
    monkeypatch.setattr(reader, "MAX_SHEET_ROWS", 5)
    p = tmp_path / "많은행.xlsx"
    wb = openpyxl.Workbook()
    for i in range(10):
        wb.active.append([f"항목 {i}", "값"])
    wb.save(str(p))
    code, out = run_main(reader, capsys, "read", "--file", str(p))
    assert code == 0 and out["truncated"] is True and isinstance(out["truncated_reason"], list)
    assert "next_commands" not in out and "--save로 전체" not in out["agent_directive"]
    assert "5행" in out["agent_directive"] and "읽기 상한" in out["agent_directive"]
    code, out = run_main(reader, capsys, "read", "--file", str(p), "--save", str(tmp_path / "o.csv"))
    assert code == 0 and isinstance(out["truncated_reason"], list) and out["agent_directive"].startswith("UNVERIFIED:")


# ---------------------------------------------------------------------------
# parse_pages 큰 범위
# ---------------------------------------------------------------------------


def test_parse_pages_huge_range_is_fast():
    root = os.path.dirname(reader.__file__)
    code = ("import sys; sys.path.insert(0, %r); import reader; "
            "print(reader.parse_pages('1-999999999', 5), reader.parse_pages('999999990-999999999,2', 5))" % root)
    res = subprocess.run([sys.executable, "-c", code], capture_output=True, text=True, timeout=20)
    assert res.stdout.strip() == "[1, 2, 3, 4, 5] [2]", res.stderr
    with pytest.raises(rt.InputError):
        reader.parse_pages("999999990-999999999", 5)
    assert json.dumps(reader.parse_pages("0-2", 5)) == "[1, 2]"


# ---------------------------------------------------------------------------
# CLI: index / search / list
# ---------------------------------------------------------------------------


def test_index_folder_counts_and_unsupported_skipped(docs):
    code, out = run_reader("index", "--folder", str(docs))
    assert code == 0 and out["status"] == "ok"
    assert (out["count"], out["indexed"], out["updated"], out["unchanged"], out["failed"]) == (3, 3, 0, 0, 0)
    assert out["skipped_types"] == {".msg": 1, ".png": 1} and out["truncated"] is False
    code, out = run_reader("index", "--folder", str(docs))
    assert (out["indexed"], out["unchanged"]) == (0, 3)
    os.utime(docs / "spec.pdf", (1, 1))
    code, out = run_reader("index", "--folder", str(docs), "--depth", "0")
    assert (out["count"], out["updated"], out["unchanged"]) == (2, 1, 1)


def test_index_max_files_continues(docs):
    code, out = run_reader("index", "--folder", str(docs), "--max-files", "1")
    assert out["indexed"] == 1 and out["remaining"] == 2 and out["truncated"] is True
    assert "--max-files" in out["next_commands"][0]["command"] and "읽기 전용" in out["agent_directive"]
    code, out = run_reader("index", "--folder", str(docs), "--max-files", "1")
    assert out["indexed"] == 1 and out["unchanged"] == 1 and out["remaining"] == 1


def test_search_without_file_uses_cache_only(docs):
    code, out = run_reader("search", "--query", "내전압")
    assert code == 1 and out["error_type"] == "not_indexed" and out["agent_directive"].startswith("STOP_RETRY:")
    run_reader("index", "--folder", str(docs))
    code, out = run_reader("search", "--query", "내전압")
    assert code == 0 and out["total_matches"] == 2 and "indexed_now" not in out
    hit = [m for m in out["matches"] if m["file"] == "규격.docx"][0]
    assert hit["section"] == ["요구사항", "가. 검증기술"] and hit["section_text"] == "요구사항 > 가. 검증기술"
    assert hit["page"] == 2 and hit["page_source"] == "page_break" and hit["loc"] == "표 1 2행"
    assert "기준: 1.5 kV 1분" in hit["snippet"]
    assert hit["evidence"] == "규격.docx p.2(추정), 요구사항 > 가. 검증기술"
    pdf_hit = [m for m in out["matches"] if m["file"] == "계획서.pdf"][0]
    assert pdf_hit["evidence"] == "계획서.pdf p.1, III. 개발 내용 > 1. 개발 내용 종합 > 가. 검증기술"


def test_search_file_auto_caches_then_uses_cache(docs):
    code, out = run_reader("search", "--query", "Purpose", "--file", str(docs / "spec.pdf"))
    assert code == 0 and out["indexed_now"] == ["spec.pdf"] and out["from_cache"] == 0
    m = out["matches"][0]
    assert (m["page"], m["number"], m["section_text"]) == (2, "1.2", "1. SCOPE > 1.2. Purpose")
    assert m["evidence"] == "spec.pdf p.2, 1. SCOPE > 1.2. Purpose (1.2)"
    code, out = run_reader("search", "--query", "Purpose", "--file", str(docs / "spec.pdf"))
    assert out["indexed_now"] == [] and out["from_cache"] == 1
    code, out = run_reader("search", "--query", "내전압", "--file", str(docs))           # 폴더: 남은 2개만 새로 등록
    assert out["indexed_now_count"] == 2 and out["from_cache"] == 1 and out["skipped_types"][".msg"] == 1
    code, out = run_reader("list", "--name", "SPEC")
    assert out["count"] == 1 and out["documents"][0]["file"] == "spec.pdf"


def test_search_folder_respects_max_files(docs):
    code, out = run_reader("search", "--query", "본문", "--file", str(docs), "--max-files", "1")
    assert out["indexed_now_count"] == 1 and out["not_indexed_remaining"] == 2
    assert any("index" in c["command"] and "--folder" in c["command"] for c in out["next_commands"])


def test_search_pages_and_section_filters(docs):
    run_reader("index", "--folder", str(docs))
    code, out = run_reader("search", "--query", "본문", "--file", str(docs / "계획서.pdf"), "--pages", "2")
    assert out["total_matches"] == 2 and all(m["page"] == 2 for m in out["matches"])
    code, out = run_reader("search", "--query", "본문", "--section", "기대 효과")
    assert [m["section_text"] for m in out["matches"]] == ["III. 개발 내용 > 2. 기대 효과"]
    code, out = run_reader("search", "--query", "온도", "--pages", "1-3")
    assert out["total_matches"] >= 1 and all(m["page"] in (1, 2, 3) for m in out["matches"])
    code, out = run_reader("search", "--query", "본문", "--section", "없는 장절")
    assert code == 0 and out["count"] == 0 and out["agent_directive"].startswith("VERIFIED: 0건.")


def test_search_save_xlsx_columns_and_md_evidence(docs, tmp_path):
    import openpyxl
    run_reader("index", "--folder", str(docs))
    out_path = tmp_path / "근거.xlsx"
    code, out = run_reader("search", "--query", "내전압|Purpose", "--save", str(out_path), "--top", "1")
    assert code == 0 and out["output_file"] == str(out_path) and out["saved"] == out["total_matches"] == 4
    wb = openpyxl.load_workbook(out["output_file"])
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    assert list(rows[0]) == ["파일", "쪽", "장절", "번호", "위치", "내용(문장/단락)", "문맥", "점수"]
    assert ws["A1"].font.bold and ws["F2"].alignment.wrap_text
    docx_row = [r for r in rows if r[0] == "규격.docx"][0]
    assert docx_row[1] == 2 and docx_row[2] == "요구사항 > 가. 검증기술" and docx_row[4] == "표 1 2행"
    purpose = [r for r in rows if r[0] == "spec.pdf" and r[3] == "1.2"][0]
    assert purpose[1] == 2 and purpose[2] == "1. SCOPE > 1.2. Purpose"
    code, out = run_reader("search", "--query", "Purpose", "--save", str(tmp_path / "근거.md"))
    text = open(out["output_file"], encoding="utf-8").read()
    assert "근거: spec.pdf p.2, 1. SCOPE > 1.2. Purpose (1.2)" in text
    code, out = run_reader("search", "--query", "Purpose", "--save", str(tmp_path / "근거.csv"))
    assert out["output_file"].endswith(".csv") and "excel.py" in out["confirm_commands"][0]["command"]


def test_search_truncated_offers_full_save(docs):
    run_reader("index", "--folder", str(docs))
    code, out = run_reader("search", "--query", "본문|body", "--top", "1")
    assert out["truncated"] is True and out["shown"] == 1 and out["count"] > 1
    assert "--save" in out["next_commands"][-1]["command"]


def test_old_format_cache_is_skipped_and_reported_when_original_is_gone(tmp_path, isolated_dirs):
    """이전 파서 형식 캐시(키가 다름)는 다시 읽지 못하면 검색하지 않고 not_searched에 알림 (1.4.10: 옛 캐시 보정 코드 삭제)."""
    src = tmp_path / "old.txt"
    _old_cache(src)
    os.remove(src)
    code, out = run_reader("search", "--query", "검증")
    # 1.4.11: 검색하지 않은 문서는 stale(옛 내용으로 검색)에 세지 않고, 원본이 없으면 index --all이 아닌 clear --missing
    assert code == 0 and out["total_matches"] == 0 and "stale_cache_count" not in out
    assert out["skipped"][0]["reason"].startswith("이전 형식의 캐시") and "원본이 없어" in out["skipped"][0]["reason"]
    assert out["agent_directive"].startswith("UNVERIFIED:") and "clear --missing" in out["agent_directive"]
    assert any("clear --missing" in c["command"] for c in out["next_commands"])
    assert not any("index --all" in c["command"] for c in out["next_commands"])
    code, out = run_reader("list")
    assert out["documents"][0]["stale"] is True and out["documents"][0]["missing"] is True


def test_stale_cache_is_reread_during_search(tmp_path, isolated_dirs):
    """원본이 있는 낡은 캐시는 검색할 때 그 자리에서 다시 읽어 새 내용으로 찾습니다 (recovered로 알림)."""
    src = tmp_path / "old.txt"
    _old_cache(src, "보증 기간 3년. 새 내용의 검증 문장")
    code, out = run_reader("search", "--query", "검증")
    assert code == 0 and out["total_matches"] == 1 and "stale_cache_count" not in out
    assert out["matches"][0]["snippet"].startswith("보증 기간 3년") and "stale" not in out["matches"][0]
    assert out["recovered"] == ["원본이 바뀐 캐시 1개 -> 다시 읽고 검색"]
    assert out["agent_directive"].startswith("VERIFIED:")
    code, out = run_reader("list")
    assert out["documents"][0]["stale"] is False and out["stale_count"] == 0


def test_list_filters_missing_and_zero(docs):
    run_reader("index", "--folder", str(docs))
    code, out = run_reader("list", "--folder", str(docs / "sub"))
    assert out["count"] == 1 and out["documents"][0]["file"] == "규격.docx"
    d = out["documents"][0]
    assert d["pages"] == 2 and d["sections"] >= 3 and d["missing"] is False and d["cached_at"]
    os.remove(docs / "spec.pdf")
    code, out = run_reader("list", "--name", "spec")
    assert out["documents"][0]["missing"] is True and out["missing_count"] == 1
    code, out = run_reader("list", "--name", "없는이름")
    assert code == 0 and out["count"] == 0 and out["truncated"] is False and out["agent_directive"].startswith("VERIFIED: 0건.")


def test_read_shows_page_section_number_and_page_filter(docs):
    code, out = run_reader("read", "--file", str(docs / "계획서.pdf"), "--pages", "2")
    assert code == 0 and {u["page"] for u in out["units"]} == {2}
    u = [x for x in out["units"] if x["text"].startswith("2. 기대 효과")][0]
    assert u["number"] == "2" and u["section"] == ["III. 개발 내용", "2. 기대 효과"] and u["kind"] == "heading"
    code, out = run_reader("read", "--file", str(docs / "계획서.pdf"), "--save", str(docs / "out.md"))
    assert "근거: 계획서.pdf p.1, III. 개발 내용" in open(out["output_file"], encoding="utf-8").read()
    code, out = run_reader("list")
    assert out["documents"][0]["pages"] == 2 and "parser_version" not in out["documents"][0]      # 현재 형식이면 넣지 않음 (1.4.97)


def test_msg_is_unsupported_like_any_unknown_type(tmp_path):
    f = tmp_path / "메일.msg"
    f.write_bytes(b"\xd0\xcf\x11\xe0")
    for args in (("read", "--file", str(f)), ("search", "--query", "x", "--file", str(f)), ("index", "--file", str(f))):
        code, out = run_reader(*args)
        assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:")
        assert "지원하지 않는 형식입니다: .msg" in out["message"] and ".eml" in out["message"]


# ---------------------------------------------------------------------------
# 검색 문법, 검색 범위, 색인 복구
# ---------------------------------------------------------------------------


def test_parse_query_separators():
    """1.4.81: 저장소 구분자 규칙 (그리고 ;, 또는 |, ;가 먼저 묶임). 따옴표 안의 ; |는 글자."""
    assert reader.parse_query('"설계 기준";온도|습도') == [["설계 기준", "온도"], ["습도"]]
    assert reader.parse_query("A ; B | C D") == [["A", "B"], ["C D"]]
    assert reader.parse_query('"a;b|c";d') == [["a;b|c", "d"]]
    assert reader.parse_query("'a;b|c';d") == [["a;b|c", "d"]]                       # 1.4.83: excel --where와 같은 작은따옴표
    assert reader.parse_query("don't stop;x") == [["don't stop", "x"]]                   # 낱말 안의 '는 글자
    assert reader.parse_query('"견적') == [["견적"]] and reader.parse_query("'unclosed;x") == [["'unclosed", "x"]]   # 짝 없는 따옴표
    assert reader.parse_query("terms and conditions") == [["terms and conditions"]]    # 소문자 낱말은 글자 그대로
    assert reader.parse_query('""') == []


def test_old_and_or_query_gives_full_command(tmp_path, capsys):
    """옛 'A AND B', 'A OR B'는 받지 않고 ;/|로 고친 전체 명령 [추천] (따옴표 안의 AND는 글자)."""
    p = tmp_path / "a.txt"
    p.write_text("x", encoding="utf-8")
    code, out = run_main(reader, capsys, "search", "--query", '정격 전압 AND 절연 OR "AND 회로"', "--file", str(p))
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:") and "[추천]" in out["message"], out
    argv = recommended_argv(out["message"], "reader")
    assert argv[argv.index("--query") + 1] == '정격 전압;절연|"AND 회로"', argv


def test_matcher_ignores_spacing_and_empty_regex():
    m = reader.build_matcher("설계기준", False)
    assert m("본 설계 기준에 따른다") == ["설계기준"] and m("설계기\n준") == ["설계기준"]
    assert reader.build_matcher('"설계 기준"', False)("설계기준을 따른다")
    assert reader.build_matcher("a*", True)("xyz") is None
    assert "**설계 기준**" in reader.make_snippet("본 설계 기준에 따른다", ["설계기준"], 20)


def test_zero_hits_with_unindexed_files_is_unverified(tmp_path):
    d = tmp_path / "d"
    d.mkdir()
    for i in range(4):
        (d / f"f{i}.txt").write_text("하자 보수" if i == 3 else f"일반 문서 {i}", encoding="utf-8")
    code, out = run_reader("search", "--query", "하자", "--file", str(d), "--max-files", "2")
    assert code == 0 and out["count"] == 0 and out["not_indexed_remaining"] == 2
    assert out["agent_directive"].startswith("UNVERIFIED:") and "등록하지 못한 파일 2개" in out["agent_directive"]
    assert out["verified"] is False


def test_folder_search_reports_excluded_subfolders(tmp_path):
    d = tmp_path / "d"
    (d / "_보관").mkdir(parents=True)
    (d / "a.txt").write_text("일반", encoding="utf-8")
    (d / "_보관" / "b.txt").write_text("하자", encoding="utf-8")
    code, out = run_reader("search", "--query", "하자", "--file", str(d))
    assert code == 0 and out["count"] == 0 and any("하위 폴더 1개" in g for g in out["not_searched"])


def test_index_is_repaired_when_cache_exists_but_index_entry_is_missing(tmp_path):
    f = tmp_path / "a.txt"
    f.write_text("보증 기간 조항", encoding="utf-8")
    run_reader("index", "--file", str(f))
    os.remove(reader._index_path())                      # 캐시는 있는데 색인만 없음 (저장 전 중단)
    code, out = run_reader("index", "--file", str(f))
    assert code == 0 and out["unchanged"] == 1
    code, out = run_reader("search", "--query", "보증")
    assert code == 0 and out["count"] == 1


def test_index_keys_differing_only_in_case_are_merged(tmp_path):
    a, b = str(tmp_path / "A.txt"), str(tmp_path / "A.txt").lower()
    with open(reader._index_path(), "w", encoding="utf-8") as fp:
        json.dump({a: {"parsed_at": "1"}, b: {"parsed_at": "2"}}, fp)
    idx = reader.load_index()
    if os.path.normcase(a) == os.path.normcase(b):          # Windows: 같은 파일 -> 최근 것 하나
        assert list(idx) == [b]
    else:
        assert len(idx) == 2


def test_hwp_text_pairs_surrogates_and_replaces_lone_ones():
    rec = "\U00020BB7田".encode("utf-16-le") + b"\x00\xd8"        # 확장 한자 + 짝 없는 서로게이트
    text = reader._hwp_para_text(rec)
    assert text == "\U00020BB7田\ufffd"
    text.encode("utf-8")


def test_read_pages_on_docx_uses_real_pages(tmp_path):
    from docx.enum.text import WD_BREAK
    d = new_docx()
    d.add_paragraph("1쪽 문단")
    d.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
    d.add_paragraph("2쪽 문단")
    path = tmp_path / "p.docx"
    d.save(str(path))
    code, out = run_reader("read", "--file", str(path), "--pages", "2")
    assert code == 0 and out["pages_mode"] == "page" and [u["text"] for u in out["units"]] == ["2쪽 문단"]


def test_numbering_helper_creates_part_when_template_has_none(tmp_path):
    """회사 PC 기본 템플릿처럼 numbering.xml이 없는 문서에서도 번호 단락을 만들 수 있어야 함 (테스트 도우미 자체 확인)."""
    from docx.opc.constants import RELATIONSHIP_TYPE as RT
    d = new_docx()
    for rid, rel in list(d.part.rels.items()):
        if rel.reltype == RT.NUMBERING:
            del d.part.rels[rid]
    d.part.__dict__.pop("numbering_part", None)
    add_numbering(d)
    numbered(d, "일반 사항", 0)
    path = tmp_path / "n.docx"
    d.save(str(path))
    units, _ = reader.parse_docx(str(path))
    assert by_text(units, "3. 일반 사항")


@pytest.mark.parametrize("query,text,hit", [
    ("SOI 1", "Activities for SOI #1 are listed", True),        # 2026-10-29: 'SOI 1'로 'SOI #1'을 못 찾던 문제
    ("SOI #1", "SOI-1 검토", True),
    ("SOI1", "SOI #1", True),
    ("SOI 1", "SOI #10 활동", False),                            # 숫자 경계는 지킴
    ("SOI 1", "SOI 21", False),
    ("1.2", "제 1.2절", True),
    ("1.2", "11.2절", False),
    ("설계기준", "설계 기준에 따른다", True),
    ("#4", "SOI #4 활동", True),
    ("#4", "4개 항목", False),                                     # 기호를 빼면 숫자만 남는 검색어는 기호 그대로
    ("C++", "C++ 코딩 규칙", True),
])
def test_search_ignores_symbols_and_spacing_but_keeps_number_boundaries(query, text, hit):
    assert (reader.build_matcher(query, False)(text) is not None) == hit


def test_snippet_highlights_symbol_insensitive_match():
    assert "[[SOI #1]]" in reader.make_snippet("Activities for SOI #1 are listed", ["SOI 1"], 30)


def test_snippet_marks_layer_b_differently():
    # 10차 검토(1.4.73): 층 A 자리만 **, 숫자가 든 검색어의 층 B 자리(표기가 다른 곳)는 [[ ]], 숫자 없는 검색어는 **
    s = reader.make_snippet("허용 1·5 V 전원", ["5 V", "전원"], 20)
    assert "[[5 V]]" in s and "**전원**" in s and "**5 V**" not in s, s
    s = reader.make_snippet("보관 -40 ℃, 동작 40 ℃", ["40 ℃"], 30)
    assert "-[[40 ℃]]" in s and "동작 **40 ℃**" in s, s


# ---------------------------------------------------------------------------
# clear: 캐시 지우기
# ---------------------------------------------------------------------------


def test_clear_requires_target_and_dry_run_changes_nothing(tmp_path):
    a, b = _cached(tmp_path, "a규격.txt", "b보고.txt")
    code, out = run_reader("clear")
    assert code == 1 and "--dry-run" in out["message"]
    code, out = run_reader("clear", "--name", "규격", "--dry-run")
    assert code == 0 and out["dry_run"] is True and out["count"] == 1 and out["planned"] == ["캐시 삭제: a규격.txt"]
    assert "clear --name 규격" in out["confirm_commands"][0]["command"] and "--dry-run" not in out["confirm_commands"][0]["command"]
    assert len(reader.load_index()) == 2


def test_clear_by_name_removes_index_and_cache_file_keeps_original(tmp_path):
    a, b = _cached(tmp_path, "a규격.txt", "b보고.txt")
    cache_a = reader._doc_cache_path(str(a))
    code, out = run_reader("clear", "--name", "규격")
    assert code == 0 and out["verified"] is True and out["remaining_in_cache"] == 1
    assert not os.path.exists(cache_a) and a.exists()                     # 원본은 그대로
    assert list(reader.load_index()) == [str(b)]
    code, out = run_reader("search", "--query", "내용")
    assert out["count"] == 1                                             # 지운 문서는 검색 대상에서 빠짐


def test_clear_missing_and_all(tmp_path):
    a, b = _cached(tmp_path, "a.txt", "b.txt")
    os.remove(a)
    code, out = run_reader("clear", "--missing")
    assert code == 0 and [d["file"] for d in out["documents"]] == ["a.txt"]
    code, out = run_reader("clear", "--missing")
    assert code == 0 and out["count"] == 0 and "0건" in out["agent_directive"]
    code, out = run_reader("clear", "--all")
    assert code == 0 and out["remaining_in_cache"] == 0 and reader.load_index() == {}


# ---------------------------------------------------------------------------
# 이전 형식 캐시 - 상한을 넘은 것, stale 이중 집계, 원본 없는 것(clear --missing), directive 안내
# ---------------------------------------------------------------------------


def test_old_caches_over_refresh_limit_have_their_own_reason(tmp_path, capsys, monkeypatch):
    monkeypatch.setattr(reader, "STALE_REFRESH_MAX", 1)
    _old_caches(tmp_path, ["a.txt", "b.txt"])
    code, out = run_main(reader, capsys, "search", "--query", "검증")
    assert code == 0 and out["total_matches"] == 1                          # 1개는 다시 읽어 검색
    reasons = [s["reason"] for s in out["skipped"]]
    assert len(reasons) == 1 and "상한 1개" in reasons[0] and "다시 읽지 못함" not in reasons[0]
    assert "stale_cache_count" not in out                                   # 검색하지 않은 문서는 stale이 아님
    assert not any("옛 내용으로 검색" in g for g in out["not_searched"])
    assert out["agent_directive"].startswith("UNVERIFIED:") and "index --all" in out["agent_directive"]
    assert any(c["command"].endswith("index --all") for c in out["next_commands"])


def test_old_cache_without_original_points_to_clear_missing(tmp_path, capsys):
    _old_caches(tmp_path, ["gone.txt"], remove=("gone.txt",))
    code, out = run_main(reader, capsys, "search", "--query", "검증")
    assert code == 0 and "원본이 없어 다시 읽을 수 없음" in out["skipped"][0]["reason"]
    assert "stale_cache_count" not in out
    assert "clear --missing" in out["agent_directive"] and "index --all" not in out["agent_directive"]
    assert [c for c in out["next_commands"] if "clear --missing" in c["command"]]
    assert not [c for c in out["next_commands"] if "index --all" in c["command"]]


# ---------------------------------------------------------------------------
# 동시에 index를 실행해도 색인 항목이 사라지지 않음 (잠금 + 끝에 다시 읽어 채움), index --all이 빠진 캐시를 찾음
# ---------------------------------------------------------------------------


def test_concurrent_index_runs_keep_all_entries(tmp_path):
    folders = []
    for name in ("a", "b"):
        d = tmp_path / name
        d.mkdir()
        for i in range(30):
            (d / f"{name}{i}.txt").write_text(f"{name} 문서 {i} 내용", encoding="utf-8")
        folders.append(str(d))
    env = dict(os.environ, PYTHONIOENCODING="utf-8")
    procs = [subprocess.Popen([sys.executable, os.path.join(TOOL_DIR, "reader.py"), "index", "--folder", f],
                              stdout=subprocess.PIPE, stderr=subprocess.DEVNULL, env=env) for f in folders]
    outs = [json.loads(p.communicate(timeout=300)[0].decode("utf-8").strip().splitlines()[-1]) for p in procs]
    assert all(o["status"] == "ok" for o in outs), outs
    index = _index_json()
    assert len(index) == 60, len(index)
    assert not os.path.exists(os.path.join(os.environ["OFFICE_CLI_CACHE_DIR"], "index.lock"))


def test_index_repairs_entries_lost_to_another_run(tmp_path, capsys, monkeypatch):
    for i in range(3):
        (tmp_path / f"d{i}.txt").write_text(f"내용 {i}", encoding="utf-8")
    real = reader._index_files

    def racing(files, max_files, **kw):
        res = real(files, max_files, **kw)
        # 다른 실행이 옛 색인으로 덮어쓴 상황: 방금 넣은 항목 하나가 사라짐
        index = reader.load_index()
        index.pop(sorted(index)[0])
        reader._save_index(index)
        return res
    monkeypatch.setattr(reader, "_index_files", racing)
    code, out = run_main(reader, capsys, "index", "--folder", str(tmp_path))
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    assert any("다시 넣음" in r for r in out["recovered"])
    assert len(_index_json()) == 3


def test_index_all_readds_cache_files_missing_from_index(tmp_path, capsys):
    keep = tmp_path / "keep.txt"
    keep.write_text("남길 문서", encoding="utf-8")
    gone = tmp_path / "gone.txt"
    gone.write_text("지울 문서", encoding="utf-8")
    other = tmp_path / "other.txt"
    other.write_text("다른 문서", encoding="utf-8")
    for p in (keep, gone, other):
        run_main(reader, capsys, "read", "--file", str(p))
    index = _index_json()
    for p in (keep, gone):                                   # 색인에서만 빠진 캐시 (중단·동시 실행)
        index.pop(next(k for k in index if os.path.basename(k) == p.name))
    with open(os.path.join(os.environ["OFFICE_CLI_CACHE_DIR"], "index.json"), "w", encoding="utf-8") as f:
        json.dump(index, f)
    gone.unlink()                                            # 원본이 없는 캐시는 다시 넣지 않음
    code, out = run_main(reader, capsys, "index", "--all")
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    assert out["readded_to_index"] == 1 and any("1개 -> 색인에 다시 넣음" in r for r in out["recovered"])
    assert sorted(os.path.basename(k) for k in _index_json()) == ["keep.txt", "other.txt"]


def test_index_lock_waits_and_clears_stale_lock(monkeypatch):
    lock = os.path.join(os.environ["OFFICE_CLI_CACHE_DIR"], "index.lock")
    with open(lock, "w") as f:
        f.write("1")
    old = time.time() - reader.INDEX_LOCK_STALE - 5
    os.utime(lock, (old, old))                               # 중단된 실행이 남긴 잠금
    with reader._index_lock():
        assert os.path.exists(lock)
    assert not os.path.exists(lock)
    with open(lock, "w") as f:                               # 살아 있는 잠금: 상한까지 기다린 뒤 잠금 없이 진행 (무한 대기 없음)
        f.write("1")
    monkeypatch.setattr(reader, "INDEX_LOCK_WAIT", 0.2)
    start = time.monotonic()
    with reader._index_lock():
        pass
    assert time.monotonic() - start < 5 and os.path.exists(lock)


# ---------------------------------------------------------------------------
# 캐시도 원본도 없는 문서는 index --all 반복 안내가 아니라 clear --missing
# ---------------------------------------------------------------------------


def test_missing_cache_and_original_is_orphan(tmp_path, capsys):
    a = tmp_path / "a.txt"
    a.write_text("검색어 있음", encoding="utf-8")
    b = tmp_path / "b.txt"
    b.write_text("다른 글", encoding="utf-8")
    for p in (a, b):
        run_main(reader, capsys, "read", "--file", str(p))
    os.remove(reader._doc_cache_path(str(b)))
    b.unlink()
    code, out = run_main(reader, capsys, "search", "--query", "검색어")
    assert code == 0 and out["missing_file_count"] == 1
    assert "clear --missing" in out["skipped"][0]["reason"]
    cmds = [c["command"] for c in out["next_commands"]]
    assert any("clear --missing --dry-run" in c for c in cmds) and not any("index --all" in c for c in cmds)
    assert "clear --missing" in out["agent_directive"] and "index --all" not in out["agent_directive"]


# ---------------------------------------------------------------------------
# search --save --full은 단락 전체 저장, 아니면 잘린 행을 알림. JSON --full의 잘림 표시
# ---------------------------------------------------------------------------


def test_search_save_full_and_clamped_context(tmp_path, capsys):
    import openpyxl
    p = tmp_path / "long.txt"
    p.write_text("시작 검색어 " + "가" * 5000, encoding="utf-8")
    code, out = run_main(reader, capsys, "search", "--query", "검색어", "--file", str(p), "--save", str(tmp_path / "f.xlsx"), "--full")
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    ws = openpyxl.load_workbook(out["output_file"]).active
    assert len(ws.cell(2, 7).value) > 5000 and "context_clamped" not in out
    code, out = run_main(reader, capsys, "search", "--query", "검색어", "--save", str(tmp_path / "c.csv"), "--full")
    rows = rt.read_csv_rows(out["output_file"], ",")
    assert code == 0 and len(rows[1][6]) > 5000
    code, out = run_main(reader, capsys, "search", "--query", "검색어", "--save", str(tmp_path / "s.xlsx"))
    assert code == 0 and out["context_clamped"] == 1 and "3000자" in out["agent_directive"]
    code, out = run_main(reader, capsys, "search", "--query", "검색어", "--full")
    m = out["matches"][0]
    assert m["text_truncated"] and m["text_chars"] > 5000 and len(m["text"]) == 2000
    assert out["truncated"] is True and "--full" in out["next_commands"][-1]["command"]


# ---------------------------------------------------------------------------
# --top, --context 상한
# ---------------------------------------------------------------------------


def test_top_and_context_are_clamped(tmp_path, capsys):
    p = tmp_path / "many.txt"
    p.write_text("\n\n".join(f"항목 {i} 검색어 " + "나" * 2000 for i in range(80)), encoding="utf-8")
    code, out = run_main(reader, capsys, "search", "--query", "검색어", "--file", str(p), "--top", "500", "--context", "100000")
    assert code == 0 and out["shown"] == 50 and out["truncated"] is True
    assert len(out["warnings"]) == 2 and all(len(m["snippet"]) < 1200 for m in out["matches"])
    assert any("--save" in c["command"] for c in out["next_commands"])


# ---------------------------------------------------------------------------
# 색인을 잠깐 읽지 못해도 빈 색인으로 덮어쓰지 않음, 손상된 색인은 캐시 문서로 다시 만듦
# ---------------------------------------------------------------------------


def test_transient_index_lock_is_retried_not_emptied(tmp_path, capsys, monkeypatch):
    folder = _docs(tmp_path)
    code, out = run_main(reader, capsys, "index", "--folder", str(folder))
    assert code == 0 and len(_index_json()) == 3
    real = reader._read_bytes
    state = {"n": 0}

    def flaky(p):
        if p.endswith("index.json") and state["n"] == 0:
            state["n"] += 1
            raise PermissionError(13, "sharing violation", p)
        return real(p)
    monkeypatch.setattr(reader, "_read_bytes", flaky)
    monkeypatch.setattr(time, "sleep", lambda s: None)
    new = folder / "new.txt"
    new.write_text("새 문서", encoding="utf-8")
    reader.load_document_status(str(new))
    assert state["n"] == 1 and len(_index_json()) == 4


def test_unreadable_index_stops_without_overwriting(tmp_path, capsys, monkeypatch):
    folder = _docs(tmp_path)
    run_main(reader, capsys, "index", "--folder", str(folder))
    before = _index_json()
    real = reader._read_bytes

    def locked(p):
        if p.endswith("index.json"):
            raise PermissionError(13, "sharing violation", p)
        return real(p)
    monkeypatch.setattr(reader, "_read_bytes", locked)
    monkeypatch.setattr(time, "sleep", lambda s: None)
    new = folder / "new.txt"
    new.write_text("새 문서", encoding="utf-8")
    doc, _ = reader.load_document_status(str(new))           # 읽기는 계속, 색인은 그대로
    assert doc["units"] and _index_json() == before
    code, out = run_main(reader, capsys, "index", "--all")
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:") and "색인" in out["message"]
    assert _index_json() == before


def test_corrupt_index_is_rebuilt_from_cache(tmp_path, capsys):
    folder = _docs(tmp_path)
    run_main(reader, capsys, "index", "--folder", str(folder))
    with open(reader._index_path(), "w", encoding="utf-8") as f:
        f.write("{깨진 색인")
    code, out = run_main(reader, capsys, "search", "--query", "토크")
    assert code == 0 and out["total_matches"] == 3, out
    assert any("색인 파일 손상" in r for r in out.get("recovered", [])), out
    assert len(_index_json()) == 3


# ---------------------------------------------------------------------------
# search --full 저장 안내는 엑셀 칸 상한보다 긴 단락이 있으면 .md
# ---------------------------------------------------------------------------


def test_full_search_with_huge_unit_suggests_md(tmp_path, capsys):
    f = tmp_path / "huge.txt"
    f.write_text("앞 단락 토크\n\n" + "토크 " + "가" * 33000 + "\n", encoding="utf-8")
    code, out = run_main(reader, capsys, "search", "--query", "토크", "--file", str(f), "--full")
    assert code == 0 and out["truncated"], out
    saves = [c["command"] for c in out["next_commands"] if "--save" in c["command"]]
    assert saves and "근거.md" in saves[0] and "xlsx" not in saves[0], saves


# ---------------------------------------------------------------------------
# 오래된 잠금은 고유 이름으로 옮긴 뒤 지움 (다른 실행의 새 잠금은 지우지 않음)
# ---------------------------------------------------------------------------


def test_stale_lock_removal_keeps_fresh_lock(tmp_path):
    lock = tmp_path / "index.lock"
    lock.write_text("1")
    old = time.time() - reader.INDEX_LOCK_STALE - 10
    os.utime(lock, (old, old))
    assert reader._remove_stale_lock(str(lock)) is True and not lock.exists()
    lock.write_text("2")                                     # 그 사이 다른 실행이 새로 만든 잠금
    assert reader._remove_stale_lock(str(lock)) is False and lock.read_text() == "2"
    assert not [p.name for p in tmp_path.iterdir() if ".stale." in p.name]
    assert reader.INDEX_LOCK_WAIT > sum(0.5 * (i + 1) for i in range(6))   # 잠근 쪽의 retry_fs 재시도보다 길게


# ---------------------------------------------------------------------------
# 디코딩되지 않는 색인은 손상으로 보고 다시 만듦 (모든 reader 명령이 실패하던 문제)
# ---------------------------------------------------------------------------


def test_undecodable_index_is_rebuilt(tmp_path, capsys):
    folder = tmp_path / "docs"
    folder.mkdir()
    for i in range(2):
        (folder / f"d{i}.txt").write_text(f"문서 {i} 토크", encoding="utf-8")
    code, out = run_main(reader, capsys, "index", "--folder", str(folder))
    assert code == 0, out
    with open(reader._index_path(), "wb") as f:
        f.write(b'{"a": \xff\xfe\x80\x80')
    code, out = run_main(reader, capsys, "search", "--query", "토크")
    assert code == 0 and out["total_matches"] == 2, out
    assert any("색인 파일 손상" in r for r in out.get("recovered", [])), out
    with open(reader._index_path(), "wb") as f:
        f.write(b'{"a": \xff\xfe\x80\x80')
    code, out = run_main(reader, capsys, "clear", "--all")
    assert code == 0, out


# ---------------------------------------------------------------------------
# search 저장 안내는 next_commands를 가리킴 (directive에 근거.xlsx를 박지 않음)
# ---------------------------------------------------------------------------


def test_search_directive_points_to_next_commands(tmp_path, capsys):
    f = tmp_path / "huge.txt"
    f.write_text("앞 단락 토크\n\n" + "토크 " + "가" * 33000 + "\n", encoding="utf-8")
    code, out = run_main(reader, capsys, "search", "--query", "토크", "--file", str(f), "--full", "--top", "1")
    assert code == 0 and out["truncated"], out
    saves = [c["command"] for c in out["next_commands"] if "--save" in c["command"]]
    assert saves and "근거.md" in saves[0], saves
    assert "근거.xlsx" not in out["agent_directive"] and "next_commands" in out["agent_directive"], out


# ---------------------------------------------------------------------------
# clear --file/--folder의 ~
# ---------------------------------------------------------------------------


def test_clear_expands_home(tmp_path, capsys, monkeypatch):
    monkeypatch.setenv("HOME", str(tmp_path))
    monkeypatch.setenv("USERPROFILE", str(tmp_path))
    p = tmp_path / "docs" / "a.txt"
    p.parent.mkdir()
    p.write_text("내용", encoding="utf-8")
    run_main(reader, capsys, "read", "--file", str(p))
    code, out = run_main(reader, capsys, "clear", "--file", "~/docs/a.txt", "--dry-run")
    assert code == 0 and out["count"] == 1
    code, out = run_main(reader, capsys, "clear", "--folder", "~/docs")
    assert code == 0 and out["count"] == 1 and out["agent_directive"].startswith("VERIFIED:")


# ---------------------------------------------------------------------------
# reader clear 미리보기 전체 목록
# ---------------------------------------------------------------------------


def test_reader_clear_dry_run_truncated_has_full_result_file(tmp_path, capsys):
    docs = tmp_path / "docs"
    docs.mkdir()
    for i in range(reader.MAX_LIST_ITEMS + 3):
        (docs / f"문서{i:02d}.txt").write_text(f"내용 {i}", encoding="utf-8")
        code, out = run_main(reader, capsys, "index", "--file", str(docs / f"문서{i:02d}.txt"))
        assert code == 0, out
    code, out = run_main(reader, capsys, "clear", "--all", "--dry-run")
    assert code == 0 and out["truncated"] is True and out["count"] == reader.MAX_LIST_ITEMS + 3, out
    path = out["full_result_file"]
    assert os.path.isfile(path)
    rows = [r for r in rt.read_csv_rows(path) if any(r)]
    assert len(rows) == reader.MAX_LIST_ITEMS + 4 and rows[0] == ["file", "path", "missing"], rows[:2]
    assert len(reader.load_index()) == reader.MAX_LIST_ITEMS + 3          # 미리보기는 지우지 않음


# ---------------------------------------------------------------------------
# reader clear applied / done CSV
# ---------------------------------------------------------------------------


def test_reader_clear_applied_only_removed_and_done_csv(tmp_path, capsys, monkeypatch):
    docs = tmp_path / "docs"
    docs.mkdir()
    n = reader.MAX_LIST_ITEMS + 2
    for i in range(n):
        (docs / f"d{i:02d}.txt").write_text(f"내용 {i}", encoding="utf-8")
    code, out = run_main(reader, capsys, "index", "--folder", str(docs))
    assert code == 0, out
    stuck = {reader._doc_cache_path(str(docs / "d00.txt")), reader._doc_cache_path(str(docs / "d01.txt"))}
    real_remove = os.remove

    def rm(p, *a, **k):
        if p in stuck:
            raise OSError("cannot remove")
        return real_remove(p, *a, **k)
    monkeypatch.setattr(reader.os, "remove", rm)
    code, out = run_main(reader, capsys, "clear", "--all")
    assert code == 0 and out["agent_directive"].startswith("UNVERIFIED:"), out
    assert not any(a.endswith(("d00.txt", "d01.txt")) for a in out["applied"]), out["applied"]
    assert len(out["applied"]) == reader.MAX_LIST_ITEMS, out["applied"]
    rows = [r for r in rt.read_csv_rows(out["full_result_file"]) if any(r)]
    assert rows[0] == ["file", "path", "missing", "removed"], rows[0]
    flags = {r[0]: r[3] for r in rows[1:]}
    assert len(flags) == n and flags["d00.txt"] in (False, "False") and flags["d02.txt"] in (True, "True"), flags


# ---------------------------------------------------------------------------
# reader --file 파일 하나의 내부 오류는 내부 오류 그대로
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("argv", [["search", "--query", "hello", "--file"], ["index", "--file"]])
def test_single_file_internal_error_is_internal(one_file, capsys, monkeypatch, isolated_dirs, argv):
    def bug(path, check_index=True):
        raise KeyError("heading")
    monkeypatch.setattr(reader, "load_document_status", bug)
    code, out = run_fresh(reader, capsys, monkeypatch, *(argv + [str(one_file)]))
    assert code == 1 and out["error_type"] == "internal" and out["agent_directive"].startswith("ERROR:"), out
    assert out["message"] == "KeyError: 'heading'"
    assert "traceback:" in error_log_text(isolated_dirs)


@pytest.mark.parametrize("argv", [["search", "--query", "hello", "--file"], ["index", "--file"]])
def test_single_file_input_error_keeps_log_message(one_file, capsys, monkeypatch, isolated_dirs, argv):
    def bad(path, check_index=True):
        raise rt.InputError("문서를 읽을 수 없습니다: ZQX 본문", log_message="문서를 읽을 수 없습니다 (본문 생략)")
    monkeypatch.setattr(reader, "load_document_status", bad)
    code, out = run_fresh(reader, capsys, monkeypatch, *(argv + [str(one_file)]))
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:") and SECRET in out["message"], out
    log = error_log_text(isolated_dirs)
    assert SECRET not in log and "본문 생략" in log, log


def test_folder_search_failure_reason_names_error_type(one_file, capsys, monkeypatch):
    (one_file.parent / "b.txt").write_text("hello", encoding="utf-8")
    real = reader.load_document_status

    def bug(path, check_index=True):
        if path.endswith("a.txt"):
            raise KeyError("heading")
        return real(path, check_index=check_index)
    monkeypatch.setattr(reader, "load_document_status", bug)
    code, out = run_fresh(reader, capsys, monkeypatch, "search", "--query", "hello", "--file", str(one_file.parent))
    assert code == 0, out
    assert [s["reason"] for s in out["skipped"]] == ["KeyError: 'heading'"], out


# ---------------------------------------------------------------------------
# read --section: 장절 경로 일부가 맞는 단위만 읽기 (search --section과 같은 규칙, 1.4.49)
# ---------------------------------------------------------------------------


def test_read_section_reads_only_matching_units_with_evidence(tmp_path, capsys):
    p = tmp_path / "spec.pdf"
    english_pdf(p)
    code, out = run_main(reader, capsys, "read", "--file", str(p), "--section", "1.1.  SCOPE")   # 대소문자·공백 무시
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:"), out
    assert out["section_filter"] == "1.1.  SCOPE" and out["count"] == len(out["units"]) >= 2
    assert all("1.1. Scope" in u["section"] for u in out["units"])
    data = [u for u in out["units"] if u["text"].startswith("Data body")][0]
    assert data["page"] == 1 and data["evidence"] == "spec.pdf p.1, 1. SCOPE > 1.1. Scope > 1.1.1. Data"
    assert out["units_total"] > out["count"]
    # --pages와 함께: 두 조건을 모두 만족하는 단위만
    code, out = run_main(reader, capsys, "read", "--file", str(p), "--section", "scope", "--pages", "2")
    assert code == 0 and out["units"] and all(u["page"] == 2 and "1. SCOPE" in u["section"] for u in out["units"])
    assert not any(u["text"].startswith("Requirement") for u in out["units"])


def test_read_section_not_found_gives_candidates_not_logged(tmp_path, capsys, monkeypatch, isolated_dirs):
    from office_helpers import assert_not_logged
    p = tmp_path / "spec.pdf"
    make_text_pdf(p, [[("1. ZQX SCOPE", 16, True), ("Scope body.", 10, False), ("2. PURPOSE", 16, True),
                       ("Purpose body.", 10, False)]], font="helv")
    code, out = run_fresh(reader, capsys, monkeypatch, "read", "--file", str(p), "--section", "zqx scop")
    assert code == 0 and out["units"], out           # 부분 일치는 그대로 읽힘
    code, out = run_fresh(reader, capsys, monkeypatch, "read", "--file", str(p), "--section", "ZQX SCOPEE")
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:") and "[후보]" in out["message"]
    assert "1. ZQX SCOPE" in out["candidates"] and "[후보] 1) 1. ZQX SCOPE" in out["message"]
    assert split_command(out["confirm_commands"][0]["command"])[-2:] == ["--section", "1. ZQX SCOPE"], out   # 고른 번호의 전체 명령
    assert_not_logged(out, isolated_dirs, marker="1. ZQX SCOPE")   # 장절 이름(문서 글)은 오류 로그에 남지 않음
    # 비슷한 장절이 없으면 문서의 장절 일부를 [후보]로
    code, out = run_fresh(reader, capsys, monkeypatch, "read", "--file", str(p), "--section", "warranty")
    assert code == 1 and "[후보]" in out["message"] and set(out["candidates"]) == {"1. ZQX SCOPE", "2. PURPOSE"}


def test_read_section_continuation_and_save_keep_section(tmp_path, capsys, monkeypatch):
    p = tmp_path / "spec.pdf"
    english_pdf(p)
    monkeypatch.setattr(rt, "DEFAULT_MAX_CHARS", 60)
    code, out = run_main(reader, capsys, "read", "--file", str(p), "--section", "1. SCOPE")
    assert code == 0 and out["truncated"] is True and out["next_commands"]
    assert all("--section '1. SCOPE'" in c["command"] for c in out["next_commands"]), out["next_commands"]
    monkeypatch.setattr(rt, "DEFAULT_MAX_CHARS", 8000)
    save = tmp_path / "근거.csv"
    code, out = run_main(reader, capsys, "read", "--file", str(p), "--section", "purpose", "--save", str(save))
    assert code == 0 and out["agent_directive"].startswith("VERIFIED:")
    body = open(out["output_file"], encoding="utf-8-sig").read()
    assert "Purpose body" in body and "Data body" not in body and "Requirement" not in body


# ---------------------------------------------------------------------------
# 검색 재현율: 조사·숫자와 단위·전각·대시, 일부·유사 일치로 넓히기, 순위, 동의어 (1.4.51)
# ---------------------------------------------------------------------------

RECALL_TEXT = "\n\n".join([
    "정격 전압은 5.0 V 이하로 한다.",                   # 1
    "작동 온도 범위는 -20℃ ~ 60℃이다.",                # 2
    "ＰＣＢ 두께는 1.6 mm로 한다.",                     # 3 전각 영문
    "절연 저항은 100 MΩ 이상이어야 한다.",              # 4
    "접지\u2013단자는 외함에 연결한다.",                 # 5 en dash
    "rated voltage shall not exceed 220 V.",           # 6
    "케이블 피복의 난연성은 KS C 3341에 따른다.",       # 7
])

# (검색어, 찾아야 하는 단락 번호들, 기대 match_mode)
RECALL_CASES = [
    # 숫자가 든 검색어가 표기가 다른 곳(조사, 5V = 5.0 V, °C = ℃)에서만 맞으면 exact_mixed (층 B, 1.4.70)
    ("정격전압은 5V", {1}, "exact_mixed"),    # 검색어의 조사, 5V = 5.0 V
    ("작동 온도;-20°C", {2}, "exact_mixed"),  # °C = ℃ (부호 없는 '20°C'는 음수 '-20℃'와 맞지 않음, 1.4.68)
    ("PCB 두께", {3}, "exact"),               # 전각 ＰＣＢ (숫자 없는 검색어는 그대로)
    ("절연저항이 100MΩ", {4}, "exact_mixed"),  # 조사, 숫자와 단위 띄어쓰기
    ("접지-단자", {5}, "exact"),              # 대시 종류
    ("정격 전압;절연", {1, 4}, "partial"),  # 둘 다 든 단락 없음 -> 하나씩 든 단락
    ("케이불 피복", {7}, "fuzzy"),            # 오타 한 글자 (한글 4음절 이상, 첫·끝 음절 같음)
]


def _recall_file(tmp_path):
    p = tmp_path / "규격_재현율.txt"
    p.write_text(RECALL_TEXT, encoding="utf-8")
    return p


def _hit_numbers(out):
    paras = RECALL_TEXT.split("\n\n")
    got = set()
    for m in out.get("matches", []) + out.get("candidates", []):
        text = m.get("text") or m.get("snippet", "")
        for i, para in enumerate(paras, 1):
            if para[:6] in text.replace("**", ""):
                got.add(i)
    return got


@pytest.mark.parametrize("query,want,mode", RECALL_CASES)
def test_search_recall_variants(tmp_path, capsys, query, want, mode):
    p = _recall_file(tmp_path)
    code, out = run_main(reader, capsys, "search", "--query", query, "--file", str(p), "--full")
    assert code == 0 and _hit_numbers(out) == want, (query, out.get("matches"))
    assert out["match_mode"] == mode
    if mode == "exact_mixed":
        # 표기가 다른 곳만: 결과에 넣되 literal false + UNVERIFIED (글자 그대로 0건)
        assert out["count"] == len(want) and out["literal_count"] == 0 and all(m["literal"] is False for m in out["matches"])
        assert out["agent_directive"].startswith(f"UNVERIFIED: 글자 그대로 일치 0건, 표기가 다른 곳 {len(want)}건"), out
    elif mode != "exact":
        # 넓혀 찾은 결과는 정확 일치 0건 + candidates + UNVERIFIED (근거로 쓰지 않게)
        assert out["count"] == 0 and out["total_matches"] == 0 and out["candidate_count"] == len(want) and "matches" not in out
        assert out["agent_directive"].startswith("UNVERIFIED: 정확히 일치한 곳은 0건") and mode in out["agent_directive"]
        assert "recovered" not in out


def test_search_partial_lists_missing_terms_and_fuzzy_word(tmp_path, capsys):
    p = _recall_file(tmp_path)
    code, out = run_main(reader, capsys, "search", "--query", "정격 전압;절연", "--file", str(p))
    assert {tuple(m["missing_terms"]) for m in out["candidates"]} == {("절연",), ("정격 전압",)}
    assert {tuple(m["matched_terms"]) for m in out["candidates"]} == {("정격 전압",), ("절연",)}
    code, out = run_main(reader, capsys, "search", "--query", "케이불 피복", "--file", str(p))
    c = out["candidates"][0]
    assert c["fuzzy_words"] == {"케이불 피복": "케이블피복"} and c["matched_terms"] == [], c   # 실제로 맞은 낱말만
    # 짧은 검색어(한글 3음절 이하)는 유사 일치로 넓히지 않음: 0건 그대로
    code, out = run_main(reader, capsys, "search", "--query", "난연셩", "--file", str(p))
    assert out["count"] == 0 and out["match_mode"] == "exact" and "recovered" not in out


def test_search_widened_directive_keeps_notes(tmp_path, capsys):
    # 회귀: 넓혀 찾은 결과의 directive가 indexed_now·stale·clamp·text_cut 안내를 덮어씀
    p = _recall_file(tmp_path)
    code, out = run_main(reader, capsys, "search", "--query", "정격 전압;절연", "--file", str(p))
    assert out["match_mode"] != "exact" and out["agent_directive"].startswith("UNVERIFIED:"), out
    assert "indexed_now" in out["agent_directive"] and "다시 실행하지 마세요" in out["agent_directive"]


@pytest.mark.parametrize("query,word", [("전압", "전류"), ("전류", "전력"), ("최대", "최소"), ("상한", "하한"),
                                        ("대기 전류", "대기 전력"), ("최대 전류", "최소 전류"), ("정격 전압", "정격 전류"),
                                        ("상한 온도", "하한 온도")])
def test_fuzzy_never_matches_different_korean_terms(query, word):
    core = reader._fuzzy_core(query)
    assert core is None or reader._fuzzy_word(f"이 장치의 {word}는 5 이하", core) is None, (query, word)


def test_search_fuzzy_daegi_current_is_not_power(tmp_path, capsys):
    # 회귀: "대기 전류"가 "대기 전력"에 VERIFIED count 2로 맞음
    p = tmp_path / "규격.md"
    p.write_text("대기 전력은 0.5 W 이하로 한다.\n\n대기 전력 측정 방법은 부록 A.\n", encoding="utf-8")
    code, out = run_main(reader, capsys, "search", "--query", "대기 전류", "--file", str(p))
    assert code == 0 and out["count"] == 0 and out["match_mode"] == "exact" and out["agent_directive"].startswith("VERIFIED: 0건"), out


def test_search_particles_do_not_overstrip_short_words():
    m = reader.build_matcher("높이", False)
    assert m("높이 10 mm") and m("높은 온도") is None          # '높이'의 '이'는 조사가 아님 (남는 글이 1자)
    assert reader.build_matcher("5V", False)("15V 전원") is None
    assert reader.build_matcher("5V", False)("5.5 V 전원") is None
    assert reader.build_matcher("5V", False)("5[V] 전원")


def test_search_ranks_best_hits_first_and_saves_score(tmp_path, capsys):
    p = tmp_path / "순위.txt"
    p.write_text("\n\n".join(["전압 언급 하나.", "다른 문단.", "전압 시험: 전압을 올리며 전압 강하를 본다. 전압 기록.",
                              "전압 한 번."]), encoding="utf-8")
    code, out = run_main(reader, capsys, "search", "--query", "전압", "--file", str(p), "--top", "1",
                         "--save", str(tmp_path / "근거.csv"))
    assert out["matches"][0]["snippet"].startswith("**전압** 시험") or "전압 시험" in out["matches"][0]["snippet"]
    assert out["matches"][0]["score"] >= max(1, out["matches"][0]["score"])
    rows = open(out["output_file"], encoding="utf-8-sig").read().splitlines()
    assert rows[0].endswith(",점수") and "전압 시험" in rows[1]


def test_search_synonym_file(tmp_path, capsys, isolated_dirs):
    p = _recall_file(tmp_path)
    (isolated_dirs["data"] / "synonyms.txt").write_text("# 동의어\n정격 전압, rated voltage\n", encoding="utf-8")
    code, out = run_main(reader, capsys, "search", "--query", "rated voltage", "--file", str(p), "--full")
    assert code == 0 and _hit_numbers(out) == {1, 6} and out["match_mode"] == "exact"
    assert out["synonyms_used"] == {"rated voltage": ["정격 전압"]}


def test_search_bad_regex_recommends_literal_search(tmp_path, capsys, monkeypatch, isolated_dirs):
    p = _recall_file(tmp_path)
    code, out = run_fresh(reader, capsys, monkeypatch, "search", "--query", "전압 (5V", "--file", str(p), "--regex")
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:") and "정규식" in out["message"]
    rec = out["message"].split("[추천] ", 1)[1]
    assert rec.startswith("python ") and "--regex" not in rec and "search" in rec
    assert "전압 (5V" not in error_log_text(isolated_dirs)        # 검색어는 오류 로그에 남지 않음
    code, out = run_fresh(reader, capsys, monkeypatch, "search", "--query", r"\d+(V", "--file", str(p), "--regex")
    assert code == 1 and "[추천]" not in out["message"]          # 정규식을 쓰려던 것이 분명하면 글자 그대로 찾기를 권하지 않음


def test_search_case_sensitive_option_removed(capsys):
    """사용자 결정(1.4.52): search --case-sensitive/-c 없음. 대소문자는 늘 무시."""
    for flag in ("--case-sensitive", "-c"):
        code, out = run_main(reader, capsys, "search", "--query", "abc", flag)
        assert code == 2 and out["status"] == "error", out
    assert reader.build_matcher("ABC", False)("xx abc yy")


# ---------------------------------------------------------------------------
# R10 CONTINUE: 긴 문서 등록은 시간 제한 전에 멈추고 같은 명령으로 이어서 (1.4.56)
# ---------------------------------------------------------------------------


def _txt_folder(tmp_path, n=3):
    folder = tmp_path / "규격"
    folder.mkdir()
    for i in range(n):
        (folder / f"doc{i}.txt").write_text(f"정격 전압 {i} V 문서 {i}\n", encoding="utf-8")
    return folder


def _shell_args(cmd):
    return split_command(cmd)[2:]


@pytest.mark.parametrize("command", ["index", "search"])
def test_long_registration_continues_with_same_command(tmp_path, capsys, monkeypatch, command):
    folder = _txt_folder(tmp_path)
    monkeypatch.setattr(rt, "time_limit", lambda: rt.TIME_RESERVE)      # 첫 파일 뒤로는 늘 시간 초과
    argv = ["index", "--folder", str(folder)] if command == "index" else ["search", "--query", "정격 전압", "--file", str(folder)]
    seen = []
    for _ in range(5):
        code, out = run_main(reader, capsys, *argv)
        assert code == 0 and out["status"] == "ok", out
        seen.append(out["agent_directive"].split(":")[0])
        if not out.get("continue"):
            break
        assert out["verified"] is False and out["agent_directive"].startswith("CONTINUE:") and out["progress"], out
        nxt = out["next_commands"][0]["command"]
        assert nxt.startswith("python ") and _shell_args(nxt) == argv, nxt
    assert seen == ["CONTINUE", "CONTINUE", "VERIFIED"], seen          # 실행마다 새 파일 1개, 세 번째에 끝
    assert len(reader.load_index()) == 3
    if command == "search":
        assert out["count"] >= 3, out


@pytest.mark.parametrize("command", ["index", "search"])
def test_failing_file_is_remembered_so_continue_ends(tmp_path, capsys, monkeypatch, command):
    # 회귀: 실패하는 느린 파일이 앞에 있으면 매 실행 같은 파일을 다시 읽고 다음 파일 앞에서 CONTINUE (끝나지 않음)
    folder = _txt_folder(tmp_path)
    (folder / "a_bad.txt").write_text("깨진 파일", encoding="utf-8")
    orig = reader.load_document_status
    calls = []

    def fake(path, *a, **k):
        calls.append(os.path.basename(path))
        if "a_bad" in path:
            raise ValueError("손상된 파일")
        return orig(path, *a, **k)
    monkeypatch.setattr(reader, "load_document_status", fake)
    monkeypatch.setattr(rt, "time_limit", lambda: rt.TIME_RESERVE)      # 첫 파일 뒤로는 늘 시간 초과
    argv = ["index", "--folder", str(folder)] if command == "index" else ["search", "--query", "정격 전압", "--file", str(folder)]
    seen, progress = [], []
    for _ in range(6):
        code, out = run_main(reader, capsys, *argv)
        assert code == 0, out
        seen.append(out["agent_directive"].split(":")[0])
        if not out.get("continue"):
            break
        progress.append(out["progress"])
    assert seen[:3] == ["CONTINUE"] * 3 and seen[3] in ("VERIFIED", "UNVERIFIED"), seen   # 실행마다 새 파일 1개 (실패도 한 단계), 네 번째에 끝
    assert calls.count("a_bad.txt") == 1, calls                         # 실패한 파일은 한 번만 읽음
    assert "새로 등록 0개, 읽지 못함 1개" in progress[0], progress
    fails = out["failures"] if command == "index" else out["skipped"]
    assert [f["file"] for f in fails] == ["a_bad.txt"] and "이전 실행에서 읽지 못함" in fails[0]["reason"], out
    assert len(reader.load_index()) == 3
    # 파일이 바뀌면 다시 읽음, 하나만 주면 바로 다시 시도
    (folder / "a_bad.txt").write_text("고친 파일 정격 전압", encoding="utf-8")
    monkeypatch.setattr(reader, "load_document_status", orig)
    code, out = run_main(reader, capsys, "index", "--folder", str(folder))
    assert code == 0 and out["failed"] == 0 and out["indexed"] == 1 and not reader.load_failed(), out


def test_single_file_retries_known_failure(tmp_path, capsys, monkeypatch):
    f = tmp_path / "bad.txt"
    f.write_text("x", encoding="utf-8")
    reader.record_failure(str(f), "손상")
    code, out = run_main(reader, capsys, "index", "--file", str(f))
    assert code == 0 and out["indexed"] == 1 and not reader.load_failed(), out


def test_stale_refresh_failure_is_remembered(tmp_path, capsys, monkeypatch):
    # 회귀: 원본이 바뀐 캐시를 다시 읽다 실패하는 파일을 매 검색마다 다시 읽음
    folder = _txt_folder(tmp_path, 2)
    code, out = run_main(reader, capsys, "index", "--folder", str(folder))
    assert code == 0 and out["indexed"] == 2
    for i in range(2):
        p = folder / f"doc{i}.txt"
        p.write_text(f"정격 전압 바뀐 {i}\n", encoding="utf-8")
        os.utime(p, (1, 2_000_000_000 + i))
    orig = reader.load_document_status
    calls = []

    def fake(path, *a, **k):
        calls.append(os.path.basename(path))
        if "doc0" in path:
            raise ValueError("손상된 파일")
        return orig(path, *a, **k)
    monkeypatch.setattr(reader, "load_document_status", fake)
    for _ in range(2):
        code, out = run_main(reader, capsys, "search", "--query", "정격 전압")
        assert code == 0, out
    assert calls.count("doc0.txt") == 1, calls
    assert [f["file"] for f in out["refresh_failures"]] == ["doc0.txt"], out
    assert "이전 실행에서 읽지 못함" in out["refresh_failures"][0]["reason"]


def test_time_limit_env(monkeypatch):
    monkeypatch.setenv("OFFICE_CLI_TIME_LIMIT", "0")
    assert rt.time_limit() == float("inf")
    monkeypatch.setenv("OFFICE_CLI_TIME_LIMIT", "")
    assert rt.time_limit() == rt.DEFAULT_TIME_LIMIT == 100.0
    monkeypatch.setenv("OFFICE_CLI_TIME_LIMIT", "45")
    assert rt.time_limit() == 45.0


def test_continue_is_not_logged(tmp_path, capsys, monkeypatch, isolated_dirs):
    folder = _txt_folder(tmp_path, 2)
    monkeypatch.setattr(rt, "time_limit", lambda: rt.TIME_RESERVE)
    code, out = run_main(reader, capsys, "index", "--folder", str(folder))
    assert out.get("continue") is True
    assert not error_log_text(isolated_dirs)


# ---------------------------------------------------------------------------
# --section 장절 번호 경계 (1.4.64)
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("sec,key,want", [
    (["2. 전기", "2.1 정격"], "2.1", True),
    (["2. 전기", "2.1 정격", "2.1.3 전압"], "2.1", True),          # 하위 절 포함
    (["2. 전기", "2.10 기타"], "2.1", False),                      # 회귀: 2.10에도 맞았음
    (["12. 부록", "12.1 표"], "2.1", False),                       # 회귀: 12.1에도 맞았음
    (["1. 개요", "1.2.1 범위"], "2.1", False),
    (["2.1. 정격"], "2.1.", True),
    (["제2장 전기"], "2", True),
    (["12. 부록"], "2", False),
    (["A.1 부록 표"], "A.1", True),
    (["1. SCOPE", "1.1. Scope"], "Scope", True),                   # 제목 글자는 부분 일치
    (["3.2 정격 전압"], "정격", True),
])
def test_section_number_keeps_boundaries(sec, key, want):
    assert reader.section_matches({"section": sec}, reader.section_key(key)) is want


def test_search_section_number_excludes_2_10(tmp_path, capsys):
    f = tmp_path / "규격.md"
    f.write_text("# 2. 전기\n\n## 2.1 정격\n\n정격 전압 5 V\n\n## 2.10 기타\n\n정격 전압 12 V\n", encoding="utf-8")
    code, out = run_main(reader, capsys, "search", "--query", "정격 전압", "--file", str(f), "--section", "2.1")
    assert code == 0 and out["count"] == 1 and "5 V" in json.dumps(out["matches"][0], ensure_ascii=False), out


# ---------------------------------------------------------------------------
# 숫자만 든 검색어, 낱말 끝 조사 (1.4.64)
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("query,text,want", [
    ("5,000", "5000 rpm", True), ("5000", "5,000 rpm", True), ("0.5", "0.50 mm", True), ("0.50", "0.5 mm", True),
    ("0.5", "10.5 mm", False), ("0.5", "0.55 mm", False), ("1.2", "11.2", False), ("1.2", "3.1.2절", False),
    ("1.2", "12", False), ("5,000", "15,000", False),
    # 소수 한 자리 맨 숫자는 장절 번호 경계, 끝 0은 단위가 붙을 때만 (1.4.65)
    ("2.1", "2.10 전압", False), ("2.1", "2.1.3 세부", True), ("2.1", "12.1 전압", False), ("0.5", "0.50 이다", False),
    ("3.14", "3.140 이다", True), ("2.1 전압", "2.10 전압", False), ("1.5 mm", "1.50 mm", True), ("1.5", "1.50 mm", True),
    # 소수 끝 0 무시는 수학적으로 같은 값일 때만 (1.4.65 회귀: '2.0'이 '2.5 V', '1.0'이 '1.03 mm', '0.0'이 '0.5'에 맞음)
    ("2.0", "2.5 V", False), ("2.0", "2.05 V", False), ("2.0", "2.0 V", True), ("2.0", "2 V", True), ("2.0", "2.00", True),
    ("1.0", "1.03 mm", False), ("1.0", "1.0.3절", False), ("3.0", "3.5", False), ("0.0", "0.5", False), ("0.0", "0", True),
    ("2.0 V", "2.5 V", False), ("1.0 mm", "1.03 mm", False), ("2.10", "2.10 전압", True), ("2.10", "2.1 V", True),
    ("2.10", "2.1 전압", False), ("2.1", "2.15 V", False), ("1,500", "1,500.", True), ("0.5", "0.5,1.0", False),
    # 전각 단위 글자 (CP949): 끝 0이 단위 앞이면 같은 값
    ("1.5", "1.50㎜", True), ("2.5", "2.50㎝", True), ("0.5", "0.50㎏", True), ("1.5", "1.50㎖", True), ("0.5", "0.50㎃", True),
])
def test_number_only_query_normalized(query, text, want):
    assert (reader.build_matcher(query, False)(text) is not None) is want


# 4차 검토 (1.4.67): 숫자 검색의 오일치가 exact VERIFIED였음 (부호·단위 경계·기호 소수 경계·천 단위 쉼표)
NUMBER_SEARCH_CASES = [
    # 검색어의 부호는 패턴에 넣고 앞쪽 경계를 지킴. 글의 U+2212·U+2010~2013 빼기도 '-'와 같음
    ("-20 ℃", "보관 온도 20 ℃", False), ("-20 ℃", "+20 ℃", False), ("-20 ℃", "-20 ℃", True), ("-20 ℃", "\u221220 ℃", True),
    ("-20 ℃", "-20.5 ℃", False), ("-20 ℃", "1-20 ℃", False),
    ("-5 V", "5 V", False), ("-5 V", "+5 V", False), ("-5 V", "-5 V", True),
    ("-20", "-20", True), ("-20", "\u221220", True), ("-20", "\u201320", True), ("-20", "-20~60", True), ("-20", "20", False),
    ("-20", "1-20", False), ("-20", "A-20", False), ("-20", "-200", False), ("-20", "-20.5", False),
    ("20 ℃", "-20 ℃", False), ("20 ℃", "\u221220 ℃", False), ("20 ℃", "+20 ℃", False), ("20", "1-20", False),   # 1.4.69: 부호·대시가 다르면 후보
    ("±5 %", "5 %", False), ("±5 %", "-5 %", False), ("±5 %", "± 5 %", True),
    ("±5", "±5", True), ("±5", "±5.0", True), ("±5", "+/-5", True), ("±5", "±5.5", False), ("±5", "±50", False),
    ("±5", "±0.5", False),
    # 끝이 단위면 뒤에 영문·단위 글자가 이어지지 않아야 함 (Ω 등 비ASCII 포함)
    ("10 m", "10 m", True), ("10 m", "10m", True), ("10 m", "10 mm", False), ("10 m", "10 mA", False), ("10 m", "10 ms", False),
    ("10 m", "10 min", False), ("10 m", "10 MΩ", False),
    ("5 V", "5 VA", False), ("5 V", "5 V", True), ("5 V", "5 Vdc", True), ("5 V", "5 V/m", True), ("5V", "5 VA", False),
    ("5 A", "5 Ah", False), ("5 A", "5 A", True), ("100 W", "100 Wh", False), ("100 W", "100 W", True),
    ("2 s", "2 sec", True), ("2 s", "2 seconds", True), ("2 s", "2 s", True),         # 같은 단위 꼬리 (1.4.68)
    # 기호 분기의 소수 경계
    ("1/2", "1/2.5", False), ("1/2", "1/2", True), ("1/2", "1/2인치", True), ("1/2", "1/20", False),
    # 맞아야 하는 것: 10% = 10.0 %, 0.5 = 0.500, 날짜 표기
    ("10%", "10.0 %", True), ("10%", "10 %", True), ("10%", "110%", False), ("10%", "10.5%", False), ("10%", "100%", False),
    ("0.5", "0.500", True), ("0.5", "0.5", True), ("0.5", "0.55", False), ("12.5", "12.500", True),
    ("2026.11.04", "2026.11.4", True), ("2026.11.04", "2026-11-04", True), ("2026.11.04", "2026/11/04", True),
    ("2026.11.04", "2026. 11. 04", True), ("2026.11.04", "2026년 11월 4일", True), ("2026.11.4", "2026.11.04", True),
    ("2026.11.04", "2026.11.041", False), ("2026.11.04", "12026.11.04", False), ("2026.11.04", "2026.11.040", False),
    # 천 단위 쉼표는 3자리 묶음일 때만
    ("1,500", "1500", True), ("1,500", "1,500", True), ("1,500", "15,00", False), ("1,500", "2,1,500", False),
    ("1,500", "1,5000", False), ("1500", "1,500 rpm", True), ("1,500,000", "1500000", True), ("1,500,000", "1,500,000", True),
]


@pytest.mark.parametrize("query,text,want", NUMBER_SEARCH_CASES)
def test_number_search_boundaries(query, text, want):
    assert (reader.build_matcher(query, False)(text) is not None) is want


# 5차 검토 (1.4.68): 보고서 사례 전부. 정상 표기가 0건이 되던 것(±·부호 앞 글자, 같은 단위 꼬리)과 오일치
NUMBER_SEARCH_CASES_5 = [
    # [치명] ± 앞 글자 제한 없음 (+/-, +- 포함)
    ("±0.5 mm", "100±0.5 mm", True), ("±0.5 mm", "Φ10±0.5mm", True), ("±5%", "5V±5%", True), ("±10%", "AC220V±10%", True),
    ("±10%", "24VDC±10%", True), ("±0.1", "공차±0.1", True), ("±0.1", "R5±0.1", True), ("±1 dB", "3 dB±1 dB", True),
    ("±2%", "1000 ppm±2%", True), ("정격 전압 5V ±5%", "정격 전압: 5V±5%", True), ("±0.1 mm", "치수 100±0.1 mm", True),
    ("5 V ±5%", "5V±5%", True), ("220V ±10%", "AC 220V±10%", True), ("±5 %", "5 V±5 %", True), ("±10 %", "정격±10 %", True),
    ("±0.5", "100+/-0.5", True), ("±0.5", "100+-0.5", True), ("5V±5%", "5V±5%", True), ("5V±5%", "5V 5%", False),
    # 빼기 부호: 앞이 한글·=·: 이면 부호, 영문·숫자면 대시
    ("-20℃", "온도-20℃", True), ("-20 ℃", "Ta=-20℃", True), ("-20", "온도:-20℃", True), ("-20", "X-20", False),
    ("20", "온도-20℃", False), ("-5", "+/-5", False),
    # [높음] (a) + 또는 부호 없는 숫자는 ±에 맞지 않음
    ("+10%", "±10%", False), ("+5V", "±5V", False), ("+5 V", "±5 V 이내", False), ("5V", "±5V", False),
    ("10%", "+/-10%", False), ("10%", "+- 10%", False), ("+5V", "5V", False), ("+5V", "+5V", True),
    # [높음] (b) 숫자(+단위) 뒤 띄어쓴 대시는 범위 ('0 - 20 ℃'의 20은 -20이 아님)
    ("-20 ℃", "0 - 20 ℃", False), ("-20 ℃", "0 \u2013 20 ℃", False), ("-20 ℃", "0 -20 ℃", False), ("-20", "10 - 20", False),
    ("-5 V", "입력 0 -5 V", False), ("-3 dB", "1 - 3 dB", False), ("-20℃", "온도 0 ~ -20℃", True),
    ("-20 ℃", "5 V - 20 ℃", False), ("20", "10 -20", False), ("100 Hz", "10 \u2212100 Hz", False), ("20", "10~-20", False),
    # [중간] 같은 단위 꼬리는 허용, 다른 단위로 이어지면 아님
    ("5 V", "5 Vrms", False), ("5 V", "5 Vp-p", False), ("5 V", "5 Vpp", False), ("5 V", "5 Vin", True), ("5 V", "5 Vcc", True),
    ("5 V", "5 Vmax", True), ("5 V", "5 Vmin", True), ("5 V", "5 VAC", True), ("5V", "5 Vrms", False),
    ("3 s", "3 sec", True), ("3 s", "3 secs", True), ("10 min", "10 mins", True), ("10 min", "10 minutes", True),
    ("10 h", "10 hr", True), ("10 h", "10 hrs", True), ("30 dB", "30 dBA", False), ("1 bar", "1 barg", False),
    ("1 MPa", "1 MPaG", False), ("5 V", "5 VA", False), ("10 m", "10 mm", False), ("10 m", "10 MΩ", False),
    ("10 m", "10 min", False), ("1 A", "1 Ah", False), ("30 dB", "30 dBm", False), ("30 dB", "30 dBi", False),
    ("1 kg", "1 kgf", False), ("100 g", "100 gf", False), ("1 W", "1 Wh", False), ("2 s", "2 sq", False),
    # (여유) 표기 차이: 단위 글자, 곱하기, 지름, 시각, 영숫자 코드 경계
    ("10 kΩ", "10 kohm", True), ("10 kΩ", "10 k\u2126", True), ("10 Ω", "10 ohm", True), ("100mm", "100㎜", True),
    ("10 mA", "10 ㎃", True), ("100 ㎜", "100 mm", True), ("10 m", "10 ㎜", False),
    ("M6x1.0", "M6×1.0", True), ("M6×1.0", "M6x1.0", True), ("M6x1.0", "M6 1.0", False), ("Φ10", "Ø10", False),
    ("08:30", "8:30", True), ("8:30", "08:30", True), ("8:30", "18:30", False), ("08:30", "08 30", False),
    ("ABC-1234", "ABC-1234", True), ("ABC-1234", "(ABC-1234)", True), ("ABC-1234", "XABC-1234", False),
    ("ABC-1234", "ABC-1234-5", False), ("ABC-1234", "ABC-12345", False),
]


@pytest.mark.parametrize("query,text,want", NUMBER_SEARCH_CASES_5)
def test_number_search_boundaries_review5(query, text, want):
    assert (reader.build_matcher(query, False)(text) is not None) is want, (query, text)


def test_number_safety_net_candidates(tmp_path, capsys):
    # 5차 검토 S: 정확 일치 0건 + 검색어에 숫자 -> 숫자 값만 같은 곳을 candidates로 (0건 VERIFIED -> '규정 없음' 단정 방지)
    p = tmp_path / "규격.txt"
    p.write_text("정격 전압: DC 5V±5%\n\n외경: Φ10±0.5mm, 길이 100±0.5 mm\n\n보관 온도 -20℃\n", encoding="utf-8")
    code, out = run_main(reader, capsys, "search", "--query", "0.5 mm", "--file", str(p))
    assert code == 0 and out["count"] == 0 and out["match_mode"] == "number" and "matches" not in out, out
    assert out["agent_directive"].startswith("UNVERIFIED: 정확히 일치한 곳은 0건, 숫자가 같은 곳 1건 - 표기를 확인하세요")
    c = out["candidates"][0]
    assert c["number_words"] == {"0.5 mm": "±0.5mm"} and c["matched_terms"] == [] and "외경" in c["snippet"], c
    code, out = run_main(reader, capsys, "search", "--query", "20 ℃", "--file", str(p))   # 부호 없는 20은 -20과 다름
    assert out["match_mode"] == "number" and out["candidates"][0]["number_words"] == {"20 ℃": "-20℃"}, out
    code, out = run_main(reader, capsys, "search", "--query", "보관;20", "--file", str(p))  # 숫자 없는 검색어는 정확히
    assert out["match_mode"] == "number" and out["candidates"][0]["matched_terms"] == ["보관"], out
    # 숫자 경계는 지킴: 느슨한 검색도 0건이면 그때만 0건 VERIFIED
    code, out = run_main(reader, capsys, "search", "--query", "0.05 mm", "--file", str(p))
    assert out["count"] == 0 and out["match_mode"] == "exact" and out["agent_directive"].startswith("VERIFIED: 0건"), out
    code, out = run_main(reader, capsys, "search", "--query", "5V ±5%", "--file", str(p))  # 정확 일치는 그대로
    assert out["count"] == 1 and out["match_mode"] == "exact", out


# 6차 검토 (1.4.69): 보수적 원칙. 정확 일치(exact, VERIFIED)는 숫자 앞 부호가 분명히 같을 때만, 부호·대시가 다르거나
# 범위인지 부호인지 애매하면 숫자 안전망 후보(number, UNVERIFIED). 범위 대시는 양옆이 띄어진 '0 - 20'만.
# 'E' = exact, 'N' = 숫자 안전망 후보, '-' = 둘 다 아님
def _match_kind(query, text):
    if reader.build_matcher(query, False)(text) is not None:
        return "E"
    return "N" if reader.number_search([({}, {"text": text})], reader.parse_query(query), 2) else "-"


NUMBER_KIND_CASES_6 = [
    # 검토 보고 q1: 글의 숫자 앞에 부호가 있는데 검색어에는 없음 (또는 다른 부호) -> 후보
    ("40 ℃", "Ta 25 -40 ℃", "N"), ("40 ℃", "보관 85 ℃ -40 ℃", "N"), ("40 °C", "Range 85 °C -40 °C", "N"),
    ("40 ℃", "25 to -40 ℃", "N"), ("40 °C", "from 85 to -40 °C", "N"), ("12 V", "전원 +5 V -12 V", "N"),
    ("12 V", "Pin 3 VEE -12 V", "N"), ("5 V", "OUT2 -5 V", "N"), ("10 %", "허용차 +5 % -10 %", "N"),
    ("+10 %", "허용차 +5 % -10 %", "N"), ("15 %", "+10 % -15 %", "N"), ("+15 %", "+10 % -15 %", "N"),
    ("0.1 mm", "공차 +0.2 mm -0.1 mm", "N"), ("0.1", "공차 +0.2 -0.1", "N"), ("0.5", "+0 -0.5", "N"),
    ("+0.5", "+0 -0.5", "N"), ("3 dB", "Gain 1 dB -3 dB", "N"), ("40", "Typ 25 Min -40", "N"),
    ("0.02 mm", "+0.05 mm -0.02 mm", "N"), ("1 A", "Offset 2 A -1 A", "N"),
    # 6번: ± 빈틈(띄어쓰기 여러 칸, + / -), 쉼표·낱말에 붙은 빼기, 시간대
    ("0.5 mm", "±  0.5 mm", "N"), ("0.5 mm", "+ / - 0.5 mm", "N"), ("+0.5 mm", "+ / - 0.5 mm", "N"),
    ("10 %", "+/-  10 %", "N"), ("-5 V", "+ / - 5 V", "N"), ("40 ℃", "25,-40 ℃", "N"), ("40℃", "Min-40℃", "N"),
    ("08:30", "UTC-08:30", "N"), ("5:30", "GMT-5:30", "N"),
    # q2: 앞에 다른 값이 있는 음수
    ("30 ℃", "보관 70 ℃ -30 ℃", "N"), ("0.5 V", "Offset 1 mV -0.5 V", "N"), ("2 mm", "X 10 mm -2 mm", "N"),
    ("55 ℃", "Tmax 125 ℃ Tmin 25 ℃ -55 ℃", "N"), ("20 dBm", "Pout 10 dBm -20 dBm", "N"), ("1 %", "±2 % -1 %", "N"),
    ("3 ㎜", "치수 10 ㎜ -3 ㎜", "N"),
    # 반대로 맞아야 할 것: 같은 부호는 exact, 범위인지 애매하면 최소한 후보
    ("-40 ℃", "25 to -40 ℃", "E"), ("-0.1 mm", "공차 +0.2 mm -0.1 mm", "N"), ("-12 V", "전원 +5 V -12 V", "N"),
    ("±0.5 mm", "±  0.5 mm", "E"), ("±0.5 mm", "+ / - 0.5 mm", "E"), ("±10 %", "+/-  10 %", "E"),
    ("-40 ℃", "25,-40 ℃", "E"), ("-40℃", "Min-40℃", "E"), ("-20", "\u221220", "E"), ("-20", "\u201320", "E"),
    ("+5 V", "+ 5 V", "E"), ("-20", "10\u201320", "N"), ("3 dB", "-3 dB", "N"), ("0.5", "±0.5", "N"),
    # 범위 대시: 양옆이 띄어진 '0 - 20'만 범위 (부호 없는 20과는 exact, -20과는 아님)
    ("20", "0 - 20", "E"), ("20 ℃", "0 - 20 ℃", "E"), ("-20", "0 - 20", "N"), ("-20 ℃", "0 - 20 ℃", "N"),
    ("20", "10 -20", "N"), ("20", "1-20", "N"), ("100 Hz", "10-100 Hz", "N"),
    # 다른 낱말 뒤 숫자의 붙임표는 구분자 ('SOI 1' = 'SOI-1'), 영숫자 코드
    ("SOI 1", "SOI-1", "E"), ("SOI 1", "SOI #1", "E"), ("ABC-1234", "ABC-1234", "E"), ("ABC-1234", "ABC-1234-5", "N"),
    ("정격 전압 5 V", "정격 전압: 5V", "E"), ("정격 전압 5 V", "정격 전압: -5V", "N"),
    # 3번: 단위 접두 대소문자 (m/M, p/P, μ = u = U+00B5). 소문자 'mohm'은 밀리로 봄
    ("10 mohm", "10 MΩ", "N"), ("10 Mohm", "10 mΩ", "N"), ("10 MΩ", "10 mohm", "N"), ("100 mohm", "100 MΩ", "N"),
    ("1 Mohm", "1 mohm", "N"), ("1 mΩ", "1 ㏁", "N"), ("10 kohm", "10 MΩ", "N"), ("5 ohm", "5 ㏁", "N"),
    ("10 mHz", "10 MHz", "N"), ("1 mW", "1 MW", "N"), ("10 MPa", "10 mPa·s", "N"), ("5 mV", "5 MV", "N"),
    ("10 mA", "10 MA", "N"), ("5 v", "5 V", "E"), ("10 KΩ", "10 kΩ", "E"), ("10 uA", "10 \u03bcA", "E"),
    ("10 \u03bcA", "10 \u00b5A", "E"), ("10 min", "10 MIN", "E"), ("1 ppm", "1 PPM", "E"), ("10 mΩ", "10 mohm", "E"),
    # 7번: 뜻이 다른 꼬리 (피크·실효·게이지·가중), 지수, 상(φ)은 exact가 아님. 같은 단위 꼬리는 그대로
    ("1x10", "1×10\u207b\u00b3", "N"), ("1x10", "1×10^-3", "N"), ("1x10", "1×10", "E"),
    ("5 V", "5 Vp-p", "N"), ("5 V", "5 Vpp", "N"), ("5 V", "5 Vp", "N"), ("10 A", "10 Apk", "N"), ("5 V", "5 Vrms", "N"),
    ("1 bar", "1 barg", "N"), ("1 bar", "1 bara", "N"), ("30 dB", "30 dBA", "N"), ("Ø10", "φ10", "N"),
    ("Ø10", "ø10", "E"), ("Φ10", "φ10", "E"), ("5 V", "5 Vdc", "E"), ("5 V", "5 VAC", "E"), ("10 s", "10 sec", "E"),
    ("10 min", "10 mins", "E"), ("10 h", "10 hrs", "E"), ("5 V", "5 VA", "N"), ("10 m", "10 mm", "N"),
]


@pytest.mark.parametrize("query,text,want", NUMBER_KIND_CASES_6)
def test_number_match_kind_review6(query, text, want):
    assert _match_kind(query, text) == want, (query, text)


@pytest.mark.parametrize("query,text", [
    ("0 0 0 0 0 0 0 1", "REGA R/W " + "0 " * 3000), ("1 1 1 1 1 3", "1 " * 600), ("1 1 1 1 3", "1 " * 2500),
    ("1 1 1 1 1 1 1 3", "1 " * 20000), ("20", "1-20 " * 20000), ("5 V", "-5 V " * 20000),
], ids=["zeros", "ones600", "ones2500", "ones20000", "range20", "minus5V"])     # 긴 글이 id에 들면 Windows 환경변수 한도 초과
def test_number_search_linear_time(query, text):
    # 6차 검토 [치명]: 숫자 안전망의 큰 정규식이 '1 1 1 ...'에서 지수 시간 -> 숫자별 위치 목록 + 40자 창 선형 탐색
    t = time.monotonic()
    reader.build_matcher(query, False)(text)
    reader.number_search([({}, {"text": text})], reader.parse_query(query), re.IGNORECASE)
    assert time.monotonic() - t < 1.0


def test_number_safety_net_needs_other_words(tmp_path, capsys):
    # 6차 검토 [중간]: 숫자 아닌 낱말(단위 빼고)도 같은 단위 안에 있어야 후보. 낱말이 없으면 number를 건너뜀
    p = tmp_path / "규격.txt"
    p.write_text("정격 전류 1 A\n\n보관 온도 -20 ℃\n", encoding="utf-8")
    code, out = run_main(reader, capsys, "search", "--query", "1 퀀텀", "--file", str(p))
    assert code == 0 and out["count"] == 0 and out["match_mode"] == "exact", out
    assert out["agent_directive"].startswith("VERIFIED: 0건"), out
    code, out = run_main(reader, capsys, "search", "--query", "보관 20 ℃", "--file", str(p))   # 낱말이 있으면 후보
    assert out["match_mode"] == "number" and out["candidates"][0]["number_words"] == {"보관 20 ℃": "-20 ℃"}, out
    code, out = run_main(reader, capsys, "search", "--query", "정격 20 ℃", "--file", str(p))   # 같은 단위에 '정격' 없음
    assert out["match_mode"] != "number", out


def test_number_safety_net_time_cap_is_not_zero_verified(tmp_path, capsys, monkeypatch):
    # 6차 검토: 숫자 안전망이 시간 상한에 걸리면 0건이어도 VERIFIED로 단정하지 않음
    p = tmp_path / "규격.txt"
    p.write_text("\n\n".join(f"항목 {i} 값 {i}" for i in range(200)) + "\n", encoding="utf-8")
    monkeypatch.setattr(reader, "NUMBER_SEARCH_SECONDS", -1.0)
    code, out = run_main(reader, capsys, "search", "--query", "없는말 7777", "--file", str(p))
    assert code == 0 and out["count"] == 0 and out["agent_directive"].startswith("UNVERIFIED:"), out
    assert any("숫자 안전망" in g for g in out["not_searched"]), out


@pytest.mark.parametrize("query,text,want", [
    ("케이블로", "케이블 연결", True),          # ㄹ 받침 뒤 '로'는 조사
    ("케이블로", "케이블은 2 m", True),         # 뗀 낱말 뒤 조사는 허용
    ("케이블로", "케이블카 운행", False),       # 뗀 낱말 뒤에 한글이 이어지면 다른 낱말
    ("프로토콜로", "프로토콜", True),
    ("서버로", "서버로 전송", True),            # 받침 없는 글자 뒤 '로'는 떼지 않고 그대로 찾음
    ("보호회로", "보호회로 시험", True),        # 낱말 끝 '로'도 그대로 맞음
    ("보호회로 시험", "보호회로의 시험", True),
    # 1.4.65 회귀: 받침 없는 '로'를 떼고도 찾아 뗀 낱말 앞부분에 맞음
    ("전기회로", "전기회사 납품", False), ("냉각수로", "냉각수 온도 80 ℃", False), ("배수로", "배수 펌프", False),
    ("보호회로", "보호회의 개최", False), ("고속도로", "고속도 카메라", False), ("보호회로 시험", "보호회 시험", False),
    ("서버로", "서버 설정", False), ("우회로", "우회 경로", False),
])
def test_particle_ro(query, text, want):
    assert (reader.build_matcher(query, False)(text) is not None) is want
    assert reader.strip_particles("케이블로 보호회로") == "케이블 보호회로"


def test_word_final_syllable_is_not_particle():
    assert reader.strip_particles("와이파이 보호회로 역효과") == "와이파이 보호회로 역효과"
    assert reader.strip_particles("정격전압이 전류가 저항은 시험을 케이블의 전압으로") == "정격전압 전류 저항 시험 케이블 전압"
    assert reader.build_matcher("와이파이", False)("와이파 모듈") is None


# 7차 검토 (1.4.70): 숫자가 든 검색어의 두 층 판정. 층 A(literal: 글자 그대로 + 엄격한 경계)만 VERIFIED,
# 층 B(정규화 매처만 찾은 곳: 띄어쓰기·단위 별칭·부호 판정)는 결과에 넣되 literal false + UNVERIFIED (exact_mixed)
def _layer(query, text):
    g = reader.build_matcher(query, False)(text)
    if g is not None:
        return "A" if reader.literal_hit(text, g) else "B"
    return "N" if reader.number_search([({}, {"text": text})], reader.parse_query(query), re.IGNORECASE) else "-"


def _unescape(x):
    return re.sub(r"\\u([0-9a-f]{4})", lambda m: chr(int(m.group(1), 16)), x).replace("\\t", "\t").replace("\\n", "\n")


def _load_literal_cases():
    out = []
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "number_literal_cases.tsv")
    with open(path, encoding="utf-8") as f:
        for line in f:
            if line.startswith("#") or "\t" not in line:
                continue
            q, t, want = line.rstrip("\n").split("\t")
            out.append((_unescape(q), _unescape(t), want))
    return out


LITERAL_CASES = _load_literal_cases()


@pytest.mark.parametrize("query,text,want", LITERAL_CASES)
def test_number_literal_layer_review7(query, text, want):
    got = _layer(query, text)
    assert (got == "A") is (want == "A"), (query, text, got)


def test_number_literal_layer_counts():
    # 6·7차 검토 오일치 쌍 전부(notA)가 표에 있고, 기본 표기(A)는 모두 층 A
    assert sum(1 for c in LITERAL_CASES if c[2] == "notA") >= 250 and sum(1 for c in LITERAL_CASES if c[2] == "A") >= 20


@pytest.mark.parametrize("query,text,want", [
    # 층 B로 찾음 (표기가 다름): 결과에는 들지만 VERIFIED가 아님
    ("5 V", "5 Vdc", "B"), ("25 \u2103", "25\u00b0C", "B"), ("-40 \u2103", "\u221240\u2103", "B"),
    ("1,500 rpm", "1500 rpm", "B"), ("10 uA", "10 \u03bcA", "B"), ("SOI 1", "SOI-1", "B"),
    ("정격 전압 5 V", "정격 전압은 5V", "B"),
    # 모두 대문자 단위는 층 B에서 대소문자를 가리지 않음 ('10 MM', '10UF', '20MA'가 후보로만 나오던 회귀)
    ("10 MM", "10 mm", "B"), ("10 MM", "10 \u339c", "B"), ("10UF", "10 uF", "B"), ("10UF", "10 \u03bcF", "B"),
    ("20MA", "20 mA", "B"), ("20MA", "20 MA", "A"),
    # 소문자·섞인 단위의 접두 대소문자는 그대로 가림
    ("10 mA", "10 MA", "N"), ("10 M\u03a9", "10 m\u03a9", "N"), ("10 mW", "10 MW", "N"),
    # 7차 검토에서 exact VERIFIED였던 전각 빼기·탭 너머 빼기: 이제 층 B(UNVERIFIED)로만
    ("40 \u2103", "보관 \uff0d40 \u2103", "B"), ("-40 \u2103", "동작 온도\t-\t40 \u2103", "B"),
])
def test_number_layer_b_and_caps_units(query, text, want):
    assert _layer(query, text) == want, (query, text)


def test_search_exact_mixed_output(tmp_path, capsys):
    p = tmp_path / "규격.txt"
    p.write_text("정격 5 Vdc 공급\n\n입력 -5 V\n\n보조 전원 5V 공급\n\n출력 5 V 이하\n", encoding="utf-8")
    code, out = run_main(reader, capsys, "search", "--query", "5 V", "--file", str(p), "--save", str(tmp_path / "근거.csv"))
    assert code == 0 and out["match_mode"] == "exact_mixed" and out["count"] == 3 and out["literal_count"] == 2, out
    # 층 A를 먼저, 항목마다 literal. 부호가 다른 '-5 V'는 결과에 없음
    assert [m["literal"] for m in out["matches"]] == [True, True, False], out["matches"]
    assert "정격" in out["matches"][2]["snippet"] and all("-5" not in m["snippet"] for m in out["matches"])
    d = out["agent_directive"]
    assert d.startswith("UNVERIFIED: 글자 그대로 일치 2건, 표기가 다른 곳 1건(literal: false, 부호·단위를 확인하세요)"), d
    assert "literal: true인 결과만 근거로" in d and "evidence" in d and out["verified"] is False, d
    rows = open(out["output_file"], encoding="utf-8-sig").read().splitlines()
    assert rows[0].endswith(",점수,표기") and rows[-1].endswith(",표기 다름"), rows
    # 층 A만이면 그대로 exact VERIFIED, literal_count = count
    code, out = run_main(reader, capsys, "search", "--query", "출력 5 V", "--file", str(p))
    assert out["match_mode"] == "exact" and out["count"] == out["literal_count"] == 1, out
    assert out["matches"][0]["literal"] is True and out["agent_directive"].startswith("VERIFIED:"), out
    # 숫자 없는 검색어는 literal 키 없이 지금 그대로
    code, out = run_main(reader, capsys, "search", "--query", "공급", "--file", str(p))
    assert out["match_mode"] == "exact" and "literal" not in out["matches"][0] and out["literal_count"] == 2, out


def test_spaced_sign_is_layer_a_and_snippet_shows_it(tmp_path, capsys):
    # 8차 검토 (1.4.71) 설계 확정: 층 A는 '검색어 글자열이 그 자리에 글자 그대로'까지만. 띄어 쓴 부호('- 40 ℃')·범위의 뜻은
    # 판정하지 않고, 스니펫이 그 일치 자리와 앞(최소 15자, 부호 포함)을 반드시 보여 줌 + directive에 앞뒤를 함께 전달하라는 안내
    p = tmp_path / "규격.txt"
    p.write_text("동작 온도: - 40 \u2103 ~ + 85 \u2103\n", encoding="utf-8")
    assert _layer("40 \u2103", "동작 온도: - 40 \u2103 ~ + 85 \u2103") == "A"
    code, out = run_main(reader, capsys, "search", "--query", "40 \u2103", "--file", str(p), "--context", "3")
    assert code == 0 and out["count"] == out["literal_count"] == 1 and out["agent_directive"].startswith("VERIFIED:"), out
    snip = out["matches"][0]["snippet"]
    assert "- **40 \u2103**" in snip and snip.startswith("동작 온도: -"), snip
    assert reader.NUM_CONTEXT_NOTE in out["agent_directive"], out["agent_directive"]
    # 층 A로만 맞은 단위(정규화 매처는 못 찾음)도 스니펫·근거 문장에 값이 들어감 (전에는 글 앞부분만)
    text = "가" * 200 + " 출력 - 5 V"
    assert "- **5 V**" in reader.make_snippet(text, ["5 V"], 80)
    # 근거 문장: 층 A 자리가 든 문장을 먼저 ('5 Vdc'는 층 B)
    text = "정격 5 Vdc 공급. 출력 5 V 이하."
    assert reader.match_sentence(text, ["5 V"]) == "출력 5 V 이하."


def test_and_snippet_shows_far_number_term(tmp_path, capsys):
    # 9차 검토 [치명]: AND 검색 스니펫이 첫 검색어 자리만 창으로 잡아, 멀리 있는 숫자 검색어의 층 A 자리가 빠짐
    # -> 검색어마다 창, 겹치지 않으면 조각을 ' / '로 이음
    text = "보관 온도 조건은 아래와 같다. " + "가" * 150 + " 동작 범위 - 40 ℃에서 시험한다."
    p = tmp_path / "규격.txt"
    p.write_text(text + "\n", encoding="utf-8")
    code, out = run_main(reader, capsys, "search", "--query", "온도;40 ℃", "--file", str(p), "--context", "20")
    assert code == 0 and out["count"] == 1, out
    snip = out["matches"][0]["snippet"]
    assert "**온도**" in snip and "- **40 ℃**에서" in snip and reader.SNIPPET_JOIN in snip, snip
    assert snip.index("**온도**") < snip.index(reader.SNIPPET_JOIN) < snip.index("**40 ℃**"), snip
    # 가까우면 조각 하나 (지금 그대로)
    assert reader.SNIPPET_JOIN not in reader.make_snippet("온도 - 40 ℃에서", ["온도", "40 ℃"], 20)


def test_number_safety_net_checks_time_every_unit(monkeypatch):
    # 7차 검토 [보통]: 시간 상한을 50단위마다 확인해 큰 단위(1MB)가 이어지면 상한을 크게 넘음 -> 단위마다 확인
    monkeypatch.setattr(reader, "NUMBER_SEARCH_SECONDS", -1.0)
    st = {}
    units = [({}, {"text": "값 1 2 3"}) for _ in range(10)]
    reader.number_search(units, reader.parse_query("1 2 3 9"), re.IGNORECASE, st)
    assert st == {"cut": True, "checked": 0}, st


def test_pymupdf_messages_follow_the_current_stderr():
    """1.4.90: pymupdf는 메시지 출력 스트림 객체를 기억함. 그때의 sys.stderr를 주면 그것이 닫힌 뒤(같은 프로세스의 시험 캡처)
    메시지를 쓰다 'I/O operation on closed file' (집 PC 저장소 전체 pytest에서 diag 시험의 pymupdf import가 실패했음)."""
    import io
    real, tmp = sys.stderr, io.StringIO()
    sys.stderr = tmp
    try:
        pymupdf = reader._import_pymupdf()
    finally:
        sys.stderr = real
    tmp.close()
    pymupdf.message("closed stderr test")          # 닫힌 tmp가 아니라 지금의 stderr로 감


# ---------------------------------------------------------------------------
# 2026-11-09 검토 A4: 쪽 정보가 없는 문서의 search --pages는 read --pages처럼 항목 번호 (전에는 모두 빼고 VERIFIED 0건)
# ---------------------------------------------------------------------------


def test_search_pages_on_document_without_pages_uses_item_numbers(tmp_path, capsys):
    p = tmp_path / "a.txt"
    p.write_text("첫 줄 안내.\n\n정격 전압은 5 V 이다.\n\n끝.\n", encoding="utf-8")
    code, out = run_main(reader, capsys, "read", "--file", str(p), "--pages", "2")
    assert code == 0 and "정격 전압" in out["units"][0]["text"], out
    code, out = run_main(reader, capsys, "search", "--query", "정격 전압", "--pages", "1-3", "--file", str(p))
    assert code == 0 and out["count"] == 1 and out["agent_directive"].startswith("VERIFIED:"), out
    assert any("항목 번호" in w for w in out["warnings"]) and "page_unknown_skipped" not in out, out
    code, out = run_main(reader, capsys, "search", "--query", "정격 전압", "--pages", "3", "--file", str(p))
    assert code == 0 and out["count"] == 0, out


# ---------------------------------------------------------------------------
# 2026-11-09 검토 A5: 2열 표는 머리글 추정 없음, 머리글 행도 단위, 다음 쪽으로 이어진 PDF 표는 앞 표의 머리글
# ---------------------------------------------------------------------------


def test_two_column_table_has_no_guessed_header(tmp_path):
    import docx
    p = tmp_path / "t.docx"
    d = new_docx()
    t = d.add_table(rows=3, cols=2)
    for i, (k, v) in enumerate([("모델명", "PSU-100"), ("입력 전압", "220 V"), ("출력 전압", "5.0 V")]):
        t.cell(i, 0).text, t.cell(i, 1).text = k, v
    d.save(str(p))
    units, _ = reader.parse_docx(str(p))
    assert [u["text"] for u in units if u["kind"] == "row"] == ["모델명 | PSU-100", "입력 전압 | 220 V", "출력 전압 | 5.0 V"]


def test_pdf_table_continued_on_next_page_keeps_first_header(tmp_path, capsys):
    p = tmp_path / "c.pdf"
    head = ["Item", "Condition", "Limit"]
    make_text_pdf(str(p), font="helv", pages=[[("TABLE", [head, ["Efficiency", "230 Vac", "90 %"], ["Hold-up", "100 Vac", "10 ms"]])],
                           [("TABLE", [["Ripple", "20 MHz BW", "50 mVp-p"], ["Noise", "20 MHz BW", "100 mVp-p"]])]])
    units, _ = reader.parse_pdf(str(p))
    rows = [u["text"] for u in units if u["kind"] == "row"]
    assert "Item: Ripple | Condition: 20 MHz BW | Limit: 50 mVp-p" in rows, rows
    assert not any(t.startswith("Ripple: ") for t in rows), rows
    code, out = run_main(reader, capsys, "search", "--query", "Ripple;50 mVp-p", "--file", str(p))
    assert code == 0 and out["count"] == 1 and out["matches"][0]["snippet"].startswith("Item: **Ripple**"), out


# ---------------------------------------------------------------------------
# 2026-11-09 검토 A6: PDF 단락 안에서 줄바꿈된 줄('1.5 times ...', '5.0 VDC ...')은 장절 제목이 아님
# ---------------------------------------------------------------------------


def test_pdf_wrapped_lines_starting_with_numbers_stay_in_paragraph(tmp_path, capsys):
    pymupdf = reader._import_pymupdf()
    p = tmp_path / "wrap.pdf"
    doc = pymupdf.open()
    page = doc.new_page()

    def put(y, t, size=10, font="helv"):
        page.insert_text((72, y), t, fontname=font, fontsize=size)
    put(72, "3. ELECTRICAL", 14, "hebo")
    put(100, "3.2 Output", 12, "hebo")
    y = 124
    for line in ["The output current shall not exceed a limit of", "1.5 times the rated current for 10 s. The supply is",
                 "5.0 VDC nominal with decoupling of", "0.1 uF capacitors placed near each load. Screws are torqued to",
                 "1.5 N-m applied to each screw at assembly time."]:
        put(y, line)
        y += 12
    put(y + 20, "3.3 Protection", 12, "hebo")
    put(y + 44, "Over current protection is provided.")
    put(y + 64, "Body text for size statistics here and more words to dominate.")
    doc.save(str(p))
    doc.close()
    units, _ = reader.parse_pdf(str(p))
    heads = [u["text"] for u in units if u["kind"] == "heading"]
    assert heads == ["3. ELECTRICAL", "3.2 Output", "3.3 Protection"], heads
    prot = [u for u in units if u["text"].startswith("Over current")][0]
    assert prot["section"] == ["3. ELECTRICAL", "3.3 Protection"], prot
    code, out = run_main(reader, capsys, "search", "--query", "limit of 1.5 times", "--file", str(p), "--section", "3.2")
    assert code == 0 and out["count"] == 1, out


# ---------------------------------------------------------------------------
# 2026-11-09 검토 A13: 검색 범위가 빠졌으면 VERIFIED로 내지 않음, DOCX 머리글·각주, 쪽 경계 문장, diff의 스캔 쪽
# ---------------------------------------------------------------------------


def test_search_deleted_source_and_legacy_files_are_not_searched(tmp_path, capsys):
    folder = tmp_path / "d"
    folder.mkdir()
    (folder / "a.txt").write_text("정격 전압 5 V", encoding="utf-8")
    (folder / "b.txt").write_text("다른 글", encoding="utf-8")
    (folder / "old.doc").write_bytes(b"\xd0\xcf\x11\xe0" + b"\0" * 64)
    (folder / "그림.png").write_bytes(b"\x89PNG")
    code, out = run_main(reader, capsys, "search", "--query", "없는말", "--file", str(folder))
    assert code == 0 and out["agent_directive"].startswith("UNVERIFIED:"), out
    assert any(".doc 1개" in g for g in out["not_searched"]) and not any(".png" in g for g in out["not_searched"]), out
    code, out = run_main(reader, capsys, "index", "--folder", str(folder))
    assert code == 0
    os.remove(folder / "b.txt")
    code, out = run_main(reader, capsys, "search", "--query", "정격 전압")
    assert code == 0 and out["count"] == 1 and out["agent_directive"].startswith("UNVERIFIED:"), out
    assert any("원본이 지워졌거나" in g for g in out["not_searched"]), out


def test_docx_header_footer_and_footnotes_are_searchable(tmp_path, capsys):
    from sample_docs import new_docx
    p = tmp_path / "h.docx"
    d = new_docx()
    d.sections[0].header.paragraphs[0].text = "문서번호 QP-123 Rev. C"
    d.sections[0].footer.paragraphs[0].text = "사내 한정 배포"
    d.add_paragraph("본문 단락입니다.")
    d.save(str(p))
    units, _ = reader.parse_docx(str(p))
    extra = [(u["loc"], u["text"], u["section"]) for u in units if u.get("extra")]
    assert extra == [("머리글", "문서번호 QP-123 Rev. C", ["머리글·바닥글"]), ("바닥글", "사내 한정 배포", ["머리글·바닥글"])], extra
    code, out = run_main(reader, capsys, "search", "--query", "QP-123", "--file", str(p))
    assert code == 0 and out["count"] == 1 and out["agent_directive"].startswith("VERIFIED:"), out
    # 각주 (python-docx로는 만들 수 없어 XML을 넣음)
    import zipfile as zf
    q = tmp_path / "f.docx"
    with zf.ZipFile(p) as zin, zf.ZipFile(q, "w") as zout:
        for item in zin.infolist():
            zout.writestr(item, zin.read(item.filename))
        w = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
        zout.writestr("word/footnotes.xml", f'<w:footnotes {w}><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:t>-</w:t></w:r></w:p>'
                      f'</w:footnote><w:footnote w:id="1"><w:p><w:r><w:t>시험 온도는 23 </w:t></w:r><w:r><w:t>°C 기준</w:t></w:r></w:p>'
                      '</w:footnote></w:footnotes>')
    units, _ = reader.parse_docx(str(q))
    assert [(u["loc"], u["text"]) for u in units if u["loc"].startswith("각주")] == [("각주 1", "시험 온도는 23 °C 기준")]


def test_search_finds_sentence_split_across_pages(tmp_path, capsys):
    pymupdf = reader._import_pymupdf()
    p = tmp_path / "seam.pdf"
    doc = pymupdf.open()
    doc.new_page().insert_text((72, 700), "The relative humidity shall not exceed", fontname="helv", fontsize=10)
    doc.new_page().insert_text((72, 72), "95 % at 40 C during storage.", fontname="helv", fontsize=10)
    doc.save(str(p))
    doc.close()
    code, out = run_main(reader, capsys, "search", "--query", "humidity shall not exceed 95", "--file", str(p))
    assert code == 0 and out["count"] == 1 and out["matches"][0]["loc"] == "p.1~p.2", out


def test_index_search_template_needs_input(tmp_path, capsys):
    """2026-11-09 검토 B2: index 뒤의 search '<검색어>' 명령은 그대로 실행하면 0건이라 needs_input."""
    (tmp_path / "a.txt").write_text("본문", encoding="utf-8")
    code, out = run_main(reader, capsys, "index", "--folder", str(tmp_path))
    item = [c for c in out["next_commands"] if "<검색어>" in c["command"]][0]
    assert code == 0 and item["needs_input"] is True, out


def test_regex_with_nested_repeat_is_refused_quickly(tmp_path, capsys):
    """2026-11-09 검토 B10: '(\\w+\\s?)+$X'는 끝나지 않아 120초 제한에 JSON 없이 끊겼음. 구분 글자가 있는 반복은 그대로."""
    p = tmp_path / "a.txt"
    p.write_text("word " * 40 + "\n\n1,234,567 원", encoding="utf-8")
    code, out = run_main(reader, capsys, "search", "--query", r"(\w+\s?)+$X", "--regex", "--file", str(p))
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:") and "반복 안의 반복" in out["message"], out
    code, out = run_main(reader, capsys, "search", "--query", r"(?:\d{1,3},)+\d{3}", "--regex", "--file", str(p))
    assert code == 0 and out["count"] == 1, out


def test_pptx_shape_of_unknown_kind_does_not_fail_the_file(tmp_path):
    """2026-11-09 검토 B11: prstGeom도 txBox도 없는 p:sp는 shape_type이 NotImplementedError라 파일 전체가 ERROR였음."""
    prs = new_pptx()
    layout = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[-1]
    s = prs.slides.add_slide(layout)
    s.shapes.add_textbox(0, 0, 100, 100).text_frame.text = "보통 글"
    sp = s.shapes.add_shape(1, 0, 0, 100, 100)
    sp.text_frame.text = "이상한 도형 글"
    for g in list(sp._element.iter("{http://schemas.openxmlformats.org/drawingml/2006/main}prstGeom")):
        g.getparent().remove(g)
    p = tmp_path / "odd.pptx"
    prs.save(str(p))
    units, _ = reader.parse_pptx(str(p))
    assert "이상한 도형 글" in units[0]["text"], units


# ---------------------------------------------------------------------------
# 2026-11-09 검토 B13 (reader): --save 확장자, 이어 읽기 범위, 같은 이름 파일 근거, 스캔 PDF index
# ---------------------------------------------------------------------------


def test_b13_reader_small_fixes(tmp_path, capsys):
    a, b = tmp_path / "A", tmp_path / "B"
    a.mkdir()
    b.mkdir()
    (a / "spec.txt").write_text("전압 5 V", encoding="utf-8")
    (b / "spec.txt").write_text("전압 12 V", encoding="utf-8")
    code, out = run_main(reader, capsys, "search", "--query", "전압", "--file", str(tmp_path))
    ev = sorted(m["evidence"] for m in out["matches"])
    assert code == 0 and ev[0].startswith("A\\spec.txt") and ev[1].startswith("B\\spec.txt"), ev
    code, out = run_main(reader, capsys, "read", "--file", str(a / "spec.txt"), "--save", str(tmp_path / "x.docx"))
    assert code == 0 and out["output_file"].endswith("x.md") and out.get("recovered"), out
    p = tmp_path / "long.txt"
    p.write_text("\n\n".join("단락 %d " % i + "가" * 3000 for i in range(1, 15)), encoding="utf-8")
    code, out = run_main(reader, capsys, "read", "--file", str(p), "--pages", "2,3,4,5,6,13")
    nxt = [c["command"] for c in out.get("next_commands", []) if "이어서" in c["purpose"] or "--pages" in c["command"]]
    assert nxt and "13" in nxt[0] and "7" not in nxt[0].split("--pages")[1].split("--")[0], nxt
    scan = _scan_pdf(tmp_path)
    code, out = run_main(reader, capsys, "index", "--file", str(scan))
    assert code == 0 and any("스캔 PDF" in w for w in out.get("warnings", [])), out


# ---------------------------------------------------------------------------
# 1.5.0 EML 읽기·검색 (사용자 결정 2026-11-10: Outlook 가져오기 대신, 헤더·본문·첨부 이름만 캐시, 메일 한 통 = 결과 하나)
# ---------------------------------------------------------------------------


def _eml(path, subject, sender, when, body, html=False, attach=(), forward=None, raw_cp949_subject=False):
    from email.message import EmailMessage
    from email.utils import format_datetime
    m = EmailMessage()
    m["Subject"] = subject
    m["From"] = sender
    m["To"] = "나 <me@corp.com>"
    m["Date"] = format_datetime(when)
    if html:
        m.set_content(body, subtype="html")
    else:
        m.set_content(body)
    for name in attach:
        m.add_attachment(b"%PDF-1.4 x", maintype="application", subtype="pdf", filename=name)
    if forward is not None:
        m.add_attachment(forward)
    data = bytes(m)
    if raw_cp949_subject:            # 인코딩 표시 없는 CP949 제목 (사내 메일 서버에서 흔함)
        data = re.sub(rb"(?m)^Subject: .*$", b"Subject: " + subject.encode("cp949"), data, count=1)
    path.write_bytes(data)
    return path


@pytest.fixture
def mails(tmp_path):
    import datetime as _dt
    from email.message import EmailMessage
    from email.utils import format_datetime
    kst = _dt.timezone(_dt.timedelta(hours=9))
    folder = tmp_path / "메일"
    (folder / "2025").mkdir(parents=True)
    _eml(folder / "a.eml", "견적 요청 (모터 200대)", "김철수 <kim@vendor.com>", _dt.datetime(2026, 3, 2, 14, 5, tzinfo=kst),
         "안녕하세요.\n\n모터 200대 견적을 부탁드립니다.\n\n납기는 4월 말입니다.", attach=["견적서_v2.pdf"])
    _eml(folder / "b.eml", "회의록 공유", "이영희 <lee@corp.com>", _dt.datetime(2026, 3, 5, 9, 0, tzinfo=kst),
         "<html><body><p>회의록을 <b>공유</b>합니다.</p><p>다음 회의는 3월 12일.</p></body></html>", html=True)
    inner = EmailMessage()
    inner["Subject"] = "원래 메일: 시험 일정"
    inner["From"] = "박민수 <park@lab.com>"
    inner["Date"] = format_datetime(_dt.datetime(2026, 2, 1, 10, 0, tzinfo=kst))
    inner.set_content("시험은 2월 20일에 합니다.")
    _eml(folder / "2025" / "c.eml", "FW: 시험 일정", "최지훈 <choi@corp.com>", _dt.datetime(2025, 12, 1, 8, 0, tzinfo=kst),
         "아래 메일 참고하세요.", forward=inner)
    _eml(folder / "d.eml", "견적 회신", "정수진 <jung@vendor.com>", _dt.datetime(2026, 1, 10, 9, 0, tzinfo=kst), "견적을 보냅니다.",
         raw_cp949_subject=True)
    return folder


def test_eml_units_headers_body_attachments_and_forwarded(mails):
    units, meta = reader.parse_eml(str(mails / "2025" / "c.eml"))
    assert [u["loc"] for u in units] == ["제목", "헤더", "본문 1", "전달된 메일 제목", "전달된 메일 헤더", "전달된 메일 본문 1"]
    assert units[0]["section"] == ["FW: 시험 일정 (2025-12-01 08:00, 최지훈)"] and meta["mail"]["forwarded"] == 1
    assert units[5]["text"] == "시험은 2월 20일에 합니다." and units[5]["section"][0].startswith("전달된 메일 원래 메일: 시험 일정")
    units, meta = reader.parse_eml(str(mails / "a.eml"))
    assert units[-1]["text"] == "첨부: 견적서_v2.pdf" and meta["mail"]["attachments"] == ["견적서_v2.pdf"]
    assert "보낸 사람: 김철수 <kim@vendor.com>" in units[1]["text"] and "날짜: 2026-03-02 14:05" in units[1]["text"]
    units, _ = reader.parse_eml(str(mails / "b.eml"))
    assert [u["text"] for u in units if u["loc"].startswith("본문")] == ["회의록을 공유 합니다.", "다음 회의는 3월 12일."], units
    units, _ = reader.parse_eml(str(mails / "d.eml"))
    assert units[0]["text"] == "견적 회신"


def test_eml_folder_search_one_result_per_mail_newest_first(mails, capsys):
    code, out = run_main(reader, capsys, "index", "--folder", str(mails))
    assert code == 0 and out["count"] == 4 and out["indexed"] == 4, out
    code, out = run_main(reader, capsys, "search", "--query", "견적", "--file", str(mails))
    assert code == 0 and out["count"] == 2 and out["agent_directive"].startswith("VERIFIED:"), out
    assert [m["file"] for m in out["matches"]] == ["a.eml", "d.eml"]                 # 메일 한 통 = 결과 하나, 최신순
    first = out["matches"][0]
    assert first["mail_hits"] >= 2 and first["mail"]["subject"] == "견적 요청 (모터 200대)" and first["path"].endswith("a.eml"), first
    assert first["evidence"].startswith("a.eml 제목, 견적 요청 (모터 200대) (2026-03-02 14:05, 김철수)"), first
    code, out = run_main(reader, capsys, "search", "--query", "견적;김철수", "--file", str(mails))    # 제목과 보낸 사람에 나뉨
    assert out["count"] == 1 and out["matches"][0]["loc"] == "메일 전체", out
    code, out = run_main(reader, capsys, "search", "--query", "견적서_v2", "--file", str(mails))      # 첨부 이름
    assert out["count"] == 1 and out["matches"][0]["loc"] == "첨부", out
    code, out = run_main(reader, capsys, "search", "--query", "2월 20일", "--file", str(mails))      # 전달된 메일 본문 (하위 폴더)
    assert out["count"] == 1 and out["matches"][0]["loc"] == "전달된 메일 본문 1", out


def test_outline_lists_headings_with_units_and_pages(tmp_path, capsys, monkeypatch):
    """검토 4장 6번 (1.5.1): 큰 문서를 8,000자씩 이어 읽지 않고 장절 목록 -> 필요한 장절만 read --section/--units."""
    p = tmp_path / "k.pdf"
    korean_pdf(p)
    code, out = run_main(reader, capsys, "outline", "--file", str(p))
    assert code == 0 and out["count"] >= 5 and out["agent_directive"].startswith("VERIFIED:"), out
    first = out["headings"][0]
    assert first["text"] == "III. 개발 내용" and first["page"] == 1 and first["level"] == 1, first
    lv2 = [h for h in out["headings"] if h["text"].startswith("1. 개발 내용 종합")][0]
    assert lv2["level"] == 2, lv2
    code, part = run_main(reader, capsys, "read", "--file", str(p), "--units", lv2["units"])
    assert code == 0 and part["units"][0]["text"].startswith("1. 개발 내용 종합"), part
    monkeypatch.setattr(reader, "OUTLINE_MAX", 2)
    code, out = run_main(reader, capsys, "outline", "--file", str(p))
    assert out["truncated"] is True and out["shown"] == 2 and out["next_commands"][0]["command"].endswith("--start 3"), out
    t = tmp_path / "plain.txt"
    t.write_text("제목 없는 글입니다.\n\n둘째 단락입니다.", encoding="utf-8")
    code, out = run_main(reader, capsys, "outline", "--file", str(t))
    assert code == 0 and out["count"] == 0 and "0건" in out["agent_directive"], out


def test_save_without_path_goes_next_to_the_source_file(tmp_path, capsys, isolated_dirs):
    """1.6.7 전역 지침 8장: 파일을 대상으로 한 결과는 그 파일과 같은 폴더, 대상 파일이 없는 결과(캐시 전체 검색)는 C:\\AIKIT\\data
    (테스트에서는 OFFICE_CLI_RESULT_DIR). 이름만 준 --save도 대상 파일 옆."""
    folder = tmp_path / "규격 모음"
    folder.mkdir()
    src = folder / "규격서.txt"
    src.write_text("1. 범위\n시험 온도는 -20 ℃ 이상이어야 한다.\n", encoding="utf-8")
    code, out = run_main(reader, capsys, "read", "--file", str(src), "--save")
    assert code == 0 and os.path.dirname(out["output_file"]) == str(folder), out
    code, out = run_main(reader, capsys, "read", "--file", str(src), "--save", "요약.md")
    assert code == 0 and out["output_file"] == str(folder / "요약.md"), out
    code, out = run_main(reader, capsys, "search", "--query", "시험 온도", "--file", str(src), "--save")
    assert code == 0 and os.path.dirname(out["output_file"]) == str(folder), out
    code, out = run_main(reader, capsys, "search", "--query", "시험 온도", "--save")      # 캐시 전체: 대상 파일 없음
    assert code == 0 and os.path.dirname(out["output_file"]) == str(isolated_dirs["data"]), out
