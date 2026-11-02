# -*- coding: utf-8 -*-
"""
ml_report.py - ML 결과 보고서 (Word .docx) generator for one trained version folder.

One integrated document covers every target of the version (regression or classification):
cover, contents, summary, data, method, per-target results with charts, version history,
operation guide and appendix. All numbers and texts are generated deterministically from the
training artifacts; an LLM may optionally override narrative paragraphs with --slots.

Inputs (all written by ml_train.py 2.2.0+):
  <data_root>/history.jsonl, <version_dir>/leaderboard.json, model_<t>.json, oof_<t>.csv, correlation.csv

Select the version with -f <data file> (2.4.0: the best version of every target, built by
ml_train.make_best when the history changed; one report per task), -m <version_dir>, or with
-d <data_root> [-v vN] (without -v: the best version of the latest experiment, i.e. same task and
target set as the latest training run).
Requires python-docx and matplotlib (both in the standard environment).
"""

import sys
import os
import re

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from ml_predict import (  # noqa: E402  shared runtime lives in ml_predict.py
    setup_console, run_cli, JsonArgumentParser, InputError,
    sanitize_filename, load_json_file, save_with_lock_fallback, warn_long_path, avoid_overwrite,
    normalize_version, strip_unit, note_recovered, read_text_file, VERSION,
)

setup_console()

import io
import json
import argparse
from datetime import datetime
from typing import Dict, Any, List, Optional, Tuple

import numpy as np
import pandas as pd

# Optional at import time so a missing package is reported as protocol JSON
try:
    import docx
    from docx.shared import Pt, Cm, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_LINE_SPACING
    from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
    from docx.enum.section import WD_ORIENT
    from docx.enum.style import WD_STYLE_TYPE
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    _DOCX_IMPORT_ERROR = None
except ImportError as _e:
    _DOCX_IMPORT_ERROR = _e

try:
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib import font_manager
    _MPL_IMPORT_ERROR = None
except ImportError as _e:
    _MPL_IMPORT_ERROR = _e


REPORT_TITLE = "ML 결과 보고서"
REPORT_VERSION = VERSION

# ==============================================================================
# 1. Design tokens (grayscale-safe slate palette + one navy accent)
# ==============================================================================

FONT_KO = "맑은 고딕"
C_INK = "1F2937"      # body text
C_MUTED = "6B7280"    # captions, secondary text
C_LINE = "CBD5E1"     # table grid
C_HEAD = "F1F5F9"     # table header / box background
C_ACCENT = "1E3A8A"   # navy: titles, champion, primary series
C_ACCENT_BG = "E0E7FF"

VERDICT_STYLE = {      # label -> (text color, background)
    "적합": ("1E3A8A", "DBEAFE"),
    "조건부": ("92400E", "FEF3C7"),
    "부적합": ("991B1B", "FEE2E2"),
}
PASS_SCORE = 0.90
CONDITIONAL_SCORE = 0.70
OVERFIT_LIMIT = 0.15

MPL_NAVY = "#1E3A8A"
MPL_GRAY = "#9CA3AF"
MPL_LIGHT = "#E5E7EB"
MPL_ORANGE = "#D97706"
MPL_INK = "#1F2937"
FIG_WIDTH_CM = 16.0


def _setup_matplotlib() -> None:
    available = {f.name for f in font_manager.fontManager.ttflist}
    for name in ("Malgun Gothic", "NanumGothic", "Noto Sans CJK KR", "AppleGothic"):
        if name in available:
            plt.rcParams["font.family"] = name
            break
    plt.rcParams.update({
        "axes.unicode_minus": False,
        "font.size": 9,
        "axes.titlesize": 10,
        "axes.labelsize": 9,
        "axes.edgecolor": "#9CA3AF",
        "axes.labelcolor": MPL_INK,
        "axes.spines.top": False,
        "axes.spines.right": False,
        "xtick.color": "#4B5563",
        "ytick.color": "#4B5563",
        "axes.grid": True,
        "grid.color": "#E5E7EB",
        "grid.linewidth": 0.6,
        "figure.dpi": 200,
        "savefig.dpi": 200,
        "savefig.bbox": "tight",
    })


def verdict_of(score: Optional[float], overfit_gap: Optional[float], metric: str = "점수") -> Tuple[str, str]:
    """Rule-based verdict shared by summary, cover and JSON output."""
    if score is None:
        return "부적합", "성능 지표 없음"
    if score >= PASS_SCORE:
        if overfit_gap is not None and overfit_gap > OVERFIT_LIMIT:
            return "조건부", f"검증 {metric}는 기준을 충족하나 학습-검증 간극 {overfit_gap:.3f} > {OVERFIT_LIMIT}"
        return "적합", f"검증 {metric} {score:.3f} >= {PASS_SCORE}"
    if score >= CONDITIONAL_SCORE:
        return "조건부", f"검증 {metric} {score:.3f} (기준 {PASS_SCORE} 미달, {CONDITIONAL_SCORE} 이상)"
    return "부적합", f"검증 {metric} {score:.3f} < {CONDITIONAL_SCORE}"


def worst_verdict(verdicts: List[str]) -> str:
    order = ["부적합", "조건부", "적합"]
    return min(verdicts, key=order.index) if verdicts else "부적합"


def fmt(v: Any, digits: int = 4) -> str:
    """Numbers in the report: `digits` decimals, but small values keep `digits` significant digits
    (same rule as round_pred in the JSON), so an MAE of 3e-06 is not shown as 0.0000."""
    if v is None:
        return "-"
    if isinstance(v, (bool, np.bool_)):
        return "예" if v else "아니오"
    if isinstance(v, (int, np.integer)):
        return f"{int(v):,}"
    if isinstance(v, (float, np.floating)):
        if not np.isfinite(v):
            return "-"
        a = abs(float(v))
        dec = digits if a == 0 else max(digits, digits - 1 - int(np.floor(np.log10(a))))
        if 0 < a < 1e-4:
            return f"{float(v):.{digits - 1}e}"
        return f"{float(v):,.{dec}f}"
    return str(v)


# ==============================================================================
# 2. Word building blocks
# ==============================================================================

def _rgb(hex_color: str) -> "RGBColor":
    return RGBColor.from_string(hex_color)


def _set_run_font(run, size: Optional[float] = None, bold: Optional[bool] = None,
                  color: Optional[str] = None, font: str = FONT_KO) -> None:
    run.font.name = font
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rfonts.set(qn(attr), font)
    if size is not None:
        run.font.size = Pt(size)
    if bold is not None:
        run.font.bold = bold
    if color is not None:
        run.font.color.rgb = _rgb(color)


def _style_font(style, size: float, bold: bool = False, color: str = C_INK) -> None:
    style.font.name = FONT_KO
    style.font.size = Pt(size)
    style.font.bold = bold
    style.font.color.rgb = _rgb(color)
    rpr = style.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rfonts.set(qn(attr), FONT_KO)
    for attr in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
        if rfonts.get(qn(attr)) is not None:
            del rfonts.attrib[qn(attr)]


# Child order required by the WordprocessingML schema. Word 2007 refuses to open documents that break it,
# and it only knows w:left/w:right (w:start/w:end are Word 2010+ names).
_TCPR_ORDER = ("cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd", "noWrap", "tcMar",
               "textDirection", "tcFitText", "vAlign", "hideMark")
_TBLPR_ORDER = ("tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize", "tblStyleColBandSize", "tblW",
                "jc", "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout", "tblCellMar", "tblLook")
_PPR_ORDER = ("pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr",
              "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap",
              "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid",
              "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc", "textDirection",
              "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange")


def _set_child(parent, el, order: Tuple[str, ...]) -> None:
    """Put el into parent at its schema position, replacing an earlier element of the same kind."""
    old = parent.find(el.tag)
    if old is not None:
        parent.remove(old)
    name = el.tag.rsplit("}", 1)[-1]
    later = {qn(f"w:{n}") for n in order[order.index(name) + 1:]}
    for child in parent:
        if child.tag in later:
            child.addprevious(el)
            return
    parent.append(el)


def _shade(cell, hex_color: str) -> None:
    tcpr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    _set_child(tcpr, shd, _TCPR_ORDER)


def _cell_margins(cell, top=60, bottom=60, left=90, right=90) -> None:
    tcpr = cell._tc.get_or_add_tcPr()
    mar = OxmlElement("w:tcMar")
    for k, v in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        el = OxmlElement(f"w:{k}")
        el.set(qn("w:w"), str(v))
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    _set_child(tcpr, mar, _TCPR_ORDER)


def _table_borders(table, color: str = C_LINE, inside: bool = True) -> None:
    tblpr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    edges = ["top", "left", "bottom", "right"] + (["insideH", "insideV"] if inside else [])
    for edge in edges:
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "4")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)
        borders.append(el)
    _set_child(tblpr, borders, _TBLPR_ORDER)


def _cell_left_bar(cell, color: str, size: int = 18) -> None:
    tcpr = cell._tc.get_or_add_tcPr()
    borders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{edge}")
        if edge == "left":
            el.set(qn("w:val"), "single")
            el.set(qn("w:sz"), str(size))
            el.set(qn("w:color"), color)
        else:
            el.set(qn("w:val"), "nil")
        borders.append(el)
    _set_child(tcpr, borders, _TCPR_ORDER)


def _fix_widths(table, widths_cm: List[float]) -> None:
    """Fixed table layout with explicit grid column widths (cell widths alone are ignored by some renderers)."""
    tblpr = table._tbl.tblPr
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    _set_child(tblpr, layout, _TBLPR_ORDER)
    for col, w in zip(table.columns, widths_cm):
        col.width = Cm(w)
    for row in table.rows:
        for cell, w in zip(row.cells, widths_cm):
            cell.width = Cm(w)


def _repeat_header(row) -> None:
    trpr = row._tr.get_or_add_trPr()
    el = OxmlElement("w:tblHeader")
    el.set(qn("w:val"), "true")
    trpr.append(el)


def _no_split(row) -> None:
    trpr = row._tr.get_or_add_trPr()
    trpr.append(OxmlElement("w:cantSplit"))


def _add_field(paragraph, instr: str, size: float = 8, color: str = C_MUTED) -> None:
    for kind, text in (("begin", None), ("instr", instr), ("separate", None), ("text", "1"), ("end", None)):
        run = paragraph.add_run()
        _set_run_font(run, size=size, color=color)
        if kind == "instr":
            it = OxmlElement("w:instrText")
            it.set(qn("xml:space"), "preserve")
            it.text = f" {text} "
            run._r.append(it)
        elif kind == "text":
            t = OxmlElement("w:t")
            t.text = text
            run._r.append(t)
        else:
            fc = OxmlElement("w:fldChar")
            fc.set(qn("w:fldCharType"), kind)
            run._r.append(fc)


def _para_border_bottom(paragraph, color: str = C_LINE, size: int = 6) -> None:
    ppr = paragraph._p.get_or_add_pPr()
    pbdr = OxmlElement("w:pBdr")
    b = OxmlElement("w:bottom")
    b.set(qn("w:val"), "single")
    b.set(qn("w:sz"), str(size))
    b.set(qn("w:space"), "1")
    b.set(qn("w:color"), color)
    pbdr.append(b)
    _set_child(ppr, pbdr, _PPR_ORDER)


# ------------------------------------------------------------------------------
# Document template (DRM-safe)
# ------------------------------------------------------------------------------
# docx.Document() opens python-docx's templates/default.docx. Corporate DRM may encrypt or replace
# files with Office extensions, leaving a substitute document without the standard styles
# ('Caption' etc.) and with placeholder text. python-docx ships the same template unpacked as plain
# XML files (templates/default-docx-template/), which DRM does not touch, so that folder is used first.

# Paragraph styles the report uses: name -> (outline level for headings or None, bold, size pt)
REPORT_STYLES = {
    "Normal": (None, False, 10),
    "Heading 1": (0, True, 15),
    "Heading 2": (1, True, 12.5),
    "Heading 3": (2, True, 11),
    "Caption": (None, False, 8.5),
    "List Bullet": (None, False, 10),
}


def docx_template_dir() -> str:
    return os.path.join(os.path.dirname(docx.__file__), "templates", "default-docx-template")


def _clear_body(doc) -> int:
    """Removes body content (keeps the final section properties); returns the number of removed blocks."""
    body = doc.element.body
    removed = 0
    for child in list(body):
        if child.tag != qn("w:sectPr"):
            body.remove(child)
            removed += 1
    return removed


def _ensure_styles(doc) -> List[str]:
    """Creates every report style missing from a damaged template (based on Normal); returns created names."""
    st = doc.styles
    have = {s.name for s in st}
    created = []
    for name, (outline, bold, size) in REPORT_STYLES.items():
        if name in have:
            continue
        style = st.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
        if name != "Normal":
            style.base_style = st["Normal"]
            style.next_paragraph_style = st["Normal"]
        style.font.bold = bold
        style.font.size = Pt(size)
        if outline is not None:
            style.paragraph_format.keep_with_next = True
            ppr = style.element.get_or_add_pPr()
            lvl = OxmlElement("w:outlineLvl")
            lvl.set(qn("w:val"), str(outline))
            _set_child(ppr, lvl, _PPR_ORDER)
        created.append(name)
    return created


def docx_from_template_dir(tdir: str):
    """Opens python-docx's unpacked template folder as a document, tolerating missing optional parts.
    On the company PC a security product removed docProps/thumbnail.jpeg from the folder, and python-docx
    fails on any part a relationship points to but that is missing. The folder is packed into an in-memory
    .docx without those relationships. Returns (document, list of dropped part names)."""
    import io
    import zipfile
    import posixpath
    import xml.etree.ElementTree as ET
    rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    ct_ns = "http://schemas.openxmlformats.org/package/2006/content-types"
    files: Dict[str, bytes] = {}
    for root, _dirs, names in os.walk(tdir):
        for n in names:
            full = os.path.join(root, n)
            with open(full, "rb") as f:
                files[os.path.relpath(full, tdir).replace(os.sep, "/")] = f.read()
    dropped: List[str] = []
    for rels_name in [k for k in files if k.endswith(".rels")]:
        # "_rels/.rels" -> source folder "", "word/_rels/document.xml.rels" -> "word"
        base = posixpath.dirname(posixpath.dirname(rels_name))
        ET.register_namespace("", rel_ns)
        tree = ET.fromstring(files[rels_name])
        changed = False
        for rel in list(tree):
            if rel.get("TargetMode") == "External":
                continue
            target = rel.get("Target", "")
            part = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join(base, target))
            if part not in files:
                tree.remove(rel)
                dropped.append(part)
                changed = True
        if changed:
            files[rels_name] = ET.tostring(tree, xml_declaration=True, encoding="UTF-8")
    if dropped and "[Content_Types].xml" in files:
        ET.register_namespace("", ct_ns)
        ct = ET.fromstring(files["[Content_Types].xml"])
        for ov in list(ct):
            if (ov.get("PartName") or "").lstrip("/") in dropped:
                ct.remove(ov)
        files["[Content_Types].xml"] = ET.tostring(ct, xml_declaration=True, encoding="UTF-8")
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for name in sorted(files, key=lambda k: (k != "[Content_Types].xml", k)):
            z.writestr(name, files[name])
    buf.seek(0)
    return docx.Document(buf), dropped


def new_report_document():
    """Opens a blank document that always has the report styles (JIT recovery for damaged templates)."""
    problems: List[str] = []
    doc = None
    tdir = docx_template_dir()
    if os.path.isdir(tdir):
        try:
            doc, dropped = docx_from_template_dir(tdir)
            if dropped:
                note_recovered(f"Word 내장 템플릿 폴더에 빠진 파일({', '.join(dropped)}) -> 제외하고 사용")
        except Exception as e:
            problems.append(f"내장 템플릿 폴더 열기 실패({type(e).__name__}: {e})")
    else:
        problems.append("내장 템플릿 폴더 없음")
    if doc is None:
        try:
            doc = docx.Document()
        except Exception as e:
            raise InputError(
                f"python-docx 기본 템플릿을 열 수 없습니다 ({'; '.join(problems)}; default.docx: {type(e).__name__}: {e}). "
                f"보안 프로그램(DRM)이 python-docx 설치 폴더의 templates 파일을 바꿨을 수 있습니다. "
                f"관리자에게 python-docx 재설치를 요청하십시오: python -m pip install --force-reinstall python-docx"
            ) from e
        note_recovered(f"Word 내장 템플릿 폴더 사용 불가({'; '.join(problems)}) -> default.docx 사용")
        # A substitute template may carry its own header/footer content: start from fresh ones
        for sec in doc.sections:
            for hf in (sec.header, sec.footer, sec.first_page_header, sec.first_page_footer):
                try:
                    hf.is_linked_to_previous = True
                except Exception:
                    pass
    doc.element.body.get_or_add_sectPr()
    removed = _clear_body(doc)
    if removed:
        note_recovered(f"Word 템플릿에 본문 {removed}개 있음(보안 프로그램 대체 문서 추정) -> 제거 후 작성")
    created = _ensure_styles(doc)
    if created:
        note_recovered(f"Word 템플릿에 스타일 없음({', '.join(created)}) -> Normal 기반으로 생성")
    return doc


class ReportDoc:
    """Thin layer over python-docx enforcing one consistent visual system."""

    def __init__(self, header_text: str):
        self.doc = new_report_document()
        self.table_no = 0
        self.figure_no = 0
        self._setup_page(header_text)
        self._setup_styles()

    # -- page / styles -------------------------------------------------------
    def _setup_page(self, header_text: str) -> None:
        sec = self.doc.sections[0]
        sec.orientation = WD_ORIENT.PORTRAIT
        sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
        sec.left_margin = sec.right_margin = Cm(2.0)
        sec.top_margin, sec.bottom_margin = Cm(2.2), Cm(2.0)
        sec.header_distance, sec.footer_distance = Cm(1.0), Cm(1.0)
        sec.different_first_page_header_footer = True  # clean cover page

        hp = sec.header.paragraphs[0]
        hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        _set_run_font(hp.add_run(header_text), size=8, color=C_MUTED)
        _para_border_bottom(hp)

        fp = sec.footer.paragraphs[0]
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _add_field(fp, "PAGE")
        _set_run_font(fp.add_run(" / "), size=8, color=C_MUTED)
        _add_field(fp, "NUMPAGES")

    def _setup_styles(self) -> None:
        st = self.doc.styles
        normal = st["Normal"]
        _style_font(normal, 10)
        normal.paragraph_format.space_after = Pt(4)
        normal.paragraph_format.line_spacing = 1.2
        for name, size, before, after in (("Heading 1", 15, 0, 10), ("Heading 2", 12.5, 14, 6), ("Heading 3", 11, 10, 4)):
            h = st[name]
            _style_font(h, size, bold=True, color=C_ACCENT if name == "Heading 1" else C_INK)
            h.paragraph_format.space_before = Pt(before)
            h.paragraph_format.space_after = Pt(after)
            h.paragraph_format.keep_with_next = True
        cap = st["Caption"]
        _style_font(cap, 8.5, color=C_MUTED)
        cap.font.italic = False
        cap.paragraph_format.space_before = Pt(3)
        cap.paragraph_format.space_after = Pt(10)

    # -- text ----------------------------------------------------------------
    def h1(self, text: str, page_break: bool = False) -> None:
        if page_break:
            self.page_break()
        p = self.doc.add_heading(text, level=1)
        p.paragraph_format.space_before = Pt(16)
        _para_border_bottom(p, color=C_ACCENT, size=8)

    def spacer(self, pt: float = 4) -> None:
        """표와 표 사이를 나누는 아주 낮은 빈 단락 (Word는 붙은 표를 하나로 합치므로 필요)."""
        p = self.doc.add_paragraph()
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        p.paragraph_format.line_spacing = Pt(pt)

    def h2(self, text: str) -> None:
        self.doc.add_heading(text, level=2)

    def h3(self, text: str) -> None:
        self.doc.add_heading(text, level=3)

    def para(self, text: str, size: float = 10, color: str = C_INK, bold: bool = False,
             align=None, space_after: Optional[float] = None) -> Any:
        p = self.doc.add_paragraph()
        if align is not None:
            p.alignment = align
        if space_after is not None:
            p.paragraph_format.space_after = Pt(space_after)
        _set_run_font(p.add_run(text), size=size, color=color, bold=bold)
        return p

    def bullets(self, items: List[str]) -> None:
        for it in items:
            p = self.doc.add_paragraph(style="List Bullet")
            p.paragraph_format.space_after = Pt(2)
            _set_run_font(p.add_run(it), size=10, color=C_INK)

    def page_break(self) -> None:
        self.doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

    def caption(self, kind: str, text: str) -> None:
        if kind == "표":
            self.table_no += 1
            label = f"표 {self.table_no}. "
        else:
            self.figure_no += 1
            label = f"그림 {self.figure_no}. "
        p = self.doc.add_paragraph(style="Caption")
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER if kind == "그림" else WD_ALIGN_PARAGRAPH.LEFT
        # Table captions sit above the table: keep them on the same page
        p.paragraph_format.keep_with_next = kind == "표"
        _set_run_font(p.add_run(label), size=8.5, color=C_INK, bold=True)
        _set_run_font(p.add_run(text), size=8.5, color=C_MUTED)

    # -- tables --------------------------------------------------------------
    def table(self, header: List[str], rows: List[List[Any]], caption: Optional[str] = None,
              widths: Optional[List[float]] = None, align: Optional[List[str]] = None,
              highlight_rows: Optional[List[int]] = None, verdict_col: Optional[int] = None,
              font_size: float = 8.5) -> None:
        """Data table: shaded header, numbers right-aligned, optional highlighted rows (0-based body index)."""
        if caption:
            # Caption above tables (engineering convention), below figures
            self.caption("표", caption)
        n_cols = len(header)
        t = self.doc.add_table(rows=1 + len(rows), cols=n_cols)
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        t.autofit = False
        _table_borders(t)
        _fix_widths(t, widths or [17.0 / n_cols] * n_cols)
        highlight_rows = set(highlight_rows or [])
        if align is None:
            align = ["left"] + ["right"] * (n_cols - 1)
        amap = {"left": WD_ALIGN_PARAGRAPH.LEFT, "right": WD_ALIGN_PARAGRAPH.RIGHT, "center": WD_ALIGN_PARAGRAPH.CENTER}

        for r_i in range(len(rows) + 1):
            row = t.rows[r_i]
            _no_split(row)
            if r_i == 0:
                _repeat_header(row)
            for c_i in range(n_cols):
                cell = row.cells[c_i]
                _cell_margins(cell)
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                p = cell.paragraphs[0]
                p.paragraph_format.space_after = Pt(0)
                p.paragraph_format.line_spacing = 1.0
                # Short tables stay on one page
                p.paragraph_format.keep_with_next = len(rows) <= 15 and r_i < len(rows)
                if r_i == 0:
                    _shade(cell, C_HEAD)
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    _set_run_font(p.add_run(str(header[c_i])), size=font_size, bold=True, color=C_INK)
                    continue
                value = rows[r_i - 1][c_i]
                text = value if isinstance(value, str) else fmt(value)
                p.alignment = amap[align[c_i]]
                bold = (r_i - 1) in highlight_rows
                color = C_ACCENT if bold else C_INK
                if verdict_col is not None and c_i == verdict_col and text in VERDICT_STYLE:
                    fg, bg = VERDICT_STYLE[text]
                    _shade(cell, bg)
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    color, bold = fg, True
                elif bold:
                    _shade(cell, "F8FAFC")
                _set_run_font(p.add_run(text), size=font_size, bold=bold, color=color)
        self.spacer()

    def kv_table(self, pairs: List[Tuple[str, Any]], caption: Optional[str] = None,
                 widths: Tuple[float, float] = (4.2, 12.8)) -> None:
        if caption:
            self.caption("표", caption)
        t = self.doc.add_table(rows=len(pairs), cols=2)
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        t.autofit = False
        _table_borders(t)
        _fix_widths(t, list(widths))
        for i, (k, v) in enumerate(pairs):
            kc, vc = t.rows[i].cells
            for c in (kc, vc):
                _cell_margins(c)
                c.paragraphs[0].paragraph_format.space_after = Pt(0)
            _shade(kc, C_HEAD)
            _set_run_font(kc.paragraphs[0].add_run(str(k)), size=9, bold=True, color=C_INK)
            text = v if isinstance(v, str) else fmt(v)
            _set_run_font(vc.paragraphs[0].add_run(text), size=9, color=C_INK)
        self.spacer()

    def metric_cards(self, cards: List[Tuple[str, str, str]]) -> None:
        """Row of KPI cards: (label, value, sub-text)."""
        t = self.doc.add_table(rows=1, cols=len(cards))
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        t.autofit = False
        _fix_widths(t, [17.0 / len(cards)] * len(cards))
        for i, (label, value, sub) in enumerate(cards):
            c = t.rows[0].cells[i]
            _shade(c, C_HEAD)
            _cell_margins(c, top=120, bottom=120, left=140, right=140)
            _cell_left_bar(c, C_ACCENT, size=16)
            p1 = c.paragraphs[0]
            p1.paragraph_format.space_after = Pt(2)
            _set_run_font(p1.add_run(label), size=8, color=C_MUTED)
            p2 = c.add_paragraph()
            p2.paragraph_format.space_after = Pt(1)
            _set_run_font(p2.add_run(value), size=15 if len(value) <= 10 else (11 if len(value) <= 16 else 9.5), bold=True, color=C_ACCENT)
            p3 = c.add_paragraph()
            p3.paragraph_format.space_after = Pt(0)
            _set_run_font(p3.add_run(sub), size=7.5, color=C_MUTED)
        self.spacer(6)

    def note(self, title: str, text: str) -> None:
        """Finding box: light background, navy left bar."""
        t = self.doc.add_table(rows=1, cols=1)
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        _fix_widths(t, [17.0])
        _no_split(t.rows[0])  # keep the finding box on one page
        c = t.rows[0].cells[0]
        _shade(c, "F8FAFC")
        _cell_margins(c, top=100, bottom=100, left=160, right=160)
        _cell_left_bar(c, C_ACCENT)
        p = c.paragraphs[0]
        p.paragraph_format.space_after = Pt(3)
        _set_run_font(p.add_run(title), size=9, bold=True, color=C_ACCENT)
        for line in [s for s in text.split("\n") if s.strip()]:
            q = c.add_paragraph()
            q.paragraph_format.space_after = Pt(1)
            q.paragraph_format.line_spacing = 1.25
            _set_run_font(q.add_run(line), size=9.5, color=C_INK)
        self.spacer(6)

    def formula_table(self, rows: List[Tuple[str, str, str]], caption: str) -> None:
        """파생변수 표: 이름 | 계산식(크고 굵게, 강조 배경) | 사용 타겟과 기여도."""
        self.caption("표", caption)
        t = self.doc.add_table(rows=1 + len(rows), cols=3)
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        t.autofit = False
        _table_borders(t)
        _fix_widths(t, [3.6, 8.6, 4.8])
        for i, h in enumerate(("파생변수", "계산식", "사용 타겟 (기여도, 순위)")):
            c = t.rows[0].cells[i]
            _shade(c, C_HEAD)
            _cell_margins(c)
            c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            _set_run_font(c.paragraphs[0].add_run(h), size=8.5, bold=True, color=C_INK)
        for r_i, (name, formula, used) in enumerate(rows, 1):
            row = t.rows[r_i]
            _no_split(row)
            for c_i, (text, size, bold, color) in enumerate(((name, 9.5, True, C_INK), (formula, 11, True, C_ACCENT),
                                                            (used, 8.5, False, C_INK))):
                c = row.cells[c_i]
                _cell_margins(c, top=90, bottom=90)
                c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                if c_i == 1:
                    _shade(c, C_ACCENT_BG)
                q = c.paragraphs[0]
                q.paragraph_format.space_after = Pt(0)
                for k, line in enumerate(text.split("\n")):
                    if k:
                        q = c.add_paragraph()
                        q.paragraph_format.space_after = Pt(0)
                    _set_run_font(q.add_run(line), size=size, bold=bold, color=color)
        self.spacer()

    # -- figures -------------------------------------------------------------
    def figure(self, fig, caption: str, width_cm: Optional[float] = None) -> None:
        if width_cm is None:
            width_cm = min(fig.get_size_inches()[0] * 2.54, FIG_WIDTH_CM)
        buf = io.BytesIO()
        fig.savefig(buf, format="png", facecolor="white")
        plt.close(fig)
        buf.seek(0)
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.keep_with_next = True
        p.paragraph_format.space_after = Pt(0)
        p.add_run().add_picture(buf, width=Cm(width_cm))
        self.caption("그림", caption)

    def save(self, path: str) -> str:
        return save_with_lock_fallback(path, self.doc.save, what="Word 보고서")


# ==============================================================================
# 3. Charts (matplotlib -> PNG)
# ==============================================================================

def _fig(h_cm: float = 6.5, w_cm: float = FIG_WIDTH_CM, ncols: int = 1):
    return plt.subplots(1, ncols, figsize=(w_cm / 2.54, h_cm / 2.54))


def _draw_actual_vs_pred(ax, oof: pd.DataFrame, target: str, r2: float,
                         excluded: Optional[pd.DataFrame] = None) -> None:
    a, p = oof["actual"].values, oof["predicted"].values
    ea = ep = np.array([])
    if excluded is not None and len(excluded):
        ea = pd.to_numeric(excluded["실제값"], errors="coerce").to_numpy(dtype=float)
        ep = pd.to_numeric(excluded["예측값"], errors="coerce").to_numpy(dtype=float)
        ok = np.isfinite(ea) & np.isfinite(ep)
        ea, ep = ea[ok], ep[ok]
    lo = float(min(a.min(), p.min(), *(ea.tolist() + ep.tolist())))
    hi = float(max(a.max(), p.max(), *(ea.tolist() + ep.tolist())))
    pad = (hi - lo) * 0.04 or 1.0
    ax.plot([lo - pad, hi + pad], [lo - pad, hi + pad], color=MPL_ORANGE, lw=1, ls="--", label="실측 = 예측")
    ax.scatter(a, p, s=9, color=MPL_NAVY, alpha=0.55, edgecolors="none")
    if len(ea):
        ax.scatter(ea, ep, s=26, marker="x", color=MPL_ORANGE, linewidths=1.2, label=f"학습 제외 행 ({len(ea)})")
    ax.set_xlim(lo - pad, hi + pad)
    ax.set_ylim(lo - pad, hi + pad)
    ax.set_xlabel(f"실측 {target}")
    ax.set_ylabel(f"예측 {target}")
    ax.text(0.03, 0.95, f"검증 R² = {r2:.4f}\nn = {len(a):,}", transform=ax.transAxes, va="top", fontsize=8,
            bbox=dict(boxstyle="round,pad=0.3", fc="white", ec=MPL_LIGHT))
    ax.legend(loc="lower right", fontsize=7, frameon=False)
    ax.set_title("실측 대비 예측 (교차검증)", loc="left")


def _draw_confusion(ax, cm: Dict[str, Any]) -> None:
    labels, mat = cm["labels"], np.array(cm["matrix"], dtype=float)
    row_sum = mat.sum(axis=1, keepdims=True)
    pct = np.divide(mat, row_sum, out=np.zeros_like(mat), where=row_sum > 0)
    ax.imshow(pct, cmap="Blues", vmin=0, vmax=1)
    ax.grid(False)
    for i in range(len(labels)):
        for j in range(len(labels)):
            ax.text(j, i, f"{int(mat[i, j]):,}\n{pct[i, j]:.0%}", ha="center", va="center", fontsize=8,
                    color="white" if pct[i, j] > 0.6 else MPL_INK)
    ax.set_xticks(range(len(labels)), labels)
    ax.set_yticks(range(len(labels)), labels)
    ax.set_xlabel("예측 클래스")
    ax.set_ylabel("실제 클래스")
    for sp in ax.spines.values():
        sp.set_visible(False)
    ax.set_title("혼동행렬 (건수, 실제 대비 비율)", loc="left")


def _draw_importance(ax, ranks: List[Dict[str, Any]], derived: List[str], top_n: int = 8) -> None:
    """기여도 막대. 파생변수는 주황색으로 구별합니다."""
    top = ranks[:top_n][::-1]
    vals = [r["pct"] for r in top]
    colors = [MPL_ORANGE if r["feature"] in derived else MPL_NAVY for r in top]
    bars = ax.barh([r["feature"] for r in top], vals, color=colors, height=0.6)
    for b, v in zip(bars, vals):
        ax.text(b.get_width() + 0.5, b.get_y() + b.get_height() / 2, f"{v:.1f}%", va="center", fontsize=7.5)
    ax.set_xlim(0, max(vals + [1]) * 1.2)
    ax.set_xlabel("기여도 (%)")
    ax.grid(axis="y", visible=False)
    ax.tick_params(axis="y", labelsize=7.5)
    ax.set_title("피처 기여도" + (" (주황: 파생변수)" if any(r["feature"] in derived for r in top) else ""), loc="left")


def chart_target_overview(t: Dict[str, Any], derived: List[str]):
    """타겟 하나를 그림 한 장으로: 왼쪽 성능(회귀: 실측 대비 예측, 분류: 혼동행렬), 오른쪽 피처 기여도."""
    c = t["champ"]
    n = min(8, len(c["feature_importances"]))
    fig, (ax1, ax2) = _fig(h_cm=max(6.5, 1.6 + 0.62 * n), ncols=2)
    if t["task"] == "regression":
        _draw_actual_vs_pred(ax1, t["oof"], t["name"], c["val_r2"], t.get("excluded_rows"))
    else:
        _draw_confusion(ax1, c["confusion_matrix"])
    _draw_importance(ax2, c["feature_importances"], derived, top_n=n)
    fig.tight_layout(w_pad=2.5)
    return fig


def chart_correlation(corr: pd.DataFrame):
    n = corr.shape[0]
    size = min(16.0, 3.5 + 0.75 * n)
    fig, ax = _fig(h_cm=size - 1.5, w_cm=size)
    im = ax.imshow(corr.values, cmap="RdBu_r", vmin=-1, vmax=1)
    ax.grid(False)
    ax.set_xticks(range(n), corr.columns, rotation=60, ha="right", fontsize=7)
    ax.set_yticks(range(n), corr.index, fontsize=7)
    if n <= 16:
        for i in range(n):
            for j in range(n):
                v = corr.values[i, j]
                if np.isfinite(v):
                    ax.text(j, i, f"{v:.2f}", ha="center", va="center", fontsize=6,
                            color="white" if abs(v) > 0.6 else MPL_INK)
    for s in ax.spines.values():
        s.set_visible(False)
    fig.colorbar(im, ax=ax, fraction=0.035, pad=0.02)
    return fig


# ==============================================================================
# 4. Artifact loading
# ==============================================================================

def _read_jsonl(path: str) -> List[Dict[str, Any]]:
    out = []
    with open(path, "r", encoding="utf-8") as f:
        for line in f:
            if line.strip():
                try:
                    out.append(json.loads(line))
                except json.JSONDecodeError:
                    sys.stderr.write("[WARN] history.jsonl 손상 라인 1건을 건너뜀\n")
    return out


def load_version(version_dir: str, target_filter: Optional[List[str]] = None) -> Dict[str, Any]:
    vdir = os.path.abspath(version_dir)
    if not os.path.isdir(vdir):
        raise FileNotFoundError(f"버전 폴더를 찾을 수 없습니다: {vdir} (학습 결과 JSON의 version_dir 값을 지정하십시오)")
    lb_path = os.path.join(vdir, "leaderboard.json")
    if not os.path.isfile(lb_path):
        raise InputError(f"leaderboard.json이 없습니다. ml_train.py가 만든 버전 폴더(vX)를 지정하십시오: {vdir}")
    data_root = os.path.dirname(vdir)
    ledger = os.path.join(data_root, "history.jsonl")
    if not os.path.isfile(ledger):
        raise InputError(f"실험 원장(history.jsonl)이 없습니다: {ledger}")
    version = os.path.basename(vdir)
    records = _read_jsonl(ledger)
    rec = [r for r in records if r.get("version") == version]
    if not rec:
        raise InputError(f"원장에 버전 '{version}' 기록이 없습니다. [사용 가능한 버전]: {[r.get('version') for r in records]}")
    record = rec[-1]
    leaderboard = load_json_file(lb_path)

    names = list(record["targets"])
    if target_filter:
        wanted, unknown = [], []
        for f in target_filter:
            hit = [t for t in names if t == f or strip_unit(t) == strip_unit(f)]
            (wanted.extend(hit) if hit else unknown.append(f))
        if unknown:
            raise InputError(f"버전 '{version}'에 타겟 {unknown}이(가) 없습니다. [추천] {names}")
        names = list(dict.fromkeys(wanted))

    excluded_all = None
    p_ex = os.path.join(vdir, "excluded_rows.csv")
    if os.path.isfile(p_ex):
        try:
            excluded_all = pd.read_csv(io.StringIO(read_text_file(p_ex)))
        except Exception as e:  # the report still works without the excluded-row markers
            sys.stderr.write(f"[WARN] 학습 제외 행 파일을 읽지 못해 그림에 표시하지 않습니다: {e}\n")

    targets = []
    for t in names:
        safe = sanitize_filename(t)
        meta = load_json_file(os.path.join(vdir, f"model_{safe}.json"))
        oof_path = os.path.join(vdir, f"oof_{safe}.csv")
        if not os.path.isfile(oof_path):
            raise InputError(f"교차검증 예측 파일이 없습니다: {oof_path}. ml-kit 2.9.1 이상으로 같은 데이터를 다시 학습한 버전만 보고서를 만들 수 있습니다.")
        bench = leaderboard["targets"][t]
        targets.append({
            "name": t, "safe": safe, "task": record["task"], "meta": meta,
            "bench": bench, "champ": bench["champion"], "board": bench["leaderboard"],
            "oof": pd.read_csv(oof_path, encoding="utf-8-sig"),
            "classes": bench["champion"].get("classes"),
            "excluded": meta.get("excluded") or {},
            "excluded_rows": (excluded_all[excluded_all["타겟"].astype(str) == t]
                              if excluded_all is not None and "타겟" in excluded_all.columns else None),
        })

    corr = None
    corr_path = os.path.join(vdir, "correlation.csv")
    if os.path.isfile(corr_path):
        corr = pd.read_csv(corr_path, index_col=0, encoding="utf-8-sig")

    lineage = [r for r in records if r.get("task") == record["task"]
               and sorted(r.get("targets", [])) == sorted(record["targets"])]
    seen, uniq = set(), []
    for r in reversed(lineage):  # last record wins for re-trained versions
        if r["version"] not in seen:
            seen.add(r["version"])
            uniq.append(r)
    lineage = sorted(uniq, key=lambda r: r.get("timestamp", ""))

    # Every training of the same task (a best version gathers targets trained in different runs)
    last = {r.get("version"): i for i, r in enumerate(records)}
    task_history = [r for i, r in enumerate(records) if last.get(r.get("version")) == i
                    and r.get("task", "regression") == record.get("task", "regression")]
    task_history = sorted(task_history, key=lambda r: r.get("timestamp", ""))

    return {"version_dir": vdir, "data_root": data_root, "version": version, "record": record,
            "leaderboard": leaderboard, "targets": targets, "corr": corr, "lineage": lineage,
            "task_history": task_history}


def merged_feature_info(targets: List[Dict[str, Any]]) -> Tuple[Dict[str, Any], Dict[str, Any], Dict[str, Any], List[str]]:
    """Union of feature schema / stats / required inputs over the targets (they differ in a best version)."""
    schema: Dict[str, Any] = {"numeric": [], "categorical": {}, "datetime": {}}
    stats: Dict[str, Any] = {}
    dstats: Dict[str, Any] = {}
    required: List[str] = []
    for t in targets:
        meta = t["meta"]
        sc = meta.get("feature_schema") or {}
        schema["numeric"] += [c for c in sc.get("numeric", []) if c not in schema["numeric"]]
        for c, v in (sc.get("categorical") or {}).items():
            schema["categorical"].setdefault(c, v)
        for c, v in (sc.get("datetime") or {}).items():
            schema["datetime"].setdefault(c, v)
        for c, v in (meta.get("feature_stats") or {}).items():
            stats.setdefault(c, v)
        for c, v in (meta.get("datetime_stats") or {}).items():
            dstats.setdefault(c, v)
        required += [f for f in meta.get("features_required", []) if f not in required]
    return schema, stats, dstats, required


# ==============================================================================
# 5. Narrative (deterministic; --slots may override)
# ==============================================================================

def _metric_info(task: str) -> Tuple[str, str]:
    return ("val_r2", "R²") if task == "regression" else ("val_f1_macro", "Macro-F1")


def exclusion_line(t: Dict[str, Any], metric: str) -> str:
    """'학습 제외 3행 (직접 1, 자동 이상치 2) | 검증 R²(제외 행 포함) 0.81' or '' when nothing was left out."""
    e = t.get("excluded") or {}
    n = int(e.get("manual", 0)) + int(e.get("auto", 0))
    parts = []
    if n:
        parts.append(f"학습 제외 {n}행 (직접 표시 {e.get('manual', 0)}, 자동 이상치 {e.get('auto', 0)})")
        if isinstance(e.get("all_rows_score"), (int, float)):
            parts.append(f"검증 {metric}(제외 행 포함) {e['all_rows_score']:.4f}")
    if e.get("forced_include"):
        parts.append(f"'포함' 표시 {e['forced_include']}행은 자동 제외하지 않음")
    if e.get("auto_note"):
        parts.append(str(e["auto_note"]))
    return " | ".join(parts)


def text_performance(t: Dict[str, Any]) -> str:
    c = t["champ"]
    if t["task"] == "regression":
        std = float(t["oof"]["actual"].std()) or 1.0
        lines = [
            f"챔피언 {c['champion_model']}의 검증 R²는 {c['val_r2']:.4f}, MAE는 {fmt(c['mae'])}로 "
            f"타겟 표준편차({fmt(std)})의 {c['mae'] / std:.0%} 수준입니다.",
        ]
        gap = c.get("overfit_gap", 0.0)
        lines.append(f"학습-검증 간극 {gap:.4f}: " + ("과적합 위험이 낮습니다." if gap <= 0.10 else
                     "과적합 경향이 있어 피처 축소나 데이터 추가가 권장됩니다."))
        skew = c.get("residual_skew", 0.0)
        if abs(skew) > 0.5:          # 치우쳤을 때만 (대칭이면 생략)
            lines.append(f"오차가 한쪽으로 치우쳐(왜도 {skew:.2f}) 높은/낮은 구간에서 편향된 예측이 나타날 수 있습니다.")
        return "\n".join(lines)
    pc = c["per_class"]
    worst = min(pc, key=lambda k: pc[k]["recall"])
    lines = [
        f"챔피언 {c['champion_model']}의 검증 Macro-F1은 {c['val_f1_macro']:.4f}, 정확도 {c['accuracy']:.4f}, "
        f"균형 정확도 {c['balanced_accuracy']:.4f}입니다" + (f" (ROC-AUC {c['roc_auc']:.4f})." if c.get("roc_auc") is not None else "."),
        f"재현율이 가장 낮은 클래스는 '{worst}'({pc[worst]['recall']:.3f})이며, 이 클래스를 놓치는 오분류에 유의해야 합니다.",
    ]
    if c.get("class_weight"):
        lines.append("클래스 불균형으로 class_weight='balanced'가 적용되었습니다.")
    th = c.get("decision_threshold")
    if th:
        m = th["oof_positive_metrics"]
        lines.append(f"양성 클래스 '{th['positive_class']}' 판정 기준 확률은 {th['threshold']:.2f}이며, "
                     f"이때 양성 정밀도 {m['precision']:.3f}, 재현율 {m['recall']:.3f}, F1 {m['f1']:.3f}입니다.")
    return "\n".join(lines)


def text_features(t: Dict[str, Any]) -> str:
    c = t["champ"]
    top = c.get("top_features", [])
    schema = t["meta"]["feature_schema"]
    lines = []
    if top:
        lines.append("예측 기여도 상위 피처는 " + ", ".join(f"'{r['feature']}'({r['pct']:.1f}%)" for r in top) + "입니다.")
    if schema.get("categorical"):
        cats = [f for f in schema["categorical"] if any(r["feature"] == f for r in c["feature_importances"][:5])]
        if cats:
            lines.append(f"범주형 피처 {cats}이(가) 상위 기여 피처에 포함되어, 범주별 운전 조건 차이가 결과에 영향을 줍니다.")
    if schema.get("datetime"):
        lines.append(f"날짜 피처 {list(schema['datetime'])}은(는) {', '.join(sum(schema['datetime'].values(), []))} 단위로 분해해 반영했습니다.")
    weak = c.get("weak_features", [])
    if weak:
        lines.append(f"기여도 1% 이하 피처 {weak[:5]}은(는) 제외해도 성능 저하가 작을 가능성이 높습니다.")
    return "\n".join(lines) or "피처 중요도 정보가 없습니다."


def pretty_formula(expr: str) -> str:
    """'x2 = a * b / c ** 2' 의 오른쪽을 읽기 쉬운 수식으로: a × b ÷ c²."""
    rhs = expr.split("=", 1)[1].strip() if "=" in expr else expr.strip()
    rhs = re.sub(r"\s*\*\*\s*2(?![\d.])", "²", rhs)
    rhs = re.sub(r"\s*\*\*\s*3(?![\d.])", "³", rhs)
    rhs = rhs.replace("**", "^").replace("//", " 몫 ").replace("*", " × ").replace("/", " ÷ ")
    rhs = re.sub(r"\bsqrt\(", "√(", rhs)
    return re.sub(r"\s+", " ", rhs).strip()


def derived_feature_rows(targets: List[Dict[str, Any]]) -> Tuple[List[Tuple[str, str, str]], List[str]]:
    """타겟들이 쓴 파생변수: [(이름, 보기 좋은 계산식, '타겟 기여도% (순위)')], 이름 목록."""
    found: Dict[str, Dict[str, Any]] = {}
    for t in targets:
        ranks = t["champ"].get("feature_importances") or []
        for expr in t["meta"].get("derived_evals") or []:
            name = expr.split("=", 1)[0].strip()
            item = found.setdefault(name, {"formula": pretty_formula(expr), "used": []})
            pos = next((k for k, r in enumerate(ranks, 1) if r["feature"] == name), None)
            if pos is not None:
                item["used"].append(f"{t['name']}: {ranks[pos - 1]['pct']:.1f}% ({pos}위)")
            else:
                item["used"].append(t["name"])
    rows = [(n, f"{n} = {v['formula']}", "\n".join(v["used"])) for n, v in found.items()]
    return rows, list(found)


# ==============================================================================
# 6. Document sections
# ==============================================================================

def build_report(ctx: Dict[str, Any], slots: Dict[str, Any], out_path: Optional[str],
                 overwrite: bool = False, export: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    _setup_matplotlib()
    rec, targets, version = ctx["record"], ctx["targets"], ctx["version"]
    task = rec["task"]
    metric_key, metric_label = _metric_info(task)
    prof = rec.get("data_profile", {})
    data_name = prof.get("source_file", "")
    task_ko = "회귀 (수치 예측)" if task == "regression" else "분류 (범주 예측)"
    slot_targets = slots.get("targets", {}) if isinstance(slots.get("targets"), dict) else {}
    slots_used: List[str] = []

    def slot(key: str, default: str, target: Optional[str] = None) -> str:
        if target is not None and isinstance(slot_targets.get(target), dict) and slot_targets[target].get(key):
            slots_used.append(f"targets.{target}.{key}")
            return str(slot_targets[target][key])
        if slots.get(key):  # document-level text (applies to every target when no target-specific one)
            if key not in slots_used:
                slots_used.append(key)
            return str(slots[key])
        return default

    verdicts = {}
    for t in targets:
        verdicts[t["name"]] = verdict_of(t["champ"].get("val_score"), t["champ"].get("overfit_gap"), metric_label)
    overall = worst_verdict([v[0] for v in verdicts.values()])
    best_src = (rec.get("best_source") or {}) if rec.get("best") else {}

    def source_line(t: Dict[str, Any]) -> str:
        src = best_src.get(t["name"]) or {}
        old = src.get("score")
        return (f"학습 조건 출처: {src.get('version', '-')} (당시 검증 {metric_label} {old:.4f}), 이번 검증 {metric_label} {t['champ']['val_score']:.4f}"
                if isinstance(old, (int, float)) else f"학습 조건 출처: {src.get('version', '-')}, 이번 검증 {metric_label} {t['champ']['val_score']:.4f}")
    better = max((r for r in ctx["lineage"] if (r.get("mean_val_score") or -1) > rec["mean_val_score"] + 1e-9),
                 key=lambda r: r["mean_val_score"], default=None)

    R = ReportDoc(f"{REPORT_TITLE}  |  {data_name}  |  {version}")
    derived_rows, derived_names = derived_feature_rows(targets)

    # -- 제목 (표지 페이지 없이 첫 페이지에 바로 요약) ---------------------------------
    R.para(REPORT_TITLE, size=22, bold=True, color=C_ACCENT, space_after=2)
    R.para(rec.get("description") or "머신러닝 모델 학습 결과", size=12, color=C_INK, space_after=2)
    head = (f"{data_name}  ·  {version} ({rec.get('timestamp', '')[:10]})  ·  {task_ko}"
            + (f"  ·  메모: {rec['tag']}" if rec.get("tag") else "") + f"  ·  작성 {datetime.now():%Y-%m-%d}")
    p = R.para(head, size=8.5, color=C_MUTED, space_after=6)
    _para_border_bottom(p, color=C_ACCENT, size=12)
    if best_src:
        R.para(f"이 버전은 타겟마다 점수가 가장 높았던 학습 조건(피처, 파생변수)으로 {len(best_src)}개 타겟을 현재 데이터로 다시 학습한 통합 버전입니다.",
               size=9, color=C_MUTED, space_after=4)

    # -- 1. 요약 ------------------------------------------------------------------
    R.h1("1. 요약")
    best = max(targets, key=lambda t: t["champ"].get("val_score") or -1)
    R.metric_cards([
        ("종합 판정", overall, f"기준 검증 {metric_label} {PASS_SCORE} 이상"),
        (f"평균 검증 {metric_label}", f"{rec['mean_val_score']:.4f}", f"타겟 {len(targets)}개 평균"),
        ("학습 데이터", f"{prof.get('rows_loaded', 0):,}행", f"피처 {len(rec.get('features_used', []))}개"
                                                           + (f" (파생 {len(derived_rows)}개)" if derived_rows else "")),
        ("최고 성능 모델", best["champ"]["champion_model"], best["name"]),
    ])
    has_excl = any(isinstance((t.get("excluded") or {}).get("all_rows_score"), (int, float)) for t in targets)
    rows = []
    for t in targets:
        c = t["champ"]
        err = (f"MAE {fmt(c['mae'])}" if task == "regression" else f"정확도 {c['accuracy']:.4f}")
        row = [t["name"], c["champion_model"], c.get("val_score")]
        if has_excl:
            row.append((t.get("excluded") or {}).get("all_rows_score", c.get("val_score")))
        rows.append(row + [err, verdicts[t["name"]][0], ", ".join(r["feature"] for r in c.get("top_features", [])[:2]) or "-"])
    if has_excl:
        R.table(["타겟", "챔피언 모델", f"검증 {metric_label} (제외 행 미포함)", f"검증 {metric_label} (제외 행 포함)", "보조 지표", "판정", "주요 피처"], rows, caption="타겟별 결과 요약",
                widths=[2.4, 2.8, 2.2, 2.2, 2.2, 1.4, 3.8], align=["left", "left", "right", "right", "right", "center", "left"],
                verdict_col=5)
    else:
        R.table(["타겟", "챔피언 모델", f"검증 {metric_label}", "보조 지표", "판정", "주요 피처"], rows, caption="타겟별 결과 요약",
                widths=[3.0, 3.6, 1.8, 2.4, 1.5, 4.7], align=["left", "left", "right", "right", "center", "left"], verdict_col=4)
    recs = rec.get("recommendations") or []
    extra = [slots[k] for k in ("lineage", "operation") if slots.get(k)]          # 없앤 장의 소견은 요약에 붙임
    for k in ("lineage", "operation"):
        if slots.get(k):
            slots_used.append(k)
    summary_default = "\n".join(
        [f"{t['name']}: {verdicts[t['name']][0]} ({verdicts[t['name']][1]})" for t in targets]
        + [f"권고: {r}" for r in recs[:2]]
        + ([f"참고: 같은 조건에서 더 높은 점수의 버전 {better['version']}({better['mean_val_score']:.4f})이 있습니다."] if better else []))
    R.note("요약 소견", "\n".join([slot("summary", summary_default)] + [str(x) for x in extra]))
    criteria = [
        f"판정 기준: 검증 {metric_label} >= {PASS_SCORE} 적합, {CONDITIONAL_SCORE} 이상 조건부, 그 미만 부적합 "
        f"(적합이라도 학습-검증 간극이 {OVERFIT_LIMIT}를 넘으면 조건부)",
        "점수 용어",
        f"- 학습 {metric_label}: 모델이 배운 행을 다시 맞혀 본 점수 (과적합 확인용)",
        f"- 검증 {metric_label}: 모델이 배우지 않은 행을 맞혀 본 점수 (5겹 교차검증, 판정 기준)",
    ]
    if has_excl:
        criteria += [f"- 검증 {metric_label} (제외 행 미포함): 학습 제외 행 없이 계산 (판정 기준)",
                     f"- 검증 {metric_label} (제외 행 포함): 학습 제외 행까지 최종 모델로 예측해 함께 계산 (이상치가 섞인 실제 데이터 기준)"]
    R.para("\n".join(criteria), size=8, color=C_MUTED, space_after=2)
    if export and export.get("result"):
        ex = export["result"]
        lines = [f"성능예측표 (엑셀 계산기): {os.path.basename(ex['output_file'])}"]
        for t, v in ex["targets"].items():
            a = v["agreement"]
            lines.append(f"- {t}: 입력 {len(v['inputs'])}개 ({', '.join(strip_unit(i) for i in v['inputs'])}), "
                         f"모델 예측과 평균 {a['agreement_pct']:.1f}% 일치, 95% 조건에서 차이 {a['p95_diff_pct']:.1f}% 이내")
        R.para("\n".join(lines), size=8.5, color=C_INK, space_after=4)

    # -- 2. 데이터와 피처 -----------------------------------------------------------------
    R.h1("2. 데이터와 피처")
    layout = prof.get("layout", {})
    layout_txt = ", ".join(f"{k.replace('_row', '')} {v}행" for k, v in layout.items() if v) or "자동 감지"
    used = ", ".join(f"{t}: {n:,}행" for t, n in prof.get("rows_used", {}).items())
    data_default = (f"{data_name}: {prof.get('rows_loaded', 0):,}행 x {prof.get('columns_loaded', 0)}열 (표 구조: {layout_txt}). "
                    f"학습 사용 {used}. 피처 결측 {prof.get('missing_imputed_cells', 0):,}칸은 학습 데이터 기준값으로 대치.")
    R.para(slot("data", data_default), size=9, space_after=4)
    excl_lines = [f"{t['name']}: {exclusion_line(t, metric_label)}" for t in targets if exclusion_line(t, metric_label)]
    if excl_lines:
        R.para("학습 제외 행: 데이터의 '제외' 열에 글자가 있는 행(직접 표시)과, 값이 중앙값에서 크게 벗어난 행"
               "(중앙값·절대편차 기준 3.5배 초과, 자동 이상치)은 학습에서 뺐습니다. '제외' 열에 '포함'이라고 쓴 행은 자동으로 빼지 않습니다. "
               "뺀 행 목록은 버전 폴더의 excluded_rows.csv에 있습니다.\n" + "\n".join(excl_lines),
               size=9, space_after=4)
    schema, stats, dstats, required_inputs = merged_feature_info(targets)
    miss = prof.get("feature_missing", {})
    frows = []
    for f in schema["numeric"]:
        st = stats.get(f, {})
        frows.append([f, "파생" if f in derived_names else "수치", f"{fmt(st.get('min'), 3)} ~ {fmt(st.get('max'), 3)}", miss.get(f, 0)])
    for f, parts in schema.get("datetime", {}).items():
        st = dstats.get(f, {})
        frows.append([f, "날짜", f"{st.get('min', '')[:10]} ~ {st.get('max', '')[:10]} ({', '.join(parts)})", miss.get(f, 0)])
    for f, cats in schema.get("categorical", {}).items():
        shown = ", ".join(cats[:6]) + (f" 외 {len(cats) - 6}개" if len(cats) > 6 else "")
        frows.append([f, "범주", f"{len(cats)}개: {shown}", miss.get(f, 0)])
    R.table(["피처", "유형", "학습 범위 / 범주", "결측(칸)"], frows, caption="입력 피처",
            widths=[4.6, 1.4, 9.0, 2.0], align=["left", "center", "left", "right"],
            highlight_rows=[k for k, r in enumerate(frows) if r[1] == "파생"])

    R.h2("2.1. 파생변수")
    if derived_rows:
        R.para("아래 식으로 계산한 값을 입력 피처로 더해 학습했습니다. 예측할 때도 같은 식이 자동으로 적용됩니다.", size=9, space_after=3)
        R.formula_table(derived_rows, caption="파생변수 계산식")
    else:
        R.para("파생변수 없음 (입력 컬럼을 그대로 사용).", size=9, color=C_MUTED)

    R.h2("2.2. 상관관계")
    corr = ctx["corr"]
    if corr is not None and corr.shape[0] >= 2:
        order = [r["feature"] for r in targets[0]["champ"]["feature_importances"]]
        numeric_targets = [t["name"] for t in targets if t["task"] == "regression" and t["name"] in corr.columns]
        cols = [cc for cc in corr.columns if cc not in numeric_targets]
        owner_rank = {f: k for k, f in enumerate(order)}
        cols = sorted(cols, key=lambda cc: min([owner_rank.get(f, 999) for f in order if cc == f or cc.startswith(f + "_")] or [999]))
        keep = cols[:12] + numeric_targets
        R.figure(chart_correlation(corr.loc[keep, keep]), "수치 입력 및 타겟 간 피어슨 상관계수 (중요도 상위 12개)")
        pairs = []
        for a_i, a in enumerate(cols):
            for b in cols[a_i + 1:]:
                v = corr.loc[a, b]
                if np.isfinite(v) and abs(v) >= 0.85:
                    pairs.append((a, b, v))
        tgt_pairs = []
        for tn in numeric_targets:
            col = corr[tn].drop(labels=[tn], errors="ignore").drop(labels=numeric_targets, errors="ignore").dropna()
            if len(col):
                top = col.abs().sort_values(ascending=False).index[:2]
                tgt_pairs.append(f"{tn}: " + ", ".join(f"{f}({col[f]:+.2f})" for f in top))
        corr_default = "\n".join((["타겟과 상관이 큰 입력: " + "; ".join(tgt_pairs)] if tgt_pairs else [])
                                 + ["다중공선성 주의(|r| >= 0.85): " + (", ".join(f"{a}-{b}({v:+.2f})" for a, b, v in pairs[:5])
                                                                      if pairs else "없음")])
        R.note("상관관계 소견", slot("correlation", corr_default))
    else:
        R.para("수치 입력이 2개 미만이라 상관관계를 표시하지 않습니다.", size=9, color=C_MUTED)

    # -- 3. 분석 방법 (한 단락) -----------------------------------------------------------
    R.h1("3. 분석 방법")
    if task == "regression":
        method = ("후보 5종(HistGradientBoosting, RandomForest, ExtraTrees, Ridge, ElasticNet)을 5겹 교차검증으로 비교해 "
                  "검증 R²가 가장 높은 모델을 챔피언으로 고르고, 학습 행 전체로 다시 학습했습니다.")
    else:
        method = ("후보 5종(HistGradientBoosting, RandomForest, ExtraTrees, LogisticRegression, KNeighbors)을 클래스 비율을 유지한 "
                  "5겹 교차검증으로 비교해 검증 Macro-F1이 가장 높은 모델을 챔피언으로 고르고, 학습 행 전체로 다시 학습했습니다.")
    R.para(method + " 전처리(수치 중앙값 대치·표준화, 범주 원-핫, 날짜 분해)는 각 겹의 학습 데이터로만 맞춰 누수를 막았고, "
           "피처 기여도는 순열 중요도입니다.", size=9)

    # -- 4. 타겟별 결과 (타겟마다 핵심만) --------------------------------------------------
    R.h1("4. 타겟별 결과")
    for i, t in enumerate(targets, 1):
        c, board = t["champ"], t["board"]
        R.h2(f"4.{i}. {t['name']}")
        v, why = verdicts[t["name"]]
        R.para(f"판정: {v}  ({why})", size=9.5, bold=True, color=VERDICT_STYLE[v][0],
               space_after=1).paragraph_format.keep_with_next = True
        if best_src:
            R.para(source_line(t), size=9, color=C_INK, space_after=1).paragraph_format.keep_with_next = True
        if exclusion_line(t, metric_label):
            R.para(exclusion_line(t, metric_label), size=9, color=C_INK, space_after=1).paragraph_format.keep_with_next = True
        if task == "regression":
            iv = c.get("prediction_interval") or {}
            half = (iv["residual_upper"] - iv["residual_lower"]) / 2 if iv else None
            cards = [("검증 R²", f"{c['val_r2']:.4f}", c["champion_model"]), ("MAE", fmt(c['mae']), "평균 절대 오차"),
                     ("학습-검증 간극", f"{c.get('overfit_gap', 0):.4f}", "0.15 초과면 과적합 주의"),
                     ("90% 예측 범위", f"±{fmt(half, 3)}" if half is not None else "-", "예측값 기준")]
            hdr = ["순위", "모델", "검증 R²", "학습 R²", "간극", "MAE"]
            body = [[e["rank"], e["model_name"], e["val_r2"], e["train_r2"], e["overfit_gap"], e["mae"]] for e in board[:3]]
        else:
            cards = [("검증 Macro-F1", f"{c['val_f1_macro']:.4f}", c["champion_model"]), ("정확도", f"{c['accuracy']:.4f}", "전체 정답 비율"),
                     ("균형 정확도", f"{c['balanced_accuracy']:.4f}", "클래스별 재현율 평균"),
                     ("ROC-AUC", fmt(c.get("roc_auc")), "확률 순위 성능")]
            hdr = ["순위", "모델", "검증 Macro-F1", "학습 Macro-F1", "간극", "정확도"]
            body = [[e["rank"], e["model_name"], e["val_f1_macro"], e["train_f1_macro"], e["overfit_gap"], e["accuracy"]]
                    for e in board[:3]]
        R.metric_cards(cards)
        R.figure(chart_target_overview(t, derived_names),
                 f"{t['name']} " + ("실측 대비 예측" if task == "regression" else "혼동행렬") + "과 피처 기여도 (교차검증)")
        R.table(hdr, body, caption=f"{t['name']} 상위 3개 모델 (교차검증)", widths=[1.2, 5.0, 2.7, 2.7, 2.6, 2.8],
                align=["center", "left"] + ["right"] * 4, highlight_rows=[0])
        if task != "regression":
            pc = c["per_class"]
            R.table(["클래스", "정밀도", "재현율", "F1", "실제 건수"],
                    [[k, m["precision"], m["recall"], m["f1"], m["support"]] for k, m in pc.items()],
                    caption=f"{t['name']} 클래스별 성능", widths=[4.0, 3.0, 3.0, 3.0, 4.0])
        perf = slot("performance", text_performance(t), target=t["name"])
        feats = slot("features", text_features(t), target=t["name"])
        R.note("소견", perf + "\n" + feats)

    # -- Save --------------------------------------------------------------------
    if not out_path:
        out_path = os.path.join(ctx["data_root"], f"ML결과보고서_{os.path.splitext(data_name)[0]}_{version}.docx")
    out_abs = os.path.abspath(out_path)
    os.makedirs(os.path.dirname(out_abs), exist_ok=True)
    warn_long_path(out_abs)
    out_abs = avoid_overwrite(out_abs, overwrite, what="보고서")
    saved = R.save(out_abs)
    unverified = verify_docx(saved, version, len(targets))

    return {
        "status": "ok",
        "report_file": saved,
        "best_version": version,
        "version": version,
        "task": task,
        "targets": [t["name"] for t in targets],
        "overall_verdict": overall,
        "verdicts": {k: v[0] for k, v in verdicts.items()},
        "figure_count": R.figure_no,
        "table_count": R.table_no,
        "slots_used": slots_used,
        "verified": not unverified,
        **({"unverified": unverified} if unverified else {}),
        "preview": f"[보고서 완료] {version} / 타겟 {len(targets)}개 / 종합 판정 {overall} / 저장: {os.path.basename(saved)}",
        "agent_directive": (("VERIFIED: " if not unverified else "UNVERIFIED: 보고서 확인 실패 (" + "; ".join(unverified)
                             + "). 다시 실행하지 말고 unverified를 그대로 알리세요. ") + f"ML result report for {version} ({len(targets)} target(s)) saved to "
                            f"'{os.path.basename(saved)}'. Overall verdict={overall}. Next step: report the file path to the user."),
    }


def parse_slots(raw: Optional[str]) -> Dict[str, Any]:
    if not raw:
        return {}
    s_raw = raw.strip()
    try:
        if os.path.isfile(s_raw):
            data = load_json_file(s_raw)
        elif s_raw.startswith("{"):
            data = json.loads(s_raw)
        else:
            raise InputError("--slots는 JSON 파일 경로 또는 '{'로 시작하는 JSON 문자열이어야 합니다.")
        if not isinstance(data, dict):
            raise InputError("슬롯 JSON은 객체({...})여야 합니다.")
        return data
    except (InputError, ValueError) as e:
        sys.stderr.write(f"[WARN] 슬롯 JSON 파싱 실패 ({e}) -> 자동 작성 소견 사용\n")
        return {}


def resolve_version_dir(args: argparse.Namespace) -> str:
    """-m <version_dir> wins; otherwise -d <data_root> with -v, or the best version of the latest experiment."""
    if args.model:
        return args.model
    if not args.data_dir:
        raise InputError("보고서를 만들 버전 폴더(-m <version_dir>) 또는 데이터 폴더(-d <data_root>)를 지정하십시오.")
    root = os.path.abspath(args.data_dir)
    if args.version:
        return os.path.join(root, normalize_version(args.version))
    ledger = os.path.join(root, "history.jsonl")
    if not os.path.isfile(ledger):
        raise InputError(f"실험 원장(history.jsonl)이 없습니다: {ledger}")
    records = _read_jsonl(ledger)
    if not records:
        raise InputError(f"실험 원장이 비어 있습니다: {ledger}")
    latest = records[-1]
    same = [r for r in records if r.get("task", "regression") == latest.get("task", "regression")
            and sorted(r.get("targets", [])) == sorted(latest.get("targets", []))
            and isinstance(r.get("mean_val_score"), (int, float))]
    best = max(same, key=lambda r: r["mean_val_score"]) if same else latest
    sys.stderr.write(f"[INFO] 보고서 버전 자동 선택: {best['version']} (최근 학습과 같은 타겟 조합 중 최고 점수)\n")
    return os.path.join(root, best["version"])


def make_export(version_dir: str, task: str, title: Optional[str], targets: Optional[List[str]] = None,
                overwrite: bool = False) -> Optional[Dict[str, Any]]:
    """2.9.2: the report and the performance calculator (ml_export.py) are made as a set (regression versions only).
    Made BEFORE the Word report so the report can state the calculator's agreement with the model (2.9.4).
    targets: the report's -t targets (None = all); overwrite: the report's --overwrite also replaces the workbook.
    A failure never fails the report: {"note": ...}."""
    if task != "regression" or os.environ.get("ML_KIT_SKIP_EXPORT") == "1":
        return None
    try:
        import ml_export  # heavy (retraining); only when a report is made
        return {"result": ml_export.execute_export(argparse.Namespace(
            model=version_dir, targets=",".join(targets) if targets else None, keep=None, inputs=None, title=title,
            save=None, overwrite=overwrite))}
    except Exception as e:
        note = f"성능예측표를 만들지 못했습니다 (보고서는 정상): {type(e).__name__}: {str(e)[:200]}"
        sys.stderr.write(f"[WARN] {note}\n")
        return {"note": note}


def verify_docx(path: str, version: str, n_targets: int) -> List[str]:
    """R1: reopens the saved report and checks it is a readable Word file with the version and every target section."""
    try:
        from docx import Document
        doc = Document(path)
    except Exception as e:
        return [f"저장한 보고서를 다시 열지 못함: {type(e).__name__}: {e}"]
    text = "\n".join(p.text for p in doc.paragraphs)
    problems = []
    if version not in text:
        problems.append(f"보고서에 버전 '{version}'이 없음")
    if len(doc.tables) < n_targets:
        problems.append(f"표가 {len(doc.tables)}개뿐임 (타겟 {n_targets}개)")
    return problems


def attach_export(result: Dict[str, Any], export: Optional[Dict[str, Any]]) -> None:
    if not export:
        return
    if export.get("note"):
        result["export_note"] = export["note"]
        result["agent_directive"] += " export_note를 그대로 알리세요."
        return
    res = export["result"]
    result["export_file"] = res["output_file"]
    result["export"] = {t: {k: v[k] for k in ("inputs", "dropped", "val_r2", "table_rows", "agreement")}
                        for t, v in res["targets"].items()}
    result["export_verified"] = res["verified"]
    if res.get("unverified"):
        result["export_unverified"] = res["unverified"]
    if res.get("warnings"):
        result["export_warnings"] = res["warnings"]
    result["preview"] += f" | 성능예측표: {os.path.basename(res['output_file'])}"
    result["agent_directive"] += (" 성능예측표(엑셀 계산기) export_file도 함께 만들었습니다: 경로, 타겟별 남긴 입력(export.<타겟>.inputs), "
                                  "모델 대비 일치도(export.<타겟>.agreement.agreement_pct)를 알리세요."
                                  + (" export_unverified 내용도 그대로 알리세요." if res.get("unverified") else "")
                                  + (" export_warnings(성능예측표 주의 사항)도 그대로 알리세요." if res.get("warnings") else ""))


def execute_report(args: argparse.Namespace) -> Dict[str, Any]:
    for err, pkg in ((_DOCX_IMPORT_ERROR, "python-docx"), (_MPL_IMPORT_ERROR, "matplotlib")):
        if err is not None:
            raise InputError(f"보고서 생성에 필요한 {pkg} 패키지가 설치되어 있지 않습니다 ({err}). docs/ENVIRONMENT.md를 참고하십시오.")
    target_filter = None
    if args.target and args.target.strip().lower() != "all":
        target_filter = [t.strip() for t in args.target.split(",") if t.strip()]
    if getattr(args, "file", None) and not args.model and not args.data_dir:
        return execute_best_report(args, target_filter)
    vdir = resolve_version_dir(args)
    ctx = load_version(vdir, target_filter)
    export = make_export(vdir, ctx["record"]["task"], getattr(args, "title", None),
                         [t["name"] for t in ctx["targets"]] if target_filter else None, overwrite=args.overwrite)
    result = build_report(ctx, parse_slots(args.slots), args.out, overwrite=args.overwrite, export=export)
    attach_export(result, export)
    return result


def _numbered_out(out: Optional[str], version: str, many: bool) -> Optional[str]:
    if not out or not many:
        return out
    stem, ext = os.path.splitext(out)
    return f"{stem}_{version}{ext or '.docx'}"


def _layout_args(args: argparse.Namespace, records: List[Dict[str, Any]], task: str) -> argparse.Namespace:
    """Data layout for the predict command: the -f options, else the latest training of the task (as --best uses)."""
    import ml_train
    layout = ml_train._latest_layout(records, task)
    return argparse.Namespace(**{k: getattr(args, k, None) if getattr(args, k, None) is not None else layout.get(k)
                                 for k in ("header_row", "unit_row", "data_row")})


def execute_best_report(args: argparse.Namespace, target_filter: Optional[List[str]]) -> Dict[str, Any]:
    """
    -f <data file>: one report per task covering EVERY target of that task. The best version (each target
    retrained with its best recipe from the history) is reused when up to date, otherwise built first.
    """
    import ml_train  # heavy (scikit-learn); only this mode needs it
    info = ml_train.ensure_best_versions(args)
    slots = parse_slots(args.slots)
    here = os.path.dirname(os.path.abspath(__file__))
    versions = info["best_versions"]
    reports: List[Dict[str, Any]] = []
    export = None
    for bv in versions:
        flt = None
        if target_filter:
            flt = [f for f in target_filter if any(f == t or strip_unit(f) == strip_unit(t) for t in bv["targets"])]
            if not flt:
                continue
        ctx = load_version(bv["version_dir"], flt)
        ex = (make_export(bv["version_dir"], bv["task"], getattr(args, "title", None),
                          [t["name"] for t in ctx["targets"]] if flt else None, overwrite=args.overwrite)
              if export is None else None)
        export = export or ex
        reports.append(build_report(ctx, slots, _numbered_out(args.out, bv["version"], len(versions) > 1),
                                    overwrite=args.overwrite, export=ex))
    if not reports:
        all_targets = [t for bv in versions for t in bv["targets"]]
        raise InputError(f"지정한 타겟 {target_filter}이(가) 최고 버전에 없습니다. [후보] {', '.join(all_targets[:10])}")

    _, records = ml_train._ledger_for(os.path.abspath(args.file))
    result = dict(reports[0])
    result.update({
        "made_best": info["made_best"],
        "best_version": reports[0]["version"],
        "best_version_dir": next(bv["version_dir"] for bv in versions if bv["version"] == reports[0]["version"]),
        "best_versions": versions,
        "best_targets": info["best_targets"],
        "report_count": len(reports),
        "best_reports": [{k: r[k] for k in ("report_file", "version", "task", "targets", "overall_verdict", "verdicts")}
                    for r in reports],
        "next_commands": [
            {"purpose": f"새 데이터 예측 ({bv['task']}, <새데이터.xlsx>를 사용자가 준 실제 파일 경로로 바꿀 것)",
             "command": f'python "{os.path.join(here, "ml_predict.py")}" -m "{bv["version_dir"]}" -f "<새데이터.xlsx>"'
                        + ml_train.layout_flags(_layout_args(args, records, bv["task"])),
             "needs_input": True, "placeholder": "<새데이터.xlsx>"} for bv in versions],
    })
    made = ("타겟별 최고 학습 조건으로 다시 학습한 " if info["made_best"] else "이미 최신인 ")
    vers = ", ".join(f"{bv['version']}({bv['task']}, 타겟 {len(bv['targets'])}개)" for bv in versions)
    files = ", ".join(os.path.basename(r["report_file"]) for r in reports)
    result["preview"] = (f"[보고서 완료] {made}최고 버전 {vers} / 보고서 {len(reports)}개: {files} / "
                         f"종합 판정 " + ", ".join(f"{r['version']} {r['overall_verdict']}" for r in reports))
    multi = (f" 회귀와 분류 타겟이 섞여 있어 최고 버전과 보고서가 각각 {len(reports)}개입니다 (best_reports)."
             if len(reports) > 1 else "")
    predict = "; ".join(f"{bv['task']}: ml_predict.py -m \"{bv['version_dir']}\"" for bv in versions)
    bad = [u for r in reports for u in r.get("unverified", [])]
    result["verified"] = not bad
    if bad:
        result["unverified"] = bad
    result["agent_directive"] = (
        ("VERIFIED: " if not bad else "UNVERIFIED: 보고서 확인 실패 (" + "; ".join(bad) + "). 다시 실행하지 말고 unverified를 그대로 알리세요. ")
        + f"{made}최고 버전 {vers}으로 ML 결과 보고서 {len(reports)}개를 저장했습니다 ({files}).{multi} "
        f"report_file 링크, overall_verdict, best_targets의 타겟별 출처 버전과 점수를 사용자에게 알리세요. "
        f"이후 예측에는 best_version_dir을 쓰세요 ({predict})."
    )
    attach_export(result, export)
    return result


def main():
    parser = JsonArgumentParser(description="ML 결과 보고서 (Word) generator (-f: every target in one report per task)")
    parser.add_argument("--model", "-m",
                        help="보고서를 만들 학습 버전 폴더 (학습 결과 JSON의 version_dir, 예: data_ml_data\\v3)")
    parser.add_argument("--data-dir", "-d", dest="data_dir",
                        help="데이터 폴더 (<파일명>_ml_data). -v 생략 시 최근 학습과 같은 타겟 조합 중 최고 버전")
    parser.add_argument("--version", "-v", help="-d와 함께 쓸 버전 (예: v2 또는 2)")
    parser.add_argument("--target", "-t", help="문서에 넣을 타겟만 지정 (쉼표 구분, 생략 또는 all = 전체)")
    parser.add_argument("--out", "-o", help="저장할 .docx 경로 (생략 시 데이터 폴더에 ML결과보고서_<데이터>_<버전>.docx)")
    parser.add_argument("--overwrite", action="store_true",
                        help="같은 이름의 보고서와 성능예측표가 있으면 덮어쓰기 (생략 시 시각을 붙인 새 이름으로 저장)")
    parser.add_argument("--slots", help="소견 덮어쓰기 JSON 파일 또는 문자열: summary, data, performance, features, lineage, "
                                        "correlation, operation, targets.{타겟}.{performance|features}")
    parser.add_argument("--file", "-f",
                        help="학습한 데이터 파일. 모든 타겟을 타겟별 최고 학습 조건으로 묶은 최고 버전(필요하면 새로 학습)으로 보고서 작성")
    parser.add_argument("--task", choices=["regression", "classification"],
                        help="-f와 함께: 이 작업 유형의 보고서만 (생략 시 이력에 있는 유형마다 1개)")
    parser.add_argument("--header-row", type=int, help="-f와 함께: 최고 버전을 새로 학습할 때 헤더 행 (생략 시 이력 값)")
    parser.add_argument("--unit-row", type=int, help="-f와 함께: 단위 행 (생략 시 이력 값)")
    parser.add_argument("--data-row", type=int, help="-f와 함께: 데이터 시작 행 (생략 시 이력 값)")
    parser.add_argument("--title", help="함께 만드는 성능예측표(엑셀)의 제목 (데이터 설명을 짧게 정리한 글, 40자 이내)")
    args = parser.parse_args()
    if not args.model and not args.data_dir and not args.file:
        parser.error("-f <데이터파일>, -m <version_dir>, -d <data_root> 중 하나를 지정해야 합니다.")
    run_cli(lambda: execute_report(args), hint="Check the version folder path (version_dir from the training result).")


if __name__ == "__main__":
    main()
