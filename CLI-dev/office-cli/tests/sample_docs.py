"""테스트용 새 docx/pptx 만들기.

회사 PC(DRM)에서는 site-packages 안의 라이브러리 기본 템플릿이 망가져 있었습니다 (2026-11-05 실기):
pptx/templates/default.pptx는 지워졌고, docx/templates/default.docx는 zip이지만 스타일 4개뿐인 다른 문서였고,
폴더 템플릿(default-docx-template)에서는 docProps/thumbnail.jpeg가 빠져 있었습니다. 그러면 인자 없는
pptx.Presentation() / docx.Document()가 실패하거나 제목 스타일이 없는 문서를 줍니다.
테스트 샘플은 늘 new_docx() / new_pptx()로 만듭니다 (pptx.Presentation() / docx.Document()를 직접 부르지 않음).
- docx: 기본 .docx가 zip이고 'Heading 1' 스타일이 있으면 그것, 아니면 같은 내용의 폴더 템플릿으로 만듭니다.
  폴더 템플릿에서 빠진 파일을 가리키는 관계는 빼고 엽니다. 둘 다 못 쓰면 그 테스트만 건너뜁니다.
- pptx: 기본 템플릿이 zip이면 그것, 아니면 이 파일이 코드로 만드는 최소 OOXML 프레젠테이션(minimal_pptx_bytes)을
  엽니다. 슬라이드 레이아웃 순서는 기본 템플릿과 같습니다 (0 제목, 1 제목+내용, 5 제목만 등 11개).
"""
import io
import os
import zipfile

import pytest

PPTX_SKIP_REASON = "python-pptx 기본 템플릿이 손상됨 (DRM 추정)"
DOCX_SKIP_REASON = "python-docx 기본 템플릿이 손상됨 (DRM 추정)"


def docx_template_paths():
    """(기본 .docx 파일, 폴더 템플릿) 경로."""
    import docx
    base = os.path.join(os.path.dirname(docx.__file__), "templates")
    return os.path.join(base, "default.docx"), os.path.join(base, "default-docx-template")


def pptx_template_path():
    import pptx
    return os.path.join(os.path.dirname(pptx.__file__), "templates", "default.pptx")


def pptx_template_ok():
    return zipfile.is_zipfile(pptx_template_path())


def docx_from_folder(folder):
    """폴더 템플릿을 메모리 속 .docx로 묶어 엽니다. 관계(.rels)가 가리키는데 없는 파일은 관계째 뺍니다."""
    import posixpath
    import xml.etree.ElementTree as ET
    import docx
    files = {}
    for root, _dirs, names in os.walk(folder):
        for n in names:
            full = os.path.join(root, n)
            with open(full, "rb") as f:
                files[os.path.relpath(full, folder).replace(os.sep, "/")] = f.read()
    ET.register_namespace("", "http://schemas.openxmlformats.org/package/2006/relationships")
    for rels in [k for k in files if k.endswith(".rels")]:
        base = posixpath.dirname(posixpath.dirname(rels))
        tree = ET.fromstring(files[rels])
        for rel in list(tree):
            target = rel.get("Target", "")
            part = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join(base, target))
            if rel.get("TargetMode") != "External" and part not in files:
                tree.remove(rel)
        files[rels] = ET.tostring(tree, xml_declaration=True, encoding="UTF-8")
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for name in sorted(files, key=lambda k: (k != "[Content_Types].xml", k)):
            z.writestr(name, files[name])
    buf.seek(0)
    return docx.Document(buf)


def _usable_docx(d):
    """기본 템플릿으로 만든 문서가 정상인지: 제목 스타일이 있고 본문이 비어 있음 (DRM PC: 스타일 4개, 빈 단락 1개)."""
    names = {s.name for s in d.styles}
    return "Heading 1" in names and not d.paragraphs


def new_docx():
    """빈 python-docx 문서. 기본 .docx가 망가졌으면 폴더 템플릿을 씁니다."""
    import docx
    default, folder = docx_template_paths()
    if zipfile.is_zipfile(default):
        try:
            d = docx.Document()
            if _usable_docx(d):
                return d
        except Exception:
            pass
    if os.path.isdir(folder):
        return docx_from_folder(folder)
    pytest.skip(DOCX_SKIP_REASON)


def new_pptx():
    """빈 python-pptx 프레젠테이션. 기본 템플릿이 zip이 아니면 코드로 만든 최소 프레젠테이션을 엽니다."""
    import pptx
    if pptx_template_ok():
        return pptx.Presentation()
    return pptx.Presentation(io.BytesIO(minimal_pptx_bytes()))


# ---------------------------------------------------------------------------
# 최소 OOXML 프레젠테이션 (라이브러리 템플릿 없이)
# ---------------------------------------------------------------------------

_NS = ('xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
       'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
       'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"')
_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
_CT = "application/vnd.openxmlformats-officedocument."
_GRP = ('<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm>'
        '<a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>')

# 기본 템플릿과 같은 순서의 레이아웃: (type, 이름, 자리 목록[(ph type, idx)])
_LAYOUTS = [
    ("title", "Title Slide", [("ctrTitle", None), ("subTitle", 1)]),
    ("obj", "Title and Content", [("title", None), (None, 1)]),
    ("secHead", "Section Header", [("title", None), ("body", 1)]),
    ("twoObj", "Two Content", [("title", None), (None, 1), (None, 2)]),
    ("twoTxTwoObj", "Comparison", [("title", None), ("body", 1), (None, 2), ("body", 3), (None, 4)]),
    ("titleOnly", "Title Only", [("title", None)]),
    ("blank", "Blank", []),
    ("objTx", "Content with Caption", [("title", None), (None, 1), ("body", 2)]),
    ("picTx", "Picture with Caption", [("title", None), ("pic", 1), ("body", 2)]),
    ("vertTx", "Title and Vertical Text", [("title", None), ("body", 1)]),
    ("vertTitleAndTx", "Vertical Title and Text", [("title", None), ("body", 1)]),
]

# 자리 기본 위치 (EMU, 4:3 화면 9144000 x 6858000)
_BOX = {"title": (457200, 274638, 8229600, 1143000), "ctrTitle": (685800, 2130425, 7772400, 1470025),
        "subTitle": (1371600, 3886200, 6400800, 1752600)}
_BODY_BOX = (457200, 1600200, 8229600, 4525963)


def _ph_sp(n, ph_type, idx, master=False):
    attrs = (f' type="{ph_type}"' if ph_type else "") + (f' idx="{idx}"' if idx is not None else "")
    x, y, cx, cy = _BOX.get(ph_type or "", _BODY_BOX)
    sppr = (f'<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>') if master else "<p:spPr/>"
    name = f"{(ph_type or 'Content').title()} Placeholder {n}"
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{n + 1}" name="{name}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
            f'<p:nvPr><p:ph{attrs}/></p:nvPr></p:nvSpPr>{sppr}<p:txBody><a:bodyPr/><a:lstStyle/><a:p>'
            '<a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>')


def _theme():
    colors = "".join(f'<a:{k}><a:srgbClr val="{v}"/></a:{k}>' for k, v in (
        ("dk1", "000000"), ("lt1", "FFFFFF"), ("dk2", "1F497D"), ("lt2", "EEECE1"), ("accent1", "4F81BD"),
        ("accent2", "C0504D"), ("accent3", "9BBB59"), ("accent4", "8064A2"), ("accent5", "4BACC6"),
        ("accent6", "F79646"), ("hlink", "0000FF"), ("folHlink", "800080")))
    font = '<a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/>'
    fill = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
    line = f'<a:ln w="9525">{fill}</a:ln>'
    return (_HEAD + '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Minimal">'
            f'<a:themeElements><a:clrScheme name="Minimal">{colors}</a:clrScheme>'
            f'<a:fontScheme name="Minimal"><a:majorFont>{font}</a:majorFont><a:minorFont>{font}</a:minorFont></a:fontScheme>'
            f'<a:fmtScheme name="Minimal"><a:fillStyleLst>{fill * 3}</a:fillStyleLst><a:lnStyleLst>{line * 3}</a:lnStyleLst>'
            '<a:effectStyleLst>' + '<a:effectStyle><a:effectLst/></a:effectStyle>' * 3 + '</a:effectStyleLst>'
            f'<a:bgFillStyleLst>{fill * 3}</a:bgFillStyleLst></a:fmtScheme></a:themeElements>'
            '<a:objectDefaults/><a:extraClrSchemeLst/></a:theme>')


def minimal_pptx_bytes():
    """python-pptx 기본 템플릿 없이 여는 빈 4:3 프레젠테이션 (마스터 1, 레이아웃 11, 테마 1)."""
    rels = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{}</Relationships>'

    def rel(i, kind, target):
        return f'<Relationship Id="rId{i}" Type="{_REL}{kind}" Target="{target}"/>'

    files = {}
    over = [("/ppt/presentation.xml", "presentationml.presentation.main+xml"),
            ("/ppt/slideMasters/slideMaster1.xml", "presentationml.slideMaster+xml"),
            ("/ppt/theme/theme1.xml", "theme+xml")]
    layout_rels, layout_ids = [], []
    for i, (kind, name, phs) in enumerate(_LAYOUTS, 1):
        sps = "".join(_ph_sp(n, t, idx) for n, (t, idx) in enumerate(phs, 1))
        files[f"ppt/slideLayouts/slideLayout{i}.xml"] = (
            _HEAD + f'<p:sldLayout {_NS} type="{kind}" preserve="1"><p:cSld name="{name}"><p:spTree>{_GRP}{sps}'
            '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>')
        files[f"ppt/slideLayouts/_rels/slideLayout{i}.xml.rels"] = _HEAD + rels.format(
            rel(1, "slideMaster", "../slideMasters/slideMaster1.xml"))
        over.append((f"/ppt/slideLayouts/slideLayout{i}.xml", "presentationml.slideLayout+xml"))
        layout_rels.append(rel(i, "slideLayout", f"../slideLayouts/slideLayout{i}.xml"))
        layout_ids.append(f'<p:sldLayoutId id="{2147483648 + i}" r:id="rId{i}"/>')
    theme_rid = len(_LAYOUTS) + 1
    master_sps = _ph_sp(1, "title", None, master=True) + _ph_sp(2, "body", 1, master=True)
    files["ppt/slideMasters/slideMaster1.xml"] = (
        _HEAD + f'<p:sldMaster {_NS}><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>'
        f'<p:spTree>{_GRP}{master_sps}</p:spTree></p:cSld>'
        '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" '
        'accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>'
        f'<p:sldLayoutIdLst>{"".join(layout_ids)}</p:sldLayoutIdLst></p:sldMaster>')
    files["ppt/slideMasters/_rels/slideMaster1.xml.rels"] = _HEAD + rels.format(
        "".join(layout_rels) + rel(theme_rid, "theme", "../theme/theme1.xml"))
    files["ppt/theme/theme1.xml"] = _theme()
    files["ppt/presentation.xml"] = (
        _HEAD + f'<p:presentation {_NS} saveSubsetFonts="1"><p:sldMasterIdLst>'
        '<p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>'
        '<p:sldSz cx="9144000" cy="6858000" type="screen4x3"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>')
    files["ppt/_rels/presentation.xml.rels"] = _HEAD + rels.format(
        rel(1, "slideMaster", "slideMasters/slideMaster1.xml") + rel(2, "theme", "theme/theme1.xml"))
    files["_rels/.rels"] = _HEAD + rels.format(rel(1, "officeDocument", "ppt/presentation.xml"))
    files["[Content_Types].xml"] = (
        _HEAD + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        + "".join(f'<Override PartName="{p}" ContentType="{_CT}{t}"/>' for p, t in over) + "</Types>")
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for name in sorted(files, key=lambda k: (k != "[Content_Types].xml", k)):
            z.writestr(name, files[name].encode("utf-8"))
    return buf.getvalue()
