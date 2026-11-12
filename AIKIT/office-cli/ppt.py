"""ppt.py - 파워포인트 작업

명령:
    read     슬라이드 내용 읽기 (제목, 본문, 표, 노트). 사용자 화면은 움직이지 않음
    info     프레젠테이션 요약 (슬라이드별 제목, 도형/표/그림 수). --all: 열린 프레젠테이션 전체 목록
    write    제목/본문/표/그림/카드 넣기. --file의 마크다운은 제목(#, ##)마다 새 슬라이드
    replace  찾아 바꾸기 (+서식). 표, 그룹 도형 포함. --dry-run: 바꾸지 않고 바뀔 곳만 미리 보기
    export   PDF / PPTX 사본 저장 (원본은 그대로)
    slide    슬라이드 추가/복제/이동/삭제/비우기/목록. --dry-run: 바꾸지 않고 미리 보기
    notes    발표자 노트 읽기 (--slide N --text/--file이면 그 슬라이드 노트 쓰기, --append는 덧붙이기)

--target: active(기본, 지금 보고 있는 프레젠테이션) | 열린 프레젠테이션 이름 | 파일 경로
--slide: 슬라이드 번호 (1부터). write에서 new는 맨 뒤에 새 슬라이드
닫힌 .pptx 읽기는 PowerPoint를 띄우지 않고 python-pptx로 읽습니다.

예:
    python ppt.py write --slide new --title "추진 전략" --text "- 안정성\n- 확장성"
    python ppt.py write --file "검색결과.md"
    python ppt.py write --slide 3 --image chart.png
    python ppt.py write --slide new --title "핵심 가치" --cards "안정성:무중단|확장성:모듈 구조|보안:폐쇄망"
    python ppt.py slide --action move --slide 1 --dest 3
"""
import json
import os
import re
import sys
from typing import Any, Dict, List, Optional, Sequence, Tuple

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _office as off  # noqa: E402
import _runtime as rt  # noqa: E402
import _markdown as md  # noqa: E402

TOOL = "ppt"

PP_LAYOUT = {"title": 1, "text": 2, "title_only": 11, "blank": 12}
PH_BODY = (2, 7)           # ppPlaceholderBody, ppPlaceholderObject
MSO_GROUP = 6
MSO_PLACEHOLDER = 14
MSO_ROUNDED_RECT = 5
PP_BULLET_UNNUMBERED = 1    # ppBulletUnnumbered (글머리표)
PP_BULLET_NUMBERED = 2      # ppBulletNumbered (1. 2. 3.)
MARGIN = 36.0
GAP = 14.0
MIN_BLOCK_HEIGHT = 80.0     # 카드·그림이 들어갈 최소 남은 높이 (모자라면 새 슬라이드)
TABLE_ROWS_PER_SLIDE = 15
MAX_NEW_SLIDES = 30


# ==============================================================================
# 1. 도형 텍스트 도우미
# ==============================================================================

def _has(shape: Any, attr: str) -> bool:
    try:
        return bool(getattr(shape, attr))
    except Exception:
        return False


def iter_text_ranges(shape: Any):
    """도형 안의 모든 TextRange (그룹은 재귀, 표는 셀마다)."""
    try:
        if shape.Type == MSO_GROUP:
            items = shape.GroupItems
            for i in range(1, items.Count + 1):
                yield from iter_text_ranges(items.Item(i))
            return
    except Exception:
        pass
    if _has(shape, "HasTable"):
        tbl = shape.Table
        for r in range(1, tbl.Rows.Count + 1):
            for c in range(1, tbl.Columns.Count + 1):
                try:
                    yield tbl.Cell(r, c).Shape.TextFrame.TextRange
                except Exception:
                    continue
        return
    if _has(shape, "HasTextFrame") and _has(shape.TextFrame, "HasText"):
        yield shape.TextFrame.TextRange


def _shape_lines(shape: Any) -> List[str]:
    if _has(shape, "HasTable"):
        tbl = shape.Table
        rows = []
        for r in range(1, tbl.Rows.Count + 1):
            cells = []
            for c in range(1, tbl.Columns.Count + 1):
                try:
                    cells.append(str(tbl.Cell(r, c).Shape.TextFrame.TextRange.Text).replace("\r", " ").strip())
                except Exception:
                    cells.append("")
            rows.append(" | ".join(cells))
        return rows
    out: List[str] = []
    for tr in iter_text_ranges(shape):
        out += [ln.strip() for ln in str(tr.Text).replace("\x0b", "\r").split("\r") if ln.strip()]
    return out


def _title_shape(slide: Any) -> Optional[Any]:
    try:
        if slide.Shapes.HasTitle:
            return slide.Shapes.Title
    except Exception:
        pass
    return None


def slide_units(slide: Any, idx: int) -> Dict[str, Any]:
    title_shape = _title_shape(slide)
    title = ""
    if title_shape is not None:
        try:
            title = str(title_shape.TextFrame.TextRange.Text).strip()
        except Exception:
            pass
    lines: List[str] = []
    for i in range(1, slide.Shapes.Count + 1):
        sh = slide.Shapes(i)
        if title_shape is not None and sh.Name == title_shape.Name:
            continue
        lines += _shape_lines(sh)
    return {"no": idx, "title": title or (lines[0] if lines else ""), "text": "\n".join(lines), "notes": _notes_text(slide)}


def slide_blocks(slide: Dict[str, Any]) -> List[Tuple[str, Any]]:
    """슬라이드 하나 -> 마크다운 블록 (제목은 # 제목, 글은 줄마다 단락, ' | '로 나눈 줄이 이어지면 표). read --save가 쓰고
    word.py·ppt.py write --file이 같은 모양으로 되읽음 (전에는 줄이 한 단락으로 합쳐지고 표 행이 1x1 표, 2026-11-09 검토 A12)."""
    blocks: List[Tuple[str, Any]] = []
    title = str(slide.get("title") or "").strip()
    lines = [ln for ln in str(slide.get("text") or "").split("\n") if ln.strip()]
    if title:
        blocks.append(("heading", (1, title)))
        if lines and lines[0].strip() == title:
            lines = lines[1:]                     # 제목 개체틀이 없는 슬라이드는 첫 줄이 제목
    table: List[List[str]] = []
    for ln in lines + [""]:
        cells = ln.split(" | ") if " | " in ln else None
        if cells is not None and (not table or len(cells) == len(table[0])):
            table.append(cells)
            continue
        if table:
            blocks.append(("table", table) if len(table) >= 2 else ("paragraph", " | ".join(table[0])))
            table = []
        if cells is not None:
            table.append(cells)
        elif ln.strip():
            blocks.append(("paragraph", ln))
    return blocks


def _notes_body(slide: Any) -> Optional[Any]:
    """발표자 노트 본문 개체틀 (ppPlaceholderBody). 없으면 None."""
    try:
        holders = slide.NotesPage.Shapes.Placeholders
        for i in range(1, int(holders.Count) + 1):
            ph = holders(i)
            try:
                if int(ph.PlaceholderFormat.Type) == 2:
                    return ph
            except Exception:
                continue
    except Exception:
        pass
    return None             # 본문 개체틀이 없으면 머리글/쪽 번호 개체틀에 쓰지 않음


def _notes_text(slide: Any) -> str:
    body = _notes_body(slide)
    if body is None:
        return ""
    try:
        return str(body.TextFrame.TextRange.Text).replace("\r", "\n").replace("\x0b", "\n").strip()
    except Exception:
        return ""


# ==============================================================================
# 2. read / info
# ==============================================================================

def _closed_pptx(target: str) -> Optional[str]:
    if not target or target.strip().lower() == "active":
        return None
    raw = os.path.abspath(os.path.expandvars(target.strip().strip('"')))
    if not os.path.isfile(raw):
        cand = os.path.join(rt.get_data_dir(), target.strip().strip('"'))
        if not os.path.isfile(cand):
            return None
        raw = cand
    if os.path.splitext(raw)[1].lower() not in (".pptx", ".pptm"):
        return None
    import zipfile
    if not zipfile.is_zipfile(raw):
        return None
    try:
        if off.win32com is not None and off.match_open("ppt", raw) is not None:
            return None
    except rt.InputError:
        return None   # 같은 이름의 열린 문서가 여럿: 파일을 읽지 않고 COM 경로에서 [후보] 오류로 안내
    except Exception:
        pass
    return raw


def _from_target(args: Any, source: Dict[str, Any]) -> List[str]:
    """transfer 다음 명령의 --from-target: 보고 있는 문서(active)가 아니면 읽은 문서 (전체 경로가 있으면 그것)."""
    if (args.target or "active").strip().lower() == "active":
        return []
    return ["--from-target", source.get("path") or args.target]


def cmd_read(args) -> Dict[str, Any]:
    import reader
    closed = _closed_pptx(args.target)
    if closed:
        units, _ = reader.parse_pptx(closed)
        slides = [{"no": i, "title": u["heading"], "text": u["text"]} for i, u in enumerate(units, 1)]
        source = {"app": "PowerPoint", "name": os.path.basename(closed), "path": closed, "source": "file(python-pptx)"}
    else:
        with off.open_target("ppt", args.target, readonly=True) as t:       # 읽기: 화면 이동 없음
            prs = t.doc
            numbers = reader.parse_pages(args.slides, prs.Slides.Count, "--slides") if prs.Slides.Count else []
            slides = [slide_units(prs.Slides(n), n) for n in numbers]
            source = t.resolved()
            total = int(prs.Slides.Count)
    if closed:
        total = len(slides)
        numbers = reader.parse_pages(args.slides, total, "--slides") if total else []
        slides = [slides[n - 1] for n in numbers]
    result: Dict[str, Any] = {"target_resolved": source, "slides_total": total}
    if args.save:
        out = rt.resolve_output_path(None if args.save is True else args.save, os.path.splitext(source["name"])[0], ".md",
                                     near=source.get("path"))
        out = rt.write_text(out, md.document_markdown([b for s in slides for b in slide_blocks(s)]))
        # 열린 문서를 바꾸는 이어 하기는 confirm_commands (1.4.75)
        result.update(output_file=out, confirm_commands=[
            rt.next_command("word.py", ["write", "--file", out], "저장한 내용을 열린 워드 문서에 넣기 (사용자가 원할 때만)"),
            rt.next_command("office.py", ["transfer", "--from-app", "ppt", "--what", "text", "--to-app", "excel"]
                            + _from_target(args, source),
                            "슬라이드 내용을 엑셀 표로 옮기기 (사용자가 원할 때만)")])
        return rt.ok_file(TOOL, "read", f"{source['name']}: 슬라이드 {len(slides)}장을 저장했습니다 ({out})", out,
                          "VERIFIED: 저장한 파일을 다시 열어 확인했습니다. " + rt.INTO_OTHER_DOC,
                          **result)
    shown: List[Dict[str, Any]] = []
    if not any(str(s["title"]).strip() or str(s["text"]).strip() for s in slides):
        # 빈 결과 명시 (ACI 설계 권장 2번, 1.4.81): 슬라이드가 없거나 글이 하나도 없음
        what = "슬라이드가 없습니다" if not slides else f"슬라이드 {len(slides)}장에 글이 없습니다 (그림·표만 있거나 빈 슬라이드)"
        result.update(slides=slides[:20], truncated=False, count=0)
        return rt.ok(TOOL, "read", f"0건: {source['name']}: {what}",
                     f"VERIFIED: 0건. {what} (찾은 항목 없음). 사용자에게 그대로 알리세요.", **result)
    used = 0
    text_cut = 0
    for s in slides:
        size = len(s["text"]) + len(s["title"]) + 20
        if shown and used + size > rt.DEFAULT_MAX_CHARS:
            break
        text, cut = rt.clamp_text(s["text"], rt.DEFAULT_MAX_CHARS)
        item = dict(s, text=text, title=rt.clamp_text(s["title"], 500)[0])
        if cut:                                   # 글이 아주 많은 슬라이드: 앞부분만 (전체는 --save 파일)
            item.update(text_truncated=True, text_chars=len(s["text"]))
            text_cut += 1
        shown.append(item)
        used += size
    truncated = len(shown) < len(slides) or bool(text_cut)
    result.update(slides=shown, truncated=truncated)
    nexts: List[Dict[str, str]] = []
    if len(shown) < len(slides):
        rest = _compress_numbers([s["no"] for s in slides[len(shown):]])     # 고른 슬라이드 중 남은 것만 (1,5,9 -> 5,9)
        nexts.append(rt.next_command("ppt.py", ["read", "--target", args.target, "--slides", rest],
                                     "나머지 슬라이드 이어서 읽기"))
    if text_cut:
        result["truncated_reason"] = f"글이 {rt.DEFAULT_MAX_CHARS}자보다 많은 슬라이드 {text_cut}장은 앞부분만 표시 (text_truncated, 전체는 --save)"
        save = ["read", "--target", args.target] + (["--slides", args.slides] if args.slides else []) + ["--save"]
        nexts.append(rt.next_command("ppt.py", save, "잘린 슬라이드까지 전체를 md 파일로 저장"))
    if nexts:
        result["next_commands"] = nexts
    return rt.ok(TOOL, "read", f"{source['name']}: 슬라이드 {len(shown)}장 표시 (전체 {total}장)",
                 "VERIFIED: 읽기 완료. 일부만 표시되었으면 next_commands로 이어서 읽으세요." if truncated else "", **result)


def _compress_numbers(nums: List[int]) -> str:
    """[2,3,4,9] -> '2-4,9' (--slides에 그대로 넣을 수 있는 모양)"""
    runs: List[List[int]] = []
    for n in sorted(set(nums)):
        if runs and n == runs[-1][1] + 1:
            runs[-1][1] = n
        else:
            runs.append([n, n])
    return ",".join(str(a) if a == b else f"{a}-{b}" for a, b in runs)


def _shape_kind(sh: Any) -> str:
    if _has(sh, "HasTable"):
        return "table"
    if _has(sh, "HasChart"):
        return "chart"
    try:
        t = int(sh.Type)
        if t in (13, 11):
            return "picture"
        if t == MSO_GROUP:
            return "group"
        if t == MSO_PLACEHOLDER:
            ph = sh.PlaceholderFormat
            if ph.Type in (1, 3):
                return "title"
            if getattr(ph, "ContainedType", None) == 13:
                return "picture"
            return "text"
        if t == 17:
            return "text"
        if t == 1:
            return "shape"
    except Exception:
        pass
    return "other"


def _list_kinds(tr: Any) -> List[str]:
    """단락별 목록 종류: none | bullet | numbered."""
    kinds = []
    try:
        n = int(tr.Paragraphs().Count)
    except Exception:
        return kinds
    for i in range(1, n + 1):
        try:
            b = tr.Paragraphs(i).ParagraphFormat.Bullet
            kinds.append("none" if not int(b.Visible) else ("numbered" if int(b.Type) == PP_BULLET_NUMBERED else "bullet"))
        except Exception:
            kinds.append("none")
    return kinds


def slide_detail(slide: Any, width: float, height: float) -> Tuple[List[Dict[str, Any]], List[str]]:
    """도형별 종류·위치·크기·글자·목록 종류와, 배치 문제 경고(슬라이드 밖, 글자 넘침, 겹침)."""
    shapes: List[Dict[str, Any]] = []
    warnings: List[str] = []
    for k in range(1, slide.Shapes.Count + 1):
        sh = slide.Shapes(k)
        kind = _shape_kind(sh)
        item: Dict[str, Any] = {"name": str(sh.Name), "kind": kind, **off.shape_box(sh)}
        if _has(sh, "HasTextFrame") and _has(sh.TextFrame, "HasText"):
            tr = sh.TextFrame.TextRange
            item["text"] = str(tr.Text).replace("\r", " / ")[:60]
            try:
                item["font_size"] = float(tr.Font.Size)
            except Exception:
                pass
            kinds = _list_kinds(tr)
            if any(x != "none" for x in kinds):
                item["lists"] = kinds
            try:
                if float(tr.BoundHeight) > float(sh.Height) + 2:
                    item["text_overflow"] = True
                    warnings.append(f"'{sh.Name}' 글자가 도형 높이를 넘칩니다")
            except Exception:
                pass
        elif kind in ("title", "text"):
            continue                                        # 빈 자리표시자는 화면에 보이지 않음
        if off.outside_area(item, width, height):
            item["outside"] = True
            warnings.append(f"'{sh.Name}'이(가) 슬라이드 밖으로 나갑니다")
        shapes.append(item)
    for i, a in enumerate(shapes):
        for b in shapes[i + 1:]:
            r = off.overlap_ratio(a, b)
            if r >= 0.1:
                warnings.append(f"'{a['name']}'와(과) '{b['name']}'이(가) 겹칩니다 ({int(r * 100)}%)")
    return shapes, warnings


def cmd_info(args) -> Dict[str, Any]:
    if args.all:
        items = [{"name": d["name"], "path": d["path"], "saved": d["saved"], "active": d["active"]}
                 for d in off.list_open("ppt")]
        if not items:
            return off.no_open_docs_result(TOOL, "info", "0건: 열린 프레젠테이션이 없습니다",
                         "VERIFIED: 0건. 열린 PowerPoint 프레젠테이션이 없습니다. 사용자에게 그대로 알리세요.",
                         presentations=items, count=0)
        return rt.ok(TOOL, "info", f"열린 프레젠테이션 {len(items)}개", presentations=items, count=len(items))
    with off.open_target("ppt", args.target, readonly=True) as t:
        prs = t.doc
        slides = []
        width, height = float(prs.PageSetup.SlideWidth), float(prs.PageSetup.SlideHeight)
        import reader
        wanted = set(reader.parse_pages(args.slides, prs.Slides.Count, "--slides")) if args.slides and prs.Slides.Count else None
        warnings_total = 0
        for i in range(1, prs.Slides.Count + 1):
            if wanted is not None and i not in wanted:
                continue
            s = prs.Slides(i)
            tables = pictures = 0
            for k in range(1, s.Shapes.Count + 1):
                sh = s.Shapes(k)
                if _has(sh, "HasTable"):
                    tables += 1
                try:
                    if sh.Type in (13, 11) or (sh.Type == MSO_PLACEHOLDER and sh.PlaceholderFormat.ContainedType == 13):
                        pictures += 1
                except Exception:
                    pass
            u = slide_units(s, i)
            entry = {"no": i, "title": u["title"][:60], "shapes": int(s.Shapes.Count), "tables": tables,
                     "pictures": pictures, "chars": len(u["text"]) + len(u["title"])}
            if args.detail:
                entry["shape_detail"], entry["warnings"] = slide_detail(s, width, height)
                warnings_total += len(entry["warnings"])
            slides.append(entry)
        size = {"width": width, "height": height}
        result = {"target_resolved": t.resolved(), "slides_total": int(prs.Slides.Count), "slides": slides,
                  "slide_size": size, "saved": off.is_saved(prs)}
    preview = f"{result['target_resolved']['name']}: 슬라이드 {result['slides_total']}장"
    if args.detail:
        result["layout_warnings"] = warnings_total
        preview += f", 배치 경고 {warnings_total}건" if warnings_total else ", 배치 경고 없음"
    return rt.ok(TOOL, "info", preview, **result)


# ==============================================================================
# 3. write
# ==============================================================================

PP_FOOTER_PLACEHOLDERS = {13, 14, 15, 16}   # ppPlaceholderSlideNumber, Header, Footer, Date


class SlideWriter:
    """슬라이드 하나에 요소를 위에서 아래로 쌓습니다 (슬라이드 경계 안에서)."""

    def __init__(self, prs: Any, slide: Any, theme: Dict[str, Any]):
        self.prs = prs
        self.slide = slide
        self.theme = theme
        self.W = float(prs.PageSetup.SlideWidth)
        self.H = float(prs.PageSetup.SlideHeight)
        self.created: List[str] = []
        self.shapes: List[Any] = []          # 이번에 넣거나 채운 도형 (다시 읽어 슬라이드 안에 있는지 확인)

    def _empty_placeholder(self, sh: Any) -> bool:
        try:
            return sh.Type == MSO_PLACEHOLDER and _has(sh, "HasTextFrame") and not _has(sh.TextFrame, "HasText")
        except Exception:
            return False

    def _footer_placeholder(self, sh: Any) -> bool:
        """바닥글·날짜·슬라이드 번호·머리글 개체틀: 슬라이드 바닥에 있어 그 아래를 기준으로 쌓으면 새 글상자가 슬라이드 밖에 놓임
        (2026-11-09 검토 C8)."""
        try:
            return sh.Type == MSO_PLACEHOLDER and int(sh.PlaceholderFormat.Type) in PP_FOOTER_PLACEHOLDERS
        except Exception:
            return False

    def next_top(self) -> float:
        bottom = MARGIN
        for i in range(1, self.slide.Shapes.Count + 1):
            sh = self.slide.Shapes(i)
            if self._empty_placeholder(sh) or self._footer_placeholder(sh):
                continue
            try:
                bottom = max(bottom, float(sh.Top) + float(sh.Height))
            except Exception:
                pass
        return bottom + GAP

    def space_left(self) -> float:
        return self.H - MARGIN - self.next_top()

    def title(self, text: str) -> None:
        ts = _title_shape(self.slide)
        if ts is not None:
            ts.TextFrame.TextRange.Text = text
        else:
            box = self.slide.Shapes.AddTextbox(1, MARGIN, MARGIN * 0.6, self.W - 2 * MARGIN, 50)
            tr = box.TextFrame.TextRange
            tr.Text = text
            tr.Font.Size = 28
            tr.Font.Bold = True
            tr.Font.Color.RGB = off.parse_color(self.theme["primary"])
            self.shapes.append(box)
        self.created.append("제목")

    def _body_placeholder(self) -> Optional[Any]:
        for i in range(1, self.slide.Shapes.Count + 1):
            sh = self.slide.Shapes(i)
            try:
                if sh.Type == MSO_PLACEHOLDER and sh.PlaceholderFormat.Type in PH_BODY and not _has(sh.TextFrame, "HasText"):
                    return sh
            except Exception:
                continue
        return None

    def body(self, lines: List[Tuple[str, str]]) -> None:
        """lines: [(텍스트, 목록 종류)]. 목록 종류는 ""(일반), "bullet", "numbered"."""
        text = "\r".join(t for t, _ in lines)
        ph = self._body_placeholder()
        if ph is not None:
            shape = ph
            shape.TextFrame.TextRange.Text = text
        else:
            top = self.next_top()
            height = max(40.0, min(self.H - MARGIN - top, 28.0 * len(lines) + 10))
            shape = self.slide.Shapes.AddTextbox(1, MARGIN, top, self.W - 2 * MARGIN, height)
            shape.TextFrame.WordWrap = -1
            tr = shape.TextFrame.TextRange
            tr.Text = text
            tr.Font.Size = 18 if len(lines) <= 6 else 14
            self.shapes.append(shape)
        tr = shape.TextFrame.TextRange
        for i, (_, kind) in enumerate(lines, 1):
            try:
                bullet = tr.Paragraphs(i).ParagraphFormat.Bullet
                bullet.Visible = -1 if kind else 0
                if kind:
                    # Visible만 켜면 글상자에 따라 번호 목록이 될 수 있어 종류를 명시합니다.
                    bullet.Type = PP_BULLET_NUMBERED if kind == "numbered" else PP_BULLET_UNNUMBERED
            except Exception:
                pass
        self.created.append(f"본문 {len(lines)}줄")

    def table(self, rows: List[List[Any]], markdown: bool = False) -> None:
        bold_cells: set = set()
        if markdown:
            rows, cells = md.table_cells(rows)
            bold_cells = {(i + 1, j + 1) for i, j in cells}
        nrows, ncols = len(rows), max(len(r) for r in rows)
        top = self.next_top()
        height = min(self.H - MARGIN - top, 26.0 * nrows)
        shape = self.slide.Shapes.AddTable(nrows, ncols, MARGIN, top, self.W - 2 * MARGIN, max(height, 26.0))
        self.shapes.append(shape)
        tbl = shape.Table
        size = 14 if nrows <= 8 else 11
        header_bg = off.parse_color(self.theme["header_bg"])
        header_fg = off.parse_color(self.theme["header_fg"])
        for r, row in enumerate(rows, 1):
            for c in range(1, ncols + 1):
                value = row[c - 1] if c - 1 < len(row) else ""
                cell = tbl.Cell(r, c).Shape
                tr = cell.TextFrame.TextRange
                tr.Text = "" if value is None else str(value)
                tr.Font.Size = size
                if r == 1:
                    tr.Font.Bold = True
                    tr.Font.Color.RGB = header_fg
                    cell.Fill.ForeColor.RGB = header_bg
                elif (r, c) in bold_cells:
                    tr.Font.Bold = True
        self.created.append(f"표 {nrows}x{ncols}")

    def image(self, path: str, width: Optional[float]) -> None:
        top = self.next_top()
        pic = self.slide.Shapes.AddPicture(FileName=path, LinkToFile=0, SaveWithDocument=-1, Left=MARGIN, Top=top)
        max_w = width or (self.W - 2 * MARGIN)
        max_h = max(60.0, self.H - MARGIN - top)
        w0, h0 = float(pic.Width), float(pic.Height)
        ratio = min(max_w / w0, max_h / h0)            # 남은 공간 안에 들어가는 최대 배율
        if not width:
            ratio = min(ratio, 1.0)                    # 너비를 지정하지 않았으면 원본보다 키우지 않음
        pic.LockAspectRatio = -1
        pic.Width = w0 * ratio
        pic.Height = h0 * ratio
        pic.Left = (self.W - float(pic.Width)) / 2
        self.shapes.append(pic)
        self.created.append(f"그림 {os.path.basename(path)}")

    def cards(self, cards: List[Tuple[str, str]]) -> None:
        """카드 높이는 80~220pt (남은 공간에 맞춤). 80pt도 남지 않으면 부르는 쪽이 새 슬라이드에 넣습니다."""
        n = len(cards)
        top = self.next_top()
        height = max(80.0, min(self.H - MARGIN - top, 220.0))
        total = self.W - 2 * MARGIN
        w = (total - GAP * (n - 1)) / n
        primary = off.parse_color(self.theme["primary"])
        light = off.parse_color(self.theme["light"])
        for i, (head, desc) in enumerate(cards):
            box = self.slide.Shapes.AddShape(MSO_ROUNDED_RECT, MARGIN + i * (w + GAP), top, w, height)
            self.shapes.append(box)
            box.Fill.ForeColor.RGB = light
            box.Line.ForeColor.RGB = primary
            box.Line.Weight = 1.5
            tr = box.TextFrame.TextRange
            tr.Text = head + ("\r" + desc if desc else "")
            tr.Font.Color.RGB = 0x333333
            tr.Font.Size = 14
            first = tr.Paragraphs(1)
            first.Font.Bold = True
            first.Font.Size = 18
            first.Font.Color.RGB = primary
        self.created.append(f"카드 {n}개")


def _new_slide(prs: Any, index: int, with_title: bool) -> Any:
    return prs.Slides.Add(index, PP_LAYOUT["title_only"] if with_title else PP_LAYOUT["blank"])


def _md_sections(text: str) -> List[Dict[str, Any]]:
    """마크다운을 슬라이드 단위로 나눕니다: '#'/'##' 제목마다 새 슬라이드, 그 밖의 블록은 본문/표.
    '---'(구분선)은 넣지 않고 슬라이드도 나누지 않음, 맨 앞 YAML 머리말도 넣지 않음 (cmd_write가 warnings에 남김)."""
    sections: List[Dict[str, Any]] = []
    cur: Optional[Dict[str, Any]] = None
    for kind, payload in md.parse_markdown_blocks(text):
        if kind in ("break", "front_matter"):
            continue
        if kind == "heading" and payload[0] <= 2:
            cur = {"title": md.plain_heading(payload[1]), "lines": [], "tables": []}
            sections.append(cur)
            continue
        if cur is None:
            cur = {"title": "", "lines": [], "tables": []}
            sections.append(cur)
        if kind == "heading":
            cur["lines"].append((md.plain_heading(payload[1]), ""))   # 3단계 이하 제목: 1·2단계 제목과 같은 규칙 (** __를 뺀 글)
        elif kind == "paragraph":
            cur["lines"].append((md.split_bold(payload)[0], ""))
        elif kind == "code":
            cur["lines"].append((payload, ""))          # 코드 줄은 글자 그대로 (** __를 굵게로 읽지 않음)
        elif kind in ("bullets", "numbered"):
            list_kind = "numbered" if kind == "numbered" else "bullet"
            cur["lines"] += [(md.split_bold(x)[0], list_kind) for x in payload]
        elif kind == "table":
            cur["tables"].append(payload)
    return [s for s in sections if s["title"] or s["lines"] or s["tables"]]


def _text_lines(text: str) -> List[Tuple[str, str]]:
    """--text: 줄마다 한 줄 (-, 1. 목록 기호와 ** __ 굵게 표시를 해석). 펜스(```, ~~~) 안의 코드 줄은 .md와 같이 글자 그대로."""
    lines = []
    fence = ""
    numbered_as_text = False
    for ln in rt.unescape_newlines(text).splitlines():
        s = ln.strip()
        if fence:
            if s.startswith(fence) and not s.lstrip(fence[0]):
                fence = ""                          # 닫는 펜스 (닫지 않으면 끝까지 코드)
            elif s:
                lines.append((ln.rstrip(), ""))
            continue
        if not s:
            continue
        m = md._FENCE.match(s)
        if m:
            fence = m.group(1)                      # 펜스 줄은 넣지 않음
            continue
        m = re.match(r"^[-*+]\s+(.*)$", s)
        n = None if m else md.MD_NUMBERED.match(s)
        if not n:
            numbered_as_text = False
        elif not numbered_as_text and not (lines and lines[-1][1] == "numbered"):
            # 번호 목록의 첫 줄: 1이 아닌 번호로 시작하면 자동 번호가 1부터 다시 매기므로 이 목록은 번호를 글로 둔 줄로 넣음
            numbered_as_text = int(n.group(1)) != 1
        if n and numbered_as_text:
            lines.append((md.split_bold(s)[0], ""))
            continue
        body = m.group(1) if m else (s[n.end():] if n else s)
        lines.append((md.split_bold(body)[0], "bullet" if m else ("numbered" if n else "")))
    return lines


def _parse_cards(raw: str) -> List[Tuple[str, str]]:
    """'제목:내용; 제목:내용' (카드 사이는 ;, 저장소 구분자 규칙). 예전 구분자 '|'만 쓴 것은 ;로 고친 명령을 [추천] (1.4.79)."""
    if "|" in raw and ";" not in raw:
        fixed = "; ".join(p.strip() for p in raw.split("|") if p.strip())
        raise rt.InputError(f"--cards의 카드 사이는 ; 로 나눕니다 ('|' 아님). [추천] {rt.retry_command({'--cards': fixed})}")
    cards = []
    for part in raw.split(";"):
        if not part.strip():
            continue
        head, _, desc = part.partition(":")
        cards.append((head.strip(), desc.strip()))
    if not 1 <= len(cards) <= 5:
        raise rt.InputError("--cards는 1~5개를 '제목:내용; 제목:내용' 형식으로 주세요.")
    return cards


def _resolve_slide(prs: Any, value: Optional[str], default_new: bool) -> Tuple[Optional[int], bool]:
    """(슬라이드 번호, 새로 만들지). value: None | 'new' | 숫자"""
    count = int(prs.Slides.Count)
    if value is None:
        return (None, True) if default_new or count == 0 else (count, False)
    if str(value).strip().lower() == "new":
        return None, True
    try:
        n = int(value)
    except ValueError:
        raise rt.InputError(f"--slide는 번호 또는 new여야 합니다: {value}", no_fallback=True)
    if not 1 <= n <= count:
        raise rt.InputError(f"슬라이드 {n}번이 없습니다 (전체 {count}장). 새 슬라이드는 --slide new", no_fallback=True)
    return n, False


def _new_text_lines(text: str) -> List[str]:
    """새 글의 줄: '- '·'* '로 시작하면 그 표시는 빼고 (도형의 글머리 서식을 그대로 씀), **굵게**는 글만."""
    out = []
    for line in rt.unescape_newlines(text).replace("\r\n", "\n").split("\n"):
        t = line.rstrip()
        m = re.match(r"^\s*[-*]\s+(.*)$", t)
        out.append(md.split_bold(m.group(1) if m else t)[0])
    while out and not out[-1].strip():
        out.pop()
    return out


def _find_shape(slide: Any, slide_no: int, name: str) -> Any:
    """슬라이드의 도형을 이름으로 (대소문자·앞뒤 공백 무시). 없으면 글이 있는 도형 이름을 [후보]로."""
    want = name.strip().casefold()
    names = []
    for k in range(1, slide.Shapes.Count + 1):
        sh = slide.Shapes(k)
        if str(sh.Name).strip().casefold() == want:
            return sh
        if _has(sh, "HasTextFrame"):
            names.append(str(sh.Name))
    raise rt.ambiguous(f"슬라이드 {slide_no}에 '{name}' 도형이 없습니다. 도형 이름은 ppt.py info --detail --slides {slide_no}의 "
                       "shape_detail name입니다.", names, log_message="도형 이름을 찾지 못함", replace=name)


TABLE_ROWS_SHOW = 100        # table 읽기에서 JSON에 담는 행 수
_CELL_KEY = re.compile(r"^R(\d+)C(\d+)$", re.I)


def _cell_text(tbl: Any, r: int, c: int) -> str:
    try:
        return str(tbl.Cell(r, c).Shape.TextFrame.TextRange.Text).replace("\r", "\n").strip()
    except Exception:
        return ""


def _pick_table(slide: Any, slide_no: int, name: Optional[str]) -> Any:
    """슬라이드의 표 도형: --shape가 있으면 그 이름, 없으면 슬라이드에 하나뿐인 표 (여럿이면 [후보])."""
    if name:
        sh = _find_shape(slide, slide_no, name)
        if not _has(sh, "HasTable"):
            raise rt.InputError(f"슬라이드 {slide_no}의 '{sh.Name}'은(는) 표가 아닙니다.")
        return sh
    tables = [slide.Shapes(k) for k in range(1, slide.Shapes.Count + 1) if _has(slide.Shapes(k), "HasTable")]
    if not tables:
        raise rt.InputError(f"슬라이드 {slide_no}에 표가 없습니다 (count 0). 표가 있는 슬라이드는 ppt.py info의 tables로 확인합니다.")
    if len(tables) > 1:
        names = [str(t.Name) for t in tables]
        raise rt.UserDecisionError(f"슬라이드 {slide_no}에 표가 {len(tables)}개 있습니다.", "어느 표인지 [후보]를 사용자에게 보여 주고 고른 번호의 "
                                   "confirm_commands를 1번 실행하세요.", candidates=names,
                                   commands=[rt.retry_command({"--shape": n}) for n in names])
    return tables[0]


def _parse_table_cells(spec: Optional[str], path: Optional[str]) -> List[Tuple[int, int, str]]:
    """--cells "R2C3=값; R3C1=값"(칸 여러 개는 ;) 또는 --file(.csv, 열: 칸, 값) -> [(행, 열, 값)]."""
    pairs: List[Tuple[str, str]] = []
    if path:
        src = rt.resolve_input_path(path, "칸 값 파일", exts=(".csv", ".txt"))
        import csv as _csv
        for row in _csv.reader(rt.read_text_file(src).splitlines()):
            if len(row) >= 2 and row[0].strip():
                pairs.append((row[0].strip(), row[1]))
        if pairs and not _CELL_KEY.match(pairs[0][0]):
            pairs = pairs[1:]                                   # 머리글 한 줄 (예: 칸,값)
    elif spec:
        if "," in spec and ";" not in spec and "=" in spec and spec.count("=") > 1:
            raise rt.InputError(f"--cells의 칸 여러 개는 ';'로 나눕니다: '{spec}'. [추천] {rt.retry_command({'--cells': spec.replace(',', ';')})}",
                                no_fallback=True)
        for part in spec.split(";"):
            if part.strip():
                k, sep, v = part.partition("=")
                if not sep:
                    raise rt.InputError(f"--cells의 '{part.strip()}'에 '='가 없습니다. 형식: \"R2C3=값; R3C1=값\"", no_fallback=True)
                pairs.append((k.strip(), v.strip()))
    out = []
    for k, v in pairs:
        m = _CELL_KEY.match(k)
        if not m:
            raise rt.InputError(f"칸 '{k}'은(는) 'R행C열'(예: R2C3) 형식이 아닙니다.", no_fallback=True)
        out.append((int(m.group(1)), int(m.group(2)), rt.unescape_newlines(v)))
    if not out:
        raise rt.InputError("쓸 칸이 없습니다. --cells \"R2C3=값\" 또는 --file <칸,값 CSV>를 주세요.", no_fallback=True)
    return out


def cmd_table(args) -> Dict[str, Any]:
    """슬라이드에 이미 있는 표 읽기(--cells·--file 없음)와 칸 쓰기 (1.6.7 기반 기능: 슬라이드 표의 수치 갱신).
    글이 있는 칸은 --overwrite가 있을 때만 바꿈 (R4), 쓴 뒤 다시 읽어 확인."""
    writing = bool(args.cells or args.file)
    cells = _parse_table_cells(args.cells, args.file) if writing else []
    with off.open_target("ppt", args.target, readonly=not writing or bool(args.dry_run), focus=not args.no_focus, tool=TOOL,
                         command="table") as t:
        prs = t.doc
        if args.slide < 1 or args.slide > int(prs.Slides.Count):
            raise rt.InputError(f"슬라이드 {args.slide}이(가) 없습니다 (1~{prs.Slides.Count}).")
        sh = _pick_table(prs.Slides(args.slide), args.slide, args.shape)
        tbl = sh.Table
        nr, nc = int(tbl.Rows.Count), int(tbl.Columns.Count)
        info: Dict[str, Any] = {"slide": args.slide, "shape": str(sh.Name), "nrows": nr, "ncols": nc}
        if not writing:
            rows = [[_cell_text(tbl, r, c) for c in range(1, nc + 1)] for r in range(1, nr + 1)]
            resolved = t.resolved()
            shown = rows[:TABLE_ROWS_SHOW]
            return rt.ok(TOOL, "table", f"슬라이드 {args.slide} 표 '{sh.Name}': {nr}행 x {nc}열", "VERIFIED: 표를 읽었습니다. 칸은 R행C열(예: R2C3)로 씁니다.",
                         rows=shown, truncated=len(rows) > len(shown), target_resolved=resolved, **info)
        bad = [f"R{r}C{c}" for r, c, _ in cells if not (1 <= r <= nr and 1 <= c <= nc)]
        if bad:
            raise rt.InputError(f"표 '{sh.Name}'({nr}행 x {nc}열) 밖의 칸: {', '.join(bad[:5])}.", no_fallback=True)
        olds = {(r, c): _cell_text(tbl, r, c) for r, c, _ in cells}
        filled = [f"R{r}C{c}('{_preview_text(olds[(r, c)], 20)}')" for r, c, v in cells if olds[(r, c)] and olds[(r, c)] != v.strip()]
        if filled and not args.overwrite and not args.dry_run:
            raise rt.UserDecisionError(f"표 '{sh.Name}'에서 글이 있는 칸 {len(filled)}개를 바꾸게 됩니다: {', '.join(filled[:5])}.",
                                       "바꾸지 않았습니다. 사용자에게 바꿀지 묻고, '예'이면 confirm_commands[0]을 1번 실행하세요.",
                                       candidates=["글이 있는 칸도 바꾸기"], commands=[rt.retry_command({"--overwrite": True})], destructive=True)
        planned = [f"R{r}C{c}: '{_preview_text(olds[(r, c)], 20)}' -> '{_preview_text(v, 20)}'" for r, c, v in cells]
        if args.dry_run:
            return rt.dry_run_result(TOOL, "table", f"미리보기: 표 '{sh.Name}' 칸 {len(cells)}개 (바꾸지 않음)", len(cells), planned,
                                     rt.dry_run_command(TOOL, f"표 '{sh.Name}' 칸 {len(cells)}개 쓰기", destructive=bool(filled), overwrite=bool(filled)),
                                     target_resolved=t.resolved(), **info)
        for r, c, v in cells:
            tbl.Cell(r, c).Shape.TextFrame.TextRange.Text = v
        wrong = [f"R{r}C{c}" for r, c, v in cells if _cell_text(tbl, r, c) != v.replace("\r", "\n").strip()]
        t.set_focus(slide=args.slide)
        resolved = t.resolved()
    preview = f"슬라이드 {args.slide} 표 '{info['shape']}'의 칸 {len(cells)}개를 썼습니다"
    if wrong:
        return rt.ok_unverified(TOOL, "table", preview, [f"다시 읽은 글이 다른 칸: {', '.join(wrong[:5])}"], target_resolved=resolved, **info)
    return rt.ok(TOOL, "table", preview, "VERIFIED: 쓴 칸을 다시 읽어 확인했습니다. 되돌리려면 office.py undo.", verified=True,
                 count=len(cells), target_resolved=resolved, **info)


def cmd_shape(args) -> Dict[str, Any]:
    """기존 도형 옮기기·크기 (pt): --left --top --width --height 중 준 것만 바꿈. 슬라이드 밖으로 나가면 바꾸지 않고
    슬라이드 안으로 맞춘 값을 [추천]. 그림처럼 비율이 잠긴 도형은 너비만 주면 높이가 따라감."""
    want = {k: getattr(args, k) for k in ("left", "top", "width", "height") if getattr(args, k) is not None}
    if not want:
        raise rt.InputError("바꿀 위치·크기가 없습니다. --left, --top, --width, --height(pt) 중 하나 이상을 주세요. "
                            f"지금 위치는 ppt.py info --detail --slides {args.slide}의 shape_detail입니다.")
    if any(v < 0 for k, v in want.items()) or any(want.get(k, 1) <= 0 for k in ("width", "height")):
        raise rt.InputError("위치는 0 이상, 너비·높이는 0보다 커야 합니다 (단위 pt).")
    with off.open_target("ppt", args.target, readonly=bool(args.dry_run), focus=not args.no_focus, tool=TOOL, command="shape") as t:
        prs = t.doc
        if args.slide < 1 or args.slide > int(prs.Slides.Count):
            raise rt.InputError(f"슬라이드 {args.slide}이(가) 없습니다 (1~{prs.Slides.Count}).")
        sh = _find_shape(prs.Slides(args.slide), args.slide, args.shape)
        sw, shh = float(prs.PageSetup.SlideWidth), float(prs.PageSetup.SlideHeight)
        before = off.shape_box(sh)
        box = {**before, **{k: float(v) for k, v in want.items()}}
        locked = bool(getattr(sh, "LockAspectRatio", 0))
        if locked and ("width" in want) != ("height" in want) and before["width"] and before["height"]:
            ratio = before["height"] / before["width"]
            if "width" in want:
                box["height"] = round(box["width"] * ratio, 1)
            else:
                box["width"] = round(box["height"] / ratio, 1)
        if off.outside_area(box, sw, shh):
            fit = dict(box)
            fit["width"], fit["height"] = min(fit["width"], sw), min(fit["height"], shh)
            fit["left"] = max(0.0, min(fit["left"], sw - fit["width"]))
            fit["top"] = max(0.0, min(fit["top"], shh - fit["height"]))
            rec = {f"--{k}": round(fit[k], 1) for k in want}
            raise rt.InputError(f"'{sh.Name}'이(가) 슬라이드({sw:g} x {shh:g}pt) 밖으로 나갑니다 (바꾸지 않음). "
                                f"[추천] {rt.retry_command(rec)}")
        info: Dict[str, Any] = {"slide": args.slide, "shape": str(sh.Name), "before": before, "after": {k: round(v, 1) for k, v in box.items()}}
        desc = ", ".join(f"{k} {before[k]:g}->{box[k]:g}" for k in ("left", "top", "width", "height") if abs(before[k] - box[k]) > 0.05)
        if not desc:
            return rt.ok(TOOL, "shape", f"슬라이드 {args.slide} '{sh.Name}': 이미 그 위치·크기라 바꾸지 않았습니다",
                         "VERIFIED: 바뀐 것 없음.", verified=True, target_resolved=t.resolved(), unchanged=True, **info)
        if args.dry_run:
            return rt.dry_run_result(TOOL, "shape", f"미리보기: 슬라이드 {args.slide} '{sh.Name}' {desc} (바꾸지 않음)", 1,
                                     [f"'{sh.Name}': {desc}"], rt.dry_run_command(TOOL, f"'{sh.Name}' 옮기기·크기"),
                                     target_resolved=t.resolved(), **info)
        both = locked and "width" in want and "height" in want
        if both:
            sh.LockAspectRatio = 0                           # 너비·높이를 둘 다 주면 비율 잠금을 잠시 풂 (잠긴 채면 서로를 다시 바꿈)
        try:
            for k in ("width", "height", "left", "top"):      # 크기 먼저 (비율 잠금 도형이 위치를 다시 잡지 않게)
                if k in want:
                    setattr(sh, k.capitalize(), float(want[k]))
        finally:
            if both:
                sh.LockAspectRatio = -1
        t.set_focus(slide=args.slide)
        got = off.shape_box(sh)
        resolved = t.resolved()
    bad = [f"{k} {got[k]:g}(준 값 {box[k]:g})" for k in want if abs(got[k] - box[k]) > 1.0]
    preview = f"슬라이드 {args.slide} '{info['shape']}': {desc}"
    info["after"] = got
    if bad:
        return rt.ok_unverified(TOOL, "shape", preview, ["다시 읽은 위치·크기가 다름: " + ", ".join(bad)], target_resolved=resolved, **info)
    return rt.ok(TOOL, "shape", preview, "VERIFIED: 위치·크기를 다시 읽어 확인했습니다. 잘못 옮겼으면 office.py undo로 되돌립니다.",
                 verified=True, target_resolved=resolved, **info)


def _write_shape(args) -> Dict[str, Any]:
    """write --slide N --shape <이름>: 그 도형(글상자·본문 개체 틀)의 글을 새 글로 바꿈. 새 글상자를 만들지 않음.
    글이 있던 도형은 --overwrite가 있을 때만 (사용자 확인). 도형의 글꼴·글머리 서식은 그대로."""
    if any([args.title, args.image, args.cards]):
        raise rt.InputError("--shape는 그 도형의 글만 바꿉니다. --title·--image·--cards는 함께 쓰지 않습니다 (제목 개체 틀도 --shape로).")
    if args.text is None and not args.file:
        raise rt.InputError("--shape에 넣을 글이 없습니다. --text \"<새 글>\" 또는 --file \"<글.md>\"를 주세요.")
    if not args.slide or not str(args.slide).isdigit():
        raise rt.InputError("--shape에는 기존 슬라이드 번호가 필요합니다 (--slide 3). 도형 이름은 ppt.py info --detail --slides <번호>로 봅니다.")
    text = rt.read_text_file(rt.resolve_input_path(args.file, "입력 파일")) if args.file else args.text
    lines = _new_text_lines(text or "")
    slide_no = int(args.slide)
    with off.open_target("ppt", args.target, readonly=bool(args.dry_run), focus=not args.no_focus,
                         tool=TOOL, command="write") as t:
        prs = t.doc
        if slide_no < 1 or slide_no > int(prs.Slides.Count):
            raise rt.InputError(f"슬라이드 {slide_no}이(가) 없습니다 (1~{prs.Slides.Count}).")
        sh = _find_shape(prs.Slides(slide_no), slide_no, args.shape)
        if not _has(sh, "HasTextFrame"):
            raise rt.InputError(f"'{sh.Name}'은(는) 글을 넣을 수 없는 도형입니다 (그림·표·차트). 글상자나 본문 개체 틀 이름을 쓰세요.")
        old = str(sh.TextFrame.TextRange.Text).replace("\r", "\n").strip() if _has(sh.TextFrame, "HasText") else ""
        new = "\n".join(lines)
        info: Dict[str, Any] = {"slide": slide_no, "shape": str(sh.Name), "old_text": _preview_text(old, 200), "new_text": _preview_text(new, 200)}
        if old == new:
            t.set_focus(slide=slide_no, changed=False)
            return rt.ok(TOOL, "write", f"슬라이드 {slide_no} '{sh.Name}': 이미 같은 글이라 바꾸지 않았습니다",
                         "VERIFIED: 바뀐 것 없음. 다시 실행하지 마세요.", verified=True, target_resolved=t.resolved(), unchanged=True, **info)
        if args.dry_run:
            return rt.dry_run_result(TOOL, "write", f"미리보기: 슬라이드 {slide_no} '{sh.Name}'의 글을 바꿉니다: "
                                     f"'{_preview_text(old, 40)}' -> '{_preview_text(new, 40)}' (바꾸지 않음)", 1,
                                     [f"슬라이드 {slide_no} '{sh.Name}': {_preview_text(old, 60)} -> {_preview_text(new, 60)}"],
                                     rt.dry_run_command(TOOL, f"슬라이드 {slide_no} '{sh.Name}' 글 바꾸기", overwrite=bool(old)),
                                     target_resolved=t.resolved(), **info)
        if old and not args.overwrite:
            raise rt.UserDecisionError(
                f"슬라이드 {slide_no} '{sh.Name}'에 글이 있습니다 ({_preview_text(old, 80)}). 바꾸면 이 글이 없어집니다. 아무것도 바꾸지 않았습니다.",
                "지금 글과 새 글을 사용자에게 보여 주고, 바꾸라고 하면 같은 번호의 confirm_commands를 1번 실행하세요.",
                candidates=[f"'{sh.Name}'의 글을 새 글로 바꾸기"], commands=[rt.retry_command({"--overwrite": True})],
                log_message=f"슬라이드 {slide_no} 도형에 글이 있음 ({rt.DOC_OMITTED})", destructive=True)
        sh.TextFrame.TextRange.Text = "\r".join(lines)
        t.set_focus(slide=slide_no)
        got = str(sh.TextFrame.TextRange.Text).replace("\r", "\n").strip()
        resolved = t.resolved()
    preview = f"슬라이드 {slide_no} '{info['shape']}'의 글을 바꿨습니다 ({len(lines)}줄)"
    if got != new.strip():
        return rt.ok_unverified(TOOL, "write", preview, [f"다시 읽은 글이 넣은 글과 다름 ({_preview_text(got, 60)})"],
                                target_resolved=resolved, **info)
    return rt.ok(TOOL, "write", preview, "VERIFIED: 도형의 글을 다시 읽어 확인했습니다. 같은 명령을 다시 실행하지 마세요. "
                 "잘못 바꿨으면 office.py undo로 되돌립니다.", verified=True, target_resolved=resolved, **info)


def cmd_write(args) -> Dict[str, Any]:
    if args.shape:
        return _write_shape(args)
    if args.dry_run:
        raise rt.dry_run_unsupported(TOOL, "write")
    if not any([args.title, args.text, args.file, args.image, args.cards]):
        raise rt.InputError("넣을 내용이 없습니다. --title, --text, --file, --image, --cards 중 하나 이상을 지정하세요.")
    theme = off.get_theme(args.theme or "navy")
    if args.title:
        args.title = md.plain_heading(args.title)   # 슬라이드 제목은 .md 제목·word --title과 같은 규칙 (** __를 뺀 글)
    md_sections: List[Dict[str, Any]] = []
    md_warnings: List[str] = []
    table_rows: Optional[List[List[str]]] = None
    if args.file:
        path = rt.resolve_input_path(args.file, "입력 파일")
        if os.path.splitext(path)[1].lower() in (".csv", ".tsv"):
            table_rows = rt.read_csv_rows(path)
            if not table_rows:
                raise rt.InputError(f"입력 표 파일이 비어 있습니다: {args.file}")
        else:
            md_text = rt.read_text_file(path)
            md_sections = _md_sections(md_text)
            if md.front_matter_lines(md_text):
                md_warnings.append(md.front_matter_warning(md.front_matter_lines(md_text)))
                rt.log("WARN", md_warnings[-1])
            if len(md_sections) > MAX_NEW_SLIDES:
                raise rt.InputError(f"파일 내용이 슬라이드 {len(md_sections)}장 분량입니다 (최대 {MAX_NEW_SLIDES}장). 내용을 줄여 주세요.")
    image = rt.resolve_input_path(args.image, "그림 파일") if args.image else None
    cards = _parse_cards(args.cards) if args.cards else None
    if md_sections and args.title:
        # --title은 첫 슬라이드의 제목 (.md가 있어도). 첫 절의 제목이 --title과 다르면 그 제목은 본문 첫 줄로
        first = dict(md_sections[0])
        if first["title"] and first["title"].strip().casefold() != args.title.strip().casefold():
            first["lines"] = [(first["title"], "")] + list(first["lines"])
        first["title"] = args.title
        md_sections[0] = first

    with off.open_target("ppt", args.target, readonly=False, focus=not args.no_focus, create=True,
                         tool=TOOL, command="write") as t:
        prs = t.doc
        created: List[str] = []
        touched: List[int] = []
        placed: List[Tuple[int, Any]] = []   # (슬라이드 번호, 이번에 넣은 도형)
        warnings: List[str] = list(md_warnings)
        slide_no, make_new = _resolve_slide(prs, args.slide, default_new=bool(md_sections or table_rows) and args.slide is None)
        insert_at = int(prs.Slides.Count) + 1 if slide_no is None else slide_no + 1
        if not make_new and slide_no is not None and not args.overwrite:
            # 기존 슬라이드의 제목을 바꾸게 되면 사용자 확인 (노트와 같은 방식). 새 슬라이드는 해당 없음
            new_title = (md_sections[0]["title"] if md_sections else args.title) or ""
            # 제목 개체틀의 글만 기존 제목 (slide_units의 title은 개체틀이 없으면 본문 첫 줄이라 제목으로 보면 안 됨)
            title_shape = _title_shape(prs.Slides(slide_no))
            old_title = ""
            if title_shape is not None:
                try:
                    old_title = str(title_shape.TextFrame.TextRange.Text).strip()
                except Exception:
                    old_title = ""
            if new_title.strip() and old_title and old_title.strip() != new_title.strip():
                raise rt.UserDecisionError(
                    f"슬라이드 {slide_no}에 이미 제목이 있습니다 ({_preview_text(old_title)}). 아무것도 쓰지 않았습니다.",
                    log_message=f"슬라이드 {slide_no}에 이미 제목이 있습니다 (제목 {rt.DOC_OMITTED}). 아무것도 쓰지 않았습니다.",
                    directive="기존 제목을 바꿀지 정해야 합니다. 제목을 사용자에게 보여 주고, 사용자가 바꾸기를 분명히 원할 때만 "
                    "--overwrite를 붙여 다시 실행하세요. 제목을 바꾸지 않으려면 --title을 빼거나 --slide new로 새 슬라이드에 쓰세요.")

        def open_slide(with_title: bool) -> Tuple[Any, int]:
            nonlocal make_new, insert_at
            if make_new:
                s = _new_slide(prs, insert_at, with_title)
                no = insert_at
                insert_at += 1
            else:
                no = slide_no
                s = prs.Slides(no)
                make_new = True              # 두 번째 섹션부터는 새 슬라이드
            touched.append(no)
            return s, no

        if md_sections:
            for k, sec in enumerate(md_sections):
                s, no = open_slide(bool(sec["title"]))
                w = SlideWriter(prs, s, theme)
                if sec["title"]:
                    w.title(sec["title"])
                if sec["lines"]:
                    w.body(sec["lines"])
                for tb in sec["tables"]:
                    w.table(tb, markdown=True)
                placed += [(no, x) for x in w.shapes]
                created.append(f"슬라이드 {no}: " + ", ".join(w.created))
        if table_rows:
            header, body_rows = table_rows[0], table_rows[1:] or [[""] * len(table_rows[0])]
            for k in range(0, len(body_rows), TABLE_ROWS_PER_SLIDE):     # 긴 표는 머리글을 반복해 여러 슬라이드로
                s, no = open_slide(bool(args.title))
                w = SlideWriter(prs, s, theme)
                if args.title:
                    w.title(args.title + (f" ({k // TABLE_ROWS_PER_SLIDE + 1})" if len(body_rows) > TABLE_ROWS_PER_SLIDE else ""))
                w.table([header] + body_rows[k:k + TABLE_ROWS_PER_SLIDE])
                placed += [(no, x) for x in w.shapes]
                created.append(f"슬라이드 {no}: " + ", ".join(w.created))
        if args.text or (args.title and not table_rows and not md_sections) or image or cards:
            # 파일 내용을 이미 넣었으면 그 마지막 슬라이드에 이어서, 아니면 대상 슬라이드에
            s, no = (prs.Slides(touched[-1]), touched[-1]) if touched else open_slide(bool(args.title))
            w = SlideWriter(prs, s, theme)
            if args.title and not table_rows and not md_sections:
                w.title(args.title)
            if args.text:
                if w._body_placeholder() is None and w.space_left() < MIN_BLOCK_HEIGHT:
                    # 앞 내용 아래 공간이 없으면 새 슬라이드에 (높이 40pt로 슬라이드 밖에 놓이던 것, 검토 C8)
                    placed += [(no, x) for x in w.shapes]
                    created.append(f"슬라이드 {no}: " + ", ".join(w.created))
                    s, no = open_slide(False)
                    w = SlideWriter(prs, s, theme)
                    warnings.append(f"앞 내용 아래에 공간이 부족해 본문을 새 슬라이드 {no}에 넣었습니다.")
                    rt.log("WARN", warnings[-1])
                w.body(_text_lines(args.text))
            for kind in (["cards"] if cards else []) + (["image"] if image else []):
                if w.space_left() < MIN_BLOCK_HEIGHT:   # 공간이 부족하면 다음 새 슬라이드에 (슬라이드 밖으로 나가지 않게)
                    placed += [(no, x) for x in w.shapes]
                    created.append(f"슬라이드 {no}: " + ", ".join(w.created))
                    s, no = open_slide(False)
                    w = SlideWriter(prs, s, theme)
                    what = "카드" if kind == "cards" else "그림"
                    warnings.append(f"앞 내용 아래에 공간이 부족해 {what}를 새 슬라이드 {no}에 넣었습니다.")
                    rt.log("WARN", warnings[-1])
                if kind == "cards":
                    w.cards(cards)
                else:
                    w.image(image, args.width)
            placed += [(no, x) for x in w.shapes]
            created.append(f"슬라이드 {no}: " + ", ".join(w.created))
        expect = [args.title or ""] + [sec["title"] or "" for sec in md_sections] \
            + [sec["lines"][0][0] for sec in md_sections if sec["lines"]] \
            + ([_text_lines(args.text)[0][0]] if args.text and _text_lines(args.text) else []) \
            + ([str(table_rows[0][0])] if table_rows else []) + ([cards[0][0]] if cards else [])
        problems = _verify_slides(prs, sorted(set(touched)), expect) + _verify_inside(prs, placed)
        t.set_focus(slide=touched[-1])
        resolved = t.resolved()
    extra: Dict[str, Any] = {"warnings": warnings} if warnings else {}
    if problems:
        return rt.ok_unverified(TOOL, "write", "; ".join(created), problems,
                                target_resolved=resolved, slides_written=sorted(set(touched)), **extra)
    return rt.ok(TOOL, "write", "; ".join(created), "VERIFIED: 쓴 슬라이드를 다시 읽어 확인했습니다. 같은 명령을 다시 실행하지 마세요.",
                 verified=True, target_resolved=resolved, slides_written=sorted(set(touched)), **extra)


def _verify_inside(prs: Any, placed: List[Tuple[int, Any]]) -> List[str]:
    """R1: 이번에 넣은 도형을 다시 읽어 슬라이드 경계 안에 있는지 확인합니다 (밖으로 나간 도형은 보이지 않음)."""
    try:
        width, height = float(prs.PageSetup.SlideWidth), float(prs.PageSetup.SlideHeight)
    except Exception as e:
        return [f"슬라이드 크기를 읽지 못함: {type(e).__name__}"]
    problems: List[str] = []
    for no, sh in placed:
        try:
            box = {"left": float(sh.Left), "top": float(sh.Top), "width": float(sh.Width), "height": float(sh.Height)}
        except Exception as e:
            problems.append(f"슬라이드 {no}의 도형 위치를 다시 읽지 못함: {type(e).__name__}")
            continue
        if off.outside_area(box, width, height):
            problems.append(f"슬라이드 {no}의 '{sh.Name}'이(가) 슬라이드 밖으로 나갑니다 "
                            f"(위 {box['top']:.0f}, 높이 {box['height']:.0f}, 슬라이드 높이 {height:.0f})")
    return problems


def _verify_slides(prs: Any, slides: List[int], expect: List[str]) -> List[str]:
    """R1: 쓴 슬라이드를 다시 읽어 제목·첫 줄·표 머리글·카드 제목이 들어 있는지 확인합니다 (공백 무시).
    expect는 넣은 글 그대로(굵게 표시는 이미 뺀 글)이므로 ** __를 빼지 않습니다 (코드 줄·제목에는 글자로 남음)."""
    try:
        got = "".join(re.sub(r"\s+", "", slide_units(prs.Slides(n), n)["title"] + slide_units(prs.Slides(n), n)["text"])
                      for n in slides)
    except Exception as e:
        return [f"쓴 슬라이드를 다시 읽지 못함: {type(e).__name__}"]
    missing = [x for x in expect if x and re.sub(r"\s+", "", x)[:20] not in got]
    return [f"다시 읽은 슬라이드에서 '{x[:20]}'을(를) 찾지 못함" for x in missing]


# ==============================================================================
# 4. replace
# ==============================================================================

def _load_map(args) -> Dict[str, str]:
    if args.map:
        raw = args.map.strip()
        if not raw.startswith("{"):
            raw = rt.read_text_file(rt.resolve_input_path(raw, "치환 목록 파일"))
        try:
            data = json.loads(raw)
        except json.JSONDecodeError as e:
            raise rt.InputError(f"--map은 JSON 객체여야 합니다 (예: {{\"진행중\": \"완료\"}}): {e}")
        if not isinstance(data, dict) or not data:
            raise rt.InputError("--map은 비어 있지 않은 JSON 객체여야 합니다.")
        mapping = {str(k): ("" if v is None else str(v)) for k, v in data.items()}
    elif args.find is not None:
        mapping = {args.find: args.find if args.replace is None else args.replace}
    else:
        raise rt.InputError("--find 또는 --map을 지정하세요.")
    if "" in mapping:
        # 빈 글을 찾으면 모든 글자 사이에 바꿀 글이 끼어들어 글 상자를 모두 망가뜨림
        raise rt.InputError("찾을 글(--find 또는 --map의 키)이 비어 있어 바꾸지 않았습니다. 빈 곳에 글을 넣으려면 ppt.py write를 쓰세요 "
                            "(예: ppt.py write --slide 2 --text \"- 항목\"). 바꿀 글이 있으면 찾을 글을 정확히 지정하세요.")
    return mapping


def replace_in_range(tr: Any, find: str, rep: str, match_case: bool, fmt: Dict[str, Any]) -> Tuple[int, int]:
    """TextRange.Replace를 반복합니다. 바꾼 글 뒤부터 이어서 찾아 무한 반복을 막습니다.

    (바꾼 곳 수, 찾았지만 이미 바꿀 글과 같아 그대로 둔 곳 수)를 돌려줍니다. 서식이 없으면 찾은 곳의 글이
    바꿀 글과 글자 그대로 같은 곳(대소문자까지)은 바꾸지 않고 세기만 합니다
    (예: --find iphone --replace iPhone에서 이미 'iPhone'인 곳. 바꾼 곳으로 세면 되돌리기 지점이 남음).
    """
    text = str(tr.Text)
    if (find not in text) if match_case else (find.lower() not in text.lower()):
        return 0, 0
    count = same = 0
    after = 0
    while count + same < 1000:
        text = str(tr.Text)
        idx = text.find(find, after) if match_case else text.lower().find(find.lower(), after)
        if idx < 0:
            break
        if not fmt and text[idx:idx + len(find)] == rep:
            same += 1                  # 이미 바꿀 글 그대로: 건드리지 않음
            after = idx + len(find)
            continue
        # PowerPoint 위치는 UTF-16 단위: 앞에 BMP 밖 글자(이모지 등)가 있으면 파이썬 글자 위치와 다름
        after_u16 = len(text[:idx].encode("utf-16-le")) // 2
        found = tr.Replace(FindWhat=find, ReplaceWhat=rep, After=after_u16, MatchCase=-1 if match_case else 0, WholeWords=0)
        if found is None:
            break
        count += 1
        if fmt:
            font = found.Font
            if "bold" in fmt:
                font.Bold = -1 if fmt["bold"] else 0
            if "italic" in fmt:
                font.Italic = -1 if fmt["italic"] else 0
            if "size" in fmt:
                font.Size = fmt["size"]
            if "color" in fmt:
                font.Color.RGB = fmt["color"]
        after = idx + len(rep)            # 다음 찾기는 파이썬 글자 위치로 (PowerPoint 위치와 섞지 않음)
        if after >= len(str(tr.Text)):
            break
    return count, same


def count_in_range(tr: Any, find: str, rep: str, match_case: bool, fmt: Dict[str, Any]) -> Tuple[int, int]:
    """replace_in_range와 같은 규칙으로 세기만 합니다 (--dry-run: 글을 바꾸지 않음).

    (바뀔 곳 수, 찾았지만 바꿀 글과 같아 그대로 둘 곳 수). 바꾼 뒤 바꾼 글 뒤부터 찾는 것은
    바꾸기 전 글에서 찾은 글 뒤부터 찾는 것과 같습니다.
    """
    text = str(tr.Text)
    hay, needle = (text, find) if match_case else (text.lower(), find.lower())
    count = same = 0
    after = 0
    while count + same < 1000:
        idx = hay.find(needle, after)
        if idx < 0:
            break
        if not fmt and text[idx:idx + len(find)] == rep:
            same += 1
        else:
            count += 1
        after = idx + len(find)
    return count, same


def _num(v: float) -> str:
    return f"{v:g}"


def _replace_command(args) -> Dict[str, Any]:
    """--dry-run을 뺀 같은 replace 명령 (confirm_commands용)."""
    argv: List[Any] = ["replace", "--target", args.target]
    if args.map:
        argv += ["--map", args.map]
    else:
        argv += ["--find", args.find]
        if args.replace is not None:
            argv += ["--replace", args.replace]
    if args.slide:
        argv += ["--slide", args.slide]
    for flag, on in (("--match-case", args.match_case), ("--bold", args.bold), ("--italic", args.italic)):
        if on:
            argv.append(flag)
    if args.size:
        argv += ["--size", _num(args.size)]
    if args.color:
        argv += ["--color", args.color]
    if args.no_focus:
        argv.append("--no-focus")
    return rt.next_command("ppt.py", argv, "찾아 바꾸기 실행", destructive=True)


def _replace_dry_run(args, mapping: Dict[str, str], fmt: Dict[str, Any]) -> Dict[str, Any]:
    """replace --dry-run: 읽기 세션으로 바뀔 곳만 셉니다 (되돌리기 지점 없음)."""
    counts: Dict[str, int] = {}
    same: Dict[str, int] = {}
    planned: List[str] = []
    slides: List[int] = []
    with off.open_target("ppt", args.target, readonly=True) as t:
        prs = t.doc
        numbers = [int(args.slide)] if args.slide else list(range(1, prs.Slides.Count + 1))
        for n in numbers:
            if not 1 <= n <= prs.Slides.Count:
                raise rt.InputError(f"슬라이드 {n}번이 없습니다 (전체 {prs.Slides.Count}장).")
            slide = prs.Slides(n)
            here: Dict[str, int] = {}
            for k in range(1, slide.Shapes.Count + 1):
                for tr in iter_text_ranges(slide.Shapes(k)):
                    for find, rep in mapping.items():
                        c, hits = count_in_range(tr, find, rep, args.match_case, fmt)
                        if hits:
                            same[find] = same.get(find, 0) + hits
                        if c:
                            here[find] = here.get(find, 0) + c
            for find, c in here.items():
                counts[find] = counts.get(find, 0) + c
                rep = mapping[find]
                what = "서식만" if rep == find else f"-> '{_preview_text(rep, 20)}'"
                planned.append(f"슬라이드 {n}: '{_preview_text(find, 20)}' {what} {c}곳")
            if here:
                slides.append(n)
        resolved = t.resolved()
    total = sum(counts.values())
    extra: Dict[str, Any] = {"target_resolved": resolved, "by_text": counts, "slides": slides}
    if same:
        extra["unchanged_same_text"] = same
    if not total:
        why = f"찾은 {sum(same.values())}곳은 바꿀 글이 찾는 글과 같아 " if same else "일치하는 문구가 없어 "
        preview = f"미리보기: {why}바뀔 것 없음 (바꾸지 않음)"
    else:
        preview = f"미리보기: {total}곳이 바뀝니다 (슬라이드 {', '.join(map(str, slides))}) (바꾸지 않음)"
    return rt.dry_run_result(TOOL, "replace", preview, total, planned, _replace_command(args), **extra)


def cmd_replace(args) -> Dict[str, Any]:
    mapping = _load_map(args)
    fmt: Dict[str, Any] = {}
    if args.bold is not None:
        fmt["bold"] = args.bold
    if args.italic is not None:
        fmt["italic"] = args.italic
    if args.size:
        fmt["size"] = args.size
    if args.color:
        fmt["color"] = off.parse_color(args.color)
    if not args.map and args.replace is None and not fmt:
        # 바꿀 글도 서식도 없으면 아무것도 바뀌지 않는데 'N곳 처리'로 보고되던 문제 (excel/word replace와 같게)
        raise rt.InputError("--replace(바꿀 글)나 서식 옵션(--bold, --color 등)이 없어 바꿀 것이 없습니다. "
                            "사용자에게 무엇으로 바꿀지 확인하세요. 글을 지우려면 --replace \"\"를 씁니다.")
    if args.dry_run:
        return _replace_dry_run(args, mapping, fmt)
    counts: Dict[str, int] = {}
    same: Dict[str, int] = {}          # 바꿀 글이 찾는 글과 같고 서식도 없어 바꾸지 않은 문구: 찾은 곳 수
    slides_changed: List[int] = []
    with off.open_target("ppt", args.target, readonly=False, focus=not args.no_focus, tool=TOOL, command="replace") as t:
        prs = t.doc
        numbers = [int(args.slide)] if args.slide else list(range(1, prs.Slides.Count + 1))
        for n in numbers:
            if not 1 <= n <= prs.Slides.Count:
                raise rt.InputError(f"슬라이드 {n}번이 없습니다 (전체 {prs.Slides.Count}장).")
            slide = prs.Slides(n)
            changed = False
            for k in range(1, slide.Shapes.Count + 1):
                for tr in iter_text_ranges(slide.Shapes(k)):
                    for find, rep in mapping.items():
                        # 찾은 곳마다 바꿀 글과 비교: 글자 그대로 같은 곳은 세기만 하고(바꾼 것으로 표시하지 않음),
                        # 대소문자만 다른 곳은 PowerPoint가 바꿀 글 그대로 써 주므로(대소문자 통일) 실제로 바꿈
                        c, hits = replace_in_range(tr, find, rep, args.match_case, fmt)
                        if hits:
                            same[find] = same.get(find, 0) + hits
                        if c:
                            counts[find] = counts.get(find, 0) + c
                            changed = True
            if changed:
                slides_changed.append(n)
        # R1: 바꾼 슬라이드를 다시 읽어 확인 (바꾼 글이 찾는 글을 품지 않으면 찾는 글이 남지 않아야 함)
        problems = []
        if slides_changed:
            text = "\n".join(slide_units(prs.Slides(n), n)["title"] + "\n" + slide_units(prs.Slides(n), n)["text"] for n in slides_changed)
            for find in counts:
                rep_text = mapping[find]
                hay, needle = (text, find) if args.match_case else (text.lower(), find.lower())
                if needle not in rep_text.lower() and needle in hay:
                    problems.append(f"'{find}'이(가) 슬라이드에 남아 있음")
                elif rep_text and rep_text not in text:
                    # 바꿀 글은 대소문자까지 그대로 있어야 함 (대소문자만 바꾸는 경우도 확인)
                    problems.append(f"'{rep_text}'을(를) 다시 읽은 슬라이드에서 찾지 못함")
            t.set_focus(slide=slides_changed[0])
        resolved = t.resolved()
    total = sum(counts.values())
    result = {"target_resolved": resolved, "replaced": total, "by_text": counts, "slides": slides_changed}
    if same:
        result["unchanged_same_text"] = same
    if not total and same:
        found = sum(same.values())
        return rt.ok(TOOL, "replace", f"찾은 {found}곳은 바꿀 글이 찾는 글과 같아 바꾼 곳이 없습니다 (0곳 바뀜)",
                     f"VERIFIED: 바꾼 곳 0곳. 찾은 {found}곳은 바꿀 글이 찾는 글과 같아 바뀐 것이 없습니다. "
                     "다시 실행하지 말고 사용자에게 무엇으로 바꿀지 확인하세요.", **result)
    if not total:
        return rt.ok(TOOL, "replace", "일치하는 문구가 없습니다",
                     "VERIFIED: 일치하는 문구가 없습니다. 검색어를 임의로 바꿔 재시도하지 말고 사용자에게 알리세요.", **result)
    preview = f"{total}곳 처리 (슬라이드 {', '.join(map(str, slides_changed))})"
    if problems:
        return rt.ok_unverified(TOOL, "replace", preview, problems, **result)
    return rt.ok(TOOL, "replace", preview, "VERIFIED: 바꾼 슬라이드를 다시 읽어 확인했습니다.", verified=True, **result)


# ==============================================================================
# 5. export
# ==============================================================================

def cmd_export(args) -> Dict[str, Any]:
    ext = os.path.splitext(args.save)[1].lower() if args.save else ".pdf"
    if ext not in (".pdf", ".pptx", ".pptm", ".ppt"):
        raise rt.InputError(f"--save 확장자는 .pdf 또는 .pptx여야 합니다: {args.save}")
    with off.open_target("ppt", args.target, readonly=True) as t:
        prs = t.doc
        if ext == ".pdf":
            off.require_pdf_export(t.app, ".pptx")
            out = rt.resolve_output_path(args.save, os.path.splitext(t.name)[0], ".pdf", near=t.path)
            try:
                # PrintRange(개체 인자)를 생략하면 pywin32가 인자 변환에서 TypeError를 냅니다. None을 명시합니다.
                prs.ExportAsFixedFormat(os.path.normpath(out), 2, PrintRange=None)   # ppFixedFormatTypePDF
            except Exception as e:
                if "기능이 설치되어 있지 않아" in str(e):
                    raise rt.InputError("PowerPoint PDF 내보내기 기능이 없습니다. --save에 .pptx를 쓰세요.")
                raise
        else:
            src_ext = os.path.splitext(t.path)[1].lower() if t.path else ".pptx"
            out = rt.resolve_output_path(args.save, os.path.splitext(t.name)[0] + "_사본", src_ext or ".pptx", near=t.path)
            prs.SaveCopyAs(os.path.normpath(out))                     # 원본 형식 그대로 사본
        resolved = t.resolved()
    return rt.ok_file(TOOL, "export", f"{resolved['name']}을(를) 저장: {out}", out,
                      "VERIFIED: 저장한 파일을 확인했습니다. 원본은 그대로입니다.", target_resolved=resolved, output_file=out)


# ==============================================================================
# 5-1. notes (발표자 노트)
# ==============================================================================

def cmd_notes(args) -> Dict[str, Any]:
    if args.text is None and not args.file:
        return rt.dry_run_read_only(_notes_read(args), args.dry_run, "notes 읽기")
    if args.dry_run:
        raise rt.dry_run_unsupported(TOOL, "notes 쓰기")
    if args.slide is None:
        raise rt.InputError("노트를 쓸 슬라이드를 --slide 번호로 지정하세요 (예: --slide 3).")
    if args.file:
        text = rt.read_text_file(rt.resolve_input_path(args.file, "노트 파일")).strip()
    else:
        text = rt.unescape_newlines(args.text).strip()
    if args.append and not text:
        raise rt.InputError("--append로 덧붙일 노트 내용이 없습니다. --text 또는 --file을 지정하세요.")
    if args.append and args.overwrite:
        raise rt.InputError("--append(덧붙이기)와 --overwrite(바꾸기) 중 하나만 지정하세요.")
    with off.open_target("ppt", args.target, readonly=False, focus=not args.no_focus, tool=TOOL, command="notes") as t:
        prs = t.doc
        count = int(prs.Slides.Count)
        if not 1 <= args.slide <= count:
            raise rt.InputError(f"--slide는 1~{count} 사이여야 합니다 (전체 {count}장).")
        slide = prs.Slides(args.slide)
        body = _notes_body(slide)
        if body is None:
            raise rt.InputError(f"슬라이드 {args.slide}의 노트 페이지에 본문 개체틀이 없어 노트를 쓸 수 없습니다. "
                                "PowerPoint의 노트 마스터에서 본문 개체틀을 켜야 합니다.")
        before = _notes_text(slide)
        if before and not args.append and not args.overwrite:
            raise rt.UserDecisionError(
                f"슬라이드 {args.slide}에 이미 노트가 있습니다 ({_preview_text(before)}). 기존 노트를 지우지 않았습니다.",
                log_message=f"슬라이드 {args.slide}에 이미 노트가 있습니다 (노트 {rt.DOC_OMITTED}). 기존 노트를 지우지 않았습니다.",
                directive="기존 노트를 지울지 정해야 합니다. 노트를 사용자에게 보여 주고, 뒤에 덧붙이려면 --append, "
                "사용자가 바꾸기를 분명히 원할 때만 --overwrite를 붙여 다시 실행하세요.")
        new = (before + "\n" + text).strip() if args.append and before else text
        same = new == before               # 같은 노트로 바꾸기(빈 노트 지우기 포함)는 바뀐 것 없음 (되돌리기 지점을 남기지 않음)
        if not same:
            body.TextFrame.TextRange.Text = new.replace("\n", "\r")
        t.set_focus(slide=args.slide, changed=not same)
        after = _notes_text(slide)
        resolved = t.resolved()
    fields = {"target_resolved": resolved, "slide": args.slide, "before": rt.clamp_text(before, 200)[0],
              "after": rt.clamp_text(after, 200)[0], "appended": bool(args.append)}
    preview = f"{resolved['name']} 슬라이드 {args.slide}: 노트를 " + ("덧붙였습니다" if args.append else ("지웠습니다" if not new else "바꿨습니다"))
    if same:
        fields["unchanged"] = True
        preview = f"{resolved['name']} 슬라이드 {args.slide}: 노트가 이미 같아 바꾸지 않았습니다 (바뀐 것 없음)"
    if re.sub(r"\s+", "", after) != re.sub(r"\s+", "", new):
        return rt.ok_unverified(TOOL, "notes", preview, ["다시 읽은 노트가 쓴 내용과 다름"], **fields)
    if same:
        return rt.ok(TOOL, "notes", preview, "VERIFIED: 노트를 다시 읽어 이미 같은 내용임을 확인했습니다. 바꾼 것이 없으니 다시 실행하지 마세요.",
                     verified=True, **fields)
    return rt.ok(TOOL, "notes", preview, "VERIFIED: 노트를 쓰고 다시 읽어 확인했습니다. 잘못 썼으면 office.py undo로 되돌릴 수 있습니다.",
                 verified=True, **fields)


def _preview_text(text: str, n: int = 40) -> str:
    t = re.sub(r"\s+", " ", text).strip()
    return t if len(t) <= n else t[:n] + "..."


def _slide_title(slide: Any) -> str:
    shape = _title_shape(slide)
    try:
        return str(shape.TextFrame.TextRange.Text).strip() if shape is not None else ""
    except Exception:
        return ""


def _notes_read(args) -> Dict[str, Any]:
    import reader
    with off.open_target("ppt", args.target, readonly=True) as t:
        prs = t.doc
        total = int(prs.Slides.Count)
        spec = args.slides or (str(args.slide) if args.slide else None)      # 읽기에서 --slide 3도 그 슬라이드만
        numbers = reader.parse_pages(spec, total, "--slides") if total else []
        items = []
        for n in numbers:
            slide = prs.Slides(n)
            note = _notes_text(slide)
            if note:
                items.append({"no": n, "title": _slide_title(slide)[:60], "notes": note})
        resolved = t.resolved()
    result: Dict[str, Any] = {"target_resolved": resolved, "slides_total": total, "count": len(items)}
    what = f"슬라이드 {spec}" if spec else "전체 슬라이드"
    if not items:
        result.update(notes=[], truncated=False)
        return rt.ok(TOOL, "notes", f"{resolved['name']}: {what}에 노트 없음 (count 0)",
                     f"VERIFIED: 0건. {what}에 발표자 노트가 없습니다. 사용자에게 그대로 알리세요.", **result)
    shown: List[Dict[str, Any]] = []
    used = 0
    text_cut = 0
    for it in items:
        size = len(it["notes"]) + len(it["title"]) + 20
        if shown and used + size > rt.DEFAULT_MAX_CHARS:
            break
        text, cut = rt.clamp_text(it["notes"], rt.DEFAULT_MAX_CHARS)
        item = dict(it, notes=text)
        if cut:
            item.update(text_truncated=True, text_chars=len(it["notes"]))
            text_cut += 1
        shown.append(item)
        used += size
    truncated = len(shown) < len(items) or bool(text_cut)
    result.update(notes=shown, truncated=truncated)
    if text_cut:
        result["truncated_reason"] = (f"{rt.DEFAULT_MAX_CHARS}자보다 긴 노트 {text_cut}개는 앞부분만 표시 (text_truncated). "
                                      "전체는 PowerPoint에서 직접 확인")
    if len(shown) < len(items):
        rest = ",".join(str(it["no"]) for it in items[len(shown):])      # 고른 슬라이드 중 남은 것만
        result["next_commands"] = [rt.next_command("ppt.py", ["notes", "--target", args.target, "--slides", rest],
                                                   "나머지 노트 이어서 읽기")]
    directive = ""
    if len(shown) < len(items):
        directive = "VERIFIED: 읽기 완료. 일부만 표시되었으면 next_commands로 이어서 읽으세요."
    if text_cut:
        directive = (directive or "VERIFIED: 읽기 완료.") + " 노트가 잘린 슬라이드(text_truncated)는 truncated_reason을 사용자에게 알리세요."
    return rt.ok(TOOL, "notes", f"{resolved['name']}: 노트 있는 슬라이드 {len(shown)}장 표시 ({what} 중 {len(items)}장)",
                 directive, **result)


# ==============================================================================
# 6. slide
# ==============================================================================

def _slide_command(args) -> Dict[str, Any]:
    """--dry-run을 뺀 같은 slide 명령 (confirm_commands용)."""
    argv: List[Any] = ["slide", "--target", args.target, "--action", args.action]
    if args.slide is not None:
        argv += ["--slide", args.slide]
    if args.dest is not None:
        argv += ["--dest", args.dest]
    if args.action == "add" and args.layout != "title_only":
        argv += ["--layout", args.layout]
    if args.all_slides:
        argv.append("--all-slides")
    if args.no_focus:
        argv.append("--no-focus")
    what = {"add": "추가", "duplicate": "복제", "move": "옮기기", "delete": "삭제", "clear": "비우기"}[args.action]
    return rt.next_command("ppt.py", argv, f"슬라이드 {what} 실행", destructive=args.action in ("delete", "clear"))


def _slide_dry_run(args) -> Dict[str, Any]:
    """slide --dry-run: 읽기 세션으로 바뀔 것만 보여 줍니다 (되돌리기 지점 없음). 검사는 실제 실행과 같습니다."""
    a = args.action
    with off.open_target("ppt", args.target, readonly=True) as t:
        prs = t.doc
        count = int(prs.Slides.Count)

        def need(n: Optional[int], what: str = "--slide") -> int:
            if n is None:
                raise rt.InputError(f"slide {a}에는 {what}가 필요합니다.")
            if not 1 <= n <= count:
                raise rt.InputError(f"슬라이드 {n}번이 없습니다 (전체 {count}장).")
            return n

        def label(n: int) -> str:
            title = slide_units(prs.Slides(n), n)["title"]
            return f"슬라이드 {n}" + (f"('{_preview_text(title, 30)}')" if title else "")

        planned: List[str] = []
        if a == "add":
            pos = count + 1 if args.slide is None else min(max(1, args.slide), count + 1)
            n_change, msg = 1, f"{pos}번 위치에 새 슬라이드를 추가합니다"
            planned.append(f"추가: {pos}번 위치 (레이아웃 {args.layout})")
        elif a == "duplicate":
            n = need(args.slide)
            n_change, msg = 1, f"{label(n)}을(를) 복제해 {n + 1}번에 넣습니다"
            planned.append(f"복제: {label(n)} -> {n + 1}번")
        elif a == "move":
            n = need(args.slide)
            dest = need(args.dest, "--dest")
            if n == dest:
                n_change, msg = 0, f"슬라이드 {n}은(는) 이미 {dest}번 위치라 바뀔 것 없음"
            else:
                n_change, msg = 1, f"{label(n)}을(를) {dest}번 위치로 옮깁니다"
                planned.append(f"옮기기: {label(n)} -> {dest}번")
        elif a == "delete":
            n = need(args.slide)
            if count == 1:
                raise rt.InputError("마지막 남은 슬라이드는 삭제할 수 없습니다.")
            shapes = int(prs.Slides(n).Shapes.Count)
            n_change, msg = 1, f"{label(n)}을(를) 삭제합니다 (요소 {shapes}개, 남는 슬라이드 {count - 1}장)"
            planned.append(f"삭제: {label(n)} (요소 {shapes}개)")
        else:   # clear
            if args.slide is None and not args.all_slides:
                raise rt.InputError("clear에는 --slide가 필요합니다. 모든 슬라이드를 비우려면 --all-slides를 명시하세요 (사용자가 요청한 경우에만).")
            numbers = list(range(1, count + 1)) if args.slide is None else [need(args.slide)]
            n_change = 0
            for n in numbers:
                shapes = int(prs.Slides(n).Shapes.Count)
                if shapes:
                    n_change += shapes
                    planned.append(f"비우기: {label(n)} 요소 {shapes}개")
            where = f"슬라이드 {', '.join(map(str, numbers))}" if len(numbers) <= 10 else f"슬라이드 {len(numbers)}장"
            msg = (f"{where}의 요소 {n_change}개를 지웁니다" if n_change else f"{where}에 지울 요소가 없어 바뀔 것 없음")
        resolved = t.resolved()
    return rt.dry_run_result(TOOL, "slide", f"미리보기: {msg} (바꾸지 않음)", n_change, planned, _slide_command(args),
                             target_resolved=resolved, slides_total=count)


def cmd_slide(args) -> Dict[str, Any]:
    a = args.action
    if a == "list":
        # list는 읽기만 하므로 --dry-run은 받고 무시 (그 사실을 warnings로)
        with off.open_target("ppt", args.target, readonly=True) as t:
            prs = t.doc
            items = [{"no": i, "title": slide_units(prs.Slides(i), i)["title"][:60]} for i in range(1, prs.Slides.Count + 1)]
            resolved = t.resolved()
        if not items:      # R2 0건
            return rt.dry_run_read_only(rt.ok(TOOL, "slide", "0건: 슬라이드가 없습니다",
                                              "VERIFIED: 0건. 이 프레젠테이션에는 슬라이드가 없습니다. 사용자에게 그대로 알리세요.",
                                              target_resolved=resolved, slides=[], count=0), args.dry_run, "slide list")
        return rt.dry_run_read_only(rt.ok(TOOL, "slide", f"슬라이드 {len(items)}장", target_resolved=resolved, slides=items,
                                          count=len(items)), args.dry_run, "slide list")
    if args.dry_run:
        return _slide_dry_run(args)
    with off.open_target("ppt", args.target, readonly=False, focus=not args.no_focus, tool=TOOL, command="slide") as t:
        prs = t.doc
        count = int(prs.Slides.Count)

        def need(n: Optional[int], what: str = "--slide") -> int:
            if n is None:
                raise rt.InputError(f"slide {a}에는 {what}가 필요합니다.")
            if not 1 <= n <= count:
                raise rt.InputError(f"슬라이드 {n}번이 없습니다 (전체 {count}장).")
            return n

        if a == "add":
            pos = count + 1 if args.slide is None else min(max(1, args.slide), count + 1)
            prs.Slides.Add(pos, PP_LAYOUT[args.layout])
            focus, msg = pos, f"{pos}번 위치에 새 슬라이드를 추가했습니다"
        elif a == "duplicate":
            n = need(args.slide)
            prs.Slides(n).Duplicate()
            focus, msg = n + 1, f"슬라이드 {n}을(를) 복제해 {n + 1}번에 넣었습니다"
        elif a == "move":
            n = need(args.slide)
            dest = need(args.dest, "--dest")
            moved_title = slide_units(prs.Slides(n), n)["title"]
            if n == dest:          # 같은 자리로 옮기기는 바뀐 것 없음 (되돌리기 지점을 남기지 않음)
                focus, msg = dest, f"슬라이드 {n}은(는) 이미 {dest}번 위치라 옮기지 않았습니다 (바뀐 것 없음)"
            else:
                prs.Slides(n).MoveTo(dest)
                focus, msg = dest, f"슬라이드 {n}을(를) {dest}번 위치로 옮겼습니다"
        elif a == "delete":
            n = need(args.slide)
            if count == 1:
                raise rt.InputError("마지막 남은 슬라이드는 삭제할 수 없습니다.")
            prs.Slides(n).Delete()
            focus, msg = min(n, count - 1), f"슬라이드 {n}을(를) 삭제했습니다 (되돌리기: office.py undo)"
        else:   # clear
            if args.slide is None and not args.all_slides:
                raise rt.InputError("clear에는 --slide가 필요합니다. 모든 슬라이드를 비우려면 --all-slides를 명시하세요 (사용자가 요청한 경우에만).")
            numbers = list(range(1, count + 1)) if args.slide is None else [need(args.slide)]
            removed = 0
            for n in numbers:
                shapes = prs.Slides(n).Shapes
                for k in range(shapes.Count, 0, -1):
                    shapes(k).Delete()
                    removed += 1
            focus, msg = numbers[0], f"슬라이드 {', '.join(map(str, numbers))}의 요소 {removed}개를 지웠습니다"
        # 빈 슬라이드 비우기, 같은 자리로 옮기기는 바뀐 것 없음
        noop = (a == "clear" and removed == 0) or (a == "move" and n == dest)
        t.set_focus(slide=focus, changed=not noop)
        resolved = t.resolved()
        total = int(prs.Slides.Count)
        # R1: 슬라이드 수(추가·복제 +1, 삭제 -1), 옮긴 슬라이드 제목, 비운 슬라이드의 요소 수를 다시 읽어 확인
        want = {"add": count + 1, "duplicate": count + 1, "delete": count - 1}.get(a, count)
        problems = [] if total == want else [f"슬라이드가 {total}장 (기대 {want}장)"]
        if a == "move" and moved_title and slide_units(prs.Slides(dest), dest)["title"] != moved_title:
            problems.append(f"{dest}번 슬라이드 제목이 옮긴 슬라이드('{moved_title[:20]}')와 다름")
        if a == "clear":
            left = sum(int(prs.Slides(n).Shapes.Count) for n in numbers)
            if left:
                problems.append(f"비운 슬라이드에 요소 {left}개가 남음")
    if problems:
        return rt.ok_unverified(TOOL, "slide", msg, problems, target_resolved=resolved, slides_total=total)
    return rt.ok(TOOL, "slide", msg, "VERIFIED: 슬라이드를 다시 읽어 확인했습니다.", target_resolved=resolved, slides_total=total)


# ==============================================================================
# 7. 진입점
# ==============================================================================

def build_cli():
    parser, sub = rt.build_parser(TOOL, "파워포인트 작업")

    def common(p, write=False):
        p.add_argument("--target", default="active", help="active(기본) | 열린 프레젠테이션 이름 | 파일 경로")
        if write:
            p.add_argument("--no-focus", action="store_true", help="작업 후 화면을 고친 슬라이드로 옮기지 않음")
            p.add_argument("--dry-run", action="store_true", help="바꾸지 않고 바뀔 것만 미리 보기 (미리보기가 없는 동작은 바꾸지 않고 멈춤, 목록은 그대로 읽음)")

    p = sub.add_parser("read", help="슬라이드 내용 읽기")
    common(p)
    p.add_argument("--slides", help="슬라이드 번호 (예: 3, 1-3, 1,4)")
    p.add_argument("--save", nargs="?", const=True, help="md 파일로 저장")

    p = sub.add_parser("info", help="프레젠테이션 요약")
    common(p)
    p.add_argument("--all", action="store_true", help="열린 프레젠테이션 전체 목록")
    p.add_argument("--detail", action="store_true",
                   help="도형별 위치/크기/글자 크기/목록 종류와 배치 경고(슬라이드 밖, 글자 넘침, 겹침)")
    p.add_argument("--slides", help="--detail 대상 슬라이드 (예: 3, 2-4). 생략 시 전체")

    p = sub.add_parser("write", help="내용 넣기")
    common(p, write=True)
    p.add_argument("--slide", help="슬라이드 번호 또는 new (기본: --file이면 new, 아니면 마지막 슬라이드)")
    p.add_argument("--title", help="슬라이드 제목")
    p.add_argument("--text", help="본문 (- 로 시작하는 줄은 글머리 기호, 줄바꿈은 \\n)")
    p.add_argument("--file", help="입력 파일 (.md: 제목마다 새 슬라이드, .csv: 표)")
    p.add_argument("--image", help="그림 파일 (남은 공간에 비율 유지로 맞춤)")
    p.add_argument("--width", type=float, help="그림 너비(pt)")
    p.add_argument("--cards", help="카드 '제목:내용; 제목:내용' (1~5개)")
    p.add_argument("--theme", help=f"색 테마: {', '.join(off.THEMES)} (기본 navy)")
    p.add_argument("--shape", help="이 도형의 글을 --text로 바꾸기 (새 글상자를 만들지 않음). 이름은 info --detail의 name")
    p.add_argument("--overwrite", action="store_true", help="기존 슬라이드의 제목·도형의 글을 바꾸기 (사용자가 원할 때만)")

    p = sub.add_parser("replace", help="찾아 바꾸기")
    common(p, write=True)
    p.add_argument("--find", help="찾을 글")
    p.add_argument("--replace", help="바꿀 글 (빈 문자열이면 삭제, 생략하면 서식만)")
    p.add_argument("--map", help='여러 개 한 번에: JSON 객체 또는 JSON 파일')
    p.add_argument("--slide", type=int, help="이 슬라이드만 (생략 시 전체)")
    p.add_argument("--match-case", action="store_true", help="대소문자 구분")
    p.add_argument("--bold", action="store_true", default=None)
    p.add_argument("--italic", action="store_true", default=None)
    p.add_argument("--size", type=float)
    p.add_argument("--color")

    p = sub.add_parser("export", help="PDF/PPTX 사본 저장")
    common(p)
    p.add_argument("--save", help="저장 경로 (.pdf 또는 .pptx, 생략 시 대상 파일 옆에 PDF, 없으면 C:\\AIKIT\\data)")

    p = sub.add_parser("notes", help="발표자 노트 읽기 (--text/--file이 있으면 --slide의 노트 쓰기)")
    common(p, write=True)
    p.add_argument("--slides", help="읽을 슬라이드 번호 (예: 3, 1-3, 1,4). 생략 시 전체")
    p.add_argument("--slide", type=int, help="노트를 쓸 슬라이드 번호")
    p.add_argument("--text", help="쓸 노트 내용 (기존 노트를 바꿈). 줄바꿈은 \\n")
    p.add_argument("--file", help="쓸 노트 내용 파일 (.txt/.md)")
    p.add_argument("--append", action="store_true", help="기존 노트 뒤에 덧붙이기")
    p.add_argument("--overwrite", action="store_true", help="기존 노트를 지우고 바꾸기 (사용자가 원할 때만)")

    p = sub.add_parser("shape", help="도형 옮기기·크기 (pt, info --detail의 위치를 보고)")
    common(p, write=True)
    p.add_argument("--slide", type=int, required=True, help="슬라이드 번호")
    p.add_argument("--shape", required=True, help="도형 이름 (info --detail의 name)")
    p.add_argument("--left", type=float, help="왼쪽 위치 (pt)")
    p.add_argument("--top", type=float, help="위쪽 위치 (pt)")
    p.add_argument("--width", type=float, help="너비 (pt, 비율이 잠긴 그림은 높이가 따라감)")
    p.add_argument("--height", type=float, help="높이 (pt)")

    p = sub.add_parser("table", help="슬라이드의 표 읽기 (--cells·--file이 있으면 칸 쓰기)")
    common(p, write=True)
    p.add_argument("--slide", type=int, required=True, help="슬라이드 번호")
    p.add_argument("--shape", help="표 도형 이름 (슬라이드에 표가 하나면 생략)")
    p.add_argument("--cells", help="쓸 칸 \"R2C3=값; R3C1=값\" (R행C열, 칸 여러 개는 ;)")
    p.add_argument("--file", help="칸 값 파일 (.csv, 열: 칸, 값. 예: R2C3,25). 많을 때")
    p.add_argument("--overwrite", action="store_true", help="글이 있는 칸도 바꾸기 (사용자가 확인했을 때만)")

    p = sub.add_parser("slide", help="슬라이드 관리")
    common(p, write=True)
    p.add_argument("--action", required=True, choices=["list", "add", "duplicate", "move", "delete", "clear"])
    p.add_argument("--slide", type=int, help="대상 슬라이드 번호 (add는 넣을 위치)")
    p.add_argument("--dest", type=int, help="move: 옮길 위치 (슬라이드 번호)")
    p.add_argument("--layout", choices=list(PP_LAYOUT), default="title_only", help="add 레이아웃 (기본 title_only)")
    p.add_argument("--all-slides", action="store_true", help="clear: 모든 슬라이드 (명시 필요)")
    return parser


HANDLERS = {"read": cmd_read, "info": cmd_info, "write": cmd_write, "replace": cmd_replace,
            "export": cmd_export, "slide": cmd_slide, "notes": cmd_notes, "shape": cmd_shape, "table": cmd_table}


def main(argv: Optional[Sequence[str]] = None) -> None:
    args = build_cli().parse_args(argv)
    rt.run_cli(TOOL, args.command, lambda: HANDLERS[args.command](args))


if __name__ == "__main__":
    main()
