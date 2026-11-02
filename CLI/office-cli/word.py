"""word.py - 워드 작업

명령:
    read     내용 읽기 (제목 구조 포함, --save로 md 저장). --pages / --paragraphs / --section으로 범위 고르기
    outline  제목 구조(목차): 수준, 단락 번호, 쪽
    comment  검토 메모 목록 (--action add: 기준 글에 메모 달기, --file CSV로 여러 개)
    revisions 변경 내용 추적(수정 기록) 목록
    info     문서 요약 (페이지/단어/글자 수, 표, 그림, 제목 목록). --all: 열린 문서 전체 목록
    write    텍스트/마크다운/표/그림/요약상자/목차 넣기 (기본: 문서 끝, --anchor로 위치 지정)
    replace  찾아 바꾸기 (+서식). 머리글/바닥글/글상자 포함
    export   PDF / DOCX 사본 저장 (원본은 그대로, --page-number는 사본에만)

--target: active(기본, 지금 보고 있는 문서) | 열린 문서 이름 | 파일 경로
닫힌 .docx 읽기는 Word를 띄우지 않고 python-docx로 읽습니다.

예:
    python word.py write --file "검색결과.md"
    python word.py write --text "## 결론\n- 수율 98% 달성" --anchor "추진 배경" --pos after
    python word.py write --image "chart.png" --width 400
    python word.py replace --find "초안" --replace "최종본" --bold
    python word.py export --save report.pdf --page-number
    python word.py read --section "1.1 Scope"
    python word.py comment --action add --anchor "시험 온도" --text "규격 4.2항과 다름"
"""
import difflib
import json
import os
import re
import sys
from typing import Any, Dict, List, Optional, Sequence, Tuple

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _office as off  # noqa: E402
import _runtime as rt  # noqa: E402

TOOL = "word"

# WdBuiltinStyle (언어와 무관한 기본 스타일 번호)
WD_STYLE_NORMAL = -1
WD_STYLE_TITLE = -63
WD_STYLE_LIST_BULLET = -49
WD_STYLE_LIST_NUMBER = -50
WD_HIGHLIGHT = {"yellow": 7, "green": 4, "turquoise": 3, "pink": 5, "blue": 2, "red": 6, "gray": 16, "grey": 16,
                "노랑": 7, "초록": 4, "분홍": 5, "파랑": 2, "빨강": 6, "회색": 16, "none": 0}
FIND_MAX = 255   # Word Find 검색어 길이 한도 (^를 ^^로 바꾼 뒤의 길이)


def _find_len(text: str) -> int:
    """Word Find에 넘기는 길이: _find가 ^를 ^^로 바꾸므로 그 길이로 셉니다."""
    return len(str(text).replace("^", "^^"))


def heading_style(level: int) -> int:
    return -(min(max(level, 1), 9) + 1)       # wdStyleHeading1 = -2 ... Heading9 = -10


# ==============================================================================
# 1. read / info
# ==============================================================================

def _closed_docx(target: str) -> Optional[str]:
    if not target or target.strip().lower() == "active":
        return None
    raw = os.path.abspath(os.path.expandvars(target.strip().strip('"')))
    if not os.path.isfile(raw):
        cand = os.path.join(rt.get_data_dir(), target.strip().strip('"'))
        if not os.path.isfile(cand):
            return None
        raw = cand
    if os.path.splitext(raw)[1].lower() not in (".docx", ".docm"):
        return None
    import zipfile
    if not zipfile.is_zipfile(raw):
        return None
    try:
        if off.win32com is not None and off.match_open("word", raw) is not None:
            return None
    except rt.InputError:
        return None   # 같은 이름의 열린 문서가 여럿: 파일을 읽지 않고 COM 경로에서 [후보] 오류로 안내
    except Exception:
        pass
    return raw


PICTURE_MARK = "[그림]"


def _replace_inline_pictures(rng: Any, raw: str) -> str:
    """본문 그림(InlineShape)은 Range.Text에서 한 글자(Word 2007은 '/', 새 버전은 '\\x01')로 나옵니다.
    그림의 실제 위치에 있는 글자만 [그림]으로 바꿉니다 (본문에 있는 진짜 '/'는 그대로)."""
    if "\x01" not in raw and "/" not in raw:
        return raw
    try:
        shapes = rng.InlineShapes
        count = int(shapes.Count)
        if not count:
            return raw
        start = int(rng.Start)
        offsets = sorted({int(shapes.Item(i).Range.Start) - start for i in range(1, count + 1)}, reverse=True)
    except Exception:
        return raw
    chars = list(raw)
    for pos in offsets:
        if 0 <= pos < len(chars) and chars[pos] in ("\x01", "/"):
            chars[pos] = PICTURE_MARK
    return "".join(chars)


def _com_units(doc: Any, with_start: bool = False) -> List[Dict[str, Any]]:
    """열린 문서의 단락을 제목 구조와 함께 읽습니다 (단락 단위 COM 호출 2회, 페이지 계산 없음). 그림은 [그림].
    with_start: 단락 시작·끝 위치(start, end)도 담습니다 (쪽 계산용, COM 호출 2회 추가)."""
    units: List[Dict[str, Any]] = []
    heading = ""
    n = 0
    for para in doc.Paragraphs:
        rng = para.Range
        raw = _replace_inline_pictures(rng, str(rng.Text))
        text = raw.rstrip("\r\x07").replace("\x07", " ").replace("\x0b", "\n").strip()
        if not text:
            continue
        n += 1
        try:
            level = int(para.OutlineLevel)
        except Exception:
            level = 10
        if 1 <= level <= 9:
            heading = text
        unit: Dict[str, Any] = {"loc": f"단락 {n}", "heading": heading, "text": text, "level": level if level <= 9 else 0}
        if with_start:
            unit["start"], unit["end"] = int(rng.Start), int(rng.End)
        units.append(unit)
    return units


WD_PAGE_NUMBER = 3          # wdActiveEndPageNumber: 문서 처음부터 센 실제 쪽 (PDF 뷰어의 쪽과 같음)
OUTLINE_MAX = 200           # outline 한 번에 보여 줄 제목 수
LIST_MAX = 100              # comment/revisions 한 번에 보여 줄 항목 수
_SPEC_RE = re.compile(r"\d+(?:-\d*)?")


def _closed_units(path: str) -> List[Dict[str, Any]]:
    """닫힌 .docx를 reader 단위로 읽어 word.py 단위(no, loc, heading, text, level, page)로 바꿉니다."""
    import reader
    units, _ = reader.parse_docx(path)
    return [{"no": i, "loc": u["loc"], "heading": u["heading"], "text": u["text"],
             "level": len(u["section"] or []) if u["kind"] == "heading" else 0, "page": u["page"]}
            for i, u in enumerate(units, 1)]


def _spec_numbers(spec: str, total: int, option: str) -> List[int]:
    """'120-140', '141-'(끝까지), '7', '1,4,7' -> 번호 목록 (1~total)."""
    parts = [p for p in str(spec).replace(" ", "").split(",") if p]
    if not parts or not all(_SPEC_RE.fullmatch(p) for p in parts):
        raise rt.InputError(f"{option} 형식이 잘못되었습니다: '{spec}' (예: 7, 120-140, 141-, 1,4,7)")
    result: List[int] = []
    for p in parts:
        a, _, b = p.partition("-")
        lo = int(a)
        hi = int(b) if b else (total if "-" in p else lo)
        result += [n for n in range(min(lo, hi), max(lo, hi) + 1) if 1 <= n <= total]
    if not result:
        raise rt.InputError(f"{option} '{spec}'에 해당하는 항목이 없습니다 (전체 1~{total}).")
    return sorted(set(result))


def _compress(nums: Sequence[int]) -> str:
    """[1,2,3,7] -> '1-3,7'"""
    runs: List[List[int]] = []          # [시작, 끝]
    for n in sorted(set(nums)):
        if runs and n == runs[-1][1] + 1:
            runs[-1][1] = n
        else:
            runs.append([n, n])
    return ",".join(str(a) if a == b else f"{a}-{b}" for a, b in runs)


def _page_bounds(doc: Any, pages: List[int], total: int) -> Dict[int, Tuple[int, int]]:
    """쪽마다 (시작 위치, 다음 쪽 시작 위치). GoTo는 쪽 수만큼만 COM을 부릅니다."""
    end = int(doc.Content.End)
    cache: Dict[int, int] = {}

    def start_of(p: int) -> int:
        if p > total:
            return end
        if p not in cache:
            cache[p] = int(doc.GoTo(What=1, Which=1, Count=p).Start)     # wdGoToPage, wdGoToAbsolute
        return cache[p]
    return {p: (start_of(p), start_of(p + 1)) for p in pages}


def _select_pages(units: List[Dict[str, Any]], bounds: Dict[int, Tuple[int, int]]) -> List[Dict[str, Any]]:
    """쪽 범위에 걸친 단락 (앞 쪽에서 이어지는 단락 포함). 각 단락에 page를 붙입니다."""
    import bisect
    starts = [u["start"] for u in units]
    picked: Dict[int, Dict[str, Any]] = {}
    for p, (a, b) in sorted(bounds.items()):
        if a >= b:
            continue
        i = bisect.bisect_right(starts, a) - 1
        if i >= 0 and units[i]["end"] <= a:       # 앞 쪽에서 끝난 단락
            i += 1
        i = max(i, 0)
        while i < len(units) and units[i]["start"] < b:
            picked.setdefault(units[i]["no"], dict(units[i], page=p))
            i += 1
    return [picked[k] for k in sorted(picked)]


def _closed_pages(units: List[Dict[str, Any]], spec: str) -> Tuple[List[Dict[str, Any]], int]:
    known = [u["page"] for u in units if u.get("page")]
    if not known:
        raise rt.InputError("이 파일에는 쪽 정보가 저장되어 있지 않아 --pages로 고를 수 없습니다. "
                            "--paragraphs 또는 --section으로 범위를 지정하세요 (또는 Word로 열어 둔 뒤 다시 실행).")
    total = max(known)
    pages = set(_spec_numbers(spec, total, "--pages"))
    return [u for u in units if u.get("page") in pages], total


def _norm(text: str) -> str:
    return re.sub(r"\s+", "", str(text)).lower()


def _find_section(units: List[Dict[str, Any]], query: str, occurrence: Optional[int]) -> List[Dict[str, Any]]:
    """제목 글(일부)로 장절 하나를 고릅니다: 그 제목부터 같은 수준 이상의 다음 제목 직전까지 (하위 절 포함)."""
    heads = [u for u in units if u.get("level")]
    if not heads:
        raise rt.InputError("이 문서에는 제목(제목 스타일이나 장절 번호)이 없어 --section으로 고를 수 없습니다. "
                            "--pages나 --paragraphs로 범위를 지정하세요.")
    key = _norm(query)
    if not key:
        raise rt.InputError("--section이 비어 있습니다. 제목 글의 일부를 지정하세요 (예: \"1.1\", \"검증 방법\").")
    # 1) 제목 전체가 같음 2) 제목이 그 글로 시작 ('1.1'은 '1.1. Scope'와 맞고 '1.1.1 Data'와는 안 맞음) 3) 제목에 들어 있음
    hits = ([h for h in heads if _norm(h["text"]) == key]
            or [h for h in heads if _norm(h["text"]).startswith(key) and not re.match(r"\.?\d", _norm(h["text"])[len(key):])]
            or [h for h in heads if key in _norm(h["text"])])
    if not hits:
        raise rt.not_found("제목", query, [h["text"][:60] for h in heads],
                           extra=" word.py outline으로 제목 목록을 확인할 수 있습니다.")
    if occurrence is not None:
        if not 1 <= occurrence <= len(hits):
            raise rt.InputError(f"--occurrence는 1~{len(hits)} 사이여야 합니다 ('{query}' 제목 {len(hits)}개).")
        head = hits[occurrence - 1]
    elif len(hits) > 1:
        raise rt.ambiguous(f"'{query}'이(가) 들어간 제목이 {len(hits)}개라 고를 수 없습니다. 사용자에게 번호를 고르게 한 뒤 "
                           "--occurrence 번호를 붙이세요.",
                           [f"{i}: 단락 {h['no']} {_preview(h['text'], 40)}" for i, h in enumerate(hits, 1)])
    else:
        head = hits[0]
    i = units.index(head)
    j = i + 1
    while j < len(units) and not (units[j].get("level") and units[j]["level"] <= head["level"]):
        j += 1
    return units[i:j]


def cmd_read(args) -> Dict[str, Any]:
    if sum(1 for x in (args.pages, args.paragraphs, args.section) if x) > 1:
        raise rt.InputError("--pages, --paragraphs, --section 중 하나만 지정하세요.")
    filtered = bool(args.pages or args.paragraphs or args.section)
    page_total = None
    selection = ""
    closed = _closed_docx(args.target)
    if closed:
        units = _closed_units(closed)
        source = {"app": "Word", "name": os.path.basename(closed), "path": closed, "source": "file(python-docx)"}
        if args.pages:
            selected, page_total = _closed_pages(units, args.pages)
    else:
        with off.open_target("word", args.target, readonly=True) as t:
            units = _com_units(t.doc, with_start=bool(args.pages))
            for i, u in enumerate(units, 1):
                u["no"] = i
            if args.pages:
                page_total = int(t.doc.ComputeStatistics(2))           # wdStatisticPages
                pages = _spec_numbers(args.pages, page_total, "--pages")
                selected = _select_pages(units, _page_bounds(t.doc, pages, page_total))
            source = t.resolved()
    if args.pages:
        selection = f"쪽 {args.pages}"
    elif args.paragraphs:
        if not units:
            raise rt.InputError("문서에 글이 있는 단락이 없습니다.")
        selected = [units[n - 1] for n in _spec_numbers(args.paragraphs, len(units), "--paragraphs")]
        selection = f"단락 {args.paragraphs}"
    elif args.section:
        selected = _find_section(units, args.section, args.occurrence)
        selection = f"장절 '{_preview(selected[0]['text'], 40)}' (단락 {selected[0]['no']}-{selected[-1]['no']})"
    else:
        selected = units[max(1, args.start) - 1:]
    name = source["name"]
    result: Dict[str, Any] = {"target_resolved": source, "units_total": len(units)}
    if filtered:
        result.update(selection=selection, count=len(selected))
    if page_total is not None:
        result["page_total"] = page_total
    if args.save:
        import reader
        save_units = selected if filtered else units
        out = rt.resolve_output_path(None if args.save is True else args.save, os.path.splitext(name)[0], ".md")
        out = rt.write_text(out, reader.units_to_markdown(name, save_units))
        result.update(output_file=out, next_commands=[
            rt.next_command("ppt.py", ["write", "--file", out], "저장한 내용을 열린 PPT에 넣기"),
            rt.next_command("office.py", ["transfer", "--from", "word", "--what", "text", "--to", "excel"],
                            "이 문서의 단락을 엑셀 표로 옮기기")])
        return rt.ok_file(TOOL, "read", f"{name}: 단락 {len(save_units)}개를 저장했습니다 ({out})", out,
                          "VERIFIED: 저장한 파일을 다시 열어 확인했습니다. 다른 문서에 넣으려면 next_commands를 그대로 실행하세요.",
                          **result)
    if filtered and not selected:
        result.update(paragraphs=[], truncated=False)
        return rt.ok(TOOL, "read", f"{name}: {selection}에 글이 있는 단락 없음 (count 0)",
                     f"VERIFIED: 0건. {selection}에는 글이 있는 단락이 없습니다 (표·그림만 있거나 빈 쪽). 사용자에게 그대로 알리세요.",
                     **result)
    shown: List[Dict[str, Any]] = []
    used = 0
    text_cut = 0
    for u in selected:
        size = len(u["text"]) + 12
        if shown and used + size > rt.DEFAULT_MAX_CHARS:
            break
        text, cut = rt.clamp_text(u["text"], rt.DEFAULT_MAX_CHARS)
        item = {"no": u["no"], "heading": u["heading"], "text": text}
        if cut:                                  # 아주 긴 단락: 앞부분만 (전체는 --save 파일)
            item.update(text_truncated=True, text_chars=len(u["text"]))
            text_cut += 1
        if args.pages and u.get("page") is not None:
            item["page"] = u["page"]
        shown.append(item)
        used += size
    truncated = len(shown) < len(selected) or bool(text_cut)
    result.update(paragraphs=shown, truncated=truncated)
    if text_cut:
        result["truncated_reason"] = f"{rt.DEFAULT_MAX_CHARS}자보다 긴 단락 {text_cut}개는 앞부분만 표시 (text_truncated, 전체는 --save)"
    if truncated and len(shown) == len(selected):
        # 단락 글만 잘림: 이어 읽을 단락은 없고 저장 파일에 전체가 있음
        save_args = ["read", "--target", args.target]
        if args.section:
            save_args += ["--section", args.section] + (["--occurrence", str(args.occurrence)] if args.occurrence else [])
        elif args.pages:
            save_args += ["--pages", args.pages]
        elif args.paragraphs:
            save_args += ["--paragraphs", args.paragraphs]
        result["next_commands"] = [rt.next_command("word.py", save_args + ["--save"], "잘린 단락까지 전체를 md 파일로 저장")]
    elif truncated:
        rest = selected[len(shown):]
        if not filtered:
            nxt = shown[-1]["no"] + 1
            result["next_commands"] = [
                rt.next_command("word.py", ["read", "--target", args.target, "--start", str(nxt)], f"{nxt}번째 단락부터 이어서 읽기"),
                rt.next_command("word.py", ["read", "--target", args.target, "--save"], "전체를 md 파일로 저장")]
        else:
            nos = [u["no"] for u in rest]
            if args.pages and nos != list(range(nos[0], nos[-1] + 1)):
                spec, opt = _compress([u["page"] for u in rest]), "--pages"    # 떨어진 쪽: 남은 쪽으로 이어 읽기
            else:
                spec, opt = _compress(nos), "--paragraphs"
            save_args = ["read", "--target", args.target]
            if args.section:
                save_args += ["--section", args.section] + (["--occurrence", str(args.occurrence)] if args.occurrence else [])
            else:
                save_args += ["--pages", args.pages] if args.pages else ["--paragraphs", args.paragraphs]
            result["next_commands"] = [
                rt.next_command("word.py", ["read", "--target", args.target, opt, spec], "나머지 이어서 읽기"),
                rt.next_command("word.py", save_args + ["--save"], "고른 범위 전체를 md 파일로 저장")]
    what = selection or "전체"
    preview = f"{name}: {what}, 단락 {len(shown)}개 표시 (" + (f"고른 범위 {len(selected)}개, " if filtered else "") + f"전체 {len(units)}개)"
    return rt.ok(TOOL, "read", preview, "VERIFIED: 읽기 완료. 일부만 표시되었으면 next_commands로 이어서 읽으세요."
                 if truncated else "", **result)


def cmd_info(args) -> Dict[str, Any]:
    if args.all:
        docs = [{"name": d["name"], "path": d["path"], "saved": d["saved"], "active": d["active"]}
                for d in off.list_open("word")]
        return rt.ok(TOOL, "info", f"열린 문서 {len(docs)}개", documents=docs)
    with off.open_target("word", args.target, readonly=True) as t:
        doc = t.doc
        stats = {}
        for key, code in (("pages", 2), ("words", 0), ("characters", 3), ("paragraphs", 4)):
            try:
                stats[key] = int(doc.ComputeStatistics(code))
            except Exception:
                pass
        tables = []
        for i in range(1, doc.Tables.Count + 1):
            try:
                tb = doc.Tables(i)
                tables.append({"no": i, "rows": int(tb.Rows.Count), "cols": int(tb.Columns.Count)})
            except Exception:
                tables.append({"no": i})
        headings = [u["text"] for u in _com_units(doc) if u.get("level")][:30]
        result = {"target_resolved": t.resolved(), "stats": stats, "tables": tables,
                  "pictures": int(doc.InlineShapes.Count), "floating_shapes": int(doc.Shapes.Count),
                  "headings": headings, "saved": off.is_saved(doc)}
        if args.detail:
            result.update(_layout_detail(doc))
    preview = (f"{result['target_resolved']['name']}: {stats.get('pages', '?')}쪽, 글자 {stats.get('characters', '?')}, "
               f"표 {len(tables)}개, 그림 {result['pictures']}개, 제목 {len(headings)}개")
    if args.detail:
        n = len(result["warnings"])
        preview += f", 배치 경고 {n}건" if n else ", 배치 경고 없음"
    return rt.ok(TOOL, "info", preview, **result)


def _layout_detail(doc: Any) -> Dict[str, Any]:
    """본문 그림 크기(쪽 너비 대비)와 제목 수준. 쪽 너비를 넘는 그림은 경고."""
    warnings: List[str] = []
    try:
        ps = doc.PageSetup
        usable = float(ps.PageWidth) - float(ps.LeftMargin) - float(ps.RightMargin)
    except Exception:
        usable = None
    pictures = []
    for i in range(1, doc.InlineShapes.Count + 1):
        try:
            s = doc.InlineShapes(i)
            w, h = float(s.Width), float(s.Height)
            item: Dict[str, Any] = {"no": i, "width": round(w, 1), "height": round(h, 1),
                                    "ratio": round(w / h, 3) if h else None}
            if usable is not None and w > usable + 0.5:
                item["wider_than_page"] = True
                warnings.append(f"그림 {i}이(가) 쪽 너비({usable:.0f}pt)보다 넓습니다 ({w:.0f}pt)")
            pictures.append(item)
        except Exception:
            pictures.append({"no": i})
    outline = [{"level": u["level"], "text": u["text"][:60]} for u in _com_units(doc) if u.get("level")][:30]
    return {"usable_width": round(usable, 1) if usable is not None else None, "picture_detail": pictures,
            "heading_levels": outline, "warnings": warnings, "layout_warnings": len(warnings)}


def cmd_outline(args) -> Dict[str, Any]:
    """제목 구조(목차): 수준, 단락 번호(read --paragraphs와 같은 번호), 쪽."""
    start = max(1, args.start)
    closed = _closed_docx(args.target)
    if closed:
        heads = [u for u in _closed_units(closed) if u["level"]]
        window = heads[start - 1:start - 1 + OUTLINE_MAX]
        source = {"app": "Word", "name": os.path.basename(closed), "path": closed, "source": "file(python-docx)"}
    else:
        with off.open_target("word", args.target, readonly=True) as t:
            units = _com_units(t.doc, with_start=True)
            for i, u in enumerate(units, 1):
                u["no"] = i
            heads = [u for u in units if u["level"]]
            window = heads[start - 1:start - 1 + OUTLINE_MAX]
            for h in window:                    # 쪽 계산은 보여 줄 제목만 (제목마다 COM 1회)
                try:
                    h["page"] = int(t.doc.Range(h["start"], h["start"]).Information(WD_PAGE_NUMBER))
                except Exception:
                    h["page"] = None
            source = t.resolved()
    name = source["name"]
    result: Dict[str, Any] = {"target_resolved": source, "count": len(heads)}
    if not heads:
        result.update(headings=[], truncated=False)
        return rt.ok(TOOL, "outline", f"{name}: 제목 없음 (count 0)",
                     "VERIFIED: 0건. 제목(제목 스타일이나 장절 번호)이 있는 단락이 없습니다. 사용자에게 그대로 알리고, "
                     "본문은 word.py read --pages 또는 --paragraphs로 읽으세요.", **result)
    shown: List[Dict[str, Any]] = []
    used = 0
    for h in window:
        item = {"paragraph": h["no"], "level": h["level"], "text": _preview(h["text"], 120), "page": h.get("page")}
        size = len(item["text"]) + 30
        if shown and used + size > rt.DEFAULT_MAX_CHARS:
            break
        shown.append(item)
        used += size
    shown_end = start - 1 + len(shown)
    truncated = shown_end < len(heads)
    result.update(headings=shown, truncated=truncated)
    if truncated:
        result["next_commands"] = [rt.next_command("word.py", ["outline", "--target", args.target, "--start", str(shown_end + 1)],
                                                   f"{shown_end + 1}번째 제목부터 이어서 보기")]
    preview = f"{name}: 제목 {len(shown)}개 표시 (전체 {len(heads)}개)"
    return rt.ok(TOOL, "outline", preview,
                 "VERIFIED: 목차 읽기 완료. 한 장절의 본문은 word.py read --section \"<제목 일부>\"로 읽으세요."
                 + (" 일부만 표시되었으면 next_commands로 이어서 보세요." if truncated else ""), **result)


def _com_date(value: Any) -> str:
    try:
        return value.strftime("%Y-%m-%d %H:%M")
    except Exception:
        return str(value or "")


def _page_of(rng: Any) -> Optional[int]:
    try:
        return int(rng.Information(WD_PAGE_NUMBER))
    except Exception:
        return None


def _clean(text: Any) -> str:
    return str(text or "").replace("\r", "\n").replace("\x07", " ").replace("\x0b", "\n").strip()


def _window(total: int, start: int) -> range:
    first = max(1, start)
    return range(first, min(total, first + LIST_MAX - 1) + 1)


def _list_result(name: str, command: str, kind: str, items: List[Dict[str, Any]], total: int, start: int,
                 target: str, **fields: Any) -> Dict[str, Any]:
    """comment/revisions 목록 응답: 글자 수 상한, 0건 안내, 이어 보기 명령."""
    shown: List[Dict[str, Any]] = []
    used = 0
    for it in items:
        size = sum(len(str(v)) for v in it.values()) + 20
        if shown and used + size > rt.DEFAULT_MAX_CHARS:
            break
        shown.append(it)
        used += size
    shown_end = max(1, start) - 1 + len(shown)
    truncated = shown_end < total
    result: Dict[str, Any] = {"count": total, "items": shown, "truncated": truncated}
    result.update(fields)
    if not total:
        return rt.ok(TOOL, command, f"{name}: {kind} 없음 (count 0)",
                     f"VERIFIED: 0건. 이 문서에는 {kind}이(가) 없습니다. 사용자에게 그대로 알리세요.", **result)
    if truncated:
        result["next_commands"] = [rt.next_command("word.py", [command, "--target", target, "--start", str(shown_end + 1)],
                                                   f"{shown_end + 1}번째부터 이어서 보기")]
    return rt.ok(TOOL, command, f"{name}: {kind} {len(shown)}개 표시 (전체 {total}개)",
                 "VERIFIED: 읽기 완료." + (" 일부만 표시되었으면 next_commands로 이어서 보세요." if truncated else ""), **result)


def _read_comments(doc: Any, numbers: Sequence[int]) -> List[Dict[str, Any]]:
    items = []
    for i in numbers:
        c = doc.Comments(i)
        item: Dict[str, Any] = {"no": i, "author": str(getattr(c, "Author", "") or ""), "date": _com_date(getattr(c, "Date", ""))}
        try:
            item["page"] = _page_of(c.Scope)
            item["target_text"] = _preview(str(c.Scope.Text), 80)
        except Exception:
            item["page"], item["target_text"] = None, ""
        item["text"] = rt.clamp_text(_clean(c.Range.Text), 1000)[0]
        try:
            parent = c.Ancestor
            if parent is not None:
                item["reply_to"] = int(parent.Index)
        except Exception:
            pass
        try:
            item["done"] = bool(c.Done)
        except Exception:
            pass
        items.append(item)
    return items


def _memo_rows(path: str) -> List[Tuple[str, str, Optional[int]]]:
    """메모 CSV: 기준 글, 메모 내용[, 몇 번째 단락]. 첫 행이 머리글이면 건너뜁니다."""
    rows = [r for r in rt.read_csv_rows(path) if any(str(v).strip() for v in r)]
    if rows and str(rows[0][0]).strip().lower() in ("anchor", "기준", "기준 글", "앵커", "위치", "대상", "대상 글"):
        rows = rows[1:]
    out: List[Tuple[str, str, Optional[int]]] = []
    for n, r in enumerate(rows, 1):
        anchor = str(r[0]).strip() if r else ""
        memo = str(r[1]).strip() if len(r) > 1 else ""
        occ_raw = str(r[2]).strip() if len(r) > 2 else ""
        if not anchor or not memo:
            raise rt.InputError(f"메모 파일 {n}번째 줄에 기준 글 또는 메모 내용이 없습니다. 열: 기준 글, 메모 내용[, 몇 번째 단락].")
        if occ_raw and not occ_raw.isdigit():
            raise rt.InputError(f"메모 파일 {n}번째 줄의 3번째 열(몇 번째 단락)은 숫자여야 합니다: '{occ_raw}'")
        out.append((anchor, memo, int(occ_raw) if occ_raw else None))
    if not out:
        raise rt.InputError("메모 파일에 내용이 없습니다. 열: 기준 글, 메모 내용[, 몇 번째 단락].")
    return out


def _comment_range(doc: Any, anchor: str, occurrence: Optional[int]) -> Any:
    """메모를 달 범위: 기준 글이 든 단락을 고르고(여러 곳이면 [후보]) 그 안의 기준 글 자체."""
    if _find_len(anchor) > FIND_MAX:
        raise rt.InputError(f"--anchor는 {FIND_MAX}자 이하여야 합니다(^는 2자로 셈). 앞부분 일부만 쓰세요.")
    para = _anchor_paragraph(doc, anchor, occurrence, "after")
    rng = doc.Range(int(para.Start), int(para.End))
    return rng if _find(rng, anchor) else para


def _comment_add(args, t: Any) -> Dict[str, Any]:
    doc = t.doc
    if args.file and os.path.splitext(args.file)[1].lower() in (".csv", ".tsv"):
        rows = _memo_rows(rt.resolve_input_path(args.file, "메모 파일"))
    else:
        if not args.anchor:
            raise rt.InputError("--anchor로 메모를 달 글을 지정하세요 (여러 개면 --file \"메모.csv\": 기준 글, 메모 내용).")
        if args.file:
            memo = rt.read_text_file(rt.resolve_input_path(args.file, "메모 파일")).strip()
        else:
            memo = rt.unescape_newlines(args.text or "").strip()
        if not memo:
            raise rt.InputError("메모 내용이 없습니다. --text \"메모\" 또는 --file \"메모.txt\"를 지정하세요.")
        rows = [(args.anchor, memo, args.occurrence)]
    if len(rows) > 1:                               # 여러 개: 전부 찾을 수 있을 때만 단다 (일부만 달리지 않게)
        problems = []
        for n, (anchor, _memo, occ) in enumerate(rows, 1):
            try:
                _comment_range(doc, anchor, occ)
            except rt.InputError as e:
                msg = str(e).replace("[추천]", "추천:").replace("[후보]", "후보:")
                problems.append(f"{n}번째 줄 '{_preview(anchor, 20)}': {msg}")
        if problems:
            raise rt.InputError(f"메모 {len(rows)}개 중 {len(problems)}개의 기준 글을 정할 수 없어 하나도 달지 않았습니다. "
                                "메모 파일의 기준 글을 고치거나 3번째 열에 몇 번째 단락인지 적은 뒤 다시 실행하세요. "
                                + " / ".join(problems[:10]))
    before = int(doc.Comments.Count)
    added = []
    for anchor, memo, occ in rows:
        rng = _comment_range(doc, anchor, occ)
        doc.Comments.Add(Range=rng, Text=memo.replace("\n", "\r"))
        added.append({"anchor": _preview(anchor, 40), "text": _preview(memo, 80), "page": _page_of(rng)})
    t.set_focus()
    after = int(doc.Comments.Count)
    texts = {_norm(_clean(doc.Comments(i).Range.Text)) for i in range(1, after + 1)}
    missing = [a["anchor"] for (_, memo, _), a in zip(rows, added) if _norm(memo) not in texts]
    resolved = t.resolved()
    fields = {"target_resolved": resolved, "added": added, "count": len(added), "comments_total": after}
    preview = f"{resolved['name']}: 메모 {len(added)}개를 달았습니다 (문서 전체 메모 {after}개)"
    if after - before != len(added) or missing:
        return rt.ok_unverified(TOOL, "comment", preview,
                                [f"메모 수가 {before} -> {after}로 기대({before + len(added)})와 다름"] if after - before != len(added)
                                else [f"다시 읽은 메모에서 찾지 못함: {', '.join(missing[:5])}"], **fields)
    return rt.ok(TOOL, "comment", preview,
                 "VERIFIED: 메모를 달고 다시 읽어 확인했습니다. 잘못 달았으면 office.py undo로 되돌릴 수 있습니다.",
                 verified=True, **fields)


def cmd_comment(args) -> Dict[str, Any]:
    # --action add를 빠뜨리고 메모 내용만 준 경우: 목록을 보여 주고 "완료"라고 하면 LLM이 메모를 단 것으로 오해합니다
    if args.action == "add" or args.anchor or args.text is not None or args.file:
        with off.open_target("word", args.target, readonly=False, focus=not args.no_focus,
                             tool=TOOL, command="comment") as t:
            return _comment_add(args, t)
    with off.open_target("word", args.target, readonly=True) as t:
        total = int(t.doc.Comments.Count)
        items = _read_comments(t.doc, _window(total, args.start))
        resolved = t.resolved()
    return _list_result(resolved["name"], "comment", "검토 메모", items, total, args.start, args.target,
                        target_resolved=resolved)


WD_REVISION = {1: "삽입", 2: "삭제", 3: "서식 변경", 4: "단락 번호 변경", 5: "필드 표시 변경", 6: "조정", 7: "충돌",
               8: "스타일 변경", 9: "바꾸기", 10: "단락 서식 변경", 11: "표 서식 변경", 12: "구역 서식 변경",
               13: "스타일 정의 변경", 14: "이동(원래 자리)", 15: "이동(새 자리)", 16: "셀 삽입", 17: "셀 삭제",
               18: "셀 병합", 19: "셀 분할", 20: "충돌 삽입", 21: "충돌 삭제"}


def cmd_revisions(args) -> Dict[str, Any]:
    with off.open_target("word", args.target, readonly=True) as t:
        doc = t.doc
        total = int(doc.Revisions.Count)
        items = []
        for i in _window(total, args.start):
            r = doc.Revisions(i)
            kind = int(getattr(r, "Type", 0) or 0)
            item: Dict[str, Any] = {"no": i, "type": WD_REVISION.get(kind, f"기타({kind})"),
                                    "author": str(getattr(r, "Author", "") or ""), "date": _com_date(getattr(r, "Date", ""))}
            try:
                item["page"] = _page_of(r.Range)
                item["text"] = rt.clamp_text(_clean(r.Range.Text), 300)[0]
            except Exception:
                item["page"], item["text"] = None, ""
            items.append(item)
        try:
            tracking = bool(doc.TrackRevisions)
        except Exception:
            tracking = None
        resolved = t.resolved()
    return _list_result(resolved["name"], "revisions", "변경 내용(변경 추적)", items, total, args.start, args.target,
                        target_resolved=resolved, track_changes=tracking)



# ==============================================================================
# 2. write
# ==============================================================================

class Writer:
    """문서의 한 위치(pos)부터 블록을 차례로 넣는 도우미. 넣은 뒤 pos는 다음 삽입 위치로 이동합니다."""

    def __init__(self, doc: Any, pos: int, theme: Optional[Dict[str, Any]]):
        self.doc = doc
        self.pos = pos
        self.theme = theme or off.THEMES["navy"] | {"name": "navy"}
        self.created: List[str] = []

    def _para(self, text: str, style: Optional[int] = None, bold_spans: Sequence[Tuple[int, int]] = ()) -> Any:
        r = self.doc.Range(self.pos, self.pos)
        r.InsertAfter(text + "\r")
        para = self.doc.Range(r.Start, r.Start + len(text))
        try:
            para.Style = self.doc.Styles(style if style is not None else WD_STYLE_NORMAL)
        except Exception:
            pass
        para.Font.Bold = False
        for a, b in bold_spans:
            self.doc.Range(r.Start + a, r.Start + b).Font.Bold = True
        self.pos = r.End
        return para

    def text(self, text: str, style: Optional[int] = None) -> None:
        plain, spans = rt.split_bold(text)
        self._para(plain, style, spans)

    def heading(self, level: int, text: str) -> None:
        self.text(text, heading_style(level))
        self.created.append(f"제목{level}")

    def items(self, items: List[str], numbered: bool) -> None:
        for it in items:
            self.text(it, WD_STYLE_LIST_NUMBER if numbered else WD_STYLE_LIST_BULLET)
        self.created.append(f"{'번호' if numbered else '글머리'} 목록 {len(items)}개")

    def table(self, rows: List[List[Any]]) -> Any:
        width = max(len(r) for r in rows)
        lines = ["\t".join(str("" if v is None else v).replace("\t", " ").replace("\r", " ").replace("\n", " ")
                           for v in (list(r) + [""] * (width - len(r)))) for r in rows]
        r = self.doc.Range(self.pos, self.pos)
        r.InsertAfter("\r".join(lines) + "\r")
        body = self.doc.Range(r.Start, r.End - 1)
        body.Style = self.doc.Styles(WD_STYLE_NORMAL)
        tbl = body.ConvertToTable(Separator=1, NumRows=len(rows), NumColumns=width)   # 탭 구분 -> 표 (셀 단위 COM 호출 없음)
        tbl.Borders.Enable = True
        try:
            tbl.AutoFitBehavior(2)            # 창 너비에 맞춤
        except Exception:
            pass
        head = tbl.Rows(1)
        head.Range.Font.Bold = True
        head.Shading.BackgroundPatternColor = off.parse_color(self.theme["header_bg"])
        head.Range.Font.Color = off.parse_color(self.theme["header_fg"])
        try:
            head.HeadingFormat = True          # 페이지가 넘어가면 머리글 행 반복
        except Exception:
            pass
        self.pos = tbl.Range.End
        if self.pos >= self.doc.Content.End - 1:     # 문서 끝 표 뒤에는 빈 단락이 하나 있음
            self.pos = self.doc.Content.End - 1
        self.created.append(f"표 {len(rows)}x{width}")
        return tbl

    def callout(self, text: str) -> None:
        r = self.doc.Range(self.pos, self.pos)
        r.InsertAfter(text.replace("\n", "\v") + "\r")
        body = self.doc.Range(r.Start, r.End - 1)
        tbl = body.ConvertToTable(Separator=1, NumRows=1, NumColumns=1)
        cell = tbl.Cell(1, 1)
        cell.Shading.BackgroundPatternColor = off.parse_color(self.theme["light"])
        tbl.Borders.Enable = False
        left = tbl.Borders(-2)                     # wdBorderLeft
        left.LineStyle = 1
        left.LineWidth = 24                        # 3pt
        left.Color = off.parse_color(self.theme["primary"])
        self.pos = tbl.Range.End
        if self.pos >= self.doc.Content.End - 1:
            self.pos = self.doc.Content.End - 1
        self.created.append("요약 상자")

    def image(self, path: str, width: Optional[float]) -> None:
        r = self.doc.Range(self.pos, self.pos)
        r.InsertAfter("\r")
        anchor = self.doc.Range(r.Start, r.Start)
        pic = self.doc.InlineShapes.AddPicture(FileName=path, LinkToFile=False, SaveWithDocument=True, Range=anchor)
        setup = self.doc.PageSetup
        usable = float(setup.PageWidth - setup.LeftMargin - setup.RightMargin)
        target_w = min(width or usable, usable)
        w0, h0 = float(pic.Width), float(pic.Height)
        if w0 > 0 and (w0 > target_w or width):
            ratio = target_w / w0                     # 가로세로 비율 유지
            pic.Width = target_w
            pic.Height = h0 * ratio
        try:
            pic.Range.ParagraphFormat.Alignment = 1   # 가운데
        except Exception:
            pass
        self.pos = pic.Range.End + 1
        self.created.append(f"그림 {os.path.basename(path)}")

    def toc(self) -> None:
        r = self.doc.Range(self.pos, self.pos)
        r.InsertAfter("\r")
        anchor = self.doc.Range(r.Start, r.Start)
        toc = self.doc.TablesOfContents.Add(Range=anchor, UseHeadingStyles=True, UpperHeadingLevel=1, LowerHeadingLevel=3)
        toc.Update()
        self.pos = toc.Range.End + 1
        self.created.append("목차")

    def page_break(self) -> None:
        r = self.doc.Range(self.pos, self.pos)
        r.InsertBreak(7)                              # wdPageBreak
        self.pos = r.End
        self.created.append("페이지 나누기")

    def markdown(self, text: str) -> None:
        for kind, payload in rt.parse_markdown_blocks(text):
            if kind == "heading":
                self.heading(*payload)
            elif kind == "paragraph":
                self.text(payload)
                if not self.created or not self.created[-1].startswith("본문"):
                    self.created.append("본문")
            elif kind in ("bullets", "numbered"):
                self.items(payload, kind == "numbered")
            elif kind == "table":
                self.table(payload)
            elif kind == "break":
                self.page_break()


def _find(rng: Any, text: str, match_case: bool = False, whole_word: bool = False) -> bool:
    """Find 옵션을 모두 명시합니다 (사용자가 마지막에 쓴 와일드카드 등 설정이 남아 오동작하는 것 방지)."""
    f = rng.Find
    f.ClearFormatting()
    # Word는 와일드카드가 꺼져 있어도 ^를 특수 코드로 읽음 (^^ -> ^, ^p 단락, ^6은 COM 오류): 글자 그대로 찾게 ^^로 바꿈
    return bool(f.Execute(FindText=text.replace("^", "^^"), MatchCase=match_case, MatchWholeWord=whole_word, MatchWildcards=False,
                          MatchSoundsLike=False, MatchAllWordForms=False, Forward=True, Wrap=0, Format=False))


ANCHOR_MAX_HITS = 50     # 앵커가 들어간 단락을 이 개수까지만 셉니다
_CONTROL = re.compile(r"[\x00-\x1f]")


def _preview(text: str, n: int = 30) -> str:
    t = re.sub(r"\s+", " ", _CONTROL.sub(" ", text)).strip().replace(",", " ")
    return t if len(t) <= n else t[:n] + "..."


def _anchor_paragraphs(doc: Any, anchor: str) -> List[Any]:
    """앵커 글이 들어간 단락 Range 목록 (문서 순서, 같은 단락은 한 번)."""
    found: List[Any] = []
    seen = set()
    pos, end = 0, int(doc.Content.End)
    while len(found) < ANCHOR_MAX_HITS and pos < end:
        rng = doc.Range(pos, end)
        if not _find(rng, anchor):
            break
        para = rng.Paragraphs(1).Range
        if int(para.Start) not in seen:
            seen.add(int(para.Start))
            found.append(para)
        pos = max(int(rng.End), pos + 1)
    return found


def _similar_anchor(doc: Any, anchor: str) -> Optional[str]:
    """찾지 못한 앵커와 가장 비슷한 문서 속 글 (그대로 다시 찾을 수 있는 원문 일부). 충분히 비슷하지 않으면 None."""
    key = anchor.strip().lower()
    if len(key) < 2:
        return None
    paras = [p for p in str(doc.Content.Text).split("\r") if p.strip()][:5000]
    scored = []
    for p in paras:
        m = difflib.SequenceMatcher(None, key, p.lower(), autojunk=False)
        size = m.find_longest_match(0, len(key), 0, len(p)).size
        if size:
            scored.append((size, p))
    best, best_text = 0.0, None
    for _, p in sorted(scored, key=lambda x: -x[0])[:5]:
        low = p.lower()
        for width in (len(key) - 1, len(key), len(key) + 1):
            if width < 2:
                continue
            for i in range(0, max(1, len(p) - width + 1)):
                piece = p[i:i + width]
                if _CONTROL.search(piece):
                    continue
                r = difflib.SequenceMatcher(None, key, low[i:i + width], autojunk=False).ratio()
                if r > best:
                    best, best_text = r, piece.strip()
    return best_text if best >= 0.6 and best_text and best_text.lower() != key else None


def _safe_anchor_fix(anchor: str, similar: str, pos: str) -> bool:
    """비슷한 글을 [추천](LLM이 확인 없이 1회 재시도)으로 줘도 되는지.
    공백·대소문자만 다르면 항상, 앵커 글자 사이에 2자 이하만 끼어든 경우(예: '2026 계획' -> '2026년 계획')는
    단락을 지우지 않는 위치에서만. 글자가 바뀌거나 빠진 경우('Q2' -> 'Q1')는 뜻이 다를 수 있어 주지 않음."""
    a = re.sub(r"\s+", "", anchor).lower()
    b = re.sub(r"\s+", "", similar).lower()
    if a == b:
        return True
    if pos == "replace" or not 0 < len(b) - len(a) <= 2:
        return False
    it = iter(b)
    return all(ch in it for ch in a)            # a가 b의 부분 수열 (끼어든 글자만 있음)


def _anchor_paragraph(doc: Any, anchor: str, occurrence: Optional[int], pos: str = "after") -> Any:
    """앵커 단락 하나. 못 찾으면 비슷한 글을 안내 (확실할 때만 [추천]), 여러 단락에 있으면 고르지 않고 [후보]로 오류."""
    hits = _anchor_paragraphs(doc, anchor)
    if not hits:
        similar = _similar_anchor(doc, anchor)
        hint = ""
        if similar and _safe_anchor_fix(anchor, similar, pos):
            hint = f" [추천] {similar}"
        elif similar:
            hint = f" 비슷한 문구: '{similar}' (뜻이 같은지 사용자에게 확인한 뒤에만 쓰세요)"
        raise rt.InputError(f"문서에서 '{anchor}'을(를) 찾을 수 없습니다. word.py read로 정확한 문구를 확인하세요." + hint,
                            no_fallback=True)
    if occurrence is not None:
        if not 1 <= occurrence <= len(hits):
            raise rt.InputError(f"--occurrence는 1~{len(hits)} 사이여야 합니다 ('{anchor}'이(가) 있는 단락 {len(hits)}개).",
                                no_fallback=True)
        return hits[occurrence - 1]
    if len(hits) > 1:
        more = "개 이상" if len(hits) >= ANCHOR_MAX_HITS else "개"
        raise rt.ambiguous(
            f"'{anchor}'이(가) 단락 {len(hits)}{more}에 있어 어느 곳인지 고를 수 없습니다. 사용자에게 번호를 고르게 한 뒤 "
            "--occurrence 번호를 붙이거나, 한 곳에만 있는 더 긴 문구를 --anchor로 쓰세요.",
            [f"{i}: {_preview(str(h.Text))}" for i, h in enumerate(hits, 1)])
    return hits[0]


def _insert_position(doc: Any, anchor: Optional[str], pos: str, occurrence: Optional[int] = None) -> Tuple[int, str]:
    """삽입 위치(문자 오프셋)와 설명. 앵커가 있으면 그 단락의 앞/뒤, replace면 앵커 단락을 지우고 그 자리.
    앵커가 여러 단락에 있으면 occurrence(1부터)로 고른 단락, 없으면 오류 (첫 번째를 임의로 고르지 않음)."""
    if not anchor:
        end = doc.Content.End - 1
        last = doc.Range(end, end).Paragraphs(1).Range
        if pos == "start":
            if _in_table(doc.Range(0, 0)):
                raise rt.InputError("문서가 표로 시작해 맨 앞에 넣으면 첫 칸 안에 들어갑니다. 표 뒤 단락의 글을 --anchor로 지정해 "
                                    "--pos before로 넣거나, Word에서 표 앞에 빈 줄을 만든 뒤 다시 실행하세요.", no_fallback=True)
            return 0, "문서 처음"
        if str(last.Text).strip():                    # 마지막 단락에 내용이 있으면 새 단락에서 시작
            doc.Range(end, end).InsertAfter("\r")
            end = doc.Content.End - 1
        return end, "문서 끝"
    if _find_len(anchor) > FIND_MAX:
        raise rt.InputError(f"--anchor는 {FIND_MAX}자 이하여야 합니다(^는 2자로 셈). 앵커 단락의 앞부분 일부만 쓰세요.")
    para = _anchor_paragraph(doc, anchor, occurrence, pos)
    if _in_table(para):
        # 셀 단락의 End는 다음 셀의 시작이라 그대로 넣으면 옆 칸에 들어감: 표 바깥(표 뒤)에 넣습니다
        if pos != "after":
            raise rt.InputError(f"'{anchor}'은(는) 표 안에 있어 --pos {pos}로 넣으면 표 칸 안에 들어갑니다. --pos after(표 뒤에 넣기)를 쓰거나, "
                                "칸 안의 글만 바꾸려면 word.py replace를 쓰세요.", no_fallback=True)
        return int(para.Tables(1).Range.End), f"'{anchor}'이(가) 있는 표 뒤"
    if pos == "before":
        return int(para.Start), f"'{anchor}' 단락 앞"
    if pos == "replace":
        start = int(para.Start)
        if int(para.End) < int(doc.Content.End):
            para.Delete()                                  # 단락 전체(단락 기호 포함)를 지우고 그 자리에 넣음
        else:
            doc.Range(para.Start, para.End - 1).Delete()  # 마지막 단락은 단락 기호를 남김
        return start, f"'{anchor}' 단락 자리"
    if int(para.End) >= int(doc.Content.End):
        # 마지막 단락 뒤에는 글을 둘 수 없어 그대로 넣으면 앵커 단락에 붙고 스타일이 바뀜: 새 단락을 먼저 만듦
        end = int(doc.Content.End) - 1
        doc.Range(end, end).InsertAfter("\r")
        return int(doc.Content.End) - 1, f"'{anchor}' 단락 뒤"
    return int(para.End), f"'{anchor}' 단락 뒤"


def _in_table(rng: Any) -> bool:
    try:
        return bool(rng.Information(12))          # wdWithInTable
    except Exception:
        return False


def cmd_write(args) -> Dict[str, Any]:
    if not any([args.text, args.file, args.image, args.toc, args.page_break]):
        raise rt.InputError("넣을 내용이 없습니다. --text, --file, --image, --toc, --page-break 중 하나 이상을 지정하세요.")
    theme = off.get_theme(args.theme)
    text = rt.unescape_newlines(args.text) if args.text else None
    file_rows = None
    file_text = None
    if args.file:
        path = rt.resolve_input_path(args.file, "입력 파일")
        if os.path.splitext(path)[1].lower() in (".csv", ".tsv"):
            file_rows = rt.read_csv_rows(path)
        else:
            file_text = rt.read_text_file(path)
    image = rt.resolve_input_path(args.image, "그림 파일") if args.image else None

    with off.open_target("word", args.target, readonly=False, focus=not args.no_focus, create=True,
                         tool=TOOL, command="write") as t:
        doc = t.doc
        start, where = _insert_position(doc, args.anchor, args.pos, args.occurrence)
        w = Writer(doc, start, theme)
        if args.page_break:
            w.page_break()
        if args.toc:
            w.toc()
        for content in (text, file_text):
            if content is None:
                continue
            if args.callout:
                w.callout(content.strip())
            elif args.heading:
                w.heading(args.heading, content.strip())
            elif args.title:
                w.text(content.strip(), WD_STYLE_TITLE)
                w.created.append("문서 제목")
            elif rt.looks_like_markdown(content):
                w.markdown(content)
            else:
                for para in [p for p in re.split(r"\n\s*\n|\n", content) if p.strip()]:
                    w.text(para.strip())
                w.created.append("본문")
        if file_rows:
            w.table(file_rows)
        if image:
            w.image(image, args.width)
        problems = _verify_inserted(doc, start, w.pos, [c for c in (text, file_text) if c])
        t.set_focus()
        if t.source == "active" and not args.no_focus:
            try:
                doc.Range(w.pos, w.pos).Select()          # 커서를 넣은 내용 끝으로
            except Exception:
                pass
        resolved = t.resolved()
    preview = f"{resolved['name']}의 {where}에 넣었습니다: {', '.join(w.created)}"
    if problems:
        return rt.ok_unverified(TOOL, "write", preview, problems, target_resolved=resolved, inserted=w.created)
    return rt.ok(TOOL, "write", preview, "VERIFIED: 넣은 범위를 다시 읽어 확인했습니다. 같은 명령을 다시 실행하지 마세요.",
                 verified=True, target_resolved=resolved, inserted=w.created)


def _verify_inserted(doc: Any, start: int, end: int, contents: List[str]) -> List[str]:
    """R1: 넣은 범위를 다시 읽어, 범위가 비어 있지 않고 넣은 글의 첫 줄이 들어 있는지 확인합니다 (마크다운 기호 제외)."""
    try:
        got = re.sub(r"[\s\x00-\x1f]+", "", str(doc.Range(start, end).Text))   # 표 칸 구분 기호도 무시
    except Exception as e:
        return [f"넣은 범위를 다시 읽지 못함: {type(e).__name__}"]
    if end <= start:
        return ["넣은 범위가 비어 있음"]
    problems = []
    for content in contents:
        first = next((ln for ln in content.splitlines() if ln.strip()), "")
        first = re.sub(r"^\s*(#+|[-*+]|\d+[.)]|>)\s*|\*\*|__|\|", "", first)
        key = re.sub(r"\s+", "", first)[:20]
        if key and key not in got:
            problems.append(f"넣은 범위에서 '{first.strip()[:20]}'을(를) 찾지 못함")
    return problems


# ==============================================================================
# 3. replace
# ==============================================================================

def _stories(doc: Any) -> List[Any]:
    """본문, 머리글/바닥글, 글상자, 각주 등 모든 텍스트 영역."""
    result = []
    try:
        for story in doc.StoryRanges:
            s = story
            while s is not None:
                result.append(s)
                try:
                    s = s.NextStoryRange
                except Exception:
                    s = None
    except Exception:
        result = [doc.Content]
    return result or [doc.Content]


def _load_map(args) -> Dict[str, str]:
    if args.map:
        raw = args.map.strip()
        if not raw.startswith("{"):
            raw = rt.read_text_file(rt.resolve_input_path(raw, "치환 목록 파일"))
        try:
            data = json.loads(raw)
        except json.JSONDecodeError as e:
            raise rt.InputError(f"--map은 JSON 객체여야 합니다 (예: {{\"초안\": \"최종\"}}): {e}")
        if not isinstance(data, dict) or not data:
            raise rt.InputError("--map은 비어 있지 않은 JSON 객체여야 합니다.")
        mapping = {str(k): ("" if v is None else str(v)) for k, v in data.items()}
    elif args.find is not None:
        mapping = {args.find: args.find if args.replace is None else args.replace}   # --replace "" 는 삭제
    else:
        raise rt.InputError("--find 또는 --map을 지정하세요.")
    if "" in mapping:
        # 빈 글을 찾으면 모든 글자 사이에 바꿀 글이 끼어들 수 있음
        raise rt.InputError("찾을 글(--find 또는 --map의 키)이 비어 있어 바꾸지 않았습니다. 빈 곳에 글을 넣으려면 word.py write를 쓰세요 "
                            "(예: word.py write --text \"내용\" --anchor \"기준 글\" --pos after). 바꿀 글이 있으면 찾을 글을 정확히 지정하세요.")
    long_keys = [k for k in mapping if _find_len(k) > FIND_MAX]
    if long_keys:
        raise rt.InputError(f"찾을 문구는 {FIND_MAX}자 이하여야 합니다: '{long_keys[0][:30]}...' 문장을 나눠서 지정하세요.")
    return mapping


def cmd_replace(args) -> Dict[str, Any]:
    mapping = _load_map(args)
    color = off.parse_color(args.color) if args.color else None
    highlight = None
    if args.highlight:
        highlight = WD_HIGHLIGHT.get(args.highlight.strip().lower())
        if highlight is None:
            raise rt.not_found("형광펜 색", args.highlight, [k for k in WD_HIGHLIGHT if k.isascii()], vocab=True)
    formatting = any(x is not None for x in (color, highlight, args.bold, args.italic, args.size))
    if not args.map and args.replace is None and not formatting:
        # 바꿀 글도 서식도 없으면 아무것도 바뀌지 않는데 'N곳 처리'로 보고되던 문제
        raise rt.InputError("--replace(바꿀 글)나 서식 옵션(--bold, --color, --highlight 등)이 없어 바꿀 것이 없습니다. "
                            "사용자에게 무엇으로 바꿀지 확인하세요. 글을 지우려면 --replace \"\"를 씁니다.")
    counts: Dict[str, int] = {}
    same: Dict[str, int] = {}          # 바꿀 글이 찾는 글과 같고 서식도 없어 바꾸지 않은 문구: 찾은 곳 수
    limit_hit = False
    with off.open_target("word", args.target, readonly=False, focus=not args.no_focus,
                         tool=TOOL, command="replace") as t:
        for find, rep in mapping.items():
            if find == rep and not formatting:
                # 바뀌는 것이 없음: 세기만 함 (이전에는 'N곳 처리'로 보고하고 바꾼 것으로 표시해 되돌리기 지점이 남았음)
                hits = _count_text(t.doc, find, args.match_case, args.whole_word)
                if hits:
                    same[find] = hits
        for story in _stories(t.doc):
            for find, rep in mapping.items():
                if find == rep and not formatting:
                    continue
                n = 0
                rng = story.Duplicate
                while _find(rng, find, args.match_case, args.whole_word):
                    if not formatting and str(rng.Text) == rep:
                        # 찾은 글이 이미 바꿀 글 그대로 (예: --find iphone --replace iPhone에서 'iPhone'): 세기만 함
                        # (이전에는 바꾼 곳으로 세어 되돌리기 지점이 남았음)
                        same[find] = same.get(find, 0) + 1
                        rng.Collapse(0)
                        continue
                    if str(rng.Text) != rep:
                        rng.Text = rep
                    if formatting:
                        if color is not None:
                            rng.Font.Color = color
                        if args.bold is not None:
                            rng.Font.Bold = args.bold
                        if args.italic is not None:
                            rng.Font.Italic = args.italic
                        if args.size:
                            rng.Font.Size = args.size
                        if highlight is not None:
                            rng.HighlightColorIndex = highlight
                    n += 1
                    if n >= 5000:
                        limit_hit = True
                        break
                    rng.Collapse(0)            # wdCollapseEnd: 바꾼 글 뒤부터 이어서 찾기 (바꾼 글을 다시 찾지 않음)
                if n:
                    counts[find] = counts.get(find, 0) + n
        # R1: 다시 찾아 확인. 바꾼 글이 찾는 글을 품지 않으면 찾는 글이 남지 않아야 하고, 아니면 바꾼 글이 그만큼 있어야 함
        problems = []
        if counts and not limit_hit:
            for find, n in counts.items():
                rep = mapping[find]
                # --match-case가 없으면 대소문자만 다른 바꿈(office -> Office)도 찾는 글이 '남은' 것으로 세지 않음
                inside = (find in rep) if args.match_case else (find.casefold() in rep.casefold())
                if not inside:
                    left = _count_text(t.doc, find, args.match_case, args.whole_word)
                    if left:
                        problems.append(f"'{find}'이(가) {left}곳 남아 있음")
                elif _count_text(t.doc, rep, True, args.whole_word) < n:
                    # 바꾼 곳은 바꿀 글이 대소문자까지 그대로 있어야 함 (대소문자만 바꾸는 경우도 확인)
                    problems.append(f"'{rep}'이(가) 바꾼 수({n})보다 적음")
        if counts:
            t.set_focus()
        resolved = t.resolved()
    total = sum(counts.values())
    result = {"target_resolved": resolved, "replaced": total, "by_text": counts}
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
    if limit_hit:
        result["limit_reached"] = True
        return rt.ok_unverified(TOOL, "replace", f"{total}곳 처리 (문구당 5000곳 상한 도달, 일부 남았을 수 있음)",
                                ["문구당 5000곳 상한에 도달해 남은 곳이 있는지 다시 찾아 확인하지 못함"], **result)
    preview = f"{total}곳 처리: " + ", ".join(f"'{k}' {v}곳" for k, v in counts.items())
    if problems:
        return rt.ok_unverified(TOOL, "replace", preview, problems, **result)
    return rt.ok(TOOL, "replace", preview, "VERIFIED: 문서를 다시 찾아 확인했습니다.", verified=True, **result)


def _count_text(doc: Any, text: str, match_case: bool, whole_word: bool, limit: int = 5000) -> int:
    """문서 전체(본문, 머리글/바닥글, 글상자 등)에서 글이 몇 곳 있는지 (다시 읽어 확인용)."""
    n = 0
    if _find_len(text) > FIND_MAX:
        # Word Find 한도(255자)를 넘는 글(긴 바꿀 글을 다시 셀 때): Word Find 대신 각 스토리 글에서 직접 셈
        needle = text if match_case else text.lower()
        for story in _stories(doc):
            hay = str(story.Text) if match_case else str(story.Text).lower()
            n += hay.count(needle)
        return min(n, limit)
    for story in _stories(doc):
        rng = story.Duplicate
        while n < limit and _find(rng, text, match_case, whole_word):
            n += 1
            rng.Collapse(0)
    return n


# ==============================================================================
# 4. export
# ==============================================================================

def _working_copy(app: Any, doc: Any) -> Tuple[Any, str]:
    """현재 메모리 상태(미저장 포함)를 창 없는 사본 문서로 엽니다. 원본은 그대로."""
    import tempfile
    fd, tmp = tempfile.mkstemp(prefix="office_word_", suffix=".xml")
    os.close(fd)
    with open(tmp, "w", encoding="utf-8") as f:
        f.write(doc.WordOpenXML)
    return app.Documents.Open(tmp, ConfirmConversions=False, AddToRecentFiles=False, Visible=False), tmp


def cmd_export(args) -> Dict[str, Any]:
    ext = os.path.splitext(args.save)[1].lower() if args.save else ".pdf"
    if ext not in (".pdf", ".docx"):
        raise rt.InputError(f"--save 확장자는 .pdf 또는 .docx여야 합니다: {args.save}")
    work = tmp = None
    with off.open_target("word", args.target, readonly=True) as t:
        if ext == ".pdf":
            off.require_pdf_export(t.app, ".docx")                # 기능이 없으면 사본을 만들기 전에 안내
        out = rt.resolve_output_path(args.save, os.path.splitext(t.name)[0], ext)
        try:
            target_doc = t.doc
            if args.page_number or ext == ".docx":
                work, tmp = _working_copy(t.app, t.doc)
                target_doc = work
            if args.page_number:
                for section in target_doc.Sections:
                    section.Footers(1).PageNumbers.Add(PageNumberAlignment=1)
            if ext == ".pdf":
                try:
                    target_doc.ExportAsFixedFormat(OutputFileName=os.path.normpath(out), ExportFormat=17,
                                                   OpenAfterExport=False, OptimizeFor=0, Range=0)
                except Exception as e:
                    if "25726" in str(e) or "기능이 설치되어 있지 않아" in str(e):
                        raise rt.InputError("Word PDF 내보내기 기능이 없습니다. --save에 .docx를 쓰세요.")
                    raise
            else:
                off.word_save_as(target_doc, os.path.normpath(out), 16)
        finally:
            if work is not None:
                try:
                    work.Close(SaveChanges=0)
                except Exception:
                    pass
            if tmp and os.path.exists(tmp):
                try:
                    os.remove(tmp)
                except OSError:
                    pass
        resolved = t.resolved()
    return rt.ok_file(TOOL, "export", f"{resolved['name']}을(를) {ext[1:].upper()}로 저장: {out}", out,
                      "VERIFIED: 저장한 파일을 확인했습니다. 원본 문서는 바뀌지 않았습니다.", target_resolved=resolved, output_file=out,
                 page_number=bool(args.page_number))


# ==============================================================================
# 5. 진입점
# ==============================================================================

def build_cli():
    parser, sub = rt.build_parser(TOOL, "워드 작업")

    def common(p, write=False):
        p.add_argument("--target", "-t", default="active", help="active(기본) | 열린 문서 이름 | 파일 경로")
        if write:
            p.add_argument("--no-focus", action="store_true", help="작업 후 화면 이동 안 함")

    p = sub.add_parser("read", help="내용 읽기")
    common(p)
    p.add_argument("--start", type=int, default=1, help="이 번호의 단락부터 읽기")
    p.add_argument("--pages", help="이 쪽만 읽기 (예: 3, 2-5, 1,4,7)")
    p.add_argument("--paragraphs", help="이 단락만 읽기 (예: 120-140, 141-, 7). 번호는 read/outline의 단락 번호")
    p.add_argument("--section", help="이 제목의 장절만 읽기 (제목 글 일부, 예: \"1.1\", \"검증 방법\"). 하위 절 포함")
    p.add_argument("--occurrence", type=int, help="--section 제목이 여러 개일 때 몇 번째인지 (오류의 [후보] 번호)")
    p.add_argument("--save", nargs="?", const=True, help="md 파일로 저장 (경로 생략 시 작업 폴더)")

    p = sub.add_parser("outline", help="제목 구조(목차): 수준, 단락 번호, 쪽")
    common(p)
    p.add_argument("--start", type=int, default=1, help="이 번호의 제목부터 보기")

    p = sub.add_parser("comment", help="검토 메모 목록 보기(list) / 메모 달기(add)")
    common(p, write=True)
    p.add_argument("--action", "-a", choices=["list", "add"], default="list", help="list(기본): 메모 목록, add: 메모 달기")
    p.add_argument("--anchor", help="add: 메모를 달 글 (문서에 있는 그대로)")
    p.add_argument("--text", help="add: 메모 내용. 줄바꿈은 \\n")
    p.add_argument("--file", "-f", help="add: 메모 내용 파일(.txt/.md) 또는 여러 메모 CSV (열: 기준 글, 메모 내용[, 몇 번째 단락])")
    p.add_argument("--occurrence", type=int, help="add: --anchor가 여러 단락에 있을 때 몇 번째인지 (오류의 [후보] 번호)")
    p.add_argument("--start", type=int, default=1, help="list: 이 번호의 메모부터 보기")

    p = sub.add_parser("revisions", help="변경 내용 추적(수정 기록) 목록")
    common(p)
    p.add_argument("--start", type=int, default=1, help="이 번호의 변경부터 보기")

    p = sub.add_parser("info", help="문서 요약")
    common(p)
    p.add_argument("--all", action="store_true", help="열린 문서 전체 목록")
    p.add_argument("--detail", action="store_true", help="그림 크기(쪽 너비 대비)와 제목 수준, 배치 경고")

    p = sub.add_parser("write", help="내용 넣기")
    common(p, write=True)
    p.add_argument("--text", help="넣을 글. 마크다운(# 제목, - 목록, | 표 |, **굵게**)은 Word 서식으로 변환. 줄바꿈은 \\n")
    p.add_argument("--file", "-f", help="넣을 파일 (.md/.txt는 글, .csv는 표)")
    p.add_argument("--image", help="넣을 그림 파일")
    p.add_argument("--width", type=float, help="그림 너비(pt). 생략 시 본문 폭에 맞춤")
    p.add_argument("--heading", type=int, choices=range(1, 7), help="--text를 이 수준의 제목으로")
    p.add_argument("--title", action="store_true", help="--text를 문서 제목 스타일로")
    p.add_argument("--callout", action="store_true", help="--text를 강조 요약 상자로")
    p.add_argument("--toc", action="store_true", help="목차 넣기")
    p.add_argument("--page-break", action="store_true", help="넣기 전에 페이지 나누기")
    p.add_argument("--anchor", help="이 글이 있는 단락을 기준으로 넣기")
    p.add_argument("--occurrence", type=int,
                   help="--anchor가 여러 단락에 있을 때 몇 번째 단락인지 (1부터, 오류의 [후보] 번호)")
    p.add_argument("--pos", choices=["after", "before", "replace", "end", "start"], default="after",
                   help="앵커 기준 위치 (앵커 없으면 end/start만 의미, 기본: 문서 끝)")
    p.add_argument("--theme", help=f"표/요약 상자 색: {', '.join(off.THEMES)}")

    p = sub.add_parser("replace", help="찾아 바꾸기")
    common(p, write=True)
    p.add_argument("--find", help="찾을 글")
    p.add_argument("--replace", help="바꿀 글 (빈 문자열이면 삭제, 생략하면 서식만 적용)")
    p.add_argument("--map", help='여러 개 한 번에: JSON 객체 또는 JSON 파일 (예: {"초안":"최종"})')
    p.add_argument("--match-case", action="store_true", help="대소문자 구분")
    p.add_argument("--whole-word", action="store_true", help="단어 전체 일치")
    p.add_argument("--bold", action="store_true", default=None, help="굵게")
    p.add_argument("--italic", action="store_true", default=None, help="기울임")
    p.add_argument("--color", help="글자색")
    p.add_argument("--size", type=float, help="글자 크기")
    p.add_argument("--highlight", help="형광펜 (yellow, green, pink, blue, red, gray, none)")

    p = sub.add_parser("export", help="PDF/DOCX 사본 저장")
    common(p)
    p.add_argument("--save", help="저장 경로 (.pdf 또는 .docx, 생략 시 작업 폴더에 PDF)")
    p.add_argument("--page-number", action="store_true", help="사본 아래쪽 가운데에 쪽 번호")
    return parser


HANDLERS = {"read": cmd_read, "info": cmd_info, "write": cmd_write, "replace": cmd_replace, "export": cmd_export,
            "outline": cmd_outline, "comment": cmd_comment, "revisions": cmd_revisions}


def main(argv: Optional[Sequence[str]] = None) -> None:
    args = build_cli().parse_args(argv)
    if args.command == "write" and not args.anchor and args.pos not in ("end", "start"):
        args.pos = "end"
    rt.run_cli(TOOL, args.command, lambda: HANDLERS[args.command](args))


if __name__ == "__main__":
    main()
