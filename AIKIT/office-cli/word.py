"""word.py - 워드 작업

명령:
    read     내용 읽기 (제목 구조 포함, --save로 md 저장). --pages / --paragraphs / --section으로 범위 고르기
    outline  제목 구조(목차): 수준, 단락 번호, 쪽
    comment  검토 메모 목록 (--action add: 기준 글에 메모 달기, --file CSV로 여러 개, reply 답글, delete 지우기)
    revisions 변경 내용 추적(수정 기록) 목록
    info     문서 요약 (페이지/단어/글자 수, 표, 그림, 제목 목록). --all: 열린 문서 전체 목록
    write    텍스트/마크다운/표/그림/요약상자/목차 넣기 (기본: 문서 끝, --anchor로 위치 지정, --dry-run)
    delete   단락 지우기 (--paragraphs, read의 단락 번호)
    format   단락 서식 (정렬, 글꼴, 크기, 굵게, 기울임, 스타일)
    table    표 행 넣기·지우기
    header   머리글·바닥글 읽기·쓰기
    link     하이퍼링크 목록·달기 (#책갈피 포함)
    bookmark 책갈피 목록·만들기
    replace  찾아 바꾸기 (+서식). 머리글/바닥글/글상자 포함 (--dry-run)
    fill     서식 문서 채우기: 값 파일(자리, 값)의 {{이름}} 자리와 표 칸(표2!R3C4)을 한 번에 (--dry-run, --overwrite)
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
import _markdown as md  # noqa: E402

TOOL = "word"

# WdBuiltinStyle (언어와 무관한 기본 스타일 번호)
WD_STYLE_NORMAL = -1
WD_STYLE_TITLE = -63
WD_STYLE_LIST_BULLET = -49
WD_STYLE_LIST_NUMBER = -50
WD_BORDER_BOTTOM = -3
WD_HIGHLIGHT = {"yellow": 7, "green": 4, "turquoise": 3, "pink": 5, "blue": 2, "red": 6, "gray": 16, "grey": 16,
                "노랑": 7, "초록": 4, "분홍": 5, "파랑": 2, "빨강": 6, "회색": 16, "none": 0}
FIND_MAX = 255   # Word Find 검색어 길이 한도 (^를 ^^로 바꾼 뒤의 길이)


def _find_len(text: str) -> int:
    """Word Find에 넘기는 길이: _find가 ^를 ^^로 바꾸므로 그 길이로 셉니다."""
    return len(str(text).replace("^", "^^"))


def _u16(text: str) -> int:
    """Word Range 위치로 센 글 길이 (UTF-16 단위: BMP 밖 글자는 2)."""
    return len(text.encode("utf-16-le", "surrogatepass")) // 2


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
    """본문 그림(InlineShape)은 Range.Text에서 한 글자('\\x01')로 나옵니다.
    그림의 실제 위치에 있는 그 글자만 [그림]으로 바꿉니다."""
    if "\x01" not in raw:
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
        if 0 <= pos < len(chars) and chars[pos] == "\x01":
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
        raise rt.ambiguous(f"'{query}'이(가) 들어간 제목이 {len(hits)}개라 고를 수 없습니다. 사용자에게 번호를 고르게 하세요 "
                           "(고른 번호의 confirm_commands가 --occurrence를 붙인 명령).",
                           [f"{i}: 단락 {h['no']} {_preview(h['text'], 40)}" for i, h in enumerate(hits, 1)],
                           option="--occurrence", values=[str(i) for i in range(1, len(hits) + 1)])
    else:
        head = hits[0]
    i = units.index(head)
    j = i + 1
    while j < len(units) and not (units[j].get("level") and units[j]["level"] <= head["level"]):
        j += 1
    return units[i:j]


def _from_target(args: Any, source: Dict[str, Any]) -> List[str]:
    """transfer 다음 명령의 --from-target: 보고 있는 문서(active)가 아니면 읽은 문서 (전체 경로가 있으면 그것)."""
    if (args.target or "active").strip().lower() == "active":
        return []
    return ["--from-target", source.get("path") or args.target]


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
        out = rt.resolve_output_path(None if args.save is True else args.save, os.path.splitext(name)[0], ".md",
                                     near=source.get("path"))
        out = rt.write_text(out, md.document_markdown(
            [("heading", (u["level"], u["text"])) if u.get("level") else ("paragraph", u["text"]) for u in save_units]))
        # 열린 문서를 바꾸는 이어 하기는 confirm_commands (1.4.75)
        result.update(output_file=out, confirm_commands=[
            rt.next_command("ppt.py", ["write", "--file", out], "저장한 내용을 열린 PPT에 넣기 (사용자가 원할 때만)"),
            rt.next_command("office.py", ["transfer", "--from-app", "word", "--what", "text", "--to-app", "excel"]
                            + _from_target(args, source),
                            "이 문서의 단락을 엑셀 표로 옮기기 (사용자가 원할 때만)")])
        return rt.ok_file(TOOL, "read", f"{name}: 단락 {len(save_units)}개를 저장했습니다 ({out})", out,
                          "VERIFIED: 저장한 파일을 다시 열어 확인했습니다. " + rt.INTO_OTHER_DOC,
                          **result)
    if not selected:
        # 빈 결과 명시 (ACI 설계 권장 2번): 필터 없이 읽은 빈 문서도 count 0 (1.4.81)
        where = selection or ("문서" if not units else f"단락 {max(1, args.start)}부터")
        result.update(paragraphs=[], truncated=False, count=0)
        return rt.ok(TOOL, "read", f"{name}: {where}에 글이 있는 단락 없음 (count 0)",
                     f"VERIFIED: 0건. {where}에는 글이 있는 단락이 없습니다 (찾은 항목 없음: 표·그림만 있거나 빈 문서·쪽). "
                     "사용자에게 그대로 알리세요.", **result)
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
        if not docs:
            return off.no_open_docs_result(TOOL, "info", "0건: 열린 문서가 없습니다",
                         "VERIFIED: 0건. 열린 Word 문서가 없습니다. 사용자에게 그대로 알리세요.", documents=docs, count=0)
        return rt.ok(TOOL, "info", f"열린 문서 {len(docs)}개", documents=docs, count=len(docs))
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
            raise rt.InputError(f"메모 파일 {n}번째 줄의 3번째 열(몇 번째 단락)은 숫자여야 합니다: '{occ_raw}'",
                                log_message=f"메모 파일 {n}번째 줄의 3번째 열(몇 번째 단락)은 숫자여야 합니다 (값 {rt.DOC_OMITTED}).")
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
                problems.append((n, f"{n}번째 줄 '{_preview(anchor, 20)}': {msg}"))
        if problems:
            head = (f"메모 {len(rows)}개 중 {len(problems)}개의 기준 글을 정할 수 없어 하나도 달지 않았습니다. "
                    "메모 파일의 기준 글을 고치거나 3번째 열에 몇 번째 단락인지 적은 뒤 다시 실행하세요. ")
            # 오류 로그에는 줄 번호만 (기준 글은 메모 파일 내용, 후보는 문서 단락)
            raise rt.InputError(head + " / ".join(p for _, p in problems[:10]),
                                log_message=head + f"기준 글을 정할 수 없는 줄: {', '.join(str(n) for n, _ in problems[:10])} "
                                                   f"(기준 글과 후보 {rt.DOC_OMITTED})")
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


def _comment_reply_delete(args) -> Dict[str, Any]:
    """comment --action reply --no N --text (답글, Word 2013+) / --action delete --no N (메모와 그 답글을 지움, --dry-run)."""
    if not args.no:
        raise rt.InputError(f"메모 번호를 --no로 주세요 (word.py comment 목록의 no). 예: --action {args.action} --no 2")
    if args.action == "reply" and not (args.text or "").strip():
        raise rt.InputError("답글 내용을 --text로 주세요.")
    with off.open_target("word", args.target, readonly=bool(args.dry_run), focus=not args.no_focus,
                         tool=TOOL, command="comment") as t:
        doc = t.doc
        total = int(doc.Comments.Count)
        if not 1 <= args.no <= total:
            raise rt.InputError(f"메모 {args.no}이(가) 없습니다 (1~{total}). 번호는 word.py comment 목록의 no입니다." if total
                                else "문서에 검토 메모가 없습니다.")
        c = doc.Comments(args.no)
        head = _read_comments(doc, [args.no])[0]
        try:
            root = c.Ancestor if c.Ancestor is not None else c
        except Exception:
            root = c
        replies = [i for i in range(1, total + 1) if i != args.no and _is_reply_of(doc.Comments(i), c)]
        info: Dict[str, Any] = {"no": args.no, "author": head["author"], "text": _preview(head["text"], 80),
                                "target_text": head.get("target_text", "")}
        if args.action == "delete":
            desc = f"메모 {args.no}({head['author']}: {_preview(head['text'], 40)})" + (f"와 답글 {len(replies)}개" if replies else "")
            if args.dry_run:
                return rt.dry_run_result(TOOL, "comment", f"미리보기: {desc}을(를) 지웁니다 (바꾸지 않음)", 1 + len(replies), [desc],
                                         rt.dry_run_command(TOOL, f"{desc} 지우기"), target_resolved=t.resolved(), **info)
            # 실제 Word의 Comment.Delete는 그 메모만 지우고 답글은 남김 (회사 PC 2026-11-11 N10): 답글을 뒤 번호부터 먼저 지움
            for i in sorted(replies, reverse=True):
                doc.Comments(i).Delete()
            c.Delete()
            left = int(doc.Comments.Count)
            t.set_focus()
            resolved = t.resolved()
            preview = f"{desc}을(를) 지웠습니다"
            if left != total - 1 - len(replies):
                return rt.ok_unverified(TOOL, "comment", preview, [f"메모 수가 {left} (지우기 전 {total})"], target_resolved=resolved, **info)
            return rt.ok(TOOL, "comment", preview, "VERIFIED: 메모 수를 다시 세어 확인했습니다. 같은 명령을 다시 실행하지 마세요 (번호가 바뀜). "
                         "잘못 지웠으면 office.py undo로 되돌립니다.", verified=True, target_resolved=resolved, **info)
        if args.dry_run:
            raise rt.dry_run_unsupported(TOOL, "comment --action reply")
        text = rt.unescape_newlines(args.text)
        try:
            root.Replies.Add(Range=root.Scope, Text=text)
        except Exception as e:
            raise rt.InputError(f"이 Word에서는 메모 답글을 달 수 없습니다 ({type(e).__name__}, Word 2013 이상 필요). "
                                "새 메모로 달려면 comment --action add를 쓰세요.", no_fallback=True)
        after = int(doc.Comments.Count)
        found = any(_clean(doc.Comments(i).Range.Text).strip() == text.strip() and _is_reply_of(doc.Comments(i), root)
                    for i in range(1, after + 1))
        t.set_focus()
        resolved = t.resolved()
    preview = f"메모 {args.no}({info['author']})에 답글을 달았습니다: {_preview(text, 40)}"
    if not found:
        return rt.ok_unverified(TOOL, "comment", preview, ["단 답글을 다시 찾지 못함"], target_resolved=resolved, **info)
    return rt.ok(TOOL, "comment", preview, "VERIFIED: 답글을 다시 읽어 확인했습니다. 같은 명령을 다시 실행하지 마세요 (답글이 하나 더 생김).",
                 verified=True, target_resolved=resolved, **info)


def _is_reply_of(c: Any, parent: Any) -> bool:
    try:
        a = c.Ancestor
        return a is not None and int(a.Index) == int(parent.Index)
    except Exception:
        return False


def cmd_comment(args) -> Dict[str, Any]:
    if args.action in ("reply", "delete"):
        return _comment_reply_delete(args)
    # --action add를 빠뜨리고 메모 내용만 준 경우: 목록을 보여 주고 "완료"라고 하면 LLM이 메모를 단 것으로 오해합니다
    if args.action == "add" or args.anchor or args.text is not None or args.file:
        if args.dry_run:
            raise rt.dry_run_unsupported(TOOL, "comment --action add")
        with off.open_target("word", args.target, readonly=False, focus=not args.no_focus,
                             tool=TOOL, command="comment") as t:
            return _comment_add(args, t)
    with off.open_target("word", args.target, readonly=True) as t:
        total = int(t.doc.Comments.Count)
        items = _read_comments(t.doc, _window(total, args.start))
        resolved = t.resolved()
    return rt.dry_run_read_only(_list_result(resolved["name"], "comment", "검토 메모", items, total, args.start, args.target,
                                             target_resolved=resolved), args.dry_run, "comment list")


WD_REVISION_DELETE = 2
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

WD_LIST_APPLY_SELECTION, WD_LIST_APPLY_FORWARD = 1, 2     # ListFormat.ApplyListTemplate ApplyTo


class Writer:
    """문서의 한 위치(pos)부터 블록을 차례로 넣는 도우미. 넣은 뒤 pos는 다음 삽입 위치로 이동합니다."""

    def __init__(self, doc: Any, pos: int, theme: Optional[Dict[str, Any]]):
        self.doc = doc
        self.pos = pos
        self.theme = theme or off.THEMES["navy"] | {"name": "navy"}
        self.created: List[str] = []
        self.warnings: List[str] = []
        self.problems: List[str] = []

    def _para(self, text: str, style: Optional[int] = None, bold_spans: Sequence[Tuple[int, int]] = ()) -> Any:
        r = self.doc.Range(self.pos, self.pos)
        r.InsertAfter(text + "\r")
        # Word 위치는 UTF-16 단위: BMP 밖 글자(이모지 등)는 2칸이라 파이썬 글자 위치를 그대로 쓰면 굵게가 밀림
        para = self.doc.Range(r.Start, r.Start + _u16(text))
        try:
            para.Style = self.doc.Styles(style if style is not None else WD_STYLE_NORMAL)
        except Exception:
            pass
        para.Font.Bold = False
        for a, b in bold_spans:
            self.doc.Range(r.Start + _u16(text[:a]), r.Start + _u16(text[:b])).Font.Bold = True
        self.pos = r.End
        return para

    def text(self, text: str, style: Optional[int] = None) -> None:
        plain, spans = md.split_bold(text)
        self._para(plain, style, spans)

    def heading(self, level: int, text: str) -> None:
        self._para(md.plain_heading(text), heading_style(level))   # 제목은 모든 수준이 ** __를 뺀 글 (굵기가 섞이지 않음)
        self.created.append(f"제목{level}")

    def items(self, items: List[str], numbered: bool) -> None:
        first = self.pos
        for it in items:
            self.text(it, WD_STYLE_LIST_NUMBER if numbered else WD_STYLE_LIST_BULLET)
        if numbered and items:
            # 목록 번호 스타일만 쓰면 문서의 앞 번호 목록에 이어 4, 5로 매겨짐: 새 목록은 1부터 (2026-11-09 검토 C2, 실기 확인)
            # ApplyTo 기본값(목록 전체)은 앞 목록까지 한 목록으로 보아 다시 1부터가 되지 않음 (회사 PC 2026-11-10 L5):
            # 넣은 단락만(1), 안 되면 여기부터(2). 첫 번호를 다시 읽어 확인
            why = ""
            for apply_to in (WD_LIST_APPLY_SELECTION, WD_LIST_APPLY_FORWARD):
                try:
                    lf = self.doc.Range(first, self.pos).ListFormat
                    lf.ApplyListTemplate(ListTemplate=lf.ListTemplate, ContinuePreviousList=False, ApplyTo=apply_to)
                    got = str(self.doc.Range(first, first).ListFormat.ListString or "").strip()
                except Exception as e:                # noqa: BLE001
                    why = type(e).__name__
                    continue
                if got.rstrip(".)") == "1":
                    why = ""
                    break
                why = f"첫 번호 '{got}'"
            if why:
                self.warnings.append(f"번호 목록을 1부터 다시 매기지 못함 ({why}): 앞 목록에 이어 번호가 매겨졌을 수 있음. "
                                     "Word에서 첫 번호를 오른쪽 클릭 > '1에서 다시 시작'을 누르라고 알리세요")
        self.created.append(f"{'번호' if numbered else '글머리'} 목록 {len(items)}개")

    def _after_table(self) -> None:
        """넣을 자리 바로 앞이 표면 빈 단락을 먼저 (표·요약 상자가 앞 표와 합쳐지고 머리글 서식이 사용자 표의 첫 행을 바꾸던 것, 검토 C3)."""
        if self.pos > 0 and _in_table(self.doc.Range(self.pos - 1, self.pos - 1)):
            self._para("")

    def _normal_paragraph(self, start: int) -> None:
        """새로 만든 단락(그림·목차 자리)의 스타일을 본문으로 (앞 단락의 제목·글머리 스타일을 물려받던 것, 검토 C4)."""
        try:
            self.doc.Range(start, start).Paragraphs(1).Style = self.doc.Styles(WD_STYLE_NORMAL)
        except Exception:
            pass

    def table(self, rows: List[List[Any]], markdown: bool = False) -> Any:
        self._after_table()
        bold_cells: List[Tuple[int, int]] = []
        if markdown:
            rows, bold_cells = md.table_cells(rows)
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
        for i, j in bold_cells:                       # 칸 전체가 **굵게**인 칸 (합계 행 등)
            try:
                tbl.Cell(i + 1, j + 1).Range.Font.Bold = True
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
        self._after_table()
        r = self.doc.Range(self.pos, self.pos)
        r.InsertAfter(text.replace("\n", "\v") + "\r")
        body = self.doc.Range(r.Start, r.End - 1)
        try:
            body.Style = self.doc.Styles(WD_STYLE_NORMAL)
        except Exception:
            pass
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
        self._normal_paragraph(r.Start)
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
        self._normal_paragraph(r.Start)
        anchor = self.doc.Range(r.Start, r.Start)
        self._toc = self.doc.TablesOfContents.Add(Range=anchor, UseHeadingStyles=True, UpperHeadingLevel=1, LowerHeadingLevel=3)
        self.pos = self._toc.Range.End + 1
        self.created.append("목차")

    def toc_update(self) -> None:
        """넣을 내용을 모두 넣은 뒤 목차 갱신 (--toc --file: 본문 전에 갱신해 빈 목차인데 VERIFIED였음, 검토 C5).
        갱신으로 목차 길이가 바뀌면 그만큼 끝 위치를 옮김."""
        toc = getattr(self, "_toc", None)
        if toc is None:
            return
        try:
            before = int(toc.Range.End)
            toc.Update()
            self.pos += int(toc.Range.End) - before
        except Exception as e:                        # noqa: BLE001
            self.problems.append(f"목차를 갱신하지 못함 ({type(e).__name__})")

    def rule(self) -> None:
        """마크다운 '---'(구분선): 아래 테두리만 있는 빈 단락 (페이지 나누기가 아님. 페이지 나누기는 --page-break)."""
        para = self._para("")
        try:
            border = para.Paragraphs(1).Borders(WD_BORDER_BOTTOM)
            border.LineStyle = 1                      # wdLineStyleSingle
            border.LineWidth = 6                      # wdLineWidth075pt
        except Exception as e:
            self.problems.append(f"구분선의 아래 테두리를 넣지 못함: {type(e).__name__}")
        self.created.append("구분선")

    def page_break(self) -> None:
        r = self.doc.Range(self.pos, self.pos)
        r.InsertBreak(7)                              # wdPageBreak
        self.pos = r.End
        self.created.append("페이지 나누기")

    def markdown(self, text: str) -> None:
        for kind, payload in md.parse_markdown_blocks(text):
            if kind == "heading":
                self.heading(*payload)
            elif kind in ("paragraph", "code"):
                if kind == "code":
                    self._para(payload, None, [])          # 코드 줄은 글자 그대로 (** __를 굵게로 읽지 않음)
                else:
                    self.text(payload)
                if not self.created or not self.created[-1].startswith("본문"):
                    self.created.append("본문")
            elif kind in ("bullets", "numbered"):
                self.items(payload, kind == "numbered")
            elif kind == "table":
                self.table(payload, markdown=True)
            elif kind == "break":
                self.rule()
            elif kind == "front_matter":
                self.warnings.append(md.front_matter_warning(payload))
                rt.log("WARN", self.warnings[-1])


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


# write --anchor에 문서의 글 대신 위치 낱말을 준 경우 (집 PC Gemma O4 '--anchor "end"'): 문서에 그 글이 없으면 같은 뜻의 위치로 [추천]
ANCHOR_END_WORDS = {"end", "끝", "문서 끝", "문서끝", "맨 끝", "맨끝", "마지막"}
ANCHOR_START_WORDS = {"start", "처음", "문서 처음", "맨 앞", "맨앞", "시작"}


def _anchor_paragraph(doc: Any, anchor: str, occurrence: Optional[int], pos: str = "after",
                      position_words: bool = False) -> Any:
    """앵커 단락 하나. 못 찾으면 비슷한 글을 안내 (확실할 때만 [추천]), 여러 단락에 있으면 고르지 않고 [후보]로 오류.
    position_words(write만): 앵커가 'end'·'끝' 같은 위치 낱말이고 문서에 없으면 --anchor를 뺀 명령(문서 끝, 처음은 --pos start)을 [추천]."""
    hits = _anchor_paragraphs(doc, anchor)
    if not hits:
        word = anchor.strip().lower()
        if position_words and pos == "after" and (word in ANCHOR_END_WORDS or word in ANCHOR_START_WORDS):
            fix = rt.retry_command({"--anchor": None, "--pos": "start" if word in ANCHOR_START_WORDS else None})
            raise rt.InputError(f"문서에서 '{anchor}'을(를) 찾을 수 없습니다. --anchor는 문서의 글을 받습니다. "
                                f"문서 {'처음' if word in ANCHOR_START_WORDS else '끝'}에 넣는 명령: [추천] {fix}", no_fallback=True)
        similar = _similar_anchor(doc, anchor)
        hint = ""
        if similar and _safe_anchor_fix(anchor, similar, pos):
            hint = f" [추천] {rt.retry_command({'--anchor': similar})}"
        elif similar:
            hint = f" 비슷한 문구: '{similar}' (뜻이 같은지 사용자에게 확인한 뒤에만 쓰세요)"
        raise rt.InputError(f"문서에서 '{anchor}'을(를) 찾을 수 없습니다. word.py read로 정확한 문구를 확인하세요." + hint,
                            no_fallback=True,
                            log_message=f"문서에서 '{anchor}'을(를) 찾을 수 없습니다." + (f" (비슷한 문구 {rt.DOC_OMITTED})" if hint else ""))
    if occurrence is not None:
        if not 1 <= occurrence <= len(hits):
            raise rt.InputError(f"--occurrence는 1~{len(hits)} 사이여야 합니다 ('{anchor}'이(가) 있는 단락 {len(hits)}개).",
                                no_fallback=True)
        return hits[occurrence - 1]
    if len(hits) > 1:
        more = "개 이상" if len(hits) >= ANCHOR_MAX_HITS else "개"
        raise rt.ambiguous(
            f"'{anchor}'이(가) 단락 {len(hits)}{more}에 있어 어느 곳인지 고를 수 없습니다. 사용자에게 번호를 고르게 하세요 "
            "(고른 번호의 confirm_commands가 --occurrence를 붙인 명령). 한 곳에만 있는 더 긴 문구를 --anchor로 써도 됩니다.",
            [f"{i}: {_preview(str(h.Text))}" for i, h in enumerate(hits, 1)],
            option="--occurrence", values=[str(i) for i in range(1, len(hits) + 1)])
    return hits[0]


_PENDING_REPLACE: Dict[str, Any] = {}   # --pos replace: 내용을 넣은 뒤 지울 앵커 단락 (_finish_replace)


def _finish_replace(doc: Any, inserted_len: int) -> Optional[str]:
    """--pos replace의 앵커 단락을 내용을 넣은 뒤에 지움. 넣은 글 만큼 뒤로 밀린 위치. 실패하면 이유."""
    if not _PENDING_REPLACE:
        return None
    a, b = _PENDING_REPLACE["replace_range"]
    last = _PENDING_REPLACE.get("last")
    _PENDING_REPLACE.clear()
    try:
        a, b = a + inserted_len, b + inserted_len
        if not last:
            doc.Range(a, b).Delete()
        else:
            doc.Range(a, b - 1).Delete()            # 마지막 단락은 단락 기호를 남김
        return None
    except Exception as e:                          # noqa: BLE001
        return f"넣은 뒤 앵커 단락을 지우지 못함 ({type(e).__name__}): 앵커 단락이 남아 있음"


def _insert_position(doc: Any, anchor: Optional[str], pos: str, occurrence: Optional[int] = None,
                     preview: Optional[Dict[str, Any]] = None) -> Tuple[int, str]:
    """삽입 위치(문자 오프셋)와 설명. 앵커가 있으면 그 단락의 앞/뒤, replace면 앵커 단락을 지우고 그 자리.
    앵커가 여러 단락에 있으면 occurrence(1부터)로 고른 단락, 없으면 오류 (첫 번째를 임의로 고르지 않음).
    preview(dict)를 주면 --dry-run: 문서를 바꾸지 않고(새 단락 만들기·앵커 단락 지우기 없음) 위치와 설명만,
    replace면 지워질 단락 글을 preview['removed']에 담습니다."""
    if not anchor:
        end = doc.Content.End - 1
        last = doc.Range(end, end).Paragraphs(1).Range
        if pos == "start":
            if _in_table(doc.Range(0, 0)):
                raise rt.InputError("문서가 표로 시작해 맨 앞에 넣으면 첫 칸 안에 들어갑니다. 표 뒤 단락의 글을 --anchor로 지정해 "
                                    "--pos before로 넣거나, Word에서 표 앞에 빈 줄을 만든 뒤 다시 실행하세요.", no_fallback=True)
            return 0, "문서 처음"
        if str(last.Text).strip() and preview is None:     # 마지막 단락에 내용이 있으면 새 단락에서 시작
            doc.Range(end, end).InsertAfter("\r")
            end = doc.Content.End - 1
        return end, "문서 끝"
    if _find_len(anchor) > FIND_MAX:
        raise rt.InputError(f"--anchor는 {FIND_MAX}자 이하여야 합니다(^는 2자로 셈). 앵커 단락의 앞부분 일부만 쓰세요.", no_fallback=True)
    para = _anchor_paragraph(doc, anchor, occurrence, pos, position_words=True)
    if _in_table(para):
        # 셀 단락의 End는 다음 셀의 시작이라 그대로 넣으면 옆 칸에 들어감: 표 바깥(표 뒤)에 넣습니다
        if pos != "after":
            raise rt.InputError(f"'{anchor}'은(는) 표 안에 있어 --pos {pos}로 넣으면 표 칸 안에 들어갑니다. --pos after(표 뒤에 넣기)를 쓰거나, "
                                "칸 안의 글만 바꾸려면 word.py replace를 쓰세요.", no_fallback=True)
        return int(para.Tables(1).Range.End), f"'{anchor}'이(가) 있는 표 뒤"
    if pos == "before":
        return int(para.Start), f"'{anchor}' 단락 앞"
    # 표 바로 위 단락: 그 End는 표 첫 칸의 시작이라 그대로 넣으면 표 첫 칸 안에 들어감 (2026-11-09 검토 C1, 실기 확인)
    before_table = int(para.End) < int(doc.Content.End) and _in_table(doc.Range(int(para.End), int(para.End)))
    if pos == "replace":
        start = int(para.Start)
        if preview is not None:
            preview["removed"] = str(para.Text)
            preview["replace_range"] = (start, int(para.End))
            return start, f"'{anchor}' 단락 자리"
        # 지우기는 내용을 넣은 뒤에 (먼저 지우면 넣다가 COM이 실패할 때 단락만 사라짐, 검토 C7): 앵커 단락 앞에 넣고 끝나면 지움
        preview_sink = {"replace_range": (start, int(para.End)), "last": int(para.End) >= int(doc.Content.End)}
        _PENDING_REPLACE.clear()
        _PENDING_REPLACE.update(preview_sink)
        return start, f"'{anchor}' 단락 자리"
    if before_table and preview is None:
        # 앵커 단락과 표 사이에 빈 단락을 만들고 거기에. 넣기 전 끝 위치를 씀: 실제 Word는 안에 글이 들어간 para Range를
        # 늘려 para.End가 다시 표 첫 칸이 됨 (회사 PC 2026-11-10 L4)
        end = int(para.End)
        doc.Range(end - 1, end - 1).InsertAfter("\r")
        return end, f"'{anchor}' 단락 뒤 (표 앞)"
    if int(para.End) >= int(doc.Content.End) and preview is None:
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


def _refresh_tocs(doc: Any, start: int, w: "Writer") -> int:
    """문서에 이미 있는 목차를 갱신. 넣은 범위 앞의 목차 길이가 바뀌면 넣은 범위(start, w.pos)를 그만큼 옮김. 새 start를 돌려줌."""
    try:
        tocs = doc.TablesOfContents
        n = int(tocs.Count)
    except Exception:
        return start
    for i in range(1, n + 1):
        try:
            toc = tocs(i)
            before_start, before_end = int(toc.Range.Start), int(toc.Range.End)
            toc.Update()
            delta = int(toc.Range.End) - before_end
            if before_start < start:
                start += delta
                w.pos += delta
        except Exception as e:                        # noqa: BLE001
            w.warnings.append(f"문서의 목차 {i}을(를) 갱신하지 못함 ({type(e).__name__}): Word에서 목차를 눌러 [필드 업데이트]")
    return start


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

    with off.open_target("word", args.target, readonly=args.dry_run, focus=not args.no_focus, create=True,
                         tool=TOOL, command="write") as t:
        doc = t.doc
        if args.dry_run:
            # 바꾸지 않음: 읽기 세션이라 되돌리기 지점도 남지 않음
            seen: Dict[str, Any] = {}
            _, where = _insert_position(doc, args.anchor, args.pos, args.occurrence, preview=seen)
            return _write_dry_run(args, t.resolved(), where, seen.get("removed"), text)
        _PENDING_REPLACE.clear()
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
                w._para(md.plain_heading(content.strip()), WD_STYLE_TITLE)
                w.created.append("문서 제목")
            elif md.looks_like_markdown(content):
                w.markdown(content)
            else:
                for para in [p for p in re.split(r"\n\s*\n|\n", content) if p.strip()]:
                    w.text(para.strip())
                w.created.append("본문")
        if file_rows:
            w.table(file_rows)
        if image:
            w.image(image, args.width)
        if args.toc:
            w.toc_update()                              # 목차는 내용을 넣은 뒤에 갱신 (먼저 갱신하면 빈 목차, 검토 C5)
        elif any(c and re.search(r"(?m)^#{1,3}\s", c) for c in (text, file_text)) or args.heading:
            start = _refresh_tocs(doc, start, w)        # 문서에 있던 목차도 새 제목을 넣으면 갱신 (검토 C5)
        gone = _finish_replace(doc, w.pos - start)
        if gone:
            w.problems.append(gone)
        mode = "raw" if args.callout else ("bold" if args.heading or args.title else "auto")
        problems = w.problems + _verify_inserted(doc, start, w.pos, [c for c in (text, file_text) if c], mode)
        t.set_focus()
        if t.source == "active" and not args.no_focus:
            try:
                doc.Range(w.pos, w.pos).Select()          # 커서를 넣은 내용 끝으로
            except Exception:
                pass
        resolved = t.resolved()
    preview = f"{resolved['name']}의 {where}에 넣었습니다: {', '.join(w.created)}"
    extra: Dict[str, Any] = {"warnings": w.warnings} if w.warnings else {}
    if problems:
        return rt.ok_unverified(TOOL, "write", preview, problems, target_resolved=resolved, inserted=w.created, **extra)
    return rt.ok(TOOL, "write", preview, "VERIFIED: 넣은 범위를 다시 읽어 확인했습니다. 같은 명령을 다시 실행하지 마세요."
                 + (" 주의: warnings를 사용자에게 알리세요." if w.warnings else ""),
                 verified=True, target_resolved=resolved, inserted=w.created, **extra)


def _write_dry_run(args, resolved: Dict[str, Any], where: str, removed: Optional[str], text: Optional[str]) -> Dict[str, Any]:
    """write --dry-run: 넣을 위치와 넣을 것(그리고 --pos replace면 지워질 단락)만 알려 줌."""
    parts: List[str] = []
    if args.page_break:
        parts.append("쪽 나누기")
    if args.toc:
        parts.append("목차")
    if text is not None:
        kind = "요약 상자" if args.callout else (f"제목{args.heading}" if args.heading else ("문서 제목" if args.title else "글"))
        parts.append(f"{kind} '{_preview(text, 30)}'")
    if args.file:
        parts.append(f"파일 {os.path.basename(args.file.strip().strip(chr(34)))}")
    if args.image:
        parts.append(f"그림 {os.path.basename(args.image.strip().strip(chr(34)))}")
    planned = ([f"지우기: 단락 '{_preview(removed, 40)}'"] if removed is not None else []) + [f"넣기: {p}" for p in parts]
    preview = f"미리보기: {resolved['name']}의 {where}에 넣습니다: {', '.join(parts)}"
    if removed is not None:
        preview += f". 지워질 단락: '{_preview(removed, 40)}'"
    replace = removed is not None
    return rt.dry_run_result(TOOL, "write", preview + " (바꾸지 않음)", 1, planned,
                             rt.dry_run_command(TOOL, "앵커 단락을 지우고 그 자리에 넣기 실행" if replace else "내용 넣기 실행",
                                                destructive=replace), target_resolved=resolved)


def _verify_inserted(doc: Any, start: int, end: int, contents: List[str], mode: str = "auto") -> List[str]:
    """R1: 넣은 범위를 다시 읽어, 범위가 비어 있지 않고 넣은 글의 첫 글이 들어 있는지 확인합니다.
    mode: 'auto'(마크다운이면 블록으로 해석, 아니면 줄마다 단락), 'bold'(--heading/--title: 굵게 표시만 뺌), 'raw'(--callout: 그대로)."""
    try:
        got = re.sub(r"[\s\x00-\x1f]+", "", str(doc.Range(start, end).Text))   # 표 칸 구분 기호도 무시
    except Exception as e:
        return [f"넣은 범위를 다시 읽지 못함: {type(e).__name__}"]
    if end <= start:
        return ["넣은 범위가 비어 있음"]
    problems = []
    for content in contents:
        first = _first_inserted(content, mode)
        key = re.sub(r"\s+", "", first)[:20]
        if key and key not in got:
            problems.append(f"넣은 범위에서 '{first.strip()[:20]}'을(를) 찾지 못함")
    return problems


def _first_inserted(content: str, mode: str) -> str:
    """넣은 글의 첫 글 (Writer가 넣는 모양 그대로: 마크다운 블록은 해석한 글, 코드 줄과 표 칸은 글자 그대로)."""
    if mode == "raw":
        return content.strip()
    if mode == "auto" and md.looks_like_markdown(content):
        for kind, payload in md.parse_markdown_blocks(content):
            if kind == "heading":
                return md.plain_heading(payload[1])
            if kind == "paragraph":
                return md.split_bold(payload)[0]
            if kind == "code":
                return payload                      # 펜스 줄은 넣지 않고 코드는 글자 그대로
            if kind in ("bullets", "numbered"):
                return md.split_bold(payload[0])[0]
            if kind == "table":
                return "".join(str(c) for c in payload[0])     # 칸은 그대로 넣음 ('\|'는 이미 글자 '|')
        return ""                                   # 구분선(빈 단락)·머리말만 있음
    if mode == "bold":
        return md.plain_heading(content.strip())
    first = next((ln.strip() for ln in content.splitlines() if ln.strip()), "")
    return md.split_bold(first)[0]


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
        # 찾을 문구의 앞부분을 보여 주지 않음 (오류 로그에 가린 값의 일부가 남았음). 몇 번째 문구인지와 길이만
        first = list(mapping).index(long_keys[0]) + 1
        raise rt.InputError(f"찾을 문구는 {FIND_MAX}자 이하여야 합니다 ({first}번째 찾을 문구가 {_find_len(long_keys[0])}자). "
                            "문장을 나눠서 지정하세요.")
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
    with off.open_target("word", args.target, readonly=args.dry_run, focus=not args.no_focus,
                         tool=TOOL, command="replace") as t:
        if args.dry_run:
            # 바꾸지 않음: 읽기 세션이라 되돌리기 지점도 남지 않음
            return _replace_dry_run(t.doc, t.resolved(), mapping, formatting, args)
        for find, rep in mapping.items():
            if find == rep and not formatting:
                # 바뀌는 것이 없음: 세기만 함 (바꾼 것으로 보고하지 않고 되돌리기 지점도 남기지 않음)
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


def _format_names(args) -> List[str]:
    names = []
    if args.bold is not None:
        names.append("굵게")
    if args.italic is not None:
        names.append("기울임")
    if args.color:
        names.append(f"글자색 {args.color}")
    if args.size:
        names.append(f"크기 {args.size:g}")
    if args.highlight:
        names.append(f"형광펜 {args.highlight}")
    return names


def _replace_plan_line(rng: Any, find: str, rep: str, fmt: str) -> str:
    """미리보기 한 줄: 찾은 곳 앞뒤 글의 바꾸기 전 -> 바꾼 뒤. 단락을 읽지 못하면 찾은 글만."""
    head = f"'{_preview(find, 20)}'"
    try:
        found = str(rng.Text)
        para = rng.Paragraphs(1).Range
        ptext = str(para.Text)
        at = int(rng.Start) - int(para.Start)
        if not 0 <= at <= len(ptext) or ptext[at:at + len(found)] != found:
            raise ValueError("위치가 맞지 않음")
        lo, hi = max(0, at - 15), at + len(found) + 15
        before = ("..." if lo else "") + ptext[lo:hi]
        after = ("..." if lo else "") + ptext[lo:at] + rep + ptext[at + len(found):hi]
        line = f"{head}: {_preview(before, 45)} -> {_preview(after, 45)}"
    except Exception:
        line = f"{head} -> '{_preview(rep, 20)}'"
    return line + (f" ({fmt})" if fmt else "")


def _replace_dry_run(doc: Any, resolved: Dict[str, Any], mapping: Dict[str, str], formatting: bool, args) -> Dict[str, Any]:
    """replace --dry-run: 실제 바꾸기와 같은 규칙으로 바뀔 곳만 세고 보여 줌 (문서는 바꾸지 않음)."""
    fmt = ", ".join(_format_names(args))
    counts: Dict[str, int] = {}
    same: Dict[str, int] = {}
    planned: List[str] = []
    limit_hit = False
    for find, rep in mapping.items():
        if find == rep and not formatting:
            # 실제 바꾸기와 같이 세기만 함 (바뀌는 것이 없음)
            hits = _count_text(doc, find, args.match_case, args.whole_word)
            if hits:
                same[find] = hits
    for story in _stories(doc):
        for find, rep in mapping.items():
            if find == rep and not formatting:
                continue
            n = 0
            rng = story.Duplicate
            while _find(rng, find, args.match_case, args.whole_word):
                if not formatting and str(rng.Text) == rep:
                    same[find] = same.get(find, 0) + 1          # 이미 바꿀 글 그대로: 바뀌지 않음
                else:
                    n += 1
                    if len(planned) <= rt.DRY_RUN_SHOW:
                        planned.append(_replace_plan_line(rng, find, rep, fmt))
                    if n >= 5000:
                        limit_hit = True
                        break
                rng.Collapse(0)
            if n:
                counts[find] = counts.get(find, 0) + n
    total = sum(counts.values())
    extra: Dict[str, Any] = {"target_resolved": resolved, "by_text": counts}
    if same:
        extra["unchanged_same_text"] = same
    if not total:
        found = sum(same.values())
        why = f"찾은 {found}곳은 이미 바꿀 글과 같습니다" if found else "일치하는 문구가 없습니다"
        return rt.dry_run_result(TOOL, "replace", f"미리보기: 바뀔 것 없음 (0곳), {why} (바꾸지 않음)", 0, planned, None, **extra)
    if limit_hit:
        extra["limit_reached"] = True
    more = " 이상" if limit_hit else ""
    preview = (f"미리보기: {total}곳{more} 바뀝니다: "
               + ", ".join(f"'{_preview(k, 20)}' {v}곳 -> '{_preview(mapping[k], 20)}'" for k, v in counts.items())
               + (f" ({fmt})" if fmt else "") + " (바꾸지 않음)")
    return rt.dry_run_result(TOOL, "replace", preview, total, planned,
                             rt.dry_run_command(TOOL, f"찾아 바꾸기 실행 ({total}곳{more})"), **extra)


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
    tracked = _has_revisions(doc)
    for story in _stories(doc):
        rng = story.Duplicate
        while n < limit and _find(rng, text, match_case, whole_word):
            if not (tracked and _deleted_revision(rng)):
                n += 1
            rng.Collapse(0)
    return n


def _has_revisions(doc: Any) -> bool:
    try:
        return int(doc.Revisions.Count) > 0
    except Exception:
        return False


def _deleted_revision(rng: Any) -> bool:
    """찾은 글이 변경 추적의 삭제 표시 글인지. Find는 삭제 표시 글도 찾아, 변경 추적 문서에서 바꾼 옛 글을 '남아 있음'으로
    세던 것 (회사 PC 2026-11-10 L9: read에는 새 글만 나오는데 replace가 UNVERIFIED)."""
    try:
        revs = rng.Revisions
        count = int(revs.Count)
        return count > 0 and all(int(revs(i).Type) == WD_REVISION_DELETE for i in range(1, count + 1))
    except Exception:
        return False


# ==============================================================================
# 3-1. fill: 서식 문서의 자리({{이름}})와 표 칸(표2!R3C4)에 값을 한 번에 채우기
# ==============================================================================

FILL_MAX = 500
_FILL_HEADER = ("자리", "위치", "place", "placeholder", "칸")
_FILL_PLACE = re.compile(r"^\{\{\s*[^{}\r\n]{1,100}?\s*\}\}$")
_FILL_CELL = re.compile(r"^\s*표\s*(\d{1,3})\s*!\s*R\s*(\d{1,4})\s*C\s*(\d{1,3})\s*$", re.IGNORECASE)
_FILL_FOUND = re.compile(r"\{\{[^{}\r\n]{1,100}?\}\}")
_CELL_END = "\r\x07"


def _fill_rows(path: str) -> List[Tuple[int, str, str]]:
    """값 파일(.csv/.tsv/.md 표): 자리, 값 -> [(줄 번호, 자리, 값)]. 첫 행이 머리글(자리/위치/place)이면 건너뜀. 값은 글자 그대로."""
    ext = os.path.splitext(path)[1].lower()
    if ext in (".csv", ".tsv"):
        rows = rt.read_csv_rows(path)
    else:
        table = next((payload for kind, payload in md.parse_markdown_blocks(rt.read_text_file(path)) if kind == "table"), None)
        if table is None:
            raise rt.InputError("값 파일에서 표를 찾지 못했습니다. CSV(열: 자리, 값)나 마크다운 표로 만드세요.")
        rows = table
    start = 1
    if rows and str(rows[0][0] if rows[0] else "").strip().lower() in _FILL_HEADER:
        rows, start = rows[1:], 2
    rt.check_two_columns(path, rows, start, "값 파일", "자리, 값", "채우지 않았습니다")
    return [(n, str(r[0] if r else "").strip(), "" if len(r) < 2 or r[1] is None else str(r[1]))
            for n, r in enumerate(rows, start) if any(str(v or "").strip() for v in r)]


_SLOT_COMMA = re.compile(r"(?:\s*[,|]\s*|\s+)(?:\{\{[^{}\r\n]{1,100}\}\}|표\s*\d{1,3}\s*!\s*R\s*\d{1,4}\s*C\s*\d{1,3}|[\w가-힣]{1,30})\s*=")


def _fully_quoted(val: str) -> bool:
    """값 전체가 한 쌍의 따옴표로 묶였는지 (첫 닫는 따옴표가 끝)."""
    return len(val) >= 2 and val[0] in "'\"" and val.find(val[0], 1) == len(val) - 1


def _slot_comma_split(val: str) -> bool:
    """값에 ',자리='가 있고 그 앞에 '='가 없으면 자리를 쉼표로 나눈 것 ('a,시험일=b'). '조건=V1=5, V2=12'의 값 'V1=5, V2=12'는 아님."""
    if _fully_quoted(val):
        return False
    m = _SLOT_COMMA.search(val)
    return bool(m) and "=" not in val[:m.start()]


def _split_slots(text: str) -> List[str]:
    """'업체명=한빛전자; 시험일=2026-11-04'를 ; 또는 줄바꿈으로 나눔. 따옴표로 묶은 값('a;b') 안의 ;는 나누지 않음."""
    parts: List[str] = []
    buf, quote = "", ""
    for ch in text:
        if quote:
            buf += ch
            if ch == quote:
                quote = ""
            continue
        if ch in "'\"" and buf.split("=", 1)[-1].strip() == "" and "=" in buf:
            quote = ch
            buf += ch
            continue
        if ch in ";\n":
            parts.append(buf)
            buf = ""
            if ch == "\n":
                parts.append("\n")              # 줄바꿈 표시: 다음 줄이 '자리='가 아니면 앞 값의 줄바꿈 (아래)
            continue
        buf += ch
    parts.append(buf)
    out: List[str] = []
    joined = False
    for p in parts:
        if p == "\n":
            joined = True
            continue
        if joined and out and "=" not in p.split("'", 1)[0].split('"', 1)[0] and p.strip():
            out[-1] += "\n" + p                  # '주소=서울\n강남구': 값 안의 줄바꿈
        elif p.strip():
            out.append(p)
        joined = False
    return out


def _slot_rows(args) -> List[Tuple[int, str, str]]:
    """--slots 글 -> [(번호, 자리, 값)]. 모양: 자리=값; 자리=값 (자리는 업체명, {{업체명}}, 표2!R3C4). JSON 객체({"업체명": "한빛전자"})도 받음.
    {{ }} 없이 적은 이름은 {{이름}} 자리로 봄. 쉼표로 나눈 것 같으면(값에 ',자리=') 채우지 않고 ;로 바꾼 명령을 [추천]."""
    text = rt.unescape_newlines(args.slots).strip()
    pairs: List[Tuple[str, str]] = []
    if text.startswith("{") and not text.startswith("{{"):
        try:
            obj = json.loads(text)
        except ValueError:
            raise rt.InputError("--slots는 '자리=값; 자리=값' 모양으로 적습니다 (예: --slots \"업체명=한빛전자; 시험일=2026-11-04\"). "
                                "JSON은 셸에서 따옴표가 깨지기 쉬워 권하지 않습니다.")
        if not isinstance(obj, dict):
            raise rt.InputError("--slots JSON은 {\"자리\": \"값\"} 객체여야 합니다. '자리=값; 자리=값' 모양을 권합니다.")
        odd = [str(k) for k, v in obj.items() if isinstance(v, (dict, list, bool))]
        if odd:
            raise rt.InputError(f"--slots JSON의 값은 글이나 숫자여야 합니다 ({', '.join(odd[:5])}). '자리=값; 자리=값' 모양을 권합니다.")
        pairs = [(str(k), "" if v is None else str(v)) for k, v in obj.items()]
    else:
        bad: List[str] = []
        for part in _split_slots(text):
            if "=" not in part:
                bad.append(f"'{part.strip()[:30]}': '자리=값' 모양이 아닙니다")
                continue
            key, val = part.split("=", 1)
            val = val.strip()
            if _fully_quoted(val):
                val = val[1:-1]
            elif _slot_comma_split(val):
                fixed = "; ".join(_SLOT_COMMA.sub(lambda m: "; " + re.sub(r"^\s*[,|]?\s*", "", m.group(0)), p.strip())
                                  if "=" in p and _slot_comma_split(p.split("=", 1)[1].strip()) else p.strip()
                                  for p in _split_slots(text))
                cmd = _fill_command(type(args)(**{**vars(args), "slots": fixed}), ["--dry-run"] if args.dry_run else [])["command"]
                raise rt.InputError(f"--slots의 자리는 ; 로 나눕니다. 쉼표로 나눈 것 같아 채우지 않았습니다 ('{key.strip()}'의 값에 '자리=' 이 들어 있음). "
                                    f"[추천] {cmd}")
            pairs.append((key.strip(), val))
        if bad:
            raise rt.InputError(f"--slots {len(bad)}개 항목이 잘못되어 채우지 않았습니다 (모양: 자리=값; 자리=값). " + " / ".join(bad[:10]))
    rows = []
    for n, (key, val) in enumerate(pairs, 1):
        if key and not _FILL_CELL.match(key) and not _FILL_PLACE.match(key):
            key = "{{" + key.strip("{} ") + "}}"
        rows.append((n, key, val))
    return rows


def _place_key(place: str) -> str:
    """자리 '{{ 업체 명 }}'의 비교 열쇠: 괄호 안 띄어쓰기를 모두 뺀 이름 ('업체명')."""
    return re.sub(r"\s+", "", place.strip()[2:-2])


def _cell_text(cell: Any) -> str:
    t = str(cell.Range.Text)
    return t[:-2] if t.endswith(_CELL_END) else t.rstrip("\r\x07")


def _cell_heads(doc: Any, ti: int, r: int, c: int) -> str:
    """표 칸 옆에 붙일 그 행 첫 칸 글과 그 열 첫 행(머리글) 글: " [행 '전압', 열 '측정']".
    병합된 표에서는 R3C4가 눈에 보이는 3행 4열과 다를 수 있어 미리보기로 맞는 칸인지 알아볼 수 있게. 못 읽으면 뺌."""
    def text(rr: int, cc: int) -> str:
        try:
            return _preview(_cell_text(doc.Tables(ti).Cell(rr, cc)), 12)
        except Exception:
            return ""
    row = text(r, 1) if c != 1 else ""
    col = text(1, c) if r != 1 else ""
    parts = ([f"행 '{row}'"] if row else []) + ([f"열 '{col}'"] if col else [])
    return f" [{', '.join(parts)}]" if parts else ""


def _doc_protected(doc: Any) -> bool:
    """편집 제한(문서 보호)이 걸렸는지. ProtectionType: -1 = 보호 없음. 읽지 못하면 보호 없음으로 봄 (쓰기에서 오류로 잡힘)."""
    try:
        return int(doc.ProtectionType) != -1
    except Exception:
        return False


def _fill_command(args, extra: Sequence[str] = ()) -> Dict[str, Any]:
    argv = ["fill", "--target", args.target] + (["--slots", args.slots] if args.slots else ["--file", args.file])
    for flag, on in (("--overwrite", args.overwrite), ("--no-focus", args.no_focus)):
        if on:
            argv.append(flag)
    argv += [x for x in extra if x not in argv]
    return rt.next_command("word.py", argv, "자리와 표 칸 채우기 실행" + (" (값이 있는 표 칸 덮어쓰기)" if "--overwrite" in argv else ""),
                           destructive="--overwrite" in argv)


FILL_HOWTO = ("값은 --slots \"<자리>=<값>; <자리>=<값>\" (자리는 양식의 {{이름}}에서 이름만 적어도 됨, 표 칸은 표2!R3C4, 자리마다 ; 로 구분) "
              "또는 --file 값.csv(열: 자리, 값)")


# ==============================================================================
# 3-2. delete (단락 지우기), table (표 행 넣기·지우기)
# ==============================================================================

DELETE_MAX = 200


def cmd_delete(args) -> Dict[str, Any]:
    """--paragraphs(read·outline의 단락 번호)의 단락을 지움. 표 안 단락은 지우지 않음 (표 행은 word.py table).
    --dry-run이면 지울 단락 글만 보여 줌. 지운 뒤 다시 읽어 남은 단락이 예상과 같은지 확인. office.py undo로 되돌림."""
    if not args.paragraphs:
        raise rt.InputError("지울 단락 번호를 --paragraphs로 주세요 (예: --paragraphs 12-14). 번호는 word.py read의 no입니다.")
    with off.open_target("word", args.target, readonly=bool(args.dry_run), focus=not args.no_focus,
                         tool=TOOL, command="delete") as t:
        doc = t.doc
        units = _com_units(doc, with_start=True)
        if not units:
            raise rt.InputError("문서에 글이 있는 단락이 없습니다.")
        nums = _spec_numbers(args.paragraphs, len(units), "--paragraphs")
        if len(nums) > DELETE_MAX:
            raise rt.InputError(f"한 번에 {DELETE_MAX}개 단락까지 지웁니다 (받은 범위 {len(nums)}개). 나눠서 지우세요.")
        picked = [units[n - 1] for n in nums]
        in_table = [n for n, u in zip(nums, picked) if _in_table(doc.Range(u["start"], u["start"]))]
        if in_table:
            raise rt.InputError(f"단락 {', '.join(map(str, in_table[:10]))}은(는) 표 안에 있어 지우지 않았습니다. 표의 행은 "
                                "word.py table --action delete-row로, 칸 글은 word.py fill로 바꾸세요.")
        planned = [f"단락 {n}: {_preview(u['text'], 60)}" for n, u in zip(nums, picked)]
        info: Dict[str, Any] = {"paragraphs": args.paragraphs, "count": len(nums), "deleted": planned[:rt.DRY_RUN_SHOW]}
        if args.dry_run:
            return rt.dry_run_result(TOOL, "delete", f"미리보기: 단락 {len(nums)}개를 지웁니다: " + " / ".join(planned[:3])
                                     + (f" 외 {len(nums) - 3}개" if len(nums) > 3 else "") + " (바꾸지 않음)", len(nums), planned,
                                     rt.dry_run_command(TOOL, f"단락 {len(nums)}개 지우기"), target_resolved=t.resolved())
        expect = [u["text"] for i, u in enumerate(units, 1) if i not in set(nums)]
        for u in sorted(picked, key=lambda x: -x["start"]):          # 뒤에서부터 (앞 위치가 밀리지 않게)
            doc.Range(u["start"], u["end"]).Delete()
        t.set_focus()
        left = [u["text"] for u in _com_units(doc)]
        resolved = t.resolved()
    preview = f"단락 {len(nums)}개를 지웠습니다: " + " / ".join(planned[:3]) + (f" 외 {len(nums) - 3}개" if len(nums) > 3 else "")
    if left != expect:
        return rt.ok_unverified(TOOL, "delete", preview, ["지운 뒤 남은 단락이 예상과 다름 (office.py undo로 되돌릴 수 있음)"],
                                target_resolved=resolved, **info)
    return rt.ok(TOOL, "delete", preview, "VERIFIED: 남은 단락을 다시 읽어 확인했습니다. 같은 명령을 다시 실행하지 마세요 (번호가 바뀌어 "
                 "다른 단락이 지워짐). 잘못 지웠으면 office.py undo로 되돌립니다.", verified=True, target_resolved=resolved, **info)


WD_ALIGN = {"left": 0, "center": 1, "right": 2, "justify": 3}
FORMAT_MAX = 500


def _style_name(rng: Any) -> str:
    try:
        st = rng.Style
        return str(getattr(st, "NameLocal", st) or "")
    except Exception:
        return ""


LINKS_SHOW = 100
_BOOKMARK_NAME = re.compile(r"^[^\W\d_][\w]{0,39}$")


def cmd_link(args) -> Dict[str, Any]:
    """하이퍼링크 목록(--anchor 없음)과 달기(--anchor 글에 --address 주소: 파일 경로, http 주소, #책갈피). 1.6.7 기반 기능.
    문서 글은 바꾸지 않고 링크만 덧붙임, 달고 나서 다시 읽어 확인."""
    if not args.anchor:
        with off.open_target("word", args.target, readonly=True, tool=TOOL, command="link") as t:
            doc = t.doc
            n = int(doc.Hyperlinks.Count)
            items = []
            for i in range(1, min(n, LINKS_SHOW) + 1):
                h = doc.Hyperlinks(i)
                sub = str(getattr(h, "SubAddress", "") or "")
                items.append({"no": i, "text": _preview(str(h.TextToDisplay or ""), 60), "address": str(h.Address or "") + (f"#{sub}" if sub else "")})
            resolved = t.resolved()
        if not n:
            return rt.ok(TOOL, "link", f"0건: {resolved['name']}에 하이퍼링크가 없습니다", "VERIFIED: 0건. 하이퍼링크가 없습니다 (찾은 항목 없음).",
                         links=[], count=0, target_resolved=resolved)
        return rt.ok(TOOL, "link", f"{resolved['name']}: 하이퍼링크 {n}개", "VERIFIED: 하이퍼링크를 읽었습니다.", links=items, count=n,
                     truncated=n > LINKS_SHOW, target_resolved=resolved)
    if not args.address:
        raise rt.InputError("링크 주소를 --address로 주세요 (파일 경로, http 주소, 책갈피는 #이름).", no_fallback=True)
    address = str(args.address).strip().strip('"')
    sub = ""
    if address.startswith("#"):
        address, sub = "", address[1:]
    with off.open_target("word", args.target, readonly=bool(args.dry_run), focus=not args.no_focus, tool=TOOL, command="link") as t:
        doc = t.doc
        rng = _comment_range(doc, args.anchor, args.occurrence)
        shown = _preview(str(rng.Text), 40)
        if args.dry_run:
            return rt.dry_run_result(TOOL, "link", f"미리보기: '{shown}'에 링크 {args.address} (바꾸지 않음)", 1,
                                     [f"'{shown}' -> {args.address}"], rt.dry_run_command(TOOL, "링크 달기"), target_resolved=t.resolved())
        before = int(doc.Hyperlinks.Count)
        doc.Hyperlinks.Add(Anchor=rng, Address=address, SubAddress=sub)
        after = int(doc.Hyperlinks.Count)
        last = doc.Hyperlinks(after) if after else None
        got = (str(getattr(last, "Address", "") or ""), str(getattr(last, "SubAddress", "") or "")) if last is not None else ("", "")
        t.set_focus()
        resolved = t.resolved()
    preview = f"'{shown}'에 링크를 달았습니다 ({args.address})"
    if after != before + 1 or got != (address, sub):
        return rt.ok_unverified(TOOL, "link", preview, ["달고 다시 읽은 링크 수나 주소가 다름"], target_resolved=resolved)
    return rt.ok(TOOL, "link", preview, "VERIFIED: 링크를 다시 읽어 확인했습니다. 되돌리려면 office.py undo.", verified=True,
                 target_resolved=resolved, address=args.address)


def cmd_bookmark(args) -> Dict[str, Any]:
    """책갈피 목록(--name 없음)과 만들기(--name 이름 --anchor 글). 같은 이름이 있으면 --overwrite가 있을 때만 옮김 (1.6.7 기반 기능).
    링크로 문서 안의 곳을 가리키려면 word.py link --address "#이름"."""
    if not args.name:
        with off.open_target("word", args.target, readonly=True, tool=TOOL, command="bookmark") as t:
            doc = t.doc
            n = int(doc.Bookmarks.Count)
            items = [{"name": str(doc.Bookmarks(i).Name), "text": _preview(str(doc.Bookmarks(i).Range.Text or ""), 60)}
                     for i in range(1, min(n, LINKS_SHOW) + 1)]
            resolved = t.resolved()
        if not n:
            return rt.ok(TOOL, "bookmark", f"0건: {resolved['name']}에 책갈피가 없습니다", "VERIFIED: 0건. 책갈피가 없습니다 (찾은 항목 없음).",
                         bookmarks=[], count=0, target_resolved=resolved)
        return rt.ok(TOOL, "bookmark", f"{resolved['name']}: 책갈피 {n}개", "VERIFIED: 책갈피를 읽었습니다.", bookmarks=items, count=n,
                     truncated=n > LINKS_SHOW, target_resolved=resolved)
    name = str(args.name).strip()
    if not _BOOKMARK_NAME.match(name):
        raise rt.InputError(f"책갈피 이름 '{name}'은(는) 쓸 수 없습니다: 글자로 시작하고 글자·숫자·밑줄만, 40자까지 (공백 불가).", no_fallback=True)
    if not args.anchor:
        raise rt.InputError("책갈피를 둘 글을 --anchor로 주세요.", no_fallback=True)
    with off.open_target("word", args.target, readonly=bool(args.dry_run), focus=not args.no_focus, tool=TOOL, command="bookmark") as t:
        doc = t.doc
        rng = _comment_range(doc, args.anchor, args.occurrence)
        exists = bool(doc.Bookmarks.Exists(name))
        if exists and not args.overwrite and not args.dry_run:
            raise rt.UserDecisionError(f"책갈피 '{name}'이(가) 이미 있습니다.", "바꾸지 않았습니다. 그 책갈피를 이 글로 옮길지 사용자에게 묻고, "
                                       "'예'이면 confirm_commands[0]을 1번 실행하세요.", candidates=["책갈피를 이 글로 옮기기"],
                                       commands=[rt.retry_command({"--overwrite": True})], destructive=True)
        shown = _preview(str(rng.Text), 40)
        if args.dry_run:
            return rt.dry_run_result(TOOL, "bookmark", f"미리보기: '{shown}'에 책갈피 '{name}' (바꾸지 않음)", 1, [f"{name} -> '{shown}'"],
                                     rt.dry_run_command(TOOL, f"책갈피 '{name}' 만들기", destructive=exists, overwrite=exists),
                                     target_resolved=t.resolved())
        doc.Bookmarks.Add(name, rng)
        ok = bool(doc.Bookmarks.Exists(name))
        t.set_focus()
        resolved = t.resolved()
    preview = f"'{shown}'에 책갈피 '{name}'을(를) 만들었습니다"
    if not ok:
        return rt.ok_unverified(TOOL, "bookmark", preview, ["만든 뒤 책갈피를 찾지 못함"], target_resolved=resolved)
    return rt.ok(TOOL, "bookmark", preview, f"VERIFIED: 책갈피를 확인했습니다. 링크로 가리키려면 word.py link --address \"#{name}\".",
                 verified=True, target_resolved=resolved, name=name)


HF_TYPES = (("기본", 1), ("첫 쪽", 2), ("짝수 쪽", 3))    # wdHeaderFooterPrimary/FirstPage/EvenPages


def _hf(sec: Any, kind: str, idx: int) -> Any:
    return sec.Headers(idx) if kind == "header" else sec.Footers(idx)


def _hf_text(hf: Any) -> str:
    return str(hf.Range.Text or "").replace("\x07", "").rstrip("\r\n")


def cmd_header(args) -> Dict[str, Any]:
    """머리글·바닥글 읽기(--text 없음)와 쓰기(--text: --section의 기본 머리글 또는 바닥글을 그 글로). 기반 기능 (1.6.7):
    문서 번호·개정 판 표기 확인과 수정. 글이 있으면 --overwrite가 있을 때만 바꿈 (R4), 쓴 뒤 다시 읽어 확인."""
    kinds = [args.kind] if args.kind else ["header", "footer"]
    names = {"header": "머리글", "footer": "바닥글"}
    if args.text is None:
        with off.open_target("word", args.target, readonly=True, tool=TOOL, command="header") as t:
            doc = t.doc
            n = int(doc.Sections.Count)
            if args.section and not 1 <= args.section <= n:
                raise rt.InputError(f"--section {args.section}: 이 문서의 구역은 1~{n}입니다.")
            items: List[Dict[str, Any]] = []
            for si in ([args.section] if args.section else range(1, n + 1)):
                sec = doc.Sections(si)
                for kind in kinds:
                    for tname, idx in HF_TYPES:
                        hf = _hf(sec, kind, idx)
                        if idx > 1 and not bool(getattr(hf, "Exists", False)):
                            continue
                        items.append({"section": si, "kind": names[kind], "type": tname, "text": _preview(_hf_text(hf), 500)})
            resolved = t.resolved()
        filled = [i for i in items if i["text"]]
        where = f"{resolved.get('name', '문서')}"
        if not filled:
            return rt.ok(TOOL, "header", f"0건: {where}에 글이 있는 {'·'.join(names[k] for k in kinds)}이(가) 없습니다",
                         f"VERIFIED: 0건. {'·'.join(names[k] for k in kinds)}이(가) 비어 있습니다 (찾은 항목 없음). 사용자에게 그대로 알리세요.",
                         items=items, count=0, sections=n, target_resolved=resolved)
        shown = "; ".join(f"구역 {i['section']} {i['kind']}({i['type']}): {_preview(i['text'], 40)}" for i in filled[:5])
        return rt.ok(TOOL, "header", f"{where}: {shown}", "VERIFIED: 머리글·바닥글을 읽었습니다.", items=items, count=len(filled),
                     sections=n, target_resolved=resolved)
    if not args.kind:
        raise rt.InputError("쓸 곳을 --kind header 또는 --kind footer로 주세요.", no_fallback=True)
    text = rt.unescape_newlines(args.text)
    si = args.section or 1
    with off.open_target("word", args.target, readonly=bool(args.dry_run), focus=not args.no_focus, tool=TOOL, command="header") as t:
        doc = t.doc
        n = int(doc.Sections.Count)
        if not 1 <= si <= n:
            raise rt.InputError(f"--section {si}: 이 문서의 구역은 1~{n}입니다.")
        hf = _hf(doc.Sections(si), args.kind, 1)
        old = _hf_text(hf)
        what = f"구역 {si} {names[args.kind]}"
        if old and old != text and not args.overwrite and not args.dry_run:
            raise rt.UserDecisionError(f"{what}에 이미 글이 있습니다: '{_preview(old, 60)}'.",
                                       "바꾸지 않았습니다. 있던 글을 새 글로 바꿀지 사용자에게 묻고, '예'이면 confirm_commands[0]을 1번 실행하세요.",
                                       candidates=["있던 글을 새 글로 바꾸기"], commands=[rt.retry_command({"--overwrite": True})], destructive=True)
        if args.dry_run:
            return rt.dry_run_result(TOOL, "header", f"미리보기: {what}을(를) '{_preview(text, 60)}'(으)로 (바꾸지 않음)", 1,
                                     [f"{what}: '{_preview(old, 60)}' -> '{_preview(text, 60)}'"],
                                     rt.dry_run_command(TOOL, f"{what} 바꾸기", destructive=bool(old), overwrite=bool(old)),
                                     target_resolved=t.resolved())
        hf.Range.Text = text
        back = _hf_text(_hf(doc.Sections(si), args.kind, 1))
        linked = si > 1 and bool(getattr(hf, "LinkToPrevious", False))
        t.set_focus()
        resolved = t.resolved()
    preview = f"{what}을(를) '{_preview(text, 60)}'(으)로 바꿨습니다" + (" (앞 구역과 연결되어 있어 앞 구역도 같이 바뀜)" if linked else "")
    if back != text.replace("\n", "\r").rstrip("\r"):
        return rt.ok_unverified(TOOL, "header", preview, [f"다시 읽은 글이 다름: '{_preview(back, 60)}'"], target_resolved=resolved, old_text=old)
    return rt.ok(TOOL, "header", preview, "VERIFIED: 쓴 뒤 다시 읽어 같은 것을 확인했습니다. 되돌리려면 office.py undo.",
                 verified=True, target_resolved=resolved, old_text=old, section=si)


def cmd_format(args) -> Dict[str, Any]:
    """--paragraphs(read의 단락 번호)의 정렬·글꼴·크기·굵게·기울임·스타일을 바꿈 (README에만, 스킬에는 적지 않음: 범위 고르기가 어려움).
    스타일을 먼저 적용하고 그 위에 글꼴 옵션. 바꾼 뒤 단락마다 다시 읽어 확인, 되돌리기 지점."""
    if not args.paragraphs:
        raise rt.InputError("바꿀 단락 번호를 --paragraphs로 주세요 (예: --paragraphs 3-5). 번호는 word.py read의 no입니다.")
    if not any([args.align, args.font, args.size, args.bold, args.italic, args.style]):
        raise rt.InputError("바꿀 서식이 없습니다. --align, --font, --size, --bold, --italic, --style 중 하나 이상을 주세요.")
    if args.size is not None and not 1 <= args.size <= 400:
        raise rt.InputError(f"--size는 1~400pt입니다 (받은 값 {args.size}).")
    with off.open_target("word", args.target, readonly=bool(args.dry_run), focus=not args.no_focus, tool=TOOL, command="format") as t:
        doc = t.doc
        units = _com_units(doc, with_start=True)
        if not units:
            raise rt.InputError("문서에 글이 있는 단락이 없습니다.")
        nums = _spec_numbers(args.paragraphs, len(units), "--paragraphs")
        if len(nums) > FORMAT_MAX:
            raise rt.InputError(f"한 번에 {FORMAT_MAX}개 단락까지 바꿉니다 (받은 범위 {len(nums)}개).")
        picked = [units[n - 1] for n in nums]
        style = None
        if args.style:
            try:
                style = doc.Styles(args.style)
            except Exception:
                raise rt.InputError(f"스타일 '{args.style}'이(가) 이 문서에 없습니다. Word 스타일 이름 그대로 쓰세요 (예: 표준, 제목 1).")
        what = ", ".join(x for x in (args.style and f"스타일 {args.style}", args.align and f"정렬 {args.align}", args.font and f"글꼴 {args.font}",
                                     args.size and f"크기 {args.size:g}", args.bold and "굵게", args.italic and "기울임") if x)
        planned = [f"단락 {n}: {_preview(u['text'], 50)}" for n, u in zip(nums, picked)]
        info: Dict[str, Any] = {"paragraphs": args.paragraphs, "count": len(nums), "format": what}
        if args.dry_run:
            return rt.dry_run_result(TOOL, "format", f"미리보기: 단락 {len(nums)}개에 {what} (바꾸지 않음)", len(nums), planned,
                                     rt.dry_run_command(TOOL, f"단락 {len(nums)}개 서식 바꾸기", destructive=False), target_resolved=t.resolved())
        problems: List[str] = []
        for n, u in zip(nums, picked):
            rng = doc.Range(u["start"], u["end"])
            if style is not None:
                rng.Style = style
            if args.align:
                rng.ParagraphFormat.Alignment = WD_ALIGN[args.align]
            if args.font:
                # 한글은 동아시아 글꼴(NameFarEast)로 그려짐: Name만 바꾸면 영문만 바뀌고 한글은 그대로
                rng.Font.Name = args.font
                try:
                    rng.Font.NameFarEast = args.font
                except Exception:
                    pass
            if args.size:
                rng.Font.Size = float(args.size)
            if args.bold:
                rng.Font.Bold = True
            if args.italic:
                rng.Font.Italic = True
            chk = doc.Range(u["start"], u["end"])
            bad = []
            if args.align and int(chk.ParagraphFormat.Alignment) != WD_ALIGN[args.align]:
                bad.append("정렬")
            if args.font and (str(chk.Font.Name) != args.font or str(getattr(chk.Font, "NameFarEast", args.font) or args.font) != args.font):
                bad.append("글꼴")
            if args.size and abs(float(chk.Font.Size or 0) - float(args.size)) > 0.01:
                bad.append("크기")
            if args.bold and chk.Font.Bold not in (True, -1, 1):
                bad.append("굵게")
            if args.italic and chk.Font.Italic not in (True, -1, 1):
                bad.append("기울임")
            if args.style and _style_name(chk).casefold() != str(getattr(style, "NameLocal", args.style)).casefold():
                bad.append("스타일")
            if bad:
                problems.append(f"단락 {n}: {', '.join(bad)}이(가) 다르게 읽힘")
        t.set_focus()
        resolved = t.resolved()
    preview = f"단락 {len(nums)}개에 {what}을(를) 적용했습니다"
    if problems:
        return rt.ok_unverified(TOOL, "format", preview, problems[:5], target_resolved=resolved, **info)
    return rt.ok(TOOL, "format", preview, "VERIFIED: 단락마다 서식을 다시 읽어 확인했습니다. 잘못 바꿨으면 office.py undo로 되돌립니다.",
                 verified=True, target_resolved=resolved, **info)


def _row_cells(row: Any) -> List[str]:
    return [c for c in str(row.Range.Text).split("\r\x07")[:-1]][: int(row.Cells.Count)]


def cmd_table(args) -> Dict[str, Any]:
    """워드 표의 행 넣기(add-row: --row 행 뒤, 생략하면 끝, --values '칸; 칸')와 지우기(delete-row: --row).
    표 번호는 word.py info의 tables no. 세로 병합된 표는 행 단위로 다룰 수 없어 멈춤."""
    if args.action == "delete-row" and not args.row:
        raise rt.InputError("지울 행 번호를 --row로 주세요 (1 = 첫 행, 머리글 포함).")
    values = [v.strip() for v in args.values.split(";")] if args.values is not None else []
    if args.action == "delete-row" and args.values is not None:
        raise rt.InputError("--values는 add-row(행 넣기)에만 씁니다. [추천] " + rt.retry_command({"--values": None}))
    with off.open_target("word", args.target, readonly=bool(args.dry_run), focus=not args.no_focus,
                         tool=TOOL, command="table") as t:
        doc = t.doc
        count = int(doc.Tables.Count)
        if count == 0:
            raise rt.InputError("문서에 표가 없습니다.")
        if not 1 <= args.table <= count:
            raise rt.InputError(f"표 {args.table}이(가) 없습니다 (1~{count}). 표 번호는 word.py info의 tables no입니다.")
        tb = doc.Tables(args.table)
        rows = int(tb.Rows.Count)
        try:
            if args.row is not None and not 1 <= args.row <= rows:
                raise rt.InputError(f"표 {args.table}에 {args.row}행이 없습니다 (1~{rows}).")
            ref = tb.Rows(args.row or rows)
            width = int(ref.Cells.Count)
            ref_text = _row_cells(ref)
        except rt.InputError:
            raise
        except Exception as e:
            raise rt.InputError(f"표 {args.table}은(는) 세로로 병합된 칸이 있어 행을 하나씩 다룰 수 없습니다 ({type(e).__name__}). "
                                "Word에서 병합을 풀거나 직접 고치라고 사용자에게 알리세요.", no_fallback=True)
        if len(values) > width:
            raise rt.InputError(f"--values가 {len(values)}칸인데 표 {args.table}의 행은 {width}칸입니다. 칸은 ; 로 나눕니다.")
        where = f"표 {args.table}"
        if args.action == "delete-row":
            desc = f"{where} {args.row}행 지우기: {_preview(' | '.join(ref_text), 60)}"
        else:
            pos = args.row if args.row else rows
            desc = f"{where} {pos}행 뒤에 행 넣기" + (f": {' | '.join(values)}" if values else " (빈 행)")
        info: Dict[str, Any] = {"table": args.table, "action": args.action, "rows_before": rows}
        if args.dry_run:
            return rt.dry_run_result(TOOL, "table", f"미리보기: {desc} (바꾸지 않음)", 1, [desc],
                                     rt.dry_run_command(TOOL, desc), target_resolved=t.resolved(), **info)
        problems: List[str] = []
        if args.action == "delete-row":
            ref.Delete()
            got_rows = int(tb.Rows.Count)
            if got_rows != rows - 1:
                problems.append(f"행 수가 {got_rows} (지우기 전 {rows})")
        else:
            before = tb.Rows(args.row + 1) if args.row and args.row < rows else None
            new = tb.Rows.Add(before) if before is not None else tb.Rows.Add()
            new_no = (args.row + 1) if before is not None else rows + 1
            for j, v in enumerate(values, 1):
                tb.Cell(new_no, j).Range.Text = v
            got_rows = int(tb.Rows.Count)
            got = _row_cells(tb.Rows(new_no))
            if got_rows != rows + 1:
                problems.append(f"행 수가 {got_rows} (넣기 전 {rows})")
            if [g.strip() for g in got[:len(values)]] != values:
                problems.append(f"새 행의 글이 다름 ({' | '.join(got)})")
            info["row"] = new_no
        info["rows_after"] = got_rows
        t.set_focus()
        resolved = t.resolved()
    preview = desc.replace("지우기:", "을(를) 지웠습니다:").replace("행 넣기", "행을 넣었습니다")
    if problems:
        return rt.ok_unverified(TOOL, "table", preview, problems, target_resolved=resolved, **info)
    return rt.ok(TOOL, "table", preview, "VERIFIED: 표를 다시 읽어 확인했습니다. 같은 명령을 다시 실행하지 마세요. "
                 "잘못 바꿨으면 office.py undo로 되돌립니다.", verified=True, target_resolved=resolved, **info)


def cmd_fill(args) -> Dict[str, Any]:
    if bool(args.slots) == bool(args.file):
        raise rt.InputError(("--slots와 --file 중 하나만 씁니다. " if args.slots else "채울 값이 없습니다. ") + FILL_HOWTO +
                            ". 채울 양식 문서는 --target으로 줍니다.")
    if args.file and os.path.splitext(args.file.strip().strip('"'))[1].lower() in (".docx", ".doc", ".docm", ".dotx", ".hwp", ".hwpx"):
        # 양식 문서를 --file로 준 경우 (회사 PC 2026-11-06 O5: 인코딩 인식 실패로 끝남)
        raise rt.InputError(f"--file은 채울 값 파일입니다. 양식 문서 '{os.path.basename(args.file.strip().strip(chr(34)))}'은(는) --target으로 주고, "
                            + FILL_HOWTO + ".")
    rows = _slot_rows(args) if args.slots else \
        _fill_rows(rt.resolve_input_path(args.file, "값 파일", exts=(".csv", ".tsv", ".md", ".txt")))
    if not rows:
        raise rt.InputError("채울 값이 없습니다. " + FILL_HOWTO + ".")
    if len(rows) > FILL_MAX:
        raise rt.InputError(f"한 번에 {FILL_MAX}곳까지 채웁니다 (받은 {len(rows)}곳). 나눠 실행하세요.")
    bad: List[str] = []
    unit = "번째 항목" if args.slots else "번째 줄"
    places: Dict[str, str] = {}
    cells: Dict[Tuple[int, int, int], str] = {}
    for n, key, val in rows:
        m = _FILL_CELL.match(key)
        if m:
            k = (int(m.group(1)), int(m.group(2)), int(m.group(3)))
            if k in cells and cells[k] != val:
                bad.append(f"{n}{unit} '{key}': 같은 칸에 다른 값이 두 번 있습니다")
            cells[k] = val
        elif _FILL_PLACE.match(key):
            if key in places and places[key] != val:
                bad.append(f"{n}{unit} '{key}': 같은 자리에 다른 값이 두 번 있습니다")
            places[key] = val
        else:
            bad.append(f"{n}{unit} '{key[:30]}': 자리는 {{{{이름}}}} 또는 표 칸 표2!R3C4 모양이어야 합니다")
    if bad:
        head = (f"--slots {len(rows)}개 중 {len(bad)}개가 잘못되어 하나도 채우지 않았습니다 (모양: 자리=값; 자리=값). " if args.slots else
                f"값 파일 {len(rows)}줄 중 {len(bad)}줄이 잘못되어 하나도 채우지 않았습니다 (열: 자리, 값). ") + "고친 뒤 다시 실행하세요. "
        raise rt.InputError(head + " / ".join(bad[:10]), log_message=head + f"잘못된 줄 {len(bad)}개 ({rt.DOC_OMITTED})")
    with off.open_target("word", args.target, readonly=args.dry_run, focus=not args.no_focus,
                         tool=TOOL, command="fill") as t:
        doc = t.doc
        if _doc_protected(doc):
            raise rt.InputError("문서에 편집 제한(문서 보호)이 걸려 있어 아무것도 채우지 않았습니다. 사용자가 Word에서 보호를 풀었다고 하면 "
                                "같은 명령을 1번 다시 실행하세요 (암호를 묻거나 추측하지 않음).", no_fallback=True)
        # 1) 모두 확인: 자리가 문서에 있는지, 표 칸이 있는지, 값이 있는 칸 (하나라도 문제면 아무것도 바꾸지 않음)
        counts = {k: _count_text(doc, k, True, False) for k in places}
        missing = [k for k, c in counts.items() if not c]
        problems_in: List[str] = []
        if missing:
            text = "".join(str(s.Text) for s in _stories(doc))
            found = sorted(set(_FILL_FOUND.findall(text)))
            # 괄호 안 띄어쓰기만 다른 자리('{{ 업체명 }}' <-> '{{업체명}}')는 문서의 자리 하나와만 맞으면 그 자리로 채움 (뜻이 같음)
            for k in missing:
                alt = [f for f in found if _place_key(f) == _place_key(k) and f not in places]
                if len(alt) == 1:
                    places = {(alt[0] if kk == k else kk): vv for kk, vv in places.items()}
                    counts = {(alt[0] if kk == k else kk): (_count_text(doc, alt[0], True, False) if kk == k else cc)
                              for kk, cc in counts.items()}
                    rt.note_recovered(f"자리 {k} -> 문서의 {alt[0]}(띄어쓰기만 다름)로 채움")
            missing = [k for k, c in counts.items() if not c]
        if missing:
            problems_in.append(f"문서에 없는 자리 {len(missing)}개: {', '.join(missing[:10])}"
                               + (f" (문서에 있는 자리: {', '.join(found[:10])})" if found else " (문서에 {{...}} 자리가 없음)"))
        try:
            n_tables = int(doc.Tables.Count)
        except Exception:
            n_tables = 0
        old: Dict[Tuple[int, int, int], str] = {}
        for (ti, r, c) in cells:
            ref = f"표{ti}!R{r}C{c}"
            if not 1 <= ti <= n_tables:
                problems_in.append(f"{ref}: 표 {ti}번이 없습니다 (표 {n_tables}개)")
                continue
            try:
                old[(ti, r, c)] = _cell_text(doc.Tables(ti).Cell(r, c))
            except Exception:
                problems_in.append(f"{ref}: 그 칸이 없습니다 (행·열 번호 또는 병합된 칸)")
        if problems_in:
            raise rt.InputError("아무것도 채우지 않았습니다. " + " / ".join(problems_in[:10]))
        conflicts = [(k, v) for k, v in old.items() if v.strip() and v != cells[k]]
        same = [k for k, v in old.items() if v == cells[k]]
        heads = {k: _cell_heads(doc, *k) for k in cells}
        planned = [f"바꾸기: {k} {counts[k]}곳 -> {_preview(v)}" for k, v in places.items()]
        for k, v in cells.items():
            ref = f"표{k[0]}!R{k[1]}C{k[2]}{heads[k]}"
            if k in same:
                planned.append(f"그대로(같은 값): {ref}")
            elif old[k].strip():
                planned.append(f"덮어쓰기: {ref} ({_preview(old[k], 15)} -> {_preview(v, 15)})")
            else:
                planned.append(f"채우기: {ref} (빈 칸 -> {_preview(v, 15)})")
        if args.dry_run:
            count = len(places) + len(cells) - len(same)
            extra: Dict[str, Any] = {"target_resolved": t.resolved()}
            if conflicts:
                extra["overwrite_cells"] = len(conflicts)
            preview = (f"미리보기: 자리 {len(places)}개, 표 칸 {len(cells)}개" + (f" (값이 있는 칸 {len(conflicts)}개)" if conflicts else "")
                       if count else f"미리보기: 바뀔 것 없음 (0곳, 이미 같은 값 {len(same)}칸)")
            return rt.dry_run_result(TOOL, "fill", preview + " (바꾸지 않음)", count, planned,
                                     _fill_command(args, ["--overwrite"] if conflicts else []),
                                     rt.DRY_RUN_DIRECTIVE + " 표 칸은 planned의 [행, 열] 글로 맞는 칸인지 함께 알리세요.", **extra)
        if conflicts and not args.overwrite:
            refs = [f"표{k[0]}!R{k[1]}C{k[2]}" for k, _ in conflicts]
            head = f"값이 있는 표 칸 {len(conflicts)}개가 덮어써집니다: "
            raise rt.UserDecisionError(
                head + ", ".join(f"{r}{heads[k]}='{_preview(v, 15)}'" for r, (k, v) in zip(refs[:5], conflicts[:5]))
                + (" ..." if len(refs) > 5 else ""),
                "채우지 않았습니다. 덮어쓸 칸과 기존 값을 사용자에게 보여 주고, 사용자가 덮어써도 된다고 하면 같은 명령에 --overwrite를 붙여 "
                "다시 실행하세요 (미리 보려면 --dry-run).",
                log_message=head + ", ".join(refs[:5]) + f" (값 {rt.DOC_OMITTED})")
        # 2) 채우기 (되돌리기 지점은 open_target이 남김)
        done = 0
        where = ""
        try:
            for story in _stories(doc):
                for key, val in places.items():
                    where = key
                    rng = story.Duplicate
                    n = 0
                    while n < 5000 and _find(rng, key, True, False):
                        rng.Text = val.replace("\r\n", "\v").replace("\n", "\v")      # 자리 안의 줄바꿈은 줄 바꿈(단락 나누지 않음)
                        t.changed = True
                        done += 1
                        n += 1
                        rng.Collapse(0)
            for k, v in cells.items():
                if k in same:
                    continue
                where = f"표{k[0]}!R{k[1]}C{k[2]}"
                doc.Tables(k[0]).Cell(k[1], k[2]).Range.Text = v.replace("\r\n", "\r").replace("\n", "\r")
                t.changed = True
                done += 1
        except Exception as e:
            info = rt._com_error_info(e)
            if info is not None and info[2] in ("busy", "gone"):
                raise
            rt.debug(f"채우기 실패: {e}")
            if done and t.source in ("active", "open"):
                left = f"앞의 {done}곳은 이미 바뀌었습니다(office.py undo로 되돌릴 수 있음). "
            elif done:
                left = "파일은 저장하지 않아 그대로입니다. "
            else:
                left = "아무것도 바꾸지 않았습니다. "
            raise rt.InputError(f"{where}을(를) 채우지 못했습니다 (문서 보호, 편집 중, 내용 컨트롤 잠금 등, {type(e).__name__}). "
                                + left + "원인을 사용자에게 알리고 해결했다고 하면 같은 명령을 1번 다시 실행하세요.",
                                log_message=f"{where if where.startswith('표') else '자리'}을(를) 채우지 못했습니다 ({type(e).__name__}). "
                                + left, no_fallback=True)
        # 3) 다시 읽어 확인 (R1)
        problems: List[str] = []
        for key, val in places.items():
            left = _count_text(doc, key, True, False)
            if left and key not in val:
                problems.append(f"{key}이(가) {left}곳 남아 있음")
        for k, v in cells.items():
            got = _cell_text(doc.Tables(k[0]).Cell(k[1], k[2])).replace("\r", "\n")
            if got != v.replace("\r\n", "\n"):
                problems.append(f"표{k[0]}!R{k[1]}C{k[2]}의 글이 쓴 값과 다름")
        if t.changed:
            t.set_focus()
        resolved = t.resolved()
    fields = {"target_resolved": resolved, "count": len(places) + len(cells) - len(same),
              "places": {k: counts[k] for k in places}, "cells": len(cells) - len(same),
              "applied": [p for p in planned if not p.startswith("그대로")][:50]}
    if same:
        fields["unchanged"] = len(same)
    if conflicts:
        fields["overwritten"] = len(conflicts)
    elif args.overwrite:      # 사용법 단순성 4번: 쓰이지 않은 옵션은 받고 warnings로 알림 (집 PC Gemma O5)
        fields["warnings"] = ["값이 있는 표 칸이 없어 --overwrite로 덮어쓴 칸은 없습니다 (빈 칸과 자리만 채움)"]
    preview = f"자리 {len(places)}개({sum(counts.values())}곳), 표 칸 {len(cells) - len(same)}개를 채웠습니다" \
        + (f" (이미 같은 값 {len(same)}칸은 그대로)" if same else "")
    if problems:
        return rt.ok_unverified(TOOL, "fill", preview, problems[:10], **fields)
    return rt.ok(TOOL, "fill", preview, "VERIFIED: 채운 곳을 다시 읽어 확인했습니다. 잘못 채웠으면 office.py undo로 되돌릴 수 있습니다.",
                 verified=True, **fields)


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
        out = rt.resolve_output_path(args.save, os.path.splitext(t.name)[0], ext, near=t.path)
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
            if tmp:
                rt.remove_file(tmp)
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
        p.add_argument("--target", default="active", help="active(기본) | 열린 문서 이름 | 파일 경로")
        if write:
            p.add_argument("--no-focus", action="store_true", help="작업 후 화면 이동 안 함")
            p.add_argument("--dry-run", action="store_true", help="바꾸지 않고 바뀔 것만 미리 보기 (미리보기가 없는 동작은 바꾸지 않고 멈춤, 목록은 그대로 읽음)")

    p = sub.add_parser("read", help="내용 읽기")
    common(p)
    p.add_argument("--start", type=int, default=1, help="이 번호의 단락부터 읽기")
    p.add_argument("--pages", help="이 쪽만 읽기 (예: 3, 2-5, 1,4,7)")
    p.add_argument("--paragraphs", help="이 단락만 읽기 (예: 120-140, 141-, 7). 번호는 read/outline의 단락 번호")
    p.add_argument("--section", help="이 제목의 장절만 읽기 (제목 글 일부, 예: \"1.1\", \"검증 방법\"). 하위 절 포함")
    p.add_argument("--occurrence", type=int, help="--section 제목이 여러 개일 때 몇 번째인지 (오류의 [후보] 번호)")
    p.add_argument("--save", nargs="?", const=True, help="md 파일로 저장 (경로 생략 시 대상 파일 옆, 없으면 C:\\AIKIT\\data)")

    p = sub.add_parser("outline", help="제목 구조(목차): 수준, 단락 번호, 쪽")
    common(p)
    p.add_argument("--start", type=int, default=1, help="이 번호의 제목부터 보기")

    p = sub.add_parser("comment", help="검토 메모 목록 보기(list) / 메모 달기(add)")
    common(p, write=True)
    p.add_argument("--action", choices=["list", "add", "reply", "delete"], default="list",
                   help="list(기본): 메모 목록, add: 메모 달기, reply: --no 메모에 답글, delete: --no 메모 지우기(답글 포함)")
    p.add_argument("--no", type=int, help="reply/delete: 메모 번호 (목록의 no)")
    p.add_argument("--anchor", help="add: 메모를 달 글 (문서에 있는 그대로)")
    p.add_argument("--text", help="add: 메모 내용. 줄바꿈은 \\n")
    p.add_argument("--file", help="add: 메모 내용 파일(.txt/.md) 또는 여러 메모 CSV (열: 기준 글, 메모 내용[, 몇 번째 단락])")
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
    p.add_argument("--file", help="넣을 파일 (.md/.txt는 글, .csv는 표)")
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

    p = sub.add_parser("delete", help="단락 지우기 (--paragraphs: read의 단락 번호, --dry-run으로 먼저 확인)")
    common(p, write=True)
    p.add_argument("--paragraphs", help="지울 단락 번호 (예: 12, 12-14, 3,7). word.py read의 no")

    p = sub.add_parser("format", help="단락 서식: 정렬·글꼴·크기·굵게·기울임·스타일 (--paragraphs)")
    common(p, write=True)
    p.add_argument("--paragraphs", help="바꿀 단락 번호 (예: 3, 3-5, 1,4). word.py read의 no")
    p.add_argument("--align", choices=list(WD_ALIGN), help="정렬")
    p.add_argument("--font", help="글꼴 이름 (예: 맑은 고딕)")
    p.add_argument("--size", type=float, help="글자 크기 (pt)")
    p.add_argument("--bold", action="store_true", help="굵게")
    p.add_argument("--italic", action="store_true", help="기울임")
    p.add_argument("--style", help="Word 스타일 이름 (예: 표준, 제목 1)")

    p = sub.add_parser("table", help="워드 표의 행 넣기·지우기")
    common(p, write=True)
    p.add_argument("--table", type=int, default=1, help="표 번호 (word.py info의 tables no, 기본 1)")
    p.add_argument("--action", required=True, choices=["add-row", "delete-row"])
    p.add_argument("--row", type=int, help="add-row: 이 행 뒤에 (생략 시 끝) / delete-row: 지울 행 (1 = 첫 행)")
    p.add_argument("--values", help="add-row: 새 행의 칸 글 '칸; 칸; 칸'")

    p = sub.add_parser("fill", help="서식 문서의 자리({{이름}})와 표 칸(표2!R3C4)을 값 파일로 한 번에 채우기")
    common(p, write=True)
    p.add_argument("--slots", help="채울 값 '자리=값; 자리=값' (예: \"업체명=한빛전자; 시험일=2026-11-04; 표2!R3C4=25\")")
    p.add_argument("--file", help="값 파일 (.csv/.md 표, 열: 자리, 값. 예: {{업체명}},한빛전자 / 표2!R3C4,25). 많을 때")
    p.add_argument("--overwrite", action="store_true", help="값이 있는 표 칸도 덮어쓰기 (사용자가 확인했을 때만)")

    p = sub.add_parser("header", help="머리글·바닥글 읽기 (--text가 있으면 --kind의 머리글 또는 바닥글 쓰기)")
    common(p, write=True)
    p.add_argument("--kind", choices=["header", "footer"], help="header(머리글) 또는 footer(바닥글). 읽기는 생략하면 둘 다")
    p.add_argument("--section", type=int, help="구역 번호 (읽기는 생략하면 모든 구역, 쓰기는 생략하면 1)")
    p.add_argument("--text", help="쓸 글 (기본 머리글·바닥글을 이 글로 바꿈). 줄바꿈은 \\n")
    p.add_argument("--overwrite", action="store_true", help="글이 있는 머리글·바닥글 바꾸기 (사용자가 확인했을 때만)")

    p = sub.add_parser("link", help="하이퍼링크 목록 (--anchor가 있으면 그 글에 --address 링크 달기)")
    common(p, write=True)
    p.add_argument("--anchor", help="링크를 달 문서의 글")
    p.add_argument("--occurrence", type=int, help="--anchor 글이 여러 곳이면 몇 번째인지 (오류의 [후보] 번호)")
    p.add_argument("--address", help="링크 주소: 파일 경로, http 주소, 문서 안 책갈피는 #이름")

    p = sub.add_parser("bookmark", help="책갈피 목록 (--name이 있으면 --anchor 글에 책갈피 만들기)")
    common(p, write=True)
    p.add_argument("--name", help="책갈피 이름 (글자로 시작, 글자·숫자·밑줄, 공백 없음)")
    p.add_argument("--anchor", help="책갈피를 둘 문서의 글")
    p.add_argument("--occurrence", type=int, help="--anchor 글이 여러 곳이면 몇 번째인지")
    p.add_argument("--overwrite", action="store_true", help="같은 이름의 책갈피를 이 글로 옮기기 (사용자가 확인했을 때만)")

    p = sub.add_parser("export", help="PDF/DOCX 사본 저장")
    common(p)
    p.add_argument("--save", help="저장 경로 (.pdf 또는 .docx, 생략 시 대상 파일 옆에 PDF, 없으면 C:\\AIKIT\\data)")
    p.add_argument("--page-number", action="store_true", help="사본 아래쪽 가운데에 쪽 번호")
    return parser


HANDLERS = {"read": cmd_read, "info": cmd_info, "write": cmd_write, "replace": cmd_replace, "export": cmd_export,
            "outline": cmd_outline, "comment": cmd_comment, "revisions": cmd_revisions, "fill": cmd_fill,
            "delete": cmd_delete, "table": cmd_table, "format": cmd_format, "header": cmd_header,
            "link": cmd_link, "bookmark": cmd_bookmark}


def main(argv: Optional[Sequence[str]] = None) -> None:
    args = build_cli().parse_args(argv)
    if args.command == "write" and not args.anchor and args.pos not in ("end", "start"):
        args.pos = "end"
    rt.run_cli(TOOL, args.command, lambda: HANDLERS[args.command](args))


if __name__ == "__main__":
    main()
