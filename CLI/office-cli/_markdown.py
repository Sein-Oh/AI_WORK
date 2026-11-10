"""office-cli 내부 모듈 (진입점 아님): 마크다운 블록 해석 (word.py / ppt.py 공통, excel.py 표 읽기).

제목·목록·표·코드 블록·굵게·구분선·맨 앞 YAML 머리말을 블록 목록으로 바꿉니다. 다른 office-cli 모듈을 import하지 않습니다.
"""
import re
from typing import Any, List, Tuple

# '__'는 CommonMark처럼 단어 안에서는 굵게 표시가 아님: 앞에 글자·숫자가 없을 때만 열고 뒤에 없을 때만 닫음
# (report__2024__final.xlsx는 그대로, __init__ 메서드는 굵은 'init'). '**'는 단어 안에서도 굵게.
_MD_BOLD = re.compile(r"\*\*(.+?)\*\*|(?<!\w)__(?=\S)(.+?)(?<=\S)__(?!\w)")
# 번호 목록 기호: 1~3자리 숫자만 ('2026. 11. 3. 회의'의 연도는 목록 번호가 아님)
MD_NUMBERED = re.compile(r"^\s*(\d{1,3})[.)]\s+")


def split_bold(text: str) -> Tuple[str, List[Tuple[int, int]]]:
    """'가 **나** 다' -> ('가 나 다', [(2, 3)]). 굵게 표시 구간(시작, 끝 오프셋)을 돌려줍니다."""
    plain: List[str] = []
    spans: List[Tuple[int, int]] = []
    pos = 0
    length = 0
    for m in _MD_BOLD.finditer(text):
        before = text[pos:m.start()]
        plain.append(before)
        length += len(before)
        inner = m.group(1) if m.group(1) is not None else m.group(2)
        spans.append((length, length + len(inner)))
        plain.append(inner)
        length += len(inner)
        pos = m.end()
    plain.append(text[pos:])
    return "".join(plain), spans


def table_cells(rows: List[List[str]]) -> Tuple[List[List[str]], List[Tuple[int, int]]]:
    """마크다운 표 칸의 ** __ 굵게 표시를 뺀 칸 글과, 칸 전체가 굵게인 칸 [(행, 열)] (0부터). 표 칸에 '**합계**'가 별표 그대로
    들어가던 것 (2026-11-09 검토 3장 끝). 일부만 굵게면 표시만 빼고 칸은 굵게 하지 않음."""
    out: List[List[str]] = []
    bold: List[Tuple[int, int]] = []
    for i, row in enumerate(rows):
        line = []
        for j, cell in enumerate(row):
            text = "" if cell is None else str(cell)
            plain, spans = split_bold(text)
            if spans and plain.strip() and sum(b - a for a, b in spans) >= len(plain.strip()):
                bold.append((i, j))
            line.append(plain)
        out.append(line)
    return out, bold


def plain_heading(text: str) -> str:
    """제목·슬라이드 제목에 넣을 글: 굵게 표시(** __)를 뺀 글 (제목은 굵기가 섞인 글을 담지 않음. word·ppt 모든 수준 같은 규칙)."""
    return split_bold(text)[0]


_TABLE_PIPE = re.compile(r"(?<!\\)\|")
_FENCE = re.compile(r"^(`{3,}(?!.*`)|~{3,})")    # 백틱 펜스의 정보 글에는 백틱이 없음 (```x```는 인라인 코드)


def parse_markdown_table(lines: List[str]) -> List[List[str]]:
    """'|'로 나눈 표 (앞에 역슬래시가 있는 '\\|'는 칸 안의 글자 '|'). 정렬 줄(머리글 다음 줄은 ':-', '-:' 등 '-' 하나부터, 그 밖은 '---' 이상)은 뺍니다."""
    rows: List[List[str]] = []
    for no, ln in enumerate(lines):
        s = ln.strip()
        if s.startswith("|"):
            s = s[1:]
        if s.endswith("|") and not s.endswith("\\|"):
            s = s[:-1]
        cells = [c.strip().replace("\\|", "|") for c in _TABLE_PIPE.split(s)]
        dash = r":?-+:?" if no == 1 else r":?-{3,}:?"   # 머리글 다음 줄은 '-' 하나부터 정렬 줄
        if cells and any(cells) and all(re.fullmatch(dash, c) for c in cells if c):
            continue
        rows.append(cells)
    width = max((len(r) for r in rows), default=0)
    return [r + [""] * (width - len(r)) for r in rows]


_YAML_LINE = re.compile(r"^(?:\s|-\s|-$|#|[^\s:#][^:]{0,60}:(?:\s|$))")


def front_matter_lines(text: str) -> int:
    """맨 앞 YAML 머리말(1번째 줄 '---', 그 뒤 'key: 값' 줄들, '---' 또는 '...' 줄로 닫힘)의 줄 수 (닫는 줄 포함). 없으면 0.
    닫는 줄이 없거나 사이에 'key:' 모양이 아닌 줄이 있으면 머리말이 아님 (첫 줄 '---'는 구분선)."""
    lines = text.lstrip("\ufeff").replace("\r\n", "\n").replace("\r", "\n").split("\n")
    if not lines or lines[0].strip() != "---":
        return 0
    keys = 0
    for i, ln in enumerate(lines[1:], 1):
        s = ln.rstrip()
        if s in ("---", "..."):
            return i + 1 if keys else 0
        if not s.strip():
            continue
        if not _YAML_LINE.match(s):
            return 0
        keys += bool(re.match(r"^[^\s:#-][^:]{0,60}:(?:\s|$)", s))
    return 0


# 줄 머리 기호를 글자 그대로 쓰는 표기 (CommonMark 역슬래시): '\# 글', '\- 글', '\| a', '1\. 글'
_ESCAPED = re.compile(r"^\\([#\-*+|>`~_])|^(\d{1,3})\\(?=[.)])")
_NEEDS_ESCAPE = re.compile(r"^(?:#{1,6}(?:\s|$)|[-*+](?:\s|$)|\||>|`{3}|~{3}|(?:-{3,}|\*{3,}|_{3,})$)")


def escape_line(text: str) -> str:
    """글 한 줄을 마크다운으로 쓸 때 줄 머리 기호가 제목·목록·표·구분선으로 읽히지 않게 (parse_markdown_blocks가 되돌림)."""
    s = text.strip()
    m = MD_NUMBERED.match(s)
    if m:
        k = m.end(1)
        return s[:k] + "\\" + s[k:]
    return "\\" + s if _NEEDS_ESCAPE.match(s) else s


def document_markdown(blocks: List[Tuple[str, Any]]) -> str:
    """('heading', (수준, 글)), ('paragraph', 글), ('table', [[칸]])을 parse_markdown_blocks가 같은 블록으로 되읽는 마크다운으로
    (word.py·ppt.py read --save -> write --file 왕복: 제목은 # 수준, 단락은 빈 줄로 나눈 한 줄, 표는 마크다운 표, 2026-11-09 검토 A12)."""
    out: List[str] = []
    for kind, value in blocks:
        if kind == "heading":
            level, text = value
            out += ["#" * max(1, min(6, int(level))) + " " + " ".join(str(text).split()), ""]
        elif kind == "table":
            rows = [[str(c).replace("|", "\\|").replace("\n", " ").strip() for c in r] for r in value]
            width = max(len(r) for r in rows)
            rows = [r + [""] * (width - len(r)) for r in rows]
            out.append("| " + " | ".join(rows[0]) + " |")
            out.append("|" + "---|" * width)
            out += ["| " + " | ".join(r) + " |" for r in rows[1:]]
            out.append("")
        else:
            text = " ".join(str(value).split())
            if text:
                out += [escape_line(text), ""]
    return "\n".join(out).rstrip("\n") + "\n"


def parse_markdown_blocks(text: str) -> List[Tuple[str, Any]]:
    """마크다운을 블록 목록으로 바꿉니다.
    ('heading', (레벨, 텍스트)), ('paragraph', 텍스트), ('bullets', [항목]), ('numbered', [항목]),
    ('table', [[셀]]), ('code', 줄), ('break', None)  # '---' 구분선(thematic break)
    ('front_matter', 줄 수): 맨 앞 YAML 머리말. 넣지 않는 블록 (쓰는 쪽이 warnings에 남김)
    """
    blocks: List[Tuple[str, Any]] = []
    para: List[str] = []
    lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    skip = front_matter_lines(text)
    if skip:
        blocks.append(("front_matter", skip))
        lines = lines[skip:]

    def flush() -> None:
        if para:
            blocks.append(("paragraph", " ".join(s.strip() for s in para)))
            para.clear()

    i = 0
    while i < len(lines):
        s = lines[i].strip()
        if not s:
            flush()
            i += 1
            continue
        esc = _ESCAPED.match(s)
        if esc:                                     # '\# 글', '1\. 글': 기호를 글자 그대로 (제목·목록·표로 읽지 않음)
            para.append((esc.group(1) or esc.group(2)) + s[esc.end():])
            i += 1
            continue
        fence = _FENCE.match(s)
        if fence:                                   # 코드 블록: 줄마다 단락 하나, 글자 그대로 (제목·목록으로 읽지 않음)
            flush()
            mark = fence.group(1)
            i += 1
            while i < len(lines):
                ln = lines[i].rstrip()
                i += 1
                if ln.strip().startswith(mark[0] * len(mark)) and not ln.strip().lstrip(mark[0]):
                    break                           # 닫는 펜스 (닫지 않으면 끝까지 코드)
                if ln.strip():
                    blocks.append(("code", ln))     # 굵게(** __) 등 마크다운 기호를 해석하지 않고 그대로
            continue
        m = re.match(r"^(#{1,6})\s+(.*)$", s)
        if m:
            flush()
            blocks.append(("heading", (len(m.group(1)), m.group(2).strip().rstrip("#").strip())))
            i += 1
            continue
        if s.startswith("|"):
            flush()
            tbl = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                tbl.append(lines[i])
                i += 1
            rows = parse_markdown_table(tbl)
            if rows:
                blocks.append(("table", rows))
            continue
        if re.match(r"^[-*+]\s+", s):
            flush()
            items = []
            while i < len(lines) and re.match(r"^\s*[-*+]\s+", lines[i]):
                items.append(re.sub(r"^\s*[-*+]\s+", "", lines[i]).strip())
                i += 1
            blocks.append(("bullets", items))
            continue
        num = MD_NUMBERED.match(s)
        if num:
            flush()
            run = []
            while i < len(lines) and MD_NUMBERED.match(lines[i]):
                run.append(lines[i].strip())
                i += 1
            first = int(num.group(1))
            if blocks and blocks[-1][0] == "numbered" and first == len(blocks[-1][1]) + 1:
                blocks[-1][1].extend(MD_NUMBERED.sub("", x).strip() for x in run)   # 빈 줄로 띄운 같은 목록의 다음 번호
            elif first == 1:
                blocks.append(("numbered", [MD_NUMBERED.sub("", x).strip() for x in run]))
            else:
                # 1이 아닌 번호로 시작하는 목록: 자동 번호는 1부터 다시 매기므로 사용자 번호를 글로 둔 단락으로 넣음
                blocks += [("paragraph", x) for x in run]
            continue
        if re.fullmatch(r"(-{3,}|\*{3,}|_{3,})", s):
            flush()
            blocks.append(("break", None))
            i += 1
            continue
        para.append(s)
        i += 1
    flush()
    return blocks


def looks_like_markdown(text: str) -> bool:
    return bool(re.search(r"(?m)^\s*(#{1,6}\s|[-*+]\s|\d{1,3}[.)]\s|\|.*\||```|~~~)", text) or "**" in text
                or front_matter_lines(text))


def front_matter_warning(lines: int) -> str:
    return f"맨 앞 YAML 머리말({lines}줄, '---'로 둘러싼 key: 값)은 문서 정보라 넣지 않았습니다."
