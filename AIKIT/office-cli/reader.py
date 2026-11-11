"""reader.py - 문서 파일 읽기, 등록(캐시), 근거 검색 (Office 프로그램 불필요)

명령:
    read    파일 하나를 읽고 캐시에 등록합니다. PDF, HWP, HWPX, DOCX, PPTX, XLSX, CSV, TXT, MD
    index   폴더(또는 파일)의 문서를 내용 출력 없이 한꺼번에 캐시에 등록합니다.
    search  캐시에 등록된 문서에서 검색합니다. 결과마다 근거(쪽, 장절, 번호)를 줍니다.
            --file을 주면 그 파일/폴더에서 아직 등록되지 않았거나 바뀐 문서를 먼저 등록한 뒤 검색합니다.
    list    캐시에 등록된 문서 목록 (--name, --folder로 거르기)
    clear   캐시 지우기 (--file, --name, --folder, --missing, --all, --dry-run). 원본 문서는 그대로

캐시 형식(모든 형식 공통): 문서마다 units 목록, 단위마다
    id, kind, loc, page, page_source, section(장절 경로), number(항목 번호), heading, text
결과를 다른 문서에 넣으려면 --save로 파일을 만든 뒤 word.py/ppt.py write --file 또는 excel.py write --file로 넘깁니다.

예:
    python reader.py index --folder "C:\\docs\\규격"
    python reader.py read --file "C:\\docs\\보고서.pdf" --pages 1-3
    python reader.py search --query "마일스톤;일정" --save "근거.xlsx"
    python reader.py search --query "단가" --file "C:\\docs\\견적" --pages 3-7 --section "SCOPE"
    python reader.py list --name 규격
"""
import bisect
import contextlib
import functools
import hashlib
import json
import os
import re
import sys
import time
import unicodedata
import zipfile
import zlib
from collections import Counter
from typing import Any, Dict, List, Optional, Tuple
from xml.etree import ElementTree as ET

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _runtime as rt  # noqa: E402
import _xlsx_formula as xf  # noqa: E402

TOOL = "reader"
SUPPORTED = {
    ".pdf": "pdf", ".hwp": "hwp", ".hwpx": "hwpx",
    ".docx": "docx", ".docm": "docx", ".pptx": "pptx", ".pptm": "pptx",
    ".xlsx": "xlsx", ".xlsm": "xlsx", ".csv": "csv", ".txt": "txt", ".md": "txt", ".eml": "eml",
}
UNIT_NAMES = {"pdf": "단락", "pptx": "슬라이드", "xlsx": "행", "csv": "행", "eml": "항목"}
LEGACY_TYPES = {".doc", ".xls", ".ppt", ".rtf", ".xlsb", ".odt", ".ods", ".odp"}   # 읽지 않는 문서 형식 (폴더 검색이면 not_searched)
PAGE_TYPES = ("pdf", "pptx")   # read --pages가 실제 쪽/슬라이드 번호인 형식 (그 밖은 항목 번호)
PARSER_VERSION = 12         # 파서가 바뀌면 올려서 기존 캐시를 무효화 (2: 공통 단위 형식, 장절/번호/쪽, 3: PDF 머리글·2단, DOCX 변경 추적·표 행 쪽, 4: 나열 항목은 제목 아님, 5: PDF 글자 없는 쪽 meta, 6: 글자 없는 쪽은 그림이 있는 쪽만, 7: 머리글을 뺀 뒤 글이 없고 그림이 쪽을 크게 덮는 쪽만, 8: 짧은 글(도장·Bates 번호)만 남은 스캔 쪽도, 보이지 않는 OCR 글자 층이 있는 쪽은 제외, 9: DOCX 글상자 글, XLSX 계산된 값이 없는 수식 칸 meta, 10: fullCalcOnLoad 파일의 수식 칸·확인 못함도 meta, 11: 표 머리글 행도 단위·2열 표·이어진 PDF 표, PDF 줄바꿈 제목, DOCX 머리글·바닥글·각주와 PDF 반복 머리글을 단위로, 12: .eml)
SCAN_IMAGE_COVER = 0.4      # 스캔 쪽으로 보는 그림 면적 비율 (쪽 면적 대비)
SCAN_TEXT_MAX = 40          # 그림이 덮는 쪽에 이보다 적은 글자(도장, Bates 번호)만 남으면 스캔 쪽
MAX_SHEET_ROWS = 20000      # xlsx 시트당 최대 행 수
MAX_FOLDER_FILES = 300      # 한 번에 새로 등록(파싱)할 최대 파일 수
MAX_FOLDER_DEPTH = 3        # 폴더 기본 탐색 깊이
MAX_FOLDER_SCAN = 20000     # 폴더에서 살펴볼 최대 파일 수
MAX_LIST_ITEMS = 20         # JSON에 담는 실패/새 등록 목록 상한
LIST_SHOW = 40              # list 명령이 JSON에 담는 문서 수 (전에는 100개, 넘으면 전체는 full_result_file)
MAX_SAVE_HITS = 5000        # --save 파일에 담는 최대 검색 결과 수
MAX_SEARCH_TOP = 50         # search --top 상한 (더 많은 결과는 --save 파일로)
MAX_SEARCH_CONTEXT = 500    # search --context(스니펫 앞뒤 글자 수) 상한
SEARCH_JSON_FULL_CHARS = 2000   # search --full의 JSON 항목 text 상한 (전체는 --save)
SAVE_CONTEXT_CHARS = 3000   # search --save의 '문맥' 칸 상한 (--full이면 단위 전체)
STALE_REFRESH_MAX = 20      # 검색 때 원본이 바뀐 캐시를 그 자리에서 다시 읽는 최대 문서 수
HEADING_MAX_CHARS = 80      # 번호로 시작하는 짧은 줄만 제목 후보 (굵게/큰 글씨는 120)
LEGACY_OFFICE = (".xls", ".doc", ".ppt")
INDEX_LOCK_WAIT = 15.0      # 색인 잠금(index.lock)을 기다리는 최대 초 (동시 실행). 잠근 쪽의 retry_fs 재시도(최대 약 10.5초)보다 길게
INDEX_LOCK_STALE = 60.0     # 이보다 오래된 잠금 파일은 중단된 실행이 남긴 것으로 보고 지움


# ==============================================================================
# 1. 번호와 장절 인식 (모든 형식 공통)
# ==============================================================================
# 번호 토큰을 "계열"로 나누고, 계열이 문서에서 처음 나온 순서로 수준을 정합니다.
#   - 처음 보는 계열은 현재 가장 깊은 제목의 한 단계 아래
#   - 이미 본 계열은 그 계열의 수준으로 돌아감 (그보다 깊은 제목은 닫힘)
#   - 1.1.1 같은 점 번호는 숫자 개수가 깊이 (끝의 .0은 빼고 셈: 1.0# -> 깊이 1)
# 글머리 기호(O, -, 동그라미/네모 기호 등)는 번호가 아니고 제목도 아닙니다.

GANADA = "가나다라마바사아자차카타파하"
CHOSUNG = "ㄱㄴㄷㄹㅁㅂㅅㅇㅈㅊㅋㅌㅍㅎ"
_ROMAN_FULLWIDTH = {chr(0x2160 + i): r for i, r in enumerate(
    ["I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X", "XI", "XII"])}
# 글머리 기호: 영문 O/o, -, *, 가운뎃점, 그리고 유니코드 기호들(코드로만 적음)
_BULLET_SPACE = "Oo-*"
_BULLET_SYMBOLS = ("\u00b7\u2022\u2023\u2043\u2219\u25cb\u25cf\u25ce\u25e6\u25a1\u25a0\u25aa\u25ab\u25c7\u25c6"
                   "\u25b6\u25b7\u25ba\u25b8\u25b9\u203b\u3147\u2218\u27a2\u2713\u2714\u221a\u2192\u21d2\u2794"
                   "\uf0a7\uf0b7\uf06c\uf0d8\uf0fc\uf076")

_GA = f"[{GANADA}]"
_NUM_PATTERNS = [
    # (계열, 정규식, 번호 표기 함수)
    ("paren_arabic", re.compile(r"\(\s*(\d{1,3})\s*\)"), lambda m: f"({m.group(1)})"),
    ("paren_ganada", re.compile(rf"\(\s*({_GA})\s*\)"), lambda m: f"({m.group(1)})"),
    ("paren_alpha", re.compile(r"\(\s*([a-zA-Z])\s*\)"), lambda m: f"({m.group(1)})"),
    ("rparen_arabic", re.compile(r"(\d{1,3})\)"), lambda m: f"{m.group(1)})"),
    ("rparen_ganada", re.compile(rf"({_GA})\)"), lambda m: f"{m.group(1)})"),
    ("rparen_alpha", re.compile(r"([a-z])\)"), lambda m: f"{m.group(1)})"),
]
_KO_UNIT = {"편": "ko_pyeon", "장": "ko_jang", "절": "ko_jeol", "관": "ko_gwan", "조": "ko_jo", "항": "ko_hang",
            "호": "ko_ho"}
_KO_RE = re.compile(r"제\s*(\d{1,4})\s*(편|장|절|관|조|항|호)(?:\s*의\s*(\d{1,3}))?(?=[\s(.:]|$)")
_WORD_RE = re.compile(r"(?i)(chapter|part|section|article)\s+(\d{1,3}|[IVX]{1,6})(?=[\s.:]|$)[.:]?")
_ANNEX_RE = re.compile(r"(?i)(annex|appendix|attachment)\s+([A-Z]|\d{1,3})(?=[\s.:)]|$)[.:]?")
_ANNEX_KO_RE = re.compile(r"(부록|별표|별지|첨부)\s*(\d{1,3}|[A-Z])(?=[\s.:)]|$)[.:]?|(부록)(?=[\s.:]|$)")
_Q_RE = re.compile(r"(Q|문항|질문|문)\s*\.?\s*(\d{1,3})(?=[\s.):]|$)[.)]?")
_ROMAN_RE = re.compile(r"([IVX]{1,6})[.)](?=\s|$)")
_ROMAN_LOWER_RE = re.compile(r"([ivx]{1,6})[.)](?=\s|$)")
_ARABIC_RE = re.compile(r"(\d{1,2}(?:\.\d{1,3}){0,6})(\.|#)?(?=\s|$)")
_GANADA_RE = re.compile(rf"({_GA})\.(?=\s|$)")
_ALPHA_UPPER_RE = re.compile(r"([A-Z])\.(?=\s)")
_ALPHA_LOWER_RE = re.compile(r"([a-z])\.(?=\s)")
_SENTENCE_END = re.compile(r"(?:[.\u3002!;]|다|니다|어요|아요|해요|세요)\s*$")
_NOUN_LIKE_END = re.compile(r"(?:음|함|임|됨)\s*$")
_LETTER = re.compile(r"[A-Za-z\uac00-\ud7a3]")
# 번호 뒤가 단위면 번호가 아니라 수치 (10.5 kV 이하, 1.5 kV 1분, 3 mm, 2 개월). '1. A형'처럼 뒤에 글자가 붙으면 단위 아님
_UNIT_AFTER = re.compile(r"\s*(?:%|\u2103|\u00b0|\u2126|\u03a9|"
                         r"[kKMGmunp\u03bc]?(?:V|A|W|Hz|Pa|bar|g|m|L|s|dB|lx|ppm|ohm|F|N|J|"
                         r"V(?:DC|AC|dc|ac|rms|p-p|pp)|A(?:DC|AC|rms)|N-?m|N\u00b7m)(?=[\s/,.)\]]|$)|"
                         r"(?:times|x|X|\u00d7)(?=\s|$)|"
                         r"(?:초|분|시간|일|주|개월|년|회|개|배|도|kg|mm|cm|km|mA|kA)(?=[\s/,.)\]]|$))")
# 목차 줄: 점선 리더나 탭 뒤에 쪽 번호 (1.2 설계 기준 ........ 12)
# 문장 중간에서 끝나는 나열 항목: "(3) Verification Review; and", "가) 시험 장비,", "1) 설계 검토 및"
_LIST_CONTINUES = re.compile(r"(?:[;,]|[;,]?\s+(?:and|or|및|또는|그리고))\s*$", re.IGNORECASE)
_TOC_LINE = re.compile(r"(?:[.\u00b7\u2026\u2025\u30fb\u2219_-]{3,}|\t)\s*\d{1,4}\s*$")


def _info(family: str, number: Optional[str], rest: str, depth: int = 1, plain: bool = False) -> Dict[str, Any]:
    return {"family": family, "number": number, "rest": rest.strip(), "depth": depth, "plain": plain}


def classify_number(line: str) -> Optional[Dict[str, Any]]:
    """줄 머리의 번호 토큰을 찾습니다. 반환: {family, number, rest, depth, plain} 또는 None.
    family 'bullet'은 글머리 기호(번호 아님). plain=True는 구분 기호 없는 숫자 하나('1 Scope')로, 굵게/큰 글씨일 때만 제목."""
    s = (line or "").strip()
    if not s:
        return None
    c, o = s[0], ord(s[0])
    if 0x2460 <= o <= 0x2473:                        # 동그라미 숫자 1~20
        return _info("circled", f"({o - 0x245F})", s[1:])
    if 0x2474 <= o <= 0x2487:                        # 괄호 숫자 1~20
        return _info("paren_arabic", f"({o - 0x2473})", s[1:])
    if 0x326E <= o <= 0x327B:                        # 동그라미 가나다
        return _info("circled_ganada", f"({GANADA[o - 0x326E]})", s[1:])
    if 0x3260 <= o <= 0x326D:                        # 동그라미 자음
        return _info("circled_jamo", f"({CHOSUNG[o - 0x3260]})", s[1:])
    if c in _ROMAN_FULLWIDTH and (len(s) == 1 or s[1] in " .)\t"):
        rest = s[1:].lstrip(".)")
        return _info("roman", _ROMAN_FULLWIDTH[c], rest)
    if c in _BULLET_SYMBOLS or (c in _BULLET_SPACE and (len(s) == 1 or s[1] in " \t")):
        return _info("bullet", None, s[1:])
    for family, pat, fmt in _NUM_PATTERNS:
        m = pat.match(s)
        if m:
            return _info(family, fmt(m), s[m.end():])
    m = _KO_RE.match(s)
    if m:
        num = f"제{m.group(1)}{m.group(2)}" + (f"의{m.group(3)}" if m.group(3) else "")
        return _info(_KO_UNIT[m.group(2)], num, s[m.end():])
    m = _WORD_RE.match(s)
    if m:
        return _info(m.group(1).lower(), f"{m.group(1)} {m.group(2)}", s[m.end():])
    m = _ANNEX_RE.match(s)
    if m:
        return _info("annex", f"{m.group(1)} {m.group(2)}", s[m.end():])
    m = _ANNEX_KO_RE.match(s)
    if m:
        num = f"{m.group(1)} {m.group(2)}" if m.group(1) else m.group(3)
        return _info("annex_ko", num, s[m.end():])
    m = _Q_RE.match(s)
    if m:
        return _info("question", f"{m.group(1)}{m.group(2)}", s[m.end():])
    m = _ROMAN_RE.match(s)
    if m:
        return _info("roman", m.group(1), s[m.end():])
    m = _ROMAN_LOWER_RE.match(s)
    if m:
        return _info("roman_lower", m.group(1), s[m.end():])
    m = _ARABIC_RE.match(s)
    if m:
        rest = s[m.end():]
        if rest.strip()[:1].isdigit() or _UNIT_AFTER.match(rest):   # 날짜/수치 (예: 10. 28., 10.5 kV)
            return None
        parts = m.group(1).split(".")
        depth_parts = list(parts)
        while len(depth_parts) > 1 and depth_parts[-1] == "0":   # 1.0# -> 깊이 1
            depth_parts.pop()
        plain = len(parts) == 1 and not m.group(2)
        return _info("arabic", m.group(1), rest, depth=len(depth_parts), plain=plain)
    m = _GANADA_RE.match(s)
    if m:
        return _info("ganada", m.group(1), s[m.end():])
    m = _ALPHA_UPPER_RE.match(s)
    if m:
        return _info("alpha_upper", m.group(1), s[m.end():])
    m = _ALPHA_LOWER_RE.match(s)
    if m:
        return _info("alpha_lower", m.group(1), s[m.end():])
    return None


def heading_title(info: Optional[Dict[str, Any]], line: str, strong: bool = False) -> Optional[str]:
    """번호로 시작하는 줄이 제목이면 장절 경로에 넣을 제목 글을, 아니면 None을 돌려줍니다.
    제목 조건: 짧은 한 줄, 문장 끝(. / 다 / 요 등)으로 끝나지 않음, 번호 뒤에 글자가 있음."""
    if info is None or info["family"] == "bullet":
        return None
    s = line.strip()
    if "\n" in s or _TOC_LINE.search(s):            # 목차 줄은 제목이 아님 (목차 쪽이 장절로 남지 않게)
        return None
    rest = info["rest"]
    if info["family"].startswith("ko_") and info["family"] in ("ko_pyeon", "ko_jang", "ko_jeol", "ko_gwan", "ko_jo"):
        # 제N조(목적) 본문... 처럼 조문은 문장이어도 구조이므로 번호와 괄호 제목만 장절로 씁니다.
        m = re.match(r"\s*(\([^)]{1,40}\))", rest)
        if len(s) > HEADING_MAX_CHARS or _SENTENCE_END.search(s):
            return info["number"] + (m.group(1) if m else "")
        return s if _LETTER.search(rest) or m else info["number"]
    if info["plain"] and not strong:
        return None
    if not _LETTER.search(rest):
        return None
    if len(s) > (120 if strong else HEADING_MAX_CHARS):
        return None
    if _SENTENCE_END.search(s) or _LIST_CONTINUES.search(s):
        return None
    if _NOUN_LIKE_END.search(s) and len(rest) >= 8:   # 개조식 문장 끝 (~함, ~임). 짧은 명사(소음)는 제외
        return None
    return s


class Outline:
    """장절 경로 스택. 계열별 수준은 문서에서 처음 나온 위치로 정합니다."""

    def __init__(self) -> None:
        self.stack: List[Dict[str, Any]] = []
        self.levels: Dict[str, int] = {}

    def path(self) -> List[str]:
        return [e["title"] for e in self.stack]

    def _place(self, level: int, title: str, family: str, size: Optional[float]) -> None:
        level = max(1, level)
        while self.stack and self.stack[-1]["level"] >= level:
            self.stack.pop()
        self.stack.append({"level": level, "title": title, "family": family, "size": size})

    def register(self, family: str, depth: int, level: int) -> None:
        """스타일(제목 N, 개요 N)로 수준이 정해진 제목의 번호 계열을 기록 (같은 계열의 일반 단락이 맞는 수준을 쓰게)."""
        if family not in self.levels:
            self.levels[family] = max(1, level - (depth - 1))

    def push_level(self, level: int, title: str, size: Optional[float] = None) -> None:
        self._place(level, title, "style", size)

    def push_family(self, family: str, depth: int, title: str, size: Optional[float] = None) -> None:
        if family in self.levels:
            level = self.levels[family] + depth - 1
        else:
            base = (self.stack[-1]["level"] if self.stack else 0) + 1
            self.levels[family] = base
            level = base + depth - 1
        # 글자 크기 확인: 상위 제목보다 분명히 큰 제목은 그 상위의 하위가 아님 -> 더 큰 제목 아래까지 올라감
        if size:
            parents = [e for e in self.stack if e["level"] < level]
            if parents and parents[-1]["size"] and size > parents[-1]["size"] + 0.5:
                keep = [e for e in parents if e["size"] and e["size"] > size + 0.5]
                new_level = (keep[-1]["level"] if keep else 0) + 1
                self.levels[family] = max(1, new_level - (depth - 1))
                level = new_level
        self._place(level, title, family, size)


class UnitBuilder:
    """모든 형식이 같은 키의 단위를 만들게 합니다."""

    def __init__(self) -> None:
        self.units: List[Dict[str, Any]] = []
        self.outline = Outline()

    def add(self, kind: str, loc: str, text: str, page: Optional[int] = None, page_source: Optional[str] = None,
            number: Optional[str] = None, section: Optional[List[str]] = None) -> Dict[str, Any]:
        sec = list(section) if section is not None else self.outline.path()
        unit = {"id": len(self.units) + 1, "kind": kind, "loc": loc, "page": page,
                "page_source": page_source if page is not None else None,
                "section": sec, "number": number, "heading": sec[-1] if sec else "", "text": text}
        self.units.append(unit)
        return unit

    def add_text(self, loc: str, text: str, page: Optional[int] = None, page_source: Optional[str] = None, *,
                 kind: str = "paragraph", strong: bool = False, size: Optional[float] = None,
                 style_level: Optional[int] = None, list_family: Optional[str] = None,
                 list_depth: int = 1, allow_heading: bool = True) -> Dict[str, Any]:
        """글 단위 하나를 넣으며 제목이면 장절 경로를 갱신합니다."""
        s = text.strip()
        first = s.split("\n", 1)[0]
        info = classify_number(first)
        number = info["number"] if info and info["family"] != "bullet" and not info["plain"] else None
        if style_level:
            if info and info["family"] != "bullet":
                number = info["number"]
                self.outline.register(list_family or info["family"], list_depth if list_family else info["depth"],
                                      style_level)
            self.outline.push_level(style_level, first[:120], size)
            return self.add("heading", loc, s, page, page_source, number)
        title = heading_title(info, s, strong) if allow_heading else None
        if title:
            assert info is not None
            fam, depth = (list_family, list_depth) if list_family else (info["family"], info["depth"])
            self.outline.push_family(fam, depth, title[:120], size)
            return self.add("heading", loc, s, page, page_source, info["number"])
        return self.add(kind, loc, s, page, page_source, number)


_ROW_ID_RE = re.compile(r"^(?:\d{1,3}(?:\.\d{1,3})*\.?|[A-Za-z]{1,8}[-_]?\d[\w.-]{0,12})$")


def _clean_cell(v: Any) -> str:
    return " ".join(str(v or "").split())


def _looks_like_header(row: List[str]) -> bool:
    cells = [c for c in row if c]
    if not cells or len(cells) < max(1, len([1 for c in row]) // 2):
        return False
    return all(len(c) <= 30 and not re.fullmatch(r"[\d.,%\s-]+", c) for c in cells)


def table_header(grid: List[List[str]]) -> Optional[List[str]]:
    """정리한 표(빈 행 없음)의 첫 행이 머리글이면 머리글 (빈 칸은 앞 칸을 이어 씀), 아니면 None.
    2열 표(항목 | 값 꼴)는 첫 행도 항목이라 머리글로 보지 않음 (2026-11-09 검토 A5: '모델명: 입력 전압 | PSU-100: 220 V')."""
    if len(grid) < 2 or max(len(r) for r in grid) <= 2 or not _looks_like_header(grid[0]):
        return None
    header, prev = [], ""
    for c in grid[0]:
        prev = c or prev
        header.append(prev)
    return header


def clean_table(rows: List[List[Any]]) -> List[List[str]]:
    """표 칸 글을 정리하고 빈 행을 뺀 것 (table_rows와 같은 정리)."""
    return [r for r in ([_clean_cell(c) for c in row] for row in rows) if any(r)]


def table_rows(rows: List[List[Any]], keys: Optional[List[Any]] = None,
               carry: Optional[List[str]] = None) -> List[Dict[str, Any]]:
    """표를 행 단위 글로 바꿉니다. 첫 행이 머리글로 보이면 '머리글: 값 | 머리글: 값' 형식.
    병합/빈 머리글 칸은 앞 칸의 머리글을 이어 쓰고, 여러 줄 칸은 공백으로 잇습니다.
    keys를 주면 행마다 같은 순서의 값을 결과의 key로 돌려줍니다 (예: 행의 쪽 번호).
    carry: 앞 쪽에서 이어진 표의 머리글 (첫 행이 그 머리글을 되풀이하면 머리글 행, 아니면 첫 행부터 데이터).
    반환: [{"row": 행 번호(1부터), "text", "number", "key"}]. 머리글 행도 머리글 칸 글을 ' | '로 이은 단위로 넣음 (검색에서 빠지지 않게)."""
    cleaned = [([_clean_cell(c) for c in r], (keys[i] if keys is not None and i < len(keys) else None))
               for i, r in enumerate(rows)]
    cleaned = [(r, k) for r, k in cleaned if any(r)]
    grid = [r for r, _ in cleaned]
    row_keys = [k for _, k in cleaned]
    if not grid:
        return []
    header: Optional[List[str]] = None
    header_row = False
    if carry is not None and len(carry) == max(len(r) for r in grid):
        header = carry
        header_row = [c for c in grid[0] if c] == [c for i, c in enumerate(carry) if c and (i == 0 or c != carry[i - 1])]
    else:
        header = table_header(grid)
        header_row = header is not None
    out: List[Dict[str, Any]] = []
    for i, r in enumerate(grid, 1):
        if header_row and i == 1:
            out.append({"row": 1, "text": " | ".join(c for c in r if c), "number": None, "key": row_keys[0]})
            continue
        parts: List[str] = []
        for j, v in enumerate(r):
            if not v or (j > 0 and v == r[j - 1] and header is not None and header[j] == header[j - 1]):
                continue                                     # 빈 칸, 가로 병합으로 반복된 칸
            h = header[j] if header is not None and j < len(header) else ""
            parts.append(f"{h}: {v}" if h and h != v else v)
        if not parts:
            continue
        first = next((c for c in r if c), "")
        number = None
        if len(first) <= 20 and _ROW_ID_RE.match(first):
            number = first.rstrip(".")
        else:
            info = classify_number(first)
            if info and info["family"] != "bullet" and not info["plain"]:
                number = info["number"]
        out.append({"row": i, "text": " | ".join(parts), "number": number, "key": row_keys[i - 1]})
    return out


# ==============================================================================
# 2. 형식별 파서: 결과는 (units, meta). units는 UnitBuilder 형식
# ==============================================================================

def _encrypted_hint(path: str) -> rt.InputError:
    return rt.InputError(
        f"파일을 열 수 없습니다. 문서 보안(DRM)으로 암호화되었거나 손상된 파일일 수 있습니다: {os.path.basename(path)}. "
        "Office 문서라면 excel.py / word.py / ppt.py read --target \"<경로>\"로 Office에서 직접 읽어 보세요."
    )


class _DamagedFile(Exception):
    """파서가 만난 파일 구조 손상 (load_document_status가 InputError로 바꿈)."""


# 파일 내용이 깨졌을 때 파서가 내는 예외: 내부 오류(ERROR:)가 아니라 입력 문제(STOP_RETRY)로 알림
_DAMAGE_ERRORS: Tuple[type, ...] = (_DamagedFile, ET.ParseError, zlib.error, zipfile.BadZipFile, EOFError)
_DAMAGE_NAMES = {"XMLSyntaxError", "PackageNotFoundError"}       # lxml / python-docx·pptx (import 없이 이름으로)


def _damaged_error(path: str, e: BaseException) -> rt.InputError:
    return rt.InputError(
        f"손상된 파일이라 읽을 수 없습니다: {os.path.basename(path)} ({type(e).__name__}: {str(e)[:120]}). "
        "같은 명령을 다시 실행하지 말고 사용자에게 알리세요. 원본을 다시 받거나 프로그램(한글/Office)에서 열어 다른 이름으로 저장하면 읽을 수 있습니다."
    )


def _pdf_tables(page) -> List[Dict[str, Any]]:
    """선으로 그린 표를 찾습니다 (pymupdf find_tables). 실패하거나 없으면 빈 목록 (단락으로만 처리)."""
    try:
        found = page.find_tables()
        tables = []
        for t in getattr(found, "tables", []) or []:
            rows = t.extract() or []
            if len(rows) >= 2 and max((len(r) for r in rows), default=0) >= 2:
                tables.append({"bbox": tuple(t.bbox), "rows": rows})
        return tables
    except Exception as e:
        rt.debug(f"PDF 표 인식 실패 (단락으로 처리): {e}")
        return []


def _inside(bbox: Tuple[float, ...], box: Tuple[float, ...]) -> bool:
    cx, cy = (bbox[0] + bbox[2]) / 2, (bbox[1] + bbox[3]) / 2
    return box[0] - 1 <= cx <= box[2] + 1 and box[1] - 1 <= cy <= box[3] + 1


def _pdf_items(page, tables: Optional[List[Dict[str, Any]]] = None) -> List[Tuple[Tuple[float, ...], Any]]:
    """페이지의 (bbox, 내용) 목록. 표 영역 안의 글 블록은 표 행과 겹치므로 뺍니다."""
    tables = tables or []
    items: List[Tuple[Tuple[float, ...], Any]] = [(tuple(t["bbox"]), {"table": t["rows"]}) for t in tables]
    for b in page.get_text("dict").get("blocks", []):
        if b.get("type", 0) != 0:
            continue
        bbox = tuple(b.get("bbox", (0, 0, 0, 0)))
        if any(_inside(bbox, t["bbox"]) for t in tables):
            continue
        lines: List[Tuple[str, float, bool]] = []
        for ln in b.get("lines", []):
            spans = [sp for sp in ln.get("spans", []) if sp.get("text", "").strip()]
            if not spans:
                continue
            text = "".join(sp.get("text", "") for sp in ln.get("spans", [])).strip()
            size = max(float(sp.get("size", 0) or 0) for sp in spans)
            bold = all((int(sp.get("flags", 0)) & 16) or "bold" in str(sp.get("font", "")).lower() for sp in spans)
            lines.append((text, size, bool(bold)))
        if lines:
            items.append((bbox, lines))
    return items


def _order_items(items: List[Tuple[Tuple[float, ...], Any]], width: float) -> List[Any]:
    """읽는 순서. 기본은 위에서 아래. 2단 편집(왼쪽 단과 오른쪽 단이 세로로 겹침)이면 가로로 걸친 블록(제목, 표) 사이마다
    왼쪽 단을 먼저, 오른쪽 단을 다음에 읽습니다 (y로만 정렬하면 두 단의 줄이 섞여 장절이 틀어짐)."""
    def by_y(xs):
        return sorted(xs, key=lambda it: (round(it[0][1], 1), it[0][0]))
    if not items or width <= 0:
        return [it[1] for it in by_y(items)]
    mid, gap = width / 2, width * 0.04
    full = [it for it in items if it[0][0] < mid - gap and it[0][2] > mid + gap]
    left = [it for it in items if it[0][2] <= mid + gap and it not in full]
    right = [it for it in items if it[0][0] >= mid - gap and it not in full and it not in left]
    overlap = any(lb[0][1] < rb[0][3] and rb[0][1] < lb[0][3] for lb in left for rb in right)
    if len(left) < 2 or len(right) < 2 or not overlap:
        return [it[1] for it in by_y(items)]
    out: List[Any] = []
    rest = left + right
    for f in by_y(full):
        band = [it for it in rest if it[0][1] < f[0][1]]
        rest = [it for it in rest if it not in band]
        out += [it[1] for it in by_y([b for b in band if b in left]) + by_y([b for b in band if b in right])]
        out.append(f[1])
    out += [it[1] for it in by_y([b for b in rest if b in left]) + by_y([b for b in rest if b in right])]
    return out


_MARGIN = 0.08          # 쪽 위/아래 8% 안에 있는 줄만 머리글/바닥글 후보


def _margin_key(item: Tuple[Tuple[float, ...], Any], height: float) -> Optional[str]:
    """머리글/바닥글 후보 블록의 비교용 글 (숫자는 #으로: '- 3 -'과 '- 4 -'는 같은 쪽 번호 줄)."""
    bbox, payload = item
    if isinstance(payload, dict) or height <= 0:
        return None
    if not (bbox[3] <= height * _MARGIN or bbox[1] >= height * (1 - _MARGIN)):
        return None
    text = " ".join(t for t, _, _ in payload)
    return re.sub(r"\d+", "#", re.sub(r"\s+", " ", text)).strip() or None


def _drop_headers_footers(pages: List[Tuple[List[Any], float]], dropped: Optional[List[Tuple[int, str]]] = None) -> int:
    """여러 쪽의 위/아래 여백에 반복되는 줄(문서 제목 머리글, 쪽 번호)을 뺍니다. 뺀 블록 수를 돌려줍니다.
    남겨 두면 매 쪽 장절 경로가 머리글로 바뀌고, 머리글 문구가 쪽마다 검색 결과에 걸립니다.
    dropped: 뺀 블록의 (쪽, 글)을 담음 (문서 번호·개정 표기가 검색되게 따로 단위로 넣기 위해, 2026-11-09 검토 A13)."""
    if len(pages) < 2:
        return 0
    seen: Counter = Counter()
    for items, height in pages:
        seen.update({k for k in (_margin_key(it, height) for it in items) if k})
    repeated = {k for k, n in seen.items() if n >= max(2, (len(pages) + 1) // 2)}
    if not repeated:
        return 0
    removed = 0
    for pno, (items, height) in enumerate(pages, 1):
        keep = [it for it in items if _margin_key(it, height) not in repeated]
        if dropped is not None:
            dropped += [(pno, " ".join(t for t, _, _ in it[1])) for it in items
                        if _margin_key(it, height) in repeated]
        removed += len(items) - len(keep)
        items[:] = keep
    return removed


class _CurrentStderr:
    """쓸 때마다 그때의 sys.stderr로 보내는 출력 대상. pymupdf는 받은 스트림 객체를 기억하므로, 그 순간의 sys.stderr를 주면
    같은 프로세스에서 stderr가 바뀌고 닫힌 뒤(시험의 출력 캡처 등) 메시지를 쓰다 'I/O operation on closed file'로 실패함
    (집 PC 저장소 전체 pytest, 1.4.90)."""

    def write(self, text: str) -> int:
        return sys.stderr.write(text)

    def flush(self) -> None:
        sys.stderr.flush()


def _import_pymupdf():
    """pymupdf를 불러오고, 라이브러리 안내/경고가 stdout(1줄 JSON 자리)에 나오지 않게 stderr로 돌립니다."""
    pymupdf = rt.require_lib("pymupdf", "pymupdf", "PDF 읽기")
    for fn, kwargs in (("set_messages", {"stream": _CurrentStderr()}), ("set_log", {"stream": _CurrentStderr()}),
                       ("no_recommend_layout", {})):
        try:
            getattr(pymupdf, fn)(**kwargs)
        except Exception:
            pass
    return pymupdf


def parse_pdf(path: str) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    import contextlib
    with contextlib.redirect_stdout(sys.stderr):    # pymupdf의 print 안내도 stderr로
        return _parse_pdf(path)


def _parse_pdf(path: str) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    pymupdf = _import_pymupdf()
    try:
        doc = pymupdf.open(path)
    except Exception:
        raise _encrypted_hint(path)
    try:
        if doc.needs_pass:
            raise rt.InputError(f"암호가 걸린 PDF입니다: {os.path.basename(path)}")
        raw = [(_pdf_items(page, _pdf_tables(page)), float(page.rect.height), float(page.rect.width)) for page in doc]
        # 그림이 쪽 면적의 큰 부분을 덮는지는 문서를 닫기 전에 셈 (작은 로고·도장은 스캔 쪽이 아님)
        image_cover = [_page_image_cover(page) for page in doc]
        # 그림을 크게 덮은 쪽에 보이지 않는 글자 층(OCR 결과)이 있으면 이미 OCR한 쪽
        ocr_layer = [cover >= SCAN_IMAGE_COVER and _page_has_ocr_layer(page) for cover, page in zip(image_cover, doc)]
        meta = {"pages": doc.page_count, "title": (doc.metadata or {}).get("title") or ""}
    finally:
        doc.close()
    dropped: List[Tuple[int, str]] = []
    removed = _drop_headers_footers([(items, h) for items, h, _ in raw], dropped)
    if removed:
        meta["header_footer_removed"] = removed
    # 스캔 쪽 후보: 그림이 쪽의 큰 부분을 덮고, 머리글·쪽 번호를 뺀 뒤 남은 글이 아주 적은 쪽
    # (스캔 위에 찍은 'RECEIVED' 도장, 한 쪽짜리 스캔의 Bates 번호처럼 머리글로 뺄 수 없는 짧은 글은 남아도 스캔 쪽).
    # 빈 쪽·간지(로고만 있는 쪽 포함)는 OCR 대상이 아님. 보이지 않는 글자 층(OCR 결과)이 있는 쪽은 이미 OCR한 쪽이라 제외.
    scanned = [pno for pno, (items, _, _) in enumerate(raw, 1)
               if image_cover[pno - 1] >= SCAN_IMAGE_COVER and not ocr_layer[pno - 1]
               and _items_chars(items) < SCAN_TEXT_MAX]
    pages = [_order_items(items, w) for items, _, w in raw]

    sizes: Counter = Counter()
    bold_chars = total_chars = 0
    for blocks in pages:
        for lines in blocks:
            if isinstance(lines, dict):
                continue
            for text, size, bold in lines:
                sizes[round(size, 1)] += len(text)
                total_chars += len(text)
                bold_chars += len(text) if bold else 0
    body = sizes.most_common(1)[0][0] if sizes else 0.0
    bold_is_signal = bold_chars < total_chars * 0.5
    meta["body_font_size"] = body

    b = UnitBuilder()
    ended_with: Optional[List[str]] = None       # 앞 쪽의 마지막 덩이가 표였으면 그 표의 머리글 (다음 쪽으로 이어진 표에 다시 씀)
    for pno, blocks in enumerate(pages, 1):
        loc = f"p.{pno}"
        table_no = 0
        carried, ended_with = ended_with, None
        for k_block, lines in enumerate(blocks):
            if isinstance(lines, dict):           # 표: 행 단위 (머리글: 값)
                table_no += 1
                grid = clean_table(lines["table"])
                carry = carried if k_block == 0 else None      # 쪽 맨 앞의 표가 앞 쪽 끝 표와 열 수가 같으면 이어진 표
                for r in table_rows(lines["table"], carry=carry):
                    b.add("row", f"p.{pno} 표 {table_no} {r['row']}행", r["text"], pno, "pdf", r["number"])
                if k_block == len(blocks) - 1 and grid:
                    ended_with = (carry if carry is not None and len(carry) == max(len(r) for r in grid)
                                  else table_header(grid))
                continue
            buf: List[str] = []

            def flush() -> None:
                if buf:
                    b.add_text(loc, " ".join(buf), pno, "pdf", allow_heading=False)
                    buf.clear()

            prev_line = ""
            for k_line, (text, size, bold) in enumerate(lines):
                strong = size >= body + 1.0 or (bold_is_signal and bold)
                info = classify_number(text)
                # 블록 안의 둘째 줄부터는 줄바꿈된 본문일 수 있음: 큰 글씨·굵게이거나 앞 줄이 문장 끝일 때만 제목·새 항목
                # ('...limit of' 다음 줄 '1.5 times the rated current for'가 장절 제목이 되던 것, 2026-11-09 검토 A6)
                wrapped = k_line > 0 and not strong and not _SENTENCE_END.search(prev_line) and not _LIST_CONTINUES.search(prev_line) \
                    and not (prev_line.rstrip().endswith(":"))
                prev_line = text
                if wrapped and info is not None and info["family"] != "bullet":
                    buf.append(text)
                elif heading_title(info, text, strong):
                    flush()
                    b.add_text(loc, text, pno, "pdf", strong=strong, size=size)
                elif info is not None:          # 번호/글머리로 시작하는 줄은 새 단락
                    flush()
                    buf.append(text)
                else:
                    buf.append(text)
            flush()
    seen_margin: set = set()
    for pno, text in dropped:                    # 반복 머리글·바닥글 (쪽 번호만인 줄은 뺌): 글마다 처음 나온 쪽에 한 번
        text = " ".join(text.split())
        if text in seen_margin or not _LETTER.search(text):
            continue
        seen_margin.add(text)
        b.add("paragraph", f"p.{pno} 머리글·바닥글", text, pno, "pdf", None, section=["머리글·바닥글"])["extra"] = True
    if not b.units or (scanned and len(scanned) == len(raw)):
        # 글이 없거나, 모든 쪽이 스캔 쪽(도장·Bates 번호 같은 짧은 글만 있음)
        meta["ocr_needed"] = True
        rt.log("WARN", "본문 글자가 없는 PDF입니다 (스캔 이미지일 가능성). OCR이 필요합니다.")
    if scanned and b.units:
        # 일부 쪽만 스캔된 PDF: 그 쪽은 검색되지 않으므로 0건이어도 '없다'고 단정하면 안 됨
        meta["pages_without_text"] = scanned
        rt.log("WARN", f"글이 거의 없고 그림이 덮는 쪽 {len(scanned)}개 (스캔 쪽일 가능성, OCR 필요): p.{_page_list(scanned)}")
    return b.units, meta


def _items_chars(items: List[Any]) -> int:
    """쪽에 남은 글자 수 (머리글·쪽 번호를 뺀 뒤). 표가 있으면 글이 있는 쪽으로 봅니다."""
    total = 0
    for _, payload in items:
        if isinstance(payload, dict):
            return SCAN_TEXT_MAX
        total += sum(len(re.sub(r"\s+", "", t)) for t, _, _ in payload)
    return total


def _page_has_ocr_layer(page: Any) -> bool:
    """보이지 않는 글자(그리기 방식 3, OCR 프로그램이 그림 위에 얹는 글자 층)가 있는지."""
    try:
        return any(span.get("type") == 3 and span.get("chars") for span in page.get_texttrace())
    except Exception:
        return False


def _page_image_cover(page: Any) -> float:
    """쪽 면적 중 그림이 덮는 비율 (0~1). 그림 면적의 합이라 겹친 그림은 두 번 셀 수 있어 1에서 자릅니다."""
    try:
        rect = page.rect
        area = float(rect.width) * float(rect.height)
        if area <= 0:
            return 0.0
        total = 0.0
        for info in page.get_image_info():
            x0, y0, x1, y1 = info.get("bbox") or (0, 0, 0, 0)
            # 쪽 밖으로 나간 부분은 셈하지 않음
            w = min(x1, rect.x1) - max(x0, rect.x0)
            h = min(y1, rect.y1) - max(y0, rect.y0)
            if w > 0 and h > 0:
                total += w * h
        return min(1.0, total / area)
    except Exception:
        return 0.0


def _number_ranges(nums: List[int]) -> str:
    """[2,3,4,9,10] -> '2-4,9-10' (--pages·--units 값)."""
    runs: List[List[int]] = []
    for n in sorted(set(int(x) for x in nums)):
        if runs and n == runs[-1][1] + 1:
            runs[-1][1] = n
        else:
            runs.append([n, n])
    return ",".join(str(a) if a == b else f"{a}-{b}" for a, b in runs)


def _page_list(pages: List[int], limit: int = 10) -> str:
    shown = ",".join(str(p) for p in pages[:limit])
    return shown + (f" 외 {len(pages) - limit}쪽" if len(pages) > limit else "")


# HWP 5.0 제어 문자: 1 WCHAR 짜리 문자 컨트롤. 나머지 1~31은 인라인/확장 컨트롤로 8 WCHAR(16바이트)를 차지합니다.
_HWP_CHAR_CONTROLS = {0, 10, 13, 24, 25, 26, 27, 28, 29, 30, 31}
_HWP_CHAR_MAP = {10: "\n", 24: "-", 30: " ", 31: " "}
_HWPTAG_PARA_TEXT = 67


def _hwp_para_text(rec: bytes) -> str:
    out: List[str] = []
    i, n = 0, len(rec)
    while i + 1 < n:
        code = rec[i] | (rec[i + 1] << 8)
        if 0xD800 <= code <= 0xDBFF and i + 3 < n:          # UTF-16 서로게이트 쌍 (확장 한자 등 BMP 밖 글자)
            low = rec[i + 2] | (rec[i + 3] << 8)
            if 0xDC00 <= low <= 0xDFFF:
                out.append(chr(0x10000 + ((code - 0xD800) << 10) + (low - 0xDC00)))
                i += 4
                continue
        if 0xD800 <= code <= 0xDFFF:                         # 짝이 없는 서로게이트: 캐시 저장(UTF-8)이 실패하지 않게
            out.append("\ufffd")
            i += 2
        elif code >= 32:
            out.append(chr(code))
            i += 2
        elif code in _HWP_CHAR_CONTROLS:
            out.append(_HWP_CHAR_MAP.get(code, ""))
            i += 2
        else:
            if code == 9:
                out.append("\t")
            i += 16
    return "".join(out)


def parse_hwp(path: str) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    olefile = rt.require_lib("olefile", "olefile", "HWP 읽기")
    if not olefile.isOleFile(path):
        raise _encrypted_hint(path)
    try:
        ole = olefile.OleFileIO(path)
    except (PermissionError, FileNotFoundError):
        raise
    except (OSError, ValueError, IndexError) as e:      # olefile은 구조 손상을 OSError로 냄
        raise _DamagedFile(str(e)) from e
    try:
        if not ole.exists("FileHeader"):                # 이름만 .hwp인 다른 OLE 문서(.doc, .xls 등)
            raise rt.InputError(f"HWP 5.0 문서가 아닙니다 (FileHeader 없음): {os.path.basename(path)}. "
                                "확장자가 실제 형식과 맞는지 확인하세요 (.doc, .xls 등을 이름만 바꾼 파일일 수 있음).")
        try:
            header = ole.openstream("FileHeader").read()
        except (PermissionError, FileNotFoundError):
            raise
        except (OSError, ValueError, IndexError) as e:
            raise _DamagedFile(str(e)) from e
        props = int.from_bytes(header[36:40], "little") if len(header) >= 40 else 0
        compressed = bool(props & 0x1)
        if props & 0x2:
            raise rt.InputError(f"암호가 걸린 HWP 문서입니다: {os.path.basename(path)}")
        if props & 0x4:
            raise rt.InputError(f"배포용 HWP 문서라 본문을 추출할 수 없습니다: {os.path.basename(path)}")

        sections = [d for d in ole.listdir() if len(d) == 2 and d[0] == "BodyText" and d[1].startswith("Section")]
        sections.sort(key=lambda d: int(re.sub(r"\D", "", d[1]) or 0))
        paragraphs: List[str] = []
        for sec in sections:
            try:
                data = ole.openstream(sec).read()
            except (PermissionError, FileNotFoundError):
                raise
            except (OSError, ValueError, IndexError) as e:
                raise _DamagedFile(str(e)) from e
            if compressed:
                data = zlib.decompress(data, -15)
            pos, size_total = 0, len(data)
            while pos + 4 <= size_total:
                h = int.from_bytes(data[pos:pos + 4], "little")
                pos += 4
                tag, size = h & 0x3FF, (h >> 20) & 0xFFF
                if size == 0xFFF:
                    size = int.from_bytes(data[pos:pos + 4], "little")
                    pos += 4
                if tag == _HWPTAG_PARA_TEXT:
                    text = _hwp_para_text(data[pos:pos + size]).strip()
                    if text:
                        paragraphs.append(text)
                pos += size
    finally:
        ole.close()
    b = UnitBuilder()
    for i, t in enumerate(paragraphs, 1):
        b.add_text(f"단락 {i}", t)
    return b.units, {"format": "HWP 5.0", "sections": len(sections)}


def _local(tag: Any) -> str:
    return str(tag).rsplit("}", 1)[-1]


# 개요 번호 형식 (HWPX numFormat / Word numFmt 공통으로 씀)
def _roman(n: int) -> str:
    vals = [(1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"), (50, "L"), (40, "XL"),
            (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")]
    out = ""
    for v, r in vals:
        while n >= v:
            out += r
            n -= v
    return out or str(n)


def _letters(n: int) -> str:
    out = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        out = chr(65 + r) + out
    return out or "A"


def format_counter(n: int, fmt: str) -> str:
    """번호 형식 이름(Word numFmt, HWPX numFormat)으로 숫자를 글자로 바꿉니다. 원문자 형식은 '(n)'으로 씁니다."""
    f = (fmt or "decimal").lower().replace("_", "")
    if n <= 0:
        n = 1
    if f in ("upperroman", "romancapital"):
        return _roman(n)
    if f in ("lowerroman", "romansmall"):
        return _roman(n).lower()
    if f in ("upperletter", "latincapital"):
        return _letters(n)
    if f in ("lowerletter", "latinsmall"):
        return _letters(n).lower()
    if f in ("ganada", "hangulsyllable"):
        return GANADA[(n - 1) % len(GANADA)]
    if f in ("chosung", "hanguljamo"):
        return CHOSUNG[(n - 1) % len(CHOSUNG)]
    if f in ("circledhangulsyllable",):
        return f"({GANADA[(n - 1) % len(GANADA)]})"
    if "circle" in f or "enclosed" in f:
        return f"({n})"
    if f == "decimalzero":
        return f"{n:02d}"
    return str(n)


def _hwpx_header_info(z: zipfile.ZipFile) -> Tuple[Dict[str, Tuple[str, str, int]], Dict[str, Dict[int, Tuple[int, str, str]]]]:
    """header.xml에서 (문단모양 id -> (머리 종류, 번호 id, 수준)), (번호 id -> {수준: (시작, 형식, 글자틀)})."""
    para: Dict[str, Tuple[str, str, int]] = {}
    numbering: Dict[str, Dict[int, Tuple[int, str, str]]] = {}
    names = [n for n in z.namelist() if n.lower().endswith("header.xml")]
    if not names:
        return para, numbering
    root = ET.fromstring(z.read(names[0]))
    for el in root.iter():
        tag = _local(el.tag)
        if tag == "paraPr" and el.get("id") is not None:
            for sub in el.iter():
                if _local(sub.tag) == "heading":
                    try:
                        para[el.get("id", "")] = (sub.get("type", "NONE").upper(), sub.get("idRef", ""),
                                                  int(sub.get("level", "0")))
                    except ValueError:
                        pass
                    break
        elif tag == "numbering" and el.get("id") is not None:
            levels: Dict[int, Tuple[int, str, str]] = {}
            for sub in el:
                if _local(sub.tag) == "paraHead":
                    try:
                        levels[int(sub.get("level", "1"))] = (int(sub.get("start", "1")), sub.get("numFormat", "DIGIT"),
                                                              sub.text or "")
                    except ValueError:
                        pass
            numbering[el.get("id", "")] = levels
    return para, numbering


def _hwpx_walk(elem: ET.Element, buf: Optional[List[str]], out: List[Dict[str, Any]]) -> None:
    """문단(<p>)마다 자기 텍스트(<t>)만 모읍니다. 표 안의 문단은 별도 문단으로, 문서 순서를 유지합니다."""
    tag = _local(elem.tag)
    if tag == "p":
        slot = len(out)
        out.append({"text": "", "para": elem.get("paraPrIDRef", ""), "page_break": elem.get("pageBreak") == "1"})
        own: List[str] = []
        for child in elem:
            _hwpx_walk(child, own, out)
        out[slot]["text"] = "".join(own)
        return
    if tag == "t" and buf is not None:
        if elem.text:
            buf.append(elem.text)
        for child in elem:
            if _local(child.tag) == "tab":
                buf.append("\t")
            if child.tail:
                buf.append(child.tail)
        return
    for child in elem:
        _hwpx_walk(child, buf, out)


def parse_hwpx(path: str) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    if not zipfile.is_zipfile(path):
        raise _encrypted_hint(path)
    with zipfile.ZipFile(path) as z:
        names = [n for n in z.namelist() if re.match(r"Contents/section\d+\.xml$", n)]
        if not names:
            raise rt.InputError(f"HWPX 본문(Contents/section*.xml)이 없습니다: {os.path.basename(path)}")
        names.sort(key=lambda n: int(re.sub(r"\D", "", n)))
        try:
            para_heads, numbering = _hwpx_header_info(z)
        except Exception as e:                      # 개요 정보가 없으면 번호 패턴으로만 인식
            rt.debug(f"HWPX header.xml 해석 실패: {e}")
            para_heads, numbering = {}, {}
        paragraphs: List[Dict[str, Any]] = []
        for name in names:
            _hwpx_walk(ET.fromstring(z.read(name)), None, paragraphs)
    use_pages = any(p["page_break"] for p in paragraphs)
    page = 1
    counters: Dict[str, List[int]] = {}
    b = UnitBuilder()
    n = 0
    for p in paragraphs:
        if p["page_break"] and b.units:
            page += 1
        text = (p["text"] or "").strip()
        if not text:
            continue
        n += 1
        kind_level = para_heads.get(p["para"])
        style_level = None
        if kind_level and kind_level[0] == "OUTLINE":
            style_level = kind_level[2] + 1
            levels = numbering.get(kind_level[1]) or {}
            fmt = levels.get(style_level)
            if fmt:                                  # 개요 번호(자동 번호)를 글 앞에 붙임: ^1.^2 -> 1.2
                cnt = counters.setdefault(kind_level[1], [0] * 11)
                cnt[style_level] += 1
                for k in range(style_level + 1, 11):
                    cnt[k] = 0

                def repl(m: "re.Match[str]") -> str:
                    lv = int(m.group(1))
                    start, numfmt, _ = levels.get(lv, (1, "DIGIT", ""))
                    return format_counter(max(cnt[lv], 1) + start - 1, numfmt)
                label = re.sub(r"\^(\d+)", repl, fmt[2]).strip()
                if label and not text.startswith(label):
                    text = f"{label} {text}"
        b.add_text(f"단락 {n}", text, page if use_pages else None, "page_break" if use_pages else None,
                   style_level=style_level)
    return b.units, {"format": "HWPX", "sections": len(names)}


# ---------------------------------------------------------------- DOCX

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def _wval(el: Any, path: str) -> Optional[str]:
    """요소 아래 path(w: 태그를 '/'로 이은 것)의 w:val 값."""
    cur = el
    for part in path.split("/"):
        if cur is None:
            return None
        cur = cur.find(_W + part)
    return None if cur is None else cur.get(_W + "val")


class _DocxNumbering:
    """numbering.xml로 Word 자동 번호 글자(예: 3.2.1.)를 계산합니다."""

    def __init__(self, document: Any) -> None:
        self.abstract: Dict[str, Dict[int, Dict[str, Any]]] = {}
        self.nums: Dict[str, Tuple[str, Dict[int, int]]] = {}
        self.counters: Dict[str, List[Optional[int]]] = {}
        self.seen: set = set()
        try:
            root = document.part.numbering_part.element
        except Exception:
            return
        for an in root.findall(_W + "abstractNum"):
            aid = an.get(_W + "abstractNumId", "")
            lvls: Dict[int, Dict[str, Any]] = {}
            for lvl in an.findall(_W + "lvl"):
                try:
                    ilvl = int(lvl.get(_W + "ilvl", "0"))
                    start = int(_wval(lvl, "start") or 1)
                except ValueError:
                    continue
                lvls[ilvl] = {"start": start, "fmt": _wval(lvl, "numFmt") or "decimal",
                              "text": _wval(lvl, "lvlText") or "", "lgl": lvl.find(_W + "isLgl") is not None,
                              "style": _wval(lvl, "pStyle")}
            self.abstract[aid] = lvls
        for num in root.findall(_W + "num"):
            nid = num.get(_W + "numId", "")
            aid = _wval(num, "abstractNumId") or ""
            overrides: Dict[int, int] = {}
            for ov in num.findall(_W + "lvlOverride"):
                so = _wval(ov, "startOverride")
                if so is not None:
                    try:
                        overrides[int(ov.get(_W + "ilvl", "0"))] = int(so)
                    except ValueError:
                        pass
            self.nums[nid] = (aid, overrides)

    def level_for_style(self, num_id: str, style_id: str) -> Optional[int]:
        aid = self.nums.get(num_id, ("", {}))[0]
        for ilvl, lv in self.abstract.get(aid, {}).items():
            if lv.get("style") == style_id:
                return ilvl
        return None

    def label(self, num_id: str, ilvl: int) -> Optional[Tuple[str, str]]:
        """(번호 글자, 계열 이름). 글머리 기호나 번호 없음은 ('', 계열)."""
        if not num_id or num_id == "0" or num_id not in self.nums:
            return None
        aid, overrides = self.nums[num_id]
        lvls = self.abstract.get(aid)
        if not lvls or ilvl not in lvls:
            return None
        cnt = self.counters.setdefault(aid, [None] * 10)
        if num_id not in self.seen:
            self.seen.add(num_id)
            for lv, start in overrides.items():
                if 0 <= lv < 10:
                    cnt[lv] = start - 1
                    for k in range(lv + 1, 10):
                        cnt[k] = None
        cur = cnt[ilvl]
        cnt[ilvl] = (cur + 1) if cur is not None else lvls[ilvl]["start"]
        for k in range(ilvl + 1, 10):
            cnt[k] = None
        info = lvls[ilvl]
        fam = f"wlist:{aid}"
        if info["fmt"] in ("bullet", "none"):
            return "", fam

        def repl(m: "re.Match[str]") -> str:
            k = int(m.group(1)) - 1
            lv = lvls.get(k, {"start": 1, "fmt": "decimal"})
            value = cnt[k] if cnt[k] is not None else lv["start"]
            return format_counter(value, "decimal" if info["lgl"] else lv["fmt"])
        return re.sub(r"%(\d)", repl, info["text"]).strip(), fam


def _style_chain(style: Any) -> List[Any]:
    chain = []
    while style is not None and len(chain) < 10:
        chain.append(style)
        style = getattr(style, "base_style", None)
    return chain


def _docx_style_info(para: Any, numbering: _DocxNumbering) -> Tuple[Optional[int], bool, Optional[Tuple[str, int]]]:
    """(제목 수준, 문서 제목 여부, (numId, ilvl)). 제목 수준: Heading N / 제목 N / 개요 수준(outlineLvl)."""
    style = para.style
    name = (style.name if style is not None else "") or ""
    level: Optional[int] = None
    is_title = name.strip().lower() in ("title", "subtitle") or name.strip() in ("제목", "부제")
    m = re.match(r"(?i)(heading|제목)\s*(\d)$", name.strip())
    if m:
        level = int(m.group(2))
    ppr = para._p.pPr
    outline = _wval(ppr, "outlineLvl") if ppr is not None else None
    chain = _style_chain(style)
    if outline is None:
        for st in chain:
            spr = st.element.find(_W + "pPr")
            outline = _wval(spr, "outlineLvl") if spr is not None else None
            if outline is not None:
                break
    if level is None and outline is not None and outline.isdigit() and int(outline) < 9:
        level = int(outline) + 1
    num: Optional[Tuple[str, int]] = None
    if ppr is not None and ppr.find(_W + "numPr") is not None:
        nid = _wval(ppr, "numPr/numId")
        ilvl = _wval(ppr, "numPr/ilvl")
        if nid:
            num = (nid, int(ilvl) if ilvl and ilvl.isdigit() else 0)
    if num is None:
        for st in chain:
            spr = st.element.find(_W + "pPr")
            nid = _wval(spr, "numPr/numId") if spr is not None else None
            if nid:
                ilvl = _wval(spr, "numPr/ilvl")
                lv = numbering.level_for_style(nid, st.style_id)
                num = (nid, lv if lv is not None else (int(ilvl) if ilvl and ilvl.isdigit() else 0))
                break
    return level, is_title, num


class _DocxPages:
    """저장된 쪽 정보로 단락의 쪽 번호를 셉니다.
    w:lastRenderedPageBreak(Word가 저장할 때 기록한 쪽 경계)가 있으면 그것만 쓰고(word_saved_layout),
    없으면 강제 쪽 나눔(w:br type=page, pageBreakBefore, 다음 쪽 구역 나누기)만 셉니다(page_break, 실제보다 적을 수 있음)."""

    def __init__(self, body: Any) -> None:
        self.mode: Optional[str] = None
        if any(True for _ in body.iter(_W + "lastRenderedPageBreak")):
            self.mode = "word_saved_layout"
        elif (any(br.get(_W + "type") == "page" for br in body.iter(_W + "br"))
              or any(True for _ in body.iter(_W + "pageBreakBefore"))
              or any(self._section_break(s) for s in body.iter(_W + "sectPr") if s.getparent().tag == _W + "pPr")):
            self.mode = "page_break"
        self.page = 1
        self.pending = 0
        self.started = False

    @staticmethod
    def _section_break(sect: Any) -> bool:
        t = _wval(sect, "type")
        return t not in ("continuous",)

    def element_page(self, el: Any) -> Optional[int]:
        """요소(단락/표)의 첫 글자가 있는 쪽. 요소 안의 쪽 경계도 반영해 다음 요소로 넘깁니다."""
        if self.mode is None:
            return None
        self.page += self.pending
        self.pending = 0
        if self.mode == "page_break" and self.started:
            ppr = el.find(_W + "pPr") if el.tag == _W + "p" else None
            if ppr is not None and ppr.find(_W + "pageBreakBefore") is not None:
                v = ppr.find(_W + "pageBreakBefore").get(_W + "val")
                if v not in ("0", "false"):
                    self.page += 1
        first: Optional[int] = None
        for node in el.iter():
            tag = node.tag
            if tag == _W + "t" and node.text and node.text.strip():
                if first is None:
                    first = self.page
                self.started = True
            elif self.mode == "word_saved_layout" and tag == _W + "lastRenderedPageBreak":
                if self.started or first is not None:
                    self.page += 1
            elif self.mode == "page_break" and tag == _W + "br" and node.get(_W + "type") == "page":
                if self.started or first is not None:
                    self.page += 1
            elif (self.mode == "page_break" and tag == _W + "sectPr" and node.getparent().tag == _W + "pPr"
                  and self._section_break(node)):
                self.pending += 1
        return first if first is not None else self.page


_SKIP_TEXT_IN = {_W + "del", _W + "moveFrom", _W + "txbxContent"}   # 지운 글, 옮기기 전 자리, 글상자(따로 단락)


def _p_text(el: Any) -> str:
    """단락(w:p)의 글. python-docx의 Paragraph.text와 달리 변경 추적으로 삽입된 글(w:ins)과 옮긴 글(w:moveTo)도 넣고,
    지운 글(w:del)은 뺍니다. 탭과 줄바꿈은 글자로."""
    out: List[str] = []

    def walk(node: Any) -> None:
        for ch in node:
            tag = ch.tag
            if tag in _SKIP_TEXT_IN:
                continue
            if tag == _W + "t":
                out.append(ch.text or "")
            elif tag == _W + "tab":
                out.append("\t")
            elif tag in (_W + "br", _W + "cr"):
                out.append("\n")
            else:
                walk(ch)
    walk(el)
    return "".join(out)


_MC_FALLBACK = "{http://schemas.openxmlformats.org/markup-compatibility/2006}Fallback"


def _textboxes(el: Any) -> List[Any]:
    """요소 안의 글상자 내용(w:txbxContent), 문서 순서. Word 2010 이후 파일은 같은 글상자를 DrawingML(mc:Choice)과
    VML(mc:Fallback) 두 벌로 저장하므로 mc:Fallback 안의 것은 뺍니다 (두 번 세지 않게). 글상자 안의 글상자도 각각 하나로
    (_p_text가 안쪽 글상자 글을 빼므로 겹치지 않음)."""
    out = []
    for tb in el.iter(_W + "txbxContent"):
        node = tb.getparent()
        while node is not None and node is not el and node.tag != _MC_FALLBACK:
            node = node.getparent()
        if node is None or node.tag != _MC_FALLBACK:
            out.append(tb)
    return out


def _textbox_lines(tb: Any) -> List[str]:
    """글상자 하나의 글: 단락마다 한 줄, 표는 행마다 칸을 탭으로 이은 한 줄. 빈 줄은 뺍니다."""
    lines = []
    for block in _body_blocks(tb):
        if _local(block.tag) == "p":
            lines.append(_p_text(block).strip())
        else:
            for tr in block.iter(_W + "tr"):
                lines.append("\t".join(" ".join(_p_text(p).strip() for p in tc.iter(_W + "p")).strip()
                                       for tc in tr.iter(_W + "tc")).strip())
    return [ln for ln in lines if ln]


def _cell_paragraphs(tc: Any) -> List[Any]:
    """표 칸의 단락 (글상자 안 단락은 빼고, 칸 안의 표 단락은 넣음): 글상자 글은 따로 단위로 둠 (두 벌 저장도 두 번 세지 않음)."""
    def in_textbox(p: Any) -> bool:
        node = p.getparent()
        while node is not None and node is not tc:
            if node.tag == _W + "txbxContent":
                return True
            node = node.getparent()
        return False
    return [p for p in tc.iter(_W + "p") if not in_textbox(p)]


def _body_blocks(parent: Any):
    """본문의 단락(w:p)과 표(w:tbl)를 문서 순서로. 콘텐츠 컨트롤(w:sdt)과 customXml 안에 든 것도 꺼냅니다."""
    for child in parent.iterchildren():
        tag = child.tag
        if tag in (_W + "p", _W + "tbl"):
            yield child
        elif tag == _W + "sdt":
            content = child.find(_W + "sdtContent")
            if content is not None:
                yield from _body_blocks(content)
        elif tag == _W + "customXml":
            yield from _body_blocks(child)


def _docx_extras(document: Any, path: str) -> List[Tuple[str, str, str]]:
    """DOCX 본문 밖의 글: 머리글·바닥글(구역마다, 앞 구역과 같으면 뺌, 같은 글은 한 번), 각주·미주 -> [(위치, 글, 장절)].
    본문만 읽으면 문서 번호·개정 표기·각주의 조건이 검색되지 않음 (2026-11-09 검토 A13). 읽지 못하면 그 부분만 뺌."""
    out: List[Tuple[str, str, str]] = []
    seen: set = set()
    try:
        for sec in document.sections:
            for what, part in (("머리글", sec.header), ("머리글", sec.first_page_header), ("머리글", sec.even_page_header),
                               ("바닥글", sec.footer), ("바닥글", sec.first_page_footer), ("바닥글", sec.even_page_footer)):
                try:
                    if part.is_linked_to_previous:
                        continue
                    lines = [" ".join(_p_text(p._p).split()) for p in part.paragraphs]
                    for t in part.tables:
                        lines += [" | ".join(" ".join(c.text.split()) for c in row.cells) for row in t.rows]
                except Exception:
                    continue
                for line in lines:
                    if line and (what, line) not in seen and _LETTER.search(line):
                        seen.add((what, line))
                        out.append((what, line, "머리글·바닥글"))
    except Exception as e:                           # noqa: BLE001
        rt.debug(f"머리글·바닥글 읽기 실패: {type(e).__name__}")
    w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    try:
        import xml.etree.ElementTree as ET
        with zipfile.ZipFile(path) as z:
            names = set(z.namelist())
            for part, tag, label in (("word/footnotes.xml", "footnote", "각주"), ("word/endnotes.xml", "endnote", "미주")):
                if part not in names:
                    continue
                root = ET.fromstring(z.read(part))
                for note in root.iter(w + tag):
                    if note.get(w + "type") in ("separator", "continuationSeparator", "continuationNotice"):
                        continue
                    text = " ".join("".join(t.text or "" for t in p.iter(w + "t")) for p in note.iter(w + "p"))
                    text = " ".join(text.split())
                    if text:
                        out.append((f"{label} {note.get(w + 'id')}", text, label))
    except Exception as e:                           # noqa: BLE001
        rt.debug(f"각주 읽기 실패: {type(e).__name__}")
    return out


def parse_docx(path: str) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    if not zipfile.is_zipfile(path):
        raise _encrypted_hint(path)
    docx = rt.require_lib("docx", "python-docx", "DOCX 읽기")
    from docx.table import Table
    from docx.text.paragraph import Paragraph
    try:
        document = docx.Document(path)
    except Exception:
        raise _encrypted_hint(path)
    numbering = _DocxNumbering(document)
    pages = _DocxPages(document.element.body)
    b = UnitBuilder()
    para_no = table_no = box_no = 0
    numbering_failed = False

    def add_textboxes(el: Any, page: Optional[int]) -> None:
        """본문 단락·표에 든 글상자 글을 그 뒤에 따로 단위로 (단락 번호는 그대로: 글상자가 없는 문서의 단위 번호가 같음)."""
        nonlocal box_no
        for tb in _textboxes(el):
            lines = _textbox_lines(tb)
            if not lines:
                continue
            box_no += 1
            for j, line in enumerate(lines, 1):
                loc = f"글상자 {box_no}" + (f" 단락 {j}" if len(lines) > 1 else "")
                b.add("paragraph", loc, line, page, pages.mode)

    for child in _body_blocks(document.element.body):
        tag = _local(child.tag)
        if tag == "p":
            para = Paragraph(child, document)
            page = pages.element_page(child)
            text = _p_text(child).strip()
            if not text:
                # 빈 번호 단락도 Word는 번호를 매김 ("3." 빈 줄): 번호 카운터를 진행해야 뒤 조항 번호가 밀리지 않음
                try:
                    _, _, num = _docx_style_info(para, numbering)
                    if num:
                        numbering.label(*num)
                except Exception:
                    pass
                add_textboxes(child, page)                 # 글상자만 있는 단락 (글상자 글은 단락 글에 없음)
                continue
            para_no += 1
            loc = f"단락 {para_no}"
            try:
                level, is_title, num = _docx_style_info(para, numbering)
                lab = numbering.label(*num) if num else None
            except Exception as e:                  # 번호 해석 실패 -> 글의 번호 패턴만 사용
                if not numbering_failed:
                    rt.debug(f"Word 번호 해석 실패, 글 패턴으로 대신함: {e}")
                numbering_failed = True
                level, is_title, lab = None, False, None
            list_family, list_depth = None, 1
            if lab:
                label, list_family = lab
                list_depth = (num[1] + 1) if num else 1
                if label and not text.startswith(label):
                    text = f"{label} {text}"
            if is_title and not level:
                b.add("heading", loc, text, page, pages.mode, None, section=[text[:120]])
            else:
                b.add_text(loc, text, page, pages.mode, style_level=level, list_family=list_family, list_depth=list_depth)
            add_textboxes(child, page)
        elif tag == "tbl":
            table_no += 1
            grid: List[List[str]] = []
            row_pages: List[Optional[int]] = []
            for row in Table(child, document).rows:
                row_pages.append(pages.element_page(row._tr))      # 여러 쪽에 걸친 표: 행마다 그 행이 있는 쪽
                cells: List[str] = []
                prev_tc = None
                for cell in row.cells:
                    # 가로 병합 칸은 같은 칸이 반복되므로 빈 칸으로 (머리글은 앞 칸을 이어 씀)
                    cells.append("" if cell._tc is prev_tc else
                                 "\n".join(_p_text(p) for p in _cell_paragraphs(cell._tc)))
                    prev_tc = cell._tc
                grid.append(cells)
            for r in table_rows(grid, row_pages):
                b.add("row", f"표 {table_no} {r['row']}행", r["text"], r["key"], pages.mode, r["number"])
            add_textboxes(child, next((p for p in row_pages if p), None))
    for loc, text, sec in _docx_extras(document, path):     # 머리글·바닥글·각주·미주 (본문 뒤, 쪽 정보 없음)
        b.add("paragraph", loc, text, None, pages.mode, None, section=[sec])["extra"] = True
    core = document.core_properties
    meta: Dict[str, Any] = {"title": core.title or "", "paragraphs": para_no, "tables": table_no,
                            "page_source": pages.mode}
    if box_no:
        meta["textboxes"] = box_no
    if pages.mode:
        meta["pages"] = max([u["page"] for u in b.units if u["page"]] or [1])
    return b.units, meta


# ---------------------------------------------------------------- PPTX / XLSX / CSV / TXT

def _pptx_shape_lines(shape, out: List[str]) -> None:
    """도형 하나의 텍스트를 줄 단위로 모읍니다 (그룹은 재귀, 표는 행 단위)."""
    # 그룹은 shape_type 대신 하위 도형(shapes)이 있는지로: 종류를 알 수 없는 도형(prstGeom도 txBox도 없는 p:sp)은
    # shape_type이 NotImplementedError라 파일 전체가 ERROR였음 (2026-11-09 검토 B11)
    if hasattr(shape, "shapes"):
        for sub in shape.shapes:
            _pptx_shape_lines(sub, out)
    elif getattr(shape, "has_table", False):
        for row in shape.table.rows:
            out.append(" | ".join(c.text.strip().replace("\n", " ") for c in row.cells))
    elif getattr(shape, "has_text_frame", False):
        for para in shape.text_frame.paragraphs:
            text = para.text.replace("\v", " ").strip()
            if text:
                out.append(text)


def parse_pptx(path: str) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    """슬라이드 하나 = 단위 하나 (ppt.py/office.py가 슬라이드 번호로 씀). 장절 = [슬라이드 제목]."""
    if not zipfile.is_zipfile(path):
        raise _encrypted_hint(path)
    pptx = rt.require_lib("pptx", "python-pptx", "PPTX 읽기")
    try:
        prs = pptx.Presentation(path)
    except Exception:
        raise _encrypted_hint(path)
    b = UnitBuilder()
    for idx, slide in enumerate(prs.slides, 1):   # presentation.xml의 실제 슬라이드 순서
        title_shape = slide.shapes.title
        title = title_shape.text.strip() if title_shape is not None and title_shape.has_text_frame else ""
        lines: List[str] = []
        for shape in slide.shapes:
            if title_shape is not None and shape.shape_id == title_shape.shape_id:
                continue
            try:
                _pptx_shape_lines(shape, lines)
            except Exception as e:                   # noqa: BLE001  도형 하나 때문에 파일 전체를 못 읽지 않게
                rt.debug(f"슬라이드 {idx} 도형을 읽지 못함: {type(e).__name__}")
        if not title and lines:
            title = lines[0]
        notes = ""
        if slide.has_notes_slide:
            notes = (slide.notes_slide.notes_text_frame.text or "").strip()
        body = "\n".join(([title] if title and (not lines or lines[0] != title) else []) + lines)
        if notes:
            body += f"\n[노트] {notes}"
        title_line = title.replace("\n", " ").strip()
        info = classify_number(title_line)
        number = info["number"] if info and info["family"] != "bullet" and not info["plain"] else None
        b.add("slide", f"슬라이드 {idx}", body.strip(), idx, "slide", number,
              section=[title_line] if title_line else [])
    return b.units, {"slides": len(prs.slides), "pages": len(prs.slides)}


def _cell_str(v: Any) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    if hasattr(v, "isoformat"):
        return v.isoformat(sep=" ") if hasattr(v, "hour") else v.isoformat()
    return str(v)


def parse_xlsx(path: str) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    if not zipfile.is_zipfile(path):
        raise _encrypted_hint(path)
    openpyxl = rt.require_lib("openpyxl", "openpyxl", "XLSX 읽기")
    try:
        wb = rt.retry_fs(lambda: openpyxl.load_workbook(path, read_only=True, data_only=True))
    except Exception:
        raise _encrypted_hint(path)
    b = UnitBuilder()
    sheets: List[str] = []
    truncated = False
    try:
        for ws in wb.worksheets:
            sheets.append(ws.title)
            count = 0
            if hasattr(ws, "reset_dimensions"):
                ws.reset_dimensions()          # 크기 정보(<dimension>)가 실제보다 작은 파일도 끝까지 (2026-11-09 검토 A3)
            for r_idx, row in enumerate(ws.iter_rows(values_only=True), 1):
                cells = [_cell_str(v) for v in row]
                while cells and cells[-1] == "":
                    cells.pop()
                if not any(cells):
                    continue
                count += 1
                if count > MAX_SHEET_ROWS:
                    truncated = True
                    break
                b.add("row", f"{ws.title}!{r_idx}행", "\t".join(cells), section=[ws.title])
    finally:
        wb.close()
    meta: Dict[str, Any] = {"sheets": sheets}
    if truncated:
        meta["rows_truncated"] = True
        rt.log("WARN", f"시트당 {MAX_SHEET_ROWS}행까지만 읽었습니다.")
    # 계산된 값이 저장되지 않은 수식 칸 (openpyxl 등이 쓰고 엑셀에서 연 적 없는 파일): data_only로는 빈 칸이라 글이 없음
    uncomputed = xf.xlsx_uncomputed_formulas(path)
    _uncomputed_from_check(meta, uncomputed)           # 확인 실패도 남김: 캐시를 쓸 때 recheck_uncomputed가 다시 확인
    if uncomputed.get("unknown"):
        rt.log("WARN", f"수식 계산 여부를 확인하지 못함: {uncomputed.get('error')}")
    elif uncomputed["count"]:
        rt.log("WARN", f"계산된 값이 저장되지 않은 수식 칸 {uncomputed['count']}개 (빈 칸으로 읽음): {', '.join(uncomputed['cells'][:5])}")
    return b.units, meta


def parse_csv(path: str) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    rows = rt.read_csv_rows(path)
    b = UnitBuilder()
    for i, r in enumerate(rows, 1):
        if any(c.strip() for c in r):
            b.add("row", f"{i}행", "\t".join(r), section=[])
    return b.units, {"rows": len(rows)}


def parse_txt(path: str) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    """빈 줄로 나눈 단락. 단락 첫 줄이 제목(마크다운 #, 또는 번호 제목)이면 제목 단위로 따로 둡니다."""
    text = rt.read_text_file(path)
    b = UnitBuilder()
    n = 0
    for block in [p.strip() for p in re.split(r"\n\s*\n", text) if p.strip()]:
        lines = block.splitlines()
        first = lines[0].strip()
        md = re.match(r"^(#{1,6})\s+(.*)$", first)
        if md:
            n += 1
            b.add_text(f"단락 {n}", md.group(2).strip().rstrip("#").strip(), style_level=len(md.group(1)))
            rest = "\n".join(lines[1:]).strip()
        elif len(lines) > 1 and heading_title(classify_number(first), first):
            n += 1
            b.add_text(f"단락 {n}", first)
            rest = "\n".join(lines[1:]).strip()
        else:
            rest = block
        if rest:
            n += 1
            b.add_text(f"단락 {n}", rest, allow_heading=len(rest.splitlines()) == 1)
    return b.units, {}


# ---------------------------------------------------------------- EML (메일 파일, 1.5.0)
# 헤더·본문·첨부 이름만 읽음 (첨부 안의 글은 읽지 않음: 사용자 결정 2026-11-10, 원본 메일은 Outlook에서 연다).
# 시험 스크립트 tools/eml_read_trial.py(회사 PC에서 잘 읽힘)와 같은 규칙.

EML_FORWARD_DEPTH = 2        # 전달된 메일(메일 안의 메일)을 읽는 깊이


def _eml_bytes(raw: bytes, declared: Optional[str] = None) -> str:
    for enc in (declared, "utf-8", "cp949", "euc-kr"):
        if not enc:
            continue
        try:
            return raw.decode(enc)
        except (UnicodeDecodeError, LookupError):
            continue
    return raw.decode("latin-1")


def _eml_header(value: Any) -> str:
    """RFC2047 인코딩과 인코딩 표시 없는 8비트(CP949 등) 헤더."""
    from email.header import decode_header, make_header
    if value is None:
        return ""
    s = str(value)
    try:
        raw = s.encode("ascii", "surrogateescape")
        if any(b > 127 for b in raw):
            return _eml_bytes(raw).strip()
    except UnicodeEncodeError:
        pass
    try:
        return " ".join(str(make_header(decode_header(s))).split())
    except Exception:
        return " ".join(s.split())


def _eml_html_text(html: str) -> str:
    import html as _html
    html = re.sub(r"(?is)<(script|style|head).*?</\1>", " ", html)
    html = re.sub(r"(?i)<br\s*/?>", "\n", html)
    html = re.sub(r"(?i)</p>|</div>|</tr>|</li>|</h\d>|</table>", "\n\n", html)       # 단락 끝: 빈 줄 (본문 단락으로 나눔)
    text = _html.unescape(re.sub(r"<[^>]+>", " ", html))
    lines = [" ".join(ln.split()) for ln in text.splitlines()]
    out: List[str] = []
    for ln in lines:                              # 빈 줄 하나로 단락을 나눔
        if ln or (out and out[-1]):
            out.append(ln)
    return "\n".join(out).strip()


def _eml_message(msg: Any) -> Dict[str, Any]:
    """email.message -> {subject, from_name, from_addr, to, cc, date(현지 표기 글), body, attachments, forwarded[메일...]}."""
    from email.utils import getaddresses, parsedate_to_datetime
    raw: Dict[str, Any] = {}
    for k, v in msg.raw_items():
        raw.setdefault(k.lower(), v)
    frm = _eml_header(raw.get("from"))
    name, addr = (getaddresses([frm]) or [("", "")])[0]
    date = ""
    if msg.get("Date"):
        try:
            date = parsedate_to_datetime(str(msg.get("Date"))).strftime("%Y-%m-%d %H:%M")   # 메일에 적힌 그대로의 시각
        except Exception:
            date = ""
    plain: List[str] = []
    html: List[str] = []
    atts: List[str] = []
    forwarded: List[Any] = []

    def walk(part: Any) -> None:
        ctype = part.get_content_type()
        if ctype == "message/rfc822":             # 전달된 메일: 본문에 섞지 않고 따로
            for sub in (part.get_payload() or []):
                forwarded.append(sub)
            return
        if part.is_multipart():
            for sub in part.get_payload():
                walk(sub)
            return
        filename = part.get_filename()
        disp = (part.get("Content-Disposition") or "").lower()
        if filename or disp.startswith("attachment"):
            atts.append(_eml_header(filename) if filename else "이름 없는 첨부")
            return
        payload = part.get_payload(decode=True)
        if payload is None:
            return
        text = _eml_bytes(payload, part.get_content_charset())
        if ctype == "text/plain":
            plain.append(text)
        elif ctype == "text/html":
            html.append(text)

    walk(msg)
    body = "\n\n".join(t.replace("\r\n", "\n").strip() for t in plain if t.strip()) or _eml_html_text("\n".join(html))
    return {"subject": _eml_header(raw.get("subject")), "from_name": name or addr, "from_addr": addr,
            "to": _eml_header(raw.get("to")), "cc": _eml_header(raw.get("cc")), "date": date, "body": body,
            "attachments": atts, "forwarded": forwarded}


def eml_label(mail: Dict[str, Any]) -> str:
    """근거·장절에 쓸 메일 표시: '제목 (2026-03-02 14:05, 김철수)'."""
    who = mail.get("from_name") or mail.get("from_addr") or "보낸 사람 없음"
    when = mail.get("date") or "날짜 없음"
    return f"{mail.get('subject') or '(제목 없음)'} ({when}, {who})"


def parse_eml(path: str) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    """메일 한 통: 제목·헤더(보낸 사람, 받는 사람, 참조, 날짜)·본문 단락·첨부 이름 단위, 전달된 메일은 그 안의 같은 단위를
    '전달된 메일'로 (장절 = 메일 표시). 첨부 안의 글은 읽지 않음."""
    import email
    from email import policy
    try:
        with open(path, "rb") as f:
            msg = email.message_from_binary_file(f, policy=policy.compat32)
    except OSError as e:
        raise rt.InputError(f"메일 파일을 열 수 없습니다: {os.path.basename(path)} ({type(e).__name__})")
    top = _eml_message(msg)
    if not (top["subject"] or top["from_addr"] or top["body"]):
        raise rt.InputError(f"메일 파일로 읽을 수 없습니다 (제목·보낸 사람·본문이 없음): {os.path.basename(path)}")
    b = UnitBuilder()

    def add_mail(mail: Dict[str, Any], prefix: str, depth: int) -> None:
        sec = [prefix + eml_label(mail)]
        b.add("heading", f"{prefix}제목", mail["subject"] or "(제목 없음)", None, None, None, section=sec)
        head = "; ".join(x for x in (f"보낸 사람: {mail['from_name']} <{mail['from_addr']}>" if mail["from_addr"] else "",
                                     f"받는 사람: {mail['to']}" if mail["to"] else "", f"참조: {mail['cc']}" if mail["cc"] else "",
                                     f"날짜: {mail['date']}" if mail["date"] else "") if x)
        if head:
            b.add("paragraph", f"{prefix}헤더", head, None, None, None, section=sec)
        for n, para in enumerate([p.strip() for p in re.split(r"\n\s*\n", mail["body"]) if p.strip()], 1):
            b.add("paragraph", f"{prefix}본문 {n}", " ".join(para.split()), None, None, None, section=sec)
        if mail["attachments"]:
            b.add("paragraph", f"{prefix}첨부", "첨부: " + ", ".join(mail["attachments"]), None, None, None, section=sec)
        if depth < EML_FORWARD_DEPTH:
            for sub in mail["forwarded"]:
                try:
                    add_mail(_eml_message(sub), "전달된 메일 ", depth + 1)
                except Exception:                    # noqa: BLE001  전달된 메일 하나를 못 읽어도 바깥 메일은 그대로
                    continue

    add_mail(top, "", 0)
    meta = {"mail": {"subject": top["subject"], "from": f"{top['from_name']} <{top['from_addr']}>" if top["from_addr"] else top["from_name"],
                     "to": top["to"], "cc": top["cc"], "date": top["date"], "attachments": top["attachments"][:MAX_LIST_ITEMS],
                     "attachment_count": len(top["attachments"]), "forwarded": len(top["forwarded"])}}
    return b.units, meta


PARSERS = {"pdf": parse_pdf, "hwp": parse_hwp, "hwpx": parse_hwpx, "docx": parse_docx,
           "pptx": parse_pptx, "xlsx": parse_xlsx, "csv": parse_csv, "txt": parse_txt, "eml": parse_eml}


def file_type(path: str) -> str:
    ext = os.path.splitext(path)[1].lower()
    if ext in LEGACY_OFFICE:
        raise rt.InputError(f"구형 Office 형식({ext})은 라이브러리로 읽을 수 없습니다. "
                            f"excel.py / word.py / ppt.py read --target \"{path}\"로 Office에서 읽으세요.")
    if ext not in SUPPORTED:
        raise rt.InputError(f"지원하지 않는 형식입니다: {ext or '(확장자 없음)'} (지원: {', '.join(sorted(SUPPORTED))})")
    return SUPPORTED[ext]


# ==============================================================================
# 3. 캐시와 색인
# ==============================================================================

def _doc_cache_path(path: str) -> str:
    key = hashlib.sha1(os.path.normcase(os.path.abspath(path)).encode("utf-8")).hexdigest()[:16]
    folder = os.path.join(rt.get_cache_dir(), "docs")
    os.makedirs(folder, exist_ok=True)
    return os.path.join(folder, f"{key}.json")


def _index_path() -> str:
    return os.path.join(rt.get_cache_dir(), "index.json")


class IndexReadError(rt.InputError):
    """색인 파일(index.json)이 있는데 읽지 못함 (다른 프로그램이 잡고 있음 등). 이때 색인을 덮어쓰지 않습니다."""


def _read_bytes(path: str) -> bytes:
    with open(path, "rb") as f:
        return f.read()


def load_index() -> Dict[str, Dict[str, Any]]:
    """색인 {원본 경로: 정보}. Windows에서 대소문자만 다른 같은 파일(C:\\Docs, c:\\docs)은 최근 것 하나로 합칩니다.
    - 일시적인 잠금(백신, 공유 위반)은 retry_fs로 기다려 다시 읽고, 끝내 못 읽으면 IndexReadError
      (빈 색인으로 보고 저장하면 등록한 문서가 색인에서 모두 빠지던 문제)
    - 내용이 손상(JSON/인코딩 오류)된 경우에만 캐시 문서들로 색인을 다시 만듭니다 (JIT 복구, recovered로 알림)"""
    p = _index_path()
    if not os.path.exists(p):
        return {}
    try:
        blob = rt.retry_fs(lambda: _read_bytes(p))
    except FileNotFoundError:
        return {}
    except OSError as e:
        raise IndexReadError(f"색인 파일을 읽지 못했습니다 ({type(e).__name__}: 다른 프로그램이 사용 중일 수 있음). "
                             "색인은 바꾸지 않았습니다. 사용자에게 알리고, 잠시 뒤 다시 요청하게 하세요.")
    try:
        # 도구가 쓰는 색인은 UTF-8. 디코딩 실패도 손상으로 보고 다시 만듦 (read_text_file은 InputError를 내
        # clear --all까지 모든 reader 명령이 실패했음)
        data = json.loads(blob.decode("utf-8-sig"))
        if not isinstance(data, dict):
            raise ValueError("색인이 객체가 아님")
    except (ValueError, UnicodeDecodeError):
        return _rebuild_index()
    merged: Dict[str, Tuple[str, Dict[str, Any]]] = {}
    for path, entry in data.items():
        if not isinstance(entry, dict):
            continue
        key = os.path.normcase(os.path.abspath(path))
        if key not in merged or str(entry.get("parsed_at", "")) > str(merged[key][1].get("parsed_at", "")):
            merged[key] = (path, entry)
    return {path: entry for path, entry in merged.values()}


def _rebuild_index() -> Dict[str, Dict[str, Any]]:
    """손상된 색인 대신 캐시 문서 파일(docs/*.json)로 색인을 다시 만들고 저장합니다 (index --all과 같은 정보).
    저장에 실패해도 다시 만든 색인을 돌려줍니다 (다음 쓰기 때 저장)."""
    cache_dir = os.path.join(rt.get_cache_dir(), "docs")
    index: Dict[str, Dict[str, Any]] = {}
    names = sorted(os.listdir(cache_dir)) if os.path.isdir(cache_dir) else []
    for name in names:
        if not name.endswith(".json"):
            continue
        cache_file = os.path.join(cache_dir, name)
        try:
            doc = json.loads(rt.read_text_file(cache_file))
            path = doc.get("file")
            if (isinstance(path, str) and isinstance(doc.get("units"), list)
                    and os.path.basename(_doc_cache_path(path)) == name):
                _index_put(index, path, _index_entry(doc, cache_file))
        except Exception:
            continue                       # 손상된 캐시 파일은 clear --all이 정리
    try:
        _save_index(index)
    except OSError as e:
        rt.log("WARN", f"다시 만든 색인을 저장하지 못했습니다 (다음 등록 때 저장): {e}")
    rt.note_recovered(f"색인 파일 손상 -> 캐시 문서 {len(index)}개로 색인을 다시 만듦")
    return index


def _save_index(index: Dict[str, Dict[str, Any]]) -> None:
    rt.safe_write_text(_index_path(), json.dumps(index, ensure_ascii=False))


def _remove_stale_lock(path: str) -> bool:
    """오래된 잠금을 고유한 이름으로 옮긴 뒤 지웁니다. 두 실행이 동시에 오래된 잠금을 보고 지우면, 늦은 쪽이
    먼저 지운 쪽이 새로 만든 잠금까지 지우던 경쟁을 막습니다: 옮긴 파일이 새 잠금이면 제자리로 돌려놓습니다."""
    moved = f"{path}.stale.{os.getpid()}"
    try:
        os.replace(path, moved)
    except OSError:
        return False                       # 다른 실행이 먼저 옮김
    try:
        if time.time() - os.path.getmtime(moved) <= INDEX_LOCK_STALE and not os.path.exists(path):
            os.replace(moved, path)        # 그 사이 다른 실행이 새로 만든 잠금이었음: 되돌려 놓음
            return False
        os.remove(moved)
    except OSError:
        pass
    return True


@contextlib.contextmanager
def _index_lock():
    """색인 읽기-고치기-쓰기를 한 실행씩만 하게 잠급니다 (동시에 index를 두 번 실행하면 항목이 사라지던 문제).
    잠금 파일을 O_CREAT|O_EXCL로 만들고(Windows·Linux 공통), 최대 INDEX_LOCK_WAIT초 기다립니다.
    INDEX_LOCK_STALE초보다 오래된 잠금은 중단된 실행이 남긴 것으로 보고 지웁니다.
    끝내 얻지 못하면 잠금 없이 진행하고 경고만 남깁니다 (index는 끝에 색인을 다시 읽어 빠진 항목을 채움)."""
    path = os.path.join(rt.get_cache_dir(), "index.lock")
    os.makedirs(os.path.dirname(path), exist_ok=True)
    deadline = time.monotonic() + INDEX_LOCK_WAIT
    acquired = False
    while True:
        try:
            fd = os.open(path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            try:
                os.write(fd, str(os.getpid()).encode("ascii"))
            finally:
                os.close(fd)
            acquired = True
            break
        except (FileExistsError, PermissionError):       # Windows: 지우는 중인 파일은 PermissionError
            try:
                if time.time() - os.path.getmtime(path) > INDEX_LOCK_STALE and _remove_stale_lock(path):
                    rt.log("WARN", f"{int(INDEX_LOCK_STALE)}초보다 오래된 색인 잠금(중단된 실행)을 지웠습니다.")
                    continue
            except OSError:
                pass
        if time.monotonic() >= deadline:
            rt.log("WARN", f"색인 잠금을 {int(INDEX_LOCK_WAIT)}초 안에 얻지 못해 잠금 없이 진행합니다.")
            break
        time.sleep(0.05)
    try:
        yield
    finally:
        if acquired:
            try:
                os.remove(path)
            except OSError:
                pass


def _index_update(change) -> Dict[str, Dict[str, Any]]:
    """잠근 상태에서 색인을 읽고 change(index)로 고친 뒤 저장합니다. 고친 색인을 돌려줍니다.
    색인을 읽지 못하면(IndexReadError) 저장하지 않고 오류를 올립니다 (읽지 못한 색인 위에 덮어쓰지 않음)."""
    with _index_lock():
        index = load_index()
        change(index)
        _save_index(index)
    return index


def _index_put(index: Dict[str, Dict[str, Any]], path: str, entry: Dict[str, Any]) -> None:
    key = os.path.normcase(path)
    for old in [k for k in index if os.path.normcase(k) == key]:
        del index[old]
    index[path] = entry


def _index_entry(doc: Dict[str, Any], cache_file: str) -> Dict[str, Any]:
    return {"type": doc["type"], "units": len(doc["units"]), "parsed_at": doc.get("parsed_at", ""),
            "chars": sum(len(u["text"]) for u in doc["units"]), "cache_file": os.path.basename(cache_file),
            "mtime": doc.get("mtime"), "size": doc.get("size"), "parser_version": doc.get("parser_version"),
            "pages": _doc_pages(doc), "sections": _doc_sections(doc)}


def _ensure_indexed(path: str, doc: Dict[str, Any], cache_file: str) -> None:
    """캐시는 최신인데 색인에 없거나 다르면(색인 저장 전 중단, 동시 실행) 색인을 고칩니다 (JIT 복구, 실패는 무시)."""
    try:
        index = load_index()
        entry = next((e for k, e in index.items() if os.path.normcase(k) == os.path.normcase(path)), None)
        if entry and all(entry.get(k) == doc.get(k) for k in ("mtime", "size", "parser_version")):
            return
        _index_update(lambda idx: _index_put(idx, path, _index_entry(doc, cache_file)))
        rt.log("WARN", f"캐시는 있는데 색인에서 빠진 문서를 색인에 다시 넣었습니다: {os.path.basename(path)}")
    except Exception:
        pass


def is_stale_doc(doc: Dict[str, Any], st: Optional[os.stat_result] = None) -> bool:
    """캐시가 이전 파서 버전이거나, 원본 파일이 캐시한 뒤 바뀌었으면 True (원본이 없으면 파서 버전만 봄)."""
    if doc.get("parser_version") != PARSER_VERSION:
        return True
    if st is None:
        try:
            st = os.stat(doc.get("file", ""))
        except OSError:
            return False
    return doc.get("mtime") != st.st_mtime or doc.get("size") != st.st_size


def _doc_pages(doc: Dict[str, Any]) -> Optional[int]:
    meta = doc.get("meta") or {}
    if isinstance(meta.get("pages"), int):
        return meta["pages"]
    pages = [u.get("page") for u in doc.get("units", []) if isinstance(u.get("page"), int)]
    return max(pages) if pages else None


def _doc_sections(doc: Dict[str, Any]) -> int:
    units = doc.get("units", [])
    heads = sum(1 for u in units if u.get("kind") == "heading")
    if heads:
        return heads
    return len({tuple(u.get("section") or []) for u in units if u.get("section")})


def load_cached_only(path: str) -> Optional[Dict[str, Any]]:
    cache_file = _doc_cache_path(path)
    if not os.path.exists(cache_file):
        return None
    try:
        doc = json.loads(rt.read_text_file(cache_file))
        return doc if isinstance(doc, dict) and isinstance(doc.get("units"), list) else None
    except Exception:
        return None


def _uncomputed_from_check(meta: Dict[str, Any], uncomputed: Dict[str, Any]) -> None:
    """xlsx_uncomputed_formulas 결과를 문서 meta에 (계산되지 않은 칸이 있거나 확인하지 못했을 때만)."""
    meta.pop("uncomputed_formulas", None)
    if uncomputed.get("unknown") or uncomputed["count"]:
        meta["uncomputed_formulas"] = uncomputed


RECHECK_MAX = 3                  # 한 실행에서 수식 칸을 다시 확인하는 문서 수 상한 (큰 엑셀이 많으면 검색 한 번이 시간 제한을 넘던 문제)
_RECHECK_RUN: Dict[str, Any] = {"run": None, "done": 0}


def _recheck_allowed() -> bool:
    """이번 실행에서 수식 칸 다시 확인을 더 해도 되는지 (상한 RECHECK_MAX, 시간 제한에서 여유를 뺀 시각 전)."""
    if _RECHECK_RUN["run"] != rt.RUN_START:              # 실행마다 새로 셈 (테스트처럼 한 프로세스에서 여러 번 실행해도)
        _RECHECK_RUN.update(run=rt.RUN_START, done=0)
    if _RECHECK_RUN["done"] >= RECHECK_MAX:
        return False
    limit = rt.time_limit()
    if limit != float("inf") and time.monotonic() - rt.RUN_START > limit - rt.TIME_RESERVE:
        return False
    _RECHECK_RUN["done"] += 1
    return True


def recheck_uncomputed(doc: Dict[str, Any], cache_file: Optional[str] = None) -> Dict[str, Any]:
    """캐시의 '수식 칸 확인 실패'(MemoryError 등 일시적 실패)는 결과로 믿지 않고 캐시를 쓸 때마다 다시 확인합니다.
    확인되면 meta를 고치고 캐시 파일도 고쳐 씁니다(실패는 무시). 다시 실패하면 그대로 (UNVERIFIED 유지).
    한 실행에서 RECHECK_MAX개까지, 시간 제한 여유 안에서만 (넘으면 확인하지 않은 채 UNVERIFIED 유지, 다음 실행에서 다시)."""
    meta = doc.get("meta") or {}
    info = meta.get("uncomputed_formulas")
    path = doc.get("file", "")
    if not (isinstance(info, dict) and info.get("unknown")) or doc.get("type") != "xlsx" or not os.path.isfile(path):
        return doc
    if not _recheck_allowed():
        rt.debug(f"수식 칸 다시 확인 상한/시간 초과로 건너뜀: {os.path.basename(path)}")
        return doc
    again = xf.xlsx_uncomputed_formulas(path)
    if again.get("unknown"):
        return doc
    _uncomputed_from_check(meta, again)
    doc["meta"] = meta
    rt.log("INFO", f"전에 확인하지 못한 수식 칸을 다시 확인했습니다: {os.path.basename(path)}")
    try:
        rt.safe_write_text(cache_file or _doc_cache_path(path), json.dumps(doc, ensure_ascii=False))
    except OSError as e:
        rt.debug(f"캐시 고쳐 쓰기 실패 (결과에는 영향 없음): {e}")
    return doc


def load_document_status(path: str, use_cache: bool = True, check_index: bool = True) -> Tuple[Dict[str, Any], str]:
    """파일을 캐시에서 불러오거나(수정 시각/크기/파서 버전이 같으면) 파싱해 캐시와 색인에 등록합니다.
    반환: (문서, 상태) 상태는 'cached'(캐시 사용), 'new'(처음 등록), 'updated'(바뀌어서 다시 등록)."""
    path = os.path.abspath(path)
    ftype = file_type(path)
    st = os.stat(path)
    cache_file = _doc_cache_path(path)
    had_cache = False
    if use_cache and os.path.exists(cache_file):
        had_cache = True
        try:
            cached = json.loads(rt.read_text_file(cache_file))
            if not is_stale_doc(cached, st):
                rt.debug(f"캐시 사용: {os.path.basename(path)}")
                if check_index:
                    _ensure_indexed(path, cached, cache_file)
                return recheck_uncomputed(cached, cache_file), "cached"
        except Exception:
            rt.log("WARN", f"캐시가 손상되어 다시 읽습니다: {os.path.basename(path)}")
    rt.log("PROGRESS", f"문서 읽는 중: {os.path.basename(path)}")
    try:
        units, meta = PARSERS[ftype](path)
    except rt.InputError:
        raise
    except _DAMAGE_ERRORS as e:
        raise _damaged_error(path, e) from e
    except Exception as e:
        if type(e).__name__ in _DAMAGE_NAMES:
            raise _damaged_error(path, e) from e
        raise
    doc = {"file": path, "type": ftype, "mtime": st.st_mtime, "size": st.st_size,
           "parser_version": PARSER_VERSION, "parsed_at": rt.timestamp(), "meta": meta, "units": units}
    try:
        rt.safe_write_text(cache_file, json.dumps(doc, ensure_ascii=False))
    except OSError as e:
        doc["_cache_failed"] = True          # index가 '캐시에 저장 확인'이라고 하지 않게
        rt.log("WARN", f"캐시 저장 실패 (결과에는 영향 없음): {e}")
        return doc, ("updated" if had_cache else "new")
    try:
        _index_update(lambda idx: _index_put(idx, path, _index_entry(doc, cache_file)))
    except (OSError, IndexReadError) as e:
        # 캐시는 저장됨: 색인은 다음 index(--all)가 캐시 파일로 채움. 읽지 못한 색인은 덮어쓰지 않음
        rt.log("WARN", f"색인에 넣지 못했습니다 (결과에는 영향 없음, 색인은 그대로): {e}")
    return doc, ("updated" if had_cache else "new")


def load_document(path: str, use_cache: bool = True) -> Dict[str, Any]:
    return load_document_status(path, use_cache)[0]


def _under(path: str, folder: str) -> bool:
    p = os.path.normcase(os.path.abspath(path))
    f = os.path.normcase(os.path.abspath(folder)).rstrip("\\/") + os.sep
    return p.startswith(f)


def _collect_folder(folder: str, depth: int = MAX_FOLDER_DEPTH,
                    excluded: Optional[List[str]] = None) -> Tuple[List[str], Dict[str, int]]:
    """폴더의 지원 형식 파일 목록과 건너뛴 형식별 개수 (예: {'.png': 2}). depth 0 = 그 폴더만.
    excluded를 주면 들어가지 않은 하위 폴더(깊이 제한, '_'/'.'로 시작)를 거기에 담습니다."""
    files: List[str] = []
    skipped: Counter = Counter()
    base_depth = folder.rstrip("\\/").count(os.sep)
    for root, dirs, names in os.walk(folder):
        if root.count(os.sep) - base_depth >= depth:
            if excluded is not None:
                excluded += [os.path.join(root, d) for d in dirs]
            dirs[:] = []
        if excluded is not None:
            excluded += [os.path.join(root, d) for d in dirs if d.startswith(("_", "~$"))]
        dirs[:] = sorted(d for d in dirs if not d.startswith((".", "~$", "_")))
        for n in sorted(names):
            if n.startswith(("~$", ".")):
                continue
            ext = os.path.splitext(n)[1].lower()
            if ext not in SUPPORTED:
                skipped[ext or "(확장자 없음)"] += 1
                continue
            files.append(os.path.join(root, n))
            if len(files) >= MAX_FOLDER_SCAN:
                rt.log("WARN", f"폴더에서 최대 {MAX_FOLDER_SCAN}개 파일까지만 살펴봅니다.")
                return files, dict(skipped)
    return files, dict(skipped)


def _folder_arg(raw: str) -> str:
    folder = os.path.abspath(os.path.expandvars(os.path.expanduser(str(raw).strip().strip('"'))))
    if not os.path.isdir(folder):
        raise rt.InputError(f"폴더를 찾을 수 없습니다: {folder}")
    return folder


# 읽지 못한 파일 기록 {정규화한 경로: {mtime, size, parser_version, reason}}. 같은 파일(수정 시각·크기·파서 버전이 같음)은
# 다음 실행에서 다시 읽지 않고 failures로만 보고합니다. 실패하는 느린 파일을 매 실행 다시 읽어 CONTINUE가 끝나지 않던 문제 (R10).
FAILED_FILE = "index_failures.json"


def _failed_path() -> str:
    return os.path.join(rt.get_cache_dir(), FAILED_FILE)


def load_failed() -> Dict[str, Dict[str, Any]]:
    try:
        data = json.loads(rt.read_text_file(_failed_path()))
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}                            # 없거나 손상: 빈 기록 (다음 실패 때 새로 씀)


def _save_failed(data: Dict[str, Dict[str, Any]]) -> None:
    try:
        rt.safe_write_text(_failed_path(), json.dumps(data, ensure_ascii=False))
    except Exception as e:
        rt.log("WARN", f"읽지 못한 파일 기록을 저장하지 못했습니다: {type(e).__name__}")


def known_failure(path: str, st: Optional[os.stat_result] = None) -> Optional[str]:
    """이 파일이 지금 모습 그대로 전에 읽기에 실패했으면 그 사유, 아니면 None."""
    rec = load_failed().get(os.path.normcase(os.path.abspath(path)))
    if not isinstance(rec, dict):
        return None
    try:
        st = st or os.stat(path)
    except OSError:
        return None
    if (rec.get("mtime") == st.st_mtime and rec.get("size") == st.st_size
            and rec.get("parser_version") == PARSER_VERSION):
        return str(rec.get("reason") or "읽기 실패")
    return None


def record_failure(path: str, reason: str) -> None:
    try:
        st = os.stat(path)
    except OSError:
        return
    data = load_failed()
    data[os.path.normcase(os.path.abspath(path))] = {"mtime": st.st_mtime, "size": st.st_size,
                                                     "parser_version": PARSER_VERSION, "reason": reason[:300]}
    _save_failed(data)


def clear_failure(path: str) -> None:
    key = os.path.normcase(os.path.abspath(path))
    data = load_failed()
    if key in data:
        del data[key]
        _save_failed(data)


KNOWN_FAILURE_NOTE = "이전 실행에서 읽지 못함 (파일이 바뀌면 다시 시도. 바로 다시 해 보려면 index --file로 이 파일 하나만): "


def _index_files(files: List[str], max_files: int, retry_failed: bool = False) -> Dict[str, Any]:
    """파일들을 캐시에 등록합니다. 새로 파싱하는 파일 수만 max_files로 제한합니다 (이미 최신인 파일은 세지 않음).
    전에 읽지 못한 파일(그대로인 것)은 retry_failed가 아니면 다시 읽지 않고 failures에만 넣습니다.
    반환: docs(불러온 문서), new/updated/cached(파일 이름 목록), failures, remaining.
    failures 항목: file, reason, error(원래 예외. 파일 하나만 줬을 때 _raise_single_failure가 그대로 다시 냄, JSON에는 넣지 않음)."""
    out: Dict[str, Any] = {"docs": [], "new": [], "updated": [], "cached": [], "failures": [], "remaining": 0}
    parsed = 0                                   # 이번 실행에서 새로 읽어 본 파일 수 (실패 포함)
    registered = 0                               # 그중 등록에 성공한 수
    failed_now = 0
    spent = 0.0                                  # 이번 실행에서 새로 읽는 데 쓴 초 (다음 파일 예상 시간)
    index = load_index()
    total = len(files)
    for i, path in enumerate(files, 1):
        try:
            st = os.stat(path)
        except OSError as e:
            out["failures"].append({"file": path, "reason": rt.error_payload(e)["message"][:300], "error": e})
            continue
        key = os.path.normcase(os.path.abspath(path))
        entry = next((e for k, e in index.items() if os.path.normcase(k) == key), None)
        fresh = bool(entry and entry.get("parser_version") == PARSER_VERSION
                     and entry.get("mtime") == st.st_mtime and entry.get("size") == st.st_size)
        if not fresh and not retry_failed:
            known = known_failure(path, st)
            if known is not None:
                out["failures"].append({"file": path, "reason": (KNOWN_FAILURE_NOTE + known)[:300],
                                        "error": rt.InputError(known)})
                continue
        if not fresh and parsed >= max_files:
            out["remaining"] += 1
            continue
        if not fresh and parsed:
            # R10: 새 파일을 읽기 전에 시계를 봄 (이번 실행에서 하나는 읽은 뒤에만: 반복하면 반드시 끝남).
            # 읽은 파일은 캐시에, 실패한 파일은 실패 기록에 저장되어 같은 명령을 다시 실행하면 건너뜀
            rt.check_time(spent / parsed, f"문서 {total}개 중 {i - 1}개 확인, 이번에 새로 등록 {registered}개"
                          + (f", 읽지 못함 {failed_now}개" if failed_now else "") + " (남은 파일은 같은 명령으로 이어서)")
        t0 = time.monotonic()
        try:
            if not fresh:
                rt.log("PROGRESS", f"[{i}/{total}] 등록 중: {os.path.basename(path)}")
            doc, status = load_document_status(path, check_index=not fresh)
        except Exception as e:
            # 내부 오류는 'KeyError: ...'처럼 종류와 함께 (str(e)만 쓰면 "'heading'"만 남음)
            reason = rt.error_payload(e)["message"][:300]
            out["failures"].append({"file": path, "reason": reason, "error": e})
            if not fresh:
                parsed += 1
                failed_now += 1
                spent += time.monotonic() - t0
                record_failure(path, reason)
            continue
        if status != "cached":
            parsed += 1
            registered += 1
            spent += time.monotonic() - t0
            clear_failure(path)
        out["docs"].append(doc)
        out[status].append(path)
    return out


def _raise_single_failure(res: Dict[str, Any]) -> None:
    """--file로 파일 하나만 줬는데 읽지 못했으면 원래 예외를 그대로 다시 냅니다 (read와 같은 결과).
    사유 글로 InputError를 새로 만들면 내부 오류(파서 버그)도 STOP_RETRY 입력 오류가 되고 traceback이 남지 않으며,
    InputError의 log_message(문서 글을 뺀 글)도 잃기 때문. 내부 오류는 ERROR와 traceback, InputError는 그대로."""
    if res["failures"]:
        raise res["failures"][0]["error"]


# ==============================================================================
# 4. 공통 출력 도우미
# ==============================================================================

def parse_pages(spec: Optional[str], total: int, option: str = "--pages") -> List[int]:
    """'3', '1-3', '1,4,6-8' -> 1부터 센 번호 목록. option은 오류 문구에 쓸 인자 이름 (ppt.py는 --slides)."""
    if not spec:
        return list(range(1, total + 1))
    result: List[int] = []
    for part in str(spec).replace(" ", "").split(","):
        if not part:
            continue
        m = re.fullmatch(r"(\d+)(?:-(\d+))?", part)
        if not m:
            raise rt.InputError(f"{option} 형식이 잘못되었습니다: '{spec}' (예: 3, 1-3, 1,4,6-8)")
        a = int(m.group(1))
        b = int(m.group(2) or a)
        lo, hi = max(1, min(a, b)), min(total, max(a, b))    # 먼저 줄임 (1-999999999를 다 돌지 않게)
        result += range(lo, hi + 1)
    if not result:
        raise rt.InputError(f"{option} '{spec}'에 해당하는 항목이 없습니다 (전체 {total}개).")
    return sorted(set(result))


def page_set(spec: Optional[str]) -> Optional[set]:
    """검색 --pages: '3-7', '2,5' -> 쪽 번호 집합 (상한 없음)."""
    if not spec:
        return None
    pages: set = set()
    for part in str(spec).replace(" ", "").split(","):
        if not part:
            continue
        m = re.fullmatch(r"(\d+)(?:-(\d+))?", part)
        if not m:
            raise rt.InputError(f"--pages 형식이 잘못되었습니다: '{spec}' (예: 3, 3-7, 1,4,6-8)")
        a, b = int(m.group(1)), int(m.group(2) or m.group(1))
        if abs(b - a) > 100000:
            raise rt.InputError(f"--pages 범위가 너무 넓습니다: '{spec}'")
        pages.update(range(min(a, b), max(a, b) + 1))
    return pages


def section_text(unit: Dict[str, Any]) -> str:
    return " > ".join(unit.get("section") or [])


def section_key(spec: Optional[str]) -> str:
    """--section 값을 비교용으로: 공백을 하나로, 소문자 (read/search 공통)."""
    return re.sub(r"\s+", " ", spec or "").strip().lower()


_SECTION_NUMBER = re.compile(r"(?:[a-z]\.)?\d+(?:\.\d+)*\.?")


def _section_number_pattern(key: str) -> "re.Pattern[str]":
    """장절 번호 '2.1'은 번호 경계를 지킴: 2.1, 2.1.3(하위 절), 제2.1절은 맞고 2.10, 12.1, 1.2.1은 맞지 않음."""
    num = re.escape(key.rstrip("."))
    return re.compile(rf"(?:(?<![\w.])|(?<=제)){num}(?!\d)")


def section_matches(unit: Dict[str, Any], key: str) -> bool:
    """장절 경로(section_text)에 --section 값이 들어 있는지 (대소문자·공백 무시). key는 section_key 결과.
    장절 번호로 보이는 값(2.1, 3, A.1)은 번호 경계를 지키고 하위 절을 포함, 제목 글자는 부분 일치."""
    text = re.sub(r"\s+", " ", section_text(unit)).lower()
    if _SECTION_NUMBER.fullmatch(key):
        return bool(_section_number_pattern(key).search(text))
    return key in text


def where_text(unit: Dict[str, Any]) -> str:
    """근거의 위치 부분: PDF/저장된 Word 쪽은 'p.N', 쪽 나눔만 센 Word 쪽은 'p.N(추정)', 그 밖은 loc."""
    page, src = unit.get("page"), unit.get("page_source")
    if page is not None and src in ("pdf", "word_saved_layout"):
        return f"p.{page}"
    if page is not None and src == "page_break":
        return f"p.{page}(추정)"
    return str(unit.get("loc", ""))


_SAME_NAMES: set = set()     # 이번 검색에서 이름이 같은 파일 (다른 폴더): 근거에 상위 폴더 이름을 붙임


def evidence_name(file: str) -> str:
    """근거에 쓸 파일 이름. 다른 폴더의 같은 이름 파일이 함께 검색되면 '상위 폴더\\이름' (구분되지 않던 것, 2026-11-09 검토 B13)."""
    name = os.path.basename(file)
    if name.lower() in _SAME_NAMES:
        parent = os.path.basename(os.path.dirname(os.path.abspath(file)))
        return f"{parent}\\{name}" if parent else name
    return name


def note_same_names(files: List[str]) -> None:
    """검색할 문서 경로들 중 이름이 겹치는 것을 기억 (evidence_name이 씀)."""
    from collections import Counter as _C
    _SAME_NAMES.clear()
    counts = _C(os.path.basename(f).lower() for f in set(os.path.abspath(x) for x in files))
    _SAME_NAMES.update(n for n, k in counts.items() if k > 1)


def evidence_text(file: str, unit: Dict[str, Any]) -> str:
    """'규격서.pdf p.12, 1. SCOPE > 1.1. Scope (1.1)' 형식의 근거 문구."""
    s = f"{evidence_name(file)} {where_text(unit)}".strip()
    sec = section_text(unit)
    if sec:
        s += f", {sec}"
    if unit.get("number"):
        s += f" ({unit['number']})"
    return s


def units_to_markdown(title: str, units: List[Dict[str, Any]], file: Optional[str] = None) -> str:
    """단위 목록을 마크다운으로. reader 단위(section 키 있음)는 '근거:' 줄을 붙입니다 (word.py/ppt.py 단위는 그대로)."""
    lines = [f"# {title}", ""]
    for u in units:
        head = u["loc"] + (f" - {u['heading']}" if u.get("heading") and u["heading"] != u["text"] else "")
        lines += [f"## {head}", ""]
        if file and "section" in u:
            lines += [f"근거: {evidence_text(file, u)}", ""]
        lines += [u["text"], ""]
    return "\n".join(lines)


def _write_confirm_commands(output_file: str) -> List[Dict[str, Any]]:
    """저장한 파일을 열린 문서에 넣는 명령 (문서를 바꿈: next_commands가 아니라 confirm_commands, 1.4.75)."""
    low = output_file.lower()
    if low.endswith(".xlsx"):
        return []
    if low.endswith(".csv"):
        return [rt.next_command("excel.py", ["write", "--target", "active", "--file", output_file],
                                "저장한 표를 열린 엑셀에 넣기 (사용자가 원할 때만)")]
    return [rt.next_command("word.py", ["write", "--target", "active", "--file", output_file],
                            "저장한 내용을 열린 워드 문서 끝에 넣기 (사용자가 원할 때만)"),
            rt.next_command("ppt.py", ["write", "--target", "active", "--file", output_file],
                            "저장한 내용을 열린 PPT에 새 슬라이드로 넣기 (사용자가 원할 때만)")]


def _save_ext(save: Any) -> str:
    if save is True or not save:
        return ".md"
    ext = os.path.splitext(str(save).strip().strip('"'))[1].lower()
    return ext if ext in (".xlsx", ".csv", ".md") else ".md"


XLSX_CELL_MAX = 32000      # 엑셀 셀 글자 상한(32767)보다 조금 작게


def _xlsx_safe(value: Any) -> Any:
    if not isinstance(value, str):
        return value
    from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
    return ILLEGAL_CHARACTERS_RE.sub("", value)[:XLSX_CELL_MAX]


def _xlsx_cut_text(rows: int) -> str:
    return f"셀 글자 상한 {XLSX_CELL_MAX}자 초과: {rows}행 잘림 (.md로 저장하면 전체)"


def write_xlsx_table(path: str, header: List[str], rows: List[List[Any]], widths: List[int],
                     sheet: str = "결과") -> Tuple[str, int]:
    """표를 xlsx로 저장합니다 (머리글 굵게, 줄바꿈, 열 너비, 첫 행 고정). 저장 후 다시 열어 행 수를 확인합니다.
    반환: (저장 경로, 셀 글자 상한에서 글이 잘린 행 수). 잘린 행이 있으면 호출하는 쪽이 UNVERIFIED로 알립니다."""
    openpyxl = rt.require_lib("openpyxl", "openpyxl", "xlsx로 저장(.csv나 .md로도 저장할 수 있음)")
    from openpyxl.styles import Alignment, Font

    def _write(p: str) -> None:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = sheet
        ws.append(header)
        for row in rows:
            ws.append([_xlsx_safe(v) for v in row])
        wrap = Alignment(wrap_text=True, vertical="top")
        for row in ws.iter_rows():
            for cell in row:
                if cell.data_type == "f":           # 문서 글이 '='로 시작해도 수식이 아닌 글로
                    cell.data_type = "s"
                cell.alignment = wrap
        for cell in ws[1]:
            cell.font = Font(bold=True)
        for i, w in enumerate(widths):
            ws.column_dimensions[openpyxl.utils.get_column_letter(i + 1)].width = w
        ws.freeze_panes = "A2"
        wb.save(p)

    out = rt.save_with_lock_fallback(path, _write, "엑셀 파일")
    wb = rt.retry_fs(lambda: openpyxl.load_workbook(out, read_only=True))
    try:
        n = sum(1 for _ in wb.active.iter_rows(values_only=True))
    finally:
        wb.close()
    if n != len(rows) + 1:
        raise RuntimeError(f"저장한 xlsx의 행 수가 다릅니다 (기대 {len(rows) + 1}, 실제 {n}): {out}")
    cut_rows = sum(1 for row in rows if any(isinstance(v, str) and len(v) > XLSX_CELL_MAX for v in row))
    return out, cut_rows


def _unit_view(n: int, u: Dict[str, Any], text: str) -> Dict[str, Any]:
    return {"no": n, "loc": u["loc"], "kind": u.get("kind"), "page": u.get("page"), "section": u.get("section") or [],
            "number": u.get("number"), "heading": u.get("heading", ""), "text": text}


def _section_not_found(spec: str, units: List[Dict[str, Any]], pages: Optional[str]) -> rt.InputError:
    """--section에 맞는 단위가 없음: 가까운 장절 이름(없으면 문서의 장절 일부)을 [후보]로 (R8).
    장절 이름은 문서 글이라 오류 로그에는 개수만 남김 (rt.ambiguous)."""
    names: List[str] = []
    for u in units:
        for t in u.get("section") or []:
            if t and t not in names:
                names.append(t)
    where = f" (--pages {pages} 안)" if pages else ""
    head = f"장절 '{spec}'에 맞는 내용이 없습니다{where}."
    if not names:
        return rt.InputError(head + " 이 문서에는 장절 정보가 없습니다. --section 없이 --pages로 읽으세요.")
    near = rt.suggest(spec, names, n=5)
    if near:
        tail = " 비슷한 장절이 있지만 같은 것인지 도구가 정할 수 없습니다. 사용자에게 고르게 하세요."
        return rt.ambiguous(head + tail, near, replace=spec, option="--section")
    return rt.ambiguous(head + " 비슷한 장절이 없습니다. 이 문서의 장절 중에서 사용자에게 고르게 하세요.", names, replace=spec,
                        option="--section")


# ==============================================================================
# 5. 명령: read
# ==============================================================================

def cmd_read(args) -> Dict[str, Any]:
    path = rt.resolve_input_path(args.file, "문서 파일")
    doc = load_document(path)
    units = doc["units"]
    total = len(units)
    name = os.path.basename(path)
    unit_name = UNIT_NAMES.get(doc["type"], "항목")
    # 쪽 정보가 있는 형식(docx/hwpx의 저장된 쪽 경계 포함)은 --pages가 실제 쪽: search의 근거 'p.12'를 그대로 읽을 수 있게
    by_page = doc["type"] in PAGE_TYPES or any(u.get("page") for u in units)
    req_pages: Optional[set] = None
    if args.units:
        numbers = parse_pages(args.units, total) if total else []
    elif by_page and args.pages:
        req_pages = set(parse_pages(args.pages, _doc_pages(doc) or 0)) if total else set()
        numbers = [i for i, u in enumerate(units, 1) if u.get("page") in req_pages]
    else:
        numbers = parse_pages(args.pages, total) if total else []
    sec_key = section_key(args.section)
    if sec_key and not (doc.get("meta") or {}).get("ocr_needed"):     # 스캔 PDF는 아래 OCR 안내로
        matched = [n for n in numbers if section_matches(units[n - 1], sec_key)]
        if not matched:
            pool = [units[n - 1] for n in numbers] if (args.pages or args.units) else units
            raise _section_not_found(args.section, pool if any(u.get("section") for u in pool) else units,
                                     args.pages or args.units)
        numbers = matched
    selected = [units[n - 1] for n in numbers]
    result: Dict[str, Any] = {"file": path, "type": doc["type"], "units_total": total, "count": len(selected),
                              "unit_name": unit_name, "pages_total": _doc_pages(doc),
                              "pages_mode": "page" if by_page else "item",
                              "sections_total": _doc_sections(doc), "meta": doc.get("meta", {})}
    if args.section:
        result["section_filter"] = args.section
    meta = doc.get("meta") or {}
    # 원본에서 읽지 않은 부분: 엑셀 시트당 행 상한 (읽기 결과도 저장 파일도 문서 전부가 아님)
    src_cut = f"원본 시트의 {MAX_SHEET_ROWS}행을 넘는 부분은 읽지 않음 (읽기 상한, --save 파일에도 없음)" \
        if meta.get("rows_truncated") else ""
    # 계산된 값이 없는 수식 칸: 빈 칸으로 읽혀 글에 없음 -> 읽기·저장 모두 UNVERIFIED (값이 계산된 적 없음)
    uncomputed = meta.get("uncomputed_formulas") or None
    formula_reason = xf.uncomputed_note(uncomputed) if uncomputed else ""
    if uncomputed:
        result.update(xf.uncomputed_fields(uncomputed))
    # 글자 없이 그림만 있는 쪽 (스캔 쪽): 읽으려는 범위 안의 쪽만 알림. 잘림이 아니라 주의 사항이라 notes에
    by_range = bool(args.units) or bool(sec_key)      # --section: 고른 단위의 첫 쪽~끝 쪽만
    blank_pages = _read_blank_pages(meta.get("pages_without_text") or [], selected, req_pages, by_range)
    notes: List[str] = []
    # 모든 쪽이 스캔 이미지(도장·Bates 번호 같은 짧은 글만 있음)이거나 읽으려는 쪽이 모두 스캔 쪽: 읽은 글은 본문이 아님
    # -> 저장해도 읽어도 UNVERIFIED (--save, 스캔 쪽만 고른 --pages도)
    got = "읽은 글은 도장·쪽 번호 같은 짧은 글뿐" if selected else "읽은 글 없음"
    ocr_reason = ""
    if meta.get("ocr_needed"):
        ocr_reason = f"스캔 이미지 PDF라 본문 글자 층이 없음 ({got}, OCR(글자 인식)한 PDF가 있어야 본문을 읽을 수 있음)"
    elif blank_pages and set(blank_pages) == _read_pages(doc, selected, req_pages, by_range):
        ocr_reason = (f"고른 쪽(p.{_page_list(blank_pages)})이 스캔 이미지라 본문 글자 층이 없음 "
                      f"({got}, OCR(글자 인식)한 PDF가 있어야 본문을 읽을 수 있음)")
    if blank_pages:
        result["pages_without_text"] = blank_pages[:50]
        notes.append(f"글이 거의 없고 그림만 있는 쪽 {len(blank_pages)}개 (스캔 쪽일 수 있음, OCR 필요): p.{_page_list(blank_pages)}")
        result["notes"] = notes

    if args.save:
        ext = _save_ext(args.save)
        save_arg = None if args.save is True else args.save
        out = rt.resolve_output_path(save_arg, os.path.splitext(name)[0], ext, near=path)
        header = ["위치", "쪽", "장절", "번호", "제목", "내용"]
        rows = [[u["loc"], u.get("page") if u.get("page") is not None else "", section_text(u), u.get("number") or "",
                 u.get("heading", ""), u["text"]] for u in selected]
        cut_cells = 0
        if ext == ".csv":
            out = rt.write_csv_rows(out, [header] + rows)
        elif ext == ".xlsx":
            out, cut_cells = write_xlsx_table(out, header, rows, [14, 6, 40, 10, 30, 80], "내용")
        else:
            out = rt.write_text(out, units_to_markdown(name, selected, path))
        result["output_file"] = out
        if selected and not ocr_reason:     # 0개 저장, 스캔 쪽의 도장·쪽 번호뿐: 다른 문서에 넣는 명령은 주지 않음
            result["confirm_commands"] = _write_confirm_commands(out)
        preview = f"{name}: {len(selected)}개 {unit_name}을(를) 파일로 저장했습니다 ({out})"
        # R1: 저장한 파일을 다시 열어 확인 (CSV는 행 수까지)
        problems = [p for p in [rt.check_saved_file(out, len(rows) + 1 if ext == ".csv" else None, delimiter=",")] if p]
        reasons: List[str] = []
        if cut_cells:
            reasons.append(_xlsx_cut_text(cut_cells))
        if src_cut:
            reasons.append(src_cut)
        if reasons:
            result.update(truncated=True, truncated_reason=reasons)
        problems += reasons + ([ocr_reason] if ocr_reason else []) + ([formula_reason] if formula_reason else [])
        if problems:
            payload = rt.ok_unverified(TOOL, "read", preview, problems, **result)
            if notes:
                payload["agent_directive"] += " 주의: " + "; ".join(notes) + "."
            return payload
        directive = ("VERIFIED: 저장한 파일을 다시 열어 확인했습니다."
                     + (" " + rt.INTO_OTHER_DOC if result.get("confirm_commands") else "")
                     if selected else "VERIFIED: 0건. 저장할 글이 없어 빈 파일을 저장했습니다. 사용자에게 그대로 알리고 다시 실행하지 마세요.")
        if notes:
            directive += " 주의: " + "; ".join(notes) + ". 사용자에게 알리세요."
        return rt.ok(TOOL, "read", preview, directive, verified=True, **result)

    if args.format == "summary":
        preview = (f"{name}: {doc['type']} 문서, {unit_name} {total}개, 쪽 {_doc_pages(doc) or '-'}, "
                   f"장절 {_doc_sections(doc)}개, 글자 수 {sum(len(u['text']) for u in units)}"
                   + (" (스캔 이미지 PDF, 본문은 OCR 필요)" if meta.get("ocr_needed") else ""))
        return rt.ok(TOOL, "read", preview, **result)

    if not selected and (meta.get("ocr_needed") or blank_pages):
        # 스캔 이미지(쪽): '읽을 글이 없다'고 단정하지 않음 (글자 층이 없어 OCR 전에는 읽을 수 없음)
        result.update(units=[], truncated=False, shown=0)
        what = "스캔 이미지 PDF" if meta.get("ocr_needed") else f"고른 쪽(p.{_page_list(blank_pages)})이 스캔 이미지"
        return rt.ok_unverified(TOOL, "read", f"{name}: {what}라 읽을 글자 층이 없습니다 (0건)",
                                [f"{what}라 글자 층이 없어 내용을 읽지 못함 (OCR(글자 인식)한 PDF가 있어야 읽을 수 있음)"],
                                **result)
    if not selected and formula_reason:
        result.update(units=[], truncated=False, shown=0)
        return rt.ok_unverified(TOOL, "read", f"{name}: 글이 없습니다 (0건, 계산된 값이 없는 수식 칸 있음)", [formula_reason],
                                **result)
    if not selected:
        result.update(units=[], truncated=False, shown=0)
        return rt.ok(TOOL, "read", f"{name}: 글이 없습니다 (0건)",
                     "VERIFIED: 0건. 이 문서에는 읽을 글이 없습니다. 사용자에게 그대로 알리고 다시 실행하지 마세요."
                     + (" 주의: " + "; ".join(notes) + "." if notes else ""), **result)

    # 본문: 상한까지 담고, 넘으면 다음 구간 명령 제공
    budget = rt.DEFAULT_MAX_CHARS
    shown: List[Dict[str, Any]] = []
    used = 0
    cut_units: List[int] = []
    for n, u in zip(numbers, selected):
        size = len(u["text"]) + len(u["loc"]) + len(section_text(u)) + 24
        if shown and used + size > budget:
            break
        text, cut = rt.clamp_text(u["text"], budget)
        view = _unit_view(n, u, text)
        if sec_key:          # 장절로 골라 읽기(규격 대조): 검토 메모에 그대로 쓸 근거 문구
            view["evidence"] = evidence_text(path, u)
        if cut:              # 한 단위의 글이 상한보다 길어 뒤가 잘림 (이어 읽기로는 그 뒤를 볼 수 없음)
            view["text_truncated"] = True
            view["text_chars"] = len(u["text"])
            cut_units.append(n)
        shown.append(view)
        used += size
    dropped = len(shown) < len(selected)
    truncated = dropped or bool(cut_units) or bool(src_cut)
    result["units"] = shown
    result["truncated"] = truncated
    result["shown"] = len(shown)
    reasons = []
    stem = os.path.splitext(name)[0]
    if cut_units:
        reasons.append(f"{unit_name} {','.join(map(str, cut_units))}의 글이 길어 앞 {budget}자만 표시 "
                       "(전체는 next_commands의 --save .md 파일에 있음)")
    if src_cut:
        reasons.append(src_cut)
    if reasons:
        result["truncated_reason"] = reasons
    nexts: List[Dict[str, str]] = []
    sec_args = ["--section", args.section] if sec_key else []
    # 저장 명령은 고른 범위를 그대로 (--section 읽기에서 문서 전체를 저장하지 않게)
    range_args = (["--units", args.units] if args.units else ["--pages", args.pages] if args.pages else []) if sec_key else []
    if dropped:
        nxt_i = len(shown)
        nxt, last = numbers[nxt_i], numbers[-1]
        cut = selected[nxt_i]
        # 남은 번호만 (고르지 않은 쪽·단위를 넣지 않음: '--pages 2,3,4,5,6,300'의 이어 읽기가 '--units 10-600'이던 것, 2026-11-09 검토 B13)
        rest_pages = sorted({u.get("page") for u in selected[nxt_i:] if u.get("page") is not None})
        if not by_page:
            cont = ["--pages", _number_ranges(numbers[nxt_i:])]
        elif cut.get("page") != selected[nxt_i - 1].get("page") and not args.units and rest_pages:
            cont = ["--pages", _number_ranges(rest_pages)]
        else:                                    # 쪽 중간에서 잘림 -> 단위 번호로 이어 읽기
            cont = ["--units", _number_ranges(numbers[nxt_i:])]
        nexts.append(rt.next_command("reader.py", ["read", "--file", path] + cont + sec_args, f"나머지 {unit_name} 이어서 읽기"))
    if dropped or cut_units:
        # 잘린 긴 글도 .md에는 전부 들어감 (.xlsx는 셀당 32000자 상한). 시트 행 상한으로 읽지 않은 부분은 저장해도 없음
        nexts.append(rt.next_command("reader.py", ["read", "--file", path] + range_args + sec_args
                                     + ["--save", f"{stem}_전체.md"],
                                     "읽은 내용 전체를 .md 파일로 저장 (잘린 글 포함, 다른 문서에 넣을 때)"
                                     + (f". 원본 시트의 {MAX_SHEET_ROWS}행을 넘는 부분은 저장 파일에도 없음" if src_cut else "")))
    if nexts:
        result["next_commands"] = nexts
    preview = (f"{name}: {unit_name} {len(shown)}개 표시 (전체 {total}개"
               + (f", 장절 '{args.section}'에 맞는 {len(selected)}개" if sec_key else "") + ")"
               + (" - 일부만 표시" if truncated else ""))
    if truncated:
        directive = "VERIFIED: 읽기 완료. 일부만 표시되었습니다 (" + "; ".join(reasons or ["길이 제한"]) + ")."
        if nexts:
            directive += " 사용자에게 알리고, 필요하면 next_commands로 이어서 읽거나 .md로 저장하세요."
        else:
            directive += " 더 읽거나 저장할 방법이 없습니다(읽기 상한). 사용자에게 그대로 알리세요."
    else:
        directive = "VERIFIED: 읽기 완료. 내용을 인용할 때는 page, section, number로 근거(파일 p.N, 장절 (번호))를 함께 알리세요."
    if sec_key:
        directive += " 근거는 units의 evidence를 그대로 쓰세요."
    if ocr_reason or formula_reason:
        return rt.ok_unverified(TOOL, "read", preview, [x for x in (ocr_reason, formula_reason) if x] + reasons, **result)
    if notes:
        directive += " 주의: " + "; ".join(notes) + ". 사용자에게 알리세요."
    return rt.ok(TOOL, "read", preview, directive, **result)


OUTLINE_MAX = 200    # outline 한 번에 담는 제목 수 (넘으면 --start로 이어서)


def cmd_outline(args) -> Dict[str, Any]:
    """문서의 장절 목록(제목 단위만): 제목마다 units(read --units 번호), page, level, number, text. 큰 문서를 8,000자씩 이어 읽지 않고
    필요한 장절만 read --section으로 읽게 (2026-11-09 검토 4장 6번)."""
    path = rt.resolve_input_path(args.file, "문서 파일")
    doc = load_document(path)
    units = doc["units"]
    heads = [(i, u) for i, u in enumerate(units, 1) if u.get("kind") == "heading"]
    start = max(1, args.start)
    shown = heads[start - 1:start - 1 + OUTLINE_MAX]
    items = []
    for k, (i, u) in enumerate(shown):
        nxt = heads[start - 1 + k + 1][0] if start - 1 + k + 1 < len(heads) else len(units) + 1
        item = {"units": f"{i}-{nxt - 1}" if nxt - 1 > i else str(i), "level": len(u.get("section") or []) or 1,
                "text": rt.clamp_text(u["text"], 200)[0]}
        if u.get("page") is not None:
            item["page"] = u["page"]
        if u.get("number"):
            item["number"] = u["number"]
        items.append(item)
    name = os.path.basename(path)
    result: Dict[str, Any] = {"file": path, "type": doc["type"], "count": len(heads), "shown": len(items),
                              "units_total": len(units), "headings": items, "truncated": start - 1 + len(items) < len(heads)}
    if result["truncated"]:
        result["next_commands"] = [rt.next_command("reader.py", ["outline", "--file", path, "--start", str(start + len(items))],
                                                   "다음 제목 이어서 보기 (읽기 전용)")]
    if not heads:
        return rt.ok(TOOL, "outline", f"{name}: 제목(장절)이 없음 (count 0, 단위 {len(units)}개)",
                     "VERIFIED: 0건. 이 문서에서 장절 제목을 찾지 못했습니다. 내용은 reader.py read --file로 읽으세요.", **result)
    preview = f"{name}: 제목 {len(heads)}개" + (f" (제목 {start}~{start + len(items) - 1} 표시)" if result["truncated"] or start > 1 else "")
    return rt.ok(TOOL, "outline", preview,
                 "VERIFIED: 장절 목록을 읽었습니다. 장절 하나를 읽으려면 reader.py read --file \"<같은 파일>\" --section \"<제목 일부>\" "
                 "또는 그 제목의 units 값으로 --units를 쓰세요." + (" 나머지 제목은 next_commands[0]." if result["truncated"] else ""),
                 **result)


def _read_pages(doc: Dict[str, Any], selected: List[Dict[str, Any]], req_pages: Optional[set], by_units: bool) -> set:
    """읽으려는 쪽: --pages면 그 쪽들, --units면 고른 단위의 첫 쪽~끝 쪽, 아니면 문서 전체 쪽."""
    if req_pages is not None:
        return set(req_pages)
    if by_units:
        got = [u["page"] for u in selected if u.get("page") is not None]
        return set(range(min(got), max(got) + 1)) if got else set()
    return set(range(1, int(_doc_pages(doc) or 0) + 1))


def _read_blank_pages(flagged: List[int], selected: List[Dict[str, Any]], req_pages: Optional[set],
                      by_units: bool) -> List[int]:
    """글자 없는 쪽 중 읽으려는 범위 안의 쪽. --pages면 그 쪽들, --units면 고른 단위의 첫 쪽~끝 쪽, 아니면 전부."""
    if not flagged:
        return []
    if req_pages is not None:
        return [p for p in flagged if p in req_pages]
    if by_units:
        got = [u["page"] for u in selected if u.get("page") is not None]
        return [p for p in flagged if got and min(got) <= p <= max(got)]
    return list(flagged)


# ==============================================================================
# 6. 명령: index
# ==============================================================================

def cmd_index(args) -> Dict[str, Any]:
    if sum(bool(x) for x in (args.folder, args.file, args.all)) != 1:
        raise rt.InputError("--folder, --file, --all 중 하나만 지정하세요. 예: reader.py index --folder \"C:\\docs\\규격\"")
    skipped_types: Dict[str, int] = {}
    missing = 0
    target = ""
    orphan_docs: List[str] = []
    if args.folder:
        target = _folder_arg(args.folder)
        depth = 50 if args.recursive else max(0, args.depth)
        files, skipped_types = _collect_folder(target, depth)
    elif args.file:
        files = [rt.resolve_input_path(args.file, "등록할 파일")]
        file_type(files[0])
        target = files[0]
    else:
        index = load_index()
        files = [p for p in index if os.path.exists(p)]
        missing = len(index) - len(files)
        target = "(등록된 문서 전체)"
        # 색인에서 빠진 캐시 파일(동시 실행, 색인 저장 전 중단)도 원본이 있으면 다시 넣음 (JIT 복구)
        orphan_docs = _orphan_cache_files(index)
        if orphan_docs:
            files += orphan_docs
            rt.log("WARN", f"색인에 없는 캐시 문서 {len(orphan_docs)}개를 찾아 다시 등록합니다.")
    res = _index_files(files, max(1, args.max_files), retry_failed=bool(args.file))
    if args.file:
        _raise_single_failure(res)
    readded = [d for d in res["docs"]
               if os.path.normcase(d["file"]) in {os.path.normcase(p) for p in orphan_docs}] if orphan_docs else []
    if readded:
        rt.note_recovered(f"색인에서 빠진 캐시 문서 {len(readded)}개 -> 색인에 다시 넣음")
    # 동시 실행 등으로 색인에서 빠진 문서를 다시 넣고 확인한 뒤에만 VERIFIED (JIT 복구)
    lost = _repair_index(res["docs"])
    failures = [{"file": os.path.basename(f["file"]), "path": f["file"], "reason": f["reason"]} for f in res["failures"]]
    result: Dict[str, Any] = {
        "target": target, "count": len(files), "indexed": len(res["new"]), "updated": len(res["updated"]),
        "unchanged": len(res["cached"]), "failed": len(failures), "failures": failures[:MAX_LIST_ITEMS],
        "remaining": res["remaining"], "truncated": res["remaining"] > 0 or len(failures) > MAX_LIST_ITEMS,
        "total_units": sum(len(d["units"]) for d in res["docs"]),
    }
    if skipped_types:
        result["skipped_types"] = skipped_types
    if missing:
        result["missing"] = missing
    if readded:
        result["readded_to_index"] = len(readded)
    nexts: List[Dict[str, str]] = []
    if res["remaining"]:
        same = ["index"] + (["--folder", target] if args.folder else ["--all"]) + ["--max-files", str(args.max_files)]
        if args.folder and args.recursive:
            same.append("--recursive")
        elif args.folder:
            same += ["--depth", str(args.depth)]
        nexts.append(rt.next_command("reader.py", same, f"남은 {res['remaining']}개 파일 이어서 등록 (읽기 전용)"))
    search_target = ["--file", target] if (args.folder or args.file) else []
    nexts.append(rt.next_command("reader.py", ["search", "--query", "<검색어>"] + search_target,
                                 "등록한 문서에서 검색 (<검색어>를 사용자 검색어로 바꿈)", needs_input=True))
    result["next_commands"] = nexts
    if not files:
        what = "폴더에 지원 형식 문서가 없습니다" if args.folder else "등록된 문서가 없습니다"
        return rt.ok(TOOL, "index", f"0건: {what}",
                     f"VERIFIED: 0건. {what}. 사용자에게 그대로 알리고 다시 실행하지 마세요.",
                     **{k: v for k, v in result.items() if k != "next_commands"})
    preview = (f"문서 {len(files)}개: 새로 등록 {result['indexed']}, 다시 등록 {result['updated']}, "
               f"변경 없음 {result['unchanged']}, 실패 {result['failed']}"
               + (f", 남음 {res['remaining']} (--max-files {args.max_files})" if res["remaining"] else ""))
    # R1: 새로 읽은 문서의 캐시 파일이 실제로 저장됐는지 확인 (캐시 저장 실패는 결과 JSON에만 보이지 않던 문제)
    unsaved = [os.path.basename(d["file"]) for d in res["docs"]
               if d.get("_cache_failed") or not os.path.isfile(_doc_cache_path(d["file"]))]
    if lost:
        result["index_missing"] = lost[:MAX_LIST_ITEMS]
        return rt.ok_unverified(TOOL, "index", preview,
                                [f"색인에 넣지 못한 파일 {len(lost)}개 (검색에서 빠질 수 있음, 다시 index 실행): "
                                 + ", ".join(lost[:5])], **result)
    if unsaved:
        result["cache_failed"] = unsaved[:MAX_LIST_ITEMS]
        return rt.ok_unverified(TOOL, "index", preview,
                                [f"캐시에 저장하지 못한 파일 {len(unsaved)}개 (다음 검색 때 다시 읽음): {', '.join(unsaved[:5])}"], **result)
    scans = [os.path.basename(d["file"]) for d in res["docs"] if (d.get("meta") or {}).get("ocr_needed")]
    if scans:
        # 글자 없는 스캔 PDF를 '등록 완료, 검색 가능'으로만 알리던 것 (2026-11-09 검토 B13)
        result["ocr_needed"] = scans[:MAX_LIST_ITEMS]
        result.setdefault("warnings", []).append(f"글자가 없는 스캔 PDF {len(scans)}개는 검색되지 않음 (OCR 필요): {', '.join(scans[:5])}")
    if res["remaining"]:
        directive = ("VERIFIED: 일부만 등록했습니다(캐시에 저장 확인). 남은 파일은 next_commands[0]을 그대로 실행해 이어서 등록하세요 "
                     "(문서를 바꾸지 않는 읽기 전용 작업). 실패한 파일은 failures를 사용자에게 알리세요.")
    else:
        directive = ("VERIFIED: 등록 완료(캐시에 저장 확인). 이제 reader.py search로 검색할 수 있습니다. "
                     "실패한 파일이 있으면 failures를, warnings(스캔 PDF 등)가 있으면 그것도 사용자에게 알리세요.")
    return rt.ok(TOOL, "index", preview, directive, **result)


def _orphan_cache_files(index: Dict[str, Dict[str, Any]]) -> List[str]:
    """색인에 없는 캐시 파일 중 원본이 아직 있는 문서의 경로 (index --all이 다시 넣을 대상)."""
    cache_dir = os.path.join(rt.get_cache_dir(), "docs")
    if not os.path.isdir(cache_dir):
        return []
    known = {os.path.basename(_doc_cache_path(p)) for p in index}
    found: List[str] = []
    for name in sorted(os.listdir(cache_dir)):
        if not name.endswith(".json") or name in known:
            continue
        try:
            path = json.loads(rt.read_text_file(os.path.join(cache_dir, name))).get("file")
        except Exception:
            continue                       # 손상된 캐시 파일은 clear --all이 정리
        if (isinstance(path, str) and os.path.isfile(path) and os.path.basename(_doc_cache_path(path)) == name
                and os.path.splitext(path)[1].lower() in SUPPORTED):
            found.append(path)
    return found


def _repair_index(docs: List[Dict[str, Any]]) -> List[str]:
    """불러온 문서가 색인에 모두 있는지 다시 읽어 확인하고, 빠진 것(다른 실행이 동시에 색인을 덮어씀)을 다시 넣습니다.
    다시 넣은 뒤에도 빠진 파일 이름 목록을 돌려줍니다 (비면 확인 완료)."""
    want = {os.path.normcase(os.path.abspath(d["file"])): d for d in docs
            if not d.get("_cache_failed") and os.path.isfile(_doc_cache_path(d["file"]))}

    def absent() -> List[Dict[str, Any]]:
        have = {os.path.normcase(os.path.abspath(k)) for k in load_index()}
        return [d for k, d in want.items() if k not in have]

    lost = absent()
    if not lost:
        return []

    def put(index: Dict[str, Dict[str, Any]]) -> None:
        for d in lost:
            _index_put(index, d["file"], _index_entry(d, _doc_cache_path(d["file"])))
    try:
        _index_update(put)
    except OSError as e:
        rt.log("WARN", f"색인 고치기 실패: {e}")
    still = absent()
    fixed = len(lost) - len(still)
    if fixed:
        rt.log("WARN", f"색인에서 빠진 문서 {fixed}개를 다시 넣었습니다 (동시 실행 등).")
        rt.note_recovered(f"색인에서 빠진 문서 {fixed}개 -> 다시 넣음 (동시 실행 등)")
    return [os.path.basename(d["file"]) for d in still]


# ==============================================================================
# 7. 명령: search (캐시 기준)
# ==============================================================================

def _bad_regex(query: str, e: "re.error", args: Any = None) -> rt.InputError:
    """--regex의 정규식 오류 (STOP_RETRY). 정규식 문법(\\d, .*, [..], |, (?..) 등)이 없는 글이면 글자 그대로 찾는 전체 명령을 [추천]."""
    pos = f"{e.pos + 1}번째 글자, " if getattr(e, "pos", None) is not None else ""
    head = f"--regex의 정규식이 잘못되었습니다 ({pos}{e.msg})."
    literal = not re.search(r"\\.|\.[*+?]|\[[^\]]*\]|\(\?|\||\{\d", query)
    if literal and args is not None:
        cmd = ["search", "--query", query]
        for flag, val in (("--file", args.file), ("--pages", args.pages), ("--section", args.section)):
            if val:
                cmd += [flag, val]
        cmd += ["--full"] if args.full else []
        rec = rt.next_command("reader.py", cmd, "")["command"]
        tail = " 정규식 문법이 없는 글이라 글자 그대로 찾으면 됩니다(--regex 없이, 기호·띄어쓰기는 무시)."
        return rt.InputError(f"{head}{tail} [추천] {rec}", log_message=f"{head}{tail} [추천 명령 생략: 검색어 {rt.DOC_OMITTED}]")
    return rt.InputError(f"{head} 정규식을 고쳐 다시 실행하세요 (기호를 글자로 찾으려면 앞에 \\를 붙임).")


def _nested_repeat(query: str) -> bool:
    """정규식에 '여러 번 반복' 안에 다시 '여러 번 반복'이 있고, 바깥 반복 안에 반드시 나오는 한 글자(구분 글자)가 없는지
    (파국적 되추적의 흔한 모양: (a+)+, (\\w+\\s?)+). '(\\d{1,3},)+'처럼 구분 글자가 있으면 괜찮음. 해석하지 못하면 False."""
    try:
        import re._parser as sre               # 파이썬 3.11+
    except Exception:                          # noqa: BLE001
        return False
    repeats = {sre.MAX_REPEAT, sre.MIN_REPEAT}
    single = {sre.LITERAL, sre.NOT_LITERAL, sre.IN, sre.ANY}

    def has_many(items: Any) -> bool:
        for op, av in items:
            if op in repeats and (av[1] is sre.MAXREPEAT or av[1] > 1):
                return True
            if op in repeats and has_many(av[2]):
                return True
            if op is sre.SUBPATTERN and has_many(av[-1]):
                return True
            if op is sre.BRANCH and any(has_many(b) for b in av[1]):
                return True
        return False

    def flat(items: Any) -> List[Any]:
        out: List[Any] = []
        for op, av in items:
            if op is sre.SUBPATTERN:
                out += flat(av[-1])
            else:
                out.append((op, av))
        return out

    def walk(items: Any) -> bool:
        for op, av in items:
            if op in repeats:
                lo, hi, sub = av
                if (hi is sre.MAXREPEAT or hi > 1) and has_many(sub) and not any(o in single for o, _ in flat(sub)):
                    return True
                if walk(sub):
                    return True
            elif op is sre.SUBPATTERN and walk(av[-1]):
                return True
            elif op is sre.BRANCH and any(walk(b) for b in av[1]):
                return True
            elif op in (sre.ASSERT, sre.ASSERT_NOT) and walk(av[1]):
                return True
        return False
    try:
        return walk(sre.parse(query))
    except Exception:                          # noqa: BLE001
        return False


def build_matcher(query: str, regex: bool, args: Any = None):
    """검색 조건을 함수로 만듭니다. 'A;B'(그리고), 'A|B'(또는) 조합 지원 (;가 |보다 먼저 묶임).
    반환: match_fn(text) -> 강조할 단어 목록 또는 None. 대소문자는 늘 무시."""
    flags = re.IGNORECASE
    if regex:
        try:
            pat = re.compile(query, flags)
        except re.error as e:
            raise _bad_regex(query, e, args)
        if _nested_repeat(query):
            # 파이썬 re는 중간에 멈출 수 없어 '(\w+\s?)+$X'처럼 반복 안의 반복은 몇 분씩 걸려 120초 제한에 JSON 없이 끊김 (2026-11-09 검토 B10)
            raise rt.InputError(f"정규식 '{query}'에 반복 안의 반복(예: (a+)+, (\\w+\\s?)*)이 있어 검색이 끝나지 않을 수 있어 실행하지 않았습니다. "
                                "안쪽이나 바깥쪽 반복 하나를 빼거나, --regex 없이 낱말 검색어(A;B, A|B)로 찾으세요.")

        def match_regex(text: str):
            m = next((m for m in pat.finditer(text) if m.group(0)), None)   # 빈 일치('a*')는 일치로 보지 않음
            return [m.group(0)] if m else None
        return match_regex

    groups = parse_query(query)
    if not groups:
        raise rt.InputError("검색어가 비어 있습니다.")

    pats = [[_term_pattern(t, flags) for t in g] for g in groups]
    lits = [[literal_finder(t) for t in g] for g in groups]

    def match_terms(text: str):
        # 숫자가 든 검색어는 정규화 매처(층 B) 또는 글자 그대로(층 A)로 맞음. 층 A로 다 맞은 묶음을 먼저 고름 (literal_hit)
        t = unicodedata.normalize("NFC", text)
        first = None
        for g, ps, ls in zip(groups, pats, lits):
            literal = True
            for p, lf in zip(ps, ls):
                lit = lf.search(t) is not None if lf is not None else False
                if not lit and not p.search(t):
                    break
                literal = literal and (lf is None or lit)
            else:
                if literal:
                    return g
                first = first or g
        return first
    return match_terms


_OLD_QUERY_OP = re.compile(r'"[^"]*"|\u201c[^\u201d]*\u201d|(?<!\S)(AND|OR)(?!\S)')


def parse_query(query: str) -> List[List[str]]:
    """'A;B|C D' -> [['A', 'B'], ['C D']]. 저장소 구분자 규칙(AGENTS.md 사용법 단순성 3번): 그리고는 ;, 또는은 |
    (;가 먼저 묶임, excel --where·ml-kit --where와 같음). 띄어 쓴 단어는 한 구문. ;나 |가 든 글은 큰따옴표로 묶음
    (따옴표는 검색어가 아님). 1.4.81: 옛 'A AND B', 'A OR B'(대문자 낱말)는 받지 않고 고친 전체 명령 [추천]."""
    text = unicodedata.normalize("NFC", query or "")
    if any(m.group(1) for m in _OLD_QUERY_OP.finditer(text)):
        fixed = _OLD_QUERY_OP.sub(lambda m: {"AND": ";", "OR": "|"}.get(m.group(1) or "", m.group(0)), text)
        fixed = re.sub(r"\s*([;|])\s*", r"\1", fixed).strip()
        raise rt.InputError(f"--query의 AND/OR는 쓰지 않습니다. 그리고는 ;, 또는은 |로 씁니다 (예: --query \"온도;습도|압력\"). "
                            f"[추천] {rt.retry_replacing(query, fixed, '--query')}",
                            log_message="--query에 옛 AND/OR 표기 [추천 명령 생략: 검색어 " + rt.DOC_OMITTED + "]")
    groups: List[List[str]] = [[]]
    for tok in _query_tokens(text):
        if tok == "|":
            groups.append([])
        elif tok == ";":
            continue
        elif len(tok) >= 2 and {'"': '"', "\u201c": "\u201d", "'": "'"}.get(tok[0]) == tok[-1]:
            if tok[1:-1].strip():
                groups[-1].append(tok[1:-1].strip())
        elif " ".join(tok.split()):
            groups[-1].append(" ".join(tok.split()))
    return [g for g in groups if g]


def _query_tokens(text: str) -> List[str]:
    """--query를 ; | 와 낱말로 나눔. 큰따옴표 묶음은 어디서나, 작은따옴표 묶음은 한 항목 전체일 때만 (1.4.83: excel --where와 같은
    'a;b' 표기. PowerShell 5.1은 큰따옴표를 인자 안에 넘기기 어려움. "don't"처럼 낱말 안의 '는 글자)."""
    out: List[str] = []
    i, n = 0, len(text)
    while i < n:
        c = text[i]
        if c in ";|":
            out.append(c)
            i += 1
            continue
        if c.isspace():
            i += 1
            continue
        close = {'"': '"', "\u201c": "\u201d", "'": "'"}.get(c)
        end = text.find(close, i + 1) if close else -1
        if close and end > i:
            rest = text[end + 1:].lstrip()
            if c != "'" or not rest or rest[0] in ";|":
                out.append(text[i:end + 1])
                i = end + 1
                continue
        if c in '"\u201c':                    # 짝 없는 큰따옴표는 버림 (검색어가 아님)
            i += 1
            continue
        j = i + 1
        while j < n and text[j] not in ';|"\u201c':
            j += 1
        out.append(text[i:j])
        i = j
    return out


_SEP = r"[\W_]*"          # 글자 사이에 끼어도 되는 것: 띄어쓰기, 줄바꿈, 기호(# - . : / ( ) 등, 대시 종류 모두)

# 검색어 낱말 끝의 조사·어미: 떼고 찾음 (글에서는 낱말 뒤에 무엇이 와도 맞음). 긴 것부터.
# 떼고 남는 글이 2자 이상일 때만 ('높이', '단가', '회의'는 그대로). '도'·'만'은 낱말 끝 글자인 경우가 많아('정확도') 떼지 않음.
# '로'는 ㄹ 받침 뒤에서만 뗌('케이블로' -> '케이블'). 뗀 낱말 뒤에는 한글이 이어지지 않아야 맞음(조사는 허용:
# '케이블은'은 맞고 '케이블카'는 아님). 받침 없는 글자 뒤의 '로'는 낱말 끝 글자인 경우가 많아('보호회로', '도로', '냉각수로')
# 떼지 않고 그대로 찾음 ('전기회로'가 '전기회사'에 맞지 않게). '으로'는 뗌.
_PARTICLES = ("에서의", "에서는", "으로는", "으로서", "으로써", "이어야", "에서", "으로", "에는", "에게", "까지", "부터",
              "보다", "과의", "와의", "이며", "이고", "이다", "은", "는", "이", "가", "을", "를", "의", "에", "와", "과", "및", "로")
# 받침에 따라 붙는 조사: 앞 글자의 받침이 맞을 때만 뗌 ('와이파이', '역효과'의 끝 글자는 조사가 아님)
_AFTER_FINAL = {"이", "은", "을", "과", "이며", "이고", "이다", "이어야", "과의"}       # 받침 있는 글자 뒤
_AFTER_OPEN = {"가", "는", "를", "와", "와의"}                                      # 받침 없는 글자 뒤
_HANGUL = re.compile(r"[\uac00-\ud7a3]")
_RIEUL = 8                                     # 종성 번호 ㄹ


def _final(ch: str) -> int:
    return (ord(ch) - 0xAC00) % 28


def _has_final(ch: str) -> bool:
    """한글 음절에 받침이 있는지."""
    return (ord(ch) - 0xAC00) % 28 != 0


def _strip_word(word: str) -> str:
    if not _HANGUL.search(word):
        return word
    for p in _PARTICLES:
        if word.endswith(p) and len(word) - len(p) >= 2 and _HANGUL.match(word[len(word) - len(p) - 1]):
            before = word[len(word) - len(p) - 1]
            if (p in _AFTER_FINAL and not _has_final(before)) or (p in _AFTER_OPEN and _has_final(before)):
                continue
            if p == "로" and _final(before) != _RIEUL:
                continue                       # 받침 없는 글자 뒤 '로'는 낱말 끝일 수 있음 ('보호회로')
            return word[: -len(p)]
    return word


def strip_particles(term: str) -> str:
    """'정격전압은 5V' -> '정격전압 5V' (낱말마다 끝의 조사를 뗌)."""
    return " ".join(_strip_word(w) for w in term.split())


_DIAMETER = "\u00d8\u00f8\u2300"          # 지름 기호: Ø ø U+2300 (서로 같음)
_PHI = "\u03a6\u03c6"                     # Φ φ: 지름으로도 상(phase, '3φ')으로도 쓰여 Ø와 같게 보지 않음 (1.4.69)


def _char(c: str) -> str:
    """글자 하나의 정규식: 영문·숫자는 전각(ＰＣＢ, ５)도 맞게, 지름 기호는 서로 같게 (Φ·φ는 Ø와 다름)."""
    if c.isascii() and c.isalnum():
        return f"[{c}{chr(ord(c) + 0xFEE0)}]"
    if c in _DIAMETER:
        return f"(?-i:[{_DIAMETER}])"
    if c in _PHI:
        return f"(?-i:[{_PHI}])"
    return re.escape(c)


_OHM_PREFIX = ("", "k", "K", "M", "G", "m")


def _unit_alias(u: str) -> Optional[str]:
    """단위 표기의 다른 글자 모음: kΩ = kohm = kΩ(U+2126), mm = ㎜, mA = ㎃. 다른 표기가 없으면 None."""
    forms: List[str] = []
    for end in ("\u03a9", "\u2126", "ohm", "옴"):
        if u.endswith(end) and u[: -len(end)] in _OHM_PREFIX:
            pre = u[: -len(end)]
            forms = [pre + "\u03a9", pre + "\u2126", pre + "ohm"] + (["옴"] if not pre else [])
            break
    caps = _caps_unit(u)                       # 모두 대문자('MM', 'MW')는 대소문자를 가리지 않음 (1.4.70)
    squares = [chr(c) for c in range(0x3380, 0x33E0)
               if unicodedata.normalize("NFKC", chr(c)) in forms + [u]
               or (caps and unicodedata.normalize("NFKC", chr(c)).upper() == u)]
    if not forms and not squares:
        return None
    # 첫 글자가 단위 접두(m, M, μ 등)면 대소문자를 가림: 'mohm'(밀리로 봄)은 'MΩ'과 다름
    alts = [r"\s*".join((_prefix_char(x, f) if i == 0 and not caps else None) or _char(x) for i, x in enumerate(f))
            for f in sorted(set(forms + [u]), key=len, reverse=True)]
    return "(?:" + "|".join(alts + [re.escape(x) for x in squares]) + ")"


_NUM_TOKEN = re.compile(r"\d{1,3}(?:,\d{3})+(?:\.\d+)?|\d+(?:\.\d+)?|[^\W\d_]|\S")


# CP949에 있는 전각 단위 글자(㎜ ㎝ ㎏ ㎖ ㎃ ㎾ ㏊ ...: U+3380~U+33DF)와 ℓ. 글은 NFKC로 바꾸지 않고 찾으므로 따로 둠
_SQUARE_UNIT = "[\u3380-\u33df\u2113]"


_MINUS = "\\-\u2010\u2011\u2012\u2013\u2212"          # 글에서 빼기로 쓰는 글자 ('-', U+2010~2013, U+2212)
_DASHES = "-\u2010\u2011\u2012\u2013\u2014\u2212"      # 숫자 앞 부호 판정에서 대시로 보는 글자 (U+2014 포함)
# 숫자 앞 부호 판정 (보수적 원칙, 1.4.69): 정확 일치(exact, VERIFIED)는 검색어 숫자와 글 숫자의 앞 부호가 분명히 같을 때만.
# 글의 숫자 바로 앞(띄어쓰기 여러 칸·'/' 건너)이 -, U+2212, U+2013, ±, +/-, + / -, +-, + 중 하나인데 검색어에 같은 부호가
# 없으면(또는 그 반대, 범위 대시인지 부호인지 애매하면) exact로 세지 않고 숫자 안전망 후보(number, UNVERIFIED)로 보냄.
# 범위 대시는 양옆이 띄어진 '0 - 20', '5 V - 10 V'만 (부호 없는 20과는 exact, 부호 있는 -20과는 아님).
_SIGN_GAP = 12                                         # 부호를 찾으며 건너뛸 띄어쓰기·'/'의 최대 글자 수
_DASH_LABEL = re.compile(r"(?<![A-Za-z])(?:min|max|typ|nom|ta|tc|tj|tamb|temp|utc|gmt|to)$", re.IGNORECASE)
_RANGE_LEFT = re.compile(r"(?<![\d.])\d+(?:[.,]\d+)?(?:[ \t]?([^\s\d]{1,4}))?\s+$")
_GLUED_RANGE_LEFT = re.compile("(?<![\\d.])\\d+(?:[.,]\\d+)?[ \u00a0]?[\u2103\u2109%\u03a9\u2126]$")   # '40 ℃'(기호 단위)


def _range_unit(u: Optional[str]) -> bool:
    """범위 대시 앞 '숫자 단위'의 단위인지 ('5 V - 10 V'의 V). 'to', 'Min' 같은 낱말은 아님."""
    if not u:
        return True
    n = unicodedata.normalize("NFKC", u)
    return n in ("%", "°", "°C", "°F") or bool(_unit_key(n) or _unit_key("°" + n))


def _text_sign(text: str, p: int) -> str:
    """text[p]에서 시작하는 숫자의 앞 부호: 'none', 'range'(띄어 쓴 범위 대시), 'hyphen'(영문·숫자에 붙은 대시: 'A-20', '1-20'),
    'minus', 'ambig'('25 -40': 범위인지 부호인지 애매), 'pm'(±, +/-, + / -, +-), 'plus'."""
    j, lim = p, max(0, p - _SIGN_GAP)
    while j > lim and (text[j - 1].isspace() or text[j - 1] == "/"):
        j -= 1
    if j == 0 or (j == lim and lim > 0):
        return "none"
    c = text[j - 1]
    if c == "\u00b1":
        return "pm"
    if c == "+":
        return "plus"
    if c not in _DASHES:
        return "none"
    d = j - 1
    k, lim = d, max(0, d - _SIGN_GAP)
    while k > lim and (text[k - 1].isspace() or text[k - 1] == "/"):
        k -= 1
    if k > 0 and text[k - 1] == "+":
        return "pm"
    after, before = text[d + 1:p], text[k:d]
    prev = text[d - 1] if d > 0 else ""
    if not after and not before and prev.isascii() and prev.isalnum():
        return "minus" if _DASH_LABEL.search(text[max(0, d - 5):d]) else "hyphen"     # 'Min-40', 'UTC-08:30'은 부호
    if not after and not before and _GLUED_RANGE_LEFT.search(text, max(0, d - 16), d):
        return "range"                                     # '40 ℃\u201385 ℃': 기호 단위 바로 뒤에 붙인 범위 대시 (1.4.72)
    left = _RANGE_LEFT.search(text[max(0, d - 20):d]) if before and "/" not in before else None
    if left and _range_unit(left.group(1)):
        return "range" if after and "/" not in after else "ambig"
    return "minus"


# 검색어 숫자의 앞 부호(왼쪽) -> exact로 셀 글 숫자의 앞 부호. 'after'는 검색어에서 다른 낱말 뒤의 부호 없는 숫자
# ('SOI 1' = 'SOI-1'). 그 밖의 조합은 exact가 아님 (숫자 안전망 후보)
_SIGN_OK: Dict[str, Tuple[str, ...]] = {
    "none": ("none", "range"), "after": ("none", "range", "hyphen"), "hyphen": ("none", "range", "hyphen"),
    "range": ("none", "range", "hyphen"), "minus": ("minus",), "ambig": ("minus", "ambig"), "pm": ("pm",), "plus": ("plus",),
}
# 검색어 숫자가 맨 앞일 때 글에서 부호까지 차지할 패턴 (뒤에 _text_sign으로 다시 확인)
_SIGN_HEAD = {"minus": rf"[{_MINUS}\u2014]\s*", "ambig": rf"[{_MINUS}\u2014]\s*",
              "pm": r"(?:±|\+\s*/\s*-|\+-)\s*", "plus": r"\+\s*"}
_MARK_CLASS: Dict[str, str] = {}                       # 숫자 표시 그룹 이름 -> 검색어 숫자의 앞 부호
_MARK_RE = re.compile(r"\(\?P<_n\d+>\)")
_NUM_END = r"(?![.,]?\d)(?!\s*(?:[\u00b2\u00b3\u00b9\u2070-\u209f]|\^|\*\*))"    # 지수(위첨자 '10-3', '10^-3')가 이어지면 다른 값


def _mark(cls: str) -> str:
    """숫자 시작 자리 표시 (빈 그룹). 일치 뒤 _text_sign으로 글의 부호를 보고 _SIGN_OK에 없으면 버림."""
    name = f"_n{len(_MARK_CLASS)}"
    _MARK_CLASS[name] = cls
    return f"(?P<{name}>)"


class _SignChecked:
    """숫자 표시가 든 검색어 정규식: 일치마다 글의 숫자 앞 부호를 확인 (re.Pattern의 search·finditer·pattern만)."""

    def __init__(self, pat: "re.Pattern[str]", names: List[str]):
        self.pat = pat
        self.pattern = _MARK_RE.sub("", pat.pattern)      # 강조용으로 합칠 때 (그룹 이름이 겹치지 않게)
        self.checks = [(n, _MARK_CLASS[n]) for n in names]

    def _ok(self, m: "re.Match[str]") -> bool:
        for name, cls in self.checks:
            p = m.start(name)
            if p >= 0 and _text_sign(m.string, p) not in _SIGN_OK[cls]:
                return False
        return True

    def search(self, text: str, pos: int = 0) -> Optional["re.Match[str]"]:
        while True:
            m = self.pat.search(text, pos)
            if m is None or self._ok(m):
                return m
            pos = m.start() + 1

    def finditer(self, text: str):
        pos = 0
        while pos <= len(text):
            m = self.search(text, pos)
            if m is None:
                return
            yield m
            pos = m.end() if m.end() > m.start() else m.end() + 1


# 끝 단위 뒤에 이어져도 같은 단위인 꼬리 ('5 Vdc', '3 sec', '10 hrs'). 다른 단위로 이어지거나 뜻이 다른 꼬리
# ('10 mm', '1 Ah', '30 dBm', '5 Vp-p'·'5 Vpp'·'10 Apk'(피크), '5 Vrms'(실효), '1 barg'(게이지), '30 dBA'(가중))는 아님.
# 키: _UNIT_TABLE의 단위(정확히) 또는 기준 단위
_ELEC_TAIL = ("max", "min", "typ", "nom", "ac", "dc", "in", "out")
_UNIT_TAILS: Dict[str, Tuple[str, ...]] = {
    "V": _ELEC_TAIL + ("o", "cc", "dd", "ee", "ss"),
    "A": _ELEC_TAIL,
    "W": ("max", "min", "typ", "nom", "ac", "dc"),
    "s": ("ec", "ecs", "econd", "econds"), "ms": ("ec", "ecs"), "us": ("ec", "ecs"), "μs": ("ec", "ecs"),
    "ns": ("ec", "ecs"), "sec": ("s", "ond", "onds"), "min": ("s", "ute", "utes"), "h": ("r", "rs", "our", "ours"),
    "hr": ("s",),
}
_UNIT_LETTER = "[A-Za-z\uff21-\uff3a\uff41-\uff5a\u03a9\u2126\u03bc\u00b5\u00b0\u2103]"


def _unit_end(unit_key: Optional[str]) -> str:
    """검색어 끝 단위 뒤 경계: 같은 단위 표기의 꼬리만 허용하고 다른 단위 글자가 이어지면 아님."""
    base = _UNIT_TABLE[unit_key][0] if unit_key in _UNIT_TABLE else ""
    tails = _UNIT_TAILS.get(unit_key or "") or _UNIT_TAILS.get(base) or ("ac", "dc")
    alts = "|".join(re.escape(t) for t in sorted(tails, key=len, reverse=True))
    return rf"(?:(?=(?:{alts})(?!{_UNIT_LETTER}))|(?!{_UNIT_LETTER}))"


# 숫자 바로 뒤 단위 접두의 대소문자 (m = 밀리, M = 메가, p = 피코, P = 페타): 대소문자가 다르면 다른 단위 ('10 mHz' != '10 MHz').
# μ·u·U+00B5는 같은 마이크로. 접두가 아닌 단위('min', 'max', 'ppm', 'psi', 'Pa', 'pH')는 대소문자 무관
_CASE_PREFIX = "mMpP"
_CASE_FREE_UNITS = {"min", "mins", "minute", "minutes", "max", "mile", "miles", "mph", "ppm", "ppb", "ppt", "psi", "psig",
                    "psia", "pcs", "pc", "pa", "pas", "ph"}
_MICRO = "u\u03bc\u00b5"


def _prefix_char(c: str, unit: str) -> Optional[str]:
    """단위 첫 글자 c(단위 글자 전체 unit)의 정규식: 대소문자를 가려야 하면 (?-i:...), 아니면 None."""
    if c in _CASE_PREFIX and unit.lower() not in _CASE_FREE_UNITS:
        return f"(?-i:{_char(c)})"
    if c in _MICRO and len(unit) >= 2:
        return f"(?-i:[{_MICRO}])"
    return None


def _caps_unit(unit: str) -> bool:
    """검색어 숫자 뒤 단위를 영문 대문자로만 썼는지 ('MM', 'UF', 'MA': 대소문자를 뜻으로 쓰지 않은 표기로 봄)."""
    return len(unit) >= 2 and unit.isascii() and unit.isalpha() and unit.isupper()


def _unit_letters(w: List[str], ti: int) -> str:
    """낱말 토큰 w의 ti부터 이어진 영문·그리스 글자 (숫자 뒤 단위 글자: '10mHz이상' -> 'mHz')."""
    out = ""
    for x in w[ti:]:
        if len(x) != 1 or not (x.isascii() and x.isalpha() or x in _MICRO or x in "\u03a9\u2126"):
            break
        out += x
    return out


def _number(tok: str) -> str:
    """숫자 하나: 천 단위 쉼표 무시(5000 = 5,000), 정수와 .0은 같음(5 = 5.0 = 5.00).
    글의 숫자가 수학적으로 같은 값일 때만 맞음: '2.0'은 '2', '2.00'과 맞고 '2.5', '2.05'와는 맞지 않음
    (정수 부분 뒤에 0이 아닌 소수가 이어지면 아님).
    소수 끝 0: 소수 둘째 자리 이상(3.14 = 3.140)은 늘 무시, 소수 한 자리(2.1)는 글에서 단위가 붙을 때만
    (0.5 = '0.50 mm' = '0.50㎜'이지만 2.1 != '2.10 전압': 장절 번호 2.10은 2.1과 다름)."""
    whole, _, frac = tok.replace(",", "").partition(".")
    body = "".join(_char(d) for d in whole)
    if len(whole) > 3:
        # 천 단위 쉼표는 3자리 묶음일 때만 ('1,500' = '1500', '15,00'·'2,1,500'은 아님)
        head = len(whole) % 3 or 3
        grouped = "".join(_char(d) for d in whole[:head]) + "".join(
            "," + "".join(_char(d) for d in whole[i:i + 3]) for i in range(head, len(whole), 3))
        body = rf"(?<!\d,)(?:{body}|{grouped})"
    sig = frac.rstrip("0")
    if len(sig) >= 2:
        return body + r"\." + "".join(_char(d) for d in sig) + "0*"
    if sig:
        unit = rf"(?=\s*[(\[]?\s*(?:{_UNIT_RE}|\u2103|{_SQUARE_UNIT}))"
        if frac != sig:                                 # 검색어가 끝 0을 적음('2.10'): 글의 '2.10', '2.100', 단위 붙은 '2.1 V'
            return body + r"\." + _char(sig) + rf"(?:0+|{unit})"
        # '0.5'는 '0.500'(끝 0 둘 이상)이나 단위 붙은 '0.50 mm'와 맞고, 장절 번호일 수 있는 맨 '2.10'(끝 0 하나)과는 아님
        return body + r"\." + _char(sig) + rf"(?:0{{2,}}|0+{unit})?"
    return body + r"(?:\.0+)?(?!\.0*[1-9])"


def _single_pattern(term: str) -> str:
    """검색어 하나의 정규식 글. 띄어쓰기·줄바꿈·기호가 달라도 찾습니다:
    'SOI 1' = 'SOI #1' = 'SOI-1' = 'SOI1', '설계기준' = '설계 기준', PDF 줄바꿈으로 갈라진 글.
    검색어 끝의 조사는 떼고('정격전압은' = '정격 전압이'), 전각 영문·숫자도 맞고, 숫자 뒤 단위는 띄어쓰기·괄호와 무관
    (5V = 5 V = 5[V] = 5.0 V), °C = ℃.
    숫자 경계는 지킵니다: 'SOI 1'은 'SOI #10'이나 'SOI 21'과, '5V'는 '15V', '5.5V'와 맞지 않습니다.
    기호를 빼면 숫자만 남는 검색어('#4', '1.2')는 기호를 그대로 찾습니다 (모든 '4'와 맞지 않게)."""
    # 전각 -> 반각, ℃ -> °C. 숫자 사이 곱하기는 x로 ('M6×1.0' = 'M6x1.0': 글에서는 x·×·* 모두 맞음)
    orig = re.sub(r"(?<=\d)\s*[\u00d7*]\s*(?=\d)", "x", unicodedata.normalize("NFKC", term)).split()
    stripped = [_strip_word(w) for w in orig]
    ro_cut = {sw for w, sw in zip(orig, stripped) if w.endswith("로") and sw == w[:-1]}   # ㄹ 받침 뒤 '로'를 뗀 낱말
    t = " ".join(stripped)
    compact = re.sub(r"\s+", "", t)
    core = re.sub(r"[\W_]+", "", compact)
    if re.fullmatch(r"\d{1,3}(?:,\d{3})+(?:\.\d+)?|\d+\.\d+", compact):
        # 숫자 하나: 천 단위 쉼표·소수 끝 0과 무관 ('5,000' = '5000', '0.5' = '0.50'), 숫자 경계는 지킴 ('0.5'는 '10.5'와 다름)
        # 소수 한 자리 맨 숫자('2.1')는 장절 번호 경계(section_matches와 같음): '2.1.3'(하위 절)은 맞고 '2.10', '12.1'은 아님
        tail = r"(?!,?\d)" if re.fullmatch(r"\d+\.[1-9]", compact) else _NUM_END
        return r"(?<![\d.])" + _mark("none") + _number(compact) + tail
    tm = re.fullmatch(r"(\d{1,2}):(\d{2})(?::(\d{2}))?", compact)
    if tm and int(tm.group(1)) <= 24 and int(tm.group(2)) < 60:
        # 시각: '08:30' = '8:30' = '8 : 30' ('18:30', '8:300', '8:30:15'(초가 있는 다른 시각)과는 다름)
        sec = rf"\s*:\s*{tm.group(3)}" if tm.group(3) else r"(?!\s*:\s*\d)"
        return rf"(?<![\d.:]){_mark('none')}0?{int(tm.group(1))}\s*:\s*{tm.group(2)}{sec}(?![.,]?\d)"
    dm = re.fullmatch(r"(\d{4})[.\-/](\d{1,2})[.\-/](\d{1,2})\.?", compact)
    if dm and 1 <= int(dm.group(2)) <= 12 and 1 <= int(dm.group(3)) <= 31:
        # 날짜: '2026.11.04' = '2026.11.4' = '2026-11-04' = '2026/11/04' = '2026년 11월 4일'
        y, mo, d = dm.group(1), str(int(dm.group(2))), str(int(dm.group(3)))
        sep = r"\s*[.\-/]\s*"
        return (rf"(?<![\d.]){_mark('none')}{y}(?:{sep}0?{mo}{sep}0?{d}|\s*년\s*0?{mo}\s*월\s*0?{d}\s*일)(?![.,]?\d)")
    sm = re.fullmatch(rf"(±|\+/-|\+-|[+{_MINUS}])?(\d{{1,3}}(?:,\d{{3}})+(?:\.\d+)?|\d+(?:\.\d+)?)(%)?", compact)
    if sm and (sm.group(1) or sm.group(3)):
        # 부호·%가 붙은 숫자 하나: '-20'은 '20', '1-20', 'A-20'과 다름, '±5'는 '±5.5'와 다름, '10%' = '10.0 %'
        cls = _text_sign(compact, sm.start(2))
        head = _SIGN_HEAD.get(cls) or r"(?<![\d.])"
        return head + _mark(cls) + _number(sm.group(2)) + (r"\s*%" if sm.group(3) else _NUM_END)
    if not core or (core.isdigit() and core != compact):
        chars = compact or t
        body = r"\s*".join(_char(c) for c in chars)      # 기호까지 그대로 ('1.2'는 '12'와 다름)
        if chars[0].isdigit():
            body = r"(?<![\d.])" + _mark("none") + body   # '1.2'가 '11.2', '3.1.2', '-1.2'의 끝과 맞지 않게
        if chars[-1].isdigit():
            body += _NUM_END                             # '1/2'가 '1/2.5'와 맞지 않게
        return body
    # 낱말마다 글자 사이에 기호·띄어쓰기를 허용하고, 한글로 끝나는 낱말 뒤에는 글의 조사를 허용 ('절연저항이 100MΩ' = '절연 저항은 100 MΩ')
    raw = [(wm.group(0), wm.start()) for wm in re.finditer(r"\S+", t)
           if any(re.match(r"[^\W_]", x) for x in _NUM_TOKEN.findall(wm.group(0)))]
    raw_words = [w for w, _ in raw]
    # 숫자마다 검색어 안의 앞 부호 (_text_sign: 'Ta -20' -> minus, '0 - 20' -> range, 'ABC-1234' -> hyphen).
    # 다른 낱말 뒤의 부호 없는 숫자는 'after' ('SOI 1'은 'SOI-1'과도 맞음)
    signs: Dict[Tuple[int, int], str] = {}
    words: List[List[str]] = []
    for wi, (w, off) in enumerate(raw):
        kept: List[str] = []
        for m in _NUM_TOKEN.finditer(w):
            if re.match(r"[^\W_]", m.group(0)):
                if m.group(0)[0].isdigit():
                    cls = _text_sign(t, off + m.start())
                    signs[(wi, len(kept))] = "after" if cls == "none" and (kept or words) else cls
                kept.append(m.group(0))
        words.append(kept)
    toks = [x for w in words for x in w]
    particle = "(?:" + "|".join(_PARTICLES) + ")?"
    pieces: List[str] = []
    prev = ""
    for wi, w in enumerate(words):
        parts: List[str] = []
        skip = 0
        for ti, tok in enumerate(w):
            if skip:
                skip -= 1
                continue
            after_num = prev[:1].isdigit()
            run = 0
            while after_num and ti + run < len(w) and not w[ti + run][0].isdigit():
                run += 1
            alias = None
            while run and not alias:                                    # 가장 긴 단위 표기부터 ('100mm이상'의 'mm')
                alias = _unit_alias("".join(w[ti:ti + run]))
                run -= 0 if alias else 1
            if alias:
                parts.append(alias)                                     # 숫자 뒤 단위: 'kΩ' = 'kohm', 'mm' = '㎜'
                skip = run - 1
                prev = w[ti + run - 1]
                continue
            if tok in ("x", "X") and after_num and ti + 1 < len(w) and w[ti + 1][0].isdigit():
                parts.append("[xX\u00d7*]")                             # 숫자 사이 곱하기: 'M6x1.0' = 'M6×1.0'
                prev = tok
                continue
            if tok[0].isdigit():
                cls = signs.get((wi, ti), "none")
                # 맨 앞 숫자의 부호는 패턴에 넣고(가운데 숫자는 낱말 사이 _SEP가 부호를 차지), 일치 뒤 글의 부호를 확인
                # ('-20 ℃'는 '20 ℃', '+20 ℃', '1-20 ℃'와 다름: _SignChecked)
                head = _SIGN_HEAD.get(cls, "") if not parts and not pieces else ""
                parts.append(head + _mark(cls) + _number(tok))
            elif after_num and _caps_unit(_unit_letters(w, ti)):
                # 모두 대문자로 쓴 단위('10 MM', '10UF', '20MA')는 접두 대소문자를 가리지 않음 (층 B 후보, 1.4.70)
                parts.append(f"[{_MICRO}Uu]" if tok == "U" else _char(tok))
            elif after_num and _prefix_char(tok, _unit_letters(w, ti)):
                parts.append(_prefix_char(tok, _unit_letters(w, ti)))  # 단위 접두 대소문자: '10 mHz' != '10 MHz'
            elif tok.upper() == "C" and prev[:1].isdigit():
                parts.append(r"(?:\u2103|[Cc\uff23\uff43])")            # 숫자 뒤 C: 25°C = 25 ℃ = 25C
            else:
                parts.append(_char(tok))
            prev = tok
        piece = _SEP.join(parts)
        if raw_words[wi] in ro_cut:
            piece += particle + r"(?![\uac00-\ud7a3])"                  # '케이블로' -> '케이블'(+조사), '케이블카'는 아님
        elif wi < len(words) - 1 and _HANGUL.match(w[-1]):
            piece += particle
        pieces.append(piece)
    body = _SEP.join(pieces)
    if re.fullmatch(r"[A-Za-z][A-Za-z0-9]*[\-_][A-Za-z0-9\-_]*\d", compact):
        # 영숫자 코드('ABC-1234'): 앞뒤로 영문·숫자가 이어진 다른 코드('XABC-1234', 'ABC-1234-5', 'ABC-12345')와 다름
        body = r"(?<![A-Za-z0-9])" + body + r"(?![\-_]?[0-9])(?![A-Za-z])"
    if toks[0][0].isdigit() and signs.get((0, 0)) not in _SIGN_HEAD:
        body = r"(?<![\d.])" + body                     # 부호('-', '±', '+')가 있으면 앞 경계는 부호 패턴이 봄 ('100±0.5 mm')
    if toks[-1][0].isdigit():
        body += _NUM_END
    else:
        last_num = max((i for i, x in enumerate(toks) if x[0].isdigit()), default=-1)
        tail = toks[last_num + 1:]
        unit = "".join(tail)
        key = _unit_key(unit) or _unit_key("°" + unit)
        if last_num >= 0 and tail and key:
            body += _unit_end(key)                           # 끝이 단위: '10 m'가 '10 mm', '10 MΩ'과 맞지 않게
    return body


# --- 동의어 파일 (<데이터 폴더>/synonyms.txt: 한 줄에 한 묶음, 쉼표로 구분) ---
_SYN_CACHE: Dict[str, Any] = {"key": None, "groups": []}


def _syn_key(term: str) -> str:
    return re.sub(r"[\W_]+", "", strip_particles(unicodedata.normalize("NFKC", term))).lower()


def synonym_groups() -> List[List[str]]:
    """동의어 묶음. 파일이 없거나 읽지 못하면 빈 목록 (검색은 그대로)."""
    path = os.path.join(rt.get_data_dir(), "synonyms.txt")
    try:
        st = os.stat(path)
    except OSError:
        return []
    key = (path, st.st_mtime_ns, st.st_size)
    if _SYN_CACHE["key"] != key:
        groups: List[List[str]] = []
        try:
            for line in rt.read_text_file(path).splitlines():
                line = line.strip()
                if not line or line.startswith("#"):
                    continue
                words = [w.strip() for w in line.split(",") if _syn_key(w)]
                if len(words) > 1:
                    groups.append(words)
        except Exception as e:
            rt.log("WARN", f"동의어 파일을 읽지 못해 쓰지 않음: {type(e).__name__}")
        _SYN_CACHE.update(key=key, groups=groups)
    return _SYN_CACHE["groups"]


def synonyms_of(term: str) -> List[str]:
    """검색어와 같은 묶음의 다른 말 (검색어 자신은 빼고)."""
    k = _syn_key(term)
    out: List[str] = []
    for g in synonym_groups():
        if any(_syn_key(w) == k for w in g):
            out += [w for w in g if _syn_key(w) != k and w not in out]
    return out


@functools.lru_cache(maxsize=512)
def _compiled(term: str, flags: int, syn: Tuple[str, ...]) -> Any:
    """검색어 정규식. 숫자가 들면 글의 숫자 앞 부호를 확인하는 _SignChecked (search·finditer·pattern은 re.Pattern과 같음)."""
    alts = [_single_pattern(t) for t in (term,) + syn]
    pat = re.compile("|".join(f"(?:{a})" for a in alts), flags)
    names = [n for n in pat.groupindex if n in _MARK_CLASS]
    return _SignChecked(pat, names) if names else pat


def _term_pattern(term: str, flags: int) -> "re.Pattern[str]":
    """검색어(와 동의어)를 글 안에서 찾는 정규식 (_single_pattern 규칙)."""
    term = unicodedata.normalize("NFC", term)
    return _compiled(term, flags, tuple(synonyms_of(term)))


# --- 숫자가 든 검색어의 두 층 판정 (1.4.70) ---
# 층 A(literal): 검색어 글자열이 글에 그대로 있고 엄격한 경계를 만족 (띄어쓰기·탭 여러 칸 = 한 칸, 숫자와 단위 사이 띄어쓰기 0~1칸 차이,
# ± 뒤 띄어쓰기만 허용. 대소문자·전각/반각·부호 글자는 그대로). 층 B: 위 정규화 매처(_single_pattern)만 찾은 곳.
# 정규화 규칙을 더 늘리지 않고, 층 B는 결과에 넣되 VERIFIED로 내지 않음 (부호·대시·단위 표기 오일치가 검토마다 새로 나왔음)
_LIT_SP = "[ \t\u00a0]"
_LIT_SIGNS = "-+\u00b1\u2213\u2010\u2011\u2012\u2013\u2014\u2015\u2212\ufe63\uff0d\uff0b\u02d7\u207a\u207b"
_LIT_SUPER = "\u00b2\u00b3\u00b9" + "".join(chr(c) for c in range(0x2070, 0x20a0))
_LIT_LETTER_AFTER = "\u03a9\u2126\u03bc\u00b5\u00b0\u2103\u2109"     # 단위 끝 뒤에 붙으면 다른 단위인 글자 (영문 밖)
_LIT_SIGN_WORD = re.compile(r"(?:영하|마이너스|minus|negative)" + _LIT_SP + r"*$", re.IGNORECASE)
_LIT_PAREN_SIGN = re.compile(r"\(\s*[" + re.escape(_LIT_SIGNS) + r"]\s*\)" + _LIT_SP + r"*$")
# 부호 없는 검색어 숫자 앞을 띄어쓰기·'/' 건너 보면 ±·+ ('± 5 V', '+ 20', '+ / - 5 V', '+/-  10 %').
# 빼기 하나만 띄어 쓴 것('0 - 5 V', 표 빈 칸 '-\t40')은 허용
_LIT_PM_GAP = re.compile(r"[+\u00b1\u2213\uff0b][ \t\u00a0/]*(?:[" + re.escape(_LIT_SIGNS) + r"])?[ \t\u00a0/]*$")
# 빼기로 시작하는 검색어 앞에 띄어 쓴 다른 값(숫자+단위)이나 'to'가 있으면 범위 대시인지 부호인지 애매
_LIT_RANGE_LEFT = re.compile(r"(?:\d(?:" + _LIT_SP + r"?[^\s\d]{1,4})?|(?<![A-Za-z])to)" + _LIT_SP + r"+$", re.IGNORECASE)
# 단위 끝 뒤의 뜻이 다른 꼬리: '(rms)', '(g)', '(A)', ' rms', ' p-p', ' peak', ' pk', ' abs', ' g'(게이지).
# 괄호는 이 목록만 막음: '(DC)', '(AC)', '(Max)', '(Typ.)'처럼 값의 뜻을 바꾸지 않는 괄호는 층 A (1.4.71)
_LIT_TAIL = re.compile(_LIT_SP + r"*[(\[](?:rms|p-p|pp|peak|pk|abs|g|a|gauge)[)\]]|"
                       + _LIT_SP + r"+(?:rms|p-p|pp|peak|pk|abs|g)(?![A-Za-z])", re.IGNORECASE)
# 띄어 쓴 단위 이어짐 (낱말이 다른 단위): '5 N m', '10 Ω cm', '10 kV A', '10 mm Hg', '2.5 mm SQ', '5 % RH', '3 dB μV',
# 띄어 쓴 '/단위'('25 ℃ /W'), 띄어 쓴 위첨자('10 m ³'). 대소문자 그대로 (1.4.71)
_LIT_UNIT_GO = re.compile("[ \u00a0]{1,2}(?:(?:H2O|Hg|Aq|SQ|R\\.H\\.?|RH|FS|f|[\u03bc\u00b5u]V|cm|mm|km|m|h|s|A|V|W|\u03a9|\u2126)"
                          r"(?![A-Za-z0-9])|/[ \u00a0]*[^\s\d+\-\u00b1\u2212]|[" + _LIT_SUPER + "])")
_LIT_UNIT_OP = "\u00b7\u22c5\u2219\u2022\u2217*"                   # 곱 단위 연산자 (가운뎃점, U+22C5, U+2219, U+2022, U+2217)
_LIT_THOUSANDS = re.compile("(?<![\\d.,])\\d{1,3}[ \u00a0\u2009\u202f]$")   # 띄어 쓴 천 단위의 앞 묶음 ('1 500 W')
# 붙여 쓴 것과 같은 공백 ('-\u00a040 ℃'): 붙는 공백, 폭 없는 공백(U+200B~200D, U+2060, U+FEFF), 전각 공백(U+3000)
_LIT_GLUED_SPACE = "\u00a0\u2009\u202f\u200b\u200c\u200d\u2060\ufeff\u3000"
# 숫자 바로 앞에 붙으면 다른 값의 일부인 글자: 곱하기('1\u00d72 mm', '1*2 mm'), 대시로 쓴 한글 모음(U+3161)
_LIT_GLUED_BEFORE = "\u00d7*\u3161"
# 빗금 글자: '/', 전각 U+FF0F, 나눗셈 U+2215, 분수 U+2044. 값 구분으로 볼 수 있는 것은 '/'뿐 (나머지는 늘 단위·분수)
_LIT_SLASHES = "/\uff0f\u2215\u2044"
# '/' 값 구분 (1.4.72): 양쪽이 '숫자+단위'('100 V/220 V', '85 ℃/85 %RH')이고 '/' 뒤 값이 1·10·100·1000(1·10^n)이 아닐 때만.
# '/' 뒤 값이 1·10^n이면 '당'(비율: '0.1 mm/1000 mm', '2 ℃/1 min', '5 %/1000 h'). 앞이 숫자뿐이면 주파수·전압만
# ('50/60 Hz', '10/50 kHz', '110/220 V'. '10/350 μs', '15/16 inch', '1.2/50 μs'는 서지 파형·분수)
# 숫자 바로 앞에 붙어도 층 A인 여는 글자 (1.4.73 허용 목록): 괄호, 범위 물결(~, U+FF5E, U+301C), '='
_LIT_OPEN_BEFORE = "([~\uff5e\u301c="
# 여는 따옴표·괄호 (1.4.75, 앞이 숫자가 아닐 때만: '1\u2018500 V'의 천 단위와 구분): U+201C, U+2018, U+300C, U+300E,
# U+3010, U+3008, U+300A. ASCII 따옴표는 짝일 때만 ('"24 V"', 짝 없는 '"5 V'·'5 V\''는 층 B)
_LIT_QUOTE_OPEN = "\"'"
_LIT_OPEN_PAIR = "\u201c\u2018\u300c\u300e\u3010\u3008\u300a"
# 단위 바로 뒤에 붙어도 층 A인 닫는 짝·구분 (1.4.75): 닫는 따옴표·괄호(U+201D, U+2019, U+300D, U+300F, U+3011, U+3009,
# U+300B), 쉼표 U+3001, 곱하기 U+00D7('100 mm×폭': 숫자 앞의 ×만 막음). ASCII 따옴표는 위처럼 짝일 때만
_LIT_CLOSE_AFTER = "\u201d\u2019\u300d\u300f\u3011\u3009\u300b\u3001\u00d7"
# 단위 뒤 괄호 중 값의 뜻이 바뀌는 것 ('(rms)' 등은 _LIT_TAIL): 부호만 '5 V(-)', 숫자만 '5 V(1)'(각주·지수),
# 한글 실효값·피크 등 '5 V(실효값)'
_LIT_PAREN_BAD_AFTER = re.compile(r"\(\s*(?:[" + re.escape(_LIT_SIGNS) + r"]|\d+|실효|피크|첨두|게이지|절대|평균)\s*(?:값)?\s*\)")
# 단위 바로 뒤에 붙은 괄호 중 값의 뜻을 바꾸지 않는 것 (허용 목록): '5 V(DC)', '10 A(Max)', '3.3 V(Typ.)'
_LIT_PAREN_OK = re.compile(r"[(\[](?:DC|AC|max\.?|min\.?|typ\.?|nom\.?|ref\.?|정격|최대|최소|이하|이상|미만|초과)[)\]]",
                           re.IGNORECASE)
# 단위 뒤에 붙여 쓴 범위 대시 뒤의 다른 값 ('10 m-20 m', '40 ℃\u201385 ℃'). '10 m-1'(지수)은 아님
_LIT_RANGE_RIGHT = re.compile("[\\d.,]+[ \u00a0]?[A-Za-z%\u00b0\u2103\u2109\u03a9\u2126\u03bc\u00b5]")
_LIT_SLASH_LEFT_UNIT = re.compile(r"(?<![\d.,])\d[\d.,]*[ \u00a0]?[^\s\d/\uff0f\u2215\u2044]{1,4}/$")
_LIT_SLASH_LEFT_NUM = re.compile(r"(?<![\d.,])\d{2,}/$")
_LIT_SLASH_NUM_UNITS = ("Hz", "kHz", "MHz", "GHz", "V", "kV", "mV")
_LIT_SLASH_RIGHT = re.compile(r"[+\-\u00b1\u2212]?(\d[\d.,]*)[ \u00a0]?[^\s\d/\uff0f\u2215\u2044]")
_LIT_PER = re.compile(r"10*(?:[.,]0+)?")                          # 1, 10, 100, 1000, 1.0 (비율의 '당' 값)
# 붙여 쓴 범위 대시 앞의 '숫자+기호 단위'는 _GLUED_RANGE_LEFT ('40 ℃\u201385 ℃'의 '40 ℃'): 영문 단위('5 V-10 V')는 하이픈 코드와 구분할 수 없어 뺌


def _lit_slash_left(text: str, s: int, q: str) -> bool:
    """text[s-1]이 '/'일 때 그 앞이 값 구분인지 ('100 V/220 V'의 '220 V', '50/60 Hz'의 '60 Hz')."""
    head = text[max(0, s - 16):s]
    num = re.match(r"[\d.,]+", q)
    if num is None:
        return False
    if _LIT_SLASH_LEFT_UNIT.search(head):
        return not _LIT_PER.fullmatch(num.group(0))
    unit = q[num.end():].strip(" \t\u00a0")
    # 앞이 숫자뿐이면 분수('15/16 V', 분모 2·4·8·16·32·64)와 비율('10/100 V', 1·10^n)은 값 구분이 아님 (10차 검토)
    value = num.group(0)
    if _LIT_PER.fullmatch(value) or value in ("2", "4", "8", "16", "32", "64"):
        return False
    return bool(_LIT_SLASH_LEFT_NUM.search(head)) and unit in _LIT_SLASH_NUM_UNITS


def _lit_slash_right(text: str, e: int) -> bool:
    """text[e]가 '/'일 때 그 뒤가 다른 값(부호·숫자+단위, 1·10^n 아님)인지 ('100 V/220 V', '+10 %/-5 %')."""
    m = _LIT_SLASH_RIGHT.match(text, e + 1)
    return bool(m) and not _LIT_PER.fullmatch(m.group(1))


def _lit_space(c: str) -> bool:
    """공백류 (띄어쓰기·탭·줄바꿈·붙는 공백·폭 없는 공백·전각 공백). 빈 글자(줄 시작·끝)는 아님."""
    return bool(c) and (c.isspace() or c in _LIT_GLUED_SPACE)


def _hangul(c: str) -> bool:
    """한글 음절 (자모 U+3161 'ㅡ'는 대시로 쓰여 빠짐)."""
    return bool(c) and "가" <= c <= "힣"


def _latin(c: str) -> bool:
    return bool(c) and (("a" <= c.lower() <= "z") or "\uff21" <= c <= "\uff3a" or "\uff41" <= c <= "\uff5a")


def _unit_start(c: str) -> bool:
    """숫자 바로 뒤에 오면 단위의 시작으로 보는 글자 (그 사이 띄어쓰기 0~1칸 차이는 층 A)."""
    return c.isalpha() or c in "%\u00b0\u2103\u2109\u2126\u03a9\u00b5\u03bc\u2113" or "\u3380" <= c <= "\u33df"


class _Literal:
    """층 A 찾기: 검색어 글자열(띄어쓰기만 느슨하게) + 앞뒤 엄격한 경계. search(text)는 경계를 만족한 첫 일치(없으면 None)."""

    def __init__(self, terms: List[str]):
        self.items = []
        for t in terms:
            toks = re.findall(r"[ \t\u00a0\u3000]+|.", t.strip(), re.S)
            out: List[str] = []
            prev = ""
            for k, tok in enumerate(toks):
                if tok[0] in " \t\u00a0\u3000":
                    nxt = toks[k + 1][0] if k + 1 < len(toks) else ""
                    # 영문 접두와 숫자 사이도 0~1칸 차이 허용 ('IP 67' = 'IP67', 'M 10' = 'M10', 1.4.71)
                    loose = ((prev.isdecimal() and _unit_start(nxt)) or prev in "\u00b1\u2213" or nxt in "\u00b1\u2213"
                             or (_latin(prev) and nxt.isdecimal()))
                    out.append(_LIT_SP + ("*" if loose else "+"))
                else:
                    # 숫자와 단위 사이, ± 앞뒤는 띄어쓰기 0~1칸 차이 허용 ('5V' = '5 V', '5V ±5%' = '5V±5%')
                    if prev not in ("", " ") and ((prev.isdecimal() and _unit_start(tok)) or tok in "\u00b1\u2213"
                                                  or (prev in "\u00b1\u2213" and tok.isdecimal())):
                        out.append(_LIT_SP + "*")
                    elif _latin(prev) and tok.isdecimal():
                        out.append(_LIT_SP + "?")
                    out.append(re.escape(tok))
                prev = tok[-1] if tok[0] not in " \t\u00a0\u3000" else " "
            if out:
                self.items.append((re.compile("".join(out)), t.strip()))

    @staticmethod
    def _ok(text: str, s: int, e: int, q: str) -> bool:
        first, last = q[0], q[-1]
        before = text[s - 1] if s else ""
        if before:
            if first.isdecimal():
                # 'DC24 V'의 '24 V'는 층 A, 단위 없는 검색어 '12'·'100'은 'AC12 단자대'·'AC100'과 다름 (1.4.72)
                dc_ac = (not last.isdecimal() and _latin(before) and text[max(0, s - 2):s] in ("DC", "AC")
                         and not _latin(text[s - 3:s - 2]))
                slash = before == "/" and not last.isdecimal() and _lit_slash_left(text, s, q)
                glued_range = before in "-\u2013" and _GLUED_RANGE_LEFT.search(text, max(0, s - 16), s - 1)
                # 허용 목록 (1.4.73): 공백류, 줄 시작, 여는 괄호, 범위 물결, '=', ':'(앞이 숫자가 아닐 때), 한글, 'DC'/'AC',
                # 값 구분 '/', 붙여 쓴 범위 대시만. 그 밖 글자(문장부호·기호·숫자·영문: '1，500 V', '1．5 V', '1·5 V',
                # '△40 ℃', '1:5 V', '1÷5 V', '½5 A')는 층 B. 금지 목록은 검토마다 새 글자가 나왔음 (10차 검토)
                # 1.4.75: 여는 따옴표·괄호, 전각 콜론 U+FF1A, 목록 머리 U+2022(줄 시작·공백 뒤)도 허용
                prev2 = text[s - 2:s - 1]
                if not (_lit_space(before) or before in _LIT_OPEN_BEFORE or _hangul(before) or dc_ac or slash
                        or glued_range or (before in ":\uff1a" and not prev2.isdecimal())
                        or (before in _LIT_OPEN_PAIR and not prev2.isdecimal())
                        or (before in _LIT_QUOTE_OPEN and text[e:e + 1] == before)
                        or (before == "\u2022" and (not prev2 or _lit_space(prev2)))):
                    return False                                   # 'DC5V'의 '5V', '50/60 Hz'의 '60 Hz'는 층 A
                if before in _LIT_GLUED_SPACE and text[s - 2:s - 1] in _LIT_SIGNS:
                    return False                                   # '-\u00a040 ℃': 붙여 쓴 빼기와 같음
                if before == "\t" and text[s - 2:s - 1] in _LIT_SIGNS and s >= 3 and text[s - 3] not in "\t\r\n":
                    return False                                   # '온도 -\t40 ℃': 같은 칸의 빼기 (표 빈 칸 '\t-\t40'은 허용)
                if re.match(r"\d{3}(?!\d)", q) and _LIT_THOUSANDS.search(text, max(0, s - 6), s):
                    return False                                   # '1 500 W', '10 000 h': 띄어 쓴 천 단위
                if _LIT_SIGN_WORD.search(text[max(0, s - 20):s]) or _LIT_PAREN_SIGN.search(text[max(0, s - 12):s]):
                    return False                                   # '영하 40 ℃', '(-)5 V'
                if _LIT_PM_GAP.search(text[max(0, s - 24):s]):
                    return False                                   # '±  0.5 mm', '+ / - 5 V', '+ 20' (띄어 쓴 ±·+)
            elif first in _LIT_SIGNS:
                # 검색어의 부호 앞: 다른 부호·영문은 아님, 숫자는 ±에만 ('100±0.5 mm'), '/'·','는 허용 ('+5 V/-5 V')
                if before in _LIT_SIGNS or before in ".^" or _latin(before) or (before.isdecimal() and first not in "\u00b1\u2213"):
                    return False
                if first not in "+\u00b1\u2213\uff0b" and (_LIT_RANGE_LEFT.search(text[max(0, s - 24):s])
                                                           or _LIT_PM_GAP.search(text[max(0, s - 24):s])):
                    return False                                   # '25 -40 ℃', '+0.2 mm -0.1 mm', '25 to -40', '+/-5'
            elif _latin(first) and (before.isdecimal() or _latin(before)):
                return False
        after = text[e] if e < len(text) else ""
        if not after:
            return True
        nxt = text[e + 1] if e + 1 < len(text) else ""
        if last.isdecimal():
            if _latin(first) and after in "-_" and nxt.isalnum():
                return False                                       # 영숫자 코드 'ABC-1234'는 'ABC-1234-5'와 다름
            if after == "-" and nxt.isdecimal():
                return False                                       # 숫자 코드 '61000-4'는 'IEC 61000-4-2', '1234'는 '1234-5'와 다름
            return not (after.isdecimal() or (after in ".," and nxt.isdecimal()) or _latin(after)
                        or after in "^*%" or after in _LIT_SLASHES or after in _LIT_SUPER)
        if _unit_start(last):
            if _latin(last) and text[e:e + 2] in ("DC", "AC") and not _latin(text[e + 2:e + 3]):
                return True                                        # '24 VDC'의 '24 V' (단위에 붙은 DC/AC)
            if after in _LIT_SLASHES:
                # '100 V/220 V', '85 ℃/85 %RH'는 값 구분, '5 V/m', '0.1 mm/1000 mm', '10 V\uff0fm'은 단위·비율 (1.4.72)
                return after == "/" and _lit_slash_right(text, e)
            # 허용 목록 (1.4.73): 공백류, 줄 끝, 닫는 괄호, ',' '.'(뒤가 숫자·영문이 아닐 때), ';', ':'(뒤가 공백류·줄 끝),
            # 범위 물결, 한글, 붙여 쓴 괄호의 허용 목록('(DC)', '(Max)', '(Typ.)'), 범위 대시('10 m-20 m').
            # 그 밖 글자('5 V㎃', '5 VA', '5 V·s', '5 V′')는 층 B (10차 검토)
            # 1.4.75: 닫는 따옴표·괄호, U+3001, 곱하기 U+00D7, '±' 뒤 숫자('220 V±10 %'), 각주 '*1', 괄호 '('(막는 꼬리
            # '(rms)' '(g)' '(A)'와 부호 괄호 '(-)' 밖: '220 V(60 Hz)', '1 MPa(10 kgf/cm²)', '2 h(1회)', '24 V(±10 %)')
            if not (_lit_space(after) or after in ")];~\uff5e\u301c" or _hangul(after)
                    or (after in ",." and not (nxt.isdecimal() or _latin(nxt)))
                    or (after in ":\uff1a" and (not nxt or _lit_space(nxt)))
                    or (after == "[" and _LIT_PAREN_OK.match(text, e))
                    or (after == "(" and not _LIT_PAREN_BAD_AFTER.match(text, e))
                    or after in _LIT_CLOSE_AFTER or (after in _LIT_QUOTE_OPEN and text[s - 1:s] == after)
                    or (after == "\u00b1" and nxt.isdecimal())
                    or (after == "*" and re.match(r"\*\d{1,2}(?![\d.,A-Za-z])", text[e:e + 4]))
                    or (after in "-\u2013\u2212" and _LIT_RANGE_RIGHT.match(text, e + 1))):
                return False
            if after in "-\u2212\u207b" and nxt.isdecimal() and not re.match(
                    "[\\d.,]+[ \u00a0]?[A-Za-z%\u00b0\u2103\u2109\u03a9\u2126\u03bc\u00b5]", text[e + 1:e + 12]):
                return False                                       # '10 m-1': 단위의 지수 (범위 '10 m-20 m'은 아님)
            if after.isdecimal() or _latin(after) or after in _LIT_LETTER_AFTER or after == "^" or after in _LIT_SUPER:
                return False
            if after == "%" and last != "%":
                return False
            if after in "-." and _latin(nxt):
                return False                                       # '5 kgf-cm'
            if after in _LIT_UNIT_OP and (_latin(nxt) or nxt in _LIT_LETTER_AFTER):
                return False                                       # '10 N·m', '5 N(U+22C5)m'
            return not (_LIT_TAIL.match(text, e) or _LIT_UNIT_GO.match(text, e))
        return True

    def finditer(self, text: str):
        """경계를 만족한 일치를 모두 (검색어 묶음 순서대로, 같은 묶음 안에서는 앞에서부터)."""
        for pat, q in self.items:
            pos = 0
            while True:
                m = pat.search(text, pos)
                if m is None:
                    break
                if self._ok(text, m.start(), m.end(), q):
                    yield m
                pos = m.start() + 1

    def search(self, text: str) -> Optional["re.Match[str]"]:
        return next(self.finditer(text), None)


@functools.lru_cache(maxsize=512)
def _literal_compiled(term: str, syn: Tuple[str, ...]) -> Optional[_Literal]:
    if not re.search(r"\d", term):
        return None
    return _Literal([term] + list(syn))


def literal_finder(term: str) -> Optional[_Literal]:
    """숫자가 든 검색어의 층 A 찾기 객체 (숫자가 없는 검색어는 None: 지금 동작 그대로)."""
    term = unicodedata.normalize("NFC", term)
    return _literal_compiled(term, tuple(synonyms_of(term)))


def literal_hit(text: str, terms: List[str]) -> bool:
    """맞은 검색어 묶음의 숫자가 든 검색어가 모두 글에 글자 그대로 있는지 (층 A). 숫자가 든 검색어가 없으면 True."""
    t = unicodedata.normalize("NFC", text)
    return all(f.search(t) is not None for f in (literal_finder(x) for x in terms) if f is not None)


def term_span(term: str, text: str, flags: int = re.IGNORECASE, loose_text: Optional[str] = None) -> Optional[Tuple[int, int]]:
    """검색어가 글에서 맞은 자리 (시작, 끝). 숫자가 든 검색어는 층 A(글자 그대로) 자리를 먼저, 없으면 층 B(정규화 매처) 자리.
    loose_text: 층 B를 찾을 글 (글자 수가 같은 변환본, 예: 빼기 대시를 '-'로 바꾼 문장). 없으면 None (1.4.71)"""
    lf = literal_finder(term)
    m = lf.search(text) if lf is not None else None
    if m is None:
        m = _term_pattern(term, flags).search(text if loose_text is None else loose_text)
    return (m.start(), m.end()) if m else None


# 숫자 검색어의 VERIFIED(층 A)는 '검색어 글자열이 그 자리에 글자 그대로 있다'까지만 확인 (앞뒤 부호·범위·이하/이상의 뜻은 판정하지 않음)
NUM_CONTEXT_NOTE = "결과를 알릴 때 snippet의 앞뒤(부호, 범위 ~, 이하/이상)를 그대로 함께 전달하세요."
SNIPPET_MIN_BEFORE = 15        # 스니펫에 늘 담는 일치 앞 글자 수 (부호 '- 40 ℃', 범위 '-10 ~', '이하'가 보이게)


SNIPPET_JOIN = " / "          # 검색어 자리가 멀어 창이 겹치지 않을 때 조각 사이 (1.4.72)


def _highlight(text: str, start: int, end: int, patterns: Tuple[Any, Any], lit_spans: List[Tuple[int, int]]) -> str:
    """text[start:end]에서 맞은 곳을 강조. 층 A 자리(숫자가 든 검색어의 글자 그대로 자리)와 숫자 없는 검색어는 **강조**,
    숫자가 든 검색어의 층 B 자리(표기가 다른 곳, 정규화 매처만 찾음)는 [[강조]] (1.4.73)."""
    piece = text[start:end]
    plain, numeric = patterns
    lits = [(a - start, b - start) for a, b in lit_spans if a >= start and b <= end]
    marks = [(a, b, "A") for a, b in lits]
    if plain is not None:
        marks += [(m.start(), m.end(), "A") for m in plain.finditer(piece) if m.group(0)]
    if numeric is not None:
        marks += [(m.start(), m.end(), "B") for m in numeric.finditer(piece)
                  if m.group(0) and not any(a < m.end() and m.start() < b for a, b in lits)]
    merged: List[List[Any]] = []
    for a, b, kind in sorted(marks):
        if merged and a <= merged[-1][1]:
            merged[-1][1] = max(merged[-1][1], b)
            if kind == "A":
                merged[-1][2] = "A"
        else:
            merged.append([a, b, kind])
    for a, b, kind in reversed(merged):
        o, c = ("**", "**") if kind == "A" else ("[[", "]]")
        piece = piece[:a] + o + piece[a:b] + c + piece[b:]
    return ("..." if start > 0 else "") + piece + ("..." if end < len(text) else "")


def make_snippet(text: str, terms: List[str], context: int) -> str:
    """검색어마다 일치 자리(숫자가 든 검색어는 층 A 자리 먼저)와 앞뒤 문맥. 앞은 최소 SNIPPET_MIN_BEFORE자. 맞은 곳은 **강조**
    (숫자가 든 검색어의 층 B 자리는 [[강조]]).
    AND 검색어들이 멀어 창이 겹치지 않으면 조각을 SNIPPET_JOIN으로 이음 ('...**온도**... / ...0 - **40 ℃**에서...')"""
    flags = re.IGNORECASE
    spans = sorted({sp for sp in (term_span(t, text, flags) for t in terms) if sp})
    if not spans:
        return text[: context * 2]
    before = max(context, SNIPPET_MIN_BEFORE)
    wins: List[List[int]] = []
    for s0, e0 in spans:
        a, b = max(0, s0 - before), min(len(text), e0 + context)
        if wins and a <= wins[-1][1]:
            wins[-1][1] = max(wins[-1][1], b)
        else:
            wins.append([a, b])
    def joined(group: List[str]) -> Any:
        return re.compile("|".join(_term_pattern(t, 0).pattern for t in sorted(group, key=len, reverse=True)), flags) if group else None
    patterns = (joined([t for t in terms if not re.search(r"\d", t)]), joined([t for t in terms if re.search(r"\d", t)]))
    lit_spans = [(m.start(), m.end()) for lf in (literal_finder(t) for t in terms) if lf is not None for m in lf.finditer(text)]
    return SNIPPET_JOIN.join(_highlight(text, a, b, patterns, lit_spans) for a, b in wins).replace("\n", " ")


_SENT_SPLIT = re.compile(r"(?<=[.!?\u3002])\s+|\n+")


def match_sentence(text: str, terms: List[str], limit: int = 1000) -> str:
    """일치한 단어가 든 문장 (문장으로 못 나누면 단락 앞부분)."""
    flags = re.IGNORECASE
    sents = _SENT_SPLIT.split(text)
    lits = [f for f in (literal_finder(t) for t in terms) if f is not None]
    for sent in sents if lits else ():
        if any(f.search(sent) for f in lits):
            return sent.strip()[:limit]                     # 층 A(글자 그대로) 자리가 든 문장을 먼저 (1.4.71)
    for sent in sents:
        if any(term_span(t, sent, flags) for t in terms):
            return sent.strip()[:limit]
    return text.strip()[:limit]


# --- 검색 넓히기 (0건일 때: 일부 검색어 -> 유사 낱말)와 순위 ---

FUZZY_MIN_LATIN = 4      # 유사 일치(한 글자 다름)로 찾는 검색어의 최소 글자 수 (영문·숫자)
FUZZY_MIN_HANGUL = 4     # 한글은 음절 수. 첫·끝 음절이 같을 때만 (전압/전류, 대기 전류/대기 전력, 최대/최소가 서로 맞지 않게)


def _edit1(a: str, b: str) -> bool:
    """두 글이 한 글자만 다른지 (바꿈·넣음·뺌 하나, 같으면 False)."""
    if a == b or abs(len(a) - len(b)) > 1:
        return False
    if len(a) == len(b):
        return sum(x != y for x, y in zip(a, b)) == 1
    s, l = (a, b) if len(a) < len(b) else (b, a)
    i = 0
    while i < len(s) and s[i] == l[i]:
        i += 1
    return s[i:] == l[i + 1:]


def _fuzzy_core(term: str) -> Optional[str]:
    """유사 일치로 찾을 수 있는 검색어면 비교용 글 (기호·띄어쓰기·조사 뺌, 소문자), 아니면 None."""
    core = _syn_key(term)
    if not core or any(ch.isdigit() for ch in core):
        return None                       # 숫자가 든 검색어(규격 번호, 값)는 한 글자 달라도 다른 것
    hangul = len(_HANGUL.findall(core))
    if hangul:
        return core if hangul >= FUZZY_MIN_HANGUL else None
    return core if len(core) >= FUZZY_MIN_LATIN else None


def _jamo(ch: str) -> Optional[Tuple[int, int, int]]:
    """한글 음절 -> (초성, 중성, 종성) 번호, 음절이 아니면 None."""
    code = ord(ch) - 0xAC00
    if not 0 <= code < 11172:
        return None
    return code // 588, (code % 588) // 28, code % 28


def _korean_typo(a: str, b: str) -> bool:
    """같은 길이의 두 글에서 다른 글자가 한글 음절이면 자모 하나만 다를 때만 오타로 봄 (젼압/전압, 케이불/케이블).
    음절이 통째로 다르면(대/소, 류/력, 상/하) 다른 말이라 맞추지 않음."""
    for x, y in zip(a, b):
        if x != y:
            jx, jy = _jamo(x), _jamo(y)
            if jx is None or jy is None:
                return True
            return sum(p != q for p, q in zip(jx, jy)) == 1
    return True


def _fuzzy_word(text: str, core: str) -> Optional[str]:
    """글에서 core와 한 글자만 다른 낱말(또는 이어진 두 낱말, 뒤 조사는 무시).
    한글이 든 검색어는 첫 글자와 끝 글자가 같고, 바뀐 음절은 자모 하나만 달라야 함: 한 음절 차이로 뜻이 반대·다른 말
    (전류/전력, 최대/최소, 상한/하한)이 맞지 않게."""
    words = re.findall(r"[^\W_]+", unicodedata.normalize("NFKC", text))
    cands = words + [a + b for a, b in zip(words, words[1:])]
    n = len(core)
    hangul = bool(_HANGUL.search(core))
    for w in cands:
        low = w.lower()
        for k in (n, n - 1, n + 1):
            if 0 < k <= len(low) and _edit1(core, low[:k]):
                if hangul and (low[0] != core[0] or low[k - 1] != core[-1]
                               or (k == n and not _korean_typo(core, low[:k]))):
                    continue
                return w[:k]
    return None


NUMBER_GAP = 40                # 숫자 안전망: 검색어의 숫자 사이에 끼어도 되는 글자 수
NUMBER_SEARCH_SECONDS = 20.0   # 숫자 안전망 전체의 시간 상한 (넘으면 남은 단위는 보지 않고 결과에 알림)


class _LooseSpan:
    """_LooseNumbers.search의 결과 (re.Match의 start·end만)."""

    def __init__(self, s: int, e: int):
        self._s, self._e = s, e

    def start(self) -> int:
        return self._s

    def end(self) -> int:
        return self._e


class _LooseNumbers:
    """숫자 안전망: 검색어의 숫자 값만 차례로 찾음 (부호·단위·붙여쓰기·기호 무시, 숫자 경계는 지킴: '5'는 '15', '5.5'와 다름).
    숫자 사이에는 NUMBER_GAP자까지 무엇이든 허용. 숫자마다 위치 목록을 만든 뒤 앞 숫자의 끝에서 NUMBER_GAP자 안에
    다음 숫자가 시작하는 사슬을 선형으로 찾음 (하나의 큰 정규식은 '1 1 1 ...' 같은 글에서 지수 시간이 걸렸음, 1.4.69)."""

    def __init__(self, pieces: List[str]):
        self.pats = [re.compile(p, re.IGNORECASE) for p in pieces]

    def search(self, text: str) -> Optional[_LooseSpan]:
        # reach: 지금 숫자까지 사슬이 닿은 일치들 [(끝, 사슬 시작)], 끝 위치 오름차순
        reach: List[Tuple[int, int]] = [(m.end(), m.start()) for m in self.pats[0].finditer(text)]
        for pat in self.pats[1:]:
            if not reach:
                return None
            ends = [e for e, _ in reach]
            nxt: List[Tuple[int, int]] = []
            for m in pat.finditer(text):
                s = m.start()
                hi = bisect.bisect_right(ends, s) - 1          # 앞 숫자의 끝이 s 이하인 마지막 일치
                if hi >= 0 and ends[hi] >= s - NUMBER_GAP:
                    nxt.append((m.end(), reach[hi][1]))
            reach = nxt
        if not reach:
            return None
        s, e = min((s0, e) for e, s0 in reach)
        return _LooseSpan(s, e)


def _loose_number_pattern(term: str) -> Optional[_LooseNumbers]:
    """숫자 안전망의 찾기 객체 (_LooseNumbers). 숫자가 없으면 None."""
    t = unicodedata.normalize("NFKC", term)
    if not re.search(r"\d", t):
        return None
    if re.search(r"\d[.:]\d+[.:\-/]\d", t):
        nums = re.findall(r"\d+", t)                          # 날짜·시각: 마디마다 ('2026.11.04', '08:30')
    else:
        nums = re.findall(r"\d{1,3}(?:,\d{3})+(?:\.\d+)?|\d+(?:\.\d+)?", t)
    pieces = []
    for n in nums[:8]:
        whole, dot, frac = n.partition(".")
        whole = whole.lstrip("0") if whole.replace(",", "").strip("0") else "0"
        pieces.append(r"(?<![\d.])(?<!\d,)0*" + _number(whole + dot + frac) + r"(?![.,]?\d)")
    return _LooseNumbers(pieces)


def _need_words(term: str) -> List[str]:
    """숫자 안전망에서도 같은 단위 안에 있어야 할 검색어 낱말: 숫자가 없고 글자가 든 낱말 중 단위('V', 'mm', '℃')가 아닌 것
    ('1 퀀텀'의 '퀀텀', '정격 전압 5 V'의 '정격', '전압')."""
    out = []
    for w in unicodedata.normalize("NFKC", term).split():
        if re.search(r"\d", w) or not re.search(r"[^\W\d_]", w):
            continue
        core = re.sub(r"[\W_]+", "", w)
        if _unit_key(core) or _unit_key("°" + core) or core.lower() in ("c", "rms") \
                or core.lower().endswith(("ohm", chr(0x3c9))):          # 'mohm', 'kΩ'(소문자로 ω)도 단위
            continue
        out.append(w)
    return out


def number_search(units: List[Tuple[Dict[str, Any], Dict[str, Any]]], groups: List[List[str]], flags: int,
                  status: Optional[Dict[str, Any]] = None) -> List[Dict[str, Any]]:
    """정확히 일치한 곳이 0건일 때의 숫자 안전망: 숫자가 든 검색어는 숫자 값만 같은 곳, 나머지 검색어는 정확히 일치.
    숫자 경계·부호 규칙이 엄격해 맞는 표기를 놓쳐도 '0건(규정 없음)'으로 단정하지 않게 후보로 줌.
    숫자가 든 검색어의 숫자 아닌 낱말(단위 빼고)도 같은 단위 안에 있어야 함 ('1 퀀텀'은 '퀀텀'이 없는 곳의 '1'과 맞지 않음).
    NUMBER_SEARCH_SECONDS를 넘으면 남은 단위는 보지 않고 status['cut'] = True.
    반환 [{doc, unit, terms(강조할 글), matched_terms, missing_terms, number_words}] (숫자가 든 검색어가 없으면 빈 목록)."""
    loose = {t: _loose_number_pattern(t) for g in groups for t in g}
    if not any(loose.values()):
        return []
    pats = {t: _term_pattern(t, flags) for g in groups for t in g}
    need = {t: [_term_pattern(w, flags) for w in _need_words(t)] for t in loose if loose[t]}
    words = {t: " ".join(_need_words(t)) for t in loose}
    rows: List[Dict[str, Any]] = []
    sign_chars = "\u00b1+/ \t" + _DASHES
    deadline = time.monotonic() + NUMBER_SEARCH_SECONDS
    for i, (doc, u) in enumerate(units):
        if time.monotonic() > deadline:              # 단위마다 확인 (큰 단위 50개를 몰아 보면 상한을 크게 넘었음, 1.4.70)
            if status is not None:
                status["cut"] = True
                status["checked"] = i
            break
        text = unicodedata.normalize("NFC", u["text"])
        for g in groups:
            if not any(loose[t] for t in g):
                continue
            exact, nw, extra = [], {}, []
            for term in g:
                if loose[term]:
                    if not all(p.search(text) for p in need[term]):
                        break                                # 숫자 밖의 낱말이 없는 곳은 후보가 아님
                    m = loose[term].search(text)
                    if not m:
                        break
                    # 실제 글: 앞의 부호(±, +/-, 빼기)와 뒤의 단위 글자까지 ('±0.5 mm', '-20℃')
                    a = m.start()
                    while a > 0 and m.start() - a < 6 and text[a - 1] in sign_chars:
                        a -= 1
                    tail = re.match(r"\s?(?:[%\u00b0\u2103]|[A-Za-z\u03a9\u2126\u03bc\u3380-\u33df]{1,6})?", text[m.end():])
                    nw[term] = text[a:m.end() + len(tail.group(0))].strip()
                    if words[term]:
                        extra.append(words[term])            # 숫자 밖의 낱말도 강조 (점수)
                elif pats[term].search(text):
                    exact.append(term)
                else:
                    break
            else:
                rows.append({"doc": doc, "unit": u, "matched_terms": exact, "missing_terms": [], "number_words": nw,
                             "terms": exact + list(nw.values()) + extra})
                break
    return rows


def widen_search(units: List[Tuple[Dict[str, Any], Dict[str, Any]]], groups: List[List[str]], flags: int,
                 status: Optional[Dict[str, Any]] = None) -> Tuple[str, List[Dict[str, Any]]]:
    """모든 검색어가 든 단위가 없을 때 넓혀 찾기. 순서: number(숫자가 든 검색어면 숫자 값만 같은 곳, number_search) ->
    partial(검색어 절반 이상, 가장 많이 든 단위) -> fuzzy(한 글자 다른 낱말 허용, 모두 또는 절반 이상).
    반환 (match_mode, [{doc, unit, terms(강조할 글), matched_terms, missing_terms, fuzzy_words 또는 number_words}]).
    status: number_search의 시간 상한 표시 (status['cut'])."""
    num_rows = number_search(units, groups, flags, status)
    if num_rows:
        return "number", num_rows
    pats = {t: _term_pattern(t, flags) for g in groups for t in g}

    def pick(rows: List[Tuple[int, Dict[str, Any]]], need_all: bool) -> List[Dict[str, Any]]:
        if need_all:
            return [r for n, r in rows if n == 0]
        best = max((r["_found"] for _, r in rows), default=0)
        return [r for _, r in rows if r["_found"] == best]

    def scan(fuzzy: bool) -> List[Tuple[int, Dict[str, Any]]]:
        rows = []
        for doc, u in units:
            text = unicodedata.normalize("NFC", u["text"])
            best_row: Optional[Tuple[int, Dict[str, Any]]] = None
            for g in groups:
                found, fw = [], {}
                for term in g:
                    if pats[term].search(text):
                        found.append(term)
                    elif fuzzy and _fuzzy_core(term):
                        w = _fuzzy_word(text, _fuzzy_core(term))
                        if w:
                            found.append(term)
                            fw[term] = w
                missing = [x for x in g if x not in found]
                if not found or (missing and (len(g) < 2 or len(found) * 2 < len(g))):
                    continue
                # matched_terms는 글자 그대로 맞은 검색어만 (비슷한 낱말로 찾은 검색어는 fuzzy_words에만)
                row = {"doc": doc, "unit": u, "matched_terms": [x for x in found if x not in fw], "missing_terms": missing,
                       "fuzzy_words": fw, "terms": [fw.get(x, x) for x in found], "_found": len(found)}
                if best_row is None or (len(missing), -len(found)) < (best_row[0], -best_row[1]["_found"]):
                    best_row = (len(missing), row)
            if best_row:
                rows.append(best_row)
        return rows

    partial = [r for r in scan(False)]
    if partial:
        return "partial", pick(partial, False)
    fuzzy = [r for r in scan(True) if r[1]["fuzzy_words"]]
    if fuzzy:
        whole = pick(fuzzy, True)
        return "fuzzy", whole or pick(fuzzy, False)
    return "exact", []


def hit_score(unit: Dict[str, Any], terms: List[str], flags: int) -> int:
    """순위 점수: 맞은 검색어 수(10점씩) + 나온 횟수(검색어마다 5번까지) + 검색어끼리 가까움(3점) + 장절 제목에 있음(3점)."""
    text = unit["text"]
    score = 10 * len(terms)
    firsts = []
    for t in terms:
        p = _term_pattern(t, flags)
        ms = []
        for m in p.finditer(text):
            ms.append(m.start())
            if len(ms) >= 5:
                break
        score += len(ms)
        if ms:
            firsts.append(ms[0])
    if len(firsts) >= 2:
        span = max(firsts) - min(firsts)
        score += 3 if span <= 50 else 1 if span <= 200 else 0
    sec = section_text(unit)
    if sec and any(_term_pattern(t, flags).search(sec) for t in terms):
        score += 3
    return score


def _not_indexed_error() -> rt.InputError:
    return rt.InputError(
        "캐시에 등록된 문서가 없어 검색할 대상이 없습니다. 문서를 먼저 등록하세요: "
        "reader.py index --folder \"<폴더>\" 또는 reader.py read --file \"<파일>\".",
        directive=("검색할 등록 문서가 없습니다. 다시 실행하지 말고 사용자에게 검색할 폴더나 파일을 물어본 뒤 "
                   "search에 --file \"<폴더 또는 파일>\"을 붙여 실행하세요 (처음 한 번 자동 등록)."),
        error_type="not_indexed",
        fields={"count": 0})


def _search_targets(args) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    """검색할 문서를 정합니다. --file이 있으면 그 파일/폴더의 미등록/변경 문서를 먼저 등록, 없으면 캐시 전체만.
    반환: (문서 목록, 정보: indexed_now, from_cache, skipped, remaining, stale, missing, skipped_types)"""
    info: Dict[str, Any] = {"indexed_now": [], "from_cache": 0, "skipped": [], "remaining": 0, "stale": [],
                            "missing": 0, "skipped_types": {}, "folder": None, "excluded_dirs": [], "refreshed": 0,
                            "reindex": 0, "orphan": 0, "refresh_failed": []}
    docs: List[Dict[str, Any]] = []
    if args.file:
        raw = os.path.expandvars(os.path.expanduser(str(args.file).strip().strip('"')))
        if os.path.isdir(raw):
            folder = os.path.abspath(raw)
            info["folder"] = folder
            files, info["skipped_types"] = _collect_folder(folder, excluded=info["excluded_dirs"])
        else:
            path = rt.resolve_input_path(raw, "검색할 파일")
            file_type(path)
            files = [path]
        res = _index_files(files, max(1, args.max_files), retry_failed=len(files) == 1 and not info["folder"])
        if len(files) == 1 and not info["folder"]:
            _raise_single_failure(res)
        docs = res["docs"]
        info["indexed_now"] = res["new"] + res["updated"]
        info["from_cache"] = len(res["cached"])
        info["remaining"] = res["remaining"]
        info["skipped"] = [{"file": f["file"], "reason": f["reason"]} for f in res["failures"]]
        return docs, info

    index = load_index()
    if not index:
        raise _not_indexed_error()
    tried = 0
    for path in sorted(index, key=lambda p: index[p].get("parsed_at", ""), reverse=True):
        doc = load_cached_only(path)
        if doc is None:
            # 검색하지 않은 문서는 skipped에만 (stale에 넣으면 '옛 내용으로 검색'으로 두 번 셈)
            if not os.path.exists(path):
                # 원본도 없으면 index --all로도 다시 읽을 수 없음 (같은 안내 반복을 막음)
                info["skipped"].append({"file": path, "reason": "캐시 파일이 없거나 손상됐고 원본도 없어 다시 읽을 수 없음 "
                                                                "-> clear --missing으로 정리"})
                info["orphan"] += 1
                info["missing"] += 1
            else:
                info["skipped"].append({"file": path, "reason": "캐시 파일이 없거나 손상됨 -> index --all로 다시 등록 필요"})
                info["reindex"] += 1
            continue
        if not os.path.exists(path):
            doc["file_missing"] = True
            info["missing"] += 1
            info["missing_searched"] = info.get("missing_searched", 0) + 1
        elif not is_stale_doc(doc):
            doc = recheck_uncomputed(doc)
        attempted = False
        known = known_failure(path) if is_stale_doc(doc) and not doc.get("file_missing") else None
        if known is not None:
            # 전에 다시 읽지 못한 원본(그대로임)은 매 실행 다시 읽지 않음 (실패하는 느린 파일 때문에 CONTINUE가 끝나지 않던 문제)
            attempted = True
            info["refresh_failed"].append({"file": path, "reason": (KNOWN_FAILURE_NOTE + known)[:300]})
        elif is_stale_doc(doc) and not doc.get("file_missing") and tried < STALE_REFRESH_MAX:
            if tried:
                # R10: 다시 읽은 캐시는 저장되고, 실패한 원본은 실패 기록에 남아 같은 명령을 다시 실행하면 건너뜀
                rt.check_time(0.0, f"원본이 바뀐 캐시 {tried}개 중 {info['refreshed']}개를 다시 읽음 (남은 캐시는 같은 명령으로 이어서)")
            tried += 1
            attempted = True
            try:
                doc, _ = load_document_status(path)     # 원본이 바뀐 캐시는 그 자리에서 다시 읽음 (JIT 복구)
                info["refreshed"] += 1
                clear_failure(path)
            except Exception as e:
                reason = rt.error_payload(e)["message"][:300]
                rt.debug(f"다시 읽기 실패, 이전 캐시로 검색: {reason}")
                record_failure(path, reason)
                info["refresh_failed"].append({"file": path, "reason": reason})
        if doc.get("parser_version") != PARSER_VERSION:
            # 이전 파서 형식의 캐시는 단위 키가 달라 검색하지 않음 (skipped에만, stale은 '옛 내용으로 검색한 문서'만)
            if doc.get("file_missing"):
                reason = "이전 형식의 캐시인데 원본이 없어 다시 읽을 수 없음 -> clear --missing으로 정리"
                info["orphan"] += 1
            elif not attempted:
                reason = (f"이전 형식의 캐시인데 검색 때 다시 읽는 상한 {STALE_REFRESH_MAX}개를 넘어 다시 읽지 않음 "
                          "-> index --all로 다시 등록 필요")
                info["reindex"] += 1
            else:
                reason = "이전 형식의 캐시인데 다시 읽지 못함 -> index --all로 다시 등록 필요"
                info["reindex"] += 1
            info["skipped"].append({"file": path, "reason": reason})
            continue
        if is_stale_doc(doc):
            doc["_stale"] = True
            info["stale"].append(path)
        docs.append(doc)
    info["from_cache"] = len(docs)
    if info["refreshed"]:
        rt.note_recovered(f"원본이 바뀐 캐시 {info['refreshed']}개 -> 다시 읽고 검색")
    return docs, info


def _page_seam_hits(docs: List[Dict[str, Any]], matcher: Any, pages: Optional[set], sec_filter: str) -> List[Dict[str, Any]]:
    """쪽 경계에 걸친 문장: 한 쪽의 마지막 단락이 문장 끝 없이 끝나고 다음 쪽 첫 단락으로 이어지면 둘을 붙여 한 번 더 찾음
    ('humidity shall not exceed' / '95 %'가 p.1, p.2로 나뉘어 0건이던 것, 2026-11-09 검토 A13). 따로는 맞지 않을 때만."""
    out: List[Dict[str, Any]] = []
    for doc in docs:
        body = [u for u in doc["units"] if u.get("page") and u.get("kind") == "paragraph" and not u.get("extra")]
        for a, b in zip(body, body[1:]):
            if b["page"] != a["page"] + 1 or _SENTENCE_END.search(a["text"]):
                continue
            if pages is not None and a["page"] not in pages and b["page"] not in pages:
                continue
            if sec_filter and not section_matches(a, sec_filter):
                continue
            joined = a["text"].rstrip() + " " + b["text"].lstrip()
            terms = matcher(joined)
            if terms is None or matcher(a["text"]) is not None or matcher(b["text"]) is not None:
                continue
            unit = dict(a, text=joined, loc=f"{a['loc']}~{b['loc']}")
            out.append({"doc": doc, "unit": unit, "terms": terms, "matched_terms": list(terms),
                        "literal": literal_hit(joined, terms)})
    return out


def _mail_whole_hits(docs: List[Dict[str, Any]], matcher: Any, hits: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """메일은 한 통이 한 검색 대상: 검색어가 제목·헤더·본문·첨부 이름에 나뉘어 있어도 맞음 ('견적;김철수'가 제목과 보낸 사람에
    나뉜 메일). 단위 하나로 이미 맞은 메일은 뺌."""
    done = {id(h["doc"]) for h in hits}
    out: List[Dict[str, Any]] = []
    for doc in docs:
        if doc.get("type") != "eml" or id(doc) in done or not doc["units"]:
            continue
        joined = " \n ".join(u["text"] for u in doc["units"] if not u["loc"].startswith("전달된 메일"))
        terms = matcher(joined)
        if terms is None:
            continue
        unit = dict(doc["units"][0], text=joined, loc="메일 전체")
        out.append({"doc": doc, "unit": unit, "terms": terms, "matched_terms": list(terms), "literal": literal_hit(joined, terms)})
    return out


def _group_mail_hits(hits: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """메일은 한 통에 결과 하나 (가장 좋은 곳, 맞은 곳 수는 mail_hits). 결과가 모두 메일이면 받은 날짜 최신순."""
    seen: Dict[int, Dict[str, Any]] = {}
    out: List[Dict[str, Any]] = []
    for h in hits:
        if h["doc"].get("type") != "eml":
            out.append(h)
            continue
        first = seen.get(id(h["doc"]))
        if first is None:
            h["mail_hits"] = 1
            seen[id(h["doc"])] = h
            out.append(h)
        else:
            first["mail_hits"] += 1
    if out and all(h["doc"].get("type") == "eml" for h in out):
        out.sort(key=lambda h: str(((h["doc"].get("meta") or {}).get("mail") or {}).get("date") or ""), reverse=True)
    return out


def cmd_search(args) -> Dict[str, Any]:
    matcher = build_matcher(args.query, args.regex, args)
    flags = re.IGNORECASE
    clamp_warnings: List[str] = []
    if args.top > MAX_SEARCH_TOP:
        clamp_warnings.append(f"--top {args.top}은 상한 {MAX_SEARCH_TOP}으로 줄임 (전체 결과는 --save 파일로)")
        args.top = MAX_SEARCH_TOP
    if args.context > MAX_SEARCH_CONTEXT or args.context < 0:
        clamp_warnings.append(f"--context {args.context}는 0~{MAX_SEARCH_CONTEXT}로 줄임 (단락 전체는 --full 또는 --save)")
        args.context = min(MAX_SEARCH_CONTEXT, max(0, args.context))
    for w in clamp_warnings:
        rt.log("WARN", w)
    pages = page_set(args.pages)
    sec_filter = section_key(args.section)
    docs, info = _search_targets(args)
    note_same_names([d["file"] for d in docs])
    hits: List[Dict[str, Any]] = []
    page_unknown = 0
    pool: List[Tuple[Dict[str, Any], Dict[str, Any]]] = []      # 조건(--pages, --section) 안의 단위 (넓혀 찾기용)
    by_number: List[str] = []         # 쪽 정보가 없는 문서: --pages를 항목 번호로 (read --pages와 같은 뜻, 2026-11-09 검토 A4)
    for doc in docs:
        doc_by_page = doc["type"] in PAGE_TYPES or any(u.get("page") for u in doc["units"])
        if pages is not None and not doc_by_page and doc["units"]:
            by_number.append(os.path.basename(doc["file"]))
        for n_unit, u in enumerate(doc["units"], 1):
            if sec_filter and not section_matches(u, sec_filter):
                continue
            if pages is not None and not u.get("extra"):     # 머리글·바닥글·각주는 쪽 범위와 상관없이 찾음
                if not doc_by_page:
                    if n_unit not in pages:
                        continue
                elif u.get("page") is None:
                    page_unknown += 1
                    continue
                elif u["page"] not in pages:
                    continue
            pool.append((doc, u))
            terms = matcher(u["text"])
            if terms is None:
                continue
            hits.append({"doc": doc, "unit": u, "terms": terms, "matched_terms": list(terms),
                         "literal": True if args.regex else literal_hit(u["text"], terms)})
    if not args.regex:
        hits += _page_seam_hits(docs, matcher, pages, sec_filter)
        hits += _mail_whole_hits(docs, matcher, hits)
    groups = [] if args.regex else parse_query(args.query)
    has_num = any(literal_finder(t) for g in groups for t in g)       # 숫자가 든 검색어: 결과마다 literal(층 A) 표시
    # 층 B(표기가 다른 일치)가 섞이면 exact_mixed: 결과에 함께 넣되 VERIFIED로 내지 않음
    mode = "exact_mixed" if any(not h["literal"] for h in hits) else "exact"
    net_status: Dict[str, Any] = {}
    if not hits and groups and pool:
        # 0건: 숫자가 든 검색어면 숫자 값만 같은 곳(안전망) -> 일부 검색어만 든 단위 -> 한 글자 다른 낱말로 넓혀 찾음.
        # 정확 일치는 여전히 0건이라 count 0, 넓혀 찾은 결과는 candidates와 UNVERIFIED로 알림 (근거로 쓰지 않게)
        mode, hits = widen_search(pool, groups, flags, net_status)
    for h in hits:
        h["score"] = hit_score(h["unit"], h["terms"], flags) if not args.regex else 0
        h["snippet"] = make_snippet(h["unit"]["text"], h["terms"], args.context)
    if not args.regex:
        # 글자 그대로 일치(층 A)를 먼저, 그 안에서 좋은 결과부터 (같은 점수는 문서 순서)
        hits.sort(key=lambda h: (not h.get("literal", True), -h["score"]))
    hits = _group_mail_hits(hits)
    syn_used = {t: synonyms_of(t) for g in groups for t in g if synonyms_of(t)}

    total = len(hits)
    shown = hits[: max(0, args.top)]
    exact_like = mode in ("exact", "exact_mixed")
    exact_total = total if exact_like else 0       # 넓혀 찾은 결과는 일치 건수에 넣지 않음
    literal_total = sum(1 for h in hits if h.get("literal", True)) if exact_like else 0
    result: Dict[str, Any] = {"query": args.query, "files_searched": len(docs), "total_matches": exact_total,
                              "count": exact_total, "literal_count": literal_total,
                              "matched": len(shown) if exact_like else 0, "shown": len(shown),
                              "truncated": total > len(shown), "match_mode": mode}
    if not exact_like:
        result["candidate_count"] = total
    if syn_used:
        result["synonyms_used"] = syn_used
    filters = {k: v for k, v in (("file", args.file), ("pages", args.pages), ("section", args.section)) if v}
    if filters:
        result["filters"] = filters
    if clamp_warnings:
        result["warnings"] = clamp_warnings
    if pages is not None and page_unknown:
        result["page_unknown_skipped"] = page_unknown
    if by_number:
        clamp_warnings.append(f"쪽 정보가 없는 문서 {len(by_number)}개({', '.join(by_number[:3])})는 --pages를 항목 번호로 봄 "
                              "(read --pages와 같음)")
        result["warnings"] = clamp_warnings
    if args.file:
        result["indexed_now_count"] = len(info["indexed_now"])
        result["indexed_now"] = [os.path.basename(p) for p in info["indexed_now"][:MAX_LIST_ITEMS]]
        result["from_cache"] = info["from_cache"]
        if info["remaining"]:
            result["not_indexed_remaining"] = info["remaining"]
    if info["skipped_types"]:
        result["skipped_types"] = info["skipped_types"]
    if info["skipped"]:
        result["skipped"] = [{"file": os.path.basename(s["file"]), "path": s["file"], "reason": s["reason"]}
                             for s in info["skipped"][:MAX_LIST_ITEMS]]
    if info["stale"]:
        result["stale_cache_count"] = len(info["stale"])
    if info["refresh_failed"]:
        # 원본이 바뀐 캐시를 다시 읽지 못한 파일 (실패 기록에 남겨 다음 실행에서 다시 읽지 않음)
        result["refresh_failures"] = [{"file": os.path.basename(f["file"]), "path": f["file"], "reason": f["reason"]}
                                      for f in info["refresh_failed"][:MAX_LIST_ITEMS]]
    if info["missing"]:
        result["missing_file_count"] = info["missing"]

    nexts: List[Dict[str, str]] = []
    if args.save and hits:
        stem = "search_" + re.sub(r"[^\w가-힣]+", "_", args.query)[:30].strip("_")
        ext = _save_ext(args.save)
        save_arg = None if args.save is True else args.save
        # 파일 하나를 찾았으면 그 파일 옆, 폴더·캐시 전체면 C:\\AIKIT\\data (폴더에 두면 다음 검색 때 결과 파일이 등록됨)
        out = rt.resolve_output_path(save_arg, stem, ext, near=args.file if args.file and os.path.isfile(str(args.file)) else None)
        saved = hits[:MAX_SAVE_HITS]
        cut_cells = 0
        clamped_context = 0
        header = ["파일", "쪽", "장절", "번호", "위치", "내용(문장/단락)", "문맥", "점수"]
        mixed_save = mode == "exact_mixed"
        if mixed_save:
            header.append("표기")                     # 글자 그대로(층 A) / 표기 다름(층 B, 부호·단위 확인)
        rows = []
        for h in saved:
            u = h["unit"]
            if args.full:
                context_text = u["text"]          # --full: 단위 전체 (xlsx는 셀 상한 32000자 확인이 따로 있음)
            else:
                context_text, cut = rt.clamp_text(u["text"], SAVE_CONTEXT_CHARS)
                clamped_context += 1 if cut else 0
            rows.append([os.path.basename(h["doc"]["file"]), u.get("page") if u.get("page") is not None else "",
                         section_text(u), u.get("number") or "", u["loc"],
                         match_sentence(u["text"], h["terms"]), context_text, h["score"]]
                        + (["글자 그대로" if h["literal"] else "표기 다름"] if mixed_save else []))
        if ext == ".csv":
            out = rt.write_csv_rows(out, [header] + rows)
        elif ext == ".xlsx":
            out, cut_cells = write_xlsx_table(out, header, rows, [22, 6, 45, 10, 14, 60, 80, 6] + [12] * mixed_save, "검색결과")
        else:
            lines = [f"# 검색 결과: {args.query}", ""]
            for h in saved:
                u = h["unit"]
                sec = section_text(u)
                lines += [f"## {os.path.basename(h['doc']['file'])} > {u['loc']}" + (f" ({sec})" if sec else "")
                          + (" [표기 다름: 부호·단위 확인]" if mixed_save and not h["literal"] else ""), "",
                          f"근거: {evidence_text(h['doc']['file'], u)}", "",
                          u["text"] if args.full else h["snippet"].replace("**", ""), ""]
            out = rt.write_text(out, "\n".join(lines))
        result.update(output_file=out, saved=len(saved))
        if clamped_context and ext in (".csv", ".xlsx"):
            result["context_clamped"] = clamped_context
            result["context_clamped_reason"] = (f"문맥 칸은 {SAVE_CONTEXT_CHARS}자까지 저장: {clamped_context}행 잘림 "
                                                "(단락 전체는 --full을 붙여 저장)")
        if len(hits) > len(saved):
            result.update(save_truncated=True,
                          save_truncated_reason=f"검색 결과 {len(hits)}건 중 앞 {len(saved)}건만 저장 (저장 상한 {MAX_SAVE_HITS}건)")
        save_problem = rt.check_saved_file(out, len(rows) + 1 if ext == ".csv" else None, delimiter=",")
        if not save_problem and cut_cells:
            save_problem = _xlsx_cut_text(cut_cells)
        if save_problem:
            result["save_problem"] = save_problem
        if _write_confirm_commands(out):
            result["confirm_commands"] = _write_confirm_commands(out)

    view = []
    text_cut = 0
    for h in shown:
        u, doc = h["unit"], h["doc"]
        item = {"file": evidence_name(doc["file"]), "path": doc["file"], "page": u.get("page"),
                "page_source": u.get("page_source"), "section": u.get("section") or [],
                "section_text": section_text(u), "number": u.get("number"), "loc": u["loc"],
                "heading": u.get("heading", ""), "evidence": evidence_text(doc["file"], u), "snippet": h["snippet"],
                "score": h["score"], "matched_terms": h["matched_terms"]}
        if h.get("missing_terms"):
            item["missing_terms"] = h["missing_terms"]
        if h.get("fuzzy_words"):
            item["fuzzy_words"] = h["fuzzy_words"]
        if h.get("number_words"):
            item["number_words"] = h["number_words"]
        if has_num and exact_like:
            item["literal"] = h["literal"]        # False: 표기가 다른 곳(띄어쓰기·단위 별칭·부호 판정으로 찾음)
        if args.full:
            item["text"], cut = rt.clamp_text(u["text"], SEARCH_JSON_FULL_CHARS)
            if cut:
                item["text_truncated"] = True     # 전체는 --save 파일 (--full이면 단위 전체 저장)
                item["text_chars"] = len(u["text"])
                text_cut += 1
        if doc.get("type") == "eml":            # 메일: 원본 eml 경로(path)로 Outlook에서 열어 봄
            item["mail"] = (doc.get("meta") or {}).get("mail", {})
            if h.get("mail_hits", 1) > 1:
                item["mail_hits"] = h["mail_hits"]
        if doc.get("file_missing"):
            item["file_missing"] = True
        if doc.get("_stale"):
            item["stale"] = True             # 원본이 바뀌기 전 내용 (근거로 쓰기 전에 다시 등록 필요)
        view.append(item)
    result["matches" if exact_like else "candidates"] = view
    if text_cut:
        result["truncated"] = True
        result["truncated_reason"] = (f"--full 항목 {text_cut}개의 text를 {SEARCH_JSON_FULL_CHARS}자에서 자름 "
                                      "(text_truncated 표시, 전체는 --save 파일)")

    if (info["stale"] or info["reindex"]) and not args.file:
        nexts.append(rt.next_command("reader.py", ["index", "--all"],
                                     f"원본이 바뀌었거나 다시 읽지 못한 캐시 {len(info['stale']) + info['reindex']}개 다시 등록 (읽기 전용)"))
    if info["orphan"]:
        # 원본이 없는 이전 형식 캐시는 index --all로도 다시 읽을 수 없음 (반복 안내를 막음). 원본 문서는 지우지 않음
        nexts.append(rt.next_command("reader.py", ["clear", "--missing", "--dry-run"],
                                     f"원본이 없는 캐시 정리 미리보기 (다시 읽을 수 없는 캐시 {info['orphan']}개 포함, 원본 문서는 그대로)"))
    if info["remaining"] and info["folder"]:
        nexts.append(rt.next_command("reader.py", ["index", "--folder", info["folder"], "--max-files", str(args.max_files)],
                                     f"아직 등록하지 못한 {info['remaining']}개 파일 등록 (읽기 전용)"))
    if result["truncated"] and not args.save:
        cmd = ["search", "--query", args.query]
        for flag, val in (("--file", args.file), ("--pages", args.pages), ("--section", args.section)):
            if val:
                cmd += [flag, val]
        if args.regex:
            cmd.append("--regex")
        if args.full:
            cmd.append("--full")
        # --full에서 엑셀 셀 상한(XLSX_CELL_MAX자)보다 긴 단락이 있으면 .xlsx는 반드시 잘림(UNVERIFIED): .md로
        long_unit = args.full and any(len(h["unit"]["text"]) > XLSX_CELL_MAX for h in hits[:MAX_SAVE_HITS])
        if long_unit:
            nexts.append(rt.next_command("reader.py", cmd + ["--save", "근거.md"],
                                         f"전체 {total}건을 .md 파일로 저장 (단락 전체, 엑셀 칸 상한 {XLSX_CELL_MAX}자보다 긴 단락이 있어 .md)"))
        else:
            nexts.append(rt.next_command("reader.py", cmd + ["--save", "근거.xlsx"],
                                         f"전체 {total}건을 엑셀 파일로 저장" + (" (단락 전체)" if args.full else "")))
    if nexts:
        result["next_commands"] = nexts

    scope = f"문서 {len(docs)}개"
    if args.file and info["indexed_now"]:
        scope += f"(새로 등록 {len(info['indexed_now'])}개)"
    notes = ""
    if info["stale"] and not args.file:
        notes += (f" 캐시 {len(info['stale'])}개는 원본이 바뀌기 전 내용입니다(결과의 stale 표시). 최신 내용이 필요하면 "
                  "next_commands의 index --all(읽기 전용)을 실행한 뒤 같은 검색을 1번 다시 하세요.")
    if info["indexed_now"]:
        notes += " 새로 등록한 파일(indexed_now)을 사용자에게 알리세요."
    if clamp_warnings:
        notes += " warnings(상한으로 줄인 옵션, --pages를 항목 번호로 본 문서)를 사용자에게 알리세요."
    if text_cut:
        notes += (f" 일부 결과의 text가 {SEARCH_JSON_FULL_CHARS}자에서 잘렸습니다(text_truncated). 전체는 "
                  + ("저장 파일(output_file)에 있습니다." if "output_file" in result else "next_commands의 --save로 저장하세요."))
    if result.get("context_clamped"):
        notes += f" 저장 파일의 {result['context_clamped_reason']}."
    gaps = _search_gaps(docs, info, pages, sec_filter)
    if pages is not None and page_unknown:
        gaps.append(f"쪽 번호를 알 수 없는 단위 {page_unknown}개는 --pages 안인지 몰라 찾지 않음 (--pages 없이 다시 검색)")
    if net_status.get("cut"):
        # 숫자 안전망이 시간 상한(NUMBER_SEARCH_SECONDS)으로 일부 단위만 봄: 0건이어도 '규정 없음'으로 단정하지 않음
        gaps.append(f"숫자 안전망 검색이 {NUMBER_SEARCH_SECONDS:.0f}초 상한으로 단위 {len(pool)}개 중 "
                    f"{net_status.get('checked', 0)}개만 확인함 (--file이나 --section으로 범위를 좁혀 다시 검색)")
    if gaps:
        result["not_searched"] = gaps
    if total == 0:
        preview = f"'{args.query}': {scope}에서 일치 항목 없음 (0건)"
        if gaps:
            # 일부를 검색하지 못했으면 "근거 없음"으로 단정하지 않음 (규격 대조에서 잘못된 판정이 됨)
            return rt.ok_unverified(TOOL, "search", preview + ", 검색 범위 불완전",
                                    ["0건이지만 검색하지 못한 부분이 있음: " + "; ".join(gaps)], **result)
        return rt.ok(TOOL, "search", preview,
                     "VERIFIED: 0건. 조건에 맞는 항목이 없습니다. 사용자에게 그대로 알리고 조건을 바꿔 다시 실행하지 마세요." + notes,
                     **result)
    if not exact_like:
        # 정확히 일치한 곳은 0건: 넓혀 찾은 결과는 후보일 뿐이라 VERIFIED로 내지 않음 ("대기 전류"가 "대기 전력"에 맞던 문제)
        what = {"partial": "일부 검색어만 든 곳, 빠진 검색어는 missing_terms",
                "number": "숫자 값만 같은 곳(부호·단위 표기는 다를 수 있음), 실제 글은 number_words"}.get(
                    mode, "한 글자 다른 낱말이 든 곳, 실제 낱말은 fuzzy_words")
        kind = "숫자가 같은 곳" if mode == "number" else "비슷한 곳"
        preview = (f"'{args.query}': {scope}에서 정확히 일치한 곳 0건, {kind} {total}건(candidates, match_mode {mode})"
                   + (f", 저장: {result['output_file']}" if "output_file" in result else ""))
        payload = rt.ok_unverified(TOOL, "search", preview, [f"정확히 일치한 곳은 0건, {kind} {total}건({what})"]
                                   + (["검색하지 못한 부분이 있음: " + "; ".join(gaps)] if gaps else []), **result)
        head = (f"UNVERIFIED: 정확히 일치한 곳은 0건, 숫자가 같은 곳 {total}건 - 표기를 확인하세요 (match_mode number: {what}). "
                if mode == "number" else
                f"UNVERIFIED: 정확히 일치한 곳은 0건, 비슷한 곳 {total}건(match_mode {mode}: {what}) - ")
        payload["agent_directive"] = (head + "candidates를 근거로 쓰지 말고, 사용자에게 "
                                      f"'정확히 일치한 곳은 없고 {kind}이 있다'고 알린 뒤 맞는지 확인을 받으세요."
                                      # stale 안내(index --all 뒤 1번 다시 검색)와 부딪히지 않게 그때는 '다시 실행 금지'를 빼고 안내를 따름
                                      + ("" if info["stale"] and not args.file else " 같은 검색을 다시 실행하지 마세요.")
                                      + notes)                # stale·indexed_now·clamp·text_cut 안내를 지우지 않음
        return payload
    preview = f"'{args.query}': {scope}에서 {total}건 일치" + (f" (상위 {len(shown)}건 표시)" if total > len(shown) else "")
    # 층 B가 섞임: 글자 그대로가 아닌 일치는 부호·단위 표기가 다를 수 있어 VERIFIED로 내지 않음 (1.4.70)
    mixed = ([f"글자 그대로 일치 {literal_total}건, 표기가 다른 곳 {total - literal_total}건(literal: false, 부호·단위를 확인하세요)"]
             if mode == "exact_mixed" else [])
    if mixed:
        preview += f" (글자 그대로 {literal_total}건, 표기가 다른 곳 {total - literal_total}건)"
    evid = " 각 결과의 근거는 evidence(파일 p.N, 장절 (번호))를 그대로 알리세요. page가 null이면 장절과 번호로 알립니다."
    if has_num:
        evid += " " + NUM_CONTEXT_NOTE
    if "output_file" in result:
        preview += f", 저장: {result['output_file']}"
        directive = "VERIFIED: 검색과 저장 완료(저장 파일을 다시 열어 확인)." + evid
        save_issues = [x for x in (result.get("save_truncated_reason"), result.get("save_problem")) if x]
        if save_issues:
            payload = rt.ok_unverified(TOOL, "search", preview, mixed + save_issues + (
                ["검색하지 못한 부분이 있음: " + "; ".join(gaps)] if gaps else []), **result)
            payload["agent_directive"] += evid
            return payload
        if result.get("confirm_commands"):
            directive += " " + rt.INTO_OTHER_DOC
    else:
        directive = "VERIFIED: 검색 완료." + evid
        if result["truncated"]:
            # 저장 이름은 next_commands에만 (긴 단락이 있으면 .md, 아니면 .xlsx: directive에 이름을 박으면 서로 달랐음)
            directive += " 일부만 표시했습니다. 전체 결과가 필요하면 next_commands의 저장 명령을 그대로 실행하세요."
    if gaps:
        payload = rt.ok_unverified(TOOL, "search", preview + ", 검색 범위 불완전",
                                   mixed + ["검색하지 못한 부분이 있어 결과가 전부가 아닐 수 있음: " + "; ".join(gaps)], **result)
        payload["agent_directive"] += evid
        return payload
    if mixed:
        payload = rt.ok_unverified(TOOL, "search", preview, mixed, **result)
        payload["agent_directive"] = (
            "UNVERIFIED: " + mixed[0] + ". literal: true인 결과만 근거로 쓰고, literal: false인 결과는 사용자에게 실제 글(snippet)을 "
            "보여 주고 맞는지 확인을 받으세요."
            + ("" if info["stale"] and not args.file else " 같은 검색을 다시 실행하지 마세요.")
            + " " + directive.split(". ", 1)[1] + notes)        # 근거(evidence)·저장·일부 표시 안내는 그대로
        return payload
    return rt.ok(TOOL, "search", preview, directive + notes, **result)


def _pages_in_filter(doc: Dict[str, Any], flagged: List[int], pages: Optional[set], sec_filter: str) -> List[int]:
    """글자 없는 쪽 중 검색 조건(--pages, --section) 안에 드는 쪽. 장절은 그 쪽 바로 앞 글의 장절로 봅니다
    (글이 없는 쪽에는 장절 정보가 없음)."""
    out = [p for p in flagged if pages is None or p in pages]
    if not sec_filter or not out:
        return out
    units = [u for u in doc.get("units") or [] if u.get("page") is not None]
    keep = []
    for p in out:
        before = [u for u in units if u["page"] < p]
        if before and section_matches(before[-1], sec_filter):
            keep.append(p)
    return keep


def _search_gaps(docs: List[Dict[str, Any]], info: Dict[str, Any], pages: Optional[set] = None,
                 sec_filter: str = "") -> List[str]:
    """검색 범위에서 빠진 것: 결과가 없거나 적어도 '없다'고 단정하면 안 되는 이유들.
    다시 할 명령(index --all, clear --missing)을 문구에 넣어 UNVERIFIED directive만 보고도 알 수 있게 합니다."""
    gaps: List[str] = []
    if info["remaining"]:
        gaps.append(f"아직 등록하지 못한 파일 {info['remaining']}개 (next_commands의 index로 등록)")
    if info["skipped"]:
        how = []
        if info.get("reindex"):
            how.append(f"{info['reindex']}개는 reader.py index --all로 다시 등록")
        if info.get("orphan"):
            how.append(f"{info['orphan']}개는 원본이 없어 다시 읽을 수 없음, reader.py clear --missing으로 정리")
        gaps.append(f"읽지 못한 파일 {len(info['skipped'])}개 (skipped" + (": " + ", ".join(how) if how else "") + ")")
    if info["stale"]:
        gaps.append(f"원본이 바뀐 뒤 다시 읽지 못한 캐시 {len(info['stale'])}개 (옛 내용으로 검색, 결과의 stale 표시, "
                    "reader.py index --all로 다시 등록)")
    ocr = [os.path.basename(d["file"]) for d in docs if (d.get("meta") or {}).get("ocr_needed")]
    if ocr:
        gaps.append(f"본문 글자가 없는 PDF {len(ocr)}개 (스캔 이미지, OCR 필요): {', '.join(ocr[:5])}")
    blank = []
    for d in docs:
        meta = d.get("meta") or {}
        if meta.get("ocr_needed"):
            continue
        flagged = _pages_in_filter(d, meta.get("pages_without_text") or [], pages, sec_filter)
        if flagged:
            blank.append((os.path.basename(d["file"]), flagged))
    if blank:
        gaps.append(f"글이 거의 없고 그림만 있는 쪽이 있는 PDF {len(blank)}개 (스캔 쪽일 수 있음, 그 쪽은 검색 안 됨, OCR 필요): "
                    + ", ".join(f"{n} p.{_page_list(p, 5)}" for n, p in blank[:5]))
    cut = [os.path.basename(d["file"]) for d in docs if (d.get("meta") or {}).get("rows_truncated")]
    if cut:
        gaps.append(f"행이 많아 {MAX_SHEET_ROWS}행까지만 읽은 엑셀 {len(cut)}개: {', '.join(cut[:5])}")
    uncomputed = [(os.path.basename(d["file"]), (d.get("meta") or {})["uncomputed_formulas"]) for d in docs
                  if (d.get("meta") or {}).get("uncomputed_formulas")]
    if uncomputed:
        gaps.append(f"계산된 값이 저장되지 않은 수식 칸이 있거나 확인하지 못한 엑셀 {len(uncomputed)}개 (그 칸은 빈 칸이나 계산되지 않은 값으로 "
                    "읽혀 검색 결과를 믿을 수 없음, "
                    "엑셀에서 연 적 없이 프로그램이 만든 파일. 엑셀에서 열어 저장한 뒤 reader.py index --all로 다시 등록): "
                    + ", ".join(f"{n} 확인 못함" if u.get("unknown") else f"{n} {u['count']}칸({', '.join(u['cells'][:3])})"
                                for n, u in uncomputed[:5]))
    if info.get("missing_searched"):
        gaps.append(f"원본이 지워졌거나 옮겨진 문서 {info['missing_searched']}개는 마지막으로 읽은 캐시로 찾음 (결과의 file_missing, "
                    "원본 위치를 확인하고 reader.py clear --missing으로 정리)")
    legacy = {k: v for k, v in (info.get("skipped_types") or {}).items() if k in LEGACY_TYPES}
    if legacy:
        gaps.append(f"읽지 않는 옛 형식 문서 {sum(legacy.values())}개 ("
                    + ", ".join(f"{k} {v}개" for k, v in sorted(legacy.items()))
                    + ", Office에서 열어 .docx·.xlsx·.pptx로 저장하면 검색됨)")
    if info["excluded_dirs"]:
        gaps.append(f"들어가지 않은 하위 폴더 {len(info['excluded_dirs'])}개 (깊이 제한 또는 '_'로 시작): "
                    + ", ".join(os.path.basename(d) for d in info["excluded_dirs"][:5]))
    return gaps


# ==============================================================================
# 8. 명령: list (캐시된 문서 조회)
# ==============================================================================

def cmd_list(args) -> Dict[str, Any]:
    index = load_index()
    name_filter = str(args.name or "").strip().lower()
    folder = _folder_arg(args.folder) if args.folder else None
    items = []
    for path, info in sorted(index.items(), key=lambda kv: kv[1].get("parsed_at", ""), reverse=True):
        if name_filter and name_filter not in os.path.basename(path).lower() and name_filter not in path.lower():
            continue
        if folder and not _under(path, folder):
            continue
        exists = os.path.exists(path)
        stale = info.get("parser_version") != PARSER_VERSION
        if exists and not stale:
            try:
                st = os.stat(path)
                stale = info.get("mtime") != st.st_mtime or info.get("size") != st.st_size
            except OSError:
                pass
        # 같은 뜻 키(read_at = cached_at, cache = stale, exists = not missing)와 문자 수는 넣지 않음 (2026-11-09 검토 5장 4번:
        # 40개에 14.6KB). parser_version은 이전 형식일 때만
        item = {"file": os.path.basename(path), "path": path, "type": info.get("type"), "units": info.get("units"),
                "pages": info.get("pages"), "sections": info.get("sections"), "cached_at": info.get("parsed_at"),
                "stale": stale, "missing": not exists}
        if info.get("parser_version", 1) != PARSER_VERSION:
            item["parser_version"] = info.get("parser_version", 1)
        items.append(item)
    limit = LIST_SHOW
    shown = items[:limit]
    stale_n = sum(1 for i in items if i["stale"])
    missing_n = sum(1 for i in items if i["missing"])
    result: Dict[str, Any] = {"documents": shown, "total": len(items), "count": len(items), "shown": len(shown),
                              "truncated": len(items) > limit, "stale_count": stale_n, "missing_count": missing_n}
    if result["truncated"]:
        out = os.path.join(rt.get_data_dir(), "results", f"cached_documents_{rt.timestamp()}.csv")
        os.makedirs(os.path.dirname(out), exist_ok=True)
        cols = ["file", "path", "type", "units", "pages", "sections", "cached_at", "stale", "missing"]
        result["full_result_file"] = rt.write_csv_rows(out, [cols] + [[i[c] for c in cols] for i in items])
    if stale_n:
        result["next_commands"] = [rt.next_command("reader.py", ["index", "--all"],
                                                   f"바뀌었거나 이전 형식인 캐시 {stale_n}개 다시 등록 (읽기 전용)")]
    cond = " ".join(x for x in (f"이름 '{args.name}'" if args.name else "", f"폴더 {folder}" if folder else "") if x)
    if not items:
        what = f"{cond}에 맞는 캐시 문서가 없습니다" if cond else "캐시에 등록된 문서가 없습니다"
        return rt.ok(TOOL, "list", f"0건: {what}",
                     f"VERIFIED: 0건. {what}. 사용자에게 그대로 알리고 조건을 바꿔 다시 실행하지 마세요.", **result)
    preview = (f"캐시 문서 {len(items)}개" + (f" ({cond})" if cond else "") + (f", 최근 {limit}개 표시" if result["truncated"] else "")
               + (f", 원본 변경/이전 형식 {stale_n}개" if stale_n else "") + (f", 원본 없음 {missing_n}개" if missing_n else ""))
    directive = "VERIFIED: 캐시 목록 조회 완료. documents의 file, pages, cached_at, stale, missing을 사용자에게 알리세요."
    if result["truncated"]:
        directive += " 전체 목록은 full_result_file에 있습니다."
    return rt.ok(TOOL, "list", preview, directive, **result)


# ==============================================================================
# 8-1. 명령: clear (캐시 지우기)
# ==============================================================================

def cmd_clear(args) -> Dict[str, Any]:
    """캐시에서 문서를 지웁니다 (색인 항목과 캐시 파일). 원본 문서는 건드리지 않습니다."""
    if not any([args.file, args.name, args.folder, args.missing, args.all]):
        raise rt.InputError("지울 대상을 정하세요: --file 파일, --name 이름일부, --folder 폴더, --missing(원본 없는 문서), --all(전체). "
                            "먼저 --dry-run으로 지울 목록을 확인할 수 있습니다.")
    if args.all and any([args.file, args.name, args.folder, args.missing]):
        raise rt.InputError("--all은 다른 조건과 함께 쓸 수 없습니다.")
    index = load_index()
    file_key = (os.path.normcase(os.path.abspath(os.path.expandvars(os.path.expanduser(str(args.file).strip().strip('"')))))
                if args.file else None)
    folder = os.path.abspath(os.path.expandvars(os.path.expanduser(str(args.folder).strip().strip('"')))) if args.folder else None
    name = str(args.name or "").strip().lower()
    targets = []
    for path in index:
        if args.all:
            targets.append(path)
            continue
        if file_key and os.path.normcase(os.path.abspath(path)) != file_key:
            continue
        if name and name not in os.path.basename(path).lower() and name not in path.lower():
            continue
        if folder and not _under(path, folder):
            continue
        if args.missing and os.path.exists(path):
            continue
        targets.append(path)
    cache_dir = os.path.join(rt.get_cache_dir(), "docs")
    orphans: List[str] = []
    if args.all and os.path.isdir(cache_dir):              # 색인에 없는 캐시 파일(중단 등으로 남은 것)도 정리
        known = {os.path.basename(_doc_cache_path(p)) for p in index}
        orphans = [os.path.join(cache_dir, f) for f in os.listdir(cache_dir) if f.endswith(".json") and f not in known]
    items = [{"file": os.path.basename(p), "path": p, "missing": not os.path.exists(p)} for p in targets]
    cond = ", ".join(x for x in (f"파일 {args.file}" if args.file else "", f"이름 '{args.name}'" if args.name else "",
                                 f"폴더 {folder}" if folder else "", "원본 없는 문서" if args.missing else "",
                                 "전체" if args.all else "") if x)
    result: Dict[str, Any] = {"count": len(items), "documents": items[:MAX_LIST_ITEMS], "shown": min(len(items), MAX_LIST_ITEMS),
                              "truncated": len(items) > MAX_LIST_ITEMS, "condition": cond}
    if orphans:
        result["orphan_cache_files"] = len(orphans)

    def write_full_list(kind: str, removed_set: Optional[set] = None) -> None:
        """R2: 잘린 목록은 전체를 파일로. 미리보기(plan)는 지울 대상, 실제 지우기(done)는 지운 뒤 removed 열(지웠는지)과 함께."""
        out = os.path.join(rt.get_data_dir(), "results", f"clear_{kind}_{rt.timestamp()}.csv")
        os.makedirs(os.path.dirname(out), exist_ok=True)
        head = ["file", "path", "missing"] + (["removed"] if removed_set is not None else [])
        rows = [[i["file"], i["path"], i["missing"]] + ([i["path"] in removed_set] if removed_set is not None else [])
                for i in items]
        result["full_result_file"] = rt.write_csv_rows(out, [head] + rows)

    if result["truncated"] and args.dry_run:
        write_full_list("plan")
    if not items and not orphans:
        return rt.ok(TOOL, "clear", f"0건: {cond}에 맞는 캐시 문서가 없습니다",
                     f"VERIFIED: 0건. {cond}에 맞는 캐시 문서가 없어 지운 것이 없습니다. 사용자에게 그대로 알리세요.", **result)
    if args.dry_run:
        cmd = ["clear"] + [x for flag, val in (("--file", args.file), ("--name", args.name), ("--folder", args.folder))
                           if val for x in (flag, val)] + (["--missing"] if args.missing else []) + (["--all"] if args.all else [])
        result.update(dry_run=True, planned=[f"캐시 삭제: {i['file']}" for i in items[:MAX_LIST_ITEMS]],
                      confirm_commands=[rt.next_command("reader.py", cmd, "미리 본 캐시를 실제로 지우기 (사용자 확인 뒤에만)")])
        return rt.ok(TOOL, "clear", f"미리보기: 캐시 문서 {len(items)}개를 지울 예정 ({cond})",
                     "VERIFIED: 미리보기만 했습니다(변경 없음). planned를 사용자에게 보여주고 사용자가 '예'라고 한 뒤에만 confirm_commands[0]을 실행하세요."
                     + (f" planned는 {len(items)}개 중 {MAX_LIST_ITEMS}개만 담았고 전체는 full_result_file에 있습니다." if result["truncated"] else ""),
                     **result)
    failed: List[str] = []
    removed: List[str] = []
    for p in targets:
        cf = _doc_cache_path(p)
        try:
            if os.path.exists(cf):
                rt.retry_fs(lambda: os.remove(cf))
        except OSError as e:
            failed.append(f"{os.path.basename(p)}: {e}")
            continue
        removed.append(p)
    for f in orphans:
        try:
            rt.retry_fs(lambda: os.remove(f))
        except OSError as e:
            failed.append(f"{os.path.basename(f)}: {e}")
    if args.all and os.path.exists(_failed_path()):       # 읽지 못한 파일 기록도 비움 (다음 등록 때 모두 다시 시도)
        try:
            rt.retry_fs(lambda: os.remove(_failed_path()))
        except OSError as e:
            failed.append(f"{FAILED_FILE}: {e}")
    # 잠근 상태에서 색인을 다시 읽어 지운 항목만 뺌 (그 사이 다른 실행이 넣은 항목을 지우지 않게)
    _index_update(lambda idx: [idx.pop(p, None) for p in removed])
    # 다시 읽어 확인: 색인에서 빠졌고 캐시 파일도 없는지
    left = load_index()
    still = [p for p in targets if p in left or os.path.exists(_doc_cache_path(p))]
    done = set(removed) - set(still)                     # 지웠고 다시 읽어도 없는 것만 (applied·removed 열)
    if result["truncated"]:
        write_full_list("done", done)
    result.update(applied=[f"캐시 삭제: {os.path.basename(p)}" for p in targets if p in done][:MAX_LIST_ITEMS],
                  remaining_in_cache=len(left))
    preview = f"캐시 문서 {len(targets) - len(still)}개를 지웠습니다 ({cond}). 원본 문서는 그대로입니다. 남은 캐시 문서 {len(left)}개"
    if still or failed:
        result["failed"] = failed[:MAX_LIST_ITEMS]
        return rt.ok_unverified(TOOL, "clear", preview, [f"지우지 못한 캐시 {len(still) or len(failed)}개"] + failed[:3], **result)
    return rt.ok(TOOL, "clear", preview,
                 "VERIFIED: 캐시를 지우고 다시 확인했습니다. 원본 문서는 그대로이며, 다시 검색하면 그때 새로 읽습니다.",
                 verified=True, **result)


# ==============================================================================
# 8-1. 규격 대조 도우미 (excel.py check가 씀): 항목 찾기, 규격 글의 수치 조건, 값 판정
# ==============================================================================

# 단위: 쓴 글자 -> (기준 단위, 배율). NFKC 뒤의 글자로 적음 (℃ -> °C, Ω는 그대로, 마이크로 기호(U+00B5) -> μ)
_UNIT_TABLE: Dict[str, Tuple[str, float]] = {}
for _base, _units in {
    "V": {"μV": 1e-6, "uV": 1e-6, "mV": 1e-3, "V": 1, "kV": 1e3, "MV": 1e6},
    "A": {"μA": 1e-6, "uA": 1e-6, "mA": 1e-3, "A": 1, "kA": 1e3},
    "W": {"mW": 1e-3, "W": 1, "kW": 1e3, "MW": 1e6},
    "VA": {"VA": 1, "kVA": 1e3, "MVA": 1e6},
    "Ω": {"mΩ": 1e-3, "Ω": 1, "kΩ": 1e3, "MΩ": 1e6, "GΩ": 1e9, "옴": 1,
          "ohm": 1, "Ohm": 1, "OHM": 1, "mOhm": 1e-3, "kohm": 1e3, "kOhm": 1e3, "Kohm": 1e3, "KOhm": 1e3,
          "Mohm": 1e6, "MOhm": 1e6, "Gohm": 1e9, "GOhm": 1e9},
    "Hz": {"Hz": 1, "kHz": 1e3, "MHz": 1e6, "GHz": 1e9},
    "°C": {"°C": 1, "도씨": 1},
    "%": {"%": 1},
    "m": {"nm": 1e-9, "μm": 1e-6, "um": 1e-6, "mm": 1e-3, "cm": 1e-2, "m": 1, "km": 1e3},
    "s": {"ns": 1e-9, "μs": 1e-6, "us": 1e-6, "ms": 1e-3, "s": 1, "sec": 1, "초": 1, "min": 60, "분": 60, "h": 3600, "hr": 3600, "시간": 3600},
    "g": {"mg": 1e-3, "g": 1, "kg": 1e3},
    "N": {"N": 1, "kN": 1e3},
    "Pa": {"Pa": 1, "kPa": 1e3, "MPa": 1e6, "bar": 1e5},
    "rpm": {"rpm": 1, "RPM": 1},
    "dB": {"dB": 1},
    "dBm": {"dBm": 1},
    "F": {"pF": 1e-12, "nF": 1e-9, "μF": 1e-6, "uF": 1e-6, "F": 1},
    "lx": {"lx": 1, "lux": 1},
}.items():
    for _u, _s in _units.items():
        _UNIT_TABLE[_u] = (_base, float(_s))
_UNIT_CI: Dict[str, str] = {}
for _u in _UNIT_TABLE:
    if "ohm" in _u.lower():
        continue                    # ohm 계열은 쓴 그대로만 (소문자 'mohm'은 밀리인지 메가인지 모름 -> 단위 모름)
    _UNIT_CI[_u.lower()] = "" if _u.lower() in _UNIT_CI else _u      # 대소문자만 다른 단위가 둘이면(mV, MV) 쓰지 않음


def _unit_key(u: Optional[str]) -> Optional[str]:
    """쓴 단위 글자 -> _UNIT_TABLE 키 (Vdc -> V, KV -> kV). 모르면 None."""
    if not u:
        return None
    for k in (u, re.sub(r"(?i)(dc|ac)$", "", u)):
        if k in _UNIT_TABLE:
            return k
        if _UNIT_CI.get(k.lower()):
            return _UNIT_CI[k.lower()]
    return None


_UNIT_ALT = "|".join(re.escape(u) for u in sorted(_UNIT_TABLE, key=len, reverse=True))
# 단위 뒤에 영문이 이어지면 단위가 아님 (Vdc는 V로 봄: dc/ac는 허용). '초' 뒤 '과'는 '초과'.
# 묶음 단위('Ω·cm', 'V/m', 'mm^2', 'mm2')도 그 단위가 아님 (모르는 단위 -> 조건으로 쓰지 않음)
_UNIT_RE = (rf"(?:{_UNIT_ALT})(?:(?<=초)(?!과)|(?<!초))(?:(?=dc|ac|DC|AC)|(?![A-Za-z]))"
            r"(?!\s*[\u00b7\u22c5*/]\s*[A-Za-z\u03a9\u03bc\u00b0])(?![\^\d])")
# 숫자는 다른 숫자의 가운데서 시작하지 않고 구분자 없이 30자를 넘으면 수치가 아님 (시리얼 번호 등): 긴 숫자열에서
# 모든 자리를 시작점으로 다시 읽어 시간이 제곱으로 늘던 문제 (1.4.69)
# 쉼표 뒤가 3자리 묶음이 아니면('1,2345', '1,5', '10,20') 수치로 보지 않음 (1.4.70)
_NUM_RE = r"[+-]?(?<!\d)(?<!\d,)(?:\d{1,3}(?:,\d{3})+|\d{1,30})(?![\d])(?!,\d)(?:\.\d{1,30}(?!\d))?"
_Q = rf"(?P<{{n}}>{_NUM_RE})\s*(?P<{{u}}>{_UNIT_RE})?"          # 숫자와 (있으면) 단위
# 단위 뒤 꼬리: 'V DC', 'V rms', 'dB(A)', 'mV(p-p)', 'V(DC)' (그 뒤의 이하·이상을 읽으려고)
_AD = (r"(?:\s*(?:AC|DC|ac|dc|rms|RMS))?"
       r"(?:\s*\(\s*(?:A|C|Z|AC|DC|ac|dc|rms|RMS|p-p|P-P|pp|pk|peak|0-p|o-p)\s*\))?")


def _q(n: str, u: str) -> str:
    return _Q.replace("{n}", n).replace("{u}", u)


# 순서가 우선순위 (앞 패턴이 차지한 글은 뒤 패턴이 다시 쓰지 않음)
_COND_PATTERNS: List[Tuple[str, "re.Pattern[str]"]] = [(k, re.compile(p, re.IGNORECASE)) for k, p in (
    ("tolp", rf"\(\s*(?P<a>{_NUM_RE})\s*(?:±|\+/-|\+-)\s*(?P<d>{_NUM_RE})\s*(?P<ud>%)?\s*\)\s*(?P<ua>{_UNIT_RE})"),
    ("between", rf"(?:between|from)\s+{_q('a', 'ua')}{_AD}\s*(?:and|to|~)\s*{_q('b', 'ub')}"),
    ("tol", rf"{_q('a', 'ua')}{_AD}\s*\(?\s*(?:±|\+/-|\+-)\s*(?P<d>{_NUM_RE})\s*(?P<ud>%|{_UNIT_RE})?\s*\)?"),
    # '±5 % 이내', '±0.1 mm 이하' -> -N~N (기준 값이 없는 대칭 한계)
    ("pmrange", rf"(?:±|\+/-|\+-)\s*{_q('a', 'ua')}{_AD}\s*(?:을|를|이|가)?\s*"
                rf"(?:이하|이내|max(?:imum)?\b|or\s+less|넘지\s*않|초과하지\s*않)"),
    ("range", rf"{_q('a', 'ua')}{_AD}\s*(?:~|\u223c|\u301c|\uff5e|\u2013|\u2014|to)\s*{_q('b', 'ub')}(?:\s*(?:까지|사이))?"),
    # 'A에서 B까지'는 범위, 'A에서 B로(를)'는 바뀌는 값('220 V에서 110 V로 강압')이라 범위가 아님
    ("range", rf"{_q('a', 'ua')}{_AD}\s*(?:에서|부터)\s*(?>{_q('b', 'ub')})(?:\s*(?:까지|사이)|(?!\s*(?:으로|로|를|을)))"),
    ("range", rf"(?<![\w.-]){_q('a', 'ua')}\s*-\s*{_q('b', 'ub')}"),
    ("upper", rf"(?:최대|max(?:imum)?\.?|≤|<=|=<|not\s+(?:more|greater)\s+than|no\s+more\s+than|not\s+exceed(?:ing)?|"
              rf"up\s+to|less\s+than\s+or\s+equal\s+to|at\s+most)\s*[:=]?\s*{_q('a', 'ua')}"),
    ("upper_x", rf"(?:<|less\s+than|below|under)\s*{_q('a', 'ua')}"),
    ("lower", rf"(?:최소|min(?:imum)?\.?|≥|>=|=>|not\s+less\s+than|no\s+less\s+than|at\s+least|greater\s+than\s+or\s+equal\s+to)"
              rf"\s*[:=]?\s*{_q('a', 'ua')}"),
    ("lower_x", rf"(?:>|greater\s+than|more\s+than|above|over)\s*{_q('a', 'ua')}"),
    ("upper", rf"{_q('a', 'ua')}{_AD}\s*(?:을|를|이|가)?\s*(?:이하|이내|max(?:imum)?\b|or\s+less|or\s+below|넘지\s*않|초과하지\s*않)"),
    ("upper_x", rf"{_q('a', 'ua')}{_AD}\s*(?:미만|보다\s*작)"),
    ("lower", rf"{_q('a', 'ua')}{_AD}\s*(?:을|를|이|가)?\s*(?:이상|min(?:imum)?\b|or\s+more|or\s+greater|or\s+above)"),
    ("lower_x", rf"{_q('a', 'ua')}{_AD}\s*(?:초과|보다\s*커|보다\s*큰)"),
    ("pm", rf"(?:±|\+/-|\+-)\s*(?P<d>{_NUM_RE})\s*(?P<ud>%|{_UNIT_RE})"),       # 기준 값 없이 떨어진 공차 '±0.1 mm', '±5 %'
    ("plain", rf"(?<![\w.]){_q('a', 'ua')}"),
)]


def _num(text: str) -> float:
    return float(text.replace(",", ""))


_SUPER = "\u2070\u00b9\u00b2\u00b3\u2074\u2075\u2076\u2077\u2078\u2079"       # 위첨자 0~9
_SUPER_MAP = str.maketrans({**{ch: str(i) for i, ch in enumerate(_SUPER)}, "\u207a": "+", "\u207b": "-"})
_MINUS_DASH = "\u2010\u2011\u2012\u2212"                    # 빼기로 쓰는 대시 (U+2013은 숫자 앞에서만)
# 지수는 2자리까지 (1E99). 3자리 이상('1E999', '1×10^999')은 계산하지 않고 _EXP_BIG가 수치가 아닌 글로 바꿈 -> 조건 없음(판단필요)
_EXP_10 = re.compile(r"(?<![\w.])(\d+(?:\.\d+)?)\s*[\u00d7xX*\u00b7]\s*10\s*"
                     rf"(?:\^\s*\(?([+-]?\d{{1,2}})(?!\d)\)?|\*\*\s*([+-]?\d{{1,2}})(?!\d)|"
                     rf"([\u207a\u207b]?[{_SUPER}]{{1,2}})(?![{_SUPER}]))")
_EXP_10_BARE = re.compile(rf"(?<![\w.\u00d7*^])10(?:\s*\^\s*([+-]?\d{{1,2}})(?!\d)|"
                          rf"([\u207a\u207b]?[{_SUPER}]{{1,2}})(?![{_SUPER}]))")
_EXP_E = re.compile(r"(?<![\w.])(\d+(?:\.\d+)?)[eE]([+-]?\d{1,2})(?![\w.])")
_EXP_BIG = re.compile(rf"(?<![\w.])(?:\d+(?:\.\d+)?\s*[\u00d7xX*\u00b7]\s*)?10\s*(?:\^\s*\(?[+-]?\d{{3,}}\)?|\*\*\s*[+-]?\d{{3,}}|"
                      rf"[\u207a\u207b]?[{_SUPER}]{{3,}})|(?<![\w.])\d+(?:\.\d+)?[eE][+-]?\d{{3,}}(?![\w.])")


def _exp_value(mant: str, exp: str) -> str:
    """가수와 10의 지수 -> 지수 없는 숫자 글 ('1', '9' -> '1000000000')."""
    import decimal
    v = decimal.Decimal(mant) * (decimal.Decimal(10) ** int(exp.translate(_SUPER_MAP)))
    out = format(v, "f")
    return out.rstrip("0").rstrip(".") if "." in out else out


def spec_norm(text: Any) -> str:
    """규격 글·칸 값의 수치 표기를 맞춤 (check 판정용). NFKC 전에 처리할 것:
    지수 '1×10'+위첨자 9, '1×10^9', '1E9' -> 계산한 숫자 (NFKC가 위첨자를 '109'로 바꾸기 전에, 지수 2자리까지:
    '1E999'는 '(지수 범위 밖)'으로 바꿔 조건으로 읽지 않음),
    단위 뒤 위첨자 'mm²' -> 'mm^2' (묶음 단위로 보고 조건에 쓰지 않음),
    빼기 대시 U+2212·U+2010~2012, 숫자 바로 앞 U+2013(앞이 숫자·단위가 아닐 때) -> '-'."""
    t = str(text or "")
    t = _EXP_BIG.sub("(지수 범위 밖)", t)
    t = _EXP_10.sub(lambda m: _exp_value(m.group(1), m.group(2) or m.group(3) or m.group(4)), t)
    t = _EXP_10_BARE.sub(lambda m: _exp_value("1", m.group(1) or m.group(2)), t)
    t = _EXP_E.sub(lambda m: _exp_value(m.group(1), m.group(2)), t)
    t = re.sub(rf"(?<=[A-Za-z\u03a9\u2126\u03bc\u00b5])[{_SUPER}]+", lambda m: "^" + m.group(0).translate(_SUPER_MAP), t)
    t = re.sub(f"[{_MINUS_DASH}]", "-", t)
    t = re.sub(r"(?<![\dA-Za-z%\u00b0\u2103\u03a9\u2126\u03bc)\]])(?<![\dA-Za-z%\u00b0\u2103\u03a9\u2126\u03bc)\]]\s)"
               r"\u2013(?=\s?\d)", "-", t)
    return unicodedata.normalize("NFKC", t).replace("\u00b5", "μ")


def unit_of(text: Any) -> Optional[Tuple[str, float, str]]:
    """'mA', '전류(mA)', '[V]', '5 V'의 단위 -> (기준 단위, 배율, 쓴 글자). 모르면 None."""
    t = spec_norm(text)
    m = re.search(rf"[\(\[]\s*({_UNIT_ALT})\s*[\)\]]", t)
    if not m:
        m = re.search(rf"(?:^|[\s\d])({_UNIT_RE})\s*$", t.strip())
    if not m:
        return None
    k = _unit_key(m.group(1).strip())
    return (_UNIT_TABLE[k][0], _UNIT_TABLE[k][1], k) if k else None


def quantity_of(value: Any) -> Optional[Tuple[float, Optional[Tuple[str, float, str]]]]:
    """칸 값 -> (숫자, 단위 또는 None). 숫자 칸, '5.1', '5.1 V', '5,000 rpm'. 숫자가 아니면 None."""
    if isinstance(value, bool) or value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value), None
    t = spec_norm(value).strip()
    m = re.fullmatch(rf"({_NUM_RE})\s*({_UNIT_RE})?\s*", t)
    if not m:
        return None
    unit = None
    k = _unit_key(m.group(2))
    if k:
        unit = (_UNIT_TABLE[k][0], _UNIT_TABLE[k][1], k)
    return _num(m.group(1)), unit


def _cond(kind: str, m: "re.Match[str]") -> Optional[Dict[str, Any]]:
    """패턴 일치 하나 -> 조건 {kind, lo, hi, lo_strict, hi_strict, base, scale, unit, text} (기준 단위 값). 단위가 없으면 None."""
    gd = m.groupdict()
    if kind == "pm":
        # 떨어진 공차: 혼자서는 판정 조건이 아님 (spec_conditions가 바로 앞 기준 값과 합치거나, 못 합치면 판단필요)
        ud = gd.get("ud")
        if ud == "%":
            return {"kind": "pm", "lo": None, "hi": None, "lo_strict": False, "hi_strict": False, "base": "%", "scale": 1.0,
                    "unit": "%", "d": _num(gd["d"]), "pct": True, "text": m.group(0).strip()}
        key = _unit_key(ud)
        if key is None:
            return None
        base, scale = _UNIT_TABLE[key]
        return {"kind": "pm", "lo": None, "hi": None, "lo_strict": False, "hi_strict": False, "base": base, "scale": scale,
                "unit": key, "d": _num(gd["d"]) * scale, "pct": False, "text": m.group(0).strip()}
    ua, ub = gd.get("ua"), gd.get("ub")
    unit = ua or ub or (gd.get("ud") if gd.get("ud") and gd.get("ud") != "%" else None)
    if kind == "tol" and not unit and gd.get("ud") == "%":
        return None                                 # '5 ± 5%'처럼 기준 값의 단위가 없으면 판단하지 않음
    if not unit:
        return None
    key = _unit_key(unit)
    if key is None:
        return None
    base, scale = _UNIT_TABLE[key]

    def val(num: str, u: Optional[str]) -> float:
        k = _unit_key(u) if u else key
        if k is None or _UNIT_TABLE[k][0] != base:
            raise ValueError("단위가 다름")
        return _num(num) * _UNIT_TABLE[k][1]
    try:
        a = val(gd["a"], ua)
        c: Dict[str, Any] = {"kind": kind, "lo": None, "hi": None, "lo_strict": False, "hi_strict": False,
                             "base": base, "scale": scale, "unit": key, "text": m.group(0).strip()}
        if kind in ("range", "between"):
            b = val(gd["b"], ub or ua)
            c.update(lo=min(a, b), hi=max(a, b), kind="range")
        elif kind == "pmrange":
            c.update(lo=-abs(a), hi=abs(a), kind="range")
        elif kind in ("tol", "tolp"):
            if kind == "tol" and ua and not gd.get("ud"):
                # '100 mm ±0.5': 공차에 단위가 없음 (mm인지 %인지 글만으로 확실하지 않음)
                c["ambiguous"] = "단위 없는 공차"
            c["kind"] = "tol"
            d = _num(gd["d"])
            ud = gd.get("ud")
            d = abs(a) * d / 100 if ud == "%" and base != "%" else val(gd["d"], ud) if ud else d * scale
            c.update(lo=a - d, hi=a + d)
        elif kind.startswith("upper"):
            c.update(hi=a, hi_strict=kind.endswith("_x"), kind="upper")
        elif kind.startswith("lower"):
            c.update(lo=a, lo_strict=kind.endswith("_x"), kind="lower")
        else:
            c.update(lo=a, hi=a)
    except (ValueError, KeyError, TypeError):
        return None
    return c


# 조건 사이를 잇는 말 (이것만 있으면 같은 항목): '5 mm 이상이며 10 mm 이하', '5 mm 이상, 10 mm 이하'
_COND_JOIN = re.compile(r"(?:으로|로)\s*(?:하고|하며|하되)|이고|이며|이면서|이되|및|그리고|또한|and|or|또는|[\s,;:()\[\]~-]", re.IGNORECASE)
# 앞 항목의 한계를 말하는 낱말 (새 항목이 아님): '정격 전압 5 V, 허용 범위 4.75~5.25 V', '변동 범위는 ...'
# '허용 전류 10 A 이하'처럼 허용·변동 뒤에 다른 낱말이 오면 새 항목.
# '사용 범위', '동작 범위', '측정 범위'는 한계 낱말이 아니라 별개 항목 ('시험 전압 500 V, 사용 범위 0~1000 V')
_LIMIT_WORD = (r"(?:(?:허용|변동)\s*)?(?:범위|공차|오차|편차)|허용\s*(?:차|치|값)|변동\s*폭"
               r"|(?:허용|변동)(?!\s*[가-힣A-Za-z])|tolerance|range|deviation")
_LIMIT_PHRASE = re.compile(rf"(?:{_LIMIT_WORD})\s*(?:은|는|이|가|의|을|를|에서|내)?", re.IGNORECASE)


_AC_DC = re.compile(r"(?<![A-Za-z])(?:AC|DC)(?![A-Za-z])", re.IGNORECASE)     # 'DC 24 V(허용 범위 DC 20.4 V ~ 26.4 V)'


def _other_item_between(t: str, a: int, b: int) -> bool:
    """t[a:b](두 조건 사이)에 다른 항목 낱말이 있는지: 잇는 말과 문장 부호를 빼고 글자가 남으면 있음
    ('5 mm 이상, 길이 100 mm 이하'의 '길이')."""
    gap = _COND_JOIN.sub("", _AC_DC.sub(" ", _LIMIT_PHRASE.sub(" ", t[min(a, b):max(a, b)])))
    return bool(re.search(r"[가-힣A-Za-z]", gap))


def spec_conditions(text: str) -> List[Dict[str, Any]]:
    """규격 글의 수치 조건 목록 (글 순서). 단위가 있는 수치만 (규격 번호 'KS C 1234', 장절 '3.2'는 빼려고).
    여러 조건을 하나로 합치는 것은 확실한 규칙 셋뿐입니다 (보수적 판정: 나머지는 그대로 여러 개 -> check가 판단필요):
    (a) 같은 단위의 아래·위 한계가 하나씩: '4.75 V 이상 5.25 V 이하' -> 4.75~5.25 V
    (b) 기준 값 바로 뒤의 떨어진 공차 '±N 단위'·'±N %' (사이에 잇는 말·괄호·허용 오차 같은 한계 낱말만): '두께 3 mm, 공차 ±0.1 mm'
    (c) 기준 값 뒤의 양쪽 한계 범위가 기준 값을 품고, 범위 앞에 한계 낱말('허용 범위' 등)이 있거나 범위가 기준 값 바로 뒤
        괄호 안에 그것만 있음: '정격 5 V, 허용 범위 4.75~5.25 V', '정격 5 V (4.75~5.25 V)'.
        '5 V, 4~6 V에서 동작', '0~10 A까지 측정 가능'처럼 한계 낱말이 없거나 쓰임새가 이어지면 합치지 않음
    두 조건 사이에 다른 항목 낱말이 있으면('두께 5 mm 이상, 길이 100 mm 이하') 합치지 않습니다.
    한쪽 한계와 기준 값('5 V, 허용 범위 4.75 V 이상')은 합치지 않습니다."""
    t = spec_norm(text)
    taken: List[Tuple[int, int]] = []
    found: List[Tuple[int, int, Dict[str, Any]]] = []
    for kind, pat in _COND_PATTERNS:
        for m in pat.finditer(t):
            s, e = m.span()
            if any(s < b and a < e for a, b in taken):
                continue
            c = _cond(kind, m)
            if c is None:
                continue
            taken.append((s, e))
            found.append((s, e, c))
    found.sort(key=lambda x: x[0])
    spans = {id(c): (s, e) for s, e, c in found}
    conds = [c for _, _, c in found]

    def apart(x: Dict[str, Any], y: Dict[str, Any]) -> bool:
        (xs, xe), (ys, ye) = spans[id(x)], spans[id(y)]
        return _other_item_between(t, xe, ys) if xs <= ys else _other_item_between(t, ye, xs)

    def joined(x: Dict[str, Any], y: Dict[str, Any], **kw: Any) -> Dict[str, Any]:
        c = dict(x, **kw)
        spans[id(c)] = (min(spans[id(x)][0], spans[id(y)][0]), max(spans[id(x)][1], spans[id(y)][1]))
        return c
    # (a)
    lows = [c for c in conds if c["kind"] == "lower"]
    highs = [c for c in conds if c["kind"] == "upper"]
    if len(lows) == 1 and len(highs) == 1 and lows[0]["base"] == highs[0]["base"] and lows[0]["lo"] <= highs[0]["hi"] \
            and not apart(lows[0], highs[0]):
        lo, hi = lows[0], highs[0]
        merged = joined(lo, hi, kind="range", hi=hi["hi"], hi_strict=hi["hi_strict"], text=f"{lo['text']} ~ {hi['text']}")
        i = min(conds.index(lo), conds.index(hi))
        conds = [c for c in conds if c is not lo and c is not hi]
        conds.insert(i, merged)
    # (b), (c): 기준 값(plain)과 바로 다음 조건
    out: List[Dict[str, Any]] = []
    i = 0
    while i < len(conds):
        c = conds[i]
        nxt = conds[i + 1] if i + 1 < len(conds) else None
        if c["kind"] == "plain" and nxt is not None and not apart(c, nxt):
            v = c["lo"]
            if nxt["kind"] == "pm" and (nxt["pct"] and c["base"] != "%" or not nxt["pct"] and nxt["base"] == c["base"]):
                d = abs(v) * nxt["d"] / 100 if nxt["pct"] else nxt["d"]
                out.append(joined(c, nxt, kind="tol", lo=v - d, hi=v + d, text=f"{c['text']} {nxt['text']}"))
                i += 2
                continue
            if nxt["kind"] == "range" and nxt["base"] == c["base"] and nxt["lo"] != nxt["hi"] and cond_holds(nxt, v) \
                    and (_LIMIT_PHRASE.search(t[spans[id(c)][1]:spans[id(nxt)][0]])
                         or (re.fullmatch(r"\s*\(\s*", t[spans[id(c)][1]:spans[id(nxt)][0]])
                             and re.match(r"\s*\)", t[spans[id(nxt)][1]:]))) \
                    and not _RANGE_USE.match(t, spans[id(nxt)][1]):
                out.append(nxt)
                i += 2
                continue
        out.append(c)
        i += 1
    res, seen = [], set()
    for c in out:
        key = (c["kind"] == "pm", c["base"], c["lo"], c["hi"], c["lo_strict"], c["hi_strict"], c.get("d"))
        if key not in seen:
            seen.add(key)
            res.append(c)
    why = _shape_doubt(t, [sp for sp in taken])
    if why:
        for c in res:
            c.setdefault("ambiguous", why)
    return res


# 범위 뒤에 이어지는 쓰임새('0~10 A까지 측정 가능', '18~36 V 입력 가능'): 허용 범위가 아니라 기능 설명
_RANGE_USE = re.compile(r"\s*(?:까지|에서)?\s*(?:측정|입력|출력|표시|설정|조정|사용|동작|선택)", re.IGNORECASE)
_BARE_NUM = re.compile(r"(?<![\w.,])(?<![A-Za-z]\s)[+-]?\d+(?:\.\d+)?(?![\d.,]*\d)")
_NUM_FOLLOW = re.compile(r"\s*(?:[^\W\d_]|[%\u00b0])")      # 숫자 뒤에 낱말·단위·기호가 붙음 ('3종', '1분', '5%')
_NUM_LABEL = re.compile(r"(?:표|그림|부록|별표|Table|Fig\.?|Figure|No\.?|Note)\s*$", re.IGNORECASE)
_DIMENSIONS = re.compile(r"\d\s*(?:[A-Za-z]{1,3}\s*)?[\u00d7xX]\s*(?:[A-Za-z]{1,2}\s*)?\d")
_DECIMAL_COMMA = re.compile(r"(?<![\d,])\d+,(?:\d{1,2}|\d{4,})(?!\d)")
_ASYM_TOL = re.compile(r"\+\s*\d+(?:\.\d+)?\s*(?:[%A-Za-z\u03a9\u03bc]{1,3}\s*)?[/,]?\s*-\s*\d")


def _shape_doubt(t: str, taken: List[Tuple[int, int]]) -> str:
    """조건 글 모양 때문에 판정에 바로 쓰면 안 되는 이유 (없으면 ''): 치수 나열 '100×200 mm', 소수점 쉼표 '1,5 mm',
    비대칭 공차 '+0.5/-0.2', 단위 있는 조건 곁의 단위 없는 수치 '4.75 5.0 5.25 V', '100 mm, ±0.5'."""
    if _DIMENSIONS.search(t):
        return "치수 나열(×)인 문장"
    if _DECIMAL_COMMA.search(t):
        return "소수점 쉼표('1,5')일 수 있는 수치"
    if _ASYM_TOL.search(t):
        return "비대칭 공차(+a/-b)인 문장"
    if taken:
        for m in _BARE_NUM.finditer(t):
            s, e = m.span()
            if any(a <= s < b for a, b in taken) or _NUM_FOLLOW.match(t, e) or _NUM_LABEL.search(t, 0, s) \
                    or re.search(r"\(\s*$", t[:s]) and re.match(r"\s*\)", t[e:]):
                continue                                 # 번호 '(1)', '표 2', 'Fig. 3'은 수치가 아님
            return f"단위 없는 수치({m.group(0)})가 함께 있음"
    return ""


def cond_text(c: Dict[str, Any]) -> str:
    """조건을 사람이 읽는 글로 (쓴 단위 기준). 예: '4.75~5.25 V', '100 MΩ 이상', '3 A 미만', '±0.1 mm'(합치지 못한 공차)."""
    s = c["scale"] or 1.0
    f = lambda v: f"{v / s:.6g}"  # noqa: E731
    u = c["unit"]
    if c["kind"] == "pm":
        return f"±{f(c['d'])} {u}"
    if c["lo"] is not None and c["hi"] is not None:
        return f"{f(c['lo'])} {u}" if c["lo"] == c["hi"] else f"{f(c['lo'])}~{f(c['hi'])} {u}"
    if c["hi"] is not None:
        return f"{f(c['hi'])} {u} {'미만' if c['hi_strict'] else '이하'}"
    return f"{f(c['lo'])} {u} {'초과' if c['lo_strict'] else '이상'}"


def cond_holds(c: Dict[str, Any], v: float) -> bool:
    eps = 1e-9 * max(1.0, abs(v), abs(c["lo"] or 0), abs(c["hi"] or 0))
    if c["lo"] is not None and (v < c["lo"] - eps or (c["lo_strict"] and v <= c["lo"] + eps)):
        return False
    if c["hi"] is not None and (v > c["hi"] + eps or (c["hi_strict"] and v >= c["hi"] - eps)):
        return False
    return True


# 다음 항목이 시작하는 자리: 절 경계(쉼표, ;, 이며, 및 ...) 뒤에 낱말(+조사)이 오고 바로 수치가 이어짐
# ('두께 5 mm 이상, 길이 100 mm 이하'의 ', 길이 '). 최대·최소 같은 조건 낱말은 새 항목이 아님
_NEXT_ITEM = re.compile(
    r"(?:[,;]|\s(?:이며|이고|이면서|및|그리고|또한)|(?<=[가-힣])(?:이며|이고|이면서))\s*"
    r"(?!(?:최대|최소|약|max|min|단|또는|or|and)\b)(?!(?:최대|최소|약)\s)"
    rf"(?!(?:{_LIMIT_WORD}))"
    r"[가-힣A-Za-z][가-힣A-Za-z0-9 ]{0,30}?(?:은|는|이|가|을|를|의|:)?\s*(?=[+-]?\d|최대|최소|max|min|≤|≥|<|>)",
    re.IGNORECASE)
_CLAUSE_END = re.compile(r"[,;]|\s(?:이며|이고|및|그리고)\s|(?<=[가-힣])(?:이며|이고)\s")


# 판정 조건으로 바로 쓰면 안 되는 문장: 불합격(거부) 기준, 값이 여럿('50/60 Hz')
_REJECT_WORDS = re.compile(r"불합격|부적합|불량|reject|fail", re.IGNORECASE)
# 부정문: '1 mA 이상이 되어서는 안 된다', '2 mm 미만이어서는 아니 된다' (조건을 뒤집어 읽어야 함)
_NEGATION = re.compile(r"서는\s*(?:안|아니)\s*(?:된|되|됨|함)|아니\s*(?:된|되|됨)|안\s*된다")
_MULTI_VALUE = re.compile(r"\d\s*/\s*\d")
# 기준 값만 있을 때 한계인지 대략인지 정할 수 없는 낱말: '최소 두께 2 mm', '최대 전류 3 A', 'Typ. 5 V', '약 5 kg', '5 kg 내외'
_HEDGE_WORDS = re.compile(r"최소|최대|최고|최저|상한|하한|내외|정도|가량|전후|대략|(?<![가-힣])약\s*[\d+-]"
                          r"|(?<![A-Za-z])(?:max(?:imum)?|min(?:imum)?|typ(?:ical)?|approx(?:imately)?|about|upper|lower"
                          r"|nominal|nom)(?![a-z])", re.IGNORECASE)
# 상황마다 다른 조건: '대기 시 10 mA 이하, 동작 시 100 mA 이하', '85 ℃ 초과 시 동작'
_WHEN_WORDS = re.compile(r"[가-힣]\s시(?:에는|에|의)?(?=[\s,)]|$)")
# 단서 조항: '10 Ω 이하. 다만, 특별 3종은 100 Ω 이하', '5 mm 이상 (단, 끝부분은 10 mm 이하)'
_PROVISO = re.compile(r"(?:^|[\s(\[])(?:다만|단)\s*[,:]|^\s*다만|(?<![A-Za-z])(?:however|except|provided\s+that)(?![a-z])",
                      re.IGNORECASE)
# 문장 나누기 (check용): 'Max. 3A', 'min. 100 MΩ'의 약어 점에서는 나누지 않음
_COND_SENT_SPLIT = re.compile(r"(?<=[.!?\u3002])(?<![Mm][Aa][Xx]\.)(?<![Mm][Ii][Nn]\.)(?<![Tt][Yy][Pp]\.)\s+|\n+")


NEAR_PRE_CHARS = 200           # 낱말 앞 같은 절에서 조건을 읽을 글자 수


def near_conditions(text: str, terms: List[str]) -> Tuple[List[Dict[str, Any]], str]:
    """찾은 낱말 가까이의 수치 조건과 그 문장. 낱말이 든 문장에서 낱말부터 다음 항목 낱말(절 경계 뒤 낱말 + 수치) 직전까지
    (200자까지, '항목 절')를 먼저, 없으면 낱말이 든 절(앞쪽 절 경계부터 낱말까지). 같은 문장의 다른 항목 조건을 섞지 않으려고.
    조건에 'ambiguous'(판정에 바로 쓰면 안 되는 이유)나 'alt'(같은 문장 다른 항목의, 기준 값을 품는 범위)를 붙일 수 있습니다."""
    flags = re.IGNORECASE
    sents = _COND_SENT_SPLIT.split(text)
    lits = {t: literal_finder(t) for t in terms}
    cands = []
    # 숫자가 든 검색어는 층 A(글자 그대로) 자리가 든 문장을 먼저, 그 자리에서 조건을 읽음 (search 스니펫과 같은 자리, 1.4.71)
    for lit_only in ((True, False) if any(lits.values()) else (False,)):
        for si, raw in enumerate(sents):
            # 빼기 대시(U+2212 등)를 '-'로 (글자 수가 같아 위치가 그대로): 다음 항목 찾기의 '[+-]?\d'가 맞게. 돌려주는 문장은 원문
            sent = re.sub(f"[{_MINUS_DASH}]", "-", raw)
            if lit_only and not any(f.search(raw) for f in lits.values() if f is not None):
                continue
            ms = [sp for t in terms for sp in [term_span(t, raw, flags, sent)] if sp]
            if ms:
                cands.append((si, raw, sent, ms))
                break
        if cands:
            break
    for si, raw, sent, ms in cands:
        start = min(m[0] for m in ms)
        end_term = max(m[1] for m in ms if m[0] == start)
        window = sent[start:start + 200]
        nxt = _NEXT_ITEM.search(window, end_term - start)
        clause = window[:nxt.start()] if nxt else window
        after = spec_conditions(clause)
        if nxt and after and all(c["kind"] == "plain" for c in after):
            # 안전장치: 항목 절에 기준 값만 남았는데 같은 문장 뒤쪽 항목에 같은 단위의 범위(아래·위 한계 둘 다)가 있고
            # 그 범위가 기준 값을 품으면 그 범위가 이 항목의 허용 범위일 수 있음 -> 기준 값만으로 불일치라 단정하지 않게 표시.
            # 이름에 '범위'가 든 뒤 항목('사용 범위', '동작 범위', '측정 범위')은 별개 항목이 분명해 보지 않음
            marks = list(_NEXT_ITEM.finditer(window, nxt.start()))
            rest: List[Dict[str, Any]] = []
            for k, m in enumerate(marks):
                if "범위" in m.group(0):
                    continue
                seg = window[m.start():marks[k + 1].start() if k + 1 < len(marks) else len(window)]
                rest += [c for c in spec_conditions(seg)
                         if c["lo"] is not None and c["hi"] is not None and c["lo"] != c["hi"]]
            for c in after:
                alt = [r for r in rest if r["base"] == c["base"] and cond_holds(r, c["lo"])]
                if alt:
                    c["alt"] = alt
        cuts = [m.end() for m in _CLAUSE_END.finditer(sent[:start])]
        pre = sent[cuts[-1] if cuts else 0:start]          # 낱말 앞 같은 절 ('최소 두께'를 '두께'로 찾았을 때 '최소')
        pre = pre[-NEAR_PRE_CHARS:]                        # 끝 200자만 (구분자 없는 긴 글에서 시간이 늘지 않게)
        if not after:
            clause = pre
            after = spec_conditions(clause)
        nxt_sent = sents[si + 1] if si + 1 < len(sents) else ""
        why = ("불합격(거부) 기준을 말하는 문장" if _REJECT_WORDS.search(clause)
               else "부정문(~해서는 안 된다)" if _NEGATION.search(clause)
               else "값이 여럿('/')인 문장" if _MULTI_VALUE.search(unicodedata.normalize("NFKC", clause))
               else "상황마다 다른 조건('~ 시')" if _WHEN_WORDS.search(pre + clause)
               else "단서 조항(다만·단)이 있는 문장" if _PROVISO.search(sent, start) or _PROVISO.match(nxt_sent)
               else "한계인지 대략인지 정할 수 없는 기준 값(최대·최소·Typ·약 등)"
               if after and all(c["kind"] == "plain" for c in after) and _HEDGE_WORDS.search(pre + " " + clause)
               else "")
        if why:
            for c in after:
                c["ambiguous"] = why
        return after, raw.strip()
    return spec_conditions(text[:400]), text.strip()[:400]


def find_item(docs: List[Dict[str, Any]], query: str, sec_key: str = "") -> Tuple[str, List[Dict[str, Any]]]:
    """등록 문서에서 항목 글을 찾습니다 (search와 같은 규칙: 띄어쓰기·조사·단위 표기 무시, 동의어, 0건이면 partial -> fuzzy).
    반환 (match_mode, 결과 [{doc, unit, terms, score, literal}]): 글자 그대로 일치(literal, 층 A)를 먼저, 그 안에서 점수 순.
    숫자가 든 항목 글이 표기가 다른 곳(층 B)에서만 맞으면 literal False (match_mode는 그대로 'exact'). 0건이면 ('exact', [])."""
    flags = re.IGNORECASE
    matcher = build_matcher(query, False)
    hits: List[Dict[str, Any]] = []
    pool: List[Tuple[Dict[str, Any], Dict[str, Any]]] = []
    for doc in docs:
        for u in doc["units"]:
            if sec_key and not section_matches(u, sec_key):
                continue
            pool.append((doc, u))
            terms = matcher(u["text"])
            if terms is not None:
                hits.append({"doc": doc, "unit": u, "terms": terms, "literal": literal_hit(u["text"], terms)})
    mode = "exact"
    status: Dict[str, Any] = {}
    if not hits and pool:
        mode, hits = widen_search(pool, parse_query(query), flags, status)
        if not hits and status.get("cut"):
            mode = "number"           # 숫자 안전망이 시간 상한으로 끝까지 보지 못함: check는 '근거없음'이 아니라 판단필요
    for h in hits:
        h["score"] = hit_score(h["unit"], h["terms"], flags)
    hits.sort(key=lambda h: (not h.get("literal", True), -h["score"]))
    return mode, hits


def load_spec_docs(spec: str, max_files: int = MAX_FOLDER_FILES) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    """--spec 파일 또는 폴더의 문서 (등록되지 않았거나 바뀐 문서는 먼저 등록, R10 CONTINUE 포함)."""
    import types as _types
    return _search_targets(_types.SimpleNamespace(file=spec, max_files=max_files))


# ==============================================================================
# 9. 진입점
# ==============================================================================

def build_cli():
    parser, sub = rt.build_parser(TOOL, "문서 파일 읽기, 등록(캐시), 근거 검색 (Office 불필요)")

    p = sub.add_parser("read", help="파일 하나 읽기 (캐시에 등록)")
    p.add_argument("--file", required=True, help="문서 경로 (파일명만 쓰면 data 폴더에서 찾음)")
    p.add_argument("--pages", help="PDF 쪽/PPT 슬라이드 번호 (그 외 형식은 항목 번호). 예: 3, 1-3, 1,4")
    p.add_argument("--units", help="항목(단위) 번호로 고르기. 예: 57-120 (이어 읽기 명령이 씀)")
    p.add_argument("--section", help="장절 경로 일부가 맞는 단위만 읽기 (대소문자 무시, search --section과 같음)")
    p.add_argument("--format", choices=["text", "summary"], default="text", help="text: 본문, summary: 요약 정보만")
    p.add_argument("--save", nargs="?", const=True, help="결과를 파일로 저장 (.md 기본, .csv/.xlsx 가능). 경로 생략 시 대상 파일 옆, 없으면 C:\\AIKIT\\data")

    o = sub.add_parser("outline", help="문서의 장절(제목) 목록: 제목마다 units·page·level (큰 문서에서 필요한 장절만 읽기)")
    o.add_argument("--file", required=True, help="문서 경로")
    o.add_argument("--start", type=int, default=1, help="몇 번째 제목부터 (이어 보기, 기본 1)")

    x = sub.add_parser("index", help="폴더/파일의 문서를 내용 출력 없이 캐시에 등록")
    x.add_argument("--folder", help="등록할 폴더")
    x.add_argument("--file", help="등록할 파일 하나")
    x.add_argument("--all", action="store_true", help="등록된 문서 전체를 다시 확인 (바뀐 파일, 이전 형식 캐시만 다시 읽음)")
    x.add_argument("--depth", type=int, default=MAX_FOLDER_DEPTH, help=f"하위 폴더 깊이 (기본 {MAX_FOLDER_DEPTH}, 0=그 폴더만)")
    x.add_argument("--recursive", action="store_true", help="하위 폴더 전체")
    x.add_argument("--max-files", type=int, default=MAX_FOLDER_FILES,
                   help=f"이번에 새로 읽을 최대 파일 수 (기본 {MAX_FOLDER_FILES}, 이미 최신인 파일은 세지 않음)")

    s = sub.add_parser("search", help="캐시에 등록된 문서에서 검색 (--file이면 그 파일/폴더를 먼저 자동 등록)")
    s.add_argument("--query", required=True, help="검색어. 그리고는 'A;B', 또는은 'A|B' (;가 먼저 묶임)")
    s.add_argument("--file", help="검색할 파일 또는 폴더 (미등록/바뀐 문서는 자동 등록). 생략 시 캐시 전체")
    s.add_argument("--pages", help="쪽 번호로 거르기. 예: 3-7 (쪽 정보가 없는 문서는 항목 번호, read --pages와 같음)")
    s.add_argument("--section", help="장절 경로 일부로 거르기 (대소문자 무시)")
    s.add_argument("--regex", action="store_true", help="정규식으로 검색")
    s.add_argument("--top", type=int, default=10, help="표시할 최대 건수 (기본 10, --save 파일에는 전체)")
    s.add_argument("--context", type=int, default=80, help="스니펫 앞뒤 글자 수 (기본 80)")
    s.add_argument("--full", action="store_true", help="스니펫 대신 일치한 단락/페이지 전체")
    s.add_argument("--max-files", type=int, default=MAX_FOLDER_FILES,
                   help=f"--file 폴더에서 이번에 새로 등록할 최대 파일 수 (기본 {MAX_FOLDER_FILES})")
    s.add_argument("--save", nargs="?", const=True, help="결과를 파일로 저장 (.md 기본, .xlsx/.csv 가능)")

    lst = sub.add_parser("list", help="캐시에 등록된 문서 목록")
    lst.add_argument("--name", help="파일 이름/경로 일부로 거르기 (대소문자 무시)")
    lst.add_argument("--folder", help="이 폴더 아래 문서만")

    c = sub.add_parser("clear", help="캐시 지우기 (원본 문서는 그대로). 다시 검색하면 새로 읽음")
    c.add_argument("--file", help="이 문서의 캐시만")
    c.add_argument("--name", help="파일 이름/경로 일부가 맞는 문서")
    c.add_argument("--folder", help="이 폴더 아래 문서")
    c.add_argument("--missing", action="store_true", help="원본 파일이 없어진 문서만")
    c.add_argument("--all", action="store_true", help="캐시 전체")
    c.add_argument("--dry-run", action="store_true", help="지우지 않고 지울 목록만 보기")
    return parser


def main(argv: Optional[List[str]] = None) -> None:
    args = build_cli().parse_args(argv)
    handlers = {"read": cmd_read, "index": cmd_index, "search": cmd_search, "list": cmd_list, "clear": cmd_clear,
                "outline": cmd_outline}
    rt.run_cli(TOOL, args.command, lambda: handlers[args.command](args))


if __name__ == "__main__":
    main()
