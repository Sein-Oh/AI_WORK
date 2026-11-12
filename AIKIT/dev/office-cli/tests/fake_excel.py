"""메모리 셀 격자로 동작하는 가짜 Excel (excel.py 명령 로직 테스트용).

지원: Range("A1:B2" / "A1:B2,A4:B4") / Range(cell, cell) / Cells(r, c), Value(읽기/쓰기, 스칼라 채우기),
Formula(쓰기 시 상대참조 문자열 그대로 저장), HasFormula/HasArray/CurrentArray(배열 수식 arrays), Row/Column/Rows.Count/Columns.Count/Address, UsedRange,
End(xlUp), 서식 속성 기록(Interior/Font/Borders/NumberFormat ...), ChartObjects, Cut/Copy/Merge,
EntireRow/EntireColumn Insert/Delete(칸 옮기기), Sort(Key1~Key3, Header), Worksheet.Sort(SortFields, SetRange, Apply),
메모(Range.Comment/AddComment, Comment.Text()/Delete(), Worksheet.Comments).
"""
import re
import types


def col_num(s):
    n = 0
    for ch in s.upper():
        n = n * 26 + ord(ch) - 64
    return n


def col_let(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


_REF = re.compile(r"(?<![A-Za-z$])([A-Z]{1,3})(\d+)(?![\d(])")


def shift_formula(formula, dr, dc):
    """상대 참조(A1)를 dr행, dc열 옮긴 수식."""
    return _REF.sub(lambda m: f"{col_let(col_num(m.group(1)) + dc)}{int(m.group(2)) + dr}", formula)


def to_r1c1(formula, r, c):
    return _REF.sub(lambda m: f"R[{int(m.group(2)) - r}]C[{col_num(m.group(1)) - c}]", formula)


class Props:
    """임의 속성을 기록하는 객체 (Interior.Color, Font.Bold 등)."""

    def __init__(self):
        object.__setattr__(self, "_d", {})

    def __getattr__(self, k):
        d = object.__getattribute__(self, "_d")
        if k not in d:
            d[k] = Props()
        return d[k]

    def __setattr__(self, k, v):
        object.__getattribute__(self, "_d")[k] = v

    def get(self, k, default=None):
        return object.__getattribute__(self, "_d").get(k, default)


class Count:
    def __init__(self, n):
        self.Count = n


class FakeFormatConditions:
    def __init__(self):
        self.items = []

    def Add(self, **kw):
        fc = Props()
        fc.kw = kw
        self.items.append(fc)
        return fc

    def Delete(self):
        self.items.clear()

    @property
    def Count(self):
        return len(self.items)


class SheetNameRefused(Exception):
    """실제 Excel이 시트 이름을 거절할 때의 COM 오류 (쓸 수 없는 글자, 31자 초과, 예약된 이름 History)."""


def check_fake_sheet_name(name):
    if re.search(r"[\[\]:*?/\\]", name) or len(name) > 31 or name.lower() == "history":
        raise SheetNameRefused(-2146827284, "Exception occurred.", (0, "Microsoft Excel", "This name is reserved.", None, 0, 0), None)


class FakeLines:
    """EntireRow / EntireColumn: Hidden은 시트의 숨긴 행·열 집합으로 (일부만 숨기면 None, 실제 Excel처럼)."""

    def __init__(self, ws, kind, keys):
        object.__setattr__(self, "ws", ws)
        object.__setattr__(self, "kind", kind)
        object.__setattr__(self, "keys", keys)

    @property
    def Hidden(self):
        hidden = self.ws.hidden[self.kind]
        n = sum(1 for k in self.keys if k in hidden)
        return True if n == len(self.keys) else (False if n == 0 else None)

    def __setattr__(self, k, v):
        if k == "Hidden":
            for key in self.keys:
                (self.ws.hidden[self.kind].add if v else self.ws.hidden[self.kind].discard)(key)
            self.ws.log.append(("Hidden", self.kind, tuple(self.keys), v))
        else:
            object.__setattr__(self, k, v)

    def _shift(self, start, n):
        """start행(열) 이상의 칸을 n만큼 옮김 (n < 0이면 당김). 수식 글의 참조는 고치지 않음."""
        i = 0 if self.kind == "rows" else 1
        for store in (self.ws.values, self.ws.formulas):
            moved = {}
            for key, v in store.items():
                if key[i] >= start:
                    nk = list(key)
                    nk[i] += n
                    moved[tuple(nk)] = v
                else:
                    moved[key] = v
            store.clear()
            store.update(moved)

    def Insert(self, Shift=None, CopyOrigin=None):
        """행·열 넣기: 첫 행(열) 앞에 len(keys)개를 넣음 (실제 Excel처럼 연속 범위만)."""
        self.ws.log.append(("Insert", self.kind, tuple(self.keys)))
        self._shift(self.keys[0], len(self.keys))

    def Delete(self, Shift=None):
        self.ws.log.append(("Delete", self.kind, tuple(self.keys)))
        i = 0 if self.kind == "rows" else 1
        lo, hi = self.keys[0], self.keys[-1]
        for store in (self.ws.values, self.ws.formulas):
            for key in [k for k in store if lo <= k[i] <= hi]:
                del store[key]
        self._shift(hi + 1, -len(self.keys))


_TYPED_NUM = re.compile(r"^[+-]?\d+(?:\.\d+)?(?:[eE][+-]?\d+)?$")
_TYPED_DATE = re.compile(r"^\d{1,2}\s*[-/]\s*\d{1,2}$")


def _typed(ws, v):
    """ws.typing이 True면 실제 Excel처럼 '·= 없이 대입한 글자를 직접 입력한 것처럼 해석 (007 -> 7, 1-2 -> 날짜).
    기본은 False (대입한 값 그대로)."""
    if not getattr(ws, "typing", False) or not isinstance(v, str) or v.startswith(("'", "=")):
        return v
    t = v.strip()
    if _TYPED_NUM.match(t):
        num = float(t)
        return int(num) if num.is_integer() and abs(num) < 1e15 else num
    if _TYPED_DATE.match(t):
        return 46000.0                             # 날짜 일련번호
    if t[:1] in ("+", "-", "@"):                   # Excel은 수식으로 계산 (+82-10-1234-5678 -> -6840), 못 하면 #NAME?
        if re.fullmatch(r"[+\-]?[\d.\s+\-*/()]+", t):
            try:
                num = float(eval(t, {"__builtins__": {}}))      # noqa: S307  시험용 산술식만
                return int(num) if num.is_integer() else num
            except Exception:
                pass
        return -2146826259                         # #NAME? (pywin32는 오류 칸을 정수로 줌)
    return v


class FakeComment:
    """칸 메모(이전 방식 메모). Text(): 읽기, Text(Text=글): 바꾸기, Text(Text=글, Start=n): n번째 글자 앞에 넣기."""

    def __init__(self, ws, r, c, text, author="검토자"):
        self.ws = ws
        self.rc = (r, c)
        self.text = text
        self.Author = author

    @property
    def Parent(self):
        return FakeRange(self.ws, [self.rc * 2])

    def Text(self, Text=None, Start=None, Overwrite=False):
        if Text is not None:
            if self.ws.refuse_notes:
                raise Exception("시트 보호")
            if Start is None:
                self.text = Text
            else:
                i = Start - 1
                self.text = self.text[:i] + Text + self.text[i + (len(Text) if Overwrite else 0):]
            self.ws.log.append(("CommentText", self.rc))
        return self.text

    def Delete(self):
        self.ws.log.append(("CommentDelete", self.rc))
        self.ws.notes.pop(self.rc, None)


class FakeComments:
    """Worksheet.Comments: 실제 Excel처럼 시트의 메모를 행 순서로."""

    def __init__(self, ws):
        self.ws = ws

    def _items(self):
        return [self.ws.notes[k] for k in sorted(self.ws.notes)]

    @property
    def Count(self):
        return len(self.ws.notes)

    def Item(self, i):
        return self._items()[i - 1]

    def __call__(self, i):
        return self.Item(i)


class FakeRange:
    def __init__(self, ws, areas):
        self.ws = ws
        self.areas = areas          # [(r1, c1, r2, c2)]

    # --- 위치 ---
    @property
    def Row(self):
        return self.areas[0][0]

    @property
    def Column(self):
        return self.areas[0][1]

    @property
    def Rows(self):
        r1, c1, r2, c2 = self.areas[0]
        rng = Count(r2 - r1 + 1)
        rng.AutoFit = lambda: self.ws.log.append(("RowsAutoFit", self.Address))
        return rng

    @property
    def Columns(self):
        r1, c1, r2, c2 = self.areas[0]
        rng = Count(c2 - c1 + 1)
        rng.AutoFit = lambda: self.ws.log.append(("AutoFit", self.Address))
        return rng

    @property
    def Address(self):
        parts = []
        for r1, c1, r2, c2 in self.areas:
            a = f"${col_let(c1)}${r1}"
            if (r1, c1) != (r2, c2):
                a += f":${col_let(c2)}${r2}"
            parts.append(a)
        return ",".join(parts)

    @property
    def Worksheet(self):
        return self.ws

    @property
    def Left(self):
        return (self.Column - 1) * 48.0

    @property
    def Top(self):
        return (self.Row - 1) * 15.0

    @property
    def Width(self):
        r1, c1, r2, c2 = self.areas[0]
        return (c2 - c1 + 1) * 48.0

    @property
    def Height(self):
        r1, c1, r2, c2 = self.areas[0]
        return (r2 - r1 + 1) * 15.0

    def Cells(self, r, c):
        return FakeRange(self.ws, [(self.Row + r - 1, self.Column + c - 1) * 2])

    def cells(self):
        for r1, c1, r2, c2 in self.areas:
            for r in range(r1, r2 + 1):
                for c in range(c1, c2 + 1):
                    yield r, c

    # --- 값 ---
    @property
    def Value(self):
        r1, c1, r2, c2 = self.areas[0]
        if (r1, c1) == (r2, c2):
            return self.ws.values.get((r1, c1))
        return tuple(tuple(self.ws.values.get((r, c)) for c in range(c1, c2 + 1)) for r in range(r1, r2 + 1))

    def _reject(self, v):
        vals = [x for row in v for x in row] if isinstance(v, tuple) else [v]
        rej = self.ws.reject_formula
        if rej and any(isinstance(x, str) and x.startswith("=") and rej(x) for x in vals):
            raise Exception("com_error: 이 수식에 문제가 있습니다")

    @property
    def Locked(self):
        keys = list(self.cells())
        n = sum(1 for k in keys if k not in self.ws.unlocked)
        return True if n == len(keys) else (False if n == 0 else None)

    @Value.setter
    def Value(self, v):
        self._reject(v)
        r1, c1, r2, c2 = self.areas[0]
        if isinstance(v, tuple):
            for i, row in enumerate(v):
                for j, x in enumerate(row):
                    if getattr(self.ws, "typing", False) and isinstance(x, str) and x.startswith("="):
                        FakeRange(self.ws, [(r1 + i, c1 + j, r1 + i, c1 + j)]).Formula = x     # 실제 Excel: =로 시작하는 글은 수식
                        continue
                    self.ws.set(r1 + i, c1 + j, _typed(self.ws, x))
        else:
            for r, c in self.cells():
                self.ws.set(r, c, _typed(self.ws, v))

    @property
    def Value2(self):
        return self.Value

    @Value2.setter
    def Value2(self, v):
        self.Value = v

    @property
    def PrefixCharacter(self):
        return "'" if (self.Row, self.Column) in self.ws.prefixed else ""

    @property
    def HasFormula(self):
        """실제 Excel처럼: 모두 수식이면 True, 하나도 없으면 False, 섞이면 None."""
        keys = list(self.cells())
        n = sum(1 for k in keys if k in self.ws.formulas)
        self.ws.log.append(("HasFormula", self.Address))
        return True if n == len(keys) else (False if n == 0 else None)

    def _array_of(self, r, c):
        for a in self.ws.arrays:
            if a[0] <= r <= a[2] and a[1] <= c <= a[3]:
                return a
        return None

    @property
    def HasArray(self):
        keys = list(self.cells())
        n = sum(1 for k in keys if self._array_of(*k))
        self.ws.log.append(("HasArray", self.Address))
        return True if n == len(keys) else (False if n == 0 else None)

    @property
    def CurrentArray(self):
        a = self._array_of(self.Row, self.Column)
        if a is None:
            raise Exception("not an array")
        return FakeRange(self.ws, [a])

    @property
    def Formula(self):
        self.ws.log.append(("FormulaRead", self.Address))
        r1, c1, r2, c2 = self.areas[0]

        def f(r, c):
            return self.ws.formulas.get((r, c), self.ws.values.get((r, c)) if self.ws.values.get((r, c)) is not None else "")
        if (r1, c1) == (r2, c2):
            return f(r1, c1)
        return tuple(tuple(f(r, c) for c in range(c1, c2 + 1)) for r in range(r1, r2 + 1))

    @Formula.setter
    def Formula(self, v):
        self._reject(v)
        if isinstance(v, tuple):                  # 되돌리기: 칸마다 원래 수식·값
            r1, c1, _, _ = self.areas[0]
            for i, row in enumerate(v):
                for j, x in enumerate(row):
                    FakeRange(self.ws, [(r1 + i, c1 + j, r1 + i, c1 + j)]).Formula = x
            return
        if not (isinstance(v, str) and v.startswith("=")):
            for r, c in self.cells():
                self.ws.formulas.pop((r, c), None)
                if v in ("", None):
                    self.ws.values.pop((r, c), None)
                    self.ws.prefixed.discard((r, c))
                else:
                    self.ws.set(r, c, _typed(self.ws, v))
            self.ws.log.append(("Formula", self.Address, v))
            return
        # 실제 Excel: 모르는 함수 이름(영어 Formula에 한글 함수 '합계(' 등)는 받아 두고 칸 값이 #NAME? (pywin32는 오류를 정수로 줌)
        bad_name = bool(re.search(r"[\uac00-\ud7a3]+\s*\(", v)) or bool(getattr(self.ws, "error_formula", None) and self.ws.error_formula(v))
        for r, c in self.cells():
            self.ws.formulas[(r, c)] = v
            self.ws.values[(r, c)] = -2146826259 if bad_name else f"<{v}>"
        self.ws.log.append(("Formula", self.Address, v))

    @property
    def Formula2(self):
        """동적 배열 수식 (Office 2021 이상). ws.no_formula2면 없는 속성 (Office 2019)."""
        if getattr(self.ws, "no_formula2", False):
            raise AttributeError("Formula2")
        return self.Formula

    @Formula2.setter
    def Formula2(self, v):
        if getattr(self.ws, "no_formula2", False):
            raise AttributeError("Formula2")
        self.Formula = v
        self.ws.log.append(("Formula2", self.Address, v))

    @property
    def FormulaR1C1(self):
        """상대 참조(A1, $ 없음)를 R[행차]C[열차]로 바꾼 수식 (복사해도 같은 값)."""
        r1, c1, r2, c2 = self.areas[0]

        def f(r, c):
            if (r, c) in self.ws.formulas:
                return to_r1c1(self.ws.formulas[(r, c)], r, c)
            v = self.ws.values.get((r, c))
            return "" if v is None else v
        if (r1, c1) == (r2, c2):
            return f(r1, c1)
        return tuple(tuple(f(r, c) for c in range(c1, c2 + 1)) for r in range(r1, r2 + 1))

    @property
    def Left(self):
        return float((self.areas[0][1] - 1) * 48)

    @property
    def Top(self):
        return float((self.areas[0][0] - 1) * 15)

    def Hyperlinks(self, i):
        links = [h for h in self.ws.links if h[0] == self.Address]
        return types.SimpleNamespace(Address=links[i - 1][1], TextToDisplay=links[i - 1][2])

    def AutoFilter(self, **kw):
        """자동 필터 (range --action filter): 건 조건을 기록하고 시트의 AutoFilterMode를 켬."""
        self.ws.AutoFilterMode = True
        self.ws.filters.append((self.Address, kw))

    # --- 서식/조작 ---
    def __getattr__(self, name):
        if name in ("Interior", "Font", "Borders_"):
            return self.ws.props(self.Address, name)
        if name == "WrapText":                   # 자동 줄바꿈: 쓴 범위 그대로 다시 읽음 (안 썼으면 False)
            return bool(self.ws.props(self.Address, "Misc").get("WrapText", False))
        if name == "HorizontalAlignment":        # 가로 맞춤: 쓴 범위 그대로 (안 썼으면 xlGeneral 1)
            return self.ws.props(self.Address, "Misc").get("HorizontalAlignment", 1)
        raise AttributeError(name)

    def __setattr__(self, k, v):
        if k in ("ws", "areas", "Value", "Value2", "Formula", "Formula2"):
            object.__setattr__(self, k, v)
        else:
            if k in ("WrapText", "HorizontalAlignment"):
                setattr(self.ws.props(self.Address, "Misc"), k, v)
            self.ws.log.append(("set", self.Address, k, v))

    def Borders(self, i):
        return self.ws.props(self.Address, f"Border{i}")

    @property
    def FormatConditions(self):
        return self.ws.fcs.setdefault(self.Address, FakeFormatConditions())

    def End(self, direction):
        """xlUp: 그 열의 마지막 값 칸. 실제 Excel처럼 필터로 숨긴 행(ws.hidden_rows)은 건너뜀."""
        r1, c1, _, _ = self.areas[0]
        hidden = getattr(self.ws, "hidden_rows", set())
        filled = [r for (r, c) in self.ws.values if c == c1 and self.ws.values[(r, c)] is not None and r not in hidden]
        last = max(filled) if filled else 1
        return FakeRange(self.ws, [(last, c1, last, c1)])

    def _copy_to(self, dest, cut=False):
        r1, c1, r2, c2 = self.areas[0]
        dr, dc = dest.areas[0][0], dest.areas[0][1]
        vals = {(r, c): self.ws.values.get((r, c)) for r in range(r1, r2 + 1) for c in range(c1, c2 + 1)}
        forms = {k: self.ws.formulas[k] for k in vals if k in self.ws.formulas}
        if cut:
            for k in vals:
                self.ws.values.pop(k, None)
                self.ws.formulas.pop(k, None)
        for (r, c), v in vals.items():
            if (r, c) in forms:
                # Excel처럼 복사하면 상대 참조가 옮긴 만큼 바뀌고 값도 달라짐 (잘라내기는 수식 그대로)
                f = forms[(r, c)] if cut else shift_formula(forms[(r, c)], dr - r1, dc - c1)
                dest.ws.formulas[(dr + r - r1, dc + c - c1)] = f
                dest.ws.values[(dr + r - r1, dc + c - c1)] = v if cut else f"<{f}>"
            elif v is not None:
                dest.ws.set(dr + r - r1, dc + c - c1, v)

    def Cut(self, Destination=None):
        self.ws.log.append(("Cut", self.Address, Destination.Address))
        self._copy_to(Destination, cut=True)

    def Copy(self, Destination=None):
        self.ws.log.append(("Copy", self.Address, Destination.Address if Destination else None))
        if Destination is not None:
            self._copy_to(Destination)

    def Merge(self):
        self.ws.log.append(("Merge", self.Address, self.ws.app.DisplayAlerts))
        self.ws.merged.add(self.Address)

    def UnMerge(self):
        self.ws.log.append(("UnMerge", self.Address))
        self.ws.merged.discard(self.Address)

    @property
    def MergeCells(self):
        return self.Address in self.ws.merged

    @property
    def EntireRow(self):
        return FakeLines(self.ws, "rows", sorted({r for a in self.areas for r in range(a[0], a[2] + 1)}))

    @property
    def EntireColumn(self):
        return FakeLines(self.ws, "cols", sorted({c for a in self.areas for c in range(a[1], a[3] + 1)}))

    def Sort(self, Key1=None, Order1=1, Key2=None, Order2=1, Key3=None, Order3=1, Header=2, MatchCase=False,
             Orientation=1, **kw):
        """Excel처럼 행 단위 정렬: 숫자 < 글자(대소문자 무시) < 논리값, 빈 칸은 오름·내림 모두 맨 뒤, 같은 값은 순서 유지."""
        keys = [(k, o) for k, o in ((Key1, Order1), (Key2, Order2), (Key3, Order3)) if k is not None]
        self.ws.log.append(("Sort", self.Address, Header, [(k.Column, o) for k, o in keys]))
        if getattr(self.ws, "sort_ignores_header", False):
            Header = 2       # 회사 PC 2026-11-10 L1: Header=xlYes를 줘도 실제 엑셀이 머리글 행까지 정렬한 일을 흉내
        self._sort_rows(keys, Header)

    def _sort_rows(self, keys, Header):
        r1, c1, r2, c2 = self.areas[0]
        start = r1 + 1 if Header == 1 else r1
        rows = [[self.ws.values.get((r, c)) for c in range(c1, c2 + 1)] for r in range(start, r2 + 1)]

        def cls(v):
            if isinstance(v, bool):
                return (2, v)
            if isinstance(v, (int, float)) or type(v).__name__ == "Decimal":    # 통화 칸은 Excel에서 숫자 (pywin32가 Decimal로 줌)
                return (0, float(v))
            return (1, str(v).casefold())
        for key, order in reversed(keys):
            j = key.Column - c1
            blank = [r for r in rows if r[j] is None or r[j] == ""]
            full = sorted((r for r in rows if not (r[j] is None or r[j] == "")), key=lambda r: cls(r[j]), reverse=order == 2)
            rows = full + blank
        for i, row in enumerate(rows):
            for j, v in enumerate(row):
                if v is None:
                    self.ws.values.pop((start + i, c1 + j), None)
                else:
                    self.ws.set(start + i, c1 + j, v)

    def ClearFormats(self):
        self.ws.log.append(("ClearFormats", self.Address))

    def ClearContents(self):
        for k in list(self.cells()):
            self.ws.values.pop(k, None)

    def Clear(self):
        self.ClearContents()

    def Select(self):
        self.ws.log.append(("Select", self.Address))

    # --- 메모 ---
    @property
    def Comment(self):
        """왼쪽 위 칸의 메모, 없으면 None (실제 Excel처럼)."""
        return self.ws.notes.get((self.Row, self.Column))

    def AddComment(self, Text=None):
        key = (self.Row, self.Column)
        if key in self.ws.notes or self.ws.refuse_notes:      # 실제 Excel: 메모가 있는 칸, 보호된 시트는 COM 오류
            raise Exception("AddComment 실패")
        self.ws.notes[key] = FakeComment(self.ws, *key, Text or "")
        self.ws.log.append(("AddComment", key))
        return self.ws.notes[key]


class FakeChartObject:
    def __init__(self, ws, name):
        self.ws = ws
        self.Name = name
        self.Chart = Props()
        self.Chart.HasTitle = False
        self.series = []
        chart = self.Chart

        def set_source(rng):
            ncols = rng.Columns.Count
            self.series = [Props() for _ in range(max(1, ncols - 1))]
            # 데이터를 넣을 때의 종류 (분산형은 이때 이미 분산형이어야 첫 열을 x값으로 씀)
            self.type_at_source = object.__getattribute__(chart, "_d").get("ChartType")
            # 실제 Excel(실기 I5): 종류를 먼저 정하고 SetSourceData를 하면 분산형 세부 종류가 74(선+표식)로 바뀜
            if self.type_at_source in (-4169, 75):
                chart.ChartType = 74
        chart.SetSourceData = set_source
        chart.SeriesCollection = self._series
        chart.ApplyDataLabels = lambda: None

    def _series(self, i=None):
        if i is None:
            obj = Count(len(self.series))
            obj.__class__ = type("S", (Count,), {"__call__": lambda s, k: self.series[k - 1]})
            return obj
        return self.series[i - 1]

    def Delete(self):
        self.ws.charts.remove(self)


class FakeChartObjects:
    def __init__(self, ws):
        self.ws = ws

    @property
    def Count(self):
        return len(self.ws.charts)

    def Item(self, i):
        return self.ws.charts[i - 1]

    def Add(self, left, top, width, height):
        obj = FakeChartObject(self.ws, f"Chart {len(self.ws.charts) + 1}")
        obj.Left, obj.Top, obj.Width, obj.Height = left, top, width, height
        self.ws.charts.append(obj)
        return obj

    def Delete(self):
        self.ws.charts.clear()


class FakeSortFields:
    def __init__(self):
        self.fields = []

    def Clear(self):
        self.fields = []

    def Add(self, Key=None, SortOn=0, Order=1, DataOption=0):
        self.fields.append((Key, Order))

    @property
    def Count(self):
        return len(self.fields)


class FakeSort:
    """Worksheet.Sort: SortFields, SetRange, Header, Apply."""

    def __init__(self, ws):
        self.ws = ws
        self.SortFields = FakeSortFields()
        self.rng = None
        self.Header = 2
        self.MatchCase = False
        self.Orientation = 1

    def SetRange(self, rng):
        self.rng = rng

    def Apply(self):
        keys = list(self.SortFields.fields)
        self.ws.log.append(("SortApply", self.rng.Address, self.Header, [(k.Column, o) for k, o in keys]))
        self.rng._sort_rows(keys, self.Header)


class FakeWorksheet:
    def __init__(self, app, wb, name):
        self.app = app
        self.wb = wb
        self._name = name
        self.hidden = {"rows": set(), "cols": set()}
        self.values = {}
        self.prefixed = set()        # '를 붙여 글자로 쓴 칸
        self.formulas = {}
        self.arrays = []             # 배열 수식 범위 [(r1, c1, r2, c2)] (칸마다 formulas에 같은 수식)
        self.styles = {}
        self.AutoFilterMode = False
        self.filters = []
        self.links = []                # (칸 주소, 링크, 보이는 글): write --link
        self.pictures = []
        ws = self

        class _Shapes:
            @property
            def Count(self):
                return len(ws.pictures)

            def AddPicture(self, Filename=None, **kw):
                ws.pictures.append((Filename, kw))
                return types.SimpleNamespace(Name=f"Picture {len(ws.pictures)}")

        class _Links:
            def Add(self, Anchor=None, Address=None, TextToDisplay=None, **kw):
                ws.links.append((Anchor.Address, Address, TextToDisplay))
                r, c = Anchor.areas[0][:2]
                ws.values[(r, c)] = TextToDisplay
        self.Shapes = _Shapes()
        self.Hyperlinks = _Links()
        self.merged = set()
        self.fcs = {}
        self.log = []
        self.charts = []
        self.activated = 0
        self.Rows = Count(1048576)
        self.Sort = FakeSort(self)
        self.notes = {}              # 칸 메모 {(r, c): FakeComment}
        self.refuse_notes = False    # True면 메모 추가·수정이 COM 오류 (보호된 시트 흉내)
        self.Comments = FakeComments(self)
        # 시트 보호: ProtectContents, 개체 보호, 허용 항목(Protection.Allow*), 잠기지 않은 칸(기본은 모두 잠김, 엑셀과 같음)
        self.ProtectContents = False
        self.ProtectDrawingObjects = False
        self.Protection = types.SimpleNamespace(**{k: False for k in (
            "AllowFormattingCells", "AllowFormattingColumns", "AllowFormattingRows", "AllowInsertingRows",
            "AllowInsertingColumns", "AllowDeletingRows", "AllowDeletingColumns", "AllowSorting")})
        self.unlocked = set()
        self.reject_formula = None   # 함수(수식) -> True면 .Formula/.Value 대입이 COM 오류 (엑셀이 받지 않는 수식 흉내)
        self.typing = False          # True면 대입한 글자를 Excel처럼 해석 (007 -> 7, 1-2 -> 날짜, _typed)

    @property
    def Name(self):
        return self._name

    @Name.setter
    def Name(self, v):
        check_fake_sheet_name(v)          # 실제 Excel처럼 쓸 수 없는 이름은 COM 오류
        self._name = v

    def set(self, r, c, v):
        if isinstance(v, str) and v.startswith("'"):      # Excel: 앞의 '는 "글자로 두기" 표시라 값에 남지 않음
            self.prefixed.add((r, c))
            v = v[1:]
        else:
            self.prefixed.discard((r, c))
        self.values[(r, c)] = v
        self.formulas.pop((r, c), None)

    def props(self, address, name):
        if (address, name) not in self.styles and "," not in address and ":" not in address:
            # 칸 하나를 읽을 때: 그 칸이 든 범위에 준 서식 (실제 Excel처럼 범위 서식이 칸에 보임, 나중에 준 것이 이김)
            r, c = self._parse(address.replace("$", ""))[0][:2]
            for (addr_, nm), pr in reversed(list(self.styles.items())):
                if nm == name and any(not isinstance(v, Props) for v in object.__getattribute__(pr, "_d").values()) and any(r1 <= r <= r2 and c1 <= c <= c2 for r1, c1, r2, c2 in self._parse(addr_.replace("$", ""))):
                    return pr
        return self.styles.setdefault((address, name), Props())

    def _parse(self, ref):
        areas = []
        for part in ref.replace("$", "").split(","):
            whole = re.fullmatch(r"([A-Z]+):([A-Z]+)|(\d+):(\d+)", part.strip().upper())
            if whole:                    # 열 전체(B:C) / 행 전체(2:3)
                if whole.group(1):
                    areas.append((1, col_num(whole.group(1)), 1048576, col_num(whole.group(2))))
                else:
                    areas.append((int(whole.group(3)), 1, int(whole.group(4)), 16384))
                continue
            m = re.fullmatch(r"([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?", part.strip().upper())
            if not m:
                raise Exception(f"bad address {ref}")
            r1, c1 = int(m.group(2)), col_num(m.group(1))
            r2, c2 = (int(m.group(4)), col_num(m.group(3))) if m.group(3) else (r1, c1)
            areas.append((r1, c1, r2, c2))
        return areas

    def Range(self, a, b=None):
        if b is not None:
            return FakeRange(self, [(a.Row, a.Column, b.Row, b.Column)])
        return FakeRange(self, self._parse(a))

    def Cells(self, r, c):
        return FakeRange(self, [(r, c, r, c)])

    @property
    def UsedRange(self):
        keys = [k for k, v in self.values.items() if v is not None]
        if not keys:
            return FakeRange(self, [(1, 1, 1, 1)])
        r1 = min(r for r, _ in keys)
        c1 = min(c for _, c in keys)
        r2 = max(r for r, _ in keys)
        c2 = max(c for _, c in keys)
        return FakeRange(self, [(r1, c1, r2, c2)])

    def ChartObjects(self):
        return FakeChartObjects(self)

    def Activate(self):
        self.activated += 1
        self.wb.active = self

    def Delete(self):
        self.wb.sheets.remove(self)

    def Copy(self, After=None):
        clone = FakeWorksheet(self.app, self.wb, self.Name + " (2)")
        clone.values = dict(self.values)
        self.wb.sheets.append(clone)


class FakeSheetsColl:
    def __init__(self, wb):
        self.wb = wb

    @property
    def Count(self):
        return len(self.wb.sheets)

    def __call__(self, key):
        if isinstance(key, int):
            return self.wb.sheets[key - 1]
        for s in self.wb.sheets:
            if s.Name == key:
                return s
        raise Exception("no sheet")

    def Add(self, After=None):
        ws = FakeWorksheet(self.wb.app, self.wb, f"Sheet{len(self.wb.sheets) + 1}")
        self.wb.sheets.append(ws)
        return ws


class FakeExcelWorkbook:
    def __init__(self, app, name="Book1.xlsx", sheets=("Sheet1",)):
        self.app = app
        self.Name = name
        self.FullName = "C:\\docs\\" + name
        self.Saved = True
        self.sheets = [FakeWorksheet(app, self, n) for n in sheets]
        self.active = self.sheets[0]
        self.Worksheets = FakeSheetsColl(self)

    @property
    def ActiveSheet(self):
        return self.active

    def SaveCopyAs(self, path):
        open(path, "wb").close()

    def Save(self):
        self.app.saves = getattr(self.app, "saves", 0) + 1
        self.Saved = True

    def SaveAs(self, path, fmt=None):
        """실제 Excel처럼: 새 파일을 쓰고 열린 문서가 그 파일이 됨 (DisplayAlerts가 켜져 있고 파일이 있으면 확인 창 = 오류로 흉내)."""
        import os
        if os.path.exists(path) and self.app.DisplayAlerts:
            raise Exception("SaveAs: '바꾸시겠습니까?' 창")
        open(path, "wb").close()
        self.FullName, self.Name, self.Saved = path, os.path.basename(path), True
        self.app.saved_as = (path, fmt)


class FakeXlApp:
    def __init__(self):
        self.DisplayAlerts = True
        self.ScreenUpdating = True
        self.EnableEvents = True
        self.Calculation = -4105
        self.Selection = Props()
        self.Selection.Address = "$A$1"
        self.ActiveWindow = Props()
        self.ActiveWindow.FreezePanes = False

    def ConvertFormula(self, Formula, FromReferenceStyle, ToReferenceStyle=None, ToAbsolute=None, RelativeTo=None):
        """A1(1) <-> R1C1(-4150). 상대 참조($ 없는 A1)만 RelativeTo 칸 기준 R[행차]C[열차]로 (실제 Excel처럼 255자 넘으면 오류)."""
        if len(Formula) > 255:
            raise Exception("ConvertFormula: 255자 초과")
        r, c = RelativeTo.Row, RelativeTo.Column
        if FromReferenceStyle == 1 and ToReferenceStyle == -4150:
            return to_r1c1(Formula, r, c)
        if FromReferenceStyle == -4150 and ToReferenceStyle == 1:
            return re.sub(r"R\[(-?\d+)\]C\[(-?\d+)\]",
                          lambda m: f"{col_let(c + int(m.group(2)))}{r + int(m.group(1))}", Formula)
        return Formula

    def Intersect(self, a, b):
        """두 범위(첫 영역)의 겹치는 범위, 없으면 None (실제 Excel처럼)."""
        ar1, ac1, ar2, ac2 = a.areas[0]
        br1, bc1, br2, bc2 = b.areas[0]
        r1, c1, r2, c2 = max(ar1, br1), max(ac1, bc1), min(ar2, br2), min(ac2, bc2)
        if r1 > r2 or c1 > c2:
            return None
        return FakeRange(a.ws, [(r1, c1, r2, c2)])
