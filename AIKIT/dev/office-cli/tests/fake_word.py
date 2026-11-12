"""문서를 문자열(단락 기호 \\r 포함)로 흉내 내는 가짜 Word (word.py 위치 계산/치환 로직 테스트용)."""
import types



class Props:
    def __init__(self, log=None, tag=None):
        object.__setattr__(self, "_log", log)
        object.__setattr__(self, "_tag", tag)

    def __setattr__(self, k, v):
        if self._log is not None:
            self._log.append((self._tag, k, v))
        object.__setattr__(self, k, v)


class FakeFind:
    def __init__(self, rng):
        self.rng = rng

    def ClearFormatting(self):
        pass

    def Execute(self, FindText="", MatchCase=False, MatchWholeWord=False, MatchWildcards=False, **kw):
        assert MatchWildcards is False
        # 실제 Word처럼 ^를 특수 코드로 읽음: ^^는 글자 ^ 하나, 그 밖의 ^x는 이 가짜에서 지원하지 않음(오류)
        if len(FindText) > 255:                          # 실제 Word: 검색어 255자 초과는 오류 5854
            raise RuntimeError("(-2147352567, 'Microsoft Word', 'string parameter too long', 5854)")
        if "^" in FindText.replace("^^", ""):
            raise RuntimeError("Word Find: 지원하지 않는 ^ 코드")
        FindText = FindText.replace("^^", "^")
        doc = self.rng.doc
        hay = doc.text if MatchCase else doc.text.lower()
        needle = FindText if MatchCase else FindText.lower()
        start = self.rng.Start
        end = self.rng.End if self.rng.End > self.rng.Start else len(doc.text)
        idx = hay.find(needle, start, end)
        if idx < 0:
            return False
        self.rng.Start, self.rng.End = idx, idx + len(FindText)
        return True


class FakeRange:
    def __init__(self, doc, start, end):
        self.doc = doc
        self.Start = start
        self.End = end
        cache = getattr(doc, "format_cache", None)      # 켜면 같은 위치의 Range가 서식을 기억 (실제 Word처럼 다시 읽힘)
        if cache is not None:
            self.Font = cache.setdefault(("font", start, end), Props(doc.log, ("font", start, end)))
            self.ParagraphFormat = cache.setdefault(("pf", start, end), Props(doc.log, ("pf", start, end)))
        else:
            self.Font = Props(doc.log, ("font", start, end))
            self.ParagraphFormat = Props(doc.log, ("pf", start, end))
        live = getattr(doc, "live_ranges", None)
        if live is not None:
            live.add(self)

    @property
    def Text(self):
        return self.doc.text[self.Start:self.End]

    @Text.setter
    def Text(self, value):
        if getattr(self.doc, "TrackRevisions", False) and self.End > self.Start:
            # 변경 추적: 옛 글은 삭제 표시로 남고 새 글이 그 뒤에 들어감 (Find는 삭제 표시 글도 찾음)
            self.doc.deleted_spans.append((self.Start, self.End))
            old_end = self.End
            self.doc.replace(old_end, old_end, value, owner=self)
            self.Start, self.End = old_end, old_end + len(value)
            return
        self.doc.replace(self.Start, self.End, value, owner=self)
        self.End = self.Start + len(value)

    @property
    def Revisions(self):
        hit = [(a, b) for a, b in getattr(self.doc, "deleted_spans", []) if a <= self.Start and self.End <= b]

        class _Revs:
            Count = len(hit)

            def __call__(self, i):
                return type("Rev", (), {"Type": 2})()
        return _Revs()

    @property
    def Find(self):
        return FakeFind(self)

    @property
    def Duplicate(self):
        return FakeRange(self.doc, self.Start, self.End)

    @property
    def Style(self):
        return getattr(self.doc, "styles_at", {}).get(self.Start)

    @Style.setter
    def Style(self, value):
        self.doc.log.append(("style", self.doc.text[self.Start:self.End], value))
        if hasattr(self.doc, "styles_at"):
            self.doc.styles_at[self.Start] = value

    def Information(self, kind):
        if kind == 12:                                   # wdWithInTable
            return any(s <= self.Start < e for s, e in self.doc.table_spans)
        raise AssertionError(kind)

    def Tables(self, i):
        s, e = next((s, e) for s, e in self.doc.table_spans if s <= self.Start < e)
        return type("T", (), {"Range": FakeRange(self.doc, s, e)})()

    def InsertAfter(self, s):
        self.doc.replace(self.End, self.End, s, owner=self)
        self.End += len(s)

    def InsertBreak(self, kind):
        self.InsertAfter("\x0c")

    def Collapse(self, direction):
        if direction == 0:
            self.Start = self.End
        else:
            self.End = self.Start

    def Delete(self):
        self.doc.replace(self.Start, self.End, "", owner=self)
        self.End = self.Start

    def Select(self):
        self.doc.log.append(("select", self.Start))

    def Paragraphs(self, i):
        t = self.doc.text
        s = t.rfind("\r", 0, self.Start) + 1
        e = t.find("\r", self.Start)
        e = len(t) if e < 0 else e + 1
        return FakeParagraph(FakeRange(self.doc, s, e))

    def ConvertToTable(self, Separator=1, NumRows=1, NumColumns=1):
        rows = self.Text.split("\r")
        self.doc.tables.append(rows)
        self.doc.table_spans.append((self.Start, self.End + 1))
        return FakeTable(self.doc, FakeRange(self.doc, self.Start, self.End + 1))

    @property
    def ListFormat(self):
        doc, start = self.doc, self.Start

        class _LF:
            ListTemplate = "번호 목록"

            def ApplyListTemplate(self, ListTemplate=None, ContinuePreviousList=True, ApplyTo=0, **kw):
                doc.log.append(("list_restart", start, ContinuePreviousList, ApplyTo))

            @property
            def ListString(self):
                # 실제 Word: 목록 번호 스타일은 문서의 앞 목록에 이어짐. ApplyTo=0(목록 전체)은 앞 목록까지 한 목록이라
                # 다시 1부터가 되지 않음 (회사 PC 2026-11-10 L5). 어떤 ApplyTo가 먹히는지는 doc.list_restart_works
                works = getattr(doc, "list_restart_works", {1, 2})
                ok = any(e[0] == "list_restart" and e[1] == start and e[2] is False and e[3] in works for e in doc.log)
                return "1." if ok else "3."
        return _LF()


class FakeParagraph:
    def __init__(self, rng):
        self.Range = rng

    def Borders(self, which):
        return Props(self.Range.doc.log, ("border", self.Range.Start, which))


class FakeTable:
    def __init__(self, doc, rng):
        self.doc = doc
        self.Range = rng
        self.Borders = Props()
        self.Borders.Enable = True

    def AutoFitBehavior(self, x):
        pass

    def Rows(self, i):
        row = Props(self.doc.log, ("row", i))
        object.__setattr__(row, "Range", FakeRange(self.doc, self.Range.Start, self.Range.Start))
        object.__setattr__(row, "Shading", Props(self.doc.log, ("rowshade", i)))
        return row

    def Cell(self, r, c):
        cell = Props()
        cell.Shading = Props(self.doc.log, ("cellshade",))
        return cell


class FakeCellRange:
    def __init__(self, grid, r, c):
        self.grid, self.r, self.c = grid, r, c

    @property
    def Text(self):
        return self.grid[self.r][self.c] + "\r\x07"            # Word: 칸 끝 표시

    @Text.setter
    def Text(self, value):
        self.grid[self.r][self.c] = value


class FakeGridTable:
    """칸 글만 있는 표 (word.py fill의 doc.Tables(i).Cell(r, c).Range.Text). None 칸은 병합으로 없는 칸."""

    def __init__(self, rows, merged_rows=False):
        self.grid = [list(r) for r in rows]
        self.merged_rows = merged_rows                       # 세로 병합: Rows(i) 접근이 COM 오류
        self.Rows = _GridRows(self)

    def Cell(self, r, c):
        if not (1 <= r <= len(self.grid) and 1 <= c <= len(self.grid[r - 1])) or self.grid[r - 1][c - 1] is None:
            raise RuntimeError("요청한 구성원이 컬렉션에 없습니다")
        return type("Cell", (), {"Range": FakeCellRange(self.grid, r - 1, c - 1)})()


class _GridRow:
    def __init__(self, table, i):
        self.table, self.i = table, i

    @property
    def Cells(self):
        return type("Cells", (), {"Count": len(self.table.grid[self.i - 1])})()

    @property
    def Range(self):
        row = self.table.grid[self.i - 1]
        return type("R", (), {"Text": "".join((c or "") + "\r\x07" for c in row) + "\r\x07"})()

    def Delete(self):
        if self.table.merged_rows:
            raise RuntimeError("세로로 병합된 칸이 있어 이 컬렉션의 개별 행에 액세스할 수 없습니다")
        del self.table.grid[self.i - 1]


class _GridRows:
    def __init__(self, table):
        self.table = table

    @property
    def Count(self):
        return len(self.table.grid)

    def __call__(self, i):
        if self.table.merged_rows:
            raise RuntimeError("세로로 병합된 칸이 있어 이 컬렉션의 개별 행에 액세스할 수 없습니다")
        if not 1 <= i <= len(self.table.grid):
            raise RuntimeError("요청한 구성원이 컬렉션에 없습니다")
        return _GridRow(self.table, i)

    def Add(self, BeforeRow=None):
        """실제 Word처럼: BeforeRow 앞(없으면 끝)에 빈 행, 칸 수는 이웃 행과 같음."""
        width = len(self.table.grid[-1]) if self.table.grid else 1
        at = len(self.table.grid) if BeforeRow is None else BeforeRow.i - 1
        self.table.grid.insert(at, [""] * width)
        return _GridRow(self.table, at + 1)


class FakeTables:
    def __init__(self, doc):
        self.doc = doc

    @property
    def Count(self):
        return len(self.doc.grid_tables)

    def __call__(self, i):
        return self.doc.grid_tables[i - 1]


class FakeComment:
    def __init__(self, scope, text, author="검토자", done=False):
        self.Scope = scope
        self.Range = Props()
        self.Range.Text = text
        self.Author = author
        self.Date = "2026-10-28 10:00"
        self.Done = done
        self.Ancestor = None
        self.coll = None                                     # FakeComments (번호·답글·지우기)

    @property
    def Index(self):
        return self.coll.items.index(self) + 1

    @property
    def Replies(self):
        parent = self

        class _Replies:
            def Add(self, Range=None, Text=""):
                if getattr(parent.coll.doc, "no_replies", False):        # Word 2010 이하: 답글 없음
                    raise AttributeError("Replies")
                r = FakeComment(parent.Scope, Text, author="나")
                r.Ancestor, r.coll = parent, parent.coll
                at = parent.coll.items.index(parent) + 1
                while at < len(parent.coll.items) and parent.coll.items[at].Ancestor is parent:
                    at += 1
                parent.coll.items.insert(at, r)
                return r
        return _Replies()

    def Delete(self):
        """실제 Word처럼: 그 메모 하나만 지움, 답글은 남음 (회사 PC 2026-11-11 N10: 메모 3 -> 2). 답글까지는 DeleteRecursively."""
        self.coll.items = [c for c in self.coll.items if c is not self]

    def DeleteRecursively(self):
        self.coll.items = [c for c in self.coll.items if c is not self and c.Ancestor is not self]


class FakeComments:
    def __init__(self, doc):
        self.doc = doc
        self.items = []

    @property
    def Count(self):
        return len(self.items)

    def __call__(self, i):
        c = self.items[i - 1]
        c.coll = self
        return c

    def Add(self, Range=None, Text=""):
        c = FakeComment(FakeRange(self.doc, Range.Start, Range.End), Text)
        c.coll = self
        self.items.append(c)
        self.items.sort(key=lambda x: x.Scope.Start)      # Word도 문서 순서로 번호를 매김
        self.doc.Saved = False
        return c


class FakeStyles:
    def __call__(self, key):
        return key


class _HFRange:
    def __init__(self, text=""):
        self.Text = text


class FakeHeaderFooter:
    """머리글·바닥글 하나 (Range.Text, Exists, LinkToPrevious)."""
    def __init__(self, text="", exists=True):
        self.Range = _HFRange(text)
        self.Exists = exists
        self.LinkToPrevious = False


class FakeSection:
    def __init__(self):
        self.headers = {1: FakeHeaderFooter(), 2: FakeHeaderFooter(exists=False), 3: FakeHeaderFooter(exists=False)}
        self.footers = {1: FakeHeaderFooter(), 2: FakeHeaderFooter(exists=False), 3: FakeHeaderFooter(exists=False)}

    def Headers(self, i):
        return self.headers[i]

    def Footers(self, i):
        return self.footers[i]


class FakeSections:
    def __init__(self, n=1):
        self.items = [FakeSection() for _ in range(n)]

    @property
    def Count(self):
        return len(self.items)

    def __call__(self, i):
        return self.items[i - 1]


class FakeHyperlinks:
    def __init__(self):
        self.items = []

    @property
    def Count(self):
        return len(self.items)

    def __call__(self, i):
        return self.items[i - 1]

    def Add(self, Anchor=None, Address="", SubAddress="", **kw):
        h = types.SimpleNamespace(Address=Address, SubAddress=SubAddress, TextToDisplay=Anchor.Text)
        self.items.append(h)
        return h


class FakeBookmarks:
    def __init__(self):
        self.items = {}

    @property
    def Count(self):
        return len(self.items)

    def __call__(self, i):
        name = list(self.items)[i - 1]
        return types.SimpleNamespace(Name=name, Range=self.items[name])

    def Exists(self, name):
        return name in self.items

    def Add(self, name, rng):
        self.items[name] = rng


class FakeWordDocument:
    def __init__(self, text="", name="보고서.docx"):
        self.text = text + "\r"
        self.log = []
        self.tables = []
        self.Name = name
        self.FullName = "C:\\docs\\" + name
        self.Saved = True
        self.Styles = FakeStyles()
        self.WordOpenXML = "<xml/>"
        self.Comments = FakeComments(self)
        self.table_spans = []        # 표 영역 [(시작, 끝)] (끝 = 표 다음 단락의 시작)
        self.grid_tables = []        # fill용 표 (FakeGridTable): doc.Tables(i).Cell(r, c)
        self.Tables = FakeTables(self)
        self.Sections = FakeSections()          # 구역과 머리글·바닥글 (word.py header)
        self.Hyperlinks = FakeHyperlinks()      # word.py link
        self.Bookmarks = FakeBookmarks()        # word.py bookmark
        # 실제 Word처럼 만들어 둔 Range가 문서 편집을 따라 움직임: 앞에 넣으면 밀리고, 안에 넣으면 늘어남
        # (끝 위치에 넣은 것은 그 Range에 들어가지 않음). 넣은 Range 자신(owner)은 각자 메서드가 맞춤
        import weakref
        self.live_ranges = weakref.WeakSet()
        self.deleted_spans = []      # 변경 추적의 삭제 표시 글 [(시작, 끝)] (TrackRevisions = True일 때 Text 바꾸기가 만듦)

    @property
    def Revisions(self):
        return type("Revs", (), {"Count": len(self.deleted_spans)})()

    def replace(self, s, e, v, owner=None):
        self.text = self.text[:s] + v + self.text[e:]
        self.Saved = False
        d = len(v) - (e - s)
        for r in list(getattr(self, "live_ranges", ())):
            if r is owner:
                continue
            if r.Start >= e and not (s == e and r.Start == s and r.End == s):
                r.Start, r.End = r.Start + d, r.End + d
            elif r.End > s and r.Start <= s:
                r.End = max(s, r.End + d) if r.End >= e else s + len(v)
        # 표 영역도 실제 Word처럼 옮김: 앞에 넣으면 밀리고, 안(표 첫 칸의 시작 포함)에 넣으면 늘어남
        self.table_spans = [(a + d, b + d) if s < a else ((a, b + d) if s < b else (a, b)) for a, b in self.table_spans]

    @property
    def TablesOfContents(self):
        doc = self

        class _Toc:
            def __init__(self, rng):
                self.Range = rng

            def Update(self):
                doc.log.append(("toc_update", doc.text))

        class _Tocs:
            items = doc.__dict__.setdefault("_tocs", [])

            @property
            def Count(self):
                return len(self.items)

            def __call__(self, i):
                return self.items[i - 1]

            def Add(self, Range=None, **kw):
                Range.InsertAfter("[목차]")
                t = _Toc(FakeRange(doc, Range.Start, Range.End))
                self.items.append(t)
                return t
        return _Tocs()

    def Range(self, s=0, e=None):
        return FakeRange(self, s, len(self.text) if e is None else e)

    @property
    def Content(self):
        return FakeRange(self, 0, len(self.text))

    @property
    def StoryRanges(self):
        story = self.Content
        story.StoryType = 1
        story.NextStoryRange = None
        return [story]

    def Activate(self):
        self.log.append(("activate",))

    @property
    def paragraphs(self):
        return [p for p in self.text.split("\r")]

    @property
    def Paragraphs(self):
        """문서 단락 (Range는 단락 기호 포함). OutlineLevel은 heading_at(시작 위치 -> 수준)에 있으면 그 수준, 아니면 본문(10)."""
        out, pos = [], 0
        levels = getattr(self, "heading_at", {})
        for part in self.text.split("\r")[:-1]:
            para = FakeParagraph(FakeRange(self, pos, pos + len(part) + 1))
            para.OutlineLevel = levels.get(part, 10)
            out.append(para)
            pos += len(part) + 1
        return out
