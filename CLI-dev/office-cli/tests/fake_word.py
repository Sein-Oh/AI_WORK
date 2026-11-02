"""문서를 문자열(단락 기호 \\r 포함)로 흉내 내는 가짜 Word (word.py 위치 계산/치환 로직 테스트용)."""


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
        self.Font = Props(doc.log, ("font", start, end))
        self.ParagraphFormat = Props(doc.log, ("pf", start, end))

    @property
    def Text(self):
        return self.doc.text[self.Start:self.End]

    @Text.setter
    def Text(self, value):
        self.doc.replace(self.Start, self.End, value)
        self.End = self.Start + len(value)

    @property
    def Find(self):
        return FakeFind(self)

    @property
    def Duplicate(self):
        return FakeRange(self.doc, self.Start, self.End)

    @property
    def Style(self):
        return None

    @Style.setter
    def Style(self, value):
        self.doc.log.append(("style", self.doc.text[self.Start:self.End], value))

    def Information(self, kind):
        if kind == 12:                                   # wdWithInTable
            return any(s <= self.Start < e for s, e in self.doc.table_spans)
        raise AssertionError(kind)

    def Tables(self, i):
        s, e = next((s, e) for s, e in self.doc.table_spans if s <= self.Start < e)
        return type("T", (), {"Range": FakeRange(self.doc, s, e)})()

    def InsertAfter(self, s):
        self.doc.replace(self.End, self.End, s)
        self.End += len(s)

    def InsertBreak(self, kind):
        self.InsertAfter("\x0c")

    def Collapse(self, direction):
        if direction == 0:
            self.Start = self.End
        else:
            self.End = self.Start

    def Delete(self):
        self.doc.replace(self.Start, self.End, "")
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
        return FakeTable(self.doc, FakeRange(self.doc, self.Start, self.End + 1))


class FakeParagraph:
    def __init__(self, rng):
        self.Range = rng


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


class FakeComment:
    def __init__(self, scope, text, author="검토자", done=False):
        self.Scope = scope
        self.Range = Props()
        self.Range.Text = text
        self.Author = author
        self.Date = "2026-10-28 10:00"
        self.Done = done
        self.Ancestor = None


class FakeComments:
    def __init__(self, doc):
        self.doc = doc
        self.items = []

    @property
    def Count(self):
        return len(self.items)

    def __call__(self, i):
        return self.items[i - 1]

    def Add(self, Range=None, Text=""):
        c = FakeComment(FakeRange(self.doc, Range.Start, Range.End), Text)
        self.items.append(c)
        self.items.sort(key=lambda x: x.Scope.Start)      # Word도 문서 순서로 번호를 매김
        self.doc.Saved = False
        return c


class FakeStyles:
    def __call__(self, key):
        return key


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

    def replace(self, s, e, v):
        self.text = self.text[:s] + v + self.text[e:]
        self.Saved = False

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
