"""가짜 PowerPoint (ppt.py 슬라이드 배치/치환/슬라이드 관리 테스트용)."""


class Obj:
    """임의 속성을 받아 두는 객체 (Font, Fill, Line, Bullet 등)."""

    def __getattr__(self, k):
        v = Obj()
        object.__setattr__(self, k, v)
        return v


class FakeTextRange:
    def __init__(self, text=""):
        self.Text = text
        self.Font = Obj()
        self.paras = {}

    def Paragraphs(self, i=None):
        p = self.paras.setdefault(i, Obj())
        p.Font = getattr(p, "Font")
        return p

    def Replace(self, FindWhat="", ReplaceWhat="", After=0, MatchCase=0, WholeWords=0):
        hay = self.Text if MatchCase else self.Text.lower()
        idx = hay.find(FindWhat if MatchCase else FindWhat.lower(), After)
        if idx < 0:
            return None
        self.Text = self.Text[:idx] + ReplaceWhat + self.Text[idx + len(FindWhat):]
        found = Obj()
        found.Start = idx + 1
        found.Font = Obj()
        return found


class FakeTextFrame:
    def __init__(self, text=""):
        self.TextRange = FakeTextRange(text)

    @property
    def HasText(self):
        return -1 if self.TextRange.Text else 0


class FakeShape:
    def __init__(self, name, left=0.0, top=0.0, width=100.0, height=40.0, text=None, kind=17, ph_type=None):
        self.Name = name
        self.Left, self.Top, self.Width, self.Height = left, top, width, height
        self.Type = kind
        self.HasTable = False
        self.HasTextFrame = text is not None
        if text is not None:
            self.TextFrame = FakeTextFrame(text)
        if ph_type is not None:
            self.PlaceholderFormat = Obj()
            self.PlaceholderFormat.Type = ph_type
        self.Fill = Obj()
        self.Line = Obj()
        self.LockAspectRatio = 0
        self.deleted = False

    def Delete(self):
        self.deleted = True


class FakeCell:
    def __init__(self):
        self.Shape = FakeShape("cell", text="")


class FakeTable:
    def __init__(self, rows, cols):
        self.Rows = Obj()
        self.Rows.Count = rows
        self.Columns = Obj()
        self.Columns.Count = cols
        self.cells = {(r, c): FakeCell() for r in range(1, rows + 1) for c in range(1, cols + 1)}

    def Cell(self, r, c):
        return self.cells[(r, c)]


class FakeShapes:
    def __init__(self, slide):
        self.slide = slide
        self.items = []

    @property
    def Count(self):
        return len([s for s in self.items if not s.deleted])

    def _live(self):
        return [s for s in self.items if not s.deleted]

    def __call__(self, i):
        return self._live()[i - 1]

    @property
    def HasTitle(self):
        return any(getattr(s, "PlaceholderFormat", None) is not None and s.PlaceholderFormat.Type in (1, 3)
                   for s in self._live())

    @property
    def Title(self):
        return [s for s in self._live() if getattr(s, "PlaceholderFormat", None) is not None
                and s.PlaceholderFormat.Type in (1, 3)][0]

    def add(self, shape):
        self.items.append(shape)
        return shape

    def AddTextbox(self, orient, left, top, width, height):
        return self.add(FakeShape(f"TextBox {len(self.items) + 1}", left, top, width, height, text=""))

    def AddTable(self, rows, cols, left, top, width, height):
        sh = FakeShape(f"Table {len(self.items) + 1}", left, top, width, height)
        sh.HasTable = True
        sh.Table = FakeTable(rows, cols)
        return self.add(sh)

    def AddPicture(self, FileName, LinkToFile, SaveWithDocument, Left, Top):
        sh = FakeShape(f"Picture {len(self.items) + 1}", Left, Top, 1600.0, 900.0, kind=13)
        sh.file = FileName
        return self.add(sh)

    def AddShape(self, kind, left, top, width, height):
        return self.add(FakeShape(f"Shape {len(self.items) + 1}", left, top, width, height, text=""))


class FakePlaceholders:
    def __init__(self, items):
        self.items = items

    @property
    def Count(self):
        return len(self.items)

    def __call__(self, i):
        return self.items[i - 1]


class FakeNotesPage:
    """노트 페이지: 1번은 슬라이드 그림, 2번이 노트 본문 개체틀 (ppPlaceholderBody=2)."""
    def __init__(self):
        self.Shapes = Obj()
        self.body = FakeShape("Notes Placeholder 2", text="", kind=14, ph_type=2)
        self.Shapes.Placeholders = FakePlaceholders([FakeShape("Slide Image 1", kind=14, ph_type=1), self.body])


class FakeSlide:
    def __init__(self, prs, layout):
        self.prs = prs
        self.layout = layout
        self.Shapes = FakeShapes(self)
        self.NotesPage = FakeNotesPage()
        if layout in (1, 2, 11):
            self.Shapes.add(FakeShape("Title 1", 36, 20, 888, 60, text="", kind=14, ph_type=1))
        if layout == 2:
            self.Shapes.add(FakeShape("Content 2", 36, 100, 888, 380, text="", kind=14, ph_type=7))

    def Duplicate(self):
        idx = self.prs.slides.index(self)
        clone = FakeSlide(self.prs, self.layout)
        self.prs.slides.insert(idx + 1, clone)

    def MoveTo(self, dest):
        self.prs.slides.remove(self)
        self.prs.slides.insert(dest - 1, self)

    def Delete(self):
        self.prs.slides.remove(self)


class FakeSlides:
    def __init__(self, prs):
        self.prs = prs

    @property
    def Count(self):
        return len(self.prs.slides)

    def __call__(self, i):
        return self.prs.slides[i - 1]

    def Add(self, index, layout):
        s = FakeSlide(self.prs, layout)
        self.prs.slides.insert(index - 1, s)
        return s


class FakePresentation:
    def __init__(self, n=0, name="발표.pptx"):
        self.Name = name
        self.FullName = "C:\\docs\\" + name
        self.Saved = -1
        self.slides = []
        self.Slides = FakeSlides(self)
        self.PageSetup = Obj()
        self.PageSetup.SlideWidth = 960.0
        self.PageSetup.SlideHeight = 540.0
        for _ in range(n):
            self.Slides.Add(len(self.slides) + 1, 11)

    def SaveCopyAs(self, path):
        open(path, "wb").close()


class FakeView:
    """ActiveWindow.View: GotoSlide로 옮긴 슬라이드 번호를 goto에 남김."""
    def __init__(self):
        self.goto = []

    def GotoSlide(self, n):
        self.goto.append(n)


class FakeWindow:
    def __init__(self, prs):
        self.View = FakeView()
        self.Presentation = prs


class FakePptWindowApp:
    """ActiveWindow 하나만 있는 PowerPoint 앱 (화면 이동 확인용)."""
    def __init__(self, prs):
        self.ActiveWindow = FakeWindow(prs)
