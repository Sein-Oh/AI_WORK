"""1.4.23 (14차 점검) 회귀 테스트.

B1: outlook 검색이 '_'를 공백으로 바꾸고 '%'를 지워 '견적서_v2', '10% 할인'이 절대 맞지 않고 0건 VERIFIED였음.
B2: 차트·시트 이름('1/4분기 매출')이 기본 파일 이름이 되면 경로가 깨졌음 (resolve_output_path에서 정리).
A1: '시트!이름'을 첫 '!'에서만 나눠, 시트 'A!B'의 차트나 다른 시트의 글자 그대로 'Sheet1!x' 차트를 찾지 못하거나 엉뚱한 차트를 골랐음.
A2: 한 시트에 이름이 같은 차트가 둘이면 [후보] 없이 첫 차트를 골랐음.
B3: 손상된 HWPX/HWP가 내부 오류(ERROR:)로 끝났음.
B4: office undo --pick 0이 가장 최근 지점을 되돌렸음.
기타: outlook list/search --limit 상한.
"""
import json
import os
import zipfile

import pytest

import _office as off
import _runtime as rt
import excel
import office
import outlook
import reader
import test_excel
from fake_office import FakeWorkbook

xl = test_excel.xl                  # 가짜 Excel 통합문서 (Sheet1, 요약)


def _run(mod, capsys, *argv):
    with pytest.raises(SystemExit) as ei:
        mod.main(list(argv))
    out = [ln for ln in capsys.readouterr().out.splitlines() if ln.strip()]
    assert len(out) == 1, out
    return ei.value.code, json.loads(out[0])


def _fill(wb):
    for ws in wb.sheets:
        for r, row in enumerate([["월", "값"], ["1월", 3], ["2월", 5]], 1):
            for c, v in enumerate(row, 1):
                ws.set(r, c, v)


def _charts(wb):
    return {str(ws.Name): [(str(o.Name), excel._chart_title(o)) for o in excel._chart_objects(ws)] for ws in wb.sheets}


class _T:
    def __init__(self, wb):
        self.doc = wb

    def sheet(self, name):
        return self.doc.Worksheets(name)


# B1 outlook 검색 -----------------------------------------------------------------------------------------------

def test_dasl_keeps_underscore_and_percent():
    f = outlook.build_search_filter("견적서_v2", None, "subject")
    assert "LIKE '%견적서_v2%'" in f
    f = outlook.build_search_filter("10% 할인", "o'neil", "all")
    assert "LIKE '%10% 할인%'" in f and "o''neil" in f


class _Att:
    Count = 0


class _Item:
    Class = 43

    def __init__(self, i, subject, body="", sender="김팀장"):
        self.EntryID = f"id{i}"
        self.Subject = subject
        self.Body = body
        self.SenderName = sender
        self.SenderEmailAddress = "kim@x.com"
        self.ReceivedTime = "2026-11-01"
        self.UnRead = False
        self.Attachments = _Att()


class _Items(list):
    def Restrict(self, flt):
        self.filter = flt
        return self                    # DASL 와일드카드처럼 넓게 돌려줌 (걸러내기는 도구가)

    def Sort(self, *a):
        pass


class _Folder:
    Name = "받은 편지함"

    def __init__(self, items):
        self.Items = items


def _outlook(monkeypatch, items):
    folder = _Folder(_Items(items))
    monkeypatch.setattr(outlook, "get_outlook", lambda: object())
    monkeypatch.setattr(outlook, "get_namespace", lambda app: object())
    monkeypatch.setattr(outlook, "resolve_folder", lambda ns, name: folder)
    return folder


def test_search_wildcard_query_keeps_only_literal_matches(monkeypatch, capsys):
    folder = _outlook(monkeypatch, [_Item(1, "견적서_v2 송부"), _Item(2, "견적서-v2 송부"), _Item(3, "회의", body="10% 할인 안내")])
    code, out = _run(outlook, capsys, "search", "--query", "견적서_v2")
    assert code == 0 and "견적서_v2" in folder.Items.filter
    assert [m["subject"] for m in out["mails"]] == ["견적서_v2 송부"]      # '_' 와일드카드로 더 찾은 메일은 뺌
    code, out = _run(outlook, capsys, "search", "--query", "10% 할인", "--where", "body")
    assert code == 0 and [m["id"] for m in out["mails"]] == ["id3"]


def test_limit_upper_bound_and_zero(monkeypatch, capsys):
    _outlook(monkeypatch, [_Item(i, f"메일 {i}") for i in range(250)])
    code, out = _run(outlook, capsys, "list", "--limit", "1000")
    assert code == 0 and len(out["mails"]) == outlook.MAX_LIMIT and out["warnings"]
    code, out = _run(outlook, capsys, "search", "--query", "메일", "--limit", "1000")
    assert code == 0 and len(out["mails"]) == outlook.MAX_LIMIT and out["warnings"]
    code, out = _run(outlook, capsys, "list", "--limit", "0")
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY")


# B2 기본 파일 이름 ----------------------------------------------------------------------------------------------

@pytest.mark.parametrize("stem", ["chart_1/4분기 매출_redraw", 'a?b*c:d"e<f>g|h\\i', "CON", "끝점. "])
def test_default_stem_is_sanitized(stem):
    path = rt.resolve_output_path(None, stem, ".png")
    assert os.path.dirname(path) == rt.get_work_dir()
    base = os.path.basename(path)
    assert not any(ch in base for ch in '/\\:*?"<>|') and not base.upper().startswith("CON_")
    with open(path, "wb") as f:                # 실제로 만들 수 있는 경로
        f.write(b"x")


def test_chart_export_with_slash_in_name():
    class _Chart:
        def Export(self, path, fmt):
            with open(path, "wb") as f:
                f.write(office.PNG_SIGNATURE + b"0" * 10)

    class _Obj:
        Chart = _Chart()
    stem = rt.safe_filename("chart_1/4분기 매출", 40)
    path = office.image_by_export(_Obj(), stem)
    assert os.path.dirname(path) == rt.get_work_dir() and "/4" not in os.path.basename(path)
    path = office.image_by_export(_Obj(), "chart_1/4분기 매출")     # 정리하지 않은 이름도 resolve_output_path가 정리
    assert os.path.dirname(path) == rt.get_work_dir()


# A1 '시트!이름' 해석 ---------------------------------------------------------------------------------------------

def test_quoted_sheet_with_bang(xl, capsys):
    app, wb = xl
    _fill(wb)
    wb.sheets[0].Name = "A"
    wb.sheets[1].Name = "A!B"
    for s in ("A", "A!B"):
        _run(excel, capsys, "chart", "--sheet", s, "--range", "A1:B3", "--title", "차트")
    code, out = _run(excel, capsys, "chart", "--action", "delete", "--name", "차트")
    assert code == 1 and "'A!B'!차트" in out["message"] and "'A'!차트" in out["message"], out
    code, out = _run(excel, capsys, "chart", "--action", "delete", "--name", "'A!B'!차트")
    assert code == 0 and _charts(wb) == {"A": [("차트", "차트")], "A!B": []}, out
    ws, obj = excel.find_chart(_T(wb), "'A'!차트")
    assert ws.Name == "A"


def test_unquoted_split_falls_back_to_later_bang(xl, capsys):
    app, wb = xl
    _fill(wb)
    wb.sheets[0].Name = "A"
    wb.sheets[1].Name = "A!B"
    _run(excel, capsys, "chart", "--sheet", "A!B", "--range", "A1:B3", "--title", "차트")
    ws, obj = excel.find_chart(_T(wb), "A!B!차트")         # 'A'에서 '!B!차트'는 없음 -> 'A!B'에서 '차트'
    assert ws.Name == "A!B"


def test_literal_bang_chart_name_on_other_sheet(xl, capsys):
    app, wb = xl
    _fill(wb)
    _run(excel, capsys, "chart", "--sheet", "요약", "--range", "A1:B3", "--title", "t1", "--name", "Sheet1!x")
    ws, obj = excel.find_chart(_T(wb), "Sheet1!x")         # Sheet1에 차트가 없어도 글자 그대로의 이름으로 찾음
    assert ws.Name == "요약" and obj.Name == "Sheet1!x"
    _run(excel, capsys, "chart", "--sheet", "Sheet1", "--range", "A1:B3", "--title", "t2", "--name", "x")
    code, out = _run(excel, capsys, "chart", "--action", "delete", "--name", "Sheet1!x")
    assert code == 1 and "[후보]" in out["message"], out          # 둘 다 맞으면 고르지 않음
    assert _charts(wb) == {"Sheet1": [("x", "t2")], "요약": [("Sheet1!x", "t1")]}
    cands = out["message"].split("[후보]")[1].strip().split(", ")
    assert len(cands) == 2
    for cand in cands:                                       # 후보는 그대로 주면 그 차트 하나로 찾아짐
        excel.find_chart(_T(wb), cand)
    code, out = _run(excel, capsys, "chart", "--action", "delete", "--name", "'Sheet1'!x")
    assert code == 0 and _charts(wb) == {"Sheet1": [], "요약": [("Sheet1!x", "t1")]}, out


def test_sheet_option_contradicts_prefix(xl, capsys):
    app, wb = xl
    _fill(wb)
    _run(excel, capsys, "chart", "--sheet", "Sheet1", "--range", "A1:B3", "--title", "매출")
    code, out = _run(excel, capsys, "chart", "--action", "delete", "--sheet", "요약", "--name", "Sheet1!매출")
    assert code == 1 and "Sheet1" in out["message"] and "요약" in out["message"] and "먼저 만드세요" not in out["message"], out
    code, out = _run(excel, capsys, "chart", "--action", "delete", "--sheet", "요약", "--name", "'Sheet1'!매출")
    assert code == 1 and "--sheet" in out["message"], out
    assert _charts(wb)["Sheet1"] == [("매출", "매출")]


# A2 한 시트의 같은 이름 차트 ---------------------------------------------------------------------------------------

def test_duplicate_names_same_sheet(xl, capsys):
    app, wb = xl
    _fill(wb)
    _run(excel, capsys, "chart", "--sheet", "Sheet1", "--range", "A1:B3", "--title", "매출")
    _run(excel, capsys, "chart", "--sheet", "Sheet1", "--range", "A1:B3", "--title", "비용")
    excel._chart_objects(wb.sheets[0])[1].Name = "매출"
    code, out = _run(excel, capsys, "chart", "--action", "delete", "--name", "매출")
    assert code == 1 and "Sheet1!#1" in out["message"] and "Sheet1!#2" in out["message"], out
    assert len(_charts(wb)["Sheet1"]) == 2
    code, out = _run(excel, capsys, "chart", "--action", "delete", "--name", "Sheet1!#2")
    assert code == 0 and out["agent_directive"].startswith("VERIFIED"), out
    assert _charts(wb)["Sheet1"] == [("매출", "매출")]


# B3 손상된 한글 파일 -----------------------------------------------------------------------------------------------

def test_damaged_hwpx_is_input_error(tmp_path):
    p = tmp_path / "깨짐.hwpx"
    with zipfile.ZipFile(p, "w") as z:
        z.writestr("Contents/section0.xml", "<hs:sec><hp:p>닫히지 않음")
    with pytest.raises(rt.InputError) as ei:
        reader.load_document_status(str(p))
    assert "손상" in str(ei.value)


def test_damaged_hwpx_cli_not_internal(tmp_path, capsys):
    p = tmp_path / "깨짐.hwpx"
    with zipfile.ZipFile(p, "w") as z:
        z.writestr("Contents/section0.xml", "<bad")
    code, out = _run(reader, capsys, "read", "--file", str(p))
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY") and out["error_type"] != "internal", out
    log = rt.error_log_path()
    assert not os.path.exists(log) or "traceback" not in open(log, encoding="utf-8").read()


class _Stream:
    def __init__(self, data):
        self.data = data

    def read(self):
        return self.data


def _fake_ole(monkeypatch, open_error=None, section=b"not zlib"):
    import olefile

    class Ole:
        def __init__(self, path):
            if open_error:
                raise open_error

        def exists(self, name):
            return True

        def openstream(self, name):
            if name == "FileHeader":
                return _Stream(b"\0" * 36 + (1).to_bytes(4, "little"))     # 압축된 본문
            return _Stream(section)

        def listdir(self):
            return [["BodyText", "Section0"]]

        def close(self):
            pass
    monkeypatch.setattr(olefile, "isOleFile", lambda p: True)
    monkeypatch.setattr(olefile, "OleFileIO", Ole)


@pytest.mark.parametrize("open_error", [None, OSError("incorrect OLE sector index")])
def test_damaged_hwp_is_input_error(tmp_path, monkeypatch, open_error):
    p = tmp_path / "깨짐.hwp"
    p.write_bytes(b"x")
    _fake_ole(monkeypatch, open_error)
    with pytest.raises(rt.InputError) as ei:
        reader.load_document_status(str(p))
    assert "손상" in str(ei.value)


# B4 undo --pick 0 ------------------------------------------------------------------------------------------------

def test_undo_pick_zero_rejected(tmp_path, capsys):
    orig = tmp_path / "a.xlsx"
    orig.write_bytes(b"current")
    off.save_snapshot("excel", FakeWorkbook(str(orig)))
    code, out = _run(office, capsys, "undo", "--pick", "0")
    assert code == 1 and "1 이상" in out["message"], out
    assert orig.read_bytes() == b"current"
    code, out = _run(office, capsys, "undo", "--pick", "-1")
    assert code == 1 and orig.read_bytes() == b"current"
