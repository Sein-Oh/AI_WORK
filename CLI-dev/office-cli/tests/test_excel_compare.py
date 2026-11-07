"""excel.py compare: 두 판의 표(BOM) 비교. 머리글·열 역할·키를 도구가 정함, 정규화, 구분별 결과, 결과 xlsx 확인, 메모 미리보기, [후보]."""
import csv
import os

import openpyxl
import pytest

import _runtime as rt
import excel
from office_helpers import recommended_argv, run_excel, split_command

HEAD = ["Level", "Part No", "Description", "Qty", "Unit", "Ref Des", "Maker", "비고"]


def _xlsx(path, rows, header=HEAD, title="BOM 목록 (rev A)", sheet="BOM", extra_sheets=()):
    wb = openpyxl.Workbook()
    for name in extra_sheets:                    # 앞쪽에 표가 아닌 시트 (도구가 표 시트를 골라야 함)
        ws = wb.create_sheet(name, 0)
        ws.append(["개정 이력"])
        ws.append(["rev A", "최초"])
    ws = wb["Sheet"]
    ws.title = sheet
    if title:
        ws.append([title])
        ws.append([])
    ws.append(list(header))
    for r in rows:
        ws.append(list(r))
    wb.save(path)
    return str(path)


def _csv(path, rows, header=HEAD):
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(header)
        w.writerows(rows)
    return str(path)


MULTI_OLD = [[1, "A-100", "메인 보드", 1, "EA", "", "삼성", ""],
             [2, "R-0010", "저항 10k 0603", 3, "EA", "R1-R3", "야게오", ""],
             [2, "C-0100", "콘덴서 100nF", 2, "EA", "C1,C2", "무라타", ""],
             [2, "U-001", "MCU STM32F103", 1, "EA", "U1", "ST", ""],
             [1, "B-200", "전원 보드", 1, "EA", "", "삼성", ""],
             [2, "R-0010", "저항 10k 0603", 2, "EA", "R10,R11", "야게오", ""],
             [2, "D-01", "다이오드 1N4148", 1, "EA", "D1", "온세미", ""]]
MULTI_NEW = [[1, "A-100", "메인 보드", 1, "EA", "", "삼성", ""],
             [2, "R-0010", "저항 10k 0603", 4, "EA", "R1-R3,R5", "야게오", ""],
             [2, "C-0100", "콘덴서 100nF", 2, "EA", "C2, C1", "무라타", "변경"],
             [2, "U-002", "MCU STM32F103C8", 1, "EA", "U1", "ST", ""],
             [1, "B-200", "전원 보드", 1, "EA", "", "삼성", ""],
             [2, "r-0010", "저항 10k 0603", 2.0, "ea", "R10-R11", "야게오", ""],
             [2, "L-01", "인덕터 10uH", 1, "EA", "L1", "TDK", ""]]


def _compare(capsys, old, new, *extra):
    return run_excel(capsys, "compare", "--target", old, "--target2", new, *extra)


def _sheet_rows(path, title):
    wb = openpyxl.load_workbook(path, read_only=True)
    try:
        return [list(r) for r in wb[title].iter_rows(values_only=True)]
    finally:
        wb.close()


def test_multi_level_bom_uses_level_parent_and_reports_every_category(tmp_path, capsys):
    old = _xlsx(tmp_path / "old.xlsx", MULTI_OLD)
    new = _xlsx(tmp_path / "new.xlsx", MULTI_NEW, title="BOM 목록 (rev B)")
    code, out = _compare(capsys, old, new)
    assert code == 0 and out["status"] == "ok" and out["verified"] is True, out
    assert out["agent_directive"].startswith("VERIFIED:")
    assert out["header_row"] == {"old": 3, "new": 3} and out["rows"] == {"old": 7, "new": 7}
    # R-0010이 두 상위 아래에 있어 품번만으로는 짝지을 수 없음 -> 레벨로 찾은 상위품번을 더함
    assert out["key_columns"] == "상위품번(레벨로 찾음)+Part No" and "레벨" in out["key_reason"]
    roles = {c["role"]: c["old"] for c in out["columns"]}
    assert roles["품번"] == "B (Part No)" and roles["수량"] == "D (Qty)" and roles["위치"] == "F (Ref Des)" and roles["레벨"] == "A (Level)"
    assert out["counts"] == {"추가": 1, "삭제": 1, "수량 변경": 1, "위치 변경": 1, "속성 변경": 1, "대체 추정": 1}, out["counts"]
    assert out["count"] == 5
    q = out["qty_changed"][0]
    assert q["key"] == "A-100 / R-0010" and (q["old"], q["new"], q["diff"]) == (3, 4, "+1") and q["row_new"] == 5
    r = out["ref_changed"][0]
    assert r["added"] == "R5" and r["removed"] == ""
    a = out["attr_changed"][0]
    assert a["key"] == "A-100 / C-0100" and a["columns"] == [{"column": "비고", "old": "", "new": "변경"}]
    # 'r-0010', 2.0, 'ea', 'R10-R11'은 대소문자·숫자·위치 범위 표기만 달라 같은 행 (바뀐 것 없음)
    assert all("B-200 / R-0010" != x["key"] for x in out["qty_changed"] + out["attr_changed"] + out["ref_changed"])
    assert out["added"][0]["part"] == "L-01" and out["removed"][0]["part"] == "D-01"
    g = out["replaced_guess"][0]
    assert (g["old_key"], g["new_key"], g["guess"]) == ("A-100 / U-001", "A-100 / U-002", True) and "같은 위치 U1" in g["reason"]
    assert "추가 1, 삭제 1, 수량 변경 1, 위치 변경 1, 속성 변경 1, 대체 추정 1" in out["preview"] and "인식한 열: 품번 B" in out["preview"]
    assert "추정(확인 필요)" in out["agent_directive"] and out["truncated"] is False
    # 결과 xlsx: 요약, 추가, 삭제, 변경(이전·새 값 나란히), 대체 추정
    path = out["output_file"]
    assert openpyxl.load_workbook(path, read_only=True).sheetnames == ["요약", "추가", "삭제", "변경", "대체 추정"]
    change = _sheet_rows(path, "변경")
    assert change[0] == ["키", "구분", "열", "이전 값", "새 값", "차이", "이전 행", "새 행"]
    assert ["A-100 / R-0010", "수량 변경", "Qty", 3, 4, "+1", 5, 5] in change
    assert _sheet_rows(path, "대체 추정")[1][7] == "추정(확인 필요)"


def test_memo_file_and_comment_preview_are_next_commands(tmp_path, capsys):
    old = _xlsx(tmp_path / "old.xlsx", MULTI_OLD)
    new = _xlsx(tmp_path / "new.xlsx", MULTI_NEW)
    _, out = _compare(capsys, old, new)
    with open(out["memo_file"], encoding="utf-8-sig") as f:
        memo = {r[0]: r[1] for r in csv.reader(f)}
    assert memo["BOM!D5"] == "비교: 수량 3 -> 4 (+1)" and memo["BOM!F5"] == "비교: 위치 추가 R5"
    assert memo["BOM!B10"].startswith("비교: 이전 판에 없는 행") and "추정" in memo["BOM!B7"]
    cmd = out["next_commands"][0]["command"]
    argv = split_command(cmd)[2:]
    assert argv[:2] == ["comment", "--target"] and argv[2] == new and "--dry-run" in argv and "confirm_commands" not in out


def test_single_level_unique_part_numbers_with_title_rows(tmp_path, capsys):
    head = ["No", "품번", "품명", "규격", "소요량", "단위", "제조사", "비고"]
    old = _xlsx(tmp_path / "a.xlsx", [[1, "P-1", "볼트", "M3x10", 4, "EA", "A사", ""],
                                      [2, "P-2", "너트", "M3", "4", "EA", "A사", ""],
                                      [3, "P-3", "와셔", "M3", 8, "EA", "B사", ""]], header=head, title="부품표")
    new = _xlsx(tmp_path / "b.xlsx", [[1, "P-1", "볼트", "M3x12", 4.0, "EA", "A사", ""],
                                      [2, "P-2", "너트", "M3", 6, "EA", "A사", ""],
                                      [3, "P-4", "스프링", "SUS", 1, "EA", "C사", ""]], header=head, title="부품표")
    code, out = _compare(capsys, old, new)
    assert code == 0 and out["key_columns"] == "품번" and out["key_reason"] == "품번이 행마다 다름"
    assert out["counts"]["수량 변경"] == 1 and out["qty_changed"][0]["diff"] == "+2"        # 문자 "4" -> 6
    assert out["attr_changed"][0]["columns"] == [{"column": "규격", "old": "M3x10", "new": "M3x12"}]   # 4 -> 4.0은 같음
    assert out["counts"]["추가"] + out["counts"]["삭제"] + out["counts"]["대체 추정"] >= 1
    assert {c["role"] for c in out["columns"]} >= {"품번", "품명", "규격", "수량", "단위", "메이커", "비고"}


def test_parent_column_composite_key(tmp_path, capsys):
    head = ["상위품번", "품번", "품명", "수량"]
    rows = [["ASSY-1", "S-1", "나사", 2], ["ASSY-2", "S-1", "나사", 3], ["ASSY-1", "K-1", "키", 1]]
    old = _xlsx(tmp_path / "a.xlsx", rows, header=head, title=None)
    rows2 = [["ASSY-1", "S-1", "나사", 2], ["ASSY-2", "S-1", "나사", 5], ["ASSY-1", "K-1", "키", 1]]
    new = _xlsx(tmp_path / "b.xlsx", rows2, header=head, title=None)
    _, out = _compare(capsys, old, new)
    assert out["key_columns"] == "상위품번+품번" and out["count"] == 1
    assert out["qty_changed"][0]["key"] == "ASSY-2 / S-1" and out["qty_changed"][0]["diff"] == "+2"


def test_no_changes_says_zero_and_writes_no_file(tmp_path, capsys):
    old = _xlsx(tmp_path / "old.xlsx", MULTI_OLD)
    same = [list(r) for r in MULTI_OLD]
    same[1][5] = "R3, R1,R2"                     # 위치 순서·띄어쓰기만 다름
    same[2][3] = "2"                             # 숫자 2 = 글자 "2"
    same[0][2] = "  메인  보드 "                  # 앞뒤·여러 칸 공백
    new = _xlsx(tmp_path / "new.xlsx", same)
    code, out = _compare(capsys, old, new)
    assert code == 0 and out["verified"] is True and out["count"] == 0, out
    assert out["agent_directive"].startswith("VERIFIED: 0건. 바뀐 것 없음") and "바뀐 것 없음" in out["preview"]
    assert "output_file" not in out and "next_commands" not in out and all(v == 0 for v in out["counts"].values())


def test_ref_designator_normalization():
    assert excel._cmp_refs("R1-R3, R5") == excel._cmp_refs("R1,R2,R3,R5") == excel._cmp_refs("r5 r1~r3") == excel._cmp_refs("R1-3;R5")
    assert excel._cmp_refs("C01-C03") == frozenset({"C01", "C02", "C03"})
    assert excel._cmp_refs("R1-C3") == frozenset({"R1-C3"})              # 머리글자가 다른 범위는 펼치지 않음
    assert len(excel._cmp_refs("R1-R99999")) == 1                        # 너무 큰 범위(오타)는 펼치지 않음
    assert excel._cmp_ref_text(excel._cmp_refs("R10,R2,R1")) == "R1,R2,R10"


def test_value_and_code_normalization():
    assert excel._cmp_value(1) == excel._cmp_value(1.0) == excel._cmp_value("1") == excel._cmp_value(" 1.0 ") == "1"
    assert excel._cmp_value("1,000") == "1000" and excel._cmp_value("007") == "007"     # 앞자리 0 글은 그대로
    assert excel._cmp_value("ＡＢＣ　１") == "ABC 1"                                      # 전각 -> 반각
    assert excel._cmp_code("0012") == "0012" != excel._cmp_code(12) and excel._cmp_code(12.0) == "12"
    assert excel._cmp_code(" ab-01 ") == excel._cmp_code("AB-01") == excel._cmp_code("ＡＢ－０１")
    assert excel._cmp_header_key("Q'ty") == "qty" and excel._cmp_header_key("수량(EA)") == "수량"
    assert excel._cmp_header_key("P/N") == "pn" and excel._cmp_header_key("Part No.") == "partno"
    assert [excel._cmp_level(v) for v in (1, "2", ".3", "..3", "**", "1.2.3", "x")] == [1, 2, 3, 3, 2, 3, None]


def test_leading_zero_codes_and_csv_input(tmp_path, capsys):
    head = ["자재코드", "자재명", "수량"]
    old = _csv(tmp_path / "old.csv", [["007", "핀", "10"], ["0100", "캡", "1"], ["12", "축", "1"]], header=head)
    new = _xlsx(tmp_path / "new.xlsx", [["007", "핀", 10], ["100", "캡", 1], ["12", "축", 2]], header=head, title=None)
    code, out = _compare(capsys, old, new)
    assert code == 0 and out["key_columns"] == "자재코드", out
    # 앞자리 0은 코드의 일부: 0100과 100은 다른 품번 (삭제 + 추가 또는 대체 추정), 007은 같은 행
    keys = [x["part"] for x in out["added"]] + [g["new_key"] for g in out["replaced_guess"]]
    assert "100" in keys and out["qty_changed"][0]["key"] == "12"
    assert "memo_file" in out                                    # 새 판이 xlsx면 메모 파일
    code, out = _compare(capsys, new, old)                       # 새 판이 CSV면 메모를 달 수 없음
    assert code == 0 and "memo_file" not in out and "next_commands" not in out


def test_ambiguous_key_gives_numbered_candidates_that_run(tmp_path, capsys):
    head = ["품번", "품명", "수량", "위치"]
    rows = [["S-1", "나사", 2, ""], ["S-1", "나사", 3, ""], ["K-1", "키", 1, "K1"]]
    old = _xlsx(tmp_path / "a.xlsx", rows, header=head, title=None)
    new = _xlsx(tmp_path / "b.xlsx", [["S-1", "나사", 2, ""], ["S-1", "나사", 4, ""], ["K-1", "키", 1, "K1"]], header=head, title=None)
    code, out = _compare(capsys, old, new)
    assert code == 1 and out["agent_directive"].startswith("STOP_RETRY:") and out["error_type"] == "ambiguous", out
    assert "[후보] 1) " in out["message"] and "2) " in out["message"]
    cmds = out["confirm_commands"]
    assert len(cmds) == len(out["candidates"]) == 2
    argv = split_command(cmds[0]["command"])
    assert argv[0] == "python" and os.path.basename(argv[1]) == "excel.py" and argv[2] == "compare"
    assert argv[argv.index("--key") + 1] == "품번"
    code, out = run_excel(capsys, *argv[2:])
    assert code == 0 and out["key_columns"] == "품번" and out["qty_changed"][0]["diff"] == "+1", out
    assert any("나온 순서대로 짝지음" in w for w in out["warnings"])


def test_unknown_part_column_asks_for_key_column(tmp_path, capsys):
    head = ["항목", "값", "메모"]
    old = _xlsx(tmp_path / "a.xlsx", [["x", 1, ""], ["y", 2, ""]], header=head, title=None)
    new = _xlsx(tmp_path / "b.xlsx", [["x", 1, ""], ["y", 3, ""]], header=head, title=None)
    code, out = _compare(capsys, old, new)
    assert code == 1 and "품번 열" in out["message"] and out["candidates"][0] == "항목", out
    argv = split_command(out["confirm_commands"][0]["command"])[2:]
    code, out = run_excel(capsys, *argv)
    assert code == 0 and out["count"] == 1 and out["attr_changed"][0]["columns"][0]["column"] == "값"


def test_key_given_as_column_letter_recommends_header_name(tmp_path, capsys):
    old = _xlsx(tmp_path / "old.xlsx", MULTI_OLD)
    new = _xlsx(tmp_path / "new.xlsx", MULTI_NEW)
    code, out = _compare(capsys, old, new, "--key", "B")
    assert code == 1 and "[추천]" in out["message"], out
    argv = recommended_argv(out["message"], "excel")
    assert argv[argv.index("--key") + 1] == "Part No"
    code, out = run_excel(capsys, *argv)
    assert code == 0 and out["key_columns"] == "Part No" and any("나온 순서대로" in w for w in out["warnings"])


def test_save_is_verified_and_never_overwrites_without_option(tmp_path, capsys):
    old = _xlsx(tmp_path / "old.xlsx", MULTI_OLD)
    new = _xlsx(tmp_path / "new.xlsx", MULTI_NEW)
    target = tmp_path / "결과.xlsx"
    code, out = _compare(capsys, old, new, "--save", str(target))
    assert code == 0 and out["output_file"] == str(target) and out["verified"] is True
    code, out = _compare(capsys, old, new, "--save", str(target))
    assert out["output_file"] != str(target) and out["recovered"] and "덮어쓰지 않고" in out["recovered"][0]
    before = os.path.getmtime(target)
    code, out = _compare(capsys, old, new, "--save", str(target), "--overwrite")
    assert code == 0 and out["output_file"] == str(target)
    assert not any("결과.xlsx" in r for r in out.get("recovered", []))          # 메모 파일(같은 초)만 새 이름일 수 있음
    assert os.path.getmtime(target) >= before
    same = _xlsx(tmp_path / "same.xlsx", MULTI_OLD)              # 0건이어도 --save를 주면 요약 파일
    code, out = _compare(capsys, old, same, "--save", str(tmp_path / "zero.xlsx"))
    assert code == 0 and out["count"] == 0 and out["verified"] is True and os.path.isfile(out["output_file"])


def test_unverified_when_result_file_does_not_read_back(tmp_path, capsys, monkeypatch):
    old = _xlsx(tmp_path / "old.xlsx", MULTI_OLD)
    new = _xlsx(tmp_path / "new.xlsx", MULTI_NEW)
    monkeypatch.setattr(excel, "_cmp_check_xlsx", lambda path, expected: "결과 파일 '추가' 시트의 행 수가 다름 (기대 2, 읽은 1)")
    code, out = _compare(capsys, old, new)
    assert code == 0 and out["verified"] is False and out["agent_directive"].startswith("UNVERIFIED:") and "추가" in out["unverified"][0]


def test_truncated_lists_point_to_full_result_file(tmp_path, capsys):
    head = ["품번", "품명", "수량"]
    old = _xlsx(tmp_path / "a.xlsx", [[f"P{i:03d}", f"부품{i}", 1] for i in range(5)], header=head, title=None)
    new = _xlsx(tmp_path / "b.xlsx", [[f"P{i:03d}", f"부품{i}", 1] for i in range(5)] +
                [[f"N{i:03d}", f"새 부품 {i * 7}", 1] for i in range(30)], header=head, title=None)
    _, out = _compare(capsys, old, new)
    assert out["counts"]["추가"] == 30 and len(out["added"]) == excel.COMPARE_SHOWN and out["truncated"] is True
    assert out["full_result_file"] == out["output_file"] and "full_result_file" in out["agent_directive"]
    assert len(_sheet_rows(out["full_result_file"], "추가")) == 31


def test_sheet_is_chosen_by_header_and_new_uses_same_name(tmp_path, capsys):
    old = _xlsx(tmp_path / "old.xlsx", MULTI_OLD, extra_sheets=("표지",))
    new = _xlsx(tmp_path / "new.xlsx", MULTI_NEW, extra_sheets=("표지", "이력"))
    _, out = _compare(capsys, old, new)
    assert out["old_sheet"] == "BOM" and out["new_sheet"] == "BOM"
    code, out = _compare(capsys, old, new, "--sheet", "없는시트")
    assert code == 1 and "시트" in out["message"]


def test_columns_only_in_one_version_are_warned(tmp_path, capsys):
    old = _xlsx(tmp_path / "a.xlsx", [["P1", "볼트", 1]], header=["품번", "품명", "수량"], title=None)
    new = _xlsx(tmp_path / "b.xlsx", [["P1", "볼트", 1, "RoHS"]], header=["품번", "품명", "수량", "환경"], title=None)
    _, out = _compare(capsys, old, new)
    assert out["count"] == 0 and any("새 판에만 있는 열" in w and "환경" in w for w in out["warnings"])
    assert "warnings" in out["agent_directive"]


def test_flat_bom_replacement_guess_by_similar_name(tmp_path, capsys):
    head = ["품번", "품명", "수량"]
    old = _xlsx(tmp_path / "a.xlsx", [["P-1", "육각 볼트 M3x10 SUS", 4], ["P-2", "평와셔 M3", 4]], header=head, title=None)
    new = _xlsx(tmp_path / "b.xlsx", [["P-1", "육각 볼트 M3x10 SUS", 4], ["P-9", "평와셔 M3 ", 4], ["P-5", "스프링", 1]],
                header=head, title=None)
    _, out = _compare(capsys, old, new)
    assert out["counts"]["대체 추정"] == 1 and out["replaced_guess"][0]["old_key"] == "P-2" and out["counts"]["추가"] == 1
    assert out["counts"]["삭제"] == 0


def test_file2_option_gets_full_fixed_command(tmp_path, capsys):
    code, out = run_excel(capsys, "compare", "--target", "a.xlsx", "--file2", "b.xlsx")
    assert code == 2 and recommended_argv(out["message"], "excel") == ["compare", "--target", "a.xlsx", "--target2", "b.xlsx"]
    assert "office-data" in out["agent_directive"]


def test_help_points_to_office_data_skill(capsys):
    code, out = run_excel(capsys, "compare", "-h")
    assert code == 0 and out["skill_file"].endswith(os.path.join("skills", "office-data", "SKILL.md"))
    assert "office-data" in out["agent_directive"]
    code, out = run_excel(capsys, "read", "-h")
    assert out["skill_file"].endswith(os.path.join("skills", "office-excel", "SKILL.md"))


def test_xls_and_same_file_are_input_errors(tmp_path, capsys):
    old = _xlsx(tmp_path / "old.xlsx", MULTI_OLD)
    xls = tmp_path / "old.xls"
    xls.write_bytes(b"\xd0\xcf\x11\xe0")
    code, out = _compare(capsys, str(xls), old)
    assert code == 1 and ".xlsx로 저장" in out["message"] and out["agent_directive"].startswith("STOP_RETRY:")
    code, out = _compare(capsys, old, old)
    assert code == 1 and "같은 파일" in out["message"]


def test_slow_first_file_continues_with_cache(tmp_path, capsys, monkeypatch):
    """R10: 새 판 읽기가 시간 제한을 넘을 것 같으면 읽은 이전 판을 캐시에 두고 CONTINUE, 같은 명령은 캐시를 다시 씀."""
    old = _xlsx(tmp_path / "old.xlsx", MULTI_OLD)
    new = _xlsx(tmp_path / "new.xlsx", MULTI_NEW)
    real = rt.check_time

    def stop(seconds, progress):
        raise rt.Unfinished(progress)
    monkeypatch.setattr(rt, "check_time", stop)
    code, out = _compare(capsys, old, new)
    assert code == 0 and out["continue"] is True and out["agent_directive"].startswith("CONTINUE:"), out
    assert split_command(out["next_commands"][0]["command"])[2:] == ["compare", "--target", old, "--target2", new]
    cache_dir = os.path.join(rt.get_cache_dir(), "compare")
    assert len(os.listdir(cache_dir)) == 1
    monkeypatch.setattr(rt, "check_time", real)
    read = []
    orig = excel._cmp_read_file
    monkeypatch.setattr(excel, "_cmp_read_file", lambda *a: read.append(a[0]) or orig(*a))
    code, out = _compare(capsys, old, new)
    assert code == 0 and out["verified"] is True and out["counts"]["수량 변경"] == 1
    assert read == [new] and os.listdir(cache_dir) == []          # 이전 판은 캐시에서, 끝나면 캐시를 지움


def test_large_bom_is_fast(tmp_path, capsys):
    """BOM 2만 행(여러 단계, 같은 품번 반복)도 몇 초 안에 끝남 (10만 행 실측: 리눅스 약 25초, 대부분 openpyxl 읽기)."""
    import time
    head = ["레벨", "품번", "품명", "수량", "위치"]

    def rows(mut):
        out = []
        for i in range(20000):
            lvl = 1 if i % 50 == 0 else 2
            part = f"ASSY{i // 50:05d}" if lvl == 1 else f"P{i % 997:06d}"
            q = 1 + i % 5 + (1 if mut and i % 1000 == 7 else 0)
            out.append([lvl, part, f"부품 {i % 997}", q, f"R{i}"])
        return out
    old = _xlsx(tmp_path / "a.xlsx", rows(False), header=head, title=None)
    new = _xlsx(tmp_path / "b.xlsx", rows(True), header=head, title=None)
    t = time.monotonic()
    code, out = _compare(capsys, old, new)
    assert code == 0 and out["counts"]["수량 변경"] == 20 and out["count"] == 20, out["counts"]
    assert time.monotonic() - t < 60
