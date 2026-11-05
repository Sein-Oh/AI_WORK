# -*- coding: utf-8 -*-
"""Gemma 종합 테스트(office-cli 묶음 O)용 샘플 파일을 만든다. 사람이나 Gemma가 실행하는 개발 도구 (agent_scenarios.py의 S1).

    python make_office_samples.py [출력 폴더]      # 출력 폴더 생략 시 현재 폴더 아래 office_샘플 폴더

- 규격대조.xlsx  : 항목(B열)·값(C열) 5행 (B3:C7, 머리글 2행). 규격서.txt와 대조하면 일치 2, 불일치 2, 근거없음 1
                   (정격 전압 5.1 일치, 소비 전류 250 불일치, 동작 온도 70 불일치, 절연 저항 150 일치, 포장 색상 근거없음)
- 규격서.txt     : 장절 번호가 있는 규격 문장 (전압 V, 전류 mA, 온도 ℃, 저항 MΩ)
- 단위표.xlsx    : 5행 머리글(항목, 값, 단위), 6~7행 항목·값. 단위 칸 C6, C7은 비어 있음 (O2에서 채움)
- 주간보고.docx  : 제목 1개와 문단 2개 (O4에서 끝에 문단을 덧붙임)
- 성적서양식.docx : 자리 {{업체명}}, {{시험일}}이 있는 양식 (O5에서 채움)
- 시험성적서.docx : 메일 첨부로 쓸 문서 (O3 수동 준비: 이 파일을 첨부한 메일을 자신에게 보냄). 본문에 "판정: 합격"
같은 이름의 파일은 새로 만든 파일로 바꾼다 (시험 폴더 전용).
"""
import os
import sys

import openpyxl
from docx import Document

SPEC = """3. 전기적 특성

3.1 정격 전압은 4.75 V 이상 5.25 V 이하이어야 한다.

3.2 소비 전류는 200 mA 이하이어야 한다.

3.3 동작 온도는 -20 ℃ 이상 60 ℃ 이하이어야 한다.

3.4 절연 저항은 100 MΩ 이상이어야 한다.
"""

CHECK_ROWS = [("정격 전압", 5.1), ("소비 전류", 250), ("동작 온도", 70), ("절연 저항", 150), ("포장 색상", "회색")]


def main(out: str) -> list:
    os.makedirs(out, exist_ok=True)
    made = []

    def keep(name):
        made.append(name)
        return os.path.join(out, name)

    with open(keep("규격서.txt"), "w", encoding="utf-8") as f:
        f.write(SPEC)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "검사"
    ws["B1"] = "부품 검사 기록 (실기 테스트용 샘플)"
    ws.append([None, "항목", "값"])
    for item, value in CHECK_ROWS:
        ws.append([None, item, value])
    wb.save(keep("규격대조.xlsx"))

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "단위"
    ws["A1"] = "측정 항목 단위표 (실기 테스트용 샘플)"
    ws["A5"], ws["B5"], ws["C5"] = "항목", "값", "단위"
    ws["A6"], ws["B6"] = "길이", 10
    ws["A7"], ws["B7"] = "온도", 25
    wb.save(keep("단위표.xlsx"))

    doc = Document()
    doc.add_heading("주간 보고", level=1)
    doc.add_paragraph("1. 이번 주 한 일: 시험 장비 점검")
    doc.add_paragraph("2. 문제점: 없음")
    doc.save(keep("주간보고.docx"))

    doc = Document()
    doc.add_heading("시험 성적서", level=1)
    doc.add_paragraph("업체명: {{업체명}}")
    doc.add_paragraph("시험일: {{시험일}}")
    doc.save(keep("성적서양식.docx"))

    doc = Document()
    doc.add_heading("시험 성적서", level=1)
    doc.add_paragraph("시험 항목: 절연 저항")
    doc.add_paragraph("측정값: 150 MΩ")
    doc.add_paragraph("판정: 합격")
    doc.save(keep("시험성적서.docx"))

    print(f"[INFO] 샘플 {len(made)}개 저장: {out} ({', '.join(made)})", file=sys.stderr)
    return made


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.getcwd(), "office_샘플"))
