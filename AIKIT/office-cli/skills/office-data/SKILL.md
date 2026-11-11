---
name: office-data
description: 엑셀·CSV 표 데이터 집계와 중복 행 찾기. 표 집계(합계·평균·최소·최대·건수, 묶음별, 조건, 상위 N), 같은 값의 중복 행 찾기. 표 합계·평균·건수, 부서별 합계, 중복 확인 요청에 사용. 두 판 비교(BOM 개정, 문서)는 이 도구에 없음. 엑셀 읽기·쓰기는 office-excel.
license: Proprietary
compatibility: opencode
metadata:
  version: "1.6.4"
  platform: windows
---

# office-data 스킬 (Windows)

## 규칙
1. 아래 표의 명령을 그대로 쓰고 `<...>` 값과 `[...]` 안의 필요한 옵션만 바꾼다. 옵션을 지어내지 않는다.
2. PowerShell에서는 명령 앞에 전역 지침의 인코딩 줄을 붙인다.
3. 명령은 **1번만** 실행하고 마지막 1줄 JSON만 읽는다. `next_commands`는 그 `command`를 글자 그대로 실행한다.
4. 파일은 전체 경로로 준다. `--target`을 빼면 엑셀에 열려 있는 통합문서(보고 있는 시트)를 읽는다.
5. 경로는 결과 JSON 값을 그대로 쓴다. 도구 코드(.py)는 읽지 않는다. 명령 표에 없는 일은 전역 지침 8장대로 한다 (8장의 "하지 않는 일"(두 파일 비교 등)은 파이썬으로도 하지 않고 직접 하는 곳만 알린다).
6. 두 판 비교(엑셀 BOM 개정, 워드·PDF 문서)는 이 도구에 없다. 파이썬으로 대신하지 말고, 워드는 검토 탭 > 비교, 엑셀은 보기 탭 > 나란히 보기로 사용자가 직접 비교하라고 알린다.

## 명령
| 할 일 | 명령 |
|---|---|
| 표 집계 (합계·평균·최소·최대·건수) | `python "C:\AIKIT\office-cli\excel.py" stats --target "<파일.xlsx 또는 .csv>"` (열: `--columns 금액,수량`, 묶음별: `--by 부서`, 조건: `--where "금액 >= 100"`, 합계 상위 N묶음: `--by 부서 --top 5`) |
| 중복 행 찾기 (읽기만) | `python "C:\AIKIT\office-cli\excel.py" duplicates --target "<파일.xlsx 또는 .csv>"` (비교할 열: `--columns 품번,규격`, 생략하면 모든 열이 같은 행) |

- `.xlsx`, `.xlsm`, `.csv`, `.tsv`를 읽는다. `.xls`는 사용자에게 엑셀에서 `.xlsx`로 저장해 달라고 한다.
- 도구가 정하는 것(사용자가 말하지 않으면 옵션을 붙이지 않는다): 제목 행 아래의 머리글 행, 시트, 숫자 열.

## 결과 처리
`agent_directive` 첫 단어별 행동은 전역 지침 4장대로. 추가로:
- `stats`는 `preview`의 숫자를 그대로 알린다. 표를 읽어 직접 더하거나 평균을 내지 않는다 (글자 열은 건수·고유값·많은 값).
- `duplicates`는 `preview`의 묶음과 엑셀 행 번호를 그대로 알린다. 지우기는 사용자가 지울 행을 말한 뒤에만 (office-excel 행 지우기).
- `warnings`가 있으면 함께 알린다. 0건이면 "찾은 항목 없음"이라고 알린다.
- `error_type: not_supported`(도구에 없는 기능)면 다시 실행하거나 파이썬으로 대신하지 않고 message의 방법을 사용자에게 알린다.
