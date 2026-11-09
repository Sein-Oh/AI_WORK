---
name: office-data
description: 엑셀·CSV 표 데이터 작업. 표 두 판(이전 판, 새 판) 비교와 BOM(부품 목록) 개정 비교 - 추가·삭제·수량 변경·위치(Ref Des) 변경·속성 변경, 품번 변경 추정, 결과 xlsx와 새 판 칸 메모. 표 집계도 앞으로 이 스킬에 둔다. BOM 비교, 개정 전후 표 비교, 바뀐 행 찾기 요청에 사용. 엑셀 읽기·쓰기는 office-excel, 문서 두 판 비교(글)는 office-spec.
license: Proprietary
compatibility: opencode
metadata:
  version: "1.4.93"
  platform: windows
---

# office-data 스킬 (Windows)

## 규칙
1. 아래 표의 명령을 그대로 쓰고 `<...>` 값과 `[...]` 안의 필요한 옵션만 바꾼다. 옵션을 지어내지 않는다.
2. PowerShell에서는 명령 앞에 전역 지침의 인코딩 줄을 붙인다.
3. 명령은 **1번만** 실행하고 마지막 1줄 JSON만 읽는다. `next_commands`는 그 `command`를 글자 그대로 실행한다. `confirm_commands`는 사용자 "예" 뒤에만 1번 (사용자가 "묻지 말고"라고 했으면 생략, 단 `"destructive": true` 항목은 늘 묻는다).
4. 파일은 전체 경로로 준다. 이전 판은 `--target`, 새 판은 `--target2` (순서: 이전 -> 새). 엑셀에 열려 있어도 저장된 파일을 읽는다.
5. 경로는 결과 JSON 값을 그대로 쓴다. 도구 코드(.py)는 읽지 않는다. 표를 직접 열어 비교하는 코드를 짜지 않는다. 명령 표에 없는 일은 전역 지침 8장(파이썬 직접)대로 한다.

## 명령
| 할 일 | 명령 |
|---|---|
| 두 판 비교 (BOM 개정) | `python "C:\CLI\office-cli\excel.py" compare --target "<이전 판.xlsx>" --target2 "<새 판.xlsx>"` |
| 결과 파일 경로를 정해 비교 | `python "C:\CLI\office-cli\excel.py" compare --target "<이전 판.xlsx>" --target2 "<새 판.xlsx>" --save "<결과.xlsx>"` |
| 시트를 정해 비교 | `python "C:\CLI\office-cli\excel.py" compare --target "<이전 판.xlsx>" --target2 "<새 판.xlsx>" --sheet "<이전 시트>" --sheet2 "<새 시트>"` |
| 짝지을 열을 정해 비교 | `python "C:\CLI\office-cli\excel.py" compare --target "<이전 판.xlsx>" --target2 "<새 판.xlsx>" --key "<머리글 이름>"` (여럿이면 `--key "상위품번,품번"`) |
| 되돌리기 (메모) | `python "C:\CLI\office-cli\office.py" undo` |

- `.xlsx`, `.xlsm`, `.csv`, `.tsv`를 읽는다. `.xls`는 사용자에게 엑셀에서 `.xlsx`로 저장해 달라고 한다.
- 도구가 정하는 것(사용자가 말하지 않으면 옵션을 붙이지 않는다): 제목 행 아래의 머리글 행, 열 역할(품번·상위품번·레벨·수량·단위·품명·규격·메이커·위치·비고), 시트, 짝지을 키(같은 품번이 여러 행이면 상위품번·레벨·위치를 더함).
- 비교는 공백·전각/반각·코드 대소문자를 무시하고, 숫자는 `1` = `1.0` = `"1"`, 위치는 `R1-R3, R5` = `R1,R2,R3,R5`로 본다. 품번의 앞자리 0은 그대로 다른 값이다.
- `--sheet`, `--sheet2`, `--key`는 사용자가 시트나 열을 말했을 때만. `--overwrite`(있는 결과 파일 바꾸기)는 사용자가 요청했을 때만.

## 결과 처리
`agent_directive` 첫 단어별 행동은 전역 지침 4장대로. 추가로:
- `preview`(구분별 건수, 짝짓기 키, 인식한 열)와 `warnings`를 그대로 알린다. 0건이면 "바뀐 것 없음"이라고 알린다.
- 구분별 목록 `added`, `removed`, `qty_changed`, `ref_changed`, `attr_changed`는 항목의 `key`와 함께 보여 준다.
- `replaced_guess`는 도구의 추정이다. "품번이 바뀐 것으로 보임(확인 필요)"이라고 알리고 확정하지 않는다.
- `truncated`이면 목록은 앞부분만이다. 전체는 `full_result_file`(요약, 추가, 삭제, 변경, 대체 추정 시트)이라고 알린다.
- `[후보] 1) ... 2) ...`(짝지을 키를 도구가 못 정함)이면 후보를 보여 주고 사용자가 고른 번호의 `confirm_commands`를 1번 실행한다.
- 사용자가 새 판에 바뀐 칸 메모를 원하면 `next_commands[0]`(미리보기, 변경 없음)을 실행해 `planned`를 보여 주고, "예" 뒤에만 그 결과의 `confirm_commands[0]`.
- `recovered`(새 이름으로 저장 등)는 그대로 알리고 `output_file`의 실제 경로를 쓴다.
