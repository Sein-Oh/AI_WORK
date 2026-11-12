---
name: excel-kit
description: 엑셀(열린 통합문서, xlsx/csv 파일) 읽기·쓰기. 값·표·수식 읽기와 쓰기, 조건 행 찾기, 행 수, 칠한 칸 찾기, 칸 메모, 그림·링크, 바꾸기, 저장·내보내기, 차트·범위를 워드/PPT로, 되돌리기. 엑셀 표 읽기·값 넣기 요청에 사용. 서식·차트·시트·행열·정렬·필터는 excel-format-kit, 표 집계·중복 찾기는 stats-kit, 워드는 word-kit, PPT는 ppt-kit, 메일은 outlook-kit.
license: Proprietary
compatibility: opencode
metadata:
  version: "1.6.7"
  platform: windows
---

# excel-kit 스킬 (Windows)

## 규칙
1. 아래 표의 명령을 그대로 쓰고 `<...>` 값과 `[...]` 안의 필요한 옵션만 바꾼다. 옵션을 지어내지 않는다.
2. PowerShell에서는 명령 앞에 전역 지침의 인코딩 줄을 붙인다.
3. 명령은 **1번만** 실행하고 마지막 1줄 JSON만 읽는다. `next_commands`는 그 `command`를 글자 그대로 실행한다. `confirm_commands`는 사용자 "예" 뒤에만 1번 (사용자가 "묻지 말고"라고 했으면 생략, 단 `"destructive": true` 항목은 늘 묻는다).
4. `--target`을 빼면 사용자가 보고 있는 통합문서다. 다른 열린 문서는 `--target "<이름>"`, 닫힌 파일은 `--target "<전체경로>"`.
5. 표와 여러 줄 값은 인자로 넣지 않는다. 파일 쓰기 도구로 UTF-8 .csv를 만들고 `--file "<그 경로>"`로 넘긴다. `$`가 든 값(수식 `=$A$1`)은 작은따옴표 `'...'`로 감싼다.
6. 경로는 결과 JSON 값을 그대로 쓴다. 도구 코드(.py)는 읽지 않는다. 명령 표에 없는 일은 전역 지침 8장대로 한다 (8장의 "나. 하지 않는 일"(인쇄, 회의 요청 등)은 파이썬으로도 하지 않고 직접 하는 곳만 알린다).

## 명령
| 할 일 | 명령 |
|---|---|
| 읽기 / 행 수 / 셀 찾기 | `python "C:\AIKIT\office-cli\excel.py" read [--target "<파일.xlsx 또는 .csv>"] [--range A1:D20] [--search 글자] [--save]` (행 수는 결과의 `rows`, 수식은 `--formulas`, 200행 넘게 보려면 `--max-rows 500`, 몇 열만 `--columns 품번,수량`, 칠한 칸·굵은 칸 목록 `--formats`, `--range`에 이름 정의 이름도 됨(목록은 `info`의 `names`)) |
| 조건에 맞는 행 | `python "C:\AIKIT\office-cli\excel.py" read --where "금액 >= 100; 부서 = 영업"` (`;` 그리고, `|` 또는, `>= <= > < = != ~`(포함) `!~`(포함 안 함), 값에 `;` `|`가 있으면 `'a;b'`) |
| 칸 하나에 값 | `python "C:\AIKIT\office-cli\excel.py" write --range A2 --value "반가워"` |
| 여러 칸에 각각 | `python "C:\AIKIT\office-cli\excel.py" write --cells "C6=ft; C7=20; 계산기!C11==A1*2"` (많으면 `--cells "칸값.csv"`(열: 칸, 값)) |
| 표 쓰기 | `python "C:\AIKIT\office-cli\excel.py" write --range A1 --file "<표.csv>"` (`--header "x,y"`, `--append`, `--sheet 시트`, 새 시트 `--sheet 이름 --new-sheet`) |
| 수식 쓰기 | `python "C:\AIKIT\office-cli\excel.py" write --range A2:B101 --formula "=ROW()-1" "=A2*2"` |
| 그림·링크 | 그림 `python "C:\AIKIT\office-cli\excel.py" write --range C3 --image "<그림.png>"` / 링크 `python "C:\AIKIT\office-cli\excel.py" write --range B2 --link "<파일 경로 또는 주소>" [--value "<보이는 글>"]` |
| 칸 메모 | 목록 `python "C:\AIKIT\office-cli\excel.py" comment [--range B2:D20]` / 넣기 `python "C:\AIKIT\office-cli\excel.py" comment --action add --range B5 --text "<의견>"` (여러 개 `--file "<메모.csv>"`(열: 칸, 메모), 있던 메모 바꾸기 `--overwrite`, 덧붙이기 `--append`) / 지우기 `--action delete --range B2:D20` (먼저 `--dry-run`) |
| 바꾸기 | `python "C:\AIKIT\office-cli\excel.py" replace --find <A> --replace <B> [--all-sheets]` (칸 전체가 같을 때만 `--whole`, 여러 쌍은 `--map "<바꿀말.json>"`(`{"초안":"최종"}`), 바꾼 칸 강조 `--bold --color red`. `--find`는 비우지 않는다. 빈 칸 채우기는 `write --value`) |
| 저장 / 내보내기 | `python "C:\AIKIT\office-cli\excel.py" save` (다른 이름 `--save "<전체경로.xlsx>"`) / 사본 `python "C:\AIKIT\office-cli\excel.py" export --save "<결과.csv|.xlsx|.pdf>"` (CSV 인코딩 `--encoding cp949`) |
| 정보 / 배치 확인 | `python "C:\AIKIT\office-cli\excel.py" info [--detail]` (`layout_warnings`가 있으면 `warnings`를 알림. 열린 엑셀에서 사용자가 선택한(마지막으로 클릭한) 칸은 `--target` 없이 실행한 결과의 `selection`, 시트는 `active_sheet`) |
| 차트·범위를 워드/PPT로 | `python "C:\AIKIT\office-cli\office.py" transfer --from-app excel --chart "<차트이름>" --to-app word,ppt` (범위는 `--range A1:D10` 대신) |
| 되돌리기 / 창 | `python "C:\AIKIT\office-cli\office.py" undo` (목록 `--list`: 고른 항목의 `command` 실행) / `python "C:\AIKIT\office-cli\office.py" windows` / `python "C:\AIKIT\office-cli\office.py" focus --app excel` / 파일을 화면에 열기 `python "C:\AIKIT\office-cli\office.py" open --file "<파일 전체 경로>"` |

## 결과 처리
`agent_directive` 첫 단어별 행동은 전역 지침 4장대로. 추가로:
- `--new-sheet`는 사용자가 새 시트를 원한다고 답했을 때만 (없는 시트는 `[후보]`의 번호 명령).
- `needs_user_decision`(없는 시트, 덮어쓸 칸)은 전역 지침 4장대로 사용자에게 묻고 원할 때만 1번 실행한다.
- `sheet_protected`이면 보호를 풀거나 암호를 묻지 않는다. 사용자가 [편집 사용]을 눌렀거나 풀었다고 하면 같은 명령 1번, `[후보]` 칸을 고르면 그 칸으로.
- 잘못 고쳤으면 `undo`로 되돌린다.
- `error_type: not_supported`(도구로 하지 않는 일)면 다시 실행하거나 파이썬으로 대신하지 않고 message의 방법을 사용자에게 알린다.
- `--new-sheet`, `--overwrite`는 사용자가 요청했을 때만. 새 파일 경로는 지어 붙이지 않는다 (사용자가 준 경로만, 안 줬으면 `--save`만: 도구가 대상 파일 옆에 저장). 같은 이름이 있으면 새 이름으로 저장되니 `output_file`·`recovered`를 알린다.
