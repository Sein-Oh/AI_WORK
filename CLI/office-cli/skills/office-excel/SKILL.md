---
name: office-excel
description: 엑셀(열린 통합문서, xlsx/csv 파일) 읽기·쓰기. 값·표·수식, 조건 행 찾기, 행 수, 차트, 서식, 시트, 행·열 넣기·지우기·정렬, 칸 메모, 내보내기, 차트·범위를 워드/PPT로, 되돌리기. 엑셀 표/그래프 요청에 사용. 규격 대조는 office-spec, 표 두 판 비교(BOM)는 office-data, 워드/PPT는 office-doc, 메일은 office-mail.
license: Proprietary
compatibility: opencode
metadata:
  version: "1.4.83"
  platform: windows
---

# office-excel 스킬 (Windows)

## 규칙
1. 아래 표의 명령을 그대로 쓰고 `<...>` 값과 `[...]` 안의 필요한 옵션만 바꾼다. 옵션을 지어내지 않는다.
2. PowerShell에서는 명령 앞에 전역 지침의 인코딩 줄을 붙인다.
3. 명령은 **1번만** 실행하고 마지막 1줄 JSON만 읽는다. `next_commands`는 그 `command`를 글자 그대로 실행한다. `confirm_commands`는 사용자 "예" 뒤에만 1번 (사용자가 "묻지 말고"라고 했으면 생략, 단 `"destructive": true` 항목은 늘 묻는다).
4. `--target`을 빼면 사용자가 보고 있는 통합문서다. 다른 열린 문서는 `--target "<이름>"`, 닫힌 파일은 `--target "<전체경로>"`.
5. 표와 여러 줄 값은 인자로 넣지 않는다. 파일 쓰기 도구로 UTF-8 .csv를 만들고 `--file "<그 경로>"`로 넘긴다. `$`가 든 값(수식 `=$A$1`)은 작은따옴표 `'...'`로 감싼다.
6. 경로는 결과 JSON 값을 그대로 쓴다. 도구 코드(.py)는 읽지 않는다. 방법이 없으면 사용자에게 알리고 멈춘다.

## 명령
| 할 일 | 명령 |
|---|---|
| 읽기 / 행 수 / 셀 찾기 | `python "C:\CLI\office-cli\excel.py" read [--target "<파일.xlsx 또는 .csv>"] [--range A1:D20] [--search 글자] [--save]` (행 수 `rows`, 수식 보기 `--formulas`) |
| 조건에 맞는 행 | `python "C:\CLI\office-cli\excel.py" read --where "금액 >= 100; 부서 = 영업"` (`;` 그리고, `|` 또는, `>= <= > < = != ~`(포함), 값에 `;` `|`가 있으면 `'a;b'`) |
| 칸 하나에 값 | `python "C:\CLI\office-cli\excel.py" write --range A2 --value "반가워"` |
| 여러 칸에 각각 | `python "C:\CLI\office-cli\excel.py" write --cells "C6=ft; C7=20; 계산기!C11==A1*2"` (많으면 `--cells "칸값.csv"`(열: 칸, 값)) |
| 표 쓰기 | `python "C:\CLI\office-cli\excel.py" write --range A1 --file "<표.csv>"` (`--header "x,y"`, `--append`, `--sheet 시트`, 새 시트 `--sheet 이름 --new-sheet`) |
| 수식 쓰기 | `python "C:\CLI\office-cli\excel.py" write --range A2:B101 --formula "=ROW()-1" "=A2*2"` |
| 차트 | `python "C:\CLI\office-cli\excel.py" chart --range A1:B101 --type line --title "<제목>"` (`--type column|bar|line|pie|scatter|area`, 목록 `--action list`, 지우기 `--action delete --name <제목>`) |
| 서식 | `python "C:\CLI\office-cli\excel.py" style --range A1:E10 --theme navy` (`--no-zebra` / `--bg yellow --bold --italic` / `--highlight ">=100:pink"` / `--number-format "#,##0" --align center --border thin --wrap`) |
| 범위 작업 | `python "C:\CLI\office-cli\excel.py" range --action freeze|unfreeze|hide|unhide|width|height|autofit|clear|copy|move|merge --range A1:D10 [--dest B2] [--size 20]` (외곽선만 `--action border --box --weight medium`) |
| 행·열 넣기·지우기 | `python "C:\CLI\office-cli\excel.py" range --action insert|delete --range 5:7` (열은 `--range B:C`. 지울 때 먼저 `--dry-run`: 지워질 값과 `#REF!`가 될 수식 칸) |
| 정렬 | `python "C:\CLI\office-cli\excel.py" range --action sort --range A1:E100 --by "금액" [--desc]` (여러 기준 `--by "부서,금액"`, 일부만 내림차순 `--desc "금액"`) |
| 칸 메모 | 목록 `python "C:\CLI\office-cli\excel.py" comment [--range B2:D20]` / 넣기 `python "C:\CLI\office-cli\excel.py" comment --action add --range B5 --text "<의견>"` (여러 개 `--file "<메모.csv>"`(열: 칸, 메모), 있던 메모 바꾸기 `--overwrite`, 덧붙이기 `--append`) / 지우기 `--action delete --range B2:D20` (먼저 `--dry-run`) |
| 시트 | `python "C:\CLI\office-cli\excel.py" sheet --action add|rename|copy|delete --name <A> [--new-name <B>]` |
| 바꾸기 | `python "C:\CLI\office-cli\excel.py" replace --find <A> --replace <B> [--all-sheets]` (`--find`는 비우지 않는다. 빈 칸 채우기는 `write --range B2:B10 --value 0`) |
| 내보내기 | `python "C:\CLI\office-cli\excel.py" export --save "<결과.csv|.xlsx|.pdf>" [--range A1:D20]` |
| 정보 / 배치 확인 | `python "C:\CLI\office-cli\excel.py" info [--detail]` (`layout_warnings`가 있으면 `warnings`를 알림) |
| 차트를 워드/PPT로 | `python "C:\CLI\office-cli\office.py" transfer --from-app excel --chart "<차트이름>" --to-app word,ppt` |
| 범위를 워드/PPT 표로 | `python "C:\CLI\office-cli\office.py" transfer --from-app excel --range A1:D10 --to-app word` |
| 되돌리기 / 창 | `python "C:\CLI\office-cli\office.py" undo` (목록 `--list`: 고른 항목의 `command` 실행) / `python "C:\CLI\office-cli\office.py" windows` / `python "C:\CLI\office-cli\office.py" focus --app excel` |

## 결과 처리
`agent_directive` 첫 단어별 행동은 전역 지침 4장대로. 추가로:
- `--new-sheet`는 사용자가 새 시트를 원한다고 답했을 때만 (없는 시트는 `[후보]`의 번호 명령).
- `needs_user_decision`(없는 시트, 지워질 값, 덮어쓸 칸)은 전역 지침 4장대로 사용자에게 묻고 원할 때만 1번 실행한다.
- `sheet_protected`이면 보호를 풀거나 암호를 묻지 않는다. 사용자가 풀었다고 하면 같은 명령 1번, `[후보]` 칸을 고르면 그 칸으로.
- 시트·행·열 삭제, `--new-sheet`, `--overwrite`는 사용자가 요청했을 때만. 새 파일은 사용자가 준 전체 경로로만 만든다. 같은 이름이 있으면 새 이름으로 저장되니 `output_file`·`recovered`를 알린다.
- 잘못 고쳤으면 `undo`로 되돌린다.
