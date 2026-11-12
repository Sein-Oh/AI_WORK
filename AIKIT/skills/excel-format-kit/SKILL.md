---
name: excel-format-kit
description: 엑셀 모양과 구조 바꾸기: 서식(색·굵게·숫자 서식·테두리·표 테마), 차트 만들기·지우기, 시트 추가·이름·복사·삭제, 행·열 넣기·지우기, 정렬, 자동 필터, 틀 고정·숨기기·병합·너비, 수식을 값으로. 엑셀 서식·차트·시트 정리 요청에 사용. 값 읽기·쓰기·메모는 excel-kit, 표 집계는 stats-kit.
license: Proprietary
compatibility: opencode
metadata:
  version: "1.6.7"
  platform: windows
---

# excel-format-kit 스킬 (Windows)

## 규칙
1. 아래 표의 명령을 그대로 쓰고 `<...>` 값과 `[...]` 안의 필요한 옵션만 바꾼다. 옵션을 지어내지 않는다.
2. PowerShell에서는 명령 앞에 전역 지침의 인코딩 줄을 붙인다.
3. 명령은 **1번만** 실행하고 마지막 1줄 JSON만 읽는다. `next_commands`는 그 `command`를 글자 그대로 실행한다. `confirm_commands`는 사용자 "예" 뒤에만 1번 (사용자가 "묻지 말고"라고 했으면 생략, 단 `"destructive": true` 항목은 늘 묻는다).
4. `--target`을 빼면 사용자가 보고 있는 통합문서다. 다른 열린 문서는 `--target "<이름>"`, 닫힌 파일은 `--target "<전체경로>"`.
5. 칸 목록이 많으면 인자로 넣지 않는다. 파일 쓰기 도구로 UTF-8 .csv를 만들고 `--file "<그 경로>"`로 넘긴다.
6. 경로는 결과 JSON 값을 그대로 쓴다. 도구 코드(.py)는 읽지 않는다. 명령 표에 없는 일은 전역 지침 8장대로 한다 (8장의 "나. 하지 않는 일"(인쇄, 회의 요청 등)은 파이썬으로도 하지 않고 직접 하는 곳만 알린다).

## 명령
| 할 일 | 명령 |
|---|---|
| 차트 | `python "C:\AIKIT\office-cli\excel.py" chart --range A1:B101 --type line --title "<제목>"` (`--type <column/bar/line/line_markers/pie/scatter/scatter_lines/area/combo>`, 값 표시 `--labels`, 범례 `--legend <bottom/right/top/left/none>`, 자리 `--at F2`, 크기 `--width 400 --height 250`, 목록 `--action list`, 지우기 `--action delete --name <제목>`) |
| 서식 | `python "C:\AIKIT\office-cli\excel.py" style --range A1:E10 --theme navy` (`--no-zebra` / `--bg yellow --bold --italic --color red` / `--highlight ">=100:pink"` / `--number-format "#,##0" --align center --border thin --wrap` / 서식 지우기 `--clear`). 떨어진 칸은 `--range "B2;D5"`, 칸이 많으면 `--file "<칸목록.csv>"`(줄마다 첫 칸이 B5) |
| 범위 작업 | `python "C:\AIKIT\office-cli\excel.py" range --action <freeze/unfreeze/hide/unhide/width/height/autofit/clear/copy/move/merge/unmerge> --range A1:D10 [--dest B2] [--size 20]` (외곽선만 `--action border --box --weight medium`) |
| 행·열 넣기·지우기 | `python "C:\AIKIT\office-cli\excel.py" range --action <insert/delete> --range 5:7` (열은 `--range B:C`. 지울 때 먼저 `--dry-run`: 지워질 값과 `#REF!`가 될 수식 칸) |
| 수식을 값으로 | `python "C:\AIKIT\office-cli\excel.py" range --action values --range A1:D20 --dry-run` (확인 뒤 `confirm_commands[0]`) |
| 정렬 | `python "C:\AIKIT\office-cli\excel.py" range --action sort --range A1:E100 --by "금액" [--desc]` (여러 기준 `--by "부서,금액"`, 일부만 내림차순 `--desc "금액"`) |
| 자동 필터 (보이게 거르기) | `python "C:\AIKIT\office-cli\excel.py" range --action filter --range A1:E100 --where "부서 = 영업; 금액 >= 100"` (같은 열 두 값 `"부서 = 영업 | 부서 = 설계"`, 끄기 `--action unfilter`) |
| 시트 | `python "C:\AIKIT\office-cli\excel.py" sheet --action <add/rename/copy/delete/select> --name <A> [--new-name <B>]` |
| 배치 확인 | `python "C:\AIKIT\office-cli\excel.py" info [--detail]` (`layout_warnings`가 있으면 `warnings`를 알림. 열린 엑셀에서 사용자가 선택한(마지막으로 클릭한) 칸은 `--target` 없이 실행한 결과의 `selection`, 시트는 `active_sheet`) |
| 되돌리기 / 창 | `python "C:\AIKIT\office-cli\office.py" undo` (목록 `--list`: 고른 항목의 `command` 실행) / `python "C:\AIKIT\office-cli\office.py" windows` / `python "C:\AIKIT\office-cli\office.py" focus --app excel` / 파일을 화면에 열기 `python "C:\AIKIT\office-cli\office.py" open --file "<파일 전체 경로>"` |

## 결과 처리
`agent_directive` 첫 단어별 행동은 전역 지침 4장대로. 추가로:
- `needs_user_decision`(없는 시트, 지워질 값, 덮어쓸 칸)은 전역 지침 4장대로 사용자에게 묻고 원할 때만 1번 실행한다.
- `sheet_protected`이면 보호를 풀거나 암호를 묻지 않는다. 사용자가 [편집 사용]을 눌렀거나 풀었다고 하면 같은 명령 1번, `[후보]` 칸을 고르면 그 칸으로.
- 시트·행·열 삭제, `--new-sheet`, `--overwrite`는 사용자가 요청했을 때만. 새 파일 경로는 지어 붙이지 않는다 (사용자가 준 경로만, 안 줬으면 `--save`만: 도구가 대상 파일 옆에 저장). 같은 이름이 있으면 새 이름으로 저장되니 `output_file`·`recovered`를 알린다.
- 잘못 고쳤으면 `undo`로 되돌린다.
- `error_type: not_supported`(도구로 하지 않는 일)면 다시 실행하거나 파이썬으로 대신하지 않고 message의 방법을 사용자에게 알린다.
