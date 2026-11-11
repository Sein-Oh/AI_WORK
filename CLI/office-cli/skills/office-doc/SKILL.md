---
name: office-doc
description: 열린 워드·파워포인트 문서 읽기·쓰기. 워드 읽기, 단락 지우기·표 행, 양식 채우기, 검토 메모·변경 추적, 바꾸기, PDF 내보내기. PPT 슬라이드 읽기·쓰기·도형 고치기, 발표자 노트, 슬라이드 추가·복제·이동. 엑셀 차트·표와 문서 내용을 앱 사이로 옮기기(transfer), 되돌리기. 보고서·발표자료 작성에 사용. 엑셀은 office-excel, 메일은 office-mail, 파일 검색·규격 근거는 office-spec.
license: Proprietary
compatibility: opencode
metadata:
  version: "1.6.3"
  platform: windows
---

# office-doc 스킬 (Windows)

## 규칙
1. 아래 표의 명령을 그대로 쓰고 `<...>` 값과 `[...]` 안의 필요한 옵션만 바꾼다. 옵션을 지어내지 않는다.
2. PowerShell에서는 명령 앞에 전역 지침의 인코딩 줄을 붙인다.
3. 명령은 **1번만** 실행하고 마지막 1줄 JSON만 읽는다. `next_commands`는 그 `command`를 글자 그대로 실행한다. `confirm_commands`는 사용자 "예" 뒤에만 1번 (사용자가 "묻지 말고"라고 했으면 생략, 단 `"destructive": true` 항목은 늘 묻는다).
4. `--target`을 빼면 사용자가 보고 있는 문서다. 다른 열린 문서는 `--target "<이름>"`, 닫힌 파일은 `--target "<전체경로>"`.
5. 여러 줄 글, 표, 큰따옴표가 든 값은 인자로 넣지 않는다. 파일 쓰기 도구로 UTF-8 .md(.csv)를 만들고 `--file "<그 경로>"`로 넘긴다. .md의 코드는 ```로 감싼다.
6. 문서 내용을 다른 문서로 옮길 때는 `--save`로 파일을 만들고 결과의 `output_file`을 `--file`로 넘긴다.
7. 경로는 결과 JSON 값을 그대로 쓴다. 도구 코드(.py)는 읽지 않는다. 명령 표에 없는 일은 전역 지침 8장대로 한다 (8장의 "하지 않는 일"(두 파일 비교 등)은 파이썬으로도 하지 않고 직접 하는 곳만 알린다).

## 명령
| 할 일 | 명령 |
|---|---|
| 워드 읽기 | `python "C:\CLI\office-cli\word.py" read [--save]` / 장절 `--section "1.1"`(하위 절 포함) / `--pages 3-5` / `--paragraphs 120-140` |
| 워드 목차(장절 목록) | `python "C:\CLI\office-cli\word.py" outline` |
| 워드 쓰기 | `python "C:\CLI\office-cli\word.py" write --file "<내용.md>"` / 한 줄 `--text "## 제목\n- 항목"` / 그림 `--image "<그림.png>"` (위치 `--anchor <글> --pos after`, 여러 곳이면 `--occurrence N`) |
| 워드 쪽 나누기 / 목차 / 요약 상자 | `python "C:\CLI\office-cli\word.py" write --page-break` / `--toc` / `--text "<요약>" --callout` |
| 워드 양식 채우기 | `python "C:\CLI\office-cli\word.py" fill --target "<양식.docx>" --slots "<자리>=<값>; <자리>=<값>" --dry-run` (자리: `{{업체명}}`이면 `업체명`, 표 칸 `표2!R3C4`. 미리보기의 `[행, 열]` 글 확인 뒤 `confirm_commands[0]`. 값이 있는 칸은 원할 때만 `--overwrite`) |
| 워드 검토 메모 | 목록 `python "C:\CLI\office-cli\word.py" comment` / 넣기 `python "C:\CLI\office-cli\word.py" comment --action add --anchor "<문서의 글>" --text "<메모>"` (여러 개 `--file "<메모.csv>"`, 열: 기준 글, 메모 내용) / 답글 `--action reply --no 2 --text "<답>"` / 지우기 `--action delete --no 2 --dry-run` |
| 워드 단락 지우기 | `python "C:\CLI\office-cli\word.py" delete --paragraphs 12-14 --dry-run` (번호는 `read`의 `no`, 확인 뒤 `confirm_commands[0]`) |
| 워드 표 행 | `python "C:\CLI\office-cli\word.py" table --table 1 --action add-row [--row 3] --values "칸; 칸"` / 지우기 `--action delete-row --row 3 --dry-run` |
| 워드 단락 서식 | `python "C:\CLI\office-cli\word.py" format --paragraphs 3-5 --align center --size 14 --bold` (`--align <left/center/right/justify>`, `--font "맑은 고딕"`, `--italic`, `--style "제목 1"`. 번호는 `read`의 `no`) |
| 워드 변경 추적 | `python "C:\CLI\office-cli\word.py" revisions` |
| 워드 바꾸기 | `python "C:\CLI\office-cli\word.py" replace --find <A> --replace <B> [--bold]` / 형광펜만 `--find <A> --highlight yellow` (대소문자 구분 `--match-case`, 단어 전체만 `--whole-word`, 여러 쌍 `--map "<바꿀말.json>"`(`{"초안":"최종"}`), 글자 `--color red --italic --size 12`) |
| 워드 정보 / PDF | `python "C:\CLI\office-cli\word.py" info [--detail]` / `python "C:\CLI\office-cli\word.py" export --save "<보고서.pdf>" [--page-number]` |
| PPT 읽기 | `python "C:\CLI\office-cli\ppt.py" read [--slides 1-3]` |
| PPT 쓰기 | `python "C:\CLI\office-cli\ppt.py" write --slide new --title "<제목>" --file "<내용.md>"` / `--text "- 항목"` / `--image "<그림.png>"` / `--cards "제목:내용; 제목:내용"` (색 `--theme <navy/slate/teal/emerald>`, 워드 표·요약 상자도 같음) |
| PPT 도형 글 바꾸기 | `python "C:\CLI\office-cli\ppt.py" write --slide 3 --shape "<도형 이름>" --text "<새 글>"` (이름은 `info --detail --slides 3`의 `name`) / 옮기기·크기 `python "C:\CLI\office-cli\ppt.py" shape --slide 3 --shape "<이름>" --left 50 --top 100 [--width 300] [--height 200]`(pt) |
| PPT 발표자 노트 | 읽기 `python "C:\CLI\office-cli\ppt.py" notes [--slides 1-3]` / 쓰기 `python "C:\CLI\office-cli\ppt.py" notes --slide 3 --text "<노트>"` (있던 노트에 덧붙이기 `--append`, 여러 줄 `--file "<노트.md>"`) |
| PPT 슬라이드 | `python "C:\CLI\office-cli\ppt.py" slide --action <add/duplicate/delete/move> --slide N [--dest M] [--layout <title/text/title_only/blank>]` |
| PPT 기타 | `python "C:\CLI\office-cli\ppt.py" replace --find <A> --replace <B>` (`--match-case`, `--map "<바꿀말.json>"`) / `python "C:\CLI\office-cli\ppt.py" export --save "<발표.pdf>"` / 배치 확인 `python "C:\CLI\office-cli\ppt.py" info --detail --slides N` |
| 엑셀 차트·범위를 워드/PPT에 | `python "C:\CLI\office-cli\office.py" transfer --from-app excel --chart "<차트이름>" --to-app word,ppt` (범위는 `--range A1:D10` 대신. 문서를 정하려면 `--from-target "<보내는 문서>" --to-target "<받는 문서>"`) |
| 워드/PPT 내용을 엑셀로 | `python "C:\CLI\office-cli\office.py" transfer --from-app word --what text --to-app excel` (표는 `--what table --table 1`) |
| 되돌리기 / 창 | `python "C:\CLI\office-cli\office.py" undo` (목록 `--list`: 고른 항목의 `command` 실행) / `python "C:\CLI\office-cli\office.py" focus --app word` |

## 결과 처리
`agent_directive` 첫 단어별 행동은 전역 지침 4장대로. 추가로:
- transfer는 받는 앱이 안 열려 있으면 아무것도 바꾸지 않고 `[추천]`에 열린 앱만 넣은 명령을 준다.
- `info --detail`의 `layout_warnings`가 있으면 `warnings`를 알린다.
- 단락·행·슬라이드 삭제, `slide --action clear --all-slides`, `--overwrite`(노트·제목·도형 글·표 칸 덮어쓰기)는 사용자가 요청했을 때만.
- `replace`의 `--find`는 비우지 않는다. 새 파일은 사용자가 준 전체 경로로만 만든다. 같은 이름이 있으면 새 이름으로 저장되니 `output_file`·`recovered`를 알린다.
- 잘못 고쳤으면 `undo`로 되돌린다. 저장한 적 없는 문서도 되돌린다.
- `error_type: not_supported`(도구에 없는 기능)면 다시 실행하거나 파이썬으로 대신하지 않고 message의 방법을 사용자에게 알린다.
