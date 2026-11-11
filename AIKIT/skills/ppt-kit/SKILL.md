---
name: ppt-kit
description: 열린 파워포인트 발표자료 읽기·쓰기. 슬라이드 읽기·쓰기(목록·표·그림·카드), 도형 글 바꾸기·옮기기·크기, 발표자 노트, 슬라이드 추가·복제·삭제·이동, 바꾸기, PDF 내보내기. 엑셀 차트·표를 PPT로, 워드 내용을 PPT로(PPT를 워드로) 옮기기, 되돌리기. 발표자료 작성에 사용. 워드는 word-kit, 엑셀은 excel-kit, 메일은 outlook-kit, 파일 검색·규격 근거는 search-kit.
license: Proprietary
compatibility: opencode
metadata:
  version: "1.6.7"
  platform: windows
---

# ppt-kit 스킬 (Windows)

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
| PPT 읽기 | `python "C:\AIKIT\office-cli\ppt.py" read [--slides 1-3]` |
| PPT 쓰기 | `python "C:\AIKIT\office-cli\ppt.py" write --slide new --title "<제목>" --file "<내용.md>"` / `--text "- 항목"` / `--image "<그림.png>"` / `--cards "제목:내용; 제목:내용"` (색 `--theme <navy/slate/teal/emerald>`) |
| PPT 도형 글 바꾸기 | `python "C:\AIKIT\office-cli\ppt.py" write --slide 3 --shape "<도형 이름>" --text "<새 글>"` (이름은 `info --detail --slides 3`의 `name`) / 옮기기·크기 `python "C:\AIKIT\office-cli\ppt.py" shape --slide 3 --shape "<이름>" --left 50 --top 100 [--width 300] [--height 200]`(pt) |
| PPT 발표자 노트 | 읽기 `python "C:\AIKIT\office-cli\ppt.py" notes [--slides 1-3]` / 쓰기 `python "C:\AIKIT\office-cli\ppt.py" notes --slide 3 --text "<노트>"` (있던 노트에 덧붙이기 `--append`, 여러 줄 `--file "<노트.md>"`) |
| PPT 슬라이드 | `python "C:\AIKIT\office-cli\ppt.py" slide --action <add/duplicate/delete/move> --slide N [--dest M] [--layout <title/text/title_only/blank>]` |
| PPT 기타 | `python "C:\AIKIT\office-cli\ppt.py" replace --find <A> --replace <B>` (`--match-case`, `--map "<바꿀말.json>"`) / `python "C:\AIKIT\office-cli\ppt.py" export --save "<발표.pdf>"` / 배치 확인 `python "C:\AIKIT\office-cli\ppt.py" info --detail --slides N` |
| 엑셀 차트·범위를 PPT에 | `python "C:\AIKIT\office-cli\office.py" transfer --from-app excel --chart "<차트이름>" --to-app ppt` (범위는 `--range A1:D10` 대신. 문서를 정하려면 `--from-target "<보내는 문서>" --to-target "<받는 문서>"`) |
| 워드 내용을 PPT로 | `python "C:\AIKIT\office-cli\word.py" read --target "<문서.docx>" --save` 후 결과의 `output_file`로 `python "C:\AIKIT\office-cli\ppt.py" write --target "<발표.pptx>" --slide new --title "<제목>" --file "<output_file>"` (transfer는 워드<->PPT를 하지 않는다) |
| PPT 내용을 워드로 | `python "C:\AIKIT\office-cli\ppt.py" read --target "<발표.pptx>" --save` 후 결과의 `output_file`로 `python "C:\AIKIT\office-cli\word.py" write --target "<문서.docx>" --file "<output_file>"` (transfer는 워드<->PPT를 하지 않는다) |
| PPT 내용을 엑셀로 | `python "C:\AIKIT\office-cli\office.py" transfer --from-app ppt --what text --to-app excel` (표는 `--what table --table 1`) |
| 되돌리기 / 창 | `python "C:\AIKIT\office-cli\office.py" undo` (목록 `--list`: 고른 항목의 `command` 실행) / `python "C:\AIKIT\office-cli\office.py" focus --app ppt` |

## 결과 처리
`agent_directive` 첫 단어별 행동은 전역 지침 4장대로. 추가로:
- transfer는 받는 앱이 안 열려 있으면 아무것도 바꾸지 않고 `[추천]`에 열린 앱만 넣은 명령을 준다.
- `info --detail`의 `layout_warnings`가 있으면 `warnings`를 알린다.
- 슬라이드 삭제, `slide --action clear --all-slides`, `--overwrite`(노트·제목·도형 글 덮어쓰기)는 사용자가 요청했을 때만.
- `replace`의 `--find`는 비우지 않는다. 새 파일은 사용자가 준 전체 경로로만 만든다. 같은 이름이 있으면 새 이름으로 저장되니 `output_file`·`recovered`를 알린다.
- 잘못 고쳤으면 `undo`로 되돌린다. 저장한 적 없는 문서도 되돌린다.
- `error_type: not_supported`(도구에 없는 기능)면 다시 실행하거나 파이썬으로 대신하지 않고 message의 방법을 사용자에게 알린다.
