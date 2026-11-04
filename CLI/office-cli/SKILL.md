---
name: office-cli
description: 열려 있는 Excel/Word/PowerPoint 문서를 읽고 쓰며(값/표/수식/차트/서식, 본문/표/그림, 슬라이드, 워드 목차·장절·쪽 읽기, 검토 메모·변경 추적, PPT 발표자 노트), 앱 사이로 차트·표·내용을 옮기고(office.py transfer), Outlook 메일을 조회·작성하고(첨부 내려받기, 답장·전달에 문서 첨부), PDF/HWP/Office 파일을 Office 없이 읽고 등록(캐시)해 쪽·장절·번호 근거와 함께 검색한다(reader.py). 작업 전 자동 스냅샷으로 되돌리기(office.py undo)를 지원한다. 엑셀 표/그래프 만들기, 보고서·발표자료 작성, 문서 내용 검색 후 다른 문서에 넣기, 규격/규정 근거 찾기, 메일 답장 작성·첨부 저장 요청 시 사용.
license: Proprietary
compatibility: opencode
metadata:
  version: "1.4.36"
  platform: windows
---

# office-cli 스킬 (Windows)

## 꼭 지킬 규칙
1. 도구 위치는 `C:\CLI\office-cli`이다. 모든 명령은 `python "C:\CLI\office-cli\<파일>.py" <명령> ...` 형식이다 (예: `python "C:\CLI\office-cli\excel.py" read`).
2. PowerShell에서는 명령 앞에 전역 지침의 인코딩 줄(`$env:PYTHONIOENCODING='utf-8'; ...`)을 붙인다.
3. 명령은 **1번만** 실행한다. 출력의 **마지막 1줄 JSON**만 읽는다. `[PROGRESS]` 줄은 무시한다.
4. 아래 표의 명령과 옵션만 쓴다. 옵션을 지어내지 않는다.
5. 대상 문서: `--target`을 생략하면 사용자가 보고 있는 문서(active)다. 다른 열린 문서는 `--target "<이름>"`, 닫힌 파일은 `--target "<전체경로>"`.
6. 표, 여러 줄 글, 큰따옴표가 든 값은 인자로 넣지 않는다. 파일 쓰기 도구로 UTF-8 파일(.csv, .md)을 만든 뒤 `--file "<그 경로>"`로 넘긴다. .md의 코드는 ```로 감싸면 줄마다 그대로 들어간다.
7. 값에 `$`가 있으면(엑셀 수식 `=$A$1` 등) 작은따옴표 `'...'`로 감싼다.
8. 파일 경로는 결과 JSON의 값을 바꾸지 않고 그대로 1번만 쓴다 (`file:///` 링크로 바꾸지 않는다). 이모지를 쓰지 않는다.
9. 도구 코드(.py)는 사용자가 코드를 열어 보라고 직접 요청할 때만 읽는다. 그 밖에는 읽지 않고, 이 문서의 템플릿으로 방법을 찾지 못하면, 하려던 일과 막힌 곳을 사용자에게 알리고 멈춘다.

## 1. 명령

| 할 일 | 명령 |
|---|---|
| 엑셀·CSV 읽기 / 행 수 / 셀 찾기 | `excel.py read [--target "<파일.xlsx 또는 .csv>"] [--range A1:D20] [--search 글자] [--save]` (결과 `rows`가 행 수). 수식 보기는 `--formulas` (결과 `formulas`) |
| 엑셀에 표 쓰기 | `excel.py write --range A1 --file "표.csv"` (`--header "x,y"`, `--append`, `--sheet 시트`, 새 시트는 `--sheet 이름 --new-sheet`) |
| 엑셀에 수식 쓰기 | `excel.py write --range A2:B101 --formula "=ROW()-1" "=A2*2"` |
| 엑셀 차트 | `excel.py chart --range A1:B101 --type line --title "제목"` (`--action list` / `--action delete --name 제목`) |
| 엑셀 서식 | `excel.py style --range A1:E10 --theme navy` / `--bg yellow --bold` / `--highlight ">=100:pink"` |
| 엑셀 기타 | `excel.py info`, `replace --find A --replace B`, `sheet --action add --name 이름`, `range --action merge --range A1:B1`, `export --save "결과.csv"` |
| 워드 읽기 / 쓰기 | `word.py read [--save]` / `word.py write --file "내용.md"` / 한 줄이면 `--text "## 제목\n- 항목"` / `--image "그림.png"` (`--anchor 글 --pos after`, 여러 곳이면 `--occurrence N`). 쪽 나누기는 `word.py write --page-break` (.md의 `---`는 구분선) |
| 워드 목차 / 범위 읽기 | `word.py outline` / `word.py read --section "1.1"` (장절, 하위 절 포함) / `--pages 3-5` / `--paragraphs 120-140` |
| 워드 검토 메모 / 변경 추적 | `word.py comment` (목록) / `word.py comment --action add --anchor "문서의 글" --text "메모"` (여러 개는 `--file "메모.csv"`, 열: 기준 글, 메모 내용) / `word.py revisions` |
| 워드 기타 | `word.py info`, `replace --find A --replace B [--bold]`, `export --save "보고서.pdf" [--page-number]` |
| PPT 읽기 / 쓰기 | `ppt.py read [--slides 1-3]` / `ppt.py write --slide new --title 제목 --file "내용.md"` / `--text "- 항목"` / `--image "그림.png"` / `--cards "제목:내용|제목:내용"` |
| PPT 발표자 노트 | 읽기 `ppt.py notes [--slides 1-3]` / 쓰기 `ppt.py notes --slide 3 --text "노트"` (기존 노트가 있으면 `--append`로 덧붙이기, 여러 줄은 `--file "노트.md"`) |
| PPT 기타 | `ppt.py info`, `replace --find A --replace B`, `slide --action move --slide 1 --to 3`, `export --save "발표.pdf"` |
| 차트를 워드/PPT에 그림으로 | `office.py transfer --from excel --chart "차트이름" --to word,ppt` |
| 엑셀 범위를 워드/PPT 표로 | `office.py transfer --from excel --range A1:D10 --to word` |
| 워드/PPT 내용을 엑셀로 | `office.py transfer --from word --what text --to excel` (`--what table --table 1`) |
| 문서 등록 (PDF/HWP/Office) | 폴더: `reader.py index --folder "<폴더>"` / 파일 하나 읽기: `reader.py read --file "<파일>" [--pages 1-3] [--save]` |
| 문서 검색 (근거 포함) | `reader.py search --query "A AND B" [--file 파일또는폴더] [--pages 3-7] [--section "장절 일부"] --save "근거.xlsx"` (띄어쓰기·기호는 무시: `SOI 1` = `SOI #1`. 사용자가 패턴을 말하면 `--regex`) |
| 캐시된 파일 조회 | `reader.py list [--name 이름일부] [--folder 폴더]` |
| 캐시 지우기 (원본은 그대로) | `reader.py clear --name 이름일부` / `--file "<파일>"` / `--folder "<폴더>"` / `--missing`(원본 없는 것) / `--all`. 먼저 `--dry-run`으로 목록을 보여 주고 확인받은 뒤 `next_commands[0]` 실행 |
| 메일 읽기 | `outlook.py current`, `search --query 말 [--from 이름]`, `read --id <id>` (`--id`를 빼면 보고 있는 메일) |
| 메일 첨부 내려받기 | 목록 `outlook.py attachments --id <id>` / 저장 `outlook.py attachments --id <id> --save` (다운로드\메일 첨부 폴더. 다른 폴더는 `--save "<전체경로 폴더>"`, 일부만 `--name 이름일부`). 결과의 `folder`와 파일 이름을 알린다 |
| 메일 쓰기 / 문서 첨부 | `outlook.py reply --body "내용" --attach "C:\전체경로\파일.pdf"`. `--attach`는 `new --to 주소 --subject 제목 --body 내용`, `forward --to 주소`에도 같다. 여러 파일은 `;`로 구분 |
| 되돌리기 / 창 | `office.py undo` (목록 `--list`: 사용자가 고른 항목의 `command`를 그대로 실행), `office.py windows`, `office.py focus --app excel` |
| 결과 배치 확인 | `ppt.py info --detail --slides N`, `excel.py info --detail`, `word.py info --detail` (`layout_warnings`가 0이 아니면 `warnings`를 알림) |

## 2. 여러 단계 작업
1. 결과 JSON에 `next_commands`가 있으면 다음 단계는 `next_commands[i].command`를 **글자 그대로** 실행한다. 명령을 직접 만들지 않는다.
2. 문서 내용을 다른 문서로 옮길 때는 내용을 복사해 인자로 넣지 않는다. `--save`로 파일을 만들고, 결과의 `output_file`을 `--file`로 넘긴다.
3. 예) 사인 그래프를 워드에 넣기:
   1. `excel.py write --range A1:B101 --header "x,sin(x)" --formula "=(ROW()-2)*4*PI()/99" "=SIN(A2)"`
   2. `excel.py chart --range A1:B101 --type line --title "사인 2주기"`
   3. 결과의 `next_commands[0]` 실행 (워드/PPT로 전달. 받는 앱이 안 열려 있으면 아무것도 바꾸지 않고 `[추천]`에 열린 앱만 넣은 명령을 줌)

## 3. 결과 처리 규칙 (반드시 따른다)

| `agent_directive` 첫 단어 | 할 일 |
|---|---|
| `VERIFIED:` | 완료. `preview`를 사용자에게 알린다. 같은 명령을 다시 실행하거나 결과를 다시 읽지 않는다. |
| `UNVERIFIED:` | 실행은 됐지만 확인하지 못했다. 다시 실행하지 않고 `unverified` 내용을 그대로 알린다. |
| `STOP_RETRY:` + message에 `[후보]` | 다시 실행하지 않는다. `[후보]`를 보여 주고 사용자가 고르게 한 뒤, 고른 값으로 실행한다 (시트: `--sheet 이름`, 앵커·장절: `--occurrence 번호`, 문서: message가 말한 옵션(`--target`, `--to-target`, `--from-target`, `--name`)에 전체경로). 새 시트는 `excel.py write`에서만, 사용자가 원한다고 답한 경우에만 `--new-sheet`. |
| `STOP_RETRY:` + message에 `[추천]` | 추천된 값으로 **1번만** 다시 실행한다. 또 실패하면 message를 그대로 알린다. |
| `STOP_RETRY:` (위 둘 없음) | 다시 실행하지 않는다. message를 그대로 알린다. `needs_user_decision`(없는 시트, 지워질 값, 덮어쓸 칸, 되돌리기 위험)이면 사용자에게 묻고, 사용자가 원할 때만 directive가 말한 옵션을 붙여 1번 실행한다. |
| `ERROR:` | 다시 실행하지 않는다. message를 그대로 알린다. |

- JSON에 `recovered`가 있으면 그 문장을 그대로 알린다. `warnings`가 있으면 함께 알린다.
- 문서 검색 결과는 항목마다 근거를 붙여 알린다: `evidence`(예: `규격.pdf p.12, 1. SCOPE > 1.1. Scope (1.1)`)를 그대로 쓴다. `page`가 null인 형식(엑셀, HWP 등)은 장절(`section_text`)과 번호(`number`), 위치(`loc`)로 알린다.
- `search` 결과가 `UNVERIFIED:`이면 `not_searched`(검색하지 못한 부분)를 그대로 알린다. 0건이어도 "근거 없음"이라고 단정하지 않는다.
- `search`는 `--file`이 없으면 등록된 문서(캐시)만 찾는다. `--file`을 주면 그 파일/폴더의 새 문서를 먼저 등록한다(`indexed_now`를 사용자에게 알린다). `not_indexed` 오류면 사용자에게 폴더를 물어 `--file "<폴더>"`로 실행한다.
- 메일은 기본으로 작성 창만 띄운다. `--send`는 사용자가 **발송을 분명히 요청한 경우에만** 붙인다.
- 아래는 사용자가 요청한 경우에만 쓴다: 시트/슬라이드 삭제, `ppt.py slide --action clear --all-slides`, `--new-sheet`(새 시트), `--overwrite`(기존 시트·기존 노트·기존 슬라이드 제목·값이 있는 칸을 덮어쓰기).
- `replace`의 `--find`는 비워 두지 않는다 (빈 칸 채우기는 `excel.py write --range B2:B10 --value 0`).
- 새 파일은 사용자가 준 전체 경로로만 만든다 (`--target "C:\폴더\새파일.xlsx"`). 파일 이름만으로는 만들지 않는다.
- 저장 경로에 같은 이름의 파일이 있으면 도구가 덮어쓰지 않고 새 이름으로 저장한다. 결과의 `output_file`과 `recovered`를 그대로 알린다.
- 잘못 고쳤으면 `office.py undo`로 되돌린다. 저장한 적 없는 문서도 그 문서를 되돌린다 (결과 `preview`를 그대로 알린다).
