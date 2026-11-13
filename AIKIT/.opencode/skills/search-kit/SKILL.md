---
name: search-kit
description: PDF/HWP/HWPX/Office/CSV/TXT 파일을 Office 없이 읽고, 폴더를 등록(캐시)해 쪽·장절·번호 근거와 함께 검색한다. 장절 읽기, 엑셀 칸에 검토 메모 남기기, 캐시 조회·지우기. 규격/규정 근거 찾기, 문서 내용 검색 요청에 사용. 열린 엑셀 편집은 excel-kit, 워드는 word-kit, PPT는 ppt-kit, 메일은 outlook-kit.
license: Proprietary
compatibility: opencode
metadata:
  version: "1.6.7"
  platform: windows
---

# search-kit 스킬 (Windows)

## 규칙
1. 아래 표의 명령을 그대로 쓰고 `<...>` 값과 `[...]` 안의 필요한 옵션만 바꾼다. 옵션을 지어내지 않는다.
2. PowerShell에서는 명령 앞에 전역 지침의 인코딩 줄을 붙인다.
3. 명령은 **1번만** 실행하고 마지막 1줄 JSON만 읽는다. `next_commands`는 그 `command`를 글자 그대로 실행한다. `confirm_commands`는 사용자 "예" 뒤에만 1번 (사용자가 "묻지 말고"라고 했으면 생략, 단 `"destructive": true` 항목은 늘 묻는다).
4. 검색한 내용을 다른 문서에 넣을 때는 내용을 복사해 인자로 넣지 않는다. `--save`로 파일을 만들고 결과의 `output_file`을 word-kit·ppt-kit·excel-kit 명령의 `--file`로 넘긴다.
5. 경로는 결과 JSON 값을 그대로 쓴다. 도구 코드(.py)는 읽지 않는다. 명령 표에 없는 일은 전역 지침 8장대로 한다 (8장의 "나. 하지 않는 일"(인쇄, 회의 요청 등)은 파이썬으로도 하지 않고 직접 하는 곳만 알린다).

## 명령
| 할 일 | 명령 |
|---|---|
| 폴더 등록 | `python "C:\AIKIT\office-cli\reader.py" index --file "<폴더>"` (하위 폴더는 깊이 3까지, 전체는 `--recursive`) |
| 파일 읽기 | `python "C:\AIKIT\office-cli\reader.py" read --file "<파일>" [--pages 1-3] [--save]` (쪽 수·장절 수만 `--format summary`) |
| 장절 목록 (큰 문서) | `python "C:\AIKIT\office-cli\reader.py" outline --file "<파일>"` |
| 장절 읽기 | `python "C:\AIKIT\office-cli\reader.py" read --file "<파일>" --section "<장절 일부>" [--pages 3-7]` (번호 `2.1`은 하위 절 포함, 이름은 일부. `units[].evidence`가 근거) |
| 검색 | `python "C:\AIKIT\office-cli\reader.py" search --query "A;B" [--file "<파일또는폴더>"] [--pages 3-7] [--section "<장절 일부>"] [--save]` (근거를 엑셀로 받으려면 `--save "근거.xlsx"`) (`;` 그리고, `|` 또는. 더 많이 `--top 30`, 단락 전체 `--full`) |
| 엑셀 칸 메모 | 목록 `python "C:\AIKIT\office-cli\excel.py" comment --target "<엑셀 파일>" [--range B2:D20]` / 넣기 `python "C:\AIKIT\office-cli\excel.py" comment --action add --target "<엑셀 파일>" --range B5 --text "<의견>"` (여러 개 `--file "<메모.csv>"`, 열: 칸, 메모. 판정·검토 의견은 사용자가 이번 요청에서 넣으라고 했으면 바로 넣고 판정·근거와 되돌리는 법(되돌리기 줄의 명령)을 알린다, 아니면 넣기 전에 보여 주고 "예"를 받는다 (전역 지침 3장 5번)) |
| 캐시 목록 | `python "C:\AIKIT\office-cli\reader.py" list [--name <이름일부>] [--folder "<폴더>"]` |
| 캐시 지우기 (원본은 그대로) | `python "C:\AIKIT\office-cli\reader.py" clear --name <이름일부> --dry-run` (`--file "<파일>"` / `--folder "<폴더>"` / `--missing`(원본 없는 것) / `--all`. 목록을 보여 주고 확인받은 뒤 `confirm_commands[0]`) |
| 되돌리기 (메모) | `python "C:\AIKIT\office-cli\office.py" undo` |

- 검색은 띄어쓰기·기호·조사·단위 표기를 무시한다 (`SOI 1` = `SOI #1`). 사용자가 패턴을 말하면 `--regex`.
- 사용자가 동의어를 주면 `C:\AIKIT\office-cli\data\synonyms.txt`에 적는다 (UTF-8, 한 줄에 한 묶음, 쉼표 구분, 있던 줄은 지우지 않음).

## 결과 처리
`agent_directive` 첫 단어별 행동은 전역 지침 4장대로. 추가로:
- 근거는 항목마다 `evidence`(예: `규격.pdf p.12, 1.1. Scope (1.1)`)를 그대로 쓴다. `page`가 null(엑셀, HWP 등)이면 장절 `section_text`, 번호 `number`, 위치 `loc`로 알린다.
- `match_mode`가 `number`(숫자 값만 같은 곳, 실제 글 `number_words`), `partial`(일부 검색어만), `fuzzy`(한 글자 다른 낱말)이면 `count: 0`이고 결과는 `candidates`다. 근거로 쓰지 말고 "정확히 일치한 곳은 없고 비슷한 곳이 있다"고 알려 사용자에게 확인받는다. 부호·대시가 다른 숫자(`40 ℃`와 `-40 ℃`)도 `number`다.
- `search`가 `UNVERIFIED:`이면 `not_searched`를 알린다. 0건이어도 "근거 없음"이라고 단정하지 않는다.
- `search`는 `--file`이 없으면 등록된 문서만 찾는다. `--file`을 주면 새 문서를 먼저 등록한다(`indexed_now`를 알림). `not_indexed` 오류면 사용자에게 폴더를 물어 `--file "<폴더>"`로 실행한다.
- `truncated`이면 `full_result_file`을, `recovered`와 `warnings`는 그대로 알린다.
- `--save` 새 파일 경로는 지어 붙이지 않는다 (사용자가 준 경로만, 안 줬으면 `--save`만: 도구가 대상 파일 옆에 저장). 같은 이름이 있으면 새 이름으로 저장되니 `output_file`·`recovered`를 알린다.
- 메모의 `--overwrite`(있던 메모 바꾸기)는 사용자가 요청했을 때만 붙인다.
