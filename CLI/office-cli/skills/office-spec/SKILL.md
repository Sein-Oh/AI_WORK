---
name: office-spec
description: PDF/HWP/HWPX/Office/CSV/TXT 파일을 Office 없이 읽고, 폴더를 등록(캐시)해 쪽·장절·번호 근거와 함께 검색한다. 장절 읽기, 두 판의 문서 비교(바뀐 숫자), 엑셀 표를 규격서와 대조해 일치·불일치와 근거를 찾고 검토 메모 남기기, 캐시 조회·지우기. 규격/규정 근거 찾기, 문서 내용 검색, 개정 비교 요청에 사용. 열린 엑셀 편집은 office-excel, 워드/PPT는 office-doc, 메일은 office-mail.
license: Proprietary
compatibility: opencode
metadata:
  version: "1.4.70"
  platform: windows
---

# office-spec 스킬 (Windows)

## 규칙
1. 아래 표의 명령을 그대로 쓰고 `<...>` 값과 `[...]` 안의 필요한 옵션만 바꾼다. 옵션을 지어내지 않는다.
2. PowerShell에서는 명령 앞에 전역 지침의 인코딩 줄(`$env:PYTHONIOENCODING='utf-8'; ...`)을 붙인다.
3. 명령은 **1번만** 실행하고 마지막 1줄 JSON만 읽는다. `next_commands`가 있으면 다음 단계는 그 `command`를 글자 그대로 실행한다.
4. 검색한 내용을 다른 문서에 넣을 때는 내용을 복사해 인자로 넣지 않는다. `--save`로 파일을 만들고 결과의 `output_file`을 office-doc·office-excel 명령의 `--file`로 넘긴다.
5. 경로는 결과 JSON 값을 그대로 쓴다. 도구 코드(.py)는 읽지 않는다. 방법이 없으면 사용자에게 알리고 멈춘다.

## 명령
| 할 일 | 명령 |
|---|---|
| 폴더 등록 | `python "C:\CLI\office-cli\reader.py" index --folder "<폴더>"` |
| 파일 읽기 | `python "C:\CLI\office-cli\reader.py" read --file "<파일>" [--pages 1-3] [--save]` |
| 장절 읽기 | `python "C:\CLI\office-cli\reader.py" read --file "<파일>" --section "<장절 일부>" [--pages 3-7]` (번호 `2.1`은 하위 절 포함, 이름은 일부. `units[].evidence`가 근거) |
| 검색 | `python "C:\CLI\office-cli\reader.py" search --query "A AND B" [--file "<파일또는폴더>"] [--pages 3-7] [--section "<장절 일부>"] --save "<근거.xlsx>"` |
| 두 판 비교 | `python "C:\CLI\office-cli\reader.py" diff --file "<이전 판>" --file2 "<새 판>" [--section "3"] [--save "<차이.xlsx>"]` |
| 엑셀 표를 규격서와 대조 | `python "C:\CLI\office-cli\excel.py" check --range B2:D50 --item-col B --value-col D --spec "<규격 파일 또는 폴더>" [--section "3.2"] [--section-col E] [--save "<대조결과.xlsx>"]` |
| 엑셀 칸 메모 | 목록 `python "C:\CLI\office-cli\excel.py" comment [--range B2:D20]` / 넣기 `python "C:\CLI\office-cli\excel.py" comment --action add --range B5 --text "<의견>"` (여러 개 `--file "<메모.csv>"`, 열: 칸, 메모) |
| 캐시 목록 | `python "C:\CLI\office-cli\reader.py" list [--name <이름일부>] [--folder "<폴더>"]` |
| 캐시 지우기 (원본은 그대로) | `python "C:\CLI\office-cli\reader.py" clear --name <이름일부> --dry-run` (`--file "<파일>"` / `--folder "<폴더>"` / `--missing`(원본 없는 것) / `--all`. 목록을 보여 주고 확인받은 뒤 `next_commands[0]`) |
| 되돌리기 (메모) | `python "C:\CLI\office-cli\office.py" undo` |

- 검색은 띄어쓰기·기호·조사·단위 표기를 무시한다 (`SOI 1` = `SOI #1`, `정격전압은 5V` = `정격 전압이 5.0 V`). 사용자가 패턴을 말하면 `--regex`.
- 사용자가 동의어를 주면 `C:\CLI\office-cli\data\synonyms.txt`에 적는다 (UTF-8, 한 줄에 한 묶음, 쉼표 구분, 있던 줄은 지우지 않음).
- 규격 대조: `check` 한 번 -> `counts`와 불일치·`needs_review` 행을 근거와 함께 알림 (`needs_review`는 사용자가 정함) -> 사용자가 원하면 `next_commands[0]`(메모 미리보기) -> 확인 뒤 그 명령에서 `--dry-run`을 뺀 명령. 한 칸만 볼 때는 위 검색 명령에 `--section "<장절>"`.

## 결과 처리
`agent_directive` 첫 단어별 행동은 전역 지침 4장대로. 추가로:
- 근거는 항목마다 `evidence`(예: `규격.pdf p.12, 1. SCOPE > 1.1. Scope (1.1)`)를 그대로 쓴다. `page`가 null(엑셀, HWP 등)이면 장절 `section_text`, 번호 `number`, 위치 `loc`로 알린다.
- `match_mode`가 `partial`(일부 검색어만)이나 `fuzzy`(한 글자 다른 낱말)이면 `count: 0`이고 결과는 `candidates`다. 근거로 쓰지 말고 "정확히 일치한 곳은 없고 비슷한 곳이 있다"고 알려 사용자에게 확인받는다.
- `search`가 `UNVERIFIED:`이면 `not_searched`를 그대로 알린다. 0건이어도 "근거 없음"이라고 단정하지 않는다.
- `search`는 `--file`이 없으면 등록된 문서만 찾는다. `--file`을 주면 새 문서를 먼저 등록한다(`indexed_now`를 알림). `not_indexed` 오류면 사용자에게 폴더를 물어 `--file "<폴더>"`로 실행한다.
- `diff`는 숫자 변경 `number_changes`를 먼저, 삭제·추가·변경 항목마다 `old_evidence`·`new_evidence`를 붙여 알린다.
- `truncated`이면 `full_result_file`을 알린다. `recovered`와 `warnings`는 그대로 알린다.
