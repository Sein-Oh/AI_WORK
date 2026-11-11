---
name: office-mail
description: Outlook 메일과 일정. 받은·보낸 메일 목록·검색(기간, 보낸 사람, 폴더), 메일 읽기, 대화 전체 보기, 일정 보기·만들기·지우기, 첨부 목록·내려받기, 내려받은 eml 폴더 검색, 새 메일·답장·전체 답장·전달 작성(문서 첨부, 참조)과 확인 후 발송. 메일 답장 작성, 첨부 저장, 일정 확인 요청에 사용. 엑셀은 office-excel, 워드/PPT는 office-doc, 내려받은 문서 검색은 office-spec.
license: Proprietary
compatibility: opencode
metadata:
  version: "1.6.4"
  platform: windows
---

# office-mail 스킬 (Windows)

## 규칙
1. 아래 표의 명령을 그대로 쓰고 `<...>` 값과 `[...]` 안의 필요한 옵션만 바꾼다. 옵션을 지어내지 않는다.
2. PowerShell에서는 명령 앞에 전역 지침의 인코딩 줄을 붙인다.
3. 명령은 **1번만** 실행하고 마지막 1줄 JSON만 읽는다. `next_commands`는 그 `command`를 글자 그대로 실행한다. `confirm_commands`는 사용자 "예" 뒤에만 1번 (사용자가 "묻지 말고"라고 했으면 생략, 단 `"destructive": true` 항목은 늘 묻는다).
4. 메일은 `list`·`search`·`thread` 결과의 `id`(`--id 3`)를 글자 그대로 쓴다. 번호는 목록마다 이어서 늘어나고 1부터 다시 매기지 않아, 옛 번호는 같은 메일이거나 "없는 번호" 오류다. 번호 오류("목록을 다시 보세요")면 `next_commands`로 목록을 다시 보고 새 `id`를 쓴다. `id`가 긴 글자(EntryID)여도 그대로 쓴다. `--id`를 빼면 사용자가 보고 있는 메일이다.
5. 긴 본문과 여러 줄 글은 인자로 넣지 않는다. 파일 쓰기 도구로 UTF-8 .md를 만들고 `--body-file "<그 경로>"`로 넘긴다.
6. 첨부 경로는 전체 경로로 쓴다. 경로는 결과 JSON 값을 그대로 쓴다. 도구 코드(.py)는 읽지 않는다. 명령 표에 없는 일은 전역 지침 8장대로 한다 (8장의 "하지 않는 일"(두 파일 비교 등)은 파이썬으로도 하지 않고 직접 하는 곳만 알린다).

## 명령
| 할 일 | 명령 |
|---|---|
| 목록 | `python "C:\AIKIT\office-cli\outlook.py" list [--unread] [--with-attachments] [--folder <sent/drafts/"폴더">] [--limit 20]` (기간 `--days 7`(오늘 포함) 또는 `--since 2026-11-01 [--until 2026-11-05]`, 검색도 같음) |
| 검색 | `python "C:\AIKIT\office-cli\outlook.py" search --query <말> [--from <이름>] [--folder sent] [--where subject]` |
| 읽기 | `python "C:\AIKIT\office-cli\outlook.py" read --id 3` / 보고 있는 메일 `python "C:\AIKIT\office-cli\outlook.py" current` |
| 대화 전체 | `python "C:\AIKIT\office-cli\outlook.py" thread --id 3` |
| 일정 보기 / 만들기 | `python "C:\AIKIT\office-cli\outlook.py" calendar --days 7` / `python "C:\AIKIT\office-cli\outlook.py" event --subject "<제목>" --start "2026-11-12 14:00" [--minutes 60] [--location "<장소>"]` (종일은 `--start 2026-11-12`, 내 일정에만 저장. 회의 요청은 없음) |
| 일정 지우기 | `python "C:\AIKIT\office-cli\outlook.py" event --action delete --subject "<제목>" --start "2026-11-12 14:00" --dry-run` (확인 뒤 `confirm_commands[0]`. 반복 일정·회의는 Outlook에서 직접) |
| 폴더 목록 | `python "C:\AIKIT\office-cli\outlook.py" folders` (더 깊이 `--depth 3`) |
| 첨부 목록 | `python "C:\AIKIT\office-cli\outlook.py" attachments --id <id>` |
| 첨부 저장 | `python "C:\AIKIT\office-cli\outlook.py" attachments --id <id> --save` (다운로드\메일 첨부 폴더. 다른 폴더 `--save "<전체경로 폴더>"`, 일부만 `--name <이름일부>` 또는 목록의 번호 `--index 1,3`) |
| 답장 | `python "C:\AIKIT\office-cli\outlook.py" reply --id 3 --body "<내용>" [--attach "C:\전체경로\파일.pdf"]` (전체 답장 `--all`, 긴 본문 `--body-file "<본문.md>"`) |
| 새 메일 | `python "C:\AIKIT\office-cli\outlook.py" new --to <주소> --subject "<제목>" --body "<내용>" [--cc <주소>] [--attach "<전체경로>"]` |
| 전달 | `python "C:\AIKIT\office-cli\outlook.py" forward --id 3 --to <주소> [--attach "<전체경로>"]` |
| eml 폴더 검색 (Outlook 없이) | `python "C:\AIKIT\office-cli\reader.py" search --query "<말>" --file "<eml 폴더>"` (`;` 그리고, `|` 또는. 처음은 등록이라 오래 걸림, `CONTINUE`면 같은 명령) |
| 초안 보내기 | `python "C:\AIKIT\office-cli\outlook.py" send --id <draft_id>` (new·reply·forward 결과의 `confirm_commands[0]`, 사용자 "예" 뒤에만) |

- eml 검색 결과는 메일 한 통에 하나다. 메일마다 `mail`(제목, 보낸 사람, 날짜, 첨부 이름)과 `path`(원본 eml)를 알리고, 원본은 사용자가 그 eml을 열어 본다고 안내한다. 첨부 안의 글은 찾지 않는다.
- `--attach` 여러 파일은 `;`로 구분한다 (`"C:\a.pdf;C:\b.xlsx"`).
- 첨부를 저장하면 결과의 `folder`와 파일 이름을 알린다. 다른 폴더는 사용자가 준 전체 경로로만. 같은 이름이 있으면 덮어쓰지 않고 새 이름으로 저장된다. 내용을 보려면 결과의 `next_commands`(reader.py read, 여럿이면 reader.py index)를 그대로 실행한다.

## 결과 처리
`agent_directive` 첫 단어별 행동은 전역 지침 4장대로. 추가로:
- `new`·`reply`·`forward`는 보내지 않는다. 초안을 저장하고 작성 창만 띄운다. 발송 순서:
  1. `new`·`reply`·`forward`를 1번 실행한다 (답장·전달은 `--id`로 원본을 지정).
  2. 결과의 `original`(제목·보낸사람)과 받는 사람·참조·제목을 사용자에게 보여 주고 "이 메일이 맞는지, 보낼지" 확인을 받는다.
  3. 사용자가 "예"(보내라)라고 하면 `confirm_commands[0]`(`send --id <draft_id>`)을 글자 그대로 1번 실행한다. 확인 전에는 실행하지 않는다.
  4. 사용자가 작성 창에서 직접 보냈으면 `send`를 실행하지 않는다. 같은 `new`·`reply`·`forward`를 다시 실행하지 않는다(초안이 하나 더 생김).
  5. `send`가 "발송하지 않았습니다"(이 도구가 만든 초안이 아님, 이미 보냄, 임시 보관함에 없음)면 다시 실행하지 않고 message를 알린다.
- "보내지 않은 초안이라 원본으로 쓸 수 없습니다"면 `next_commands[0]`으로 목록을 보고 원본 메일의 `id`로 다시 실행한다.
- `reply`·`forward`·`attachments` 결과의 `original`(또는 `mail`)의 제목·보낸사람을 사용자에게 알린다.
- `elevated_process`이면 다시 실행하지 않고 message를 그대로 알린다.
- 0건(`count: 0`)이면 "찾은 메일 없음"과 조건(폴더, 기간)을 알린다.
- `recovered`와 `warnings`(겹치는 일정 등)가 있으면 그대로 알린다.
- `error_type: not_supported`(도구에 없는 기능)면 다시 실행하거나 파이썬으로 대신하지 않고 message의 방법을 사용자에게 알린다.
