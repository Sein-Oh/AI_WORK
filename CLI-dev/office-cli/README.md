# office-cli 1.4

Excel, Word, PowerPoint, Outlook을 AI 에이전트(26B~30B 경량 LLM)가 명령 한 줄로 다루게 하는 CLI 도구입니다.
PDF, HWP, Office 파일은 Office 없이 읽고, 캐시에 등록해 쪽·장절·번호 근거와 함께 검색합니다.

- 에이전트용 요약: [`SKILL.md`](../../CLI/office-cli/SKILL.md)
- 변경 이력과 남은 할 일: [`CHANGELOG.md`](CHANGELOG.md)
- 1.0 설계 배경과 결정 기록: [`docs/history/office-cli_WORK_PLAN_1.0.md`](../../docs/history/office-cli_WORK_PLAN_1.0.md)

## 1. 구성

```text
CLI/office-cli/  실행에 필요한 파일만 (사내 PC C:\CLI\office-cli)
├── SKILL.md     에이전트용 스킬
├── excel.py     엑셀   read / info / write / replace / style / export / sheet / range / chart
├── word.py      워드   read / info / outline / write / replace / comment / revisions / export
├── ppt.py       PPT    read / info / write / notes / replace / export / slide
├── outlook.py   메일   list / search / read / current / folders / new / reply / forward
├── reader.py    파일   read / index / search / list / clear   (Office 불필요: PDF, HWP, HWPX, DOCX, PPTX, XLSX, CSV, TXT, MD)
├── office.py    공통   windows / focus / transfer / undo
├── _runtime.py  공통 런타임 (응답 계약, 경로, 입출력, 마크다운 해석)
└── _office.py   Office COM 공통 (세션, 스냅샷, 색/테마)

CLI-dev/office-cli/  개발·실기 테스트용 (사내 PC C:\CLI-dev\office-cli, C:\CLI와 나란히)
├── README.md / CHANGELOG.md / TEST_CHECKLIST.md / requirements.txt
├── tools/run_win_tests.py  Windows 실기 테스트 러너 (판단 규칙 tools/win_test_rules.py)
├── tools/office-cli-test/SKILL.md  실기 테스트를 에이전트가 돌리는 스킬 (스킬 폴더 office-cli-test에 SKILL.md만 복사)
└── tests/       오프라인 테스트 (가짜 COM + 실제 샘플 파일, 도구는 ../../CLI/office-cli에서 불러옴)
```

설치 (폐쇄망은 [`docs/WINDOWS_OFFLINE.md`](../../docs/WINDOWS_OFFLINE.md) 참고):
```bat
python -m pip install -r CLI-dev/office-cli/requirements.txt
```

## 2. 공통 규칙

### 출력 계약 (저장소 ACI 계약)
- stdout은 정확히 1줄 JSON입니다. 진행 로그는 stderr(`[PROGRESS]`, `[WARN]` 등)입니다.
- 종료 코드: 0 성공, 1 오류, 2 인자 오류, 130 중단.
- `agent_directive`:
  - `VERIFIED:` 완료
  - `UNVERIFIED:` 실행은 됐지만 결과를 확인하지 못함 (예: 검색하지 못한 부분 `not_searched`). 다시 실행하지 않고 `unverified`를 알림
  - `STOP_RETRY:` 입력이나 대상 문제. `[추천]`이 있으면 그 값으로 1회만 다시 시도
  - `STOP_RETRY:` + `[후보]`(`error_type: "ambiguous"`, `candidates` 목록): 후보가 둘 이상이라 도구가 고르지 않음. 사용자가 고른 값으로 다시 실행
  - `ERROR:` 내부 오류
- 공통 키: `status`, `tool`, `command`, `preview`, `agent_directive`. 오류 시 `error_type`, `message`가 추가됩니다. COM 오류는 `hresult`도 들어갑니다.
- `target_resolved`: 실제로 작업한 문서(앱, 이름, 경로, 시트/슬라이드)
- `next_commands`: 이어서 실행할 명령 `[{purpose, command}]`. 완성된 형태라 그대로 실행하면 됩니다.

### 대상 문서 (`--target`)
| 값 | 의미 |
|---|---|
| `active` (기본) | 그 앱에서 사용자가 보고 있는 문서 |
| 열린 문서 이름 | 예: `보고.xlsx` (대소문자 무시). 같은 이름의 문서가 둘 이상 열려 있으면 고르지 않고 `[후보]`(전체 경로)로 오류 |
| 파일 경로 | 닫힌 파일. 읽기는 라이브러리로, 쓰기는 사용자 창과 분리된 백그라운드 인스턴스로 열고 저장 후 닫음 |

### 화면과 안전
- **읽기는 화면을 움직이지 않습니다.** 활성 시트, 스크롤, PPT 슬라이드가 그대로입니다.
- **쓰기는 고친 시트/슬라이드로만 이동합니다.** `--no-focus`를 주면 이동하지 않습니다.
- **쓰기 전 스냅샷을 자동으로 만듭니다**(문서별 5개 보관, 50MB 초과 문서는 생략). 같은 경로에 다른 문서를 새로 만들면 이전 문서의 스냅샷은 되돌리기 대상에서 빠집니다(파일 속 만든 시각으로 판별).
  - 원본 문서는 저장하거나 경로를 바꾸지 않습니다.
  - `office.py undo`로 되돌립니다. 결과 `restored_as`:
    | 경우 | `restored_as` | 동작 |
    |---|---|---|
    | 저장된 문서, 열려 있음 | `reopened_in_office` | 현재 상태를 `pre_undo_backup`에 보관, 닫고 스냅샷으로 바꾼 뒤 다시 엶 |
    | 저장된 문서, 닫혀 있음 | `file_replaced` | 파일을 스냅샷으로 바꿈 (기존 파일은 `pre_undo_backup`) |
    | 저장한 적 없는 워드 문서, 열려 있음 | `restored_in_place` | 같은 창에서 내용만 스냅샷으로 바꿈 (현재 내용은 `pre_undo_backup` .xml, 문서는 저장하지 않음) |
    | 저장한 적 없는 엑셀/PPT 문서, 열려 있음 | `replaced_untitled_document` | 현재 상태를 `pre_undo_backup`에 보관, 스냅샷을 새 제목 없는 문서로 열고 바뀐 문서는 저장하지 않고 닫음 (문서 수 그대로, 창 이름은 바뀜) |
    | 저장한 적 없는 문서가 닫혔거나 같은 이름의 저장 안 된 문서가 둘 이상 | `new_untitled_document` | 고르지 않고 스냅샷을 새 제목 없는 문서로만 엶. `preview`에 이유를 적음 |
  - 저장한 적 없는 문서는 스냅샷 때의 문서 이름(`문서1`, `Book1`)으로 찾습니다(1.2.2, 시나리오 12). 이전에는 항상 새 문서로 열어 바뀐 문서가 그대로 남았습니다.
  - 스냅샷을 만들지 못한 쓰기(DRM, 50MB 초과)와 새로 만든 파일에 쓰기는 "되돌릴 지점 없음" 표식을 남깁니다. 직전 작업이 그렇다면 `undo`는 더 오래된 지점으로 넘어가지 않고 `needs_user_decision`(사용자가 `--list`에서 고른 `--pick N`만 실행).
  - 되돌리기 지점 뒤에 파일이 다시 저장되었으면(지점 2분 뒤 이후의 수정 시각) 그 사이 저장한 내용이 사라지므로 `needs_user_decision`. 사용자가 원하면 `--overwrite`.
    `--pick N`으로 앞 지점을 고르면 그 사이 이 도구가 저장한 것은 사용자 저장으로 보지 않습니다(같은 문서의 가장 새 지점 시각과 비교).
  - `--pick N`으로 되돌리면 같은 문서의 더 새 지점(스냅샷, 되돌릴 지점 없음 표식)도 함께 지웁니다(`removed_newer_points`). 남겨 두면 다음 `undo`가 되돌린 뒤의 상태로 앞으로 갔습니다(1.4.10).
  - 열린 문서를 닫은 뒤 파일 복사가 실패하면(백신·잠금) 원본을 다시 열고 `undo_failed` 오류에 `pre_undo_backup`을 줍니다.
- **읽기는 아무것도 만들거나 저장하지 않습니다.** 없는 시트나 파일을 지정하면 `[추천]`과 함께 오류를 냅니다.
- **추측해서 쓰지 않습니다.** 쓰기 위치가 확실하지 않으면 아무것도 바꾸지 않고 오류를 냅니다.
  - 엑셀 쓰기에서 없는 시트: 새로 만들지 않습니다. 표기만 다른 시트(`Sheet 1`과 `Sheet1`)가 있으면 `[추천]`. 그 밖에는 `error_type: "needs_user_decision"`으로 있는 시트 목록을 알리고, 비슷한 기존 시트(`Sheet2`와 `Sheet1`)가 있으면 그 이름만 `[후보]`에 넣습니다. 새 시트는 `--new-sheet`로만 만듭니다.
    - message와 `[후보]`에는 `--new-sheet`나 "만드세요" 같은 명령문을 넣지 않습니다. 경량 LLM이 그 문장을 따라 사용자에게 묻지 않고 새 시트를 만들었기 때문입니다(1.2.2, 시나리오 10). `--new-sheet`는 `agent_directive`에만 "사용자가 만들라고 답한 경우에만"이라는 조건과 함께 적습니다.
    - 예: `시트 'Sheet9'이(가) 없습니다 (있는 시트: Sheet1). 새 시트를 만들지 않았고 아무 데도 쓰지 않았습니다. 이름이 비슷한 기존 시트에 쓸지, 새 시트를 만들지 사용자에게 물어보세요. [후보] Sheet1`
  - 워드 `--anchor`가 여러 단락에 있으면 `[후보]`(번호와 단락 앞부분)로 오류. `--occurrence N`으로 고릅니다. 찾지 못하면 비슷한 문구를 안내합니다 (공백·대소문자 차이 등 확실할 때만 `[추천]`, `--pos replace`에서는 더 엄격).
  - `office.py transfer --to excel --sheet 이름`은 새 시트 이름입니다(transfer는 항상 새 시트에 받습니다). 이미 있으면 덮어쓰지 않고 비어 있는 이름을 `[추천]`, 덮어쓰기는 `--overwrite`.
  - 쓰기 `--target`에 파일 이름만 주었는데 열린 문서도 파일도 아니면 새 파일을 만들지 않고 "찾을 수 없음"(열린 문서가 있으면 `[후보]`). 새 파일은 전체 경로를 줄 때만 만듭니다(1.4.10, 이전에는 현재 폴더에 만들어 거기에 씀).
- **메일**: 기본은 작성 창만 띄웁니다. `--send`는 명시했을 때만 쓰고, 수신자 주소를 확인한 뒤 발송합니다.

### 파일과 폴더
| 항목 | 기본 위치 | 바꾸는 방법 |
|---|---|---|
| data 폴더 (파일명만 쓴 입력/출력) | `<도구폴더>\data` | 환경변수 `OFFICE_CLI_DATA_DIR` 또는 `config.json`의 `data_dir` |
| 작업 폴더 (`--save` 경로 생략 시, 7일 후 자동 정리) | `data\_work` | - |
| 캐시 (등록한 문서 색인, 스냅샷) | `<도구폴더>\cache` | `OFFICE_CLI_CACHE_DIR` 또는 `cache_dir` |
| 설정 (`transfer --probe` 결과 등) | `<도구폴더>\config.json` | - |
| 오류 로그 (실패한 실행 자동 기록) | `data\logs\errors.txt` | data 폴더를 따라감 |

- 입력 파일은 절대경로 -> data 폴더 -> 작업 폴더 순서로 찾습니다. 없으면 비슷한 파일명을 `[추천]`으로 알려 줍니다.
- 텍스트 입력은 UTF-8, UTF-16, CP949를 자동으로 판별합니다. CSV 출력은 기본 utf-8-sig입니다(Excel에서 한글 정상).
- 결과가 길면 기본 200행 또는 약 8,000자에서 자르고 `truncated: true`와 이어 읽기 명령을 줍니다.

### 오류 로그
- 실패한 실행(exit 1, 인자 오류 exit 2)은 자동으로 `<data 폴더>\logs\errors.txt`에 덧붙습니다. `C:\CLI`에 배치했다면 `C:\CLI\office-cli\data\logs\errors.txt`(`config.json`의 `data_dir`나 `OFFICE_CLI_DATA_DIR`를 바꿨으면 그 폴더 아래).
- 항목: 시각, 도구 버전, Python, OS / 실행한 명령과 인자 / `error_type`과 message / 재시도 한도를 넘긴 복구 시도(`복구 시도:`) / 내부 오류(`ERROR:`)일 때만 traceback. 문서 본문은 기록하지 않습니다.
- 내용 옵션의 값은 `<N자 생략>`으로 남깁니다: `--body --text --subject --to --cc --bcc --from --query --find --replace --value --values --formula --header --title --cards --anchor --search --new-name`(`office.py`의 `--from/--to`, `ppt.py`의 `--to`는 앱 이름·위치라 그대로). 경로와 그 밖의 인자는 그대로 남깁니다(1.4.10).
- 성공, 사용자 중단(exit 130)은 기록하지 않습니다. 2MB를 넘으면 `errors.old.txt`로 넘기고 새로 시작합니다. 로그 경로는 stdout JSON에 넣지 않습니다(사람과 개발 에이전트의 사후 분석용).
- 실기 테스트 러너의 `errors_<시각>.txt`는 테스트 한 번의 판정 근거를 모은 별도 파일이고, 그 실행 중에 쌓인 도구 오류 로그(`<테스트 폴더>\_data\logs\errors.txt`)도 끝에 함께 옮겨 담습니다.

## 3. 명령 요약

### excel.py
| 명령 | 주요 옵션 |
|---|---|
| `read` | `--range`, `--sheet`, `--search 글자`(셀 주소 반환), `--max-rows`, `--save [경로.csv]`. `--target`이 CSV/TSV 파일이면 Excel 없이 읽음(인코딩 자동 판별, 숫자 변환, `source: file(csv)`). 닫힌 xlsx도 열린 문서와 같게 읽음: 범위를 생략하면 데이터가 시작하는 칸부터(UsedRange처럼), 거꾸로 준 범위(`C5:A1`)는 `A1:C5`, 크기 정보가 없는 파일도 실제 데이터 끝까지 |
| `info` | 시트별 사용 범위와 차트. `--all`: 열린 통합문서 전체. `--detail`: 차트 종류·계열 수·위치와 데이터를 가리는 차트 경고 |
| `write` | `--range` 시작 셀 또는 범위. 입력은 `--value`, `--values JSON`, `--file (.csv/.md 표/.txt)`, `--formula F1 [F2 ...]` 중 하나. `--header "a,b"`, `--append`(쓸 열 전체에서 가장 아래 행 다음에, 시작 행의 머리글과 같으면 생략), `--sheet`, `--new-sheet`(없는 시트를 새로 만들 때만). 쓸 칸에 값이나 수식(빈 글을 내는 수식 포함)이 있으면 아무것도 쓰지 않고 `needs_user_decision`(사용자가 원하면 `--overwrite`). `--append`도 쓸 칸을 확인하고, `--formula`/`--value`도 쓸 열 수만큼 가장 아래 행을 찾습니다. 모든 검사를 마친 뒤에 씁니다(실패하면 시트는 그대로). 쓴 뒤 칸을 다시 읽어 비교합니다(같으면 `VERIFIED:`, 다르면 `UNVERIFIED:` + 다른 칸, 1.4.6). 숫자만 `--formula`로 주면 `--values`를 `[추천]`. Excel이 다른 값으로 바꿔 버릴 글자(앞자리 0 코드 `01234`, 15자리 초과 숫자, `1-2`·`3/4`, `(100)`, `1E5`)는 글자 그대로 씁니다 |
| `replace` | `--find/--replace` 또는 `--map JSON`, `--whole`, `--all-sheets`, 서식 옵션. 수식 셀은 내용 보존. `--replace`도 서식 옵션도 없으면 바꿀 것이 없어 STOP_RETRY. `--range '요약'!A1:D9`처럼 시트를 주면 그 시트에서 읽고 씀. `--all-sheets`와 시트 이름이 든 `--range`는 함께 쓸 수 없음 |
| `style` | `--theme navy\|slate\|teal\|emerald`, `--bg`, `--color`, `--bold`, `--align`, `--number-format`, `--border`, `--highlight ">=100:pink"`(조건부 서식), `--clear` |
| `export` | `--save 경로.csv\|.xlsx\|.pdf`, `--columns A,C,E` |
| `sheet` | `--action list\|add\|rename\|copy\|delete\|select --name --new-name` |
| `range` | `--action autofit\|width\|height\|merge\|unmerge\|freeze\|unfreeze\|hide\|unhide\|clear\|copy\|move\|border`. `clear`는 `--range` 필수(시트 전체를 실수로 지우지 않게). `merge`는 왼쪽 위 밖의 칸에 값이 하나라도 있으면(병합하면 지워짐), `copy`/`move`는 대상에 값이 있으면 `needs_user_decision`(사용자가 원하면 `--overwrite`). `copy`는 대상의 수식(R1C1)을 원본과 비교해 확인(상대 참조는 값이 달라도 같은 수식). `autofit`은 비교할 값이 없어 `UNVERIFIED:`. `hide`/`unhide`/`autofit`의 `--axis`를 생략하면 `5:7` 같은 행 주소는 행, 그 밖은 열 |
| `chart` | `--action create\|list\|delete`, `--range`, `--type column\|bar\|line\|line_markers\|pie\|area\|scatter\|scatter_lines\|combo`, `--title`, `--at`, `--labels`, `--legend`, `--theme` |

- 수식이 1개면 범위 전체에 넣고 상대참조가 자동으로 조정됩니다. 여러 개면 열마다 하나씩 넣습니다.
- 첫 열이 숫자인 `line` 차트는 x축이 숫자 축이 되도록 `scatter_lines`로 그립니다(예: 사인 그래프).
- `--at`을 생략하면 차트는 시트의 데이터(사용 범위) 오른쪽에 놓여 다른 데이터를 가리지 않습니다.

### word.py
| 명령 | 주요 옵션 |
|---|---|
| `read` | `--start N`, `--save [경로.md]`. 제목 구조(개요 수준) 포함. 본문 그림은 `[그림]`으로 표시. 범위 고르기(하나만): `--pages 3-5`(Word가 계산한 실제 쪽, 앞 쪽에서 이어지는 단락 포함, 결과에 `page`), `--paragraphs 120-140`(`141-`는 끝까지), `--section "제목 일부"`(그 제목부터 같은 수준 이상의 다음 제목 전까지, 하위 절 포함) |
| `outline` | 제목 구조(목차): 제목마다 `paragraph`(read `--paragraphs`와 같은 번호), `level`, `text`, `page`. 200개씩, `--start N`. 제목이 없으면 count 0 |
| `comment` | `--anchor`/`--text`/`--file`이 있으면 `--action add`가 없어도 메모 달기. `--action list`(기본): 메모마다 `no, author, date, page, target_text`(메모를 단 글), `text, done, reply_to`. 100개씩 `--start`. `--action add --anchor 글 --text 메모`(또는 `--file 메모.txt`): 기준 글이 든 단락을 `write`의 앵커와 같은 규칙으로 고르고(여러 곳이면 `[후보]` + `--occurrence`) 그 글 자체에 메모를 답니다. `--file 메모.csv`(열: 기준 글, 메모 내용[, 몇 번째 단락])는 여러 메모를 한 번에 달며, 기준 글 하나라도 정할 수 없으면 하나도 달지 않습니다. 단 뒤 다시 읽어 확인(`verified`) |
| `revisions` | 변경 내용 추적 목록: `no, type`(삽입/삭제/서식 변경 등), `author, date, page, text`, `track_changes`(추적 켜짐 여부). 100개씩 `--start`. 읽기만 합니다(수락·거부는 Word에서) |
| `info` | 쪽/단어/글자 수, 표, 그림, 제목 목록. `--all`. `--detail`: 그림 크기(쪽 너비 대비)와 제목 수준, 쪽보다 넓은 그림 경고 |
| `write` | `--text`(마크다운 자동 변환, 줄바꿈 `\n`), `--file (.md/.txt, .csv는 표)`, `--image`, `--width`, `--heading N`, `--title`, `--callout`, `--toc`, `--page-break`, `--anchor 글 --pos after\|before\|replace`, `--occurrence N`(앵커가 여러 단락에 있을 때), `--theme`. 앵커가 표 칸 안이면 `after`는 표 뒤에 넣고 `before`/`replace`는 거절(칸 안에 들어가지 않게). 문서가 표로 시작하면 `--pos start` 거절 |
| `replace` | `--find/--replace`(빈 문자열은 삭제) 또는 `--map`, `--match-case`, `--whole-word`, `--bold`, `--color`, `--highlight`. 머리글/바닥글/글상자 포함. 문구당 5,000곳 상한에 닿으면 남은 곳을 확인하지 못해 `UNVERIFIED:`(`limit_reached`). `--match-case`가 없으면 대소문자만 바꾸는 치환(`office` -> `Office`)도 확인됨. `--replace`도 서식 옵션도 없으면 바꿀 것이 없어 STOP_RETRY |
| `export` | `--save 경로.pdf\|.docx`, `--page-number`(사본에만) |

### ppt.py
| 명령 | 주요 옵션 |
|---|---|
| `read` | `--slides 1-3`, `--save`. 길어서 잘리면 고른 슬라이드 중 남은 번호로 이어 읽는 명령(`--slides 1,5,9` -> `5,9`) |
| `info` | 슬라이드별 제목, 도형/표/그림 수. `--all`. `--detail [--slides 2-4]`: 도형별 위치·크기·글자 크기·목록 종류와 배치 경고(슬라이드 밖, 글자 넘침, 겹침) |
| `write` | `--slide N\|new`, `--title`, `--text`(`-` 줄은 글머리표, `1.` 줄은 번호), `--file (.md: 제목마다 새 슬라이드, .csv: 표)`, `--image`, `--cards "제목:내용\|..."`, `--theme`, `--overwrite`. `--title`과 `.md`를 함께 주면 `--title`이 첫 슬라이드 제목(첫 절의 제목은 본문 첫 줄). 기존 슬라이드에 이미 다른 제목이 있으면 아무것도 쓰지 않고 `needs_user_decision`(사용자가 원하면 `--overwrite`, 새 슬라이드는 해당 없음) |
| `replace` | `--find/--replace` 또는 `--map`, `--slide`, 서식 옵션. 표와 그룹 도형 포함 |
| `export` | `--save 경로.pdf\|.pptx`. PDF 기능이 없는 Office(2007 기본 설치 등)는 `.pptx`로 저장하라는 STOP_RETRY를 돌려줍니다 |
| `slide` | `--action list\|add\|duplicate\|move\|delete\|clear --slide --to --layout`. `clear`는 `--slide` 또는 `--all-slides`가 필요 |
| `notes` | 발표자 노트. 읽기: `--slides 1-3`(생략 시 전체, `--slide 3`도 됨), 노트가 있는 슬라이드만 `no, title, notes`. 쓰기: `--slide N --text 노트`(줄바꿈 `\n`) 또는 `--file 노트.md`. 기존 노트가 있으면 `--append`(뒤에 덧붙이기) 또는 `--overwrite`(바꾸기)가 필요하고, 둘 다 없으면 지우지 않고 `needs_user_decision`. 노트 본문 개체틀이 없으면 쓰지 않음. 쓴 뒤 다시 읽어 확인(`verified`, `before`, `after`) |

### outlook.py
- 조회: `list [--folder --unread --limit]`, `search --query [--from --where all|subject|body]`, `read --id [--save-attach --save]`, `current`, `folders`
- 작성: `new --to --subject --body [--cc --attach --body-file --send]`, `reply [--id --all --body]`, `forward [--id --to --body]`
- `--attach`: 여러 개는 `;`로 구분(쉼표도 됨). 값 전체가 있는 파일이면 하나로 봅니다(`보고서,최종.pdf`). `--body`의 글자 `\n`만 줄바꿈으로 바꾸고(Windows 경로 안의 `\n`은 그대로), `--body-file` 내용은 그대로 씁니다.
- `read --save`: 파일 이름에 쓸 수 없는 글자(`RE:`의 `:` 등)는 `_`로 바꾸고, 저장한 파일을 다시 읽어 확인합니다.
- 첨부는 `data\_work\mail_attachments\`에 저장합니다. 같은 이름 첨부는 번호를 붙이고, 본문 삽입 이미지와 실행 파일 형식은 저장하지 않습니다.

### reader.py
용도: (1) 작성 중인 문서(Word/Excel/PPT)가 규정·규격서(PDF/Word)에 맞는지 확인, (2) 여러 문서에서 키워드를 찾아 근거(쪽, 장절, 문항/조항 번호)와 함께 모으기.

| 명령 | 하는 일 |
|---|---|
| `index --folder 폴더 [--depth 3 \| --recursive] [--max-files 300]` | 폴더의 문서를 내용 출력 없이 캐시에 등록. 결과: `count`, `indexed`, `updated`, `unchanged`, `failed`, `failures`(20개까지), `remaining`, `truncated`, `skipped_types`(지원하지 않는 확장자별 개수) |
| `index --file 파일` / `index --all` | 파일 하나 등록 / 등록된 문서 전체를 다시 확인(바뀐 파일, 이전 형식 캐시만 다시 읽음) |
| `read --file 경로 [--pages 1-3 \| --units 57-120] [--format summary] [--save .md\|.csv\|.xlsx]` | 파일 하나 읽기(캐시에 등록). 단위마다 `page`, `section`, `number`, `kind`. `--pages`는 쪽 정보가 있으면(PDF, PPTX, 쪽 경계가 저장된 DOCX/HWPX) 실제 쪽, 없으면 항목 번호(`pages_mode`: `page`/`item`). 한 단위가 약 8,000자보다 길면 그 단위에 `text_truncated`, `text_chars`, 결과에 `truncated: true`, `truncated_reason`과 `--save` 명령(잘린 뒤쪽은 저장 파일에 있음). 엑셀 시트 행 상한(20,000행)에 닿은 문서도 `truncated`(저장하면 `UNVERIFIED:`). 글자가 없는 PDF 쪽은 `pages_without_text`. `--save`는 저장한 파일을 다시 열어 확인(CSV는 행 수까지) |
| `search --query "A AND B" \| "A OR B" [--file 파일또는폴더] [--pages 3-7] [--section 장절일부] [--regex --top --full --max-files] [--save 근거.xlsx\|.csv\|.md]` | 캐시에서 검색 |
| `clear [--file 파일 \| --name 이름일부 \| --folder 폴더 \| --missing \| --all] [--dry-run]` | 캐시 지우기(색인 항목과 캐시 파일). 원본 문서는 그대로이고, 다시 검색하면 그때 새로 읽습니다. `--missing`은 원본이 없어진 문서만, `--all`은 색인에 없는 남은 캐시 파일까지. `--dry-run`은 `planned`와 실제로 지우는 명령(`next_commands[0]`)만 줌. 지운 뒤 다시 읽어 확인(`verified`, `applied`, `remaining_in_cache`) |
| `list [--name 이름일부] [--folder 폴더]` | 캐시된 문서 목록: 파일, 경로, 형식, 단위 수, 쪽 수, 장절 수, `cached_at`, `stale`(원본 변경 또는 이전 형식), `missing`(원본 없음) |

- 지원 형식: PDF, HWP, HWPX, DOCX/DOCM, PPTX/PPTM, XLSX/XLSM, CSV, TXT, MD. 그 밖의 확장자는 "지원하지 않는 형식입니다"(STOP_RETRY). EML은 1.3.0에서 뺐습니다 (Outlook 통합 도구에서 별도로 다룰 예정).
- `--max-files`는 이번 실행에서 **새로 읽는** 파일 수의 상한입니다. 이미 최신인 파일은 세지 않으므로 같은 명령을 다시 실행하면 남은 파일부터 이어서 등록합니다(`next_commands[0]`). 등록은 원본을 바꾸지 않는 읽기 전용 작업입니다.
- DRM으로 암호화된 Office 파일은 라이브러리로 열 수 없으므로, `excel.py`/`word.py`/`ppt.py read --target 경로`로 Office에서 읽으라고 안내합니다.

#### 검색과 캐시
- `--file` 없음: 캐시에 등록된 문서 전체만 검색합니다. 등록된 문서가 없으면 `error_type: not_indexed`(STOP_RETRY).
  원본이 바뀌었거나 이전 파서 형식인 캐시는 그 자리에서 다시 읽고(20개까지, `recovered`로 알림) 검색합니다. 다시 읽지 못한 캐시(원본 없음, 한도 초과)는
  옛 내용으로 검색하고 결과 항목에 `stale: true`, `stale_cache_count`와 `index --all` 명령. 원본이 지워진 문서는 `file_missing: true`.
- 이전 파서 형식의 캐시를 다시 읽지 못하면 검색하지 않고 `skipped`(다시 등록 필요)로 알립니다(1.4.10, 옛 캐시의 키를 채우던 보정 삭제).
- 검색 범위가 불완전하면(등록하지 못한 파일, 읽지 못한 파일, 다시 읽지 못한 캐시, 글자 없는 스캔 PDF, 일부 쪽에 글자가 없는 PDF, 20,000행에서 자른 엑셀, 들어가지 않은 하위 폴더)
  `not_searched` 목록과 함께 `UNVERIFIED:`로 응답합니다. 0건이어도 "근거 없음"으로 단정하지 않게 하기 위해서입니다.
- 검색어: `A AND B`, `A OR B`(대소문자 무관, AND가 먼저 묶임), `"설계 기준"`(따옴표는 구문, 검색어에 넣지 않음), 연산자 없이 이어진 단어는 한 구문.
  띄어쓰기, 줄바꿈, 기호(`# - . : / ( )` 등)는 무시하고 비교합니다: `SOI 1` = `SOI #1` = `SOI-1`, `설계기준` = `설계 기준`, PDF 줄바꿈으로 갈라진 글.
  숫자 경계는 지킵니다: `SOI 1`은 `SOI #10`, `SOI 21`과 맞지 않고, `1.2`는 `11.2`, `3.1.2`와 맞지 않습니다.
  기호를 빼면 숫자만 남는 검색어(`#4`, `1.2`)는 기호까지 그대로 찾습니다. 더 복잡한 패턴은 `--regex`(예: `"SOI\s*#?\d"`). `--regex`의 빈 일치는 일치로 보지 않습니다.
- `--file 파일`: 캐시가 없거나 낡았으면(수정 시각/크기/파서 버전) 먼저 읽어 등록하고, 최신이면 캐시를 씁니다. `--file 폴더`: 폴더(기본 3단계 하위)의 미등록/변경 문서를 `--max-files`까지 등록한 뒤 그 폴더 문서만 검색합니다.
  결과에 `indexed_now`(새로 등록한 파일 이름, 20개까지), `indexed_now_count`, `from_cache`, 남은 파일이 있으면 `not_indexed_remaining`과 `index --folder` 명령.
- 거르기: `--pages 3-7`(쪽 정보가 없는 단위는 빼고 `page_unknown_skipped`로 개수 보고), `--section "SCOPE"`(장절 경로 일부, 대소문자 무시).
- 결과 항목: `file`, `path`, `page`, `page_source`, `section`(목록), `section_text`(` > `로 이음), `number`, `loc`, `heading`, `evidence`, `snippet`, `--full`이면 `text`.
  공통: `total_matches`(= `count`), `matched`(= `shown`), `truncated`, `files_searched`, `output_file`, `next_commands`. 잘렸으면 `--save "근거.xlsx"`로 전체를 저장하는 명령을 줍니다.
- `evidence` 형식: `규격서.pdf p.12, 1. SCOPE > 1.1. Scope (1.1)`. 쪽이 없으면 `loc`(예: `매출.xlsx 매출!12행`), Word 쪽을 쪽 나눔으로만 셌으면 `p.3(추정)`.
- `--save`: `.xlsx`(머리글 굵게, 줄바꿈, 열 너비, 첫 행 고정), `.csv`, `.md`(기본). 열: 파일, 쪽, 장절, 번호, 위치, 내용(문장/단락), 문맥. `.md`에는 항목마다 `근거:` 줄. 저장 파일에는 `--top`과 관계없이 전체(최대 5,000건)를 담습니다. 5,000건을 넘으면 `save_truncated`와 함께 `UNVERIFIED:`. 저장한 파일은 다시 열어 확인합니다(`save_problem`이 있으면 `UNVERIFIED:`).
- `index`: 새로 읽은 문서의 캐시 파일이 실제로 저장됐는지 확인하고, 저장하지 못한 파일이 있으면 `cache_failed`와 함께 `UNVERIFIED:`.

#### 캐시 형식 (모든 형식 공통, `PARSER_VERSION` 4)
문서마다 `cache/docs/<해시>.json` = `{file, type, mtime, size, parser_version, parsed_at, meta, units}`, 색인 `cache/index.json`(경로별 형식, 단위 수, 쪽 수, 장절 수, 수정 시각, 크기, 파서 버전).

| 단위 키 | 뜻 |
|---|---|
| `id` | 1부터 순서 |
| `kind` | `heading`(제목), `paragraph`(단락, 글머리 항목 포함), `row`(표 행, 엑셀/CSV 행), `slide`(슬라이드) |
| `loc` | 사람이 읽는 위치: `p.12`, `p.3 표 1 2행`, `단락 37`, `표 3 5행`, `슬라이드 3`, `매출!12행` |
| `page`, `page_source` | 쪽 번호와 출처(아래 표). 모르면 둘 다 null |
| `section` | 위에서 아래로의 장절 경로. 예: `["III. 개발 내용", "1. 개발 내용 종합", "가. 검증기술"]`. 제목 단위는 자기 자신까지 포함 |
| `number` | 단위 글 머리의 번호 토큰. 예: `1.1.1`, `가`, `III`, `제12조`, `Q3`, `(1)`, `1)`. 끝의 `.`/`#`은 뺌(`1.0#` -> `1.0`). 원문자 입력은 `(1)`로 |
| `heading` | 하위 호환: `section`의 마지막 값(없으면 빈 글) |
| `text` | 본문 |

- 표는 **행 단위**입니다. 첫 행이 머리글로 보이면(짧은 글, 숫자만 있는 칸 없음) 각 행을 `항목: 내전압 | 기준: 1.5 kV 1분 | 시험방법: KS C 1234`처럼 만들어, 기준 값을 검색하면 그 요구사항 행이 머리글과 함께 나옵니다.
  병합/빈 머리글 칸은 앞 칸 머리글을 이어 쓰고, 여러 줄 칸은 공백으로 잇습니다. 행의 `number`는 첫 칸이 번호(`1`, `1.2`, `REQ-001`)일 때. DOCX는 표 단위 대신 행 단위만 둡니다(같은 내용이 두 번 걸리지 않게).
  PDF는 pymupdf `find_tables()`로 선이 있는 표를 찾아 행 단위로 넣고, 표 영역의 글은 단락에서 뺍니다. 표를 못 찾으면 단락으로 처리합니다.
- 이전 형식(`parser_version` 1) 캐시도 검색됩니다: 빈 키를 채웁니다(`p.N`/`슬라이드 N`에서 쪽, `heading`에서 장절). 다시 읽으면(`read`, `index`) 새 형식으로 바뀝니다.

| 형식 | `page` / `page_source` | 장절 |
|---|---|---|
| PDF | 정확한 쪽 / `pdf` (쪽마다 단락·줄 묶음 단위) | 번호 패턴 + 글자 크기/굵게 |
| DOCX | Word가 저장한 쪽 경계 `w:lastRenderedPageBreak`가 있으면 그것으로 / `word_saved_layout`. 없으면 강제 쪽 나눔(`w:br type=page`, `pageBreakBefore`, 다음 쪽 구역 나누기)만 세어 / `page_break`(실제보다 적을 수 있어 근거에 `(추정)`). 둘 다 없으면 null | 제목 스타일(Heading N, 제목 N, 개요 수준) + Word 자동 번호(numbering.xml) + 번호 패턴 |
| PPTX | 슬라이드 번호 / `slide` | `[슬라이드 제목]` |
| HWPX | 문단의 쪽 나눔(`pageBreak`)이 있으면 / `page_break`, 없으면 null | 개요 수준(문단 모양의 `heading type=OUTLINE`)과 개요 번호, 없으면 번호 패턴 |
| HWP, TXT/MD | null | 번호 패턴 (MD는 `#` 수준) |
| XLSX / CSV | null | `[시트 이름]` / 없음 |

#### 제목(장절)과 번호 인식
줄 머리의 번호를 **계열**로 나누고, 계열의 수준은 문서에서 처음 나온 위치로 정합니다.
- 계열: 로마 숫자(`III.`, 전각 로마 숫자), 점 번호(`1.` `1.1` `1.1.` `1.0#` `1.1#`: 숫자 개수가 깊이, 끝의 `.0`은 세지 않음), 가나다(`가.`), `1)`, `(1)`, `가)`, `(가)`, `a)`, `(a)`, `A.`, `a.`, 원문자(입력만), `제N편/장/절/관/조/항/호`, `Chapter/Part/Section/Article N`, `Annex/Appendix N`, `부록/별표/별지/첨부 N`, 문항(`Q3`, `문3`, `문항 3`).
- 수준: 처음 보는 계열은 현재 가장 깊은 제목의 한 단계 아래. 이미 본 계열이 다시 나오면 그 수준으로 돌아가고 더 깊은 제목은 닫습니다. 점 번호는 `기준 수준 + 깊이 - 1`.
  PDF에서는 상위 제목보다 글자가 분명히 큰 제목을 그 하위로 두지 않습니다.
- 글머리 기호(`O`, `o`, `-`, `*`, `·`, 동그라미/네모/세모 기호, 참고 표시 등)는 번호가 아니고(`number` null) 제목도 아닙니다. 지금 장절 아래의 단락입니다.
- 제목 조건(오탐 방지): 한 줄, 80자 이하(굵게/큰 글씨는 120자), 번호 뒤에 글자가 있음, 문장 끝(`.`, `다`, `니다`, 긴 줄의 `~함/~임/~음/~됨`)으로 끝나지 않음. 문장 속 숫자, 날짜(`2026. 10. 28.`), 수치(`3.5kg`)는 번호가 아닙니다.
  구분 기호 없는 숫자(`2 REQUIREMENTS`)는 PDF에서 본문보다 크거나 굵을 때만 제목입니다. Word 제목 스타일은 이 조건과 관계없이 제목입니다.
  `제N조(제목) 본문...`처럼 긴 조문은 `제N조(제목)`만 장절에 넣습니다. 번호가 있지만 제목이 아닌 단락(`1. 시스템은 ...한다.`)도 `number`는 채웁니다.
- 예: 아래 세 형식 모두 올바르게 중첩됩니다 (`tests/test_reader_evidence.py`).
  ```text
  III. 개발 내용            -> [III. 개발 내용]
  1. 개발 내용 종합         -> [III. 개발 내용, 1. 개발 내용 종합]
  가. 검증기술              -> [..., 가. 검증기술]
    O 프로그램 전체 연구 목표  -> 단락 (number null), 장절은 위와 같음
  1. SCOPE / 1.1. Scope / 1.1.1. Data / 1.2. Purpose   -> 1.2는 [1. SCOPE, 1.2. Purpose]
  1.0# Introduction / 1.1# Purpose / 1.1.1# Information -> 1.1.1#은 3단계
  ```
- 한 문서는 번호 체계 하나를 쓴다고 가정합니다. 같은 글자 크기로 서로 다른 체계를 섞은 문서는 먼저 나온 체계 아래로 들어갈 수 있습니다.

### office.py
- `windows`: 열린 Office 문서 목록과 실행 중인 Office 버전, PDF 내보내기 가능 여부(`office`). PDF 기능이 없는 Office에서는 `export --save *.pdf`가 시도 없이 바로 안내합니다
- `focus --app excel|word|ppt [--name]`: 창을 앞으로 가져옵니다. 최대화 상태는 유지합니다.
- `transfer`: 아래 4절 참고
- `undo [--app] [--list] [--pick N] [--overwrite]`

## 4. 앱 간 전달 (transfer)과 회사 DRM

| 전달 | 명령 | 방식 |
|---|---|---|
| 엑셀 차트 -> 워드/PPT | `--from excel --chart 이름 --to word,ppt` | 그림 (아래 3방식) |
| 엑셀 범위 -> 워드/PPT | `--from excel --range A1:D10 --to word` | 네이티브 표 (값 전달, 클립보드 없음) |
| 워드/PPT -> 엑셀 | `--from word --what text\|table [--table N] --to excel` | 새 시트에 표로 |

차트 그림은 다음 순서로 자동 시도하고, 결과 JSON의 `method`에 실제로 쓴 방식을 적습니다.
1. `export`: Excel의 `Chart.Export` PNG를 씁니다. **파일 앞부분이 PNG가 아니면(DRM 암호화) 건너뜁니다.**
2. `redraw`: 차트의 계열 값을 COM으로 읽고 matplotlib으로 다시 그립니다. Python이 쓴 파일을 쓰고 클립보드는 쓰지 않습니다. 원본 서식은 근사치입니다.
3. `relay`: 엑셀 -> 워드 -> PPT 클립보드 경유입니다.

회사 PC에서 처음 한 번 `python office.py transfer --probe`를 실행하면, 세 방식을 실제로 시험해 대상 앱별로 되는 순서를 `config.json`에 저장합니다.

## 5. 체이닝 예

```bat
:: 사인 2주기 그래프 -> 워드/PPT
python excel.py write --range A1:B101 --header "x,sin(x)" --formula "=(ROW()-2)*4*PI()/99" "=SIN(A2)"
python excel.py chart --range A1:B101 --type line --title "사인 2주기"
python office.py transfer --from excel --chart "사인 2주기" --to word,ppt

:: 워드/PPT 내용 -> 엑셀
python office.py transfer --from word --what text --to excel     :: 새 시트에 받음 (--sheet 이름, 기존 시트에 덮어쓰기는 --overwrite)
python office.py transfer --from ppt --what text --to excel

:: 문서 등록 -> 근거 검색 -> 엑셀로 저장 / 워드에 넣기
python reader.py index --folder "C:\docs\규격"
python reader.py search --query "내전압 AND 기준" --save "근거.xlsx"
python reader.py search --query "마일스톤" --file "C:\docs\기획" --save
python word.py write --file "<결과의 output_file>"
```

## 6. 테스트

```bat
python -m pytest -q CLI-dev/office-cli          :: 오프라인 (가짜 COM + 실제 샘플 파일)
python CLI-dev/office-cli/tools/run_win_tests.py  :: Windows + Office 실기 (자세한 내용: TEST_CHECKLIST.md)
```
- 실기 테스트는 테스트 폴더에 `report_*.txt`(요약)와 `errors_*.txt`(분석용 상세 로그)를 남깁니다. 문제가 있으면 `errors_*.txt`를 그대로 Claude에게 전달하세요.
  `errors_*.txt`에는 러너의 자동 복구(닫힌 A 문서 다시 열기), 창/프로세스 정리 기록(어떤 PID에 Quit/WM_CLOSE를 보냈고 무엇을 보호했는지), 도구 오류 로그가 함께 들어 있습니다.
- 러너는 테스트 문서(A.xlsx/A.docx/A.pptx)가 열린 프로세스를 끝내지 않습니다. 정리 판단 규칙은 `tools/win_test_rules.py`(오프라인 테스트 `tests/test_win_runner.py`).
- DRM PC에서 python-pptx 기본 템플릿(`pptx/templates/default.pptx`)이 암호화돼 있으면 그 템플릿이 필요한 오프라인 테스트만 "python-pptx 기본 템플릿이 손상됨 (DRM 추정)"으로 건너뜁니다. python-docx는 폴더 템플릿으로 대신 만듭니다. 도구 코드는 기본 템플릿을 쓰지 않습니다.
- 도구는 `C:\CLI\office-cli`에 둡니다 (SKILL.md의 명령 경로 기준). 실기 테스트 파일은 `C:\CLI-dev`(이 저장소의 `CLI-dev` 폴더)에 둡니다. OpenCode 에이전트에게 실기 테스트를 맡기려면 스킬 폴더 `office-cli-test`에 `C:\CLI-dev\office-cli\tools\office-cli-test\SKILL.md`를 복사합니다. 에이전트는 `--agent`로 실행해 결과 JSON 1줄만 읽습니다.
- 대상은 Office 16.0(2016/2019/2021/365)입니다. Word 2007의 `SaveAs2` 부재는 `SaveAs`로 대체하므로 Office 2007에서도 실기 테스트가 돕니다. 다만 PDF 추가 기능이 없으면 PDF 내보내기 항목은 `ENV-SKIP`으로 기록됩니다.
