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
├── outlook.py   메일   list / search / read / current / folders / attachments / new / reply / forward
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

- 관리자 권한(높은 무결성 수준)으로 실행하면 일반 권한으로 연 Office 창·문서에 COM으로 닿지 못합니다(회사 PC에서 관리자 PowerShell의 OpenCode).
  이때 "열린 문서 없음"으로 끝내지 않고(1.4.40): `active` 대상, 열린 문서 이름 대상, `office.py transfer`의 받는 문서를 찾지 못하면
  `error_type: "elevated_process"`와 "관리자가 아닌 일반 PowerShell 창에서 OpenCode를 다시 실행" 안내(STOP_RETRY, 다시 실행하지 말 것).
  `info --all`·`office.py windows`의 0건은 확인한 0건이 아니므로 `UNVERIFIED:`와 같은 안내. Outlook 연결 실패 message에도 붙입니다.
  판단은 `ctypes.windll.shell32.IsUserAnAdmin()`(`_runtime.is_elevated`, Windows 밖이나 실패하면 일반 권한으로 봄).

### 화면과 안전
- **읽기는 화면을 움직이지 않습니다.** 활성 시트, 스크롤, PPT 슬라이드가 그대로입니다.
- **쓰기는 고친 시트/슬라이드로만 이동합니다.** `--no-focus`를 주면 이동하지 않습니다.
- **쓰기 전 스냅샷을 자동으로 만듭니다**(문서별 5개 보관, 50MB 초과 문서는 생략). 같은 경로에 다른 문서를 새로 만들면 이전 문서의 스냅샷은 되돌리기 대상에서 빠집니다(파일 속 만든 시각으로 판별).
  - 원본 문서는 저장하거나 경로를 바꾸지 않습니다.
  - 쓰기 전에 멈춘 실행(입력 오류, `needs_user_decision`)은 만든 되돌리기 지점을 지웁니다. 남기면 다음 `undo`가 바뀐 것 없는 그 지점으로 되돌려 실제로 바꾼 작업이 그대로 남았습니다(1.4.12).
    지우는 것은 문서를 바꾸지 않았을 때만입니다(1.4.13): 열린 문서는 바꾼 표시가 없고 저장 상태가 그대로일 때(원래 저장 안 된 문서는 입력 오류일 때만),
    도구가 열거나 만든 파일은 디스크 파일이 그대로일 때(오류나 저장 실패로 저장하지 않고 닫음). 그래서 쓰기 명령은 입력 확인(시트, 칸 주소, 수식, `--highlight` 등)을
    모두 마친 뒤에 문서를 바꿉니다(`excel write --new-sheet`도 쓰기 직전에 시트를 만듦).
  - 바꾼 것 없이 성공한 실행(0곳 바꿈, `chart --action delete --all`에 차트 0개, 이미 있는 시트 `sheet add`, `sheet select`, 빈 슬라이드 `slide clear`)도
    같은 기준으로 지점을 지웁니다(1.4.14, 이전에는 남아 다음 `undo`가 바뀐 것 없는 이 지점으로 되돌리고 `VERIFIED`로 답해 실제로 바꾼 작업이 그대로 남음).
    문서를 바꾸는 명령은 `t.set_focus()`(바꾸지 않았으면 `changed=False`)나 `t.changed = True`로 표시합니다.
    1.4.15부터 같은 글로 바꾸는 `replace`(`--map {"A": "A"}` 포함, 서식 없음), 같은 자리로 `slide move`·`range copy/move`, 같은 노트로 `notes --overwrite`,
    이미 그 상태인 `range`(병합 없는 `unmerge`, 고정 없는 `unfreeze`, 같은 `width`/`height`,
    사용 범위 밖이거나 값이 없는 칸의 `clear --type contents`)도 바꾸지 않고 "0곳/0칸 바뀜" 또는 "바뀐 것 없음"(`unchanged: true`)으로 답하며 지점을 남기지 않습니다.
    이미 그 상태인지는 COM 속성(`MergeCells`, `FreezePanes`, `ColumnWidth`, `Intersect(UsedRange)` 등)으로 확인하고, 알 수 없으면 바꾼 것으로 봅니다.
    1.4.16부터 `hide`/`unhide`(일부만 숨긴 범위에서 `Hidden`이 섞임 대신 False일 수 있음)와 `clear` 전체·서식(사용 범위 밖 칸에도
    데이터 유효성·메모가 있을 수 있음)은 확인하지 않고 늘 바꾼 것으로 봅니다(되돌리기 지점이 하나 더 남을 수 있음).
  - 도구가 연 닫힌 파일은 명령이 바꾼 것(`t.changed`)이 있을 때만 저장합니다(1.4.15). 휘발성 수식(NOW, RAND 등)이 든 파일은 열기만 해도
    저장 안 된 상태라, 이전에는 바꾼 것 없는 성공도 사용자 파일을 다시 저장하고 바뀐 것 없는 지점을 남겼습니다.
  - 오래된 지점 정리(5개 유지)는 명령이 문서를 바꾼 뒤에만 합니다. 거절된 실행은 옛 지점을 밀어내지 않습니다(1.4.13).
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
    - `--new-sheet`가 있는 명령은 `excel write`뿐입니다. 그 밖의 명령(`replace`, `style`, `range`, `chart` 등)의 없는 시트는 읽기와 같은 "찾을 수 없음"(`[추천]`/`[후보]`)이고 `--new-sheet`를 안내하지 않습니다(1.4.12, 이전에는 그 명령에 없는 옵션을 붙이라고 안내).
  - 워드 `--anchor`가 여러 단락에 있으면 `[후보]`(번호와 단락 앞부분)로 오류. `--occurrence N`으로 고릅니다. 찾지 못하면 비슷한 문구를 안내합니다 (공백·대소문자 차이 등 확실할 때만 `[추천]`, `--pos replace`에서는 더 엄격).
  - `office.py transfer --to excel --sheet 이름`은 새 시트 이름입니다(transfer는 새 시트에 받습니다. 방금 만든 새 파일만 예외, 아래). 이미 있으면 덮어쓰지 않고 비어 있는 이름을 `[추천]`, 덮어쓰기는 `--overwrite`.
    방금 만든 새 파일이면 기본 시트 이름(`Sheet1` 등)은 새 시트를 만들지 않고 그 빈 기본 시트에 받습니다(1.4.16은 기본 시트가 1개일 때만, 1.4.17부터 여럿이어도).
  - 쓰기 `--target`에 파일 이름만 주었는데 열린 문서도 파일도 아니면 새 파일을 만들지 않고 "찾을 수 없음"(열린 문서가 있으면 `[후보]`). 새 파일은 전체 경로를 줄 때만 만듭니다(1.4.10, 이전에는 현재 폴더에 만들어 거기에 씀).
    전체 경로인지는 `%USERPROFILE%`·`~`를 푼 뒤에 판단하고(`%USERPROFILE%\Desktop\새.xlsx`도 새로 만듦), 없는 전체 경로는 열린 문서를 `[후보]`로 주지 않습니다(1.4.11).
- **메일**: 기본은 작성 창만 띄웁니다. `--send`는 명시했을 때만 쓰고, 수신자 주소를 확인한 뒤 발송합니다.

### 파일과 폴더
| 항목 | 기본 위치 | 바꾸는 방법 |
|---|---|---|
| data 폴더 (파일명만 쓴 입력/출력) | `<도구폴더>\data` | 환경변수 `OFFICE_CLI_DATA_DIR` 또는 `config.json`의 `data_dir` |
| 작업 폴더 (`--save` 경로 생략 시 `<이름>_<시각>`, 같은 초에 겹치면 `_2`를 붙이고 `recovered`, 7일 후 자동 정리. 이름의 쓸 수 없는 글자 `\ / : * ? " < > \|`는 `_`, 1.4.23) | `data\_work` | - |
| 캐시 (등록한 문서 색인, 스냅샷) | `<도구폴더>\cache` | `OFFICE_CLI_CACHE_DIR` 또는 `cache_dir` |
| 설정 (`transfer --probe` 결과 등) | `<도구폴더>\config.json` | 환경변수 `OFFICE_CLI_CONFIG` (1.4.33) |
| 오류 로그 (실패한 실행 자동 기록) | `data\logs\errors.txt` | data 폴더를 따라감 |
| 메일 첨부 (`outlook.py attachments --save`) | `<다운로드>\메일 첨부\<시각>_<제목>` (자동 정리 안 함, En-Genius 실행도 같음, 1.4.34) | `--save "<전체경로 폴더>"` |

- En-Genius 스킬(보관: `archive/engenius`)은 `C:\CLI\office-cli`의 도구를 실행하므로 OpenCode와 같은 `data`, `cache`, `config.json`을 씁니다.
  도구 파일 자체가 트레이의 임시 폴더(`%TEMP%\engenius-skill\<id>`)에서 실행될 때만(실행마다 지워짐) 기본 위치를
  `%USERPROFILE%\Documents\En-Genius\CLI-data\office-cli\`로 바꿉니다 (1.4.33, 1.4.35부터 이 경우만). 환경변수가 있으면 그것이 먼저입니다.

- 입력 파일은 절대경로 -> data 폴더 -> 작업 폴더 순서로 찾습니다. 없으면 비슷한 파일명을 `[추천]`으로 알려 줍니다.
- 텍스트 입력은 UTF-8, UTF-16, CP949를 자동으로 판별합니다. CSV 출력은 기본 utf-8-sig입니다(Excel에서 한글 정상).
- 결과가 길면 기본 200행 또는 약 8,000자에서 자르고 `truncated: true`와 이어 읽기 명령을 줍니다.
  엑셀 칸 값과 찾은 값은 하나에 500자(`value_truncated`), 슬라이드 글·워드 단락·PPT 노트는 하나에 약 8,000자(`text_truncated`, `text_chars`)에서 자릅니다.
  잘린 것이 있으면 `truncated: true`, `truncated_reason`과 `--save` 명령(저장 파일에는 전체, PPT 노트는 저장 명령이 없어 PowerPoint에서 확인)을 줍니다(1.4.12, 이전에는 첫 행·첫 슬라이드가 아무리 커도 그대로 넣음).
- 도구가 쓴 CSV는 다시 읽을 때 `,`로 읽습니다(구분자 추측이 `| ` 같은 칸 때문에 `|`를 골라 행 수가 틀려 `UNVERIFIED`가 되던 문제, 1.4.12).

### 오류 로그
- 실패한 실행(exit 1, 인자 오류 exit 2)은 자동으로 `<data 폴더>\logs\errors.txt`에 덧붙습니다. `C:\CLI`에 배치했다면 `C:\CLI\office-cli\data\logs\errors.txt`(`config.json`의 `data_dir`나 `OFFICE_CLI_DATA_DIR`를 바꿨으면 그 폴더 아래).
- 항목: 시각, 도구 버전, Python, OS / 실행한 명령과 인자 / `error_type`과 message / 재시도 한도를 넘긴 복구 시도(`복구 시도:`) / 내부 오류(`ERROR:`)일 때만 traceback. 문서 본문은 기록하지 않습니다.
- 내용 옵션의 값은 `<N자 생략>`으로 남깁니다: `--body --text --subject --to --cc --bcc --from --query --find --replace --value --values --formula --header --title --cards --anchor --search --new-name --map --section`, `excel.py`의 `--highlight`(셀 값이 들어감, 1.4.26)(`office.py`의 `--from/--to`, `ppt.py`의 `--to`는 앱 이름·위치, `word.py`의 `--highlight`는 색 이름이라 그대로). 경로와 그 밖의 인자는 그대로 남깁니다(1.4.10).
  짧은 이름(`-q` = `--query`, `-q값`)과 줄여 쓴 이름(`--que`, `--bod=값`)도 가립니다(1.4.11).
  `결과:` 줄의 message는 문서 글을 뺀 글입니다(stdout JSON의 message는 그대로, 1.4.27·1.4.28).
  1. 문서에서 가져온 글(칸 값, 슬라이드 제목·노트, 시트·문서·파일·메일 폴더 이름 목록, 후보 단락·제목, 비슷한 문구, 메모 파일의 기준 글)을
     message에 넣는 오류는 그 글을 뺀 `log_message`를 함께 만들고(`InputError`/`UserDecisionError`의 `log_message`, `ambiguous()`·`not_found()`는
     후보를 개수만 남김, 결과 dict는 `_log_message` 키), 로그에는 그것을 씁니다. 뺀 자리는 `<문서 내용 생략>`, `[후보 N개 생략]` 등(1.4.28).
  2. 그 위에 가린 인자 값을 같은 `<N자 생략>`으로 바꿉니다. 10자 이상인 값은 10자 이상의 앞부분(잘린 미리보기)도 가리고,
     1~2자 값(`--value 1`)은 따옴표 안에 있을 때만 가립니다(`Sheet1`의 1을 가리지 않게. 따옴표 없이 쓴 1~2자 값은 로그에 남을 수 있음, 1.4.28).
  3. 빠뜨린 곳을 위한 대비로 표식(`_runtime.LOG_CUT_MARKERS`: `[후보]`, `[추천]`, `후보:`, `추천:`, `비슷한 문구:`, `사용 가능한 <종류>:`(종류에 공백 가능))
     부터 끝까지를 `<문서에서 가져온 후보·문구 N자 생략>`으로 자릅니다.
  4. 내부 오류의 traceback도 가린 인자 값을 바꿉니다(표식 자르기는 하지 않음, 1.4.28). 가리기에 실패해도 로그 기록은 오류를 삼킵니다.
  - 남는 것: 가리지 않는 인자로 준 이름(`--sheet`, `--target`, `--name` 등)과 경로·열린 문서 하나의 파일 이름, Office COM 오류 설명, 파서 예외 글(손상 파일)은 그대로 남습니다.
  - 모든 파서는 줄여 쓴 옵션을 받지 않습니다(`allow_abbrev=False`, 인자 오류 exit 2).
- 성공, 사용자 중단(exit 130)은 기록하지 않습니다. 2MB를 넘으면 `errors.old.txt`로 넘기고 새로 시작합니다. 로그 경로는 stdout JSON에 넣지 않습니다(사람과 개발 에이전트의 사후 분석용).
- 실기 테스트 러너의 `errors_<시각>.txt`는 테스트 한 번의 판정 근거를 모은 별도 파일이고, 그 실행 중에 쌓인 도구 오류 로그(`<테스트 폴더>\_data\logs\errors.txt`)도 끝에 함께 옮겨 담습니다.

## 3. 명령 요약

### excel.py
| 명령 | 주요 옵션 |
|---|---|
| `read` | `--range`, `--sheet`, `--search 글자`(셀 주소 반환), `--formulas`(수식 칸 목록, 아래), `--where "금액 >= 100; 부서 = 영업 | 부서 = 개발"`(조건에 맞는 행만, `;` 그리고, `|` 또는, 아래), `--max-rows`, `--save [경로.csv]`(저장한 CSV를 다시 읽어 빈 행을 뺀 행 수 확인). `--target`이 CSV/TSV 파일이면 Excel 없이 읽음(인코딩 자동 판별, 숫자 변환, `source: file(csv)`). 닫힌 xlsx도 열린 문서와 같게 읽음: 범위를 생략하면 데이터가 시작하는 칸부터(UsedRange처럼), 거꾸로 준 범위(`C5:A1`)는 `A1:C5`, 크기 정보가 없는 파일도 실제 데이터 끝까지(시트를 훑는 것은 범위 생략·`B:B`·`5:7`일 때만) |
| `info` | 시트별 사용 범위와 차트. `--all`: 열린 통합문서 전체. `--detail`: 차트 종류·계열 수·위치와 데이터를 가리는 차트 경고 |
| `write` | `--range` 시작 셀 또는 범위. 입력은 `--value`, `--values JSON`, `--file (.csv/.md 표/.txt)`(`.md`는 word·ppt와 같은 마크다운 해석으로 첫 번째 표만: 코드 블록 안 `|` 줄과 두 번째 표는 빼고 정렬 줄 `|--|`는 빼며 `\|`는 칸 글자 `|`. 표가 없으면 한 줄 = 한 행, 1.4.26), `--formula F1 [F2 ...]` 중 하나. `--header "a,b"`, `--append`(쓸 열 전체에서 가장 아래 행 다음에, 시작 행의 머리글과 같으면 생략. `--formula`의 상대 참조는 실제로 쓴 행 기준으로 옮기고 `$` 참조는 그대로, 옮기지 못하면 쓰지 않고 STOP_RETRY, 1.4.24), `--sheet`, `--new-sheet`(없는 시트를 새로 만들 때만). 쓸 칸에 값이나 수식(빈 글을 내는 수식 포함)이 있으면 아무것도 쓰지 않고 `needs_user_decision`(사용자가 원하면 `--overwrite`). `--append`도 쓸 칸을 확인하고, `--formula`/`--value`도 쓸 열 수만큼 가장 아래 행을 찾습니다. 모든 검사를 마친 뒤에 씁니다(실패하면 시트는 그대로). 쓴 뒤 칸을 다시 읽어 비교합니다(같으면 `VERIFIED:`, 다르면 `UNVERIFIED:` + 다른 칸, 1.4.6). 숫자만 `--formula`로 주면 `--values`를 `[추천]`. Excel이 다른 값으로 바꿔 버릴 글자(앞자리 0 코드 `01234`, 15자리 초과 숫자, `1-2`·`3/4`, `(100)`, `1E5`)는 글자 그대로 씁니다 |
| `replace` | `--find/--replace` 또는 `--map JSON`, `--whole`, `--all-sheets`, 서식 옵션. 수식 셀은 내용 보존. `--replace`도 서식 옵션도 없으면 바꿀 것이 없어 STOP_RETRY. 찾은 칸이 있어도 바꿀 글이 같고 서식이 없으면(수식 칸 포함) 바꾼 칸으로 세지 않음(`changed_cells` 0, `unchanged_same_text`, 1.4.15). `--range '요약'!A1:D9`처럼 시트를 주면 그 시트에서 읽고 씀. `--all-sheets`와 시트 이름이 든 `--range`는 함께 쓸 수 없음. 찾을 글이 빈 문자열(`--find ""`, `--map` 키 `""`)이면 STOP_RETRY(빈 칸 채우기는 `write`, 1.4.17) |
| `style` | `--theme navy\|slate\|teal\|emerald`, `--bg`, `--color`, `--bold`, `--align`, `--number-format`, `--border`, `--highlight ">=100:pink"`(조건부 서식. 값의 `"`와 `포함:` 규칙의 `~ ? *`는 글자 그대로 찾음, 1.4.24), `--clear` |
| `export` | `--save 경로.csv\|.xlsx\|.pdf`, `--columns A,C,E` |
| `sheet` | `--action list\|add\|rename\|copy\|delete\|select --name --new-name`. 엑셀에서 쓸 수 없는 이름(`/ \ ? * [ ] :`, 31자 초과, 빈 이름, 앞뒤 `'`, 엑셀이 예약한 `History`(대소문자 무관, 1.4.15))은 아무것도 바꾸기 전에 STOP_RETRY와 `[추천]`(쓸 수 없는 글자를 `_`로 바꾸고 31자로 자른 이름, `History`는 `History_`, 있는 시트와 겹치면 번호, 1.4.14). `write --new-sheet`, `office.py transfer --sheet`도 같음(`transfer`는 받는 문서의 시트 이름을 안 뒤에 확인해 `[추천]`이 있는 시트와 겹치지 않음, 새로 만들 파일은 만들기 전에 확인, 1.4.15). `copy`의 기본 이름은 `<시트 28자> 복사` |
| `range` | `--action autofit\|width\|height\|merge\|unmerge\|freeze\|unfreeze\|hide\|unhide\|clear\|copy\|move\|border\|insert\|delete\|sort`(insert/delete/sort는 아래, 1.4.37). `clear`는 `--range` 필수(시트 전체를 실수로 지우지 않게). `merge`는 왼쪽 위 밖의 칸에 값이 하나라도 있으면(병합하면 지워짐), `copy`/`move`는 대상에 값이 있으면 `needs_user_decision`(사용자가 원하면 `--overwrite`). `copy`는 대상의 수식(R1C1)을 원본과 비교해 확인(상대 참조는 값이 달라도 같은 수식). `autofit`은 비교할 값이 없어 `UNVERIFIED:`. `hide`/`unhide`/`autofit`의 `--axis`를 생략하면 `5:7` 같은 행 주소는 행, 그 밖은 열. 이미 그 상태면(병합 없는 `unmerge`, 고정 없는 `unfreeze`, 값이 없는 범위의 `clear --type contents`, 같은 자리로 `copy`/`move` 등) 바꾸지 않고 "바뀐 것 없음"(`unchanged: true`, 1.4.15). `hide`/`unhide`와 `clear`(전체·서식)는 엑셀이 일부만 숨김·사용 범위 밖 유효성 검사를 알려 주지 않을 수 있어 늘 바꾼 것으로 봅니다(1.4.16) |
| `chart` | `--action create\|list\|delete`, `--range`, `--type column\|bar\|line\|line_markers\|pie\|area\|scatter\|scatter_lines\|combo`, `--title`, `--at`, `--labels`, `--legend`, `--theme` |

- 수식이 1개면 범위 전체에 넣고 상대참조가 자동으로 조정됩니다. 여러 개면 열마다 하나씩 넣습니다.
- 첫 열이 숫자인 `line` 차트는 x축이 숫자 축이 되도록 `scatter_lines`로 그립니다(예: 사인 그래프).
- `--at`을 생략하면 차트는 시트의 데이터(사용 범위) 오른쪽에 놓여 다른 데이터를 가리지 않습니다.
- 칸 주소는 엑셀 범위(1~1,048,576행, A~XFD열) 안이어야 합니다. `A0`, `0:3`, `XFE1`은 STOP_RETRY(1.4.12, 이전에는 닫힌 파일에서 그대로 읽음).
- 닫힌 xlsx에서 범위를 직접 주면(`A1:C5`) 데이터 뒤의 빈 행까지 그 범위만큼 `rows`를 셉니다(열린 문서와 같음, 1.4.12). 범위를 생략하면 끝쪽 빈 행은 뺍니다.
  단 계산된 값이 없는 수식 칸이 있는 행·열(값 없이 수식만 있는 합계 행 등)은 빈 칸으로 읽혀도 빼지 않습니다(1.4.32, 아래).
  준 범위가 200만 칸을 넘으면(`A1:XFD1048576`) 데이터가 있는 끝까지만 읽고 `recovered`로 알리며, 그래도 넘으면 STOP_RETRY와 `[추천] --range`(1.4.13).
  파일에 적힌 크기(`<dimension>`)가 서식만 있는 먼 칸(예: `Z300000`)까지라 그것으로도 넘으면 값을 훑어 실제로 값이 있는 끝을 찾습니다(1.4.14, 이전에는 STOP_RETRY).
- 읽기 미리보기의 잘림 표시: 500자보다 긴 값은 "긴 값 N칸 잘림", 첫 행이 표시 상한(약 8,000자)을 넘어 뺀 칸은 "뒤쪽 N열 생략"으로 따로 알립니다(1.4.13).
- 닫힌 xlsx의 계산된 값이 없는 수식 칸(openpyxl 같은 프로그램이 쓰고 엑셀에서 한 번도 열지 않은 파일)은 값이 계산된 적이 없어 빈 칸으로 읽힙니다.
  읽은 범위에 그런 칸이 있으면 `read`(`--search`, `--save` 포함), `export --save .csv`, `office.py transfer --from excel --range`가 `UNVERIFIED:`와
  `uncomputed_formulas`(칸 주소 20개까지), `uncomputed_count`, 20개를 넘으면 `uncomputed_truncated: true`로 답하고 엑셀에서 파일을 열어 저장한 뒤
  다시 실행하라고 알립니다(1.4.31, 이전에는 빈 값으로 `VERIFIED`). 빈 칸이 있거나 아래 `fullCalcOnLoad` 표시가 있을 때만 시트 XML의 수식 칸을 확인합니다.
  - 범위를 데이터로 정할 때(범위 생략, `B:B`·`5:7`, 200만 칸을 넘는 범위 줄이기)는 끝쪽 행·열을 자르기 전에 시트 전체의 그런 칸을 찾아
    그 행·열을 남깁니다(1.4.32, 이전에는 `=SUM` 합계 행이 빠지고 `VERIFIED`, 수식만 있는 시트는 머리글만, 크기 정보가 없는 파일은 끝쪽 수식 열도 빠짐).
  - `xl/workbook.xml`의 `<calcPr fullCalcOnLoad="1">`(또는 `"true"`)은 "열 때 모두 다시 계산" 표시입니다. xlsxwriter(`<v>0</v>`을 씀)·openpyxl 같은
    프로그램이 쓰며, 그 파일의 수식 칸은 값이 있어도(0 등) 모두 계산된 값이 없는 칸으로 봅니다(1.4.32). 엑셀은 저장할 때 이 표시를 지우므로
    엑셀에서 저장한 파일은 해당 없습니다.
  - 값이 있는 끝에서 1,000행(열) 넘게 떨어진 계산된 값이 없는 수식 칸(데이터 A1:C10과 `Z200000`)은 범위를 데이터로 정할 때 넣지 않습니다(1.4.44).
    전에는 범위 생략 읽기가 그 칸까지 빈 행을 채웠고, `-r A:Z`는 "너무 큽니다"와 그 칸을 빼는 `[추천]`을 냈습니다. 이제 데이터 범위만 읽고,
    요청한 범위 안의 먼 칸은 `far_uncomputed_formulas`(20개까지), `far_uncomputed_count`, `UNVERIFIED:`와 그 칸만 읽는 명령(`next_commands` 끝,
    `--range 시트!Z200000`)으로 알립니다(`read`, `export --save .csv`). 요청한 열 밖이면(`-r A:C`) 알리지 않습니다.
    칸 위치는 10만 개까지 보고, 넘으면 전처럼 모두 범위에 넣습니다.
  - 배열 수식(`<f t="array" ref="C1:C3">`)은 파일에 왼쪽 위 칸에만 `<f>`가 있지만 `ref` 범위의 모든 칸을 수식 칸으로 셉니다(1.4.43). 다른 칸에
    계산된 값이 없거나(openpyxl은 왼쪽 위 칸만 씀) `fullCalcOnLoad`이면 계산된 값이 없는 칸입니다(전에는 왼쪽 위 칸만 셈).
  - 수식 칸 확인 자체가 실패하면(파일 손상, MemoryError 등) 0개로 보지 않고 `uncomputed_check: "unknown"`과
    `UNVERIFIED: ... 수식 계산 여부를 확인하지 못함`으로 답합니다(1.4.32, 이전에는 `VERIFIED`). reader는 `meta.uncomputed_formulas.unknown`.
    reader 캐시에 남은 확인 실패는 결과로 다시 쓰지 않고, 캐시를 쓸 때(`read`, 색인 검색) 원본을 다시 확인해 확인되면 캐시도 고칩니다(1.4.42,
    이전에는 일시적 실패(MemoryError 등)가 파일이 바뀌거나 캐시를 지울 때까지 계속 `UNVERIFIED:`).
- `read --search`: `count`(찾은 칸 전체 수, `total_matches`와 같음), `shown`, `truncated`. 0건이면 `VERIFIED: 0건.` 안내(1.4.31).
- `read --formulas`(1.4.36): 값은 그대로 두고, 읽은 범위의 수식 칸만 `formulas`에 더합니다.
  - 항목: `{"cell": "B5", "formula": "=SUM(B2:B4)", "value": 600}`. 읽기는 한 시트라 `cell`에 시트 이름을 붙이지 않습니다(시트는 `sheet`).
    배열 수식은 배열마다 한 번(배열 안에서 처음 만난 칸, 보통 왼쪽 위) `"array": true`, `"array_range": "C1:C3"`. 닫힌 파일의 데이터 표(What-If)는
    `=TABLE(행 입력,열 입력)`과 `data_table: true`. 500자보다 긴 수식·값은 앞부분만(`value_truncated`).
  - `formula_count`(전체 수), `formulas_truncated`. 200개를 넘으면 앞 200개만 담고 전체 목록 CSV(`cell, formula, value, array`)를
    `<데이터 폴더>\results\excel_formulas_<시트>_<시각>.csv`에 써 `full_result_file`로 줍니다(R2). `--save`와 함께면 저장한 값 CSV 옆
    `<이름>_formulas.csv`에 늘 전체 목록을 쓰고(`full_result_file`) 다시 읽어 행 수를 확인합니다(R1, 맞지 않으면 `UNVERIFIED:`).
  - 수식이 없으면 `formula_count: 0`, `formulas: []`와 "수식 0개: ...에는 수식이 없습니다" 안내. CSV는 수식을 저장하지 않아 늘 0개(오류 아님).
  - 열린 문서(COM): 범위의 `HasFormula`가 False면 수식을 읽지 않습니다. 수식 글은 범위 전체를 한 번에 읽습니다: `Formula2`(엑셀 365, 동적 배열을
    그대로)가 없거나 오류면 `Formula`(영어 함수 이름, A1 주소). 2차원 결과에서 상수는 값, 수식은 `=`로 시작하는 글이라 칸마다 부르지 않고,
    값도 같은 글인 칸(`'=`로 둔 글자 상수일 수 있음)만 `HasFormula`로 확인합니다. 배열 수식은 범위의 `HasArray`가 False가 아닐 때만 칸마다
    확인합니다(2,000칸까지, 넘으면 `warnings`).
  - 닫힌 xlsx/xlsm: 값을 읽은 뒤 같은 범위만 openpyxl `data_only=False`(읽기 전용)로 한 번 더 읽습니다. 수식 칸은 openpyxl이 수식으로 읽은 칸
    (`data_type` `f`: `=`로 시작하는 글, `ArrayFormula`, `DataTableFormula`)만이라 `=`로 시작하는 글자 상수는 빠집니다. 공유 수식은 openpyxl이
    칸마다 펼칩니다. `value`는 파일에 저장된 계산된 값이고, 계산되지 않은 수식 칸은 `null`이며 계산되지 않은 수식 안내(`UNVERIFIED:`,
    `uncomputed_formulas`, 1.4.31/1.4.32)는 그대로입니다. 배열 수식의 왼쪽 위 칸이 읽은 범위 밖이면 그 배열은 값만 나옵니다(파일에 수식이 왼쪽 위에만 있음).
  - Office 없이 수식을 읽을 수 없는 닫힌 파일(.xls, .xlsb, zip이 아닌 xlsx)은 Office(pywin32)가 없으면 `STOP_RETRY:`로 엑셀에서 연 뒤 다시
    실행하라고 알립니다. Office가 있으면 다른 읽기처럼 엑셀로 열어 읽습니다.
  - `--search`와 함께면 수식 글에서도 찾고(대소문자 무시), 찾은 수식 칸에 `formula`를 붙입니다.
  - 확인(R1)은 다른 읽기와 같습니다: 읽은 결과가 확인이라 `VERIFIED:`. 수식 글은 문서 내용이라 오류 로그에 남기지 않습니다(수식 읽기 실패는 예외 종류만).
- `read --where "<열> <연산자> <값>; ... | ..."`(1.4.37, 또는 1.4.45): 읽은 범위의 첫 행을 머리글로 보고, 조건에 맞는 행만 돌려줍니다. 열린 문서, 닫힌 xlsx, CSV 모두.
  - `;`는 그리고(AND), `|`는 또는(OR)이고 `;`가 먼저 묶입니다: `부서 = 영업; 금액 > 250 | 부서 = 기획`은 (영업이고 250 초과) 또는 기획. 두 묶음에 맞는 행도 한 번만.
  - 열은 머리글 이름(대소문자·공백·`_-` 무시) 또는 열 문자(`B`). 이름을 찾지 못하면 머리글 이름 `[후보]`(오류 로그에는 남기지 않음).
  - 연산자 `>= <= > < = != ~`(`~`는 포함). 숫자 칸과 숫자 값은 숫자로(`1,000`도 숫자), 날짜 칸은 `YYYY-MM-DD`(시각 포함 가능) 값과 날짜로,
    그 밖은 앞뒤 공백을 뺀 글자로 대소문자 없이 비교합니다. 크기 비교인데 칸이 숫자가 아니면(빈 칸, 글자) 맞지 않습니다. `금액 = `은 빈 칸.
    값의 앞뒤 따옴표는 뺍니다. 값에 `;`나 `|`가 있으면 따옴표로 감쌉니다(`메모 = 'a;b'`, PowerShell에서는 작은따옴표). 따옴표 안은 나누거나 고치지 않습니다(1.4.45).
  - 결과: `values`(첫 행은 머리글, 다음은 맞는 행), `row_numbers`(`values`와 같은 순서의 시트 행 번호, 첫 값은 머리글 행), `count`(맞는 행 수),
    `data_rows`(머리글 밖 행 수), `where`. 0건이면 `count: 0`, `values`는 머리글만, `VERIFIED: 0건.`(R2).
  - 잘림과 `--save`는 `read`와 같습니다: `--max-rows`를 넘으면 `truncated`와 `--where`가 든 저장 명령(`next_commands`), `--save`는 머리글과 맞는 행만 CSV
    (행 번호 열은 넣지 않음)로 쓰고 다시 읽어 확인.
  - 고칠 수 있는 표기(`==`, `=>`, `=<`, `<>`, `AND`, `&&`, `OR`, `||`, `또는`, `contains`, `포함`)는 STOP_RETRY와 고친 식 `[추천]`. 읽을 수 없는 식은 형식과 예를 알리고
    `[추천]` 없이 STOP_RETRY. `--search`, `--formulas`와 함께 쓸 수 없습니다.
- `range --action insert|delete --range 5:7|B:C`(1.4.37): 행·열 넣기·지우기. `5`, `B`처럼 하나도 됩니다. `A5:C7` 같은 칸 범위는 행인지 열인지
  `--axis rows|cols`가 필요합니다(없으면 `[후보]`).
  - `insert`: 범위 앞에 같은 수만큼 넣고 아래(오른쪽)는 밀립니다.
  - `delete`: 지울 행·열에 값이나 수식이 있는 칸이 있으면 지우지 않고 `needs_user_decision`(칸 수와 주소 5개까지, 값은 stdout과 오류 로그 어디에도
    보이지 않음). 사용자가 원하면 `--overwrite`. `--dry-run`은 바꾸지 않고 칸 수(`filled_cells`), 주소 20개(`filled_addresses`)와
    실행 명령(`next_commands[0]`, 값이 있으면 `--overwrite` 포함)을 줍니다(되돌리기 지점도 남기지 않음).
  - `delete`는 다른 칸의 수식이 지울 행·열 안에만 있는 칸이나 범위를 참조하면(지우면 `#REF!`) 값이 없는 행·열이어도 `needs_user_decision`(1.4.41).
    결과 `broken_refs`(수), `broken_ref_cells`(`시트!주소` 20개까지, 넘으면 `broken_ref_cells_truncated`), message에는 5개까지. 수식 글은 보이지 않습니다.
    같은 통합문서의 다른 시트에서 이 시트를 이름으로 참조하는 수식도 봅니다(시트 이름 대소문자 무시). 범위 일부만 지우면(`SUM(C2:C5)`에서 3행) 범위가
    줄어들 뿐이라 세지 않고, 지울 행·열 안의 수식, 따옴표 글, 다른 통합문서 참조, 이름 정의·`INDIRECT` 같은 간접 참조는 세지 않습니다.
    통합문서 전체가 200만 칸을 넘으면 넘는 시트는 훑지 않고 `warnings`로 알립니다. `--overwrite`로 지우면 preview에 `#REF!`가 된 수식 수.
  - 다른 쓰기처럼 되돌리기 지점을 남기고(`office.py undo`), 닫힌 파일은 엑셀로 열어 고칩니다.
  - 확인(R1): 사용 범위 끝 행(열)이 넣은·지운 수만큼 바뀌었는지, 옮겨진 칸 하나(수식이 아닌 값이 있는 첫 칸, `moved`)가 기대한 자리에 같은 값인지.
- `range --action sort --range A1:E100 --by "<열>" [--desc] [--header auto|yes|no]`(1.4.37): 행 단위 정렬. 기준 3개까지는 Range.Sort, 4개부터 64개까지는 `Worksheet.Sort.SortFields`(엑셀 2007 이상, 정렬한 뒤 시트에 남는 정렬 기준은 지움, 1.4.46).
  - `--by`는 머리글 이름 또는 열 문자, 여러 기준은 쉼표(`--by "부서,금액"`). `--desc`만 쓰면 모든 기준이 내림차순, `--desc "금액"`이면 그 기준만.
  - 머리글: `--by`가 머리글 이름이면 첫 행이 머리글. 열 문자면 `--header auto`(기본)가 판단합니다: 첫 행 칸이 모두 글자이고, 굵게이거나 아래에 숫자·날짜가
    있는 열이 있으면 머리글. 결과 `header`, `header_decided_by`.
  - 열 전체(`A:E`)는 사용 범위까지만. 병합된 칸이 있으면 STOP_RETRY. 범위 바로 옆 열의 같은 행에 값이 있으면(표를 가로지름) 정렬하지 않고
    `needs_user_decision`과 표 전체 범위. 이미 정렬되어 있으면 바꾸지 않음(`unchanged: true`, 되돌리기 지점 없음).
  - 확인(R1): 다시 읽어 행 수, 머리글 행, 행 묶음(정렬 전과 같은 행들)과 기준 순서(숫자 < 글자 < 논리값, 빈 칸은 맨 뒤, 글자는 대소문자 무시)를 봅니다.
    엑셀 비교 규칙과 달라 순서를 확인하지 못하면 `UNVERIFIED:`.

### word.py
| 명령 | 주요 옵션 |
|---|---|
| `read` | `--start N`, `--save [경로.md]`(저장한 파일을 다시 열어 확인). 제목 구조(개요 수준) 포함. 본문 그림은 `[그림]`으로 표시. 범위 고르기(하나만): `--pages 3-5`(Word가 계산한 실제 쪽, 앞 쪽에서 이어지는 단락 포함, 결과에 `page`), `--paragraphs 120-140`(`141-`는 끝까지), `--section "제목 일부"`(그 제목부터 같은 수준 이상의 다음 제목 전까지, 하위 절 포함) |
| `outline` | 제목 구조(목차): 제목마다 `paragraph`(read `--paragraphs`와 같은 번호), `level`, `text`, `page`. 200개씩, `--start N`. 제목이 없으면 count 0 |
| `comment` | `--anchor`/`--text`/`--file`이 있으면 `--action add`가 없어도 메모 달기. `--action list`(기본): 메모마다 `no, author, date, page, target_text`(메모를 단 글), `text, done, reply_to`. 100개씩 `--start`. `--action add --anchor 글 --text 메모`(또는 `--file 메모.txt`): 기준 글이 든 단락을 `write`의 앵커와 같은 규칙으로 고르고(여러 곳이면 `[후보]` + `--occurrence`) 그 글 자체에 메모를 답니다. `--file 메모.csv`(열: 기준 글, 메모 내용[, 몇 번째 단락])는 여러 메모를 한 번에 달며, 기준 글 하나라도 정할 수 없으면 하나도 달지 않습니다. 단 뒤 다시 읽어 확인(`verified`) |
| `revisions` | 변경 내용 추적 목록: `no, type`(삽입/삭제/서식 변경 등), `author, date, page, text`, `track_changes`(추적 켜짐 여부). 100개씩 `--start`. 읽기만 합니다(수락·거부는 Word에서) |
| `info` | 쪽/단어/글자 수, 표, 그림, 제목 목록. `--all`. `--detail`: 그림 크기(쪽 너비 대비)와 제목 수준, 쪽보다 넓은 그림 경고 |
| `write` | `--text`(마크다운 자동 변환, 줄바꿈 `\n`), `--file (.md/.txt, .csv는 표)`, `--image`, `--width`, `--heading N`, `--title`, `--callout`, `--toc`, `--page-break`, `--anchor 글 --pos after\|before\|replace`, `--occurrence N`(앵커가 여러 단락에 있을 때), `--theme`. 앵커가 표 칸 안이면 `after`는 표 뒤에 넣고 `before`/`replace`는 거절(칸 안에 들어가지 않게). 문서가 표로 시작하면 `--pos start` 거절 |
| `replace` | `--find/--replace`(빈 문자열은 삭제) 또는 `--map`, `--match-case`, `--whole-word`, `--bold`, `--color`, `--highlight`. 머리글/바닥글/글상자 포함. 문구당 5,000곳 상한에 닿으면 남은 곳을 확인하지 못해 `UNVERIFIED:`(`limit_reached`). `--match-case`가 없으면 대소문자만 바꾸는 치환(`office` -> `Office`)도 확인됨. `--replace`도 서식 옵션도 없으면 바꿀 것이 없어 STOP_RETRY. 바꿀 글이 찾는 글과 같고 서식이 없으면 바꾸지 않고 `replaced` 0, 찾은 곳 수는 `unchanged_same_text`(1.4.15). 대소문자만 다른 치환(`iphone` -> `iPhone`)에서 이미 바꿀 글 그대로인 곳은 바꾸지 않고 `unchanged_same_text`로 셈(1.4.17). 찾을 글이 빈 문자열이면 STOP_RETRY(1.4.17) |
| `export` | `--save 경로.pdf\|.docx`, `--page-number`(사본에만) |

마크다운(`word write`의 `.md`·`--text`, `ppt write`의 `.md`) 해석(1.4.24, 코드 줄의 `**`·`__`를 굵게로 읽지 않음 1.4.25): 펜스 코드 블록(백틱 3개 또는 `~~~`, 닫지 않으면 끝까지)은 줄마다 단락 하나로
글자 그대로 넣고(들여쓰기 유지, 빈 줄 생략, 안의 `#`·`-`·`1.`을 제목·목록으로 읽지 않음, 펜스 줄은 넣지 않음).
`ppt write --text`는 마크다운 블록 해석 없이 줄마다 한 줄(`-`/`1.` 목록, `**` 굵게 표시 제거)이고, 펜스 안 코드 줄만 위와 같이 글자 그대로입니다(1.4.26).
다시 읽기 확인은 넣은 첫 글(구분선은 건너뜀, 코드 줄·표 칸은 글자 그대로)로 비교합니다(1.4.26). 표의 정렬 줄은
머리글 다음 줄이면 `-` 하나부터(`:-`, `-:`), 칸 안의 `\|`는 글자 `|`입니다.

구분선과 YAML 머리말(1.4.31):
- 줄 하나가 `---`(`***`, `___`)이면 구분선입니다. word는 아래 테두리만 있는 빈 단락을 넣고(`inserted`의 `구분선`), 페이지 나누기가 아닙니다
  (1.4.30까지는 페이지 나누기). 페이지를 나누려면 `word.py write --page-break`를 따로 실행합니다. ppt `.md`는 구분선을 넣지 않고 슬라이드도
  나누지 않습니다(슬라이드는 `#`·`##` 제목마다). `ppt write --text`의 `---` 줄은 글자 그대로 한 줄입니다.
- 1번째 줄이 `---`이고 `key: 값` 줄들 뒤에 `---`(또는 `...`) 줄로 닫히면 YAML 머리말로 보고 넣지 않으며 `warnings`에 남깁니다(word `.md`·`--text`,
  ppt `.md`). 닫는 줄이 없거나 사이에 `key:` 모양이 아닌 줄이 있으면 머리말이 아니고 첫 `---`는 구분선입니다.
- word 굵게 위치는 Word처럼 UTF-16 단위로 셉니다(앞에 이모지 같은 BMP 밖 글자가 있어도 굵게가 밀리지 않음, 1.4.31).

굵게 표시와 목록 번호(1.4.27):
- 본문의 `**x**`, `__x__`는 굵은 x. `__`는 CommonMark처럼 앞에 글자·숫자가 없을 때만 열고 뒤에 없을 때만 닫습니다(`report__2024__final.xlsx`는 그대로,
  `__init__ 메서드`는 굵은 `init`. 코드 이름은 코드 블록이나 `--callout`으로 넣으면 글자 그대로). `**`는 단어 안에서도 굵게입니다.
- 제목(word 모든 수준, `--heading`, `--title`)과 PowerPoint 슬라이드 제목(`.md`의 `#`·`##`, `ppt write --title`, 3단계 이하 제목 줄)은 같은 규칙으로
  굵게 표시를 뺀 글만 넣습니다(제목은 굵기가 섞인 글을 담지 않음).
- 번호 목록 기호는 1~3자리 숫자(`1.`, `2)`)만 봅니다. `2026. 11. 3. 회의`는 본문입니다. 1이 아닌 번호로 시작하는 목록(`3. 셋`)은 자동 번호가 1부터
  다시 매기므로 번호를 글로 둔 일반 단락(PowerPoint는 목록 기호 없는 줄)으로 넣습니다. 빈 줄로 띄운 `1.` / `2.`는 한 목록입니다.

### ppt.py
| 명령 | 주요 옵션 |
|---|---|
| `read` | `--slides 1-3`, `--save`(저장한 파일을 다시 열어 확인). 길어서 잘리면 고른 슬라이드 중 남은 번호로 이어 읽는 명령(`--slides 1,5,9` -> `5,9`) |
| `info` | 슬라이드별 제목, 도형/표/그림 수. `--all`. `--detail [--slides 2-4]`: 도형별 위치·크기·글자 크기·목록 종류와 배치 경고(슬라이드 밖, 글자 넘침, 겹침) |
| `write` | `--slide N\|new`, `--title`, `--text`(`-` 줄은 글머리표, `1.` 줄은 번호, 펜스 안 코드 줄은 글자 그대로 1.4.26), `--file (.md: 제목마다 새 슬라이드, .csv: 표)`, `--image`, `--cards "제목:내용\|..."`, `--theme`, `--overwrite`. 카드·그림은 앞 내용 아래에 80pt도 남지 않으면 새 슬라이드에 넣고 `warnings`로 알림, 이번에 넣은 도형이 슬라이드 밖으로 나가면 `UNVERIFIED:`(1.4.24). `--title`과 `.md`를 함께 주면 `--title`이 첫 슬라이드 제목(첫 절의 제목은 본문 첫 줄). 첫 절의 제목이 `--title`과 같으면(공백·대소문자 무시) 본문에 다시 넣지 않음. 기존 슬라이드의 제목 개체틀에 이미 다른 제목이 있으면 아무것도 쓰지 않고 `needs_user_decision`(사용자가 원하면 `--overwrite`, 새 슬라이드와 제목 개체틀이 없는 슬라이드는 해당 없음) |
| `replace` | `--find/--replace` 또는 `--map`, `--slide`, 서식 옵션. 표와 그룹 도형 포함. `--replace`도 서식 옵션도 없으면 STOP_RETRY, 바꿀 글이 찾는 글과 같고 서식이 없으면 바꾸지 않고 `replaced` 0(`unchanged_same_text`, 1.4.15). 찾은 곳마다 비교해 이미 바꿀 글 그대로(대소문자까지)인 곳은 `unchanged_same_text`, 실제로 바뀐 곳만 `replaced`(1.4.17). 찾을 글이 빈 문자열이면 STOP_RETRY(1.4.17) |
| `export` | `--save 경로.pdf\|.pptx`. PDF 기능이 없는 Office(2007 기본 설치 등)는 `.pptx`로 저장하라는 STOP_RETRY를 돌려줍니다 |
| `slide` | `--action list\|add\|duplicate\|move\|delete\|clear --slide --to --layout`. `clear`는 `--slide` 또는 `--all-slides`가 필요. 같은 자리로 `move`(`--slide 2 --to 2`)는 바꾸지 않고 "바뀐 것 없음"(1.4.15). `list`는 `count`, 0장이면 `VERIFIED: 0건.`(1.4.27) |
| `notes` | 발표자 노트. 읽기: `--slides 1-3`(생략 시 전체, `--slide 3`도 됨), 노트가 있는 슬라이드만 `no, title, notes`. 쓰기: `--slide N --text 노트`(줄바꿈 `\n`) 또는 `--file 노트.md`. 기존 노트가 있으면 `--append`(뒤에 덧붙이기) 또는 `--overwrite`(바꾸기)가 필요하고, 둘 다 없으면 지우지 않고 `needs_user_decision`. 노트 본문 개체틀이 없으면 쓰지 않음. 쓴 뒤 다시 읽어 확인(`verified`, `before`, `after`). 이미 같은 노트면 쓰지 않고 `unchanged: true`(1.4.15) |

### outlook.py
- 조회: `list [--folder --unread --limit --days --since]`, `search --query [--from --where all|subject|body --days --since]`, `read --id [--save]`, `current [--save]`, `folders`
- 받은 날짜(1.4.37): `--days N`은 오늘 포함 최근 N일(1 = 오늘), `--since YYYY-MM-DD`는 그 날짜 0시부터. 함께 쓰면 STOP_RETRY. 결과에 `period`
  (`since`, `until`, `days`)와 preview의 기간. `list`는 Jet `[ReceivedTime] >= 'YYYY-MM-DD HH:MM'`(현지 시각, 월/일 순서가 지역 설정에 따라
  뒤바뀌지 않는 형식), `search`는 DASL `urn:schemas:httpmail:datereceived`(UTC로 바꾼 시각)를 같은 `@SQL=`에 AND로 붙이고, 받은 시각을 메일마다
  다시 확인합니다(내림차순이라 기간 밖 메일을 만나면 멈춤). Outlook이 날짜 조건을 거부하면 날짜 없는 필터로 거른 뒤 받은 시각으로 거르고
  `recovered`. `search`는 `--days`/`--since`만으로도 됩니다. 잘못된 날짜는 STOP_RETRY: 고칠 수 있는 표기(`2026/10/1`, `20261001`, `2026년 10월 1일`)와
  `--days 0`/`-7`은 `[추천]`, 없는 날짜·미래 날짜·읽을 수 없는 글은 `[추천]` 없이 예만.
- 첨부: `attachments [--id] [--save [폴더]] [--name 이름일부] [--index 1,3]` (1.4.34)
- `--folder`의 폴더를 찾지 못하면 그 단계의 하위 폴더 이름(끝 이름)끼리 비교해 `[추천]`(표기만 다름)이나 `[후보]`를 주고, 보여 줄 때만
  전체 경로(`계정\폴더\하위`)로 바꿉니다. 1.4.28까지는 전체 경로끼리 비교해 같은 계정 이름 때문에 관계없는 폴더가 모두 `[후보]`였습니다.
  오류 로그에는 계정 이름(보통 메일 주소) 대신 `<메일 계정>`을 남깁니다(1.4.29). message 끝의 `(경로: ...)`도 로그에서는
  계정 이름 자리를 `<메일 계정>`으로 바꿉니다(1.4.30, `not_found`의 `log_extra`. 이전에는 `--folder` 원문이 그대로 남았음).
- `--limit`은 1~200(기본 10). 200을 넘으면 200통으로 줄이고 `warnings`로 알립니다. 0 이하는 STOP_RETRY(1.4.23).
- `list`, `search`의 `count`는 `--limit` 안에서 가져온 메일 수입니다. 0통이면 `count: 0`과 `VERIFIED: 0건.` 안내(조건을 바꿔 다시 실행하지 않음, 1.4.27).
- `search`: 검색어의 `%`, `_`는 글자 그대로 찾습니다. DASL `LIKE`에서는 와일드카드라 더 넓게 찾으므로, 그런 검색어는 제목·본문·보낸사람을
  다시 읽어 글자 그대로 들어 있는 메일만 남깁니다(대소문자 무시, 읽지 못한 칸은 남김). 1.4.22까지는 `_`를 공백으로 바꾸고 `%`를 지워
  `견적서_v2`, `10% 할인`이 0건이었습니다(1.4.23).
- 작성: `new --to --subject --body [--cc --attach --body-file --send]`, `reply [--id --all --body --attach]`, `forward [--id --to --body --attach]`
- `--attach`: 여러 개는 `;`로 구분(쉼표도 됨). 값 전체가 있는 파일이면 하나로 봅니다(`보고서,최종.pdf`). 나눈 조각이 있는 파일이 아니면 뒤 조각을 이어 붙여 봅니다(폴더 이름에 `;`가 든 경로 여러 개). `--body`의 글자 `\n`만 줄바꿈으로 바꾸고(Windows 경로 안의 `\n`은 그대로), `--body-file` 내용은 그대로 씁니다.
- `read --save`, `current --save`: 파일 이름에 쓸 수 없는 글자(`RE:`의 `:` 등)는 `_`로 바꾸고, 저장한 파일을 다시 읽어 확인합니다(확인하지 못하면 `UNVERIFIED:`).
- `read`/`current`는 첨부를 저장하지 않습니다. 첨부가 있으면 `attachment_names`(본문 삽입 이미지를 뺀 이름, 20개까지,
  `attachment_names_truncated`), 본문 삽입 이미지 수 `inline_images`, `next_commands`의 "첨부 내려받기"
  (`attachments --id <id> --save`)를 줍니다. `--save`와 함께면 "메일 내용을 워드에 넣기"가 먼저입니다. `--save-attach`는 1.4.34에서 없앴습니다.
- `attachments` (1.4.34): `--id`를 빼면 `reply`처럼 열린 창이나 목록에서 선택한 메일입니다. 다음 명령에는 그 메일의 id를 넣습니다.
  - `--save` 없이: 목록 `attachments`(항목마다 `index` 1부터, `name`, `size` 바이트(모르면 null), `hidden_inline` 본문 삽입 이미지, `blocked` 실행 파일 형식),
    `count`, `truncated`(50개 상한). 0개면 `count: 0`과 `VERIFIED: 0건. 이 메일에는 첨부 없음`.
  - `--save`: 기본 폴더는 `<다운로드>\메일 첨부\<YYYYMMDD_HHMMSS>_<제목 40자>`입니다. 다운로드는 Windows 셸 폴더 설정(옮긴 경우 그 위치),
    없으면 `%USERPROFILE%\Downloads`. data 폴더와 관계없고 자동 정리하지 않습니다.
    `--save "<폴더>"`는 전체 경로만 받고(없으면 만듦), 상대 경로는 STOP_RETRY(사용자에게 전체 경로를 물음).
  - 고르기: `--name`(이름 일부, 대소문자 무시, `;`로 여러 개, 옵션 반복 가능), `--index 1,3`(`2-4`도 됨). 둘을 함께 주면 합칩니다.
    맞는 첨부가 없는 값이 하나라도 있으면 아무것도 저장하지 않고 첨부 이름을 `[후보]`로 STOP_RETRY. `--save` 없이 고르기만 주면 STOP_RETRY.
  - 실행 파일 형식(`.exe .bat .cmd .js .vbs .ps1 .msi .lnk` 등)은 골라도 저장하지 않습니다. 본문 삽입 이미지(서명 그림 등)는
    전부 저장할 때는 건너뛰고 `--name`/`--index`로 직접 고르면 저장합니다. 건너뛴 첨부는 `skipped`(`name`, `reason`).
  - 덮어쓰지 않습니다. 같은 이름 파일이 있으면(폴더에 있거나 같은 메일에 같은 이름 첨부) ` (2)`를 붙이고 `recovered`로 알립니다.
    파일 이름에 쓸 수 없는 글자는 `_`(`rt.safe_filename`).
  - 결과: `saved`(전체 경로), `skipped`, `count`(저장한 수), `folder`(저장한 것이 없으면 null). 저장한 것이 0개면 `VERIFIED: 0건.`과 이유.
  - 확인(R1): 저장한 파일이 있는지 다시 보고, Outlook이 크기를 알려 준 첨부는 빈 파일이 아닌지 봅니다. Outlook의 `Size`에는 메일 속성 몫이 더해져
    파일보다 조금 크므로 같은 크기는 요구하지 않습니다. 저장 실패, 저장 뒤 파일 없음, 빈 파일이 있으면 `UNVERIFIED:`와 그 첨부 이름.
- `new/reply/forward --attach`: 본문 파일·첨부 파일을 Outlook에 연결하기 전에 확인합니다(없으면 아무것도 만들지 않고 STOP_RETRY, 1.4.13).
  전체 경로가 아닌 파일 이름은 data 폴더에서 찾고(없으면 data 폴더의 비슷한 이름 `[추천]`/`[후보]`), 전체 경로 파일이 없으면 같은 폴더의
  비슷한 파일을 전체 경로로 `[후보]`(폴더가 없으면 그렇게 알림, 1.4.34). 붙인 뒤 `Attachments.Count`가 붙인 수만큼 늘었는지 다시 셉니다
  (전달은 원본 첨부가 그대로 있고 늘어난 수로 비교). 맞지 않거나 붙이다 실패하면 작성 창만 띄우고 `UNVERIFIED:`, `--send`였으면
  보내지 않고 작성 창을 띄운 뒤 STOP_RETRY(`attach_problems`, 1.4.34).

### reader.py
용도: (1) 작성 중인 문서(Word/Excel/PPT)가 규정·규격서(PDF/Word)에 맞는지 확인, (2) 여러 문서에서 키워드를 찾아 근거(쪽, 장절, 문항/조항 번호)와 함께 모으기.

| 명령 | 하는 일 |
|---|---|
| `index --folder 폴더 [--depth 3 \| --recursive] [--max-files 300]` | 폴더의 문서를 내용 출력 없이 캐시에 등록. 결과: `count`, `indexed`, `updated`, `unchanged`, `failed`, `failures`(20개까지), `remaining`, `truncated`, `skipped_types`(지원하지 않는 확장자별 개수) |
| `index --file 파일` / `index --all` | 파일 하나 등록 / 등록된 문서 전체를 다시 확인(바뀐 파일, 이전 형식 캐시만 다시 읽음). `--all`은 색인에서 빠진 캐시 파일 중 원본이 있는 문서도 색인에 다시 넣습니다(`readded_to_index`, `recovered`) |
| `read --file 경로 [--pages 1-3 \| --units 57-120] [--format summary] [--save .md\|.csv\|.xlsx]` | 파일 하나 읽기(캐시에 등록). 단위마다 `page`, `section`, `number`, `kind`. `--pages`는 쪽 정보가 있으면(PDF, PPTX, 쪽 경계가 저장된 DOCX/HWPX) 실제 쪽, 없으면 항목 번호(`pages_mode`: `page`/`item`). 한 단위가 약 8,000자보다 길면 그 단위에 `text_truncated`, `text_chars`, 결과에 `truncated: true`, `truncated_reason`(늘 목록)과 `.md`로 저장하는 명령(잘린 뒤쪽은 저장 파일에 있음). 엑셀 시트 행 상한(20,000행)에 닿은 문서도 `truncated`지만 저장해도 그 뒤는 없으므로 저장 명령을 주지 않고 상한을 알림(저장하면 `UNVERIFIED:`). 읽으려는 범위에 글자 없이 그림만 있는 PDF 쪽(스캔 쪽)이 있으면 `pages_without_text`와 `notes`(잘림이 아님). `--save`는 저장한 파일을 다시 열어 확인(CSV는 행 수까지). `.xlsx`는 셀당 32,000자에서 자르므로 잘린 행이 있으면 `UNVERIFIED:`(`.md`는 전체) |
| `search --query "A AND B" \| "A OR B" [--file 파일또는폴더] [--pages 3-7] [--section 장절일부] [--regex --top --full --max-files] [--save 근거.xlsx\|.csv\|.md]` | 캐시에서 검색 |
| `clear [--file 파일 \| --name 이름일부 \| --folder 폴더 \| --missing \| --all] [--dry-run]` | 캐시 지우기(색인 항목과 캐시 파일, `--file`/`--folder`의 `~`·환경 변수는 풀어서 비교). 원본 문서는 그대로이고, 다시 검색하면 그때 새로 읽습니다. `--missing`은 원본이 없어진 문서만, `--all`은 색인에 없는 남은 캐시 파일까지. `--dry-run`은 `planned`와 실제로 지우는 명령(`next_commands[0]`)만 줌(20개를 넘으면 잘리고 전체 목록은 `full_result_file` CSV, 1.4.26). 지운 뒤 다시 읽어 확인(`verified`, `applied`, `remaining_in_cache`). `applied`와 완료 CSV(`clear_done_...`, 20개를 넘을 때)는 지운 뒤 만들고, CSV의 `removed` 열이 실제로 지웠는지(1.4.27) |
| `list [--name 이름일부] [--folder 폴더]` | 캐시된 문서 목록: 파일, 경로, 형식, 단위 수, 쪽 수, 장절 수, `cached_at`, `stale`(원본 변경 또는 이전 형식), `missing`(원본 없음) |

- 손상된 파일(깨진 XML, 압축 해제 실패, OLE 구조 오류)은 내부 오류가 아니라 STOP_RETRY "손상된 파일이라 읽을 수 없습니다"(1.4.23).
  FileHeader가 없는 OLE 파일(이름만 `.hwp`인 `.doc`/`.xls` 등)은 STOP_RETRY "HWP 5.0 문서가 아닙니다"(1.4.24).
- `--pages`/`--units`(ppt `--slides`)의 범위는 문서 크기로 먼저 줄입니다(`1-999999999`도 바로 끝남, 1.4.24).
- 지원 형식: PDF, HWP, HWPX, DOCX/DOCM, PPTX/PPTM, XLSX/XLSM, CSV, TXT, MD. 그 밖의 확장자는 "지원하지 않는 형식입니다"(STOP_RETRY). EML은 1.3.0에서 뺐습니다 (Outlook 통합 도구에서 별도로 다룰 예정).
- `--max-files`는 이번 실행에서 **새로 읽는** 파일 수의 상한입니다. 이미 최신인 파일은 세지 않으므로 같은 명령을 다시 실행하면 남은 파일부터 이어서 등록합니다(`next_commands[0]`). 등록은 원본을 바꾸지 않는 읽기 전용 작업입니다.
- DRM으로 암호화된 Office 파일은 라이브러리로 열 수 없으므로, `excel.py`/`word.py`/`ppt.py read --target 경로`로 Office에서 읽으라고 안내합니다.
- DOCX 글상자(`w:txbxContent`)의 글은 그 글상자가 든 단락(표) 바로 뒤에 `글상자 N`(여러 단락이면 `글상자 N 단락 M`) 단위로 넣습니다. Word 2010 이후
  파일은 같은 글상자를 DrawingML과 VML(`mc:Fallback`) 두 벌로 저장하므로 `mc:Fallback` 쪽은 읽지 않습니다(한 번만). 단락 번호(`단락 N`)는 글상자가
  없는 문서와 같습니다. 닫힌 파일의 `word.py read`도 같습니다(1.4.31, 이전에는 글상자 글을 읽지 않아 검색이 `VERIFIED` 0건).
- XLSX의 계산된 값이 없는 수식 칸(엑셀에서 연 적 없는 파일)은 `meta.uncomputed_formulas`에 남기고 `read`는 `UNVERIFIED:`(`uncomputed_formulas`),
  `search`는 `not_searched`에 넣어 `UNVERIFIED:`입니다(1.4.31). 파서 버전 9(1.4.31)의 캐시는 이 두 가지를 반영해 다시 읽습니다.

#### 검색과 캐시
- `--file` 없음: 캐시에 등록된 문서 전체만 검색합니다. 등록된 문서가 없으면 `error_type: not_indexed`(STOP_RETRY).
  원본이 바뀌었거나 이전 파서 형식인 캐시는 그 자리에서 다시 읽고(20개까지, `recovered`로 알림) 검색합니다. 원본이 바뀌었는데 다시 읽지 못한 캐시(한도 초과, 읽기 실패)는
  옛 내용으로 검색하고 결과 항목에 `stale: true`, `stale_cache_count`와 `index --all` 명령. 원본이 지워진 문서는 `file_missing: true`.
- 이전 파서 형식의 캐시는 단위 키가 달라 다시 읽지 못하면 검색하지 않고 `skipped`로만 알립니다(`stale`에는 세지 않음). 이유는 셋입니다(1.4.11):
  검색 때 다시 읽는 상한 20개를 넘음 / 다시 읽기 실패 -> `index --all` 명령, 원본이 없어 다시 읽을 수 없음 -> `clear --missing --dry-run` 명령(`index --all`로는 다시 읽을 수 없어 같은 안내가 반복됨).
  캐시 파일이 없거나 손상됐는데 원본도 없는 문서도 같습니다(`missing_file_count`에 세고 `clear --missing --dry-run`, 1.4.12).
  할 명령은 `not_searched`와 `UNVERIFIED:` directive 문구에도 넣습니다.
- 검색 범위가 불완전하면(등록하지 못한 파일, 읽지 못한 파일, 다시 읽지 못한 캐시, 글자 없는 스캔 PDF, `--pages`/`--section` 안에 글자 없이 그림만 있는 쪽이 있는 PDF, 20,000행에서 자른 엑셀, 들어가지 않은 하위 폴더)
  `not_searched` 목록과 함께 `UNVERIFIED:`로 응답합니다. 0건이어도 "근거 없음"으로 단정하지 않게 하기 위해서입니다.
  PDF의 "글자 없는 쪽"(`pages_without_text`)은 그림이 쪽 면적의 40% 이상을 덮고, 머리글·쪽 번호를 뺀 뒤 남은 글이 40자보다 적은 쪽입니다
  (스캔 위에 찍은 도장, 한 쪽짜리 스캔의 Bates 번호처럼 짧은 글만 남은 쪽 포함, 1.4.13). 모든 쪽이 그런 쪽이면 `ocr_needed`.
  빈 쪽·간지(쪽마다 있는 작은 로고만 있는 쪽 포함)는 해당 없고, 그림 위에 보이지 않는 글자 층(OCR 결과)이 있는 쪽도 해당 없습니다.
  `read`는 `ocr_needed` 문서나 고른 쪽이 스캔 쪽뿐이면 `UNVERIFIED`(스캔 이미지라 OCR 필요)로 답합니다(1.4.13, 이전에는 "읽을 글이 없습니다" `VERIFIED`).
  `--save`도 같고, 도장·쪽 번호만 남은 스캔 쪽(`--pages 3`)도 고른 쪽이 모두 스캔 쪽이면 `UNVERIFIED`입니다. 그때와 0개 저장이면 다른 문서에 넣는 `next_commands`를 주지 않습니다(1.4.14).
  글이 없는 쪽의 장절은 바로 앞 글의 장절로 봅니다.
- 검색어: `A AND B`, `A OR B`(대소문자 무관, AND가 먼저 묶임), `"설계 기준"`(따옴표는 구문, 검색어에 넣지 않음), 연산자 없이 이어진 단어는 한 구문.
  띄어쓰기, 줄바꿈, 기호(`# - . : / ( )` 등)는 무시하고 비교합니다: `SOI 1` = `SOI #1` = `SOI-1`, `설계기준` = `설계 기준`, PDF 줄바꿈으로 갈라진 글.
  숫자 경계는 지킵니다: `SOI 1`은 `SOI #10`, `SOI 21`과 맞지 않고, `1.2`는 `11.2`, `3.1.2`와 맞지 않습니다.
  기호를 빼면 숫자만 남는 검색어(`#4`, `1.2`)는 기호까지 그대로 찾습니다. 더 복잡한 패턴은 `--regex`(예: `"SOI\s*#?\d"`). `--regex`의 빈 일치는 일치로 보지 않습니다.
- `--file 파일`: 캐시가 없거나 낡았으면(수정 시각/크기/파서 버전) 먼저 읽어 등록하고, 최신이면 캐시를 씁니다. 그 파일을 읽지 못하면 `read`와 같은
  오류입니다(입력 문제는 `STOP_RETRY:`, 파서 내부 오류는 `ERROR:`와 traceback, `index --file`도 같음, 1.4.29). 폴더의 실패한 파일은 `skipped`의
  `reason`(내부 오류는 `KeyError: ...`처럼 예외 종류와 함께). `--file 폴더`: 폴더(기본 3단계 하위)의 미등록/변경 문서를 `--max-files`까지 등록한 뒤 그 폴더 문서만 검색합니다.
  결과에 `indexed_now`(새로 등록한 파일 이름, 20개까지), `indexed_now_count`, `from_cache`, 남은 파일이 있으면 `not_indexed_remaining`과 `index --folder` 명령.
- 거르기: `--pages 3-7`(쪽 정보가 없는 단위는 빼고 `page_unknown_skipped`로 개수 보고), `--section "SCOPE"`(장절 경로 일부, 대소문자 무시).
- 결과 항목: `file`, `path`, `page`, `page_source`, `section`(목록), `section_text`(` > `로 이음), `number`, `loc`, `heading`, `evidence`, `snippet`, `--full`이면 `text`.
  공통: `total_matches`(= `count`), `matched`(= `shown`), `truncated`, `files_searched`, `output_file`, `next_commands`. 잘렸으면 `--save "근거.xlsx"`로 전체를 저장하는 명령을 줍니다
  (`--full`에서 엑셀 칸 상한 32,000자보다 긴 단락이 있으면 `--save "근거.md"`, 1.4.13). directive는 저장 이름을 적지 않고 `next_commands`의 저장 명령을 가리킵니다(1.4.14, 이전에는 늘 `근거.xlsx`라 `.md` 명령과 달랐음).
- `evidence` 형식: `규격서.pdf p.12, 1. SCOPE > 1.1. Scope (1.1)`. 쪽이 없으면 `loc`(예: `매출.xlsx 매출!12행`), Word 쪽을 쪽 나눔으로만 셌으면 `p.3(추정)`.
- `--save`: `.xlsx`(머리글 굵게, 줄바꿈, 열 너비, 첫 행 고정), `.csv`, `.md`(기본). 열: 파일, 쪽, 장절, 번호, 위치, 내용(문장/단락), 문맥. `.md`에는 항목마다 `근거:` 줄. 저장 파일에는 `--top`과 관계없이 전체(최대 5,000건)를 담습니다. 5,000건을 넘으면 `save_truncated`와 함께 `UNVERIFIED:`. 저장한 파일은 다시 열어 확인합니다(`save_problem`이 있으면 `UNVERIFIED:`). `.xlsx`는 셀당 32,000자에서 자르므로 잘린 행이 있으면 `save_problem`(`.md`는 전체).
- `index`: 새로 읽은 문서의 캐시 파일이 실제로 저장됐는지 확인하고, 저장하지 못한 파일이 있으면 `cache_failed`와 함께 `UNVERIFIED:`.
  끝에 색인을 다시 읽어 불러온 문서가 모두 있는지 확인하고, 빠진 것(동시에 실행한 다른 `index`가 덮어씀)은 다시 넣어 `recovered`로 알립니다. 다시 넣지 못하면 `index_missing`과 `UNVERIFIED:`.
- 색인 읽기-고치기-쓰기는 `cache/index.lock`으로 한 실행씩만 합니다(1.4.12, 이전에는 두 `index`를 동시에 실행하면 항목이 사라짐). 최대 15초 기다리고(1.4.13, 잠근 쪽의 `retry_fs` 재시도보다 길게), 60초보다 오래된 잠금은 중단된 실행이 남긴 것으로 보고 고유 이름으로 옮긴 뒤 지웁니다(옮긴 것이 새 잠금이면 되돌림). 15초 안에 얻지 못하면 잠금 없이 진행합니다(위 확인이 빠진 항목을 채움).
- 색인 파일을 읽지 못하면(공유 위반 등) `retry_fs`로 다시 읽고, 끝내 못 읽으면 색인을 바꾸지 않고 STOP_RETRY입니다(`read`는 문서를 읽고 색인 넣기만 건너뜀). 빈 색인으로 보고 덮어쓰지 않습니다.
  내용이 손상(JSON·인코딩 오류)되었을 때만 캐시 문서 파일(`cache/docs/*.json`)로 색인을 다시 만들고 `recovered`로 알립니다(1.4.13).
  색인은 바이트로 읽어 UTF-8로 풀고, 풀리지 않는 내용도 손상으로 봅니다(1.4.14, 이전에는 입력 오류가 나 `clear --all`까지 모든 reader 명령이 실패).
- `--top`은 50, `--context`는 500까지(넘으면 줄이고 `warnings`). `--full`의 JSON `text`는 2,000자까지(잘리면 `text_truncated`, `text_chars`, `truncated: true`, 저장 명령에 `--full`).
- `--save .xlsx/.csv`의 문맥 칸: `--full`이면 단락 전체(`.xlsx`는 셀 상한 32,000자 확인), 아니면 3,000자까지이고 잘린 행 수를 `context_clamped`와 directive로 알립니다(1.4.12, 이전에는 `--full`을 무시하고 말없이 자름).

#### 캐시 형식 (모든 형식 공통, `PARSER_VERSION` 8)
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
- `PARSER_VERSION`이 다른(이전 형식) 캐시는 그대로 검색하지 않습니다(1.4.10에서 빈 키를 채우던 보정 삭제). 검색 때 원본에서 다시 읽고(20개까지), 다시 읽지 못하면 `skipped`로 알립니다(위 "검색과 캐시"). `read`, `index`로 다시 읽으면 새 형식으로 바뀝니다. 5: PDF 글자 없는 쪽 `meta.pages_without_text`, 6: 그 쪽을 그림이 있는 쪽으로 좁힘(1.4.11), 7: 머리글·쪽 번호를 뺀 뒤 글이 없고 그림이 쪽 면적의 40% 이상을 덮는 쪽(1.4.12), 8: 남은 글이 40자보다 적은 쪽도(도장, Bates 번호), OCR 글자 층이 있는 쪽은 제외(1.4.13).

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
- 예: 아래 세 형식 모두 올바르게 중첩됩니다 (`tests/test_reader.py`).
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
- `undo [--app] [--list] [--pick N] [--overwrite]`: `--pick`은 1 이상(0 이하는 STOP_RETRY, 1.4.23).
  `--list`의 `count`·`shown`은 되돌릴 수 있는 지점(항목에 `command`가 있음)만 세고, 되돌릴 수 없는 작업 수는 `barriers`. 되돌릴 지점이 0개면 `VERIFIED: 0건.`(1.4.27)
  `--list` 항목은 되돌릴 수 있는 지점 앞 20개와 그보다 최근의 되돌릴 수 없는 작업(최대 20개)이라, 되돌릴 수 없는 작업이 20개 넘게 앞에 쌓여도
  되돌릴 수 있는 지점이 빠지지 않습니다(`truncated`는 빠진 항목이 있을 때 true). `--pick` 없는 `undo`는 되돌릴 수 있는 지점이 하나도 없으면
  `VERIFIED: 0건.`(`count: 0`, `restored: false`), 직전 작업만 되돌릴 수 없고 더 앞 지점이 있으면 `needs_user_decision`입니다(1.4.28).

## 4. 앱 간 전달 (transfer)과 회사 DRM

| 전달 | 명령 | 방식 |
|---|---|---|
| 엑셀 차트 -> 워드/PPT | `--from excel --chart 이름 --to word,ppt` | 그림 (아래 3방식) |
| 엑셀 범위 -> 워드/PPT | `--from excel --range A1:D10 --to word` | 네이티브 표 (값 전달, 클립보드 없음) |
| 워드/PPT -> 엑셀 | `--from word --what text\|table [--table N] --to excel` | 새 시트에 표로 |

닫힌 docx의 `--table N`은 `reader.py read`·Word(COM)와 같은 번호입니다: 콘텐츠 컨트롤(`w:sdt`)·customXml 안의 표도 문서 순서로 셉니다
(1.4.31, 이전에는 본문 바로 아래 표만 세어 다른 표를 옮겼음). 글상자 안의 표는 세지 않습니다(Word `Tables`와 같음).

차트 그림은 다음 순서로 자동 시도하고, 결과 JSON의 `method`에 실제로 쓴 방식을 적습니다.
1. `export`: Excel의 `Chart.Export` PNG를 씁니다. **파일 앞부분이 PNG가 아니면(DRM 암호화) 건너뜁니다.**
2. `redraw`: 차트의 계열 값을 COM으로 읽고 matplotlib으로 다시 그립니다. Python이 쓴 파일을 쓰고 클립보드는 쓰지 않습니다. 원본 서식은 근사치입니다.
   제목·계열 이름·항목 이름의 `$`는 글자 그대로 그립니다(수식으로 읽지 않음, 1.4.29).
3. `relay`: 엑셀 -> 워드 -> PPT 클립보드 경유입니다. 받는 문서와 앵커를 찾은 뒤에 클립보드를 씁니다(1.4.29).
   PPT로 보낼 때는 `--slide`도 확인한 뒤에 클립보드에 복사하고 임시 워드를 띄웁니다(1.4.30, 이전에는
   2장짜리에 `--slide 9`면 클립보드를 덮어쓴 뒤 오류). `--to-target ""`는 export·redraw처럼 활성 문서입니다(1.4.30, 이전에는 `KeyError`).
   열린 PPT·활성 PPT는 찾기만 하고 슬라이드를 확인하며, 쓰기 세션(되돌리기 지점)은 임시 워드 단계 뒤에 엽니다. 닫힌 파일은 열어서 확인합니다
   (멈추면 파일이 그대로라 지점은 지움), 새로 만들 파일은 슬라이드 0장으로 확인합니다(만들지 않음). 1.4.30은 쓰기 세션을 먼저 열어, 임시 워드
   단계가 실패하면 저장 안 된 PPT에 바꾼 것 없는 되돌리기 지점이 남고 다음 `undo`가 그 지점을 먼저 골랐습니다(1.4.31).

방식을 바꿔도 같은 문제(받는 문서를 찾지 못함: 오타 `[추천]`, `[후보]`, 확장자, 읽기 전용 / 앵커 / 슬라이드 번호)는 다음 방식으로 넘어가지 않고
그 오류 그대로 멈춥니다. `[추천]`은 message 끝에 있고 `--probe` 안내를 붙이지 않습니다(1.4.29, 이전에는 다음 방식까지 시도해
relay가 사용자의 클립보드를 덮어쓰고 `[추천]`이 `; relay: ...` 앞에 묻혔음). 모든 방식이 실패한 오류의 오류 로그에는 내부 오류의 글 대신
`<방식>: <예외 종류>`만 남깁니다.

회사 PC에서 처음 한 번 `python office.py transfer --probe`를 실행하면, 세 방식을 실제로 시험해 대상 앱별로 되는 순서를 `config.json`에 저장합니다.

여러 앱에 보낼 때(`--to word,ppt`, 1.4.20):
- 무엇이든 넣기 전에 앱마다 받는 문서(각 앱의 활성 문서 또는 `--to-target`)와 앱별 인자(`--anchor`, `--slide`)를 확인합니다.
  못 찾은 문서가 있으면 아무것도 바꾸지 않고 STOP_RETRY(`error_type: target_not_open`)와 찾은 앱만으로 바꾼 같은 명령
  `[추천] python ...office.py transfer ... --to word`를 냅니다. 에이전트는 그대로 1회 실행하고 못 넣은 앱을 사용자에게 알립니다.
- `--to-target`은 `--to`의 모든 앱에 함께 쓰이므로 파일 경로를 줄 때는 `--to`에 앱 하나만 둡니다.
  열린 문서가 없을 때 안내는 받는 쪽 `--to-target`, 보내는 쪽 `--from-target`입니다(transfer에는 `--target`이 없음).
- 넣은 뒤 한 앱만 실패하면(늦게 드러난 대상 문제, COM 오류) 오류로 끝내지 않고 `UNVERIFIED`와 `succeeded`/`failed`/`failures`로 답합니다.
  같은 명령을 다시 실행하면 넣은 앱에 또 들어가므로, 실패한 앱에만 다시 넣으려면 원인을 고친 뒤 `--to <그 앱>`으로 실행합니다.
  되돌리기 지점은 실제로 바뀐 문서에만 남습니다.
- `excel chart`의 `next_commands`는 문서가 열린 받는 앱만 `--to`에 넣습니다(둘 다 열렸거나 둘 다 없으면 `word,ppt`).
  1.4.21부터 `--sheet <차트가 있는 시트>`도 넣습니다.

차트 찾기(`--chart`, `chart --action delete --name`, 1.4.21): 이름 또는 제목(대소문자 무시), 없으면 번호.
`--sheet`를 생략하면 모든 시트에서 찾고, 이름이 같은 차트를 제목이 같은 차트보다 우선합니다. 그래도 둘 이상이면(다른 시트의 같은 이름/제목)
하나를 고르지 않고 STOP_RETRY와 `[후보] 시트!이름`을 냅니다. 사용자가 고른 `시트!이름`을 그대로 차트 이름으로 주면 됩니다
(앞부분이 있는 시트 이름일 때 `--sheet`와 같음). 이름이 같은 열린 문서가 여럿일 때의 `[후보]` 안내도 그 명령의 옵션
(`--to-target`/`--from-target`)을 말합니다.
1.4.23부터:
- `'시트'!이름`(작은따옴표 안의 `''`는 `'`)은 그 시트에서만 찾습니다. 시트·차트 이름에 `!`나 `'`가 있으면 `[후보]`가 이 형식입니다.
- 따옴표 없는 `시트!이름`은 `!` 자리마다 나눠 보고, 앞부분이 있는 시트이고 그 시트에서 찾아질 때만 씁니다. 글자 그대로를 차트 이름으로도 찾아
  (다른 시트의 차트 이름이 `Sheet1!x`일 수 있음) 둘 다 찾아지면 하나를 고르지 않고 `[후보]`를 냅니다.
- 한 시트에 이름이 같은 차트가 여럿이면 `[후보] 시트!#번호`(그 시트의 몇 번째 차트). `#번호`(또는 번호)를 그대로 받습니다.
  1.4.24부터 후보가 아닌 차트까지 그 시트의 모든 차트로 셉니다(제목으로 찾은 차트의 이름이 다른 차트와 같아도 후보마다 차트 하나).
- `--sheet`가 `시트!이름`의 시트와 다르고 달리 찾은 차트가 없으면 두 시트를 말하는 STOP_RETRY. 차트가 없다는 안내는 어느 시트인지 말합니다.
- 차트 그림 파일 이름(`chart_<차트 이름>`)의 `/` 같은 글자는 `_`로 바꿉니다(제목 `1/4분기 매출`로 만든 차트를 옮기지 못하던 문제).

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
:: 러너는 인자를 먼저 해석한 뒤 pywin32를 확인합니다: --agent -h는 JSON exit 0, 인자 오류는 exit 2, pywin32가 없으면 exit 1 (1.4.29)
```
- 실기 테스트는 테스트 폴더에 `report_*.txt`(요약)와 `errors_*.txt`(분석용 상세 로그)를 남깁니다. 문제가 있으면 `errors_*.txt`를 그대로 Claude에게 전달하세요.
  `errors_*.txt`에는 러너의 자동 복구(닫힌 A 문서 다시 열기), 창/프로세스 정리 기록(어떤 PID에 Quit/WM_CLOSE를 보냈고 무엇을 보호했는지), 도구 오류 로그가 함께 들어 있습니다.
- 러너는 테스트 문서(A.xlsx/A.docx/A.pptx)가 열린 프로세스를 끝내지 않습니다. 정리 판단 규칙은 `tools/win_test_rules.py`(오프라인 테스트 `tests/test_win_runner.py`).
- DRM PC에서 python-pptx 기본 템플릿(`pptx/templates/default.pptx`)이 암호화돼 있으면 그 템플릿이 필요한 오프라인 테스트만 "python-pptx 기본 템플릿이 손상됨 (DRM 추정)"으로 건너뜁니다. python-docx는 폴더 템플릿으로 대신 만듭니다. 도구 코드는 기본 템플릿을 쓰지 않습니다.
- 도구는 `C:\CLI\office-cli`에 둡니다 (SKILL.md의 명령 경로 기준). 실기 테스트 파일은 `C:\CLI-dev`(이 저장소의 `CLI-dev` 폴더)에 둡니다. OpenCode 에이전트에게 실기 테스트를 맡기려면 스킬 폴더 `office-cli-test`에 `C:\CLI-dev\office-cli\tools\office-cli-test\SKILL.md`를 복사합니다. 에이전트는 `--agent`로 실행해 결과 JSON 1줄만 읽습니다.
- 대상은 Office 16.0(2016/2019/2021/365)입니다. Word 2007의 `SaveAs2` 부재는 `SaveAs`로 대체하므로 Office 2007에서도 실기 테스트가 돕니다. 다만 PDF 추가 기능이 없으면 PDF 내보내기 항목은 `ENV-SKIP`으로 기록됩니다.
