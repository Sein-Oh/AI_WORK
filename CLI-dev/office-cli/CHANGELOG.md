# CHANGELOG - office-cli

형식: 버전별 `구조 / 기능 / 수정` + 하단 `남은 할 일`. 1.0.0부터 관리합니다.

## [1.4.10] - 2026-11-01

전체 검토에서 찾은 오류 수정. 회귀 테스트: `tests/test_review_v1410.py`.

### 수정 (엑셀)
- `excel replace --range '요약'!A1:D9`: 읽기는 `--range`의 시트, 쓰기는 지금 시트에 하고 `VERIFIED`였던 문제. 이제 범위의 시트에서 읽고 씁니다. `--all-sheets`와 시트 이름이 든 `--range`는 함께 쓸 수 없음(STOP_RETRY).
- `excel replace --find X`만 주면(바꿀 글·서식 없음) 아무것도 바뀌지 않는데 "N개 셀 처리" `VERIFIED`였던 문제. 이제 STOP_RETRY(`word replace`도 같게).
- `excel range --action merge`: 왼쪽 위가 비고 다른 칸 하나에만 값이 있으면 확인 없이 병합해 그 값이 지워지던 문제. 왼쪽 위 밖의 값이 하나라도 있으면 `needs_user_decision`.
- 덮어쓰기 확인(`write`, `range copy/move`, `merge`): 빈 글("")을 내는 수식 칸을 빈 칸으로 보고 확인 없이 덮어쓰던 문제. 수식도 확인합니다(`--append`의 마지막 행 찾기도).
- `excel write --append --formula/--value`: 첫 열만 보고 마지막 행을 정해 옆 열의 데이터를 덮어쓰던 문제. 범위의 열 수(또는 수식 개수)만큼 보고, `--append`에도 덮어쓰기 확인을 합니다.
- `excel range --action copy`: 상대 참조 수식은 값이 달라져 늘 `UNVERIFIED`였던 문제. 수식(R1C1)으로 비교합니다.
- `excel range --action autofit`: 확인하지 않고 `VERIFIED`(verified false)였던 것을 `UNVERIFIED`로.
- 닫힌 xlsx·CSV 읽기를 열린 문서와 같게: 거꾸로 준 범위(`C5:A1`)가 빈 결과였던 문제, 크기 정보(`<dimension>`)가 없는 파일에서 `B:B`가 `B:D`로·`5:7`이 끝 열 없이 읽히던 문제(실제 데이터 범위를 셈), 범위를 생략하면 A1부터 읽어 `rows`가 열린 문서와 달랐던 문제(데이터가 시작하는 칸부터).

### 수정 (워드·PPT)
- `word replace`: 문구당 5,000곳 상한에 닿으면 확인 없이 `VERIFIED`였던 것을 `UNVERIFIED`로. `--match-case` 없이 대소문자만 바꾸면(`office` -> `Office`) `UNVERIFIED`로 나오던 문제.
- `ppt.py write --slide new --title 제목 --file 내용.md`: `.md`가 있으면 `--title`이 버려지던 문제. 첫 슬라이드 제목으로 넣고 첫 절의 제목은 본문 첫 줄로.
- `ppt.py write --slide N --title`: 기존 슬라이드의 다른 제목을 확인 없이 바꾸던 문제. `needs_user_decision`(노트와 같은 방식), 사용자가 원하면 새 인자 `--overwrite`. 새 슬라이드는 해당 없음.
- `ppt.py read --slides 1,5,9`가 잘리면 이어 읽기 명령이 `2-9`였던 문제. 고른 슬라이드 중 남은 번호(`5,9`)로.

### 수정 (공통·office·outlook)
- `_runtime.ok()`가 `verified=False`를 받으면 내부 오류(ValueError). 확인하지 못한 결과는 `ok_unverified`로 무엇을 확인하지 못했는지 적습니다. `office.py transfer` 일부 실패도 `UNVERIFIED`로.
- `_runtime.ok_file`: 크기만 보던 것을 형식별로 다시 열어 확인(xlsx openpyxl, docx python-docx, pptx python-pptx, CSV 행 수, md/txt 읽기, PDF 머리). 열지 못하면 `UNVERIFIED`.
- 오류 로그(`errors.txt`)에 메일 본문·받는 사람·제목, `--text`, 검색어 등 내용 옵션 값이 그대로 남던 문제. `<N자 생략>`으로 남기고 경로와 그 밖의 인자는 그대로.
- 쓰기 `--target`에 파일 이름만 주었는데 열린 문서도 파일도 아니면 현재 폴더에 새 파일을 만들던 문제. "찾을 수 없음"(열린 문서가 있으면 `[후보]`), 새 파일은 전체 경로일 때만.
- `office.py undo --pick N`(N>1): 고른 지점만 지워 더 새 지점이 남아 다음 `undo`가 앞으로 가던 문제. 같은 문서의 더 새 지점도 지웁니다(`removed_newer_points`). 수정 시각 비교는 같은 문서의 가장 새 지점 기준(그 사이 도구 저장을 사용자 저장으로 보지 않음).
- `outlook.py read --save`: 제목을 그대로 파일 이름에 써 `RE:`·`/`·`?` 등에서 실패하던 문제(첨부 저장과 같은 정리). 저장한 파일을 다시 읽어 확인.
- `outlook.py --body`: `\n` 변환을 두 번 해 `--body-file`의 `C:\new\data`까지 깨던 문제. `--body`만 `unescape_newlines`(경로 보존), `--body-file`은 그대로.
- `outlook.py --attach`: 쉼표로만 나눠 이름에 쉼표가 든 파일을 첨부하지 못하던 문제. 값 전체가 있는 파일이면 하나로, 여러 개는 `;`(쉼표도 됨).

### 수정 (reader)
- `read`: 한 단위가 약 8,000자보다 길어 잘려도 `truncated: false`였던 문제. `truncated: true`, 단위의 `text_truncated`·`text_chars`, `truncated_reason`, 항상 `--save` 명령(잘린 뒤쪽은 저장 파일에 있음).
- 일부 쪽만 스캔된 PDF: 문서 전체에 글자가 없을 때만 알려 0건 검색이 `VERIFIED`였던 문제. 글자 없는 쪽을 `meta.pages_without_text`로 남기고 `search`의 `not_searched`와 `read` 결과에 알립니다(`PARSER_VERSION` 5, 기존 캐시는 다시 읽음).
- 엑셀 시트 행 상한(20,000행)에서 자른 문서를 `read`/`read --save`가 그대로 `VERIFIED`로 알리던 문제(`truncated`, 저장은 `UNVERIFIED`). `search --save`가 5,000건에서 멈춰도 `VERIFIED`였던 문제(`save_truncated`, `UNVERIFIED`).
- `read --save`, `search --save`: 저장한 파일을 다시 열어 확인(CSV는 행 수). 확인하지 못하면 `UNVERIFIED`.
- `index`: 캐시 저장에 실패해도 "캐시에 저장 확인"이던 문제. 캐시 파일을 확인해 `cache_failed`와 `UNVERIFIED`.
- 옛 캐시의 빈 키를 채우던 `normalize_doc` 삭제 (베타, 이전 형식 캐시는 `PARSER_VERSION`으로 다시 읽음). 다시 읽지 못한 이전 형식 캐시는 검색하지 않고 `skipped`로 알립니다.

### 정리
- pyflakes: `ppt.py`의 쓰지 않는 `nonlocal slide_no`, `word.py` `_compress`를 대입 전 참조 없이 다시 씀. 테스트의 BOM 글자를 `"\ufeff"`로.

## [1.4.9] - 2026-10-31

### 수정
- 닫힌 Word 파일 작업: Word 2007은 `ReadOnly`와 `PasswordDocument`를 함께 넘기면 암호 없는 파일도 "명령을 정상적으로 종료할 수 없습니다"로 열지 못했습니다. 파일 구조로 암호 여부를 먼저 판별해, 암호가 없다고 확인된 파일에는 넘기지 않습니다. 판별하지 못한 파일(DRM 등)은 전과 같이 넘깁니다.
- 암호 Word 문서: Word 2007의 오류 문구에는 '암호'가 없어 `Office 오류`로 나오던 것을 "열기 암호가 걸린 문서" 안내로 바꿈.

## [1.4.8] - 2026-10-30

### 변경 (R1 사후 확인 마무리)
- 쓰기 명령은 모두 결과를 다시 읽어 확인한 경우에만 `VERIFIED:` + `verified: true`, 아니면 `UNVERIFIED:` + `unverified`.
  - `excel replace`: 바꾼 칸의 값. `excel style`: 준 서식(배경·글자색·굵게·기울임·크기·숫자 서식), 테마 머리글 색, 조건부 강조.
  - `excel sheet`: 시트 목록 (추가·이름 변경·복사·삭제·선택). `excel range`: 복사·이동한 대상 값, 지운 칸, 병합 상태, 너비·높이, 숨김, 틀 고정 (자동 맞춤은 비교할 값이 없음).
  - `excel chart`: 만든 차트가 목록에 있고 제목이 같은지, 지운 차트가 목록에 없는지.
  - `word replace`: 문서를 다시 찾아 찾는 글이 남지 않았는지 (바꾼 글이 찾는 글을 품으면 바꾼 글이 그만큼 있는지). `ppt replace`: 바꾼 슬라이드를 다시 읽어 같은 확인.
  - `ppt slide`: 슬라이드 수, 옮긴 슬라이드 제목, 비운 슬라이드의 요소 수.
  - `transfer` 차트: 넣기 전후 그림 수 (붙여넣기 방식 relay는 `UNVERIFIED`).
  - `office.py undo`: 되돌린 원본 파일이 스냅샷과 같은지.
  - `outlook --send`: 실제 발송은 Outlook이 처리하므로 `UNVERIFIED`(보낸 편지함 확인 안내). 작성 창은 받는 사람·제목을 다시 읽어 `VERIFIED`.
- `_runtime.ok()`는 더 이상 `VERIFIED:`를 대신 붙이지 않음 (다른 문구가 오면 내부 오류). 읽기·조회 명령은 읽은 결과가 곧 확인이라 기본 문구 유지.

## [1.4.7] - 2026-10-30

### 변경 (R1 사후 확인)
- 내보내기(`excel export`, `word export`, `ppt export`): 저장한 파일이 있고 비어 있지 않은지 확인 (`_runtime.ok_file`).
- `office.py transfer`: 엑셀 범위를 워드·PPT 표로 넣으면 넣은 범위·슬라이드에서 머리글을, 워드·PPT 내용을 엑셀로 넣으면 넣은 칸을 다시 읽어 확인.

## [1.4.6] - 2026-10-30

### 변경 (R1 사후 확인)
- `excel write`: 쓴 칸을 다시 읽어 비교 (수식은 첫 칸의 수식, 값은 칸마다. 앞의 `'`는 값에 남지 않으므로 빼고 비교, Excel이 날짜 등으로 바꾼 글자는 비어 있지 않으면 인정).
  같으면 `VERIFIED:` + `verified: true`, 다르면 `UNVERIFIED:` + `unverified`(다른 칸).
- `word write`: 넣은 범위를 다시 읽어 비어 있지 않고 넣은 글의 첫 줄(마크다운 기호 제외)이 들어 있는지 확인.
- `ppt write`: 쓴 슬라이드를 다시 읽어 제목·첫 줄·표 머리글·카드 제목이 들어 있는지 확인.
- 위 명령의 "다시 읽어서 확인하지 마세요" 문구를 "다시 읽어 확인했습니다"로.

## [1.4.5] - 2026-10-30

전체 검토와 회사 PC 오류 로그에서 찾은 오류 수정.

### 수정
- 문서: 문자 규칙(유니코드 화살표, 긴 줄표) 정리, README 구성도(word outline/comment/revisions, ppt notes, reader clear), 결과 처리 표에 `UNVERIFIED:`, 깨진 SKILL.md 링크.
- `excel write`: 머리글을 먼저 쓰고 수식 검사에서 실패해 머리글만 남던 문제. 모든 검사를 마친 뒤 씀. 수식이 여러 개인데 범위가 없으면
  머리글 행부터의 범위를 `[추천]` (이전 예시를 따르면 한 행씩 밀려 `#VALUE!`인데 ok).
- `excel write`: 값이 있는 칸을 확인 없이 덮어쓰던 문제 (SKILL.md "덮어쓸 칸"과 불일치). 이제 `needs_user_decision`, 사용자가 원하면 `--overwrite`.
- `excel write --formula 73.4 49.9`처럼 숫자만 주면 `--values`를 `[추천]`.
- `office.py transfer --to excel`: 워드/PPT 글이 숫자·날짜·수식으로 바뀌던 문제 ("001", "1-2", "(100)", "=..."는 글자 그대로).
- `excel read`: 닫힌 xlsx·CSV에서 `B:B`, `5:7` 같은 열·행 범위를 읽지 못하던 문제.
- 워드·PPT `--text`의 `\n` 변환이 Windows 경로(`C:\new\data`)를 깨던 문제.
- `office.py undo --list`: 항목마다 그대로 실행할 `command`(같은 `--app` 포함). 필터에 따라 번호가 달라 다른 문서가 되돌려질 수 있었음.

### 기능
- 인자 오류에 `[추천]`: 없는 명령(`excel.py list` -> `info`, 오타 -> 가까운 명령), 다른 옵션 이름(`office.py undo --target` -> `--app`),
  따옴표 없이 띄어쓰기가 든 경로(큰따옴표 안내).
- 열린 문서가 없으면 `needs_user_decision`: 사용자에게 전체 경로를 물어 `--target`으로 1번 실행 (경로 추측 금지).

## [1.4.4] - 2026-10-30

### 변경
- SKILL.md 규칙 9: 도구 코드(.py)는 사용자가 코드를 열어 보라고 직접 요청할 때만 읽음 (ml-kit과 같게).

## [1.4.3] - 2026-10-29

### 구조
- 폴더 정리: `CLI/office-cli/`에는 실행에 필요한 파일과 `SKILL.md`만 남김 (`excel.py`, `word.py`, `ppt.py`, `outlook.py`, `reader.py`, `office.py`, `_runtime.py`, `_office.py`).
  `README.md`, `CHANGELOG.md`, `TEST_CHECKLIST.md`, `requirements.txt`, `tests/`, `tools/`는 `CLI-dev/office-cli/`로 옮김.
  공통 계약 `CLI/README.md`와 `run_all_tests.py/.bat`도 `CLI-dev/`로.
- 테스트와 실기 러너는 옆 폴더 `../CLI/office-cli`의 도구를 시험. 사내 PC에서는 실기 테스트 때만 `CLI-dev`를 `C:\CLI-dev`로 `C:\CLI`와 나란히 둠.
- CLI 인자, JSON 키, 도구 동작은 그대로.
- `pyproject.toml` 삭제: pytest `testpaths` 2줄뿐이었고 반입 검사에 자주 걸림.
- 실기 테스트 스킬(`office-cli-test`)의 러너 경로: `C:\CLI-dev\office-cli\tools\run_win_tests.py`.

### 수정
- `-h` 결과 JSON에 `skill_file`(SKILL.md 경로)과 "도구 코드(.py)는 읽지 말고, 모르면 사용자에게 물어보라"는 안내 추가.
- SKILL.md 규칙 9: 도구 코드(.py)를 읽지 않고, 방법을 못 찾으면 사용자에게 알리고 멈춤.

## [1.4.2] - 2026-10-29

### 기능
- 새 명령 `reader.py clear`: 캐시 지우기(색인 항목과 캐시 파일, 원본 문서는 그대로). 대상은 `--file`, `--name`, `--folder`, `--missing`(원본이 없어진 문서), `--all`(색인에 없는 남은 캐시 파일 포함).
  `--dry-run`은 지울 목록(`planned`)과 실제로 지우는 명령을 줍니다. 지운 뒤 색인과 캐시 파일을 다시 확인해 `verified`.
- `reader.py search`: 띄어쓰기·줄바꿈에 더해 기호(`# - . : / ( )` 등)도 무시하고 찾습니다. `SOI 1`로 `SOI #1`, `SOI-1`을 찾습니다(이전에는 `SOI #1`로만 찾힘).
  숫자 경계는 지켜 `SOI 1`이 `SOI #10`과, `1.2`가 `11.2`·`3.1.2`와 맞지 않습니다. 기호를 빼면 숫자만 남는 검색어(`#4`, `1.2`)는 기호까지 그대로 찾습니다.
  결과 강조(`**`)와 저장 파일의 문장도 같은 기준입니다.

### 문서
- SKILL.md(office-cli, office-cli-test) 보고 규칙: 파일 경로를 `file:///` 링크로 바꾸지 않고 결과 JSON 값을 그대로 1번만 쓰게 함.
  긴 대화에서 26B 모델이 한글 경로를 링크로 변환하다가 같은 조각(`l'al l'al ...`)을 끝없이 반복한 사례 대응.

## [1.4.1] - 2026-10-29

배경: 회사 PC 실기 (2026-10-29 로그). 실기 러너 71/71 통과(E3, K1~K5 포함). 오프라인 테스트 실패 4건·오류 10건과 OpenCode 대화 6개에서 나온 문제를 고쳤습니다.

### 수정
- `ppt.py notes --file`, `word.py comment --file` 등: Windows 메모장 형식(CRLF) 파일을 넣으면 줄마다 빈 줄이 하나 더 생기던 문제. `read_text_file`이 줄바꿈을 `\n`으로 맞춥니다.
- `excel.py read --target <CSV>`: CSV/TSV도 Excel 없이 읽습니다(인코딩 자동 판별, 숫자 변환, `--range`). 이전에는 확장자 오류가 나 Gemma가 pandas 코드를 직접 짰습니다.
- reader 제목 인식: `(3) Verification Review; and`처럼 `;` `,` `and/or/및/또는`으로 끝나는 나열 항목은 장절 제목이 아닙니다(검색 근거의 장절이 나열 항목으로 바뀌던 문제). `PARSER_VERSION` 4.

### 테스트
- 회사 PC처럼 python-docx 기본 템플릿에 numbering.xml이 없으면 번호 문단을 만드는 테스트 도우미가 `NotImplementedError`로 실패하던 문제(오류 10건, 실패 3건). 도우미가 번호 파트를 직접 만듭니다. 도구 동작과는 무관합니다(reader는 번호 파트가 없어도 읽음).

## [1.4.0] - 2026-10-28

배경: 동료가 만드는 office MCP와 기능을 비교해 필요한 것만 가져왔습니다. 규격서·보고서 검토(장절을 골라 읽고, 검토 결과를 메모로 남기고, 수정 기록을 확인)와 발표 준비(노트)에 씁니다.

### 기능 (추가, 기존 인자·키 유지)
- `word.py read`에 범위 고르기 `--pages 3-5`, `--paragraphs 120-140`(`141-`는 끝까지), `--section "제목 일부"`(+ `--occurrence N`). 셋 중 하나만 씁니다.
  - `--pages`: Word가 계산한 실제 쪽(`GoTo`, `Information(wdActiveEndPageNumber)`). 앞 쪽에서 이어지는 단락도 넣고 단락마다 `page`. 결과에 `page_total`. 닫힌 .docx는 reader 캐시와 같은 쪽 정보(저장된 쪽 경계)로 고르고, 쪽 정보가 없으면 `--paragraphs`/`--section`을 안내합니다.
  - `--section`: 제목 전체가 같은 것 -> 제목이 그 글로 시작하는 것(`1.1`은 `1.1. Scope`와 맞고 `1.1.1 Data`와는 안 맞음) -> 제목에 들어 있는 것 순으로 찾습니다. 여럿이면 `[후보]`, 없으면 비슷한 제목 `[추천]`.
  - 결과에 `selection`, `count`. 잘리면 나머지를 `--paragraphs a-b`(또는 `--pages`)로 이어 읽는 명령과 고른 범위 전체를 `--save`하는 명령. 0건은 `count: 0`과 안내.
- 새 명령 `word.py outline`: 제목마다 `paragraph`(read `--paragraphs`와 같은 번호), `level`, `text`, `page`. 200개씩 `--start`.
- 새 명령 `word.py comment`: 검토 메모 목록(`--action list`, 기본) / 메모 달기(`--action add --anchor 글 --text 메모` 또는 `--file`).
  메모는 기준 글이 든 단락을 `write --anchor`와 같은 규칙으로 고른 뒤 그 글 자체에 답니다. `--file 메모.csv`(열: 기준 글, 메모 내용[, 몇 번째 단락])는 여러 메모를 한 번에 달고, 기준 글 하나라도 정할 수 없으면 하나도 달지 않습니다(어느 줄이 왜 안 되는지 message). 단 뒤 메모 수와 내용을 다시 읽어 확인합니다(`verified`, 다르면 `UNVERIFIED:`). `office.py undo`로 되돌릴 수 있습니다.
- 새 명령 `word.py revisions`: 변경 내용 추적 목록(`type` 삽입/삭제/서식 변경 등, `author, date, page, text`, `track_changes`). 읽기만 합니다.
- 새 명령 `ppt.py notes`: 발표자 노트 읽기(`--slides`, 노트가 있는 슬라이드만, 0건 안내) / 쓰기(`--slide N --text` 또는 `--file`, `--append`). 쓴 뒤 다시 읽어 확인(`verified`, `before`, `after`).
- 공통 런타임: `ok_unverified()` (R1의 `verified: false` + `UNVERIFIED:` 응답). 이번에 추가한 쓰기 명령(`comment --action add`, `notes` 쓰기)부터 `verified` 키를 씁니다.

### 점검 반영: 공통 런타임 (모든 명령에 영향)
전체 코드 점검에서 경량 LLM을 잘못 이끌 수 있는 공통 동작을 고쳤습니다 (R3, R4, R5, R7, R8).
- `-h`/`--help`: 도움말은 stderr, stdout에는 `{"status":"ok","command":"help","verified":true,...}` 1줄 (이전에는 stdout에 사람용 글만 나와 JSON 파싱 실패). 서브명령도 같음.
- `[추천]` 기준: 표기만 다른(대소문자·공백·`_`·`-`) 이름이 하나일 때만 `[추천]`. 부분·유사 일치는 뜻이 다를 수 있어 `[후보]`(사용자가 고름, `error_type: ambiguous`).
  예: 시트 `2024` -> 이전 `[추천] 2024_1분기`, 이제 `[후보] 2024_1분기, 2024_2분기`. 첨부 `견적서.pdf`가 없을 때 `견적서_구버전.pdf`를 추천해 `--send`와 함께 다른 파일이 발송될 수 있던 문제.
  도구가 정한 옵션 값(색 이름, 테마, 형광펜 색, 앱 이름)의 오타는 가까운 값 하나를 계속 `[추천]`.
- 오류 분류: 일반 `ValueError`/`KeyError`는 코드 버그일 수 있어 `ERROR:`(내부 오류, 오류 로그에 traceback). 입력 문제는 `InputError`로 냅니다.
- 출력 파일: `--save`, `export` 등의 저장 경로에 같은 이름의 파일이 있으면 덮어쓰지 않고 `_YYYYMMDD_HHMMSS`를 붙인 새 이름으로 저장합니다. 결과에 `recovered`.
  잠긴 파일은 잠깐 기다려 보고(`retry_fs`) 그래도 잠겨 있으면 새 이름으로 저장하고 `recovered`. 이미 시각이 붙은 이름에 시각을 두 번 붙이지 않습니다.
- `note_recovered()`와 응답의 `recovered` 키(성공·오류 모두). `retry_fs`는 일시적 오류(접근 거부, 공유·잠금 위반)만 재시도하고, 없는 파일 같은 영구 오류는 바로 올립니다.

### 점검 반영: 1.4.0 새 기능
- `ppt.py notes` 쓰기: 기존 노트가 있으면 `--append` 또는 `--overwrite`가 있어야 바꿉니다. 둘 다 없으면 노트를 지우지 않고 `needs_user_decision`.
  노트 본문 개체틀이 없는 파일에서 다른 개체틀(머리글·쪽 번호)에 쓰던 문제. 읽기에서 `--slide 3`도 그 슬라이드만 읽습니다. 이어 읽기 명령이 고르지 않은 슬라이드를 넣던 문제.
- `word.py comment`: `--anchor`/`--text`/`--file`을 주고 `--action add`를 빠뜨려도 메모를 답니다 (이전에는 목록만 보여 주고 완료로 응답).
- `word.py read --paragraphs`에 떨어진 범위(예: `1-12,20`)를 주고 결과가 잘리면 오류가 나던 문제. 이어 읽기는 남은 단락 번호로 줍니다.
- `ppt.py read/info/notes --slides` 형식 오류 문구가 `--pages`를 안내하던 문제.

### 점검 반영 2: 데이터 손상 방지
- 엑셀 쓰기(`write`, `replace`): Excel이 COM으로 받은 글자를 직접 입력한 것처럼 다시 해석해 값이 바뀌던 문제. 앞자리 0 코드(`01234` -> 1234),
  15자리 초과 숫자, 월-일·분수처럼 보이는 글(`1-2`, `3/4` -> 날짜), 괄호 숫자(`(100)` -> -100), 지수(`1E5`)는 앞에 `'`를 붙여 글자 그대로 씁니다(`'`는 셀 값에 남지 않음).
  15자리를 넘는 숫자 글자는 숫자로 바꾸지 않습니다(정밀도 손실). 연도가 있는 날짜(`2026-10-28`)는 계속 날짜로 바뀝니다.
- `excel.py range --action clear`: `--range`를 빼면 시트 전체를 지우던 문제. `--range`가 필요합니다.
- `range --action merge`: 값이 있는 칸이 둘 이상이면 왼쪽 위 밖의 값이 지워지므로 병합하지 않고 `needs_user_decision`(지워질 칸과 값). `--overwrite`로만 병합.
- `range --action copy/move`: 대상 범위에 값이 있으면 덮어쓰지 않고 `needs_user_decision`(원본과 겹치는 칸은 제외). `--overwrite`로만 덮어씀.
- `range --action hide/unhide/autofit`: `--axis`를 생략하고 `5:7`처럼 행 주소를 주면 시트의 모든 열이 숨겨지던 문제. 행 주소면 행으로 판단합니다.
- `write --append`: 시작 열 하나만 보고 마지막 행을 찾아 다른 열의 데이터를 덮어쓰던 문제(쓸 열 전체에서 가장 아래 행). 머리글 비교를 1행이 아니라 시작 행과 하고, `--header`도 같은 머리글이면 생략. 시작 행 위쪽에 쓰던 문제.
- `word.py write --anchor`: 앵커가 표 칸 안이면 `after`는 옆 칸이 아니라 표 뒤에 넣고, `before`/`replace`는 거절합니다. 문서 마지막 단락이 앵커일 때 앵커 단락에 글이 붙고 스타일이 바뀌던 문제. 표로 시작하는 문서의 `--pos start` 거절.
- `office.py undo`:
  - 스냅샷을 만들지 못한 쓰기와 새 파일 쓰기는 "되돌릴 지점 없음" 표식을 남깁니다. 직전 작업이 그렇다면 더 오래된 지점으로 넘어가 그 사이 작업까지 지우지 않고 `needs_user_decision`(`--list`에 `undo_available: false`).
  - 되돌리기 지점 뒤에 파일이 다시 저장되었으면 `needs_user_decision`(`--overwrite`로만 되돌림).
  - 열린 문서를 닫은 뒤 복사가 실패하면 원본을 다시 열고 `undo_failed` 오류에 `pre_undo_backup` 경로. 복사는 `retry_fs`로 재시도.
- 쓰기 대상(`--target`, transfer 대상)이 열린 문서도 파일도 아니고 이름이 비슷한 열린 문서가 있으면 새 파일을 만들지 않고 `[후보]`. 새 파일은 전체 경로일 때만 바로 만듭니다.

### 점검 반영 3: reader 근거 정확도 (`PARSER_VERSION` 3, 이전 캐시는 다시 읽을 때 갱신)
- PDF: 여러 쪽의 위/아래 여백(8%)에 반복되는 줄(문서 제목 머리글, 쪽 번호)을 뺍니다(`meta.header_footer_removed`). 이전에는 머리글이 매 쪽 장절 경로를 끊고 검색 결과에 쪽마다 걸렸습니다.
- PDF 2단 편집: 가로로 걸친 블록(제목, 표) 사이마다 왼쪽 단을 먼저 읽습니다. 이전에는 y 좌표로만 정렬해 두 단의 줄이 섞여 장절이 틀어졌습니다.
- 제목 인식: 번호 뒤가 단위인 수치 줄(`10.5 kV 이하`, `3 mm`, `2 개월`)과 목차 줄(`1.2 설계 기준 ....... 12`, 탭 뒤 쪽 번호)은 제목이 아닙니다.
- DOCX: 변경 추적으로 삽입된 글(`w:ins`, `w:moveTo`)을 넣고 지운 글은 뺍니다. 콘텐츠 컨트롤(`w:sdt`) 안의 본문도 읽습니다. 여러 쪽에 걸친 표는 행마다 그 행의 쪽(이전에는 모두 표가 시작한 쪽).
  빈 번호 단락도 번호 카운터를 진행해 뒤 조항 번호가 밀리지 않습니다.
- HWP: 확장 한자 등 BMP 밖 글자(서로게이트 쌍)를 바르게 읽고, 짝 없는 글자는 대체 문자로 바꿔 캐시 저장이 실패하지 않습니다.
- `search`:
  - 검색 범위가 불완전하면(등록하지 못한 파일, 읽지 못한 파일, 다시 읽지 못한 캐시, 스캔 PDF, 20,000행에서 자른 엑셀, 들어가지 않은 하위 폴더) `not_searched`와 함께 `UNVERIFIED:`. 이전에는 일부를 못 찾고도 "VERIFIED: 0건, 다시 실행하지 마세요"였습니다.
  - 원본이 바뀐 캐시는 검색할 때 그 자리에서 다시 읽습니다(20개까지, `recovered`). 다시 읽지 못한 것은 결과 항목에 `stale: true`.
  - 검색어: 따옴표 구문, 소문자 `and`/`or`, 띄어쓰기·줄바꿈 무시(`설계기준` = `설계 기준`), NFC 정규화, `--regex`의 빈 일치 제외.
- 색인: 캐시는 최신인데 색인에 없으면(저장 전 중단, 동시 실행) 다시 등록할 때 색인을 고칩니다. Windows에서 경로 대소문자만 다른 항목은 하나로 합칩니다.
- `read --pages`: 쪽 정보가 있는 DOCX/HWPX에서도 실제 쪽으로 고릅니다(`search`의 근거 `p.12`를 그대로 읽을 수 있게). 결과에 `pages_mode`.
- `excel.py read`(닫힌 xlsx): 크기 정보(`<dimension>`)가 없는 파일을 A1 한 칸만 읽고 성공으로 응답하던 문제.

### 점검 반영 4: 실제 환경 견고성 (회사 PC 실기 확인 필요)
- 닫힌 파일을 보이지 않는 Office로 열 때 암호 창이나 열기 매크로 창이 뜨면 끝나지 않던 문제:
  Excel/Word는 틀린 암호를 넘겨 바로 "열기 암호가 걸린 문서" 오류(STOP_RETRY)가 나게 하고, 백그라운드 인스턴스의 매크로를 끕니다(`AutomationSecurity` 3).
  Excel은 읽기 전용 권장 창도 띄우지 않습니다. PowerPoint와 DRM 인증 창은 아직 막지 못합니다(남은 할 일).
- 쓰기 대상 파일을 다른 사람·프로그램이 열고 있어 읽기 전용으로 열리면, 고친 내용을 저장할 수 없으므로 쓰기 전에 멈춥니다(이전에는 저장 실패 또는 저장 안 된 채 성공 응답).
- Office가 재계산·자동 저장으로 잠깐 바빠 호출을 거부하면 COM 메시지 필터가 15초까지 기다렸다 다시 부릅니다(`recovered`). 넘으면 기존 `office_busy` 오류.
- Outlook `--folder`: `inbox\프로젝트`처럼 기본 폴더 별칭으로 시작하는 경로를 지원하고, 없는 하위 폴더의 후보를 전체 경로로 줍니다(일부만 주면 상위 폴더 전체를 검색함).
- `office.py focus`: 창 전환이 실패해도 Alt 키가 눌린 채 남지 않습니다.

### 테스트 러너 (`tools/run_win_tests.py`)
- K 그룹 추가(K1~K5): outline·장절, 쪽·단락 범위, 메모 달기·목록, 변경 추적 목록, PPT 노트. 가짜 COM 테스트로는 Word 쪽 계산(`GoTo`, `Information`)과 `Comments.Add`, 노트 개체틀이 실제와 같은지 확인할 수 없어 회사 PC 실기로 확인합니다.

## [1.3.0] - 2026-10-28

배경: 회사 PC에서 reader.py의 두 가지 쓰임새를 지원합니다.
1. 작성 중인 문서(주로 Word/Excel, 가끔 PPT)가 규정·규격서(주로 PDF, Word)에 맞는지 확인.
2. 여러 문서에서 키워드를 찾아 관련 정보를 근거와 함께 모으기. 근거는 쪽 번호, 장절 경로, 문항/조항 번호가 필수입니다.
이전 캐시는 형식마다 단위 모양이 달랐고(PDF는 쪽 하나가 단위, 표는 통째로), 장절은 Word 제목 스타일만 알아서 근거를 댈 수 없었습니다.

### 기능 (추가, 기존 인자·키 유지)
- 공통 캐시 형식(`PARSER_VERSION` 2): 모든 형식의 단위가 같은 키 `id, kind, loc, page, page_source, section, number, heading, text`를 가집니다.
  `heading`은 하위 호환용(`section`의 마지막 값). 다시 읽으면(`read`, `index`) 캐시가 새 형식으로 바뀝니다.
- 장절/번호 인식: 번호 토큰을 계열(로마 숫자, 점 번호 `1.1.1`/`1.0#`, 가나다, `1)`, `(1)`, `가)`, `(가)`, `제N장/절/조`, `Chapter N`, `Annex A`, `부록 1`, 문항 `Q3` 등)로 나누고 문서에서 처음 나온 순서로 수준을 정합니다.
  글머리 기호(`O`, `-`, 동그라미/네모 기호 등)는 번호도 제목도 아닙니다. 짧은 한 줄이고 문장 끝으로 끝나지 않을 때만 제목(오탐 방지).
  PDF는 글자 크기/굵게(본문 크기 = 가장 흔한 크기)를 제목 신호로, DOCX는 제목 스타일·개요 수준과 Word 자동 번호(numbering.xml의 `lvlText`, 시작 값, decimal/로마/가나다 형식)를 씁니다.
  자동 번호는 글 앞에 붙여(`3.2.1 세부 조건`) 검색과 근거에 나옵니다. 번호 해석이 실패하면 글의 번호 패턴으로 대신합니다. HWPX는 개요 수준과 개요 번호, HWP/TXT는 번호 패턴.
- 쪽 번호: PDF는 쪽마다 단락/줄 묶음 단위로 정확한 쪽. DOCX는 `w:lastRenderedPageBreak`(Word가 저장한 쪽 경계, `word_saved_layout`)가 있으면 그것으로,
  없으면 강제 쪽 나눔만 세어(`page_break`, 근거에 `(추정)`), 둘 다 없으면 null. PPTX는 슬라이드 번호(`slide`), HWPX는 문단 쪽 나눔이 있을 때만. XLSX/CSV/HWP/TXT는 null.
- 표는 행 단위(`kind: row`): 첫 행이 머리글이면 `항목: 내전압 | 기준: 1.5 kV 1분 | 시험방법: KS C 1234`. 병합/빈 머리글은 앞 칸을 이어 쓰고 여러 줄 칸은 공백으로 잇습니다.
  DOCX 표(`loc` `표 3 5행`), PDF 표(pymupdf `find_tables()`, `p.3 표 1 2행`, 표 영역의 글은 단락에서 뺌). 행 `number`는 첫 칸이 번호(`1`, `1.2`, `REQ-001`)일 때.
- 새 명령 `reader.py index --folder 폴더 [--depth N | --recursive] [--max-files 300]`, `index --file 파일`, `index --all`: 내용 출력 없이 캐시에 등록.
  결과 `count, indexed, updated, unchanged, failed, failures(20개), remaining, truncated, skipped_types, total_units`. `--max-files`는 새로 읽는 파일 수만 세므로 같은 명령을 다시 실행하면 이어서 등록합니다.
- `search`: 결과마다 `page, page_source, section, section_text, number, evidence`(`규격서.pdf p.12, 1. SCOPE > 1.1. Scope (1.1)`). 새 인자 `--pages 3-7`, `--section "장절 일부"`, `--max-files`.
  `--save`에 `.xlsx` 추가(머리글 굵게, 줄바꿈, 열 너비; 기본은 계속 `.md`). 열: 파일, 쪽, 장절, 번호, 위치, 내용(문장/단락), 문맥. `.md`는 항목마다 `근거:` 줄.
  저장 파일에는 `--top`과 관계없이 전체 결과(최대 5,000건)를 담습니다(이전에는 표시한 건수만). R2 키 `count`, `shown`, `truncated` 추가, 잘리면 `--save "근거.xlsx"` 명령.
- `search --file 파일/폴더`: 지금처럼 그 대상을 검색하되, 캐시가 없거나 낡은(수정 시각/크기/파서 버전) 파일만 읽어 등록하고 최신이면 캐시를 씁니다. 결과에 `indexed_now`(새로 등록한 파일, 20개까지), `indexed_now_count`, `from_cache`.
  폴더는 `--max-files`(기본 300)까지 새로 읽고 남으면 `not_indexed_remaining`과 `index --folder` 명령.
- `search`(`--file` 없음)는 캐시만 검색합니다. 이전 형식 캐시도 빈 키를 채워(`p.N`/`슬라이드 N`에서 쪽, `heading`에서 장절) 검색하고, 원본이 바뀌었거나 이전 형식이면 `stale_cache_count`와 `index --all` 명령을 줍니다.
  등록된 문서가 없으면 `error_type: not_indexed`(STOP_RETRY: 사용자에게 폴더를 물어 `--file`로 실행).
- `read`: 단위마다 `kind, page, section, number`. `--save`에 `.xlsx` 추가, `.md`에 `근거:` 줄. `read --pages`는 PDF에서 쪽 번호로 고릅니다(PDF 단위가 쪽에서 단락으로 바뀜). 새 인자 `--units`(단위 번호, 쪽 중간에서 잘렸을 때 이어 읽기 명령이 씀).
- `list`: 캐시된 파일 조회. 문서마다 `type, units, pages, sections, cached_at, parser_version, cache(current|stale), stale, missing`. 새 인자 `--name 이름일부`, `--folder 폴더`. R2 키 `count, shown, truncated`, 100개를 넘으면 `full_result_file`(CSV).
- 공통 런타임: `InputError(directive=, error_type=, fields=)`로 오류 응답에 전용 안내와 키(예: `next_commands`)를 붙일 수 있습니다.
- PDF를 읽을 때 pymupdf의 안내/경고 출력(예: pymupdf_layout 권장 문구)이 stdout으로 나와 1줄 JSON 계약을 깨지 않게 stderr로 돌립니다.

### 제거
- EML 읽기 기능 제거 (Outlook 통합 도구에서 별도로 다룰 예정). `.eml`은 다른 미지원 확장자와 같은 "지원하지 않는 형식입니다" 오류이고, `index --folder`는 `skipped_types`에 셉니다.

### 동작 변경
- `search`(`--file` 없음)는 이제 캐시만 봅니다. 이전에는 색인된 문서를 검색할 때마다 원본이 바뀌었으면 다시 읽었습니다. 최신 내용이 필요하면 `index --all` 또는 `search --file`.
- 표가 표 하나 = 단위 하나에서 행 하나 = 단위 하나로 바뀌었습니다 (`loc` `표 1` -> `표 1 2행`). PDF 단위는 쪽에서 단락으로 바뀌었습니다 (`loc`은 그대로 `p.N`).
- `read`의 PDF `unit_name`이 `페이지`에서 `단락`으로 바뀌었습니다. `--pages`의 뜻(PDF 쪽, PPT 슬라이드, 그 밖은 항목 번호)은 그대로입니다.

### 스킬
- SKILL.md reader 행: 문서 등록(`index --folder`, `read --file`), 근거 검색(`--pages`, `--section`, `--save "근거.xlsx"`), 캐시된 파일 조회(`list --name`). 근거를 `evidence`(파일 p.N, 장절 (번호))로 알리고, page가 null이면 장절과 번호로 알리는 규칙.

### 테스트
- `tests/test_reader_evidence.py`: 번호 계열 분류, 글머리, 제목 오탐 방지, 세 가지 번호 체계 중첩(PDF, 글자 크기 포함), DOCX 제목 스타일·자동 번호·쪽 나눔·저장된 쪽 경계,
  PPTX/XLSX/HWPX 장절과 쪽, 표 행 형식(병합 머리글, 여러 줄 칸), PDF 선 표, `index` 집계와 이어서 등록, 캐시 기준 검색과 자동 등록, `--pages`/`--section`, xlsx/md/csv 저장, 이전 캐시 정규화와 stale, `list` 필터, `.eml` 미지원.
- `tests/test_reader.py`: EML 샘플과 테스트 제거, 표 행 단위와 폴더 검색의 `from_cache`에 맞춤.

## [1.2.2] - 2026-10-28

사내 경량 LLM(Gemma, OpenCode + PowerShell)으로 `docs/opencode/SCENARIOS.md`를 실행한 시나리오 테스트 결과 반영 (시나리오 10, 12).

### 스킬
- 도구 위치를 `C:\CLI\office-cli`로 고정: SKILL.md와 office-cli-test SKILL.md의 명령 경로를 실제 경로로 적음 (office-cli 옆 배치 조건 없앰).

### 수정
- (시나리오 10) Sheet1만 있는 빈 통합문서에서 "Sheet9에 '테스트'라고 써줘" -> Gemma가 사용자에게 묻지 않고 Sheet9를 만들어 썼습니다.
  원인: 없는 시트 오류(`_office.missing_sheet`)의 message에 "새 시트로 만들려면 같은 명령에 --new-sheet를 붙이세요"가 있었고, 비슷한 시트가 있으면 `[후보]`에 "(새 시트 만들기: --new-sheet)"를 섞어 넣어 LLM이 그것을 골랐습니다.
  - message와 `[후보]`에는 사실만 적습니다: 없는 시트, 있는 시트 목록, 아무 데도 쓰지 않았음, 사용자에게 물어볼 것. `--new-sheet`와 명령문을 넣지 않습니다.
  - 표기만 다른 시트(대소문자·공백·_-.)가 하나면 지금처럼 `[추천]`(error_type `input`). 그 밖에는 새 error_type `needs_user_decision`(첫 단어는 `STOP_RETRY:` 그대로). 비슷한 기존 시트가 있으면 그 이름만 `[후보]`와 `candidates`에 넣습니다.
  - `agent_directive`: `STOP_RETRY: 다시 실행하지 마세요. message를 사용자에게 그대로 보여 주고, 새 시트 '<이름>'을(를) 만들지 물어보세요. 사용자가 만들라고 답한 경우에만 같은 명령에 --new-sheet를 붙여 실행하세요.` (`[후보]`가 있으면 기존 시트를 고르는 경우도 안내)
  - 공통 런타임에 `UserDecisionError`(InputError 하위 클래스, 사용자 결정이 필요한 오류) 추가.
  - SKILL.md 3절 `[후보]` 행에서 `--new-sheet`를 고를 값 목록에서 빼고 "사용자가 새 시트를 원한다고 답한 경우에만"으로 바꿨습니다.
  - `office.py transfer --to excel --sheet X`에서 X가 이미 있을 때의 `[추천] X 2`는 유지합니다. transfer는 원래 항상 새 시트에 받고(`--sheet` 생략 시에도 `Word 내용 2`처럼 비어 있는 이름), 기존 시트를 덮어쓰는 길은 `--overwrite`뿐이라 추천대로 다시 실행해도 기존 데이터가 바뀌지 않습니다.
- (시나리오 12) 저장한 적 없는 새 워드 문서(`문서1`)에 엑셀 차트를 transfer한 뒤 "방금 한 거 되돌려줘" -> `office.py undo`가 빈 워드 창을 새로 열고, 차트가 들어간 문서는 그대로 남았습니다.
  원인: 저장한 적 없는 문서의 스냅샷은 항상 "새 제목 없는 문서"로 열었습니다(1.0 설계).
  - 스냅샷 때의 문서 이름으로 열린 문서 중 저장 안 된 문서를 찾습니다(`_office.open_untitled`, 스냅샷 meta에 `untitled` 추가).
  - 워드: 현재 내용을 `pre_undo_backup`(.xml)에 보관하고, 같은 창에서 내용만 스냅샷으로 바꿉니다(flat XML 스냅샷은 `InsertXML`, 그 밖은 `InsertFile`). `restored_as: "restored_in_place"`.
  - 엑셀/PPT: 현재 상태를 `pre_undo_backup`에 보관하고 스냅샷을 새 제목 없는 문서로 연 뒤, 바뀐 문서를 저장하지 않고 닫습니다(문서 수 그대로). `restored_as: "replaced_untitled_document"`.
  - 바뀐 문서가 닫혔거나 같은 이름의 저장 안 된 문서가 둘 이상이면 고르지 않고 예전처럼 새 문서로만 엽니다(`new_untitled_document`). 이유를 `preview`에 적습니다.
  - `preview`와 `agent_directive`에 무엇을 했는지 적습니다(예: `'문서1'을(를) 같은 창에서 되돌렸습니다`). 결과에 `restored_name`(되돌린 내용이 있는 문서 이름) 추가. 기존 키(`restored`, `app`, `name`, `restored_as`, `pre_undo_backup`) 유지.

### 실기 테스트 러너 (`tools/run_win_tests.py`)
- C3(저장 안 한 새 통합문서 undo): 기대값을 "통합문서 수 +1"에서 "수 그대로, 바뀐 통합문서는 닫힘, 되돌린 통합문서의 A1이 쓰기 전 값"으로 바꿨습니다.
- C8 추가: 새 워드 문서에 쓰고 undo -> 같은 창에서 문장이 사라지고 문서 수 그대로 (시나리오 12).

### 테스트
- 없는 시트: message·`[후보]`에 `--new-sheet`와 명령문이 없음, directive가 사용자에게 묻게 함, `[추천]`(표기만 다름) 유지, `[후보]`는 실제 시트만, 시나리오 10 재현.
- 저장 안 한 문서 undo: 워드 같은 창 복원(새 문서 없음, 백업 기록), 엑셀/PPT 바꿔 열고 바뀐 문서 닫기, 문서가 닫혀 있으면 예전 동작, 같은 이름 둘이면 고르지 않음.

## [1.2.1] - 2026-10-28

회사 PC 실기 테스트(Windows 11, Office 16.0 / Microsoft 365, WinPython 3.13.13, DRM) 결과 반영과 오류 로그(R6).

### 스킬
- SKILL.md를 PowerShell 기준으로 정리: 꼭 지킬 규칙 8개(인코딩 줄, 1회 실행, 표·여러 줄·큰따옴표 값은 파일로, `$`는 작은따옴표), PowerShell에서 따옴표가 사라지는 `--values` JSON 예시를 `--file "표.csv"`로 교체, `UNVERIFIED:` 행 추가, 유니코드 화살표 제거. 도구 동작은 바뀌지 않습니다.

### 기능 (추가만, 기존 인자·키 유지)
- 오류 로그: 실패한 실행(exit 1, 인자 오류 exit 2)을 `<data 폴더>/logs/errors.txt`에 자동으로 덧붙입니다 (`CLI-dev/README.md` R6, 기준 doc-classifier `write_error_log`).
  머리줄 `===== <시각> | office-cli <버전> | Python <버전> | <OS>`, `명령:`, `결과: <error_type> | <message>`, 한도를 넘긴 복구 시도는 `복구 시도:`, 내부 오류(`ERROR:`)만 `traceback:`.
  성공과 사용자 중단(130)은 기록하지 않고, 2MB를 넘으면 `errors.old.txt`로 넘깁니다(1세대). 회전 실패와 기록 실패는 따로 삼켜 stdout 1줄 JSON에 영향이 없습니다. 로그 경로는 JSON에 넣지 않습니다.
- `_runtime.VERSION`("1.2.1"), `record_recovery_attempt()`. `retry_fs`가 한도를 넘기면 `복구 시도:`에 남깁니다. SKILL.md `metadata.version`과 같은지 테스트합니다.

### 실기 테스트 러너 (`tools/run_win_tests.py`)
- A5(두 번째 Excel 정리)가 A.xlsx를 연 사용자 Excel까지 닫던 문제. 이후 B2, B3, C1, C2, D3, E, H1, H3, I1, I2가 `DISP_E_BADINDEX`(Workbooks("A.xlsx") 없음)로 실패했습니다.
  원인: 정리 스크립트가 "실행 시작 뒤 생긴 Excel/PowerPoint의 숨은 창 또는 제목 빈 창"에 WM_CLOSE를 보냈는데, A.xlsx를 연 Excel도 러너가 띄운 것이고 Office 16에는 숨은 XLMAIN 창이 있습니다 (Office 2007에서는 드러나지 않음).
  - 문서가 남아 있는 프로세스(COM 문서 수 > 0, 제목 있는 창, 창 제목에 A.xlsx 등)는 Quit도 WM_CLOSE도 하지 않습니다.
  - A5는 Excel을 띄우기 전후의 PID를 비교해 자기가 띄운 PID만 끝냅니다. 종료 정리는 실행 시작 뒤 생긴 빈 프로세스만, 실행 전부터 있던 Office는 문서가 없어도 끝내지 않습니다.
  - 같은 인스턴스가 두 번 잡히면 건너뛰고, Office 16의 빈 창 제목(`Excel`, `PowerPoint`)도 인식합니다.
  - 판단 규칙은 `tools/win_test_rules.py`로 분리해 오프라인 테스트(`tests/test_win_runner.py`, 가짜 Office/pywin32로 회사 PC 상황 재현)를 붙였습니다.
- JIT 복구: 그룹 시작 전과 문서 조회 때 A.xlsx/A.docx/A.pptx가 닫혀 있으면 준비 단계와 같은 방법(`os.startfile`)으로 다시 열고 `[WARN] A.xlsx가 닫혀 있어 다시 열었습니다`를 남깁니다.
  다시 열지 못하면 그 그룹의 남은 항목을 "테스트 문서 A.xlsx가 닫혀 다시 열지 못함" 하나의 이유로 바로 ERROR 처리합니다.
- `com_retry`가 '그 이름의 항목 없음'(`DISP_E_BADINDEX`)을 재시도하지 않습니다. 중첩된 재시도(최대 10x10x10회)로 멈춘 것처럼 보여 Ctrl+C로 끝내야 했습니다.
- 열린 문서 조회가 첫 번째 Excel 인스턴스만 보지 않고 모든 인스턴스를 찾습니다.
- `errors_*.txt`에 "러너 자동 복구와 창/프로세스 정리"(다시 연 문서, 닫은 문서, Quit/WM_CLOSE를 보낸 PID와 창, 보호한 PID)와 "도구 오류 로그"(이번 실행 중 `<테스트 폴더>/_data/logs/errors.txt`에 추가된 부분) 절을 넣습니다. 종료 정리 뒤에 한 번 더 써서 정리 기록까지 담습니다.

### 테스트
- (회사 PC 2차 실기 후 보완) 테스트용 docx를 폴더 템플릿으로 만들 때, 폴더에서 빠진 파일(회사 PC: `docProps/thumbnail.jpeg`)을 가리키는 관계는 빼고 엽니다 (`tests/sample_docs.py`). 도구 동작은 바뀌지 않습니다.
- DRM PC에서 python-pptx 기본 템플릿(`site-packages/pptx/templates/default.pptx`)이 zip이 아니라(암호화) `pptx.Presentation()`이 실패하던 문제(1 FAILED + 12 ERROR).
  템플릿이 필요한 테스트만 "python-pptx 기본 템플릿이 손상됨 (DRM 추정)"으로 건너뛰고, reader 샘플(docx/xlsx/pdf/eml/hwpx)은 형식별로 만들어 나머지 테스트는 그대로 돕니다.
  python-docx는 기본 .docx가 손상돼 있으면 폴더 템플릿(`docx/templates/default-docx-template`)으로 만듭니다 (`tests/sample_docs.py`). 도구 코드는 인자 없는 `Presentation()`/`Document()`를 쓰지 않습니다.

## [1.2.0] - 2026-10-28

저장소 지침(AGENTS.md 2026-10-28) 반영 1차: 사용자 확인 없이 덮어쓰거나 추측하던 동작을 멈춥니다.
근거 `docs/history/COMPLIANCE_REVIEW_2026-10-28.md` 2절, 구현 `CLI-dev/README.md` R8(`[추천]`/`[후보]`).

### 기능 (추가만, 기존 인자·키 유지)
- 후보가 둘 이상이라 도구가 고를 수 없을 때의 오류 형식: `error_type: "ambiguous"`, `candidates`(최대 10개), message 끝 `[후보] a, b, c`,
  directive `STOP_RETRY: 후보가 여러 개라 ... 재시도하지 말고 [후보]를 사용자에게 보여주고 고르게 한 뒤 ...`. SKILL.md 결과 처리 표에 추가.
- `excel.py write --new-sheet`: 없는 시트를 새로 만들어 쓸 때만 붙입니다. 이미 있는 이름에 붙이면 만들지 않고 오류.
- `word.py write --occurrence N`, `office.py transfer --occurrence N`: `--anchor`가 여러 단락에 있을 때 몇 번째 단락인지 (오류의 `[후보]` 번호).
- `office.py transfer --to excel --overwrite`: `--sheet`로 준 시트가 이미 있을 때 그 시트에 덮어쓰기.

### 수정 (동작 변경)
- `excel.py write --sheet`(또는 `--range '시트'!A1`)에 없는 시트 이름을 주면 그 이름으로 새 시트를 조용히 만들던 문제. 오타가 새 시트가 되었습니다.
  이제 만들지 않고 오류를 냅니다.
  - 대소문자·공백·`_-.`만 다른 시트가 하나면 `[추천]` (예: `Sheet 1` -> `Sheet1`).
  - 비슷하지만 다른 이름이면 `[후보]`로 사용자에게 고르게 합니다 (예: `Sheet2`를 `Sheet1`의 오타로 단정해 덮어쓰지 않음).
  - 비슷한 것이 없으면 기존 시트 목록과 `--new-sheet` 안내.
- `office.py transfer --to excel --sheet 이름`에 이미 있는 시트를 주면 A1부터 덮어쓰던 문제. 덮어쓰지 않고 비어 있는 이름을 `[추천]`합니다(`--overwrite`로 덮어쓰기). `--sheet`를 생략하면 지금처럼 새 시트(`Word 내용 2` 등).
- `word.py write --anchor`가 여러 단락에 맞으면 첫 번째를 고르고, `--pos replace`면 그 단락을 지우던 문제. 아무것도 바꾸지 않고 `[후보]`(번호와 단락 앞부분 30자)로 오류를 냅니다. 같은 단락에 여러 번 나오면 한 곳으로 셉니다.
  `office.py transfer --anchor`도 같습니다. 이 오류는 차트 전달 방식을 바꿔도 같으므로 다음 방식으로 넘어가지 않습니다.
- `--anchor`를 찾지 못하면 문서에서 가장 비슷한 문구(그대로 다시 찾을 수 있는 원문 일부)를 안내합니다.
  공백·대소문자만 다를 때, 또는 앵커 글자 사이에 2자 이하만 끼어들고(`2026 계획` -> `2026년 계획`) `--pos replace`가 아닐 때만 `[추천]`.
  글자가 바뀌거나 빠진 문구(`Q2` -> `Q1`)는 뜻이 다를 수 있어 `[추천]` 없이 "비슷한 문구: ..."로만 알립니다.

### 실기 테스트 러너
- I그룹이 미리 만들지 않은 `I_DATA`, `I_BIG` 시트에 쓰던 첫 write에 `--new-sheet` 추가 (1.2.0 동작 변경으로 실패하던 문제).

### 다음으로 미룸
- `--save`(CSV/MD/텍스트)와 COM 내보내기가 기존 출력 파일을 덮어쓰는 문제는 공통 런타임 반영(R4) 때 함께 처리 (남은 할 일).
- 열린 문서 이름(`--target 보고.xlsx`, `office.py focus --name`)이 여러 문서에 맞으면(다른 폴더의 같은 파일명, 다른 Excel 창의 `Book1`) 첫 번째를 고르던 문제. `[후보]`(전체 경로)로 오류를 냅니다. 전체 경로를 주면 그 문서로 정해집니다.

## [1.1.0] - 2026-10-27

### 구조 (이름 변경, 동작 불변)
- 도구 이름을 `atomic-cli`에서 `office-cli`로 바꿨습니다. 진입점 6개로 합치면서 "작은 도구 여러 개"라는 뜻이 더 이상 맞지 않기 때문입니다.
  - 폴더 `CLI/atomic-cli` -> `CLI/office-cli`, 스킬 이름(SKILL.md `name`) `office-cli`. OpenCode 스킬 폴더도 `office-cli`로 바꿔 배치하세요.
  - 환경변수 `ATOMIC_DATA_DIR` / `ATOMIC_CACHE_DIR` / `ATOMIC_VERBOSE` -> `OFFICE_CLI_DATA_DIR` / `OFFICE_CLI_CACHE_DIR` / `OFFICE_CLI_VERBOSE`. 예전 이름은 읽지 않습니다.
  - 실기 테스트 기본 폴더 `C:\atomic_test` -> `C:\office_test`.
- 내부 함수 `_runtime.atomic_write_text`(중단돼도 기존 파일이 깨지지 않는 저장)를 `safe_write_text`로 바꿨습니다. 코드에 `atomic` 이름이 남지 않습니다.
- 명령, 인자, JSON 키, `config.json`, `cache`/`data` 폴더 구조는 그대로입니다.

### 테스트 러너 (`tools/run_win_tests.py`)
- 분석용 오류 로그 `errors_<시각>.txt`를 `report_*.txt`와 함께 항상 만듭니다. 환경(Windows, Python, 라이브러리, Office 버전과 PDF 가능 여부, `config.json`, 러너 옵션)과, 실패/오류 항목마다 실행한 도구 명령 전체, 종료 코드, 걸린 시간, stdout JSON 전체, stderr 전체, traceback, 그 항목의 진행 로그를 담습니다. 기존 보고서는 JSON 400자, stderr 600자에서 잘려 원인을 알기 어려웠습니다.
- 준비 단계 실패(테스트 파일 만들기, 창 열기), 러너 자체 오류, Ctrl+C 중단에도 그 시점까지의 오류 로그를 남깁니다.
- `--agent`: 에이전트(OpenCode 스킬)용 실행. `--auto`와 같고, 진행 로그는 stderr로, stdout에는 결과 JSON 1줄(`summary`, `failed`, `report_file`, `error_log`, `agent_directive`)만 씁니다. 준비 실패는 exit 1, 중단은 130.
- 실기 테스트 스킬 `tools/office-cli-test/SKILL.md` 추가. 경량 모델이 사용자 확인 후 러너를 1회 실행하고 결과와 오류 로그 경로를 보고합니다.

## [1.0.1] - 2026-10-27

Windows 실기 테스트(집 PC: Windows 11 Pro 26200, Python 3.13.13, Office 2007 32비트) 결과 반영. 최종 65개 중 PASS 63, ENV-SKIP 2(PDF 추가 기능 없음).

### 기능 (추가만, 기존 인자·키 유지)
- `excel.py / word.py / ppt.py info --detail`: 결과가 제대로 보이는지 도구로 확인. PPT는 도형별 위치·크기·글자 크기·목록 종류와 경고(슬라이드 밖, 글자 넘침, 겹침, `--slides`로 범위 지정), 엑셀은 차트 종류·계열 수·위치와 데이터를 가리는 차트 경고, 워드는 그림 크기(쪽 너비 대비)와 제목 수준. 경고 수는 `layout_warnings`.
- `office.py windows`가 실행 중인 Office 버전과 PDF 내보내기 가능 여부(`office`)를 알려 줍니다. PDF 기능이 없는 Office(추가 기능 없는 2007 등)에서는 `export --save *.pdf`가 시도하지 않고 바로 STOP_RETRY로 안내합니다.
- 되돌리기 스냅샷에 문서 식별값(파일 속 만든 시각 `docProps/core.xml`)을 저장합니다. 같은 경로에 다른 문서를 새로 만들면 이전 문서의 스냅샷은 undo 목록에서 빠지고 정리됩니다(Office 저장으로는 바뀌지 않음).

### 수정
- `excel.py chart`를 `--at` 없이 만들면 데이터 범위 바로 오른쪽에 놓여 시트의 다른 데이터를 가리던 문제. 사용 범위 오른쪽에 놓습니다 (실기 J2에서 발견).
- `ppt.py export --save *.pdf`가 항상 내부 오류(`TypeError: The Python instance can not be converted to a COM object`)를 내던 문제. `ExportAsFixedFormat`의 개체 인자 `PrintRange`를 생략하면 pywin32가 인자 변환에서 실패하므로 `PrintRange=None`을 명시합니다.
- PowerPoint 2007(12.x)에 PDF 추가 기능이 없을 때 E_FAIL이 내부 오류로 나가던 문제. `.pptx`로 저장하라는 STOP_RETRY 입력 오류로 안내합니다. 다른 버전의 E_FAIL은 그대로 Office 오류로 보고합니다.
- Word 2007에 없는 `SaveAs2` 호출(새 파일 만들기, `word.py export --save *.docx`, undo의 Flat XML 복원)이 AttributeError로 실패하던 문제. 공통 `_office.word_save_as()`가 `SaveAs`로 대체합니다.
- `ppt.py write`의 `-` 글머리 목록이 "1. 2." 번호 목록으로 나오던 문제. 글머리 표시만 켜고 종류를 정하지 않아 PowerPoint 2007 글상자가 번호 목록으로 만들었습니다. 종류(글머리표/번호)를 명시하고, 마크다운과 `--text`의 `1.` 번호 목록도 구분합니다.
- `excel.py chart`의 분산형 종류가 틀리던 문제. `--type scatter`가 표식만이 아니라 선+표식(74)으로, `scatter_lines`(숫자 x축 `line` 자동 전환 포함)가 표식 없는 선(75) 대신 선+표식으로 그려졌습니다. `scatter` 상수를 -4169로 고치고, 종류를 데이터 지정 앞뒤로 모두 정합니다(앞: 첫 열을 x값으로, 뒤: Excel 2007이 바꾼 세부 종류 복원).
- 열린 워드 문서를 읽을 때 본문 그림 자리가 글자 `/`(Word 2007, 새 버전은 `\x01`)로 나오던 문제. `word.py read`와 `office.py transfer --from word`에서 그림 위치의 글자만 `[그림]`으로 바꿉니다(본문의 진짜 `/`는 유지).

### 테스트 러너 (`tools/run_win_tests.py`)
- 테스트 파일 준비가 Word 2007에서 `SaveAs2`로 멈추던 문제.
- 더블클릭으로 연 Excel은 포커스를 잃기 전까지 ROT에 등록되지 않아(`GetActiveObject` 실패, MS KB 238610) "A.xlsx가 열리지 않았습니다"로 중단되던 문제. 도구와 같은 창 핸들 바인딩(`_office.find_apps`)을 씁니다.
- 이전 실행의 스냅샷이 남아 같은 경로로 다시 만든 A 파일에 undo가 이어져 C7이 두 번째 실행부터 실패하던 문제. 도구 캐시/data를 테스트 폴더의 `_cache`(실행마다 비움), `_data`로 분리합니다.
- PDF 기능이 없는 PC의 D1, D4를 FAIL 대신 `ENV-SKIP`으로 기록합니다.
- I그룹 27개 추가 (모두 자동 판정): 엑셀 CSV/`--values`/`--append` 쓰기, `--search`·`--save`, 300행 잘림, 차트 종류·목록·삭제, 조건부 서식·숫자 서식·지우기, 시트 이름변경·복사·삭제, 범위 복사·숨김·틀 고정, 워드 그림·CSV 표·앵커 앞/바꾸기·제목 수준·쪽 나누기·목차·치환(굵게, `--map`, 바닥글), PPT 그림·카드·마크다운·긴 표·슬라이드 관리·표/그룹 치환·`.pptx` 사본, 엑셀 범위 -> 워드/PPT 표, 워드 표 -> 엑셀, redraw 강제, `undo --list/--pick`, 닫힌 파일 undo, 창 전환 시 최대화 유지. 집 PC(Office 2007): 60개 중 PASS 60(ENV-SKIP 2 제외).
- J그룹 3개 추가: `info --detail`로 PPT 전체 슬라이드, 엑셀 차트, 워드 그림의 배치 경고가 없는지 판정.
- 끝나면 A.pptx 모든 슬라이드와 A.xlsx 모든 차트를 `shots\`에 PNG로 저장하고, 테스트 폴더 문서와 실행 중 생긴 제목 없는 문서를 저장하지 않고 닫습니다(`--keep-open`으로 유지). 문서가 남지 않은 앱은 끝내고, `Quit`을 받고도 남는 빈 Excel/PowerPoint(이 실행 뒤에 생긴 것만)는 창 닫기로 정리합니다. 정리는 러너가 COM 참조를 놓은 뒤 별도 프로세스에서 합니다.

## [1.0.0] - 2026-10-27

첫 배포. 기존 마이크로 CLI 34개를 진입점 6개로 다시 설계했습니다. 이전 버전과의 CLI 호환은 없습니다.

### 구조
- 진입점 6개: `excel.py`, `word.py`, `ppt.py`, `outlook.py`, `reader.py`, `office.py`. 공통 모듈 2개: `_runtime.py`(응답 계약, 경로, 입출력), `_office.py`(COM 세션, 스냅샷, 색/테마).
- 저장소 ACI 계약 준수: stdout 1줄 JSON, exit 0/1/2/130, `VERIFIED:` / `STOP_RETRY:` / `ERROR:`, 인자 오류도 JSON.
- 닫힌 파일 읽기는 라이브러리(openpyxl, python-docx, python-pptx, pymupdf, olefile)로, Office를 띄우지 않습니다. 열린 문서와 쓰기는 COM입니다.
- 결과에 `target_resolved`(실제 작업한 문서), `next_commands`(다음 명령), 이름 오타 시 `[추천]`을 넣습니다.

### 기능
- 화면 규칙: 읽기는 사용자 화면(활성 시트, 스크롤, 슬라이드)을 움직이지 않습니다. 쓰기는 고친 시트/슬라이드로만 이동합니다(`--no-focus`로 끔).
- 쓰기 전 자동 스냅샷과 `office.py undo`(`--list`, `--pick`). 원본은 저장하지 않으며 복원 직전 상태를 백업합니다.
- 차트 전달 3방식(`export` / `redraw` / `relay`)을 자동으로 이어서 시도합니다(DRM으로 PNG가 암호화되는 환경 대응). `--probe`로 이 PC에서 되는 방식을 저장합니다.
- 엑셀: 범위 수식(열마다 수식), 머리글, 이어 쓰기, 테마, 실제 조건부 서식, 숫자 x축 꺾은선의 분산형 자동 전환.
- 워드: 마크다운(제목/목록/표/굵게)을 기본 스타일로 변환, 표는 ConvertToTable로 빠르게, 그림 비율 유지, 앵커 단락 앞/뒤/교체.
- PPT: 마크다운 제목마다 새 슬라이드, 긴 표는 머리글을 반복해 나눔, 그림은 남은 공간에 비율 유지, 카드.
- reader: 읽은 문서 자동 색인과 전체 검색(AND/OR/정규식), `--save`로 md/csv.
- outlook: 기본은 작성 창만 표시, `--send`는 명시 요청 시에만 수신자 확인 후 발송.

### 수정 (기존 코드 대비)
- 파일 대상 작업이 사용자 Excel/Word 창을 숨기던 문제 (전용 인스턴스 사용).
- 읽기 작업이 없는 시트/파일을 만들고 저장하던 문제, PPT 읽기만 해도 마지막 슬라이드로 이동하던 문제.
- Word 스냅샷이 원본을 저장하거나 경로를 바꾸던 문제, 활성 PPT는 스냅샷이 없던 문제, undo가 실패해도 성공으로 보고하던 문제.
- 한 행 범위 처리 오류, 범위 이동이 겹치면 데이터가 지워지던 문제, 치환이 수식을 덮어쓰던 문제.
- 워드 일반 텍스트 쓰기가 항상 실패하던 문제, 앵커 뒤 삽입이 단락을 쪼개던 문제, 빈 문자열 치환이 무시되던 문제.
- HWP 제어 문자 해석 오류, PPT 슬라이드 순서와 단어 쪼개짐, EML 한글 헤더와 전달 메일 본문 섞임.
- Outlook 검색 필터 구문 오류, 첨부 덮어쓰기, 서명 사라짐.

## 남은 할 일

### 1.0 제거 기능 (필요하면 다시 추가할 후보)
- 스타일 DNA 추출 (`office_theme --action extract`). 테마 프리셋은 유지.
- 검색 결과를 Office로 바로 넣기 (`search --to`). 현재는 `reader.py search --save` 후 `write --file`.
- 검색어가 파일명과 같으면 파일을 반환하던 처리.
- 샌드박스 가상 데스크톱 전환 (exebox, `SetThreadDesktop`).
- `clean fix-env`, JIT 자가복구의 "유령 프로세스 정리".
- 새 문서 만들기 / 파일 열기 (`office_window --new/--open`).
- EML 읽기·마스터 JSON 빌드·스레드는 Outlook 통합 도구에서 별도로 다룸 (reader.py에서는 1.3.0에서 제거).
- 워드 그림/도형 -> PPT 전달 (`office_bridge` word->ppt), PPT -> 워드 전달.
- Outlook 날짜 범위 검색 (`--since`).

### 1.4.0 점검에서 남긴 것
- PowerPoint 닫힌 파일의 열기 암호, DRM 인증 창이 뜨는 문서: 백그라운드에서 멈출 수 있음 (감시 시간 초과 후 정리 필요).
- 엑셀: 열린 문서의 열 전체 범위(`A:A`)를 사용 범위로 자르기 (닫힌 파일·CSV는 1.4.5), 조건부 강조 `포함:`의 상대참조(활성 셀 기준), 같은 이름 차트가 여러 시트에 있을 때 `[후보]`, `--map` 연쇄 치환, 보호된 시트 사전 확인, 잘못된 시트 이름 검사 전에 시트가 만들어지는 문제.
- 워드: 열린 문서의 자동 번호 제목(`ListFormat.ListString`)이 `outline`/`--section`에 나오지 않음(닫힌 파일과 번호 체계 다름), 목차(TOC) 안의 앵커 후보 구별, 머리글 안 글상자 치환. (`replace --find`만 준 경우는 1.4.10)
- reader: HWP(바이너리) 개요 번호, PDF `find_tables` 성능(선이 있는 쪽만), 색인 저장 일괄화, BOM 없는 UTF-16 판별.
- transfer relay 실패 시 넣은 빈 단락 정리.

### 확인 필요
- Windows 실기 검증: 회사 PC(Office 16)에서 1.4.1 실기 러너 71/71 통과. 남은 것: 1.4.5 쓰기 변경(write 덮어쓰기 확인, transfer 글자 유지)의 실기 확인, 눈 확인 항목.
- 실제 HWP 5.0 파일과 DRM 문서로 `reader.py` 검증. 실제 규격서(PDF/Word/HWPX)로 장절·쪽 인식 확인 (1.3.0).
- 회사 PC에서 `office.py transfer --probe` 결과 확인. (집 PC Office 2007: export, redraw 성공 / relay는 "워드에 붙여넣기 안 됨"으로 실패)
- 남은 Office 2007 실기 확인: `.xls/.doc/.ppt`(OOXML 아님)는 문서 식별값이 없어 같은 경로 새 파일 판별을 하지 않습니다.

### 저장소 지침 반영 (AGENTS.md, 2026-10-28)
상세 근거: `docs/history/COMPLIANCE_REVIEW_2026-10-28.md`. 공통 키·함수는 `CLI-dev/README.md` "공통 런타임 구현 규칙"(R1~R9)대로 구현. 우선순위 순.
- [ ] `recovered` 나머지: transfer 방식 대체, 오래된 스냅샷 정리 (잠금·기존 파일 새 이름 저장, 낡은 캐시는 1.4.0)
- [ ] `require_lib()`: openpyxl(excel.py:215, reader.py:382), python-docx(reader.py:291, office.py:465), python-pptx(reader.py:329,346) 실패 시 STOP_RETRY + `python -m pip install ...` (회귀 테스트)
- [ ] `remove_tree` 추가, `os.remove`(_office.py:314, _runtime.py:311, word.py:593)에 적용. (1.4.0: undo 복원 복사는 `retry_fs`, retry_fs는 일시적 오류만 재시도 완료)
- [ ] COM 내보내기(excel.py, word.py, ppt.py export)의 잠금 대비 `save_with_lock_fallback`. (1.4.0: 기존 출력 파일은 새 이름 저장 + `recovered` 완료)
- [ ] `--save`로 기존 파일을 덮어쓰기를 원할 때의 `--overwrite` 인자 (기존 파일 새 이름 저장은 1.4.0)
- [ ] `--dry-run`: excel replace/range clear·move/sheet delete/chart delete --all, word replace/write --pos replace, ppt replace/slide delete·clear. (1.4.0: 스냅샷이 없는 쓰기는 undo가 건너뛰지 않음, `undo --list`의 `undo_available: false`. 쓰기 결과에도 표시하는 것은 남음)
- [ ] 빈 결과 `count: 0` + "찾은 항목 없음": excel read 빈 범위, word/ppt read 0단위, outlook list, office windows, info --all, chart list, undo --list
- [ ] 잘리는 결과에 `truncated` + 전체 결과 파일 경로: excel read --search(50), excel replace cells(30), undo --list(20), outlook list/search. `body_truncated` 옆에 `truncated`
- [ ] next_commands: excel write의 차트 명령 자리표시 `--title 차트` 제거, word·ppt의 transfer 명령에 `--from-target`
- [ ] tools/run_win_tests.py `--agent` 모드: JsonArgumentParser와 최상위 예외 처리, 로그 태그 `[준비]` `[주의]` `[확인]` -> `[INFO]` `[WARN]`
- [ ] 인자 이름 정리 (베타라 별칭 없이 바꿈): `-r`/`-c`/`-n` 충돌, `--limit` 대 `--top`, `--pages` 대 `--slides`
