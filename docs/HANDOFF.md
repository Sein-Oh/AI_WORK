# 작업 인계 (2026-11-02, 클라우드 세션 -> 회사 PC)

회사 PC에서 저장소 파일을 받은 뒤, 새 Claude Code(또는 OpenCode) 대화를 열고 맨 아래 "붙여넣을 첫 메시지"를 보내면 이어서 진행할 수 있습니다.
이 파일은 작업이 이어지면 최신 상태로 고치거나 지웁니다.

## 1. 지금 상태
- main 기준 도구 버전: **ml-kit 2.9.58**, **office-cli 1.4.62**
- 저장소 규칙은 `AGENTS.md` 한 곳 (Claude Code는 `CLAUDE.md`가 불러옴). 특히:
  - 답변은 한국어, 이모지·장식 기호 금지 (CP949 안전 문자만)
  - 베타 단계: 이전 버전 호환 코드를 넣지 않음. CLI 인자·JSON 키는 바꿔도 되고, 바꾸면 SKILL.md·README·CHANGELOG·테스트를 같은 커밋에서
  - 리팩터링과 기능 변경은 별도 커밋/PR, 버그 수정에는 회귀 테스트
  - `archive/`는 수정하지 않음
- 배포 파일에 사용자 이름이 든 전체 경로를 넣지 않음 (성능예측표 등은 `<데이터>_ml_data\vN`처럼 끝부분만)
- 실기 로그는 사용자가 GitHub 웹으로 `docs/test_logs/`에 올림 (git 명령으로는 올리지 않음, `.gitignore` 처리). 프로젝트 마무리 때 저장소 이력에서 정리 예정
- 클라우드 점검: 17~25차 점검(검토 -> 수정 -> 재검사 반복) 완료 (2026-11-03). 마지막 차수의 회귀·높음 결함은 모두 수정, 경미·중간은 각 도구 CHANGELOG "남은 할 일"
- 클라우드 테스트: 전체 pytest 1549 passed, 1 skipped (ml-kit 2.9.36, office-cli 1.4.32 기준). 회사 PC 수는 `run_all_tests.py`의 pytest 결과로 확인

## 2. 최근 바뀐 것 (클라우드 전체 점검 반영, 실기에서 볼 것)
| 버전 | 내용 |
|---|---|
| office-cli 1.4.10 | `ok()`가 `verified=False`를 거부하고 저장 파일은 형식별로 다시 열어 확인. `excel replace`가 `--range`의 시트에 씀. 결과가 ""인 수식 칸도 덮어쓰기 확인. `ppt write --title`로 다른 제목을 바꾸려면 `--overwrite`. 파일 이름만 준 쓰기 대상은 새로 만들지 않음(전체 경로만). reader 잘림·글자 없는 PDF 쪽을 알림. 오류 로그에서 본문·받는 사람 등 내용 값은 `<N자 생략>` |
| office-cli 1.4.11 | 문서 색인 형식이 바뀜(PARSER_VERSION 6): 처음 검색 때 도구가 `index --all`을 안내하면 그대로 실행 |
| office-cli 1.4.12 | 문서 색인 형식이 다시 바뀜(PARSER_VERSION 7, PDF 스캔 쪽 판정): 처음 검색 때 도구가 `index --all`을 안내하면 그대로 실행. 동시에 실행한 `index`도 색인 항목을 잃지 않음 |
| office-cli 1.4.13 | 문서 색인 형식이 다시 바뀜(PARSER_VERSION 8, 도장·Bates 번호만 남은 스캔 쪽도 판정): 처음 검색 때 도구가 `index --all`을 안내하면 그대로 실행. 입력 오류로 멈춘 쓰기는 문서(새 시트·서식 포함)를 바꾸지 않고 되돌리기 지점도 밀어내지 않음. 색인을 잠깐 못 읽어도 비우지 않음. 모든 쪽이 스캔인 PDF의 `read`는 `UNVERIFIED`(OCR 필요). 첨부 저장 실패는 `UNVERIFIED` |
| office-cli 1.4.14 | 바꾼 것 없이 성공한 쓰기(0곳 바꿈, 차트 0개 삭제, 이미 있는 시트 add)는 되돌리기 지점을 남기지 않음(다음 `undo`가 실제로 바꾼 작업을 되돌림). 엑셀에서 쓸 수 없는 시트 이름(`/ \ ? * [ ] :`, 31자 초과)은 시트를 만들기 전에 STOP_RETRY와 `[추천]` 이름. 스캔 PDF를 `read --save`하거나 스캔 쪽만 고르면 `UNVERIFIED`. 깨진 색인 파일도 다시 만듦 |
| office-cli 1.4.15 | 엑셀이 예약한 시트 이름 `History`(대소문자 무관)도 시트를 만들기 전에 STOP_RETRY와 `[추천] History_`. `transfer --sheet`의 `[추천]`이 받는 문서의 있는 시트와 겹치지 않음. 바꾼 것 없이 성공한 명령(같은 글로 바꾸는 `replace`, 같은 자리로 `slide move`, 숨긴 것 없는 `unhide`, 고정 없는 `unfreeze`, 빈 범위 `clear` 등)은 "0곳/0칸 바뀜" 또는 "바뀐 것 없음"으로 답하고 되돌리기 지점을 남기지 않음. 휘발성 수식(NOW 등)이 든 닫힌 파일도 바꾼 것이 없으면 다시 저장하지 않음. `ppt replace --find`만 주면 STOP_RETRY |
| office-cli 1.4.16 | `ppt replace`로 대소문자만 다른 글을 통일하는 치환(`iphone` -> `iPhone`)을 실제로 바꿈. 새 파일로 `transfer --sheet Sheet1`처럼 기본 시트 이름을 주면 그 시트에 받음. `range hide`/`unhide`와 `clear`(전체·서식)는 늘 바꾼 것으로 봄(되돌리기 지점을 잃지 않게) |
| office-cli 1.4.17 | `excel`/`word`/`ppt replace`의 찾을 글이 빈 문자열이면 STOP_RETRY(빈 칸 채우기는 `write`). 기본 시트가 여럿인 새 파일로 `transfer --sheet Sheet1`도 그 시트에 받음. 대소문자만 다른 치환(`iphone` -> `iPhone`)에서 이미 바꿀 글인 곳은 바꾼 곳으로 세지 않음(`unchanged_same_text`) |
| office-cli 1.4.18 | 이모지 같은 BMP 밖 글자 뒤의 글도 `ppt replace`가 모두 바꿈. Word 찾기에서 `^`를 글자 그대로 찾음(`감사합니다^^`, `10^6`) |
| office-cli 1.4.19 | Word 치환에서 바꿀 글이 255자를 넘어도 바꾼 뒤 확인이 오류로 끝나지 않음(Word 찾기 대신 직접 셈). 찾을 글·앵커 길이는 `^`를 2자로 셈 |
| office-cli 1.4.20 | `transfer --to word,ppt`는 무엇이든 넣기 전에 받는 문서를 모두 확인: 한 앱만 열려 있으면 아무것도 바꾸지 않고 STOP_RETRY와 `[추천]`(열린 앱만 `--to word` 등, 그대로 1회 실행). 넣은 뒤 한 앱만 실패하면 `UNVERIFIED`(`succeeded`/`failed`, 다시 실행하지 않음). 열린 문서가 없을 때 안내는 `--to-target`/`--from-target`(없는 `--target`을 안내하지 않음). `excel chart`의 `next_commands`는 열린 앱만 `--to`에 넣음 |
| office-cli 1.4.21 | 시트를 주지 않은 차트 찾기(`transfer --chart`, `chart --action delete --name`)는 이름이 같은 차트를 제목보다 우선하고, 그래도 여러 시트에 있으면 하나를 고르지 않고 `[후보] 시트!이름`(그 형식을 그대로 받음). `excel chart`의 `next_commands`에 `--sheet`. 이름이 같은 열린 문서가 여럿일 때 안내는 그 명령의 옵션(`--to-target`/`--from-target`/`--name`) |
| office-cli 1.4.22 | 차트 `[후보]`의 `시트!이름`을 `--sheet`와 함께 줘도, 차트 이름에 `!`가 있어도 찾음. SKILL의 문서 `[후보]` 안내가 message가 말한 옵션(`--to-target` 등)을 따르게 함 |
| office-cli 1.4.23 | outlook 검색어의 `%`·`_`를 글자 그대로 찾음(`견적서_v2`가 0건이던 문제). 기본 파일 이름의 `/ : ?` 등을 `_`로(차트 `1/4분기 매출` 전달 실패). 차트 `'시트'!이름` 형식, 글자 그대로 `!`가 든 차트 이름과의 `[후보]`, 한 시트의 같은 이름 차트는 `[후보] 시트!#번호`. 손상된 HWP/HWPX는 STOP_RETRY. `undo --pick 0` 거절, outlook `--limit` 상한 200 |
| office-cli 1.4.24 | 차트 `[후보]`는 그 시트의 모든 차트로 이름이 겹치는지 보고 `#번호`로(후보마다 차트 하나). FileHeader 없는 OLE `.hwp`는 STOP_RETRY "HWP 5.0 문서가 아닙니다". `excel style --highlight` 값의 `"`·`~ ? *`를 글자 그대로. `excel write --append --formula`는 상대 참조를 실제 쓴 행 기준으로. 마크다운 펜스 코드 블록(백틱 3개)을 줄마다 글자 그대로, 표 정렬 줄 `:-`, 칸 글자 `\|`. `ppt write --cards`는 공간이 모자라면 새 슬라이드에, 슬라이드 밖 도형은 UNVERIFIED. reader `--units 1-999999999`가 멈추지 않음 |
| office-cli 1.4.25 | `.md` 코드 블록 줄의 `__init__`, `x**2`를 굵게로 읽지 않고 그대로 넣음(Word·PowerPoint). 줄 첫머리 인라인 코드 ```` ```x``` ````는 코드 블록 시작이 아님. 코드 블록으로 시작하는 파일도 다시 읽기 확인 |
| office-cli 1.4.26 | `ppt write`가 코드 줄·제목의 `__`·`**`(`def __init__`, `## __init__ 설명`)로 거짓 UNVERIFIED가 되지 않음. `word write`는 `---`로 시작하는 `.md`, 첫 표 행의 `\|`도 다시 읽기 확인. `ppt write --text`의 펜스 코드 줄은 글자 그대로, `###` 제목의 `**`는 굵게 표시로 뺌. `excel write --file .md`는 첫 번째 표만(코드 블록·두 번째 표 제외). 0건 결과(`office windows`, `undo --list`, `info --all`, `chart --action list`)는 count 0과 "0건". `reader clear --dry-run`이 20개를 넘으면 `full_result_file`. 오류 로그에서 `--section`, excel `--highlight` 값을 가림 |
| office-cli 1.4.27 | 제목 굵게 표시 규칙을 word·ppt 모든 수준(슬라이드 제목, `--title` 포함)에서 하나로(`**x**`, `__x__`는 표시를 뺀 글). 단어 안의 `__`(`report__2024__final.xlsx`)는 그대로. `2026. 11. 3.`처럼 4자리 이상 숫자는 번호 목록이 아니고, 1이 아닌 번호로 시작하는 목록은 사용자 번호를 글로 둔 단락. 오류 로그 message에서 가린 인자 값과 문서에서 가져온 후보·문구를 뺌. `undo --list`(되돌릴 수 없는 작업만), `ppt slide --action list`(0장), `outlook list/search`(0통)는 `count`와 `VERIFIED: 0건`. `reader clear`의 `applied`·완료 CSV는 실제로 지운 것만(`removed` 열) |
| office-cli 1.4.28 | 오류 로그에 문서 글이 남지 않게 구조를 바꿈: 문서 글(칸 값, 제목, 노트, 시트·문서·파일·메일 폴더 이름 목록, 후보 단락)을 message에 넣는 오류는 그 글을 뺀 `log_message`를 주고, 표식 자르기(`사용 가능한 <공백 포함 종류>:`, `후보:`, `추천:` 추가)는 대비로 남김. 가린 값의 10자 이상 앞부분과 traceback도 가림, 1~2자 값은 따옴표 안만. `undo`는 되돌릴 수 없는 작업만 있으면 `VERIFIED: 0건.`, `undo --list`는 그런 작업이 20개 넘게 앞에 있어도 되돌릴 수 있는 지점을 보여 줌 |
| office-cli 1.4.29 | `transfer --chart`는 받는 문서를 찾지 못하면(오타 `[추천]`, `[후보]`, 확장자) 다음 방식(redraw, relay)으로 넘어가지 않고 원래 오류 그대로 멈춤(`[추천]`이 message 끝, 클립보드를 덮어쓰지 않음). outlook 폴더 `[후보]`는 끝 이름끼리 비교(같은 메일 계정 때문에 관계없는 폴더가 후보가 되지 않음), 오류 로그에 메일 계정 이름을 남기지 않음. 차트 다시 그리기는 `$`를 수식으로 읽지 않음. `reader search/index --file <파일 하나>`의 내부 오류는 `ERROR:`. `run_win_tests.py --agent -h`는 pywin32 없이도 JSON exit 0 |
| office-cli 1.4.30 | outlook `--folder` 오류의 `(경로: ...)`도 오류 로그에서는 계정 이름 자리를 `<메일 계정>`으로(`not_found`의 `log_extra`). `transfer --method relay --to ppt`는 받는 PPT의 `--slide`를 확인한 뒤에 클립보드·임시 워드를 씀(없는 슬라이드로 클립보드를 덮어쓰지 않음). relay `--to-target ""`는 활성 문서(`KeyError` 아님) |
| office-cli 1.4.31 | relay `--to ppt`는 열린·활성 PPT의 `--slide`를 찾기만으로 확인하고 쓰기 세션(되돌리기 지점)은 임시 워드 단계 뒤에 엶(워드 단계가 실패해도 바꾼 것 없는 지점이 남지 않음). 닫힌 xlsx의 계산된 값이 없는 수식 칸(openpyxl이 쓰고 엑셀에서 연 적 없는 파일)은 excel read/export, transfer 범위, reader read/search에서 UNVERIFIED와 칸 목록(`uncomputed_formulas`). docx 글상자 글을 따로 단위로 읽음(두 벌 저장은 한 번). 닫힌 docx 표 번호는 콘텐츠 컨트롤 안의 표도 셈(read·COM과 같음). word 굵게 위치는 UTF-16 단위. excel read --search에 `count`와 `VERIFIED: 0건.`. 마크다운 `---`는 word에서 구분선(아래 테두리 빈 단락), 맨 앞 YAML 머리말은 넣지 않고 `warnings` |
| office-cli 1.4.32 | 닫힌 xlsx를 범위 없이 읽을 때(`excel read`·`--search`·`--save`, `export`, transfer, `B:B`, 큰 범위 줄이기) 계산된 값이 없는 수식 칸의 행·열(합계 행 등)을 자르지 않고 남기며 UNVERIFIED. workbook.xml `fullCalcOnLoad`(xlsxwriter·openpyxl)면 모든 수식 칸을 계산되지 않은 칸으로 봄(`<v>0</v>`을 값으로 믿지 않음, 엑셀은 저장할 때 지움). 수식 칸 확인이 실패하면(MemoryError 포함) `uncomputed_check: "unknown"`과 UNVERIFIED "수식 계산 여부를 확인하지 못함". reader 파서 버전 10 |
| office-cli 1.4.33 | En-Genius 스킬로 실행하면(`ENGENIUS_OUT`이 있으면) `config.json`, data, cache를 `<사용자>\Documents\En-Genius\CLI-data\office-cli`에 둠 (트레이가 스킬 폴더를 실행마다 지우므로). `OFFICE_CLI_CONFIG`로 설정 파일 위치 지정 가능. OpenCode·`C:\CLI` 실행은 그대로 |
| office-cli 1.4.34 | 새 명령 `outlook.py attachments`: 첨부 목록(`index`, `name`, `size`, `hidden_inline`), `--save [전체 경로 폴더]`로 내려받기(기본 `%USERPROFILE%\Downloads\메일 첨부\<시각>_<제목>`, 자동 정리 안 함), `--name 일부`/`--index 1,3`으로 고르기(맞는 것이 없으면 `[후보]`), 같은 이름은 ` (2)`와 `recovered`. `read/current --save-attach` 없앰(대신 `attachment_names`와 `next_commands` "첨부 내려받기"). `new/reply/forward --attach`는 없는 전체 경로 파일이면 메일을 만들기 전에 같은 폴더의 비슷한 파일 `[후보]`, 붙인 뒤 `Attachments.Count`로 확인(못 하면 UNVERIFIED, `--send`면 보내지 않음) |
| office-cli 1.4.35 | En-Genius 스킬 zip은 지시문과 `run.py`만 담고 `C:\CLI\office-cli`를 실행: OpenCode와 같은 data, cache, config.json을 씀 (1.4.33의 고정 폴더는 도구가 트레이 임시 폴더에서 돌 때만) |
| office-cli 1.4.36 | `excel.py read --formulas`: 읽은 범위의 수식 칸 목록(`formulas`: `cell`, `formula`, `value`, 배열 수식은 `array`·`array_range`), `formula_count`, 200개 상한(`formulas_truncated`, 전체는 `full_result_file` CSV). 열린 문서는 Formula2(없으면 Formula)를 범위 한 번에, 닫힌 xlsx는 openpyxl. SKILL.md가 바뀌어 En-Genius 스킬 zip을 다시 가져와야 함 |
| office-cli 1.4.40~1.4.48 | 관리자 권한 실행이면 "열린 문서 없음" 대신 권한 안내(`elevated_process`), `range delete` 전 `#REF!`가 될 수식 칸 알림, reader 캐시의 확인 실패 다시 확인, 배열 수식 범위 전체를 수식 칸으로, 멀리 떨어진 계산 안 된 수식 칸 처리, `read --where`의 또는(`\|`)과 따옴표 값, 정렬 기준 4개 이상, 라이브러리 없음은 `missing_library`, next_commands 다듬기 |
| office-cli 1.4.49~1.4.51 | 규격서 대조 작업: `reader.py read --section`(캐시 문서 장절 읽기), `excel.py comment`(칸 메모 목록·달기·지우기, CSV로 여러 개), `reader.py search` 재현율(조사·단위 표기·전각 허용, 0건이면 부분/오타 일치로 넓힘 `match_mode`, 점수 순위, `data\synonyms.txt` 동의어) |
| office-cli 1.4.52~1.4.61 | `transfer --method`·`search -c` 삭제, 여러 칸 한 번에 `excel.py write --cells`, 시트 보호 `sheet_protected`와 거부된 수식 `[추천]`(되돌림), reader 등록·검색 `CONTINUE:`(`OFFICE_CLI_TIME_LIMIT`), 규격 대조 `excel.py check`, 열기 감시 30초 `drm_or_dialog`, `reader.py diff`(판 비교), 메일 짧은 `id`(긴 값은 `entry_id`)·`thread`·`calendar --days`, `word.py fill --file` |
| office-cli 1.4.62 | 통합 스킬 `office-cli`(16KB)를 작업별 스킬 4개로 나눔: `office-excel`, `office-doc`, `office-mail`, `office-spec` (`CLI/office-cli/skills/<이름>/SKILL.md`, 각 6KB 이하, 명령은 전체 명령). `apply_skills.py`가 4개를 설치하고 옛 `skills\office-cli` 폴더는 보관 후 지움. 도구 동작은 그대로 (`-h` 결과의 `skill_file`만 해당 작업 스킬) |
| office-cli 1.4.39 | 실제 사용 오류 로그 반영: `office.py write`, `office.py transfer --text`, `reader.py read --section`에 맞는 명령(워드/PPT/엑셀 write, reader search --section) 안내, `excel.py write --cell`에 `[추천] --range` |
| office-cli 1.4.38 | `opencode run` 기록에서: 엑셀 `write`의 `--text`에 `[추천] --value`(전에는 `--target`), 표 파일 못 찾음 후보에 표 파일만, SKILL.md에 칸 하나 쓰기(`--value`) 줄 |
| office-cli 1.4.37 | SKILL.md에 빠졌던 기존 기능(메일 list/folders·search --folder/--where·reply --all·new --cc·--body-file, PPT slide add/duplicate/delete/move, 엑셀 sheet/range/export/style/replace --all-sheets·차트 종류, 워드 --toc/--callout·replace --highlight)을 넣고 SKILL.md의 명령·옵션이 argparse에 있는지 테스트로 대조. 새 기능: `outlook.py list|search --days N|--since YYYY-MM-DD`(period), `excel.py range --action insert|delete`(값이 있으면 needs_user_decision, `--dry-run`, `--overwrite`), `range --action sort --by "부서,금액" [--desc]`, `read --where "금액 >= 100; 부서 = 영업"`(`row_numbers`) |
| ml-kit 2.9.11 | 날짜 열을 학습·예측에서 같은 방식으로 해석(YYYYMMDD 포함). `-i` 값은 글자로 받아 범주 코드 `007` 유지. 머리글·단위 행 옵션이 예측 `next_commands`에도 들어감. `-i` 6건 이상은 결과 CSV 저장 |
| ml-kit 2.9.12 | 성능예측표: 작은 값 출력의 자리수, "실측 없는 영역" 기준을 표 해상도에 맞춤, **입력은 타겟마다 최대 8개**, 타겟별 `--inputs "T1=a,b;T2=c"`. `ml_export`·`ml_map`에 `--overwrite`. `ml_map`은 열마다 그 열에 값이 있는 행으로 계산. 예측값은 작은 값이면 유효숫자 4자리 |
| ml-kit 2.9.13 | 타겟 값이 모두 같으면 `STOP_RETRY`(학습 안 함). 완전히 같은 행은 교차검증에서 같은 fold로 묶고 `warnings`에 중복 개수. **`[추천]` 뒤가 `python`으로 시작하는 전체 명령이면 그대로 1회 실행**(레이아웃 옵션을 뺀 예측, 성능예측표 `--inputs` 전체. 전역 지침 `CLI-dev/opencode/AGENTS.md`도 바뀜). `ml_map` 여러 구할 열은 `basis_by_column`. 작은 입력값(1e-6 단위)의 학습 범위·`--find` 결과 |
| ml-kit 2.9.14 | 숫자 입력에 쉼표(`-i "압력=500,000"`)가 있으면 `STOP_RETRY` + 쉼표를 뺀 전체 명령 `[추천]`. 예측 JSON의 `drift_warnings` 키 삭제(`extrapolation_warnings`는 20개까지, 전체 수 `extrapolation_warning_count`). 한 버전이 모든 타겟의 최고면 `ml_report -f`가 다시 학습하지 않고 그 버전을 씀(`made_best: false`). 머리글 행 안내는 열 이름 `[추천]`/`[후보]` 또는 `--header-row`. 입력+타겟이 다른 행이 10개 미만이면 `STOP_RETRY` |
| ml-kit 2.9.15 | 클래스가 많고 행이 적은 분류(10~15행)도 내부 오류 없이 학습. 숫자가 대부분인 단위 행은 `STOP_RETRY`, `[추천]` 머리글 행에 단위 행도 함께(전체 명령). 단위 행이 있는 새 파일을 옵션 없이 예측하면 `--unit-row`를 넣은 `[추천]` 명령. `--history`의 `best_version`(새 키 `best_version_dir`)과 보고서 `-d`는 모든 타겟의 최고가 있는 버전. 엑셀이 없는 PC의 보고서는 성능예측표를 확인하지 못해 `UNVERIFIED:` |
| ml-kit 2.9.16 | 범주·글자 타겟이 든 첫 데이터 행을 단위 행으로 보지 않음(`[추천]`에 틀린 `--unit-row`가 없고, 그런 `--unit-row`는 `STOP_RETRY`). 단위 없이 학습한 모델로 단위 행이 있는 새 파일을 예측하면 `--data-row`를 넣은 `[추천]` 명령. `ml_map --points`도 머리글 행 `[추천]` |
| ml-kit 2.9.17 | 단위 행 판단을 그 아래 값 최대 10개의 과반으로 함: 첫 값이 `측정불가`·`NG`·`12%`이거나 날짜 열(`yyyy-mm-dd`)이 있는 단위 행도 `--unit-row`로 받고, 옵션 없이 예측해도 단위 행을 건너뛰는 `[추천]`. 빈 행이나 머리글 위의 `--unit-row`는 `STOP_RETRY`(이전에는 진짜 단위 행을 데이터로 예측). 제목 행을 머리글로, 진짜 머리글을 단위 행으로 보던 `[추천]`(두 번 다시 실행) 수정 |
| ml-kit 2.9.18 | 숫자가 든 칸이 있는 행은 단위 행이 아님(첫 행이 `측정불가, 측정불가, A, 61.66`이면 데이터). 첫 데이터 행이 비었거나 `측정불가`여도 그 위의 단위 행을 건너뛰는 `[추천]`. 제목 줄이 좁은 CSV를 읽고 머리글 행 `[추천]`. 단위가 든 모델로 제목+머리글 파일을 예측하면 `[추천] --header-row`. 페이지마다 반복된 머리글(단위) 행은 빼고 `warnings`, 합계·평균 행은 빼지 않고 `warnings`('제외' 열 안내) |
| ml-kit 2.9.19 | 단위 행을 칸마다 아래 값과 비교해 판단: 타겟 열이 없는 예측 파일의 첫 행 `측정불가, 측정불가, A`는 데이터(모든 행 예측), 배율 숫자(`°C, 1000, , %`)와 이름 없는 머리글 아래 `비고`가 든 단위 행은 단위 행. `제외` 열에 표시한 합계 행과 집계 범주(`평균`/`최대`)는 합계 행 경고 없음. `;` CSV의 소수점 쉼표(`175,6`)는 소수점으로 읽고 `warnings`(천 단위와 섞이면 `STOP_RETRY`) |
| ml-kit 2.9.20 | `;` CSV의 `27,8448`(소수 4자리)을 278448로 읽던 문제 수정(정규식 이름이 겹쳐 있었음). `0,149`는 항상 소수, `1.346,8`(점 천 단위 + 소수점 쉼표)은 1346.8로 읽고 `warnings`(다른 형식과 섞이면 `STOP_RETRY`). 합계·소계 행은 열에서 차지하는 비율과 관계없이 경고(교대별 소계 25%, 빈 열의 합계 1개), 집계 범주로 보는 것은 `평균`류만. 단위 `ng`(나노그램)는 빈 값 표시가 아님(`NG`만). 큰 파일 읽기가 빨라짐(10만 행 x 32열 CSV 약 4.8초 -> 1초) |
| ml-kit 2.9.21 | 평균·합계 행을 값으로 판단: 바로 위 데이터 행들의 평균(합계)과 같은 `평균`(`합계`) 행은 비율과 관계없이 경고(교대별 `A조, B조, C조, 평균`). 통계 이름 열(`집계` = 합계/평균/최대/최소, `Sum/Avg`, 모든 행이 `평균`)은 값이 맞지 않으면 경고 없음. 소수점 쉼표 `;` CSV의 `1.300`(점으로 묶인 정수)을 1.3으로 읽던 문제: 같은 열과 파일이 일관되면 1300으로 읽고 `warnings`, 아니면 `STOP_RETRY` |
| ml-kit 2.9.22 | 합계 글자 행은 값과 관계없이 다시 항상 경고(하단 `합계`/`평균`/`최대`/`최소` 블록, `소계` 아래 총 `합계`, `평균`이 범주인 열의 `소계`), 통계 이름 열(합계·집계 글자가 값 칸의 절반 이상)만 경고 없음. 평균 행 값 비교가 천 단위 쉼표(`2,452`)를 숫자로 읽음. 통계 이름 열 10만 행 읽기 111초 -> 0.3초. 소수점 쉼표 CSV의 점 묶음 숫자(`1.300`, `2.125`)는 항상 `STOP_RETRY`(2.9.21의 1300 읽기 없앰), 모델에 소수점 쉼표 판단을 남겨 예측 파일·`ml_map --points` 조회 파일도 같게 읽음 |
| ml-kit 2.9.23 | 세로로 병합된 일자·라인 열 아래의 `합계`/`소계` 행도 다시 경고(통계 이름 열은 최대·최소·평균 같은 말이 함께 있을 때만). 엑셀 글자 칸의 `175,6`처럼 천 단위가 아닌 쉼표 숫자는 1756으로 바꾸지 않고 읽을 수 없는 값으로 경고 |
| ml-kit 2.9.24 | 일자 병합 일일 보고서의 일별 `합계`는 맨 아래 `최대`/`최소` 줄이 있어도 경고. `0,149`처럼 0으로 시작하는 쉼표 숫자는 천 단위가 아님(읽을 수 없는 값). 학습에서 숫자 열의 글자 칸을 중앙값으로 채우면 `warnings`에 알림 |
| ml-kit 2.9.25 | 읽지 못한 숫자 칸 경고는 학습하는 행(타겟 있음, `제외` 아님, 머리글 아래 단위 행 아님)과 모델이 숫자로 읽는 열(`--eval` 원본 열 포함)만: SKILL 예제의 단위 행 거짓 경고 5개 없앰, 열마다 1번, 5열을 넘으면 개수로. 합계 글자 넓힘(`일계`, `월계`, `누계`, `총 합계`, `전체합계`, `합계:`, `Total:`, `소계(A)`, `주간계`, `평균값`, `AVG.`; `설계`·`단계`·`온도계`는 아님). 선택한 피처 칸의 절반 이상이 소수점 쉼표 글자(`0,149`)면 `STOP_RETRY`, 그보다 적으면 경고 |
| ml-kit 2.9.26 | 2.9.25 회귀 수정: 범주 열의 `A계`/`B계`, `생산계`/`품질계`, `레벨계`/`pH계`를 합계 행으로 모든 행 경고하던 문제. 넓힌 합계 글자(`...합계`, 짧은 `...계`, 낱말 `Total`/`Sum`, `평균값`)는 열 안에서 합계 행처럼 보일 때만(채운 칸 절반 이하, 행 수 절반 미만, 글자 2개 이하) 경고, 정확한 낱말(`합계`, `소계`, `일계` 등)은 그대로. `SUM-01`, `Total_Loss`, `평균온도`는 합계·평균 글자 아님 |
| ml-kit 2.9.27 | 2.9.26 회귀 수정: 일자 병합 일일 보고서의 `일합계` + `월합계`/`총 합계`, 빈 열의 `총 합계` 1개를 경고하지 않던 문제. 넓힌 합계 글자는 값으로 판단: 행의 숫자 칸 절반 이상이 바로 위 데이터 행의 합, 위 합계 행의 합(`월합계` = `일합계`들), 위 모든 데이터 행의 합이면 경고, 합계 행이 적은 열에서는 숫자가 데이터보다 훨씬 크면 경고. 범주 값(`생산계`/`품질팀`, 센서 사이 `pH계`, `A계`/`B계`)은 경고 없음. `-i 'k=v'`의 같은 이름 두 번, `온도`와 `온도(°C)`를 함께 주면 `STOP_RETRY` |
| ml-kit 2.9.28 | 2.9.27 회귀 수정: 공정 열(`온도`/`압력`/`속도`)은 평균, `생산량`만 합인 `총 합계`/`월합계`/`일합계` 행을 경고하지 않던 문제 (블록 평균과 같은 칸은 판단에서 빼고, 합 1칸 이상 + 합·평균 칸이 절반 이상이면 경고). `-i`의 단위가 다른 이름(`압력(MPa)`, 모델은 `압력(bar)`)은 이름 바꾸기 `[추천]` 없이 모델 단위로 값을 환산하라는 `STOP_RETRY`. `온도`+`온도(°C)` 동시 입력은 모델 입력(ml_map은 데이터 열)과 맞출 때 판단: 모델에 둘 다 있으면 사용, `온도(°C)`만 있으면 `STOP_RETRY`. JSON `-i`의 같은 키 두 번, 같은 피처로 가는 별칭 쌍도 `STOP_RETRY` |
| ml-kit 2.9.29 | 단위가 다른 이름은 기본 이름이 달라도(`압력(MPa)`, 모델은 `반응압력(bar)`) `[추천]` 없이 값 환산 `STOP_RETRY`, `[후보]`의 단위 다른 후보에 "(단위 다름: 값 환산 필요)" (ml_map `-i`·`--find`·`--points`, `--find -i` 포함). `--find`의 목표 단위가 타겟과 다르면(`생산량(kg)`, 타겟 `생산량(t)`) `STOP_RETRY`. `--find`의 `-i`·목표에 여러 행 JSON 배열은 `STOP_RETRY`. 2.9.28 회귀 수정: 일정한 설정값 열이 평균 칸으로 세어져 범주 행(`A계`)을 합계 행으로 경고하던 문제 (블록 안에서 값이 바뀌는 열만 평균 칸) |
| ml-kit 2.9.30 | 타겟 이름(`--find` 목표, `-y`, `ml_report.py --target`)은 열 이름처럼 맞춤: `수율`이 `수율(1차)`·`수율(최종)`에 모두 맞으면 `[후보]` (이전에는 `--find`가 하나를 말없이 고름). 예측 `-i`의 모델이 쓰지 않는 이름(`압력(MPa)`, `습도`)은 `STOP_RETRY` (이전에는 버림). `ml_map.py --points`의 데이터 열과 비슷한 이름 열(`반응 압력(bar)`, `압력(MPa)`)은 `STOP_RETRY`. 합계 행: 블록 안에서 일정한 열의 값과 같은 칸은 판단에서 빼고 합 칸은 0이 아닌 수 2개 이상 (설정값을 적은 교대 보고서 `일합계` 경고) |
| ml-kit 2.9.31 | `[추천]`은 오타일 때만(`sure_typo`: 같은 이름, 한 이름이 다른 이름 안, 자모 하나 차이. 숫자가 다르거나 짧은 영문 이름의 길이가 다르면 아님), 아니면 `[후보]` (`습도`/`속도`, `T1`/`T2`, `CO2`/`O2`). 입력 열 오타 하나가 두 입력과 비슷하면 `[후보]` 하나, `--eval` 이름 `[추천]`/`[후보]` 문구. `ml_map.py --points`의 구할 열 이름 열(`생산량_실측`)은 실측값 열로 통과 (2.9.30 회귀), `-y`가 없으면 `[후보]`. 같은 단위 띄어쓰기(`수율( 1차)`)는 모든 이름에서 정확한 이름. 성능예측표는 단위를 뺀 이름이 같은 타겟에 전체 이름. 합계 행 규칙은 그대로 (거짓 경고 위험 문서화) |
| ml-kit 2.9.32 | 2.9.31 회귀 수정: 다른 옵션이 쓰는 열을 빼고 찾아 단위 없는 이름(`수율`, 열은 `수율(1차)`·`수율(최종)`, `-y 수율(최종)`)이 말없이 `수율(1차)`가 되던 문제(`--points` 기준 열, `-i`의 `[추천]`). `[추천]`은 확실한 오타만: 흔한 글자 하나를 다른 글자로(`속두`/`속도`, `강도`/`경도`), 영문 글자 하나, 짧은 영문 이름, 열 이름 앞부분(`온도`/`온도차`), 더 긴 이름(`설정온도`/`온도`), 다른 괄호 단위·구분은 `[후보]`. 다른 옵션이 쓰는 열은 후보가 아님(`-x`/`-y`, `--find`/`--vary`/`-i`, `--eval`/타겟, `--keep`). `--positive-class` 없는 라벨은 `[후보] 불량, 양품`. `--points`의 `-y` 없는 `[후보]`에도 `-y` 안내, `온두`/`온도(°C)` 비교 |
| ml-kit 2.9.33 | 이름 맞추기의 `[추천]`을 없앰: 철자가 비슷한 열·피처·타겟 이름은 하나여도 `[후보]`(사용자에게 물음; 2.9.32의 `속도` -> `[추천] '가속도'`, `수율` -> `[추천] '함수율(%)'`처럼 다른 양이 될 수 있음). `sure_typo`와 자모·KS X 1001·짧은 영문 규칙을 지움. 같은 열로 확실한 이름(띄어쓰기, 괄호 모양, `°C`/`℃`/`C`)은 `[INFO]`와 함께 그대로 맞춤. `ml_map.py --points`의 `-y` 이름이 든 실측값·메모 열(`수율 실측`, `목표수율`) 통과 (2.9.32 회귀), `-y` 없음 안내는 열 이름 오류에만. 숫자 이름 열(`2`)이 `--eval` 수식의 숫자를 가로채면 `STOP_RETRY`. `64.4%` 같은 숫자는 `%`를 떼고 읽고 경고. `ml_export.py --inputs` 상수 입력과 `--best --task` 없는 이력은 `[후보]`, 다른 옵션이 쓰는 열은 `[후보]`에 넣지 않음 |
| ml-kit 2.9.34 | 2.9.33 회귀 수정: `ml_map.py --points`에서 `-y` 이름이 든 기준 열(`냉각수 온도(°C)`와 `-y 온도`, `함수율%`와 `-y 수율`)이 실측값·메모 열로 말없이 빠지던 문제(비슷한 열 검사를 먼저, `[후보]`), `-i`와 `--points`의 `%` 값(`함수율=10%`, `12%`)을 데이터처럼 읽음(경고). 한 숫자 열에 `%` 칸과 `%` 없는 숫자 칸이 섞이면(`11.9%`, `0.119`) 학습·예측·표 보간 모두 `STOP_RETRY`. 숫자로 읽히는 단위 없는 이름(`2(mm)`의 `2`)도 `--eval` 수식 이름으로 쓰지 않고 수식에 그 숫자가 있으면 `STOP_RETRY` |
| ml-kit 2.9.35 | 2.9.34 회귀 수정: `ml_map.py --points`의 실측값·메모 열(`수율_실측`, `목표수율`, `-y 수율`)이 쓰지 않은 `함수율` 열 때문에 `[후보] '함수율'`로 멈추던 문제(두 이름에 든 `-y` 이름을 빼고 비교, `함수율%`는 그대로 멈춤). `--eval`의 단위까지 쓴 숫자 이름 열(`0.5(mm)*2(mm)`, `온도+2(mm)`)이 `STOP_RETRY`되던 문제와 더 긴 숫자 안의 숫자(`온도*2.5`와 `2(mm)` 열)를 그 열로 보던 문제. 학습의 섞인 `%` 검사는 `제외` 표시 행을 빼고 |
| ml-kit 2.9.36 | 2.9.35 회귀 수정: `ml_map.py --points`에서 위치·단이 다른 키 열(`온도_1`과 데이터 `온도(1)`, `온도_상`과 `온도(상)`, `압력1`과 `압력(1단)`)을 말없이 빼고 `VERIFIED`로 내던 문제. `-y` 이름을 뺀 나머지를 기호 없이 비교하고 한쪽이 다른 쪽을 포함해도 비슷한 열로 보아 `[후보]`로 멈춤 |
| ml-kit 2.9.37 | En-Genius 스킬로 실행하면 데이터 폴더(오류 로그)를 `<사용자>\Documents\En-Genius\CLI-data\ml-kit\data`에 둠. `ML_KIT_DATA_DIR`가 있으면 그것이 먼저. 모델·보고서는 지금처럼 데이터 파일 옆 |
| ml-kit 2.9.38 | En-Genius 스킬 zip은 지시문과 `run.py`만 담고 `C:\CLI\ml-kit`를 실행: OpenCode와 같은 데이터 폴더 |
| ml-kit 2.9.41~2.9.51 | .xlsx 백분율 서식 칸을 화면 값(11.9)으로 읽음(재학습 필요), `ml_map` 묻지 않은 열의 % 섞임은 빼고 계산, CSV 소수점 쉼표를 열마다 판단, `-i`의 `온도`+`온도(C)` 안내, 라이브러리 없음 `missing_library`, `--history` 0건 ok와 `--overwrite` 원장 중복 제거, `--dtypes` 이름 맞추기, 없는 버전 `[추천]`/`[후보]`, 목록 상한 `truncated`, 로그 태그 `[PROGRESS] 45%`, 인자 이름 `--save`(예측·보고서), 보고서 `--targets/-y` |
| ml-kit 2.9.52~2.9.58 | `--meta`·`--encoding`·보고서 `-d/-v` 삭제, CSV는 UTF-8 BOM, 학습 결과 `excluded_reasons`/`excluded_preview`, 예측 파일 레이아웃 그 자리 복구(`recovered`), 누수 `[후보]` 전체 명령 둘, `-i` 빠진 입력은 `needs_user_decision`+중앙값 명령, `--exclude-where`, 여러 타겟 학습 `CONTINUE:`, `--sweep`과 `--maximize/--minimize --vary --where` |
| ml-kit 2.9.40 | 회사 PC Gemma 시험 R1·R2가 opencode 명령 시간 제한(120초)에 끊김 -> 성능예측표 후보 학습을 저장하며 100초 전에 `CONTINUE:`로 멈추고 같은 명령으로 이어서 (R10). SKILL·전역 지침에 `CONTINUE:` 행동, 채점기는 CONTINUE 뒤 재실행을 반복으로 세지 않음 |
| ml-kit 2.9.39 | 학습 결과에 타겟별 중요 입력 `top_features`(최고 모델의 순열 중요도 비중, 최대 5개, 원래 열 이름)와 preview의 "중요 입력". SKILL.md에 버전 비교(`--history`), 파일 예측 엑셀 저장(`--out`), 파생변수(`--eval`), 예측 범위·확률 열 추가. SKILL.md 옵션이 argparse에 있는지 검사하는 테스트 |
| 개발 도구 | 내보내기는 `cmd /c "opencode export <세션ID> > chat.json"` (PowerShell 5.1의 `\| Out-File`이나 `>`는 한글을 깨뜨림, 채점기가 깨진 대화를 찾으면 [ERROR]). "한 번에 보내기"는 시나리오별로 나눠 채점(참고). 최종 답변은 시나리오 명령 뒤의 글. 러너 `--agent` 모드 JSON 계약(도움말 포함 `verified`), office 러너 `unverified_count`. 3차: 한 번에 보내기에서 가운데 시나리오의 재실행이 다음 시나리오로 넘어가지 않음, X2 답변 정규식이 틀린 답("바꿀 수 있습니다")을 받지 않음, `C:\CLI-dev`에서 이름만 쓴 러너 실행·glob 파일 목록·`*.md` grep은 위반 아님. 4차: X2는 마크다운으로 시작하는 "아니요"도 정답, "네,"로 시작하거나 "바꿀 수 있습니다"면 오답(`answer_none`), `C:\CLI-dev` 아래 상대 경로 러너 실행 허용, glob 패턴에 `.py`가 있으면(`*.{py,md}` 포함) 위반, `*.md` grep도 path가 .py 파일이면 위반. 5·6차: X2 오답은 "바꿀 수 있습니다", "코드·설정 파일을 고치면", "네"만 있는 줄로 판정(정중한 "네,"나 "예:" 예시 줄로 시작하는 정답은 통과), 채점 요약 표의 정규식 `\|` 처리. 7차: X2 판정을 실제 답변 형태 66개 말뭉치(`CLI-dev/tests/x2_answers.json`)로 맞춤 ("변경할 수는 없"·"불가" 정답 인정, 지어낸 `--outlier`류 옵션·코드/설정 파일 수정 안내는 오답). 8차: X2 답변 판정은 참고(최종은 --judge나 사람이 확인), 말뭉치 확대, P2/M1/M3/P3/R1/P1 답변 규칙 보완. 9차: M3 "범위를 초과", P3 "운전 폭"·"operating window" 인정, P2는 "90% 범위"만으로는 통과하지 않음 |

## 3. 회사 PC에서 할 일
상세: `docs/opencode/실기테스트.md` (1. 준비 -> 2. 자동 점검 -> 3. Gemma(선택) -> 4. 눈 확인 -> 5. 올릴 파일).

1. 준비: `C:\CLI\ml-kit`, `C:\CLI\office-cli`, `C:\CLI-dev`를 지우고 새로 붙여넣음(`C:\CLI\python`, `opencode.exe`는 둠).
   시험 폴더는 모두 `C:\CLI-dev\work\` 아래라 함께 지워짐. `python apply_skills.py`(C:\CLI-dev)로 OpenCode 스킬(ml-kit, office 4개: office-excel·office-doc·office-mail·office-spec)과 전역 지침(`CLI-dev/opencode/AGENTS.md`)을 적용하고 OpenCode 다시 시작 (옛 `skills\office-cli` 폴더는 보관 후 지워짐)
2. 자동 점검 (명령 하나, 질문 없음, 40~70분):
   ```powershell
   cd C:\CLI-dev
   python run_all_tests.py --label 회사PC
   ```
   - 환경 진단 -> pytest -> ml-kit 베타 점검(F7 엑셀 재계산 포함) -> office-cli 실기 러너 -> 도구 오류 로그. 결과는 `result_회사PC_<시각>.md` 하나
   - 맨 위 요약 표와 실패·ENV-SKIP·UNVERIFIED 목록을 봄. 베타 점검 F2(보고서·성능예측표 같은 폴더)와 F7을 꼭 확인
3. Gemma 종합 테스트(선택): 2의 명령에 `--gemma` (일반 권한 창). 시나리오를 `opencode run`으로 보내고 채점해 같은 결과 파일에 넣음
4. 눈 확인 10줄: 성능예측표(범위 밖 표시, 파이썬 없는 PC에서 문서 보안 아래 열기), Word `.md` 쓰기의 구분선, 엑셀 `read --formulas`·`--where`(날짜 칸),
   `range sort`·`insert`, relay 전달 실패 때 클립보드·되돌리기, Outlook `--days`(날짜 조건 형식, 1.4.37 남은 할 일)와 `attachments --save`
5. `result_회사PC_<시각>.md`, 눈 확인 O/X를 GitHub 웹으로 `docs/test_logs/`에 올림

회사 PC에서만 확인할 수 있는 것: 문서 보안(DRM) 영향, Office 16 차트 전달 방식, Gemma의 요청 처리, 사내 API 채점(`--judge`).

## 4. 지난 실기 결과 요약
- 집 PC(2026-10-31, Office 2007, ml-kit 2.9.10 / office-cli 1.4.9): beta check 27/27(F7 포함), pytest 통과, ml-kit 실기 14/14, office-cli 실기 PASS 69·ENV-SKIP 2(PDF 추가 기능 없음), UNVERIFIED 0건, 눈 확인 V1~V7·V9~V10 통과
- 회사 PC(2026-10-29~30): beta check 26/27(F2), Gemma 12/14(X1·X2에서 도구 코드 읽기 -> 지침 보강, X2 시나리오 변경)
- 회사 PC(2026-11-04, `run_all_tests.py --only gemma`, opencode CLI 2.2.5, ml-kit 2.9.39): Gemma 12/14. R1(보고서)·R2(성능예측표)가 opencode 명령 시간 제한 120초에 끊겨 결과 JSON 없음(R1은 같은 명령 4번) -> 2.9.40 나눠서 진행(`CONTINUE:`). 나머지 12건 통과, X2 답변 판정 통과(참고)

## 5. 남은 일
- 회사 PC 실기 결과 분석과 수정
- 각 도구 `CLI-dev/<tool>/CHANGELOG.md` 하단 "남은 할 일"
- 프로젝트 마무리 때 `docs/test_logs/` 등 로그 이력 정리 (저장소를 다시 만드는 방식)
- 도구 개발을 마친 뒤 다시 판단하기로 한 검토 (2026-11-01, 지금은 하지 않음):
  - MCP 전환: CLI 유지. 다시 본다면 CLI를 하위 프로세스로 부르는 얇은 MCP 층부터
  - 전용 에이전트: 사내 에이전트 En-Genius에 스킬로 붙이는 패키지를 만들었으나 업데이트를 멈추고 `archive/engenius`에 보관 (2026-11-04).
    다시 쓸 때 원래 자리로 옮기고 zip을 새로 만듦 (archive/README.md). 도구 쪽 En-Genius 대응(임시 폴더 실행 시 고정 데이터 폴더)은 그대로 둠
  - ml-kit 웹 서버: 동료 수요(예측이면 성능예측표로 충분)와 IT 정책·문서 보안 확인 후
  - `opencode run` 자동화: 사용 중. opencode CLI v2.2.5, 일반 권한 창 (관리자 창에서는 열린 문서를 못 봄), 기본 모델은 opencode.jsonc의 `model`.
    `run_all_tests.py --gemma`가 시나리오 14개를 한 세션에 보내고 `grade_opencode.py --runs`로 채점 (2026-11-04 회사 PC 12/14, R1·R2는 2.9.40에서 수정)

## 6. 붙여넣을 첫 메시지
```text
AI_WORK 저장소 작업을 이어서 해 줘. 먼저 AGENTS.md와 docs/HANDOFF.md를 읽고 지금 상태를 파악해.
지금은 회사 PC(Windows 11, WinPython 3.13.13, Office 16, 문서 보안)에서 실기 테스트를 하려고 해.
1. HANDOFF.md 3장대로 준비를 같이 하고, 자동 점검 `python run_all_tests.py --label 회사PC`(C:\CLI-dev에서) 한 번을 실행해 줘. 명령은 python만 쓰고, 실행 전에 무엇을 할지 짧게 알려 줘.
2. 끝나면 결과 파일 result_회사PC_<시각>.md 맨 위의 요약 표와 실패·ENV-SKIP·UNVERIFIED 목록을 표로 정리해 줘. 원인 추정과 고칠 곳(파일:줄)도.
3. 고칠 것은 바로 고치지 말고 계획을 먼저 보여 줘.
답변은 한국어로, 이모지 없이.
```
