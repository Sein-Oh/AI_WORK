# CHANGELOG - bom-kit

형식: 버전별 `구조 / 기능 / 수정` + 하단 `남은 할 일`.

## [0.2.0] - 2026-11-14

### 기능
- `history` 명령과 기본 스킬 `bom-history` (사용자 결정 2026-11-14: BOM 변경 이력 추적). diff마다 옛 판 옆 결과 폴더에 기록
  (`diff_<새 판>_<시각>_기록.json`: 시각, 두 판의 경로·수정 시각·크기, 건수, 결과 파일)을 남기고, `history --file <아무 판>`이 비교로 이어진
  판들의 기록을 시간순으로 모아 `변경이력_<시각>.csv` 하나로. 기록은 결과 폴더에 있어 도구를 다시 반입해도 남음. setup 기본 스킬 16개.
  시험 `test_history_collects_the_diffs_of_connected_versions`.

### 수정 (Fable 검토 2026-11-14: 비표준 BOM에서 틀린 답을 VERIFIED로 냄, 재현함)
- 점 표기 Level(`.1`, `..2`)을 모두 버리고 "두 판이 같습니다"였던 것: 점 표기를 읽음.
- 한 판만 수량 열 이름이 다르면(`Qty`/`수량`) 비교하지 않고 "두 판이 같습니다"였던 것: 같은 뜻의 열(Level·품번·수량·품명·단위)을 짝지어 비교.
- 같은 상위 아래 둘이던 품번이 하나가 되면 `#1`을 떼어 짝이 어긋나 "추가 3건, 삭제 6건"이던 것: 짝 맞추는 키는 늘 `#순번`,
  구분자는 품번의 `/`·`#`과 겹치지 않는 것으로 (품번 `A/100`이 경로와 겹치던 것도). 화면과 변경내용 파일은 전처럼.
- `VERIFIED` 조건: Level·품번을 읽지 못한 행이 있거나 한쪽 판에만 있는 열을 비교하지 못했으면 preview 첫 줄 `주의:`와 `UNVERIFIED:`
  (warnings는 뒤에 있어 경량 모델이 빠뜨림). where-used도 같은 규칙.
- 시험 `test_irregular_boms_are_compared_or_said_plainly`.

### 수정 (집 PC 실기 2026-11-14, `docs/test_logs/집PC_bom_20261114/결과.md`)
- 결과 폴더와 작업 폴더를 나눔 (사용자 결정 2026-11-14, Fable 검토 B): 결과는 원본 옆 `<원본 이름>_bom\`에 실행마다 시각을 붙인 새 이름
  (`diff_<새 판>_<시각>_요약.txt`·`_변경내용.csv`·`_칠할칸_옛판/새판.csv`, `쓰인곳_<찾는 말>_<시각>.csv`)으로 두고 덮어쓰지 않으며(R4), pandas-kit 작업 폴더는
  캐시라 실행마다 지운다. 0.1.1은 `<파일 이름>_<시각>` 폴더가 실행마다 쌓였고(B3~B9 뒤 5개), 검토 중 넣었던 "원본·명령당 1개" 규칙은
  이미 알려 준 `output_file`과 칠할 칸 파일을 지워(다른 `--find`, 같은 옛 판의 다른 새 판) 다음 턴의 `confirm_commands`가 실패했다.
  결과 JSON의 `work` 대신 `result_folder`.
- 안의 파일이 열린 작업 폴더 (B9): `shutil.rmtree`가 표시 파일·스냅샷부터 지우고 열린 csv에서 멈춰, 표시 파일이 없어 다시는 못 지우는
  폴더가 남았다. 이제 폴더 이름부터 바꿔 보고(Windows는 열린 파일이 든 폴더 이름을 못 바꿈) 안 되면 손대지 않고,
  `works.json`의 "못 지운 폴더"로 기억했다가 다음 실행 때 다시 지운다 (`works.json`은 임시 파일 뒤 바꿔 넣기).
- 두 실행을 동시에 하면 둘 다 멈추던 것 (Fable 검토 A2): pandas-kit 데이터 폴더(지금 작업 포인터)를 실행마다 따로 쓴다.
  멈춘 실행의 message에 pandas-kit 명령 문장이 새어 나오던 것도 뺌.
- `where-used`가 30곳을 넘으면 preview 끝에 `(나머지 N줄은 <output_file>)` (diff와 같은 모양). 전에는 preview를 그대로 전하면 30곳만 보였다.
- `where-used`에 숨긴 행이 걸리면 `warnings`에 엑셀 행과 품번 (B4: 숨긴 폐기 품목 110099가 '나사' 결과와 합계에 들어감).
  pandas-kit 0.2.10 `select`의 `hidden_rows`를 씀.
- 틀린 명령 이름의 message에서 argparse 영어 문장(`argument cmd: invalid choice ...`)을 빼고 한국어만.
- 시험: `test_results_stay_in_one_result_folder_and_work_folders_are_removed`, `test_two_bom_runs_at_once_both_finish`,
  `test_a_work_folder_with_an_open_file_is_left_whole_and_removed_next_time`,
  `test_more_than_thirty_places_say_where_the_rest_is`, `test_a_word_in_the_name_finds_every_part_and_counts_them_apart`(숨긴 행 경고),
  `test_misspelt_command_name_gives_candidates`(Windows에서 bash 따옴표로 감싼 경로를 못 나누던 시험 고침).

## [0.1.1] - 2026-11-14

### 수정 (저장소 점검에서 직접 써 보며)
- Level 칸에 `inf`, `nan` 글자가 있으면 내부 오류(`OverflowError`, `ERROR:`)가 나던 것: Level이 아닌 행으로 보고 `warnings`로 알림.
  수량의 `inf`, `nan`도 숫자가 아닌 수량으로 봄.
- 원본 옆에 작업 폴더가 쌓이던 것 (실행할 때마다 `<파일 이름>_<시각>` 폴더 하나, `--find`만 바꿔 6번이면 6개, 멈춘 실행도 빈 폴더를 남김):
  멈춘 실행(`STOP_RETRY:`)이 만든 폴더는 지우고, 같은 명령(같은 파일·시트·범위·열·찾는 말)을 다시 실행하면 앞 실행의 폴더를 지운다.
  다른 `--find`, 다른 파일의 결과 폴더는 그대로 둔다 (`output_file`을 나중에 열 수 있게). pandas-kit 표시 파일이 있는 폴더만 지우고,
  파일이 열려 있어 못 지우면 그대로 둔다 (stderr `[WARN]`). 어느 실행이 어느 폴더였는지는 `<데이터 폴더>/works.json`(최근 200개).
- 틀린 명령 이름(`dif`)에 `[후보]`와 고친 전체 명령 (`_bom_runtime.py`, pandas-kit 0.2.9와 같음).

### 스킬
- `mbom-diff`, `where-used`의 `metadata.version`을 도구 버전과 같게 (전에는 스킬 스크립트 시절의 0.3.0, 0.2.0이 남아 실기 러너가
  bom-kit 버전을 0.3.0으로 적음). 시험 `test_skill_versions_match_the_tool`.
- 시험: `test_work_folders_do_not_pile_up_next_to_the_original`, `test_level_that_is_not_a_finite_number_is_a_skipped_row`,
  `test_misspelt_command_name_gives_candidates`.

## [0.1.0] - 2026-11-14

### 구조
- 새 도구 (사용자 결정 2026-11-14): mbom-diff·where-used 스킬 폴더의 스크립트 3개(`mbom_path.py`, `mbom_summary.py`, `where_used.py`)에
  따로 있던 같은 계산(Level로 상위·경로, 열 찾기, csv 읽기)을 `AIKIT/bom-kit/bom.py` 한 곳으로 (한쪽만 고치면 두 스킬이 같은 BOM을 다르게 읽음).
  공통 런타임 `_bom_runtime.py`는 pandas-kit `_pd_runtime.py`에서 복사 (표준 라이브러리만). 스킬은 그대로 둘 (필요한 순서만 읽음).
- 스킬 스크립트와 달리 공통 계약을 따름: 인자 오류·도움말 JSON, 같은 뜻 옵션 `[추천]`, 오류 로그, 잠긴 파일 새 이름, 인코딩 자동 판별,
  pandas-kit 명령의 셸 따옴표(`quote_arg`).
- 명령 하나로 (사용자 결정 2026-11-14, Gemma급 검토): 스킬이 pandas-kit 명령과 스크립트를 13단계(diff)·4단계(where-used)로 잇던 것을
  `bom.py diff`·`bom.py where-used` 하나로. bom-kit이 pandas-kit을 실행(import 아님, 데이터 폴더 따로)하고 최종 결과만 줌.
  중간 명령의 `next_commands`·지시를 경량 모델이 따라가 멈추는 위험, 긴 경로 옮겨 적기, 결과 JSON 13개가 쌓이는 문제를 없앰.
  스킬은 4KB -> 2.8KB.

### 기능
- `diff`, `where-used` (README). 결과는 스킬 스크립트와 같음 (v2 -> v2_mod 요약, where-used 4가지 찾기를 나란히 돌려 확인).
- pandas-kit이 시트·표 범위를 물으면 bom.py 전체 명령으로, 열을 못 찾으면 열 `[후보]`로. pandas-kit 경고는 사실만 (명령을 권하는 문장은 뺌).
- 시험: `test_mbom_diff.py`, `test_where_used.py`(pandas-kit 시험에서 옮김), `test_contract.py` (사용자의 pandas-kit 지금 작업을 바꾸지 않는지 포함).

## 남은 할 일
- 누적 수량·총소요량(제품 N대에 품목별 합계), BOM 한 판 점검(같은 상위 아래 같은 품번, 수량 0, 단위 섞임).
- Level 표기 `.1`, `..2`(점으로 깊이), 상위 품번 열이 이미 있는 BOM.
- 사용 가이드 HTML (`AIKIT/guide/`), 실기 러너·Gemma 시나리오 (지금은 `run_all_tests.py`가 버전과 오류 로그만 붙임).
