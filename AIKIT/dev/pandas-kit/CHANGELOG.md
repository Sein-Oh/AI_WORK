# CHANGELOG - pandas-kit

형식: 버전별 `구조 / 기능 / 수정` + 하단 `남은 할 일`.

## [0.2.1] - 2026-11-14

### 수정 (MBOM 두 판으로 기반 명령을 써 보며 찾은 '결과가 틀리는데 알리지 않음', 개선점_MBOM.md E)
- 글자 열에 `=`로 숫자를 주면(`Part Number=125`, `Op No=10`) 글자 그대로 비교해 0건이었는데 말이 없었음 -> 숫자로 같은 값이 다르게 적힌 칸 수와
  적힌 값(`000125`, `0010`)을 `warnings`로, 적힌 그대로 쓰는 예. 비교 방식은 그대로 (값을 바꾸지 않음).
- 날짜 열 조건 값이 연-월-일이 아니면(`2026-13`, `2026-12`, 없는 날 `2026-02-30`) 글자 비교로 그럴듯한 결과를 냈음 -> `STOP_RETRY` (연-월-일로).
- `join`(양쪽에 있는 행만): 짝이 없어 빠진 행 수를 알리지 않았음 (MBOM과 Routing: 110행 중 104행) -> 양쪽 빠진 행 수 `warnings`.
- `join`의 오른쪽 키 열 `[후보]`: 철자 순이라 맞는 열(`Item`)이 둘째였음 -> 왼쪽 키와 겹치는 값이 많은 열부터, 후보마다 겹치는 값 수.
- `only`: 키가 행마다 다르지 않으면(MBOM `Part Number` 34행 중복) 지운 행의 키가 다른 행에 남아 있을 때 결과에서 빠지는데 알리지 않았음 -> `warnings`.
- `diff`: 키가 같은 행이 있어 멈출 때 어느 키·행인지 몰랐음 -> 예 3개(`000125 (엑셀 행 16, 30, 35, 38, 58 등)`). BOM 용어(상위 품번) 안내는 뺌 (업무 판단은 스킬).
- 표 위의 다른 내용 경고가 한 칸짜리 제목 줄을 세지 않아 "3행(엑셀 행 2~4)"이었음 -> "4행(엑셀 행 1~4)".
- `export` xlsx: 날짜만 있는 열이 엑셀에서 `2026-01-01 00:00:00`으로 보였음 -> 날짜 형식 `yyyy-mm-dd` (시각이 있으면 시각까지).
- 시험: `test_mbom_title_rows_counted_and_number_written_differently_is_warned`, `test_mbom_date_condition_must_be_a_full_date`,
  `test_mbom_join_says_unmatched_rows_and_ranks_key_candidates_by_shared_values`, `test_mbom_only_and_diff_with_repeated_keys`,
  `test_mbom_export_keeps_dates_as_dates`.

## [0.2.0] - 2026-11-13

### 시험 자료
- `tools/make_mbom_samples.py`: 집 PC에서 쓴 MBOM 두 판을 클라우드에서도 같게 만듦 (xlsx는 반입 규칙상 저장소에 넣지 않음). 시험 `test_mbom_samples_are_made_and_read`.

### 구조
- 성격을 완결형 작업 도구에서 기반 도구로 바꿈 (사용자 결정 2026-11-13, office-cli와 같은 개념): 표를 고르고·보고·비교하고·잇고·저장하는
  단위 기능만 둔다. BOM 분석(Level로 상위 찾기, 누적 수량, 총소요량) 같은 업무 판단은 이 명령들을 조합한 스킬이 한다.
  MBOM으로 써 보며 넣었던 `tree`·`rollup`과 `diff`의 BOM 안내는 같은 이유로 넣지 않음 (MBOM 스킬로).

### 기능 (MBOM_DR200.xlsx로 직접 써 보며 찾은 기반 기능 부족)
- 표 찾기: 빈 행으로 나뉜 내용 중 가장 큰 것을 표로 (제목 블록 `공장·Rev·작성일`을 표로 잡던 것), 표 위의 다른 내용은 `warnings`.
- `group --by --sum`: 값별 행 수·합계·엑셀 행 목록, 단위 섞임 `warnings`. `sort --by --desc`: 정렬(같은 값은 표 순서, 빈 칸은 끝).
- `show --start`: 21행부터 보기, 더 있으면 다음 20행 `next_commands`.
- 작업 폴더 이름: `pandas_YYMMDDHHMM` -> `<파일 이름>_YYMMDDHHMM`, 같은 분이면 `_2` (사용자 결정, 어느 파일의 작업인지 이름으로 보임).
- `--on "왼쪽 열=오른쪽 열"`: 이름이 다른 키(`Part Number=Item`), `엑셀 행` 키로 스킬이 계산한 csv를 원래 표에 붙이기.
  머리글 `엑셀 행`처럼 도구의 행 번호 열 이름은 `엑셀 행번호`로 바꿔 읽음. 키의 정수는 종류와 상관없이 같게 봄.
- 시험: `test_title_block_above_is_not_taken_as_the_table`, `test_join_with_different_key_names_and_back_by_excel_row`,
  `test_group_counts_sums_and_warns_mixed_units`, `test_sort_keeps_excel_rows_and_puts_empty_cells_last`, show `--start` (기존 시험에 더함).

### 수정
- `diff`: 한 판에만 소수 수량이 있으면 열 종류가 달라(정수/소수) 2와 2.0을 바뀐 칸으로 셌음 (numpy 정수를 숫자로 보지 않던 것).
  시험 `test_diff_same_number_in_int_and_decimal_columns_is_not_a_change`.

## [0.1.2] - 2026-11-13

### 수정 (MBOM_DR200.xlsx로 직접 써 보며 찾은 '결과가 틀리는데 알리지 않음' 5가지)
- 숫자 사이 글자로 저장된 숫자 한 칸이 열 전체를 글자로 만들던 것 -> 숫자 열로 읽고 `warnings`(칸 수, 엑셀 행). 숫자 사이 숫자가 아닌 칸이 조금 섞인 글자 열도 `warnings`.
- 글자 열에 숫자 값으로 `>` `<` `>=` `<=`: 글자 순서 비교("18" > "3"이 거짓) -> 숫자로 읽히는 칸끼리 숫자 비교, 숫자가 아닌 칸 수 `warnings`.
- `!=`·`!~` 갈래만 `|`로 이으면 "어느 것도 아님" (`Item Type!=RAW|CONS`가 모든 행을 남기던 것).
- 값 없는 조건(`Qty>=`, `Desc~`)이 모든 행을 내던 것 -> 거절 (빈 칸은 `열=`·`열!=`만).
- `chart`: 단위 열(`UoM` 등)에 단위가 여러 가지면 `warnings` + 단위 하나로 고르는 전체 명령.
- 시험: `test_numbers_written_as_text_and_numeric_conditions_on_text_columns`, `test_not_equal_with_choices_means_none_of_them_and_empty_value_needs_equal`, `test_chart_warns_when_units_are_mixed`.

## [0.1.1] - 2026-11-13

### 수정
- 집 PC 실기(2026-11-13): `join`·`only`의 preview·계보에 받침과 맞지 않는 조사("df1와 df2를")가 나옴 -> "df1, df2을(를) <키> 기준으로 짝 맞춤", "(<키> 기준으로 df1과(와) 비교)". 시험 `test_join_only_stack_keep_both_row_numbers`.
- 시험: `test_diff_adds_deletes_changed_cells_with_rows_and_paint_files`가 명령을 `split()`으로 나눠 Windows 경로의 따옴표가 남아 실패 -> `shlex.split`.

## [0.1.0] - 2026-11-13

### 구조
- 새 도구 (완결형 작업 도구, 설계 `설계.md`, 결정 `docs/DECISIONS.md` 2026-11-13). `AIKIT/pandas-kit/pandas_kit.py`(진입점),
  `_pd_runtime.py`(ml-kit `_ml_runtime.py`에서 필요한 부분만 같은 이름·같은 동작으로 복사: `run_cli`, `JsonArgumentParser`, `quote_arg`,
  오류 로그, `retry_fs`, `save_with_lock_fallback`, `read_text_file`). 기본 스킬 `pandas-kit` (setup 기본 스킬 13개, `.gitignore`, 지침).

### 기능
- 첫 판 명령: `select`, `show`, `join`, `diff`, `only`, `stack`, `undo`, `new`, `export`, `chart` (README.md).
- 표 찾기(제목 한 줄 건너뜀, 빈 행에서 끝), 엑셀 행 번호·열 글자 유지, 값은 바꾸지 않음(앞자리 0 코드는 글자), 병합 칸·숨긴 행·저장 안 된 수식 `warnings`.
- 상태는 작업 폴더의 파일: 단계 기록(원본) + pickle 스냅샷(캐시, 깨지면 기록으로 다시 만듦), 원본 변경 감지, 옛 표 기준 경고.
- `diff`: 바뀐 곳마다 한 행 + 원본 파일별 칠할 칸 목록과 office-cli 칠하기 명령(`confirm_commands`).
- `chart`: 표를 차트 그림(.png)으로 (막대·선·산점·원·히스토그램·상자, 사용자 요청 2026-11-13). 같은 x는 합계, 막대 30개·원 12조각까지, 한글 글꼴 찾기, 다시 열어 확인.

## 남은 할 일
- 다음 판 명령: `replay --file <새 판>`(같은 단계를 다른 파일에), 명시적 `fill`·`trim`. (`where-used`·`rollup` 같은 BOM 계산은 도구가 아니라 MBOM 스킬로)
- 긴 작업 나눠서 진행(`CONTINUE:`, R10): 첫 판은 한 번에 끝냄 (큰 파일 시간 측정 뒤 넣음).
- 범위 밖 주소(`C300`처럼 읽은 범위 밖)를 받으면 범위를 넓혀 다시 읽는 `[추천]`: 첫 판은 `--columns`의 열 글자가 표에 없으면 열 목록만 알림.
- 사용 가이드 HTML (`AIKIT/guide/`), 실기 러너·Gemma 시나리오.
- 자유 코드 `eval`(README에만, 대형 모델용).
