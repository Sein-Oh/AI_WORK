# CHANGELOG - ml-kit

이 문서는 2.0.0부터 관리합니다. 형식: 버전별 `구조 / 기능 / 호환` + 하단 `남은 할 일`.

> 테스트 파일 정리 (2026-11-04, 도구 동작 변경 없음): 차수별 테스트 파일(`test_review_v29xx.py`, `test_layout_matrix_v29xx.py`,
> `test_features_v21.py` 등 46개)을 기능별 12개(`test_contract`, `test_safety`, `test_train`, `test_predict`, `test_names`, `test_map`,
> `test_report_export`, `test_loading_layout`, `test_loading_totals`, `test_loading_numbers`, `test_layout_matrix`, `test_skill_docs`)로
> 합쳤습니다. 공용 도구는 `tests/mlkit_helpers.py`, 여러 파일이 쓰는 fixture는 `tests/conftest.py`. 같은 입력·같은 기대값의 중복 검사
> 13개를 빼고, 같은 데이터로 학습하던 모델은 한 번만 학습합니다. 아래 버전 항목의 테스트 파일 이름은 당시 이름입니다.
>
> 회사 PC 실기(2026-11-05, Windows) 테스트 수정 (도구 동작 변경 없음): `[추천]`·`next_commands` 명령을 `shlex.split`(posix)으로 나눠
> Windows 경로의 백슬래시가 지워져 11개가 실패하던 것을 `mlkit_helpers.split_command`(Windows 명령줄 규칙)로. conftest의
> `pytest_make_parametrize_id`로 긴 글자 매개변수 id를 60자로 자름 (Windows 환경변수 32767자 한도).

## [2.9.68] - 2026-11-05

`next_commands`·`[추천]` 전체 명령의 인자 따옴표를 실제 실행 셸에 맞춤 (office-cli 1.4.76과 같은 코드, 복사).

### 버그 수정
- `quote_arg`(PART C)가 cmd 기준 `"..."` + `\"`라 PowerShell에서 깨지던 값(끝이 `\`인 폴더, 값 안의 `"`, `$`·백틱, cmd의 `%NAME%`):
  `pwsh`가 PATH에 있으면 PowerShell 7, 없으면 Windows PowerShell 5.1(OpenCode 셸 도구와 같은 순서, `ML_KIT_SHELL`로 강제)에 맞춰
  작은따옴표 `'...'`로 감싸고, 5.1이면 값 안의 `"`·끝 백슬래시·빈 값을 미리 바꿈. 단순 토큰(영숫자·한글·`_-.:\/,=+@!`)은 그대로.
- `ml_train.py`가 따로 `"..."`를 붙이던 `_q`와 `ml_report.py`의 예측 명령(`-m`, `-f <새데이터.xlsx>`)도 `quote_arg`로.
- 테스트: `mlkit_helpers.split_command`가 셸 시뮬레이터(`CLI-dev/tests/shell_sim.py`)를 씀. `test_contract`에서 `ml_predict -i` 채우기 명령과
  `ml_train` 타겟 누수 `[후보]` 명령(`--exclude-where`, `$`·`"`·`'`·`%`가 든 `--desc`)을 셸 4종마다 끝까지 확인.

## [2.9.67] - 2026-11-05

6차 독립 검토(10000행 x 1000열 회귀)의 [치명] 발견 수정. 항목마다 회귀 테스트를 추가했습니다.

### 버그 수정
- [치명] 큰 데이터의 나눠 진행이 끝나지 않음 (R10): 실행마다 학습 전 고정 비용(읽기 8초, 검증 26초, 상관행렬 38초, 타겟 준비 20초, 합 50~94초)을
  다시 하고, 그 실행의 첫 단계(fold 1개, 시계 없음)를 더하면 120초를 넘어 강제 종료 -> JSON 없음 -> 진행 없음 -> 같은 명령을 끝없이 반복했음.
  - 학습 전 준비를 한 단계로 저장 (`_prepare_run`, `_prep_key`, `_prep_load`): 첫 모델 전에 하는 일(읽기, 검증·타입 정제, 피처 판별, 상관행렬, `%` 검사,
    학습 제외 행, 타겟별 저장 키)을 한 번에 하고, 시간 제한이 있고 `PREP_SAVE_SECONDS`(5초) 이상 걸렸으면 `.resume_train/<키>.prep.joblib`에 통째로 저장
    (대체 위치 JIT 복구 포함). 키는 데이터 파일 바이트, 데이터를 정하는 인자(목록 파일은 내용), 행 인자, 시드, 도구·scikit-learn·pandas·numpy 버전.
    이어 실행은 준비를 다시 하지 않음. 준비를 저장한 실행은 그것이 한 단계라 다음 단계 전에 시계를 보고, 이미 제한 - 여유를 넘었으면 첫 학습 단계 전에 멈춤
    (`progress`: `학습: 학습 준비(데이터 읽기, 검증, 상관행렬) 저장됨 (첫 학습 단계 전에 멈춤, 약 N단계 남음)`). 상관행렬도 이 준비에 들어 한 번만 계산.
    같은 객체를 저장했다 읽으므로 결과(리더보드, OOF, 모델, `top_features`, 패키지 파일)가 한 번에 학습한 것과 같음.
  - 숫자 열 읽기(`coerce_numeric`)를 칸마다 한 번에 처리(정규식은 `%`가 있는 칸만, 결측 표시는 집합 대조)하고 `%` 칸 찾기(`percent_cells`, `percent_mix`)도
    `%`가 없는 칸은 건너뜀: 같은 결과로 준비 시간 94초 -> 53초 (10000행 x 1000열).
  - 시작 표시 (`step_begin`/`step_end`, PART C): 실행의 첫 단계(시계 없이 하는 준비, fold, 1위 전체 적합, 중요도 첫 열, 성능예측표 첫 후보)는 시작 전에
    `<단계>.start`(시도 횟수, 그때 제한)를 남기고 저장하면 지움(실행이 오류로 끝나도 지움: 밖에서 끊긴 실행만 남김). 표시만 남은 단계를 다시 만나면 2번째까지는
    다시 시도하고(일시적 부하로 한 번 끊길 수 있음), 같은 제한 이하에서 2번 끊긴 단계는 시작하지 않고 `STOP_RETRY`:
    `이 데이터는 한 단계(...)가 명령 시간 제한(100초) 안에 끝나지 않습니다 ... [추천] <같은 명령> --time-limit 570` (명령 실행 도구 timeout 600000으로).
    늘린 제한에서는 다시 시도하고, 570초로도 끊겼으면 `[추천]` 없이 멈춤. 정리 패턴(`_RESUME_FILE_RE`)에 `<키>.prep.joblib`, `<단계>.start` 추가.
  - 안전망 감시 스레드 (`run_cli`, `start_watchdog`): 명령 시작부터 제한 + 15초(`WATCHDOG_GRACE`, 기본 115초)에 아직 끝나지 않았으면 1줄 JSON을 직접 내고
    `os._exit`: 이 실행이 저장한 단계가 있으면 `CONTINUE`(`<단계> 도중 시간 제한(100초)에 닿아 멈춤. 그 전 단계까지 저장됨`), 표시한 첫 단계의 첫 끊김이면
    `CONTINUE`(`그 단계 다시 시도 1/2`), 아니면 위 `STOP_RETRY`(오류 로그 기록). stdout은 락과 표시(`emit_json_response`가 이미 썼으면 쓰지 않음)로 한 번만
    쓰고, 버전 폴더 커밋 중이면 4초 기다림(`WATCHDOG_HOLD`). 제한이 없거나 ml-kit 명령이 아닌 실행(같은 프로세스 테스트)에는 감시 없음.
  - 단일 타겟도 끝난 타겟을 저장 (제한 - 여유를 넘었을 때): 마지막 패키징 도중 감시가 끊어도 다음 실행이 학습을 다시 하지 않음.

### 기능
- 모든 ml-kit 명령이 `--time-limit 초`를 받음 (`JsonArgumentParser`, `ML_KIT_TIME_LIMIT`보다 우선, 0 = 없음). 위 `STOP_RETRY`의 `[추천]` 명령용.
- `progress`에 남은 단계 수 추정 `약 N단계 남음`(후보 x fold + 전체 적합 + 중요도, 뒤 타겟은 27단계씩).
- CONTINUE 반복 상한: SKILL.md·전역 지침(`CLI-dev/opencode/AGENTS.md`, office-cli와 공통 문구)·CLI-dev/README.md R10의 "최대 20번"을
  "`progress`가 바뀌는 동안 반복, 바로 앞과 같은 `progress`가 나오면(2번 연속) 멈추고 알림, 안전 상한 60번"으로. office-cli의 `progress`도 실행마다
  처리 개수가 바뀌어 같은 규칙이 맞음(office-cli 파일은 바꾸지 않음).

### 테스트
- `test_train.py`: `test_preparation_is_saved_and_reused_by_the_continued_run`(준비 저장 -> 첫 단계 전 CONTINUE -> 이어 실행은 준비를 다시 하지 않고
  한 번에 학습한 것과 같은 버전), `test_preparation_key_follows_data_and_arguments`, `test_step_mark_stops_a_step_longer_than_the_limit`,
  `test_time_limit_argument_overrides_the_environment`, `test_watchdog_writes_one_json_line_at_the_limit`(작은 제한 + 느린 가짜 단계: 저장 있음 CONTINUE,
  표시 단계 1번째 CONTINUE, 2번째 STOP_RETRY와 오류 로그, 늘린 제한 다시 시도, 저장 없음 STOP_RETRY, 빨리 끝난 명령은 감시 없음, 늘 1줄),
  `test_step_mark_is_removed_when_the_run_fails_by_itself`.

### 실측
리눅스 4코어, 에이전트 흉내(같은 명령을 다시 실행, 각 실행을 `timeout -s KILL 120`으로 끊음):
- `big.csv`(10000행 x 1000열 회귀, 한 번에 994초, 준비 53초): 기본 제한(100초)으로 17번에 `VERIFIED`. 실행당 17~89초, 모두 1줄 JSON, 강제 종료 0번
  (2.9.66: 실행 6번째부터 120초 강제 종료 반복). 1번째 실행이 준비를 저장하고 HistGB fold 3개까지, 이어 실행은 준비를 다시 쓰는 데 4~6초.
  리더보드·`oof_y.csv`·`correlation.csv`·`model_dataset.csv`(바이트)·`model_y.json`·최종 모델 예측·`top_features`·`warnings`가 한 번에 학습한 것과 같음.
- 같은 데이터 `ML_KIT_TIME_LIMIT=50`(fold 하나 72초가 제한보다 김): 4번째 실행에서 감시가 65초에 `CONTINUE`(저장된 진행 있음), 5번째 `CONTINUE`(다시 시도 1/2),
  6번째 `STOP_RETRY` + `[추천] ... --time-limit 570`과 오류 로그 1건. 모두 1줄 JSON, 무한 반복 없음.
- `w300.csv`(5000행 x 300열): 2번(2.9.66: 5번), `h3.csv`(60000행 x 16열): 1번(35초). 두 경우 모두 결과가 한 번에 학습한 것과 같음.

## [2.9.66] - 2026-11-04

### 기능
- 학습을 후보 모델 안에서도 교차검증 fold 단위로 나눠 진행 (R10). 지금까지는 후보 모델 단위로만 저장해, 후보 하나의 교차검증(fold 5개)이
  명령 시간 제한(120초)을 넘으면 그 후보를 매번 fold 1부터 다시 해 끝나지 않았음(무응답 반복). 이제 fold가 끝날 때마다 그 fold의 결과
  (OOF 예측 조각, 학습 점수, 걸린 시간)를 `.resume_train/<키>.m<번호>.f<fold>.joblib`에 저장(`_cv_folds`, 대체 위치 JIT 복구 포함)하고 다음 fold 전에
  시계를 봄: 이 후보의 fold 실측(저장된 것 포함) 최대값 x 1.2, 후보의 첫 fold는 이번 실행의 다른 fold 시간과 저장된 후보 시간 / fold 수의 최대값 x 1.2,
  값이 없으면 진행. 시계는 이번 실행에서 한 단계(fold, 후보, 전체 적합, 타겟)를 끝낸 뒤에만 봄(`_step_done`, 실행마다 최소 1단계). 모든 fold가 끝난
  후보는 지금처럼 `<키>.m<번호>.joblib`로 합쳐 저장하고 fold 파일은 지움. 손상되었거나 맞지 않는 fold 파일은 버리고 그 fold만 다시 학습(`[WARN]`).
  분할기·시드·셔플·중복 행 묶음 그대로라 리더보드·OOF·최종 모델 예측·중요도가 한 번에 학습한 것과 같음. `progress`:
  `타겟 '<이름>' 후보 모델 k/5개 끝남, '<모델>' 교차검증 fold j/5개 끝남, 저장됨 (fold j+1 전에 멈춤)`.
  대체 위치 정리(`resume_sweep_old`/`resume_clear`)의 도구 파일 이름에 `<키>.m<번호>.f<fold>.joblib` 추가.
  후보 전 시간 확인(`_candidate_check`)은 fold 확인으로 바뀌고, 1위 전체 적합 전 확인은 `_fit_check`.
  테스트: `test_train.py` `test_slow_candidate_is_split_by_fold_and_matches_one_run`(회귀·분류 x 중복 행 묶음),
  `test_killed_run_keeps_saved_folds_and_damaged_fold_is_trained_again`, `test_fold_clock_check_estimates`,
  명령 단위 동일성 `_assert_same_version`(여러 타겟 `test_training_stops_at_time_limit_and_the_same_command_continues`,
  `--best` `test_best_continue_skips_finished_task_and_checks_time`: 리더보드·OOF 파일 바이트·최종 모델 예측), 실행 횟수 상한 갱신
  (`test_report_export.py` `test_report_best_retraining_continues` 포함).
  실측(리눅스 4코어, 각 실행을 timeout으로 끊어 에이전트 제한 흉내): `h3.csv`(60000행 x 16열, 한 번에 43초) `ML_KIT_TIME_LIMIT=30`(여유를 빼면 0초:
  실행마다 1단계)·timeout 50초 42번(fold 25, 전체 적합 1, 중요도 열 16), 실행당 11~14초, 모두 1줄 JSON. `=60`·timeout 80초 2번(28초, 26초).
  `w300.csv`(5000행 x 300열, 한 번에 165초) `=60`·timeout 80초 15번(실행당 20~32초, 모두 JSON). 세 경우 모두 리더보드·`oof_y.csv`(바이트)·
  최종 모델 예측이 한 번에 학습한 것과 같음.

## [2.9.65] - 2026-11-04

5차 독립 검토의 ml-kit 발견 수정. 항목마다 회귀 테스트를 추가했습니다.

### 버그 수정
- [치명] 순열 중요도(`_permutation_means`): 다른 단계를 하지 않은 실행은 처음 열의 1/5(`IMPORTANCE_MAX_RUNS`)을 시계 없이 계산하고, 끝낸 열은
  멈출 때만 저장해 600열 데이터(HistGB, KNN)에서 120초 강제 종료가 매번 반복되었음(JSON 없음, 진행 없음). 이제 시계 없이 하는 열은 첫 열 1개뿐(R10 진행 보장
  최소)이고 그 뒤는 열마다 시계를 봄. 끝낸 열은 `IMPORTANCE_SAVE_SECONDS`(12초)마다 저장(처음 저장할 때 1위 모델 전체 적합도 한 번 저장)하므로
  밖에서 강제 종료되어도 다음 실행이 이어 감. 열 시간 추정은 첫 열 뒤부터 이번 실행의 실측 열 시간(최대값 x 1.2)으로 바꿈(전에는 첫 예측 추정과의
  최대값이라 첫 예측이 느리면 계속 일찍 멈춤). `IMPORTANCE_MAX_RUNS` 삭제. 테스트: `test_train.py`
  `test_slow_wide_importance_stops_by_the_clock_from_the_second_column`, `test_killed_importance_keeps_columns_saved_on_the_way`,
  `test_wide_target_last_step_is_split_and_matches_one_run`·`test_wide_data_command_split_runs_end_like_one_run`(상한 갱신).
- [개선] 1위 모델 전체 적합 전 시간 추정을 가장 느린 후보(지금까지는 저장된 후보 시간의 최대값 x 1.2)가 아니라 1위 모델 자신의 교차검증 시간
  x 1.6 / (fold 수 - 1)로 바꿈(`_fit_estimate`): 진 RandomForest가 느리면 빠른 전체 적합 앞에서 멈춰 한 실행을 더 썼음. 5000행 x 300열 타겟 1개
  (`w300.csv`, 120초 강제 종료 조건)의 실행 횟수: 2.9.64 12번(첫 실행 JSON 없음) -> 2.9.65 5번(모두 JSON, 실행당 21~102초).
  테스트: `test_last_step_estimate_uses_saved_candidate_times`, `test_single_target_is_split_by_candidate_model`(갱신).
- [치명] 대체 저장 위치 정리(`resume_sweep_old`)가 `<ML_KIT_DATA_DIR>/resume/` 아래 모든 파일·폴더를 이름 확인 없이 3일 지나면 지웠음
  (`ML_KIT_DATA_DIR`를 사용자 폴더로 두면 `resume\2025\cv.docx` 삭제). 이제 데이터 폴더의 대체 폴더 이름은 `ml-kit-resume`, 지우는 것은 12자리
  16진수 하위 폴더 안의 도구 파일 이름(`<키>.joblib`, `<키>.m<번호>.joblib`, `<키>.fit.joblib`, `<키>.imp.joblib`, `best_run.json`, 각 `.tmp`)뿐이고
  그 밖은 손대지 않음. 끝난 실행의 대체 위치 정리(`resume_clear`)도 같은 이름만. 옛 `resume` 폴더는 쓰지도 지우지도 않음.
  테스트: `test_train.py` `test_resume_sweep_never_touches_user_files`, `test_resume_find_newest_and_old_fallbacks_swept`(파일 이름 갱신).
- [경미] 코드 열 규칙 (a): `01, 001, 02, 002`처럼 같은 수를 0 개수만 다르게 적은 칸이 있어도 코드 열이 아니어서 `c = 001`이 `01` 행까지 뺐음.
  이제 (a)는 같은 수를 다르게 적은 칸이 둘 이상(그중 하나는 앞자리 0)이면 코드 열. 테스트: `test_code_column_judgment`(사례 추가),
  `test_where_thousands_and_code_columns`.
- [경미] `--unit-row`의 중량 단위 `킬로그램중`, `그램중`, `톤중`을 '중'으로 끝나는 상태 낱말로 보아 경고·거부했음. 단위 목록(`_KOREAN_UNIT_WORDS`,
  상태 규칙보다 먼저 봄)에 추가. 테스트: `test_loading_layout.py` `test_named_unit_row_with_korean_units_is_read`(사례 추가).

## [2.9.64] - 2026-11-04

4차 독립 검토의 ml-kit 발견 수정. 항목마다 회귀 테스트를 추가했습니다.

### 버그 수정
- [치명] 학습의 마지막 단계(1위 모델 전체 적합 + 순열 중요도)가 열이 많으면(5000행 x 200~300열) 120초를 넘어 JSON 없이 끊기고, 다시 실행해도
  같은 단계를 반복했음. 이제 순열 중요도를 직접 계산하며(열마다 난수 (시드, 열 번호), 예측 5번) 전체 적합 뒤 예측 1번의 시간을 재어 열마다 시계를
  보고, 넘을 것 같으면 전체 적합 모델(`<키>.fit.joblib`)과 끝낸 열(`<키>.imp.joblib`)을 저장하고 `CONTINUE`(`progress`: `순열 중요도 k/N열 끝남`).
  실행마다 열의 1/5 이상(`IMPORTANCE_MAX_RUNS` 5)을 하므로 유한 횟수에 끝나고, 나눠 계산한 결과(리더보드·모델·중요도)는 한 번에 학습한 것과 같음.
  후보 저장 파일에 실측 시간을 함께 남겨 다음 후보·전체 적합 전 추정은 저장된 후보 시간(재사용 포함)의 최대값을 씀(후보를 모두 재사용한 실행 포함).
  `sklearn.inspection.permutation_importance` 대신 `_permutation_means`(같은 방법: 기준 점수 - 열을 섞은 점수의 평균, 회귀 R2·분류 Macro-F1).
  공통 `time_would_pass`(PART C). 결과 JSON `top_features`의 비중은 반올림해도 합이 1을 넘지 않게 맞춤(가장 큰 값에서 뺌). 테스트: `test_train.py` `test_wide_target_last_step_is_split_and_matches_one_run`(회귀·분류),
  `test_last_step_estimate_uses_saved_candidate_times`, `test_full_fit_after_reused_candidates_is_checked`(갱신).
- [높음, 2.9.63 회귀] `--unit-row`로 준 정상 한글 단위 행(`세/백만원/cm`, `억원`, `킬로/km`, `기가바이트/Mbps`, `개수/EA`, `초/센티`,
  `점/퍼센트포인트`, `킬로` 한 칸)을 낱말 목록에 없다고 거부했음. 이제 사용자가 준 행의 짧은 한글 낱말(공백 없는 6자 이하, 아래 열이 대부분 숫자)은
  단위이고, 거부는 상태·결측 표시 낱말(`점검중`, `센서고장`, `정지` 등)이 글자 칸의 절반 이상일 때만(아래 열의 값과 같은 칸이 든 행은 전처럼 데이터 행).
  경고도 그런 낱말이 있을 때만. 단위 목록 보강(세, 살, 억원, 만원, 백만원, 천원, 원, 킬로, 센티, 미리, 바이트~기가바이트, 개수, 마리, 명, 점,
  퍼센트포인트, 배, 대, 건). 테스트: `test_loading_layout.py` `test_named_unit_row_with_korean_units_is_read`; `점검중,센서고장,정지` 거부는 그대로.
- [중간] 코드 열 규칙 (b)(앞자리 0 칸 50% 이상)가 월 `01`~`12`를 코드 열로 보아 `--exclude-where "월 = 1"`이 0행(학습은 숫자 1로 읽음). (b)를 없앰:
  코드 열은 (a) `007`과 `7`처럼 같은 수를 다르게 적었거나 (b) 모든 칸이 3자리 이상 같은 자릿수 코드일 때만. `ml_map`도 같은 함수.
  테스트: `test_train.py` `test_where_on_zero_padded_month_column_compares_numbers`, `test_code_column_judgment`(기대값 정리).
- [낮음] 중간 저장: `resume_find`가 여러 위치 중 가장 최근 파일을 읽고, 대체 위치에 저장하면 앞 위치의 옛 파일을 지움(`best_run.json` 고정 이름).
  명령 시작 때 대체 폴더 전체에서 3일 넘은 항목을 지움(`resume_sweep_old`). 저장 실패 `STOP_RETRY`의 directive가 "열 이름과 --eval 식" 대신
  쓰기 권한·디스크·경로 길이를 안내. 테스트: `test_resume_find_newest_and_old_fallbacks_swept`, `test_resume_saved_nowhere_stops_with_stop_retry`(보강).

## [2.9.63] - 2026-11-04

3차 독립 검토의 ml-kit 발견 수정. 항목마다 회귀 테스트를 추가했습니다.

### 버그 수정
- [치명] 중간 저장 실패 때 2.9.62가 "나누지 않고 끝까지" 진행해 큰 데이터(6만 행)에서 120초를 넘겨 JSON 없이 끊겼음. JIT 복구로 바꿈:
  `.resume_train`/`best_run.json`/성능예측표 `_fits`를 원래 위치에 저장하지 못하면 데이터 폴더 `resume\<해시 12자>\`, 그다음 `%TEMP%\ml-kit-resume\<해시>\`에
  저장하고 `recovered`와 `[WARN]`. 이어 실행은 모든 위치를 찾고, 끝나면 대체 위치도 지움. 어디에도 저장하지 못하면 `check_time`이 멈출 지점에서
  `STOP_RETRY`(쓰기 권한·디스크·경로 길이 안내, 1줄 JSON). 저장 파일 이름은 키 해시 16자(전 32자)로 줄여 MAX_PATH 여유.
  공통 함수 `resume_places`/`resume_save`/`resume_find`/`resume_clear`/`save_failed_warning`(PART C), `split_off`·`SPLIT_OFF` 삭제.
  테스트: `test_train.py` `test_blocked_resume_folder_continues_in_fallback`(타겟 1개·2개), `test_resume_saved_nowhere_stops_with_stop_retry`,
  `test_resume_save_checks_the_file`, `test_report_export.py` `test_export_blocked_fit_cache_continues_in_fallback`.
- [중간] `is_code_column`이 앞자리 0 칸 하나로 코드 열이라 판정해 `ml_map`이 숫자 열(`05,10,15,20`)을 보간하지 않고 틀린 `[후보]`. 이제 앞자리 0 칸이 있고
  (a) 같은 숫자의 0 없는 표기가 함께 있거나, (b) 앞자리 0 칸이 절반 이상이거나, (c) 모든 칸이 3자리 이상 같은 자릿수 코드일 때만 코드 열.
  `--where`/`--exclude-where`와 `ml_map`이 같은 함수. 테스트: `test_train.py` `test_code_column_judgment`,
  `test_where_on_two_digit_number_column_compares_numbers`, `test_map.py` `test_two_digit_numbers_with_a_leading_zero_are_interpolated`.
- [중간] 코드 열에서 `--exclude-where "월 = 1"`이 `01` 행(학습은 숫자 1)을 말없이 빼지 않았음. 이제 `warnings`에 `'01' N행은(는) 표기가 달라 빼지 않음`
  (`!=`는 `함께 뺌`) + `[추천] <전체 명령>`, `next_commands[0]`에 그 표기를 포함한 조건(`월 = 1 | 월 = 01`)의 명령. 테스트:
  `test_where_on_code_column_warns_other_notation`(추천 명령을 실행하면 경고 없이 12행을 뺌).
- [낮음~중간] 데이터 행(`점검중,센서고장,정지`)을 `--unit-row`로 주면 받아 피처 이름이 `온도(점검중)`. 이제 한글 단위 낱말 목록(`_KOREAN_UNIT_WORDS`,
  `--unit-row`로 준 행에만 씀)에 없는 한글 낱말이 그 행 글자 칸의 절반 이상이면 `STOP_RETRY` "단위 행이 아닙니다" + `[추천]`(--unit-row를 뺀 명령),
  일부면 받아들이고 `warnings`에 `2행을 단위로 읽음: 압력(점검중) ... 데이터 행이면 --unit-row를 빼세요`. 도구가 스스로 판단할 때의 목록
  (`_KOREAN_UNITS`)은 그대로. 테스트: `test_loading_layout.py` `test_data_row_given_as_unit_row_is_refused`,
  `test_some_odd_words_in_unit_row_are_read_with_a_warning`; `test_korean_unit_words_in_named_unit_row`(도씨·킬로그램·기압·섭씨·회전·평)는 그대로 통과.

## [2.9.62] - 2026-11-04

2차 독립 검토의 ml-kit 발견 수정. 회귀 테스트는 수정 전 실패를 확인했습니다.

### 버그 수정
- [치명] `--exclude-where`/`--where`: 숫자 열에 `측정불가` 같은 결측 표시 한 칸만 있어도 글자 열로 보아 `압력 = 2`가 `2.0`과 맞지 않고
  `!= 2`가 모든 행에 맞아 틀린 행을 빼고도 `VERIFIED`였음. 이제 `=`/`!=`의 숫자 값은 숫자로 읽히는 칸과 숫자로 비교(`1,000` = `1000`,
  학습과 같은 `coerce_numeric`)하고, 글자 비교는 앞자리 0 코드일 때만(코드 열, 또는 글자 칸이 있는 열에 `= 01`). 결측 표시 칸은 열 판정에서 뺌.
  `= 2`, `!= 2`, `!= 2; 온도 > 180`의 행 수는 2.9.59와 같음. 공통 함수 `is_code_column`(PART C).
  테스트: `test_train.py` `test_where_number_column_with_missing_mark_compares_numbers`, `test_where_thousands_and_code_columns`.
- [치명] 학습 나눠 진행(R10): 중간 저장(`_resume_save`, 후보 모델과 타겟)이 실패해도 `[WARN]`만 내고 시간을 기록해 `저장됨`이라며 `CONTINUE:`를 냈고,
  같은 명령을 반복해도 끝나지 않았음. 이제 저장 뒤 파일이 있는지 확인하고, 실패하면 그 실행은 나누지 않고 끝까지 진행(`split_off`,
  `check_time`이 멈추지 않음) + `warnings`에 `중간 저장 실패로 나누지 않음: ...`. `--best`의 `best_run.json`, 성능예측표 후보 저장(`_fits`)도 같음.
  타겟 1개·여러 개 모두. 테스트: `test_failed_resume_save_is_not_split`(`.resume_train`을 파일로 만들어 저장 실패), `test_resume_save_checks_the_file`.
- [중요] 목록에 없는 한글 단위(`도씨`, `킬로그램`, `기압`, `섭씨`, `회전`, `평`)를 메모로 보아 `--unit-row 2`가 `STOP_RETRY`(틀린 이유와 `[추천]`)였고,
  예측은 단위 행을 데이터로 예측(`VERIFIED`). 이제 사용자가 준 `--unit-row` 행은 공백 없는 6자 이하 한글 낱말도 단위
  (`unit_row_verdict(..., given=True)`). 도구가 스스로 판단할 때는 그대로 한글 낱말을 단위로 보지 않고(`점검중`, `센서고장` 첫 데이터 행 보존),
  예측 레이아웃 복구는 `--unit-row`로 읽어 모델 이름(`온도(도씨)`)이 모두 있을 때만 그 행을 단위로 씀.
  테스트: `test_loading_layout.py` `test_korean_unit_words_in_named_unit_row`(학습 성공, 예측 `sample_count` = 데이터 행 수).
- [기존] `ml_map.py`가 코드 `007`과 `7`을 같은 숫자 7로 합쳐 보간(`VERIFIED`). 이제 앞자리 0 코드가 있는 열은 숫자 열이 아니라(`_is_numeric_col`이
  `is_code_column`을 봄) `-i`의 같은 코드 행만 쓰고, 없는 코드(`코드=07`)는 `[후보]`(10개까지). 테스트: `test_map.py`
  `test_code_column_is_matched_exactly_not_interpolated`.

### 구조
- 시간 확인: 이번 실행에서 한 단계(후보 또는 타겟)를 끝낸 뒤면 데이터 읽기 직후 첫 후보, 후보를 모두 다시 쓴 타겟의 1위 전체 적합·순열 중요도 앞에서도
  `check_time`(전에는 이번 실행에서 후보를 학습했을 때만). 아직 아무 단계도 하지 않은 실행은 멈추지 않음(저장 없이 멈추면 끝나지 않음, R10).
  테스트: `test_full_fit_after_reused_candidates_is_checked`.
- `--best`가 오류로 끝나면 `.resume_train/best_run.json`을 지움(나눠 진행 `CONTINUE:`와 사용자 중단은 남김). 테스트: `test_best_run_record_is_removed_after_an_error`.

## [2.9.61] - 2026-11-04

학습 나눠 진행(R10) 버그 수정 (`ml_train.py`). 회귀 테스트는 수정 전 실패를 확인했습니다(단일 타겟은 새 동작).

### 버그 수정
- [중요] `--best`를 `CONTINUE`로 나눌 때, 회귀 최고 버전을 만든 뒤 분류 작업 유형에서 멈추면 다음 실행이 회귀 최고 버전을 다시 만들어 원장에
  중복(v3, v5)되었음. 이제 만든 최고 버전을 `.resume_train/best_run.json`에 적고, 이어 실행 때 그 버전이 `best_is_fresh`면 건너뜀
  (결과의 `best_versions`·`best_targets`에는 그대로 들어감). 다 끝나면 기록을 지움.
- [중요] 타겟 학습 시간 기록(`fit_seconds`)이 `train_version` 호출마다 새로 시작해, 두 번째 작업 유형의 첫 타겟은 시간 확인 없이 학습했음.
  이제 명령 단위(`RUN_FIT_SECONDS`)라 `--best`와 `ml_report.py -f`(최고 버전 다시 학습)도 같은 확인.
  테스트: `test_train.py` `test_best_continue_skips_finished_task_and_checks_time`.

### 기능 (검토 항목 7)
- 타겟이 1개면 나눌 곳이 없어 100초를 넘으면 JSON 없이 끊길 수 있었음. 후보 모델(5개)마다 교차검증 결과를 `.resume_train/<키>.m<번호>.joblib`에
  저장하고 다음 후보와 1위 모델 전체 적합 전에 `check_time`(가장 오래 걸린 후보 x 1.2). 같은 명령을 다시 실행하면 저장된 후보를 씀.
  결과(점수, 모델)는 한 번에 학습한 것과 같음. 성능예측표(`ml_export.py`)의 후보 학습은 이 저장을 쓰지 않음(자체 `_fits`).
  테스트: `test_single_target_is_split_by_candidate_model`. 기존 나눠 진행 테스트 2개는 실행 횟수 대신 진행 문구와 결과로 확인.

## [2.9.60] - 2026-11-04

독립 검토에서 나온 예측 쪽 버그 수정 (`ml_predict.py`, 학습의 `--exclude-where` 포함). 회귀 테스트는 수정 전 실패를 확인했습니다.

### 버그 수정
- [치명] `--sweep`, `--maximize/--minimize`, `--find`의 `-i` 글자 입력을 학습 범주와 대조하지 않았음: `-i "라인=라인C"`나 `라인=c`가 어느 범주도 아닌
  기준값으로 예측되고 경고 없이 `VERIFIED`. 이제 공통 함수 `match_categories`/`fixed_category_value`(PART C)로 대조:
  대소문자·앞뒤 공백·전각만 다르고 학습 범주 하나와 맞으면 그 범주로 맞추고 `recovered`, 그 밖은 `STOP_RETRY` + `[후보] <학습 범주>`(10개까지).
  일반 예측(`-i`, `-f`)도 같은 함수(`align_categories`)로 맞추고, 남은 값만 `unseen_category_warnings`. 빠진 글자 입력의 `[후보]`도 10개까지(R8).
  같은 `recovered` 문구는 한 번만 적음(`note_recovered`, 타겟이 여럿일 때).
  테스트: `test_predict.py` `test_whatif_unknown_category_stops_with_candidates`, `test_whatif_category_case_width_space_recovered`,
  `test_predict_category_case_matched_like_whatif`, `test_match_categories_ambiguous_not_fixed`.
- [중요] 예측 파일 레이아웃 자동 복구가 첫 데이터 행에 글자(`점검중`, `센서고장`)가 섞였다고 단위·메모 행으로 보고 그 데이터 행을 버렸음.
  숫자·날짜가 하나라도 든 행과 글자 칸이 단위 모양이 아닌 행(한글 낱말)은 데이터로 남김(`_unit_shaped`, 단위 행 판단 `unit_row_verdict`도 같은 기준).
  행을 건너뛰면 원본 행 번호와 내용을 `warnings`에. 레이아웃 옵션을 주지 않았을 때 `recovered`는 `머리글 N행, 데이터 K행부터로 읽음`
  ("레이아웃 옵션이 이 파일과 달라"는 옵션을 줬을 때만). 테스트: `test_loading_layout.py` `test_first_data_row_with_text_marks_is_kept`,
  `test_skipped_unit_row_is_in_warnings`, 기존 `recovered` 기대값 5곳.
- [중요] `--exclude-where`/`--where`의 표기 고치기가 값 안의 `포함`/`and`/`or`까지 바꿔 틀린 `[추천]`을 냈음(`비고 = 불량 포함 시료` -> `비고 = 불량 ~ 시료`).
  이제 조건을 잇는 자리에서만 바꾸고(`_where_words`), 다른 표기를 고친 `[추천]`에서는 그런 값을 작은따옴표로 감싸고 "값을 작은따옴표로 감싸세요" 안내.
  테스트: `test_train.py` `test_where_words_inside_value_are_kept`.
- [경미] `!=`가 빈 칸까지 맞는 행으로 셌음 -> 빈 칸은 맞지 않음(글자 칸은 숫자 값과 다른 값이라 맞음, README 9)). 크기 비교(`> < >= <=`) 값이
  숫자·날짜가 아니면 조용히 0행이었음 -> `9 0`처럼 공백만 빼면 숫자면 `[추천]`(전체 명령), 그 밖은 `STOP_RETRY`. `= 01`은 숫자 열이면 숫자로,
  글자 열(글자 칸이나 앞자리 0 코드가 있는 열)이면 글자로 비교. 테스트: `test_where_not_equal_keeps_blank_and_size_value_checked`,
  `test_where_code_text_column_compares_as_text`.
- [경미] 조건 찾기(`--maximize/--minimize`) directive "찾은 조합을 모델로 다시 예측해 조건을 확인했습니다"가 사실이 아니었음. 이제 답마다 적힌
  `inputs`로 다시 예측해 `predictions`와 `--where`를 확인하고(`_recheck_solutions`), 다르면 `UNVERIFIED` + `unverified`.
  테스트: `test_predict.py` `test_search_rechecks_answers`.

### 문서
- SKILL.md: `--exclude-where` 값의 작은따옴표, 4-C단계 글자 입력 `[후보]`/`recovered` 한 줄씩. README: 레이아웃 복구 문구·건너뛴 행, `--exclude-where` 규칙, 범주 맞추기, 조건 찾기 확인.

## [2.9.59] - 2026-11-04

결과 문구를 짧은 한국어로 통일 (동작 변경 없음). 소비자인 경량 LLM(Gemma 26B)이 사용자에게 한국어로 전달하는데 `agent_directive`·`message`·
`preview`·경고에 영어와 한국어가 섞여 있었음. 첫 단어(`VERIFIED:` `UNVERIFIED:` `CONTINUE:` `STOP_RETRY:` `ERROR:`), `[추천]`/`[후보]` 표지,
JSON 키, `error_type` 값, 열·모델·옵션 이름은 그대로.

### 바뀐 대표 문구 (전 -> 후)
- 학습: `VERIFIED: v1 competitive training completed. Champion models and leaderboard packaged into 'v1'. Task=regression, Mean 검증 R2=0.9486. Next step: ...`
  -> `VERIFIED: v1 학습을 마쳤습니다 (regression, 평균 검증 R2=0.9486). ... 다음: 사용자가 원하면 next_commands[0]을 실행하거나 예측으로 넘어갑니다. ...`
- 학습 preview: `(Delta: +0.0123)`, `소요시간: 8.61s`, `산출물 저장:` -> `(변화: +0.0123)`, `소요 8.61초`, `저장:`
- 이력: `VERIFIED: History summary returned. Do not re-read history.jsonl ...` -> `VERIFIED: 학습 이력입니다. history.jsonl을 다시 읽지 마세요.`
- 예측: `VERIFIED: Prediction completed for 3 target(s) with extrapolation warning (...). Output saved to x.csv. Next step: report inference results to user.`
  -> `VERIFIED: 타겟 3개를 예측했습니다. 학습 범위 밖 입력이 있어 예측은 참고용입니다 (extrapolation_warnings). 결과 파일: x.csv. 위 경고를 함께 알리세요.`
  (저확신, 빈 입력, 학습 때 없던 종류, 실측 비교 경고, 잘린 직접 입력도 같은 식의 한국어 문장)
- 예측 preview·`source`: `[추론 완료] 소스: Inline Input (1 record(s)) | 타겟 예측 요약: [...] | 인라인 반환` -> `[예측 완료] 입력: 직접 입력 1건 | 예측 요약: [...]`
- 외삽 경고: `피처 'X' 상한 초과 외삽(Extrapolation) 감지: ... (범위 대비 +50.0% 초과 이탈, 범위 밖 3건)` -> `입력 'X' 학습 범위 초과: ... (범위의 +50.0%, 범위 밖 3건)`
- 보고서: `ML result report for v1 (1 target(s)) saved to '...'. Overall verdict=적합. Next step: report the file path to the user.`
  -> `v1 ML 결과 보고서(타겟 1개)를 '...'에 저장했습니다. 종합 판정 적합. report_file 경로를 사용자에게 알리세요.` (`report_file 링크` -> `경로`)
- 입력 오류: `STOP_RETRY: Argument or data schema validation failed. If the message contains '[추천]', re-run ONCE ...`
  -> `STOP_RETRY: 입력이나 데이터 문제로 멈췄습니다. message에 [추천]이 있으면 그 값으로 바꿔 1번만 다시 실행하세요 (... 전체 명령이니 그대로 1번 실행). 없으면 message를 그대로 알리세요.`
  뒤에 붙는 안내(`Check column names or formula syntax.` 등) -> `확인할 것: 열 이름과 --eval 식.` 등.
- 인자 오류: `STOP_RETRY: Invalid CLI arguments. Check the option names in SKILL.md ...` -> `STOP_RETRY: 명령 인자가 잘못되었습니다. 다시 실행하지 말고 message를 그대로 알리세요 (옵션 이름은 SKILL.md 기준).`
- 중단: `STOP_RETRY: Execution was interrupted. Ask the user before running again.` -> `STOP_RETRY: 실행이 중단되었습니다. 다시 실행하기 전에 사용자에게 물으세요.`
- 내부 오류: `ERROR: Unexpected internal failure. Do not retry; ...` -> `ERROR: 도구 내부 오류입니다. 다시 실행하지 말고 message를 그대로 알리세요.`
- 누수: `[DATA LEAKAGE DETECTED] ...` -> `타겟 누수: ...`
- `UNVERIFIED:` 끝맺음을 `다시 실행하지 말고 unverified를 그대로 알리세요.`로 통일 (`재시도하지 말고 ...` 혼용 제거).
- 모델 읽기 실패 `Joblib 모델 로드 실패` -> `모델 파일을 읽지 못했습니다`, `-h` 설명 문구, 예측 stderr `[WARN]` 머리말도 한국어.

테스트: `test_contract.py`의 누수·`[추천]` 전체 명령 문구 기대값을 새 문구로.

## [2.9.58] - 2026-11-04

조건별 예측(what-if)과 제약 안의 최적 조건 찾기. 이전에는 조건을 바꿔 보려면 에이전트가 `-i` JSON 배열을 직접 만들고 결과를 비교해야 했고,
"불순물 100 이하에서 생산효율 최대" 같은 요청은 방법이 없었음. 회귀 테스트: `tests/test_predict.py` `test_sweep_*` 3개,
`test_maximize_with_constraint`, `test_minimize_impossible_constraint_is_zero`, `test_whatif_argument_rules` (2.9.57에서 6개 모두 실패).

### 새 기능 (`ml_predict.py`, 회귀 타겟만, 단독 배포 그대로)
- `--sweep "<입력>=<시작>:<끝>:<간격>" [-i "<고정 입력>"] [-y <타겟>] [--save <표>]`: 입력 하나를 바꿔 가며 예측.
  `<시작>:<끝>`만 주면 21점, `<입력>`만 주면 학습 범위 21점, 500점까지(넘으면 `[추천] <간격을 넓힌 전체 명령>`).
  - 표(xlsx, 데이터 폴더의 `sweep_<입력>_<시각>.xlsx`)와 그래프(같은 이름 `.png`, 90% 띠, 학습 범위 음영. matplotlib이 없으면 표만과 경고).
  - JSON `points`(25점까지 고르게, 넘으면 `truncated`, `full_result_file`), 점마다 `in_range`, `<타겟>_90%`; `trend`(타겟마다 `direction`
    increasing/decreasing/peak/valley/flat/mixed, `max`, `min`, `text`); `out_of_range_points`와 외삽 경고.
- `--maximize|--minimize "<타겟>" --vary "<입력1>,<입력2>" [--where "<제약>"] [-i "<고정 입력>"]`: 학습 범위 안에서만 격자 탐색(약 2만 번 평가)과
  좌표 탐색. `--where`는 `--exclude-where`와 같은 식으로 타겟 예측값·입력에 겁니다. `solutions`(5개까지, 좋은 순, 서로 다른 조합,
  `inputs`, `predictions`, `interval_90`), `feasible_count`, `nearest_rows`(첫 답에 가까운 학습 행 5개). 0건이면 `VERIFIED: 0건`과 `best_without_where`.
- 말하지 않은 숫자 입력은 `--find`처럼 학습 중앙값(`filled_defaults`, `warnings`), 글자·날짜 입력은 `[후보]`로 멈춤.
- 인자 규칙: `--where`는 `--maximize`/`--minimize`와만, `--sweep`은 `--find`·`--maximize`·`-f`·`--vary`와 함께 쓰지 않음 (인자 오류).
- SKILL.md 4-C단계 (명령 2개와 알릴 것), 1단계 표에 한 줄.

## [2.9.57] - 2026-11-04

학습도 나눠서 진행 (CLI-dev/README.md R10 `CONTINUE:`). 회사 PC Gemma 시험 R1, R2처럼 타겟이 여러 개인 학습은 120초를 넘겨 결과 JSON 없이
끊기고 에이전트가 같은 명령을 처음부터 반복했음. 회귀 테스트: `tests/test_train.py`
`test_training_stops_at_time_limit_and_the_same_command_continues`, `test_damaged_saved_target_is_trained_again`,
`tests/test_report_export.py` `test_report_best_retraining_continues` (2.9.56에서 3개 모두 실패: 시간 제한을 보지 않음).

### 바뀐 동작
- `ml_train.py`(일반 학습, `--best`)와 `ml_report.py -f`의 최고 버전 다시 학습: 타겟이 2개 이상이면 5개 모델 경쟁을 마친 타겟마다
  `<데이터>_ml_data/.resume_train/<키>.joblib`에 저장하고, 다음 타겟을 학습하기 전에 시계를 봅니다
  (`check_time`, 이번 실행에서 가장 오래 걸린 타겟 시간 x 1.2). 제한(`ML_KIT_TIME_LIMIT`, 기본 100초, 여유 30초)을 넘을 것 같으면
  `status: ok`, `continue: true`, `progress`("학습: 타겟 1/3개 끝남, 저장됨 ..."), `next_commands: [같은 명령]`, `CONTINUE:` directive.
- 같은 명령을 다시 실행하면 저장된 타겟을 다시 쓰고(`[INFO]`) 다음 타겟부터 학습합니다. 끝난 실행의 결과에 `resumed_targets`(앞 실행이 학습한 타겟).
  키는 학습 입력 그 자체(학습 행과 값, 스키마, 작업 유형, 시드, 양성 클래스, 도구·scikit-learn 버전)의 해시라, 데이터나 조건이 바뀌면 다시 학습합니다.
- 멈춘 실행은 스테이징 폴더를 지우고 버전 폴더와 원장 기록을 남기지 않습니다. 버전 번호와 `history.jsonl`은 끝나는 실행이 한 번만 씁니다
  (원자적 커밋은 그대로). 멈춤은 비정상 중단이 아니므로 `[WARN] 비정상 중단` 대신 `[STATUS]`만. 오류 로그에도 남기지 않습니다.
- 커밋 뒤 이번 실행이 쓴 저장 파일과 3일 넘은 저장 파일을 지웁니다. 손상된 저장 파일은 버리고 다시 학습합니다(`[WARN]`, JIT).
- 실행마다 새 타겟을 최소 1개 학습하므로 반복하면 반드시 끝납니다. 타겟이 1개면 나눌 단계가 없어 저장하지 않습니다.
- SKILL.md 규칙 6: "학습, 보고서, 성능예측표는 오래 걸리면 도구가 나눠서 진행한다".

## [2.9.56] - 2026-11-04

조건으로 학습에서 뺄 행 (`--exclude-where`). 회사 PC 시험 A3("생산효율이 90보다 큰 행은 빼고 다시 학습해줘")에서 에이전트가 할 수 있는 일이
`제외` 열에 직접 적으라는 안내뿐이었음. 회귀 테스트: `tests/test_train.py` `test_exclude_where_*` 4개, `test_where_mask_rules`
(2.9.55에서 5개 모두 실패: 옵션 없음).

### 새 기능
- `ml_train.py --exclude-where "<조건>"`: office-cli `excel.py read --where`와 같은 식 (`;` 그리고, `|` 또는(`;`가 먼저), 연산자 `>= <= > < = != ~`,
  `;`·`|`가 든 값은 따옴표). 숫자는 숫자로(`1,500` = 1500), `YYYY-MM-DD`는 날짜로, 그 밖은 글자로(대소문자 무시) 비교. 데이터 파일은 바꾸지 않습니다.
  - 열 이름은 `-x`/`-y`와 같이 맞춤(`resolve_column_names`: 단위 없는 고유한 이름, 비슷한 이름은 `[후보]`). `==`, `<>`, `AND`, `OR`는
    `[추천] <고친 식의 전체 명령>`, 읽을 수 없는 식은 형식 안내.
  - 조건에 맞는 행은 모든 타겟의 학습에서 빠짐: `exclusions.<타겟>.condition`(행 수)·`condition_text`, 이유 `조건: <열 이름을 맞춘 식>`
    (`excluded_rows.csv` 구분 `조건`, `excluded_preview`의 `kind` `조건`, `excluded_reasons`), preview `학습 제외 N행(직접 a, 조건 b, 자동 c)`.
    `제외` 열에 이유를 적은 행은 직접 표시로 셈.
  - `제외` 열이 `포함`인 행은 조건에 맞아도 학습에 넣고 `warnings`에 그 행을 적음. 조건에 맞는 행이 없으면 `warnings`에 알림.
  - 조건으로 빼서 학습 행이 10개 미만이면 `STOP_RETRY`(남는 행 수).
  - 원장 `exclude_where`(맞춘 식), 개선 `next_commands`의 재학습 명령과 `--best`(`target_recipes.<타겟>.exclude_where`)가 같은 조건을 이어받음.
    `model_dataset.csv`의 `학습제외_<타겟>`에도 표시되어 예측의 실측 비교값이 그 행을 쓰지 않습니다.
- 공통 런타임(PART C)에 조건 식 함수: `parse_where`, `resolve_where`, `where_mask`, `where_text` (예측의 조건 검색에도 씀).
- 보고서: 학습 제외 행 줄에 조건 행 수와 식.
- SKILL.md 2단계: 옵션 표에 한 줄, 행 빼기 안내를 `--exclude-where`(조건) / `제외` 열(특정 행)로. `agent_scenarios.yaml` A3를 이 옵션으로 채점
  (샘플 데이터의 생산효율 최대가 84라 기준을 80으로), 메시지 팩(`docs/opencode/자동테스트`) 다시 만듦.

## [2.9.55] - 2026-11-04

사용자가 정할 일을 도구가 고를 수 있는 명령으로 바꿔 줌. 회귀 테스트: `tests/test_contract.py`
`test_train_target_in_features_offers_both_commands`, `test_predict_missing_inputs_offer_median_fill`, `tests/test_train.py`
`test_prediction_uses_source_and_caps` (2.9.54에서 3개 모두 실패).

### 바뀐 동작
- 학습: 같은 열을 `-x`와 `-y`에 함께 주면(`[DATA LEAKAGE DETECTED]`) `STOP_RETRY` + `[후보] 1) <그 열을 -y에서 뺀 전체 명령> / 2) <-x에서 뺀 전체 명령>`.
  두 명령은 `next_commands`(`purpose`, `command`)에도 있고 다른 옵션은 그대로입니다. `-y`가 그 열뿐이면 2)만, `-x`가 그 열뿐이면 1)만.
  이전에는 "종속변수는 독립변수로 사용될 수 없습니다"만 알려 에이전트가 명령을 직접 고쳐야 했습니다.
- 예측: `-i`(한 건)에 숫자 모델 입력이 빠지면 `status: error`, `error_type: "needs_user_decision"`, `missing_inputs`, `fill_values`(학습 중앙값),
  `next_commands[0]` = 빠진 입력을 중앙값으로 채운 같은 명령(`needs_user_confirm: true`). 도구는 채워서 실행하지 않습니다 (사용자 값은 도구가 알 수 없음).
  `agent_directive`는 `STOP_RETRY:`로 시작하고 사용자에게 묻고 원할 때만 `next_commands[0].command`를 그대로 1번 실행하라고 합니다.
  이름 오타·모델 입력이 아닌 이름, 빠진 글자·날짜 입력, 여러 건은 이전 안내 그대로입니다.
- 공통 런타임: `InputError(message, extra=..., error_type=...)`. `extra`의 키는 오류 JSON에 들어가고, `error_type="needs_user_decision"`은
  위 directive를 씁니다 (office-cli의 `needs_user_decision`과 같은 이름).
- SKILL.md: 4단계에 빠진 입력 안내 한 줄, 결과 처리 표에 `needs_user_decision` 행과 전체 명령 `[후보]`(고른 번호의 `next_commands` 실행).

## [2.9.54] - 2026-11-04

예측 파일의 레이아웃이 학습 파일과 다를 때 도구가 그 자리에서 고쳐 읽음 (AGENTS.md JIT 자가복구). 회귀 테스트:
`tests/test_loading_layout.py` `test_training_layout_on_plain_file_is_recovered`, `test_two_level_header_hint`,
`test_several_header_rows_with_every_input_stop_with_candidates`, `test_layout_recovery_keeps_the_unit_row`,
`test_predict_finds_model_names_with_units_in_a_header_unit_pair`, `test_predict_plain_template_on_a_unit_row_file`,
`test_unit_row_file_for_a_unitless_model`, `test_title_and_unit_row_file_for_a_unitless_model`, `tests/test_train.py`
`test_predict_next_commands_have_no_layout` (2.9.53에서 9개 모두 실패: STOP_RETRY [추천], 예측 명령에 레이아웃 옵션).

### 바뀐 동작
- `ml_predict.py -f`: 레이아웃이 파일과 맞지 않을 때 모델 입력이 모두 든 레이아웃이 **하나로 정해지면** 그 레이아웃으로 다시 읽고 `status: ok`,
  `recovered: ["레이아웃 옵션이 이 파일과 달라 머리글 N행으로 읽음"]`(단위 행이면 `, 단위 M행`, 건너뛴 행이 있으면 `, 데이터 K행부터`).
  이전에는 `STOP_RETRY` + `[추천] <고친 전체 명령>`이었습니다. 대상 (판단 규칙은 이전 `[추천]`과 같음, `_layout_fix`):
  - 준 레이아웃 옵션(학습 때 옵션을 붙인 경우)으로는 데이터가 없거나 모델 열이 없고, 옵션 없이 읽으면 모델 열이 모두 있음.
  - 모델 입력이 모두 든 머리글이 위쪽의 다른 한 행에 있음 (두 줄 머리글, 병합된 머리글, 제목 행).
  - 옵션 없이 읽은 첫 데이터 행이 단위·메모 행 (단위가 든 모델은 `--unit-row`, 단위 없는 모델은 `--data-row`로 건너뜀).
  모델 입력이 모두 든 행이 둘 이상이면 고르지 않고 `STOP_RETRY` + `[후보] --header-row 2, --header-row 4`(사용자에게 물음).
  준 옵션으로 모델 열이 모두 읽히면 옵션을 그대로 씁니다(복구 없음). 모델 입력의 일부만 다른 행에 있으면 이전처럼 `[추천]`.
- 학습 결과와 `--best`, 보고서 결과의 예측 명령(`next_commands`)에서 `--header-row/--unit-row/--data-row`를 뺐습니다. 새 파일은 구조가
  다를 수 있고, 도구가 파일마다 머리글을 찾습니다. SKILL.md 4단계의 "학습 때 옵션을 예측에도 붙인다"를 "붙이지 않는다"로.
- 내부: `ml_train.layout_flags`, `ml_report._layout_args`와 그 테스트를 없앰. `_unit_row_hint`는 `_layout_fix`에 합침.

## [2.9.53] - 2026-11-04

회사 PC 시험: 에이전트가 CP949로 저장된 `excluded_rows.csv`를 깨진 글자로 읽고 제외 이유를 지어냄. 회귀 테스트: `tests/test_train.py`
`test_marks_auto_outliers_and_all_rows_score`, `test_multiclass_predict_batch_columns` (2.9.52에서 2개 모두 실패).

### 바뀐 동작
- 도구가 쓰는 CSV(예측 결과 `--save .csv`와 기본 `pred_*.csv`, `inline_pred_*.csv`, `excluded_rows.csv`, `model_dataset.csv`)를 CP949 대신
  BOM이 붙은 UTF-8(`utf-8-sig`)로 저장합니다. Excel은 더블클릭으로 그대로 엽니다. 도구의 CSV 읽기는 인코딩을 자동 판별하므로 이전 파일도 읽습니다.
- 학습 결과 JSON에 제외 행을 직접 담습니다 (제외 행이 있을 때만): `excluded_count`, `excluded_reasons`(이유별 행 수: `제외` 열의 글,
  자동은 `자동 이상치: <열>`), `excluded_preview`(앞 20행: `target`, `row`, `kind` `직접`/`자동`, `reason`, `value`),
  `excluded_preview_truncated`. `agent_directive`는 이 키를 그대로 알리고 CSV를 열지 말라고 합니다.
- SKILL.md 2단계: `excluded_reasons`와 `excluded_preview`를 그대로 알리고 `excluded_rows_file`은 열지 않음.

## [2.9.52] - 2026-11-04

사용자 결정: 쓰지 않는 옵션 정리 (베타, 별칭 없이 없앰). 회귀 테스트: `tests/test_contract.py` `test_old_option_names_are_argument_errors`
(새 7개, 2.9.51에서 모두 실패: 옵션을 받음).

### 없앤 옵션 (주면 인자 오류, exit 2)
- `ml_predict.py --meta`: 개별 `.joblib`을 주면 모델 정보는 늘 같은 폴더의 같은 이름 `.json`(`model_<타겟>.json`)에서 읽습니다.
- `ml_predict.py --encoding`, `ml_train.py --encoding`: 결과 CSV와 `model_dataset.csv`의 인코딩은 도구가 정합니다.
- `ml_report.py -d/--data-dir`, `-v/--version`: 보고서는 `-f <데이터파일>`(권장, 최고 버전) 또는 `-m <version_dir>`만. 없는 `-m` 폴더의
  `[추천]`/`[후보]`(2.9.48)는 그대로이고, `[추천]` 명령은 `-m`만 바꿉니다.
- 이 옵션들만 시험하던 테스트(`-v 1` 숫자 버전, `-d`의 최고 버전 자동 선택 2개)를 지우고, 보고서 테스트와 실기 러너(M6, M7)는 `-m <version_dir>`로 바꿈.
  `ml_train.py --version/-v`(학습 버전 지정)는 그대로입니다.

## [2.9.51] - 2026-11-04

저장소 지침 반영: 인자 이름 통일 (AGENTS.md ACI 설계 권장 5, 베타라 별칭 없이 바꿈). 회귀 테스트: `tests/test_contract.py`
`test_old_option_names_are_argument_errors`, `test_report_targets_and_save` (2.9.50에서 6개 중 5개 실패), `tests/test_skill_docs.py`.

### 바뀐 인자 (옛 이름은 인자 오류, exit 2)
- 결과 저장: `ml_predict.py --out`/`-o`, `ml_report.py --out`/`-o` -> `--save` (`ml_map.py`, `ml_export.py`와 같음, office-cli도 `--save`).
  예측의 확장자 오류 문구도 "결과 파일(--save)", `[추천] <이름>.csv`는 `--save`의 값입니다.
- 보고서 타겟: `ml_report.py --target`/`-t` -> `--targets`/`-y` (학습·예측·성능예측표·표 보간과 같음). `--target`은 argparse가 `--targets`의
  줄임으로 받습니다(같은 뜻, 따로 만든 별칭 아님).
- SKILL.md(예측 엑셀 저장 `--save`), README 옵션 표·예시, 테스트, 실기 러너(`tools/run_win_tests.py` M5)를 함께 고침.
  `agent_scenarios.yaml`과 `docs/opencode`에는 바뀐 옵션을 쓰는 곳이 없습니다.

## [2.9.50] - 2026-11-04

저장소 지침 반영: 로그 태그. 회귀 테스트: `tests/test_contract.py` `test_stderr_uses_only_shared_log_tags` (2.9.49에서 실패).

### 바뀐 동작 (stderr만, stdout JSON 그대로)
- 학습 진행 줄 `[PROGRESS  45%] [=========-----------] <내용>`(터미널에서는 `\r` 덮어쓰기)을 `[PROGRESS] 45% <내용>` 한 줄씩으로.
- `--history` 표 머리 `[ML EXPERIMENT HISTORY LEDGER] Root: ...`, `[PROJECT DESCRIPTION] ...`을 `[STATUS] 실험 이력 (history.jsonl): ...`,
  `[STATUS] 설명: ...`으로 (허용 태그는 `[INFO] [PROGRESS] [WARN] [ERROR] [STATUS]`뿐).
- 실기 러너(`tools/run_win_tests.py`)의 `[준비]`를 `[INFO]`로. 러너 M14 검사(stdout 1줄 JSON, stderr `\r` 없음)는 진행 줄 모양을 읽지 않아 그대로입니다.

## [2.9.49] - 2026-11-04

저장소 지침 반영: 출력 양 제한(R2). 회귀 테스트: `tests/test_train.py` "JSON 목록 상한 + truncated" (2.9.48에서 3개 모두 실패).

### 바뀐 동작
- `ml_train.py --history`의 `records`는 최근 30건(`HISTORY_SHOWN`)까지. `count`(= `record_count`, 전체), `shown`, `truncated`, 잘렸으면
  `full_result_file`(`<데이터>_ml_data/history.json`, 전체 기록)과 `agent_directive`에 "최근 N건만" 안내. 이전에는 기록 수만큼 모두.
- 학습 결과의 `feature_types`는 종류(numeric, categorical, datetime)마다 30개(`FEATURE_TYPES_SHOWN`)까지. 넘으면 `truncated: true`, `feature_count`,
  `full_result_file`(`history.json`). 학습 결과에 `truncated`가 늘 있습니다(안 잘렸으면 false). 원장 기록은 그대로 전체.
- 오류 메시지의 열 목록(`[사용 가능한 컬럼 목록]`, `[모델 입력 목록]`, 보고서의 `[사용 가능한 버전]`)은 40개(`MESSAGE_LIST_MAX`)까지와 `외 N개 (전체 M개)`
  (PART C `name_list`). 열이 500개인 표의 오류 한 번이 경량 LLM의 문맥을 채우지 않게.

## [2.9.48] - 2026-11-04

저장소 지침 반영: 보고서 버전 없음의 `[추천]`. 회귀 테스트: `tests/test_contract.py` `test_report_unknown_version_*` 3개
(2.9.47에서 3개 모두 실패).

### 수정
- `ml_report.py -d <폴더> -v <없는 버전>`과 `-m <없는 버전 폴더>`가 "버전 폴더를 찾을 수 없습니다"만 알려 에이전트가 고칠 값을 몰랐던 문제.
  같은 데이터 폴더에 학습한 버전(`leaderboard.json`이 있는 `vN`)이 하나면 `[추천]` 그 폴더를 `-m`으로 준 전체 명령(다른 옵션 유지, `-d`/`-v` 뺌),
  여럿이면 `[후보] v1, v3`(다른 버전은 다른 모델이라 사용자가 고름), 없으면 먼저 학습하라는 안내.
- 예측 `--targets`(`-y`) 불일치는 바꾸지 않았습니다: 이미 `resolve_column_names`로 단위 없는 고유한 이름은 자동으로 맞추고, 비슷한 이름은 `[후보]`,
  없으면 모델의 타겟 목록을 알립니다. 다른 타겟을 `[추천]`하면 다른 양을 예측하게 되므로 2.9.33 규칙대로 `[추천]`을 두지 않습니다.

## [2.9.47] - 2026-11-04

저장소 지침 반영: `--dtypes` 이름 맞추기. 회귀 테스트: `tests/test_train.py` "--dtypes 이름과 형식 확인" (2.9.46에서 2개 모두 실패).

### 수정
- `--dtypes`의 열 이름이 정확히 같을 때만 적용하고, 다른 이름(`설비코드` / 열 `설비코드(번호)`)과 모르는 형식(`strr`)은 말없이 무시하던 문제.
  새 `resolve_dtypes`가 열 이름을 `-x`/`-y`와 같은 규칙(`resolve_column_names`: 정확한 이름, 단위 없는 고유한 이름, 비슷한 이름은 `[후보]`)으로 맞추고,
  형식을 확인합니다: 형식 오타는 `[추천] --dtypes "<고친 전체 값>"`, 비슷한 형식이 없으면 쓸 수 있는 형식(float, int, str, datetime) 안내,
  `:`가 없으면 형식 없음 안내. 원장의 `dtypes`에는 맞춘 전체 이름과 표준 형식이 남습니다.
- `int` 형식이 빈 칸을 0으로 채우던 것(`fillna(0)`)을 없앰: 숫자 열로 읽고 빈 칸은 다른 숫자 열처럼 중앙값으로 채웁니다.
- 지침의 "`[추천]`"은 이름에는 쓰지 않았습니다: 비슷한 이름은 다른 양일 수 있어 `[후보]`(2.9.33 규칙). `[추천]`은 형식 오타에만.

## [2.9.46] - 2026-11-04

저장소 지침 반영: 빈 결과(R2)와 원장. 회귀 테스트: `tests/test_contract.py` `test_history_without_records_is_ok_zero`,
`test_blank_rows_only_input_says_zero_rows`, `tests/test_safety.py` `test_existing_version_refused_then_overwrite` (2.9.45에서 3개 모두 실패).

### 바뀐 동작
- `ml_train.py --history`에 학습 기록이 없으면 `status: no_history`(그 밖의 키 없음) 대신 `status: ok`, `verified: true`, `record_count: 0`, `count: 0`,
  `records: []`, `no_history: true`, `agent_directive` `VERIFIED: 0건. 이 데이터 파일의 학습 기록(history.jsonl)이 없습니다. ...`.
- 머리글 아래가 모두 빈 행인 입력은 "머리글(N행) 아래의 M행이 모두 빈 행이라 데이터가 0행입니다"로 멈춥니다 (공통 로더라 학습·예측·표 보간 같음).
  이전에는 학습이 "숫자로 변환 가능한 값이 하나도 없습니다 ... --task classification", 표 보간이 "숫자 기준 열이 하나 이상 필요합니다"로 원인과 다르게 안내.

### 수정
- `--version vX --overwrite`로 버전을 바꾸면 원장에 같은 version 기록이 둘 남아, 새 기록의 `delta_score`·`is_best`가 바뀐 옛 폴더의 점수와 비교되고
  `--history`·보고서가 옛 기록을 볼 수 있던 문제. 원장에 그 버전 기록이 있으면 빼고 새 기록을 붙여 다시 씁니다(원자적 쓰기, stderr `[INFO]`),
  `delta_score`·`is_best`도 그 버전의 옛 기록을 빼고 계산합니다.

## [2.9.45] - 2026-11-04

저장소 지침 반영(R5): 라이브러리 없음. 회귀 테스트: `tests/test_contract.py` "라이브러리 없음" (2.9.44에서 7개 중 6개 실패).

### 수정
- 진입점 최상단 import(numpy, pandas, joblib, scikit-learn)가 실패하면 traceback만 나오고 stdout JSON이 없던 문제. `ml_predict.py`는 핵심
  라이브러리를 감싸서 불러오고(`CORE_IMPORT_ERROR`, 타입 표시는 `from __future__ import annotations`), 모든 진입점이 먼저
  `check_core_libraries`, 나머지 무거운 import는 `exit_missing_library`로 감쌉니다. 응답: exit 1, `error_type: "missing_library"`,
  `missing_package`, `install_command: "python -m pip install <패키지>"`, `agent_directive` `STOP_RETRY: <기능>에 필요한 <패키지>이(가) 설치되어
  있지 않습니다. install_command를 사용자에게 전달하세요. 직접 설치하지 마세요.` 오류 로그에도 남습니다.
- 실행 중의 `ImportError`(scipy, openpyxl 등)도 `ERROR:`(내부 오류) 대신 같은 `missing_library` `STOP_RETRY`. 보고서의 python-docx·matplotlib 없음도 같은 형식.
- PART C에 `MissingLibraryError`, `missing_library_error`, `require_lib`, `exit_missing_library`, `check_core_libraries`, `PIP_NAMES`(sklearn -> scikit-learn 등).

## [2.9.44] - 2026-11-04

23차 점검 L1. 회귀 테스트: `tests/test_names.py` `test_alias_pair_with_degree_spelling_refused` (2.9.43에서 3개 모두 실패).

### 수정
- 예측 `-i`에 `온도`와 `온도(C)`를 함께 주면(모델 입력 `온도(°C)`) 같은 열의 두 값이라는 `[후보]` 대신 "'온도'은(는) 모델 입력이 아님"으로
  안내하던 문제. `check_inline_aliases`가 단위 붙은 이름도 입력 열 맞추기(`align_input_columns`)와 같은 도 기호 동등성(`_same_unit_name`:
  `°C`·`℃`·`C`, 띄어쓰기)으로 모델 입력에 맞춥니다: `[후보] 온도=180, 온도(C)=182`. `온도(℃)`, `온도( °C)`도 같습니다.

## [2.9.43] - 2026-11-04

9차 점검 CSV-4. 회귀 테스트: `tests/test_loading_numbers.py` "소수점 쉼표 판단은 열마다" (2.9.42에서 5개 중 3개 실패).

### 수정
- 탭·`;` CSV에서 글자 열의 `1,2`(라인 목록)가 파일 전체를 소수점 쉼표로 정해, 다른 열의 천 단위 숫자(`1,000`)나 점 묶음 숫자(`2.125`) 때문에
  파일 전체가 `STOP_RETRY`이던 문제 (2.9.18은 학습). 소수점 쉼표 판단을 열마다 합니다(`_comma_columns`, PART C):
  - 채운 칸의 절반 이상이 숫자가 아닌 글자 열은 근거가 아니고 바꾸지 않습니다 (`1,2`, `A`, `1,2,3` 열의 `1,2`는 글자 그대로).
  - 소수점 쉼표가 없는 열에 점 소수(`61.5`)가 있으면 파일의 소수점은 `.`: 어느 열도 소수점 쉼표로 읽지 않고 `1,000`은 1000, `2.125`는 2.125.
    `1,2`만 든 열은 글자 그대로라 학습 피처로 쓰면 기존 "소수점 쉼표로 쓴 숫자 글자" 검사가 `--dtypes "<열>:str"`를 안내합니다.
  - 점 소수가 없는 파일은 지금처럼 소수점 쉼표 열과 `1,000`·`2.125`가 함께 있으면 `STOP_RETRY` (유럽식 파일이면 1000배 다름, 2.9.22 규칙).

### 바뀐 동작 (의도한 변경)
- `;`·탭 CSV에 소수점 쉼표 열(`170,5`)과 점 소수 열(`60.25`)이 함께 있으면 2.9.42는 둘 다 소수로 읽었지만, 이제 소수점 쉼표 열은 글자로 남고
  학습 피처로 쓰면 "소수점 쉼표로 쓴 숫자 글자" `STOP_RETRY`입니다 (한 파일에 두 소수점 규칙이 섞임).
- 학습 판단(`decimal_comma`)을 따르는 예측 파일·`--points` 파일도 글자 열은 바꾸지 않습니다.

## [2.9.42] - 2026-11-04

24차 확인 검사 LOW-3. 회귀 테스트: `tests/test_map.py` "묻지 않은 값 열의 섞인 %" (2.9.41에서 3개 중 1개 실패).

### 수정
- `ml_map.py`를 `-y` 없이 실행하면 묻지 않은 값 열 하나에 `%` 칸과 `%` 없는 숫자 칸이 섞여 있어도 조회 전체가 `STOP_RETRY`이던 문제.
  `-y`가 없으면 섞인 값 열은 빼고 나머지 값 열로 계산하며 `warnings`에 "값 열 ['<열>']은(는) ... 섞여 있어 계산하지 않았습니다"를 넣습니다.
  기준 열(`-i`, `--points`), `-y`로 지정한 열, 역산의 `--find`·`--vary` 열은 지금처럼 멈춥니다. 값 열이 모두 섞였으면 멈춥니다.
- PART C에 `percent_mix`(섞인 열 목록, `check_percent_mix`가 씀).

## [2.9.41] - 2026-11-04

23차 점검 M1. 회귀 테스트: `tests/test_loading_numbers.py` ".xlsx 백분율 서식 칸" (2.9.40에서 5개 모두 실패).

### 수정
- .xlsx의 백분율 서식 칸(Excel 화면 `11.9%`, 저장 값 0.119)을 분수(0.119)로, 같은 표를 CSV로 저장한 `11.9%`는 11.9로 읽어 두 저장 형식이
  100배 다르던 문제. 공통 로더(`read_raw_table`)가 첫 시트에서 서식에 `%`가 있는 숫자 칸을 Excel이 보여 주는 글자(`11.9%`, 유효숫자 15자리)로
  바꿔 CSV와 같은 규칙으로 읽습니다: 학습(피처·타겟), 예측 입력 파일, `ml_map.py`(데이터·조회 파일), `ml_export.py`(학습 스냅숏) 모두 11.9.
  `warnings`의 "'<열>' 열의 N칸 ... %를 떼고 숫자로 읽었습니다"에 ".xlsx의 백분율 서식 칸도 화면에 보이는 숫자로 읽음"을 더했습니다.
- 한 열에 백분율 서식 칸과 일반 숫자 칸이 섞이면 CSV의 `11.9%`/`0.119`처럼 `check_percent_mix`가 `STOP_RETRY`(안내에 "열 전체를 같은 셀 서식으로").
- 따옴표·`\`로 쓴 `%`(`0.0"%"`)는 글자일 뿐이라 값 그대로. 서식 표에 백분율 서식이 없는 통합 문서는 다시 읽지 않습니다(속도 그대로).

### 바뀐 동작 (의도한 변경)
- 백분율 서식 열로 학습한 모델은 이제 11.9 배율입니다. 2.9.40 이전에 .xlsx 백분율 열로 학습한 모델은 다시 학습하십시오 (베타, 호환 없음).

## [2.9.40] - 2026-11-04

회사 PC `run_all_tests.py --gemma`(Gemma, opencode CLI 2.2.5)에서 R1(보고서)과 R2(성능예측표)가 결과 JSON 없이 끝남:
OpenCode의 명령 실행 도구가 명령을 120초에 끊고(Gemma가 SKILL.md의 시간 제한 600000을 주지 않음), 이 PC에서 성능예측표의
입력 고르기(후보마다 다시 학습)는 2~8분 걸림. Gemma는 같은 명령을 4번 반복했고 매번 처음부터 다시 학습해 끝나지 못함.
회귀 테스트: `tests/test_report_export.py` "명령 시간 제한" (2.9.39에서 시간 제한 시험 실패).

### 기능
- 긴 작업 나눠서 진행 (CLI-dev/README.md R10): `ml_export.py`의 후보 학습을 하나씩 `<버전>/export/_fits/`에 저장하고(데이터, 타겟,
  입력, 파생 수식, seed, ml-kit·scikit-learn 버전이 같을 때만 다시 씀), 다음 학습 전에 시계를 봄. 명령 시간 제한(`ML_KIT_TIME_LIMIT`,
  기본 100초, 0 = 없음)에서 여유 30초를 넘길 것 같으면 멈추고 `status: ok`, `verified: false`, `continue: true`, `progress`,
  `next_commands`(같은 명령), `agent_directive` `CONTINUE: ...`(exit 0, 오류 로그 없음). 같은 명령을 다시 실행하면 저장된 후보를 쓰고
  이어서 학습. 실행마다 새 후보를 최소 1개 학습하므로 반복하면 반드시 끝남. `ml_report.py`가 부르는 성능예측표도 같음
  (`Unfinished`는 BaseException이라 보고서의 "성능예측표 실패는 note" 처리에 삼켜지지 않음).
- 저장된 후보는 다른 `--keep`으로 다시 만들 때도 같은 입력 조합이면 다시 씀 (R2처럼 보고서 뒤 성능예측표 다시 만들기가 빨라짐).
- 손상된 저장 파일은 지우고 다시 학습 (JIT). 저장 실패는 `[WARN]`만 (다음 실행이 느려질 뿐).
- 사람이 돌리는 `run_beta_check.py`, `tools/run_win_tests.py`와 테스트는 `ML_KIT_TIME_LIMIT=0`.

### 문서
- SKILL.md: `CONTINUE:` 행동(사용자에게 묻지 말고 `next_commands`를 그대로 다시 실행, 최대 10번), 규칙 3·6에 예외 표시.
  전역 지침(`CLI-dev/opencode/AGENTS.md`) 4장 표에도 같은 줄. 채점기(`grade_opencode.py`)는 CONTINUE 뒤의 같은 명령을 반복 위반으로 세지 않음.

## [2.9.39] - 2026-11-04

학습 직후 "어떤 변수가 제일 중요해?"에 바로 답하도록 중요 입력을 결과 JSON에 넣고, 쓰이던 기능을 SKILL.md에 적음.
회귀 테스트: `tests/test_review_v2939.py` (2.9.38에서 7개 중 6개 실패, SKILL.md 옵션 검사는 감시용이라 통과).

### 기능
- `ml_train.py` 결과 JSON에 `top_features`: 타겟별 중요 입력 `[{"feature": <원래 열 이름>, "importance": <비중>}]`, 큰 순서로 최대 5개.
  최고 모델의 순열 중요도(`leaderboard.json`·Word 보고서와 같은 값, 새 계산 없음)를 모든 입력의 합이 1이 되게 나눈 비중입니다.
  글자 입력은 원래 열 하나(원-핫 이름 아님), 날짜 부분은 날짜 열 하나로 셉니다. 회귀·분류·여러 타겟·`--best` 모두.
  `preview`에 `중요 입력: 온도 0.41, 압력 0.22, ...`(타겟마다 3개)를 적고, `agent_directive`가 그것을 알리라고 합니다.
- 순열 중요도 계산이 실패했거나 모든 입력이 0이면 그 타겟은 `top_features: []`, `warnings`에 이유. 실패하면 모든 입력을 같은 값으로
  채우는 `leaderboard.json`·보고서 쪽 동작은 그대로입니다 (결과 JSON만 비움).
- `leaderboard.json`의 `feature_importances`·`top_features` 항목에 `share`(비중, 소수 넷째 자리)를 더함.

### 문서
- SKILL.md: 2단계 "알릴 것"에 `top_features`(중요 입력), 옵션 표에 `--eval "<이름> = <식>; ..."`(숫자로 읽히는 열 이름은 멈춤, 2.9.33~2.9.35),
  3단계에 버전 비교 `ml_train.py -f "<데이터>" --history`(`records`의 `version`, `mean_val_score`, `delta_score`, `tag`를 표로),
  4단계에 파일 예측 엑셀 저장 `--out "<전체경로.xlsx>"`, 분류 한 건의 `probabilities`, 결과 파일의 예측 범위·확률 열 이름. 1단계 표에 "버전 비교".
- 사용 가이드 3.1·5·6.2에 예시, README 8.1에 `top_features`.

### 테스트
- `test_review_v2939.py`: 회귀·분류·여러 타겟의 `top_features`(정렬, 비중 범위, 원래 열 이름, 글자 열 `라인`이 한 항목, 입력이 5개 이하이면 합 1),
  계산 실패·0일 때 `[]`과 이유, SKILL.md의 백틱 명령에 있는 모든 `--옵션`이 그 스크립트의 argparse에 있는지(앞으로 문서와 코드가 어긋나지 않게).

## [2.9.38] - 2026-11-04

En-Genius 스킬 배포 구조 변경에 맞춤. 회귀 테스트: `tests/test_review_v2937.py` (고침).

### 바뀐 동작
- En-Genius 스킬 zip(`docs/engenius/ml-kit.zip`)에는 지시문과 `run.py`만 들어가고, `run.py`가 `C:\CLI\ml-kit`의 도구를 실행합니다.
  그래서 `ENGENIUS_OUT`이 있어도 도구 폴더의 데이터 폴더(OpenCode와 같은 것)를 씁니다. 고정 폴더(2.9.37)는 도구 파일이 트레이 임시 폴더에서 실행될 때만.
- 도구 업데이트 때 스킬 zip을 다시 가져올 필요는 `SKILL.md` 본문(버전 줄 제외)이 바뀐 경우뿐입니다. `run.py`가 지문을 비교해 알려 줍니다.

## [2.9.37] - 2026-11-04

En-Genius(사내 에이전트) 스킬 실행 대응. 회귀 테스트: `tests/test_review_v2937.py` (2.9.36에서 4개 중 2개 실패).

### 기능
- En-Genius 트레이는 스킬 파일을 실행마다 임시 폴더에 풀고 끝나면 지웁니다. 환경변수 `ENGENIUS_OUT`이 있으면(트레이가 넣음) 데이터 폴더(오류 로그)를
  `%USERPROFILE%\Documents\En-Genius\CLI-data\ml-kit\data`에 둡니다. `ML_KIT_DATA_DIR`가 있으면 그것이 먼저입니다.
  모델(`<데이터>_ml_data`)과 보고서는 지금처럼 데이터 파일 옆에 만들어집니다.

## [2.9.36] - 2026-11-03

25차 확인 검사에서 찾은 2.9.35 회귀 수정. 회귀 테스트: `tests/test_review_v2936.py` (2.9.35에서 8개 중 6개 실패).

### 수정
- 2.9.35 회귀(높음): `ml_map.py --points`에서 조회 파일 열과 비슷한 데이터 열에 모두 `-y`의 단위를 뺀 이름이 들었을 때, 그 이름을 뺀
  나머지를 기호·괄호를 남긴 채 비교해 `온도_1`/`온도(1)`(`_1`/`(1)`), `온도_상`/`온도(상)`, `압력1`/`압력(1단)`, `수율_1`/`수율(1차)`가 비슷하지
  않다고 보았습니다. 그래서 다른 위치·단의 키 열을 말없이 빼고 결과를 `VERIFIED`로 냈습니다. 이제 나머지는 기호·괄호·`_`를 빼고 비교하고,
  한쪽이 다른 쪽을 포함해도 비슷한 열로 보아 `[후보]`로 멈춥니다. `수율_실측`/`함수율`(`실측`/`함`), `목표수율`은 계속 실측값·메모 열로 붙습니다.
- 바뀐 동작: 데이터에 없는 번호의 열(`온도1`, 데이터는 `온도2`만, `-y 온도`)과 `압력 A`/`압력 B`는 메모 열로 붙습니다 (2.9.34는 `[후보]`로 멈춤).
  다른 센서의 값이라 기준 열로 쓰지 않는 것이 맞습니다.

## [2.9.35] - 2026-11-03

24차 확인 검사에서 찾은 2.9.34 회귀 수정. 회귀 테스트: `tests/test_review_v2935.py` (2.9.34에서 23개 중 12개 실패).

### 수정
- NEW-1 (2.9.34 회귀): `ml_map.py --points`의 실측값·메모 열(`수율_실측`, `목표수율`, `-y 수율`)이 쓰지 않은 데이터 열 `함수율`과
  `수율`만 같은데도 `[후보] '함수율'`로 멈추던 문제. 조회 파일 열과 비슷한 열에 모두 `-y`의 단위를 뺀 이름이 들었으면 그 이름을 빼고
  같은 비슷한 열 검사를 다시 합니다(`_check_near_points`의 `targets`). 나머지가 비슷하거나 둘 다 비면 멈추고(`함수율%`/`함수율`: `함%`/`함`,
  `냉각수 온도(°C)`/`냉각수온도`, `-y 온도`), 아니면 그대로 붙습니다(`수율_실측`/`함수율`: `_실측`/`함`).
- NEW-2 (2.9.34 회귀): `compile_formula`의 숫자 이름 검사가 단위까지 쓴 전체 이름(`--eval "합=0.5(mm)*2(mm);차=온도+2(mm)"`)도
  `STOP_RETRY`하던 문제. 수식에서 숫자로 시작하는 전체 열 이름을 먼저 빼고 남은 숫자만 검사합니다. `온도*2`(`2(mm)` 열)는 그대로 멈춥니다.
- LOW-1: 더 긴 숫자 안의 숫자(`2(mm)` 열과 `온도*2.5`, `5(mm)` 열과 `온도*0.5`)를 그 열로 보던 문제. 숫자로 시작하는 이름은 앞뒤에
  숫자·`.`이 붙으면 이름으로 보지 않습니다(검사와 이름 바꾸기 모두, `2` 열과 `온도*2.5`도 숫자 2.5).
- LOW-2: 학습의 섞인 `%` 검사(`check_percent_mix`)가 `제외` 열에 표시한 행에도 멈추던 문제. 표시한 행(`포함` 제외)을 뺀 행에서 검사합니다.
  예측은 그대로.

### 바뀐 기대값 (의도한 변경)
- `test_review_v2934.py`의 `test_number_alias_in_formula_stops`에서 `온도+2.5`(`2(mm)` 열): 2.9.35는 숫자 2.5로 컴파일(LOW-1).
- `test_review_v2933.py`의 `test_number_named_column_in_formula_stops`에서 `온도*3 + 2.5`(`2` 열): 같은 이유로 숫자 2.5로 컴파일(2.9.33/2.9.34는 `STOP_RETRY`).

### 테스트
- `test_review_v2935.py`: 실측값 열 5쌍 통과(`-y 수율`, `-y 수율,생산량`, 데이터에 `함수율`), 비슷한 열 5개 `[후보] '함수율'`(`함수율%`,
  `함수 율(%)`, `함수율_수율`, `함수율 수율`, `수율실측함수율`), `냉각수 온도(°C)` 멈춤, 전체 숫자 이름 수식과 학습 `--eval`, 숫자 안의 숫자,
  `온도*2`·`온도*0.5` 멈춤, `제외` 행의 `%` 칸(학습 통과)과 `포함` 행의 `%` 칸(멈춤).

## [2.9.34] - 2026-11-03

23차 점검에서 찾은 오류 수정. 회귀 테스트: `tests/test_review_v2934.py` (2.9.33에서 21개 중 15개 실패).

### 수정
- R1 (2.9.33 회귀): `ml_map.py --points`의 실측값·메모 열 통과(`_check_near_points`)가 이름에 `-y` 타겟 이름이 들었는지만 보아(`t in name`),
  `-y` 이름이 든 진짜 기준 열(`냉각수 온도(°C)`와 `-y 온도`, 데이터는 `냉각수온도`; `함수율%`와 `-y 수율`)을 말없이 빼고 틀린 결과를
  `VERIFIED`로 내던 문제. 모든 열에 비슷한 열 검사(`suggest_columns(..., bare=False)`, 기준으로 쓰지 않은 데이터 열)를 먼저 하고, 비슷한 열이
  있으면 2.9.32처럼 `[후보]`로 멈춥니다. 비슷한 열이 없을 때만 실측값·메모 열로 붙습니다(`수율_실측`, `목표수율`, `생산량_실측`은 그대로 통과).
- R2 (2.9.33 회귀): `ml_map.py`의 `-i "함수율=10%"`와 `--points`의 `12%` 값이 "숫자가 아닙니다"로 멈추던 문제(`_split_keys`가 `float(v)`).
  `coerce_numeric`과 같은 `%` 떼기(`_strip_percent`)로 읽고(10, 0.1이 아님), `warnings`에 알립니다(`-i '<이름>'의 값 '10%'은(는) %를 떼고 ...`,
  `조회 파일 '<열>' 열의 N칸 ...`).
- H1: 한 숫자 열에 `%` 칸과 `%` 없는 숫자 칸이 섞이면(`11.9%`와 `0.119`) 말없이 학습하던 문제. 새 `check_percent_mix`(PART C)가 `STOP_RETRY`:
  열 이름, 두 형식의 예 하나씩, 열을 모두 `%` 또는 모두 `%` 없는 숫자(한 가지 배율)로 고치라는 안내. 학습(숫자 피처, `--eval` 원본 열, 회귀 타겟),
  예측(숫자 입력, 추론 전), `ml_map.py`(기준·구할 열, `--find` 목표, `--points` 기준 열, 계산 전). 모든 칸에 `%`가 있는 열은 2.9.33과 같습니다.
- H2: `compile_formula`의 숫자 이름 검사가 전체 열 이름만 보아 `2(mm)`의 단위 없는 이름 `2`가 수식 `온도*2`의 숫자 2를 말없이 그 열로 바꾸던 문제.
  숫자로 읽히는 단위 없는 이름은 수식 이름으로 쓰지 않고, 그 열의 전체 이름이나 단위 없는 이름이 수식에 숫자로 있으면 `STOP_RETRY`.

### 바뀐 기대값 (의도한 변경)
- `test_layout_matrix_v2917.py`의 `L12_pct_text`(타겟 `수율`의 한 칸이 `12%`, 나머지는 숫자): 학습(`test_train_with_correct_options`,
  `test_train_without_options`)은 섞인 `%`의 `STOP_RETRY`(2.9.33은 12로 읽고 학습). `[추천]`이 없는 `STOP_RETRY`는 따라 실행하지 않음(`_run_with_one_retry`).

### 테스트
- `test_review_v2934.py`: `냉각수 온도(°C)`/`-y 온도`, `함수율%`/`-y 수율`과 `-y 수율(최종)`의 `[후보]`, 실측값 열 4쌍 통과(결과 파일에 그 열),
  `_check_near_points` 양방향, `-i 함수율=10%`(순방향·`--find`)와 `--points 12%`가 `%` 없는 데이터와 같은 값과 경고, `check_percent_mix`,
  학습 피처·타겟과 예측 입력 파일과 `ml_map.py` 데이터·조회 파일의 섞인 `%` `STOP_RETRY`(모두 `%`인 입력은 통과, 버전 폴더 안 남음),
  `2(mm)`/`0.5(mm)`의 단위 없는 이름과 수식 숫자(`온도*2`, `온도*0.5`, `온도+2.5`), 학습 `--eval`.
- 이전 점검 입력 폴더(r8~r23, 3986개 파일: .csv/.xlsx/.xls/.txt/.tsv, 결과 파일 포함) 불러오기 비교: 열, 행 수, 숫자 열 판단이 바뀐 파일 없음.
  섞인 `%` 열이 있는 파일 31개(그 열을 쓰면 새로 `STOP_RETRY`): r8/r9/r14 `L12_pct_text.xlsx` 18개(같은 파일의 사본, `수율` 열의 `12%` 칸 하나와
  `73.39` 같은 숫자 칸; 2.9.33은 12로 읽고 학습), 그 예측 결과 CSV 12개(같은 열을 그대로 담음), r23 `pctmix.csv`(`함수율`).

## [2.9.33] - 2026-11-03

22차 점검에서 찾은 오류 수정. 회귀 테스트: `tests/test_review_v2933.py` (2.9.32에서 38개 중 33개 실패).

### 설계 (이름 맞추기의 `[추천]`을 없앰)
- 2.9.28~2.9.32의 다섯 차례 오타 판단(`sure_typo`) 수정은 매번 새 예외를 만들었습니다. 마지막은 끝 포함 규칙으로 `속도` -> `[추천] '가속도'`,
  `수율` -> `[추천] '함수율(%)'`: 에이전트는 `[추천]`을 스스로 다시 실행하므로 값이 다른 양으로 말없이 들어갔습니다.
- 새 규칙: 철자가 비슷한 열·피처·타겟 이름(`difflib`)은 후보가 하나여도 `[추천]`이 아닌 `[후보]`(사용자에게 물음)이고, 메시지에 "비슷한 이름은 다른
  양일 수 있으니 어느 열인지 사용자에게 물으십시오"를 붙입니다. `resolve_column_names`(학습 `-x`/`-y`, 예측 `-y`/`--find`, `ml_map.py`의 모든 이름과
  `--points`, `ml_export.py`, `ml_report.py --target`), 예측 `-i`의 모르는 이름(`check_inline_unknown`), 빠진 입력 열(`_typo_hints`), `--eval` 수식
  이름(`_validate_eval_ast`)이 모두 같습니다.
- 같은 열로 확실히 정해지는 이름은 묻지 않고 그대로 맞춥니다(`[INFO] 컬럼명 자동 매칭`): 정확한 이름, 단위 없이 쓴 고유한 이름, 단위를 뺀 이름
  (대소문자·띄어쓰기 무시)과 단위(띄어쓰기, 괄호 모양 `[bar]`/`(bar)`, `°C`·`℃`·`C`)가 같은 이름. 새로 맞추는 것은 도 기호 표기(`온도(C)`, `온도(℃)`
  -> `온도(°C)`, 이전에는 `[추천]`)뿐이고, 그런 열이 둘이면 `[후보]`. 단위의 대소문자(`mg`/`Mg`)는 다른 단위. 머리글 행 찾기(`names_in_cells`)도 같음.
- 지운 것: `sure_typo`, `one_or_ask`, `_osa`, `_common_syllable`(자모, KS X 1001 완성형, 짧은 영문 규칙). `_same_unit_name`이 도 기호 표기까지 맡습니다.
- 그대로: 다른 옵션이 쓰는 열(`exclude=`)은 후보가 아님, 단위가 다른 이름 하나는 값 환산 안내(`unit_hint`), 주어진 이름마다 묶은 `[후보]`.
- 이름이 아닌 `[추천]`을 하나씩 확인: 다시 실행해 확인한 머리글·단위·데이터 행 명령, 쉼표를 뺀 `-i` 명령, `--version <다음>`, `-o <이름>.csv`,
  8개를 넘는 `--inputs`의 줄인 안(다른 타겟의 `--inputs` 유지)은 그대로 `[추천]`. 아래 둘은 사용자가 뜻한 것과 달라 `[후보]`로 바꿈:
  `ml_export.py --inputs`의 입력이 모두 값이 하나뿐일 때(L1), `ml_train.py --best --task`에 이력이 없는 학습(다른 학습의 최고 버전이 됨).

### 수정
- R1 (2.9.32 회귀): `ml_map.py --points`에서 `-y` 타겟과 단위를 뺀 이름이 같은 열이 있으면(`-y 수율(최종)` 옆 `수율(1차)`, `-y 생산량`도) 실측값·메모 열
  `수율 실측`, `수율_실측`, `목표수율`, `수율목표`를 단위를 뺀 이름 비교로 `수율(1차)`와 "비슷한" 열로 보고 멈추던 문제. `_check_near_points`는
  `-y` 타겟의 단위를 뺀 이름이 든 열을 실측값·메모 열로 통과시키고(2.9.31처럼), 비슷한 열 판단에는 단위를 뺀 이름 비교를 쓰지 않습니다
  (`suggest_columns(..., bare=False)`; 그 비교는 `[후보]` 목록을 만들 때만).
- R2 (2.9.32 회귀): `-y` 없는 `--points`의 "-y가 없어 ..." 안내가 숫자가 아닌 기준 값, 같은 열을 가리키는 두 이름 같은 다른 오류에도 붙던 문제.
  `resolve_column_names`의 이름 오류(새 `ColumnNameError`, `InputError`의 하위 클래스)에만 붙입니다(`_ask_without_y`).
- H1: 위 설계. `속도`/`가속도`, `수율`/`함수율(%)`은 `[후보]`.
- H2: 숫자로 읽히는 이름의 열(`1`, `2`)이 있으면 `--eval` 수식의 숫자(`온도*2`, `2.5`)가 그 열로 바뀌어 말없이 다른 값을 계산하던 문제(`compile_formula`).
  그 숫자가 수식에 있으면 `STOP_RETRY`(수식에는 열 이름을 감싸 쓰는 방법이 없어 열 이름을 바꾸라는 안내).
- M1: `%`를 붙여 쓴 숫자(`64.4%`, Excel 백분율 칸을 CSV로 저장한 값)를 숫자로 읽지 못해 피처는 "ID/일련번호로 보입니다"나 범주, 타겟은 "숫자로 변환
  가능한 값이 하나도 없습니다 ... --task classification"(틀린 안내)이던 문제. `coerce_numeric`(PART C, 학습·예측·표 보간 공통)이 끝의 `%`와 앞뒤 공백을
  떼고 쓰인 그대로 읽습니다(64.4, 0.644가 아님). 열마다 `warnings`에 "'<열>' 열의 N칸('64.4%')은 %를 떼고 숫자로 읽었습니다 (64.4% -> 64.4, 100으로
  나누지 않음)"(`percent_cells`, `percent_warning`): 학습(숫자 피처, `--eval` 원본 열, 회귀 타겟), 예측(숫자 입력, `warnings`), `ml_map.py`(쓴 열).
  피처 종류·ID 판단 전에 숫자로 읽히고, 단위 행 판단(`_numeric_share`, `_number`)과 합계 행 값 비교(`_number_matrix`)도 같게 읽습니다.
  읽을 수 없는 숫자 경고(`unparsed_number_cells`)에는 `%` 숫자가 더 이상 들지 않습니다.
- L1: `ml_export.py --inputs`의 입력이 모두 값이 하나뿐이면 `[추천] <다른 입력 8개>`가 사용자의 선택과 다른 타겟의 `--inputs`를 바꾸던 문제.
  `[후보] <값이 변하는 입력>`과 사용자에게 묻는 안내.
- L2: 다른 옵션이 쓰는 열이 `[후보]`에 다시 들어가던 문제(같은 이름의 열이면 `s = same`으로 되살림). `exclude=`의 열은 `[후보]`에 넣지 않고,
  안내 문구를 "이 명령에서 이미 다른 이름이나 옵션에 쓰인 열"로 고침(`ml_export.py --keep "압력,압룍"`처럼 같은 옵션의 다른 이름도 맞게).

### 바꾼 기대값 (의도한 변경)
- `test_cli_smoke.py::test_train_column_typo_suggests`: `반응온도(C)`는 이제 맞으므로 오타 `반응온도x`로 바꾸고 `[추천]` -> `[후보]`.
- `test_export_v290.py`: `--keep 온도x`는 `[추천]` -> `[후보] '온도'`.
- `test_layout_matrix_v2918.py::test_model_names_with_units_match_header_without_units`: `names_in_cells(["온도(C)"], ["온도(°C)"])`는
  이제 참(같은 단위), 다른 단위 검사는 `온도(K)`.
- `test_review_v296.py::test_inputs_and_keep_drop_constant_with_hint`: `[추천] a` -> `[후보] a` (L1).
- `test_review_v2913.py::test_inline_typo_points_at_the_user_key`: `도핑농돜` -> `[후보]`.
- `test_review_v2914.py::test_header_hint_not_the_current_row`: `온도(C)`는 맞으므로 오타 `온도x`로 바꾸고 `[추천]` -> `[후보]` (`ml_train.py`, `ml_map.py -i`);
  `test_find_header_row_rules`: `온도(C)`는 `온도(°C)` 머리글 행을 찾음(이전에는 None), 다른 단위 검사는 `온도(K)`.
- `test_review_v2929.py::test_same_unit_typo_still_suggested`: 예측 `-i 압력(bar)`(모델 `반응압력(bar)`)과 `ml_map.py -i 반응압룍(bar)`은 `[후보]`.
- `test_review_v2930.py`: `test_points_near_name_refused`(`반응압룍(bar)`)와 `test_inline_typo_key_suggested`(`속돜`)는 `[후보]`.
- `test_review_v2931.py`: `test_points_other_word_is_no_recommendation`(`속돜`), `test_inline_keys_never_suggest_the_value_column`(`압룍(bar)`),
  `test_other_word_is_a_candidate`(`sure_typo` 검사 뺌), `test_typo_still_recommended`(6쌍 모두 `[후보]`; 맞는 `반응 온도`, `반응온도(C)`는
  `test_typo_resolve_messages`로), `test_train_and_predict_other_word`(`습돜`), `test_one_name_for_two_inputs`(`속돜`), `test_eval_unknown_name_wording`
  (`온돜`), `test_spaced_unit_is_exact`(`온도(C)`는 맞음).
- `test_review_v2932.py`: `test_inline_unitless_name_next_to_the_target_asks`와 `test_train_x_typo_never_suggests_the_target`(`[후보]`에 `-y` 열
  `수율(최종)`이 없음, L2), `test_not_a_sure_typo`/`test_sure_typo`(`sure_typo` 검사 뺌, `SURE` 8쌍은 `[후보]`, `반응온도(C)`는 맞으므로 뺌),
  `test_common_syllable_typo_cli_is_a_candidate`(`습돜`), `test_eval_never_suggests_a_target`(`온돜`), `test_export_keep_typo_never_suggests_a_given_input`
  (`압룍`), `test_unitless_name_against_unit_columns`(`온돜`).

### 테스트
- `test_review_v2933.py`: `sure_typo` 없음, 철자가 비슷한 9쌍 `[후보]`(`속도`/`가속도`, `수율`/`함수율(%)`, `속두`, `압룍`, `온도x`, `tmep`, `압력`/`반응압력(bar)`,
  `생산량(kg)`/`생산량`, `온돜`), 같은 열 7쌍 자동 맞춤(`온도(C)`, `온도(℃)`, `TEMP(°C)`/`temp(℃)`, `압력 [bar]`, `수율( 1차)`, `반응 온도(C)`, `생산효율`),
  단위 대소문자·같은 열 둘은 `[후보]`, 학습 `-x 속도`/`-y 수율`(데이터 `가속도`, `함수율(%)`), 학습 `-x 온도(C)`, `ml_map.py -i 압룍`, `--points` 실측값 열
  4종 x `-y` 2종 통과와 가까운 이름 `솎도`의 `[후보]`, `-y` 없음 안내(숫자 아님·같은 열 둘에는 없음, 이름 오류에는 있음), 숫자 이름 열과 수식,
  `%` 숫자(`coerce_numeric`, 학습 피처·타겟 경고, `%` 없는 같은 데이터와 같은 점수, 예측 경고와 같은 예측값, `ml_map.py` 경고),
  `ml_export.py --inputs` 상수 입력 `[후보]`, `exclude=` 열이 `[후보]`에 없음과 `--keep` 안내, `--best --task` `[후보]`.
- 이름 쌍 비교(새로 만든 96 x 74, 7030쌍의 `resolve_column_names`; 제조 열 이름 74개 + 오타·다른 표기 22개): 2.9.32 `[추천]` 24, `[후보]` 282,
  자동 맞춤 6, 유사 없음 6718 -> 2.9.33 `[추천]` 0, `[후보]` 303, 자동 맞춤 9, 유사 없음 6718. `[추천]` -> `[후보]` 21쌍(`속도`/`가속도`, `속도`/`사출속도`,
  `함량`/`수분함량`, `효율`/`생산효율(%)`, `강도`/`인장강도`, `전력`/`소비전력`, `압룍`/`압력(bar)`, `온도x`/`온도(°C)`, `tmep`/`Temp`, `생산량(kg)`/`생산량` 등),
  `[추천]` -> 자동 맞춤 3쌍(`온도(C)`, `온도(℃)` -> `온도(°C)`, `반응온도(C)` -> `반응온도(°C)`). 이름 맞추기의 `[추천]`은 이제 없습니다(같은 열은 맞추고,
  나머지는 묻기).
- 이전 점검 입력 폴더(r8~r22, 1407개 파일) 불러오기 비교: 열, 행 수, 합계 행 결과가 바뀐 파일 없음. `%` 숫자를 읽게 된 파일 21개:
  r22 `pct.csv`(`습도(%)`), `pct2.csv`(`수분(%)`), `pcty.csv`(`수율(최종)`)는 숫자 열이 됨(이전에는 글자), r8/r9/r14 `L12_pct_text.xlsx` 18개(같은
  파일의 사본)는 `수율` 열의 `12%` 칸 하나를 12로 읽음(이전에는 빈 값).

## [2.9.32] - 2026-11-03

21차 점검에서 찾은 오류 수정. 회귀 테스트: `tests/test_review_v2932.py` (2.9.31에서 42개 중 32개 실패).
원칙: 에이전트는 `[추천]`을 스스로 1번 다시 실행하므로, 다른 양을 가리키거나 다시 실행하면 실패하는 `[추천]`은 오류입니다. 확실하지 않으면 `[후보]`.

### 수정
- 2.9.31 회귀: 다른 옵션이 쓰는 열을 빼고(`exclude=`) 찾아 단위 없는 이름이 말없이 하나로 맞던 문제. 데이터 열 `수율(1차)`, `수율(최종)`과
  `-y 수율(최종)`에서 `ml_map.py --points`의 `수율` 열이 `수율(1차)` 기준 열이 되어 `VERIFIED`를 내고, `-i 수율=40`은 `[추천] '수율(1차)'`
  (2.9.30은 둘 다 `[후보]`). `--points`는 정확한 이름과 같은 단위의 띄어쓰기 차이만 먼저 기준 열로 보고, 구할 열의 이름(단위를 뺀 이름)인 열은
  실측값 열입니다(`_points_keys`). `resolve_column_names`는 뺀 열이 단위를 뺀 이름이 같거나 다른 후보만큼 비슷하면 `[추천]` 없이 `[후보]`(같은
  이름의 열 모두)와 "다른 옵션에 쓰인 열" 안내를 줍니다.
- `[추천]`의 오타 판단(`sure_typo`, 점검자 제안 B = 제안 C + 흔한 글자 규칙): 다음은 `[후보]`입니다.
  - 흔한 한글 글자 하나를 다른 글자로 바꾼 이름: 주어진 글자가 KS X 1001 완성형 2350자(EUC-KR 2바이트)이면 다른 낱말일 수 있음
    (`속두`/`속도`, `습두`/`습도`, `온두`/`온도`, `강도`/`경도`, `점도`/`정도`, `생산랑`/`생산량`). 드문 글자(`압룍`/`압력`)는 `[추천]`.
  - 3글자 이하 영문 이름은 단위·대소문자·띄어쓰기만 다를 때만 맞음(`O2`/`T2`, `T1`/`X1`, `RMP`/`RPM`).
  - 포함은 주어진 이름이 열 이름의 끝일 때만(`압력`/`반응압력`): 열 이름의 앞부분(`온도`/`온도차`, `Ti`/`Time`, `Press`/`Pressure`)과 주어진 이름이 더
    긴 경우(`설정온도`/`온도`, `소비전력`/`전력`: 데이터에 없는 더 좁은 양일 수 있음)는 `[후보]`.
  - 같은 길이에서 영문 글자 하나만 다름(`TempA`/`TempB`, `온도A`/`온도B`, `Heat`/`Head`, `Spped`/`Speed`).
  - 두 이름에 모두 괄호가 있고 단위·구분이 다름(`수율(1차)`/`수율(최종)`, `생산량(kg)`/`생산량(t)`, `압력(bar)`/`압력(MPa)`; `°C`/`C`는 같은 단위).
  `[추천]`은 단위를 뺀 이름이 같을 때, 끝 포함, 글자 넣기·빼기(`온도x`), 자리바꿈(`tmep`), 드문 글자로 바뀐 자모 하나(`압룍`)만 남습니다.
- 다른 옵션이 쓰는 열은 오타 후보가 아님 (한 명령 안에서 다시 실행하면 실패하거나 다른 열을 쓰게 되는 `[추천]`):
  - `ml_train.py`: `-y`를 먼저 맞추고 `-x`의 오타 후보에서 타겟을, `-y`의 오타 후보에서 `-x` 열을 뺌(`-x 생산랑 -y 생산량`의 `[추천] '생산량'`은
    다시 실행하면 데이터 누수 오류).
  - `ml_predict.py --find`: `-i`의 오타 후보에서 `--vary` 입력을, `--vary`의 오타 후보에서 `-i` 입력을 뺌.
  - `ml_map.py --find`: `--vary`의 오타 후보에서 목표 열과 `-i` 열을, 목표의 오타 후보에서 `--vary`·`-i` 열을 뺌. `ml_map.py`의 `-y` 오타 후보에서 `-i` 열을 뺌.
  - `--eval`: 수식의 모르는 이름은 타겟을 후보로 주지 않음(`_validate_eval_ast`에 타겟 전달). 타겟만 비슷하면 "타겟이라 수식에 쓸 수 없음".
  - 추가로 찾은 곳: `ml_map.py --points`의 비슷한 이름 열 검사(`_check_near_points`)는 기준 열·구할 열이 다른 후보만큼 비슷하면 `[후보]`,
    `ml_export.py --keep`/`--inputs`(`_check_known`)는 한 이름씩 맞추므로 다른 이름이 이미 맞춘 입력을 후보에서 뺌, 예측 `-i`의 모르는 이름
    (`check_inline_unknown`)과 입력 열 오타(`_typo_hints`)는 이미 준 입력도 비슷하면 `[후보]`.
  - 새 함수 `known_columns(names, columns)`: 오류 없이 맞는 열(다른 옵션의 `exclude=`로 전달).
- `ml_train.py --positive-class`: 클래스에 없는 라벨이면 `[추천] ['불량', '양품']`(파이썬 목록, 다시 실행할 값이 없음)이던 문제. `[후보] 불량, 양품`
  (고르는 것은 사용자 몫).
- `ml_map.py --points`(`-y` 없음): `-y` 안내(기준 열이면 머리글 고치기, 실측값이면 `-y`)를 `[추천]`이 `[후보]`로 바뀐 경우에만 붙이던 문제. 처음부터
  `[후보]`인 경우에도 붙입니다(`_ask_without_y`).
- 단위 없는 이름은 단위를 뺀 열 이름과도 비교(`suggest_columns`, `_typo_hints`): `온두`/`온도(°C)`가 "유사 컬럼 없음"이던 문제. 위 규칙대로 `[후보]`.

### 바꾼 기대값 (의도한 변경, 모두 `[추천]` -> `[후보]`)
- `test_features_v21.py::test_positive_class_validation`: `--positive-class` 없는 라벨은 `[후보]`.
- `test_review_v2913.py::test_inline_typo_points_at_the_user_key`: `도핑농두` -> `[후보]`, `[추천]` 검사는 드문 글자 `도핑농돜`로.
- `test_review_v2930.py::test_inline_typo_key_suggested`: `속두` -> `[후보]`, `[추천]` 검사는 `속돜`로.
- `test_review_v2931.py`: `test_points_other_word_is_no_recommendation`(`속두`, `[추천]`은 `속돜`), `test_typo_still_recommended`의 `속두`,
  `도핑농두`(-> `속돜`, `도핑농돜`)와 `생산량(kg)`/`생산량(t)`(-> `생산량(kg)`/`생산량`, 다른 단위는 `[후보]`), `test_typo_resolve_messages`(`속두` -> `속돜`),
  `test_train_and_predict_other_word`(`습두` -> `[후보]`, `[추천]`은 `습돜`), `test_one_name_for_two_inputs`(`속두` -> `[후보]`, `[추천]`은 `속돜`),
  `test_eval_unknown_name_wording`(`온두` -> `[후보]`, `[추천]`은 `온돜`).

### 테스트
- `test_review_v2932.py`: `--points`의 `수율` 열(실측값 열, 기준은 `온도`, `압력`), `-i 수율`의 `[후보]` 둘, `sure_typo` `[후보]` 20쌍과 `[추천]` 9쌍,
  예측 `-i`의 `습두`/`습돜`, 학습 `-x`/`-y`, `--find`의 `-i`/`--vary`, `ml_map.py`의 `--find`/`--vary`/`-y`/`-i`, `--eval` 타겟, `ml_export._check_known`,
  `known_columns`, `--positive-class` `[후보]`, `--points` `-y` 없음 `[후보]`의 `-y` 안내, `온두`/`온도(°C)`.
- 이름 쌍 비교(2.9.31과 같은 96 x 74, 7030쌍의 `resolve_column_names`): `[추천]` 99 -> 45, `[후보]` 187 -> 322, 유사 없음 6726 -> 6645.
  `[추천]` -> `[후보]` 58쌍(주어진 이름이 더 긺 17: `설정온도`/`온도`, `반응시간`/`시간`; 괄호 단위·구분 다름 16: `수율(1차)`/`수율(최종)`,
  `압력(bar)`/`압력(MPa)`; 흔한 글자 13: `속두`, `강도`/`경도`, `생산랑`; 짧은 영문 9: `T1`/`X1`, `RMP`/`RPM`; 앞부분 2: `Press`/`Pressure`; 영문 글자 1:
  `Spped`), 유사 없음 -> `[후보]` 77쌍(단위 뺀 비교: `온두`/`온도(°C)`, `습도`/`온도(°C)`), 유사 없음 -> `[추천]` 4쌍(단위 뺀 비교로 찾은 확실한 오타:
  `압룍`/`압력(bar)`, `압룍`/`압력(MPa)`, `온도x`/`온도(°C)`, `압력`/`반응압력(bar)`). `[후보]` -> `[추천]`은 없음(열 하나, `exclude=` 없는 비교).

## [2.9.31] - 2026-11-03

20차 점검에서 찾은 오류 수정. 회귀 테스트: `tests/test_review_v2931.py` (2.9.30에서 31개 중 28개 실패, 새 함수 `sure_typo` 검사 8개 포함).

### 수정
- `ml_map.py --points` (2.9.30 회귀): 비슷한 이름 검사가 `-y` 구할 열도 후보로 봐서, 실측값 열 `생산량_실측`에 `[추천] '생산량'`을 주고 그대로 다시
  실행하면 실패하던 문제. 구할 열(`-y`로 맞춘 열)은 기준 열 후보와 오타 후보에서 모두 뺍니다. 구할 열의 이름인 열(`생산량_실측`, `생산량(실측)`,
  `생산량`)은 실측값·메모 열로 결과에 그대로 붙습니다(2.9.29처럼). 이전부터: `생산량(실측)`을 단위가 `실측`인 기준 열로 읽어 "기준(-i)" `[추천]`을
  주던 문제도 같음. `-y`가 없으면 기준으로 쓰지 않은 숫자 열이 모두 구할 열이라 `[추천]` 대신 `[후보]`(기준 열이면 머리글을 고치고, 실측값이면 `-y`를
  붙여 다시 실행). 조회 파일 열의 오류는 "조회 파일(--points) 기준"으로 적습니다.
  같은 이유로 `ml_map.py -i`의 오타 후보에서 `-y` 열과 `--find`의 목표·`--vary` 열을 뺍니다(`-i "생산랑=470" -y 생산량`의 `[추천] '생산량'`은 다시 실행하면
  구할 열이 없어 실패).
- `[추천]`이 다른 낱말·숫자 사이에도 나오던 문제(`습도`->`속도`, `온도`->`농도`, `T1`->`T2`, `온도1`->`온도2`, `반응온도`->`반응농도`, `Temp`->`Time`).
  에이전트는 `[추천]`을 스스로 다시 실행하므로 다른 열의 값이 말없이 쓰였습니다. 새 `sure_typo(name, cand)`(단위를 뺀 이름으로 비교): 같은 이름,
  숫자가 다르면 아님, 짧은 쪽이 2글자 이상이고 한 이름이 다른 이름 안에 있음, 아니면 한글 자모(NFD) 기준 최적 문자열 정렬 거리 1 이하.
  3글자 이하 영문 이름은 길이도 같아야 합니다(`CO2`/`O2`, `pH`/`H`, `Ax`/`A`는 `[후보]`). 아니면 후보 하나도 `[후보]`(사용자에게 물음).
  `resolve_column_names`(열·피처·타겟 이름, `ml_map.py`, `ml_train.py`/`ml_report.py`/`ml_export.py` 이름 맞추기), 예측 `-i`(`check_inline_unknown`),
  입력 열 오타(`_typo_hints`), `--eval` 수식 이름(`_validate_eval_ast`)이 모두 같습니다(`difflib`를 쓰는 곳은 이 넷뿐).
  남은 위험: 자모 하나 차이인 실제 낱말(`강도`/`경도`, `점도`/`정도`)과 한 글자만 다른 2글자 영문 이름(`O2`/`T2`, `T1`/`X1`)은 여전히 `[추천]`.
- 입력 열 오타(`_typo_hints`): 주어진 이름 하나(`솔도`)가 빠진 모델 입력 둘(`속도`, `습도`)과 비슷하면 서로 다른 `[추천]` 둘을 주던 문제.
  이름마다 묶어 2개 이상이면 `[후보] 속도, 습도` 하나.
- `--eval` 수식의 모르는 이름: `[추천] ['수율(1차)', '수율(최종)']`처럼 파이썬 목록으로 적던 문제. 오타인 후보 하나는 `[추천] <이름>`, 아니면
  `[후보] a, b`.
- `ml_report.py --target "수율( 1차)"`가 `[후보]`(`-y`는 맞춤)이던 문제. 같은 단위의 띄어쓰기 차이 맞추기를 `resolve_target_models`에서
  `resolve_column_names`로 옮겨, 단위를 뺀 이름이 같고 띄어쓰기를 뺀 단위가 같으면 정확한 이름입니다(`-y`, `--target`, `-i`, `-x`, `--points`,
  예측 `-i`와 입력 열(`align_input_columns`) 모두. 괄호 모양 `[bar]`/`(bar)`와 이름 안 띄어쓰기 `반응 압력(bar)`도 같은 이름).
  단위 표기 차이(`온도(℃)`, `온도(C)` / `온도(°C)`)는 자동으로 맞추지 않고 `[추천]`(2.9.30의 예측 `-y`/`--find`는 맞췄음, 다시 실행하면 같음).
- `ml_export.py`: 단위를 뺀 이름이 같은 두 타겟(`수율(1차)`, `수율(최종)`)을 머리글(`예측 수율` 두 번), 시트 이름(`표_수율`, `표_수율_2`), 일치도 문장
  (`- 수율:` 두 번), 계산기 결과 이름에서 구별할 수 없던 문제. 그런 타겟은 전체 이름을 씁니다(`target_labels`).

### 합계 행 (바꾸지 않음)
- 2.9.30 규칙 그대로. 알려진 거짓 경고: 작은 정수 열의 값 하나가 우연히 위 두 행의 합(`불량수` `2` = `1` + `1`)이고 나머지가 일정한 설정값이면 범주 행
  `A계`도 경고합니다. 합계 행을 놓치면 그 행이 학습에 들어가 더 나빠, 경고 하나가 더 나오는 쪽을 택했습니다. `test_review_v2931.py`가 이 동작을 고정.

### 테스트
- `test_review_v2931.py`: `--points` 구할 열 이름 열(`-y` 있음: 통과와 결과 시트 유지, `-y` 없음: `[후보]`), 다른 숫자 열 `[후보]`와 오타 `[추천]`의 다시
  실행 성공, 다른 낱말·숫자·짧은 영문 9쌍 `[후보]`, 오타 8쌍 `sure_typo`(`속두`, `압력(bar)`/`반응압력(bar)`, 띄어쓰기, `반응온도(C)`, `온도x`, `도핑농두`,
  `tmep`), 학습 `-x` 자모 오타 `[추천]`, 예측 `-i` 다른 낱말, `ml_map.py -i`의 `-y`·`--find` 열 후보 제외, `솔도` 하나의 `[후보]`, `--eval` 문구 3종, 합계 행 거짓 경고 고정, `수율( 1차)`
  (`load_version`, 예측 `-y`, `-x`), 성능예측표 두 타겟 머리글·일치도 문장·시트 이름.
- 의도적으로 바꾼 기대값: `test_map_v260.py::test_errors_and_contract`의 `-i "Ax=10"`(열 `A`)은 `[추천]` -> `[후보] 'A'`(짧은 영문 이름의 길이 차이).
  `test_review_v2930.py::test_points_near_name_refused`의 `반응 압력(bar)`과 `test_review_v2929.py::test_same_unit_typo_still_suggested`의
  `반응압력( bar)`, `test_review_v2928.py::test_other_unit_no_rename_suggestion`의 예측 `-i "압력(bar )=3"`은 `[추천]` -> 바로 맞춤(같은 단위의 띄어쓰기 차이). 두 테스트의 `[추천]` 검사는 오타 `반응압룍(bar)`로 바꿈.
- 이전 점검 입력 폴더(r8~r20, 1681개 파일) 불러오기 비교: 열, 행 수, 합계 행 결과가 바뀐 파일 없음.
- 이름 쌍 비교(주는 이름 96개(제조 열 이름 74개 + 오타 22개) x 열 이름 74개, 7030쌍의 `resolve_column_names`): `[추천]` 289 -> 99. 바뀐 190쌍: `[추천]` -> `[후보]` 187쌍(다른 낱말·숫자:
  `온도`/`농도`, `반응시간`/`냉각시간`, `전압`/`전력`, `유량`/`유속`, `T1`/`T2`, `Load`/`Lot` 등, 오타가 아닌 먼 후보: `속두`/`유속`, `온도x`/`반응온도`)과
  `[추천]` -> 바로 맞춤 3쌍(`반응압력 (bar)`, `수율( 1차)`, `압력 [bar]`). `[후보]` -> `[추천]`은 이 비교(열 하나, `exclude=` 없음)에는 없었음.
  (2.9.32 정정: 다른 옵션이 쓰는 열을 빼는 곳에는 있었음. `-y 수율(최종)` 옆 `ml_map.py -i 수율`은 2.9.30 `[후보]` -> 2.9.31 `[추천] '수율(1차)'`,
  `--points`의 `수율` 열은 2.9.30 `[후보]` -> 2.9.31 경고 없이 `수율(1차)` 기준 열. 2.9.32에서 고침.)

## [2.9.30] - 2026-11-03

19차 점검에서 찾은 오류 수정. 회귀 테스트: `tests/test_review_v2930.py` (2.9.29에서 15개 중 11개 실패).

### 수정
- `ml_predict.py --find "수율=65"`(타겟 `수율(1차)`, `수율(최종)`): 단위를 뺀 이름으로 둘 다 맞아 그중 하나를 경고 없이 풀던 문제. 타겟 이름은 열 이름처럼
  맞춥니다(`resolve_column_names`: 정확한 이름, 아니면 단위를 뺀 이름이 하나만 같을 때). 여럿이면 `ml_map.py --find`처럼 `STOP_RETRY` + `[후보]`.
  예측 `-y`와 `ml_report.py --target`(최고 버전 보고서 포함)도 같아, `-y 수율(1차)`가 `수율(최종)`까지 예측하던 문제도 고침. 단위가 다른 이름
  (`-y 생산량(kg)`, 타겟 `생산량(t)`)은 자동으로 맞추지 않습니다(`[추천]`, `--find`는 값 환산 안내). 같은 단위의
  띄어쓰기·표기 차이(`생산량( t )`)는 그대로 맞춥니다. `-m`에 모델 파일 하나를 주면 이름이 그 타겟과 맞아야 합니다.
  비슷한 이름 후보(`suggest_columns`)는 단위를 뺀 이름이 같은 열을 길이와 관계없이 모두 앞에 넣습니다(`수율` / `수율(최종공정)`).
- `ml_map.py --points`: 기준 열이 아닌 숫자 열이 기준으로 쓰지 않은 데이터 열과 비슷하면(`반응 압력(bar)`, `압력(MPa)`, 데이터는 `반응압력(bar)`) 메모 열처럼
  붙여 압력을 쓴 것처럼 보이는 결과를 내던 문제. `-i`와 같은 판단으로 `STOP_RETRY`: `[추천]`(조회 파일 머리글 고치기), 단위가 다르면 값 환산 안내.
  어느 데이터 열과도 비슷하지 않은 열(메모, 번호)은 그대로 붙습니다. 2.9.29의 "`--points`도 같음"은 데이터 열과 단위 앞 이름이 같은 열만이었습니다.
- `ml_predict.py -i`: 어느 모델도 쓰지 않는 이름(`압력(bar)=3,압력(MPa)=0.5`의 MPa, `습도=50`, JSON의 `id`)을 경고 없이 버리던 문제(이전부터).
  `STOP_RETRY`: 같은 이름의 다른 단위는 값 환산 안내, `-i`에 없는 모델 입력과 비슷하면 `[추천]`, 여럿이면 `[후보]`, 없으면 `[모델 입력 목록]`과 함께
  사용자에게 묻도록 안내. 파일 입력(`-f`)의 다른 열은 그대로. `--find -i`와 `ml_map.py -i`는 이전부터 거부.
- 모르는 이름의 오타 후보(`resolve_column_names`)에서 다른 이름이 이미 맞춘 열은 뺍니다(같은 이름의 다른 단위는 남김): `--find`/`ml_map.py`의
  `-i "압력=3,속도=45,습도=50"`이 `[추천] '속도'`를 줘서 다시 실행하면 습도 값을 속도로 쓰던 문제.
- 합계 행 경고 (2.9.29 trade-off): 주간/야간 + `일합계`에 그 날의 설정값(`설정온도`, `설정압력`, 하루 안에서 일정)을 그대로 적고 `생산량`만 합한 교대
  보고서를 경고하지 않던 문제(2.9.28은 경고). 바로 위 행들과 위 모든 행의 합 판단에서, 그 행들 안에서 일정한 열의 값과 같은 칸은 숫자 칸 수에서 빼고
  (합도 범주도 말해 주지 않음), 합 칸은 그 열에 0이 아닌 숫자가 2개 이상일 때만 셉니다(0/1 열에 1이 하나면 그 값이 곧 합). 일정한 설정값 + 0/1 열의
  `A계` 범주 행(2.9.29 대상)은 그대로 경고 없음.

### 테스트
- `test_review_v2930.py`: 두 타겟 `--find`/`-y`/`ml_report.load_version`의 `[후보]`와 정확한 이름, 긴 단위 이름 후보, `--points` 비슷한 이름 열(띄어쓰기,
  다른 단위)과 메모·번호 열 통과, 예측 `-i`의 다른 단위·모르는 이름·JSON 키·오타 `[추천]`, 파일 입력 다른 열 유지, `--find -i`/`ml_map.py -i`의
  이미 준 이름 오타 후보 없음, 교대 보고서 `일합계` 10행 경고(csv, xlsx), r18 `fp_const` 경고 없음, `_sums_match`의 일정한 열과 0이 아닌 수.
- 합계 행 무작위 비교(`fuzz_sum.py`/`fuzz_alt.py`, 1500개씩): 점검자 시안과 결과가 모두 같음. 2.9.29 대비 합 모드 1499 -> 1485(줄어든 14개는 모두
  숫자 열이 0/1 열 하나뿐인 표: 한 칸뿐인 근거), 평균 섞음 998 -> 1037, 범주 402 -> 302(오경고 감소), 평균 모드 같음.
- 이전 점검 입력 폴더(r8~r19, 1664개 파일) 비교: 합계 행 결과가 바뀐 파일은 둘. `shift_setpoint.csv`(경고 없음 -> `일합계` 10행, 의도한 수정),
  r17 `FPE_count01_A계.xlsx`(`A계` 2행 경고 -> 경고 없음: 0/1 횟수 열에 1이 하나뿐인 우연한 합, r17에서 오경고로 남겨 둔 것).

## [2.9.29] - 2026-11-03

18차 점검에서 찾은 오류 수정. 회귀 테스트: `tests/test_review_v2929.py` (2.9.28에서 12개 중 10개 실패).

### 수정
- `-i "압력(MPa)=0.3"`(모델은 `반응압력(bar)`): 단위 앞 이름이 다르면 2.9.28의 단위 검사가 빠져 이름 바꾸기 `[추천] '반응압력(bar)'`을 주고, 다시 실행하면
  MPa 값을 bar로 읽어 `VERIFIED`로 틀린 예측을 내던 문제. 두 이름에 모두 단위가 있고 다르면(`°C`·`℃`·`C`와 띄어쓰기 차이는 같은 단위) 이름과 관계없이
  `[추천]` 없이 모델 단위로 값을 환산해 달라는 `STOP_RETRY`. 예측 `-i`·파일 입력 열, `--find -i`, `ml_map.py`의 `-i`·`--find`가 같음 (`--points`는 데이터 열과
  단위 앞 이름이 같은 열만, 다른 열은 2.9.30에서 고침).
- `ml_map.py -i`, `--find -i`의 `[후보]`(데이터 열·모델 입력과 맞추는 곳)에 단위가 다른 후보 표시가 없어, 사용자가 `압력(bar)`를 고르면 0.3 MPa를 0.3 bar로
  쓰던 문제. 예측 `-i`처럼 단위가 다른 후보마다 "(단위 다름: 값 환산 필요)". 값이 없는 이름(`-y`, `-x`, `--vary`)은 표시하지 않음.
- `ml_predict.py --find "생산량(kg)=500"`(타겟 `생산량(t)`): 목표 단위를 보지 않고 500 t로 풀어 `VERIFIED`를 내던 문제. `ml_map.py --find`처럼 값 환산
  `STOP_RETRY`. 단위 없는 목표와 같은 단위(띄어쓰기 차이 포함)는 그대로.
- 합계 행 경고 (2.9.28 회귀): 일정한 열은 어느 행들의 평균과도 같아, 일정한 설정값 열(`설정온도=180`, `설정압력=3.5`)과 우연히 위 두 행의 합인 0/1 열 하나로
  범주 행(`공통`, `공통`, `A계` 반복)의 `A계`를 합계 행으로 경고하던 문제. 평균 칸은 그 행들 안에서 값이 바뀌는 열만 셉니다(`_summary_by_value`와 같은 조건).
  공정 열이 평균이고 `생산량`만 합인 2.9.28 대상 행은 그대로 경고.
- `ml_predict.py --find`: `-i`와 목표에 여러 객체의 JSON 배열을 주면 첫 객체만 경고 없이 쓰던 문제(이전부터). `ml_map.py`처럼 한 행으로 달라는 `STOP_RETRY`
  (객체 하나의 배열은 받음).

### 테스트
- `test_review_v2929.py`: 단위 앞 이름이 다른 단위 다른 이름(예측 `-i`, 파일 입력 열, `ml_map.py -i`/`--find`/`--points`, `--find -i`), 같은 단위 오타의
  `[추천]` 유지, `[후보]` 단위 표시(`ml_map.py`, `--find -i`, 값 없는 이름은 표시 없음), `--find` 목표 단위, 일정한 설정값 열 + 0/1 열 범주 행 경고 없음
  (csv, xlsx), 평균 3열 + 합 1열 `총 합계` 경고 유지, `_sums_match`의 `varies`, `--find` 여러 행 JSON 배열 거부와 한 객체 배열 사용.
- 이전 점검 입력 폴더(r8~r18, 2819개 파일) 비교: 합계 행 결과가 바뀐 파일은 `fp_const.csv` 하나(`A계` 3행 경고 -> 경고 없음, 의도한 수정).

## [2.9.28] - 2026-11-03

17차 점검에서 찾은 오류 수정. 회귀 테스트: `tests/test_review_v2928.py` (2.9.27에서 17개 실패).

### 수정
- 합계 행 경고 (2.9.27 회귀): 넓힌 합계 글자 행은 숫자 칸의 절반 이상이 합이어야 경고해, 공정 열(`온도`/`압력`/`속도`)은 평균이고 `생산량`만 합인
  `총 합계`/`전체합계`/`월합계` 행(2.9.26은 경고), 일자 병합 일일 보고서 맨 아래의 같은 `월합계`/`총 합계` 행, 평균 3열 + 합 1열인 `일합계` 행을
  경고하지 않던 문제. 같은 행들의 평균과 같은 칸(같은 허용 오차)은 판단에서 뺍니다: 합 칸 1칸 이상, (합 칸 + 평균 칸) * 2 >= 숫자 칸이면 경고.
  평균 칸만 있고 합 칸이 없으면 경고하지 않습니다. 데이터보다 훨씬 큰 숫자 규칙은 그대로(정확한 합이 없는 범주 행이 평균 근처 값으로 경고되지 않게).
- `-i "압력(MPa)=0.3"`(모델은 `압력(bar)`): 이름 바꾸기 `[추천] '압력(bar)'`을 줘서, 다시 실행하면 MPa 값을 bar로 읽어 `VERIFIED`로 틀린 예측을 내던
  문제(2.9.13부터). 두 이름에 모두 단위가 있고 다르면 `[추천]` 없이 모델 단위로 값을 환산해 달라는 `STOP_RETRY`(모델 단위 이름 포함).
  `°C`·`℃`·`C`와 띄어쓰기 차이는 같은 단위(`[추천]` 그대로), 대소문자가 다르면 다른 단위. `[후보]`가 여럿이면 단위가 다른 후보에 "(단위 다름: 값 환산 필요)". 파일 입력 열, `--find`의 `-i`, `ml_map.py`의 `-i`·`--find`·`--points`도 같음.
- `-i "온도=180,온도(°C)=182"` (2.9.27): 모델을 보지 않고 거부해, 피처에 `온도`와 `온도(°C)`가 둘 다 있는 모델을 `-i "k=v"`로 쓸 수 없던 문제.
  모델 입력과 이름을 맞추는 곳(`ml_map.py`는 데이터 열)에서 판단: 단위 없는 이름이 그 자체로 피처/열이 아니고 단위 붙은 이름과 같은 입력으로 갈 때만
  `STOP_RETRY` + `[후보]`.
- JSON `-i`: 한 객체의 같은 키 두 번(`{"온도":180,"온도":190}`)과, 같은 피처로 가는 별칭 쌍(`{"온도":180,"온도(°C)":190}`, 모델은 `온도(°C)`)에서 한 값을
  경고 없이 쓰던 문제(이전부터). `"k=v"`처럼 `STOP_RETRY` + `[후보]`. .json 파일 경로 입력도 같음.

### 테스트
- `test_review_v2928.py`: 평균 3열 + 합 1열 `총 합계`/`합계`/`전체합계`/`Grand Total`/`월합계`, 비율 열, 교대 3개 일자 병합 일일 보고서 4종(11행 모두),
  범주 행(`A계` 편차, `1계`/`2계`, 평균과 같은 온도 + 합이 아닌 생산량) 경고 없음, `_sums_match` 단위 검사, 단위 다른 이름(ml_predict, ml_map, `--find`),
  `온도`+`온도(°C)` 두 피처 모델 사용, 한 피처 모델 거부(k=v, JSON, JSON 배열, ml_map, `--find`), JSON 같은 키(문자열, .json 파일).
  의도적으로 바꾼 기대값: `test_review_v2927.py::test_kv_alias_and_full_name_refused`는 `parse_inline_input` 대신 `check_same_column`으로 검사
  (분석 단계에서는 모델을 몰라 거부하지 않음).

## [2.9.27] - 2026-11-03

16차 점검에서 찾은 오류 수정. 회귀 테스트: `tests/test_review_v2927.py` (2.9.26에서 19개 실패).

### 수정
- 합계 행 경고 (2.9.26 회귀): 넓힌 합계 글자(`...합계`, 짧은 `...계`, 낱말 `Total`/`Sum`)를 채운 칸 수로만 판단해, 일자가 병합된 일일 보고서의
  `일합계`(날마다) + 맨 아래 `월합계`/`총 합계`/`전체합계`, `주간계` + `월간계`(합계 행이 일자 칸보다 많음)를 하나도 경고하지 않고(2.9.25는 11개),
  일자·라벨 열이 비고 `총 합계` 1칸만 있는 파일도 경고하지 않던 문제. 한편 범주 값이 열의 절반 미만이면(`생산계` 9 + `품질팀` 11, 센서 사이의 `pH계`,
  `레벨계`, `A계`/`B계` + `공통`) 모든 행을 경고(2.9.25부터). 글자 수 규칙을 바꾸는 것으로는 한쪽이 다시 깨져, 넓힌 합계 글자 행은 숫자로 판단합니다:
  숫자 칸의 절반 이상(1칸 이상)이 바로 위 데이터 행들의 합, 위 합계 행 가운데 한 글자 행들의 합(`월합계` = `일합계`들), 위 모든 데이터 행의 합이면 경고
  (칸에 보이는 자릿수의 반올림 또는 상대 0.1% 이내, 2행 이상이고 모두 0은 아닌 열만). 합이 아니어도 넓힌 글자 행이 열 안에서 적으면(글자 2개 이하,
  다른 채운 칸 수 + 1 이하, 행 수의 절반 미만) 숫자가 데이터보다 훨씬 클 때(최대 + max(범위, 최대/2) 초과) 경고: 숫자를 고친 `월합계 9999`,
  `주간계 2100`, 작은 파일의 `총 합계 999`. 범주 값은 숫자가 데이터 행과 같은 범위라 경고 없음. 정확한 낱말(`합계`, `소계`, `계`, `일계`, `누계`, `Total` 등),
  통계 이름 열, `평균값`/`평균치`(적을 때만) 규칙은 그대로. 교대가 1개라 `일합계`가 그 행과 같은 값인 보고서는 경고하지 않습니다(한 행은 자기 자신의 합).
- `-i "k=v"`: 같은 이름이 두 번 나오면(`온도=195,압력=3.8;온도=165,압력=2.1`) 뒤 값으로 덮어써 한 행만 `VERIFIED`로 예측하던 문제(2.9.26 이전부터).
  `STOP_RETRY`(한 행 입력, 여러 행은 JSON 배열이나 .json 파일 안내). 단위 없는 이름과 같은 이름의 단위 붙은 이름(`온도=180,온도(°C)=190`)을 함께 주면
  `온도(°C)` 값만 경고 없이 쓰던 문제: `STOP_RETRY` + `[후보] 온도=180, 온도(°C)=190`. `--find`, `ml_map.py -i`도 같음.

### 테스트
- `test_review_v2927.py`: 일자 병합 일일 보고서(`일합계` + `월합계`/`총 합계`/`전체합계`, 교대 3개, 날짜 칸, `주간계` + `월간계`), `일합계`의 합인 `월합계`,
  근무조 열의 `주간계`/`월간계`, `비고` 열과 빈 라벨 열의 `총 합계`, 4~6행 파일의 `총 합계`(일자 있음/없음), 범주 값(`A계`/`B계`, `생산계`/`품질팀`, `pH계`,
  `A계`/`B계`/`공통`, `레벨계`, r15 L1/L2/L3/L6) 경고 없음, 합이 맞는 일일 합계 글자 11개, 학습 경고, `-i` 같은 이름·단위 이름 거부(ml_map 포함).
  의도적으로 바꾼 기대값: `test_review_v2926.py::test_loose_words_with_many_texts_are_a_category`는 `A합계`/`B합계`/`C합계` 행이 위 두 행의 합이라
  이제 합계 행(숫자로 판단). 범주 검사는 데이터 행과 같은 범위의 숫자로 바꾸고(경고 없음), 합인 경우는 12행 경고로 검사.

## [2.9.26] - 2026-11-03

15차 점검에서 찾은 오류 수정. 회귀 테스트: `tests/test_review_v2926.py` (2.9.25에서 11개 실패).

### 수정
- 합계 행 경고 (2.9.25 회귀): 4글자 이하로 `계`로 끝나는 말을 모두 합계 글자로 보아 범주 열(`A계`/`B계`, `1계`/`2계`, `생산계`/`품질계`/`설비계`,
  `레벨계`/`pH계`/`유량계`, `고온계`, `저항계` 등)의 모든 행을 합계 행으로 경고하던 문제. 학습·표 보간은 모든 행을 `제외`에 표시하라고, 예측은 모든 예측값이
  데이터 행의 예측이 아니라고 알림. 보통 낱말 목록으로는 다 막을 수 없어, 합계 글자를 정확한 낱말(`합계`, `소계`, `총계`, `계`, `총합`, `일계`/`월계`/`주계`/
  `연계`/`년계`/`누계`, `Total`, `Sum`, `Subtotal`, `평균`, `Avg`, `Mean`)과 넓힌 글자(`...합계`, 짧은 `...계`, 낱말로 든 `Total`/`Sum`, `평균값`/`평균치`)로 나눕니다.
  넓힌 글자는 열 안에서 데이터 사이의 합계 행처럼 보일 때만 셉니다: 채운 칸의 절반 이하, 행 수의 절반 미만, 서로 다른 글자 2개 이하
  (일일 보고서의 `총 합계`, `주간계`, `Grand Total`은 그대로 경고). 정확한 낱말과 통계 이름 열 규칙은 그대로.
- 같은 경고: `Total`/`Sum`이 글자 안에 붙어 있어도(`SUM-01`, `Lot-Sum`, `sum_a`, `Total_Loss`) 합계 글자로, `평균`으로 시작하는 6글자 이하의 이름
  (`평균온도`)을 평균 글자로 보던 문제. `Total`/`Sum`은 띄어 쓴 낱말일 때만, `평균` 뒤에는 `값`/`치`만 평균 글자입니다.

## [2.9.25] - 2026-11-03

14차 점검에서 찾은 오류 수정. 회귀 테스트: `tests/test_review_v2925.py` (2.9.24에서 21개 실패).

### 수정
- 읽지 못한 숫자 칸 경고 (2.9.24 회귀): 불러온 모든 행을 봐서, 머리글 아래에서 자동으로 찾은 단위 행(타겟이 비어 학습하지 않음)의 `°C`, `bar`,
  `wt%`, `%`, `L/min`을 SKILL 2단계 예제에서 거짓 경고 5개로 알리고, 경고 5개 상한 때문에 뒤쪽 열의 진짜 경고가 빠지던 문제.
  학습하는 행(타겟이 있고 `제외`에 표시되지 않은 행, 단위 행 아님)만 봅니다. 타겟이 여러 개여도 열마다 한 번, 5열을 넘으면 "N개 더"로 알립니다.
- 같은 경고: 채운 칸의 절반 이상이 숫자인 열을 모두 봐서 범주 피처(금형 코드 `101`/`A1`, 소수점 쉼표가 40%라 범주가 된 열)와
  `--dtypes "<열>:datetime"` 열도 "중앙값으로 채웠습니다"로 알리던 문제. 모델이 숫자로 읽는 열(숫자 피처)만 봅니다.
- 같은 경고: `--eval` 원본 열(`dT=T1-T2`의 `T1`)의 `175,6` 칸을 알리지 않던 문제. 원본 열도 봅니다.
- 합계 행 경고: 합계·평균 글자를 정확히 같을 때만 찾아 일일 보고서의 `일계`, `월계`, `누계`, `일합계`, `총 합계`, `전체합계`, `합계:`, `Total:`,
  `소계(A)`, `주간계`, `평균값`, `AVG.` 행을 경고 없이 학습(`VERIFIED`)하던 문제. 띄어쓰기·괄호 부분·끝의 `:`·`.`를 빼고, `합계`로 끝나는 말,
  기간 합계(`일계`/`월계`/`주계`/`연계`/`년계`/`누계`), 4글자 이하로 `계`로 끝나는 말(`설계`, `단계`, `통계`, 계기 이름 등 보통 낱말은 아님),
  `Total`/`Sum`이 낱말로 든 짧은 글을 합계 글자로, `평균`으로 시작하는 짧은 말을 평균 글자로 봅니다. 통계 이름 열 규칙은 그대로.
- 소수점 쉼표 글자 피처: 칸이 모두 `0,149`, `175,6` 같은 글자인 피처를 경고 없이 값 종류가 많은 범주로 학습하던 문제(13차부터 남음).
  선택한 피처의 채운 칸 절반 이상이 그런 글자면 `STOP_RETRY`(쉼표를 점으로 바꾸거나 숫자 칸으로 저장 안내), 그보다 적으면 경고.
  `--dtypes`로 형식을 정한 열은 보지 않습니다.

## [2.9.24] - 2026-11-03

13차 점검에서 찾은 오류 수정. 회귀 테스트: `tests/test_review_v2923.py` (2.9.24 부분).

### 수정
- 합계 행 경고 (2.9.22 회귀, 2.9.23 수정이 덜 됨): 일자가 병합된 일일 보고서에 일별 `합계`와 맨 아래 `평균`/`최대`/`최소` 묶음이 같은 열에 있으면
  `최대` 한 칸 때문에 통계 이름 열로 보아 일별 합계 행이 경고 없이 학습되던 문제. 통계 이름 열은 집계·평균 글자가 합계 글자의 절반 이상일 때만입니다.
- 숫자 읽기: `0,149`, `00,150`처럼 0으로 시작하는 쉼표 숫자를 천 단위로 보아 149로 읽거나 `-i`에서 `0149`를 `[추천]`하던 문제.
  천 단위 묶음은 1~9로 시작할 때만입니다(읽을 수 없는 값으로 둠).
- 학습: 숫자 열(채운 칸의 절반 이상이 숫자)의 글자 칸(`175,6` 등)을 경고 없이 중앙값으로 채우던 문제. 결과 `warnings`에 열 이름, 칸 수, 예를 알립니다.

## [2.9.23] - 2026-11-03

12차 점검에서 찾은 오류 수정. 회귀 테스트: `tests/test_review_v2923.py`.

### 수정
- 합계 행 경고 (2.9.22 회귀): 합계·소계 글자가 채운 칸의 절반 이상이면 통계 이름 열로 보아 경고하지 않았는데, 최대·최소·평균 같은 말이
  없어도 그렇게 봐서 세로로 병합된 일자·라인 열(맨 위 칸만 채워짐) 아래의 `합계`/`소계` 행이 경고 없이 학습되던 문제.
  통계 이름 열은 집계 단어(최대·최소 등)나 평균 단어가 함께 있을 때만으로 봅니다.
- 숫자 읽기: `coerce_numeric`이 쉼표를 모두 지워 엑셀 글자 칸이나 JSON `-i`의 `175,6`(소수점 쉼표)이 1756으로 바뀌던 문제.
  천 단위 묶음(`1,500`, `1,234,567.5`)의 쉼표만 지우고, 그 밖의 쉼표 숫자는 읽을 수 없는 값(결측 경고)으로 둡니다.

## [2.9.22] - 2026-11-03

11차 점검에서 찾은 합계 행 경고 누락(2.9.21 회귀), 통계 이름 열 읽기 속도, 천 단위 쉼표 평균 행, 소수점 쉼표 CSV의 점 묶음 숫자 문제 수정.
이번에는 추측보다 단순하고 안전한 동작을 택함.

### 수정 (학습·예측·표 보간)
- 합계 글자 행 경고 누락 (2.9.21): 합계 글자(`합계`, `소계`, `총계`, `계`, `Total`, `Sum`, `Subtotal`)를 다른 집계 글자가 있는 열에서 위 행들의 합과 같을 때만
  경고해, 하단 블록(`합계`/`평균`/`최대`/`최소`)의 비율 열(`불량률`)이 든 `합계` 행, `소계` 묶음 아래 총 `합계` 행, `평균`이 범주인 열(`개별`/`평균`)의 `소계`
  행을 경고 없이 학습(`VERIFIED`)하던 문제. 합계 글자는 값과 관계없이 다시 항상 경고(2.9.20과 같음). 예외는 통계 이름 열뿐: 행의 절반 이상에 값이 있고,
  합계·집계 글자가 2칸 이상이면서 값이 든 칸의 절반 이상이거나 값이 든 칸이 모두 합계·평균·집계 글자인 열(`집계` = `합계`/`평균`/`최대`/`최소`, `Sum`/`Avg`).
  평균 글자의 값 비교(2.9.21)는 그대로.
- 읽기 속도 (2.9.21): 값 비교가 연속된 합계·집계 행을 후보마다 거슬러 올라가(행 수의 제곱) 통계 이름 열 10만 행 CSV 읽기가 111초(2.9.20은 0.7초).
  위 데이터 행 범위를 한 번에 계산하고, 비교할 행이 없으면 숫자 표도 만들지 않음 (같은 PC에서 0.31초, 2.9.20 0.37초).
- 천 단위 쉼표 평균 행: Excel `#,##0` 형식으로 저장한 CSV(`"2,452"`)의 교대별 `평균` 행을 값 비교에서 글자로 보고 경고하지 않던 문제. 학습과 같은 숫자 읽기
  (`coerce_numeric`)로 비교.
- 소수점 쉼표 CSV의 점 묶음 숫자: 2.9.21의 `1.300` -> 1300 읽기를 없앰. 소수점 쉼표(`175,5`, 탭 CSV의 라인 목록 `1,2` 포함)나 `1.346,8`이 확인된 `;`·탭 CSV에
  쉼표 없이 점으로 묶인 숫자(`1.300`, `2.125`)가 있으면 항상 `STOP_RETRY`(한 가지 숫자 형식으로 또는 .xlsx로 저장 안내). 2.9.21은 탭 CSV의 `1,2`와 두께
  `2.125`를 함께 두면 2125로 학습(`VERIFIED` + 경고).
- 학습·예측 읽기 차이: 소수점 쉼표 파일로 학습한 모델로 소수점 쉼표 칸이 없는 예측 파일(`1.300`)을 예측하면 1.3으로 읽고 `VERIFIED`(학습은 천 단위)이던 문제.
  학습 판단을 `model_<타겟>.json`의 `decimal_comma`에 남기고, 예측 파일은 구분 기호와 관계없이 같은 방식으로 읽음: 소수점 쉼표 칸이 없어도 `1.300`, `1,500`이
  있으면 `STOP_RETRY`. `ml_map.py --points` 조회 파일도 `-f` 데이터 파일의 판단을 따름. 이전 모델(키 없음)은 소수점 쉼표가 아닌 것으로 봄(베타, 다시 학습).

### 테스트
- `test_review_v2922.py`: 통계 이름 열 읽기 속도(5천 행 대비 4만 행 비율), 하단 블록·`소계` 묶음·`평균` 범주 열의 합계 행 경고, 통계 이름 열과 빈 `비고` 열의
  `소계`, 천 단위 쉼표 교대 평균 행, 탭 CSV `1,2` + `2.125` 거부, 모델·`--points`의 소수점 쉼표 판단. 2.9.21에서 8개 실패.
- 의도적으로 바꾼 기대값: `test_review_v2921.py`의 `1.300` 읽기 두 검사(이제 거부, `test_dot_grouped_integers_refused`, `..._train_refused`)와 통계 이름 열의
  진짜 합계 검사(이제 경고 없음, `test_statistic_column_total_not_value_checked`), `test_review_v2920.py`의 `1.346,8` 파일 속 `1.346`(이제 거부).

## [2.9.21] - 2026-11-02

10차 점검에서 찾은 평균 행 경고 누락, 통계 이름 열의 거짓 경고, 소수점 쉼표 CSV의 점 천 단위 정수 문제 수정.

### 수정 (학습·예측·표 보간)
- 평균 행을 값으로 판단: 교대 보고서(`구분` = `A조, B조, C조, 평균`, 날마다 반복)의 `평균`이 열의 25%라 범주 값으로 보고 경고 없이 학습(`VERIFIED`)하던 문제
  (2.9.19부터, `소계`가 함께 있으면 `소계`만 경고). 평균 글자 행의 숫자 칸이 모두 바로 위 데이터 행들(위의 합계·평균·집계 행이나 첫 행까지, 2행 이상)의
  평균과 같으면(칸에 보이는 자릿수의 반올림 또는 상대 0.1% 이내, 한 열 이상은 값이 달라야 함) 비율과 관계없이 경고. 합계 글자 행은 합과 같으면 확정.
  평균 값이 위 행들의 평균과 다른 범주 열(`평균`/`측정`)은 그대로 경고 없음.
- 거짓 경고 (2.9.20): 통계 이름 열(`집계` = `합계`/`평균`/`최대`/`최소`, `통계` = `Sum`/`Avg` 번갈아)의 `합계`·`Sum` 행을 모두 경고하고, 모든 행이 `평균`인 열의
  모든 행을 경고하던 문제. 다른 집계 글자가 있거나 평균 글자가 범주인 열(모든 행이 같은 평균 글자 포함)의 합계 글자는 위 행들의 합과 같을 때만 경고.
- 소수점 쉼표가 확인된 `;`·탭 CSV(`175,5`)에서 쉼표 없이 점으로 묶인 정수(`1.300`)를 1.3으로 학습(`VERIFIED`)하던 문제 (2.9.20 이전부터).
  모두 1~9로 시작하고(`1.300`, `12.345.678`), 같은 열의 점이 든 칸이 모두 점 천 단위 숫자이며, 파일에 다른 소수점 숫자(`1.5`, `0.300`)가 없으면 1300으로 읽고
  `warnings`("쉼표 없는 숫자 점(.)을 천 단위 구분 기호로 읽었습니다", 칸 수와 예). 아니면 `STOP_RETRY`. `1.346,8`과 함께 있는 `1.346`도 같은 규칙(이전에는
  `STOP_RETRY`, 이제 1346 + `warnings`; `test_review_v2920.py`의 해당 검사를 바꿈).

### 테스트
- `test_review_v2921.py`: 교대별 평균 행(소계 포함, 학습 포함) 경고, 평균 범주는 경고 없음, `집계`·`Sum/Avg`·모든 행 `평균` 열은 경고 없음과 진짜 합계는 경고,
  `1.300` CSV(학습 포함)와 섞인 형식(`1.5`, `0.300`, 같은 열의 다른 점 글자) 거부. 2.9.20에서 7개 실패.

## [2.9.20] - 2026-11-02

9차 점검에서 찾은 소수점 쉼표 CSV의 값 오류, 합계 행 경고 회귀, 단위 `ng`, 읽기 속도 문제 수정.

### 수정 (학습·예측·표 보간)
- `;` CSV의 소수 4자리 이상 쉼표 숫자(`27,8448`)를 천 단위로 보고 278448로 학습(`VERIFIED`)하거나, `175,6`과 함께 있으면 거짓 `STOP_RETRY`이던 문제 (2.9.19).
  `ml_predict.py`가 `_THOUSANDS_RE`를 두 번 정의해(앞은 처음부터 끝까지, 뒤는 `-i` 입력용) 소수점 쉼표 판단이 뒤의 정규식으로 앞부분만 비교했음.
  이름을 나누고(`_CSV_THOUSANDS_RE`, `_KV_THOUSANDS_RE`) 전체 일치(`fullmatch`)로 비교.
- 0으로 시작하는 쉼표 숫자(`0,149`)는 천 단위가 될 수 없어 소수점 쉼표로 봄: 쉼표 숫자가 모두 `0,149` 꼴인 파일을 149로 학습(`VERIFIED`, 경고 없음)하던 문제와
  `176,0`과 함께 있으면 거짓 `STOP_RETRY`이던 문제.
- 점이 천 단위인 유럽식 숫자(`1.346,8`)를 1.3468로 학습(`VERIFIED`)하던 문제. 쉼표와 점이 든 숫자 칸이 모두 그 형식이고 쉼표 없이 점으로 묶인 칸(`1.346`)이
  없으면 1346.8로 읽고 `warnings`("숫자 점(.)을 천 단위 구분 기호로 읽었습니다", 칸 수와 예). 다른 형식(`1,346.8`, `1.346`)과 섞이면 `STOP_RETRY`.
- 합계 행 경고 회귀 (2.9.19): 합계 글자가 열의 과반이거나 20% 이상이면 범주 값으로 보아 교대별 `소계`(25%), 2라인 + `합계`(33%), 빈 `비고` 열의 `합계` 1개를
  경고 없이 데이터로 학습(`VERIFIED`)하던 문제. 범주로 보는 것은 평균 글자(`평균`, `Average`, `Avg`, `Mean`)만이고, 과반 기준은 그 글자 말고도 값이 든 칸이 있을 때만.
  합계 글자(`합계`, `소계`, `총계`, `계`, `Total`, `Sum`, `Subtotal`)는 항상 경고. `집계` = `평균`/`최대` 열은 그대로 경고 없음.
- 빈 값 표시(`MISSING_MARKS`)에서 `ng`를 뺌: 단위 행 `ng, ng`의 `--unit-row 2`를 "빈 칸과 빈 값 표시뿐"이라 거부하고, `ng, °C` 단위의 이름이 `농도`(단위 없음)이던
  문제 (2.9.19). 대소문자를 구분해 `NG`만 빈 값 표시.

### 성능
- 큰 파일 읽기가 느리던 문제: pandas가 열 연산마다 `df.attrs`의 행 번호 목록(`source_rows`, 행마다 1개)을 깊은 복사함. 합계 행 판단(`_summary_rows`)을
  배열 한 번 읽기로 바꾸고(숫자 칸은 건너뜀), 행 번호·머리글 후보 목록은 복사하지 않는 목록(`_SharedList`)으로 둠(학습·예측의 열 연산도 빨라짐).
  10만 행 x 32열 CSV 읽기(`load_data_file`) 2.9.18 4.8초, 2.9.19 6.0초 -> 1.0초.

### 테스트
- `test_review_v2920.py`: 모듈 수준 이름이 두 번 정의되지 않음(AST), `27,8448`·`0,149`·`1.346,8` CSV(학습 포함)와 섞인 형식 거부, 교대별 소계·2라인 합계·
  빈 열의 합계 1개 경고, 평균 범주와 `제외` 표시는 경고 없음, 단위 `ng` 행(학습 포함)과 `NG` 빈 값 표시, `attrs`를 복사하지 않음. 2.9.19에서 9개 실패.

## [2.9.19] - 2026-11-02

8차 점검에서 찾은 단위 행 판단·합계 행 경고·소수점 쉼표 CSV 문제 수정. 단위 행 판단은 칸마다 그 열 아래 값과 비교해 정하도록 다시 설계하고
2.9.17/2.9.18 매트릭스를 모두 통과.

### 수정 (학습·예측·표 보간)
- 단위 행 판단을 칸 단위로 다시 설계 (`unit_row_verdict`, `_looks_like_unit_row`가 씀). 칸마다 그 열 아래 값(반복된 머리글·단위 행은 뺌)과 비교해
  단위 칸(아래 값에 없는 짧은 글자, 아래 숫자 범위에서 크게 벗어난 배율 숫자 `1000`), 데이터 칸(범위 안의 숫자·날짜, 아래에도 있는 글자 `A`),
  빈 칸(빈 값 표시 `측정불가`, `NG`, `결측`, `없음` 등 `MISSING_MARKS`)으로 나누고, 데이터 칸이 없고 단위 글자 칸이 있을 때만(기존의 열 과반 기준과 함께) 단위 행.
  - 타겟 열이 없는 예측 파일의 첫 행 `측정불가, 측정불가, A`(`NG, NG, A`)를 단위 행으로 보고 진짜 데이터 행을 뺀 `[추천] --data-row 3`이던 문제 (2.9.18은 숫자 칸에 의존).
    이제 모든 행을 예측(`VERIFIED`). 범주 열이 없는 `측정불가, 측정불가`는 결측 값인 데이터 행.
  - 첫 데이터 행이 `측정불가, 측정불가, A, <빈 타겟>`인 학습 파일의 `--unit-row 2`를 받아 `반응온도(측정불가)`, `라인(A)` 이름으로 학습(`VERIFIED`)하던 문제.
    "아래 데이터와 같은 값이 든 데이터 행" `STOP_RETRY` + `[추천] <--unit-row를 뺀 전체 명령>`.
  - 배율 숫자가 든 단위 행(`°C, 1000, , %`)을 숫자가 든 데이터 행이라 거부해 `[추천]`이 단위 행을 빼고, 예측은 단위 행을 데이터(`압력=1000`)로 예측(`VERIFIED`)하던
    문제 (2.9.18 회귀). 이제 단위 행.
  - 이름 없는 머리글 칸 아래의 `비고`가 든 단위 행을 거부해, 옵션 없는 예측이 단위 행을 데이터로 예측(7행, `VERIFIED`)하던 문제 (2.9.18 회귀, 2.9.17은
    `[추천] --data-row 3`). 이름 없는 머리글 칸 아래 칸은 그 열 아래에 숫자·날짜가 없으면 판단에서 뺌(있으면 제목 행 아래의 머리글이라 거부). 이제 `[추천] --unit-row 2`
    (단위가 든 모델) / `--data-row 3`(단위 없는 모델), 6행. 같은 파일의 학습 `--unit-row 2`도 받음.
- 예측의 첫 행 비율(`_first_row_text_share`)과 메모 행 판단이 빈 값 표시(`측정불가`, `NG`)를 빈 칸으로 봄. 단위 행 기준에 맞지 않아도 첫 행의 숫자 입력 칸 과반이
  글자(숫자로 시작하는 `<0.1`, `12%`는 아님)이고 `--data-row <그 행+1>`로 읽은 첫 행이 숫자면 `--data-row`를 추천 (2.9.17의 대체 경로, `_fit_unit_layout`).
- 단위 행의 빈 값 표시(`NG`, `N/A`)는 이름에 붙이지 않음 (`_with_unit`, 이전에는 `-`만).
- 합계 행 경고가 안내대로 `제외` 열에 `합계`를 적은 뒤에도 계속 나오던 문제 (학습·표 보간). `제외` 열에 `포함`이 아닌 표시가 있는 행과 `제외` 열 자체의 글자는 보지 않음.
  `제외` 칸 해석(`0`/`FALSE`는 표시 없음)은 `ml_predict.exclude_mark` 하나 (학습이 씀).
- 범주 입력 열 `집계`(`평균`/`최대`)의 `평균` 21행을 합계·평균 행으로 경고하던 문제. 합계 글자가 그 열의 값이 든 칸의 과반이거나, 3칸 이상이면서 20% 이상이거나,
  다른 집계 글자(`최대`, `최소`, `중앙값` 등)와 같은 열이면 범주 값으로 봄. 몇 번 반복된 소계 행은 그대로 경고.
- 쉼표가 아닌 구분 기호(`;`, 탭)의 CSV에서 소수점 쉼표(`175,6`)를 천 단위 구분 기호로 지워 1756으로 학습(`VERIFIED`)하던 문제 (이전부터). 천 단위로 읽을 수 없는
  쉼표 숫자가 있으면 `<숫자>,<숫자>` 칸을 소수점으로 읽고 `warnings`("숫자 쉼표를 소수점으로 읽었습니다", 칸 수와 예). 천 단위로도 읽히는 칸(`1,500`)이
  함께 있으면 `STOP_RETRY` (값이 1000배 달라질 수 있음). 쉼표 CSV와 소수점 쉼표가 없는 파일은 그대로.
- 행이 적은 `;` CSV(제목 줄 + 머리글 + 단위 + 소수점 쉼표 숫자 6행)의 구분 기호를 `,`로 잘못 골라 예측이 열을 못 찾던 문제. `csv.Sniffer`의 결과보다
  표 폭으로 나뉘는 줄이 더 많은 구분 기호가 있으면 그것을 씀 (`_sniff_delimiter`).

### 테스트
- `test_layout_matrix_v2919.py`: 레이아웃 2개(이름 없는 머리글 칸 아래 `비고`가 든 단위 행, 배율 숫자 단위 행)의 학습(맞는 옵션, 옵션 없음)·예측(단위 있는/없는 모델,
  템플릿·`next_commands`), 타겟 열이 없는 예측 파일의 빈 값 표시 첫 행 3개, `측정불가` 첫 행의 `--unit-row` 거부, 칸 단위 규칙, `제외` 표시와 집계 범주의
  합계 행 경고(학습·표 보간), 소수점 쉼표 CSV(소수점으로 읽기, 행이 적은 파일의 구분 기호, 천 단위와 섞이면 거부, 쉼표 CSV 그대로). 2.9.18에서 22개 실패.

## [2.9.18] - 2026-11-02

7차 점검에서 찾은 단위 행 판단 회귀와 레이아웃 읽기 문제 수정.

### 수정 (학습·예측·표 보간)
- 숫자가 든 칸이 있는 첫 데이터 행(`측정불가, 측정불가, A, 61.66`)을 단위 행으로 받던 문제 (2.9.17 회귀). `--unit-row 2`로
  `반응온도(측정불가)` 같은 이름으로 학습(`VERIFIED`)하고, 예측은 진짜 데이터 행을 빼는 `[추천] --data-row 3`이었음.
  이제 단위 행은 값이 든 칸에 숫자·날짜가 없어야 함 (`_looks_like_unit_row`). 이런 `--unit-row`는 "숫자·날짜가 든 칸이 있는 데이터 행"이라는 `STOP_RETRY`.
- 예측의 레이아웃 확인이 첫 데이터 행 하나만 봐서, 첫 데이터 행의 숫자 입력이 비었거나 모두 `측정불가`이면 그 위의 단위 행을 데이터로 예측(`VERIFIED`,
  6행 대신 7행)하던 문제. 이제 머리글 바로 아래 행을 파일에서 직접 판단: 단위 행이면 `--unit-row`(단위가 든 모델) 또는 `--data-row`(단위 없는 모델),
  긴 글자의 메모 행이면 `--data-row`, 숫자가 든 칸이 있으면 데이터 행이라 추천하지 않음 (`_fit_unit_layout`). 첫 행 비율(`_first_row_text_share`)은
  숫자 입력이 든 첫 행(최대 10행)으로 봄.
- 표보다 좁은 제목 줄이 있는 CSV(줄마다 칸 수가 다름)를 "Expected 1 fields in line 2, saw 4" `STOP_RETRY`로 못 읽어 머리글 행 안내가 없던 문제.
  `csv.reader`로 가장 넓은 줄의 칸 수를 세어 모든 줄을 그 폭으로 읽음 (인코딩 자동 판별은 그대로).
- 단위가 든 모델 이름(`반응온도(°C)`)이 단위 없는 머리글 칸(`반응온도`)과 맞지 않아, 제목 행 + 머리글 파일을 예측하면 `[추천] --header-row`가 없던 문제.
  예측의 머리글 행 찾기(`find_header_layout`, `names_in_cells(model_names=True)`)는 예측의 열 이름 맞추기와 같이 단위를 뗀 이름으로도 맞춤.
  학습·표 보간의 `-x`/`-y` 이름은 그대로(단위는 이름의 일부).

### 기능
- 데이터 중간에 머리글 행과 똑같은 행(페이지마다 반복된 머리글)과 바로 다음의 단위 행 반복은 빼고 `warnings`에 행 번호 (학습·예측·표 보간).
  값이 없는 행이라 데이터 값이 바뀌지 않고, 숫자 열이 숫자로 남음 (이전에는 숫자 열이 글자 열이 되거나 그 행을 예측).
- 합계·평균 행(`합계`, `총계`, `평균`, `소계`, `계`, `Total`, `Sum`, `Average` 같은 글자 칸과 숫자가 같은 행)은 빼지 않고(데이터 값의 판단은 사용자)
  `warnings`에 행 번호. 학습·표 보간은 "'제외' 열에 이유를 적고 다시 실행" 안내, 예측은 그 행의 예측값이 데이터 행의 예측이 아니라는 안내.
  타겟 열의 글자(분류 라벨 `계`)는 첫 열일 때만 봄.

### 테스트
- `test_layout_matrix_v2918.py`: 2.9.17 매트릭스에 레이아웃 8개 추가(첫 행 숫자 입력이 모두 `측정불가`(단위 행 있음/없음), 단위 행 아래 첫 행이 타겟만,
  제목 줄이 좁은 CSV(단위 행 있음/없음), 제목 2칸 + 머리글을 단위 모델로 예측, 반복된 머리글+단위 행, 끝의 합계 행). 같은 기준(맞는 옵션의 학습,
  옵션 없는 학습, 단위 있는/없는 모델의 템플릿·`next_commands` 예측)과 규칙 단위 테스트.

## [2.9.17] - 2026-11-02

6차 점검에서 찾은 단위 행·머리글 행 판단 오류 수정. 2.9.14부터 매번 회귀한 부분이라 기준을 하나로 정리하고 레이아웃 매트릭스 테스트로 고정.

### 수정 (학습·예측·표 보간)
- 단위 행 판단이 바로 아래 한 행만 봐서, 첫 데이터 행에 글자 값 하나(`측정불가`, `NG`, `<0.1`, `12%`, 날짜 열의 날짜)가 있으면 진짜 단위 행을 거부하던 문제
  (2.9.16 회귀). 맞는 `--unit-row 2`가 "단위 행이 아니라 데이터 행"이라는 앞뒤가 맞지 않는 `STOP_RETRY`였고 `[추천]`은 단위 행을 빼고,
  옵션 없는 템플릿 예측은 단위 행을 데이터로 예측(`VERIFIED`). 이제 단위 칸이 든 열마다 그 아래 값이 든 칸 최대 10개의 과반이 숫자나 날짜(날짜 칸 포함)인지 보고,
  단위 칸의 과반이 숫자가 아닌 글자이고 그 열들의 과반이 아래가 숫자·날짜면 단위 행 (`_looks_like_unit_row`). 단위 칸은 짧은 글자(12자 이하, 띄어쓰기 1번 이하)만,
  바로 아래 값이 든 행이 그 열들에서 모두 글자면(단위 행 위의 머리글) 단위 행이 아님.
- 주어진 `--unit-row`가 빈 행이거나 머리글 행보다 아래가 아니면 말없이 받아(단위 없음), 진짜 단위 행을 데이터로 학습·예측(`VERIFIED`)하던 문제
  (머리글 1행 + 단위 2행으로 학습한 모델의 `next_commands`를 제목 행이 있는 파일에 쓴 경우). `STOP_RETRY`, 예측은 파일을 다시 읽어 확인한 `[추천]`
  (단위가 든 모델은 `--unit-row 4`, 단위 없는 모델은 `--data-row 5`).
- 제목 행(칸 2개 이상)과 진짜 머리글을 "머리글+단위 행 쌍"으로 봐서 `[추천]`을 두 번 따라야 하던 문제. 단위 칸은 모두 이름이 든 머리글 칸 아래에 있고
  (단위 행이 머리글 행보다 넓지 않음) 머리글 행은 대부분 글자여야 함 (`_header_unit_pair`). 주어진 `--unit-row`가 제목 행 아래의 진짜 머리글이면
  `[추천] <--header-row 그 행, --unit-row를 뺀 전체 명령>`.
- 예측의 레이아웃 확인(`_fit_unit_layout`)이 첫 데이터 행의 숫자 입력이 모두 숫자여야 맞다고 봐서, 첫 값이 `측정불가`인 파일에 추천을 못 하던 문제.
  숫자 입력 칸의 절반 이상이 숫자면 데이터 행, 과반이 글자면 단위·메모 행 (`_first_row_text_share` 하나로 판단).

### 테스트
- `test_layout_matrix_v2917.py`: 실제 모양의 레이아웃 16개(머리글만, 머리글+단위, 제목+머리글+단위, 병합 두 줄 머리글, 메모 행, 데이터 앞 빈 행,
  일부 열만 단위, `℃`, 빈 첫 값, `%` 글자 값, 제목+머리글, 범주가 앞인 글자 타겟 데이터(단위 있음/없음), 첫 행 `NG`/`측정불가`, 날짜 단위, 날짜 머리글만)마다
  맞는 옵션의 학습, 옵션 없는 학습, 단위 있는/없는 모델의 템플릿·`next_commands` 예측이 `VERIFIED`이거나 `[추천]` 1번으로 맞는 결과가 되는지 확인.

## [2.9.16] - 2026-11-02

5차 점검에서 찾은 머리글·단위 행 판단 오류 수정.

### 수정 (학습·예측·표 보간)
- 범주 입력과 글자 타겟이 든 두 줄 머리글 파일에서 첫 데이터 행(`B, M3, kim, 182.8, 중`, 숫자 절반 미만)을 단위 행으로 보고
  `[추천] --header-row 2 --unit-row 3`을 내, 그대로 따르면 `온도(182.8)`, `등급(중)` 같은 이름으로 학습(`VERIFIED`)하던 문제 (2.9.15 회귀).
  단위 행은 숫자가 절반 미만이고, 단위가 든 열마다 바로 아래 칸이 숫자나 빈 칸인 행만 (`_looks_like_unit_row`, `-`·`N/A`는 빈 칸).
  머리글 행 추천(`header_candidates`), 머리글+단위 행 쌍 찾기, `--unit-row` 검사가 모두 이 기준. 주어진 `--unit-row`가 이 기준에 맞지 않으면 `STOP_RETRY`,
  머리글+단위 행 쌍이 없으면 `[추천] <--unit-row를 뺀 전체 명령>`. 제목 행 아래 머리글이 단위 행으로 잡히던 일도 없어짐.

### 수정 (예측)
- 단위 없이 학습한 모델로 머리글 1행 + 단위 2행인 새 파일을 예측하면 `--unit-row 2`로 읽은 이름(`반응온도(°C)`)이 모델과 달라 안내가 없고
  단위 행을 데이터로 예측(`VERIFIED`)하던 문제 (2.9.15 미완). `--unit-row`로 맞지 않으면 단위 행을 건너뛴 `--data-row <단위 행+1>`로 읽어 확인하고
  `STOP_RETRY` + `[추천] <--data-row를 붙인 전체 명령>`. 레이아웃 옵션이 틀린 경우와 머리글 행 안내(제목 행 + 머리글 + 단위 행 파일,
  `--header-row N --data-row N+2`)도 같은 확인 (`_fit_unit_layout`).

### 수정 (표 보간)
- `ml_map --points`: 데이터 파일의 머리글 행이 틀려(제목 행) 조회 파일 열을 찾지 못하면 안내가 없던 문제. `-i`와 같은 `[추천] --header-row N`.

## [2.9.15] - 2026-11-02

4차 점검에서 찾은 학습·예측·보고서·표 보간 오류 수정.

### 수정 (학습)
- 클래스 3개 이상, 10~15행 분류(중복 행을 묶은 경우 포함)가 내부 오류 "n_splits=3 cannot be greater than the number of members in each class"이던 문제.
  `LogisticRegressionCV(cv=3)`의 내부 교차검증이 학습 fold의 작은 클래스에 맞지 않았음. 후보 모델 모두를 점검해 학습 fold(와 마지막 전체 적합)마다
  설정을 맞춤 (`fit_safe_model`): 가장 작은 클래스가 3행 미만이면 내부 `StratifiedKFold(2)`, 1행이면 일반 `LogisticRegression`, KNN 이웃 수는 행 수 이하,
  ElasticNetCV 내부 교차검증은 fold당 2행 이상. 값이 몇 개뿐인 열이 학습 fold에서 모두 비면 입력이 0개가 되던 내부 오류도 `keep_empty_features`로 막음.
  입력 검사를 통과한 작은 데이터(클래스 x 행, 중복)를 매개변수 테스트로 확인.
- `--unit-row M`과 함께 받은 `[추천] --header-row N`을 따르면 단위 행이 데이터 행이 되어 `반응온도(192.58)` 같은 이름으로 학습(`VERIFIED`)하던 문제.
  단위 행 값이 절반 이상 숫자면 `STOP_RETRY` (파일의 머리글+단위 행 쌍을 찾으면 `[추천] <--header-row N --unit-row N+1 전체 명령>`).
  열 이름 머리글 추천은 그 행 아래가 단위 행이면 `--unit-row N+1`을 함께, 실행에 `--unit-row`/`--data-row`가 있는데 단위 행이 없으면 그 옵션을 뺀 전체 명령 (제목 행 파일 포함).
  `resolve_column_names`의 인자 `header_candidates`는 `layout`(`df.attrs`)으로 바뀜. `header_candidates` 항목에 아래 단위 행과 합친 이름이 더해짐.
- xlsx 머리글의 정수 숫자(연도)가 `2023.0`이 되던 문제. `2023`.

### 수정 (예측)
- 머리글 1행 + 단위 2행인 새 파일을 옵션 없는 템플릿으로 예측하면 단위 행을 빈 값으로 채워 예측(`VERIFIED`)하던 문제. 레이아웃 옵션이 없어도 첫 데이터 행을 확인해,
  숫자 입력 칸이 대부분 글자이고 `--unit-row <머리글+1>`로 읽으면 맞으면 `STOP_RETRY` + `[추천] <--unit-row를 붙인 전체 명령>`. `-` 같은 빈 값 표시는 단위로 보지 않음.
- `--unit-row`로 학습한 모델(`반응온도(°C)`)을 머리글 `반응온도`, 단위가 다음 행인 새 파일(제목 행 아래)에 쓰면 머리글 안내가 없고
  `'bar(3)' -> '압축압력(bar)'` 같은 틀린 이름 바꾸기 `[추천]`이 나오던 문제. 아래 단위 행과 합친 이름으로도 찾아 `[추천] <--header-row N --unit-row N+1 전체 명령>`,
  단위 자리에 숫자가 든 열 이름(`bar(3)`)은 오타 후보로 쓰지 않음.

### 수정 (보고서·이력)
- `--history`의 `best_version`이 마지막 기록과 타겟 조합이 같은 기록 중 최고라, 보고서가 v1을 재사용(`made_best: false`)한 뒤 타겟 하나만 다시 학습하면
  그 기록(나쁘거나 일부 타겟뿐)을 가리키던 문제 (2.9.14 회귀, SKILL 4단계가 그 버전으로 예측). `ml_report.py -d`(버전 생략)도 같음.
  둘 다 `task_best_version` 하나를 씀: 최신 최고 버전 기록, 아니면 모든 타겟의 최고 점수가 있는 한 버전(보고서 재사용과 같은 규칙), 나뉘어 있으면
  최고 타겟이 가장 많은 버전(같으면 타겟을 더 많이 가진 버전, 그다음 나중 버전)과 `best_version_note`. 새 키 `best_version_dir`, `best_by_task`도 같은 값
  (`version_dir`, `complete`, `missing_targets`). SKILL 4단계는 `best_version_dir`을 씀.
- 보고서 숫자가 5e-5~1e-3에서 유효숫자 1자리(`0.00012345` -> `0.0001`)이고 1e-3 근처에서 자리 수가 뛰던 문제 (2.9.14 회귀).
  소수 자리 = max(4, 3 - floor(log10 |v|)), 4자리를 넘는 끝 0은 지움 (`0.0001235`, `0.0066`), 1e-5 미만만 지수 형식. 카드·표·학습 범위·90% 범위 모두 같은 함수.
- 엑셀이 없는 PC에서 보고서 결과가 `export_verified: false`인데 `agent_directive`는 `VERIFIED:`이던 문제 (R1). 성능예측표의 `unverified`를
  결과 `unverified`에 `성능예측표: ...`로 넣고 `verified: false`, `UNVERIFIED:`.

### 수정 (표 보간)
- `ml_map`: 흩어진 기록의 "기준으로 주지 않은 열" 경고가 `-i`의 글자 열로 거른 행에서 값이 없는 열까지 나열하던 문제. 거른 행으로 숫자 열을 판단.

## [2.9.14] - 2026-11-02

3차 점검에서 찾은 학습·예측·표 보간·보고서·성능예측표 오류 수정.

### 수정 (학습)
- 분류 + 중복 행: 어떤 클래스의 행이 모두 같으면(서로 다른 행 1개) `StratifiedGroupKFold`가 그 클래스를 한 fold에만 넣어, 이진 분류는 내부 오류(shape mismatch),
  다중 분류는 그 클래스 재현율 0이 말없이 나오던 문제. 그런 클래스가 있거나 묶은 결과 어떤 학습 fold에 빠지는 클래스가 있으면 그 타겟만 일반 `StratifiedKFold`로 검증하고
  중복 경고에 이유를 덧붙임. 학습 fold가 5행보다 작으면 KNN 이웃 수를 줄임.
- 회귀 + 심한 중복: 학습 fold가 1~2행이 되어 ElasticNetCV가 내부 오류(ValueError)를 내던 문제. 입력+타겟이 서로 다른 행이 10개 미만이면
  `STOP_RETRY` "입력+타겟이 다른 행이 N개뿐", 학습 fold가 작으면 ElasticNetCV 내부 교차검증 수를 줄임.
- 열 이름을 못 찾을 때의 `[추천] --header-row N`이 틀리던 문제: 지금 머리글 행을 추천("1행에 있습니다"), 단위를 뗀 비교로 `온도(C)`를 `온도(°C)` 행에서 찾음,
  글자 값이 든 데이터 행('압력' 범주)을 머리글로 봄. 지금 머리글 행은 빼고, 데이터 시작 행보다 위이거나 숫자가 절반 미만인 행만, 열 이름 매칭과 같은 규칙
  (정확히 같거나, 단위 없이 쓴 이름이 단위를 뗀 머리글 하나와 같을 때)으로 찾음. 못 찾으면 비슷한 열 하나는 `[추천] '온도(°C)'`, 여럿은 `[후보]` (학습·예측·표 보간).
  `df.attrs["head_rows"]`는 `header_candidates`로 바뀜.

### 수정 (예측)
- 레이아웃 옵션이 맞지 않는 파일의 `[추천]` 명령이 단위 행이 있는 파일에서 `--unit-row` 없이 나와, 그대로 실행하면 단위 행을 데이터로 예측하던 문제.
  옵션 없이 읽은 첫 데이터 행의 숫자 입력이 숫자인지 확인하고, 아니면 `--unit-row <머리글+1>`로 읽어 맞을 때만 그 옵션을 붙여 추천, 둘 다 안 맞으면 추천하지 않음.
- `-i "온도=180,190"`이 말없이 180190으로 예측(`VERIFIED`)되던 문제. 숫자 입력 값에 쉼표가 있으면 받지 않음: 천 단위 형식이면 `STOP_RETRY` +
  `[추천] <쉼표를 뺀 -i로 바꾼 전체 명령>`(천 단위 구분 기호로 봤다고 알림), 아니면(`1,5`) `[추천]` 없이 `STOP_RETRY`. 글자 입력은 이어 붙인 글자 그대로.
  `--find`의 목표값(`y=1,000`)·`-i`, `ml_map.py`의 `-i`·`--find`도 같은 함수(`check_comma_numbers`)로 같은 `STOP_RETRY`.
- 추천 명령(`_predict_command`)의 `-m`/`-f`/`-o`/`--meta`를 절대 경로로 따옴표 처리 (끝의 `\`나 상대 경로가 Windows에서 명령을 깨지 않게).
- 외삽 경고가 상한 없이 늘던 문제 (입력 200개 -> 400건, 같은 목록이 `drift_warnings`와 `extrapolation_warnings`에 두 번). SKILL.md가 읽는 `extrapolation_warnings`만 남기고
  최대 20건, `extrapolation_warning_count`(전체), `extrapolation_warnings_truncated`. `drift_warnings` 키는 지움 (베타).
- `STOP_RETRY` 안내 문구: "[추천]의 이름으로 1번 다시 실행" -> "[추천]의 이름, 옵션 또는 명령 (python으로 시작하면 전체 명령이니 그대로 1번 실행)".

### 수정 (보고서)
- 모든 타겟의 최고 점수가 이미 한 버전(예: 세 타겟을 한 번에 학습한 v1)에 있는데 `ml_report.py -f`가 같은 조건으로 다시 학습한 "최고" v2를 만들던 문제.
  타겟별 최고 버전이 모두 같은 버전이고 데이터 파일 내용이 그때와 같으면 그 버전을 그대로 씀 (`made_best: false`, 같은 조건을 같은 데이터로 다시 학습하면 같은 모델).
- 숫자 표시가 저장한 값보다 자리를 늘리던 문제 (`0.0066` -> `0.006600`). 소수 4자리로 쓰고, 소수 4자리로 0이 되는 값만 유효숫자 4자리 지수 형식. 카드와 표가 같은 규칙.
- 학습의 중복 행 경고가 보고서에 없던 문제. "2. 데이터와 피처"에 타겟별 중복 행 수와 검증 방식.

### 수정 (표 보간·성능예측표)
- `ml_map`: `-i`의 글자 열로 거른 뒤 값이 없는 구할 열이 "숫자 열이 아닙니다" 오류(틀린 문구), `-y "Y1,Y2"`면 전체 실패, `-y` 없으면 말없이 빠지던 문제.
  숫자 열 판단은 거르기 전 표로 하고, 그 열은 `MultiTable`이 빼고 `warnings` (남는 열이 없을 때만 오류, 오류 문구에 열 이름).
- `ml_map`: 뺀 구할 열이 "기준으로 주지 않은 열"(영향이 섞인 값) 경고에 들어가던 문제. 처음 구할 열 목록으로 판단.
- `MultiTable`: `data_range`·`nearest_rows`가 뺀 열에만 값이 있는 행까지 쓰던 문제. 남은 열 계산에 쓴 행만. `out_of_range`가 열 묶음마다 덮어써지던 문제: 내용이 다르면 `<열>: <내용>`.
- `ml_map --points`: 결과 시트에 `근거`와 `근거_<열>`이 행마다 섞이던 문제. 실행마다 한 번 정함. CSV 조회 파일의 기준 값이 글자로 쓰이던 문제: 숫자 열은 숫자.
- 성능예측표: 타겟이 하나일 때도 타겟 이름의 `,` `;` `=`를 검사해 후보 명령을 만들지 않던 문제 (타겟 하나면 `--inputs`에 타겟 이름을 쓰지 않음).

## [2.9.13] - 2026-11-02

2차 점검에서 찾은 학습·예측·보고서·표 보간·성능예측표 오류 수정.

### 수정 (학습·예측)
- 학습 범위(`feature_stats`)를 소수 4자리로 저장해 작은 입력(1e-6~9e-6)이 최소=최대=중앙값=0이 되던 문제. 예측마다 외삽 경고("학습 최대=0.00 ... +500%"),
  `--find`의 `training_range` [0, 0]·0건, `filled_defaults` 0. 유효숫자 12자리로 저장하고, `--find`의 답·운전 구간·`training_range`·`reachable`·`filled_defaults`는
  소수 6자리(작은 값은 유효숫자 6자리), 외삽 경고의 값은 유효숫자 4자리.
- 교차검증이 그냥 KFold라 완전히 같은 행(입력+타겟)이 학습 fold와 검증 fold에 나뉘어 점수가 부풀던 문제 (무작위 30행 x 4번 복사가 R² 1.0, `goal_met`).
  같은 행은 같은 fold로 묶고(회귀 `GroupKFold`, 분류 `StratifiedGroupKFold`) 결과 `warnings`에 중복 개수. 중복이 없으면 fold는 이전과 같음.
- 회귀 타겟 값이 모두 같으면 R² 1.0과 `goal_met`이 나오던 문제. `STOP_RETRY` "타겟 'y' 값이 모두 같아(5.0) 학습할 수 없습니다" (분류는 이미 클래스 1개 오류).
- 예측 `-i "...,압력=500,000"`이 `000`을 버리던 문제. `=`가 없는 조각은 앞 값에 쉼표로 붙이고(천 단위 구분 기호는 숫자 변환에서 지움), 맨 앞 조각이면 `STOP_RETRY`.
- 학습 레이아웃 옵션(`--header-row`/`--unit-row`/`--data-row`)이 든 예측 `next_commands`를 구조가 다른 새 파일에 쓰면 빠져나갈 방법 없이 `STOP_RETRY`이던 문제.
  옵션으로 읽은 결과에 데이터 행이나 모델 열이 없고 옵션 없이 읽으면 모델 열이 모두 있으면 `[추천] <옵션을 뺀 같은 명령>` (말없이 바꾸지 않음).
- 두 줄(병합) 머리글에서 열 이름을 못 찾으면 힌트가 없던 문제. 위쪽 15행 중 그 이름이 모두 있는 행을 찾아 `[추천] --header-row N` (학습·예측·표 보간).
- 예측 `-i`의 이름 오타 안내가 반대 방향이던 문제 ("'온도' -> [추천] ['온두']"). "-i의 '온두' -> [추천] '온도'"(고칠 이름 -> 모델 이름), 비슷한 이름이 여럿이면 `[후보]`.
- `parse_datetime`이 숫자 20240103.0은 읽고 글자 "20240103.0"은 읽지 못하던 문제. 끝의 `.0`을 떼고 8자리 날짜로 읽음.
- 잘린 인라인 결과(`-i` 6건 이상)에 `full_result_file`이 없던 문제 (R2). `output_file`과 같은 경로.

### 수정 (보고서·표 보간·성능예측표)
- Word 보고서가 작은 타겟의 MAE·표준편차·90% 범위·순위표를 소수 고정 자리(`0.0000`)로 쓰던 문제. 소수 4자리 기본, 작은 값은 유효숫자 4자리(1e-4 아래는 `3.200e-06`).
- `ml_map` 여러 구할 열(`MultiTable`): 모든 열의 근거(`basis`)가 첫 열 기준이던 문제. 결과에 `basis_by_column`, 엑셀은 열마다 쓴 행이 다르면 `근거_<열>` 열.
- `ml_map`: 값이 1개뿐인 구할 열 하나 때문에 전체가 실패하던 문제. 그 열만 빼고 `warnings`, 모든 열이 그럴 때만 오류. 같은 경고가 열 묶음마다 반복되던 것을 한 번으로,
  `data_range`·`nearest_rows`는 첫 열이 있는 행 대신 기준 열이 있는 모든 행으로.
- 성능예측표 `--inputs` 9개 이상의 `[추천]`이 넘은 타겟 하나만 담아, 그대로 실행하면 다른 타겟의 `--inputs`가 사라지던 문제. 모든 타겟의 `--inputs`를 담고(넘은 타겟은 줄인 안),
  넘은 타겟이 여럿이면 한 번에 모두 줄임.
- 성능예측표: 이름에 `,` `;` `=`가 든 타겟·입력은 `--inputs`로 쓸 수 없는데 깨진 명령을 만들던 문제. 그런 후보는 `command` 대신 `note`, `[추천]` 대신 줄인 입력 목록과 안내 (README 제한 사항).

## [2.9.12] - 2026-11-01

전체 검토에서 찾은 성능예측표·보고서·표 보간 오류 수정.

### 수정
- 성능예측표: 표의 예측값을 소수 6자리로 반올림하고 표시 형식·90% 범위 글자를 `0.00`으로 고정해, 작은 타겟(2e-6~8e-6, 0.005 근처)은 계산기 값이 틀리거나 0.00인데
  일치도는 100%로 나오던 문제. 표에는 유효숫자 12자리로 넣고 일치도도 저장한 값으로 계산, 표시 형식은 타겟 크기에 맞춰 유효숫자 약 4자리
  (기본 `0.00`, 작으면 소수 자리를 늘리고 1e-5 아래는 `0.000E+00`). 학습의 90% 범위 잔차와 MAE·RMSE도 소수 자리 대신 유효숫자로 반올림.
- 성능예측표: "참고용 (실측 없는 영역)" 기준이 고정 0.15라, 입력 6개 이상(표가 성김)이면 학습 행의 75~100%가 참고용으로 나오던 문제.
  기준 = 0.15 + 0.5 × 표 칸 대각선(√k/(n-1)). 학습 행은 거의 표시되지 않고 실측이 없는 영역은 그대로 표시됨. 결과 `uncovered_dist`, 정보 시트.
- 성능예측표: 입력 13개 이상이면 엑셀 열 한도를 넘거나 메모리가 부족해 JSON 없이 끝나던 문제. 사용자 결정으로 입력은 타겟마다 최대 8개.
  자동 최적화는 8개 이하 안 중에서 고르고(9개 이상 후보는 표 없이 점수만), `--inputs` 9개 이상은 8개 이하로 줄여 다시 학습한 안을 `STOP_RETRY` + `[추천]`, `--keep` 8개 초과는 `STOP_RETRY`.
- 표 보간(`ml_map`, `-y` 생략): 다른 숫자 열에 빈칸이 있는 행을 모든 열 계산에서 빼 값이 달라지던 문제 (A=50에서 250 대신 252.5).
  기준 열이 빈 행만 빼고, 구할 열마다 그 열에 값이 있는 행으로 따로 계산 (`MultiTable`).
- 성능예측표 `--keep`/`--inputs`: 이름을 모든 타겟의 입력에서 찾아, 타겟마다 입력이 다른 최고 버전에서 실패하던 문제. 그 입력이 있는 타겟에만 적용하고 어느 타겟에도 없을 때만 오류.
- `ml_export`, `ml_map`에 `--overwrite` (같은 이름 파일이 있을 때 복구 안내가 말하던 옵션이 없었음). 보고서 `--overwrite`는 함께 만드는 성능예측표에도 적용.
- 보고서 결과가 성능예측표 `warnings`를 버리던 문제: `export_warnings`로 넘기고 `agent_directive`에서 알리라고 함.
- 성능예측표 `next_commands`를 8개에서 말없이 잘랐고, 다른 안 명령이 `-y <한 타겟>`이라 그 타겟만 든 파일이 만들어지던 문제.
  후보마다 `command`(다른 타겟은 지금 입력 유지, 타겟별 `--inputs "<타겟>=a,b;<타겟2>=c"`), `next_commands_count`, `truncated`.
- 성능예측표 시트 이름 중복 검사가 대소문자를 구분하던 문제 (엑셀은 구분하지 않음).
- 보고서 `-t`로 타겟을 골라도 성능예측표는 모든 타겟으로 만들던 문제.
- 예측 결과(`pred_`, 90% 범위, 실측 비교값, 반대 찾기 확인값)를 소수 4자리로 반올림해 작은 타겟이 0으로 나오던 문제.
  소수 4자리를 기본으로 하되 작은 값은 유효숫자 4자리까지 남김 (`round_pred`).
- pyflakes 정리 (보고서의 쓰지 않는 변수).

## [2.9.11] - 2026-11-01

전체 검토에서 찾은 학습·예측 오류 수정.

### 수정
- 학습 `--dtypes "일자:datetime"`이 예측과 다른 날짜 해석(`pd.to_datetime`)을 써서, `20240103` 같은 8자리 날짜로 학습하면 예측 때 날짜가 모두 빈 값으로 대치되던 문제.
  학습(강제 지정, 자동 판별, `datetime_stats`)과 예측이 `ml_predict.py` PART C의 `parse_datetime` 하나를 쓰고, 이 함수가 `YYYYMMDD` 8자리(숫자 또는 글자)도 읽음.
- 예측 `-i "k=v"`가 값을 모두 숫자로 바꿔 범주 코드 `007`이 `7`이 되던 문제 (학습 때 없던 범주 경고, `.json` 입력과 다른 예측). 값은 글자로 받고 숫자 입력만 나중에 숫자로 바꿈. `--find`의 `-i`도 같음.
- 이름이 있는데 값이 모두 빈 열을 읽을 때 버리던 문제. 학습 `-x 온도1`(빈 열)이 `[추천] ['온도2']`를 내고, 예측은 빈 값 채우기 대신 "필수 피처 누락"이었음.
  이제 머리글이 빈 열(`col_<번호>`)만 버리고, 이름 있는 빈 열은 학습 "값이 하나도 없습니다" 오류, 예측은 기준값 채우기와 `missing_input_warning`.
- 학습을 `--header-row`/`--unit-row`/`--data-row`로 했을 때 `next_commands`의 예측 명령에 이 옵션이 없어, 같은 구조의 새 파일에서 단위 행을 데이터로 예측하던 문제 (`--best`와 보고서 `-f` 결과의 예측 명령 포함, 보고서는 `-f`의 옵션 또는 이력 값).
- 예측 `-i`에 6건 이상을 주면 평균·최소·최대만 주고 결과 파일이 없던 문제. `-o`가 없으면 결과 CSV를 저장(`.json` 입력이면 그 옆, 아니면 `<데이터>_ml_data\inline_pred_<시각>.csv`)하고 `output_file`, `truncated: true`, `count`, `shown`.
  5건 이하 인라인 결과는 `truncated: false`, `count`.
- 같은 머리글 이름 바꾸기가 다시 중복을 만들던 문제 (`a, a, a_1` -> `a, a_1, a_1`). 붙인 이름이 이미 있으면 다음 번호.
- 예측 `-o`에 `.csv`/`.xlsx`가 아닌 확장자를 주면 저장 후 `UNVERIFIED`이던 문제. 시작할 때 `STOP_RETRY` + `[추천] <이름>.csv`.
- 학습 `next_commands`가 이름의 큰따옴표를 작은따옴표로 바꿔 열 이름이 달라지던 문제. 큰따옴표가 든 피처·타겟 목록은 `next_features_*.txt`/`next_targets.txt` 파일로 넘기고,
  `--eval`·`--dtypes`·`--positive-class`에 큰따옴표가 필요한 재학습 명령은 넣지 않음.
- `--history` 결과에 `verified: true` (R1).

### 삭제 (베타, 이전 버전 호환)
- 예측의 `feature_schema` 없는 모델용 분기와 2.0 이전 모델 주석. README의 예전 사본·1.x 모델 설명을 현재 동작으로 고침.
- 보고서의 "2.2.0 이상" 안내를 예측과 같은 "2.9.1 이상으로 다시 학습"으로.

## [2.9.10] - 2026-10-31

### 수정
- 성능예측표: 상태·판단 칸의 조건부 서식 색이 나오지 않던 문제. Excel은 조건부 서식의 단색 채우기에서 `bgColor`를 읽습니다. 정상은 초록, 범위 밖·참고용은 주황 굵은 글자로 표시합니다 (계산기, 일괄계산).

## [2.9.9] - 2026-10-31

집 PC(Office 2007) 실기에서 찾은 문제를 고쳤습니다.

### 수정
- 보고서: Word 2007에서 열리지 않던 문제. 표 칸 여백에 Word 2010 이후 이름(`w:start`/`w:end`)을 쓰고 `w:tcPr`·`w:tblPr`·`w:pPr` 하위 요소 순서가 스키마와 달랐습니다. `left`/`right`와 스키마 순서로 씁니다.
- 예측: 입력 4개 이상 모델의 실측 비교값이 입력 행마다 학습 데이터 열을 다시 변환해 2만 행 예측이 15분을 넘던 문제. 범주별로 한 번만 변환합니다. 계산 결과는 같습니다.

## [2.9.8] - 2026-10-30

### 변경 (R1 사후 확인)
- 예측: 결과 파일(CSV/xlsx)을 저장한 뒤 다시 읽어 행·열 수와 열 이름을 비교. 같으면 `VERIFIED:` + `verified: true`, 다르거나 못 읽으면 `UNVERIFIED:` + `unverified`.
- 보고서: 저장한 Word 파일을 다시 열어 버전 표기와 표 수(타겟 수 이상)를 확인. `verified`/`unverified`, `-f`는 보고서 전체 기준.

## [2.9.7] - 2026-10-30

베타 정책(AGENTS.md: 이전 버전 호환 불필요)에 따라 예전 버전 호환 코드를 지웠습니다. 이전 버전 폴더는 같은 데이터로 다시 학습해서 씁니다.

### 삭제
- 2.9.1 이전 모델(모델 정보에 `input_columns_required`나 `feature_schema`가 없음)과 정보 파일이 없는 모델: 대체 예측 경로 대신 `STOP_RETRY` "다시 학습".
- 2.7.0 이전 학습 데이터 사본(`원본행` 열 없음)의 줄 번호 대체: 실측 비교값·`--find` 실측 근거를 계산하지 않고 다시 학습 안내.
- 결과·원장의 예전 회귀 전용 키 `mean_val_r2`, `delta_r2`, `--history`의 `best_mean_val_r2` (`mean_val_score`, `delta_score`, `best_mean_val_score` 사용).
- `ml_report.py`: 2.1 소견 이름(`slot1_residual` 등), `--history`/`-H` 별칭(`--data-dir`/`-d`), `-d`에 history.json 파일 경로 받기.
- 성능예측표가 만드는 `export/model_<타겟>.json`에도 `input_columns_required`.

## [2.9.6] - 2026-10-30

전체 검토에서 찾은 오류 수정.

### 수정
- 성능예측표: 계산기의 학습 범위·"범위 밖" 판단·기본값이 학습에서 뺀 행(자동 이상치, 직접 표시)까지 포함하던 문제. 이제 표가 실제로 덮는 범위
  (출력 여러 개면 그 입력을 쓰는 모든 표의 공통 범위)라 "정상"인데 몰래 외삽되는 일이 없음. 정보 시트의 "뺀 입력" 범위도 학습에 쓴 행 기준.
- 성능예측표: 값이 하나뿐인 입력(예: pH가 모두 7.0)이 있으면 만들지 못하던 문제. 그 입력은 표에서 빼고 `warnings`로 알림
  (`--inputs`에 그런 입력만 있으면 `[추천]`과 함께 STOP_RETRY).
- 성능예측표: 입력 8개 이상이면 보간 수식이 Excel 인수 한도(255개)를 넘던 문제. 꼭짓점 항을 연속 칸에 두고 범위 합계(`SUM(M10:T10)`)로.
- 성능예측표: 엑셀 재계산 확인에서 파일 열기가 실패하면 숨은 EXCEL.EXE가 남던 문제 (항상 종료).
- 성능예측표: 파생변수의 원본 입력을 글자 포함으로 찾던 문제 ('온도'가 '반응온도' 안에서 찾아짐). 이름 전체로 비교하고, 다른 파생변수로 만든 파생변수는 원본이 빠지면 함께 뺌.
- `제외` 열의 `0`, `FALSE`는 표시 없음으로 봄 (0/1 표시를 쓰면 0인 행을 모두 잃던 문제).
- 가까운 실측 5행(ml_map, `--find`)의 숫자가 CSV에서 글자로 나오던 문제.
- `--find`의 `-i` 고정 입력이 숫자가 아니면 영어 오류 원문 대신 어느 입력인지 알려 주는 STOP_RETRY.
- 내부 오류(ValueError, KeyError)를 입력 오류(STOP_RETRY)로 분류하던 문제: 이제 `ERROR:`와 오류 로그 traceback (R5).
- ml_map 반대 찾기의 `truncated`는 항상 있고, 잘렸으면 `shown`도.
- 학습 후 안내의 보고서 명령 설명이 `--title`을 말하지만 명령에는 없던 불일치 (제목은 데이터 설명에서 자동).

### 문서
- requirements.txt에 matplotlib (보고서 그림). README: 성능예측표·excluded_rows.csv 폴더 구조, 보고서 `--title`, 결과 키 `export_*`. SKILL.md 결과 처리 표에 `[후보]`.

## [2.9.5] - 2026-10-30

### 변경 (실기 테스트)
- 학습 제외 행 목록을 `vX/excluded_rows.csv`로 저장 (이전 `.xlsx`). 다른 산출물처럼 CSV, 한국어 엑셀에서 바로 열리는 CP949. 보고서도 CSV를 읽음.
- SKILL.md 규칙 9: 도구 코드(.py)는 사용자가 코드를 열어 보라고 직접 요청할 때만 읽음 (실기 테스트에서 허용으로 결정).

## [2.9.4] - 2026-10-30

### 기능
- 성능예측표의 모델 대비 일치도: 학습 범위 안 무작위 조건 2,000개에서 표 값과 모델 예측을 비교해 평균 일치도(%), 95% 조건의 차이, 최대 차이,
  모델 90% 범위 안 비율. 출력이 0 근처를 지나면 출력 범위 대비로 계산. 계산기 하단 안내, 정확도·정보 시트, `ml_export` 결과 `agreement`.
- 보고서: 성능예측표를 Word보다 먼저 만들어, 요약 아래에 성능예측표 파일명과 타겟별 남긴 입력·일치도 한 줄. 결과 `export.<타겟>.agreement`.

## [2.9.3] - 2026-10-30

### 수정 (실기 테스트)
- 성능예측표의 `표_<타겟>` 시트(모델로 계산한 격자 데이터)를 완전히 숨김(veryHidden). 계산 수식은 그대로 이 시트를 참조. 사용법 시트 문구 수정.

## [2.9.2] - 2026-10-30

사용자 요청 (실기 테스트): 보고서와 성능예측표는 세트. 보고서를 요청하면 성능예측표가 함께 만들어지고, 같은 폴더(history가 있는
`<데이터>_ml_data`)에 저장.

### 기능
- `ml_report.py`가 회귀 버전이면 `ml_export.py`(입력 최적화 + 엑셀)를 함께 실행. 결과 `export_file`, `export`(타겟별 남긴 입력 등),
  `export_verified`/`export_unverified`. 성능예측표 실패는 보고서를 막지 않고 `export_note`. 새 인자 `--title`(성능예측표 제목). 분류는 만들지 않음.
- `ml_export.py` 기본 저장 위치를 보고서와 같은 `<데이터>_ml_data` 폴더로, 이름을 `성능예측표_<데이터>_<버전>.xlsx`로 (보고서 `ML결과보고서_<데이터>_<버전>.docx`와 짝).
- 학습 결과 안내: 다음 명령은 "보고서와 성능예측표 생성" 하나로 합침 (별도 `ml_export.py` 명령 없앰). `ml_export.py`는 `--keep`/`--inputs`로 다시 만들 때 사용.
- 개발: 테스트에서는 `ML_KIT_SKIP_EXPORT=1`(conftest)로 보고서마다 성능예측표를 만들지 않음 (`test_report_set_v292.py`만 켬).

## [2.9.1] - 2026-10-30

### 수정 (실기 테스트 2026-10-30)
- 보고서 판정 기준 아래 점수 용어 설명을 한 줄씩 나눔 ("점수 용어" 다음에 학습/검증/제외 행 미포함·포함을 각각 한 줄).
- 학습 결과 안내에 성능예측표 추가 (회귀만, 2.9.2에서 보고서 명령으로 합침).
- 버그: 파생변수의 원본 열이 `-x`에 없으면(예: `--eval "A상한 = clip(A, -1e12, 250)"`에 A를 `-x`에서 뺌) 학습은 되지만 예측이 모두 실패.
  모델 정보에 사용자가 넣어야 할 원본 입력 `input_columns_required`(자기 피처 + 파생변수 원본 열)를 기록하고, 원본 열을 `model_dataset.csv`와
  `feature_stats`에 넣음. 예측(별칭 맞추기, 빠진 원본 입력은 분명한 오류), 실측 비교값, 반대 찾기(`--vary`에 원본 열), 성능예측표가 이 목록을 씀.
  이전 모델은 지금처럼 피처 중 파생변수가 아닌 것을 원본 입력으로 봄. 회귀 테스트 추가.

## [2.9.0] - 2026-10-29

사용자 요청: 성능 예측 프로그램을 모델+ml_predict 대신 엑셀만으로 배포. 모델로 운용 영역의 표를 만들고 엑셀 기능으로 맵핑.
입력 수(표 크기)는 입력 조합 최적화로 줄이고, 꼭 포함할 입력은 지정. 일반 사용자용 화면, 데이터는 수정 불가(암호·숨김 없음).

### 기능 (추가)
- 새 진입점 `ml_export.py -m <버전> [--keep] [--inputs] [--title] [-y] [--save]`: 원본 숫자 입력을 다시 학습하며 하나씩 줄여
  추천 조합(최고 검증 R² - 0.005 이내에서 입력 최소)을 고르고, 격자 표(최대 20만 행)와 엑셀 계산기를 만듦.
- 엑셀: 계산기(노란 입력 칸, 예측값, 90% 범위, 학습 범위 밖·실측 없는 영역 표시), 일괄계산 500행, 사용법, 표, 정확도, 정보.
  수식은 INDEX/MATCH만(엑셀 2016+), 매크로 없음. 모든 시트 보호(암호 없음)와 구조 잠금, 노란 칸만 입력 가능.
- 확인: 파일 다시 읽기(시트, 잠금, 입력 칸, 표), 수식 계산을 파이썬으로 재현해 비교, Windows+엑셀이면 엑셀로 계산해 비교.
- 출력이 여러 개면 출력마다 따로 입력을 최적화하고(학습 -y 순서 유지), 계산기는 모든 출력의 입력을 합쳐 받으며, 출력은 자기 입력이 다 차면 계산.
- 실험 근거 (CLI-dev 기록): 직선형 모델은 표 보간 오차 0, 비선형 입력 2개는 표가 오히려 정확(트리 계단 완화),
  입력 4개는 축 16점(6.5만 행)이면 모델과 같은 수준. 입력 6개 -> 4개로 줄여도 검증 R² 0.980 -> 0.980.

## [2.8.0] - 2026-10-29

사용자 요청: 모델 기준 반대 찾기 1단계 (회귀 타겟, 입력 1개), 운전 구간도 함께.

### 기능 (추가)
- `ml_predict.py --find "<타겟>=<값>" --vary <입력> [-i 고정 입력]`: 학습 범위를 2,000단계로 훑어 목표값을 지나는 답을 모두 찾음.
  가까운 답(범위의 3%)은 구간으로 묶음, 운전 구간(목표가 90% 예측 오차 안), 도달 가능한 범위, 학습 데이터 기준 답(숫자 입력 3개 이하),
  가까운 실측 5행. 고정하지 않은 숫자 입력은 학습 중앙값, 글자 입력은 [후보]로 묻기. 답 주변을 다시 예측해 확인.
- `_crossings`를 공통 부분으로 옮겨 `ml_map.py`와 함께 씀.

## [2.7.0] - 2026-10-29

사용자 요청: ml_map 보간값이 예측할 때 비교용으로 같이 나오면 좋겠다 (도구를 따로 부르지 않고).

### 구조
- 표 보간 계산(`Table`, `nearest_rows`)을 `ml_predict.py` 공통 부분(C8)으로 옮기고 `ml_map.py`가 가져다 씀 (동작 불변, 별도 커밋). 단독 배포 유지.

### 기능 (추가)
- 회귀 예측마다 실측 비교값(`map_check`): 버전 폴더 `model_dataset.csv`로 같은 조건을 보간 (숫자 입력 3개 이하: 표 보간,
  4개 이상: 가까운 실측 5행 거리 가중평균). 파일 예측은 `map_<타겟>`, `map_diff_<타겟>`, `map_방식_<타겟>` 열.
- `map_check_warning`: 비교값이 90% 예측 범위 밖 (외삽값, 먼 이웃 가중평균은 제외). `map_check_notes`: 비교값을 못 낸 이유.
- 학습: `model_dataset.csv`에 `원본행`, `학습제외_<타겟>` 열 추가 (비교 근거의 엑셀 행 번호, 학습 제외 행 빼기).

## [2.6.0] - 2026-10-29

사용자 요청: 예측 결과를 검증하려고 전통적인 맵핑(위아래 가까운 값의 비율로 사잇값)이 필요하다. 반대 찾기도 필요.
엑셀 수식(office-cli)은 흩어진 기록에서 끝 두 점 외삽이 터지고 COM 때문에 느려, 파일을 직접 읽는 ml-kit 도구로 결정.

### 기능 (추가)
- 새 진입점 `ml_map.py`: 사잇값(`-i`), 반대 찾기(`--find`, `--vary`), 여러 점(`--points`), 결과 엑셀(`--save`).
  격자는 기준마다 위아래 비율 보간과 끝 구간 연장, 흩어진 기록은 위아래 2점(기준 1개) 또는 둘러싼 실측 3~4점(기준 2~3개) 보간과
  전체 추세선 외삽. 답마다 쓴 원본 행과 비율(`basis`), 가까운 실측 5행, 범위 밖 정도(데이터 폭 대비), 외삽 경고.
- 반대 찾기는 답을 모두 돌려주고 보간식에 다시 넣어 확인. 0건이면 도달 가능한 범위.
- '제외' 열 표시 행은 쓰지 않음. 데이터 파일은 고치지 않음.
- `requirements.txt`에 scipy (표준 환경 1.18.1, scikit-learn 의존성으로 이미 설치).
- SKILL.md 6단계, README 6-B, 사용 가이드 장 추가.

## [2.5.0] - 2026-10-29

사용자 요청: 엑셀에서 이상치를 눈으로 보고 지운 뒤 학습하다 보니 보고되는 R²가 실제보다 좋게 보인다. 이상치를 자동으로 찾고,
이상치를 포함한 성능도 볼 수 있어야 한다. 열 이름은 `제외`로 고정, 인자는 늘리지 않는다, 분류도 포함.

### 기능 (추가, 옵션 없음)
- 학습 제외 행: 데이터의 `제외` 열(빈칸: 자동 검사, `포함`: 자동 제외 안 함, 그 밖의 글자: 직접 제외)과 자동 이상치
  (로버스트 z > 3.5, 회귀는 타겟+수치 피처, 분류는 수치 피처)를 학습에서 뺌. 데이터 파일은 고치지 않음.
  후보가 10% 초과, 학습 행 10개 미만, 분류 클래스가 2행 미만이 되는 경우는 자동 제외하지 않음.
- 제외 행 포함 검증 점수: 학습 제외 행을 최종 모델로 예측해 교차검증 예측과 합친 점수(`all_rows_score`). 판정은 그대로 제외 행 미포함 검증 점수 기준.
- `vX/excluded_rows.xlsx`: 뺀 행 목록(원본 엑셀 행 번호, 구분, 이유, 실제값, 예측값).
- 결과 JSON `exclusions`, `excluded_rows_file`, 원장 `exclusions`, 모델 json `excluded`. 공통 로더가 원본 행 번호(`attrs["source_rows"]`)를 기록.
- 보고서: 요약 표 "검증 R² (제외 행 포함)" 열, 데이터 장 제외 규칙, 타겟 장 한 줄, 그림에 제외 행 주황색 X.

### 수정
- 점수 용어 통일: 학습 R²(배운 행을 다시 맞힌 점수, 과적합 확인용) / 검증 R²(배우지 않은 행을 맞힌 점수, 판정 기준).
  학습 제외 행이 있으면 검증 R²를 "제외 행 미포함"(판정)과 "제외 행 포함"으로 나눠 적음. 분류는 Macro-F1.
  보고서의 "교차검증 R²", "검증 점수", "Macro-F1", "학습 F1", "R² =" 표기와 학습 결과 preview의 "Val R2"를 이 용어로 바꿈.
- 학습 범위(외삽 경고 기준)와 타겟 통계를 학습에 쓴 행으로만 계산 (이상치가 학습 범위를 넓히지 않게).

### 호환
- CLI 인자 변경 없음. JSON 키는 추가만. `제외` 열이 없는 데이터도 자동 이상치 검사는 동작하므로, 이상치가 있던 데이터는
  같은 조건으로 다시 학습하면 검증 점수가 달라질 수 있음 (제외 행 포함 검증 점수로 이전과 비교 가능).

## [2.4.2] - 2026-10-29

### 구조
- 폴더 정리: `CLI/ml-kit/`에는 실행에 필요한 파일과 `SKILL.md`만 남김 (`ml_train.py`, `ml_predict.py`, `ml_report.py`).
  `README.md`, `CHANGELOG.md`, `requirements.txt`, `tests/`, `tools/`는 `CLI-dev/ml-kit/`로 옮김.
  공통 계약 `CLI/README.md`와 `run_all_tests.py/.bat`도 `CLI-dev/`로.
- 테스트와 실기 러너는 옆 폴더 `../CLI/ml-kit`의 도구를 시험. 사내 PC에서는 실기 테스트 때만 `CLI-dev`를 `C:\CLI-dev`로 `C:\CLI`와 나란히 둠.
- CLI 인자, JSON 키, 도구 동작은 그대로.

### 수정
- `-h/--help`: 도움말 텍스트는 stderr, stdout은 JSON 1줄(`command: "help"`, `skill_file`, SKILL.md를 쓰라는 안내). 이전에는 도움말이 stdout으로 나가 마지막 줄 JSON이 없었음 (남은 할 일에서 이동).
- SKILL.md 규칙 9: 도구 코드(.py)를 읽지 않고, 방법을 못 찾으면 사용자에게 알리고 멈춤 (BEAM 에이전트가 `-h` 후 코드를 읽은 사례).
- openpyxl 미설치 안내 문구: `requirements.txt를 설치` -> `python -m pip install openpyxl`로 설치 (requirements.txt가 도구 폴더에 없음).

## [2.4.1] - 2026-10-29

사용자 요청: 보고서 내용이 너무 많고 빈 공백이 많다. 핵심만 간결하게, 파생변수 수식은 눈에 잘 들어오게.

### 기능 (보고서 구성 변경, CLI 인자·JSON 키는 그대로)
- `ml_report.py` 문서를 간결하게 다시 구성 (샘플 기준 회귀 4쪽, 분류 3쪽).
  - 표지·목차·쪽 나눔 없앰. 제목 블록 아래 바로 1. 요약.
  - 5. 버전 이력, 6. 운영 가이드, 부록 B(산출물·실행 환경) 없앰. 부록 A 상관관계는 2.2로 옮김.
  - 2.1. 파생변수: `--eval` 계산식을 원래 컬럼 이름으로 `압력비 = 압축압력(bar) ÷ 촉매비율(wt%)`처럼
    ×, ÷, ², ³, √ 기호로 강조 표시하고 사용 타겟별 기여도·순위를 함께 표시. 피처 표에서도 유형 "파생"을 강조.
  - 4. 타겟별 결과: 타겟마다 판정 한 줄, 지표 카드, 성능 그림과 피처 기여도를 나란히 놓은 그림 1장(파생변수는 주황색),
    상위 3개 모델 표, 소견 1개. 잔차 진단·타겟 분포·ROC 그림과 예측 구간 문장은 뺌.
  - 표 캡션과 짧은 표(15행 이하)가 쪽 경계에서 갈라지지 않게 함.
- `--slots`: 모든 키를 그대로 받음. 없어진 장의 `lineage`/`operation`은 요약 소견 끝에 붙임.

### 문서
- SKILL.md 보고 규칙: 파일 경로를 `file:///` 링크로 바꾸지 않고 결과 JSON 값을 그대로 1번만 쓰게 함.
  긴 대화에서 26B 모델이 한글 경로를 링크로 변환하다가 같은 조각(`l'al l'al ...`)을 끝없이 반복한 사례 대응.

## [2.4.0] - 2026-10-28

실제 사용(회사 PC, OpenCode + 26B Gemma)에서 입력 3개·회귀 타겟 6개 데이터를 LLM이 타겟 하나씩, 조건을 바꿔 가며 학습해
v1~v7이 타겟별로 흩어졌고 보고서도 버전마다 따로 나왔습니다. 사용자는 항상 모든 타겟이 들어간 보고서 1개를 원하고,
LLM이 타겟·버전을 고르기를 원하지 않습니다. 사용자와 정한 설계("make_best"):
타겟마다 검증 점수가 가장 좋았던 기록의 학습 조건을 가져와 **현재 데이터 파일**로 모든 타겟을 함께 다시 학습하고,
평범한 새 버전 1개(태그 `best`)에 담은 뒤 그 버전으로 보고서를 만듭니다.

### 기능 (추가만)
- `ml_train.py -f <데이터> --best [--task regression|classification]`: 최고 버전 만들기.
  - 타겟별 선택: 회귀 `val_r2`, 분류 `val_f1_macro`가 가장 높은 기록, 같은 점수면 나중 기록. 버전 폴더가 없거나 불완전한 기록은 제외.
    `--overwrite`로 같은 버전 기록이 2개면 마지막 기록만 봄.
  - 학습 조건(피처, `--eval` 수식, `--dtypes`, 양성 클래스)은 타겟마다 다를 수 있고, 모델 종류는 평소처럼 5개 모델 경쟁으로 다시 고름.
  - 작업 유형마다 버전 1개 (회귀와 분류가 섞이면 2개, 결과에 명시). `--task`로 하나만.
  - 헤더/단위/데이터 행은 주지 않으면 그 유형의 마지막 기록 값. 이력이 없으면 `STOP_RETRY`(먼저 학습).
  - 점수 하락 경고는 하지 않고, 타겟별 원래 점수(출처 버전)와 이번 점수를 결과 JSON과 보고서 표에 함께 적음.
- `ml_report.py -f <데이터> [--task ...] [--header-row ...]`: 작업 유형마다 최고 버전으로 보고서 1개.
  최고 버전이 최신이면(그 뒤에 같은 유형의 학습 기록 없음, 폴더 온전, 데이터 파일 SHA-256 동일) 그대로 쓰고, 아니면 같은 프로세스에서 먼저 만듦.
  `-m`, `-d` 동작은 그대로.
- 결과 JSON 새 키: `made_best`, `best_version_dir`, `best_versions`, `best_targets`(`{타겟: {task, source_version, source_score, new_score}}`),
  보고서 `-f`는 `report_count`, `best_reports`도. 기존 키(`report_file`, `version`, `best_version` 등)는 첫 번째 보고서 기준으로 뜻이 같음.
  (`reports`는 2.2.0에서 없앤 옛 키라 다른 뜻으로 다시 쓰지 않고 `best_reports`로 이름 지음.)
- 원장 새 키: 모든 학습 `data_fingerprint`(`size`, `mtime`, `sha256`), 지정 시 `dtypes`, `positive_class`.
  최고 버전 기록은 `best: true`, `best_source`(`{타겟: {version, score[, selected_version]}}`), `target_recipes`.
  이전 최고 버전의 조건이 다시 뽑히면 `best_source`는 원래(최고 버전이 아닌) 버전을 가리켜 출처를 끝까지 추적.
  `model_<타겟>.json`에도 `best_source`. `--history`의 `records`에 `best: true`.
- 보고서(최고 버전): 표지 "학습 방식", 요약 "타겟별 학습 조건 출처" 표, 타겟마다 "학습 조건 출처: vK (당시 점수 x.xxxx), 이번 점수 y.yyyy"와 입력 피처,
  버전 이력에 "버전별 타겟 점수" 표와 타겟별 추이 그림. 입력 피처 구성과 예측 필수 입력은 모든 타겟의 합집합 (일반 버전은 이전과 같음).
- 학습 목표 달성 시 `next_commands`의 보고서 명령을 `ml_report.py -m <vN>`에서 `ml_report.py -f <데이터>`로 바꿈.

### 구조
- `ml_train.py`의 학습 본체를 `train_version(args, task, groups, best_info)`로 분리. 그룹 = 학습 조건 1개(피처, 수식, dtypes, 양성 클래스)와 그 타겟들.
  일반 학습은 그룹 1개라 산출물이 이전과 같고, `--best`는 조건마다 그룹을 만들어 한 스테이징 폴더(한 `TransactionGuard`)에 모든 타겟을 학습.
  `correlation.csv`, `model_dataset.csv`, 원장 `features_used`·`feature_types`는 그룹들의 합집합(같은 이름 컬럼은 첫 그룹 값).
- `ml_report.py -f`는 `ml_train`을 그때만 import (보고서 `-m`/`-d`는 scikit-learn을 불러오지 않음).

### 스킬
- 5단계 보고서를 `ml_report.py -f "<데이터파일>"` 한 줄로. LLM은 타겟·버전을 고르지 않음. `-m` 안내 제거 (도구에서는 계속 동작).
- 4단계 예측은 보고서를 만들었으면 `best_version_dir`, 아니면 마지막 학습의 `version_dir`.
- 2단계: 여러 타겟을 `-y`에 함께 넣는 것은 권장으로 두고, 따로 학습해도 보고서가 합친다고 안내. `split_targets_detected` 안내 삭제(기능 없음).

### 테스트
- `tests/test_best_v240.py`: 단일 타겟 학습 4회(한 타겟은 최고 점수가 최신 기록이 아님) -> `--best`가 타겟 3개 버전 1개, 타겟별 피처·수식, `best_source`, 새 점수;
  합친 피처로 `ml_predict -i` 예측; `--history`; `ml_report -f` 재사용(`made_best: false`) -> 새 학습 후 재생성 -> 데이터 변경 후 재생성;
  회귀+분류 혼합 시 버전·보고서 2개와 `--task`; 이력 없음 `STOP_RETRY`; `-m` 결과에 새 키 없음.

### 호환
- 기존 CLI 인자·JSON 키 삭제나 뜻 변경 없음. `ml_train.py --task` 기본값을 `None`으로 바꿨지만 학습에서는 그대로 `regression`.
- 산출물 레이아웃(`<stem>_ml_data/history.jsonl`, `vX/`)은 그대로. 2.4.0 이전 기록도 `--best` 후보 (`dtypes`가 기록되지 않아 자동 판별로 다시 학습).

## [2.3.2] - 2026-10-28

회사 PC 2차 실기(2.3.1, 실기 M1~M14 모두 통과) 결과 반영. 오프라인 테스트 3건 실패 수정.

### 스킬
- 도구 위치를 `C:\CLI\ml-kit`으로 고정: SKILL.md 명령 경로를 `<SKILL_DIR>` 대신 실제 경로로 적음 (스킬 폴더에는 연결 또는 SKILL.md만).
- SKILL.md를 PowerShell 기준으로 정리: 전역 지침의 인코딩 줄 사용, 시간 제한 600000, 실행 전에 물어볼 값(파일, 타겟, 피처), 컬럼 이름 확인 방법, `version_dir`을 모를 때 `--history`로 찾기, 파일은 파일 쓰기 도구로 UTF-8 저장.

### 수정
- 회사 PC의 python-docx 템플릿 폴더(`templates/default-docx-template`)에서 `docProps/thumbnail.jpeg`가 빠져 있어
  python-docx가 폴더를 열지 못했습니다(FileNotFoundError). 보고서는 그때 `default.docx`로 넘어가 성공했지만, 폴더 템플릿을 쓰는 테스트가 실패했습니다.
  이제 폴더를 메모리 속 .docx로 묶으면서 빠진 파일을 가리키는 관계를 빼고 엽니다 (`docx_from_template_dir`).
  빠진 파일이 있으면 `recovered`에 "Word 내장 템플릿 폴더에 빠진 파일(...) -> 제외하고 사용"을 남깁니다.
- 참고(회사 PC 진단): `default.docx`는 스타일 4개짜리 2716바이트 문서, `pptx/templates/default.pptx`는 파일 없음.

### 테스트
- 폴더 템플릿에서 thumbnail.jpeg를 지운 상황 재현 테스트 추가. 템플릿 손상 테스트는 폴더 템플릿 상태와 상관없이 동작.

## [2.3.1] - 2026-10-28

회사 PC(Windows 11, WinPython 3.13.13, 보안 프로그램 DRM) 실기에서 `ml_report.py`가 모든 보고서에서
`no style with name 'Caption'`(STOP_RETRY, exit 1)으로 실패한 문제 수정, 오류 로그(R6) 추가.

### 수정
- 보고서 Word 템플릿을 DRM에 안전하게 엽니다. 원인 추정: DRM이 Office 확장자 파일을 암호화·대체해
  python-docx의 `templates/default.docx`가 표준 스타일이 없는 대체 문서가 됨 (같은 PC의 python-pptx `default.pptx`는 zip이 아니었음).
  - python-docx가 함께 설치하는 압축 해제 템플릿 폴더 `templates/default-docx-template`(XML 파일, DRM 대상 아님)를 먼저 쓰고,
    없거나 열리지 않으면 `default.docx`를 씁니다.
  - 보고서가 쓰는 스타일(`Normal`, `Heading 1~3`, `Caption`, `List Bullet`)이 없으면 `Normal` 기반으로 만들고,
    템플릿에 본문(대체 문서의 안내 문구 등)이 있으면 지운 뒤 작성합니다. 대체 템플릿의 머리글·바닥글도 새로 만듭니다.
  - 위 복구는 stderr `[WARN]`과 결과 JSON `recovered`에 적습니다. 두 템플릿 모두 열 수 없으면 `STOP_RETRY`와
    `python -m pip install --force-reinstall python-docx` 안내.

### 기능 (추가만)
- 오류 로그: 실패한 실행(exit 1, 인자 오류 exit 2)을 `<도구 폴더>\data\logs\errors.txt`에 자동으로 덧붙입니다
  (`CLI-dev/README.md` R6, doc-classifier `write_error_log` 복사). 데이터 폴더는 환경변수 `ML_KIT_DATA_DIR`로 바꿀 수 있습니다.
  - 항목: `===== <시각> | ml-kit <버전> | Python <버전> | <OS>`, `명령:`(스크립트와 인자), `결과: <error_type> | <message>`,
    자동 복구가 한도를 넘긴 경우 `복구 시도:`, 내부 오류(`ERROR:`)일 때만 `traceback:`.
  - 성공, 사용자 중단(exit 130)은 기록하지 않습니다. 2MB를 넘으면 `errors.old.txt`로 넘깁니다 (1세대).
    회전 실패와 기록 실패는 따로 삼켜 결과 JSON 1줄 계약을 깨지 않습니다. 로그 경로는 결과 JSON에 넣지 않습니다.
  - `retry_fs`가 재시도 한도를 넘기면 `record_recovery_attempt()`로 `복구 시도:` 줄에 남깁니다 (R4).
- 공통 런타임(`ml_predict.py` PART C)에 `VERSION = "2.3.1"`. SKILL.md `metadata.version`, 보고서 표지의 ml-kit 버전과 같음을 테스트로 확인합니다.

### 테스트
- `tests/test_report_template.py`: 템플릿 폴더 정상 사용, 손상된 `default.docx`(스타일 없음 + 안내 문구) + 템플릿 폴더 없음에서
  보고서 성공·캡션 존재·안내 문구 없음, 두 템플릿 모두 손상 시 `STOP_RETRY`.
- `tests/test_error_log_v231.py`: exit 1(traceback 없음)/내부 오류(traceback)/exit 2 기록, 성공·130 미기록, 회전, 회전 실패,
  기록 실패 시 JSON 유지, `VERSION` == SKILL.md.
- `tests/conftest.py`: 테스트 중 오류 로그가 실제 도구 폴더 `data\`에 쌓이지 않게 `ML_KIT_DATA_DIR`을 임시 폴더로 지정
  (하위 프로세스도 상속).

### 호환
- CLI 인자와 JSON 키 변경 없음. 실패 시 `data\logs\errors.txt` 파일이 새로 생깁니다 (`CLI/*/data/`는 git 제외).

## [2.3.0] - 2026-10-28

저장소 지침(AGENTS.md "JIT 자가복구" 3번: 덮어쓰기·삭제는 사용자 확인 필요) 반영 1차. 근거: `docs/history/COMPLIANCE_REVIEW_2026-10-28.md` 2절.
구현은 `CLI-dev/README.md` "공통 런타임 구현 규칙" R1(사후 확인), R3(`recovered`), R4(덮어쓰기 방지).

### 수정 (사용자 확인 없이 지우거나 덮어쓰던 동작)
- 학습 시작 시 `leaderboard.json`/`model_dataset.csv`가 없는 `vN` 폴더를 **통째로 지우던** 동작을 없앴습니다.
  이제 지우지 않고 `[WARN]`과 결과 JSON `incomplete_versions`로 알리며, 자동 버전 번호는 그 번호를 건너뜁니다.
  도구가 만든 임시 폴더 `.tmp_*`만 계속 자동 청소합니다.
- `--version`에 이미 있는 버전을 주면 **아무 말 없이 교체하던** 동작을 막았습니다.
  `STOP_RETRY` + `[추천] --version v<다음>`을 내고, 교체는 새 인자 `--overwrite`가 있을 때만 합니다.
  교체 중에는 기존 폴더를 `.bak_vX`로 보관했다가 새 버전이 원장에 기록된 뒤에 지우며, 실패·중단 시 되돌립니다
  (프로세스가 죽어 남은 `.bak_vX`는 다음 학습 때 `vX`가 없으면 복원).
- 원장(`history.jsonl`) 기록을 버전 폴더 커밋 **뒤로** 옮겼습니다. 커밋이 실패하면 원장에 없는 버전이 남지 않고,
  원장 기록이 실패하면 방금 만든 버전 폴더를 롤백합니다.
  `history.jsonl`에 한 줄이 기록되는 즉시 그 버전을 확정하므로, 이후 `history.json` 동기화 중에 중단돼도 폴더를 지우거나 되돌리지 않습니다
  (원장과 폴더가 항상 일치).
- `--overwrite` 교체 시 이전 덮어쓰기의 백업 `.bak_vX`가 남아 있으면 지우지 않고 `STOP_RETRY`로 멈춥니다 (유일한 정상본일 수 있음).
- `ml_predict.py --out`의 파일이 이미 있으면 덮어쓰던 동작, `ml_report.py`가 같은 이름(`ML결과보고서_<데이터>_<버전>.docx`)의
  기존 보고서를 덮어쓰던 동작을 막았습니다. 시각을 붙인 새 이름으로 저장하고 `recovered`에 적습니다. 덮어쓰기는 새 인자 `--overwrite`.
- 실기 러너 `tools/run_win_tests.py`가 `--dir` 폴더를 확인 없이 `rmtree` 하던 동작을 막았습니다.
  러너가 만든 폴더(표식 파일 `.created_by_ml-kit-tests`)나 빈 폴더만 비우고, 다른 폴더면 지우지 않고 준비 단계 오류로 멈춥니다.
- 잠금 때문에 새 이름으로 저장할 때 이름에 시각이 두 번 붙던 문제 (`_pred_<시각>_<시각>.csv`): 기존 시각을 바꿔 붙입니다.
  같은 초에 같은 이름이 있으면 `_2`를 붙여 기존 파일을 덮지 않습니다.

### 기능 (추가만)
- 학습 결과 JSON `verified`: 커밋 후 버전 폴더를 다시 읽어 확인(leaderboard, 데이터 스냅샷, 모델 다시 load, 메타, 원장 마지막 기록).
  확인 실패 시 `verified: false`, `unverified` 사유 목록, `agent_directive`는 `UNVERIFIED:`.
- 모든 도구 결과 JSON `recovered` (있을 때만): 결과가 달라진 자동 복구 문구 목록. 공통 런타임 `note_recovered()`가 stderr `[WARN]`과 함께 모읍니다.
  이번 버전에서 기록하는 것: 기존 파일 덮어쓰기 회피, 잠긴 파일 새 이름 저장.
- 새 인자 `--overwrite`: `ml_train.py`(기존 버전 교체), `ml_predict.py`(`--out` 덮어쓰기), `ml_report.py`(같은 이름 보고서 덮어쓰기).
- SKILL.md: `UNVERIFIED:` 행동, `recovered` 전달, `--overwrite`는 사용자가 허락했을 때만 붙이는 규칙.

### 실기 테스트 러너
- M5/M7은 `--overwrite`를 주고 잠긴 파일을 덮어쓰려 할 때 새 이름 저장 + `recovered`의 '잠김'을 확인합니다
  (덮어쓰기 방지 때문에 잠금 경로가 검사되지 않는 것을 막음).
- 이전 러너가 만든 `C:\ml_test\ml 테스트`에는 표식 파일이 없으므로, 처음 한 번은 그 폴더를 직접 지우거나 다른 `--dir`을 주세요.

### 호환
- 기존 CLI 인자와 JSON 키는 그대로입니다. 다만 위 세 경우(이미 있는 `--version`, 이미 있는 `--out` 파일, 같은 이름 보고서)는
  기존처럼 교체·덮어쓰기하려면 `--overwrite`를 붙여야 합니다 (지침의 "덮어쓰기는 `--overwrite`가 있을 때만", R4).

## [2.2.0] - 2026-10-27

### 실기
- Windows 실기 테스트 러너 `tools/run_win_tests.py` (M1~M14). 집 PC(Windows 11, Python 3.13.13, scikit-learn 1.9.0, Office 2007) 14개 모두 통과:
  Excel이 만든 보고서형 xlsx(제목/단위 행, 날짜 셀, 범주, 한글+공백 경로), Excel 저장 CP949 CSV, 입력 xlsx가 Excel에 열린 상태,
  결과 CSV를 Excel로 열 때 한글, 결과 CSV/보고서가 열려 잠긴 상태(새 이름 저장), 보고서를 Word로 열기,
  `next_commands`를 cmd.exe/PowerShell에 그대로 실행, 단독 배포 폴더, 2만 행(학습 10.6초, 예측 3.6초),
  동시 학습 잠금, 학습 강제 종료 뒤 복구, 오류/출력 계약.

### 기능
- **ML 결과 보고서** (`ml_report.py` 새로 작성): 버전 폴더 하나의 **모든 타겟을 Word 문서 1개**로 만듭니다. 회귀·분류 모두 지원.
  - 구성: 표지, 목차, 요약(타겟별 판정), 데이터, 분석 방법, 타겟별 결과(모델 비교, 성능 검증, 피처 영향), 버전 이력, 운영 가이드, 부록.
  - 그림: 모델 비교, 실측 대비 예측, 잔차 진단, 피처 기여도 / 분류는 혼동행렬, 클래스별 지표, ROC·판정 기준별 F1(이진), 버전 추이, 상관 히트맵.
  - 판정: 0.90 이상 적합, 0.70 이상 조건부, 미만 부적합 (과적합 간극 0.15 초과 시 조건부). 종합 판정은 가장 낮은 타겟 기준.
  - 사용법: `-m <version_dir>` 추가. 기존 `-d <data_root> [-v vN]`, `-t`, `--slots`, `-o` 유지
    (`-v` 없으면 최근 학습과 같은 task·타겟 조합 중 최고 버전). 예전 슬롯 이름(`slot1_residual` 등)도 받음.
  - 결과 JSON: `report_file`, `best_version` 유지 + `version`, `task`, `targets`, `overall_verdict`, `verdicts`,
    `figure_count`, `table_count`, `slots_used` 추가.
- 학습: 챔피언 교차검증 예측값 `vX/oof_<타겟>.csv`, 원장 `data_profile`(불러온/사용 행, 타겟 결측 행, 피처별 결측) 저장.
- 학습 결과 JSON에 `goal_met`(모든 타겟 0.90 이상) 추가. `next_commands`의 보고서 명령은 `ml_report.py -m <이번 버전>`이며 분류에도 제공.

### 변경
- 목표 달성 판단과 권고를 평균 점수가 아니라 **가장 낮은 타겟** 기준으로 합니다 (보고서 판정과 일치).
  다중 타겟에서 "목표 달성" 문구와 타겟별 판정이 어긋나던 문제 해결.
- 다중 타겟 학습의 "저기여 피처 제외" 권고·명령은 **모든 타겟에서 기여도가 낮은 피처**만 대상으로 합니다
  (한 타겟에서만 약한 피처를 빼서 다른 타겟 성능을 떨어뜨리던 문제 해결).

### 실기 테스트 러너 (`tools/run_win_tests.py`)
- AI 판독용 상세 로그 `tools/test_logs/testlog_<라벨>_<시각>.txt`를 **항상** 만듭니다 (준비 단계 실패, 중단 시에도).
  판정 표, 실행 환경(OS, Python, 패키지·Office 버전, 코드 git 버전, 인코딩), 항목별 통과 기준,
  실패 항목의 호출별 명령·종료 코드·소요 시간·전체 JSON·stderr 끝부분·traceback·관련 폴더 상태, 출력 계약 위반, 전체 호출 기록.
  `--label 집PC`로 파일 이름에 PC 라벨을 붙이고, 파일을 그대로 전달하거나 커밋하면 AI가 원인을 판단할 수 있습니다.

### 스킬
- `SKILL.md`를 26B~30B 경량 모델용으로 다시 작성: 요청별 단계 선택 표, 그대로 채우는 명령 템플릿,
  JSON 값별로 사용자에게 알릴 내용, `agent_directive` 첫 단어별 행동 표.

### 호환
- 보고서의 2.1 타겟별 키 `target`, `metrics`, `slots_filled`, `reports`, `target_count`는 문서 통합으로 제거 (사용자 승인).
- 2.2.0 이전에 학습한 버전(`oof_*.csv` 없음)은 보고서 생성 시 `STOP_RETRY`로 재학습을 안내합니다. 예측은 그대로 가능합니다.

## [2.1.2] - 2026-10-27

### 수정
- 같은 데이터로 다른 타겟을 학습하면 `delta_score`/`is_best`가 **다른 타겟의 이전 기록**과 비교되던 문제.
  이제 같은 task + 같은 타겟 조합의 기록끼리만 비교합니다 (새 타겟 조합의 첫 학습은 `delta_score: "base"`).
- `--history`의 `best_version`/`best_mean_val_score`도 가장 최근 학습과 같은 task + 타겟 조합 중 최고 버전을 가리킵니다.
  (`best_by_task`는 기존대로 task별 최고)

## [2.1.1] - 2026-10-27

### 기능 (추가만, 기존 키 유지)
- `next_commands`의 각 명령에 `needs_input`(bool)을 붙입니다. 자리표시자(`<새데이터.xlsx>`)가 있는 명령은 `needs_input: true`와 `placeholder`를 갖고, 바로 실행 가능한 명령 뒤에 둡니다.
  목표 성능 달성(회귀) 시 `next_commands[0]`은 이제 보고서 명령입니다. 실기 테스트에서 SKILL대로 `next_commands[0]`을 그대로 실행하면 자리표시자 때문에 실패하던 문제를 막습니다.
- `agent_directive`와 SKILL.md에 `needs_input` 처리 규칙(사용자가 준 경로로 바꾸거나 먼저 묻기)을 넣었습니다.

## [2.1.0] - 2026-10-27

train/predict 완성 단계. CLI 인자와 기존 JSON 키는 모두 유지(추가만).

### 기능
- **범주형 피처**: 문자 컬럼(라인, 설비코드 등)을 자동 감지해 파이프라인 안에서 원-핫 인코딩.
  예측 시 학습 때 없던 범주는 `unseen_category_warning`. 종류가 너무 많은 ID형 컬럼은 `STOP_RETRY`.
- **날짜 피처**: Excel 날짜·`YYYY-MM-DD[ HH:MM]` 문자열을 연/월/일/요일/시 피처로 자동 분해.
  예측 시 학습 기간 밖 날짜는 외삽 경고.
- 피처 유형 판별 결과는 `feature_types`(결과 JSON·원장)와 `feature_schema`(모델 메타)에 기록. `--dtypes`로 강제 지정 가능.
- **선형 모델 자동 튜닝**: RidgeCV, ElasticNetCV, LogisticRegressionCV. 선택된 값은 리더보드 `params`.
- **회귀 예측 범위**: 챔피언 OOF 잔차 분위수 기반 90% 범위. 예측 컬럼 `pred_lo_<타겟>`, `pred_hi_<타겟>`, JSON `intervals`.
- **이진 분류 판정 기준**: `--positive-class <라벨>` 지정 시 양성 F1이 최대가 되는 확률 기준을 OOF로 찾아 저장·적용.
- **예측 입력 결측 경고**: 빈 값/숫자 아닌 값이 대치되면 `missing_input_warning`, `missing_input_counts`.
- **다음 명령 제공**: 결과 JSON `next_commands` (피처 제외 재학습, 교차항 추가, 목표 달성 시 예측·보고서 명령).
  피처가 15개를 넘으면 목록을 `vX/next_features_*.txt`로 저장해 경로로 전달.
- **동시 실행 잠금**: 같은 데이터에 학습이 겹치면 `STOP_RETRY` (OS 파일 잠금, 비정상 종료 시 자동 해제).

### 성능
- CSV: 구분자 자동 감지(`,` `;` 탭 `|`) 후 고속 C 파서 사용.
- 순열 중요도는 최대 5,000행 표본으로 계산. 1만 행 초과 시 RF/ET 트리마다 1만 행 표본, 5만 행 초과 시 KNN 제외.
- 측정 (이 개발 환경 4코어): 2만 행 × 20피처 80초 -> 46초(정확도 동일), 300행 × 150피처 39초 -> 11초.

### 원장 크기
- 피처가 30개를 넘으면 원장의 상관행렬은 중요도 상위 30개 + 타겟만 저장(`correlation_truncated: true`),
  전체 행렬은 `vX/correlation.csv`, 피처-타겟 상관은 `target_correlations`. (150피처 기준 기록 1건 370KB -> 31KB)

### 호환
- 2.1.0 이전 모델(`feature_schema` 없음)은 기존 수치 배열 방식으로 예측. 새 모델도 `ml_predict.py` 단독 실행 유지.

## [2.0.0] - 2026-10-27

### 구조
- `ml_common.py`를 `ml_predict.py`에 병합 (PART C: 공통 런타임). `.py` 파일 3개 구성.
- `ml_predict.py`는 다른 ml-kit 파일 없이 단독 실행됩니다. **배포 = `ml_predict.py` + `vX\` 폴더 복사.**
  `ml_train.py`, `ml_report.py`는 `ml_predict.py`에서 공통 기능을 import합니다.

### 기능
- **분류 학습** `ml_train.py --task classification` (기본값 `regression`, 사용자 지정).
  - 5개 모델: HistGradientBoosting, RandomForest, ExtraTrees, LogisticRegression, KNeighbors
  - Stratified K-Fold, 챔피언 = OOF Macro-F1 최고
  - 지표: Macro-F1, Accuracy, Balanced Accuracy, ROC-AUC, 혼동행렬, 클래스별 Precision/Recall/F1
  - 최소 클래스 비율 20% 미만이면 `class_weight='balanced'` 자동 적용 (`[WARN]`)
  - 라벨은 문자열로 정규화 (`1`, `1.0`, `"1"` -> `"1"`), 클래스별 최소 2개 미만이면 `STOP_RETRY`
- **분류 예측** `ml_predict.py`: 메타데이터의 `task`로 자동 분기.
  - 출력 컬럼 `pred_<타겟>`, `proba_<타겟>_<클래스>`, `conf_<타겟>`
  - JSON `predictions`, `probabilities`, `predicted_class_counts`, `low_confidence_warning`(확신도 0.6 미만)
- 결과 JSON/원장/메타/리더보드에 `task`, `primary_metric`, `mean_val_score`, `delta_score` 추가.
  성능 변화(`delta`)와 `is_best`는 같은 task 기록끼리만 비교.
- `--history`: Task/Score 표시, `best_by_task` 반환.
- `ml_report.py`: 분류 기록은 `STOP_RETRY`로 "다음 버전 지원" 안내 (회귀 보고서는 동일).

### 호환
- 회귀의 기존 키(`mean_val_r2`, `delta_r2`, `val_r2` 등)와 CLI 인자는 그대로 유지.
- `task` 키가 없는 1.x 모델/원장은 회귀로 처리.

### 스킬
- OpenCode Agent Skills 형식 적용: 스킬 이름 `ml-kit`, 폴더 이름도 `CLI/ml-kit/`로 변경 (OpenCode는 폴더명 = 스킬명 필수).
  frontmatter는 `name`, `description`, `license`, `compatibility`, `metadata`만 사용. 설치 방법은 README "OpenCode 스킬 설치".

### 문서
- 이전 이력 문서(ROADMAP 등)를 정리하고 이 CHANGELOG로 일원화. `requirements-dev.txt`, `install_offline.bat` 삭제.

## 남은 할 일
- [ ] 성능예측표 엑셀을 문서 보안 아래에서 파이썬 없는 PC로 열기 (`docs/opencode/실기테스트.md` 4장 눈 확인 2번). 엑셀 재계산은 회사 PC 통과 (2.9.4)
- [ ] `-x` 생략 시 피처 자동 선택 (ID/일련번호 배제 규칙)
- [ ] 무거운 import 지연 로딩 (`--history` 응답 속도)
- [ ] 260자 넘는 긴 경로
- [ ] `--best`: 조건이 다른 타겟들이 같은 이름의 파생변수를 다른 수식으로 만들면 `model_dataset.csv`·`correlation.csv`에는 첫 수식 값만 들어감 (모델은 타겟별로 정확)
- [ ] 회사 PC에서 6타겟 실제 데이터로 `ml_report.py -f` 소요 시간 확인

### 저장소 지침 반영 (AGENTS.md 2026-10-28, 우선순위 순)
상세 근거: `docs/history/COMPLIANCE_REVIEW_2026-10-28.md`. 공통 키·함수는 `CLI-dev/README.md` "공통 런타임 구현 규칙"(R1~R9)대로 구현.
- [ ] `recovered` 나머지: cp949 -> utf-8-sig 폴백, .xls -> .xlsx 저장 (잠금 새 이름 저장과 덮어쓰기 회피는 2.3.0), 폴백에도 `[WARN]`
- [ ] 파괴적 작업 미리보기 `--dry-run` (R8): 학습 `--overwrite`, 예측·보고서 `--overwrite` 대상
- [ ] 테스트 보강: exit 130 롤백, 성공 경로 `VERIFIED:` 검사 확대, 빈 데이터 (덮어쓰기 차단은 2.3.0)
