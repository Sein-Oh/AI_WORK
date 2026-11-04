---
name: ml-kit
description: 표(Excel/CSV) 데이터로 수치 예측(회귀) 또는 범주 예측(분류) 모델을 학습하고(ml_train.py), 새 데이터를 예측하고(ml_predict.py), Word로 ML 결과 보고서(ml_report.py)를 만든다. 모델 없이 표에서 사이값을 선형 보간하거나 목표값이 되는 조건을 반대로 찾는다(ml_map.py). 데이터 정제, 5개 모델 경쟁, 버전 관리, 입력 경고를 도구가 자동 처리하므로 명령 1회 실행 후 JSON 결과만 읽으면 된다. 예측 모델 만들기, 성능 개선, 새 조건 예측값 계산, 결과 보고서 요청 시 사용.
license: Proprietary
compatibility: opencode
metadata:
  version: "2.9.51"
  platform: windows
---

# ml-kit 스킬

## 꼭 지킬 규칙
1. 명령은 아래 **템플릿을 그대로** 쓰고 `<...>` 부분만 바꾼다. 옵션을 지어내지 않는다.
2. PowerShell에서는 명령 앞에 전역 지침의 인코딩 줄(`$env:PYTHONIOENCODING='utf-8'; ...`)을 붙인다.
3. 명령은 **1번만** 실행한다 (예외: 결과가 `CONTINUE:`이면 표대로 다시 실행). 결과는 출력의 **마지막 1줄 JSON**만 읽는다. `[PROGRESS]` 줄은 무시한다.
4. JSON의 `agent_directive` 첫 단어로 다음 행동을 정한다 (맨 아래 표).
5. 같은 명령을 다시 실행하지 않는다. 결과 파일을 다시 열어 확인하지 않는다.
6. 학습은 몇 분 걸린다. 명령 실행 도구의 시간 제한은 600000(10분)으로 준다. 보고서와 성능예측표는 오래 걸리면 도구가 나눠서 진행한다 (`CONTINUE:`).
7. `--overwrite`는 사용자가 "덮어써", "교체해"라고 말했을 때만 붙인다. 스스로 붙이지 않는다.
8. 파일 경로는 결과 JSON의 값을 바꾸지 않고 그대로 1번만 쓴다 (`file:///` 링크로 바꾸지 않는다). 이모지를 쓰지 않는다.
9. 도구 코드(.py)는 사용자가 코드를 열어 보라고 직접 요청할 때만 읽는다. 그 밖에는 읽지 않고, 이 문서의 템플릿으로 방법을 찾지 못하면, 하려던 일과 막힌 곳을 사용자에게 알리고 멈춘다.

도구 위치는 `C:\CLI\ml-kit` (office-cli는 `C:\CLI\office-cli`)이다. 아래 경로를 그대로 쓴다.

## 1단계: 무엇을 할지 고르기

| 사용자 요청 | 할 일 |
|---|---|
| "모델 만들어줘", "예측 모델 학습" | 2단계 학습 |
| "성능 올려봐", "다시 해봐" | 3단계 개선 |
| "버전별로 비교해줘", "지금까지 학습 이력" | 3단계 끝의 이력 명령 |
| "이 조건이면 얼마야?", "새 데이터 예측" | 4단계 예측 |
| 모델을 만든 뒤 "생산효율이 50이 되려면 온도는?" (목표값이 되는 조건) | 4-B단계 모델 반대 찾기 |
| "보고서 만들어줘" | 5단계 보고서 |
| "성능예측표", "엑셀 계산기로 만들어줘" | 5단계 보고서 (성능예측표가 함께 만들어짐). 입력을 바꿔 다시 만들 때만 7단계 |
| 모델 없이 "데이터(표)에서 ~일 때 값", "보간", "맵핑", "표에서 ~가 50이 되는 온도는?" | 6단계 표 보간·반대 찾기 |

실행 전에 꼭 알아야 할 것: 데이터 파일 경로, 타겟(맞힐 컬럼), 피처(입력 컬럼). 모르면 사용자에게 묻는다.
컬럼 이름을 모르는 xlsx는 office-cli로 앞부분만 읽는다: `python "C:\CLI\office-cli\excel.py" read --target "<데이터파일>" --range A1:Z5`

## 2단계: 학습

타겟이 숫자(수율, 온도, 강도)면 회귀, 글자나 등급(양품/불량, A/B/C)이면 분류다. 모르면 사용자에게 묻는다.

타겟이 여러 개면 `-y`에 쉼표로 함께 넣어 1번에 학습하는 것을 권장한다 (예: `-y "수율,강도,두께"`).
타겟별로 따로 학습해도 괜찮다. 보고서(5단계)가 모든 타겟을 자동으로 합친다.

회귀:
```powershell
python "C:\CLI\ml-kit\ml_train.py" -f "<데이터파일>" -x "<피처1>,<피처2>,<피처3>" -y "<타겟1>,<타겟2>" --desc "<데이터 설명 한 줄>"
```
분류:
```powershell
python "C:\CLI\ml-kit\ml_train.py" -f "<데이터파일>" -x "<피처1>,<피처2>,<피처3>" -y "<타겟1>,<타겟2>" --task classification --desc "<데이터 설명 한 줄>"
```

필요할 때만 덧붙이는 옵션:

| 상황 | 덧붙일 옵션 |
|---|---|
| 엑셀 위쪽에 제목/단위 행이 있다 (예: 4행이 머리글, 5행이 단위) | `--header-row 4 --unit-row 5` |
| 분류에서 놓치면 안 되는 값이 있다 (예: 불량) | `--positive-class 불량` |
| 숫자로 된 코드(101, 102)를 종류로 다뤄야 한다 | `--dtypes "설비코드:str"` |
| 사용자가 계산한 입력(비율, 차이)을 원한다 | `--eval "<이름> = <식>; <이름2> = <식2>"` (예: `--eval "온도압력비 = 온도 / 압력"`). 숫자로 읽히는 열 이름(`2`, `2(mm)`)이 식에 있으면 도구가 멈춘다 |
| 피처가 10개 이상이다 | 파일 쓰기 도구로 한 줄에 피처 하나씩 쓴 `features.txt`(UTF-8)를 데이터 파일 옆에 만들고 `-x "<그 경로>"` |
| 숫자 타겟과 등급 타겟이 섞여 있다 | 회귀 타겟끼리 1번, 분류 타겟끼리 `--task classification`으로 1번 (모두 2번) |

- 컬럼명의 단위는 빼도 된다 (`반응온도` = `반응온도(°C)`).
- 날짜 컬럼과 글자 컬럼(라인, 설비)도 `-x`에 그대로 넣는다.
- 이상치는 도구가 자동으로 찾아 학습에서 뺀다. 옵션은 없다. 데이터의 `제외` 열은 도구가 읽는 표시용 열이다 (피처에 넣어도 도구가 뺀다).
  사용자가 행을 빼거나 넣고 싶다고 하면 데이터 파일을 직접 고치지 말고, 엑셀의 `제외` 열에 이유(뺄 때) 또는 `포함`(자동 제외를 막을 때)을 적으라고 안내한다.

결과 JSON에서 사용자에게 알릴 것: `version`, `mean_val_score`(검증 점수, 1에 가까울수록 좋음), `top_features`(타겟별 중요 입력과 비중, 합이 1), `recommendations`의 첫 문장. `warnings`(예: 중복 행, 반복된 머리글 행, 합계·평균 행)가 있으면 그대로.
`exclusions`가 있으면 타겟별 학습 제외 행 수(직접 `manual`, 자동 `auto`), 제외 행 포함 검증 점수 `all_rows_score`, `excluded_rows_file`(뺀 행 목록)도 알린다.
다음 단계를 위해 기억할 것: `version_dir`, `goal_met`, `next_commands`.

## 3단계: 성능 개선

1. `goal_met`이 `true`이면 개선을 멈추고 5단계 보고서로 간다.
2. `goal_met`이 `false`이면 `next_commands[0].command`를 **글자 그대로** 실행한다. 명령을 직접 만들지 않는다.
3. 새 결과의 `delta_score`(양수=좋아짐)를 알린다.
4. 1~3을 최대 5번 반복한다.

`needs_input`이 `true`인 명령에는 `<새데이터.xlsx>` 같은 자리표시자가 있다. 사용자가 준 실제 경로로 바꿔 실행하고, 경로를 모르면 먼저 묻는다.

버전 비교·이력: `python "C:\CLI\ml-kit\ml_train.py" -f "<데이터파일>" --history` -> `records`의 `version`, `mean_val_score`, `delta_score`, `tag`를 표로 알린다 (`truncated`가 true이면 최근 30건만이라고 알리고 `full_result_file` 경로도).

## 4단계: 예측

한 건:
```powershell
python "C:\CLI\ml-kit\ml_predict.py" -m "<version_dir>" -i "<피처1>=<값>,<피처2>=<값>,<피처3>=<값>"
```
파일 전체:
```powershell
python "C:\CLI\ml-kit\ml_predict.py" -m "<version_dir>" -f "<새데이터파일>"
```
결과를 엑셀로 받고 싶다고 하면: `python "C:\CLI\ml-kit\ml_predict.py" -m "<version_dir>" -f "<새데이터파일>" --save "<전체경로.xlsx>"`
- 학습 때 `--header-row`/`--unit-row`를 줬으면 파일 예측에도 같은 옵션을 붙인다 (학습 결과 `next_commands`의 예측 명령에 들어 있음).
- `-i`에는 모델 입력만 넣는다. 모델 입력이 아닌 이름(`습도`, 다른 단위 `압력(MPa)`)이 있으면 도구가 멈춘다 (`[후보]`나 목록을 사용자에게 보여 주고 묻는다).
- `-i "k=v"`는 한 건이다. 여러 건은 JSON 배열(`-i "[{...},{...}]"`)이나 .json 파일로 준다 (같은 이름이 두 번이면 도구가 멈춘다).
- `-i`에 6건 이상을 넣으면 건별 값 대신 `output_file`(결과 파일)을 알린다 (`truncated: true`).
- `-i`의 숫자 값에는 쉼표를 쓰지 않는다 (`압력=3000`, `3,000` 아님). 쉼표가 있으면 도구가 멈추고 `[추천]` 명령을 준다 (`--find`, `ml_map.py`도 같음).
- `<version_dir>`: 보고서를 만들었으면 보고서 결과의 `best_version_dir`, 아니면 마지막 학습 결과의 `version_dir`을 그대로 쓴다. 모르면 `python "C:\CLI\ml-kit\ml_train.py" -f "<데이터파일>" --history`를 실행해 `best_version_dir`을 쓴다 (`best_version_note`가 있으면 그대로 알린다).
- 알릴 것: 한 건이면 `predictions_summary.<타겟>.predictions`, 회귀는 `intervals`(90% 예측 범위), 분류는 `probabilities`(종류별 확률)도. 파일이면 `output_file`
  (열 `pred_<타겟>`, 회귀 `pred_lo_<타겟>`/`pred_hi_<타겟>` = 90% 예측 범위, 분류 `proba_<타겟>_<종류>` = 확률).
- 회귀는 `map_check`(실측 비교값: 학습 데이터에서 같은 조건을 찾은 값)의 `value`, `diff`(예측-비교값), `method`를 "실측 비교값"으로 함께 알린다.
  `map_check_warning`이 true이면 `map_check_warnings`를 그대로 알린다. `map_check_notes`가 있으면 비교값을 못 낸 이유로 짧게 알린다.
- 아래 값이 `true`이면 옆의 문장을 덧붙인다:

| JSON 값 | 덧붙일 말 |
|---|---|
| `drift_warning` | 입력값이 학습 범위를 벗어나 예측을 참고용으로만 써야 함 (`extrapolation_warnings` 내용 전달, 최대 20건이니 `extrapolation_warning_count` 전체 건수도) |
| `unseen_category_warning` | 학습 때 없던 종류가 있음 (`unseen_category_warnings` 내용 전달) |
| `missing_input_warning` | 빈 값이 있어 기준값으로 채움 (`missing_input_warnings` 내용 전달) |
| `low_confidence_warning` | 분류 확신도가 낮은 건이 있음 (`low_confidence_warnings` 내용 전달) |

## 4-B단계: 모델 반대 찾기 (목표값이 되는 입력값)

```powershell
python "C:\CLI\ml-kit\ml_predict.py" -m "<version_dir>" --find "<타겟>=<목표값>" --vary "<찾을입력>" -i "<고정입력1>=<값>,<고정입력2>=<값>"
```
- `-i`에는 사용자가 말한 나머지 입력만 넣는다. 말하지 않은 숫자 입력은 도구가 학습 중앙값으로 채운다. 글자 입력(라인 등)이 빠지면 `[후보]`가 나오니 사용자에게 묻는다.
- `--find`와 `-i`는 한 건만 준다 (JSON 배열이 여러 건이면 도구가 멈춘다). 여러 경우는 경우마다 따로 실행한다.
- 알릴 것: `solutions`(답 모두, 하나만 고르지 않음), `operating_windows`(목표가 예측 오차 안에 드는 운전 구간), `measured`(학습 데이터 기준 답),
  `nearest_rows`(가까운 실측 5행). 0건이면 `reachable`. `warnings`는 반드시 함께. 모델이 예측한 조건이지 보장된 조건이 아니라고 덧붙인다.

## 5단계: ML 결과 보고서 (Word)와 성능예측표 (엑셀)

```powershell
python "C:\CLI\ml-kit\ml_report.py" -f "<데이터파일>" --title "<제목>"
```
- 숫자 타겟(회귀)이면 성능예측표(엑셀 계산기)도 함께 만들어진다. `<제목>`은 학습 때 데이터 설명(`description`)을 30자 안팎의 짧은 제목으로 정리해 넣는다 (예: "반응 공정 생산효율 예측").
  입력을 줄이며 다시 학습하므로 몇 분 걸린다 (시간 제한 600000).
- 도구가 이력에서 타겟마다 가장 좋았던 학습 조건으로 모든 타겟을 묶은 최고 모델을 자동으로 만들고(필요할 때만 다시 학습, 몇 분 걸릴 수 있으니 시간 제한 600000) 보고서 1개를 만든다. 타겟이나 버전을 고르지 않는다.
- 회귀와 분류 타겟이 섞여 있으면 보고서가 2개(`best_reports`)다.
- 알릴 것: `report_file`, `export_file`(성능예측표)과 `export`의 타겟별 `inputs`(남긴 입력)·`agreement.agreement_pct`(모델 대비 일치도 %), `overall_verdict`(적합/조건부/부적합), `made_best`(true면 이번에 다시 학습함),
  `best_targets`의 타겟별 `source_version`(학습 조건 출처), `source_score`(그때 점수), `new_score`(이번 점수)를 짧게.
- `export_note`, `export_warnings`가 있으면 그대로 알린다 (보고서는 정상). 성능예측표를 다 확인하지 못하면(엑셀이 없는 PC) `UNVERIFIED:`이고 `unverified`를 그대로 알린다.
- 이후 예측(4단계)에는 결과의 `best_version_dir`을 쓴다.
- 사용자가 보고서 문장을 직접 정해 주면, 파일 쓰기 도구로 `slots.json`(UTF-8)을 만들고 명령 끝에 `--slots "<그 경로>"`를 붙인다.
  형식: `{"summary": "<요약 문장>", "operation": "<운영 문장>"}`

## 6단계: 표 보간·반대 찾기 (ml_map.py, 학습 불필요)

사잇값 (기준 열과 값은 -i, 구할 열은 -y, -y를 빼면 나머지 숫자 열 전부):
```powershell
python "C:\CLI\ml-kit\ml_map.py" -f "<데이터파일>" -i "<기준열>=<값>" -y "<구할열>"
```
반대 찾기 (목표 열이 목표값이 되는 기준 값, 다른 기준은 -i로 고정, 고정할 것이 없으면 -i를 뺀다):
```powershell
python "C:\CLI\ml-kit\ml_map.py" -f "<데이터파일>" --find "<목표열>=<목표값>" --vary "<찾을기준열>" -i "<고정열>=<값>"
```
- 머리글이 1행이 아니면 2단계처럼 `--header-row`, `--unit-row`를 붙인다. 여러 점은 기준 열만 든 조회 파일로 `--points "<조회파일>"`.
  조회 파일에 구할 열의 실측값 열(예: `생산량_실측`)이 있으면 `-y "<구할열>"`을 꼭 붙인다 (그 열은 결과에 그대로 붙는다).
- 알릴 것: `preview`, `points[0].method`와 `basis`(쓴 원본 행과 비율, `basis_by_column`이 있으면 열마다), `nearest_rows`(가까운 실측 5행)를 표로.
  반대 찾기는 `solutions`를 모두(하나만 고르지 않음), 0건이면 `reachable`. `output_file`이 있으면 그 경로.
- `extrapolated`가 true이거나 `warnings`가 있으면 그 내용을 반드시 함께 알린다 (범위 밖 외삽값, 흩어진 기록).

## 7단계: 성능예측표 다시 만들기 (ml_export.py)

성능예측표는 5단계 보고서와 함께 만들어진다. 사용자가 꼭 넣을 입력을 말하거나 다른 입력 조합을 원할 때만 이 명령을 쓴다.
```powershell
python "C:\CLI\ml-kit\ml_export.py" -m "<best_version_dir>" --keep "<입력1>,<입력2>" --title "<제목>"
```
- 입력을 줄이며 다시 학습하므로 몇 분 걸린다 (시간 제한 600000). 같은 이름 파일이 있으면 시각을 붙인 새 이름으로 저장된다.
- 알릴 것: `output_file`, 타겟별 `inputs`(남긴 입력)와 `dropped`(뺀 입력), `val_r2`, `table_rows`.
  사용자가 다른 입력 조합을 원하면 `candidates`를 표로 보여 주고, 고른 안의 `command`를 그대로 실행한다.
- 입력은 타겟마다 8개까지다. 더 많이 넣으면 `STOP_RETRY`와 `[추천]` 조합이 나온다.

## 결과 처리 (모든 단계 공통)

| `agent_directive` 첫 단어 | 할 일 |
|---|---|
| `VERIFIED:` | 성공. `preview`와 위에서 정한 값을 사용자에게 알리고 끝낸다. |
| `UNVERIFIED:` | 실행은 됐지만 결과를 확인하지 못했다. 다시 실행하지 않는다. `unverified` 내용을 그대로 알린다. |
| `CONTINUE:` | 아직 끝나지 않았다 (시간 제한 때문에 나눠서 진행, 한 일은 저장됨). 사용자에게 묻지 말고 `next_commands`의 명령을 그대로 다시 실행한다. `VERIFIED:`나 `UNVERIFIED:`가 나올 때까지 반복한다 (최대 10번, 그래도 `CONTINUE:`이면 `progress`를 알리고 멈춤). |
| `STOP_RETRY:` 이고 `message`에 `[추천]`이 있다 | `[추천]` 뒤의 값(`--header-row 2` 같은 옵션 값)으로 바꿔 **1번만** 다시 실행한다. `[추천]` 뒤가 `python`으로 시작하는 명령이면 그 명령을 그대로 실행한다. 또 실패하면 `message`를 그대로 알린다. |
| `STOP_RETRY:` 이고 `message`에 `[후보]`가 있다 | 다시 실행하지 않는다. `[후보]`를 보여 주고 사용자가 고른 값으로 1번 실행한다 (후보가 하나여도 묻는다: 비슷한 열 이름은 다른 양일 수 있음). "(단위 다름: 값 환산 필요)"가 붙은 후보를 고르면 그 단위로 환산한 값도 사용자에게 받는다. |
| `STOP_RETRY:` 이고 `[추천]`, `[후보]`가 없다 | 다시 실행하지 않는다. `message`를 그대로 알린다. `install_command`가 있으면 그 명령을 알린다(직접 설치하지 않는다). |
| `ERROR:` | 다시 실행하지 않는다. `message`를 그대로 알린다. |

결과 JSON에 `recovered`가 있으면 (예: 같은 이름 파일이 있어 새 이름으로 저장) 그 문장을 사용자에게 그대로 알린다.
