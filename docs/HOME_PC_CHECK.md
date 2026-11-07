# 집 PC 점검 (Gemma 없이, ml-kit 2.9.71 · office-cli 1.4.79)

집 PC에는 Gemma가 없으므로 Gemma가 고르는 일(맞는 스킬 고르기, 번호 `[후보]` 처리)은 회사 PC에서 본다.
집 PC에서는 그 앞단을 확인한다: 도구가 준 고친 명령이 실제 PowerShell 5.1에서 그대로 실행되는지, 스킬 4개가 설치되는지, 기존 기능이 그대로인지.

## 1. 준비 (5분)
1. GitHub에서 저장소 zip을 받아 `CLI`를 `C:\CLI`로, `CLI-dev`를 `C:\CLI-dev`로 덮어쓴다 (예전 폴더는 지우고 붙여넣기 권장, `C:\CLI\python`은 그대로).
2. `C:\CLI\setup.bat` 더블클릭 -> 마지막 줄 `FAIL 0` 확인. 이번에는 `옛 ml-kit 스킬: 보관 후 지움`과 `ml-train`, `ml-predict`, `ml-report`, `ml-map`
   스킬 복사 줄이 보여야 한다. opencode.jsonc 관련 `[WARN]`은 집 PC에서는 무시해도 된다.
3. Excel, Word, PowerPoint 창은 모두 닫아 둔다.

## 2. 사용법 단순화 점검 (명령 1개, 5~10분)
PowerShell에서:
```
python C:\CLI-dev\run_simplify_check.py --label 집PC
```
틀린 명령을 일부러 실행하고, 도구가 준 고친 명령(`[추천]` 또는 `[후보]` 번호의 명령)을 Gemma처럼 PowerShell에 그대로 넣어 실행해 성공하는지 본다.

| 항목 | 확인 |
|---|---|
| M1 | 열 이름 오타 -> `[후보] 1) ...`과 같은 번호의 명령 -> 학습 |
| M2 · M3 | 예전 줄임말 `--target`, 별칭 `--description` -> 고친 전체 명령 -> 학습 |
| M4 · M5 | `--dtypes` 오타, 두 줄 머리글 엑셀 -> 고친 전체 명령 -> 학습 |
| M6 | 예측 `-m <데이터 파일>` -> 도구가 최고 버전을 고름 |
| M7 | `--save` 확장자 -> 고친 명령 -> 예측 저장 |
| M8 | `--dry-run` 교체 미리보기와 그 확인 명령 |
| M9 | 성능예측표 `-m <데이터 파일>`과 `--keep` 오타 `[후보]` |
| O1 · O2 | 닫힌 엑셀의 없는 시트 `[후보]`, `--where` 쉼표 `[추천]` (경로에 띄어쓰기·`$`·괄호) |
| O3 | `transfer` 옛 `--from`/`--to` -> `--from-app`/`--to-app` |
| O4 · O5 · O6 | 닫힌 엑셀 `--cells`, 워드 `--slots`, PPT `--cards` 구분자 실수 -> 고친 명령을 실제 Office로 실행 |
| K1 | OpenCode 설정 폴더에 ml-kit 스킬 4개, 옛 `skills\ml-kit` 없음 |

결과: `C:\CLI-dev\result_simplify_집PC_<시각>.md`

## 3. 전체 회귀 점검 (명령 1개, 1~2시간, 시간이 되면)
```
python C:\CLI-dev\run_all_tests.py --label 집PC --only diag,pytest,beta,ml,office
```
도중에 PC를 쓰지 않는다 (Office 창이 열렸다 닫힌다). 결과: `C:\CLI-dev\result_집PC_<시각>.md`

## 4. 결과 올리기
위 결과 파일(1~2개)을 저장소 `docs/test_logs/`에 올린다. 분석과 수정은 에이전트가 한다.

## 회사 PC에서 따로 볼 것 (Gemma)
`python C:\CLI-dev\run_all_tests.py --label 회사PC --only gemma`: Gemma가 요청마다 맞는 ml 스킬(학습·예측·보고서·보간)을 고르는지,
예측에 `-m "<데이터파일>"`을 쓰는지, 번호 `[후보]`를 사용자에게 보여 주고 고른 번호의 명령을 실행하는지.
