# 집 PC 점검 (Gemma 없이, ml-kit 2.9.72 · office-cli 1.4.80)

집 PC에는 Gemma가 없으므로 Gemma가 고르는 일(맞는 스킬 고르기, 번호 `[후보]` 처리)은 회사 PC에서 본다.
집 PC에서는 그 앞단을 확인한다: 도구가 준 고친 명령이 실제 PowerShell 5.1에서 그대로 실행되는지, 스킬 4개가 설치되는지, 기존 기능이 그대로인지.

## 1. 준비 (5분)
집 PC는 WinPython이 아니라 같은 버전(3.13)의 공식 Python이다. `setup.bat`은 `C:\CLI\python\python.exe`(WinPython)를 찾으므로 쓰지 않고
`setup.py`를 공식 Python으로 직접 실행한다 (`C:\CLI\python`이 없으면 실행 중인 Python으로 확인함).
1. GitHub에서 저장소 zip을 받아 `CLI`를 `C:\CLI`로, `CLI-dev`를 `C:\CLI-dev`로 덮어쓴다 (예전 폴더는 지우고 붙여넣기 권장).
2. PowerShell에서 `python`이 공식 Python인지 확인한다 (점검 스크립트가 고친 명령을 `python ...`으로 실행하므로 PATH의 `python`이어야 함):
   ```
   python -c "import sys; print(sys.executable, sys.version)"
   ```
   3.13.x가 나오면 된다. `python`을 못 찾으면 공식 설치 프로그램의 "Add python.exe to PATH"가 꺼져 있는 것이다.
   라이브러리가 없으면 `pip_install.txt`의 2-1 명령(인터넷 설치)을 그 PowerShell에서 실행한다.
3. 스킬 설치 (PATH는 건드리지 않음):
   ```
   python C:\CLI\setup.py --no-path
   ```
   `옛 ml-kit 스킬: 보관 후 지움`(예전에 설치했다면)과 `ml-train`, `ml-predict`, `ml-report`, `ml-map`, `office-data`(새 BOM 비교 스킬) 복사 줄이 보여야 한다.
   `opencode.exe: 없음` `[FAIL]` 1개와 opencode.jsonc·rg(ripgrep) `[WARN]`은 집 PC 환경 차이라 무시한다. 그 밖의 `[FAIL]`이 있으면 결과와 함께 알린다.
4. Excel, Word, PowerPoint 창은 모두 닫아 둔다.

## 2. 사용법 단순화 + 새 기능 점검 (명령 1개, 10~15분)
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
| N1 | ml-kit 2.9.72: `-x` 없이 학습 -> 도구가 입력을 고름 (ID·상수·타겟 복사 열은 이유와 함께 뺌), 재학습 명령 실행 |
| N2 | `--suggest` 다음 실험 조건 (불확실한 곳, `--maximize`), 실험계획 엑셀, 상한 넘음 `[추천]` |
| N3 | 한 건 예측의 입력별 영향 (`explain`, preview의 높인/낮춘 입력) |
| N4 | 실측이 있는 새 데이터 -> `actual_check`와 재학습 알림, 재학습 명령 실행, 같은 데이터면 알림 없음 |
| B1~B5 | office-cli 1.4.80 BOM 비교 `excel compare`: 다단 BOM(상위품번+품번 키), 수량·위치·메이커 변경, 품번 변경 추정, 결과 엑셀, 바뀐 칸 메모 미리보기(Excel), 같은 파일, 표기만 다른 판은 0건 |
| K1 | OpenCode 설정 폴더에 ml-kit 스킬 4개와 office-data, 옛 `skills\ml-kit` 없음 |

결과: `C:\CLI-dev\result_simplify_집PC_<시각>.md`

## 3. 전체 회귀 점검 (명령 1개, 1~2시간, 시간이 되면)
```
python C:\CLI-dev\run_all_tests.py --label 집PC --only diag,pytest,beta,ml,office
```
도중에 PC를 쓰지 않는다 (Office 창이 열렸다 닫힌다). 결과: `C:\CLI-dev\result_집PC_<시각>.md`
`diag` 단계의 `C:\CLI\python`·opencode 관련 FAIL은 공식 Python을 쓰는 집 PC의 환경 차이다 (지난 집 PC 결과와 같음).

## 4. 결과 올리기
위 결과 파일(1~2개)을 저장소 `docs/test_logs/`에 올린다. 분석과 수정은 에이전트가 한다.

## 회사 PC에서 따로 볼 것 (Gemma)
`python C:\CLI-dev\run_all_tests.py --label 회사PC --only gemma`: Gemma가 요청마다 맞는 ml 스킬(학습·예측·보고서·보간)을 고르는지,
예측에 `-m "<데이터파일>"`을 쓰는지, 번호 `[후보]`를 사용자에게 보여 주고 고른 번호의 명령을 실행하는지.
