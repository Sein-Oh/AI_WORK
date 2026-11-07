# 집 PC 점검 (Gemma 없이, ml-kit 2.9.73 · office-cli 1.4.82)

집 PC에는 Gemma가 없으므로 Gemma가 고르는 일(맞는 스킬 고르기, 번호 `[후보]` 처리)은 회사 PC에서 본다.
집 PC에서는 그 앞단을 확인한다: 도구가 준 고친 명령이 실제 PowerShell 5.1에서 그대로 실행되는지, 스킬 4개가 설치되는지, 기존 기능이 그대로인지.

## 1. 준비 (5분)
집 PC도 회사 PC와 같이 WinPython 3.13을 `C:\CLI\python`에 두고, Office는 Office 16(Professional Plus 2019, x64)이다 (2026-11-07부터).
1. GitHub에서 저장소 zip을 받아 `CLI`를 `C:\CLI`로, `CLI-dev`를 `C:\CLI-dev`로 덮어쓴다 (예전 폴더는 지우고 붙여넣기 권장, `C:\CLI\python`은 둠).
2. `C:\CLI\setup.bat`을 오른쪽 클릭 -> 관리자 권한으로 실행 (스킬 복사와 시스템 PATH 등록).
   `C:\CLI\python`에 라이브러리가 없으면(`라이브러리 import` `[FAIL]`) `pip_install.txt`의 2-1 명령을 `C:\CLI\python\python.exe -m pip ...`로 실행한 뒤
   setup.bat을 다시 실행한다.
3. 새 PowerShell 창에서 확인한다 (점검 스크립트가 고친 명령을 `python ...`으로 실행하므로 PATH의 `python`이어야 함):
   ```
   python -c "import sys; print(sys.executable, sys.version)"
   ```
   `C:\CLI\python\python.exe`와 3.13.x가 나오면 된다. `WindowsApps\python.exe`가 나오면 PATH 등록 전 창이니 새 창을 연다.
   setup.bat 결과의 `opencode.exe: 없음` `[FAIL]` 1개와 opencode.jsonc·rg(ripgrep) `[WARN]`은 집 PC 환경 차이라 무시한다.
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
| C1 | ml-kit 2.9.73: 옵션 오타 `--data-ro` -> `[후보] 1) --data-row`와 같은 번호의 명령 -> 학습 |
| C2 | office-cli 1.4.81: `reader search --query "A AND B"` -> 고친 전체 명령 `[추천]`(`A;B`) -> 검색 |
| C3 | 닫힌 워드에 같은 제목 둘 -> `[후보] 1) 2)`와 `--occurrence 1` 명령 -> 읽기 |
| E1 | 닫힌 엑셀 빈 범위 읽기 -> `count 0`과 "찾은 항목 없음" |
| D1~D4 | `--dry-run` 미리보기: 엑셀 replace·시트 지우기, 워드 replace, PPT 슬라이드 지우기. 미리보기 뒤 파일이 그대로인지, `confirm_commands[0]`을 실행하면 실제로 바뀌는지 (Office) |
| K1 | OpenCode 설정 폴더에 ml-kit 스킬 4개와 office-data, 옛 `skills\ml-kit` 없음 |

결과: `C:\CLI-dev\result_simplify_집PC_<시각>.md`

## 3. 전체 회귀 점검 (명령 1개, 1~2시간, 시간이 되면)
```
python C:\CLI-dev\run_all_tests.py --label 집PC --only diag,pytest,beta,ml,office
```
도중에 PC를 쓰지 않는다 (Office 창이 열렸다 닫힌다). 결과: `C:\CLI-dev\result_집PC_<시각>.md`
`diag` 단계는 PASS여야 한다 (opencode가 없어도 diag는 실패로 보지 않음).

## 4. 결과 올리기
위 결과 파일(1~2개)을 저장소 `docs/test_logs/`에 올린다. 분석과 수정은 에이전트가 한다.

## 회사 PC에서 따로 볼 것 (Gemma)
`python C:\CLI-dev\run_all_tests.py --label 회사PC --only gemma`: Gemma가 요청마다 맞는 ml 스킬(학습·예측·보고서·보간)을 고르는지,
예측에 `-m "<데이터파일>"`을 쓰는지, 번호 `[후보]`를 사용자에게 보여 주고 고른 번호의 명령을 실행하는지.
