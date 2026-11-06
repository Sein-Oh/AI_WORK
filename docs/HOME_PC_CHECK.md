# 집 PC 점검 (Gemma 없이, office-cli 1.4.77 · ml-kit 2.9.69)

집 PC에는 Gemma가 없으므로 gemma 단계만 빼고 나머지를 모두 실행한다. 실제 Office(Excel, Word, PowerPoint)가 있어야 office 단계가 돈다.

## 1. 준비 (5분)
1. GitHub에서 저장소 zip을 받아 `CLI`를 `C:\CLI`로, `CLI-dev`를 `C:\CLI-dev`로 덮어쓴다 (예전 폴더는 지우고 붙여넣기 권장. `C:\CLI\python`은 그대로 둔다).
2. `C:\CLI\setup.bat` 더블클릭 -> 마지막 줄 `FAIL 0` 확인. opencode.jsonc 관련 `[WARN]`은 집 PC에서는 무시해도 된다.
   라이브러리 `[FAIL]`이 있으면 `pip_install.txt`의 2-1 명령(인터넷 설치)을 WinPython 명령 프롬프트에서 실행한다.
3. Excel, Word, PowerPoint 창은 모두 닫아 둔다 (러너가 시험 파일을 직접 연다).

## 2. 실행 (명령 1개, 1~2시간)
PowerShell 또는 명령 프롬프트에서:

```
python C:\CLI-dev\run_all_tests.py --label 집PC --only diag,pytest,beta,ml,office
```

- 도중에 PC를 쓰지 않는다 (Office 창이 열렸다 닫힌다). 질문(y/n)은 나오지 않는다.
- pytest 단계의 ml-kit이 가장 오래 걸린다 (회사 PC 기준 1시간 이상).

## 3. 결과 올리기
- 끝나면 `C:\CLI-dev\result_집PC_<시각>.md` 파일 1개가 생긴다. 그 파일을 저장소 `docs/test_logs/`에 올린다. 분석과 수정은 에이전트가 한다.

## 이번에 새로 확인되는 항목
| 항목 | 내용 |
|---|---|
| A2 · A3 · I26 | 닫힌 엑셀·워드 파일 쓰기 (창 유지, 프로세스 정리, undo) |
| I28 | 닫힌 파일 `write --cells` + 메모 달기·목록이 성공으로 보고되는지 (회사 PC O2의 0x800a01a8 실패) |
| I29 | `--cells`를 쉼표·띄어쓰기·세로줄로 나누면 쓰지 않고 `;`로 고친 명령 `[추천]`, `--where` 조건 사이 `,`도 `[추천]` |
| I30 | `excel check`가 값 열 오른쪽 `단위` 열을 자동으로 읽어 판정 (회사 PC O1) |
| I31 | `word fill --slots` (값 파일 없이), 양식을 `--file`로 주면 `--target` 안내 (회사 PC O5) |
| E8 · E10 | `write --cells` 덮어쓰기 확인, 규격 대조 |
| ml 단계 | 실제 Excel이 만든 파일로 학습·예측·보고서·성능예측표 |
| pytest | ml-kit 2.9.69의 저장 뒤 감시 응답, 구분자 `[추천]`, 인자 오류 `[추천]`/`[후보]` |

Gemma로만 확인할 수 있는 것(O1, O2, O5 대화 시나리오)은 회사 PC에서 `--only gemma`로 따로 확인한다.
