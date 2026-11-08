# CLI 실기 점검 결과: homePC_gemma4_r3

- 시작 2026-11-08 10:23:08, 끝 10:23:55
- 도구: ml-kit 2.9.77, office-cli 1.4.83 / Python 3.13.13 / Windows-11-10.0.26200-SP0
- 이 파일 하나를 그대로 올리세요. 편집하지 마세요.

## 요약

| 단계 | 결과 | PASS | FAIL | SKIP | 시간(초) | 비고 |
|---|---|---|---|---|---|---|
| 1. Gemma 종합 테스트 (opencode run) | PASS | 2 | 0 | 0 | 47 | 세션 ses_e473f6a7bffe9opJ6IqcwUQXgd, 기록 폴더 C:\CLI-dev\work\gemma_runs_20261108_102308, 시험 폴더 C:\CLI-dev\work\gemma |
| 도구 오류 로그 | - | - | - | - | - | ml-kit 0건, office-cli 1건 |

### 실패 0건

- 없음

### 건너뜀 (ENV-SKIP, SKIP) 0건

- 없음

### UNVERIFIED (도구가 결과를 확인하지 못한 호출) 0건

- 없음

## 1. Gemma 종합 테스트 (opencode run): PASS

세션 ses_e473f6a7bffe9opJ6IqcwUQXgd, 기록 폴더 C:\CLI-dev\work\gemma_runs_20261108_102308, 시험 폴더 C:\CLI-dev\work\gemma

### 채점 결과 (grade_opencode)

```text
# OpenCode 종합 테스트 채점 (homePC_gemma4_r3)

- 채점 시각: 2026-11-08 10:23
- 대화 파일: opencode run 기록 C:\CLI-dev\work\gemma_runs_20261108_102308 (2건, 세션 ses_e473f6a7bffe9opJ6IqcwUQXgd)
- 대화 파일: 요청별 모델: A1=openrouter/google/gemma-4-26b-a4b-it, P1=openrouter/google/gemma-4-26b-a4b-it
- 결과: 채점 2건 중 통과 2건 (100%), 대화에 없음 18건

| 요청 | 결과 | 명령 수 | 걸린 초 | 실패 항목 |
|---|---|---|---|---|
| S0 | 없음 | 0 |  | 대화에 이 요청이 없음 |
| S1 | 없음 | 0 |  | 대화에 이 요청이 없음 |
| A1 | 통과 | 3 | 29 |  |
| A2 | 없음 | 0 |  | 대화에 이 요청이 없음 |
| A3 | 없음 | 0 |  | 대화에 이 요청이 없음 |
| M1 | 없음 | 0 |  | 대화에 이 요청이 없음 |
| M2 | 없음 | 0 |  | 대화에 이 요청이 없음 |
| M3 | 없음 | 0 |  | 대화에 이 요청이 없음 |
| P1 | 통과 | 1 | 8 |  |
| P2 | 없음 | 0 |  | 대화에 이 요청이 없음 |
| P3 | 없음 | 0 |  | 대화에 이 요청이 없음 |
| R1 | 없음 | 0 |  | 대화에 이 요청이 없음 |
| R2 | 없음 | 0 |  | 대화에 이 요청이 없음 |
| O1 | 없음 | 0 |  | 대화에 이 요청이 없음 |
| O2 | 없음 | 0 |  | 대화에 이 요청이 없음 |
| O3 | 없음 | 0 |  | 수동 전용 요청 (자동 실행에서는 보내지 않음) |
| O4 | 없음 | 0 |  | 대화에 이 요청이 없음 |
| O5 | 없음 | 0 |  | 대화에 이 요청이 없음 |
| X1 | 없음 | 0 |  | 대화에 이 요청이 없음 |
| X2 | 없음 | 0 |  | 대화에 이 요청이 없음 |

## 요청별 기록

### A1

요청: [A1] C:\CLI-dev\work\gemma\공정_다중출력.xlsx 로 생산효율, 불순물, 소비전력을 예측하는 모델 만들어줘. 입력은 반응온도, 압축압력, 촉매비율, 수분함량, 유량, 분진이야. 엑셀 위쪽에 제목이 있고 머리글은 3행, 단위는 4행이야.

- [OK] 명령: ml_train\.py
- [OK] 도구 결과 ok
- [OK] 답변: R2|R²|점수
- [OK] 패키지 설치 시도 없음
- [OK] python 외 실행기(py, python3) 없음
- [OK] 임의 파이썬 코드 실행 없음 (python -c, 표준 입력, 직접 만든 .py)
- [OK] 도구 코드(.py) 읽기 없음
- [OK] .py 파일 만들기·고치기 없음
- [OK] office-cli에 없는 명령·옵션 없음 (excel write --text/--cell, office.py write, transfer --source/--text/--to/--from, ppt slide --to, undo --target, excel list)
- [OK] 같은 명령 3번 이상 반복 없음
- [OK] 답변에 이모지·file:///·같은 말 반복 없음

명령:
- `skill`
- `python "C:\CLI\office-cli\excel.py" read --target "C:\CLI-dev\work\gemma\공정_다중출력.xlsx" --range A1:Z5` -> ok / VERIFIED: 작업이 완료되었습니다. 같은 명령을 다시 실행하거나 결과를 다시 조회하지 마세요.
- `python "C:\CLI\ml-kit\ml_train.py" -f "C:\CLI-dev\work\gemma\공정_다중출력.xlsx" -y "생산효율,불순물,소비전력" -x "반응온도,압축압력,촉매비율,수분함량,유량,분진" --header-row 3 --unit-row 4 --desc "생산효율, 불순물, 소비전력 예측 모델 학습"` -> ok / VERIFIED: v1 학습을 마쳤습니다 (regression, 평균 검증 R2=0.9715). 학습에서 뺀

최종 답변:

> 생산효율, 불순물, 소비전력을 예측하는 모델(v1) 학습을 완료했습니다. 모든 타겟의 검증 R2가 0.90 이상으로 성능이 우수합니다.
> 
> [학습 결과]
> - 평균 검증 R2: 0.9715
> - 타겟별 성능 (ExtraTrees):
>   - 생산효율(%): 0.9287
>   - 불순물(ppm): 0.9895
>   - 소비전력(kW): 0.9964
> - 중요 입력 (기여도):
>   - 생산효율(%): 반응온도(0.81), 수분함량(0.10), 촉매비율(0.08)
>   - 불순물(ppm): 반응온도(0.71), 압축압력(0.29)
>   - 소비전력(kW): 유량(0.79), 압축압력(0.21)
> 
> [학습 제외 데이터]
> 총 16개 행이 제외되었습니다.
> - 생산효율(%): 14개 (직접 1, 자동 이상치 13)
> - 불순물(ppm): 1개 (직접 1)
> - 소비전력(kW): 1개 (직접 1)
> * 제외 상세 내역: `C:\CLI-dev\work\gemma\공정_다중출력_ml_data\v1\excluded_rows.csv`
> 
> [다음 작업 추천]
> 1. **보고서 및 엑셀 계산기 생성**: 학습 결과 보고서(Word)와 엑셀에서 바로 예측 가능한 계산기를 만듭니다.
>    `python C:\CLI\ml-kit\ml_report.py -f C:\CLI-dev\work\gemma\공정_다중출력.xlsx`
> 2. **새 데이터 예측**: 새로운 엑셀 파일의 값을 예측합니다.
> 3. **단일 조건 예측**: 입력값을 직접 지정하여 결과를 확인합니다.
> 
> 보고서 생성을 진행할까요?

### P1

요청: [P1] 방금 만든 모델로 반응온도 182, 압축압력 3, 촉매비율 1, 수분함량 20, 유량 10, 분진 0.5 일 때 생산효율, 불순물, 소비전력 예측해줘.

- [OK] 명령: ml_predict\.py
- [OK] 명령: (-m|--model)\s
- [OK] 명령: (-i|--input)\s
- [OK] 도구 결과 ok
- [OK] 답변: 실측 또는 측정
- [OK] 패키지 설치 시도 없음
- [OK] python 외 실행기(py, python3) 없음
- [OK] 임의 파이썬 코드 실행 없음 (python -c, 표준 입력, 직접 만든 .py)
- [OK] 도구 코드(.py) 읽기 없음
- [OK] .py 파일 만들기·고치기 없음
- [OK] office-cli에 없는 명령·옵션 없음 (excel write --text/--cell, office.py write, transfer --source/--text/--to/--from, ppt slide --to, undo --target, excel list)
- [OK] 같은 명령 3번 이상 반복 없음
- [OK] 답변에 이모지·file:///·같은 말 반복 없음

명령:
- `python "C:\CLI\ml-kit\ml_predict.py" -m "C:\CLI-dev\work\gemma\공정_다중출력_ml_data\v1" -i "반응온도(°C)=182,압축압력(bar)=3,촉매비율(wt%)=1,수분함량(%)=20,유량(L/min)=10,분진(mg)=0.5"` -> ok / VERIFIED: 타겟 3개를 예측했습니다. predictions_summary의 예측값과 회귀는 inter

최종 답변:

> 입력하신 조건에 따른 예측 결과입니다.
> 
> [예측 결과]
> - **생산효율(%): 76.33** (90% 범위: 75.08 ~ 77.51)
>   - 실측 비교값: 75.36 (차이: +0.97, 방법: 가까운 실측 5행 거리 가중평균)
>   - 높인 입력: 압축압력 (+0.014)
>   - 낮춘 입력: 반응온도 (-0.839), 촉매비율 (-0.131)
> - **불순물(ppm): 81.61** (90% 범위: 78.21 ~ 85.41)
>   - 실측 비교값: 83.86 (차이: -2.26, 방법: 가까운 실측 5행 거리 가중평균)
>   - 높인 입력: 없음
>   - 낮춘 입력: 반응온도 (-2.74), 압축압력 (-0.146)
> - **소비전력(kW): 42.93** (90% 범위: 42.22 ~ 43.71)
>   - 실측 비교값: 45.82 (차이: -2.89, 방법: 가까운 실측 5행 거리 가중평균)
>   - 높인 입력: 유량 (+0.641)
>   - 낮춘 입력: 압축압력 (-0.304)
```

### opencode run 실행 기록

```text
S0: 러너가 직접 실행 make_beta_samples.py -> exit 0 
S1: 러너가 직접 실행 make_office_samples.py -> exit 0 
A1: 모델 openrouter/google/gemma-4-26b-a4b-it, exit 0, 31.7초, 세션 ses_e473f6a7bffe9opJ6IqcwUQXgd
P1: 모델 openrouter/google/gemma-4-26b-a4b-it, exit 0, 12.8초, 세션 ses_e473f6a7bffe9opJ6IqcwUQXgd
```

### 요청별 모델

```text
A1: openrouter/google/gemma-4-26b-a4b-it
P1: openrouter/google/gemma-4-26b-a4b-it
```

## 도구 오류 로그 (<도구>/data/logs/errors.txt)

ml-kit 0건, office-cli 1건

### office-cli errors.txt

```text
파일: C:\CLI\office-cli\data\logs\errors.txt
===== 2026-11-07T22:07:15 | office-cli 1.4.82 | Python 3.13.13 | Windows-11-10.0.26200-SP0
명령: excel.py info --file C:\CLI-dev\work\gemma_mini_20261107_220632\공정_다중출력.xlsx --detail
결과: input | CLI 인자 오류: unrecognized arguments: --file C:\CLI-dev\work\gemma_mini_20261107_220632\공정_다중출력.xlsx. 'info' 명령에는 --file 옵션이 없어 같은 뜻으로 고쳤습니다 (--file -> --target). <문서에서 가져온 후보·문구 118자 생략>
```
