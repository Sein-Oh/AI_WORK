# 사용법 단순화 실기 점검: 집PC

- 시각: 2026-11-12T17:39:36, 걸린 시간 162초
- PC: Windows-11-10.0.26200-SP0, Python 3.13.13, 셸: PowerShell (Windows)
- 결과: PASS 34, FAIL 0, SKIP 0

| 항목 | 내용 | 결과 | 상세 |
|---|---|---|---|
| M1 | 열 이름 오타 -> [후보] 1) + confirm_commands[0]을 셸로 실행 -> 학습 | PASS |  |
| M2 | 예전 줄임말 --target -> 인자 오류 + [추천] --targets 전체 명령 -> 학습 | PASS |  |
| M3 | --description -> [추천] --desc 전체 명령 -> 학습 | PASS |  |
| M4 | --dtypes 형식 오타 -> [추천] 전체 명령 -> 학습 | PASS |  |
| M5 | 두 줄 머리글 엑셀 -> [추천] --header-row 2 전체 명령 -> 학습 | PASS |  |
| M6 | 예측 --model <데이터 파일> -> 최종 버전 자동 (model_chosen) | PASS | 최종 버전 v4 (마지막 학습) |
| M7 | --save .txt -> [추천] .csv 전체 명령 -> 예측 저장 | PASS |  |
| M8 | --dry-run: v1 교체 미리보기 (replaced_files, destructive) | PASS |  |
| M8b | 미리보기의 confirm_commands[0]을 셸로 실행 -> v1 교체 | PASS |  |
| M9 | 성능예측표 --model <데이터 파일> + --targets 오타 -> [후보] 1) -> 셸로 실행 | PASS |  |
| N1 | --features 없이 학습 -> 도구가 입력을 고름 (ID·상수·타겟 복사 열은 이유와 함께 뺌) | PASS | 제외=['상수', '시료번호', '효율복사'], 고른 입력=['반응온도(°C)', '소비전력(kW)', '압축압력(bar)', '촉매비율(wt%)', '수분함량(%)'] |
| N1b | 학습 결과의 next_commands[0](--features가 적힌 재학습)을 셸로 실행 | PASS |  |
| N2 | --suggest 3 -> 실험 조건 3개와 실험계획 엑셀 | PASS | 다음 실험 제안 3개 [불확실한 곳 (예측 90% 범위가 넓고 실측과 먼 곳)]: 1) 반응온도=220.2, 소비전력=42.8, 압축압력=3.64, 촉매비율=1.427, 수분함량=29.78 -> 생산효율 44.7142 (90% 폭 12.3342) / 2) 반응온도=277.5, 소비전력=88.5, 압축압력=3.8, 촉매비율=0.638, 수분함량=28.51 - |
| N2b | --suggest 3 --maximize -> 좋은 쪽 실험 조건 | PASS |  |
| N2c | --suggest 99(상한 넘음) -> [추천] 전체 명령 -> 실행 | PASS |  |
| N3 | 한 건 예측 -> 입력별 영향(explain)과 preview의 '높인 입력/낮춘 입력' | PASS | [예측 완료] 입력: 직접 입력 1건 / 예측: 생산효율(%) 50.7665 (90% 범위 48.382 ~ 53.2946, 실측 비교값 51.709 (차이 -0.9425)) / 입력별 영향 생산효율: 높인 입력: 수분함량 +0.0982 / 낮춘 입력: 반응온도 -1.24, 촉매비율 -0.858 |
| N4 | 실측이 있는 새 데이터(값이 크게 다름) -> actual_check와 재학습 알림 | PASS | {"생산효율(%)": {"column": "생산효율(%)", "rows": 30, "rmse": 15.2325, "mae": 15.171, "r2": -3.107, "validation": {"rmse": 1.59442, "mae": 1.30781, "r2": 0.9393}, "text": "생산효율 RMSE 15.23 (학습 검증 1.594), R2 -3 |
| N4b | 재학습 명령을 셸로 실행 -> 새 데이터로 학습 | PASS |  |
| N4c | 학습과 같은 데이터 -> 재학습 알림 없음 | PASS | {"생산효율(%)": {"column": "생산효율(%)", "rows": 120, "rmse": 1.4953, "mae": 1.2171, "r2": 0.9466, "validation": {"rmse": 1.59442, "mae": 1.30781, "r2": 0.9393}, "text": "생산효율 RMSE 1.495 (학습 검증 1.594), R2 0. |
| O1 | 닫힌 엑셀 없는 시트 -> [후보] 1) 요약 + confirm_commands[0]을 셸로 실행 -> 읽기 | PASS |  |
| O2 | --where 조건 사이 쉼표 -> [추천] ; 전체 명령 -> 읽기 (1건) | PASS |  |
| O3 | transfer 옛 --from/--to -> 인자 오류 + [추천] --from-app/--to-app (실행은 안 함) | PASS | python C:\AIKIT\office-cli\office.py transfer --from-app excel --range A1:B2 --to-app word |
| O4 | 닫힌 엑셀 --cells 쉼표 -> [추천] ; 전체 명령 -> 셸로 실행 (Excel) | PASS |  |
| O5 | 닫힌 워드 --slots 쉼표 -> [추천] ; 전체 명령 -> 셸로 실행 (Word) | PASS |  |
| O6 | 닫힌 PPT --cards '|' -> [추천] ; 전체 명령 -> 셸로 실행 (PowerPoint) | PASS |  |
| C1 | ml-kit 옵션 오타 --data-ro -> [후보] 1) --data-row + confirm_commands[0] -> 학습 | PASS |  |
| C2 | reader --query 옛 AND -> [추천] ; 전체 명령 -> 검색 (1건 이상) | PASS |  |
| C3 | 닫힌 워드 같은 제목 둘 -> [후보] 1) 2) + confirm_commands[0](--occurrence 1) -> 읽기 | PASS |  |
| E1 | 닫힌 엑셀 빈 범위 읽기 -> count 0 + '찾은 항목 없음' | PASS | {'status': 'ok', 'tool': 'excel', 'command': 'read', 'preview': '0건: 요약!H10:I12에 값이 있는 칸이 없습니다 (3행 x 2열 모두 빈 칸)', 'target_resolved': {'app': 'Excel', 'name': '보고 $자료 (1).xlsx', 'path': 'C:\\AIKIT\\dev |
| D1 | 닫힌 엑셀 replace --dry-run -> 그대로 -> confirm_commands[0] -> 바뀜 (Excel) | PASS |  |
| D2 | 닫힌 엑셀 sheet delete --dry-run -> 그대로 -> confirm_commands[0] -> 시트 지움 (Excel) | PASS |  |
| D3 | 닫힌 워드 replace --dry-run -> 그대로 -> confirm_commands[0] -> 바뀜 (Word) | PASS |  |
| D4 | 닫힌 PPT slide delete --dry-run -> 그대로 -> confirm_commands[0] -> 지움 (PowerPoint) | PASS |  |
| K1 | C:\AIKIT\skills에 기본 스킬 12개 (setup.py 뒤) | PASS | 없는 스킬 [] |
