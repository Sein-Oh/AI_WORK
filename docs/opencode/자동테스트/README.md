# OpenCode 종합 테스트 메시지 팩

agent_scenarios.py에서 만든 파일입니다. 직접 고치지 말고 agent_scenarios.py를 고친 뒤 `python grade_opencode.py --make-pack`으로 다시 만듭니다.

## 1. 도구 점검 (정답을 아는 샘플로 도구만 확인, 5~10분)

PowerShell에서 직접 실행하거나 Gemma에게 아래 메시지를 붙여넣습니다. 결과는 `C:\CLI-dev\beta_check_회사PC_<시각>.md`.

```powershell
python "C:\CLI-dev\run_beta_check.py" --label 회사PC
```

```text
python "C:\CLI-dev\run_beta_check.py" --label 회사PC --agent 를 시간 제한 1800000으로 1번만 실행하고, 결과 JSON의 passed, failed, skipped, failed_items, report_file을 그대로 알려줘.
```

## 2. Gemma 종합 테스트 (요청 처리 확인)

1. 시험 폴더 `C:\CLI-dev\work\gemma`가 이미 있으면 탐색기에서 지웁니다 (이전 모델이 남아 있으면 버전 번호가 달라집니다).
2. OpenCode에서 새 대화를 열고, 아래 묶음 파일의 메시지를 순서대로 붙여넣습니다.
   - 하나씩: 메시지 1개를 보내고 답을 받은 뒤 다음 메시지를 보냅니다 (권장).
   - 한 번에: 묶음 전체를 메시지 1개로 보냅니다 (긴 작업을 이어서 처리하는지 볼 때). 요청이 한 턴에 섞이므로 채점은 명령 순서로 나눈 참고 결과입니다.
   - 묶음은 S -> A -> M -> P -> O -> X 순서입니다. P와 R은 A에서 만든 모델을 씁니다. O는 S1이 만든 샘플을 닫힌 파일 경로로 쓰고 Excel·Word가 있어야 합니다.
   - 수동 준비가 필요한 요청(O3)은 묶음 파일의 '수동 준비'를 먼저 하고 하나씩 보냅니다. `run_all_tests.py --gemma`는 이 요청을 보내지 않습니다.
3. 끝나면 대화를 JSON으로 내보냅니다. PowerShell에서 (세션ID는 OpenCode 화면이나 `opencode session list`에서 확인):

```powershell
cmd /c "opencode export <세션ID> > chat.json"
```

   cmd의 `>`는 opencode가 쓴 UTF-8을 그대로 저장합니다. PowerShell의 `| Out-File`이나 `>`로 저장하면 한글이 깨지고 채점기가 [ERROR]로 멈춥니다.
   OpenCode 앱에 내보내기 메뉴가 있으면 그것으로 저장해도 됩니다.
4. `C:\CLI-dev`에서 채점합니다:

```powershell
python grade_opencode.py --export "<내보낸 대화.json>" --label 회사PC
```

   여러 대화로 나눠 했으면 `--export a.json b.json`처럼 함께 넣습니다.
   사내 LLM 참고 점수를 붙이려면 `$env:JUDGE_BASE_URL='http://<서버>/v1'` 설정 후 `--judge` (통과 여부에는 넣지 않음).
5. 만들어진 `grade_회사PC_<시각>.md`와 `beta_check_회사PC_<시각>.md`를 전달합니다.

Gemma가 질문하면 답해도 됩니다. 표시([ID])가 없는 메시지는 앞 요청에 이어진 대화로 채점합니다.

## 묶음

- [S 준비](S_준비.md): S0, S1
- [A 학습과 제외 행](A_학습과_제외_행.md): A1, A2, A3
- [M 표 보간과 반대 찾기 (ml_map)](M_표_보간과_반대_찾기.md): M1, M2, M3
- [P 예측과 모델 반대 찾기](P_예측과_모델_반대_찾기.md): P1, P2, P3
- R 보고서와 성능예측표: 뺌 (보고서·성능예측표는 개발 중이라 시험에서 뺌 (2026-11-08))
- [O office-cli 문서 작업 (닫힌 파일)](O_office-cli_문서_작업.md): O1, O2, O3, O4, O5
- [X 하면 안 되는 일](X_하면_안_되는_일.md): X1, X2
