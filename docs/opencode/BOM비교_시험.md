# BOM 비교 스킬 만들기 연습 (집 PC)

회사에서는 사내 LLM이 [`BOM비교_스킬만들기_가이드.md`](../../AIKIT/dev/office-cli/tools/BOM비교_스킬만들기_가이드.md)를 읽고
회사 BOM을 보며 사용자와 함께 비교 스킬(`bom-compare`)을 만든다. 우리 저장소에는 맞대기 스크립트를 두지 않는다
(회사 BOM을 본 적이 없으므로 기준과 스크립트는 회사에서 만든다, 사용자 결정 2026-11-12).
집 PC에서는 정답을 아는 샘플로 같은 가이드를 미리 따라 해 보고, 가이드가 막히는 곳을 찾는다.

## 1. 준비
1. `C:\AIKIT`을 main 최신으로 다시 배치.
2. 샘플 만들기: `python "C:\AIKIT\dev\office-cli\tools\make_bom_samples.py"`
   -> `C:\AIKIT\dev\work\bom\`에 `BOM_A_Rev1.xlsx`(옛 판), `BOM_A_Rev2.xlsx`(새 판), `BOM_A_Rev2_중복.xlsx`, `BOM_A_정답.csv`.
   모양: 1행 제목, 2행 빈 행, 3행 머리글(Level, Part Number, Description, Spec, Qty, Unit, Remark), Level 0~2,
   중간 조립체 12개, 여러 조립체에 반복되는 부품, 정답 15건, 함정 8가지 (실행하면 목록이 나온다).

## 2. 연습
OpenCode 새 대화에 가이드 맨 아래 첫 메시지를 붙여넣는다 (옛 판 `BOM_A_Rev1.xlsx`, 새 판 `BOM_A_Rev2.xlsx`).
3장 질문에는 샘플 기준으로 답한다: 같은 자리 = 상위 조립체 경로 + Part Number, 비교 열 Description·Spec·Qty, 여러 줄은 실제 중복으로 멈춤.
만들어진 스크립트의 결과를 정답과 비교:
```powershell
python "C:\AIKIT\dev\office-cli\tools\make_bom_samples.py" --check "C:\AIKIT\data\bom_work\새판_비교용_비교결과.csv"
```
(`--check`는 결과 CSV의 앞 네 열 `구분, 상위 조립체, 품번, 바뀐 열`만 본다.) 중복 판은 새 판 자리에 `BOM_A_Rev2_중복.xlsx`를 주고 멈추는지 본다.

볼 것 (O/X): 1) 가이드 순서대로 하는가 (2장 도구로 모양을 읽고, 3장에서 묻고, 스크립트·SKILL.md를 저장 전에 보여 줌)
2) 스크립트가 엑셀을 직접 열지 않고 `read --save` CSV를 쓰는가 3) 5장 시험 셋을 하는가 4) `--check`가 `[OK]` 5) 만든 SKILL.md로 새 대화에서 비교가 되는가.
막힌 곳은 가이드 7장 보고 형식으로 `docs/test_logs/집PC_BOM비교_<날짜>/`에 남긴다.
