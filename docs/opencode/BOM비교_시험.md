# BOM 비교 시험 (집 PC, 사용자 조합 스킬)

기반 도구(office-cli)를 이어 쓰는 사용자 조합 스킬 `bom-compare`를 집 PC에서 시험한다. 정답을 아는 샘플로 먼저 확인하고,
그다음 샘플을 회사 BOM과 비슷하게 고쳐 다시 해 본다.

| 파일 (저장소) | 배포 PC | 무엇 |
|---|---|---|
| `AIKIT/dev/office-cli/tools/bom-compare/` | `C:\AIKIT\dev\office-cli\tools\bom-compare\` | 스킬 원본 (`SKILL.md`, `scripts\compare.py`) |
| `AIKIT/dev/office-cli/tools/make_bom_samples.py` | `C:\AIKIT\dev\office-cli\tools\make_bom_samples.py` | 샘플 생성과 정답 확인 (`--check`) |

역할 나눔 (전역 지침 10장): 읽기·칠하기·메모·화면에 열기는 office-cli, 비교 기준은 `SKILL.md` "판단 기준",
두 표를 행 단위로 맞대는 계산은 스킬의 `scripts\compare.py` (기반 도구 범위 밖이라 스킬에 둠).

## 1. 준비
1. `C:\AIKIT`을 main 최신으로 다시 배치 (stats·duplicates가 제목 행 아래 머리글을 찾는 수정 포함).
2. 스킬 설치: `C:\AIKIT\dev\office-cli\tools\bom-compare` 폴더를 통째로 `C:\AIKIT\skills\bom-compare`로 복사 -> OpenCode 다시 시작.
3. 샘플 만들기:
   ```powershell
   python "C:\AIKIT\dev\office-cli\tools\make_bom_samples.py"
   ```
   `C:\AIKIT\dev\work\bom\`에 `BOM_A_Rev1.xlsx`(옛 판), `BOM_A_Rev2.xlsx`(새 판), `BOM_A_Rev2_중복.xlsx`, `BOM_A_정답.csv`.

샘플 모양: 1행 제목, 2행 빈 행, 3행 머리글(번호, 품번, 품명, 규격, 수량, 단위, 비고), 옛 판 42개·새 판 41개 부품.
정답 11건 (추가 2, 삭제 3, 변경 6). 함정 7가지는 정답에 넣지 않는다 (걸리면 정답에 없는 결과가 나옴):
행 순서 바뀜, 번호 다시 매김, 비고 바뀜, 품번 공백·소문자(` p-0006 `), 수량 `3`과 `3.0`, 규격 대소문자(`M3X15`), 앞자리 0(`0102`와 `102`는 다른 품번).

## 2. 시험 A: 명령을 사람이 직접 (스킬 순서가 맞는지)
`C:\AIKIT\dev\work\bom`에서 SKILL.md "순서" 1~5단계 명령을 차례로 실행한다 (`<...>`는 앞 단계 결과 JSON 값).
1~3단계는 클라우드에서 확인함(닫힌 xlsx는 Office 없이 읽음): `compare.py`가 `추가 2, 삭제 3, 변경 6개 부품(6칸)`.
정답 확인:
```powershell
python "C:\AIKIT\dev\office-cli\tools\make_bom_samples.py" --check "C:\AIKIT\dev\work\bom\BOM_A_Rev2_비교용_비교결과.csv"
```
`[OK] 정답과 같습니다`이면 맞대기가 맞음. 5단계(엑셀 쓰기·칠하기·메모)는 실제 Excel이 필요해 집 PC에서 처음 확인한다. 볼 것:
- `BOM_A_Rev2_비교결과.xlsx`가 생기고 11줄 + 머리글.
- `BOM_A_Rev2.xlsx`에서 노란 칸 6개(E6, D11, E14, C17, E28, D42), 연두 칸 2개(B43, B44), 노란 칸마다 `[BOM 비교] 수량: 4 -> 9 (옛 판 6행)` 같은 메모.
- 각 명령이 `VERIFIED:`인지 (`UNVERIFIED:`면 그 결과의 `unverified`를 기록).

## 3. 시험 B: OpenCode에게 요청 (모델이 스킬을 고르고 순서대로 하는지)
OpenCode 새 대화에서:
```text
C:\AIKIT\dev\work\bom\BOM_A_Rev1.xlsx 가 옛 판이고 BOM_A_Rev2.xlsx 가 새 판이야. 두 BOM 비교해줘.
```
볼 것 (O/X):
1. bom-compare 스킬을 고르는가 (excel-kit만 쓰거나 파이썬으로 직접 비교하지 않는가).
2. 1~4단계를 묻지 않고 이어서 하고, 결과 표를 보여 준 뒤 5단계 전에 "새 판이 바뀝니다"를 묻는가.
3. 결과 수를 compare.py `preview` 그대로 말하는가 (추가 2, 삭제 3, 변경 6). 표를 읽어 다시 세지 않는가.
4. "예" 뒤 칠하기·메모가 시험 A와 같은가. 옛 판(`BOM_A_Rev1.xlsx`)은 바뀌지 않았는가.
5. 함정 7가지를 바뀐 것으로 말하지 않는가.

중복 판:
```text
C:\AIKIT\dev\work\bom\BOM_A_Rev1.xlsx 와 BOM_A_Rev2_중복.xlsx 비교해줘.
```
2단계에서 `P-0003`이 6행과 45행에 두 번 있다고 알리고 멈추는가 (맞대기·칠하기를 하지 않음).

## 4. 회사 BOM과 비슷하게 바꾸기
- 샘플만 바꿀 때: `make_bom_samples.py` 맨 위 `COLUMNS`(열 이름), `TITLE`, `PARTS`(옛 판 부품), `CHANGES`(새 판에서 바뀌는 것 = 정답)를 고친다.
  회사 BOM 엑셀을 직접 고쳐서 써도 된다 (그때는 `--check` 대신 결과 표를 눈으로 본다).
- 열 이름이나 비교 기준이 바뀌면 세 곳을 같이 고친다:
  1. `compare.py` 맨 위 "기준": `KEY`(품번 열), `COMPARE`(비교할 열), `NUMBER_COLS`, `IGNORE_CASE_COLS`
  2. `SKILL.md` "판단 기준" 문장 (모델이 사용자에게 설명하는 근거)
  3. `SKILL.md` 2단계의 `--columns 품번` (품번 열 이름이 다르면)
- 고친 스킬은 `C:\AIKIT\skills\bom-compare`에 다시 복사하고 OpenCode를 다시 시작한다. 스킬 하나 8KB 이하.

## 5. 결과 남기기
시험 A·B의 O/X와 이상한 점(명령, 결과 JSON의 `agent_directive`·`unverified`)을 `docs/test_logs/집PC_BOM비교_<날짜>/`에 메모로 남긴다.
회사 BOM을 쓴 경우 파일 이름·품번은 옮기지 않는다 (AGENTS.md 작업 절차 4).
