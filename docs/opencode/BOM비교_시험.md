# BOM 비교 시험 (집 PC, 사용자 조합 스킬)

기반 도구(office-cli)를 이어 쓰는 사용자 조합 스킬 `bom-compare`(0.2, 계층 BOM)를 집 PC에서 시험한다. 정답을 아는 샘플로 먼저 확인하고,
그다음 샘플을 회사 BOM과 비슷하게 고쳐 다시 해 본다.

| 파일 (저장소) | 배포 PC | 무엇 |
|---|---|---|
| `AIKIT/dev/office-cli/tools/bom-compare/` | `C:\AIKIT\dev\office-cli\tools\bom-compare\` | 스킬 원본 (`SKILL.md`, `scripts\compare.py`) |
| `AIKIT/dev/office-cli/tools/make_bom_samples.py` | `C:\AIKIT\dev\office-cli\tools\make_bom_samples.py` | 샘플 생성과 정답 확인 (`--check`) |

역할 나눔 (전역 지침 10장): 읽기·칠하기·메모·화면에 열기는 office-cli, 비교 기준은 `SKILL.md` "판단 기준",
두 표를 자리(상위 조립체 + Part Number)마다 맞대는 계산과 같은 자리 중복 검사는 스킬의 `scripts\compare.py` (기반 도구 범위 밖이라 스킬에 둠).

## 1. 준비
1. `C:\AIKIT`을 main 최신으로 다시 배치 (stats·duplicates가 제목 행 아래 머리글을 찾는 수정 포함).
2. 스킬 설치: `C:\AIKIT\dev\office-cli\tools\bom-compare` 폴더를 통째로 `C:\AIKIT\skills\bom-compare`로 복사 -> OpenCode 다시 시작.
3. 샘플 만들기:
   ```powershell
   python "C:\AIKIT\dev\office-cli\tools\make_bom_samples.py"
   ```
   `C:\AIKIT\dev\work\bom\`에 `BOM_A_Rev1.xlsx`(옛 판), `BOM_A_Rev2.xlsx`(새 판), `BOM_A_Rev2_중복.xlsx`, `BOM_A_정답.csv`.

샘플 모양: 1행 제목, 2행 빈 행, 3행 머리글(Level, Part Number, Description, Spec, Qty, Unit, Remark), 계층 BOM.
Level 0 최상위 조립체 `TA-1000` 1개 -> Level 1 중간 조립체 `SA-01`~`SA-12` -> Level 2 부품 (조립체마다 5~8개). 옛 판 77행(부품 자리 64개).
공통 부품(볼트 `P-0001`, 와셔 `P-0003`, 케이블 타이 `P-0006` 등)은 여러 조립체에 반복해서 쓰인다 (중복이 아님).
정답 15건 (추가 4, 삭제 4, 변경 7자리): `SA-03` 아래 `P-0001` Qty (다른 조립체의 `P-0001`은 그대로), `SA-05` 자체 Qty,
`P-0010` Spec·`P-0018` Description 변경(쓰인 자리마다 한 줄), `P-0012`를 `SA-04`에서 `SA-06`으로 옮김(삭제 + 추가), 앞자리 0(`0102` 삭제, `102` 추가) 등.
함정 8가지는 정답에 넣지 않는다 (걸리면 정답에 없는 결과가 나옴): 중간 조립체 순서, 조립체 안 부품 순서, Remark, Part Number 공백·소문자(` p-0004 `),
Qty `8`과 `8.0`, Spec 대소문자(`d12 l120`), 반복 부품, 앞자리 0. 목록은 `make_bom_samples.py`를 실행하면 나온다.

## 2. 시험 A: 명령을 사람이 직접 (스킬 순서가 맞는지)
`C:\AIKIT\dev\work\bom`에서 SKILL.md "순서" 1~4단계 명령을 차례로 실행한다 (`<...>`는 앞 단계 결과 JSON 값).
1~2단계는 Office 없이도 됨(닫힌 xlsx는 openpyxl로 읽음, pytest `test_bom_compare`): `compare.py`가 `추가 4, 삭제 4, 변경 7자리(7칸)`.
정답 확인:
```powershell
python "C:\AIKIT\dev\office-cli\tools\make_bom_samples.py" --check "C:\AIKIT\dev\work\bom\BOM_A_Rev2_비교용_비교결과.csv"
```
`[OK] 정답과 같습니다`이면 맞대기가 맞음. 4단계(엑셀 쓰기·칠하기·메모)는 실제 Excel이 필요하다. 볼 것:
- `BOM_A_Rev2_비교결과.xlsx`가 생기고 15줄 + 머리글 (열: 구분, 상위 조립체, Part Number, 바뀐 열, 옛 값, 새 값, 옛 행, 새 행).
- `BOM_A_Rev2.xlsx`에서 노란 칸 7개(D16, E21, C25, E31, E47, D58, C70), 연두 칸 4개(B30, B42, B55, B80),
  노란 칸마다 `[BOM 비교] Qty: 4 -> 6 (SA-03 아래, 옛 판 21행)` 같은 메모.
- 각 명령이 `VERIFIED:`인지 (`UNVERIFIED:`면 그 결과의 `unverified`를 기록).

## 3. 시험 B: OpenCode에게 요청 (모델이 스킬을 고르고 순서대로 하는지)
OpenCode 새 대화에서:
```text
C:\AIKIT\dev\work\bom\BOM_A_Rev1.xlsx 가 옛 판이고 BOM_A_Rev2.xlsx 가 새 판이야. 두 BOM 비교해줘.
```
볼 것 (O/X):
1. bom-compare 스킬을 고르는가 (excel-kit만 쓰거나 파이썬으로 직접 비교하지 않는가).
2. 1~3단계를 묻지 않고 이어서 하고, 결과 표를 보여 준 뒤 4단계 전에 "새 판이 바뀝니다"를 묻는가.
3. 결과 수를 compare.py `preview` 그대로 말하는가 (추가 4, 삭제 4, 변경 7자리). 표를 읽어 다시 세지 않는가.
4. "예" 뒤 칠하기·메모가 시험 A와 같은가. 옛 판(`BOM_A_Rev1.xlsx`)은 바뀌지 않았는가.
5. 함정 8가지를 바뀐 것으로 말하지 않는가. 특히 여러 조립체에 반복되는 부품을 "중복"이라 하지 않는가.

중복 판:
```text
C:\AIKIT\dev\work\bom\BOM_A_Rev1.xlsx 와 BOM_A_Rev2_중복.xlsx 비교해줘.
```
2단계에서 `SA-03` 아래 `P-0016`이 19행과 20행에 두 번 있다고 알리고 멈추는가 (결과 표·칠하기를 하지 않음).

## 4. 회사 BOM과 비슷하게 바꾸기
- 샘플만 바꿀 때: `make_bom_samples.py` 맨 위 `COLUMNS`(열 이름), `TITLE`, `PARTS`(부품 마스터), `ASSEMBLIES`(중간 조립체와 그 아래 부품),
  `CHANGES`(새 판에서 바뀌는 것 = 정답)를 고친다.
  회사 BOM 엑셀을 직접 고쳐서 써도 된다 (그때는 `--check` 대신 결과 표를 눈으로 본다).
- 열 이름이나 비교 기준이 바뀌면 세 곳을 같이 고친다:
  1. `compare.py` 맨 위 "기준": `LEVEL`(계층 열, 평면 BOM이면 `None`), `KEY`(Part Number 열), `COMPARE`(비교할 열), `NUMBER_COLS`, `IGNORE_CASE_COLS`
  2. `SKILL.md` "판단 기준" 문장 (모델이 사용자에게 설명하는 근거)
- 고친 스킬은 `C:\AIKIT\skills\bom-compare`에 다시 복사하고 OpenCode를 다시 시작한다. 스킬 하나 8KB 이하.

회사 PC에서 회사 BOM으로 맞출 때는 회사 PC OpenCode에 `C:\AIKIT\dev\office-cli\tools\BOM비교_회사데이터_가이드.md`를 읽게 한다
(복사본으로만 시험, 기준은 사용자 확인 뒤 고침, 정답을 아는 시험 판으로 먼저 확인, 회사 값 없이 보고).

## 5. 결과 남기기
시험 A·B의 O/X와 이상한 점(명령, 결과 JSON의 `agent_directive`·`unverified`)을 `docs/test_logs/집PC_BOM비교_<날짜>/`에 메모로 남긴다.
회사 BOM을 쓴 경우 파일 이름·품번은 옮기지 않는다 (AGENTS.md 작업 절차 4).
