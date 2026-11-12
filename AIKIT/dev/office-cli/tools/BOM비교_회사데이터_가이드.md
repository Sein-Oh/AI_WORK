# BOM 비교 스킬을 회사 BOM에 맞추기 (회사 PC OpenCode용 가이드)

이 문서는 회사 PC의 OpenCode 모델이 읽고 그대로 따라 하는 절차다. 사람은 맨 아래 "사용자가 붙여넣을 첫 메시지"만 보면 된다.

## 0. 목표와 지킬 것
목표: 샘플로 만든 스킬 `bom-compare`(`C:\AIKIT\skills\bom-compare`)가 회사 BOM 두 판을 바르게 비교하도록 "기준"만 맞춘다.

지킬 것 (하나라도 어기면 멈추고 사용자에게 묻는다):
1. 원본 BOM 파일은 열기만 한다. 시험과 칠하기는 `C:\AIKIT\data\bom_work`에 복사한 파일로만 한다.
2. 고치는 파일은 둘뿐이다: `C:\AIKIT\skills\bom-compare\scripts\compare.py`의 "기준" 상수 블록, `C:\AIKIT\skills\bom-compare\SKILL.md`의 "판단 기준" 문장과 `description`.
   그 밖의 코드는 고치지 않는다. 기준으로 안 되면 6장대로 보고하고 멈춘다.
3. 비교 기준(어느 열로 같은 부품을 가를지, 어느 열을 비교할지)은 업무 판단이다. 근거(읽은 머리글과 앞 행)를 보여 주고 사용자가 "예"라고 한 뒤에 고친다.
4. 회사 BOM의 파일 이름·품번·값은 보고서나 저장소에 옮기지 않는다. 6장 보고는 열 이름과 수만 쓴다.
5. 명령은 `python "C:\AIKIT\..."` 전체 경로로, 이 문서에 적힌 그대로 실행한다. 결과 처리는 전역 지침 4장을 따른다.

## 1. 복사본 만들기
사용자에게 옛 판과 새 판 파일 경로를 받는다. 그다음 실행:
```powershell
New-Item -ItemType Directory -Force "C:\AIKIT\data\bom_work"
Copy-Item "<옛 판 원본.xlsx>" "C:\AIKIT\data\bom_work\옛판.xlsx"
Copy-Item "<새 판 원본.xlsx>" "C:\AIKIT\data\bom_work\새판.xlsx"
```
이후 모든 명령은 `C:\AIKIT\data\bom_work\옛판.xlsx`, `C:\AIKIT\data\bom_work\새판.xlsx`로 한다.

## 2. BOM 모양 읽기 (읽기만, 묻지 않고 진행)
```powershell
python "C:\AIKIT\office-cli\excel.py" info --target "C:\AIKIT\data\bom_work\새판.xlsx"
python "C:\AIKIT\office-cli\excel.py" read --target "C:\AIKIT\data\bom_work\새판.xlsx" --range A1:Z20
```
시트가 여러 개면 BOM이 있는 시트를 사용자에게 묻고, `read`에 `--sheet "<시트>"`를 붙여 다시 읽는다.
읽은 `values`에서 아래 표를 채운다. 모르는 칸은 비워 두고 사용자에게 묻는다.

| 확인할 것 | 샘플 값 (지금 compare.py) | 회사 BOM에서 찾을 것 |
|---|---|---|
| 머리글 행 | 3행 (위는 제목·빈 행) | 열 이름이 모인 행. 위 15행 안이어야 함 |
| 계층 열 `LEVEL` | `Level` (0, 1, 2) | 숫자(0, 1, 2) 또는 점 표기(.1, ..2)인 열. 없으면 평면 BOM. `1.1`·`1.2.1` 같은 번호식이면 지금 compare.py가 잘못 읽으니(`1.1`을 1로) 6장대로 보고 |
| 부품 열 `KEY` | `Part Number` | 품번 열 |
| 비교할 열 `COMPARE` | `Description`, `Spec`, `Qty` | 품명, 규격, 수량에 해당하는 열 |
| 숫자로 비교할 열 `NUMBER_COLS` | `Qty` | 수량 열 (3과 3.0을 같게) |
| 대소문자·공백 무시 열 `IGNORE_CASE_COLS` | `Spec` | 규격 열 |

## 3. 사용자에게 기준 확인 (업무 판단)
2장의 표와 읽은 앞 행 3~5줄을 보여 주고 묻는다:
"이 BOM은 <계층 열>로 상위 조립체를 정하고, 상위 조립체와 <부품 열>이 같으면 같은 자리로 봅니다. 비교할 열은 <열들>입니다. 이대로 할까요?"
- "예"이면 4장으로.
- 사용자가 열을 바꾸면 바꾼 대로 다시 보여 주고 "예"를 받는다.
- 계층 열이 없으면(평면 BOM) "품번만으로 맞댑니다"로 묻는다 (`LEVEL = None`).

## 4. 기준 고치기
`C:\AIKIT\skills\bom-compare\scripts\compare.py`의 아래 블록만 고친다 (따옴표 안 열 이름은 머리글과 글자까지 같게):
```python
LEVEL = "Level"
KEY = "Part Number"
COMPARE = ["Description", "Spec", "Qty"]
NUMBER_COLS = {"Qty"}
IGNORE_CASE_COLS = {"Spec"}
```
`C:\AIKIT\skills\bom-compare\SKILL.md`에서 고칠 곳 (열 이름만 바꾸고 문장 구조는 그대로):
- `description`의 `Part Number`, `Description, Spec, Qty`
- "판단 기준" 1~4번의 열 이름 (`Level`, `Part Number`, `Description`, `Spec`, `Qty`, `Unit`, `Remark`)
- 고친 뒤 SKILL.md가 8KB(8192바이트) 이하인지 본다. SKILL.md를 고쳤으면 OpenCode를 다시 시작해야 모델이 새 기준을 읽는다고 사용자에게 알린다.

## 5. 시험 (정답을 아는 판으로 먼저, 그다음 실제 두 판)
### 5-1. 같은 판끼리 (0건이어야 함)
```powershell
python "C:\AIKIT\office-cli\excel.py" read --target "C:\AIKIT\data\bom_work\새판.xlsx" --save "C:\AIKIT\data\bom_work\새판_비교용.csv"
python "C:\AIKIT\skills\bom-compare\scripts\compare.py" --old "C:\AIKIT\data\bom_work\새판_비교용.csv" --new "C:\AIKIT\data\bom_work\새판_비교용.csv" --new-range "<read 결과 range>" --old-range "<read 결과 range>"
```
기대: `VERIFIED: 0건`. 중복으로 멈추면 6장 "중복"으로.

### 5-2. 정답을 아는 시험 판 (값 2칸을 바꾼 복사본)
```powershell
Copy-Item "C:\AIKIT\data\bom_work\새판.xlsx" "C:\AIKIT\data\bom_work\시험판.xlsx"
```
5-1에서 읽은 행 중 부품 행 둘을 고른다. 하나는 수량 칸, 하나는 규격 칸. 바꿀 칸과 새 값을 보여 주고 "예"를 받은 뒤 실행 (복사본이어도 `--overwrite`는 전역 지침 5장):
```powershell
python "C:\AIKIT\office-cli\excel.py" write --target "C:\AIKIT\data\bom_work\시험판.xlsx" --range <수량 칸> --value <다른 수> --overwrite
python "C:\AIKIT\office-cli\excel.py" write --target "C:\AIKIT\data\bom_work\시험판.xlsx" --range <규격 칸> --value "<다른 규격>" --overwrite
python "C:\AIKIT\office-cli\excel.py" read --target "C:\AIKIT\data\bom_work\시험판.xlsx" --save "C:\AIKIT\data\bom_work\시험판_비교용.csv"
python "C:\AIKIT\skills\bom-compare\scripts\compare.py" --old "C:\AIKIT\data\bom_work\새판_비교용.csv" --new "C:\AIKIT\data\bom_work\시험판_비교용.csv" --new-range "<시험판 range>" --old-range "<새판 range>"
```
기대: `preview`가 `추가 0, 삭제 0, 변경 2자리(2칸)`, `result_file`에 바꾼 두 칸의 상위 조립체·열이 맞게 나온다. 다르면 6장.

### 5-3. 실제 두 판
SKILL.md "순서" 1~3단계를 `옛판.xlsx`, `새판.xlsx`로 한다 (4단계 칠하기는 사용자가 원할 때만, 복사본에만). 2장에서 시트를 정했으면 이 장의 모든 `read`에 `--sheet "<시트>"`를 붙인다.
`preview`와 결과 표 앞 10줄을 보여 주고, 사용자가 아는 변경(설계 변경 통보 등) 몇 건이 결과에 있는지 사용자에게 확인받는다.

## 6. 문제가 생기면 (고치지 말고 보고)
| 증상 | 뜻 | 할 일 |
|---|---|---|
| "머리글(...)을 찾지 못했습니다" | 4장 열 이름이 머리글과 다름, 또는 머리글이 15행 아래 | 머리글 글자를 다시 읽어 4장을 고침 (공백·괄호까지 같게). 15행 아래면 보고 |
| "Level 값 ...을(를) 읽지 못했습니다" | 계층 열이 숫자·점 표기가 아님 (예: 1.2.1, 글자) | 보고 |
| "같은 상위 조립체 아래 같은 ...가 두 번" | 실제 중복이거나, 같은 부품을 위치(Item No 등)별로 여러 줄 쓰거나, 같은 하위 조립체가 여러 상위 조립체에 쓰임(지금은 바로 위 조립체만 봄) | 행 번호와 그 위 조립체들을 보여 주고 어느 쪽인지 묻는다. 실제 중복이 아니면 보고 |
| 5-1이 0건이 아님 | 같은 판인데 차이가 나옴 (기준 열 해석 문제) | 결과 앞 5줄의 "바뀐 열"을 보고 |
| 5-2가 2칸이 아님 | 기준이 BOM에 맞지 않음 | 결과 전체 수와 "바뀐 열" 목록을 보고 |

보고 형식 (개발자에게 전달, 회사 값·파일 이름·품번은 쓰지 않음):
```text
[BOM 비교 회사 적용 보고]
머리글 행: <n>행, 열: <머리글 목록>, 시트: 여러 개/하나
기준: LEVEL=<값>, KEY=<값>, COMPARE=<값>, NUMBER_COLS=<값>, IGNORE_CASE_COLS=<값>
5-1 같은 판: <preview 또는 오류 message 앞부분>
5-2 시험 판: <preview>
5-3 실제 두 판: <preview>, 사용자 확인 <O/X>
문제: <6장 증상, 계층 열 표기 예(값 대신 모양만: 예 "1.2.1")>
```

## 사용자가 붙여넣을 첫 메시지
```text
C:\AIKIT\dev\office-cli\tools\BOM비교_회사데이터_가이드.md 를 읽고 그대로 해 줘.
옛 판은 <옛 판 파일 전체 경로>, 새 판은 <새 판 파일 전체 경로>야. 원본은 바꾸지 말고, 기준을 바꾸기 전에는 나에게 물어 줘.
```
