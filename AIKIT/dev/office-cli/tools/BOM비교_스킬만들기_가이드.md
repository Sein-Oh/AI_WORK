# BOM 비교 스킬 만들기 (사내 LLM용 가이드)

이 문서는 회사 PC의 OpenCode 모델이 읽고 따라 하는 절차다. 사람은 맨 아래 "사용자가 붙여넣을 첫 메시지"만 보면 된다.
목표: 사용자와 함께 회사 BOM 두 판(옛 판, 새 판)을 비교하는 조합 스킬 `bom-compare`를 `C:\AIKIT\skills\bom-compare`에 만든다.
회사 BOM은 이 가이드를 쓴 사람이 본 적이 없다. 모양은 2장에서 도구로 읽어 알아내고, 기준은 3장에서 사용자에게 정해 받는다.

## 0. 역할과 지킬 것
| 일 | 누가 |
|---|---|
| BOM 읽기, CSV 저장, 결과 표 쓰기, 칸 칠하기, 메모 | office-cli 기본 도구 (`excel.py`, 명령은 이 문서에 적힌 그대로) |
| 두 판을 행 단위로 맞대는 계산 | 이 가이드로 만드는 `C:\AIKIT\skills\bom-compare\scripts\compare.py` (전역 지침 8·10장) |
| 같은 부품의 기준, 비교할 열, 여러 줄 처리 | 사용자 (업무 판단, 근거를 보여 주고 "예"를 받는다) |

1. 원본 BOM은 읽기만 한다. 시험과 칠하기는 `C:\AIKIT\data\bom_work`에 복사한 파일로만 한다.
2. 스크립트는 엑셀 파일을 직접 열지 않는다. 읽기는 `excel.py read --save`가 만든 CSV, 쓰기·칠하기·메모는 `excel.py` 명령으로 한다.
3. 회사 BOM의 파일 이름·품번·값은 보고서나 다른 파일에 옮기지 않는다. 7장 보고는 열 이름과 수만 쓴다.
4. 스크립트와 SKILL.md는 저장하기 전에 내용을 보여 주고 "예"를 받는다 (전역 지침 10장). 고친 뒤에는 5장 시험을 처음부터 다시 한다.
5. 같은 오류로 스크립트를 2번 고쳐도 안 되면 멈추고 7장대로 보고한다 (전역 지침 8장 5번).

## 1. 복사본 만들기
사용자에게 옛 판과 새 판 파일 경로를 받는다.
```powershell
New-Item -ItemType Directory -Force "C:\AIKIT\data\bom_work"
Copy-Item "<옛 판 원본.xlsx>" "C:\AIKIT\data\bom_work\옛판.xlsx"
Copy-Item "<새 판 원본.xlsx>" "C:\AIKIT\data\bom_work\새판.xlsx"
```

## 2. BOM 모양 읽기 (읽기만, 묻지 않고 진행)
```powershell
python "C:\AIKIT\office-cli\excel.py" info --target "C:\AIKIT\data\bom_work\새판.xlsx"
python "C:\AIKIT\office-cli\excel.py" read --target "C:\AIKIT\data\bom_work\새판.xlsx" --range A1:Z30
```
시트가 여러 개면 BOM 시트를 사용자에게 묻고, 이후 모든 `read`·`style`·`comment`에 `--sheet "<시트>"`를 붙인다. 읽은 `values`로 아래 표를 채운다.

| 확인할 것 | 볼 곳 |
|---|---|
| 머리글 행 | 열 이름이 모인 행 (위에 제목·빈 행이 있을 수 있음) |
| 계층 열과 표기 | 숫자(1, 2, 3 ...)인지, 다른 표기(`.1`, `1.2.1`)인지. 계층 열이 없으면 평면 BOM |
| 품번 열 | 부품을 가르는 열 |
| 비교할 열 | 품명, 규격, 수량 등 (사용자가 정함) |
| 위치 번호 열 | Item No, 위치, Ref 같은 열이 있는지 |

같은 품번이 여러 줄인지 본다 (부품이 여러 곳에 쓰이는지, 같은 조립체 안에 여러 줄인지 알기 위해):
```powershell
python "C:\AIKIT\office-cli\excel.py" duplicates --target "C:\AIKIT\data\bom_work\새판.xlsx" --columns "<품번 열>"
```
하위 조립체 하나가 쓰인 곳을 보려면 (`<하위 조립체 품번>`은 위 결과나 2장 read에서 계층 중간에 있는 품번):
```powershell
python "C:\AIKIT\office-cli\excel.py" read --target "C:\AIKIT\data\bom_work\새판.xlsx" --where "<품번 열> = <하위 조립체 품번>"
```

## 3. 기준 정하기 (업무 판단, 사용자에게 묻기)
2장에서 읽은 앞 행과 아래 질문을 보여 주고 답을 받는다. 괄호는 권하는 답이다.
1. 같은 부품의 기준: "상위 조립체와 품번이 같으면 같은 자리로 보고 비교할까요?"
   - 하위 조립체가 여러 곳에 쓰이면(2장 `--where` 결과가 여러 행) 상위 조립체는 바로 위 하나가 아니라 맨 위부터의 경로로 한다
     (권함. 바로 위만 보면 같은 하위 조립체 아래 부품이 "두 번 나온다"고 잘못 겹친다).
   - 계층 열이 없으면 품번만으로 맞댄다.
2. 비교할 열과 비교 방법: 수량은 숫자로(3 = 3.0), 규격은 공백·대소문자 무시(권함), 그 밖의 열과 행 순서는 비교하지 않음.
3. 같은 자리(경로 + 품번)가 한 판에 여러 줄일 때: 행 번호를 보여 주고 고르게 한다. 예:
   ```text
   Level 3  SA-03  구동부 조립
   Level 4  P-0001 볼트 M4x10  Qty 4  (위치 10, 모터 고정)
   Level 4  P-0001 볼트 M4x10  Qty 2  (위치 30, 커버 고정)
   ```
   1) 실제 중복이라 멈추고 알림  2) 수량을 더해 한 줄로 보고 비교  3) 위치 번호 열까지 같아야 같은 자리
4. 앞자리 0: 0102와 102는 다른 품번으로 본다(권함).

## 4. 맞대기 스크립트 만들기 (`scripts\compare.py`)
3장의 답을 스크립트 맨 위 상수로 모으고(열 이름, 비교할 열, 숫자 열, 여러 줄 처리), 사용자가 기준을 바꿀 때 상수만 고치게 한다.
- 입력: `--old`, `--new` (두 판을 `excel.py read --save`로 저장한 CSV), `--new-range`, `--old-range` (그 read 결과의 `range` 값 그대로).
  CSV 첫 줄은 시트의 `range` 첫 칸이다. 빈 행은 빈 줄로 남는다. 그래서 엑셀 행 = `range` 첫 행 + CSV 줄 번호(0부터), 열도 같은 방식.
  CSV는 `utf-8-sig`로 읽는다. 머리글은 위 15줄 안에서 3장의 열 이름이 모두 있는 줄로 찾는다.
- 계층: 행을 위에서부터 읽으며 자기보다 계층 값이 작은 가장 가까운 위 행들을 쌓아 경로를 만든다. 읽지 못한 계층 값은 몇 번째 줄인지 알리고 멈춘다.
- 품번은 앞뒤 공백·대소문자를 무시해 맞대고, 앞자리 0은 그대로 둔다.
- 출력 (새 판 CSV와 같은 폴더):
  - `<새 판>_비교결과.csv`: 열 순서 `구분, 상위 조립체, 품번, 바뀐 열, 옛 값, 새 값, 옛 행, 새 행`. 구분은 추가·삭제·변경. 변경은 열마다 한 줄, 같은 자리는 넣지 않음
  - `<새 판>_칠할칸_변경.csv`, `<새 판>_칠할칸_추가.csv`: 머리글 `칸`, 줄마다 칸 주소 하나 (예: E21)
  - `<새 판>_메모.csv`: 머리글 `칸, 메모`, 예: `E21, [BOM 비교] Qty: 4 -> 6 (SA-03 아래, 옛 판 21행)`
- 출력 마지막 줄은 JSON 1줄: `status`(ok/error), `preview`(예: `추가 4, 삭제 4, 변경 7자리(7칸)`), `counts`, 위 파일 경로들.
  멈출 때는 `"status": "error"`와 `message`(무엇이 몇 행에서). 진행 글은 stderr.

## 5. 시험 (정답을 아는 것부터)
두 판을 CSV로 저장한다 (`--range`를 붙이지 않는다. 결과의 `output_file`·`range`를 스크립트에 그대로 넘긴다):
```powershell
python "C:\AIKIT\office-cli\excel.py" read --target "C:\AIKIT\data\bom_work\새판.xlsx" --save "C:\AIKIT\data\bom_work\새판_비교용.csv"
python "C:\AIKIT\office-cli\excel.py" read --target "C:\AIKIT\data\bom_work\옛판.xlsx" --save "C:\AIKIT\data\bom_work\옛판_비교용.csv"
```
1. 같은 판끼리 (`--old`, `--new` 모두 새판 CSV): 0건이어야 한다. 아니면 기준이나 계층 읽기가 틀렸다.
2. 값 2칸을 바꾼 시험 판: `새판.xlsx`를 `시험판.xlsx`로 복사하고, 부품 행 둘의 수량 칸과 규격 칸을 바꾼다. 바꿀 칸과 값을 보여 주고 "예"를 받은 뒤:
   ```powershell
   python "C:\AIKIT\office-cli\excel.py" write --target "C:\AIKIT\data\bom_work\시험판.xlsx" --range <수량 칸> --value <다른 수> --overwrite
   python "C:\AIKIT\office-cli\excel.py" write --target "C:\AIKIT\data\bom_work\시험판.xlsx" --range <규격 칸> --value "<다른 규격>" --overwrite
   python "C:\AIKIT\office-cli\excel.py" read --target "C:\AIKIT\data\bom_work\시험판.xlsx" --save "C:\AIKIT\data\bom_work\시험판_비교용.csv"
   ```
   새판 CSV와 맞대면 변경 2칸, 상위 조립체와 열이 맞아야 한다. 여러 곳에 쓰이는 하위 조립체 안의 칸을 하나 고르면 그 자리 하나만 나와야 한다.
3. 실제 두 판 (옛판, 새판): `preview`와 결과 앞 10줄을 보여 주고, 사용자가 아는 변경(설계 변경 통보 등)이 결과에 있는지 확인받는다.

## 6. SKILL.md 쓰기
5장을 모두 통과하면 `C:\AIKIT\skills\bom-compare\SKILL.md`를 쓴다 (frontmatter `name: bom-compare`, `description`에 "BOM 비교" 요청에 쓴다고, 8KB 이하).
- "판단 기준": 3장의 답을 문장으로 (모델이 사용자에게 설명할 근거).
- "순서": 1) 두 판 `read --save` 2) `compare.py` (멈추면 `message`를 알리고 끝) 3) `preview`와 결과 앞 10줄을 보여 주고
  "결과 표를 새 엑셀로 만들고 새 판에 칠하고 메모를 달까요? (새 판 파일이 바뀝니다)"를 묻기 4) "예"이면 아래 명령 5) 원하면 화면에 열기.
  ```text
  python "C:\AIKIT\office-cli\excel.py" write --target "<새 판 폴더>\<새 판 이름>_비교결과.xlsx" --range A1 --file "<result 파일>"
  python "C:\AIKIT\office-cli\excel.py" style --target "<새 판.xlsx>" --sheet "<시트>" --file "<칠할칸_변경 파일>" --bg yellow
  python "C:\AIKIT\office-cli\excel.py" style --target "<새 판.xlsx>" --sheet "<시트>" --file "<칠할칸_추가 파일>" --bg lightgreen
  python "C:\AIKIT\office-cli\excel.py" comment --action add --target "<새 판.xlsx>" --sheet "<시트>" --file "<메모 파일>"
  python "C:\AIKIT\office-cli\office.py" open --file "<새 판.xlsx>"
  ```
- "하지 않는 것": 옛 판은 바꾸지 않음, 기준을 대화 중에 바꾸지 않음, 결과 수는 `preview` 그대로.
- 다 쓴 뒤 OpenCode를 다시 시작해야 새 스킬이 보인다고 사용자에게 알린다.

## 7. 보고 (개발자에게 전달, 회사 값·파일 이름·품번은 쓰지 않음)
```text
[BOM 비교 스킬 만들기 보고]
모양: 머리글 <n>행, 열 <머리글 목록>, 시트 <하나/여러 개>, 계층 표기 <모양만, 예: 1~5 숫자>
기준: 같은 자리 <경로+품번 등>, 비교 열 <목록>, 여러 줄 <멈춤/합계/위치 번호>
시험: 같은 판 <preview>, 시험 판 <preview>, 실제 두 판 <preview>, 사용자 확인 <O/X>
막힌 곳: <도구 명령과 message 앞부분, 스크립트를 고친 횟수>
```

## 사용자가 붙여넣을 첫 메시지
```text
C:\AIKIT\dev\office-cli\tools\BOM비교_스킬만들기_가이드.md 를 읽고 그대로 해 줘.
옛 판은 <옛 판 파일 전체 경로>, 새 판은 <새 판 파일 전체 경로>야. 원본은 바꾸지 말고, 기준을 정하거나 파일을 저장하기 전에는 나에게 물어 줘.
```
