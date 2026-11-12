---
name: bom-compare
description: 두 판의 BOM(엑셀 부품표, Level 계층) 비교. 상위 조립체와 Part Number 기준으로 추가·삭제·변경(Description, Spec, Qty)을 찾아 결과 표를 만들고, 사용자가 원하면 새 판에 바뀐 칸을 칠하고 메모를 단다. "BOM 비교해줘", "두 판 부품표 차이" 요청에 사용. 사용자 조합 스킬(기본 도구 excel.py를 이어 씀).
license: Proprietary
compatibility: opencode
metadata:
  version: "0.2"
  platform: windows
---

# bom-compare (사용자 조합 스킬)

## 판단 기준 (이 스킬의 업무 규칙. 바꾸려면 사용자가 이 문단과 scripts\compare.py 맨 위 "기준"을 함께 고친다)
1. 같은 자리: 상위 조립체와 Part Number가 모두 같으면 같은 자리다. 상위 조립체는 Level로 정한다 (자기보다 Level이 작은 가장 가까운 위 행).
   같은 부품이 여러 조립체에 쓰이는 것은 정상이며 자리마다 따로 비교한다. Part Number의 앞뒤 공백과 대소문자는 무시하고, 앞자리 0은 다른 값으로 본다 (0102와 102는 다른 품번).
2. 비교할 열: Description, Spec, Qty. 그 밖의 열(Unit, Remark 등)과 행 순서는 비교하지 않는다.
3. Qty는 숫자로 비교한다 (3과 3.0은 같음). Spec은 공백과 대소문자를 무시한다 (M3x10 = M3X10).
4. 같은 판에서 같은 상위 조립체 아래 같은 Part Number가 두 번 이상 있으면 비교하지 않고 그 자리와 행을 사용자에게 알린다 (어느 줄이 맞는지는 사용자가 정한다).
5. 결과 구분: 추가(새 판에만), 삭제(옛 판에만), 변경(비교 열 값이 다름, 열마다 한 줄). 다른 조립체로 옮긴 부품은 옛 자리 삭제와 새 자리 추가로 나온다. 같은 자리는 결과에 넣지 않는다.

## 순서
명령은 그대로 실행하고 결과 처리는 전역 지침 4장, 승인은 5장을 따른다. 1~3단계는 읽기만 하므로 묻지 않고 이어서 한다.
1. 두 판을 CSV로 저장한다 (열 고르기 옵션을 붙이지 않는다):
   `python "C:\AIKIT\office-cli\excel.py" read --target "<옛 판.xlsx>" --save "<옛 판 이름>_비교용.csv"`
   `python "C:\AIKIT\office-cli\excel.py" read --target "<새 판.xlsx>" --save "<새 판 이름>_비교용.csv"`
   결과의 `output_file`, `range`, `sheet`를 다음 단계에 그대로 쓴다.
2. 맞댄다 (판단 기준 4의 중복도 여기서 본다):
   `python "C:\AIKIT\skills\bom-compare\scripts\compare.py" --old "<1단계 옛 판 output_file>" --new "<1단계 새 판 output_file>" --new-range "<1단계 새 판 range>" --old-range "<1단계 옛 판 range>"`
   중복으로 멈추면(`STOP_RETRY:`) `message`(자리와 엑셀 행)를 그대로 알리고 여기서 끝낸다.
3. `preview`(추가·삭제·변경 수)와 `result_file`의 앞 10줄을 표로 보여 준다. 0건이면 그대로 알리고 끝낸다.
   이어서 묻는다: "결과 표를 새 엑셀 파일로 만들고, 새 판에 바뀐 칸을 칠하고 메모를 달까요? (새 판 파일이 바뀝니다)"
4. "예"이면 차례로 실행한다:
   `python "C:\AIKIT\office-cli\excel.py" write --target "<새 판 폴더>\<새 판 이름>_비교결과.xlsx" --range A1 --file "<2단계 result_file>"`
   `python "C:\AIKIT\office-cli\excel.py" style --target "<새 판.xlsx>" --sheet "<1단계 새 판 sheet>" --file "<2단계 changed_cells_file>" --bg yellow`
   `python "C:\AIKIT\office-cli\excel.py" style --target "<새 판.xlsx>" --sheet "<1단계 새 판 sheet>" --file "<2단계 added_cells_file>" --bg lightgreen`
   `python "C:\AIKIT\office-cli\excel.py" comment --action add --target "<새 판.xlsx>" --sheet "<1단계 새 판 sheet>" --file "<2단계 memo_file>"`
   칠할 칸이나 메모가 0개인 파일(`changed_cells`·`added_cells`가 0)은 그 명령을 건너뛴다.
5. 사용자가 보겠다고 하면 연다: `python "C:\AIKIT\office-cli\office.py" open --file "<새 판.xlsx>"`

## 하지 않는 것
- 옛 판은 바꾸지 않는다. 새 판은 4단계("예" 뒤)의 칠하기와 메모만 바꾼다. 잘못 바꿨으면 `python "C:\AIKIT\office-cli\office.py" undo`.
- 판단 기준에 없는 열을 비교하거나 기준을 대화 중에 바꾸지 않는다. 사용자가 기준을 바꾸고 싶다고 하면 이 파일과 compare.py를 고칠지 묻는다.
- 결과 표를 읽어 직접 세거나 다시 맞대지 않는다. 수는 compare.py의 `preview`를 그대로 쓴다.
