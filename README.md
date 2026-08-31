# excel_proc

CSV를 서식 있는 엑셀로 변환하는 도구. 같은 일을 하는 두 가지 구현이 있다.

## index.html (권장)

`index.html`을 더블클릭해서 브라우저로 열면 끝이다. 설치도 빌드도 필요 없고,
인터넷 연결 없이도 동작한다 (ExcelJS 4.4.0이 파일 안에 인라인되어 있다).

- **입력**: 숫자 12개 열로 이루어진 CSV. 각 줄 끝의 빈 열(trailing comma)은 허용한다.
- **출력**: `<원본 파일명>_result.xlsx`

## CSV_to_Excel_변환기.xlsm (원본)

`index.html`이 이식된 원본 VBA 매크로 워크북이다. 당분간 계속 보관한다.
로직은 `modCSVImport.bas`에 있다.

## 테스트

```
node test/fidelity.js
```

`index.html`이 만드는 워크북이 `test/reference.xlsx`(VBA 도구로 만든 정답 셀 값·서식)와
셀 단위로 완전히 같은지 검증한다.

**`test/reference.xlsx`는 절대 다시 만들지 말 것.** 이 파일은 정답지다. 코드에 맞춰
정답지를 새로 뽑으면 테스트가 아무것도 증명하지 못하게 된다.
