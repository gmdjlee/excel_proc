# excel_proc

CSV를 서식 있는 엑셀로 변환하는 도구. 같은 일을 하는 세 가지 구현이 있다.

## index.html (권장)

`index.html`을 더블클릭해서 브라우저로 열면 끝이다. 설치도 빌드도 필요 없고,
인터넷 연결 없이도 동작한다 (ExcelJS 4.4.0이 파일 안에 인라인되어 있다).

- **입력**: 숫자로만 이루어진 CSV. 열 개수는 파일의 첫 줄에서 자동으로 감지하며, 같은
  파일 안의 모든 줄은 그 개수와 같아야 한다. 각 줄 끝의 빈 열(trailing comma)은 허용한다.
- **출력**: `<원본 파일명>_result.xlsx`

## Electron 데스크톱 앱 (사내망 등 브라우저 파일 첨부가 막힌 환경용)

`index.html`은 손대지 않고 그대로 담아 `.exe`로 패키징한 것뿐이다. 일부 사내
보안 정책은 브라우저 페이지의 **URL이 첨부 허용 목록에 등록되어 있는지**를
검사해서 `<input type="file">` 사용을 막는데, 이 검사는 페이지 URL을 대상으로
하므로 URL 자체가 없는 독립 실행 파일에는 적용되지 않는다.

```
npm install       # 개발 PC에서 한 번. electron/electron-builder 를 받는다
npm start         # 개발 중 바로 실행해서 확인 (electron .)
npm run build     # release/ 에 포터블 exe 를 만든다
```

완성된 `release/CSV-Excel-Converter 1.0.0.exe` 하나만 사용 환경에 전달하면
된다 — 그 PC에는 node/npm이 전혀 필요 없다. 약 100MB로 `index.html` 단독
버전보다 훨씬 크므로, 메일보다는 파일 서버/공유 폴더로 배포하는 편이 낫다.

`main.js`가 하는 일은 창 하나를 띄우고 `index.html`을 그대로 여는 것뿐이라,
CSV 파싱·레이아웃·xlsx 내보내기 로직은 브라우저판과 완전히 동일하다.

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
