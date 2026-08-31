# CSV → 서식 xlsx 변환기: Excel VBA → HTML SPA 마이그레이션 설계

작성일: 2026-08-31
대상: `CSV_to_Excel_변환기.xlsm` / `modCSVImport.bas` → `index.html`

## 1. 목적과 범위

### 1.1 현행

Excel 매크로 통합문서가 CSV 파일을 읽어, 정해진 레이아웃과 서식이 입혀진 `.xlsx` 파일로 다시 씁니다. 실행에 Excel과 매크로 사용 허용이 필요합니다.

### 1.2 이행 후

브라우저에서 열리는 단일 HTML 파일이 같은 일을 합니다. 추가로 변환 결과를 화면에 그대로 보여줍니다. Excel 설치가 필요 없습니다.

### 1.3 이번 범위에 포함

- CSV 파싱 (BOM, CRLF, 후행 빈 열 허용 — 현행 `ParseCSV`와 동일 판정)
- 레이아웃 계산 (현행 `BuildSheet`와 동일 결과)
- 화면 미리보기 — 배경색·병합·테두리까지 엑셀과 동일한 모양
- `.xlsx` 내보내기 — VBA 산출물과 셀 단위로 동일
- 여러 CSV 동시 처리 (현행 `RunBatch` 대체)

### 1.4 범위에서 제외

12열 고정 해제, 인쇄/PDF 전용 스타일, 드래그앤드롭, 다국어, zip 묶음 내려받기. 필요해지면 별도 작업으로 추가합니다.

### 1.5 기존 VBA 도구

`CSV_to_Excel_변환기.xlsm`과 `modCSVImport.bas`는 **손대지 않고 그대로 둡니다.** SPA가 검증을 통과한 뒤 폐기 여부를 따로 결정합니다. 그때까지 VBA 산출물이 SPA의 정답지 역할을 합니다 (4장).

## 2. 타당성 근거

설계 전에 프로브로 확인한 사실입니다.

| 항목 | 결과 |
|---|---|
| ExcelJS 4.4가 쓰는 채우기 XML | `<fgColor theme="5" tint="0.5999938962981048"/>` — VBA 산출물과 동일 |
| 병합 · 열 너비 · 행 높이 · 테두리 · 정렬 | 전부 재현됨 |
| **ExcelJS 내장 테마** | **Office 2007 테마 (accent2 = `C0504D`). 그대로 두면 색이 다름** |
| 테마 교체 (`wb._themes = {theme1: xml}`) | 정상 동작. accent2 `ED7D31`, accent4 `FFC000`, accent5 `5B9BD5` — 템플릿과 일치 |
| `file://` 에서의 File API · Blob 내려받기 | 동작. 네트워크 요청 없음 |

## 3. 아키텍처

### 3.1 핵심 결정: 모델 하나, 렌더러 둘

미리보기와 xlsx가 서로 어긋나는 것이 이 마이그레이션의 유일한 실질적 위험입니다. 레이아웃 로직을 **한 번만** 이식하고, 두 출력이 그 결과를 각자 그립니다.

```
CSV 텍스트 ──parseCSV──▶ number[n][12]
                            │
                            ▼  buildModel(rows, title, label)
                    ┌───────────────┐
                    │  레이아웃 모델  │   ← BuildSheet를 이식한 유일한 지점
                    └───────┬───────┘
                    ┌───────┴───────┐
              renderPreview      toXlsx
                    │               │
              HTML <table>      .xlsx Blob
```

검토했으나 버린 안:

- **렌더러 2개 독립 작성** — 그룹 병합과 사이드 테이블 계산이 두 벌이 되어 반드시 어긋납니다.
- **HTML을 먼저 그리고 DOM을 훑어 xlsx 생성** — CSS 색을 테마 색으로 역매핑해야 하고, 표 구조 변경에 취약합니다.

### 3.2 레이아웃 모델

```js
{
  cells:  [ {r, c, v, s} ],      // r, c 는 1-based (엑셀과 동일). v 는 number|string|null
  merges: [ {r1, c1, r2, c2} ],
  colWidths: { 16: 5.125 },      // 엑셀 저장값 기준
  rowHeight: 16.5,
  maxRow, maxCol
}
```

`s` 는 스타일 토큰이며 아래 표가 유일한 정의입니다. 두 렌더러가 같은 표를 읽으므로 색을 바꾸면 화면과 엑셀이 함께 바뀝니다.

| 토큰 | VBA | OOXML | 화면 색 | 용도 |
|---|---|---|---|---|
| `accent2` | `xlThemeColorAccent2`, tint 0.6 | `theme=5, tint=0.5999938962981048` | `#F8CBAD` | 제목, 머리글, A/B 그룹 |
| `accent5` | `xlThemeColorAccent5`, tint 0.6 | `theme=8, tint=0.5999938962981048` | `#BDD7EE` | Q2 |
| `accent4` | `xlThemeColorAccent4`, tint 0.8 | `theme=7, tint=0.7999816888943144` | `#FFF2CC` | R2 |
| `gray` | `xlThemeColorDark1`, tint −0.15 | `theme=0, tint=-0.1499984740745262` | `#D9D9D9` | 라벨, P열 |
| `data` | 채우기 없음 | — | 흰색 | 값 |

화면 색은 Excel이 해당 테마+tint를 실제로 렌더링한 RGB를 읽어 확정한 값입니다.

모든 셀에 공통 적용: 맑은 고딕 11, 가로·세로 가운데 정렬, 사방 얇은 테두리.

### 3.3 파일 구성

```
excel_proc/
  index.html                     ← 앱 전체 (더블클릭 실행)
  test/fidelity.js               ← 자동 검증 스크립트 (의존성 없음)
  test/reference.xlsx            ← VBA가 생성한 정답 파일 (고정 픽스처)
  random_data.csv                ← 기존 샘플, 그대로 사용
  CSV_to_Excel_변환기.xlsm       ← 변경 없음
  modCSVImport.bas               ← 변경 없음
```

`index.html` 안의 스크립트 블록 3개:

| 블록 | 내용 | 비고 |
|---|---|---|
| `<script>` | ExcelJS 4.4 `exceljs.bare.min.js` 인라인 | 862KB, 한 줄. 버전이 파일에 고정됨 |
| `<script id="app">` | `parseCSV` · `buildModel` · `toXlsx` | **DOM 접근 금지** (4.2 참조). `window.APP` 에 노출 |
| `<script id="ui">` | 파일 입력, 목록, 미리보기 렌더링, 내려받기 | DOM 담당 |

`app` 블록은 전역 `ExcelJS` 를 참조합니다. 브라우저에서는 인라인 번들이 그 전역을 채우고, 테스트에서는 같은 번들을 샌드박스에 올려 채웁니다 (4.2).

## 4. 검증

### 4.1 자동 검증 — `node test/fidelity.js`

`random_data.csv` 를 SPA 코드로 변환한 뒤, **VBA가 만든 `test/reference.xlsx` 와 셀 단위로 대조**합니다. 양쪽 모두 ExcelJS로 읽어 비교하므로 Excel 설치가 필요 없습니다.

비교 항목: 셀 값, 채우기(theme + tint), 테두리 4변, 폰트(이름·크기), 정렬, 병합 목록, 열 너비, 행 높이.

같은 스크립트에서 파서 판정도 확인합니다.

| 입력 | 기대 |
|---|---|
| `...,406` (12열) | 통과 |
| `...,406,` (후행 콤마) | 통과, 위와 동일 결과 |
| `...,406 ,   ` (콤마 + 공백) | 통과 |
| `...,406,,` (빈 열 2개) | 통과 |
| 11열 | 실패 |

### 4.2 단일 HTML 파일의 테스트 방법

테스트가 `index.html` 에서 **스크립트 블록을 추출해 `vm` 샌드박스에서 실행**합니다. 인라인된 ExcelJS 번들과 `app` 블록을 순서대로 넣으면 됩니다.

```js
const fs = require('fs'), vm = require('vm');
const html = fs.readFileSync('index.html', 'utf8');
const sandbox = { console, setTimeout, clearTimeout, Buffer, process, TextEncoder, TextDecoder, URL };
sandbox.window = sandbox; sandbox.self = sandbox; sandbox.global = sandbox;
vm.createContext(sandbox);
for (const m of html.matchAll(/<script[^>]*>([\s\S]*?)<\/script>/g)) vm.runInContext(m[1], sandbox);
// → sandbox.ExcelJS, sandbox.APP.{parseCSV, buildModel, toXlsx}
```

이 방식의 이점:

- **의존성 설치가 없습니다.** `package.json` 도 `node_modules` 도 만들지 않습니다
- 별도로 받은 라이브러리가 아니라 **실제로 배포되는 번들 그 자체**를 검증합니다
- 소스가 중복되지 않고 빌드 단계도 생기지 않습니다

`ui` 블록은 DOM을 건드리므로 샌드박스에서 실행하면 실패합니다. 따라서 `app` 블록에는 DOM 접근을 두지 않고, 추출 시 `ui` 블록은 `id` 로 걸러냅니다.

프로브로 확인 완료: 샌드박스에서 `ExcelJS.Workbook` 생성, `writeBuffer()`, 그리고 `xlsx.load()` 로 읽어들인 채우기가 `{theme:5, tint:0.6}` 으로 보존되는 것까지 동작합니다 (4.1의 대조가 성립하는 근거).

### 4.3 픽스처 생성

`test/reference.xlsx` 는 현행 `.xlsm` 을 Excel COM으로 구동해 한 번 생성한 뒤 커밋합니다. 제목 `8/28 111G`, 라벨 `XX`, 입력 `random_data.csv`. 이후 재생성하지 않습니다 — 정답이 움직이면 검증이 무의미해집니다.

### 4.4 수동 확인

`index.html` 을 열어 `random_data.csv` 를 넣고, 미리보기를 Excel에서 연 `reference.xlsx` 와 나란히 놓고 눈으로 대조합니다.

## 5. 레이아웃 사양

`BuildSheet` 에서 그대로 옮깁니다. n = CSV 행 수, 열은 항상 12개.

열 배치: A=1, B=2, 데이터 12열 = C..N(3..14), O=15(빈 칸막이), P=16, Q=17, R=18.

| 위치 | 내용 | 스타일 |
|---|---|---|
| A1:O1 (병합) | 제목 | `accent2` |
| Q1:R1 (병합) | 라벨 | `gray` |
| C2..N2 | `A`,`B` 반복 (홀수번째 = `A`) | `accent2` |
| Q2 / R2 | `A` / `B` | `accent5` / `accent4` |

본문, i = 1..n:

- `headerRow = 3 + (i-1)*2`, `dataRow = headerRow + 1`
- `headerRow` 의 C..N = 1..12 (`accent2`), O = 빈 값 (`accent2`)
- `dataRow` 의 C..N = 데이터 값 (`data`), O = 빈 값 (`data`)

그룹 라벨:

- B열 — CSV 2행씩 묶어 병합, 값은 그룹 번호, `accent2`
- A열 — CSV 4행씩 묶어 병합, 값은 그룹 번호, `accent2`
- 그룹 g 의 범위: `rFrom = 3 + (first-1)*2`, `rTo = 3 + (last-1)*2 + 1` (`last` 는 n을 넘지 않음)

사이드 표, 총 `6n` 행:

- P1, P2 — 값 없이 `gray`
- Q 목록 = 홀수 열(1,3,5,7,9,11) 값 전체, R 목록 = 짝수 열 값 전체. 수집 순서는 `k = 1..6` 이 바깥, `i = 1..n` 이 안쪽
- 두 목록을 각각 **내림차순 정렬**
- `i = 1..6n` 에 대해 행 `2+i`: P = i (`gray`), Q = Q목록[i] (`data`), R = R목록[i] (`data`)

VBA는 T/U열을 임시 작업 공간으로 쓴 뒤 지웠습니다. 모델에서는 JS 정렬로 대체하므로 T/U는 존재하지 않습니다.

치수:

- P열 너비 — VBA `ColumnWidth = 4.5`, 엑셀 저장값 `5.125`. **ExcelJS에는 저장값 `5.125` 를 씁니다.** 나머지 열은 기본값
- 모든 행 높이 16.5 (행 1 ~ `2 + 6n`)

## 6. 화면과 조작

```
[CSV 선택 (다중)]   제목 [8/28 111G]   라벨 [XX]

파일                        ┌─ 미리보기 ─────────────────┐
▸ random_data.csv  ✓ 12행   │ ████ 8/28 111G █████│░ XX ░│
  b.csv            ✓  8행   │ █A█│█B█│█1█│█2█│ …        │
  c.csv            ✗ 3행이  │  1 │ 1 │465│996│ …        │
                     11열   └────────────────────────────┘
[선택 내려받기]  [전체 내려받기]
```

- 제목·라벨은 입력 즉시 미리보기에 반영 (VBA는 InputBox라 재실행이 필요했음)
- 기본값은 현행과 동일: 제목 `8/28 111G`, 라벨 `XX`
- 내려받기 파일명은 현행과 동일: `<원본이름>_result.xlsx`
- 전체 내려받기는 개별 xlsx를 순차 저장. zip으로 묶지 않습니다 (라이브러리 절약)
- 실패한 파일은 목록에서 ✗ 로 표시되고 내려받기 대상에서 제외

미리보기 렌더링:

- 모델을 `<table>` 로 그리고, 병합은 `colspan`/`rowspan` 으로 표현하며 덮인 셀은 건너뜁니다
- 모델에 없는 셀은 값·테두리 없이 비웁니다
- 열 너비 환산: `px = round(저장너비 × 7) + 5` (기본 8.43 → 64px, P열 5.125 → 41px). 상수 `7` 과 `5` 는 한 곳에 모아 조정 가능하게 둡니다

## 7. 오류 처리

| 상황 | 현행 VBA | SPA |
|---|---|---|
| 12열이 아닌 행 | "각 행이 정확히 12개의 값을 가져야 합니다" (위치 불명) | **몇 번째 행인지 표시** |
| 숫자가 아닌 값 | `CDbl` 런타임 오류 | 행·열을 지목, 해당 파일만 ✗ |
| 빈 파일 / 유효 행 없음 | 실패 | ✗ 표시 |
| 한 파일이 실패 | 나머지 파일은 계속 처리 | 동일 |

판정 기준 자체는 VBA와 같습니다. 실패 위치를 알려주는 것만 개선했습니다.

인코딩은 UTF-8로 읽습니다 (`File.text()`). CP949 파일이라도 숫자와 콤마는 ASCII이므로 파싱에 영향이 없습니다.

## 8. 알려진 위험과 대응

| 위험 | 영향 | 대응 |
|---|---|---|
| `wb._themes` 는 ExcelJS 공개 API가 아님 | 라이브러리 업그레이드 시 색이 틀어질 수 있음 | ExcelJS를 인라인하므로 버전이 파일에 고정됨. `ponytail:` 주석으로 대체 경로(ARGB 직접 지정)를 명시 |
| Normal 스타일 폰트 차이 (맑은 고딕 vs Calibri) | P열 너비가 미세하게 다르게 렌더링될 수 있음 | 4.1 검증이 열 너비를 비교하므로 검출됨. 검출되면 저장 너비 수치로 보정 |
| 값 없이 스타일만 있는 셀 (O열, P1/P2) | ExcelJS가 빈 셀의 스타일을 생략할 수 있음 | 4.1 검증이 해당 셀의 채우기를 비교. 누락되면 명시적 스타일 지정으로 처리 |
| 862KB 인라인 | `index.html` 이 약 1MB | 한 줄 minified라 이후 diff에 영향 없음. 최초 커밋에만 나타남 |
| 열 너비 → px 환산이 폰트에 의존 | 미리보기 열 폭이 엑셀과 약간 다를 수 있음 | 환산 상수를 한 곳에 모아 조정 가능하게 (6장) |

## 9. 완료 기준

1. `node test/fidelity.js` 통과 — `random_data.csv` 산출물이 `reference.xlsx` 와 모든 비교 항목에서 일치
2. 파서 엣지케이스 5종 통과 (4.1)
3. `index.html` 을 `file://` 로 열어 CSV 다중 선택 → 미리보기 → 개별/전체 내려받기가 동작
4. 미리보기와 Excel에서 연 산출물이 육안으로 동일
5. 기존 `.xlsm` / `.bas` 무변경
