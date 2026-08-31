# CSV → 엑셀 변환기 SPA 마이그레이션 구현 계획

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Excel VBA 변환기를 더블클릭으로 열리는 단일 `index.html` 로 옮긴다. 화면에 결과를 그대로 보여주고, VBA 산출물과 셀 단위로 같은 `.xlsx` 를 내보낸다.

**Architecture:** `BuildSheet` 를 레이아웃 모델로 **한 번만** 이식하고, HTML 미리보기와 xlsx 내보내기가 그 모델을 각자 렌더링한다. 스타일 토큰 하나가 CSS 색과 `{theme, tint}` 를 함께 소유하므로 두 출력이 어긋날 수 없다. 검증은 VBA가 만든 고정 픽스처와의 셀 단위 대조다.

**Tech Stack:** 순수 브라우저 JS (빌드 도구 없음), ExcelJS 4.4.0 (`exceljs.bare.min.js` 인라인), node `vm` 모듈 (테스트, 의존성 없음), Excel COM + PowerShell (픽스처 생성 1회)

**Spec:** `docs/superpowers/specs/2026-08-31-csv-spa-migration-design.md`

## Global Constraints

모든 태스크에 적용된다. 값은 스펙에서 그대로 옮긴 것이다.

- **결과물은 파일 하나** — `index.html`. `package.json`, `node_modules`, 빌드 스크립트를 만들지 않는다.
- **기존 VBA 파일 무변경** — `CSV_to_Excel_변환기.xlsm`, `modCSVImport.bas` 를 수정하지 않는다.
- **데이터 열 수는 12 고정.**
- **스타일 토큰과 값** — 이 표가 유일한 정의다. 다른 곳에서 색을 하드코딩하지 않는다.

  | 토큰 | 화면 색 | OOXML theme | OOXML tint |
  |---|---|---|---|
  | `accent2` | `#F8CBAD` | `5` | `0.5999938962981048` |
  | `accent5` | `#BDD7EE` | `8` | `0.5999938962981048` |
  | `accent4` | `#FFF2CC` | `7` | `0.7999816888943144` |
  | `gray` | `#D9D9D9` | `0` | `-0.1499984740745262` |
  | `data` | `#FFFFFF` | 채우기 없음 | — |

- **모든 셀 공통 서식** — 맑은 고딕 11pt, 가로·세로 가운데 정렬, 사방 `thin` 테두리.
- **P열 너비는 저장값 `5.125`** (VBA `ColumnWidth = 4.5` 에 대응). 나머지 열은 기본값.
- **모든 행 높이 16.5pt**, 행 1 부터 `2 + 6n` 까지 (n = CSV 행 수).
- **내려받기 파일명** — `<원본파일이름>_result.xlsx`.
- **기본 입력값** — 제목 `8/28 111G`, 라벨 `XX`.
- **테스트 명령은 하나** — `node test/fidelity.js`. 태스크마다 이 파일에 검증 구획을 덧붙인다.

---

## 파일 구성

| 파일 | 책임 |
|---|---|
| `index.html` | 앱 전체. 스크립트 블록 3개 (ExcelJS 번들 / `app` 순수 로직 / `ui` DOM 글루) |
| `test/fidelity.js` | 전 구간 검증. `index.html` 에서 스크립트를 뽑아 `vm` 샌드박스에서 실행 |
| `test/reference.xlsx` | VBA가 만든 정답 픽스처. 최초 1회 생성 후 고정 |
| `test/make-reference.ps1` | 픽스처 생성 스크립트. 출처 기록용, 재실행하지 않음 |
| `random_data.csv` | 기존 샘플. 변경 없음 |

`app` 블록에는 **DOM 접근을 두지 않는다.** 샌드박스에서 실행되어야 하기 때문이다. `ui` 블록만 `document` 를 만진다.

---

### Task 1: 부트스트랩 — 픽스처, HTML 뼈대, 테스트 하네스

앱 로직을 쓰기 전에 "테스트가 `index.html` 안의 코드를 실행할 수 있다" 를 먼저 세운다. 이후 모든 태스크가 이 하네스 위에서 돈다.

**Files:**
- Create: `index.html`
- Create: `test/fidelity.js`
- Create: `test/make-reference.ps1`
- Create: `test/reference.xlsx` (스크립트가 생성)

**Interfaces:**
- Consumes: 없음
- Produces:
  - `test/fidelity.js` 의 `loadApp()` → `sandbox` 객체. `sandbox.ExcelJS` 와 `sandbox.APP` 를 노출
  - `test/fidelity.js` 의 `check(name, cond, detail)` → 실패를 세는 단언 헬퍼
  - `index.html` 의 `<script id="app">` 블록이 `window.APP` 객체를 노출

- [ ] **Step 1: 픽스처 생성 스크립트를 쓴다**

`test/make-reference.ps1` 을 만든다. 주석을 영어로 쓰는 이유는 PowerShell 5.1 이 BOM 없는 UTF-8 스크립트를 ANSI로 읽어 한글이 깨지기 때문이다. 파일명도 리터럴 대신 glob 으로 찾는다 (통합문서 이름에 한글이 있어 같은 이유로 깨진다).

```powershell
# Generates test/reference.xlsx from the existing VBA workbook.
# Run ONCE. Do not regenerate - this file is the answer key for the SPA's
# fidelity test, and a moving answer key makes the test meaningless.
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$xlsm = (Get-ChildItem $root -Filter '*.xlsm').FullName
$csv  = Join-Path $root 'random_data.csv'
$out  = Join-Path $PSScriptRoot 'reference.xlsx'
Remove-Item $out -Force -ErrorAction SilentlyContinue

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.AutomationSecurity = 1
$wb = $xl.Workbooks.Open($xlsm)
$ok = $xl.Run('BuildFromCSV', $csv, '8/28 111G', 'XX', $out)
"BuildFromCSV = $ok"
$wb.Close($false)
$xl.Quit()
```

- [ ] **Step 2: 픽스처를 생성한다**

Run: `powershell -ExecutionPolicy Bypass -File test/make-reference.ps1`

Expected: `BuildFromCSV = True` 가 출력되고 `test/reference.xlsx` 가 생긴다.

확인: `ls -l test/reference.xlsx` — 13KB 내외.

Excel이 없거나 COM 이 막혀 실패하면 여기서 멈추고 보고한다. 픽스처 없이는 7번 태스크를 할 수 없다.

- [ ] **Step 3: ExcelJS 번들과 테마 XML을 준비한다**

저장소 밖 임시 디렉터리에서 받는다. 저장소에는 `package.json` 을 만들지 않는다.

```bash
TMP=$(mktemp -d) && cd "$TMP" && npm init -y >/dev/null && npm i exceljs@4.4.0 >/dev/null
cp node_modules/exceljs/dist/exceljs.bare.min.js /tmp/exceljs.bare.min.js
echo "bundle bytes: $(wc -c < /tmp/exceljs.bare.min.js)"      # 862631 근처
echo "script-tag hazard: $(grep -c '</script' /tmp/exceljs.bare.min.js)"   # 반드시 0
```

`script-tag hazard` 가 0 이 아니면 인라인이 HTML을 깨뜨린다. 그때는 번들을 `.replace('</script', '<\\/script')` 로 치환해 넣는다. 4.4.0 에서는 0 이다.

테마 XML을 JS 문자열 리터럴 한 줄로 뽑는다. 저장소 루트에서:

```bash
python -c "
import json, zipfile
x = zipfile.ZipFile('test/reference.xlsx').read('xl/theme/theme1.xml').decode('utf-8')
open('/tmp/theme-line.txt','w',encoding='utf-8').write('const OFFICE_THEME_XML = ' + json.dumps(x) + ';')
"
wc -c /tmp/theme-line.txt    # 8500 근처
```

- [ ] **Step 4: 실패하는 테스트를 쓴다**

`test/fidelity.js` 를 만든다.

```js
'use strict';
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = path.join(__dirname, '..');

let failures = 0;
function check(name, cond, detail) {
  if (cond) {
    console.log('  ok   ' + name);
  } else {
    console.log('  FAIL ' + name + (detail ? ' — ' + detail : ''));
    failures++;
  }
}
function section(title) {
  console.log('\n== ' + title + ' ==');
}

// index.html 에서 ui 블록을 뺀 스크립트를 순서대로 샌드박스에 올린다.
// ui 블록은 document 를 만지므로 샌드박스에서 실행하면 실패한다.
function loadApp() {
  const html = fs.readFileSync(path.join(ROOT, 'index.html'), 'utf8');
  const sandbox = {
    console, setTimeout, clearTimeout, Buffer, process,
    TextEncoder, TextDecoder, URL,
  };
  sandbox.window = sandbox;
  sandbox.self = sandbox;
  sandbox.global = sandbox;
  vm.createContext(sandbox);
  const re = /<script(?![^>]*\bid="ui")[^>]*>([\s\S]*?)<\/script>/g;
  for (const m of html.matchAll(re)) vm.runInContext(m[1], sandbox);
  return sandbox;
}

async function main() {
  section('Task 1: 하네스');
  const app = loadApp();
  check('ExcelJS 가 샌드박스에 올라온다', typeof app.ExcelJS === 'object', 'typeof=' + typeof app.ExcelJS);
  check('APP 이 노출된다', typeof app.APP === 'object', 'typeof=' + typeof app.APP);

  console.log(failures === 0 ? '\n전체 통과' : '\n실패 ' + failures + '건');
  process.exit(failures ? 1 : 0);
}

main().catch(e => { console.error(e); process.exit(1); });
```

- [ ] **Step 5: 실패를 확인한다**

Run: `node test/fidelity.js`

Expected: FAIL — `ENOENT: no such file or directory, open '...index.html'`

- [ ] **Step 6: index.html 뼈대를 만든다**

아래 내용으로 `index.html` 을 만든다. `BUNDLE` 자리에는 `/tmp/exceljs.bare.min.js` 의 내용을, `THEME_LINE` 자리에는 `/tmp/theme-line.txt` 의 한 줄을 그대로 넣는다.

```html
<!doctype html>
<html lang="ko">
<head>
<meta charset="utf-8">
<title>CSV → 엑셀 변환기</title>
<style>
</style>
</head>
<body>

<script>BUNDLE</script>

<script id="app">
'use strict';
(function (root) {

THEME_LINE

root.APP = {
  OFFICE_THEME_XML: OFFICE_THEME_XML,
};

})(typeof window !== 'undefined' ? window : this);
</script>

<script id="ui">
</script>
</body>
</html>
```

- [ ] **Step 7: 테스트를 통과시킨다**

Run: `node test/fidelity.js`

Expected:
```
== Task 1: 하네스 ==
  ok   ExcelJS 가 샌드박스에 올라온다
  ok   APP 이 노출된다

전체 통과
```

- [ ] **Step 8: 커밋**

```bash
git add index.html test/fidelity.js test/make-reference.ps1 test/reference.xlsx
git commit -m "Add SPA skeleton with sandboxed test harness and VBA reference fixture"
```

---

### Task 2: `parseCSV`

`modCSVImport.bas` 의 `ParseCSV` 를 옮긴다. 판정 기준은 동일하게 두되, 실패한 행 번호를 알려준다.

**Files:**
- Modify: `index.html` — `app` 블록
- Modify: `test/fidelity.js`

**Interfaces:**
- Consumes: Task 1 의 `loadApp()`, `check()`, `section()`
- Produces: `APP.parseCSV(text)` → `{rows: number[][]}` 또는 `{error: string}`. `rows[i]` 는 길이 12 의 숫자 배열

- [ ] **Step 1: 실패하는 테스트를 쓴다**

`test/fidelity.js` 의 `main()` 안, Task 1 구획 뒤에 넣는다.

```js
  section('Task 2: parseCSV');
  const P = app.APP.parseCSV;
  const base = '465,996,97,596,776,652,442,626,637,96,744,406';

  check('12열 한 줄', JSON.stringify(P(base).rows) === JSON.stringify([[465,996,97,596,776,652,442,626,637,96,744,406]]));
  check('후행 콤마', JSON.stringify(P(base + ',').rows) === JSON.stringify(P(base).rows));
  check('콤마 + 공백', JSON.stringify(P(base + ' ,   ').rows) === JSON.stringify(P(base).rows));
  check('빈 열 2개', JSON.stringify(P(base + ',,').rows) === JSON.stringify(P(base).rows));
  check('BOM 과 CRLF', JSON.stringify(P('﻿' + base + '\r\n').rows) === JSON.stringify(P(base).rows));
  check('빈 줄은 건너뛴다', P(base + '\n\n' + base).rows.length === 2);

  const short = P(base.split(',').slice(0, 11).join(','));
  check('11열은 실패', !!short.error, JSON.stringify(short));
  check('11열 오류에 줄 번호', /1/.test(short.error || ''), short.error);

  const nan = P('465,x,97,596,776,652,442,626,637,96,744,406');
  check('숫자가 아니면 실패', !!nan.error, JSON.stringify(nan));

  const hole = P('465,,97,596,776,652,442,626,637,96,744,406');
  check('중간 빈 칸은 0 이 아니라 실패', !!hole.error, JSON.stringify(hole));

  const empty = P('\n\n');
  check('유효 행이 없으면 실패', !!empty.error, JSON.stringify(empty));

  const twoLines = P(base + '\n' + base.split(',').slice(0, 11).join(','));
  check('둘째 줄 오류는 줄 번호 2', /2/.test(twoLines.error || ''), twoLines.error);
```

- [ ] **Step 2: 실패를 확인한다**

Run: `node test/fidelity.js`

Expected: FAIL — `TypeError: P is not a function`

- [ ] **Step 3: 구현한다**

`index.html` 의 `app` 블록, `root.APP = {...}` 위에 넣는다.

```js
// modCSVImport.bas 의 ParseCSV 와 동일한 판정. 실패 위치만 덧붙인다.
function parseCSV(text) {
  const lines = text.replace(/^﻿/, '').replace(/\r/g, '').split('\n');
  const rows = [];
  for (let i = 0; i < lines.length; i++) {
    let t = lines[i].trim();
    // 후행 빈 열 허용: 줄 끝 콤마와 그 뒤 공백·탭을 벗겨낸다
    while (t.length && (t.endsWith(',') || t.endsWith(' ') || t.endsWith('\t'))) {
      t = t.slice(0, -1);
    }
    if (t === '') continue;

    const parts = t.split(',');
    if (parts.length !== 12) {
      return { error: (i + 1) + '번째 줄이 ' + parts.length + '열입니다 (12열이어야 합니다)' };
    }
    // Number('') 은 0 이므로 빈 칸을 먼저 걸러야 한다 (VBA CDbl("") 은 오류)
    const nums = parts.map(p => (p.trim() === '' ? NaN : Number(p.trim())));
    const bad = nums.findIndex(v => !Number.isFinite(v));
    if (bad >= 0) {
      return { error: (i + 1) + '번째 줄 ' + (bad + 1) + '번째 값이 숫자가 아닙니다: "' + parts[bad].trim() + '"' };
    }
    rows.push(nums);
  }
  if (rows.length === 0) return { error: '유효한 데이터 행이 없습니다' };
  return { rows: rows };
}
```

`root.APP` 에 등록한다.

```js
root.APP = {
  OFFICE_THEME_XML: OFFICE_THEME_XML,
  parseCSV: parseCSV,
};
```

- [ ] **Step 4: 통과를 확인한다**

Run: `node test/fidelity.js`

Expected: Task 2 구획의 12개 항목이 모두 `ok`.

- [ ] **Step 5: 커밋**

```bash
git add index.html test/fidelity.js
git commit -m "Port ParseCSV to the SPA with row-level error locations"
```

---

### Task 3: `buildModel` — 1~2행과 본문 그리드

`BuildSheet` 이식의 첫 조각. 모델 구조와 상수를 여기서 확정한다.

**Files:**
- Modify: `index.html` — `app` 블록
- Modify: `test/fidelity.js`

**Interfaces:**
- Consumes: `APP.parseCSV`
- Produces:
  - `APP.buildModel(rows, titleText, labelText)` → `{cells, merges, colWidths, rowHeight, maxRow, maxCol}`
    - `cells`: `{r, c, v, s}[]` — `r`,`c` 는 1-based, `v` 는 `number|string|null`, `s` 는 스타일 토큰
    - `merges`: `{r1, c1, r2, c2}[]` — 마스터 셀만 `cells` 에 있고 덮인 셀은 넣지 않는다
    - `colWidths`: `{[col:number]: number}`
  - `APP.STYLES` — 토큰 → `{css, theme, tint}`
  - `APP.COL` — `{A:1, B:2, C:3, O:15, P:16, Q:17, R:18}`

- [ ] **Step 1: 실패하는 테스트를 쓴다**

`test/fidelity.js` 에 이어 붙인다. 먼저 공용 헬퍼를 `main()` 밖, `loadApp()` 아래에 둔다.

```js
// 모델에서 셀 하나를 집어낸다
function cellAt(model, r, c) {
  return model.cells.find(x => x.r === r && x.c === c) || null;
}
function hasMerge(model, r1, c1, r2, c2) {
  return model.merges.some(m => m.r1 === r1 && m.c1 === c1 && m.r2 === r2 && m.c2 === c2);
}
```

`main()` 안, Task 2 구획 뒤:

```js
  section('Task 3: buildModel 머리글과 본문');
  const rows12 = app.APP.parseCSV(fs.readFileSync(path.join(ROOT, 'random_data.csv'), 'utf8')).rows;
  check('샘플이 12행으로 파싱된다', rows12.length === 12, 'n=' + rows12.length);

  const M = app.APP.buildModel(rows12, '8/28 111G', 'XX');
  const at = (r, c) => cellAt(M, r, c);

  check('A1 = 제목, accent2', at(1, 1) && at(1, 1).v === '8/28 111G' && at(1, 1).s === 'accent2', JSON.stringify(at(1, 1)));
  check('A1:O1 병합', hasMerge(M, 1, 1, 1, 15));
  check('제목 병합에 덮인 셀은 모델에 없다', at(1, 5) === null);
  check('Q1 = 라벨, gray', at(1, 17) && at(1, 17).v === 'XX' && at(1, 17).s === 'gray', JSON.stringify(at(1, 17)));
  check('Q1:R1 병합', hasMerge(M, 1, 17, 1, 18));

  check('C2 = A, accent2', at(2, 3) && at(2, 3).v === 'A' && at(2, 3).s === 'accent2');
  check('D2 = B', at(2, 4) && at(2, 4).v === 'B');
  check('N2 = B (12번째)', at(2, 14) && at(2, 14).v === 'B');
  check('Q2 = A, accent5', at(2, 17) && at(2, 17).v === 'A' && at(2, 17).s === 'accent5');
  check('R2 = B, accent4', at(2, 18) && at(2, 18).v === 'B' && at(2, 18).s === 'accent4');

  check('C3 = 1 (머리글 행)', at(3, 3) && at(3, 3).v === 1 && at(3, 3).s === 'accent2');
  check('N3 = 12', at(3, 14) && at(3, 14).v === 12);
  check('C4 = 첫 값 465, data', at(4, 3) && at(4, 3).v === 465 && at(4, 3).s === 'data');
  check('N4 = 마지막 값 406', at(4, 14) && at(4, 14).v === 406);
  check('마지막 데이터 행 N26 = 99', at(26, 14) && at(26, 14).v === 99, JSON.stringify(at(26, 14)));
  check('O3 = 빈 칸막이, accent2', at(3, 15) && at(3, 15).v === null && at(3, 15).s === 'accent2');
  check('O4 = 빈 칸막이, data', at(4, 15) && at(4, 15).v === null && at(4, 15).s === 'data');
```

- [ ] **Step 2: 실패를 확인한다**

Run: `node test/fidelity.js`

Expected: FAIL — `app.APP.buildModel is not a function`

- [ ] **Step 3: 구현한다**

`app` 블록의 `parseCSV` 아래에 넣는다.

```js
const COL = { A: 1, B: 2, C: 3, O: 15, P: 16, Q: 17, R: 18 };
const NCOL = 12;                 // 데이터 열 수 (C..N)
const FONT_NAME = '맑은 고딕';
const FONT_SIZE = 11;

const STYLES = {
  accent2: { css: '#F8CBAD', theme: 5,    tint:  0.5999938962981048 },
  accent5: { css: '#BDD7EE', theme: 8,    tint:  0.5999938962981048 },
  accent4: { css: '#FFF2CC', theme: 7,    tint:  0.7999816888943144 },
  gray:    { css: '#D9D9D9', theme: 0,    tint: -0.1499984740745262 },
  data:    { css: '#FFFFFF', theme: null, tint:  0 },
};

// BuildSheet 를 이식한 유일한 지점. 미리보기와 xlsx 가 이 결과를 각자 그린다.
function buildModel(rows, titleText, labelText) {
  const n = rows.length;
  const cells = [];
  const merges = [];
  const put = (r, c, v, s) => cells.push({ r: r, c: c, v: v, s: s });

  // 1행: 제목 A1:O1, 라벨 Q1:R1. 병합에 덮인 셀은 넣지 않는다 —
  // ExcelJS 가 마스터 스타일을 슬레이브에 그대로 전파하고, 미리보기는 colspan 으로 처리한다.
  put(1, COL.A, titleText, 'accent2');
  merges.push({ r1: 1, c1: COL.A, r2: 1, c2: COL.O });
  put(1, COL.Q, labelText, 'gray');
  merges.push({ r1: 1, c1: COL.Q, r2: 1, c2: COL.R });

  // 2행: C..N 에 A/B 반복
  for (let i = 0; i < NCOL; i++) put(2, COL.C + i, i % 2 === 0 ? 'A' : 'B', 'accent2');
  put(2, COL.Q, 'A', 'accent5');
  put(2, COL.R, 'B', 'accent4');

  // 본문: CSV 행마다 [머리글 행][데이터 행] 두 줄
  for (let i = 0; i < n; i++) {
    const headerRow = 3 + i * 2;
    const dataRow = headerRow + 1;
    for (let c = 0; c < NCOL; c++) {
      put(headerRow, COL.C + c, c + 1, 'accent2');
      put(dataRow,   COL.C + c, rows[i][c], 'data');
    }
    put(headerRow, COL.O, null, 'accent2');   // 빈 칸막이
    put(dataRow,   COL.O, null, 'data');
  }

  return {
    cells: cells,
    merges: merges,
    colWidths: {},
    rowHeight: 16.5,
    maxRow: 2 + 2 * n,
    maxCol: COL.R,
  };
}
```

`root.APP` 에 `buildModel`, `STYLES`, `COL` 을 추가한다.

- [ ] **Step 4: 통과를 확인한다**

Run: `node test/fidelity.js`

Expected: Task 3 구획 18개 항목이 모두 `ok`.

- [ ] **Step 5: 커밋**

```bash
git add index.html test/fidelity.js
git commit -m "Build the header rows and main grid of the layout model"
```

---

### Task 4: `buildModel` — A/B 그룹 병합

CSV 2행마다 B열, 4행마다 A열에 그룹 번호를 병합해 넣는다.

**Files:**
- Modify: `index.html` — `app` 블록
- Modify: `test/fidelity.js`

**Interfaces:**
- Consumes: Task 3 의 `buildModel`, `COL`
- Produces: 변경 없음 — 같은 `buildModel` 이 `merges` 와 `cells` 를 더 채운다

- [ ] **Step 1: 실패하는 테스트를 쓴다**

n=12 이면 B 그룹은 6개, A 그룹은 3개다. 그룹 g 의 행 범위는 `rFrom = 3 + (first-1)*2`, `rTo = 3 + (last-1)*2 + 1`.

```js
  section('Task 4: A/B 그룹 병합');
  check('B 그룹 1: B3:B6 = 1', hasMerge(M, 3, 2, 6, 2) && at(3, 2) && at(3, 2).v === 1 && at(3, 2).s === 'accent2');
  check('B 그룹 2: B7:B10 = 2', hasMerge(M, 7, 2, 10, 2) && at(7, 2) && at(7, 2).v === 2);
  check('B 그룹 6: B23:B26 = 6', hasMerge(M, 23, 2, 26, 2) && at(23, 2) && at(23, 2).v === 6);
  check('B 그룹 수 = 6', M.merges.filter(m => m.c1 === 2).length === 6);

  check('A 그룹 1: A3:A10 = 1', hasMerge(M, 3, 1, 10, 1) && at(3, 1) && at(3, 1).v === 1 && at(3, 1).s === 'accent2');
  check('A 그룹 2: A11:A18 = 2', hasMerge(M, 11, 1, 18, 1) && at(11, 1) && at(11, 1).v === 2);
  check('A 그룹 3: A19:A26 = 3', hasMerge(M, 19, 1, 26, 1) && at(19, 1) && at(19, 1).v === 3);
  check('A 그룹 수 = 3', M.merges.filter(m => m.c1 === 1 && m.r1 !== 1).length === 3);

  check('병합 총 개수 = 11', M.merges.length === 11, 'merges=' + M.merges.length);

  // 나누어떨어지지 않는 n 에서 마지막 그룹이 잘려야 한다
  const M5 = app.APP.buildModel(rows12.slice(0, 5), 'T', 'L');
  check('n=5 의 마지막 B 그룹은 B11:B12', hasMerge(M5, 11, 2, 12, 2), JSON.stringify(M5.merges.filter(m => m.c1 === 2)));
  check('n=5 의 마지막 A 그룹은 A11:A12', hasMerge(M5, 11, 1, 12, 1), JSON.stringify(M5.merges.filter(m => m.c1 === 1 && m.r1 !== 1)));
```

- [ ] **Step 2: 실패를 확인한다**

Run: `node test/fidelity.js`

Expected: Task 4 구획 전부 FAIL, `병합 총 개수 = 11` 은 `merges=2`.

- [ ] **Step 3: 구현한다**

`buildModel` 의 본문 루프 뒤, `return` 앞에 넣는다.

```js
  // A/B 그룹 라벨 — B는 CSV 2행씩, A는 4행씩 묶어 병합
  const GROUPS = [{ col: COL.B, size: 2 }, { col: COL.A, size: 4 }];
  for (let gi = 0; gi < GROUPS.length; gi++) {
    const col = GROUPS[gi].col, size = GROUPS[gi].size;
    for (let g = 0; g * size < n; g++) {
      const first = g * size + 1;
      const last = Math.min((g + 1) * size, n);
      const rFrom = 3 + (first - 1) * 2;
      const rTo   = 3 + (last  - 1) * 2 + 1;
      put(rFrom, col, g + 1, 'accent2');
      merges.push({ r1: rFrom, c1: col, r2: rTo, c2: col });
    }
  }
```

- [ ] **Step 4: 통과를 확인한다**

Run: `node test/fidelity.js`

Expected: Task 4 구획 11개 항목이 모두 `ok`.

- [ ] **Step 5: 커밋**

```bash
git add index.html test/fidelity.js
git commit -m "Add the merged A/B group labels to the layout model"
```

---

### Task 5: `buildModel` — P/Q/R 사이드 표와 치수

홀수 열 값을 Q에, 짝수 열 값을 R에 모아 각각 내림차순으로 놓는다. VBA는 T/U열을 임시 작업 공간으로 썼지만 여기서는 JS 정렬로 대체한다.

**Files:**
- Modify: `index.html` — `app` 블록
- Modify: `test/fidelity.js`

**Interfaces:**
- Consumes: Task 4 의 `buildModel`
- Produces: `buildModel` 이 `colWidths = {16: 5.125}` 와 `maxRow = 2 + 6n` 을 반환

- [ ] **Step 1: 실패하는 테스트를 쓴다**

```js
  section('Task 5: 사이드 표와 치수');
  check('P1 = 값 없는 gray', at(1, 16) && at(1, 16).v === null && at(1, 16).s === 'gray', JSON.stringify(at(1, 16)));
  check('P2 = 값 없는 gray', at(2, 16) && at(2, 16).v === null && at(2, 16).s === 'gray');
  check('P3 = 1, gray', at(3, 16) && at(3, 16).v === 1 && at(3, 16).s === 'gray');
  check('마지막 P행 = 72', at(74, 16) && at(74, 16).v === 72, JSON.stringify(at(74, 16)));

  // 홀수 열(1,3,5,7,9,11) 최대는 989, 짝수 열(2,4,...) 최대는 996
  check('Q3 = 홀수 열 최대 989', at(3, 17) && at(3, 17).v === 989 && at(3, 17).s === 'data', JSON.stringify(at(3, 17)));
  check('R3 = 짝수 열 최대 996', at(3, 18) && at(3, 18).v === 996 && at(3, 18).s === 'data');

  const qCol = [];
  for (let r = 3; r <= 74; r++) qCol.push(at(r, 17).v);
  check('Q는 72개', qCol.length === 72);
  check('Q는 내림차순', qCol.every((v, i) => i === 0 || qCol[i - 1] >= v));
  const qExpected = [];
  for (let k = 0; k < 6; k++) for (let i = 0; i < 12; i++) qExpected.push(rows12[i][2 * k]);
  qExpected.sort((a, b) => b - a);
  check('Q 내용이 홀수 열 전체와 일치', JSON.stringify(qCol) === JSON.stringify(qExpected));

  const rCol = [];
  for (let r = 3; r <= 74; r++) rCol.push(at(r, 18).v);
  const rExpected = [];
  for (let k = 0; k < 6; k++) for (let i = 0; i < 12; i++) rExpected.push(rows12[i][2 * k + 1]);
  rExpected.sort((a, b) => b - a);
  check('R 내용이 짝수 열 전체와 일치', JSON.stringify(rCol) === JSON.stringify(rExpected));

  check('P열 너비 5.125', M.colWidths[16] === 5.125, JSON.stringify(M.colWidths));
  check('행 높이 16.5', M.rowHeight === 16.5);
  check('maxRow = 2 + 6n = 74', M.maxRow === 74, 'maxRow=' + M.maxRow);
  check('maxCol = 18', M.maxCol === 18);
```

- [ ] **Step 2: 실패를 확인한다**

Run: `node test/fidelity.js`

Expected: Task 5 구획 전부 FAIL. `maxRow=26`, `colWidths={}`.

- [ ] **Step 3: 구현한다**

그룹 병합 루프 뒤, `return` 앞에 넣는다.

```js
  // 사이드 표 P/Q/R — Q는 홀수 열 값 전체, R은 짝수 열 값 전체를 내림차순으로
  put(1, COL.P, null, 'gray');
  put(2, COL.P, null, 'gray');
  const qs = [], rs = [];
  for (let k = 0; k < 6; k++) {
    for (let i = 0; i < n; i++) {
      qs.push(rows[i][2 * k]);       // 1,3,5,7,9,11 번째 열 (0-based 0,2,4,...)
      rs.push(rows[i][2 * k + 1]);   // 2,4,6,8,10,12 번째 열
    }
  }
  qs.sort((a, b) => b - a);
  rs.sort((a, b) => b - a);
  for (let i = 0; i < qs.length; i++) {
    const r = 3 + i;
    put(r, COL.P, i + 1, 'gray');
    put(r, COL.Q, qs[i], 'data');
    put(r, COL.R, rs[i], 'data');
  }
```

`return` 문을 고친다.

```js
  return {
    cells: cells,
    merges: merges,
    colWidths: { 16: 5.125 },   // VBA ColumnWidth 4.5 의 엑셀 저장값
    rowHeight: 16.5,
    maxRow: 2 + 6 * n,          // 사이드 표가 본문(2+2n)보다 항상 길다
    maxCol: COL.R,
  };
```

- [ ] **Step 4: 통과를 확인한다**

Run: `node test/fidelity.js`

Expected: Task 5 구획 15개 항목이 모두 `ok`.

- [ ] **Step 5: 커밋**

```bash
git add index.html test/fidelity.js
git commit -m "Add the sorted P/Q/R side table and sheet dimensions"
```

---

### Task 6: `toXlsx`

모델을 ExcelJS 워크북으로 옮긴다. 테마 교체가 여기 들어간다.

**Files:**
- Modify: `index.html` — `app` 블록
- Modify: `test/fidelity.js`

**Interfaces:**
- Consumes: `APP.buildModel`, `APP.STYLES`, `APP.OFFICE_THEME_XML`, 전역 `ExcelJS`
- Produces: `APP.toXlsx(model)` → `Promise<Buffer|ArrayBuffer>` — xlsx 바이트

- [ ] **Step 1: 실패하는 테스트를 쓴다**

```js
  section('Task 6: toXlsx');
  const buf = await app.APP.toXlsx(M);
  check('바이트가 나온다', buf && buf.byteLength > 5000, 'bytes=' + (buf && buf.byteLength));

  const back = new app.ExcelJS.Workbook();
  await back.xlsx.load(buf);
  const bws = back.getWorksheet('Sheet1');
  check('시트 이름 Sheet1', !!bws);

  const fillOf = c => (c.fill && c.fill.fgColor)
    ? 'theme=' + c.fill.fgColor.theme + ' tint=' + Math.round(c.fill.fgColor.tint * 1e6) / 1e6
    : 'none';

  check('A1 값', bws.getCell('A1').value === '8/28 111G');
  check('A1 채우기 accent2', fillOf(bws.getCell('A1')) === 'theme=5 tint=0.599994', fillOf(bws.getCell('A1')));
  check('E1 (병합 슬레이브) 도 accent2', fillOf(bws.getCell('E1')) === 'theme=5 tint=0.599994', fillOf(bws.getCell('E1')));
  check('Q1 채우기 gray', fillOf(bws.getCell('Q1')) === 'theme=0 tint=-0.149998', fillOf(bws.getCell('Q1')));
  check('Q2 채우기 accent5', fillOf(bws.getCell('Q2')) === 'theme=8 tint=0.599994', fillOf(bws.getCell('Q2')));
  check('R2 채우기 accent4', fillOf(bws.getCell('R2')) === 'theme=7 tint=0.799982', fillOf(bws.getCell('R2')));
  check('C4 는 채우기 없음', fillOf(bws.getCell('C4')) === 'none', fillOf(bws.getCell('C4')));
  check('O3 은 값 없이 accent2', bws.getCell('O3').value === null && fillOf(bws.getCell('O3')) === 'theme=5 tint=0.599994');

  check('C4 폰트', bws.getCell('C4').font.name === '맑은 고딕' && bws.getCell('C4').font.size === 11, JSON.stringify(bws.getCell('C4').font));
  check('C4 정렬', bws.getCell('C4').alignment.horizontal === 'center' && bws.getCell('C4').alignment.vertical === 'middle');
  check('C4 사방 테두리', ['top','left','bottom','right'].every(k => bws.getCell('C4').border[k].style === 'thin'));

  check('병합 11개', bws.model.merges.length === 11, JSON.stringify(bws.model.merges));
  check('P열 너비 5.125', bws.getColumn(16).width === 5.125, 'w=' + bws.getColumn(16).width);
  check('1행 높이 16.5', bws.getRow(1).height === 16.5);
  check('74행 높이 16.5', bws.getRow(74).height === 16.5);

  // 테마 교체가 먹었는지 — 안 먹으면 accent2 가 C0504D (Office 2007) 로 남는다
  const themeXml = back._themes && back._themes.theme1;
  check('테마가 Office 2013+ (accent2=ED7D31)', /ED7D31/.test(themeXml || ''), 'theme len=' + (themeXml || '').length);
```

- [ ] **Step 2: 실패를 확인한다**

Run: `node test/fidelity.js`

Expected: FAIL — `app.APP.toXlsx is not a function`

- [ ] **Step 3: 구현한다**

`app` 블록의 `buildModel` 아래에 넣는다.

```js
const THIN = { style: 'thin' };
const BORDER_ALL = { top: THIN, left: THIN, bottom: THIN, right: THIN };

function toXlsx(model) {
  const wb = new ExcelJS.Workbook();
  // ponytail: _themes 는 ExcelJS 공개 API가 아니다. 번들을 인라인해 버전을 고정했으므로
  // 지금은 안전하다. 라이브러리를 갈아끼워 색이 틀어지면 STYLES 의 theme/tint 를
  // 해결된 ARGB (accent2=F8CBAD 등) 로 바꾸는 것이 대체 경로다.
  wb._themes = { theme1: OFFICE_THEME_XML };

  const ws = wb.addWorksheet('Sheet1');

  // 스타일을 먼저, 병합을 나중에. ExcelJS 가 마스터 스타일을 슬레이브에 전파한다.
  for (let i = 0; i < model.cells.length; i++) {
    const item = model.cells[i];
    const cell = ws.getCell(item.r, item.c);
    if (item.v !== null) cell.value = item.v;
    cell.font = { name: FONT_NAME, size: FONT_SIZE };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    cell.border = BORDER_ALL;
    const st = STYLES[item.s];
    if (st.theme !== null) {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { theme: st.theme, tint: st.tint } };
    }
  }

  for (let i = 0; i < model.merges.length; i++) {
    const m = model.merges[i];
    ws.mergeCells(m.r1, m.c1, m.r2, m.c2);
  }

  const widths = Object.keys(model.colWidths);
  for (let i = 0; i < widths.length; i++) {
    ws.getColumn(Number(widths[i])).width = model.colWidths[widths[i]];
  }
  for (let r = 1; r <= model.maxRow; r++) ws.getRow(r).height = model.rowHeight;

  return wb.xlsx.writeBuffer();
}
```

`root.APP` 에 `toXlsx` 를 추가한다.

- [ ] **Step 4: 통과를 확인한다**

Run: `node test/fidelity.js`

Expected: Task 6 구획 18개 항목이 모두 `ok`.

`테마가 Office 2013+` 가 FAIL 이면 `OFFICE_THEME_XML` 이 제대로 삽입되지 않은 것이다. Task 1 Step 3 의 추출을 다시 한다.

- [ ] **Step 5: 커밋**

```bash
git add index.html test/fidelity.js
git commit -m "Render the layout model to xlsx with the workbook theme replaced"
```

---

### Task 7: VBA 산출물과의 정합성 대조

여기가 마이그레이션이 맞는지 판정하는 지점이다. 앞의 테스트들은 "내가 의도한 대로 만들었나" 를 봤고, 이 테스트는 "VBA와 같은가" 를 본다.

**Files:**
- Modify: `test/fidelity.js`
- Modify: `index.html` — 대조에서 차이가 나오면 그것만 수정

**Interfaces:**
- Consumes: `APP.toXlsx`, `test/reference.xlsx`
- Produces: 없음 — 검증 전용

- [ ] **Step 1: 대조 테스트를 쓴다**

```js
  section('Task 7: reference.xlsx 와 셀 단위 대조');
  const refWb = new app.ExcelJS.Workbook();
  await refWb.xlsx.load(fs.readFileSync(path.join(ROOT, 'test', 'reference.xlsx')));
  const rws = refWb.worksheets[0];

  // tint 는 VBA 와 ExcelJS 가 쓰는 소수 자릿수가 달라 반올림해 비교한다
  const facet = {
    value: c => (c.value === null || c.value === undefined) ? null
      : (c.value && c.value.richText ? c.value.richText.map(t => t.text).join('') : c.value),
    fill: c => (c.fill && c.fill.fgColor)
      ? 'theme=' + c.fill.fgColor.theme + ' tint=' + Math.round((c.fill.fgColor.tint || 0) * 1e6) / 1e6
      : 'none',
    border: c => ['top','left','bottom','right'].map(k => (c.border && c.border[k]) ? c.border[k].style : '-').join(','),
    font: c => c.font ? (c.font.name + '/' + c.font.size) : '-',
    align: c => c.alignment ? (c.alignment.horizontal + '/' + c.alignment.vertical) : '-',
  };

  const names = Object.keys(facet);
  const diffs = [];
  for (let r = 1; r <= 74; r++) {
    for (let c = 1; c <= 18; c++) {
      for (let f = 0; f < names.length; f++) {
        const k = names[f];
        const a = facet[k](bws.getCell(r, c));
        const b = facet[k](rws.getCell(r, c));
        if (String(a) !== String(b)) {
          diffs.push('R' + r + 'C' + c + ' ' + k + ': SPA=' + a + ' REF=' + b);
        }
      }
    }
  }
  check('셀 서식·값 차이 0건', diffs.length === 0, diffs.slice(0, 12).join(' | ') + (diffs.length > 12 ? ' … 총 ' + diffs.length + '건' : ''));

  const sortRefs = a => a.slice().sort();
  check('병합 목록 일치',
    JSON.stringify(sortRefs(bws.model.merges)) === JSON.stringify(sortRefs(rws.model.merges)),
    'SPA=' + JSON.stringify(sortRefs(bws.model.merges)) + ' REF=' + JSON.stringify(sortRefs(rws.model.merges)));

  const wdiff = [];
  for (let c = 1; c <= 18; c++) {
    const a = bws.getColumn(c).width, b = rws.getColumn(c).width;
    if (String(a) !== String(b)) wdiff.push('col' + c + ': SPA=' + a + ' REF=' + b);
  }
  check('열 너비 일치', wdiff.length === 0, wdiff.join(' | '));

  const hdiff = [];
  for (let r = 1; r <= 74; r++) {
    const a = bws.getRow(r).height, b = rws.getRow(r).height;
    if (String(a) !== String(b)) hdiff.push('row' + r + ': SPA=' + a + ' REF=' + b);
  }
  check('행 높이 일치', hdiff.length === 0, hdiff.slice(0, 8).join(' | '));
```

- [ ] **Step 2: 대조를 돌린다**

Run: `node test/fidelity.js`

차이가 나오면 출력이 정확히 어느 셀의 어느 항목인지 알려준다. 스펙 8장에 적어둔 예상 차이와 대응:

| 나올 수 있는 차이 | 대응 |
|---|---|
| 열 너비 불일치 | 참조 파일의 16번 열 너비는 `5.125` 로 확인되어 있다. 다르게 나오면 `buildModel` 의 `colWidths[16]` 을 REF 값에 맞춘다 |
| 특정 셀 `fill: SPA=none REF=theme=...` | 그 셀을 `buildModel` 이 빠뜨린 것이다. 스펙 5장 레이아웃 사양과 대조 |
| `font` 불일치 | 문자열 정규화 문제. 공백이나 자모 분리 확인 |
| 값이 `null` vs `''` | `toXlsx` 에서 `v !== null` 판정 확인 |

**비교 범위를 넓히지 말 것.** 참조 파일은 Excel이 쓴 것이라 우리가 지정하지 않는 속성을 더 갖고 있다. 확인된 것만 적으면:

- `font` — REF는 `{size, name, color:{theme:1}, family:3, charset:129, scheme:'minor'}`, SPA는 `{name, size}`. 그래서 `facet.font` 는 이름과 크기만 본다. 객체 전체를 비교하도록 "개선" 하면 의미 없는 실패가 난다
- `fill` — REF는 `bgColor:{indexed:64}` 를 갖고 SPA는 없다. 솔리드 채우기에서 `bgColor` 는 렌더링에 영향이 없으므로 `facet.fill` 은 `fgColor` 만 본다
- `tint` — REF는 `0.59999389629810485`, SPA는 `0.5999938962981048` 로 자릿수가 다르다. 그래서 반올림해 비교한다

차이를 고칠 때는 **`index.html` 만** 고친다. `reference.xlsx` 는 절대 다시 만들지 않는다.

- [ ] **Step 3: 차이 0건을 확인한다**

Run: `node test/fidelity.js`

Expected:
```
== Task 7: reference.xlsx 와 셀 단위 대조 ==
  ok   셀 서식·값 차이 0건
  ok   병합 목록 일치
  ok   열 너비 일치
  ok   행 높이 일치

전체 통과
```

- [ ] **Step 4: 커밋**

```bash
git add test/fidelity.js index.html
git commit -m "Verify SPA output matches the VBA reference workbook cell for cell"
```

---

### Task 8: UI — 파일 선택, 목록, 미리보기

여기부터는 `ui` 블록과 CSS다. 샌드박스 테스트 범위 밖이므로 브라우저에서 확인한다. `node test/fidelity.js` 는 계속 통과해야 한다 (`ui` 블록이 걸러지는지 확인하는 의미도 있다).

**Files:**
- Modify: `index.html` — `<style>`, `<body>` 마크업, `ui` 블록

**Interfaces:**
- Consumes: `APP.parseCSV`, `APP.buildModel`, `APP.STYLES`
- Produces: `ui` 블록 내부 상태 `entries` — `{name, base, rows, error}[]`

- [ ] **Step 1: CSS를 넣는다**

`<style>` 안에 넣는다. 색은 Global Constraints 의 표와 같은 값이다.

```css
body { font: 14px "맑은 고딕", "Malgun Gothic", sans-serif; margin: 0; padding: 16px; }
header { display: flex; gap: 16px; align-items: center; flex-wrap: wrap; margin-bottom: 12px; }
header input[type=text] { font: inherit; padding: 4px 6px; }
main { display: flex; gap: 16px; align-items: flex-start; }
#list { list-style: none; margin: 0; padding: 0; min-width: 220px; }
#list li { padding: 6px 8px; cursor: pointer; border: 1px solid transparent; }
#list li.sel { background: #eef3fb; border-color: #9db8dd; }
#list li.bad { color: #b00; cursor: default; }
#list .note { display: block; font-size: 12px; color: #666; }
#preview { overflow: auto; max-width: calc(100vw - 300px); max-height: 78vh; }
footer { margin-top: 12px; display: flex; gap: 8px; }

table.sheet { border-collapse: collapse; table-layout: fixed; font: 11pt "맑은 고딕", "Malgun Gothic", sans-serif; }
table.sheet td { border: 1px solid #000; text-align: center; vertical-align: middle; overflow: hidden; white-space: nowrap; padding: 0; }
table.sheet td.empty { border: none; }
table.sheet td.s-accent2 { background: #F8CBAD; }
table.sheet td.s-accent5 { background: #BDD7EE; }
table.sheet td.s-accent4 { background: #FFF2CC; }
table.sheet td.s-gray    { background: #D9D9D9; }
table.sheet td.s-data    { background: #FFFFFF; }
```

- [ ] **Step 2: 마크업을 넣는다**

`<body>` 안, 스크립트 앞에 넣는다.

```html
<header>
  <input type="file" id="files" multiple accept=".csv">
  <label>제목 <input type="text" id="title" value="8/28 111G"></label>
  <label>라벨 <input type="text" id="label" value="XX"></label>
</header>
<main>
  <ul id="list"></ul>
  <div id="preview"></div>
</main>
<footer>
  <button id="dl-one" disabled>선택 내려받기</button>
  <button id="dl-all" disabled>전체 내려받기</button>
</footer>
```

- [ ] **Step 3: 미리보기 렌더러를 `ui` 블록에 쓴다**

```js
'use strict';
(function () {
  const filesEl = document.getElementById('files');
  const titleEl = document.getElementById('title');
  const labelEl = document.getElementById('label');
  const listEl = document.getElementById('list');
  const previewEl = document.getElementById('preview');
  const dlOneEl = document.getElementById('dl-one');
  const dlAllEl = document.getElementById('dl-all');

  let entries = [];      // {name, base, rows, error}
  let selected = -1;

  // 엑셀 저장 너비 → 화면 px. 폰트에 따라 달라지므로 여기 두 상수만 조정하면 된다.
  const PX_PER_WIDTH = 7;
  const PX_PADDING = 5;
  const DEFAULT_WIDTH = 8.43;
  const colPx = w => Math.round(w * PX_PER_WIDTH) + PX_PADDING;

  const escapeHtml = s => String(s).replace(/[&<>"]/g,
    ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[ch]));

  function renderPreview(model) {
    const covered = new Set();
    const span = new Map();
    for (const m of model.merges) {
      span.set(m.r1 + ',' + m.c1, { rowspan: m.r2 - m.r1 + 1, colspan: m.c2 - m.c1 + 1 });
      for (let r = m.r1; r <= m.r2; r++) {
        for (let c = m.c1; c <= m.c2; c++) {
          if (!(r === m.r1 && c === m.c1)) covered.add(r + ',' + c);
        }
      }
    }
    const byPos = new Map(model.cells.map(x => [x.r + ',' + x.c, x]));

    let html = '<table class="sheet"><colgroup>';
    for (let c = 1; c <= model.maxCol; c++) {
      const w = model.colWidths[c] === undefined ? DEFAULT_WIDTH : model.colWidths[c];
      html += '<col style="width:' + colPx(w) + 'px">';
    }
    html += '</colgroup><tbody>';

    for (let r = 1; r <= model.maxRow; r++) {
      html += '<tr style="height:' + (model.rowHeight / 0.75) + 'px">';   // pt → px
      for (let c = 1; c <= model.maxCol; c++) {
        const key = r + ',' + c;
        if (covered.has(key)) continue;
        const sp = span.get(key);
        const attrs = sp ? ' rowspan="' + sp.rowspan + '" colspan="' + sp.colspan + '"' : '';
        const cell = byPos.get(key);
        if (!cell) { html += '<td class="empty"' + attrs + '></td>'; continue; }
        const text = cell.v === null ? '' : escapeHtml(cell.v);
        html += '<td class="s-' + cell.s + '"' + attrs + '>' + text + '</td>';
      }
      html += '</tr>';
    }
    return html + '</tbody></table>';
  }

  function draw() {
    listEl.innerHTML = entries.map((e, i) => {
      const cls = e.error ? 'bad' : (i === selected ? 'sel' : '');
      const note = e.error ? e.error : e.rows.length + '행';
      return '<li class="' + cls + '" data-i="' + i + '">' +
        (e.error ? '✗ ' : '✓ ') + escapeHtml(e.name) +
        '<span class="note">' + escapeHtml(note) + '</span></li>';
    }).join('');

    const ok = entries.filter(e => !e.error);
    dlOneEl.disabled = selected < 0;
    dlAllEl.disabled = ok.length === 0;

    if (selected < 0) { previewEl.innerHTML = ''; return; }
    const model = APP.buildModel(entries[selected].rows, titleEl.value, labelEl.value);
    previewEl.innerHTML = renderPreview(model);
  }

  filesEl.addEventListener('change', async () => {
    entries = [];
    for (const f of filesEl.files) {
      const res = APP.parseCSV(await f.text());
      entries.push({
        name: f.name,
        base: f.name.replace(/\.[^.]*$/, ''),
        rows: res.rows || null,
        error: res.error || null,
      });
    }
    selected = entries.findIndex(e => !e.error);
    draw();
  });

  listEl.addEventListener('click', ev => {
    const li = ev.target.closest('li');
    if (!li) return;
    const i = Number(li.dataset.i);
    if (entries[i].error) return;
    selected = i;
    draw();
  });

  titleEl.addEventListener('input', draw);
  labelEl.addEventListener('input', draw);
})();
```

- [ ] **Step 4: 샌드박스 테스트가 여전히 통과하는지 확인한다**

`ui` 블록은 `document` 를 쓰므로, 추출 정규식이 이 블록을 제대로 걸러야 한다.

Run: `node test/fidelity.js`

Expected: 전체 통과. `document is not defined` 가 나오면 `<script id="ui">` 의 속성 표기가 정규식과 다른 것이다 (작은따옴표나 공백 확인).

- [ ] **Step 5: 브라우저에서 확인한다**

`index.html` 을 더블클릭해 연다. `random_data.csv` 를 선택하고 다음을 확인한다.

1. 목록에 `✓ random_data.csv` / `12행`
2. 미리보기 1행에 주황 배경(`#F8CBAD`)의 `8/28 111G` 가 O열까지 이어지고, 그 오른쪽에 회색 `XX`
3. Q2가 연한 파랑, R2가 연한 노랑
4. 3행이 `1..12`, 4행이 `465 996 97 …`
5. A열이 8행씩, B열이 4행씩 병합되어 번호가 보인다
6. 오른쪽 P/Q/R 표가 74행까지 이어지고 Q3=989, R3=996
7. 제목 입력을 고치면 미리보기 1행이 즉시 바뀐다

Excel에서 `test/reference.xlsx` 를 열어 나란히 놓고 모양이 같은지 본다. 열 폭이 눈에 띄게 다르면 `PX_PER_WIDTH` / `PX_PADDING` 을 조정한다.

- [ ] **Step 6: 커밋**

```bash
git add index.html
git commit -m "Add the file list and the sheet-faithful preview"
```

---

### Task 9: UI — 내려받기와 최종 확인

**Files:**
- Modify: `index.html` — `ui` 블록

**Interfaces:**
- Consumes: `APP.toXlsx`, Task 8 의 `entries` / `selected`
- Produces: 없음 — 최종 태스크

- [ ] **Step 1: 내려받기를 `ui` 블록에 쓴다**

Task 8 의 `titleEl.addEventListener('input', draw);` 앞에 넣는다.

```js
  async function download(entry) {
    const model = APP.buildModel(entry.rows, titleEl.value, labelEl.value);
    const buf = await APP.toXlsx(model);
    const blob = new Blob([buf], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = entry.base + '_result.xlsx';   // VBA 와 같은 규칙
    a.click();
    URL.revokeObjectURL(url);
  }

  dlOneEl.addEventListener('click', async () => {
    if (selected < 0) return;
    dlOneEl.disabled = true;
    try { await download(entries[selected]); } finally { draw(); }
  });

  dlAllEl.addEventListener('click', async () => {
    dlAllEl.disabled = true;
    try {
      // 순차 저장. 브라우저가 다중 내려받기 허용을 한 번 물을 수 있다.
      for (const e of entries) if (!e.error) await download(e);
    } finally { draw(); }
  });
```

- [ ] **Step 2: 샌드박스 테스트를 다시 돌린다**

Run: `node test/fidelity.js`

Expected: 전체 통과.

- [ ] **Step 3: 브라우저에서 단일 내려받기를 확인한다**

`index.html` 을 열고 `random_data.csv` 선택 → `선택 내려받기`.

1. `random_data_result.xlsx` 가 저장된다
2. Excel에서 열어 `test/reference.xlsx` 와 나란히 비교 — 색, 병합, 테두리, P열 폭이 같다

- [ ] **Step 4: 배치와 오류 경로를 확인한다**

검증용 CSV 두 개를 만든다.

```bash
cp random_data.csv /tmp/second.csv
head -3 random_data.csv | sed 's/,[0-9]*,$/,/' > /tmp/broken.csv
```

세 파일을 한 번에 선택하고 확인한다.

1. `random_data.csv` ✓, `second.csv` ✓, `broken.csv` ✗ 에 몇 번째 줄이 몇 열인지 표시
2. `broken.csv` 를 눌러도 선택되지 않는다
3. `전체 내려받기` → 정상 파일 2개만 저장되고 `broken` 은 빠진다
4. 제목을 `<b>x</b>` 로 바꿔도 미리보기가 깨지지 않고 그대로 글자로 보인다

- [ ] **Step 5: 완료 기준을 점검한다**

스펙 9장 대조.

```bash
node test/fidelity.js && git status --short
```

1. `node test/fidelity.js` 전체 통과
2. 파서 엣지케이스 5종 통과 (Task 2 구획)
3. `file://` 에서 다중 선택 → 미리보기 → 개별/전체 내려받기 동작 (Step 3, 4)
4. 미리보기와 산출물이 육안으로 동일 (Step 3)
5. `git status` 에 `.xlsm` / `.bas` 변경 없음

- [ ] **Step 6: 커밋**

```bash
git add index.html
git commit -m "Add single and batch xlsx download to the SPA"
```

---

## 완료 후

기존 VBA 도구는 그대로 남는다. 폐기 여부는 SPA를 실제로 며칠 써 본 뒤 따로 판단한다 (스펙 1.5).

`test/reference.xlsx` 는 이후 어떤 이유로도 다시 만들지 않는다. VBA 쪽 레이아웃을 바꿔야 하는 날이 오면, 그때는 픽스처 재생성이 아니라 "무엇이 왜 달라지는가" 를 먼저 적는다.
