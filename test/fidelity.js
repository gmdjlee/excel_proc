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
    TextEncoder, TextDecoder, URL, Uint8Array, ArrayBuffer,
  };
  sandbox.window = sandbox;
  sandbox.self = sandbox;
  sandbox.global = sandbox;
  vm.createContext(sandbox);
  const re = /<script(?![^>]*\bid="ui")[^>]*>([\s\S]*?)<\/script>/g;
  for (const m of html.matchAll(re)) vm.runInContext(m[1], sandbox);
  return sandbox;
}

// 모델에서 셀 하나를 집어낸다
function cellAt(model, r, c) {
  return model.cells.find(x => x.r === r && x.c === c) || null;
}
function hasMerge(model, r1, c1, r2, c2) {
  return model.merges.some(m => m.r1 === r1 && m.c1 === c1 && m.r2 === r2 && m.c2 === c2);
}

async function main() {
  section('Task 1: 하네스');
  const app = loadApp();
  check('ExcelJS 가 샌드박스에 올라온다', typeof app.ExcelJS === 'object', 'typeof=' + typeof app.ExcelJS);
  check('APP 이 노출된다', typeof app.APP === 'object', 'typeof=' + typeof app.APP);

  section('Task 2: parseCSV');
  const P = app.APP.parseCSV;
  const base = '465,996,97,596,776,652,442,626,637,96,744,406';

  check('12열 한 줄', JSON.stringify(P(base).rows) === JSON.stringify([[465,996,97,596,776,652,442,626,637,96,744,406]]));
  check('후행 콤마', JSON.stringify(P(base + ',').rows) === JSON.stringify(P(base).rows));
  check('콤마 + 공백', JSON.stringify(P(base + ' ,   ').rows) === JSON.stringify(P(base).rows));
  check('빈 열 2개', JSON.stringify(P(base + ',,').rows) === JSON.stringify(P(base).rows));
  check('BOM 과 CRLF', JSON.stringify(P('﻿' + base + '\r\n').rows) === JSON.stringify(P(base).rows));
  check('빈 줄은 건너뛴다', P(base + '\n\n' + base).rows.length === 2);

  const eightCol = '1,2,3,4,5,6,7,8';
  const okEight = P(eightCol + '\n' + eightCol);
  check('12열이 아니어도 스스로 일관되면 성공 (열 개수 동적 감지)', !okEight.error && okEight.rows.length === 2, JSON.stringify(okEight));

  const mismatch = P(eightCol + '\n' + '1,2,3,4,5,6,7');
  check('감지된 열 개수(8)와 다른 줄은 실패', !!mismatch.error && /8열이어야/.test(mismatch.error), mismatch.error);
  check('열 개수 불일치 오류에 줄 번호', /^2번째/.test(mismatch.error || ''), mismatch.error);

  const nan = P('465,x,97,596,776,652,442,626,637,96,744,406');
  check('숫자가 아니면 실패', !!nan.error, JSON.stringify(nan));

  const hole = P('465,,97,596,776,652,442,626,637,96,744,406');
  check('중간 빈 칸은 0 이 아니라 실패', !!hole.error, JSON.stringify(hole));

  const empty = P('\n\n');
  check('유효 행이 없으면 실패', !!empty.error, JSON.stringify(empty));

  const twoLines = P(base + '\n' + base.split(',').slice(0, 11).join(','));
  check('둘째 줄 오류는 줄 번호 2', /^2번째/.test(twoLines.error || ''), twoLines.error);

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

  check('모델의 모든 스타일 토큰이 STYLES 에 있다',
    M.cells.every(c => app.APP.STYLES[c.s] !== undefined),
    [...new Set(M.cells.map(c => c.s))].join(','));

  section('셀 편집 지원 — buildModel 이 본문 데이터 셀에 원본 위치를 표시한다');
  check('C4(첫 데이터 값) 에 편집 태그', at(4, 3) && at(4, 3).edit && at(4, 3).edit.row === 0 && at(4, 3).edit.col === 0,
    JSON.stringify(at(4, 3) && at(4, 3).edit));
  check('N26(마지막 데이터 값) 에 편집 태그', at(26, 14) && at(26, 14).edit && at(26, 14).edit.row === 11 && at(26, 14).edit.col === 11,
    JSON.stringify(at(26, 14) && at(26, 14).edit));
  check('머리글 셀(C3)엔 편집 태그 없음', at(3, 3) && at(3, 3).edit === undefined);
  check('사이드 표(Q3)엔 편집 태그 없음', at(3, 17) && at(3, 17).edit === undefined);
  check('O열 칸막이(O4)엔 편집 태그 없음', at(4, 15) && at(4, 15).edit === undefined);
  const taggedCount = M.cells.filter(c => c.edit).length;
  check('편집 태그가 붙은 셀 수 = 12행 × 12열 = 144', taggedCount === 144, 'tagged=' + taggedCount);

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
  check('mainMaxCol = 15 (O)', M.mainMaxCol === 15, 'mainMaxCol=' + M.mainMaxCol);

  section('열 개수 동적 감지 — 12열이 아닌 파일의 레이아웃');
  const rows8 = [[1, 2, 3, 4, 5, 6, 7, 8], [11, 12, 13, 14, 15, 16, 17, 18]];
  const M8 = app.APP.buildModel(rows8, 'T8', 'L8');
  const at8 = (r, c) => cellAt(M8, r, c);
  check('mainMaxCol = C(3) + NCOL(8) = 11', M8.mainMaxCol === 11, 'mainMaxCol=' + M8.mainMaxCol);
  check('maxCol = mainMaxCol + 3 = 14', M8.maxCol === 14, 'maxCol=' + M8.maxCol);
  check('제목 병합이 A1:11열1 로 줄어든다', hasMerge(M8, 1, 1, 1, 11));
  check('라벨 병합이 13~14열로 밀린다', hasMerge(M8, 1, 13, 1, 14));
  check('colWidths 키가 칸막이+1(12) 로 밀린다', M8.colWidths[12] === 5.125 && M8.colWidths[16] === undefined, JSON.stringify(M8.colWidths));
  check('마지막 데이터 열(10번째) 값', at8(4, 10) && at8(4, 10).v === 8, JSON.stringify(at8(4, 10)));
  check('칸막이(11열) 는 값 없는 data', at8(4, 11) && at8(4, 11).v === null && at8(4, 11).s === 'data');
  check('Q(13열) 첫 값 = 홀수 열 중 최대(17)', at8(3, 13) && at8(3, 13).v === 17, JSON.stringify(at8(3, 13)));

  const rows7 = [[10, 20, 30, 40, 50, 60, 70]];   // 홀수 개(7) — 마지막 열은 짝이 없다
  const M7 = app.APP.buildModel(rows7, 'T7', 'L7');
  const q7 = [], r7 = [];
  for (let r = 3; r < 3 + 4; r++) {
    q7.push(cellAt(M7, r, M7.mainMaxCol + 2).v);
    r7.push(cellAt(M7, r, M7.mainMaxCol + 3).v);
  }
  check('홀수 열 개수(7) — Q는 4개(짝 없는 마지막 열 포함)', JSON.stringify(q7) === JSON.stringify([70, 50, 30, 10]), JSON.stringify(q7));
  check('홀수 열 개수(7) — R은 3개, 남는 칸은 값 없음', JSON.stringify(r7) === JSON.stringify([60, 40, 20, null]), JSON.stringify(r7));

  section('행 선택 필터링과 정렬 방향');
  const sel2 = rows12.map((_, i) => i < 2);   // 앞의 2행만 선택
  const Msel = app.APP.buildModel(rows12, '8/28 111G', 'XX', sel2, 'desc');
  check('선택 2행 → maxRow 은 본문 높이(26) — 사이드 표(12)보다 크다', Msel.maxRow === 26, 'maxRow=' + Msel.maxRow);

  const qSelExpected = [], rSelExpected = [];
  for (let k = 0; k < 6; k++) {
    for (let i = 0; i < 12; i++) {
      if (!sel2[i]) continue;
      qSelExpected.push(rows12[i][2 * k]);
      rSelExpected.push(rows12[i][2 * k + 1]);
    }
  }
  qSelExpected.sort((a, b) => b - a);
  rSelExpected.sort((a, b) => b - a);

  const qSelActual = [];
  for (let r = 3; r < 3 + qSelExpected.length; r++) qSelActual.push(cellAt(Msel, r, 17).v);
  check('선택 2행의 Q 값이 그 두 행만으로 계산된다', JSON.stringify(qSelActual) === JSON.stringify(qSelExpected), JSON.stringify(qSelActual));
  check('선택 범위를 넘는 행엔 Q 셀이 없다', cellAt(Msel, 3 + qSelExpected.length, 17) === null);

  const Masc = app.APP.buildModel(rows12, '8/28 111G', 'XX', sel2, 'asc');
  const qAscExpected = qSelExpected.slice().reverse();
  const qAscActual = [];
  for (let r = 3; r < 3 + qAscExpected.length; r++) qAscActual.push(cellAt(Masc, r, 17).v);
  check('오름차순 지정 시 Q 가 오름차순', JSON.stringify(qAscActual) === JSON.stringify(qAscExpected), JSON.stringify(qAscActual));

  const Mnone = app.APP.buildModel(rows12, '8/28 111G', 'XX', rows12.map(() => false), 'desc');
  check('선택 0행 → 사이드 표 없음, maxRow 은 본문 높이', cellAt(Mnone, 3, 17) === null && Mnone.maxRow === 26, 'maxRow=' + Mnone.maxRow);
  check('선택 0행이어도 P1/P2 헤더는 남는다', cellAt(Mnone, 1, 16) !== null && cellAt(Mnone, 2, 16) !== null);

  const Mexplicit = app.APP.buildModel(rows12, '8/28 111G', 'XX', rows12.map(() => true), 'desc');
  check('selected/sortDir 인자를 생략하면 전체선택+내림차순과 동일', JSON.stringify(Mexplicit) === JSON.stringify(M), 'differs');

  section('편집 모드 — 셀 단위 제외');
  const cellAllOn = rows12.map(row => row.map(() => true));
  const Mcell0 = app.APP.buildModel(rows12, '8/28 111G', 'XX', undefined, undefined, undefined, undefined, cellAllOn);
  check('cellIncluded 를 명시적으로 전체 true 로 주면 생략과 동일', JSON.stringify(Mcell0) === JSON.stringify(M), 'differs');

  const cellExcl = rows12.map(row => row.map(() => true));
  cellExcl[0][0] = false;   // 첫 행 첫 값(홀수 열 1번째, Q 로 감) 만 제외
  const Mexcl = app.APP.buildModel(rows12, '8/28 111G', 'XX', undefined, undefined, undefined, undefined, cellExcl);

  const qColExcl = [], rColExcl = [];
  for (let r = 3; r <= 74; r++) {
    const qc = cellAt(Mexcl, r, 17), rc = cellAt(Mexcl, r, 18);
    qColExcl.push(qc ? qc.v : undefined);
    rColExcl.push(rc ? rc.v : undefined);
  }
  check('Q 는 제외한 값 하나만큼 71개로 줄고 마지막 칸은 값이 빈다', qColExcl.filter(v => v !== null).length === 71 && qColExcl[71] === null, JSON.stringify(qColExcl[71]));
  check('R 은 셀 제외의 영향을 받지 않고 72개 그대로', rColExcl.filter(v => v !== null).length === 72 && rColExcl[71] !== null);

  const lastQCell = cellAt(Mexcl, 74, 17);
  check('짧아진 쪽의 남는 칸도 서식(data)은 유지된다', lastQCell && lastQCell.v === null && lastQCell.s === 'data', JSON.stringify(lastQCell));
  check('maxRow 은 더 긴 쪽(R, 72개)에 맞춰 74 그대로', Mexcl.maxRow === 74, 'maxRow=' + Mexcl.maxRow);

  const selAllButRow0 = rows12.map((_, i) => i !== 0);
  const Mrowoff = app.APP.buildModel(rows12, '8/28 111G', 'XX', selAllButRow0, 'desc', undefined, undefined, cellAllOn);
  const qRowoff = [];
  for (let r = 3; r <= 74; r++) { const c = cellAt(Mrowoff, r, 17); if (c && c.v !== null) qRowoff.push(c.v); }
  check('행 체크가 해제되면 셀 포함 플래그가 true 여도 그 행 값은 제외된다', qRowoff.length === 66, 'len=' + qRowoff.length);

  section('A/B 헤더 텍스트 커스터마이즈');
  check('aText/bText 생략 시 기본값 A/B', at(2, 3).v === 'A' && at(2, 4).v === 'B' && at(2, 17).v === 'A' && at(2, 18).v === 'B');
  const Mcustom = app.APP.buildModel(rows12, '8/28 111G', 'XX', undefined, undefined, '매수', '매도');
  const atC = (r, c) => cellAt(Mcustom, r, c);
  check('C2 커스텀 라벨', atC(2, 3).v === '매수' && atC(2, 3).s === 'accent2', JSON.stringify(atC(2, 3)));
  check('D2 커스텀 라벨', atC(2, 4).v === '매도');
  check('N2 커스텀 라벨(12번째, 짝수)', atC(2, 14).v === '매도');
  check('Q2 커스텀 라벨, accent5 유지', atC(2, 17).v === '매수' && atC(2, 17).s === 'accent5');
  check('R2 커스텀 라벨, accent4 유지', atC(2, 18).v === '매도' && atC(2, 18).s === 'accent4');

  const Mempty = app.APP.buildModel(rows12, '8/28 111G', 'XX', undefined, undefined, '', '');
  check('빈 문자열을 명시하면 그대로 빈 헤더 (undefined 만 기본값으로 대체)', cellAt(Mempty, 2, 3).v === '' && cellAt(Mempty, 2, 4).v === '');

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

  console.log(failures === 0 ? '\n전체 통과' : '\n실패 ' + failures + '건');
  process.exit(failures ? 1 : 0);
}

main().catch(e => { console.error(e); process.exit(1); });
