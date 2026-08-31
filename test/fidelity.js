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

  const short = P(base.split(',').slice(0, 11).join(','));
  check('11열은 실패', !!short.error, JSON.stringify(short));
  check('11열 오류에 줄 번호', /^1번째/.test(short.error || ''), short.error);

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

  console.log(failures === 0 ? '\n전체 통과' : '\n실패 ' + failures + '건');
  process.exit(failures ? 1 : 0);
}

main().catch(e => { console.error(e); process.exit(1); });
