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

  console.log(failures === 0 ? '\n전체 통과' : '\n실패 ' + failures + '건');
  process.exit(failures ? 1 : 0);
}

main().catch(e => { console.error(e); process.exit(1); });
