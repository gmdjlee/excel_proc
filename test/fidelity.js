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
