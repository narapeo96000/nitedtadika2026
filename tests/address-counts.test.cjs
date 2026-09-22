const { test } = require('node:test');
const assert = require('node:assert/strict');
const vm = require('node:vm');
const { readFileSync } = require('node:fs');
const path = require('node:path');
const root = path.join(__dirname, '..');
const html = readFileSync(path.join(root, 'index.html'), 'utf8');
function client() {
  const nodes = new Map();
  const get = id => {
    if (!nodes.has(id)) nodes.set(id, { value: '', innerHTML: '', innerText: '', classList: { add() {} } });
    return nodes.get(id);
  };
  const ctx = vm.createContext({ document: { getElementById: get }, autoReg: { addr: { search: 'addrSearch', selected: null } },
    tadikaList: [], escHtml: value => String(value || ''), tadikaAddr: t => t.address || '', typeBadge: () => '', statusBadge: () => '' });
  vm.runInContext(html.slice(html.indexOf('    function addrCountNumber('), html.indexOf('    // ---- สถิติระบบนิเทศออนไลน์ (หน้าแรก)')), ctx);
  return { ctx, get };
}
test('O/P/Q and R/S/T map to male/female/total from ADDR_TADEKA row 6 only', () => {
  const row = Array(22).fill('');
  row[0] = 'T001'; row[3] = 'Test center';
  [11, 12, 30, 41, 42, 90].forEach((value, i) => { row[14 + i] = value; });
  const ctx = vm.createContext({ SpreadsheetApp: { openById: () => ({ getSheetByName: name => {
    assert.equal(name, 'ADDR_TADEKA');
    return { getLastRow: () => 6, getRange: (...range) => { assert.deepEqual(range, [6, 1, 1, 22]); return { getValues: () => [row] }; } };
  } }) } });
  vm.runInContext(readFileSync(path.join(root, 'apps-script/รหัส.js'), 'utf8'), ctx);
  const result = JSON.parse(JSON.stringify(ctx.getTadikaList()));
  assert.deepEqual(result.data[0].teachers, { male: 11, female: 12, total: 30 });
  assert.deepEqual(result.data[0].students, { male: 41, female: 42, total: 90 });
});
test('counts preserve recorded totals and zero, distinguish blanks, format numeric strings', () => {
  const { ctx } = client();
  assert.equal(ctx.formatAddrBreakdown({ male: 3, female: 4, total: 0 }), '3 / 4 / <b>0</b>');
  assert.equal(ctx.formatAddrBreakdown({ male: 3, female: 4, total: '' }), '3 / 4 / <b>—</b>');
  assert.equal(ctx.formatAddrBreakdown({ male: '1,000', female: ' 2 ', total: '1,100' }), '1,000 / 2 / <b>1,100</b>');
  assert.equal(ctx.formatAddrBreakdown(), '— / — / <b>—</b>');
  for (const value of [null, undefined, '', ' ', 'bad', -1]) assert.equal(ctx.formatAddrCount(value), '—');
});
test('table shows two gender breakdown columns, keeps foreigners and sorts by recorded totals', () => {
  const { ctx, get } = client();
  ctx.tadikaList = [
    { id: 'A', name: 'A', teachers: { male: 8, female: 2, total: 10 }, students: { male: 1, female: 1, total: 99 }, foreignStudents: 0 },
    { id: 'B', name: 'B', teachers: { male: 1, female: 20, total: 21 }, students: { male: 5, female: 6, total: 11 }, foreignStudents: '' }
  ];
  ctx.renderAddrHeader(); ctx.renderAddrTable();
  assert.equal((get('addrHead').innerHTML.match(/<th /g) || []).length, 10);
  assert.match(get('addrBody').innerHTML, /ผู้สอน \(ชาย \/ หญิง \/ รวม\)/);
  assert.match(get('addrBody').innerHTML, /1 \/ 1 \/ <b>99<\/b>/);
  assert.match(get('addrBody').innerHTML, /ต่างชาติ: 0/);
  ctx.sortAddr('students');
  assert.ok(get('addrBody').innerHTML.indexOf('🆔 B') < get('addrBody').innerHTML.indexOf('🆔 A'));
  ctx.sortAddr('students');
  assert.ok(get('addrBody').innerHTML.indexOf('🆔 A') < get('addrBody').innerHTML.indexOf('🆔 B'));
  get('addrSearch').value = 'missing'; ctx.renderAddrTable();
  assert.match(get('addrBody').innerHTML, /colspan="5"/);
});
