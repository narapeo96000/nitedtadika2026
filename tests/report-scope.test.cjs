const { test } = require('node:test');
const assert = require('node:assert/strict');
const { readFileSync } = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');
const root = path.join(__dirname, '..');
const html = readFileSync(path.join(root, 'index.html'), 'utf8');
const backend = readFileSync(path.join(root, 'apps-script/รหัส.js'), 'utf8');

function server() {
  const readSheets = [];
  const rows = ['001', '002', '', ' 001 '].map(id => ['2026-09-22', id, 'Same name', '', 1, 2, 3, 4, 10, 7, '', '{}', 'Inspector']);
  const ctx = vm.createContext({
    CacheService: { getScriptCache: () => ({ get: key => key === 'reportSession:valid' ? '{"username":"tester"}' : null }) },
    SpreadsheetApp: { openById: () => ({ getSheetByName: name => {
      readSheets.push(name);
      return { getLastRow: () => rows.length + 1, getRange: () => ({ getValues: () => rows }) };
    } }) }
  });
  vm.runInContext(backend, ctx);
  ctx.formatDate = x => x;
  ctx.getSystemSettings = () => ({ geminiKey: 'test-placeholder' });
  ctx.callGeminiAPI = () => { ctx.aiCalls++; return 'Summary'; };
  ctx.aiCalls = 0;
  return { ctx, readSheets };
}

test('backend returns only exact selected Tadika ID, never Pondok or same-name centers', () => {
  const { ctx, readSheets } = server();
  const result = ctx.getReportEvaluations({ sessionToken: 'valid', type: 'ตาดีกา', centerId: '001' });
  assert.equal(result.success, true);
  assert.deepEqual(Array.from(result.data, r => r.id), ['001', '001']);
  assert.deepEqual(readSheets, ['DATA_TADEKA']);
});

test('backend rejects missing ID, broad type and expired sessions before reading sheets', () => {
  const { ctx, readSheets } = server();
  for (const filters of [
    { sessionToken: 'valid', type: 'ตาดีกา' },
    { sessionToken: 'valid', type: 'ตาดีกา', centerId: ' ' },
    { sessionToken: 'valid', type: 'ทั้งหมด', centerId: '001' },
    { sessionToken: 'valid', type: 'ปอเนาะ', centerId: '001' },
    { type: 'ตาดีกา', centerId: '001' }
  ]) assert.equal(ctx.getReportEvaluations(filters).success, false);
  assert.deepEqual(readSheets, []);
});

test('AI rejects mixed-center and mixed-type records, including after the 100-record cap', () => {
  const { ctx } = server();
  const good = { id: '001', type: 'ตาดีกา' };
  const payload = { sessionToken: 'valid', report: { filters: { centerId: '001', type: 'ตาดีกา' } }, records: [good] };
  assert.equal(ctx.generateReportAI(payload).success, true);
  for (const bad of [{ id: '002', type: 'ตาดีกา' }, { id: '001', type: 'ปอเนาะ' }, {}, null]) {
    assert.equal(ctx.generateReportAI({ ...payload, records: [...Array(100).fill(good), bad] }).success, false);
  }
  assert.equal(ctx.generateReportAI({ ...payload, report: {} }).success, false);
  assert.equal(ctx.aiCalls, 1);
});

function client() {
  const elements = new Map();
  const pending = [];
  const element = id => {
    if (!elements.has(id)) {
      const classes = new Set(id === 'reportPreviewCard' ? ['section-hidden'] : []);
      elements.set(id, { value: id === 'reportKind' ? 'all' : '', textContent: '', innerHTML: '',
        classList: { add: x => classes.add(x), remove: x => classes.delete(x), contains: x => classes.has(x) },
        addEventListener() {}, scrollIntoView() {} });
    }
    return elements.get(id);
  };
  const ctx = vm.createContext({
    selectedTadika: { id: '001', type: 'ตาดีกา', name: 'Center A' }, currentUser: { sessionToken: 'valid' },
    document: { getElementById: element },
    api: (action, payload) => new Promise(resolve => pending.push({ action, payload, resolve })),
    escHtml: x => String(x || '').replaceAll('<', '&lt;'), renderChatReply: x => x,
    Swal: { fire() {} }, window: { print: () => ctx.printCount++ }, printCount: 0
  });
  const start = html.indexOf('    let reportSourceRecords =');
  const end = html.indexOf('    // ---- ผู้ดูแลระบบ:', start);
  assert.ok(start > 0 && end > start);
  vm.runInContext(html.slice(start, end), ctx);
  return { ctx, element, pending };
}
const record = (id = '001', type = 'ตาดีกา') => ({ id, type, name: 'Center ' + id, timestamp: '2026-09-22', details: {}, pct: 50 });
const finish = req => req.resolve({ success: true, data: [record(), record('002'), record('001', 'ปอเนาะ')] });

test('preview and PDF use the selected center; stale filters block printing', async () => {
  const { ctx, element, pending } = client();
  const task = ctx.buildReportPreview();
  assert.equal(pending[0].payload.centerId, '001');
  finish(pending[0]);
  const result = await task;
  assert.equal(result.records.length, 1);
  assert.match(element('reportPreview').innerHTML, /Center A/);
  assert.doesNotMatch(element('reportPreview').innerHTML, /Center 002|ปอเนาะ/);
  ctx.printReportPDF();
  assert.equal(ctx.printCount, 1);
  element('reportFrom').value = '2026-09-23';
  ctx.printReportPDF();
  assert.equal(ctx.printCount, 1);
});

test('no center means no API request; empty selected-center results remain empty', async () => {
  const { ctx, element, pending } = client();
  ctx.selectedTadika = null;
  assert.equal(await ctx.buildReportPreview(), null);
  assert.equal(pending.length, 0);
  ctx.selectedTadika = { id: '003', type: 'ตาดีกา', name: 'Empty center' };
  const task = ctx.buildReportPreview();
  pending[0].resolve({ success: true, data: [] });
  assert.equal((await task).records.length, 0);
  assert.match(element('reportPreview').innerHTML, /Empty center/);
});

test('late preview responses cannot restore a previous center', async () => {
  const { ctx, element, pending } = client();
  const task = ctx.buildReportPreview();
  ctx.selectedTadika = { id: '002', type: 'ตาดีกา', name: 'Center B' };
  ctx.invalidateReport();
  finish(pending[0]);
  assert.equal(await task, null);
  assert.equal(element('reportPreview').innerHTML, '');
  assert.ok(element('reportPreviewCard').classList.contains('section-hidden'));
});

test('AI sends selected ID only and ignores responses after center/filter changes or close', async () => {
  for (const change of ['center', 'filter', 'close']) {
    const { ctx, element, pending } = client();
    const task = ctx.generateReportSummary();
    finish(pending[0]);
    await new Promise(resolve => setImmediate(resolve));
    assert.equal(pending[1].action, 'generateReportAI');
    assert.equal(pending[1].payload.report.filters.centerId, '001');
    assert.deepEqual(Array.from(pending[1].payload.records, r => r.id), ['001']);
    if (change === 'center') ctx.selectedTadika = { id: '002', type: 'ตาดีกา', name: 'Center B' };
    if (change === 'filter') element('reportKind').value = 'general';
    if (change === 'close') ctx.closeReportMenu(); else ctx.invalidateReport();
    pending[1].resolve({ success: true, text: 'Old center summary' });
    await task;
    assert.equal(element('reportPreview').innerHTML, '');
    assert.doesNotMatch(element('reportAiOutput').innerHTML, /Old center summary/);
  }
});

test('all inline scripts and backend parse', () => {
  for (const match of html.matchAll(/<script(?:\s[^>]*)?>([\s\S]*?)<\/script>/g)) new vm.Script(match[1]);
  new vm.Script(backend);
});
