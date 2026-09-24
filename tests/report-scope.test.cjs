const { test } = require('node:test');
const assert = require('node:assert/strict');
const { readFileSync } = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');
const { installTemplates } = require('./report-fixtures.cjs');
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
      const attributes = new Map();
      elements.set(id, { id, value: id === 'reportKind' ? 'all' : '', textContent: '', innerHTML: '',
        classList: { add: x => classes.add(x), remove: x => classes.delete(x), contains: x => classes.has(x) },
        setAttribute: (key, value) => attributes.set(key, value), removeAttribute: key => attributes.delete(key), getAttribute: key => attributes.get(key),
        addEventListener() {}, scrollIntoView() {} });
    }
    return elements.get(id);
  };
  const ctx = vm.createContext({
    selectedTadika: { id: '001', type: 'ตาดีกา', name: 'Center A' }, currentUser: { sessionToken: 'valid' },
    tadikaList: [
      { id: '001', type: 'ตาดีกา', name: 'Center A', dist: 'เมือง', subdist: 'บางนาค' },
      { id: '002', type: 'ตาดีกา', name: 'Center B', dist: 'ตากใบ', subdist: 'เจ๊ะเห' },
      { id: '003', type: 'ตาดีกา', name: 'Empty center' },
      { id: '004', type: 'ปอเนาะ', name: 'Pondok' }
    ],
    document: { getElementById: element },
    api: (action, payload) => new Promise(resolve => pending.push({ action, payload, resolve })),
    escHtml: x => String(x || '').replaceAll('<', '&lt;'), renderChatReply: x => x,
    Swal: { fire() {} }, window: { print: () => ctx.printCount++ }, printCount: 0
  });
  const start = html.indexOf('    let reportSourceRecords =');
  const end = html.indexOf('    // ---- ผู้ดูแลระบบ:', start);
  assert.ok(start > 0 && end > start);
  installTemplates(ctx);
  vm.runInContext(html.slice(start, end), ctx);
  ctx.openReportMenu();
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
  element('reportCenter').value = '';
  assert.equal(await ctx.buildReportPreview(), null);
  assert.equal(pending.length, 0);
  element('reportCenter').value = '003';
  const task = ctx.buildReportPreview();
  pending[0].resolve({ success: true, data: [] });
  assert.equal((await task).records.length, 0);
  assert.match(element('reportPreview').innerHTML, /Empty center/);
});

test('late preview responses cannot restore a previous center', async () => {
  const { ctx, element, pending } = client();
  const task = ctx.buildReportPreview();
  element('reportCenter').value = '002';
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
    if (change === 'center') element('reportCenter').value = '002';
    if (change === 'filter') element('reportKind').value = 'form8';
    if (change === 'close') ctx.closeReportMenu(); else ctx.invalidateReport();
    pending[1].resolve({ success: true, text: 'Old center summary' });
    await task;
    assert.equal(element('reportPreview').innerHTML, '');
    assert.doesNotMatch(element('reportAiOutput').innerHTML, /Old center summary/);
  }
});

test('report center picker is independent from the evaluation center and retained on reopen', async () => {
  const { ctx, element, pending } = client();
  assert.equal(element('reportCenter').value, '001');
  assert.doesNotMatch(element('reportCenterOptions').innerHTML, /Pondok/);
  element('reportCenter').value = '002';
  ctx.invalidateReport();
  const task = ctx.buildReportPreview();
  assert.equal(pending[0].payload.centerId, '002');
  finish(pending[0]);
  const preview = await task;
  assert.equal(preview.records.length, 1);
  assert.equal(preview.records[0].id, '002');
  assert.match(element('reportPreview').innerHTML, /Center B/);
  assert.doesNotMatch(element('reportPreview').innerHTML, /ภาพรวมผลการนิเทศ|Center 001/);
  assert.equal(ctx.selectedTadika.id, '001');
  ctx.closeReportMenu();
  ctx.openReportMenu();
  assert.equal(element('reportCenter').value, '002');
});

test('unknown or Pondok selections cannot issue report requests', async () => {
  const { ctx, element, pending } = client();
  for (const id of ['missing', '004']) {
    element('reportCenter').value = id;
    assert.equal(await ctx.buildReportPreview(), null);
  }
  assert.equal(pending.length, 0);
});

test('search filters by name, ID, district, subdistrict and combined terms', () => {
  const { ctx, element } = client();
  for (const query of ['center b', '002', 'ตากใบ', 'เจ๊ะเห', '  CENTER   ตากใบ  ']) {
    element('reportCenterSearch').value = query;
    ctx.filterReportCenters();
    assert.match(element('reportCenterOptions').innerHTML, /Center B/);
    assert.doesNotMatch(element('reportCenterOptions').innerHTML, /Center A|Empty center|Pondok/);
    assert.equal(element('reportCenter').value, '');
    assert.equal(element('reportCenterCount').textContent, 'พบ 1 ตาดีกา จากทั้งหมด 3 แห่ง');
    assert.equal(ctx.selectedTadika.id, '001');
  }
});

test('no search matches clears selection and blocks reports; clearing search restores choices without auto-select', async () => {
  const { ctx, element, pending } = client();
  element('reportCenterSearch').value = 'not found';
  ctx.filterReportCenters();
  assert.match(element('reportCenterOptions').innerHTML, /ไม่พบตาดีกาที่ค้นหา/);
  assert.equal(element('reportCenter').value, '');
  assert.equal(await ctx.buildReportPreview(), null);
  assert.equal(pending.length, 0);
  element('reportCenterSearch').value = '';
  ctx.filterReportCenters();
  assert.match(element('reportCenterOptions').innerHTML, /Center A/);
  assert.match(element('reportCenterOptions').innerHTML, /Center B/);
  assert.equal(element('reportCenter').value, '');
});

test('editing search requires re-selection and invalidates pending previews', async () => {
  const { ctx, element, pending } = client();
  const task = ctx.buildReportPreview();
  element('reportCenterSearch').value = 'เมือง';
  ctx.filterReportCenters();
  assert.equal(element('reportCenter').value, '');
  finish(pending[0]);
  assert.equal(await task, null);
  assert.equal(element('reportPreview').innerHTML, '');
});

test('one search box chooses a center by click or keyboard without changing the evaluation center', () => {
  const { ctx, element } = client();
  const key = name => ctx.reportCenterKey({ key: name, preventDefault() {} });
  element('reportCenterSearch').value = 'ตากใบ';
  ctx.filterReportCenters();
  assert.equal(ctx.reportContext(), null);
  ctx.chooseReportCenter(0);
  assert.equal(ctx.reportContext().centerId, '002');
  assert.equal(element('reportCenterSearch').value, 'Center B · รหัส 002');
  assert.equal(element('reportCenterSearch').getAttribute('aria-expanded'), 'false');
  element('reportCenterSearch').value = '';
  ctx.filterReportCenters();
  key('ArrowDown'); key('ArrowDown'); key('Enter');
  assert.equal(ctx.reportContext().centerId, '002');
  assert.equal(ctx.selectedTadika.id, '001');
  ctx.openReportCenterSuggestions();
  key('Enter');
  assert.equal(ctx.reportContext().centerId, '002');
  ctx.openReportCenterSuggestions();
  key('Escape');
  assert.equal(element('reportCenterSearch').getAttribute('aria-expanded'), 'false');
  assert.equal(ctx.reportContext().centerId, '002');
  element('reportCenterSearch').value = 'not found';
  ctx.filterReportCenters(); key('Enter');
  assert.equal(ctx.reportContext(), null);
});

test('all inline scripts and backend parse', () => {
  for (const match of html.matchAll(/<script(?:\s[^>]*)?>([\s\S]*?)<\/script>/g)) new vm.Script(match[1]);
  new vm.Script(backend);
});
