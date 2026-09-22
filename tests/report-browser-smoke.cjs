// Full-page browser smoke test; every request is fulfilled locally with synthetic data.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { chromium } = require('playwright');
const { fixture, html, root } = require('./report-fixtures.cjs');
(async () => {
  const browser = await chromium.launch({ channel: 'chrome', headless: true });
  try {
    const page = await browser.newPage({ serviceWorkers: 'block', viewport: { width: 390, height: 844 } });
    const errors = [];
    page.on('pageerror', error => errors.push(error.message));
    const { record, general } = fixture();
    const center = { id: record.id, name: record.name, type: record.type, teachers: {}, students: {} };
    await page.addInitScript(() => { window.tailwind = {}; window.Swal = { fire: async () => ({}), close() {}, showLoading() {} }; });
    await page.route('**/*', async route => {
      const url = new URL(route.request().url());
      if (url.hostname === 'report.test') {
        if (url.pathname === '/') return route.fulfill({ contentType: 'text/html', body: html });
        if (url.pathname === '/report-templates.js') return route.fulfill({ contentType: 'application/javascript', body: fs.readFileSync(path.join(root, 'report-templates.js'), 'utf8') });
      }
      if (url.hostname === 'script.google.com') {
        const request = route.request().method() === 'POST' ? route.request().postDataJSON() : { action: url.searchParams.get('action') };
        let response = { success: false };
        if (request.action === 'getTadikaList') response = { success: true, data: [center] };
        if (request.action === 'getReportEvaluations') {
          assert.equal(request.payload.centerId, record.id);
          response = { success: true, data: [record, general] };
        }
        if (request.action === 'generateReportAI') response = { success: true, text: 'สรุปทดสอบ ไม่ได้เรียก AI จริง' };
        return route.fulfill({ contentType: 'application/json', body: JSON.stringify(response), headers: { 'access-control-allow-origin': '*' } });
      }
      return route.fulfill({ contentType: 'application/javascript', body: '' });
    });
    await page.goto('https://report.test/');
    await page.waitForFunction(() => tadikaList.length === 1);
    await page.evaluate(() => {
      currentUser = { sessionToken: 'synthetic-session' };
      document.getElementById('loginSection').classList.add('section-hidden');
      document.getElementById('appSection').classList.remove('section-hidden');
      openReportMenu();
    });
    await page.locator('#reportCenterSearch').fill('TEST-001');
    await page.locator('#reportCenter').selectOption(record.id);
    for (const [kind, no] of [['form1', '1'], ['form2', '2'], ['form3', '3'], ['form4', '4'], ['form5', '5'], ['general', '6']]) {
      await page.locator('#reportKind').selectOption(kind);
      await page.getByRole('button', { name: '🔎 แสดงตัวอย่าง', exact: true }).click();
      await page.waitForFunction(() => document.getElementById('reportStatus').textContent.startsWith('แสดงตัวอย่างแล้ว'));
      assert.equal(await page.locator('#reportPreview article').count(), 1);
      assert.equal(await page.locator('#reportPreview article').getAttribute('data-report-form'), no);
    }
    assert.match(await page.locator('#reportPreview').innerText(), /อบรมการใช้สื่อดิจิทัล/);
    await page.getByRole('button', { name: '✨ สรุปด้วย AI', exact: true }).click();
    await page.waitForFunction(() => document.getElementById('reportStatus').textContent.startsWith('AI สรุปเรียบร้อย'));
    assert.match(await page.locator('#reportAiOutput').innerText(), /ไม่ได้เรียก AI จริง/);
    assert.equal(await page.evaluate(() => document.documentElement.scrollWidth > innerWidth + 2), false);
    await page.locator('#reportCenterSearch').fill('not found');
    assert.equal(await page.locator('#reportPreview article').count(), 0);
    assert.deepEqual(errors, []);
    console.log('Browser smoke passed: six form layouts, mobile search, isolated AI mock, stale preview clearing, no JS errors.');
  } finally { await browser.close(); }
})().catch(error => { console.error(error); process.exitCode = 1; });
