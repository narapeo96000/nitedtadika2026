// Local-only visual QA with synthetic data; no sessions, sheet reads or AI requests.
const fs = require('node:fs');
const path = require('node:path');
const { fixture, html, root } = require('./report-fixtures.cjs');
const { chromium } = require('playwright');
(async () => {
  const output = path.join(root, 'tmp/pdfs');
  fs.mkdirSync(output, { recursive: true });
  const browser = await chromium.launch({ channel: 'chrome', headless: true });
  try {
    const page = await browser.newPage({ viewport: { width: 1000, height: 1000 } });
    // Use the real DOM and print rules, but strip executable scripts and remote assets.
    const isolatedHtml = html.replace(/<script\b[^>]*>[\s\S]*?<\/script>/gi, '').replace(/<link\b[^>]*>/gi, '');
    const { templates, config, record, general, context } = fixture();
    for (const kind of ['form1', 'form2', 'form3', 'form4', 'form5', 'general']) {
      const content = templates.render([kind === 'general' ? general : record], { ...context, kind }, config);
      await page.setContent(isolatedHtml);
      await page.evaluate(content => {
        document.getElementById('loginSection').classList.add('section-hidden');
        document.getElementById('appSection').classList.remove('section-hidden');
        document.getElementById('reportPreviewCard').classList.remove('section-hidden');
        document.getElementById('reportPreview').innerHTML = '<div class="report-forms">' + content + '</div>';
        // Exercise a long, visible sibling to detect hidden-content blank print pages.
        const extra = document.createElement('div'); extra.style.height = '6000px'; extra.textContent = 'NON_REPORT_CONTENT';
        document.getElementById('appSection').appendChild(extra);
      }, content);
      await page.evaluate(() => document.fonts.ready);
      await page.pdf({ path: path.join(output, kind + '.pdf'), format: 'A4', printBackground: true, preferCSSPageSize: true });
      await page.locator('#reportPreviewCard').screenshot({ path: path.join(output, kind + '-screen.png') });
      const overflow = await page.locator('.report-table').evaluateAll(tables => tables.filter(t => t.scrollWidth > t.clientWidth + 2).length);
      if (overflow) throw new Error(kind + ': table overflow');
      console.log(kind + ': rendered A4 and screen preview');
    }
  } finally { await browser.close(); }
})().catch(error => { console.error(error); process.exitCode = 1; });
