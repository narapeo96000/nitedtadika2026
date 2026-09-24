const { test } = require('node:test');
const assert = require('node:assert/strict');
const { fixture } = require('./report-fixtures.cjs');

test('each scored form uses all and only its live-schema questions, evidence, metadata and notes', () => {
  const { templates, config, record, context } = fixture();
  for (const f of config.forms) {
    const output = templates.render([record], { ...context, kind: 'form' + f.no }, config);
    assert.match(output, new RegExp('data-report-form="' + f.no + '"'));
    assert.ok(output.includes(f.name));
    f.sections.forEach(s => s.items.forEach(item => { assert.ok(output.includes(item.text)); assert.ok(output.includes(item.evidence)); }));
    assert.equal((output.match(/class="report-mark"/g) || []).length, f.sections.flatMap(s => s.items).length * 4);
    assert.ok(output.includes('พบหลักฐานจากแผนการสอน'));
    assert.ok(output.includes(record['score' + f.no] + ' / ' + f.pts));
    config.forms.filter(other => other.no !== f.no).forEach(other => assert.ok(!output.includes(other.sections[0].items[0].text)));
  }
});

test('general form preserves N/A, zero, notes, help and follow-up checkbox details', () => {
  const { templates, config, general, context } = fixture();
  const output = templates.render([general], { ...context, kind: 'form8' }, config);
  assert.equal((output.match(/>X<\/td>/g) || []).length, 12);
  assert.equal((output.match(/aria-label="N\/A">X/g) || []).length, 2);
  assert.equal((output.match(/aria-label="0">X/g) || []).length, 2);
  for (const value of ['คำอธิบายคะแนน', 'อบรมการใช้สื่อดิจิทัล', 'กำลังทำ', 'รอรับสื่อเพิ่มเติม', 'บันทึกการสังเกตและผลงาน', 'ลงชื่อผู้นิเทศ']) assert.ok(output.includes(value));
  assert.doesNotMatch(output, /ร้อยละเฉลี่ย|คะแนนรวมทั้งหมด/);
});

test('missing scores are not inferred from totals; zero remains an actual marked answer', () => {
  const { templates, config, record, context } = fixture();
  record.details.answers = {};
  let output = templates.render([record], context, config);
  assert.doesNotMatch(output, />X<\/td>/);
  assert.match(output, /ไม่มีคะแนนรายข้อที่บันทึกไว้/);
  record.details.answers = { 1: [0, null, '', undefined, 'unexpected'] };
  output = templates.render([record], context, config);
  assert.equal((output.match(/>X<\/td>/g) || []).length, 1);
  assert.match(output, /นอกเกณฑ์แบบนี้/);
});

test('summary form includes recorded development plan; all renders separate forms and visits', () => {
  const { templates, config, record, general, context } = fixture();
  let output = templates.render([record], { ...context, kind: 'form5' }, config);
  assert.ok(output.includes('ติดตามผู้เรียนรายบุคคล'));
  assert.ok(output.includes('ตุลาคม 2569'));
  output = templates.render([record, general], { ...context, kind: 'all' }, config);
  assert.equal((output.match(/<article /g) || []).length, 6);
});

test('stored text cannot inject HTML; renderer never mutates source records', () => {
  const { templates, config, record, context } = fixture();
  record.details.notes[1][0] = '<img src=x onerror=alert(1)>';
  record.name = '<script>bad()</script>';
  const before = JSON.stringify(record);
  const output = templates.render([record], context, config);
  assert.ok(output.includes('&lt;img'));
  assert.doesNotMatch(output, /<img|<script>/);
  assert.equal(JSON.stringify(record), before);
});
