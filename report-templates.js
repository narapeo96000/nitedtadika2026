/* Read-only report templates. Share the live form schemas, never infer missing answers. */
(function (root) {
  'use strict';
  const escape = value => String(value == null ? '' : value).replace(/[&<>"']/g,
    c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const present = value => value !== undefined && value !== null && String(value).trim() !== '';
  const text = value => present(value) ? escape(value) : '<span class="report-missing">ไม่ระบุ</span>';
  const colors = { 1: '#059669', 2: '#2563eb', 3: '#d97706', 4: '#7c3aed', 5: '#db2777', 6: '#0891b2' };
  const field = (label, value) => `<div class="report-field"><b>${escape(label)}</b><div class="report-value">${text(value)}</div></div>`;
  const title = value => `<h4 class="report-section-title">${escape(value)}</h4>`;
  function header(record, name, no, context) {
    return `<header class="report-form-header"><div>รายงานผลการนิเทศ · ตาดีกาที่เลือก</div><h2>${escape(name)}</h2>
      <h3>${escape(record.name || context.centerName)} · รหัส ${escape(record.id)}</h3>
      <div class="report-meta-grid">${field('วันที่บันทึก', record.timestamp)}${field('ผู้นิเทศ', record.supervisor)}</div></header>`;
  }
  function wrap(record, name, no, context, content) {
    return `<article class="report-form" data-report-form="${no}" style="--report-color:${colors[no]}">
      ${header(record, name, no, context)}${content}
      <footer class="report-form-footer">ใช้ข้อมูลที่บันทึกไว้เพื่อการพัฒนาและให้ความช่วยเหลือ ไม่ใช้เพื่อจัดอันดับ</footer></article>`;
  }
  function metadata(schema, values) {
    if (!schema.meta) return '';
    return (schema.meta.title ? title(schema.meta.title) : '') + '<div class="report-meta-grid">' + schema.meta.fields.map(fd => {
      const key = fd.id === 'f2_note' ? 'context' : fd.id.replace(/^(f\d|g)_/, '');
      return field(fd.label, values[key]);
    }).join('') + '</div>';
  }
  function assessment(schema, answers, notes, general) {
    const scores = Array.isArray(answers) ? answers : [];
    const comments = Array.isArray(notes) ? notes : [];
    // Keep out-of-scale stored values visible rather than silently changing them.
    const scale = general ? ['2', '1', '0', 'N/A'] : ['3', '2', '1', '0'];
    let offset = 0;
    return schema.sections.map((section, sectionIndex) => {
      const rows = section.items.map((item, i) => {
        const index = offset++;
        const raw = scores[index];
        const value = raw === 'NA' ? 'N/A' : String(raw);
        const valid = present(raw) && scale.includes(value);
        const warning = !present(raw) ? 'ไม่มีคะแนนรายข้อที่บันทึกไว้' : !valid ? 'ค่าที่บันทึกไว้: ' + String(raw) + ' (นอกเกณฑ์แบบนี้)' : '';
        return `<tr><td class="report-number">${i + 1}</td><td><div>${escape(item.text)}</div>
          <div class="report-evidence">หลักฐานที่ควรพิจารณา: ${escape(item.evidence)}</div></td>
          ${scale.map(score => `<td class="report-mark" aria-label="${escape(score)}">${valid && value === score ? 'X' : ''}</td>`).join('')}
          <td class="report-value">${text(comments[index])}${warning ? `<div class="report-missing">${escape(warning)}</div>` : ''}</td></tr>`;
      }).join('');
      return title(general ? `${sectionIndex + 4}. ส่วนที่ ${section.code} ${section.title}` : `${section.code} ${section.title}`) +
        `<table class="report-table report-assessment"><colgroup><col style="width:5%"><col style="width:45%">${scale.map(() => '<col style="width:6%">').join('')}<col style="width:26%"></colgroup>
        <thead><tr><th>ข้อ</th><th>ประเด็นที่พิจารณา / หลักฐาน</th>${scale.map(score => `<th>${score}</th>`).join('')}<th>สิ่งที่พบ / ข้อสังเกต</th></tr></thead><tbody>${rows}</tbody></table>`;
    }).join('');
  }
  function standard(record, schema, context) {
    const d = record.details || {};
    const values = d['form' + schema.no] || {};
    const score = record['score' + schema.no];
    return wrap(record, `แบบที่ ${schema.no} ${schema.name}`, schema.no, context,
      metadata(schema, values) + '<p class="report-instruction">ทำเครื่องหมาย X ตามคะแนนรายข้อที่บันทึกไว้ (0–3) ช่องที่ไม่มีข้อมูลจะไม่ถือเป็นคะแนน 0</p>' +
      assessment(schema, (d.answers || {})[schema.no], (d.notes || {})[schema.no], false) +
      field('คะแนนรวมแบบนี้ (ตามบันทึก)', present(score) ? score + ' / ' + schema.pts : null) +
      (schema.no === 2 ? title('ข้อเสนอแนะหลังสังเกต') + field('ข้อเสนอแนะสำหรับครู', values.suggestion) : ''));
  }
  function checks(options, selected) {
    const choices = Array.isArray(selected) ? selected : [];
    return `<div class="report-checks">${options.map(option => `<div><span class="report-checkbox">${choices.includes(option) ? 'X' : '&nbsp;'}</span> ${escape(option)}</div>`).join('')}</div>`;
  }
  function general(record, config, context) {
    const g = (record.details || {}).generalRecord || {};
    const summary = g.summary || {};
    const byId = id => (config.generalFields.find(fd => fd.id === id) || { label: id }).label;
    const fields = ids => ids.map(id => field(byId(id), summary[id])).join('');
    const legend = [['2 ทำได้ชัดเจน', 'พบการดำเนินงานจริงและมีหลักฐานหรือเห็นจากการปฏิบัติ'],
      ['1 กำลังพัฒนา', 'มีการดำเนินงานแล้วบางส่วน แต่ยังไม่ต่อเนื่องหรือยังไม่ชัดเจน'],
      ['0 ต้องได้รับการช่วยเหลือ', 'ยังไม่พบการดำเนินงาน หรือพบปัญหาที่ควรได้รับการช่วยเหลือ'],
      ['N/A', 'ไม่สามารถพิจารณาได้ในครั้งนี้']];
    return wrap(record, config.generalSchema.name + ' (หน้า 6–10)', 6, context,
      title('1. จุดมุ่งหมาย') + '<p>ติดตามการนำหลักสูตรไปใช้และการจัดการเรียนรู้ เพื่อค้นหาจุดแข็ง ประเด็นพัฒนา และความต้องการการสนับสนุน โดยใช้หลักฐานจากสภาพจริง</p>' +
      title('2. ข้อมูลทั่วไป') + metadata({ meta: { ...config.generalSchema.meta, title: '' } }, g.meta || {}) +
      title('3. วิธีใช้แบบนิเทศและคำอธิบายคะแนน') + '<p>ถามผู้รับผิดชอบหรือครู · ดูข้อมูลและหลักฐานที่จำเป็น · สังเกตชั้นเรียนประมาณ 20–30 นาที · สะท้อนผลและกำหนดสิ่งที่จะติดตาม</p>' +
      '<table class="report-table"><thead><tr><th>ระดับ</th><th>ความหมาย</th></tr></thead><tbody>' + legend.map(row => `<tr><td>${row[0]}</td><td>${row[1]}</td></tr>`).join('') + '</tbody></table>' +
      assessment(config.generalSchema, g.answers, g.notes, true) +
      title('7. สรุปผลการนิเทศแบบกระชับ') + fields(['g_strengths', 'g_gaps']) +
      title('7.3 ศูนย์ต้องการความช่วยเหลือเรื่องใด') + checks(config.helpOptions, summary.helpOptions) + field('รายละเอียดความช่วยเหลือ / อื่น ๆ', summary.g_help) +
      title('8. ข้อตกลงเพื่อการพัฒนา') + fields(['g_agreement', 'g_action', 'g_supporter', 'g_followDate', 'g_followLook']) +
      title('9. แบบติดตามครั้งต่อไป') + '<table class="report-table"><thead><tr><th>คำถามติดตาม</th><th>ผลการติดตาม</th></tr></thead><tbody>' +
      `<tr><td>1. เรื่องที่ตกลงจะพัฒนา ได้ดำเนินการแล้วหรือไม่</td><td>${checks(config.followOptions, summary.followOptions)}${field('อื่น ๆ (ระบุ)', summary.followOther)}</td></tr>` +
      `<tr><td>2. หลังดำเนินการแล้ว มีอะไรเปลี่ยนแปลง</td><td class="report-value">${text(summary.g_change)}</td></tr>` +
      `<tr><td>3. มีหลักฐานอะไรที่แสดงถึงการเปลี่ยนแปลง</td><td class="report-value">${text(summary.g_evidence)}</td></tr></tbody></table>` +
      fields(['g_followStatus']) + '<div class="report-signatures"><div>ลงชื่อผู้รับผิดชอบศูนย์ ........................................</div><div>ลงชื่อผู้นิเทศ ........................................</div><div>วันที่ ........................................</div></div>');
  }
  function summaryForm(record, config, context) {
    const d = record.details || {}, values = d.form5 || {};
    const plan = Array.isArray(d.actionPlan) ? d.actionPlan : [];
    const extra = [['help', 'ความช่วยเหลือที่ต้องการ'], ['agreement', 'ข้อตกลงเพื่อการพัฒนา'], ['centerAction', 'สิ่งที่ศูนย์จะดำเนินการ'],
      ['supporter', 'ผู้สนับสนุน'], ['followDate', 'กำหนดติดตาม'], ['followLook', 'ประเด็นที่ติดตาม'], ['followStatus', 'ผลการติดตาม'],
      ['change', 'การเปลี่ยนแปลง'], ['evidence', 'หลักฐาน'], ['nextAction', 'สิ่งที่จะดำเนินการต่อ'], ['summary', 'สรุปสำหรับผู้นิเทศ']];
    return wrap(record, 'แบบที่ 5 สรุปผลรายศูนย์และแผนพัฒนา', 5, context,
      title('สรุปผลการนิเทศ') + '<table class="report-table"><thead><tr><th>แบบนิเทศ</th><th>คะแนนตามบันทึก</th></tr></thead><tbody>' +
      config.forms.map(f => `<tr><td>แบบที่ ${f.no} ${escape(f.name)}</td><td>${text(record['score' + f.no])} / ${f.pts}</td></tr>`).join('') +
      `<tr><td>คะแนนรวม / ร้อยละ / ระดับ (ตามบันทึกเดิม)</td><td>${text(record.totalScore)} / ${text(record.pct)}% / ${text(record.level)}</td></tr></tbody></table>` +
      config.summaryFields.map(fd => field(fd.label, values[fd.id.replace('f5_', '')])).join('') +
      title('แผนพัฒนาร่วมกัน (Action Plan)') + '<table class="report-table"><thead><tr><th>ประเด็นพัฒนา</th><th>กิจกรรม / วิธีดำเนินการ</th><th>ผู้รับผิดชอบ</th><th>กำหนดเสร็จ</th><th>หลักฐานความสำเร็จ</th></tr></thead><tbody>' +
      (plan.length ? plan.map(row => '<tr>' + ['topic', 'activity', 'owner', 'deadline', 'evidence'].map(key => `<td class="report-value">${text(row[key])}</td>`).join('') + '</tr>').join('') : '<tr><td colspan="5">ไม่มีแผนพัฒนาที่บันทึกไว้</td></tr>') + '</tbody></table>' +
      extra.filter(([key]) => present(values[key])).map(([key, label]) => field(label, values[key])).join('') +
      field('เอกสาร / รูปภาพประกอบ (ชื่อไฟล์ตามบันทึก)', values.files));
  }
  function hasGeneral(record) {
    const g = (record.details || {}).generalRecord || {};
    return Object.values(g.meta || {}).some(present) || Object.values(g.summary || {}).some(v => Array.isArray(v) ? v.length : present(v)) ||
      (Array.isArray(g.answers) && g.answers.some(v => present(v) && String(v) !== '0')) || (Array.isArray(g.notes) && g.notes.some(present));
  }
  function hasLegacy(record) {
    const d = record.details || {};
    return !hasGeneral(record) || Object.values(d.answers || {}).some(Array.isArray) ||
      [1, 2, 3, 4, 5].some(n => Object.values(d['form' + n] || {}).some(present)) ||
      [1, 2, 3, 4].some(n => Number(record['score' + n]) > 0);
  }
  function render(records, context, config) {
    if (!records.length) return '<p class="report-empty">ไม่พบข้อมูลตามเงื่อนไขที่เลือก</p>';
    return records.map(record => {
      if (context.kind === 'form8' || context.kind === 'general') return general(record, config, context);
      if (context.kind === 'form5') return summaryForm(record, config, context);
      const schema = config.forms.find(f => 'form' + f.no === context.kind);
      if (schema) return standard(record, schema, context);
      if (context.kind === 'summary') return wrap(record, 'ข้อมูลประกอบรายงานสรุปด้วย AI', 5, context,
        field('แบบนิเทศตามบันทึก', record.formType) + field('จุดเด่น', ((record.details || {}).form5 || {}).strengths || (((record.details || {}).generalRecord || {}).summary || {}).g_strengths));
      return (hasLegacy(record) ? config.forms.map(f => standard(record, f, context)).join('') + summaryForm(record, config, context) : '') +
        (hasGeneral(record) ? general(record, config, context) : '');
    }).join('');
  }
  root.ReportTemplates = { render, hasGeneral, hasLegacy };
})(globalThis);
