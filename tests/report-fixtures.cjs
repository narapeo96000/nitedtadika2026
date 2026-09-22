const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const root = path.join(__dirname, '..');
const html = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
function installTemplates(ctx) {
  vm.runInContext(html.slice(html.indexOf('    const S = '), html.indexOf('    let currentUser =')), ctx);
  for (const name of ['GENERAL_HELP_OPTIONS', 'GENERAL_FOLLOW_OPTIONS']) {
    vm.runInContext(html.match(new RegExp('    const ' + name + ' = [^;]+;'))[0], ctx);
  }
  vm.runInContext(fs.readFileSync(path.join(root, 'report-templates.js'), 'utf8'), ctx);
  return vm.runInContext('({forms:FORMS, generalSchema:GENERAL_SCHEMA, generalFields:GENERAL_FIELDS, summaryFields:FORM5_FIELDS, helpOptions:GENERAL_HELP_OPTIONS, followOptions:GENERAL_FOLLOW_OPTIONS})', ctx);
}
function fixture() {
  const ctx = vm.createContext({});
  const config = installTemplates(ctx);
  const record = { id: 'TEST-001', name: 'ศูนย์ตาดีกาตัวอย่าง (ข้อมูลทดสอบ)', type: 'ตาดีกา', timestamp: '2026-09-22 09:30', supervisor: 'ผู้นิเทศตัวอย่าง',
    score1: 30, score2: 36, score3: 34, score4: 10, totalScore: 100, pct: 67, level: 'ดี', formType: 'แบบที่ 1-5 (ครบทุกด้าน)',
    details: { answers: {}, notes: {}, form2: { teacher: 'ครูตัวอย่าง', subject: 'คุณธรรมและการอยู่ร่วมกัน', class: 'ชั้นปีที่ 3', duration: '30 นาที', context: 'สังเกตการทำงานกลุ่ม', suggestion: 'เปิดโอกาสให้ผู้เรียนทุกคนได้อธิบายเหตุผล' },
      form3: { subject: 'ศาสนปฏิบัติ', class: 'ชั้นปีที่ 3', period: 'ภาคเรียนที่ 1' }, form4: { sampled: 10, passed: 8, pct: 80 },
      form5: { strengths: 'ครูใช้สื่อประกอบการสอนอย่างเหมาะสม\nผู้เรียนมีส่วนร่วมในการฝึกปฏิบัติ', gaps: 'ควรเพิ่มการติดตามผู้เรียนรายบุคคล', causes: 'ตัวอย่างข้อมูลเก่าที่บันทึกไว้', resources: 'สื่อประกอบการสอน' },
      actionPlan: [{ topic: 'ติดตามผู้เรียนรายบุคคล', activity: 'จัดทำบันทึกพัฒนาการและนัดสะท้อนผลกับครู', owner: 'ครูประจำชั้น', deadline: 'ตุลาคม 2569', evidence: 'บันทึกพัฒนาการรายบุคคล' }] } };
  config.forms.forEach(f => {
    record.details.answers[f.no] = f.sections.flatMap(s => s.items.map((_, i) => i % 4));
    record.details.notes[f.no] = f.sections.flatMap(s => s.items.map(() => 'พบหลักฐานจากแผนการสอนและการสังเกตชั้นเรียน\nควรติดตามผลต่อเนื่อง'));
  });
  const general = structuredClone(record);
  for (let no = 1; no <= 4; no++) general['score' + no] = 0;
  general.details = { generalRecord: { answers: [2, 1, 0, 'N/A', 2, 1, 2, 1, 0, 'N/A', 1, 2], notes: Array(12).fill('สังเกตการปฏิบัติจริงและสนทนากับครู'),
    meta: { date: '2026-09-22', contact: 'ผู้รับผิดชอบตัวอย่าง', levels: '1-6', students: 80, teachers: 8, teacher: 'ครูตัวอย่าง', subject: 'ศาสนปฏิบัติ', class: 'ชั้นปีที่ 3', duration: '30 นาที' },
    summary: { g_strengths: 'ผู้เรียนมีส่วนร่วมและฝึกปฏิบัติจริง', g_gaps: 'ติดตามผู้เรียนที่ต้องการความช่วยเหลือ', helpOptions: ['การพัฒนาครู', 'อื่น ๆ'], g_help: 'อบรมการใช้สื่อดิจิทัล',
      g_agreement: 'พัฒนากิจกรรมฝึกปฏิบัติ', g_action: 'ครูร่วมกันออกแบบกิจกรรม', g_supporter: 'คณะกรรมการศูนย์', g_followDate: 'ตุลาคม 2569', g_followLook: 'ผลงานและการปฏิบัติของผู้เรียน',
      followOptions: ['กำลังทำ', 'อื่น ๆ'], followOther: 'รอรับสื่อเพิ่มเติม', g_change: 'ผู้เรียนกล้าแสดงออกมากขึ้น', g_evidence: 'บันทึกการสังเกตและผลงาน', g_followStatus: 'ติดตามความก้าวหน้าในครั้งต่อไป' } } };
  return { config, record, general, templates: ctx.ReportTemplates, context: { centerId: record.id, centerName: record.name, type: 'ตาดีกา', from: '', to: '', kind: 'form1' } };
}
module.exports = { installTemplates, fixture, html, root };
