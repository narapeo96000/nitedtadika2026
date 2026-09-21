const SHEET_ID = '1qk9eLhwKgPvh2fLwWNSthV4JKyDkJGqJojhLDus5460';
// ชื่อชีตตาม Google Sheets: แยกข้อมูลตามประเภทศูนย์
const SHEET_DATA_TADEKA = 'DATA_TADEKA';
const SHEET_ADDR_TADEKA  = 'ADDR_TADEKA';
const SHEET_DATA_PONDOK = 'DATA_PONDOK';
const SHEET_ADDR_PONDOK = 'ADDR_PONDOK';
const SHEET_USERS = 'USERS';
const TYPE_TADEKA = 'ตาดีกา';
const TYPE_PONDOK = 'ปอเนาะ';
const ADDR_SHEETS = [
  { name: SHEET_ADDR_TADEKA, type: TYPE_TADEKA, startRow: 6 },
  { name: SHEET_ADDR_PONDOK, type: TYPE_PONDOK, startRow: 3 }
];
const DATA_SHEETS = [
  { name: SHEET_DATA_TADEKA, type: TYPE_TADEKA },
  { name: SHEET_DATA_PONDOK, type: TYPE_PONDOK }
];
function getAddrSheet(type) { return type === TYPE_PONDOK ? SHEET_ADDR_PONDOK : SHEET_ADDR_TADEKA; }
function getDataSheet(type) { return type === TYPE_PONDOK ? SHEET_DATA_PONDOK : SHEET_DATA_TADEKA; }

// โครงสร้างใหม่ของชีต DATA (15 คอลัมน์ รองรับการประเมิน 4 ด้าน + สรุป)
const DATA_HEADERS = ['Timestamp', 'ID ศูนย์', 'ชื่อศูนย์', 'ประเภทการประเมิน', 'คะแนนแบบ1', 'คะแนนแบบ2', 'คะแนนแบบ3', 'คะแนนแบบ4', 'รวม/150', 'ร้อยละ', 'ระดับ', 'รายละเอียด', 'ผู้นิเทศ', 'แก้ไขครั้งล่าสุด', 'ผู้แก้ไขล่าสุด'];

function doGet(e) {
  // GET API สำหรับข้อมูลอ่านอย่างเดียว: /exec?action=getTadikaList
  if (e && e.parameter && e.parameter.action) {
    try {
      const action = String(e.parameter.action || '').trim();
      let payload = e.parameter.payload || '';
      if (payload) {
        try { payload = JSON.parse(payload); } catch (_) {}
      } else if (e.parameter.id) {
        payload = e.parameter.id;
      } else if (e.parameter.username) {
        payload = { username: e.parameter.username };
      } else {
        payload = {};
      }
      return jsonResponse(routeAction(action, payload));
    } catch (error) {
      return jsonResponse({ success: false, message: error.message });
    }
  }

  return HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html lang="th"><head><meta charset="UTF-8">' +
    '<meta name="viewport" content="width=device-width, initial-scale=1">' +
    '<title>ระบบนิเทศออนไลน์จังหวัดนราธิวาส</title>' +
    '<style>body{font-family:"Sarabun",sans-serif;display:flex;align-items:center;justify-content:center;min-height:100vh;margin:0;background:#ecfdf5;color:#065f46}.box{text-align:center;background:#fff;padding:48px;border-radius:16px;box-shadow:0 10px 30px rgba(0,0,0,.08)}a{display:inline-block;margin-top:16px;padding:12px 24px;background:#059669;color:#fff;text-decoration:none;border-radius:10px}</style>' +
    '</head><body><div class="box">' +
    '<h1>🕌 ระบบนิเทศออนไลน์จังหวัดนราธิวาส</h1>' +
    '<p>บริการนี้เป็น API ของระบบนิเทศออนไลน์ตาดีกา</p>' +
    '<p style="color:#64748b;font-size:14px">หน้าเว็บหลักถูกเปิดใช้งานผ่านหน้าจอระบบแยกต่างหาก</p>' +
    '</div></body></html>'
  );
}

function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);
    return jsonResponse(routeAction(req.action, req.payload));
  } catch (error) {
    return jsonResponse({ success: false, message: error.message });
  }
}

function jsonResponse(value) {
  return ContentService.createTextOutput(JSON.stringify(value || { success: false, message: 'ไม่พบผลลัพธ์' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// จุดรวม routing เดียวกันสำหรับ GET และ POST เพื่อป้องกันพฤติกรรมของ API ไม่ตรงกัน
function routeAction(action, payload) {
  action = String(action || '').trim();
  if (action === 'login') return loginUser(payload || {});
  if (action === 'register') return registerUser(payload || {});
  if (action === 'getTadikaList') return getTadikaList();
  if (action === 'getStats') return getStats();
  if (action === 'getUsers') return getUsers(payload || {});
  if (action === 'setUserStatus') return setUserStatus(payload || {});
  if (action === 'getTadikaData') return getTadikaData(payload);
  if (action === 'getEvaluations') return getEvaluations(payload);
  if (action === 'saveEvaluation') return saveEvaluation(payload || {});
  if (action === 'savePin') return savePin(payload || {});
  if (action === 'chat') return processChatbot(payload);
  return { success: false, message: 'ไม่รองรับ action: ' + action };
}

const USERS_HEADERS = ['Username','Password','ชื่อ-นามสกุล','เบอร์โทร','สถานะ','บทบาท'];

// --- เตรียมชีต USERS ให้มีคอลัมน์ สถานะ/บทบาท + สร้างบัญชี admin ถ้ายังไม่มี ---
function ensureUsersSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_USERS);
  if(!sheet) {
    sheet = ss.insertSheet(SHEET_USERS);
    sheet.getRange(1, 1, 1, USERS_HEADERS.length).setValues([USERS_HEADERS]);
    sheet.appendRow(['admin','admin123','ผู้ดูแลระบบ','-','ใช้งาน','ผู้ดูแลระบบ']);
    return sheet;
  }
  const head = sheet.getRange(1, 1, 1, sheet.getLastColumn() || 1).getValues()[0];
  if(String(head[0]).trim() !== 'Username' || String(head[4] || '').trim() !== 'สถานะ') {
    sheet.getRange(1, 1, 1, USERS_HEADERS.length).setValues([USERS_HEADERS]);
  }
  const lastRow = sheet.getLastRow();
  let hasAdmin = false;
  if(lastRow >= 1) {
    const rows = sheet.getRange(1, 1, lastRow, 6).getValues();
    for(let i = 1; i < rows.length; i++) {
      const u = String(rows[i][0]).trim();
      if(u === '') continue;
      if(u.toLowerCase() === 'admin') hasAdmin = true;
      if(String(rows[i][4] || '').trim() === '') sheet.getRange(i + 1, 5).setValue('ใช้งาน');
      if(String(rows[i][5] || '').trim() === '') {
        sheet.getRange(i + 1, 6).setValue(u.toLowerCase() === 'admin' ? 'ผู้ดูแลระบบ' : 'ผู้ใช้');
      }
    }
  }
  if(!hasAdmin) sheet.appendRow(['admin','admin123','ผู้ดูแลระบบ','-','ใช้งาน','ผู้ดูแลระบบ']);
  return sheet;
}

// --- ตรวจสอบ Login (ต้องมีสถานะ = ใช้งาน จึงจะเข้าได้) ---
function loginUser(data) {
  const sheet = ensureUsersSheet();
  const rows = sheet.getDataRange().getValues();
  const inputUser = String(data.username).trim();
  const inputPass = String(data.password).trim();

  for(let i = 1; i < rows.length; i++) {
    let sheetUser = String(rows[i][0]).trim();
    let sheetPass = String(rows[i][1]).trim();

    if(sheetUser === inputUser && sheetPass === inputPass) {
      const status = String(rows[i][4] || '').trim() || 'ใช้งาน';
      const role = String(rows[i][5] || '').trim() || (sheetUser.toLowerCase() === 'admin' ? 'ผู้ดูแลระบบ' : 'ผู้ใช้');
      if(status === 'รออนุมัติ') {
        return { success: false, message: 'บัญชีของคุณยังรอการอนุมัติจากผู้ดูแลระบบ กรุณารอผู้ดูแลระบบอนุมัติก่อนเข้าสู่ระบบ' };
      }
      if(status === 'ระงับ') {
        return { success: false, message: 'บัญชีของคุณถูกระงับการใช้งาน กรุณาติดต่อผู้ดูแลระบบ' };
      }
      return {
        success: true,
        userData: { username: sheetUser, fname: rows[i][2], tel: rows[i][3], role: role, status: status }
      };
    }
  }
  return {
    success: false,
    message: "Username หรือ Password ไม่ถูกต้อง"
  };
}

// --- สมัครสมาชิกใหม่ (บันทึกสถานะ = รออนุมัติ รอผู้ดูแลระบบ) ---
function registerUser(data) {
  const sheet = ensureUsersSheet();
  const rows = sheet.getDataRange().getValues();
  const u = String(data.username).trim();
  const p = String(data.password).trim();
  const f = String(data.fname).trim();
  const tel = String(data.tel || '').trim();

  if(!u || !p || !f) return {success: false, message: "กรอกข้อมูลไม่ครบถ้วน (ต้องมี Username, Password และชื่อ-นามสกุล)"};

  for(let i = 1; i < rows.length; i++) {
    if(String(rows[i][0]).trim().toLowerCase() === u.toLowerCase()) {
      return {success: false, message: "Username นี้ถูกใช้งานแล้ว กรุณาใช้ชื่ออื่น"};
    }
  }

  sheet.appendRow([u, p, f, tel, 'รออนุมัติ', 'ผู้ใช้']);
  return {success: true, message: "สมัครสมาชิกเรียบร้อย!<br>บัญชีของคุณ<b>รอการอนุมัติจากผู้ดูแลระบบ</b> จึงจะเข้าสู่ระบบได้"};
}

// --- ผู้ดูแลระบบ: ดึงรายชื่อผู้ใช้ทั้งหมด ---
function getUsers(data) {
  const sheet = ensureUsersSheet();
  const rows = sheet.getDataRange().getValues();
  const admin = String(data.username || '').trim();
  let isAdmin = false;
  for(let i = 1; i < rows.length; i++) {
    if(String(rows[i][0]).trim() === admin && String(rows[i][5] || '').trim() === 'ผู้ดูแลระบบ') { isAdmin = true; break; }
  }
  if(!isAdmin) return {success: false, message: 'ไม่มีสิทธิ์ใช้งาน (เฉพาะผู้ดูแลระบบ)'};
  const list = [];
  for(let i = 1; i < rows.length; i++) {
    if(String(rows[i][0]).trim() === '') continue;
    list.push({
      row: i + 1,
      username: rows[i][0],
      fname: rows[i][2],
      tel: rows[i][3],
      status: String(rows[i][4] || '').trim() || 'ใช้งาน',
      role: String(rows[i][5] || '').trim() || 'ผู้ใช้'
    });
  }
  return {success: true, data: list};
}

// --- ผู้ดูแลระบบ: อนุมัติ / ระงับ / เปิดใช้งานบัญชี ---
function setUserStatus(data) {
  const sheet = ensureUsersSheet();
  const rows = sheet.getDataRange().getValues();
  const admin = String(data.admin || '').trim();
  const target = String(data.username || '').trim();
  const status = String(data.status || '').trim();

  let isAdmin = false, targetRow = -1;
  for(let i = 1; i < rows.length; i++) {
    const u = String(rows[i][0]).trim();
    if(u === admin && String(rows[i][5] || '').trim() === 'ผู้ดูแลระบบ') isAdmin = true;
    if(u === target) targetRow = i + 1;
  }
  if(!isAdmin) return {success: false, message: 'ไม่มีสิทธิ์ใช้งาน (เฉพาะผู้ดูแลระบบ)'};
  if(targetRow < 1) return {success: false, message: 'ไม่พบบัญชีผู้ใช้นี้'};
  if(status === 'ใช้งาน' || status === 'ระงับ') {
    sheet.getRange(targetRow, 5).setValue(status);
    return {success: true, message: (status === 'ใช้งาน' ? '✅ เปิดใช้งาน' : '⛔ ระงับ') + 'บัญชี "' + target + '" เรียบร้อย'};
  }
  return {success: false, message: 'สถานะไม่ถูกต้อง'};
}

// --- ดึงรายชื่อศูนย์มาให้เลือก โดยแต่ละชีตมีแถวเริ่มต้นต่างกัน ---
function getTadikaList() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let list = [];

  ADDR_SHEETS.forEach(cfg => {
    const sheet = ss.getSheetByName(cfg.name);
    if(!sheet) return;
    const startRow = cfg.startRow || 6;
    const lastRow = sheet.getLastRow();
    if(lastRow >= startRow) {
      // คอลัมน์: A=ID, B=สถานะ, C=มัสยิด, D=ชื่อศูนย์, E=ประธาน, J=ที่อยู่, K=ตำบล, L=อำเภอ,
      // M=โทร, N=ขนาด, O/P/Q=ครู ช/ญ/รวม, R/S/T=นักเรียน ช/ญ/รวม, U=ละติจูด, V=ลองจิจูด
      const rows = sheet.getRange(startRow, 1, lastRow - startRow + 1, 22).getValues();
      for(let i = 0; i < rows.length; i++) {
        if(rows[i][0] != "") {
          list.push({
            id: rows[i][0],
            type: cfg.type,
            status: rows[i][1],
            name: rows[i][3],
            mosque: rows[i][2],
            head: rows[i][4],
            phone: rows[i][12],
            size: rows[i][13],
            teachers: { male: rows[i][14], female: rows[i][15], total: rows[i][16] },
            students: { male: rows[i][17], female: rows[i][18], total: rows[i][19] },
            address: rows[i][9],
            subdist: rows[i][10],
            dist: rows[i][11],
            lat: rows[i][20],
            lng: rows[i][21]
          });
        }
      }
    }
  });
  return {success: true, data: list};
}

// --- สถิติระบบนิเทศออนไลน์ (นับเฉพาะตาดีกาที่ได้รับการนิเทศ จาก ADDR_TADEKA + DATA_TADEKA) ---
function getStats() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let totalTadika = 0, totalEval = 0, totalUsers = 0, sumPct = 0;
  let registered = 0, unregistered = 0;
  let tMale = 0, tFemale = 0, tTotal = 0, sMale = 0, sFemale = 0, sTotal = 0;
  const levelCounts = { 'ดีมาก': 0, 'ดี': 0, 'พอใช้': 0, 'ต้องปรับปรุง': 0, 'ไม่ระบุ': 0 };
  const latest = [];
  const evaluatedIds = {};

  // 1) อ่าน DATA_TADEKA ก่อน: รวบรวม ID ศูนย์ที่ได้รับการนิเทศ + สถิติการนิเทศ
  const data = ss.getSheetByName(SHEET_DATA_TADEKA);
  if(data) {
    const lastRow = data.getLastRow();
    if(lastRow >= 2) {
      // หาตำแหน่งคอลัมน์จาก header เพื่อกันข้อมูลเลื่อนคอลัมน์
      const header = data.getRange(1, 1, 1, 15).getValues()[0];
      const colOf = name => {
        for(let i = 0; i < header.length; i++) if(String(header[i]).trim() === name) return i;
        return -1;
      };
      const ciId = colOf('ID ศูนย์'), ciName = colOf('ชื่อศูนย์'), ciPct = colOf('ร้อยละ'), ciLvl = colOf('ระดับ'), ciTs = colOf('Timestamp');
      const vals = data.getRange(2, 1, lastRow - 1, 15).getValues();
      totalEval = vals.length;
      vals.forEach(r => {
        const id = String(r[ciId] || '').trim();
        if(id) evaluatedIds[id] = true;
        const pct = Number(r[ciPct]);
        let lvl = String(r[ciLvl] || '').trim();
        if(!lvl || !(lvl in levelCounts)) lvl = 'ไม่ระบุ';
        levelCounts[lvl]++;
        if(!isNaN(pct)) sumPct += pct;
        latest.push({ name: String(r[ciName] || ''), timestamp: formatDate(r[ciTs]), pct: isNaN(pct) ? 0 : pct, level: lvl });
      });
    }
  }
  totalTadika = Object.keys(evaluatedIds).length;
  latest.sort((a, b) => a.timestamp < b.timestamp ? 1 : -1);

  // 2) นับข้อมูลศูนย์จาก ADDR_TADEKA เฉพาะศูนย์ที่ได้รับการนิเทศ
  // ADDRESS: A=ID, B=สถานะ, J=เลขที่/หมู่/ถนน, O/P/Q=ครู ช/ญ/รวม, R/S/T=นักเรียน ช/ญ/รวม
  const toNum = v => { const n = Number(String(v).replace(/,/g, '').trim()); return isNaN(n) ? 0 : n; };
  const addr = ss.getSheetByName(SHEET_ADDR_TADEKA);
  if(addr) {
    const lastRow = addr.getLastRow();
    if(lastRow >= 6) {
      const vals = addr.getRange(6, 1, lastRow - 5, 22).getValues();
      vals.forEach(r => {
        const id = String(r[0]).trim();
        if(id === '' || !evaluatedIds[id]) return;
        const status = String(r[1] || '').trim();
        if(status) {
          if(status.includes('จดทะเบียน') && !status.includes('ไม่')) registered++;
          else unregistered++;
        }
        const tm = toNum(r[14]), tf = toNum(r[15]), tt = toNum(r[16]);
        tMale += tm; tFemale += tf; tTotal += tt || (tm + tf);
        const sm = toNum(r[17]), sf = toNum(r[18]), st = toNum(r[19]);
        sMale += sm; sFemale += sf; sTotal += st || (sm + sf);
      });
    }
  }

  const users = ss.getSheetByName(SHEET_USERS);
  if(users) {
    const lastRow = users.getLastRow();
    if(lastRow >= 2) totalUsers = lastRow - 1;
  }

  // 3) นับศูนย์ตาดีกาทั้งหมดจาก ADDR_TADEKA (สำหรับการ์ด ข้อมูลตาดีกา (TADEKA) — ไม่รวมปอเนาะ)
  let tkTotal = 0, tkReg = 0, tkUnreg = 0;
  let tkTMale = 0, tkTFemale = 0, tkTTotal = 0, tkSMale = 0, tkSFemale = 0, tkSTotal = 0;
  if(addr) {
    const lastRow = addr.getLastRow();
    if(lastRow >= 6) {
      const vals = addr.getRange(6, 1, lastRow - 5, 22).getValues();
      vals.forEach(r => {
        if(String(r[0]).trim() === '') return;
        tkTotal++;
        const status = String(r[1] || '').trim();
        if(status) {
          if(status.includes('จดทะเบียน') && !status.includes('ไม่')) tkReg++;
          else tkUnreg++;
        }
        const tm = toNum(r[14]), tf = toNum(r[15]), tt = toNum(r[16]);
        tkTMale += tm; tkTFemale += tf; tkTTotal += tt || (tm + tf);
        const sm = toNum(r[17]), sf = toNum(r[18]), st = toNum(r[19]);
        tkSMale += sm; tkSFemale += sf; tkSTotal += st || (sm + sf);
      });
    }
  }

  return {
    success: true,
    data: {
      totalTadika: totalTadika,
      registered: registered,
      unregistered: unregistered,
      teachers: { male: tMale, female: tFemale, total: tTotal },
      students: { male: sMale, female: sFemale, total: sTotal },
      totalEval: totalEval,
      totalUsers: totalUsers,
      avgPct: totalEval ? Math.round(sumPct / totalEval) : 0,
      levelCounts: levelCounts,
      latest: latest.slice(0, 10),
      tkTotal: tkTotal,
      tkReg: tkReg,
      tkUnreg: tkUnreg,
      tkTeachers: { male: tkTMale, female: tkTFemale, total: tkTTotal },
      tkStudents: { male: tkSMale, female: tkSFemale, total: tkSTotal }
    }
  };
}

// --- ดึงข้อมูลรายละเอียดของศูนย์ที่เลือก (ค้นหาจากทั้ง 2 ชีต ADDR) ---
function getTadikaData(id) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  id = String(id).trim();
  for(const cfg of ADDR_SHEETS) {
    const sheet = ss.getSheetByName(cfg.name);
    if(!sheet) continue;
    const startRow = cfg.startRow || 6;
    const lastRow = sheet.getLastRow();
    if(lastRow < startRow) continue;
    const rows = sheet.getRange(startRow, 1, lastRow - startRow + 1, 23).getValues();
    for(let i = 0; i < rows.length; i++) {
      if(String(rows[i][0]).trim() === id) {
        return {
          success: true,
          data: {
            row: i + startRow,
            type: cfg.type,
            id: rows[i][0], status: rows[i][1], mosque: rows[i][2], name: rows[i][3],
            head: rows[i][4], eduSec: rows[i][5], eduRel: rows[i][6], foundedDate: rows[i][7],
            regNum: rows[i][8], address: rows[i][9], subdist: rows[i][10], dist: rows[i][11],
            phone: rows[i][12], size: rows[i][13], tMale: rows[i][14], tFemale: rows[i][15],
            tTotal: rows[i][16], sMale: rows[i][17], sFemale: rows[i][18], sTotal: rows[i][19],
            lat: rows[i][20], lng: rows[i][21], rooms: rows[i][22]
          }
        };
      }
    }
  }
  return {success: false, message: "ไม่พบข้อมูลศูนย์"};
}

// --- บันทึกพิกัด GPS ของศูนย์ (ชีต ADDR ที่ตรงตามประเภท: คอลัมน์ U=ละติจูด, V=ลองจิจูด) ---
function savePin(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const id = String(data && data.id ? data.id : '').trim();
  const lat = String(data && data.lat != null ? data.lat : '').trim();
  const lng = String(data && data.lng != null ? data.lng : '').trim();
  if(!id) return {success: false, message: "ไม่พบรหัสศูนย์ (ID)"};
  if(!lat || !lng || isNaN(Number(lat)) || isNaN(Number(lng))) {
    return {success: false, message: "พิกัดไม่ถูกต้อง (ต้องเป็นตัวเลขละติจูด/ลองจิจูด)"};
  }
  for(const cfg of ADDR_SHEETS) {
    const sheet = ss.getSheetByName(cfg.name);
    if(!sheet) continue;
    const startRow = cfg.startRow || 6;
    const lastRow = sheet.getLastRow();
    if(lastRow < startRow) continue;
    const rows = sheet.getRange(startRow, 1, lastRow - startRow + 1, 22).getValues();
    for(let i = 0; i < rows.length; i++) {
      if(String(rows[i][0]).trim() === id) {
        const row = i + startRow;
        sheet.getRange(row, 21).setValue(Number(lat));
        sheet.getRange(row, 22).setValue(Number(lng));
        return {success: true, message: "บันทึกพิกัดเรียบร้อย (U=ละติจูด, V=ลองจิจูด)", row: row};
      }
    }
  }
  return {success: false, message: "ไม่พบข้อมูลศูนย์นี้ในชีต ADDRESS"};
}

// --- เตรียมชีต DATA (ตาดีกา/ปอเนาะ) ให้มี Header ครบตาม DATA_HEADERS (15 คอลัมน์) ---
function ensureDataSheet(type) {
  const name = getDataSheet(type);
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(name);
  if(!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, DATA_HEADERS.length).setValues([DATA_HEADERS]);
    return sheet;
  }
  sheet.getRange(1, 1, 1, DATA_HEADERS.length).setValues([DATA_HEADERS]);
  return sheet;
}

function formatDate(d) {
  if(!d) return '';
  if(!(d instanceof Date)) d = new Date(d);
  const pad = n => ('0' + n).slice(-2);
  return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()) + ' ' +
         pad(d.getHours()) + ':' + pad(d.getMinutes()) + ':' + pad(d.getSeconds());
}

// --- ดึงประวัติการนิเทศของศูนย์ที่เลือก (ค้นหาจากทั้ง 2 ชีต DATA) ---
function getEvaluations(tadikaId) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  tadikaId = String(tadikaId).trim();
  const list = [];

  for(const cfg of DATA_SHEETS) {
    const sheet = ss.getSheetByName(cfg.name);
    if(!sheet) continue;
    const lastRow = sheet.getLastRow();
    if(lastRow < 2) continue;
    const rows = sheet.getRange(2, 1, lastRow - 1, DATA_HEADERS.length).getValues();
    for(let i = 0; i < rows.length; i++) {
      if(String(rows[i][1]).trim() === tadikaId) {
        let details = null;
        try { details = JSON.parse(rows[i][11] || 'null'); } catch(e) { details = null; }
        list.push({
          row: i + 2,
          type: cfg.type,
          timestamp: formatDate(rows[i][0]),
          id: rows[i][1],
          name: rows[i][2],
          formType: rows[i][3],
          score1: rows[i][4],
          score2: rows[i][5],
          score3: rows[i][6],
          score4: rows[i][7],
          totalScore: rows[i][8],
          pct: rows[i][9],
          level: rows[i][10],
          details: details,
          supervisor: rows[i][12],
          lastEdit: formatDate(rows[i][13]),
          lastEditor: rows[i][14]
        });
      }
    }
  }
  list.sort((a, b) => a.timestamp < b.timestamp ? 1 : -1);
  return {success: true, data: list};
}

// --- บันทึก/แก้ไขผลนิเทศ + อัปเดตข้อมูลศูนย์ (เลือกชีตตามประเภทตาดีกา/ปอเนาะ) ---
function saveEvaluation(payload) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const t = payload.tadikaData;
  const e = payload.evalData;
  // ประเภทศูนย์: มาจากข้อมูลศูนย์ (เพิ่ม), หรือจากประวัติที่กำลังแก้ไข (แก้ไข)
  const type = (t && t.type) || (payload.editType && payload.editType.trim()) || TYPE_TADEKA;
  const dataType = (payload.editRow && !isNaN(payload.editRow) && payload.editType && payload.editType.trim())
    ? payload.editType.trim() : type;

  // อัปเดตข้อมูลศูนย์ลงชีต ADDR ที่ตรงตามประเภท
  const addressSheet = ss.getSheetByName(getAddrSheet(type));
  const updateValues = [[
    t.status, t.mosque, t.name, t.head, t.eduSec, t.eduRel, t.foundedDate, t.regNum,
    t.address, t.subdist, t.dist, t.phone, t.size, t.tMale, t.tFemale, t.tTotal,
    t.sMale, t.sFemale, t.sTotal, t.lat, t.lng, t.rooms
  ]];
  if(addressSheet && t.row) {
    addressSheet.getRange(t.row, 2, 1, 22).setValues(updateValues);
  }

  // บันทึก/แก้ไขผลนิเทศลงชีต DATA ที่ตรงตามประเภท
  const dataSheet = ensureDataSheet(dataType);
  const supervisor = String(payload.supervisor);
  const now = new Date();
  const detailsJSON = JSON.stringify({
    answers: e.answers || {},
    notes: e.notes || {},
    generalRecord: e.generalRecord || {},
    form1: e.form1 || {},
    form2: e.form2 || {},
    form3: e.form3 || {},
    form4: e.form4 || {},
    form5: e.form5 || {},
    actionPlan: e.actionPlan || [],
    comment: e.comment || '',
    strengths: e.strengths || '',
    improve: e.improve || ''
  });

  if(payload.editRow && !isNaN(payload.editRow)) {
    const row = Number(payload.editRow);
    dataSheet.getRange(row, 4).setValue(e.formType);
    dataSheet.getRange(row, 5).setValue(e.score1);
    dataSheet.getRange(row, 6).setValue(e.score2);
    dataSheet.getRange(row, 7).setValue(e.score3);
    dataSheet.getRange(row, 8).setValue(e.score4);
    dataSheet.getRange(row, 9).setValue(e.totalScore);
    dataSheet.getRange(row, 10).setValue(e.pct);
    dataSheet.getRange(row, 11).setValue(e.level);
    dataSheet.getRange(row, 12).setValue(detailsJSON);
    // ยืนยันคอลัมน์ M:N:O ให้ตรงกับหัวตารางทุกครั้งที่แก้ไขรายการเดิม
    dataSheet.getRange(row, 13, 1, 3).setValues([[supervisor, now, supervisor]]);
    return {
      success: true,
      message: 'แก้ไขผลการนิเทศเรียบร้อยแล้ว!<br>แก้ไขครั้งล่าสุด: ' + formatDate(now) + ' โดย ' + supervisor,
      lastEdit: formatDate(now),
      lastEditor: supervisor
    };
  }

  dataSheet.appendRow([now, t.id, t.name, e.formType, e.score1, e.score2, e.score3, e.score4, e.totalScore, e.pct, e.level, detailsJSON, supervisor, now, supervisor]);
  return {success: true, message: 'อัปเดตข้อมูลศูนย์ และบันทึกผลการนิเทศเรียบร้อยแล้ว!', lastEdit: formatDate(now), lastEditor: supervisor};
}

// ============================================================
// แชทบอท "น้องศึกษา" — ใช้ Gemini API
// API Key เก็บที่: Apps Script > Project Settings > Script Properties
//   ชื่อคีย์: geminiKey หรือ GEMINI_API_KEY (รองรับทั้งสอง)
// (ไม่ฝัง key ในโค้ด เพื่อกันรั่วไหลบน GitHub สาธารณะ)
// ============================================================

// --- อ่านการตั้งค่าระบบจาก Script Properties (รองรับทั้ง geminiKey และ GEMINI_API_KEY) ---
function getGeminiApiKey() {
  const k = PropertiesService.getScriptProperties().getProperty('geminiKey') || PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  return k ? String(k).trim() : '';
}

function getSystemSettings() {
  const p = PropertiesService.getScriptProperties();
  const geminiKey = p.getProperty('geminiKey') || p.getProperty('GEMINI_API_KEY') || '';
  return {
    geminiKey: String(geminiKey).trim(),
    models: ['gemini-3.5-flash', 'gemini-3-flash', 'gemini-3.1-flash-lite', 'gemini-1.5-flash']
  };
}

function processChatbot(userMessage) {
  const msg = String(userMessage || '').toLowerCase();

  // อ่านการตั้งค่าระบบ (รวม Gemini API Key)
  const settings = getSystemSettings();

  // ถ้ายังไม่ได้ตั้ง API Key ให้ตอบแบบออฟไลน์ (ไม่ต้องใช้ Gemini)
  if (!settings.geminiKey || settings.geminiKey.trim() === '') {
    return { reply: offlineChatReply(msg) };
  }

  // ดึงข้อมูลสถิติ + รายละเอียดศูนย์จาก Sheets มาเป็น Context ให้ AI ตอบเสมอ
  const contextData = fetchStatisticsForAI();

  // ส่งคำถามผู้ใช้ พร้อมข้อมูลบริบทไปให้ Gemini คิดคำตอบ
  // (ส่งกลับเป็น Markdown ธรรมดา — หน้าเว็บจะแปลงเป็น HTML ปลอดภัยเอง)
  return { reply: callGeminiAPI(userMessage, contextData, settings) };
}

// คำตอบสำรองเมื่อยังไม่ได้ตั้งค่า API Key หรือ API ขัดข้อง
function offlineChatReply(msg) {
  if (msg.includes('สวัสดี') || msg.includes('hello') || msg.includes('hi') || msg.includes('เป็นใคร') || msg.includes('อะไร') || msg.includes('ชื่ออะไร') || msg.length <= 10) {
    return "อัสลามุอะลัยกุม! ขอสันติจงมีแด่ท่าน ฉันคือ \"น้องศึกษา\" ผู้ช่วย AI ยินดีให้คำปรึกษาเกี่ยวกับชุดเครื่องมือนิเทศออนไลน์สำหรับศูนย์ตาดีกา จังหวัดนราธิวาสครับ<br>สอบถามเรื่องเกณฑ์การให้คะแนน การจัดการเรียนรู้ หรือข้อมูลศูนย์ได้เลยครับ";
  }
  if (msg.includes('เกณฑ์') || msg.includes('คะแนน')) {
    return "แบบบันทึกสำหรับผู้นิเทศทั่วไปมี 2 ด้าน รวม 12 ข้อครับ: ด้านที่ 1 การนำหลักสูตรไปใช้ 6 ข้อ และด้านที่ 2 การจัดการเรียนรู้ของครู 6 ข้อ ใช้เกณฑ์ 2 = ทำได้ชัดเจน, 1 = กำลังพัฒนา, 0 = ต้องได้รับการช่วยเหลือ และ N/A = พิจารณาไม่ได้ในครั้งนี้ โดยใช้เพื่อค้นหาจุดแข็งและวางแผนช่วยเหลือ ไม่ใช้จัดอันดับศูนย์หรือบุคลากรครับ";
  }
  if (msg.includes('ระดับ')) {
    return "ระบบแสดงภาพรวมเพื่อช่วยกำหนดการสนับสนุน ไม่ใช่การจัดอันดับครับ โดยพิจารณาจากสิ่งที่พบจริง หลักฐาน การสนทนา และการสังเกตชั้นเรียนร่วมกัน";
  }
  if (msg.includes('แบบที่ 4') || msg.includes('สุ่ม')) {
    return "แบบบันทึกนี้ให้สังเกตผู้เรียนว่ามีส่วนร่วม คิด สื่อสาร และฝึกปฏิบัติจริงหรือไม่ หากพิจารณาไม่ได้ในครั้งนี้ให้เลือก N/A และบันทึกเหตุผลหรือหลักฐานไว้ครับ";
  }
  if (msg.includes('แบบที่ 5') || msg.includes('แผนพัฒนา')) {
    return "ส่วนสรุปให้บันทึกจุดแข็งไม่เกิน 3 เรื่อง ประเด็นพัฒนา 1–2 เรื่อง ความช่วยเหลือที่ต้องการ ข้อตกลง สิ่งที่จะดำเนินการ ผู้สนับสนุน และสิ่งที่จะติดตามครั้งต่อไปครับ";
  }
  if (msg.includes('กี่') || msg.includes('จำนวน') || msg.includes('สถิติ') || msg.includes('ทั้งหมด')) {
    return "ขณะนี้ยังไม่ได้เชื่อมต่อ AI สำหรับข้อมูลสถิติ กรุณาตั้งค่า Gemini API Key ก่อนครับ (หรือถามเรื่องเกณฑ์การประเมินได้ทันที)";
  }
  return "สวัสดีครับ ผมน้องศึกษา พร้อมให้คำแนะนำเรื่องเกณฑ์การประเมินตาดีกา 5 แบบฟอร์มได้เลยครับ";
}

// ---------------------------------------------------------
// ฟังก์ชันสำหรับคำนวณและสรุปสถิติจาก Sheet ให้เป็นข้อความ Text
//
// โครงสร้างชีต ADDR_TADEKA (เริ่มข้อมูลจริงที่แถว 6) และ ADDR_PONDOK (เริ่มแถว 3):
//   A=ID, B=สถานะ, C=ชื่อมัสยิด, D=ชื่อศูนย์, E=ประธานศูนย์,
//   F=วุฒิ(สามัญ), G=วุฒิ(ศาสนา), H=วันก่อตั้ง, I=เลขจดทะเบียน,
//   J=เลขที่/หมู่/ถนน, K=ตำบล, L=อำเภอ, M=โทรศัพท์, N=ขนาดศูนย์,
//   O=ผู้สอน(ชาย), P=ผู้สอน(หญิง), Q=ผู้สอน(รวม), R=ผู้เรียน(ชาย),
//   S=ผู้เรียน(หญิง), T=ผู้เรียน(รวม)
// ---------------------------------------------------------
function fetchStatisticsForAI() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  // --- ดึงข้อมูลตาดีกา ---
  const sheetTadika = ss.getSheetByName(SHEET_ADDR_TADEKA);
  const tadikaStats = { count: 0, tMale: 0, tFemale: 0, sMale: 0, sFemale: 0 };
  const tadikaLines = [];

  if (sheetTadika) {
    const lastRow = sheetTadika.getLastRow();
    if (lastRow >= 6) {
      const dataT = sheetTadika.getRange(6, 1, lastRow - 5, 20).getValues();
      for (let i = 0; i < dataT.length; i++) {
        if (String(dataT[i][0]).trim() != "") { // นับเฉพาะแถวที่มี ID
          tadikaStats.count++;
          tadikaStats.tMale   += Number(dataT[i][14]) || 0;
          tadikaStats.tFemale += Number(dataT[i][15]) || 0;
          tadikaStats.sMale   += Number(dataT[i][17]) || 0;
          tadikaStats.sFemale += Number(dataT[i][18]) || 0;
          tadikaLines.push(
            "• " + (dataT[i][3] || '-') +
            " | ประธาน: " + (dataT[i][4] || '-') +
            " | ที่ตั้ง: " + [dataT[i][10], dataT[i][11]].filter(Boolean).join(' ') +
            " | โทร: " + (dataT[i][12] || '-') +
            " | ครู " + (dataT[i][16] || 0) + " คน | นร. " + (dataT[i][19] || 0) + " คน"
          );
        }
      }
    }
  }

  // สร้างข้อความ Context ส่งให้ AI รับรู้
  let contextText = "" +
    "[ข้อมูลจริงจากฐานข้อมูลระบบนิเทศออนไลน์ตาดีกา — ใช้ตัวเลขเหล่านี้ตอบ ไม่ควรแต่งตัวเลข]\n\n" +
    "1. ศูนย์ตาดีกา: ทั้งหมด " + tadikaStats.count + " แห่ง\n" +
    "   - ครูผู้สอนรวม " + (tadikaStats.tMale + tadikaStats.tFemale) + " คน (ชาย " + tadikaStats.tMale + " / หญิง " + tadikaStats.tFemale + ")\n" +
    "   - ผู้เรียนรวม " + (tadikaStats.sMale + tadikaStats.sFemale) + " คน (ชาย " + tadikaStats.sMale + " / หญิง " + tadikaStats.sFemale + ")\n\n" +
    "รายชื่อศูนย์ตาดีกา (ข้อมูลสำหรับตอบคำถามรายศูนย์):\n" + tadikaLines.join('\n');

  return contextText;
}

// ---------------------------------------------------------
// ฟังก์ชันเรียก Gemini API (ฝัง Context ไปด้วย)
// ---------------------------------------------------------
function callGeminiAPI(userMessage, contextData, settings) {
  const API_KEY = (settings && settings.geminiKey) ? settings.geminiKey : getGeminiApiKey();
  // ลำดับโมเดลที่ต้องการใช้ (ตัวแรกดีที่สุด ถ้าล้มเหลวจะถอยไปตัวถัดไปอัตโนมัติ)
  const MODELS = (settings && settings.models && settings.models.length) ? settings.models : ['gemini-3.5-flash', 'gemini-3-flash', 'gemini-3.1-flash-lite', 'gemini-1.5-flash'];

  const systemPrompt = "คุณคือ \"น้องศึกษา\" ผู้ช่วยอัจฉริยะ AI บุคลิกสุภาพ เป็นมิตร กระตือรือร้น ให้เกียรติผู้ใช้งาน ใช้ภาษาไทยที่ถูกต้อง เป็นทางการแต่นุ่มนวล และลงท้ายประโยคด้วย \"ครับ/ค่ะ\" เสมอ ตอบให้กระชับ ตรงประเด็น ใช้ Bullet points จัดรูปแบบให้อ่านง่าย กฎสำคัญ: เมื่อใดก็ตามที่ผู้ใช้ทักทาย (เช่น \"สวัสดี\", \"hello\", \"hi\") หรือเป็นการสนทนาเริ่มต้น ให้เริ่มคำตอบด้วย \"อัสลามุอะลัยกุม! ขอสันติจงมีแด่ท่าน\" เสมอ ตามด้วยแนะนำตัว \"ฉันคือ \"น้องศึกษา\" ผู้ช่วย AI ยินดีให้คำปรึกษาเกี่ยวกับชุดเครื่องมือนิเทศออนไลน์สำหรับศูนย์ตาดีกา จังหวัดนราธิวาส\" และเสนอความช่วยเหลือ ตัวอย่างคำตอบแรกสุดของบทสนทนา คือ \"อัสลามุอะลัยกุม! ขอสันติจงมีแด่ท่าน ฉันคือ \"น้องศึกษา\" ผู้ช่วย AI ยินดีให้คำปรึกษาเกี่ยวกับชุดเครื่องมือนิเทศออนไลน์สำหรับศูนย์ตาดีกา จังหวัดนราธิวาส ฉันพร้อมตอบคำถาม เกณฑ์การให้คะแนน การจัดการเรียนรู้ หรือมีอะไรให้ช่วย สอบถามได้เลย ครับ/ค่ะ 😊\"\n\n" +
    "หน้าที่หลักของคุณคือ:\n" +
    "ให้ข้อมูล คำแนะนำ และตอบคำถามที่เกี่ยวข้องกับการศึกษาเอกชนในจังหวัดนราธิวาส โดยอ้างอิงจากฐานข้อมูลและระเบียบปฏิบัติที่ถูกต้อง ดังนี้:\n\n" +
    "1. ข้อมูลสถิติสถานศึกษา (ต้องให้ข้อมูลตามที่ผู้ใช้ถาม ทั้งภาพรวมและรายศูนย์):\n" +
    "- ข้อมูลจำนวนครูผู้สอน (แยกชาย/หญิง/รวม)\n" +
    "- ข้อมูลจำนวนนักเรียน (แยกชาย/หญิง/รวม)\n" +
    "- ข้อมูลพื้นฐาน: ชื่อศูนย์, ประธานศูนย์, ที่ตั้ง, เบอร์โทรศัพท์, ขนาดของศูนย์\n\n" +
    "2. ความรู้เกี่ยวกับประเภทสถานศึกษาเอกชน:\n" +
    "- โรงเรียนเอกชนในระบบ (สามัญศึกษา, อาชีวศึกษา)\n" +
    "- โรงเรียนเอกชนนอกระบบ (กวดวิชา, ศิลปะและกีฬา, วิชาชีพ, สร้างเสริมทักษะชีวิต)\n" +
    "- โรงเรียนเอกชนสอนศาสนาอิสลาม (ประเภท 15(1), 15(2))\n" +
    "- โรงเรียนเอกชนที่สอนควบคู่ (สามัญ-ศาสนา)\n" +
    "- สถาบันศึกษาปอเนาะ (ระดับการเปิดสอน: อิบติดาอียะฮฺ, มุตะวัซซิเฎาะฮฺ, อาลียะฮฺ; เน้นการจัดการเรียนรู้ 8 สาระการเรียนรู้พื้นฐาน; การรายงานผล ปอ.1, ปอ.2, ปอ.3 และการมีส่วนร่วมของครอบครัว/ชุมชน)\n" +
    "- ศูนย์การศึกษาอิสลามประจำมัสยิด (ตาดีกา)\n\n" +
    "3. กฎหมาย ระเบียบ และหลักสูตร:\n" +
    "- พ.ร.บ. โรงเรียนเอกชน พ.ศ. 2550 (และที่แก้ไขเพิ่มเติม)\n" +
    "- ระเบียบที่เกี่ยวข้องกับปอเนาะ, ตาดีกา และมัสยิด\n" +
    "- หลักสูตรอิสลามศึกษาที่ใช้สอนในตาดีกา (เช่น หลักสูตรฟัรฎูอีนประจำมัสยิด)\n" +
    "- หลักสูตรและการจัดการเรียนรู้ในปอเนาะ (8 สาระการเรียนรู้พื้นฐาน)\n\n" +
    "4. ระบบนิเทศออนไลน์ตาดีกาและ ศอม. จังหวัดนราธิวาส:\n" +
    "- อธิบายวิธีการเข้าใช้งานระบบ การล็อกอิน\n" +
    "- อธิบายแบบบันทึกสำหรับผู้นิเทศทั่วไป 2 ด้าน 12 ข้อ และส่วนสรุป/ติดตามผล\n" +
    "- อธิบายเกณฑ์ 2 = ทำได้ชัดเจน, 1 = กำลังพัฒนา, 0 = ต้องได้รับการช่วยเหลือ, N/A = พิจารณาไม่ได้ โดยไม่ใช้เพื่อจัดอันดับ\n\n" +
    "ข้อควรระวังและขอบเขต (Boundaries):\n" +
    "- หากคำถามไม่อยู่ในขอบเขตการศึกษาเอกชนจังหวัดนราธิวาส ให้ตอบอย่างสุภาพว่า \"ขออภัยค่ะ น้องศึกษามีข้อมูลเฉพาะด้านการจัดการศึกษาเอกชน ศูนย์ตาดีกา สถาบันศึกษาปอเนาะ และระบบนิเทศออนไลน์ของจังหวัดนราธิวาสเท่านั้นค่ะ\"\n" +
    "- หากผู้ใช้ถามข้อมูลสถิติ ให้ตอบเป็นตารางหรือ Bullet points เพื่อให้อ่านง่าย\n" +
    "- ห้ามให้คำแนะนำทางการแพทย์ การเงิน หรือเรื่องส่วนตัวเด็ดขาด\n" +
    "- หากไม่ทราบข้อมูลที่แน่ชัด ให้แนะนำผู้ใช้ติดต่อกลุ่มงานที่เกี่ยวข้องของสำนักงานการศึกษาเอกชนจังหวัดนราธิวาส\n\n" +
    "รูปแบบการตอบ (Output Format):\n" +
    "- ใช้ Markdown ในการจัดรูปแบบ (เช่น ทำตัวหนาที่หัวข้อ, ทำรายการ (List), หรือสร้างตารางถ้าจำเป็น)\n" +
    "- ตอบให้กระชับ ตรงประเด็น แต่อ่านแล้วรู้สึกถึงความช่วยเหลือและเป็นมิตร";

  // นำคำถามผู้ใช้ มารวมกับข้อมูลสถิติที่ดึงมาได้
  let promptText = systemPrompt + "\n\n";
  if (contextData !== "") {
    promptText += "โปรดใช้ข้อมูลความจริงต่อไปนี้ในการตอบคำถามผู้ใช้ หากตัวเลขเป็น 0 แปลว่ายังไม่ได้บันทึกข้อมูล:\n" + contextData + "\n\n";
  }
  promptText += "คำถามจากผู้ใช้: " + userMessage;

  const payload = {
    contents: [{ parts: [{ text: promptText }] }],
    generationConfig: {
      temperature: 0.3, // เน้นตอบตรงข้อมูลความจริง
      maxOutputTokens: 500,
    }
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  // ลองเรียกแต่ละโมเดลตามลำดับ: ตัวไหนตอบสำเร็จให้ใช้ทันที ถ้า error/quota ให้ถอยไปตัวถัดไป
  for (let i = 0; i < MODELS.length; i++) {
    const API_URL = 'https://generativelanguage.googleapis.com/v1beta/models/' + MODELS[i] + ':generateContent?key=' + API_KEY;
    try {
      const response = UrlFetchApp.fetch(API_URL, options);
      const data = JSON.parse(response.getContentText());
      if (data.candidates && data.candidates.length > 0) {
        return data.candidates[0].content.parts[0].text;
      }
      // ยังไม่สำเร็จ เก็บข้อความ error ของโมเดลสุดท้ายไว้แจ้งถ้าครบทุกตัวแล้ว
      if (i === MODELS.length - 1) {
        if (data.error) {
          return "AI แจ้งข้อผิดพลาด (รหัส " + (data.error.code || '?') + "): " + (data.error.message || 'ไม่ทราบสาเหตุ');
        }
        return "ขออภัยค่ะ ขณะนี้น้องศึกษาไม่สามารถดึงข้อมูลได้ค่ะ";
      }
    } catch (error) {
      if (i === MODELS.length - 1) {
        return "AI ขัดข้องชั่วคราว: " + error;
      }
      // ลองโมเดลถัดไป
    }
  }
  return "ขออภัยค่ะ ขณะนี้น้องศึกษาไม่สามารถดึงข้อมูลได้ค่ะ";
}
