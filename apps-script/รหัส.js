const SHEET_ID = '1qk9eLhwKgPvh2fLwWNSthV4JKyDkJGqJojhLDus5460'; 
const SHEET_NAME1 = 'DATA';
const SHEET_NAME2 = 'ADDRESS';
const SHEET_NAME3 = 'USERS';

// โครงสร้างใหม่ของชีต DATA (15 คอลัมน์ รองรับการประเมิน 4 ด้าน + สรุป)
const DATA_HEADERS = ['Timestamp', 'ID ศูนย์', 'ชื่อศูนย์', 'ประเภทการประเมิน', 'คะแนนแบบ1', 'คะแนนแบบ2', 'คะแนนแบบ3', 'คะแนนแบบ4', 'รวม/150', 'ร้อยละ', 'ระดับ', 'รายละเอียด', 'ผู้นิเทศ', 'แก้ไขครั้งล่าสุด', 'ผู้แก้ไขล่าสุด'];

function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);
    const action = req.action;
    const payload = req.payload;
    let res = {};

    if (action === 'login') {
      res = loginUser(payload);
    } else if (action === 'register') {
      res = registerUser(payload);
    } else if (action === 'getTadikaList') {
      res = getTadikaList();
    } else if (action === 'getStats') {
      res = getStats();
    } else if (action === 'getTadikaData') {
      res = getTadikaData(payload);
    } else if (action === 'getEvaluations') {
      res = getEvaluations(payload);
    } else if (action === 'saveEvaluation') {
      res = saveEvaluation(payload);
    } else if (action === 'savePin') {
      res = savePin(payload);
    } else if (action === 'chat') {
      res = processChatbot(payload);
    }

    return ContentService.createTextOutput(JSON.stringify(res))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// --- ตรวจสอบ Login ---
function loginUser(data) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('USERS');
  if(!sheet) return {success: false, message: "ไม่พบชีต USERS"};
  
  const rows = sheet.getDataRange().getValues(); 
  const inputUser = String(data.username).trim();
  const inputPass = String(data.password).trim();

  for(let i = 1; i < rows.length; i++) {
    let sheetUser = String(rows[i][0]).trim();
    let sheetPass = String(rows[i][1]).trim();
    
    if(sheetUser === inputUser && sheetPass === inputPass) {
      return {
        success: true, 
        userData: { fname: rows[i][2], tel: rows[i][3] }
      };
    }
  }
  
  return {
    success: false, 
    message: "Username หรือ Password ไม่ถูกต้อง\n\n(ข้อมูลที่ระบบได้รับ: User='" + inputUser + "', Pass='" + inputPass + "')"
  };
}

// --- สมัครสมาชิกใหม่ (เพิ่มผู้ใช้ลงชีต USERS) ---
function registerUser(data) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME3);
  if(!sheet) return {success: false, message: "ไม่พบชีต USERS"};

  const rows = sheet.getDataRange().getValues();
  const u = String(data.username).trim();
  const p = String(data.password).trim();
  const f = String(data.fname).trim();
  const tel = String(data.tel || '').trim();

  if(!u || !p || !f) return {success: false, message: "กรอกข้อมูลไม่ครบถ้วน (ต้องมี Username, Password และชื่อ-นามสกุล)"};

  for(let i = 1; i < rows.length; i++) {
    if(String(rows[i][0]).trim() === u) {
      return {success: false, message: "Username นี้ถูกใช้งานแล้ว กรุณาใช้ชื่ออื่น"};
    }
  }

  sheet.appendRow([u, p, f, tel]);
  return {success: true, message: "สมัครสมาชิกเรียบร้อยแล้ว! กรุณาเข้าสู่ระบบ"};
}

// --- ดึงรายชื่อศูนย์มาให้เลือก (Autocomplete) ---
function getTadikaList() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME2);
  if(!sheet) return {success: false, message: "ไม่พบชีต ADDRESS"};
  
  const lastRow = sheet.getLastRow();
  if(lastRow < 6) return {success: true, data: []};
  
  // คอลัมน์: A=ID, C=มัสยิด, D=ชื่อศูนย์, J=ที่อยู่, K=ตำบล, L=อำเภอ, U=ละติจูด, V=ลองจิจูด
  const rows = sheet.getRange(6, 1, lastRow - 5, 22).getValues(); 
  let list = [];
  for(let i = 0; i < rows.length; i++) {
    if(rows[i][0] != "") {
      list.push({
        id: rows[i][0],
        name: rows[i][3],
        mosque: rows[i][2],
        address: rows[i][9],
        subdist: rows[i][10],
        dist: rows[i][11],
        lat: rows[i][20],
        lng: rows[i][21]
      });
    }
  }
  return {success: true, data: list};
}

// --- สถิติระบบนิเทศออนไลน์ (นับจากชีต ADDRESS/DATA/USERS) ---
function getStats() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let totalTadika = 0, totalEval = 0, totalUsers = 0, sumPct = 0;
  const levelCounts = { 'ดีมาก': 0, 'ดี': 0, 'พอใช้': 0, 'ต้องปรับปรุง': 0 };
  const latest = [];

  const addr = ss.getSheetByName(SHEET_NAME2);
  if(addr) {
    const lastRow = addr.getLastRow();
    if(lastRow >= 6) {
      const vals = addr.getRange(6, 1, lastRow - 5, 1).getValues();
      totalTadika = vals.filter(r => String(r[0]).trim() !== '').length;
    }
  }

  const data = ss.getSheetByName(SHEET_NAME1);
  if(data) {
    const lastRow = data.getLastRow();
    if(lastRow >= 2) {
      const vals = data.getRange(2, 1, lastRow - 1, 15).getValues();
      totalEval = vals.length;
      vals.forEach(r => {
        const pct = Number(r[9]);
        if(!isNaN(pct)) {
          sumPct += pct;
          const lvl = String(r[10] || '').trim();
          if(lvl in levelCounts) levelCounts[lvl]++;
        }
        latest.push({ name: String(r[2] || ''), timestamp: formatDate(r[0]), pct: isNaN(pct) ? 0 : pct, level: r[10] });
      });
      latest.sort((a, b) => a.timestamp < b.timestamp ? 1 : -1);
    }
  }

  const users = ss.getSheetByName(SHEET_NAME3);
  if(users) {
    const lastRow = users.getLastRow();
    if(lastRow >= 2) totalUsers = lastRow - 1;
  }

  return {
    success: true,
    data: {
      totalTadika: totalTadika,
      totalEval: totalEval,
      totalUsers: totalUsers,
      avgPct: totalEval ? Math.round(sumPct / totalEval) : 0,
      levelCounts: levelCounts,
      latest: latest.slice(0, 5)
    }
  };
}

// --- ดึงข้อมูลรายละเอียดของศูนย์ที่เลือก ---
function getTadikaData(id) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME2);
  const lastRow = sheet.getLastRow();
  const rows = sheet.getRange(6, 1, lastRow - 5, 22).getValues(); 
  
  for(let i = 0; i < rows.length; i++) {
    if(rows[i][0] == id) {
      return {
        success: true,
        data: {
          row: i + 6,
          id: rows[i][0], status: rows[i][1], mosque: rows[i][2], name: rows[i][3], 
          head: rows[i][4], eduSec: rows[i][5], eduRel: rows[i][6], foundedDate: rows[i][7], 
          regNum: rows[i][8], address: rows[i][9], subdist: rows[i][10], dist: rows[i][11], 
          phone: rows[i][12], size: rows[i][13], tMale: rows[i][14], tFemale: rows[i][15], 
          tTotal: rows[i][16], sMale: rows[i][17], sFemale: rows[i][18], sTotal: rows[i][19],
          lat: rows[i][20], lng: rows[i][21]
        }
      };
    }
  }
  return {success: false, message: "ไม่พบข้อมูลศูนย์"};
}

// --- บันทึกพิกัด GPS ของศูนย์ (ชีต ADDRESS: คอลัมน์ U=ละติจูด, V=ลองจิจูด) ---
function savePin(data) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME2);
  if(!sheet) return {success: false, message: "ไม่พบชีต ADDRESS"};
  const id = String(data && data.id ? data.id : '').trim();
  const lat = String(data && data.lat != null ? data.lat : '').trim();
  const lng = String(data && data.lng != null ? data.lng : '').trim();
  if(!id) return {success: false, message: "ไม่พบรหัสศูนย์ (ID)"};
  if(!lat || !lng || isNaN(Number(lat)) || isNaN(Number(lng))) {
    return {success: false, message: "พิกัดไม่ถูกต้อง (ต้องเป็นตัวเลขละติจูด/ลองจิจูด)"};
  }
  const lastRow = sheet.getLastRow();
  if(lastRow < 6) return {success: false, message: "ชีต ADDRESS ยังไม่มีข้อมูลศูนย์"};
  const rows = sheet.getRange(6, 1, lastRow - 5, 22).getValues();
  for(let i = 0; i < rows.length; i++) {
    if(String(rows[i][0]).trim() === id) {
      const row = i + 6;
      sheet.getRange(row, 21).setValue(Number(lat));
      sheet.getRange(row, 22).setValue(Number(lng));
      return {success: true, message: "บันทึกพิกัดเรียบร้อย (U=ละติจูด, V=ลองจิจูด)", row: row};
    }
  }
  return {success: false, message: "ไม่พบข้อมูลศูนย์นี้ในชีต ADDRESS"};
}

// --- เตรียมชีต DATA ให้มี Header ครบตาม DATA_HEADERS (15 คอลัมน์) ---
function ensureDataSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME1);
  if(!sheet) {
    sheet = ss.insertSheet(SHEET_NAME1);
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

// --- ดึงประวัติการนิเทศของศูนย์ที่เลือก (สำหรับกลับมาแก้ไขได้) ---
function getEvaluations(tadikaId) {
  const sheet = ensureDataSheet();
  const lastRow = sheet.getLastRow();
  if(lastRow < 2) return {success: true, data: []};

  const rows = sheet.getRange(2, 1, lastRow - 1, DATA_HEADERS.length).getValues();
  const list = [];
  for(let i = 0; i < rows.length; i++) {
    if(String(rows[i][1]).trim() == String(tadikaId).trim()) {
      let details = null;
      try { details = JSON.parse(rows[i][11] || 'null'); } catch(e) { details = null; }
      list.push({
        row: i + 2,
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
  list.sort((a, b) => a.timestamp < b.timestamp ? 1 : -1);
  return {success: true, data: list};
}

// --- บันทึก/แก้ไขผลนิเทศ + อัปเดตข้อมูลศูนย์ (ประทับเวลาและผู้แก้ไขทุกครั้ง) ---
function saveEvaluation(payload) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  
  // 5.1 อัปเดตข้อมูลศูนย์ลงชีต ADDRESS
  const addressSheet = ss.getSheetByName(SHEET_NAME2);
  const t = payload.tadikaData;
  const updateValues = [[
    t.status, t.mosque, t.name, t.head, t.eduSec, t.eduRel, t.foundedDate, t.regNum,
    t.address, t.subdist, t.dist, t.phone, t.size, t.tMale, t.tFemale, t.tTotal,
    t.sMale, t.sFemale, t.sTotal, t.lat, t.lng
  ]];
  if(addressSheet && t.row) {
    addressSheet.getRange(t.row, 2, 1, 21).setValues(updateValues);
  }

  // 5.2 บันทึก/แก้ไขผลนิเทศลงชีต DATA
  const dataSheet = ensureDataSheet();
  const e = payload.evalData;
  const supervisor = String(payload.supervisor);
  const now = new Date();
  const detailsJSON = JSON.stringify({ answers: e.answers || {}, comment: e.comment || '', strengths: e.strengths || '', improve: e.improve || '' });

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
    dataSheet.getRange(row, 14).setValue(now);
    dataSheet.getRange(row, 15).setValue(supervisor);
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

function processChatbot(userMessage) {
  const msg = String(userMessage || '').toLowerCase();
  let reply = "สวัสดีครับ ผมชื่อน้องศึกษา พร้อมให้คำแนะนำเรื่องเกณฑ์การประเมินหรือวิธีการกรอกคะแนนได้เลยนะครับ";
  if (msg.includes('เกณฑ์')) {
    reply = "เกณฑ์การประเมินมี 4 ด้าน รวมคะแนนเต็ม 150 คะแนนครับ:\n" +
      "1) ด้านการบริหารจัดการศูนย์ (40 คะแนน)\n" +
      "2) ด้านครูผู้สอน (40 คะแนน)\n" +
      "3) ด้านผู้เรียน (35 คะแนน)\n" +
      "4) ด้านอาคารสถานที่และสิ่งแวดล้อม (35 คะแนน)\n" +
      "รวมแล้วคิดเป็นร้อยละและจัดระดับผลการนิเทศครับ";
  } else if (msg.includes('ระดับ')) {
    reply = "การจัดระดับผลการนิเทศครับ:\n" +
      "- ร้อยละ 80 ขึ้นไป  → ดีมาก\n" +
      "- ร้อยละ 60-79   → ดี\n" +
      "- ร้อยละ 40-59   → พอใช้\n" +
      "- ต่ำกว่าร้อยละ 40 → ต้องปรับปรุง";
  } else if (msg.includes('คะแนน') || msg.includes('กรอก')) {
    reply = "กรอกคะแนนโดยเลือกจำนวนคะแนนในแต่ละข้อ (สูงสุดข้อละ 5 คะแนน) ในแต่ละด้านครับ ระบบจะคำนวณคะแนนรวม ร้อยละ และระดับผลการนิเทศให้อัตโนมัติ";
  } else if (msg.includes('แก้ไข')) {
    reply = "สามารถกลับมาแก้ไขผลการนิเทศได้จากตารางประวัติการนิเทศของศูนย์นั้น โดยกดปุ่ม \"แก้ไข\" แล้วบันทึกใหม่ ระบบจะประทับเวลาและชื่อผู้แก้ไขล่าสุดให้ครับ";
  } else if (msg.includes('สวัสดี') || msg.includes('hello') || msg.includes('hi')) {
    reply = "สวัสดีครับ ผมน้องศึกษา พร้อมช่วยเหลือเรื่องการนิเทศออนไลน์เสมอนะครับ";
  }
  return { reply: reply };
}
