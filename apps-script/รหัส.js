const SHEET_ID = '1qk9eLhwKgPvh2fLwWNSthV4JKyDkJGqJojhLDus5460'; 
const SHEET_NAME1 = 'DATA';
const SHEET_NAME2 = 'ADDRESS';
const SHEET_NAME3 = 'USERS';

// คอลัมน์ทั้งหมดของชีต DATA (คอลัมน์ 10-11 ใช้ติดตามการแก้ไข)
const DATA_HEADERS = ['Timestamp', 'ID ศูนย์', 'ชื่อศูนย์', 'ประเภทการประเมิน', 'คะแนน 1', 'คะแนน 2', 'รวม', 'ผู้นิเทศ', 'ข้อเสนอแนะ', 'แก้ไขครั้งล่าสุด', 'ผู้แก้ไขล่าสุด'];

function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);
    const action = req.action;
    const payload = req.payload;
    let res = {};

    if (action === 'login') {
      res = loginUser(payload);
    } else if (action === 'getTadikaList') {
      res = getTadikaList();
    } else if (action === 'getTadikaData') {
      res = getTadikaData(payload);
    } else if (action === 'getEvaluations') {
      res = getEvaluations(payload);
    } else if (action === 'saveEvaluation') {
      res = saveEvaluation(payload);
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

// --- ขั้นตอนที่ 1: ตรวจสอบ Login ---
function loginUser(data) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('USERS');
  if(!sheet) return {success: false, message: "ไม่พบชีต USERS"};
  
  // สมมติว่าแถวที่ 1 เป็น Header ข้อมูลเริ่มแถวที่ 2
  const rows = sheet.getDataRange().getValues(); 
  
  // แปลงค่าที่ส่งมาจากหน้าเว็บให้เป็น String และตัดช่องว่างซ้ายขวา (Trim)
  const inputUser = String(data.username).trim();
  const inputPass = String(data.password).trim();

  for(let i = 1; i < rows.length; i++) {
    // แปลงข้อมูลที่ดึงมาจาก Sheet ให้เป็น String และตัดช่องว่างซ้ายขวาด้วย
    let sheetUser = String(rows[i][0]).trim();
    let sheetPass = String(rows[i][1]).trim();
    
    // เปรียบเทียบข้อมูลที่ Clean แล้ว
    if(sheetUser === inputUser && sheetPass === inputPass) {
      return {
        success: true, 
        userData: { fname: rows[i][2], tel: rows[i][3] }
      };
    }
  }
  
  // หากไม่สำเร็จ ให้พ่นค่าที่หน้าเว็บส่งมากลับไปใน message เพื่อช่วยให้เรา Debug ได้
  return {
    success: false, 
    message: "Username หรือ Password ไม่ถูกต้อง\n\n(ข้อมูลที่ระบบได้รับ: User='" + inputUser + "', Pass='" + inputPass + "')"
  };
}

// --- ขั้นตอนที่ 2: ดึงรายชื่อศูนย์มาให้เลือก ---
function getTadikaList() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME2);
  if(!sheet) return {success: false, message: "ไม่พบชีต ADDRESS"};
  
  const lastRow = sheet.getLastRow();
  if(lastRow < 6) return {success: true, data: []};
  
  // ดึงข้อมูลคอลัมน์ A ถึง D เริ่มจากแถวที่ 6
  const rows = sheet.getRange(6, 1, lastRow - 5, 4).getValues(); 
  let list = [];
  for(let i = 0; i < rows.length; i++) {
    if(rows[i][0] != "") { // ถ้ามีรหัส ID
      list.push({ id: rows[i][0], name: rows[i][3] }); // A=ID, D=ชื่อศูนย์
    }
  }
  return {success: true, data: list};
}

// --- ขั้นตอนที่ 3: ดึงข้อมูลรายละเอียดของศูนย์ที่เลือก ---
function getTadikaData(id) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME2);
  const lastRow = sheet.getLastRow();
  // ดึงข้อมูลตั้งแต่ A (1) ถึง T (20) เริ่มแถว 6
  const rows = sheet.getRange(6, 1, lastRow - 5, 20).getValues(); 
  
  for(let i = 0; i < rows.length; i++) {
    if(rows[i][0] == id) {
      return {
        success: true,
        data: {
          row: i + 6, // เก็บเลขแถวไว้ใช้อัปเดตข้อมูล
          id: rows[i][0], status: rows[i][1], mosque: rows[i][2], name: rows[i][3], 
          head: rows[i][4], eduSec: rows[i][5], eduRel: rows[i][6], foundedDate: rows[i][7], 
          regNum: rows[i][8], address: rows[i][9], subdist: rows[i][10], dist: rows[i][11], 
          phone: rows[i][12], size: rows[i][13], tMale: rows[i][14], tFemale: rows[i][15], 
          tTotal: rows[i][16], sMale: rows[i][17], sFemale: rows[i][18], sTotal: rows[i][19]
        }
      };
    }
  }
  return {success: false, message: "ไม่พบข้อมูลศูนย์"};
}

// --- เตรียมชีต DATA ให้มี Header ครบตาม DATA_HEADERS ---
function ensureDataSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME1);
  if(!sheet) {
    sheet = ss.insertSheet(SHEET_NAME1);
    sheet.getRange(1, 1, 1, DATA_HEADERS.length).setValues([DATA_HEADERS]);
    return sheet;
  }
  const headers = sheet.getRange(1, 1, 1, DATA_HEADERS.length).getValues()[0];
  const isEmpty = headers.every(h => String(h).trim() == '');
  if(isEmpty) {
    sheet.getRange(1, 1, 1, DATA_HEADERS.length).setValues([DATA_HEADERS]);
  } else {
    // เติมคอลัมน์ติดตามการแก้ไข หากยังไม่มี (รองรับชีตเดิมที่ยังไม่มีคอลัมน์ 10-11)
    if(headers[9] == undefined || String(headers[9]).trim() == '') sheet.getRange(1, 10).setValue('แก้ไขครั้งล่าสุด');
    if(headers[10] == undefined || String(headers[10]).trim() == '') sheet.getRange(1, 11).setValue('ผู้แก้ไขล่าสุด');
  }
  return sheet;
}

function formatDate(d) {
  if(!d) return '';
  if(!(d instanceof Date)) d = new Date(d);
  const pad = n => ('0' + n).slice(-2);
  return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()) + ' ' +
         pad(d.getHours()) + ':' + pad(d.getMinutes()) + ':' + pad(d.getSeconds());
}

// --- ขั้นตอนที่ 4: ดึงประวัติการนิเทศของศูนย์ที่เลือก (สำหรับกลับมาแก้ไขได้) ---
function getEvaluations(tadikaId) {
  const sheet = ensureDataSheet();
  const lastRow = sheet.getLastRow();
  if(lastRow < 2) return {success: true, data: []};

  const rows = sheet.getRange(2, 1, lastRow - 1, DATA_HEADERS.length).getValues();
  const list = [];
  for(let i = 0; i < rows.length; i++) {
    if(String(rows[i][1]).trim() == String(tadikaId).trim()) {
      list.push({
        row: i + 2, // เลขแถวจริงในชีต DATA ใช้อัปเดตเมื่อแก้ไข
        timestamp: formatDate(rows[i][0]),
        id: rows[i][1],
        name: rows[i][2],
        formType: rows[i][3],
        score1: rows[i][4],
        score2: rows[i][5],
        totalScore: rows[i][6],
        supervisor: rows[i][7],
        comment: rows[i][8],
        lastEdit: formatDate(rows[i][9]),
        lastEditor: rows[i][10]
      });
    }
  }
  list.sort((a, b) => a.timestamp < b.timestamp ? 1 : -1);
  return {success: true, data: list};
}

// --- ขั้นตอนที่ 5: บันทึกอัปเดตข้อมูลศูนย์ และ บันทึก/แก้ไขผลนิเทศ (ประทับเวลา + ชื่อผู้แก้ไขทุกครั้ง) ---
function saveEvaluation(payload) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  
  // 5.1 อัปเดตข้อมูลศูนย์ลงชีต ADDRESS
  const addressSheet = ss.getSheetByName(SHEET_NAME2);
  const t = payload.tadikaData;
  const updateValues = [[
    t.status, t.mosque, t.name, t.head, t.eduSec, t.eduRel, t.foundedDate, t.regNum,
    t.address, t.subdist, t.dist, t.phone, t.size, t.tMale, t.tFemale, t.tTotal,
    t.sMale, t.sFemale, t.sTotal
  ]];
  if(addressSheet && t.row) {
    addressSheet.getRange(t.row, 2, 1, 19).setValues(updateValues);
  }

  // 5.2 บันทึก/แก้ไขผลนิเทศลงชีต DATA
  const dataSheet = ensureDataSheet();
  const e = payload.evalData;
  const supervisor = String(payload.supervisor);
  const now = new Date();

  if(payload.editRow && !isNaN(payload.editRow)) {
    // แก้ไขบันทึกเดิม: อัปเดตค่าพร้อมประทับเวลาผู้แก้ไข (เก็บวันแรก/ผู้นิเทศคนแรกไว้)
    const row = Number(payload.editRow);
    dataSheet.getRange(row, 4).setValue(e.formType);
    dataSheet.getRange(row, 5).setValue(e.score1);
    dataSheet.getRange(row, 6).setValue(e.score2);
    dataSheet.getRange(row, 7).setValue(e.totalScore);
    dataSheet.getRange(row, 9).setValue(e.comment);
    dataSheet.getRange(row, 10).setValue(now);
    dataSheet.getRange(row, 11).setValue(supervisor);
    return {
      success: true,
      message: 'แก้ไขผลการนิเทศเรียบร้อยแล้ว!<br>แก้ไขครั้งล่าสุด: ' + formatDate(now) + ' โดย ' + supervisor,
      lastEdit: formatDate(now),
      lastEditor: supervisor
    };
  }

  // บันทึกใหม่: ประทับเวลา + ผู้นิเทศ (และใช้เป็นผู้แก้ไขครั้งแรกด้วย)
  dataSheet.appendRow([now, t.id, t.name, e.formType, e.score1, e.score2, e.totalScore, supervisor, e.comment, now, supervisor]);
  return {success: true, message: 'อัปเดตข้อมูลศูนย์ และบันทึกผลการนิเทศเรียบร้อยแล้ว!', lastEdit: formatDate(now), lastEditor: supervisor};
}

function processChatbot(userMessage) {
  // บอทยังคงทำงานแบบเดิม
  const msg = userMessage.toLowerCase();
  let reply = "น้องศึกษาพร้อมให้คำแนะนำค่ะ";
  if (msg.includes('เกณฑ์')) reply = "เกณฑ์การนิเทศมี 5 รูปแบบค่ะ...";
  return { reply: reply };
}