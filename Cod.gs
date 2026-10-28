// ==================================================
// CONFIG & SETUP
// ==================================================
const PROP_KEY_SHEET_ID = 'SHEET_ID';
const PROP_KEY_ADMIN_HASH = 'ADMIN_PASSWORD_HASH';

// โครงตารางของแต่ละชีต (ใช้ตอนสร้างชีตใหม่เท่านั้น)
const SHEET_DEFS = {
  'Students':   ['student_id', 'student_name', 'student_class', 'created_at'],
  'Subjects':   ['subject_name', 'subject_class', 'created_at'],
  'Attendance': ['attendance_date', 'attendance_student_id', 'attendance_subject',
                 'attendance_status', 'attendance_notes', 'created_at']
};

// ⚠️ รันฟังก์ชันนี้ "ครั้งเดียว" เพื่อตั้งค่าเริ่มต้น
//    (รันซ้ำได้: จะคงค่า Sheet ID / รหัสผ่าน Admin เดิมไว้ ไม่รีเซ็ต)
function setup() {
  const props = PropertiesService.getScriptProperties();

  // ตั้งค่า Sheet ID ของคุณ ถ้ายังไม่ได้ตั้ง
  if (!props.getProperty(PROP_KEY_SHEET_ID)) {
    props.setProperty(PROP_KEY_SHEET_ID, '1ya65apZf-4ODvqMFdswWsd0haSVoSvow0Xew633zF1w');
  }

  // ตั้งรหัสผ่าน Admin ถ้ายังไม่ได้ตั้ง (เปลี่ยน 'admin1234' เป็นรหัสที่ต้องการ)
  if (!props.getProperty(PROP_KEY_ADMIN_HASH)) {
    props.setProperty(PROP_KEY_ADMIN_HASH, hashPassword_('admin1234'));
    Logger.log('✅ Setup complete. Admin password is: admin1234');
  } else {
    Logger.log('✅ Setup: คงค่า Sheet ID และรหัสผ่าน Admin เดิมไว้ (ไม่ทับ)');
  }

  ensureSheets_();
}

// สร้างชีตที่ยังไม่มี พร้อมหัวตาราง
function ensureSheets_() {
  const ss = getSpreadsheet_();
  Object.keys(SHEET_DEFS).forEach(function (name) {
    if (!ss.getSheetByName(name)) {
      const headers = SHEET_DEFS[name];
      ss.insertSheet(name).getRange(1, 1, 1, headers.length).setValues([headers]);
      Logger.log('🆕 สร้างชีตใหม่: ' + name);
    }
  });
}

// ดึงชีตแบบแจ้งข้อผิดพลาดชัดเจน (ป้องกันอ่านผิดชีต/ผิด Sheet ID แบบเงียบ ๆ)
function getSheet_(name) {
  const sh = getSpreadsheet_().getSheetByName(name);
  if (!sh) {
    throw new Error('ไม่พบชีต "' + name + '" ในสเปรดชีต (Sheet ID: ' + getSheetId_() +
      ') — ตรวจ Sheet ID หรือรัน setup() เพื่อสร้างชีต');
  }
  return sh;
}

function getSheetId_() {
  return PropertiesService.getScriptProperties().getProperty(PROP_KEY_SHEET_ID);
}

function getSpreadsheet_() {
  return SpreadsheetApp.openById(getSheetId_());
}

// ==================================================
// UTILITIES
// ==================================================
function hashPassword_(password) {
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256, password, Utilities.Charset.UTF_8
  );
  return digest.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

function verifyAdmin_(password) {
  const storedHash = PropertiesService.getScriptProperties().getProperty(PROP_KEY_ADMIN_HASH);
  return storedHash && storedHash === hashPassword_(password);
}

function todayThaiStr_() {
  return Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd');
}

function normalizeDate_(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Bangkok', 'yyyy-MM-dd');
  return String(v);
}

// ==================================================
// WEB APP ENTRY POINT
// ==================================================
function doGet(e) {
  // สำหรับทดสอบการเชื่อมต่อ
  if (e && e.parameter && e.parameter.action) {
    try {
      const action = e.parameter.action;
      const data = e.parameter.data ? JSON.parse(e.parameter.data) : null;
      const result = handleAction_(action, data);
      return ContentService.createTextOutput(JSON.stringify({ ok: true, data: result })).setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({ ok: false, error: err.message })).setMimeType(ContentService.MimeType.JSON);
    }
  }
  return ContentService.createTextOutput('API is running');
}

function doPost(e) {
  try {
    // ✅ อ่าน body ที่ส่งมาเป็น text/plain (หลีกเลี่ยง CORS preflight)
    const payload = JSON.parse(e.postData.contents);
    const result = handleAction_(payload.action, payload.data);
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, data: result }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function handleAction_(action, data) {
  switch (action) {
    case 'getAllData':       return getAllData_();
    case 'addStudent':       return addStudent_(data); 
    case 'addBulkStudents':  return addBulkStudents_(data); 
    case 'updateStudent':    return updateStudent_(data); 
    case 'deleteStudent':    return deleteStudent_(data.id); 
    case 'addSubject':       return addSubject_(data); 
    case 'addBulkSubjects':  return addBulkSubjects_(data); 
    case 'updateSubject':    return updateSubject_(data); 
    case 'deleteSubject':    return deleteSubject_(data.id); 
    case 'saveAttendance':   return saveAttendance_(data); 
    case 'updateAttendance': return updateAttendance_(data); 
    case 'deleteAttendance': return deleteAttendance_(data.id); 
    case 'updateSheetId':
      if (!verifyAdmin_(data.password)) throw new Error('รหัสผ่าน Admin ไม่ถูกต้อง');
      PropertiesService.getScriptProperties().setProperty(PROP_KEY_SHEET_ID, data.sheetId);
      return { ok: true, message: 'บันทึก Sheet ID ใหม่แล้ว' };
    default: throw new Error('ไม่รู้จัก action: ' + action);
  }
}

// ==================================================
// CRUD: STUDENTS
// ==================================================
function getAllData_() {
  const ss = getSpreadsheet_();
  const result = { students: [], subjects: [], attendance: [] };

  const sSh = getSheet_('Students');
  if (sSh.getLastRow() > 1) {
    const rows = sSh.getDataRange().getValues().slice(1);
    result.students = rows.map((r, i) => ({
      _id: 'STU_' + (i + 2), // ใช้ row index เป็น id
      student_id: String(r[0]),
      student_name: String(r[1]),
      student_class: String(r[2]),
      created_at: r[3] ? String(r[3]) : ''
    }));
  }

  const subSh = getSheet_('Subjects');
  if (subSh.getLastRow() > 1) {
    const rows = subSh.getDataRange().getValues().slice(1);
    result.subjects = rows.map((r, i) => ({
      _id: 'SUB_' + (i + 2),
      subject_name: String(r[0]),
      subject_class: String(r[1]),
      created_at: r[2] ? String(r[2]) : ''
    }));
  }

  const aSh = getSheet_('Attendance');
  if (aSh.getLastRow() > 1) {
    const rows = aSh.getDataRange().getValues().slice(1);
    result.attendance = rows.map((r, i) => ({
      _id: 'ATT_' + (i + 2),
      attendance_date: normalizeDate_(r[0]),
      attendance_student_id: String(r[1]),
      attendance_subject: String(r[2]),
      attendance_status: String(r[3]),
      attendance_notes: String(r[4] || ''),
      created_at: r[5] ? String(r[5]) : ''
    }));
  }
  return result;
}

function addStudent_(data) {
  const sh = getSheet_('Students');
  const existing = sh.getDataRange().getValues().slice(1);
  if (existing.some(r => String(r[0]) === String(data.student_id) && String(r[2]) === String(data.student_class))) {
    throw new Error('มีนักเรียนเลขที่นี้ในชั้นนี้แล้ว');
  }
  sh.appendRow([data.student_id, data.student_name, data.student_class, new Date()]);
  return { ok: true };
}

function addBulkStudents_(data) {
  const sh = getSheet_('Students');
  const existing = new Set(sh.getDataRange().getValues().slice(1).map(r => String(r[0]) + '|' + String(r[2])));
  const rows = [];
  let added = 0, skipped = 0;
  
  data.list.forEach(s => {
    const key = String(s.student_id) + '|' + String(s.student_class);
    if (existing.has(key)) { skipped++; return; }
    rows.push([s.student_id, s.student_name, s.student_class, new Date()]);
    existing.add(key);
    added++;
  });
  
  if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, 4).setValues(rows);
  return { ok: true, added, skipped };
}

function updateStudent_(data) {
  const sh = getSheet_('Students');
  const rowIndex = parseInt(data.id.replace('STU_', ''), 10);
  if (isNaN(rowIndex)) throw new Error('ID ไม่ถูกต้อง');
  
  const all = sh.getDataRange().getValues();
  const dup = all.slice(1).some((r, i) => 
    (i + 2) !== rowIndex && String(r[0]) === String(data.student_id) && String(r[2]) === String(data.student_class)
  );
  if (dup) throw new Error('เลขที่ซ้ำในชั้นนี้');
  
  sh.getRange(rowIndex, 1, 1, 3).setValues([[data.student_id, data.student_name, data.student_class]]);
  return { ok: true };
}

function deleteStudent_(id) {
  const sh = getSheet_('Students');
  const rowIndex = parseInt(id.replace('STU_', ''), 10);
  if (isNaN(rowIndex)) throw new Error('ID ไม่ถูกต้อง');
  sh.deleteRow(rowIndex);
  return { ok: true };
}

// ==================================================
// CRUD: SUBJECTS
// ==================================================
function addSubject_(data) {
  const sh = getSheet_('Subjects');
  const existing = sh.getDataRange().getValues().slice(1);
  if (existing.some(r => String(r[0]) === String(data.subject_name) && String(r[1]) === String(data.subject_class))) {
    throw new Error('มีวิชานี้ในชั้นนี้แล้ว');
  }
  sh.appendRow([data.subject_name, data.subject_class, new Date()]);
  return { ok: true };
}

function addBulkSubjects_(data) {
  const sh = getSheet_('Subjects');
  const existing = new Set(sh.getDataRange().getValues().slice(1).map(r => String(r[0]) + '|' + String(r[1])));
  const rows = [];
  let added = 0, skipped = 0;
  
  data.list.forEach(name => {
    const key = name + '|' + data.subject_class;
    if (existing.has(key)) { skipped++; return; }
    rows.push([name, data.subject_class, new Date()]);
    existing.add(key);
    added++;
  });
  
  if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, 3).setValues(rows);
  return { ok: true, added, skipped };
}

function updateSubject_(data) {
  const sh = getSheet_('Subjects');
  const rowIndex = parseInt(data.id.replace('SUB_', ''), 10);
  if (isNaN(rowIndex)) throw new Error('ID ไม่ถูกต้อง');
  
  const all = sh.getDataRange().getValues();
  const dup = all.slice(1).some((r, i) =>
    (i + 2) !== rowIndex && String(r[0]) === String(data.subject_name) && String(r[1]) === String(data.subject_class)
  );
  if (dup) throw new Error('วิชาซ้ำในชั้นนี้');
  
  sh.getRange(rowIndex, 1, 1, 2).setValues([[data.subject_name, data.subject_class]]);
  return { ok: true };
}

function deleteSubject_(id) {
  const sh = getSheet_('Subjects');
  const rowIndex = parseInt(id.replace('SUB_', ''), 10);
  if (isNaN(rowIndex)) throw new Error('ID ไม่ถูกต้อง');
  sh.deleteRow(rowIndex);
  return { ok: true };
}

// ==================================================
// CRUD: ATTENDANCE
// ==================================================
function saveAttendance_(records) {
  const sh = getSheet_('Attendance');
  const all = sh.getDataRange().getValues();
  const keyIndex = new Map(); // date|student|subject → rowIndex

  all.slice(1).forEach((r, i) => {
    keyIndex.set(normalizeDate_(r[0]) + '|' + String(r[1]) + '|' + String(r[2]), i + 2);
  });

  let updated = 0, inserted = 0;
  const toAppend = [];

  records.forEach(rec => {
    const key = rec.date + '|' + String(rec.student_id) + '|' + rec.subject;
    if (keyIndex.has(key)) {
      const row = keyIndex.get(key);
      sh.getRange(row, 4, 1, 2).setValues([[rec.status, rec.notes || '']]);
      updated++;
    } else {
      toAppend.push([rec.date, String(rec.student_id), rec.subject, rec.status, rec.notes || '', new Date()]);
    }
  });

  if (toAppend.length) {
    sh.getRange(sh.getLastRow() + 1, 1, toAppend.length, 6).setValues(toAppend);
    inserted = toAppend.length;
  }
  return { ok: true, updated, inserted };
}

function updateAttendance_(data) {
  const sh = getSheet_('Attendance');
  const rowIndex = parseInt(data.id.replace('ATT_', ''), 10);
  if (isNaN(rowIndex)) throw new Error('ID ไม่ถูกต้อง');
  sh.getRange(rowIndex, 4, 1, 2).setValues([[data.attendance_status, data.attendance_notes || '']]);
  return { ok: true };
}

function deleteAttendance_(id) {
  const sh = getSheet_('Attendance');
  const rowIndex = parseInt(id.replace('ATT_', ''), 10);
  if (isNaN(rowIndex)) throw new Error('ID ไม่ถูกต้อง');
  sh.deleteRow(rowIndex);
  return { ok: true };
}
