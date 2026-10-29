// ==================================================
// CONFIG & SETUP
// ==================================================
const PROP_KEY_SHEET_ID = 'SHEET_ID';
const PROP_KEY_ADMIN_HASH = 'ADMIN_PASSWORD_HASH';
const PROP_KEY_LOGIN_USERNAME = 'APP_LOGIN_USERNAME';
const PROP_KEY_LOGIN_PASSWORD_HASH = 'APP_LOGIN_PASSWORD_HASH';
const INITIAL_LOGIN_USERNAME = 'FirstStar';
const INITIAL_LOGIN_PASSWORD_HASH = '4bbe438f18a5e2b59bdf0472eac2a86a957c4e874b002d81f94867692703c646';
const AUTH_SESSION_TTL_SECONDS = 21600; // 6 hours (Apps Script CacheService maximum)
const AUTH_LOGIN_FAILURE_KEY = 'APP_AUTH_LOGIN_FAILURES';
const AUTH_LOGIN_FAILURE_LIMIT = 5;
const AUTH_LOGIN_LOCK_SECONDS = 900;
const CLASS_LIST = ['ป.1', 'ป.2', 'ป.3', 'ป.4', 'ป.5', 'ป.6', 'ม.1', 'ม.2', 'ม.3'];
const ATTENDANCE_STATUS_LIST = ['มา', 'สาย', 'ลา', 'ขาด'];

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
    Logger.log('✅ Setup complete. รหัสผ่าน Admin เดิมถูกตั้งค่าแล้ว');
  }

  // Pin the sole authorized account; store only the password hash.
  props.setProperty(PROP_KEY_LOGIN_USERNAME, INITIAL_LOGIN_USERNAME);
  props.setProperty(PROP_KEY_LOGIN_PASSWORD_HASH, INITIAL_LOGIN_PASSWORD_HASH);
  Logger.log('✅ Setup: บัญชีล็อกอินถูกกำหนดเป็น FirstStar แล้ว');

  ensureSheets_();
}

// สร้างชีตที่ยังไม่มี พร้อมหัวตาราง
function ensureSheets_() {
  const ss = getSpreadsheet_();
  const kinds = Object.keys(SHEET_DEFS);
  kinds.forEach(function (name) {
    if (!ss.getSheetByName(name)) {
      const headers = SHEET_DEFS[name];
      ss.insertSheet(name).getRange(1, 1, 1, headers.length).setValues([headers]);
      Logger.log('🆕 สร้างชีตใหม่: ' + name);
    }
  });
  // เตรียมแท็บแยกตามชั้นสำหรับข้อมูลใหม่
  CLASS_LIST.forEach(function (cls) {
    kinds.forEach(function (kind) {
      const name = classSheetName_(kind, cls);
      let sh = ss.getSheetByName(name);
      if (!sh) sh = ss.insertSheet(name);
      const headers = kind === 'Attendance'
        ? SHEET_DEFS[kind].concat(['attendance_student_class'])
        : SHEET_DEFS[kind];
      if (sh.getLastRow() === 0) sh.getRange(1, 1, 1, headers.length).setValues([headers]);
      else if (kind === 'Attendance' && sh.getLastColumn() < headers.length) {
        sh.getRange(1, headers.length).setValue(headers[headers.length - 1]);
      }
    });
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

function createAuthSession_(username) {
  const token = Utilities.getUuid() + Utilities.getUuid();
  CacheService.getScriptCache().put('APP_AUTH_' + token, username, AUTH_SESSION_TTL_SECONDS);
  return { token: token, username: username, expiresIn: AUTH_SESSION_TTL_SECONDS };
}

function getAuthUser_(token) {
  if (!token || typeof token !== 'string' || token.length > 100) return '';
  return CacheService.getScriptCache().get('APP_AUTH_' + token) || '';
}

function requireAuth_(token) {
  const username = getAuthUser_(token);
  if (!username) throw new Error('กรุณาเข้าสู่ระบบก่อนเพิ่มหรือแก้ไขข้อมูล');
  return username;
}

function withScriptLock_(callback) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    return callback();
  } finally {
    lock.releaseLock();
  }
}

function login_(data) {
  const cache = CacheService.getScriptCache();
  const attemptedUsername = String(data && data.username || '').trim().slice(0, 80);
  const failureKey = AUTH_LOGIN_FAILURE_KEY + '_' + hashPassword_(attemptedUsername).slice(0, 16);
  const failedAttempts = parseInt(cache.get(failureKey) || '0', 10);
  const username = INITIAL_LOGIN_USERNAME;
  const passwordHash = INITIAL_LOGIN_PASSWORD_HASH;
  const credentialsMatch = !!data && attemptedUsername === username &&
    hashPassword_(String(data.password || '')) === passwordHash;

  // A valid owner login must recover immediately, even if earlier typos triggered the cooldown.
  if (credentialsMatch) {
    cache.remove(failureKey);
    return createAuthSession_(username);
  }
  if (failedAttempts >= AUTH_LOGIN_FAILURE_LIMIT) {
    throw new Error('พยายามเข้าสู่ระบบผิดหลายครั้ง กรุณารอ 15 นาทีแล้วลองใหม่');
  }

  cache.put(failureKey, String(failedAttempts + 1), AUTH_LOGIN_LOCK_SECONDS);
  throw new Error('ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง');
}

function logout_(token) {
  if (token && typeof token === 'string' && token.length <= 100) {
    CacheService.getScriptCache().remove('APP_AUTH_' + token);
  }
  return { ok: true };
}

function todayThaiStr_() {
  return Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd');
}

function normalizeDate_(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Bangkok', 'yyyy-MM-dd');
  return String(v);
}

function normalizeClass_(c) {
  if (!c) return '';
  let s = String(c).toLowerCase().replace(/\s+/g, '').replace(/ชั้น/g, '');
  s = s.replace(/[๐-๙]/g, function (d) {
    return String.fromCharCode(d.charCodeAt(0) - 0x0E50 + 48);
  });
  const match = s.match(/(?:ป\.|ประถม|ปีที่|ม\.|มัธยม|ที่)?(\d+)/);
  if (!match) return s;
  const number = match[1];
  if (s.indexOf('ม.') !== -1 || s.indexOf('มัธยม') !== -1) return 'ม.' + number;
  if (s.indexOf('ป.') !== -1 || s.indexOf('ประถม') !== -1) return 'ป.' + number;
  return number;
}

// ==================================================
// WEB APP ENTRY POINT
// ==================================================
function doGet(e) {
  // สำหรับทดสอบการเชื่อมต่อ
  if (e && e.parameter && e.parameter.action) {
    try {
      const action = e.parameter.action;
      if (action !== 'getAllData' && action !== 'getAuthStatus') {
        throw new Error('คำสั่งเปลี่ยนข้อมูลและเข้าสู่ระบบต้องส่งผ่าน POST');
      }
      const data = e.parameter.data ? JSON.parse(e.parameter.data) : null;
      const result = handleAction_(action, data, e.parameter.sessionToken);
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
    const result = handleAction_(payload.action, payload.data, payload.sessionToken);
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, data: result }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function handleAction_(action, data, sessionToken) {
  if (action === 'login') return login_(data);
  if (action === 'logout') return logout_(sessionToken);
  if (action === 'getAuthStatus') return { authenticated: !!getAuthUser_(sessionToken) };
  if (action === 'getAllData') return getAllDataByClass_();
  // Anyone with the link may add attendance; changing an existing record requires login.
  if (action === 'saveAttendance') return saveAttendanceByClass_(data, !!getAuthUser_(sessionToken));

  // Student/subject management, attendance corrections and system settings require a valid session.
  requireAuth_(sessionToken);
  return withScriptLock_(function () {
    switch (action) {
      case 'addStudent':       return addStudentByClass_(data);
      case 'addBulkStudents':  return addBulkStudentsByClass_(data);
      case 'updateStudent':    return updateStudentByClass_(data);
      case 'deleteStudent':    return deleteByClassId_(data.id, 'STU_');
      case 'addSubject':       return addSubjectByClass_(data);
      case 'addBulkSubjects':  return addBulkSubjectsByClass_(data);
      case 'updateSubject':    return updateSubjectByClass_(data);
      case 'deleteSubject':    return deleteByClassId_(data.id, 'SUB_');
      case 'updateAttendance': return updateAttendanceByClass_(data);
      case 'deleteAttendance': return deleteByClassId_(data.id, 'ATT_');
      case 'updateSheetId':
        if (!data || !/^[A-Za-z0-9_-]{20,120}$/.test(String(data.sheetId || ''))) {
          throw new Error('รูปแบบ Google Sheet ID ไม่ถูกต้อง');
        }
        PropertiesService.getScriptProperties().setProperty(PROP_KEY_SHEET_ID, data.sheetId);
        return { ok: true, message: 'บันทึก Sheet ID ใหม่แล้ว' };
      default: throw new Error('ไม่รู้จัก action: ' + action);
    }
  });
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

// ==================================================
// CLASS-SCOPED TABS (new API implementation)
// ==================================================
// Existing tabs are read together with class tabs. New records are written to the
// matching class tab; old tabs stay intact until migrateExistingDataToClassTabs()
// verifies every row has a destination.
function classSheetName_(baseName, cls) {
  const normalized = normalizeClass_(cls);
  if (CLASS_LIST.indexOf(normalized) === -1) throw new Error('ชั้นเรียนไม่ถูกต้อง: ' + cls);
  return baseName + '-' + normalized;
}

function dataSheetEntries_(baseName) {
  const ss = getSpreadsheet_();
  const entries = [];
  const legacy = ss.getSheetByName(baseName);
  if (legacy) entries.push({ name: baseName, cls: '', sheet: legacy });
  CLASS_LIST.forEach(function (cls) {
    const name = baseName + '-' + cls;
    const sh = ss.getSheetByName(name);
    if (sh) entries.push({ name: name, cls: cls, sheet: sh });
  });
  return entries;
}

function ensureClassSheet__(baseName, cls) {
  const name = classSheetName_(baseName, cls);
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  let headers = SHEET_DEFS[baseName].slice();
  if (baseName === 'Attendance') headers.push('attendance_student_class');
  if (sh.getLastRow() === 0) sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  else if (baseName === 'Attendance' && sh.getLastColumn() < headers.length) {
    sh.getRange(1, headers.length).setValue(headers[headers.length - 1]);
  }
  return sh;
}

function encodeRowId_(prefix, row, sheetName) {
  return prefix + row + '@' + sheetName;
}

function decodeRowId_(id, prefix, legacySheet) {
  const value = String(id || '');
  if (value.indexOf(prefix) !== 0) throw new Error('ID ไม่ถูกต้อง');
  const body = value.slice(prefix.length);
  const parts = body.split('@');
  const row = parseInt(parts[0], 10);
  const sheetName = parts.length > 1 ? parts.slice(1).join('@') : legacySheet;
  if (!Number.isInteger(row) || row < 2 || !sheetName) throw new Error('ID ไม่ถูกต้อง');
  const sh = getSpreadsheet_().getSheetByName(sheetName);
  if (!sh) throw new Error('ไม่พบแท็บข้อมูล ' + sheetName + ' กรุณารีเฟรชข้อมูล');
  return { row: row, sheet: sh, sheetName: sheetName };
}

function appendUnique_(sh, row, keyCols) {
  const vals = sh.getDataRange().getValues();
  const key = keyCols.map(function (i) { return String(row[i] == null ? '' : row[i]); }).join('|');
  const exists = vals.slice(1).some(function (r) {
    return keyCols.map(function (i) { return String(r[i] == null ? '' : r[i]); }).join('|') === key;
  });
  if (!exists) sh.appendRow(row);
  return !exists;
}

function getAllDataByClass_() {
  const result = { students: [], subjects: [], attendance: [] };
  const studentClasses = {};
  const subjectClasses = {};
  dataSheetEntries_('Students').forEach(function (entry) {
    entry.sheet.getDataRange().getValues().slice(1).forEach(function (r, i) {
      if (!r || String(r[0]) === '') return;
      const cls = normalizeClass_(r[2] || entry.cls);
      result.students.push({
        _id: encodeRowId_('STU_', i + 2, entry.name),
        student_id: String(r[0]), student_name: String(r[1]), student_class: cls,
        created_at: r[3] ? String(r[3]) : ''
      });
      const studentKey = String(r[0]);
      if (!studentClasses[studentKey]) studentClasses[studentKey] = [];
      if (cls && studentClasses[studentKey].indexOf(cls) === -1) studentClasses[studentKey].push(cls);
    });
  });

  dataSheetEntries_('Subjects').forEach(function (entry) {
    entry.sheet.getDataRange().getValues().slice(1).forEach(function (r, i) {
      if (!r || String(r[0]) === '') return;
      const cls = normalizeClass_(r[1] || entry.cls);
      result.subjects.push({
        _id: encodeRowId_('SUB_', i + 2, entry.name),
        subject_name: String(r[0]), subject_class: cls,
        created_at: r[2] ? String(r[2]) : ''
      });
      const subjectKey = String(r[0]);
      if (!subjectClasses[subjectKey]) subjectClasses[subjectKey] = [];
      if (cls && subjectClasses[subjectKey].indexOf(cls) === -1) subjectClasses[subjectKey].push(cls);
    });
  });

  dataSheetEntries_('Attendance').forEach(function (entry) {
    entry.sheet.getDataRange().getValues().slice(1).forEach(function (r, i) {
      if (!r || !r[0]) return;
      let cls = normalizeClass_(r[6] || entry.cls || '');
      if (!cls) {
        const knownStudentClasses = studentClasses[String(r[1])] || [];
        const knownSubjectClasses = subjectClasses[String(r[2])] || [];
        const candidates = knownStudentClasses.filter(function (c) { return knownSubjectClasses.indexOf(c) !== -1; });
        if (candidates.length === 1) cls = candidates[0];
        else if (knownStudentClasses.length === 1) cls = knownStudentClasses[0];
      }
      result.attendance.push({
        _id: encodeRowId_('ATT_', i + 2, entry.name),
        attendance_date: normalizeDate_(r[0]), attendance_student_id: String(r[1]),
        attendance_subject: String(r[2]), attendance_status: String(r[3]),
        attendance_notes: String(r[4] || ''), created_at: r[5] ? String(r[5]) : '',
        attendance_student_class: cls || ''
      });
    });
  });

  // A migration may leave legacy tabs visible until every row is safely assigned.
  // Prefer class-tab rows and suppress duplicates while retaining unresolved legacy records.
  result.students = dedupeRows_(result.students, function (r) { return normalizeClass_(r.student_class) + '|' + r.student_id; });
  result.subjects = dedupeRows_(result.subjects, function (r) { return normalizeClass_(r.subject_class) + '|' + r.subject_name; });
  result.attendance = dedupeRows_(result.attendance, function (r) {
    const cls = normalizeClass_(r.attendance_student_class);
    // Keep ambiguous legacy rows distinct; their class cannot safely form a dedupe key.
    return cls
      ? cls + '|' + r.attendance_date + '|' + r.attendance_student_id + '|' + r.attendance_subject
      : 'legacy|' + r._id;
  });
  return result;
}

function dedupeRows_(rows, keyOf) {
  const index = {};
  rows.forEach(function (row) { index[keyOf(row)] = row; });
  return Object.keys(index).map(function (key) { return index[key]; });
}

function addStudentByClass_(data) {
  const cls = normalizeClass_(data.student_class);
  const dup = getAllDataByClass_().students.some(function (s) {
    return s.student_id === String(data.student_id) && normalizeClass_(s.student_class) === cls;
  });
  if (dup) throw new Error('มีนักเรียนเลขที่นี้ในชั้นนี้แล้ว');
  ensureClassSheet__('Students', cls).appendRow([data.student_id, data.student_name, cls, new Date()]);
  return { ok: true };
}

function addBulkStudentsByClass_(data) {
  const result = { ok: true, added: 0, skipped: 0 };
  (data.list || []).forEach(function (student) {
    try {
      addStudentByClass_(student);
      result.added++;
    } catch (err) {
      if (String(err.message).indexOf('ซ้ำ') !== -1) result.skipped++;
      else throw err;
    }
  });
  return result;
}

function updateStudentByClass_(data) {
  const loc = decodeRowId_(data.id, 'STU_', 'Students');
  const row = loc.sheet.getRange(loc.row, 1, 1, 4).getValues()[0];
  if (!row || String(row[0]) === '') throw new Error('ไม่พบข้อมูลนักเรียน กรุณารีเฟรช');
  const cls = normalizeClass_(data.student_class);
  const duplicate = getAllDataByClass_().students.some(function (s) {
    return s._id !== data.id && String(s.student_id) === String(data.student_id) && normalizeClass_(s.student_class) === cls;
  });
  if (duplicate) throw new Error('เลขที่ซ้ำในชั้นนี้');
  const destName = classSheetName_('Students', cls);
  if (loc.sheetName === destName) {
    loc.sheet.getRange(loc.row, 1, 1, 4).setValues([[data.student_id, data.student_name, cls, row[3]]]);
  } else {
    ensureClassSheet__('Students', cls).appendRow([data.student_id, data.student_name, cls, row[3] || new Date()]);
    loc.sheet.deleteRow(loc.row);
  }
  return { ok: true };
}

function addSubjectByClass_(data) {
  const cls = normalizeClass_(data.subject_class);
  const duplicate = getAllDataByClass_().subjects.some(function (s) {
    return String(s.subject_name) === String(data.subject_name) && normalizeClass_(s.subject_class) === cls;
  });
  if (duplicate) throw new Error('มีวิชานี้ในชั้นนี้แล้ว');
  ensureClassSheet__('Subjects', cls).appendRow([data.subject_name, cls, new Date()]);
  return { ok: true };
}

function addBulkSubjectsByClass_(data) {
  const result = { ok: true, added: 0, skipped: 0 };
  (data.list || []).forEach(function (name) {
    try {
      addSubjectByClass_({ subject_name: name, subject_class: data.subject_class });
      result.added++;
    } catch (err) {
      if (String(err.message).indexOf('ซ้ำ') !== -1) result.skipped++;
      else throw err;
    }
  });
  return result;
}

function updateSubjectByClass_(data) {
  const loc = decodeRowId_(data.id, 'SUB_', 'Subjects');
  const row = loc.sheet.getRange(loc.row, 1, 1, 3).getValues()[0];
  if (!row || String(row[0]) === '') throw new Error('ไม่พบข้อมูลวิชา กรุณารีเฟรช');
  const cls = normalizeClass_(data.subject_class);
  const duplicate = getAllDataByClass_().subjects.some(function (s) {
    return s._id !== data.id && String(s.subject_name) === String(data.subject_name) && normalizeClass_(s.subject_class) === cls;
  });
  if (duplicate) throw new Error('วิชาซ้ำในชั้นนี้');
  const destName = classSheetName_('Subjects', cls);
  if (loc.sheetName === destName) {
    loc.sheet.getRange(loc.row, 1, 1, 3).setValues([[data.subject_name, cls, row[2]]]);
  } else {
    ensureClassSheet__('Subjects', cls).appendRow([data.subject_name, cls, row[2] || new Date()]);
    loc.sheet.deleteRow(loc.row);
  }
  return { ok: true };
}

function deleteByClassId_(id, prefix) {
  const baseName = prefix === 'STU_' ? 'Students' : prefix === 'SUB_' ? 'Subjects' : 'Attendance';
  const loc = decodeRowId_(id, prefix, baseName);
  if (loc.row > loc.sheet.getLastRow()) throw new Error('ไม่พบข้อมูล กรุณารีเฟรช');
  loc.sheet.deleteRow(loc.row);
  return { ok: true };
}

function isValidAttendanceDate_(value) {
  if (typeof value !== 'string' || !/^\d{4}-\d{2}-\d{2}$/.test(value)) return false;
  const parts = value.split('-').map(Number);
  const date = new Date(Date.UTC(parts[0], parts[1] - 1, parts[2]));
  return date.getUTCFullYear() === parts[0] && date.getUTCMonth() === parts[1] - 1 && date.getUTCDate() === parts[2];
}

function attendanceKey_(cls, date, studentId, subject) {
  return JSON.stringify([normalizeClass_(cls), date, String(studentId), String(subject)]);
}

function saveAttendanceByClass_(records, allowUpdates) {
  if (!Array.isArray(records) || records.length === 0 || records.length > 500) {
    throw new Error('รายการเช็คชื่อต้องมี 1–500 รายการ');
  }

  return withScriptLock_(function () {
    const data = getAllDataByClass_();
    const students = new Set(data.students.map(function (student) {
      return attendanceKey_(student.student_class, '', student.student_id, '');
    }));
    const subjects = new Set(data.subjects.map(function (subject) {
      return attendanceKey_(subject.subject_class, '', '', subject.subject_name);
    }));
    const seen = new Set();
    const validated = records.map(function (record) {
      if (!record || typeof record !== 'object' || Array.isArray(record)) throw new Error('รูปแบบรายการเช็คชื่อไม่ถูกต้อง');
      let cls = normalizeClass_(record.cls || '');
      const studentId = String(record.student_id == null ? '' : record.student_id).trim();
      const subject = String(record.subject == null ? '' : record.subject).trim();
      const date = record.date;
      const status = record.status;
      const notes = record.notes == null ? '' : record.notes;

      if (!studentId || studentId.length > 40) throw new Error('เลขที่นักเรียนไม่ถูกต้อง');
      if (!subject || subject.length > 120) throw new Error('รายวิชาไม่ถูกต้อง');
      if (!cls) {
        const possibleClasses = data.students.filter(function (student) {
          return String(student.student_id) === studentId && subjects.has(attendanceKey_(student.student_class, '', '', subject));
        }).map(function (student) { return normalizeClass_(student.student_class); });
        const uniqueClasses = Array.from(new Set(possibleClasses));
        if (uniqueClasses.length !== 1) throw new Error('ไม่สามารถระบุชั้นเรียนจากข้อมูลนักเรียนและวิชาได้');
        cls = uniqueClasses[0];
      }
      if (CLASS_LIST.indexOf(cls) === -1) throw new Error('ชั้นเรียนไม่ถูกต้อง');
      if (!isValidAttendanceDate_(date)) throw new Error('รูปแบบวันที่ไม่ถูกต้อง');
      if (ATTENDANCE_STATUS_LIST.indexOf(status) === -1) throw new Error('สถานะการเข้าเรียนไม่ถูกต้อง');
      if (typeof notes !== 'string' || notes.length > 200 || /^\s*[=+\-@]/.test(notes)) {
        throw new Error('หมายเหตุไม่ถูกต้องหรือยาวเกิน 200 ตัวอักษร');
      }
      if (!students.has(attendanceKey_(cls, '', studentId, ''))) throw new Error('ไม่พบนักเรียนในชั้นที่ระบุ');
      if (!subjects.has(attendanceKey_(cls, '', '', subject))) throw new Error('รายวิชาไม่ตรงกับชั้นเรียนที่ระบุ');

      const key = attendanceKey_(cls, date, studentId, subject);
      if (seen.has(key)) throw new Error('มีรายการนักเรียนซ้ำในคำขอเดียวกัน');
      seen.add(key);
      return { cls: cls, studentId: studentId, subject: subject, date: date, status: status, notes: notes, key: key };
    });

    const existingByKey = new Map();
    const unresolvedLegacyKeys = new Set();
    data.attendance.forEach(function (row) {
      const cls = normalizeClass_(row.attendance_student_class);
      if (!cls) {
        unresolvedLegacyKeys.add(JSON.stringify([
          normalizeDate_(row.attendance_date), String(row.attendance_student_id), String(row.attendance_subject)
        ]));
        return;
      }
      const key = attendanceKey_(cls, row.attendance_date, row.attendance_student_id, row.attendance_subject);
      existingByKey.set(key, row);
    });

    const insertsByClass = {};
    const updates = [];
    let inserted = 0, updated = 0, skipped = 0;
    validated.forEach(function (record) {
      const match = existingByKey.get(record.key);
      if (match) {
        if (!allowUpdates) {
          if (match.attendance_status !== record.status || String(match.attendance_notes || '') !== record.notes) {
            throw new Error('รายการนี้ถูกบันทึกแล้ว หากต้องการแก้ไขสถานะ กรุณาเข้าสู่ระบบ');
          }
          skipped++;
          return;
        }
        updates.push({ match: match, record: record });
        updated++;
      } else {
        const legacyKey = JSON.stringify([record.date, record.studentId, record.subject]);
        if (unresolvedLegacyKeys.has(legacyKey)) {
          throw new Error('พบรายการเดิมที่ยังระบุชั้นเรียนไม่ได้ กรุณาให้ผู้ดูแลตรวจสอบก่อนบันทึกซ้ำ');
        }
        if (!insertsByClass[record.cls]) insertsByClass[record.cls] = [];
        insertsByClass[record.cls].push([
          record.date, record.studentId, record.subject, record.status,
          record.notes, new Date(), record.cls
        ]);
        inserted++;
      }
    });

    // Do not mutate any sheet until every incoming row has passed validation.
    // Resolve every target row first, so a stale/missing _id cannot leave a
    // partially applied batch behind.
    const preparedUpdates = updates.map(function (item) {
      return { loc: decodeRowId_(item.match._id, 'ATT_', 'Attendance'), item: item };
    });
    preparedUpdates.forEach(function (prepared) {
      const loc = prepared.loc;
      const item = prepared.item;
      loc.sheet.getRange(loc.row, 4, 1, 2).setValues([[item.record.status, item.record.notes]]);
      if (!item.match.attendance_student_class) {
        if (loc.sheet.getLastColumn() < 7) loc.sheet.getRange(1, 7).setValue('attendance_student_class');
        loc.sheet.getRange(loc.row, 7).setValue(item.record.cls);
      }
    });
    Object.keys(insertsByClass).forEach(function (cls) {
      const sh = ensureClassSheet__('Attendance', cls);
      const rows = insertsByClass[cls];
      sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    });
    return { ok: true, inserted: inserted, updated: updated, skipped: skipped };
  });
}

function updateAttendanceByClass_(data) {
  const loc = decodeRowId_(data.id, 'ATT_', 'Attendance');
  if (loc.row > loc.sheet.getLastRow()) throw new Error('ไม่พบข้อมูลรายการ กรุณารีเฟรช');
  loc.sheet.getRange(loc.row, 4, 1, 2).setValues([[data.attendance_status, data.attendance_notes || '']]);
  return { ok: true };
}

// One-time safe migration: copies data to per-class tabs and archives legacy tabs
// only if every source row was assigned. Unassigned rows remain in the original tab.
function migrateExistingDataToClassTabs() {
  const ss = getSpreadsheet_();
  const report = { students: 0, subjects: 0, attendance: 0, unresolved: [] };
  const studentClasses = {};
  const subjectClasses = {};
  let canArchive = { Students: true, Subjects: true, Attendance: true };

  const studentsSheet = ss.getSheetByName('Students');
  if (studentsSheet && studentsSheet.getLastRow() > 1) {
    studentsSheet.getDataRange().getValues().slice(1).forEach(function (r) {
      if (!r || String(r[0]) === '') return;
      const cls = normalizeClass_(r[2]);
      if (CLASS_LIST.indexOf(cls) === -1) { canArchive.Students = false; report.unresolved.push('นักเรียน ' + r[0] + ': ชั้นไม่ถูกต้อง'); return; }
      const dest = ensureClassSheet__('Students', cls);
      if (appendUnique_(dest, [r[0], r[1], cls, r[3]], [0])) report.students++;
      const key = String(r[0]);
      if (!studentClasses[key]) studentClasses[key] = [];
      if (studentClasses[key].indexOf(cls) === -1) studentClasses[key].push(cls);
    });
  }
  // Include already split students when resolving attendance rows.
  dataSheetEntries_('Students').filter(function (x) { return x.name !== 'Students'; }).forEach(function (entry) {
    entry.sheet.getDataRange().getValues().slice(1).forEach(function (r) {
      if (!r || String(r[0]) === '') return;
      const cls = normalizeClass_(r[2] || entry.cls);
      const key = String(r[0]);
      if (!studentClasses[key]) studentClasses[key] = [];
      if (studentClasses[key].indexOf(cls) === -1) studentClasses[key].push(cls);
    });
  });

  const subjectsSheet = ss.getSheetByName('Subjects');
  if (subjectsSheet && subjectsSheet.getLastRow() > 1) {
    subjectsSheet.getDataRange().getValues().slice(1).forEach(function (r) {
      if (!r || String(r[0]) === '') return;
      const cls = normalizeClass_(r[1]);
      if (CLASS_LIST.indexOf(cls) === -1) { canArchive.Subjects = false; report.unresolved.push('วิชา ' + r[0] + ': ชั้นไม่ถูกต้อง'); return; }
      const dest = ensureClassSheet__('Subjects', cls);
      if (appendUnique_(dest, [r[0], cls, r[2]], [0])) report.subjects++;
      const key = String(r[0]);
      if (!subjectClasses[key]) subjectClasses[key] = [];
      if (subjectClasses[key].indexOf(cls) === -1) subjectClasses[key].push(cls);
    });
  }
  dataSheetEntries_('Subjects').filter(function (x) { return x.name !== 'Subjects'; }).forEach(function (entry) {
    entry.sheet.getDataRange().getValues().slice(1).forEach(function (r) {
      if (!r || String(r[0]) === '') return;
      const cls = normalizeClass_(r[1] || entry.cls);
      const key = String(r[0]);
      if (!subjectClasses[key]) subjectClasses[key] = [];
      if (subjectClasses[key].indexOf(cls) === -1) subjectClasses[key].push(cls);
    });
  });

  const attendanceSheet = ss.getSheetByName('Attendance');
  if (attendanceSheet && attendanceSheet.getLastRow() > 1) {
    attendanceSheet.getDataRange().getValues().slice(1).forEach(function (r) {
      if (!r || !r[0]) return;
      let cls = normalizeClass_(r[6] || '');
      if (!cls) {
        const sc = studentClasses[String(r[1])] || [];
        const sub = subjectClasses[String(r[2])] || [];
        const candidates = sc.filter(function (candidate) { return sub.indexOf(candidate) !== -1; });
        if (candidates.length === 1) cls = candidates[0];
        else if (sc.length === 1) cls = sc[0];
      }
      if (CLASS_LIST.indexOf(cls) === -1) {
        canArchive.Attendance = false;
        report.unresolved.push('เช็คชื่อ ' + normalizeDate_(r[0]) + ' เลขที่ ' + r[1] + ' วิชา ' + r[2] + ': ระบุชั้นไม่ได้');
        return;
      }
      const dest = ensureClassSheet__('Attendance', cls);
      const dataRow = [r[0], r[1], r[2], r[3], r[4] || '', r[5] || new Date(), cls];
      if (appendUnique_(dest, dataRow, [0, 1, 2])) report.attendance++;
    });
  }

  ['Students', 'Subjects', 'Attendance'].forEach(function (name) {
    const source = ss.getSheetByName(name);
    const mayArchive = name === 'Attendance' ? canArchive.Attendance : canArchive[name];
    if (source && mayArchive) {
      source.setName(name + ' (Archive)');
      Logger.log('เก็บแท็บเดิมเป็นสำรอง: ' + name + ' (Archive)');
    }
  });
  report.unresolvedCount = report.unresolved.length;
  Logger.log(JSON.stringify(report));
  return report;
}
