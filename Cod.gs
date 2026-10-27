// ==================================================
// CONFIG & SETUP
// ==================================================
const PROP_KEY_SHEET_ID    = 'SHEET_ID';
const PROP_KEY_ADMIN_HASH  = 'ADMIN_PASSWORD_HASH';
const PROP_KEY_API_TOKEN   = 'API_TOKEN';
const PROP_KEY_SCHEMA_VER  = 'SCHEMA_VERSION';
const CURRENT_SCHEMA_VER   = '2';

const SHEET_STUDENTS   = 'Students';
const SHEET_SUBJECTS   = 'Subjects';
const SHEET_ATTENDANCE = 'Attendance';

// Schema ใหม่
const HEADERS = {
  Students:   ['student_id','student_name','student_class','created_at','uid'],
  Subjects:   ['subject_name','subject_class','created_at','uid'],
  Attendance: ['attendance_date','attendance_student_id','attendance_student_class',
               'attendance_subject','attendance_status','attendance_notes','created_at','uid']
};

function setup() {
  const props = PropertiesService.getScriptProperties();
  props.setProperty(PROP_KEY_SHEET_ID, '1ya65apZf-4ODvqMFdswWsd0haSVoSvow0Xew633zF1w');
  props.setProperty(PROP_KEY_ADMIN_HASH, hashPassword_('admin1234'));
  if (!props.getProperty(PROP_KEY_API_TOKEN)) {
    props.setProperty(PROP_KEY_API_TOKEN, Utilities.getUuid());
  }
  ensureSheets_();
  props.setProperty(PROP_KEY_SCHEMA_VER, CURRENT_SCHEMA_VER);
  Logger.log('✅ Setup complete');
  Logger.log('Admin password: admin1234');
  Logger.log('API Token: ' + props.getProperty(PROP_KEY_API_TOKEN));
  Logger.log('👉 คัดลอก API Token ไปใส่ใน index.html (const API_TOKEN)');
}

function ensureSheets_() {
  const ss = getSpreadsheet_();
  Object.keys(HEADERS).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    const headers = HEADERS[name];
    const lastCol = sh.getLastColumn();
    const current = lastCol > 0 ? sh.getRange(1, 1, 1, lastCol).getValues()[0] : [];
    if (current.length < headers.length || String(current[0]) !== headers[0]) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
  });
}

function getSheetId_() {
  const id = PropertiesService.getScriptProperties().getProperty(PROP_KEY_SHEET_ID);
  if (!id) throw new Error('ยังไม่ได้รัน setup() หรือ SHEET_ID หาย');
  return id;
}
function getApiToken_() {
  return PropertiesService.getScriptProperties().getProperty(PROP_KEY_API_TOKEN) || '';
}
function getSpreadsheet_() { return SpreadsheetApp.openById(getSheetId_()); }

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
  const stored = PropertiesService.getScriptProperties().getProperty(PROP_KEY_ADMIN_HASH);
  return stored && password && stored === hashPassword_(password);
}
function normalizeDate_(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Bangkok', 'yyyy-MM-dd');
  const s = String(v).trim();
  // ISO date?
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
  return s;
}
function newUid_() { return Utilities.getUuid(); }

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    return fn();
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// ==================================================
// WEB APP ENTRY
// ==================================================
function doGet(e) {
  if (e && e.parameter && e.parameter.action) {
    try {
      const action = e.parameter.action;
      const data   = e.parameter.data ? JSON.parse(e.parameter.data) : null;
      return jsonOut_({ ok: true, data: handleAction_(action, data) });
    } catch (err) {
      return jsonOut_({ ok: false, error: err.message });
    }
  }
  return ContentService.createTextOutput('API is running');
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonOut_({ ok: false, error: 'Missing request body' });
    }
    const payload = JSON.parse(e.postData.contents);
    const result = handleAction_(payload.action, payload.data);
    return jsonOut_({ ok: true, data: result });
  } catch (err) {
    return jsonOut_({ ok: false, error: err.message });
  }
}

// ==================================================
// ROUTER + AUTH
// ==================================================
const PUBLIC_ACTIONS = ['getAllData', 'ping'];

function handleAction_(action, data) {
  // Auth (ยกเว้น public actions)
  if (!PUBLIC_ACTIONS.includes(action)) {
    const expected = getApiToken_();
    const provided = data && data.apiToken;
    if (!expected) throw new Error('ยังไม่ได้ตั้งค่า API_TOKEN — รัน setup() ก่อน');
    if (provided !== expected) throw new Error('Unauthorized: API Token ไม่ถูกต้อง');
  }

  switch (action) {
    case 'ping':                return { ok: true, ts: new Date().toISOString() };
    case 'getAllData':          return getAllData_();
    case 'getConfig':           return getConfig_();
    case 'addStudent':          return withLock_(() => addStudent_(data));
    case 'addBulkStudents':     return withLock_(() => addBulkStudents_(data));
    case 'updateStudent':       return withLock_(() => updateStudent_(data));
    case 'deleteStudent':       return withLock_(() => deleteStudent_(data.id));
    case 'addSubject':          return withLock_(() => addSubject_(data));
    case 'addBulkSubjects':     return withLock_(() => addBulkSubjects_(data));
    case 'updateSubject':       return withLock_(() => updateSubject_(data));
    case 'deleteSubject':       return withLock_(() => deleteSubject_(data.id));
    case 'saveAttendance':      return withLock_(() => saveAttendance_(data));
    case 'updateAttendance':    return withLock_(() => updateAttendance_(data));
    case 'deleteAttendance':    return withLock_(() => deleteAttendance_(data.id));

    case 'updateSheetId': {
      if (!verifyAdmin_(data.password)) throw new Error('รหัสผ่าน Admin ไม่ถูกต้อง');
      // rate limit
      const cache = CacheService.getScriptCache();
      const key = 'upd_sheet_attempts';
      const n = parseInt(cache.get(key) || '0', 10);
      if (n >= 5) throw new Error('พยายามบ่อยเกินไป ลองใหม่ภายหลัง');
      cache.put(key, String(n + 1), 300);
      PropertiesService.getScriptProperties().setProperty(PROP_KEY_SHEET_ID, data.sheetId);
      ensureSheets_();
      return { ok: true, message: 'บันทึก Sheet ID ใหม่แล้ว', sheetId: data.sheetId };
    }

    case 'changeAdminPassword': {
      if (!verifyAdmin_(data.oldPassword)) throw new Error('รหัสผ่านเดิมไม่ถูกต้อง');
      if (!data.newPassword || data.newPassword.length < 6) {
        throw new Error('รหัสใหม่ต้องอย่างน้อย 6 ตัวอักษร');
      }
      PropertiesService.getScriptProperties()
        .setProperty(PROP_KEY_ADMIN_HASH, hashPassword_(data.newPassword));
      return { ok: true, message: 'เปลี่ยนรหัสผ่านแล้ว' };
    }

    default: throw new Error('ไม่รู้จัก action: ' + action);
  }
}

function getConfig_() {
  const p = PropertiesService.getScriptProperties();
  return {
    sheetId: p.getProperty(PROP_KEY_SHEET_ID),
    schemaVersion: p.getProperty(PROP_KEY_SCHEMA_VER)
  };
}

// ==================================================
// MIGRATION (รันครั้งเดียวต่อ schema)
// ==================================================
function migrateIfNeeded_() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty(PROP_KEY_SCHEMA_VER) === CURRENT_SCHEMA_VER) return;
  withLock_(() => {
    // re-check หลังได้ lock
    if (props.getProperty(PROP_KEY_SCHEMA_VER) === CURRENT_SCHEMA_VER) return;

    const ss = getSpreadsheet_();

    // ---- Students: เพิ่ม uid ถ้ายังไม่มี ----
    const sSh = ss.getSheetByName(SHEET_STUDENTS);
    if (sSh && sSh.getLastRow() > 1) {
      const data = sSh.getDataRange().getValues();
      const header = data[0];
      const hasUid = String(header[4] || '').toLowerCase() === 'uid';
      if (!hasUid) {
        const rows = data.slice(1);
        const out = [HEADERS.Students];
        rows.forEach(r => {
          out.push([r[0], r[1], r[2], r[3] || new Date(), newUid_()]);
        });
        sSh.clearContents();
        sSh.getRange(1, 1, out.length, 5).setValues(out);
      }
    }

    // ---- Subjects ----
    const subSh = ss.getSheetByName(SHEET_SUBJECTS);
    if (subSh && subSh.getLastRow() > 1) {
      const data = subSh.getDataRange().getValues();
      const header = data[0];
      const hasUid = String(header[3] || '').toLowerCase() === 'uid';
      if (!hasUid) {
        const rows = data.slice(1);
        const out = [HEADERS.Subjects];
        rows.forEach(r => {
          out.push([r[0], r[1], r[2] || new Date(), newUid_()]);
        });
        subSh.clearContents();
        subSh.getRange(1, 1, out.length, 4).setValues(out);
      }
    }

    // ---- Attendance: เพิ่ม student_class + uid ----
    const aSh = ss.getSheetByName(SHEET_ATTENDANCE);
    if (aSh && aSh.getLastRow() > 1) {
      const data = aSh.getDataRange().getValues();
      // ตรวจว่า migrate แล้วหรือยัง
      const isMigrated = data[0].length >= 8 && String(data[0][2]).toLowerCase() === 'attendance_student_class';
      if (!isMigrated) {
        // map: student_id -> class (ถ้าซ้ำ เลือกตัวแรก)
        const stuMap = new Map();
        if (sSh && sSh.getLastRow() > 1) {
          sSh.getDataRange().getValues().slice(1).forEach(r => {
            const sid = String(r[0]);
            if (!stuMap.has(sid)) stuMap.set(sid, String(r[2]));
          });
        }
        const out = [HEADERS.Attendance];
        data.slice(1).forEach(r => {
          if (!r[0]) return;
          const sid = String(r[1]);
          const cls = stuMap.get(sid) || '';
          out.push([
            normalizeDate_(r[0]),
            sid,
            cls,
            String(r[2] || ''),         // subject
            String(r[3] || ''),         // status
            String(r[4] || ''),         // notes
            r[5] || new Date(),         // created_at
            newUid_()
          ]);
        });
        aSh.clearContents();
        aSh.getRange(1, 1, out.length, 8).setValues(out);
      }
    }

    props.setProperty(PROP_KEY_SCHEMA_VER, CURRENT_SCHEMA_VER);
  });
}

// ==================================================
// GET ALL DATA
// ==================================================
function getAllData_() {
  migrateIfNeeded_();
  const ss = getSpreadsheet_();
  const result = { students: [], subjects: [], attendance: [] };

  const sSh = ss.getSheetByName(SHEET_STUDENTS);
  if (sSh && sSh.getLastRow() > 1) {
    const rows = sSh.getDataRange().getValues().slice(1);
    result.students = rows.filter(r => r[0] !== '').map(r => ({
      _id: String(r[4] || newUid_()),
      student_id: String(r[0]),
      student_name: String(r[1]),
      student_class: String(r[2]),
      created_at: r[3] ? String(r[3]) : ''
    }));
  }

  const subSh = ss.getSheetByName(SHEET_SUBJECTS);
  if (subSh && subSh.getLastRow() > 1) {
    const rows = subSh.getDataRange().getValues().slice(1);
    result.subjects = rows.filter(r => r[0] !== '').map(r => ({
      _id: String(r[3] || newUid_()),
      subject_name: String(r[0]),
      subject_class: String(r[1]),
      created_at: r[2] ? String(r[2]) : ''
    }));
  }

  const aSh = ss.getSheetByName(SHEET_ATTENDANCE);
  if (aSh && aSh.getLastRow() > 1) {
    const rows = aSh.getDataRange().getValues().slice(1);
    result.attendance = rows.filter(r => r[0] !== '').map(r => ({
      _id: String(r[7] || newUid_()),
      attendance_date: normalizeDate_(r[0]),
      attendance_student_id: String(r[1]),
      attendance_student_class: String(r[2] || ''),
      attendance_subject: String(r[3] || ''),
      attendance_status: String(r[4] || ''),
      attendance_notes: String(r[5] || ''),
      created_at: r[6] ? String(r[6]) : ''
    }));
  }
  return result;
}

// ==================================================
// HELPERS: row lookup by UID
// ==================================================
function findRowByUid_(sheet, uid, uidColIndex) {
  if (!uid) return -1;
  const last = sheet.getLastRow();
  if (last < 2) return -1;
  const vals = sheet.getRange(2, uidColIndex, last - 1, 1).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]) === String(uid)) return i + 2;
  }
  return -1;
}

// ==================================================
// CRUD: STUDENTS
// ==================================================
function addStudent_(data) {
  const sh = getSpreadsheet_().getSheetByName(SHEET_STUDENTS);
  const existing = sh.getDataRange().getValues().slice(1);
  const dup = existing.some(r =>
    String(r[0]) === String(data.student_id) &&
    String(r[2]) === String(data.student_class)
  );
  if (dup) throw new Error('มีนักเรียนเลขที่นี้ในชั้นนี้แล้ว');
  sh.appendRow([data.student_id, data.student_name, data.student_class, new Date(), newUid_()]);
  return { ok: true };
}

function addBulkStudents_(data) {
  const sh = getSpreadsheet_().getSheetByName(SHEET_STUDENTS);
  const existing = new Set(
    sh.getDataRange().getValues().slice(1).map(r => String(r[0]) + '|' + String(r[2]))
  );
  const rows = [];
  let added = 0, skipped = 0;
  data.list.forEach(s => {
    const key = String(s.student_id) + '|' + String(s.student_class);
    if (existing.has(key)) { skipped++; return; }
    rows.push([s.student_id, s.student_name, s.student_class, new Date(), newUid_()]);
    existing.add(key);
    added++;
  });
  if (rows.length) {
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, 5).setValues(rows);
  }
  return { ok: true, added, skipped };
}

function updateStudent_(data) {
  const ss = getSpreadsheet_();
  const sh = ss.getSheetByName(SHEET_STUDENTS);
  const rowIndex = findRowByUid_(sh, data.id, 5);
  if (rowIndex < 0) throw new Error('ไม่พบนักเรียน (UID ไม่ถูกต้อง)');

  const all = sh.getDataRange().getValues();
  const dup = all.slice(1).some((r, i) =>
    (i + 2) !== rowIndex &&
    String(r[0]) === String(data.student_id) &&
    String(r[2]) === String(data.student_class)
  );
  if (dup) throw new Error('เลขที่ซ้ำในชั้นนี้');

  // อ่านค่าเดิม
  const oldRow = sh.getRange(rowIndex, 1, 1, 3).getValues()[0];
  const oldId = String(oldRow[0]);
  const oldClass = String(oldRow[2]);
  const newId = String(data.student_id);
  const newClass = String(data.student_class);

  sh.getRange(rowIndex, 1, 1, 3).setValues([[data.student_id, data.student_name, data.student_class]]);

  // Cascade update attendance ถ้า id หรือ class เปลี่ยน
  if (oldId !== newId || oldClass !== newClass) {
    const aSh = ss.getSheetByName(SHEET_ATTENDANCE);
    const last = aSh.getLastRow();
    if (last > 1) {
      const vals = aSh.getRange(2, 1, last - 1, 8).getValues();
      let changed = 0;
      for (let i = 0; i < vals.length; i++) {
        if (String(vals[i][1]) === oldId && String(vals[i][2]) === oldClass) {
          vals[i][1] = newId;
          vals[i][2] = newClass;
          changed++;
        }
      }
      if (changed > 0) {
        aSh.getRange(2, 1, vals.length, 8).setValues(vals);
      }
    }
  }
  return { ok: true };
}

function deleteStudent_(uid) {
  const ss = getSpreadsheet_();
  const sh = ss.getSheetByName(SHEET_STUDENTS);
  const rowIndex = findRowByUid_(sh, uid, 5);
  if (rowIndex < 0) throw new Error('ไม่พบนักเรียน (UID ไม่ถูกต้อง)');
  sh.deleteRow(rowIndex);

  // ลบ attendance ของนักเรียนคนนี้ด้วย
  const aSh = ss.getSheetByName(SHEET_ATTENDANCE);
  const last = aSh.getLastRow();
  if (last > 1) {
    const vals = aSh.getRange(2, 1, last - 1, 8).getValues();
    const kept = vals.filter(r => String(r[7]) !== String(uid));
    if (kept.length !== vals.length) {
      aSh.getRange(2, 1, vals.length, 8).clearContent();
      if (kept.length) aSh.getRange(2, 1, kept.length, 8).setValues(kept);
    }
  }
  return { ok: true };
}

// ==================================================
// CRUD: SUBJECTS
// ==================================================
function addSubject_(data) {
  const sh = getSpreadsheet_().getSheetByName(SHEET_SUBJECTS);
  const existing = sh.getDataRange().getValues().slice(1);
  if (existing.some(r =>
    String(r[0]) === String(data.subject_name) &&
    String(r[1]) === String(data.subject_class))) {
    throw new Error('มีวิชานี้ในชั้นนี้แล้ว');
  }
  sh.appendRow([data.subject_name, data.subject_class, new Date(), newUid_()]);
  return { ok: true };
}

function addBulkSubjects_(data) {
  const sh = getSpreadsheet_().getSheetByName(SHEET_SUBJECTS);
  const existing = new Set(
    sh.getDataRange().getValues().slice(1).map(r => String(r[0]) + '|' + String(r[1]))
  );
  const rows = [];
  let added = 0, skipped = 0;
  data.list.forEach(name => {
    const key = name + '|' + data.subject_class;
    if (existing.has(key)) { skipped++; return; }
    rows.push([name, data.subject_class, new Date(), newUid_()]);
    existing.add(key);
    added++;
  });
  if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, 4).setValues(rows);
  return { ok: true, added, skipped };
}

function updateSubject_(data) {
  const sh = getSpreadsheet_().getSheetByName(SHEET_SUBJECTS);
  const rowIndex = findRowByUid_(sh, data.id, 4);
  if (rowIndex < 0) throw new Error('ไม่พบวิชา (UID ไม่ถูกต้อง)');

  const all = sh.getDataRange().getValues();
  const dup = all.slice(1).some((r, i) =>
    (i + 2) !== rowIndex &&
    String(r[0]) === String(data.subject_name) &&
    String(r[1]) === String(data.subject_class)
  );
  if (dup) throw new Error('วิชาซ้ำในชั้นนี้');

  sh.getRange(rowIndex, 1, 1, 2).setValues([[data.subject_name, data.subject_class]]);
  return { ok: true };
}

function deleteSubject_(uid) {
  const sh = getSpreadsheet_().getSheetByName(SHEET_SUBJECTS);
  const rowIndex = findRowByUid_(sh, uid, 4);
  if (rowIndex < 0) throw new Error('ไม่พบวิชา');
  sh.deleteRow(rowIndex);
  return { ok: true };
}

// ==================================================
// CRUD: ATTENDANCE
// ==================================================
function saveAttendance_(records) {
  const sh = getSpreadsheet_().getSheetByName(SHEET_ATTENDANCE);
  const all = sh.getDataRange().getValues();
  const keyIndex = new Map(); // date|student_id|class|subject → rowIndex
  all.slice(1).forEach((r, i) => {
    const k = normalizeDate_(r[0]) + '|' + String(r[1]) + '|' + String(r[2]) + '|' + String(r[3]);
    keyIndex.set(k, i + 2);
  });

  let updated = 0, inserted = 0;
  const toAppend = [];

  records.forEach(rec => {
    const key = rec.date + '|' + String(rec.student_id) + '|' + String(rec.student_class) + '|' + rec.subject;
    if (keyIndex.has(key)) {
      const row = keyIndex.get(key);
      sh.getRange(row, 5, 1, 2).setValues([[rec.status, rec.notes || '']]);
      updated++;
    } else {
      toAppend.push([
        rec.date, String(rec.student_id), String(rec.student_class),
        rec.subject, rec.status, rec.notes || '', new Date(), newUid_()
      ]);
    }
  });

  if (toAppend.length) {
    sh.getRange(sh.getLastRow() + 1, 1, toAppend.length, 8).setValues(toAppend);
    inserted = toAppend.length;
  }
  return { ok: true, updated, inserted };
}

function updateAttendance_(data) {
  const sh = getSpreadsheet_().getSheetByName(SHEET_ATTENDANCE);
  const rowIndex = findRowByUid_(sh, data.id, 8);
  if (rowIndex < 0) throw new Error('ไม่พบรายการ (UID ไม่ถูกต้อง)');
  sh.getRange(rowIndex, 5, 1, 2).setValues([[data.attendance_status, data.attendance_notes || '']]);
  return { ok: true };
}

function deleteAttendance_(uid) {
  const sh = getSpreadsheet_().getSheetByName(SHEET_ATTENDANCE);
  const rowIndex = findRowByUid_(sh, uid, 8);
  if (rowIndex < 0) throw new Error('ไม่พบรายการ');
  sh.deleteRow(rowIndex);
  return { ok: true };
}
