// ==================================================
// CONFIG & SETUP
// ==================================================
const PROP_KEY_SHEET_ID = 'SHEET_ID';
const PROP_KEY_ADMIN_HASH = 'ADMIN_PASSWORD_HASH';
const PROP_KEY_LOGIN_USERNAME = 'APP_LOGIN_USERNAME';
const PROP_KEY_LOGIN_PASSWORD_HASH = 'APP_LOGIN_PASSWORD_HASH';
const PROP_KEY_ADMIN_USERS = 'APP_ADMIN_USERS';
const MIN_ADMIN_PASSWORD_DIGITS = 4;
const MIN_ADMIN_PASSWORD_LETTERS = 2;
const INITIAL_LOGIN_USERNAME = 'FirstStar';
const INITIAL_LOGIN_PASSWORD_HASH = '4bbe438f18a5e2b59bdf0472eac2a86a957c4e874b002d81f94867692703c646';
const AUTH_SESSION_TTL_SECONDS = 21600; // 6 hours (Apps Script CacheService maximum)
const AUTH_LOGIN_FAILURE_KEY = 'APP_AUTH_LOGIN_FAILURES';
const AUTH_LOGIN_FAILURE_LIMIT = 5;
const AUTH_LOGIN_LOCK_SECONDS = 900;
// เวอร์ชันโค้ดฝั่งเซิร์ฟเวอร์ — ใช้ตรวจว่า deployment ที่ใช้งานอยู่อัปเดตแล้วหรือยัง
const APP_CODE_VERSION = '2026-11-19-edit-attendance';
const CLASS_LIST = ['ป.1', 'ป.2', 'ป.3', 'ป.4', 'ป.5', 'ป.6', 'ม.1', 'ม.2', 'ม.3'];
const ATTENDANCE_STATUS_LIST = ['มา', 'สาย', 'ลา', 'ขาด'];

// action ที่เปลี่ยนข้อมูลจริง → ต้องล้างแคชหลังทำงาน (อ่านอย่างเดียวไม่ต้องล้าง)
const MUTATING_ACTIONS = new Set([
  'changePassword', 'addAdmin', 'removeAdmin',
  'addStudent', 'addBulkStudents', 'updateStudent', 'deleteStudent',
  'addSubject', 'addBulkSubjects', 'updateSubject', 'deleteSubject',
  'updateAttendance', 'deleteAttendance', 'updateSheetId'
]);

// ช่วงเวลาเช็คชื่อประจำวัน — ต้องตรงกับ PERIOD_DEFS ใน index.html
const ATTENDANCE_PERIOD_LIST = ['07.30-08.10', '08.15-08.20', '08.30-09.00', '09.00-10.00',
  '10.00-11.30', '12.30-13.00', '13.00-14.00', '14.00-15.00', '15.00-15.30'];
// กิจกรรมประจำช่วงเวลา (ไม่ต้องมีในชีต Subjects แต่บันทึกเป็นวิชาได้)
const ATTENDANCE_ACTIVITY_LIST = ['ปรนนิบัติสถานที่', 'สติเคารพฯ', 'MT (Mind training)',
  'PT (Personality talk)', 'AAR'];

// โครงตารางของแต่ละชีต (ใช้ตอนสร้างชีตใหม่เท่านั้น)
const SHEET_DEFS = {
  'Students':   ['student_id', 'student_name', 'student_class', 'created_at'],
  'Subjects':   ['subject_name', 'subject_class', 'created_at'],
  'Attendance': ['attendance_date', 'attendance_student_id', 'attendance_subject',
                 'attendance_status', 'attendance_notes', 'created_at']
};

// หัวตารางชีต Attendance ครบชุด = 6 คอลัมน์หลัก + ชั้นเรียน (7) + ช่วงเวลาเช็คชื่อ (8)
function attendanceHeaders_() {
  return SHEET_DEFS.Attendance.concat(['attendance_student_class', 'attendance_period']);
}

// เติมหัวคอลัมน์ที่ยังขาด (เช่นเพิ่ม attendance_period ให้ชีตเดิม) โดยไม่ทับข้อมูลเดิม
function ensureHeaderTail_(sh, headers) {
  const last = sh.getLastColumn();
  if (last >= headers.length) return;
  const missing = headers.slice(last);
  sh.getRange(1, last + 1, 1, missing.length).setValues([missing]);
}

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

  // บัญชีหลักของระบบ: ตั้งครั้งแรกเท่านั้น (รัน setup() ซ้ำจะไม่ทับรหัสที่เปลี่ยนแล้ว)
  if (!props.getProperty(PROP_KEY_LOGIN_USERNAME)) {
    props.setProperty(PROP_KEY_LOGIN_USERNAME, INITIAL_LOGIN_USERNAME);
  }
  if (!props.getProperty(PROP_KEY_LOGIN_PASSWORD_HASH)) {
    props.setProperty(PROP_KEY_LOGIN_PASSWORD_HASH, INITIAL_LOGIN_PASSWORD_HASH);
  }
  Logger.log('✅ Setup: บัญชีหลักของระบบพร้อมใช้งาน (' + ownerUsername_() + ')');

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
      const headers = kind === 'Attendance' ? attendanceHeaders_() : SHEET_DEFS[kind];
      if (sh.getLastRow() === 0) sh.getRange(1, 1, 1, headers.length).setValues([headers]);
      else ensureHeaderTail_(sh, headers);
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

// เปิด Spreadsheet ครั้งเดียวต่อ 1 คำขอ (เดิมเปิดซ้ำหลายสิบครั้ง = ช้ามาก)
let SS_MEMO_ = null;
function getSpreadsheet_() {
  if (!SS_MEMO_) SS_MEMO_ = SpreadsheetApp.openById(getSheetId_());
  return SS_MEMO_;
}

// ---------- แคชข้อมูลทั้งหมด (แบ่งเป็นก้อน ก้อนละ ≤ 25,000 ตัวอักษร) ----------
const DATA_CACHE_KEY = 'APP_DATA_CACHE_V1';
const DATA_CACHE_TTL = 600;   // วินาที (มี trigger รีเฟรชทุก 5 นาที ถ้าติดตั้งด้วย installCacheRefreshTrigger)
const DATA_CACHE_CHUNK = 25000;      // อักษรไทย 3 ไบต์ → ~75KB/ก้อน ต่ำกว่าเพดาน 100KB/คีย์
const DATA_CACHE_MAX_CHUNKS = 400;   // ~10M ตัวอักษร (~40,000 แถว) — กันแคชหยุดเงียบเมื่อข้อมูลโต
const DATA_CACHE_GEN_KEY = DATA_CACHE_KEY + '_gen';
const DATA_CACHE_PUT_BATCH = 20;

// generation ปัจจุบันของแคช — เปลี่ยนทุกครั้งที่ข้อมูลถูกแก้
function cacheGeneration_() {
  try { return CacheService.getScriptCache().get(DATA_CACHE_GEN_KEY) || ''; } catch (err) { return ''; }
}

function cacheGetData_() {
  try {
    const c = CacheService.getScriptCache();
    const gen = c.get(DATA_CACHE_GEN_KEY);
    if (!gen) return null;                 // ไม่มี generation = ยังไม่เคยแคช หรือถูก invalidation
    const n = parseInt(c.get(DATA_CACHE_KEY + '_' + gen + '_n') || '0', 10);
    if (!n || n > DATA_CACHE_MAX_CHUNKS) return null;
    const keys = [];
    for (let i = 0; i < n; i++) keys.push(DATA_CACHE_KEY + '_' + gen + '_' + i);
    const parts = c.getAll(keys);
    let s = '';
    for (let i = 0; i < n; i++) {
      const p = parts[keys[i]];
      if (p == null) return null;
      s += p;
    }
    return JSON.parse(s);
  } catch (err) { return null; }
}

// เขียนแคชชุดใหม่ลง key ของ generation ใหม่ แล้วค่อยสลับ generation (compare-and-set)
// expectedGen = generation ที่อ่านตอนเริ่มคำขอ; ถ้ามีการบันทึกแทรกกลางทาง จะไม่เขียนทับข้อมูลใหม่
function cachePutData_(obj, expectedGen) {
  try {
    const c = CacheService.getScriptCache();
    const baseGen = expectedGen === undefined ? cacheGeneration_() : String(expectedGen || '');
    if ((c.get(DATA_CACHE_GEN_KEY) || '') !== baseGen) return false; // ข้อมูลถูกแก้ระหว่างอ่าน

    const s = JSON.stringify(obj);
    const gen = Utilities.getUuid().replace(/-/g, '');
    const n = Math.ceil(s.length / DATA_CACHE_CHUNK);
    if (!n || n > DATA_CACHE_MAX_CHUNKS) {
      console.warn('cachePutData_: skipped, chunks=' + n + ' max=' + DATA_CACHE_MAX_CHUNKS);
      return false;
    }
    const keys = [DATA_CACHE_KEY + '_' + gen + '_n'];
    const values = [String(n)];
    for (let i = 0; i < n; i++) {
      keys.push(DATA_CACHE_KEY + '_' + gen + '_' + i);
      values.push(s.slice(i * DATA_CACHE_CHUNK, (i + 1) * DATA_CACHE_CHUNK));
    }
    for (let i = 0; i < keys.length; i += DATA_CACHE_PUT_BATCH) {
      const puts = {};
      const end = Math.min(i + DATA_CACHE_PUT_BATCH, keys.length);
      for (let j = i; j < end; j++) puts[keys[j]] = values[j];
      try {
        c.putAll(puts, DATA_CACHE_TTL);
      } catch (batchErr) {
        // บางสภาพแวดล้อมจำกัดขนาด putAll → เขียนทีละคีย์แทน
        Object.keys(puts).forEach(function (k) { c.put(k, puts[k], DATA_CACHE_TTL); });
      }
    }
    // ตรวจซ้ำก่อนสลับ generation: ถ้ามีคนบันทึกระหว่างเขียน ให้ทิ้งชุดนี้
    if ((c.get(DATA_CACHE_GEN_KEY) || '') !== baseGen) {
      c.removeAll(keys);
      return false;
    }
    c.put(DATA_CACHE_GEN_KEY, gen, DATA_CACHE_TTL);
    return true;
  } catch (err) {
    console.warn('cachePutData_ failed: ' + err.message);
    return false; // แคชล้มเหลวได้ ไม่กระทบการทำงาน
  }
}

// invalidation: สลับ generation ทันที (ชุดเก่าเข้าถึงไม่ได้) แล้วค่อยลบ key ชุดเก่า
function cacheClearData_() {
  try {
    const c = CacheService.getScriptCache();
    const oldGen = c.get(DATA_CACHE_GEN_KEY);
    const stamp = 'x' + Utilities.getUuid().replace(/-/g, '');
    c.put(DATA_CACHE_GEN_KEY, stamp, DATA_CACHE_TTL);
    if (!oldGen) return;
    const n = parseInt(c.get(DATA_CACHE_KEY + '_' + oldGen + '_n') || '0', 10);
    const keys = [DATA_CACHE_KEY + '_' + oldGen + '_n'];
    for (let i = 0; i < n && i < DATA_CACHE_MAX_CHUNKS; i++) {
      keys.push(DATA_CACHE_KEY + '_' + oldGen + '_' + i);
    }
    c.removeAll(keys);
  } catch (err) {}
}

function getAllDataCached_(force) {
  const t0 = Date.now();
  const startedGen = cacheGeneration_();
  if (!force) {
    const cached = cacheGetData_();
    if (cached) {
      cached._meta = { cached: true, serverMs: Date.now() - t0 };
      return cached;
    }
  }
  SHEETS_READ_COUNT_ = 0;
  const data = getAllDataByClass_();
  const cacheWritten = cachePutData_(data, startedGen);
  data._meta = {
    cached: false, serverMs: Date.now() - t0,
    sheetsRead: SHEETS_READ_COUNT_, cacheWritten: cacheWritten
  };
  return data;
}

// รันด้วย time-driven trigger ทุก 5 นาที: อุ่นสคริปต์ + เติมแคชไว้ล่วงหน้า
// ผู้ใช้จึงแทบไม่ต้องรออ่านชีตจริง (เดิมต้องรอทุกครั้งที่แคชหมดอายุ)
function refreshDataCache() {
  const gen = cacheGeneration_();
  const data = getAllDataByClass_();
  cachePutData_(data, gen);
}

// ⚠️ รันฟังก์ชันนี้ "ครั้งเดียว" ใน Apps Script เพื่อติดตั้ง trigger (รันซ้ำได้ ไม่สร้างซ้ำ)
function installCacheRefreshTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'refreshDataCache') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('refreshDataCache').timeBased().everyMinutes(5).create();
  Logger.log('✅ ติดตั้ง trigger refreshDataCache ทุก 5 นาทีแล้ว');
}

// อ่านเฉพาะแถวข้อมูลที่มีจริง (ข้ามชีตว่าง และอ่านเฉพาะคอลัมน์ที่ใช้)
let SHEETS_READ_COUNT_ = 0;
function readRows_(sh, cols) {
  SHEETS_READ_COUNT_++;
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  const width = Math.min(cols, sh.getMaxColumns());
  return sh.getRange(2, 1, lastRow - 1, width).getValues();
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

// ---------- บัญชีผู้ดูแลระบบ (เจ้าของระบบ 1 บัญชี + ผู้ดูแลที่เพิ่มได้หลายคน) ----------
// เก็บเฉพาะ hash ของรหัสผ่าน ไม่เก็บรหัสจริง
function normalizeUsername_(value) {
  return String(value == null ? '' : value).trim().slice(0, 40);
}

function ownerUsername_() {
  return PropertiesService.getScriptProperties().getProperty(PROP_KEY_LOGIN_USERNAME) || INITIAL_LOGIN_USERNAME;
}

function ownerPasswordHash_() {
  return PropertiesService.getScriptProperties().getProperty(PROP_KEY_LOGIN_PASSWORD_HASH) || INITIAL_LOGIN_PASSWORD_HASH;
}

function adminAccounts_() {
  try {
    const raw = PropertiesService.getScriptProperties().getProperty(PROP_KEY_ADMIN_USERS);
    const list = raw ? JSON.parse(raw) : [];
    if (!Array.isArray(list)) return [];
    return list.filter(function (a) { return a && a.username && a.passwordHash; });
  } catch (err) {
    return [];
  }
}

function saveAdminAccounts_(list) {
  PropertiesService.getScriptProperties().setProperty(PROP_KEY_ADMIN_USERS, JSON.stringify(list));
}

function validateAdminUsername_(username) {
  if (!/^[A-Za-z0-9._-]{3,40}$/.test(username)) {
    throw new Error('ชื่อผู้ใช้ต้องเป็นตัวอักษรอังกฤษ/ตัวเลข 3–40 ตัว (ใช้ . _ - ได้)');
  }
}

function countChars_(value, pattern) {
  const matches = String(value).match(pattern);
  return matches ? matches.length : 0;
}

// รหัสผ่านผู้ดูแลระบบ: ตัวเลขอย่างน้อย 4 ตัว และตัวอักษรอย่างน้อย 2 ตัว
function validateNewPassword_(password) {
  const value = String(password == null ? '' : password);
  const digits = countChars_(value, /[0-9]/g);
  const letters = countChars_(value, /[A-Za-z\u0E00-\u0E7F]/g);
  if (digits < MIN_ADMIN_PASSWORD_DIGITS || letters < MIN_ADMIN_PASSWORD_LETTERS) {
    throw new Error('รหัสผ่านต้องมีตัวเลขอย่างน้อย ' + MIN_ADMIN_PASSWORD_DIGITS +
      ' ตัว และตัวอักษรอย่างน้อย ' + MIN_ADMIN_PASSWORD_LETTERS + ' ตัว');
  }
  if (value.length > 100) throw new Error('รหัสผ่านยาวเกิน 100 ตัวอักษร');
  return value;
}

function isOwnerUsername_(username) {
  return normalizeUsername_(username).toLowerCase() === ownerUsername_().toLowerCase();
}

// หาบัญชีจากชื่อผู้ใช้ (ไม่สนตัวพิมพ์เล็ก/ใหญ่) — คืน null ถ้าไม่มี
function accountForUsername_(username) {
  const name = normalizeUsername_(username);
  if (!name) return null;
  if (isOwnerUsername_(name)) {
    return { username: ownerUsername_(), role: 'owner', passwordHash: ownerPasswordHash_(), updated_at: '' };
  }
  const lower = name.toLowerCase();
  const found = adminAccounts_().filter(function (a) {
    return String(a.username).toLowerCase() === lower;
  })[0];
  if (!found) return null;
  return {
    username: found.username,
    role: 'admin',
    passwordHash: found.passwordHash,
    updated_at: found.updated_at || ''
  };
}

function createAuthSession_(username, role) {
  const token = Utilities.getUuid() + Utilities.getUuid();
  const session = { username: username, role: role || 'admin' };
  CacheService.getScriptCache().put('APP_AUTH_' + token, JSON.stringify(session), AUTH_SESSION_TTL_SECONDS);
  return { token: token, username: username, role: session.role, expiresIn: AUTH_SESSION_TTL_SECONDS };
}

// อ่าน session จาก token; ถ้าบัญชีถูกลบไปแล้ว session จะใช้ไม่ได้ทันที
function getAuthSession_(token) {
  if (!token || typeof token !== 'string' || token.length > 100) return {};
  const raw = CacheService.getScriptCache().get('APP_AUTH_' + token);
  if (!raw) return {};
  let parsed = null;
  try { parsed = JSON.parse(raw); } catch (_) { parsed = { username: raw }; }
  if (!parsed || !parsed.username) return {};
  if (!accountForUsername_(parsed.username)) return {};
  return { username: parsed.username, role: parsed.role || 'admin' };
}

function getAuthUser_(token) {
  return getAuthSession_(token).username || '';
}

function changePassword_(data, token) {
  const session = getAuthSession_(token);
  if (!session.username) throw new Error('กรุณาเข้าสู่ระบบก่อนเปลี่ยนรหัสผ่าน');
  const account = accountForUsername_(session.username);
  if (!account) throw new Error('ไม่พบบัญชีผู้ใช้ในระบบ กรุณาเข้าสู่ระบบใหม่');
  const currentPassword = String(data && data.currentPassword || '');
  if (hashPassword_(currentPassword) !== account.passwordHash) {
    throw new Error('รหัสผ่านปัจจุบันไม่ถูกต้อง');
  }
  const nextPassword = validateNewPassword_(data && data.newPassword);
  const nextHash = hashPassword_(nextPassword);
  if (nextHash === account.passwordHash) throw new Error('รหัสผ่านใหม่ต้องต่างจากรหัสเดิม');
  const now = new Date().toISOString();

  if (account.role === 'owner') {
    PropertiesService.getScriptProperties().setProperty(PROP_KEY_LOGIN_PASSWORD_HASH, nextHash);
  } else {
    const lower = account.username.toLowerCase();
    saveAdminAccounts_(adminAccounts_().map(function (a) {
      if (String(a.username).toLowerCase() !== lower) return a;
      return {
        username: a.username,
        passwordHash: nextHash,
        created_at: a.created_at || '',
        updated_at: now
      };
    }));
  }
  return { ok: true, username: account.username, updated_at: now };
}

function listAdmins_(token) {
  const session = getAuthSession_(token);
  if (!session.username) throw new Error('กรุณาเข้าสู่ระบบก่อน');
  const admins = [{ username: ownerUsername_(), role: 'owner', updated_at: '' }];
  adminAccounts_().forEach(function (a) {
    admins.push({ username: a.username, role: 'admin', updated_at: a.updated_at || '' });
  });
  return {
    current: session.username,
    role: session.role || '',
    admins: admins,
    minPasswordDigits: MIN_ADMIN_PASSWORD_DIGITS,
    minPasswordLetters: MIN_ADMIN_PASSWORD_LETTERS
  };
}

function addAdminAccount_(data) {
  const username = normalizeUsername_(data && data.username);
  validateAdminUsername_(username);
  const password = validateNewPassword_(data && data.password);
  const lower = username.toLowerCase();
  const exists = isOwnerUsername_(username) || adminAccounts_().some(function (a) {
    return String(a.username).toLowerCase() === lower;
  });
  if (exists) throw new Error('มีชื่อผู้ใช้นี้ในระบบแล้ว');
  const list = adminAccounts_();
  list.push({
    username: username,
    passwordHash: hashPassword_(password),
    created_at: new Date().toISOString(),
    updated_at: ''
  });
  saveAdminAccounts_(list);
  return { ok: true, username: username };
}

function removeAdminAccount_(data, token) {
  const session = getAuthSession_(token);
  if (!session.username) throw new Error('กรุณาเข้าสู่ระบบก่อน');
  const username = normalizeUsername_(data && data.username);
  if (!username) throw new Error('ระบุชื่อผู้ใช้ที่ต้องการลบ');
  if (isOwnerUsername_(username)) throw new Error('ลบบัญชีหลักของระบบไม่ได้');
  if (username.toLowerCase() === String(session.username).toLowerCase()) {
    throw new Error('ลบบัญชีที่กำลังใช้งานอยู่ไม่ได้');
  }
  const list = adminAccounts_();
  const next = list.filter(function (a) {
    return String(a.username).toLowerCase() !== username.toLowerCase();
  });
  if (next.length === list.length) throw new Error('ไม่พบบัญชีผู้ใช้นี้');
  saveAdminAccounts_(next);
  return { ok: true, username: username };
}

function requireAuth_(token) {
  const username = getAuthUser_(token);
  if (!username) throw new Error('กรุณาเข้าสู่ระบบก่อนเพิ่มหรือแก้ไขข้อมูล');
  return username;
}

function withScriptLock_(callback) {
  const lock = LockService.getScriptLock();
  // บัฟเฟอร์คิว: รอคิวเขียนข้อมูลสูงสุด ~25 วินาที เพื่อรองรับครูหลายคนบันทึกพร้อมกัน
  const deadline = Date.now() + 25000;
  let acquired = false;
  while (!acquired) {
    const remain = deadline - Date.now();
    if (remain <= 0) throw new Error('ระบบมีคิวบันทึกข้อมูลจากผู้ใช้อื่นหนาแน่น กรุณากดบันทึกอีกครั้ง');
    try {
      lock.waitLock(Math.min(remain, 5000));
      acquired = true;
    } catch (err) {
      if (Date.now() >= deadline) {
        throw new Error('ระบบมีคิวบันทึกข้อมูลจากผู้ใช้อื่นหนาแน่น กรุณากดบันทึกอีกครั้ง');
      }
    }
  }
  try {
    return callback();
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

function login_(data) {
  const cache = CacheService.getScriptCache();
  const attemptedUsername = String(data && data.username || '').trim().slice(0, 80);
  const failureKey = AUTH_LOGIN_FAILURE_KEY + '_' + hashPassword_(attemptedUsername.toLowerCase()).slice(0, 16);
  const failedAttempts = parseInt(cache.get(failureKey) || '0', 10);
  const account = attemptedUsername ? accountForUsername_(attemptedUsername) : null;
  const credentialsMatch = !!account && !!data &&
    hashPassword_(String(data.password || '')) === account.passwordHash;

  // A valid login must recover immediately, even if earlier typos triggered the cooldown.
  if (credentialsMatch) {
    cache.remove(failureKey);
    return createAuthSession_(account.username, account.role);
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
      if (action !== 'getAllData' && action !== 'getAuthStatus' && action !== 'version') {
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
  if (action === 'getAuthStatus') {
    const session = getAuthSession_(sessionToken);
    return { authenticated: !!session.username, username: session.username || '', role: session.role || '' };
  }
  if (action === 'version') return { version: APP_CODE_VERSION, adminAccounts: true };
  if (action === 'getAllData') return getAllDataCached_(!!(data && data.force));
  // Anyone with the link may add attendance; changing an existing record requires login.
  if (action === 'saveAttendance') return saveAttendanceByClass_(data, !!getAuthUser_(sessionToken));

  // Student/subject management, attendance corrections and system settings require a valid session.
  requireAuth_(sessionToken);
  if (action === 'getSheetInfo') {
    const ss = getSpreadsheet_();
    return { sheetId: getSheetId_(), name: ss.getName(), tabs: ss.getSheets().length };
  }
  try {
    return withScriptLock_(function () {
    switch (action) {
      case 'changePassword':   return changePassword_(data, sessionToken);
      case 'listAdmins':       return listAdmins_(sessionToken);
      case 'addAdmin':         return addAdminAccount_(data);
      case 'removeAdmin':      return removeAdminAccount_(data, sessionToken);
      case 'addStudent':       return addStudentByClass_(data);
      case 'addBulkStudents':  return addBulkStudentsByClass_(data);
      case 'updateStudent':    return updateStudentByClass_(data);
      case 'deleteStudent':    return deleteByClassId_(data.id, 'STU_', data.expect);
      case 'addSubject':       return addSubjectByClass_(data);
      case 'addBulkSubjects':  return addBulkSubjectsByClass_(data);
      case 'updateSubject':    return updateSubjectByClass_(data);
      case 'deleteSubject':    return deleteByClassId_(data.id, 'SUB_', data.expect);
      case 'updateAttendance': return updateAttendanceByClass_(data);
      case 'deleteAttendance': return deleteByClassId_(data.id, 'ATT_', data.expect);
      case 'updateSheetId':
        if (!data || !/^[A-Za-z0-9_-]{20,120}$/.test(String(data.sheetId || ''))) {
          throw new Error('รูปแบบ Google Sheet ID ไม่ถูกต้อง');
        }
        PropertiesService.getScriptProperties().setProperty(PROP_KEY_SHEET_ID, data.sheetId);
        return { ok: true, message: 'บันทึก Sheet ID ใหม่แล้ว' };
      default: throw new Error('ไม่รู้จัก action: ' + action);
    }
    });
  } finally {
    // ล้างเฉพาะ action ที่เขียนข้อมูลจริง (listAdmins/getSheetInfo อ่านอย่างเดียว)
    if (MUTATING_ACTIONS.has(action)) cacheClearData_();
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

// ดึงรายชื่อชีตทั้งหมดด้วยการเรียกครั้งเดียว (เดิมเรียก getSheetByName ~30 ครั้งต่อคำขอ)
let SHEET_MAP_ = null;
function getSheetMap_() {
  if (!SHEET_MAP_) {
    SHEET_MAP_ = {};
    getSpreadsheet_().getSheets().forEach(function (sh) { SHEET_MAP_[sh.getName()] = sh; });
  }
  return SHEET_MAP_;
}

function dataSheetEntries_(baseName) {
  const map = getSheetMap_();
  const entries = [];
  if (map[baseName]) entries.push({ name: baseName, cls: '', sheet: map[baseName] });
  CLASS_LIST.forEach(function (cls) {
    const name = baseName + '-' + cls;
    if (map[name]) entries.push({ name: name, cls: cls, sheet: map[name] });
  });
  return entries;
}

function ensureClassSheet__(baseName, cls) {
  const name = classSheetName_(baseName, cls);
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); SHEET_MAP_ = null; }
  const headers = baseName === 'Attendance' ? attendanceHeaders_() : SHEET_DEFS[baseName].slice();
  if (sh.getLastRow() === 0) sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  else ensureHeaderTail_(sh, headers);
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
    readRows_(entry.sheet, 4).forEach(function (r, i) {
      if (!r || String(r[0]) === '') return;
      const cls = normalizeClass_(r[2] || entry.cls);
      result.students.push({
        _id: encodeRowId_('STU_', i + 2, entry.name),
        student_id: String(r[0]), student_name: String(r[1]), student_class: cls
      });
      const studentKey = String(r[0]);
      if (!studentClasses[studentKey]) studentClasses[studentKey] = [];
      if (cls && studentClasses[studentKey].indexOf(cls) === -1) studentClasses[studentKey].push(cls);
    });
  });

  dataSheetEntries_('Subjects').forEach(function (entry) {
    readRows_(entry.sheet, 3).forEach(function (r, i) {
      if (!r || String(r[0]) === '') return;
      const cls = normalizeClass_(r[1] || entry.cls);
      result.subjects.push({
        _id: encodeRowId_('SUB_', i + 2, entry.name),
        subject_name: String(r[0]), subject_class: cls
      });
      const subjectKey = String(r[0]);
      if (!subjectClasses[subjectKey]) subjectClasses[subjectKey] = [];
      if (cls && subjectClasses[subjectKey].indexOf(cls) === -1) subjectClasses[subjectKey].push(cls);
    });
  });

  dataSheetEntries_('Attendance').forEach(function (entry) {
    readRows_(entry.sheet, 8).forEach(function (r, i) {
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
        attendance_notes: String(r[4] || ''),
        attendance_student_class: cls || '',
        attendance_period: String(r[7] || '')
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
      ? cls + '|' + r.attendance_date + '|' + r.attendance_student_id + '|' + r.attendance_subject + '|' + (r.attendance_period || '')
      : 'legacy|' + r._id;
  });
  // แถวเก่าที่ยังไม่มีช่วงเวลาและมีแถวของวัน/วิชาเดียวกันที่ระบุช่วงเวลาแล้ว → ไม่ต้องแสดงซ้ำ
  const attendanceKeyOf = function (r) {
    return normalizeClass_(r.attendance_student_class) + '|' + r.attendance_date + '|' +
      r.attendance_student_id + '|' + r.attendance_subject;
  };
  const periodFilled = new Set(result.attendance.filter(function (r) {
    return !!r.attendance_period;
  }).map(attendanceKeyOf));
  result.attendance = result.attendance.filter(function (r) {
    return !!r.attendance_period || !periodFilled.has(attendanceKeyOf(r));
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

// อ่านข้อมูลครั้งเดียว แล้วเขียนทีเดียวต่อชั้น (เดิมอ่านทั้งไฟล์ซ้ำทุกคน → ช้ามาก)
function addBulkStudentsByClass_(data) {
  const existing = new Set(getAllDataByClass_().students.map(function (s) {
    return normalizeClass_(s.student_class) + '|' + String(s.student_id);
  }));
  const byClass = {};
  let added = 0, skipped = 0;
  (data.list || []).forEach(function (s) {
    const cls = normalizeClass_(s.student_class);
    if (CLASS_LIST.indexOf(cls) === -1) throw new Error('ชั้นเรียนไม่ถูกต้อง: ' + s.student_class);
    const key = cls + '|' + String(s.student_id);
    if (existing.has(key)) { skipped++; return; }
    existing.add(key);
    if (!byClass[cls]) byClass[cls] = [];
    byClass[cls].push([s.student_id, s.student_name, cls, new Date()]);
    added++;
  });
  Object.keys(byClass).forEach(function (cls) {
    const sh = ensureClassSheet__('Students', cls);
    const rows = byClass[cls];
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, 4).setValues(rows);
  });
  return { ok: true, added: added, skipped: skipped };
}

function updateStudentByClass_(data) {
  const loc = decodeRowId_(data.id, 'STU_', 'Students');
  const row = loc.sheet.getRange(loc.row, 1, 1, 4).getValues()[0];
  if (!row || String(row[0]) === '') throw new Error('ไม่พบข้อมูลนักเรียน กรุณารีเฟรช');
  if (!rowMatchesExpect_(loc.sheet, loc.row, data.expect, 2)) throw new Error('ข้อมูลมีการเปลี่ยนแปลงแล้ว กรุณารีเฟรชแล้วลองใหม่');
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
  const cls = normalizeClass_(data.subject_class);
  if (CLASS_LIST.indexOf(cls) === -1) throw new Error('ชั้นเรียนไม่ถูกต้อง: ' + data.subject_class);
  const existing = new Set(getAllDataByClass_().subjects.map(function (s) {
    return normalizeClass_(s.subject_class) + '|' + String(s.subject_name);
  }));
  const rows = [];
  let skipped = 0;
  (data.list || []).forEach(function (name) {
    const key = cls + '|' + String(name);
    if (existing.has(key)) { skipped++; return; }
    existing.add(key);
    rows.push([name, cls, new Date()]);
  });
  if (rows.length) {
    const sh = ensureClassSheet__('Subjects', cls);
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, 3).setValues(rows);
  }
  return { ok: true, added: rows.length, skipped: skipped };
}

function updateSubjectByClass_(data) {
  const loc = decodeRowId_(data.id, 'SUB_', 'Subjects');
  const row = loc.sheet.getRange(loc.row, 1, 1, 3).getValues()[0];
  if (!row || String(row[0]) === '') throw new Error('ไม่พบข้อมูลวิชา กรุณารีเฟรช');
  if (!rowMatchesExpect_(loc.sheet, loc.row, data.expect, 1)) throw new Error('ข้อมูลมีการเปลี่ยนแปลงแล้ว กรุณารีเฟรชแล้วลองใหม่');
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

// ป้องกัน ID ที่เป็นเลขแถวล้าสมัย (มีคนลบแถวอื่นไปแล้ว แถวเลื่อน) → ตรวจว่าแถวตรงกับที่ผู้ใช้เห็นจริง
function rowMatchesExpect_(sh, row, expect, cols) {
  if (expect == null || expect === '') return true;
  if (row > sh.getLastRow()) return false;
  const vals = sh.getRange(row, 1, 1, cols).getValues()[0];
  return vals.map(function (v) { return normalizeDate_(v); }).join('|') === String(expect);
}

function deleteByClassId_(id, prefix, expect) {
  const baseName = prefix === 'STU_' ? 'Students' : prefix === 'SUB_' ? 'Subjects' : 'Attendance';
  const loc = decodeRowId_(id, prefix, baseName);
  if (loc.row > loc.sheet.getLastRow()) throw new Error('ไม่พบข้อมูล กรุณารีเฟรช');
  const cols = prefix === 'ATT_' ? 3 : prefix === 'STU_' ? 2 : 1;
  if (!rowMatchesExpect_(loc.sheet, loc.row, expect, cols)) {
    throw new Error('ข้อมูลมีการเปลี่ยนแปลงแล้ว กรุณารีเฟรชแล้วลองใหม่ (ป้องกันลบผิดแถว)');
  }
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
    const seenPeriodSlots = new Map(); // ชั้น|วัน|ช่วงเวลา → วิชาที่เลือก
    const validated = records.map(function (record) {
      if (!record || typeof record !== 'object' || Array.isArray(record)) throw new Error('รูปแบบรายการเช็คชื่อไม่ถูกต้อง');
      let cls = normalizeClass_(record.cls || '');
      const studentId = String(record.student_id == null ? '' : record.student_id).trim();
      const subject = String(record.subject == null ? '' : record.subject).trim();
      const date = record.date;
      const status = record.status;
      const notes = record.notes == null ? '' : record.notes;
      const period = String(record.period == null ? '' : record.period).trim();

      if (!studentId || studentId.length > 40) throw new Error('เลขที่นักเรียนไม่ถูกต้อง');
      if (!subject || subject.length > 120) throw new Error('รายวิชาไม่ถูกต้อง');
      if (period && ATTENDANCE_PERIOD_LIST.indexOf(period) === -1) throw new Error('ช่วงเวลาเช็คชื่อไม่ถูกต้อง');
      if (!cls) {
        const isActivity = ATTENDANCE_ACTIVITY_LIST.indexOf(subject) !== -1;
        const possibleClasses = data.students.filter(function (student) {
          return String(student.student_id) === studentId &&
            (isActivity || subjects.has(attendanceKey_(student.student_class, '', '', subject)));
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
      // วิชาตามตารางสอนต้องมีในชีต Subjects; กิจกรรมประจำช่วงเวลาบันทึกได้โดยไม่ต้องมีในตารางสอน
      if (!subjects.has(attendanceKey_(cls, '', '', subject)) && ATTENDANCE_ACTIVITY_LIST.indexOf(subject) === -1) {
        throw new Error('รายวิชาไม่ตรงกับชั้นเรียนที่ระบุ');
      }

      const key = attendanceKey_(cls, date, studentId, subject) + '|' + period;
      if (seen.has(key)) throw new Error('มีรายการนักเรียนซ้ำในคำขอเดียวกัน');
      seen.add(key);
      // ในคำขอเดียวกันก็ต้องเป็นวิชาเดียวต่อช่วงเวลา
      if (period) {
        const slot = JSON.stringify([cls, normalizeDate_(date), period]);
        const chosen = seenPeriodSlots.get(slot);
        if (chosen && chosen !== subject) {
          throw new Error('ช่วงเวลา ' + period + ' ต้องเช็คด้วยวิชาเดียวกันทั้งห้อง (พบ ' + chosen + ' และ ' + subject + ')');
        }
        seenPeriodSlots.set(slot, subject);
      }
      return {
        cls: cls, studentId: studentId, subject: subject, date: date, status: status, notes: notes,
        period: period, key: key,
        legacyKey: JSON.stringify([cls, normalizeDate_(date), studentId, subject])
      };
    });

    const existingByKey = new Map();
    const legacyByKey = new Map();
    const unresolvedLegacyKeys = new Set();
    data.attendance.forEach(function (row) {
      const cls = normalizeClass_(row.attendance_student_class);
      if (!cls) {
        unresolvedLegacyKeys.add(JSON.stringify([
          normalizeDate_(row.attendance_date), String(row.attendance_student_id), String(row.attendance_subject)
        ]));
        return;
      }
      const key = attendanceKey_(cls, row.attendance_date, row.attendance_student_id, row.attendance_subject) +
        '|' + String(row.attendance_period || '');
      existingByKey.set(key, row);
      // แถวที่บันทึกไว้ก่อนมีระบบช่วงเวลา (ช่องช่วงเวลาว่าง) → จับคู่ด้วย วัน/เลขที่/วิชา
      if (!row.attendance_period) {
        legacyByKey.set(JSON.stringify([cls, normalizeDate_(row.attendance_date),
          String(row.attendance_student_id), String(row.attendance_subject)]), row);
      }
    });

    // 1 ช่วงเวลา ของ 1 ห้อง ใน 1 วัน ต้องเป็นวิชาเดียวกัน
    // (วิชาเดิมเลือกซ้ำได้ในอีกช่วงเวลา แต่ห้ามแยกวิชาในเวลาเดียวกัน แม้มาจากหลายเครื่อง)
    const periodSubjects = new Map();
    data.attendance.forEach(function (row) {
      const cls = normalizeClass_(row.attendance_student_class);
      const period = String(row.attendance_period || '');
      if (!cls || !period || !row.attendance_subject) return;
      const key = JSON.stringify([cls, normalizeDate_(row.attendance_date), period]);
      if (!periodSubjects.has(key)) periodSubjects.set(key, String(row.attendance_subject));
    });

    const insertsByClass = {};
    const updates = [];
    let inserted = 0, updated = 0, skipped = 0;
    validated.forEach(function (record) {
      if (record.period) {
        const usedSubject = periodSubjects.get(JSON.stringify([record.cls, normalizeDate_(record.date), record.period]));
        if (usedSubject && usedSubject !== record.subject) {
          throw new Error('ช่วงเวลา ' + record.period + ' เช็คชื่อด้วยวิชา ' + usedSubject + ' แล้ว (1 ช่วงเวลาเช็คได้เพียงวิชาเดียว)');
        }
      }
      const match = existingByKey.get(record.key) || legacyByKey.get(record.legacyKey);
      if (match) {
        const sameValues = match.attendance_status === record.status &&
          String(match.attendance_notes || '') === record.notes;
        const samePeriod = String(match.attendance_period || '') === record.period;
        // แถวเก่าที่ยังไม่มีช่วงเวลา → เติมช่วงเวลาได้โดยไม่ต้องเข้าสู่ระบบ (ค่าเดิมไม่เปลี่ยน)
        const periodOnlyFill = sameValues && !match.attendance_period && !!record.period;
        if (!allowUpdates && !periodOnlyFill) {
          if (!sameValues || !samePeriod) {
            throw new Error('รายการนี้ถูกบันทึกแล้ว หากต้องการแก้ไขสถานะ กรุณาเข้าสู่ระบบ');
          }
          skipped++;
          return;
        }
        if (sameValues && samePeriod) { skipped++; return; }
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
          record.notes, new Date(), record.cls, record.period
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
    const headerFixed = {};
    preparedUpdates.forEach(function (prepared) {
      const loc = prepared.loc;
      const item = prepared.item;
      loc.sheet.getRange(loc.row, 4, 1, 2).setValues([[item.record.status, item.record.notes]]);
      if (!item.match.attendance_student_class) {
        if (loc.sheet.getLastColumn() < 7) loc.sheet.getRange(1, 7).setValue('attendance_student_class');
        loc.sheet.getRange(loc.row, 7).setValue(item.record.cls);
      }
      if (!item.match.attendance_period && item.record.period) {
        if (!headerFixed[loc.sheetName]) {
          headerFixed[loc.sheetName] = true;
          if (loc.sheet.getLastColumn() < 8) loc.sheet.getRange(1, 8).setValue('attendance_period');
        }
        loc.sheet.getRange(loc.row, 8).setValue(item.record.period);
      }
    });
    Object.keys(insertsByClass).forEach(function (cls) {
      const sh = ensureClassSheet__('Attendance', cls);
      const rows = insertsByClass[cls];
      sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    });
    cacheClearData_();
    return { ok: true, inserted: inserted, updated: updated, skipped: skipped };
  });
}

function updateAttendanceByClass_(data) {
  const loc = decodeRowId_(data.id, 'ATT_', 'Attendance');
  if (loc.row > loc.sheet.getLastRow()) throw new Error('ไม่พบข้อมูลรายการ กรุณารีเฟรช');
  if (!rowMatchesExpect_(loc.sheet, loc.row, data.expect, 3)) throw new Error('ข้อมูลมีการเปลี่ยนแปลงแล้ว กรุณารีเฟรชแล้วลองใหม่');

  const lastRow = loc.sheet.getLastRow();
  const row = loc.sheet.getRange(loc.row, 1, 1, 8).getValues()[0];
  const date = normalizeDate_(row[0]);
  const studentId = String(row[1] == null ? '' : row[1]).trim();
  const oldSubject = String(row[2] == null ? '' : row[2]).trim();
  const oldPeriod = String(row[7] == null ? '' : row[7]).trim();

  const subject = data.attendance_subject == null ? oldSubject : String(data.attendance_subject).trim();
  const period = data.attendance_period == null ? oldPeriod : String(data.attendance_period).trim();
  const status = data.attendance_status == null ? String(row[3] == null ? '' : row[3]).trim() : String(data.attendance_status).trim();
  const notes = data.attendance_notes == null ? String(row[4] == null ? '' : row[4]) : String(data.attendance_notes).trim();

  if (!subject || subject.length > 120) throw new Error('รายวิชาไม่ถูกต้อง');
  if (ATTENDANCE_STATUS_LIST.indexOf(status) === -1) throw new Error('สถานะการเข้าเรียนไม่ถูกต้อง');
  if (period && ATTENDANCE_PERIOD_LIST.indexOf(period) === -1) throw new Error('ช่วงเวลาเช็คชื่อไม่ถูกต้อง');
  if (notes.length > 200 || /^\s*[=+\-@]/.test(notes)) throw new Error('หมายเหตุไม่ถูกต้องหรือยาวเกิน 200 ตัวอักษร');

  // แก้วิชา/ช่วงเวลา → กันข้อมูลชนกันจากหลายเครื่อง (ตรวจจากชีตจริงก่อนเขียน)
  if (subject !== oldSubject || period !== oldPeriod) {
    if (lastRow > 1) {
      const rows = loc.sheet.getRange(2, 1, lastRow - 1, 8).getValues();
      for (let i = 0; i < rows.length; i++) {
        const r = rows[i];
        if (i + 2 === loc.row) continue;
        if (normalizeDate_(r[0]) !== date) continue;
        const rSubject = String(r[2] == null ? '' : r[2]).trim();
        const rPeriod = String(r[7] == null ? '' : r[7]).trim();
        const rStudent = String(r[1] == null ? '' : r[1]).trim();
        if (period && rPeriod === period && rSubject && rSubject !== subject) {
          throw new Error('ช่วงเวลา ' + period + ' เช็คชื่อด้วยวิชา ' + rSubject + ' แล้ว (1 ช่วงเวลาเช็คได้เพียงวิชาเดียว)');
        }
        if (rStudent === studentId && rSubject === subject && rPeriod === period) {
          throw new Error('มีรายการเช็คชื่อของนักเรียนคนนี้ วิชาและช่วงเวลานี้อยู่แล้ว กรุณาลบรายการซ้ำก่อน');
        }
      }
    }
  }

  loc.sheet.getRange(loc.row, 3, 1, 3).setValues([[subject, status, notes]]);
  if (period !== oldPeriod) {
    if (period && loc.sheet.getLastColumn() < 8) loc.sheet.getRange(1, 8).setValue('attendance_period');
    loc.sheet.getRange(loc.row, 8).setValue(period);
  }
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
      const dataRow = [r[0], r[1], r[2], r[3], r[4] || '', r[5] || new Date(), cls, String(r[7] || '')];
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
