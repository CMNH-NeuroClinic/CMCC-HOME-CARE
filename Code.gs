/******************************************************************************
 * CMCC Home Care by CMNH — ระบบหลังบ้านโฮมแคร์ (สร้างจาก HomeCare Suite) (Google Apps Script + Google Sheets)
 * ------------------------------------------------------------------------------
 * ไฟล์เดียวรองรับ 3 หน้าเว็บ:
 *   1) admin.html     หลังบ้านแอดมิน: เคส/แพ็กเกจ, ใบสรุปค่าบริการ, พนักงาน, การชำระเงิน, สัญญา
 *   2) caregiver.html แอปผู้ดูแล: บันทึกการดูแลรายวัน, เช็คอิน, SOS, ตารางเวร
 *   3) sign.html      หน้าลูกค้าเปิดอ่านและเซ็นสัญญาบริการ (ไม่ต้องล็อกอิน ใช้ลิงก์ + รหัส 6 หลัก)
 *
 * ติดตั้ง (ทำครั้งเดียว — ดูคู่มือฉบับเต็มในไฟล์ INSTALL):
 *   1. สร้าง Google Sheet ใหม่ > ส่วนขยาย > Apps Script > ลบโค้ดเดิม แล้ววางไฟล์นี้ทั้งไฟล์
 *   2. แก้ INITIAL_ADMIN_PASSWORD ด้านล่างเป็นรหัสผ่านแอดมิน (อย่างน้อย 8 ตัวอักษร)
 *   3. เลือกฟังก์ชัน setup แล้วกด Run (อนุญาตสิทธิ์ตามที่ Google ถาม)
 *   4. การทำให้ใช้งานได้ > การทำให้ใช้งานได้รายการใหม่ > เว็บแอป
 *        ดำเนินการในฐานะ: ฉัน · ผู้มีสิทธิ์เข้าถึง: ทุกคน
 *   5. คัดลอกลิงก์ .../exec ไปใส่ในตัวแปร API_URL ของ admin.html, caregiver.html และ sign.html
 *
 * ความปลอดภัย:
 *   - ทุกคำขอข้อมูลใช้ POST + session token (ไม่มีรหัสผ่านใน URL)
 *   - แอดมินใช้รหัสผ่าน (เก็บแบบ hash ใน Script Properties) · ผู้ดูแลใช้รหัสส่วนตัว ≥ 6 ตัว
 *   - รูปเช็คอินเก็บแบบส่วนตัวใน Drive ไม่แชร์ลิงก์สาธารณะ
 ******************************************************************************/

var INITIAL_ADMIN_PASSWORD = 'CHANGE-ME';   // <<< เปลี่ยนก่อนรัน setup() (อย่างน้อย 8 ตัวอักษร)
var PRODUCT_NAME = 'CMCC Home Care';        // ชื่อระบบ (แสดงท้ายรายงาน) — ชื่อร้านตั้งในหน้า "ตั้งค่า" ของแอดมิน
var TZ = 'Asia/Bangkok';
var SESSION_SECONDS = 21600;                // เข้าสู่ระบบค้างได้ 6 ชั่วโมง (ขีดจำกัดของ CacheService)
var SS = SpreadsheetApp.getActiveSpreadsheet();
var CURRENT_SESSION_ = null;

/* ============================== ติดตั้งครั้งแรก ============================== */

function onOpen() {
  SpreadsheetApp.getUi().createMenu(PRODUCT_NAME)
    .addItem('ติดตั้ง/ซ่อมชีทที่ขาด (ปลอดภัย ไม่ลบข้อมูล)', 'setup')
    .addToUi();
}

function setup() {
  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('ADMIN_PW_HASH')) {
    if (!INITIAL_ADMIN_PASSWORD || INITIAL_ADMIN_PASSWORD === 'CHANGE-ME' || String(INITIAL_ADMIN_PASSWORD).length < 8) {
      throw new Error('กรุณาแก้ INITIAL_ADMIN_PASSWORD ที่ต้นไฟล์เป็นรหัสผ่านของคุณ (อย่างน้อย 8 ตัวอักษร) แล้วรัน setup ใหม่');
    }
    setAdminPassword_(String(INITIAL_ADMIN_PASSWORD));
  }
  ensureSystemSheets_();
  hcSetupSheets_();
  cgSetupSheets_();
  var st = readSettingsRaw_();
  Object.keys(DEFAULT_SETTINGS).forEach(function (k) { if (st[k] === undefined) setSetting_(k, DEFAULT_SETTINGS[k]); });
  if (!readSettingsRaw_().driveFolderId) setSetting_('driveFolderId', DriveApp.createFolder(PRODUCT_NAME + ' - เอกสาร').getId());
  var first = SS.getSheetByName('Sheet1') || SS.getSheetByName('ชีต1');
  if (first && SS.getSheets().length > 1 && first.getLastRow() === 0) SS.deleteSheet(first);
  Logger.log('ติดตั้งเรียบร้อย ขั้นต่อไป: การทำให้ใช้งานได้ > เว็บแอป');
}

// ลืมรหัสผ่านแอดมิน: แก้ INITIAL_ADMIN_PASSWORD แล้วรันฟังก์ชันนี้
function resetAdminPassword() {
  if (!INITIAL_ADMIN_PASSWORD || INITIAL_ADMIN_PASSWORD === 'CHANGE-ME' || String(INITIAL_ADMIN_PASSWORD).length < 8) throw new Error('แก้ INITIAL_ADMIN_PASSWORD ก่อน (อย่างน้อย 8 ตัวอักษร)');
  setAdminPassword_(String(INITIAL_ADMIN_PASSWORD));
  Logger.log('ตั้งรหัสผ่านใหม่แล้ว');
}

var SYS_SCHEMA_ = {
  Settings: ['Key', 'Value'],
  Contracts: ['ContractID', 'CaseID', 'Type', 'Status', 'ClientName', 'ContactName', 'ContactPhone', 'Snapshot', 'ContentHash',
              'SignToken', 'SignCode', 'SignExpires', 'SignedAt', 'SignerName', 'SignerRole', 'SignerRelation', 'PdpaConsent',
              'SignedUA', 'PdfUrl', 'PdfId', 'SignatureFileId', 'CreatedAt', 'CreatedBy', 'Note', 'CancelReason',
              'StaffIDs', 'StaffSnapshot', 'FirstOpenedAt', 'LastOpenedAt', 'OpenCount',
              'SignerIdNo', 'SignerExtra', 'StaffID'],
  Log: ['At', 'Action', 'By', 'Ref', 'Detail'],
  Officers: ['OfficerID', 'Name', 'CodeHash', 'Active', 'CreatedAt', 'LastLogin']
};

function ensureSystemSheets_() {
  Object.keys(SYS_SCHEMA_).forEach(function (name) {
    var sh = SS.getSheetByName(name);
    if (sh) {
      var cols0 = SYS_SCHEMA_[name];
      var have = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0].map(String);
      var miss = cols0.filter(function (c) { return have.indexOf(c) === -1; });
      if (miss.length) sh.getRange(1, sh.getLastColumn() + 1, 1, miss.length).setValues([miss]).setFontWeight('bold');
      return;
    }
    sh = SS.insertSheet(name);
    var cols = SYS_SCHEMA_[name];
    sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold').setBackground('#E7EFEC');
    sh.setFrozenRows(1);
    sh.getRange(2, 1, 9999, cols.length).setNumberFormat('@');
  });
}

/* ============================== ทางเข้า Web App ============================== */

function doGet() {
  return out_({ ok: true, success: true, name: PRODUCT_NAME + ' API' });
}

function doPost(e) {
  var res;
  try {
    if (!e || !e.postData) throw userError_('ต้องเรียกผ่านหน้าเว็บเท่านั้น');
    var req = JSON.parse(e.postData.contents);
    res = route_(req);
  } catch (err) {
    if (!err.userMessage) console.error(err && err.stack ? err.stack : err);
    var msg = err.userMessage || ('เกิดข้อผิดพลาดในระบบ: ' + (err && err.message ? err.message : err));
    res = { ok: false, success: false, error: msg, message: msg, code: err.code || '' };
  }
  return out_(res);
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function userError_(msg, code) {
  var e = new Error(msg);
  e.userMessage = msg;
  e.code = code || '';
  return e;
}
function fail_(msg, code) { throw userError_(msg, code); }

/* ---------- ตารางคำสั่ง: role = public | caregiver (ผู้ดูแลหรือแอดมิน) | admin ---------- */

var ROUTES_ = {
  // สาธารณะ
  'public.branding':      { role: 'public', fn: function () { return publicBranding_(); } },
  'auth.adminLogin':      { role: 'public', fn: function (d) { return adminLogin_(d); } },
  'auth.caregiverLogin':  { role: 'public', fn: function (d) { return caregiverLogin_(d); } },
  'auth.officerLogin':    { role: 'public', fn: function (d) { return officerLogin_(d); } },
  'auth.me':              { role: 'admin', perm: 'any', fn: function (d, s) { return { success: true, role: s.role, name: s.name, perms: sessionPerms_(s) }; } },
  'officers.list':        { role: 'admin', fn: function () { return officersList_(); } },
  'officers.save':        { role: 'admin', lock: true, fn: function (d) { return officersSave_(d); } },
  'officers.perms':       { role: 'admin', lock: true, fn: function (d) { return officersSetPerms_(d); } },
  'sign.get':             { role: 'public', fn: function (d) { return signGet_(d); } },
  'sign.submit':          { role: 'public', lock: true, fn: function (d) { return signSubmit_(d); } },
  'auth.logout':          { role: 'caregiver', fn: function (d, s, token) { CacheService.getScriptCache().remove('sess_' + token); return { success: true }; } },

  // ผู้ดูแล (แอดมินเรียกได้ด้วย)
  'care.patients':        { role: 'caregiver', fn: function (d, s) { return { success: true, patients: getPatients(s.role === 'caregiver' ? s.code : null) }; } },
  'care.vitalThresholds': { role: 'caregiver', fn: function () { return getVitalThresholds(); } },
  'care.existingRecords': { role: 'caregiver', fn: function (d, s) { assertPatient_(s, d.patientName); return getExistingRecordsForTime(d.patientName, d.date, d.time); } },
  'care.dailyTimesSummary': { role: 'caregiver', fn: function (d, s) { assertPatient_(s, d.patientName); return getTimesSummaryForDate_(d.patientName, d.date); } },
  'care.generateReport':  { role: 'caregiver', lock: true, fn: function (d, s) { assertPatient_(s, d.patientName); return generateDailyReport(d.patientName, d.date); } },
  'care.submitRecord':    { role: 'caregiver', lock: true, fn: function (d, s) {
      if (s.role === 'caregiver') { d.caregiverCode = s.code; d.caregiverName = s.name; }
      else { d.caregiverCode = ''; d.caregiverName = d.caregiverName || s.name; }
      return submitRecord(d);
    } },
  'care.checkIn':         { role: 'caregiverOnly', lock: true, fn: function (d, s) { d.caregiverCode = s.code; return checkIn(d); } },
  'care.sos':             { role: 'caregiverOnly', lock: true, fn: function (d, s) { d.caregiverCode = s.code; return triggerSOS(d); } },
  'care.mySchedule':      { role: 'caregiverOnly', fn: function (d, s) { return getSchedule(d.dateFrom, d.dateTo, s.code); } },

  // แอดมิน: ฝั่งบันทึกการดูแล
  'care.dailyRecords':    { role: 'admin', perm: 'reports', fn: function (d) { return getDailyRecordsForView(d.patientName, d.dateFrom || d.date, d.dateTo || d.dateFrom || d.date); } },
  'care.activityLog':     { role: 'admin', perm: 'reports', fn: function (d) { return getActivityLog(parseInt(d.limit, 10) || 300); } },
  'care.checkIns':        { role: 'admin', perm: 'reports', fn: function (d) { return getCheckIns(parseInt(d.limit, 10) || 200); } },
  'care.checkinPhoto':    { role: 'admin', perm: 'reports', fn: function (d) { return getCheckinPhoto(d.photoId); } },
  'care.caregiversList':  { role: 'admin', perm: 'staff.logins', fn: function () { return getCaregiversList(); } },
  'care.schedule':        { role: 'admin', perm: 'reports', fn: function (d) { return getSchedule(d.dateFrom, d.dateTo, d.caregiverCode || ''); } },
  'care.addSchedule':     { role: 'admin', perm: 'reports', lock: true, fn: function (d) { return addScheduleEntry(d); } },
  'care.deleteSchedule':  { role: 'admin', perm: 'reports', lock: true, fn: function (d) { return deleteScheduleEntry(d); } },
  'care.emergencies':     { role: 'admin', perm: 'reports', fn: function (d) { return getEmergencies(parseInt(d.limit, 10) || 100); } },
  'care.ackEmergency':    { role: 'admin', perm: 'reports', lock: true, fn: function (d) { return acknowledgeEmergency(d); } },
  'care.payRates':        { role: 'admin', perm: 'payroll', fn: function () { return getPayRates(); } },
  'care.savePayRates':    { role: 'admin', perm: 'payroll', lock: true, fn: function (d) { return savePayRates(d); } },
  'care.saveVitalThresholds': { role: 'admin', lock: true, fn: function (d) { return saveVitalThresholds(d); } },
  'care.payroll':         { role: 'admin', perm: 'payroll', fn: function (d) { return getPayrollSummary(d.dateFrom, d.dateTo); } },
  'care.registerCaregiver': { role: 'admin', perm: 'staff.logins', lock: true, fn: function (d) { return registerCaregiver(d); } },
  'care.registerPatient': { role: 'admin', perm: 'cases.edit', lock: true, fn: function (d) { return registerPatient(d); } },
  'care.patientsAdmin':   { role: 'admin', perm: 'reports', fn: function () { return getPatientsAdmin_(); } },
  'care.savePatient':     { role: 'admin', perm: 'reports', lock: true, fn: function (d) { return savePatientAdmin_(d); } },

  // แอดมิน: เคส / พนักงาน / การชำระเงิน (ล็อกกันชนกันอยู่ใน hcPost_ เองแล้ว)
  'cases.list':           { role: 'admin', perm: 'cases.edit', fn: function () { return hcGet_({ parameter: {} }); } },
  'cases.history':        { role: 'admin', perm: 'cases.edit', fn: function () { return hcGet_({ parameter: { type: 'history' } }); } },
  'cases.save':           { role: 'admin', perm: 'cases.edit', fn: function (d, s) {
      delete d.entity; d.recordedBy = d.recordedBy || s.name;
      if (Array.isArray(d.staffIds)) d.staffIds = d.staffIds.join(',');
      if (!d.deleteId) applyPriceApproval_(d, s);
      var res = hcPost_(d);
      if (res && res.ok && d.priceStatus === 'รออนุมัติ' && d._priceNewlyPending) notifyPricePending_(d, s, res.id);
      if (res && res.ok && !d.deleteId && d.clientName) ensurePatientFromCase_(d);
      return res;
    } },
  'cases.approvePrice':   { role: 'admin', lock: true, fn: function (d, s) { return approvePrice_(d, s); } },
  'staff.list':           { role: 'admin', perm: 'staff.edit', fn: function (d, s) { var r = hcGet_({ parameter: { type: 'staff' } }); if (!canDo_(s, 'staff.pay')) r.staff.forEach(function (x) { STAFF_PAY_HEADERS_.forEach(function (h) { delete x[h]; }); }); return r; } },
  'staff.save':           { role: 'admin', perm: 'staff.edit', fn: function (d) { d.entity = 'staff'; return hcPost_(d); } },
  'staff.savePhoto':      { role: 'admin', perm: 'staff.edit', lock: true, fn: function (d) { return staffSavePhoto_(d); } },
  'staff.photos':         { role: 'admin', perm: 'staff.edit', fn: function () { return staffPhotos_(); } },
  'staff.profilePreview': { role: 'admin', perm: 'staff.edit', fn: function (d) { return { success: true, profiles: buildStaffSnapshot_(d.staffIds || [], true), watermark: getWatermarkDataUrl_(), logo: getLogoDataUrl_(), company: getSettings_().companyName }; } },
  'payments.list':        { role: 'admin', perm: 'accounts', fn: function () { return hcGet_({ parameter: { type: 'payments' } }); } },
  'payments.save':        { role: 'admin', perm: 'accounts', fn: function (d, s) { d.entity = 'payment'; d.recordedBy = d.recordedBy || s.name; return hcPost_(d); } },

  // แอดมิน: สัญญา
  'contracts.list':       { role: 'admin', perm: 'contracts.view', fn: function (d, s) { return contractsList_(s); } },
  'contracts.createEmployment': { role: 'admin', perm: 'staff.pay', lock: true, fn: function (d, s) { return contractsCreateEmployment_(d, s); } },
  'contracts.create':     { role: 'admin', perm: 'contracts.send', lock: true, fn: function (d, s) { return contractsCreate_(d, s); } },
  'contracts.resend':     { role: 'admin', perm: 'contracts.send', lock: true, fn: function (d, s) { return contractsResend_(d, s); } },
  'contracts.cancel':     { role: 'admin', perm: 'contracts.cancel', lock: true, fn: function (d, s) { return contractsCancel_(d, s); } },

  // แอดมิน: ตั้งค่า
  'settings.get':         { role: 'admin', perm: 'any', fn: function () { return settingsForAdmin_(); } },
  'settings.save':        { role: 'admin', lock: true, fn: function (d) { return settingsSave_(d); } },
  'settings.saveLogo':    { role: 'admin', lock: true, fn: function (d) { return settingsSaveLogo_(d); } },
  'admin.changePassword': { role: 'admin', lock: true, fn: function (d) { return changeAdminPassword_(d); } },
  'system.testPdf':       { role: 'admin', fn: function () { return systemTestPdf_(); } }
};

function route_(req) {
  var action = String(req.action || '');
  var route = ROUTES_[action];
  if (!route) fail_('ไม่รู้จักคำสั่ง: ' + action);
  var data = req.data || req.payload || {};
  if (typeof data !== 'object') data = {};
  _settingsCache = null;

  var s = null;
  if (route.role !== 'public') {
    s = readSession_(req.token);
    if (route.role === 'admin' && s.role !== 'admin') {
      // พนักงานออฟฟิศ (officer) ทำได้เฉพาะคำสั่งที่เจ้าของเปิดสิทธิ์ไว้
      if (s.role !== 'officer' || !route.perm || (route.perm !== 'any' && sessionPerms_(s).indexOf(route.perm) === -1)) fail_('บัญชีของคุณไม่มีสิทธิ์ทำรายการนี้ — ติดต่อเจ้าของกิจการ', 'FORBIDDEN');
      guardOfficer_(action, data, s);
    }
    if (route.role === 'caregiverOnly' && s.role !== 'caregiver') fail_('คำสั่งนี้สำหรับผู้ดูแลเท่านั้น', 'FORBIDDEN');
    if (s.role === 'admin' || s.role === 'officer') data.adminCode = 'session';   // ฟังก์ชันเดิมบางตัวเช็คว่ามี adminCode — ตัวจริงเช็คผ่าน session
    else delete data.adminCode;
  }
  CURRENT_SESSION_ = s;

  var lock = null;
  if (route.lock) { lock = LockService.getScriptLock(); lock.waitLock(25000); }
  try {
    var result = route.fn(data, s, req.token);
    if (result && typeof result === 'object' && !Array.isArray(result)) {
      if (result.ok === undefined && result.success !== undefined) result.ok = result.success;
      if (result.success === undefined && result.ok !== undefined) result.success = result.ok;
      if (result.ok === undefined) { result.ok = true; result.success = true; }
      if (result.error === undefined && result.message !== undefined && result.ok === false) result.error = result.message;
      if (result.message === undefined && result.error !== undefined) result.message = result.error;
    }
    if (route.lock && s) log_(action, s.name, data.id || data.ContractID || data.patientName || '', '');
    return result;
  } finally {
    if (lock) lock.releaseLock();
  }
}

/* ============================== เข้าสู่ระบบ / session ============================== */

function hashPw_(pw, salt) {
  return sha256hex_(salt + '|' + pw);
}
function setAdminPassword_(pw) {
  var salt = Utilities.getUuid();
  var props = PropertiesService.getScriptProperties();
  props.setProperty('ADMIN_PW_SALT', salt);
  props.setProperty('ADMIN_PW_HASH', hashPw_(pw, salt));
}
function checkAdminPassword_(pw) {
  var props = PropertiesService.getScriptProperties();
  var salt = props.getProperty('ADMIN_PW_SALT'), h = props.getProperty('ADMIN_PW_HASH');
  return !!(salt && h && hashPw_(String(pw || ''), salt) === h);
}

function newSession_(obj) {
  var token = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
  CacheService.getScriptCache().put('sess_' + token, JSON.stringify(obj), SESSION_SECONDS);
  return token;
}

function readSession_(token) {
  token = String(token || '');
  var raw = token.length >= 40 ? CacheService.getScriptCache().get('sess_' + token) : null;
  if (!raw) fail_('หมดเวลาเข้าสู่ระบบ กรุณาเข้าสู่ระบบใหม่', 'AUTH');
  var s = JSON.parse(raw);
  if (s.role === 'officer') {
    var off = findOfficer_(s.officerId);
    if (!off || off.Active !== 'Y') { CacheService.getScriptCache().remove('sess_' + token); fail_('บัญชีนี้ถูกปิดใช้งานแล้ว', 'AUTH'); }
    s.name = off.Name;
  }
  if (s.role === 'caregiver') {
    // ผู้ดูแลที่ถูกลบออกจากชีท Caregivers จะใช้งานต่อไม่ได้ทันที
    var cg = verifyCaregiverCode(s.code);
    if (!cg.success) { CacheService.getScriptCache().remove('sess_' + token); fail_('บัญชีนี้ถูกปิดใช้งานแล้ว', 'AUTH'); }
    s.name = cg.name;
  }
  return s;
}

function throttle_(key, max, seconds) {
  var cache = CacheService.getScriptCache();
  var n = Number(cache.get(key) || 0);
  if (n >= max) fail_('ใส่รหัสผิดหลายครั้ง กรุณารอ ' + Math.round(seconds / 60) + ' นาทีแล้วลองใหม่');
  return { bump: function () { cache.put(key, String(n + 1), seconds); }, clear: function () { cache.remove(key); } };
}

function adminLogin_(d) {
  var t = throttle_('fail_admin', 8, 600);
  if (!checkAdminPassword_(d.password)) { t.bump(); fail_('รหัสผ่านไม่ถูกต้อง'); }
  t.clear();
  var name = String(d.name || '').trim().slice(0, 60) || 'แอดมิน';
  var sess = { role: 'admin', name: name };
  return { success: true, token: newSession_(sess), name: name, role: 'admin', perms: sessionPerms_(sess) };
}

function caregiverLogin_(d) {
  var t = throttle_('fail_cg', 25, 600);
  var code = String(d.code || '').trim();
  var cg = code ? verifyCaregiverCode(code) : { success: false };
  if (!cg.success) { t.bump(); fail_('รหัสไม่ถูกต้อง'); }
  t.clear();
  return { success: true, token: newSession_({ role: 'caregiver', code: code, name: cg.name }), name: cg.name, assignedPatients: cg.assignedPatients };
}

function changeAdminPassword_(d) {
  if (!checkAdminPassword_(d.current)) fail_('รหัสผ่านปัจจุบันไม่ถูกต้อง');
  var next = String(d.next || '');
  if (next.length < 8) fail_('รหัสผ่านใหม่ต้องยาวอย่างน้อย 8 ตัวอักษร');
  setAdminPassword_(next);
  return { success: true };
}

// ผู้ดูแลเข้าถึงได้เฉพาะคนไข้ที่ได้รับมอบหมาย (ถ้าไม่ได้ผูกไว้ = ทุกคน)
function assertPatient_(s, patientName) {
  if (!patientName) fail_('กรุณาระบุชื่อคนไข้');
  if (s.role !== 'caregiver') return;
  var cg = verifyCaregiverCode(s.code);
  if (cg.assignedPatients && cg.assignedPatients.indexOf(String(patientName).trim()) === -1) fail_('คุณไม่มีสิทธิ์ดูข้อมูลของคนไข้ท่านนี้', 'FORBIDDEN');
}


/* ============================== พนักงานออฟฟิศ + สิทธิ์การใช้งาน ==============================
 * เจ้าของ (รหัสผ่าน) = ทำได้ทุกอย่าง · พนักงานออฟฟิศ (รหัสส่วนตัว 6 หลัก) = ทำได้ตามสิทธิ์ที่เจ้าของติ๊กไว้
 */
var PERMISSIONS_ = [
  ['cases.edit',       'ขาย: เปิด/แก้ไขเคส กรอกราคา ออกใบเสนอราคา', true],
  ['cases.delete',     'ลบเคส', false],
  ['staff.edit',       'ประวัติเจ้าหน้าที่: เพิ่ม/แก้ไขประวัติ', true],
  ['staff.pay',        'เห็นและแก้ค่าตอบแทนของพนักงาน', false],
  ['staff.logins',     'สร้างบัญชีเข้าแอปผู้ดูแล', true],
  ['contracts.view',   'ดูรายการสัญญา', true],
  ['contracts.send',   'ทำสัญญาและส่งลิงก์ให้ลูกค้าเซ็น', true],
  ['contracts.cancel', 'ยกเลิกสัญญา', false],
  ['accounts',         'บัญชี: ดูยอดค้างและรับชำระเงิน', true],
  ['payments.delete',  'ลบรายการรับเงิน', false],
  ['payroll',          'เงินเดือนผู้ดูแล', false],
  ['reports',          'รายงาน: SOS เช็คอิน รายงานการดูแล', true]
];
var DEFAULT_OFFICER_PERMS_ = PERMISSIONS_.filter(function (p) { return p[2]; }).map(function (p) { return p[0]; });
var STAFF_PAY_HEADERS_ = ['อัตราค่าตอบแทนต่อชั่วโมง (บาท)', 'ประเภทค่าตอบแทน', 'อัตราค่าตอบแทนต่อเวร (บาท)', 'หมายเหตุ'];
var STAFF_PAY_KEYS_ = ['hourlyRate', 'payType', 'shiftRate', 'notes'];

function sessionPerms_(s) {
  if (!s) return [];
  if (s.role === 'admin') return PERMISSIONS_.map(function (p) { return p[0]; }).concat(['settings']);
  if (s.role !== 'officer') return [];
  var allowed = PERMISSIONS_.map(function (p) { return p[0]; });
  return String(getSettings_().officerPerms || '').split(',').filter(function (k) { return allowed.indexOf(k) >= 0; });
}
function canDo_(s, perm) { return sessionPerms_(s).indexOf(perm) >= 0; }

// ตรวจรายละเอียดเพิ่มสำหรับพนักงานออฟฟิศ (สิทธิ์ย่อยภายในคำสั่งเดียวกัน)
function guardOfficer_(action, d, s) {
  if (action === 'cases.save' && d.deleteId && !canDo_(s, 'cases.delete')) fail_('บัญชีของคุณลบเคสไม่ได้ — ติดต่อเจ้าของกิจการ', 'FORBIDDEN');
  if (action === 'payments.save' && (d.deleteId || d.id) && !canDo_(s, 'payments.delete')) fail_('บัญชีของคุณลบหรือแก้รายการรับเงินไม่ได้ — ติดต่อเจ้าของกิจการ', 'FORBIDDEN');
  if (action === 'staff.save' && !canDo_(s, 'staff.pay')) STAFF_PAY_KEYS_.forEach(function (k) { delete d[k]; });
  d.recordedBy = s.name;   // พนักงานใช้ชื่อบัญชีตัวเองเสมอ แก้ชื่อผู้บันทึกไม่ได้
}

function officersSheet_() { return SS.getSheetByName('Officers') || (ensureSystemSheets_(), SS.getSheetByName('Officers')); }
function readOfficers_() {
  var sh = officersSheet_(), cols = SYS_SCHEMA_.Officers, last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, cols.length).getValues().map(function (r, i) {
    var o = { _row: i + 2 }; cols.forEach(function (k, j) { o[k] = r[j] instanceof Date ? Utilities.formatDate(r[j], TZ, 'yyyy-MM-dd HH:mm:ss') : String(r[j]); }); return o;
  }).filter(function (o) { return o.OfficerID; });
}
function findOfficer_(id) { var l = readOfficers_(); for (var i = 0; i < l.length; i++) if (l[i].OfficerID === String(id)) return l[i]; return null; }
function writeOfficer_(o) {
  var sh = officersSheet_(), row = SYS_SCHEMA_.Officers.map(function (k) { return o[k] == null ? '' : o[k]; });
  if (o._row) sh.getRange(o._row, 1, 1, row.length).setValues([row]); else sh.appendRow(row);
}
function officerHash_(code) { return sha256hex_('officer|' + PropertiesService.getScriptProperties().getProperty('ADMIN_PW_SALT') + '|' + String(code).trim()); }

function officerLogin_(d) {
  var t = throttle_('fail_officer', 15, 600);
  var code = String(d.code || '').trim();
  var hsh = code.length >= 6 ? officerHash_(code) : '';
  var off = hsh ? readOfficers_().filter(function (o) { return o.CodeHash === hsh && o.Active === 'Y'; })[0] : null;
  if (!off) { t.bump(); fail_('รหัสพนักงานไม่ถูกต้อง'); }
  t.clear();
  off.LastLogin = nowStr_(); writeOfficer_(off);
  var sess = { role: 'officer', name: off.Name, officerId: off.OfficerID };
  return { success: true, token: newSession_(sess), name: off.Name, role: 'officer', perms: sessionPerms_(sess) };
}

function officersList_() {
  return { success: true,
    officers: readOfficers_().map(function (o) { return { id: o.OfficerID, name: o.Name, active: o.Active === 'Y', createdAt: o.CreatedAt, lastLogin: o.LastLogin }; }),
    permissions: PERMISSIONS_.map(function (p) { return { key: p[0], label: p[1] }; }),
    officerPerms: String(getSettings_().officerPerms).split(',').filter(Boolean) };
}

// เพิ่มพนักงาน / ออกรหัสใหม่ / ปิด-เปิดบัญชี — รหัสแสดงครั้งเดียวตอนออก (ระบบเก็บเป็น hash)
function officersSave_(d) {
  var o, code = '';
  if (d.id) { o = findOfficer_(d.id); if (!o) fail_('ไม่พบบัญชีพนักงาน'); }
  else {
    var name = String(d.name || '').trim().slice(0, 60);
    if (name.length < 2) fail_('กรุณาใส่ชื่อพนักงาน');
    o = { OfficerID: 'OF' + Date.now().toString(36).toUpperCase(), Name: name, Active: 'Y', CreatedAt: nowStr_() };
  }
  if (!d.id || d.resetCode) {
    var hashes = readOfficers_().map(function (x) { return x.CodeHash; });
    do { code = String(Math.floor(100000 + Math.random() * 900000)); } while (hashes.indexOf(officerHash_(code)) >= 0);
    o.CodeHash = officerHash_(code);
  }
  if (d.active !== undefined) o.Active = d.active ? 'Y' : 'N';
  if (d.name && d.id) o.Name = String(d.name).trim().slice(0, 60);
  writeOfficer_(o);
  var res = officersList_(); res.newCode = code; res.codeFor = o.Name;
  return res;
}

function officersSetPerms_(d) {
  var allowed = PERMISSIONS_.map(function (p) { return p[0]; });
  var list = (Array.isArray(d.perms) ? d.perms : []).filter(function (k) { return allowed.indexOf(k) >= 0; });
  setSetting_('officerPerms', list.join(',') || 'none');
  return officersList_();
}

/* ============================== ตั้งค่าร้าน ============================== */

var DEFAULT_WATERMARK_PNG_ = 'iVBORw0KGgoAAAANSUhEUgAAAjAAAAIwCAMAAACvL6FdAAAA/1BMVEX+/v7////U8PDl5eXT09PX7Ovg39/c3+Dg3+Da4N7g4N8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAUf8aYAAAAQHRSTlP/AP///////////wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAALnm1UgAAGdxJREFUeNrt3Yty5DYOBVCQEGeS/f8P3ozdY7ftlsQXgAsSrNpNUslYbPH4gqQeTSnaW6O7Fqfo/TwFko4WYMJJyAkwklK2dUNhJdQEGGUqO7GhwBJoAoydleXVUGB5b/muBZo1wUxFMgNPgPGvJc9qG5qhrbBkmbYTGtpES5Zve5ih9bVkzba8GVpaS7ZpK5uhZbVk27aqGVpSS8ZoK5qh5bRkrLaaGQotYWZxMK6w3KEJMKFlbTO0ApfspS1Ahtxzyb6adzIUWsLMgmDktehejfJLhrbm0nkz3c5kyCcXGygKvQgwWFpE7t7dygxtw0X4nu9dyNAOXNSeFNiADHnigmxldi8DjD4Xw4eS1iVDXrz4wTKtxwFGiQvhtPXIkAMuXrXM6H2AkeRCqG0hMrQMF8Juq5AhaC+raBn8NAFmLhfy0/yToeASZDyC6Tyx5LG5nsqQ43ghv81vyFBwCTLewHScSlqhdXzQALMvF6dkyF81opWau7pEzuKF1mu+QoaCS5DxAqb1tNG6zY8YQvGyM5eOz78fmMbzRes3HyFDLuKF9mgeQoaCS5CBB9N0kmivhi6GrL0ElxEyO4BpOTu0Z0MOGQKOF9q34YYMhZcQgwum4aTQ7g20LBFmvFC0ljO0KJjqsxFWWk/SimAovAiKoeXABJc1QobCS4gBBFN7CsJHGxl9MaTuJbiIkVkFTMTLQiFDMF5CRScZXTEU5SjKEhSYiBflkHEOJrwsJobCS4jBAVP1YcPBDDJaYijiJUIGBEx4WVEMRTmKsgQBJrysKYbCS4gBAFPz6WLcBciIi6HwEmLMwYSXdcWQjZcYbzkysmJI1EvEi3XIOAATXpYWQ+ElxFiCCS+Li6HwEmLswMTyCHKxBAsmvKwvhsJLiDECE152EEPhJcRYgwkvgGLgwISXPcSQVkGKITURM70oUXgJMfpgwss2Yii8hBhtMCt5yR1tKzEUXjqZPLeC8inlxZD8gnp1LU9q8MUAgHHtJc9tpZTVxdC+XmZb+fzrymJoVy95eivPeJYVQ7IT3m20fINjWptEJ760n5cs28qjOBnGjKQYkixIG3LBOAWCRYm28pK1vViVJjkxtJGXbNGMyIiJoW28mGh5mFlIjBiY4GKcMnBgPHnJ5s2AjIwY2qAgGVP5+JslihKFl2VTRkQMrV2Q/vQl4zTlvTwgMF4mMBmsKV+ZFBBD4UXbjG8xtHY9ghSTPYuZDwbnKgCqGNWQQQDjIWD+9GNgn+2tLbFemh0x5M0Lyy2OylP78lG++5mAya2YETAGXphlpi8PJndH/9A0IYIKhBgFMJYTmFov1IylqRf/UfkSRbnZTjEW0x8x5KkgVXp574j0dcHxyU7xWJRoTS91YCb0acxM8SeG3ExguLoe1YXLtNEYMYMgRhSM3QSmwQupYfkeNJ7EdEYMOSlIE70I9fBzAZVbJsI6WzLzihK5KEjc4oWsbuPvXGwXV0WJPBSkSi/5Foz4uHSRKZ6KEs0KGAAvfzth+ogQ/yVTavdjAMQIgTEqSK1ezK9wNaTMg5bGBexJRYngC1Lb/MVmB3q4MmlMfecUJQIvSNzqBeT6OTsTIwDm6sCyXtgdl7bCVOzEtEcMQRekVi8EdXdOKf/+vTiJkTEzihIhF6Q6L/n0lJB1K+V428vzIGYyGIuCVJ0v+XXnCKD9IQO0ITMeMYQbMA31KL/oHWE0Lr8ryBQlMeMRQ+BeuOt8EFJ7r0vl5hYupZ4PixkHA+Hl6z0wBNbe6tLVkunp3xkWpYlg9AOmMV+wX03z2JUpCPOY0YghzBlvlZeMt5K+3JXB2MEbnPcSZMA016NHTcq4Lwau3vqVFjMYMQTspacgEW6rEfN+BxayGAKc8fZ7Iej2uCMPcBozF4x2wHR7Ifj2mPsWx0WJ4AKGe8EQuRFjnDGyYHx4ISftbiJTCrgYWDALxssHGfN3yEiCgfbytx+uvoW0qizBiiGsPTvuAePtW2vrxIDu3hFUwFR5+bvB627+8mUTz3ji2x0xhLSk5vqAyZ9giHyKMc2Y7qU1AQUMtxSkj4vTRK7FFBgxM8AgT2By9uulbrEEWZTIacB0xUvzE07iYgpSURoHg1yQ2nvBfLw1FDU+xcCAaZvwtnfioeVhxklVcgbGxgvXftSWH318bzgbMtd3+8KJIWcB837oli4crxtHxoiBQQqYVjB8nDYfYmwipguMYsBw84yXhr3giLlUgxYx5CZg8sdfJnkBEoNXlDrAoAVM7pjvXnvxMY8Bi5geMEYFqf3i0XE4EMNQYvrB6AUM14JpvXrEt2COis6J47q/do0UMeQkYHLHft29lxsFrBNId2KgIoacBExuv1uqImCuI0athhmKaY4Ysg6YyoLUcfwaLxcCWHFldQcm40QMWa+pG2YwjYcfA8OqE2XEiGkCAxYwuQNMVUU6TQztpdWNGJxZDNkGDNeA6Xv1yzECRn8xXqz2exsjhhwEzANM47sZRsAY7BDz5aVrnFmMLRiuBtN+9Eow3PNn9YtSgQYDFzC544bMATBsskN8fVUJRQzBB0ynlxEwE3aI52eMBzDCva4uSF0H7wfDh81FqJuXVGFMe8nDDKbr4P1gRneIhYoSRsSQm4AhGTAEBMaqKLVEDDmYwfQduxsMH4dRTSKjmtQSMXZguAZM7n9BQ/dO72EH5uZZJXAwvgOmEgxjgTl/FUixWVlXgIEKmJFDS05hxMBcrpQKQMRYgWHpgOm/H8YUzOVKCRsMQMAMvmFKbNtO9G5gkL2YajDLVKTue3qBwQBEDDkImO7T1LebYgwGPGJswHBTwPQPTteoszEYwoqYGzCAU96Bk8Q9g852O72Pz11Kht2KIdyKNOPI3BMS1mAIa2UNAIbVwFyI4dmrqy3BAFWkWQfm5iFn44C5EiN3r2ZlTSL0gJlwuOaEMH8mGzhiNgDT8circcBcioECc/qHXXt5P2hLOpi/V6YgLJRe1iTaIGBmb+BovCTkYhZjGzH6YLgKjM63H3aJ0XmpzOmNMUhg1q1IE8UovYTo/eV3eDWJMAOGDAPm0auL5ThrvIMetCYRckUyCBj+nNeeLscPlVfQn85iDLZiTsDAVaQ5y6LT2vLz6Px1JcSvlkesVZ7O7qQqpjWJgCtSHg+Kpi1b/hkmz+y4b09HIGIsaxLhVqT+83I0ND5HVvlQpBSZ84vW6GCcBQwfR4cYrlgRse5GHsBu7yUYwgFTRO/LfKXi/A1l/LfjFo+1Wdekn5MYwl0jiVejp9E+RptqxFhOYnTBiAeMlRfldwxBgAGqSN2/RXZeRMRkwJpEqItqb15ExJSTjRjDmkSAFan3d8jWi4SY1zWpWNYkAlxUl76jWnuREaMMhqrBgE1hPHoREGO/d/d9EkOGFWkqGO4b36leBPZjTsTYTWIAwRS3XgTE+AOjHzBdr87pG1we1CH/fdj2WzEnYAgKTDbyws0zFhYWk80j5tskhiCnMEYFqfEHvTw0L16T9MBwPZiiEDAvvXSA+fFHdGoSLJhVA4ZfjzT3LKLFI+bVAwRWOzGIYHQChmsKTN2ui3DEvL5T0xaM6ynMz2lER0GiATAsGTHF+mGTr2IIcAozGDB8Hzk/CxKNgPn2x1SmvQUTDDkAw+37cWdeesGQbMQYf38SCBipilQxqeHTiWovGMmI0S5JNWAICcxYReIKMKcB0w2G1MEYJAz9BAMwhSmDYCqWTXy+rukGQ5I1Cek+Tf9gXi1R2gKGZ4M5tgXjb85bsTHDF3PUfjCsDcZo1qsEhuvBTKhI1BQwNAWMYE3KJcBIzXn5HsxVwIyAYeVJjCWYdbbtKsBcroEngTnWAvMsZjswlwEzAoaUF9Yl44GxmPPSEJj7Ocz1gmYWGJZPGKOtOwKb8xZ5MNzxfoYaMNqTmADTB6ZtH4Z/BMysSa9oTcICQ0Bghp8XuH55x4uAOQJMy8UBwprz9rSTFwRxDZg//+gADMwXn1+DMQiYrp9e/3AQ/5ybHjJgFLbuJF9KuxeY6zei8vcl0zQwyrPeADOpJl28VvX7uL79Q4DZDUzzKH9dMgUYNDAsDKbtqQF+ft03zwRD344ivNdrBobcg+FWMfJgju+zo/mzXgsw9AwmuwVDrc9E87McOTATxRTVi0kX62qymcLMfjwQFMzhFQyhgZn+AXte0/HwIQnmCDCgYFqGmj7AHOJgWA5MXhUMq4BpGOvHf8saYI4AgwqmfrQ/wPBkMPzi7VQsM4ABRo/MZ7LMBkMvLnIGGGQwVSNOH2CO2WD4xW03JCfGBswC2zDfj3f3nstDCsxzrPAhPolRBPO8dbfANszro55djvwWAjPBPH3jrG8wtB2Y2y0bITA/fhQHGM9gmH+sYwIMFhg2AMMfJejrMV/erBlgAgxXPDAQYALMZmAowIiCefHAgCMwuvfDBJhXYI4AE2A2AhMlSRUMB5gAczfpDTATwJjs2wGskl7v3bfeTb4umFsQ64M5fj7DFmACzAWY528YoNf3HwSYAPN10sHf7nG4egP82CQmwEz0YgTmeHToa5YEmADDt1Wk4ktJnm+j4QCzJ5i3N4Hw6XtXz1lw8yRGHAwFGA0wlzlx8e8DzG5gTr8X5wYMB5iNwfAlGG4EcyiCKXuDYT9gKv5dgAkwASbA3IM5PIGhAAMNhgJMgBkGwwEmwFTsDQeYAAMPpnLOywFGDwwvACYSBgYMm4Kp9bJswrAXMJX/NsBsCOYIMAFmFTCViyQOMDBgyAUYCjAB5hUY2g4MLwGGA8y+YO7uwjQFoz6FCTCuwZR9wFCAUaxIASbABBhUMId7MLw0GHYBpuaHKoDRD5gAIwLm0AGTA0yAAa9IF19PseC6enEwZDmFsQPDASbABBgJMLlUVqQAE2CMAibA3IK5faFHgNEBYyFmRzAcYMDANG71zpzDBJgAI7DNuz4YDjCoFSnABJgAs2dJIlQwOcBggKkOGMkpzBMYw40YMTHTwBAcGLIB85kvacWatBQYi4CBBcMBZtoMJsDogjkcg+EAE2BMKxIuGA4wiBUpwEiBOaTBFJOK1AUmBxgAMNUBIzqFeQlmoYhZB0zGCpg3MCnAOAFD5mASABgOMGNgOMAEmPaKtBQYdTHLgCk2FSnArJAwpBcwnWBygDEGU4zA5DowVpMYDjD9YFQD5gEmBZiqjX91MBmvIimB0RazCJhS7SXABBjDKW83GLWaxAEGOWACTIAZA2M+6539iZcAY+blqiJpgaEA0wOmVC2ptwTDAaY/YHDAKNYkDjCtYLS8mIChACO1pLYEY1mTOMCcjJqyl6sLA3/BpABjBObsg2dAMAkMDO8F5vqzN3iBArNQxECBufvwBcgLLpjZmzE8DIZEwFR8+IwOBqMmSdxKxfPA8BQwNR/eMmCuK5IiGDIB8zl0fFWqusG0Pipb9+nL7fmHBQMYMV1/lI8zFhcF5/vPHH1Pb+2HL6VUe1GuSGhgePgH3JxBbvkGvxlgerYuS31Bkg+YGzCECYab21kvzkf8XgM3fyMb9xRk04C53rb7BAMyiXl5Ari/3Yq54sC3+cTXWcZdHz03TGAEZn3XAXMPRjtieBqWs5N6miF8r6G6INHzt9S39fPjzBfCq0iqYJojhqe110ttvgoQvp0D8czfk48+FlMv7WDMaxLP59KwArvFwI1emvtXGrwoVqQfYGAmMY+zwBKtto/D/0kH+o4VksS+1U3AVICxEMNyjRRbR79KaZjxGlQkTDDCDZDL05LatCD1gKHlwaiY6etOyw6MakV6AWajiJE309kV/vOkQLk65dYBszEYwYud3b1omsDAgKF9wIi8OaK/B/n9SaRi5uW+Ij2D0VhYA4qZ/LzCyMEf9ajU/XBNMKkJTF4czLyzP3bctwlMOfeiHzAAYGDFjA/B6CHfJjDVXpDA0KZghsZh+GD8/iy1pZeKKcwXMFvXpLHhmHAQQC8vAkYfjBMxLYMy5ce/F6RaL2BgKMBUD82kn/u+xVvIMmBqKtJXMBExjXDm/ThEL68CJsCM3Yk1S1+bFzwwojXJt5gvwzXNSwHwUlWRvoGJiGm9WjjNS6n3YhkwJmAWEDNyH+nP8wHhZQxMREztuM/xkoG9XIKJiJEW88LL5fUjsIAJMLpg6JWXjOBlFEyIkRDzapzK5fzF3ssNmIiY5vEf2/tr9GIdMFZgdowYQvYyDiYiZjIYp15uwUTEiIChMy8E4aU6YC7AhJiJYE68lDYvAAETYFQmvehehsBo1STaJmDOTgCOl/qKdAUmImbO5aRTLxnfCxQY2txL27k5cMGc16QQ03yHA3V5afpJAl7OK9IlmIiYCgcLeGkJmJdgImLqFfRxaT8rBBIw12DEI4YwRr2hNy1/womXpoBxCoZkwEztPXnxMgEMvJgp4TTe92M/L+Zghh4blC1Qd/0+lLwcDsDoTXtHn9TQmdy86LQWF/2AuZry3oLBjZgJTwf18Hn/nT8uuUz1Ip0vjQFzBkYxYqY8DSaC5ThrQnMji3xpDZhbMMARI/A1Fiazoev+E1bAdIABEsMEZ8a9l14wmhEz8aF2Ci+yAdMDBluMJZrpn5/gAuYcjGrE0OxRcsXlrLcH4QVMFxioiDkfKO/xYuFlBIzmynpwdAGubQp8crIEkwbAeBajhEbiYxNkwFyB0Y0YEhyy8DItYKrA+BdDoN/xxobzl76AQQJD0gPnJ15UvMwHs5wYsvmq2vae4HrBAkM6w4fORelLKQXA+BPDagcSmrwo3M8w5GU5MKx2MKFPqeVFBoxHMaxyQKmPiO5lAAyuGPkvrxH8fGZe5oDRjxiz6cVOXPoDph6MMzGMOKTswcsgGIuIIes1r8z9TIzjpT9g7sE4FqM4ABM+0uHCCyiYaXtrbrioepEF41sMu9Ci3M8BL4NgPIhhD1yMvcwFYxMxUy8TQmv5dRy6HRzxMgrGixjGDZdfylzGAqYKjJWYyfciQJYib14awegWJaS7V6Q6rl2ORvbs6sGYRYzR2zzUOv2POpfRgKkEk6wiRuS+Spj+Hr/051aDXiaAcShGCk1TD/7372EwFc9KYNYTM91M6+EPl17awegXJclnRMz6Z8JleMbbAMYyYqSfRDPomg2XCQHTAWZBMX1q+o91AHqZDsa2KCk+Vy/eEysuMwpSCxjbokSrfHUxlJfmgtQHJsR45DKlIDWBuYwYFTHuyfxjx+V60JIIGPOi5FyMJZdJBakXjJ0Yx2SO36hexMAAFCW3ZA7TeJlVkLrBhBhXXG6GSw4MRFHyR8aay7yC1AwGoyj5ImPPZV5BGgBjLcYJmf8BcJlYkNrBAIkhD+ECwGWql3YwKNMYB2SO419CaDO9dIBJOBEDTQYiXCZPYAbBQIjBNHPAcJlbkLrAwIlBI1NwtEz30gUGahoDR+bPV0ISIYMZ8SIAJhudGRAtBeqtNNMDpg8MphhzMwdSKRLy0gkGVYwxGaLlvfSCgRVjZ4ZoBy9CYLLxmQotFSOjCwZcjJaZgopFyks/GOCiJI/m16/jz3oIF4tUQRoB40HMXzR/9kYmXn8+jn/AsYh5kQOTkU7e30Xv+LoZP1hEC9IQGE9intR0uPl1fP458tGkvAyB8SbmeVpzHMcNnuOzsScqsl7GwHgU813O70edKe82fpf//va//73/nzso4l6mgfEn5nxBRd7b3VjYgVlIzEJN0ssomLuiFGJAvEwqSONgQsxeXsbBhJitvEwAE2J28jIXTIhB9wIBJsRs5GUKmBCzj5c5YELMNl4mgQkxu3iZBSbEbOJlGpjbxXWIsfCSZ3uZBybEbOFFBEyIwfMCCCbE7OBlJpgQs4GXqWBCzPpe5oIJMct7mQymQkyQUeEi5WU2mBCzuJfpYELM2l7mgwkxS3sRABNiVvYiASa1TeGjyS2PJOIg2YgJMgrxIpIGKcSEF3MwIWZVL1JgqsQEmdlc5L2IgQkxa3qRAxNilvQiCKZOTJCZxkXFiySYELOgF1EwUZYWK0fyYCJk1ooXeTAhZjEv4mBCzFpe5MFUigky/Vw0vSiASRQhoxYvCr//CUZMkBmOF41f/5SUyYQYt+VID0yEzBrxogemWkyQaeBi4EUNTIhZw4semERBRpCL3i9+SilCJuIFEky9mCDTdobSomASBRkBLrq/9CklzJDJwQUuXgzAhBjfXvTBJAoyE7no/8KnlIBDJgcXpHgxApMoyEzhYvLbnuDF5OCC48UITJBxysUOTKIgM3YK0mZgWkMmb84FIl5swSQKMr2fPW0JpllM3pULjhdbMEHGGxdzMKn97OXNtPz4vGlrMB0hsxCZjg9rPlwpBZng4gtMT11yb6brM6YAsykZr1xgwKS+U5rX1fLzo6UAMx4y/sz0fiyYYUopyAQXn2D6yeSltABzQQOT+k9yXkUL6uQFE8wQmbyAFnAugGDGyACiGes93uik5ICMWzOD/UYcm5QWJAOBZrjHmCOTvIhpHQBTNDP6mgKMOhkTNHN6CTssKXki0zMaimqm9Q54UFLagYy8mpndgh6SlLYhI6Vmcn/AByTBt8ljNNGNQC/wRyMln2RGR2sEjtihPYxFSluTqbcjfzQfI5GSYzJzR1GrkV8ujsCsYsa1FmdgzsiQcy2uxsAXmFMy5FeLsxHwBsarmTW0+ARzTgYUzXl3PZ775LO5MbOUFsdgLs0QPBa/p90xmGszhIvF9Tn3DebGDAFi8X7C3YO5M6Os5qYvC5zstETTvRLUeTVqjTOdVmlal5k7L3ovc5rTSk3x9oSW2yOWOsVpsaZyX0vLQVY7v2nBJnJTVM8PXfHcpkWb+XNsq57XtHILKwEGHc3ypzPt0YJKgAFSs9M5TPu1kBJglOVsfM7+D+F2Tx2n9JyMAAAAAElFTkSuQmCC';
var DEFAULT_LOGO_PNG_ = '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';

var DEFAULT_SETTINGS = {
  companyName: 'CMCC Home Care by CMNH',
  companyAddress: 'เลขที่ 199/414 หมู่ 2 หมู่บ้านสวนนนทรี ต.หนองจ๊อม อ.สันทราย จ.เชียงใหม่ 50210',
  companyPhone: '080-124-8695, 062-353-0909',
  companyLine: '',
  taxId: '0105556026571',
  bankInfo: 'ธนาคารกสิกรไทย สาขามีโชคพลาซ่า (สันทราย)\nชื่อบัญชี บจก.ซีเอ็มเอ็นเอช 2012\nเลขที่บัญชี 065-3-35069-4',
  brandColor: '#0B8587',
  signerName: 'นางกรรณิการ์ วิฑูรย์',
  signerTitle: 'กรรมการบริษัท',
  notifyEmail: 'chiangmailcmnh@gmail.com',
  signPageUrl: '',
  signValidDays: '7',
  requireSignCode: '1',
  sharePublic: '1',
  docFooter: 'เอกสารนี้ไม่ใช่ใบกำกับภาษี',
  contractTitleMonthly: 'หนังสือสัญญาว่าจ้าง (รายเดือน)',
  contractTitleDaily: 'หนังสือสัญญาว่าจ้าง (รายวัน)',
  contractTitleConsent: 'ใบยินยอมรับบริการ',
  contractTitleAddendum: 'บันทึกข้อตกลงแนบท้ายสัญญา',
  termsMonthly: '',
  termsDaily: '',
  termsConsent: '',
  pdpaText: '',
  catalog: '',
  officerPerms: '',
  legalName: 'บริษัท ซีเอ็มเอ็นเอช 2012 จำกัด',
  providerParty: 'บริษัท ซีเอ็มเอ็นเอช 2012 จำกัด โดย นางกรรณิการ์ วิฑูรย์ กรรมการบริษัท ในนาม Chiangmai Excellent Caregiver Center (CMCC Home Care by CMNH)',
  requireSignerId: '1',
  contractTitleEmployment: 'สัญญาจ้างพนักงานเฝ้าไข้',
  employmentPosition: 'พนักงานเฝ้าไข้',
  termsEmployment: '',
  pdpaEmployment: '',
  profileConsentText: ''
};

var _settingsCache = null;

function readSettingsRaw_() {
  var map = {};
  var sh = SS.getSheetByName('Settings');
  if (!sh || sh.getLastRow() < 2) return map;
  sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (r) { if (r[0] !== '') map[String(r[0])] = r[1]; });
  return map;
}

function getSettings_() {
  if (_settingsCache) return _settingsCache;
  var raw = readSettingsRaw_(), out = {};
  Object.keys(DEFAULT_SETTINGS).forEach(function (k) {
    var v = (raw[k] === undefined || raw[k] === '') ? DEFAULT_SETTINGS[k] : String(raw[k]);
    out[k] = v;
  });
  if (!out.termsMonthly) out.termsMonthly = DEFAULT_TERMS_MONTHLY;
  if (!out.termsDaily) out.termsDaily = DEFAULT_TERMS_DAILY;
  if (!out.termsConsent) out.termsConsent = DEFAULT_TERMS_CONSENT;
  if (!out.pdpaText) out.pdpaText = DEFAULT_PDPA;
  if (!out.catalog) out.catalog = JSON.stringify(DEFAULT_CATALOG);
  if (!out.termsEmployment) out.termsEmployment = DEFAULT_TERMS_EMPLOYMENT;
  if (!out.pdpaEmployment) out.pdpaEmployment = DEFAULT_PDPA_EMPLOYMENT;
  if (!out.profileConsentText) out.profileConsentText = DEFAULT_PROFILE_CONSENT;
  if (!out.legalName) out.legalName = out.companyName;
  if (!out.providerParty) out.providerParty = out.legalName === out.companyName ? out.companyName : out.legalName + ' ในนาม ' + out.companyName;
  if (!out.officerPerms) out.officerPerms = DEFAULT_OFFICER_PERMS_.join(',');
  out.driveFolderId = raw.driveFolderId ? String(raw.driveFolderId) : '';
  out.logoFileId = raw.logoFileId ? String(raw.logoFileId) : '';
  out.watermarkFileId = raw.watermarkFileId ? String(raw.watermarkFileId) : '';
  _settingsCache = out;
  return out;
}

function setSetting_(key, value) {
  var sh = SS.getSheetByName('Settings') || (ensureSystemSheets_(), SS.getSheetByName('Settings'));
  var last = sh.getLastRow();
  if (last >= 2) {
    var keys = sh.getRange(2, 1, last - 1, 1).getValues();
    for (var i = 0; i < keys.length; i++) {
      if (String(keys[i][0]) === key) { sh.getRange(i + 2, 2).setValue(value); _settingsCache = null; return; }
    }
  }
  sh.appendRow([key, value]);
  _settingsCache = null;
}

function getCompanyName_() { return getSettings_().companyName; }

function settingsForAdmin_() {
  var st = getSettings_(), out = {};
  Object.keys(DEFAULT_SETTINGS).forEach(function (k) { out[k] = st[k]; });
  var raw = readSettingsRaw_();
  return { success: true, settings: out, logo: getLogoDataUrl_(), watermark: getWatermarkDataUrl_(), productName: PRODUCT_NAME,
    setup: { hasLogo: !!st.logoFileId, pdfTestedAt: raw.pdfTestedAt ? String(raw.pdfTestedAt) : '' } };
}

function settingsSave_(d) {
  var s = d.settings || d;
  if (s.catalog !== undefined && s.catalog !== '') {
    try { JSON.parse(s.catalog); } catch (e) { fail_('ราคา/แพ็กเกจ (catalog) ไม่ใช่ JSON ที่ถูกต้อง: ' + e.message); }
  }
  if (s.brandColor !== undefined && s.brandColor && !/^#[0-9a-fA-F]{6}$/.test(s.brandColor)) fail_('สีแบรนด์ต้องเป็นรหัสสี เช่น #2F6F62');
  Object.keys(s).forEach(function (k) {
    if (DEFAULT_SETTINGS[k] === undefined || k === 'officerPerms') return;
    var v = String(s[k] === null || s[k] === undefined ? '' : s[k]);
    if (v.length > 45000) fail_('ข้อความในช่อง ' + k + ' ยาวเกินไป');
    setSetting_(k, v);
  });
  return settingsForAdmin_();
}

function settingsSaveLogo_(d) {
  var m = String(d.dataUrl || '').match(/^data:(image\/(png|jpeg|webp));base64,(.+)$/);
  if (!m) fail_('รองรับเฉพาะรูป PNG, JPG หรือ WEBP');
  if (m[3].length > 700000) fail_('ไฟล์โลโก้ใหญ่เกินไป (ไม่เกิน ~500KB)');
  var st = getSettings_();
  var file = driveFolder_('branding').createFile(Utilities.newBlob(Utilities.base64Decode(m[3]), m[1], 'logo'));
  if (st.logoFileId) { try { DriveApp.getFileById(st.logoFileId).setTrashed(true); } catch (e) { /* ข้าม */ } }
  setSetting_('logoFileId', file.getId());
  // ลายน้ำวงกลม (หน้าเว็บแอดมินสร้างจากโลโก้ให้อัตโนมัติ) ใช้เป็นพื้นหลังหน้าประวัติเจ้าหน้าที่
  var w = String(d.watermarkDataUrl || '').match(/^data:image\/png;base64,(.+)$/);
  if (w && w[1].length < 900000) {
    var wf = driveFolder_('branding').createFile(Utilities.newBlob(Utilities.base64Decode(w[1]), 'image/png', 'watermark.png'));
    if (st.watermarkFileId) { try { DriveApp.getFileById(st.watermarkFileId).setTrashed(true); } catch (e) { /* ข้าม */ } }
    setSetting_('watermarkFileId', wf.getId());
  }
  CacheService.getScriptCache().remove('logo_dataurl');
  CacheService.getScriptCache().remove('wm_dataurl');
  return settingsForAdmin_();
}

function getLogoBlob_() {
  var id = getSettings_().logoFileId;
  if (id) { try { return DriveApp.getFileById(id).getBlob(); } catch (e) { /* ไฟล์หาย ใช้โลโก้ตั้งต้น */ } }
  return Utilities.newBlob(Utilities.base64Decode(DEFAULT_LOGO_PNG_), 'image/png', 'logo.png');
}

function getLogoDataUrl_() {
  var cache = CacheService.getScriptCache();
  var hit = cache.get('logo_dataurl');
  if (hit) return hit;
  var b = getLogoBlob_();
  var url = 'data:' + b.getContentType() + ';base64,' + Utilities.base64Encode(b.getBytes());
  if (url.length < 95000) cache.put('logo_dataurl', url, 21600);
  return url;
}

function getWatermarkDataUrl_() {
  var cache = CacheService.getScriptCache();
  var hit = cache.get('wm_dataurl');
  if (hit) return hit;
  var id = getSettings_().watermarkFileId, url = '';
  if (id) { try { var b = DriveApp.getFileById(id).getBlob(); url = 'data:image/png;base64,' + Utilities.base64Encode(b.getBytes()); } catch (e) { url = ''; } }
  if (!url) url = 'data:image/png;base64,' + DEFAULT_WATERMARK_PNG_;
  if (url.length < 95000) cache.put('wm_dataurl', url, 21600);
  return url;
}

function publicBranding_() {
  var st = getSettings_();
  return { success: true, companyName: st.companyName, companyPhone: st.companyPhone, brandColor: st.brandColor, productName: PRODUCT_NAME, logo: getLogoDataUrl_() };
}

function driveFolder_(sub) {
  var st = getSettings_();
  var root;
  try { root = st.driveFolderId ? DriveApp.getFolderById(st.driveFolderId) : null; } catch (e) { root = null; }
  if (!root) { root = DriveApp.createFolder(PRODUCT_NAME + ' - เอกสาร'); setSetting_('driveFolderId', root.getId()); }
  if (!sub) return root;
  var it = root.getFoldersByName(sub);
  return it.hasNext() ? it.next() : root.createFolder(sub);
}

function log_(action, by, ref, detail) {
  try {
    var sh = SS.getSheetByName('Log');
    if (sh) sh.appendRow([nowStr_(), action, by || '', String(ref || ''), String(detail || '')]);
  } catch (e) { /* log พลาดไม่ให้งานหลักล้ม */ }
}

/* ============================== ตัวช่วยทั่วไป ============================== */

function nowStr_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'); }
function todayStr_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
var TH_MONTHS_SHORT_ = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];
function thDate_(s) {
  if (!s) return '-';
  var p = String(s).slice(0, 10).split('-');
  if (p.length !== 3) return String(s);
  return (+p[2]) + ' ' + TH_MONTHS_SHORT_[+p[1] - 1] + ' ' + (+p[0] + 543);
}
function thDateTime_(s) { return s ? thDate_(s) + ' ' + String(s).slice(11, 16) + ' น.' : '-'; }
function fmt0_(n) {
  var v = Math.round(Number(n || 0) * 100) / 100;
  var parts = v.toFixed(v % 1 === 0 ? 0 : 2).split('.');
  return parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',') + (parts[1] ? '.' + parts[1] : '');
}
function esc_(s) { return String(s === undefined || s === null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }
function br_(s) { return esc_(s).replace(/\r?\n/g, '<br>'); }
function sha256hex_(str) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, str, Utilities.Charset.UTF_8);
  return bytes.map(function (b) { return ('0' + (b < 0 ? b + 256 : b).toString(16)).slice(-2); }).join('');
}

function bahtText_(n) {
  n = Math.round(Number(n) * 100) / 100;
  if (isNaN(n)) return '';
  var neg = n < 0; n = Math.abs(n);
  var baht = Math.floor(n), sat = Math.round((n - baht) * 100);
  if (sat === 100) { baht += 1; sat = 0; }
  var out = (baht === 0 ? 'ศูนย์' : thaiInt_(baht)) + 'บาท';
  out += sat === 0 ? 'ถ้วน' : thaiInt_(sat) + 'สตางค์';
  return (neg ? 'ลบ' : '') + out;
}
function thaiInt_(n) {
  var D = ['ศูนย์', 'หนึ่ง', 'สอง', 'สาม', 'สี่', 'ห้า', 'หก', 'เจ็ด', 'แปด', 'เก้า'];
  var P = ['', 'สิบ', 'ร้อย', 'พัน', 'หมื่น', 'แสน'];
  var s = String(n), groups = [];
  while (s.length > 6) { groups.unshift(s.slice(-6)); s = s.slice(0, -6); }
  groups.unshift(s);
  var out = '';
  groups.forEach(function (g, gi) {
    var len = g.length, txt = '';
    for (var i = 0; i < len; i++) {
      var dgt = +g[i], p = len - i - 1;
      if (!dgt) continue;
      var higher = gi > 0 || /[1-9]/.test(g.slice(0, i));
      if (p === 0 && dgt === 1 && higher) txt += 'เอ็ด';
      else if (p === 1 && dgt === 2) txt += 'ยี่';
      else if (p === 1 && dgt === 1) { /* สิบ */ }
      else txt += D[dgt];
      txt += P[p];
    }
    out += txt;
    if (gi < groups.length - 1) out += 'ล้าน';
  });
  return out;
}


/* ============================== ข้อความสัญญาตั้งต้น (ร่าง — ให้ทนายตรวจก่อนใช้จริง) ==============================
 * ขึ้นบรรทัดใหม่ = ย่อหน้าใหม่ · ขึ้นต้นด้วย "# " = หัวข้อ · ช่อง "(ใส่ข้อมูล...)" ต้องแก้ให้ครบ ระบบจะไม่ให้ส่งสัญญาถ้ายังเหลือ
 * ตัวแปรที่ใช้ได้: {{companyName}} {{companyPhone}} {{clientName}} {{address}} {{contactName}} {{contactRelation}}
 *   {{packageName}} {{serviceFormat}} {{startDate}} {{contractEndDate}} {{monthlyFee}} {{dailyRate}} {{admitRate}}
 *   {{assessmentFee}} {{deposit}} {{depositText}} {{dailyPkgRate}} {{dailyPkgDays}} {{dailyPkgTotal}}
 *   {{dailyPkgRange}} {{serviceItems}} {{serviceTotal}} {{serviceTotalText}}
 */
var CMCC_TERMS_ = "# คู่สัญญา\nสัญญาว่าจ้างฉบับนี้จัดทำขึ้น ณ {{address}} ระหว่าง {{payerName}} ซึ่งต่อไปนี้ในสัญญาฉบับนี้เรียกว่า \"ผู้ว่าจ้าง\" ฝ่ายหนึ่ง กับ {{providerParty}} ซึ่งต่อไปนี้เรียกว่า \"ผู้รับจ้าง\" หรือ \"ผู้ให้บริการ\" อีกฝ่ายหนึ่ง\nโดยรับจ้างดูแล {{clientName}} ซึ่งต่อไปนี้ในสัญญาฉบับนี้เรียกว่า \"ผู้ป่วย\" หรือ \"ผู้รับบริการ\" โดยมี {{caregiverNames}} ซึ่งต่อไปนี้เรียกว่า \"ผู้ดูแลผู้ป่วย\" (ประวัติและหลักสูตรที่จบตามเอกสารแนบท้าย) คู่สัญญาทั้งสองและผู้เกี่ยวข้องตกลงกันดังต่อไปนี้\n# ข้อ 1 หน้าที่ของผู้ดูแลผู้ป่วย\n\"ผู้ว่าจ้าง\" ตกลงว่าจ้างให้ \"ผู้ให้บริการ\" เป็นตัวแทนจัดหา \"ผู้ดูแลผู้ป่วย\" ให้แก่ \"ผู้รับบริการ\" โดย \"ผู้ดูแลผู้ป่วย\" มีหน้าที่ดังต่อไปนี้\n1. ดูแลสุขภาวะและความสะอาดของร่างกาย ป้อนอาหาร ให้ยา และดูแลความต้องการขั้นพื้นฐานในการดำรงชีวิตของ \"ผู้ป่วย\" ตามขอบเขตที่ระบุเฉพาะในสัญญาฉบับนี้\n2. ทำความสะอาดสิ่งแวดล้อมทั่วไปเฉพาะพื้นที่ที่ผู้ป่วยพักอาศัย (เตียงนอน ห้องที่ผู้ป่วยอยู่อาศัย ห้องน้ำที่ผู้ป่วยใช้เป็นประจำ)\n3. ดูแลและทำความสะอาดเสื้อผ้า เครื่องมือหรืออุปกรณ์ที่เกี่ยวข้องกับผู้ป่วย ได้แก่ เครื่องดูดเสมหะ อุปกรณ์ให้อาหารทางสายยาง แก้วยา ภาชนะรับประทานอาหาร และเวชภัณฑ์ที่เกี่ยวข้องกับผู้ป่วยเท่านั้น\n4. ดูแลผู้ป่วยตามระยะเวลาทำงานตามที่ระบุ ทั้งที่บ้านและโรงพยาบาล\n5. รายงานผลสุขภาพประจำวัน และอาการต่าง ๆ ของ \"ผู้ป่วย\" ให้ \"ผู้ว่าจ้าง\" รับทราบข้อมูล\nสิ่งที่ไม่เกี่ยวข้องกับการดูแลผู้ป่วย ห้าม \"ผู้ว่าจ้าง\" \"ผู้ป่วย\" หรือ \"ผู้ที่ไม่เกี่ยวข้อง\" ในสัญญานี้มอบหมายให้ \"ผู้ดูแลผู้ป่วย\" ทำ ได้แก่\n1. ห้ามให้ \"ผู้ดูแลผู้ป่วย\" ทำความสะอาดบ้าน ล้างห้องน้ำ เช็ดกระจก ล้างมุ้งลวด ทำความสะอาดศาลาภายในบ้าน ดูแลสัตว์เลี้ยง เว้นแต่ที่ได้ระบุไว้ตามข้อ 1\n2. ห้ามให้ \"ผู้ดูแลผู้ป่วย\" ซักเสื้อผ้าหรือทำความสะอาดเครื่องนุ่งห่มของบุคคลอื่นในบ้าน ยกเว้นเสื้อผ้าของผู้ป่วยเท่านั้น หากมีการปฏิบัติงานนอกเหนือจากที่ตกลง \"ผู้ว่าจ้าง\" จะชำระค่าแรงให้ \"ผู้ดูแลผู้ป่วย\" ตามค่าแรงงานขั้นต่ำรายวัน เพิ่มเติมจากค่าจ้างบริการดูแลผู้ป่วย\n3. ห้าม \"ผู้ว่าจ้าง\" ชำระเงินค่าจ้าง ให้เบิกเงินล่วงหน้า ให้กู้ยืมเงิน หรือให้ยืมทรัพย์สินมีค่าแก่ \"ผู้ดูแลผู้ป่วย\" หากฝ่าฝืน \"ผู้ให้บริการ\" จะไม่รับผิดชอบใด ๆ ทั้งสิ้น\n4. ห้าม \"ผู้ว่าจ้าง\" ติดต่อจ้างงานโดยตรงกับ \"ผู้ดูแลผู้ป่วย\" ที่ \"ผู้ให้บริการ\" จัดหามาให้บริการ หากละเมิด \"ผู้ว่าจ้าง\" ยินดีให้ \"ผู้ให้บริการ\" ยึดเงินค่าประกันตามจำนวนที่ตกลงกัน\n# ข้อ 2 ระยะเวลา รูปแบบการจ้าง และค่าบริการ\n\"ผู้ว่าจ้าง\" และ \"ผู้ให้บริการ\" เริ่มใช้สัญญาฉบับนี้มีผลตั้งแต่วันที่ {{startDate}} รูปแบบการจ้างงาน{{employmentMode}} {{formatText}}\n{{rateText}}\n{{depositClause}}\n# รายละเอียดเพิ่มเติม\n1. \"ผู้ดูแลผู้ป่วย\" ทำงาน {{workDaysPerWeek}} วันต่อสัปดาห์ และหยุด {{dayOffPerWeek}} วันต่อสัปดาห์ โดยวันหยุดคือ {{weeklyDayOff}}\n2. หาก \"ผู้ว่าจ้าง\" ต้องการให้ผู้ดูแลหรือผู้ดูแลทดแทนทำงานในวันหยุด และผู้ดูแลยินยอม ตกลงจ่ายค่าทำงานพิเศษรายวัน วันละ {{extraDayRate}} บาท โดยไม่รวมในค่าบริการรายเดือน\n3. กรณีจ้างงาน 24 ชั่วโมง ผู้รับบริการจัดที่พักอาศัยที่เหมาะสมให้แก่ผู้ดูแล\n4. ผู้รับบริการจัดหาอาหารให้แก่ผู้ดูแล: {{caregiverMeals}}\n# สิทธิของผู้ดูแลผู้ป่วย\n1. ได้รับการจัดสรรเวลาพักผ่อนที่เหมาะสม เช่น เวลานอน และเวลาทำกิจวัตรส่วนตัวจากผู้รับบริการ\n2. \"ผู้ว่าจ้าง\" ต้องไม่มอบหมายงานอื่นนอกเหนือจากการดูแล \"ผู้ป่วย\" เช่น เลี้ยงสัตว์ เก็บกวาดบ้าน ล้างรถ รดน้ำต้นไม้ ตัดหญ้า เช็ดกระจก ล้างห้องน้ำ ดูแลเด็กเล็ก ซักผ้าให้บุคคลอื่น และอื่น ๆ ซึ่งไม่ใช่การให้บริการผู้ป่วยที่ระบุในสัญญา\n3. ห้าม \"ผู้ว่าจ้าง\" \"ผู้ป่วย\" หรือบุคคลในครอบครัวและบุคคลที่เกี่ยวข้องกับ \"ผู้ว่าจ้าง\" ล่วงละเมิดทางเพศ \"ผู้ดูแลผู้ป่วย\" หาก \"ผู้ให้บริการ\" ได้รับแจ้งข้อมูลจาก \"ผู้ดูแลผู้ป่วย\" จะดำเนินคดีความทางกฎหมายทั้งแพ่งและอาญา\n4. ห้าม \"ผู้ว่าจ้าง\" \"ผู้ป่วย\" หรือบุคคลในครอบครัวและบุคคลที่เกี่ยวข้องกับ \"ผู้ว่าจ้าง\" ทำร้ายร่างกายและจิตใจ \"ผู้ดูแลผู้ป่วย\" ทั้งการกระทำและวาจา หาก \"ผู้ให้บริการ\" ได้รับแจ้งข้อมูลจาก \"ผู้ดูแลผู้ป่วย\" จะดำเนินคดีความทางกฎหมาย\n5. มีวันหยุดพิเศษนักขัตฤกษ์ 5 วันต่อปี (ไม่นับรวมวันหยุดประจำสัปดาห์) ได้แก่ 31 ธันวาคม, 1 มกราคม, 13 เมษายน, 14 เมษายน และ 15 เมษายน หาก \"ผู้ว่าจ้าง\" ต้องการให้ผู้ให้บริการจัดหาผู้ดูแลทดแทนในวันดังกล่าว ตกลงจ่ายค่าทำงานพิเศษรายวัน วันละ {{extraDayRate}} บาท โดยต้องแจ้งความประสงค์ล่วงหน้าอย่างน้อย 2 สัปดาห์ ผู้ว่าจ้างรับผิดชอบค่าตอบแทน \"ผู้ดูแลผู้ป่วย\" ทดแทน ซึ่งไม่รวมในค่าบริการรายเดือน\nหาก \"ผู้ว่าจ้าง\" ละเมิดหรือกระทำผิดข้อตกลงข้อใดข้อหนึ่ง \"ผู้ให้บริการ\" ถือว่า \"ผู้ว่าจ้าง\" ผิดสัญญาและสามารถริบเงินประกันการจ้างงานได้\n# ข้อ 3 ความเสียหายต่อทรัพย์สิน\nการกระทำใด ๆ ที่ \"ผู้ป่วย\" หรือ \"ผู้ว่าจ้าง\" ทำให้เกิดความเสียหายต่อทรัพย์สิน หรือทำความเสียหายใด ๆ แก่ \"ผู้ดูแลผู้ป่วย\" \"ผู้ว่าจ้าง\" จะเป็นผู้รับผิดชอบชดเชยให้ \"ผู้ดูแลผู้ป่วย\" ตามราคาความเสียหายจริงที่เกิดขึ้นและพิสูจน์ได้\n# ข้อ 4 กรณีผู้ดูแลผู้ป่วยหยุดปฏิบัติงาน\n1. กรณี \"ผู้ดูแลผู้ป่วย\" ลาออกจากงานหรือไม่สามารถปฏิบัติงานได้ \"ผู้ให้บริการ\" จะแจ้งเหตุเป็นข้อความให้ \"ผู้ว่าจ้าง\" ทราบ พร้อมจัดเตรียมผู้ดูแลทดแทนให้โดยเร็ว\n2. กรณี \"ผู้ดูแลผู้ป่วย\" หยุดปฏิบัติงานโดยมีการวางแผนล่วงหน้า \"ผู้ให้บริการ\" จะแจ้ง \"ผู้ว่าจ้าง\" ล่วงหน้า 3 วัน หาก \"ผู้ว่าจ้าง\" ยินดีให้ผู้ให้บริการจัดหาผู้ดูแลทดแทน ให้ดำเนินการได้ แต่หากผู้ว่าจ้างไม่ยินยอม จะหักค่าบริการรายเดือนในวันที่ \"ผู้ดูแลผู้ป่วย\" หยุดปฏิบัติงาน\n3. กรณี \"ผู้ดูแลผู้ป่วย\" หยุดปฏิบัติงานฉุกเฉิน เช่น ประสบอุบัติเหตุ ไม่สบาย หรือมีเหตุอันสมควรที่ไม่สามารถปฏิบัติหน้าที่ได้ \"ผู้ให้บริการ\" จะแจ้งผู้ว่าจ้างทราบ หากผู้ว่าจ้างต้องการผู้ดูแลทดแทน ผู้ให้บริการจะจัดเตรียมให้ โดยหากไม่มีผู้ดูแลทดแทน จะไม่คิดค่าบริการในวันนั้น\n# ข้อ 5 การยุติการว่าจ้าง\nหาก \"ผู้ว่าจ้าง\" \"ผู้ป่วย\" หรือ \"บุคคลที่เกี่ยวข้อง\" มีการกระทำอันไม่เหมาะสมต่อ \"ผู้ดูแลผู้ป่วย\" หรือมีพฤติกรรมที่ก่อให้เกิดความเสี่ยง \"ผู้ให้บริการ\" อาจยุติการให้บริการและดำเนินคดีตามกฎหมายทันที และจะไม่มีการคืนเงินประกันตามเงื่อนไขในสัญญาให้แก่ผู้ว่าจ้าง โดย \"ผู้ว่าจ้าง\" ต้องปฏิบัติดังนี้\n1. ชำระค่าบริการเต็มจำนวนในเดือนนั้น\n2. ชำระค่าบริการอื่น เช่น ค่าวันหยุด ค่าล่วงเวลา ค่าวันหยุดนักขัตฤกษ์ และอื่น ๆ\n# ข้อ 6 การบอกเลิกการจ้างงาน\n1. \"ผู้ว่าจ้าง\" ต้องแจ้ง \"ผู้ให้บริการ\" ล่วงหน้า 30 วัน เว้นแต่ผู้ป่วยถึงแก่กรรม\n2. กรณีผู้ว่าจ้างตกลงว่าจ้างบริการแบบรายเดือน แต่ขอสิ้นสุดการว่าจ้างก่อนครบสัญญาจ้างแบบรายเดือน ผู้ให้บริการจะคิดค่าบริการเป็นรายวัน วันละ {{extraDayRate}} บาท นับตั้งแต่วันที่เริ่มให้บริการตามสัญญา\n# ข้อ 7 การมีผลของสัญญาและการลงนามอิเล็กทรอนิกส์\nคู่สัญญาตกลงให้มีผลการชำระค่าบริการเมื่อเริ่มงาน และเห็นชอบตามข้อความทั้งหมดในสัญญาฉบับนี้ คู่สัญญาตกลงให้การลงลายมือชื่ออิเล็กทรอนิกส์ในสัญญานี้มีผลผูกพันตามกฎหมายว่าด้วยธุรกรรมทางอิเล็กทรอนิกส์ ให้ถือว่าไฟล์ PDF ที่ระบบจัดเก็บเป็นต้นฉบับ และ \"ผู้ว่าจ้าง\" ได้รับสำเนาทางอิเล็กทรอนิกส์";
var DEFAULT_TERMS_MONTHLY_GENERIC_ = [
  '# 1. บริการที่ตกลง',
  '{{providerParty}} (ต่อไปนี้เรียกว่า "ผู้ให้บริการ") ตกลงจัดผู้ดูแลไปให้บริการดูแล {{clientName}} (ต่อไปนี้เรียกว่า "ผู้รับบริการ") ตามแพ็กเกจ {{packageName}} รูปแบบ {{serviceFormat}} ณ {{address}} เริ่มตั้งแต่วันที่ {{startDate}} ตามแผนการดูแลที่ผู้ให้บริการประเมินร่วมกับผู้ว่าจ้าง',
  '# 2. ค่าบริการและการชำระเงิน',
  'ค่าบริการเดือนละ {{monthlyFee}} บาท ผู้ว่าจ้างชำระล่วงหน้าก่อนเริ่มบริการของแต่ละเดือน ภายใน (ใส่ข้อมูล: จำนวนวัน) วันนับจากวันที่ได้รับใบแจ้งหนี้',
  '# 3. วันหยุดของผู้ดูแล',
  'ผู้ดูแลมีวันหยุด (ใส่ข้อมูล: จำนวนวัน) วันต่อสัปดาห์ หากผู้ว่าจ้างต้องการผู้ดูแลแทนในวันหยุดหรือวันหยุดนักขัตฤกษ์ คิดค่าบริการเพิ่มตามอัตรารายวัน {{dailyRate}} บาทต่อวัน',
  '# 4. การใช้บริการไม่ครบเดือนและการยกเลิกบริการ',
  'หากผู้ว่าจ้างใช้บริการไม่ครบเดือน หรือยกเลิกบริการก่อนครบกำหนด ให้คิดค่าบริการตามจำนวนวันที่ใช้จริงในอัตรารายวัน {{dailyRate}} บาทต่อวัน รวมค่าประเมินหน้างาน {{assessmentFee}} บาท โดยผู้ว่าจ้างต้องแจ้งยกเลิกล่วงหน้าอย่างน้อย (ใส่ข้อมูล: จำนวนวัน) วัน',
  '# 5. กรณีผู้รับบริการเข้ารับการรักษาในโรงพยาบาล',
  'ระหว่างที่ผู้รับบริการเข้ารับการรักษาตัวในโรงพยาบาล หากผู้ว่าจ้างต้องการคงตัวผู้ดูแลเดิมไว้ ให้คิดค่าบริการร้อยละ 50 ของอัตรารายวัน ({{admitRate}} บาทต่อวัน) ตามจำนวนวันที่เข้ารับการรักษา',
  '# 6. เงินประกันสัญญา',
  'ผู้ว่าจ้างวางเงินประกันสัญญา {{deposit}} บาท ({{depositText}}) ผู้ให้บริการจะคืนเงินประกันภายใน (ใส่ข้อมูล: จำนวนวัน) วันหลังสิ้นสุดสัญญา โดยมีสิทธิ์หักค่าบริการค้างชำระหรือค่าเสียหายที่เกิดขึ้นจริง พร้อมแจ้งรายการที่หักให้ทราบ',
  '# 7. ขอบเขตงานของผู้ดูแล',
  'ผู้ดูแลให้บริการตามแผนการดูแล เช่น ช่วยกิจวัตรประจำวัน ดูแลความสะอาด การรับประทานอาหารและยาตามคำสั่งแพทย์ วัดและบันทึกสัญญาณชีพ ไม่รวมงานบ้านทั่วไปที่ไม่เกี่ยวกับผู้รับบริการ และไม่ทำหัตถการที่ต้องใช้ใบประกอบวิชาชีพ เว้นแต่ผู้ให้บริการจัดบุคลากรที่มีคุณสมบัติตามกฎหมายมาดำเนินการ',
  '# 8. การเปลี่ยนตัวผู้ดูแล',
  'ผู้ให้บริการอาจจัดผู้ดูแลแทนชั่วคราวเพื่อให้บริการต่อเนื่อง ผู้ว่าจ้างขอเปลี่ยนตัวผู้ดูแลได้โดยแจ้งเหตุผล และผู้ให้บริการจะจัดหาผู้ดูแลใหม่ภายใน (ใส่ข้อมูล: จำนวนวัน) วัน',
  '# 9. การไม่ว่าจ้างผู้ดูแลโดยตรง',
  'ระหว่างสัญญาและภายใน (ใส่ข้อมูล: จำนวนเดือน) เดือนหลังสิ้นสุดสัญญา ผู้ว่าจ้างตกลงไม่ว่าจ้างผู้ดูแลที่ผู้ให้บริการจัดมาโดยตรง หากฝ่าฝืน ผู้ว่าจ้างตกลงชำระค่าธรรมเนียม (ใส่ข้อมูล: จำนวนเงิน) บาท',
  '# 10. เหตุฉุกเฉินและข้อจำกัด',
  'ผู้ดูแลไม่ใช่แพทย์และไม่วินิจฉัยโรค เมื่อผู้รับบริการมีอาการฉุกเฉิน ผู้ดูแลจะโทรแจ้งหน่วยแพทย์ฉุกเฉิน 1669 และแจ้งผู้ว่าจ้างทันที ผู้ว่าจ้างยินยอมให้ผู้ให้บริการประสานการรักษาเบื้องต้นในกรณีดังกล่าว',
  '# 11. การบอกเลิกสัญญา',
  'คู่สัญญาฝ่ายใดฝ่ายหนึ่งบอกเลิกสัญญาได้โดยแจ้งเป็นลายลักษณ์อักษรหรือข้อความทางอิเล็กทรอนิกส์ล่วงหน้าอย่างน้อย (ใส่ข้อมูล: จำนวนวัน) วัน หากฝ่ายใดผิดสัญญาในสาระสำคัญ อีกฝ่ายบอกเลิกสัญญาได้ทันที',
  '# 12. การลงนามอิเล็กทรอนิกส์',
  'คู่สัญญาตกลงให้การลงลายมือชื่ออิเล็กทรอนิกส์ตามสัญญานี้มีผลผูกพันตามกฎหมายว่าด้วยธุรกรรมทางอิเล็กทรอนิกส์ และให้ถือว่าไฟล์ PDF ที่ระบบจัดเก็บไว้เป็นต้นฉบับสัญญา'
].join('\n');

var DEFAULT_TERMS_MONTHLY = CMCC_TERMS_;
var DEFAULT_TERMS_DAILY_GENERIC_ = [
  '# 1. บริการที่ตกลง',
  '{{providerParty}} (ต่อไปนี้เรียกว่า "ผู้ให้บริการ") ตกลงจัดผู้ดูแลไปให้บริการดูแล {{clientName}} ตามแพ็กเกจ {{packageName}} รูปแบบ {{serviceFormat}} ณ {{address}} จำนวน {{dailyPkgDays}} วัน {{dailyPkgRange}}',
  '# 2. ค่าบริการ',
  'อัตรารายวัน {{dailyPkgRate}} บาท รวมเป็นเงิน {{dailyPkgTotal}} บาท ผู้ว่าจ้างชำระล่วงหน้าก่อนเริ่มบริการ',
  '# 3. การเปลี่ยนแปลงและการยกเลิก',
  'ผู้ว่าจ้างแจ้งเลื่อนหรือยกเลิกบริการล่วงหน้าอย่างน้อย (ใส่ข้อมูล: จำนวนชั่วโมง) ชั่วโมง หากแจ้งช้ากว่านี้ ผู้ให้บริการมีสิทธิ์คิดค่าบริการ 1 วัน',
  '# 4. เหตุฉุกเฉินและข้อจำกัด',
  'ผู้ดูแลไม่ใช่แพทย์และไม่วินิจฉัยโรค เมื่อผู้รับบริการมีอาการฉุกเฉิน ผู้ดูแลจะโทรแจ้งหน่วยแพทย์ฉุกเฉิน 1669 และแจ้งผู้ว่าจ้างทันที',
  '# 5. การไม่ว่าจ้างผู้ดูแลโดยตรง',
  'ภายใน (ใส่ข้อมูล: จำนวนเดือน) เดือนหลังสิ้นสุดบริการ ผู้ว่าจ้างตกลงไม่ว่าจ้างผู้ดูแลที่ผู้ให้บริการจัดมาโดยตรง',
  '# 6. การลงนามอิเล็กทรอนิกส์',
  'คู่สัญญาตกลงให้การลงลายมือชื่ออิเล็กทรอนิกส์ตามสัญญานี้มีผลผูกพันตามกฎหมายว่าด้วยธุรกรรมทางอิเล็กทรอนิกส์ และให้ถือว่าไฟล์ PDF ที่ระบบจัดเก็บไว้เป็นต้นฉบับ'
].join('\n');

var DEFAULT_TERMS_DAILY = CMCC_TERMS_;
var DEFAULT_TERMS_CONSENT = [
  '# 1. บริการที่ขอรับ',
  'ผู้ลงนามขอรับบริการจาก {{providerParty}} สำหรับ {{clientName}} ดังนี้: {{serviceItems}} ณ {{address}}',
  '# 2. ค่าบริการ',
  'ค่าบริการรวม {{serviceTotal}} บาท ({{serviceTotalText}}) ชำระตามที่ผู้ให้บริการแจ้ง',
  '# 3. การรับทราบและยินยอม',
  'ผู้ลงนามได้รับคำอธิบายขั้นตอน ประโยชน์ และความเสี่ยงที่อาจเกิดขึ้นของบริการข้างต้นแล้ว และยินยอมให้บุคลากรของผู้ให้บริการดำเนินการ ผู้ลงนามแจ้งประวัติแพ้ยาและโรคประจำตัวของผู้รับบริการตามความจริงแล้ว',
  '# 4. เหตุฉุกเฉิน',
  'หากเกิดเหตุฉุกเฉินระหว่างให้บริการ ผู้ให้บริการจะโทรแจ้งหน่วยแพทย์ฉุกเฉิน 1669 และแจ้งผู้ลงนามทันที',
  '# 5. การลงนามอิเล็กทรอนิกส์',
  'การลงลายมือชื่ออิเล็กทรอนิกส์ในเอกสารนี้มีผลผูกพันตามกฎหมายว่าด้วยธุรกรรมทางอิเล็กทรอนิกส์'
].join('\n');

var DEFAULT_PDPA = [
  'ข้าพเจ้ายินยอมให้ {{legalName}} ({{companyName}}) เก็บรวบรวม ใช้ และเปิดเผยข้อมูลส่วนบุคคลและข้อมูลสุขภาพของผู้รับบริการ ได้แก่ อาการ โรคประจำตัว การใช้ยา สัญญาณชีพ บันทึกการดูแลประจำวัน รูปถ่ายขณะเช็คอินหน้าที่พัก และตำแหน่งที่ตั้งเมื่อมีการแจ้งเหตุฉุกเฉิน',
  'เพื่อวัตถุประสงค์ในการวางแผนและให้บริการดูแล รายงานอาการแก่ญาติที่ผู้ว่าจ้างแจ้งไว้ ประสานบุคลากรทางการแพทย์เมื่อจำเป็น และการเรียกเก็บค่าบริการ',
  'ผู้ให้บริการจะเก็บรักษาข้อมูลไว้ (ใส่ข้อมูล: จำนวนปี) ปีหลังสิ้นสุดบริการ ข้าพเจ้าขอเข้าถึง แก้ไข หรือถอนความยินยอมได้โดยติดต่อ {{companyPhone}} ทั้งนี้การถอนความยินยอมอาจทำให้ผู้ให้บริการไม่สามารถให้บริการต่อได้'
].join('\n');


var DEFAULT_TERMS_EMPLOYMENT = [
  '# ข้อ 1 ตำแหน่งและสถานะการจ้าง',
  'นายจ้างตกลงรับลูกจ้างเข้าทำงานในตำแหน่ง "{{position}}" และลูกจ้างตกลงทำงานภายใต้การบังคับบัญชา ระเบียบ ข้อบังคับการทำงาน มาตรฐานบริการ และคำสั่งอันชอบด้วยกฎหมายของนายจ้าง ตามประเภทการจ้าง วันเริ่มงาน และการทดลองงานที่ระบุในตารางด้านบน',
  '# ข้อ 2 สถานที่และรูปแบบการปฏิบัติงาน',
  'ลูกจ้างยินยอมปฏิบัติงาน ณ ศูนย์ของนายจ้าง บ้านพักของผู้รับบริการ โรงพยาบาล หรือสถานที่อื่นในพื้นที่ให้บริการตามที่นายจ้างมอบหมาย โดยการเปลี่ยนสถานที่หรือเคสต้องมีเหตุผลเกี่ยวกับการดำเนินงาน ตามรูปแบบงาน เวลาปกติ และวันหยุดที่ระบุในตารางด้านบน',
  '# ข้อ 3 หน้าที่และขอบเขตงาน',
  'ลูกจ้างมีหน้าที่ดูแลผู้ป่วยหรือผู้สูงอายุตาม Care Plan ใบมอบหมายงาน และระดับความสามารถที่นายจ้างรับรอง โดยรวมถึง',
  '• ดูแลสุขอนามัยร่างกาย การแต่งกาย การขับถ่าย การเคลื่อนย้าย และกิจวัตรประจำวันของผู้ป่วยอย่างสุภาพและปลอดภัย',
  '• เตรียมและช่วยรับประทานอาหารหรือให้อาหารตามแผนการดูแล รวมทั้งบันทึกปริมาณอาหารและน้ำ',
  '• ช่วยให้ยาตามรายการและเวลาที่นายจ้าง ผู้ประกอบวิชาชีพ หรือครอบครัวกำหนด โดยไม่ปรับยาเอง และรายงานเมื่อพบความผิดปกติ',
  '• สังเกตอาการ วัดและบันทึกข้อมูลพื้นฐานตามที่มอบหมาย รายงานอาการเปลี่ยนแปลง อุบัติเหตุ หรือเหตุฉุกเฉินทันที',
  '• ดูแลความสะอาดเฉพาะเตียง ห้องพัก พื้นที่โดยรอบ เสื้อผ้า และอุปกรณ์ที่เกี่ยวข้องกับผู้ป่วย',
  '• ดูแลอุปกรณ์ที่ได้รับมอบหมาย เช่น อุปกรณ์ให้อาหารทางสายยาง เครื่องดูดเสมหะ ภาชนะยา และเวชภัณฑ์ โดยปฏิบัติตามการฝึกอบรมและขอบเขตวิชาชีพ',
  '• อยู่ดูแลผู้ป่วยตลอดเวลาปฏิบัติงาน ส่งเวร และจัดทำรายงานประจำวันในระบบของนายจ้างให้ครบถ้วน',
  '• ประสานงานกับครอบครัว ทีมพยาบาล และผู้บังคับบัญชาตามช่องทางที่นายจ้างกำหนด',
  '# ข้อ 4 ข้อจำกัดและข้อห้ามในการปฏิบัติงาน',
  '• ห้ามทำหัตถการหรือการพยาบาลที่เกินคุณวุฒิ การฝึกอบรม หรือคำสั่งมอบหมาย',
  '• ห้ามทอดทิ้งผู้ป่วย หลับในเวลาที่ต้องเฝ้าระวัง ดื่มสุรา ใช้สารเสพติด หรือปฏิบัติงานในสภาพไม่พร้อม',
  '• ห้ามใช้ความรุนแรง ข่มขู่ ดูหมิ่น ล่วงละเมิด หรือแสวงหาประโยชน์จากผู้ป่วยและครอบครัว',
  '• ห้ามรับงานส่วนตัวจากผู้รับบริการโดยตรง รับค่าจ้าง เงินกู้ เงินล่วงหน้า ของมีค่า หรือทำธุรกรรมกับผู้รับบริการโดยไม่ได้รับอนุญาตจากนายจ้าง',
  '• ห้ามนำบุคคลภายนอกเข้าพื้นที่ปฏิบัติงาน หรือเผยแพร่ภาพ เสียง ข้อมูลสุขภาพ และเรื่องภายในครอบครัว',
  '• ห้ามทำงานแม่บ้านทั่วไป ดูแลสัตว์ ล้างรถ ทำสวน หรือทำงานของบุคคลอื่น เว้นแต่เป็นงานที่เกี่ยวข้องกับผู้ป่วยและนายจ้างอนุมัติ',
  'หากผู้รับบริการมอบหมายงานนอกขอบเขต ลูกจ้างต้องแจ้งผู้บังคับบัญชาและรอคำสั่งจากนายจ้าง ไม่ตกลงค่าตอบแทนหรือรับงานเพิ่มด้วยตนเอง',
  '# ข้อ 5 ค่าจ้างและสิทธิประโยชน์',
  'นายจ้างจ่ายค่าจ้างตามอัตราและกำหนดจ่ายที่ระบุในตารางด้านบน ค่าตอบแทนอื่น เช่น ค่าเดินทาง ค่าอาหาร ค่าเคส และเบี้ยเลี้ยง เป็นไปตามประกาศหรือใบมอบหมายงานของนายจ้าง ประกันสังคมดำเนินการตามหลักเกณฑ์และระยะเวลาที่กฎหมายกำหนด',
  '# ข้อ 6 เวลาพัก วันหยุด และการลา',
  'ลูกจ้างมีสิทธิได้รับเวลาพัก วันหยุดประจำสัปดาห์ วันหยุดตามประเพณี วันหยุดพักผ่อนประจำปี และการลาตามกฎหมาย การลาต้องแจ้งตามช่องทางและระยะเวลาที่กำหนด เว้นแต่เจ็บป่วยหรือมีเหตุฉุกเฉินซึ่งให้แจ้งโดยเร็วที่สุด พร้อมส่งหลักฐานเมื่อเข้าเงื่อนไขที่กำหนด',
  '# ข้อ 7 การลาออก ขาดงาน และการส่งมอบเคส',
  '• ลูกจ้างที่ประสงค์ลาออกต้องยื่นหนังสือล่วงหน้าไม่น้อยกว่า 30 วัน หรือเป็นไปตามรอบการจ่ายค่าจ้างและกฎหมาย เพื่อให้นายจ้างจัดบุคลากรทดแทนและส่งมอบงาน',
  '• ลูกจ้างต้องไม่ละทิ้งเคสหรือออกจากสถานที่ก่อนส่งมอบผู้ป่วยให้ผู้รับผิดชอบ เว้นแต่มีเหตุคุกคามต่อชีวิตหรือความปลอดภัย ซึ่งต้องติดต่อผู้บังคับบัญชาหรือหน่วยฉุกเฉินทันที',
  '• เมื่อสิ้นสุดกะหรือสิ้นสุดการจ้าง ลูกจ้างต้องคืนยา เอกสาร กุญแจ อุปกรณ์ เงินทดรอง และทรัพย์สินทั้งหมด พร้อมรายงานสถานะผู้ป่วยอย่างครบถ้วน',
  '# ข้อ 8 มาตรฐานบริการ วินัย และการประเมินผล',
  'ลูกจ้างต้องแต่งกายตามระเบียบ ตรงต่อเวลา รักษาความสะอาด สื่อสารสุภาพ บันทึกข้อมูลตามจริง เข้าร่วมการปฐมนิเทศและอบรม และยอมรับการนิเทศงานหรือประเมินสมรรถนะที่เกี่ยวข้องกับหน้าที่ นายจ้างอาจตักเตือน พักการมอบหมายเคส หรือดำเนินการทางวินัยตามข้อบังคับและกฎหมายเมื่อพบการฝ่าฝืน',
  '# ข้อ 9 ความปลอดภัยของผู้ป่วยและเหตุฉุกเฉิน',
  'เมื่อผู้ป่วยมีอาการผิดปกติ อุบัติเหตุ หายใจลำบาก สำลัก ชัก หมดสติ หรืออยู่ในภาวะเสี่ยง ลูกจ้างต้องช่วยเหลือตามความรู้และแผนฉุกเฉิน ติดต่อผู้บังคับบัญชา ครอบครัว และหน่วยแพทย์ฉุกเฉิน 1669 ตามลำดับความเร่งด่วน ห้ามปกปิด แก้ไขย้อนหลัง หรือรายงานข้อมูลอันเป็นเท็จ',
  '# ข้อ 10 การรักษาความลับและข้อมูลส่วนบุคคล',
  'ลูกจ้างต้องรักษาความลับของนายจ้าง ผู้ป่วย ผู้รับบริการ และครอบครัว รวมถึงข้อมูลสุขภาพ ภาพถ่าย ที่อยู่ รหัสเข้าถึงระบบ เอกสาร รายงาน ราคา และข้อมูลธุรกิจ ใช้ข้อมูลเฉพาะเพื่อปฏิบัติงานตามคำสั่ง และส่งคืนหรือลบเมื่อหมดความจำเป็นตามระเบียบ หน้าที่นี้ยังคงอยู่หลังสิ้นสุดการจ้าง ทั้งนี้ให้ดำเนินการตามกฎหมายคุ้มครองข้อมูลส่วนบุคคล',
  '# ข้อ 11 ทรัพย์สิน ความเสียหาย และการหักเงิน',
  'ลูกจ้างต้องใช้และเก็บรักษาทรัพย์สิน เครื่องแบบ อุปกรณ์ เอกสาร และเงินทดรองด้วยความระมัดระวัง หากเกิดความเสียหายต้องรายงานทันที การพิจารณาความรับผิดให้ยึดข้อเท็จจริง เจตนา ความประมาท และกฎหมาย นายจ้างจะไม่หักค่าจ้างหรือเงินอื่นใด เว้นแต่เป็นกรณีและดำเนินการตามเงื่อนไขที่กฎหมายอนุญาต',
  '# ข้อ 12 การสิ้นสุดสัญญาและการเลิกจ้าง',
  'การสิ้นสุดสัญญา การบอกกล่าวล่วงหน้า ค่าจ้างแทนการบอกกล่าว ค่าชดเชย และการจ่ายเงินเมื่อสิ้นสุดการจ้างให้เป็นไปตามกฎหมายและข้อบังคับการทำงาน นายจ้างอาจเลิกจ้างโดยไม่จ่ายค่าชดเชยได้เฉพาะกรณีที่กฎหมายกำหนดและเมื่อมีข้อเท็จจริงรองรับ การไม่ผ่านทดลองงานถือเป็นการเลิกจ้างและต้องดำเนินการตามสิทธิที่กฎหมายกำหนด',
  '# ข้อ 13 การลงนามอิเล็กทรอนิกส์',
  'คู่สัญญาตกลงให้การลงลายมือชื่ออิเล็กทรอนิกส์ในสัญญานี้มีผลผูกพันตามกฎหมายว่าด้วยธุรกรรมทางอิเล็กทรอนิกส์ ให้ถือว่าไฟล์ PDF ที่ระบบจัดเก็บเป็นต้นฉบับ และลูกจ้างขอรับสำเนาจากนายจ้างได้'
].join('\n');

var DEFAULT_PDPA_EMPLOYMENT = [
  '{{legalName}} ในฐานะนายจ้าง เก็บรวบรวมและใช้ข้อมูลส่วนบุคคลของลูกจ้าง ได้แก่ ชื่อ เลขประจำตัวประชาชน ที่อยู่ ช่องทางติดต่อ บัญชีรับเงิน รูปถ่าย ประวัติการศึกษา การอบรม และการทำงาน เพื่อการจ้างงาน การจ่ายค่าจ้าง ประกันสังคม ภาษี และการมอบหมายงาน ตามสัญญาจ้างและหน้าที่ตามกฎหมาย',
  'บริษัทเก็บรักษาข้อมูลตลอดระยะเวลาการจ้างและตามระยะเวลาที่กฎหมายกำหนดหลังสิ้นสุดการจ้าง ลูกจ้างขอเข้าถึง แก้ไข หรือสอบถามเกี่ยวกับข้อมูลของตนได้ที่ฝ่ายบุคคล โทร {{companyPhone}}'
].join('\n');

var DEFAULT_PROFILE_CONSENT = 'เปิดเผยชื่อ ชื่อเล่น รูปถ่าย อายุ การศึกษา ใบอนุญาต การอบรม ประสบการณ์ และทักษะ แก่ผู้รับบริการและญาติ เพื่อแนบท้ายสัญญาบริการ (ไม่บังคับ ไม่มีผลต่อการจ้างงาน และถอนความยินยอมได้ทุกเมื่อโดยแจ้งบริษัท)';

// ราคา/แพ็กเกจตั้งต้น (ตัวอย่าง) — แก้ได้ในหน้าตั้งค่าของแอดมิน
var DEFAULT_CATALOG = {
  _note: 'ราคาตั้งต้น — กรุณาตรวจและแก้ให้ตรงราคาจริงของ CMCC ในหน้า ตั้งค่า > ราคาและแพ็กเกจ',
  packages: {
    elderly:  { name: 'ผู้สูงอายุทั่วไป', tag: 'Elderly Care', suited: 'ช่วยเหลือตัวเองได้น้อย · อยู่ลำพัง · พักฟื้นหลังผ่าตัด',
                includes: ['ดูแลกิจวัตรประจำวัน อาบน้ำ เปลี่ยนเสื้อผ้า', 'เตรียมอาหารและดูแลการรับประทานยา', 'วัดสัญญาณชีพและรายงานอาการแก่ครอบครัว'],
                price: { '8': [16000, 18000], '12': [16000, 18000], '24': [18000, 20000] }, daily: { '8': 800, '12': 900, '24': 1300 } },
    feed:     { name: 'ผู้ป่วยให้อาหารทางสาย', tag: 'Tube Feeding', suited: 'NG Tube · PEG Tube',
                includes: ['เตรียมอาหารและให้อาหารทางสาย', 'ดูแลความสะอาดสายและอุปกรณ์', 'เฝ้าระวังการสำลัก'],
                price: { '8': [16000, 18000], '12': [17000, 19000], '24': [20000, 22000] }, daily: { '8': 800, '12': 900, '24': 1300 } },
    stroke:   { name: 'ผู้ป่วยโรคหลอดเลือดสมอง', tag: 'Stroke Recovery', suited: 'อ่อนแรงครึ่งซีก · พาร์กินสัน',
                includes: ['ฝึกการเคลื่อนไหวเบื้องต้นตามแผน', 'พลิกตะแคงตัว ป้องกันแผลกดทับ', 'ติดตามความก้าวหน้า'],
                price: { '8': [16000, 18000], '12': [17000, 19000], '24': [19000, 22000] }, daily: { '8': 800, '12': 900, '24': 1300 } },
    memory:   { name: 'ผู้ป่วยภาวะสมองเสื่อม', tag: 'Memory Care', suited: 'อัลไซเมอร์ · หลงลืม · เดินหลง',
                includes: ['เฝ้าระวังการเดินหลงและพลัดตกหกล้ม', 'กิจกรรมกระตุ้นความจำ', 'รายงานพฤติกรรมประจำวัน'],
                price: { '8': [17000, 18000], '12': [17000, 19000], '24': [19000, 22000] }, daily: { '8': 800, '12': 900, '24': 1300 } }
  },
  formats: [
    { key: '8', label: '8 ชั่วโมง (ไป-กลับ)' },
    { key: '12', label: '12 ชั่วโมง (ไป-กลับ)' },
    { key: '24', label: '24 ชั่วโมง (ค้างคืน)' }
  ],
  zones: { A: 'Zone A · ใกล้ (≤10 กม.)', B: 'Zone B · ระยะกลาง (10–20 กม.)', C: 'Zone C · ระยะไกล (20–35 กม.)' },
  procedures: {
    ng:       { name: 'เปลี่ยนสายให้อาหาร (NG Tube)', price: { A: 700, B: 1000, C: 1300 } },
    foley:    { name: 'เปลี่ยนสายสวนปัสสาวะ (Foley)', price: { A: 900, B: 1200, C: 1500 } },
    dressing: { name: 'ทำแผล', price: { A: 700, B: 1000, C: 1300 } },
    blood:    { name: 'เจาะเลือดที่บ้าน', price: { A: 600, B: 900, C: 1200 } }
  },
  courses: {
    ptot: { name: 'PT-OT Service Package (กายภาพ/กิจกรรมบำบัด)', homeZoneRate: { A: 0, B: 100, C: 200 }, singleSessionRate: 700,
          options: { s5: { name: '5 ครั้ง', sessions: 5, price: 3300 }, s10: { name: '10 ครั้ง', sessions: 10, price: 6400 } } },
    aging: { name: 'Active Aging Wellness Program (PTA)', homeZoneRate: { A: 0, B: 100, C: 200 }, singleSessionRate: 500,
          options: { s8: { name: '8 ครั้ง', sessions: 8, price: 3600 }, s12: { name: '12 ครั้ง', sessions: 12, price: 5200 } } },
    tms: { name: 'TMS / PMS (ที่ศูนย์ CMNH)', homeZoneRate: { A: 0, B: 0, C: 0 }, singleSessionRate: 0,
          options: { s10: { name: '10 ครั้ง', sessions: 10, price: 0 } } }
  },
  depositSuggest: { '24': 10000, other: 5000 },
  assessmentFee: 1000
};



/* ============================== อนุมัติราคา ==============================
 * พนักงานกรอกราคาต่ำกว่าราคาตั้งต้นในแคตตาล็อก → เคสเป็น "รออนุมัติ" ทำสัญญาไม่ได้จนกว่าเจ้าของอนุมัติ
 * คำนวณเกณฑ์ที่ฝั่งเซิร์ฟเวอร์เสมอ (ไม่เชื่อค่าที่หน้าเว็บส่งมา)
 */
var PRICE_FIELDS_ = ['monthlyFee', 'dailyRate', 'deposit', 'assessmentFee', 'dailyPkgRate', 'procedureTotal', 'courseTotalPrice'];
var PRICE_FIELD_HEADERS_ = { monthlyFee: 'ค่าบริการต่อเดือน (บาท)', dailyRate: 'อัตรารายวัน (บาท)', deposit: 'เงินประกันสัญญา (บาท)', assessmentFee: 'ค่าประเมินหน้างาน (บาท)',
  dailyPkgRate: 'อัตรารายวัน (แพ็กเกจรายวัน) (บาท)', procedureTotal: 'ค่าบริการหัตถการ (บาท)', courseTotalPrice: 'ราคาคอร์สรวม (บาท)' };

function priceIssues_(d) {
  var cat; try { cat = JSON.parse(getSettings_().catalog); } catch (e) { cat = DEFAULT_CATALOG; }
  var issues = [], type = String(d.serviceTypeLabel || '');
  function below(label, val, floor) { if (floor > 0 && num_(val) < floor) issues.push(label + ' ' + fmt0_(num_(val)) + ' บาท (เกณฑ์ขั้นต่ำ ' + fmt0_(floor) + ')'); }
  function findBy(obj, name) { var k = Object.keys(obj || {}).filter(function (x) { return obj[x].name === name; })[0]; return k ? obj[k] : null; }
  var fmt = (cat.formats || []).filter(function (f) { return f.label === d.serviceFormat; })[0];
  var zoneKey = Object.keys(cat.zones || {}).filter(function (z) { return cat.zones[z] === d.zoneLabel; })[0];
  if (type === 'แพ็กเกจรายเดือน' || type === 'แพ็กเกจรายวัน') {
    var pkg = findBy(cat.packages, d.packageName);
    if (!pkg || !fmt) { issues.push('แพ็กเกจหรือรูปแบบบริการไม่ตรงกับแคตตาล็อก'); return issues; }
    var daily = num_((pkg.daily || {})[fmt.key]);
    if (type === 'แพ็กเกจรายเดือน') {
      below('ค่าบริการต่อเดือน', d.monthlyFee, num_(((pkg.price || {})[fmt.key] || [])[0]));
      below('อัตรารายวัน', d.dailyRate, daily);
      var ds = cat.depositSuggest || {};
      below('เงินประกันสัญญา', d.deposit, num_(ds[fmt.key] != null ? ds[fmt.key] : ds.other));
    } else below('อัตรารายวัน', d.dailyPkgRate, daily);
    below('ค่าประเมินหน้างาน', d.assessmentFee, num_(cat.assessmentFee));
  } else if (type === 'หัตถการ') {
    if (!zoneKey) { issues.push('โซนบริการไม่ตรงกับแคตตาล็อก'); return issues; }
    var floor = 0;
    String(d.procedureListText || '').split(/\s*,\s*/).filter(Boolean).forEach(function (nm) {
      var pr = findBy(cat.procedures, nm);
      if (!pr) issues.push('หัตถการ "' + nm + '" ไม่อยู่ในแคตตาล็อก'); else floor += num_((pr.price || {})[zoneKey]);
    });
    below('ค่าหัตถการรวม', d.procedureTotal, floor);
  } else if (type === 'คอร์ส') {
    var co = findBy(cat.courses, d.courseProgramLabel), op = co ? findBy(co.options, d.courseOptionLabel) : null;
    if (!op || !zoneKey) { issues.push('คอร์สหรือโซนไม่ตรงกับแคตตาล็อก'); return issues; }
    below('ราคาคอร์สรวม', d.courseTotalPrice, num_(op.price) + num_((co.homeZoneRate || {})[zoneKey]) * num_(op.sessions));
  }
  return issues;
}

function applyPriceApproval_(d, s) {
  var prev = d.id ? findCase_(d.id) : null;
  // ค่าสถานะอนุมัติตั้งโดยเซิร์ฟเวอร์เท่านั้น
  delete d.priceStatus; delete d.priceIssues; delete d.priceApprovedBy; delete d.priceApprovedAt; delete d.priceNote;
  var keep = function () {
    d.priceStatus = prev ? prev['สถานะราคา'] : ''; d.priceIssues = prev ? prev['รายการที่ต่ำกว่าเกณฑ์'] : '';
    d.priceApprovedBy = prev ? prev['อนุมัติราคาโดย'] : ''; d.priceApprovedAt = prev ? prev['อนุมัติราคาเมื่อ'] : ''; d.priceNote = prev ? prev['หมายเหตุการอนุมัติ'] : '';
  };
  if (s.role === 'admin') {
    // เจ้าของเป็นผู้กำหนดราคาเอง ถือว่าอนุมัติ
    var wasPending = prev && (prev['สถานะราคา'] === 'รออนุมัติ' || prev['สถานะราคา'] === 'ไม่อนุมัติ');
    var issuesO = priceIssues_(d);
    if (issuesO.length || wasPending) { d.priceStatus = 'อนุมัติแล้ว'; d.priceIssues = issuesO.join(' · '); d.priceApprovedBy = s.name; d.priceApprovedAt = nowStr_(); d.priceNote = 'เจ้าของกำหนดราคาเอง'; }
    else keep();
    return;
  }
  var issues = priceIssues_(d);
  if (!issues.length) { d.priceStatus = ''; d.priceIssues = ''; d.priceApprovedBy = ''; d.priceApprovedAt = ''; d.priceNote = ''; return; }
  var samePrices = prev && PRICE_FIELDS_.every(function (k) { return num_(prev[PRICE_FIELD_HEADERS_[k]]) === num_(d[k]); }) &&
    String(prev['แพ็กเกจ']) === String(d.packageName || '') && String(prev['รูปแบบบริการ']) === String(d.serviceFormat || '');
  if (samePrices && prev['สถานะราคา'] === 'อนุมัติแล้ว') { keep(); return; }   // แก้ข้อมูลอื่น ราคาเดิมที่อนุมัติแล้วยังใช้ได้
  d.priceStatus = 'รออนุมัติ'; d.priceIssues = issues.join(' · '); d.priceApprovedBy = ''; d.priceApprovedAt = ''; d.priceNote = '';
  d._priceNewlyPending = !(prev && prev['สถานะราคา'] === 'รออนุมัติ' && samePrices);
}

function notifyPricePending_(d, s, caseId) {
  var st = getSettings_();
  String(st.notifyEmail || '').split(',').map(function (x) { return x.trim(); }).filter(Boolean).forEach(function (email) {
    try { MailApp.sendEmail(email, 'รออนุมัติราคา — ' + d.clientName, s.name + ' บันทึกเคส ' + d.clientName + ' (' + caseId + ') ด้วยราคาต่ำกว่าเกณฑ์\n' + d.priceIssues + '\n\nเข้าหลังบ้าน > ขาย / เคส > รออนุมัติราคา'); } catch (e) { /* ข้าม */ }
  });
}

function approvePrice_(d, s) {
  var sh = getSheet(SHEET_RECORDS), row = findRowIndexById(sh, d.id);
  if (row === -1) fail_('ไม่พบเคส');
  var r = findCase_(d.id);
  if (r['สถานะราคา'] !== 'รออนุมัติ') fail_('เคสนี้ไม่ได้รออนุมัติราคา');
  var note = String(d.note || '').trim().slice(0, 300);
  if (!d.approve && !note) fail_('กรุณาระบุเหตุผลที่ไม่อนุมัติ เพื่อให้พนักงานแก้ไขได้ถูก');
  var vals = { 'สถานะราคา': d.approve ? 'อนุมัติแล้ว' : 'ไม่อนุมัติ', 'อนุมัติราคาโดย': s.name, 'อนุมัติราคาเมื่อ': nowStr_(), 'หมายเหตุการอนุมัติ': note };
  Object.keys(vals).forEach(function (hname) { sh.getRange(row, RECORD_HEADERS.indexOf(hname) + 1).setValue(vals[hname]); });
  appendHistory(d.approve ? 'อนุมัติราคา' : 'ไม่อนุมัติราคา', d.id, r['ชื่อผู้รับบริการ'], s.name, new Date());
  return { success: true, id: d.id, priceStatus: vals['สถานะราคา'] };
}

/* ============================== ประวัติเจ้าหน้าที่ (ข้อมูลส่วนกลาง → แนบท้ายสัญญา) ============================== */

function findStaff_(id) {
  var rows = readSheetAsObjects(SHEET_STAFF, STAFF_HEADERS);
  for (var i = 0; i < rows.length; i++) if (String(rows[i]['id']) === String(id)) return rows[i];
  return null;
}

function staffSavePhoto_(d) {
  var stf = findStaff_(d.id);
  if (!stf) fail_('ไม่พบเจ้าหน้าที่');
  var m = String(d.dataUrl || '').match(/^data:(image\/(jpeg|png|webp));base64,(.+)$/);
  if (!m) fail_('รองรับเฉพาะรูป JPG, PNG หรือ WEBP');
  if (m[3].length > 400000) fail_('รูปใหญ่เกินไป');
  var file = driveFolder_('staff-photos').createFile(Utilities.newBlob(Utilities.base64Decode(m[3]), m[1], 'staff_' + d.id));
  // ไม่ลบรูปเก่า: สัญญาที่ออกไปแล้วยังอ้างถึงรูปเดิมได้
  var sh = getSheet(SHEET_STAFF), row = findRowIndexById(sh, d.id);
  sh.getRange(row, STAFF_HEADERS.indexOf('รูปถ่าย (file id)') + 1).setValue(file.getId());
  return { success: true, photoFileId: file.getId() };
}

function photoDataUrl_(fileId) {
  if (!fileId) return '';
  try { var b = DriveApp.getFileById(fileId).getBlob(); return 'data:' + b.getContentType() + ';base64,' + Utilities.base64Encode(b.getBytes()); } catch (e) { return ''; }
}

function staffPhotos_() {
  var out = {};
  readSheetAsObjects(SHEET_STAFF, STAFF_HEADERS).forEach(function (r) {
    var id = r['รูปถ่าย (file id)'];
    if (id) out[r['id']] = photoDataUrl_(id);
  });
  return { success: true, photos: out };
}

function splitList_(v) { return String(v || '').split(/\r?\n|,/).map(function (x) { return x.trim(); }).filter(Boolean); }

// เฉพาะข้อมูลที่ลูกค้าควรเห็น — ไม่มีค่าตอบแทน/หมายเหตุภายใน · เบอร์โทรแสดงเมื่อเปิดสิทธิ์ไว้
function staffPublicProfile_(r) {
  var nowBE = new Date().getFullYear() + 543;
  var by = parseInt(r['ปีเกิด (พ.ศ.)'], 10);
  return {
    id: String(r['id']),
    name: r['ชื่อ-นามสกุล'] || '', nickname: r['ชื่อเล่น'] || '', role: r['ตำแหน่ง'] || '',
    age: by > 2400 && by < nowBE ? String(nowBE - by) : '',
    education: r['การศึกษา'] || '', licenses: splitList_(r['ใบอนุญาต/ประกาศนียบัตร']),
    experienceYears: String(r['ประสบการณ์ (ปี)'] || ''), skills: splitList_(r['ทักษะ']), languages: splitList_(r['ภาษา']),
    trainings: splitList_(r['การอบรม']), bio: String(r['แนะนำตัว'] || '').slice(0, 700),
    backgroundCheck: r['ตรวจประวัติเมื่อ'] ? thDate_(r['ตรวจประวัติเมื่อ']) : '',
    since: r['วันที่เริ่มงาน'] ? thDate_(r['วันที่เริ่มงาน']) : '',
    phone: r['แสดงเบอร์โทรในประวัติ'] === 'แสดง' ? (r['เบอร์โทร'] || '') : '',
    photoFileId: r['รูปถ่าย (file id)'] || ''
  };
}

function buildStaffSnapshot_(ids, preview) {
  if (!Array.isArray(ids)) ids = splitList_(ids);
  ids = ids.filter(Boolean).slice(0, 3);
  return ids.map(function (id, i) {
    var r = findStaff_(id);
    if (!r) fail_('ไม่พบเจ้าหน้าที่ ' + id);
    if (!preview && r['ยินยอมเปิดเผยประวัติแก่ลูกค้า'] !== 'ยินยอม') fail_('เจ้าหน้าที่ ' + r['ชื่อ-นามสกุล'] + ' ยังไม่ได้ให้ความยินยอมเปิดเผยประวัติแก่ลูกค้า (แก้ที่หน้า ประวัติเจ้าหน้าที่)');
    if (String(r['สถานะ'] || '') === 'ลาออก') fail_('เจ้าหน้าที่ ' + r['ชื่อ-นามสกุล'] + ' มีสถานะลาออกแล้ว');
    var p = staffPublicProfile_(r);
    p.duty = i === 0 ? 'ผู้ดูแลหลัก' : 'ผู้ดูแลสำรอง';
    if (preview) p.photo = photoDataUrl_(p.photoFileId);
    return p;
  });
}

// เคสใหม่ → ลงชื่อคนไข้ในชีท Patients ให้แอปผู้ดูแลเห็นทันที (ถ้ายังไม่มี)
function ensurePatientFromCase_(d) {
  var name = String(d.clientName || '').trim();
  if (!name) return;
  var sheet = ensurePatientsSheet_();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) if (String(data[i][1]).trim() === name) return;
  var zone = String(d.zoneLabel || '').replace(/^Zone\s*/i, '').slice(0, 1);
  sheet.appendRow(['P' + ('000' + data.length).slice(-3), name, zone, '', '']);
}

/* ============================== สัญญาบริการ + ลงนามอิเล็กทรอนิกส์ ============================== */

function readContracts_() {
  var sh = SS.getSheetByName('Contracts') || (ensureSystemSheets_(), SS.getSheetByName('Contracts'));
  var cols = SYS_SCHEMA_.Contracts, last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, cols.length).getValues().map(function (r, i) {
    var o = { _row: i + 2 };
    cols.forEach(function (k, j) { o[k] = r[j] instanceof Date ? Utilities.formatDate(r[j], TZ, 'yyyy-MM-dd HH:mm:ss') : r[j]; });
    return o;
  }).filter(function (o) { return o.ContractID; });
}

function saveContract_(o) {
  var sh = SS.getSheetByName('Contracts');
  var row = SYS_SCHEMA_.Contracts.map(function (k) { return o[k] === undefined || o[k] === null ? '' : o[k]; });
  if (o._row) sh.getRange(o._row, 1, 1, row.length).setValues([row]);
  else { sh.appendRow(row); o._row = sh.getLastRow(); }
}

function findContract_(col, val) {
  var list = readContracts_();
  for (var i = 0; i < list.length; i++) if (String(list[i][col]) === String(val)) return list[i];
  return null;
}

// เลขที่สัญญาแยกชุด: SC = สัญญาลูกค้า, EC = สัญญาจ้างพนักงาน
function nextContractNo_(prefix) {
  prefix = prefix || 'SC';
  var be = new Date().getFullYear() + 543;
  var key = (prefix === 'SC' ? 'seq_contract_' : 'seq_' + prefix.toLowerCase() + '_') + be;
  var n = (parseInt(readSettingsRaw_()[key], 10) || 0) + 1;
  setSetting_(key, String(n));
  return prefix + be + '-' + ('0000' + n).slice(-4);
}

var CONTRACT_TYPE_LABEL_ = { monthly: 'สัญญารายเดือน', daily: 'สัญญารายวัน', consent: 'ใบยินยอมรับบริการ', addendum: 'ข้อตกลงแนบท้าย', employment: 'สัญญาจ้างพนักงาน' };

function maskId_(id) { id = String(id || ''); return id.length === 13 ? id.charAt(0) + '-xxxx-xxxxx-' + id.slice(10, 12) + '-' + id.charAt(12) : ''; }

function contractsList_(s) {
  var canEmp = !s || canDo_(s, 'staff.pay');
  var list = readContracts_().filter(function (c) { return c.Type !== 'employment' || canEmp; }).map(function (c) {
    return {
      kind: c.Type === 'employment' ? 'employment' : 'customer', staffId: c.StaffID || '', signerIdMasked: maskId_(c.SignerIdNo),
      contractId: c.ContractID, caseId: c.CaseID, type: c.Type, typeLabel: CONTRACT_TYPE_LABEL_[c.Type] || c.Type,
      status: c.Status, clientName: c.ClientName, contactName: c.ContactName,
      signExpires: c.SignExpires, signedAt: c.SignedAt, signerName: c.SignerName, signerRole: c.SignerRole,
      signerRelation: c.SignerRelation, pdfUrl: c.PdfUrl, createdAt: c.CreatedAt, note: c.Note, cancelReason: c.CancelReason,
      firstOpenedAt: c.FirstOpenedAt || '', lastOpenedAt: c.LastOpenedAt || '', openCount: Number(c.OpenCount || 0),
      // รหัสและลิงก์ส่งให้ลูกค้า แสดงเฉพาะสัญญาที่รอเซ็นอยู่
      staff: parseStaffSnap_(c).map(function (x) { return { id: x.id, name: x.name, nickname: x.nickname, duty: x.duty }; }),
      signCode: c.Status === 'awaiting_sign' ? c.SignCode : '',
      signLink: c.Status === 'awaiting_sign' ? signLink_(c.SignToken) : ''
    };
  });
  list.sort(function (a, b) { return String(b.createdAt).localeCompare(String(a.createdAt)); });
  return { success: true, contracts: list };
}

function parseStaffSnap_(c) { try { return JSON.parse(c.StaffSnapshot || '[]'); } catch (e) { return []; } }

function signLink_(token) {
  var base = String(getSettings_().signPageUrl || '').replace(/[?#].*$/, '');
  return base ? base + '?t=' + token : '';
}

function findCase_(caseId) {
  var rows = readSheetAsObjects(SHEET_RECORDS, RECORD_HEADERS);
  for (var i = 0; i < rows.length; i++) if (String(rows[i]['id']) === String(caseId)) return rows[i];
  return null;
}

function num_(v) { var n = Number(String(v === undefined ? '' : v).replace(/,/g, '')); return isNaN(n) ? 0 : n; }

function hoursOf_(fmt) { var m = String(fmt || '').match(/(\d+)\s*ชั่วโมง/); return m ? Number(m[1]) : 0; }

function caseVars_(r, st, staffNames) {
  var v = caseVarsBase_(r, st);
  var hours = hoursOf_(r['รูปแบบบริการ']);
  var ws = String(r['เวลาเริ่มงาน'] || '').slice(0, 5), we = String(r['เวลาเลิกงาน'] || '').slice(0, 5);
  var deposit = num_(r['เงินประกันสัญญา (บาท)']);
  var monthly = r['ประเภทเคส'] === 'แพ็กเกจรายเดือน';
  v.providerParty = st.providerParty; v.legalName = st.legalName;
  v.signerName = st.signerName || '-'; v.signerTitle = st.signerTitle || '';
  v.payerName = v.contactName;
  v.caregiverNames = staffNames || 'ผู้ดูแลตามที่ผู้ให้บริการจัดให้';
  v.hoursPerDay = hours ? String(hours) : '-';
  v.workTime = hours >= 24 ? 'อยู่ประจำบ้าน 24 ชั่วโมง' : ((ws && we) ? 'ตั้งแต่เวลา ' + ws + ' น. ถึง ' + we + ' น.' : 'ตามเวลาที่ตกลงกัน');
  v.employmentMode = monthly ? 'รายเดือน' : 'รายวัน';
  v.formatText = hours >= 24
    ? 'แบบอยู่ประจำบ้าน (24 ชั่วโมง) โดยกำหนดช่วงเวลาพักรับประทานอาหารทั้ง 3 มื้อ และช่วงเวลาพักผ่อนตามความเหมาะสม (อย่างน้อย 6–8 ชั่วโมงต่อวัน)'
    : 'แบบไป–กลับ ระยะเวลาปฏิบัติงาน ' + (hours || '-') + ' ชั่วโมงต่อวัน ' + v.workTime + ' โดยกำหนดช่วงเวลาพักรับประทานอาหาร 2 มื้อ';
  v.monthlyFeeText = bahtText_(num_(r['ค่าบริการต่อเดือน (บาท)']));
  v.rateText = monthly
    ? 'อัตราค่าจ้าง ' + v.monthlyFee + ' บาท (' + v.monthlyFeeText + ') ต่อเดือน'
    : 'อัตราค่าจ้าง ' + v.dailyPkgRate + ' บาทต่อวัน จำนวน ' + v.dailyPkgDays + ' วัน รวม ' + v.dailyPkgTotal + ' บาท';
  v.extraDayRate = monthly ? v.dailyRate : v.dailyPkgRate;
  v.workDaysPerWeek = String(r['วันทำงานต่อสัปดาห์'] || '6');
  v.dayOffPerWeek = String(Math.max(0, 7 - (Number(r['วันทำงานต่อสัปดาห์']) || 6)));
  v.weeklyDayOff = String(r['วันหยุดประจำสัปดาห์'] || 'ตามที่ตกลงกัน');
  v.caregiverMeals = String(r['อาหารผู้ดูแล'] || 'ตามที่ตกลงกัน');
  v.depositClause = deposit > 0
    ? 'เงินประกันการว่าจ้าง ' + v.deposit + ' บาท (' + v.depositText + ') ชำระในวันที่เริ่มจ้างงานและทำสัญญา โดย "ผู้ว่าจ้าง" จะได้รับเงินประกันคืนเต็มจำนวน เมื่อบอกเลิกจ้างล่วงหน้า 1 เดือน และไม่ค้างชำระค่าบริการหรือผิดข้อตกลงในสัญญา'
    : 'สัญญาฉบับนี้ไม่มีเงินประกันการว่าจ้าง';
  return v;
}

function caseVarsBase_(r, st) {
  var dailyRate = num_(r['อัตรารายวัน (บาท)']);
  var deposit = num_(r['เงินประกันสัญญา (บาท)']);
  var procTotal = num_(r['ค่าบริการหัตถการ (บาท)']);
  var courseTotal = num_(r['ราคาคอร์สรวม (บาท)']);
  var serviceItems = [r['รายการหัตถการ'], [r['โปรแกรมคอร์ส'], r['แพ็กเกจคอร์ส']].filter(String).join(' ')].filter(function (x) { return String(x || '').trim(); }).join(' · ') || '-';
  var serviceTotal = procTotal + courseTotal;
  var rs = r['วันที่เริ่ม (แพ็กเกจรายวัน)'], re = r['วันที่สิ้นสุด (แพ็กเกจรายวัน)'];
  return {
    companyName: st.companyName, companyPhone: st.companyPhone || '-',
    clientName: r['ชื่อผู้รับบริการ'] || '-', address: r['ที่อยู่ให้บริการ'] || '-',
    contactName: r['ผู้ติดต่อ'] || '-', contactRelation: r['ความสัมพันธ์'] || '-', contactPhone: r['เบอร์โทรผู้ติดต่อ'] || '-',
    packageName: r['แพ็กเกจ'] || '-', serviceFormat: r['รูปแบบบริการ'] || '-',
    startDate: thDate_(r['วันเริ่มบริการ']), contractEndDate: r['วันครบกำหนดสัญญา'] ? thDate_(r['วันครบกำหนดสัญญา']) : 'ไม่กำหนด',
    monthlyFee: fmt0_(r['ค่าบริการต่อเดือน (บาท)']), dailyRate: fmt0_(dailyRate), admitRate: fmt0_(dailyRate / 2),
    assessmentFee: fmt0_(r['ค่าประเมินหน้างาน (บาท)']), deposit: fmt0_(deposit), depositText: bahtText_(deposit),
    dailyPkgRate: fmt0_(r['อัตรารายวัน (แพ็กเกจรายวัน) (บาท)']), dailyPkgDays: String(r['จำนวนวัน (แพ็กเกจรายวัน)'] || '-'),
    dailyPkgTotal: fmt0_(r['ยอดรวมแพ็กเกจรายวัน (บาท)']),
    dailyPkgRange: (rs && re) ? ('ระหว่างวันที่ ' + thDate_(rs) + ' ถึง ' + thDate_(re)) : ('เริ่มวันที่ ' + thDate_(r['วันเริ่มบริการ'])),
    serviceItems: serviceItems, serviceTotal: fmt0_(serviceTotal), serviceTotalText: bahtText_(serviceTotal)
  };
}

function renderTpl_(s, vars) {
  return String(s).replace(/\{\{(\w+)\}\}/g, function (m, k) { return vars[k] !== undefined ? vars[k] : m; });
}

function textToBlocks_(text, vars) {
  var blocks = [];
  String(text || '').split(/\r?\n/).forEach(function (line) {
    line = line.trim();
    if (!line) return;
    if (line.indexOf('# ') === 0) blocks.push({ h: renderTpl_(line.slice(2), vars) });
    else blocks.push({ p: renderTpl_(line, vars) });
  });
  return blocks;
}

function buildContractSnapshot_(type, r, st, extra) {
  var v = caseVars_(r, st, extra && extra._staffNames);
  var title, terms, summary;
  var company = [st.companyName, st.companyAddress, st.companyPhone ? 'โทร ' + st.companyPhone : ''].filter(String).join(' · ');
  var client = v.clientName + (r['อายุ'] ? ' อายุ ' + r['อายุ'] + ' ปี' : '');
  var payer = v.contactName !== '-' ? v.contactName + ' (' + v.contactRelation + ') โทร ' + v.contactPhone : '-';
  if (type === 'monthly') {
    title = st.contractTitleMonthly; terms = st.termsMonthly;
    summary = [['ผู้ให้บริการ', company], ['ผู้รับบริการ', client], ['ผู้ว่าจ้าง/ผู้ติดต่อ', payer], ['สถานที่ให้บริการ', v.address],
      ['แพ็กเกจ', v.packageName + ' · ' + v.serviceFormat], ['วันเริ่มบริการ', v.startDate], ['วันครบกำหนดสัญญา', v.contractEndDate],
      ['ค่าบริการ', v.monthlyFee + ' บาท/เดือน · อัตรารายวัน ' + v.dailyRate + ' บาท/วัน'], ['เงินประกันสัญญา', v.deposit + ' บาท']];
  } else if (type === 'daily') {
    title = st.contractTitleDaily; terms = st.termsDaily;
    summary = [['ผู้ให้บริการ', company], ['ผู้รับบริการ', client], ['ผู้ว่าจ้าง/ผู้ติดต่อ', payer], ['สถานที่ให้บริการ', v.address],
      ['แพ็กเกจ', v.packageName + ' · ' + v.serviceFormat], ['ช่วงเวลาให้บริการ', v.dailyPkgDays + ' วัน · ' + v.dailyPkgRange],
      ['ค่าบริการ', v.dailyPkgRate + ' บาท/วัน · รวม ' + v.dailyPkgTotal + ' บาท']];
  } else if (type === 'consent') {
    title = st.contractTitleConsent; terms = st.termsConsent;
    summary = [['ผู้ให้บริการ', company], ['ผู้รับบริการ', client], ['ผู้ติดต่อ', payer], ['สถานที่ให้บริการ', v.address],
      ['บริการ', v.serviceItems], ['ค่าบริการ', v.serviceTotal + ' บาท']];
  } else if (type === 'addendum') {
    var ref = String(extra.refContractId || '').trim();
    var text = String(extra.addendumText || '').trim();
    if (!ref) fail_('กรุณาระบุเลขที่สัญญาฉบับเดิม');
    if (text.length < 10) fail_('กรุณาพิมพ์รายละเอียดข้อตกลงที่เปลี่ยนแปลง');
    title = st.contractTitleAddendum;
    terms = '# ข้อตกลงเพิ่มเติม\nบันทึกนี้เป็นส่วนหนึ่งของสัญญาเลขที่ ' + ref + ' ข้อความใดในสัญญาเดิมที่ขัดกับบันทึกนี้ ให้ใช้ข้อความในบันทึกนี้แทน\n' + text +
      '\n# การลงนามอิเล็กทรอนิกส์\nการลงลายมือชื่ออิเล็กทรอนิกส์ในบันทึกนี้มีผลผูกพันตามกฎหมายว่าด้วยธุรกรรมทางอิเล็กทรอนิกส์';
    summary = [['ผู้ให้บริการ', company], ['ผู้รับบริการ', client], ['ผู้ว่าจ้าง/ผู้ติดต่อ', payer], ['อ้างอิงสัญญาเดิม', ref]];
  } else fail_('ประเภทสัญญาไม่ถูกต้อง');

  var pdpaParas = textToBlocks_(st.pdpaText, v).map(function (b) { return b.p || b.h; });
  var snap = { title: title, summary: summary, blocks: textToBlocks_(terms, v), pdpa: pdpaParas,
               company: st.companyName, signerName: st.signerName, signerTitle: st.signerTitle,
               contactName: v.contactName !== '-' ? v.contactName : '', contactRelation: v.contactRelation !== '-' ? v.contactRelation : '', clientName: v.clientName };
  var leftovers = JSON.stringify(snap).match(/\(ใส่ข้อมูล[^)]*\)/g);
  if (leftovers) fail_('ในข้อความสัญญาหรือข้อความ PDPA ยังมีช่องที่ต้องกรอก: ' + leftovers.slice(0, 3).join(', ') + ' — แก้ที่ ตั้งค่า > ข้อความสัญญา');
  var missingVar = JSON.stringify(snap.blocks).match(/\{\{\w+\}\}/g);
  if (missingVar) fail_('ข้อความสัญญามีตัวแปรที่ระบบไม่รู้จัก: ' + missingVar.slice(0, 3).join(', '));
  return snap;
}

function snapshotHash_(no, snap, staff) {
  return sha256hex_(JSON.stringify({ no: no, t: snap.title, s: snap.summary, b: snap.blocks, p: snap.pdpa, staff: staff || [] }));
}

function newSignCredentials_(c) {
  var st = getSettings_();
  c.SignToken = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
  c.SignCode = String(Math.floor(100000 + Math.random() * 900000));
  var days = parseInt(st.signValidDays, 10) || 7;
  c.SignExpires = Utilities.formatDate(new Date(Date.now() + days * 86400000), TZ, 'yyyy-MM-dd HH:mm:ss');
}

function signMessage_(c) {
  var st = getSettings_();
  if (c.Type === 'employment') return 'เรียน คุณ' + c.ClientName + '\n' + st.legalName + ' ส่ง' + st.contractTitleEmployment + ' เลขที่ ' + c.ContractID + ' ให้อ่านและลงนามทางออนไลน์\n' +
    signLink_(c.SignToken) + '\n' + (st.requireSignCode === '1' ? 'รหัสเปิดเอกสาร: ' + c.SignCode + '\n' : '') + 'ลิงก์ใช้ได้ถึง ' + thDateTime_(c.SignExpires) + '\nสอบถามฝ่ายบุคคล โทร ' + (st.companyPhone || '-');
  return 'เรียน คุณ' + (c.ContactName || c.ClientName) + '\n' +
    st.companyName + ' ส่ง' + (CONTRACT_TYPE_LABEL_[c.Type] || 'สัญญา') + ' เลขที่ ' + c.ContractID + ' ของ ' + c.ClientName + ' ให้อ่านและลงนามทางออนไลน์\n' +
    signLink_(c.SignToken) + '\n' +
    (st.requireSignCode === '1' ? 'รหัสเปิดเอกสาร: ' + c.SignCode + '\n' : '') +
    'ลิงก์ใช้ได้ถึง ' + thDateTime_(c.SignExpires) + '\nสอบถามโทร ' + (st.companyPhone || '-');
}

function contractsCreate_(d, s) {
  var st = getSettings_();
  var missing = [];
  if (!st.companyName || st.companyName === 'ชื่อบริษัทของคุณ') missing.push('ชื่อบริษัท');
  if (!String(st.companyAddress).trim()) missing.push('ที่อยู่บริษัท');
  if (!String(st.companyPhone).trim()) missing.push('เบอร์โทรบริษัท');
  if (!String(st.signerName).trim()) missing.push('ชื่อผู้มีอำนาจลงนาม');
  if (!String(st.signPageUrl).trim()) missing.push('ลิงก์หน้าเซ็นสัญญา (sign.html)');
  if (missing.length) fail_('ไปที่ ตั้งค่า และกรอกก่อน: ' + missing.join(', '));

  var r = findCase_(d.caseId);
  if (!r) fail_('ไม่พบเคส ' + d.caseId);
  if (r['สถานะราคา'] === 'รออนุมัติ') fail_('เคสนี้มีราคาต่ำกว่าเกณฑ์และยังรอเจ้าของอนุมัติ — ทำสัญญาได้หลังอนุมัติแล้ว');
  if (r['สถานะราคา'] === 'ไม่อนุมัติ') fail_('เจ้าของไม่อนุมัติราคาของเคสนี้ — แก้ราคาในเคสก่อนทำสัญญา');
  var type = d.type;
  if (!type) {
    var t = String(r['ประเภทเคส'] || '');
    type = t === 'แพ็กเกจรายเดือน' ? 'monthly' : t === 'แพ็กเกจรายวัน' ? 'daily' : 'consent';
  }
  var staffIds = Array.isArray(d.staffIds) ? d.staffIds : splitList_(d.staffIds || r['รหัสเจ้าหน้าที่ผู้ดูแล']);
  var staff = type === 'addendum' && !d.staffIds ? [] : buildStaffSnapshot_(staffIds, false);
  d._staffNames = staff.map(function (x) { return x.name + (x.role ? ' (' + x.role + ')' : ''); }).join(', ');
  var snap = buildContractSnapshot_(type, r, st, d);
  var no = nextContractNo_();
  var c = {
    ContractID: no, CaseID: String(r['id']), Type: type, Status: 'awaiting_sign',
    ClientName: r['ชื่อผู้รับบริการ'] || '', ContactName: r['ผู้ติดต่อ'] || '', ContactPhone: r['เบอร์โทรผู้ติดต่อ'] || '',
    Snapshot: JSON.stringify(snap), ContentHash: snapshotHash_(no, snap, staff),
    StaffIDs: staff.map(function (x) { return x.id; }).join(','), StaffSnapshot: JSON.stringify(staff),
    CreatedAt: nowStr_(), CreatedBy: s.name, Note: String(d.note || '').slice(0, 300)
  };
  if (c.Snapshot.length > 48000) fail_('ข้อความสัญญายาวเกินไป กรุณาย่อข้อความในตั้งค่า');
  newSignCredentials_(c);
  saveContract_(c);
  return { success: true, contractId: no, link: signLink_(c.SignToken), code: c.SignCode, expires: c.SignExpires, message: signMessage_(c), requireCode: st.requireSignCode === '1' };
}

function contractsResend_(d, s) {
  var c = findContract_('ContractID', d.contractId);
  if (!c) fail_('ไม่พบสัญญา');
  if (c.Type === 'employment' && s && !canDo_(s, 'staff.pay')) fail_('บัญชีของคุณไม่มีสิทธิ์จัดการสัญญาจ้างพนักงาน', 'FORBIDDEN');
  if (c.Status !== 'awaiting_sign') fail_('ออกลิงก์ใหม่ได้เฉพาะสัญญาที่รอลงนาม');
  newSignCredentials_(c);
  saveContract_(c);
  CacheService.getScriptCache().remove('sf_' + String(c.SignToken).slice(0, 16));
  return { success: true, contractId: c.ContractID, link: signLink_(c.SignToken), code: c.SignCode, expires: c.SignExpires, message: signMessage_(c) };
}

function contractsCancel_(d, s) {
  var c = findContract_('ContractID', d.contractId);
  if (!c) fail_('ไม่พบสัญญา');
  if (c.Type === 'employment' && s && !canDo_(s, 'staff.pay')) fail_('บัญชีของคุณไม่มีสิทธิ์จัดการสัญญาจ้างพนักงาน', 'FORBIDDEN');
  if (c.Status === 'cancelled') fail_('สัญญานี้ยกเลิกไปแล้ว');
  var reason = String(d.reason || '').trim();
  if (!reason) fail_('กรุณาระบุเหตุผลที่ยกเลิก');
  c.Status = 'cancelled'; c.CancelReason = reason; c.SignToken = ''; c.SignCode = '';
  saveContract_(c);
  return { success: true };
}

/* ---------- หน้าเซ็น (สาธารณะ: ใช้ token ในลิงก์ + รหัส 6 หลัก) ---------- */

function checkSignAccess_(d, needCode) {
  var token = String(d.token || '');
  if (token.length < 40) fail_('ลิงก์ไม่ถูกต้อง');
  var c = findContract_('SignToken', token);
  if (!c) fail_('ลิงก์ไม่ถูกต้องหรือถูกยกเลิกแล้ว กรุณาติดต่อผู้ให้บริการ');
  if (c.Status === 'signed') fail_('เอกสารนี้ลงนามเรียบร้อยแล้ว', 'SIGNED');
  if (c.Status !== 'awaiting_sign') fail_('ลิงก์นี้ใช้ไม่ได้แล้ว กรุณาติดต่อผู้ให้บริการเพื่อขอลิงก์ใหม่');
  if (c.SignExpires && String(c.SignExpires) < nowStr_()) fail_('ลิงก์หมดอายุแล้ว กรุณาติดต่อผู้ให้บริการเพื่อขอลิงก์ใหม่', 'EXPIRED');
  if (needCode && getSettings_().requireSignCode === '1') {
    var cache = CacheService.getScriptCache(), key = 'sf_' + token.slice(0, 16);
    var tries = Number(cache.get(key) || 0);
    if (tries >= 5) fail_('ใส่รหัสผิดหลายครั้ง กรุณาติดต่อผู้ให้บริการเพื่อขอรหัสใหม่', 'LOCKED');
    if (String(d.code || '').trim() !== String(c.SignCode)) {
      cache.put(key, String(tries + 1), 3600);
      fail_('รหัสเปิดเอกสารไม่ถูกต้อง (เหลือ ' + (4 - tries) + ' ครั้ง)', 'BADCODE');
    }
  }
  return c;
}

function signGet_(d) {
  var st = getSettings_();
  var brand = publicBranding_();
  var c = checkSignAccess_(d, false);
  if (st.requireSignCode === '1' && !String(d.code || '').trim()) return { success: true, needCode: true, brand: brand };
  c = checkSignAccess_(d, true);
  // บันทึกว่าลูกค้าเปิดอ่านแล้ว (ให้แอดมินรู้ว่าควรโทรตามหรือส่งลิงก์ใหม่)
  try {
    var t = nowStr_();
    if (!c.FirstOpenedAt) c.FirstOpenedAt = t;
    c.LastOpenedAt = t; c.OpenCount = Number(c.OpenCount || 0) + 1;
    saveContract_(c);
  } catch (e) { /* บันทึกไม่ได้ไม่เป็นไร ให้ลูกค้าอ่านต่อ */ }
  var snap = JSON.parse(c.Snapshot);
  return {
    success: true, needCode: false, brand: brand, no: c.ContractID, typeLabel: CONTRACT_TYPE_LABEL_[c.Type] || '',
    title: snap.title, summary: snap.summary, blocks: snap.blocks, pdpa: snap.pdpa, hash: c.ContentHash, expires: c.SignExpires,
    staff: parseStaffSnap_(c).map(function (x) { var o = JSON.parse(JSON.stringify(x)); o.photo = photoDataUrl_(x.photoFileId); delete o.photoFileId; return o; }),
    watermark: getWatermarkDataUrl_(),
    kind: c.Type === 'employment' ? 'employment' : 'customer',
    askId: c.Type === 'employment' || st.requireSignerId === '1',
    profileConsentText: c.Type === 'employment' ? snap.profileConsentText || '' : '',
    privacyTitle: c.Type === 'employment' ? 'ประกาศความเป็นส่วนตัวสำหรับลูกจ้าง' : 'ความยินยอมเรื่องข้อมูลส่วนบุคคลและข้อมูลสุขภาพ',
    signerHint: { name: snap.contactName || '', relation: snap.contactRelation || '', clientName: snap.clientName, phone: snap.contactPhone || '' }
  };
}

// ตรวจเลขประจำตัวประชาชนไทย 13 หลัก (ตรวจหลักสุดท้ายตามสูตรกรมการปกครอง)
function thaiIdOk_(id) {
  id = String(id || '').replace(/\D/g, '');
  if (id.length !== 13) return false;
  var sum = 0; for (var i = 0; i < 12; i++) sum += Number(id.charAt(i)) * (13 - i);
  return (11 - sum % 11) % 10 === Number(id.charAt(12));
}

function signSubmit_(d) {
  var c = checkSignAccess_(d, true);
  if (c.Type === 'employment') return signSubmitEmployment_(c, d);
  var name = String(d.name || '').trim().slice(0, 120);
  if (name.length < 2) fail_('กรุณาพิมพ์ชื่อ-นามสกุลผู้ลงนาม');
  var idNo = String(d.idNo || '').replace(/\D/g, '');
  if (getSettings_().requireSignerId === '1' && !thaiIdOk_(idNo)) fail_('เลขประจำตัวประชาชนไม่ถูกต้อง กรุณาตรวจสอบอีกครั้ง');
  var role = d.role === 'self' ? 'self' : 'representative';
  var relation = String(d.relation || '').trim().slice(0, 60);
  if (role === 'representative' && !relation) fail_('กรุณาระบุความสัมพันธ์กับผู้รับบริการ');
  if (d.agree !== true) fail_('กรุณาติ๊กยืนยันว่าได้อ่านและยอมรับเอกสาร');
  if (d.pdpa !== true) fail_('กรุณาติ๊กให้ความยินยอมเรื่องข้อมูลส่วนบุคคลและข้อมูลสุขภาพ');
  var sig = String(d.signature || '');
  if (sig.indexOf('data:image/png;base64,') !== 0 || sig.length < 1500 || sig.length > 900000) fail_('กรุณาเซ็นชื่อในช่องลายเซ็น');
  var snap = JSON.parse(c.Snapshot);
  if (snapshotHash_(c.ContractID, snap, parseStaffSnap_(c)) !== c.ContentHash || d.hash !== c.ContentHash) fail_('เนื้อหาเอกสารมีการเปลี่ยนแปลงระหว่างที่คุณเปิดอ่าน กรุณารีเฟรชหน้านี้แล้วอ่านใหม่อีกครั้ง');

  var signedAt = nowStr_();
  var ua = String(d.ua || '').slice(0, 200);
  var sigFile = driveFolder_('signatures').createFile(Utilities.newBlob(Utilities.base64Decode(sig.split(',')[1]), 'image/png', c.ContractID + '_signature.png'));
  var html = contractHtml_(c, snap, { name: name, role: role, relation: relation, at: signedAt, sig: sig, ua: ua, idNo: idNo });
  var pdf = savePdf_(html, c.ContractID + '_' + c.ClientName, 'contracts');

  c.Status = 'signed'; c.SignedAt = signedAt; c.SignerName = name; c.SignerRole = role; c.SignerRelation = relation;
  c.PdpaConsent = 'Y'; c.SignedUA = ua; c.PdfUrl = pdf.url; c.PdfId = pdf.id; c.SignatureFileId = sigFile.getId(); c.SignCode = '';
  c.SignerIdNo = idNo;
  saveContract_(c);
  log_('sign.submit', name, c.ContractID, relation);

  var st = getSettings_();
  String(st.notifyEmail || '').split(',').map(function (x) { return x.trim(); }).filter(Boolean).forEach(function (email) {
    try { MailApp.sendEmail(email, 'ลงนามแล้ว ' + c.ContractID + ' — ' + c.ClientName, name + ' ลงนาม' + (CONTRACT_TYPE_LABEL_[c.Type] || 'สัญญา') + ' ' + c.ContractID + ' เมื่อ ' + thDateTime_(signedAt) + '\n' + pdf.url); } catch (e) { /* ข้าม */ }
  });
  return { success: true, signed: true, no: c.ContractID, pdfUrl: st.sharePublic === '1' ? pdf.url : '' };
}


/* ============================== สัญญาจ้างพนักงาน (บริษัท ↔ ลูกจ้าง) ============================== */

var EMP_MODES_ = ['ไป–กลับ', 'พักประจำบ้านผู้รับบริการ', 'ประจำโรงพยาบาล', 'หมุนเวียนตามเคส'];
var EMP_HOLIDAYS_ = ['1 มกราคม', '13 เมษายน', '14 เมษายน', '15 เมษายน', '31 ธันวาคม'];
var EMP_CASES_ = [['general', 'Case ทั่วไป'], ['feed', 'Case Feed'], ['suction', 'Case Feed suction'], ['special', 'Case หัตถการพิเศษ']];

function wageLine_(w) {
  w = w || {};
  var parts = EMP_CASES_.filter(function (k) { return num_(w[k[0]]) > 0; }).map(function (k) { return k[1] + ' ' + fmt0_(w[k[0]]) + ' บาท/วัน'; });
  return parts.join(' · ');
}

function contractsCreateEmployment_(d, s) {
  var st = getSettings_();
  var missing = [];
  if (!String(st.companyAddress).trim()) missing.push('ที่อยู่บริษัท');
  if (!String(st.signerName).trim()) missing.push('ชื่อผู้มีอำนาจลงนาม');
  if (!String(st.signPageUrl).trim()) missing.push('ลิงก์หน้าเซ็นสัญญา (sign.html)');
  if (missing.length) fail_('ไปที่ ตั้งค่า และกรอกก่อน: ' + missing.join(', '));
  var r = findStaff_(d.staffId);
  if (!r) fail_('ไม่พบเจ้าหน้าที่');
  if (!d.startDate) fail_('กรุณาระบุวันเริ่มงาน');
  var modes = (d.modes || []).filter(function (m) { return EMP_MODES_.indexOf(m) >= 0; });
  if (!modes.length) fail_('กรุณาเลือกรูปแบบงานอย่างน้อย 1 แบบ');
  var wc = wageLine_(d.wageCommute), wl = wageLine_(d.wageLiveIn);
  if (!wc && !wl) fail_('กรุณากรอกค่าจ้างอย่างน้อย 1 รายการ');
  var payDay = parseInt(d.payDay, 10);
  if (!(payDay >= 1 && payDay <= 31)) fail_('กรุณาระบุวันที่จ่ายค่าจ้าง (1–31)');

  var nowBE = new Date().getFullYear() + 543, by = parseInt(r['ปีเกิด (พ.ศ.)'], 10);
  var age = by > 2400 && by < nowBE ? (nowBE - by) + ' ปี' : '-';
  var probation = parseInt(d.probationDays, 10) || 0;
  var probText = 'ไม่มี';
  if (probation > 0) {
    var p0 = new Date(d.startDate + 'T00:00:00'); var p1 = new Date(p0.getTime() + (probation - 1) * 86400000);
    probText = probation + ' วัน (ตั้งแต่ ' + thDate_(d.startDate) + ' ถึง ' + thDate_(Utilities.formatDate(p1, TZ, 'yyyy-MM-dd')) + ') โดยสิทธิแรงงานยังเป็นไปตามกฎหมาย';
  }
  var typeText = d.employmentType === 'fixed'
    ? 'มีกำหนดระยะเวลา ตั้งแต่ ' + thDate_(d.fixedFrom) + ' ถึง ' + thDate_(d.fixedTo)
    : 'ไม่มีกำหนดระยะเวลา';
  if (d.employmentType === 'fixed' && (!d.fixedFrom || !d.fixedTo)) fail_('กรุณาระบุช่วงวันที่ของการจ้างแบบมีกำหนดระยะเวลา');
  var modeText = modes.map(function (m) { return m === 'ไป–กลับ' && d.commuteFrom && d.commuteTo ? m + ' ตั้งแต่ ' + d.commuteFrom + ' น. ถึง ' + d.commuteTo + ' น.' : m; }).join(' · ');
  var holidays = (d.holidays || []).filter(function (x) { return EMP_HOLIDAYS_.indexOf(x) >= 0; });
  var position = String(d.position || st.employmentPosition || 'พนักงานเฝ้าไข้').slice(0, 80);
  var place = String(d.place || st.companyAddress).replace(/\s*\n\s*/g, ' ').slice(0, 200);

  var vars = { legalName: st.legalName, companyName: st.companyName, position: position, companyPhone: st.companyPhone || '-' };
  var summary = [
    ['วันที่ทำสัญญา', thDate_(todayStr_())], ['สถานที่ทำสัญญา', place],
    ['นายจ้าง', st.legalName + (st.taxId ? ' · เลขทะเบียนนิติบุคคล ' + st.taxId : '')],
    ['สำนักงาน', String(st.companyAddress).replace(/\s*\n\s*/g, ' ')],
    ['ผู้แทนบริษัท', 'โดย ' + st.signerName + (st.signerTitle ? ' ตำแหน่ง ' + st.signerTitle : '')],
    ['ลูกจ้าง', r['ชื่อ-นามสกุล'] + (r['ชื่อเล่น'] ? ' (ชื่อเล่น ' + r['ชื่อเล่น'] + ')' : '') + ' · อายุ ' + age],
    ['ตำแหน่ง', position], ['ประเภทการจ้าง', typeText], ['วันเริ่มงาน', thDate_(d.startDate)], ['ทดลองงาน', probText],
    ['รูปแบบงาน', modeText],
    ['เวลาปกติ', (d.hoursPerDay || '-') + ' ชั่วโมง/วัน และ ' + (d.daysPerWeek || '-') + ' วัน/สัปดาห์'],
    ['วันหยุดประจำสัปดาห์', String(d.weeklyDayOff || '-').slice(0, 80)],
    ['วันหยุดพิเศษประจำปี (เฉพาะเฝ้ารายเดือน)', holidays.length ? holidays.join(', ') : '-']
  ];
  if (wc) summary.push(['ค่าจ้าง ไป–กลับ หมุนเวียนตามเคส', wc]);
  if (wl) summary.push(['ค่าจ้าง พักประจำบ้านผู้รับบริการ (รายเดือน)', wl]);
  summary.push(['กำหนดจ่าย', 'ทุกวันที่ ' + payDay + ' ของเดือน โดย' + (d.payMethod === 'other' ? (' ' + String(d.payMethodOther || 'วิธีอื่น').slice(0, 60)) : 'โอนเข้าบัญชีลูกจ้าง')]);

  var blocks = textToBlocks_(st.termsEmployment, vars);
  var annex = [{ h: 'เอกสารแนบท้ายสัญญา: รายละเอียดตำแหน่ง' + position }];
  annex.push({ p: 'รหัสพนักงาน: ' + (d.employeeCode || r['id']) + ' · หน่วยงาน/ทีม: ' + (d.team || '-') });
  annex.push({ p: 'คุณวุฒิ: ' + ((d.quals || []).join(', ') || '-') });
  annex.push({ p: 'ทักษะที่รับรอง: ' + ((d.skills || []).concat(d.skillOther ? [d.skillOther] : []).join(', ') || '-') });
  annex.push({ p: 'ข้อจำกัดการมอบหมาย: ' + (String(d.limits || '').trim() || 'ไม่มี') });
  annex.push({ p: 'เอกสารที่ลูกจ้างได้รับ: ' + ((d.docs || []).join(' · ') || '-') });
  annex.push({ p: 'ลูกจ้างรับรองว่าได้รับ อ่าน และเข้าใจเอกสารข้างต้นแล้ว' });
  var snap = {
    title: st.contractTitleEmployment, summary: summary, blocks: blocks.concat(annex),
    pdpa: textToBlocks_(st.pdpaEmployment, vars).map(function (b) { return b.p || b.h; }),
    profileConsentText: renderTpl_(st.profileConsentText, vars),
    company: st.legalName, signerName: st.signerName, signerTitle: st.signerTitle,
    contactName: r['ชื่อ-นามสกุล'], contactPhone: r['เบอร์โทร'] || '', clientName: r['ชื่อ-นามสกุล']
  };
  var miss = JSON.stringify(snap).match(/\{\{\w+\}\}/g);
  if (miss) fail_('ข้อความสัญญาจ้างมีตัวแปรที่ระบบไม่รู้จัก: ' + miss.slice(0, 3).join(', '));
  var no = nextContractNo_('EC');
  var c = {
    ContractID: no, CaseID: '', StaffID: String(r['id']), Type: 'employment', Status: 'awaiting_sign',
    ClientName: r['ชื่อ-นามสกุล'], ContactName: r['ชื่อ-นามสกุล'], ContactPhone: r['เบอร์โทร'] || '',
    Snapshot: JSON.stringify(snap), ContentHash: snapshotHash_(no, snap, []), StaffSnapshot: '[]',
    CreatedAt: nowStr_(), CreatedBy: s.name, Note: String(d.note || '').slice(0, 300)
  };
  newSignCredentials_(c);
  saveContract_(c);
  return { success: true, contractId: no, link: signLink_(c.SignToken), code: c.SignCode, expires: c.SignExpires, message: signMessage_(c) };
}

function signSubmitEmployment_(c, d) {
  var idNo = String(d.idNo || '').replace(/\D/g, '');
  if (!thaiIdOk_(idNo)) fail_('เลขประจำตัวประชาชนไม่ถูกต้อง กรุณาตรวจสอบอีกครั้ง');
  var name = String(d.name || '').trim().slice(0, 120);
  if (name.length < 2) fail_('กรุณาพิมพ์ชื่อ-นามสกุล');
  var ex = d.extra || {};
  var info = {
    address: String(ex.address || '').trim().slice(0, 300), phone: String(ex.phone || '').trim().slice(0, 40), line: String(ex.line || '').trim().slice(0, 60),
    nationality: String(ex.nationality || '').trim().slice(0, 40),
    bankName: String(ex.bankName || '').trim().slice(0, 80), bankBranch: String(ex.bankBranch || '').trim().slice(0, 80),
    accountName: String(ex.accountName || '').trim().slice(0, 120), accountNo: String(ex.accountNo || '').replace(/[^\d-]/g, '').slice(0, 30)
  };
  if (info.address.length < 8) fail_('กรุณากรอกที่อยู่ให้ครบ');
  if (info.phone.replace(/\D/g, '').length < 9) fail_('กรุณากรอกเบอร์โทรศัพท์');
  if (d.agree !== true) fail_('กรุณาติ๊กยืนยันว่าได้อ่านและยอมรับสัญญาจ้าง');
  if (d.pdpa !== true) fail_('กรุณาติ๊กรับทราบประกาศความเป็นส่วนตัว');
  var sig = String(d.signature || '');
  if (sig.indexOf('data:image/png;base64,') !== 0 || sig.length < 1500 || sig.length > 900000) fail_('กรุณาเซ็นชื่อในช่องลายเซ็น');
  var snap = JSON.parse(c.Snapshot);
  if (snapshotHash_(c.ContractID, snap, []) !== c.ContentHash || d.hash !== c.ContentHash) fail_('เนื้อหาเอกสารมีการเปลี่ยนแปลงระหว่างที่คุณเปิดอ่าน กรุณารีเฟรชหน้านี้แล้วอ่านใหม่อีกครั้ง');
  var profileConsent = d.profileConsent === true;

  var signedAt = nowStr_(), ua = String(d.ua || '').slice(0, 200);
  var sigFile = driveFolder_('signatures').createFile(Utilities.newBlob(Utilities.base64Decode(sig.split(',')[1]), 'image/png', c.ContractID + '_signature.png'));
  var html = employmentHtml_(c, snap, { name: name, idNo: idNo, info: info, at: signedAt, sig: sig, ua: ua, profileConsent: profileConsent });
  var pdf = savePdf_(html, c.ContractID + '_' + c.ClientName, 'employment-contracts');
  c.Status = 'signed'; c.SignedAt = signedAt; c.SignerName = name; c.SignerRole = 'employee'; c.SignerRelation = '';
  c.PdpaConsent = profileConsent ? 'Y' : 'N'; c.SignedUA = ua; c.PdfUrl = pdf.url; c.PdfId = pdf.id; c.SignatureFileId = sigFile.getId(); c.SignCode = '';
  c.SignerIdNo = idNo; c.SignerExtra = JSON.stringify(info);
  saveContract_(c);
  // อัปเดตประวัติเจ้าหน้าที่: ความยินยอมเปิดเผยประวัติ + เบอร์โทร (ถ้ายังไม่มี)
  try {
    var sh = getSheet(SHEET_STAFF), row = findRowIndexById(sh, c.StaffID);
    if (row > 0) {
      sh.getRange(row, STAFF_HEADERS.indexOf('ยินยอมเปิดเผยประวัติแก่ลูกค้า') + 1).setValue(profileConsent ? 'ยินยอม' : 'ไม่ยินยอม');
      var pc = STAFF_HEADERS.indexOf('เบอร์โทร') + 1;
      if (!String(sh.getRange(row, pc).getValue()).trim()) sh.getRange(row, pc).setValue(info.phone);
    }
  } catch (e) { /* ข้าม */ }
  log_('sign.employment', name, c.ContractID, profileConsent ? 'profile-consent' : '');
  var st = getSettings_();
  String(st.notifyEmail || '').split(',').map(function (x) { return x.trim(); }).filter(Boolean).forEach(function (email) {
    try { MailApp.sendEmail(email, 'พนักงานลงนามสัญญาจ้างแล้ว ' + c.ContractID + ' — ' + c.ClientName, name + ' ลงนามสัญญาจ้าง ' + c.ContractID + ' เมื่อ ' + thDateTime_(signedAt) + '\n' + pdf.url); } catch (e) { /* ข้าม */ }
  });
  return { success: true, signed: true, no: c.ContractID, pdfUrl: st.sharePublic === '1' ? pdf.url : '' };
}

function employmentHtml_(c, snap, sign) {
  var st = getSettings_(), i = sign.info;
  var h = '<table><tr><td style="width:64px"><img src="' + getLogoDataUrl_() + '" style="width:56px;height:56px"></td>' +
    '<td><div class="brand">' + esc_(snap.company) + '</div><div class="sm">' + (st.legalName && st.legalName !== snap.company ? esc_(st.legalName) + '<br>' : '') + br_(st.companyAddress) + (st.companyPhone ? ' · โทร ' + esc_(st.companyPhone) : '') + '</div></td>' +
    '<td style="text-align:right;vertical-align:top" class="sm">เลขที่สัญญา <b>' + esc_(c.ContractID) + '</b><br>ออกเมื่อ ' + esc_(thDate_(c.CreatedAt)) + '<br>เอกสารภายในบริษัท</td></tr></table>';
  h += '<div class="c" style="font-size:19px;font-weight:bold;margin:12px 0 2px">' + esc_(snap.title) + '</div><div class="c sm" style="margin-bottom:8px">ระหว่าง ' + esc_(snap.company) + ' กับพนักงานรายบุคคล</div>';
  h += '<table class="sum">';
  snap.summary.forEach(function (r) { h += '<tr><td class="k">' + esc_(r[0]) + '</td><td>' + esc_(r[1]) + '</td></tr>'; });
  h += '<tr><td class="k">ข้อมูลลูกจ้าง</td><td>เลขบัตรประชาชน ' + esc_(sign.idNo) + (i.nationality ? ' · สัญชาติ ' + esc_(i.nationality) : '') + '<br>ที่อยู่ ' + esc_(i.address) + '<br>โทร ' + esc_(i.phone) + (i.line ? ' · LINE ' + esc_(i.line) : '') + '</td></tr>';
  if (i.accountNo) h += '<tr><td class="k">บัญชีรับเงิน</td><td>' + esc_([i.bankName, i.bankBranch ? 'สาขา ' + i.bankBranch : '', i.accountName ? 'ชื่อบัญชี ' + i.accountName : '', 'เลขที่ ' + i.accountNo].filter(String).join(' · ')) + '</td></tr>';
  h += '</table>';
  snap.blocks.forEach(function (b) { h += b.h !== undefined ? '<h3>' + esc_(b.h) + '</h3>' : '<p>' + esc_(b.p) + '</p>'; });
  h += '<h3>ประกาศความเป็นส่วนตัวสำหรับลูกจ้าง</h3>';
  snap.pdpa.forEach(function (p) { h += '<p>' + esc_(p) + '</p>'; });
  h += '<p><b>☑ ลูกจ้างรับทราบประกาศความเป็นส่วนตัวแล้ว</b></p>';
  h += '<p><b>' + (sign.profileConsent ? '☑ ลูกจ้างยินยอม' : '☐ ลูกจ้างไม่ยินยอม') + '</b>ให้เปิดเผยประวัติโดยย่อแก่ผู้รับบริการ: ' + esc_(snap.profileConsentText) + '</p>';
  h += '<table style="margin-top:26px"><tr>' +
    '<td class="c" style="width:50%;vertical-align:bottom"><div style="font-size:15px;height:58px;line-height:58px">' + esc_(snap.signerName) + '</div>_________________________<br>' + esc_(snap.signerName) + '<br>' + esc_(snap.signerTitle) + ' · นายจ้าง<br><span class="sm">ลงนามโดยการออกเอกสารผ่านระบบ</span></td>' +
    '<td class="c" style="vertical-align:bottom"><img src="' + sign.sig + '" style="height:58px"><br>_________________________<br>' + esc_(sign.name) + '<br>ลูกจ้าง<br><span class="sm">ลงนามอิเล็กทรอนิกส์</span></td></tr></table>';
  h += '<div class="box sm"><b>บันทึกการลงนามอิเล็กทรอนิกส์</b><br>เลขที่สัญญา ' + esc_(c.ContractID) + ' · ลงนามเมื่อ ' + esc_(thDateTime_(sign.at)) +
    '<br>ผู้ลงนาม: ' + esc_(sign.name) + ' (ลูกจ้าง) · ยืนยันตัวตนด้วยลิงก์เฉพาะบุคคล' + (st.requireSignCode === '1' ? ' และรหัส 6 หลัก' : '') + ' และเลขประจำตัวประชาชน' +
    '<br>รหัสตรวจสอบเนื้อหา (SHA-256): ' + esc_(c.ContentHash) + '<br>อุปกรณ์ที่ใช้ลงนาม: ' + esc_(sign.ua || '-') + '</div>';
  return pdfShell_(h);
}

/* ============================== PDF ============================== */

function savePdf_(html, fileName, folderKey) {
  var blob = Utilities.newBlob(html, 'text/html', fileName + '.html').getAs('application/pdf').setName(fileName + '.pdf');
  var file = driveFolder_(folderKey).createFile(blob);
  if (getSettings_().sharePublic === '1') file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { url: file.getUrl(), id: file.getId() };
}

function pdfShell_(body) {
  var color = getSettings_().brandColor || '#2F6F62';
  var css = '@page{size:A4;margin:14mm}' +
    'body{font-family:"Sarabun","Noto Sans Thai",Tahoma,sans-serif;font-size:13px;color:#1F2A2A;line-height:1.55}' +
    'table{border-collapse:collapse;width:100%}' +
    '.brand{font-size:18px;font-weight:bold;color:' + color + '}' +
    '.sm{font-size:11px;color:#4B5857}' +
    '.sum td{border:1px solid #CFDAD7;padding:5px 7px;vertical-align:top}' +
    '.sum td.k{width:27%;background:#F1F6F4;font-weight:bold}' +
    'h3{font-size:14px;margin:12px 0 2px;color:' + color + '}p{margin:3px 0}' +
    '.box{border:1px solid #CFDAD7;padding:8px 10px;margin-top:12px}' +
    '.c{text-align:center}';
  return '<html><head><meta charset="utf-8"><style>' + css + '</style></head><body>' + body + '</body></html>';
}

function contractHtml_(c, snap, sign) {
  var st = getSettings_();
  var h = '<table><tr><td style="width:64px"><img src="' + getLogoDataUrl_() + '" style="width:56px;height:56px"></td>' +
    '<td><div class="brand">' + esc_(snap.company) + '</div><div class="sm">' + (st.legalName && st.legalName !== snap.company ? esc_(st.legalName) + '<br>' : '') + br_(st.companyAddress) + (st.companyPhone ? ' · โทร ' + esc_(st.companyPhone) : '') + (st.taxId ? ' · เลขประจำตัวผู้เสียภาษี ' + esc_(st.taxId) : '') + '</div></td>' +
    '<td style="text-align:right;vertical-align:top" class="sm">เลขที่ <b>' + esc_(c.ContractID) + '</b><br>ออกเมื่อ ' + esc_(thDate_(c.CreatedAt)) + '</td></tr></table>';
  h += '<div class="c" style="font-size:19px;font-weight:bold;margin:12px 0 8px">' + esc_(snap.title) + '</div>';
  h += '<table class="sum">';
  snap.summary.forEach(function (r) { h += '<tr><td class="k">' + esc_(r[0]) + '</td><td>' + esc_(r[1]) + '</td></tr>'; });
  h += '</table>';
  snap.blocks.forEach(function (b) { h += b.h !== undefined ? '<h3>' + esc_(b.h) + '</h3>' : '<p>' + esc_(b.p) + '</p>'; });
  h += '<h3>ความยินยอมเกี่ยวกับข้อมูลส่วนบุคคลและข้อมูลสุขภาพ</h3>';
  snap.pdpa.forEach(function (p) { h += '<p>' + esc_(p) + '</p>'; });
  h += '<p><b>☑ ผู้ลงนามให้ความยินยอมตามข้อความข้างต้น</b> (ติ๊กแยกจากการยอมรับเอกสาร)</p>';
  var roleText = sign.role === 'self' ? 'ผู้รับบริการ' : 'ผู้แทนผู้รับบริการ (' + esc_(sign.relation) + ')';
  h += '<table style="margin-top:26px"><tr>' +
    '<td class="c" style="width:50%;vertical-align:bottom"><div style="font-size:15px;height:58px;line-height:58px">' + esc_(snap.signerName) + '</div>' +
    '_________________________<br>' + esc_(snap.signerName) + '<br>' + esc_(snap.signerTitle) + ' · ผู้ให้บริการ<br><span class="sm">ลงนามโดยการออกเอกสารผ่านระบบ</span></td>' +
    '<td class="c" style="vertical-align:bottom"><img src="' + sign.sig + '" style="height:58px"><br>_________________________<br>' + esc_(sign.name) + '<br>' + roleText + (sign.idNo ? '<br><span class="sm">เลขประจำตัวประชาชน ' + esc_(sign.idNo) + '</span>' : '') + '<br><span class="sm">ลงนามอิเล็กทรอนิกส์</span></td>' +
    '</tr></table>';
  h += '<div class="box sm"><b>บันทึกการลงนามอิเล็กทรอนิกส์</b><br>เลขที่เอกสาร ' + esc_(c.ContractID) + ' · ลงนามเมื่อ ' + esc_(thDateTime_(sign.at)) +
    '<br>ผู้ลงนาม: ' + esc_(sign.name) + ' ในฐานะ' + roleText + '<br>ยืนยันตัวตนด้วยลิงก์เฉพาะบุคคล' + (st.requireSignCode === '1' ? ' และรหัส 6 หลัก' : '') +
    '<br>รหัสตรวจสอบเนื้อหา (SHA-256): ' + esc_(c.ContentHash) + '<br>อุปกรณ์ที่ใช้ลงนาม: ' + esc_(sign.ua || '-') + '</div>';
  if (st.docFooter) h += '<p class="sm c" style="margin-top:12px">' + esc_(st.docFooter) + '</p>';
  parseStaffSnap_(c).forEach(function (p) { h += staffProfileHtml_(p, c.ContractID, snap.company); });
  return pdfShell_(h);
}

// หน้าประวัติเจ้าหน้าที่ 1 หน้า/คน พื้นหลังลายน้ำวงกลมของแบรนด์ (ภาพจางสำเร็จรูป ไม่พึ่ง CSS opacity)
function staffProfileHtml_(p, contractNo, company) {
  var color = getSettings_().brandColor || '#2F6F62';
  var wm = getWatermarkDataUrl_();
  var photo = photoDataUrl_(p.photoFileId);
  function row(k, v) { return v ? '<tr><td style="width:30%;color:#5A6866;padding:5px 0;vertical-align:top">' + esc_(k) + '</td><td style="padding:5px 0">' + v + '</td></tr>' : ''; }
  function list(a) { return a && a.length ? a.map(esc_).join('<br>') : ''; }
  var h = '<div style="page-break-before:always"></div>';
  h += '<table width="100%" background="' + wm + '" style="background-image:url(' + wm + ');background-repeat:no-repeat;background-position:center 58%;background-size:78%;height:255mm"><tr><td style="vertical-align:top;padding:4mm 2mm">';
  h += '<table><tr><td class="sm">เอกสารแนบท้าย ' + esc_(contractNo) + '</td><td class="sm" style="text-align:right">' + esc_(company) + '</td></tr></table>';
  h += '<div class="c" style="margin-top:14px">' + (photo ? '<img src="' + photo + '" style="width:120px;height:120px;border-radius:60px;object-fit:cover;border:4px solid ' + color + '">' : '') + '</div>';
  h += '<div class="c" style="font-size:22px;font-weight:bold;margin-top:8px">' + esc_(p.name) + (p.nickname ? ' (' + esc_(p.nickname) + ')' : '') + '</div>';
  h += '<div class="c" style="color:' + color + ';font-weight:bold">' + esc_(p.role || 'ผู้ดูแล') + ' · ' + esc_(p.duty) + '</div>';
  if (p.bio) h += '<p class="c" style="margin:12px 30px 0;font-style:italic">“' + esc_(p.bio) + '”</p>';
  h += '<h3 style="margin-top:18px">ประวัติโดยย่อ</h3><table>' +
    row('อายุ', p.age ? esc_(p.age) + ' ปี' : '') + row('ประสบการณ์', p.experienceYears ? esc_(p.experienceYears) + ' ปี' : '') +
    row('การศึกษา', esc_(p.education)) + row('ใบอนุญาต/ประกาศนียบัตร', list(p.licenses)) + row('การอบรม', list(p.trainings)) +
    row('ทักษะ', list(p.skills)) + row('ภาษา', esc_((p.languages || []).join(', '))) +
    row('ร่วมงานกับเราตั้งแต่', esc_(p.since)) + row('ตรวจประวัติแล้วเมื่อ', esc_(p.backgroundCheck)) + row('เบอร์ติดต่อ', esc_(p.phone)) + '</table>';
  h += '<p class="sm" style="margin-top:16px">เจ้าหน้าที่ให้ความยินยอมเปิดเผยประวัตินี้แก่ผู้รับบริการเพื่อประกอบการตัดสินใจ ห้ามนำไปเผยแพร่ต่อ</p>';
  h += '</td></tr></table>';
  return h;
}

function systemTestPdf_() {
  var html = pdfShell_('<div class="brand">ทดสอบ PDF ภาษาไทย</div><p>ตัวอักษรไทย: กขคงจฉชซ ฌญฎฏ ก็ ก่ ก้ ก๊ ก๋ ๐๑๒๓๔๕๖๗๘๙</p><p>ตัวเลข: 1,234,567.89 บาท (' + esc_(bahtText_(1234567.89)) + ')</p><p>ถ้าอ่านข้อความนี้ออกครบทุกตัว แปลว่า PDF ใช้งานได้</p>' +
    '<table width="100%" background="' + getWatermarkDataUrl_() + '" style="background-image:url(' + getWatermarkDataUrl_() + ');background-repeat:no-repeat;background-position:center;background-size:60%;height:120mm"><tr><td class="c">ถ้าเห็นลายน้ำวงกลมจาง ๆ ด้านหลังข้อความนี้ แปลว่าหน้าประวัติเจ้าหน้าที่แสดงลายน้ำได้</td></tr></table>');
  var pdf = savePdf_(html, 'ทดสอบ-PDF-' + todayStr_(), 'test');
  setSetting_('pdfTestedAt', nowStr_());
  return { success: true, url: pdf.url };
}

/* ============================== โมดูล: เคส / พนักงาน / การชำระเงิน ============================== */
/* ---------------- CONFIG ---------------- */
const SHEET_RECORDS = 'Records';
const SHEET_HISTORY = 'History';
const SHEET_STAFF = 'Staff';
const SHEET_PAYMENTS = 'Payments';
const DEDUPE_WINDOW_SECONDS = 300; // 5 นาที กันบันทึกซ้ำจากการกดปุ่มซ้ำ/เน็ตกระตุก

// หัวตารางชีต Payments
const PAYMENT_HEADERS = [
  'id',
  'บันทึกเมื่อ',
  'แก้ไขล่าสุดเมื่อ',
  'เคส id',
  'ชื่อผู้รับบริการ',
  'จำนวนเงิน (บาท)',
  'วันที่ชำระ',
  'ช่องทางชำระ',
  'หมายเหตุ',
  'ผู้บันทึก'
];

// map: key จากฟอร์มชำระเงิน (ภาษาอังกฤษ) -> หัวตารางชีต Payments (ภาษาไทย)
const PAYMENT_PAYLOAD_TO_HEADER = {
  recordId: 'เคส id',
  clientName: 'ชื่อผู้รับบริการ',
  amount: 'จำนวนเงิน (บาท)',
  paymentDate: 'วันที่ชำระ',
  method: 'ช่องทางชำระ',
  note: 'หมายเหตุ',
  recordedBy: 'ผู้บันทึก'
};

// หัวตารางชีต Staff
const STAFF_HEADERS = [
  'id',
  'บันทึกเมื่อ',
  'แก้ไขล่าสุดเมื่อ',
  'ชื่อ-นามสกุล',
  'ชื่อเล่น',
  'ตำแหน่ง',
  'เบอร์โทร',
  'อัตราค่าตอบแทนต่อชั่วโมง (บาท)',
  'สถานะ',
  'วันที่เริ่มงาน',
  'หมายเหตุ',
  // ---- เพิ่มใหม่: ประเภทค่าตอบแทน (รายชั่วโมง / เหมาต่อเวร) — ต่อท้ายเสมอ ห้ามแทรกกลาง ----
  'ประเภทค่าตอบแทน',
  'อัตราค่าตอบแทนต่อเวร (บาท)',
  // ---- ประวัติเจ้าหน้าที่ (แนบท้ายสัญญาได้) ----
  'รูปถ่าย (file id)',
  'ปีเกิด (พ.ศ.)',
  'การศึกษา',
  'ใบอนุญาต/ประกาศนียบัตร',
  'ประสบการณ์ (ปี)',
  'ทักษะ',
  'ภาษา',
  'การอบรม',
  'แนะนำตัว',
  'ตรวจประวัติเมื่อ',
  'ยินยอมเปิดเผยประวัติแก่ลูกค้า',
  'แสดงเบอร์โทรในประวัติ'
];

// หัวตารางที่ลูกค้าเห็นได้ในประวัติแนบท้ายสัญญา (ห้ามใส่ค่าตอบแทน/หมายเหตุภายใน)
const STAFF_PUBLIC_FIELDS_ = ['ชื่อ-นามสกุล', 'ชื่อเล่น', 'ตำแหน่ง', 'ปีเกิด (พ.ศ.)', 'การศึกษา', 'ใบอนุญาต/ประกาศนียบัตร', 'ประสบการณ์ (ปี)', 'ทักษะ', 'ภาษา', 'การอบรม', 'แนะนำตัว', 'ตรวจประวัติเมื่อ', 'วันที่เริ่มงาน'];

// map: key จากฟอร์มพนักงาน (ภาษาอังกฤษ) -> หัวตารางชีต Staff (ภาษาไทย)
const STAFF_PAYLOAD_TO_HEADER = {
  fullName: 'ชื่อ-นามสกุล',
  nickname: 'ชื่อเล่น',
  role: 'ตำแหน่ง',
  phone: 'เบอร์โทร',
  hourlyRate: 'อัตราค่าตอบแทนต่อชั่วโมง (บาท)',
  status: 'สถานะ',
  startDate: 'วันที่เริ่มงาน',
  notes: 'หมายเหตุ',
  payType: 'ประเภทค่าตอบแทน',
  shiftRate: 'อัตราค่าตอบแทนต่อเวร (บาท)',
  photoFileId: 'รูปถ่าย (file id)',
  birthYear: 'ปีเกิด (พ.ศ.)',
  education: 'การศึกษา',
  licenses: 'ใบอนุญาต/ประกาศนียบัตร',
  experienceYears: 'ประสบการณ์ (ปี)',
  skills: 'ทักษะ',
  languages: 'ภาษา',
  trainings: 'การอบรม',
  bio: 'แนะนำตัว',
  backgroundCheckDate: 'ตรวจประวัติเมื่อ',
  profileConsent: 'ยินยอมเปิดเผยประวัติแก่ลูกค้า',
  showPhone: 'แสดงเบอร์โทรในประวัติ'
};

// หัวตารางชีต Records — ลำดับนี้คือ "ความจริง" ของโครงสร้างข้อมูล
// ห้ามสลับลำดับหลังจากมีข้อมูลแล้ว (จะเติมคอลัมน์ใหม่ต่อท้ายแทน)
const RECORD_HEADERS = [
  'id',
  'บันทึกเมื่อ',
  'แก้ไขล่าสุดเมื่อ',
  'ชื่อผู้รับบริการ',
  'อายุ',
  'เพศ',
  'ที่อยู่ให้บริการ',
  'ผู้ติดต่อ',
  'ความสัมพันธ์',
  'เบอร์โทรผู้ติดต่อ',
  'ประเภทเคส',
  'แพ็กเกจ',
  'รูปแบบบริการ',
  'อุปกรณ์ทางการแพทย์',
  'อาการ/ความต้องการเพิ่มเติม',
  'โซนบริการ',
  'รายการหัตถการ',
  'ค่าบริการหัตถการ (บาท)',
  'โปรแกรมคอร์ส',
  'แพ็กเกจคอร์ส',
  'จำนวนครั้งทั้งหมด',
  'จำนวนครั้งที่ใช้ไป',
  'ราคาคอร์สรวม (บาท)',
  'วันเริ่มบริการ',
  'วันครบกำหนดสัญญา',
  'เจ้าหน้าที่ผู้ดูแล',
  'สถานะ',
  'ค่าบริการต่อเดือน (บาท)',
  'อัตรารายวัน (บาท)',
  'ค่าประเมินหน้างาน (บาท)',
  'อัตราค่าบริการวันหยุด (บาท/วัน)',
  'จำนวนวันที่คิดเพิ่ม',
  'ค่าบริการวันหยุด (บาท)',
  'เงินประกันสัญญา (บาท)',
  'สถานะเงินประกัน',
  'หมายเหตุ',
  'ผู้บันทึก',
  'ออกเอกสารในนาม',
  'รหัสคำขอ', // clientRequestId — ใช้กันบันทึกซ้ำ ไม่แสดงในแอป
  // ---- เพิ่มใหม่: แพ็กเกจรายวัน (ต่อท้ายเสมอ ห้ามแทรกกลาง) ----
  'อัตรารายวัน (แพ็กเกจรายวัน) (บาท)',
  'วิธีนับจำนวนวัน (แพ็กเกจรายวัน)',
  'จำนวนวัน (แพ็กเกจรายวัน)',
  'วันที่เริ่ม (แพ็กเกจรายวัน)',
  'วันที่สิ้นสุด (แพ็กเกจรายวัน)',
  'ยอดรวมแพ็กเกจรายวัน (บาท)',
  'รหัสเจ้าหน้าที่ผู้ดูแล',
  // ---- อนุมัติราคา: พนักงานกรอกราคาต่ำกว่าเกณฑ์ → รอเจ้าของอนุมัติ ----
  'สถานะราคา',
  'รายการที่ต่ำกว่าเกณฑ์',
  'อนุมัติราคาโดย',
  'อนุมัติราคาเมื่อ',
  'หมายเหตุการอนุมัติ',
  // ---- เงื่อนไขการทำงานของผู้ดูแล (ใช้ในสัญญา) ----
  'เวลาเริ่มงาน',
  'เวลาเลิกงาน',
  'วันทำงานต่อสัปดาห์',
  'วันหยุดประจำสัปดาห์',
  'อาหารผู้ดูแล'
];

const HISTORY_HEADERS = ['เวลา', 'การกระทำ', 'id', 'ชื่อผู้รับบริการ', 'ผู้บันทึก/แก้ไข'];

// map: key ที่ส่งมาจากฟอร์ม (ภาษาอังกฤษ) -> หัวตารางในชีต (ภาษาไทย)
const PAYLOAD_TO_HEADER = {
  clientName: 'ชื่อผู้รับบริการ',
  age: 'อายุ',
  gender: 'เพศ',
  address: 'ที่อยู่ให้บริการ',
  contactName: 'ผู้ติดต่อ',
  contactRelation: 'ความสัมพันธ์',
  contactPhone: 'เบอร์โทรผู้ติดต่อ',
  serviceTypeLabel: 'ประเภทเคส',
  packageName: 'แพ็กเกจ',
  serviceFormat: 'รูปแบบบริการ',
  notes: 'อาการ/ความต้องการเพิ่มเติม',
  zoneLabel: 'โซนบริการ',
  procedureListText: 'รายการหัตถการ',
  procedureTotal: 'ค่าบริการหัตถการ (บาท)',
  courseProgramLabel: 'โปรแกรมคอร์ส',
  courseOptionLabel: 'แพ็กเกจคอร์ส',
  courseSessionsTotal: 'จำนวนครั้งทั้งหมด',
  courseSessionsUsed: 'จำนวนครั้งที่ใช้ไป',
  courseTotalPrice: 'ราคาคอร์สรวม (บาท)',
  startDate: 'วันเริ่มบริการ',
  contractEndDate: 'วันครบกำหนดสัญญา',
  caregiver: 'เจ้าหน้าที่ผู้ดูแล',
  status: 'สถานะ',
  monthlyFee: 'ค่าบริการต่อเดือน (บาท)',
  dailyRate: 'อัตรารายวัน (บาท)',
  assessmentFee: 'ค่าประเมินหน้างาน (บาท)',
  holidayChargeRate: 'อัตราค่าบริการวันหยุด (บาท/วัน)',
  holidayChargeDays: 'จำนวนวันที่คิดเพิ่ม',
  holidayChargeAmount: 'ค่าบริการวันหยุด (บาท)',
  deposit: 'เงินประกันสัญญา (บาท)',
  depositStatus: 'สถานะเงินประกัน',
  remarks: 'หมายเหตุ',
  recordedBy: 'ผู้บันทึก',
  billingEntity: 'ออกเอกสารในนาม',
  clientRequestId: 'รหัสคำขอ',
  // ---- เพิ่มใหม่: แพ็กเกจรายวัน ----
  dailyPkgRate: 'อัตรารายวัน (แพ็กเกจรายวัน) (บาท)',
  dailyPkgDaysMode: 'วิธีนับจำนวนวัน (แพ็กเกจรายวัน)',
  dailyPkgDays: 'จำนวนวัน (แพ็กเกจรายวัน)',
  dailyPkgRangeStart: 'วันที่เริ่ม (แพ็กเกจรายวัน)',
  dailyPkgRangeEnd: 'วันที่สิ้นสุด (แพ็กเกจรายวัน)',
  dailyPkgTotal: 'ยอดรวมแพ็กเกจรายวัน (บาท)',
  staffIds: 'รหัสเจ้าหน้าที่ผู้ดูแล',
  priceStatus: 'สถานะราคา',
  priceIssues: 'รายการที่ต่ำกว่าเกณฑ์',
  priceApprovedBy: 'อนุมัติราคาโดย',
  priceApprovedAt: 'อนุมัติราคาเมื่อ',
  priceNote: 'หมายเหตุการอนุมัติ',
  workStart: 'เวลาเริ่มงาน',
  workEnd: 'เวลาเลิกงาน',
  workDaysPerWeek: 'วันทำงานต่อสัปดาห์',
  weeklyDayOff: 'วันหยุดประจำสัปดาห์',
  caregiverMeals: 'อาหารผู้ดูแล'
};

/* ---------------- ติดตั้งชีทฝั่งเคส/พนักงาน/การชำระเงิน (ปลอดภัย รันซ้ำได้ ไม่ลบข้อมูล) ---------------- */
function hcSetupSheets_() {
  [[SHEET_RECORDS, RECORD_HEADERS], [SHEET_HISTORY, HISTORY_HEADERS], [SHEET_STAFF, STAFF_HEADERS], [SHEET_PAYMENTS, PAYMENT_HEADERS]].forEach(function (pair) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName(pair[0]);
    if (sh) {
      const lastCol = Math.max(sh.getLastColumn(), 1);
      const have = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
      const missing = pair[1].filter(function (h) { return have.indexOf(h) === -1; });
      if (missing.length) {
        const start = sh.getLastColumn() + 1;
        sh.getRange(1, start, 1, missing.length).setValues([missing]).setFontWeight('bold').setBackground('#2F6F62').setFontColor('#FFFFFF');
        sh.getRange(2, start, 9999, missing.length).setNumberFormat('@');
      }
      return;
    }
    sh = ss.insertSheet(pair[0]);
    sh.getRange(1, 1, 1, pair[1].length).setValues([pair[1]]).setFontWeight('bold').setBackground('#2F6F62').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
    // Plain text ทั้งชีท กันวันที่/เบอร์โทรถูกแปลงชนิดข้อมูลเอง
    sh.getRange(2, 1, 9999, pair[1].length).setNumberFormat('@');
  });
}

/* ---------------- HTTP ENTRY POINTS ---------------- */
function hcGet_(e) {
  try {
    const type = e && e.parameter && e.parameter.type;
    if (type === 'history') {
      return jsonResponse({ ok: true, history: readSheetAsObjects(SHEET_HISTORY, HISTORY_HEADERS) });
    }
    if (type === 'staff') {
      return jsonResponse({ ok: true, staff: readSheetAsObjects(SHEET_STAFF, STAFF_HEADERS) });
    }
    if (type === 'payments') {
      return jsonResponse({ ok: true, payments: readSheetAsObjects(SHEET_PAYMENTS, PAYMENT_HEADERS) });
    }
    return jsonResponse({ ok: true, records: readSheetAsObjects(SHEET_RECORDS, RECORD_HEADERS) });
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err) });
  }
}

function hcPost_(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {

    if (payload.entity === 'staff') {
      return handleStaffPost(payload);
    }
    if (payload.entity === 'payment') {
      return handlePaymentPost(payload);
    }

    // กันบันทึกซ้ำ: ถ้า clientRequestId เดิมเคยประมวลผลไปแล้วภายใน 5 นาที ให้คืน id เดิมโดยไม่เขียนซ้ำ
    const cache = CacheService.getScriptCache();
    const dedupeKey = 'req_' + (payload.clientRequestId || '');
    if (payload.clientRequestId) {
      const cached = cache.get(dedupeKey);
      if (cached) return jsonResponse({ ok: true, id: cached, deduped: true });
    }

    const sheet = getSheet(SHEET_RECORDS);
    const now = new Date();
    let resultId;

    if (payload.deleteId) {
      // ------- ลบเรคคอร์ด -------
      resultId = String(payload.deleteId);
      const rowIndex = findRowIndexById(sheet, resultId);
      if (rowIndex === -1) {
        return jsonResponse({ ok: false, error: 'ไม่พบ id ที่ต้องการลบ: ' + resultId });
      }
      const clientNameCol = RECORD_HEADERS.indexOf('ชื่อผู้รับบริการ') + 1;
      const clientName = sheet.getRange(rowIndex, clientNameCol).getValue();
      sheet.deleteRow(rowIndex);
      appendHistory('ลบ', resultId, clientName, payload.recordedBy, now);
      return jsonResponse({ ok: true, id: resultId, deleted: true });
    }

    if (payload.id) {
      // ------- แก้ไขเรคคอร์ดเดิม -------
      resultId = String(payload.id);
      const rowIndex = findRowIndexById(sheet, resultId);
      if (rowIndex === -1) {
        return jsonResponse({ ok: false, error: 'ไม่พบ id ที่ต้องการแก้ไข: ' + resultId });
      }
      const rowValues = buildRowValues(payload, resultId, null, now);
      sheet.getRange(rowIndex, 1, 1, RECORD_HEADERS.length).setValues([rowValues]);
      appendHistory('แก้ไข', resultId, payload.clientName, payload.recordedBy, now);
    } else {
      // ------- สร้างเรคคอร์ดใหม่ -------
      resultId = generateId('REC');
      const rowValues = buildRowValues(payload, resultId, now, now);
      sheet.appendRow(rowValues);
      appendHistory('บันทึกใหม่', resultId, payload.clientName, payload.recordedBy, now);
    }

    if (payload.clientRequestId) {
      cache.put(dedupeKey, resultId, DEDUPE_WINDOW_SECONDS);
    }

    return jsonResponse({ ok: true, id: resultId });
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/* ---------------- STAFF CRUD ---------------- */
function handleStaffPost(payload) {
  const sheet = getSheet(SHEET_STAFF);
  const now = new Date();
  let resultId;

  if (payload.deleteId) {
    resultId = String(payload.deleteId);
    const rowIndex = findRowIndexById(sheet, resultId);
    if (rowIndex === -1) {
      return jsonResponse({ ok: false, error: 'ไม่พบพนักงาน id: ' + resultId });
    }
    sheet.deleteRow(rowIndex);
    return jsonResponse({ ok: true, id: resultId, deleted: true });
  }

  if (payload.id) {
    resultId = String(payload.id);
    const rowIndex = findRowIndexById(sheet, resultId);
    if (rowIndex === -1) {
      return jsonResponse({ ok: false, error: 'ไม่พบพนักงาน id: ' + resultId });
    }
    const rowValues = buildStaffRowValues(payload, resultId, null, now);
    sheet.getRange(rowIndex, 1, 1, STAFF_HEADERS.length).setValues([rowValues]);
  } else {
    resultId = generateId('STF');
    const rowValues = buildStaffRowValues(payload, resultId, now, now);
    sheet.appendRow(rowValues);
  }

  return jsonResponse({ ok: true, id: resultId });
}

function buildStaffRowValues(payload, id, createdAt, updatedAt) {
  const bag = {};
  const sentHeaders = {};
  Object.keys(STAFF_PAYLOAD_TO_HEADER).forEach(key => {
    const header = STAFF_PAYLOAD_TO_HEADER[key];
    const val = payload[key];
    sentHeaders[header] = val !== undefined;
    bag[header] = val === undefined || val === null ? '' : val;
  });

  return STAFF_HEADERS.map(h => {
    if (h === 'id') return id;
    if (h === 'บันทึกเมื่อ') return createdAt ? formatDate(createdAt) : keepExistingGeneric(SHEET_STAFF, STAFF_HEADERS, id, h);
    if (h === 'แก้ไขล่าสุดเมื่อ') return formatDate(updatedAt);
    if (!createdAt && sentHeaders[h] === false) return keepExistingGeneric(SHEET_STAFF, STAFF_HEADERS, id, h);
    return bag.hasOwnProperty(h) ? bag[h] : '';
  });
}

/* ---------------- PAYMENTS CRUD (การชำระเงิน/ใบเสร็จ) ---------------- */
function handlePaymentPost(payload) {
  const sheet = getSheet(SHEET_PAYMENTS);
  const now = new Date();
  let resultId;

  if (payload.deleteId) {
    resultId = String(payload.deleteId);
    const rowIndex = findRowIndexById(sheet, resultId);
    if (rowIndex === -1) {
      return jsonResponse({ ok: false, error: 'ไม่พบรายการชำระเงิน id: ' + resultId });
    }
    sheet.deleteRow(rowIndex);
    return jsonResponse({ ok: true, id: resultId, deleted: true });
  }

  // กันบันทึกรายการชำระเงินที่ไม่ได้ผูกกับเคสไหนเลย (recordId ว่าง) — ป้องกันข้อมูลกำพร้าที่คำนวณยอดชำระผิด
  if (!payload.recordId) {
    return jsonResponse({ ok: false, error: 'ไม่พบรหัสเคส (recordId) ในคำขอนี้ — ปฏิเสธการบันทึกเพื่อป้องกันข้อมูลไม่สมบูรณ์' });
  }

  if (payload.id) {
    resultId = String(payload.id);
    const rowIndex = findRowIndexById(sheet, resultId);
    if (rowIndex === -1) {
      return jsonResponse({ ok: false, error: 'ไม่พบรายการชำระเงิน id: ' + resultId });
    }
    const rowValues = buildPaymentRowValues(payload, resultId, null, now);
    sheet.getRange(rowIndex, 1, 1, PAYMENT_HEADERS.length).setValues([rowValues]);
  } else {
    resultId = generateId('PAY');
    const rowValues = buildPaymentRowValues(payload, resultId, now, now);
    sheet.appendRow(rowValues);
  }

  return jsonResponse({ ok: true, id: resultId });
}

function buildPaymentRowValues(payload, id, createdAt, updatedAt) {
  const bag = {};
  Object.keys(PAYMENT_PAYLOAD_TO_HEADER).forEach(key => {
    const header = PAYMENT_PAYLOAD_TO_HEADER[key];
    const val = payload[key];
    bag[header] = val === undefined || val === null ? '' : val;
  });

  return PAYMENT_HEADERS.map(h => {
    if (h === 'id') return id;
    if (h === 'บันทึกเมื่อ') return createdAt ? formatDate(createdAt) : keepExistingGeneric(SHEET_PAYMENTS, PAYMENT_HEADERS, id, h);
    if (h === 'แก้ไขล่าสุดเมื่อ') return formatDate(updatedAt);
    return bag.hasOwnProperty(h) ? bag[h] : '';
  });
}

/* ---------------- HELPERS ---------------- */
function getSheet(name) {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) { hcSetupSheets_(); sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name); }
  if (!sheet) throw new Error('ไม่พบชีตชื่อ "' + name + '" — กรุณารัน setup ก่อน');
  return sheet;
}

function readSheetAsObjects(sheetName, headers) {
  const sheet = getSheet(sheetName);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  // อ่านตามจำนวนคอลัมน์จริงในชีต (เผื่อชีตยังไม่ได้ migrate คอลัมน์ใหม่ หรือมีคอลัมน์มากกว่า headers ที่รู้จัก)
  const numCols = Math.min(headers.length, sheet.getLastColumn());
  const values = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
  return values
    .filter(row => row.some(cell => cell !== '' && cell !== null))
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => {
        let v = i < row.length ? row[i] : '';
        if (v instanceof Date) v = formatDate(v);
        obj[h] = v === null || v === undefined ? '' : v;
      });
      return obj;
    })
    .reverse(); // ล่าสุดขึ้นก่อน
}

function findRowIndexById(sheet, id) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2; // +2: offset ของ header row + 1-index
  }
  return -1;
}

function buildRowValues(payload, id, createdAt, updatedAt) {
  const bag = {};
  Object.keys(PAYLOAD_TO_HEADER).forEach(key => {
    const header = PAYLOAD_TO_HEADER[key];
    let val = payload[key];
    if (Array.isArray(val)) val = val.join(', ');
    bag[header] = val === undefined || val === null ? '' : val;
  });

  return RECORD_HEADERS.map(h => {
    if (h === 'id') return id;
    if (h === 'บันทึกเมื่อ') return createdAt ? formatDate(createdAt) : keepExisting(id, h);
    if (h === 'แก้ไขล่าสุดเมื่อ') return formatDate(updatedAt);
    return bag.hasOwnProperty(h) ? bag[h] : '';
  });
}

// เมื่อแก้ไขเรคคอร์ดเดิม ต้องคง "บันทึกเมื่อ" (วันที่สร้างครั้งแรก) เอาไว้ ไม่ทับด้วยเวลาปัจจุบัน
function keepExisting(id, headerName) {
  return keepExistingGeneric(SHEET_RECORDS, RECORD_HEADERS, id, headerName);
}

// เวอร์ชันทั่วไป ใช้ได้กับทุกชีต (Records, Staff, ฯลฯ)
function keepExistingGeneric(sheetName, headers, id, headerName) {
  const sheet = getSheet(sheetName);
  const rowIndex = findRowIndexById(sheet, id);
  if (rowIndex === -1) return '';
  const colIndex = headers.indexOf(headerName) + 1;
  const existing = sheet.getRange(rowIndex, colIndex).getValue();
  return existing instanceof Date ? formatDate(existing) : existing;
}

function appendHistory(action, id, clientName, recordedBy, when) {
  const sheet = getSheet(SHEET_HISTORY);
  sheet.appendRow([formatDate(when), action, id, clientName || '', recordedBy || '']);
}

function generateId(prefix) {
  const p = prefix || 'REC';
  return p + '-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'GMT+7', 'yyMMdd') +
    '-' + Math.floor(Math.random() * 9000 + 1000);
}

function formatDate(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone() || 'GMT+7', "yyyy-MM-dd'T'HH:mm:ss");
}

// เดิมคืนค่า ContentService — ตอนนี้ router ใน Main เป็นคนแปลงเป็น JSON จึงคืน object ตรงๆ
function jsonResponse(obj) {
  return obj;
}

/* ============================== โมดูล: บันทึกการดูแล (แอปผู้ดูแล) ============================== */
const PATIENTS_SHEET_NAME = 'Patients';
const CAREGIVERS_SHEET_NAME = 'Caregivers';
const CHECKINS_SHEET_NAME = 'CheckIns';
const SCHEDULE_SHEET_NAME = 'Schedule';
const EMERGENCIES_SHEET_NAME = 'Emergencies';
const PAYRATES_SHEET_NAME = 'PayRates';
const VITALTHRESHOLDS_SHEET_NAME = 'VitalThresholds';

// ติดตั้งชีทฝั่งบันทึกการดูแล (เรียกจาก setup())
function cgSetupSheets_() {
  ensurePatientsSheet_();
  ensureCaregiversSheet_();
  ensureCheckInsSheet_();
  ensureScheduleSheet_();
  ensureEmergenciesSheet_();
  ensurePayRatesSheet_();
  ensureVitalThresholdsSheet_();
}

function ensurePatientsSheet_() {
  let sheet = SS.getSheetByName(PATIENTS_SHEET_NAME);
  if (sheet) return sheet;
  sheet = SS.insertSheet(PATIENTS_SHEET_NAME);
  sheet.appendRow(['PatientID', 'PatientName', 'Zone', 'Categories', 'FamilyEmail']);
  sheet.appendRow(['P001', 'ตัวอย่างคนไข้ (แก้/ลบแถวนี้ได้)', 'A', '', '']);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
  sheet.autoResizeColumns(1, 5);
  return sheet;
}

function ensureCaregiversSheet_() {
  let sheet = SS.getSheetByName(CAREGIVERS_SHEET_NAME);
  if (sheet) return sheet;
  sheet = SS.insertSheet(CAREGIVERS_SHEET_NAME);
  sheet.appendRow(['Code', 'Name', 'AssignedPatients']);
  sheet.appendRow([String(Math.floor(100000 + Math.random() * 900000)), 'ผู้ดูแลตัวอย่าง (แก้/ลบแถวนี้ได้)', '']);
  sheet.getRange(2, 1, 998, 1).setNumberFormat('@');
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, 3).setFontWeight('bold');
  sheet.autoResizeColumns(1, 3);
  return sheet;
}

function ensureCheckInsSheet_() {
  let sheet = SS.getSheetByName(CHECKINS_SHEET_NAME);
  if (sheet) return sheet;
  sheet = SS.insertSheet(CHECKINS_SHEET_NAME);
  sheet.appendRow(['วันที่', 'เวลา', 'ชื่อผู้ดูแล', 'ชื่อคนไข้', 'เวรที่เข้า', 'รหัสไฟล์รูปถ่าย', 'บันทึกเมื่อ']);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, 7).setFontWeight('bold');
  sheet.autoResizeColumns(1, 7);
  return sheet;
}

function ensureScheduleSheet_() {
  let sheet = SS.getSheetByName(SCHEDULE_SHEET_NAME);
  if (sheet) return sheet;
  sheet = SS.insertSheet(SCHEDULE_SHEET_NAME);
  sheet.appendRow(['ID', 'วันที่', 'เวรที่มอบหมาย', 'ชื่อคนไข้', 'รหัสผู้ดูแล', 'ชื่อผู้ดูแล', 'หมายเหตุ', 'จัดตารางเมื่อ']);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, 8).setFontWeight('bold');
  sheet.autoResizeColumns(1, 8);
  return sheet;
}

function ensureEmergenciesSheet_() {
  let sheet = SS.getSheetByName(EMERGENCIES_SHEET_NAME);
  if (sheet) return sheet;
  sheet = SS.insertSheet(EMERGENCIES_SHEET_NAME);
  sheet.appendRow(['ID', 'วันที่', 'เวลา', 'ชื่อผู้ดูแล', 'คนไข้', 'ข้อความ', 'ละติจูด', 'ลองจิจูด', 'สถานะ', 'รับทราบเมื่อ', 'รับทราบโดย', 'บันทึกเมื่อ', 'ประเภท']);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, 13).setFontWeight('bold');
  sheet.autoResizeColumns(1, 13);
  return sheet;
}

// ค่าปกติของสัญญาณชีพ — ใช้ค่าอ้างอิงทั่วไปเป็นค่าเริ่มต้นเท่านั้น แอดมินควรปรับตามคำสั่งแพทย์ของคนไข้แต่ละราย
// ค่าอ้างอิงตามแนวทางที่คลินิกใช้จริง — แอดมินยังปรับได้ทุกค่าผ่านหน้าเว็บ
const DEFAULT_VITAL_THRESHOLDS_ = [
  { metric: 'bp_systolic', label: 'ความดันโลหิต (ตัวบน)', min: 90, max: 140, unit: 'mmHg' },
  { metric: 'bp_diastolic', label: 'ความดันโลหิต (ตัวล่าง)', min: 60, max: 90, unit: 'mmHg' },
  { metric: 'heart_rate', label: 'ชีพจร (PR)', min: 60, max: 100, unit: 'ครั้ง/นาที' },
  { metric: 'respiratory_rate', label: 'อัตราการหายใจ (RR)', min: 16, max: 24, unit: 'ครั้ง/นาที' },
  { metric: 'temperature', label: 'อุณหภูมิ (T)', min: 36.5, max: 37.4, unit: '°C' },
  { metric: 'spo2', label: 'ออกซิเจนปลายนิ้ว (SpO2)', min: 95, max: 100, unit: '%' },
  { metric: 'sugar', label: 'ระดับน้ำตาลในเลือด (DTX)', min: 90, max: 190, unit: 'mg%' }
];

function ensureVitalThresholdsSheet_() {
  let sheet = SS.getSheetByName(VITALTHRESHOLDS_SHEET_NAME);
  if (!sheet) {
    sheet = SS.insertSheet(VITALTHRESHOLDS_SHEET_NAME);
    sheet.appendRow(['ตัวชี้วัด', 'ชื่อ', 'ค่าต่ำสุดปกติ', 'ค่าสูงสุดปกติ', 'หน่วย']);
    DEFAULT_VITAL_THRESHOLDS_.forEach(function (t) { sheet.appendRow([t.metric, t.label, t.min, t.max, t.unit]); });
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
    sheet.autoResizeColumns(1, 5);
    return sheet;
  }
  // ชีทมีอยู่แล้ว (เช่นสร้างไว้ก่อนเพิ่มตัวชี้วัดใหม่) — เติมเฉพาะแถวที่ยังไม่มี ไม่แตะค่าที่แอดมินปรับไว้แล้ว
  const lastRow = sheet.getLastRow();
  const existing = lastRow >= 2
    ? sheet.getRange(2, 1, lastRow - 1, 1).getValues().map(function (r) { return String(r[0] || ''); })
    : [];
  DEFAULT_VITAL_THRESHOLDS_.forEach(function (t) {
    if (existing.indexOf(t.metric) === -1) {
      sheet.appendRow([t.metric, t.label, t.min, t.max, t.unit]);
    }
  });
  return sheet;
}

// อัตราค่าตอบแทนต่อเวร — แอดมินตั้งราคาผ่านหน้าเว็บได้ (ไม่ต้องมาแก้ชีทเอง)
const SHIFT_TYPES_ = [
  'เวรเช้า 8 ชั่วโมง', 'เวรเช้า 10 ชั่วโมง', 'เวรเช้า 12 ชั่วโมง',
  'เวรดึก 8 ชั่วโมง', 'เวรดึก 10 ชั่วโมง', 'เวรดึก 12 ชั่วโมง',
  'เวรทั้งวัน 24 ชั่วโมง'
];

function ensurePayRatesSheet_() {
  let sheet = SS.getSheetByName(PAYRATES_SHEET_NAME);
  if (sheet) return sheet;
  sheet = SS.insertSheet(PAYRATES_SHEET_NAME);
  sheet.appendRow(['เวร', 'อัตรา (บาท/เวร)']);
  SHIFT_TYPES_.forEach(function (s) { sheet.appendRow([s, 0]); });
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, 2).setFontWeight('bold');
  sheet.autoResizeColumns(1, 2);
  return sheet;
}

// ---------------------------------------------------------------------------
// Auth
// ---------------------------------------------------------------------------
function verifyCaregiverCode(code) {
  const sheet = ensureCaregiversSheet_();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(code).trim() && String(code).trim() !== '') {
      const assignedStr = String(data[i][2] || '').trim();
      const assignedPatients = assignedStr
        ? assignedStr.split(',').map(function (s) { return s.trim(); }).filter(Boolean)
        : null; // null = ดูแลได้ทุกคนไข้
      return { success: true, name: data[i][1], assignedPatients: assignedPatients };
    }
  }
  return { success: false, message: 'รหัสไม่ถูกต้อง' };
}

// แอดมินยืนยันตัวตนผ่านรหัสผ่านและ session แล้ว (ดู Main) — ฟังก์ชันนี้แค่ตรวจว่า session ปัจจุบันเป็นแอดมิน
function verifyAdminCode() {
  if (CURRENT_SESSION_ && (CURRENT_SESSION_.role === 'admin' || CURRENT_SESSION_.role === 'officer')) return { success: true, name: CURRENT_SESSION_.name };
  return { success: false, message: 'ไม่มีสิทธิ์แอดมิน' };
}

// ---------------------------------------------------------------------------
// Patients
// ---------------------------------------------------------------------------
function getPatients(caregiverCode) {
  const sheet = ensurePatientsSheet_();
  const data = sheet.getDataRange().getValues();

  let allowedNames = null; // null = ไม่จำกัด (แอดมิน หรือ ผู้ดูแลที่ไม่ได้ผูกคนไข้เฉพาะ)
  if (caregiverCode) {
    const cg = verifyCaregiverCode(caregiverCode);
    if (cg.success) allowedNames = cg.assignedPatients;
  }

  const patients = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][1]) continue;
    const name = String(data[i][1]);
    if (allowedNames && allowedNames.indexOf(name.trim()) === -1) continue;
    const catStr = String(data[i][3] || '').trim();
    patients.push({
      id: data[i][0],
      name: data[i][1],
      zone: data[i][2] || '',
      categories: catStr ? catStr.split(',').map(function (s) { return s.trim(); }).filter(Boolean) : null // null = ทุกหมวด
    });
  }
  return patients;
}

// ---------------------------------------------------------------------------
// Records
// ---------------------------------------------------------------------------
// ---------------------------------------------------------------------------
// แอดมิน: ลงทะเบียนผู้ดูแล/คนไข้ใหม่ — ต้องแนบ adminCode ที่ผ่านการยืนยันแล้วมาด้วยทุกครั้ง
// ---------------------------------------------------------------------------
function registerCaregiver(payload) {
  if (!payload || !payload.adminCode) return { success: false, message: 'ไม่มีสิทธิ์ดำเนินการ' };
  const admin = verifyAdminCode(payload.adminCode);
  if (!admin.success) return { success: false, message: 'รหัสแอดมินไม่ถูกต้อง' };

  const code = String(payload.code || '').trim();
  const name = String(payload.name || '').trim();
  if (!code || !name) return { success: false, message: 'กรุณากรอกรหัสและชื่อผู้ดูแล' };
  if (code.length < 6) return { success: false, message: 'รหัสผู้ดูแลต้องยาวอย่างน้อย 6 ตัว เพื่อความปลอดภัย' };

  const sheet = ensureCaregiversSheet_();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === code) {
      return { success: false, message: 'รหัสนี้มีผู้ใช้งานอยู่แล้ว กรุณาตั้งรหัสอื่น' };
    }
  }

  const assignedPatients = Array.isArray(payload.assignedPatients) ? payload.assignedPatients : [];
  sheet.appendRow(['', name, assignedPatients.join(', ')]);
  sheet.getRange(sheet.getLastRow(), 1).setNumberFormat('@').setValue(code);
  return { success: true };
}

// แอดมิน: รายชื่อผู้รับบริการพร้อมโซน หมวดที่บันทึก และอีเมลญาติ (ใช้ส่ง SOS / ค่าผิดปกติ)
const CARE_CATEGORY_KEYS_ = ['vitals', 'nutrition', 'medicine', 'excretion', 'hygiene', 'mobility', 'devices', 'general', 'incident'];
function getPatientsAdmin_() {
  const data = ensurePatientsSheet_().getDataRange().getValues();
  const out = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][1]) continue;
    out.push({
      id: String(data[i][0] || ''), name: String(data[i][1]), zone: String(data[i][2] || ''),
      categories: String(data[i][3] || '').split(',').map(function (s) { return s.trim(); }).filter(Boolean),
      familyEmail: String(data[i][4] || '')
    });
  }
  return { success: true, patients: out, categoryKeys: CARE_CATEGORY_KEYS_ };
}

function savePatientAdmin_(d) {
  const name = String(d.name || '').trim();
  if (!name) fail_('กรุณาระบุชื่อผู้รับบริการ');
  const emails = String(d.familyEmail || '').split(',').map(function (x) { return x.trim(); }).filter(Boolean);
  emails.forEach(function (e) { if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e)) fail_('อีเมลไม่ถูกต้อง: ' + e); });
  const cats = (Array.isArray(d.categories) ? d.categories : []).filter(function (k) { return CARE_CATEGORY_KEYS_.indexOf(k) >= 0; });
  const zone = String(d.zone || '').trim().slice(0, 40);
  const sheet = ensurePatientsSheet_();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if ((d.id && String(data[i][0]) === String(d.id)) || (!d.id && String(data[i][1]).trim() === name)) {
      // ไม่ให้เปลี่ยนชื่อ เพราะชื่อผูกกับชีทบันทึกและการมอบหมายผู้ดูแล
      sheet.getRange(i + 1, 3, 1, 3).setValues([[zone, cats.join(','), emails.join(', ')]]);
      return getPatientsAdmin_();
    }
  }
  const newId = 'P' + ('000' + data.length).slice(-3);
  sheet.appendRow([newId, name, zone, cats.join(','), emails.join(', ')]);
  return getPatientsAdmin_();
}

function registerPatient(payload) {
  if (!payload || !payload.adminCode) return { success: false, message: 'ไม่มีสิทธิ์ดำเนินการ' };
  const admin = verifyAdminCode(payload.adminCode);
  if (!admin.success) return { success: false, message: 'รหัสแอดมินไม่ถูกต้อง' };

  const name = String(payload.name || '').trim();
  if (!name) return { success: false, message: 'กรุณากรอกชื่อคนไข้' };

  const sheet = ensurePatientsSheet_();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim() === name) {
      return { success: false, message: 'มีคนไข้ชื่อนี้อยู่แล้วในระบบ' };
    }
  }

  const nextNum = data.length; // header กินแถวแรกไปแล้ว พอดีเป็นลำดับถัดไป
  const newId = 'P' + ('000' + nextNum).slice(-3);
  const categories = Array.isArray(payload.categories) ? payload.categories : [];
  sheet.appendRow([newId, name, payload.zone || '', categories.join(','), payload.familyEmail || '']);
  return { success: true, id: newId };
}

// ---------------------------------------------------------------------------
// เช็คอินเข้างาน: ผู้ดูแลถ่ายรูปตัวเองหน้าบ้านคนไข้ + เลือกคนไข้ + เลือกเวรที่เข้า
// ---------------------------------------------------------------------------
function checkIn(payload) {
  if (!payload || !payload.caregiverCode) return { success: false, message: 'กรุณาเข้าสู่ระบบก่อน' };
  const cg = verifyCaregiverCode(payload.caregiverCode);
  if (!cg.success) return { success: false, message: 'รหัสผู้ดูแลไม่ถูกต้อง' };

  const patientName = String(payload.patientName || '').trim();
  const shiftType = String(payload.shiftType || '').trim();
  if (!patientName) return { success: false, message: 'กรุณาเลือกคนไข้' };
  if (!shiftType) return { success: false, message: 'กรุณาเลือกเวรที่เข้า' };
  if (cg.assignedPatients && cg.assignedPatients.indexOf(patientName) === -1) {
    return { success: false, message: 'คุณไม่มีสิทธิ์เช็คอินให้คนไข้ท่านนี้' };
  }
  if (!payload.photoDataUrl) return { success: false, message: 'กรุณาถ่ายรูปก่อนเช็คอิน' };

  const match = String(payload.photoDataUrl).match(/^data:(image\/[\w+.-]+);base64,(.+)$/);
  if (!match) return { success: false, message: 'รูปภาพไม่ถูกต้อง กรุณาถ่ายใหม่อีกครั้ง' };
  const mimeType = match[1];
  const base64Data = match[2];

  const now = new Date();
  const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm');

  const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, 'checkin.jpg');
  const folder = driveFolder_('checkin-photos');
  const fileName = 'checkin_' + cg.name + '_' + Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
  const file = folder.createFile(blob).setName(fileName);
  // รูปเก็บแบบส่วนตัว ไม่แชร์ลิงก์สาธารณะ — แอดมินเปิดดูผ่าน care.checkinPhoto
  const sheet = ensureCheckInsSheet_();
  sheet.appendRow([dateStr, timeStr, cg.name, patientName, shiftType, file.getId(), now]);

  return { success: true, time: timeStr, date: dateStr };
}

// สำหรับแอดมิน: ดึงประวัติเช็คอินล่าสุด (ใหม่ไปเก่า) มาแสดงพร้อมรูป
function getCheckIns(limit) {
  const sheet = ensureCheckInsSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: true, records: [] };

  const values = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
  let records = values.map(function (r) {
    const loggedAtDate = r[6] instanceof Date ? r[6] : null;
    return {
      date: formatDate_(r[0]),
      time: formatTimeValue_(r[1]),
      caregiverName: String(r[2] || ''),
      patientName: String(r[3] || ''),
      shiftType: String(r[4] || ''),
      photoId: String(r[5] || ''),
      loggedAtSort: loggedAtDate ? loggedAtDate.getTime() : 0
    };
  });
  records.sort(function (a, b) { return b.loggedAtSort - a.loggedAtSort; });
  if (limit) records = records.slice(0, limit);
  records.forEach(function (r) { delete r.loggedAtSort; });

  return { success: true, records: records };
}

// สำหรับแอดมิน: เปิดดูรูปเช็คอิน 1 รูป (ส่งกลับเป็น data URL ไม่ต้องแชร์ไฟล์สาธารณะ)
function getCheckinPhoto(photoId) {
  const id = String(photoId || '').trim();
  if (!id) return { success: false, message: 'ไม่พบรูป' };
  const sheet = ensureCheckInsSheet_();
  const ids = sheet.getLastRow() >= 2 ? sheet.getRange(2, 6, sheet.getLastRow() - 1, 1).getValues().map(function (r) { return String(r[0]); }) : [];
  if (ids.indexOf(id) === -1) return { success: false, message: 'ไม่พบรูป' };
  const blob = DriveApp.getFileById(id).getBlob();
  return { success: true, dataUrl: 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes()) };
}

// ---------------------------------------------------------------------------
// ตารางเวรล่วงหน้า — แอดมินจัดเวรผู้ดูแลแต่ละคนไข้ไว้ล่วงหน้าเป็นรายวัน
// ---------------------------------------------------------------------------

// รายชื่อ+รหัสผู้ดูแลทั้งหมด (ให้แอดมินเลือกตอนจัดตาราง) — ต้องแนบ adminCode เพราะมีรหัสล็อกอินอยู่ในผลลัพธ์
function getCaregiversList(adminCode) {
  const admin = verifyAdminCode(adminCode);
  if (!admin.success) return { success: false, message: 'รหัสแอดมินไม่ถูกต้อง' };

  const sheet = ensureCaregiversSheet_();
  const data = sheet.getDataRange().getValues();
  const caregivers = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][1]) continue;
    caregivers.push({ code: String(data[i][0] || ''), name: String(data[i][1] || '') });
  }
  return { success: true, caregivers: caregivers };
}

function addScheduleEntry(payload) {
  if (!payload || !payload.adminCode) return { success: false, message: 'ไม่มีสิทธิ์ดำเนินการ' };
  const admin = verifyAdminCode(payload.adminCode);
  if (!admin.success) return { success: false, message: 'รหัสแอดมินไม่ถูกต้อง' };

  const date = String(payload.date || '').trim();
  const shiftType = String(payload.shiftType || '').trim();
  const patientName = String(payload.patientName || '').trim();
  const caregiverCode = String(payload.caregiverCode || '').trim();
  if (!date || !shiftType || !patientName || !caregiverCode) {
    return { success: false, message: 'กรุณากรอกวันที่ เวร คนไข้ และผู้ดูแลให้ครบ' };
  }

  const cg = verifyCaregiverCode(caregiverCode);
  if (!cg.success) return { success: false, message: 'ไม่พบผู้ดูแลคนนี้' };

  const sheet = ensureScheduleSheet_();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (formatDate_(data[i][1]) === date && String(data[i][2]) === shiftType && String(data[i][4]) === caregiverCode) {
      return { success: false, message: 'ผู้ดูแลคนนี้มีเวรนี้ในวันนี้อยู่แล้ว' };
    }
  }

  const now = new Date();
  const id = 'S' + now.getTime() + Math.floor(Math.random() * 1000);
  sheet.appendRow([id, date, shiftType, patientName, caregiverCode, cg.name, String(payload.note || ''), now]);

  return { success: true, id: id };
}

function deleteScheduleEntry(payload) {
  if (!payload || !payload.adminCode) return { success: false, message: 'ไม่มีสิทธิ์ดำเนินการ' };
  const admin = verifyAdminCode(payload.adminCode);
  if (!admin.success) return { success: false, message: 'รหัสแอดมินไม่ถูกต้อง' };
  const id = String(payload.id || '').trim();
  if (!id) return { success: false, message: 'ไม่พบรายการที่จะลบ' };

  const sheet = ensureScheduleSheet_();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === id) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, message: 'ไม่พบรายการนี้แล้ว (อาจถูกลบไปก่อนหน้า)' };
}

// ดึงตารางเวรในช่วงวันที่ที่กำหนด — ใส่ caregiverCode เพื่อกรองเฉพาะเวรของผู้ดูแลคนนั้น (ฝั่งผู้ดูแลดูเวรตัวเอง)
// ไม่ใส่ caregiverCode = ดึงทั้งหมด (ใช้ฝั่งแอดมิน)
function getSchedule(dateFrom, dateTo, caregiverCode) {
  const sheet = ensureScheduleSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: true, entries: [] };

  const values = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
  let entries = values.map(function (r) {
    return {
      id: String(r[0] || ''),
      date: formatDate_(r[1]),
      shiftType: String(r[2] || ''),
      patientName: String(r[3] || ''),
      caregiverCode: String(r[4] || ''),
      caregiverName: String(r[5] || ''),
      note: String(r[6] || '')
    };
  });

  entries = entries.filter(function (e) {
    if (dateFrom && e.date < dateFrom) return false;
    if (dateTo && e.date > dateTo) return false;
    if (caregiverCode && e.caregiverCode !== caregiverCode) return false;
    return true;
  });

  entries.sort(function (a, b) {
    if (a.date !== b.date) return a.date < b.date ? -1 : 1;
    return a.shiftType < b.shiftType ? -1 : (a.shiftType > b.shiftType ? 1 : 0);
  });

  return { success: true, entries: entries };
}

// ---------------------------------------------------------------------------
// ปุ่มฉุกเฉิน (SOS) — ผู้ดูแลกดค้างเพื่อแจ้งเหตุ ระบบส่งอีเมลแจ้งร้านและญาติคนไข้ทันที
// ---------------------------------------------------------------------------
function triggerSOS(payload) {
  if (!payload || !payload.caregiverCode) return { success: false, message: 'กรุณาเข้าสู่ระบบก่อน' };
  const cg = verifyCaregiverCode(payload.caregiverCode);
  if (!cg.success) return { success: false, message: 'รหัสผู้ดูแลไม่ถูกต้อง' };

  const now = new Date();
  const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm');
  const patientName = String(payload.patientName || '').trim();
  const message = String(payload.message || '').trim();
  const lat = payload.lat != null ? String(payload.lat) : '';
  const lng = payload.lng != null ? String(payload.lng) : '';

  const sheet = ensureEmergenciesSheet_();
  const id = 'E' + now.getTime() + Math.floor(Math.random() * 1000);
  sheet.appendRow([id, dateStr, timeStr, cg.name, patientName, message, lat, lng, 'ยังไม่รับทราบ', '', '', now, 'SOS']);

  const mapsLink = (lat && lng) ? ('https://maps.google.com/?q=' + lat + ',' + lng) : '';
  const alertText =
    '🚨 แจ้งเหตุฉุกเฉิน ' + getCompanyName_() + '\n' +
    'ผู้ดูแล: ' + cg.name + '\n' +
    (patientName ? ('คนไข้: ' + patientName + '\n') : '') +
    'เวลา: ' + dateStr + ' ' + timeStr + ' น.\n' +
    (message ? ('ข้อความ: ' + message + '\n') : '') +
    (mapsLink ? ('ตำแหน่ง: ' + mapsLink) : '');

  const recipients = collectEmergencyRecipients_(patientName);
  recipients.emails.forEach(function (email) {
    sendEmergencyEmail_(email, '🚨 แจ้งเหตุฉุกเฉิน — ' + cg.name, alertText);
  });

  return { success: true, id: id, notified: recipients.emails.length };
}

// รวบรวมอีเมลแจ้งเตือน: อีเมลร้าน (ตั้งค่า > notifyEmail คั่นด้วย comma) + อีเมลญาติของคนไข้คนนั้น (ชีท Patients)
// หมายเหตุ: LINE Notify ปิดบริการแล้ว (31 มี.ค. 2025) จึงแจ้งเตือนทางอีเมลเท่านั้น
function collectEmergencyRecipients_(patientName) {
  const emails = String(getSettings_().notifyEmail || '').split(',').map(function (x) { return x.trim(); }).filter(Boolean);
  if (patientName) {
    const pdata = ensurePatientsSheet_().getDataRange().getValues();
    for (let i = 1; i < pdata.length; i++) {
      if (String(pdata[i][1]).trim() === patientName) {
        const famEmail = String(pdata[i][4] || '').trim();
        if (famEmail) emails.push(famEmail);
        break;
      }
    }
  }
  return { emails: emails };
}

function sendEmergencyEmail_(email, subject, body) {
  if (!email) return;
  try { MailApp.sendEmail(email, subject, body); } catch (e) { /* ไม่ให้การส่งพลาดหยุดทั้งฟังก์ชัน */ }
}

function acknowledgeEmergency(payload) {
  if (!payload || !payload.adminCode) return { success: false, message: 'ไม่มีสิทธิ์ดำเนินการ' };
  const admin = verifyAdminCode(payload.adminCode);
  if (!admin.success) return { success: false, message: 'รหัสแอดมินไม่ถูกต้อง' };
  const id = String(payload.id || '').trim();
  if (!id) return { success: false, message: 'ไม่พบรายการ' };

  const sheet = ensureEmergenciesSheet_();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === id) {
      const now = new Date();
      const ackAt = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
      sheet.getRange(i + 1, 9, 1, 3).setValues([['รับทราบแล้ว', ackAt, admin.name]]);
      return { success: true };
    }
  }
  return { success: false, message: 'ไม่พบรายการนี้แล้ว' };
}

function getEmergencies(limit) {
  const sheet = ensureEmergenciesSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: true, records: [] };

  const values = sheet.getRange(2, 1, lastRow - 1, 13).getValues();
  let records = values.map(function (r) {
    const sortDate = r[11] instanceof Date ? r[11] : null;
    return {
      id: String(r[0] || ''),
      date: String(r[1] || ''),
      time: String(r[2] || ''),
      caregiverName: String(r[3] || ''),
      patientName: String(r[4] || ''),
      message: String(r[5] || ''),
      lat: String(r[6] || ''),
      lng: String(r[7] || ''),
      status: String(r[8] || ''),
      ackAt: String(r[9] || ''),
      ackBy: String(r[10] || ''),
      type: String(r[12] || '') || 'SOS',
      sortTs: sortDate ? sortDate.getTime() : 0
    };
  });
  records.sort(function (a, b) { return b.sortTs - a.sortTs; });
  if (limit) records = records.slice(0, limit);
  records.forEach(function (r) { delete r.sortTs; });

  return { success: true, records: records };
}

// ---------------------------------------------------------------------------
// ค่าปกติสัญญาณชีพ + แจ้งเตือนอัตโนมัติเมื่อบันทึกค่าผิดปกติ (ขึ้นรวมกับหน้าแจ้งเหตุฉุกเฉินฝั่งแอดมิน)
// ---------------------------------------------------------------------------
function getVitalThresholds() {
  const sheet = ensureVitalThresholdsSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: true, thresholds: [] };
  const values = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  const thresholds = values.map(function (r) {
    return { metric: String(r[0] || ''), label: String(r[1] || ''), min: Number(r[2]), max: Number(r[3]), unit: String(r[4] || '') };
  }).filter(function (t) { return t.metric; });
  return { success: true, thresholds: thresholds };
}

function saveVitalThresholds(payload) {
  if (!payload || !payload.adminCode) return { success: false, message: 'ไม่มีสิทธิ์ดำเนินการ' };
  const admin = verifyAdminCode(payload.adminCode);
  if (!admin.success) return { success: false, message: 'รหัสแอดมินไม่ถูกต้อง' };
  const thresholds = Array.isArray(payload.thresholds) ? payload.thresholds : [];
  if (thresholds.length === 0) return { success: false, message: 'ไม่มีข้อมูลค่าปกติ' };

  const sheet = ensureVitalThresholdsSheet_();
  const byMetric = {};
  thresholds.forEach(function (t) { byMetric[String(t.metric || '')] = t; });

  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const values = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
    for (let i = 0; i < values.length; i++) {
      const metric = String(values[i][0] || '');
      if (byMetric.hasOwnProperty(metric)) {
        sheet.getRange(i + 2, 3, 1, 2).setValues([[Number(byMetric[metric].min) || 0, Number(byMetric[metric].max) || 0]]);
        delete byMetric[metric];
      }
    }
  }
  return { success: true };
}

// ดึงตัวเลขทั้งหมดจากข้อความอิสระที่ผู้ดูแลพิมพ์ เช่น "120/80" -> [120, 80], "36.5 องศา" -> [36.5]
function extractNumbers_(text) {
  const matches = String(text || '').match(/-?\d+(\.\d+)?/g);
  return matches ? matches.map(Number) : [];
}

// รายการที่ตรงชื่อในหมวดสัญญาณชีพ -> คีย์ค่าปกติ (ยกเว้นความดันโลหิตซึ่งมีตัวบน/ตัวล่างแยกกัน จัดการเป็นกรณีพิเศษ)
const VITAL_METRIC_MAP_ = {
  'อัตราการเต้นของหัวใจ': 'heart_rate',
  'อัตราการหายใจ': 'respiratory_rate',
  'อุณหภูมิ': 'temperature',
  'ออกซิเจนปลายนิ้ว (SpO2)': 'spo2',
  'ระดับน้ำตาลในเลือด': 'sugar'
};

// ตรวจรายการที่บันทึกว่ามีค่าสัญญาณชีพตัวไหนหลุดกรอบปกติไหม คืนเป็นข้อความอธิบายทีละรายการ
function checkVitalsAbnormal_(items) {
  const thresholdsRes = getVitalThresholds();
  const byMetric = {};
  thresholdsRes.thresholds.forEach(function (t) { byMetric[t.metric] = t; });

  const alerts = [];
  (items || []).forEach(function (it) {
    if (it.category !== 'สัญญาณชีพ' || !it.detail) return;

    if (it.item === 'ความดันโลหิต') {
      const nums = extractNumbers_(it.detail);
      if (nums.length >= 2) {
        const sys = nums[0], dia = nums[1];
        const tSys = byMetric['bp_systolic'], tDia = byMetric['bp_diastolic'];
        if (tSys && (sys < tSys.min || sys > tSys.max)) {
          alerts.push('ความดันโลหิตตัวบน ' + sys + ' ' + tSys.unit + ' (ค่าปกติ ' + tSys.min + '-' + tSys.max + ')');
        }
        if (tDia && (dia < tDia.min || dia > tDia.max)) {
          alerts.push('ความดันโลหิตตัวล่าง ' + dia + ' ' + tDia.unit + ' (ค่าปกติ ' + tDia.min + '-' + tDia.max + ')');
        }
      }
      return;
    }

    const metricKey = VITAL_METRIC_MAP_[it.item];
    if (!metricKey || !byMetric[metricKey]) return;
    const nums = extractNumbers_(it.detail);
    if (nums.length === 0) return;
    const val = nums[0];
    const t = byMetric[metricKey];
    if (val < t.min || val > t.max) {
      alerts.push(t.label + ' ' + val + ' ' + t.unit + ' (ค่าปกติ ' + t.min + '-' + t.max + ')');
    }
  });

  return alerts;
}

// บันทึกแจ้งเตือนค่าสัญญาณชีพผิดปกติ เข้าไปในชีท Emergencies ชุดเดียวกับ SOS (แยกด้วยคอลัมน์ ประเภท)
// แล้วส่งแจ้งเตือนช่องทางเดียวกับ SOS (อีเมล + LINE ไปหาแอดมิน/ญาติของคนไข้)
function logVitalAlert_(patientName, caregiverName, alertLines) {
  const now = new Date();
  const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm');
  const message = alertLines.join(' / ');

  const sheet = ensureEmergenciesSheet_();
  const id = 'V' + now.getTime() + Math.floor(Math.random() * 1000);
  sheet.appendRow([id, dateStr, timeStr, caregiverName, patientName, message, '', '', 'ยังไม่รับทราบ', '', '', now, 'VITALS']);

  const alertText =
    '⚠️ ค่าสัญญาณชีพผิดปกติ — ' + getCompanyName_() + '\n' +
    'คนไข้: ' + patientName + '\n' +
    'ผู้บันทึก: ' + caregiverName + '\n' +
    'เวลา: ' + dateStr + ' ' + timeStr + ' น.\n' +
    alertLines.join('\n');

  const recipients = collectEmergencyRecipients_(patientName);
  recipients.emails.forEach(function (email) {
    sendEmergencyEmail_(email, '⚠️ ค่าสัญญาณชีพผิดปกติ — ' + patientName, alertText);
  });
}

// บัญชี/ค่าตอบแทนผู้ดูแล — คำนวณจากประวัติเช็คอินจริง (ไม่ใช่ตารางที่วางแผนไว้) x อัตราต่อเวร
// ---------------------------------------------------------------------------
function getPayRates() {
  const sheet = ensurePayRatesSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: true, rates: [] };
  const values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  const rates = values.map(function (r) {
    return { shiftType: String(r[0] || ''), rate: Number(r[1]) || 0 };
  }).filter(function (r) { return r.shiftType; });
  return { success: true, rates: rates };
}

function savePayRates(payload) {
  if (!payload || !payload.adminCode) return { success: false, message: 'ไม่มีสิทธิ์ดำเนินการ' };
  const admin = verifyAdminCode(payload.adminCode);
  if (!admin.success) return { success: false, message: 'รหัสแอดมินไม่ถูกต้อง' };
  const rates = Array.isArray(payload.rates) ? payload.rates : [];
  if (rates.length === 0) return { success: false, message: 'ไม่มีข้อมูลอัตราค่าตอบแทน' };

  const sheet = ensurePayRatesSheet_();
  const byType = {};
  rates.forEach(function (r) { byType[String(r.shiftType || '')] = Number(r.rate) || 0; });

  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    for (let i = 0; i < values.length; i++) {
      const type = String(values[i][0] || '');
      if (byType.hasOwnProperty(type)) {
        sheet.getRange(i + 2, 2).setValue(byType[type]);
        delete byType[type];
      }
    }
  }
  // เวรใหม่ที่ยังไม่เคยมีในชีท ให้เพิ่มแถวใหม่ต่อท้าย
  Object.keys(byType).forEach(function (type) {
    sheet.appendRow([type, byType[type]]);
  });

  return { success: true };
}

// ดึงยอดค่าตอบแทนของผู้ดูแลแต่ละคน ในช่วงวันที่ที่กำหนด อิงจากประวัติเช็คอินจริง (CheckIns)
function getPayrollSummary(dateFrom, dateTo) {
  const rateRes = getPayRates();
  const rateMap = {};
  rateRes.rates.forEach(function (r) { rateMap[r.shiftType] = r.rate; });

  const sheet = ensureCheckInsSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: true, rows: [], grandTotal: 0 };

  const values = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
  const byCaregiver = {};
  const order = [];

  values.forEach(function (r) {
    const date = formatDate_(r[0]);
    if (dateFrom && date < dateFrom) return;
    if (dateTo && date > dateTo) return;
    const caregiverName = String(r[2] || '');
    const shiftType = String(r[4] || '');
    if (!caregiverName) return;

    if (!byCaregiver[caregiverName]) {
      byCaregiver[caregiverName] = { caregiverName: caregiverName, shifts: {}, total: 0 };
      order.push(caregiverName);
    }
    const rate = rateMap.hasOwnProperty(shiftType) ? rateMap[shiftType] : 0;
    if (!byCaregiver[caregiverName].shifts[shiftType]) {
      byCaregiver[caregiverName].shifts[shiftType] = { shiftType: shiftType, count: 0, rate: rate, subtotal: 0 };
    }
    byCaregiver[caregiverName].shifts[shiftType].count += 1;
    byCaregiver[caregiverName].shifts[shiftType].subtotal += rate;
    byCaregiver[caregiverName].total += rate;
  });

  let grandTotal = 0;
  const rows = order.map(function (name) {
    const c = byCaregiver[name];
    grandTotal += c.total;
    return {
      caregiverName: c.caregiverName,
      shifts: Object.keys(c.shifts).map(function (t) { return c.shifts[t]; }),
      total: c.total
    };
  });
  rows.sort(function (a, b) { return b.total - a.total; });

  return { success: true, rows: rows, grandTotal: grandTotal };
}

// ชีทบันทึกของคนไข้แต่ละคนขึ้นต้นด้วย LOG_PREFIX_ กันชื่อชนกับชีทระบบ
const LOG_PREFIX_ = 'บันทึก_';
function patientSheet_(patientName) {
  const safe = String(patientName || '').replace(/[\[\]\*\?\/\\:]/g, ' ').trim().slice(0, 90);
  return SS.getSheetByName(LOG_PREFIX_ + safe);
}

function getOrCreatePatientSheet_(patientName) {
  let sheet = patientSheet_(patientName);
  if (!sheet) {
    const safe = String(patientName || '').replace(/[\[\]\*\?\/\\:]/g, ' ').trim().slice(0, 90);
    sheet = SS.insertSheet(LOG_PREFIX_ + safe);
    sheet.appendRow(['วันที่', 'เวลา', 'หมวด', 'รายการ', 'รายละเอียด/ค่า', 'บันทึกโดย', 'เวลาบันทึกจริง']);
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, 7);
  }
  return sheet;
}

/**
 * payload = {
 *   patientName: string,
 *   date: 'yyyy-MM-dd',
 *   time: 'HH:mm',
 *   caregiverName: string,
 *   items: [{ category: string, item: string, detail: string }]
 * }
 *
 * ถ้าวันที่+เวลา+หมวด+รายการ ตรงกับที่เคยบันทึกไว้แล้ว (เช่น ผู้ดูแลกลับมาแก้ไขรายการรอบเดิม)
 * จะ "แก้ไขทับ" แถวเดิมแทนการเพิ่มแถวใหม่ซ้ำ — รายการที่ไม่เคยมีจะถูกเพิ่มใหม่ตามปกติ
 */
function submitRecord(payload) {
  if (!payload || !payload.patientName || !payload.items || payload.items.length === 0) {
    return { success: false, message: 'ไม่มีรายการที่ติ๊กเลือกไว้' };
  }

  // ตรวจสิทธิ์: ถ้าผู้ดูแลถูกผูกไว้กับคนไข้เฉพาะราย ห้ามบันทึกให้คนไข้นอกรายชื่อ
  // (กันไว้สองชั้น — ฝั่งเว็บกรองไม่ให้เลือกอยู่แล้ว แต่ฝั่งนี้คือด่านจริงที่ยืนยันสิทธิ์)
  if (payload.caregiverCode) {
    const cg = verifyCaregiverCode(payload.caregiverCode);
    if (!cg.success) {
      return { success: false, message: 'รหัสผู้ดูแลไม่ถูกต้อง' };
    }
    if (cg.assignedPatients && cg.assignedPatients.indexOf(String(payload.patientName).trim()) === -1) {
      return { success: false, message: 'คุณไม่มีสิทธิ์บันทึกข้อมูลของคนไข้ท่านนี้' };
    }
  }

  const sheet = getOrCreatePatientSheet_(payload.patientName);
  const now = new Date();

  const lastRow = sheet.getLastRow();
  const existing = lastRow >= 2 ? sheet.getRange(2, 1, lastRow - 1, 7).getValues() : [];

  let added = 0;
  let updated = 0;
  const newRows = [];

  payload.items.forEach(function (it) {
    let matchIdx = -1;
    for (let i = 0; i < existing.length; i++) {
      const r = existing[i];
      if (formatDate_(r[0]) === payload.date &&
          formatTimeValue_(r[1]) === payload.time &&
          String(r[2] || '') === it.category &&
          String(r[3] || '') === it.item) {
        matchIdx = i;
        break;
      }
    }
    if (matchIdx !== -1) {
      const sheetRow = matchIdx + 2; // +1 เพราะแถวหัวตาราง, +1 เพราะ index เริ่มที่ 0
      sheet.getRange(sheetRow, 5, 1, 3).setValues([[it.detail || '', payload.caregiverName, now]]);
      updated++;
    } else {
      newRows.push([payload.date, payload.time, it.category, it.item, it.detail || '', payload.caregiverName, now]);
      added++;
    }
  });

  if (newRows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, 7).setValues(newRows);
  }

  const abnormal = checkVitalsAbnormal_(payload.items);
  if (abnormal.length > 0) {
    logVitalAlert_(payload.patientName, payload.caregiverName, abnormal);
  }

  return { success: true, count: added + updated, added: added, updated: updated };
}

/**
 * ดึงรายการที่เคยบันทึกไว้แล้วของคนไข้ 1 คน ตรงวันที่+เวลาที่ระบุเป๊ะๆ
 * ใช้ตอนผู้ดูแลกด "โหลดข้อมูลที่เคยบันทึกไว้" เพื่อดึงมาแก้ไขก่อนบันทึกทับ
 */
function getExistingRecordsForTime(patientName, dateStr, timeStr) {
  if (!patientName) return { success: false, message: 'กรุณาระบุชื่อคนไข้' };
  const sheet = patientSheet_(patientName);
  if (!sheet) return { success: true, items: [] };

  const data = sheet.getDataRange().getValues();
  const items = data.slice(1)
    .filter(function (r) {
      return formatDate_(r[0]) === dateStr && formatTimeValue_(r[1]) === timeStr;
    })
    .map(function (r) {
      return {
        category: String(r[2] || ''),
        item: String(r[3] || ''),
        detail: String(r[4] || ''),
        caregiver: String(r[5] || '')
      };
    });

  return { success: true, items: items };
}

// สรุปว่าวันนั้นบันทึกไว้กี่เวลา แต่ละเวลามีกี่รายการ — ให้ผู้ดูแลเลือกย้อนกลับไปแก้ไขได้ง่ายๆ โดยไม่ต้องจำเวลาที่พิมพ์ไว้
function getTimesSummaryForDate_(patientName, dateStr) {
  if (!patientName) return { success: false, message: 'กรุณาระบุชื่อคนไข้' };
  const grouped = getGroupedRecords_(patientName, dateStr);
  if (!grouped) return { success: false, message: 'ไม่พบข้อมูลคนไข้นี้' };

  const times = grouped.times.map(function (t) {
    return { time: t, count: grouped.byTime[t].length };
  });

  return { success: true, date: dateStr, dateThai: formatThaiDate_(dateStr), times: times };
}

// ---------------------------------------------------------------------------
// Daily report -> PDF (บันทึก PDF ไว้ใน Google Drive ให้แอดมินกดดาวน์โหลด/ส่งเข้า LINE เอง)
// ดีไซน์: หัวกระดาษโลโก้ร้าน + กล่องข้อมูลคนไข้ + การ์ดแยกตามช่วงเวลา สีธีมทีล
// ---------------------------------------------------------------------------

// --- ธีมสี / ฟอนต์ ---
const TEAL = '#14a89c';
const TEAL_DARK = '#0b7a70';
const DARK = '#2c2c2c';
const GRAY = '#8a8a8a';
const LIGHT_BG = '#eef7f6';
const ROW_ALT = '#f6fbfb';
const BORDER_LIGHT = '#dcebe9';
const FONT_BODY = 'Sarabun';
const FONT_HEAD = 'Prompt';

// อ่าน+จัดกลุ่มบันทึกของคนไข้ 1 คนในวันที่กำหนด ใช้ร่วมกันทั้งฝั่งสร้าง PDF และฝั่งแสดงผลหน้าเว็บของแอดมิน
// คืนค่า null ถ้าไม่พบชีทคนไข้นี้เลย, หรือ { rows, byTime, times } (rows อาจว่างได้ถ้าวันนั้นไม่มีบันทึก)
function getGroupedRecords_(patientName, dateStr) {
  const sheet = patientSheet_(patientName);
  if (!sheet) return null;

  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1).filter(function (r) {
    return formatDate_(r[0]) === dateStr;
  });

  // จัดกลุ่มตามเวลา (ใช้ formatTimeValue_ กันบั๊กเวลาที่เก็บเป็น Date แล้วโชว์ปี ค.ศ. 1899)
  const byTime = {};
  rows.forEach(function (r) {
    const t = formatTimeValue_(r[1]);
    if (!byTime[t]) byTime[t] = [];
    byTime[t].push(r);
  });
  const times = Object.keys(byTime).sort();

  return { rows: rows, byTime: byTime, times: times };
}

// ---------------------------------------------------------------------------
// เหมือน getGroupedRecords_ แต่รับช่วงวันที่ (dateFrom..dateTo) แทนวันเดียว ใช้กับหน้าดูรายงานของแอดมิน
// คืนค่า null ถ้าไม่พบชีทคนไข้นี้เลย, หรือ { rows, byDate, dates } (rows อาจว่างได้ถ้าช่วงนั้นไม่มีบันทึก)
// ---------------------------------------------------------------------------
function getGroupedRecordsRange_(patientName, dateFrom, dateTo) {
  const sheet = patientSheet_(patientName);
  if (!sheet) return null;

  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1).filter(function (r) {
    const d = formatDate_(r[0]);
    return d >= dateFrom && d <= dateTo;
  });

  const byDate = {};
  rows.forEach(function (r) {
    const d = formatDate_(r[0]);
    const t = formatTimeValue_(r[1]);
    if (!byDate[d]) byDate[d] = {};
    if (!byDate[d][t]) byDate[d][t] = [];
    byDate[d][t].push(r);
  });
  const dates = Object.keys(byDate).sort();

  return { rows: rows, byDate: byDate, dates: dates };
}

// ---------------------------------------------------------------------------
// สำหรับแอดมิน: ดึงบันทึกของคนไข้ 1 คนในช่วงวันที่ที่เลือก (วันเดียวหรือหลายวันก็ได้) เป็น JSON
// ให้หน้าเว็บแสดงผล/พิมพ์เอง (ไม่ต้องสร้างไฟล์ PDF) — จัดกลุ่มเป็นรายวัน แต่ละวันแยกตามช่วงเวลาอีกที
// ---------------------------------------------------------------------------
function getDailyRecordsForView(patientName, dateFrom, dateTo) {
  if (!patientName) return { success: false, message: 'กรุณาระบุชื่อคนไข้' };
  const grouped = getGroupedRecordsRange_(patientName, dateFrom, dateTo);
  if (!grouped) return { success: false, message: 'ไม่พบข้อมูลคนไข้นี้' };

  const days = grouped.dates.map(function (d) {
    const rowsForDay = grouped.rows.filter(function (r) { return formatDate_(r[0]) === d; });
    return {
      date: d,
      dateThai: formatThaiDate_(d),
      itemGroups: buildItemGroupsForDay_(rowsForDay)
    };
  });

  const meta = getPatientMeta_(patientName);
  const caregiverNames = Array.from(new Set(grouped.rows.map(function (r) { return String(r[5] || ''); }).filter(Boolean)));
  const docSuffix = dateFrom === dateTo ? dateFrom : (dateFrom + '_' + dateTo);

  return {
    success: true,
    patientName: patientName,
    zone: meta.zone,
    dateFrom: dateFrom,
    dateTo: dateTo,
    dateFromThai: formatThaiDate_(dateFrom),
    dateToThai: formatThaiDate_(dateTo),
    issueDateThai: formatThaiDate_(formatDate_(new Date())),
    docNumber: buildDocNumber_(meta.id, docSuffix),
    caregiverNames: caregiverNames,
    days: days,
    message: days.length === 0 ? 'ไม่มีบันทึกของช่วงวันที่นี้' : ''
  };
}

// ---------------------------------------------------------------------------
// สำหรับแอดมิน: ประวัติการบันทึกล่าสุดของผู้ดูแลทุกคน ทุกคนไข้ (ไล่อ่านทุกชีทคนไข้)
// ---------------------------------------------------------------------------
function getActivityLog(limit) {
  const sheets = SS.getSheets();
  let all = [];

  sheets.forEach(function (sheet) {
    const sheetName = sheet.getName();
    if (sheetName.indexOf(LOG_PREFIX_) !== 0) return;
    const name = sheetName.slice(LOG_PREFIX_.length);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;
    const values = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
    values.forEach(function (r) {
      if (!r[0]) return;
      const loggedAtDate = r[6] instanceof Date ? r[6] : null;
      all.push({
        patientName: name,
        date: formatDate_(r[0]),
        time: formatTimeValue_(r[1]),
        category: String(r[2] || ''),
        item: String(r[3] || ''),
        detail: String(r[4] || ''),
        caregiverName: String(r[5] || ''),
        loggedAt: loggedAtDate ? Utilities.formatDate(loggedAtDate, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') : String(r[6] || ''),
        loggedAtSort: loggedAtDate ? loggedAtDate.getTime() : 0
      });
    });
  });

  all.sort(function (a, b) { return b.loggedAtSort - a.loggedAtSort; });
  if (limit) all = all.slice(0, limit);
  all.forEach(function (r) { delete r.loggedAtSort; });

  return { success: true, records: all };
}

function generateDailyReport(patientName, dateStr) {
  const grouped = getGroupedRecords_(patientName, dateStr);
  if (!grouped) return { success: false, message: 'ไม่พบข้อมูลคนไข้นี้' };
  if (grouped.rows.length === 0) {
    return { success: false, message: 'ไม่มีบันทึกของวันที่นี้' };
  }
  const meta = getPatientMeta_(patientName);
  const docNumber = buildDocNumber_(meta.id, dateStr);
  const caregiverNames = Array.from(new Set(grouped.rows.map(function (r) { return String(r[5] || ''); }).filter(Boolean)));
  const itemGroups = buildItemGroupsForDay_(grouped.rows);

  const doc = DocumentApp.create('รายงานประจำวัน_' + patientName + '_' + dateStr);
  const body = doc.getBody();
  body.setMarginTop(30).setMarginBottom(30).setMarginLeft(40).setMarginRight(40);

  // ---------- หัวกระดาษ: โลโก้+ชื่อศูนย์ (ซ้าย) / เลขที่เอกสาร+วันที่ออกเอกสาร (ขวา) ----------
  const headerTable = body.appendTable([['', '']]);
  headerTable.setBorderWidth(0);
  const headerRow = headerTable.getRow(0);

  const brandCell = headerRow.getCell(0);
  brandCell.setWidth(330);
  brandCell.setVerticalAlignment(DocumentApp.VerticalAlignment.TOP);
  const brandRow = brandCell.appendTable([['', '']]);
  brandRow.setBorderWidth(0);
  const logoCell = brandRow.getRow(0).getCell(0);
  logoCell.setWidth(46);
  logoCell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
  logoCell.appendImage(getLogoBlob_())
    .setWidth(38).setHeight(38);
  const brandTextCell = brandRow.getRow(0).getCell(1);
  brandTextCell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
  const nameP = brandTextCell.getChild(0).asParagraph();
  nameP.setFontFamily(FONT_HEAD).setSpacingAfter(0).setSpacingBefore(0);
  nameP.editAsText().setText(getCompanyName_());
  nameP.editAsText().setBold(true).setFontSize(14).setForegroundColor(TEAL_DARK);
  const subP = brandTextCell.appendParagraph('รายงานการดูแลผู้ป่วยประจำวัน');
  subP.setFontFamily(FONT_BODY).setSpacingAfter(0).setSpacingBefore(0);
  subP.editAsText().setFontSize(9).setForegroundColor(GRAY);
  brandCell.getChild(0).asParagraph().setSpacingAfter(0);

  const docInfoCell = headerRow.getCell(1);
  docInfoCell.setVerticalAlignment(DocumentApp.VerticalAlignment.TOP);
  const docNumP = docInfoCell.getChild(0).asParagraph();
  docNumP.setAlignment(DocumentApp.HorizontalAlignment.RIGHT).setSpacingAfter(2).setSpacingBefore(0);
  const docNumT = docNumP.editAsText();
  let dOff = 0;
  dOff = styleAppend_(docNumT, dOff, 'เลขที่เอกสาร ', { color: GRAY, size: 9, font: FONT_BODY });
  dOff = styleAppend_(docNumT, dOff, docNumber, { bold: true, color: DARK, size: 9.5, font: FONT_BODY });
  const issueP = docInfoCell.appendParagraph('');
  issueP.setAlignment(DocumentApp.HorizontalAlignment.RIGHT).setSpacingAfter(0).setSpacingBefore(0);
  const issueT = issueP.editAsText();
  let iOff = 0;
  iOff = styleAppend_(issueT, iOff, 'วันที่ออกเอกสาร ', { color: GRAY, size: 9, font: FONT_BODY });
  iOff = styleAppend_(issueT, iOff, formatThaiDate_(formatDate_(new Date())), { bold: true, color: DARK, size: 9.5, font: FONT_BODY });

  // เส้นคาดสีทีลใต้หัวกระดาษ
  appendDivider_(body, TEAL, 3);
  body.appendParagraph('').setSpacingBefore(0).setSpacingAfter(8);

  // ---------- สองคอลัมน์: ผู้รับบริการ (ซ้าย) / สรุปข้อมูลการบันทึก (ขวา) ----------
  const twoColTable = body.appendTable([['', '']]);
  twoColTable.setBorderWidth(0);
  const twoColRow = twoColTable.getRow(0);

  const patientCell = twoColRow.getCell(0);
  const pLabel = patientCell.getChild(0).asParagraph();
  pLabel.setFontFamily(FONT_BODY).setSpacingAfter(3).setSpacingBefore(0);
  pLabel.editAsText().setText('ผู้รับบริการ');
  pLabel.editAsText().setBold(true).setFontSize(9.5).setForegroundColor(TEAL_DARK);
  const pName = patientCell.appendParagraph(patientName);
  pName.setFontFamily(FONT_HEAD).setSpacingAfter(0).setSpacingBefore(0);
  pName.editAsText().setBold(true).setFontSize(13).setForegroundColor(DARK);
  if (meta.zone) {
    const pZone = patientCell.appendParagraph('โซน ' + meta.zone);
    pZone.setFontFamily(FONT_BODY).setSpacingAfter(0).setSpacingBefore(0);
    pZone.editAsText().setFontSize(9).setForegroundColor(GRAY);
  }

  const summaryCell = twoColRow.getCell(1);
  const sLabel = summaryCell.getChild(0).asParagraph();
  sLabel.setFontFamily(FONT_BODY).setSpacingAfter(3).setSpacingBefore(0);
  sLabel.editAsText().setText('สรุปข้อมูลการบันทึกวันนี้');
  sLabel.editAsText().setBold(true).setFontSize(9.5).setForegroundColor(TEAL_DARK);
  const sLine1 = summaryCell.appendParagraph('จำนวนหัวข้อที่บันทึก: ' + itemGroups.length + ' หัวข้อ');
  sLine1.setFontFamily(FONT_BODY).setSpacingAfter(0).setSpacingBefore(0);
  sLine1.editAsText().setFontSize(9.5).setForegroundColor(DARK);
  const sLine2 = summaryCell.appendParagraph('จำนวนรายการทั้งหมด: ' + grouped.rows.length + ' รายการ');
  sLine2.setFontFamily(FONT_BODY).setSpacingAfter(0).setSpacingBefore(0);
  sLine2.editAsText().setFontSize(9.5).setForegroundColor(DARK);

  body.appendParagraph('').setSpacingAfter(6);

  // ---------- กล่องข้อมูลแบบมีเส้นขอบ (label ซ้าย / value ขวา) ----------
  const detailRows = [
    ['วันที่บันทึก', formatThaiDate_(dateStr)],
    ['จำนวนผู้ดูแลที่บันทึกวันนี้', caregiverNames.length + ' คน'],
    ['ชื่อผู้ดูแล', caregiverNames.join(', ') || '-']
  ];
  const detailTable = body.appendTable(detailRows);
  detailTable.setBorderColor(BORDER_LIGHT).setBorderWidth(0.75);
  for (let i = 0; i < detailRows.length; i++) {
    const row = detailTable.getRow(i);
    const labelCell = row.getCell(0);
    const valueCell = row.getCell(1);
    labelCell.setWidth(150).setPaddingTop(6).setPaddingBottom(6).setPaddingLeft(10);
    valueCell.setPaddingTop(6).setPaddingBottom(6).setPaddingRight(10);

    const labelP = labelCell.getChild(0).asParagraph();
    labelP.setFontFamily(FONT_BODY);
    labelP.editAsText().setForegroundColor(GRAY).setFontSize(9.5);

    const valueP = valueCell.getChild(0).asParagraph();
    valueP.setFontFamily(FONT_BODY).setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    valueP.editAsText().setBold(true).setForegroundColor(DARK).setFontSize(9.5);
  }
  body.appendParagraph('').setSpacingAfter(10);

  // ---------- หัวข้อรายการ ----------
  const listTitleP = body.appendParagraph('รายการที่บันทึกวันนี้');
  listTitleP.setFontFamily(FONT_HEAD).setSpacingBefore(0).setSpacingAfter(8);
  listTitleP.editAsText().setBold(true).setFontSize(13).setForegroundColor(TEAL_DARK);

  // ---------- การ์ดแยกตามหัวข้อรายการ (แต่ละหัวข้อโชว์ทุกครั้งที่บันทึกในวันนั้น เรียงตามเวลา) ----------
  itemGroups.forEach(function (grp) {
    // แถบหัวข้อ (พื้นสีทีล ตัวอักษรขาว) — ไอคอนหมวด + ชื่อรายการ
    const headerBar = body.appendTable([[categoryIcon_(grp.category) + '  ' + grp.item]]);
    headerBar.setBorderWidth(0);
    const headerCell = headerBar.getRow(0).getCell(0);
    headerCell.setBackgroundColor(TEAL);
    headerCell.setPaddingTop(6).setPaddingBottom(6).setPaddingLeft(10);
    const headerP = headerCell.getChild(0).asParagraph();
    headerP.setFontFamily(FONT_HEAD);
    headerP.editAsText().setBold(true).setForegroundColor('#ffffff').setFontSize(12);

    // ตารางของทุกครั้งที่บันทึกหัวข้อนี้ในวันนั้น (ลายเส้นบางสีทีลอ่อน + แถบสีสลับแถว)
    let occTable = null;
    grp.occurrences.forEach(function (occ, idx) {
      let cell;
      if (idx === 0) {
        occTable = body.appendTable([['']]);
        cell = occTable.getRow(0).getCell(0);
      } else {
        cell = occTable.appendTableRow().appendTableCell();
      }
      fillOccurrenceCell_(cell, occ, idx);
    });
    occTable.setBorderColor(BORDER_LIGHT).setBorderWidth(0.6);

    body.appendParagraph('').setSpacingAfter(8);
  });

  // ---------- ท้ายรายงาน ----------
  appendDivider_(body, BORDER_LIGHT, 1);
  body.appendParagraph('').setSpacingBefore(0).setSpacingAfter(4);
  const thankP = body.appendParagraph('ขอบคุณที่ไว้วางใจให้ ' + getCompanyName_() + ' ดูแลคนที่คุณรัก');
  thankP.setFontFamily(FONT_BODY).setAlignment(DocumentApp.HorizontalAlignment.CENTER).setSpacingAfter(4).setSpacingBefore(0);
  thankP.editAsText().setForegroundColor(TEAL_DARK).setFontSize(10).setItalic(true);
  const footNote = body.appendParagraph(
    'สร้างรายงานโดยระบบ ' + PRODUCT_NAME + ' เมื่อ ' +
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'd MMM yyyy HH:mm')
  );
  footNote.setFontFamily(FONT_BODY).setAlignment(DocumentApp.HorizontalAlignment.CENTER).setSpacingBefore(0);
  footNote.editAsText().setFontSize(8).setForegroundColor(GRAY);

  doc.saveAndClose();

  const docFile = DriveApp.getFileById(doc.getId());
  const pdfBlob = docFile.getAs('application/pdf');
  const pdfFile = driveFolder_('reports').createFile(pdfBlob).setName('รายงาน_' + patientName + '_' + dateStr + '.pdf');
  docFile.setTrashed(true); // เก็บเฉพาะ PDF ไม่เก็บ Doc ต้นฉบับ

  // แชร์ลิงก์เพื่อส่งต่อให้ญาติทาง LINE ได้ — ปิดได้ที่ ตั้งค่า > sharePublic
  if (getSettings_().sharePublic === '1') pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return {
    success: true,
    viewUrl: pdfFile.getUrl(),
    downloadUrl: 'https://drive.google.com/uc?export=download&id=' + pdfFile.getId()
  };
}

// อ่านข้อมูลเสริมของคนไข้ (รหัส/โซน) จากชีท Patients เพื่อใช้ในหัวรายงาน
function getPatientMeta_(patientName) {
  const sheet = ensurePatientsSheet_();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim() === String(patientName).trim()) {
      return { id: String(data[i][0] || ''), zone: String(data[i][2] || '') };
    }
  }
  return { id: '', zone: '' };
}

// สร้างเลขที่เอกสารของรายงาน เช่น RPT-20260918-P001
function buildDocNumber_(patientId, dateStr) {
  const datePart = String(dateStr).replace(/-/g, '');
  return 'RPT-' + datePart + (patientId ? '-' + patientId : '');
}

// เติมเนื้อหา 1 แถว = 1 ครั้งที่บันทึกหัวข้อนี้ (เวลา + ค่าที่บันทึก + ใครบันทึก — แยกบรรทัดย่อยถ้ามีมากกว่า 1 ค่า)
function fillOccurrenceCell_(cell, occ, idx) {
  const parts = splitDetailParts_(occ.detail);

  cell.setBackgroundColor(idx % 2 === 0 ? '#ffffff' : ROW_ALT);
  cell.setPaddingTop(7).setPaddingBottom(7).setPaddingLeft(12).setPaddingRight(10);

  const p1 = cell.getChild(0).asParagraph();
  p1.setFontFamily(FONT_BODY).setSpacingBefore(0).setSpacingAfter(parts.length > 1 ? 2 : 1);
  const t1 = p1.editAsText();
  let off = 0;
  off = styleAppend_(t1, off, 'เวลา ' + occ.time + ' น.', { bold: true, color: TEAL_DARK, size: 11 });
  if (parts.length === 1) {
    off = styleAppend_(t1, off, '   —   ' + parts[0], { color: DARK, size: 11 });
  } else if (parts.length === 0 && occ.detail) {
    off = styleAppend_(t1, off, '   —   ' + occ.detail, { color: DARK, size: 11 });
  }

  if (parts.length > 1) {
    parts.forEach(function (part) {
      const sub = cell.appendParagraph('•  ' + part);
      sub.setFontFamily(FONT_BODY).setSpacingBefore(0).setSpacingAfter(1).setIndentStart(16);
      sub.editAsText().setForegroundColor(DARK).setFontSize(10.5);
    });
  }

  const p2 = cell.appendParagraph('บันทึกโดย ' + occ.caregiver);
  p2.setFontFamily(FONT_BODY).setSpacingBefore(0).setSpacingAfter(0);
  p2.editAsText().setItalic(true).setForegroundColor(GRAY).setFontSize(9);
}

// ต่อข้อความสไตล์ต่างๆ เข้าไปใน Text object ทีละท่อน แล้วคืน offset ถัดไป
function styleAppend_(textElement, offset, str, opts) {
  if (!str) return offset;
  textElement.appendText(str);
  const start = offset;
  const end = offset + str.length - 1;
  if (opts.bold !== undefined) textElement.setBold(start, end, opts.bold);
  if (opts.italic) textElement.setItalic(start, end, true);
  if (opts.color) textElement.setForegroundColor(start, end, opts.color);
  if (opts.size) textElement.setFontSize(start, end, opts.size);
  if (opts.font) textElement.setFontFamily(start, end, opts.font);
  return offset + str.length;
}

// เส้นคาดสีบางๆ คั่นแต่ละส่วน (จำลองด้วยตาราง 1 ช่องพื้นสี)
function appendDivider_(body, color, heightPt) {
  const bar = body.appendTable([['']]);
  bar.setBorderWidth(0);
  const cell = bar.getRow(0).getCell(0);
  cell.setBackgroundColor(color);
  cell.setPaddingTop(heightPt / 2).setPaddingBottom(heightPt / 2).setPaddingLeft(0).setPaddingRight(0);
  cell.getChild(0).asParagraph().editAsText().setFontSize(1);
}

// เลือกไอคอนตามคำในชื่อหมวด
function categoryIcon_(category) {
  const map = [
    ['โภชนาการ', '🍽️'],
    ['ขับถ่าย', '🚻'],
    ['ยา', '💊'],
    ['สุขอนามัย', '🧼'],
    ['เคลื่อนไหว', '🚶'],
    ['แผลกดทับ', '🛡️'],
    ['สัญญาณชีพ', '❤️'],
    ['หัตถการ', '🩺'],
    ['อุปกรณ์', '🩺'],
    ['เหตุการณ์', '⚠️']
  ];
  for (let i = 0; i < map.length; i++) {
    if (category.indexOf(map[i][0]) !== -1) return map[i][1];
  }
  return '📋';
}

// แยกค่ารายละเอียดที่อาจรวมหลายช่องไว้ (คั่นด้วย '  —  ' ตามฝั่งเว็บ) กลับเป็นแต่ละท่อน
function splitDetailParts_(detail) {
  if (!detail) return [];
  return String(detail).split('  —  ').map(function (s) { return s.trim(); }).filter(Boolean);
}

// จัดกลุ่มบันทึกของวันเดียว "ตามหัวข้อรายการ" แทนที่จะเรียงตามเวลา — แต่ละหัวข้อ (เช่น "ทานอาหารแล้ว")
// จะมีลิสต์ของทุกครั้งที่บันทึกในวันนั้นอยู่ข้างใน เรียงตามเวลา ส่วนหัวข้อเรียงตามเวลาที่บันทึกครั้งแรกของวัน
function buildItemGroupsForDay_(rows) {
  const groupMap = {};
  const order = [];
  rows.forEach(function (r) {
    const category = String(r[2] || '');
    const item = String(r[3] || '');
    const key = category + '|||' + item;
    if (!groupMap[key]) {
      groupMap[key] = { category: category, item: item, occurrences: [] };
      order.push(key);
    }
    groupMap[key].occurrences.push({
      time: formatTimeValue_(r[1]),
      detail: String(r[4] || ''),
      caregiver: String(r[5] || '')
    });
  });
  order.forEach(function (key) {
    groupMap[key].occurrences.sort(function (a, b) { return a.time < b.time ? -1 : (a.time > b.time ? 1 : 0); });
  });
  order.sort(function (a, b) {
    const ta = groupMap[a].occurrences[0].time;
    const tb = groupMap[b].occurrences[0].time;
    return ta < tb ? -1 : (ta > tb ? 1 : 0);
  });
  return order.map(function (key) { return groupMap[key]; });
}

// แปลงค่าเวลาที่อาจเก็บเป็น Date (บั๊กเดิม: โชว์ "Sat Dec 30 1899") ให้เป็น HH:mm เสมอ
function formatTimeValue_(t) {
  if (Object.prototype.toString.call(t) === '[object Date]') {
    return Utilities.formatDate(t, Session.getScriptTimeZone(), 'HH:mm');
  }
  const s = String(t);
  const m = s.match(/^(\d{1,2}):(\d{2})/);
  return m ? (('0' + m[1]).slice(-2) + ':' + m[2]) : s;
}

// แปลง yyyy-MM-dd เป็นรูปแบบไทยอ่านง่าย เช่น "18 กันยายน 2569"
const THAI_MONTHS_ = ['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน',
  'กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];
function formatThaiDate_(dateStr) {
  const parts = String(dateStr).split('-');
  if (parts.length !== 3) return dateStr;
  const y = parseInt(parts[0], 10) + 543;
  const m = parseInt(parts[1], 10) - 1;
  const d = parseInt(parts[2], 10);
  if (isNaN(y) || isNaN(m) || isNaN(d) || m < 0 || m > 11) return dateStr;
  return d + ' ' + THAI_MONTHS_[m] + ' ' + y;
}

function formatDate_(d) {
  if (Object.prototype.toString.call(d) === '[object Date]') {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(d);
}

