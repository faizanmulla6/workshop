/**
 * EK ROZAH TARBIYATI WORKSHOP — Google Sheets + Apps Script Backend
 * Version 5 (fixed)
 *
 * What changed vs v3/v4:
 *  - Fixed API routing (adminStats_/adminList_ did not exist).
 *  - Admin PIN no longer travels in URLs; all admin calls are POST / google.script.run.
 *  - PIN brute-force lock (5 wrong tries -> 15 min lock). PIN can live in Script Properties.
 *  - Half-finished registrations (page closed / refreshed before upload) can RESUME
 *    instead of getting "mobile already registered".
 *  - Registration close date + optional seat limit enforced on the server.
 *  - Screenshot upload is locked, validated (magic bytes) and safe against double submit.
 *  - Public lookup no longer leaks the Drive screenshot link / admin notes.
 *  - adminUpdate no longer wipes remarks when only status is changed.
 *  - Mobile column stored as text; timezone forced to Asia/Kolkata.
 *  - viewport meta tag added (Apps Script strips it otherwise -> tiny/zoomed mobile page).
 *
 * DEPLOY: after pasting this file you MUST create a NEW VERSION of the deployment
 * (Deploy -> Manage deployments -> pencil -> Version: New version -> Deploy).
 * Otherwise the old code keeps running on the same /exec URL.
 */

const CONFIG = {
  SPREADSHEET_ID: '18x8dAcWAGniFsrsV0YGAdVmBWT3aEIUEoPe7TNw-7cI',
  REGISTRATION_SHEET: 'Registrations',
  AUDIT_SHEET: 'Audit_Log',
  PAYMENT_FOLDER_NAME: 'Ek_Rozah_Tarbiyati_Workshop_Payments',

  // Better: Project Settings -> Script Properties -> add ADMIN_PIN. It overrides this value.
  ADMIN_PIN: '3249',

  PROGRAM_NAME: 'Tahaffuz-e-Aqeedah Wa Manhaj',
  PROGRAM_SHORT: 'Ek Rozah Tarbiyati Workshop',
  REGISTRATION_PREFIX: 'TRW',
  FEE: 50,
  UPI_ID: 'faizanmulla6@ybl',
  PAYEE_NAME: 'Ek Rozah Tarbiyati Workshop',

  VENUE: 'Masjid-e-Umar Khalifa',
  VENUE_ADDRESS: 'Sumaiyya Nagar, Nai Zindagi, Solapur',
  PROGRAM_DATE: '18 October 2026, Sunday',
  REPORTING_TIME: '09:30 AM',
  PROGRAM_TIME: 'Subah 9:30 baje se Namaz-e-Isha tak',
  LAST_REGISTRATION_DATE: '16 October 2026',
  CONTACT_MARD: '70284 88819 / 99608 05017',
  ORGANISATION: 'Zilayi Jamiat Ahle Hadees, Solapur',

  TIMEZONE: 'Asia/Kolkata',
  REGISTRATION_CLOSES_ON: '2026-10-16',   // yyyy-MM-dd, inclusive. '' = never closes
  MAX_SEATS: 0,                           // 0 = unlimited. Cancelled rows do not count.
  MAX_PER_MOBILE: 1,                      // set 3-4 if families register on one mobile number
  MAX_SCREENSHOT_BYTES: 5 * 1024 * 1024
};

const HEADERS = [
  'Registration No','Timestamp','Name','Age','Gender','Location (City)',
  'Mobile Number','Language','Registration Fee','Payment Screenshot URL',
  'UTR / Transaction ID','Payment Status','Registration Status',
  'Venue','Program Date','Reporting Time','Program Time',
  'Last Registration Date','Remarks','Admin Note','Admin Updated By','Admin Updated At'
];

// Column indexes (0-based) — keeps the code readable
const COL = {
  REG:0, TS:1, NAME:2, AGE:3, GENDER:4, CITY:5, MOBILE:6, LANG:7, FEE:8, SHOT:9,
  UTR:10, PAY:11, STATUS:12, VENUE:13, DATE:14, REPORT:15, TIME:16, LAST:17,
  REMARKS:18, NOTE:19, BY:20, AT:21
};

/* ============================ ENTRY POINTS ============================ */

function doGet(e) {
  // Tiny public JSON endpoint (config only). Everything else is POST.
  if (e && e.parameter && e.parameter.api === '1') {
    return jsonResponse_(publicConfig());
  }
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle(CONFIG.PROGRAM_NAME + ' — Registration')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * External frontend (Netlify/Vercel) calls:  POST  {"fn":"submitRegistration","args":[ ... ]}
 * (Content-Type text/plain to avoid a CORS pre-flight.)
 */
function doPost(e) {
  try {
    const body = parsePostBody_(e);
    const api = {
      config: publicConfig,
      submitRegistration: submitRegistration,
      uploadPaymentScreenshot: uploadPaymentScreenshot,
      getRegistration: getRegistration,
      adminDashboard: adminDashboard,
      adminUpdate: adminUpdate
    };
    const fn = api[body.fn];
    if (!fn) return jsonResponse_({ok:false, message:'Unknown action.'});
    return jsonResponse_(fn.apply(null, body.args || []));
  } catch (err) {
    return jsonResponse_({ok:false, message: msg_(err)});
  }
}

/* ============================== SETUP ================================ */

function setupSheets() {
  const sh = getRegSheet_();
  sh.setFrozenRows(1);
  sh.getRange(2, COL.MOBILE + 1, Math.max(sh.getMaxRows() - 1, 1), 1).setNumberFormat('@');

  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let audit = ss.getSheetByName(CONFIG.AUDIT_SHEET);
  if (!audit) audit = ss.insertSheet(CONFIG.AUDIT_SHEET);
  if (audit.getLastRow() === 0) {
    audit.getRange(1, 1, 1, 6).setValues([['Timestamp','Action','Registration No','Details','User','Status']]);
    audit.setFrozenRows(1);
  }
  getPaymentFolder_();
  return 'Sheets and payment folder are ready.';
}

/* ========================== PUBLIC API (safe) ======================== */
// Every public function returns {ok:true,...} or {ok:false,message}. Never throws to the client.

function publicConfig() {
  return {
    ok:true,
    programName:CONFIG.PROGRAM_NAME, programShort:CONFIG.PROGRAM_SHORT,
    fee:CONFIG.FEE, upiId:CONFIG.UPI_ID, payeeName:CONFIG.PAYEE_NAME,
    venue:CONFIG.VENUE, venueAddress:CONFIG.VENUE_ADDRESS,
    programDate:CONFIG.PROGRAM_DATE, reportingTime:CONFIG.REPORTING_TIME,
    programTime:CONFIG.PROGRAM_TIME, lastRegistrationDate:CONFIG.LAST_REGISTRATION_DATE,
    contact:CONFIG.CONTACT_MARD, organisation:CONFIG.ORGANISATION
  };
}

function submitRegistration(data)                       { return safe_(() => submitRegistration_(data)); }
function uploadPaymentScreenshot(regNo, mobile, file)   { return safe_(() => uploadPaymentScreenshot_(regNo, mobile, file)); }
function getRegistration(regNo, mobile)                 { return safe_(() => getRegistration_(regNo, mobile)); }
function adminDashboard(pin)                            { return safe_(() => adminDashboard_(pin)); }
function adminUpdate(pin, regNo, remarks, status)       { return safe_(() => adminUpdate_(pin, regNo, remarks, status)); }

/* ============================ REGISTRATION =========================== */

function submitRegistration_(data) {
  const v = validateRegistration_(data);

  if (isClosed_()) {
    return {ok:false, code:'CLOSED',
      message:'Registration is closed (last date ' + CONFIG.LAST_REGISTRATION_DATE + ').'};
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = getRegSheet_();
    const values = sh.getDataRange().getValues();
    const nameKey = normName_(v.name);

    let activeSeats = 0, sameMobile = 0;
    for (let i = 1; i < values.length; i++) {
      const r = values[i];
      const status = String(r[COL.STATUS] || '').toUpperCase();
      if (status === 'CANCELLED') continue;
      activeSeats++;

      if (normalizeMobile_(r[COL.MOBILE]) === v.mobile) {
        if (normName_(r[COL.NAME]) === nameKey) {
          // Same person registering again
          if (status === 'CONFIRMED') {
            return {ok:false, code:'ALREADY_REGISTERED', registrationNo:r[COL.REG],
              message:'You are already registered. Your Registration No. is ' + r[COL.REG] +
                      '. Use "My Registration" to view it.'};
          }
          // Not paid yet -> RESUME the same registration
          audit_('REGISTRATION_RESUMED', r[COL.REG], v.mobile);
          return {ok:true, resumed:true, registrationNo:r[COL.REG], fee:CONFIG.FEE, upiId:CONFIG.UPI_ID};
        }
        sameMobile++;
      }
    }

    if (sameMobile >= CONFIG.MAX_PER_MOBILE) {
      return {ok:false, code:'DUPLICATE',
        message:'This mobile number has already been used for registration.'};
    }
    if (CONFIG.MAX_SEATS > 0 && activeSeats >= CONFIG.MAX_SEATS) {
      return {ok:false, code:'FULL', message:'Sorry, all seats are filled.'};
    }

    const regNo = nextRegistrationNumber_(sh);
    const rowValues = [
      regNo, new Date(), v.name, v.age, v.gender, v.city, v.mobile, v.language,
      CONFIG.FEE, '', '', 'AWAITING_SCREENSHOT', 'INCOMPLETE',
      CONFIG.VENUE, CONFIG.PROGRAM_DATE, CONFIG.REPORTING_TIME, CONFIG.PROGRAM_TIME,
      CONFIG.LAST_REGISTRATION_DATE, '', '', '', ''
    ];
    const newRow = sh.getLastRow() + 1;
    sh.getRange(newRow, COL.MOBILE + 1).setNumberFormat('@');   // keep mobile as text
    sh.getRange(newRow, 1, 1, rowValues.length).setValues([rowValues]);

    audit_('REGISTRATION_CREATED', regNo, JSON.stringify({name:v.name, mobile:v.mobile, language:v.language}));
    return {ok:true, registrationNo:regNo, fee:CONFIG.FEE, upiId:CONFIG.UPI_ID};
  } finally {
    lock.releaseLock();
  }
}

function uploadPaymentScreenshot_(registrationNo, mobile, fileData) {
  if (!registrationNo || !mobile || !fileData || !fileData.data) {
    throw new Error('Registration number, mobile number and screenshot are required.');
  }
  const contentType = String(fileData.type || 'image/jpeg').toLowerCase();
  if (!/^image\/(jpeg|jpg|png|webp)$/.test(contentType)) {
    throw new Error('Please upload a JPG, PNG or WEBP image.');
  }

  const base64 = String(fileData.data).split(',').pop();
  const bytes = Utilities.base64Decode(base64);
  if (bytes.length > CONFIG.MAX_SCREENSHOT_BYTES) {
    throw new Error('Screenshot is too large. Please upload an image below 5 MB.');
  }
  if (!looksLikeImage_(bytes)) throw new Error('This file is not a valid image.');

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sh = getRegSheet_();
    const values = sh.getDataRange().getValues();
    let row = -1;
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][COL.REG]).trim() === String(registrationNo).trim() &&
          normalizeMobile_(values[i][COL.MOBILE]) === normalizeMobile_(mobile)) { row = i + 1; break; }
    }
    if (row < 0) throw new Error('Registration not found.');
    if (String(values[row - 1][COL.STATUS]).toUpperCase() === 'CANCELLED') {
      throw new Error('This registration was cancelled. Please contact the organisers.');
    }

    const safeReg = String(registrationNo).replace(/[^A-Za-z0-9_-]/g, '_');
    const ext = contentType.split('/')[1].replace('jpeg', 'jpg');
    const filename = safeReg + '_payment_' +
      Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyyMMdd_HHmmss') + '.' + ext;

    const file = getPaymentFolder_().createFile(Utilities.newBlob(bytes, contentType, filename));
    file.setDescription('Payment screenshot for registration ' + registrationNo);

    sh.getRange(row, COL.SHOT + 1).setValue(file.getUrl());
    sh.getRange(row, COL.PAY + 1).setValue('PAYMENT_SCREENSHOT_RECEIVED');
    sh.getRange(row, COL.STATUS + 1).setValue('CONFIRMED');
    sh.getRange(row, COL.REMARKS + 1).setValue('Payment screenshot received.');
    sh.getRange(row, COL.AT + 1).setValue(new Date());

    audit_('PAYMENT_SCREENSHOT_RECEIVED', registrationNo, file.getUrl());
    return publicObject_(sh.getRange(row, 1, 1, HEADERS.length).getValues()[0], true);
  } finally {
    lock.releaseLock();
  }
}

function getRegistration_(registrationNo, mobile) {
  if (!registrationNo || !mobile) throw new Error('Registration number and mobile number are required.');
  const sh = getRegSheet_();
  const values = sh.getDataRange().getValues();
  const reg = String(registrationNo).trim().toUpperCase();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][COL.REG]).trim().toUpperCase() === reg &&
        normalizeMobile_(values[i][COL.MOBILE]) === normalizeMobile_(mobile)) {
      return publicObject_(values[i], true);
    }
  }
  throw new Error('Registration not found. Please check Registration No. and Mobile Number.');
}

/* ================================ ADMIN ============================== */

function adminDashboard_(pin) {
  requireAdmin_(pin);
  const sh = getRegSheet_();
  const rows = sh.getLastRow() < 2 ? [] : sh.getDataRange().getValues().slice(1).map(adminObject_);
  return {ok:true, stats:computeStats_(rows), rows:rows};
}

function adminUpdate_(pin, registrationNo, remarks, status) {
  requireAdmin_(pin);
  if (status && ['CONFIRMED','CANCELLED','INCOMPLETE'].indexOf(status) < 0) throw new Error('Invalid status.');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = getRegSheet_();
    const data = sh.getDataRange().getValues();
    let row = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][COL.REG]).trim() === String(registrationNo).trim()) { row = i + 1; break; }
    }
    if (row < 0) throw new Error('Registration not found.');

    if (status) sh.getRange(row, COL.STATUS + 1).setValue(status);
    if (remarks !== undefined && remarks !== null) sh.getRange(row, COL.NOTE + 1).setValue(clean_(remarks));
    sh.getRange(row, COL.BY + 1).setValue('Admin');
    sh.getRange(row, COL.AT + 1).setValue(new Date());

    audit_('ADMIN_UPDATE', registrationNo, JSON.stringify({status:status || null, remarks:remarks || null}), 'Admin');
    return {ok:true, row:adminObject_(sh.getRange(row, 1, 1, HEADERS.length).getValues()[0])};
  } finally {
    lock.releaseLock();
  }
}

function computeStats_(rows) {
  const eq = (a, b) => String(a || '').trim().toLowerCase() === b;
  const n = (key, val) => rows.filter(r => eq(r[key], val)).length;
  return {
    total: rows.length,
    confirmed: n('registrationStatus','confirmed'),
    incomplete: n('registrationStatus','incomplete'),
    cancelled: n('registrationStatus','cancelled'),
    male: n('gender','male'), female: n('gender','female'),
    solapur: n('city','solapur'), outsideSolapur: n('city','outside solapur'),
    urdu: n('language','urdu'), english: n('language','english'), hindi: n('language','hindi')
  };
}

function getAdminPin_() {
  return String(PropertiesService.getScriptProperties().getProperty('ADMIN_PIN') || CONFIG.ADMIN_PIN);
}

function requireAdmin_(pin) {
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('admin_pin_fails') || 0);
  if (fails >= 5) throw new Error('Too many wrong PIN attempts. Try again after 15 minutes.');
  if (String(pin || '') !== getAdminPin_()) {
    cache.put('admin_pin_fails', String(fails + 1), 900);
    throw new Error('Invalid admin PIN.');
  }
  cache.remove('admin_pin_fails');
}

/* ============================== HELPERS ============================== */

function safe_(fn) {
  try { return fn(); }
  catch (err) { return {ok:false, message: msg_(err)}; }
}
function msg_(err) { return (err && err.message) ? err.message : String(err); }

function getRegSheet_() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sh = ss.getSheetByName(CONFIG.REGISTRATION_SHEET) || ss.insertSheet(CONFIG.REGISTRATION_SHEET);
  ensureHeaders_(sh);
  return sh;
}

function getPaymentFolder_() {
  const folders = DriveApp.getFoldersByName(CONFIG.PAYMENT_FOLDER_NAME);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(CONFIG.PAYMENT_FOLDER_NAME);
}

function isClosed_() {
  if (!CONFIG.REGISTRATION_CLOSES_ON) return false;
  const today = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd');
  return today > CONFIG.REGISTRATION_CLOSES_ON;
}

function parsePostBody_(e) {
  if (!e || !e.postData) throw new Error('No POST data received.');
  return JSON.parse(e.postData.contents || '{}');
}

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function validateRegistration_(data) {
  if (!data) throw new Error('No registration data received.');
  const name = clean_(data.name).replace(/\s+/g, ' ');
  const age = Number(data.age);
  const mobile = normalizeMobile_(data.mobile);
  const gender = pick_(data.gender, ['Male','Female']);
  const city = pick_(data.city, ['Solapur','Outside Solapur']);
  const language = pick_(data.language, ['Urdu','English','Hindi']);

  if (!name || name.length < 2) throw new Error('Name is required.');
  if (name.length > 80) throw new Error('Name is too long.');
  if (!age || age < 16 || age > 120) throw new Error('Registration is for persons above 15 years of age.');
  if (!gender) throw new Error('Please select Male or Female.');
  if (!city) throw new Error('Please select a valid location option.');
  if (!/^[6-9]\d{9}$/.test(mobile)) throw new Error('Enter a valid 10-digit Indian mobile number.');
  if (!language) throw new Error('Please select a valid language.');
  return {name:name, age:Math.floor(age), gender:gender, city:city, mobile:mobile, language:language};
}

// case-insensitive match against allowed list; returns canonical value or ''
function pick_(val, allowed) {
  const s = clean_(val).toLowerCase();
  for (let i = 0; i < allowed.length; i++) if (allowed[i].toLowerCase() === s) return allowed[i];
  return '';
}

function nextRegistrationNumber_(sh) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return CONFIG.REGISTRATION_PREFIX + '-000001';
  const values = sh.getRange(2, 1, lastRow - 1, 1).getValues();
  let max = 0;
  values.forEach(function (v) {
    const m = String(v[0]).match(/(\d+)$/);
    if (m) max = Math.max(max, Number(m[1]));
  });
  return CONFIG.REGISTRATION_PREFIX + '-' + String(max + 1).padStart(6, '0');
}

// What a participant may see (no Drive link, no admin notes)
function publicObject_(r) {
  return {
    ok:true,
    registrationNo:r[COL.REG], timestamp:formatDate_(r[COL.TS]), name:r[COL.NAME], age:r[COL.AGE],
    gender:r[COL.GENDER], city:r[COL.CITY], mobile:normalizeMobile_(r[COL.MOBILE]), language:r[COL.LANG],
    fee:r[COL.FEE], paymentStatus:r[COL.PAY], registrationStatus:r[COL.STATUS],
    venue:r[COL.VENUE], programDate:r[COL.DATE], reportingTime:r[COL.REPORT], programTime:r[COL.TIME],
    lastRegistrationDate:r[COL.LAST], hasScreenshot:!!r[COL.SHOT]
  };
}

// What the admin may see
function adminObject_(r) {
  const o = publicObject_(r);
  delete o.ok;
  o.screenshotUrl = r[COL.SHOT];
  o.utr = r[COL.UTR];
  o.remarks = r[COL.REMARKS];
  o.adminNote = r[COL.NOTE];
  return o;
}

function ensureHeaders_(sh) {
  if (sh.getLastRow() === 0) { sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]); return; }
  const current = sh.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  for (let i = 0; i < HEADERS.length; i++) {
    if (current[i] !== HEADERS[i]) { sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]); return; }
  }
}

function audit_(action, registrationNo, details, user) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sh = ss.getSheetByName(CONFIG.AUDIT_SHEET) || ss.insertSheet(CONFIG.AUDIT_SHEET);
    if (sh.getLastRow() === 0) sh.appendRow(['Timestamp','Action','Registration No','Details','User','Status']);
    sh.appendRow([new Date(), action, registrationNo, details || '', user || 'Participant', 'SUCCESS']);
  } catch (e) { /* audit must never break a registration */ }
}

// JPEG / PNG / WEBP signature check (Apps Script bytes are signed, so mask with 0xFF)
function looksLikeImage_(b) {
  if (!b || b.length < 12) return false;
  const x = function (i) { return b[i] & 0xFF; };
  const jpeg = x(0) === 0xFF && x(1) === 0xD8;
  const png = x(0) === 0x89 && x(1) === 0x50 && x(2) === 0x4E && x(3) === 0x47;
  const webp = x(0) === 0x52 && x(1) === 0x49 && x(2) === 0x46 && x(3) === 0x46 &&
               x(8) === 0x57 && x(9) === 0x45 && x(10) === 0x42 && x(11) === 0x50;
  return jpeg || png || webp;
}

function normalizeMobile_(v) { return String(v == null ? '' : v).replace(/\D/g, '').slice(-10); }
function normName_(v) { return String(v == null ? '' : v).trim().toLowerCase().replace(/\s+/g, ' '); }
function clean_(v) { return String(v == null ? '' : v).trim().replace(/[<>]/g, ''); }
function formatDate_(v) {
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v)) {
    return Utilities.formatDate(v, CONFIG.TIMEZONE, 'dd-MM-yyyy HH:mm:ss');
  }
  return String(v || '');
}
