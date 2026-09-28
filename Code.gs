/**
 * EK ROZAH TARBIYATI WORKSHOP — Google Sheets + Apps Script Backend
 * Version 3:
 * - Registration form
 * - Screenshot-based payment submission
 * - Auto-confirmation after screenshot upload
 * - Admin dashboard protected by PIN 3249
 * - Same Google Sheet backend
 * - API endpoints for custom Netlify/Vercel frontend
 *
 * IMPORTANT:
 * For the easiest deployment, you can host index.html through Apps Script.
 * If you host index.html on Netlify/Vercel, put the deployed Apps Script Web App
 * URL into BACKEND_URL in index.html.
 */

const CONFIG = {
  SPREADSHEET_ID: '18x8dAcWAGniFsrsV0YGAdVmBWT3aEIUEoPe7TNw-7cI',
  REGISTRATION_SHEET: 'Registrations',
  AUDIT_SHEET: 'Audit_Log',
  PAYMENT_FOLDER_NAME: 'Ek_Rozah_Tarbiyati_Workshop_Payments',

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
  ORGANISATION: 'Zilayi Jamiat Ahle Hadees, Solapur'
};

const HEADERS = [
  'Registration No','Timestamp','Name','Age','Gender','Location (City)',
  'Mobile Number','Language','Registration Fee','Payment Screenshot URL',
  'UTR / Transaction ID','Payment Status','Registration Status',
  'Venue','Program Date','Reporting Time','Program Time',
  'Last Registration Date','Remarks','Admin Note','Admin Updated By','Admin Updated At'
];

function doGet(e) {
  // JSON/JSONP API for externally hosted frontend
  if (e && e.parameter && e.parameter.api === '1') {
    const action = e.parameter.action || '';
    const callback = e.parameter.callback || '';
    let result;

    try {
      if (action === 'config') result = publicConfig_();
      else if (action === 'lookup') result = getRegistration(e.parameter.registrationNo, e.parameter.mobile);
      else if (action === 'adminStats') result = adminStats_(e.parameter.pin);
      else if (action === 'adminList') result = adminList_(e.parameter.pin);
      else result = {ok:false, message:'Unknown API action.'};
    } catch (err) {
      result = {ok:false, message:err.message || String(err)};
    }

    return jsonResponse_(result, callback);
  }

  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle(CONFIG.PROGRAM_NAME + ' — Registration')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    const body = parsePostBody_(e);

    if (body.action === 'register') {
      return jsonResponse_(submitRegistration(body.data || {}));
    }

    if (body.action === 'uploadScreenshot') {
      return jsonResponse_(uploadPaymentScreenshot(
        body.registrationNo, body.mobile, body.fileData
      ));
    }

    if (body.action === 'optionalUTR') {
      return jsonResponse_(submitOptionalUTR(
        body.registrationNo, body.mobile, body.utr
      ));
    }

    if (body.action === 'adminUpdate') {
      return jsonResponse_(adminUpdate(
        body.pin, body.registrationNo, body.remarks, body.status
      ));
    }

    return jsonResponse_({ok:false, message:'Unknown POST action.'});
  } catch (err) {
    return jsonResponse_({ok:false, message:err.message || String(err)});
  }
}

function setupSheets() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  let sh = ss.getSheetByName(CONFIG.REGISTRATION_SHEET);
  if (!sh) sh = ss.insertSheet(CONFIG.REGISTRATION_SHEET);
  ensureHeaders_(sh);
  sh.setFrozenRows(1);

  let audit = ss.getSheetByName(CONFIG.AUDIT_SHEET);
  if (!audit) audit = ss.insertSheet(CONFIG.AUDIT_SHEET);
  if (audit.getLastRow() === 0) {
    audit.getRange(1, 1, 1, 6).setValues([[
      'Timestamp','Action','Registration No','Details','User','Status'
    ]]);
    audit.setFrozenRows(1);
  }

  getPaymentFolder_();
  return 'Sheets and payment folder are ready.';
}

function publicConfig_() {
  return {
    ok:true,
    programName:CONFIG.PROGRAM_NAME,
    programShort:CONFIG.PROGRAM_SHORT,
    fee:CONFIG.FEE,
    upiId:CONFIG.UPI_ID,
    payeeName:CONFIG.PAYEE_NAME,
    venue:CONFIG.VENUE,
    venueAddress:CONFIG.VENUE_ADDRESS,
    programDate:CONFIG.PROGRAM_DATE,
    reportingTime:CONFIG.REPORTING_TIME,
    programTime:CONFIG.PROGRAM_TIME,
    lastRegistrationDate:CONFIG.LAST_REGISTRATION_DATE,
    contact:CONFIG.CONTACT_MARD,
    organisation:CONFIG.ORGANISATION
  };
}

function submitRegistration(data) {
  validateRegistration_(data);

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sh = ss.getSheetByName(CONFIG.REGISTRATION_SHEET) || ss.insertSheet(CONFIG.REGISTRATION_SHEET);
    ensureHeaders_(sh);

    const values = sh.getDataRange().getValues();
    const mobile = normalizeMobile_(data.mobile);

    for (let i = 1; i < values.length; i++) {
      if (normalizeMobile_(values[i][6]) === mobile &&
          String(values[i][12] || '').toUpperCase() !== 'CANCELLED') {
        return {
          ok:false, code:'DUPLICATE',
          message:'This mobile number is already registered.',
          registrationNo:values[i][0]
        };
      }
    }

    const regNo = nextRegistrationNumber_(sh);

    sh.appendRow([
      regNo, new Date(), clean_(data.name), Number(data.age), clean_(data.gender),
      clean_(data.city), mobile, clean_(data.language), CONFIG.FEE, '',
      '', 'AWAITING_SCREENSHOT', 'INCOMPLETE',
      CONFIG.VENUE, CONFIG.PROGRAM_DATE, CONFIG.REPORTING_TIME,
      CONFIG.PROGRAM_TIME, CONFIG.LAST_REGISTRATION_DATE, '', '', '', ''
    ]);

    audit_('REGISTRATION_CREATED', regNo,
      JSON.stringify({name:clean_(data.name), mobile:mobile, language:clean_(data.language)}));

    return {ok:true, registrationNo:regNo, fee:CONFIG.FEE, upiId:CONFIG.UPI_ID};
  } finally {
    lock.releaseLock();
  }
}

function uploadPaymentScreenshot(registrationNo, mobile, fileData) {
  if (!registrationNo || !mobile || !fileData || !fileData.data) {
    throw new Error('Registration number, mobile number and screenshot are required.');
  }

  const reg = getRegistration(registrationNo, mobile);
  const contentType = fileData.type || 'image/jpeg';

  if (!/^image\/(jpeg|jpg|png|webp)$/i.test(contentType)) {
    throw new Error('Please upload a JPG, PNG or WEBP image.');
  }

  const base64 = String(fileData.data).split(',').pop();
  const bytes = Utilities.base64Decode(base64);

  if (bytes.length > 5 * 1024 * 1024) {
    throw new Error('Screenshot is too large. Please upload an image below 5 MB.');
  }

  const safeReg = String(registrationNo).replace(/[^A-Za-z0-9_-]/g, '_');
  const ext = contentType.split('/')[1].replace('jpeg','jpg');
  const filename = safeReg + '_payment_' +
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss') + '.' + ext;

  const blob = Utilities.newBlob(bytes, contentType, filename);
  const folder = getPaymentFolder_();
  const file = folder.createFile(blob);
  file.setDescription('Payment screenshot for registration ' + registrationNo);

  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sh = ss.getSheetByName(CONFIG.REGISTRATION_SHEET);
  const values = sh.getDataRange().getValues();

  let row = -1;
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim() === String(registrationNo).trim() &&
        normalizeMobile_(values[i][6]) === normalizeMobile_(mobile)) {
      row = i + 1;
      break;
    }
  }

  if (row < 0) {
    file.setTrashed(true);
    throw new Error('Registration not found.');
  }

  // Registration is automatically confirmed once screenshot is submitted.
  sh.getRange(row, 10).setValue(file.getUrl());
  sh.getRange(row, 12).setValue('PAYMENT_SCREENSHOT_RECEIVED');
  sh.getRange(row, 13).setValue('CONFIRMED');
  sh.getRange(row, 19).setValue('Payment screenshot received.');
  sh.getRange(row, 22).setValue(new Date());

  audit_('PAYMENT_SCREENSHOT_RECEIVED', registrationNo, file.getUrl());

  return getRegistration(registrationNo, mobile);
}

function submitOptionalUTR(registrationNo, mobile, utr) {
  if (!utr) return getRegistration(registrationNo, mobile);

  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sh = ss.getSheetByName(CONFIG.REGISTRATION_SHEET);
  const data = sh.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][10]).trim().toUpperCase() === String(utr).trim().toUpperCase()) {
      throw new Error('This UTR / transaction ID has already been submitted.');
    }
  }

  let row = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(registrationNo).trim() &&
        normalizeMobile_(data[i][6]) === normalizeMobile_(mobile)) {
      row = i + 1;
      break;
    }
  }
  if (row < 0) throw new Error('Registration not found.');

  sh.getRange(row, 11).setValue(clean_(utr));
  audit_('OPTIONAL_UTR_ADDED', registrationNo, 'UTR: ' + clean_(utr));
  return getRegistration(registrationNo, mobile);
}

function getRegistration(registrationNo, mobile) {
  if (!registrationNo || !mobile) throw new Error('Registration number and mobile number are required.');

  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sh = ss.getSheetByName(CONFIG.REGISTRATION_SHEET);
  if (!sh) throw new Error('Registration sheet not found.');

  const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim() === String(registrationNo).trim() &&
        normalizeMobile_(values[i][6]) === normalizeMobile_(mobile)) {
      return rowToObject_(values[i]);
    }
  }
  throw new Error('Registration not found. Please check Registration No. and Mobile Number.');
}

function adminStats(pin) {
  requireAdmin_(pin);
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sh = ss.getSheetByName(CONFIG.REGISTRATION_SHEET);
  if (!sh || sh.getLastRow() < 2) {
    return {ok:true,total:0,confirmed:0,incomplete:0,cancelled:0,male:0,female:0,
      solapur:0,outsideSolapur:0,urdu:0,english:0,hindi:0};
  }

  const rows = sh.getDataRange().getValues().slice(1);
  const count = (idx, val) => rows.filter(r => String(r[idx]).trim().toLowerCase() === String(val).trim().toLowerCase()).length;

  return {
    ok:true,
    total:rows.length,
    confirmed:count(12,'CONFIRMED'),
    incomplete:count(12,'INCOMPLETE'),
    cancelled:count(12,'CANCELLED'),
    male:count(4,'Male'),
    female:count(4,'Female'),
    solapur:count(5,'Solapur'),
    outsideSolapur:count(5,'Outside solapur'),
    urdu:count(7,'Urdu'),
    english:count(7,'English'),
    hindi:count(7,'Hindi')
  };
}

function adminList(pin) {
  requireAdmin_(pin);
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sh = ss.getSheetByName(CONFIG.REGISTRATION_SHEET);
  if (!sh || sh.getLastRow() < 2) return {ok:true, rows:[]};

  const rows = sh.getDataRange().getValues().slice(1);
  return {
    ok:true,
    rows:rows.map(rowToObject_)
  };
}

function adminUpdate(pin, registrationNo, remarks, status) {
  requireAdmin_(pin);

  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sh = ss.getSheetByName(CONFIG.REGISTRATION_SHEET);
  const data = sh.getDataRange().getValues();

  let row = -1;
  for (let i=1;i<data.length;i++) {
    if (String(data[i][0]).trim() === String(registrationNo).trim()) { row=i+1; break; }
  }
  if (row < 0) throw new Error('Registration not found.');

  if (status && !['CONFIRMED','CANCELLED','INCOMPLETE'].includes(status)) {
    throw new Error('Invalid status.');
  }

  if (status) sh.getRange(row,13).setValue(status);
  sh.getRange(row,20).setValue(clean_(remarks || ''));
  sh.getRange(row,21).setValue('Admin');
  sh.getRange(row,22).setValue(new Date());

  audit_('ADMIN_UPDATE', registrationNo, JSON.stringify({status,remarks}), 'Admin');
  return rowToObject_(sh.getRange(row,1,1,HEADERS.length).getValues()[0]);
}

function requireAdmin_(pin) {
  if (String(pin || '') !== String(CONFIG.ADMIN_PIN)) throw new Error('Invalid admin PIN.');
}

function getPaymentFolder_() {
  const folders = DriveApp.getFoldersByName(CONFIG.PAYMENT_FOLDER_NAME);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(CONFIG.PAYMENT_FOLDER_NAME);
}

function parsePostBody_(e) {
  if (!e || !e.postData) throw new Error('No POST data received.');
  const raw = e.postData.contents || '';
  try { return JSON.parse(raw); }
  catch (_) {
    const p = e.parameter || {};
    return p;
  }
}

function jsonResponse_(obj, callback) {
  const text = JSON.stringify(obj);
  if (callback && /^[A-Za-z_$][\w$]*$/.test(callback)) {
    return ContentService.createTextOutput(callback + '(' + text + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(text).setMimeType(ContentService.MimeType.JSON);
}

function validateRegistration_(data) {
  if (!data) throw new Error('No registration data received.');

  const name = clean_(data.name);
  const age = Number(data.age);
  const gender = clean_(data.gender);
  const city = clean_(data.city);
  const mobile = normalizeMobile_(data.mobile);
  const language = clean_(data.language);

  if (!name) throw new Error('Name is required.');
  if (!age || age < 16) throw new Error('Registration is for persons above 15 years of age.');
  if (!['Male','Female'].includes(gender)) throw new Error('Please select Male or Female.');
  if (!['Solapur','Outside solapur'].includes(city)) throw new Error('Please select a valid city option.');
  if (!/^[6-9]\d{9}$/.test(mobile)) throw new Error('Enter a valid 10-digit Indian mobile number.');
  if (!['Urdu','English','Hindi'].includes(language)) throw new Error('Please select a valid language.');
}

function nextRegistrationNumber_(sh) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return CONFIG.REGISTRATION_PREFIX + '-000001';
  const values = sh.getRange(2,1,lastRow-1,1).getValues().flat();
  let max=0;
  values.forEach(v=>{const m=String(v).match(/(\d+)$/);if(m)max=Math.max(max,Number(m[1]));});
  return CONFIG.REGISTRATION_PREFIX + '-' + String(max+1).padStart(6,'0');
}

function rowToObject_(r) {
  return {
    registrationNo:r[0], timestamp:formatDate_(r[1]), name:r[2], age:r[3],
    gender:r[4], city:r[5], mobile:r[6], language:r[7], fee:r[8],
    screenshotUrl:r[9], utr:r[10], paymentStatus:r[11], registrationStatus:r[12],
    venue:r[13], programDate:r[14], reportingTime:r[15], programTime:r[16],
    lastRegistrationDate:r[17], remarks:r[18], adminNote:r[19]
  };
}

function ensureHeaders_(sh) {
  if (sh.getLastRow() === 0 || sh.getRange(1,1,1,HEADERS.length).getValues()[0].join('') === '') {
    sh.getRange(1,1,1,HEADERS.length).setValues([HEADERS]);
  } else {
    const current = sh.getRange(1,1,1,HEADERS.length).getValues()[0];
    let changed=false;
    for(let i=0;i<HEADERS.length;i++) if(current[i] !== HEADERS[i]) { changed=true; break; }
    if(changed) sh.getRange(1,1,1,HEADERS.length).setValues([HEADERS]);
  }
}

function audit_(action, registrationNo, details, user) {
  const ss=SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sh=ss.getSheetByName(CONFIG.AUDIT_SHEET)||ss.insertSheet(CONFIG.AUDIT_SHEET);
  if(sh.getLastRow()===0) sh.appendRow(['Timestamp','Action','Registration No','Details','User','Status']);
  sh.appendRow([new Date(),action,registrationNo,details||'',user||'Participant','SUCCESS']);
}

function normalizeMobile_(v){return String(v||'').replace(/\D/g,'').slice(-10);}
function clean_(v){return String(v==null?'':v).trim().replace(/[<>]/g,'');}
function formatDate_(v){
  if(Object.prototype.toString.call(v)==='[object Date]'&&!isNaN(v))
    return Utilities.formatDate(v,Session.getScriptTimeZone(),'dd-MM-yyyy HH:mm:ss');
  return String(v||'');
}
