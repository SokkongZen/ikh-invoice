// ===== ikh-DB : EDU-ikh Invoice Backend (v3: report + Mark Paid + password/token) =====
// SETUP: Project Settings -> Script Properties -> add  ADMIN_PASSWORD = <your password>
// Every request must carry a valid token (obtained from ?action=login&pw=...).
// Columns: A Date | B Invoice No | C Saved At | D Student | E Phone | F Location
//          G Subtotal | H Discount % | I Total | J Validity Start | K Due Date
//          L Subjects | M Time In | N Time Out | O Paid | P Paid Date
//          Q Unit Prices | R Qtys   (comma-joined, one value per subject row)
var HEADERS = ['Date','Invoice No','Saved At','Student','Phone','Location',
               'Subtotal','Discount %','Total','Validity Start','Due Date',
               'Subjects','Time In','Time Out','Paid','Paid Date','Unit Prices','Qtys'];
var FIRST_DATA_ROW = 2;
var BASE_SEQ = 1221;
var TZ = 'Asia/Phnom_Penh';

function getSheet_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
  } else if (String(sheet.getRange(1, 15).getValue()) !== 'Paid') {
    sheet.getRange(1, 15, 1, 2).setValues([['Paid', 'Paid Date']]);
  }
  if (String(sheet.getRange(1, 17).getValue()) !== 'Unit Prices') {
    sheet.getRange(1, 17, 1, 2).setValues([['Unit Prices', 'Qtys']]);
  }
  sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#e7e5e4');
  return sheet;
}

function output_(obj, e) {
  var json = JSON.stringify(obj);
  var cb = e && e.parameter && e.parameter.callback;
  if (cb && /^[A-Za-z0-9_$.]+$/.test(cb)) {
    return ContentService.createTextOutput(cb + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function readRows_(sheet) {
  var last = sheet.getLastRow();
  var count = last >= FIRST_DATA_ROW ? last - FIRST_DATA_ROW + 1 : 0;
  if (!count) return [];
  return sheet.getRange(FIRST_DATA_ROW, 1, count, HEADERS.length).getDisplayValues();
}

function num_(v) {
  return parseFloat(String(v || '0').replace(/[^0-9.-]+/g, '')) || 0;
}


// ---------- AUTH ----------
var TOKEN_TTL_MS = 12 * 60 * 60 * 1000;   // token valid 12 hours
var MAX_FAILS = 5;                         // failed logins allowed ...
var LOCK_SECONDS = 600;                    // ... per 10 minutes

function adminPw_() {
  return PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD') || '';
}

function hex_(bytes) {
  return bytes.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}

function sign_(msg, pw) {
  return hex_(Utilities.computeHmacSha256Signature(msg, pw));
}

function makeToken_(pw) {
  var exp = String(Date.now() + TOKEN_TTL_MS);
  return exp + '.' + sign_(exp, pw);
}

function validToken_(token) {
  var pw = adminPw_();
  if (!pw || !token) return false;
  var parts = String(token).split('.');
  if (parts.length !== 2) return false;
  if (!(Number(parts[0]) > Date.now())) return false;
  var expected = sign_(parts[0], pw), got = parts[1], diff = expected.length ^ got.length;
  for (var i = 0; i < expected.length; i++) diff |= expected.charCodeAt(i) ^ (got.charCodeAt(i) || 0);
  return diff === 0;
}

function login_(e) {
  var pw = adminPw_();
  if (!pw) return { ok: false, error: 'ADMIN_PASSWORD is not set in Script Properties' };
  var cache = CacheService.getScriptCache();
  var fails = Number(cache.get('fails') || 0);
  if (fails >= MAX_FAILS) return { ok: false, error: 'locked', message: 'Too many attempts. Try again in 10 minutes.' };
  var given = String((e.parameter && e.parameter.pw) || '');
  var diff = given.length ^ pw.length;
  for (var i = 0; i < pw.length; i++) diff |= pw.charCodeAt(i) ^ (given.charCodeAt(i) || 0);
  if (diff === 0) {
    cache.remove('fails');
    return { ok: true, token: makeToken_(pw) };
  }
  cache.put('fails', String(fails + 1), LOCK_SECONDS);
  Utilities.sleep(1000);
  return { ok: false, error: 'wrong' };
}


// ---------- STUDY LOG (daily lesson / exam-paper entries; separate tab, created automatically) ----------
var STUDY_HEADERS = ['ID','Date','Student','Subject','Type','Title','A','B','Weakness','Saved At'];

function studySheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('StudyLog');
  if (!sh) {
    sh = ss.insertSheet('StudyLog');
    sh.getRange(1, 1, 1, STUDY_HEADERS.length).setValues([STUDY_HEADERS]).setFontWeight('bold').setBackground('#e7e5e4');
    sh.setFrozenRows(1);
  }
  return sh;
}

function s_(v, max) { return String(v == null ? '' : v).replace(/[\r\n]+/g, ' ').trim().slice(0, max || 200); }

function studyRowById_(sh, id) {
  var last = sh.getLastRow();
  if (last < 2 || !id) return 0;
  var ids = sh.getRange(2, 1, last - 1, 1).getDisplayValues();
  for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return 2 + i;
  return 0;
}

function listStudy_() {
  var sh = studySheet_(), last = sh.getLastRow(), out = [];
  if (last >= 2) {
    var start = Math.max(2, last - 2999);
    sh.getRange(start, 1, last - start + 1, STUDY_HEADERS.length).getDisplayValues().forEach(function (r) {
      if (!r[0]) return;
      out.push({ id: r[0], date: r[1], student: r[2], subject: r[3], type: r[4] === 'E' ? 'E' : 'L', title: r[5],
        a: r[6] === '' ? null : num_(r[6]), b: r[7] === '' ? null : num_(r[7]), weak: r[8] });
    });
  }
  return { entries: out };
}

function saveStudy_(p) {
  var x = p.entry || {};
  if (!s_(x.id) || !s_(x.student) || !s_(x.subject) || !s_(x.title)) throw new Error('entry needs id, student, subject and title');
  var hasNum = x.a !== null && x.a !== undefined && x.a !== '' && x.b !== null && x.b !== undefined && x.b !== '';
  var a = hasNum ? Number(x.a) : '', b = hasNum ? Number(x.b) : '';
  if (hasNum && (!(a >= 0) || !(b > 0) || a > b)) throw new Error('invalid numbers');
  var type = x.type === 'E' ? 'E' : 'L';
  if (type === 'E' && !hasNum) throw new Error('exam paper needs a score');
  var sh = studySheet_();
  var row = [s_(x.id, 60), s_(x.date, 20), s_(x.student, 80), s_(x.subject, 40), type, s_(x.title, 200), a, b, s_(x.weak, 200),
             Utilities.formatDate(new Date(), TZ, 'dd-MM-yyyy HH:mm:ss')];
  var t = studyRowById_(sh, row[0]) || Math.max(sh.getLastRow() + 1, 2);
  sh.getRange(t, 1, 1, 6).setNumberFormat('@');
  sh.getRange(t, 9, 1, 2).setNumberFormat('@');
  sh.getRange(t, 1, 1, row.length).setValues([row]);
  return json_({ ok: true, row: t });
}

function deleteStudy_(p) {
  var sh = studySheet_(), row = studyRowById_(sh, p.id);
  if (!row) return json_({ ok: false, error: 'entry not found' });
  sh.deleteRow(row);
  return json_({ ok: true });
}

// ---------- GET ----------
function doGet(e) {
  try {
    var action = e && e.parameter && e.parameter.action;
    if (action === 'login') return output_(login_(e), e);
    if (!validToken_(e && e.parameter && e.parameter.token)) return output_({ error: 'unauthorized' }, e);
    if (action === 'study') return output_(listStudy_(), e);
    var sheet = getSheet_();
    var data = readRows_(sheet);

    // របាយការណ៍ប្រចាំខែ: ?action=report
    if (action === 'report') {
      var invoices = [];
      for (var i = 0; i < data.length; i++) {
        var r = data[i];
        if (!String(r[1] || '').trim() && !String(r[3] || '').trim()) continue;
        invoices.push({
          row: FIRST_DATA_ROW + i,
          date: r[0], invoiceNo: r[1], savedAt: r[2], name: r[3],
          phone: r[4], location: r[5],
          subtotal: num_(r[6]), discount: num_(r[7]), total: num_(r[8]),
          validityStart: r[9], dueDate: r[10], subjects: r[11],
          paid: String(r[14] || '').toLowerCase() === 'yes',
          paidDate: r[15] || ''
        });
      }
      return output_({ invoices: invoices }, e);
    }

    // លំនាំដើម: បញ្ជីសិស្ស + លេខរៀងវិក្កយបត្រ
    var students = [], seen = {};
    for (var j = data.length - 1; j >= 0; j--) {
      var name = String(data[j][3] || '').trim();
      if (!name || seen[name.toLowerCase()]) continue;
      seen[name.toLowerCase()] = true;
      students.push({
        studentId: 'S-' + (FIRST_DATA_ROW + j),
        name: name,
        phone: String(data[j][4] || ''),
        location: String(data[j][5] || ''),
        price: num_(data[j][6]),
        discount: num_(data[j][7]),
        subject: String(data[j][11] || ''),
        timeIn: String(data[j][12] || ''),
        timeOut: String(data[j][13] || ''),
        prices: String(data[j][16] || ''),
        qtys: String(data[j][17] || '')
      });
    }
    return output_({ nextSeq: BASE_SEQ + data.length, students: students }, e);
  } catch (err) {
    return output_({ error: String(err), nextSeq: BASE_SEQ, students: [], invoices: [] }, e);
  }
}

// ---------- POST ----------
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var p = JSON.parse(e.postData.contents);
    if (!validToken_(p.token)) return json_({ ok: false, error: 'unauthorized' });
    if (p.action === 'saveStudy') return saveStudy_(p);
    if (p.action === 'deleteStudy') return deleteStudy_(p);
    var sheet = getSheet_();
    if (p.action === 'markPaid') return markPaid_(sheet, p);
    return saveInvoice_(sheet, p);
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function saveInvoice_(sheet, p) {
  var row = [p.date, p.invoiceNumber, p.realTime, p.buyerName, p.buyerPhone, p.buyerAddress,
             Number(p.subtotal) || 0, Number(p.discountRate) || 0, Number(p.grandTotal) || 0,
             p.validityStart, p.dueDate, p.subjects, p.timeIns, p.timeOuts];

  var target = 0, last = sheet.getLastRow();
  if (last >= FIRST_DATA_ROW) {
    var nums = sheet.getRange(FIRST_DATA_ROW, 2, last - FIRST_DATA_ROW + 1, 1).getValues();
    for (var i = 0; i < nums.length; i++) {
      if (String(nums[i][0]) === String(p.invoiceNumber)) { target = FIRST_DATA_ROW + i; break; }
    }
  }
  if (!target) target = Math.max(last + 1, FIRST_DATA_ROW);

  sheet.getRange(target, 1, 1, 6).setNumberFormat('@');
  sheet.getRange(target, 10, 1, 9).setNumberFormat('@');
  sheet.getRange(target, 1, 1, row.length).setValues([row]);
  sheet.getRange(target, 17, 1, 2).setValues([[p.prices || '', p.qtys || '']]);   // Paid / Paid Date (O, P) stay untouched
  return json_({ ok: true, row: target });
}

// items: [{row, invoiceNumber}], paid: true|false
function markPaid_(sheet, p) {
  var items = p.items || [];
  var paid = p.paid !== false;
  var today = Utilities.formatDate(new Date(), TZ, 'dd-MM-yyyy');
  var updated = 0;
  var last = sheet.getLastRow();

  items.forEach(function (it) {
    var row = Number(it.row) || 0;
    // ផ្ទៀងផ្ទាត់ថាជួរនេះគឺវិក្កយបត្រត្រឹមត្រូវ បើមិនត្រូវ ស្វែងរកតាមលេខវិក្កយបត្រ
    if (row < FIRST_DATA_ROW || row > last ||
        String(sheet.getRange(row, 2).getDisplayValue()) !== String(it.invoiceNumber)) {
      row = 0;
      if (last >= FIRST_DATA_ROW) {
        var nums = sheet.getRange(FIRST_DATA_ROW, 2, last - FIRST_DATA_ROW + 1, 1).getDisplayValues();
        for (var i = 0; i < nums.length; i++) {
          if (String(nums[i][0]) === String(it.invoiceNumber)) { row = FIRST_DATA_ROW + i; break; }
        }
      }
    }
    if (!row) return;
    var rng = sheet.getRange(row, 15, 1, 2);
    rng.setNumberFormat('@');
    rng.setValues([[paid ? 'Yes' : '', paid ? today : '']]);
    updated++;
  });
  return json_({ ok: true, updated: updated });
}
