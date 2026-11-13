// ===== ikh-DB : EDU-ikh Invoice Backend (v4: report + Mark Paid + password/token + lessons & homework) =====
// SETUP: Project Settings -> Script Properties -> add  ADMIN_PASSWORD = <your password>
// Every request must carry a valid token (obtained from ?action=login&pw=...).
// Columns: A Date | B Invoice No | C Saved At | D Student | E Phone | F Location
//          G Subtotal | H Discount % | I Total | J Validity Start | K Due Date
//          L Subjects | M Time In | N Time Out | O Paid | P Paid Date
var HEADERS = ['Date','Invoice No','Saved At','Student','Phone','Location',
               'Subtotal','Discount %','Total','Validity Start','Due Date',
               'Subjects','Time In','Time Out','Paid','Paid Date'];
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


// ---------- LESSONS & HOMEWORK (separate tabs, created automatically) ----------
var LESSON_HEADERS = ['ID','Date','Student','Grade','Subject','Topic','Hours','Errors','Saved At'];
var HW_HEADERS = ['ID','Assigned','Student','Subject','Topic','Due','Status','Sent Date'];

function tab_(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#e7e5e4');
    sh.setFrozenRows(1);
  }
  return sh;
}

function s_(v, max) { return String(v == null ? '' : v).replace(/[\r\n]+/g, ' ').trim().slice(0, max || 200); }

function findRowById_(sh, id) {
  var last = sh.getLastRow();
  if (last < 2 || !id) return 0;
  var ids = sh.getRange(2, 1, last - 1, 1).getDisplayValues();
  for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return 2 + i;
  return 0;
}

function listLessons_() {
  var ls = tab_('Lessons', LESSON_HEADERS), hs = tab_('Homework', HW_HEADERS);
  var lessons = [], homework = [];
  var ll = ls.getLastRow();
  if (ll >= 2) {
    var start = Math.max(2, ll - 399);
    ls.getRange(start, 1, ll - start + 1, LESSON_HEADERS.length).getDisplayValues().forEach(function (r, i) {
      if (!r[0]) return;
      lessons.push({ row: start + i, id: r[0], date: r[1], student: r[2], grade: r[3], subject: r[4], topic: r[5],
        hours: num_(r[6]), errors: r[7] ? String(r[7]).split(' | ') : [] });
    });
  }
  var hl = hs.getLastRow();
  if (hl >= 2) {
    var hstart = Math.max(2, hl - 599);
    hs.getRange(hstart, 1, hl - hstart + 1, HW_HEADERS.length).getDisplayValues().forEach(function (r, i) {
      if (!r[0]) return;
      homework.push({ row: hstart + i, id: r[0], assigned: r[1], student: r[2], subject: r[3], topic: r[4], due: r[5],
        sent: String(r[6]).toLowerCase() === 'sent', sentDate: r[7] });
    });
  }
  return { lessons: lessons, homework: homework };
}

function saveLesson_(p) {
  var L = p.lesson || {}, H = p.homework || null;
  if (!L.id || !s_(L.student) || !s_(L.topic)) throw new Error('lesson needs id, student and topic');
  var ls = tab_('Lessons', LESSON_HEADERS);
  var errors = (L.errors || []).map(function (x) { return s_(x, 60); }).filter(String).join(' | ');
  var lrow = [s_(L.id, 60), s_(L.date, 20), s_(L.student, 80), s_(L.grade, 4), s_(L.subject, 40), s_(L.topic, 200),
              Number(L.hours) || 0, errors, Utilities.formatDate(new Date(), TZ, 'dd-MM-yyyy HH:mm:ss')];
  var lt = findRowById_(ls, lrow[0]) || Math.max(ls.getLastRow() + 1, 2);
  ls.getRange(lt, 1, 1, 6).setNumberFormat('@');
  ls.getRange(lt, 8, 1, 2).setNumberFormat('@');
  ls.getRange(lt, 1, 1, lrow.length).setValues([lrow]);

  var hrowNo = 0;
  if (H && H.id && s_(H.student)) {
    var hs = tab_('Homework', HW_HEADERS);
    var hrow = [s_(H.id, 60), s_(H.assigned, 20), s_(H.student, 80), s_(H.subject, 40), s_(H.topic, 200), s_(H.due, 20), 'Pending', ''];
    hrowNo = findRowById_(hs, hrow[0]);
    if (hrowNo) {   // re-save: keep the current status
      hrow[6] = String(hs.getRange(hrowNo, 7).getDisplayValue() || 'Pending');
      hrow[7] = String(hs.getRange(hrowNo, 8).getDisplayValue() || '');
    } else hrowNo = Math.max(hs.getLastRow() + 1, 2);
    hs.getRange(hrowNo, 1, 1, hrow.length).setNumberFormat('@');
    hs.getRange(hrowNo, 1, 1, hrow.length).setValues([hrow]);
  }
  return json_({ ok: true, lessonRow: lt, homeworkRow: hrowNo });
}

// {id, sent: true|false}
function homeworkStatus_(p) {
  var hs = tab_('Homework', HW_HEADERS);
  var row = findRowById_(hs, p.id);
  if (!row) return json_({ ok: false, error: 'homework not found' });
  var sent = p.sent !== false;
  var today = Utilities.formatDate(new Date(), TZ, 'dd-MM-yyyy');
  var rng = hs.getRange(row, 7, 1, 2);
  rng.setNumberFormat('@');
  rng.setValues([[sent ? 'Sent' : 'Pending', sent ? today : '']]);
  return json_({ ok: true });
}

// ---------- GET ----------
function doGet(e) {
  try {
    var action = e && e.parameter && e.parameter.action;
    if (action === 'login') return output_(login_(e), e);
    if (!validToken_(e && e.parameter && e.parameter.token)) return output_({ error: 'unauthorized' }, e);
    if (action === 'lessons') return output_(listLessons_(), e);
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
        subject: String(data[j][11] || ''),
        timeIn: String(data[j][12] || ''),
        timeOut: String(data[j][13] || '')
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
    if (p.action === 'saveLesson') return saveLesson_(p);
    if (p.action === 'homeworkStatus') return homeworkStatus_(p);
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
  sheet.getRange(target, 10, 1, 7).setNumberFormat('@');
  sheet.getRange(target, 1, 1, row.length).setValues([row]);
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
