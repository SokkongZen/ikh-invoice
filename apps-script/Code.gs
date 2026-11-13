// ===== ikh-DB : EDU-ikh Invoice Backend (v2: report + Mark Paid) =====
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

// ---------- GET ----------
function doGet(e) {
  try {
    var action = e && e.parameter && e.parameter.action;
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
