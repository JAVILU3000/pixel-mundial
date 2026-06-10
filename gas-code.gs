// ═══════════════════════════════════════════════════════════════════
//  MUNDIAL ETERNO 2026 — Google Apps Script
//  VERSION 08 JUNIO 2026 — LA QUE FUNCIONABA
// ═══════════════════════════════════════════════════════════════════

var SHEET_NAME = 'Ventas';
var FOLDER_ID  = '1XX8xw94p8_h8A5Ub8mnON59-j-XxPo0K';

var PAYPHONE_TOKEN    = 'UhVj1XFoBjyzu0UWaoaWQZzZM1jOmsfVLYvz1DNG9uTzNLCx-x_MQh0qN3Gr7pHmYPGT9I8-glVCDQi7QMLAk7rewPg91ZbHRwfSQmi7cyQLxCqUnuJYbAJEdKqryTmA9LI6M9Y5LYyQLvcakR_OiRsmUQvrbtO_lXatX1qh-zYWQsBifyPuXk4Oj7_oYioHRNEFsshhtta_FEuH9zLaubw7DDdJcPETOVwbARQZpYQodDaGnJMmvMBsM5u6_RXoxAsbm4yrVW83iRjTUfX15h5LQePnhtTPjIlT3b2L1GnrlrK0uyb8ySs0w6WVXsnj9VDSjRU3D1CiJvZPhw7gHyhMcu8';
var PAYPHONE_STORE_ID = '0d38c7cf-abd6-46d7-b6ae-ba497c02ce79';
var PAYPHONE_URL      = 'https://paymentbox.payphonetodoesposible.com/api/confirm';

var COLS = ['fecha','transactionId','monto','x','y','w','h','tamano','nombre','imagenUrl','status','payphoneId','error'];

function doGet(e) {
  var action = (e && e.parameter) ? e.parameter.action : '';
  var out = (action === 'getPixels') ? gasGetPixels() : { error: 'unknown action' };
  return ContentService
    .createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var data;
  try { data = JSON.parse(e.postData.contents); }
  catch(err) { return jsonOut({ status: 'error', message: 'invalid JSON' }); }

  switch (data.action) {
    case 'uploadImage':     return jsonOut(gasUploadImage(data));
    case 'confirmPayment':  return jsonOut(gasSavePayment(data));
    case 'confirmPayphone': return jsonOut(gasConfirmPayphone(data));
    default:                return jsonOut({ status: 'error', message: 'unknown action' });
  }
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(COLS);
  } else if (sheet.getLastRow() === 0) {
    sheet.appendRow(COLS);
  }
  return sheet;
}

function colIndex(sheet) {
  var headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0];
  var idx = {};
  headers.forEach(function(h, i) { if (h) idx[String(h).trim()] = i + 1; });
  return idx;
}

function appendRow(sheet, data) {
  var idx = colIndex(sheet);
  COLS.forEach(function(col) {
    if (!idx[col]) {
      var next = sheet.getLastColumn() + 1;
      sheet.getRange(1, next).setValue(col);
      idx[col] = next;
    }
  });
  var row = new Array(sheet.getLastColumn()).fill('');
  Object.keys(data).forEach(function(k) {
    if (idx[k]) row[idx[k] - 1] = (data[k] !== undefined && data[k] !== null) ? data[k] : '';
  });
  sheet.appendRow(row);
}

function gasGetPixels() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetNames = [SHEET_NAME, 'Hoja 1'];
    var pixels = [];

    sheetNames.forEach(function(name) {
      var sheet = ss.getSheetByName(name);
      if (!sheet) return;
      var lastRow = sheet.getLastRow();
      if (lastRow < 2) return;

      var idx = colIndex(sheet);
      var numCols = sheet.getLastColumn();
      var rows = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();

      rows.forEach(function(row) {
        var g = function(col) { return idx[col] ? row[idx[col] - 1] : ''; };
        var status    = String(g('status'));
        var imagenUrl = String(g('imagenUrl') || '');
        var x = parseInt(g('x'), 10);
        var y = parseInt(g('y'), 10);
        var w = parseInt(g('w'), 10);
        var h = parseInt(g('h'), 10);

        if (isNaN(x) || isNaN(y) || !w || !h) return;
        if (status !== 'Approved' && status !== 'PENDING_VERIFICATION') return;

        pixels.push({ x: x, y: y, w: w, h: h,
                      nombre: String(g('nombre') || ''), imagenUrl: imagenUrl });
      });
    });

    return { pixels: pixels };
  } catch(err) {
    return { pixels: [], error: err.toString() };
  }
}

function gasConfirmPayphone(data) {
  try {
    var payload = {
      id        : parseInt(data.id, 10),
      clientTxId: data.clientTransactionId
    };
    Logger.log('confirmPayphone → payload: ' + JSON.stringify(payload));

    var resp = UrlFetchApp.fetch(PAYPHONE_URL, {
      method            : 'post',
      contentType       : 'application/json',
      headers           : { 'Authorization': 'Bearer ' + PAYPHONE_TOKEN },
      payload           : JSON.stringify(payload),
      muteHttpExceptions: true
    });

    var code = resp.getResponseCode();
    var body = resp.getContentText();
    Logger.log('confirmPayphone ← HTTP ' + code + ' | ' + body);

    var result;
    try { result = JSON.parse(body); } catch(e) { result = {}; }

    if (code !== 200) {
      return { transactionStatus: 'Error', httpStatus: code, body: body };
    }

    if (result.transactionStatus === 'Approved') {
      var ref    = result.reference || '';
      var xMatch = ref.match(/X=(\d+)/);
      var yMatch = ref.match(/Y=(\d+)/);
      var wMatch = ref.match(/W=(\d+)/);
      var hMatch = ref.match(/H=(\d+)/);
      var px = {
        x: xMatch ? parseInt(xMatch[1]) : parseInt(data.x || 0, 10),
        y: yMatch ? parseInt(yMatch[1]) : parseInt(data.y || 0, 10),
        w: wMatch ? parseInt(wMatch[1]) : parseInt(data.w || 1, 10),
        h: hMatch ? parseInt(hMatch[1]) : parseInt(data.h || 1, 10)
      };

      gasSavePayment({
        fecha        : new Date().toISOString(),
        transactionId: result.transactionId || data.id || data.clientTransactionId,
        monto        : (result.amount || 0) / 100,
        x: px.x, y: px.y, w: px.w, h: px.h,
        tamano       : px.w * px.h,
        nombre       : data.nombre || result.email || result.phoneNumber || '',
        imagenUrl    : data.imagenUrl || '',
        status       : 'Approved',
        payphoneId   : data.id || ''
      });
    }

    return { transactionStatus: result.transactionStatus || 'Unknown' };

  } catch(err) {
    return { transactionStatus: 'Error', error: err.toString() };
  }
}

function gasSavePayment(data) {
  try {
    var sheet = getSheet();
    appendRow(sheet, {
      fecha        : data.fecha         || new Date().toISOString(),
      transactionId: data.transactionId || '',
      monto        : data.monto         || 0,
      x            : data.x             !== undefined ? data.x : '',
      y            : data.y             !== undefined ? data.y : '',
      w            : data.w             || '',
      h            : data.h             || '',
      tamano       : data.tamano        || '',
      nombre       : data.nombre        || '',
      imagenUrl    : data.imagenUrl     || '',
      status       : data.status        || '',
      payphoneId   : data.payphoneId    || '',
      error        : data.error         || ''
    });
    return { status: 'ok' };
  } catch(err) {
    return { status: 'error', message: err.toString() };
  }
}

function gasUploadImage(data) {
  try {
    var folder;
    try {
      folder = DriveApp.getFolderById(FOLDER_ID);
    } catch(e) {
      var iter = DriveApp.getFoldersByName('MundialEterno-Imagenes');
      folder = iter.hasNext() ? iter.next() : DriveApp.createFolder('MundialEterno-Imagenes');
    }
    var bytes = Utilities.base64Decode(data.imageBase64);
    var blob  = Utilities.newBlob(bytes, 'image/jpeg', data.filename || 'pixel.jpg');
    var file  = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var url = 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w800';
    return { status: 'ok', url: url };
  } catch(err) {
    return { status: 'error', message: err.toString() };
  }
}
