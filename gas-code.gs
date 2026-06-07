// ═══════════════════════════════════════════════════════════════════
//  MUNDIAL ETERNO 2026 — Google Apps Script
//  Pega TODO este código reemplazando lo que tengas actualmente.
// ═══════════════════════════════════════════════════════════════════

// ── CONFIGURA ESTOS DOS VALORES ────────────────────────────────────
var SHEET_NAME = 'Ventas';   // nombre exacto de la pestaña en Sheets
var FOLDER_ID  = '1XX8xw94p8_h8A5Ub8mnON59-j-XxPo0K';
// El ID está en la URL de la carpeta: drive.google.com/drive/folders/ESTE_ES_EL_ID
// ──────────────────────────────────────────────────────────────────

var PAYPHONE_TOKEN    = 'eLg7ZOaSKrgqkghpcf1p_Jk-Bu8uEGBleiDS4NHY3IyTXAsKNpiOJgFyLu-Cjq33q1Gz_ecaTSHEQK57QRtYLNCHRAlVQRsxLEk6sYQ-FvvvnAeXAFj8QQJ-742TSdY3l5DHQFeS7rjre6ElzFPhwZs0FfVYgksgZ_nLcEjm3Zw74mgG6Yn3kzFRD5gt2HU5Hyv2diWfkXbDWgSIVoSl1V3GItNaTF5dbheyB2hy-q3xWm_1Djn0nnlKkdeMaY9_wK3Va-GiUadg3MYmeDDResZIq5OQTh5mO_ygS4wXERRiEbk8MR0X-CyT1l6DGkXCNYh0K8nd66MOr_jrWxui4MUCPno';
var PAYPHONE_STORE_ID = '0d38c7cf-abd6-46d7-b6ae-ba497c02ce79';
var PAYPHONE_URL      = 'https://paymentbox.payphonetodoesposible.com/api/confirm';

// Columnas del spreadsheet (deben coincidir con los encabezados de la fila 1)
var COLS = ['fecha','transactionId','monto','x','y','w','h','tamano','nombre','imagenUrl','status','payphoneId','error'];

// ── ROUTER GET ─────────────────────────────────────────────────────
function doGet(e) {
  var action = (e && e.parameter) ? e.parameter.action : '';
  var out = (action === 'getPixels') ? gasGetPixels() : { error: 'unknown action' };
  return ContentService
    .createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── ROUTER POST ────────────────────────────────────────────────────
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

// ── HELPERS ────────────────────────────────────────────────────────
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

// Lee los encabezados de fila 1 y devuelve { nombreColumna: posición1Indexed }
function colIndex(sheet) {
  var headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0];
  var idx = {};
  headers.forEach(function(h, i) { if (h) idx[String(h).trim()] = i + 1; });
  return idx;
}

// Añade una fila mapeando por nombre de columna; crea columnas faltantes
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

// ── GET: todos los píxeles vendidos ───────────────────────────────
function gasGetPixels() {
  try {
    var sheet = getSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { pixels: [] };

    var idx = colIndex(sheet);
    var numCols = sheet.getLastColumn();
    var rows = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
    var pixels = [];

    rows.forEach(function(row) {
      var g = function(col) { return idx[col] ? row[idx[col] - 1] : ''; };
      var status    = String(g('status'));
      var imagenUrl = String(g('imagenUrl') || '');
      var x = parseInt(g('x'), 10);
      var y = parseInt(g('y'), 10);
      var w = parseInt(g('w'), 10);
      var h = parseInt(g('h'), 10);

      // Solo píxeles con dimensiones válidas y estado aprobado/pendiente
      if (isNaN(x) || isNaN(y) || !w || !h) return;
      if (status !== 'Approved' && status !== 'PENDING_VERIFICATION') return;

      pixels.push({ x: x, y: y, w: w, h: h,
                    nombre: String(g('nombre') || ''), imagenUrl: imagenUrl });
    });

    return { pixels: pixels };
  } catch(err) {
    return { pixels: [], error: err.toString() };
  }
}

// ── POST: confirmar pago con Payphone desde el servidor ───────────
function gasConfirmPayphone(data) {
  try {
    var payload = {
      id                 : parseInt(data.id, 10),
      clientTransactionId: data.clientTransactionId
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

// ── POST: guardar pago en Sheets directamente ─────────────────────
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

// ── POST: subir imagen a Google Drive ─────────────────────────────
function gasUploadImage(data) {
  try {
    var folder;
    try {
      folder = DriveApp.getFolderById(FOLDER_ID);
    } catch(e) {
      // FOLDER_ID inválido o sin acceso → buscar por nombre o crear
      Logger.log('gasUploadImage: getFolderById falló (' + e + '). Buscando por nombre...');
      var iter = DriveApp.getFoldersByName('MundialEterno-Imagenes');
      folder = iter.hasNext() ? iter.next() : DriveApp.createFolder('MundialEterno-Imagenes');
    }
    Logger.log('gasUploadImage: folder=' + folder.getName() + ' id=' + folder.getId());
    var bytes = Utilities.base64Decode(data.imageBase64);
    var blob  = Utilities.newBlob(bytes, 'image/jpeg', data.filename || 'pixel.jpg');
    var file  = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var url = 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w800';
    Logger.log('gasUploadImage: ok → ' + url);
    return { status: 'ok', url: url };
  } catch(err) {
    Logger.log('gasUploadImage ERROR: ' + err);
    return { status: 'error', message: err.toString() };
  }
}
