/**
 * LITPA — backend de Google Sheets para el cuestionario anónimo.
 *
 * El formulario envía ítems numerados desde q0 (maestro, texto abierto o Sí/No)
 * más un comentario final. Si se añaden preguntas (hasta 40), la hoja crea columnas sola.
 *
 * Cómo activarlo (una sola vez):
 * 1. Cree una Hoja de cálculo en Google Drive (ej. "Evaluacion LITPA").
 * 2. Extensiones > Apps Script. Borre el código de ejemplo y pegue este archivo.
 * 3. Guardar. En setup, pulse ▶ y autorice la cuenta.
 * 4. Implementar > Nueva implementación > Tipo: Aplicación web
 *      - Descripción: cuestionario
 *      - Ejecutar como: Yo
 *      - Quién tiene acceso: Cualquier persona
 * 5. Copie la URL de la aplicación web (/exec) y péguela en js/config.js
 *    (window.LITPA_SHEETS_URL) y vuelva a generar cuestionario.html si lo usa.
 *
 * Si creó un proyecto de Apps Script suelto (no desde la Hoja), pegue el ID
 * de la hoja entre las comillas de SPREADSHEET_ID (Archivo > Configuración).
 */

var SHEET_NAME = "Respuestas";
var TOTAL_PREGUNTAS = 21; // q0..q20
var MAX_PREGUNTAS = 40;
var MAX_TEXTO = 2000;
var MAX_BODY = 80000;
var SPREADSHEET_ID = "";

function setup() {
  var sheet = ensureSheet_();
  return "Listo: hoja \"" + sheet.getName() + "\". Ahora implemente como Aplicación web.";
}

function doGet(e) {
  try {
    var action = e && e.parameter && e.parameter.action;
    if (action === "resultados") {
      return jsonpOrJson_(getResultados_(), e);
    }
    return jsonpOrJson_({ ok: true, total: countRows_() }, e);
  } catch (err) {
    return jsonpOrJson_({ ok: false, error: String(err.message || err) }, e);
  }
}

function getResultados_() {
  var sheet = ensureSheet_();
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) {
    return { ok: true, total: 0, preguntas: [] };
  }
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var total = data.length;

  var obsIdx = indexOfHeader_(headers, "observacion");
  var qStart = 6;
  var qEnd = obsIdx >= 0 ? obsIdx : lastCol;

  var preguntas = [];
  for (var col = qStart; col < qEnd; col++) {
    var si = 0, no = 0;
    var respuestas = [];
    var esTexto = false;
    for (var row = 0; row < data.length; row++) {
      var val = String(data[row][col] || "").trim().toLowerCase();
      if (!val) continue;
      if (val === "si" || val === "sí") { si++; }
      else if (val === "no") { no++; }
      else { esTexto = true; respuestas.push(String(data[row][col] || "").trim()); }
    }
    if (esTexto && respuestas.length > 0) {
      preguntas.push({ texto: String(headers[col] || "pregunta_" + (col - qStart + 1)), tipo: "texto", respuestas: respuestas });
    } else {
      preguntas.push({ texto: String(headers[col] || "pregunta_" + (col - qStart + 1)), tipo: "si-no", si: si, no: no });
    }
  }

  var observaciones = [];
  if (obsIdx >= 0) {
    for (var r = 0; r < data.length; r++) {
      var obs = String(data[r][obsIdx] || "").trim();
      if (obs) observaciones.push(obs);
    }
  }

  return { ok: true, total: total, preguntas: preguntas, observaciones: observaciones };
}

function doPost(e) {
  try {
    var payload = parsePayload_(e);
    return jsonpOrJson_(saveRow_(payload), e);
  } catch (err) {
    return jsonpOrJson_({ ok: false, error: String(err.message || err) }, e);
  }
}

function parsePayload_(e) {
  var raw = "";
  if (e && e.postData && e.postData.contents) {
    raw = e.postData.contents;
  } else if (e && e.parameter && e.parameter.payload) {
    raw = e.parameter.payload;
  }
  if (!raw) {
    throw new Error("Sin datos");
  }
  if (raw.length > MAX_BODY) {
    throw new Error("Envío demasiado largo");
  }
  var payload = JSON.parse(raw);
  if (!payload || typeof payload !== "object") {
    throw new Error("JSON inválido");
  }
  return payload;
}

function saveRow_(payload) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sheet = ensureSheet_();
    SpreadsheetApp.flush();

    var numero = countRows_() + 1;
    var servidor = new Date();
    var isoServidor = Utilities.formatDate(servidor, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
    var id = String((payload && payload.id) || Utilities.getUuid());
    var isoCliente = safeCell_((payload && payload.fecha_cliente) || "");
    var epochCliente = Number((payload && payload.epoch_cliente) || 0);
    var observacion = safeCell_((payload && payload.observacion) || "").slice(0, MAX_TEXTO);

    var answers = (payload && payload.respuestas) || {};
    var tipos = (payload && payload.tipos) || [];
    var enviadas = Number(payload && payload.total);
    if (!enviadas || enviadas < 1) {
      enviadas = TOTAL_PREGUNTAS;
    }
    if (enviadas > MAX_PREGUNTAS) {
      enviadas = MAX_PREGUNTAS;
    }

    ensureQuestionHeaders_(sheet, enviadas);
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var obs = indexOfHeader_(headers, "observacion");
    var qCount = obs >= 6 ? obs - 6 : enviadas;

    var row = [id, numero, isoServidor, servidor.getTime(), isoCliente, epochCliente];

    // El formulario numera desde q0 (maestros, texto o sí/no).
    for (var i = 0; i < qCount; i++) {
      if (i >= enviadas) {
        row.push("");
        continue;
      }
      var rawAns = String(answers["q" + i] || "");
      var tipo = String(tipos[i] || "si-no");
      if (tipo === "texto" || tipo === "maestros") {
        row.push(safeCell_(rawAns).slice(0, MAX_TEXTO));
        continue;
      }
      var value = rawAns.toLowerCase();
      if (value !== "si" && value !== "no") {
        throw new Error("Cuestionario incompleto");
      }
      row.push(value);
    }

    // Comentario opcional de la pregunta de maestros (si viene).
    var comMaestros = "";
    for (var c = 0; c < enviadas; c++) {
      if (tipos[c] === "maestros") {
        comMaestros = safeCell_(String(answers["q" + c + "_comentario"] || "")).slice(0, MAX_TEXTO);
        break;
      }
    }
    if (comMaestros && observacion) {
      observacion = "[Maestros] " + comMaestros + " | " + observacion;
    } else if (comMaestros) {
      observacion = "[Maestros] " + comMaestros;
    }
    row.push(observacion);

    sheet.appendRow(row);
    SpreadsheetApp.flush();

    var recibidos = countRows_();
    return { ok: true, numero: numero, total: recibidos, id: id, recibido: isoServidor };
  } finally {
    lock.releaseLock();
  }
}

function countRows_() {
  var sheet = ensureSheet_();
  return Math.max(sheet.getLastRow() - 1, 0);
}

function indexOfHeader_(headers, name) {
  var i;
  for (i = 0; i < headers.length; i++) {
    if (String(headers[i]) === name) {
      return i;
    }
  }
  return -1;
}

function ensureQuestionHeaders_(sheet, total) {
  if (sheet.getLastRow() === 0) {
    return;
  }
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var obs = indexOfHeader_(headers, "observacion");
  var existing = obs >= 6 ? obs - 6 : Math.max(lastCol - 7, 0);
  if (total <= existing) {
    return;
  }
  var need = total - existing;
  var i;
  if (obs >= 0) {
    sheet.insertColumnsBefore(obs + 1, need);
    for (i = 0; i < need; i++) {
      sheet.getRange(1, obs + 1 + i).setValue("pregunta_" + (existing + i));
    }
  } else {
    for (i = 0; i < need; i++) {
      sheet.getRange(1, lastCol + 1 + i).setValue("pregunta_" + (existing + i));
    }
  }
}

function getSpreadsheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) {
    return ss;
  }
  var id = String(SPREADSHEET_ID || "").trim();
  if (!id) {
    id = PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID") || "";
  }
  if (id) {
    return SpreadsheetApp.openById(id);
  }
  throw new Error(
    "Abra este archivo desde la Hoja (Extensiones > Apps Script) o ponga el ID en SPREADSHEET_ID."
  );
}

function ensureSheet_() {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    var headers = [
      "id_unico",
      "numero",
      "recibido_iso",
      "epoch_ms",
      "cliente_iso",
      "cliente_epoch_ms",
    ];
    var i;
    for (i = 0; i < TOTAL_PREGUNTAS; i++) {
      headers.push("pregunta_" + i);
    }
    headers.push("observacion");
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
  }
  return sheet;
}

function safeCell_(value) {
  var text = String(value == null ? "" : value);
  if (/^[=+\-@]/.test(text)) {
    return "'" + text;
  }
  return text;
}

function jsonpOrJson_(obj, e) {
  var json = JSON.stringify(obj);
  var callback = e && e.parameter && e.parameter.callback;
  if (callback && /^[A-Za-z_][A-Za-z0-9_]*$/.test(callback)) {
    return ContentService.createTextOutput(callback + "(" + json + ")").setMimeType(
      ContentService.MimeType.JAVASCRIPT
    );
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}
