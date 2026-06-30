// ================================================
//  YA! Chipacitos — Google Apps Script (Sync)
//  Guarda y carga todos los datos de la app via JSONP
// ================================================

const SHEET_ID_SYNC = '17L_oBTiJfJY7JKE85jl7ZU4xgObplDt6H_Q_zGkmSdY';

function wrap(callback, data) {
  const json = JSON.stringify(data);
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const cb   = e.parameter.callback || '';
  const tipo = e.parameter.tipo || '';
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID_SYNC);
    if (tipo === 'sync_guardar') {
      const datos = JSON.parse(decodeURIComponent(e.parameter.datos || '{}'));
      return wrap(cb, syncGuardar(ss, datos));
    }
    if (tipo === 'sync_leer') {
      return wrap(cb, syncLeer(ss));
    }
    return wrap(cb, { ok: true, status: 'YA! Chipacitos — Sync API online ✅' });
  } catch(err) {
    return wrap(cb, { ok: false, error: err.message });
  }
}

function getHojaSync(ss) {
  let sheet = ss.getSheetByName('Sync');
  if (!sheet) {
    sheet = ss.insertSheet('Sync');
    sheet.getRange(1,1,1,2).setValues([['CLAVE','VALOR']]);
    sheet.getRange(1,1,1,2).setBackground('#CC3300').setFontColor('#FFB800').setFontWeight('bold');
    sheet.setColumnWidth(1, 280);
    sheet.setColumnWidth(2, 600);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function syncGuardar(ss, datos) {
  const sheet = getHojaSync(ss);
  const last = sheet.getLastRow();
  if (last > 1) sheet.deleteRows(2, last - 1);
  const filas = [['_sync_fecha', new Date().toISOString()]];
  Object.entries(datos).forEach(([k,v]) => filas.push([k, v]));
  sheet.getRange(2, 1, filas.length, 2).setValues(filas);
  return { ok: true, guardados: filas.length };
}

function syncLeer(ss) {
  const sheet = ss.getSheetByName('Sync');
  if (!sheet || sheet.getLastRow() < 2) return { ok: true, datos: {} };
  const vals = sheet.getRange(2, 1, sheet.getLastRow()-1, 2).getValues();
  const datos = {};
  vals.forEach(([k,v]) => { if (k && k !== '_sync_fecha') datos[k] = String(v); });
  return { ok: true, datos };
}
