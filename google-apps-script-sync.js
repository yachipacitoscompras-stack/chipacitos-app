// ================================================
//  YA! Chipacitos — Google Apps Script (Sync)
//  Guarda y carga todos los datos de la app como JSON
//  en una hoja "Sync" de la misma planilla de compras.
// ================================================

const SHEET_ID_SYNC = '17L_oBTiJfJY7JKE85jl7ZU4xgObplDt6H_Q_zGkmSdY'; // ← misma sheet que compras

function okResp(extra)  { return ContentService.createTextOutput(JSON.stringify({ok:true,  ...extra})).setMimeType(ContentService.MimeType.JSON); }
function errResp(msg)   { return ContentService.createTextOutput(JSON.stringify({ok:false, error:msg})).setMimeType(ContentService.MimeType.JSON); }

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss   = SpreadsheetApp.openById(SHEET_ID_SYNC);

    if (data.tipo === 'sync_guardar') return syncGuardar(ss, data.datos);
    if (data.tipo === 'sync_leer')    return syncLeer(ss);

    return errResp('Tipo no soportado: ' + data.tipo);
  } catch(err) {
    return errResp(err.message);
  }
}

function getHojaSync(ss) {
  let sheet = ss.getSheetByName('Sync');
  if (!sheet) {
    sheet = ss.insertSheet('Sync');
    sheet.getRange(1, 1, 1, 2).setValues([['CLAVE', 'VALOR']]);
    sheet.getRange(1, 1, 1, 2).setBackground('#CC3300').setFontColor('#FFB800').setFontWeight('bold');
    sheet.setColumnWidth(1, 280);
    sheet.setColumnWidth(2, 600);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function syncGuardar(ss, datos) {
  const sheet = getHojaSync(ss);

  // Limpiar datos anteriores (mantiene cabecera)
  const last = sheet.getLastRow();
  if (last > 1) sheet.deleteRows(2, last - 1);

  // Agregar timestamp + todos los pares clave/valor
  const filas = [['_sync_fecha', new Date().toISOString()]];
  Object.entries(datos).forEach(([k, v]) => filas.push([k, v]));

  sheet.getRange(2, 1, filas.length, 2).setValues(filas);

  return okResp({ guardados: filas.length });
}

function syncLeer(ss) {
  const sheet = ss.getSheetByName('Sync');
  if (!sheet || sheet.getLastRow() < 2) return okResp({ datos: {} });

  const vals = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  const datos = {};
  vals.forEach(([k, v]) => { if (k && k !== '_sync_fecha') datos[k] = String(v); });

  const fechaFila = vals.find(([k]) => k === '_sync_fecha');
  const fecha = fechaFila ? fechaFila[1] : null;

  return okResp({ datos, fecha });
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'YA! Chipacitos — Sync API online ✅' }))
    .setMimeType(ContentService.MimeType.JSON);
}
