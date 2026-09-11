const SHEET_NAME = 'Hoja 1';
const GD_SHEET_NAME = 'GDs';

function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // GET para GDs
  if (e.parameter.action === 'get_gds') {
    const sheet = ss.getSheetByName(GD_SHEET_NAME) || ss.insertSheet(GD_SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return jsonResponse([]);
    const rows = data.slice(1).map(r => { try { return JSON.parse(r[1]); } catch(e) { return null; } }).filter(Boolean);
    return jsonResponse(rows);
  }

  // GET para OCs (default)
  const sheet = ss.getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return jsonResponse([]);
  const rows = data.slice(1).map(r => { try { return JSON.parse(r[0]); } catch(e) { return null; } }).filter(Boolean);
  return jsonResponse(rows);
}

function doPost(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  const payload = JSON.parse(e.postData.contents);

  // ── GUARDAR GD ──────────────────────────────────────────────────────────
  if (payload.action === 'save_gd') {
    const gdSheet = ss.getSheetByName(GD_SHEET_NAME) || ss.insertSheet(GD_SHEET_NAME);
    const gd = payload.gd;
    const data = gdSheet.getDataRange().getValues();
    let found = false;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === gd.id) {
        gdSheet.getRange(i + 1, 1, 1, 2).setValues([[gd.id, JSON.stringify(gd)]]);
        found = true; break;
      }
    }
    if (!found) gdSheet.appendRow([gd.id, JSON.stringify(gd)]);
    return jsonResponse({ ok: true });
  }

  // ── ELIMINAR GD ─────────────────────────────────────────────────────────
  if (payload.action === 'delete_gd') {
    const gdSheet = ss.getSheetByName(GD_SHEET_NAME) || ss.insertSheet(GD_SHEET_NAME);
    const data = gdSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === payload.id) { gdSheet.deleteRow(i + 1); break; }
    }
    return jsonResponse({ ok: true });
  }

  // ── GUARDAR TODO (fallback) ──────────────────────────────────────────────
  if (payload.action === 'save_all') {
    sheet.clearContents();
    sheet.appendRow(['data']);
    payload.ocs.forEach(oc => sheet.appendRow([JSON.stringify(oc)]));
    return jsonResponse({ ok: true });
  }

  // ── GUARDAR OC ───────────────────────────────────────────────────────────
  if (payload.action === 'save_oc') {
    const oc = payload.oc;
    const data = sheet.getDataRange().getValues();
    let found = false;
    for (let i = 1; i < data.length; i++) {
      try {
        const row = JSON.parse(data[i][0]);
        if (row.id === oc.id) {
          sheet.getRange(i + 1, 1).setValue(JSON.stringify(oc));
          found = true; break;
        }
      } catch(e) {}
    }
    if (!found) sheet.appendRow([JSON.stringify(oc)]);
    return jsonResponse({ ok: true });
  }

  // ── ELIMINAR OC ──────────────────────────────────────────────────────────
  if (payload.action === 'delete_oc') {
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      try {
        const row = JSON.parse(data[i][0]);
        if (row.id === payload.id) { sheet.deleteRow(i + 1); break; }
      } catch(e) {}
    }
    return jsonResponse({ ok: true });
  }

  // ── SUBIR ARCHIVO ────────────────────────────────────────────────────────
  if (payload.action === 'upload_file') {
    const folders = DriveApp.getFoldersByName('Miglo ERP - Facturas');
    if (!folders.hasNext()) return jsonResponse({ ok: false, error: 'Carpeta no encontrada' });
    const folder = folders.next();
    const blob = Utilities.newBlob(
      Utilities.base64Decode(payload.data),
      payload.mimeType,
      payload.filename
    );
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
    return jsonResponse({ ok: true, url: file.getUrl(), id: file.getId() });
  }

  return jsonResponse({ ok: false });
}

// Backend conectado vía clasp - verificación
function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

function migrarPermisosExistentes() {
  const folders = DriveApp.getFoldersByName('Miglo ERP - Facturas');
  if (!folders.hasNext()) { Logger.log('❌ Carpeta no encontrada'); return; }
  const folder = folders.next();
  const files = folder.getFiles();
  let count = 0;
  while (files.hasNext()) {
    const file = files.next();
    file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
    Logger.log('✅ Actualizado: ' + file.getName());
    count++;
  }
  Logger.log('Listo. ' + count + ' archivos actualizados.');
}