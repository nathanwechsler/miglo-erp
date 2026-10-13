const SHEET_NAME = 'Hoja 1';
const GD_SHEET_NAME = 'GDs';
const FACTURAS_FOLDER_ID = '1aiQKdQVhUcdGnZnipI78vPkm5b_nW5dx';
const ALERT_EMAIL = 'nathan.wechsler@primechef.cl';
const OC_EMITIDAS_SHEET_NAME = 'OC Emitidas';
const OC_EMITIDAS_HEADERS = ['N° OC', 'Fecha', 'Proveedor', 'RUT', 'Subtotal', 'IVA', 'Total', 'Emitida por', 'Timestamp', 'JSON'];
const OC_NUM_PISO = 1120; // la primera OC emitida es la 1121
const OC_IVA = 0.19;

function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // GET para GDs
    if (e.parameter.action === 'get_gds') {
      const sheet = ss.getSheetByName(GD_SHEET_NAME) || ss.insertSheet(GD_SHEET_NAME);
      const data = sheet.getDataRange().getValues();
      if (data.length <= 1) return jsonResponse([]);
      const rows = data.slice(1).map(r => { try { return JSON.parse(r[1]); } catch(e) { return null; } }).filter(Boolean);
      return jsonResponse(rows);
    }

    // GET siguiente N° de OC (vista previa, no reserva)
    if (e.parameter.action === 'get_next_oc') {
      return jsonResponse({ ok: true, next: maxOcNum_(getOcEmitidasSheet_(ss)) + 1 });
    }

    // GET para OCs (default)
    const sheet = ss.getSheetByName(SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return jsonResponse([]);
    const rows = data.slice(1).map(r => { try { return JSON.parse(r[0]); } catch(e) { return null; } }).filter(Boolean);
    return jsonResponse(rows);
  } catch (err) {
    return errorResponse(err);
  }
}

function doPost(e) {
  try {
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

    // ── EMITIR OC (asigna correlativo) ──────────────────────────────────────
    if (payload.action === 'emitir_oc') {
      validarOcEmitida_(payload.oc);
      return conLockOc_(() => {
        const ocSheet = getOcEmitidasSheet_(ss);
        const num = maxOcNum_(ocSheet) + 1;
        ocSheet.appendRow(filaOcEmitida_(num, payload.oc, payload.usuario || ''));
        SpreadsheetApp.flush();
        return jsonResponse({ ok: true, num: num });
      });
    }

    // ── ACTUALIZAR OC EMITIDA (reemisión con el mismo N°) ───────────────────
    if (payload.action === 'actualizar_oc') {
      const num = Number(payload.num);
      if (!num) throw new Error('Falta el N° de la OC a actualizar');
      validarOcEmitida_(payload.oc);
      return conLockOc_(() => {
        const ocSheet = getOcEmitidasSheet_(ss);
        const data = ocSheet.getDataRange().getValues();
        for (let i = 1; i < data.length; i++) {
          if (Number(data[i][0]) === num) {
            // Se conserva quién la emitió originalmente (columna 8)
            ocSheet.getRange(i + 1, 1, 1, OC_EMITIDAS_HEADERS.length)
              .setValues([filaOcEmitida_(num, payload.oc, data[i][7])]);
            SpreadsheetApp.flush();
            return jsonResponse({ ok: true, num: num });
          }
        }
        throw new Error('La OC N° ' + num + ' no existe en "' + OC_EMITIDAS_SHEET_NAME + '"');
      });
    }

    // ── SUBIR ARCHIVO ────────────────────────────────────────────────────────
    if (payload.action === 'upload_file') {
      const folder = DriveApp.getFolderById(FACTURAS_FOLDER_ID);
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
  } catch (err) {
    return errorResponse(err);
  }
}

// ── OC EMITIDAS: helpers ─────────────────────────────────────────────────
function getOcEmitidasSheet_(ss) {
  let sheet = ss.getSheetByName(OC_EMITIDAS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(OC_EMITIDAS_SHEET_NAME);
    sheet.appendRow(OC_EMITIDAS_HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// Mayor N° de la columna A leído como número (ignora celdas vacías o no
// numéricas), con piso OC_NUM_PISO.
function maxOcNum_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return OC_NUM_PISO;
  return sheet.getRange(2, 1, lastRow - 1, 1).getValues().reduce((max, r) => {
    const n = Number(r[0]);
    return isFinite(n) && n > max ? n : max;
  }, OC_NUM_PISO);
}

function validarOcEmitida_(oc) {
  if (!oc || !String(oc.proveedor || '').trim()) throw new Error('La OC no tiene proveedor');
  if (!Array.isArray(oc.items) || !oc.items.length) throw new Error('La OC no tiene productos');
}

// Totales calculados en el servidor con la misma regla del PDF (IVA redondeado).
function filaOcEmitida_(num, oc, emitidaPor) {
  const subtotal = oc.items.reduce((a, it) => a + (Number(it.un) || 0) * (Number(it.valor) || 0), 0);
  const iva = Math.round(subtotal * OC_IVA);
  const completa = Object.assign({}, oc, { num: num });
  return [num, oc.fecha || '', oc.proveedor, oc.rut || '', subtotal, iva, subtotal + iva,
          emitidaPor || '', new Date(), JSON.stringify(completa)];
}

// Ejecuta fn bajo el lock del script para que leer el máximo y escribir la
// fila sea atómico aunque dos personas emitan al mismo tiempo.
function conLockOc_(fn) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (e) {
    throw new Error('El servidor está ocupado emitiendo otra OC. Intenta de nuevo en unos segundos.');
  }
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

// Backend conectado vía clasp - verificación
function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

// Respuesta de error uniforme: el frontend la distingue de una respuesta
// exitosa por el flag `error` y muestra `message` directamente al usuario.
function errorResponse(err) {
  return jsonResponse({ error: true, message: String(err && err.message ? err.message : err) });
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

// ── CANARIO DE AUTORIZACIÓN ──────────────────────────────────────────────
// Corre 1x/día vía trigger de tiempo (ver configurarTriggerCanario más abajo).
// Existe porque el 2026-09-15 el ERP quedó caído en silencio: Google Workspace
// pidió reautenticación de scopes sensibles (Drive/Sheets) y nadie se enteró
// hasta que alguien abrió el frontend. Esta función toca Sheets y Drive con
// la misma autorización que usa doGet/doPost; si falla, avisa por email en
// vez de esperar a que un usuario vea la pantalla de carga pegada.
function chequeoAutorizacion() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.getSheetByName(SHEET_NAME).getRange(1, 1).getValue();
    DriveApp.getFolderById(FACTURAS_FOLDER_ID);
  } catch (err) {
    MailApp.sendEmail(
      ALERT_EMAIL,
      '[Miglo ERP] Apps Script necesita reautorización',
      'El chequeo diario de autorización del ERP OC/GD falló con este error:\n\n' +
      (err && err.message ? err.message : err) +
      '\n\nProbablemente Google Workspace pidió reautenticación de scopes sensibles (Drive/Sheets), ' +
      'igual que el 2026-09-15. Para arreglarlo: abre https://script.google.com/d/' +
      ScriptApp.getScriptId() + '/edit, selecciona cualquier función en el dropdown de ejecución ' +
      '(por ejemplo chequeoAutorizacion) y dale a Ejecutar para forzar el re-consentimiento.'
    );
  }
}

// Crea el trigger diario de chequeoAutorizacion. Idempotente: si ya existe
// un trigger apuntando a esa función, no crea uno duplicado. Hay que
// ejecutarla manualmente UNA VEZ desde el editor de Apps Script (clasp push
// solo sube código, no lo ejecuta) — al correrla la primera vez Google
// pedirá autorizar el envío de mails, es esperable.
function configurarTriggerCanario() {
  const yaExiste = ScriptApp.getProjectTriggers().some(
    t => t.getHandlerFunction() === 'chequeoAutorizacion'
  );
  if (yaExiste) { Logger.log('El trigger ya existe, no se crea uno nuevo.'); return; }
  ScriptApp.newTrigger('chequeoAutorizacion')
    .timeBased()
    .everyDays(1)
    .atHour(7)
    .create();
  Logger.log('Trigger diario creado (07:00 America/Santiago aprox).');
}