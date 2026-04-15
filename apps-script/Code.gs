const SHEET_VEHICULOS = 'vehiculos';
const SHEET_SERVICIOS = 'servicios';

const PROP_SPREADSHEET_ID = 'SPREADSHEET_ID';
const PROP_TOKEN_READ = 'TOKEN_READ';
const PROP_TOKEN_WRITE = 'TOKEN_WRITE';
const PROP_TOKEN_LEGACY = 'TOKEN';

function doGet(e) {
  try {
    const params = e && e.parameter ? e.parameter : {};
    const access = getAccessLevel_(params.token);

    if (!canRead_(access)) {
      return jsonOut_({ ok: false, error: getInvalidTokenMessage_() });
    }

    const what = normalizeSimple_(params.what);
    const ss = getSpreadsheet_();

    if (what === 'veh') {
      return jsonOut_({
        ok: true,
        rows: getSheetObjects_(ss, SHEET_VEHICULOS),
      });
    }

    if (what === 'srv') {
      return jsonOut_({
        ok: true,
        rows: getSheetObjects_(ss, SHEET_SERVICIOS),
      });
    }

    return jsonOut_({ ok: false, error: 'Parametro what invalido' });
  } catch (err) {
    return jsonOut_({
      ok: false,
      error: 'Error en doGet: ' + err.message,
    });
  }
}

function doPost(e) {
  let lock = null;

  try {
    const payload = getJsonPayload_(e);
    const access = getAccessLevel_(payload.token);

    if (!canWrite_(access)) {
      return jsonOut_({ ok: false, error: getInvalidTokenMessage_() });
    }

    const action = normalizeSimple_(payload.action);
    const target = normalizeSimple_(payload.target);
    const ss = getSpreadsheet_();

    lock = LockService.getDocumentLock();
    lock.waitLock(30000);

    if (action === 'add' && target === 'veh') {
      return jsonOut_(appendRowToSheet_(ss, SHEET_VEHICULOS, payload.row || {}));
    }

    if (action === 'add' && target === 'srv') {
      return jsonOut_(appendRowToSheet_(ss, SHEET_SERVICIOS, payload.row || {}));
    }

    if (action === 'upsert_proforma' && target === 'srv') {
      return jsonOut_(upsertProformaRows_(ss, payload.proforma_num, payload.rows || []));
    }

    if (action === 'delete_proforma' && target === 'srv') {
      return jsonOut_(deleteProformaRows_(ss, payload.proforma_num));
    }

    return jsonOut_({ ok: false, error: 'Accion no soportada' });
  } catch (err) {
    return jsonOut_({
      ok: false,
      error: 'Error en doPost: ' + err.message,
    });
  } finally {
    if (lock) {
      try {
        lock.releaseLock();
      } catch (e2) {}
    }
  }
}

function upsertProformaRows_(ss, proformaNum, rowObjs) {
  const sh = ss.getSheetByName(SHEET_SERVICIOS);
  if (!sh) {
    return { ok: false, error: 'No existe la hoja ' + SHEET_SERVICIOS };
  }

  const target = String(proformaNum || '').trim().toUpperCase();
  if (!target) {
    return { ok: false, error: 'Falta proforma_num' };
  }

  if (!Array.isArray(rowObjs) || !rowObjs.length) {
    return { ok: false, error: 'No hay filas para guardar' };
  }

  const lastColumn = sh.getLastColumn();
  if (lastColumn < 1) {
    return { ok: false, error: 'La hoja ' + SHEET_SERVICIOS + ' no tiene encabezados' };
  }

  const headers = sh.getRange(1, 1, 1, lastColumn).getValues()[0].map(function (h) {
    return String(h).trim();
  });

  const idxProforma = findHeaderIndex_(headers, ['proforma_num']);
  if (idxProforma === -1) {
    return { ok: false, error: 'No existe la columna proforma_num en la hoja servicios' };
  }

  const rows = rowObjs.map(function (rowObj) {
    const row = headers.map(function (header) {
      return getValueForHeader_(rowObj, header);
    });

    row[idxProforma] = cleanValue_(target);
    return row;
  });

  const deleted = deleteProformaRowsFromSheet_(sh, target, headers, idxProforma);
  const startRow = sh.getLastRow() + 1;
  sh.getRange(startRow, 1, rows.length, headers.length).setValues(rows);

  return {
    ok: true,
    proforma_num: target,
    deleted: deleted,
    inserted: rows.length,
  };
}

function deleteProformaRows_(ss, proformaNum) {
  const sh = ss.getSheetByName(SHEET_SERVICIOS);
  if (!sh) {
    return { ok: false, error: 'No existe la hoja ' + SHEET_SERVICIOS };
  }

  const values = sh.getDataRange().getValues();
  if (!values.length) {
    return { ok: true, deleted: 0 };
  }

  const headers = values[0].map(function (h) {
    return String(h).trim();
  });

  const idxProforma = findHeaderIndex_(headers, ['proforma_num']);
  if (idxProforma === -1) {
    return { ok: false, error: 'No existe la columna proforma_num en la hoja servicios' };
  }

  const target = String(proformaNum || '').trim().toUpperCase();
  if (!target) {
    return { ok: false, error: 'Falta proforma_num' };
  }

  return {
    ok: true,
    deleted: deleteProformaRowsFromSheet_(sh, target, headers, idxProforma),
  };
}

function deleteProformaRowsFromSheet_(sh, target, headers, idxProforma) {
  const values = sh.getDataRange().getValues();
  if (!values.length) {
    return 0;
  }

  let deleted = 0;
  for (let r = values.length - 1; r >= 1; r--) {
    const current = String(values[r][idxProforma] || '').trim().toUpperCase();
    if (current === target) {
      sh.deleteRow(r + 1);
      deleted++;
    }
  }

  return deleted;
}

function appendRowToSheet_(ss, sheetName, rowObj) {
  const sh = ss.getSheetByName(sheetName);
  if (!sh) {
    return { ok: false, error: 'No existe la hoja ' + sheetName };
  }

  const lastColumn = sh.getLastColumn();
  if (lastColumn < 1) {
    return { ok: false, error: 'La hoja ' + sheetName + ' no tiene encabezados' };
  }

  const headers = sh.getRange(1, 1, 1, lastColumn).getValues()[0].map(function (h) {
    return String(h).trim();
  });

  const row = headers.map(function (header) {
    return getValueForHeader_(rowObj, header);
  });

  const rowNumber = sh.getLastRow() + 1;
  sh.getRange(rowNumber, 1, 1, row.length).setValues([row]);

  return {
    ok: true,
    sheet: sheetName,
    rowNumber: rowNumber,
  };
}

function getSheetObjects_(ss, sheetName) {
  const sh = ss.getSheetByName(sheetName);
  if (!sh) {
    return [];
  }

  const values = sh.getDataRange().getValues();
  if (!values.length) {
    return [];
  }

  const headers = values[0].map(function (h) {
    return String(h).trim();
  });

  const tz = ss.getSpreadsheetTimeZone() || Session.getScriptTimeZone() || 'America/Lima';
  const rows = [];

  for (let r = 1; r < values.length; r++) {
    const obj = {};
    let hasAnyValue = false;

    for (let c = 0; c < headers.length; c++) {
      const key = headers[c] || ('col_' + (c + 1));
      let val = values[r][c];

      if (val instanceof Date) {
        val = Utilities.formatDate(val, tz, "yyyy-MM-dd'T'HH:mm:ss");
      }

      if (val !== '' && val !== null && val !== undefined) {
        hasAnyValue = true;
      }

      obj[key] = val;
    }

    if (hasAnyValue) {
      rows.push(obj);
    }
  }

  return rows;
}

function getValueForHeader_(rowObj, header) {
  if (!rowObj || typeof rowObj !== 'object') {
    return '';
  }

  const normalizedHeader = normalizeKey_(header);

  if (Object.prototype.hasOwnProperty.call(rowObj, header)) {
    return cleanValue_(rowObj[header]);
  }

  const keys = Object.keys(rowObj);
  for (let i = 0; i < keys.length; i++) {
    const k = keys[i];
    if (normalizeKey_(k) === normalizedHeader) {
      return cleanValue_(rowObj[k]);
    }
  }

  return '';
}

function findHeaderIndex_(headers, possibleNames) {
  const normalizedPossible = possibleNames.map(function (x) {
    return normalizeKey_(x);
  });

  for (let i = 0; i < headers.length; i++) {
    const normalizedHeader = normalizeKey_(headers[i]);
    if (normalizedPossible.indexOf(normalizedHeader) !== -1) {
      return i;
    }
  }

  return -1;
}

function cleanValue_(value) {
  if (value === null || value === undefined) {
    return '';
  }

  if (typeof value === 'string') {
    const cleaned = value.replace(/\r\n/g, '\n').trim();
    if (/^[=+\-@]/.test(cleaned)) {
      return "'" + cleaned;
    }
    return cleaned;
  }

  return value;
}

function normalizeKey_(text) {
  return String(text || '')
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '');
}

function normalizeSimple_(text) {
  return String(text || '').trim().toLowerCase();
}

function getSpreadsheet_() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) {
    return active;
  }

  const spreadsheetId = String(
    PropertiesService.getScriptProperties().getProperty(PROP_SPREADSHEET_ID) || ''
  ).trim();

  if (!spreadsheetId) {
    throw new Error('No hay hoja activa y falta configurar SPREADSHEET_ID');
  }

  return SpreadsheetApp.openById(spreadsheetId);
}

function getJsonPayload_(e) {
  const raw = e && e.postData && e.postData.contents ? e.postData.contents : '{}';

  try {
    return JSON.parse(raw || '{}');
  } catch (err) {
    throw new Error('JSON invalido');
  }
}

function getAccessLevel_(token) {
  const currentToken = String(token || '').trim();
  if (!currentToken) {
    return '';
  }

  const props = PropertiesService.getScriptProperties();
  const writeToken = String(props.getProperty(PROP_TOKEN_WRITE) || '').trim();
  const readToken = String(props.getProperty(PROP_TOKEN_READ) || '').trim();
  const legacyToken = String(props.getProperty(PROP_TOKEN_LEGACY) || '').trim();

  if (writeToken && currentToken === writeToken) {
    return 'write';
  }

  if (readToken && currentToken === readToken) {
    return 'read';
  }

  if (legacyToken && currentToken === legacyToken) {
    return 'write';
  }

  return '';
}

function canRead_(access) {
  return access === 'read' || access === 'write';
}

function canWrite_(access) {
  return access === 'write';
}

function getInvalidTokenMessage_() {
  const props = PropertiesService.getScriptProperties();
  const hasAnyToken =
    props.getProperty(PROP_TOKEN_WRITE) ||
    props.getProperty(PROP_TOKEN_READ) ||
    props.getProperty(PROP_TOKEN_LEGACY);

  if (!hasAnyToken) {
    return 'Configura TOKEN_WRITE y/o TOKEN_READ en Script Properties';
  }

  return 'Token invalido';
}

function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
