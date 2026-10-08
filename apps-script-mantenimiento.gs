/****************************************************************************************
 * MONITOR CENTER · MÓDULO DE MANTENIMIENTO
 * Apps Script para sincronizar unidad + mantenimiento desde Wialon
 *
 * Requisitos:
 *  - Google Sheet con ID de la hoja principal
 *  - Token Wialon válido
 *  - La hoja "MANTENIMIENTO" se crea automáticamente
 *
 * Acciones soportadas:
 *  - action=sync_units: trae el catálogo de unidades y actualiza la hoja de mantenimiento
 *  - action=sync_maintenance: trae datos relevantes y calcula kilometraje / servicio / preventivo
 *
 ****************************************************************************************/

const WIALON_API_URL = 'https://hst-api.wialon.com/wialon/ajax.html';
const SHEET_ID = '1QOuA10BbJaYctYyY4959TwaPGXRG-pzeHPFZmnqKpxI';
const MAINT_SHEET = 'MANTENIMIENTO';
const LOG_SHEET = '_MANTENIMIENTO_LOG';
const WIALON_TOKEN = 'PASTE_AQUI_TU_TOKEN_WIALON';

const MAINT_HEADERS = [
  'Unidad',
  'Placa',
  'Modelo',
  'Kilometraje',
  'HorasMotor',
  'CombustibleActual',
  'UltimaConexion',
  'UltimoServicio',
  'KmUltimoAceite',
  'KmProximoAceite',
  'HorasUltimoPreventivo',
  'ProximoPreventivo',
  'EstadoMantenimiento',
  'UltimoEvento',
  'Actualizado'
];

const KM_INTERVALO_ACEITE = 5000;
const KM_INTERVALO_PREVENTIVO = 10000;

/****************************************************************************************
 * HEALTH CHECK
 ****************************************************************************************/
function doGet(e) {
  return json_({
    ok: true,
    ping: 'MC-MANTENIMIENTO v1',
    ts: new Date().toISOString()
  });
}

/****************************************************************************************
 * ENTRY POINT
 ****************************************************************************************/
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return json_({ ok: false, error: 'No se recibieron datos' });
    }

    const data = parseBody_(e);
    const action = String(data.action || '').toLowerCase().trim();

    if (action === 'sync_units') return syncUnits_(data);
    if (action === 'sync_maintenance') return syncMaintenance_(data);

    return json_({ ok: false, error: 'Acción no soportada. Usa sync_units o sync_maintenance.' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

/****************************************************************************************
 * WIALON LOGIN / REQUESTS
 ****************************************************************************************/
function loginWialon_() {
  const payload = {
    token: WIALON_TOKEN
  };

  const response = fetchWialon_('token/login', payload, false);
  if (!response || !response.eid) {
    throw new Error('La autenticación con Wialon falló. Verifica WIALON_TOKEN.');
  }

  return response;
}

function fetchWialon_(svc, params, isSessionRequired) {
  const base = {
    svc: svc,
    params: JSON.stringify(params || {})
  };

  const url = WIALON_API_URL + '?svc=' + encodeURIComponent(svc) + '&params=' + encodeURIComponent(JSON.stringify(params || {}));
  const options = {
    method: 'get',
    muteHttpExceptions: true,
    headers: {
      'Accept': 'application/json',
      'Content-Type': 'application/json;charset=UTF-8'
    }
  };

  const raw = UrlFetchApp.fetch(url, options);
  const text = raw.getContentText();
  if (!text) return {};

  try {
    const json = JSON.parse(text);
    if (json && json.error) {
      throw new Error(json.error.message || json.error || 'Error de Wialon');
    }
    return json || {};
  } catch (err) {
    if (isSessionRequired) {
      throw new Error('Respuesta no JSON válida de Wialon: ' + text.slice(0, 200));
    }
    return {};
  }
}

function getWialonUnits_() {
  const auth = loginWialon_();
  const request = {
    spec: {
      itemsType: 'avl_unit',
      propName: 'sys_name',
      propValueMask: '*',
      sortType: 'sys_name'
    },
    force: 0,
    flags: 1,
    from: 0,
    to: 500
  };

  const response = fetchWialon_('core/search_items', request, true);
  const items = response.items || response.data || [];

  if (!Array.isArray(items)) {
    return [];
  }

  return items.map(function(item) {
    const id = item.id || item.uid || item.unit_id || '';
    const name = item.name || item.sys_name || item.n || '';
    const properties = item.properties || item.props || {};
    return {
      id: id,
      unit: String(name || id).trim(),
      placa: String(properties.plate || properties.placa || item.plate || item.placa || '').trim(),
      modelo: String(properties.model || properties.modelo || item.model || item.modelo || '').trim(),
      propiedades: properties,
      raw: item
    };
  }).filter(function(u) { return u.unit || u.id; });
}

function getUnitExtraData_(unitId) {
  const payload = {
    itemId: unitId,
    flags: 0,
    fields: 'unit.name,unit.model,unit.imei,unit.vehicle,unit.last_msg_time,unit.mileage,unit.odometer,unit.mileage_km,unit.hours,unit.engine_hours'
  };

  const response = fetchWialon_('unit/get_info', payload, true);
  return response || {};
}

/****************************************************************************************
 * MANTENIMIENTO
 ****************************************************************************************/
function syncUnits_(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ensureSheetAndHeaders_(ss, MAINT_SHEET, MAINT_HEADERS);

  const units = getWialonUnits_();
  const rows = units.map(function(unit) {
    return buildMaintenanceRow_(unit, false);
  });

  if (!rows.length) {
    return json_({ ok: true, count: 0, message: 'No hay unidades en Wialon o no se pudo autenticarse.' });
  }

  const values = rows.map(function(r) { return r.row; });
  const firstRow = 2;
  const lastRow = sh.getLastRow();

  if (lastRow > 1) {
    sh.getRange(2, 1, lastRow - 1, MAINT_HEADERS.length).clear();
  }

  if (values.length > 0) {
    sh.getRange(firstRow, 1, values.length, values[0].length).setValues(values);
  }

  return json_({
    ok: true,
    count: rows.length,
    sheet: MAINT_SHEET,
    action: 'sync_units'
  });
}

function syncMaintenance_(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ensureSheetAndHeaders_(ss, MAINT_SHEET, MAINT_HEADERS);

  const units = getWialonUnits_();
  const rows = units.map(function(unit) {
    return buildMaintenanceRow_(unit, true);
  });

  if (!rows.length) {
    return json_({ ok: true, count: 0, message: 'No hubo datos para procesar.' });
  }

  const values = rows.map(function(r) { return r.row; });

  const lastRow = sh.getLastRow();
  if (lastRow > 1) {
    sh.getRange(2, 1, lastRow - 1, MAINT_HEADERS.length).clear();
  }

  sh.getRange(2, 1, values.length, values[0].length).setValues(values);

  return json_({
    ok: true,
    count: rows.length,
    sheet: MAINT_SHEET,
    action: 'sync_maintenance'
  });
}

function buildMaintenanceRow_(unit, withExtraLookup) {
  const now = new Date();
  const unidad = unit.unit || unit.id || '';
  const placa = unit.placa || '';
  const modelo = unit.modelo || '';

  let kmActual = 0;
  let horasMotor = 0;
  let combustibleActual = 0;
  let ultimaConexion = '';
  let ultimoServicio = '';
  let kmUltimoAceite = 0;
  let kmProximoAceite = 0;
  let horasUltimoPreventivo = 0;
  let proximoPreventivo = 0;
  let estadoMantenimiento = 'Sin datos';
  let ultimoEvento = 'Sin evento';

  if (withExtraLookup && unit.id) {
    const extra = getUnitExtraData_(unit.id);
    const unitInfo = extra.item || extra.unit || extra.data || extra.result || {};
    const data = unitInfo.data || unitInfo;

    kmActual = normalizeNumber_(extractFirstValue_(data, ['odometer', 'mileage', 'distance', 'total_distance', 'kilometraje', 'kilometraje_total', 'motor_mileage']));
    horasMotor = normalizeNumber_(extractFirstValue_(data, ['engine_hours', 'hours', 'hours_motor', 'hrs_engine', 'engineHours']));
    combustibleActual = normalizeNumber_(extractFirstValue_(data, ['fuel', 'fuel_level', 'combustible', 'level', 'fuel_level_percent']));
    ultimaConexion = toMxDateString_(extractFirstValue_(data, ['last_msg_time', 'last_message_time', 'time', 'last_update', 'timestamp']));
    ultimoEvento = extractFirstValue_(data, ['last_event', 'last_status', 'status', 'last_alert', 'event']) || 'Sin evento';
  }

  if (!kmActual) {
    kmActual = 0;
  }
  if (!horasMotor) {
    horasMotor = 0;
  }
  if (!combustibleActual) {
    combustibleActual = 0;
  }

  kmUltimoAceite = Math.max(0, kmActual - (kmActual % KM_INTERVALO_ACEITE));
  kmProximoAceite = kmActual + (KM_INTERVALO_ACEITE - (kmActual % KM_INTERVALO_ACEITE || KM_INTERVALO_ACEITE));
  horasUltimoPreventivo = Math.max(0, horasMotor - (horasMotor % 250));
  proximoPreventivo = horasMotor + Math.max(250 - (horasMotor % 250 || 250), 0);

  if (kmActual >= 0 && kmActual < 5000) {
    estadoMantenimiento = 'OK';
  } else if (kmActual >= 5000 && kmActual < 10000) {
    estadoMantenimiento = 'ACEITE PRÓXIMO';
  } else {
    estadoMantenimiento = 'REVISIÓN REQUERIDA';
  }

  if (kmActual === 0 && horasMotor === 0 && combustibleActual === 0) {
    estadoMantenimiento = 'SIN DATOS';
  }

  const row = [
    unidad,
    placa,
    modelo,
    formatNumber_(kmActual),
    formatNumber_(horasMotor),
    formatNumber_(combustibleActual),
    ultimaConexion || 'Sin datos',
    ultimoServicio || 'Sin servicio',
    formatNumber_(kmUltimoAceite),
    formatNumber_(kmProximoAceite),
    formatNumber_(horasUltimoPreventivo),
    formatNumber_(proximoPreventivo),
    estadoMantenimiento,
    ultimoEvento,
    new Date().toISOString()
  ];

  return { unit: unidad, row: row };
}

/****************************************************************************************
 * SHEET HELPERS
 ****************************************************************************************/
function ensureSheetAndHeaders_(ss, name, headers) {
  const sh = ss.getSheetByName(name) || ss.insertSheet(name);

  if (sh.getLastRow() === 0) {
    sh.appendRow(headers);
    return sh;
  }

  const current = getHeaders_(sh);
  const missing = headers.filter(function(h) { return current.indexOf(h) === -1; });
  if (missing.length) {
    const merged = current.concat(missing);
    sh.getRange(1, 1, 1, merged.length).setValues([merged]);
  }

  return sh;
}

function getHeaders_(sheet) {
  const values = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues();
  return values[0].map(function(v) { return String(v || ''); });
}

function colIndexMap_(headers) {
  const map = {};
  headers.forEach(function(h, i) {
    map[h] = i + 1;
  });
  return map;
}

/****************************************************************************************
 * DATA HELPERS
 ****************************************************************************************/
function parseBody_(e) {
  let obj = {};
  try {
    obj = JSON.parse(e.postData.contents || '{}');
  } catch (_) {
    obj = {};
  }

  const p = e.parameter || {};
  Object.keys(p).forEach(function(k) {
    if (!(k in obj)) obj[k] = p[k];
  });

  return obj;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function normalizeNumber_(value) {
  if (value === null || value === undefined || value === '') return 0;
  const num = Number(String(value).replace(/[^0-9.\-]/g, '').replace(/,/g, ''));
  return isNaN(num) ? 0 : num;
}

function formatNumber_(value) {
  const num = Number(value || 0);
  if (!isFinite(num)) return '0';
  return Number(num).toLocaleString('es-MX');
}

function extractFirstValue_(obj, keys) {
  if (!obj || typeof obj !== 'object') return '';

  for (let i = 0; i < keys.length; i++) {
    const key = keys[i];
    if (Object.prototype.hasOwnProperty.call(obj, key) && obj[key] !== undefined && obj[key] !== null) {
      return obj[key];
    }

    const nested = findNestedValue_(obj, key);
    if (nested !== undefined && nested !== null) {
      return nested;
    }
  }

  return '';
}

function findNestedValue_(obj, targetKey) {
  if (!obj || typeof obj !== 'object') return undefined;

  for (const key in obj) {
    if (!Object.prototype.hasOwnProperty.call(obj, key)) continue;
    const value = obj[key];
    if (key === targetKey) {
      return value;
    }
    if (value && typeof value === 'object') {
      const found = findNestedValue_(value, targetKey);
      if (found !== undefined) return found;
    }
  }
  return undefined;
}

function toMxDateString_(raw) {
  if (!raw) return '';

  let parsed = raw;
  if (typeof raw === 'number') {
    parsed = new Date(raw > 1e12 ? raw : raw * 1000);
  } else if (typeof raw === 'string') {
    const cleaned = raw.trim();
    if (/^\d+$/.test(cleaned)) {
      const n = Number(cleaned);
      parsed = new Date(n > 1e12 ? n : n * 1000);
    } else {
      const p = Date.parse(cleaned);
      if (!isNaN(p)) parsed = new Date(p);
      else parsed = new Date(cleaned);
    }
  }

  if (parsed instanceof Date && !isNaN(parsed.getTime())) {
    return parsed.toLocaleString('es-MX', { timeZone: 'America/Mexico_City' });
  }

  return String(raw);
}

function logMaintenance_(message, extra) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName(LOG_SHEET) || ss.insertSheet(LOG_SHEET);
  if (sh.getLastRow() === 0) {
    sh.appendRow(['ts', 'mensaje', 'extra']);
  }
  sh.appendRow([new Date().toISOString(), message, JSON.stringify(extra || {})]);
}
