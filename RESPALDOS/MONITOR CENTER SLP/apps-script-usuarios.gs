/************ CONFIG ************/
const SHEET_ID     = '1QOuA10BbJaYctYyY4959TwaPGXRG-pzeHPFZmnqKpxI';
const SHEET_NAME   = 'EVENTOS'; // exactamente este nombre
const USERS_SHEET  = 'USUARIOS';
const SECRET_TOKEN = 'b3f9d1c2-8a5e-4b6b-9412-1e0f5f1f2c93';

const HEADERS = ['Folio','Fecha','Unidad','Evento','Mapa','Estatus','Comentarios'];

/************ SALUD (para verificar deployment) ************/
function doGet(e){
  return json_({ ok:true, ping:'MC-API v2', ts:new Date().toISOString() });
}

/************ ENTRYPOINT ************/
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return json_({ ok:false, error:'No se recibieron datos' });
    }

    const data = parseBody_(e); // mezcla JSON y form-urlencoded
    const action = String(data.action || '').toLowerCase().trim();

    if (action === 'users_upsert' || action === 'users_delete') {
      if (String(data.token || '') !== SECRET_TOKEN) {
        return json_({ ok:false, error:'Token invalido' });
      }

      const ssUsers = SpreadsheetApp.openById(SHEET_ID);
      const shUsers = ssUsers.getSheetByName(USERS_SHEET);
      if (!shUsers) return json_({ ok:false, error:'No existe la hoja USUARIOS' });

      if (action === 'users_upsert') return handleUsersUpsert_(shUsers, data);
      if (action === 'users_delete') return handleUsersDelete_(shUsers, data);
    }

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sh = ensureSheetAndHeaders_(ss, SHEET_NAME, HEADERS);

    // Forzar UPDATE si viene action=update o hay folio
    const wantsUpdate =
      String(data.action || '').toLowerCase() === 'update' ||
      (data.folio && String(data.folio).trim() !== '');

    Logger.log('doPost: action=%s folio=%s wantsUpdate=%s',
               data.action, data.folio, wantsUpdate);

    if (wantsUpdate) {
      if (String(data.token || '') !== SECRET_TOKEN) {
        return json_({ ok:false, error:'Token invalido' });
      }
      return handleUpdate_(sh, data);   // <-- NO exige Unidad/Evento
    }

    // Por defecto: CREATE (alta desde Wialon)
    return handleCreate_(sh, data, e);

  } catch (err) {
    return json_({ ok:false, error:String(err) });
  }
}

/************ HANDLERS ************/
function handleCreate_(sheet, data, e) {
  const N = normalizeKeys_(data);

  const fechaRaw  = pick_(data, ['Fecha','fecha','msg_time','pos_time']) || '';
  const unidadRaw = pick_(data, ['Unidad','unidad','unit','device_id'])  || '';
  const eventoRaw = pick_(data, ['Evento','evento','event'])             || '';
  const mapaRaw   = pick_(data, ['Mapa','mapa','google_link','url'])     || '';

  if (!unidadRaw || !eventoRaw) {
    return json_({ ok:false, error:'Faltan datos obligatorios: Unidad y Evento' });
  }

  const fecha  = toLocalDateTimeMX_(fechaRaw);
  const unidad = decodeURIComponent(String(unidadRaw)).replace(/%20/g,' ');
  const evento = decodeURIComponent(String(eventoRaw)).replace(/%20/g,' ');
  const mapa   = mapaRaw ? decodeURIComponent(String(mapaRaw)) : '';

  const lastRow   = sheet.getLastRow();
  const lastFolio = (lastRow > 1) ? String(sheet.getRange(lastRow, 1).getValue()) : 'FC000000';
  const nextFolio = nextFolio_(lastFolio);

  sheet.appendRow([nextFolio, fecha, unidad, evento, mapa, '', '']);

  if ((e.parameter && e.parameter.debug === '1') || N.debug === '1') {
    logIncoming_(e.postData.contents, e.postData.type || '', nextFolio);
  }

  return json_({ ok:true, folio: nextFolio });
}

function handleUpdate_(sheet, data) {
  const folio = String(data.folio || '').trim();
  if (!folio) return json_({ ok:false, error:'Falta folio' });

  const headers = getHeaders_(sheet);
  const colMap  = colIndexMap_(headers); // {Folio:1, Fecha:2, ...}
  if (!colMap.Folio) return json_({ ok:false, error:'Columna "Folio" no existe' });

  const rowIndex = findRowByFolio_(sheet, folio, colMap.Folio);
  if (rowIndex <= 0) return json_({ ok:false, error:'Folio no encontrado' });

  const estatus    = (data.estatus ?? '').toString().trim();
  const comentario = (data.comentario ?? '').toString().trim();

  const colEstatus     = colMap.Estatus     || ensureColumn_(sheet, headers, 'Estatus');
  const colComentarios = colMap.Comentarios || ensureColumn_(sheet, headers, 'Comentarios');

  const row = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (estatus)    row[colEstatus - 1]     = estatus;
  if (comentario) row[colComentarios - 1] = comentario;

  sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);

  return json_({ ok:true, folio: folio });
}

function handleUsersUpsert_(sheet, data) {
  const headers = getHeaders_(sheet);
  const headerMap = buildUsersHeaderMap_(headers);

  if (headerMap.usuario < 0 || headerMap.password < 0 || headerMap.rol < 0) {
    return json_({ ok:false, error:'La hoja USUARIOS debe tener columnas de usuario, password y rol' });
  }

  const usuario = String(data.usuario || '').trim();
  const password = String(data.password || '').trim();
  const rol = String(data.rol || '').trim();
  const correo = String(data.correo || '').trim();
  const estado = String(data.estado || '').trim().toUpperCase();
  const originalUsuario = String(data.originalUsuario || '').trim().toLowerCase();
  let rowNumber = parseInt(data.rowNumber, 10);

  if (!usuario || !password || !rol) {
    return json_({ ok:false, error:'Faltan datos obligatorios del usuario' });
  }

  const lastCol = Math.max(sheet.getLastColumn(), headers.length);
  const dataRows = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getValues() : [];

  if (!(rowNumber >= 2)) {
    const foundOriginal = dataRows.findIndex(function(r){
      return String(r[headerMap.usuario] || '').trim().toLowerCase() === originalUsuario;
    });
    if (foundOriginal >= 0) rowNumber = foundOriginal + 2;
  }

  if (!(rowNumber >= 2)) {
    const foundByUsuario = dataRows.findIndex(function(r){
      return String(r[headerMap.usuario] || '').trim().toLowerCase() === usuario.toLowerCase();
    });
    if (foundByUsuario >= 0) rowNumber = foundByUsuario + 2;
  }

  if (!(rowNumber >= 2)) rowNumber = sheet.getLastRow() + 1;

  const row = rowNumber <= sheet.getLastRow()
    ? sheet.getRange(rowNumber, 1, 1, lastCol).getValues()[0]
    : new Array(lastCol).fill('');

  row[headerMap.usuario] = usuario;
  row[headerMap.password] = password;
  row[headerMap.rol] = rol;
  if (headerMap.correo >= 0) row[headerMap.correo] = correo;
  if (headerMap.estado >= 0) row[headerMap.estado] = estado || 'ACTIVO';

  sheet.getRange(rowNumber, 1, 1, row.length).setValues([row]);
  return json_({ ok:true, status:'OK', rowNumber: rowNumber });
}

function handleUsersDelete_(sheet, data) {
  const headers = getHeaders_(sheet);
  const headerMap = buildUsersHeaderMap_(headers);
  if (headerMap.usuario < 0) {
    return json_({ ok:false, error:'La hoja USUARIOS no tiene columna de usuario' });
  }

  let rowNumber = parseInt(data.rowNumber, 10);
  const usuario = String(data.usuario || '').trim().toLowerCase();

  if (!(rowNumber >= 2)) {
    const lastCol = Math.max(sheet.getLastColumn(), headers.length);
    const dataRows = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getValues() : [];
    const found = dataRows.findIndex(function(r){
      return String(r[headerMap.usuario] || '').trim().toLowerCase() === usuario;
    });
    if (found >= 0) rowNumber = found + 2;
  }

  if (!(rowNumber >= 2) || rowNumber > sheet.getLastRow()) {
    return json_({ ok:false, error:'No se encontro el usuario a eliminar' });
  }

  sheet.deleteRow(rowNumber);
  return json_({ ok:true, status:'OK', rowNumber: rowNumber });
}

/************ HELPERS ************/
function parseBody_(e){
  let obj = {};
  try { obj = JSON.parse(e.postData.contents || '{}'); } catch (_) {}
  const p = e.parameter || {};
  Object.keys(p).forEach(function(k){ if (!(k in obj)) obj[k] = p[k]; });
  return obj;
}
function normalizeKeys_(o){ const out={}; Object.keys(o||{}).forEach(function(k){ out[k.toLowerCase()]=o[k]; }); return out; }
function pick_(obj, keys){
  for (var i=0;i<keys.length;i++){
    const k = keys[i]; if (obj.hasOwnProperty(k)) return obj[k];
    const low=k.toLowerCase(); if (obj.hasOwnProperty(low)) return obj[low];
  }
  return null;
}
function ensureSheetAndHeaders_(ss, name, headers){
  const sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getLastRow() === 0){
    sh.appendRow(headers);
  } else {
    const cur = getHeaders_(sh);
    let changed=false;
    headers.forEach(function(h){ if (cur.indexOf(h)===-1){ cur.push(h); changed=true; }});
    if (changed) sh.getRange(1,1,1,cur.length).setValues([cur]);
  }
  return sh;
}
function getHeaders_(sheet){
  return sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0].map(String);
}
function colIndexMap_(headers){ const m={}; headers.forEach(function(h,i){ m[h]=i+1; }); return m; }
function ensureColumn_(sheet, headers, name){
  const cur=headers.slice(); cur.push(name);
  sheet.getRange(1,1,1,cur.length).setValues([cur]);
  return cur.length;
}
function findRowByFolio_(sheet, folio, folioCol){
  const last=sheet.getLastRow(); if (last<2) return -1;
  const vals=sheet.getRange(2, folioCol, last-1, 1).getValues();
  for (var i=0;i<vals.length;i++){ if (String(vals[i][0]).trim()===folio) return 2+i; }
  return -1;
}
function nextFolio_(lastFolio){
  var n=parseInt(String(lastFolio).replace(/^FC/,''),10); if(!isFinite(n)) n=0; n++;
  return 'FC'+n.toString().padStart(6,'0');
}
function toLocalDateTimeMX_(raw){
  if (!raw) return new Date().toLocaleString('es-MX',{timeZone:'America/Mexico_City'});
  const n=Number(raw); let d;
  if (isFinite(n) && String(raw).trim()!==''){ d=new Date(n>1e12?n:n*1000); }
  else { const p=Date.parse(String(raw)); d=isNaN(p)?new Date():new Date(p); }
  return d.toLocaleString('es-MX',{timeZone:'America/Mexico_City'});
}
function buildUsersHeaderMap_(headers) {
  return {
    usuario: findUsersHeaderIndex_(headers, ['USUARIO', 'NOMBRE', 'LOGIN']),
    password: findUsersHeaderIndex_(headers, ['PASSWORD', 'CONTRASENA', 'CLAVE', 'PASS']),
    rol: findUsersHeaderIndex_(headers, ['ROL', 'PERFIL']),
    correo: findUsersHeaderIndex_(headers, ['CORREO', 'EMAIL']),
    estado: findUsersHeaderIndex_(headers, ['ESTADO', 'STATUS'])
  };
}
function findUsersHeaderIndex_(headers, aliases) {
  for (var i = 0; i < aliases.length; i++) {
    for (var j = 0; j < headers.length; j++) {
      if (normalizeHeader_(headers[j]) === aliases[i]) return j;
    }
  }
  return -1;
}
function normalizeHeader_(value) {
  return String(value || '')
    .trim()
    .toUpperCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
}
function json_(obj){
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/************ (Opcional) LOG ************/
function logIncoming_(raw, ctype, folio) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName('_Log') || ss.insertSheet('_Log');
  if (sh.getLastRow() === 0) sh.appendRow(['ts','folio','contentType','raw']);
  sh.appendRow([new Date(), folio, ctype, raw]);
}
