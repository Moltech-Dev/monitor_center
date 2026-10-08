const SECRET_TOKEN = 'b3f9d1c2-8a5e-4b6b-9412-1e0f5f1f2c93';
const SHEET_USUARIOS = 'USUARIOS';

function doPost(e) {
  try {
    const p = (e && e.parameter) || {};
    if (p.token !== SECRET_TOKEN) return jsonOut({ ok: false, error: 'Token invalido' });

    switch ((p.action || '').trim()) {
      case 'users_upsert':
        return jsonOut(upsertUsuario_(p));
      case 'users_delete':
        return jsonOut(deleteUsuario_(p));
      default:
        return jsonOut({ ok: false, error: 'Accion no soportada' });
    }
  } catch (err) {
    return jsonOut({ ok: false, error: err && err.message ? err.message : String(err) });
  }
}

function upsertUsuario_(p) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_USUARIOS);
  if (!sh) throw new Error('No existe la hoja USUARIOS');

  const lastCol = Math.max(sh.getLastColumn(), 5);
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(normalizeHeader_);
  const idx = {
    usuario: findHeaderIndex_(headers, ['USUARIO', 'NOMBRE', 'LOGIN']),
    password: findHeaderIndex_(headers, ['PASSWORD', 'CONTRASENA', 'CLAVE', 'PASS']),
    rol: findHeaderIndex_(headers, ['ROL', 'PERFIL']),
    correo: findHeaderIndex_(headers, ['CORREO', 'EMAIL']),
    estado: findHeaderIndex_(headers, ['ESTADO', 'STATUS'])
  };

  if ([idx.usuario, idx.password, idx.rol].some(v => v < 0)) {
    throw new Error('La hoja USUARIOS debe tener columnas para usuario, password y rol');
  }

  const usuario = String(p.usuario || '').trim();
  const password = String(p.password || '').trim();
  const rol = String(p.rol || '').trim();
  const correo = String(p.correo || '').trim();
  const estado = String(p.estado || '').trim().toUpperCase();
  const requestedRow = parseInt(p.rowNumber, 10);
  const originalUsuario = String(p.originalUsuario || '').trim().toLowerCase();

  if (!usuario || !password || !rol) throw new Error('Faltan datos obligatorios del usuario');

  const data = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues() : [];
  let targetRow = Number.isFinite(requestedRow) && requestedRow >= 2 ? requestedRow : 0;

  if (!targetRow) {
    const found = data.findIndex(r => String(r[idx.usuario] || '').trim().toLowerCase() === originalUsuario);
    if (found >= 0) targetRow = found + 2;
  }

  if (!targetRow) {
    const duplicate = data.findIndex(r => String(r[idx.usuario] || '').trim().toLowerCase() === usuario.toLowerCase());
    if (duplicate >= 0) targetRow = duplicate + 2;
  }

  if (!targetRow) targetRow = sh.getLastRow() + 1;

  const row = targetRow <= sh.getLastRow()
    ? sh.getRange(targetRow, 1, 1, lastCol).getValues()[0]
    : new Array(lastCol).fill('');

  row[idx.usuario] = usuario;
  row[idx.password] = password;
  row[idx.rol] = rol;
  if (idx.correo >= 0) row[idx.correo] = correo;
  if (idx.estado >= 0) row[idx.estado] = estado || 'ACTIVO';

  sh.getRange(targetRow, 1, 1, lastCol).setValues([row]);
  return { ok: true, status: 'OK', rowNumber: targetRow };
}

function deleteUsuario_(p) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_USUARIOS);
  if (!sh) throw new Error('No existe la hoja USUARIOS');

  const lastCol = Math.max(sh.getLastColumn(), 5);
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(normalizeHeader_);
  const idxUsuario = findHeaderIndex_(headers, ['USUARIO', 'NOMBRE', 'LOGIN']);
  if (idxUsuario < 0) throw new Error('No se encontro la columna de usuario');

  const requestedRow = parseInt(p.rowNumber, 10);
  const usuario = String(p.usuario || '').trim().toLowerCase();

  let targetRow = Number.isFinite(requestedRow) && requestedRow >= 2 ? requestedRow : 0;
  if (!targetRow) {
    const data = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues() : [];
    const found = data.findIndex(r => String(r[idxUsuario] || '').trim().toLowerCase() === usuario);
    if (found >= 0) targetRow = found + 2;
  }

  if (!targetRow || targetRow > sh.getLastRow()) throw new Error('No se encontro el usuario a eliminar');

  sh.deleteRow(targetRow);
  return { ok: true, status: 'OK', rowNumber: targetRow };
}

function findHeaderIndex_(headers, aliases) {
  for (var i = 0; i < aliases.length; i++) {
    var idx = headers.indexOf(aliases[i]);
    if (idx >= 0) return idx;
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

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
