/****************************************************
 * TTC Monitor Center - home.js (completo)
 * Fuente de datos: Google Sheets (API v4, solo lectura)
 * - Lee EVENTOS, USUARIOS y "Prioridad Eventos"
 * - Home (alertas) se oculta al ir a Usuarios/Dashboard y vuelve sticky al regresar
 * - Prioridad SOLO basada en la pestaña "Prioridad Eventos" (semáforo)
 * - Modal de detalle con acciones de guardar/atender (requiere WebApp paso 3)
 * - Mapa Leaflet dentro del modal (lee coords desde URL de Google Maps)
 * - Guardar comentario convierte estatus a "Seguimiento" si estaba pendiente
 * - Filtro por tarjetas (Todos / Pendientes / Seguimiento) con toggle
 * - Efecto ripple + accesibilidad por teclado en tarjetas
 * - Dirección en modal (geocodificación inversa con Nominatim)
 * - Botón WhatsApp: número fijo + emojis compatibles (api.whatsapp.com)
 * - Orden por fecha: MÁS RECIENTES ARRIBA al cargar/recargar
 ****************************************************/

/***************  CONFIGURACIÓN  ***************/
const SHEET_ID = '1QOuA10BbJaYctYyY4959TwaPGXRG-pzeHPFZmnqKpxI';
const API_KEY  = 'AIzaSyBZGnqBc7zMxADfA19HADzYLtv2F0AjMcs';

// Rangos (ajusta nombres de pestañas si cambian)
const RANGE_EVENTOS    = 'EVENTOS!A1:Z';
const RANGE_USUARIOS   = 'USUARIOS!A1:Z';
const RANGE_PRIORIDAD  = 'Prioridad Eventos!A1:B';   // Evento | Prioridad

// === ESCRITURA (Paso 3: Apps Script) ===
const WEBAPP_UPDATE_URL = 'https://script.google.com/macros/s/AKfycbz96d64MLf7MJPVd4CcYXOemhBJd7U4HIXB86CaVxo0E-Q9UgCASNIDHddjTTpdu0u6/exec';
const WEBAPP_USERS_URL  = WEBAPP_UPDATE_URL;
const UPDATE_SECRET     = 'b3f9d1c2-8a5e-4b6b-9412-1e0f5f1f2c93'; // = SECRET_TOKEN en Apps Script

// === WhatsApp (E.164 SIN "+" ni espacios) ===
const WHATS_NUMBER = '+525645032901';
const WHATS_PREFIX = '🔔 *MolTech* informa que tenemos una:';

// Zona horaria MX para timestamps de comentarios
const TIMEZONE_MX = 'America/Mexico_City';

/***************  ESTADO  ***************/
let PRIORITY_MAP = new Map(); // key: evento normalizado, value: 'alta' | 'media' | 'baja'
let EVENTS = [];              // lista en memoria para abrir detalle
let CURRENT = null;           // evento abierto en modal + metadatos de edición protegida
let CURRENT_FILTER = 'TODOS'; // 'TODOS' | 'PENDIENTES' | 'SEGUIMIENTO'
let USERS = [];               // usuarios en memoria para administrar
let USERS_META = { headers: [], rawHeaders: [], fields: {} };

// Cache de direcciones (lat,lng => texto)
const ADDRESS_CACHE = new Map();

// ======= (a) MAPA (Leaflet) =======
let MAP = { map:null, marker:null };

/** Extrae lat/lng desde una URL de Google Maps. */
function getLatLngFromUrl(url){
  if(!url) return null;
  try{
    const u = new URL(url);
    const q = u.searchParams.get('q');
    if (q) {
      const [latS, lngS] = q.split(',').map(s=>s.trim());
      const lat = parseFloat(latS), lng = parseFloat(lngS);
      if (!Number.isNaN(lat) && !Number.isNaN(lng)) return {lat, lng};
    }
  }catch(e){ /* probar patrón @lat,lng */ }
  const m = (url||'').match(/@(-?\d+(?:\.\d+)?),\s*(-?\d+(?:\.\d+)?)/);
  if (m) return {lat: parseFloat(m[1]), lng: parseFloat(m[2])};
  return null;
}

/** Renderiza/actualiza el mapa dentro del modal. Si no hay coords, oculta el contenedor. */
function renderMapFor(ev){
  const wrap = document.getElementById('detMapWrap');
  const el   = document.getElementById('detMap');
  if(!wrap || !el) return;

  const pos = getLatLngFromUrl(ev.mapa);
  if(!pos){ wrap.style.display='none'; return; }
  wrap.style.display='';

  if(!MAP.map){
    MAP.map = L.map(el,{zoomControl:true});
    L.tileLayer('https://{s}.basemaps.cartocdn.com/rastertiles/voyager/{z}/{x}/{y}{r}.png', {
      maxZoom: 20,
      attribution: '&copy; <a href="https://carto.com/">CARTO</a>'
    }).addTo(MAP.map);
    MAP.marker = L.marker(pos).addTo(MAP.map);
  }else{
    MAP.marker.setLatLng(pos);
  }
  MAP.map.setView(pos, 16);
  setTimeout(()=> MAP.map && MAP.map.invalidateSize(), 150);
}

/** Desmonta el mapa para liberar memoria al cerrar el modal. */
function destroyMap(){
  if(MAP.map){
    MAP.map.remove();
    MAP.map = null; MAP.marker = null;
  }
}

/***************  UTILIDADES  ***************/
const $ = (sel) => document.querySelector(sel);
function setText(id, val){ const el = document.getElementById(id); if(el) el.textContent = String(val); }
function escapeHtml(v){
  return String(v ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/* ==== Timestamps (zona MX) ==== */
function nowMx(){
  return new Date(new Date().toLocaleString('en-US', { timeZone: TIMEZONE_MX }));
}
function pad2(n){ return String(n).padStart(2,'0'); }
function formatTsMx(d){
  const y = d.getFullYear();
  const m = pad2(d.getMonth()+1);
  const day = pad2(d.getDate());
  const hh = pad2(d.getHours());
  const mm = pad2(d.getMinutes());
  const ss = pad2(d.getSeconds());
  return `${y}-${m}-${day} ${hh}:${mm}:${ss}`;
}

/** Convierte serial de Excel/Sheets a Date */
function excelSerialToDate(n){
  const v = parseFloat(n);
  if (Number.isNaN(v)) return null;
  // Excel base 1899-12-30
  return new Date(Math.round((v - 25569) * 86400 * 1000));
}
/** Devuelve ms de una fecha en múltiples formatos (ES/ISO/serial). */
function parseDateSmartToMs(input){
  if (input == null) return 0;

  if (input instanceof Date) return input.getTime();
  if (typeof input === 'number') {
    const d = excelSerialToDate(input);
    return d ? d.getTime() : 0;
  }

  let s = String(input).trim();
  if (!s) return 0;

  // Nativo (ISO u otros entendibles)
  const tryNative = new Date(s);
  if (!Number.isNaN(tryNative.getTime())) return tryNative.getTime();

  // Normalizar separadores/comas y am/pm en español
  s = s.replace(/,/g, ' ').replace(/\s+/g, ' ').trim();

  // dd/mm/yyyy [hh:mm[:ss]] [am/pm | a. m. / p. m.]
  const re = /(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(a\.?\s?m\.?|p\.?\s?m\.?|am|pm)?)?/i;
  const m = s.match(re);
  if (m){
    let [, d, mo, y, hh='0', mm='0', ss='0', ap=''] = m;
    d  = parseInt(d,10);
    mo = parseInt(mo,10);
    y  = parseInt(y,10); if (y < 100) y += 2000;

    let H = parseInt(hh,10), M = parseInt(mm,10), S = parseInt(ss,10);
    if (ap){
      ap = ap.toLowerCase();
      const isPM = ap.startsWith('p');
      const isAM = ap.startsWith('a');
      if (isPM && H < 12) H += 12;
      if (isAM && H === 12) H = 0;
    }
    const dt = new Date(y, mo-1, d, H, M, S);
    if (!Number.isNaN(dt.getTime())) return dt.getTime();
  }

  // Serial de Excel en string
  if (/^\d+(\.\d+)?$/.test(s)){
    const d = excelSerialToDate(s);
    return d ? d.getTime() : 0;
  }

  return 0;
}

/***************  PRIORIDADES (desde hoja)  ***************/
async function cargarMapaPrioridad(){
  try{
    const values = await fetchSheet(RANGE_PRIORIDAD);
    const rows = sheetToObjects(values); // espera headers: EVENTO | PRIORIDAD
    PRIORITY_MAP = new Map();
    rows.forEach(r=>{
      const ev = normVal(r.EVENTO || '');
      const pr = (r.PRIORIDAD || '').toLowerCase();
      if(ev && /^(alta|media|baja)$/.test(pr)) PRIORITY_MAP.set(ev, pr);
    });
  }catch(err){
    console.warn('No se pudo cargar "Prioridad Eventos":', err);
    PRIORITY_MAP = new Map(); // sin mapa; default en mapEvento => 'baja'
  }
}

/***************  MAPEO Y RENDER  ***************/
function norm(s){
  if(!s) return '';
  const map = {'Á':'A','É':'E','Í':'I','Ó':'O','Ú':'U','Ä':'A','Ë':'E','Ï':'I','Ö':'O','Ü':'U','Ñ':'N'};
  return s.trim().toUpperCase().replace(/[ÁÉÍÓÚÄËÏÖÜÑ]/g, m=>map[m]||m).replace(/[^\w]+/g,'_');
}
function normVal(s){
  if(!s) return '';
  const map = {'á':'a','é':'e','í':'i','ó':'o','ú':'u','ä':'a','ë':'e','ï':'i','ö':'o','ü':'u','ñ':'n',
               'Á':'A','É':'E','Í':'I','Ó':'O','Ú':'U','Ä':'A','Ë':'E','Ï':'I','Ö':'O','Ü':'U','Ñ':'N'};
  return s.trim().replace(/[ÁáÉéÍíÓóÚúÄäËëÏïÖöÜüÑñ]/g, m=>map[m]||m)
          .replace(/\s+/g,' ').toUpperCase();
}
function sheetToObjects(values){
  if(!values || !values.length) return [];
  const headers = values[0].map(norm);
  return values.slice(1).map(row=>{
    const o = {};
    headers.forEach((h,i)=>{ o[h] = (row[i] ?? '').toString().trim(); });
    return o;
  });
}
async function fetchSheet(rangeA1){
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}/values/${encodeURIComponent(rangeA1)}?key=${API_KEY}`;
  const res = await fetch(url, { headers: {'Cache-Control':'no-cache'} });
  if(!res.ok) throw new Error(`Sheets error ${res.status}: ${await res.text()}`);
  const j = await res.json();
  return j.values || [];
}

function mapEvento(o){
  const eventoTxt = o.EVENTO || o.DESCRIPCION || o.NOTA || '';
  const fromMap   = PRIORITY_MAP.get(normVal(eventoTxt));
  const prioridad = (fromMap || 'baja').toLowerCase();

  const generado    = o.GENERADO || o.FECHA || o.MSG_TIME || o.POS_TIME || '';
  const unidad      = o.UNIDAD    || o.DEVICE_ID || o.ID_UNIDAD || '';
  const estatusRaw  = o.ESTATUS || o.SEGUIMIENTO || o.ESTADO || o.STATUS || '';
  const folio       = o.FOLIO || o.ID || '';
  const mapa        = o.MAPA || o.MAP || o.URL || '';
  const comentarios = o.COMENTARIOS || o.COMENTARIO || '';

  return {
    folio,
    prioridad,
    generado,
    descripcion: eventoTxt,   // se muestra como "Evento"
    unidad,
    mapa,
    estatus: estatusRaw || 'Sin atender',
    comentarios
  };
}

/** Semáforo prioridad */
function iconPrioridad(p){
  let color = '#43a047'; // baja
  if (p === 'media') color = '#ffb300';   // amarillo
  if (p === 'alta')  color = '#e53935';   // rojo
  const title = p ? (p[0].toUpperCase()+p.slice(1)) : '';
  return `<i class="fa-solid fa-circle" style="color:${color}" title="${title}"></i>`;
}

/** Colores por estatus */
function statusPillColor(status){
  const s = normalizeStatus(status);
  if (s.includes('SEGUIM')) return '#FFD700';                 // Seguimiento: amarillo/dorado
  if (s === 'ATENDIDO' || s === 'CERRADO') return '#2e7d32';  // Atendido: verde
  return '#777';                                              // Pendiente: gris
}
function normalizeStatus(sRaw){
  return (sRaw||'').toString().trim().toUpperCase();
}
function isPendingStatus(sRaw){
  const s = normalizeStatus(sRaw);
  return (!s) || s==='PENDIENTE' || s==='SIN ATENDER' || s==='ABIERTA' || s==='OPEN';
}

function renderEventos(eventos){
  const tbody = document.getElementById('tbodyEventos');
  if(!tbody) return;
  tbody.innerHTML = '';
  eventos.forEach(ev=>{
    const tr = document.createElement('tr');
    const eventoTxt = (ev.descripcion || '');
    const pillColor = statusPillColor(ev.estatus);
    tr.innerHTML = `
      <td>${iconPrioridad(ev.prioridad)}</td>
      <td>${ev.generado ?? ''}</td>
      <td title="${eventoTxt}">${eventoTxt.slice(0,40)}${eventoTxt.length>40?'…':''}</td>
      <td>${ev.unidad ?? ''}</td>
      <td>
        <span class="pill" style="display:inline-block;background:${pillColor};color:#fff;padding:.2rem .6rem;border-radius:999px;font-size:.8rem">
          ${ev.estatus ?? 'Sin atender'}
        </span>
      </td>
    `;
    tr.style.cursor = 'pointer';
    tr.addEventListener('click', ()=> openDetalle(ev));
    tbody.appendChild(tr);
  });
}

function computeCounters(eventos){
  const todos = eventos.length;
  const pendientes = eventos.filter(e=> isPendingStatus(e.estatus) ).length;
  const enSeg = eventos.filter(e=>{
    const s = normalizeStatus(e.estatus);
    return s.includes('SEGUIM') || s === 'EN PROCESO' || s === 'ATENDIENDO';
  }).length;
  return {todos, pendientes, seguimiento: enSeg};
}

/***************  DIRECCIÓN (reverse geocode) ***************/
async function reverseGeocode(lat, lng){
  const key = `${lat.toFixed(6)},${lng.toFixed(6)}`;
  if (ADDRESS_CACHE.has(key)) return ADDRESS_CACHE.get(key);

  const url = `https://nominatim.openstreetmap.org/reverse?lat=${encodeURIComponent(lat)}&lon=${encodeURIComponent(lng)}&format=jsonv2&zoom=18&addressdetails=1&accept-language=es`;
  const res = await fetch(url, { headers: { 'Accept':'application/json' } });
  if(!res.ok) throw new Error(`Nominatim ${res.status}`);
  const j = await res.json();

  let texto = j.display_name || '';
  if (!texto && j.address){
    const a = j.address;
    texto = [
      a.road || a.pedestrian || a.cycleway || a.path,
      a.neighbourhood || a.suburb || a.village || a.town || a.city,
      a.state,
      a.postcode,
      a.country
    ].filter(Boolean).join(', ');
  }
  if (!texto) texto = 'Dirección no disponible';

  ADDRESS_CACHE.set(key, texto);
  return texto;
}

function setDireccionText(t){
  const el = document.getElementById('detDireccion');
  if (el) el.textContent = t || '—';
}

async function renderDireccion(ev){
  const dirEl = document.getElementById('detDireccion');
  if (!dirEl) return;
  dirEl.textContent = 'Buscando…';

  const pos = getLatLngFromUrl(ev.mapa);
  if (!pos){ dirEl.textContent = 'No disponible'; return; }

  try{
    const texto = await reverseGeocode(pos.lat, pos.lng);
    dirEl.textContent = texto;
  }catch(err){
    console.warn('Reverse geocode error', err);
    dirEl.textContent = 'No disponible';
  }
}

/***************  FILTRO POR TARJETAS  ***************/
function getFilteredEvents(){
  if (!Array.isArray(EVENTS)) return [];
  switch (CURRENT_FILTER){
    case 'PENDIENTES':
      return EVENTS.filter(e=> isPendingStatus(e.estatus));
    case 'SEGUIMIENTO':
      return EVENTS.filter(e=>{
        const s = normalizeStatus(e.estatus);
        return s.includes('SEGUIM') || s==='EN PROCESO' || s==='ATENDIENDO';
      });
    case 'TODOS':
    default:
      return EVENTS.slice();
  }
}

/* Ripple helper (onda al click) */
function attachRipple(el){
  if (!el) return;
  el.addEventListener('click', function(e){
    const rect = el.getBoundingClientRect();
    const size = Math.max(rect.width, rect.height);
    const cx = rect.width/2, cy = rect.height/2;
    const x = (e.clientX ? e.clientX - rect.left : cx) - size/2;
    const y = (e.clientY ? e.clientY - rect.top  : cy) - size/2;

    const span = document.createElement('span');
    span.className = 'ripple';
    span.style.width = span.style.height = `${size}px`;
    span.style.left = `${x}px`;
    span.style.top  = `${y}px`;

    el.querySelectorAll('.ripple').forEach(n=>n.remove());
    el.appendChild(span);
    setTimeout(()=> span.remove(), 650);
  });
}

function updateCardUI(){
  const map = {
    'TODOS': document.getElementById('cardTodos'),
    'PENDIENTES': document.getElementById('cardPendientes'),
    'SEGUIMIENTO': document.getElementById('cardSeguimiento'),
  };
  Object.entries(map).forEach(([key, el])=>{
    if(!el) return;
    const isActive = (CURRENT_FILTER === key);
    el.classList.toggle('active', isActive);
    el.setAttribute('aria-pressed', String(isActive));
    if (!isActive){ el.style.outline=''; el.style.boxShadow=''; }
  });
}

/* Toggle: si clicas la misma tarjeta, vuelve a 'TODOS' */
function setFilter(kind){
  const next = (kind || 'TODOS').toUpperCase();
  CURRENT_FILTER = (CURRENT_FILTER === next) ? 'TODOS' : next;
  updateCardUI();
  renderEventos(getFilteredEvents());
}

/***************  GUARD de sesión  ***************/
(function guard(){
  const u = sessionStorage.getItem('USER');
  if(!u){ window.location.href = 'index.html'; return; }
  const chip = document.getElementById('chipUser');
  if (chip) chip.textContent = u;
})();
document.getElementById('btnLogout')?.addEventListener('click', ()=>{
  sessionStorage.clear(); window.location.href = 'index.html';
});

/***************  NAVEGACIÓN (Home/Reportes/Usuarios)  ***************/
const linkInicio    = document.getElementById('linkInicio');
const linkDashboard = document.getElementById('linkDashboard');
const linkUsuarios  = document.getElementById('linkUsuarios');

const secInicio     = document.getElementById('secInicio');
const secDashboard  = document.getElementById('secDashboard');
const secUsuarios   = document.getElementById('secUsuarios');

// Bloques del Home para controlar sticky y scroll
const toolbarAlerts = document.getElementById('toolbarAlerts') || document.querySelector('#secInicio .top-toolbar');
const tablaScroll   = document.querySelector('#secInicio .tabla-scroll');

function activate(link){
  [linkInicio, linkDashboard, linkUsuarios].forEach(a=>a?.classList.remove('active'));
  link?.classList.add('active');
}

/* Ocultamiento/Mostrado “duro” de Home */
function hideHomeHard(){
  if (toolbarAlerts) toolbarAlerts.classList.remove('sticky');
  if (tablaScroll)   tablaScroll.scrollTop = 0;
  secInicio?.classList.add('oculto');
  if (secInicio) secInicio.style.display = 'none';
  const tbody = document.getElementById('tbodyEventos');
  if (tbody) tbody.innerHTML = '';
}
function showHomeHard(){
  if (secInicio){
    secInicio.style.display = '';
    secInicio.classList.remove('oculto');
  }
  if (toolbarAlerts) toolbarAlerts.classList.add('sticky');
  if (tablaScroll)   tablaScroll.scrollTop = 0;
}

function showOnly(section){
  [secInicio, secDashboard, secUsuarios].forEach(s=>{
    if(!s) return;
    s.classList.add('oculto');
    s.style.display = 'none';
  });
  if (section){
    section.classList.remove('oculto');
    section.style.display = '';
  }
  if (section === secInicio) showHomeHard(); else hideHomeHard();
}

function go(link, section){
  activate(link);
  showOnly(section);
  if (section === secInicio) renderEventos(getFilteredEvents());
  if (section === secUsuarios) cargarUsuarios().catch(console.error);
  if (section === secDashboard) cargarDashboardModule();
}

/* Listeners navegación */
function bindNav(a, targetSection){
  if(!a) return;
  a.addEventListener('click', e=>{
    e.preventDefault();
    e.stopImmediatePropagation();
    go(a, targetSection);
  });
}
bindNav(linkInicio,   secInicio);
bindNav(linkDashboard, secDashboard);
bindNav(linkUsuarios, secUsuarios);

/***************  Búsqueda en tabla EVENTOS  ***************/
(function bindSearch(){
  const input = document.getElementById('busquedaEventos');
  if (!input) return;
  input.addEventListener('input', function(){
    const filtro = this.value.toLowerCase();
    document.querySelectorAll('#tablaEventos tbody tr').forEach(tr=>{
      tr.style.display = tr.textContent.toLowerCase().includes(filtro) ? '' : 'none';
    });
  });
})();

/***************  MODAL DETALLE ***************/
function showModal(show=true){
  const bg = document.getElementById('modalBg');
  const m  = document.getElementById('modalCaso');
  if(!bg || !m) return;
  bg.classList.toggle('oculto', !show);
  m.classList.toggle('oculto',  !show);
}

/* --- BLOQUEO DE HISTORIAL Y PREFIJO AUTO --- */
function seedLockedPrefix(base, user){
  const sep = base && !base.endsWith('\n') ? '\n' : '';
  const prefix = `[${formatTsMx(nowMx())}] ${user ? (user + ':') : ''} `;
  const locked = (base || '') + (base ? sep : '') + prefix;
  return { locked, editStart: locked.length };
}
function clampCaretToEditStart(ta){
  if (!CURRENT) return;
  const guard = CURRENT.editStartIdx ?? 0;
  const s = ta.selectionStart ?? 0;
  const e = ta.selectionEnd ?? 0;
  if (s < guard || e < guard){
    ta.setSelectionRange(guard, guard);
  }
}

/* Handlers que protegen el prefijo e historial */
function attachGuardHandlers(ta){
  // keydown: evita backspace/delete/home/edición antes del guard
  ta.onkeydown = (e)=>{
    if (!CURRENT) return;
    const guard = CURRENT.editStartIdx ?? 0;
    const s = ta.selectionStart ?? 0;
    const en = ta.selectionEnd ?? 0;

    // Home -> al inicio de la zona editable
    if (e.key === 'Home'){
      e.preventDefault();
      ta.setSelectionRange(guard, guard);
      return;
    }

    // Backspace/Delete en zona protegida
    if ((e.key === 'Backspace' && s <= guard && en <= guard) ||
        (e.key === 'Delete' && s < guard)){
      e.preventDefault();
      ta.setSelectionRange(guard, guard);
      return;
    }

    // Cualquier edición con cursor antes del guard
    const editingKey = (e.key?.length === 1) || ['Enter','Tab'].includes(e.key);
    if (editingKey && s < guard){
      e.preventDefault();
      ta.setSelectionRange(guard, guard);
      return;
    }
  };

  // input: si alguien logró modificar el prefijo, lo reponemos
  ta.oninput = ()=>{
    if (!CURRENT) return;
    const prefix = CURRENT.lockPrefix || '';
    if (!ta.value.startsWith(prefix)){
      // Conserva solo lo que esté después del prefijo original
      const rest = ta.value.slice(prefix.length);
      ta.value = prefix + rest;
    }
    clampCaretToEditStart(ta);
  };

  // click / mouseup / keyup: fuerza caret al mínimo permitido
  const fixCaret = ()=> clampCaretToEditStart(ta);
  ta.onclick = fixCaret;
  ta.onmouseup = fixCaret;
  ta.onkeyup = fixCaret;
}

function openDetalle(ev){
  CURRENT = ev;

  document.getElementById('detTitulo').textContent  = ev.descripcion || '—';
  document.getElementById('detFecha').textContent   = ev.generado || '—';
  document.getElementById('detUnidad').textContent  = ev.unidad || '—';
  const est = document.getElementById('detEstatus');
  est.textContent = ev.estatus || 'Sin atender';

  const aMapa = document.getElementById('detMapa');
  if (aMapa) {
    if (ev.mapa) { aMapa.href = ev.mapa; aMapa.style.pointerEvents = 'auto'; aMapa.style.opacity = 1; }
    else         { aMapa.href = '#';      aMapa.style.pointerEvents = 'none'; aMapa.style.opacity = .5; }
  }

  const ta = document.getElementById('detComentario');
  if (ta){
    const base = ev.comentarios || '';
    const user = sessionStorage.getItem('USER') || '';
    const { locked, editStart } = seedLockedPrefix(base, user);

    ta.value = locked;          // historial + timestamp/usuario listo
    ta.placeholder = '';        // sin leyendas
    CURRENT.lockPrefix   = locked;     // prefijo bloqueado
    CURRENT.editStartIdx = editStart;  // desde aquí se puede escribir
    CURRENT.baseComments = base;       // historial “puro” previo

    // coloca el cursor al final (inicio zona editable) y adjunta guardas
    setTimeout(()=>{
      try{ ta.setSelectionRange(editStart, editStart); }catch(_){}
      attachGuardHandlers(ta);
    }, 0);
  }

  // Mapa + dirección
  renderMapFor(ev);
  renderDireccion(ev);

  showModal(true);
}

function closeDetalle(){
  destroyMap();
  showModal(false);
  CURRENT = null;
}

document.getElementById('btnCloseModal')?.addEventListener('click', closeDetalle);
document.getElementById('modalBg')?.addEventListener('click', closeDetalle);
document.addEventListener('keydown', (e)=>{ if(e.key==='Escape') closeDetalle(); });

async function actualizarCaso({ folio, estatus, comentario }) {
  if (!WEBAPP_UPDATE_URL) { alert('Configura WEBAPP_UPDATE_URL'); return; }
  if (!folio) { alert('Folio no encontrado.'); return; }

  const form = new URLSearchParams();
  form.set('action', 'update');
  form.set('token', UPDATE_SECRET);
  form.set('folio', folio);
  form.set('estatus', estatus || '');
  form.set('comentario', comentario || '');

  const btnG = document.getElementById('btnGuardar');
  const btnA = document.getElementById('btnAtendido');
  [btnG, btnA].forEach(b=>{ if(b){ b.disabled = true; b.style.opacity = .6; }});

  try{
    const res = await fetch(WEBAPP_UPDATE_URL, { method:'POST', body: form });
    const txt = await res.text();
    let j; try { j = JSON.parse(txt); } catch { throw new Error('Respuesta no JSON del WebApp: ' + txt.slice(0,200)); }
    const success = (j && (j.ok === true || j.status === 'OK'));
    if (!success) throw new Error(j.error || j.message || 'Error al actualizar');

    await cargarDashboard();
    renderEventos(getFilteredEvents());
    closeDetalle();
  }catch(err){
    console.error(err);
    alert('No se pudo guardar: ' + err.message);
  }finally{
    [btnG, btnA].forEach(b=>{ if(b){ b.disabled = false; b.style.opacity = 1; }});
  }
}

// Botones del modal
document.getElementById('btnGuardar')?.addEventListener('click', ()=>{
  if(!CURRENT) return;
  const ta = document.getElementById('detComentario');
  const guard = CURRENT.editStartIdx ?? 0;
  const newPart = (ta.value || '').slice(guard).trim();

  // Si no escribió nada después del prefijo, no guardamos
  if (!newPart){
    alert('Escribe tu comentario después del timestamp y usuario.');
    try{ ta.setSelectionRange(guard, guard); }catch(_){}
    ta.focus();
    return;
  }

  // El contenido completo del textarea YA es el hilo correcto (historial + timestamp/usuario + texto)
  const thread = (ta.value || '').trimEnd();

  let nuevoEstatus = CURRENT.estatus;
  if (newPart && isPendingStatus(CURRENT.estatus)) {
    nuevoEstatus = 'Seguimiento';
  }

  // Mantén el historial local
  CURRENT.comentarios = thread;

  actualizarCaso({ folio: CURRENT.folio, estatus: nuevoEstatus, comentario: thread });
});

document.getElementById('btnAtendido')?.addEventListener('click', ()=>{
  if(!CURRENT) return;
  const ta = document.getElementById('detComentario');
  const guard = CURRENT.editStartIdx ?? 0;
  const newPart = (ta.value || '').slice(guard).trim();

  // Si no escribió nada, de todos modos cerramos como Atendido con solo el prefijo
  const thread = (ta.value || '').trimEnd();
  CURRENT.comentarios = thread;

  actualizarCaso({ folio: CURRENT.folio, estatus: 'Atendido', comentario: thread });
});

// guardar con Ctrl/Cmd+Enter
document.getElementById('detComentario')?.addEventListener('keydown', (e)=>{
  if ((e.ctrlKey || e.metaKey) && e.key === 'Enter'){
    e.preventDefault();
    document.getElementById('btnGuardar')?.click();
  }
});

// Copiar dirección
document.getElementById('btnCopyDir')?.addEventListener('click', ()=>{
  const t = document.getElementById('detDireccion')?.textContent?.trim();
  if (!t || t === '—' || t === 'Buscando…') return;
  navigator.clipboard?.writeText(t).catch(()=>{});
});

/***************  WHATSAPP: construir y abrir mensaje ***************/
function buildWhatsMessage(ev){
  const titulo = (ev?.descripcion || 'Alerta').replace(/\s+/g,' ').trim();
  const unidad = (ev?.unidad || '—').toString().trim();

  // Dirección actual mostrada en el modal (o link si aún no se resuelve)
  let direccion = (document.getElementById('detDireccion')?.textContent || '').trim();
  if (!direccion || direccion === '—' || /No disponible|Buscando…/i.test(direccion)){
    const pos = getLatLngFromUrl(ev?.mapa || '');
    direccion = pos ? `https://maps.google.com/?q=${pos.lat},${pos.lng}` : 'No disponible';
  }

  const iconNotif = '⚠️';
  const iconUnidad = '🚚';
  const iconPinTxt = '📌';
  const iconMapa   = '🌐';

  const lineaDireccion = /^https?:\/\//i.test(direccion)
    ? `${iconMapa} *Mapa:* ${direccion}`
    : `${iconPinTxt} *Dirección:* "${direccion}"`;

  return `${WHATS_PREFIX}
${iconNotif} *Notificación:* "${titulo}"
${iconUnidad} *Unidad:* "${unidad}"
${lineaDireccion}`;
}

// Usar api.whatsapp.com
function openWhatsToNumber(phoneE164, text){
  const url = `https://api.whatsapp.com/send?phone=${phoneE164}&text=${encodeURIComponent(text)}`;
  window.open(url, '_blank');
}

document.getElementById('btnWhats')?.addEventListener('click', ()=>{
  const msg = buildWhatsMessage(CURRENT);
  openWhatsToNumber(WHATS_NUMBER, msg);
});

/***************  CARGAS ***************/
async function cargarDashboard(){
  try{
    const values = await fetchSheet(RANGE_EVENTOS);
    const objetos = sheetToObjects(values);
    EVENTS = objetos.map(mapEvento);

    /* === ORDENAR POR FECHA (más recientes primero) === */
    EVENTS.sort((a, b) => parseDateSmartToMs(b.generado) - parseDateSmartToMs(a.generado));

    const k = computeCounters(EVENTS);
    setText('nTodos',       k.todos);
    setText('nPendientes',  k.pendientes);
    setText('nSeguimiento', k.seguimiento);

    // Render según filtro activo
    renderEventos(getFilteredEvents());
  }catch(err){
    console.error('Dashboard error:', err);
    setText('nTodos',0); setText('nPendientes',0); setText('nSeguimiento',0);
    const tbody = document.getElementById('tbodyEventos');
    if (tbody){
      const cols = document.querySelectorAll('#tablaEventos thead th').length || 5;
      tbody.innerHTML = `<tr><td colspan="${cols}" style="text-align:center;color:#a00">No se pudieron cargar las alertas.</td></tr>`;
    }
  }
}

async function cargarUsuarios(){
  const tbody = document.getElementById('tbodyUsuarios');
  if (!tbody) return;
  tbody.innerHTML = '';
  try{
    const values = await fetchSheet(RANGE_USUARIOS);
    const objs = sheetToObjects(values);

    const rows = objs.map(o=>({
      usuario: o.USUARIO || o.NOMBRE || o.LOGIN || '',
      rol:     o.ROL || o.PERFIL || '',
      correo:  o.CORREO || o.EMAIL || '',
      estado:  (o.ESTADO || o.STATUS || '').toString().toUpperCase()
    }));

    if(!rows.length){
      tbody.innerHTML = `<tr><td colspan="4" style="text-align:center;color:#666">Sin usuarios para mostrar.</td></tr>`;
      return;
    }

    rows.forEach(u=>{
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td>${u.usuario || '—'}</td>
        <td>${u.rol || '—'}</td>
        <td>${u.correo || '—'}</td>
        <td>${u.estado || '—'}</td>
      `;
      tbody.appendChild(tr);
    });
  }catch(err){
    console.error('Usuarios error:', err);
    tbody.innerHTML = `<tr><td colspan="4" style="text-align:center;color:#a00">Error al cargar usuarios.</td></tr>`;
  }
}

(function bindUsersUI(){
  document.getElementById('btnNuevoUsuario')?.addEventListener('click', ()=> openUserModal('create'));
  document.getElementById('btnRecargarUsuarios')?.addEventListener('click', ()=> cargarUsuarios().catch(console.error));
  document.getElementById('usersSearch')?.addEventListener('input', renderUsuariosTable);
  document.getElementById('btnCloseUserModal')?.addEventListener('click', closeUserModal);
  document.getElementById('btnCancelarUserModal')?.addEventListener('click', closeUserModal);
  document.getElementById('userModalBg')?.addEventListener('click', closeUserModal);
  document.getElementById('userForm')?.addEventListener('submit', onSubmitUserForm);
  document.addEventListener('keydown', (e)=>{ if (e.key === 'Escape') closeUserModal(); });
  document.getElementById('tbodyUsuarios')?.addEventListener('click', (e)=>{
    const btnEdit = e.target.closest('[data-user-edit]');
    if (btnEdit){
      const user = findUserByName(btnEdit.getAttribute('data-user-edit'));
      if (user) openUserModal('edit', user);
      return;
    }
    const btnDelete = e.target.closest('[data-user-delete]');
    if (btnDelete){
      onDeleteUsuario(btnDelete.getAttribute('data-user-delete'));
    }
  });
  syncUsersAdminUI();
})();

function isUsersAdmin(){
  return (sessionStorage.getItem('ROLE') || '').trim().toLowerCase() === 'admin';
}

function resolveUserFields(headers){
  const pick = (...aliases)=> aliases.find(a=> headers.includes(a)) || '';
  return {
    usuario: pick('USUARIO', 'NOMBRE', 'LOGIN'),
    password: pick('PASSWORD', 'CONTRASENA', 'CLAVE', 'PASS'),
    rol: pick('ROL', 'PERFIL'),
    correo: pick('CORREO', 'EMAIL'),
    estado: pick('ESTADO', 'STATUS')
  };
}

function valueByAlias(headers, row, aliases){
  for (const alias of aliases){
    const idx = headers.indexOf(alias);
    if (idx >= 0) return (row[idx] ?? '').toString().trim();
  }
  return '';
}

function buildUserRow(headers, row, idx){
  return {
    rowNumber: idx + 2,
    usuario: valueByAlias(headers, row, ['USUARIO', 'NOMBRE', 'LOGIN']),
    password: valueByAlias(headers, row, ['PASSWORD', 'CONTRASENA', 'CLAVE', 'PASS']),
    rol: valueByAlias(headers, row, ['ROL', 'PERFIL']),
    correo: valueByAlias(headers, row, ['CORREO', 'EMAIL']),
    estado: valueByAlias(headers, row, ['ESTADO', 'STATUS']).toUpperCase(),
  };
}

function getFilteredUsers(){
  const q = (document.getElementById('usersSearch')?.value || '').trim().toLowerCase();
  if (!q) return USERS.slice();
  return USERS.filter(u => [u.usuario, u.rol, u.correo, u.estado].some(v => String(v || '').toLowerCase().includes(q)));
}

function setUsersStatus(text){
  const el = document.getElementById('usersStatus');
  if (el) el.textContent = text;
}

function syncUsersAdminUI(){
  const showAdmin = isUsersAdmin();
  document.querySelectorAll('.admin-only').forEach(el=>{
    el.classList.toggle('oculto-admin', !showAdmin);
  });
}

function findUserByName(usuario){
  const key = (usuario || '').trim().toLowerCase();
  return USERS.find(u => (u.usuario || '').trim().toLowerCase() === key) || null;
}

function toggleUserModal(show){
  const bg = document.getElementById('userModalBg');
  const modal = document.getElementById('modalUsuario');
  if (!bg || !modal) return;
  bg.classList.toggle('oculto', !show);
  modal.classList.toggle('oculto', !show);
}

function openUserModal(mode, user=null){
  if (!isUsersAdmin()) return;

  const form = document.getElementById('userForm');
  const title = document.getElementById('userModalTitle');
  const btn = document.getElementById('btnGuardarUsuario');
  if (!form || !title || !btn) return;

  form.reset();
  form.dataset.mode = mode || 'create';
  form.dataset.rowNumber = user?.rowNumber ? String(user.rowNumber) : '';
  form.dataset.originalUsuario = user?.usuario || '';

  title.textContent = mode === 'edit' ? 'Editar usuario' : 'Nuevo usuario';
  btn.innerHTML = mode === 'edit'
    ? '<i class="fas fa-floppy-disk"></i> Guardar cambios'
    : '<i class="fas fa-user-plus"></i> Crear usuario';

  document.getElementById('userUsuario').value = user?.usuario || '';
  document.getElementById('userPassword').value = user?.password || '';
  document.getElementById('userRol').value = user?.rol || 'monitorista';
  document.getElementById('userCorreo').value = user?.correo || '';
  document.getElementById('userEstado').value = user?.estado || 'ACTIVO';

  toggleUserModal(true);
  setTimeout(()=> document.getElementById('userUsuario')?.focus(), 0);
}

function closeUserModal(){
  toggleUserModal(false);
}

function renderUsuariosTable(){
  const tbody = document.getElementById('tbodyUsuarios');
  if (!tbody) return;

  syncUsersAdminUI();
  const rows = getFilteredUsers();
  const cols = isUsersAdmin() ? 5 : 4;

  if (!USERS.length){
    tbody.innerHTML = `<tr><td colspan="${cols}" style="text-align:center;color:#666">Sin usuarios para mostrar.</td></tr>`;
    setUsersStatus('No hay usuarios en la hoja.');
    return;
  }

  if (!rows.length){
    tbody.innerHTML = `<tr><td colspan="${cols}" style="text-align:center;color:#666">No hay coincidencias para la busqueda actual.</td></tr>`;
    setUsersStatus(`Mostrando 0 de ${USERS.length} usuarios.`);
    return;
  }

  tbody.innerHTML = '';
  rows.forEach(u=>{
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${escapeHtml(u.usuario || '-')}</td>
      <td>${escapeHtml(u.rol || '-')}</td>
      <td>${escapeHtml(u.correo || '-')}</td>
      <td>${escapeHtml(u.estado || '-')}</td>
      ${isUsersAdmin() ? `
        <td>
          <div class="users-actions">
            <button type="button" class="btn-icon btn-edit" data-user-edit="${escapeHtml(u.usuario)}">Editar</button>
            <button type="button" class="btn-icon btn-delete" data-user-delete="${escapeHtml(u.usuario)}">Eliminar</button>
          </div>
        </td>` : ''}
    `;
    tbody.appendChild(tr);
  });

  setUsersStatus(`Mostrando ${rows.length} de ${USERS.length} usuarios.`);
}

async function postUserAction(action, payload){
  if (!WEBAPP_USERS_URL) throw new Error('Configura WEBAPP_USERS_URL');

  const form = new URLSearchParams();
  form.set('action', action);
  form.set('token', UPDATE_SECRET);
  Object.entries(payload || {}).forEach(([key, value])=>{
    if (value != null) form.set(key, String(value));
  });

  const res = await fetch(WEBAPP_USERS_URL, { method:'POST', body: form });
  const txt = await res.text();
  let j;
  try { j = JSON.parse(txt); }
  catch { throw new Error('Respuesta no JSON del WebApp: ' + txt.slice(0,200)); }

  const success = (j && (j.ok === true || j.status === 'OK'));
  if (!success) throw new Error(j.error || j.message || 'Error al escribir usuarios');
  return j;
}

async function saveUsuario(payload){
  return postUserAction('users_upsert', payload);
}

async function deleteUsuario(payload){
  return postUserAction('users_delete', payload);
}

async function onSubmitUserForm(e){
  e.preventDefault();
  if (!isUsersAdmin()) return;

  const form = e.currentTarget;
  const usuario = document.getElementById('userUsuario')?.value.trim() || '';
  const password = document.getElementById('userPassword')?.value.trim() || '';
  const rol = document.getElementById('userRol')?.value.trim() || '';
  const correo = document.getElementById('userCorreo')?.value.trim() || '';
  const estado = (document.getElementById('userEstado')?.value || '').trim().toUpperCase();

  if (!usuario || !password || !rol || !estado){
    alert('Completa usuario, contrasena, rol y estado.');
    return;
  }

  const btn = document.getElementById('btnGuardarUsuario');
  if (btn){ btn.disabled = true; btn.style.opacity = .65; }

  try{
    await saveUsuario({
      rowNumber: form.dataset.rowNumber || '',
      originalUsuario: form.dataset.originalUsuario || '',
      usuario,
      password,
      rol,
      correo,
      estado
    });
    closeUserModal();
    await cargarUsuarios();
  }catch(err){
    console.error(err);
    alert('No se pudo guardar el usuario: ' + err.message);
  }finally{
    if (btn){ btn.disabled = false; btn.style.opacity = 1; }
  }
}

async function onDeleteUsuario(usuario){
  if (!isUsersAdmin()) return;
  const user = findUserByName(usuario);
  if (!user) return;

  const ok = confirm(`Se eliminara el usuario "${user.usuario}". Esta accion borrara la fila en Google Sheets.`);
  if (!ok) return;

  try{
    await deleteUsuario({
      rowNumber: user.rowNumber,
      usuario: user.usuario
    });
    await cargarUsuarios();
  }catch(err){
    console.error(err);
    alert('No se pudo eliminar el usuario: ' + err.message);
  }
}

async function cargarUsuarios(){
  const tbody = document.getElementById('tbodyUsuarios');
  if (!tbody) return;
  tbody.innerHTML = '';
  try{
    const values = await fetchSheet(RANGE_USUARIOS);
    const headers = (values[0] || []).map(norm);
    const rows = values.slice(1);

    USERS_META = {
      headers,
      rawHeaders: values[0] || [],
      fields: resolveUserFields(headers)
    };

    USERS = rows.map((row, idx)=> buildUserRow(headers, row, idx));
    renderUsuariosTable();
  }catch(err){
    console.error('Usuarios error:', err);
    USERS = [];
    USERS_META = { headers: [], rawHeaders: [], fields: {} };
    tbody.innerHTML = `<tr><td colspan="${isUsersAdmin() ? 5 : 4}" style="text-align:center;color:#a00">Error al cargar usuarios.</td></tr>`;
    setUsersStatus('No se pudieron cargar los usuarios.');
  }
}

/***************  Inicio ***************/
(function bindCards(){
  const cTodos = document.getElementById('cardTodos');
  const cPend  = document.getElementById('cardPendientes');
  const cSeg   = document.getElementById('cardSeguimiento');

  [cTodos, cPend, cSeg].forEach(attachRipple);

  cTodos?.addEventListener('click', ()=> setFilter('TODOS'));
  cPend?.addEventListener('click',  ()=> setFilter('PENDIENTES'));
  cSeg?.addEventListener('click',   ()=> setFilter('SEGUIMIENTO'));

  [cTodos, cPend, cSeg].forEach(el=>{
    el?.setAttribute('tabindex','0');
    el?.setAttribute('role','button');
    el?.addEventListener('keydown', (e)=>{
      if(e.key === 'Enter' || e.key === ' '){
        e.preventDefault();
        el.click();
      }
    });
  });
})();

/***************  DASHBOARD  ***************/
const BAR_COLORS = ['#3b82f6','#ef4444','#f59e0b','#10b981','#8b5cf6','#ec4899','#06b6d4','#f97316','#6366f1','#14b8a6'];

function getReportFiltered(){
  const desde = document.getElementById('rptDesde')?.value;
  const hasta = document.getElementById('rptHasta')?.value;
  if (!EVENTS.length) return [];

  let list = EVENTS.slice();
  if (desde){
    const dMs = new Date(desde + 'T00:00:00').getTime();
    list = list.filter(e => parseDateSmartToMs(e.generado) >= dMs);
  }
  if (hasta){
    const hMs = new Date(hasta + 'T23:59:59').getTime();
    list = list.filter(e => parseDateSmartToMs(e.generado) <= hMs);
  }
  return list;
}

function renderBarChart(containerId, dataMap, maxBars){
  const el = document.getElementById(containerId);
  if (!el) return;
  const entries = [...dataMap.entries()].sort((a,b) => b[1] - a[1]).slice(0, maxBars || 10);
  if (!entries.length){ el.innerHTML = '<div class="rpt-empty">Sin datos para mostrar</div>'; return; }
  const max = entries[0][1] || 1;
  el.innerHTML = entries.map(([label, count], i) => {
    const pct = Math.max((count / max) * 100, 8);
    const color = BAR_COLORS[i % BAR_COLORS.length];
    return `<div class="rpt-bar-row">
      <span class="rpt-bar-label" title="${label}">${label}</span>
      <div class="rpt-bar-track"><div class="rpt-bar-fill" style="width:${pct}%;background:${color}">${count}</div></div>
    </div>`;
  }).join('');
}

function getEventIconClass(evento){
  const txt = normalizeStatus(evento);
  if (!txt) return 'fa-solid fa-bell';
  if (txt.includes('COMBUST')) return 'fa-solid fa-gas-pump';
  if (txt.includes('VELOC')) return 'fa-solid fa-gauge-high';
  if (txt.includes('CONEX') || txt.includes('DESCONEX') || txt.includes('SENAL') || txt.includes('SEÑAL')) return 'fa-solid fa-wifi';
  if (txt.includes('BATER')) return 'fa-solid fa-battery-quarter';
  if (txt.includes('PUERTA')) return 'fa-solid fa-door-open';
  if (txt.includes('IGNIC') || txt.includes('MOTOR')) return 'fa-solid fa-key';
  if (txt.includes('SOS') || txt.includes('PANIC') || txt.includes('PANICO') || txt.includes('PÁNICO')) return 'fa-solid fa-triangle-exclamation';
  if (txt.includes('FRENO')) return 'fa-solid fa-circle-stop';
  if (txt.includes('RUTA') || txt.includes('GEO') || txt.includes('CERCA')) return 'fa-solid fa-location-dot';
  return 'fa-solid fa-bell';
}

function eventLabelHtml(evento){
  const txt = evento || 'Sin evento';
  return `<span class="event-label"><i class="${getEventIconClass(txt)}" aria-hidden="true"></i><span>${txt}</span></span>`;
}

function renderEventos(eventos){
  const tbody = document.getElementById('tbodyEventos');
  if(!tbody) return;
  tbody.innerHTML = '';
  eventos.forEach(ev=>{
    const tr = document.createElement('tr');
    const eventoTxt = (ev.descripcion || '');
    const eventoShort = `${eventoTxt.slice(0,40)}${eventoTxt.length > 40 ? '...' : ''}`;
    const pillColor = statusPillColor(ev.estatus);
    tr.innerHTML = `
      <td>${iconPrioridad(ev.prioridad)}</td>
      <td>${ev.generado ?? ''}</td>
      <td title="${eventoTxt}">${eventLabelHtml(eventoShort)}</td>
      <td>${ev.unidad ?? ''}</td>
      <td>
        <span class="pill" style="display:inline-block;background:${pillColor};color:#fff;padding:.2rem .6rem;border-radius:999px;font-size:.8rem">
          ${ev.estatus ?? 'Sin atender'}
        </span>
      </td>
    `;
    tr.style.cursor = 'pointer';
    tr.addEventListener('click', ()=> openDetalle(ev));
    tbody.appendChild(tr);
  });
}

function renderBarChart(containerId, dataMap, maxBars){
  const el = document.getElementById(containerId);
  if (!el) return;
  const entries = [...dataMap.entries()].sort((a,b) => b[1] - a[1]).slice(0, maxBars || 10);
  if (!entries.length){ el.innerHTML = '<div class="rpt-empty">Sin datos para mostrar</div>'; return; }
  const max = entries[0][1] || 1;
  el.innerHTML = entries.map(([label, count], i) => {
    const pct = Math.max((count / max) * 100, 8);
    const color = BAR_COLORS[i % BAR_COLORS.length];
    const labelHtml = containerId === 'chartEvento' ? eventLabelHtml(label) : label;
    return `<div class="rpt-bar-row">
      <span class="rpt-bar-label" title="${label}">${labelHtml}</span>
      <div class="rpt-bar-track"><div class="rpt-bar-fill" style="width:${pct}%;background:${color}">${count}</div></div>
    </div>`;
  }).join('');
}

function renderDashboardDetail(list){
  const tbody = document.getElementById('tbodyDashboardDetalle');
  if (!tbody) return;
  tbody.innerHTML = '';

  if (!list.length){
    tbody.innerHTML = `<tr><td colspan="7" style="text-align:center;color:#666">Sin alertas para mostrar en el rango seleccionado.</td></tr>`;
    return;
  }

  list.forEach(e => {
    const tr = document.createElement('tr');
    const lc = parseLastComment(e.comentarios);
    const usuarioHtml = lc.usuario
      ? `<span class="rpt-user-chip"><i class="fas fa-user"></i>${lc.usuario}</span>`
      : '—';

    tr.innerHTML = `
      <td>${e.folio || '—'}</td>
      <td>${e.generado || '—'}</td>
      <td>${e.unidad || '—'}</td>
      <td title="${e.descripcion || ''}">${e.descripcion || '—'}</td>
      <td>${e.estatus || 'Sin atender'}</td>
      <td>${usuarioHtml}</td>
      <td>${lc.fecha || '—'}</td>
    `;
    tbody.appendChild(tr);
  });
}

function cargarDashboardModule(){
  const list = getReportFiltered();

  // Conteos por estatus
  let pend = 0, seg = 0, atend = 0;
  list.forEach(e => {
    const s = normalizeStatus(e.estatus);
    if (s === 'ATENDIDO' || s === 'CERRADO') atend++;
    else if (s.includes('SEGUIM') || s === 'EN PROCESO' || s === 'ATENDIENDO') seg++;
    else pend++;
  });

  setText('rptTotal', list.length);
  setText('rptPendientes', pend);
  setText('rptSeguimiento2', seg);
  setText('rptAtendidos', atend);

  // Agrupar por tipo de evento
  const byEvento = new Map();
  list.forEach(e => {
    const key = e.descripcion || 'Sin evento';
    byEvento.set(key, (byEvento.get(key) || 0) + 1);
  });
  renderBarChart('chartEvento', byEvento, 10);

  // Agrupar por unidad
  const byUnidad = new Map();
  list.forEach(e => {
    const key = e.unidad || 'Sin unidad';
    byUnidad.set(key, (byUnidad.get(key) || 0) + 1);
  });
  renderBarChart('chartUnidad', byUnidad, 10);

  // Agrupar por usuario que tomó la alerta
  const byUsuario = new Map();
  list.forEach(e => {
    const lc = parseLastComment(e.comentarios);
    const key = lc.usuario || 'Sin atención';
    byUsuario.set(key, (byUsuario.get(key) || 0) + 1);
  });
  renderBarChart('chartUsuario', byUsuario, 10);

  renderDashboardDetail(list);
}

/* Generar dashboard con filtros */
document.getElementById('btnGenReport')?.addEventListener('click', () => cargarDashboardModule());

/* Escapa un campo para CSV (RFC 4180) */
function csvCell(val){
  const s = String(val ?? '');
  if (s.includes(',') || s.includes('"') || s.includes('\n'))
    return '"' + s.replace(/"/g, '""') + '"';
  return s;
}

/* Extrae último usuario y fecha del campo comentarios.
   Formato esperado: [2026-03-24 11:13:15] admin: texto */
function parseLastComment(comentarios){
  if (!comentarios) return { usuario: '', fecha: '' };
  const matches = [...comentarios.matchAll(/\[(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2})\]\s*([^:\n]+):/g)];
  if (!matches.length) return { usuario: '', fecha: '' };
  const last = matches[matches.length - 1];
  return { fecha: (last[1] || '').trim(), usuario: (last[2] || '').trim() };
}

/* Exportar CSV */
document.getElementById('btnExportCSV')?.addEventListener('click', () => {
  const list = getReportFiltered();
  if (!list.length){ alert('No hay datos para exportar.'); return; }

  const headers = ['Folio','Fecha Generado','Unidad','Evento','Estatus','Prioridad','Atendido por','Fecha Atencion','Comentarios'];
  const rows = list.map(e => {
    const lc = parseLastComment(e.comentarios);
    return [
      csvCell(e.folio),
      csvCell(e.generado),
      csvCell(e.unidad),
      csvCell(e.descripcion),
      csvCell(e.estatus),
      csvCell(e.prioridad),
      csvCell(lc.usuario),
      csvCell(lc.fecha),
      csvCell(e.comentarios)
    ];
  });

  const csv = [headers.join(','), ...rows.map(r => r.join(','))].join('\n');
  const blob = new Blob(['\uFEFF' + csv], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `AlerTrack_Dashboard_${new Date().toISOString().slice(0,10)}.csv`;
  a.click();
  URL.revokeObjectURL(url);
});

/* Defaults: último mes */
(function initReportDates(){
  const hoy = new Date();
  const hace30 = new Date(hoy); hace30.setDate(hace30.getDate() - 30);
  const fmt = d => d.toISOString().slice(0,10);
  const rptDesde = document.getElementById('rptDesde');
  const rptHasta = document.getElementById('rptHasta');
  if (rptDesde) rptDesde.value = fmt(hace30);
  if (rptHasta) rptHasta.value = fmt(hoy);
})();

/***************  Inicio ***************/
(async function init(){
  go(linkInicio, secInicio);
  await cargarMapaPrioridad();
  await cargarDashboard();
  updateCardUI();
  // Refresco opcional:
  // setInterval(async ()=>{ await cargarMapaPrioridad(); await cargarDashboard(); }, 60000);
})();
