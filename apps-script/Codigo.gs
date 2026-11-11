/**
 * Énfasis 2026 · Barrio El Castor — conexión entre la página y esta hoja.
 *
 * La página (plan-estaca-merida.github.io) lee y guarda aquí lo que anota cada líder.
 * Cada registro ocupa una fila. La última columna, "datos (no editar)", es la copia
 * completa que usa la página; las otras columnas son para que tú lo leas fácil.
 *
 * No edites ni borres filas a mano: hazlo desde la página.
 *
 * Cómo publicarla (una sola vez): Implementar → Nueva implementación → Aplicación web,
 * "Ejecutar como: Yo" y "Quién tiene acceso: Cualquier persona". Copia la URL que termina en /exec.
 * Si cambias este código: Implementar → Gestionar implementaciones → editar → Versión: Nueva.
 */

var ARCHIVO = '1oRVPA8Huor9HN1l_z2rwun_F0oLoW1aTe1IP2SDLBIM'; // Énfasis 2026 · Barrio El Castor (datos)

var HOJAS = { personas: 'Personas', semanas: 'Semanas', lideres: 'Lideres' };

var BARRIOS = {
  'la-vuelta': 'La Vuelta', 'las-americas': 'Las Américas', 'el-castor': 'El Castor', 'ejido': 'Ejido',
  'lagunillas': 'Lagunillas', 'santa-barbara': 'Santa Bárbara', 'vigia': 'Vigía', 'chama': 'Chama', 'estaca': 'Estaca'
};
// Barrios que pueden guardar en esta hoja. Prueba: solo El Castor. Nunca se tocan registros de otros barrios.
var PERMITIDOS = ['el-castor'];

var ROLES = { obispado: 'Obispado', eq: 'Cuórum de Élderes', rs: 'Sociedad de Socorro', mj: 'Mujeres Jóvenes', pri: 'Primaria' };

// Columnas legibles de cada hoja (después de "id" y antes de "Actualizado" y "datos").
var COLUMNAS = {
  personas: function (d) { return [barrio_(d.barrio), t_(d.nombre), t_(d.org), t_(d.ord), t_(d.estado)]; },
  semanas: function (d) {
    return [barrio_(d.barrio), t_(d.fecha), ROLES[d.rol] || t_(d.rol || 'Registro general'),
      n_(d.bc), n_(d.sa), n_(d.bp), n_(d.inv), n_(d.sel), n_(d.bpm), n_(d.ent), n_(d.pad)];
  },
  lideres: function (d) { return [barrio_(d.barrio), t_(d.nombre), t_(d.org), t_(d.llamamiento)]; }
};

function barrio_(s) { return BARRIOS[s] || t_(s); }
// Evita que un texto que empieza con = + - @ se convierta en fórmula.
function t_(v) { var s = String(v == null ? '' : v); return /^[=+\-@]/.test(s) ? "'" + s : s; }
function n_(v) { return +v || 0; }

/* ---------- lectura ---------- */
// ?a=v  → solo la versión (muy rápido; la página pregunta cada pocos segundos)
// ?a=all → todos los datos
function doGet(e) {
  try {
    var a = (e && e.parameter && e.parameter.a) || 'all';
    if (a === 'v') return json_({ ok: true, v: version_() });
    var out = {};
    Object.keys(HOJAS).forEach(function (col) { out[col] = leer_(col); });
    return json_({ ok: true, data: out, v: version_(), t: new Date().toISOString() });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

/* ---------- escritura: set, update o delete de un registro ---------- */
function doPost(e) {
  var lock = LockService.getScriptLock();
  var tengo = false;
  try {
    if (!e || !e.postData || e.postData.contents.length > 50000) return json_({ ok: false, error: 'registro demasiado grande' });
    var req = JSON.parse(e.postData.contents);
    var col = req.col, id = String(req.id || ''), op = req.op;
    if (!Object.prototype.hasOwnProperty.call(HOJAS, col)) return json_({ ok: false, error: 'colección desconocida' });
    if (!/^[A-Za-z0-9_-]{1,120}$/.test(id)) return json_({ ok: false, error: 'id inválido' });
    if (req.data != null && (typeof req.data !== 'object' || Array.isArray(req.data))) return json_({ ok: false, error: 'datos inválidos' });

    // Uno a la vez: si varios líderes guardan al mismo tiempo, esperan su turno (hasta 25 s).
    if (!lock.tryLock(25000)) return json_({ ok: false, error: 'lock' });
    tengo = true;

    var sh = hoja_(col);
    var fila = buscar_(sh, id);
    var ahora = new Date();
    var previo = fila ? JSON.parse(sh.getRange(fila, sh.getLastColumn()).getValue() || '{}') : null;
    if (previo && PERMITIDOS.indexOf(previo.barrio) < 0) return json_({ ok: false, error: 'barrio no permitido' });

    if (op === 'delete') {
      if (fila) { sh.deleteRow(fila); subir_(); }
      return json_({ ok: true });
    }

    var datos;
    if (op === 'set') {
      datos = req.data || {};
    } else if (op === 'update') {
      if (!fila) return json_({ ok: false, error: 'not_found' });
      datos = merge_(previo, req.data || {});
    } else {
      return json_({ ok: false, error: 'operación desconocida' });
    }

    if (PERMITIDOS.indexOf(datos.barrio) < 0) return json_({ ok: false, error: 'barrio no permitido' });

    var valores = [id].concat(COLUMNAS[col](datos), [ahora, JSON.stringify(datos)]);
    if (fila) sh.getRange(fila, 1, 1, valores.length).setValues([valores]);
    else sh.appendRow(valores);
    SpreadsheetApp.flush();
    subir_();
    return json_({ ok: true, data: datos });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    if (tengo) { try { lock.releaseLock(); } catch (x) {} }
  }
}

/* ---------- versión: cambia con cada guardado ---------- */
function version_() { return PropertiesService.getScriptProperties().getProperty('VERSION') || '0'; }
function subir_() { PropertiesService.getScriptProperties().setProperty('VERSION', String(Date.now())); }

/* ---------- ayudantes ---------- */
function hoja_(col) {
  var sh = SpreadsheetApp.openById(ARCHIVO).getSheetByName(HOJAS[col]);
  if (!sh) throw new Error('Falta la pestaña ' + HOJAS[col]);
  return sh;
}

function leer_(col) {
  var sh = hoja_(col), n = sh.getLastRow(), res = {};
  if (n < 2) return res;
  var ultima = sh.getLastColumn();
  var ids = sh.getRange(2, 1, n - 1, 1).getValues();
  var js = sh.getRange(2, ultima, n - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    var id = String(ids[i][0] || '');
    if (!id || !js[i][0]) continue;
    try { res[id] = JSON.parse(js[i][0]); } catch (x) {}
  }
  return res;
}

function buscar_(sh, id) {
  var n = sh.getLastRow();
  if (n < 2) return 0;
  var hit = sh.getRange(2, 1, n - 1, 1).createTextFinder(id).matchEntireCell(true).findNext();
  return hit ? hit.getRow() : 0;
}

function merge_(a, b) {
  var o = {};
  Object.keys(a).forEach(function (k) { o[k] = a[k]; });
  Object.keys(b).forEach(function (k) {
    var v = b[k];
    if (k === '__proto__' || k === 'constructor' || k === 'prototype') return;
    if (v && typeof v === 'object' && !Array.isArray(v) && o[k] && typeof o[k] === 'object' && !Array.isArray(o[k])) o[k] = merge_(o[k], v);
    else o[k] = v;
  });
  return o;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
