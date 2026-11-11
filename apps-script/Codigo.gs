/**
 * Énfasis 2026 · Barrio El Castor — conexión entre la página y esta hoja.
 *
 * La página (plan-estaca-merida.github.io) lee y guarda aquí lo que anota cada líder.
 * Cada registro ocupa una fila. La última columna, "datos (no editar)", es la copia
 * completa que usa la página; las otras columnas son para que tú lo leas fácil.
 *
 * No edites ni borres filas a mano: hazlo desde la página.
 */

var ARCHIVO = '1oRVPA8Huor9HN1l_z2rwun_F0oLoW1aTe1IP2SDLBIM'; // Énfasis 2026 · Barrio El Castor (datos)

var HOJAS = { personas: 'Personas', semanas: 'Semanas', lideres: 'Lideres' };

var BARRIOS = {
  'la-vuelta': 'La Vuelta', 'las-americas': 'Las Américas', 'el-castor': 'El Castor', 'ejido': 'Ejido',
  'lagunillas': 'Lagunillas', 'santa-barbara': 'Santa Bárbara', 'vigia': 'Vigía', 'chama': 'Chama', 'estaca': 'Estaca'
};
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

/* ---------- lectura: la página pide todo ---------- */
function doGet(e) {
  try {
    var out = {};
    Object.keys(HOJAS).forEach(function (col) { out[col] = leer_(col); });
    return json_({ ok: true, data: out, t: new Date().toISOString() });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

/* ---------- escritura: set, update o delete de un registro ---------- */
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var req = JSON.parse(e.postData.contents);
    var col = req.col, id = String(req.id || ''), op = req.op;
    if (!HOJAS[col]) return json_({ ok: false, error: 'colección desconocida' });
    if (!id || id.length > 120) return json_({ ok: false, error: 'id inválido' });
    if (e.postData.contents.length > 50000) return json_({ ok: false, error: 'registro demasiado grande' });

    var sh = hoja_(col);
    var fila = buscar_(sh, id);
    var ahora = new Date();

    if (op === 'delete') {
      if (fila) sh.deleteRow(fila);
      return json_({ ok: true });
    }

    var datos;
    if (op === 'set') {
      datos = req.data || {};
    } else if (op === 'update') {
      if (!fila) return json_({ ok: false, error: 'not_found' });
      var previo = JSON.parse(sh.getRange(fila, sh.getLastColumn()).getValue() || '{}');
      datos = merge_(previo, req.data || {});
    } else {
      return json_({ ok: false, error: 'operación desconocida' });
    }

    var valores = [id].concat(COLUMNAS[col](datos), [ahora, JSON.stringify(datos)]);
    if (fila) sh.getRange(fila, 1, 1, valores.length).setValues([valores]);
    else sh.appendRow(valores);
    return json_({ ok: true, data: datos });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

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
    if (v && typeof v === 'object' && !Array.isArray(v) && o[k] && typeof o[k] === 'object' && !Array.isArray(o[k])) o[k] = merge_(o[k], v);
    else o[k] = v;
  });
  return o;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
