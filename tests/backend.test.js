// Prueba de humo del Apps Script con una hoja simulada.  Uso: node tests/backend.test.js
const fs = require('fs'), path = require('path'), assert = require('assert');
const OPS = ['Pendiente', 'Confirmado', 'No Asistirá', 'Enviará representate', 'No contesta varias veces'];
const SEP = process.env.SEP || ';';
function hoja(nombre, datos) {
  const sh = {
    d: datos, nombre, reglas: [], fondo: {}, val: {},
    getName: () => nombre,
    getLastRow() { let r = sh.d.length; while (r > 0 && sh.d[r - 1].every(v => v === '' || v == null)) r--; return r; },
    getLastColumn() { return Math.max(1, ...sh.d.map(f => { let c = f.length; while (c > 0 && (f[c - 1] === '' || f[c - 1] == null)) c--; return c; })); },
    getMaxRows: () => 1157, getMaxColumns: () => 21,
    getRange(r, c, nr = 1, nc = 1) {
      if (typeof r === 'string') { const m = r.match(/([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?/); const L = s => s.charCodeAt(0) - 64; c = L(m[1]); const r1 = +m[2]; nr = m[4] ? +m[4] - r1 + 1 : 1; nc = m[3] ? L(m[3]) - c + 1 : 1; r = r1; }
      const fil = (rr) => { while (sh.d.length < rr) sh.d.push([]); };
      const rg = {
        getRow: () => r, getColumn: () => c, getNumColumns: () => nc,
        getValues() { return Array.from({length: nr}, (_, i) => Array.from({length: nc}, (_, j) => ((sh.d[r - 1 + i] || [])[c - 1 + j] ?? ''))); },
        getValue() { return rg.getValues()[0][0]; },
        setValues(v) { v.forEach((f, i) => { fil(r + i); f.forEach((x, j) => sh.d[r - 1 + i][c - 1 + j] = x); }); return rg; },
        setValue(x) { fil(r); sh.d[r - 1][c - 1] = x; return rg; },
        setFormula(f) { fil(r); sh.d[r - 1][c - 1] = f === '=SUM(2' + SEP + '5)' ? 7 : f; if (/^=/.test(f) && f !== '=SUM(2,5)' && f !== '=SUM(2;5)') sh.formulas = (sh.formulas || []).concat(f); return rg; },
        clearContent() { for (let i = 0; i < nr; i++) for (let j = 0; j < nc; j++) if (sh.d[r - 1 + i]) sh.d[r - 1 + i][c - 1 + j] = ''; return rg; },
        getDataValidation: () => ({ getCriteriaType: () => 'VALUE_IN_LIST', getCriteriaValues: () => [OPS] }),
        copyTo() {}, setBackground(b) { sh.fondo[r + ',' + c] = b; return rg; }
      };
      ['setFontColor', 'setFontWeight', 'setVerticalAlignment', 'setWrap', 'setNumberFormat', 'setFontSize'].forEach(k => rg[k] = () => rg);
      return rg;
    },
    appendRow(f) { sh.d.splice(sh.getLastRow(), 0, f.slice()); },
    insertColumnAfter() {}, setColumnWidth() {}, setFrozenRows() {}, clear() { sh.d = []; },
    getConditionalFormatRules: () => sh.reglas, setConditionalFormatRules(r) { sh.reglas = r; }
  };
  return sh;
}
const base = hoja('BASE INVITADOS', [
  ['Nombres y Apelllidos', 'Cargo', 'Institución', 'Categoría', 'Teléfono', 'Confirmación de Asistencia'],
  ['Luis Fernando Laredo Arellano', 'Diputado Nacional ', 'Cámara de Diputados', '02. Autoridades nacionales', '70733255', ''],
  ['Diego Andres Brañez Leaños', 'Diputado Nacional ', 'Cámara de Diputados', '02. Autoridades nacionales', '67400684', 'Pendiente'],
  ['Saúl Solís', 'Director', 'GAM Cocapata', '03. Autoridades departamentales y municipales', '72201997', 'No contesta varias veces'],
  ['María Laredo Paz', 'Gerente', 'Empresa X', '06. Sector empresarial y productivo', '', '']
]);
const hojas = { 'BASE INVITADOS': base };
// Planilla original: A vacía, B nombre … I confirmación, J representante designado; nombre repetido.
const orig = hoja('Lista de Invitados', [
  ['', 'Nombres y Apelllidos', 'Cargo', 'Institución', 'Categoría', 'Responsable de Contacto', 'Teléfono', 'Invitación enviada', 'Confirmación de Asistencia', 'Representante Designado'],
  ['', 'Ing. Rolando Lopez Lopez', 'Vicerrector', 'Unifranz', '01', '', '', 'Si', '', ''],
  ['', 'Luis Fernando Laredo Arellano', 'Diputado Nacional', 'Cámara de Diputados', '02', '', '70733255', 'Si', '', ''],
  ['', 'Diego Andres Brañez Leaños', 'Diputado Nacional', 'Cámara de Diputados', '02', '', '', 'Si', '', 'Puesto a mano'],
  ['', 'Saúl Solís', 'Director', 'GAM Cocapata', '03', '', '', '', '', ''],
  ['', 'María Laredo Paz', 'Gerente', 'Empresa X', '06', '', '', '', '', ''],
  ['', 'María Laredo Paz', 'Gerente', 'Empresa X', '06', '', '', '', '', '']
]);
let originalFalla = false;
const props = {};
global.SpreadsheetApp = {
  getActive: () => ({ getSheetByName: n => hojas[n] || null, insertSheet: n => (hojas[n] = hoja(n, [])) }),
  getActiveSheet: () => base, flush() {},
  openById: id => { if (originalFalla) throw new Error('sin acceso'); assert.strictEqual(id, '1Nh9ecHGAW7B4HxX6hllduHJahHAwtoH7VHgJMnYbjes'); return { getSheetByName: n => n === 'Lista de Invitados' ? orig : null }; },
  DataValidationCriteria: { VALUE_IN_LIST: 'VALUE_IN_LIST' }, CopyPasteType: { PASTE_FORMAT: 1 },
  newConditionalFormatRule() { const o = { whenTextEqualTo: t => (o.t = t, o), setBackground: () => o, setFontColor: () => o, setRanges: () => o, build: () => ({ t: o.t, getRanges: () => [] }) }; return o; },
  getUi() { throw new Error('sin interfaz'); }
};
global.PropertiesService = { getScriptProperties: () => ({ getProperty: k => props[k] ?? null, setProperty: (k, v) => props[k] = v, deleteProperty: k => delete props[k] }) };
global.LockService = { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) };
global.Utilities = { formatDate: d => d.toISOString() };
global.ContentService = { createTextOutput: t => ({ t, setMimeType() { return this; } }), MimeType: { JSON: 1, JAVASCRIPT: 2 } };
global.Logger = { log() {} };
global.Session = { getActiveUser: () => ({ getEmail: () => '' }) };
eval(fs.readFileSync(path.join(__dirname, '..', 'apps-script', 'Codigo.gs'), 'utf8').replace(/^const /gm, 'var '));

// Configuración inicial
configuracionInicial();
assert.deepStrictEqual(base.d[0].slice(6, 9), ['Representante', 'Cargo del representante', 'Respondió en la app']);
assert.ok(hojas.RESUMEN && hojas.REGISTRO, 'crea RESUMEN y REGISTRO');
assert.strictEqual(base.reglas.length, 5, '5 reglas de color');
assert.ok(hojas.RESUMEN.formulas.every(f => SEP === ';' ? !/,(?=(?:[^"]*"[^"]*")*[^"]*$)/.test(f) : true), 'fórmulas con ;');
console.log('Ejemplo de fórmula RESUMEN:', hojas.RESUMEN.formulas[2]);

// Búsqueda
const ping = JSON.parse(doGet({ parameter: { action: 'ping' } }).t); assert.ok(ping.ok && ping.abierta);
let r = buscar_('laredo'); assert.strictEqual(r.resultados.length, 2);
r = buscar_('arellano luis'); assert.strictEqual(r.resultados.length, 1); assert.strictEqual(r.resultados[0].f, 2);
r = buscar_('BRANEZ'); assert.strictEqual(r.resultados[0].n, 'Diego Andres Brañez Leaños');
r = buscar_('saul'); assert.strictEqual(r.resultados[0].e, 'pendiente', 'No contesta = pendiente para la app');
assert.ok(!('tel' in r.resultados[0]) && !JSON.stringify(r).includes('7220'), 'no expone teléfonos');
assert.strictEqual(buscar_('la').resultados.length, 0, 'mínimo 3 letras');
const jp = doGet({ parameter: { action: 'buscar', q: 'diego', callback: 'cb1' } }).t; assert.ok(jp.startsWith('cb1('));

// Responder
let x = responder_({ f: '2', n: 'Luis Fernando Laredo Arellano', op: 'confirmado' }); assert.ok(x.ok);
assert.strictEqual(base.d[1][5], 'Confirmado'); assert.ok(base.d[1][8] instanceof Date);
x = responder_({ f: '3', n: 'Diego Andres Brañez Leaños', op: 'representante', rep: 'Ana Pérez', rcargo: 'Asesora' }); assert.ok(x.ok);
assert.strictEqual(base.d[2][5], 'Enviará representate'); assert.strictEqual(base.d[2][6], 'Ana Pérez'); assert.strictEqual(base.d[2][7], 'Asesora');
x = responder_({ f: '3', n: 'Diego Andres Brañez Leaños', op: 'no_asistira' }); assert.ok(x.ok);
assert.strictEqual(base.d[2][5], 'No Asistirá'); assert.strictEqual(base.d[2][6], '', 'borra representante al cambiar');
// fila desplazada (se reordenó la base): la encuentra por nombre
x = responder_({ f: '2', n: 'María Laredo Paz', op: 'confirmado' }); assert.ok(x.ok); assert.strictEqual(base.d[4][5], 'Confirmado');
assert.strictEqual(responder_({ f: '2', n: 'Nadie Inexistente', op: 'confirmado' }).codigo, 'no_encontrado');
assert.strictEqual(responder_({ f: '2', n: 'Luis Fernando Laredo Arellano', op: 'representante', rep: '' }).codigo, 'representante');
assert.strictEqual(responder_({ f: '2', n: 'Luis Fernando Laredo Arellano', op: 'xx' }).codigo, 'opcion');
// nuevo invitado
x = responder_({ f: '0', n: '=Pedro Nuevo', op: 'confirmado', nuevo: '1', cargo: 'Gerente', inst: 'ACME', tel: '+591 700-00000' }); assert.ok(x.ok && x.nuevo);
const fn = base.d.findIndex(f => f[0] === "'=Pedro Nuevo"); assert.ok(fn > 0, 'agrega con apóstrofo'); assert.strictEqual(base.d[fn][3], '99. Registrado desde la app'); assert.strictEqual(base.d[fn][5], 'Confirmado');
assert.strictEqual(fn, 5, 'se agrega justo después del último invitado');
assert.strictEqual(hojas.REGISTRO.getLastRow(), 1 + 5, 'REGISTRO con 5 respuestas');
// ── Copia en la planilla original ──
assert.strictEqual(orig.d[2][8], 'Confirmado', 'Laredo confirmado en la original (col I)');
assert.strictEqual(orig.d[3][8], 'No Asistirá', 'Brañez: no asistirá en la original');
assert.strictEqual(orig.d[3][9], '', 'el representante que puso la app se borra al cambiar a «no asistirá»');
assert.strictEqual(orig.d[5][8], 'Confirmado', 'María Laredo (1.ª aparición) en la original');
assert.strictEqual(orig.d[6][8], '', '2.ª aparición intacta');
assert.ok(hojas.REGISTRO.d.some(f => f[8] === 'No encontrado en la planilla original'), 'nuevo invitado: anotado como no encontrado');
assert.ok(hojas.REGISTRO.d.some(f => f[8] === 'OK · fila 3'));
// representante → J; luego confirma → J se borra solo si lo puso la app
x = responder_({ f: '5', n: 'Saúl Solís', op: 'representante', rep: 'Ana Pérez', rcargo: 'Asesora' }); assert.ok(x.ok);
assert.strictEqual(orig.d[4][8], 'Enviará representate'); assert.strictEqual(orig.d[4][9], 'Ana Pérez (Asesora)');
x = responder_({ f: '5', n: 'Saúl Solís', op: 'confirmado' }); assert.ok(x.ok);
assert.strictEqual(orig.d[4][8], 'Confirmado'); assert.strictEqual(orig.d[4][9], '', 'J borrado porque lo había puesto la app');
orig.d[4][9] = 'Escrito a mano';
x = responder_({ f: '5', n: 'Saúl Solís', op: 'confirmado' }); assert.ok(x.ok);
assert.strictEqual(orig.d[4][9], 'Escrito a mano', 'no borra lo que el equipo escribió a mano');
// si la planilla original falla, la respuesta igual se guarda
originalFalla = true;
x = responder_({ f: '5', n: 'Saúl Solís', op: 'no_asistira' }); assert.ok(x.ok); assert.strictEqual(base.d[3][5], 'No Asistirá');
assert.ok(hojas.REGISTRO.d[hojas.REGISTRO.getLastRow() - 1][8].startsWith('Error: sin acceso'));
originalFalla = false;
// sincronizar desde el menú
orig.d.slice(1).forEach(f => { f[8] = ''; });
sincronizarOriginal();
assert.strictEqual(orig.d[2][8], 'Confirmado'); assert.strictEqual(orig.d[4][8], 'No Asistirá');
console.log('✔ Copia en la planilla original OK');
// cerrada
props.CERRADA = 'SI'; assert.strictEqual(responder_({ f: '2', n: 'Luis Fernando Laredo Arellano', op: 'confirmado' }).codigo, 'cerrada'); delete props.CERRADA;
console.log('✔ Todas las pruebas del script pasaron (SEP="' + SEP + '")');
