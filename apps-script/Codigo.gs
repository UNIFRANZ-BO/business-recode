// ════════════════════════════════════════════════════════════════════════════
//  BUSINESS RECODE · Confirmación de asistencia — UNIFRANZ Cochabamba
//  Va pegado en la hoja «INVITADOS BUSSINES RECODE» (Extensiones → Apps Script).
//  Implementar como aplicación web: Ejecutar como «Yo» · Acceso «Cualquier persona».
//  Para actualizar: Implementar → Administrar implementaciones → ✏️ → Nueva versión.
// ════════════════════════════════════════════════════════════════════════════
const VERSION = '1.2';
// Planilla original del equipo (la que manda): cada respuesta se copia ahí al instante y cada 10 minutos esta hoja
// se actualiza desde ella (nombres nuevos, borrados y estados). Ver «Planilla original» más abajo.
const ORIGINAL = { id: '1Nh9ecHGAW7B4HxX6hllduHJahHAwtoH7VHgJMnYbjes', hoja: 'Lista de Invitados' };
const NOTA_REP = 'Representante (vía app): ';
const MARCA_ORIGINAL = 'Ya está en la planilla original';   // nota en el nombre de quien se registró solo
const MINUTOS_ACTUALIZACION = 10;                           // cada cuánto esta hoja se pone igual a la original
const HOJA_BASE = 'BASE INVITADOS';           // si cambias el nombre de la pestaña, cámbialo aquí
const H = { RESUMEN: 'RESUMEN', REGISTRO: 'REGISTRO' };
const MAX_NUEVOS = 60;                        // tope de personas que se registran solas («No encuentro mi nombre»)
const CAT_NUEVO = '99. Registrado desde la app';

// Columnas de la base: se buscan por el nombre del encabezado (fila 1), no por la letra.
const COLS = {
  nombre: ['nombres y apelllidos', 'nombres y apellidos', 'nombre completo', 'nombre'],
  cargo: ['cargo'],
  inst: ['institucion', 'institución', 'empresa'],
  cat: ['categoria', 'categoría'],
  tel: ['telefono', 'teléfono', 'celular'],
  estado: ['confirmacion de asistencia', 'confirmación de asistencia', 'confirmacion', 'estado']
};
// Columnas que agrega «Configuración inicial» a la derecha (si no existen).
const EXTRA = { rep: 'Representante', repCargo: 'Cargo del representante', fecha: 'Respondió en la app' };
const ESTADOS_DEFECTO = ['Pendiente', 'Confirmado', 'No Asistirá', 'Enviará representate', 'No contesta varias veces'];

// Ejecutar UNA vez desde el editor (lista de funciones → autorizar → ▶ Ejecutar) para conceder el permiso
// de escribir en la planilla original. En la ventana de Google marca TODAS las casillas («Seleccionar todo»).
function autorizar() {
  const sh = SpreadsheetApp.openById(ORIGINAL.id).getSheetByName(ORIGINAL.hoja);
  aviso_('Permiso concedido', 'Listo: el script ya puede escribir en la planilla original (pestaña «' + sh.getName() + '», ' + sh.getLastRow() + ' filas).\n\nNo necesitas hacer nada más.');
}

// ─────────────────────────── Aplicación web ───────────────────────────
function doGet(e) {
  const p = (e && e.parameter) || {};
  let out;
  try {
    const a = p.action || 'ping';
    if (a === 'ping') out = { ok: true, servicio: 'business-recode', version: VERSION, abierta: abierta_() };
    else if (a === 'buscar') out = buscar_(p.q);
    else if (a === 'responder') out = responder_(p);
    else out = { ok: false, error: 'Acción desconocida' };
  } catch (err) { out = { ok: false, error: String(err && err.message || err) }; }
  return salida_(out, p.callback);
}
function doPost(e) {
  let out;
  try { out = responder_(JSON.parse((e && e.postData && e.postData.contents) || '{}')); }
  catch (err) { out = { ok: false, error: String(err && err.message || err) }; }
  return salida_(out);
}
function salida_(obj, cb) {
  const json = JSON.stringify(obj);
  if (cb && /^[A-Za-z_$][\w$]{0,80}$/.test(cb)) return ContentService.createTextOutput(cb + '(' + json + ');').setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}
function abierta_() { return PropertiesService.getScriptProperties().getProperty('CERRADA') !== 'SI'; }

// Busca por nombre: cada palabra escrita debe coincidir con el inicio de una palabra del nombre
// (sin importar tildes, mayúsculas ni el orden). Devuelve como máximo 6 coincidencias.
// Nunca devuelve teléfonos ni categorías.
function buscar_(q) {
  const toks = norm_(q).split(' ').filter(Boolean);
  if (norm_(q).replace(/ /g, '').length < 3) return { ok: true, resultados: [] };
  const b = base_(), res = [];
  for (let i = 0; i < b.filas.length; i++) {
    const f = b.filas[i], n = norm_(f[b.c.nombre]);
    if (!n) continue;
    const pal = n.split(' ');
    let pts = 0, ok = true;
    for (const t of toks) {
      if (pal.some(w => w === t)) pts += 3;
      else if (pal.some(w => w.startsWith(t))) pts += 2;
      else if (t.length >= 4 && n.indexOf(t) >= 0) pts += 1;
      else { ok = false; break; }
    }
    if (ok) res.push({ pts, fila: i + 2, f });
  }
  res.sort((x, y) => y.pts - x.pts);
  return {
    ok: true, abierta: abierta_(),
    resultados: res.slice(0, 6).map(r => ({
      f: r.fila,
      n: limpio_(r.f[b.c.nombre]),
      c: b.c.cargo >= 0 ? limpio_(r.f[b.c.cargo]) : '',
      i: b.c.inst >= 0 ? limpio_(r.f[b.c.inst]) : '',
      e: codigoEstado_(r.f[b.c.estado]),
      r: b.c.rep >= 0 ? limpio_(r.f[b.c.rep]) : ''
    }))
  };
}

// Guarda la respuesta. p: {f: fila, n: nombre, op: confirmado|no_asistira|representante,
// rep, rcargo, nuevo: '1' (no está en la base) con cargo, inst, tel, disp}
function responder_(p) {
  if (!abierta_()) return { ok: false, codigo: 'cerrada' };
  const op = String(p.op || '');
  if (['confirmado', 'no_asistira', 'representante'].indexOf(op) < 0) return { ok: false, codigo: 'opcion' };
  const nombre = limpio_(p.n).slice(0, 120);
  const rep = limpio_(p.rep).slice(0, 120), repCargo = limpio_(p.rcargo).slice(0, 120);
  if (nombre.length < 3) return { ok: false, codigo: 'nombre' };
  if (op === 'representante' && rep.length < 3) return { ok: false, codigo: 'representante' };

  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  let hecho;
  try {
    const b = base_(), sh = b.sh;
    // 1) Ubicar la fila: la que mandó la app, si el nombre sigue siendo el mismo; si no, por nombre exacto.
    let fila = 0, f = parseInt(p.f, 10);
    if (f >= 2 && f - 2 < b.filas.length && norm_(b.filas[f - 2][b.c.nombre]) === norm_(nombre)) fila = f;
    if (!fila) for (let i = 0; i < b.filas.length; i++) if (norm_(b.filas[i][b.c.nombre]) === norm_(nombre)) { fila = i + 2; break; }
    let nuevo = false;
    if (!fila) {
      if (p.nuevo !== '1' && p.nuevo !== true) return { ok: false, codigo: 'no_encontrado' };
      const pr = PropertiesService.getScriptProperties(), cuenta = +(pr.getProperty('NUEVOS') || 0);
      if (cuenta >= MAX_NUEVOS) return { ok: false, codigo: 'limite' };
      fila = ultimaFilaNombre_(sh, b.c.nombre) + 1;
      const fila1 = new Array(sh.getLastColumn()).fill('');
      const pon = (col, v) => { if (col >= 0) fila1[col] = seguro_(v); };
      pon(b.c.nombre, nombre); pon(b.c.cargo, limpio_(p.cargo).slice(0, 120)); pon(b.c.inst, limpio_(p.inst).slice(0, 160));
      pon(b.c.cat, CAT_NUEVO); pon(b.c.tel, limpio_(p.tel).replace(/[^\d +\-]/g, '').slice(0, 30));
      fila1[b.c.estado] = '';
      sh.getRange(fila, 1, 1, fila1.length).setValues([fila1]);
      sh.getRange(fila, b.c.nombre + 1).setBackground('#fff2cc');   // amarillo: revisar
      pr.setProperty('NUEVOS', String(cuenta + 1));
      nuevo = true;
    }
    // 2) Escribir estado (exactamente como figura en la lista desplegable), representante y fecha.
    const previo = nuevo ? {} : { rep: b.c.rep >= 0 ? b.filas[fila - 2][b.c.rep] : '', repCargo: b.c.repCargo >= 0 ? b.filas[fila - 2][b.c.repCargo] : '' };
    const estado = textoEstado_(sh, b.c.estado, op), ahora = new Date();
    sh.getRange(fila, b.c.estado + 1).setValue(estado);
    if (b.c.rep >= 0) sh.getRange(fila, b.c.rep + 1).setValue(op === 'representante' ? seguro_(rep) : '');
    if (b.c.repCargo >= 0) sh.getRange(fila, b.c.repCargo + 1).setValue(op === 'representante' ? seguro_(repCargo) : '');
    if (b.c.fecha >= 0) sh.getRange(fila, b.c.fecha + 1).setValue(ahora).setNumberFormat('dd/MM/yyyy HH:mm');
    SpreadsheetApp.flush();
    hecho = { b, fila, nuevo, previo, estado, ahora };
  } finally { lock.releaseLock(); }
  // 3) Fuera del candado: copia en la planilla original y REGISTRO (si la original demora, no frena a nadie).
  const h = hecho;
  const esp = espejo_(datosEspejo_(h.b, h.fila, nombre, h.nuevo ? { cargo: p.cargo, inst: p.inst, tel: p.tel } : null, op, rep, repCargo, h.previo));
  if (h.nuevo && esp.indexOf('Agregado') === 0) try { h.b.sh.getRange(h.fila, h.b.c.nombre + 1).setNote(MARCA_ORIGINAL); } catch (e) {}
  try { registro_([h.ahora, h.fila, seguro_(nombre), h.estado, op === 'representante' ? seguro_(rep) : '', op === 'representante' ? seguro_(repCargo) : '', h.nuevo ? 'Sí' : '', limpio_(p.disp).slice(0, 20), esp]); } catch (e) {}
  return { ok: true, estado: op, nuevo: h.nuevo, fecha: Utilities.formatDate(h.ahora, 'America/La_Paz', 'dd/MM/yyyy HH:mm') };
}

// ─────────────────────────── Utilidades de la base ───────────────────────────
function base_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(HOJA_BASE);
  if (!sh) throw new Error('No existe la pestaña «' + HOJA_BASE + '»');
  const lc = sh.getLastColumn(), lr = sh.getLastRow();
  const enc = sh.getRange(1, 1, 1, lc).getValues()[0].map(norm_);
  const col = alias => { for (const a of alias) { const k = enc.indexOf(norm_(a)); if (k >= 0) return k; } return -1; };
  const c = {};
  Object.keys(COLS).forEach(k => c[k] = col(COLS[k]));
  c.rep = col([EXTRA.rep]); c.repCargo = col([EXTRA.repCargo]); c.fecha = col([EXTRA.fecha]);
  if (c.nombre < 0) throw new Error('No encuentro la columna de nombres en la fila 1');
  if (c.estado < 0) throw new Error('No encuentro la columna «Confirmación de Asistencia»');
  const filas = lr > 1 ? sh.getRange(2, 1, lr - 1, lc).getValues() : [];
  return { sh, c, filas };
}
function ultimaFilaNombre_(sh, colNombre) {
  const lr = sh.getLastRow(); if (lr < 2) return 1;
  const v = sh.getRange(2, colNombre + 1, lr - 1, 1).getValues();
  for (let i = v.length - 1; i >= 0; i--) if (String(v[i][0]).trim()) return i + 2;
  return 1;
}
// Opciones de la lista desplegable de la columna de confirmación (se leen de la hoja; si no hay, las de siempre).
function opcionesEstado_(sh, colEstado) {
  try {
    const r = sh.getRange(2, colEstado + 1).getDataValidation();
    if (r && r.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) return r.getCriteriaValues()[0].map(String);
  } catch (e) {}
  return ESTADOS_DEFECTO;
}
function claveEstado_(t) {
  const n = norm_(t);
  if (!n) return 'pendiente';
  if (n.indexOf('represent') >= 0) return 'representante';
  if (n.indexOf('no contesta') >= 0) return 'no_contesta';
  if (n.indexOf('no asist') >= 0 || n.indexOf('no podra') >= 0) return 'no_asistira';
  if (n.indexOf('confirm') >= 0) return 'confirmado';
  return 'pendiente';
}
function codigoEstado_(t) { const k = claveEstado_(t); return k === 'no_contesta' ? 'pendiente' : k; }
function textoEstado_(sh, colEstado, op) {
  const ops = opcionesEstado_(sh, colEstado);
  for (const o of ops) if (claveEstado_(o) === op) return o;
  return { confirmado: 'Confirmado', no_asistira: 'No Asistirá', representante: 'Enviará representate' }[op];
}
function registro_(fila) {
  const sh = SpreadsheetApp.getActive().getSheetByName(H.REGISTRO) || hojaRegistro_();
  if (String(sh.getRange(1, 9).getValue()) !== 'Planilla original') { sh.getRange(1, 9).setValue('Planilla original'); encabezado_(sh.getRange(1, 9)); sh.setColumnWidth(9, 260); }
  sh.appendRow(fila);
  sh.getRange(sh.getLastRow(), 1).setNumberFormat('dd/MM/yyyy HH:mm:ss');
}
function hojaRegistro_() {
  const sh = SpreadsheetApp.getActive().insertSheet(H.REGISTRO);
  const enc = ['Fecha y hora', 'Fila en la base', 'Nombre', 'Respuesta', 'Representante', 'Cargo del representante', '¿Nuevo?', 'Dispositivo', 'Planilla original'];
  sh.getRange(1, 1, 1, enc.length).setValues([enc]);
  encabezado_(sh.getRange(1, 1, 1, enc.length));
  sh.setFrozenRows(1); sh.setColumnWidth(3, 260); sh.setColumnWidth(5, 220); sh.setColumnWidth(6, 200);
  return sh;
}
function norm_(s) { return String(s === null || s === undefined ? '' : s).toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^a-z0-9ñ ]/g, ' ').replace(/\s+/g, ' ').trim(); }
function limpio_(s) { return String(s === null || s === undefined ? '' : s).replace(/\s+/g, ' ').trim(); }
function seguro_(s) { s = String(s || ''); return /^[=+\-@]/.test(s) ? "'" + s : s; }
function encabezado_(r) { r.setBackground('#0c343d').setFontColor('#ffffff').setFontWeight('bold').setVerticalAlignment('middle').setWrap(true); }

// ─────────────────────────── Menú «Invitación» ───────────────────────────
function onOpen() {
  try {
    SpreadsheetApp.getUi().createMenu('Invitación')
      .addItem('Configuración inicial (una vez)', 'configuracionInicial')
      .addItem('Ver cuántos confirmaron', 'verEstado')
      .addSeparator()
      .addItem('Cerrar confirmaciones', 'cerrar')
      .addItem('Abrir confirmaciones', 'abrir')
      .addSeparator()
      .addItem('Reiniciar la respuesta de la fila seleccionada', 'reiniciarFila')
      .addItem('Reconstruir RESUMEN', 'reconstruirResumen')
      .addItem('Actualizar desde la planilla original ahora', 'actualizarDesdeOriginal')
      .addToUi();
  } catch (e) {}
}
function configuracionInicial() {
  const b = base_(), sh = b.sh;
  // 1) Columnas nuevas a la derecha (no toca las existentes).
  const enc = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(norm_);
  [EXTRA.rep, EXTRA.repCargo, EXTRA.fecha].forEach(t => {
    if (enc.indexOf(norm_(t)) >= 0) return;
    let col = 1; for (let k = enc.length; k >= 1; k--) if (enc[k - 1]) { col = k + 1; break; }   // justo después del último encabezado
    if (col > sh.getMaxColumns()) sh.insertColumnAfter(sh.getMaxColumns());
    sh.getRange(1, col).setValue(t);
    sh.getRange(1, b.c.estado + 1).copyTo(sh.getRange(1, col), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    sh.setColumnWidth(col, t === EXTRA.fecha ? 150 : 230);
    enc[col - 1] = norm_(t);
  });
  // 2) Colores en la columna de confirmación.
  colores_(sh, b.c.estado);
  // 3) Pestañas RESUMEN y REGISTRO.
  if (!SpreadsheetApp.getActive().getSheetByName(H.REGISTRO)) hojaRegistro_();
  reconstruirResumen(true);
  PropertiesService.getScriptProperties().deleteProperty('CERRADA');
  aviso_('Listo', 'Se agregaron las columnas «Representante», «Cargo del representante» y «Respondió en la app», los colores de la columna de confirmación y las pestañas RESUMEN y REGISTRO.\n\nLas confirmaciones están ABIERTAS.');
}
function colores_(sh, colEstado) {
  const L = colEstado + 1, rango = sh.getRange(2, L, sh.getMaxRows() - 1, 1);
  const propias = sh.getConditionalFormatRules().filter(r => !r.getRanges().some(x => x.getColumn() === L && x.getNumColumns() === 1));
  const ops = opcionesEstado_(sh, colEstado), regla = (txt, bg, fg) => SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(txt).setBackground(bg).setFontColor(fg).setRanges([rango]).build();
  const tono = { confirmado: ['#b7e1cd', '#0b5d1e'], no_asistira: ['#f4c7c3', '#8a1c12'], representante: ['#c9daf8', '#1c3f8a'], no_contesta: ['#d9d9d9', '#434343'], pendiente: ['#fce8b2', '#7a5a00'] };
  const nuevas = ops.map(o => { const t = tono[claveEstado_(o)]; return t ? regla(o, t[0], t[1]) : null; }).filter(Boolean);
  sh.setConditionalFormatRules(nuevas.concat(propias));
}
function reconstruirResumen(silencio) {
  const ss = SpreadsheetApp.getActive(), b = base_();
  let sh = ss.getSheetByName(H.RESUMEN);
  if (sh) sh.clear(); else sh = ss.insertSheet(H.RESUMEN);
  sepFormulas_(sh);
  const q = "'" + HOJA_BASE + "'!", L = c => colLetra_(c + 1);
  const N = q + L(b.c.nombre) + '2:' + L(b.c.nombre), E = q + L(b.c.estado) + '2:' + L(b.c.estado), C = b.c.cat >= 0 ? q + L(b.c.cat) + '2:' + L(b.c.cat) : null;
  const ops = opcionesEstado_(b.sh, b.c.estado);
  sh.getRange('A1').setValue('BUSINESS RECODE · Confirmaciones').setFontSize(14).setFontWeight('bold');
  sh.getRange('A2').setFormula(fx_('="Actualizado: "&TEXT(NOW(),"dd/mm/yyyy hh:mm")')).setFontColor('#666666');
  sh.getRange('A4:B4').setValues([['Estado', 'Invitados']]); encabezado_(sh.getRange('A4:B4'));
  let r = 5;
  sh.getRange(r, 1).setValue('Total de invitados'); sh.getRange(r, 2).setFormula(fx_('=COUNTA(' + N + ')')); r++;
  const filasEstado = {};
  ops.forEach(o => {
    sh.getRange(r, 1).setValue(o);
    if (claveEstado_(o) === 'pendiente') sh.getRange(r, 2).setFormula(fx_('=COUNTIFS(' + N + ',"<>",' + E + ',"' + o + '")+COUNTIFS(' + N + ',"<>",' + E + ',"")'));
    else sh.getRange(r, 2).setFormula(fx_('=COUNTIFS(' + N + ',"<>",' + E + ',"' + o + '")'));
    filasEstado[claveEstado_(o)] = r; r++;
  });
  const fc = filasEstado.confirmado, fr = filasEstado.representante;
  sh.getRange(r, 1).setValue('Asistentes previstos (confirmados + representantes)').setFontWeight('bold');
  sh.getRange(r, 2).setFormula('=' + (fc ? 'B' + fc : '0') + '+' + (fr ? 'B' + fr : '0')).setFontWeight('bold').setBackground('#b7e1cd');
  r++;
  sh.getRange(r, 1).setValue('Ya respondieron (%)'); sh.getRange(r, 2).setFormula(fx_('=IFERROR((B5-' + (filasEstado.pendiente ? 'B' + filasEstado.pendiente : '0') + (filasEstado.no_contesta ? '-B' + filasEstado.no_contesta : '') + ')/B5,0)')).setNumberFormat('0%');
  r += 2;
  // Por categoría
  if (C) {
    const cats = [...new Set(b.filas.map(f => limpio_(f[b.c.cat])).filter(Boolean))].sort();
    if (cats.indexOf(CAT_NUEVO) < 0) cats.push(CAT_NUEVO);
    const encC = ['Categoría', 'Invitados'].concat(ops).concat(['Asistentes previstos']);
    sh.getRange(r, 1, 1, encC.length).setValues([encC]); encabezado_(sh.getRange(r, 1, 1, encC.length)); r++;
    cats.forEach(cat => {
      sh.getRange(r, 1).setValue(cat);
      sh.getRange(r, 2).setFormula(fx_('=COUNTIFS(' + N + ',"<>",' + C + ',A' + r + ')'));
      ops.forEach((o, k) => {
        const extra = claveEstado_(o) === 'pendiente' ? '+COUNTIFS(' + N + ',"<>",' + C + ',$A' + r + ',' + E + ',"")' : '';
        sh.getRange(r, 3 + k).setFormula(fx_('=COUNTIFS(' + N + ',"<>",' + C + ',$A' + r + ',' + E + ',"' + o + '")' + extra));
      });
      const kc = ops.findIndex(o => claveEstado_(o) === 'confirmado'), kr = ops.findIndex(o => claveEstado_(o) === 'representante');
      sh.getRange(r, 3 + ops.length).setFormula('=' + (kc >= 0 ? colLetra_(3 + kc) + r : '0') + '+' + (kr >= 0 ? colLetra_(3 + kr) + r : '0'));
      r++;
    });
  }
  sh.setColumnWidth(1, 380); for (let k = 2; k <= 3 + ops.length; k++) sh.setColumnWidth(k, 120);
  sh.setFrozenRows(0);
  if (silencio !== true) aviso_('RESUMEN', 'Se reconstruyó la pestaña RESUMEN.');
}
function verEstado() {
  const b = base_(), cuenta = {};
  b.filas.forEach(f => { if (!limpio_(f[b.c.nombre])) return; const k = claveEstado_(f[b.c.estado]); cuenta[k] = (cuenta[k] || 0) + 1; });
  const t = Object.values(cuenta).reduce((a, x) => a + x, 0);
  aviso_('Business Recode', 'Confirmaciones ' + (abierta_() ? 'ABIERTAS' : 'CERRADAS') + '\n\n' +
    'Invitados: ' + t + '\n' +
    '✔ Confirmados: ' + (cuenta.confirmado || 0) + '\n' +
    '⇄ Enviarán representante: ' + (cuenta.representante || 0) + '\n' +
    '✖ No asistirán: ' + (cuenta.no_asistira || 0) + '\n' +
    '… Pendientes: ' + (cuenta.pendiente || 0) + '\n' +
    '☎ No contestan: ' + (cuenta.no_contesta || 0) + '\n\n' +
    'Asistentes previstos: ' + ((cuenta.confirmado || 0) + (cuenta.representante || 0)));
}
function cerrar() { PropertiesService.getScriptProperties().setProperty('CERRADA', 'SI'); aviso_('Cerrado', 'La app ya no acepta confirmaciones. Para reabrir: menú Invitación → Abrir confirmaciones.'); }
function abrir() { PropertiesService.getScriptProperties().deleteProperty('CERRADA'); aviso_('Abierto', 'La app vuelve a aceptar confirmaciones.'); }
function reiniciarFila() {
  const b = base_(), sh = SpreadsheetApp.getActiveSheet();
  if (sh.getName() !== HOJA_BASE) return aviso_('Atención', 'Primero haz clic en una fila de la pestaña «' + HOJA_BASE + '».');
  const fila = sh.getActiveRange().getRow();
  if (fila < 2) return aviso_('Atención', 'Haz clic en la fila de un invitado (no en el encabezado).');
  const nombre = sh.getRange(fila, b.c.nombre + 1).getValue();
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('Reiniciar respuesta', '¿Dejar como «Pendiente» a ' + nombre + ' (fila ' + fila + ') y borrar su representante?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  const previo = { rep: b.c.rep >= 0 ? sh.getRange(fila, b.c.rep + 1).getValue() : '', repCargo: b.c.repCargo >= 0 ? sh.getRange(fila, b.c.repCargo + 1).getValue() : '' };
  sh.getRange(fila, b.c.estado + 1).setValue(textoEstadoPendiente_(sh, b.c.estado));
  [b.c.rep, b.c.repCargo, b.c.fecha].forEach(c => { if (c >= 0) sh.getRange(fila, c + 1).clearContent(); });
  const esp = espejo_(datosEspejo_(b, fila, nombre, null, 'pendiente', '', '', previo));
  registro_([new Date(), fila, nombre, 'Reiniciada por el administrador', '', '', '', 'Hoja', esp]);
}
function textoEstadoPendiente_(sh, c) { const o = opcionesEstado_(sh, c).find(x => claveEstado_(x) === 'pendiente'); return o || 'Pendiente'; }

// ─────────────────────────── Planilla original (es la que manda) ───────────────────────────
// La original es la base del equipo. Cada respuesta se escribe ahí al instante (columna «Confirmación de Asistencia»,
// representante como nota) y cada 10 minutos esta hoja se pone igual a la original (nombres nuevos, borrados, estados).
// Ojo: en la original, «Representante Designado» (J) es el responsable de Unifranz (lista desplegable): NO se toca.
function original_() {
  const sh = SpreadsheetApp.openById(ORIGINAL.id).getSheetByName(ORIGINAL.hoja);
  if (!sh) throw new Error('No existe la pestaña «' + ORIGINAL.hoja + '» en la planilla original');
  const lc = sh.getLastColumn(), lr = sh.getLastRow();
  const enc = sh.getRange(1, 1, 1, lc).getValues()[0].map(norm_);
  const col = alias => { for (const a of alias) { const k = enc.indexOf(norm_(a)); if (k >= 0) return k; } return -1; };
  const c = { num: enc[0] === '' ? 0 : -1, nombre: col(COLS.nombre), cargo: col(COLS.cargo), inst: col(COLS.inst), cat: col(COLS.cat), tel: col(COLS.tel), estado: col(COLS.estado) };
  if (c.nombre < 0 || c.estado < 0) throw new Error('No encuentro las columnas de nombre o confirmación en la original');
  const v = lr > 1 ? sh.getRange(2, 1, lr - 1, lc).getValues() : [];
  return { sh, c, v };
}
// Datos del cruce: nombre, cargo e institución del invitado y en qué repetición de ese nombre está
// (si un nombre figura 2 veces, la 2.ª aparición de esta hoja corresponde a la 2.ª de la original).
function datosEspejo_(b, fila, nombre, extra, op, rep, repCargo, previo) {
  const k = norm_(nombre); let ordinal = 0;
  for (let i = 0; i < fila - 2 && i < b.filas.length; i++) if (norm_(b.filas[i][b.c.nombre]) === k) ordinal++;
  const f = b.filas[fila - 2] || [];
  return { nombre, ordinal, op, rep, repCargo, previo: previo || {}, nuevo: !!extra,
    cargo: extra ? extra.cargo : (b.c.cargo >= 0 ? f[b.c.cargo] : ''), inst: extra ? extra.inst : (b.c.inst >= 0 ? f[b.c.inst] : ''),
    tel: extra ? extra.tel : '' };
}
function textoRep_(rep, cargo) { rep = limpio_(rep); cargo = limpio_(cargo); return rep ? rep + (cargo ? ' (' + cargo + ')' : '') : ''; }
function filaEnOriginal_(o, d) {
  const idx = []; o.v.forEach((f, i) => { if (norm_(f[o.c.nombre]) === norm_(d.nombre)) idx.push(i); });
  if (!idx.length) return -1;
  if (idx.length === 1) return idx[0];
  const fino = idx.filter(j => (o.c.cargo < 0 || norm_(o.v[j][o.c.cargo]) === norm_(d.cargo)) && (o.c.inst < 0 || norm_(o.v[j][o.c.inst]) === norm_(d.inst)));
  return fino.length === 1 ? fino[0] : idx[Math.min(d.ordinal || 0, idx.length - 1)];
}
// Regla de validación correcta de la columna de confirmación (la que tiene «Confirmado»), para reparar celdas
// que tengan copiada otra lista (pasa en la original desde la fila 117).
function reglaEstado_(o) {
  if (o.regla !== undefined) return o.regla;
  o.regla = null;
  const rs = o.sh.getRange(2, o.c.estado + 1, Math.max(1, Math.min(o.v.length, 400)), 1).getDataValidations();
  for (const r of rs) { const x = r[0]; if (x && listaDe_(x).some(t => claveEstado_(t) === 'confirmado')) { o.regla = x; break; } }
  return o.regla;
}
function listaDe_(regla) {
  try { if (regla.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) return regla.getCriteriaValues()[0].map(String); } catch (e) {}
  return [];
}
// Escribe el estado en una fila de la original (con la opción exacta de su lista) y la nota del representante.
function escribirEstadoOriginal_(o, fila, op, rep, repCargo) {
  const celda = o.sh.getRange(fila, o.c.estado + 1);
  let ops = listaDe_(celda.getDataValidation() || { getCriteriaType: () => '' });
  if (ops.length && !ops.some(t => claveEstado_(t) === op)) {               // la celda tiene otra lista: se repara
    const buena = reglaEstado_(o);
    if (buena) { celda.setDataValidation(buena); ops = listaDe_(buena); }
  }
  if (!ops.length) ops = ESTADOS_DEFECTO;
  const texto = op === 'pendiente' ? (ops.find(t => claveEstado_(t) === 'pendiente') || '') : (ops.find(t => claveEstado_(t) === op) || '');
  celda.setValue(texto);
  const nota = String(celda.getNote() || '');
  if (op === 'representante') celda.setNote(NOTA_REP + textoRep_(rep, repCargo));
  else if (nota.indexOf(NOTA_REP) === 0) celda.setNote('');                 // solo borra la nota que puso la app
}
// Agrega a la original a alguien que se registró solo en la app. Devuelve la fila.
function agregarAOriginal_(o, d) {
  let ult = 1; for (let i = o.v.length - 1; i >= 0; i--) if (limpio_(o.v[i][o.c.nombre])) { ult = i + 2; break; }
  const fila = ult + 1;
  const pon = (col, val) => { if (col >= 0 && limpio_(val)) o.sh.getRange(fila, col + 1).setValue(seguro_(limpio_(val))); };
  if (o.c.num >= 0 && ult >= 2) { const n = parseInt(o.v[ult - 2][o.c.num], 10); if (n > 0) o.sh.getRange(fila, o.c.num + 1).setValue(n + 1); }
  pon(o.c.nombre, d.nombre); pon(o.c.cargo, d.cargo); pon(o.c.inst, d.inst); pon(o.c.cat, CAT_NUEVO); pon(o.c.tel, d.tel);
  o.sh.getRange(fila, o.c.nombre + 1).setBackground('#fff2cc');
  const vacia = new Array(o.v[0] ? o.v[0].length : o.sh.getLastColumn()).fill('');
  vacia[o.c.nombre] = d.nombre; o.v[fila - 2] = vacia;
  return fila;
}
// Copia una respuesta en la original. Devuelve un texto para REGISTRO; nunca lanza error.
function espejo_(d) {
  try {
    if (!ORIGINAL.id) return '';
    const o = original_();
    const i = filaEnOriginal_(o, d);
    let fila, agregado = false;
    if (i >= 0) fila = i + 2;
    else if (d.nuevo) { fila = agregarAOriginal_(o, d); agregado = true; }
    else return 'No encontrado en la planilla original';
    escribirEstadoOriginal_(o, fila, d.op, d.rep, d.repCargo);
    SpreadsheetApp.flush();                                                  // si la hoja rechaza algo, se captura aquí
    return (agregado ? 'Agregado a la original · fila ' : 'OK · fila ') + fila;
  } catch (e) { return 'Error: ' + String(e && e.message || e).slice(0, 200); }
}

// Pone esta hoja igual a la original. La llama el activador cada 10 minutos y el menú.
// · Nombres nuevos en la original → se agregan aquí. · Borrados en la original → se quitan de aquí.
// · Estado: manda la original, salvo respuestas de la app más nuevas que la última actualización (esas se copian allá).
function actualizarDesdeOriginal(e) {
  const manual = !(e && typeof e === 'object') && e !== true;
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(60000)) { if (manual) aviso_('Ocupado', 'Hay respuestas guardándose. Intenta en un minuto.'); return 'ocupado'; }
  let res;
  try {
    const pr = PropertiesService.getScriptProperties(), ultima = +(pr.getProperty('ULTIMA_ACTUALIZACION') || 0), inicio = Date.now();
    const o = original_(), b = base_(), sh = b.sh, W = sh.getLastColumn();
    const opsApp = opcionesEstado_(sh, b.c.estado), textoApp = k => opsApp.find(t => claveEstado_(t) === k) || '';
    const nN = b.filas.length, notas = nN ? sh.getRange(2, b.c.nombre + 1, nN, 1).getNotes().map(r => r[0]) : [];
    const porNombre = {}; b.filas.forEach((f, i) => { const k = norm_(f[b.c.nombre]); if (k) (porNombre[k] = porNombre[k] || []).push(i); });
    const vistos = {}, usados = {}, filas = [], marcas = [], amarillos = [], aOriginal = [];
    let agregados = 0, cambios = 0;
    o.v.forEach((f, j) => {
      const k = norm_(f[o.c.nombre]); if (!k) return;
      const n = vistos[k] = (vistos[k] || 0) + 1, ia = (porNombre[k] || [])[n - 1];
      const prev = ia !== undefined ? b.filas[ia] : null;
      if (prev) usados[ia] = true; else agregados++;
      const fila = prev ? prev.slice(0, W) : new Array(W).fill('');
      while (fila.length < W) fila.push('');
      const pon = (cApp, cOrig) => { if (cApp >= 0 && cOrig >= 0) fila[cApp] = f[cOrig]; };
      pon(b.c.nombre, o.c.nombre); pon(b.c.cargo, o.c.cargo); pon(b.c.inst, o.c.inst); pon(b.c.cat, o.c.cat); pon(b.c.tel, o.c.tel);
      const opOrig = claveEstado_(f[o.c.estado]), opApp = prev ? claveEstado_(prev[b.c.estado]) : 'pendiente';
      const fecha = prev && b.c.fecha >= 0 ? prev[b.c.fecha] : '';
      const appReciente = fecha instanceof Date && fecha.getTime() > ultima && ['confirmado', 'no_asistira', 'representante'].indexOf(opApp) >= 0;
      if (appReciente && opApp !== opOrig) {
        aOriginal.push({ fila: j + 2, op: opApp, rep: prev[b.c.rep], repCargo: b.c.repCargo >= 0 ? prev[b.c.repCargo] : '' });
      } else {
        const t = limpio_(f[o.c.estado]) ? textoApp(opOrig) : '';
        if (prev && t !== String(prev[b.c.estado])) cambios++;
        fila[b.c.estado] = t;
        if (opOrig !== 'representante' && opOrig !== opApp) { if (b.c.rep >= 0) fila[b.c.rep] = ''; if (b.c.repCargo >= 0) fila[b.c.repCargo] = ''; }
      }
      filas.push(fila); marcas.push(prev ? notas[ia] || '' : ''); amarillos.push(limpio_(fila[b.c.cat]) === CAT_NUEVO);
    });
    // Filas que solo están en esta hoja: los registrados solos que aún no pasaron a la original se agregan allá;
    // el resto (borrados en la original) se quitan.
    let quitados = 0;
    b.filas.forEach((f, i) => {
      if (usados[i] || !limpio_(f[b.c.nombre])) return;
      if (limpio_(f[b.c.cat]) === CAT_NUEVO && notas[i] !== MARCA_ORIGINAL) {
        const op = claveEstado_(f[b.c.estado]);
        const fo = agregarAOriginal_(o, { nombre: limpio_(f[b.c.nombre]), cargo: f[b.c.cargo], inst: f[b.c.inst], tel: b.c.tel >= 0 ? f[b.c.tel] : '' });
        if (op !== 'pendiente') aOriginal.push({ fila: fo, op, rep: f[b.c.rep], repCargo: b.c.repCargo >= 0 ? f[b.c.repCargo] : '' });
        filas.push(f.slice(0, W)); marcas.push(MARCA_ORIGINAL); amarillos.push(true);
      } else quitados++;
    });
    aOriginal.forEach(x => escribirEstadoOriginal_(o, x.fila, x.op, x.rep, x.repCargo));
    // Escribir esta hoja solo si algo cambió.
    const antes = JSON.stringify(b.filas.filter(f => limpio_(f[b.c.nombre])).map(f => f.slice(0, W)));
    if (antes !== JSON.stringify(filas) || quitados || agregados) {
      const total = Math.max(nN, filas.length);
      if (total) { sh.getRange(2, 1, total, W).clearContent(); sh.getRange(2, b.c.nombre + 1, total, 1).setBackground(null).clearNote(); }
      if (filas.length) {
        sh.getRange(2, 1, filas.length, W).setValues(filas.map(f => f.map(x => typeof x === 'string' ? seguro_(x) : x)));
        sh.getRange(2, b.c.nombre + 1, filas.length, 1).setNotes(marcas.map(m => [m])).setBackgrounds(amarillos.map(a => [a ? '#fff2cc' : null]));
        if (b.c.fecha >= 0) sh.getRange(2, b.c.fecha + 1, filas.length, 1).setNumberFormat('dd/MM/yyyy HH:mm');
      }
    }
    SpreadsheetApp.flush();
    pr.setProperty('ULTIMA_ACTUALIZACION', String(inicio));
    // Si aparecieron categorías nuevas, rehacer RESUMEN.
    const cats = x => JSON.stringify([...new Set(x.map(f => limpio_(f[b.c.cat])).filter(Boolean))].sort());
    if (b.c.cat >= 0 && cats(b.filas) !== cats(filas)) reconstruirResumen(true);
    res = 'Invitados: ' + filas.length + ' · nuevos: ' + agregados + ' · quitados: ' + quitados + ' · estados traídos de la original: ' + cambios + ' · respuestas copiadas a la original: ' + aOriginal.length;
    if (agregados || quitados || cambios || aOriginal.length) registro_([new Date(), '', 'Actualización desde la planilla original', '', '', '', '', manual ? 'Menú' : 'Automática', res]);
  } catch (err) {
    res = 'Error: ' + String(err && err.message || err).slice(0, 300);
    try { registro_([new Date(), '', 'Actualización desde la planilla original', '', '', '', '', manual ? 'Menú' : 'Automática', res]); } catch (e2) {}
  } finally { lock.releaseLock(); }
  if (manual) aviso_('Actualizado desde la planilla original', res);
  return res;
}
// Ejecutar UNA vez desde el editor: crea el activador de cada 10 minutos y hace la primera actualización.
// Google pedirá permiso para «ejecutar cuando no estás presente»: marca «Seleccionar todo».
function activarActualizacion() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'actualizarDesdeOriginal').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('actualizarDesdeOriginal').timeBased().everyMinutes(MINUTOS_ACTUALIZACION).create();
  const r = actualizarDesdeOriginal(true);
  aviso_('Actualización automática activada', 'Cada ' + MINUTOS_ACTUALIZACION + ' minutos esta hoja se pone igual a la planilla original.\n\nPrimera actualización:\n' + r);
}

// ─────────────────────────── Fórmulas (hoja en español usa «;») ───────────────────────────
function sepFormulas_(sh) {
  try {
    const c = sh.getRange(1, 1);
    for (const s of [',', ';']) {
      c.setFormula('=SUM(2' + s + '5)'); SpreadsheetApp.flush();
      if (c.getValue() === 7) { c.clearContent(); PropertiesService.getScriptProperties().setProperty('SEP', s); return s; }
    }
    c.clearContent();
  } catch (e) {}
  return sep_();
}
function sep_() { try { return PropertiesService.getScriptProperties().getProperty('SEP') || ','; } catch (e) { return ','; } }
function fx_(f) { return sep_() === ',' ? f : f.replace(/("[^"]*")|,/g, (m, q) => q || ';'); }
function colLetra_(n) { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
function aviso_(titulo, texto) { try { SpreadsheetApp.getUi().alert(titulo, texto, SpreadsheetApp.getUi().ButtonSet.OK); } catch (e) { Logger.log(titulo + '\n' + texto); } }
