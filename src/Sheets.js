/** Acceso a hojas: todo por nombre de cabecera. Errores con mensajes que dicen qué hacer. */

let _ss = null, _cfg = null, _letras = {}, _pyc = null;

function ss_() {
  if (_ss) return _ss;
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  _ss = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
  if (!_ss) throw new Error('No encuentro la hoja de cálculo. Ejecuta "setup" una vez desde la hoja o guarda SPREADSHEET_ID en las propiedades del script.');
  return _ss;
}

/** Fórmula en el formato regional de la hoja (";" en España). TODA fórmula que se escribe pasa por aquí. */
function loc_(f) {
  if (_pyc === null) _pyc = usaPuntoYComa(ss_().getSpreadsheetLocale());
  return localizarFormula(f, _pyc);
}
function locFila_(fila) { return fila.map(v => (typeof v === 'string' && v.charAt(0) === '=' ? loc_(v) : v)); }

function hoja_(nombre) {
  const sh = ss_().getSheetByName(nombre);
  if (!sh) throw new Error(`No existe la pestaña "${nombre}". Ejecuta Counting Cars ▸ Preparar hoja (setup).`);
  return sh;
}

function colLetra_(n) {
  let s = '';
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

/** {cabecera: letra} según la fila 1 REAL de la pestaña (si alguien inserta una columna, las fórmulas se ajustan). */
function letras_(nombre) {
  if (_letras[nombre]) return _letras[nombre];
  const esq = ESQUEMA[nombre], o = {};
  const sh = ss_().getSheetByName(nombre);
  const cab = sh && sh.getLastColumn() > 0 ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0] : [];
  cab.forEach((h, i) => { if (h !== '') o[String(h).trim()] = colLetra_(i + 1); });
  esq.cabeceras.forEach((h, i) => { if (!o[h]) o[h] = colLetra_(i + 1); });
  return (_letras[nombre] = o);
}

function filaConDatos_(vals) { return vals.some(x => x !== '' && x !== false && x != null); }

function tieneDatos_(esq, v) {
  return esq.entradas.some(h => { const x = v[h]; return x !== '' && x !== null && x !== undefined && x !== false; });
}

/**
 * Lee una tabla con UNA sola lectura (cabecera + datos). La cabecera está en la fila 1, salvo que el esquema
 * diga otra cosa con `filaCabecera` (Trabajos: encima lleva el panel de resumen).
 * Devuelve {sh, map, filas:[{fila, v:{cabecera: valor}}], libre, leidas, valores}; leidas = filas de datos leídas
 * (con o sin datos); valores = lo leído tal cual, fila 1 incluida (valores[r - 1] es la fila r).
 */
function leerTabla_(nombre) {
  const sh = hoja_(nombre), esq = ESQUEMA[nombre], fc = esq.filaCabecera || 1;
  const todo = sh.getDataRange().getValues();
  const map = {};
  (todo[fc - 1] || []).forEach((h, i) => { if (h !== '') map[String(h).trim()] = i + 1; });
  esq.cabeceras.forEach(h => {
    if (!map[h]) throw new Error(`Falta la columna "${h}" en la pestaña "${nombre}". Ejecuta Counting Cars ▸ Reparar fórmulas y formato.`);
  });
  const filas = [];
  let libre = fc + 1;
  for (let i = fc; i < todo.length; i++) {
    const v = {};
    for (const h in map) v[h] = todo[i][map[h] - 1];
    if (tieneDatos_(esq, v)) { filas.push({ fila: i + 1, v }); libre = i + 2; }
  }
  return { nombre, sh, map, filas, libre, esq, leidas: todo.length - fc, valores: todo };
}

function anchoTabla_(tabla) { return Math.max.apply(null, Object.keys(tabla.map).map(h => tabla.map[h])); }

/** Añade filas (objetos {cabecera: valor}) al final de la tabla con UNA sola escritura. Las columnas de fórmula se rellenan solas. */
function agregarFilas_(tabla, objs) {
  if (!objs.length) return [];
  const ancho = anchoTabla_(tabla), inicio = tabla.libre;
  const formulas = typeof FORMULAS !== 'undefined' ? FORMULAS[tabla.nombre] : null;
  const data = objs.map((o, k) => {
    const arr = new Array(ancho).fill('');
    for (const h in tabla.map) {
      if (o[h] !== undefined) arr[tabla.map[h] - 1] = o[h];
      else if (formulas && formulas[h]) arr[tabla.map[h] - 1] = loc_(formulas[h](inicio + k));
    }
    return arr;
  });
  const ultima = inicio + data.length - 1;
  if (ultima > tabla.sh.getMaxRows()) tabla.sh.insertRowsAfter(tabla.sh.getMaxRows(), ultima - tabla.sh.getMaxRows() + 100);
  tabla.sh.getRange(inicio, 1, data.length, ancho).setValues(data);
  (tabla.esq.casillas || []).forEach(h => tabla.sh.getRange(inicio, tabla.map[h], data.length, 1).setDataValidation(checkbox_()));
  const filas = objs.map((o, k) => {
    tabla.filas.push({ fila: inicio + k, v: Object.assign({}, o) });
    return inicio + k;
  });
  tabla.libre = inicio + data.length;
  return filas;
}

/** Cambia celdas concretas de una fila existente y mantiene el modelo en memoria al día. */
function actualizarFila_(tabla, fila, cambios) {
  for (const h in cambios) {
    if (!tabla.map[h]) throw new Error(`Columna "${h}" inexistente en "${tabla.nombre}"`);
    tabla.sh.getRange(fila, tabla.map[h]).setValue(cambios[h]);
  }
  const f = tabla.filas.find(x => x.fila === fila);
  if (f) Object.assign(f.v, cambios);
}

/**
 * Pone en la fila las fórmulas que falten y sus casillas (fila escrita a mano, pegada encima o insertada).
 * actuales: fórmulas de la fila si ya se han leído (así esta función sólo escribe).
 */
function asegurarFila_(tabla, fila, actuales) {
  const formulas = FORMULAS[tabla.nombre];
  if (formulas) {
    actuales = actuales || tabla.sh.getRange(fila, 1, 1, anchoTabla_(tabla)).getFormulas()[0];
    for (const h in formulas) if (tabla.map[h] && !actuales[tabla.map[h] - 1]) {
      tabla.sh.getRange(fila, tabla.map[h]).setFormula(loc_(formulas[h](fila)));
    }
  }
  (tabla.esq.casillas || []).forEach(h => tabla.sh.getRange(fila, tabla.map[h]).setDataValidation(checkbox_()));
}

// ---- Fechas (Date <-> "yyyy-mm-dd") ----
function aISO_(v) {
  if (v instanceof Date && !isNaN(v)) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  const s = String(v == null ? '' : v).trim();
  return isoValid(s) ? s : '';
}
function aFecha_(iso) { return iso && isoValid(iso) ? new Date(Number(iso.slice(0, 4)), Number(iso.slice(5, 7)) - 1, Number(iso.slice(8, 10))) : ''; }
function hoyISO_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }

// ---- Config ----
function cfg_(clave) {
  if (!_cfg) {
    _cfg = {};
    try {
      const t = leerTabla_(HOJA.CONFIG);
      t.filas.forEach(f => { _cfg[String(f.v['Clave']).trim()] = f.v['Valor']; });
    } catch (e) { _cfg = {}; }
  }
  if (_cfg[clave] !== undefined && _cfg[clave] !== '') return _cfg[clave];
  const d = CONFIG_DEFECTO.find(x => x[0] === clave);
  return d ? d[1] : '';
}
function cfgNum_(clave) { const n = parseNumber(cfg_(clave)); return Number.isFinite(n) ? n : parseNumber(CONFIG_DEFECTO.find(x => x[0] === clave)[1]); }

function guardarCfg_(clave, valor) {
  const t = leerTabla_(HOJA.CONFIG);
  const f = t.filas.find(x => String(x.v['Clave']).trim() === clave);
  if (f) actualizarFila_(t, f.fila, { 'Valor': valor });
  else agregarFilas_(t, [{ 'Clave': clave, 'Valor': valor, 'Descripción': '' }]);
  if (_cfg) _cfg[clave] = valor;
}
