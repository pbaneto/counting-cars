/** Acceso a hojas: todo por nombre de cabecera. Errores con mensajes que dicen qué hacer. */

let _ss = null, _cfg = null, _letras = {}, _pyc = null, _filaCab = {};

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

/** Un error de fórmula (#REF!…) no es un dato: suele ser una fórmula vieja que ha perdido su columna. */
function tieneDatos_(esq, v) {
  return esq.entradas.some(h => { const x = v[h]; return x !== '' && x !== null && x !== undefined && x !== false && !esErrorFormula_(x); });
}
function esErrorFormula_(x) { return typeof x === 'string' && /^#(REF!|N\/A|VALUE!|DIV\/0!|NAME\?|NUM!|NULL!|ERROR!)$/.test(x); }

/**
 * Lee una tabla con UNA sola lectura (cabecera + datos). La cabecera está en la fila 1, salvo que el esquema
 * diga otra cosa con `filaCabecera` (Trabajos: encima lleva el panel de resumen).
 * Devuelve {sh, map, filas:[{fila, v:{cabecera: valor}}], libre, leidas, valores}; leidas = filas de datos leídas
 * (con o sin datos); valores = lo leído tal cual, fila 1 incluida (valores[r - 1] es la fila r).
 */
/** todo: valores de la pestaña ya leídos (getDataRange().getValues()), para no volver a leerla. */
function leerTabla_(nombre, todo) {
  const sh = hoja_(nombre), esq = ESQUEMA[nombre];
  todo = todo || sh.getDataRange().getValues();
  const pos = filaCabecera_(nombre, todo), fc = pos.fila;
  if (pos.movida) todo = sh.getDataRange().getValues();
  const map = {};
  (todo[fc - 1] || []).forEach((h, i) => { if (h !== '') map[String(h).trim()] = i + 1; });
  esq.cabeceras.forEach(h => {
    if (!map[h]) throw new Error(`Falta la columna "${h}" en la pestaña "${nombre}". El programa busca las columnas por su nombre: si la has renombrado, vuelve a ponerle "${h}" o pide que se cambie en el código.`);
  });
  const filas = [];
  let libre = fc + 1;
  for (let i = fc; i < todo.length; i++) {
    const v = {};
    for (const h in map) v[h] = todo[i][map[h] - 1];
    if (tieneDatos_(esq, v)) { filas.push({ fila: i + 1, v }); libre = i + 2; }
  }
  return { nombre, sh, map, filas, libre, esq, fc, leidas: todo.length - fc, valores: todo };
}

/**
 * Fila de la cabecera de una tabla. Se busca (la primera cabecera en la columna A) en vez de darla por fija, porque
 * alguien puede insertar o borrar filas encima:
 *  - Trabajos (cabeceraMovil): vale cualquier fila desde filaCabecera, p. ej. con una fila en blanco entre el panel y
 *    la cabecera. Si está más arriba (se ha borrado una fila del panel), se insertan filas y se rehace el panel.
 *  - Las demás: se quitan las filas vacías de encima para devolverla a su sitio (sus formatos usan filas fijas).
 * Si no la encuentra y hay datos, o encima hay datos, para con un error: escribir con la cabecera fuera de su sitio
 * pisaría datos (oct 2026: una fila en blanco encima de la cabecera de Trabajos hizo escribir las cabeceras en esa
 * fila y fórmulas encima de Matrícula, Factura y Pagado).
 * todo: valores de la pestaña ya leídos. Devuelve {fila, movida}; movida = se han insertado o borrado filas.
 */
function filaCabecera_(nombre, todo) {
  const esq = ESQUEMA[nombre], fc = esq.filaCabecera || 1, primera = esq.cabeceras[0];
  const vacia = r => (r || []).every(x => x === '' || x == null);
  let h = 0;
  for (let r = 1; r <= Math.min(todo.length, fc + 30); r++) if (String(todo[r - 1][0]).trim() === primera) { h = r; break; }
  if (h === fc || (h > fc && esq.cabeceraMovil)) return { fila: (_filaCab[nombre] = h), movida: false };
  if (!h) {
    if (todo.slice(fc).every(vacia)) return { fila: (_filaCab[nombre] = fc), movida: false };  // hoja nueva: aún no hay tabla
    throw new Error(`No encuentro la cabecera de "${nombre}" ("${primera}" en la columna A). ` +
      'Si se ha borrado, deshazlo (Ctrl+Z) o recupérala desde Archivo ▸ Historial de versiones.');
  }
  const sh = hoja_(nombre);
  if (h < fc) {
    sh.insertRowsBefore(h, fc - h);
    if (nombre === HOJA.TRAB) panelResumenTrabajos_(sh, esq.filasFormato, letras_(nombre), fc);
  } else {
    const vacias = [];
    for (let r = h - 1; r >= 1 && vacias.length < h - fc; r--) if (vacia(todo[r - 1])) vacias.push(r);
    if (vacias.length < h - fc) throw new Error(`En "${nombre}" hay filas con datos encima de la cabecera (fila ${h}). Muévelas o bórralas para que la cabecera vuelva a la fila ${fc}.`);
    vacias.forEach(r => sh.deleteRow(r));  // de abajo arriba: borrar una no mueve las que quedan por borrar
  }
  log_('AVISO', 'filaCabecera', nombre, `La cabecera estaba en la fila ${h}: devuelta a la fila ${fc}.`);
  return { fila: (_filaCab[nombre] = fc), movida: true };
}

/** Fila de la cabecera ya localizada en esta ejecución (si no, se lee la tabla). */
function filaCab_(nombre) { return _filaCab[nombre] || leerTabla_(nombre).fc; }

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
  const formulas = FORMULAS[tabla.nombre], puestas = [];
  if (formulas) {
    actuales = actuales || tabla.sh.getRange(fila, 1, 1, anchoTabla_(tabla)).getFormulas()[0];
    for (const h in formulas) if (tabla.map[h] && !actuales[tabla.map[h] - 1]) {
      tabla.sh.getRange(fila, tabla.map[h]).setFormula(loc_(formulas[h](fila)));
      puestas.push(h);
    }
  }
  (tabla.esq.casillas || []).forEach(h => tabla.sh.getRange(fila, tabla.map[h]).setDataValidation(checkbox_()));
  return puestas;  // columnas cuya fórmula faltaba (o se había pisado escribiendo un valor) y se ha vuelto a poner
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
