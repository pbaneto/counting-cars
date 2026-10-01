/**
 * Mini simulador de Apps Script para pruebas de humo en Node: carga TODOS los archivos de src/ en un único
 * ámbito global (como hace Apps Script) y simula Sheets/Drive/UrlFetch en memoria.
 * No evalúa fórmulas (las celdas de fórmula valen '' salvo HYPERLINK): sirve para ejecutar los flujos y detectar errores.
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const SRC = path.join(__dirname, '..', 'src');
const ORDEN = ['Config', 'Logic', 'Prompts', 'Sheets', 'Log', 'Formulas', 'Drive', 'Gemini', 'Trabajos', 'Piezas', 'Albaranes', 'Facturas', 'Abonos',
  'Triggers', 'Setup', 'Resumen', 'Diagnostico', 'Seed'];

function chain() {
  const f = function () { return p; };
  const p = new Proxy(f, { get: (t, k) => (k === 'then' ? undefined : (k === Symbol.toPrimitive ? () => '' : p)), apply: () => p });
  return p;
}

function colNum(s) { let n = 0; for (const ch of s) n = n * 26 + ch.charCodeAt(0) - 64; return n; }
function parseA1(a1) {
  const m = /^([A-Z]+)(\d+)?(?::([A-Z]+)(\d+)?)?$/.exec(a1);
  const c1 = colNum(m[1]), r1 = Number(m[2] || 1), c2 = m[3] ? colNum(m[3]) : c1, r2 = m[3] ? Number(m[4] || 5000) : r1;
  return [r1, c1, r2 - r1 + 1, c2 - c1 + 1];
}

class Rango {
  constructor(sh, r, c, nr, nc) { Object.assign(this, { sh, r, c, nr, nc }); return new Proxy(this, { get: (t, k) => (k in t ? t[k] : chain()) }); }
  _each(fn) { for (let i = 0; i < this.nr; i++) for (let j = 0; j < this.nc; j++) fn(this.r + i, this.c + j, i, j); }
  // Registro de E/S: en Sheets una lectura hecha tras escribir (sin flush) espera al recálculo de fórmulas.
  _w() { this.sh.ss.io.escrito = true; return this; }
  _r(que) { this.sh.ss.io.leer(`${this.sh.name}.${que}`); }
  getValues() { this._r('getValues'); const o = []; for (let i = 0; i < this.nr; i++) { o.push([]); for (let j = 0; j < this.nc; j++) { const x = this.sh.cell(this.r + i, this.c + j); o[i].push(x ? x.v : ''); } } return o; }
  getFormulas() { this._r('getFormulas'); const o = []; for (let i = 0; i < this.nr; i++) { o.push([]); for (let j = 0; j < this.nc; j++) { const x = this.sh.cell(this.r + i, this.c + j); o[i].push(x ? x.f : ''); } } return o; }
  getValue() { this._r('getValue'); const x = this.sh.cell(this.r, this.c); return x ? x.v : ''; }
  getFormula() { this._r('getFormula'); const x = this.sh.cell(this.r, this.c); return x ? x.f : ''; }
  getRow() { return this.r; } getColumn() { return this.c; } getNumRows() { return this.nr; } getNumColumns() { return this.nc; }
  getSheet() { return this.sh; } getA1Notation() { return `R${this.r}C${this.c}`; }
  setValues(a) { this._each((r, c, i, j) => this.sh.put(r, c, a[i][j])); return this._w(); }
  setValue(v) { this._each((r, c) => this.sh.put(r, c, v)); return this._w(); }
  setFormulas(a) { this._each((r, c, i, j) => this.sh.put(r, c, a[i][j])); return this._w(); }
  setFormula(f) { this.sh.put(this.r, this.c, f); return this._w(); }
  clearContent() { this._each((r, c) => this.sh.grid.delete(r + ',' + c)); return this._w(); }
  clearFormat() { return this._w(); }
  getDataValidations() { this._r('getDataValidations'); const o = []; for (let i = 0; i < this.nr; i++) { o.push([]); for (let j = 0; j < this.nc; j++) { const x = this.sh.validaciones.get((this.r + i) + ',' + (this.c + j)); o[i].push(x || null); } } return o; }
  clearDataValidations() { this.sh.validacionesLimpiadas.push(this.c); this._each((r, c) => this.sh.validaciones.delete(r + ',' + c)); return this._w(); }
  /** Como Sheets: una casilla de verificación nunca está vacía, vale FALSE aunque nadie la haya tocado. */
  setDataValidation(regla) { if (regla && regla.casilla) this._each((r, c) => { if (!this.sh.cell(r, c)) this.sh.put(r, c, false); }); return this._w(); }
  setNumberFormat() { return this._w(); }
  // Celdas combinadas como en Sheets: combinar o separar un rango que corta una combinada a medias da error.
  _cortadas() {
    const dentro = m => m.r >= this.r && m.c >= this.c && m.r + m.nr <= this.r + this.nr && m.c + m.nc <= this.c + this.nc;
    const toca = m => m.r < this.r + this.nr && this.r < m.r + m.nr && m.c < this.c + this.nc && this.c < m.c + m.nc;
    if (this.sh.combinadas.some(m => toca(m) && !dentro(m))) throw new Error('Debes seleccionar todas las celdas de un intervalo combinado para combinarlas o separarlas.');
    this.sh.combinadas = this.sh.combinadas.filter(m => !dentro(m));
  }
  merge() { this._cortadas(); this.sh.combinadas.push({ r: this.r, c: this.c, nr: this.nr, nc: this.nc }); return this; }
  insertCells(dim) { if (dim !== 'ROWS') throw new Error('sólo ROWS'); this.sh.desplazarCeldas(this.r, this.c, this.nc, this.nr); return this; }
  deleteCells(dim) { if (dim !== 'ROWS') throw new Error('sólo ROWS'); this.sh.desplazarCeldas(this.r, this.c, this.nc, -this.nr); return this; }
  breakApart() { this._cortadas(); return this; }
}

class Hoja {
  constructor(ss, name, id) {
    this.ss = ss; this.name = name; this.id = id; this.grid = new Map(); this.maxRows = 1000; this.validacionesLimpiadas = []; this.validaciones = new Map(); this.combinadas = [];
    return new Proxy(this, { get: (t, k, rcv) => (k in t || typeof k === 'symbol' ? Reflect.get(t, k, t) : chain()) });
  }
  cell(r, c) { return this.grid.get(r + ',' + c); }
  put(r, c, v) {
    if (r > this.maxRows) this.maxRows = r;
    if (typeof v === 'string' && v.startsWith('=')) {
      const h = /^=HYPERLINK\("([^"]*)"[,;]"([^"]*)"\)$/.exec(v);
      this.grid.set(r + ',' + c, { v: h ? h[2] : '', f: v });
    } else if (v === '' || v == null) this.grid.delete(r + ',' + c);
    else this.grid.set(r + ',' + c, { v, f: '' });
  }
  getName() { return this.name; } getSheetId() { return this.id; }
  getMaxRows() { this.ss.io.leer(`${this.name}.getMaxRows`); return this.maxRows; }
  insertRowsAfter(n, k) { this.maxRows += k; this.ss.io.escrito = true; }
  /** Desplaza hacia abajo, como Sheets, todo lo que esté en la fila r o a partir de ella. */
  insertRowsBefore(r, k) {
    this.combinadas.forEach(m => { if (m.r >= r) m.r += k; else if (m.r + m.nr > r) m.nr += k; });
    const nuevo = new Map();
    this.grid.forEach((v, key) => {
      const [row, col] = key.split(',').map(Number);
      nuevo.set((row >= r ? row + k : row) + ',' + col, v);
    });
    this.grid = nuevo;
    this.maxRows += k;
    this.ss.io.escrito = true;
  }
  deleteRows(r, k) { for (let i = 0; i < k; i++) this.deleteRow(r); }
  /** Desplaza sólo las columnas c..c+nc-1 desde la fila r: k > 0 inserta k celdas (baja), k < 0 borra -k celdas (sube). */
  desplazarCeldas(r, c, nc, k) {
    const nuevo = new Map();
    this.grid.forEach((v, key) => {
      const [row, col] = key.split(',').map(Number);
      const enCols = col >= c && col < c + nc;
      if (!enCols || row < r) { nuevo.set(key, v); return; }
      if (k < 0 && row < r - k) return;  // celdas borradas
      nuevo.set((row + k) + ',' + col, v);
    });
    this.grid = nuevo;
    this.ss.io.escrito = true;
  }
  /** Borra la fila r y sube todo lo de debajo, como Sheets. */
  deleteRow(r) {
    this.combinadas = this.combinadas.filter(m => !(m.r === r && m.nr === 1));
    this.combinadas.forEach(m => { if (m.r > r) m.r--; else if (m.r + m.nr > r) m.nr--; });
    const nuevo = new Map();
    this.grid.forEach((v, key) => {
      const [row, col] = key.split(',').map(Number);
      if (row !== r) nuevo.set((row > r ? row - 1 : row) + ',' + col, v);
    });
    this.grid = nuevo;
    this.maxRows -= 1;
    this.ss.io.escrito = true;
  }
  /** Desplaza a la derecha, como Sheets, todo lo que esté en la columna c o a partir de ella. */
  insertColumnBefore(c) {
    this.combinadas.forEach(m => { if (m.c >= c) m.c++; else if (m.c + m.nc > c) m.nc++; });  // una combinada que la contiene crece
    const nuevo = new Map();
    this.grid.forEach((v, k) => {
      const [r, col] = k.split(',').map(Number);
      nuevo.set(r + ',' + (col >= c ? col + 1 : col), v);
    });
    this.grid = nuevo;
    this.ss.io.escrito = true;
  }
  getLastRow() { this.ss.io.leer(`${this.name}.getLastRow`); let m = 0; this.grid.forEach((x, k) => { const r = Number(k.split(',')[0]); if (r > m) m = r; }); return m; }
  getMaxColumns() { return Math.max(26, this.getLastColumn()); }
  getLastColumn() { let m = 0; this.grid.forEach((x, k) => { const c = Number(k.split(',')[1]); if (c > m) m = c; }); return m; }
  getDataRange() { return new Rango(this, 1, 1, Math.max(this.getLastRow(), 1), Math.max(this.getLastColumn(), 1)); }
  getRange(a, b, c, d) { if (typeof a === 'string') { const [r, cc, nr, nc] = parseA1(a); return new Rango(this, r, cc, nr, nc); } return new Rango(this, a, b, c || 1, d || 1); }
  getProtections() { return []; } getCharts() { return []; } newChart() { return chain(); } insertChart() {} removeChart() {}
  /** Como Sheets: una regla de formato condicional no puede leer otra pestaña, ni directamente ni con un rango con nombre. */
  setConditionalFormatRules(reglas) {
    reglas.forEach(({ formula }) => {
      const f = String(formula).replace(/"[^"]*"/g, '""');
      const ajeno = Object.keys(this.ss.namedRanges).filter(n => this.ss.namedRanges[n] !== this.name && new RegExp(`\\b${n}\\b`).test(f));
      if (f.includes('!') || ajeno.length) throw new Error(`La regla de formato condicional no puede hacer referencia a una hoja diferente. (${this.name}: ${formula})`);
    });
  }
  clear() { this.grid.clear(); }
  valor(r, c) { const x = this.cell(r, c); return x ? x.v : ''; }
}

function crearEntorno(opts = {}) {
  const log = { toasts: [], alerts: [], logger: [], fetch: [], console: [] };
  let idSeq = 1;
  const io = { escrito: false, lecturasTrasEscribir: [], leer(que) { if (io.escrito) io.lecturasTrasEscribir.push(que); } };
  const ss = { sheets: [], namedRanges: {}, io, getSpreadsheetLocale: () => { io.leer('getSpreadsheetLocale'); return opts.locale || 'en_US'; }, toast: (m, t) => log.toasts.push(m), getId: () => 'SS', setSpreadsheetTimeZone() {},
    setNamedRange(nombre, rango) { ss.namedRanges[nombre] = rango.getSheet().getName(); }, setActiveSheet() {}, moveActiveSheet() {} };
  const validacion = () => {
    const b = { casilla: false, requireCheckbox() { b.casilla = true; return b; }, build() { return { casilla: b.casilla }; } };
    const p = new Proxy(b, { get: (t, k) => (k in t ? t[k] : () => p) });
    return p;
  };
  const reglaCF = () => {
    const b = { formula: '', whenFormulaSatisfied(f) { b.formula = f; return b; }, setBackground() { return b; }, setRanges() { return b; }, build() { return { formula: b.formula }; } };
    return b;
  };
  ss.getSheets = () => ss.sheets;
  ss.getSheetByName = n => ss.sheets.find(s => s.name === n) || null;
  ss.insertSheet = n => { const s = new Hoja(ss, n, idSeq++); ss.sheets.push(s); return s; };
  ss.deleteSheet = s => { ss.sheets = ss.sheets.filter(x => x !== s); };
  ss.insertSheet('Hoja 1');

  // Drive simulado
  const carpetas = {}, archivos = {};
  const mkFolder = (id, nombre) => {
    const f = { id, nombre, ficheros: [], getId: () => id, getName: () => nombre,
      getFiles() { let i = 0; const l = f.ficheros; return { hasNext: () => i < l.length, next: () => l[i++] }; },
      getParents() { return { hasNext: () => true, next: () => carpetas.PADRE }; },
      createFolder(n) { const nf = mkFolder('F' + idSeq++, n); return nf; } };
    carpetas[id] = f; return f;
  };
  mkFolder('PADRE', 'padre');
  const mkFile = (id, nombre, mime, carpetaId) => {
    const file = { id, nombre, carpeta: carpetaId, getId: () => id, getName: () => nombre, getMimeType: () => mime,
      getBlob: () => ({ getContentType: () => mime, getBytes: () => Object.assign([1, 2, 3], { nombre }) }),
      moveTo(dest) { carpetas[file.carpeta].ficheros = carpetas[file.carpeta].ficheros.filter(x => x !== file); dest.ficheros.push(file); file.carpeta = dest.id; } };
    archivos[id] = file; carpetas[carpetaId].ficheros.push(file); return file;
  };

  const props = {};
  const ctx = {
    console: { log: m => log.console.push(String(m)), warn: m => log.console.push(String(m)), error: m => log.console.push(String(m)) },
    Logger: { log: m => log.logger.push(m) },
    SpreadsheetApp: { getActiveSpreadsheet: () => ss, openById: () => ss, flush() { io.escrito = false; }, newDataValidation: validacion, newConditionalFormatRule: reglaCF,
      ProtectionType: { RANGE: 'RANGE' }, CopyPasteType: { PASTE_FORMAT: 'F', PASTE_DATA_VALIDATION: 'V' }, Dimension: { ROWS: 'ROWS', COLUMNS: 'COLUMNS' }, getUi: () => ({ alert: (a, b) => log.alerts.push([a, b]), createMenu: chain, prompt: () => ({ getSelectedButton: () => 'CANCEL' }), ButtonSet: { OK: 1, OK_CANCEL: 2 }, Button: { OK: 'OK' } }) },
    DriveApp: { getFolderById: id => { if (!carpetas[id]) throw new Error('carpeta inexistente ' + id); return carpetas[id]; },
      getFileById: id => archivos[id], getRootFolder: () => carpetas.PADRE },
    UrlFetchApp: { fetchAll: reqs => reqs.map(r => { log.fetch.push(r); const x = opts.gemini(r); return { getResponseCode: () => x.code || 200, getContentText: () => x.body }; }) },
    Utilities: { formatDate: (d, tz, f) => { const s = new Intl.DateTimeFormat('sv-SE', { timeZone: tz, year: 'numeric', month: '2-digit', day: '2-digit' }).format(d); return f === 'yyyy-MM-dd' ? s : s; },
      base64Encode: b => (b && b.nombre) || 'AAAA', sleep() {} },  // el "contenido" de un PDF simulado es su nombre
    PropertiesService: { getScriptProperties: () => ({ getProperty: k => props[k] || null, setProperty: (k, v) => { props[k] = String(v); } }) },
    LockService: { getScriptLock: () => ({ tryLock: () => true, releaseLock() {} }) },
    Session: { getEffectiveUser: () => ({ getEmail: () => 'test@example.com' }) },
    ScriptApp: { newTrigger: chain, getProjectTriggers: () => [], deleteTrigger() {} },
  };
  vm.createContext(ctx);
  ORDEN.concat(opts.privado ? ['private'] : []).forEach(n => {
    const f = n === 'private' ? path.join(opts.privadoRuta) : path.join(SRC, n + '.js');
    vm.runInContext(fs.readFileSync(f, 'utf8'), ctx, { filename: f });
  });
  const run = code => vm.runInContext(code, ctx);
  props.SPREADSHEET_ID = 'SS';
  return { ctx, ss, log, run, props, mkFolder, mkFile, carpetas, archivos };
}

module.exports = { crearEntorno };
