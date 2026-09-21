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
  getValues() { const o = []; for (let i = 0; i < this.nr; i++) { o.push([]); for (let j = 0; j < this.nc; j++) { const x = this.sh.cell(this.r + i, this.c + j); o[i].push(x ? x.v : ''); } } return o; }
  getFormulas() { const o = []; for (let i = 0; i < this.nr; i++) { o.push([]); for (let j = 0; j < this.nc; j++) { const x = this.sh.cell(this.r + i, this.c + j); o[i].push(x ? x.f : ''); } } return o; }
  getValue() { const x = this.sh.cell(this.r, this.c); return x ? x.v : ''; }
  getFormula() { const x = this.sh.cell(this.r, this.c); return x ? x.f : ''; }
  getRow() { return this.r; } getColumn() { return this.c; } getNumRows() { return this.nr; } getNumColumns() { return this.nc; }
  getSheet() { return this.sh; } getA1Notation() { return `R${this.r}C${this.c}`; }
  setValues(a) { this._each((r, c, i, j) => this.sh.put(r, c, a[i][j])); return this; }
  setValue(v) { this._each((r, c) => this.sh.put(r, c, v)); return this; }
  setFormulas(a) { this._each((r, c, i, j) => this.sh.put(r, c, a[i][j])); return this; }
  setFormula(f) { this.sh.put(this.r, this.c, f); return this; }
  clearContent() { this._each((r, c) => this.sh.grid.delete(r + ',' + c)); return this; }
  setNumberFormat() { return this; }
}

class Hoja {
  constructor(ss, name, id) {
    this.ss = ss; this.name = name; this.id = id; this.grid = new Map(); this.maxRows = 1000;
    return new Proxy(this, { get: (t, k, rcv) => (k in t || typeof k === 'symbol' ? Reflect.get(t, k, t) : chain()) });
  }
  cell(r, c) { return this.grid.get(r + ',' + c); }
  put(r, c, v) {
    if (r > this.maxRows) this.maxRows = r;
    if (typeof v === 'string' && v.startsWith('=')) {
      const h = /^=HYPERLINK\("([^"]*)","([^"]*)"\)$/.exec(v);
      this.grid.set(r + ',' + c, { v: h ? h[2] : '', f: v });
    } else if (v === '' || v == null) this.grid.delete(r + ',' + c);
    else this.grid.set(r + ',' + c, { v, f: '' });
  }
  getName() { return this.name; } getSheetId() { return this.id; }
  getMaxRows() { return this.maxRows; }
  insertRowsAfter(n, k) { this.maxRows += k; }
  getLastRow() { let m = 0; this.grid.forEach((x, k) => { const r = Number(k.split(',')[0]); if (r > m) m = r; }); return m; }
  getLastColumn() { let m = 0; this.grid.forEach((x, k) => { const c = Number(k.split(',')[1]); if (c > m) m = c; }); return m; }
  getRange(a, b, c, d) { if (typeof a === 'string') { const [r, cc, nr, nc] = parseA1(a); return new Rango(this, r, cc, nr, nc); } return new Rango(this, a, b, c || 1, d || 1); }
  getProtections() { return []; } getCharts() { return []; } newChart() { return chain(); } insertChart() {} removeChart() {}
  clear() { this.grid.clear(); }
  valor(r, c) { const x = this.cell(r, c); return x ? x.v : ''; }
}

function crearEntorno(opts = {}) {
  const log = { toasts: [], alerts: [], logger: [], fetch: [] };
  let idSeq = 1;
  const ss = { sheets: [], toast: (m, t) => log.toasts.push(m), getId: () => 'SS', setSpreadsheetTimeZone() {}, setNamedRange() {}, setActiveSheet() {}, moveActiveSheet() {} };
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
      getBlob: () => ({ getContentType: () => mime, getBytes: () => [1, 2, 3] }),
      moveTo(dest) { carpetas[file.carpeta].ficheros = carpetas[file.carpeta].ficheros.filter(x => x !== file); dest.ficheros.push(file); file.carpeta = dest.id; } };
    archivos[id] = file; carpetas[carpetaId].ficheros.push(file); return file;
  };

  const props = {};
  const ctx = {
    console, Logger: { log: m => log.logger.push(m) },
    SpreadsheetApp: { getActiveSpreadsheet: () => ss, openById: () => ss, newDataValidation: chain, newConditionalFormatRule: chain,
      ProtectionType: { RANGE: 'RANGE' }, getUi: () => ({ alert: (a, b) => log.alerts.push([a, b]), createMenu: chain, prompt: () => ({ getSelectedButton: () => 'CANCEL' }), ButtonSet: { OK: 1, OK_CANCEL: 2 }, Button: { OK: 'OK' } }) },
    DriveApp: { getFolderById: id => { if (!carpetas[id]) throw new Error('carpeta inexistente ' + id); return carpetas[id]; },
      getFileById: id => archivos[id], getRootFolder: () => carpetas.PADRE },
    UrlFetchApp: { fetchAll: reqs => reqs.map(r => { log.fetch.push(r); const x = opts.gemini(r); return { getResponseCode: () => x.code || 200, getContentText: () => x.body }; }) },
    Utilities: { formatDate: (d, tz, f) => { const s = new Intl.DateTimeFormat('sv-SE', { timeZone: tz, year: 'numeric', month: '2-digit', day: '2-digit' }).format(d); return f === 'yyyy-MM-dd' ? s : s; },
      base64Encode: () => 'AAAA', sleep() {} },
    PropertiesService: { getScriptProperties: () => ({ getProperty: k => props[k] || null, setProperty: (k, v) => { props[k] = v; } }) },
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
