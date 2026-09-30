/**
 * LÓGICA PURA: sin Sheets, sin Drive, sin red.
 * Se ejecuta igual en Apps Script y en Node (tests/logic.test.js).
 * Las fechas viajan como texto ISO "yyyy-mm-dd" para no depender de zonas horarias.
 */

const IVA_DEFECTO = 0.21;

function round2(n) { return Math.round((Number(n) + Number.EPSILON) * 100) / 100; }

/** Acepta 12.5, "12,50", "1.234,56", "12.50". Devuelve NaN si no es número. */
function parseNumber(v) {
  if (typeof v === 'number') return v;
  if (v == null) return NaN;
  let s = String(v).trim().replace(/[€\s]/g, '');
  if (s === '') return NaN;
  if (s.indexOf(',') >= 0) s = s.replace(/\./g, '').replace(',', '.');
  return Number(s);
}

/** "7853 kcc" -> "7853KCC" */
function normPlate(s) { return String(s == null ? '' : s).toUpperCase().replace(/[^A-Z0-9]/g, ''); }

/** Nº de albarán como texto: sólo el texto tal cual, sin espacios. */
function normAlbaran(s) { return String(s == null ? '' : s).trim().replace(/\s+/g, ''); }

/** "01000443645" -> "443645" (los abonos de RM anteponen 01000 al nº de albarán original). */
function albaranOrigen(s) {
  const d = String(s == null ? '' : s).replace(/\D/g, '');
  return d.length > 6 && d.indexOf('01000') === 0 ? d.slice(5) : d;
}

/**
 * {nº albarán: matrícula} a partir de pares [albarán, matrícula] en orden de prioridad (gana el primero que la tenga).
 * Sirve para los abonos de RM: su bloque en la factura trae la matrícula vacía, pero su albarán original sí la tiene.
 */
function mapaMatriculas(pares) {
  const m = {};
  pares.forEach(([a, p]) => { const k = normAlbaran(a), v = normPlate(p); if (k && v && !m[k]) m[k] = v; });
  return m;
}

/** Clave para comparar referencias de pieza. */
function refKey(s) { return String(s == null ? '' : s).toUpperCase().replace(/[^A-Z0-9]/g, ''); }

/** Prefijo del nº de trabajo: último dígito + letras finales. 7853KCC -> "3KCC". */
function jobPrefix(plate) {
  const p = normPlate(plate);
  const m = p.match(/^(.*?)([A-Z]+)$/);
  if (!m) return p;
  const digit = /\d$/.test(m[1]) ? m[1].slice(-1) : '';
  return digit + m[2];
}

/** Siguiente nº de trabajo libre para un prefijo: usa el máximo existente + 1 (nunca reutiliza). */
function nextJobNumber(prefix, existingNumbers) {
  let max = 0;
  for (const n of existingNumbers) {
    const m = String(n).match(/^(.*)-(\d+)$/);
    if (m && m[1] === prefix) max = Math.max(max, Number(m[2]));
  }
  return prefix + '-' + (max + 1);
}

/** Trabajo sin pagar más reciente (mayor contador) de esa matrícula, o null. jobs: [{num, plate, pagado}] */
function pickOpenJob(plate, jobs) {
  const p = normPlate(plate);
  let best = null, bestN = -1;
  for (const j of jobs) {
    if (normPlate(j.plate) !== p || j.pagado) continue;
    const m = String(j.num).match(/-(\d+)$/);
    const n = m ? Number(m[1]) : 0;
    if (n > bestN) { best = j; bestN = n; }
  }
  return best;
}

function isoValid(iso) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(iso || ''));
  if (!m) return false;
  const d = new Date(Date.UTC(+m[1], +m[2] - 1, +m[3]));
  return d.getUTCFullYear() === +m[1] && d.getUTCMonth() === +m[2] - 1 && d.getUTCDate() === +m[3];
}

function daysBetween(isoA, isoB) {
  const a = isoA.split('-').map(Number), b = isoB.split('-').map(Number);
  return Math.round((Date.UTC(b[0], b[1] - 1, b[2]) - Date.UTC(a[0], a[1] - 1, a[2])) / 86400000);
}

/** 1 si día 1-15, 2 en el resto. */
function quincenaDe(iso) { return Number(iso.slice(8, 10)) <= 15 ? 1 : 2; }

function ultimoDiaMes(y, m) { return new Date(Date.UTC(y, m, 0)).getUTCDate(); }

function rangoQuincena(year, month, q) {
  const mm = String(month).padStart(2, '0');
  return { desde: `${year}-${mm}-${q === 1 ? '01' : '16'}`, hasta: `${year}-${mm}-${q === 1 ? '15' : String(ultimoDiaMes(year, month))}` };
}

function esResiduo(linea) { return /SIGAUS/i.test(String(linea && linea.descripcion || '')); }

/** Líneas de un albarán que van a la pestaña Piezas (sin residuos SIGAUS ni líneas sin importe). */
function lineasParaPiezas(lineas) {
  return (lineas || []).filter(l => !esResiduo(l) && Number.isFinite(parseNumber(l.importe)));
}

/**
 * Validación de un albarán leído por Gemini.
 * errors  -> no se puede crear la fila (se mueve a Errores).
 * warnings-> se crea la fila y se avisa en "Nota escaneo".
 */
function validarAlbaran(doc, iva) {
  iva = iva == null ? IVA_DEFECTO : iva;
  const errors = [], warnings = [];
  const total = parseNumber(doc.total), base = parseNumber(doc.base_imponible);
  if (!Number.isFinite(total) || total <= 0) errors.push('No se ha podido leer el total del albarán');
  if (!doc.numero_albaran) warnings.push('Sin nº de albarán');
  if (!isoValid(doc.fecha)) warnings.push('Fecha ilegible: se usa la fecha de hoy');
  if (!normPlate(doc.matricula)) warnings.push('Sin matrícula');
  if (Number.isFinite(total) && Number.isFinite(base) && Math.abs(round2(base * (1 + iva)) - total) > 0.02) {
    warnings.push(`IVA no cuadra: base ${base} → total ${total}`);
  }
  const lineas = doc.lineas || [];
  const suma = round2(lineas.reduce((a, l) => a + (Number.isFinite(parseNumber(l.importe)) ? parseNumber(l.importe) : 0), 0));
  if (Number.isFinite(base) && lineas.length && Math.abs(suma - base) > 0.05) {
    warnings.push(`La suma de líneas (${suma}) no cuadra con la base (${base})`);
  }
  if (!lineas.length) warnings.push('Sin líneas de pieza');
  lineas.forEach((l, i) => {
    if (esResiduo(l) || !l.referencia) return;
    const q = parseNumber(l.cantidad), p = parseNumber(l.precio_unitario), d = parseNumber(l.descuento_pct) || 0, imp = parseNumber(l.importe);
    if ([q, p, imp].every(Number.isFinite) && Math.abs(q * p * (1 - d / 100) - imp) > 0.03) {
      warnings.push(`Línea ${i + 1} (${l.referencia}): ${q}×${p} con ${d}% dto ≠ ${imp}`);
    }
  });
  return { errors, warnings };
}

/** Validación de una factura quincenal de RM leída por Gemini. */
function validarFactura(doc, iva) {
  iva = iva == null ? IVA_DEFECTO : iva;
  const errors = [], warnings = [];
  const base = parseNumber(doc.base_imponible), total = parseNumber(doc.total), cuota = parseNumber(doc.iva_importe);
  if (!doc.numero_factura) errors.push('No se ha podido leer el nº de factura');
  if (!Number.isFinite(total) || total <= 0) errors.push('No se ha podido leer el total de la factura');
  if (!isoValid(doc.fecha_factura)) errors.push('Fecha de factura ilegible');
  const albs = doc.albaranes || [];
  if (!albs.length) errors.push('La factura no contiene albaranes');
  if (Number.isFinite(base) && Number.isFinite(cuota) && Number.isFinite(total)) {
    if (Math.abs(base + cuota - total) > 0.02) warnings.push(`Base ${base} + IVA ${cuota} ≠ total ${total}`);
    if (Math.abs(round2(base * iva) - cuota) > 0.02) warnings.push(`El IVA ${cuota} no es el ${iva * 100}% de la base ${base}`);
  } else if (Number.isFinite(base) && Number.isFinite(total) && Math.abs(round2(base * (1 + iva)) - total) > 0.02) {
    warnings.push(`Base ${base} ×${1 + iva} ≠ total ${total}`);
  }
  let sumaAlb = 0;
  albs.forEach(a => {
    const imp = parseNumber(a.importe);
    const sl = round2((a.lineas || []).reduce((s, l) => s + (Number.isFinite(parseNumber(l.importe)) ? parseNumber(l.importe) : 0), 0));
    if (Number.isFinite(imp)) sumaAlb += imp;
    if (Number.isFinite(imp) && Math.abs(sl - imp) > 0.03) warnings.push(`Albarán ${a.numero_albaran}: líneas ${sl} ≠ importe ${imp}`);
  });
  if (Number.isFinite(base) && albs.length && Math.abs(round2(sumaAlb) - base) > 0.05) {
    warnings.push(`La suma de albaranes (${round2(sumaAlb)}) no cuadra con la base imponible (${base})`);
  }
  return { errors, warnings };
}

/** Quincena a la que pertenece una factura: la de la fecha más reciente de sus albaranes de compra. */
function periodoFactura(doc) {
  const compras = (doc.albaranes || []).filter(a => !a.es_abono && isoValid(a.fecha));
  const fechas = (compras.length ? compras : (doc.albaranes || []).filter(a => isoValid(a.fecha))).map(a => a.fecha).sort();
  const iso = fechas.length ? fechas[fechas.length - 1] : doc.fecha_factura;
  return { year: Number(iso.slice(0, 4)), month: Number(iso.slice(5, 7)), quincena: quincenaDe(iso) };
}

/**
 * Claves de la tabla de Abonos: cada fila lleva en una columna oculta las claves de lo que la originó, separadas por ";".
 *  P|albarán|referencia|n  pieza de Piezas con Reembolso ✓ (n = nº de orden entre las piezas con el mismo albarán y referencia)
 *  A|factura|albarán origen|referencia|n  línea de abono de una factura RM
 * Una fila que es a la vez pieza pedida y abono recibido lleva las dos.
 */
function ponerClaves(items, prefijo, partes) {
  const vistos = {};
  items.forEach(it => {
    const base = [prefijo].concat(partes(it)).join('|');
    vistos[base] = (vistos[base] || 0) + 1;
    it.clave = base + '|' + vistos[base];
  });
  return items;
}
function refODesc_(ref, desc) { return refKey(ref) || 'D' + refKey(desc); }
function clavesPiezas(piezas) { return ponerClaves(piezas, 'P', p => [normAlbaran(p.albaran), refODesc_(p.ref, p.desc)]); }
function clavesAbonos(abonos) { return ponerClaves(abonos, 'A', a => [String(a.factura || '').trim(), normAlbaran(a.albaranOrigen), refODesc_(a.ref, a.desc)]); }
function partirClaves(s) { return String(s == null ? '' : s).split(';').map(x => x.trim()).filter(Boolean); }

/**
 * Sincroniza la tabla EDITABLE de Abonos sin tocar lo que ya hay escrito:
 *  - pieza marcada que no está en la tabla → fila nueva "Sin abonar";
 *  - línea de abono nueva (ni en la tabla ni vista antes) → rellena la fila "Sin abonar" de su pieza (mismo albarán +
 *    referencia, o mismo importe) y la pasa a "Abonada"; si no hay pieza pedida, fila nueva "Sin solicitar";
 *  - quitar: claves P de piezas desmarcadas → se borran sus filas.
 * De las filas existentes sólo se rellenan celdas VACÍAS, y el Estado sólo cambia si seguía en "Sin abonar".
 *  existentes: filas de la tabla, en orden {clave, estado, albaran, ref, sinIva, fechaAbono, factura, nota}
 *  marcadas / todasPiezas: {clave, albaran, ref, desc, sinIva, fechaReembolso, matricula}
 *  abonos: {clave, factura, fecha, albaranOrigen, ref, desc, importe, matricula}
 *  vistas: Set de claves A añadidas alguna vez (una fila de abono borrada a mano no vuelve a aparecer)
 *  matriculas: {nº albarán: matrícula} para rellenar la Matrícula de las filas que la tengan vacía
 * Devuelve { nuevas (en el orden en que van arriba: la más reciente primero), cambios: [{i, v}], borrar: [i], registrar: [claves A] }.
 */
function sincronizarAbonos(existentes, marcadas, todasPiezas, abonos, vistas, quitar, matriculas) {
  quitar = new Set(quitar || []);
  const borrar = [], enTabla = new Set(), cambios = {};
  existentes.forEach((f, i) => {
    const ks = partirClaves(f.clave);
    if (ks.some(k => quitar.has(k))) { borrar.push(i); return; }
    ks.forEach(k => enTabla.add(k));
  });
  const nuevas = [];
  marcadas.forEach(p => {
    if (enTabla.has(p.clave) || quitar.has(p.clave)) return;
    enTabla.add(p.clave);
    nuevas.push({ clave: p.clave, fechaAbono: '', descripcion: p.desc, sinIva: round2(parseNumber(p.sinIva)), estado: 'Sin abonar',
      albaran: normAlbaran(p.albaran), referencia: p.ref || '', matricula: p.matricula || '', fechaSolicitud: p.fechaReembolso || '', factura: '', nota: '' });
  });
  // Candidatas a recibir un abono: filas de pieza pedida (clave P) que aún no tienen abono (clave A).
  const candidatas = [];
  existentes.forEach((f, i) => {
    const ks = partirClaves(f.clave);
    if (borrar.indexOf(i) < 0 && ks.some(k => k.indexOf('P|') === 0) && !ks.some(k => k.indexOf('A|') === 0)) {
      candidatas.push({ existente: i, albaran: f.albaran, ref: f.ref, sinIva: f.sinIva, estado: f.estado, vacias: f });
    }
  });
  nuevas.forEach(n => candidatas.push({ nueva: n, albaran: n.albaran, ref: n.referencia, sinIva: n.sinIva, estado: n.estado }));
  const buscar = (a, porRef) => candidatas.findIndex(c => normAlbaran(c.albaran) === normAlbaran(a.albaranOrigen) && (porRef
    ? refKey(c.ref) && refKey(c.ref) === refKey(a.ref)
    : Math.abs(Math.abs(parseNumber(a.importe)) - parseNumber(c.sinIva)) < 0.02));
  const registrar = [];
  const ordenados = abonos.map((a, i) => Object.assign({ _i: i }, a)).sort((x, y) => String(x.fecha).localeCompare(String(y.fecha)) || x._i - y._i);
  ordenados.forEach(a => {
    if (enTabla.has(a.clave) || vistas.has(a.clave)) return;
    enTabla.add(a.clave);
    registrar.push(a.clave);
    const sin = round2(Math.abs(parseNumber(a.importe)));
    let j = buscar(a, true);
    if (j < 0) j = buscar(a, false);
    if (j >= 0) {
      const c = candidatas.splice(j, 1)[0], pedido = round2(parseNumber(c.sinIva));
      const nota = Math.abs(pedido - sin) > 0.02 ? `Importe abonado distinto: se pidió ${pedido} sin IVA` : '';
      if (c.nueva) {
        Object.assign(c.nueva, { clave: c.nueva.clave + ';' + a.clave, fechaAbono: a.fecha, estado: 'Abonada', factura: a.factura, nota });
        if (!c.nueva.matricula) c.nueva.matricula = a.matricula || '';
        return;
      }
      const f = c.vacias, v = { clave: partirClaves(f.clave).concat(a.clave).join(';') };
      if (f.estado === 'Sin abonar') v.estado = 'Abonada';
      if (!f.fechaAbono) v.fechaAbono = a.fecha;
      if (!f.factura) v.factura = a.factura;
      if (!f.nota && nota) v.nota = nota;
      cambios[c.existente] = v;
      return;
    }
    const existe = todasPiezas.some(p => normAlbaran(p.albaran) === normAlbaran(a.albaranOrigen) && refKey(p.ref) && refKey(p.ref) === refKey(a.ref));
    nuevas.push({ clave: a.clave, fechaAbono: a.fecha, descripcion: a.desc, sinIva: sin, estado: 'Sin solicitar', albaran: a.albaranOrigen,
      referencia: a.ref || '', matricula: a.matricula || '', fechaSolicitud: '', factura: a.factura,
      nota: existe ? 'La pieza existe en Piezas pero no tiene Reembolso marcado' : '' });
  });
  // Matrícula vacía en una fila existente: se rellena con la de su albarán (hueco, no pisa nada escrito).
  existentes.forEach((f, i) => {
    const m = (matriculas || {})[normAlbaran(f.albaran)];
    if (borrar.indexOf(i) < 0 && m && !String(f.matricula == null ? '' : f.matricula).trim()) cambios[i] = Object.assign(cambios[i] || {}, { matricula: m });
  });
  const fecha = f => f.fechaAbono || f.fechaSolicitud || '';
  const orden = nuevas.map((f, i) => ({ f, i })).sort((x, y) => fecha(y.f).localeCompare(fecha(x.f)) || x.i - y.i).map(x => x.f);
  return { nuevas: orden, cambios: Object.keys(cambios).map(i => ({ i: Number(i), v: cambios[i] })), borrar, registrar };
}

/**
 * Migración a la tabla editable: calcula las claves de las filas que ya había (generadas por la versión anterior),
 * emparejando cada fila con su línea de abono (factura + albarán + referencia, o descripción e importe) y su pieza marcada.
 * Devuelve un array paralelo a `filas` con la clave de cada una ('' si no se reconoce: se queda como fila manual).
 */
function clavesDeFilasAntiguas(filas, marcadas, abonos) {
  const usadasA = new Set(), usadasP = new Set();
  return filas.map(f => {
    const ks = [];
    if (String(f.factura || '').trim()) {
      const a = abonos.find(x => !usadasA.has(x.clave) && String(x.factura).trim() === String(f.factura).trim() && normAlbaran(x.albaranOrigen) === normAlbaran(f.albaran)
        && (refKey(x.ref) ? refKey(x.ref) === refKey(f.ref) : refKey(x.desc) === refKey(f.descripcion) && Math.abs(Math.abs(parseNumber(x.importe)) - parseNumber(f.sinIva)) < 0.02));
      if (a) { usadasA.add(a.clave); ks.push(a.clave); }
    }
    if (f.fechaSolicitud || f.estado === 'Sin abonar') {
      const p = marcadas.find(x => !usadasP.has(x.clave) && normAlbaran(x.albaran) === normAlbaran(f.albaran) && refODesc_(x.ref, x.desc) === refODesc_(f.ref, f.descripcion));
      if (p) { usadasP.add(p.clave); ks.push(p.clave); }
    }
    return ks.join(';');
  });
}

/**
 * Segundo escaneo de un albarán que ya existe: decide qué piezas marcar como reembolso.
 * existentes: [{ref, desc, reembolso}]  lineas: líneas de Gemini (con .reembolso)
 * Devuelve { marcar: [índices en existentes], añadir: [líneas nuevas], yaMarcadas: n }
 */
function aplicarReembolsos(existentes, lineas) {
  const usadas = new Set(), marcar = [], anadir = [];
  let yaMarcadas = 0;
  lineasParaPiezas(lineas).filter(l => l.reembolso).forEach(l => {
    const k = refKey(l.referencia), kd = refKey(l.descripcion);
    let idx = -1;
    for (let i = 0; i < existentes.length; i++) {
      if (usadas.has(i)) continue;
      const e = existentes[i];
      if ((k && refKey(e.ref) === k) || (!k && kd && refKey(e.desc) === kd)) { idx = i; break; }
    }
    if (idx < 0) { anadir.push(l); return; }
    usadas.add(idx);
    if (existentes[idx].reembolso) yaMarcadas++; else marcar.push(idx);
  });
  return { marcar, anadir, yaMarcadas };
}

/** Busca una fila manual (sin nº de albarán ni PDF) que corresponda al albarán escaneado. rows: [{row, plate, total, albaran, pdf}] */
function buscarFilaManual(rows, plate, total) {
  const p = normPlate(plate);
  if (!p) return null;
  return rows.find(r => !r.albaran && !r.pdf && normPlate(r.plate) === p && Math.abs(parseNumber(r.total) - total) < 0.02) || null;
}

/** Bajo este nº de caracteres no se busca por coincidencia parcial: demasiado ambiguo para ser útil. */
const MATRICULA_LARGO_MIN_BUSQUEDA = 3;

/**
 * Resuelve lo que se ha escrito en la celda Matrícula contra la lista de coches conocidos, buscando la
 * combinación en CUALQUIER posición (no sólo al principio, a diferencia del desplegable nativo de Sheets).
 * plate: ya pasada por normPlate. coches: [{plate, ...}] con plate ya normalizada.
 * Devuelve { tipo, candidatos }:
 *  - 'exacta'  -> `plate` ya es una matrícula real; no hay que tocar la celda.
 *  - 'unica'   -> `plate` no es exacta pero aparece en una sola matrícula de coches: se puede autocompletar con candidatos[0].
 *  - 'varias'  -> aparece en más de una: no se adivina, se informa de las candidatas.
 *  - 'ninguna' -> no aparece en ninguna (o el texto es demasiado corto para buscar): sin cambios.
 */
function resolverMatricula(plate, coches) {
  if (!plate) return { tipo: 'ninguna', candidatos: [] };
  if (coches.some(c => c.plate === plate)) return { tipo: 'exacta', candidatos: [] };
  if (plate.length < MATRICULA_LARGO_MIN_BUSQUEDA) return { tipo: 'ninguna', candidatos: [] };
  const candidatos = coches.filter(c => c.plate.indexOf(plate) >= 0);
  if (candidatos.length === 1) return { tipo: 'unica', candidatos };
  if (candidatos.length > 1) return { tipo: 'varias', candidatos };
  return { tipo: 'ninguna', candidatos: [] };
}

/**
 * Apps Script escribe las fórmulas con la sintaxis de la configuración regional de la hoja.
 * En España (es_ES) los argumentos se separan con ";" y los decimales llevan ",": =SI(A1>0,5;1;2).
 * Convierte una fórmula escrita con "," y "." (formato inglés) al formato con ";" y "," decimal. Respeta el texto entre comillas.
 */
function localizarFormula(f, puntoYComa) {
  if (!puntoYComa || typeof f !== 'string' || f.charAt(0) !== '=') return f;
  let out = '', enTexto = false;
  for (let i = 0; i < f.length; i++) {
    const ch = f.charAt(i);
    if (ch === '"') { enTexto = !enTexto; out += ch; continue; }
    if (enTexto) { out += ch; continue; }
    if (ch === ',') out += ';';
    else if (ch === '.' && /\d/.test(f.charAt(i - 1)) && /\d/.test(f.charAt(i + 1))) out += ',';
    else out += ch;
  }
  return out;
}

/** ¿Este locale (p. ej. "es_ES", "en_US", "de_CH") usa ";" como separador de argumentos? */
function usaPuntoYComa(locale) {
  const [lang, pais] = String(locale || 'en_US').split(/[_-]/);
  const idiomas = ['es', 'de', 'fr', 'it', 'pt', 'nl', 'ru', 'pl', 'tr', 'sv', 'da', 'nb', 'no', 'fi', 'cs', 'sk', 'hu', 'ro', 'bg', 'el', 'uk', 'hr', 'sl', 'sr', 'lt', 'lv', 'et', 'id', 'vi', 'ca', 'eu', 'gl'];
  if (idiomas.indexOf(lang) < 0) return false;
  const excepciones = { es: ['MX', 'US', 'PR', 'DO', 'GT', 'HN', 'NI', 'PA', 'SV'], de: ['CH', 'LI'] };
  return !(excepciones[lang] && excepciones[lang].indexOf(pais) >= 0);
}

/** Ordena/actualiza: fila en el Resumen para un mes (1-12). */
const MESES = ['ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'sep', 'oct', 'nov', 'dic'];

if (typeof module !== 'undefined') {
  module.exports = { IVA_DEFECTO, MESES, round2, parseNumber, normPlate, normAlbaran, albaranOrigen, mapaMatriculas, refKey, jobPrefix, nextJobNumber,
    pickOpenJob, isoValid, daysBetween, quincenaDe, ultimoDiaMes, rangoQuincena, esResiduo, lineasParaPiezas, validarAlbaran,
    validarFactura, periodoFactura, clavesPiezas, clavesAbonos, partirClaves, sincronizarAbonos, clavesDeFilasAntiguas, aplicarReembolsos, buscarFilaManual, resolverMatricula, localizarFormula, usaPuntoYComa };
}
