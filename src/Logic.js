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
 * Construye las filas de la tabla grande de Abonos.
 *  piezas:      piezas de RM con Reembolso marcado  {albaran, ref, desc, sinIva, fechaReembolso, matricula}
 *  todasPiezas: todas las piezas RM (para avisar si un abono corresponde a una pieza sin marcar) {albaran, ref}
 *  abonos:      líneas de abono de facturas RM      {factura, fecha, albaranOrigen, ref, desc, importe (negativo o positivo), matricula}
 * Emparejado: mismo albarán + misma referencia; si no, mismo albarán + mismo importe.
 */
function construirAbonos(piezas, todasPiezas, abonos, iva) {
  iva = iva == null ? IVA_DEFECTO : iva;
  const usadas = new Set(), filas = [];
  const conIva = x => round2(x * (1 + iva));
  const ordenados = abonos.map((a, i) => Object.assign({ _i: i }, a)).sort((a, b) => String(a.fecha).localeCompare(String(b.fecha)) || a._i - b._i);
  const casadas = [];
  const buscar = (a, porRef) => {
    for (let i = 0; i < piezas.length; i++) {
      if (usadas.has(i) || normAlbaran(piezas[i].albaran) !== normAlbaran(a.albaranOrigen)) continue;
      if (porRef ? (refKey(piezas[i].ref) && refKey(piezas[i].ref) === refKey(a.ref))
                 : Math.abs(Math.abs(parseNumber(a.importe)) - parseNumber(piezas[i].sinIva)) < 0.02) return i;
    }
    return -1;
  };
  ordenados.forEach(a => {
    let i = buscar(a, true);
    if (i < 0) i = buscar(a, false);
    if (i >= 0) usadas.add(i);
    casadas.push([a, i]);
  });
  casadas.forEach(([a, i]) => {
    const sin = round2(Math.abs(parseNumber(a.importe)));
    const fila = {
      fechaAbono: a.fecha, descripcion: a.desc, sinIva: sin, conIva: conIva(sin), estado: i >= 0 ? 'Abonada' : 'Sin solicitar',
      albaran: a.albaranOrigen, referencia: a.ref || '', matricula: a.matricula || '',
      fechaSolicitud: i >= 0 ? (piezas[i].fechaReembolso || '') : '', factura: a.factura, nota: '',
    };
    if (i >= 0) {
      const pedido = round2(parseNumber(piezas[i].sinIva));
      if (Math.abs(pedido - sin) > 0.02) fila.nota = `Importe abonado distinto: se pidió ${pedido} sin IVA`;
      if (!fila.matricula) fila.matricula = piezas[i].matricula || '';
    } else if (todasPiezas.some(p => normAlbaran(p.albaran) === normAlbaran(a.albaranOrigen) && refKey(p.ref) && refKey(p.ref) === refKey(a.ref))) {
      fila.nota = 'La pieza existe en Piezas pero no tiene Reembolso marcado';
    }
    filas.push(fila);
  });
  piezas.forEach((p, i) => {
    if (usadas.has(i)) return;
    const sin = round2(parseNumber(p.sinIva));
    filas.push({
      fechaAbono: '', descripcion: p.desc, sinIva: sin, conIva: conIva(sin), estado: 'Sin abonar',
      albaran: p.albaran, referencia: p.ref || '', matricula: p.matricula || '',
      fechaSolicitud: p.fechaReembolso || '', factura: '', nota: '',
    });
  });
  const clave = f => f.fechaAbono || f.fechaSolicitud || '9999';
  return filas.map((f, i) => Object.assign({ _i: i }, f)).sort((a, b) => clave(a).localeCompare(clave(b)) || a._i - b._i)
    .map(f => { delete f._i; return f; });
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
  module.exports = { IVA_DEFECTO, MESES, round2, parseNumber, normPlate, normAlbaran, albaranOrigen, refKey, jobPrefix, nextJobNumber,
    pickOpenJob, isoValid, daysBetween, quincenaDe, ultimoDiaMes, rangoQuincena, esResiduo, lineasParaPiezas, validarAlbaran,
    validarFactura, periodoFactura, construirAbonos, aplicarReembolsos, buscarFilaManual, localizarFormula, usaPuntoYComa };
}
