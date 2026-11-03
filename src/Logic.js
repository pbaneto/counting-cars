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

/**
 * Un albarán deja su trabajo porque se ha corregido la matrícula: qué se hace con el trabajo viejo.
 *  - 'dejar': aún tiene albaranes, o tiene una factura o un pagado que no tienen adónde ir.
 *  - 'renombrar': la matrícula no tiene trabajo abierto: el viejo pasa a serlo, con su factura, pagado y mes.
 *  - 'borrar': no tiene nada que conservar.
 *  - 'mover': su factura y pagado pasan al trabajo abierto de la matrícula, y se borra.
 */
function destinoTrabajoViejo({ quedan, conCobro, hayMatricula, hayAbierto, abiertoConFactura }) {
  if (quedan > 0) return 'dejar';
  if (!hayMatricula) return conCobro ? 'dejar' : 'borrar';
  if (!hayAbierto) return 'renombrar';
  if (!conCobro) return 'borrar';
  return abiertoConFactura ? 'dejar' : 'mover';
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
function mesDe(iso) { return MESES[Number(iso.slice(5, 7)) - 1]; }

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
  if (!Number.isFinite(total) || total === 0 || (total < 0 && !doc.es_abono)) errors.push('No se ha podido leer el total del albarán');
  if (doc.totalCalculado) warnings.push(AVISO_SIN_TOTAL);
  if (!doc.numero_albaran) warnings.push('Sin nº de albarán');
  if (!isoValid(doc.fecha)) warnings.push('Fecha ilegible: se usa la fecha de hoy');
  if (!normPlate(doc.matricula)) warnings.push('Sin matrícula');
  // Base × IVA = total y suma de líneas = base: ver cuadreAlbaran (si no cuadran, el albarán no se añade).
  const lineas = doc.lineas || [];
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

const AVISO_SIN_TOTAL = '⚠ Sin total en el albarán: importe calculado sumando las líneas + IVA (puede faltar una 2ª hoja)';

/**
 * Albarán con líneas pero sin total impreso (casillas de abajo en blanco): base = suma de las líneas, total = base + IVA,
 * con totalCalculado para avisar. Si alguna línea no tiene importe no se calcula: la suma saldría corta sin saberlo.
 */
function completarTotal(doc, iva) {
  iva = iva == null ? IVA_DEFECTO : iva;
  const lineas = doc.lineas || [];
  if (doc.es_abono || parseNumber(doc.total) || !lineas.length || !lineas.every(l => Number.isFinite(parseNumber(l.importe)))) return doc;
  const base = round2(lineas.reduce((a, l) => a + parseNumber(l.importe), 0));
  if (base <= 0) return doc;
  return Object.assign({}, doc, { base_imponible: base, total: round2(base * (1 + iva)), totalCalculado: true });
}

/**
 * ¿Este escaneo de un albarán ya registrado es la hoja que le faltaba? Devuelve:
 *  - 'completa': la fila lleva el aviso de sin total, el escaneo trae total impreso y precio guardado + sus líneas con IVA = ese total;
 *  - 'no cuadra': igual, pero las cuentas no dan;
 *  - '': no es una 2ª hoja (reescaneo de una 1ª hoja o de un albarán normal): se trata como siempre.
 */
function hojaQueFalta(precioGuardado, nota, doc, iva) {
  iva = iva == null ? IVA_DEFECTO : iva;
  if (String(nota || '').indexOf(AVISO_SIN_TOTAL) < 0 || doc.totalCalculado || !parseNumber(doc.total)) return '';
  const nuevas = (doc.lineas || []).reduce((a, l) => a + (Number.isFinite(parseNumber(l.importe)) ? parseNumber(l.importe) : 0), 0);
  return Math.abs(parseNumber(precioGuardado) + nuevas * (1 + iva) - parseNumber(doc.total)) <= 0.05 ? 'completa' : 'no cuadra';
}

function quitarAvisoSinTotal(nota) {
  return String(nota || '').replace(AVISO_SIN_TOTAL, '').replace(/^[\s;|]+/, '').replace(/[\s;|]+$/, '').replace(/;\s*;/, ';');
}

/**
 * ¿Cuadran los importes del albarán leído? Devuelve la lista de problemas ([] = cuadra):
 *  - base imponible × (1 + IVA) = total (±0,05 €);
 *  - suma de los importes de las líneas = base imponible (o total / (1 + IVA) si no se lee la base) (±0,05 €).
 * Si no hay ni base ni líneas no se puede comprobar nada y se da por bueno.
 */
function cuadreAlbaran(doc, iva) {
  iva = iva == null ? IVA_DEFECTO : iva;
  const problemas = [], total = parseNumber(doc.total), base = parseNumber(doc.base_imponible), lineas = doc.lineas || [];
  const eur = x => String(round2(x)).replace('.', ',');
  if (Number.isFinite(total) && Number.isFinite(base) && Math.abs(round2(base * (1 + iva)) - total) > 0.05) {
    problemas.push(`base ${eur(base)} + IVA = ${eur(base * (1 + iva))}, pero el total es ${eur(total)}`);
  }
  if (lineas.length && Number.isFinite(total)) {
    const suma = round2(lineas.reduce((a, l) => a + (Number.isFinite(parseNumber(l.importe)) ? parseNumber(l.importe) : 0), 0));
    const ref = Number.isFinite(base) ? base : round2(total / (1 + iva));
    if (Math.abs(suma - ref) > 0.05) problemas.push(`las líneas suman ${eur(suma)}, pero la base es ${eur(ref)}`);
  }
  return problemas;
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
 * Clave de una línea de un albarán de ABONO de RM: A|albarán de abono|referencia|n (n = nº de orden entre las líneas
 * iguales). La misma clave sale del albarán de abono escaneado y de la factura quincenal que luego lo incluye (en la
 * factura, el "Nº albarán" del bloque de abono), así que llegue primero uno u otro, la línea no se duplica. No lleva el
 * albarán original: Gemini lo lee mal a veces en el escaneo (4740419 por 470419).
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
function clavesAbonos(abonos) { return ponerClaves(abonos, 'A', a => [normAlbaran(a.albaranAbono), refODesc_(a.ref, a.desc)]); }

/**
 * Líneas de abono ya guardadas que la lectura de una factura se ha saltado (Gemini pierde a veces las que siguen en la hoja
 * siguiente). Son de la factura si su albarán de abono es un bloque de ella y con ellas el bloque da justo su importe impreso.
 *  bloques: [{ numero, importe, leidas: [claves de las líneas leídas], suma: importe de las leídas }]
 *  conocidas: [{ f, clave, albaranAbono, importe }]
 * Devuelve { [nº del bloque]: [conocidas que le faltaban] }.
 */
function lineasOmitidasDeBloques(bloques, conocidas) {
  const out = {};
  bloques.forEach(b => {
    const faltan = conocidas.filter(c => normAlbaran(c.albaranAbono) === normAlbaran(b.numero) && b.leidas.indexOf(c.clave) < 0);
    const suma = faltan.reduce((a, c) => a + parseNumber(c.importe), b.suma);
    if (faltan.length && Math.abs(suma - parseNumber(b.importe)) <= 0.03) out[normAlbaran(b.numero)] = faltan;
  });
  return out;
}

/** "ID pieza" de una línea de abono: "NO" = no corresponde a ninguna pieza (no se empareja sola). */
const SIN_PIEZA = 'NO';
function idPiezaDe(v) { const s = String(v == null ? '' : v).trim(); return s.toUpperCase() === SIN_PIEZA ? SIN_PIEZA : s; }

/**
 * Empareja cada línea de abono SIN "ID pieza" con su pieza: mismo albarán (el original del abono) y misma referencia;
 * si no, una pieza pedida con el mismo importe; si no, la ÚNICA pieza pedida con esa referencia en otro albarán, pedida
 * antes del abono (RM a veces abona sobre otro albarán, o sobre ninguno: "NO CODIFICA PUESTO EL ORIGINAL"). Cada pieza
 * recibe un solo abono. Un ID ya puesto (por el script o a mano) no se cambia nunca, aunque la pieza ya no exista.
 *  piezas: [{id, albaran, ref, sinIva, solicitada, fechaSolicitud}]   abonos: [{idPieza, albaranOrigen, ref, importe, fecha}]
 * Devuelve [{i (índice en abonos), id}] con los emparejamientos nuevos.
 */
function emparejarAbonos(piezas, abonos) {
  const usadas = new Set(abonos.map(a => idPiezaDe(a.idPieza)).filter(Boolean));
  const mismoImporte = (x, y) => Math.abs(Math.abs(parseNumber(x)) - Math.abs(parseNumber(y))) < 0.02;
  const nuevos = [];
  abonos.map((a, i) => ({ a, i })).filter(x => !idPiezaDe(x.a.idPieza))
    .sort((x, y) => String(x.a.fecha || '').localeCompare(String(y.a.fecha || '')) || x.i - y.i)
    .forEach(({ a, i }) => {
      const libres = piezas.filter(p => !usadas.has(String(p.id)) && normAlbaran(p.albaran) === normAlbaran(a.albaranOrigen));
      const porRef = libres.filter(p => refKey(a.ref) && refKey(p.ref) === refKey(a.ref));
      let p = porRef.find(x => x.solicitada) || porRef[0] || libres.find(x => x.solicitada && mismoImporte(x.sinIva, a.importe));
      if (!p && refKey(a.ref) && a.fecha) {
        const otras = piezas.filter(x => !usadas.has(String(x.id)) && x.solicitada && refKey(x.ref) === refKey(a.ref) && x.fechaSolicitud && x.fechaSolicitud <= a.fecha);
        if (otras.length === 1) p = otras[0];
      }
      if (!p) return;
      usadas.add(String(p.id));
      nuevos.push({ i, id: p.id });
    });
  return nuevos;
}

/**
 * Filas de la tabla de Abonos, calculadas de cero (la tabla no se edita: se rehace entera cada vez).
 *  - una por línea de abono: "Abonada" si su pieza está pedida (Reembolso o Abonada a mano); si no, "Sin solicitar";
 *  - una por pieza pedida sin abono: "Abonada (a mano)" si tiene fecha en "Abonada a mano"; si no, "Sin abonar".
 *  piezas: [{id, albaran, ref, desc, sinIva, solicitada, fechaSolicitud, abonadaAMano, matricula}] (sólo las de RM)
 *  abonos: [{idPieza, fecha, factura, albaranOrigen, ref, desc, importe, matricula}]
 *  matriculas: {nº albarán: matrícula} para las filas sin matrícula.
 *  desde: fecha ISO; las filas anteriores (fecha de abono o, si no hay, de solicitud) no salen, salvo las "Sin abonar".
 * Lo más reciente arriba.
 */
function informeAbonos(piezas, abonos, matriculas, desde) {
  matriculas = matriculas || {};
  const porId = {};
  piezas.forEach(p => { porId[String(p.id)] = p; });
  const pedida = p => p.solicitada || !!p.abonadaAMano;
  const filas = [], abonadas = new Set();
  abonos.forEach(a => {
    const id = idPiezaDe(a.idPieza), p = id && id !== SIN_PIEZA ? porId[id] : null;
    const sin = round2(Math.abs(parseNumber(a.importe) || 0));
    const notas = [];
    if (p && pedida(p)) {
      const pedido = round2(parseNumber(p.sinIva));
      if (Number.isFinite(pedido) && Math.abs(pedido - sin) > 0.02) notas.push(`Importe abonado distinto: se pidió ${pedido} sin IVA`);
    } else if (p) notas.push('La pieza existe en Piezas pero no tiene Reembolso marcado');
    else if (id && id !== SIN_PIEZA) notas.push(`La pieza ${id} ya no está en Piezas`);
    if (p && normAlbaran(a.albaranOrigen) !== normAlbaran(p.albaran)) notas.push(a.albaranOrigen ? `RM lo abona sobre el albarán ${normAlbaran(a.albaranOrigen)}` : 'RM lo abona sin decir el albarán');
    const nota = notas.join('; ');
    if (p) abonadas.add(String(p.id));
    const albaran = normAlbaran(p ? p.albaran : a.albaranOrigen);
    filas.push({ idPieza: p ? p.id : '', fechaAbono: a.fecha || '', descripcion: (p && p.desc) || a.desc || '', sinIva: sin,
      estado: p && pedida(p) ? 'Abonada' : 'Sin solicitar', albaran, referencia: (p && p.ref) || a.ref || '',
      matricula: (p && normPlate(p.matricula)) || matriculas[albaran] || normPlate(a.matricula),
      fechaSolicitud: p && pedida(p) ? p.fechaSolicitud || '' : '', factura: a.factura || '', nota });
  });
  piezas.filter(p => pedida(p) && !abonadas.has(String(p.id))).forEach(p => {
    const albaran = normAlbaran(p.albaran), sin = round2(parseNumber(p.sinIva));
    filas.push({ idPieza: p.id, fechaAbono: p.abonadaAMano || '', descripcion: p.desc || '', sinIva: Number.isFinite(sin) ? sin : '',
      estado: p.abonadaAMano ? 'Abonada (a mano)' : 'Sin abonar', albaran, referencia: p.ref || '',
      matricula: normPlate(p.matricula) || matriculas[albaran] || '', fechaSolicitud: p.fechaSolicitud || '', factura: '', nota: '' });
  });
  // Misma fecha: lo que va después en Piezas o en Líneas RM (lo último que se ha añadido) va antes.
  const fecha = f => f.fechaAbono || f.fechaSolicitud || '';
  return filas.map((f, i) => ({ f, i })).filter(({ f }) => !desde || f.estado === 'Sin abonar' || !fecha(f) || fecha(f) >= desde).sort((x, y) => fecha(y.f).localeCompare(fecha(x.f)) || y.i - x.i).map(x => x.f);
}

/**
 * ¿Lo escrito a mano junto a una línea del albarán es la marca de reembolso? Sólo una R sola (a veces con un guion
 * delante: "—R"). Un nombre ("CHEMA"), una matrícula ("MV821"), números o "R10" no lo son.
 */
function esMarcaReembolso(texto) { return /^[-—–\s.]*R[\s.]*$/i.test(String(texto == null ? '' : texto)); }

/** Nº de letras que hay que cambiar, añadir o quitar para pasar de un texto a otro (Levenshtein). */
function distanciaTexto(a, b) {
  let fila = Array.from({ length: b.length + 1 }, (_, j) => j);
  for (let i = 1; i <= a.length; i++) {
    const nueva = [i];
    for (let j = 1; j <= b.length; j++) nueva[j] = Math.min(fila[j] + 1, nueva[j - 1] + 1, fila[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
    fila = nueva;
  }
  return fila[b.length];
}

/**
 * Segundo escaneo de un albarán que ya existe: decide qué piezas marcar como reembolso. Todas son de ese albarán.
 * Una línea con R es la pieza de su misma referencia. Si no la hay (la referencia se ha podido leer mal), la de su mismo
 * importe; entre varias con ese importe, la de referencia más parecida. Si empatan no se elige: mejor preguntar que marcar otra.
 * existentes: [{ref, desc, importe, reembolso}]  lineas: líneas de Gemini (con .reembolso)
 * Devuelve { marcar: [índices en existentes], anadir: [líneas nuevas], dudosas: [líneas sin pieza clara], yaMarcadas: n }
 */
function aplicarReembolsos(existentes, lineas) {
  const usadas = new Set(), marcar = [], anadir = [], dudosas = [];
  let yaMarcadas = 0;
  const libres = cumple => existentes.map((_, i) => i).filter(i => !usadas.has(i) && cumple(existentes[i]));
  const asignar = idx => {
    usadas.add(idx);
    if (existentes[idx].reembolso) yaMarcadas++; else marcar.push(idx);
  };
  // Primero las referencias exactas, para que una línea mal leída no se quede la pieza de otra.
  lineasParaPiezas(lineas).filter(l => l.reembolso).filter(l => {
    const k = refKey(l.referencia), exacta = k ? libres(e => refKey(e.ref) === k) : [];
    if (exacta.length) asignar(exacta[0]);
    return !exacta.length;
  }).forEach(l => {
    const k = refKey(l.referencia), kd = refKey(l.descripcion), imp = parseNumber(l.importe);
    let cand = libres(e => Math.abs(parseNumber(e.importe) - imp) < 0.01);
    if (!cand.length && !k && kd) cand = libres(e => refKey(e.desc) === kd);
    if (!cand.length) { anadir.push(l); return; }
    const dist = cand.map(i => distanciaTexto(k || kd, k ? refKey(existentes[i].ref) : refKey(existentes[i].desc)));
    const mejores = cand.filter((_, n) => dist[n] === Math.min.apply(null, dist));
    if (mejores.length > 1) dudosas.push(l); else asignar(mejores[0]);
  });
  return { marcar, anadir, dudosas, yaMarcadas };
}

/** Busca una fila manual (sin nº de albarán ni PDF) que corresponda al albarán escaneado. rows: [{row, plate, total, albaran, pdf}] */
function buscarFilaManual(rows, plate, total) {
  const p = normPlate(plate);
  if (!p) return null;
  return rows.find(r => !r.albaran && !r.pdf && normPlate(r.plate) === p && Math.abs(parseNumber(r.total) - total) < 0.02) || null;
}

/**
 * Albarán ya registrado del que este parece el mismo con el nº mal leído: misma matrícula, fecha e importe, y un solo dígito distinto.
 * El importe separa los albaranes seguidos de un mismo coche; con más dígitos distintos puede ser una compra repetida de verdad.
 * rows: [{ fila, albaran, plate, fecha, total }]
 */
function buscarNumeroMalLeido(rows, num, plate, fecha, total) {
  const n = normAlbaran(num), p = normPlate(plate);
  if (!n || !p || !isoValid(fecha)) return null;
  const unDigito = a => a.length === n.length && [...a].filter((c, i) => c !== n[i]).length === 1;
  return rows.find(r => unDigito(normAlbaran(r.albaran)) && normPlate(r.plate) === p && r.fecha === fecha && Math.abs(parseNumber(r.total) - total) < 0.02) || null;
}

/** Bajo este nº de caracteres no se busca por coincidencia parcial: demasiado ambiguo para ser útil. */
const MATRICULA_LARGO_MIN_BUSQUEDA = 3;
// Para añadirla sola a Coches: más corta suele ser un trozo o un error de escritura ("4", "DARE").
const MATRICULA_LARGO_MIN_COCHE = 5;

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

/** Compara dos versiones "1.2.3": negativo si a < b, 0 si iguales, positivo si a > b. */
function compararVersiones(a, b) {
  const pa = String(a).split('.').map(Number), pb = String(b).split('.').map(Number);
  for (let i = 0; i < 3; i++) { const d = (pa[i] || 0) - (pb[i] || 0); if (d) return d; }
  return 0;
}

/** ¿Hay que volver a aplicar el diseño? Sí si cambia la versión mayor o la menor (no por un parche). */
function cambiaDiseno(a, b) {
  const pa = String(a).split('.'), pb = String(b).split('.');
  return pa[0] !== pb[0] || pa[1] !== pb[1];
}

/** Ordena/actualiza: fila en el Resumen para un mes (1-12). */
const MESES = ['ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'sep', 'oct', 'nov', 'dic'];

if (typeof module !== 'undefined') {
  module.exports = { IVA_DEFECTO, MESES, round2, parseNumber, normPlate, normAlbaran, albaranOrigen, mapaMatriculas, refKey, jobPrefix, nextJobNumber,
    pickOpenJob, destinoTrabajoViejo, isoValid, daysBetween, quincenaDe, mesDe, MATRICULA_LARGO_MIN_COCHE, ultimoDiaMes, rangoQuincena, esResiduo, lineasParaPiezas, validarAlbaran, AVISO_SIN_TOTAL, completarTotal, hojaQueFalta, quitarAvisoSinTotal, cuadreAlbaran,
    validarFactura, periodoFactura, clavesAbonos, lineasOmitidasDeBloques, SIN_PIEZA, idPiezaDe, emparejarAbonos, informeAbonos, esMarcaReembolso, aplicarReembolsos, buscarFilaManual, buscarNumeroMalLeido, resolverMatricula, localizarFormula, usaPuntoYComa, compararVersiones, cambiaDiseno };
}
