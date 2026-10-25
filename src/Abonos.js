/**
 * Pestaña Abonos: un INFORME que se rehace entero desde los datos, así que nunca se desajusta con ellos:
 *  - Piezas: la solicitud (Reembolso ✓ + Fecha reembolso) y "Abonada a mano"; cada pieza tiene un ID fijo.
 *  - Líneas RM: los abonos de RM (de las facturas quincenales y de los albaranes de abono escaneados), cada uno con el
 *    "ID pieza" de la pieza abonada. El emparejamiento se hace una sola vez (emparejarAbonos) y se guarda ahí, donde se
 *    puede corregir a mano.
 *  - Estado, precio, descripción… se calculan (informeAbonos en Logic.js). Lo que se escriba a mano en la tabla se pierde.
 * Las columnas de la tabla se buscan por el nombre de su cabecera: se pueden reordenar a mano (colsAbonos_).
 * A la derecha van el resumen por quincena y el panel de pendientes (Setup.js ▸ montarResumenAbonos_).
 */

const CAMPOS_ABONOS_ = {
  fechaAbono: 'Fecha abono', descripcion: 'Descripción pieza', sinIva: 'Precio sin IVA', conIva: 'Precio con IVA', estado: 'Estado',
  albaran: 'Nº albarán', referencia: 'Referencia', matricula: 'Matrícula', fechaSolicitud: 'Fecha solicitud', factura: 'Factura RM',
  nota: 'Nota', dias: 'Días pendiente', idPieza: 'ID pieza',
};
// Columnas con fórmula: no cuentan al comparar la tabla que hay con la que debería haber.
const FORMULAS_ABONOS_ = ['conIva', 'dias'];

let _colsAbonos = null;
/**
 * {campo: nº de columna} según la cabecera de la tabla (fila 1, columnas A-M). Hoja nueva (cabecera vacía): el orden de
 * ABONOS.cabTabla. fila1: la cabecera si ya se ha leído.
 */
function colsAbonos_(fila1) {
  if (_colsAbonos) return _colsAbonos;
  const ct = ABONOS.cabTabla;
  fila1 = fila1 || hoja_(HOJA.ABONOS).getRange(ABONOS.filaCabTabla, 1, 1, ct.length).getValues()[0];
  const cab = fila1.every(x => x === '' || x == null) ? ct : fila1.map(x => String(x).trim()), c = {};
  Object.keys(CAMPOS_ABONOS_).forEach(k => {
    const i = cab.indexOf(CAMPOS_ABONOS_[k]);
    if (i < 0) throw new Error(`Falta la columna "${CAMPOS_ABONOS_[k]}" en la tabla de Abonos (columnas A-${colLetra_(ct.length)}). Si la has renombrado o sacado de la tabla, vuelve a ponerla.`);
    c[k] = i + 1;
  });
  return (_colsAbonos = c);
}
function letraAbonos_(campo) { return colLetra_(colsAbonos_()[campo]); }

/** Rehace la tabla de Abonos. tabP: Piezas si quien llama ya la ha leído. Devuelve {filas, escrito}. */
function actualizarInformeAbonos_(tabP) {
  const crono = cronometro_('actualizarInformeAbonos_');
  const d = leerParaAbonos_(tabP, crono);
  crono.fin();
  return escribirAbonos_(d);
}

/**
 * Todo lo que el informe necesita LEER, para hacerlo antes de escribir nada: en Sheets, una lectura hecha después de una
 * escritura espera a que se recalculen las fórmulas que dependen de lo escrito.
 * tabP: la tabla de Piezas si quien llama ya la ha leído (y la mantiene al día con lo que escriba después).
 */
function leerParaAbonos_(tabP, crono) {
  if (recolocarTablaAbonos_()) { montarAbonos_(); crono.paso('recolocar tabla de Abonos'); }
  const leida = tabla => `${tabla.filas.length} de ${tabla.leidas} filas`;
  if (!tabP) { tabP = leerTabla_(HOJA.PIEZAS); crono.paso(`leer Piezas (${leida(tabP)})`); }
  const tabL = leerTabla_(HOJA.LINEAS); crono.paso(`leer Líneas RM (${leida(tabL)})`);
  const tabA = leerTabla_(HOJA.ALB); crono.paso(`leer Albaranes (${leida(tabA)})`);
  const sh = hoja_(HOJA.ABONOS), tb = ABONOS.filaTabla, ncol = ABONOS.cabTabla.length, ultFila = sh.getLastRow(), maxFilas = sh.getMaxRows();
  // getLastRow cuenta también el resumen de la derecha: se quitan las filas vacías del final de la tabla.
  const leido = sh.getRange(ABONOS.filaCabTabla, 1, Math.max(ultFila, tb) - ABONOS.filaCabTabla + 1, ncol).getValues();
  colsAbonos_(leido[0]);
  const tabla = leido.slice(tb - ABONOS.filaCabTabla);
  while (tabla.length && tabla[tabla.length - 1].every(x => x === '' || x == null)) tabla.pop();
  loc_('');  // la configuración regional también es una lectura: se guarda ya para escribir las fórmulas luego
  const desde = aISO_(cfg_('ABONOS_DESDE'));
  crono.paso(`leer Abonos (${tabla.length} filas)`);
  return { tabP, tabL, tabA, sh, tabla, maxFilas, desde };
}

/** Piezas de RM (las de proveedor "Otros" no se reclaman), en el formato de informeAbonos. */
function piezasParaAbonos_(tabP, tabA, plateDe) {
  const proveedorDe = {};
  tabA.filas.forEach(f => { const n = normAlbaran(f.v['Nº albarán']); if (n) proveedorDe[n] = f.v['Proveedor']; });
  return tabP.filas.filter(p => String(proveedorDe[normAlbaran(p.v['Nº albarán'])] || 'RM') !== 'Otros').map(p => {
    const alb = normAlbaran(p.v['Nº albarán']);
    return { id: p.v['ID'], fila: p.fila, albaran: alb, ref: p.v['Referencia pieza'], desc: p.v['Descripción'], sinIva: p.v['Precio descontado sin IVA'],
      solicitada: p.v['Reembolso'] === true && !!alb, fechaSolicitud: aISO_(p.v['Fecha reembolso']), abonadaAMano: aISO_(p.v['Abonada a mano']),
      matricula: p.v['Matrícula'] || plateDe[alb] || '' };
  });
}

/**
 * {nº albarán: matrícula}: primero la pestaña Albaranes (la que se corrige a mano), después las líneas de compra de las
 * facturas RM, que traen la matrícula de TODOS los albaranes aunque no se hayan escaneado. Se copia tal cual la pone RM.
 */
function matriculasDeAlbaranes_(tabA, tabL) {
  const pares = tabA.filas.map(f => [f.v['Nº albarán'], f.v['Matrícula']]);
  if (tabL) tabL.filas.forEach(f => { if (f.v['Tipo'] === 'Compra') pares.push([f.v['Nº albarán'], f.v['Matrícula']]); });
  return mapaMatriculas(pares);
}

/** Líneas de abono de Líneas RM con su clave A (Logic.js ▸ clavesAbonos), en el orden de la pestaña. */
function lineasAbono_(tabL) {
  return clavesAbonos(tabL.filas.filter(f => f.v['Tipo'] === 'Abono').map(f => ({
    f, albaranAbono: normAlbaran(f.v['Nº albarán']), albaranOrigen: albaranOrigen(f.v['Albarán origen']), ref: f.v['Referencia'], desc: f.v['Descripción'],
  })));
}

/**
 * Aplica lo leído en leerParaAbonos_. No lee nada de la hoja. Escribe sólo lo que cambia: IDs de piezas nuevas, la fecha
 * de reembolso que falte, el "ID pieza" de los abonos recién emparejados y la tabla, si no es ya la que debe ser.
 */
function escribirAbonos_(d) {
  const crono = cronometro_('escribirAbonos_');
  const { tabP, tabL, tabA, sh } = d, hoy = hoyISO_(), tb = ABONOS.filaTabla, ncol = ABONOS.cabTabla.length;
  const ids = asegurarIdsPiezas_(tabP);
  const plateDe = matriculasDeAlbaranes_(tabA, tabL);
  const piezas = piezasParaAbonos_(tabP, tabA, plateDe);
  // Marcada sin pasar por la edición de la celda (p. ej. pegando): su fecha de reembolso es hoy.
  const sinFecha = piezas.filter(p => p.solicitada && !p.fechaSolicitud);
  sinFecha.forEach(p => { p.fechaSolicitud = hoy; actualizarFila_(tabP, p.fila, { 'Fecha reembolso': aFecha_(hoy) }); });
  crono.paso(`preparar piezas (${piezas.filter(p => p.solicitada).length} pedidas, ${ids} IDs nuevos, ${sinFecha.length} fechas escritas)`);

  const lineas = lineasAbono_(tabL);
  const abonos = lineas.map(({ f, albaranAbono, albaranOrigen: orig, ref, desc }) => ({
    idPieza: f.v['ID pieza'], fecha: aISO_(f.v['Fecha abono']) || aISO_(f.v['Fecha albarán']), factura: String(f.v['Nº factura'] || '').trim(),
    albaranAbono, albaranOrigen: orig, ref, desc, importe: parseNumber(f.v['Importe sin IVA']), matricula: plateDe[orig] || f.v['Matrícula'],
  }));
  const nuevos = emparejarAbonos(piezas, abonos);
  nuevos.forEach(({ i, id }) => { abonos[i].idPieza = id; actualizarFila_(tabL, lineas[i].f.fila, { 'ID pieza': id }); });
  const filas = informeAbonos(piezas, abonos, plateDe, d.desde);
  crono.paso(`cruce (${abonos.length} abonos, ${nuevos.length} emparejados ahora)`);

  const valores = filas.map((f, j) => filaAbono_(f, tb + j));
  const escrito = !mismaTablaAbonos_(d.tabla, valores);
  if (escrito) {
    if (tb + valores.length - 1 > d.maxFilas) sh.insertRowsAfter(d.maxFilas, tb + valores.length - 1 - d.maxFilas + 100);  // al final: no mueve nada
    if (valores.length) sh.getRange(tb, 1, valores.length, ncol).setValues(valores);
    // Sólo las celdas de la tabla (A-M): a la derecha, en las mismas filas, está el resumen.
    if (d.tabla.length > valores.length) sh.getRange(tb + valores.length, 1, d.tabla.length - valores.length, ncol).clearContent();
  }
  // Sólo a Ejecuciones: escribirlo en Registro en cada edición costaba ~300 ms y llenaba Registro de filas INFO.
  crono.fin(`${escrito ? 'tabla rehecha' : 'tabla ya al día'} | ` + ESTADOS_ABONO.map(e => e + ' ' + filas.filter(f => f.estado === e).length).join(', '));
  return { filas, escrito };
}

function filaAbono_(f, r) {
  const c = colsAbonos_(), L = letraAbonos_, fila = new Array(ABONOS.cabTabla.length).fill('');
  // "Precio con IVA" y "Días pendiente" son fórmulas: la primera usa el IVA de Config y la segunda se actualiza sola cada día.
  const v = { fechaAbono: aFecha_(f.fechaAbono), descripcion: f.descripcion, sinIva: f.sinIva, conIva: loc_(`=ROUND($${L('sinIva')}${r}*(1+IVA),2)`),
    estado: f.estado, albaran: f.albaran, referencia: f.referencia, matricula: f.matricula, fechaSolicitud: aFecha_(f.fechaSolicitud),
    factura: f.factura, nota: f.nota, dias: loc_(`=IF($${L('estado')}${r}="Sin abonar",TODAY()-$${L('fechaSolicitud')}${r},"")`), idPieza: f.idPieza };
  Object.keys(v).forEach(k => { fila[c[k] - 1] = v[k]; });
  return fila;
}

/** ¿La tabla leída ya es la que se va a escribir? Compara sólo los datos (no las fórmulas). */
function mismaTablaAbonos_(leida, nueva) {
  if (leida.length !== nueva.length) return false;
  const c = colsAbonos_(), cols = Object.keys(CAMPOS_ABONOS_).filter(k => FORMULAS_ABONOS_.indexOf(k) < 0).map(k => c[k] - 1);
  const norm = x => (x instanceof Date ? aISO_(x) : typeof x === 'number' ? String(round2(x)) : String(x == null ? '' : x).trim());
  return leida.every((fila, i) => cols.every(j => norm(fila[j]) === norm(nueva[i][j])));
}

/**
 * Líneas de un albarán de ABONO escaneado → Líneas RM, sin Nº factura (la quincenal lo traerá después): fecha de abono =
 * la del albarán de abono. Sin la línea de residuos (SIGAUS): esa llega con la factura. Las que ya estén (albarán
 * escaneado dos veces, o su factura ya procesada) no se repiten. Devuelve cuántas ha añadido.
 */
function guardarAbonosEscaneados_(tabL, tabA, doc, hoy) {
  const num = normAlbaran(doc.numero_albaran), fecha = aFecha_(doc.fecha || hoy), plateDe = matriculasDeAlbaranes_(tabA, tabL);
  const ya = new Set(lineasAbono_(tabL).map(x => x.clave));
  const nuevas = clavesAbonos(doc.lineas.filter(l => !esResiduo(l)).map(l => ({ l, albaranAbono: num, albaranOrigen: albaranOrigen(l.albaran_origen), ref: l.referencia, desc: l.descripcion })));
  const filas = nuevas.filter(a => !ya.has(a.clave)).map(({ l, albaranOrigen: orig }) => ({
    'Nº factura': '', 'Nº albarán': num, 'Fecha albarán': fecha, 'Tipo': 'Abono', 'Matrícula': plateDe[orig] || '', 'Albarán origen': orig,
    'Referencia': l.referencia, 'Descripción': l.descripcion, 'Cantidad': l.cantidad == null ? '' : l.cantidad, 'Precio': l.precio_unitario == null ? '' : l.precio_unitario,
    'Descuento': l.descuento_pct == null ? '' : l.descuento_pct / 100, 'Importe sin IVA': l.importe == null ? 0 : l.importe, 'Fecha abono': fecha,
  }));
  agregarFilas_(tabL, filas);
  return filas.length;
}

/**
 * Edición a mano en la tabla de Abonos: se vuelve a poner lo que toca (es un informe). El resumen y el panel de la
 * derecha (el año, por ejemplo) sí se pueden cambiar.
 */
function editarAbonos_(r0, n, c0) {
  if (c0 > ABONOS.cabTabla.length || r0 + n <= ABONOS.filaTabla) return;
  if (actualizarInformeAbonos_().escrito) toast_('Abonos se rellena solo: cambia el dato en Piezas (o el "ID pieza" de la línea en Líneas RM).', 'Abonos', 8);
}

/**
 * La tabla de Abonos va arriba del todo (cabecera en la fila 1). Si alguien inserta filas enteras encima, se quitan
 * (sólo si están vacías en la tabla) y devuelve true: quien llama debe volver a montar la pestaña (montarAbonos_),
 * porque el resumen de la derecha también se ha movido. Nunca borra datos de la tabla.
 */
function recolocarTablaAbonos_() {
  const sh = hoja_(HOJA.ABONOS), ct = ABONOS.cabTabla;
  const n = Math.min(Math.max(sh.getLastRow(), 1), 80);
  const vals = sh.getRange(1, 1, n, ct.length).getValues();
  const vacia = r => r.every(x => x === '' || x == null);
  const idx = vals.findIndex(r => r.indexOf(CAMPOS_ABONOS_.fechaAbono) >= 0 && r.indexOf(CAMPOS_ABONOS_.estado) >= 0);
  if (idx < 0) {
    if (vals.every(vacia)) return false;  // hoja nueva: aún no hay tabla
    throw new Error('No encuentro la cabecera de la tabla de Abonos ("Fecha abono" … "Estado" en la fila 1). ' +
      'Si se ha borrado, deshazlo (Ctrl+Z) o recupera la fila desde Archivo ▸ Historial de versiones.');
  }
  const real = idx + 1;
  if (real === ABONOS.filaCabTabla) return false;
  if (!vals.slice(0, real - 1).every(vacia)) {
    throw new Error(`En Abonos hay datos en las filas 1-${real - 1}, encima de la cabecera de la tabla. Muévelos o bórralos para que la cabecera quede en la fila 1.`);
  }
  sh.deleteRows(1, real - 1);
  log_('AVISO', 'recolocarTablaAbonos', `${HOJA.ABONOS}!A${real}`, `La cabecera de la tabla estaba en la fila ${real}: quitadas las filas vacías de encima.`);
  return true;
}

/** Filas enteras insertadas o borradas a mano en Abonos: el resumen vuelve a su sitio y la tabla se rehace. */
function revisarAbonosTrasCambioDeFilas_() {
  if (recolocarTablaAbonos_()) montarAbonos_();
  else if (!resumenAbonosEnSuSitio_(hoja_(HOJA.ABONOS))) montarResumenAbonos_();
  actualizarInformeAbonos_();
}

/** Menú: Actualizar Abonos. */
function actualizarAbonos() {
  ejecutar_('actualizarAbonos', () => conBloqueo_(10, () => {
    actualizarHoja_(false);
    const r = actualizarInformeAbonos_();
    const pendientes = r.filas.filter(f => f.estado === 'Sin abonar').length;
    toast_(`${r.escrito ? 'Abonos actualizado' : 'Abonos ya estaba al día'}: ${r.filas.length} fila(s), ${pendientes} sin abonar.`);
  }));
}
