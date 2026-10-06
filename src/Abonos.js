/**
 * Pestaña Abonos.
 *  - Resumen por quincena (filas 4-27, fórmulas creadas en setup).
 *  - Tabla grande (cabecera en la fila 30): EDITABLE, nunca se regenera. Cada sincronización sólo:
 *      · añade ARRIBA las filas nuevas (pieza con Reembolso ✓ → "Sin abonar"; abono de factura RM sin pedir → "Sin solicitar"),
 *      · rellena las celdas VACÍAS de la fila de una pieza cuando llega su abono (y "Sin abonar" → "Abonada"),
 *      · borra la fila de una pieza que se desmarca en Piezas.
 *    Y al revés: borrar a mano una fila de pieza aquí la desmarca en Piezas (alCambiar / editarAbonos_).
 *    Cada fila lleva sus claves en la columna oculta "Clave" (Logic.js ▸ clavesPiezas / clavesAbonos).
 *  Las columnas de la tabla se buscan por el nombre de su cabecera: se pueden reordenar a mano (colsAbonos_).
 */

const CAMPOS_ABONOS_ = {
  fechaAbono: 'Fecha abono', descripcion: 'Descripción pieza', sinIva: 'Precio sin IVA', conIva: 'Precio con IVA', estado: 'Estado',
  albaran: 'Nº albarán', referencia: 'Referencia', matricula: 'Matrícula', fechaSolicitud: 'Fecha solicitud', factura: 'Factura RM',
  nota: 'Nota', dias: 'Días pendiente', clave: 'Clave',
};
// Campos que escribe una persona o la sincronización (no las fórmulas): si están todos vacíos, la fila se ha borrado.
const ENTRADAS_ABONOS_ = ['fechaAbono', 'descripcion', 'sinIva', 'estado', 'albaran', 'referencia', 'matricula', 'fechaSolicitud', 'factura', 'nota'];

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

/** escaneados: líneas de albaranes de ABONO recién escaneados (ver abonosEscaneados_), que aún no están en ninguna factura. */
function sincronizarAbonos_(escaneados) {
  const crono = cronometro_('sincronizarAbonos_');
  const d = leerParaAbonos_(null, crono);
  crono.fin();
  return escribirAbonos_(d, null, escaneados);
}

/**
 * Todo lo que la sincronización necesita LEER, para hacerlo antes de escribir nada: en Sheets, una lectura hecha
 * después de una escritura espera a que se recalculen las fórmulas que dependen de lo escrito.
 * tabP: la tabla de Piezas si quien llama ya la ha leído (y la mantiene al día con lo que escriba después).
 */
function leerParaAbonos_(tabP, crono) {
  if (recolocarTablaAbonos_()) { montarAbonos_(); crono.paso('recolocar tabla de Abonos'); }
  const leida = tabla => `${tabla.filas.length} de ${tabla.leidas} filas`;
  if (!tabP) { tabP = leerTabla_(HOJA.PIEZAS); crono.paso(`leer Piezas (${leida(tabP)})`); }
  const tabL = leerTabla_(HOJA.LINEAS); crono.paso(`leer Líneas RM (${leida(tabL)})`);
  const tabA = leerTabla_(HOJA.ALB); crono.paso(`leer Albaranes (${leida(tabA)})`);
  const fechaFactura = {};
  leerTabla_(HOJA.FACT).filas.forEach(f => { fechaFactura[String(f.v['Nº factura']).trim()] = aISO_(f.v['Fecha factura']); });
  const sh = hoja_(HOJA.ABONOS), tb = ABONOS.filaTabla, ncol = ABONOS.cabTabla.length, ultFila = sh.getLastRow(), maxFilas = sh.getMaxRows();
  // getLastRow cuenta también el resumen de la derecha: se quitan las filas vacías del final de la tabla.
  const leido = sh.getRange(ABONOS.filaCabTabla, 1, Math.max(ultFila, tb) - ABONOS.filaCabTabla + 1, ncol).getValues();
  colsAbonos_(leido[0]);
  const tabla = leido.slice(tb - ABONOS.filaCabTabla);
  while (tabla.length && tabla[tabla.length - 1].every(x => x === '' || x == null)) tabla.pop();
  const shC = ss_().getSheetByName(HOJA.CLAVES), ultC = shC ? shC.getLastRow() : 0;
  const vistas = new Set(ultC ? shC.getRange(1, 1, ultC, 1).getValues().map(r => String(r[0])).filter(Boolean) : []);
  loc_('');  // la configuración regional también es una lectura: se guarda ya para escribir las fórmulas luego
  crono.paso(`leer Abonos (${tabla.length} filas, ${vistas.size} claves vistas)`);
  return { tabP, tabL, tabA, sh, tabla, vistas, ultFila: tb + tabla.length - 1, maxFilas, fechaFactura };
}

/**
 * Piezas de RM (las de proveedor "Otros" no se reclaman) con su clave P. La numeración de la clave cuenta TODAS las
 * piezas, marcadas o no, para que no cambie al marcar o desmarcar otra pieza del mismo albarán y referencia.
 */
function piezasParaAbonos_(tabP, tabA, tabL) {
  const proveedorDe = {};
  tabA.filas.forEach(f => { const n = normAlbaran(f.v['Nº albarán']); if (n) proveedorDe[n] = f.v['Proveedor']; });
  const plateDe = matriculasDeAlbaranes_(tabA, tabL);
  const todas = [];
  tabP.filas.forEach(p => {
    const alb = normAlbaran(p.v['Nº albarán']);
    if (String(proveedorDe[alb] || 'RM') === 'Otros') return;
    todas.push({ fila: p.fila, marcada: p.v['Reembolso'] === true && !!alb, albaran: alb, ref: p.v['Referencia pieza'], desc: p.v['Descripción'],
      sinIva: p.v['Precio descontado sin IVA'], fechaReembolso: aISO_(p.v['Fecha reembolso']), matricula: p.v['Matrícula'] || plateDe[alb] || '' });
  });
  return { todas: clavesPiezas(todas), plateDe };
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

/**
 * Aplica a la tabla de Abonos lo leído en leerParaAbonos_. No lee nada de la hoja.
 * filasDesmarcadas: filas de Piezas a las que se acaba de quitar el Reembolso (su fila de Abonos se borra).
 * escaneados: líneas de albaranes de abono recién escaneados (opcional).
 */
function escribirAbonos_(d, filasDesmarcadas, escaneados) {
  const crono = cronometro_('escribirAbonos_');
  const { tabP, tabL, tabA, sh } = d;
  const hoy = hoyISO_(), tb = ABONOS.filaTabla, ncol = ABONOS.cabTabla.length;
  const { todas, plateDe } = piezasParaAbonos_(tabP, tabA, tabL);
  const marcadas = todas.filter(p => p.marcada);
  let fechasEscritas = 0;
  marcadas.forEach(p => {
    if (!p.fechaReembolso) { p.fechaReembolso = hoy; actualizarFila_(tabP, p.fila, { 'Fecha reembolso': aFecha_(hoy) }); fechasEscritas++; }
  });
  const quitar = todas.filter(p => (filasDesmarcadas || []).indexOf(p.fila) >= 0).map(p => p.clave);
  crono.paso(`preparar piezas (${marcadas.length} marcadas, ${fechasEscritas} fechas escritas)`);

  const deFacturas = abonosDeLineas_(tabL, plateDe, d.fechaFactura);
  const abonos = deFacturas.concat(clavesAbonos((escaneados || []).map(a => Object.assign({}, a, { matricula: plateDe[normAlbaran(a.albaranOrigen)] || '' }))));

  const c = colsAbonos_(), g = (r, k) => r[c[k] - 1];
  const existentes = d.tabla.map(r => ({ clave: g(r, 'clave'), fechaAbono: g(r, 'fechaAbono'), descripcion: g(r, 'descripcion'), sinIva: g(r, 'sinIva'),
    estado: g(r, 'estado'), albaran: g(r, 'albaran'), ref: g(r, 'referencia'), matricula: g(r, 'matricula'), fechaSolicitud: g(r, 'fechaSolicitud'),
    factura: g(r, 'factura'), nota: g(r, 'nota') }));
  const res = sincronizarAbonos(existentes, marcadas, todas, abonos, d.vistas, quitar, { matriculas: plateDe });
  crono.paso(`cruce (${abonos.length} abonos)`);

  // 1) Celdas vacías de filas existentes (antes de borrar o insertar: los índices aún valen).
  res.cambios.forEach(({ i, v }) => Object.keys(v).forEach(k => {
    sh.getRange(tb + i, c[k]).setValue(k === 'fechaAbono' || k === 'fechaSolicitud' ? aFecha_(v[k]) : v[k]);
  }));
  // 2) Filas de piezas desmarcadas, de abajo arriba. Sólo las celdas de la tabla (A-M): una fila entera se llevaría
  //    también la fila del resumen que hay a la derecha.
  res.borrar.slice().sort((a, b) => b - a).forEach(i => sh.getRange(tb + i, 1, 1, ncol).deleteCells(SpreadsheetApp.Dimension.ROWS));
  // 3) Filas nuevas ARRIBA, justo debajo de la cabecera, desplazando hacia abajo sólo las celdas de la tabla.
  const k = res.nuevas.length;
  if (k) {
    if (d.ultFila + k > d.maxFilas) sh.insertRowsAfter(d.maxFilas, d.ultFila + k - d.maxFilas + 100);  // al final: no mueve nada
    sh.getRange(tb, 1, k, ncol).insertCells(SpreadsheetApp.Dimension.ROWS);
    const modelo = sh.getRange(tb + k, 1, 1, ncol), destino = sh.getRange(tb, 1, k, ncol);
    modelo.copyTo(destino, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    modelo.copyTo(destino, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
    destino.setValues(res.nuevas.map((f, j) => filaAbono_(f, tb + j)));
  }
  if (res.registrar.length) registrarClavesAbonos_(res.registrar);
  crono.paso(`escribir (${k} nuevas, ${res.cambios.length} completadas, ${res.borrar.length} borradas)`);
  // Sólo a Ejecuciones: escribirlo en Registro en cada edición costaba ~300 ms y llenaba Registro de filas INFO.
  crono.fin(ESTADOS_ABONO.map(e => e + ' ' + res.nuevas.filter(f => f.estado === e).length).join(', '));
  return res;
}

/**
 * Líneas de abono de las facturas RM, con su clave A. Fecha de abono = fecha de la factura (si la pieza la marca la
 * factura; si antes llegó el albarán de abono escaneado, esa fila ya tiene la fecha del albarán y no se cambia).
 */
function abonosDeLineas_(tabL, plateDe, fechaFactura) {
  return clavesAbonos(tabL.filas.filter(f => f.v['Tipo'] === 'Abono').map(f => ({
    factura: f.v['Nº factura'], fecha: (fechaFactura || {})[String(f.v['Nº factura']).trim()] || aISO_(f.v['Fecha albarán']),
    albaranAbono: normAlbaran(f.v['Nº albarán']), albaranOrigen: albaranOrigen(f.v['Albarán origen']),
    ref: f.v['Referencia'], desc: f.v['Descripción'], importe: parseNumber(f.v['Importe sin IVA']),
    // El bloque de abono de la factura trae la matrícula vacía: la buena es la de su albarán original.
    matricula: plateDe[albaranOrigen(f.v['Albarán origen'])] || normPlate(f.v['Matrícula']),
  })));
}

function filaAbono_(f, r) {
  const c = colsAbonos_(), L = letraAbonos_, fila = new Array(ABONOS.cabTabla.length).fill('');
  // "Precio con IVA" y "Días pendiente" son fórmulas: la primera usa el IVA de Config y la segunda se actualiza sola cada día.
  const v = { fechaAbono: aFecha_(f.fechaAbono), descripcion: f.descripcion, sinIva: f.sinIva, conIva: loc_(`=ROUND($${L('sinIva')}${r}*(1+IVA),2)`),
    estado: f.estado, albaran: f.albaran, referencia: f.referencia, matricula: f.matricula, fechaSolicitud: aFecha_(f.fechaSolicitud),
    factura: f.factura, nota: f.nota, dias: loc_(`=IF($${L('estado')}${r}="Sin abonar",TODAY()-$${L('fechaSolicitud')}${r},"")`), clave: f.clave };
  Object.keys(v).forEach(k => { fila[c[k] - 1] = v[k]; });
  return fila;
}

/** Claves de abonos de factura ya añadidos alguna vez: si se borra su fila a mano, no vuelve a aparecer. */
function registrarClavesAbonos_(claves) {
  const ss = ss_();
  let sh = ss.getSheetByName(HOJA.CLAVES);
  if (!sh) { sh = ss.insertSheet(HOJA.CLAVES); sh.hideSheet(); }
  sh.getRange(sh.getLastRow() + 1, 1, claves.length, 1).setValues(claves.map(c => [c]));
}

/**
 * Líneas de un albarán de ABONO escaneado, en el formato de abonosDeLineas_ (sin factura todavía): fecha de abono =
 * fecha del albarán de abono. Sin la línea de residuos (SIGAUS): esa ya llegará con la factura.
 */
function abonosEscaneados_(doc, hoy) {
  return doc.lineas.filter(l => !esResiduo(l)).map(l => ({
    factura: '', fecha: doc.fecha || hoy, albaranAbono: normAlbaran(doc.numero_albaran), albaranOrigen: albaranOrigen(l.albaran_origen),
    ref: l.referencia, desc: l.descripcion, importe: parseNumber(l.importe),
  }));
}

/**
 * Filas borradas a mano en Abonos: las piezas marcadas cuya fila ya no está se desmarcan en Piezas.
 * Una fila de abono sin pieza (Sin solicitar) no afecta a nada más.
 */
function desmarcarPiezasBorradas_() {
  if (recolocarTablaAbonos_()) montarAbonos_();
  const sh = hoja_(HOJA.ABONOS), tb = ABONOS.filaTabla, ult = sh.getLastRow();
  const enTabla = new Set();
  if (ult >= tb) sh.getRange(tb, colsAbonos_().clave, ult - tb + 1, 1).getValues().forEach(r => partirClaves(r[0]).forEach(k => enTabla.add(k)));
  const tabP = leerTabla_(HOJA.PIEZAS);
  const { todas } = piezasParaAbonos_(tabP, leerTabla_(HOJA.ALB), null);
  const quitar = todas.filter(p => p.marcada && !enTabla.has(p.clave));
  quitar.forEach(p => actualizarFila_(tabP, p.fila, { 'Reembolso': false, 'Fecha reembolso': '' }));
  if (quitar.length) {
    const lista = quitar.map(p => `${p.albaran} ${p.ref || p.desc}`).join(', ');
    log_('INFO', 'desmarcarPiezasBorradas', HOJA.PIEZAS, `Fila borrada en Abonos → Reembolso desmarcado: ${lista}`);
    toast_(`Reembolso desmarcado en Piezas: ${lista}`, 'Abonos', 8);
  }
}

/**
 * Edición a mano en la tabla de Abonos. Vaciar una fila entera (seleccionarla y pulsar Suprimir) cuenta como borrarla:
 * se quita el hueco y, si era de una pieza, se desmarca en Piezas. Cualquier otro cambio se respeta tal cual.
 */
function editarAbonos_(r0, n, c0) {
  if (c0 > ABONOS.cabTabla.length) return;  // resumen / panel de la derecha: no es la tabla
  // Si la tabla estaba descolocada, la fila editada ya no es la que era: sólo se recoloca y se revisan las piezas.
  if (recolocarTablaAbonos_()) { montarAbonos_(); desmarcarPiezasBorradas_(); return; }
  const sh = hoja_(HOJA.ABONOS), tb = ABONOS.filaTabla, ncol = ABONOS.cabTabla.length;
  const desde = Math.max(r0, tb), hasta = r0 + n, ult = sh.getLastRow();
  if (hasta <= desde) return;
  const vals = sh.getRange(desde, 1, hasta - desde, ncol).getValues();
  const cols = ENTRADAS_ABONOS_.map(k => colsAbonos_()[k]), vacias = [];
  vals.forEach((v, i) => { if (cols.every(c => v[c - 1] === '' || v[c - 1] == null)) vacias.push(desde + i); });
  if (!vacias.length) return;
  // Lo que queda por debajo de la última fila con algo ya está en blanco. De abajo arriba para no mover las siguientes.
  vacias.filter(r => r <= ult).sort((a, b) => b - a).forEach(r => sh.getRange(r, 1, 1, ncol).deleteCells(SpreadsheetApp.Dimension.ROWS));
  desmarcarPiezasBorradas_();
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
  const idx = vals.findIndex(r => r.indexOf(CAMPOS_ABONOS_.fechaAbono) >= 0 && r.indexOf(CAMPOS_ABONOS_.clave) >= 0);
  if (idx < 0) {
    if (vals.every(vacia)) return false;  // hoja nueva: aún no hay tabla
    throw new Error('No encuentro la cabecera de la tabla de Abonos ("Fecha abono" … "Clave" en la fila 1). ' +
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

/**
 * Filas enteras insertadas o borradas a mano en Abonos: la tabla se revisa (y se desmarcan las piezas cuya fila se ha
 * borrado) y, si el resumen de la derecha se ha descolocado, se vuelve a montar.
 */
function revisarAbonosTrasCambioDeFilas_(borradas) {
  if (recolocarTablaAbonos_()) montarAbonos_();
  else if (!resumenAbonosEnSuSitio_(hoja_(HOJA.ABONOS))) montarResumenAbonos_();
  if (borradas) desmarcarPiezasBorradas_();
}

/** Menú: Actualizar Abonos (añade lo que falte; no cambia ni borra nada de lo que ya hay). */
function actualizarAbonos() {
  ejecutar_('actualizarAbonos', () => conBloqueo_(10, () => {
    actualizarHoja_(false);
    const r = sincronizarAbonos_();
    toast_(r.nuevas.length ? `Abonos: ${r.nuevas.length} fila(s) nueva(s) arriba.` : 'Abonos ya estaba al día.');
  }));
}
