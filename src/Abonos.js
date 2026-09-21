/**
 * Pestaña Abonos.
 *  - Resumen por quincena (filas 4-27, fórmulas creadas en setup).
 *  - Tabla grande (desde la fila 31): se RECONSTRUYE entera desde Piezas (Reembolso marcado) + Líneas RM (abonos de las facturas).
 *    Por eso no se edita a mano: se corrige en Piezas (check de Reembolso) o reescaneando la factura.
 */

function reconstruirAbonos_() {
  const tabP = leerTabla_(HOJA.PIEZAS), tabL = leerTabla_(HOJA.LINEAS), tabA = leerTabla_(HOJA.ALB);
  const hoy = hoyISO_(), iva = cfgNum_('IVA');
  const plateDe = {};
  tabA.filas.forEach(f => { const n = normAlbaran(f.v['Nº albarán']); if (n) plateDe[n] = f.v['Matrícula']; });

  const marcadas = [], todas = [];
  tabP.filas.forEach(p => {
    if (String(p.v['Proveedor'] || 'RM') === 'Otros') return;
    const alb = normAlbaran(p.v['Nº albarán']);
    const item = { albaran: alb, ref: p.v['Referencia pieza'], desc: p.v['Descripción'], sinIva: p.v['Precio descontado sin IVA'],
      fechaReembolso: aISO_(p.v['Fecha reembolso']), matricula: p.v['Matrícula'] || plateDe[alb] || '' };
    todas.push(item);
    if (p.v['Reembolso'] === true && alb) {
      if (!item.fechaReembolso) { item.fechaReembolso = hoy; actualizarFila_(tabP, p.fila, { 'Fecha reembolso': aFecha_(hoy) }); }
      marcadas.push(item);
    }
  });

  const abonos = tabL.filas.filter(f => f.v['Tipo'] === 'Abono').map(f => ({
    factura: f.v['Nº factura'], fecha: aISO_(f.v['Fecha albarán']), albaranOrigen: albaranOrigen(f.v['Albarán origen']),
    ref: f.v['Referencia'], desc: f.v['Descripción'], importe: parseNumber(f.v['Importe sin IVA']), matricula: plateDe[albaranOrigen(f.v['Albarán origen'])] || '',
  }));

  const filas = construirAbonos(marcadas, todas, abonos, iva);
  const sh = hoja_(HOJA.ABONOS), ini = ABONOS.filaTabla, ncol = ABONOS.cabTabla.length;
  const ult = Math.max(sh.getLastRow(), ini);
  sh.getRange(ini, 1, ult - ini + 1, ncol).clearContent();
  if (filas.length) {
    const necesarias = ini + filas.length - 1;
    if (necesarias > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), necesarias - sh.getMaxRows() + 50);
    sh.getRange(ini, 1, filas.length, ncol).setValues(filas.map(f => [aFecha_(f.fechaAbono), f.descripcion, f.sinIva, f.conIva, f.estado,
      f.albaran, f.referencia, f.matricula, aFecha_(f.fechaSolicitud), f.factura, f.nota]));
  }
  log_('INFO', 'reconstruirAbonos', '', `${filas.length} filas: ${ESTADOS_ABONO.map(e => e + ' ' + filas.filter(f => f.estado === e).length).join(', ')}`);
  return filas;
}

/** Menú: Actualizar Abonos. */
function actualizarAbonos() {
  ejecutar_('actualizarAbonos', () => conBloqueo_(10, () => {
    const f = reconstruirAbonos_();
    toast_(`Abonos actualizado: ${f.length} filas.`);
  }));
}
