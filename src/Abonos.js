/**
 * Pestaña Abonos.
 *  - Resumen por quincena (filas 4-27, fórmulas creadas en setup).
 *  - Tabla grande (desde la fila 31): se RECONSTRUYE entera desde Piezas (Reembolso marcado) + Líneas RM (abonos de las facturas).
 *    Por eso no se edita a mano: se corrige en Piezas (check de Reembolso) o reescaneando la factura.
 */

/** tabP: la tabla de Piezas si quien llama ya la ha leído y la ha mantenido al día (evita leerla dos veces). */
function reconstruirAbonos_(tabP) {
  const crono = cronometro_('reconstruirAbonos_');
  const leida = tabla => `${tabla.filas.length} de ${tabla.leidas} filas`;
  if (tabP) crono.paso('Piezas ya leída');
  else { tabP = leerTabla_(HOJA.PIEZAS); crono.paso(`leer Piezas (${leida(tabP)})`); }
  const tabL = leerTabla_(HOJA.LINEAS); crono.paso(`leer Líneas RM (${leida(tabL)})`);
  const tabA = leerTabla_(HOJA.ALB); crono.paso(`leer Albaranes (${leida(tabA)})`);
  const hoy = hoyISO_();
  const plateDe = {}, proveedorDe = {};
  tabA.filas.forEach(f => {
    const n = normAlbaran(f.v['Nº albarán']);
    if (n) { plateDe[n] = f.v['Matrícula']; proveedorDe[n] = f.v['Proveedor']; }
  });

  const marcadas = [], todas = [];
  let fechasEscritas = 0;
  tabP.filas.forEach(p => {
    const alb = normAlbaran(p.v['Nº albarán']);
    if (String(proveedorDe[alb] || 'RM') === 'Otros') return;
    const item = { albaran: alb, ref: p.v['Referencia pieza'], desc: p.v['Descripción'], sinIva: p.v['Precio descontado sin IVA'],
      fechaReembolso: aISO_(p.v['Fecha reembolso']), matricula: p.v['Matrícula'] || plateDe[alb] || '' };
    todas.push(item);
    if (p.v['Reembolso'] === true && alb) {
      if (!item.fechaReembolso) { item.fechaReembolso = hoy; actualizarFila_(tabP, p.fila, { 'Fecha reembolso': aFecha_(hoy) }); fechasEscritas++; }
      marcadas.push(item);
    }
  });
  crono.paso(`preparar piezas (${marcadas.length} marcadas, ${fechasEscritas} fechas escritas)`);

  const abonos = tabL.filas.filter(f => f.v['Tipo'] === 'Abono').map(f => ({
    factura: f.v['Nº factura'], fecha: aISO_(f.v['Fecha albarán']), albaranOrigen: albaranOrigen(f.v['Albarán origen']),
    ref: f.v['Referencia'], desc: f.v['Descripción'], importe: parseNumber(f.v['Importe sin IVA']), matricula: plateDe[albaranOrigen(f.v['Albarán origen'])] || '',
  }));

  const filas = construirAbonos(marcadas, todas, abonos);
  crono.paso(`cruce (${abonos.length} abonos)`);
  const sh = hoja_(HOJA.ABONOS), ini = ABONOS.filaTabla, ncol = ABONOS.cabTabla.length;
  const ult = Math.max(sh.getLastRow(), ini);
  sh.getRange(ini, 1, ult - ini + 1, ncol).clearContent();
  crono.paso(`borrar (${ult - ini + 1} filas)`);
  if (filas.length) {
    const necesarias = ini + filas.length - 1;
    if (necesarias > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), necesarias - sh.getMaxRows() + 50);
    // 'Días pendiente' es una fórmula (no un valor calculado aquí) para que se actualice sola día a día sin rehacer Abonos.
    sh.getRange(ini, 1, filas.length, ncol).setValues(filas.map((f, i) => {
      const r = ini + i;
      // "Precio con IVA" como fórmula con el IVA de Config: así no hay que leer Config en cada reconstrucción.
      return [aFecha_(f.fechaAbono), f.descripcion, f.sinIva, loc_(`=ROUND($C${r}*(1+IVA),2)`), f.estado, f.albaran, f.referencia, f.matricula,
        aFecha_(f.fechaSolicitud), f.factura, f.nota, loc_(`=IF($E${r}="Sin abonar",TODAY()-$I${r},"")`)];
    }));
  }
  crono.paso(`escribir (${filas.length} filas)`);
  // Sólo a Ejecuciones: escribirlo en Registro en cada edición costaba ~300 ms y llenaba Registro de filas INFO.
  crono.fin(ESTADOS_ABONO.map(e => e + ' ' + filas.filter(f => f.estado === e).length).join(', '));
  return filas;
}

/** Menú: Actualizar Abonos. */
function actualizarAbonos() {
  ejecutar_('actualizarAbonos', () => conBloqueo_(10, () => {
    const f = reconstruirAbonos_();
    toast_(`Abonos actualizado: ${f.length} filas.`);
  }));
}
