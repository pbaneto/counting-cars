/**
 * Pestaña Abonos.
 *  - Resumen por quincena (filas 4-27, fórmulas creadas en setup).
 *  - Tabla grande (desde la fila 31): se RECONSTRUYE entera desde Piezas (Reembolso marcado) + Líneas RM (abonos de las facturas).
 *    Por eso no se edita a mano: se corrige en Piezas (check de Reembolso) o reescaneando la factura.
 */

function reconstruirAbonos_() {
  const crono = cronometro_('reconstruirAbonos_');
  const d = leerParaAbonos_(null, crono);
  crono.fin();
  return escribirAbonos_(d);
}

/**
 * Todo lo que la reconstrucción necesita LEER, para hacerlo antes de escribir nada: en Sheets, una lectura hecha
 * después de una escritura espera a que se recalculen las fórmulas que dependen de lo escrito.
 * tabP: la tabla de Piezas si quien llama ya la ha leído (y la mantiene al día con lo que escriba después).
 */
function leerParaAbonos_(tabP, crono) {
  const leida = tabla => `${tabla.filas.length} de ${tabla.leidas} filas`;
  if (!tabP) { tabP = leerTabla_(HOJA.PIEZAS); crono.paso(`leer Piezas (${leida(tabP)})`); }
  const tabL = leerTabla_(HOJA.LINEAS); crono.paso(`leer Líneas RM (${leida(tabL)})`);
  const tabA = leerTabla_(HOJA.ALB); crono.paso(`leer Albaranes (${leida(tabA)})`);
  const sh = hoja_(HOJA.ABONOS), ultFila = sh.getLastRow(), maxFilas = sh.getMaxRows();
  loc_('');  // la configuración regional también es una lectura: se guarda ya para escribir las fórmulas luego
  crono.paso('leer tamaño de Abonos');
  return { tabP, tabL, tabA, sh, ultFila, maxFilas };
}

/** Rehace la tabla grande de Abonos con lo leído en leerParaAbonos_. No lee nada de la hoja. */
function escribirAbonos_(d) {
  const crono = cronometro_('escribirAbonos_');
  const { tabP, tabL, tabA, sh } = d;
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
  const ini = ABONOS.filaTabla, ncol = ABONOS.cabTabla.length;
  const ult = Math.max(d.ultFila, ini);
  sh.getRange(ini, 1, ult - ini + 1, ncol).clearContent();
  crono.paso(`borrar (${ult - ini + 1} filas)`);
  if (filas.length) {
    const necesarias = ini + filas.length - 1;
    if (necesarias > d.maxFilas) sh.insertRowsAfter(d.maxFilas, necesarias - d.maxFilas + 50);
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
