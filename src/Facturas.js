/** Facturas quincenales de RM: leer con Gemini, validar el IVA, guardar cabecera + líneas y reconstruir Abonos. */

function procesarFacturasRM() {
  ejecutar_('procesarFacturasRM', () => conBloqueo_(10, () => {
    const archivos = listarArchivos_('CARPETA_FACTURAS_RM', 5);
    if (!archivos.length) { avisar_('No hay facturas nuevas en la carpeta de facturas RM.'); return; }
    toast_(`Leyendo ${archivos.length} factura(s) con Gemini…`, 'Facturas RM', 60);
    const resultados = leerConGemini_(archivos, false);
    const cfgErr = resultados.find(r => !r.ok && r.config);
    if (cfgErr) throw new Error(cfgErr.error);
    const lineas = [];
    archivos.forEach((f, i) => {
      const r = resultados[i], nombre = f.getName();
      try {
        if (!r.ok) { log_('ERROR', 'procesarFacturasRM', nombre, `${r.error} (el PDF sigue en su carpeta; vuelve a intentarlo)`); lineas.push(`${nombre}: fallo temporal, reintenta`); return; }
        const res = registrarFactura_(f, r.doc, r.truncado);
        if (res.fatal) { moverAErrores_(f); lineas.push(`${nombre}: NO legible (${res.estado}). Movida a Errores`); return; }
        moverArchivo_(f, 'CARPETA_FACTURAS_RM_PROCESADAS');
        lineas.push(`${res.numero} (${MESES[res.periodo.month - 1]} quincena ${res.periodo.quincena}): ${res.estado}`);
      } catch (e) { log_('ERROR', 'procesarFacturasRM', nombre, (e && e.stack) || e); lineas.push(`${nombre}: error, mira Registro`); }
    });
    reconstruirAbonos_();
    avisar_(lineas.join('\n'), 'Facturas RM');
  }));
}

/** Guarda (o reemplaza) una factura. Devuelve {fatal, estado, numero, periodo}. */
function registrarFactura_(archivo, doc, truncado) {
  const ref = archivo.getName(), iva = cfgNum_('IVA');
  const v = validarFactura(doc, iva);
  if (truncado) v.warnings.push('La respuesta de Gemini se cortó: puede faltar información');
  if (v.errors.length) { log_('ERROR', 'registrarFactura', ref, v.errors.join('; ')); return { fatal: true, estado: v.errors.join('; ') }; }
  v.warnings.forEach(w => log_('AVISO', 'registrarFactura', `${doc.numero_factura}`, w));
  const per = periodoFactura(doc);
  const estado = v.warnings.length ? '⚠ ' + v.warnings.join('; ') : 'OK';

  // Reemplazo idempotente: quitar las líneas anteriores de esta misma factura.
  const tabL = leerTabla_(HOJA.LINEAS);
  const previas = tabL.filas.filter(f => String(f.v['Nº factura']).trim() === doc.numero_factura);
  if (previas.length) {
    const resto = tabL.filas.filter(f => String(f.v['Nº factura']).trim() !== doc.numero_factura).map(f => f.v);
    vaciarTabla_(tabL);
    agregarFilas_(tabL, resto);
  }
  const filas = [];
  doc.albaranes.forEach(a => {
    const tipo = a.es_abono ? 'Abono' : 'Compra';
    const ls = a.lineas.length ? a.lineas : [{ referencia: '', descripcion: '(sin detalle de líneas)', importe: a.importe, albaran_origen: '' }];
    ls.forEach(l => filas.push({
      'Nº factura': doc.numero_factura, 'Nº albarán': a.numero_albaran, 'Fecha albarán': aFecha_(a.fecha), 'Matrícula': a.matricula, 'Tipo': tipo,
      'Albarán origen': tipo === 'Abono' ? l.albaran_origen : '', 'Referencia': l.referencia, 'Descripción': l.descripcion,
      'Cantidad': l.cantidad == null ? '' : l.cantidad, 'Precio': l.precio_unitario == null ? '' : l.precio_unitario,
      'Descuento': l.descuento_pct == null ? '' : l.descuento_pct / 100, 'Importe sin IVA': l.importe == null ? 0 : l.importe,
    }));
  });
  agregarFilas_(tabL, filas);

  const tabF = leerTabla_(HOJA.FACT);
  const fila = { 'Nº factura': doc.numero_factura, 'Fecha factura': aFecha_(doc.fecha_factura), 'Año': per.year, 'Mes': per.month, 'Quincena': per.quincena,
    'Base imponible': doc.base_imponible, 'IVA': doc.iva_importe, 'Total': doc.total, 'Estado': estado, 'Ver PDF': enlacePdf_(archivo),
    'Procesada el': new Date(), 'Nº albaranes': doc.albaranes.length };
  const prev = tabF.filas.find(f => String(f.v['Nº factura']).trim() === doc.numero_factura);
  if (prev) actualizarFila_(tabF, prev.fila, fila); else agregarFilas_(tabF, [fila]);
  log_('INFO', 'registrarFactura', doc.numero_factura, `${doc.albaranes.length} albaranes, total ${doc.total}: ${estado}`);
  return { fatal: false, estado, numero: doc.numero_factura, periodo: per };
}

/** Borra el contenido de las columnas de entrada (deja fórmulas y formato). */
function vaciarTabla_(tabla) {
  const ult = tabla.sh.getLastRow();
  if (ult > 1) tabla.esq.entradas.forEach(h => tabla.sh.getRange(2, tabla.map[h], ult - 1, 1).clearContent());
  tabla.filas = [];
  tabla.libre = 2;
}

/** Botón "Reescanear factura" de la pestaña Abonos: vuelve a leer con Gemini la(s) factura(s) de esa quincena. */
function reescanearPeriodo_(anio, mes, q) {
  ejecutar_('reescanearFactura', () => conBloqueo_(20, () => {
    const tabF = leerTabla_(HOJA.FACT);
    const filas = tabF.filas.filter(f => Number(f.v['Año']) === anio && Number(f.v['Mes']) === mes && Number(f.v['Quincena']) === q);
    if (!filas.length) { avisar_(`No hay ninguna factura RM registrada para ${MESES[mes - 1]} ${anio}, quincena ${q}.`); return; }
    const out = [];
    filas.forEach(f => {
      const id = idDeEnlace_(tabF.sh.getRange(f.fila, tabF.map['Ver PDF']).getFormula());
      if (!id) { out.push(`${f.v['Nº factura']}: no encuentro el PDF (columna Ver PDF)`); return; }
      const archivo = DriveApp.getFileById(id);
      toast_(`Volviendo a leer ${f.v['Nº factura']} con Gemini…`, 'Reescanear factura', 60);
      const r = leerConGemini_([archivo], false)[0];
      if (!r.ok) { if (r.config) throw new Error(r.error); out.push(`${f.v['Nº factura']}: ${r.error}`); return; }
      const res = registrarFactura_(archivo, r.doc, r.truncado);
      out.push(res.fatal ? `${f.v['Nº factura']}: ilegible (${res.estado})` : `${res.numero}: ${res.estado}`);
    });
    reconstruirAbonos_();
    avisar_(out.join('\n'), 'Reescanear factura');
  }));
}
