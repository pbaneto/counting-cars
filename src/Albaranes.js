/** Procesar albaranes de la carpeta Entrada: Gemini en paralelo -> filas en Albaranes/Trabajos/Piezas -> PDF a Procesados. */

function procesarAlbaranes() {
  ejecutar_('procesarAlbaranes', () => conBloqueo_(10, () => {
    const archivos = listarArchivos_('CARPETA_ENTRADA', cfgNum_('MAX_ARCHIVOS'));
    if (!archivos.length) { avisar_('No hay albaranes nuevos en la carpeta Entrada.'); return; }
    toast_(`Leyendo ${archivos.length} albarán(es) con Gemini…`, 'Procesar albaranes', 30);
    const resultados = leerConGemini_(archivos, true);
    const cfgErr = resultados.find(r => !r.ok && r.config);
    if (cfgErr) throw new Error(cfgErr.error);

    const ctx = contextoAlbaranes_();
    const cuenta = { creado: 0, reembolso: 0, duplicado: 0, error: 0, pendiente: 0 };
    archivos.forEach((f, i) => {
      const r = resultados[i], nombre = f.getName();
      try {
        if (!r.ok) { cuenta.pendiente++; log_('ERROR', 'procesarAlbaranes', nombre, `${r.error} (el PDF sigue en Entrada; vuelve a intentarlo)`); return; }
        const estado = procesarDocAlbaran_(ctx, f, r.doc);
        cuenta[estado]++;
        if (estado === 'error') moverAErrores_(f); else moverArchivo_(f, 'CARPETA_PROCESADOS');
      } catch (e) {
        cuenta.pendiente++;
        log_('ERROR', 'procesarAlbaranes', nombre, `${(e && e.stack) || e} (el PDF sigue en Entrada)`);
      }
    });
    if (ctx.tocaAbonos) reconstruirAbonos_();
    const msg = `Albaranes nuevos: ${cuenta.creado}\nReembolsos marcados: ${cuenta.reembolso}\nDuplicados ignorados: ${cuenta.duplicado}\n` +
      `No legibles (carpeta Errores): ${cuenta.error}\nPendientes por fallo temporal: ${cuenta.pendiente}` +
      (cuenta.error + cuenta.pendiente ? '\n\nDetalle en la pestaña Registro.' : '');
    log_('INFO', 'procesarAlbaranes', '', msg.replace(/\n+/g, ' | '));
    avisar_(msg, 'Procesar albaranes');
  }));
}

function contextoAlbaranes_() {
  const alb = leerTabla_(HOJA.ALB), trab = leerTabla_(HOJA.TRAB), piezas = leerTabla_(HOJA.PIEZAS), coches = leerTabla_(HOJA.COCHES);
  const idsPdf = new Set();
  const ult = alb.sh.getLastRow();
  if (ult > 1) alb.sh.getRange(2, alb.map['Ver PDF'], ult - 1, 1).getFormulas().forEach(r => { const id = idDeEnlace_(r[0]); if (id) idsPdf.add(id); });
  return { alb, trab, piezas, idsPdf, iva: cfgNum_('IVA'), hoy: hoyISO_(), tocaAbonos: false,
    coches: new Set(coches.filas.map(f => normPlate(f.v['Matrícula']))) };
}

/** Devuelve 'creado' | 'reembolso' | 'duplicado' | 'error'. */
function procesarDocAlbaran_(ctx, archivo, doc) {
  const ref = archivo.getName(), num = doc.numero_albaran, hoy = ctx.hoy;
  if (ctx.idsPdf.has(archivo.getId())) { log_('AVISO', 'procesarAlbaranes', ref, 'Este PDF ya estaba registrado: se archiva sin cambios'); return 'duplicado'; }
  if (!doc.es_albaran) { log_('ERROR', 'procesarAlbaranes', ref, 'El documento no parece un albarán: se mueve a Errores'); return 'error'; }
  const v = validarAlbaran(doc, ctx.iva);
  if (v.errors.length) { log_('ERROR', 'procesarAlbaranes', ref, v.errors.join('; ') + ': se mueve a Errores'); return 'error'; }
  v.warnings.forEach(w => log_('AVISO', 'procesarAlbaranes', ref, w));
  const lineasPiezas = lineasParaPiezas(doc.lineas);

  const existente = num ? ctx.alb.filas.find(f => normAlbaran(f.v['Nº albarán']) === num) : null;
  if (existente) return procesarSegundoEscaneo_(ctx, archivo, doc, existente, lineasPiezas);

  const fecha = doc.fecha || hoy;
  const nota = v.warnings.join('; ');
  const manuales = ctx.alb.filas.map(f => ({ row: f.fila, plate: f.v['Matrícula'], total: f.v['Precio con IVA'], albaran: normAlbaran(f.v['Nº albarán']), pdf: f.v['Ver PDF'] }));
  const manual = buscarFilaManual(manuales, doc.matricula, doc.total);
  let filaAlb;
  if (manual) {
    filaAlb = manual.row;
    actualizarFila_(ctx.alb, filaAlb, { 'Nº albarán': num, 'Fecha albarán': aFecha_(fecha), 'Proveedor': doc.proveedor, 'Ver PDF': enlacePdf_(archivo),
      'Nota escaneo': ('Vinculado a la fila que ya estaba escrita a mano. ' + nota).trim() });
    log_('INFO', 'procesarAlbaranes', `${HOJA.ALB}!${filaAlb}`, `Albarán ${num} vinculado a la fila manual (misma matrícula e importe)`);
  } else {
    const trabajo = doc.matricula ? asignarTrabajo_(ctx.trab, doc.matricula, fecha) : { num: '' };
    if (doc.matricula && !ctx.coches.has(doc.matricula)) log_('AVISO', 'procesarAlbaranes', ref, `La matrícula ${doc.matricula} no está en Coches`);
    filaAlb = agregarFilas_(ctx.alb, [{ 'Fecha escaneo': aFecha_(hoy), 'Fecha albarán': aFecha_(fecha), 'Proveedor': doc.proveedor, 'Nº albarán': num,
      'Nº trabajo': trabajo.num, 'Matrícula': doc.matricula, 'Precio con IVA': doc.total, 'Ver PDF': enlacePdf_(archivo), 'Nota escaneo': nota }])[0];
    if (trabajo.num) ponerDesplegableTrabajo_(ctx.alb, filaAlb, ctx.trab, doc.matricula);
  }
  ctx.idsPdf.add(archivo.getId());
  if (num) {
    agregarFilas_(ctx.piezas, lineasPiezas.map(l => filaPieza_(l, num, hoy, 'Escaneo')));
    if (lineasPiezas.some(l => l.reembolso)) ctx.tocaAbonos = true;
  } else if (lineasPiezas.length) {
    log_('AVISO', 'procesarAlbaranes', `${HOJA.ALB}!${filaAlb}`, 'Albarán sin número: no se guardan sus piezas (no se podrían reembolsar). Escribe el número y vuelve a escanearlo si hace falta');
  }
  return 'creado';
}

/** El albarán ya existe: con R -> marca esas piezas; sin R -> escaneo duplicado, no se añade nada. */
function procesarSegundoEscaneo_(ctx, archivo, doc, existente, lineasPiezas) {
  const ref = archivo.getName(), num = doc.numero_albaran, hoy = ctx.hoy;
  const piezasAlb = piezasDeAlbaran_(ctx.piezas, num);
  const hayR = lineasPiezas.some(l => l.reembolso);
  if (!piezasAlb.length && lineasPiezas.length) {  // el primer escaneo se quedó sin piezas: se completan
    agregarFilas_(ctx.piezas, lineasPiezas.map(l => filaPieza_(l, num, hoy, 'Escaneo')));
    log_('INFO', 'procesarAlbaranes', ref, `Albarán ${num} ya existía sin piezas: piezas añadidas`);
    if (hayR) ctx.tocaAbonos = true;
    return hayR ? 'reembolso' : 'duplicado';
  }
  if (!hayR) { log_('AVISO', 'procesarAlbaranes', ref, `Albarán ${num} ya existe y no lleva R: escaneo duplicado, se ignora`); return 'duplicado'; }
  const plan = aplicarReembolsos(piezasAlb.map(p => ({ ref: p.v['Referencia pieza'], desc: p.v['Descripción'], reembolso: p.v['Reembolso'] === true })), doc.lineas);
  plan.marcar.forEach(i => actualizarFila_(ctx.piezas, piezasAlb[i].fila, { 'Reembolso': true, 'Fecha reembolso': aFecha_(hoy) }));
  if (plan.anadir.length) agregarFilas_(ctx.piezas, plan.anadir.map(l => filaPieza_(l, num, hoy, 'Escaneo')));
  if (!plan.marcar.length && !plan.anadir.length) { log_('AVISO', 'procesarAlbaranes', ref, `Albarán ${num}: las piezas con R ya estaban marcadas`); return 'duplicado'; }
  const resumen = `Reembolso ${hoy.slice(8)}/${hoy.slice(5, 7)}: ${plan.marcar.length + plan.anadir.length} pieza(s)`;
  const previa = String(existente.v['Nota escaneo'] || '');
  actualizarFila_(ctx.alb, existente.fila, { 'Nota escaneo': previa ? previa + ' | ' + resumen : resumen });
  log_('INFO', 'procesarAlbaranes', `${HOJA.ALB}!${existente.fila}`, `${resumen} (albarán ${num})`);
  ctx.tocaAbonos = true;
  return 'reembolso';
}
