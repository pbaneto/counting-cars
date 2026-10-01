/** Procesar albaranes de la carpeta Entrada: Gemini en paralelo -> filas en Albaranes/Trabajos/Piezas -> PDF a Procesados. */

function procesarAlbaranes() {
  ejecutar_('procesarAlbaranes', () => conBloqueo_(10, () => {
    const archivos = listarArchivos_('CARPETA_ENTRADA', cfgNum_('MAX_ARCHIVOS'));
    if (!archivos.length) { avisar_('No hay albaranes nuevos en la carpeta Entrada.'); return; }
    // Antes de gastar llamadas a Gemini: si la hoja no está al día (columnas nuevas), leerTabla_ lo dice ya.
    const ctx = contextoAlbaranes_();
    toast_(`Leyendo ${archivos.length} albarán(es) con Gemini…`, 'Procesar albaranes', 30);
    const resultados = leerConGemini_(archivos, true);
    const cfgErr = resultados.find(r => !r.ok && r.config);
    if (cfgErr) throw new Error(cfgErr.error);
    releerDescuadrados_(archivos, resultados, ctx.iva);

    const cuenta = { creado: 0, reembolso: 0, duplicado: 0, abono: 0, descuadre: 0, error: 0, pendiente: 0 };
    const descuadres = [], mover = [];
    let fallo = null;
    archivos.forEach((f, i) => {
      const r = resultados[i], nombre = f.getName();
      if (fallo) { cuenta.pendiente++; return; }
      if (!r.ok) { cuenta.pendiente++; log_('ERROR', 'procesarAlbaranes', nombre, `${r.error} (el PDF sigue en Entrada; vuelve a intentarlo)`); return; }
      try {
        const res = procesarDocAlbaran_(ctx, f, r.doc);
        cuenta[res.estado]++;
        if (res.estado === 'descuadre') descuadres.push(`${nombre}: ${res.motivo}`);  // se queda en Entrada
        else mover.push({ f, errores: res.estado === 'error' });
      } catch (e) {
        // Un fallo al escribir puede venir de un albarán anterior (Sheets avisa en la siguiente lectura): no se mueve ninguno.
        fallo = { nombre, e };
        cuenta.pendiente++;
      }
    });
    // Los PDF se mueven sólo después de guardar la hoja: si Sheets rechaza algo, lo dice aquí y no se mueve ninguno.
    // Volver a procesar es seguro: un albarán ya guardado lleva el enlace a su PDF y se reconoce como ya registrado.
    if (!fallo) { try { SpreadsheetApp.flush(); } catch (e) { fallo = { nombre: '', e }; } }
    if (fallo) {
      log_('ERROR', 'procesarAlbaranes', fallo.nombre, `${(fallo.e && fallo.e.stack) || fallo.e} | No se ha movido ningún PDF: siguen todos en Entrada.`);
    } else {
      mover.forEach(({ f, errores }) => {
        try { if (errores) moverAErrores_(f); else moverArchivo_(f, 'CARPETA_PROCESADOS'); }
        catch (e) { log_('AVISO', 'procesarAlbaranes', f.getName(), `No se ha podido mover el PDF: ${e.message}. La próxima vez se reconocerá como ya registrado.`); }
      });
      if (ctx.tocaAbonos) sincronizarAbonos_();
    }
    const linea = (n, txt) => (n ? `${txt}: ${n}\n` : '');
    let msg = linea(cuenta.creado, 'Albaranes nuevos') + linea(cuenta.reembolso, 'Reembolsos marcados') + linea(cuenta.duplicado, 'Duplicados ignorados') +
      linea(cuenta.abono, 'Albaranes de abono (llegarán en la factura RM)') + linea(cuenta.error, 'No legibles (carpeta Errores)') +
      linea(cuenta.pendiente, 'Sin procesar, siguen en Entrada');
    if (descuadres.length) msg += `\nNo cuadran los importes, siguen en Entrada (${descuadres.length}):\n` + descuadres.map(d => '• ' + d).join('\n') + '\n';
    if (fallo) msg = '⚠ No se ha podido guardar en la hoja: no se ha movido ningún PDF, siguen todos en Entrada. Detalle en Registro.';
    else if (cuenta.error + cuenta.pendiente) msg += '\nDetalle en la pestaña Registro.';
    log_('INFO', 'procesarAlbaranes', '', msg.trim().replace(/\n+/g, ' | '));
    avisar_(msg.trim() || 'Nada que hacer.', 'Procesar albaranes');
  }));
}

/**
 * Albaranes cuyos importes no cuadran (cuadreAlbaran): segunda lectura con Gemini diciéndole qué no cuadraba.
 * Si la segunda lectura cuadra, se usa esa; si no, se queda la primera (y el albarán se quedará en Entrada).
 */
function releerDescuadrados_(archivos, resultados, iva) {
  const repetir = [];
  resultados.forEach((r, i) => {
    if (r.ok && r.doc.es_albaran && !r.doc.es_abono && cuadreAlbaran(r.doc, iva).length) repetir.push(i);
  });
  if (!repetir.length) return;
  toast_(`Releyendo ${repetir.length} albarán(es) cuyos importes no cuadran…`, 'Procesar albaranes', 30);
  const pistas = repetir.map(i => 'ATENCIÓN: en una primera lectura de este albarán los importes no cuadraban (' +
    cuadreAlbaran(resultados[i].doc, iva).join('; ') + '). Vuelve a leer con cuidado la base imponible, el total y el importe de CADA línea, ' +
    'sin saltarte ninguna (incluidas SIGAUS, portes o residuos).');
  const segundas = leerConGemini_(repetir.map(i => archivos[i]), true, pistas);
  segundas.forEach((s, k) => {
    if (s.ok && !cuadreAlbaran(s.doc, iva).length) {
      resultados[repetir[k]] = s;
      log_('INFO', 'procesarAlbaranes', archivos[repetir[k]].getName(), 'Importes corregidos en la segunda lectura');
    }
  });
}

function contextoAlbaranes_() {
  const alb = leerTabla_(HOJA.ALB), trab = leerTabla_(HOJA.TRAB), piezas = leerTabla_(HOJA.PIEZAS), coches = leerTabla_(HOJA.COCHES);
  const idsPdf = new Set();
  const ult = alb.sh.getLastRow();
  if (ult > 1) alb.sh.getRange(2, alb.map['Ver PDF'], ult - 1, 1).getFormulas().forEach(r => { const id = idDeEnlace_(r[0]); if (id) idsPdf.add(id); });
  return { alb, trab, piezas, idsPdf, iva: cfgNum_('IVA'), hoy: hoyISO_(), tocaAbonos: false,
    coches: new Set(coches.filas.map(f => normPlate(f.v['Matrícula']))) };
}

/** Devuelve { estado: 'creado' | 'reembolso' | 'duplicado' | 'abono' | 'descuadre' | 'error', motivo }. */
function procesarDocAlbaran_(ctx, archivo, doc) {
  const ref = archivo.getName(), num = doc.numero_albaran, hoy = ctx.hoy;
  if (ctx.idsPdf.has(archivo.getId())) { log_('AVISO', 'procesarAlbaranes', ref, 'Este PDF ya estaba registrado: se archiva sin cambios'); return { estado: 'duplicado' }; }
  if (!doc.es_albaran) { log_('ERROR', 'procesarAlbaranes', ref, 'El documento no parece un albarán: se mueve a Errores'); return { estado: 'error' }; }
  if (doc.es_abono) {
    // Devolución de RM: el abono ya llega en la factura quincenal (pestaña Abonos); no es un gasto del taller.
    log_('INFO', 'procesarAlbaranes', ref, `Albarán de abono ${num || ''} (${doc.total} €): no se añade, llegará en la factura RM`);
    return { estado: 'abono' };
  }
  const v = validarAlbaran(doc, ctx.iva);
  if (v.errors.length) { log_('ERROR', 'procesarAlbaranes', ref, v.errors.join('; ') + ': se mueve a Errores'); return { estado: 'error' }; }
  const lineasPiezas = lineasParaPiezas(doc.lineas);

  const existente = num ? ctx.alb.filas.find(f => normAlbaran(f.v['Nº albarán']) === num) : null;
  if (existente) return procesarSegundoEscaneo_(ctx, archivo, doc, existente, lineasPiezas);

  const problemas = cuadreAlbaran(doc, ctx.iva);
  if (problemas.length) {
    const motivo = problemas.join('; ');
    log_('AVISO', 'procesarAlbaranes', ref, `Los importes no cuadran (${motivo}): no se añade y el PDF se queda en Entrada`);
    return { estado: 'descuadre', motivo };
  }
  v.warnings.forEach(w => log_('AVISO', 'procesarAlbaranes', ref, w));

  const fecha = doc.fecha || hoy;
  const nota = v.warnings.join('; ');
  const manuales = ctx.alb.filas.map(f => ({ row: f.fila, plate: f.v['Matrícula'], total: f.v['Precio con IVA'], albaran: normAlbaran(f.v['Nº albarán']), pdf: f.v['Ver PDF'] }));
  const manual = buscarFilaManual(manuales, doc.matricula, doc.total);
  let filaAlb;
  if (manual) {
    filaAlb = manual.row;
    actualizarFila_(ctx.alb, filaAlb, { 'Fecha escaneo': aFecha_(hoy), 'Nº albarán': num, 'Fecha albarán': aFecha_(fecha), 'Proveedor': doc.proveedor,
      'Ver PDF': enlacePdf_(archivo), 'Nota escaneo': ('Vinculado a la fila que ya estaba escrita a mano. ' + nota).trim() });
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
    agregarFilas_(ctx.piezas, lineasPiezas.map(l => filaPieza_(l, num, hoy, 'Escaneo'))).length;
    if (lineasPiezas.some(l => l.reembolso)) ctx.tocaAbonos = true;
  } else if (lineasPiezas.length) {
    log_('AVISO', 'procesarAlbaranes', `${HOJA.ALB}!${filaAlb}`, 'Albarán sin número: no se guardan sus piezas (no se podrían reembolsar). Escribe el número y vuelve a escanearlo si hace falta');
  }
  return { estado: 'creado' };
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
    return { estado: hayR ? 'reembolso' : 'duplicado' };
  }
  if (!hayR) { log_('AVISO', 'procesarAlbaranes', ref, `Albarán ${num} ya existe y no lleva R: escaneo duplicado, se ignora`); return { estado: 'duplicado' }; }
  const plan = aplicarReembolsos(piezasAlb.map(p => ({ ref: p.v['Referencia pieza'], desc: p.v['Descripción'], reembolso: p.v['Reembolso'] === true })), doc.lineas);
  if (!plan.marcar.length && !plan.anadir.length) { log_('AVISO', 'procesarAlbaranes', ref, `Albarán ${num}: las piezas con R ya estaban marcadas`); return { estado: 'duplicado' }; }
  plan.marcar.forEach(i => actualizarFila_(ctx.piezas, piezasAlb[i].fila, { 'Reembolso': true, 'Fecha reembolso': aFecha_(hoy) }));
  if (plan.anadir.length) agregarFilas_(ctx.piezas, plan.anadir.map(l => filaPieza_(l, num, hoy, 'Escaneo')));
  const resumen = `Reembolso ${hoy.slice(8)}/${hoy.slice(5, 7)}: ${plan.marcar.length + plan.anadir.length} pieza(s)`;
  const previa = String(existente.v['Nota escaneo'] || '');
  actualizarFila_(ctx.alb, existente.fila, { 'Nota escaneo': previa ? previa + ' | ' + resumen : resumen });
  log_('INFO', 'procesarAlbaranes', `${HOJA.ALB}!${existente.fila}`, `${resumen} (albarán ${num})`);
  ctx.tocaAbonos = true;
  return { estado: 'reembolso' };
}
