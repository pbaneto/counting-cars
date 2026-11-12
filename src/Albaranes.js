/** Procesar albaranes de la carpeta Entrada: Gemini en paralelo -> filas en Albaranes/Trabajos/Piezas -> PDF a Procesados. */

function procesarAlbaranes() {
  ejecutar_('procesarAlbaranes', () => conBloqueo_(10, () => {
    actualizarHoja_(false);
    const archivos = listarArchivos_('CARPETA_ENTRADA', cfgNum_('MAX_ARCHIVOS'));
    if (!archivos.length) { avisar_('No hay albaranes nuevos en la carpeta Entrada.'); return; }
    // Antes de gastar llamadas a Gemini: si la hoja no está al día (columnas nuevas), leerTabla_ lo dice ya.
    const ctx = contextoAlbaranes_();
    toast_(`Leyendo ${archivos.length} albarán(es) con Gemini…`, 'Procesar albaranes', 30);
    const resultados = leerConGemini_(archivos, true);
    const cfgErr = resultados.find(r => !r.ok && r.config);
    if (cfgErr) throw new Error(cfgErr.error);
    releerDescuadrados_(archivos, resultados, ctx.iva);

    const cuenta = { creado: 0, completado: 0, reembolso: 0, duplicado: 0, abono: 0, descuadre: 0, error: 0, pendiente: 0 };
    const descuadres = [], parecidos = [], dudas = [], mover = [];
    let fallo = null;
    // Las hojas sin total, primero: así la 2ª hoja de un albarán encuentra la 1ª aunque se haya escaneado antes.
    const conTotal = i => (resultados[i].ok && !parseNumber(resultados[i].doc.total) ? 0 : 1);
    archivos.map((_, i) => i).sort((a, b) => conTotal(a) - conTotal(b) || a - b).forEach(i => {
      const f = archivos[i], r = resultados[i], nombre = f.getName();
      if (fallo) { cuenta.pendiente++; return; }
      if (!r.ok) { cuenta.pendiente++; log_('ERROR', 'procesarAlbaranes', nombre, `${r.error} (el PDF sigue en Entrada; vuelve a intentarlo)`); return; }
      try {
        const res = procesarDocAlbaran_(ctx, f, r.doc);
        cuenta[res.estado]++;
        if (res.parecido) parecidos.push(res.parecido);
        if (res.duda) dudas.push(res.duda);
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
    if (!fallo) {
      try {
        if (ctx.tocaAbonos) actualizarInformeAbonos_(ctx.piezas);
        ordenarAlbaranes_();
        SpreadsheetApp.flush();
      } catch (e) { fallo = { nombre: '', e }; }
    }
    if (fallo) {
      log_('ERROR', 'procesarAlbaranes', fallo.nombre, `${(fallo.e && fallo.e.stack) || fallo.e} | No se ha movido ningún PDF: siguen todos en Entrada.`);
    } else {
      mover.forEach(({ f, errores }) => {
        try { if (errores) moverAErrores_(f); else moverArchivo_(f, 'CARPETA_PROCESADOS'); }
        catch (e) { log_('AVISO', 'procesarAlbaranes', f.getName(), `No se ha podido mover el PDF: ${e.message}. La próxima vez se reconocerá como ya registrado.`); }
      });
    }
    const linea = (n, txt) => (n ? `${txt}: ${n}\n` : '');
    let msg = linea(cuenta.creado, 'Albaranes nuevos') + linea(cuenta.completado, 'Albaranes completados con su 2ª hoja') + linea(cuenta.reembolso, 'Reembolsos marcados') + linea(cuenta.duplicado, 'Duplicados ignorados') +
      linea(cuenta.abono, 'Albaranes de abono (a Líneas RM y Abonos)') + linea(cuenta.error, 'No legibles (carpeta Errores)') +
      linea(cuenta.pendiente, 'Sin procesar, siguen en Entrada');
    if (descuadres.length) msg += `\nNo cuadran los importes, siguen en Entrada (${descuadres.length}):\n` + descuadres.map(d => '• ' + d).join('\n') + '\n';
    if (parecidos.length) msg += `\nParecen el mismo albarán con el nº mal leído (${parecidos.length}):\n` + parecidos.map(p => '• ' + p).join('\n') + '\n';
    if (dudas.length) msg += `\nR sin marcar, no se sabe qué pieza es: márcala a mano en Piezas (${dudas.length}):\n` + dudas.map(d => '• ' + d).join('\n') + '\n';
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
  const pdfs = pdfsPorFila_(alb), idsPdf = new Set(Object.keys(pdfs).reduce((a, k) => a.concat(pdfs[k]), []));
  return { alb, trab, piezas, pdfs, idsPdf, iva: cfgNum_('IVA'), hoy: hoyISO_(), tocaAbonos: false, lineas: null, coches };
}

/** Devuelve { estado: 'creado' | 'completado' | 'reembolso' | 'duplicado' | 'abono' | 'descuadre' | 'error', motivo }. */
function procesarDocAlbaran_(ctx, archivo, doc) {
  const ref = archivo.getName(), num = doc.numero_albaran, hoy = ctx.hoy;
  if (ctx.idsPdf.has(archivo.getId())) { log_('AVISO', 'procesarAlbaranes', ref, 'Este PDF ya estaba registrado: se archiva sin cambios'); return { estado: 'duplicado' }; }
  if (!doc.es_albaran) { log_('ERROR', 'procesarAlbaranes', ref, 'El documento no parece un albarán: se mueve a Errores'); return { estado: 'error' }; }
  if (doc.es_abono) {
    // Devolución de RM: no es un gasto del taller, así que no va a Albaranes. Sus líneas van a Líneas RM (sin factura:
    // la quincenal las traerá después) y salen en Abonos: "Abonada" si la pieza se pidió; si no, "Sin solicitar".
    if (!num) log_('AVISO', 'procesarAlbaranes', ref, 'Albarán de abono sin número: cuando llegue la factura RM puede salir repetido en Abonos');
    ctx.lineas = ctx.lineas || leerTabla_(HOJA.LINEAS);
    const n = guardarAbonosEscaneados_(ctx.lineas, ctx.alb, doc, hoy);
    if (!doc.lineas.length) log_('AVISO', 'procesarAlbaranes', ref, `Albarán de abono ${num}: no se ha leído ninguna línea`);
    ctx.tocaAbonos = true;
    log_('INFO', 'procesarAlbaranes', ref, `Albarán de abono ${num} (${doc.total} €): ${n} línea(s) nueva(s) en Líneas RM`);
    return { estado: 'abono' };
  }
  doc = completarTotal(doc, ctx.iva);
  const v = validarAlbaran(doc, ctx.iva);
  if (v.errors.length) { log_('ERROR', 'procesarAlbaranes', ref, v.errors.join('; ') + ': se mueve a Errores'); return { estado: 'error' }; }
  const lineasPiezas = lineasParaPiezas(doc.lineas);

  const existente = num ? ctx.alb.filas.find(f => normAlbaran(f.v['Nº albarán']) === num) : null;
  if (existente) {
    const hoja = hojaQueFalta(existente.v['Precio con IVA'], existente.v['Nota escaneo'], doc, ctx.iva);
    if (hoja === 'completa') return completarConSegundaHoja_(ctx, archivo, doc, existente, lineasPiezas);
    if (hoja === 'no cuadra') {
      const motivo = `parece la 2ª hoja del albarán ${num}, pero ${existente.v['Precio con IVA']} € más sus líneas no dan el total de ${doc.total} €`;
      log_('AVISO', 'procesarAlbaranes', ref, `${motivo}: no se añade y el PDF se queda en Entrada`);
      return { estado: 'descuadre', motivo };
    }
    return procesarSegundoEscaneo_(ctx, archivo, doc, existente, lineasPiezas);
  }

  const igual = buscarNumeroMalLeido(ctx.alb.filas.map(f => ({ fila: f, albaran: f.v['Nº albarán'], plate: f.v['Matrícula'], fecha: aISO_(f.v['Fecha albarán']), total: f.v['Precio con IVA'] })),
    num, doc.matricula, doc.fecha, parseNumber(doc.total));
  if (igual) {
    const registrado = normAlbaran(igual.albaran), parecido = `${ref}: leído como ${num}, igual que el ${registrado} (matrícula, fecha e importe)`;
    log_('AVISO', 'procesarAlbaranes', ref, `Parece el albarán ${registrado} con el nº mal leído (${num}): misma matrícula, fecha e importe. Se trata como un segundo escaneo`);
    return Object.assign(procesarSegundoEscaneo_(ctx, archivo, Object.assign({}, doc, { numero_albaran: registrado }), igual.fila, lineasPiezas), { parecido });
  }

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
    anadirCoche_(ctx.coches, doc.matricula);
    filaAlb = agregarFilas_(ctx.alb, [{ 'Fecha escaneo': aFecha_(hoy), 'Fecha albarán': aFecha_(fecha), 'Proveedor': doc.proveedor, 'Nº albarán': num,
      'Nº trabajo': trabajo.num, 'Matrícula': doc.matricula, 'Precio con IVA': doc.total, 'Ver PDF': enlacePdf_(archivo), 'Nota escaneo': nota }])[0];
    if (trabajo.num) ponerDesplegableTrabajo_(ctx.alb, filaAlb, ctx.trab, doc.matricula);
  }
  ctx.idsPdf.add(archivo.getId());
  ctx.pdfs[filaAlb] = [archivo.getId()];
  // Albarán que no se reconoció como ya registrado (oct 2026: los que perdieron su número o su PDF y se reescanearon)
  // pero cuyas piezas ya están: no se duplican.
  if (num && piezasDeAlbaran_(ctx.piezas, num).length) {
    log_('AVISO', 'procesarAlbaranes', `${HOJA.ALB}!${filaAlb}`, `El albarán ${num} ya tenía sus piezas en Piezas: no se vuelven a añadir`);
  } else if (num) {
    agregarFilas_(ctx.piezas, filasPiezas_(ctx.piezas, lineasPiezas, num, hoy, 'Escaneo'));
    if (lineasPiezas.some(l => l.reembolso)) ctx.tocaAbonos = true;
  } else if (lineasPiezas.length) {
    log_('AVISO', 'procesarAlbaranes', `${HOJA.ALB}!${filaAlb}`, 'Albarán sin número: no se guardan sus piezas (no se podrían reembolsar). Escribe el número y vuelve a escanearlo si hace falta');
  }
  return { estado: 'creado' };
}

/** La hoja que le faltaba a un albarán guardado sin total: pone el total impreso, añade sus piezas y quita el aviso. */
function completarConSegundaHoja_(ctx, archivo, doc, existente, lineasPiezas) {
  const num = doc.numero_albaran;
  actualizarFila_(ctx.alb, existente.fila, { 'Precio con IVA': doc.total, 'Nota escaneo': quitarAvisoSinTotal(existente.v['Nota escaneo']) });
  agregarFilas_(ctx.piezas, filasPiezas_(ctx.piezas, lineasPiezas, num, ctx.hoy, 'Escaneo'));
  if (lineasPiezas.some(l => l.reembolso)) ctx.tocaAbonos = true;
  ctx.pdfs[existente.fila] = (ctx.pdfs[existente.fila] || []).concat(archivo.getId());
  ponerPdfs_(ctx.alb, existente.fila, ctx.pdfs[existente.fila]);
  ctx.idsPdf.add(archivo.getId());
  log_('INFO', 'procesarAlbaranes', `${HOJA.ALB}!${existente.fila}`, `Albarán ${num} completado con su 2ª hoja (${archivo.getName()}): total ${doc.total} €, ${lineasPiezas.length} pieza(s) nueva(s)`);
  return { estado: 'completado' };
}

/** El albarán ya existe: con R -> marca esas piezas; sin R -> escaneo duplicado, no se añade nada. */
function procesarSegundoEscaneo_(ctx, archivo, doc, existente, lineasPiezas) {
  const ref = archivo.getName(), num = doc.numero_albaran, hoy = ctx.hoy;
  const piezasAlb = piezasDeAlbaran_(ctx.piezas, num);
  const hayR = lineasPiezas.some(l => l.reembolso);
  if (!piezasAlb.length && lineasPiezas.length) {  // el primer escaneo se quedó sin piezas: se completan
    agregarFilas_(ctx.piezas, filasPiezas_(ctx.piezas, lineasPiezas, num, hoy, 'Escaneo'));
    log_('INFO', 'procesarAlbaranes', ref, `Albarán ${num} ya existía sin piezas: piezas añadidas`);
    if (hayR) ctx.tocaAbonos = true;
    return { estado: hayR ? 'reembolso' : 'duplicado' };
  }
  if (!hayR) { log_('AVISO', 'procesarAlbaranes', ref, `Albarán ${num} ya existe y no lleva R: escaneo duplicado, se ignora`); return { estado: 'duplicado' }; }
  const plan = aplicarReembolsos(piezasAlb.map(p => ({ ref: p.v['Referencia pieza'], desc: p.v['Descripción'], importe: p.v['Precio descontado sin IVA'], reembolso: p.v['Reembolso'] === true })), doc.lineas);
  const duda = plan.dudosas.length ? `${ref}: albarán ${num}, R en ${plan.dudosas.map(l => l.referencia || l.descripcion).join(', ')}` : '';
  if (duda) log_('AVISO', 'procesarAlbaranes', ref, `Albarán ${num}: hay R en ${plan.dudosas.map(l => l.referencia || l.descripcion).join(', ')}, pero varias piezas tienen ese importe y referencias igual de parecidas. No se marca ninguna: márcala a mano en Piezas`);
  if (!plan.marcar.length && !plan.anadir.length) {
    if (!duda) log_('AVISO', 'procesarAlbaranes', ref, `Albarán ${num}: las piezas con R ya estaban marcadas`);
    return { estado: 'duplicado', duda };
  }
  plan.marcar.forEach(i => actualizarFila_(ctx.piezas, piezasAlb[i].fila, { 'Reembolso': true, 'Fecha reembolso': aFecha_(hoy) }));
  if (plan.anadir.length) agregarFilas_(ctx.piezas, filasPiezas_(ctx.piezas, plan.anadir, num, hoy, 'Escaneo'));
  const resumen = `Reembolso ${hoy.slice(8)}/${hoy.slice(5, 7)}: ${plan.marcar.length + plan.anadir.length} pieza(s)`;
  const previa = String(existente.v['Nota escaneo'] || '');
  actualizarFila_(ctx.alb, existente.fila, { 'Nota escaneo': previa ? previa + ' | ' + resumen : resumen });
  log_('INFO', 'procesarAlbaranes', `${HOJA.ALB}!${existente.fila}`, `${resumen} (albarán ${num})`);
  ctx.tocaAbonos = true;
  return { estado: 'reembolso', duda };
}

/**
 * Ordena Albaranes por Fecha albarán (a igual fecha, por Fecha escaneo; sin fecha, al final). Con el sort de Sheets cada
 * fila se mueve entera: fórmulas, enlaces a los PDF y desplegables van con ella. Si ya está en orden no se toca.
 */
function ordenarAlbaranes_() {
  const t = leerTabla_(HOJA.ALB);
  if (t.filas.length < 2) return false;
  const ini = t.fc + 1, fin = t.filas[t.filas.length - 1].fila;
  const cFa = t.map['Fecha albarán'], cFe = t.map['Fecha escaneo'];
  const clave = v => (v instanceof Date && !isNaN(v) ? v.getTime() : (v === '' || v == null ? Infinity : NaN));
  const filas = t.valores.slice(ini - 1, fin).map(r => [clave(r[cFa - 1]), clave(r[cFe - 1])]);
  const enOrden = filas.every((k, i) => !i || (!isNaN(k[0]) && !isNaN(k[1]) && !isNaN(filas[i - 1][0]) && !isNaN(filas[i - 1][1]) &&
    (filas[i - 1][0] < k[0] || (filas[i - 1][0] === k[0] && filas[i - 1][1] <= k[1]))));
  if (enOrden) return false;
  sinFiltro_(t.sh);
  t.sh.getRange(ini, 1, fin - ini + 1, t.sh.getLastColumn()).sort([{ column: cFa, ascending: true }, { column: cFe, ascending: true }]);
  return true;
}
