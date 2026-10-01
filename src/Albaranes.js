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
    const descuadres = [];
    let parado = '';
    archivos.forEach((f, i) => {
      const r = resultados[i], nombre = f.getName();
      if (parado) { cuenta.pendiente++; return; }
      if (!r.ok) { cuenta.pendiente++; log_('ERROR', 'procesarAlbaranes', nombre, `${r.error} (el PDF sigue en Entrada; vuelve a intentarlo)`); return; }
      // Cada albarán es una unidad: se escribe, se guarda, se relee para comprobarlo y sólo entonces se mueve el PDF.
      // Si algo falla, se deshace lo que haya escrito ese albarán, el PDF se queda en Entrada y se para el lote.
      const deshacer = nuevoDeshacer_(ctx);
      let res;
      try {
        res = procesarDocAlbaran_(ctx, f, r.doc, deshacer);
        if (res.escrito) { SpreadsheetApp.flush(); verificarAlbaran_(ctx, res); }
      } catch (e) {
        let deshecho = 'deshecho lo que había escrito';
        try { deshacer.aplicar(); SpreadsheetApp.flush(); } catch (e2) { deshecho = `NO se ha podido deshacer (${e2.message}): revisa las últimas filas de Albaranes, Piezas y Trabajos`; }
        cuenta.pendiente++;
        parado = nombre;
        log_('ERROR', 'procesarAlbaranes', nombre, `${(e && e.stack) || e} | ${deshecho}. El PDF sigue en Entrada y se paran los siguientes.`);
        return;
      }
      cuenta[res.estado]++;
      if (res.estado === 'descuadre') { descuadres.push(`${nombre}: ${res.motivo}`); return; }  // se queda en Entrada
      try { if (res.estado === 'error') moverAErrores_(f); else moverArchivo_(f, 'CARPETA_PROCESADOS'); }
      catch (e) { log_('AVISO', 'procesarAlbaranes', nombre, `Añadido a la hoja, pero no se ha podido mover el PDF: ${e.message}. La próxima vez se reconocerá como ya registrado.`); }
    });
    if (ctx.tocaAbonos) sincronizarAbonos_();
    const linea = (n, txt) => (n ? `${txt}: ${n}\n` : '');
    let msg = linea(cuenta.creado, 'Albaranes nuevos') + linea(cuenta.reembolso, 'Reembolsos marcados') + linea(cuenta.duplicado, 'Duplicados ignorados') +
      linea(cuenta.abono, 'Albaranes de abono (llegarán en la factura RM)') + linea(cuenta.error, 'No legibles (carpeta Errores)') +
      linea(cuenta.pendiente, 'Sin procesar, siguen en Entrada');
    if (descuadres.length) msg += `\nNo cuadran los importes, siguen en Entrada (${descuadres.length}):\n` + descuadres.map(d => '• ' + d).join('\n') + '\n';
    if (parado) msg += `\n⚠ Se ha parado en ${parado}: no se ha podido guardar en la hoja. Detalle en Registro.`;
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

/**
 * Registro de lo que escribe un albarán, para poder deshacerlo si algo falla: filas añadidas al final de Albaranes,
 * Piezas y Trabajos, y celdas cambiadas en filas que ya existían.
 */
function nuevoDeshacer_(ctx) {
  const tablas = [ctx.alb, ctx.piezas, ctx.trab].map(t => ({ t, libre: t.libre, n: t.filas.length }));
  const cambios = [], ids = [];
  return {
    actualizar(t, fila, nuevos) {
      const f = t.filas.find(x => x.fila === fila), antes = {};
      Object.keys(nuevos).forEach(h => { antes[h] = f ? f.v[h] : ''; });
      cambios.push({ t, fila, antes });
      actualizarFila_(t, fila, nuevos);
    },
    pdf(id) { ids.push(id); ctx.idsPdf.add(id); },
    aplicar() {
      cambios.reverse().forEach(c => actualizarFila_(c.t, c.fila, c.antes));
      tablas.forEach(({ t, libre, n }) => {
        if (t.libre <= libre) return;
        t.sh.getRange(libre, 1, t.libre - libre, anchoTabla_(t)).clearContent().clearDataValidations();
        t.libre = libre;
        t.filas.length = n;
      });
      ids.forEach(id => ctx.idsPdf.delete(id));
    },
  };
}

/** Relee de la hoja lo que acaba de escribir un albarán. Si no está, lanza un error (y el albarán se deshace). */
function verificarAlbaran_(ctx, res) {
  const fallos = [];
  if (res.filaAlb) {
    const t = ctx.alb, v = t.sh.getRange(res.filaAlb, 1, 1, anchoTabla_(t)).getValues()[0], g = h => v[t.map[h] - 1];
    if (normAlbaran(g('Nº albarán')) !== res.num) fallos.push(`nº de albarán (${g('Nº albarán')})`);
    if (!(Math.abs(parseNumber(g('Precio con IVA')) - res.total) <= 0.01)) fallos.push(`total (${g('Precio con IVA')})`);
    if (normPlate(g('Matrícula')) !== normPlate(res.matricula)) fallos.push(`matrícula (${g('Matrícula')})`);
    if (idDeEnlace_(t.sh.getRange(res.filaAlb, t.map['Ver PDF']).getFormula()) !== res.idPdf) fallos.push('enlace al PDF');
  }
  if (res.piezas && res.piezas.n) {
    const col = ctx.piezas.sh.getRange(res.piezas.desde, ctx.piezas.map['Nº albarán'], res.piezas.n, 1).getValues();
    if (col.some(r => normAlbaran(r[0]) !== res.num)) fallos.push('piezas');
  }
  (res.marcadas || []).forEach(fila => {
    if (ctx.piezas.sh.getRange(fila, ctx.piezas.map['Reembolso']).getValue() !== true) fallos.push(`reembolso en Piezas!${fila}`);
  });
  if (fallos.length) throw new Error(`Después de guardar, la hoja no tiene lo esperado: ${fallos.join(', ')}`);
}

function contextoAlbaranes_() {
  const alb = leerTabla_(HOJA.ALB), trab = leerTabla_(HOJA.TRAB), piezas = leerTabla_(HOJA.PIEZAS), coches = leerTabla_(HOJA.COCHES);
  const idsPdf = new Set();
  const ult = alb.sh.getLastRow();
  if (ult > 1) alb.sh.getRange(2, alb.map['Ver PDF'], ult - 1, 1).getFormulas().forEach(r => { const id = idDeEnlace_(r[0]); if (id) idsPdf.add(id); });
  return { alb, trab, piezas, idsPdf, iva: cfgNum_('IVA'), hoy: hoyISO_(), tocaAbonos: false,
    coches: new Set(coches.filas.map(f => normPlate(f.v['Matrícula']))) };
}

/**
 * Devuelve { estado: 'creado' | 'reembolso' | 'duplicado' | 'abono' | 'descuadre' | 'error', escrito, ... } con lo
 * necesario para verificarAlbaran_. Todo lo que escribe pasa por `deshacer` (agregarFilas_ al final de la tabla, o
 * deshacer.actualizar para filas que ya existían).
 */
function procesarDocAlbaran_(ctx, archivo, doc, deshacer) {
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
  if (existente) return procesarSegundoEscaneo_(ctx, archivo, doc, existente, lineasPiezas, deshacer);

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
  let filaAlb, matricula = doc.matricula;
  if (manual) {
    filaAlb = manual.row;
    matricula = manual.plate;
    deshacer.actualizar(ctx.alb, filaAlb, { 'Fecha escaneo': aFecha_(hoy), 'Nº albarán': num, 'Fecha albarán': aFecha_(fecha), 'Proveedor': doc.proveedor,
      'Ver PDF': enlacePdf_(archivo), 'Nota escaneo': ('Vinculado a la fila que ya estaba escrita a mano. ' + nota).trim() });
    log_('INFO', 'procesarAlbaranes', `${HOJA.ALB}!${filaAlb}`, `Albarán ${num} vinculado a la fila manual (misma matrícula e importe)`);
  } else {
    const trabajo = doc.matricula ? asignarTrabajo_(ctx.trab, doc.matricula, fecha) : { num: '' };
    if (doc.matricula && !ctx.coches.has(doc.matricula)) log_('AVISO', 'procesarAlbaranes', ref, `La matrícula ${doc.matricula} no está en Coches`);
    filaAlb = agregarFilas_(ctx.alb, [{ 'Fecha escaneo': aFecha_(hoy), 'Fecha albarán': aFecha_(fecha), 'Proveedor': doc.proveedor, 'Nº albarán': num,
      'Nº trabajo': trabajo.num, 'Matrícula': doc.matricula, 'Precio con IVA': doc.total, 'Ver PDF': enlacePdf_(archivo), 'Nota escaneo': nota }])[0];
    if (trabajo.num) ponerDesplegableTrabajo_(ctx.alb, filaAlb, ctx.trab, doc.matricula);
  }
  deshacer.pdf(archivo.getId());
  const piezas = { desde: ctx.piezas.libre, n: 0 };
  if (num) {
    piezas.n = agregarFilas_(ctx.piezas, lineasPiezas.map(l => filaPieza_(l, num, hoy, 'Escaneo'))).length;
    if (lineasPiezas.some(l => l.reembolso)) ctx.tocaAbonos = true;
  } else if (lineasPiezas.length) {
    log_('AVISO', 'procesarAlbaranes', `${HOJA.ALB}!${filaAlb}`, 'Albarán sin número: no se guardan sus piezas (no se podrían reembolsar). Escribe el número y vuelve a escanearlo si hace falta');
  }
  return { estado: 'creado', escrito: true, filaAlb, num, total: doc.total, matricula, idPdf: archivo.getId(), piezas };
}

/** El albarán ya existe: con R -> marca esas piezas; sin R -> escaneo duplicado, no se añade nada. */
function procesarSegundoEscaneo_(ctx, archivo, doc, existente, lineasPiezas, deshacer) {
  const ref = archivo.getName(), num = doc.numero_albaran, hoy = ctx.hoy;
  const piezasAlb = piezasDeAlbaran_(ctx.piezas, num);
  const hayR = lineasPiezas.some(l => l.reembolso);
  if (!piezasAlb.length && lineasPiezas.length) {  // el primer escaneo se quedó sin piezas: se completan
    const piezas = { desde: ctx.piezas.libre, n: agregarFilas_(ctx.piezas, lineasPiezas.map(l => filaPieza_(l, num, hoy, 'Escaneo'))).length };
    log_('INFO', 'procesarAlbaranes', ref, `Albarán ${num} ya existía sin piezas: piezas añadidas`);
    if (hayR) ctx.tocaAbonos = true;
    return { estado: hayR ? 'reembolso' : 'duplicado', escrito: true, num, piezas };
  }
  if (!hayR) { log_('AVISO', 'procesarAlbaranes', ref, `Albarán ${num} ya existe y no lleva R: escaneo duplicado, se ignora`); return { estado: 'duplicado' }; }
  const plan = aplicarReembolsos(piezasAlb.map(p => ({ ref: p.v['Referencia pieza'], desc: p.v['Descripción'], reembolso: p.v['Reembolso'] === true })), doc.lineas);
  if (!plan.marcar.length && !plan.anadir.length) { log_('AVISO', 'procesarAlbaranes', ref, `Albarán ${num}: las piezas con R ya estaban marcadas`); return { estado: 'duplicado' }; }
  const marcadas = plan.marcar.map(i => piezasAlb[i].fila);
  marcadas.forEach(fila => deshacer.actualizar(ctx.piezas, fila, { 'Reembolso': true, 'Fecha reembolso': aFecha_(hoy) }));
  const piezas = { desde: ctx.piezas.libre, n: plan.anadir.length ? agregarFilas_(ctx.piezas, plan.anadir.map(l => filaPieza_(l, num, hoy, 'Escaneo'))).length : 0 };
  const resumen = `Reembolso ${hoy.slice(8)}/${hoy.slice(5, 7)}: ${plan.marcar.length + plan.anadir.length} pieza(s)`;
  const previa = String(existente.v['Nota escaneo'] || '');
  deshacer.actualizar(ctx.alb, existente.fila, { 'Nota escaneo': previa ? previa + ' | ' + resumen : resumen });
  log_('INFO', 'procesarAlbaranes', `${HOJA.ALB}!${existente.fila}`, `${resumen} (albarán ${num})`);
  ctx.tocaAbonos = true;
  return { estado: 'reembolso', escrito: true, num, piezas, marcadas };
}
