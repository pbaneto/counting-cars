/** Facturas quincenales de RM: leer con Gemini, validar el IVA, guardar cabecera + líneas (sus abonos salen en Abonos). */

function procesarFacturasRM() {
  ejecutar_('procesarFacturasRM', () => conBloqueo_(10, () => {
    actualizarHoja_(false);
    const archivos = listarArchivos_('CARPETA_FACTURAS_RM', 5);
    if (!archivos.length) { avisar_('No hay facturas nuevas en la carpeta de facturas RM.'); return; }
    toast_(`Leyendo ${archivos.length} factura(s) con Gemini…`, 'Facturas RM', 60);
    const resultados = leerConGemini_(archivos, false);
    const cfgErr = resultados.find(r => !r.ok && r.config);
    if (cfgErr) throw new Error(cfgErr.error);
    releerFacturasDescuadradas_(archivos, resultados);
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
    actualizarInformeAbonos_();
    avisar_(lineas.join('\n'), 'Facturas RM');
  }));
}

/** Bloques cuyas líneas no suman el importe del bloque (suele ser un bloque partido entre dos hojas). */
function bloquesDescuadrados_(doc) { return validarFactura(doc, cfgNum_('IVA')).warnings.filter(w => /^Albarán .*: líneas/.test(w)); }

/** Segunda lectura de las facturas con bloques descuadrados; se queda la lectura con menos bloques descuadrados. */
function releerFacturasDescuadradas_(archivos, resultados) {
  const repetir = resultados.map((r, i) => (r.ok && bloquesDescuadrados_(r.doc).length ? i : -1)).filter(i => i >= 0);
  if (!repetir.length) return;
  toast_(`Releyendo ${repetir.length} factura(s) con bloques que no cuadran…`, 'Facturas RM', 60);
  const pistas = repetir.map(i => 'ATENCIÓN: en una primera lectura no cuadraban estos bloques (' + bloquesDescuadrados_(resultados[i].doc).join('; ') +
    '). Seguramente faltan líneas que siguen en la hoja siguiente, debajo de la cabecera de columnas y sin fila de ALBARAN: son del mismo bloque.');
  leerConGemini_(repetir.map(i => archivos[i]), false, pistas).forEach((s, k) => {
    const i = repetir[k];
    if (s.ok && bloquesDescuadrados_(s.doc).length < bloquesDescuadrados_(resultados[i].doc).length) {
      resultados[i] = s;
      log_('INFO', 'procesarFacturasRM', archivos[i].getName(), 'Bloques corregidos en la segunda lectura');
    }
  });
}

/** Guarda (o reemplaza) una factura. Devuelve {fatal, estado, numero, periodo}. */
function registrarFactura_(archivo, doc, truncado) {
  const ref = archivo.getName(), iva = cfgNum_('IVA');
  const v = validarFactura(doc, iva);
  if (truncado) v.warnings.push('La respuesta de Gemini se cortó: puede faltar información');
  if (v.errors.length) { log_('ERROR', 'registrarFactura', ref, v.errors.join('; ')); return { fatal: true, estado: v.errors.join('; ') }; }
  const per = periodoFactura(doc);

  let tabL = leerTabla_(HOJA.LINEAS);
  // Lo que ya se sabía de cada línea de abono (su pieza y su fecha de abono) pasa a la línea nueva con la misma clave: de
  // esta misma factura si se procesa otra vez, o de su albarán de abono escaneado antes, que deja de estar suelto.
  const sabido = {}, sustituidas = new Set(), conocidas = [];
  lineasAbono_(tabL).forEach(x => {
    const fra = String(x.f.v['Nº factura']).trim();
    if (fra === doc.numero_factura || fra === '') { sabido[x.clave] = x.f; conocidas.push({ f: x.f, clave: x.clave, albaranAbono: x.albaranAbono, importe: x.f.v['Importe sin IVA'] }); }
  });
  tabL.filas.forEach(f => { if (String(f.v['Nº factura']).trim() === doc.numero_factura) sustituidas.add(f); });
  // Los bloques de abono traen "MATRICULA:" vacía: se pone la del albarán original (de esta factura, de otra ya
  // procesada o de la pestaña Albaranes). Si el original tampoco la tiene, se queda vacía.
  const matriculaDe = mapaMatriculas(doc.albaranes.filter(a => !a.es_abono).map(a => [a.numero_albaran, a.matricula])
    .concat(tabL.filas.filter(f => f.v['Tipo'] === 'Compra').map(f => [f.v['Nº albarán'], f.v['Matrícula']]))
    .concat(leerTabla_(HOJA.ALB).filas.map(f => [f.v['Nº albarán'], f.v['Matrícula']])));
  const filas = [];
  doc.albaranes.forEach(a => {
    const tipo = a.es_abono ? 'Abono' : 'Compra';
    const ls = a.lineas.length ? a.lineas : [{ referencia: '', descripcion: '(sin detalle de líneas)', importe: a.importe, albaran_origen: '' }];
    ls.forEach(l => filas.push({
      'Nº factura': doc.numero_factura, 'Nº albarán': a.numero_albaran, 'Fecha albarán': aFecha_(a.fecha), 'Tipo': tipo,
      'Matrícula': a.matricula || (tipo === 'Abono' ? matriculaDe[albaranOrigen(l.albaran_origen)] || '' : ''),
      'Albarán origen': tipo === 'Abono' ? l.albaran_origen : '', 'Referencia': l.referencia, 'Descripción': l.descripcion,
      'Cantidad': l.cantidad == null ? '' : l.cantidad, 'Precio': l.precio_unitario == null ? '' : l.precio_unitario,
      'Descuento': l.descuento_pct == null ? '' : l.descuento_pct / 100, 'Importe sin IVA': l.importe == null ? 0 : l.importe,
    }));
  });
  const fechaFactura = aFecha_(doc.fecha_factura);
  const leidas = clavesAbonos(filas.filter(f => f['Tipo'] === 'Abono').map(f => ({ f, albaranAbono: normAlbaran(f['Nº albarán']), albaranOrigen: albaranOrigen(f['Albarán origen']),
    ref: f['Referencia'], desc: f['Descripción'] })));
  const delBloque = n => leidas.filter(x => x.albaranAbono === normAlbaran(n));
  const omitidas = lineasOmitidasDeBloques(doc.albaranes.filter(a => a.es_abono).map(a => ({ numero: a.numero_albaran, importe: a.importe,
    leidas: delBloque(a.numero_albaran).map(x => x.clave), suma: delBloque(a.numero_albaran).reduce((s, x) => s + (parseNumber(x.f['Importe sin IVA']) || 0), 0) })), conocidas);
  const completados = Object.keys(omitidas);
  completados.forEach(n => omitidas[n].forEach(({ f }) => {
    const copia = { 'Nº factura': doc.numero_factura };
    ['Nº albarán', 'Fecha albarán', 'Tipo', 'Matrícula', 'Albarán origen', 'Referencia', 'Descripción', 'Cantidad', 'Precio', 'Descuento', 'Importe sin IVA', 'ID pieza', 'Fecha abono'].forEach(h => { copia[h] = f.v[h]; });
    filas.push(copia);
    sustituidas.add(f);
    log_('INFO', 'registrarFactura', doc.numero_factura, `Albarán ${n}: la lectura se ha saltado ${f.v['Referencia'] || f.v['Descripción']}; se conserva la línea ya guardada`);
  }));
  leidas.forEach(({ f, clave }) => {
    const antes = sabido[clave];
    if (antes) sustituidas.add(antes);
    f['ID pieza'] = antes ? antes.v['ID pieza'] : '';
    // Escaneado antes sin fecha: la del bloque de la factura, que es la del albarán de abono.
    f['Fecha abono'] = antes ? antes.v['Fecha abono'] || f['Fecha albarán'] : fechaFactura;
  });
  // Reemplazo idempotente: fuera las líneas anteriores de esta misma factura y los abonos escaneados que ya trae.
  if (sustituidas.size) {
    borrarFilas_(tabL.sh, [...sustituidas].map(f => f.fila));
    tabL = leerTabla_(HOJA.LINEAS);
  }
  agregarFilas_(tabL, filas);

  const avisos = v.warnings.filter(w => !completados.some(n => w.indexOf(`Albarán ${n}: líneas`) === 0));
  avisos.forEach(w => log_('AVISO', 'registrarFactura', `${doc.numero_factura}`, w));
  const estado = avisos.length ? '⚠ ' + avisos.join('; ') : 'OK';
  const tabF = leerTabla_(HOJA.FACT);
  const fila = { 'Nº factura': doc.numero_factura, 'Fecha factura': aFecha_(doc.fecha_factura), 'Año': per.year, 'Mes': per.month, 'Quincena': per.quincena,
    'Base imponible': doc.base_imponible, 'IVA': doc.iva_importe, 'Total': doc.total, 'Estado': estado, 'Ver PDF': enlacePdf_(archivo),
    'Procesada el': new Date(), 'Nº albaranes': doc.albaranes.length };
  const prev = tabF.filas.find(f => String(f.v['Nº factura']).trim() === doc.numero_factura);
  if (prev) actualizarFila_(tabF, prev.fila, fila); else agregarFilas_(tabF, [fila]);
  log_('INFO', 'registrarFactura', doc.numero_factura, `${doc.albaranes.length} albaranes, total ${doc.total}: ${estado}`);
  return { fatal: false, estado, numero: doc.numero_factura, periodo: per };
}
