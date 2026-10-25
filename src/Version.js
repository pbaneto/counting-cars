/**
 * Versiones de la hoja. La hoja guarda en qué versión está (propiedad VERSION_HOJA) y el código sabe la suya (VERSION,
 * en Config.js). Al empezar una acción del menú o al abrir la hoja, si no coinciden:
 *  - se ejecutan, en orden y una sola vez, las migraciones de las versiones intermedias (MIGRACIONES);
 *  - si cambia la versión mayor o la menor, se vuelve a aplicar el diseño (aplicarDiseno_);
 *  - se apunta la versión nueva y se deja constancia en Registro.
 * Si coinciden, sólo se lee una propiedad.
 */

const PROP_VERSION_ = 'VERSION_HOJA';
// Diseño empezado y no terminado (p. ej. un error a medias): se repite la próxima vez aunque la versión ya esté apuntada.
const PROP_DISENO_PENDIENTE_ = 'DISENO_PENDIENTE';

/**
 * Migraciones de datos, por versión: { '2.0.0': () => { ... } }. Deben ser idempotentes (si se cortan a medias, se
 * repiten enteras). Antes de escribir en celdas con desplegable, quitarlo: un valor fuera de su lista hace fallar la
 * escritura (oct 2026: la 3.0.0 escribió fórmulas donde había quedado el desplegable de Estado y cortó el diseño). Cuando una ya se ha ejecutado en la hoja real, se borra de aquí en el siguiente cambio: la
 * versión guardada en la hoja impide que vuelva a hacer falta.
 */
const MIGRACIONES = {
  // Arreglos de lo que estropearon los filtros puestos en Albaranes, Trabajos y Líneas RM (ver sinFiltro_ en Sheets.js):
  //  - Líneas RM: fechas de albarán y de abono que se perdieron (datos en src/private.js);
  //  - el albarán 497296, cuya fecha corregida en la 6.0.0 no se guardó;
  //  - Piezas: las piezas de los albaranes reescaneados el 19/10 están dos veces. Si TODAS las líneas de un albarán están
  //    exactamente dos veces, se queda una de cada par: la enlazada a un abono, si no la pedida, si no la que tiene marca.
  // Y además: fuera las claves de Config que ya no se usan y el resumen que había a la derecha de Abonos (ahora en Cuadre RM).
  '7.0.0': () => {
    const rec = typeof RECUPERAR_FECHAS_7_0_0 === 'undefined' ? {} : RECUPERAR_FECHAS_7_0_0;
    const escaneados = new Set(typeof ABONOS_ESCANEADOS_7_0_0 === 'undefined' ? [] : ABONOS_ESCANEADOS_7_0_0);
    const fechaFactura = {};
    leerTabla_(HOJA.FACT).filas.forEach(f => { fechaFactura[String(f.v['Nº factura']).trim()] = aISO_(f.v['Fecha factura']); });
    const tabL = leerTabla_(HOJA.LINEAS);
    let fechas = 0;
    tabL.filas.forEach(f => {
      const fra = String(f.v['Nº factura'] || '').trim(), alb = normAlbaran(f.v['Nº albarán']);
      let fa = aISO_(f.v['Fecha albarán']);
      if (!fa && rec[`${fra}|${alb}`]) { fa = rec[`${fra}|${alb}`]; f.v['Fecha albarán'] = aFecha_(fa); fechas++; }
      if (f.v['Tipo'] === 'Abono' && !aISO_(f.v['Fecha abono'])) {
        const v = !fra || escaneados.has(alb) ? fa : fechaFactura[fra] || fa;
        if (v) f.v['Fecha abono'] = aFecha_(v);
      }
    });
    ['Fecha albarán', 'Fecha abono'].forEach(h => escribirColumna_(tabL, h));

    const alb = leerTabla_(HOJA.ALB);
    const mal = alb.filas.find(f => normAlbaran(f.v['Nº albarán']) === '497296' && aISO_(f.v['Fecha albarán']) === '2026-02-10');
    if (mal) actualizarFila_(alb, mal.fila, { 'Fecha albarán': aFecha_('2026-10-02') });

    const tabP = leerTabla_(HOJA.PIEZAS), enlazadas = new Set(tabL.filas.map(f => idPiezaDe(f.v['ID pieza'])).filter(Boolean));
    const porAlbaran = {};
    tabP.filas.forEach(p => { const n = normAlbaran(p.v['Nº albarán']); if (n) (porAlbaran[n] = porAlbaran[n] || []).push(p); });
    const firma = p => [refKey(p.v['Referencia pieza']), refKey(p.v['Descripción']), p.v['Cantidad'], round2(parseNumber(p.v['Precio descontado sin IVA']))].join('|');
    const valor = p => (enlazadas.has(String(p.v['ID'])) ? 4 : 0) + (p.v['Reembolso'] === true ? 2 : 0) + (p.v['Marca'] ? 1 : 0);
    const borrar = [];
    Object.keys(porAlbaran).forEach(n => {
      const pares = {};
      porAlbaran[n].forEach(p => (pares[firma(p)] = pares[firma(p)] || []).push(p));
      if (!Object.keys(pares).every(k => pares[k].length === 2)) return;
      Object.keys(pares).forEach(k => {
        const [queda, sobra] = pares[k].slice().sort((x, y) => valor(y) - valor(x));
        const fechas = [queda, sobra].map(p => aISO_(p.v['Fecha reembolso'])).filter(Boolean).sort();
        const cambios = {};
        if (sobra.v['Reembolso'] === true && queda.v['Reembolso'] !== true) cambios['Reembolso'] = true;
        if (fechas.length && fechas[0] !== aISO_(queda.v['Fecha reembolso'])) cambios['Fecha reembolso'] = aFecha_(fechas[0]);
        if (!queda.v['Marca'] && sobra.v['Marca']) cambios['Marca'] = sobra.v['Marca'];
        if (!queda.v['Abonada a mano'] && sobra.v['Abonada a mano']) cambios['Abonada a mano'] = sobra.v['Abonada a mano'];
        if (Object.keys(cambios).length) actualizarFila_(tabP, queda.fila, cambios);
        borrar.push(sobra.fila);
      });
    });
    borrarFilas_(tabP.sh, borrar);

    const conocidas = new Set(CONFIG_DEFECTO.map(c => c[0]));
    const tabC = leerTabla_(HOJA.CONFIG);
    borrarFilas_(tabC.sh, tabC.filas.filter(f => !conocidas.has(String(f.v['Clave']).trim())).map(f => f.fila));

    const shAb = hoja_(HOJA.ABONOS), resto = shAb.getMaxColumns() - ABONOS.cabTabla.length;
    if (resto > 0) shAb.getRange(1, ABONOS.cabTabla.length + 1, shAb.getMaxRows(), resto).breakApart().clearContent().clearFormat().clearDataValidations();
    shAb.setConditionalFormatRules((shAb.getConditionalFormatRules() || []).filter(r => {
      try { return r.getRanges().some(g => g.getColumn() <= ABONOS.cabTabla.length); } catch (e) { return true; }
    }));
    log_('INFO', 'migración 7.0.0', '', `${fechas} fechas recuperadas en Líneas RM; ${borrar.length} piezas duplicadas borradas` + (mal ? '; albarán 497296: fecha corregida' : ''));
    actualizarInformeAbonos_();
  },
};

/** Escribe una columna entera de la tabla (de la primera a la última fila con datos) con los valores de tabla.filas. */
function escribirColumna_(tabla, h) {
  if (!tabla.filas.length) return;
  const desde = tabla.filas[0].fila, hasta = tabla.filas[tabla.filas.length - 1].fila, porFila = {};
  tabla.filas.forEach(f => { porFila[f.fila] = f.v[h]; });
  sinFiltro_(tabla.sh);
  tabla.sh.getRange(desde, tabla.map[h], hasta - desde + 1, 1)
    .setValues(Array.from({ length: hasta - desde + 1 }, (_, k) => [porFila[desde + k] == null ? '' : porFila[desde + k]]));
}

/** Pone la hoja al día con el código. forzarDiseno: aplicar el diseño aunque la versión no haya cambiado (menú). */
function actualizarHoja_(forzarDiseno) {
  const props = PropertiesService.getScriptProperties();
  const antes = props.getProperty(PROP_VERSION_) || '1.0.0';  // hojas anteriores al sistema de versiones: 1.0.0
  const pendiente = props.getProperty(PROP_DISENO_PENDIENTE_) === '1';
  if (antes === VERSION && !forzarDiseno && !pendiente) return false;
  if (compararVersiones(antes, VERSION) > 0) {
    log_('AVISO', 'actualizarHoja', '', `La hoja está en la versión ${antes}, más nueva que el código (${VERSION}): no se toca`);
    return false;
  }
  Object.keys(MIGRACIONES).filter(v => compararVersiones(v, antes) > 0 && compararVersiones(v, VERSION) <= 0).sort(compararVersiones).forEach(v => {
    MIGRACIONES[v]();
    props.setProperty(PROP_VERSION_, v);
    log_('INFO', 'actualizarHoja', v, `Migración ${v} aplicada`);
  });
  if (forzarDiseno || pendiente || cambiaDiseno(antes, VERSION)) {
    props.setProperty(PROP_DISENO_PENDIENTE_, '1');
    aplicarDiseno_();
    props.deleteProperty(PROP_DISENO_PENDIENTE_);
  }
  if (antes !== VERSION) {
    props.setProperty(PROP_VERSION_, VERSION);
    log_('INFO', 'actualizarHoja', '', `Hoja actualizada de la versión ${antes} a la ${VERSION}`);
  }
  return true;
}

/** Menú: versión del código y de la hoja. */
function verVersion() {
  const hoja = PropertiesService.getScriptProperties().getProperty(PROP_VERSION_) || '1.0.0';
  avisar_(`Versión del código: ${VERSION}\nVersión de la hoja: ${hoja}`, 'Counting Cars');
}
