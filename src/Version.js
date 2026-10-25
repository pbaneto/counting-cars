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
  // Abonos pasa a ser un informe que se rehace entero (ver Abonos.js). Lo que sólo estaba en su tabla pasa a su sitio:
  //  - Piezas: columnas "Abonada a mano" (tras "Fecha reembolso") e "ID" (al final), con un ID para cada pieza;
  //  - Líneas RM: columnas "ID pieza" y "Fecha abono"; cada línea de abono recibe la pieza y la fecha de abono que tenía
  //    su fila en Abonos, y los albaranes de abono escaneados que aún no venían en ninguna factura se añaden como líneas;
  //  - Abonos: la columna oculta "Clave" pasa a ser "ID pieza", sin el desplegable de Estado; fuera "Abonos (claves)".
  // Sólo rellena lo que esté vacío y no añade dos veces la misma línea: se puede repetir entera.
  '5.0.0': () => {
    const ss = ss_(), shP = hoja_(HOJA.PIEZAS), shL = hoja_(HOJA.LINEAS), shAb = hoja_(HOJA.ABONOS);
    const cabDe = sh => sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0].map(x => String(x).trim());
    let cab = cabDe(shP);
    if (cab.indexOf('Abonada a mano') < 0) {
      const c = cab.indexOf('Fecha reembolso') + 1;
      if (!c) throw new Error('Migración 5.0.0: no encuentro la columna "Fecha reembolso" en Piezas.');
      shP.insertColumnBefore(c + 1);
      shP.getRange(1, c + 1).setValue('Abonada a mano');
      cab = cabDe(shP);
    }
    const ponerCabecera = (sh, h, cab) => {
      if (cab.indexOf(h) >= 0) return;
      const c = ESQUEMA[sh.getName()].cabeceras.indexOf(h) + 1;
      if (cab[c - 1]) throw new Error(`Migración 5.0.0: la columna ${colLetra_(c)} de "${sh.getName()}" debería quedar libre para "${h}" y tiene "${cab[c - 1]}".`);
      sh.getRange(1, c).setValue(h);
    };
    ponerCabecera(shP, 'ID', cab);
    ponerCabecera(shL, 'ID pieza', cabDe(shL));
    ponerCabecera(shL, 'Fecha abono', cabDe(shL));
    reiniciarCaches_();

    const tabP = leerTabla_(HOJA.PIEZAS), tabL = leerTabla_(HOJA.LINEAS), tabA = leerTabla_(HOJA.ALB);
    const fechaFactura = {};
    leerTabla_(HOJA.FACT).filas.forEach(f => { fechaFactura[String(f.v['Nº factura']).trim()] = f.v['Fecha factura']; });
    const ncol = ABONOS.cabTabla.length, ult = shAb.getLastRow();
    const tabla = ult >= ABONOS.filaTabla ? shAb.getRange(ABONOS.filaCabTabla, 1, ult - ABONOS.filaCabTabla + 1, ncol).getValues() : [[]];
    const cabAb = tabla[0].map(x => String(x).trim()), col = h => cabAb.indexOf(h);
    const iClave = col('Clave') >= 0 ? col('Clave') : col('ID pieza');
    asegurarIdsPiezas_(tabP);

    // Claves de la versión anterior: P|albarán|referencia|n contaba sólo las piezas de RM, en el orden de Piezas.
    const proveedorDe = {};
    tabA.filas.forEach(f => { const n = normAlbaran(f.v['Nº albarán']); if (n) proveedorDe[n] = f.v['Proveedor']; });
    const idDeP = {};
    clavesPiezas(tabP.filas.filter(p => String(proveedorDe[normAlbaran(p.v['Nº albarán'])] || 'RM') !== 'Otros')
      .map(p => ({ id: p.v['ID'], albaran: p.v['Nº albarán'], ref: p.v['Referencia pieza'], desc: p.v['Descripción'] }))).forEach(p => { idDeP[p.clave] = p.id; });
    const lineas = lineasAbono_(tabL), lineaDeA = {};
    lineas.forEach(x => { lineaDeA[x.clave] = x.f; });

    const cambios = new Map(), sueltas = [];
    const cambiar = (f, h, v) => { if (v === '' || v == null || (f.v[h] !== '' && f.v[h] != null)) return; f.v[h] = v; cambios.set(f, true); };
    tabla.slice(1).forEach(r => {
      const ks = iClave >= 0 ? partirClaves(r[iClave]) : [], kA = ks.find(k => k.indexOf('A|') === 0), kP = ks.find(k => k.indexOf('P|') === 0);
      if (!kA) return;
      const id = kP && idDeP[kP] !== undefined ? idDeP[kP] : '', fecha = r[col('Fecha abono')];
      const f = lineaDeA[kA];
      if (f) { cambiar(f, 'ID pieza', id); cambiar(f, 'Fecha abono', fecha); return; }
      // Albarán de abono escaneado que aún no ha llegado en ninguna factura: sólo estaba en Abonos.
      const [, albAbono, orig] = kA.split('|');
      sueltas.push({ 'Nº factura': r[col('Factura RM')] || '', 'Nº albarán': albAbono, 'Fecha albarán': fecha, 'Tipo': 'Abono', 'Matrícula': r[col('Matrícula')],
        'Albarán origen': orig, 'Referencia': r[col('Referencia')], 'Descripción': r[col('Descripción pieza')], 'Importe sin IVA': -Math.abs(parseNumber(r[col('Precio sin IVA')]) || 0),
        'ID pieza': id, 'Fecha abono': fecha });
      lineaDeA[kA] = { v: {} };
    });
    // Líneas de abono que no estaban en la tabla: la fecha de su factura.
    lineas.forEach(({ f }) => cambiar(f, 'Fecha abono', fechaFactura[String(f.v['Nº factura']).trim()] || f.v['Fecha albarán']));
    ['ID pieza', 'Fecha abono'].forEach(h => {
      if (!tabL.filas.length) return;
      const desde = tabL.filas[0].fila, hasta = tabL.filas[tabL.filas.length - 1].fila, porFila = {};
      tabL.filas.forEach(f => { porFila[f.fila] = f.v[h]; });
      shL.getRange(desde, tabL.map[h], hasta - desde + 1, 1).setValues(Array.from({ length: hasta - desde + 1 }, (_, k) => [porFila[desde + k] == null ? '' : porFila[desde + k]]));
    });
    agregarFilas_(tabL, sueltas);

    if (iClave >= 0 && cabAb[iClave] === 'Clave') shAb.getRange(ABONOS.filaCabTabla, iClave + 1).setValue(CAMPOS_ABONOS_.idPieza);
    shAb.getRange(ABONOS.filaTabla, 1, Math.max(shAb.getMaxRows() - ABONOS.filaTabla + 1, 1), ncol).clearDataValidations();
    const shC = ss.getSheetByName('Abonos (claves)');
    if (shC) ss.deleteSheet(shC);
    _colsAbonos = null;
    log_('INFO', 'migración 5.0.0', '', `${tabP.filas.length} piezas con ID; ${cambios.size} líneas de abono completadas; ${sueltas.length} abonos escaneados pasados a Líneas RM`);
    actualizarInformeAbonos_(tabP);
  },
};

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
