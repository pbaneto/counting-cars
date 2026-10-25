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
  // Abonos y Piezas empiezan en octubre de 2026: lo de antes era el piloto.
  //  - Abonos sólo enseña lo de octubre en adelante (Config ▸ ABONOS_DESDE) y lo que siga sin abonar. Las líneas de las
  //    facturas de antes se quedan en Líneas RM: son las que cuadran con Facturas RM.
  //  - Fuera de Piezas las de albaranes anteriores, salvo las pedidas: restan en el precio de su albarán (y de su trabajo)
  //    y su abono puede llegar en una factura de octubre.
  //  - El albarán 497296 se leyó con fecha 10/02/26 en vez de 02/10/26 (así viene en la factura FCR 26422).
  '6.0.0': () => {
    const desde = '2026-10-01';
    const alb = leerTabla_(HOJA.ALB);
    const mal = alb.filas.find(f => normAlbaran(f.v['Nº albarán']) === '497296' && aISO_(f.v['Fecha albarán']) === '2026-02-10');
    if (mal) actualizarFila_(alb, mal.fila, { 'Fecha albarán': aFecha_('2026-10-02') });
    if (!aISO_(cfg_('ABONOS_DESDE'))) guardarCfg_('ABONOS_DESDE', aFecha_(desde));
    const fechaDe = {};
    alb.filas.forEach(f => { const n = normAlbaran(f.v['Nº albarán']); if (n) fechaDe[n] = aISO_(f.v['Fecha albarán']); });
    const tabP = leerTabla_(HOJA.PIEZAS);
    const fuera = tabP.filas.filter(p => {
      const fa = fechaDe[normAlbaran(p.v['Nº albarán'])];
      return fa && fa < desde && p.v['Reembolso'] !== true && !p.v['Abonada a mano'];
    });
    const ids = new Set(fuera.map(p => String(p.v['ID'])));
    fuera.map(p => p.fila).sort((a, b) => b - a).forEach(r => tabP.sh.deleteRow(r));
    const tabL = leerTabla_(HOJA.LINEAS);
    const sueltas = tabL.filas.filter(f => ids.has(String(f.v['ID pieza'])));
    // "NO" y no vacío: vacío, se emparejaría solo (p. ej. por importe) con otra pieza del mismo albarán.
    sueltas.forEach(f => actualizarFila_(tabL, f.fila, { 'ID pieza': SIN_PIEZA }));
    log_('INFO', 'migración 6.0.0', '', `${fuera.length} piezas anteriores a octubre borradas (${sueltas.length} abonos que apuntaban a ellas, sin pieza)` +
      (mal ? '; albarán 497296: fecha corregida a 02/10/2026' : ''));
    actualizarInformeAbonos_();
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
