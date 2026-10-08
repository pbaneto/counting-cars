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
  // Albaranes escaneados el 7 y el 8 de octubre de 2026 que quedaron sin fechas, enlace al PDF ni fórmulas. Los datos (nº
  // albarán, fecha de escaneo, fecha del albarán, PDF) se han sacado de los PDF y están en src/private.js. Sólo rellena
  // lo que esté vacío; las fórmulas las vuelve a poner el diseño, que se aplica justo después por cambiar la versión mayor.
  '4.0.0': () => {
    if (typeof RECUPERAR_ALBARANES_4_0_0 === 'undefined') return;
    const alb = leerTabla_(HOJA.ALB), trab = leerTabla_(HOJA.TRAB);
    RECUPERAR_ALBARANES_4_0_0.forEach(([num, escaneo, fecha, idPdf]) => {
      const f = alb.filas.find(x => normAlbaran(x.v['Nº albarán']) === num);
      if (!f) { log_('AVISO', 'migración 4.0.0', num, `No encuentro el albarán ${num} en Albaranes`); return; }
      const cambios = {};
      if (!f.v['Fecha escaneo']) cambios['Fecha escaneo'] = aFecha_(escaneo);
      if (!f.v['Fecha albarán']) cambios['Fecha albarán'] = aFecha_(fecha);
      if (!f.v['Ver PDF']) cambios['Ver PDF'] = enlacePdfId_(idPdf);
      actualizarFila_(alb, f.fila, cambios);
      if (f.v['Nº trabajo'] && f.v['Matrícula']) ponerDesplegableTrabajo_(alb, f.fila, trab, f.v['Matrícula']);
    });
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
