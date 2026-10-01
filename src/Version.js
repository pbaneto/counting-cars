/**
 * Versiones de la hoja. La hoja guarda en qué versión está (propiedad VERSION_HOJA) y el código sabe la suya (VERSION,
 * en Config.js). Al empezar una acción del menú o al abrir la hoja, si no coinciden:
 *  - se ejecutan, en orden y una sola vez, las migraciones de las versiones intermedias (MIGRACIONES);
 *  - si cambia la versión mayor o la menor, se vuelve a aplicar el diseño (aplicarDiseno_);
 *  - se apunta la versión nueva y se deja constancia en Registro.
 * Si coinciden, sólo se lee una propiedad.
 */

const PROP_VERSION_ = 'VERSION_HOJA';

/**
 * Migraciones de datos, por versión: { '2.0.0': () => { ... } }. Deben ser idempotentes (si se cortan a medias, se
 * repiten enteras). Cuando una ya se ha ejecutado en la hoja real, se borra de aquí en el siguiente cambio: la
 * versión guardada en la hoja impide que vuelva a hacer falta.
 */
const MIGRACIONES = {};

/** Pone la hoja al día con el código. forzarDiseno: aplicar el diseño aunque la versión no haya cambiado (menú). */
function actualizarHoja_(forzarDiseno) {
  const props = PropertiesService.getScriptProperties();
  const antes = props.getProperty(PROP_VERSION_) || '1.0.0';  // hojas anteriores al sistema de versiones: 1.0.0
  if (antes === VERSION && !forzarDiseno) return false;
  if (compararVersiones(antes, VERSION) > 0) {
    log_('AVISO', 'actualizarHoja', '', `La hoja está en la versión ${antes}, más nueva que el código (${VERSION}): no se toca`);
    return false;
  }
  Object.keys(MIGRACIONES).filter(v => compararVersiones(v, antes) > 0 && compararVersiones(v, VERSION) <= 0).sort(compararVersiones).forEach(v => {
    MIGRACIONES[v]();
    props.setProperty(PROP_VERSION_, v);
    log_('INFO', 'actualizarHoja', v, `Migración ${v} aplicada`);
  });
  if (forzarDiseno || cambiaDiseno(antes, VERSION)) aplicarDiseno_();
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
