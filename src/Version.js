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
const MIGRACIONES = {
  // Abonos: en la hoja real se movió "Matrícula" a la columna C y luego el diseño volvió a escribir encima las cabeceras
  // en el orden por defecto, y las filas nuevas se escribieron en ese orden. Se vuelve al orden de los datos movidos
  // (cabecera incluida) y se pasan a él las filas escritas en el orden por defecto. Desde 3.0.0 las columnas de Abonos
  // se buscan por su nombre, así que no puede volver a pasar.
  '3.0.0': () => {
    const sh = hoja_(HOJA.ABONOS), ct = ABONOS.cabTabla, n = ct.length, ult = sh.getLastRow();
    if (ult < 2) return;
    const vals = sh.getRange(1, 1, ult, n).getValues();
    const est = x => ESTADOS_ABONO.indexOf(String(x)) >= 0, vacia = r => r.every(x => x === '' || x == null);
    if (vals[0].join('|') !== ct.join('|') || !vals.slice(1).some(r => est(r[5]) && !est(r[4]))) return;
    // ct: 0 Fecha abono, 1 Descripción, 2 Precio sin IVA, 3 Precio con IVA, 4 Estado, 5 Nº albarán, 6 Referencia, 7 Matrícula, 8…12 igual
    const aMovida = r => [r[0], r[1], r[7], r[2], r[3], r[4], r[5], r[6]].concat(r.slice(8));
    const movida = r => est(r[5]) || (!est(r[4]) && typeof r[2] !== 'number');
    const filas = vals.slice(1).map((r, i) => {
      if (vacia(r)) return r;
      const f = movida(r) ? r.slice() : aMovida(r), fila = i + 2;
      f[4] = loc_(`=ROUND($D${fila}*(1+IVA),2)`);
      f[11] = loc_(`=IF($F${fila}="Sin abonar",TODAY()-$I${fila},"")`);
      return f;
    });
    sh.getRange(1, 1, 1, n).setValues([aMovida(ct)]);
    sh.getRange(2, 1, filas.length, n).setValues(filas);
    _colsAbonos = null;
  },
};

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
