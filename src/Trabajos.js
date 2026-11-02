/** Trabajos: crear y asignar. Un coche puede tener varios trabajos sin pagar a la vez. */

function listaTrabajos_(tabTrab) {
  return tabTrab.filas.map(f => ({
    num: String(f.v['Nº trabajo']).trim(), plate: normPlate(f.v['Matrícula']), pagado: f.v['Pagado'] === true, fila: f.fila,
  }));
}

/** Crea un trabajo nuevo para la matrícula (nº = último dígito + letras + contador) y devuelve su número. */
function crearTrabajo_(tabTrab, matricula, fechaISO) {
  const num = nextJobNumber(jobPrefix(matricula), tabTrab.filas.map(f => f.v['Nº trabajo']));
  const fecha = fechaISO || hoyISO_();
  agregarFilas_(tabTrab, [{ 'Nº trabajo': num, 'Mes': mesDe(fecha), 'Quincena': quincenaDe(fecha), 'Matrícula': normPlate(matricula), 'Pagado': false }]);
  return num;
}

/** Trabajo sin pagar más reciente de la matrícula; si no hay, crea uno. Devuelve {num, creado}. */
function asignarTrabajo_(tabTrab, matricula, fechaISO) {
  const abierto = pickOpenJob(matricula, listaTrabajos_(tabTrab));
  if (abierto) return { num: abierto.num, creado: false };
  return { num: crearTrabajo_(tabTrab, matricula, fechaISO), creado: true };
}

/**
 * Se ha corregido la matrícula de un albarán: decide su trabajo nuevo y recoloca el viejo (ver destinoTrabajoViejo).
 * viejo: su fila en Trabajos, si existe. Devuelve el nº de trabajo del albarán ('' si se ha quedado sin matrícula).
 * Puede borrar filas: después hay que volver a leer Trabajos.
 */
function reasignarAlbaran_(filaAlb, tabTrab, viejo, plate, fechaISO) {
  const abierto = plate ? pickOpenJob(plate, listaTrabajos_(tabTrab)) : null;
  const filaAbierto = abierto && tabTrab.filas.find(f => f.fila === abierto.fila);
  const numViejo = viejo && String(viejo.v['Nº trabajo']).trim();
  const quedan = viejo ? leerTabla_(HOJA.ALB).filas.filter(f => f.fila !== filaAlb && mismoTrabajo_(f.v['Nº trabajo'], numViejo)).length : 0;
  const accion = !viejo ? 'dejar' : destinoTrabajoViejo({
    quedan, conCobro: viejo.v['Factura'] !== '' || viejo.v['Pagado'] === true,
    hayMatricula: !!plate, hayAbierto: !!abierto, abiertoConFactura: !!abierto && filaAbierto.v['Factura'] !== '',
  });

  if (accion === 'renombrar') {
    const num = nextJobNumber(jobPrefix(plate), tabTrab.filas.map(f => f.v['Nº trabajo']));
    actualizarFila_(tabTrab, viejo.fila, { 'Nº trabajo': num, 'Matrícula': plate });
    log_('INFO', 'alEditar', `${HOJA.TRAB}!${viejo.fila}`, `Trabajo ${numViejo} corregido: ahora es ${num} (${plate})`);
    return num;
  }
  const num = !plate ? '' : abierto ? abierto.num : crearTrabajo_(tabTrab, plate, fechaISO);
  if (accion === 'mover') {
    actualizarFila_(tabTrab, abierto.fila, { 'Factura': viejo.v['Factura'], 'Pagado': viejo.v['Pagado'] === true });
    toast_(`La factura y el pagado de ${numViejo} han pasado a ${num}. Compruébalos.`, '🔁 Trabajo corregido', 8);
  }
  if (accion === 'mover' || accion === 'borrar') {
    sinFiltro_(tabTrab.sh);
    tabTrab.sh.deleteRow(viejo.fila);
    log_('INFO', 'alEditar', HOJA.TRAB, `Trabajo ${numViejo} borrado: se ha quedado sin albaranes` + (accion === 'mover' ? ` (factura y pagado pasan a ${num})` : ''));
  } else if (viejo && quedan) {
    toast_(`${numViejo} sigue con ${quedan} albarán(es). Si también son de otra matrícula, corrígelos.`, '🔁 Trabajo corregido', 8);
  } else if (viejo) {
    toast_(`${numViejo} se queda sin albaranes y conserva su factura o pagado: revísalo.`, '⚠ Trabajo sin albaranes', 10);
  }
  return num;
}

/**
 * Se ha corregido la matrícula en Trabajos: el trabajo entero cambia de coche. Sus albaranes pasan a esa matrícula
 * y al nº de trabajo nuevo. Devuelve las matrículas que tenían antes (vacío si no tiene albaranes).
 */
function cambiarCocheDeAlbaranes_(numViejo, numNuevo, tabTrab, plate) {
  const tabAlb = leerTabla_(HOJA.ALB), antes = new Set();
  tabAlb.filas.filter(f => mismoTrabajo_(f.v['Nº trabajo'], numViejo)).forEach(f => {
    antes.add(normPlate(f.v['Matrícula']));
    actualizarFila_(tabAlb, f.fila, { 'Nº trabajo': numNuevo, 'Matrícula': plate });
    ponerDesplegableTrabajo_(tabAlb, f.fila, tabTrab, plate);
  });
  return Array.from(antes);
}

function mismoTrabajo_(a, b) { return String(a).trim().toUpperCase() === String(b).trim().toUpperCase(); }

/** Desplegable de la celda "Nº trabajo" con los trabajos de esa matrícula + NUEVO. */
function ponerDesplegableTrabajo_(tabAlb, fila, tabTrab, matricula) {
  const nums = listaTrabajos_(tabTrab).filter(j => j.plate === normPlate(matricula)).map(j => j.num).concat(['NUEVO']);
  const regla = SpreadsheetApp.newDataValidation().requireValueInList(nums, true).setAllowInvalid(true)
    .setHelpText('Elige el trabajo de este coche, o NUEVO para abrir uno nuevo.').build();
  tabAlb.sh.getRange(fila, tabAlb.map['Nº trabajo']).setDataValidation(regla);
}

/**
 * Matrícula que no está en Coches: se añade una fila sólo con la matrícula y el cliente y el coche se rellenan a mano.
 * tabCoches: la tabla de Coches ya leída (se mantiene al día). Devuelve true si la ha añadido; false si ya estaba o si es
 * demasiado corta para ser una matrícula.
 */
function anadirCoche_(tabCoches, matricula) {
  const plate = normPlate(matricula);
  if (plate.length < MATRICULA_LARGO_MIN_COCHE || tabCoches.filas.some(f => normPlate(f.v['Matrícula']) === plate)) return false;
  agregarFilas_(tabCoches, [{ 'Matrícula': plate }]);
  log_('INFO', 'anadirCoche', HOJA.COCHES, `Matrícula ${plate} añadida a Coches: falta rellenar el cliente y el coche`);
  return true;
}

/** Quita de Coches una matrícula errónea: sólo si no tiene cliente ni coche y ningún albarán ni trabajo la usa. */
function quitarCocheSinUso_(matricula) {
  const plate = normPlate(matricula), tab = leerTabla_(HOJA.COCHES);
  const fila = tab.filas.find(f => normPlate(f.v['Matrícula']) === plate);
  if (!fila || fila.v['Cliente'] !== '' || fila.v['Coche'] !== '') return;
  if ([HOJA.ALB, HOJA.TRAB].some(h => leerTabla_(h).filas.some(f => normPlate(f.v['Matrícula']) === plate))) return;
  sinFiltro_(tab.sh);
  tab.sh.deleteRow(fila.fila);
  log_('INFO', 'alEditar', HOJA.COCHES, `Matrícula ${plate} quitada de Coches: estaba sin datos y ya nadie la usa`);
}
