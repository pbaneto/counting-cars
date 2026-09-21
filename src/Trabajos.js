/** Trabajos: crear y asignar. Un coche puede tener varios trabajos sin pagar a la vez. */

function listaTrabajos_(tabTrab) {
  return tabTrab.filas.map(f => ({
    num: String(f.v['Nº trabajo']).trim(), plate: normPlate(f.v['Matrícula']), pagado: f.v['Pagado'] === true,
    apertura: aISO_(f.v['Fecha apertura']), fila: f.fila,
  }));
}

/** Crea un trabajo nuevo para la matrícula (nº = último dígito + letras + contador) y devuelve su número. */
function crearTrabajo_(tabTrab, matricula, fechaISO) {
  const num = nextJobNumber(jobPrefix(matricula), tabTrab.filas.map(f => f.v['Nº trabajo']));
  agregarFilas_(tabTrab, [{ 'Nº trabajo': num, 'Fecha apertura': aFecha_(fechaISO || hoyISO_()), 'Matrícula': normPlate(matricula), 'Pagado': false }]);
  return num;
}

/** Trabajo sin pagar más reciente de la matrícula; si no hay, crea uno. Devuelve {num, creado}. */
function asignarTrabajo_(tabTrab, matricula, fechaISO) {
  const abierto = pickOpenJob(matricula, listaTrabajos_(tabTrab));
  if (abierto) return { num: abierto.num, creado: false };
  return { num: crearTrabajo_(tabTrab, matricula, fechaISO), creado: true };
}

/** Desplegable de la celda "Nº trabajo" con los trabajos de esa matrícula + NUEVO. */
function ponerDesplegableTrabajo_(tabAlb, fila, tabTrab, matricula) {
  const nums = listaTrabajos_(tabTrab).filter(j => j.plate === normPlate(matricula)).map(j => j.num).concat(['NUEVO']);
  const regla = SpreadsheetApp.newDataValidation().requireValueInList(nums, true).setAllowInvalid(true)
    .setHelpText('Elige el trabajo de este coche, o NUEVO para abrir uno nuevo.').build();
  tabAlb.sh.getRange(fila, tabAlb.map['Nº trabajo']).setDataValidation(regla);
}
