/** Menú y edición manual. `alEditar` es un trigger INSTALABLE (lo crea setup) porque necesita Drive y bloqueos. */

function alAbrir() {
  SpreadsheetApp.getUi().createMenu('Counting Cars')
    .addItem('Procesar albaranes (carpeta Entrada)', 'procesarAlbaranes')
    .addItem('Procesar facturas RM', 'procesarFacturasRM')
    .addItem('Actualizar Abonos', 'actualizarAbonos')
    .addSeparator()
    .addItem('Diagnóstico (buscar problemas)', 'diagnostico')
    .addItem('Reparar fórmulas y formato', 'repararFormulas')
    .addSeparator()
    .addItem('Configurar API key de Gemini', 'configurarApiKey')
    .addItem('Preparar hoja (primera vez)', 'setup')
    .addItem('Cargar coches y datos del piloto', 'cargarDatosIniciales')
    .addToUi();
}

function alEditar(e) {
  if (!e || !e.range) return;
  const sh = e.range.getSheet(), nombre = sh.getName();
  if (nombre === HOJA.ABONOS) { editarAbonos_(e.range); return; }
  if ([HOJA.ALB, HOJA.TRAB, HOJA.PIEZAS, HOJA.COCHES].indexOf(nombre) < 0) return;
  ejecutar_('alEditar', () => {
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(20000)) { toast_('Counting Cars está ocupado con otro proceso. Repite la edición en un momento.', '⚠'); return; }
    try {
      const r0 = e.range.getRow(), n = Math.min(e.range.getNumRows(), 300), c0 = e.range.getColumn(), nc = e.range.getNumColumns();
      if (nombre === HOJA.ALB) editarAlbaranes_(r0, n, c0, nc);
      else if (nombre === HOJA.TRAB) editarTrabajos_(r0, n, c0, nc);
      else if (nombre === HOJA.PIEZAS) editarPiezas_(r0, n, c0, nc);
      else editarCoches_(r0, n);
    } finally { lock.releaseLock(); }
  }, true);
}

function tocada_(tab, h, c0, nc) { const c = tab.map[h]; return c >= c0 && c < c0 + nc; }

function editarAlbaranes_(r0, n, c0, nc) {
  const tab = leerTabla_(HOJA.ALB), hoy = hoyISO_(), ancho = anchoTabla_(tab);
  let trab = null, coches = null;
  for (let r = Math.max(r0, 2); r < r0 + n; r++) {
    const vals = tab.sh.getRange(r, 1, 1, ancho).getValues()[0];
    const g = h => vals[tab.map[h] - 1];
    const set = (h, v) => { tab.sh.getRange(r, tab.map[h]).setValue(v); vals[tab.map[h] - 1] = v; };

    const plate = normPlate(g('Matrícula'));
    if (String(g('Matrícula')) !== plate) set('Matrícula', plate);
    const num = normAlbaran(g('Nº albarán'));
    if (String(g('Nº albarán')) !== num) { tab.sh.getRange(r, tab.map['Nº albarán']).setNumberFormat('@'); set('Nº albarán', num); }
    const precio = parseNumber(g('Precio con IVA'));
    const jobVal = String(g('Nº trabajo')).trim().toUpperCase();

    if (jobVal === 'NUEVO' && !plate) { set('Nº trabajo', ''); toast_('Escribe primero la matrícula para abrir un trabajo nuevo.', '⚠ Falta matrícula'); continue; }
    if (!(plate && precio > 0)) continue;

    if (!g('Fecha escaneo')) set('Fecha escaneo', aFecha_(hoy));
    if (!g('Fecha albarán')) set('Fecha albarán', g('Fecha escaneo') || aFecha_(hoy));
    if (!g('Proveedor')) set('Proveedor', 'RM');
    asegurarFormulasFila_(tab, r);

    trab = trab || leerTabla_(HOJA.TRAB);
    const actual = trab.filas.find(f => String(f.v['Nº trabajo']).trim().toUpperCase() === jobVal);
    const cambioPlaca = tocada_(tab, 'Matrícula', c0, nc) && actual && normPlate(actual.v['Matrícula']) !== plate;
    if (jobVal === '' || jobVal === 'NUEVO' || cambioPlaca) {
      const fecha = aISO_(g('Fecha albarán')) || hoy;
      const num2 = jobVal === 'NUEVO' ? crearTrabajo_(trab, plate, fecha) : asignarTrabajo_(trab, plate, fecha).num;
      set('Nº trabajo', num2);
      ponerDesplegableTrabajo_(tab, r, trab, plate);
      log_('INFO', 'alEditar', `${HOJA.ALB}!${r}`, `Trabajo ${num2} asignado a ${plate}`);
    }
    coches = coches || new Set(leerTabla_(HOJA.COCHES).filas.map(f => normPlate(f.v['Matrícula'])));
    if (!coches.has(plate)) toast_(`La matrícula ${plate} no está en la pestaña Coches.`, '⚠ Matrícula desconocida', 8);
  }
}

function editarTrabajos_(r0, n, c0, nc) {
  const tab = leerTabla_(HOJA.TRAB), hoy = hoyISO_(), ancho = anchoTabla_(tab);
  let coches = null;
  for (let r = Math.max(r0, 2); r < r0 + n; r++) {
    const vals = tab.sh.getRange(r, 1, 1, ancho).getValues()[0];
    const g = h => vals[tab.map[h] - 1];
    const set = (h, v) => { tab.sh.getRange(r, tab.map[h]).setValue(v); vals[tab.map[h] - 1] = v; };
    const plate = normPlate(g('Matrícula'));
    if (String(g('Matrícula')) !== plate) set('Matrícula', plate);
    if (!plate) continue;
    asegurarFormulasFila_(tab, r);
    if (String(g('Nº trabajo')).trim() === '') {
      const otros = tab.filas.filter(f => f.fila !== r).map(f => f.v['Nº trabajo']);
      set('Nº trabajo', nextJobNumber(jobPrefix(plate), otros));
      if (!g('Fecha apertura')) set('Fecha apertura', aFecha_(hoy));
      log_('INFO', 'alEditar', `${HOJA.TRAB}!${r}`, `Trabajo ${g('Nº trabajo')} creado a mano para ${plate}`);
    }
    coches = coches || new Set(leerTabla_(HOJA.COCHES).filas.map(f => normPlate(f.v['Matrícula'])));
    if (!coches.has(plate)) toast_(`La matrícula ${plate} no está en la pestaña Coches.`, '⚠ Matrícula desconocida', 8);
  }
}

function editarPiezas_(r0, n, c0, nc) {
  const tab = leerTabla_(HOJA.PIEZAS), hoy = hoyISO_(), ancho = anchoTabla_(tab);
  let toca = false;
  for (let r = Math.max(r0, 2); r < r0 + n; r++) {
    const vals = tab.sh.getRange(r, 1, 1, ancho).getValues()[0];
    const g = h => vals[tab.map[h] - 1];
    const set = (h, v) => { tab.sh.getRange(r, tab.map[h]).setValue(v); vals[tab.map[h] - 1] = v; };
    const num = normAlbaran(g('Nº albarán'));
    if (String(g('Nº albarán')) !== num) { tab.sh.getRange(r, tab.map['Nº albarán']).setNumberFormat('@'); set('Nº albarán', num); }

    if (tocada_(tab, 'Reembolso', c0, nc)) {
      if (g('Reembolso') === true) {
        if (!num) { set('Reembolso', false); toast_('Para pedir un reembolso hay que poner antes el nº de albarán de la pieza.', '⚠ Falta el nº de albarán', 10); continue; }
        if (!g('Fecha reembolso')) set('Fecha reembolso', aFecha_(hoy));
      } else set('Fecha reembolso', '');
      toca = true;
    } else if (g('Reembolso') === true) toca = true;

    if (!g('Origen') && (num || g('Descripción') || g('Precio descontado sin IVA') !== '')) set('Origen', 'Manual');
    asegurarFormulasFila_(tab, r);
  }
  if (toca) reconstruirAbonos_();
}

function editarCoches_(r0, n) {
  const tab = leerTabla_(HOJA.COCHES);
  for (let r = Math.max(r0, 2); r < r0 + n; r++) {
    const c = tab.sh.getRange(r, tab.map['Matrícula']), v = c.getValue();
    if (v === '') continue;
    const p = normPlate(v);
    if (String(v) !== p) c.setValue(p);
    if (tab.filas.filter(f => normPlate(f.v['Matrícula']) === p).length > 1) toast_(`La matrícula ${p} está repetida en Coches.`, '⚠ Duplicada', 8);
  }
}

/** Casilla "Reescanear factura" del resumen de Abonos. */
function editarAbonos_(rango) {
  const c = ABONOS.cabResumen.indexOf('Reescanear factura') + 1;
  if (rango.getColumn() > c || rango.getColumn() + rango.getNumColumns() - 1 < c) return;
  const sh = rango.getSheet();
  for (let r = Math.max(rango.getRow(), ABONOS.filaIni); r < rango.getRow() + rango.getNumRows() && r < ABONOS.filaIni + ABONOS.filas; r++) {
    const celda = sh.getRange(r, c);
    if (celda.getValue() !== true) continue;
    celda.setValue(false);
    const k = r - ABONOS.filaIni, anio = Number(sh.getRange(ABONOS.celdaAnio).getValue());
    reescanearPeriodo_(anio, Math.floor(k / 2) + 1, (k % 2) + 1);
  }
}
