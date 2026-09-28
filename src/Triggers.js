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
  if ([HOJA.ALB, HOJA.TRAB, HOJA.PIEZAS, HOJA.COCHES].indexOf(nombre) < 0) return;
  ejecutar_('alEditar', () => {
    if (e.source) _ss = e.source;  // la hoja ya viene abierta en el evento: evita abrirla otra vez por ID
    const crono = cronometro_(`alEditar (${nombre})`);
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(20000)) { toast_('Counting Cars está ocupado con otro proceso. Repite la edición en un momento.', '⚠'); return; }
    crono.paso('esperar bloqueo');
    try {
      const r0 = e.range.getRow(), n = Math.min(e.range.getNumRows(), 300), c0 = e.range.getColumn(), nc = e.range.getNumColumns();
      if (nombre === HOJA.ALB) editarAlbaranes_(r0, n, c0, nc);
      else if (nombre === HOJA.TRAB) editarTrabajos_(r0, n, c0, nc);
      else if (nombre === HOJA.PIEZAS) editarPiezas_(r0, n, c0, nc);
      else editarCoches_(r0, n);
      crono.paso(`editar ${nombre}`);
      crono.fin();
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
    if (filaConDatos_(vals)) asegurarFormulasFila_(tab, r);

    if (jobVal === 'NUEVO' && !plate) { set('Nº trabajo', ''); toast_('Escribe primero la matrícula para abrir un trabajo nuevo.', '⚠ Falta matrícula'); continue; }
    if (!(plate && precio > 0)) continue;

    if (!g('Fecha escaneo')) set('Fecha escaneo', aFecha_(hoy));
    if (!g('Fecha albarán')) set('Fecha albarán', g('Fecha escaneo') || aFecha_(hoy));
    if (!g('Proveedor')) set('Proveedor', 'RM');

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
    if (filaConDatos_(vals)) asegurarFormulasFila_(tab, r);
    if (!plate) continue;
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
  const crono = cronometro_('editarPiezas_');
  const tab = leerTabla_(HOJA.PIEZAS), hoy = hoyISO_(), ancho = anchoTabla_(tab);
  crono.paso(`leer Piezas (${tab.filas.length} de ${tab.leidas} filas)`);
  let toca = false;
  for (let r = Math.max(r0, 2); r < r0 + n; r++) {
    const vals = tab.sh.getRange(r, 1, 1, ancho).getValues()[0];
    const fila = tab.filas.find(f => f.fila === r);  // se mantiene al día para pasar la tabla a reconstruirAbonos_ sin releerla
    const g = h => vals[tab.map[h] - 1];
    const set = (h, v) => { tab.sh.getRange(r, tab.map[h]).setValue(v); vals[tab.map[h] - 1] = v; if (fila) fila.v[h] = v; };
    if (filaConDatos_(vals)) asegurarFormulasFila_(tab, r);
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
  }
  crono.paso(`revisar ${n} fila(s) editada(s)`);
  if (toca) { reconstruirAbonos_(tab); crono.paso('reconstruir Abonos'); }
  crono.fin(toca ? '' : 'sin cambios en reembolsos: Abonos no se toca');
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
