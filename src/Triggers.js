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
    .addSeparator()
    .addItem(`Versión ${VERSION}`, 'verVersion')
    .addToUi();
  // Si se ha subido una versión nueva del código, la hoja se pone al día al abrirla (si nadie más está trabajando).
  ejecutar_('alAbrir', () => {
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(1000)) return;
    try { if (actualizarHoja_(false)) toast_(`Hoja actualizada a la versión ${VERSION}.`); } finally { lock.releaseLock(); }
  }, true);
}

function alEditar(e) {
  if (!e || !e.range) return;
  const sh = e.range.getSheet(), nombre = sh.getName();
  if ([HOJA.ALB, HOJA.TRAB, HOJA.PIEZAS, HOJA.COCHES, HOJA.ABONOS, HOJA.LINEAS].indexOf(nombre) < 0) return;
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
      else if (nombre === HOJA.ABONOS) editarAbonos_(r0, n, c0);
      else if (nombre === HOJA.LINEAS) editarLineas_(r0, n, c0, nc);
      else editarCoches_(r0, n);
      crono.paso(`editar ${nombre}`);
      crono.fin();
    } finally { lock.releaseLock(); }
  }, true);
}

/**
 * Trigger INSTALABLE onChange: insertar o borrar filas no dispara onEdit. En Abonos, el resumen de la derecha (que
 * comparte filas con la tabla) se vuelve a montar si se ha movido y la tabla se rehace. Borrar filas de Piezas o de
 * Líneas RM también cambia Abonos. Los cambios hechos por el propio script no disparan este trigger.
 */
function alCambiar(e) {
  if (!e || ['REMOVE_ROW', 'INSERT_ROW'].indexOf(e.changeType) < 0) return;
  const activa = e.source && e.source.getActiveSheet(), nombre = activa && activa.getName();
  if (nombre === HOJA.ABONOS) {
    ejecutar_('alCambiar', () => {
      if (e.source) _ss = e.source;
      conBloqueo_(20, () => revisarAbonosTrasCambioDeFilas_());
    }, true);
  } else if (ESQUEMA[nombre]) {
    // Filas insertadas o borradas encima de la cabecera: leer la tabla la devuelve a su sitio (ver filaCabecera_).
    ejecutar_('alCambiar', () => {
      if (e.source) _ss = e.source;
      conBloqueo_(20, () => {
        const tab = leerTabla_(nombre);
        if (e.changeType === 'REMOVE_ROW' && (nombre === HOJA.PIEZAS || nombre === HOJA.LINEAS)) actualizarInformeAbonos_(nombre === HOJA.PIEZAS ? tab : null);
      });
    }, true);
  }
}

function tocada_(tab, h, c0, nc) { const c = tab.map[h]; return c >= c0 && c < c0 + nc; }

/**
 * Columnas automáticas (grises): si alguien escribe encima, asegurarFila_ vuelve a poner la fórmula; aquí sólo se avisa.
 * (Antes eran protecciones con aviso, pero Sheets lo sacaba también al cambiar anchos, ordenar, etc.)
 */
function avisarColumnaAutomatica_(tab, puestas, c0, nc) {
  const pisadas = puestas.filter(h => tocada_(tab, h, c0, nc));
  if (pisadas.length) toast_(`"${pisadas.join('", "')}" se rellena sola: se ha vuelto a poner su fórmula.`, 'Columna automática', 6);
}

function editarAlbaranes_(r0, n, c0, nc) {
  const tab = leerTabla_(HOJA.ALB), hoy = hoyISO_(), ancho = anchoTabla_(tab);
  // La matrícula y el proveedor de un albarán salen en Abonos (las piezas de "Otros" no se reclaman a RM).
  const tocaAbonos = ['Nº albarán', 'Matrícula', 'Proveedor'].some(h => tocada_(tab, h, c0, nc));
  let trab = null, cochesTab = null;
  // Se carga una sola vez por lote de filas editadas, y sólo si hace falta (alguna fila con matrícula).
  const tabCoches_ = () => cochesTab || (cochesTab = leerTabla_(HOJA.COCHES));
  const coches_ = () => tabCoches_().filas.map(f => ({ plate: normPlate(f.v['Matrícula']), cliente: f.v['Cliente'], coche: f.v['Coche'] }));
  for (let r = Math.max(r0, 2); r < r0 + n; r++) {
    const vals = tab.sh.getRange(r, 1, 1, ancho).getValues()[0];
    const g = h => vals[tab.map[h] - 1];
    const set = (h, v) => { sinFiltro_(tab.sh); tab.sh.getRange(r, tab.map[h]).setValue(v); vals[tab.map[h] - 1] = v; };

    const escrita = normPlate(g('Matrícula'));
    if (String(g('Matrícula')) !== escrita) set('Matrícula', escrita);
    let plate = escrita, ambigua = false;
    if (plate) {
      const res = resolverMatricula(plate, coches_());
      ambigua = res.tipo === 'varias';
      if (res.tipo === 'unica') {
        plate = res.candidatos[0].plate;
        set('Matrícula', plate);
        const detalle = [res.candidatos[0].coche, res.candidatos[0].cliente].filter(Boolean).join(' · ');
        toast_(`"${escrita}" → ${plate}${detalle ? ' (' + detalle + ')' : ''}`, '✓ Matrícula completada', 4);
      } else if (res.tipo === 'varias') {
        const lista = res.candidatos.slice(0, 8).map(c => c.plate).join(', ');
        toast_(`"${escrita}" coincide con ${res.candidatos.length} matrículas: ${lista}${res.candidatos.length > 8 ? '…' : ''}. Escribe más letras o dígitos para acotar.`,
          '🔎 Varias coincidencias', 8);
      }
    }
    const num = normAlbaran(g('Nº albarán'));
    if (String(g('Nº albarán')) !== num) { tab.sh.getRange(r, tab.map['Nº albarán']).setNumberFormat('@'); set('Nº albarán', num); }
    const precio = parseNumber(g('Precio con IVA'));
    const jobVal = String(g('Nº trabajo')).trim().toUpperCase();
    if (filaConDatos_(vals)) avisarColumnaAutomatica_(tab, asegurarFila_(tab, r), c0, nc);

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
    if (!ambigua && anadirCoche_(tabCoches_(), plate)) toast_(`${plate} añadida a la pestaña Coches: rellena allí el cliente y el coche.`, '🚗 Coche nuevo', 8);
    else if (!tabCoches_().filas.some(f => normPlate(f.v['Matrícula']) === plate)) toast_(`La matrícula ${plate} no está en la pestaña Coches.`, '⚠ Matrícula desconocida', 8);
  }
  if (tocaAbonos) { SpreadsheetApp.flush(); actualizarInformeAbonos_(); }
}

function editarTrabajos_(r0, n, c0, nc) {
  const tab = leerTabla_(HOJA.TRAB), hoy = hoyISO_(), ancho = anchoTabla_(tab);
  let coches = null;
  // El panel "Resumen (según filtro)" vive por encima de la cabecera real: una edición ahí no es una fila de datos.
  for (let r = Math.max(r0, tab.fc + 1); r < r0 + n; r++) {
    const vals = tab.sh.getRange(r, 1, 1, ancho).getValues()[0];
    const g = h => vals[tab.map[h] - 1];
    const set = (h, v) => { sinFiltro_(tab.sh); tab.sh.getRange(r, tab.map[h]).setValue(v); vals[tab.map[h] - 1] = v; };
    const plate = normPlate(g('Matrícula'));
    if (String(g('Matrícula')) !== plate) set('Matrícula', plate);
    if (filaConDatos_(vals)) avisarColumnaAutomatica_(tab, asegurarFila_(tab, r), c0, nc);
    if (!plate) continue;
    if (String(g('Nº trabajo')).trim() === '') {
      const otros = tab.filas.filter(f => f.fila !== r).map(f => f.v['Nº trabajo']);
      set('Nº trabajo', nextJobNumber(jobPrefix(plate), otros));
      log_('INFO', 'alEditar', `${HOJA.TRAB}!${r}`, `Trabajo ${g('Nº trabajo')} creado a mano para ${plate}`);
    }
    // Mes y Quincena: se rellenan solos con la fecha de hoy sólo si están vacíos (después se pueden cambiar a mano).
    if (g('Mes') === '') set('Mes', mesDe(hoy));
    if (g('Quincena') === '') set('Quincena', quincenaDe(hoy));
    coches = coches || leerTabla_(HOJA.COCHES);
    if (anadirCoche_(coches, plate)) toast_(`${plate} añadida a la pestaña Coches: rellena allí el cliente y el coche.`, '🚗 Coche nuevo', 8);
    else if (!coches.filas.some(f => normPlate(f.v['Matrícula']) === plate)) toast_(`La matrícula ${plate} no está en la pestaña Coches.`, '⚠ Matrícula desconocida', 8);
  }
}

/**
 * Primero TODAS las lecturas, luego todas las escrituras y al final se guarda de una vez: en Sheets, una lectura
 * hecha después de escribir espera a que se recalculen las fórmulas que dependen de Piezas (Albaranes, Trabajos,
 * Resumen, Abonos), y eso costaba más de un segundo por edición.
 * Abonos se rehace siempre: cualquier dato de una pieza (precio, referencia, albarán…) puede salir allí. Si no cambia
 * nada, no se escribe.
 */
function editarPiezas_(r0, n, c0, nc) {
  const crono = cronometro_('editarPiezas_');
  const tab = leerTabla_(HOJA.PIEZAS), hoy = hoyISO_(), ancho = anchoTabla_(tab);
  crono.paso(`leer Piezas (${tab.filas.length} de ${tab.leidas} filas)`);
  const desde = Math.max(r0, 2), hasta = r0 + n;
  if (hasta <= desde) { crono.fin('sólo la cabecera'); return; }
  const formulas = tab.sh.getRange(desde, 1, hasta - desde, ancho).getFormulas();
  const editadas = [];
  for (let r = desde; r < hasta; r++) {
    const leida = tab.valores[r - 1] || [];
    editadas.push({ r, formulas: formulas[r - desde], vals: Array.from({ length: ancho }, (_, i) => (leida[i] === undefined ? '' : leida[i])) });
  }
  crono.paso(`leer fórmulas de ${editadas.length} fila(s)`);
  const paraAbonos = leerParaAbonos_(tab, crono);  // con la tabla de Piezas ya leída; se mantiene al día abajo

  editadas.forEach(({ r, vals, formulas: formulasFila }) => {
    const fila = tab.filas.find(f => f.fila === r);
    const g = h => vals[tab.map[h] - 1];
    const set = (h, v) => { sinFiltro_(tab.sh); tab.sh.getRange(r, tab.map[h]).setValue(v); vals[tab.map[h] - 1] = v; if (fila) fila.v[h] = v; };
    if (filaConDatos_(vals)) avisarColumnaAutomatica_(tab, asegurarFila_(tab, r, formulasFila), c0, nc);
    const num = normAlbaran(g('Nº albarán'));
    if (String(g('Nº albarán')) !== num) { tab.sh.getRange(r, tab.map['Nº albarán']).setNumberFormat('@'); set('Nº albarán', num); }

    if (tocada_(tab, 'Reembolso', c0, nc)) {
      if (g('Reembolso') === true) {
        if (!num) { set('Reembolso', false); toast_('Para pedir un reembolso hay que poner antes el nº de albarán de la pieza.', '⚠ Falta el nº de albarán', 10); return; }
        if (!g('Fecha reembolso')) set('Fecha reembolso', aFecha_(hoy));
      } else set('Fecha reembolso', '');
    }
    if (!g('Origen') && (num || g('Descripción') || g('Precio descontado sin IVA') !== '')) set('Origen', 'Manual');
  });
  crono.paso(`revisar ${editadas.length} fila(s) editada(s)`);
  escribirAbonos_(paraAbonos);
  crono.paso('Abonos');
  SpreadsheetApp.flush();
  crono.paso('guardar cambios y recalcular');
  crono.fin();
}

/**
 * Edición a mano en Líneas RM (normalmente, corregir a qué pieza corresponde un abono en "ID pieza"): se comprueba el
 * ID escrito y se rehace Abonos.
 */
function editarLineas_(r0, n, c0, nc) {
  const tab = leerTabla_(HOJA.LINEAS);
  if (r0 + n <= 2) return;
  if (tocada_(tab, 'ID pieza', c0, nc)) {
    const ids = new Set(leerTabla_(HOJA.PIEZAS).filas.map(f => String(f.v['ID'])));
    for (let r = Math.max(r0, 2); r < r0 + n; r++) {
      const c = tab.sh.getRange(r, tab.map['ID pieza']), v = c.getValue(), id = idPiezaDe(v);
      if (id === SIN_PIEZA && v !== SIN_PIEZA) c.setValue(SIN_PIEZA);
      else if (id && id !== SIN_PIEZA && !ids.has(id)) toast_(`No hay ninguna pieza con ID ${id} en Piezas. Escribe el ID de la pieza, o ${SIN_PIEZA} si el abono no es de ninguna.`, '⚠ ID pieza', 10);
    }
  }
  actualizarInformeAbonos_();
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
