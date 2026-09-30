/** Menú y edición manual. `alEditar` es un trigger INSTALABLE (lo crea setup) porque necesita Drive y bloqueos. */

function alAbrir() {
  SpreadsheetApp.getUi().createMenu('Counting Cars')
    .addItem('Procesar albaranes (carpeta Entrada)', 'procesarAlbaranes')
    .addItem('Procesar facturas RM', 'procesarFacturasRM')
    .addItem('Actualizar Abonos', 'actualizarAbonos')
    .addItem('Recuperar estados de Abonos (historial)', 'recuperarEstadosAbonos')
    .addSeparator()
    .addItem('Diagnóstico (buscar problemas)', 'diagnostico')
    .addItem('Reparar fórmulas y formato', 'repararFormulas')
    .addSeparator()
    .addItem('Configurar API key de Gemini', 'configurarApiKey')
    .addItem('Preparar hoja (primera vez)', 'setup')
    .addToUi();
}

function alEditar(e) {
  if (!e || !e.range) return;
  const sh = e.range.getSheet(), nombre = sh.getName();
  if ([HOJA.ALB, HOJA.TRAB, HOJA.PIEZAS, HOJA.COCHES, HOJA.ABONOS].indexOf(nombre) < 0) return;
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
      else if (nombre === HOJA.ABONOS) editarAbonos_(r0, n);
      else editarCoches_(r0, n);
      crono.paso(`editar ${nombre}`);
      crono.fin();
    } finally { lock.releaseLock(); }
  }, true);
}

/**
 * Trigger INSTALABLE onChange: borrar filas no dispara onEdit. Si se borran filas en Abonos, las piezas cuya fila ya
 * no está se desmarcan en Piezas. Los cambios hechos por el propio script no disparan este trigger.
 */
function alCambiar(e) {
  if (!e || e.changeType !== 'REMOVE_ROW') return;
  const activa = e.source && e.source.getActiveSheet();
  if (!activa || activa.getName() !== HOJA.ABONOS) return;
  ejecutar_('alCambiar', () => {
    if (e.source) _ss = e.source;
    conBloqueo_(20, desmarcarPiezasBorradas_);
  }, true);
}

function tocada_(tab, h, c0, nc) { const c = tab.map[h]; return c >= c0 && c < c0 + nc; }

function editarAlbaranes_(r0, n, c0, nc) {
  const tab = leerTabla_(HOJA.ALB), hoy = hoyISO_(), ancho = anchoTabla_(tab);
  let trab = null, cochesTab = null;
  // Se carga una sola vez por lote de filas editadas, y sólo si hace falta (alguna fila con matrícula).
  const coches_ = () => cochesTab || (cochesTab = leerTabla_(HOJA.COCHES).filas.map(
    f => ({ plate: normPlate(f.v['Matrícula']), cliente: f.v['Cliente'], coche: f.v['Coche'] })));
  for (let r = Math.max(r0, 2); r < r0 + n; r++) {
    const vals = tab.sh.getRange(r, 1, 1, ancho).getValues()[0];
    const g = h => vals[tab.map[h] - 1];
    const set = (h, v) => { tab.sh.getRange(r, tab.map[h]).setValue(v); vals[tab.map[h] - 1] = v; };

    const escrita = normPlate(g('Matrícula'));
    if (String(g('Matrícula')) !== escrita) set('Matrícula', escrita);
    let plate = escrita;
    if (plate) {
      const res = resolverMatricula(plate, coches_());
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
    if (filaConDatos_(vals)) asegurarFila_(tab, r);

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
    if (!coches_().some(c => c.plate === plate)) toast_(`La matrícula ${plate} no está en la pestaña Coches.`, '⚠ Matrícula desconocida', 8);
  }
}

function editarTrabajos_(r0, n, c0, nc) {
  const tab = leerTabla_(HOJA.TRAB), hoy = hoyISO_(), ancho = anchoTabla_(tab);
  let coches = null;
  // El panel "Resumen (según filtro)" vive por encima de la cabecera real: una edición ahí no es una fila de datos.
  for (let r = Math.max(r0, tab.esq.filaCabecera + 1); r < r0 + n; r++) {
    const vals = tab.sh.getRange(r, 1, 1, ancho).getValues()[0];
    const g = h => vals[tab.map[h] - 1];
    const set = (h, v) => { tab.sh.getRange(r, tab.map[h]).setValue(v); vals[tab.map[h] - 1] = v; };
    const plate = normPlate(g('Matrícula'));
    if (String(g('Matrícula')) !== plate) set('Matrícula', plate);
    if (filaConDatos_(vals)) asegurarFila_(tab, r);
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

/**
 * Primero TODAS las lecturas, luego todas las escrituras y al final se guarda de una vez: en Sheets, una lectura
 * hecha después de escribir espera a que se recalculen las fórmulas que dependen de Piezas (Albaranes, Trabajos,
 * Resumen, Abonos), y eso costaba más de un segundo por edición.
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
  const iReembolso = tab.map['Reembolso'] - 1;
  const puedeTocar = tocada_(tab, 'Reembolso', c0, nc) || editadas.some(e => e.vals[iReembolso] === true);
  let paraAbonos = puedeTocar ? leerParaAbonos_(tab, crono) : null;  // lee con la tabla de Piezas ya leída; se mantiene al día abajo

  let toca = false;
  const desmarcadas = [];
  editadas.forEach(({ r, vals, formulas: formulasFila }) => {
    const fila = tab.filas.find(f => f.fila === r);
    const g = h => vals[tab.map[h] - 1];
    const set = (h, v) => { tab.sh.getRange(r, tab.map[h]).setValue(v); vals[tab.map[h] - 1] = v; if (fila) fila.v[h] = v; };
    if (filaConDatos_(vals)) asegurarFila_(tab, r, formulasFila);
    const num = normAlbaran(g('Nº albarán'));
    if (String(g('Nº albarán')) !== num) { tab.sh.getRange(r, tab.map['Nº albarán']).setNumberFormat('@'); set('Nº albarán', num); }

    if (tocada_(tab, 'Reembolso', c0, nc)) {
      if (g('Reembolso') === true) {
        if (!num) { set('Reembolso', false); toast_('Para pedir un reembolso hay que poner antes el nº de albarán de la pieza.', '⚠ Falta el nº de albarán', 10); return; }
        if (!g('Fecha reembolso')) set('Fecha reembolso', aFecha_(hoy));
      } else { set('Fecha reembolso', ''); desmarcadas.push(r); }
      toca = true;
    } else if (g('Reembolso') === true) toca = true;

    if (!g('Origen') && (num || g('Descripción') || g('Precio descontado sin IVA') !== '')) set('Origen', 'Manual');
  });
  crono.paso(`revisar ${editadas.length} fila(s) editada(s)`);
  if (toca) {
    paraAbonos = paraAbonos || leerParaAbonos_(tab, crono);
    escribirAbonos_(paraAbonos, desmarcadas);
    crono.paso('sincronizar Abonos');
  }
  SpreadsheetApp.flush();
  crono.paso('guardar cambios y recalcular');
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
