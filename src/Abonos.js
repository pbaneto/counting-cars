/**
 * Pestaña Abonos.
 *  - Resumen por quincena (filas 4-27, fórmulas creadas en setup).
 *  - Tabla grande (cabecera en la fila 30): EDITABLE, nunca se regenera. Cada sincronización sólo:
 *      · añade ARRIBA las filas nuevas (pieza con Reembolso ✓ → "Sin abonar"; abono de factura RM sin pedir → "Sin solicitar"),
 *      · rellena las celdas VACÍAS de la fila de una pieza cuando llega su abono (y "Sin abonar" → "Abonada"),
 *      · borra la fila de una pieza que se desmarca en Piezas.
 *    Y al revés: borrar a mano una fila de pieza aquí la desmarca en Piezas (alCambiar / editarAbonos_).
 *    Cada fila lleva sus claves en la columna oculta "Clave" (Logic.js ▸ clavesPiezas / clavesAbonos).
 */

const COL_ABONOS_ = { fechaAbono: 1, descripcion: 2, sinIva: 3, estado: 5, albaran: 6, referencia: 7, matricula: 8, fechaSolicitud: 9, factura: 10, nota: 11, clave: 13 };
// Marca de que la tabla ya es editable (migrarAbonos_ hecho). No se usa la cabecera "Clave": montarAbonos_ la escribe.
const PROP_ABONOS_EDITABLE_ = 'ABONOS_EDITABLE';
// Columnas que escribe una persona o la sincronización (no las fórmulas D y L): si están todas vacías, la fila se ha borrado.
const COLS_ENTRADA_ABONOS_ = [1, 2, 3, 5, 6, 7, 8, 9, 10, 11];

function sincronizarAbonos_() {
  const crono = cronometro_('sincronizarAbonos_');
  const d = leerParaAbonos_(null, crono);
  crono.fin();
  return escribirAbonos_(d);
}

/**
 * Todo lo que la sincronización necesita LEER, para hacerlo antes de escribir nada: en Sheets, una lectura hecha
 * después de una escritura espera a que se recalculen las fórmulas que dependen de lo escrito.
 * tabP: la tabla de Piezas si quien llama ya la ha leído (y la mantiene al día con lo que escriba después).
 */
function leerParaAbonos_(tabP, crono) {
  if (recolocarTablaAbonos_()) crono.paso('recolocar tabla de Abonos');
  const leida = tabla => `${tabla.filas.length} de ${tabla.leidas} filas`;
  if (!tabP) { tabP = leerTabla_(HOJA.PIEZAS); crono.paso(`leer Piezas (${leida(tabP)})`); }
  const tabL = leerTabla_(HOJA.LINEAS); crono.paso(`leer Líneas RM (${leida(tabL)})`);
  const tabA = leerTabla_(HOJA.ALB); crono.paso(`leer Albaranes (${leida(tabA)})`);
  const conocidas = matriculasConocidas_(); crono.paso(`leer Coches (${conocidas.size})`);
  const sh = hoja_(HOJA.ABONOS), tb = ABONOS.filaTabla, ncol = ABONOS.cabTabla.length, ultFila = sh.getLastRow();
  const tabla = ultFila >= tb ? sh.getRange(tb, 1, ultFila - tb + 1, ncol).getValues() : [];
  const shC = ss_().getSheetByName(HOJA.CLAVES), ultC = shC ? shC.getLastRow() : 0;
  const vistas = new Set(ultC ? shC.getRange(1, 1, ultC, 1).getValues().map(r => String(r[0])).filter(Boolean) : []);
  loc_('');  // la configuración regional también es una lectura: se guarda ya para escribir las fórmulas luego
  crono.paso(`leer Abonos (${tabla.length} filas, ${vistas.size} claves vistas)`);
  const migrar = PropertiesService.getScriptProperties().getProperty(PROP_ABONOS_EDITABLE_) !== '1';
  return { tabP, tabL, tabA, sh, tabla, vistas, migrar, conocidas };
}

/**
 * Piezas de RM (las de proveedor "Otros" no se reclaman) con su clave P. La numeración de la clave cuenta TODAS las
 * piezas, marcadas o no, para que no cambie al marcar o desmarcar otra pieza del mismo albarán y referencia.
 */
function piezasParaAbonos_(tabP, tabA, tabL, conocidas) {
  const proveedorDe = {};
  tabA.filas.forEach(f => { const n = normAlbaran(f.v['Nº albarán']); if (n) proveedorDe[n] = f.v['Proveedor']; });
  const { plateDe, basura } = matriculasDeAlbaranes_(tabA, tabL, conocidas);
  const todas = [];
  tabP.filas.forEach(p => {
    const alb = normAlbaran(p.v['Nº albarán']);
    if (String(proveedorDe[alb] || 'RM') === 'Otros') return;
    todas.push({ fila: p.fila, marcada: p.v['Reembolso'] === true && !!alb, albaran: alb, ref: p.v['Referencia pieza'], desc: p.v['Descripción'],
      sinIva: p.v['Precio descontado sin IVA'], fechaReembolso: aISO_(p.v['Fecha reembolso']), matricula: p.v['Matrícula'] || plateDe[alb] || '' });
  });
  return { todas: clavesPiezas(todas), plateDe, basura };
}

/** Matrículas de la pestaña Coches (normalizadas): cuentan como válidas aunque no tengan forma de matrícula (TALLER…). */
function matriculasConocidas_() {
  return new Set(leerTabla_(HOJA.COCHES).filas.map(f => normPlate(f.v['Matrícula'])).filter(Boolean));
}

/** Validador de matrículas: con forma de matrícula española o dada de alta en Coches. */
function validadorMatricula_(conocidas) { return p => pareceMatricula(p) || (!!conocidas && conocidas.has(normPlate(p))); }

/**
 * {nº albarán: matrícula}: primero la pestaña Albaranes (la que se corrige a mano), después las líneas de compra de las
 * facturas RM, que traen la matrícula de TODOS los albaranes aunque no se hayan escaneado. Sólo valores que sean una
 * matrícula (validadorMatricula_): en el campo MATRICULA de RM a veces hay un bastidor o una referencia.
 * basura: {albarán: ese valor que no es matrícula}, para limpiar lo que copió una versión anterior.
 */
function matriculasDeAlbaranes_(tabA, tabL, conocidas) {
  const valida = validadorMatricula_(conocidas);
  const pares = tabA.filas.map(f => [f.v['Nº albarán'], f.v['Matrícula']]);
  if (tabL) tabL.filas.forEach(f => { if (f.v['Tipo'] === 'Compra') pares.push([f.v['Nº albarán'], f.v['Matrícula']]); });
  const basura = mapaMatriculas(pares, p => !valida(p));
  return { plateDe: mapaMatriculas(pares, valida), basura };
}

/**
 * Aplica a la tabla de Abonos lo leído en leerParaAbonos_. No lee nada de la hoja.
 * filasDesmarcadas: filas de Piezas a las que se acaba de quitar el Reembolso (su fila de Abonos se borra).
 */
function escribirAbonos_(d, filasDesmarcadas) {
  const crono = cronometro_('escribirAbonos_');
  const { tabP, tabL, tabA, sh } = d;
  const hoy = hoyISO_(), tb = ABONOS.filaTabla, ncol = ABONOS.cabTabla.length;
  const { todas, plateDe, basura } = piezasParaAbonos_(tabP, tabA, tabL, d.conocidas);
  const marcadas = todas.filter(p => p.marcada);
  let fechasEscritas = 0;
  marcadas.forEach(p => {
    if (!p.fechaReembolso) { p.fechaReembolso = hoy; actualizarFila_(tabP, p.fila, { 'Fecha reembolso': aFecha_(hoy) }); fechasEscritas++; }
  });
  const quitar = todas.filter(p => (filasDesmarcadas || []).indexOf(p.fila) >= 0).map(p => p.clave);
  crono.paso(`preparar piezas (${marcadas.length} marcadas, ${fechasEscritas} fechas escritas)`);

  const abonos = abonosDeLineas_(tabL, plateDe, d.conocidas);
  if (d.migrar) { migrarAbonos_(d, marcadas, abonos); crono.paso('migrar a tabla editable'); }

  const existentes = d.tabla.map(r => ({ clave: r[12], fechaAbono: r[0], descripcion: r[1], sinIva: r[2], estado: r[4], albaran: r[5], ref: r[6], matricula: r[7], factura: r[9], nota: r[10] }));
  const res = sincronizarAbonos(existentes, marcadas, todas, abonos, d.vistas, quitar, { matriculas: plateDe, basura });
  crono.paso(`cruce (${abonos.length} abonos)`);

  // 1) Celdas vacías de filas existentes (antes de borrar o insertar: los índices aún valen).
  res.cambios.forEach(({ i, v }) => Object.keys(v).forEach(k => {
    sh.getRange(tb + i, COL_ABONOS_[k]).setValue(k === 'fechaAbono' ? aFecha_(v[k]) : v[k]);
  }));
  // 2) Filas de piezas desmarcadas, de abajo arriba.
  res.borrar.slice().sort((a, b) => b - a).forEach(i => sh.deleteRow(tb + i));
  // 3) Filas nuevas ARRIBA, justo debajo de la cabecera. Los rangos del resumen, del panel y de los colores empiezan
  //    en la cabecera, así que crecen solos al insertar aquí.
  const k = res.nuevas.length;
  if (k) {
    sh.insertRowsBefore(tb, k);
    const modelo = sh.getRange(tb + k, 1, 1, ncol), destino = sh.getRange(tb, 1, k, ncol);
    modelo.copyTo(destino, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    modelo.copyTo(destino, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
    destino.setValues(res.nuevas.map((f, j) => filaAbono_(f, tb + j)));
  }
  if (res.registrar.length) registrarClavesAbonos_(res.registrar);
  crono.paso(`escribir (${k} nuevas, ${res.cambios.length} completadas, ${res.borrar.length} borradas)`);
  // Sólo a Ejecuciones: escribirlo en Registro en cada edición costaba ~300 ms y llenaba Registro de filas INFO.
  crono.fin(ESTADOS_ABONO.map(e => e + ' ' + res.nuevas.filter(f => f.estado === e).length).join(', '));
  return res;
}

/** Líneas de abono de las facturas RM, con su clave A. */
function abonosDeLineas_(tabL, plateDe, conocidas) {
  const valida = validadorMatricula_(conocidas);
  return clavesAbonos(tabL.filas.filter(f => f.v['Tipo'] === 'Abono').map(f => ({
    factura: f.v['Nº factura'], fecha: aISO_(f.v['Fecha albarán']), albaranOrigen: albaranOrigen(f.v['Albarán origen']),
    ref: f.v['Referencia'], desc: f.v['Descripción'], importe: parseNumber(f.v['Importe sin IVA']),
    // El bloque de abono de la factura trae la matrícula vacía: la buena es la de su albarán original.
    matricula: plateDe[albaranOrigen(f.v['Albarán origen'])] || (valida(f.v['Matrícula']) ? normPlate(f.v['Matrícula']) : ''),
  })));
}

function filaAbono_(f, r) {
  // "Precio con IVA" y "Días pendiente" son fórmulas: la primera usa el IVA de Config y la segunda se actualiza sola cada día.
  return [aFecha_(f.fechaAbono), f.descripcion, f.sinIva, loc_(`=ROUND($C${r}*(1+IVA),2)`), f.estado, f.albaran, f.referencia, f.matricula,
    aFecha_(f.fechaSolicitud), f.factura, f.nota, loc_(`=IF($E${r}="Sin abonar",TODAY()-$I${r},"")`), f.clave];
}

/** Claves de abonos de factura ya añadidos alguna vez: si se borra su fila a mano, no vuelve a aparecer. */
function registrarClavesAbonos_(claves) {
  const ss = ss_();
  let sh = ss.getSheetByName(HOJA.CLAVES);
  if (!sh) { sh = ss.insertSheet(HOJA.CLAVES); sh.hideSheet(); }
  sh.getRange(sh.getLastRow() + 1, 1, claves.length, 1).setValues(claves.map(c => [c]));
}

/**
 * Paso único de la tabla regenerada (versión anterior) a la tabla editable: monta la columna "Clave" y los rangos
 * que crecen al insertar arriba, pone la clave a cada fila existente y da por vistas las líneas de abono
 * que ya estaban en la tabla. No cambia ningún valor de las filas.
 */
function migrarAbonos_(d, marcadas, abonos) {
  montarAbonos_();
  const filas = d.tabla.map(r => ({ fechaSolicitud: r[8], estado: r[4], descripcion: r[1], sinIva: r[2], albaran: r[5], ref: r[6], factura: r[9] }));
  const claves = clavesDeFilasAntiguas(filas, marcadas, abonos);
  if (claves.length) d.sh.getRange(ABONOS.filaTabla, ABONOS.cabTabla.length, claves.length, 1).setValues(claves.map(c => [c]));
  claves.forEach((c, i) => { d.tabla[i][12] = c; });
  // Sólo las líneas de abono reconocidas en la tabla: si alguna no estaba (p. ej. una factura procesada justo ahora), se añade.
  const nuevas = [].concat(...claves.map(partirClaves)).filter(k => k.indexOf('A|') === 0 && !d.vistas.has(k));
  nuevas.forEach(k => d.vistas.add(k));
  if (nuevas.length) registrarClavesAbonos_(nuevas);
  const sinClave = filas.filter((f, i) => !claves[i] && (f.factura || f.albaran || f.descripcion)).length;
  PropertiesService.getScriptProperties().setProperty(PROP_ABONOS_EDITABLE_, '1');
  log_('INFO', 'migrarAbonos', HOJA.ABONOS, `Tabla editable: ${claves.filter(Boolean).length} filas con clave, ${sinClave} sin reconocer (se quedan como filas manuales)`);
}

/**
 * Filas borradas a mano en Abonos: las piezas marcadas cuya fila ya no está se desmarcan en Piezas.
 * Una fila de abono sin pieza (Sin solicitar) no afecta a nada más.
 */
function desmarcarPiezasBorradas_() {
  recolocarTablaAbonos_();
  const sh = hoja_(HOJA.ABONOS), tb = ABONOS.filaTabla, ncol = ABONOS.cabTabla.length, ult = sh.getLastRow();
  if (PropertiesService.getScriptProperties().getProperty(PROP_ABONOS_EDITABLE_) !== '1') return;  // sin migrar: no se sabe qué fila es de qué pieza
  const enTabla = new Set();
  if (ult >= tb) sh.getRange(tb, ncol, ult - tb + 1, 1).getValues().forEach(r => partirClaves(r[0]).forEach(k => enTabla.add(k)));
  const tabP = leerTabla_(HOJA.PIEZAS);
  const { todas } = piezasParaAbonos_(tabP, leerTabla_(HOJA.ALB), null, null);
  const quitar = todas.filter(p => p.marcada && !enTabla.has(p.clave));
  quitar.forEach(p => actualizarFila_(tabP, p.fila, { 'Reembolso': false, 'Fecha reembolso': '' }));
  if (quitar.length) {
    const lista = quitar.map(p => `${p.albaran} ${p.ref || p.desc}`).join(', ');
    log_('INFO', 'desmarcarPiezasBorradas', HOJA.PIEZAS, `Fila borrada en Abonos → Reembolso desmarcado: ${lista}`);
    toast_(`Reembolso desmarcado en Piezas: ${lista}`, 'Abonos', 8);
  }
}

/**
 * Edición a mano en la tabla de Abonos. Vaciar una fila entera (seleccionarla y pulsar Suprimir) cuenta como borrarla:
 * se quita el hueco y, si era de una pieza, se desmarca en Piezas. Cualquier otro cambio se respeta tal cual.
 */
function editarAbonos_(r0, n) {
  // Si la tabla estaba descolocada, la fila editada ya no es la que era: sólo se recoloca y se revisan las piezas.
  if (recolocarTablaAbonos_()) { desmarcarPiezasBorradas_(); return; }
  const sh = hoja_(HOJA.ABONOS), tb = ABONOS.filaTabla, ncol = ABONOS.cabTabla.length;
  const desde = Math.max(r0, tb), hasta = r0 + n, ult = sh.getLastRow();
  if (hasta <= desde) return;
  const vals = sh.getRange(desde, 1, hasta - desde, ncol).getValues();
  const vacias = [];
  vals.forEach((v, i) => { if (COLS_ENTRADA_ABONOS_.every(c => v[c - 1] === '' || v[c - 1] == null)) vacias.push(desde + i); });
  if (!vacias.length) return;
  // Lo que queda por debajo de la última fila con algo ya está en blanco. De abajo arriba para no mover las siguientes.
  vacias.filter(r => r <= ult).sort((a, b) => b - a).forEach(r => sh.deleteRow(r));
  desmarcarPiezasBorradas_();
}

/**
 * La tabla de Abonos va en posiciones fijas (título en ABONOS.filaTitulo, cabecera en ABONOS.filaCabTabla). Si alguien
 * borra o inserta filas por encima de la cabecera, todo se desplaza y escribir en las filas fijas pisaría datos.
 * Aquí se busca la cabecera real ("Fecha abono" | "Descripción pieza") y se vuelve a poner en su sitio insertando o
 * quitando filas VACÍAS por encima de ella; nunca se toca una fila con datos. Devuelve true si ha movido algo.
 */
function recolocarTablaAbonos_() {
  const sh = hoja_(HOJA.ABONOS), esperada = ABONOS.filaCabTabla, ct = ABONOS.cabTabla;
  const n = Math.min(Math.max(sh.getLastRow(), 1), esperada + 60);
  const vals = sh.getRange(1, 1, n, ct.length).getValues();
  const idx = vals.findIndex(r => r[0] === ct[0] && r[1] === ct[1]);
  if (idx < 0) {
    if (PropertiesService.getScriptProperties().getProperty(PROP_ABONOS_EDITABLE_) !== '1') return false;  // hoja nueva: aún no hay tabla
    throw new Error('No encuentro la cabecera de la tabla de Abonos ("Fecha abono", "Descripción pieza"). ' +
      'Si se ha borrado, deshazlo (Ctrl+Z) o recupera la fila desde Archivo ▸ Historial de versiones.');
  }
  const real = idx + 1;
  if (real === esperada) return false;
  const titulo = 'Piezas reembolsadas y abonos de RM';
  if (real < esperada) {
    sh.insertRowsBefore(real, esperada - real);
  } else {
    // Filas sobrantes entre el resumen y la cabecera: sólo si están vacías (o son el título).
    const libres = [];
    for (let r = ABONOS.filaIni + ABONOS.filas; r < real; r++) {
      const v = vals[r - 1], vacia = v.every(x => x === '' || x == null) || String(v[0]).indexOf(titulo) === 0;
      if (vacia) libres.push(r);
    }
    const sobran = real - esperada;
    if (libres.length < sobran) throw new Error(`La tabla de Abonos está desplazada ${sobran} fila(s) hacia abajo y no hay filas vacías que quitar encima de la cabecera (fila ${real}). Borra a mano las filas que sobren entre el resumen y la cabecera.`);
    libres.slice(-sobran).sort((a, b) => b - a).forEach(r => sh.deleteRow(r));
  }
  sh.getRange(ABONOS.filaTitulo, 1).setValue(TITULO_TABLA_ABONOS).setFontWeight('bold').setFontSize(11);
  log_('AVISO', 'recolocarTablaAbonos', `${HOJA.ABONOS}!A${real}`, `La cabecera de la tabla estaba en la fila ${real} en vez de la ${esperada} (se habían borrado o insertado filas encima). Recolocada sin tocar los datos.`);
  return true;
}

/** Menú: Actualizar Abonos (añade lo que falte; no cambia ni borra nada de lo que ya hay). */
function actualizarAbonos() {
  ejecutar_('actualizarAbonos', () => conBloqueo_(10, () => {
    const r = sincronizarAbonos_();
    toast_(r.nuevas.length ? `Abonos: ${r.nuevas.length} fila(s) nueva(s) arriba.` : 'Abonos ya estaba al día.');
  }));
}

// ---- Recuperación de estados desde el historial de versiones ----

const MIME_XLSX_ = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

/**
 * Menú / editor: recupera el Estado que tenían las filas de Abonos antes de perderse (hasta el 30/09/2026 la tabla se
 * regeneraba entera y borraba los estados puestos a mano). Recorre el historial de versiones de la hoja de la más
 * reciente hacia atrás, se queda con la primera versión que tenga alguna fila "Abonada" y copia su Estado a la misma
 * fila de ahora (por clave) SÓLO si ahora sigue en un estado automático. El detalle queda en la pestaña "Recuperación Abonos".
 */
function recuperarEstadosAbonos() {
  ejecutar_('recuperarEstadosAbonos', () => conBloqueo_(30, () => {
    const t0 = Date.now(), id = ss_().getId(), token = ScriptApp.getOAuthToken();
    const api = (url, opts) => {
      const r = UrlFetchApp.fetch(url, Object.assign({ headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true }, opts || {}));
      if (r.getResponseCode() >= 300) throw new Error(`Drive ${r.getResponseCode()}: ${r.getContentText().slice(0, 300)}`);
      return r;
    };
    let revs = [], pagina = '';
    do {
      const j = JSON.parse(api(`https://www.googleapis.com/drive/v3/files/${id}/revisions?pageSize=200&fields=nextPageToken,revisions(id,modifiedTime)` +
        (pagina ? '&pageToken=' + encodeURIComponent(pagina) : '')).getContentText());
      revs = revs.concat(j.revisions || []);
      pagina = j.nextPageToken || '';
    } while (pagina);
    revs.sort((a, b) => String(b.modifiedTime).localeCompare(String(a.modifiedTime)));
    toast_(`Revisando el historial (${revs.length} versiones)…`, 'Recuperar Abonos', 60);

    // Versiones ya revisadas sin ninguna "Abonada" en una ejecución anterior (si se cortó por tiempo): no se repiten.
    const props = PropertiesService.getScriptProperties(), PROP = 'RECUPERAR_ABONOS_SIN_ABONADA';
    const sinAbonada = new Set(JSON.parse(props.getProperty(PROP) || '[]'));
    const revisadas = [];
    let elegida = null, viejas = null, cortada = false;
    for (const rev of revs) {
      if (sinAbonada.has(rev.id)) continue;
      if (Date.now() - t0 > 240000) { cortada = true; break; }  // límite de 6 min de Apps Script: se sigue en otra ejecución
      const filas = abonosDeVersion_(api, id, rev.id);
      const n = filas.filter(r => r[4] === 'Abonada').length;
      revisadas.push([rev.modifiedTime, rev.id, filas.length, n]);
      if (n > 0) { elegida = rev; viejas = filas; break; }
      sinAbonada.add(rev.id);
    }
    props.setProperty(PROP, JSON.stringify(Array.from(sinAbonada)));

    recolocarTablaAbonos_();
    const tabP = leerTabla_(HOJA.PIEZAS), tabA = leerTabla_(HOJA.ALB), tabL = leerTabla_(HOJA.LINEAS), conocidas = matriculasConocidas_();
    const { todas, plateDe } = piezasParaAbonos_(tabP, tabA, tabL, conocidas);
    const abonos = abonosDeLineas_(tabL, plateDe, conocidas);
    const sh = hoja_(HOJA.ABONOS), tb = ABONOS.filaTabla, ncol = ABONOS.cabTabla.length, ult = sh.getLastRow();
    const actuales = ult >= tb ? sh.getRange(tb, 1, ult - tb + 1, ncol).getValues() : [];

    let plan = [];
    if (elegida) {
      const filas = viejas.map(r => ({ fechaSolicitud: r[8], estado: r[4], descripcion: r[1], sinIva: r[2],
        albaran: normAlbaran(r[5]), ref: r[6], factura: r[9] }));
      const claves = clavesDeFilasAntiguas(filas, todas.filter(p => p.marcada), abonos);
      // Si esa versión ya tenía la columna Clave, se usa la suya.
      plan = planRecuperacion(viejas.map((r, i) => ({ clave: r[12] || claves[i], estado: r[4] })), actuales.map(r => ({ clave: r[12], estado: r[4] })));
      plan.filter(x => x.accion === 'cambiar').forEach(x => sh.getRange(tb + x.i, COL_ABONOS_.estado).setValue(x.antes));
    }
    const cambiadas = plan.filter(x => x.accion === 'cambiar').length;

    const ss = ss_();
    const inf = ss.getSheetByName('Recuperación Abonos') || ss.insertSheet('Recuperación Abonos');
    inf.clear();
    const salida = [
      ['Recuperación de estados de Abonos desde el historial de versiones', '', '', '', '', ''],
      ['Ejecutada', new Date(), '', '', '', ''],
      ['Versión usada', elegida ? elegida.modifiedTime + ' (id ' + elegida.id + ')' : 'ninguna con filas "Abonada" entre las revisadas', '', '', '', ''],
      ['Estados cambiados', cambiadas, '', '', '', ''],
      ['', '', '', '', '', ''],
      ['Versiones revisadas (de la más reciente hacia atrás)', 'Id', 'Filas en Abonos', 'Filas "Abonada"', '', ''],
    ].concat(revisadas.map(r => r.concat(['', ''])))
      .concat([['', '', '', '', '', ''], ['Fila actual', 'Clave', 'Estado en la versión', 'Estado antes de recuperar', 'Acción', '']])
      .concat(plan.map(x => [x.i >= 0 ? tb + x.i : '', x.clave, x.antes, x.ahora, x.accion, '']));
    inf.getRange(1, 1, salida.length, 6).setValues(salida);
    log_('INFO', 'recuperarEstadosAbonos', elegida ? elegida.modifiedTime : '', `${cambiadas} estados recuperados; ${revisadas.length} versiones revisadas`);
    avisar_(elegida
      ? `Versión usada: ${elegida.modifiedTime}\nEstados recuperados: ${cambiadas}\n\nDetalle en la pestaña "Recuperación Abonos".`
      : cortada ? `Se ha acabado el tiempo tras revisar ${revisadas.length} versiones sin encontrar filas "Abonada". Vuelve a ejecutarlo: sigue donde lo ha dejado.`
        : `No hay ninguna versión con filas "Abonada" en el historial. Detalle en "Recuperación Abonos".`, 'Recuperar Abonos');
  }));
}

/** Filas de la tabla de Abonos de una versión antigua de la hoja (exporta esa versión, la abre como copia temporal y la borra). */
function abonosDeVersion_(api, id, revId) {
  const meta = JSON.parse(api(`https://www.googleapis.com/drive/v3/files/${id}/revisions/${revId}?fields=exportLinks`).getContentText());
  const url = (meta.exportLinks || {})[MIME_XLSX_];
  if (!url) return [];
  const xlsx = api(url).getBlob().getBytes();
  const limite = 'cc' + Date.now();
  const cabecera = Utilities.newBlob(`--${limite}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n` +
    JSON.stringify({ name: `tmp recuperación Abonos ${revId}`, mimeType: 'application/vnd.google-apps.spreadsheet' }) +
    `\r\n--${limite}\r\nContent-Type: ${MIME_XLSX_}\r\n\r\n`).getBytes();
  const pie = Utilities.newBlob(`\r\n--${limite}--`).getBytes();
  const copia = JSON.parse(api('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart', {
    method: 'post', contentType: `multipart/related; boundary=${limite}`, payload: cabecera.concat(xlsx, pie),
  }).getContentText());
  try {
    const sh = SpreadsheetApp.openById(copia.id).getSheetByName(HOJA.ABONOS);
    if (!sh) return [];
    const vals = sh.getDataRange().getValues();
    const h = vals.findIndex(r => r[0] === ABONOS.cabTabla[0] && r[1] === ABONOS.cabTabla[1]);
    if (h < 0) return [];
    return vals.slice(h + 1).filter(r => r.slice(0, 12).some(x => x !== '' && x != null)).map(r => r.concat(new Array(13).fill('')).slice(0, 13));
  } finally {
    DriveApp.getFileById(copia.id).setTrashed(true);
  }
}
