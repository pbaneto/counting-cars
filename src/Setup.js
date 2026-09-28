/**
 * setup(): crea/repara pestañas, cabeceras, fórmulas, formato, validaciones, colores y triggers.
 * Es IDEMPOTENTE: se puede ejecutar las veces que haga falta; nunca borra datos de entrada.
 */

const ORDEN_HOJAS = [HOJA.ALB, HOJA.TRAB, HOJA.PIEZAS, HOJA.ABONOS, HOJA.COCHES, HOJA.RESUMEN, HOJA.FACT, HOJA.LINEAS, HOJA.CONFIG, HOJA.REG];
const PROTECCION = 'counting-cars:auto';

function setup() {
  ejecutar_('setup', () => conBloqueo_(30, () => {
    const props = PropertiesService.getScriptProperties();
    if (typeof PRIVATE !== 'undefined' && PRIVATE.SPREADSHEET_ID) props.setProperty('SPREADSHEET_ID', PRIVATE.SPREADSHEET_ID);
    reiniciarCaches_();
    ss_().setSpreadsheetTimeZone(TZ);
    crearHojas_();
    prepararConfig_();
    prepararTablas_();
    montarAbonos_();
    montarResumen_();
    instalarTriggers_();
    log_('INFO', 'setup', '', 'Hoja preparada');
    avisar_('Hoja preparada.\n\nSiguientes pasos:\n1) Menú Counting Cars ▸ Configurar API key de Gemini\n2) Menú Counting Cars ▸ Cargar coches y datos del piloto\n3) Menú Counting Cars ▸ Procesar albaranes', 'Counting Cars');
  }));
}

/** Menú: reescribe fórmulas, formato y colores sin tocar los datos. Añade a Config las claves nuevas (sin pisar valores). */
function repararFormulas() {
  ejecutar_('repararFormulas', () => conBloqueo_(30, () => {
    reiniciarCaches_();
    prepararConfig_();
    prepararTablas_();
    montarAbonos_();
    toast_('Fórmulas y formato reparados.');
  }));
}

function reiniciarCaches_() { _ss = null; _cfg = null; _letras = {}; _pyc = null; }

function crearHojas_() {
  const ss = ss_();
  ORDEN_HOJAS.forEach(n => { if (!ss.getSheetByName(n)) ss.insertSheet(n); });
  ss.getSheets().forEach(sh => {  // hoja por defecto vacía
    if (ORDEN_HOJAS.indexOf(sh.getName()) < 0 && sh.getName() !== HOJA.DIAG && sh.getLastRow() <= 1 && ss.getSheets().length > 1) ss.deleteSheet(sh);
  });
  ORDEN_HOJAS.forEach((n, i) => { ss.setActiveSheet(ss.getSheetByName(n)); ss.moveActiveSheet(i + 1); });
  ss.setActiveSheet(ss.getSheetByName(HOJA.ALB));
}

function estiloCabecera_(sh, n) {
  sh.getRange(1, 1, 1, n).setBackground(COLORES.cabecera).setFontColor('#ffffff').setFontWeight('bold').setVerticalAlignment('middle').setWrap(true);
  sh.setFrozenRows(1);
}

function prepararConfig_() {
  const sh = hoja_(HOJA.CONFIG);
  escribirCabeceras_(sh, ESQUEMA['Config'].cabeceras);
  estiloCabecera_(sh, 3);
  const t = leerTabla_(HOJA.CONFIG);
  const carpetas = (typeof PRIVATE !== 'undefined' && PRIVATE.CARPETAS) || {};
  CONFIG_DEFECTO.forEach(([clave, valor, desc]) => {
    const f = t.filas.find(x => String(x.v['Clave']).trim() === clave);
    const inicial = valor === '' && carpetas[clave] ? carpetas[clave] : valor;
    if (!f) agregarFilas_(t, [{ 'Clave': clave, 'Valor': inicial, 'Descripción': desc }]);
    else { if (f.v['Valor'] === '' && inicial !== '') actualizarFila_(t, f.fila, { 'Valor': inicial }); actualizarFila_(t, f.fila, { 'Descripción': desc }); }
  });
  const t2 = leerTabla_(HOJA.CONFIG), ss = ss_();
  [['IVA', 'IVA'], ['DIAS_AVISO', 'DIAS_AVISO_TRABAJO'], ['TOL_CUADRE', 'TOLERANCIA_CUADRE'], ['DIAS_AVISO_REEMB', 'DIAS_AVISO_REEMBOLSO']].forEach(([nombre, clave]) => {
    const f = t2.filas.find(x => String(x.v['Clave']).trim() === clave);
    ss.setNamedRange(nombre, sh.getRange(f.fila, t2.map['Valor']));
  });
  sh.setColumnWidth(1, 240); sh.setColumnWidth(2, 340); sh.setColumnWidth(3, 620);
  _cfg = null;
}

/** Cabeceras, fórmulas, formato, validaciones y colores de las pestañas de tabla. */
function prepararTablas_() {
  const ss = ss_();
  ['Albaranes', 'Trabajos', 'Piezas', 'Coches', 'Facturas RM', 'Líneas RM', 'Registro'].forEach(nombre => {
    const sh = ss.getSheetByName(nombre), esq = ESQUEMA[nombre], n = esq.cabeceras.length;
    escribirCabeceras_(sh, esq.cabeceras);
    estiloCabecera_(sh, n);
  });
  _letras = {};
  ['Albaranes', 'Trabajos', 'Piezas', 'Líneas RM'].forEach(escribirFormulas_);
  formatoAlbaranes_(); formatoTrabajos_(); formatoPiezas_(); formatoCoches_(); formatoFacturas_(); formatoLineas_(); formatoRegistro_();
}

/**
 * Escribe cada cabecera sólo si esa columna ya la tiene (no hace nada) o está vacía (la rellena).
 * Si la columna tiene OTRA cabecera distinta, no la pisa: lanza un error claro en vez de desalinear en
 * silencio los datos de las filas de abajo con el nombre nuevo (p. ej. una columna "Quincena" heredada
 * de una versión anterior del esquema, que ya no existe en ESQUEMA pero seguía teniendo datos reales).
 */
function escribirCabeceras_(sh, cabeceras) {
  const ancho = Math.max(sh.getLastColumn(), cabeceras.length);
  const actual = ancho > 0 ? sh.getRange(1, 1, 1, ancho).getValues()[0] : [];
  cabeceras.forEach((h, i) => {
    if (actual[i] === h) return;
    if (actual[i]) throw new Error(`"${sh.getName()}": la columna ${colLetra_(i + 1)} tiene la cabecera "${actual[i]}" en vez de "${h}". ` +
      'Corrígelo a mano (renombra o mueve esa columna) antes de reparar, para no desalinear los datos de las filas de abajo.');
    sh.getRange(1, i + 1).setValue(h);
  });
}

/**
 * Fórmulas y casillas SÓLO en las filas con datos: miles de filas vacías con fórmula o casilla hacían que cada
 * lectura de la pestaña (y cada recálculo) arrastrara todas esas filas. Las filas nuevas las reciben al crearlas
 * (agregarFilas_) o al editarlas a mano (alEditar ▸ asegurarFila_). Quita las de las filas vacías.
 */
function escribirFormulas_(nombre) {
  const sh = hoja_(nombre), n = ESQUEMA[nombre].filasFormato;
  if (sh.getMaxRows() < n + 1) sh.insertRowsAfter(sh.getMaxRows(), n + 1 - sh.getMaxRows());
  const t = leerTabla_(nombre), ultDatos = t.libre - 1, ultHoja = sh.getLastRow();
  Object.keys(FORMULAS[nombre]).forEach(h => {
    if (ultDatos >= 2) sh.getRange(2, t.map[h], ultDatos - 1, 1).setFormulas(Array.from({ length: ultDatos - 1 }, (_, i) => [loc_(FORMULAS[nombre][h](i + 2))]));
    if (ultHoja > ultDatos) limpiarFormulasSobrantes_(t, h, ultDatos + 1, ultHoja);
  });
  // Por debajo de los datos una casilla sólo puede valer FALSE (TRUE contaría como dato), así que se quita sin perder nada.
  (t.esq.casillas || []).forEach(h => {
    if (ultDatos >= 2) sh.getRange(2, t.map[h], ultDatos - 1, 1).setDataValidation(checkbox_());
    if (ultHoja > ultDatos) sh.getRange(ultDatos + 1, t.map[h], ultHoja - ultDatos, 1).clearDataValidations().clearContent();
  });
}

/** Vacía una columna calculada por debajo de la última fila con datos. Si ahí hay un valor escrito a mano, no la toca y avisa. */
function limpiarFormulasSobrantes_(t, h, desde, hasta) {
  const rango = t.sh.getRange(desde, t.map[h], hasta - desde + 1, 1);
  const formulas = rango.getFormulas(), valores = rango.getValues();
  const aMano = valores.findIndex((v, i) => !formulas[i][0] && v[0] !== '' && v[0] !== false);
  if (aMano >= 0) {
    log_('AVISO', 'repararFormulas', `${t.nombre}!${colLetra_(t.map[h])}${desde + aMano}`, `Valor escrito a mano en la columna calculada "${h}" por debajo de los datos: esa columna no se limpia`);
    return;
  }
  rango.clearContent();
}

function colDe_(sh, h) { return sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].indexOf(h) + 1; }

/** Aplica formato a una columna por nombre: n filas desde la 2. */
function colFmt_(sh, h, n, o) {
  const c = colDe_(sh, h);
  if (!c) return;
  const r = sh.getRange(2, c, n, 1);
  if (o.fmt) r.setNumberFormat(o.fmt);
  if (o.gris) r.setBackground(COLORES.gris);
  if (o.ancho) sh.setColumnWidth(c, o.ancho);
  if (o.validacion) r.setDataValidation(o.validacion);
  if (o.gris) proteger_(r);
}

function proteger_(rango) { rango.protect().setDescription(PROTECCION).setWarningOnly(true); }

function limpiarProtecciones_(sh) {
  sh.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => { if (p.getDescription() === PROTECCION) p.remove(); });
}

function regla_(sh, a1, formula, color) {
  return SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(loc_(formula)).setBackground(color).setRanges([sh.getRange(a1)]).build();
}

const FMT = { fecha: 'dd/mm/yyyy', fechaHora: 'dd/mm/yyyy hh:mm', euro: '#,##0.00 "€"', texto: '@', pct: '0.##%' };
function checkbox_() { return SpreadsheetApp.newDataValidation().requireCheckbox().build(); }
function listaValidacion_(vals) { return SpreadsheetApp.newDataValidation().requireValueInList(vals, true).setAllowInvalid(false).build(); }
function matriculaValidacion_() {
  return SpreadsheetApp.newDataValidation().requireValueInRange(hoja_(HOJA.COCHES).getRange('A2:A5000'), true).setAllowInvalid(true)
    .setHelpText('Esta matrícula no está en la pestaña Coches. Puedes continuar, pero conviene añadir el coche.').build();
}

function formatoAlbaranes_() {
  const sh = hoja_(HOJA.ALB), n = ESQUEMA['Albaranes'].filasFormato, l = letras_(HOJA.ALB);
  limpiarProtecciones_(sh);
  colFmt_(sh, 'Fecha escaneo', n, { fmt: FMT.fecha, ancho: 105 });
  colFmt_(sh, 'Fecha albarán', n, { fmt: FMT.fecha, ancho: 105 });
  colFmt_(sh, 'Quincena', n, { gris: true, ancho: 75 });
  colFmt_(sh, 'Proveedor', n, { validacion: listaValidacion_(['RM', 'Otros']), ancho: 85 });
  colFmt_(sh, 'Nº albarán', n, { fmt: FMT.texto, ancho: 95 });
  colFmt_(sh, 'Nº trabajo', n, { ancho: 95 });
  colFmt_(sh, 'Matrícula', n, { validacion: matriculaValidacion_(), ancho: 100 });
  colFmt_(sh, 'Precio con IVA', n, { fmt: FMT.euro, ancho: 110 });
  colFmt_(sh, 'Precio facturable', n, { fmt: FMT.euro, gris: true, ancho: 120 });
  colFmt_(sh, 'Coche', n, { gris: true, ancho: 190 });
  colFmt_(sh, 'Cliente', n, { gris: true, ancho: 150 });
  colFmt_(sh, 'Ver PDF', n, { ancho: 80 });
  colFmt_(sh, 'Avisos', n, { gris: true, ancho: 380 });
  colFmt_(sh, 'Nota escaneo', n, { ancho: 320 });
  const R = h => `${l[h]}2:${l[h]}${n + 1}`;
  sh.setConditionalFormatRules([
    regla_(sh, R('Avisos'), `=$${l['Avisos']}2<>""`, COLORES.naranja),
    regla_(sh, R('Nota escaneo'), `=$${l['Nota escaneo']}2<>""`, COLORES.naranja),
    regla_(sh, R('Coche'), `=LEFT($${l['Coche']}2,1)="⚠"`, COLORES.naranja),
  ]);
}

function formatoTrabajos_() {
  const sh = hoja_(HOJA.TRAB), n = ESQUEMA['Trabajos'].filasFormato, l = letras_(HOJA.TRAB);
  limpiarProtecciones_(sh);
  colFmt_(sh, 'Nº trabajo', n, { ancho: 95 });
  colFmt_(sh, 'Fecha apertura', n, { fmt: FMT.fecha, ancho: 105 });
  colFmt_(sh, 'Matrícula', n, { validacion: matriculaValidacion_(), ancho: 100 });
  colFmt_(sh, 'Coche', n, { gris: true, ancho: 190 });
  colFmt_(sh, 'Cliente', n, { gris: true, ancho: 150 });
  colFmt_(sh, 'Recambios', n, { fmt: FMT.euro, gris: true, ancho: 110 });
  colFmt_(sh, 'Recambios facturables', n, { fmt: FMT.euro, gris: true, ancho: 150 });
  colFmt_(sh, 'Factura', n, { fmt: FMT.euro, ancho: 110 });
  colFmt_(sh, 'Beneficio', n, { fmt: FMT.euro, gris: true, ancho: 110 });
  colFmt_(sh, 'Pagado', n, { ancho: 80 });
  colFmt_(sh, 'Avisos', n, { gris: true, ancho: 340 });
  const R = h => `${l[h]}2:${l[h]}${n + 1}`;
  sh.setConditionalFormatRules([
    regla_(sh, R('Pagado'), `=AND($${l['Nº trabajo']}2<>"",$${l['Pagado']}2<>TRUE)`, COLORES.rojo),
    regla_(sh, R('Pagado'), `=$${l['Pagado']}2=TRUE`, COLORES.verde),
    regla_(sh, R('Avisos'), `=LEFT($${l['Avisos']}2,1)="⚠"`, COLORES.naranja),
    regla_(sh, R('Avisos'), `=LEFT($${l['Avisos']}2,1)="ℹ"`, COLORES.azul),
    regla_(sh, R('Coche'), `=LEFT($${l['Coche']}2,1)="⚠"`, COLORES.naranja),
  ]);
}

function formatoPiezas_() {
  const sh = hoja_(HOJA.PIEZAS), n = ESQUEMA['Piezas'].filasFormato, l = letras_(HOJA.PIEZAS);
  limpiarProtecciones_(sh);
  colFmt_(sh, 'Reembolso', n, { ancho: 85 });
  colFmt_(sh, 'Matrícula', n, { gris: true, ancho: 100 });
  colFmt_(sh, 'Nº albarán', n, { fmt: FMT.texto, ancho: 95 });
  colFmt_(sh, 'Referencia pieza', n, { ancho: 150 });
  colFmt_(sh, 'Descripción', n, { ancho: 300 });
  colFmt_(sh, 'Marca', n, { ancho: 100 });
  colFmt_(sh, 'Cantidad', n, { ancho: 75 });
  colFmt_(sh, 'Precio base', n, { fmt: FMT.euro, ancho: 100 });
  colFmt_(sh, 'Descuento aplicado', n, { fmt: FMT.pct, ancho: 100 });
  colFmt_(sh, 'Precio descontado sin IVA', n, { fmt: FMT.euro, ancho: 130 });
  colFmt_(sh, 'Precio descontado con IVA', n, { fmt: FMT.euro, gris: true, ancho: 130 });
  colFmt_(sh, 'Fecha reembolso', n, { fmt: FMT.fecha, ancho: 110 });
  colFmt_(sh, 'Proveedor', n, { gris: true, ancho: 85 });
  colFmt_(sh, 'Origen', n, { ancho: 80 });
  colFmt_(sh, 'Avisos', n, { gris: true, ancho: 340 });
  const ult = colLetra_(ESQUEMA['Piezas'].cabeceras.length);
  sh.setConditionalFormatRules([
    regla_(sh, `A2:${ult}${n + 1}`, `=$${l['Reembolso']}2=TRUE`, COLORES.amarillo),
    regla_(sh, `${l['Avisos']}2:${l['Avisos']}${n + 1}`, `=$${l['Avisos']}2<>""`, COLORES.naranja),
  ]);
}

function formatoCoches_() {
  const sh = hoja_(HOJA.COCHES);
  sh.setColumnWidth(1, 110); sh.setColumnWidth(2, 220); sh.setColumnWidth(3, 240);
  sh.getRange(2, 1, 5000, 1).setNumberFormat(FMT.texto);
}

function formatoFacturas_() {
  const sh = hoja_(HOJA.FACT), l = letras_(HOJA.FACT), n = 200;
  colFmt_(sh, 'Fecha factura', n, { fmt: FMT.fecha, ancho: 105 });
  colFmt_(sh, 'Base imponible', n, { fmt: FMT.euro, ancho: 110 });
  colFmt_(sh, 'IVA', n, { fmt: FMT.euro, ancho: 100 });
  colFmt_(sh, 'Total', n, { fmt: FMT.euro, ancho: 110 });
  colFmt_(sh, 'Estado', n, { ancho: 480 });
  colFmt_(sh, 'Procesada el', n, { fmt: FMT.fechaHora, ancho: 130 });
  sh.setConditionalFormatRules([regla_(sh, `${l['Estado']}2:${l['Estado']}${n + 1}`, `=LEFT($${l['Estado']}2,1)="⚠"`, COLORES.naranja)]);
}

function formatoLineas_() {
  const sh = hoja_(HOJA.LINEAS), n = ESQUEMA['Líneas RM'].filasFormato, l = letras_(HOJA.LINEAS);
  limpiarProtecciones_(sh);
  colFmt_(sh, 'Nº albarán', n, { fmt: FMT.texto, ancho: 95 });
  colFmt_(sh, 'Fecha albarán', n, { fmt: FMT.fecha, ancho: 105 });
  colFmt_(sh, 'Albarán origen', n, { fmt: FMT.texto, ancho: 105 });
  colFmt_(sh, 'Descripción', n, { ancho: 300 });
  colFmt_(sh, 'Descuento', n, { fmt: FMT.pct });
  colFmt_(sh, 'Importe sin IVA', n, { fmt: FMT.euro, ancho: 120 });
  colFmt_(sh, 'Conciliación', n, { gris: true, ancho: 190 });
  sh.setConditionalFormatRules([regla_(sh, `${l['Conciliación']}2:${l['Conciliación']}${n + 1}`, `=LEFT($${l['Conciliación']}2,1)="⚠"`, COLORES.naranja)]);
}

function formatoRegistro_() {
  const sh = hoja_(HOJA.REG), l = letras_(HOJA.REG);
  sh.setColumnWidth(1, 140); sh.setColumnWidth(2, 70); sh.setColumnWidth(3, 170); sh.setColumnWidth(4, 200); sh.setColumnWidth(5, 800);
  sh.getRange(2, 1, 3000, 1).setNumberFormat('dd/mm/yyyy hh:mm:ss');
  sh.setConditionalFormatRules([
    regla_(sh, 'A2:E3000', `=$${l['Nivel']}2="ERROR"`, COLORES.rojo),
    regla_(sh, 'A2:E3000', `=$${l['Nivel']}2="AVISO"`, COLORES.naranja),
  ]);
}

/** Resumen por quincena (arriba) + panel "Pendientes de RM" + tabla grande de piezas reembolsadas y abonos (debajo). */
function montarAbonos_() {
  const sh = hoja_(HOJA.ABONOS), a = letras_(HOJA.ALB), f = letras_(HOJA.FACT);
  limpiarProtecciones_(sh);
  const cab = ABONOS.cabResumen, ini = ABONOS.filaIni, tb = ABONOS.filaTabla, fin = tb + ABONOS.maxTabla - 1;
  if (sh.getMaxRows() < fin) sh.insertRowsAfter(sh.getMaxRows(), fin - sh.getMaxRows());
  // Limpia la zona de resumen+panel (filas 1..27) antes de reescribirla: si una versión anterior tenía más o menos
  // columnas, no se queda una cabecera, fórmula o validación fantasma en una columna que ya no se reescribe.
  sh.getRange(1, 1, fin, 20).clearDataValidations();
  sh.getRange(1, 1, ini + ABONOS.filas - 1, 20).clearContent().clearFormat();
  sh.getRange('A1').setValue('Año').setFontWeight('bold').setHorizontalAlignment('right');
  if (sh.getRange(ABONOS.celdaAnio).getValue() === '') sh.getRange(ABONOS.celdaAnio).setValue(2026);
  sh.getRange(ABONOS.celdaAnio).setFontWeight('bold').setBackground(COLORES.amarillo).setNumberFormat('0');
  sh.getRange('D1').setValue('Cuadre: (Recambios − Abonado) = Total factura  ·  Naranja = revisar la factura').setFontStyle('italic').setFontColor('#666666');

  sh.getRange(ABONOS.filaCabResumen, 1, 1, cab.length).setValues([cab]).setBackground(COLORES.cabecera).setFontColor('#ffffff').setFontWeight('bold').setWrap(true).setVerticalAlignment('middle');
  // Rangos de la tabla grande (piezas/abonos), usados tanto por el cuadre por quincena como por el panel de pendientes.
  const rg = col => `$${col}$${tb}:$${col}$${fin}`, ESTADO_ = rg('E'), CONIVA_ = rg('D'), DIAS_ = rg('L');
  for (let k = 0; k < ABONOS.filas; k++) {
    const r = ini + k, mes = Math.floor(k / 2) + 1, q = (k % 2) + 1;
    if (q === 1) { sh.getRange(r, 1, 2, 1).merge().setValue(MESES[mes - 1]).setVerticalAlignment('middle').setHorizontalAlignment('center').setFontWeight('bold'); }
    const desde = `DATE($B$1,${mes},${q === 1 ? 1 : 16})`, hasta = q === 1 ? `DATE($B$1,${mes},15)` : `EOMONTH(DATE($B$1,${mes},1),0)`;
    const rangoFecha = col => `${col},">="&${desde},${col},"<="&${hasta}`;
    const fA = `Albaranes!$${a['Fecha albarán']}:$${a['Fecha albarán']}`;
    const recambios = `=SUMIFS(Albaranes!$${a['Precio con IVA']}:$${a['Precio con IVA']},Albaranes!$${a['Proveedor']}:$${a['Proveedor']},"RM",${rangoFecha(fA)})`;
    const abonado = `=SUMIFS(${CONIVA_},${ESTADO_},"Abonada",${rangoFecha(rg('A'))})+SUMIFS(${CONIVA_},${ESTADO_},"Sin solicitar",${rangoFecha(rg('A'))})`;
    const crit = `'Facturas RM'!$${f['Año']}:$${f['Año']},$B$1,'Facturas RM'!$${f['Mes']}:$${f['Mes']},${mes},'Facturas RM'!$${f['Quincena']}:$${f['Quincena']},${q}`;
    const totalFactura = `=IF(COUNTIFS(${crit})=0,"",SUMIFS('Facturas RM'!$${f['Total']}:$${f['Total']},${crit}))`;
    // Sólo compara la factura de esta quincena con lo que ELLA MISMA contiene (compras + sus propios abonos): no depende de cuándo se pidió el reembolso.
    const estado = `=IF(E${r}="",IF(AND(C${r}=0,D${r}=0),"","· Sin factura escaneada"),` +
      `IF(ABS(C${r}-D${r}-E${r})>TOL_CUADRE,"⚠ La factura no cuadra con los albaranes (dif. "&TEXT(E${r}-(C${r}-D${r}),"0.00")&" €): ¿falta algún albarán o hay un error de escaneo?","✔ Cuadra"))`;
    sh.getRange(r, 2, 1, 5).setValues([locFila_([q, recambios, abonado, totalFactura, estado])]);
  }
  sh.getRange(ini, 2, ABONOS.filas, 1).setHorizontalAlignment('center');
  sh.getRange(ini, 3, ABONOS.filas, 3).setNumberFormat(FMT.euro).setBackground(COLORES.gris);
  sh.getRange(ini, 6, ABONOS.filas, 1).setBackground(COLORES.gris);
  proteger_(sh.getRange(ini, 1, ABONOS.filas, cab.length));

  // ---- Panel "Pendientes de RM": piezas 'Sin abonar' de toda la tabla, no atadas a la quincena en que se pidieron ----
  // El umbral se copia a una celda de ESTA pestaña (UMBRAL_): una regla de formato condicional no puede leer Config.
  const pc = ABONOS.panelCol, pv = pc + 1, UMBRAL_ = `$${colLetra_(pv)}$4`;
  sh.getRange(1, pc, 1, 2).merge().setValue('Pendientes de RM').setBackground(COLORES.cabecera).setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(2, pc).setValue('Piezas sin abonar');
  sh.getRange(2, pv).setValue(loc_(`=COUNTIF(${ESTADO_},"Sin abonar")`));
  sh.getRange(3, pc).setValue('Importe pendiente');
  sh.getRange(3, pv).setValue(loc_(`=SUMIF(${ESTADO_},"Sin abonar",${CONIVA_})`)).setNumberFormat(FMT.euro);
  sh.getRange(4, pc).setValue('Aviso a partir de (días)');
  sh.getRange(4, pv).setValue('=DIAS_AVISO_REEMB');
  sh.getRange(5, pc).setValue('Fuera de plazo');
  sh.getRange(5, pv).setValue(loc_(`=COUNTIFS(${ESTADO_},"Sin abonar",${DIAS_},">"&${UMBRAL_})`));
  sh.getRange(6, pc).setValue('Más antigua (días)');
  sh.getRange(6, pv).setValue(loc_(`=IFERROR(MAXIFS(${DIAS_},${ESTADO_},"Sin abonar"),0)`));
  sh.getRange(2, pc, 5, 1).setFontWeight('bold');
  sh.getRange(2, pv, 5, 1).setHorizontalAlignment('center');
  sh.setColumnWidth(pc, 170); sh.setColumnWidth(pv, 90);

  sh.getRange(ABONOS.filaTitulo, 1).setValue('Piezas reembolsadas y abonos de RM — se rellena sola desde Piezas (Reembolso ✓) y las facturas RM. No editar a mano.').setFontWeight('bold').setFontSize(11);
  const ct = ABONOS.cabTabla;
  sh.getRange(ABONOS.filaCabTabla, 1, 1, ct.length).setValues([ct]).setBackground(COLORES.cabecera).setFontColor('#ffffff').setFontWeight('bold').setWrap(true).setVerticalAlignment('middle');
  const n = ABONOS.maxTabla;
  sh.getRange(tb, 1, n, 1).setNumberFormat(FMT.fecha);
  sh.getRange(tb, 3, n, 2).setNumberFormat(FMT.euro);
  sh.getRange(tb, 5, n, 1).setDataValidation(listaValidacion_(ESTADOS_ABONO));
  sh.getRange(tb, 6, n, 1).setNumberFormat(FMT.texto);
  sh.getRange(tb, 9, n, 1).setNumberFormat(FMT.fecha);
  sh.getRange(tb, 12, n, 1).setNumberFormat('0');
  [90, 80, 130, 130, 130, 150, 110, 110, 120, 110, 120, 100].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.setColumnWidth(2, 260); sh.setColumnWidth(8, 420); sh.setColumnWidth(11, 260);
  sh.setFrozenRows(ABONOS.filaCabResumen);
  sh.setConditionalFormatRules([
    regla_(sh, `A${ini}:${colLetra_(cab.length)}${ini + ABONOS.filas - 1}`, `=LEFT($F${ini},1)="⚠"`, COLORES.naranja),
    regla_(sh, `${colLetra_(pc)}5:${colLetra_(pv)}5`, `=${colLetra_(pv)}5>0`, COLORES.naranja),
    regla_(sh, `A${tb}:L${fin}`, `=AND($E${tb}="Sin abonar",$L${tb}>${UMBRAL_})`, COLORES.naranja),
    regla_(sh, `A${tb}:L${fin}`, `=$E${tb}="Abonada"`, COLORES.verde),
    regla_(sh, `A${tb}:L${fin}`, `=$E${tb}="Sin abonar"`, COLORES.rojo),
    regla_(sh, `A${tb}:L${fin}`, `=$E${tb}="Sin solicitar"`, COLORES.amarillo),
  ]);
}

function instalarTriggers_() {
  const ss = ss_();
  ScriptApp.getProjectTriggers().forEach(t => { if (['alEditar', 'alAbrir'].indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('alEditar').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('alAbrir').forSpreadsheet(ss).onOpen().create();
}
