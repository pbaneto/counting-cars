/**
 * Diseño de la hoja: pestañas, cabeceras, fórmulas, formato, validaciones, colores y triggers. Nunca borra datos.
 * Respeta lo que se haya cambiado a mano: anchos de columna (sólo se ponen al crear la hoja), filtros, filas fijas de
 * más y reglas de color propias.
 * Se aplica solo cuando cambia la versión del código (ver Version.js); el menú "Reparar fórmulas y formato" lo fuerza.
 */

const ORDEN_HOJAS = [HOJA.ALB, HOJA.TRAB, HOJA.PIEZAS, HOJA.ABONOS, HOJA.COCHES, HOJA.RESUMEN, HOJA.FACT, HOJA.LINEAS, HOJA.CONFIG, HOJA.REG, HOJA.CLAVES];

/** true sólo al preparar una hoja nueva (setup): es cuando se ponen los anchos de columna. */
let _hojaNueva = false;

function setup() {
  ejecutar_('setup', () => conBloqueo_(30, () => {
    const props = PropertiesService.getScriptProperties();
    if (typeof PRIVATE !== 'undefined' && PRIVATE.SPREADSHEET_ID) props.setProperty('SPREADSHEET_ID', PRIVATE.SPREADSHEET_ID);
    reiniciarCaches_();
    ss_().setSpreadsheetTimeZone(TZ);
    crearHojas_();
    _hojaNueva = true;
    try { aplicarDiseno_(); } finally { _hojaNueva = false; }
    props.setProperty(PROP_VERSION_, VERSION);
    log_('INFO', 'setup', '', `Hoja preparada (versión ${VERSION})`);
    avisar_('Hoja preparada.\n\nSiguientes pasos:\n1) Menú Counting Cars ▸ Configurar API key de Gemini\n2) Si hace falta cargar coches y datos del piloto, ejecuta cargarDatosIniciales() desde el editor de Apps Script (no está en el menú)\n3) Menú Counting Cars ▸ Procesar albaranes', 'Counting Cars');
  }));
}

/** Menú: vuelve a aplicar el diseño (p. ej. si se ha borrado una fórmula a mano). No toca los datos. */
function repararFormulas() {
  ejecutar_('repararFormulas', () => conBloqueo_(30, () => {
    actualizarHoja_(true);
    toast_('Fórmulas y formato reparados.');
  }));
}

/** Todo el diseño, con los tiempos de cada paso en Registro. */
function aplicarDiseno_() {
  const crono = cronometro_('aplicarDiseno');
  reiniciarCaches_();
  prepararConfig_(); crono.paso('Config');
  prepararTablas_(crono);
  montarAbonos_(); crono.paso('Abonos');
  montarResumen_(); crono.paso('Resumen');
  instalarTriggers_(); crono.paso('triggers');
  log_('INFO', 'aplicarDiseno', '', crono.fin());
}

/** Ancho de columna: sólo en una hoja nueva. Después manda el ancho que haya puesto cada uno. */
function ancho_(sh, c, w) { if (_hojaNueva) sh.setColumnWidth(c, w); }

/** Filas fijas: como mínimo las que necesita la cabecera; si alguien ha fijado más, se respetan. */
function fijarFilas_(sh, n) { if (sh.getFrozenRows() < n) sh.setFrozenRows(n); }

/**
 * Reglas de color: se cambian las nuestras (las que tienen la misma fórmula, sin contar números de fila) y se
 * conservan las que haya añadido alguien a mano.
 */
function ponerReglas_(sh, reglas) {
  const norm = f => String(f || '').replace(/(\$?[A-Z]{1,3})\$?\d+/g, '$1').replace(/\s+/g, '').toUpperCase();
  const nuestras = new Set(reglas.map(r => norm(r.formula)));
  const formulaDe = r => { try { const c = r.getBooleanCondition(); return c ? c.getCriteriaValues()[0] : ''; } catch (e) { return ''; } };
  const ajenas = (sh.getConditionalFormatRules() || []).filter(r => !nuestras.has(norm(formulaDe(r))));
  sh.setConditionalFormatRules(ajenas.concat(reglas.map(r => r.regla)));
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

function estiloCabecera_(sh, n, filaCab) {
  filaCab = filaCab || 1;
  sh.getRange(filaCab, 1, 1, n).setBackground(COLORES.cabecera).setFontColor('#ffffff').setFontWeight('bold').setVerticalAlignment('middle').setWrap(true);
  fijarFilas_(sh, filaCab);  // deja fijo también lo que haya encima (el panel de resumen, si lo hay)
}

function prepararConfig_() {
  const sh = hoja_(HOJA.CONFIG);
  escribirCabeceras_(sh, ESQUEMA['Config'].cabeceras);
  const t = leerTabla_(HOJA.CONFIG);
  let añadidas = false;
  const carpetas = (typeof PRIVATE !== 'undefined' && PRIVATE.CARPETAS) || {};
  CONFIG_DEFECTO.forEach(([clave, valor, desc]) => {
    const f = t.filas.find(x => String(x.v['Clave']).trim() === clave);
    const inicial = valor === '' && carpetas[clave] ? carpetas[clave] : valor;
    if (!f) { agregarFilas_(t, [{ 'Clave': clave, 'Valor': inicial, 'Descripción': desc }]); añadidas = true; return; }
    // Sólo se escribe lo que cambia: cada escritura innecesaria hace que la siguiente lectura espere al recálculo.
    if (f.v['Valor'] === '' && inicial !== '') actualizarFila_(t, f.fila, { 'Valor': inicial });
    if (f.v['Descripción'] !== desc) actualizarFila_(t, f.fila, { 'Descripción': desc });
  });
  const t2 = añadidas ? leerTabla_(HOJA.CONFIG) : t, ss = ss_();
  [['IVA', 'IVA'], ['DIAS_AVISO', 'DIAS_AVISO_TRABAJO'], ['DIAS_AVISO_REEMB', 'DIAS_AVISO_REEMBOLSO']].forEach(([nombre, clave]) => {
    const f = t2.filas.find(x => String(x.v['Clave']).trim() === clave);
    ss.setNamedRange(nombre, sh.getRange(f.fila, t2.map['Valor']));
  });
  estiloCabecera_(sh, 3);
  ancho_(sh, 1, 240); ancho_(sh, 2, 340); ancho_(sh, 3, 620);
  _cfg = null;
}



/**
 * Cabeceras, fórmulas, formato, validaciones y colores de las pestañas de tabla.
 * Primero se LEE todo (valores y fórmulas de cada pestaña) y después se escribe sólo lo que cambia: en Sheets, cada
 * lectura hecha tras una escritura espera a que se recalcule la hoja, y con las SUMIFS de Trabajos eso son segundos.
 */
function prepararTablas_(crono) {
  const paso = etiqueta => crono && crono.paso(etiqueta);
  const ss = ss_(), TABLAS = ['Albaranes', 'Trabajos', 'Piezas', 'Coches', 'Facturas RM', 'Líneas RM', 'Registro'];
  const leido = {};
  TABLAS.forEach(nombre => {
    const sh = ss.getSheetByName(nombre), rango = sh.getDataRange(), conFormulas = !!FORMULAS[nombre];
    leido[nombre] = { sh, valores: rango.getValues(), formulas: conFormulas ? rango.getFormulas() : null, maxRows: conFormulas ? sh.getMaxRows() : 0 };
  });
  loc_('');  // el idioma de la hoja también es una lectura: se guarda ya para escribir las fórmulas luego
  paso('leer');
  let cabNuevas = false;
  TABLAS.forEach(nombre => {
    const esq = ESQUEMA[nombre], l = leido[nombre];
    if (escribirCabeceras_(l.sh, esq.cabeceras, esq.filaCabecera, l.valores)) cabNuevas = true;
    estiloCabecera_(l.sh, esq.cabeceras.length, esq.filaCabecera);
  });
  // escribirCabeceras_ garantiza que cada cabecera está en su posición de ESQUEMA: las letras salen de ahí, sin leer la hoja.
  _letras = {};
  TABLAS.forEach(nombre => { const o = {}; ESQUEMA[nombre].cabeceras.forEach((h, i) => { o[h] = colLetra_(i + 1); }); _letras[nombre] = o; });
  paso('cabeceras');
  // Si se ha escrito alguna cabecera (pestaña recién creada), lo leído ya no vale: escribirFormulas_ vuelve a leer.
  ['Albaranes', 'Trabajos', 'Piezas', 'Líneas RM'].forEach(nombre => { escribirFormulas_(nombre, cabNuevas ? null : leido[nombre]); paso(`fórmulas ${nombre}`); });
  [['Albaranes', formatoAlbaranes_], ['Trabajos', formatoTrabajos_], ['Piezas', formatoPiezas_], ['Coches', formatoCoches_],
    ['Facturas RM', formatoFacturas_], ['Líneas RM', formatoLineas_], ['Registro', formatoRegistro_]].forEach(([nombre, fn]) => { fn(); paso(`formato ${nombre}`); });
}

/**
 * Escribe cada cabecera sólo si esa columna ya la tiene (no hace nada) o está vacía (la rellena).
 * Si la columna tiene OTRA cabecera distinta, no la pisa: lanza un error claro en vez de desalinear en
 * silencio los datos de las filas de abajo con el nombre nuevo (p. ej. una columna "Quincena" heredada
 * de una versión anterior del esquema, que ya no existe en ESQUEMA pero seguía teniendo datos reales).
 */
function escribirCabeceras_(sh, cabeceras, filaCab, valores) {
  filaCab = filaCab || 1;
  let actual;
  if (valores) actual = valores[filaCab - 1] || [];
  else { const ancho = Math.max(sh.getLastColumn(), cabeceras.length); actual = ancho > 0 ? sh.getRange(filaCab, 1, 1, ancho).getValues()[0] : []; }
  let escritas = false;
  cabeceras.forEach((h, i) => {
    if (actual[i] === h) return;
    if (actual[i]) throw new Error(`"${sh.getName()}": la columna ${colLetra_(i + 1)} tiene la cabecera "${actual[i]}" en vez de "${h}". ` +
      'Corrígelo a mano (renombra o mueve esa columna) antes de reparar, para no desalinear los datos de las filas de abajo.');
    sh.getRange(filaCab, i + 1).setValue(h);
    escritas = true;
  });
  return escritas;
}

/**
 * Fórmulas y casillas SÓLO en las filas con datos: miles de filas vacías con fórmula o casilla hacían que cada
 * lectura de la pestaña (y cada recálculo) arrastrara todas esas filas. Las filas nuevas las reciben al crearlas
 * (agregarFilas_) o al editarlas a mano (alEditar ▸ asegurarFila_). Quita las de las filas vacías.
 * leido: {sh, valores, formulas, maxRows} ya leídos por prepararTablas_ (si no, se leen aquí). Una columna sólo se
 * reescribe si alguna de sus fórmulas ha cambiado: reescribirla obliga a recalcular todo lo que depende de ella.
 */
function escribirFormulas_(nombre, leido) {
  const esq = ESQUEMA[nombre], n = esq.filasFormato, fc = esq.filaCabecera || 1;
  const sh = leido ? leido.sh : hoja_(nombre);
  const maxRows = leido ? leido.maxRows : sh.getMaxRows();
  if (maxRows < n + fc) sh.insertRowsAfter(maxRows, n + fc - maxRows);
  const t = leerTabla_(nombre, leido && leido.valores), ultDatos = t.libre - 1, ultHoja = t.valores.length;
  const formulas = leido ? leido.formulas : (ultHoja ? sh.getRange(1, 1, ultHoja, Math.max(1, sh.getLastColumn())).getFormulas() : []);
  const actual = (r, c) => (formulas[r - 1] || [])[c - 1] || '';
  Object.keys(FORMULAS[nombre]).forEach(h => {
    const c = t.map[h];
    if (ultDatos >= fc + 1) {
      const deseadas = Array.from({ length: ultDatos - fc }, (_, i) => FORMULAS[nombre][h](i + fc + 1));
      if (deseadas.some((f, i) => !mismaFormula_(actual(i + fc + 1, c), f))) sh.getRange(fc + 1, c, ultDatos - fc, 1).setFormulas(deseadas.map(f => [loc_(f)]));
    }
    if (ultHoja > ultDatos) limpiarFormulasSobrantes_(t, h, ultDatos + 1, ultHoja, actual);
  });
  // Por debajo de los datos una casilla sólo puede valer FALSE (TRUE contaría como dato), así que se quita sin perder nada.
  (t.esq.casillas || []).forEach(h => {
    if (ultDatos >= fc + 1) sh.getRange(fc + 1, t.map[h], ultDatos - fc, 1).setDataValidation(checkbox_());
    if (ultHoja > ultDatos) sh.getRange(ultDatos + 1, t.map[h], ultHoja - ultDatos, 1).clearDataValidations().clearContent();
  });
}

/** ¿La fórmula que hay en la celda es ya la deseada? (Sheets puede devolverla con ; o , según el idioma de la hoja). */
function mismaFormula_(actual, deseada) {
  if (!actual) return false;
  if (actual === deseada || actual === loc_(deseada)) return true;
  const norm = f => String(f).replace(/\s+/g, '').toUpperCase();
  return norm(actual) === norm(deseada) || norm(actual) === norm(loc_(deseada));
}

/**
 * Vacía una columna calculada por debajo de la última fila con datos, usando lo ya leído (sin volver a leer la hoja).
 * Si ahí hay un valor escrito a mano, no la toca y avisa. Si ya está vacía, no escribe nada.
 */
function limpiarFormulasSobrantes_(t, h, desde, hasta, formulaEn) {
  const c = t.map[h];
  let hayAlgo = false;
  for (let r = desde; r <= hasta; r++) {
    const v = (t.valores[r - 1] || [])[c - 1], f = formulaEn(r, c);
    if (!f && v !== '' && v !== false && v != null) {
      log_('AVISO', 'repararFormulas', `${t.nombre}!${colLetra_(c)}${r}`, `Valor escrito a mano en la columna calculada "${h}" por debajo de los datos: esa columna no se limpia`);
      return;
    }
    if (f || (v !== '' && v != null)) hayAlgo = true;
  }
  if (hayAlgo) t.sh.getRange(desde, c, hasta - desde + 1, 1).clearContent();
}

/**
 * Columna de una cabecera según ESQUEMA, sin leer la hoja: escribirCabeceras_ ya ha garantizado que coinciden.
 * Antes se leía la fila de cabeceras en cada llamada, y cada lectura entre escrituras obliga a Sheets a esperar
 * al recálculo de toda la hoja (unas 60 veces por reparación).
 */
function colDe_(sh, h) { return ESQUEMA[sh.getName()].cabeceras.indexOf(h) + 1; }

/**
 * Aplica formato a una columna por nombre: n filas desde la siguiente a la cabecera. Una columna gris
 * (calculada) nunca debe llevar una validación manual, así que si no se le pasa una, se quita cualquier resto
 * de una versión anterior (p. ej. de un bug de columnas desalineadas que dejó una validación de otra columna
 * pegada aquí).
 */
function colFmt_(sh, h, n, o, filaCab) {
  filaCab = filaCab || 1;
  const c = colDe_(sh, h);
  if (!c) return;
  const r = sh.getRange(filaCab + 1, c, n, 1);
  if (o.fmt) r.setNumberFormat(o.fmt);
  if (o.gris) r.setBackground(COLORES.gris);
  if (o.ancho) ancho_(sh, c, o.ancho);
  if (o.validacion) r.setDataValidation(o.validacion);
  else if (o.gris) r.clearDataValidations();
}



function regla_(sh, a1, formula, color) {
  const f = loc_(formula);
  return { formula: f, regla: SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(f).setBackground(color).setRanges([sh.getRange(a1)]).build() };
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
  colFmt_(sh, 'Fecha escaneo', n, { fmt: FMT.fecha, ancho: 105 });
  colFmt_(sh, 'Fecha albarán', n, { fmt: FMT.fecha, ancho: 105 });
  colFmt_(sh, 'Mes', n, { gris: true, ancho: 70 });
  colFmt_(sh, 'Quincena', n, { gris: true, ancho: 75 });
  colFmt_(sh, 'Proveedor', n, { validacion: listaValidacion_(['RM', 'Otros']), ancho: 85 });
  colFmt_(sh, 'Nº albarán', n, { fmt: FMT.texto, ancho: 95 });
  colFmt_(sh, 'Nº trabajo', n, { ancho: 95 });
  colFmt_(sh, 'Matrícula', n, { validacion: matriculaValidacion_(), ancho: 100 });
  colFmt_(sh, 'Precio con IVA', n, { fmt: FMT.euro, ancho: 110 });
  colFmt_(sh, 'Precio - abonos', n, { fmt: FMT.euro, gris: true, ancho: 120 });
  colFmt_(sh, 'Coche', n, { gris: true, ancho: 190 });
  colFmt_(sh, 'Cliente', n, { gris: true, ancho: 150 });
  colFmt_(sh, 'Ver PDF', n, { ancho: 80 });
  colFmt_(sh, 'Avisos', n, { gris: true, ancho: 380 });
  colFmt_(sh, 'Nota escaneo', n, { ancho: 320 });
  const R = h => `${l[h]}2:${l[h]}${n + 1}`;
  ponerReglas_(sh, [
    regla_(sh, R('Avisos'), `=$${l['Avisos']}2<>""`, COLORES.naranja),
    regla_(sh, R('Nota escaneo'), `=$${l['Nota escaneo']}2<>""`, COLORES.naranja),
    regla_(sh, R('Coche'), `=LEFT($${l['Coche']}2,1)="⚠"`, COLORES.naranja),
  ]);
}

function formatoTrabajos_() {
  const sh = hoja_(HOJA.TRAB), n = ESQUEMA['Trabajos'].filasFormato, l = letras_(HOJA.TRAB), fc = ESQUEMA['Trabajos'].filaCabecera;
  const cf = (h, o) => colFmt_(sh, h, n, o, fc);
  cf('Nº trabajo', { ancho: 95 });
  cf('Fecha apertura', { fmt: FMT.fecha, ancho: 105 });
  cf('Mes', { gris: true, ancho: 70 });
  cf('Quincena', { validacion: listaValidacion_(['1', '2']), ancho: 80 });
  sh.getRange(fc + 1, colDe_(sh, 'Quincena'), n, 1).setHorizontalAlignment('center');
  cf('Matrícula', { validacion: matriculaValidacion_(), ancho: 100 });
  cf('Coche', { gris: true, ancho: 190 });
  cf('Cliente', { gris: true, ancho: 150 });
  cf('Recambios', { fmt: FMT.euro, gris: true, ancho: 110 });
  cf('Precio - abono', { fmt: FMT.euro, gris: true, ancho: 150 });
  cf('Recambios facturables RM', { fmt: FMT.euro, gris: true, ancho: 130 });
  cf('Recambios facturables Otros', { fmt: FMT.euro, gris: true, ancho: 130 });
  cf('Factura', { fmt: FMT.euro, ancho: 110 });
  cf('Beneficio', { fmt: FMT.euro, gris: true, ancho: 110 });
  cf('Pagado', { ancho: 80 });
  cf('Avisos', { gris: true, ancho: 340 });
  const R = h => `${l[h]}${fc + 1}:${l[h]}${fc + n}`;
  ponerReglas_(sh, [
    regla_(sh, R('Pagado'), `=AND($${l['Nº trabajo']}${fc + 1}<>"",$${l['Pagado']}${fc + 1}<>TRUE)`, COLORES.rojo),
    regla_(sh, R('Pagado'), `=$${l['Pagado']}${fc + 1}=TRUE`, COLORES.verde),
    regla_(sh, R('Avisos'), `=LEFT($${l['Avisos']}${fc + 1},1)="⚠"`, COLORES.naranja),
    regla_(sh, R('Avisos'), `=LEFT($${l['Avisos']}${fc + 1},1)="ℹ"`, COLORES.azul),
    regla_(sh, R('Coche'), `=LEFT($${l['Coche']}${fc + 1},1)="⚠"`, COLORES.naranja),
  ]);
  panelResumenTrabajos_(sh, n, l, fc);
}

/**
 * Panel "Resumen (según filtro)", encima de la cabecera real de Trabajos (ocupa las filas 1 a filaCabecera-1):
 * una cabecera por métrica y, debajo, su valor con SUBTOTAL, que suma sólo las filas que el filtro de la
 * cabecera deja visibles (sin filtro puesto, son todas). Morosos necesita además el truco SUMPRODUCT +
 * SUBTOTAL(103, OFFSET(...)) porque SUBTOTAL solo no admite una condición (Pagado = falso): se usa la columna
 * "Nº trabajo" para saber qué filas están visibles porque nunca está en blanco en una fila real, a diferencia
 * de "Factura" (un trabajo sin facturar todavía), que daría una visibilidad falsa (0) por error.
 */
function panelResumenTrabajos_(sh, n, l, fc) {
  const rango = h => `Trabajos!$${l[h]}$${fc + 1}:$${l[h]}$${fc + n}`;
  const primeraN = `Trabajos!$${l['Nº trabajo']}$${fc + 1}`, rangoN = rango('Nº trabajo');
  const visibles = `SUBTOTAL(103,OFFSET(${primeraN},ROW(${rangoN})-ROW(${primeraN}),0,1))`;
  const cols = [
    ['Trabajos', `=SUBTOTAL(103,${rangoN})`, '0'],
    ['Recambios RM', `=SUBTOTAL(109,${rango('Recambios facturables RM')})`, FMT.euro],
    ['Recambios Otros', `=SUBTOTAL(109,${rango('Recambios facturables Otros')})`, FMT.euro],
    ['Ingresos', `=SUBTOTAL(109,${rango('Factura')})`, FMT.euro],
    ['Morosos', `=SUMPRODUCT((${rango('Pagado')}=FALSE)*${visibles}*${rango('Factura')})`, FMT.euro],
    ['Beneficio', `=SUBTOTAL(109,${rango('Beneficio')})`, FMT.euro],
  ];
  // Las filas del panel (encima de la cabecera) son sólo del panel: se limpian ENTERAS antes de rehacerlo. Si se ha
  // insertado una columna (p. ej. Quincena), el título combinado ha crecido y combinar sólo una parte daría error.
  sh.getRange(1, 1, fc - 1, sh.getMaxColumns()).breakApart().clearContent().clearFormat().clearDataValidations();
  sh.getRange(1, 1, 1, cols.length).merge().setValue('Resumen (según filtro)')
    .setBackground(COLORES.cabecera).setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(2, 1, 1, cols.length).setValues([cols.map(c => c[0])])
    .setBackground(COLORES.cabecera).setFontColor('#ffffff').setFontWeight('bold').setWrap(true).setHorizontalAlignment('center');
  cols.forEach(([, formula, fmt], i) => sh.getRange(3, 1 + i).setValue(loc_(formula)).setNumberFormat(fmt));
  sh.getRange(3, 1, 1, cols.length).setHorizontalAlignment('center').setBackground(COLORES.gris);
}

function formatoPiezas_() {
  const sh = hoja_(HOJA.PIEZAS), n = ESQUEMA['Piezas'].filasFormato, l = letras_(HOJA.PIEZAS);
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
  ponerReglas_(sh, [
    regla_(sh, `A2:${ult}${n + 1}`, `=$${l['Reembolso']}2=TRUE`, COLORES.amarillo),
    regla_(sh, `${l['Avisos']}2:${l['Avisos']}${n + 1}`, `=$${l['Avisos']}2<>""`, COLORES.naranja),
  ]);
}

function formatoCoches_() {
  const sh = hoja_(HOJA.COCHES);
  ancho_(sh, 1, 110); ancho_(sh, 2, 220); ancho_(sh, 3, 240);
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
  ponerReglas_(sh, [regla_(sh, `${l['Estado']}2:${l['Estado']}${n + 1}`, `=LEFT($${l['Estado']}2,1)="⚠"`, COLORES.naranja)]);
}

function formatoLineas_() {
  const sh = hoja_(HOJA.LINEAS), n = ESQUEMA['Líneas RM'].filasFormato, l = letras_(HOJA.LINEAS);
  colFmt_(sh, 'Nº albarán', n, { fmt: FMT.texto, ancho: 95 });
  colFmt_(sh, 'Fecha albarán', n, { fmt: FMT.fecha, ancho: 105 });
  colFmt_(sh, 'Albarán origen', n, { fmt: FMT.texto, ancho: 105 });
  colFmt_(sh, 'Descripción', n, { ancho: 300 });
  colFmt_(sh, 'Descuento', n, { fmt: FMT.pct });
  colFmt_(sh, 'Importe sin IVA', n, { fmt: FMT.euro, ancho: 120 });
  colFmt_(sh, 'Conciliación', n, { gris: true, ancho: 190 });
  const ultCol = colLetra_(ESQUEMA['Líneas RM'].cabeceras.length);
  ponerReglas_(sh, [
    regla_(sh, `${l['Conciliación']}2:${l['Conciliación']}${n + 1}`, `=LEFT($${l['Conciliación']}2,1)="⚠"`, COLORES.naranja),
    regla_(sh, `A2:${ultCol}${n + 1}`, `=$${l['Tipo']}2="Abono"`, COLORES.amarillo),  // fila entera de las líneas de abono
  ]);
}

function formatoRegistro_() {
  const sh = hoja_(HOJA.REG), l = letras_(HOJA.REG);
  ancho_(sh, 1, 140); ancho_(sh, 2, 70); ancho_(sh, 3, 170); ancho_(sh, 4, 200); ancho_(sh, 5, 800);
  sh.getRange(2, 1, 3000, 1).setNumberFormat('dd/mm/yyyy hh:mm:ss');
  ponerReglas_(sh, [
    regla_(sh, 'A2:E3000', `=$${l['Nivel']}2="ERROR"`, COLORES.rojo),
    regla_(sh, 'A2:E3000', `=$${l['Nivel']}2="AVISO"`, COLORES.naranja),
  ]);
}

/** Pestaña Abonos: tabla grande arriba a la izquierda; resumen por quincena y panel "Pendientes de RM" a la derecha. */
function montarAbonos_() {
  recolocarTablaAbonos_();  // antes de escribir en posiciones fijas: si la tabla se ha desplazado, se pisarían datos
  montarResumenAbonos_();
  montarTablaAbonos_();
}

/** Última fila que cubren formatos y reglas de la tabla: crece con los datos. */
function finTablaAbonos_(sh) { return Math.max(ABONOS.filaTabla + ABONOS.maxTabla - 1, sh.getLastRow()); }

/**
 * Resumen por quincena + panel "Pendientes de RM", a la derecha de la tabla. Comparte filas con la tabla, así que si
 * alguien inserta o borra filas enteras se descoloca: alCambiar lo detecta (resumenAbonosEnSuSitio_) y lo vuelve a montar.
 * Las fórmulas usan columnas enteras de la tabla (E:E, D:D…): no les afecta insertar o borrar filas.
 */
function montarResumenAbonos_() {
  const sh = hoja_(HOJA.ABONOS), a = letras_(HOJA.ALB), f = letras_(HOJA.FACT);
  const cab = ABONOS.cabResumen, ini = ABONOS.filaIni, c0 = ABONOS.colResumen, pc = ABONOS.panelCol, pv = pc + 1;
  const L = i => colLetra_(c0 + i);  // 0 = Mes, 1 = Quincena, 2 = Recambios, 3 = Abonado, 4 = Total factura, 5 = Diferencia
  const anio = anioAbonos_(sh);
  // Las columnas de la derecha (O-W) son sólo del resumen y el panel: se limpian ENTERAS antes de reescribirlas, así
  // ninguna celda combinada queda a medias aunque se hayan insertado o borrado filas.
  sh.getRange(1, c0, sh.getMaxRows(), pv - c0 + 1).breakApart().clearContent().clearFormat().clearDataValidations();
  sh.getRange(1, c0).setValue('Año').setFontWeight('bold').setHorizontalAlignment('right');
  sh.getRange(ABONOS.celdaAnio).setValue(anio).setFontWeight('bold').setBackground(COLORES.amarillo).setNumberFormat('0');
  const A = `$${colLetra_(c0 + 1)}$1`;  // celda del año (ABONOS.celdaAnio, P1) en absoluto

  sh.getRange(ABONOS.filaCabResumen, c0, 1, cab.length).setValues([cab]).setBackground(COLORES.cabecera).setFontColor('#ffffff').setFontWeight('bold').setWrap(true).setVerticalAlignment('middle');
  const col = h => `$${h}:$${h}`, ESTADO_ = col('E'), CONIVA_ = col('D'), DIAS_ = col('L'), FECHA_ = col('A');
  for (let k = 0; k < ABONOS.filas; k++) {
    const r = ini + k, mes = Math.floor(k / 2) + 1, q = (k % 2) + 1;
    if (q === 1) sh.getRange(r, c0, 2, 1).merge().setValue(MESES[mes - 1]).setVerticalAlignment('middle').setHorizontalAlignment('center').setFontWeight('bold');
    const desde = `DATE(${A},${mes},${q === 1 ? 1 : 16})`, hasta = q === 1 ? `DATE(${A},${mes},15)` : `EOMONTH(DATE(${A},${mes},1),0)`;
    const rangoFecha = c => `${c},">="&${desde},${c},"<="&${hasta}`;
    const fA = `Albaranes!$${a['Fecha albarán']}:$${a['Fecha albarán']}`;
    const recambios = `=SUMIFS(Albaranes!$${a['Precio con IVA']}:$${a['Precio con IVA']},Albaranes!$${a['Proveedor']}:$${a['Proveedor']},"RM",${rangoFecha(fA)})`;
    const abonado = `=SUMIFS(${CONIVA_},${ESTADO_},"Abonada",${rangoFecha(FECHA_)})+SUMIFS(${CONIVA_},${ESTADO_},"Sin solicitar",${rangoFecha(FECHA_)})`;
    const crit = `'Facturas RM'!$${f['Año']}:$${f['Año']},${A},'Facturas RM'!$${f['Mes']}:$${f['Mes']},${mes},'Facturas RM'!$${f['Quincena']}:$${f['Quincena']},${q}`;
    const totalFactura = `=IF(COUNTIFS(${crit})=0,"",SUMIFS('Facturas RM'!$${f['Total']}:$${f['Total']},${crit}))`;
    // Diferencia informativa, sin aviso: los abonos de RM suelen llegar en la factura siguiente.
    const diferencia = `=IF(${L(4)}${r}="","",ROUND(${L(4)}${r}-(${L(2)}${r}-${L(3)}${r}),2))`;
    sh.getRange(r, c0 + 1, 1, 5).setValues([locFila_([q, recambios, abonado, totalFactura, diferencia])]);
  }
  sh.getRange(ini, c0 + 1, ABONOS.filas, 1).setHorizontalAlignment('center');
  sh.getRange(ini, c0 + 2, ABONOS.filas, 4).setNumberFormat(FMT.euro).setBackground(COLORES.gris);
  [60, 75, 120, 120, 120, 100].forEach((w, i) => ancho_(sh, c0 + i, w));
  ancho_(sh, c0 - 1, 20);  // separación con la tabla

  // ---- Panel "Pendientes de RM": piezas 'Sin abonar' de toda la tabla, no atadas a la quincena en que se pidieron ----
  // El umbral se copia a una celda de ESTA pestaña: una regla de formato condicional no puede leer Config.
  const UMBRAL_ = `$${colLetra_(pv)}$4`;
  sh.getRange(1, pc, 1, 2).merge().setValue('Pendientes de RM').setBackground(COLORES.cabecera).setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(2, pc, 5, 2).setValues([
    ['Piezas sin abonar', loc_(`=COUNTIF(${ESTADO_},"Sin abonar")`)],
    ['Importe pendiente', loc_(`=SUMIF(${ESTADO_},"Sin abonar",${CONIVA_})`)],
    ['Aviso a partir de (días)', '=DIAS_AVISO_REEMB'],
    ['Fuera de plazo', loc_(`=COUNTIFS(${ESTADO_},"Sin abonar",${DIAS_},">"&${UMBRAL_})`)],
    ['Más antigua (días)', loc_(`=IFERROR(MAXIFS(${DIAS_},${ESTADO_},"Sin abonar"),0)`)],
  ]);
  sh.getRange(3, pv).setNumberFormat(FMT.euro);
  sh.getRange(2, pc, 5, 1).setFontWeight('bold');
  sh.getRange(2, pv, 5, 1).setHorizontalAlignment('center');
  ancho_(sh, pc - 1, 20); ancho_(sh, pc, 170); ancho_(sh, pv, 90);
  reglasAbonos_(sh);
}

/** Año del resumen de Abonos: el de su celda, o el que haya junto a la etiqueta "Año" si el bloque se ha desplazado. */
function anioAbonos_(sh) {
  const c0 = ABONOS.colResumen, v = sh.getRange(1, c0, 12, 2).getValues();
  const fila = v.find(r => r[0] === 'Año' && r[1] !== '');
  const x = fila ? fila[1] : sh.getRange(ABONOS.celdaAnio).getValue();
  return x === '' || x == null ? 2026 : x;
}

/** ¿Sigue el resumen de la derecha en su sitio? (filas insertadas o borradas enteras lo desplazan). */
function resumenAbonosEnSuSitio_(sh) {
  const n = ABONOS.filas, v = sh.getRange(1, ABONOS.colResumen, ABONOS.filaIni + n, 2).getValues();
  if (v[0][0] !== 'Año' || v[ABONOS.filaCabResumen - 1][0] !== ABONOS.cabResumen[0]) return false;
  for (let k = 0; k < n; k++) if (Number(v[ABONOS.filaIni - 1 + k][1]) !== (k % 2) + 1) return false;
  return v[ABONOS.filaIni - 1 + n][1] === '';
}

/** Tabla grande: cabecera en la fila 1 (fija), formatos, desplegable de Estado, columna Clave oculta y filtro. */
function montarTablaAbonos_() {
  const sh = hoja_(HOJA.ABONOS), tb = ABONOS.filaTabla, ct = ABONOS.cabTabla, fin = finTablaAbonos_(sh), maxRows = sh.getMaxRows();
  if (maxRows < fin) sh.insertRowsAfter(maxRows, fin - maxRows);
  sh.getRange(ABONOS.filaCabTabla, 1, 1, ct.length).setValues([ct]).setBackground(COLORES.cabecera).setFontColor('#ffffff').setFontWeight('bold').setWrap(true).setVerticalAlignment('middle');
  const n = fin - tb + 1;
  sh.getRange(tb, 1, n, ct.length).clearDataValidations();
  sh.getRange(tb, 1, n, 1).setNumberFormat(FMT.fecha);
  sh.getRange(tb, 3, n, 2).setNumberFormat(FMT.euro);
  sh.getRange(tb, 5, n, 1).setDataValidation(listaValidacion_(ESTADOS_ABONO));
  sh.getRange(tb, 6, n, 1).setNumberFormat(FMT.texto);
  sh.getRange(tb, 9, n, 1).setNumberFormat(FMT.fecha);
  sh.getRange(tb, 12, n, 1).setNumberFormat('0');
  [90, 260, 130, 130, 130, 150, 110, 110, 120, 110, 260, 100].forEach((w, i) => ancho_(sh, i + 1, w));
  sh.hideColumns(ct.length);  // "Clave": une cada fila con su pieza / línea de abono; no se toca a mano
  // Filtro que incluye la columna oculta: ordenar con él mueve la clave junto con su fila.
  if (!sh.getFilter()) sh.getRange(ABONOS.filaCabTabla, 1, fin - ABONOS.filaCabTabla + 1, ct.length).createFilter();
  const shC = ss_().getSheetByName(HOJA.CLAVES);
  if (shC) shC.hideSheet();
  fijarFilas_(sh, ABONOS.filaCabTabla);
  reglasAbonos_(sh);
}

function reglasAbonos_(sh) {
  const tb = ABONOS.filaTabla, fin = finTablaAbonos_(sh), pc = ABONOS.panelCol, pv = pc + 1, UMBRAL_ = `$${colLetra_(pv)}$4`;
  ponerReglas_(sh, [
    regla_(sh, `${colLetra_(pc)}5:${colLetra_(pv)}5`, `=${colLetra_(pv)}5>0`, COLORES.naranja),
    regla_(sh, `A${tb}:L${fin}`, `=AND($E${tb}="Sin abonar",$L${tb}>${UMBRAL_})`, COLORES.naranja),
    regla_(sh, `A${tb}:L${fin}`, `=$E${tb}="Abonada"`, COLORES.verde),
    regla_(sh, `A${tb}:L${fin}`, `=$E${tb}="Sin abonar"`, COLORES.rojo),
    regla_(sh, `A${tb}:L${fin}`, `=$E${tb}="Sin solicitar"`, COLORES.amarillo),
  ]);
}

function instalarTriggers_() {
  const ss = ss_();
  ScriptApp.getProjectTriggers().forEach(t => { if (['alEditar', 'alAbrir', 'alCambiar'].indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('alEditar').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('alCambiar').forSpreadsheet(ss).onChange().create();  // borrar filas en Abonos desmarca la pieza
  ScriptApp.newTrigger('alAbrir').forSpreadsheet(ss).onOpen().create();
}
