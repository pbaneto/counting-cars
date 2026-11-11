const test = require('node:test');
const assert = require('node:assert/strict');
const { crearEntorno, Filtro } = require('./gas-mock');

const ESQ = (() => { const c = {}; require('vm').createContext(c); require('vm').runInContext(require('fs').readFileSync(require('path').join(__dirname, '..', 'src', 'Config.js'), 'utf8') + ';this.ESQUEMA = ESQUEMA;', c); return c.ESQUEMA; })();
/** Nº de columna por cabecera (Albaranes / Trabajos), para no depender del orden de las columnas. */
const A = h => ESQ['Albaranes'].cabeceras.indexOf(h) + 1, T = h => ESQ['Trabajos'].cabeceras.indexOf(h) + 1;
const letra = n => String.fromCharCode(64 + n);

const PRIV = require('path').join(__dirname, 'fixtures-private.js');
const ESTADOS = ['Abonada', 'Sin abonar', 'Sin solicitar'];
require('fs').writeFileSync(PRIV, `const PRIVATE = { SPREADSHEET_ID: 'SS',
  CARPETAS: { CARPETA_ENTRADA: 'ENT', CARPETA_PROCESADOS: 'PROC', CARPETA_FACTURAS_RM: 'FRA', CARPETA_FACTURAS_RM_PROCESADAS: 'FRAP' },
  COCHES: [['1234ABC','Cli A','Peugeot 208'],['4321GHJ','Cli B','Mini'],['5678DEF','Cli C','Peugeot']],
  PILOTO: { fecha: '2026-09-15', albaranes: [{proveedor:'RM',matricula:'4321GHJ',importe:100},{proveedor:'RM',matricula:'4321GHJ',importe:50},{proveedor:'Otros',matricula:'5678DEF',importe:30}],
            trabajos: [{matricula:'4321GHJ',pagado:true,factura:500}] },
  RESUMEN: { anio: 2026, mesEnVivo: 9, y2026: Array.from({length:12},()=>[1,2,3,4]), y2025: Array.from({length:12},()=>[1,2,3,4,5,6]), y2024: Array.from({length:12},()=>[1,2,3,4]),
    fijos: { banco: [['A',1]], gastos: [['B',2]] } } };`);

const albaranRaw = (extra) => Object.assign({ es_albaran: true, proveedor: 'RM', numero_albaran: '462446', fecha: '2026-09-15', matricula: '1234ABC',
  base_imponible: 36.81, iva_importe: 7.73, total: 44.54,
  lineas: [
    { referencia: 'DAYCO6PK1090EE', descripcion: 'CORREA ESTRIADA', marca: 'DAYCO', cantidad: 1, precio_unitario: 28.27, descuento_pct: 60, importe: 11.31, reembolso: false },
    { referencia: 'KRAFF47054', descripcion: 'HIDROIL 775', marca: 'KRAFFT', cantidad: 2, precio_unitario: 23.10, descuento_pct: 45, importe: 25.41, reembolso: false },
    { referencia: '', descripcion: 'SIGAUS (SIG. RD 679/2006)', marca: '', cantidad: 2, precio_unitario: 0.05, descuento_pct: 0, importe: 0.09, reembolso: false },
  ] }, extra);

/** Gemini simulado: la respuesta depende del nombre del PDF que "envía" cada petición (se guarda en el payload por el mock de Blob). */
function entorno(respuestas) {
  let e;
  e = crearEntorno({ privado: true, privadoRuta: PRIV, gemini: req => {
    const parts = JSON.parse(req.payload).contents[0].parts;
    const nombre = parts[0].inline_data.data !== 'AAAA' ? parts[0].inline_data.data : e.actual.shift();
    e.peticiones.push({ nombre, pista: parts[2] ? parts[2].text : '' });
    // Una lista de respuestas = una por lectura (la segunda lectura de un albarán que no cuadra recibe la siguiente)
    const r = Array.isArray(respuestas[nombre]) ? respuestas[nombre].shift() : respuestas[nombre];
    return { code: 200, body: JSON.stringify({ candidates: [{ content: { parts: [{ text: JSON.stringify(r) }] }, finishReason: 'STOP' }] }) };
  } });
  e.actual = []; e.peticiones = [];
  e.run("PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY','k'.repeat(30))");
  e.run('setup()');
  e.run('cargarDatosIniciales()');
  return e;
}

let contadorId = 0;
const esFecha = x => Object.prototype.toString.call(x) === '[object Date]';
function subir(e, nombres, carpeta) {
  e.actual.push(...nombres);
  nombres.forEach(n => e.mkFile('id' + (++contadorId), n, 'application/pdf', carpeta));
}
const tabla = (e, hoja) => e.run(`(() => { const t = leerTabla_('${hoja}'); return t.filas.map(f => f.v); })()`);
const plano = x => JSON.parse(JSON.stringify(x));  // de los objetos del simulador (otro contexto de vm) a objetos normales

test('setup crea pestañas, cabeceras, fórmulas y configuración', () => {
  const e = entorno({});
  const nombres = e.ss.getSheets().map(s => s.name);
  for (const n of ['Albaranes', 'Trabajos', 'Piezas', 'Abonos', 'Cuadre RM', 'Coches', 'Resumen', 'Facturas RM', 'Líneas RM', 'Config', 'Registro']) assert.ok(nombres.includes(n), n);
  assert.ok(!nombres.includes('Hoja 1'));
  const alb = e.ss.getSheetByName('Albaranes');
  assert.equal(alb.valor(1, A('Nº albarán')), 'Nº albarán');
  assert.equal(alb.valor(1, A('Mes')), 'Mes', 'columna nueva "Mes" junto a "Fecha albarán"');
  assert.match(alb.cell(2, A('Mes')).f, /CHOOSE\(MONTH\(/);
  assert.match(alb.cell(2, A('Quincena')).f, new RegExp(`^=IF\\(${letra(A('Fecha albarán'))}2=""`));          // Quincena
  assert.match(alb.cell(2, A('Precio - abonos')).f, /SUMIFS\(Piezas!/);       // Precio - abonos
  assert.match(alb.cell(2, A('Avisos')).f, /Nº de albarán duplicado/);           // Avisos en las filas con datos (piloto: filas 2-4)
  assert.equal(alb.cell(5, A('Avisos')), undefined, 'sin fórmulas en filas vacías: leerlas y recalcularlas era lo lento');
  assert.equal(alb.getLastRow(), 4);
  // Casillas sólo en filas con datos: una casilla vacía vale FALSE y haría leer miles de filas
  const trab = e.ss.getSheetByName('Trabajos');
  assert.equal(trab.getLastRow(), 6, 'Trabajos: panel (filas 1-3) + cabecera (fila 4) + 2 trabajos del piloto (5-6)');
  assert.equal(e.ss.getSheetByName('Piezas').getLastRow(), 1, 'Piezas vacía: sin casillas "Reembolso" por debajo');
  assert.equal(trab.valor(6, T('Pagado')), false, 'la fila nueva lleva su casilla (FALSE = sin pagar), columna M');
  // Cabecera real en la fila 4 (el panel ocupa la 1-3), datos desde la 5
  assert.equal(trab.valor(4, T('Mes')), 'Mes', 'Mes junto a Nº trabajo');
  assert.equal(trab.valor(4, T('Quincena')), 'Quincena', 'Quincena a la derecha de Mes');
  assert.equal(trab.valor(5, T('Mes')), 'sep', 'rellenado con la fecha del piloto (15/09)');
  assert.equal(trab.valor(5, T('Quincena')), 1, 'rellenada con la fecha del piloto (15/09 → 1ª quincena)');
  assert.equal(trab.cell(5, T('Mes')).f, '', 'Mes es un valor, no una fórmula: se puede cambiar a mano');
  assert.equal(trab.cell(5, T('Quincena')).f, '', 'Quincena es un valor, no una fórmula: se puede cambiar a mano');
  assert.equal(trab.valor(4, T('Matrícula')), 'Matrícula');
  assert.match(trab.cell(5, T('Recambios facturables RM')).f, /SUMIFS\(Albaranes!.*"RM"\)/);      // Recambios facturables RM
  assert.match(trab.cell(5, T('Recambios facturables Otros')).f, /SUMIFS\(Albaranes!.*"Otros"\)/);  // Recambios facturables Otros
  // Panel "Resumen (según filtro)" encima de la cabecera: título (1), etiquetas (2), valores (3)
  assert.equal(trab.valor(1, 1), 'Resumen (según filtro)');
  assert.equal(trab.valor(2, 1), 'Trabajos'); assert.equal(trab.valor(2, 5), 'Morosos');
  assert.match(trab.cell(3, 5).f, /SUMPRODUCT/);
  // Ninguna columna calculada (gris) se queda con una validación residual; las de entrada no se tocan
  assert.deepEqual(new Set(trab.validacionesLimpiadas), new Set([1, ...['Coche', 'Cliente', 'Recambios', 'Precio - abono', 'Recambios facturables RM', 'Recambios facturables Otros', 'Beneficio', 'Avisos'].map(T)]),
    'Coche, Cliente, Recambios, Recambios facturables, Recambios facturables RM/Otros, Beneficio, Avisos y la fila de valores del panel');
  const cfg = tabla(e, 'Config');
  assert.equal(cfg.find(x => x.Clave === 'CARPETA_ENTRADA').Valor, 'ENT');
  assert.equal(cfg.find(x => x.Clave === 'MODELO_GEMINI').Valor, 'gemini-3.5-flash-lite');
  assert.equal(e.ss.getSheetByName('Abonos').valor(1, 1), 'Fecha abono', 'la tabla arriba del todo');
  const cuadre = e.ss.getSheetByName('Cuadre RM');
  assert.deepEqual([cuadre.valor(1, 1), cuadre.valor(1, 2), cuadre.valor(3, 1)], ['Año', 2026, 'Mes']);
  // Fechas de cada quincena dentro de la fórmula (sin columnas Desde/Hasta): ene 1ª = 1-15, ene 2ª = 16-fin de mes
  assert.match(cuadre.cell(4, 3).f, /SUMIFS\(Albaranes!.*">="&DATE\(\$B\$1,1,1\).*"<="&DATE\(\$B\$1,1,15\)/);
  assert.match(cuadre.cell(5, 4).f, /SUMIFS\(Abonos!\$D:\$D,Abonos!\$E:\$E,"Abonada",Abonos!\$A:\$A,">="&DATE\(\$B\$1,1,16\).*"<="&EOMONTH\(DATE\(\$B\$1,1,1\),0\)/);
  assert.match(cuadre.cell(4, 6).f, /^=IF\(E4="","",ROUND\(E4-\(C4-D4\),2\)\)$/); // Diferencia informativa, sin aviso
  assert.equal(cuadre.valor(27, 2), 2, 'dic, 2ª quincena, en la fila 27');
  assert.equal(cuadre.cell(4, 9).f, '=DIAS_AVISO_REEMB');
  assert.match(cuadre.cell(5, 9).f, /COUNTIFS\(Abonos!.*">"&I4\)/);
  e.run('repararFormulas()');
  const errores = e.log.console.filter(m => /^\[ERROR\]/.test(m));
  assert.deepEqual(errores, [], 'setup y repararFormulas deben terminar sin errores');
});

test('repararFormulas quita las fórmulas de las filas vacías (hoja antigua) sin tocar datos ni valores a mano', () => {
  const e = entorno({});
  const alb = e.ss.getSheetByName('Albaranes'), p = e.ss.getSheetByName('Piezas');
  // Como la hoja real antes del cambio: fórmulas rellenadas por adelantado muy por debajo de los datos
  for (let r = 5; r <= 1501; r++) { alb.put(r, A('Quincena'), `=IF(B${r}="","",1)`); alb.put(r, A('Avisos'), `=IF(G${r}="","",1)`); }
  p.put(4001, 15, '=IF(C4001="","",1)');
  for (let r = 2; r <= 4001; r++) p.put(r, 1, false);   // casillas "Reembolso" rellenadas hasta la 4001 (valen FALSE)
  const trab = e.ss.getSheetByName('Trabajos');
  // El panel ocupa 1-3, la cabecera la 4, los 2 trabajos del piloto la 5-6: el relleno de sobra empieza en la 7
  for (let r = 7; r <= 801; r++) trab.put(r, T('Pagado'), false); // "Pagado" (columna M) hasta la 801
  alb.put(900, A('Coche'), 'escrito a mano');                 // valor (no fórmula) en la columna calculada "Coche"
  e.run('repararFormulas()');
  assert.equal(alb.cell(700, A('Quincena')), undefined, 'Quincena vacía por debajo de los datos');
  assert.equal(alb.cell(1501, A('Avisos')), undefined, 'Avisos vacía por debajo de los datos');
  assert.equal(p.cell(4001, 15), undefined);
  assert.match(alb.cell(4, A('Avisos')).f, /Nº de albarán duplicado/, 'las filas con datos conservan sus fórmulas');
  assert.equal(alb.valor(900, A('Coche')), 'escrito a mano', 'no borra un valor escrito a mano');
  assert.ok(e.log.console.some(l => new RegExp(`\\[AVISO\\] repararFormulas Albaranes!${letra(A('Coche'))}900`).test(l)));
  assert.equal(p.getLastRow(), 1);
  assert.equal(trab.getLastRow(), 6, 'Trabajos queda con sus 2 filas de datos (panel 1-3, cabecera 4, datos 5-6)');
  assert.equal(trab.valor(5, T('Pagado')), true, 'el Pagado de las filas con datos no se toca');
});

test('repararFormulas conserva el año elegido en Cuadre RM', () => {
  const e = entorno({});
  const c = e.ss.getSheetByName('Cuadre RM');
  c.put(1, 2, 2027);
  e.run('repararFormulas()');
  assert.equal(c.valor(1, 2), 2027);
  assert.match(c.cell(4, 3).f, /\$B\$1/);
});

test('repararFormulas sólo reescribe las fórmulas que han cambiado', () => {
  const e = entorno({});
  const trab = e.ss.getSheetByName('Trabajos');
  const escritas = [];
  e.ss.sheets.forEach(sh => {
    const put = sh.put.bind(sh);
    sh.put = (r, c, v) => { if (typeof v === 'string' && v.charAt(0) === '=') escritas.push(`${sh.name}!${r},${c}`); return put(r, c, v); };
  });
  const tablas = ['Albaranes!', 'Trabajos!', 'Piezas!', 'Líneas RM!'];
  const enTablas = () => escritas.filter(x => tablas.some(t => x.startsWith(t)) && !x.startsWith('Trabajos!3,'));  // fila 3 = panel
  e.run('repararFormulas()');
  assert.deepEqual(enTablas(), [], 'sin cambios: ninguna fórmula de las tablas se reescribe');
  const buena = trab.cell(5, T('Beneficio')).f;
  trab.put(5, T('Beneficio'), '=1');  // alguien ha pisado la fórmula de Beneficio
  escritas.length = 0;
  e.run('repararFormulas()');
  assert.equal(trab.cell(5, T('Beneficio')).f, buena, 'la fórmula pisada se restaura');
  assert.ok(enTablas().every(x => x.startsWith('Trabajos!') && x.endsWith(',' + T('Beneficio'))), 'sólo se reescribe esa columna');
});

test('repararFormulas añade a Config las claves nuevas sin pisar los valores existentes', () => {
  const e = entorno({});
  const cfg = e.ss.getSheetByName('Config');
  const fila = clave => e.run(`leerTabla_('Config').filas.find(f => f.v['Clave'] === '${clave}').fila`);
  const r = fila('DIAS_AVISO_REEMBOLSO');
  [1, 2, 3].forEach(c => cfg.put(r, c, ''));            // hoja montada antes de existir la clave
  cfg.put(fila('MAX_ARCHIVOS'), 2, 20);                 // valor cambiado a mano por el usuario
  e.run('repararFormulas()');
  const t = tabla(e, 'Config');
  assert.equal(t.find(x => x.Clave === 'DIAS_AVISO_REEMBOLSO').Valor, 45);
  assert.equal(t.find(x => x.Clave === 'MAX_ARCHIVOS').Valor, 20);
});

test('cargarDatosIniciales: coches y piloto (un trabajo por matrícula) y no pisa datos', () => {
  const e = entorno({});
  assert.equal(tabla(e, 'Coches').length, 3);
  const trab = tabla(e, 'Trabajos'), alb = tabla(e, 'Albaranes');
  assert.equal(alb.length, 3);
  assert.deepEqual(Array.from(trab.map(t => t['Nº trabajo'])).sort(), ['1GHJ-1', '8DEF-1']);
  assert.ok(trab.some(t => t['Nº trabajo'] === '1GHJ-1' && t['Pagado'] === true && t['Factura'] === 500));
  assert.ok(alb.every(a => a['Nº trabajo']));
  const antes = e.log.alerts.length;
  e.run('cargarDatosIniciales()');
  assert.equal(tabla(e, 'Albaranes').length, 3);
  assert.ok(e.log.alerts.length > antes);
});

test('procesarAlbaranes: nuevo, segundo escaneo con R, duplicado sin R y no-albarán', () => {
  const R = {
    'a1.pdf': albaranRaw(),
    'a2.pdf': albaranRaw({ lineas: albaranRaw().lineas.map((l, i) => Object.assign({}, l, { reembolso: i === 0, manuscrito: i === 0 ? 'R' : '' })) }),
    'a3.pdf': albaranRaw(),
    'a4.pdf': { es_albaran: false, proveedor: 'Otros', numero_albaran: '', fecha: '', matricula: '', total: null, lineas: [] },
  };
  const e = entorno(R);
  e.mkFolder('ENT', 'Entrada'); // ya creada por el mock de setup? no: las carpetas de Drive son del mock
  e.mkFolder('PROC', 'Procesados'); e.mkFolder('FRA', 'Fras'); e.mkFolder('FRAP', 'FrasP');
  subir(e, ['a1.pdf', 'a2.pdf', 'a3.pdf', 'a4.pdf'], 'ENT');
  e.run('procesarAlbaranes()');

  const alb = tabla(e, 'Albaranes').filter(a => a['Nº albarán'] === '462446');
  assert.equal(alb.length, 1, 'un solo albarán aunque se escanee 3 veces');
  assert.equal(alb[0]['Nº trabajo'], '4ABC-1');
  assert.equal(alb[0]['Precio con IVA'], 44.54);
  assert.match(alb[0]['Nota escaneo'], /Reembolso 21\/09|Reembolso \d\d\/\d\d/);
  const piezas = tabla(e, 'Piezas').filter(p => p['Nº albarán'] === '462446');
  assert.equal(piezas.length, 2, 'la línea SIGAUS no es una pieza');
  assert.equal(piezas.filter(p => p['Reembolso'] === true).length, 1);
  assert.equal(piezas.find(p => p['Reembolso'] === true)['Referencia pieza'], 'DAYCO6PK1090EE');
  assert.ok(piezas.every(p => p['Origen'] === 'Escaneo'));
  // Abonos: la pieza reembolsada aparece como "Sin abonar"
  const ab = e.ss.getSheetByName('Abonos');
  assert.equal(ab.valor(2, 5), 'Sin abonar');
  assert.equal(ab.valor(2, 3), 11.31);
  assert.equal(ab.cell(2, 4).f, '=ROUND($C2*(1+IVA),2)');  // Precio con IVA: fórmula con el IVA de Config
  assert.match(ab.cell(2, 12).f, /TODAY\(\)/);  // Días pendiente: fórmula viva, no un valor fijo
  // El nº de albarán de Abonos y de Piezas es un enlace a su PDF
  const pdf = e.ss.getSheetByName('Albaranes').cell(2 + tabla(e, 'Albaranes').findIndex(x => x['Nº albarán'] === '462446'), 13).f.match(/"(https:[^"]+)"/)[1];
  assert.equal(ab.cell(2, 6).url, pdf);
  assert.ok([2, 3].every(r => e.ss.getSheetByName('Piezas').cell(r, 3).url === pdf));
  // Drive: 3 en Procesados, el no-albarán en Errores
  assert.equal(e.carpetas.PROC.ficheros.length, 3);
  assert.equal(e.carpetas.ENT.ficheros.length, 0);
  const errores = Object.values(e.carpetas).find(f => f.nombre === 'Errores');
  assert.equal(errores.ficheros.length, 1);
  const reg = tabla(e, 'Registro').map(r => r['Mensaje']).join('\n');
  assert.match(reg, /duplicado/);
  assert.match(reg, /no parece un albarán/);
});

test('procesarAlbaranes: el mismo albarán con el nº mal leído no se duplica y su R marca la pieza del correcto', () => {
  const conR = albaranRaw({ numero_albaran: '462846', lineas: albaranRaw().lineas.map((l, i) => Object.assign({}, l, { reembolso: i === 0, manuscrito: i === 0 ? 'R' : '' })) });
  const e = entorno({ 'a1.pdf': albaranRaw(), 'a2.pdf': albaranRaw({ numero_albaran: '462846' }), 'a3.pdf': conR, 'a4.pdf': albaranRaw({ numero_albaran: '462999' }) });
  ['ENT', 'PROC', 'FRA', 'FRAP'].forEach(id => e.mkFolder(id, id));
  subir(e, ['a1.pdf', 'a2.pdf', 'a3.pdf', 'a4.pdf'], 'ENT');
  e.run('procesarAlbaranes()');

  const nums = tabla(e, 'Albaranes').map(a => a['Nº albarán']).filter(n => /^462/.test(n));
  assert.deepEqual([...nums], ['462446', '462999'], 'con dos dígitos distintos es otro albarán');
  const piezas = tabla(e, 'Piezas');
  assert.equal(piezas.filter(p => p['Nº albarán'] === '462846').length, 0);
  assert.equal(piezas.filter(p => p['Nº albarán'] === '462446' && p['Reembolso'] === true).length, 1);
  assert.equal(e.carpetas.ENT.ficheros.length, 0);
  assert.match(tabla(e, 'Registro').map(r => r['Mensaje']).join('\n'), /Parece el albarán 462446 con el nº mal leído \(462846\)/);
});

test('procesarAlbaranes: albarán sin total se añade con la suma de las líneas y un aviso; escanearlo otra vez no lo duplica', () => {
  const sinTotal = albaranRaw({ total: null, base_imponible: null, iva_importe: null });
  const e = entorno({ 'a1.pdf': sinTotal, 'a2.pdf': sinTotal });
  ['ENT', 'PROC', 'FRA', 'FRAP'].forEach(id => e.mkFolder(id, id));
  subir(e, ['a1.pdf', 'a2.pdf'], 'ENT');
  e.run('procesarAlbaranes()');

  const alb = tabla(e, 'Albaranes').filter(a => a['Nº albarán'] === '462446');
  assert.equal(alb.length, 1);
  assert.equal(alb[0]['Precio con IVA'], 44.54);
  assert.match(alb[0]['Nota escaneo'], /Sin total en el albarán/);
  assert.equal(tabla(e, 'Piezas').filter(p => p['Nº albarán'] === '462446').length, 2);
  assert.equal(e.carpetas.PROC.ficheros.length, 2);
  assert.equal(e.carpetas.ENT.ficheros.length, 0);
});

test('procesarAlbaranes: la 2ª hoja completa el albarán aunque llegue antes y repetida; si no cuadra se queda en Entrada', () => {
  const [correa, aceite, sigaus] = albaranRaw().lineas;
  const hoja1 = extra => albaranRaw(Object.assign({ total: null, base_imponible: null, iva_importe: null, lineas: [correa] }, extra));
  const hoja2 = extra => albaranRaw(Object.assign({ lineas: [aceite, sigaus] }, extra));
  const e = entorno({
    'h2.pdf': hoja2(), 'h1.pdf': hoja1(), 'h1-otra.pdf': hoja1(), 'h2-otra.pdf': hoja2(),
    'mal1.pdf': hoja1({ numero_albaran: '999999' }), 'mal2.pdf': hoja2({ numero_albaran: '999999', total: 500, base_imponible: null }),
  });
  ['ENT', 'PROC', 'FRA', 'FRAP'].forEach(id => e.mkFolder(id, id));
  subir(e, ['h2.pdf', 'h1.pdf', 'h1-otra.pdf', 'h2-otra.pdf', 'mal1.pdf', 'mal2.pdf'], 'ENT');
  e.run('procesarAlbaranes()');

  const alb = tabla(e, 'Albaranes').filter(a => a['Nº albarán'] === '462446');
  assert.equal(alb.length, 1);
  assert.equal(alb[0]['Precio con IVA'], 44.54);
  assert.equal(alb[0]['Nota escaneo'], '');
  assert.equal(tabla(e, 'Piezas').filter(p => p['Nº albarán'] === '462446').map(p => p['Referencia pieza']).join(), 'DAYCO6PK1090EE,KRAFF47054');

  const mal = tabla(e, 'Albaranes').find(a => a['Nº albarán'] === '999999');
  assert.match(mal['Nota escaneo'], /Sin total en el albarán/);
  assert.equal(mal['Precio con IVA'], 13.69);
  assert.equal(e.carpetas.ENT.ficheros.map(f => f.getName()).join(), 'mal2.pdf');
  assert.equal(e.carpetas.PROC.ficheros.length, 5);
});

test('albarán de dos hojas: «Ver PDF» enlaza las dos, en el mismo lote o en otro, y un reescaneo de cualquiera se reconoce', () => {
  const [correa, aceite, sigaus] = albaranRaw().lineas;
  const hoja1 = albaranRaw({ total: null, base_imponible: null, iva_importe: null, lineas: [correa] });
  const hoja2 = albaranRaw({ lineas: [aceite, sigaus] });
  const e = entorno({ 'h1.pdf': hoja1, 'h2.pdf': hoja2, 'n.pdf': albaranRaw({ numero_albaran: '777777', matricula: '5678DEF' }) });
  ['ENT', 'PROC', 'FRA', 'FRAP'].forEach(id => e.mkFolder(id, id));
  subir(e, ['h1.pdf', 'n.pdf'], 'ENT');
  e.run('procesarAlbaranes()');
  subir(e, ['h2.pdf'], 'ENT');
  e.run('procesarAlbaranes()');

  const id = n => e.carpetas.PROC.ficheros.find(f => f.nombre === n).getId(), url = n => `https://drive.google.com/file/d/${id(n)}/view`;
  const sh = e.ss.getSheetByName('Albaranes'), fila = 2 + tabla(e, 'Albaranes').findIndex(a => a['Nº albarán'] === '462446');
  const celda = sh.cell(fila, 13);
  assert.equal(celda.v, 'Hoja 1 · Hoja 2');
  assert.equal(celda.f, '', 'la fórmula de la 1ª hoja desaparece');
  assert.deepEqual(plano(celda.enlaces), [{ texto: 'Hoja 1', url: url('h1.pdf') }, { texto: 'Hoja 2', url: url('h2.pdf') }]);
  assert.match(sh.cell(2 + tabla(e, 'Albaranes').findIndex(a => a['Nº albarán'] === '777777'), 13).f, /HYPERLINK/, 'el de una hoja sigue igual');
  assert.ok(e.run('contextoAlbaranes_().idsPdf').has(id('h2.pdf')));
  assert.ok(e.run('contextoAlbaranes_().idsPdf').has(id('h1.pdf')));

  e.run('actualizarAbonos()');
  assert.equal(e.ss.getSheetByName('Piezas').cell(2 + tabla(e, 'Piezas').findIndex(p => p['Nº albarán'] === '462446'), 3).url, url('h1.pdf'), 'Piezas enlaza la 1ª hoja');
});

test('procesarAlbaranes: R con la referencia mal leída marca la pieza del albarán en vez de crear otra; si hay duda, no marca ninguna', () => {
  const [correa, aceite, sigaus] = albaranRaw().lineas;
  const dos = { lineas: [correa, Object.assign({}, correa, { referencia: 'DAYCO6PK1090FF' }), aceite, sigaus], base_imponible: 48.12, iva_importe: 10.11, total: 58.23 };
  const conR = (referencia, numero_albaran) => albaranRaw(Object.assign({}, dos, { numero_albaran, lineas: [Object.assign({}, correa, { referencia, reembolso: true, manuscrito: 'R' }), dos.lineas[1], aceite, sigaus] }));
  const e = entorno({
    'a1.pdf': albaranRaw({ numero_albaran: '100001', lineas: [correa, aceite, sigaus] }),
    'a2.pdf': albaranRaw({ numero_albaran: '100001', lineas: [Object.assign({}, aceite, { referencia: 'KRAFFO7054', reembolso: true, manuscrito: 'R' }), correa, sigaus] }),
    'b1.pdf': albaranRaw(Object.assign({ numero_albaran: '200002', matricula: '5678DEF' }, dos)),
    'b2.pdf': conR('DAYCO6PK1O90EE', '200002'),
    'c1.pdf': albaranRaw(Object.assign({ numero_albaran: '300003', matricula: '4321GHJ' }, dos)),
    'c2.pdf': conR('XXXXXXXXXXXXXX', '300003'),
  });
  ['ENT', 'PROC', 'FRA', 'FRAP'].forEach(id => e.mkFolder(id, id));
  subir(e, ['a1.pdf', 'a2.pdf', 'b1.pdf', 'b2.pdf', 'c1.pdf', 'c2.pdf'], 'ENT');
  e.run('procesarAlbaranes()');

  const de = n => tabla(e, 'Piezas').filter(p => p['Nº albarán'] === n);
  assert.equal(de('100001').length, 2, 'un importe único: no se crea otra pieza');
  assert.equal(de('100001').find(p => p['Reembolso'] === true)['Referencia pieza'], 'KRAFF47054');
  assert.equal(de('200002').length, 3);
  assert.equal(de('200002').find(p => p['Reembolso'] === true)['Referencia pieza'], 'DAYCO6PK1090EE', 'mismo importe: la referencia más parecida');
  assert.equal(de('300003').length, 3);
  assert.equal(de('300003').filter(p => p['Reembolso'] === true).length, 0, 'empate: no se marca ninguna');
  assert.match(e.log.alerts[e.log.alerts.length - 1][1], /R sin marcar, no se sabe qué pieza es: márcala a mano en Piezas \(1\):\n• c2\.pdf: albarán 300003/);
});

test('un filtro puesto (filas ocultas) no hace perder fechas ni fórmulas al escribir, y se deja como estaba', () => {
  const e = entorno({ 'a1.pdf': albaranRaw(), 'fra.pdf': facturaRaw() });
  ['ENT', 'PROC', 'FRA', 'FRAP'].forEach(id => e.mkFolder(id, id));
  const alb = e.ss.getSheetByName('Albaranes'), l = e.ss.getSheetByName('Líneas RM');
  // Como en la hoja real: Albaranes filtrada por Mes = oct y Líneas RM por una factura antigua (esconden también las filas vacías)
  alb.getRange(1, A('Mes'), 1500, 2).createFilter().setColumnFilterCriteria(A('Mes'), Filtro.visibles(['oct']));
  l.getRange(1, 1, 3000, 1).createFilter().setColumnFilterCriteria(1, Filtro.visibles(['FCR 99']));
  subir(e, ['a1.pdf'], 'ENT');
  e.run('procesarAlbaranes()');
  const fila = 2 + tabla(e, 'Albaranes').findIndex(a => a['Nº albarán'] === '462446');
  assert.ok(esFecha(alb.valor(fila, A('Fecha albarán'))), 'Fecha albarán guardada');
  assert.match(alb.cell(fila, A('Ver PDF')).f, /HYPERLINK/, 'enlace al PDF guardado');
  assert.match(alb.cell(fila, A('Mes')).f, /CHOOSE/, 'fórmula del Mes guardada');
  subir(e, ['fra.pdf'], 'FRA');
  e.run('procesarFacturasRM()');
  assert.ok(plano(tabla(e, 'Líneas RM')).every(x => x['Fecha albarán']), 'todas las líneas con su fecha');
  assert.deepEqual(plano(alb.getFilter().criterios), { [A('Mes')]: { visibles: ['oct'] } }, 'el filtro vuelve a estar puesto');
  assert.deepEqual(plano(l.getFilter().criterios), { 1: { visibles: ['FCR 99'] } });
});

test('procesarAlbaranes vincula una fila manual con el mismo importe y matrícula', () => {
  const e = entorno({ 'b1.pdf': albaranRaw({ numero_albaran: '999111', matricula: '4321GHJ', total: 100, base_imponible: 82.64, iva_importe: 17.36,
    lineas: [{ referencia: 'X1', descripcion: 'pieza', importe: 82.64, reembolso: false, cantidad: 1, precio_unitario: 82.64, descuento_pct: 0 }] }) });
  ['ENT', 'PROC'].forEach(id => e.mkFolder(id, id));
  subir(e, ['b1.pdf'], 'ENT');
  e.run('procesarAlbaranes()');
  const alb = tabla(e, 'Albaranes');
  assert.equal(alb.length, 3, 'no crea fila nueva: adopta la manual');
  const a = alb.find(x => x['Nº albarán'] === '999111');
  assert.equal(a['Precio con IVA'], 100);
  assert.match(a['Nota escaneo'], /Vinculado/);
  assert.equal(tabla(e, 'Piezas').filter(p => p['Nº albarán'] === '999111').length, 1);
  assert.equal(e.run("aISO_(leerTabla_('Albaranes').filas.find(f => f.v['Nº albarán'] === '999111').v['Fecha escaneo'])"), e.run('hoyISO_()'), 'fecha de escaneo = hoy');
});

test('procesarAlbaranes: si los importes no cuadran se relee; si siguen sin cuadrar se queda en Entrada con un aviso', () => {
  const malo = albaranRaw({ numero_albaran: '700001', total: 50 });                      // base 36,81 + IVA ≠ 50
  const bien = albaranRaw({ numero_albaran: '700012' });
  const malo3 = albaranRaw({ numero_albaran: '700003', total: 50 });
  const e = entorno({ 'c1.pdf': [malo, albaranRaw({ numero_albaran: '700001' })], 'c2.pdf': [malo3, malo3], 'c3.pdf': bien });
  ['ENT', 'PROC'].forEach(id => e.mkFolder(id, id));
  subir(e, ['c1.pdf', 'c2.pdf', 'c3.pdf'], 'ENT');
  e.run('procesarAlbaranes()');
  const nums = tabla(e, 'Albaranes').map(a => a['Nº albarán']);
  assert.ok(nums.includes('700001'), 'c1: la segunda lectura cuadra y se usa');
  assert.ok(nums.includes('700012'));
  assert.ok(!nums.includes('700003'), 'c2 no se añade');
  assert.deepEqual(e.peticiones.filter(p => p.pista).map(p => p.nombre).sort(), ['c1.pdf', 'c2.pdf'], 'segunda lectura sólo de los que no cuadran');
  assert.match(e.peticiones.find(p => p.pista).pista, /no cuadraban/);
  assert.deepEqual(e.carpetas.ENT.ficheros.map(f => f.nombre), ['c2.pdf'], 'el que sigue sin cuadrar se queda en Entrada');
  assert.equal(e.carpetas.PROC.ficheros.length, 2);
  const aviso = e.log.alerts[e.log.alerts.length - 1][1];
  assert.match(aviso, /No cuadran los importes, siguen en Entrada \(1\):\n• c2\.pdf: base 36,81 \+ IVA = 44,54, pero el total es 50/);
});

test('procesarAlbaranes: un albarán de ABONO va a Líneas RM (no a Albaranes); su pieza pedida sale Abonada con la fecha del abono', () => {
  const abono = albaranRaw({ numero_albaran: '486962', fecha: '2026-09-28', es_abono: true, total: -72.94, base_imponible: -60.28, matricula: '',
    lineas: [
      { referencia: 'BOSCHF026400517', descripcion: 'S0517 FILTRO', importe: -15.14, reembolso: false, albaran_origen: '01000478440' },
      { referencia: 'ELEVA05', descripcion: 'ELEVALUNAS', importe: -45.14, reembolso: false, albaran_origen: '01000476000' },
      { referencia: '', descripcion: 'SIGAUS', importe: -0.05, reembolso: false, albaran_origen: '01000476000' },
    ] });
  const linea = (ref, orig, imp) => ({ referencia: ref, descripcion: 'X', importe: imp, cantidad: -1, albaran_origen: orig });
  const fra = { numero_factura: 'FCR 7', fecha_factura: '2026-09-30', base_imponible: -65.33, iva_importe: -13.72, total: 1, albaranes: [
    { numero_albaran: '486962', fecha: '2026-09-28', matricula: '', importe: -65.28, es_abono: true,
      lineas: [linea('BOSCHF026400517', '01000478440', -15.14), linea('ELEVA05', '01000476000', -45.14), linea('OTRA1', '01000470000', -5)] }] };
  const e = entorno({ 'ab.pdf': abono, 'ab2.pdf': abono, 'fra.pdf': fra, 'fra2.pdf': fra });
  ['ENT', 'PROC', 'FRA', 'FRAP'].forEach(id => e.mkFolder(id, id));
  const p = e.ss.getSheetByName('Piezas'), ab = e.ss.getSheetByName('Abonos');
  p.put(2, 3, '478440'); p.put(2, 4, 'BOSCHF026400517'); p.put(2, 5, 'FILTRO AIRE'); p.put(2, 10, 15.14); p.put(2, 1, true);
  e.ctx.__r = p.getRange(2, 1); e.run('alEditar({ range: __r })');
  assert.equal(ab.valor(2, 5), 'Sin abonar');
  const idPieza = p.valor(2, 17);
  assert.equal(idPieza, 1);
  subir(e, ['ab.pdf'], 'ENT');
  e.run('procesarAlbaranes()');
  assert.ok(!tabla(e, 'Albaranes').some(a => a['Nº albarán'] === '486962'), 'no va a Albaranes');
  assert.equal(e.carpetas.PROC.ficheros.length, 1);
  assert.match(e.log.alerts[e.log.alerts.length - 1][1], /Albaranes de abono \(a Líneas RM y Abonos\): 1/);
  const lineas = () => tabla(e, 'Líneas RM').filter(l => l['Tipo'] === 'Abono');
  assert.deepEqual(plano(lineas().map(l => [l['Nº factura'], l['Referencia'], l['Albarán origen'], e.run(`aISO_(new Date(${+l['Fecha abono']}))`)])),
    [['', 'BOSCHF026400517', '478440', '2026-09-28'], ['', 'ELEVA05', '476000', '2026-09-28']], 'sin la línea SIGAUS');
  const filas = () => [2, 3, 4, 5].map(r => ({ ref: ab.valor(r, 7), estado: ab.valor(r, 5), fecha: e.run(`aISO_(hoja_('Abonos').getRange(${r}, 1).getValue())`), fra: ab.valor(r, 10) })).filter(f => f.ref);
  assert.deepEqual(filas(), [
    { ref: 'ELEVA05', estado: 'Sin solicitar', fecha: '2026-09-28', fra: '' },
    { ref: 'BOSCHF026400517', estado: 'Abonada', fecha: '2026-09-28', fra: '' }]);
  // Escanear otra vez el mismo albarán de abono no repite nada
  subir(e, ['ab2.pdf'], 'ENT');
  e.run('procesarAlbaranes()');
  assert.equal(lineas().length, 2);
  // Se corrige a mano a qué pieza corresponde un abono ("no" = de ninguna): se respeta aunque se reprocese la factura
  const l = e.ss.getSheetByName('Líneas RM'), filaEleva = 2 + tabla(e, 'Líneas RM').findIndex(x => x['Referencia'] === 'ELEVA05');
  l.put(filaEleva, 14, 'no'); e.ctx.__l = l.getRange(filaEleva, 14); e.run('alEditar({ range: __l })');
  assert.equal(l.valor(filaEleva, 14), 'NO');
  // Llega la factura quincenal con ese abono: las líneas escaneadas pasan a ser de la factura, con su pieza y su fecha
  subir(e, ['fra.pdf'], 'FRA');
  e.run('procesarFacturasRM()');
  assert.deepEqual(plano(lineas().map(x => [x['Nº factura'], x['Referencia'], x['ID pieza']])),
    [['FCR 7', 'BOSCHF026400517', idPieza], ['FCR 7', 'ELEVA05', 'NO'], ['FCR 7', 'OTRA1', '']], 'no se duplican');
  assert.deepEqual(filas(), [
    { ref: 'OTRA1', estado: 'Sin solicitar', fecha: '2026-09-30', fra: 'FCR 7' },
    { ref: 'ELEVA05', estado: 'Sin solicitar', fecha: '2026-09-28', fra: 'FCR 7' },
    { ref: 'BOSCHF026400517', estado: 'Abonada', fecha: '2026-09-28', fra: 'FCR 7' }], 'por la factura: fecha de la factura');
  for (let r = 2; l.valor(r, 5); r++) assert.match(l.cell(r, 13).f, /^=IF\(/, 'Conciliación sigue siendo una fórmula al rehacer Líneas RM');
  subir(e, ['fra2.pdf'], 'FRA');
  e.run('procesarFacturasRM()');
  assert.deepEqual(plano(lineas().map(x => x['ID pieza'])), [idPieza, 'NO', ''], 'reprocesar la factura conserva los emparejamientos');
  assert.equal(filas().length, 3);
});

test('procesarAlbaranes: si la hoja no se puede guardar no se mueve ningún PDF, y volver a procesar no duplica nada', () => {
  const e = entorno({ 'd1.pdf': albaranRaw({ numero_albaran: '800001', matricula: '9999ZZZ' }), 'd2.pdf': albaranRaw({ numero_albaran: '800002' }) });
  ['ENT', 'PROC'].forEach(id => e.mkFolder(id, id));
  subir(e, ['d1.pdf', 'd2.pdf'], 'ENT');
  // Como pasó con la validación "sólo fórmulas": Sheets rechaza lo escrito al guardar
  let veces = 0;
  const flush = e.ctx.SpreadsheetApp.flush;
  e.ctx.SpreadsheetApp.flush = () => { if (++veces === 1) throw new Error('Columna automática: se rellena sola.'); flush(); };
  e.run('procesarAlbaranes()');
  assert.deepEqual(e.carpetas.ENT.ficheros.map(f => f.nombre), ['d1.pdf', 'd2.pdf'], 'nada pasa a Procesados');
  assert.equal(e.carpetas.PROC.ficheros.length, 0);
  assert.match(e.log.alerts[e.log.alerts.length - 1][1], /No se ha podido guardar en la hoja: no se ha movido ningún PDF/);
  // Volver a procesar: lo que sí se había guardado se reconoce por el enlace a su PDF y no se duplica
  subir(e, [], 'ENT');
  e.run('procesarAlbaranes()');
  assert.equal(e.carpetas.PROC.ficheros.length, 2);
  assert.equal(e.carpetas.ENT.ficheros.length, 0);
  ['800001', '800002'].forEach(n => assert.equal(tabla(e, 'Albaranes').filter(a => a['Nº albarán'] === n).length, 1, n));
  assert.equal(tabla(e, 'Piezas').filter(p => p['Nº albarán'] === '800001').length, 2);
});

test('repararFormulas rehace el panel de Trabajos aunque su título combinado haya crecido (columna insertada)', () => {
  const e = entorno({});
  const trab = e.ss.getSheetByName('Trabajos');
  trab.combinadas = [{ r: 1, c: 1, nr: 1, nc: 7 }];  // como tras insertar una columna dentro del panel
  const fallos = () => e.log.alerts.filter(a => /Ha fallado/.test(String(a[0]) + String(a[1])));
  e.run('repararFormulas()');
  assert.deepEqual(fallos(), []);
  assert.deepEqual(trab.combinadas, [{ r: 1, c: 1, nr: 1, nc: 6 }], 'título del panel rehecho');
  e.run('repararFormulas()');
  assert.deepEqual(fallos(), []);
});

test('el diseño respeta lo cambiado a mano: anchos de columna, filtros, filas fijas y reglas de color propias', () => {
  const e = entorno({});
  const alb = e.ss.getSheetByName('Albaranes');
  const anchos = [], fijas = [];
  alb.setColumnWidth = (c, w) => anchos.push([c, w]);
  alb.getFrozenRows = () => 3; alb.setFrozenRows = n => fijas.push(n);
  const mia = { formula: '=$A2="x"', getBooleanCondition: () => ({ getCriteriaValues: () => ['=$A2="x"'] }) };
  alb.reglas = (alb.reglas || []).concat([mia]);
  e.run('repararFormulas()');
  assert.deepEqual(anchos, [], 'no cambia anchos de columna');
  assert.deepEqual(fijas, [], 'no quita filas fijas de más');
  assert.ok(alb.reglas.includes(mia), 'conserva la regla de color propia');
  const P = letra(A('Proveedor'));
  ['RM', 'Otros'].forEach(v => assert.ok(alb.reglas.some(r => r.formula === `=$${P}2="${v}"`), `color para Proveedor ${v}`));
  const formulas = alb.reglas.map(r => r.formula);
  assert.equal(formulas.length, new Set(formulas).size, 'sin reglas nuestras duplicadas');
});

test('versiones: al subir la versión se ejecutan las migraciones pendientes una vez y se reaplica el diseño', () => {
  const e = entorno({});
  assert.equal(e.props.VERSION_HOJA, e.run('VERSION'), 'setup apunta la versión');
  const hechas = [];
  e.ctx.__hechas = hechas;
  e.run(`MIGRACIONES['0.9.5'] = () => __hechas.push('0.9.5'); MIGRACIONES['1.0.0'] = () => __hechas.push('1.0.0'); MIGRACIONES['99.0.0'] = () => __hechas.push('99.0.0');`);
  e.props.VERSION_HOJA = e.run('VERSION');
  e.log.console.length = 0;
  e.run('actualizarAbonos()');
  assert.ok(!e.log.console.some(l => /aplicarDiseno/.test(l)), 'misma versión: no se toca el diseño');
  assert.deepEqual(hechas, []);
  e.props.VERSION_HOJA = '0.9.0';
  e.run('actualizarAbonos()');
  assert.deepEqual(hechas, ['0.9.5', '1.0.0'], 'en orden, sólo las que hay entre la versión de la hoja y la del código');
  assert.equal(e.props.VERSION_HOJA, e.run('VERSION'));
  assert.ok(e.log.console.some(l => /aplicarDiseno/.test(l)), 'cambio de versión menor/mayor: se reaplica el diseño');
  assert.ok(e.log.console.some(l => /Hoja actualizada de la versión 0\.9\.0/.test(l)));
  hechas.length = 0;
  e.run('actualizarAbonos()');
  assert.deepEqual(hechas, [], 'una sola vez');
});

test('migración 9.0.0: Albaranes y Trabajos pasan al orden nuevo sin perder datos y ocultan las columnas sobrantes', () => {
  const e = entorno({});
  const alb = e.ss.getSheetByName('Albaranes'), trab = e.ss.getSheetByName('Trabajos');
  const antes = { Albaranes: plano(tabla(e, 'Albaranes')), Trabajos: plano(tabla(e, 'Trabajos')) };
  assert.ok(antes.Albaranes.length && antes.Trabajos.length, 'hay datos que mover');
  // Hoja de la versión 8: columnas en el orden antiguo, nada oculto y el título del panel combinado
  e.run(`(() => {
    const viejo = { Albaranes: ['Fecha escaneo', 'Fecha albarán', 'Mes', 'Quincena', 'Proveedor', 'Nº albarán', 'Nº trabajo', 'Matrícula',
        'Precio con IVA', 'Precio - abonos', 'Coche', 'Cliente', 'Ver PDF', 'Avisos', 'Nota escaneo'],
      Trabajos: ['Nº trabajo', 'Mes', 'Quincena', 'Matrícula', 'Coche', 'Cliente', 'Recambios', 'Precio - abono', 'Recambios facturables RM',
        'Recambios facturables Otros', 'Factura', 'Beneficio', 'Pagado', 'Avisos'] };
    Object.keys(viejo).forEach(n => {
      const nuevo = ESQUEMA[n].cabeceras, vis = ESQUEMA[n].visibles;
      ESQUEMA[n].cabeceras = viejo[n]; ESQUEMA[n].visibles = 0;
      try { reordenarColumnas_(n); } finally { ESQUEMA[n].cabeceras = nuevo; ESQUEMA[n].visibles = vis; }
    });
  })()`);
  alb.ocultas = []; trab.ocultas = [];
  alb.getRange(1, 1, 1500, 15).createFilter().setColumnFilterCriteria(3, Filtro.visibles(['sep']));  // Mes en la columna C antigua
  trab.combinadas.push({ r: 1, c: 1, nr: 1, nc: 6 });
  assert.equal(alb.valor(1, 1), 'Fecha escaneo'); assert.equal(trab.valor(4, 1), 'Nº trabajo');
  e.props.VERSION_HOJA = '8.0.2';
  e.run('actualizarAbonos()');
  assert.deepEqual(e.log.alerts.filter(a => /Ha fallado/.test(String(a[0]) + String(a[1]))), []);
  assert.equal(e.props.VERSION_HOJA, e.run('VERSION'));
  ['Albaranes', 'Trabajos'].forEach(n => {
    const sh = n === 'Albaranes' ? alb : trab, fc = n === 'Albaranes' ? 1 : 4, esq = ESQ[n];
    assert.deepEqual(esq.cabeceras.map((h, i) => sh.valor(fc, i + 1)), esq.cabeceras, `${n}: cabeceras en el orden nuevo`);
    assert.deepEqual(plano(tabla(e, n)), antes[n], `${n}: cada dato sigue bajo su cabecera`);
    assert.deepEqual(plano(sh.ocultas).sort((a, b) => a - b), plano(esq.cabeceras.slice(esq.visibles).map((h, i) => esq.visibles + i + 1)), `${n}: columnas sobrantes ocultas`);
  });
  assert.equal(trab.valor(1, 1), 'Resumen (según filtro)', 'panel rehecho');
  assert.deepEqual(plano(alb.getFilter().criterios), { [A('Mes')]: { visibles: ['sep'] } }, 'el criterio del filtro sigue en Mes');
  e.props.VERSION_HOJA = '8.0.2';
  e.run('actualizarAbonos()');
  assert.deepEqual(plano(tabla(e, 'Albaranes')), antes.Albaranes, 'repetir la migración no cambia nada');
});

test('atajo de Coches: selecciona la Matrícula de la primera fila libre', () => {
  const e = entorno({});
  const coches = e.ss.getSheetByName('Coches');
  const libre = tabla(e, 'Coches').length + 2;
  e.run('irAFilaLibreCoches()');
  assert.deepEqual(plano(coches.activa), [libre, 1]);
  coches.put(libre, 1, '0000AAA');
  e.run('irAFilaLibreCoches()');
  assert.deepEqual(plano(coches.activa), [libre + 1, 1], 'tras añadir un coche, la siguiente');
});

test('escribir a mano en una columna automática vuelve a poner su fórmula y avisa', () => {
  const e = entorno({});
  const alb = e.ss.getSheetByName('Albaranes');
  const buena = alb.cell(2, A('Coche')).f;  // Coche
  alb.put(2, A('Coche'), 'Otro coche');
  e.ctx.__r = alb.getRange(2, A('Coche')); e.run('alEditar({ range: __r })');
  assert.equal(alb.cell(2, A('Coche')).f, buena);
  assert.ok(e.log.toasts.some(t => /"Coche" se rellena sola/.test(t)));
});

test('Trabajos: Mes y Quincena se rellenan solos con la fecha de hoy y se pueden cambiar a mano', () => {
  const e = entorno({});
  const trab = e.ss.getSheetByName('Trabajos');
  trab.put(7, T('Matrícula'), '1234ABC');
  e.ctx.__r = trab.getRange(7, T('Matrícula')); e.run('alEditar({ range: __r })');
  assert.equal(trab.valor(7, T('Mes')), e.run('mesDe(hoyISO_())'));
  assert.equal(trab.valor(7, T('Quincena')), e.run('quincenaDe(hoyISO_())'));
  trab.put(7, T('Mes'), 'ene'); trab.put(7, T('Quincena'), 1);
  e.ctx.__r = trab.getRange(7, T('Mes'), 1, 2); e.run('alEditar({ range: __r })');
  assert.equal(trab.valor(7, T('Mes')), 'ene', 'el cambio a mano se respeta');
  e.run('repararFormulas()');
  assert.equal(trab.valor(7, T('Mes')), 'ene'); assert.equal(trab.valor(7, T('Quincena')), 1);
});

test('si el diseño se corta a medias, se repite la próxima vez aunque la versión ya esté apuntada', () => {
  const e = entorno({});
  e.run("PropertiesService.getScriptProperties().setProperty('VERSION_HOJA', '3.0.0')");
  e.run("var __orig = aplicarDiseno_; aplicarDiseno_ = () => { throw new Error('validación'); }");
  assert.throws(() => e.run('actualizarHoja_(false)'), /validación/);
  e.run('aplicarDiseno_ = __orig');
  assert.equal(e.run('actualizarHoja_(false)'), true, 'lo vuelve a intentar');
  assert.equal(e.run("PropertiesService.getScriptProperties().getProperty('DISENO_PENDIENTE')"), null);
  assert.equal(e.run('actualizarHoja_(false)'), false, 'y una vez terminado ya no');
});

test('Abonos con columnas reordenadas a mano: cada dato va bajo su cabecera, se respeta el orden y no quedan desplegables', () => {
  const e = entorno({});
  const ab = e.ss.getSheetByName('Abonos'), p = e.ss.getSheetByName('Piezas');
  const orden = ['Fecha abono', 'Descripción pieza', 'Matrícula', 'Precio sin IVA', 'Precio con IVA', 'Estado', 'Nº albarán', 'Referencia',
    'Fecha solicitud', 'Factura RM', 'Nota', 'Días pendiente', 'ID pieza'];
  orden.forEach((h, i) => ab.put(1, i + 1, h));
  e.run('repararFormulas()');
  assert.deepEqual(orden.map((h, i) => ab.valor(1, i + 1)), orden, 'reparar no vuelve a poner el orden por defecto');
  assert.ok([5, 6].every(c => !((ab.validaciones.get('3,' + c) || {}).lista)), 'sin desplegable de Estado: la tabla no se edita');
  p.put(2, 1, true); p.put(2, 3, '123456'); p.put(2, 4, 'REF9'); p.put(2, 5, 'FILTRO'); p.put(2, 10, 10);
  e.ctx.__r = p.getRange(2, 1); e.run('alEditar({ range: __r })');
  const fila = Object.fromEntries(orden.map((h, i) => [h, ab.valor(2, i + 1)]));
  assert.equal(fila['Estado'], 'Sin abonar'); assert.equal(fila['Nº albarán'], '123456'); assert.equal(fila['Referencia'], 'REF9');
  assert.equal(fila['Precio sin IVA'], 10); assert.equal(fila['ID pieza'], 1);
  assert.equal(ab.cell(2, 5).f, '=ROUND($D2*(1+IVA),2)');
  e.run('actualizarAbonos()');
  assert.equal(ab.valor(3, 2), '', 'no se duplica');
});

test('Trabajos: fórmulas viejas por debajo de los datos en columnas de entrada (#REF! en Matrícula) se limpian y no cuentan como trabajos', () => {
  const e = entorno({});
  const trab = e.ss.getSheetByName('Trabajos');
  for (let r = 7; r <= 300; r++) { trab.grid.set(r + ',' + T('Matrícula'), { v: '#REF!', f: '=IF(#REF!="","",1)' }); trab.put(r, T('Factura'), `=IF(A${r}="","",1)`); }
  assert.equal(tabla(e, 'Trabajos').length, 2, 'las filas con #REF! no son trabajos');
  e.run('repararFormulas()');
  assert.equal(trab.cell(7, T('Matrícula')), undefined, 'Matrícula limpia');
  assert.equal(trab.cell(300, T('Factura')), undefined, 'Factura limpia');
  assert.equal(trab.valor(6, T('Matrícula')), tabla(e, 'Trabajos')[1]['Matrícula'], 'los datos no se tocan');
  trab.put(7, T('Matrícula'), '1234ABC');
  e.ctx.__r = trab.getRange(7, T('Matrícula')); e.run('alEditar({ range: __r })');
  assert.ok(trab.valor(7, T('Nº trabajo')), 'el trabajo nuevo va justo debajo de los datos');
});

test('Trabajos: con una fila en blanco entre el panel y la cabecera, reparar y editar siguen funcionando', () => {
  const e = entorno({});
  const trab = e.ss.getSheetByName('Trabajos');
  trab.insertRowsBefore(4, 1);
  const antes = [6, 7].map(r => ['Nº trabajo', 'Mes', 'Quincena', 'Matrícula', 'Factura', 'Pagado'].map(T).map(c => trab.valor(r, c)));
  e.run('repararFormulas()');
  assert.deepEqual(e.log.console.filter(m => /^\[ERROR\]/.test(m)), []);
  assert.equal(trab.valor(4, T('Nº trabajo')), '', 'la fila en blanco se respeta');
  assert.equal(trab.valor(5, T('Nº trabajo')), 'Nº trabajo');
  assert.deepEqual([6, 7].map(r => ['Nº trabajo', 'Mes', 'Quincena', 'Matrícula', 'Factura', 'Pagado'].map(T).map(c => trab.valor(r, c))), antes, 'los datos no se tocan');
  assert.match(trab.cell(6, T('Beneficio')).f, new RegExp(letra(T('Factura')) + '6'), 'Beneficio en su fila');
  trab.put(8, T('Matrícula'), '1234ABC');
  e.ctx.__r = trab.getRange(8, T('Matrícula')); e.run('alEditar({ range: __r })');
  assert.ok(trab.valor(8, T('Nº trabajo')), 'trabajo creado a mano en la fila nueva');
  assert.equal(tabla(e, 'Trabajos').length, 3);
});

test('Albaranes: una fila en blanco encima de la cabecera se quita sola; con datos encima, error sin escribir nada', () => {
  const e = entorno({});
  const alb = e.ss.getSheetByName('Albaranes');
  alb.insertRowsBefore(1, 1);
  e.run('repararFormulas()');
  assert.equal(alb.valor(1, A('Fecha escaneo')), 'Fecha escaneo', 'la cabecera vuelve a la fila 1');
  assert.equal(tabla(e, 'Albaranes').length, 3);
  alb.insertRowsBefore(1, 1); alb.put(1, A('Mes'), 'nota');
  const celdas = alb.grid.size;
  assert.throws(() => e.run("leerTabla_('Albaranes')"), /filas con datos encima de la cabecera/);
  assert.equal(alb.grid.size, celdas, 'no escribe nada');
});

test('edición manual en Albaranes: fecha, proveedor y trabajo automáticos; NUEVO abre otro trabajo', () => {
  const e = entorno({});
  const alb = e.ss.getSheetByName('Albaranes');
  const fila = 10;
  alb.put(fila, A('Matrícula'), '1234 abc'); alb.put(fila, A('Precio con IVA'), 44.54);
  const rango = alb.getRange(fila, A('Matrícula'), 1, A('Precio con IVA') - A('Matrícula') + 1);
  e.ctx.__rango = rango;
  e.run('alEditar({ range: __rango })');
  const t = tabla(e, 'Albaranes').find(a => a['Matrícula'] === '1234ABC');
  assert.ok(t, 'fila procesada');
  assert.ok(esFecha(t['Fecha escaneo']));
  assert.ok(esFecha(t['Fecha albarán']));
  assert.equal(t['Proveedor'], 'RM');
  assert.equal(t['Nº trabajo'], '4ABC-1');
  // Segunda fila del mismo coche -> mismo trabajo abierto
  alb.put(11, A('Matrícula'), '1234ABC'); alb.put(11, A('Precio con IVA'), 20);
  e.ctx.__r2 = alb.getRange(11, A('Matrícula'), 1, A('Precio con IVA') - A('Matrícula') + 1); e.run('alEditar({ range: __r2 })');
  assert.equal(tabla(e, 'Albaranes').filter(a => a['Nº trabajo'] === '4ABC-1').length, 2);
  // NUEVO -> segundo trabajo abierto en paralelo
  alb.put(11, A('Nº trabajo'), 'NUEVO');
  e.ctx.__r3 = alb.getRange(11, A('Nº trabajo')); e.run('alEditar({ range: __r3 })');
  assert.equal(tabla(e, 'Albaranes').find(a => a['Precio con IVA'] === 20)['Nº trabajo'], '4ABC-2');
  assert.equal(tabla(e, 'Trabajos').filter(x => x['Matrícula'] === '1234ABC').length, 2);
});

test('una matrícula que no está en Coches se añade allí (sólo la matrícula) para rellenar cliente y coche a mano', () => {
  const e = entorno({});
  const alb = e.ss.getSheetByName('Albaranes'), trab = e.ss.getSheetByName('Trabajos');
  const enCoches = p => tabla(e, 'Coches').filter(c => c['Matrícula'] === p);
  alb.put(10, A('Matrícula'), '9999 xyz'); alb.put(10, A('Precio con IVA'), 20);
  e.ctx.__r = alb.getRange(10, A('Matrícula'), 1, A('Precio con IVA') - A('Matrícula') + 1); e.run('alEditar({ range: __r })');
  assert.equal(enCoches('9999XYZ').map(c => c['Cliente'] + '|' + c['Coche']).join(), '|', 'añadida, sin cliente ni coche');
  assert.ok(e.log.toasts.some(t => /9999XYZ añadida a la pestaña Coches/.test(t)));
  alb.put(11, A('Matrícula'), '9999XYZ'); alb.put(11, A('Precio con IVA'), 30);
  e.ctx.__r = alb.getRange(11, A('Matrícula'), 1, A('Precio con IVA') - A('Matrícula') + 1); e.run('alEditar({ range: __r })');
  assert.equal(enCoches('9999XYZ').length, 1, 'no se repite');
  trab.put(7, T('Matrícula'), '8888WWW');
  e.ctx.__r = trab.getRange(7, T('Matrícula')); e.run('alEditar({ range: __r })');
  assert.equal(enCoches('8888WWW').length, 1, 'también desde Trabajos');
});

test('corregir la matrícula de un albarán: cambia de trabajo sin perder la factura y sin dejar restos', () => {
  const e = entorno({});
  const alb = e.ss.getSheetByName('Albaranes');
  const editar = (hoja, fila, cols, vals) => {
    const sh = e.ss.getSheetByName(hoja);
    cols.forEach((c, i) => sh.put(fila, c, vals[i]));
    e.ctx.__r = sh.getRange(fila, Math.min(...cols), 1, Math.max(...cols) - Math.min(...cols) + 1);
    e.run('alEditar({ range: __r })');
  };
  const job = fila => alb.valor(fila, A('Nº trabajo'));
  const trabajo = num => plano(tabla(e, 'Trabajos')).find(t => t['Nº trabajo'] === num);
  const factura = (num, v, pagado) => e.run(`(() => { const t = leerTabla_('Trabajos'); actualizarFila_(t, t.filas.find(f => f.v['Nº trabajo'] === '${num}').fila, { 'Factura': ${v}, 'Pagado': ${pagado} }); })()`);
  const enCoches = p => plano(tabla(e, 'Coches')).some(c => c['Matrícula'] === p);

  // Único albarán, sin trabajo abierto en la buena: el trabajo se convierte, con su factura; la errónea sale de Coches.
  editar('Albaranes', 10, [A('Matrícula'), A('Precio con IVA')], ['1111ERR', 50]);
  factura('1ERR-1', 500, true);
  editar('Albaranes', 10, [A('Matrícula')], ['2222NEW']);
  assert.equal(job(10), '2NEW-1');
  assert.equal(trabajo('1ERR-1'), undefined);
  assert.equal(trabajo('2NEW-1')['Factura'], 500);
  assert.equal(trabajo('2NEW-1')['Pagado'], true);
  assert.ok(!enCoches('1111ERR') && enCoches('2222NEW'));

  // Sin precio también cambia; hacia un trabajo abierto, la factura pasa con él y el viejo se borra.
  editar('Albaranes', 11, [A('Matrícula'), A('Precio con IVA')], ['3333ERR', 20]);
  editar('Albaranes', 12, [A('Matrícula'), A('Precio con IVA')], ['1234ABC', 30]);
  factura('3ERR-1', 200, false);
  alb.put(11, A('Precio con IVA'), '');
  editar('Albaranes', 11, [A('Matrícula')], ['1234ABC']);
  assert.equal(job(11), '4ABC-1');
  assert.equal(trabajo('3ERR-1'), undefined);
  assert.equal(trabajo('4ABC-1')['Factura'], 200);

  // El destino ya tiene factura: no se mezclan y el viejo se queda, avisando.
  editar('Albaranes', 13, [A('Matrícula'), A('Precio con IVA')], ['4444ERR', 20]);
  factura('4ERR-1', 100, false);
  e.log.toasts.length = 0;
  editar('Albaranes', 13, [A('Matrícula')], ['1234ABC']);
  assert.equal(job(13), '4ABC-1');
  assert.equal(trabajo('4ERR-1')['Factura'], 100);
  assert.equal(trabajo('4ABC-1')['Factura'], 200);
  assert.ok(e.log.toasts.some(t => /4ERR-1 se queda sin albaranes/.test(t)));

  // El viejo aún tiene otro albarán: no se toca.
  editar('Albaranes', 14, [A('Matrícula'), A('Precio con IVA')], ['5555ERR', 20]);
  editar('Albaranes', 15, [A('Matrícula'), A('Precio con IVA')], ['5555ERR', 20]);
  factura('5ERR-1', 300, false);
  editar('Albaranes', 14, [A('Matrícula')], ['6666NEW']);
  assert.equal(job(14), '6NEW-1');
  assert.equal(trabajo('5ERR-1')['Factura'], 300);
  assert.equal(trabajo('6NEW-1')['Factura'], '');

  // Coincide con varias matrículas: no cambia nada ni aparece un trabajo basura.
  const numTrabajos = plano(tabla(e, 'Trabajos')).length;
  e.ss.getSheetByName('Coches').put(e.ss.getSheetByName('Coches').getLastRow() + 1, 1, '4321GHK');
  editar('Albaranes', 15, [A('Matrícula')], ['432']);
  assert.equal(job(15), '5ERR-1');
  assert.equal(plano(tabla(e, 'Trabajos')).length, numTrabajos);

  // Trabajo que no existe en Trabajos: se recoloca igual.
  alb.put(15, A('Nº trabajo'), '7XXX-9');
  editar('Albaranes', 15, [A('Matrícula')], ['7777NEW']);
  assert.equal(job(15), '7NEW-1');

  // Matrícula borrada: el albarán sale del trabajo y, sin factura, este desaparece.
  editar('Albaranes', 15, [A('Matrícula')], ['']);
  assert.equal(job(15), '');
  assert.equal(trabajo('7NEW-1'), undefined);
});

test('corregir la matrícula en Trabajos cambia el trabajo entero: nº, albaranes y factura', () => {
  const e = entorno({});
  const alb = e.ss.getSheetByName('Albaranes'), trab = e.ss.getSheetByName('Trabajos');
  [10, 11].forEach(r => { alb.put(r, A('Matrícula'), '1111ERR'); alb.put(r, A('Precio con IVA'), 50); e.ctx.__r = alb.getRange(r, A('Matrícula'), 1, A('Precio con IVA') - A('Matrícula') + 1); e.run('alEditar({ range: __r })'); });
  const fila = e.run("(() => { const t = leerTabla_('Trabajos'); const f = t.filas.find(x => x.v['Nº trabajo'] === '1ERR-1'); actualizarFila_(t, f.fila, { 'Factura': 500 }); return f.fila; })()");

  const editar = plate => { trab.put(fila, T('Matrícula'), plate); e.ctx.__r = trab.getRange(fila, T('Matrícula')); e.run('alEditar({ range: __r })'); };
  const filas = () => [10, 11].map(r => `${alb.valor(r, A('Nº trabajo'))} ${alb.valor(r, A('Matrícula'))}`).join();

  editar('2221ERR');
  assert.equal(filas(), '1ERR-1 2221ERR,1ERR-1 2221ERR', 'mismo prefijo: conserva el nº y corrige los albaranes');
  editar('2222NEW');
  assert.equal(filas(), '2NEW-1 2222NEW,2NEW-1 2222NEW');
  const t = plano(tabla(e, 'Trabajos')).find(x => x['Nº trabajo'] === '2NEW-1');
  assert.equal(t['Factura'], 500);
  assert.equal(t['Matrícula'], '2222NEW');
  assert.ok(!plano(tabla(e, 'Coches')).some(c => ['1111ERR', '2221ERR'].includes(c['Matrícula'])));
  assert.ok(plano(tabla(e, 'Coches')).some(c => c['Matrícula'] === '2222NEW'));
});

test('Matrícula en Albaranes: busca por cualquier combinación de letras o dígitos, no sólo al principio', () => {
  const e = entorno({});  // Coches del piloto: 1234ABC, 4321GHJ, 5678DEF
  const alb = e.ss.getSheetByName('Albaranes');

  // Coincidencia única por las LETRAS (van al final: el desplegable nativo de Sheets sólo busca desde el principio)
  alb.put(10, A('Matrícula'), 'ghj');
  e.ctx.__unica = alb.getRange(10, A('Matrícula')); e.run('alEditar({ range: __unica })');
  assert.equal(alb.valor(10, A('Matrícula')), '4321GHJ', 'se autocompleta con la única matrícula que contiene "GHJ"');
  assert.ok(e.log.toasts.some(t => t.includes('"GHJ" → 4321GHJ')), 'avisa de qué matrícula ha puesto');

  // Texto demasiado corto (< 3): no se busca, se deja tal cual y sigue el aviso habitual de "no está en Coches"
  // (ese aviso sólo se comprueba cuando la fila ya tiene precio, igual que sin este cambio)
  e.log.toasts.length = 0;
  alb.put(11, A('Matrícula'), '4'); alb.put(11, A('Precio con IVA'), 44.54);
  e.ctx.__corto = alb.getRange(11, A('Matrícula'), 1, A('Precio con IVA') - A('Matrícula') + 1); e.run('alEditar({ range: __corto })');
  assert.equal(alb.valor(11, A('Matrícula')), '4', 'texto demasiado corto: se deja tal cual, sin adivinar');
  assert.equal(e.log.toasts.length, 1);
  assert.match(e.log.toasts[0], /no está en la pestaña Coches/);

  // Coincide con varias: no adivina, deja lo escrito y lista las candidatas (se añade un coche a propósito)
  e.log.toasts.length = 0;
  const coches = e.ss.getSheetByName('Coches');
  const filaNueva = coches.getLastRow() + 1;
  coches.put(filaNueva, 1, '4321GHK'); coches.put(filaNueva, 2, 'Otro'); coches.put(filaNueva, 3, 'Otro coche');
  alb.put(12, A('Matrícula'), '432');
  e.ctx.__varias = alb.getRange(12, A('Matrícula')); e.run('alEditar({ range: __varias })');
  assert.equal(alb.valor(12, A('Matrícula')), '432', 'coincide con varias: no adivina, deja lo escrito');
  assert.equal(e.log.toasts.length, 1);
  assert.match(e.log.toasts[0], /"432" coincide con 2 matrículas: 4321GHJ, 4321GHK/);

  // Matrícula ya exacta: no hay búsqueda ni aviso de "completada"
  e.log.toasts.length = 0;
  alb.put(13, A('Matrícula'), '5678DEF');
  e.ctx.__exacta = alb.getRange(13, A('Matrícula')); e.run('alEditar({ range: __exacta })');
  assert.equal(alb.valor(13, A('Matrícula')), '5678DEF');
  assert.ok(!e.log.toasts.some(t => t.includes('completada')), 'ya era exacta: no hay aviso de autocompletado');
});

test('edición en Piezas: reembolso exige nº de albarán y rellena la fecha; añade la fila a Abonos', () => {
  const e = entorno({});
  const p = e.ss.getSheetByName('Piezas');
  p.put(2, 1, true); p.put(2, 5, 'Pieza a mano'); p.put(2, 10, 50);
  e.ctx.__p = p.getRange(2, 1); e.run('alEditar({ range: __p })');
  assert.equal(p.valor(2, 1), false, 'sin nº de albarán se desmarca');
  p.put(2, 3, '123456'); p.put(2, 1, true);
  e.log.console.length = 0;
  e.ss.io.escrito = false; e.ss.io.lecturasTrasEscribir.length = 0;
  e.ctx.__p2 = p.getRange(2, 1); e.run('alEditar({ range: __p2 })');
  // Todas las lecturas antes de la primera escritura: en Sheets, leer tras escribir espera al recálculo
  assert.deepEqual(Array.from(e.ss.io.lecturasTrasEscribir), [], 'ninguna lectura después de escribir al marcar una pieza');
  // Tiempos en Apps Script ▸ Ejecuciones: una línea por paso del camino "marcar pieza"
  const linea = pre => e.log.console.find(l => l.startsWith(pre)) || '';
  assert.match(linea('⏱ alEditar (Piezas)'), /esperar bloqueo \d+ms, editar Piezas \d+ms \| total \d+ms/);
  assert.match(linea('⏱ editarPiezas_'), new RegExp('leer Piezas \\(1 de 1 filas\\).*leer fórmulas de 1 fila\\(s\\).*leer Líneas RM \\(0 de 0 filas\\).*' +
    'leer Albaranes \\(3 de 3 filas\\).*leer Abonos \\(0 filas\\).*revisar 1 fila\\(s\\) editada\\(s\\).*Abonos.*guardar cambios y recalcular \\d+ms'));
  assert.match(linea('⏱ escribirAbonos_'), /preparar piezas \(1 pedidas, 0 IDs nuevos, 0 fechas escritas\).*cruce \(0 abonos, 0 emparejados ahora\).*tabla rehecha \|.*Sin abonar 1/);
  assert.match(linea('■ alEditar'), /^■ alEditar: \d+ms \(escribir en Registro \d+ms\)$/);
  assert.equal(e.log.console.filter(l => /^\[INFO\] reconstruirAbonos/.test(l)).length, 0, 'el resumen ya no se escribe en Registro');
  assert.equal(p.valor(2, 1), true);
  assert.ok(esFecha(p.valor(2, 12)), 'fecha de reembolso');
  assert.equal(p.valor(2, 15), 'Manual');
  assert.match(p.cell(2, 16).f, /Falta el nº de albarán/, 'la fila escrita a mano recibe sus fórmulas');
  assert.equal(p.valor(2, 17), 1, 'y su ID');
  assert.equal(e.ss.getSheetByName('Abonos').valor(2, 5), 'Sin abonar');
  // Pieza de un albarán de proveedor "Otros": no se reclama a RM
  const alb = e.ss.getSheetByName('Albaranes');
  alb.put(5, A('Proveedor'), 'Otros'); alb.put(5, A('Nº albarán'), '999'); alb.put(5, A('Matrícula'), '1234ABC'); alb.put(5, A('Precio con IVA'), 20);
  p.put(3, 3, '999'); p.put(3, 5, 'De otro proveedor'); p.put(3, 10, 10); p.put(3, 1, true);
  e.ctx.__p4 = p.getRange(3, 1); e.run('alEditar({ range: __p4 })');
  assert.equal(e.ss.getSheetByName('Abonos').valor(2, 6), '123456');
  assert.equal(e.ss.getSheetByName('Abonos').valor(3, 6), '', 'la pieza de "Otros" no entra en Abonos');
  // desmarcar limpia la fecha y borra su fila de Abonos
  p.put(2, 1, false);
  e.ctx.__p3 = p.getRange(2, 1); e.run('alEditar({ range: __p3 })');
  assert.equal(p.valor(2, 12), '');
  assert.equal(e.ss.getSheetByName('Abonos').valor(2, 5), '');
});

// ---- Factura RM ----
function facturaRaw() {
  const alb = (n, f, m, imp, ab, ls) => ({ numero_albaran: n, fecha: f, matricula: m, importe: imp, es_abono: ab, lineas: ls });
  return { numero_factura: 'FCR 00001', fecha_factura: '2026-09-15', base_imponible: 100, iva_importe: 21, total: 121, albaranes: [
    alb('443645', '2026-09-03', '5678DEF', 130.10 + 53.90, false, [{ referencia: 'VARTAA8', descripcion: 'A8 AGM', importe: 130.10, albaran_origen: '' }, { referencia: 'PEUGE1', descripcion: 'CAJA', importe: 53.90, albaran_origen: '' }]),
    alb('446334', '2026-09-04', '', -130.10, true, [{ referencia: 'VARTAA8', descripcion: 'A8 AGM', importe: -130.10, cantidad: -1, albaran_origen: '01000443645' }]),
    alb('446877', '2026-09-07', '', -6.79, true, [{ referencia: 'WALKE80477', descripcion: 'CLAMP', importe: -6.79, cantidad: -1, albaran_origen: '01000439139' }]),
  ] };
}

test('factura RM: las líneas de abono que la lectura se salta se quedan en la factura si cuadran con el bloque, también al reprocesarla', () => {
  const l = (referencia, importe) => ({ referencia, descripcion: 'PIEZA ' + referencia, cantidad: -1, precio_unitario: -importe, descuento_pct: 0, importe, albaran_origen: '01000470000', reembolso: false });
  const escaneado = (numero_albaran, lineas) => albaranRaw({ numero_albaran, es_abono: true, matricula: '', fecha: '2026-10-07', total: -100, base_imponible: -82.64, lineas });
  const bloque = (numero_albaran, importe, lineas) => ({ numero_albaran, fecha: '2026-10-07', matricula: '', importe, es_abono: true, lineas });
  const fra = () => Object.assign(facturaRaw(), { numero_factura: 'FCR 2', fecha_factura: '2026-10-15', base_imponible: -140, iva_importe: -29.4, total: 169.4, albaranes: [
    bloque('600001', -140, [l('A', -54.81), l('B', -38.25)]),
    bloque('600002', -50, [l('X', -30)]),
    bloque('600003', -10, [l('N', -10)]),
  ] });
  const e = entorno({ 'ab1.pdf': escaneado('600001', [l('C', -9.14), l('D', -5.98), l('E', -31.82)]), 'ab2.pdf': escaneado('600002', [l('Y', -5)]), 'ab3.pdf': escaneado('600009', [l('Z', -7)]), 'f1.pdf': fra(), 'f2.pdf': fra() });
  ['FRA', 'FRAP', 'ENT', 'PROC'].forEach(id => e.mkFolder(id, id));
  subir(e, ['ab1.pdf', 'ab2.pdf', 'ab3.pdf'], 'ENT');
  e.run('procesarAlbaranes()');
  subir(e, ['f1.pdf'], 'FRA');
  e.run('procesarFacturasRM()');
  subir(e, ['f2.pdf'], 'FRA');
  e.run('procesarFacturasRM()');

  const lineas = plano(tabla(e, 'Líneas RM')), facturaDe = r => lineas.filter(x => x['Referencia'] === r).map(x => x['Nº factura']).join();
  assert.equal(lineas.filter(x => x['Nº albarán'] === '600001').length, 5, 'ni se pierden ni se duplican al reprocesar');
  ['A', 'B', 'C', 'D', 'E'].forEach(r => assert.equal(facturaDe(r), 'FCR 2', r));
  assert.equal(facturaDe('Y'), '', 'con ella el bloque no da su importe: se queda sin factura');
  assert.equal(facturaDe('Z'), '', 'otro albarán de abono');
  const estado = plano(tabla(e, 'Facturas RM')).find(x => x['Nº factura'] === 'FCR 2')['Estado'];
  assert.doesNotMatch(estado, /600001/, 'bloque completo: sin aviso');
  assert.match(estado, /600002/);
});

test('factura RM: guarda líneas, detecta abonos (Abonada / Sin solicitar) y avisa si el IVA no cuadra', () => {
  const e = entorno({ 'fra.pdf': facturaRaw() });
  ['FRA', 'FRAP', 'ENT', 'PROC'].forEach(id => e.mkFolder(id, id));
  // Pieza escaneada de 5678DEF con reembolso solicitado (albarán 443645)
  const p = e.ss.getSheetByName('Piezas');
  p.put(2, 1, true); p.put(2, 3, '443645'); p.put(2, 4, 'VARTAA8'); p.put(2, 5, 'A8 AGM'); p.put(2, 10, 130.10); p.put(2, 12, e.run('new Date(2026, 8, 3)'));
  subir(e, ['fra.pdf'], 'FRA');
  e.run('procesarFacturasRM()');
  const fr = tabla(e, 'Facturas RM');
  assert.equal(fr.length, 1);
  assert.equal(fr[0]['Nº factura'], 'FCR 00001');
  assert.equal(fr[0]['Mes'], 9); assert.equal(fr[0]['Quincena'], 1);
  assert.match(fr[0]['Estado'], /^⚠/, 'la base de prueba (100) no cuadra con los albaranes: debe avisar');
  assert.equal(tabla(e, 'Líneas RM').length, 4);
  const ab = e.ss.getSheetByName('Abonos');
  const estados = Array.from([2, 3].map(r => ab.valor(r, 5))).sort();
  assert.equal(ab.valor(4, 5), '');
  assert.deepEqual(estados, ['Abonada', 'Sin solicitar']);
  // El bloque de abono trae la matrícula vacía: se toma la del albarán original (compra 443645 de la misma factura)
  const lineaAbono = tabla(e, 'Líneas RM').find(x => x['Tipo'] === 'Abono' && x['Referencia'] === 'VARTAA8');
  assert.equal(lineaAbono['Matrícula'], '5678DEF');
  assert.equal([2, 3].map(r => ab.valor(r, 8)).filter(Boolean).join(), '5678DEF', 'y llega a Abonos; la del 439139 (no está en ninguna parte) queda vacía');
  assert.equal(e.carpetas.FRAP.ficheros.length, 1);
  // Reprocesar la misma factura no duplica líneas
  subir(e, ['fra.pdf'], 'FRA');
  e.run('procesarFacturasRM()');
  assert.equal(tabla(e, 'Líneas RM').length, 4);
  assert.equal(tabla(e, 'Facturas RM').length, 1);
});

test('Abonos: marcar en Piezas una pieza cuyo abono ya llegó en la factura convierte su fila "Sin solicitar" en "Abonada" (no añade otra)', () => {
  const e = entorno({ 'fra.pdf': facturaRaw() });
  ['FRA', 'FRAP', 'ENT', 'PROC'].forEach(id => e.mkFolder(id, id));
  const p = e.ss.getSheetByName('Piezas'), ab = e.ss.getSheetByName('Abonos');
  p.put(2, 1, false); p.put(2, 3, '443645'); p.put(2, 4, 'VARTAA8'); p.put(2, 5, 'A8 AGM'); p.put(2, 10, 130.10);  // sin marcar todavía
  subir(e, ['fra.pdf'], 'FRA');
  e.run('procesarFacturasRM()');
  const filas = () => [2, 3, 4, 5].map(r => ({ ref: ab.valor(r, 7), estado: ab.valor(r, 5), nota: ab.valor(r, 11), id: ab.valor(r, 13) })).filter(f => f.estado);
  const antes = filas();
  assert.deepEqual(antes.find(f => f.ref === 'VARTAA8'), { ref: 'VARTAA8', estado: 'Sin solicitar', nota: 'La pieza existe en Piezas pero no tiene Reembolso marcado', id: 1 });
  p.put(2, 1, true);
  e.ctx.__r = p.getRange(2, 1); e.run('alEditar({ range: __r })');
  const despues = filas();
  assert.equal(despues.length, antes.length, 'no se añade ninguna fila');
  assert.deepEqual(despues.find(x => x.ref === 'VARTAA8'), { ref: 'VARTAA8', estado: 'Abonada', nota: '', id: 1 });
  e.run('actualizarAbonos()');
  assert.equal(filas().length, antes.length, 'y sigue sin duplicarse');
});

test('Abonos es un informe: lo nuevo arriba; editar una pieza cambia su fila sin duplicarla; lo escrito a mano en Abonos se deshace', () => {
  const e = entorno({});
  const p = e.ss.getSheetByName('Piezas'), ab = e.ss.getSheetByName('Abonos');
  const editar = (fila, col, v) => { p.put(fila, col, v); e.ctx.__r = p.getRange(fila, col); e.run('alEditar({ range: __r })'); };
  const marcar = (fila, alb, ref, precio) => { p.put(fila, 3, alb); p.put(fila, 4, ref); p.put(fila, 5, 'Pieza ' + ref); p.put(fila, 10, precio); editar(fila, 1, true); };
  const refs = () => [2, 3, 4, 5].map(r => ab.valor(r, 7));
  marcar(2, '100', 'AAA', 10);
  marcar(3, '200', 'BBB', 20);
  assert.deepEqual(refs(), ['BBB', 'AAA', '', ''], 'la última pieza marcada va la primera');
  assert.equal(ab.valor(2, 13), 2, 'ID de la pieza en la columna oculta');
  assert.equal(ab.cell(2, 4).f, '=ROUND($C2*(1+IVA),2)');
  // Lo que se reportó: cambiar la referencia o el albarán de una pieza pedida añadía otra fila. Ahora cambia la suya.
  editar(3, 4, 'BBB2');
  editar(3, 3, '201');
  editar(3, 10, 25);
  assert.deepEqual(refs(), ['BBB2', 'AAA', '', '']);
  assert.deepEqual([ab.valor(2, 6), ab.valor(2, 3)], ['201', 25]);
  // Escribir en la tabla de Abonos no vale: se vuelve a poner lo que toca y se avisa
  ab.put(3, 5, 'Abonada'); ab.put(3, 11, 'Llamé a RM');
  e.ctx.__a = ab.getRange(3, 5); e.run('alEditar({ range: __a })');
  assert.deepEqual([ab.valor(3, 5), ab.valor(3, 11)], ['Sin abonar', '']);
  assert.match(e.log.toasts[e.log.toasts.length - 1], /Abonos se rellena solo/);
  // "Abonada a mano" en Piezas
  editar(2, 13, e.run('new Date(2026, 9, 20)'));
  assert.equal(ab.valor(e.run(`[2,3].find(r => hoja_('Abonos').getRange(r, 7).getValue() === 'AAA')`), 5), 'Abonada (a mano)');
  // Desmarcar en Piezas quita su fila
  editar(3, 1, false);
  assert.deepEqual(refs(), ['AAA', '', '', '']);
  assert.equal(p.valor(3, 12), '', 'sin fecha de reembolso');
  // Borrar la fila con "Eliminar fila" (trigger onChange): vuelve
  ab.deleteRow(2);
  e.ss.getActiveSheet = () => ab;
  e.ctx.__e = { changeType: 'REMOVE_ROW', source: e.ss }; e.run('alCambiar(__e)');
  assert.deepEqual(refs(), ['AAA', '', '', '']);
  assert.equal(p.valor(2, 1), true, 'la pieza sigue pedida');
  // Borrar la fila de la pieza en Piezas también la quita de Abonos
  p.deleteRow(2);
  e.ss.getActiveSheet = () => p;
  e.ctx.__e = { changeType: 'REMOVE_ROW', source: e.ss }; e.run('alCambiar(__e)');
  assert.deepEqual(refs(), ['', '', '', '']);
});

test('Piezas: cada pieza tiene un ID fijo; una fila copiada recibe uno nuevo y un ID borrado no se reutiliza', () => {
  const e = entorno({});
  const p = e.ss.getSheetByName('Piezas');
  const editar = (fila, col, v) => { p.put(fila, col, v); e.ctx.__r = p.getRange(fila, col); e.run('alEditar({ range: __r })'); };
  editar(2, 5, 'Uno'); editar(3, 5, 'Dos');
  assert.deepEqual([p.valor(2, 17), p.valor(3, 17)], [1, 2]);
  for (let c = 1; c <= 17; c++) if (p.cell(2, c)) p.put(4, c, p.valor(2, c));  // copiar y pegar la fila 2 en la 4
  e.ctx.__r = p.getRange(4, 1, 1, 17); e.run('alEditar({ range: __r })');
  assert.equal(p.valor(4, 17), 3);
  assert.equal(p.valor(2, 17), 1, 'la original conserva el suyo');
  p.deleteRow(4);
  editar(4, 5, 'Cuatro');
  assert.equal(p.valor(4, 17), 4, 'el 3 no se reutiliza');
});

test('Abonos: el abono de la factura convierte en Abonada la fila de su pieza; corregir el "ID pieza" en Líneas RM la mueve', () => {
  const e = entorno({ 'fra.pdf': facturaRaw(), 'fra2.pdf': facturaRaw() });
  ['FRA', 'FRAP', 'ENT', 'PROC'].forEach(id => e.mkFolder(id, id));
  const p = e.ss.getSheetByName('Piezas'), ab = e.ss.getSheetByName('Abonos'), l = e.ss.getSheetByName('Líneas RM');
  p.put(2, 3, '443645'); p.put(2, 4, 'VARTAA8'); p.put(2, 5, 'A8 AGM'); p.put(2, 10, 130.10); p.put(2, 1, true);
  e.ctx.__r = p.getRange(2, 1); e.run('alEditar({ range: __r })');
  // Otra pieza pedida del albarán 439139, con otra referencia y otro importe: el abono WALKE80477 no se le asigna solo
  p.put(3, 3, '439139'); p.put(3, 4, 'OTRA'); p.put(3, 5, 'ABRAZADERA'); p.put(3, 10, 7); p.put(3, 1, true);
  e.ctx.__r = p.getRange(3, 1); e.run('alEditar({ range: __r })');
  subir(e, ['fra.pdf'], 'FRA');
  e.run('procesarFacturasRM()');
  const filas = () => [2, 3, 4, 5].map(r => [ab.valor(r, 7), ab.valor(r, 5), ab.valor(r, 10)]).filter(f => f[1]);
  assert.deepEqual(filas(), [['OTRA', 'Sin abonar', ''], ['WALKE80477', 'Sin solicitar', 'FCR 00001'], ['VARTAA8', 'Abonada', 'FCR 00001']]);
  assert.ok(esFecha(ab.valor(4, 1)), 'fecha del abono');
  assert.equal(ab.valor(4, 8), '5678DEF', 'matrícula del albarán original');
  // Se corrige a mano: el abono WALKE80477 era de la pieza 2 (ABRAZADERA)
  const fila = 2 + tabla(e, 'Líneas RM').findIndex(x => x['Referencia'] === 'WALKE80477');
  l.put(fila, 14, 2); e.ctx.__l = l.getRange(fila, 14); e.run('alEditar({ range: __l })');
  assert.deepEqual(filas(), [['OTRA', 'Abonada', 'FCR 00001'], ['VARTAA8', 'Abonada', 'FCR 00001']]);
  assert.match(ab.valor(2, 11), /Importe abonado distinto: se pidió 7 sin IVA/);
  // Un ID que no existe: aviso
  l.put(fila, 14, 99); e.ctx.__l = l.getRange(fila, 14); e.run('alEditar({ range: __l })');
  assert.match(e.log.toasts[e.log.toasts.length - 1], /No hay ninguna pieza con ID 99/);
  l.put(fila, 14, 2); e.ctx.__l = l.getRange(fila, 14); e.run('alEditar({ range: __l })');
  // Reprocesar la factura no duplica nada y conserva la corrección
  subir(e, ['fra2.pdf'], 'FRA');
  e.run('procesarFacturasRM()');
  assert.deepEqual(filas(), [['OTRA', 'Abonada', 'FCR 00001'], ['VARTAA8', 'Abonada', 'FCR 00001']]);
});

test('Abonos: filas insertadas encima de la tabla o dentro se quitan; matrícula tal cual la pone RM; sin líneas SIGAUS', () => {
  const e = entorno({});
  const p = e.ss.getSheetByName('Piezas'), ab = e.ss.getSheetByName('Abonos'), l = e.ss.getSheetByName('Líneas RM');
  // Compra con un bastidor en el campo MATRICULA y su abono (con la línea de residuos)
  l.put(2, 1, 'FCR 9'); l.put(2, 2, '302751'); l.put(2, 4, 'ZFA2300000'); l.put(2, 5, 'Compra'); l.put(2, 7, 'DAYCO5PK1090'); l.put(2, 12, 9.24);
  [['DAYCO5PK1090', 'CORREA', -9.24], ['', 'SIGAUS (SIG. RD 679/2006)', -0.05]].forEach(([ref, desc, imp], k) => {
    const r = 3 + k; l.put(r, 1, 'FCR 9'); l.put(r, 2, '314672'); l.put(r, 3, e.run('new Date(2026, 5, 22)')); l.put(r, 5, 'Abono'); l.put(r, 6, '01000302751'); l.put(r, 7, ref); l.put(r, 8, desc); l.put(r, 12, imp);
  });
  e.run('actualizarAbonos()');
  assert.equal(ab.valor(1, 1), 'Fecha abono');
  assert.deepEqual([ab.valor(2, 7), ab.valor(2, 8), ab.valor(3, 2)], ['DAYCO5PK1090', 'ZFA2300000', ''], 'la matrícula tal cual la pone RM; la línea SIGAUS no sale');
  ab.insertRowsBefore(1, 1);
  e.ss.getActiveSheet = () => ab;
  e.ctx.__e = { changeType: 'INSERT_ROW', source: e.ss }; e.run('alCambiar(__e)');
  assert.deepEqual([ab.valor(1, 1), ab.valor(2, 7)], ['Fecha abono', 'DAYCO5PK1090'], 'cabecera de vuelta en la fila 1');
  ab.insertRowsBefore(2, 1);
  e.ctx.__e = { changeType: 'INSERT_ROW', source: e.ss }; e.run('alCambiar(__e)');
  assert.equal(ab.valor(2, 7), 'DAYCO5PK1090', 'la tabla se rehace sin el hueco');
  p.put(2, 3, '100'); p.put(2, 4, 'AAA'); p.put(2, 5, 'Pieza'); p.put(2, 10, 10); p.put(2, 1, true);
  e.ctx.__r = p.getRange(2, 1); e.run('alEditar({ range: __r })');
  assert.deepEqual([2, 3, 4].map(r => ab.valor(r, 7)), ['AAA', 'DAYCO5PK1090', '']);
});

test('diagnóstico lista problemas con enlaces', () => {
  const e = entorno({});
  const alb = e.ss.getSheetByName('Albaranes');
  alb.put(20, A('Matrícula'), 'ZZZ9999'); alb.put(20, A('Precio con IVA'), 10);
  e.mkFolder('ENT', 'ENT'); e.mkFolder('FRA', 'FRA');
  e.run('diagnostico()');
  const d = e.ss.getSheetByName('Diagnóstico');
  const textos = []; for (let r = 2; r < 12; r++) textos.push(d.valor(r, 4));
  assert.ok(textos.some(t => /ZZZ9999 no está en Coches/.test(t)), textos.join('|'));
  assert.ok(textos.some(t => /Sin nº de trabajo/.test(t)));
});

test('con hoja en español (es_ES) TODAS las fórmulas y reglas se escriben con ";"', () => {
  let e = crearEntorno({ privado: true, privadoRuta: PRIV, locale: 'es_ES', gemini: () => ({}) });
  e.run('setup()');
  const separadorIngles = f => /(?:^|[^\d]),|,(?:[^\d]|$)/.test(f.replace(/"[^"]*"/g, '""'));  // una coma que no sea decimal (entre dígitos)
  // Las fórmulas por fila ya no se escriben en filas vacías: se comprueban todas tal como se escribirían en la fila 2.
  const porFila = e.run("Object.keys(FORMULAS).flatMap(n => Object.keys(FORMULAS[n]).map(h => [n + ' ▸ ' + h, loc_(FORMULAS[n][h](2))]))");
  assert.ok(porFila.length >= 15);
  assert.ok(porFila.some(([k, f]) => k === 'Albaranes ▸ Quincena' && f === `=IF(${letra(A('Fecha albarán'))}2="";"";IF(DAY(${letra(A('Fecha albarán'))}2)<=15;1;2))`));
  for (const [k, f] of porFila) assert.ok(!separadorIngles(f), `${k}: ${f}`);
  for (const [hoja, col, fila] of [['Cuadre RM', 3, 4], ['Cuadre RM', 4, 4], ['Cuadre RM', 6, 4], ['Cuadre RM', 9, 5], ['Resumen', 2, 13], ['Resumen', 3, 13], ['Trabajos', 5, 3]]) {
    const f = e.ss.getSheetByName(hoja).cell(fila, col).f;
    assert.ok(f && !separadorIngles(f), `${hoja} ${fila},${col}: ${f}`);
  }
  e.run('cargarDatosIniciales()');
  assert.ok(!separadorIngles(e.ss.getSheetByName('Albaranes').cell(2, 14).f), 'y las filas añadidas también');
  e.mkFolder('ENT', 'E'); e.mkFolder('PROC', 'P');
  assert.equal(e.run("enlacePdf_({ getId: () => 'abc' })"), '=HYPERLINK("https://drive.google.com/file/d/abc/view";"Ver PDF")');
});
