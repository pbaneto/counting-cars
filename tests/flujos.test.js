const test = require('node:test');
const assert = require('node:assert/strict');
const { crearEntorno } = require('./gas-mock');

const PRIV = require('path').join(__dirname, 'fixtures-private.js');
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

test('setup crea pestañas, cabeceras, fórmulas y configuración', () => {
  const e = entorno({});
  const nombres = e.ss.getSheets().map(s => s.name);
  for (const n of ['Albaranes', 'Trabajos', 'Piezas', 'Abonos', 'Coches', 'Resumen', 'Facturas RM', 'Líneas RM', 'Config', 'Registro']) assert.ok(nombres.includes(n), n);
  assert.ok(!nombres.includes('Hoja 1'));
  const alb = e.ss.getSheetByName('Albaranes');
  assert.equal(alb.valor(1, 6), 'Nº albarán');
  assert.equal(alb.valor(1, 3), 'Mes', 'columna nueva "Mes" junto a "Fecha albarán"');
  assert.match(alb.cell(2, 3).f, /CHOOSE\(MONTH\(/);
  assert.match(alb.cell(2, 4).f, /^=IF\(B2=""/);          // Quincena
  assert.match(alb.cell(2, 10).f, /SUMIFS\(Piezas!/);       // Precio - abonos
  assert.match(alb.cell(2, 14).f, /DIAS_AVISO/);           // Avisos en las filas con datos (piloto: filas 2-4)
  assert.equal(alb.cell(5, 14), undefined, 'sin fórmulas en filas vacías: leerlas y recalcularlas era lo lento');
  assert.equal(alb.getLastRow(), 4);
  // Casillas sólo en filas con datos: una casilla vacía vale FALSE y haría leer miles de filas
  const trab = e.ss.getSheetByName('Trabajos');
  assert.equal(trab.getLastRow(), 6, 'Trabajos: panel (filas 1-3) + cabecera (fila 4) + 2 trabajos del piloto (5-6)');
  assert.equal(e.ss.getSheetByName('Piezas').getLastRow(), 1, 'Piezas vacía: sin casillas "Reembolso" por debajo');
  assert.equal(trab.valor(6, 14), false, 'la fila nueva lleva su casilla (FALSE = sin pagar), columna N');
  // Cabecera real en la fila 4 (el panel ocupa la 1-3), datos desde la 5
  assert.equal(trab.valor(4, 3), 'Mes', 'columna nueva "Mes" junto a "Fecha apertura"');
  assert.equal(trab.valor(4, 4), 'Quincena', 'Quincena a la derecha de Mes');
  assert.equal(trab.valor(5, 4), 1, 'rellenada con la fecha de apertura del piloto (15/09 → 1ª quincena)');
  assert.equal(trab.cell(5, 4).f, '', 'Quincena es un valor, no una fórmula: se puede cambiar a mano');
  assert.equal(trab.valor(4, 5), 'Matrícula');
  assert.match(trab.cell(5, 3).f, /CHOOSE\(MONTH\(/);
  assert.match(trab.cell(5, 10).f, /SUMIFS\(Albaranes!.*"RM"\)/);      // Recambios facturables RM
  assert.match(trab.cell(5, 11).f, /SUMIFS\(Albaranes!.*"Otros"\)/);  // Recambios facturables Otros
  // Panel "Resumen (según filtro)" encima de la cabecera: título (1), etiquetas (2), valores (3)
  assert.equal(trab.valor(1, 1), 'Resumen (según filtro)');
  assert.equal(trab.valor(2, 1), 'Trabajos'); assert.equal(trab.valor(2, 5), 'Morosos');
  assert.match(trab.cell(3, 5).f, /SUMPRODUCT/);
  // Ninguna columna calculada (gris) se queda con una validación residual; las de entrada no se tocan
  assert.deepEqual(new Set(trab.validacionesLimpiadas), new Set([1, 3, 6, 7, 8, 9, 10, 11, 13, 15]),
    'Mes, Coche, Cliente, Recambios, Recambios facturables, Recambios facturables RM/Otros, Beneficio, Avisos y la fila de valores del panel');
  const cfg = tabla(e, 'Config');
  assert.equal(cfg.find(x => x.Clave === 'CARPETA_ENTRADA').Valor, 'ENT');
  assert.equal(cfg.find(x => x.Clave === 'MODELO_GEMINI').Valor, 'gemini-3.5-flash-lite');
  const abonos = e.ss.getSheetByName('Abonos');
  assert.equal(abonos.valor(1, 1), 'Fecha abono', 'la tabla grande arriba del todo');
  assert.match(abonos.cell(4, 17).f, /SUMIFS\(Albaranes!/);  // resumen por quincena a la derecha (columna Q)
  // Fechas de cada quincena dentro de la fórmula (sin columnas Desde/Hasta): ene 1ª = 1-15, ene 2ª = 16-fin de mes
  assert.match(abonos.cell(4, 17).f, /">="&DATE\(\$P\$1,1,1\).*"<="&DATE\(\$P\$1,1,15\)/);
  assert.match(abonos.cell(5, 18).f, /SUMIFS\(\$D:\$D,\$E:\$E,"Abonada",\$A:\$A,">="&DATE\(\$P\$1,1,16\).*"<="&EOMONTH\(DATE\(\$P\$1,1,1\),0\)/);
  assert.match(abonos.cell(4, 20).f, /^=IF\(S4="","",ROUND\(S4-\(Q4-R4\),2\)\)$/); // Diferencia informativa, sin aviso
  assert.equal(abonos.valor(1, 15), 'Año'); assert.equal(abonos.valor(1, 16), 2026);
  assert.equal(abonos.cell(4, 23).f, '=DIAS_AVISO_REEMB');  // panel: umbral copiado de Config a esta pestaña
  assert.match(abonos.cell(5, 23).f, /COUNTIFS.*\$W\$4/);    // panel: fuera de plazo, contra el umbral de la misma pestaña
  e.run('repararFormulas()');
  const errores = e.log.console.filter(m => /^\[ERROR\]/.test(m));
  assert.deepEqual(errores, [], 'setup y repararFormulas deben terminar sin errores');
});

test('repararFormulas quita las fórmulas de las filas vacías (hoja antigua) sin tocar datos ni valores a mano', () => {
  const e = entorno({});
  const alb = e.ss.getSheetByName('Albaranes'), p = e.ss.getSheetByName('Piezas');
  // Como la hoja real antes del cambio: fórmulas rellenadas por adelantado muy por debajo de los datos
  for (let r = 5; r <= 1501; r++) { alb.put(r, 4, `=IF(B${r}="","",1)`); alb.put(r, 14, `=IF(G${r}="","",1)`); }
  p.put(4001, 15, '=IF(C4001="","",1)');
  for (let r = 2; r <= 4001; r++) p.put(r, 1, false);   // casillas "Reembolso" rellenadas hasta la 4001 (valen FALSE)
  const trab = e.ss.getSheetByName('Trabajos');
  // El panel ocupa 1-3, la cabecera la 4, los 2 trabajos del piloto la 5-6: el relleno de sobra empieza en la 7
  for (let r = 7; r <= 801; r++) trab.put(r, 14, false); // "Pagado" (columna N) hasta la 801
  alb.put(900, 11, 'escrito a mano');                 // valor (no fórmula) en la columna calculada "Coche"
  e.run('repararFormulas()');
  assert.equal(alb.cell(700, 4), undefined, 'Quincena vacía por debajo de los datos');
  assert.equal(alb.cell(1501, 14), undefined, 'Avisos vacía por debajo de los datos');
  assert.equal(p.cell(4001, 15), undefined);
  assert.match(alb.cell(4, 14).f, /DIAS_AVISO/, 'las filas con datos conservan sus fórmulas');
  assert.equal(alb.valor(900, 11), 'escrito a mano', 'no borra un valor escrito a mano');
  assert.ok(e.log.console.some(l => /\[AVISO\] repararFormulas Albaranes!K900/.test(l)));
  assert.equal(p.getLastRow(), 1);
  assert.equal(trab.getLastRow(), 6, 'Trabajos queda con sus 2 filas de datos (panel 1-3, cabecera 4, datos 5-6)');
  assert.equal(trab.valor(5, 14), true, 'el Pagado de las filas con datos no se toca');
});

test('repararFormulas conserva el año elegido en Abonos', () => {
  const e = entorno({});
  const ab = e.ss.getSheetByName('Abonos');
  ab.put(1, 16, 2027);
  e.run('repararFormulas()');
  assert.equal(ab.valor(1, 16), 2027);
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
  const buena = trab.cell(5, 13).f;
  trab.put(5, 13, '=1');  // alguien ha pisado la fórmula de Beneficio
  escritas.length = 0;
  e.run('repararFormulas()');
  assert.equal(trab.cell(5, 13).f, buena, 'la fórmula pisada se restaura');
  assert.ok(enTablas().every(x => x.startsWith('Trabajos!') && x.endsWith(',13')), 'sólo se reescribe esa columna');
});

test('repararFormulas añade a Config las claves nuevas sin pisar los valores existentes', () => {
  const e = entorno({});
  const cfg = e.ss.getSheetByName('Config');
  const fila = clave => e.run(`leerTabla_('Config').filas.find(f => f.v['Clave'] === '${clave}').fila`);
  const r = fila('DIAS_AVISO_REEMBOLSO');
  [1, 2, 3].forEach(c => cfg.put(r, c, ''));            // hoja montada antes de existir la clave
  cfg.put(fila('DIAS_AVISO_TRABAJO'), 2, 20);           // valor cambiado a mano por el usuario
  e.run('repararFormulas()');
  const t = tabla(e, 'Config');
  assert.equal(t.find(x => x.Clave === 'DIAS_AVISO_REEMBOLSO').Valor, 45);
  assert.equal(t.find(x => x.Clave === 'DIAS_AVISO_TRABAJO').Valor, 20);
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

test('repararTrabajos: recupera Matrícula cruzando con Albaranes y Factura/Pagado desde el piloto', () => {
  const e = entorno({});
  e.run('repararTrabajos()');
  const trab = tabla(e, 'Trabajos');
  const j1 = trab.find(t => t['Nº trabajo'] === '1GHJ-1'), j2 = trab.find(t => t['Nº trabajo'] === '8DEF-1');
  assert.equal(j1['Matrícula'], '4321GHJ');
  assert.equal(j1['Pagado'], true);
  assert.equal(j1['Factura'], 500);
  assert.equal(j2['Matrícula'], '5678DEF');
  assert.equal(j2['Pagado'], false);
  assert.equal(j2['Factura'], '');
});

test('procesarAlbaranes: nuevo, segundo escaneo con R, duplicado sin R y no-albarán', () => {
  const R = {
    'a1.pdf': albaranRaw(),
    'a2.pdf': albaranRaw({ lineas: albaranRaw().lineas.map((l, i) => Object.assign({}, l, { reembolso: i === 0 })) }),
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
  // Drive: 3 en Procesados, el no-albarán en Errores
  assert.equal(e.carpetas.PROC.ficheros.length, 3);
  assert.equal(e.carpetas.ENT.ficheros.length, 0);
  const errores = Object.values(e.carpetas).find(f => f.nombre === 'Errores');
  assert.equal(errores.ficheros.length, 1);
  const reg = tabla(e, 'Registro').map(r => r['Mensaje']).join('\n');
  assert.match(reg, /duplicado/);
  assert.match(reg, /no parece un albarán/);
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
  const bien = albaranRaw({ numero_albaran: '700002' });
  const malo3 = albaranRaw({ numero_albaran: '700003', total: 50 });
  const e = entorno({ 'c1.pdf': [malo, albaranRaw({ numero_albaran: '700001' })], 'c2.pdf': [malo3, malo3], 'c3.pdf': bien });
  ['ENT', 'PROC'].forEach(id => e.mkFolder(id, id));
  subir(e, ['c1.pdf', 'c2.pdf', 'c3.pdf'], 'ENT');
  e.run('procesarAlbaranes()');
  const nums = tabla(e, 'Albaranes').map(a => a['Nº albarán']);
  assert.ok(nums.includes('700001'), 'c1: la segunda lectura cuadra y se usa');
  assert.ok(nums.includes('700002'));
  assert.ok(!nums.includes('700003'), 'c2 no se añade');
  assert.deepEqual(e.peticiones.filter(p => p.pista).map(p => p.nombre).sort(), ['c1.pdf', 'c2.pdf'], 'segunda lectura sólo de los que no cuadran');
  assert.match(e.peticiones.find(p => p.pista).pista, /no cuadraban/);
  assert.deepEqual(e.carpetas.ENT.ficheros.map(f => f.nombre), ['c2.pdf'], 'el que sigue sin cuadrar se queda en Entrada');
  assert.equal(e.carpetas.PROC.ficheros.length, 2);
  const aviso = e.log.alerts[e.log.alerts.length - 1][1];
  assert.match(aviso, /No cuadran los importes, siguen en Entrada \(1\):\n• c2\.pdf: base 36,81 \+ IVA = 44,54, pero el total es 50/);
});

test('procesarAlbaranes: un albarán de ABONO no va a Albaranes; marca Abonada la pieza pedida (fecha del abono) y lo demás va a Abonos', () => {
  const abono = albaranRaw({ numero_albaran: '486962', fecha: '2026-09-28', es_abono: true, total: -72.94, base_imponible: -60.28, matricula: '',
    lineas: [
      { referencia: 'BOSCHF026400517', descripcion: 'S0517 FILTRO', importe: -15.14, reembolso: false, albaran_origen: '01000478440' },
      { referencia: 'ELEVA05', descripcion: 'ELEVALUNAS', importe: -45.14, reembolso: false, albaran_origen: '01000476000' },
    ] });
  const e = entorno({ 'ab.pdf': abono });
  ['ENT', 'PROC'].forEach(id => e.mkFolder(id, id));
  const p = e.ss.getSheetByName('Piezas'), ab = e.ss.getSheetByName('Abonos');
  p.put(2, 3, '478440'); p.put(2, 4, 'BOSCHF026400517'); p.put(2, 5, 'FILTRO AIRE'); p.put(2, 10, 15.14); p.put(2, 1, true);
  e.ctx.__r = p.getRange(2, 1); e.run('alEditar({ range: __r })');
  assert.equal(ab.valor(2, 5), 'Sin abonar');
  subir(e, ['ab.pdf'], 'ENT');
  e.run('procesarAlbaranes()');
  assert.ok(!tabla(e, 'Albaranes').some(a => a['Nº albarán'] === '486962'), 'no va a Albaranes');
  assert.equal(e.carpetas.PROC.ficheros.length, 1);
  assert.match(e.log.alerts[e.log.alerts.length - 1][1], /Albaranes de abono \(a la pestaña Abonos\): 1/);
  const filas = [2, 3, 4].map(r => ({ desc: ab.valor(r, 2), estado: ab.valor(r, 5), fecha: e.run(`aISO_(hoja_('Abonos').getRange(${r}, 1).getValue())`), fra: ab.valor(r, 10) }));
  const pedida = filas.find(f => f.desc === 'FILTRO AIRE');
  assert.deepEqual([pedida.estado, pedida.fecha, pedida.fra], ['Abonada', '2026-09-28', ''], 'pieza pedida → Abonada con la fecha del albarán de abono');
  const otra = filas.find(f => f.desc === 'ELEVALUNAS');
  assert.equal(otra.estado, 'Sin solicitar');
  assert.equal(filas.filter(f => f.desc).length, 2);
  // Llega la factura quincenal con ese abono: no se duplica, sólo se rellena "Factura RM"
  const l = e.ss.getSheetByName('Líneas RM'), fr = e.ss.getSheetByName('Facturas RM');
  fr.put(2, 1, 'FCR 7'); fr.put(2, 2, e.run('new Date(2026, 8, 30)'));
  [['BOSCHF026400517', '01000478440', -15.14], ['ELEVA05', '01000476000', -45.14], ['OTRA1', '01000470000', -5]].forEach(([ref, orig, imp], k) => {
    const r = 2 + k; l.put(r, 1, 'FCR 7'); l.put(r, 2, '486962'); l.put(r, 3, e.run('new Date(2026, 8, 28)')); l.put(r, 5, 'Abono'); l.put(r, 6, orig); l.put(r, 7, ref); l.put(r, 8, 'X'); l.put(r, 12, imp);
  });
  e.run('actualizarAbonos()');
  const todas = [2, 3, 4, 5].map(r => ({ ref: ab.valor(r, 7), estado: ab.valor(r, 5), fecha: e.run(`aISO_(hoja_('Abonos').getRange(${r}, 1).getValue())`), fra: ab.valor(r, 10) }));
  assert.equal(todas.filter(f => f.ref).length, 3, 'sólo se añade la línea que no venía en el albarán escaneado');
  assert.deepEqual(todas.find(f => f.ref === 'BOSCHF026400517'), { ref: 'BOSCHF026400517', estado: 'Abonada', fecha: '2026-09-28', fra: 'FCR 7' });
  assert.deepEqual(todas.find(f => f.ref === 'OTRA1'), { ref: 'OTRA1', estado: 'Sin solicitar', fecha: '2026-09-30', fra: 'FCR 7' }, 'por la factura: fecha de la factura');
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
  const formulas = alb.reglas.map(r => r.formula);
  assert.equal(formulas.length, new Set(formulas).size, 'sin reglas nuestras duplicadas');
});

test('versiones: al subir la versión se ejecutan las migraciones pendientes una vez y se reaplica el diseño', () => {
  const e = entorno({});
  assert.equal(e.props.VERSION_HOJA, e.run('VERSION'), 'setup apunta la versión');
  const hechas = [];
  e.ctx.__hechas = hechas;
  e.run(`MIGRACIONES['0.9.5'] = () => __hechas.push('0.9.5'); MIGRACIONES['1.0.0'] = () => __hechas.push('1.0.0'); MIGRACIONES['9.0.0'] = () => __hechas.push('9.0.0');`);
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

test('escribir a mano en una columna automática vuelve a poner su fórmula y avisa', () => {
  const e = entorno({});
  const alb = e.ss.getSheetByName('Albaranes');
  const buena = alb.cell(2, 11).f;  // Coche
  alb.put(2, 11, 'Otro coche');
  e.ctx.__r = alb.getRange(2, 11); e.run('alEditar({ range: __r })');
  assert.equal(alb.cell(2, 11).f, buena);
  assert.ok(e.log.toasts.some(t => /"Coche" se rellena sola/.test(t)));
});

test('Trabajos: Quincena se rellena sola con la fecha de apertura y se puede cambiar a mano', () => {
  const e = entorno({});
  const trab = e.ss.getSheetByName('Trabajos');
  trab.put(7, 2, e.run('new Date(2026, 8, 20)')); trab.put(7, 5, '1234ABC');
  e.ctx.__r = trab.getRange(7, 5); e.run('alEditar({ range: __r })');
  assert.equal(trab.valor(7, 4), 2, '20/09 → 2ª quincena');
  trab.put(7, 4, 1);
  e.ctx.__r = trab.getRange(7, 4); e.run('alEditar({ range: __r })');
  assert.equal(trab.valor(7, 4), 1, 'el cambio a mano se respeta');
  e.run('repararFormulas()');
  assert.equal(trab.valor(7, 4), 1);
});

test('edición manual en Albaranes: fecha, proveedor y trabajo automáticos; NUEVO abre otro trabajo', () => {
  const e = entorno({});
  const alb = e.ss.getSheetByName('Albaranes');
  const fila = 10;
  alb.put(fila, 8, '1234 abc'); alb.put(fila, 9, 44.54);
  const rango = alb.getRange(fila, 8, 1, 2);
  e.ctx.__rango = rango;
  e.run('alEditar({ range: __rango })');
  const t = tabla(e, 'Albaranes').find(a => a['Matrícula'] === '1234ABC');
  assert.ok(t, 'fila procesada');
  assert.ok(esFecha(t['Fecha escaneo']));
  assert.ok(esFecha(t['Fecha albarán']));
  assert.equal(t['Proveedor'], 'RM');
  assert.equal(t['Nº trabajo'], '4ABC-1');
  // Segunda fila del mismo coche -> mismo trabajo abierto
  alb.put(11, 8, '1234ABC'); alb.put(11, 9, 20);
  e.ctx.__r2 = alb.getRange(11, 8, 1, 2); e.run('alEditar({ range: __r2 })');
  assert.equal(tabla(e, 'Albaranes').filter(a => a['Nº trabajo'] === '4ABC-1').length, 2);
  // NUEVO -> segundo trabajo abierto en paralelo
  alb.put(11, 7, 'NUEVO');
  e.ctx.__r3 = alb.getRange(11, 7); e.run('alEditar({ range: __r3 })');
  assert.equal(tabla(e, 'Albaranes').find(a => a['Precio con IVA'] === 20)['Nº trabajo'], '4ABC-2');
  assert.equal(tabla(e, 'Trabajos').filter(x => x['Matrícula'] === '1234ABC').length, 2);
});

test('Matrícula en Albaranes: busca por cualquier combinación de letras o dígitos, no sólo al principio', () => {
  const e = entorno({});  // Coches del piloto: 1234ABC, 4321GHJ, 5678DEF
  const alb = e.ss.getSheetByName('Albaranes');

  // Coincidencia única por las LETRAS (van al final: el desplegable nativo de Sheets sólo busca desde el principio)
  alb.put(10, 8, 'ghj');
  e.ctx.__unica = alb.getRange(10, 8); e.run('alEditar({ range: __unica })');
  assert.equal(alb.valor(10, 8), '4321GHJ', 'se autocompleta con la única matrícula que contiene "GHJ"');
  assert.ok(e.log.toasts.some(t => t.includes('"GHJ" → 4321GHJ')), 'avisa de qué matrícula ha puesto');

  // Texto demasiado corto (< 3): no se busca, se deja tal cual y sigue el aviso habitual de "no está en Coches"
  // (ese aviso sólo se comprueba cuando la fila ya tiene precio, igual que sin este cambio)
  e.log.toasts.length = 0;
  alb.put(11, 8, '4'); alb.put(11, 9, 44.54);
  e.ctx.__corto = alb.getRange(11, 8, 1, 2); e.run('alEditar({ range: __corto })');
  assert.equal(alb.valor(11, 8), '4', 'texto demasiado corto: se deja tal cual, sin adivinar');
  assert.equal(e.log.toasts.length, 1);
  assert.match(e.log.toasts[0], /no está en la pestaña Coches/);

  // Coincide con varias: no adivina, deja lo escrito y lista las candidatas (se añade un coche a propósito)
  e.log.toasts.length = 0;
  const coches = e.ss.getSheetByName('Coches');
  const filaNueva = coches.getLastRow() + 1;
  coches.put(filaNueva, 1, '4321GHK'); coches.put(filaNueva, 2, 'Otro'); coches.put(filaNueva, 3, 'Otro coche');
  alb.put(12, 8, '432');
  e.ctx.__varias = alb.getRange(12, 8); e.run('alEditar({ range: __varias })');
  assert.equal(alb.valor(12, 8), '432', 'coincide con varias: no adivina, deja lo escrito');
  assert.equal(e.log.toasts.length, 1);
  assert.match(e.log.toasts[0], /"432" coincide con 2 matrículas: 4321GHJ, 4321GHK/);

  // Matrícula ya exacta: no hay búsqueda ni aviso de "completada"
  e.log.toasts.length = 0;
  alb.put(13, 8, '5678DEF');
  e.ctx.__exacta = alb.getRange(13, 8); e.run('alEditar({ range: __exacta })');
  assert.equal(alb.valor(13, 8), '5678DEF');
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
    'leer Albaranes \\(3 de 3 filas\\).*leer Abonos \\(0 filas, 0 claves vistas\\).*revisar 1 fila\\(s\\) editada\\(s\\).*sincronizar Abonos.*guardar cambios y recalcular \\d+ms'));
  assert.match(linea('⏱ escribirAbonos_'), /preparar piezas \(1 marcadas.*cruce .*escribir \(1 nuevas, 0 completadas, 0 borradas\).*Sin abonar 1/);
  assert.match(linea('■ alEditar'), /^■ alEditar: \d+ms \(escribir en Registro \d+ms\)$/);
  assert.equal(e.log.console.filter(l => /^\[INFO\] reconstruirAbonos/.test(l)).length, 0, 'el resumen ya no se escribe en Registro');
  assert.equal(p.valor(2, 1), true);
  assert.ok(esFecha(p.valor(2, 12)), 'fecha de reembolso');
  assert.equal(p.valor(2, 14), 'Manual');
  assert.match(p.cell(2, 15).f, /Falta el nº de albarán/, 'la fila escrita a mano recibe sus fórmulas');
  assert.equal(e.ss.getSheetByName('Abonos').valor(2, 5), 'Sin abonar');
  // Pieza de un albarán de proveedor "Otros": no se reclama a RM
  const alb = e.ss.getSheetByName('Albaranes');
  alb.put(5, 5, 'Otros'); alb.put(5, 6, '999'); alb.put(5, 8, '1234ABC'); alb.put(5, 9, 20);
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

test('Abonos editable: lo nuevo va arriba y no pisa lo cambiado a mano; borrar una fila desmarca la pieza', () => {
  const e = entorno({});
  const p = e.ss.getSheetByName('Piezas'), ab = e.ss.getSheetByName('Abonos');
  const marcar = (fila, alb, ref, precio) => {
    p.put(fila, 3, alb); p.put(fila, 4, ref); p.put(fila, 5, 'Pieza ' + ref); p.put(fila, 10, precio); p.put(fila, 1, true);
    e.ctx.__r = p.getRange(fila, 1); e.run('alEditar({ range: __r })');
  };
  marcar(2, '100', 'AAA', 10);
  marcar(3, '200', 'BBB', 20);
  assert.equal(ab.valor(2, 7), 'BBB', 'la última pieza marcada va la primera');
  assert.equal(ab.valor(3, 7), 'AAA');
  assert.equal(ab.valor(2, 13), 'P|200|BBB|1', 'clave en la columna oculta');
  assert.equal(ab.cell(2, 4).f, '=ROUND($C2*(1+IVA),2)');
  // Cambios a mano: se respetan en la siguiente sincronización
  ab.put(3, 5, 'Abonada'); ab.put(3, 11, 'Llamé a RM');
  ab.put(4, 2, 'Fila escrita a mano'); ab.put(4, 5, 'Sin solicitar');
  marcar(4, '300', 'CCC', 30);
  assert.deepEqual([2, 3, 4, 5].map(r => ab.valor(r, 7)), ['CCC', 'BBB', 'AAA', '']);
  assert.equal(ab.valor(4, 5), 'Abonada');
  assert.equal(ab.valor(4, 11), 'Llamé a RM');
  assert.equal(ab.valor(5, 2), 'Fila escrita a mano');
  e.run('actualizarAbonos()');
  assert.equal(ab.valor(6, 2), '', 'actualizar no duplica nada');

  // Desmarcar en Piezas borra sólo su fila
  p.put(3, 1, false); e.ctx.__r = p.getRange(3, 1); e.run('alEditar({ range: __r })');
  assert.deepEqual([2, 3, 4].map(r => ab.valor(r, 7)), ['CCC', 'AAA', '']);
  assert.equal(ab.valor(3, 5), 'Abonada', 'la fila cambiada a mano sigue igual');

  // Vaciar una fila en Abonos (seleccionarla y Suprimir) = borrarla: se quita el hueco y se desmarca la pieza
  for (let c = 1; c <= 13; c++) ab.put(2, c, '');
  e.ctx.__a = ab.getRange(2, 1, 1, 13); e.run('alEditar({ range: __a })');
  assert.equal(p.valor(4, 1), false, 'CCC desmarcada en Piezas');
  assert.equal(p.valor(4, 12), '', 'y sin fecha de reembolso');
  assert.equal(ab.valor(2, 7), 'AAA', 'sin hueco');

  // Borrar la fila con "Eliminar fila" (trigger onChange)
  ab.deleteRow(2);
  e.ss.getActiveSheet = () => ab;
  e.ctx.__e = { changeType: 'REMOVE_ROW', source: e.ss }; e.run('alCambiar(__e)');
  assert.equal(p.valor(2, 1), false, 'AAA desmarcada en Piezas');
  assert.equal(p.valor(3, 1), false);
  assert.equal(ab.valor(2, 2), 'Fila escrita a mano', 'la fila manual no afecta a nada');
  // Borrar una fila entera también movía el resumen de la derecha: se ha vuelto a montar en su sitio
  assert.equal(ab.valor(3, 15), 'Mes');
  assert.equal(ab.valor(1, 16), 2026);
  assert.match(ab.cell(4, 17).f, /SUMIFS\(Albaranes!/);
  assert.equal(ab.valor(27, 16), 2, 'dic, 2ª quincena, en la fila 27');
});

test('Abonos editable: el abono de la factura completa la fila de su pieza; una fila de abono borrada no vuelve', () => {
  const e = entorno({ 'fra.pdf': facturaRaw(), 'fra2.pdf': facturaRaw() });
  ['FRA', 'FRAP', 'ENT', 'PROC'].forEach(id => e.mkFolder(id, id));
  const p = e.ss.getSheetByName('Piezas'), ab = e.ss.getSheetByName('Abonos');
  p.put(2, 3, '443645'); p.put(2, 4, 'VARTAA8'); p.put(2, 5, 'A8 AGM'); p.put(2, 10, 130.10); p.put(2, 1, true);
  e.ctx.__r = p.getRange(2, 1); e.run('alEditar({ range: __r })');
  ab.put(2, 11, 'pedido por teléfono');
  subir(e, ['fra.pdf'], 'FRA');
  e.run('procesarFacturasRM()');
  // Arriba el "Sin solicitar" nuevo; debajo, la fila de la pieza completada con el abono y su nota intacta
  assert.equal(ab.valor(2, 5), 'Sin solicitar');
  assert.equal(ab.valor(3, 5), 'Abonada');
  assert.equal(ab.valor(3, 10), 'FCR 00001');
  assert.ok(esFecha(ab.valor(3, 1)), 'fecha del abono');
  assert.equal(ab.valor(3, 11), 'pedido por teléfono');
  assert.equal(ab.valor(4, 5), '');
  // Matrícula vacía en una fila ya existente: se rellena al sincronizar (sin pisar una escrita a mano)
  ab.put(3, 8, ''); ab.put(2, 8, 'A MANO');
  e.run('actualizarAbonos()');
  assert.equal(ab.valor(3, 8), '5678DEF');
  assert.equal(ab.valor(2, 8), 'A MANO');
  // Se borra a mano la fila "Sin solicitar" y se reprocesa la factura: no vuelve y Piezas no cambia
  ab.deleteRow(2);
  e.ss.getActiveSheet = () => ab;
  e.ctx.__e = { changeType: 'REMOVE_ROW', source: e.ss }; e.run('alCambiar(__e)');
  assert.equal(p.valor(2, 1), true, 'borrar una fila sin pieza no desmarca nada');
  subir(e, ['fra2.pdf'], 'FRA');
  e.run('procesarFacturasRM()');
  assert.equal(ab.valor(2, 5), 'Abonada');
  assert.equal(ab.valor(3, 5), '', 'la fila borrada no reaparece');
});

test('Abonos: tabla arriba y resumen a la derecha; filas insertadas encima se quitan; matrícula tal cual la pone RM', () => {
  const e = entorno({});
  const p = e.ss.getSheetByName('Piezas'), ab = e.ss.getSheetByName('Abonos'), l = e.ss.getSheetByName('Líneas RM');
  // Compra con un bastidor en el campo MATRICULA y su abono
  l.put(2, 1, 'FCR 9'); l.put(2, 2, '302751'); l.put(2, 4, 'ZFA2300000'); l.put(2, 5, 'Compra'); l.put(2, 7, 'DAYCO5PK1090'); l.put(2, 12, 9.24);
  l.put(3, 1, 'FCR 9'); l.put(3, 2, '314672'); l.put(3, 3, e.run('new Date(2026, 5, 22)')); l.put(3, 5, 'Abono'); l.put(3, 6, '01000302751'); l.put(3, 7, 'DAYCO5PK1090'); l.put(3, 8, 'CORREA'); l.put(3, 12, -9.24);
  e.run('actualizarAbonos()');
  assert.equal(ab.valor(1, 1), 'Fecha abono');
  assert.equal(ab.valor(2, 7), 'DAYCO5PK1090');
  assert.equal(ab.valor(2, 8), 'ZFA2300000', 'la matrícula tal cual la pone RM');
  assert.equal(ab.valor(3, 15), 'Mes', 'el resumen no se mueve al insertar filas en la tabla');
  assert.equal(ab.valor(4, 16), 1);
  // Alguien inserta una fila entera encima de la cabecera: se quita y el resumen vuelve a su sitio
  ab.insertRowsBefore(1, 1);
  e.ss.getActiveSheet = () => ab;
  e.ctx.__e = { changeType: 'INSERT_ROW', source: e.ss }; e.run('alCambiar(__e)');
  assert.equal(ab.valor(1, 1), 'Fecha abono', 'cabecera de vuelta en la fila 1');
  assert.equal(ab.valor(2, 7), 'DAYCO5PK1090');
  assert.equal(ab.valor(1, 15), 'Año'); assert.equal(ab.valor(3, 15), 'Mes');
  // Una fila entera insertada DENTRO de la tabla: los datos se quedan, el resumen se vuelve a montar en su sitio
  ab.insertRowsBefore(2, 1);
  e.ctx.__e = { changeType: 'INSERT_ROW', source: e.ss }; e.run('alCambiar(__e)');
  assert.equal(ab.valor(3, 7), 'DAYCO5PK1090');
  assert.equal(ab.valor(3, 15), 'Mes'); assert.equal(ab.valor(27, 16), 2); assert.equal(ab.valor(28, 16), '');
  // Marcar una pieza: la fila nueva entra arriba desplazando sólo la tabla
  p.put(2, 3, '100'); p.put(2, 4, 'AAA'); p.put(2, 5, 'Pieza'); p.put(2, 10, 10); p.put(2, 1, true);
  e.ctx.__r = p.getRange(2, 1); e.run('alEditar({ range: __r })');
  assert.deepEqual([2, 3, 4].map(r => ab.valor(r, 7)), ['AAA', '', 'DAYCO5PK1090']);
  assert.equal(ab.valor(3, 15), 'Mes'); assert.equal(ab.valor(4, 16), 1);
});

test('diagnóstico lista problemas con enlaces', () => {
  const e = entorno({});
  const alb = e.ss.getSheetByName('Albaranes');
  alb.put(20, 8, 'ZZZ9999'); alb.put(20, 9, 10);
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
  assert.ok(porFila.some(([k, f]) => k === 'Albaranes ▸ Quincena' && f === '=IF(B2="";"";IF(DAY(B2)<=15;1;2))'));
  for (const [k, f] of porFila) assert.ok(!separadorIngles(f), `${k}: ${f}`);
  for (const [hoja, col, fila] of [['Abonos', 17, 4], ['Abonos', 18, 4], ['Abonos', 20, 4], ['Abonos', 23, 3], ['Resumen', 2, 13], ['Resumen', 3, 13], ['Trabajos', 5, 3]]) {
    const f = e.ss.getSheetByName(hoja).cell(fila, col).f;
    assert.ok(f && !separadorIngles(f), `${hoja} ${fila},${col}: ${f}`);
  }
  e.run('cargarDatosIniciales()');
  assert.ok(!separadorIngles(e.ss.getSheetByName('Albaranes').cell(2, 14).f), 'y las filas añadidas también');
  e.mkFolder('ENT', 'E'); e.mkFolder('PROC', 'P');
  assert.equal(e.run("enlacePdf_({ getId: () => 'abc' })"), '=HYPERLINK("https://drive.google.com/file/d/abc/view";"Ver PDF")');
});
