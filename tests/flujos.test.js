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
    const nombre = e.actual.shift();
    const r = respuestas[nombre];
    return { code: 200, body: JSON.stringify({ candidates: [{ content: { parts: [{ text: JSON.stringify(r) }] }, finishReason: 'STOP' }] }) };
  } });
  e.actual = [];
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
  assert.equal(alb.valor(1, 5), 'Nº albarán');
  assert.match(alb.cell(2, 3).f, /^=IF\(B2=""/);          // Quincena
  assert.match(alb.cell(2, 9).f, /SUMIFS\(Piezas!/);       // Precio facturable
  assert.match(alb.cell(1500, 13).f, /DIAS_AVISO/);        // Avisos hasta la fila 1500
  const cfg = tabla(e, 'Config');
  assert.equal(cfg.find(x => x.Clave === 'CARPETA_ENTRADA').Valor, 'ENT');
  assert.equal(cfg.find(x => x.Clave === 'MODELO_GEMINI').Valor, 'gemini-3.5-flash-lite');
  const abonos = e.ss.getSheetByName('Abonos');
  assert.match(abonos.cell(4, 3).f, /SUMIFS\(Albaranes!/);
  assert.match(abonos.cell(4, 8).f, /TOL_CUADRE/);
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
  assert.equal(ab.valor(31, 5), 'Sin abonar');
  assert.equal(ab.valor(31, 3), 11.31);
  assert.equal(ab.valor(31, 4), 13.69);
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
});

test('edición manual en Albaranes: fecha, proveedor y trabajo automáticos; NUEVO abre otro trabajo', () => {
  const e = entorno({});
  const alb = e.ss.getSheetByName('Albaranes');
  const fila = 10;
  alb.put(fila, 7, '1234 abc'); alb.put(fila, 8, 44.54);
  const rango = alb.getRange(fila, 7, 1, 2);
  e.ctx.__rango = rango;
  e.run('alEditar({ range: __rango })');
  const t = tabla(e, 'Albaranes').find(a => a['Matrícula'] === '1234ABC');
  assert.ok(t, 'fila procesada');
  assert.ok(esFecha(t['Fecha escaneo']));
  assert.ok(esFecha(t['Fecha albarán']));
  assert.equal(t['Proveedor'], 'RM');
  assert.equal(t['Nº trabajo'], '4ABC-1');
  // Segunda fila del mismo coche -> mismo trabajo abierto
  alb.put(11, 7, '1234ABC'); alb.put(11, 8, 20);
  e.ctx.__r2 = alb.getRange(11, 7, 1, 2); e.run('alEditar({ range: __r2 })');
  assert.equal(tabla(e, 'Albaranes').filter(a => a['Nº trabajo'] === '4ABC-1').length, 2);
  // NUEVO -> segundo trabajo abierto en paralelo
  alb.put(11, 6, 'NUEVO');
  e.ctx.__r3 = alb.getRange(11, 6); e.run('alEditar({ range: __r3 })');
  assert.equal(tabla(e, 'Albaranes').find(a => a['Precio con IVA'] === 20)['Nº trabajo'], '4ABC-2');
  assert.equal(tabla(e, 'Trabajos').filter(x => x['Matrícula'] === '1234ABC').length, 2);
});

test('edición en Piezas: reembolso exige nº de albarán y rellena la fecha; reconstruye Abonos', () => {
  const e = entorno({});
  const p = e.ss.getSheetByName('Piezas');
  p.put(2, 1, true); p.put(2, 5, 'Pieza a mano'); p.put(2, 10, 50);
  e.ctx.__p = p.getRange(2, 1); e.run('alEditar({ range: __p })');
  assert.equal(p.valor(2, 1), false, 'sin nº de albarán se desmarca');
  p.put(2, 3, '123456'); p.put(2, 1, true);
  e.ctx.__p2 = p.getRange(2, 1); e.run('alEditar({ range: __p2 })');
  assert.equal(p.valor(2, 1), true);
  assert.ok(esFecha(p.valor(2, 12)), 'fecha de reembolso');
  assert.equal(p.valor(2, 14), 'Manual');
  assert.equal(e.ss.getSheetByName('Abonos').valor(31, 5), 'Sin abonar');
  // desmarcar limpia la fecha y la fila de Abonos
  p.put(2, 1, false);
  e.ctx.__p3 = p.getRange(2, 1); e.run('alEditar({ range: __p3 })');
  assert.equal(p.valor(2, 12), '');
  assert.equal(e.ss.getSheetByName('Abonos').valor(31, 5), '');
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
  const estados = Array.from([31, 32].map(r => ab.valor(r, 5))).sort();
  assert.equal(ab.valor(33, 5), '');
  assert.deepEqual(estados, ['Abonada', 'Sin solicitar']);
  assert.equal(e.carpetas.FRAP.ficheros.length, 1);
  // Reprocesar la misma factura no duplica líneas
  subir(e, ['fra.pdf'], 'FRA');
  e.run('procesarFacturasRM()');
  assert.equal(tabla(e, 'Líneas RM').length, 4);
  assert.equal(tabla(e, 'Facturas RM').length, 1);
});

test('reescanear una quincena vuelve a leer la factura y reemplaza sus líneas', () => {
  const e = entorno({ 'fra.pdf': facturaRaw() });
  ['FRA', 'FRAP', 'ENT', 'PROC'].forEach(id => e.mkFolder(id, id));
  subir(e, ['fra.pdf'], 'FRA');
  e.run('procesarFacturasRM()');
  const corregida = facturaRaw(); corregida.albaranes[2].importe = -7; corregida.albaranes[2].lineas[0].importe = -7;
  e.actual.push('fra.pdf');
  // respuesta distinta en el reescaneo
  e.ctx.__cor = corregida;
  const antes = e.log.fetch.length;
  e.run(`reescanearPeriodo_(2026, 9, 1)`);
  assert.equal(e.log.fetch.length, antes + 1);
});

test('diagnóstico lista problemas con enlaces', () => {
  const e = entorno({});
  const alb = e.ss.getSheetByName('Albaranes');
  alb.put(20, 7, 'ZZZ9999'); alb.put(20, 8, 10);
  e.mkFolder('ENT', 'ENT'); e.mkFolder('FRA', 'FRA');
  e.run('diagnostico()');
  const d = e.ss.getSheetByName('Diagnóstico');
  const textos = []; for (let r = 2; r < 12; r++) textos.push(d.valor(r, 4));
  assert.ok(textos.some(t => /ZZZ9999 no está en Coches/.test(t)), textos.join('|'));
  assert.ok(textos.some(t => /Sin nº de trabajo/.test(t)));
});
