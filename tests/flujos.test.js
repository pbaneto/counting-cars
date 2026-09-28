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
  assert.match(alb.cell(2, 13).f, /DIAS_AVISO/);           // Avisos en las filas con datos (piloto: filas 2-4)
  assert.equal(alb.cell(5, 13), undefined, 'sin fórmulas en filas vacías: leerlas y recalcularlas era lo lento');
  assert.equal(alb.getLastRow(), 4);
  // Casillas sólo en filas con datos: una casilla vacía vale FALSE y haría leer miles de filas
  const trab = e.ss.getSheetByName('Trabajos');
  assert.equal(trab.getLastRow(), 3, 'Trabajos: 2 trabajos del piloto, sin casillas "Pagado" por debajo');
  assert.equal(e.ss.getSheetByName('Piezas').getLastRow(), 1, 'Piezas vacía: sin casillas "Reembolso" por debajo');
  assert.equal(trab.valor(3, 11), false, 'la fila nueva lleva su casilla (FALSE = sin pagar), ahora en la columna K');
  assert.equal(trab.valor(1, 3), 'Mes', 'columna nueva "Mes" junto a "Fecha apertura"');
  assert.match(trab.cell(2, 3).f, /CHOOSE\(MONTH\(/);
  // Ninguna columna calculada (gris) se queda con una validación residual; las de entrada no se tocan
  assert.deepEqual(new Set(trab.validacionesLimpiadas), new Set([3, 5, 6, 7, 8, 10, 12]), 'Mes, Coche, Cliente, Recambios, Recambios facturables, Beneficio, Avisos');
  const cfg = tabla(e, 'Config');
  assert.equal(cfg.find(x => x.Clave === 'CARPETA_ENTRADA').Valor, 'ENT');
  assert.equal(cfg.find(x => x.Clave === 'MODELO_GEMINI').Valor, 'gemini-3.5-flash-lite');
  const abonos = e.ss.getSheetByName('Abonos');
  assert.match(abonos.cell(4, 3).f, /SUMIFS\(Albaranes!/);
  // Fechas de cada quincena dentro de la fórmula (sin columnas Desde/Hasta): ene 1ª = 1-15, ene 2ª = 16-fin de mes
  assert.match(abonos.cell(4, 3).f, /">="&DATE\(\$B\$1,1,1\).*"<="&DATE\(\$B\$1,1,15\)/);
  assert.match(abonos.cell(5, 4).f, /">="&DATE\(\$B\$1,1,16\).*"<="&EOMONTH\(DATE\(\$B\$1,1,1\),0\)/);
  assert.equal(abonos.cell(3, 7), undefined);               // ya no hay columna "Desde"
  assert.match(abonos.cell(4, 6).f, /TOL_CUADRE/);          // Estado (cuadre factura vs albaranes)
  assert.equal(abonos.cell(4, 12).f, '=DIAS_AVISO_REEMB');  // panel: umbral copiado de Config a esta pestaña
  assert.match(abonos.cell(5, 12).f, /COUNTIFS.*\$L\$4/);    // panel: fuera de plazo, contra el umbral de la misma pestaña
  e.run('repararFormulas()');
  const errores = e.log.console.filter(m => /^\[ERROR\]/.test(m));
  assert.deepEqual(errores, [], 'setup y repararFormulas deben terminar sin errores');
});

test('repararFormulas quita las fórmulas de las filas vacías (hoja antigua) sin tocar datos ni valores a mano', () => {
  const e = entorno({});
  const alb = e.ss.getSheetByName('Albaranes'), p = e.ss.getSheetByName('Piezas');
  // Como la hoja real antes del cambio: fórmulas rellenadas por adelantado muy por debajo de los datos
  for (let r = 5; r <= 1501; r++) { alb.put(r, 3, `=IF(B${r}="","",1)`); alb.put(r, 13, `=IF(G${r}="","",1)`); }
  p.put(4001, 15, '=IF(C4001="","",1)');
  for (let r = 2; r <= 4001; r++) p.put(r, 1, false);   // casillas "Reembolso" rellenadas hasta la 4001 (valen FALSE)
  const trab = e.ss.getSheetByName('Trabajos');
  for (let r = 4; r <= 801; r++) trab.put(r, 11, false); // y "Pagado" (columna K tras la migración de "Mes") hasta la 801
  alb.put(900, 10, 'escrito a mano');                 // valor (no fórmula) en la columna calculada "Coche"
  e.run('repararFormulas()');
  assert.equal(alb.cell(700, 3), undefined, 'Quincena vacía por debajo de los datos');
  assert.equal(alb.cell(1501, 13), undefined, 'Avisos vacía por debajo de los datos');
  assert.equal(p.cell(4001, 15), undefined);
  assert.match(alb.cell(4, 13).f, /DIAS_AVISO/, 'las filas con datos conservan sus fórmulas');
  assert.equal(alb.valor(900, 10), 'escrito a mano', 'no borra un valor escrito a mano');
  assert.ok(e.log.console.some(l => /\[AVISO\] repararFormulas Albaranes!J900/.test(l)));
  assert.equal(p.getLastRow(), 1);
  assert.equal(trab.getLastRow(), 3, 'Trabajos queda con sus 2 filas de datos');
  assert.equal(trab.valor(2, 11), true, 'el Pagado de las filas con datos no se toca');
});

test('Trabajos: la migración a la columna "Mes" no desalinea una hoja de la versión anterior', () => {
  const e = entorno({});  // entorno() ya corre setup() una vez: Trabajos ya tiene "Mes" en su sitio (D=Matrícula, K=Pagado)
  const trab = e.ss.getSheetByName('Trabajos');
  const matriculaAntes = trab.valor(2, 4), pagadoAntes = trab.valor(2, 11);
  assert.equal(matriculaAntes, '4321GHJ');
  // Deshace la migración a mano para simular una hoja real todavía sin la columna "Mes" (Matrícula en la C, como antes)
  const vieja = new Map();
  trab.grid.forEach((v, k) => {
    const [r, c] = k.split(',').map(Number);
    if (c === 3) return;  // la columna "Mes" no existía
    vieja.set(r + ',' + (c > 3 ? c - 1 : c), v);
  });
  trab.grid = vieja;
  assert.equal(trab.valor(2, 3), matriculaAntes, 'hoja "vieja": Matrícula está en la C, sin columna Mes');

  e.run('repararFormulas()');
  assert.equal(trab.valor(1, 3), 'Mes', 'ahora la C es "Mes"');
  assert.match(trab.cell(2, 3).f, /CHOOSE\(MONTH\(/);
  assert.equal(trab.valor(2, 4), matriculaAntes, 'Matrícula se ha desplazado a la D con su valor intacto');
  assert.equal(trab.valor(2, 11), pagadoAntes, 'Pagado se ha desplazado con su valor intacto');

  // Repetir la reparación es idempotente: no vuelve a insertar otra columna de más
  const anchoAntes = trab.getLastColumn();
  e.run('repararFormulas()');
  assert.equal(trab.getLastColumn(), anchoAntes, 'segunda reparación: no inserta otra columna');
  assert.equal(trab.valor(2, 4), matriculaAntes, 'Matrícula sigue en la D');
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
  assert.equal(ab.valor(31, 5), 'Sin abonar');
  assert.equal(ab.valor(31, 3), 11.31);
  assert.equal(ab.cell(31, 4).f, '=ROUND($C31*(1+IVA),2)');  // Precio con IVA: fórmula con el IVA de Config
  assert.match(ab.cell(31, 12).f, /TODAY\(\)/);  // Días pendiente: fórmula viva, no un valor fijo
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

test('Matrícula en Albaranes: busca por cualquier combinación de letras o dígitos, no sólo al principio', () => {
  const e = entorno({});  // Coches del piloto: 1234ABC, 4321GHJ, 5678DEF
  const alb = e.ss.getSheetByName('Albaranes');

  // Coincidencia única por las LETRAS (van al final: el desplegable nativo de Sheets sólo busca desde el principio)
  alb.put(10, 7, 'ghj');
  e.ctx.__unica = alb.getRange(10, 7); e.run('alEditar({ range: __unica })');
  assert.equal(alb.valor(10, 7), '4321GHJ', 'se autocompleta con la única matrícula que contiene "GHJ"');
  assert.ok(e.log.toasts.some(t => t.includes('"GHJ" → 4321GHJ')), 'avisa de qué matrícula ha puesto');

  // Texto demasiado corto (< 3): no se busca, se deja tal cual y sigue el aviso habitual de "no está en Coches"
  // (ese aviso sólo se comprueba cuando la fila ya tiene precio, igual que sin este cambio)
  e.log.toasts.length = 0;
  alb.put(11, 7, '4'); alb.put(11, 8, 44.54);
  e.ctx.__corto = alb.getRange(11, 7, 1, 2); e.run('alEditar({ range: __corto })');
  assert.equal(alb.valor(11, 7), '4', 'texto demasiado corto: se deja tal cual, sin adivinar');
  assert.equal(e.log.toasts.length, 1);
  assert.match(e.log.toasts[0], /no está en la pestaña Coches/);

  // Coincide con varias: no adivina, deja lo escrito y lista las candidatas (se añade un coche a propósito)
  e.log.toasts.length = 0;
  const coches = e.ss.getSheetByName('Coches');
  const filaNueva = coches.getLastRow() + 1;
  coches.put(filaNueva, 1, '4321GHK'); coches.put(filaNueva, 2, 'Otro'); coches.put(filaNueva, 3, 'Otro coche');
  alb.put(12, 7, '432');
  e.ctx.__varias = alb.getRange(12, 7); e.run('alEditar({ range: __varias })');
  assert.equal(alb.valor(12, 7), '432', 'coincide con varias: no adivina, deja lo escrito');
  assert.equal(e.log.toasts.length, 1);
  assert.match(e.log.toasts[0], /"432" coincide con 2 matrículas: 4321GHJ, 4321GHK/);

  // Matrícula ya exacta: no hay búsqueda ni aviso de "completada"
  e.log.toasts.length = 0;
  alb.put(13, 7, '5678DEF');
  e.ctx.__exacta = alb.getRange(13, 7); e.run('alEditar({ range: __exacta })');
  assert.equal(alb.valor(13, 7), '5678DEF');
  assert.ok(!e.log.toasts.some(t => t.includes('completada')), 'ya era exacta: no hay aviso de autocompletado');
});

test('edición en Piezas: reembolso exige nº de albarán y rellena la fecha; reconstruye Abonos', () => {
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
    'leer Albaranes \\(3 de 3 filas\\).*leer tamaño de Abonos.*revisar 1 fila\\(s\\) editada\\(s\\).*reconstruir Abonos.*guardar cambios y recalcular \\d+ms'));
  assert.match(linea('⏱ escribirAbonos_'), /preparar piezas \(1 marcadas.*cruce .*borrar .*escribir \(1 filas\).*Sin abonar 1/);
  assert.match(linea('■ alEditar'), /^■ alEditar: \d+ms \(escribir en Registro \d+ms\)$/);
  assert.equal(e.log.console.filter(l => /^\[INFO\] reconstruirAbonos/.test(l)).length, 0, 'el resumen ya no se escribe en Registro');
  assert.equal(p.valor(2, 1), true);
  assert.ok(esFecha(p.valor(2, 12)), 'fecha de reembolso');
  assert.equal(p.valor(2, 14), 'Manual');
  assert.match(p.cell(2, 15).f, /Falta el nº de albarán/, 'la fila escrita a mano recibe sus fórmulas');
  assert.equal(e.ss.getSheetByName('Abonos').valor(31, 5), 'Sin abonar');
  // Pieza de un albarán de proveedor "Otros": no se reclama a RM
  const alb = e.ss.getSheetByName('Albaranes');
  alb.put(5, 4, 'Otros'); alb.put(5, 5, '999'); alb.put(5, 7, '1234ABC'); alb.put(5, 8, 20);
  p.put(3, 3, '999'); p.put(3, 5, 'De otro proveedor'); p.put(3, 10, 10); p.put(3, 1, true);
  e.ctx.__p4 = p.getRange(3, 1); e.run('alEditar({ range: __p4 })');
  assert.equal(e.ss.getSheetByName('Abonos').valor(31, 6), '123456');
  assert.equal(e.ss.getSheetByName('Abonos').valor(32, 6), '', 'la pieza de "Otros" no entra en Abonos');
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

test('con hoja en español (es_ES) TODAS las fórmulas y reglas se escriben con ";"', () => {
  let e = crearEntorno({ privado: true, privadoRuta: PRIV, locale: 'es_ES', gemini: () => ({}) });
  e.run('setup()');
  const separadorIngles = f => /(?:^|[^\d]),|,(?:[^\d]|$)/.test(f.replace(/"[^"]*"/g, '""'));  // una coma que no sea decimal (entre dígitos)
  // Las fórmulas por fila ya no se escriben en filas vacías: se comprueban todas tal como se escribirían en la fila 2.
  const porFila = e.run("Object.keys(FORMULAS).flatMap(n => Object.keys(FORMULAS[n]).map(h => [n + ' ▸ ' + h, loc_(FORMULAS[n][h](2))]))");
  assert.ok(porFila.length >= 15);
  assert.ok(porFila.some(([k, f]) => k === 'Albaranes ▸ Quincena' && f === '=IF(B2="";"";IF(DAY(B2)<=15;1;2))'));
  for (const [k, f] of porFila) assert.ok(!separadorIngles(f), `${k}: ${f}`);
  for (const [hoja, col, fila] of [['Abonos', 3, 4], ['Abonos', 4, 4], ['Abonos', 6, 4], ['Abonos', 12, 3], ['Resumen', 2, 13]]) {
    const f = e.ss.getSheetByName(hoja).cell(fila, col).f;
    assert.ok(f && !separadorIngles(f), `${hoja} ${fila},${col}: ${f}`);
  }
  e.run('cargarDatosIniciales()');
  assert.ok(!separadorIngles(e.ss.getSheetByName('Albaranes').cell(2, 13).f), 'y las filas añadidas también');
  e.mkFolder('ENT', 'E'); e.mkFolder('PROC', 'P');
  assert.equal(e.run("enlacePdf_({ getId: () => 'abc' })"), '=HYPERLINK("https://drive.google.com/file/d/abc/view";"Ver PDF")');
});
