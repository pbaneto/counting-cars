const test = require('node:test');
const assert = require('node:assert/strict');
const L = require('../src/Logic.js');

test('normPlate y jobPrefix', () => {
  assert.equal(L.normPlate(' 1233-kcc '), '1233KCC');
  assert.equal(L.jobPrefix('1233KCC'), '3KCC');
  assert.equal(L.jobPrefix('0123abc'), '3ABC');
  assert.equal(L.jobPrefix('M9122HT'), '2HT');
  assert.equal(L.jobPrefix('SO2402C'), '2C');
  assert.equal(L.jobPrefix('RUEDAS'), 'RUEDAS');
  assert.equal(L.jobPrefix('105627'), '105627');
});

test('nextJobNumber usa máximo + 1 y no reutiliza huecos', () => {
  assert.equal(L.nextJobNumber('3KCC', []), '3KCC-1');
  assert.equal(L.nextJobNumber('3KCC', ['3KCC-1', '3KCC-3', '5ABC-9']), '3KCC-4');
});

test('pickOpenJob: el sin pagar más reciente; coexisten varios', () => {
  const jobs = [
    { num: '3KCC-1', plate: '1233KCC', pagado: true },
    { num: '3KCC-2', plate: '1233KCC', pagado: false },
    { num: '3KCC-3', plate: '1233KCC', pagado: false },
    { num: '3KCC-1', plate: '4573KCC', pagado: false },
  ];
  assert.equal(L.pickOpenJob('1233 kcc', jobs).num, '3KCC-3');
  assert.equal(L.pickOpenJob('0000AAA', jobs), null);
  assert.equal(L.pickOpenJob('1233KCC', [jobs[0]]), null);
});

test('fechas y quincenas', () => {
  assert.equal(L.quincenaDe('2026-09-15'), 1);
  assert.equal(L.quincenaDe('2026-09-16'), 2);
  assert.deepEqual(L.rangoQuincena(2026, 2, 2), { desde: '2026-02-16', hasta: '2026-02-28' });
  assert.deepEqual(L.rangoQuincena(2028, 2, 2), { desde: '2028-02-16', hasta: '2028-02-29' });
  assert.equal(L.isoValid('2026-02-30'), false);
  assert.equal(L.daysBetween('2026-08-01', '2026-09-15'), 45);
});

test('parseNumber y albaranOrigen', () => {
  assert.equal(L.parseNumber('1.234,56'), 1234.56);
  assert.equal(L.parseNumber('11,31'), 11.31);
  assert.equal(L.parseNumber(3), 3);
  assert.ok(Number.isNaN(L.parseNumber('')));
  assert.equal(L.albaranOrigen('01000443645'), '443645');
  assert.equal(L.albaranOrigen('443645'), '443645');
});

const albaranReal = {
  numero_albaran: '462446', fecha: '2026-09-15', matricula: '1234ABC', base_imponible: 36.81, iva_importe: 7.73, total: 44.54,
  lineas: [
    { referencia: 'DAYCO6PK1090EE', descripcion: 'CORREA ESTRIADA ELASTICA', cantidad: 1, precio_unitario: 28.27, descuento_pct: 60, importe: 11.31, reembolso: true },
    { referencia: 'KRAFF47054', descripcion: 'HIDROIL 775 (SAE 75 W 80 W)', cantidad: 2, precio_unitario: 23.10, descuento_pct: 45, importe: 25.41, reembolso: false },
    { referencia: '', descripcion: 'SIGAUS (SIG. RD 679/2006)', cantidad: 2, precio_unitario: 0.05, descuento_pct: 0, importe: 0.09, reembolso: false },
  ],
};

test('validarAlbaran: albarán real 462446 cuadra sin avisos', () => {
  const v = L.validarAlbaran(albaranReal);
  assert.deepEqual(v, { errors: [], warnings: [] });
});

test('validarAlbaran: total ilegible; un total negativo sólo vale en un albarán de abono', () => {
  assert.equal(L.validarAlbaran({ lineas: [] }).errors.length, 1);
  assert.equal(L.validarAlbaran(Object.assign({}, albaranReal, { total: -10 })).errors.length, 1);
  assert.equal(L.validarAlbaran(Object.assign({}, albaranReal, { total: -10, es_abono: true })).errors.length, 0);
});

test('cuadreAlbaran: base + IVA = total y suma de líneas = base', () => {
  assert.deepEqual(L.cuadreAlbaran(albaranReal), []);
  const p = L.cuadreAlbaran(Object.assign({}, albaranReal, { total: 50, base_imponible: 40 }));
  assert.equal(p.length, 2);
  assert.match(p[0], /base 40 \+ IVA = 48,4, pero el total es 50/);
  assert.match(p[1], /las líneas suman 36,81, pero la base es 40/);
  // Sin base: las líneas se comparan con total / 1,21
  assert.deepEqual(L.cuadreAlbaran(Object.assign({}, albaranReal, { base_imponible: null })), []);
  assert.equal(L.cuadreAlbaran(Object.assign({}, albaranReal, { base_imponible: null, total: 60 })).length, 1);
  // Sin base ni líneas no hay nada que comprobar
  assert.deepEqual(L.cuadreAlbaran({ total: 12, lineas: [] }), []);
});

test('lineasParaPiezas descarta residuos SIGAUS', () => {
  assert.equal(L.lineasParaPiezas(albaranReal.lineas).length, 2);
});

test('aplicarReembolsos: marca, no duplica y añade lo que falta', () => {
  const existentes = [
    { ref: 'DAYCO6PK1090EE', desc: 'CORREA', reembolso: false },
    { ref: 'KRAFF47054', desc: 'HIDROIL', reembolso: true },
  ];
  const lineas = [
    { referencia: 'DAYCO6PK1090EE', descripcion: 'x', importe: 11.31, reembolso: true },
    { referencia: 'KRAFF47054', descripcion: 'y', importe: 25.41, reembolso: true },
    { referencia: 'NUEVA1', descripcion: 'z', importe: 5, reembolso: true },
    { referencia: 'OTRA', descripcion: 'w', importe: 1, reembolso: false },
  ];
  const r = L.aplicarReembolsos(existentes, lineas);
  assert.deepEqual(r.marcar, [0]);
  assert.equal(r.yaMarcadas, 1);
  assert.equal(r.anadir.length, 1);
  assert.equal(r.anadir[0].referencia, 'NUEVA1');
});

test('buscarFilaManual adopta sólo filas manuales con misma matrícula e importe', () => {
  const rows = [
    { row: 5, plate: '1234ABC', total: 44.54, albaran: '', pdf: '' },
    { row: 6, plate: '1234ABC', total: 44.54, albaran: '111', pdf: '' },
  ];
  assert.equal(L.buscarFilaManual(rows, '1234 abc', 44.54).row, 5);
  assert.equal(L.buscarFilaManual(rows, '1234ABC', 50), null);
  assert.equal(L.buscarFilaManual(rows, '', 44.54), null);
});

test('resolverMatricula: busca la combinación en cualquier posición, no sólo al principio', () => {
  const coches = [{ plate: '7853KCC', cliente: 'Jonthan', coche: 'Range Rover' }, { plate: '1233KCC', cliente: 'Ana', coche: 'Golf' },
    { plate: '6605HVB', cliente: 'Tomas', coche: 'Mini Cooper' }];
  assert.equal(L.resolverMatricula('', coches).tipo, 'ninguna');
  assert.equal(L.resolverMatricula('7853KCC', coches).tipo, 'exacta', 'ya es una matrícula real: no se toca');
  assert.equal(L.resolverMatricula('HV', coches).tipo, 'ninguna', 'demasiado corto para buscar (< 3)');
  const unica = L.resolverMatricula('HVB', coches);
  assert.equal(unica.tipo, 'unica');
  assert.equal(unica.candidatos[0].plate, '6605HVB');
  const soloLetras = L.resolverMatricula('KCC', coches);  // las letras van al final: el desplegable nativo no las encuentra
  assert.equal(soloLetras.tipo, 'varias');
  assert.deepEqual(soloLetras.candidatos.map(c => c.plate).sort(), ['1233KCC', '7853KCC']);
  assert.equal(L.resolverMatricula('999', coches).tipo, 'ninguna');
});

// ---- Factura real FCR 00001 (importes sin IVA por albarán) ----
const compras = [247.70, 30.31, 75.26, 4.61, 47.33, 299.74, 6.79, 117.53, 25.69, 19.96, 14.95, 109.52, 69.06, 55.13, 7.78, 14.66, 23.75, 184.00, 26.14, 2.77,
  183.80, 67.16, 169.08, 46.36, 0.02, 49.61, 143.57, 24.12, 5.35, 47.48, 104.12, 132.72, 17.58, 23.75, 18.77, 8.22, 74.12, 114.92, 27.68, 94.32, 29.98,
  4.65, 240.69, 21.61, 116.95, 60.26, 139.09, 27.68, 133.88, 36.81, 26.69, 25.78, 13.43];
const abonosImp = [-130.10, -248.02, -5.35, -29.98, -139.09];
function facturaReal(extra) {
  const albaranes = compras.map((imp, i) => ({ numero_albaran: String(437000 + i), fecha: i < 20 ? '2026-09-01' : '2026-09-15', matricula: 'X', importe: imp, es_abono: false, lineas: [{ importe: imp }] }))
    .concat(abonosImp.map((imp, i) => ({ numero_albaran: String(446000 + i), fecha: '2026-09-07', matricula: '', importe: imp, es_abono: true, lineas: [{ importe: imp }] })));
  return Object.assign({ numero_factura: 'FCR 00001', fecha_factura: '2026-09-15', base_imponible: 3060.39, iva_importe: 642.68, total: 3703.07, albaranes }, extra);
}

test('validarFactura: la factura real cuadra (58 albaranes, base 3060,39)', () => {
  assert.equal(facturaReal().albaranes.length, 58);
  assert.deepEqual(L.validarFactura(facturaReal()), { errors: [], warnings: [] });
});

test('validarFactura detecta un fallo de escaneo en el total', () => {
  const v = L.validarFactura(facturaReal({ total: 3703.70 }));
  assert.equal(v.errors.length, 0);
  assert.ok(v.warnings.some(w => /total/.test(w)));
  const v2 = L.validarFactura(facturaReal({ base_imponible: 3000, iva_importe: 630, total: 3630 }));
  assert.ok(v2.warnings.some(w => /suma de albaranes/.test(w)));
});

test('periodoFactura: quincena de la última compra', () => {
  assert.deepEqual(L.periodoFactura(facturaReal()), { year: 2026, month: 9, quincena: 1 });
});

const piezasEj = () => [
  { id: 1, albaran: '443645', ref: 'VARTAA8', desc: 'A8 AGM VARTA', sinIva: 130.10, solicitada: true, fechaSolicitud: '2026-09-04', abonadaAMano: '', matricula: '5678DEF' },
  { id: 2, albaran: '462446', ref: 'DAYCO6PK1090EE', desc: 'CORREA', sinIva: 11.31, solicitada: true, fechaSolicitud: '2026-09-16', abonadaAMano: '', matricula: '1234ABC' },
];
const abonosEj = () => [
  { idPieza: '', factura: 'FCR 00001', fecha: '2026-09-04', albaranAbono: '446334', albaranOrigen: '443645', ref: 'VARTAA8', desc: 'A8 AGM VARTA 60 AH', importe: -130.10 },
  { idPieza: '', factura: 'FCR 00001', fecha: '2026-09-07', albaranAbono: '446877', albaranOrigen: '439139', ref: 'WALKE80477', desc: 'SPECIAL CLAMP', importe: -6.79 },
];
const emparejar = (piezas, abonos) => { L.emparejarAbonos(piezas, abonos).forEach(({ i, id }) => { abonos[i].idPieza = id; }); return abonos; };

test('clavesAbonos: la misma línea del escaneo y de la factura, aunque el albarán original se lea distinto', () => {
  const k = a => L.clavesAbonos([a])[0].clave;
  assert.equal(k({ albaranAbono: '504271', albaranOrigen: '4740419', ref: 'AUTEL 06' }), k({ albaranAbono: '504271', albaranOrigen: '470419', ref: 'AUTEL06' }));
  assert.deepEqual(L.clavesAbonos([{ albaranAbono: '1', ref: 'X' }, { albaranAbono: '1', ref: 'X' }, { albaranAbono: '1', desc: 'Filtro' }]).map(a => a.clave),
    ['A|1|X|1', 'A|1|X|2', 'A|1|DFILTRO|1']);
});

test('emparejarAbonos: mismo albarán y referencia; si no, una pieza pedida con el mismo importe; una pieza, un abono', () => {
  const piezas = piezasEj().concat([{ id: 3, albaran: '1', ref: 'ABC', sinIva: 10, solicitada: true }, { id: 4, albaran: '1', ref: 'OTRA', sinIva: 10, solicitada: false }]);
  const abonos = abonosEj().concat([
    { idPieza: '', fecha: '2026-09-02', albaranOrigen: '1', ref: 'XYZ', importe: -10 },
    { idPieza: '', fecha: '2026-09-03', albaranOrigen: '1', ref: 'XYZ2', importe: -10 },
  ]);
  assert.deepEqual(L.emparejarAbonos(piezas, abonos), [{ i: 2, id: 3 }, { i: 0, id: 1 }], 'por importe sólo con piezas pedidas, y la pieza 3 ya está cogida');
  // Una pieza sin pedir, con la misma referencia: se empareja (así, si luego se marca, el abono ya es suyo)
  assert.deepEqual(L.emparejarAbonos([{ id: 9, albaran: '439139', ref: 'WALKE80477', solicitada: false }], abonosEj()), [{ i: 1, id: 9 }]);
});

test('emparejarAbonos: un ID ya puesto no se cambia nunca, ni "NO" (sin pieza)', () => {
  const abonos = abonosEj();
  abonos[0].idPieza = 'no';
  assert.deepEqual(L.emparejarAbonos(piezasEj(), abonos), []);
  abonos[0].idPieza = 77;  // la pieza ya no existe: no se busca otra
  assert.deepEqual(L.emparejarAbonos(piezasEj(), abonos), []);
  abonos[0].idPieza = 2;  // puesto a mano a otra pieza: se respeta, y esa pieza ya no está libre
  assert.deepEqual(L.emparejarAbonos([{ id: 2, albaran: '443645', ref: 'VARTAA8', solicitada: true }], [abonos[0], Object.assign({}, abonos[0], { idPieza: '' })]), []);
});

test('informeAbonos: Abonada, Sin abonar, Sin solicitar y Abonada (a mano); lo más reciente arriba', () => {
  const piezas = piezasEj().concat([{ id: 5, albaran: '500', ref: 'M', desc: 'Manual', sinIva: 20, solicitada: true, fechaSolicitud: '2026-09-20', abonadaAMano: '2026-09-25' }]);
  const filas = L.informeAbonos(piezas, emparejar(piezas, abonosEj()), { 439139: '9999ZZZ' });
  assert.deepEqual(filas.map(f => [f.estado, f.referencia]), [
    ['Abonada (a mano)', 'M'], ['Sin abonar', 'DAYCO6PK1090EE'], ['Sin solicitar', 'WALKE80477'], ['Abonada', 'VARTAA8']]);
  const abonada = filas[3];
  assert.deepEqual([abonada.idPieza, abonada.fechaAbono, abonada.fechaSolicitud, abonada.factura, abonada.descripcion, abonada.matricula, abonada.nota],
    [1, '2026-09-04', '2026-09-04', 'FCR 00001', 'A8 AGM VARTA', '5678DEF', '']);
  assert.equal(filas[2].matricula, '9999ZZZ', 'sin pieza: la matrícula de su albarán');
  assert.equal(filas[0].fechaAbono, '2026-09-25');
});

test('informeAbonos: editar la pieza (precio, referencia, albarán) cambia su fila, nunca añade otra', () => {
  const piezas = piezasEj(), abonos = emparejar(piezas, abonosEj());
  const antes = L.informeAbonos(piezas, abonos);
  Object.assign(piezas[1], { ref: 'OTRAREF', albaran: '999999', sinIva: 12 });
  Object.assign(piezas[0], { desc: 'BATERIA', sinIva: 131 });
  const despues = L.informeAbonos(piezas, abonos);
  assert.equal(despues.length, antes.length);
  const sinAbonar = despues.find(f => f.idPieza === 2);
  assert.deepEqual([sinAbonar.referencia, sinAbonar.albaran, sinAbonar.sinIva], ['OTRAREF', '999999', 12]);
  const abonada = despues.find(f => f.idPieza === 1);
  assert.deepEqual([abonada.estado, abonada.descripcion, abonada.sinIva], ['Abonada', 'BATERIA', 130.1], 'el importe es lo abonado');
  assert.match(abonada.nota, /Importe abonado distinto: se pidió 131/);
});

test('informeAbonos: desmarcar una pieza quita su fila; si ya tenía abono, éste pasa a "Sin solicitar"', () => {
  const piezas = piezasEj(), abonos = emparejar(piezas, abonosEj());
  piezas.forEach(p => { p.solicitada = false; });
  const filas = L.informeAbonos(piezas, abonos);
  assert.deepEqual(filas.map(f => f.estado), ['Sin solicitar', 'Sin solicitar']);
  assert.match(filas.find(f => f.idPieza === 1).nota, /no tiene Reembolso marcado/);
  // La pieza se borra de Piezas: su abono sigue, sin pieza
  const sinPieza = L.informeAbonos([], abonos);
  assert.match(sinPieza.find(f => f.referencia === 'VARTAA8').nota, /La pieza 1 ya no está en Piezas/);
  assert.equal(L.informeAbonos([], [Object.assign(abonos[0], { idPieza: 'NO' })])[0].nota, '');
});

test('emparejarAbonos: RM abona sobre otro albarán, o sin albarán: vale la única pieza pedida con esa referencia, pedida antes', () => {
  const bomba = { id: 25, albaran: '336950', ref: 'DRI715520576', desc: 'BOMBA DE DIRECCIÓN', sinIva: 147.75, solicitada: true, fechaSolicitud: '2026-09-30' };
  const sensor = { id: 34, albaran: '489996', ref: 'ERA560020A', desc: 'SENSOR', sinIva: 18.75, solicitada: true, fechaSolicitud: '2026-10-01' };
  const abonos = [
    { idPieza: '', fecha: '2026-10-15', albaranOrigen: '332174', ref: 'DRI715520576', importe: -147.75 },
    { idPieza: '', fecha: '2026-10-15', albaranOrigen: '', ref: 'ERA560020A', importe: -18.75 },
    { idPieza: '', fecha: '2026-07-31', albaranOrigen: '390518', ref: 'ERA560020A', importe: -18.75 },  // abono de julio: anterior a la petición
  ];
  assert.deepEqual(L.emparejarAbonos([bomba, sensor], abonos), [{ i: 0, id: 25 }, { i: 1, id: 34 }]);
  assert.deepEqual(L.emparejarAbonos([sensor, Object.assign({}, sensor, { id: 35, albaran: '1' })], [abonos[1]]), [], 'dos pedidas con esa referencia: no se adivina');
  abonos[0].idPieza = 25;
  assert.match(L.informeAbonos([bomba], [abonos[0]])[0].nota, /RM lo abona sobre el albarán 332174/);
});

test('informeAbonos: desde una fecha, salvo lo que sigue sin abonar', () => {
  const piezas = piezasEj();
  piezas[1].fechaSolicitud = '2026-09-20';
  const filas = L.informeAbonos(piezas, emparejar(piezas, abonosEj()), {}, '2026-09-05');
  assert.deepEqual(filas.map(f => [f.estado, f.referencia]), [['Sin abonar', 'DAYCO6PK1090EE'], ['Sin solicitar', 'WALKE80477']]);
  piezas[1].fechaSolicitud = '2026-08-01';
  assert.equal(L.informeAbonos(piezas, [], {}, '2026-09-05')[0].estado, 'Sin abonar', 'lo pendiente sale siempre');
});

test('esMarcaReembolso: sólo una R sola', () => {
  assert.deepEqual(['R', '—R', ' - R ', 'r', 'CHEMA', 'MV821', 'R10', 'RR', ''].map(L.esMarcaReembolso), [true, true, true, true, false, false, false, false, false]);
});

test('versiones: comparar y saber si cambia el diseño', () => {
  assert.ok(L.compararVersiones('1.0.2', '1.0.10') < 0, 'por número, no por texto');
  assert.ok(L.compararVersiones('2.0.0', '1.9.9') > 0);
  assert.equal(L.compararVersiones('1.2.3', '1.2.3'), 0);
  assert.equal(L.cambiaDiseno('1.0.1', '1.0.2'), false, 'un parche no toca la hoja');
  assert.equal(L.cambiaDiseno('1.0.2', '1.1.0'), true);
  assert.equal(L.cambiaDiseno('1.4.0', '2.0.0'), true);
});

test('localizarFormula: ";" y decimal con coma fuera de las comillas (Sheets en español)', () => {
  assert.equal(L.localizarFormula('=IF(A1="a,b",0.05,2)', true), '=IF(A1="a,b";0,05;2)');
  assert.equal(L.localizarFormula('=HYPERLINK("https://x.y/d/1","Ver PDF")', true), '=HYPERLINK("https://x.y/d/1";"Ver PDF")');
  assert.equal(L.localizarFormula('=SUMIFS(A:A,B:B,">="&C1)', false), '=SUMIFS(A:A,B:B,">="&C1)');
  assert.equal(L.localizarFormula('texto, 1.5', true), 'texto, 1.5');
  assert.equal(L.usaPuntoYComa('es_ES'), true);
  assert.equal(L.usaPuntoYComa('es_MX'), false);
  assert.equal(L.usaPuntoYComa('en_US'), false);
  assert.equal(L.usaPuntoYComa('de_CH'), false);
});
