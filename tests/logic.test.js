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

test('validarAlbaran detecta total ilegible, IVA y suma de líneas', () => {
  assert.equal(L.validarAlbaran({ lineas: [] }).errors.length, 1);
  const malo = Object.assign({}, albaranReal, { total: 50, base_imponible: 40 });
  const w = L.validarAlbaran(malo).warnings.join('|');
  assert.match(w, /IVA no cuadra/);
  assert.match(w, /suma de líneas/);
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

const piezasEj = () => L.clavesPiezas([
  { albaran: '443645', ref: 'VARTAA8', desc: 'A8 AGM VARTA', sinIva: 130.10, fechaReembolso: '2026-09-04', matricula: '5678DEF' },
  { albaran: '462446', ref: 'DAYCO6PK1090EE', desc: 'CORREA', sinIva: 11.31, fechaReembolso: '2026-09-16', matricula: '1234ABC' },
]);
const abonosEj = () => L.clavesAbonos([
  { factura: 'FCR 00001', fecha: '2026-09-04', albaranOrigen: '443645', ref: 'VARTAA8', desc: 'A8 AGM VARTA 60 AH', importe: -130.10 },
  { factura: 'FCR 00001', fecha: '2026-09-07', albaranOrigen: '439139', ref: 'WALKE80477', desc: 'SPECIAL CLAMP', importe: -6.79 },
]);

test('claves: estables y distintas para piezas repetidas del mismo albarán', () => {
  const p = L.clavesPiezas([{ albaran: '1', ref: 'ab-1' }, { albaran: '1', ref: 'AB1' }, { albaran: '1', ref: '', desc: 'Filtro' }]);
  assert.deepEqual(p.map(x => x.clave), ['P|1|AB1|1', 'P|1|AB1|2', 'P|1|DFILTRO|1']);
  assert.equal(L.clavesAbonos([{ factura: 'F 1', albaranOrigen: '9', ref: 'X' }])[0].clave, 'A|F 1|9|X|1');
  assert.deepEqual(L.partirClaves(' P|1|A|1 ; A|F|1|A|1 '), ['P|1|A|1', 'A|F|1|A|1']);
});

test('sincronizarAbonos: tabla vacía → piezas "Sin abonar" y abonos (Abonada / Sin solicitar), lo más reciente arriba', () => {
  const piezas = piezasEj(), todas = piezas.concat(L.clavesPiezas([{ albaran: '439139', ref: 'WALKE80477' }]));
  const r = L.sincronizarAbonos([], piezas, todas, abonosEj(), new Set());
  assert.deepEqual(r.nuevas.map(f => f.estado), ['Sin abonar', 'Sin solicitar', 'Abonada']);
  const abonada = r.nuevas[2];
  assert.equal(abonada.fechaSolicitud, '2026-09-04');
  assert.equal(abonada.factura, 'FCR 00001');
  assert.equal(abonada.clave, 'P|443645|VARTAA8|1;A|FCR 00001|443645|VARTAA8|1');
  assert.match(r.nuevas[1].nota, /no tiene Reembolso/);
  assert.deepEqual(r.registrar.sort(), ['A|FCR 00001|439139|WALKE80477|1', 'A|FCR 00001|443645|VARTAA8|1']);
});

test('sincronizarAbonos: no toca lo que ya hay; sólo rellena huecos de la fila de la pieza al llegar su abono', () => {
  const piezas = piezasEj();
  const existentes = [
    { clave: piezas[0].clave, estado: 'Sin abonar', albaran: '443645', ref: 'VARTAA8', sinIva: 130.10, fechaAbono: '', factura: '', nota: 'mi nota' },
    { clave: piezas[1].clave, estado: 'Abonada', albaran: '462446', ref: 'DAYCO6PK1090EE', sinIva: 11.31, fechaAbono: '', factura: '', nota: '' },  // cambiado a mano
    { clave: 'A|FCR 00001|439139|WALKE80477|1', estado: 'Abonada', albaran: '439139', ref: 'WALKE80477' },  // "Sin solicitar" cambiado a mano
  ];
  const r = L.sincronizarAbonos(existentes, piezas, piezas, abonosEj(), new Set(['A|FCR 00001|439139|WALKE80477|1']));
  assert.equal(r.nuevas.length, 0);
  assert.deepEqual(r.cambios, [{ i: 0, v: { clave: 'P|443645|VARTAA8|1;A|FCR 00001|443645|VARTAA8|1', estado: 'Abonada', fechaAbono: '2026-09-04', factura: 'FCR 00001' } }]);
  // Una segunda sincronización no cambia nada más
  existentes[0].clave = r.cambios[0].v.clave;
  const r2 = L.sincronizarAbonos(existentes, piezas, piezas, abonosEj(), new Set(r.registrar.concat('A|FCR 00001|439139|WALKE80477|1')));
  assert.deepEqual([r2.nuevas.length, r2.cambios.length, r2.borrar.length], [0, 0, 0]);
});

test('sincronizarAbonos: una fila de abono borrada no vuelve; una pieza desmarcada borra su fila', () => {
  const piezas = piezasEj();
  const r = L.sincronizarAbonos([], [], piezas, abonosEj(), new Set(['A|FCR 00001|439139|WALKE80477|1', 'A|FCR 00001|443645|VARTAA8|1']));
  assert.equal(r.nuevas.length, 0);
  const existentes = [{ clave: 'X' }, { clave: piezas[1].clave + ';A|F|1|X|1', estado: 'Abonada' }, { clave: '' }];
  const r2 = L.sincronizarAbonos(existentes, [piezas[0]], piezas, [], new Set(), [piezas[1].clave]);
  assert.deepEqual(r2.borrar, [1]);
  assert.equal(r2.nuevas.length, 1, 'la otra pieza marcada, que faltaba, sí se añade');
});

test('mapaMatriculas y sincronizarAbonos: rellena sólo las matrículas vacías', () => {
  const m = L.mapaMatriculas([['100', ''], ['100', '1234 abc'], ['100', '9999ZZZ'], ['', 'X']]);
  assert.deepEqual(m, { 100: '1234ABC' });
  const existentes = [{ clave: 'x', albaran: '100', matricula: '' }, { clave: 'y', albaran: '100', matricula: 'A MANO' }, { clave: 'z', albaran: '200', matricula: '' }];
  const r = L.sincronizarAbonos(existentes, [], [], [], new Set(), [], { matriculas: m });
  assert.deepEqual(r.cambios, [{ i: 0, v: { matricula: '1234ABC' } }]);
});

test('sincronizarAbonos empareja por importe si la referencia no coincide', () => {
  const p = L.clavesPiezas([{ albaran: '1', ref: 'ABC', desc: 'd', sinIva: 10, fechaReembolso: '2026-09-01' }]);
  const r = L.sincronizarAbonos([], p, p, L.clavesAbonos([{ factura: 'F', fecha: '2026-09-02', albaranOrigen: '1', ref: 'XYZ', desc: 'd', importe: -10 }]), new Set());
  assert.equal(r.nuevas.length, 1);
  assert.equal(r.nuevas[0].estado, 'Abonada');
});

test('clavesDeFilasAntiguas: reconoce las filas de la tabla generada por la versión anterior', () => {
  const piezas = piezasEj(), abonos = abonosEj();
  const filas = [
    { factura: 'FCR 00001', albaran: '443645', ref: 'VARTAA8', fechaSolicitud: '2026-09-04', estado: 'Abonada' },
    { factura: 'FCR 00001', albaran: '439139', ref: 'WALKE80477', estado: 'Sin solicitar' },
    { albaran: '462446', ref: 'DAYCO6PK1090EE', fechaSolicitud: '2026-09-16', estado: 'Sin abonar' },
    { descripcion: 'fila escrita a mano' },
  ];
  assert.deepEqual(L.clavesDeFilasAntiguas(filas, piezas, abonos), [
    'A|FCR 00001|443645|VARTAA8|1;P|443645|VARTAA8|1', 'A|FCR 00001|439139|WALKE80477|1', 'P|462446|DAYCO6PK1090EE|1', '']);
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
