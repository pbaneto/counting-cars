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

test('construirAbonos: Abonada / Sin abonar / Sin solicitar', () => {
  const piezas = [
    { albaran: '443645', ref: 'VARTAA8', desc: 'A8 AGM VARTA', sinIva: 130.10, fechaReembolso: '2026-09-04', matricula: '5678DEF' },
    { albaran: '462446', ref: 'DAYCO6PK1090EE', desc: 'CORREA', sinIva: 11.31, fechaReembolso: '2026-09-16', matricula: '1234ABC' },
  ];
  const todas = piezas.concat([{ albaran: '439139', ref: 'WALKE80477' }]);
  const abonos = [
    { factura: 'FCR 00001', fecha: '2026-09-04', albaranOrigen: '443645', ref: 'VARTAA8', desc: 'A8 AGM VARTA 60 AH', importe: -130.10 },
    { factura: 'FCR 00001', fecha: '2026-09-07', albaranOrigen: '439139', ref: 'WALKE80477', desc: 'SPECIAL CLAMP', importe: -6.79 },
    { factura: 'FCR 00001', fecha: '2026-09-07', albaranOrigen: '396363', ref: 'EFI8115', desc: 'CABLES', importe: -22.51 },
  ];
  const f = L.construirAbonos(piezas, todas, abonos);
  const por = e => f.filter(x => x.estado === e);
  assert.equal(por('Abonada').length, 1);
  assert.equal(por('Abonada')[0].conIva, 157.42);
  assert.equal(por('Abonada')[0].fechaSolicitud, '2026-09-04');
  assert.equal(por('Sin abonar').length, 1);
  assert.equal(por('Sin abonar')[0].albaran, '462446');
  assert.equal(por('Sin solicitar').length, 2);
  assert.match(por('Sin solicitar').find(x => x.albaran === '439139').nota, /sin Reembolso marcado|no tiene Reembolso/);
  assert.equal(f[0].fechaAbono, '2026-09-04'); // ordenadas por fecha
});

test('construirAbonos empareja por importe si la referencia no coincide', () => {
  const f = L.construirAbonos([{ albaran: '1', ref: 'ABC', desc: 'd', sinIva: 10, fechaReembolso: '2026-09-01' }], [],
    [{ factura: 'F', fecha: '2026-09-02', albaranOrigen: '1', ref: 'XYZ', desc: 'd', importe: -10 }]);
  assert.equal(f.length, 1);
  assert.equal(f[0].estado, 'Abonada');
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
