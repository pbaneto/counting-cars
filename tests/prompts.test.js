const test = require('node:test');
const assert = require('node:assert/strict');
const P = require('../src/Prompts.js');
const L = require('../src/Logic.js');

test('normalizarAlbaran limpia tipos y matrícula', () => {
  const d = P.normalizarAlbaran({ es_albaran: true, proveedor: 'RM', numero_albaran: ' 462446 ', fecha: '2026-09-15', matricula: '1234 abc',
    total: '44,54', base_imponible: 36.81, lineas: [{ referencia: 'DAYCO 6PK1090EE', descripcion: ' CORREA  ESTRIADA ', importe: '11,31', reembolso: true }] });
  assert.equal(d.numero_albaran, '462446');
  assert.equal(d.matricula, '1234ABC');
  assert.equal(d.total, 44.54);
  assert.equal(d.lineas[0].referencia, 'DAYCO6PK1090EE');
  assert.equal(d.lineas[0].descripcion, 'CORREA ESTRIADA');
  assert.equal(d.lineas[0].reembolso, true);
});

test('normalizarAlbaran: fecha inválida -> "" y proveedor por defecto RM', () => {
  const d = P.normalizarAlbaran({ fecha: '15/09/26', lineas: [] });
  assert.equal(d.fecha, '');
  assert.equal(d.proveedor, 'RM');
});

test('normalizarFactura marca abonos y limpia albaran_origen', () => {
  const d = P.normalizarFactura({ numero_factura: 'FCR  00001', fecha_factura: '2026-09-15', total: 10, albaranes: [
    { numero_albaran: '446334', fecha: '2026-09-04', matricula: '', importe: -130.1, es_abono: false,
      lineas: [{ referencia: 'VARTAA8', descripcion: 'A8', importe: -130.1, albaran_origen: '01000443645' }] }] });
  assert.equal(d.numero_factura, 'FCR 00001');
  assert.equal(d.albaranes[0].es_abono, true);
  assert.equal(d.albaranes[0].lineas[0].albaran_origen, '443645');
});

test('los esquemas son JSON serializable y requieren los campos clave', () => {
  assert.ok(JSON.stringify(P.SCHEMA_ALBARAN).length > 100);
  assert.ok(P.SCHEMA_FACTURA.required.includes('albaranes'));
  assert.ok(P.PROMPT_ALBARAN.includes('reembolso'));
  assert.ok(L.round2);
});
