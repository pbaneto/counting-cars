/**
 * Copia este archivo a src/private.js y rellénalo. src/private.js está en .gitignore: NO se sube a GitHub,
 * pero `clasp push` sí lo sube a tu proyecto privado de Apps Script.
 */
const PRIVATE = {
  SPREADSHEET_ID: '<ID de la hoja de cálculo>',
  CARPETAS: {
    CARPETA_ENTRADA: '<ID carpeta Drive Entrada>',
    CARPETA_PROCESADOS: '<ID carpeta Drive Procesados>',
    CARPETA_FACTURAS_RM: '<ID carpeta facturas RM>',
    CARPETA_FACTURAS_RM_PROCESADAS: '<ID carpeta facturas RM procesadas>',
  },
  COCHES: [['1234ABC', 'Cliente', 'Marca modelo']],
  PILOTO: { fecha: '2026-09-15', albaranes: [{ proveedor: 'RM', matricula: '1234ABC', importe: 100 }], trabajos: [{ matricula: '1234ABC', pagado: false, factura: null }] },
  RESUMEN: {
    anio: 2026, mesEnVivo: 9,
    y2026: [[0, 0, 0, 0]],   // 12 filas [recambios, morosos, ingreso, hipoteca]
    y2025: [], y2024: [],    // 12 filas [rec, mor, ing, hip, beneficio, beneficio con gastos] / [rec, mor, ing, beneficio]
    fijos: { banco: [['Concepto', 0]], gastos: [['Concepto', 0]] },
  },
};
