/**
 * Nombres de pestañas, columnas y valores por defecto. UN solo sitio donde mirar.
 * Regla de oro: el código busca columnas por el NOMBRE de su cabecera (fila 1), nunca por letra.
 * "entradas" = columnas que escribe una persona o el script; el resto son fórmulas (color gris).
 */
const TZ = 'Europe/Madrid';

const HOJA = {
  ALB: 'Albaranes', TRAB: 'Trabajos', PIEZAS: 'Piezas', ABONOS: 'Abonos', COCHES: 'Coches', RESUMEN: 'Resumen',
  FACT: 'Facturas RM', LINEAS: 'Líneas RM', CONFIG: 'Config', REG: 'Registro', DIAG: 'Diagnóstico',
};

const ESQUEMA = {
  'Albaranes': {
    cabeceras: ['Fecha escaneo', 'Fecha albarán', 'Quincena', 'Proveedor', 'Nº albarán', 'Nº trabajo', 'Matrícula', 'Precio con IVA',
      'Precio facturable', 'Coche', 'Cliente', 'Ver PDF', 'Avisos', 'Nota escaneo'],
    entradas: ['Fecha escaneo', 'Fecha albarán', 'Proveedor', 'Nº albarán', 'Nº trabajo', 'Matrícula', 'Precio con IVA', 'Ver PDF', 'Nota escaneo'],
    filasFormulas: 1500,
  },
  'Trabajos': {
    cabeceras: ['Nº trabajo', 'Fecha apertura', 'Matrícula', 'Coche', 'Cliente', 'Recambios', 'Recambios facturables', 'Factura', 'Beneficio', 'Pagado', 'Avisos'],
    entradas: ['Nº trabajo', 'Fecha apertura', 'Matrícula', 'Factura', 'Pagado'],
    filasFormulas: 800,
  },
  'Piezas': {
    cabeceras: ['Reembolso', 'Matrícula', 'Nº albarán', 'Referencia pieza', 'Descripción', 'Marca', 'Cantidad', 'Precio base', 'Descuento aplicado',
      'Precio descontado sin IVA', 'Precio descontado con IVA', 'Fecha reembolso', 'Proveedor', 'Origen', 'Avisos'],
    entradas: ['Reembolso', 'Nº albarán', 'Referencia pieza', 'Descripción', 'Marca', 'Cantidad', 'Precio base', 'Descuento aplicado',
      'Precio descontado sin IVA', 'Fecha reembolso', 'Origen'],
    filasFormulas: 4000,
  },
  'Coches': {
    cabeceras: ['Matrícula', 'Cliente', 'Coche'],
    entradas: ['Matrícula', 'Cliente', 'Coche'],
    filasFormulas: 0,
  },
  'Facturas RM': {
    cabeceras: ['Nº factura', 'Fecha factura', 'Año', 'Mes', 'Quincena', 'Base imponible', 'IVA', 'Total', 'Estado', 'Ver PDF', 'Procesada el', 'Nº albaranes'],
    entradas: ['Nº factura', 'Fecha factura', 'Año', 'Mes', 'Quincena', 'Base imponible', 'IVA', 'Total', 'Estado', 'Ver PDF', 'Procesada el', 'Nº albaranes'],
    filasFormulas: 0,
  },
  'Líneas RM': {
    cabeceras: ['Nº factura', 'Nº albarán', 'Fecha albarán', 'Matrícula', 'Tipo', 'Albarán origen', 'Referencia', 'Descripción', 'Cantidad', 'Precio',
      'Descuento', 'Importe sin IVA', 'Conciliación'],
    entradas: ['Nº factura', 'Nº albarán', 'Fecha albarán', 'Matrícula', 'Tipo', 'Albarán origen', 'Referencia', 'Descripción', 'Cantidad', 'Precio',
      'Descuento', 'Importe sin IVA'],
    filasFormulas: 3000,
  },
  'Registro': {
    cabeceras: ['Fecha y hora', 'Nivel', 'Función', 'Referencia', 'Mensaje'],
    entradas: ['Fecha y hora', 'Nivel', 'Función', 'Referencia', 'Mensaje'],
    filasFormulas: 0,
  },
  'Config': {
    cabeceras: ['Clave', 'Valor', 'Descripción'],
    entradas: ['Clave', 'Valor', 'Descripción'],
    filasFormulas: 0,
  },
};

/** Valores por defecto de la pestaña Config: [clave, valor, descripción] */
const CONFIG_DEFECTO = [
  ['MODELO_GEMINI', 'gemini-3.5-flash-lite', 'Modelo de Gemini que lee albaranes y facturas'],
  ['IVA', 0.21, 'IVA aplicado a las piezas (0,21 = 21 %)'],
  ['DIAS_AVISO_TRABAJO', 30, 'Avisar si un albarán se suma a un trabajo sin pagar abierto hace más de estos días'],
  ['TOLERANCIA_CUADRE', 0.5, 'Diferencia máxima (€) para dar por cuadrada una quincena de Abonos (redondeos de IVA)'],
  ['PARALELISMO', 8, 'Nº de PDFs que se envían a Gemini a la vez'],
  ['MAX_ARCHIVOS', 30, 'Máximo de albaranes que se procesan por ejecución (límite de 6 min de Apps Script)'],
  ['CARPETA_ENTRADA', '', 'ID de la carpeta de Drive con los albaranes escaneados'],
  ['CARPETA_PROCESADOS', '', 'ID de la carpeta de Drive donde pasan los albaranes procesados'],
  ['CARPETA_FACTURAS_RM', '', 'ID de la carpeta con las facturas quincenales de RM por procesar'],
  ['CARPETA_FACTURAS_RM_PROCESADAS', '', 'ID de la carpeta de facturas RM ya procesadas'],
  ['CARPETA_ERRORES', '', 'ID de la carpeta de PDFs que no se han podido leer (se crea sola)'],
];

/** Posiciones fijas de la pestaña Abonos (resumen arriba, tabla grande debajo). */
const ABONOS = {
  celdaAnio: 'B1', filaCabResumen: 3, filaIni: 4, filas: 24, filaTitulo: 29, filaCabTabla: 30, filaTabla: 31, maxTabla: 3000,
  cabResumen: ['Mes', 'Quincena', 'Recambios totales RM', 'Reembolso solicitado', 'Reembolso abonado', 'Recambios − solicitado',
    'Total factura RM', 'Estado', 'Reescanear factura', 'Desde', 'Hasta'],
  cabTabla: ['Fecha abono', 'Descripción pieza', 'Precio sin IVA', 'Precio con IVA', 'Estado', 'Nº albarán', 'Referencia', 'Matrícula',
    'Fecha solicitud', 'Factura RM', 'Nota'],
};

const ESTADOS_ABONO = ['Abonada', 'Sin abonar', 'Sin solicitar'];
const COLORES = {
  cabecera: '#1f3a5f', gris: '#efefef', rojo: '#f4cccc', verde: '#d9ead3', amarillo: '#fff2cc', naranja: '#fce5cd', azul: '#cfe2f3',
};
