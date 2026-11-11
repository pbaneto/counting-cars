/**
 * Nombres de pestañas, columnas y valores por defecto. UN solo sitio donde mirar.
 * Regla de oro: el código busca columnas por el NOMBRE de su cabecera (fila 1), nunca por letra.
 * "entradas" = columnas que escribe una persona o el script; el resto son fórmulas (color gris).
 * "casillas" = columnas con casilla de verificación. Como las fórmulas, sólo van en filas con datos: una casilla
 * vale siempre TRUE/FALSE, así que en filas vacías hace que Sheets cuente miles de filas "ocupadas" al leer.
 */
const TZ = 'Europe/Madrid';

/**
 * Versión del código (MAYOR.MENOR.PARCHE). Al cambiar algo, súbela:
 *  - MAYOR: cambia dónde están los datos (columna nueva con datos, claves…): añade su migración en Version.js.
 *  - MENOR: sólo diseño (formato, colores, fórmulas, paneles): la hoja lo reaplica sola.
 *  - PARCHE: arreglo de código que no toca la hoja.
 */
const VERSION = '9.0.0';

const HOJA = {
  ALB: 'Albaranes', TRAB: 'Trabajos', PIEZAS: 'Piezas', ABONOS: 'Abonos', CUADRE: 'Cuadre RM', COCHES: 'Coches', RESUMEN: 'Resumen',
  FACT: 'Facturas RM', LINEAS: 'Líneas RM', CONFIG: 'Config', REG: 'Registro', DIAG: 'Diagnóstico',
};

const ESQUEMA = {
  'Albaranes': {
    cabeceras: ['Mes', 'Quincena', 'Proveedor', 'Coche', 'Cliente', 'Matrícula', 'Nº trabajo', 'Precio con IVA', 'Precio - abonos',
      'Fecha escaneo', 'Fecha albarán', 'Nº albarán', 'Ver PDF', 'Avisos', 'Nota escaneo'],
    visibles: 9,  // las primeras n cabeceras; el resto se ocultan (ver ocultarColumnas_)
    entradas: ['Fecha escaneo', 'Fecha albarán', 'Proveedor', 'Nº albarán', 'Nº trabajo', 'Matrícula', 'Precio con IVA', 'Ver PDF', 'Nota escaneo'],
    filasFormato: 1500,
  },
  'Trabajos': {
    // Mes y Quincena: valores (no fórmulas) para poder cambiarlos a mano; se rellenan con la fecha del primer albarán.
    cabeceras: ['Mes', 'Quincena', 'Coche', 'Cliente', 'Matrícula', 'Recambios', 'Precio - abono', 'Factura', 'Beneficio', 'Pagado',
      'Nº trabajo', 'Recambios facturables RM', 'Recambios facturables Otros', 'Avisos'],
    visibles: 10,
    entradas: ['Nº trabajo', 'Mes', 'Quincena', 'Matrícula', 'Factura', 'Pagado'],
    casillas: ['Pagado'],
    filasFormato: 800,
    filaCabecera: 4,  // filas 1-3: panel "Resumen (según filtro)" encima de la cabecera real
    cabeceraMovil: true,  // puede bajar si se insertan filas entre el panel y la cabecera (ver filaCabecera_)
  },
  'Piezas': {
    // ID: número fijo de cada pieza, lo pone el script (ver idsPiezas_). No cuenta como dato: una fila vaciada a mano que
    // conserve su ID es una fila vacía.
    cabeceras: ['Reembolso', 'Matrícula', 'Nº albarán', 'Referencia pieza', 'Descripción', 'Marca', 'Cantidad', 'Precio base', 'Descuento aplicado',
      'Precio descontado sin IVA', 'Precio descontado con IVA', 'Fecha reembolso', 'Abonada a mano', 'Proveedor', 'Origen', 'Avisos', 'ID'],
    entradas: ['Reembolso', 'Nº albarán', 'Referencia pieza', 'Descripción', 'Marca', 'Cantidad', 'Precio base', 'Descuento aplicado',
      'Precio descontado sin IVA', 'Fecha reembolso', 'Abonada a mano', 'Origen'],
    casillas: ['Reembolso'],
    filasFormato: 4000,
  },
  'Coches': {
    cabeceras: ['Matrícula', 'Cliente', 'Coche'],
    entradas: ['Matrícula', 'Cliente', 'Coche'],
    filasFormato: 0,
  },
  'Facturas RM': {
    cabeceras: ['Nº factura', 'Fecha factura', 'Año', 'Mes', 'Quincena', 'Base imponible', 'IVA', 'Total', 'Estado', 'Ver PDF', 'Procesada el', 'Nº albaranes'],
    entradas: ['Nº factura', 'Fecha factura', 'Año', 'Mes', 'Quincena', 'Base imponible', 'IVA', 'Total', 'Estado', 'Ver PDF', 'Procesada el', 'Nº albaranes'],
    filasFormato: 0,
  },
  'Líneas RM': {
    // Líneas de las facturas RM y de los albaranes de ABONO escaneados (estos, sin Nº factura hasta que llega su factura).
    // En las de abono: "ID pieza" = la pieza abonada (la pone el script; se puede corregir a mano, "NO" = sin pieza) y
    // "Fecha abono" = la del albarán de abono si se escaneó antes, o la de la factura.
    cabeceras: ['Nº factura', 'Nº albarán', 'Fecha albarán', 'Matrícula', 'Tipo', 'Albarán origen', 'Referencia', 'Descripción', 'Cantidad', 'Precio',
      'Descuento', 'Importe sin IVA', 'Conciliación', 'ID pieza', 'Fecha abono'],
    entradas: ['Nº factura', 'Nº albarán', 'Fecha albarán', 'Matrícula', 'Tipo', 'Albarán origen', 'Referencia', 'Descripción', 'Cantidad', 'Precio',
      'Descuento', 'Importe sin IVA', 'ID pieza', 'Fecha abono'],
    filasFormato: 3000,
  },
  'Registro': {
    cabeceras: ['Fecha y hora', 'Nivel', 'Función', 'Referencia', 'Mensaje'],
    entradas: ['Fecha y hora', 'Nivel', 'Función', 'Referencia', 'Mensaje'],
    filasFormato: 0,
  },
  'Config': {
    cabeceras: ['Clave', 'Valor', 'Descripción'],
    entradas: ['Clave', 'Valor', 'Descripción'],
    filasFormato: 0,
  },
};

/** Valores por defecto de la pestaña Config: [clave, valor, descripción] */
const CONFIG_DEFECTO = [
  ['MODELO_GEMINI', 'gemini-3.5-flash-lite', 'Modelo de Gemini que lee albaranes y facturas'],
  ['URL_GEMINI', 'https://generativelanguage.googleapis.com', 'Base de la API de Gemini. Para medir el gasto con LiteLLM: https://<servidor-litellm>/gemini (y su clave virtual como API key)'],
  ['IVA', 0.21, 'IVA aplicado a las piezas (0,21 = 21 %)'],
  ['DIAS_AVISO_REEMBOLSO', 45, 'Avisar en Abonos si una pieza pedida lleva más de estos días sin que RM la abone'],
  ['ABONOS_DESDE', '', 'Abonos sólo enseña lo abonado o pedido desde esta fecha (vacía = todo). Lo pendiente de abonar sale siempre'],
  ['PARALELISMO', 8, 'Nº de PDFs que se envían a Gemini a la vez'],
  ['MAX_ARCHIVOS', 30, 'Máximo de albaranes que se procesan por ejecución (límite de 6 min de Apps Script)'],
  ['CARPETA_ENTRADA', '', 'ID de la carpeta de Drive con los albaranes escaneados'],
  ['CARPETA_PROCESADOS', '', 'ID de la carpeta de Drive donde pasan los albaranes procesados'],
  ['CARPETA_FACTURAS_RM', '', 'ID de la carpeta con las facturas quincenales de RM por procesar'],
  ['CARPETA_FACTURAS_RM_PROCESADAS', '', 'ID de la carpeta de facturas RM ya procesadas'],
  ['CARPETA_ERRORES', '', 'ID de la carpeta de PDFs que no se han podido leer (se crea sola)'],
];

/**
 * Pestaña Abonos: sólo la tabla, un informe que el script rehace entero desde Piezas y Líneas RM (ver Abonos.js).
 * Cabecera en la fila 1 (fija).
 */
const ABONOS = {
  filaCabTabla: 1, filaTabla: 2, maxTabla: 3000,
  cabTabla: ['Fecha abono', 'Descripción pieza', 'Precio sin IVA', 'Precio con IVA', 'Estado', 'Nº albarán', 'Referencia', 'Matrícula',
    'Fecha solicitud', 'Factura RM', 'Nota', 'Días pendiente', 'ID pieza'],  // ID pieza: columna oculta
};

/**
 * Pestaña Cuadre RM: resumen por quincena (filas 4-27; el año, en B1) y, a su derecha, el panel "Pendientes de RM", que
 * no depende de la quincena en que se pidió el reembolso. Todo son fórmulas sobre Abonos, Albaranes y Facturas RM.
 */
const CUADRE = {
  celdaAnio: 'B1', filaCab: 3, filaIni: 4, filas: 24, panelCol: 8,  // H (etiqueta) / I (valor)
  cab: ['Mes', 'Quincena', 'Recambios totales RM', 'Reembolso abonado', 'Total factura RM', 'Diferencia'],
};

const ESTADOS_ABONO = ['Abonada', 'Abonada (a mano)', 'Sin abonar', 'Sin solicitar'];
const COLORES = {
  cabecera: '#1f3a5f', gris: '#efefef', rojo: '#f4cccc', verde: '#d9ead3', amarillo: '#fff2cc', naranja: '#fce5cd', azul: '#cfe2f3',
};
// Estados de Abonos: colores vivos que se distinguen también con daltonismo (protanopía, deuteranopía y tritanopía)
// y se leen con texto negro.
const COLORES_ABONOS = {
  abonada: '#00e5ff', abonadaMano: '#b388ff', sinSolicitar: '#ffff00', sinAbonar: '#ff9100', sinAbonarVencida: '#ff1744',
};
