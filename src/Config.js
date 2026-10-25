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
const VERSION = '6.0.0';

const HOJA = {
  ALB: 'Albaranes', TRAB: 'Trabajos', PIEZAS: 'Piezas', ABONOS: 'Abonos', COCHES: 'Coches', RESUMEN: 'Resumen',
  FACT: 'Facturas RM', LINEAS: 'Líneas RM', CONFIG: 'Config', REG: 'Registro', DIAG: 'Diagnóstico',
};

const ESQUEMA = {
  'Albaranes': {
    cabeceras: ['Fecha escaneo', 'Fecha albarán', 'Mes', 'Quincena', 'Proveedor', 'Nº albarán', 'Nº trabajo', 'Matrícula', 'Precio con IVA',
      'Precio - abonos', 'Coche', 'Cliente', 'Ver PDF', 'Avisos', 'Nota escaneo'],
    entradas: ['Fecha escaneo', 'Fecha albarán', 'Proveedor', 'Nº albarán', 'Nº trabajo', 'Matrícula', 'Precio con IVA', 'Ver PDF', 'Nota escaneo'],
    filasFormato: 1500,
  },
  'Trabajos': {
    // Mes y Quincena: valores (no fórmulas) para poder cambiarlos a mano; se rellenan con la fecha del primer albarán.
    cabeceras: ['Nº trabajo', 'Mes', 'Quincena', 'Matrícula', 'Coche', 'Cliente', 'Recambios', 'Precio - abono', 'Recambios facturables RM',
      'Recambios facturables Otros', 'Factura', 'Beneficio', 'Pagado', 'Avisos'],
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
 * Posiciones fijas de la pestaña Abonos. La tabla grande va arriba del todo (cabecera en la fila 1, que queda fija) en
 * las columnas A-M; el resumen por quincena y el panel de pendientes van a su derecha (columnas O en adelante).
 * La tabla es un informe: el script la rehace entera desde Piezas y Líneas RM (ver Abonos.js).
 */
const ABONOS = {
  filaCabTabla: 1, filaTabla: 2, maxTabla: 3000,
  colResumen: 15, celdaAnio: 'P1', filaCabResumen: 3, filaIni: 4, filas: 24,  // O: Mes, P: Quincena … T: Diferencia
  panelCol: 22, // V (etiqueta) / W (valor): panel "Pendientes de RM", no depende de la quincena en que se pidió el reembolso
  cabResumen: ['Mes', 'Quincena', 'Recambios totales RM', 'Reembolso abonado', 'Total factura RM', 'Diferencia'],
  cabTabla: ['Fecha abono', 'Descripción pieza', 'Precio sin IVA', 'Precio con IVA', 'Estado', 'Nº albarán', 'Referencia', 'Matrícula',
    'Fecha solicitud', 'Factura RM', 'Nota', 'Días pendiente', 'ID pieza'],  // ID pieza: columna oculta
};

const ESTADOS_ABONO = ['Abonada', 'Abonada (a mano)', 'Sin abonar', 'Sin solicitar'];
const COLORES = {
  cabecera: '#1f3a5f', gris: '#efefef', rojo: '#f4cccc', verde: '#d9ead3', amarillo: '#fff2cc', naranja: '#fce5cd', azul: '#cfe2f3',
};
