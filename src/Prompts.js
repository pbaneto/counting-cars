/**
 * Prompts y esquemas JSON para Gemini + normalización de la respuesta (puro, testeable en Node).
 * Cambiar aquí cómo se lee un albarán o una factura NO toca el resto del código.
 */
if (typeof require !== 'undefined' && typeof normPlate === 'undefined') Object.assign(globalThis, require('./Logic.js'));

const GT = { OBJ: 'OBJECT', ARR: 'ARRAY', STR: 'STRING', NUM: 'NUMBER', BOOL: 'BOOLEAN' };

const SCHEMA_LINEA_ALBARAN = {
  type: GT.OBJ,
  properties: {
    referencia: { type: GT.STR }, descripcion: { type: GT.STR }, marca: { type: GT.STR },
    cantidad: { type: GT.NUM, nullable: true }, precio_unitario: { type: GT.NUM, nullable: true },
    descuento_pct: { type: GT.NUM, nullable: true }, importe: { type: GT.NUM, nullable: true }, reembolso: { type: GT.BOOL },
  },
  required: ['referencia', 'descripcion', 'importe', 'reembolso'],
};

const SCHEMA_ALBARAN = {
  type: GT.OBJ,
  properties: {
    es_albaran: { type: GT.BOOL },
    proveedor: { type: GT.STR, enum: ['RM', 'Otros'] },
    numero_albaran: { type: GT.STR }, fecha: { type: GT.STR }, matricula: { type: GT.STR },
    base_imponible: { type: GT.NUM, nullable: true }, iva_importe: { type: GT.NUM, nullable: true }, total: { type: GT.NUM, nullable: true },
    lineas: { type: GT.ARR, items: SCHEMA_LINEA_ALBARAN },
  },
  required: ['es_albaran', 'proveedor', 'numero_albaran', 'fecha', 'matricula', 'total', 'lineas'],
};

const PROMPT_ALBARAN = `Eres un asistente que lee albaranes de recambios de automoción escaneados. Vienen de un taller mecánico en España; el proveedor habitual es Repuestos Miguel ("RM"). Devuelve SOLO el JSON del esquema.

REGLAS
1. es_albaran: true si el documento es un albarán de entrega de recambios; false si es otra cosa (factura, hoja en blanco, ticket, foto ilegible...).
2. proveedor: "RM" si es Repuestos Miguel (logo RM, C/ Munich 7, Las Rozas); en cualquier otro caso "Otros".
3. numero_albaran: el número junto a la palabra ALBARAN (sólo dígitos, p. ej. 462446). NO lo confundas con el código de cliente, el CIF ni el nº de pedido.
4. fecha: la que sigue a "Fecha:". Formato dd/mm/aa: conviértela a yyyy-mm-dd (aa de dos cifras = 20aa). Ej.: 15/09/26 → 2026-09-15.
5. matricula: la que sigue a "Matrícula:", en mayúsculas y sin espacios ni guiones. Cópiala tal cual, sin corregirla ni inventarla. Si está vacía, "".
6. Importes: números con punto decimal, sin símbolos. Formato español 1.234,56 → 1234.56. base_imponible = casilla "Base Imp."; iva_importe = "Imp. IVA"; total = "Total" (recuadro grande abajo a la derecha, IVA incluido).
7. lineas: una por cada fila de la tabla (Artículo | Descripción | Cantidad | Pvp | % Dto | Importe) que tenga importe.
   - referencia: columna Artículo, sin espacios (en la línea de residuos SIGAUS suele estar vacía: "").
   - descripcion: columna Descripción completa de esa fila. Si continúa en una fila siguiente sin importe, únela.
   - cantidad, precio_unitario (columna Pvp), descuento_pct (60,00 → 60), importe (columna Importe, sin IVA, tal cual aparece).
   - Incluye la línea de residuos "SIGAUS". NO incluyas textos sin importe (p. ej. "ATENCION PRECIO POR PAREJA") ni la publicidad.
   - marca: fabricante de la pieza si se deduce del prefijo del código de artículo (DAYCO6PK… → DAYCO, BOSCH…, CORTECO…, KRAFF… → KRAFFT) o de la descripción; si no, "".
8. reembolso: true SÓLO si junto a la descripción de ESA línea hay una letra R escrita A MANO (bolígrafo, normalmente rojo o azul, a veces precedida de un guion largo "—R"). Ignora las R impresas y cualquier otra anotación manuscrita fuera de las líneas de pieza (números, "R10", garabatos en la cabecera). En caso de duda, false.
9. No inventes nada: si un dato no se lee, usa "" (texto) o null (número).`;

const SCHEMA_LINEA_FACTURA = {
  type: GT.OBJ,
  properties: {
    referencia: { type: GT.STR }, descripcion: { type: GT.STR },
    cantidad: { type: GT.NUM, nullable: true }, precio_unitario: { type: GT.NUM, nullable: true },
    descuento_pct: { type: GT.NUM, nullable: true }, importe: { type: GT.NUM }, albaran_origen: { type: GT.STR },
  },
  required: ['referencia', 'descripcion', 'importe', 'albaran_origen'],
};

const SCHEMA_FACTURA = {
  type: GT.OBJ,
  properties: {
    numero_factura: { type: GT.STR }, fecha_factura: { type: GT.STR },
    base_imponible: { type: GT.NUM, nullable: true }, iva_importe: { type: GT.NUM, nullable: true }, total: { type: GT.NUM, nullable: true },
    albaranes: {
      type: GT.ARR,
      items: {
        type: GT.OBJ,
        properties: {
          numero_albaran: { type: GT.STR }, fecha: { type: GT.STR }, matricula: { type: GT.STR }, importe: { type: GT.NUM },
          es_abono: { type: GT.BOOL }, lineas: { type: GT.ARR, items: SCHEMA_LINEA_FACTURA },
        },
        required: ['numero_albaran', 'fecha', 'matricula', 'importe', 'es_abono', 'lineas'],
      },
    },
  },
  required: ['numero_factura', 'fecha_factura', 'total', 'albaranes'],
};

const PROMPT_FACTURA = `Eres un asistente que lee facturas quincenales de Repuestos Miguel (RM) a un taller mecánico en España. Devuelve SOLO el JSON del esquema.

ESTRUCTURA DE LA FACTURA
- Cabecera de cada hoja: "Factura Nº" (p. ej. FCR 00001) y "Fecha Factura" (dd/mm/aaaa → yyyy-mm-dd).
- El cuerpo es una lista de BLOQUES. Cada bloque empieza con una fila en negrita: "ALBARAN Nº 437606 del 01/09/26 MATRICULA: 1233KCC 247,70". A continuación van las líneas de pieza de ese albarán (Artículo | Descripción | Cantidad | Precio | Dto. | Importe).
- Algunos bloques son ABONOS (devoluciones): importe negativo, cantidades negativas y, bajo cada línea, un texto "ABONO CORRESP. AL ALB. Nº 01000443645 FECHA 03/09/2026 15:37:51". Ese número (con su prefijo 01000) es el albaran_origen de la línea; la fecha de ese texto NO es la fecha del bloque.
- Un bloque puede partirse entre dos hojas: si una hoja termina con la fila cabecera de un albarán y las líneas aparecen al principio de la hoja siguiente, son del MISMO albarán. Las líneas sin cabecera al principio de una hoja pertenecen al último albarán de la hoja anterior. No dupliques albaranes.
- El pie de la ÚLTIMA hoja tiene: Gestión Residuos, BASE IMPONIBLE, 21 % I.V.A. y TOTAL FACTURA.

REGLAS
1. numero_factura tal cual ("FCR 00001"). base_imponible, iva_importe y total del pie de la última hoja (números con punto decimal: 3.060,39 → 3060.39).
2. albaranes: uno por bloque, en el orden de la factura.
   - numero_albaran: sólo dígitos tras "ALBARAN Nº". fecha: la del propio bloque ("del 01/09/26" → 2026-09-01). matricula: la del bloque en mayúsculas sin espacios ("" si está vacía; cópiala tal cual aunque sea rara, p. ej. "ACEITE").
   - importe: cifra a la derecha de la fila cabecera, sin IVA y con su signo. es_abono: true si el bloque es un abono.
3. lineas de cada bloque: referencia (sin espacios; vacía en la línea de residuos SIGAUS), descripcion (sólo la descripción de la pieza, sin el texto "ABONO CORRESP..."), cantidad (negativa en abonos), precio_unitario, descuento_pct (40,00 → 40), importe (con signo), albaran_origen (sólo en abonos; "" en el resto). Incluye las líneas SIGAUS.
4. No inventes nada: si no se lee, "" o null.`;

function limpiarLinea_(l) {
  const n = v => { const x = parseNumber(v); return Number.isFinite(x) ? x : null; };
  return {
    referencia: String(l.referencia || '').replace(/\s+/g, '').toUpperCase(),
    descripcion: String(l.descripcion || '').replace(/\s+/g, ' ').trim(),
    marca: String(l.marca || '').trim(),
    cantidad: n(l.cantidad), precio_unitario: n(l.precio_unitario), descuento_pct: n(l.descuento_pct), importe: n(l.importe),
    reembolso: l.reembolso === true, albaran_origen: albaranOrigen(l.albaran_origen),
  };
}

/** Respuesta cruda de Gemini (albarán) -> objeto limpio y tipado. */
function normalizarAlbaran(raw) {
  const n = v => { const x = parseNumber(v); return Number.isFinite(x) ? x : null; };
  return {
    es_albaran: raw.es_albaran !== false,
    proveedor: raw.proveedor === 'Otros' ? 'Otros' : 'RM',
    numero_albaran: normAlbaran(raw.numero_albaran).replace(/\D/g, ''),
    fecha: isoValid(raw.fecha) ? raw.fecha : '',
    matricula: normPlate(raw.matricula),
    base_imponible: n(raw.base_imponible), iva_importe: n(raw.iva_importe), total: n(raw.total),
    lineas: (raw.lineas || []).map(limpiarLinea_),
  };
}

function normalizarFactura(raw) {
  const n = v => { const x = parseNumber(v); return Number.isFinite(x) ? x : null; };
  return {
    numero_factura: String(raw.numero_factura || '').replace(/\s+/g, ' ').trim(),
    fecha_factura: isoValid(raw.fecha_factura) ? raw.fecha_factura : '',
    base_imponible: n(raw.base_imponible), iva_importe: n(raw.iva_importe), total: n(raw.total),
    albaranes: (raw.albaranes || []).map(a => ({
      numero_albaran: normAlbaran(a.numero_albaran).replace(/\D/g, ''),
      fecha: isoValid(a.fecha) ? a.fecha : '',
      matricula: normPlate(a.matricula),
      importe: n(a.importe),
      es_abono: a.es_abono === true || (n(a.importe) != null && n(a.importe) < 0),
      lineas: (a.lineas || []).map(limpiarLinea_),
    })),
  };
}

if (typeof module !== 'undefined') {
  module.exports = { SCHEMA_ALBARAN, SCHEMA_FACTURA, PROMPT_ALBARAN, PROMPT_FACTURA, normalizarAlbaran, normalizarFactura };
}
