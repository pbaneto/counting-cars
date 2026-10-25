/** Piezas: filas desde un albarán escaneado, marcado de reembolso e ID fijo de cada pieza. */

function filaPieza_(l, num, hoy, origen) {
  return {
    'Reembolso': l.reembolso === true, 'Nº albarán': num, 'Referencia pieza': l.referencia, 'Descripción': l.descripcion, 'Marca': l.marca || '',
    'Cantidad': l.cantidad == null ? '' : l.cantidad, 'Precio base': l.precio_unitario == null ? '' : l.precio_unitario,
    'Descuento aplicado': l.descuento_pct == null ? '' : l.descuento_pct / 100,
    'Precio descontado sin IVA': l.importe, 'Fecha reembolso': l.reembolso === true ? aFecha_(hoy) : '', 'Origen': origen,
  };
}

/** Filas nuevas de Piezas para las líneas de un albarán, cada una con su ID. */
function filasPiezas_(tabP, lineas, num, hoy, origen) {
  const ids = reservarIdsPieza_(tabP, lineas.length);
  return lineas.map((l, k) => Object.assign(filaPieza_(l, num, hoy, origen), { 'ID': ids[k] }));
}

function piezasDeAlbaran_(tabPiezas, num) { return tabPiezas.filas.filter(p => normAlbaran(p.v['Nº albarán']) === num); }

const PROP_ULTIMO_ID_ = 'ULTIMO_ID_PIEZA';

/**
 * n IDs nuevos. Nunca se reutiliza uno, aunque se borre la pieza que lo tenía: una línea de abono que apuntaba a ella
 * no debe pasar a otra pieza. tabP: Piezas ya leída (cuenta también los IDs escritos a mano).
 */
function reservarIdsPieza_(tabP, n) {
  if (!n) return [];
  const props = PropertiesService.getScriptProperties();
  let ult = Number(props.getProperty(PROP_ULTIMO_ID_)) || 0;
  tabP.filas.forEach(f => { const id = Number(f.v['ID']); if (Number.isFinite(id) && id > ult) ult = id; });
  const ids = [];
  for (let k = 0; k < n; k++) ids.push(++ult);
  props.setProperty(PROP_ULTIMO_ID_, String(ult));
  return ids;
}

/**
 * Pone ID a las piezas que no lo tienen (escritas a mano) o lo tienen repetido (fila copiada y pegada: la primera se
 * queda el suyo). Una sola escritura. Devuelve cuántas ha cambiado.
 */
function asegurarIdsPiezas_(tabP) {
  const vistos = new Set(), faltan = [];
  tabP.filas.forEach(f => {
    const id = String(f.v['ID'] == null ? '' : f.v['ID']).trim();
    if (!id || vistos.has(id)) faltan.push(f); else vistos.add(id);
  });
  if (!faltan.length) return 0;
  const ids = reservarIdsPieza_(tabP, faltan.length), c = tabP.map['ID'];
  faltan.forEach((f, k) => { f.v['ID'] = ids[k]; });
  const desde = faltan[0].fila, hasta = faltan[faltan.length - 1].fila, porFila = {};
  tabP.filas.forEach(f => { porFila[f.fila] = f.v['ID']; });
  const col = [];
  for (let r = desde; r <= hasta; r++) col.push([r in porFila ? porFila[r] : ((tabP.valores[r - 1] || [])[c - 1] == null ? '' : tabP.valores[r - 1][c - 1])]);
  tabP.sh.getRange(desde, c, col.length, 1).setValues(col);
  return faltan.length;
}
