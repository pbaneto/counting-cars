/** Piezas: filas desde un albarán escaneado y marcado de reembolso. */

function filaPieza_(l, num, hoy, origen) {
  return {
    'Reembolso': l.reembolso === true, 'Nº albarán': num, 'Referencia pieza': l.referencia, 'Descripción': l.descripcion, 'Marca': l.marca || '',
    'Cantidad': l.cantidad == null ? '' : l.cantidad, 'Precio base': l.precio_unitario == null ? '' : l.precio_unitario,
    'Descuento aplicado': l.descuento_pct == null ? '' : l.descuento_pct / 100,
    'Precio descontado sin IVA': l.importe, 'Fecha reembolso': l.reembolso === true ? aFecha_(hoy) : '', 'Origen': origen,
  };
}

function piezasDeAlbaran_(tabPiezas, num) { return tabPiezas.filas.filter(p => normAlbaran(p.v['Nº albarán']) === num); }
