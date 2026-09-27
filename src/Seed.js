/**
 * Carga de datos iniciales desde src/private.js (que NO está en el repositorio): coches y piloto de la 1ª quincena de septiembre.
 * Sólo rellena pestañas que estén vacías: no pisa nada.
 */

function cargarDatosIniciales() {
  ejecutar_('cargarDatosIniciales', () => conBloqueo_(30, () => {
    if (typeof PRIVATE === 'undefined') throw new Error('No encuentro src/private.js con los datos iniciales (ver private.example.js).');
    const out = [];
    out.push(cargarCoches_());
    out.push(cargarPiloto_());
    avisar_(out.join('\n'), 'Datos iniciales');
  }));
}

function cargarCoches_() {
  const tab = leerTabla_(HOJA.COCHES);
  if (tab.filas.length) return `Coches: ya hay ${tab.filas.length} filas, no se toca.`;
  const filas = (PRIVATE.COCHES || []).map(c => ({ 'Matrícula': normPlate(c[0]), 'Cliente': c[1], 'Coche': c[2] }));
  tab.sh.getRange(2, tab.map['Matrícula'], filas.length, 1).setNumberFormat('@');
  agregarFilas_(tab, filas);
  log_('INFO', 'cargarDatosIniciales', HOJA.COCHES, `${filas.length} coches cargados`);
  return `Coches: ${filas.length} cargados.`;
}

/** Piloto: albaranes y trabajos de la 1ª quincena del Excel antiguo (sin nº de albarán ni PDF: se enlazarán al escanear). */
function cargarPiloto_() {
  const p = PRIVATE.PILOTO;
  if (!p) return 'Piloto: sin datos.';
  const alb = leerTabla_(HOJA.ALB), trab = leerTabla_(HOJA.TRAB);
  if (alb.filas.length || trab.filas.length) return `Piloto: Albaranes/Trabajos ya tienen datos, no se toca.`;
  const fecha = aFecha_(p.fecha), nums = [], jobDe = {};
  const trabajos = p.trabajos.map(t => {
    const plate = normPlate(t.matricula), num = nextJobNumber(jobPrefix(plate), nums);
    nums.push(num); jobDe[plate] = num;
    return { 'Nº trabajo': num, 'Fecha apertura': fecha, 'Matrícula': plate, 'Factura': t.factura == null ? '' : t.factura, 'Pagado': t.pagado === true };
  });
  p.albaranes.forEach(a => {
    const plate = normPlate(a.matricula);
    if (!jobDe[plate]) {
      const num = nextJobNumber(jobPrefix(plate), nums);
      nums.push(num); jobDe[plate] = num;
      trabajos.push({ 'Nº trabajo': num, 'Fecha apertura': fecha, 'Matrícula': plate, 'Pagado': false });
    }
  });
  agregarFilas_(trab, trabajos);
  const filas = p.albaranes.map(a => ({
    'Fecha escaneo': fecha, 'Fecha albarán': fecha, 'Proveedor': a.proveedor === 'Otros' ? 'Otros' : 'RM', 'Nº trabajo': jobDe[normPlate(a.matricula)],
    'Matrícula': normPlate(a.matricula), 'Precio con IVA': a.importe, 'Nota escaneo': 'Cargado del Excel antiguo (piloto 1ª quincena): sin nº de albarán ni PDF',
  }));
  agregarFilas_(alb, filas);
  log_('INFO', 'cargarDatosIniciales', HOJA.ALB, `Piloto: ${filas.length} albaranes y ${trabajos.length} trabajos`);
  return `Piloto: ${filas.length} albaranes y ${trabajos.length} trabajos cargados.`;
}

/**
 * Repara Trabajos tras el bug de la columna "Quincena" fantasma (commit 0ec8e88): las cabeceras se
 * habían desplazado una columna respecto a los datos ya escritos, así que "Matrícula" apuntaba al dato
 * viejo de Quincena, "Factura" al viejo Recambios facturables y "Pagado"/"Beneficio" a texto de fórmulas.
 * Recupera Matrícula cruzando con Albaranes (que no se tocó) y Factura/Pagado desde los datos originales
 * del piloto en src/private.js, y quita la columna sobrante que dejó el desplazamiento.
 * A propósito NO está en el menú: ejecutar sólo una vez desde el editor de Apps Script. Volver a lanzarlo
 * después de tener Factura/Pagado reales pisaría esos datos con los del piloto.
 */
function repararTrabajos_() {
  ejecutar_('repararTrabajos', () => conBloqueo_(30, () => {
    if (typeof PRIVATE === 'undefined' || !PRIVATE.PILOTO) throw new Error('No encuentro src/private.js con los datos del piloto.');
    const tab = leerTabla_(HOJA.TRAB), alb = leerTabla_(HOJA.ALB);
    const matriculaDe = {};
    alb.filas.forEach(a => {
      const n = String(a.v['Nº trabajo'] || '').trim();
      if (n && !matriculaDe[n]) matriculaDe[n] = normPlate(a.v['Matrícula']);
    });
    const piloto = {};
    (PRIVATE.PILOTO.trabajos || []).forEach(t => {
      const p = normPlate(t.matricula);
      if (!(p in piloto)) piloto[p] = { pagado: t.pagado === true, factura: t.factura == null ? '' : t.factura };
    });

    let arregladas = 0;
    const sinAlbaran = [];
    tab.filas.forEach(t => {
      const num = String(t.v['Nº trabajo'] || '').trim();
      const plate = matriculaDe[num];
      if (!plate) { sinAlbaran.push(num); return; }
      const datos = piloto[plate] || { pagado: false, factura: '' };
      actualizarFila_(tab, t.fila, { 'Matrícula': plate, 'Factura': datos.factura, 'Pagado': datos.pagado });
      arregladas++;
    });

    const nCab = ESQUEMA['Trabajos'].cabeceras.length, extra = tab.sh.getLastColumn() - nCab;
    if (extra > 0) tab.sh.deleteColumns(nCab + 1, extra);

    reiniciarCaches_();
    prepararTablas_();
    const detalle = `${arregladas} trabajo(s) reparado(s)` + (sinAlbaran.length ? `; sin albarán para cruzar (revisar a mano): ${sinAlbaran.join(', ')}` : '');
    log_('INFO', 'repararTrabajos', '', detalle);
    avisar_(detalle, 'Reparar Trabajos');
  }));
}
