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
