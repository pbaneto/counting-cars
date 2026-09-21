/** Diagnóstico: recorre todas las pestañas y lista los problemas con un enlace a la fila afectada. */

function diagnostico() {
  ejecutar_('diagnostico', () => {
    reiniciarCaches_();
    const problemas = [];
    const ss = ss_();
    const add = (sev, hoja, fila, texto) => problemas.push([sev, hoja, fila, texto]);

    const coches = leerTabla_(HOJA.COCHES), alb = leerTabla_(HOJA.ALB), trab = leerTabla_(HOJA.TRAB), piezas = leerTabla_(HOJA.PIEZAS);
    const fact = leerTabla_(HOJA.FACT), lin = leerTabla_(HOJA.LINEAS);
    const cochesSet = new Set(coches.filas.map(f => normPlate(f.v['Matrícula'])));
    const jobs = {}; trab.filas.forEach(f => { jobs[String(f.v['Nº trabajo']).trim()] = f; });
    const dias = cfgNum_('DIAS_AVISO_TRABAJO'), hoy = hoyISO_();

    // Albaranes
    const numsAlb = {};
    alb.filas.forEach(f => {
      const v = f.v, plate = normPlate(v['Matrícula']), num = normAlbaran(v['Nº albarán']), job = String(v['Nº trabajo']).trim();
      if (v['Precio con IVA'] === '' || !(parseNumber(v['Precio con IVA']) > 0)) add('ERROR', HOJA.ALB, f.fila, 'Falta el precio con IVA');
      if (!plate) add('ERROR', HOJA.ALB, f.fila, 'Falta la matrícula' + (num ? ` (albarán ${num})` : ''));
      else {
        if (!cochesSet.has(plate)) add('AVISO', HOJA.ALB, f.fila, `La matrícula ${plate} no está en Coches`);
        if (!job) add('ERROR', HOJA.ALB, f.fila, 'Sin nº de trabajo');
        else if (!jobs[job]) add('ERROR', HOJA.ALB, f.fila, `El trabajo ${job} no existe en Trabajos`);
        else {
          if (normPlate(jobs[job].v['Matrícula']) !== plate) add('ERROR', HOJA.ALB, f.fila, `El trabajo ${job} es de otra matrícula (${jobs[job].v['Matrícula']})`);
          const ap = aISO_(jobs[job].v['Fecha apertura']), es = aISO_(v['Fecha escaneo']);
          if (ap && es && jobs[job].v['Pagado'] !== true && daysBetween(ap, es) > dias) add('AVISO', HOJA.ALB, f.fila, `Sumado al trabajo ${job}, abierto hace ${daysBetween(ap, es)} días: ¿es un trabajo nuevo?`);
        }
      }
      if (num) { if (numsAlb[num]) add('ERROR', HOJA.ALB, f.fila, `Nº de albarán ${num} duplicado (también en fila ${numsAlb[num]})`); else numsAlb[num] = f.fila; }
      if (String(v['Nota escaneo'] || '').trim()) add('INFO', HOJA.ALB, f.fila, 'Nota del escaneo: ' + v['Nota escaneo']);
    });

    // Trabajos
    trab.filas.forEach(f => {
      const num = String(f.v['Nº trabajo']).trim();
      if (!num) add('ERROR', HOJA.TRAB, f.fila, 'Falta el nº de trabajo');
      if (!alb.filas.some(a => String(a.v['Nº trabajo']).trim() === num)) add('INFO', HOJA.TRAB, f.fila, `El trabajo ${num} no tiene albaranes`);
    });

    // Piezas
    piezas.filas.forEach(f => {
      const num = normAlbaran(f.v['Nº albarán']);
      if (f.v['Reembolso'] === true && !num) add('ERROR', HOJA.PIEZAS, f.fila, 'Reembolso marcado sin nº de albarán');
      if (num && !numsAlb[num]) add('AVISO', HOJA.PIEZAS, f.fila, `El albarán ${num} no está en Albaranes`);
      if (f.v['Reembolso'] === true && f.v['Precio descontado sin IVA'] === '') add('ERROR', HOJA.PIEZAS, f.fila, 'Reembolso marcado sin precio');
    });

    // Facturas y líneas
    fact.filas.forEach(f => { if (String(f.v['Estado']).indexOf('OK') !== 0) add('AVISO', HOJA.FACT, f.fila, `Factura ${f.v['Nº factura']}: ${f.v['Estado']}`); });
    const noEsc = {};
    lin.filas.forEach(f => {
      const c = String(f.v['Conciliación'] || '');
      if (c.indexOf('no escaneado') > 0 && !noEsc[f.v['Nº albarán']]) { noEsc[f.v['Nº albarán']] = 1; add('AVISO', HOJA.LINEAS, f.fila, `Albarán ${f.v['Nº albarán']} (${f.v['Matrícula'] || 'sin matrícula'}) está en la factura ${f.v['Nº factura']} pero no en Albaranes`); }
      if (c.indexOf('Importe distinto') > 0 && !noEsc['d' + f.v['Nº albarán']]) { noEsc['d' + f.v['Nº albarán']] = 1; add('AVISO', HOJA.LINEAS, f.fila, `Albarán ${f.v['Nº albarán']}: el importe de la factura no coincide con Albaranes`); }
    });

    // Carpetas pendientes
    ['CARPETA_ENTRADA', 'CARPETA_FACTURAS_RM'].forEach(k => {
      try { const n = listarArchivos_(k, 0).length; if (n) add('INFO', 'Drive', '', `${n} archivo(s) esperando en ${k}`); }
      catch (e) { add('ERROR', HOJA.CONFIG, '', e.message); }
    });
    if (!PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY')) add('ERROR', 'Configuración', '', 'Falta la API key de Gemini (menú Configurar API key)');

    escribirDiagnostico_(problemas, ss);
    log_('INFO', 'diagnostico', '', `${problemas.length} hallazgos`);
    toast_(problemas.length ? `${problemas.length} hallazgos. Mira la pestaña Diagnóstico.` : 'Todo en orden ✔');
  });
}

function escribirDiagnostico_(problemas, ss) {
  let sh = ss.getSheetByName(HOJA.DIAG) || ss.insertSheet(HOJA.DIAG);
  sh.clear();
  sh.getRange(1, 1, 1, 4).setValues([['Gravedad', 'Pestaña', 'Fila', 'Problema']]).setBackground(COLORES.cabecera).setFontColor('#ffffff').setFontWeight('bold');
  sh.setFrozenRows(1);
  const orden = { ERROR: 0, AVISO: 1, INFO: 2 };
  problemas.sort((a, b) => orden[a[0]] - orden[b[0]]);
  if (problemas.length) {
    const filas = problemas.map(p => {
      const dest = ss.getSheetByName(p[1]);
      const enlace = dest && p[2] ? loc_(`=HYPERLINK("#gid=${dest.getSheetId()}&range=A${p[2]}","Ir a fila ${p[2]}")`) : (p[2] || '');
      return [p[0], p[1], enlace, p[3]];
    });
    sh.getRange(2, 1, filas.length, 4).setValues(filas);
    const col = { ERROR: COLORES.rojo, AVISO: COLORES.naranja, INFO: COLORES.azul };
    sh.getRange(2, 1, filas.length, 1).setBackgrounds(problemas.map(p => [col[p[0]]]));
  } else sh.getRange(2, 1).setValue('✔ Sin problemas');
  sh.setColumnWidth(1, 80); sh.setColumnWidth(2, 110); sh.setColumnWidth(3, 110); sh.setColumnWidth(4, 800);
  sh.setTabColor('#cc0000');
}

/** Comprueba, sin escribir nada, que esta cuenta puede abrir la hoja, las carpetas de Drive y que hay API key. Muestra el informe en el registro de ejecución (editor) o en un cuadro (hoja). */
function probarAcceso() {
  const out = [];
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID') || (typeof PRIVATE !== 'undefined' && PRIVATE.SPREADSHEET_ID) || '';
  out.push(`Cuenta: ${Session.getEffectiveUser().getEmail()}`);
  try { out.push(`Hoja: OK "${SpreadsheetApp.openById(id).getName()}"`); } catch (e) { out.push(`Hoja: NO accesible (${e.message})`); }
  const carpetas = (typeof PRIVATE !== 'undefined' && PRIVATE.CARPETAS) || {};
  Object.keys(carpetas).forEach(k => {
    try { out.push(`${k}: OK "${DriveApp.getFolderById(carpetas[k]).getName()}"`); } catch (e) { out.push(`${k}: NO accesible (${e.message})`); }
  });
  out.push(`API key de Gemini: ${PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY') ? 'guardada' : 'FALTA'}`);
  const informe = out.join('\n');
  Logger.log(informe);  // en el editor: se ve en "Registro de ejecución" (abajo)
  avisar_(informe, 'Comprobación de acceso');  // desde la hoja: cuadro emergente
  return informe;
}
