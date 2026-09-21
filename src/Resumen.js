/**
 * Pestaña Resumen (equivale a la pestaña TOTAL del Excel antiguo).
 *  - Año en curso: meses anteriores a MES_EN_VIVO = valores importados del Excel antiguo; el resto se calcula desde Trabajos
 *    (un trabajo se imputa al mes de su Fecha de apertura).
 *  - Años anteriores (2025, 2024): valores históricos.
 *  - 3 gráficos de barras con el beneficio mensual.
 * Los importes históricos vienen de src/private.js (no está en el repositorio).
 */

const RES = {
  anioFila: 3, cabFila: 4, ini: 5, total: 17,
  cab: ['Mes', 'Recambios', 'Morosos', 'Ingreso', 'Hipoteca', 'Beneficio', 'Beneficio contando gastos fijos', 'Banco fin de mes'],
  y25: { anio: 20, cab: 21, ini: 22, total: 34 }, y24: { anio: 37, cab: 38, ini: 39, total: 51 },
  fijos: { fila: 4, colBanco: 10 },  // J4
};

function datosResumen_() {
  const d = (typeof PRIVATE !== 'undefined' && PRIVATE.RESUMEN) || {};
  const vacio = () => Array.from({ length: 12 }, () => [null, null, null, null, null, null]);
  return {
    anio: d.anio || 2026, mesEnVivo: d.mesEnVivo || 9,
    y26: d.y2026 || vacio(), y25: d.y2025 || vacio(), y24: d.y2024 || vacio(),
    banco: d.fijos ? d.fijos.banco : [['Autónomo', 0]], gastos: d.fijos ? d.fijos.gastos : [['Gastos', 0]],
  };
}

function montarResumen_() {
  const sh = hoja_(HOJA.RESUMEN), d = datosResumen_(), t = letras_(HOJA.TRAB);
  limpiarProtecciones_(sh);
  sh.getCharts().forEach(c => sh.removeChart(c));
  sh.getRange('A1').setValue('Resumen').setFontSize(16).setFontWeight('bold');
  sh.getRange('C1').setValue('Ingresos, recambios y beneficio por mes. Los trabajos se imputan al mes de su fecha de apertura.').setFontStyle('italic').setFontColor('#666666');

  const estiloCab = (fila, n) => sh.getRange(fila, 1, 1, n).setBackground(COLORES.cabecera).setFontColor('#ffffff').setFontWeight('bold').setWrap(true).setVerticalAlignment('middle');
  const titAnio = (fila, anio) => sh.getRange(fila, 1).setValue(anio).setFontSize(14).setFontWeight('bold').setHorizontalAlignment('left').setNumberFormat('0');

  // ---- Gastos fijos (tabla pequeña) ----
  const F = RES.fijos, fc = F.colBanco;
  sh.getRange(F.fila, fc, 1, 4).setValues([['Banco', 'Importe', 'Gastos', 'Importe']]);
  sh.getRange(F.fila, fc, 1, 4).setBackground(COLORES.cabecera).setFontColor('#ffffff').setFontWeight('bold');
  const n = Math.max(d.banco.length, d.gastos.length, 4);
  for (let i = 0; i < n; i++) {
    const b = d.banco[i] || ['', ''], g = d.gastos[i] || ['', ''];
    sh.getRange(F.fila + 1 + i, fc, 1, 4).setValues([[b[0], b[1], g[0], g[1]]]);
  }
  const filaTot = F.fila + 1 + n;
  sh.getRange(filaTot, fc, 1, 4).setValues([['Total', `=SUM(K${F.fila + 1}:K${filaTot - 1})`, 'Total', `=SUM(M${F.fila + 1}:M${filaTot - 1})`]]).setFontWeight('bold');
  sh.getRange(F.fila + 1, fc + 1, n + 1, 1).setNumberFormat(FMT.euro);
  sh.getRange(F.fila + 1, fc + 3, n + 1, 1).setNumberFormat(FMT.euro);
  const TOT_FIJOS = `($K$${filaTot}+$M$${filaTot})`, BANCO = `$K$${filaTot}`;

  // ---- Año en curso ----
  titAnio(RES.anioFila, d.anio);
  sh.getRange(RES.cabFila, 1, 1, 8).setValues([RES.cab]); estiloCab(RES.cabFila, 8);
  const fApert = `Trabajos!$${t['Fecha apertura']}:$${t['Fecha apertura']}`;
  for (let m = 1; m <= 12; m++) {
    const r = RES.ini + m - 1, s = d.y26[m - 1];
    const enVivo = m >= d.mesEnVivo;
    const rango = `${fApert},">="&DATE($A$${RES.anioFila},${m},1),${fApert},"<="&EOMONTH(DATE($A$${RES.anioFila},${m},1),0)`;
    const suma = h => `SUMIFS(Trabajos!$${t[h]}:$${t[h]},${rango})`;
    const rec = enVivo ? `=${suma('Recambios facturables')}` : (s[0] || 0);
    const ing = enVivo ? `=${suma('Factura')}` : (s[2] || 0);
    const mor = enVivo ? `=D${r}-SUMIFS(Trabajos!$${t['Factura']}:$${t['Factura']},${rango},Trabajos!$${t['Pagado']}:$${t['Pagado']},TRUE)` : (s[1] || 0);
    sh.getRange(r, 1, 1, 8).setValues([locFila_([MESES[m - 1], rec, mor, ing, s[3] || 0, `=D${r}-B${r}`, `=F${r}-E${r}-${TOT_FIJOS}`, `=${BANCO}+E${r}`])]);
  }
  sh.getRange(RES.total, 6, 1, 2).setValues([['Total año', `=SUM(G${RES.ini}:G${RES.ini + 11})`]]).setFontWeight('bold');
  sh.getRange(RES.ini, 2, 12, 7).setNumberFormat(FMT.euro);
  sh.getRange(RES.total, 7).setNumberFormat(FMT.euro);
  sh.getRange(RES.ini, 6, 12, 3).setBackground(COLORES.gris);
  sh.getRange(RES.ini, 2, 12, 1).setBackground('#ffffff');
  sh.getRange(RES.anioFila, 3).setValue(`Meses hasta ${MESES[d.mesEnVivo - 2] || '—'}: importados del Excel antiguo (valores fijos). Desde ${MESES[d.mesEnVivo - 1]}: se calculan desde Trabajos.`).setFontStyle('italic').setFontColor('#666666');

  // ---- 2025 ----
  const y25 = RES.y25;
  titAnio(y25.anio, 2025);
  sh.getRange(y25.cab, 1, 1, 7).setValues([RES.cab.slice(0, 7)]); estiloCab(y25.cab, 7);
  d.y25.forEach((s, i) => sh.getRange(y25.ini + i, 1, 1, 7).setValues([[MESES[i], s[0], s[1], s[2], s[3], s[4], s[5]]]));
  sh.getRange(y25.total, 6, 1, 2).setValues([['Total año', `=SUM(G${y25.ini}:G${y25.ini + 11})`]]).setFontWeight('bold');
  sh.getRange(y25.ini, 2, 12, 6).setNumberFormat(FMT.euro); sh.getRange(y25.total, 7).setNumberFormat(FMT.euro);

  // ---- 2024 ---- (Excel antiguo: Recambios 1 + Recambios 2, sólo Beneficio sin gastos fijos)
  const y24 = RES.y24;
  titAnio(y24.anio, 2024);
  sh.getRange(y24.cab, 1, 1, 7).setValues([['Mes', 'Recambios', 'Morosos', 'Ingreso', '', 'Beneficio', '']]); estiloCab(y24.cab, 7);
  d.y24.forEach((s, i) => sh.getRange(y24.ini + i, 1, 1, 7).setValues([[MESES[i], s[0], s[1], s[2], null, s[3], null]]));
  sh.getRange(y24.total, 5, 1, 2).setValues([['Total año', `=SUM(F${y24.ini}:F${y24.ini + 11})`]]).setFontWeight('bold');
  sh.getRange(y24.ini, 2, 12, 5).setNumberFormat(FMT.euro); sh.getRange(y24.total, 6).setNumberFormat(FMT.euro);
  sh.getRange(y24.total + 1, 1).setValue('2024: Recambios = suma de las dos columnas del Excel antiguo; Beneficio sin gastos fijos.').setFontStyle('italic').setFontColor('#666666');

  // ---- Gráficos ----
  const chart = (titulo, fCab, fIni, colBen, fila) => {
    const c = sh.newChart().asColumnChart()
      .addRange(sh.getRange(fCab, 1, 13, 1)).addRange(sh.getRange(fCab, colBen, 13, 1)).setNumHeaders(1)
      .setTitle(titulo).setPosition(fila, 10, 0, 0).setOption('legend', { position: 'none' })
      .setOption('width', 640).setOption('height', 290).setOption('colors', ['#1f3a5f']).build();
    sh.insertChart(c);
  };
  chart(`Beneficio mensual ${d.anio} (contando gastos fijos)`, RES.cabFila, RES.ini, 7, 13);
  chart('Beneficio mensual 2025 (contando gastos fijos)', y25.cab, y25.ini, 7, 29);
  chart('Beneficio mensual 2024 (sin gastos fijos)', y24.cab, y24.ini, 6, 45);

  [70, 120, 110, 120, 110, 120, 170, 140, 30, 140, 110, 140, 110].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.setFrozenRows(0);
}
