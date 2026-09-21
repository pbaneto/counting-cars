/**
 * Fórmulas de las columnas automáticas (grises). Todas por fila: funcionan con ordenar, filtrar e insertar filas.
 * Se construyen con las letras REALES de las cabeceras (letras_), así que insertar una columna no las rompe al reparar.
 * Named ranges usados: IVA, DIAS_AVISO, TOL_CUADRE (apuntan a la pestaña Config).
 */

const FORMULAS = {
  'Albaranes': {
    'Quincena': r => { const a = letras_(HOJA.ALB); return `=IF(${a['Fecha albarán']}${r}="","",IF(DAY(${a['Fecha albarán']}${r})<=15,1,2))`; },
    'Precio facturable': r => {
      const a = letras_(HOJA.ALB), p = letras_(HOJA.PIEZAS);
      const suma = `SUMIFS(Piezas!$${p['Precio descontado con IVA']}:$${p['Precio descontado con IVA']},Piezas!$${p['Nº albarán']}:$${p['Nº albarán']},${a['Nº albarán']}${r},Piezas!$${p['Reembolso']}:$${p['Reembolso']},TRUE)`;
      return `=IF(${a['Precio con IVA']}${r}="","",${a['Precio con IVA']}${r}-IF(${a['Nº albarán']}${r}="",0,${suma}))`;
    },
    'Coche': r => {
      const a = letras_(HOJA.ALB), c = letras_(HOJA.COCHES);
      return `=IF(${a['Matrícula']}${r}="","",IFERROR(VLOOKUP(${a['Matrícula']}${r},Coches!$A:$C,${indice_(HOJA.COCHES, 'Coche')},FALSE),"⚠ Matrícula no está en Coches"))`;
    },
    'Cliente': r => {
      const a = letras_(HOJA.ALB);
      return `=IF(${a['Matrícula']}${r}="","",IFERROR(VLOOKUP(${a['Matrícula']}${r},Coches!$A:$C,${indice_(HOJA.COCHES, 'Cliente')},FALSE),""))`;
    },
    'Avisos': r => {
      const a = letras_(HOJA.ALB), t = letras_(HOJA.TRAB), c = letras_(HOJA.COCHES);
      const E = `${a['Nº albarán']}${r}`, F = `${a['Nº trabajo']}${r}`, G = `${a['Matrícula']}${r}`, H = `${a['Precio con IVA']}${r}`, A = `${a['Fecha escaneo']}${r}`;
      const fila = `MATCH(${F},Trabajos!$${t['Nº trabajo']}:$${t['Nº trabajo']},0)`;
      const idx = h => `INDEX(Trabajos!$${t[h]}:$${t[h]},${fila})`;
      const dias = `(${A}-${idx('Fecha apertura')})`;
      return `=IF(AND(${G}="",${H}="",${E}=""),"",` +
        `IF(${H}="","⚠ Falta el precio",` +
        `IF(${G}="","⚠ Falta la matrícula",` +
        `IF(COUNTIF(Coches!$${c['Matrícula']}:$${c['Matrícula']},${G})=0,"⚠ La matrícula no está en Coches",` +
        `IF(${F}="","⚠ Sin nº de trabajo",` +
        `IF(COUNTIF(Trabajos!$${t['Nº trabajo']}:$${t['Nº trabajo']},${F})=0,"⚠ El trabajo "&${F}&" no existe",` +
        `IF(${idx('Matrícula')}<>${G},"⚠ El trabajo "&${F}&" es de otra matrícula",` +
        `IF(AND(${E}<>"",COUNTIF($${a['Nº albarán']}:$${a['Nº albarán']},${E})>1),"⚠ Nº de albarán duplicado",` +
        `IF(IFERROR(AND(${A}<>"",${idx('Pagado')}<>TRUE,${dias}>DIAS_AVISO),FALSE),"⚠ Trabajo "&${F}&" abierto hace "&${dias}&" días: ¿es un trabajo nuevo? Elige NUEVO en Nº trabajo",""))))))))) `.trim();
    },
  },

  'Trabajos': {
    'Coche': r => { const t = letras_(HOJA.TRAB); return `=IF(${t['Matrícula']}${r}="","",IFERROR(VLOOKUP(${t['Matrícula']}${r},Coches!$A:$C,${indice_(HOJA.COCHES, 'Coche')},FALSE),"⚠ Matrícula no está en Coches"))`; },
    'Cliente': r => { const t = letras_(HOJA.TRAB); return `=IF(${t['Matrícula']}${r}="","",IFERROR(VLOOKUP(${t['Matrícula']}${r},Coches!$A:$C,${indice_(HOJA.COCHES, 'Cliente')},FALSE),""))`; },
    'Recambios': r => {
      const t = letras_(HOJA.TRAB), a = letras_(HOJA.ALB);
      return `=IF(${t['Nº trabajo']}${r}="","",SUMIFS(Albaranes!$${a['Precio con IVA']}:$${a['Precio con IVA']},Albaranes!$${a['Nº trabajo']}:$${a['Nº trabajo']},${t['Nº trabajo']}${r}))`;
    },
    'Recambios facturables': r => {
      const t = letras_(HOJA.TRAB), a = letras_(HOJA.ALB);
      return `=IF(${t['Nº trabajo']}${r}="","",SUMIFS(Albaranes!$${a['Precio facturable']}:$${a['Precio facturable']},Albaranes!$${a['Nº trabajo']}:$${a['Nº trabajo']},${t['Nº trabajo']}${r}))`;
    },
    'Beneficio': r => { const t = letras_(HOJA.TRAB); return `=IF(OR(${t['Nº trabajo']}${r}="",${t['Factura']}${r}=""),"",${t['Factura']}${r}-${t['Recambios facturables']}${r})`; },
    'Avisos': r => {
      const t = letras_(HOJA.TRAB), a = letras_(HOJA.ALB), c = letras_(HOJA.COCHES);
      const N = `${t['Nº trabajo']}${r}`, B = `${t['Fecha apertura']}${r}`, C = `${t['Matrícula']}${r}`, G = `${t['Recambios facturables']}${r}`, H = `${t['Factura']}${r}`, J = `${t['Pagado']}${r}`;
      return `=IF(${N}="","",` +
        `IF(${C}="","⚠ Falta la matrícula",` +
        `IF(COUNTIF(Coches!$${c['Matrícula']}:$${c['Matrícula']},${C})=0,"⚠ La matrícula no está en Coches",` +
        `IF(COUNTIF($${t['Nº trabajo']}:$${t['Nº trabajo']},${N})>1,"⚠ Nº de trabajo duplicado",` +
        `IF(COUNTIF(Albaranes!$${a['Nº trabajo']}:$${a['Nº trabajo']},${N})=0,"ℹ Sin albaranes",` +
        `IF(IFERROR(AND(${J}<>TRUE,${B}<>"",TODAY()-${B}>DIAS_AVISO),FALSE),"ℹ Sin pagar desde hace "&(TODAY()-${B})&" días",` +
        `IF(AND(${H}<>"",${H}<${G}),"ℹ Factura menor que los recambios (pérdida)","")))))))`;
    },
  },

  'Piezas': {
    'Matrícula': r => {
      const p = letras_(HOJA.PIEZAS), a = letras_(HOJA.ALB);
      return `=IF(${p['Nº albarán']}${r}="","",IFERROR(INDEX(Albaranes!$${a['Matrícula']}:$${a['Matrícula']},MATCH(${p['Nº albarán']}${r},Albaranes!$${a['Nº albarán']}:$${a['Nº albarán']},0)),""))`;
    },
    'Precio descontado con IVA': r => { const p = letras_(HOJA.PIEZAS); return `=IF(${p['Precio descontado sin IVA']}${r}="","",ROUND(${p['Precio descontado sin IVA']}${r}*(1+IVA),2))`; },
    'Proveedor': r => {
      const p = letras_(HOJA.PIEZAS), a = letras_(HOJA.ALB);
      return `=IF(${p['Nº albarán']}${r}="","",IFERROR(INDEX(Albaranes!$${a['Proveedor']}:$${a['Proveedor']},MATCH(${p['Nº albarán']}${r},Albaranes!$${a['Nº albarán']}:$${a['Nº albarán']},0)),""))`;
    },
    'Avisos': r => {
      const p = letras_(HOJA.PIEZAS), a = letras_(HOJA.ALB);
      const A = `${p['Reembolso']}${r}`, C = `${p['Nº albarán']}${r}`, J = `${p['Precio descontado sin IVA']}${r}`;
      return `=IF(AND(${A}<>TRUE,${C}=""),"",` +
        `IF(${C}="","⚠ Falta el nº de albarán (obligatorio para el reembolso)",` +
        `IF(COUNTIF(Albaranes!$${a['Nº albarán']}:$${a['Nº albarán']},${C})=0,"⚠ El albarán "&${C}&" no está en Albaranes",` +
        `IF(AND(${A}=TRUE,${J}=""),"⚠ Falta el precio de la pieza",""))))`;
    },
  },

  'Líneas RM': {
    'Conciliación': r => {
      const l = letras_(HOJA.LINEAS), a = letras_(HOJA.ALB);
      const F = l['Nº factura'], B = l['Nº albarán'], L = l['Importe sin IVA'];
      return `=IF(${l['Tipo']}${r}<>"Compra","",` +
        `IF(COUNTIF(Albaranes!$${a['Nº albarán']}:$${a['Nº albarán']},${B}${r})=0,"⚠ Albarán no escaneado",` +
        `IF(ABS(ROUND(SUMIFS($${L}:$${L},$${F}:$${F},${F}${r},$${B}:$${B},${B}${r})*(1+IVA),2)-SUMIFS(Albaranes!$${a['Precio con IVA']}:$${a['Precio con IVA']},Albaranes!$${a['Nº albarán']}:$${a['Nº albarán']},${B}${r}))>0.05,"⚠ Importe distinto","OK")))`;
    },
  },
};

/** Posición (1, 2, 3…) de una cabecera dentro de su pestaña, para VLOOKUP. */
function indice_(hoja, cabecera) {
  const l = letras_(hoja)[cabecera];
  let n = 0;
  for (let i = 0; i < l.length; i++) n = n * 26 + (l.charCodeAt(i) - 64);
  return n;
}
