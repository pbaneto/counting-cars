/** Carpetas y archivos de Drive. Los IDs de carpeta viven en la pestaña Config, no en el código. */

function carpeta_(clave) {
  const id = String(cfg_(clave) || '').trim();
  if (!id) throw new Error(`Falta el ID de carpeta "${clave}" en la pestaña Config.`);
  try { return DriveApp.getFolderById(id); }
  catch (e) { throw new Error(`No puedo abrir la carpeta "${clave}" (ID ${id}). ¿Tienes acceso con esta cuenta? ${e.message}`); }
}

/** PDFs e imágenes de una carpeta, por orden de nombre (el escáner los nombra con fecha y hora). */
function listarArchivos_(clave, max) {
  const it = carpeta_(clave).getFiles(), out = [];
  while (it.hasNext()) {
    const f = it.next(), mt = f.getMimeType();
    if (mt === 'application/pdf' || /^image\/(jpeg|png)$/.test(mt)) out.push(f);
  }
  out.sort((a, b) => a.getName().localeCompare(b.getName()));
  return max ? out.slice(0, max) : out;
}

function moverArchivo_(archivo, clave) { archivo.moveTo(carpeta_(clave)); }

/** Carpeta "Errores" (junto a Entrada). Se crea la primera vez que hace falta. */
function moverAErrores_(archivo) {
  let id = String(cfg_('CARPETA_ERRORES') || '').trim();
  if (!id) {
    const padres = carpeta_('CARPETA_ENTRADA').getParents();
    const padre = padres.hasNext() ? padres.next() : DriveApp.getRootFolder();
    id = padre.createFolder('Errores').getId();
    guardarCfg_('CARPETA_ERRORES', id);
  }
  archivo.moveTo(DriveApp.getFolderById(id));
}

function enlacePdf_(archivo) { return enlacePdfId_(archivo.getId()); }
function urlPdf_(id) { return `https://drive.google.com/file/d/${id}/view`; }
function enlacePdfId_(id) { return loc_(`=HYPERLINK("${urlPdf_(id)}","Ver PDF")`); }

function idDeEnlace_(texto) { const m = /\/d\/([\w-]+)/.exec(String(texto || '')); return m ? m[1] : ''; }

/** Celda con un enlace por hoja: "Hoja 1 · Hoja 2". */
function textoHojasPdf_(ids) {
  const trozos = ids.map((_, i) => `Hoja ${i + 1}`), sep = ' · ', b = SpreadsheetApp.newRichTextValue().setText(trozos.join(sep));
  let pos = 0;
  ids.forEach((id, i) => { b.setLinkUrl(pos, pos + trozos[i].length, urlPdf_(id)); pos += trozos[i].length + sep.length; });
  return b.build();
}

/** Pone en "Ver PDF" los PDF del albarán: un HYPERLINK si es una hoja, un enlace por hoja si son más. */
function ponerPdfs_(tabla, fila, ids) {
  if (ids.length < 2) { actualizarFila_(tabla, fila, { 'Ver PDF': enlacePdfId_(ids[0]) }); return; }
  sinFiltro_(tabla.sh);
  tabla.sh.getRange(fila, tabla.map['Ver PDF']).clearContent().setRichTextValue(textoHojasPdf_(ids));
}

/** {nº de fila: [IDs de sus PDF]} de la columna "Ver PDF": sale de la fórmula o, si la celda tiene varias hojas, de sus enlaces. */
function pdfsPorFila_(tabla) {
  const desde = tabla.fc + 1, n = tabla.libre - desde, out = {};
  if (n < 1) return out;
  const rango = tabla.sh.getRange(desde, tabla.map['Ver PDF'], n, 1), formulas = rango.getFormulas(), ricos = rango.getRichTextValues();
  formulas.forEach((r, i) => {
    const una = idDeEnlace_(r[0]), rico = ricos[i][0];
    const ids = una ? [una] : rico ? rico.getRuns().map(x => idDeEnlace_(x.getLinkUrl())).filter(Boolean) : [];
    if (ids.length) out[desde + i] = ids;
  });
  return out;
}
