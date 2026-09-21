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

function enlacePdf_(archivo) { return loc_(`=HYPERLINK("https://drive.google.com/file/d/${archivo.getId()}/view","Ver PDF")`); }

function idDeEnlace_(texto) { const m = /\/d\/([\w-]+)/.exec(String(texto || '')); return m ? m[1] : ''; }
