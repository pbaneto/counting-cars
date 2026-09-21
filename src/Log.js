/** Registro de incidencias: cada aviso/error queda en la pestaña "Registro" con función y referencia (pestaña!fila o archivo). */

let _buffer = [];

function log_(nivel, funcion, referencia, mensaje) {
  _buffer.push([new Date(), nivel, funcion, referencia || '', String(mensaje).slice(0, 1500)]);
  Logger.log(`[${nivel}] ${funcion} ${referencia || ''} ${mensaje}`);
}

function volcarLog_() {
  if (!_buffer.length) return;
  const pend = _buffer;
  _buffer = [];
  try {
    const sh = hoja_(HOJA.REG);
    const ini = Math.max(sh.getLastRow(), 1) + 1;
    sh.getRange(ini, 1, pend.length, 5).setValues(pend);
  } catch (e) { Logger.log('No se pudo escribir en Registro: ' + e.message); }
}

/** Avisa al usuario con un cuadro si hay interfaz; si no (editor, trigger), sólo al log. */
function avisar_(msg, titulo) {
  try { SpreadsheetApp.getUi().alert(titulo || 'Counting Cars', msg, SpreadsheetApp.getUi().ButtonSet.OK); }
  catch (e) { Logger.log(`[AVISO] ${msg}`); }
}

function toast_(msg, titulo, seg) {
  try { ss_().toast(msg, titulo || 'Counting Cars', seg || 6); } catch (e) { Logger.log(`[TOAST] ${msg}`); }
}

/** Envuelve un punto de entrada: captura cualquier error, lo anota en Registro y se lo cuenta al usuario. */
function ejecutar_(nombre, fn, silencioso) {
  try { return fn(); }
  catch (e) {
    log_('ERROR', nombre, '', (e && e.stack) || e);
    const msg = e && e.message ? e.message : String(e);
    if (silencioso) toast_(`${msg} (detalle en la pestaña Registro)`, '⚠ Error', 10);
    else avisar_(`Ha fallado "${nombre}":\n${msg}\n\nDetalle en la pestaña Registro.`, 'Error');
  } finally { volcarLog_(); }
}

function conBloqueo_(segundos, fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(segundos * 1000)) throw new Error('Hay otro proceso en marcha. Espera un minuto y vuelve a intentarlo.');
  try { return fn(); } finally { lock.releaseLock(); }
}
