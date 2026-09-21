/** Llamadas a Gemini: en paralelo (fetchAll), con reintentos sólo de lo que falla y esquema JSON obligatorio. */

function claveGemini_() {
  const k = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!k) throw new Error('Falta la API key de Gemini. Usa Counting Cars ▸ Configurar API key de Gemini.');
  return k;
}

function construirPeticion_(archivo, esAlbaran, clave) {
  const blob = archivo.getBlob();
  const body = {
    contents: [{ parts: [
      { inline_data: { mime_type: blob.getContentType(), data: Utilities.base64Encode(blob.getBytes()) } },
      { text: esAlbaran ? PROMPT_ALBARAN : PROMPT_FACTURA },
    ] }],
    generationConfig: { temperature: 0, responseMimeType: 'application/json', responseSchema: esAlbaran ? SCHEMA_ALBARAN : SCHEMA_FACTURA, maxOutputTokens: 32768 },
  };
  return {
    url: `https://generativelanguage.googleapis.com/v1beta/models/${cfg_('MODELO_GEMINI')}:generateContent`,
    method: 'post', contentType: 'application/json', headers: { 'x-goog-api-key': clave },
    payload: JSON.stringify(body), muteHttpExceptions: true,
  };
}

/** {ok:true, doc} | {ok:false, transitorio:boolean, error}. Nunca lanza. */
function interpretarRespuesta_(resp, esAlbaran) {
  const code = resp.getResponseCode(), texto = resp.getContentText();
  if (code === 429 || code >= 500) return { ok: false, transitorio: true, error: `Gemini HTTP ${code}: ${texto.slice(0, 200)}` };
  // 400/401/403/404: problema de configuración (clave, modelo, permisos), no del documento: hay que parar, no descartar el PDF.
  if (code !== 200) return { ok: false, transitorio: false, config: true, error: `Gemini HTTP ${code} (revisa la API key y el modelo en Config): ${texto.slice(0, 300)}` };
  try {
    const j = JSON.parse(texto), cand = j.candidates && j.candidates[0];
    if (!cand || !cand.content) return { ok: false, transitorio: true, error: `Gemini no devolvió contenido (${JSON.stringify(j.promptFeedback || cand && cand.finishReason || '')})` };
    const raw = JSON.parse(cand.content.parts.map(p => p.text || '').join(''));
    return { ok: true, doc: esAlbaran ? normalizarAlbaran(raw) : normalizarFactura(raw), truncado: cand.finishReason === 'MAX_TOKENS' };
  } catch (e) { return { ok: false, transitorio: true, error: 'Respuesta de Gemini no válida: ' + e.message }; }
}

/**
 * archivos: [DriveApp.File]. Devuelve un resultado por archivo, en el mismo orden.
 * Lotes de PARALELISMO peticiones simultáneas; hasta 3 intentos para los errores transitorios (429/5xx/timeout).
 */
function leerConGemini_(archivos, esAlbaran) {
  const clave = claveGemini_(), n = Math.max(1, Math.floor(cfgNum_('PARALELISMO')));
  const resultados = new Array(archivos.length);
  for (let ini = 0; ini < archivos.length; ini += n) {
    let pendientes = [];
    for (let i = ini; i < Math.min(ini + n, archivos.length); i++) pendientes.push(i);
    for (let intento = 1; intento <= 3 && pendientes.length; intento++) {
      if (intento > 1) Utilities.sleep(3000 * intento);
      let respuestas;
      try { respuestas = UrlFetchApp.fetchAll(pendientes.map(i => construirPeticion_(archivos[i], esAlbaran, clave))); }
      catch (e) { respuestas = null; pendientes.forEach(i => { resultados[i] = { ok: false, transitorio: true, error: 'Fallo de red: ' + e.message }; }); }
      if (respuestas) pendientes.forEach((i, k) => { resultados[i] = interpretarRespuesta_(respuestas[k], esAlbaran); });
      pendientes = pendientes.filter(i => !resultados[i].ok && resultados[i].transitorio);
    }
  }
  return resultados;
}

function configurarApiKey() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('API key de Gemini', 'Pega la clave. Se guarda sólo en las propiedades privadas del script (no en la hoja ni en el repositorio).', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const k = r.getResponseText().trim();
  if (k.length < 20) { ui.alert('La clave parece demasiado corta. No se ha guardado.'); return; }
  PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', k);
  ui.alert('API key guardada.');
}
