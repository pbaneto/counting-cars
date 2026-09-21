#!/usr/bin/env node
/**
 * Prueba local del prompt con un PDF real (no toca Sheets ni Drive).
 *   GEMINI_API_KEY=$(cat ~/.counting-cars.key) node tools/probar-gemini.js albaran ruta/al.pdf [modelo]
 *   GEMINI_API_KEY=... node tools/probar-gemini.js factura ruta/factura.pdf
 * La clave sólo se lee del entorno; nunca se guarda en el repo.
 */
const fs = require('fs');
const P = require('../src/Prompts.js');
const L = require('../src/Logic.js');

async function main() {
  const [modo, ruta, modelo = 'gemini-3.5-flash-lite'] = process.argv.slice(2);
  const key = process.env.GEMINI_API_KEY;
  if (!key || !['albaran', 'factura'].includes(modo) || !ruta) {
    console.error('Uso: GEMINI_API_KEY=... node tools/probar-gemini.js <albaran|factura> <archivo.pdf> [modelo]');
    process.exit(2);
  }
  const esAlb = modo === 'albaran';
  const body = {
    contents: [{ parts: [
      { inline_data: { mime_type: 'application/pdf', data: fs.readFileSync(ruta).toString('base64') } },
      { text: esAlb ? P.PROMPT_ALBARAN : P.PROMPT_FACTURA },
    ] }],
    generationConfig: { temperature: 0, responseMimeType: 'application/json', responseSchema: esAlb ? P.SCHEMA_ALBARAN : P.SCHEMA_FACTURA, maxOutputTokens: 32768 },
  };
  const t0 = Date.now();
  const r = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${modelo}:generateContent`, {
    method: 'POST', headers: { 'Content-Type': 'application/json', 'x-goog-api-key': key }, body: JSON.stringify(body),
  });
  const j = await r.json();
  if (!r.ok) { console.error('HTTP', r.status, JSON.stringify(j).slice(0, 600)); process.exit(1); }
  const texto = j.candidates[0].content.parts.map(p => p.text || '').join('');
  const doc = esAlb ? P.normalizarAlbaran(JSON.parse(texto)) : P.normalizarFactura(JSON.parse(texto));
  console.log(JSON.stringify(doc, null, 1));
  console.log('--- validación:', JSON.stringify(esAlb ? L.validarAlbaran(doc) : L.validarFactura(doc)));
  console.log(`--- ${((Date.now() - t0) / 1000).toFixed(1)} s, tokens:`, JSON.stringify(j.usageMetadata));
}
main().catch(e => { console.error(e); process.exit(1); });
