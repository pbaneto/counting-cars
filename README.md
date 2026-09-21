# counting-cars

Gestión de albaranes, trabajos, piezas y abonos de un taller mecánico sobre **Google Sheets + Apps Script**.
Los albaranes escaneados (PDF en una carpeta de Drive) y las facturas quincenales del proveedor los lee **Gemini**;
el resto son fórmulas de la hoja. Diseñado para que sea sencillo de entender y fácil de depurar.

## Cómo funciona (las reglas de negocio)

| Pestaña | Qué es |
|---|---|
| **Albaranes** | Una fila por albarán (escaneado o a mano). Suma su precio al *trabajo* sin pagar de esa matrícula, o abre uno nuevo. |
| **Trabajos** | Un trabajo = una reparación de un coche. Nº = último dígito + letras + contador (`7853KCC` → `3KCC-1`). Recambios y beneficio salen por fórmula. Puede haber varios trabajos sin pagar del mismo coche. |
| **Piezas** | Líneas de cada albarán. El check **Reembolso** resta esa pieza del precio facturable del albarán y del trabajo. |
| **Abonos** | Arriba, cuadre por quincena (albaranes RM ↔ reembolsos ↔ factura RM). Debajo, piezas reembolsadas y abonadas (`Abonada` / `Sin abonar` / `Sin solicitar`). |
| **Coches** | Matrícula, cliente y coche (se rellena a mano). |
| **Resumen** | Ingresos, recambios y beneficio por mes, gastos fijos y gráficos anuales. |
| *Facturas RM, Líneas RM* | Datos leídos de las facturas quincenales (una fila por línea). |
| *Config, Registro, Diagnóstico* | Ajustes, historial de incidencias y buscador de problemas. |

Reglas clave:

- **Precio facturable = precio con IVA − piezas con Reembolso ✓ de ese albarán.** No se guarda: es una fórmula. Un segundo escaneo con **R** manuscrita sólo marca la pieza.
- Albarán que ya existe **con R** → marca las piezas. **Sin R** → escaneo duplicado: se ignora.
- Si una fila escrita a mano (sin nº de albarán) coincide en matrícula e importe con un albarán escaneado, se **vincula** en vez de duplicarse.
- Un trabajo se imputa al mes de su **fecha de apertura**.
- **Cuadre de Abonos por quincena:** `Recambios RM − Abonado = Total factura` (integridad del escaneo) y `Solicitado = Abonado` (nada pendiente ni abonado sin pedir). Los abonos de RM indican el albarán original, así que se emparejan por **nº de albarán + referencia**.

## Depurar

- Cada aviso o error queda en **Registro** (función, archivo/fila, mensaje).
- Menú **Counting Cars ▸ Diagnóstico** lista todos los problemas con enlace a la fila.
- Columnas grises = fórmulas; **Avisos** explica qué falla en esa fila. **Reparar fórmulas y formato** las restaura.
- El código busca columnas por el **nombre de su cabecera**, nunca por letra.

## Estructura

```
src/            Código de Apps Script (se sube con clasp)
  Logic.js      Lógica pura (matrículas, trabajos, validación de IVA, abonos): probada en Node
  Prompts.js    Prompts y esquemas JSON de Gemini
  Config.js     Pestañas, columnas y valores por defecto
  Formulas.js   Fórmulas de las columnas automáticas
  Setup.js      Prepara la hoja (idempotente) · Triggers.js  Menú y edición manual
  Albaranes.js Facturas.js Abonos.js Resumen.js Diagnostico.js Seed.js ...
tests/          Pruebas (node --test): lógica + flujos completos con un simulador de Apps Script
tools/          probar-gemini.js (probar el prompt con un PDF real), generar-private.py
```

## Puesta en marcha

1. `npm i -g @google/clasp` · `clasp login`
2. Copia `.clasp.json.example` a `.clasp.json` y pon el `scriptId`.
3. Copia `private.example.js` a `src/private.js` (ID de la hoja, IDs de carpetas de Drive, datos iniciales). **No se sube a git.**
4. `clasp push -f`
5. En el editor de Apps Script ejecuta `setup` una vez (pide permisos). Después, en la hoja: **Counting Cars ▸ Configurar API key de Gemini**.

## Pruebas

```
npm test                                        # lógica + flujos simulados
GEMINI_API_KEY=... node tools/probar-gemini.js albaran ruta/albaran.pdf
```
