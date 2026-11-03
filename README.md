# counting-cars

Gestión de albaranes, trabajos, piezas y abonos de un taller mecánico sobre **Google Sheets + Apps Script**.
Los albaranes escaneados (PDF en una carpeta de Drive) y las facturas quincenales del proveedor los lee **Gemini**;
el resto son fórmulas de la hoja. Diseñado para que sea sencillo de entender y fácil de depurar.

## Cómo funciona (las reglas de negocio)

| Pestaña | Qué es |
|---|---|
| **Albaranes** | Una fila por albarán (escaneado o a mano). Suma su precio al *trabajo* sin pagar de esa matrícula, o abre uno nuevo. |
| **Trabajos** | Un trabajo = una reparación de un coche. Nº = último dígito + letras + contador (`7853KCC` → `3KCC-1`). Recambios y beneficio salen por fórmula. Puede haber varios trabajos sin pagar del mismo coche. |
| **Piezas** | Líneas de cada albarán, cada una con un **ID** fijo. El check **Reembolso** resta esa pieza del precio facturable del albarán y del trabajo. **Abonada a mano** (fecha): RM la ha abonado por otra vía. |
| **Abonos** | **Informe**: se rehace solo desde Piezas y Líneas RM, no se edita. Piezas pedidas y abonos de RM (`Abonada` / `Abonada (a mano)` / `Sin abonar` / `Sin solicitar`) desde la fecha `ABONOS_DESDE` de Config (lo pendiente sale siempre), lo más reciente arriba. El nº de albarán es un enlace a su PDF (también en Piezas). |
| **Cuadre RM** | Por quincena: albaranes RM ↔ abonado ↔ factura RM (el año, en B1); y lo pendiente de RM. |
| **Coches** | Matrícula, cliente y coche (se rellena a mano). |
| **Resumen** | Ingresos, recambios y beneficio por mes, gastos fijos y gráficos anuales. |
| *Facturas RM, Líneas RM* | Datos leídos de las facturas quincenales (una fila por línea) y de los albaranes de abono escaneados. En cada abono, **ID pieza** = la pieza abonada (se pone sola; se corrige a mano, `NO` = de ninguna). |
| *Config, Registro, Diagnóstico* | Ajustes, historial de incidencias y buscador de problemas. |

Reglas clave:

- **Precio - abonos (Albaranes) = precio con IVA − piezas con Reembolso ✓ de ese albarán.** No se guarda: es una fórmula. En Trabajos, su suma es «Precio - abono». Un segundo escaneo con **R** manuscrita sólo marca la pieza. Sólo cuenta una R sola junto a la línea: nombres, matrículas o números escritos a mano no son reembolsos.
- Albarán que ya existe **con R** → marca las piezas. **Sin R** → escaneo duplicado: se ignora. La pieza de la R se busca dentro del albarán: por referencia; si no la hay, por importe (entre varias con ese importe, la de referencia más parecida). Si empatan no se marca ninguna y se avisa.
- Albarán con **misma matrícula, fecha e importe** que otro y un solo dígito distinto en el nº → es el mismo con el nº mal leído: se trata como segundo escaneo del que ya estaba.
- Albarán de **dos hojas** (dos PDF): la 2ª completa la fila y «Ver PDF» pasa a tener un enlace por hoja (`Hoja 1 · Hoja 2`). El enlace del nº de albarán en Piezas y Abonos abre la 1ª.
- Si una fila escrita a mano (sin nº de albarán) coincide en matrícula e importe con un albarán escaneado, se **vincula** en vez de duplicarse.
- **Corregir una matrícula en Albaranes** pasa el albarán al trabajo abierto de esa matrícula o a uno nuevo. Si el trabajo viejo se queda vacío, se convierte en el nuevo (no hay trabajo abierto) o se borra; su factura y pagado pasan al trabajo abierto, salvo que este ya tenga factura: entonces se queda con un aviso. **Corregirla en Trabajos** cambia el trabajo entero: nº y matrícula de todos sus albaranes.
- Un trabajo se imputa a su columna **Mes** (se rellena con la fecha del primer albarán y se puede cambiar a mano).
- **Cuadre de Abonos por quincena:** `Recambios RM − Abonado = Total factura` (integridad del escaneo) y `Solicitado = Abonado` (nada pendiente ni abonado sin pedir). Los abonos de RM indican el albarán original, así que se emparejan una vez con su pieza por **nº de albarán + referencia** (o, si no, el importe de una pieza pedida) y queda guardado en Líneas RM. El total abonado sólo cuenta lo que trae RM, no «Abonada a mano».

## Depurar

- **Filtros**: un filtro de pestaña con criterios oculta filas (también las vacías de debajo) y Sheets no guarda las fechas ni las fórmulas que un script escribe en filas ocultas. Por eso el script quita los criterios antes de escribir y los vuelve a poner al terminar (`sinFiltro_`). Si alguna vez desaparecen datos, mirar lo primero si hay un filtro puesto.
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

## Versiones

`VERSION` (en `src/Config.js`) sigue el formato MAYOR.MENOR.PARCHE y la hoja guarda la suya. Al abrir la hoja o al lanzar una acción del menú, si no coinciden, la hoja se pone al día sola (`src/Version.js`):

- **MAYOR**: cambia dónde están los datos. Se añade su migración en `MIGRACIONES`; cuando ya ha corrido en la hoja real, se borra del código.
- **MENOR**: sólo diseño (formato, colores, fórmulas). Se reaplica el diseño.
- **PARCHE**: arreglo que no toca la hoja.

El diseño respeta lo cambiado a mano: anchos de columna (sólo se ponen al crear la hoja), filtros, filas fijas de más y reglas de color propias. Cada versión se etiqueta en git (`v1.0.0`).

