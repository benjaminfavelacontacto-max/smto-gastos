# CLAUDE.md

This file provides guidance to Claude Code when working in this repository. Read it AND `CONVENTIONS.md` at the start of every session.

## What this is

SMTO Gastos — web app that automates Mexican expense reports for SMTO Engineering. Parses CFDI invoices (XML+PDF) and tickets (PDF/image via OCR), reconciles against bank statements (Clara MXN / Clara USA), exports premium Excel + ZIP with renamed files. UI is in Spanish; domain is Mexican tax/accounting.

Production: https://smto-app.vercel.app
Repo: github.com/benjaminfavelacontacto-max/smto-gastos
Local: /Users/benjaminfavela/Documents/SMTO/SMTO-Cuenta-de-Gastos/
Current version: v9.14

## Stack

- Frontend: React 18, Vite, JSZip, SheetJS (XLSX)
- Backend: Python serverless functions on Vercel (api/*.py)
- Python: pinned to 3.11 via .python-version
- Deps: openpyxl, Pillow, numpy, PyMuPDF
- No tests, linter, or TypeScript configured

## Commands

```bash
npm install        # frontend deps
npm run dev        # Vite at http://localhost:5173
npm run build      # production build → /dist
npm run preview    # serve /dist locally
```

Deployment: `git push` → Vercel auto-deploys in ~30s. No manual deploy commands.

## File layout

- `src/App.jsx` — main frontend logic (CFDI parser, matching, UI)
- `src/saldosCotejo.js` — cotejo de colaboradores especiales contra "Saldos AAAA": parser, T/C mensual del Saldos, match por factura (renglones con varias facturas, sufijo "-1", folios sin serie o truncados) y pagos en exhibiciones (prorrateo + COBRADO/FACTURADO)
- `src/presupuestoDinamico.js` — llena un mes del "SMTo - Presupuesto Dinámico" (hoja Real) desde la Hoja de trabajo + el Saldos: reglas tipo→renglón, semanas de la fila 12, fórmulas `=USD*<col>203+MXN`, escritura directa en el XML (conserva el logo), recálculo de todas las fórmulas (SUM) y cotejo de saldos al cierre
- `src/almacenLocal.js` — IndexedDB del navegador para el último Saldos y el último Presupuesto generado (no salen de la computadora)
- `src/components/` — PremiumModal, PremiumButton, and other small UI pieces
- `src/App.css` — full stylesheet
- `api/export-excel.py` — generates the premium .xlsx via openpyxl
- `api/ocr-ticket.py` — receipt OCR via Groq API (free). Rasterizes PDFs to image with PyMuPDF first since Groq vision only accepts images.
- `api/requirements.txt` — Python deps
- `api/TEMPLATE.xls`, `api/logo.png` — Excel template assets
- `public/logo.png` — frontend logo
- `CONVENTIONS.md` — immutable formatting + matching rules (MUST read)
- `.python-version` — pins Python 3.11
- `vercel.json` — Vercel config (no explicit runtime; auto-detected)

## Environment

- `SMTO_GROQ_API_KEY` — set in Vercel env vars; used by api/ocr-ticket.py (free Groq vision model `qwen/qwen3.6-27b`; the previous `meta-llama/llama-4-scout-17b-16e-instruct` was deprecated by Groq on 2026-07-17). Name is SMTO-specific to avoid colliding with a shared `GROQ_API_KEY` from another project.

## Features in production

1. CFDI XML parser: IVA, ISR, ISH, IEPS, retenciones, EDC combustible (GasNGo), Volare RFC exception
2. Bank reconciliation:
   - Clara MXN CSV: smartAmountMatch ±$0.01 (Pass 1) + delta-tip Pass 2 (cargo − total = propina, 5–35%, solo categorías de comida)
   - Clara USA CSV: match by authorization code
3. OCR for PDF/image tickets via Groq API (free) — only fires for PDFs without a matching XML. PDFs are rasterized to image server-side (PyMuPDF) before OCR since Groq vision only reads images. Image-OCR receipts are wrapped into single-page PDFs at ZIP export so they ride along with the standard buildFileName naming.
4. Multi-currency SIN columna extra (diseño aprobado por Victor Aceves
   2026-07-31; la variante con columna MONEDA fue RECHAZADA y revertida —
   commits `7b952f7`/`3f12c96`). El layout del Excel se queda en 22 columnas
   (B→W): la divisa se ve DENTRO de MONTO M.E. (col R, antes "MONTO USD") vía
   number_format con símbolo nativo (¥50 / €82.50 / RM 480.00 — el valor sigue
   siendo número, las plantillas de contabilidad no se rompen). El OCR detecta
   el ISO por símbolo + país/idioma (¥ se desambigua CNY vs JPY; $ USD/CAD/MXN)
   y `normaliza_moneda()` en export-excel.py blinda alias ("RM"→MYR). T/C: una
   celda editable POR divisa en la fila 7, ancladas de R7 hacia la izquierda —
   con UNA divisa el reporte es idéntico al clásico (label O7:Q7 + input R7);
   cada fila convierte con la celda de SU moneda. Tarjeta KPI R-T: con una
   divisa suma SUM(R) con su símbolo; con varias muestra "M. EXTRANJERA (MXN)"
   = SUMIF(S>0, O) (equivalente ya convertido; sumar montos nativos mezclados
   no significa nada). El import acepta ambos encabezados (MONTO M.E. / USD).
5. Propina (tip) split into its own sub-row in Excel
6. Collaborator selector: 50 people across 4 categories (Admin, Socio, Servicio, Ventas)
7. Expense types filtered by category (TIPOS_VENTAS vs TIPOS_NORMALES)
8. Drag & drop XMLs/PDFs/images onto the table
9. Import previous Excel to continue a report
10. Export ZIP: Excel + renamed XMLs/PDFs (see CONVENTIONS.md → buildFileName)
11. PremiumModal replaces all native alert/confirm. El modal de conciliación
    (`ConciliacionModal`) está escrito en lenguaje llano a propósito: la gente
    leía "Sin factura / Pendiente" como un adeudo. Trae banner "Esto no es un
    adeudo", 5 chips con hint, 4 pestañas (Conciliados / Cargos sin comprobante
    / No están en el estado / Revisar), nota explicativa por pestaña, leyenda de
    colores y aritmética "factura + propina = cargo" al expandir un match. Si se
    tocan esos textos, mantener el encuadre de "papeleo faltante", no de deuda.
12. Validation counter per checkbox
13. fechaCobro editable with date picker
14. Sortable, resizable columns; sticky PROVEEDOR
15. "Movimientos del banco" (solo Alejandro Olivar): propone desde el Saldos
    los egresos del mes de las cuentas de cheques que no tienen factura en el
    reporte (comisiones, pagos de crédito, fondo de ahorro, regalías,
    préstamos de nómina, no comprobados, IMSS, ISR, Telcel) con el proveedor, RFC,
    IVA y T/C como los captura contabilidad (`movimientosSinFactura` en
    src/saldosCotejo.js). El resto de egresos sin factura sale como opcional.
    Desde v9.09 se abre solo al aplicar "Cotejar con Saldos" (botón "Ver
    movimientos sin factura", reusa el Saldos ya leído) y cada movimiento se
    puede editar antes de agregarlo (proveedor, RFC, tipo, concepto, fecha,
    póliza, monto, IVA; `aplicarEdicionMovimiento` recalcula los montos).
16. El Saldos se puede cargar en los dos formatos: el "crudo" (Saldos 2026 (N))
    y el libro de Tesorería "Saldos 2026 Nuevo V1.x" (resumen arriba de cada
    pestaña, columnas extra Saldo/Categoría/Titular, "Sin clasificar" = tipo
    vacío). El V1.x no trae la tabla "Tipo Cambio": el T/C del mes se captura
    en el campo de los modales de Cotejar / Movimientos del banco. El V1.7
    (formato final) conserva fórmulas del crudo en Ingreso/Egreso
    ("=6000*1.16") sin valor calculado: el parser lee con `sheetStubs` y
    evalúa las fórmulas aritméticas (`valorCelda` en saldosCotejo.js).
17. Ingresos (v9.10, solo Alejandro Olivar): el modal de Movimientos del banco
    tiene pestañas Egresos / Ingresos. `ingresosDelMes` propone los depósitos
    de las cuentas de cheques (cobros a clientes con factura "Fact N0465" y PO
    del concepto, devoluciones, intereses, préstamos; nunca traspasos ni venta
    de bonos) con cliente/RFC de `CLIENTES_INGRESOS`. Viven en el state
    `ingresos` (NO en `lista`: no son gastos ni entran a KPIs) y se exportan a
    la hoja "Ingresos" del Excel con el layout de contabilidad (B..L, T/C en
    L1 ligado a 'Reporte SMTO'!R7). El import lee esa hoja de vuelta.
    Botón "Cuentas por Cobrar" (registro de facturas emitidas, pestaña
    "Facturas"): `parseCuentasPorCobrar` + `completarIngresoConCxC` ponen
    cliente, RFC, tipo (Type → Consultoría/Partes/Equipo/…), IVA (Tax ≠ 0) y,
    en cuentas USD, el total facturado cuando la diferencia es comisión del
    banco o la explica una nota de crédito (que se agrega como renglón
    negativo, como contabilidad). En cuentas en pesos se deja el depósito: el
    archivo guarda en USD facturas cobradas en pesos. Los números de factura
    se repiten entre años: se elige por fecha de pago.
    Intereses de Monex: una línea por cuenta (Monex USD / Monex MXN Cheques)
    al último día del mes = rendimiento de bonos (cada "Venta Bonos" menos
    su "Compra Bonos" anterior, en el mes de la venta) + los "Intereses" que
    ya trae la cuenta. Oct-2026: USD 18.68 y MXN 4,489.51 + 2,895.77 =
    7,385.28, igual que contabilidad.

18. Barra de acciones de los perfiles especiales (COLABORADORES_ESPECIALES):
    arriba la "Secuencia del reporte" numerada — 1 Cargar Carpeta, 2 Cotejar
    con Saldos y, sólo Alejandro, 3 Agregar Nómina, 4 Ingresos (sube Cuentas por Cobrar) y
    5 Presupuesto Dinámico —
    cada paso en verde con ✓ al terminarlo y el siguiente pendiente
    resaltado; abajo "Herramientas" (Manual, Importar reporte, Cargar Foto,
    Validar Banco, Movimientos del banco) y Exportar. Los demás perfiles
    conservan la barra de siempre. Con ingresos en el reporte, pestañas
    "Gastos | Ingresos" sobre la tabla (los ingresos ocupan el lugar de la
    tabla). La tabla conserva min-height 320px y, en pantallas bajas, la
    página (.app) se desplaza en vez de aplastarla.

19. Presupuesto Dinámico (v9.13, solo Alejandro Olivar, paso 5): sube la Hoja
    de trabajo del mes (la que exporta la app: hoja de egresos "Reporte SMTO"
    o "Egresos" + hoja "Ingresos") y descarga "SMTo - Presupuesto Dinámico
    <Mes> <Año>.xlsx" con el mes lleno por semana y cotejado con el Saldos.
    - El Saldos NO se vuelve a pedir: se guarda en IndexedDB del navegador al
      subirlo en Cotejar con Saldos, Agregar Nómina, Movimientos del banco o
      en el propio modal. El Presupuesto base se sube una vez y después cada
      presupuesto generado queda como base del siguiente mes. Se guarda por
      navegador y computadora (no en servidor); se pierde si se borran los
      datos del sitio, en incógnito o en Safari tras 7 días sin abrir la app.
      Al guardar se pide `navigator.storage.persist()` para que el navegador
      no lo borre por falta de espacio.
    - Si el archivo de Saldos guardado se modificó (File.lastModified) antes
      del cierre del mes de la Hoja, el modal avisa antes de generar: "Subir
      Saldos actualizado" (y genera solo con la Hoja ya elegida) o "Generar
      así" (el aviso queda en el resultado).
    - Fuentes: Ingresos (filas 15-44) por F. PAGO y Egresos (55-107) por
      F. COBRO de la Hoja; Consolidado bancario (121-229) del Saldos (categoría
      Traspaso → Traspasos±; créditos Monex/Konfio/BBVA con su convención);
      T/C de la Hoja en la fila 203. Lo que está en USD va como fórmula
      `=USD*<col>203(+MXN)`, como se capturaba a mano.
    - Se escribe directo en el XML de la hoja Real (logo, estilos, paneles e
      impresión intactos), se quita calcChain, fullCalcOnLoad=1 y se
      recalculan aquí todas las fórmulas de Real (sólo usa SUM y NOW) para que
      el archivo traiga los totales correctos en cualquier visor.
    - Cotejo en el modal: ingresos y egresos contra la Hoja (hasta $1 es
      redondeo de dólares) y saldo al cierre de cada cuenta contra el Saldos,
      al centavo (saldo base "Saldo base histórico" + movimientos ≤ fin de mes;
      NO usar "Apertura del ejercicio", que sigue el año de F11 de la pestaña).
    - Partidas cuyo tipo no tiene renglón → se listan y se puede generar sin
      ellas. Reglas y validación (junio y septiembre a mano; octubre 2026
      idéntico al validado) en ~/Documents/SMTO/Presupuesto Dinamico/.
    - Decisiones de Benjamín (10-nov-2026): comisiones ViTrox completas a la
      fila 15 (Ventas) salvo reparto manual; VIE Technologies con ViTrox;
      Equipo vendido a persona física → Venta Activos.
    - v9.14 — Hojas de varios meses y formatos viejos:
      - Lee tres formatos: el de la app (con fórmulas), el mismo con valores
        pegados, y el viejo de contabilidad (EGRESOS con FECHA DE COBRO /
        MONEDA EXT / TIPO CAMBIO; INGRESOS con FECHA DE PAGO / Cuenta / USD).
        Sin fórmulas, un egreso es en USD si su T/C es el "de lista" de su mes
        (el que más se repite, con ≤4 decimales: el del Fondo de Ahorro es un
        T/C real y queda en pesos) y MONEDA EXT × T/C = TOTAL; un ingreso es en
        USD si entró a una cuenta USD. El monto en USD es el del renglón (cada
        mes de una Hoja vieja trae su propio T/C).
      - Si la Hoja trae varios meses, el modal pide cuáles llenar (marca los
        que tienen ≥5% de los renglones). Para meses que el presupuesto base ya
        tenía (≤ "lleno hasta") se conserva su T/C y el resultado compara
        "antes vs ahora" por mes + "Diferencias (CSV)" por celda. La base sólo
        se reemplaza sola con un mes nuevo; si no, botón "Guardar como base".
      - Reglas nuevas sacadas de la captura a mano: Gasolina Serv → Gasolina
        Servicio (72); Renta Auto → Viáticos Servicio; Seguro, Comercio Ext,
        ISR IVA, Pago Préstamo → su renglón; Préstamo (egreso) → Préstamos
        (105); ScienScope también como "SCIENTIFIC TECH…"; un pago de un
        principal con tipo Comisiones cuyo concepto no es de comisiones (ni
        nota de crédito) va a su Consultoría.
      - Prueba (10-nov-2026) con las Hojas viejas ene-abr y may-oct contra el
        presupuesto capturado a mano: totales de ingresos y egresos iguales en
        los 10 meses (centavos). Lo distinto es reparto entre renglones por
        criterio de captura (ViTrox Ventas/Servicio; Soporte vs Distribución
        caso por caso) y 2 montos: mayo Aduana +7,660.02 y julio Comisiones de
        banco −4.64 en lo capturado. Reporte en ~/Downloads/Prueba Presupuesto Dinámico/.

## Cómo agregar un colaborador

Todo en `src/App.jsx`, en DOS lugares (ejemplo: Alberth Cruz, Servicio, terminación Clara 2900):

1. `POLIZAS_CLARA` — folio Clara de 4 dígitos (columna PÓLIZA del Excel): `'Alberth Cruz': '2900'`. Sin tarjeta Clara queda 'N/A', no hace falta entrada.
2. Lista de colaboradores — `{ categoria: 'Servicio', nombre: 'Alberth Cruz' }`, dentro del bloque de su categoría (Admin, Socio, Servicio, Ventas).

"Soporte" = categoría `Servicio`. Los socios que se tratan como soporte van aparte en `SOCIOS_COMO_SOPORTE`. Si es empleado con nómina, su RFC va en `EMPLEADOS_RFC`. Verificar con `npm run build`.

## Excel output design

- Colors: SMTO black #050505 + green #59D39B, EXCEL_GREEN borders #00B050
- 5 KPI cards with green borders
- Embedded logo
- Zebra-striped table, propina as sub-row beneath parent
- Native datetime objects (no date warnings)
- SUM formulas in totals row
- Columns (B→W, 22 en total — NO agregar columnas: contabilidad tiene plantillas
  que leen por letra y la variante de 23 columnas fue rechazada): RFC,
  PROVEEDOR, TIPO, PÓLIZA, FACTURA, F. FACTURA, F. COBRO, CONCEPTO, IMPORTE,
  IVA, ISR, ISH/IEPS, RETENCIÓN, TOTAL, FORMA PAGO, BANCO, MONTO M.E. (R,
  símbolo nativo vía number_format), T/C, USUARIO, COBRADO, FACTURADO,
  DIFERENCIA

## Architecture notes

The app is fully client-side except for two Python serverless functions (Excel export, OCR). Files never leave the user's machine except when OCR is explicitly invoked.

**State model:** a single `lista` array of "gasto" objects. No normalization, reducer, or context. Mutations go through `setLista(prev => prev.map(...))`. IDs are random base36 from `genId()`. Manual rows use `uuid: 'MANUAL'`.

**CFDI parsing is namespace-agnostic:** walks all elements and matches on `localName` (`comprobante`, `emisor`, `impuestos`, `timbrefiscaldigital`, `concepto`). Attribute reads tolerate both casings via `ga(el, 'Pascal', 'lower')`. PDFs match XMLs by identical base filename OR UUID substring — keep both paths.

**Date handling:** stored internally as `YYYY-MM-DD`, displayed/edited as `MM-DD-YYYY`, exported as `DD/MM/YYYY`. `parseDateRobusto` accepts all three plus `YY` shorthand. Any new date field round-trips through these helpers.

**fechaCobro behavior is asymmetric:**
- Before reconciliation: equals fechaFac
- After reconciliation: equals dCSV (bank charge date)

## Cost control (Groq API — free tier)

OCR runs on Groq's free vision model, so there is no per-call dollar cost. The
gating below still applies to avoid unnecessary calls and respect Groq free-tier
rate limits.

- XML alone → parsed locally, NO API call
- XML + PDF → PDF linked to XML, NO OCR
- PDF without matching XML → user confirmation modal BEFORE calling OCR
- Image (jpg/png/heic/webp/bmp/gif) → compressed client-side to max 2000px width @ 85% quality, then OCR with user confirmation
- Image OCR receipts → original (compressed) image stored on the gasto as `imageDataURL`, then converted to a single-page PDF at ZIP-export time via jsPDF (loaded from CDN in `index.html`). The PDF filename follows the `buildFileName` convention.

**Batch OCR + Groq TPM limit:** the free tier of `qwen/qwen3.6-27b` caps at **8000 tokens/min (TPM)**, and each OCR request is ~6000 tokens, so only ~1 receipt/min fits — a 14-image batch takes ~10 min. All OCR loops run **sequentially** (never `Promise.all` for vision calls) and share:
- a progress indicator (`ocrProgress` state → both overlays show "N de M");
- client-side pacing (`OCR_PACING_MS`, ~5s between requests) to avoid firing the batch as one burst;
- the server's `Retry-After`-aware 429 retry in `ocr-ticket.py` (up to 5×), which is the precise rate-limit backoff — pacing is only a smoother. Failed files are reported to the user, not dropped silently. For real throughput, Groq's paid Dev Tier is the fix.

## ⚠️ Producción vs entorno de prueba — REGLA OBLIGATORIA

Hay un entorno de pruebas (rediseño, ramas tipo `persistencia-local`, PR #7)
que vive en ramas separadas y **NUNCA se debe mezclar con producción** hasta
que Benjamín lo apruebe explícitamente. Incidente del 2026-10-30: una rama de
producción se creó desde el checkout local, que estaba en la rama de prueba,
y el PR #8 subió el rediseño completo a producción (se revirtió en #10).

**Cada cambio a producción sigue estos pasos, sin excepción:**

1. **Nunca partir del checkout local.** La carpeta del proyecto puede estar en
   una rama de prueba o tener cambios de prueba sin commitear (CLAUDE.md,
   CONVENTIONS.md, CONTEXTO_SMTO_GASTOS.md, etc.). Hacer `git fetch` y crear
   un worktree aparte desde `origin/main`:
   `git worktree add -b <rama> <carpeta-temporal> origin/main`
2. Editar y commitear **sólo dentro de ese worktree**. Nunca `git add` de
   archivos del checkout local que no hiciste tú en esta tarea.
3. **Antes de push/PR, verificar el contenido:**
   - `git log --oneline origin/main..HEAD` → sólo tus commits.
   - `git diff --stat origin/main...HEAD` → sólo los archivos que pretendías
     tocar. Si aparece cualquier otro archivo (p. ej. `src/borrador.js`,
     cambios masivos en `App.jsx`/`App.css`), DETENERSE y avisar.
4. PR a `main`, squash merge.
5. Verificar producción: https://smto-app.vercel.app debe decir
   "Reporte de Gastos SMTO" + la versión nueva (no el título del rediseño).
6. Borrar el worktree temporal. No tocar ni stashear los cambios de prueba
   del checkout local.

Nunca fusionar, cherry-pickear ni rebasar desde ramas de prueba hacia `main`.

## Deploy workflow

1. Crear worktree limpio desde `origin/main` (ver regla de arriba)
2. Claude Code edits files there
3. `git commit -m "describe change"`
4. `git push` + PR a `main` + squash merge
5. Vercel auto-deploys (~30s)
6. Verify at production URL

## Versioning

- Increment minor version on every meaningful change
- Update the version badge in the UI header
- El número de versión es el de PRODUCCIÓN. El entorno de prueba no consume
  ni sube versiones de producción.
- Current: v9.14

## Things to be careful about

- `buildFileName()` in App.jsx is the SINGLE source of truth for renaming. Never duplicate or inline it elsewhere.
- Folio must be preserved EXACTLY as it comes from CFDI (no upper/lower-case): DB1616, CUUMXA110440, FAC102026491.
- Never call OCR for PDFs that already have an XML match.
- Tip variants only apply to food/restaurant/bar categories.
- Pass order in validarBanco is load-bearing — see CONVENTIONS.md.
- Preserve all Spanish strings in user-facing text.

## Browser support

- `webkitdirectory` attribute set imperatively via ref
- `navigator.clipboard.writeText` requires HTTPS or localhost
- `URL.createObjectURL` for opening PDFs in new tabs
- No polyfills. Target: Chrome/Edge/Firefox/Safari 14+
