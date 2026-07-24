# SMTO Gastos — Convenciones del Proyecto

Claude Code MUST follow these rules in every modification. Never deviate.

## Renombrado de archivos (ZIP export)

The single source of truth is `buildFileName()` in `src/App.jsx`. NEVER duplicate or inline rename logic elsewhere.

**Format:** `Proveedor Folio Tipo MM-DD-YY`

- Proveedor → Title Case, full name (not truncated), brackets stripped, no double dashes
- Folio → preserved EXACTLY as-is from CFDI (DB1616, CUUMXA110440, FAC102026491). Never upper/lower-case.
- Tipo → gasto type (Vuelo, Hotel, Transporte, Herramienta, Consumo, etc.) from `g.tipo`, Title Case. Falls back to concepto if tipo missing.
- Fecha → `MM-DD-YY` derived from `fechaFac`
- Separator between blocks: single space
- Forbidden chars: `/ \ : * ? " < > | ( ) [ ] { }`

Examples:
- `Aerocomidas 66901114763782 Consumo 03-20-26.pdf`
- `Fideicomiso Irrevocable DB1616 CUUMXA110440 Hotel 03-19-26.pdf`
- `Grupo Ferreteria Calzada FAC102026491 Herramienta 03-20-26.pdf`

## Excel export (api/export-excel.py)

- **PROVEEDOR siempre en MAYÚSCULAS** en el Excel exportado (`.upper()` en la
  celda), sin importar la fuente (XML emisor, OCR, manual) ni el usuario de la
  plantilla. Esto NO afecta el Title Case del nombre de archivo del ZIP (ver
  arriba), que es una regla distinta.
- **COBRADO nunca se inventa después de una conciliación.** Si la sesión tuvo
  corrida de banco (algún gasto con `montoCobrado > 0` o `hizoMatch`), las
  filas Clara sin match muestran "SIN CONCILIAR" en COBRADO (ámbar) y
  DIFERENCIA vacía — jamás el estimado factura+propina, que pintaba un falso
  $0.00. El estimado solo aplica en exports SIN validación bancaria.

## fechaCobro

- New gasto (CFDI / OCR / manual) → `fechaCobro = fechaFac`
- After validarBanco match → `fechaCobro = dCSV` (bank charge date from CSV)

## Matching de banco (validarBanco)

- Pass 0: auth code exact match (OCR tickets only)
- Pass 1: `smartAmountMatch` with tolerance ±$0.01 ONLY. OCR tickets probe BOTH
  the OCR subtotal AND the OCR total (`totalOCR` = montoExtranjero ‖ totalCFDI) —
  el banco cobra el total con impuestos, así que un recibo con IVA y sin propina
  (hotel, comida) o un Uber sin subtotal cuadra por el total. El subtotal+propina
  sigue cubriendo comidas con propina.
- Pass 2: delta-tip match, ONLY for tip-eligible (factura, cargo) pairs — if the
  bank charge exceeds the invoice total by 5%–35%, the delta IS the propina
  (exact to the cent, taken from the real charge). Covers atypical tips (16%)
  and hand-rounded amounts that the Pass 1 ladder misses. MUST run after
  passes 0–1 so exact matches claim their rows first. Same-currency gate: una
  factura en divisa (GBP/EUR/USD) solo se compara contra el "Monto original" del
  cargo en ESA misma moneda; una factura MXN, contra el "Monto en MXN". Nunca
  cruzar divisas (evita propinas fantasma subtotal-en-divisa vs cargo-en-MXN).
- Pass 3: near-exact match — charge vs (total + detected propina) differ by
  ≤1% AND ≤$20 (hand-keyed terminal amounts, minor adjustments). Binds with
  confidence 70 (< 80) so it lands in the modal's "Revisión" tab, and the
  Excel DIFERENCIA column surfaces the real gap. MUST run last.
- No fuzzy amount passes beyond these. If no pass binds, leave unmatched.
- Tip variants (Pass 1 ladder AND Pass 2 delta) tested ONLY for gastos in
  food/restaurant/bar categories (Clara "Categoría de Compra" first, keyword
  sniff fallback). A hotel/gas/uber charge can never bind via phantom tip.
- Pass 1 common tip percentages: 10, 12, 13, 15, 16, 18, 20, 22, 25.

## Cost control (Anthropic API)

- XML alone → parsed locally, NO API call
- XML + PDF → PDF linked to XML, NO OCR, NO cost
- PDF without matching XML → user confirmation modal BEFORE calling OCR
- Image (jpg/png/heic/webp) → compressed client-side to max 2000px width @ 85% quality, then OCR with user confirmation

## Versionado

Increment minor version on every change. Update the version badge in the UI.

## Workflow

1. Read this file before any change
2. Make changes
3. Verify buildFileName isn't being shadowed
4. Commit with descriptive message
5. Push (Vercel auto-deploys)
