# SMTO Gastos — Convenciones del Proyecto

Claude Code MUST follow these rules in every modification. Never deviate.

## Renombrado de archivos (ZIP export)

The single source of truth is `buildFileName()` in `src/App.jsx`. NEVER duplicate or inline rename logic elsewhere.

**Format:** `Proveedor Folio Concepto-Segundo MM-DD-YY`

Rules:
- Proveedor → Title Case, max 40 chars, brackets stripped, no double dashes
- Folio → preserved EXACTLY as-is from CFDI (DB1616, CUUMXA110440, FAC102026491). Never upper/lower-case.
- Concepto → Title Case, first 2 words, joined by `-` (e.g. `Habitacion-Sencilla`)
- Fecha → `MM-DD-YY` derived from `fechaFac`
- Separator between blocks: single space (NOT underscore)
- Forbidden chars: `/ \ : * ? " < > | ( ) [ ] { }`

Examples:
- `Aerocomidas 66901114763782 Consumo 03-20-26.pdf`
- `Fideicomiso Irrevocable DB1616 CUUMXA110440 Habitacion-Sencilla 03-19-26.pdf`
- `Grupo Ferreteria Calzada FAC102026491 Hc168243-Urrea 03-20-26.pdf`

## fechaCobro

- New gasto (CFDI / OCR / manual) → `fechaCobro = fechaFac`
- After validarBanco match → `fechaCobro = dCSV` (bank charge date from CSV)

## Matching de banco (validarBanco)

- Pass 0: auth code exact match (OCR tickets only)
- Pass 1: `smartAmountMatch` with tolerance ±$0.01 ONLY
- No fuzzy passes. If exact match (subtotal + common tip %) not found, leave unmatched.
- Tip variants tested ONLY for gastos in food/restaurant/bar categories.
- Common tip percentages: 10, 12, 13, 15, 18, 20, 22, 25.

## Versionado

Increment minor version on every change. Update the version badge in the UI.

## Workflow

1. Read this file before any change
2. Make changes
3. Verify buildFileName isn't being shadowed
4. Commit with descriptive message
5. Push (Vercel auto-deploys)
