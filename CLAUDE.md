# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

```bash
npm install        # install deps (React 18, Vite 5, JSZip — that's it)
npm run dev        # Vite dev server at http://localhost:5173
npm run build      # production build → /dist
npm run preview    # serve the built /dist locally
```

No tests, linter, formatter, or TypeScript are configured. Deployment is via Vercel (a `.vercel/` link exists); Vercel auto-detects Vite — no extra config.

## Architecture

This is a **single-page, fully client-side** React app for reconciling Mexican **CFDI invoice XMLs** against a **bank statement CSV**. No backend, no API, no persistence — everything runs in the browser. Files never leave the user's machine.

**The entire app is one file: `src/App.jsx` (~580 lines).** `src/main.jsx` just mounts it; `src/App.css` is the entire stylesheet. When making changes, expect to edit App.jsx; resist the urge to split it up unless a feature genuinely requires it.

The UI is in **Spanish** and the domain is **Mexican tax/accounting** (CFDI, RFC, propina, IVA, retenciones). Preserve Spanish strings in user-facing text.

### Data flow (one pass through App.jsx tells the story)

1. **`cargar`** — user picks a folder via `<input webkitdirectory>`. Files are split into `.xml` (CFDI invoices) and `.pdf` (visual copies). Each XML is parsed with `parseCFDI`.

2. **`parseCFDI`** — CFDI XMLs use SAT namespaces (`cfdi:`, `tfd:`). Parsing is **namespace-agnostic**: it walks all elements and matches on `localName` (`comprobante`, `emisor`, `impuestos`, `timbrefiscaldigital`, `concepto`). Attribute reads use `ga(el, 'Pascal', 'lower')` to tolerate both casings. PDFs are matched to XMLs by **identical base filename** OR **UUID substring** — keep both paths when changing matcher logic.

3. **`clasificarGasto`** — keyword classifier that returns one of `Vuelo | Hotel | Transporte | Herramienta | Consumo`. Pure string matching against `proveedor + concepto`; add new keywords here, don't refactor into a config object unless asked.

4. **`validarBanco`** — two-pass reconciliation against a bank CSV (auto-detects `,` / `;` / `\t` separator):
   - **Pass 1**: exact total match (±$5) within ±30 days.
   - **Pass 2**: propina (tip) detection — bank charge is greater than invoice total by up to 25%; the difference is auto-filled as `montoPropina`.
   - Matches set `hizoMatch=true`, which colors the row blue. The two passes run **per bank line**, not globally — order matters; don't reorder them.

5. **`update`** has implicit coupling: editing `totalCFDI`, `propinaPorcentaje`, or `montoPropina` recomputes the other propina fields. If you add a new derived field, add its recompute branch here.

6. **`copiar`** writes TSV to clipboard (for Excel paste). **`exportar`** uses JSZip to bundle a CSV report plus the original XML/PDF files **renamed** to `Proveedor-Factura-Concepto-MM-DD-YY`. The CSV starts with a BOM (`﻿`) so Excel reads UTF-8 correctly — don't drop it.

### Date handling — read before touching

Dates are stored internally as `YYYY-MM-DD` but **displayed and edited as `MM-DD-YYYY`** in the table (see `GastoRow`'s `dateDisplay` / `onDateChange`). Exports use `DD/MM/YYYY` (Mexican). `parseDateRobusto` accepts all three plus `YY` shorthand. Any new date field should round-trip through these helpers.

### State model

A single `lista` array of "gasto" objects holds everything; there's no normalization, no reducer, no context. Every mutation goes through `setLista(prev => prev.map(...))`. IDs are random base36 from `genId()`. Rows added via `agregarManual` use `uuid: 'MANUAL'` and `xmlFile: null` — `exportar` skips file-renaming for these, which is intentional.

### Browser API dependencies

- `webkitdirectory` attribute — set imperatively via `ref` because React doesn't recognize the lowercase variant.
- `navigator.clipboard.writeText` — requires HTTPS or localhost.
- `URL.createObjectURL` for opening PDFs in a new tab.

These are not polyfilled. The README lists Chrome/Edge/Firefox/Safari 14+ as the support matrix.
