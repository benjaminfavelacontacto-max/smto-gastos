# CLAUDE.md

This file provides guidance to Claude Code when working in this repository. Read it AND `CONVENTIONS.md` at the start of every session.

## What this is

SMTO Gastos — web app that automates Mexican expense reports for SMTO Engineering. Parses CFDI invoices (XML+PDF) and tickets (PDF/image via OCR), reconciles against bank statements (Clara MXN / Clara USA), exports premium Excel + ZIP with renamed files. UI is in Spanish; domain is Mexican tax/accounting.

Production: https://smto-app.vercel.app
Repo: github.com/benjaminfavelacontacto-max/smto-gastos
Local: /Users/benjaminfavela/Documents/SMTO/smto-app/
Current version: v8.52

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

- `SMTO_GROQ_API_KEY` — set in Vercel env vars; used by api/ocr-ticket.py (free Groq vision model `meta-llama/llama-4-scout-17b-16e-instruct`). Name is SMTO-specific to avoid colliding with a shared `GROQ_API_KEY` from another project.

## Features in production

1. CFDI XML parser: IVA, ISR, ISH, IEPS, retenciones, EDC combustible (GasNGo), Volare RFC exception
2. Bank reconciliation:
   - Clara MXN CSV: smartAmountMatch ±$0.01 (Pass 1) + delta-tip Pass 2 (cargo − total = propina, 5–35%, solo categorías de comida)
   - Clara USA CSV: match by authorization code
3. OCR for PDF/image tickets via Groq API (free) — only fires for PDFs without a matching XML. PDFs are rasterized to image server-side (PyMuPDF) before OCR since Groq vision only reads images. Image-OCR receipts are wrapped into single-page PDFs at ZIP export so they ride along with the standard buildFileName naming.
4. Multi-currency: USD, EUR, JPY, etc. with per-line exchange rate
5. Propina (tip) split into its own sub-row in Excel
6. Collaborator selector: 50 people across 4 categories (Admin, Socio, Servicio, Ventas)
7. Expense types filtered by category (TIPOS_VENTAS vs TIPOS_NORMALES)
8. Drag & drop XMLs/PDFs/images onto the table
9. Import previous Excel to continue a report
10. Export ZIP: Excel + renamed XMLs/PDFs (see CONVENTIONS.md → buildFileName)
11. PremiumModal replaces all native alert/confirm
12. Validation counter per checkbox
13. fechaCobro editable with date picker
14. Sortable, resizable columns; sticky PROVEEDOR

## Excel output design

- Colors: SMTO black #050505 + green #59D39B, EXCEL_GREEN borders #00B050
- 5 KPI cards with green borders
- Embedded logo
- Zebra-striped table, propina as sub-row beneath parent
- Native datetime objects (no date warnings)
- SUM formulas in totals row
- Columns: RFC, PROVEEDOR, TIPO, FACTURA, F.FACTURA, F.COBRO, CONCEPTO, IMPORTE, IVA, RETENCIÓN, TOTAL, FORMA PAGO, MONTO EXT, T/C

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

## Deploy workflow

1. Claude Code edits files
2. `git commit -m "describe change"`
3. `git push`
4. Vercel auto-deploys (~30s)
5. Verify at production URL

## Versioning

- Increment minor version on every meaningful change
- Update the version badge in the UI header
- Current: v8.52

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
