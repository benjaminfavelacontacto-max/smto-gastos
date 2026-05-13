import { useState, useRef, useMemo, useEffect } from 'react'
import JSZip from 'jszip'
import { motion, AnimatePresence } from 'framer-motion'
import { CheckCircle2, X, CreditCard, Target, Sparkles, AlertTriangle, FileText, FileSpreadsheet, Package, Check } from 'lucide-react'

const TIPOS_GASTO = [
  'Aduana','Avión','Avión Ventas','Casetas','Casetas Ventas','Celular','COGS',
  'Consumo','Consumo Viáticos','Devolución','Envios','Estacionamiento',
  'Estacionamiento Ventas','Gasolina','Gasolina Ventas','Gasolina Ventas Viáticos',
  'Gasolina Viáticos','Gastos Rep','Gastos Rep Ventas','Gastos Rep Viáticos',
  'Herramientas','Hotel','Hotel Ventas','IT & SW','Manto Auto','Marketing',
  'No Comprobado','Papelería','PC','Rechazada','Renta Auto','Renta Oficina',
  'Taxi','Taxi Ventas','Traspaso','Uniformes'
]

const autoDetectTipo = (descripcion) => {
  const d = (descripcion || '').toLowerCase()
  if (d.includes('hotel') || d.includes('hospedaje')) return 'Hotel'
  if (d.includes('vuelo') || d.includes('avion') || d.includes('tarifa aerea')) return 'Avión'
  if (d.includes('taxi') || d.includes('uber') || d.includes('didi')) return 'Taxi'
  if (d.includes('gasolina') || d.includes('combustible')) return 'Gasolina'
  if (d.includes('caseta') || d.includes('autopista') || d.includes('peaje')) return 'Casetas'
  if (d.includes('renta') && (d.includes('auto') || d.includes('veh'))) return 'Renta Auto'
  if (d.includes('renta') && (d.includes('oficina') || d.includes('local'))) return 'Renta Oficina'
  if (d.includes('estacionamiento') || d.includes('parking')) return 'Estacionamiento'
  if (d.includes('celular') || d.includes('telefon') || d.includes('telcel')) return 'Celular'
  if (d.includes('herramienta') || d.includes('ferreteria')) return 'Herramientas'
  if (d.includes('software') || d.includes('licencia') || d.includes('suscripci')) return 'IT & SW'
  if (d.includes('marketing') || d.includes('publicidad')) return 'Marketing'
  if (d.includes('uniforme')) return 'Uniformes'
  if (d.includes('envio') || d.includes('paquete') || d.includes('flete') || d.includes('dhl') || d.includes('fedex')) return 'Envios'
  if (d.includes('aduana')) return 'Aduana'
  if (d.includes('mantenimiento')) return 'Manto Auto'
  return 'Consumo'
}

/* Single source of truth for table columns.
   `getValue` overrides simple `g[key]` lookup (used by Total Final). */
const COLUMNS = [
  { key: 'check',             label: '',             width: 52,  sortable: true,  type: 'string' },
  { key: 'status',            label: 'Estado',       width: 110, sortable: true,  type: 'string' },
  { key: 'fechaFac',          label: 'Fecha Factura',width: 115, sortable: true,  type: 'date'   },
  { key: 'fechaCobro',        label: 'Fecha Cobro',  width: 120, sortable: true,  type: 'date'   },
  { key: 'noFactura',         label: 'Factura',      width: 120, sortable: true,  type: 'string' },
  { key: 'proveedor',         label: 'Proveedor',    width: 260, sortable: true,  type: 'string' },
  { key: 'concepto',          label: 'Concepto',     width: 140, sortable: true,  type: 'string' },
  { key: 'tipo',              label: 'Tipo',         width: 100, sortable: true,  type: 'string' },
  { key: 'importe',           label: 'Subtotal',     width: 120, sortable: true,  type: 'number' },
  { key: 'iva',               label: 'IVA',          width: 110, sortable: true,  type: 'number' },
  { key: 'isrTrasladado',     label: 'ISR/ISH/IEPS', width: 135, sortable: true,  type: 'number' },
  { key: 'retencionISR',      label: 'Ret. ISR',     width: 110, sortable: true,  type: 'number' },
  { key: 'retencionIVA',      label: 'Ret. IVA',     width: 110, sortable: true,  type: 'number' },
  { key: 'retenciones',       label: 'Reten.',       width: 110, sortable: true,  type: 'number' },
  { key: 'totalCFDI',         label: 'Total Fac.',   width: 125, sortable: true,  type: 'number' },
  { key: 'formaPago',         label: 'Forma de Pago', width: 160, sortable: true,  type: 'string' },
  { key: 'propinaPorcentaje', label: 'Prop. %',      width: 95,  sortable: true,  type: 'number' },
  { key: 'montoPropina',      label: 'Prop. $',      width: 105, sortable: true,  type: 'number' },
  { key: 'totalFinal',        label: 'Total Final',  width: 130, sortable: true,  type: 'number',
    getValue: g => g.totalCFDI + g.montoPropina },
]

/* ═══════════════════════════════════════════════════
   UTILIDADES
═══════════════════════════════════════════════════ */

const genId = () => Math.random().toString(36).slice(2, 11)

/* Display dates as MM-DD-YY app-wide. Internal storage stays YYYY-MM-DD
   (HTML5 date input requirement). parseDateDisplay reverses for storage. */
const formatDateDisplay = (dateStr) => {
  if (!dateStr) return ''
  let d, m, y
  if (dateStr.includes('-') && dateStr.length === 10) {
    // YYYY-MM-DD
    const [yyyy, mm, dd] = dateStr.split('-')
    d = dd; m = mm; y = yyyy.slice(-2)
  } else if (dateStr.includes('/')) {
    // DD/MM/YYYY
    const parts = dateStr.split('/')
    d = parts[0]; m = parts[1]; y = (parts[2] || '').slice(-2)
  } else if (dateStr.includes('-') && dateStr.length <= 8) {
    // Already MM-DD-YY
    return dateStr
  } else {
    return dateStr
  }
  return `${m.padStart(2, '0')}-${d.padStart(2, '0')}-${y}`
}

const parseDateDisplay = (s) => {
  if (!s) return ''
  const parts = s.split('-')
  if (parts.length !== 3) return s
  const [mm, dd, y] = parts
  if (y.length === 2) return `20${y}-${mm.padStart(2, '0')}-${dd.padStart(2, '0')}`
  if (y.length === 4) return `${y}-${mm.padStart(2, '0')}-${dd.padStart(2, '0')}`
  return s
}

// Human-readable file size for the export-success card.
const formatBytes = (b) => {
  if (b < 1024) return `${b} B`
  if (b < 1024 * 1024) return `${(b / 1024).toFixed(1)} KB`
  return `${(b / 1024 / 1024).toFixed(2)} MB`
}

function parseDateRobusto(text) {
  if (!text) return null
  const chars = text.trim().replace(/[^0-9/\-]/g, '')
  const sep = chars.includes('/') ? '/' : '-'
  const parts = chars.split(sep)
  if (parts.length < 3) return null
  const [p0, p1, p2raw] = parts
  const p2 = p2raw.slice(0, 4)
  let y, m, d
  if      (p0.length === 4)  { y = p0; m = p1; d = p2 }
  else if (p2.length === 4)  { y = p2; m = p1; d = p0 }
  else if (p2.length === 2)  { y = '20' + p2; m = p1; d = p0 }
  else return null
  const date = new Date(`${y}-${m.padStart(2, '0')}-${d.padStart(2, '0')}T12:00:00`)
  return isNaN(date) ? null : date
}

function parseCSVLine(line, sep) {
  const result = []; let cur = ''; let inQ = false
  for (const ch of line) {
    if (ch === '"') inQ = !inQ
    else if (ch === sep && !inQ) { result.push(cur); cur = '' }
    else cur += ch
  }
  result.push(cur)
  return result
}

function parseCFDI(xmlText, xmlFile, pdfFiles) {
  // RFCs whose <Retencion Impuesto="001"> is actually an ISR trasladado
  // (their PDFs label it as a line-item tax, not a withholding).
  const RFC_ISR_COMO_TRASLADO = ['AUEJ040528DNA']

  let doc
  try { doc = new DOMParser().parseFromString(xmlText, 'text/xml') } catch { return null }

  const ga = (el, ...names) => {
    if (!el) return null
    for (const n of names) { const v = el.getAttribute(n); if (v !== null) return v }
    return null
  }

  let comp = null, emisor = null, timbre = null, conceptoEl = null
  for (const el of doc.querySelectorAll('*')) {
    const ln = el.localName.toLowerCase()
    if (ln === 'comprobante'        && !comp)       comp = el
    if (ln === 'emisor'             && !emisor)     emisor = el
    if (ln === 'timbrefiscaldigital'&& !timbre)     timbre = el
    if (ln === 'concepto'           && !conceptoEl) conceptoEl = el
  }

  if (!comp || !emisor) return null
  const rfc = ga(emisor, 'Rfc', 'RFC') || ''
  if (!rfc) return null

  // Concepto — first line of the first Concepto's Descripcion, capped at 80
  // chars. Replaces the old keyword-based categorizer (Vuelo/Hotel/etc.) so
  // each row shows the actual product description from the XML.
  const descripcionRaw = conceptoEl ? (ga(conceptoEl, 'Descripcion', 'descripcion') || '') : ''
  // Title-case so ALL-CAPS or all-lowercase descriptions render naturally.
  const toTitleCase = str => str.toLowerCase().replace(/(?:^|\s|\/|-)\S/g, c => c.toUpperCase())
  const descripcionFirstLine = descripcionRaw.split(/[\n\r]/)[0].trim()
  const concepto = toTitleCase(descripcionFirstLine.slice(0, 80) || 'Consumo')

  const proveedor = ga(emisor, 'Nombre', 'NOMBRE') || 'Proveedor'
  const uuid = ga(timbre, 'UUID', 'uuid') || ''

  // Taxes — CFDI 4.0 has TWO <Impuestos> elements: one per <Concepto> (line-item
  // totals) and one at the root (overall totals). We MUST read the root one only.
  // querySelectorAll('*') would match the first per-Concepto node and inflate sums.
  let rootImp = null
  for (const node of comp.childNodes) {
    if (node.nodeType === 1 && node.localName && node.localName.toLowerCase() === 'impuestos') {
      rootImp = node
      break
    }
  }

  // Helpers — case-insensitive, namespace-agnostic, no-throw on missing nodes.
  const findChild = (parent, name) => {
    if (!parent) return null
    for (const c of parent.childNodes) {
      if (c.nodeType === 1 && c.localName && c.localName.toLowerCase() === name) return c
    }
    return null
  }
  // CFDI 4.0 names the tax-code attr "Impuesto" on Traslado/Retencion; older
  // formats sometimes use "TipoImpuesto". Try Impuesto first.
  const readTipo    = el => (ga(el, 'Impuesto', 'TipoImpuesto', 'tipoImpuesto') || '').trim()
  const readImporte = el => parseFloat(ga(el, 'Importe', 'importe') || '0') || 0
  // `match` can be a TipoImpuesto string (exact match) or a predicate
  // function (e.g. t => t !== '002' to capture every non-IVA traslado).
  const sumByTipo = (container, childTag, match) => {
    if (!container) return 0
    const test = typeof match === 'function' ? match : t => t === match
    let total = 0
    for (const el of container.childNodes) {
      if (el.nodeType !== 1 || !el.localName || el.localName.toLowerCase() !== childTag) continue
      if (test(readTipo(el))) total += readImporte(el)
    }
    return total
  }

  const trasladosBox   = findChild(rootImp, 'traslados')
  const retencionesBox = findChild(rootImp, 'retenciones')

  // Per-invoice Retencion log for in-browser verification.
  const retEls = []
  if (retencionesBox) {
    for (const r of retencionesBox.childNodes) {
      if (r.nodeType === 1 && r.localName && r.localName.toLowerCase() === 'retencion') retEls.push(r)
    }
  }
  console.log('[parseCFDI]', xmlFile.name, '— Retencion elements found:', retEls.length)
  for (const r of retEls) {
    console.log('[parseCFDI]   Retencion tipo=', readTipo(r), 'importe=', readImporte(r))
  }

  // The parent container determines the bucket — Traslado nodes can only feed
  // iva/isrTrasladado, Retencion nodes can only feed retencionISR/retencionIVA.
  let   iva           = sumByTipo(trasladosBox,   'traslado',  '002')
  let   isrTrasladado = sumByTipo(trasladosBox,   'traslado',  '001')  // ISR from regular Traslados
  let   retencionISR  = sumByTipo(retencionesBox, 'retencion', '001')
  const retencionIVA  = sumByTipo(retencionesBox, 'retencion', '002')
  let   retenciones   = retencionISR + retencionIVA

  // Per-RFC override: some providers' <Retencion Impuesto="001"> is actually
  // a trasladado ISR (the invoice line-item tax), not a withholding.
  if (RFC_ISR_COMO_TRASLADO.includes(rfc)) {
    isrTrasladado += retencionISR
    retencionISR   = 0
    retenciones    = retencionIVA
  }

  // Local-tax complement (ISH/IEPS/etc) — flat scan, since nested scoped
  // getElementsByTagName misses implocal:-namespaced descendants in some DOM
  // implementations. Element name is the only thing we trust. ALL TrasladosLocales
  // feed isrTrasladado regardless of ImpLocTrasladado code; ALL RetencionesLocales
  // add to the retenciones total.
  for (const el of doc.querySelectorAll('*')) {
    if (!el.localName) continue
    const ln = el.localName.toLowerCase()
    if      (ln === 'trasladoslocales')  isrTrasladado += parseFloat(ga(el, 'Importe', 'importe') || '0') || 0
    else if (ln === 'retencioneslocales') retenciones  += parseFloat(ga(el, 'Importe', 'importe') || '0') || 0
  }

  // Fields the EDC override may rewrite — pull them out before the override block.
  // importe is the NET subtotal (after Descuento). Some CFDIs (e.g. Volaris) declare
  // a Descuento on <Comprobante>; subtracting it makes the invariant
  // importe + iva ≈ totalCFDI hold on the row.
  const descuento    = parseFloat(ga(comp, 'Descuento', 'descuento') || '0') || 0
  let importe        = (parseFloat(ga(comp, 'SubTotal', 'subtotal') || '0') || 0) - descuento
  let totalCFDI      = parseFloat(ga(comp, 'Total',    'total')    || '0') || 0
  let conceptoClasif = concepto   // raw first-line Descripcion (EDC may override)

  // EDC (Estado de Cuenta de Combustible — GasNGo & similar). The regular
  // <cfdi:Comprobante> has placeholder amounts (SubTotal=1.00, Total=0.00);
  // real figures live on <Dispersion> inside <cfdi:Addenda>. IEPS feeds the
  // ISR/ISH/IEPS bucket per spec.
  let isEDC = false, dispersion = null
  for (const el of doc.querySelectorAll('*')) {
    if (!el.localName) continue
    const ln = el.localName.toLowerCase()
    if      (ln === 'estadodecuentacombustible')   isEDC = true
    else if (ln === 'dispersion' && !dispersion)   dispersion = el
  }
  if (isEDC && dispersion) {
    importe         = parseFloat(ga(dispersion, 'GralImporte',  'gralImporte')  || '0') || 0
    iva             = parseFloat(ga(dispersion, 'GralImpuesto', 'gralImpuesto') || '0') || 0
    isrTrasladado  += parseFloat(ga(dispersion, 'GralIEPS',     'gralIEPS')     || '0') || 0
    totalCFDI       = parseFloat(ga(dispersion, 'GralTotal',    'gralTotal')    || '0') || 0
    conceptoClasif  = 'Combustible'
  }

  // Buscar PDF asociado — 3 estrategias en orden:
  //   1. Coincidencia exacta de nombre base
  //   2. UUID dentro del nombre del PDF
  //   3. Coincidencia "fuzzy": tras normalizar (lowercase + solo alfanuméricos),
  //      uno contiene al otro, o comparten los primeros 15 caracteres.
  const base    = xmlFile.name.replace(/\.xml$/i, '').toLowerCase()
  const norm    = s => s.toLowerCase().replace(/[^a-z0-9]/g, '')
  const xmlNorm = norm(base)
  const pdfFile = pdfFiles.find(f => {
    const pdfBase = f.name.replace(/\.pdf$/i, '').toLowerCase()
    if (pdfBase === base) return true
    if (uuid && f.name.toUpperCase().includes(uuid.toUpperCase())) return true
    const pdfNorm = norm(pdfBase)
    if (xmlNorm.length >= 6 && pdfNorm.length >= 6) {
      if (xmlNorm.includes(pdfNorm) || pdfNorm.includes(xmlNorm)) return true
    }
    if (xmlNorm.length >= 15 && pdfNorm.length >= 15 && xmlNorm.slice(0, 15) === pdfNorm.slice(0, 15)) return true
    return false
  }) || null
  console.log('[parseCFDI]', xmlFile.name, '— PDF match:', pdfFile ? pdfFile.name : 'NONE')

  const fechaFac = (ga(comp, 'Fecha', 'fecha') || '').slice(0, 10)
  return {
    id: genId(),
    rfc,
    proveedor,
    noFactura: (ga(comp, 'Serie', 'serie') || '') + (ga(comp, 'Folio', 'folio') || 'SN'),
    fechaFac,
    concepto:   conceptoClasif,
    tipo: autoDetectTipo(descripcionFirstLine),
    importe,
    iva,
    isrTrasladado,
    retencionISR,
    retencionIVA,
    retenciones,
    totalCFDI,
    propinaPorcentaje: 0,
    montoPropina: 0,
    fechaCobro: fechaFac,
    formaPago:  ga(comp, 'FormaPago') || '04',
    uuid,
    tienePDF: !!pdfFile,
    pdfFile,
    xmlFile,
    hizoMatch: false,
    checkManual: false,
  }
}

/* ═══════════════════════════════════════════════════
   COMPONENTE: BOTÓN PREMIUM
═══════════════════════════════════════════════════ */

function PremiumButton({ title, icon, variant = 'primary', isDisabled = false, onClick }) {
  return (
    <button
      className={`btn-premium ${variant}${isDisabled ? ' disabled' : ''}`}
      onClick={isDisabled ? undefined : onClick}
      disabled={isDisabled}
    >
      <span className="btn-icon">{icon}</span>
      {title}
    </button>
  )
}

/* ═══════════════════════════════════════════════════
   COMPONENTE: FILA DE LA TABLA
═══════════════════════════════════════════════════ */

function GastoRow({ g, upd, openPDF }) {
  // Display ↔ storage: app-wide formatDateDisplay/parseDateDisplay handle
  // the MM-DD-YY ↔ YYYY-MM-DD round-trip.
  const dateDisplay  = formatDateDisplay(g.fechaFac)
  const onDateChange = v => upd('fechaFac', parseDateDisplay(v))

  // Per-row toggle for the Fecha Cobro cell: span (MM-DD-YY) when blurred,
  // native date picker (YYYY-MM-DD) when focused.
  const [editingCobro, setEditingCobro] = useState(false)
  // type="date" needs YYYY-MM-DD on its value attr. Bank-matched DD/MM/YYYY
  // from legacy in-memory data still parses correctly via this guard.
  const cobroIso = (() => {
    if (!g.fechaCobro) return ''
    if (g.fechaCobro.includes('/')) {
      const [d, m, y] = g.fechaCobro.split('/')
      return `${y}-${m.padStart(2, '0')}-${d.padStart(2, '0')}`
    }
    return g.fechaCobro
  })()

  const NumCell = ({ field, prefix, suffix, format = true, compact = false }) => {
    const v = g[field]
    // Money fields render as toFixed(2) (e.g. 13.07). `compact` rounds to 2
    // decimals but drops trailing zeros (e.g. 13 instead of 13.00) — used for
    // propinaPorcentaje and montoPropina so editing feels natural. 0 → '' for
    // cleanliness.
    const display = v
      ? compact ? +Number(v).toFixed(2)
      : format  ? Number(v).toFixed(2)
      :           v
      : ''
    return (
      <div className="num-cell">
        {prefix && <span className="sym">{prefix}</span>}
        <input
          type="number"
          className="cell-in"
          value={display}
          placeholder="0"
          onChange={e => upd(field, parseFloat(e.target.value) || 0)}
        />
        {suffix && <span className="sym">{suffix}</span>}
      </div>
    )
  }

  return (
    <tr className={g.hizoMatch ? 'row-match' : ''}>
      {/* Checkbox */}
      <td className="td-chk">
        <input
          type="checkbox"
          checked={g.checkManual}
          onChange={e => upd('checkManual', e.target.checked)}
        />
      </td>

      {/* Estado: PDF + Banco — pill badges */}
      <td className="td-status">
        <div className="status-row">
          <button
            className={`pill ${g.tienePDF ? 'pill-green' : 'pill-muted'}`}
            onClick={g.tienePDF ? () => openPDF(g.pdfFile) : undefined}
            disabled={!g.tienePDF}
            title={g.tienePDF ? 'Clic para abrir el PDF' : 'Sin PDF'}
          >
            PDF
          </button>
          <span
            className={`pill ${g.hizoMatch ? 'pill-blue' : 'pill-dim'}`}
            title={g.hizoMatch ? 'Conciliado con el Banco' : 'Pendiente en Banco'}
          >
            {g.hizoMatch ? '✓ OK' : '·'}
          </span>
        </div>
      </td>

      {/* Fecha Factura */}
      <td className="td-fecha">
        <input className="cell-in" value={dateDisplay} onChange={e => onDateChange(e.target.value)} />
      </td>

      {/* Fecha Cobro — span shows MM-DD-YY; click swaps to date picker
          (which uses YYYY-MM-DD natively). Blur returns to display mode. */}
      <td>
        {editingCobro ? (
          <input
            type="date"
            className="cell-date-input"
            autoFocus
            value={cobroIso}
            onChange={e => upd('fechaCobro', e.target.value)}
            onBlur={() => setEditingCobro(false)}
          />
        ) : (
          <span
            className="cell-cobro-display"
            onClick={() => setEditingCobro(true)}
          >
            {formatDateDisplay(g.fechaCobro) || <span className="cell-cobro-empty">—</span>}
          </span>
        )}
      </td>

      {/* Factura */}
      <td>
        <div className="fac-cell">
          <input className="cell-in is-mono" value={g.noFactura} onChange={e => upd('noFactura', e.target.value)} />
          {g.tienePDF && (
            <button className="ext-btn" onClick={() => openPDF(g.pdfFile)} title="Ver Factura">
              <svg width="12" height="12" viewBox="0 0 12 12" fill="none"><path d="M5 2H2a1 1 0 00-1 1v7a1 1 0 001 1h7a1 1 0 001-1V7" stroke="currentColor" strokeWidth="1.3" strokeLinecap="round"/><path d="M7.5 1H11v3.5M11 1L5.5 6.5" stroke="currentColor" strokeWidth="1.3" strokeLinecap="round" strokeLinejoin="round"/></svg>
            </button>
          )}
        </div>
      </td>

      {/* Proveedor */}
      <td>
        <input className="cell-in is-bold" value={g.proveedor} onChange={e => upd('proveedor', e.target.value)} />
      </td>

      {/* Concepto */}
      <td>
        <input className="cell-in is-dim" value={g.concepto} onChange={e => upd('concepto', e.target.value)} />
      </td>

      {/* Tipo — categoría de gasto auto-detectada del XML, editable vía dropdown */}
      <td>
        <select
          className="cell-select"
          value={g.tipo || ''}
          onChange={e => upd('tipo', e.target.value)}
        >
          <option value="">— Tipo —</option>
          {TIPOS_GASTO.map(t => (
            <option key={t} value={t}>{t}</option>
          ))}
        </select>
      </td>

      {/* Subtotal */}
      <td><NumCell field="importe"     prefix="$" /></td>

      {/* IVA */}
      <td><NumCell field="iva"         prefix="$" /></td>

      {/* ISR trasladado */}
      <td><NumCell field="isrTrasladado" prefix="$" /></td>

      {/* Ret. ISR */}
      <td><NumCell field="retencionISR" prefix="$" /></td>

      {/* Ret. IVA */}
      <td><NumCell field="retencionIVA" prefix="$" /></td>

      {/* Retenciones (total) */}
      <td><NumCell field="retenciones" prefix="$" /></td>

      {/* Total Factura */}
      <td><NumCell field="totalCFDI"   prefix="$" /></td>

      {/* Forma de Pago */}
      <td>
        <select
          className="cell-select"
          value={g.formaPago}
          onChange={e => upd('formaPago', e.target.value)}
        >
          <option value="04">04 - Tarjeta de Crédito</option>
          <option value="02">02 - Efectivo</option>
          <option value="03">03 - Transferencia</option>
          <option value="01">01 - Efectivo (otro)</option>
        </select>
      </td>

      {/* Propina % — compact: rounds to 2 decimals, drops trailing zeros for typing comfort */}
      <td><NumCell field="propinaPorcentaje" suffix="%" compact /></td>

      {/* Propina $ */}
      <td><NumCell field="montoPropina" prefix="$" compact /></td>

      {/* Total Final */}
      <td>
        <span className={`total-val${g.hizoMatch ? ' is-blue' : ''}`}>
          ${(g.totalCFDI + g.montoPropina).toFixed(2)}
        </span>
      </td>
    </tr>
  )
}

/* ═══════════════════════════════════════════════════
   COMPONENTE: MODAL DE CONCILIACIÓN (glass + framer-motion)
═══════════════════════════════════════════════════ */

// Eased count-up — drives the big KPI numbers from 0 → target on mount.
function useCountUp(target, duration = 900) {
  const [v, setV] = useState(0)
  useEffect(() => {
    let raf, start
    const step = t => {
      if (start === undefined) start = t
      const p = Math.min((t - start) / duration, 1)
      const eased = 1 - Math.pow(1 - p, 3)  // ease-out cubic
      setV(target * eased)
      if (p < 1) raf = requestAnimationFrame(step)
    }
    raf = requestAnimationFrame(step)
    return () => cancelAnimationFrame(raf)
  }, [target, duration])
  return v
}

function KpiCard({ variants, accent, Icon, value }) {
  return (
    <motion.div
      className={`cm-kpi cm-kpi-${accent}`}
      variants={variants}
      whileHover={{ y: -3, scale: 1.02 }}
      transition={{ type: 'spring', stiffness: 260, damping: 20 }}
    >
      <div className="cm-kpi-icon"><Icon size={16} strokeWidth={2.2} /></div>
      <div className="cm-kpi-value">{Math.round(value)}</div>
    </motion.div>
  )
}

function ConciliacionModal({ data, onClose }) {
  const cBanco    = useCountUp(data.bancoRows)
  const cMatches  = useCountUp(data.matches)
  const cPropinas = useCountUp(data.propinas)

  const stagger = {
    hidden: { opacity: 0 },
    show:   { opacity: 1, transition: { staggerChildren: 0.08, delayChildren: 0.15 } },
  }
  const item = {
    hidden: { opacity: 0, y: 14 },
    show:   { opacity: 1, y: 0, transition: { type: 'spring', stiffness: 240, damping: 22 } },
  }

  return (
    <motion.div
      className="cm-overlay"
      initial={{ opacity: 0 }}
      animate={{ opacity: 1 }}
      exit={{ opacity: 0 }}
      transition={{ duration: 0.2 }}
      onClick={onClose}
    >
      <motion.div
        className="cm-modal"
        initial={{ opacity: 0, y: 32, scale: 0.94 }}
        animate={{ opacity: 1, y: 0, scale: 1 }}
        exit={{ opacity: 0, y: 16, scale: 0.96 }}
        transition={{ type: 'spring', stiffness: 260, damping: 26 }}
        onClick={e => e.stopPropagation()}
      >
        <button className="cm-close" onClick={onClose} aria-label="Cerrar">
          <X size={16} />
        </button>

        {/* Header */}
        <div className="cm-header">
          <motion.div
            className="cm-orb"
            initial={{ scale: 0, rotate: -20 }}
            animate={{ scale: 1, rotate: 0 }}
            transition={{ type: 'spring', stiffness: 220, damping: 14, delay: 0.1 }}
          >
            <CheckCircle2 size={30} strokeWidth={2.2} />
          </motion.div>
          <h2 className="cm-title">Conciliación Terminada</h2>
          <p className="cm-subtitle">
            Procesamos {data.bancoRows} {data.bancoRows === 1 ? 'cargo' : 'cargos'} del estado de cuenta
          </p>
        </div>

        {/* KPI grid */}
        <motion.div className="cm-kpi-grid" variants={stagger} initial="hidden" animate="show">
          <KpiCard variants={item} accent="blue"   Icon={CreditCard} value={cBanco} />
          <KpiCard variants={item} accent="green"  Icon={Target}     value={cMatches} />
          <KpiCard variants={item} accent="purple" Icon={Sparkles}   value={cPropinas} />
        </motion.div>
        <motion.div
          className="cm-kpi-labels"
          initial={{ opacity: 0 }}
          animate={{ opacity: 1 }}
          transition={{ delay: 0.4 }}
        >
          <span>Cargos en banco</span>
          <span>Matches exitosos</span>
          <span>Propinas detectadas</span>
        </motion.div>

        {/* Cargos sin factura */}
        {data.sinFactura.length > 0 && (
          <motion.section
            className="cm-section"
            initial={{ opacity: 0, y: 12 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ delay: 0.45 }}
          >
            <div className="cm-section-header cm-warn">
              <AlertTriangle size={14} strokeWidth={2.4} />
              <span>Cargos sin factura</span>
              <span className="cm-count-pill cm-warn">{data.sinFactura.length}</span>
            </div>
            <div className="cm-alert-list">
              {data.sinFactura.slice(0, 8).map((s, i) => (
                <motion.div
                  key={i}
                  className="cm-alert"
                  initial={{ opacity: 0, x: -8 }}
                  animate={{ opacity: 1, x: 0 }}
                  transition={{ delay: 0.5 + i * 0.04, type: 'spring', stiffness: 280, damping: 24 }}
                >
                  <div className="cm-alert-date">{formatDateDisplay(s.fecha)}</div>
                  <div className="cm-alert-desc" title={s.descripcion || ''}>
                    {s.descripcion || 'Sin descripción'}
                  </div>
                  <div className="cm-alert-amount">
                    ${s.monto.toLocaleString('es-MX', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                  </div>
                </motion.div>
              ))}
              {data.sinFactura.length > 8 && (
                <div className="cm-more">+{data.sinFactura.length - 8} más</div>
              )}
            </div>
          </motion.section>
        )}

        {/* Facturas sin cargo */}
        {data.facturasSinCargo > 0 && (
          <motion.section
            className="cm-section"
            initial={{ opacity: 0, y: 12 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ delay: 0.55 }}
          >
            <div className="cm-section-header cm-neutral">
              <FileText size={14} strokeWidth={2.4} />
              <span>Facturas sin cargo en banco</span>
              <span className="cm-count-pill cm-neutral">{data.facturasSinCargo}</span>
            </div>
          </motion.section>
        )}

        {/* CTA */}
        <motion.button
          className="cm-cta"
          onClick={onClose}
          whileHover={{ y: -1 }}
          whileTap={{ scale: 0.98 }}
          initial={{ opacity: 0, y: 8 }}
          animate={{ opacity: 1, y: 0 }}
          transition={{ delay: 0.62 }}
        >
          Listo
        </motion.button>
      </motion.div>
    </motion.div>
  )
}

/* ═══════════════════════════════════════════════════
   COMPONENTE: MODAL DE ÉXITO AL EXPORTAR (Excel / ZIP)
═══════════════════════════════════════════════════ */

// Generic success modal — driven entirely by the `data` object so it can
// represent any export flow. Caller passes a pre-formatted `meta` string
// (e.g. "17 registros · 24.5 KB") and an optional muted `note` footer.
function ExportSuccessModal({ data, onClose }) {
  const { title, subtitle, filename, meta, Icon, note } = data
  return (
    <motion.div
      className="cm-overlay"
      initial={{ opacity: 0 }}
      animate={{ opacity: 1 }}
      exit={{ opacity: 0 }}
      transition={{ duration: 0.2 }}
      onClick={onClose}
    >
      <motion.div
        className="cm-modal cm-modal-narrow"
        initial={{ opacity: 0, y: 32, scale: 0.94 }}
        animate={{ opacity: 1, y: 0, scale: 1 }}
        exit={{ opacity: 0, y: 16, scale: 0.96 }}
        transition={{ type: 'spring', stiffness: 260, damping: 26 }}
        onClick={e => e.stopPropagation()}
      >
        <button className="cm-close" onClick={onClose} aria-label="Cerrar">
          <X size={16} />
        </button>

        {/* Header */}
        <div className="cm-header">
          <motion.div
            className="cm-orb"
            initial={{ scale: 0, rotate: -20 }}
            animate={{ scale: 1, rotate: 0 }}
            transition={{ type: 'spring', stiffness: 220, damping: 14, delay: 0.1 }}
          >
            <CheckCircle2 size={30} strokeWidth={2.2} />
          </motion.div>
          <h2 className="cm-title">{title}</h2>
          <p className="cm-subtitle">{subtitle}</p>
        </div>

        {/* File card */}
        <motion.div
          className="cm-file-card"
          initial={{ opacity: 0, y: 14 }}
          animate={{ opacity: 1, y: 0 }}
          transition={{ delay: 0.25, type: 'spring', stiffness: 260, damping: 24 }}
          whileHover={{ y: -2 }}
        >
          <div className="cm-file-icon">
            <Icon size={20} strokeWidth={2} />
          </div>
          <div className="cm-file-info">
            <div className="cm-file-name">{filename}</div>
            <div className="cm-file-meta">{meta}</div>
          </div>
          <motion.div
            className="cm-file-check"
            initial={{ scale: 0 }}
            animate={{ scale: 1 }}
            transition={{ delay: 0.45, type: 'spring', stiffness: 320, damping: 16 }}
            aria-hidden="true"
          >
            <Check size={14} strokeWidth={3} />
          </motion.div>
        </motion.div>

        {/* Optional footnote — used by ZIP export to reassure about originals. */}
        {note && (
          <motion.div
            className="cm-note"
            initial={{ opacity: 0, y: 6 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ delay: 0.42 }}
          >
            {note}
          </motion.div>
        )}

        {/* CTA */}
        <motion.button
          className="cm-cta"
          onClick={onClose}
          whileHover={{ y: -1 }}
          whileTap={{ scale: 0.98 }}
          initial={{ opacity: 0, y: 8 }}
          animate={{ opacity: 1, y: 0 }}
          transition={{ delay: note ? 0.55 : 0.5 }}
        >
          Listo
        </motion.button>
      </motion.div>
    </motion.div>
  )
}

/* ═══════════════════════════════════════════════════
   APP PRINCIPAL
═══════════════════════════════════════════════════ */

export default function App() {
  const [lista,         setLista]         = useState([])
  const [carpetaNombre, setCarpetaNombre] = useState('Ninguna carpeta seleccionada')
  const [alerta,        setAlerta]        = useState(null)
  const [conciliacion,  setConciliacion]  = useState(null)
  const [exportExito,   setExportExito]   = useState(null)
  const [loading,       setLoading]       = useState(false)
  const [isDragging,    setIsDragging]    = useState(false)
  // Index-based fixed pixel widths — order matches COLUMNS positions:
  // [0] checkbox, [1] estado, [2] fecha factura, [3] fecha cobro,
  // [4] factura, [5] proveedor, [6] concepto, [7] tipo, [8] subtotal,
  // [9] iva, [10] isr/ish/ieps, [11] ret.isr, [12] ret.iva, [13] reten,
  // [14] total fac, [15] forma pago, [16] prop%, [17] prop$, [18] total final
  const [colWidths, setColWidths] = useState([40, 110, 115, 120, 120, 260, 140, 100, 120, 110, 135, 110, 110, 110, 125, 160, 95, 105, 130])
  const [sort,          setSort]          = useState({ field: null, dir: 'asc' })

  const folderRef = useRef(null)
  const bancoRef  = useRef(null)

  // ── Métricas (cards en el encabezado de la tabla) ──
  const metrics = useMemo(() => {
    const sum = field => lista.reduce((s, g) => s + (g[field] || 0), 0)
    return {
      totalFacturado:   sum('totalCFDI'),
      ivaTotal:         sum('iva'),
      retencionesTotal: sum('retenciones'),
      sinCobrar:        lista.filter(g => !g.fechaCobro).length,
      count:            lista.length,
    }
  }, [lista])
  const fmtMoney = n => `$${n.toLocaleString('es-MX', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`

  // ── Sort — three-state cycle (asc → desc → unsorted), null-safe, with
  //          date detection that accepts both YYYY-MM-DD and DD/MM/YYYY.
  const sortedLista = useMemo(() => {
    if (!sort.field) return lista
    const col = COLUMNS.find(c => c.key === sort.field)
    if (!col) return lista
    const get = col.getValue || (g => g[sort.field])
    const toDate = s => {
      if (typeof s !== 'string' || !s) return 0
      if (/^\d{4}-\d{2}-\d{2}/.test(s)) return new Date(s).getTime()
      if (/^\d{2}\/\d{2}\/\d{4}/.test(s)) {
        const [d, m, y] = s.split('/')
        return new Date(`${y}-${m}-${d}`).getTime()
      }
      return 0
    }
    return [...lista].sort((a, b) => {
      const va = get(a), vb = get(b)
      // Null/undefined/empty → push to the end regardless of direction
      const aEmpty = va == null || va === ''
      const bEmpty = vb == null || vb === ''
      if (aEmpty && bEmpty) return 0
      if (aEmpty) return 1
      if (bEmpty) return -1
      let cmp
      if (col.type === 'number')      cmp = (parseFloat(va) || 0) - (parseFloat(vb) || 0)
      else if (col.type === 'date')   cmp = toDate(va) - toDate(vb)
      else                            cmp = String(va).localeCompare(String(vb), 'es')
      return sort.dir === 'asc' ? cmp : -cmp
    })
  }, [lista, sort])

  // Three-state cycle: first click → asc, second → desc, third → unsorted.
  const toggleSort = field => setSort(s => {
    if (s.field === field) {
      if (s.dir === 'asc')  return { field, dir: 'desc' }
      if (s.dir === 'desc') return { field: null, dir: null }
    }
    return { field, dir: 'asc' }
  })

  // ── Column resize (drag right edge of th) ──
  const startResize = (colIndex, e) => {
    e.preventDefault(); e.stopPropagation()
    const startX = e.clientX
    const startWidth = colWidths[colIndex]
    const onMouseMove = ev => {
      const delta = ev.clientX - startX
      const newWidth = Math.max(40, startWidth + delta)
      setColWidths(prev => {
        const next = [...prev]
        next[colIndex] = newWidth
        return next
      })
    }
    const onMouseUp = () => {
      document.removeEventListener('mousemove', onMouseMove)
      document.removeEventListener('mouseup',   onMouseUp)
    }
    document.addEventListener('mousemove', onMouseMove)
    document.addEventListener('mouseup',   onMouseUp)
  }

  // ── Actualizar campo de un gasto (con recálculo de propina) ──
  const update = (id, field, value) =>
    setLista(prev => prev.map(g => {
      if (g.id !== id) return g
      const u = { ...g, [field]: value }
      if (field === 'totalCFDI' && u.propinaPorcentaje > 0)
        u.montoPropina = Math.round(value * u.propinaPorcentaje / 100 * 100) / 100
      if (field === 'propinaPorcentaje')
        u.montoPropina = Math.round(g.totalCFDI * value / 100 * 100) / 100
      if (field === 'montoPropina' && g.totalCFDI > 0)
        u.propinaPorcentaje = Math.round((value / g.totalCFDI) * 10000) / 100
      if (field === 'retencionISR') u.retenciones = value + (u.retencionIVA || 0)
      if (field === 'retencionIVA') u.retenciones = value + (u.retencionISR || 0)
      return u
    }))

  // ── Cargar carpeta (XMLs + PDFs) ──
  // Shared file-processing pipeline — used by folder picker and drag/drop.
  const processFiles = async (files, folderName) => {
    if (!files.length) return
    const xmls = files.filter(f => f.name.toLowerCase().endsWith('.xml'))
    const pdfs = files.filter(f => f.name.toLowerCase().endsWith('.pdf'))
    setCarpetaNombre(folderName)
    setLoading(true)
    const nueva = []
    for (const f of xmls) {
      try {
        const text = await f.text()
        const g = parseCFDI(text, f, pdfs)
        if (g) nueva.push(g)
      } catch {}
    }
    setLista(nueva)
    setLoading(false)
  }

  const cargar = async e => {
    const files = Array.from(e.target.files)
    const folder = files[0]?.webkitRelativePath?.split('/')[0] || 'Carpeta'
    await processFiles(files, folder)
    e.target.value = ''
  }

  const onDropFiles = async e => {
    e.preventDefault()
    setIsDragging(false)
    const files = Array.from(e.dataTransfer.files)
    const folder = files[0]?.webkitRelativePath?.split('/')[0] || 'Archivos arrastrados'
    await processFiles(files, folder)
  }

  // ── Validar estado de cuenta bancario ──
  const validarBanco = async e => {
    const file = e.target.files[0]; if (!file) return
    const raw = await file.text()
    const content = raw.replace(/\r\n/g, '\n').replace(/\r/g, '\n')
    const lines = content.split('\n')
    const sample = content.slice(0, 500)
    let sep = ','
    if (!sample.includes(',')) sep = sample.includes(';') ? ';' : '\t'

    // Clara-platform CSVs have a fixed column layout (date col 0, MXN amount
    // col 5, merchant col 2). Detecting by header so we can use fixed indices
    // instead of the generic scanner that mistakes card numbers and billing
    // periods for amounts. "Transacci" intentionally drops the accented ó so
    // we survive encoding mishaps.
    const isClara = (lines[0] || '').includes('Fecha de Transacci')

    let matches = 0, propinas = 0
    const sinFactura = []
    const nl = lista.map(g => ({ ...g, hizoMatch: false, fechaCobro: '' }))
    const formatCobro = d => {
      // Store as YYYY-MM-DD so the native date picker on the row accepts it
      // directly and formatDateDisplay produces MM-DD-YY for the read view.
      const dd   = String(d.getDate()).padStart(2, '0')
      const mm   = String(d.getMonth() + 1).padStart(2, '0')
      const yyyy = d.getFullYear()
      return `${yyyy}-${mm}-${dd}`
    }
    const cleanNum = s => parseFloat(String(s || '').replace(/[$,\s]/g, ''))

    // Step 1 — parse every CSV row into a flat list { dCSV, amounts[], descripcion }.
    const csvRows = []
    for (let li = 0; li < lines.length; li++) {
      const line = lines[li]
      if (!line.trim()) continue
      if (isClara && li === 0) continue   // skip Clara header row

      const cols = parseCSVLine(line, sep)
      let dCSV = null
      let amounts = []
      let descripcion = null

      if (isClara) {
        dCSV = parseDateRobusto(cols[0] || '')
        const mxn  = cleanNum(cols[5])
        const orig = cleanNum(cols[3])
        const amount = (!isNaN(mxn) && mxn) || (!isNaN(orig) && orig) || 0
        if (!dCSV || !amount) continue
        amounts = [Math.abs(amount)]
        descripcion = (cols[2] || '').trim().slice(0, 60)
      } else {
        for (const cell of cols) {
          const c = cell.trim(); if (!c) continue
          if (!dCSV) { const d = parseDateRobusto(c); if (d) { dCSV = d; continue } }
          if (c.includes('/') && !c.includes('$')) continue
          const s = c.replace(/[$,]/g, '').trim()
          const v = parseFloat(s)
          if (!isNaN(v) && Math.abs(v) > 0) amounts.push(Math.abs(v))
          else {
            const n = s.replace(/[^0-9.\-]/g, '')
            const v2 = parseFloat(n)
            if (!isNaN(v2) && Math.abs(v2) > 0) amounts.push(Math.abs(v2))
          }
        }
        if (!dCSV || !amounts.length) continue
        descripcion = line.trim().slice(0, 60)
      }
      csvRows.push({ dCSV, amounts, descripcion, matched: false })
    }
    const bancoRows = csvRows.length

    // Date is a TIEBREAKER, not a filter. Pick the invoice whose fechaFac is
    // closest in time to the bank row's date. Invoices with no parseable date
    // get pushed to the back via Infinity.
    const closestByDate = (candIdx, dCSV) => {
      if (candIdx.length === 1) return candIdx[0]
      return candIdx.reduce((bestIdx, ci) => {
        const dG_c = parseDateRobusto(nl[ci].fechaFac)
        const dG_b = parseDateRobusto(nl[bestIdx].fechaFac)
        const diff_c = dG_c ? Math.abs(dCSV - dG_c) : Infinity
        const diff_b = dG_b ? Math.abs(dCSV - dG_b) : Infinity
        return diff_c < diff_b ? ci : bestIdx
      })
    }

    // Helper: try to match one row's amounts under a given predicate. On a
    // match, hand off to `apply` which mutates nl[idx] and returns the
    // pass-specific bookkeeping increments.
    const tryPass = (predicate, apply) => {
      for (const row of csvRows) {
        if (row.matched) continue
        for (const monto of row.amounts) {
          const candidates = []
          for (let i = 0; i < nl.length; i++) {
            if (nl[i].hizoMatch) continue
            if (predicate(nl[i], monto)) candidates.push(i)
          }
          if (!candidates.length) continue
          const idx = closestByDate(candidates, row.dCSV)
          apply(idx, monto, row.dCSV)
          row.matched = true
          break
        }
      }
    }

    // Pass 1 — exact match (±$5) against totalCFDI + montoPropina.
    tryPass(
      (inv, m) => Math.abs(inv.totalCFDI + inv.montoPropina - m) <= 5,
      (idx, _m, dCSV) => {
        nl[idx].hizoMatch = true
        nl[idx].fechaCobro = formatCobro(dCSV)
        nl[idx].formaPago = '04'  // bank-matched → card transaction
        matches++
      }
    )

    // Pass 2 — propina: bank charge = invoice base + tip.
    //   Standard invoices (≥$100): tip is 5–25% of base, with a $5 absolute
    //     floor so a sub-$5 nick doesn't masquerade as a 5% "propina" on a
    //     small invoice (Pass 1's ±$5 already absorbs those).
    //   Small invoices (<$100): the percentage window collapses, so allow
    //     any positive tip up to $30 absolute.
    tryPass(
      (inv, m) => {
        const base = inv.totalCFDI
        if (base <= 0) return false
        const diff = m - base
        if (diff <= 0) return false
        if (base < 100) return diff <= 30
        const minPropina = Math.max(base * 0.05, 5.0)
        const maxPropina = base * 0.25
        return diff >= minPropina && diff <= maxPropina
      },
      (idx, m, dCSV) => {
        const base = nl[idx].totalCFDI
        const prop = Math.round((m - base) * 100) / 100
        const pct  = base > 0 ? Math.round((prop / base) * 10000) / 100 : 0
        nl[idx].hizoMatch = true
        nl[idx].fechaCobro = formatCobro(dCSV)
        nl[idx].formaPago = '04'  // bank-matched → card transaction
        nl[idx].montoPropina = prop
        nl[idx].propinaPorcentaje = pct
        matches++; propinas++
      }
    )

    // Pass 3 — relaxed amount (±$1) against totalCFDI only (no propina).
    // Catches rounding nicks (e.g. 5851.38 vs 5851.25) on invoices without
    // a propina already recorded.
    tryPass(
      (inv, m) => Math.abs(inv.totalCFDI - m) <= 1.0,
      (idx, _m, dCSV) => {
        nl[idx].hizoMatch = true
        nl[idx].fechaCobro = formatCobro(dCSV)
        nl[idx].formaPago = '04'  // bank-matched → card transaction
        matches++
      }
    )

    // Collect unmatched rows for the result modal.
    for (const row of csvRows) {
      if (row.matched) continue
      sinFactura.push({
        fecha: formatCobro(row.dCSV),
        monto: Math.max(...row.amounts),
        descripcion: row.descripcion,
      })
    }

    setLista(nl)
    setConciliacion({
      bancoRows,
      matches,
      propinas,
      sinFactura,
      facturasSinCargo: nl.length - matches,
    })
    e.target.value = ''
  }

  // ── Agregar fila manual ──
  const agregarManual = () => {
    const hoy = new Date().toISOString().slice(0, 10)
    setLista(prev => [...prev, {
      id: genId(),
      rfc: 'PUBLICO GENERAL', proveedor: 'Escribe aquí...', noFactura: 'Ticket',
      fechaFac: hoy, concepto: 'Consumo', tipo: 'Factura', importe: 0, iva: 0, isrTrasladado: 0,
      retencionISR: 0, retencionIVA: 0, retenciones: 0, totalCFDI: 0,
      propinaPorcentaje: 0, montoPropina: 0, fechaCobro: hoy, formaPago: '01', uuid: 'MANUAL',
      tienePDF: false, pdfFile: null, xmlFile: null, hizoMatch: false, checkManual: false,
    }])
  }

  // ── Copiar a portapapeles (TSV para Excel) ──
  const copiar = () => {
    const hdr = 'RFC PROVEEDOR\tPROVEEDOR\tNO. DE FACTURA\tFECHA FAC.\tCONCEPTO\tIMPORTE (MXP)\tIVA\tISR\tRET. ISR\tRET. IVA\tRET/ ISR IVA\tTOTAL CFDI\tGastos en USD\tTipo de Cambio\tTotal Checking\tFORMA DE PAGO\tFECHA DE COBRO\n'
    const rows = lista.flatMap(g => {
      const fac   = formatDateDisplay(g.fechaFac)
      const cobro = formatDateDisplay(g.fechaCobro) || 'Pendiente'
      const r  = `${g.rfc}\t${g.proveedor.replace(/\t/g,' ')}\t${g.noFactura}\t${fac}\t${g.concepto.replace(/\t/g,' ')}\t${g.importe.toFixed(2)}\t${g.iva.toFixed(2)}\t${(g.isrTrasladado||0).toFixed(2)}\t${(g.retencionISR||0).toFixed(2)}\t${(g.retencionIVA||0).toFixed(2)}\t${g.retenciones.toFixed(2)}\t${g.totalCFDI.toFixed(2)}\t\t\t\t${g.formaPago}\t${cobro}\n`
      const p  = g.montoPropina > 0
        ? `\t${g.proveedor} - PROPINA\t\t${fac}\tPROPINA\t${g.montoPropina.toFixed(2)}\t0.00\t0.00\t0.00\t0.00\t0.00\t${g.montoPropina.toFixed(2)}\t\t\t\t${g.formaPago}\t${cobro}\n`
        : ''
      return [r, p]
    })
    navigator.clipboard.writeText(hdr + rows.join(''))
      .then(() => setAlerta('¡Copiado al portapapeles! 🎉\n\nVe a tu Excel, haz clic en la celda donde quieres los datos y presiona Ctrl+V (o Cmd+V en Mac) para pegar.'))
      .catch(() => setAlerta('Error al copiar. Verifica los permisos del navegador.'))
  }

  // ── Exportar ZIP con CSV + archivos renombrados ──
  const exportar = async () => {
    const zip    = new JSZip()
    const folder = zip.folder('Reporte_Gastos')
    let csv = '\uFEFFRFC PROVEEDOR,PROVEEDOR,NO. DE FACTURA,FECHA FAC.,CONCEPTO,IMPORTE (MXP),IVA,ISR,RET. ISR,RET. IVA,RET/ ISR IVA,TOTAL CFDI,Gastos en USD,Tipo de Cambio,Total Checking,FORMA DE PAGO,FECHA DE COBRO\n'
    let r = 0

    for (const g of lista) {
      const fac   = formatDateDisplay(g.fechaFac)          // MM-DD-YY for CSV cells
      const fa    = fac                                    // also used for ZIP filename
      const cobro = formatDateDisplay(g.fechaCobro) || 'Pendiente'
      csv += `${g.rfc},${g.proveedor.replace(/,/g,' ')},${g.noFactura},${fac},${g.concepto.replace(/,/g,' ')},${g.importe.toFixed(2)},${g.iva.toFixed(2)},${(g.isrTrasladado||0).toFixed(2)},${(g.retencionISR||0).toFixed(2)},${(g.retencionIVA||0).toFixed(2)},${g.retenciones.toFixed(2)},${g.totalCFDI.toFixed(2)},,,,${g.formaPago},${cobro}\n`
      if (g.montoPropina > 0)
        csv += `,${g.proveedor} - PROPINA,,${fac},PROPINA,${g.montoPropina.toFixed(2)},0.00,0.00,0.00,0.00,0.00,${g.montoPropina.toFixed(2)},,,,${g.formaPago},${cobro}\n`

      if (g.xmlFile) {
        const titleCase = s => s.toLowerCase().replace(/(^|\s)\S/g, c => c.toUpperCase())
        const pf  = titleCase(g.proveedor.replace(/,/g, ''))
        const nom = `${pf} ${g.noFactura} ${g.concepto} ${fa}`.replace(/[\/\\:*?"<>|]/g, '')
        folder.file(`${nom}.xml`, await g.xmlFile.arrayBuffer()); r++
        if (g.pdfFile) folder.file(`${nom}.pdf`, await g.pdfFile.arrayBuffer())
      }
    }

    folder.file('Reporte_Gastos_Final.csv', csv)
    const blob = await zip.generateAsync({ type: 'blob' })
    const filename = 'Reporte_Gastos_Empaquetado.zip'
    const a = Object.assign(document.createElement('a'), {
      href: URL.createObjectURL(blob),
      download: filename,
    })
    a.click(); URL.revokeObjectURL(a.href)
    setExportExito({
      title: '¡ZIP Generado!',
      subtitle: 'Paquete descargado con CSV + facturas renombradas',
      filename,
      meta: `${r} ${r === 1 ? 'factura' : 'facturas'} · ${formatBytes(blob.size)}`,
      Icon: Package,
      note: 'Tus archivos originales siguen intactos.',
    })
  }

  // ── Exportar a Excel ──
  // Defers to /api/export-excel (Python serverless function on Vercel) so the
  // .xls file gets full template formatting via xlrd + xlutils + xlwt —
  // something community SheetJS can't preserve in the browser.
  const exportarExcel = async () => {
    if (!lista.length) return
    try {
      setAlerta(null)
      const response = await fetch('/api/export-excel', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(lista),
      })
      if (!response.ok) {
        try {
          const errData = await response.json()
          setAlerta('Error al generar Excel:\n\n' + (errData.error || response.status) + '\n\n' + (errData.trace || '').slice(0, 300))
        } catch {
          setAlerta('Error al generar Excel: API error ' + response.status)
        }
        return
      }
      const blob = await response.blob()
      const url = URL.createObjectURL(blob)
      const filename = 'Reporte_Gastos_SMTO.xlsx'
      const a = Object.assign(document.createElement('a'), { href: url, download: filename })
      a.click()
      URL.revokeObjectURL(url)
      setExportExito({
        title: '¡Excel Generado!',
        subtitle: 'Reporte descargado en formato oficial SMTO',
        filename,
        meta: `${lista.length} ${lista.length === 1 ? 'registro' : 'registros'} · ${formatBytes(blob.size)}`,
        Icon: FileSpreadsheet,
      })
    } catch (err) {
      setAlerta('Error al generar Excel: ' + err.message)
    }
  }

  // ── Abrir PDF en nueva pestaña ──
  const openPDF = pdfFile => {
    if (!pdfFile) return
    try {
      const url = URL.createObjectURL(pdfFile)
      const win = window.open(url, '_blank')
      if (!win) setAlerta('No se pudo abrir el PDF.\n\nVerifica que el navegador permita ventanas emergentes para este sitio.')
    } catch (err) {
      setAlerta(`Error al abrir el PDF:\n\n${err && err.message ? err.message : String(err)}`)
    }
  }

  /* ── RENDER ── */
  return (
    <div className="app">

      {/* ─── CABECERA ─── */}
      <div className="header">
        <div className="header-logo">
          <img src="/logo.png" alt="SMTO" style={{ height: '54px', width: 'auto', objectFit: 'contain' }} />
        </div>
        <div className="header-info">
          <h1 className="header-title">Reporte de Gastos SMTO<span className="version-badge">v4.9</span></h1>
          <div className="header-sub">
            <span className="sub-folder">
              <svg width="13" height="11" viewBox="0 0 13 11" fill="currentColor" style={{marginRight:4,verticalAlign:'middle'}}><path d="M1 2.5A1.5 1.5 0 012.5 1H5l1.5 1.5H11A1.5 1.5 0 0112.5 4V9A1.5 1.5 0 0111 10.5H2A1.5 1.5 0 01.5 9V2.5z" fill="currentColor"/></svg>
              {carpetaNombre}
            </span>
            <span className="sub-dot">•</span>
            <span className="sub-count">
              <svg width="12" height="12" viewBox="0 0 12 12" fill="currentColor" style={{marginRight:4,verticalAlign:'middle'}}><rect x=".5" y="1.5" width="11" height="9" rx="1.5" fill="none" stroke="currentColor" strokeWidth="1.2"/><path d="M3 4.5h6M3 6.5h4" stroke="currentColor" strokeWidth="1.2" strokeLinecap="round"/></svg>
              {lista.length} registros
            </span>
          </div>
        </div>
      </div>

      <div className="divider" />

      {/* ─── BARRA DE ACCIONES ─── */}
      <div className="action-bar">
        <div className="action-group">
          <PremiumButton title="Manual"         icon="＋"  variant="ghost"     onClick={agregarManual} />
          <PremiumButton title="Cargar Carpeta" icon="📂" variant="primary"   onClick={() => folderRef.current?.click()} />
          <PremiumButton title="Validar Banco"  icon="🏦" variant="secondary" onClick={() => bancoRef.current?.click()} />
        </div>
        <div className="action-group">
          <PremiumButton title="Copiar a Excel"  icon="📋" variant="copy"    isDisabled={!lista.length} onClick={copiar} />
          <PremiumButton title="Exportar a Excel" icon="📊" variant="primary" isDisabled={!lista.length} onClick={exportarExcel} />
          <PremiumButton title="Exportar a ZIP"  icon="📦" variant="export"  isDisabled={!lista.length} onClick={exportar} />
        </div>
      </div>

      {/* ─── BARRA DE MÉTRICAS ─── */}
      <div className="metrics-bar">
        <div className="metric-card" style={{ '--accent': '#0A84FF' }}>
          <div className="metric-label">Total Facturado</div>
          <div className="metric-value">{fmtMoney(metrics.totalFacturado)}</div>
        </div>
        <div className="metric-card" style={{ '--accent': '#BF5AF2' }}>
          <div className="metric-label">IVA Total</div>
          <div className="metric-value">{fmtMoney(metrics.ivaTotal)}</div>
        </div>
        <div className="metric-card" style={{ '--accent': '#FF9500' }}>
          <div className="metric-label">Retenciones</div>
          <div className="metric-value">{fmtMoney(metrics.retencionesTotal)}</div>
        </div>
        <div className="metric-card" style={{ '--accent': '#30D158' }}>
          <div className="metric-label">Registros</div>
          <div className="metric-value">{metrics.count}</div>
        </div>
        <div className="metric-card" style={{ '--accent': '#FF453A' }}>
          <div className="metric-label">Por Corroborar</div>
          <div className="metric-value">{metrics.sinCobrar}</div>
        </div>
      </div>

      {/* Inputs de archivo ocultos */}
      <input
        ref={el => { folderRef.current = el; if (el) { el.setAttribute('webkitdirectory',''); el.setAttribute('directory','') } }}
        type="file" multiple style={{ display: 'none' }}
        onChange={cargar}
      />
      <input
        ref={bancoRef}
        type="file" accept=".csv,.txt,.tsv" style={{ display: 'none' }}
        onChange={validarBanco}
      />

      <div className="divider" />

      {/* ─── TABLA ─── */}
      <div className="table-wrap">
        {loading ? (
          <div className="loading-msg">
            <div className="loading-spinner" />
            <div className="loading-text">Procesando facturas XML…</div>
          </div>
        ) : lista.length === 0 ? (
          <div
            className={`onboarding ${isDragging ? 'is-dragging' : ''}`}
            onDragOver={e => { e.preventDefault(); if (!isDragging) setIsDragging(true) }}
            onDragLeave={e => { if (!e.currentTarget.contains(e.relatedTarget)) setIsDragging(false) }}
            onDrop={onDropFiles}
          >
            <div className="onboarding-glow" />
            <div className="onboarding-card">
              <div className="onboarding-icon">
                {isDragging ? (
                  <svg width="36" height="36" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round">
                    <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                    <polyline points="17 8 12 3 7 8"/>
                    <line x1="12" y1="3" x2="12" y2="15"/>
                  </svg>
                ) : (
                  <svg width="36" height="36" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round">
                    <path d="M6 14l1.45-2.9A2 2 0 0 1 9.24 10H20a2 2 0 0 1 1.94 2.5l-1.55 6a2 2 0 0 1-1.94 1.5H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h3.93a2 2 0 0 1 1.66.9l.82 1.2a2 2 0 0 0 1.66.9H18a2 2 0 0 1 2 2v2"/>
                  </svg>
                )}
              </div>
              <div className="onboarding-title">Carga tus facturas XML</div>
              <div className="onboarding-sub">
                {isDragging
                  ? 'Suelta aquí tus archivos'
                  : 'Arrastra una carpeta o selecciónala para comenzar'}
              </div>
              <button className="onboarding-cta" onClick={() => folderRef.current?.click()}>
                Cargar Carpeta
              </button>
              <div className="onboarding-hint">Compatible con CFDI 3.3 y 4.0 · XML + PDF</div>
            </div>
          </div>
        ) : (
          <table style={{
            // Sticky-column left offsets — derived from indices 1 (Estado) and
            // 2 (Fecha) so resizing them shifts the pinned cells correctly.
            '--sl-fecha':     `${colWidths[1]}px`,
          }}>
            <thead>
              <tr>
                {COLUMNS.map((col, idx) => (
                  <th
                    key={col.key}
                    className={col.key === 'status' ? 'th-status' : col.key === 'fechaFac' ? 'th-fecha' : undefined}
                    style={{ width: colWidths[idx], cursor: col.sortable ? 'pointer' : undefined }}
                    onClick={col.sortable ? () => toggleSort(col.key) : undefined}
                  >
                    {col.label}
                    {sort.field === col.key && sort.dir && (
                      <span className="sort-arrow">{sort.dir === 'asc' ? '▲' : '▼'}</span>
                    )}
                    <div
                      className="col-resizer"
                      onMouseDown={e => startResize(idx, e)}
                      onClick={e => e.stopPropagation()}
                    />
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {sortedLista.map(g => (
                <GastoRow
                  key={g.id}
                  g={g}
                  upd={(field, val) => update(g.id, field, val)}
                  openPDF={openPDF}
                />
              ))}
            </tbody>
          </table>
        )}
      </div>

      {/* ─── MODAL ALERTA (texto simple — usado por copy/export/PDF) ─── */}
      {alerta && (
        <div className="overlay" onClick={() => setAlerta(null)}>
          <div className="modal" onClick={e => e.stopPropagation()}>
            <div className="modal-title">Resultado</div>
            <pre className="modal-body">{alerta}</pre>
            <button className="modal-ok" onClick={() => setAlerta(null)}>OK</button>
          </div>
        </div>
      )}

      {/* ─── MODAL CONCILIACIÓN BANCARIA (premium glass) ─── */}
      <AnimatePresence>
        {conciliacion && (
          <ConciliacionModal data={conciliacion} onClose={() => setConciliacion(null)} />
        )}
      </AnimatePresence>

      {/* ─── MODAL ÉXITO AL EXPORTAR EXCEL (premium glass) ─── */}
      <AnimatePresence>
        {exportExito && (
          <ExportSuccessModal data={exportExito} onClose={() => setExportExito(null)} />
        )}
      </AnimatePresence>
    </div>
  )
}
