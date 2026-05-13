import { useState, useRef, useMemo, useEffect } from 'react'
import JSZip from 'jszip'
import * as XLSX from 'xlsx'
import { motion, AnimatePresence } from 'framer-motion'
import { CheckCircle2, X, CreditCard, Target, Sparkles, AlertTriangle, FileText, FileSpreadsheet, Package, Check, Plus, Link2 } from 'lucide-react'

/* Two type lists — picked by colaborador.categoria. Ventas/Socio see the
   sales-flavored list (Hotel Ventas, Gasolina Ventas, …), everyone else
   (Admin/Servicio) sees the operational list. */
const TIPOS_VENTAS = [
  'Avión Ventas',
  'Casetas Ventas',
  'Estacionamiento Ventas',
  'Gasolina Ventas',
  'Gasolina Viáticos',
  'Gastos Rep (Representación)',
  'Gastos Rep Viáticos',
  'Herramientas Ventas',
  'Hotel Ventas',
  'Marketing',
]

const TIPOS_NORMALES = [
  'Avión',
  'Casetas',
  'Celular',
  'Consumo',
  'Consumo Viáticos',
  'Envíos',
  'Estacionamiento',
  'Gasolina',
  'Herramientas',
  'Hotel',
  'IT & SW (Software/Sistemas)',
  'Manto Auto (Mantenimiento)',
  'No Comprobado',
  'Papelería',
  'PC (Equipo)',
  'Rechazada',
  'Renta Auto',
  'Renta Oficina',
  'Taxi',
  'Traspaso',
  'Uniformes',
]

const getTiposForColaborador = (colaborador) => {
  if (!colaborador) return TIPOS_NORMALES
  const cat = colaborador.categoria
  if (cat === 'Ventas' || cat === 'Socio') return TIPOS_VENTAS
  return TIPOS_NORMALES  // Admin, Servicio
}

const autoDetectTipo = (descripcion, categoria) => {
  const d = (descripcion || '').toLowerCase()
  const isVentas = categoria === 'Ventas' || categoria === 'Socio'

  if (d.includes('hotel') || d.includes('hospedaje')) return isVentas ? 'Hotel Ventas' : 'Hotel'
  if (d.includes('vuelo') || d.includes('avion') || d.includes('tarifa aerea')) return isVentas ? 'Avión Ventas' : 'Avión'
  if (d.includes('taxi') || d.includes('uber') || d.includes('didi')) return isVentas ? 'Gastos Rep (Representación)' : 'Taxi'
  if (d.includes('gasolina') || d.includes('combustible')) return isVentas ? 'Gasolina Ventas' : 'Gasolina'
  if (d.includes('caseta') || d.includes('autopista')) return isVentas ? 'Casetas Ventas' : 'Casetas'
  if (d.includes('renta') && (d.includes('auto') || d.includes('veh'))) return 'Renta Auto'
  if (d.includes('renta') && (d.includes('oficina') || d.includes('local'))) return 'Renta Oficina'
  if (d.includes('estacionamiento') || d.includes('parking')) return isVentas ? 'Estacionamiento Ventas' : 'Estacionamiento'
  if (d.includes('celular') || d.includes('telefon') || d.includes('telcel')) return 'Celular'
  if (d.includes('herramienta') || d.includes('ferreteria')) return isVentas ? 'Herramientas Ventas' : 'Herramientas'
  if (d.includes('software') || d.includes('licencia') || d.includes('suscripci')) return 'IT & SW (Software/Sistemas)'
  if (d.includes('marketing') || d.includes('publicidad')) return 'Marketing'
  if (d.includes('uniforme')) return 'Uniformes'
  if (d.includes('envio') || d.includes('paquete') || d.includes('flete') || d.includes('dhl') || d.includes('fedex')) return 'Envíos'
  if (d.includes('mantenimiento')) return 'Manto Auto (Mantenimiento)'
  return isVentas ? 'Gastos Rep (Representación)' : 'Consumo'
}

/* Roster shown by the first-run collaborator selector modal.
   Grouped into 4 categorías; the modal filters by name OR categoría. */
const COLABORADORES = [
  { categoria: 'Admin', nombre: 'Leticia Solis' },
  { categoria: 'Admin', nombre: 'Victor Aceves' },
  { categoria: 'Admin', nombre: 'José Luis Falcón' },
  { categoria: 'Admin', nombre: 'Marco Valencia' },
  { categoria: 'Admin', nombre: 'Daniel Covarrubias' },
  { categoria: 'Socio', nombre: 'Edie Haro' },
  { categoria: 'Socio', nombre: 'David Delgado' },
  { categoria: 'Socio', nombre: 'Isaias Valencia' },
  { categoria: 'Socio', nombre: 'Sigifredo Olivas' },
  { categoria: 'Socio', nombre: 'Alejandro Olivar' },
  { categoria: 'Socio', nombre: 'Rosy Corral' },
  { categoria: 'Servicio', nombre: 'Heriberto Chacón' },
  { categoria: 'Servicio', nombre: 'Eduardo Carranco' },
  { categoria: 'Servicio', nombre: 'Daniel Gutierrez' },
  { categoria: 'Servicio', nombre: 'James Tisoto' },
  { categoria: 'Servicio', nombre: 'Misael Cruz' },
  { categoria: 'Servicio', nombre: 'Benjamin Favela' },
  { categoria: 'Servicio', nombre: 'Viviana Perez' },
  { categoria: 'Servicio', nombre: 'Juan Francisco Cuellar' },
  { categoria: 'Servicio', nombre: 'Juan Carlos Virgen' },
  { categoria: 'Servicio', nombre: 'Paola Gutierrez' },
  { categoria: 'Servicio', nombre: 'David de Jesus Delgado' },
  { categoria: 'Servicio', nombre: 'Omar Monclova' },
  { categoria: 'Servicio', nombre: 'Antonio Uribe' },
  { categoria: 'Servicio', nombre: 'Natividad Garcia' },
  { categoria: 'Servicio', nombre: 'Raydel Baltazar' },
  { categoria: 'Servicio', nombre: 'Miguel Castillo' },
  { categoria: 'Servicio', nombre: 'David Lopez' },
  { categoria: 'Servicio', nombre: 'David Castillo' },
  { categoria: 'Servicio', nombre: 'Dario Lopez' },
  { categoria: 'Servicio', nombre: 'Moises Padilla' },
  { categoria: 'Servicio', nombre: 'Emmanuel Haro' },
  { categoria: 'Servicio', nombre: 'Ricardo Pacheco Glez.' },
  { categoria: 'Ventas', nombre: 'Emmanuel Navarro' },
  { categoria: 'Ventas', nombre: 'Carlos Ponce' },
  { categoria: 'Ventas', nombre: 'Gemma Gonzalez' },
  { categoria: 'Ventas', nombre: 'Armando Torres' },
  { categoria: 'Ventas', nombre: 'Juan Carlos Santoyo' },
  { categoria: 'Ventas', nombre: 'Ricardo Pacheco' },
  { categoria: 'Ventas', nombre: 'Mariana Gonzalez' },
  { categoria: 'Ventas', nombre: 'Cynthia Diaz' },
  { categoria: 'Ventas', nombre: 'Julio Torres' },
  { categoria: 'Ventas', nombre: 'Mauricio Rodriguez' },
  { categoria: 'Ventas', nombre: 'Cindy Montaño' },
  { categoria: 'Ventas', nombre: 'Hector Duarte' },
  { categoria: 'Ventas', nombre: 'Juan Sotomayor' },
  { categoria: 'Ventas', nombre: 'Marco Alvarado' },
  { categoria: 'Ventas', nombre: 'Miranda Navarro' },
  { categoria: 'Ventas', nombre: 'Marco Sanchez' },
]

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
  { key: 'montoUSD',          label: 'Monto USD',    width: 110, sortable: true,  type: 'number' },
  { key: 'tipoCambio',        label: 'T/C',          width: 80,  sortable: true,  type: 'number' },
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

// Read a File/Blob as raw base64 (no data-URL prefix). Used to ship
// images and PDFs to the /api/ocr-ticket endpoint.
const fileToBase64 = (file) => new Promise((resolve, reject) => {
  const reader = new FileReader()
  reader.onload = () => {
    const result = reader.result || ''
    const idx = String(result).indexOf(',')
    resolve(idx >= 0 ? String(result).slice(idx + 1) : String(result))
  }
  reader.onerror = reject
  reader.readAsDataURL(file)
})

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

function parseCFDI(xmlText, xmlFile, pdfFiles, colaborador) {
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
    tipo: autoDetectTipo(descripcionFirstLine, colaborador?.categoria),
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
    validado: false,
    montoUSD: 0,
    tipoCambio: 0,
    moneda: 'MXN',
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

function GastoRow({ g, upd, openPDF, onDelete, tiposList }) {
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
    <tr className={[g.hizoMatch && 'row-match', g.isNew && 'row-new'].filter(Boolean).join(' ')}>
      {/* Checkbox */}
      <td className="td-chk">
        <input
          type="checkbox"
          checked={g.validado || false}
          onChange={e => upd('validado', e.target.checked)}
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
      <td>
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

      {/* Proveedor — sticky column (pinned left so it stays visible when scrolling) */}
      <td className="td-proveedor">
        <input className="cell-in is-bold" value={g.proveedor} onChange={e => upd('proveedor', e.target.value)} />
        {g.montoUSD > 0 && <span className="badge-usd">USD</span>}
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
          {tiposList.map(t => (
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

      {/* Monto USD — editing it auto-derives Tipo de Cambio from totalCFDI. */}
      <td>
        <input
          type="number"
          className="cell-in"
          value={g.montoUSD || ''}
          placeholder="0.00"
          step="0.01"
          onChange={e => {
            const usd = parseFloat(e.target.value) || 0
            const tc = usd > 0 && g.totalCFDI > 0
              ? +(g.totalCFDI / usd).toFixed(2)
              : g.tipoCambio
            upd('montoUSD', usd)
            upd('tipoCambio', tc)
          }}
        />
      </td>

      {/* Tipo de Cambio — manual override (or auto-set when Monto USD changes). */}
      <td>
        <input
          type="number"
          className="cell-in"
          value={g.tipoCambio || ''}
          placeholder="0.00"
          step="0.01"
          onChange={e => upd('tipoCambio', parseFloat(e.target.value) || 0)}
        />
      </td>

      {/* Eliminar — sticky-right action column; button only visible on row hover */}
      <td className="td-delete">
        <button
          className="btn-delete-row"
          onClick={onDelete}
          title="Eliminar registro"
        >
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
            <polyline points="3 6 5 6 21 6"/>
            <path d="M19 6l-1 14a2 2 0 01-2 2H8a2 2 0 01-2-2L5 6"/>
            <path d="M10 11v6M14 11v6"/>
            <path d="M9 6V4h6v2"/>
          </svg>
        </button>
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
   COMPONENTE: SELECTOR DE COLABORADOR (modal de bienvenida)
═══════════════════════════════════════════════════ */

function ColaboradorModal({ onSelect }) {
  const [search, setSearch] = useState('')

  const filtered = COLABORADORES.filter(c =>
    c.nombre.toLowerCase().includes(search.toLowerCase()) ||
    c.categoria.toLowerCase().includes(search.toLowerCase())
  )

  const categorias = ['Admin', 'Socio', 'Servicio', 'Ventas']

  return (
    <div className="colab-overlay">
      <div className="colab-modal">
        <div className="colab-modal-header">
          <img src="/logo.png" alt="SMTO" style={{ height: 44, marginBottom: 12 }} />
          <h2 className="colab-title">¿Quién realiza este reporte?</h2>
          <p className="colab-subtitle">Selecciona tu nombre para continuar</p>
        </div>

        <div className="colab-search-wrap">
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" className="colab-search-icon">
            <circle cx="11" cy="11" r="8"/><path d="m21 21-4.35-4.35"/>
          </svg>
          <input
            className="colab-search"
            placeholder="Buscar colaborador..."
            value={search}
            onChange={e => setSearch(e.target.value)}
            autoFocus
          />
        </div>

        <div className="colab-list">
          {categorias.map(cat => {
            const items = filtered.filter(c => c.categoria === cat)
            if (items.length === 0) return null
            return (
              <div key={cat} className="colab-group">
                <div className="colab-group-label">{cat}</div>
                {items.map(c => (
                  <button
                    key={c.nombre}
                    className="colab-item"
                    onClick={() => onSelect(c)}
                  >
                    <span className="colab-avatar">{c.nombre.charAt(0)}</span>
                    <span className="colab-nombre">{c.nombre}</span>
                    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" className="colab-arrow">
                      <path d="m9 18 6-6-6-6"/>
                    </svg>
                  </button>
                ))}
              </div>
            )
          })}
        </div>
      </div>
    </div>
  )
}

/* ═══════════════════════════════════════════════════
   COMPONENTE: MODAL DE IMPORTACIÓN COMPLETADA
═══════════════════════════════════════════════════ */

// Premium import success modal. Shares the .cm-* glass shell with the
// other premium modals (orb, cta, overlay). The body adds a "+N" hero,
// a total-registros stat card, and a 100% progress bar with a brand glow.
function ImportSuccessModal({ data, onClose }) {
  const cAdded = useCountUp(data.added)
  const cTotal = useCountUp(data.total)

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
          <h2 className="cm-title">Importación completada</h2>
          <p className="cm-subtitle">
            {data.mode === 'replace'
              ? 'Reporte importado correctamente'
              : 'Reporte actualizado correctamente'}
          </p>
        </div>

        {/* Hero "+N" */}
        <motion.div
          className="ism-hero"
          initial={{ opacity: 0, y: 14 }}
          animate={{ opacity: 1, y: 0 }}
          transition={{ delay: 0.25, type: 'spring', stiffness: 240, damping: 22 }}
        >
          <div className="ism-plus">
            <span className="ism-plus-sign">+</span>
            <span className="ism-plus-num">{Math.round(cAdded)}</span>
          </div>
          <div className="ism-hero-label">facturas importadas</div>
        </motion.div>

        {/* Total registros card */}
        <motion.div
          className="ism-total"
          initial={{ opacity: 0, y: 10 }}
          animate={{ opacity: 1, y: 0 }}
          transition={{ delay: 0.35, type: 'spring', stiffness: 260, damping: 24 }}
          whileHover={{ y: -2 }}
        >
          <div className="ism-total-label">Registros totales</div>
          <div className="ism-total-value">{Math.round(cTotal)}</div>
        </motion.div>

        {/* Progress bar — fills 0 → 100% with a slight delay, glowing green→cyan */}
        <motion.div
          className="ism-progress"
          initial={{ opacity: 0 }}
          animate={{ opacity: 1 }}
          transition={{ delay: 0.45 }}
        >
          <div className="ism-progress-track">
            <motion.div
              className="ism-progress-fill"
              initial={{ width: 0 }}
              animate={{ width: '100%' }}
              transition={{ delay: 0.5, duration: 1, ease: 'easeOut' }}
            />
          </div>
          <div className="ism-progress-label">
            <span>Procesado</span>
            <span>100%</span>
          </div>
        </motion.div>

        {/* CTA */}
        <motion.button
          className="cm-cta"
          onClick={onClose}
          whileHover={{ y: -1 }}
          whileTap={{ scale: 0.98 }}
          initial={{ opacity: 0, y: 8 }}
          animate={{ opacity: 1, y: 0 }}
          transition={{ delay: 0.65 }}
        >
          Continuar
        </motion.button>
      </motion.div>
    </motion.div>
  )
}

/* ═══════════════════════════════════════════════════
   COMPONENTE: MODAL DE ARCHIVOS PROCESADOS (drag-and-drop)
═══════════════════════════════════════════════════ */

// Drag/drop success modal. Same .cm-* glass shell + KpiCard grid as the
// conciliacion modal, plus the progress bar from the import modal —
// three count-up stats (Nuevas / Actualizadas / PDFs) under the orb.
function DropSuccessModal({ data, onClose }) {
  const cAdded   = useCountUp(data.added)
  const cUpdated = useCountUp(data.updated)
  const cPdfs    = useCountUp(data.pdfs)

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
          <h2 className="cm-title">Archivos procesados</h2>
          <p className="cm-subtitle">Tu reporte está actualizado</p>
        </div>

        {/* KPI grid — 3 count-up stats with brand-tinted icons */}
        <motion.div className="cm-kpi-grid" variants={stagger} initial="hidden" animate="show">
          <KpiCard variants={item} accent="green"  Icon={Plus}     value={cAdded} />
          <KpiCard variants={item} accent="blue"   Icon={Link2}    value={cUpdated} />
          <KpiCard variants={item} accent="purple" Icon={FileText} value={cPdfs} />
        </motion.div>
        <motion.div
          className="cm-kpi-labels"
          initial={{ opacity: 0 }}
          animate={{ opacity: 1 }}
          transition={{ delay: 0.4 }}
        >
          <span>Nuevas</span>
          <span>Actualizadas</span>
          <span>PDFs recibidos</span>
        </motion.div>

        {/* 100% progress bar — same shape as the import modal */}
        <motion.div
          className="ism-progress"
          initial={{ opacity: 0 }}
          animate={{ opacity: 1 }}
          transition={{ delay: 0.5 }}
        >
          <div className="ism-progress-track">
            <motion.div
              className="ism-progress-fill"
              initial={{ width: 0 }}
              animate={{ width: '100%' }}
              transition={{ delay: 0.55, duration: 1, ease: 'easeOut' }}
            />
          </div>
          <div className="ism-progress-label">
            <span>Procesado</span>
            <span>100%</span>
          </div>
        </motion.div>

        {/* CTA */}
        <motion.button
          className="cm-cta"
          onClick={onClose}
          whileHover={{ y: -1 }}
          whileTap={{ scale: 0.98 }}
          initial={{ opacity: 0, y: 8 }}
          animate={{ opacity: 1, y: 0 }}
          transition={{ delay: 0.7 }}
        >
          Continuar
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
  const [colaborador,   setColaborador]   = useState(null)
  const [showColabModal, setShowColabModal] = useState(true)
  const [colabSearch,   setColabSearch]   = useState('')
  const [importSuccess, setImportSuccess] = useState(false)
  const [importSummary, setImportSummary] = useState(null)
  const [dropSummary,   setDropSummary]   = useState(null)
  const [ocrLoading,    setOcrLoading]    = useState(false)
  const [loading,       setLoading]       = useState(false)
  const [isDraggingOver, setIsDraggingOver] = useState(false)
  // Index-based fixed pixel widths — order matches COLUMNS positions:
  // [0] checkbox, [1] estado, [2] fecha factura, [3] fecha cobro,
  // [4] factura, [5] proveedor, [6] concepto, [7] tipo, [8] subtotal,
  // [9] iva, [10] isr/ish/ieps, [11] ret.isr, [12] ret.iva, [13] reten,
  // [14] total fac, [15] forma pago, [16] prop%, [17] prop$, [18] total final
  const [colWidths, setColWidths] = useState([40, 110, 115, 120, 120, 260, 140, 100, 120, 110, 135, 110, 110, 110, 125, 160, 95, 105, 130, 110, 80])
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

  // TIPO dropdown list — depends on the selected colaborador's categoría.
  // Ventas/Socio get the sales-flavored list; Admin/Servicio get the
  // operational list; null (modal still open) falls back to a merged
  // sorted union so existing rows can still render their saved tipo.
  const tiposList = useMemo(() => getTiposForColaborador(colaborador), [colaborador])

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
        const g = parseCFDI(text, f, pdfs, colaborador)
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

  // ── Drag/drop XML+PDF onto the table to merge into the current report ──
  // Differs from processFiles (folder picker) in that it MERGES instead of
  // replacing: existing rows matched by RFC+noFactura are refreshed with
  // the new XML/PDF attachment; truly new rows are appended.
  const handleDragOver = (e) => {
    e.preventDefault()
    e.stopPropagation()
    setIsDraggingOver(true)
  }

  const handleDragLeave = (e) => {
    e.stopPropagation()
    // Only clear when the cursor leaves the container entirely, not when
    // it crosses into a child element.
    if (!e.currentTarget.contains(e.relatedTarget)) {
      setIsDraggingOver(false)
    }
  }

  // Posts a base64 image/PDF to /api/ocr-ticket and shapes the parsed
  // receipt into a full gasto row. Reads `colaborador` from closure so
  // autoDetectTipo gets routed to the right category list. USD receipts
  // land montoUSD with importe/totalCFDI = 0 (the user fills the MXN side
  // from their card statement later).
  const extractReceiptData = async (base64, mediaType, fileName) => {
    const response = await fetch('/api/ocr-ticket', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ base64, mediaType }),
    })
    if (!response.ok) {
      const err = await response.json().catch(() => ({}))
      throw new Error(err.error || `OCR ${response.status}`)
    }
    const parsed = await response.json()
    const uuid = crypto.randomUUID()
    const isUSD = parsed.moneda === 'USD'
    const today = new Date().toISOString().slice(0, 10)
    const subtotal = Number(parsed.subtotal) || 0
    const iva = Number(parsed.iva) || 0
    const total = Number(parsed.total) || 0
    const propina = Number(parsed.propina) || 0

    return {
      id: genId(),
      rfc: '',
      proveedor: parsed.proveedor || '',
      noFactura: parsed.folio ? String(parsed.folio) : ('TKT-' + uuid.slice(0, 6).toUpperCase()),
      fechaFac: parsed.fecha || today,
      concepto: parsed.concepto || '',
      tipo: autoDetectTipo(parsed.concepto || '', colaborador?.categoria),
      importe: isUSD ? 0 : subtotal,
      iva: isUSD ? 0 : iva,
      isrTrasladado: 0,
      retencionISR: 0,
      retencionIVA: 0,
      retenciones: 0,
      totalCFDI: isUSD ? 0 : total,
      propinaPorcentaje: 0,
      montoPropina: isUSD ? 0 : propina,
      fechaCobro: parsed.fecha || today,
      formaPago: parsed.formaPago || '04',
      uuid,
      tienePDF: false,
      pdfFile: null,
      xmlFile: null,
      hizoMatch: false,
      validado: false,
      montoUSD: isUSD ? total : 0,
      tipoCambio: 0,
      moneda: isUSD ? 'USD' : 'MXN',
    }
  }

  const handleDrop = async (e) => {
    e.preventDefault()
    e.stopPropagation()
    setIsDraggingOver(false)

    const files = Array.from(e.dataTransfer.files)
    if (!files.length) return

    const xmlFiles   = files.filter(f => f.name.toLowerCase().endsWith('.xml'))
    const pdfFiles   = files.filter(f => f.name.toLowerCase().endsWith('.pdf'))
    const imageFiles = files.filter(f => /\.(jpg|jpeg|png|webp)$/i.test(f.name.toLowerCase()))

    const allNewGastos = []

    // STEP 1 — parse XMLs locally (always free). parseCFDI takes the full
    // signature so xmlFile + auto-PDF-link + auto-tipo all work.
    for (const file of xmlFiles) {
      try {
        const text = await file.text()
        const gasto = parseCFDI(text, file, pdfFiles, colaborador)
        if (gasto) {
          gasto.isNew = true
          allNewGastos.push(gasto)
        }
      } catch (err) {
        console.warn('XML parse error:', file.name, err)
      }
    }

    // STEP 2 — figure out which dropped PDFs are still loose. parseCFDI's
    // built-in matcher already attached the obvious ones (g.pdfFile is set);
    // try a second pass against folio/RFC for the rest.
    const unmatchedPDFs = []
    for (const pdfFile of pdfFiles) {
      const pdfBase = pdfFile.name.replace(/\.pdf$/i, '').toLowerCase()

      // Already attached by parseCFDI?
      if (allNewGastos.some(g => g.pdfFile && g.pdfFile.name.toLowerCase() === pdfFile.name.toLowerCase())) {
        continue
      }

      const matched = allNewGastos.find(g => {
        const xmlBase = (g.xmlFile?.name || '').replace(/\.xml$/i, '').toLowerCase()
        const folio   = (g.noFactura || '').toLowerCase()
        const rfc     = (g.rfc || '').toLowerCase()
        return (
          pdfBase === xmlBase ||
          (folio   && pdfBase.includes(folio)) ||
          (rfc     && pdfBase.includes(rfc))   ||
          (xmlBase && xmlBase.includes(pdfBase))
        )
      })
      if (matched) {
        matched.pdfFile  = pdfFile
        matched.tienePDF = true
      } else {
        unmatchedPDFs.push(pdfFile)
      }
    }

    // STEP 3 — OCR pass for unmatched PDFs + dropped images. Asks before
    // spending money. Errors per file route through the plain alerta modal.
    const ocrFiles = [...unmatchedPDFs, ...imageFiles]
    if (ocrFiles.length > 0) {
      const userConfirmed = window.confirm(
        `Se detectaron ${ocrFiles.length} archivo(s) sin XML:\n\n` +
        ocrFiles.map(f => `• ${f.name}`).join('\n') +
        `\n\n⚡ ¿Procesarlos con OCR (IA)?\n` +
        `Costo estimado: ~$${(ocrFiles.length * 0.01).toFixed(2)} USD`
      )

      if (userConfirmed) {
        setOcrLoading(true)
        for (const file of ocrFiles) {
          try {
            const base64 = await fileToBase64(file)
            const ext = file.name.split('.').pop().toLowerCase()
            const mediaType = ext === 'pdf'
              ? 'application/pdf'
              : `image/${ext === 'jpg' ? 'jpeg' : ext}`

            const gasto = await extractReceiptData(base64, mediaType, file.name)
            if (gasto) {
              // Keep the source PDF attached so it rides the ZIP export.
              if (mediaType === 'application/pdf') {
                gasto.pdfFile = file
                gasto.tienePDF = true
              }
              gasto.isNew = true
              allNewGastos.push(gasto)
            }
          } catch (err) {
            console.warn('OCR error:', file.name, err)
            setAlerta(`Error procesando ${file.name}: ${err.message}`)
          }
        }
        setOcrLoading(false)
      }
    }

    // STEP 4 — bail with the right message. If nothing parseable was dropped
    // at all, show "Solo se aceptan...". If OCR was declined or every file
    // failed, exit silently (the user already knows).
    if (allNewGastos.length === 0) {
      if (xmlFiles.length === 0 && ocrFiles.length === 0) {
        setAlerta('Solo se aceptan archivos XML, PDF, JPG, PNG o WEBP.')
      }
      return
    }

    // STEP 5 — single setLista commit, then surface the premium drop modal.
    setLista(prev => {
      const merged = [...prev]
      const existingKeys = new Set(prev.map(g => `${g.rfc}|${g.noFactura}`))
      let added = 0
      let updated = 0

      for (const newG of allNewGastos) {
        const key = `${newG.rfc}|${newG.noFactura}`
        if (existingKeys.has(key)) {
          const idx = merged.findIndex(g => g.rfc === newG.rfc && g.noFactura === newG.noFactura)
          if (idx !== -1) {
            // Refresh existing row with whatever the new payload carries
            // (xmlFile, pdfFile, fresher fields), keep isNew for the flash.
            merged[idx] = { ...merged[idx], ...newG, isNew: true }
            updated++
          }
        } else {
          merged.push(newG)
          existingKeys.add(key)
          added++
        }
      }

      // Premium glass result modal — only opens when something actually changed.
      if (added > 0 || updated > 0 || pdfFiles.length > 0) {
        setTimeout(() => setDropSummary({
          added,
          updated,
          pdfs: pdfFiles.length,
        }), 100)
      }

      return merged
    })

    // Clear isNew flags after the row entrance animation finishes so
    // subsequent renders don't replay it.
    setTimeout(() => {
      setLista(l => l.map(g => g.isNew ? { ...g, isNew: false } : g))
    }, 1500)
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
      tienePDF: false, pdfFile: null, xmlFile: null, hizoMatch: false, validado: false,
      montoUSD: 0, tipoCambio: 0, moneda: 'MXN',
    }])
  }

  // ── Copiar a portapapeles (TSV para Excel) ──
  const copiar = () => {
    const hdr = 'RFC PROVEEDOR\tPROVEEDOR\tNO. DE FACTURA\tFECHA FAC.\tCONCEPTO\tIMPORTE (MXP)\tIVA\tISR\tRET. ISR\tRET. IVA\tRET/ ISR IVA\tTOTAL CFDI\tGastos en USD\tTipo de Cambio\tTotal Checking\tFORMA DE PAGO\tFECHA DE COBRO\n'
    const rows = lista.flatMap(g => {
      const fac   = formatDateDisplay(g.fechaFac)
      const cobro = formatDateDisplay(g.fechaCobro) || 'Pendiente'
      const r  = `${g.rfc}\t${g.proveedor.replace(/\t/g,' ')}\t${g.noFactura}\t${fac}\t${g.concepto.replace(/\t/g,' ')}\t${g.importe.toFixed(2)}\t${g.iva.toFixed(2)}\t${(g.isrTrasladado||0).toFixed(2)}\t${(g.retencionISR||0).toFixed(2)}\t${(g.retencionIVA||0).toFixed(2)}\t${g.retenciones.toFixed(2)}\t${g.totalCFDI.toFixed(2)}\t${(g.montoUSD||0).toFixed(2)}\t${(g.tipoCambio||0).toFixed(2)}\t\t${g.formaPago}\t${cobro}\n`
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
      csv += `${g.rfc},${g.proveedor.replace(/,/g,' ')},${g.noFactura},${fac},${g.concepto.replace(/,/g,' ')},${g.importe.toFixed(2)},${g.iva.toFixed(2)},${(g.isrTrasladado||0).toFixed(2)},${(g.retencionISR||0).toFixed(2)},${(g.retencionIVA||0).toFixed(2)},${g.retenciones.toFixed(2)},${g.totalCFDI.toFixed(2)},${(g.montoUSD||0).toFixed(2)},${(g.tipoCambio||0).toFixed(2)},,${g.formaPago},${cobro}\n`
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

  // ── Importar Excel previamente exportado ──
  // Reads a Reporte_Gastos_SMTO.xlsx and reconstructs the gastos list.
  // The export writes data starting at row 10 (table header on row 9); we
  // walk rows downward until we hit an empty RFC or the "TOTAL CUENTA"
  // sentinel. The Excel only carries the rolled-up RETENCIÓN column, so
  // retencionISR/IVA breakdown defaults to 0 — caller can edit manually.
  const handleImportExcel = async (e) => {
    const file = e.target.files[0]
    if (!file) return
    e.target.value = ''

    try {
      const arrayBuffer = await file.arrayBuffer()
      const wb = XLSX.read(arrayBuffer, { type: 'array', cellStyles: true })
      const ws = wb.Sheets[wb.SheetNames[0]]

      // Data starts at row 10 (0-indexed: r=9). Columns:
      // B=RFC, C=PROVEEDOR, D=TIPO, E=FACTURA, F=F.FACTURA, G=F.COBRO,
      // H=CONCEPTO, I=IMPORTE, J=IVA, K=RETENCIÓN, L=TOTAL, M=FORMA PAGO
      const gastos = []
      let row = 10

      while (true) {
        const rfc = ws[XLSX.utils.encode_cell({ r: row - 1, c: 1 })]?.v
        if (!rfc || rfc === 'TOTAL CUENTA') break

        const getVal = (col) => ws[XLSX.utils.encode_cell({ r: row - 1, c: col })]?.v ?? ''

        // Forma de pago: cell carries the "04 - Tarjeta de Crédito" label,
        // so strip back to the leading two-digit code.
        const formaLabel = String(getVal(12))
        const formaCode = formaLabel.startsWith('04') ? '04'
          : formaLabel.startsWith('03') ? '03'
          : formaLabel.startsWith('02') ? '02'
          : formaLabel.startsWith('01') ? '01' : '04'

        // Dates round-trip as MM-DD-YY strings in the export; convert back
        // to YYYY-MM-DD for internal use.
        const parseDate = (d) => {
          if (!d) return ''
          const s = String(d)
          if (s.includes('-') && s.length === 8) {
            const [mm, dd, yy] = s.split('-')
            return `20${yy}-${mm}-${dd}`
          }
          return s
        }

        // Numeric fields. The export writes column L as an Excel formula
        // `=I+J-K`, and SheetJS returns 0/undefined for formula cells when
        // there is no cached evaluated value — so fall back to recomputing
        // the total from I+J-K if L came back empty.
        const importe     = Number(getVal(8))  || 0
        const iva         = Number(getVal(9))  || 0
        const retenciones = Number(getVal(10)) || 0
        const lRaw        = Number(getVal(11)) || 0
        const totalCFDI   = lRaw || (importe + iva - retenciones)

        gastos.push({
          // id is required by React (table key) and the delete handler;
          // tienePDF/pdfFile/xmlFile/hizoMatch/isrTrasladado keep the row
          // shape identical to parseCFDI-produced rows so every cell
          // renders correctly.
          id: genId(),
          rfc: String(rfc || ''),
          proveedor: String(getVal(2) || ''),
          tipo: String(getVal(3) || ''),
          noFactura: String(getVal(4) || ''),
          fechaFac: parseDate(getVal(5)),
          fechaCobro: parseDate(getVal(6)),
          concepto: String(getVal(7) || ''),
          importe,
          iva,
          isrTrasladado: 0,
          retencionISR: 0,
          retencionIVA: 0,
          retenciones,
          totalCFDI,
          formaPago: formaCode,
          propinaPorcentaje: 0,
          montoPropina: 0,
          uuid: crypto.randomUUID(),
          tienePDF: false,
          pdfFile: null,
          xmlFile: null,
          hizoMatch: false,
          validado: false,
          // New USD columns sit at N (idx 13) + O (idx 14) in the exported
          // xlsx; round-trip them here so re-importing keeps the values.
          montoUSD: Number(getVal(13)) || 0,
          tipoCambio: Number(getVal(14)) || 0,
          moneda: (Number(getVal(13)) || 0) > 0 ? 'USD' : 'MXN',
        })
        row++
      }

      if (gastos.length === 0) {
        setAlerta('No se encontraron registros válidos en el archivo Excel.')
        return
      }

      // Ask user: append or replace.
      const action = window.confirm(
        `Se encontraron ${gastos.length} registros en el Excel.\n\n` +
        `¿Deseas AGREGAR a los ${lista.length} registros actuales?\n\n` +
        `OK = Agregar al reporte actual\n` +
        `Cancelar = Reemplazar todo`
      )

      // Tag every imported row as isNew so each one picks up the green-flash
      // row-fade-in animation; flag is cleared after the animation settles.
      const tagged = gastos.map(g => ({ ...g, isNew: true }))

      if (action) {
        setLista(prev => [...prev, ...tagged])
      } else {
        setLista(tagged)
      }

      // Premium success modal (count-up + progress bar + glow). Replaces the
      // plain-text alerta success line; errors still route through setAlerta.
      setImportSummary({
        added: gastos.length,
        total: action ? lista.length + gastos.length : gastos.length,
        mode: action ? 'append' : 'replace',
      })

      // Flash the toolbar button green for 2.5s and clear the isNew flags
      // after 1s so subsequent renders do not replay the row animation.
      setImportSuccess(true)
      setTimeout(() => setImportSuccess(false), 2500)
      setTimeout(() => {
        setLista(prev => prev.map(g => g.isNew ? { ...g, isNew: false } : g))
      }, 1000)
    } catch (err) {
      setAlerta('Error al importar Excel: ' + err.message)
    }
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
        body: JSON.stringify({ gastos: lista, colaborador: colaborador?.nombre || '' }),
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

      {/* ─── MODAL DE BIENVENIDA: selección de colaborador ─── */}
      {showColabModal && (
        <ColaboradorModal onSelect={c => {
          setColaborador(c)
          setShowColabModal(false)
        }} />
      )}

      {/* ─── CABECERA ─── */}
      <div className="header">
        <div className="header-logo">
          <img src="/logo.png" alt="SMTO" style={{ height: '54px', width: 'auto', objectFit: 'contain' }} />
        </div>
        <div className="header-info">
          <h1 className="header-title">Reporte de Gastos SMTO<span className="version-badge">v7.5</span></h1>
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
          {colaborador && (
            <button className="colab-chip" onClick={() => setShowColabModal(true)} title="Cambiar colaborador">
              <span className="colab-chip-avatar">{colaborador.nombre.charAt(0)}</span>
              <span>{colaborador.nombre}</span>
              <span className="colab-chip-cat">{colaborador.categoria}</span>
            </button>
          )}
        </div>
      </div>

      <div className="divider" />

      {/* ─── BARRA DE ACCIONES ─── */}
      <div className="action-bar">
        <div className="action-group">
          <PremiumButton title="Manual"         icon="＋"  variant="ghost"     onClick={agregarManual} />
          <PremiumButton
            title={importSuccess ? 'Importado' : 'Importar Excel'}
            variant={importSuccess ? 'success' : 'ghost'}
            onClick={() => document.getElementById('import-xlsx-input').click()}
            icon={importSuccess ? (
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
                <polyline points="20 6 9 17 4 12"/>
              </svg>
            ) : (
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4"/>
                <polyline points="17 8 12 3 7 8"/>
                <line x1="12" y1="3" x2="12" y2="15"/>
              </svg>
            )}
          />
          <input
            id="import-xlsx-input"
            type="file"
            accept=".xlsx"
            style={{ display: 'none' }}
            onChange={handleImportExcel}
          />
          <PremiumButton title="Cargar Carpeta" icon="📂" variant="primary"   onClick={() => folderRef.current?.click()} />
          <PremiumButton title="Validar Banco"  icon="🏦" variant="secondary" onClick={() => bancoRef.current?.click()} />
          {(() => {
            const total = lista.length
            const validados = lista.filter(g => g.validado).length
            const allDone = total > 0 && validados === total
            return total > 0 ? (
              <div className={`validacion-counter${allDone ? ' all-done' : ''}`}>
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
                  <polyline points="20 6 9 17 4 12"/>
                </svg>
                <span>
                  {allDone
                    ? `✓ ${total} facturas validadas`
                    : `${validados} / ${total} validadas`}
                </span>
              </div>
            ) : null
          })()}
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
      <div
        className={`table-wrap table-container${isDraggingOver ? ' drag-over' : ''}`}
        onDragOver={handleDragOver}
        onDragLeave={handleDragLeave}
        onDrop={handleDrop}
      >
        {loading ? (
          <div className="loading-msg">
            <div className="loading-spinner" />
            <div className="loading-text">Procesando facturas XML…</div>
          </div>
        ) : lista.length === 0 ? (
          <div
            className={`onboarding ${isDraggingOver ? 'is-dragging' : ''}`}
            onDragOver={handleDragOver}
            onDragLeave={handleDragLeave}
            onDrop={handleDrop}
          >
            <div className="onboarding-glow" />
            <div className="onboarding-card">
              <div className="onboarding-icon">
                {isDraggingOver ? (
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
                {isDraggingOver
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
            // Sticky-column left offsets — Proveedor pins right after the
            // Estado column. Resizing Estado shifts the pinned Proveedor cell
            // correctly because we read its width from colWidths.
            '--sl-proveedor': `${colWidths[1]}px`,
          }}>
            <thead>
              <tr>
                {COLUMNS.map((col, idx) => (
                  <th
                    key={col.key}
                    className={col.key === 'status' ? 'th-status' : col.key === 'proveedor' ? 'th-proveedor' : undefined}
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
                {/* Eliminar — sticky-right action column, no label, fixed 40px width */}
                <th className="th-delete" style={{ width: 40 }} />
              </tr>
            </thead>
            <tbody>
              {sortedLista.map(g => (
                <GastoRow
                  key={g.id}
                  g={g}
                  upd={(field, val) => update(g.id, field, val)}
                  onDelete={() => setLista(prev => prev.filter(x => x.id !== g.id))}
                  openPDF={openPDF}
                  tiposList={tiposList}
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

      {/* ─── MODAL IMPORTACIÓN COMPLETADA (premium glass) ─── */}
      <AnimatePresence>
        {importSummary && (
          <ImportSuccessModal data={importSummary} onClose={() => setImportSummary(null)} />
        )}
      </AnimatePresence>

      {/* ─── MODAL ARCHIVOS PROCESADOS (drag-and-drop, premium glass) ─── */}
      <AnimatePresence>
        {dropSummary && (
          <DropSuccessModal data={dropSummary} onClose={() => setDropSummary(null)} />
        )}
      </AnimatePresence>

      {/* ─── OCR LOADING OVERLAY ─── */}
      {ocrLoading && (
        <div className="cm-overlay" style={{ pointerEvents: 'auto' }}>
          <div className="loading-msg">
            <div className="loading-spinner" />
            <div className="loading-text">Procesando con OCR…</div>
          </div>
        </div>
      )}
    </div>
  )
}
