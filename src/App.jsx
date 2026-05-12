import { useState, useRef, useMemo } from 'react'
import JSZip from 'jszip'

/* Single source of truth for table columns.
   `getValue` overrides simple `g[key]` lookup (used by Total Final). */
const COLUMNS = [
  { key: 'check',             label: '',             width: 52,  sortable: false },
  { key: 'status',            label: 'Estado',       width: 110, sortable: false },
  { key: 'fechaFac',          label: 'Fecha',        width: 115, sortable: true,  type: 'date'   },
  { key: 'noFactura',         label: 'Factura',      width: 120, sortable: true,  type: 'string' },
  { key: 'proveedor',         label: 'Proveedor',    width: 260, sortable: true,  type: 'string' },
  { key: 'concepto',          label: 'Concepto',     width: 140, sortable: true,  type: 'string' },
  { key: 'importe',           label: 'Subtotal',     width: 120, sortable: true,  type: 'number' },
  { key: 'iva',               label: 'IVA',          width: 110, sortable: true,  type: 'number' },
  { key: 'isrTrasladado',     label: 'ISR/ISH/IEPS', width: 135, sortable: true,  type: 'number' },
  { key: 'retencionISR',      label: 'Ret. ISR',     width: 110, sortable: true,  type: 'number' },
  { key: 'retencionIVA',      label: 'Ret. IVA',     width: 110, sortable: true,  type: 'number' },
  { key: 'retenciones',       label: 'Reten.',       width: 110, sortable: true,  type: 'number' },
  { key: 'totalCFDI',         label: 'Total Fac.',   width: 125, sortable: true,  type: 'number' },
  { key: 'propinaPorcentaje', label: 'Prop. %',      width: 95,  sortable: true,  type: 'number' },
  { key: 'montoPropina',      label: 'Prop. $',      width: 105, sortable: true,  type: 'number' },
  { key: 'totalFinal',        label: 'Total Final',  width: 130, sortable: true,  type: 'number',
    getValue: g => g.totalCFDI + g.montoPropina },
  { key: 'fechaCobro',        label: 'Fecha Cobro',  width: 120, sortable: true,  type: 'string' },
]

/* ═══════════════════════════════════════════════════
   UTILIDADES
═══════════════════════════════════════════════════ */

const genId = () => Math.random().toString(36).slice(2, 11)

function clasificarGasto(proveedor, concepto) {
  const t = `${proveedor} ${concepto}`.toUpperCase()
  const has = (...kws) => kws.some(k => t.includes(k))
  if (has('VUELO','AVIACION','CONCESIONARIA VUELA','AEROENLACES','AEROMEXICO','VIVA AEROBUS','AEREA','AEROPUERTO','BOLETO')) return 'Vuelo'
  if (has('HOTEL','HOSPEDAJE','HILTON','MARRIOTT','HOLIDAY','CITY EXPRESS','HABITACION','INN')) return 'Hotel'
  if (has('UBER','DIDI','TAXI','PEAJE','CASETA','AUTOPISTA','FONADIN','CAPUFE','ESTACIONAMIENTO')) return 'Transporte'
  if (has('HOME DEPOT','FERRETERIA','MATERIAL','TRUPER','GRAINGER','HERRAMIENTA')) return 'Herramienta'
  if (has('GASNGO','COMBUSTIBLE','GASOLINA','MAGNA','PREMIUM','DIESEL')) return 'Combustible'
  return 'Consumo'
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

  // Concepto
  let concepto = 'Consumo'
  if (conceptoEl) {
    const desc = ga(conceptoEl, 'Descripcion', 'descripcion') || ''
    const clean = []
    for (const p of desc.split(/\s+/).filter(Boolean)) {
      if (p.length === 8 && /^\d+$/.test(p)) continue
      if (p.toUpperCase() === 'NO' && clean.length === 0) continue
      clean.push(p)
    }
    concepto = clean.slice(0, 4).join(' ') || 'Consumo'
  }

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
  let conceptoClasif = clasificarGasto(proveedor, concepto)

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

  return {
    id: genId(),
    rfc,
    proveedor,
    noFactura:  ga(comp, 'Folio', 'folio') || 'SN',
    fechaFac:  (ga(comp, 'Fecha', 'fecha') || '').slice(0, 10),
    concepto:   conceptoClasif,
    importe,
    iva,
    isrTrasladado,
    retencionISR,
    retencionIVA,
    retenciones,
    totalCFDI,
    propinaPorcentaje: 0,
    montoPropina: 0,
    fechaCobro: '',
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
  // Máscara visual: YYYY-MM-DD ↔ MM-DD-YYYY
  const dateDisplay = (() => {
    const p = g.fechaFac.split('-')
    return p.length === 3 && p[0].length === 4
      ? `${p[1]}-${p[2]}-${p[0]}`
      : g.fechaFac
  })()

  const onDateChange = v => {
    const p = v.split('-')
    upd('fechaFac', p.length === 3 && p[2].length === 4
      ? `${p[2]}-${p[0]}-${p[1]}`
      : v)
  }

  const NumCell = ({ field, prefix, suffix, format = true }) => {
    const v = g[field]
    // Money fields always render as toFixed(2); 0 renders as '' to keep cells clean.
    const display = v ? (format ? Number(v).toFixed(2) : v) : ''
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

      {/* Fecha */}
      <td className="td-fecha">
        <input className="cell-in" value={dateDisplay} onChange={e => onDateChange(e.target.value)} />
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
      <td className="td-proveedor">
        <input className="cell-in is-bold" value={g.proveedor} onChange={e => upd('proveedor', e.target.value)} />
      </td>

      {/* Concepto */}
      <td>
        <input className="cell-in is-dim" value={g.concepto} onChange={e => upd('concepto', e.target.value)} />
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

      {/* Propina % — kept unformatted so the user can type fractional percentages */}
      <td><NumCell field="propinaPorcentaje" suffix="%" format={false} /></td>

      {/* Propina $ */}
      <td><NumCell field="montoPropina" prefix="$" /></td>

      {/* Total Final */}
      <td>
        <span className={`total-val${g.hizoMatch ? ' is-blue' : ''}`}>
          ${(g.totalCFDI + g.montoPropina).toFixed(2)}
        </span>
      </td>

      {/* Fecha Cobro — populated by validarBanco, but editable */}
      <td>
        <input
          className={`cell-in is-cobro${g.fechaCobro ? ' is-filled' : ''}`}
          value={g.fechaCobro || ''}
          placeholder="Pendiente"
          onChange={e => upd('fechaCobro', e.target.value)}
        />
      </td>
    </tr>
  )
}

/* ═══════════════════════════════════════════════════
   APP PRINCIPAL
═══════════════════════════════════════════════════ */

export default function App() {
  const [lista,         setLista]         = useState([])
  const [carpetaNombre, setCarpetaNombre] = useState('Ninguna carpeta seleccionada')
  const [alerta,        setAlerta]        = useState(null)
  const [loading,       setLoading]       = useState(false)
  const [isDragging,    setIsDragging]    = useState(false)
  const [colWidths,     setColWidths]     = useState(() =>
    Object.fromEntries(COLUMNS.map(c => [c.key, c.width]))
  )
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
      count:            lista.length,
    }
  }, [lista])
  const fmtMoney = n => `$${n.toLocaleString('es-MX', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`

  // ── Sort ──
  const sortedLista = useMemo(() => {
    if (!sort.field) return lista
    const col = COLUMNS.find(c => c.key === sort.field)
    if (!col) return lista
    const get = col.getValue || (g => g[sort.field])
    return [...lista].sort((a, b) => {
      const va = get(a), vb = get(b)
      let cmp
      if (col.type === 'number') cmp = (parseFloat(va) || 0) - (parseFloat(vb) || 0)
      else                       cmp = String(va ?? '').toLowerCase().localeCompare(String(vb ?? '').toLowerCase())
      return sort.dir === 'asc' ? cmp : -cmp
    })
  }, [lista, sort])

  const toggleSort = field => setSort(s =>
    s.field === field ? { field, dir: s.dir === 'asc' ? 'desc' : 'asc' } : { field, dir: 'asc' }
  )

  // ── Column resize (drag right edge of th) ──
  const startResize = (key, e) => {
    e.preventDefault(); e.stopPropagation()
    const startX = e.clientX
    const startW = colWidths[key]
    const onMove = ev => setColWidths(w => ({ ...w, [key]: Math.max(30, startW + ev.clientX - startX) }))
    const onUp   = () => {
      window.removeEventListener('mousemove', onMove)
      window.removeEventListener('mouseup',   onUp)
    }
    window.addEventListener('mousemove', onMove)
    window.addEventListener('mouseup',   onUp)
  }

  // ── Actualizar campo de un gasto (con recálculo de propina) ──
  const update = (id, field, value) =>
    setLista(prev => prev.map(g => {
      if (g.id !== id) return g
      const u = { ...g, [field]: value }
      if (field === 'totalCFDI' && u.propinaPorcentaje > 0)
        u.montoPropina = value * u.propinaPorcentaje / 100
      if (field === 'propinaPorcentaje')
        u.montoPropina = g.totalCFDI * value / 100
      if (field === 'montoPropina' && g.totalCFDI > 0)
        u.propinaPorcentaje = (value / g.totalCFDI) * 100
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

    let bancoRows = 0, matches = 0, propinas = 0
    const nl = lista.map(g => ({ ...g, hizoMatch: false, fechaCobro: '' }))
    const formatCobro = d => {
      const dd   = String(d.getDate()).padStart(2, '0')
      const mm   = String(d.getMonth() + 1).padStart(2, '0')
      const yyyy = d.getFullYear()
      return `${dd}/${mm}/${yyyy}`
    }

    for (const line of lines) {
      if (!line.trim()) continue
      const cols = parseCSVLine(line, sep)
      let dCSV = null; const amounts = []

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
      bancoRows++; let found = false

      for (const csv of amounts) {
        if (found) break
        // Paso 1: match exacto (±$5)
        for (let i = 0; i < nl.length; i++) {
          if (nl[i].hizoMatch) continue
          const tot = nl[i].totalCFDI + nl[i].montoPropina
          const dG  = parseDateRobusto(nl[i].fechaFac)
          if (Math.abs(tot - csv) <= 5 && dG && Math.abs(dCSV - dG) / 86400000 <= 30) {
            nl[i].hizoMatch = true
            nl[i].fechaCobro = formatCobro(dCSV)
            matches++; found = true; break
          }
        }
        if (found) break
        // Paso 2: detección automática de propina (csv > total hasta 25% extra)
        for (let i = 0; i < nl.length; i++) {
          if (nl[i].hizoMatch) continue
          const inv = nl[i].totalCFDI
          const dG  = parseDateRobusto(nl[i].fechaFac)
          if (csv > inv + 2 && csv <= inv * 1.25 && dG && Math.abs(dCSV - dG) / 86400000 <= 30) {
            nl[i].hizoMatch = true
            nl[i].fechaCobro = formatCobro(dCSV)
            const prop = csv - inv
            nl[i].montoPropina = prop
            if (inv > 0) nl[i].propinaPorcentaje = (prop / inv) * 100
            matches++; propinas++; found = true; break
          }
        }
      }
    }

    setLista(nl)
    setAlerta(
      `✅ Conciliación Terminada:\n\n` +
      `💳 Cargos en Banco: ${bancoRows}\n` +
      `🎯 Matches exitosos: ${matches}\n` +
      `🪄 Propinas detectadas automáticamente: ${propinas}\n\n` +
      `⚠️ Cargos del banco SIN factura: ${bancoRows - matches}\n` +
      `⚠️ Facturas SIN cargo en tarjeta: ${nl.length - matches}`
    )
    e.target.value = ''
  }

  // ── Agregar fila manual ──
  const agregarManual = () => {
    const hoy = new Date().toISOString().slice(0, 10)
    setLista(prev => [...prev, {
      id: genId(),
      rfc: 'PUBLICO GENERAL', proveedor: 'Escribe aquí...', noFactura: 'Ticket',
      fechaFac: hoy, concepto: 'Consumo', importe: 0, iva: 0, isrTrasladado: 0,
      retencionISR: 0, retencionIVA: 0, retenciones: 0, totalCFDI: 0,
      propinaPorcentaje: 0, montoPropina: 0, fechaCobro: '', formaPago: '01', uuid: 'MANUAL',
      tienePDF: false, pdfFile: null, xmlFile: null, hizoMatch: false, checkManual: false,
    }])
  }

  // ── Copiar a portapapeles (TSV para Excel) ──
  const copiar = () => {
    const hdr = 'RFC PROVEEDOR\tPROVEEDOR\tNO. DE FACTURA\tFECHA FAC.\tCONCEPTO\tIMPORTE (MXP)\tIVA\tISR\tRET. ISR\tRET. IVA\tRET/ ISR IVA\tTOTAL CFDI\tGastos en USD\tTipo de Cambio\tTotal Checking\tFORMA DE PAGO\tFECHA DE COBRO\n'
    const rows = lista.flatMap(g => {
      const f  = g.fechaFac.split('-')
      const mx = f.length === 3 ? `${f[2]}/${f[1]}/${f[0]}` : g.fechaFac
      const cobro = g.fechaCobro || 'Pendiente'
      const r  = `${g.rfc}\t${g.proveedor.replace(/\t/g,' ')}\t${g.noFactura}\t${mx}\t${g.concepto.replace(/\t/g,' ')}\t${g.importe.toFixed(2)}\t${g.iva.toFixed(2)}\t${(g.isrTrasladado||0).toFixed(2)}\t${(g.retencionISR||0).toFixed(2)}\t${(g.retencionIVA||0).toFixed(2)}\t${g.retenciones.toFixed(2)}\t${g.totalCFDI.toFixed(2)}\t\t\t\t${g.formaPago}\t${cobro}\n`
      const p  = g.montoPropina > 0
        ? `\t${g.proveedor} - PROPINA\t\t${mx}\tPROPINA\t${g.montoPropina.toFixed(2)}\t0.00\t0.00\t0.00\t0.00\t0.00\t${g.montoPropina.toFixed(2)}\t\t\t\t${g.formaPago}\t${cobro}\n`
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
      const f   = g.fechaFac.split('-')
      const mx  = f.length === 3 ? `${f[2]}/${f[1]}/${f[0]}` : g.fechaFac
      const fa  = f.length === 3 ? `${f[1]}-${f[2]}-${f[0].slice(-2)}` : g.fechaFac.replace(/\//g, '-')
      const cobro = g.fechaCobro || 'Pendiente'
      csv += `${g.rfc},${g.proveedor.replace(/,/g,' ')},${g.noFactura},${mx},${g.concepto.replace(/,/g,' ')},${g.importe.toFixed(2)},${g.iva.toFixed(2)},${(g.isrTrasladado||0).toFixed(2)},${(g.retencionISR||0).toFixed(2)},${(g.retencionIVA||0).toFixed(2)},${g.retenciones.toFixed(2)},${g.totalCFDI.toFixed(2)},,,,${g.formaPago},${cobro}\n`
      if (g.montoPropina > 0)
        csv += `,${g.proveedor} - PROPINA,,${mx},PROPINA,${g.montoPropina.toFixed(2)},0.00,0.00,0.00,0.00,0.00,${g.montoPropina.toFixed(2)},,,,${g.formaPago},${cobro}\n`

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
    const a = Object.assign(document.createElement('a'), {
      href: URL.createObjectURL(blob),
      download: 'Reporte_Gastos_Empaquetado.zip',
    })
    a.click(); URL.revokeObjectURL(a.href)
    setAlerta(`¡Éxito Total! 📦\n\nSe ha generado el archivo ZIP con tu CSV y ${r} facturas (PDF y XML) correctamente renombradas.\n\nNota: Tus archivos originales siguen intactos.`)
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
          <h1 className="header-title">Reporte de Gastos SMTO<span className="version-badge">v1.9</span></h1>
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
          <PremiumButton title="Copiar a Excel" icon="📋" variant="copy"   isDisabled={!lista.length} onClick={copiar} />
          <PremiumButton title="Exportar a ZIP" icon="📦" variant="export" isDisabled={!lista.length} onClick={exportar} />
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
            // Sticky-column left offsets derived from current column widths, so
            // resizing Estado / Fecha shifts the pinned cells correctly.
            '--sl-fecha':     `${colWidths.status}px`,
            '--sl-proveedor': `${colWidths.status + colWidths.fechaFac}px`,
          }}>
            <thead>
              <tr>
                {COLUMNS.map(col => (
                  <th
                    key={col.key}
                    className={col.key === 'status' ? 'th-status' : col.key === 'fechaFac' ? 'th-fecha' : col.key === 'proveedor' ? 'th-proveedor' : undefined}
                    style={{ width: colWidths[col.key], cursor: col.sortable ? 'pointer' : undefined }}
                    onClick={col.sortable ? () => toggleSort(col.key) : undefined}
                  >
                    {col.label}
                    {sort.field === col.key && (
                      <span style={{ marginLeft: 4, fontSize: '0.85em' }}>{sort.dir === 'asc' ? '▲' : '▼'}</span>
                    )}
                    <div
                      className="col-resizer"
                      onMouseDown={e => startResize(col.key, e)}
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

      {/* ─── MODAL ALERTA ─── */}
      {alerta && (
        <div className="overlay" onClick={() => setAlerta(null)}>
          <div className="modal" onClick={e => e.stopPropagation()}>
            <div className="modal-title">Resultado</div>
            <pre className="modal-body">{alerta}</pre>
            <button className="modal-ok" onClick={() => setAlerta(null)}>OK</button>
          </div>
        </div>
      )}
    </div>
  )
}
