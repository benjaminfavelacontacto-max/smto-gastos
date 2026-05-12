import { useState, useRef, useMemo } from 'react'
import JSZip from 'jszip'

/* Single source of truth for table columns.
   `getValue` overrides simple `g[key]` lookup (used by Total Final). */
const COLUMNS = [
  { key: 'check',             label: '',            width: 30,  sortable: false },
  { key: 'status',            label: 'Estado',      width: 68,  sortable: false },
  { key: 'fechaFac',          label: 'Fecha',       width: 92,  sortable: true,  type: 'date'   },
  { key: 'noFactura',         label: 'Factura',     width: 104, sortable: true,  type: 'string' },
  { key: 'proveedor',         label: 'Proveedor',   width: 155, sortable: true,  type: 'string' },
  { key: 'concepto',          label: 'Concepto',    width: 118, sortable: true,  type: 'string' },
  { key: 'importe',           label: 'Subtotal',    width: 88,  sortable: true,  type: 'number' },
  { key: 'iva',               label: 'IVA',         width: 76,  sortable: true,  type: 'number' },
  { key: 'isrTrasladado',     label: 'ISR',         width: 84,  sortable: true,  type: 'number' },
  { key: 'retencionISR',      label: 'Ret. ISR',    width: 84,  sortable: true,  type: 'number' },
  { key: 'retencionIVA',      label: 'Ret. IVA',    width: 84,  sortable: true,  type: 'number' },
  { key: 'retenciones',       label: 'Reten.',      width: 84,  sortable: true,  type: 'number' },
  { key: 'totalCFDI',         label: 'Total Fac.',  width: 96,  sortable: true,  type: 'number' },
  { key: 'propinaPorcentaje', label: 'Prop. %',     width: 78,  sortable: true,  type: 'number' },
  { key: 'montoPropina',      label: 'Prop. $',     width: 88,  sortable: true,  type: 'number' },
  { key: 'totalFinal',        label: 'Total Final', width: 102, sortable: true,  type: 'number',
    getValue: g => g.totalCFDI + g.montoPropina },
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

  let iva = parseFloat(ga(rootImp, 'TotalImpuestosTrasladados', 'totalImpuestosTrasladados') || '0') || 0
  let isrTrasladado = 0
  let retencionISR  = 0
  let retencionIVA  = 0
  const needIva = !iva

  if (rootImp) {
    for (const c of rootImp.childNodes) {
      if (c.nodeType !== 1 || !c.localName) continue
      const ln = c.localName.toLowerCase()
      if (ln === 'traslados') {
        for (const t of c.childNodes) {
          if (t.nodeType !== 1 || (t.localName || '').toLowerCase() !== 'traslado') continue
          // CFDI 4.0 actually names this attribute "Impuesto" on Traslado/Retencion.
          // Some tooling/spec write-ups call it "TipoImpuesto" — accept both.
          const tipo    = (ga(t, 'TipoImpuesto', 'tipoImpuesto', 'Impuesto', 'impuesto') || '').trim()
          const importe = parseFloat(ga(t, 'Importe', 'importe') || '0') || 0
          if      (tipo === '001')             isrTrasladado += importe
          else if (tipo === '002' && needIva)  iva           += importe
        }
      } else if (ln === 'retenciones') {
        const retEls = []
        for (const r of c.childNodes) {
          if (r.nodeType === 1 && (r.localName || '').toLowerCase() === 'retencion') retEls.push(r)
        }
        console.log('[parseCFDI]', xmlFile.name, '— Retencion elements found:', retEls.length)
        for (const r of retEls) {
          const tipo    = (ga(r, 'TipoImpuesto', 'tipoImpuesto', 'Impuesto', 'impuesto') || '').trim()
          const importe = parseFloat(ga(r, 'Importe', 'importe') || '0') || 0
          console.log('[parseCFDI]   Retencion tipo=', tipo, 'importe=', importe)
          if      (tipo === '001') retencionISR += importe
          else if (tipo === '002') retencionIVA += importe
        }
      }
    }
  }
  const retenciones = retencionISR + retencionIVA

  // Buscar PDF asociado (por nombre base o UUID)
  const base = xmlFile.name.replace(/\.xml$/i, '').toLowerCase()
  const pdfFile = pdfFiles.find(f =>
    f.name.replace(/\.pdf$/i, '').toLowerCase() === base ||
    (uuid && f.name.toUpperCase().includes(uuid.toUpperCase()))
  ) || null

  return {
    id: genId(),
    rfc,
    proveedor,
    noFactura:  ga(comp, 'Folio', 'folio') || 'SN',
    fechaFac:  (ga(comp, 'Fecha', 'fecha') || '').slice(0, 10),
    concepto:   clasificarGasto(proveedor, concepto),
    importe:    parseFloat(ga(comp, 'SubTotal', 'subtotal') || '0') || 0,
    iva,
    isrTrasladado,
    retencionISR,
    retencionIVA,
    retenciones,
    totalCFDI:  parseFloat(ga(comp, 'Total', 'total')            || '0') || 0,
    propinaPorcentaje: 0,
    montoPropina: 0,
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

function PremiumButton({ title, icon, color, isSecondary = false, isDisabled = false, onClick }) {
  const [hov, setHov] = useState(false)

  const solidStyle = (!isSecondary && !isDisabled) ? {
    background: color,
    transform: hov ? 'scale(1.035)' : 'scale(1)',
    boxShadow: hov
      ? `0 5px 18px ${color}80`
      : `0 2px 8px  ${color}55`,
  } : {}

  return (
    <button
      className={`btn-premium${isSecondary ? ' secondary' : ''}${isDisabled ? ' disabled' : ''}`}
      style={solidStyle}
      onClick={isDisabled ? undefined : onClick}
      disabled={isDisabled}
      onMouseEnter={() => !isDisabled && setHov(true)}
      onMouseLeave={() => setHov(false)}
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

  const NumCell = ({ field, prefix, suffix }) => (
    <div className="num-cell">
      {prefix && <span className="sym">{prefix}</span>}
      <input
        type="number"
        className="cell-in"
        value={g[field] || ''}
        placeholder="0"
        onChange={e => upd(field, parseFloat(e.target.value) || 0)}
      />
      {suffix && <span className="sym">{suffix}</span>}
    </div>
  )

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

      {/* Estado: PDF + Banco */}
      <td className="td-status">
        <div className="status-row">
          <button
            className={`dot${g.tienePDF ? ' dot-pdf' : ' dot-empty'}`}
            onClick={g.tienePDF ? () => openPDF(g.pdfFile) : undefined}
            disabled={!g.tienePDF}
            title={g.tienePDF ? 'Clic para abrir el PDF' : 'Sin PDF'}
          >
            {g.tienePDF ? (
              <svg width="11" height="13" viewBox="0 0 11 13" fill="none"><path d="M1 1h6l3 3v8H1V1z" stroke="white" strokeWidth="1.4" strokeLinejoin="round"/><path d="M7 1v3h3" stroke="white" strokeWidth="1.4"/></svg>
            ) : (
              <svg width="11" height="13" viewBox="0 0 11 13" fill="none"><path d="M1 1h6l3 3v8H1V1z" stroke="currentColor" strokeWidth="1.4" strokeLinejoin="round"/><path d="M7 1v3h3" stroke="currentColor" strokeWidth="1.4"/></svg>
            )}
          </button>
          <div
            className={`dot${g.hizoMatch ? ' dot-match' : ' dot-empty'}`}
            title={g.hizoMatch ? 'Conciliado con el Banco' : 'Pendiente en Banco'}
          >
            {g.hizoMatch ? (
              <svg width="11" height="9" viewBox="0 0 11 9" fill="none"><path d="M1 4.5L4 7.5L10 1" stroke="white" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round"/></svg>
            ) : (
              <svg width="12" height="11" viewBox="0 0 12 11" fill="none"><rect x="1" y="3" width="10" height="7" rx="1.5" stroke="currentColor" strokeWidth="1.3"/><path d="M3 3V2a3 3 0 016 0v1" stroke="currentColor" strokeWidth="1.3"/></svg>
            )}
          </div>
        </div>
      </td>

      {/* Fecha */}
      <td>
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
      <td>
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

      {/* Propina % */}
      <td><NumCell field="propinaPorcentaje" suffix="%" /></td>

      {/* Propina $ */}
      <td><NumCell field="montoPropina" prefix="$" /></td>

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
   APP PRINCIPAL
═══════════════════════════════════════════════════ */

export default function App() {
  const [lista,         setLista]         = useState([])
  const [carpetaNombre, setCarpetaNombre] = useState('Ninguna carpeta seleccionada')
  const [alerta,        setAlerta]        = useState(null)
  const [loading,       setLoading]       = useState(false)
  const [colWidths,     setColWidths]     = useState(() =>
    Object.fromEntries(COLUMNS.map(c => [c.key, c.width]))
  )
  const [sort,          setSort]          = useState({ field: null, dir: 'asc' })

  const folderRef = useRef(null)
  const bancoRef  = useRef(null)

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
  const cargar = async e => {
    const files = Array.from(e.target.files)
    if (!files.length) return
    const xmls = files.filter(f => f.name.toLowerCase().endsWith('.xml'))
    const pdfs = files.filter(f => f.name.toLowerCase().endsWith('.pdf'))
    const folder = files[0]?.webkitRelativePath?.split('/')[0] || 'Carpeta'
    setCarpetaNombre(folder)
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
    e.target.value = ''
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
    const nl = lista.map(g => ({ ...g, hizoMatch: false }))

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
            nl[i].hizoMatch = true; matches++; found = true; break
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
      propinaPorcentaje: 0, montoPropina: 0, formaPago: '01', uuid: 'MANUAL',
      tienePDF: false, pdfFile: null, xmlFile: null, hizoMatch: false, checkManual: false,
    }])
  }

  // ── Copiar a portapapeles (TSV para Excel) ──
  const copiar = () => {
    const hdr = 'RFC PROVEEDOR\tPROVEEDOR\tNO. DE FACTURA\tFECHA FAC.\tCONCEPTO\tIMPORTE (MXP)\tIVA\tISR\tRET. ISR\tRET. IVA\tRET/ ISR IVA\tTOTAL CFDI\tGastos en USD\tTipo de Cambio\tTotal Checking\tFORMA DE PAGO\tFECHA DE COBRO\n'
    const rows = lista.flatMap(g => {
      const f  = g.fechaFac.split('-')
      const mx = f.length === 3 ? `${f[2]}/${f[1]}/${f[0]}` : g.fechaFac
      const r  = `${g.rfc}\t${g.proveedor.replace(/\t/g,' ')}\t${g.noFactura}\t${mx}\t${g.concepto.replace(/\t/g,' ')}\t${g.importe.toFixed(2)}\t${g.iva.toFixed(2)}\t${(g.isrTrasladado||0).toFixed(2)}\t${(g.retencionISR||0).toFixed(2)}\t${(g.retencionIVA||0).toFixed(2)}\t${g.retenciones.toFixed(2)}\t${g.totalCFDI.toFixed(2)}\t\t\t\t${g.formaPago}\tPendiente\n`
      const p  = g.montoPropina > 0
        ? `\t${g.proveedor} - PROPINA\t\t${mx}\tPROPINA\t${g.montoPropina.toFixed(2)}\t0.00\t0.00\t0.00\t0.00\t0.00\t${g.montoPropina.toFixed(2)}\t\t\t\t${g.formaPago}\tPendiente\n`
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
      csv += `${g.rfc},${g.proveedor.replace(/,/g,' ')},${g.noFactura},${mx},${g.concepto.replace(/,/g,' ')},${g.importe.toFixed(2)},${g.iva.toFixed(2)},${(g.isrTrasladado||0).toFixed(2)},${(g.retencionISR||0).toFixed(2)},${(g.retencionIVA||0).toFixed(2)},${g.retenciones.toFixed(2)},${g.totalCFDI.toFixed(2)},,,,${g.formaPago},Pendiente\n`
      if (g.montoPropina > 0)
        csv += `,${g.proveedor} - PROPINA,,${mx},PROPINA,${g.montoPropina.toFixed(2)},0.00,0.00,0.00,0.00,0.00,${g.montoPropina.toFixed(2)},,,,${g.formaPago},Pendiente\n`

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
    const url = URL.createObjectURL(pdfFile)
    window.open(url, '_blank')
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
          <h1 className="header-title">Reporte de Gastos SMTO</h1>
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
          <PremiumButton title="Manual"        icon="＋" isSecondary onClick={agregarManual} />
          <PremiumButton title="Cargar Carpeta" icon="📂" color="#007AFF" onClick={() => folderRef.current?.click()} />
          <PremiumButton title="Validar Banco"  icon="🏦" color="#5856D6" onClick={() => bancoRef.current?.click()} />
        </div>
        <div className="action-group">
          <PremiumButton title="Copiar a Excel" icon="📋" color="#FF9500" isDisabled={!lista.length} onClick={copiar} />
          <PremiumButton title="Exportar a ZIP" icon="📦" color="#34C759" isDisabled={!lista.length} onClick={exportar} />
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
          <div className="center-msg">
            <div className="spinner" />
            <div>Procesando archivos XML…</div>
          </div>
        ) : lista.length === 0 ? (
          <div className="center-msg">
            <div className="empty-icon">
              <svg width="64" height="56" viewBox="0 0 64 56" fill="none" opacity=".3">
                <path d="M4 12A4 4 0 018 8h14l6 6h26a4 4 0 014 4v28a4 4 0 01-4 4H8a4 4 0 01-4-4V12z" stroke="#1c1c1e" strokeWidth="3" fill="none"/>
              </svg>
            </div>
            <div className="empty-title">Sin registros</div>
            <div className="empty-sub">Haz clic en "Cargar Carpeta" para importar tus facturas XML</div>
          </div>
        ) : (
          <table>
            <thead>
              <tr>
                {COLUMNS.map(col => (
                  <th
                    key={col.key}
                    style={{ width: colWidths[col.key], position: 'relative', cursor: col.sortable ? 'pointer' : undefined }}
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
