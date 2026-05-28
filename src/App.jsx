import { useState, useRef, useMemo, useEffect } from 'react'
import JSZip from 'jszip'
import * as XLSX from 'xlsx'
import { motion, AnimatePresence } from 'framer-motion'
import { CheckCircle2, X, CreditCard, Target, Sparkles, AlertTriangle, FileText, FileSpreadsheet, Package, Check, Plus, Link2, Search, Download, ArrowRight, ChevronDown, XCircle, AlertCircle } from 'lucide-react'
import { autoDetectTipo } from './tipoRules'

/* Two type lists — picked by colaborador.categoria. Ventas/Socio see the
   sales-flavored list (Hotel Ventas, Gasolina Ventas, …), everyone else
   (Admin/Servicio) sees the operational list. */
const TIPOS_VENTAS = [
  'Aduana',
  'Automóvil',
  'Avión Ventas',
  'Casetas Ventas',
  'COGS',
  'Comercio Exterior',
  'Consultoría',
  'Contador',
  'Envíos',
  'Estacionamiento Ventas',
  'Gasolina Ventas',
  'Gasolina Viáticos',
  'Gastos Rep (Representación)',
  'Gastos Rep Viáticos',
  'Herramientas Ventas',
  'Hotel Ventas',
  'IT & SW (Software/Sistemas)',
  'Marketing',
  'Renta Oficina',
]

const TIPOS_NORMALES = [
  'Aduana',
  'Automóvil',
  'Avión',
  'Casetas',
  'Celular',
  'COGS',
  'Comercio Exterior',
  'Consultoría',
  'Consumo',
  'Consumo Viáticos',
  'Contador',
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

/* Extended list for Alejandro Olivar and Victor Aceves */
const TIPOS_ESPECIALES = [
  '3% ISN',
  'Aduana',
  'Autobus',
  'Automóvil',
  'Avión',
  'COGS',
  'Caseta',
  'Celular',
  'Comercio Exterior',
  'Comisión Banco',
  'Consultoría',
  'Contador',
  'Consumo',
  'Consumo Viáticos',
  'Curso',
  'Envío',
  'Envíos',
  'Estacionamiento',
  'Fondo de Ahorro',
  'Gasolina',
  'Gasolina Viáticos',
  'Herramienta',
  'Hotel',
  'IMSS',
  'ISR',
  'IT&SW',
  'IT & SW (Software/Sistemas)',
  'Manto Auto',
  'Marketing',
  'No Comprobado',
  'Nómina',
  'PC',
  'Papelería',
  'Pasaporte o Visa',
  'Permiso',
  'Préstamo',
  'Regalía',
  'Renta Auto',
  'Renta Oficina',
  'Seguro',
  'Taxi',
  'Transporte',
  'Uniforme',
]

const COLABORADORES_ESPECIALES = ['Alejandro Olivar', 'Victor Aceves', 'Miranda Navarro', 'Olivia Gil']

const getTiposForColaborador = (colaborador) => {
  if (!colaborador) return TIPOS_NORMALES
  if (COLABORADORES_ESPECIALES.includes(colaborador.nombre)) return TIPOS_ESPECIALES
  const cat = colaborador.categoria
  if (cat === 'Ventas' || cat === 'Socio') return TIPOS_VENTAS
  return TIPOS_NORMALES  // Admin, Servicio
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
  { categoria: 'Socio', nombre: 'Olivia Gil' },
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
  { categoria: 'Admin', nombre: 'Miranda Navarro' },
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

// Re-encode an image File at most `maxWidth` px wide, JPEG `quality`,
// before shipping to OCR. Phone photos arrive at 5–10 MB which routinely
// breaks the 4.5 MB Vercel serverless body limit (HTTP 413). This brings
// them under the cap with negligible quality loss for OCR. Non-image
// files pass through untouched.
const compressImage = async (file, maxWidth = 2000, quality = 0.85) => {
  if (!file.type.startsWith('image/')) return file
  return new Promise((resolve, reject) => {
    const img = new Image()
    const reader = new FileReader()
    reader.onload = (e) => {
      img.onload = () => {
        let { width, height } = img
        if (width > maxWidth) {
          height = Math.round(height * (maxWidth / width))
          width = maxWidth
        }
        const canvas = document.createElement('canvas')
        canvas.width = width
        canvas.height = height
        const ctx = canvas.getContext('2d')
        ctx.drawImage(img, 0, 0, width, height)
        canvas.toBlob(
          (blob) => {
            const compressed = new File(
              [blob],
              file.name.replace(/\.(heic|heif|webp|png)$/i, '.jpg'),
              { type: 'image/jpeg' }
            )
            resolve(compressed)
          },
          'image/jpeg',
          quality
        )
      }
      img.onerror = reject
      img.src = e.target.result
    }
    reader.onerror = reject
    reader.readAsDataURL(file)
  })
}

// Wrap an image data URL into a single-page PDF (data URL out). Used at
// ZIP-export time so an OCR'd photo lands in Facturas/ with the same
// PROVEEDOR_FOLIO_TIPO_FECHA naming as a CFDI-linked PDF. jsPDF is
// loaded via CDN script tag in index.html — kept out of the bundle so
// the main JS chunk stays lean.
const imageToPDF = async (imageDataURL) => {
  return new Promise((resolve, reject) => {
    const img = new Image()
    img.onload = () => {
      try {
        const orientation = img.width > img.height ? 'l' : 'p'
        const { jsPDF } = window.jspdf || {}
        if (!jsPDF) throw new Error('jsPDF not loaded')
        const pdf = new jsPDF({ orientation, unit: 'px', format: [img.width, img.height] })
        pdf.addImage(imageDataURL, 'JPEG', 0, 0, img.width, img.height)
        const pdfBlob = pdf.output('blob')
        const reader = new FileReader()
        reader.onload = () => resolve(reader.result)
        reader.onerror = reject
        reader.readAsDataURL(pdfBlob)
      } catch (err) {
        reject(err)
      }
    }
    img.onerror = reject
    img.src = imageDataURL
  })
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

// Read a File/Blob as a full data URL (e.g. "data:application/pdf;base64,…").
// Used to persist PDF bytes onto the gasto so the ZIP export is independent
// of the original File reference, which can become unreadable in some
// browsers once the source <input> / drop event is GC'd.
const fileToDataURL = (file) => new Promise((resolve, reject) => {
  const reader = new FileReader()
  reader.onload = () => resolve(String(reader.result || ''))
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
  let descripcionRaw = conceptoEl ? (ga(conceptoEl, 'Descripcion', 'descripcion') || '') : ''
  let claveProdServ  = conceptoEl ? (ga(conceptoEl, 'ClaveProdServ', 'claveprodserv') || '') : ''
  // Fallback por regex: cuando la Descripcion contiene comillas literales (ej. 15.6" FHD...)
  // el atributo XML queda truncado en el DOM pero el texto crudo sí lo tiene.
  if (!descripcionRaw) {
    const dm = xmlText.match(/[Dd]escripcion="([^"]+)"/); if (dm) descripcionRaw = dm[1]
  }
  if (!claveProdServ) {
    const cm = xmlText.match(/ClaveProdServ="([^"]+)"/i); if (cm) claveProdServ = cm[1]
  }
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
  let   retencionIVA  = sumByTipo(retencionesBox, 'retencion', '002')
  let   retenciones   = retencionISR + retencionIVA

  // Fallback por regex sobre el XML crudo: cuando la Descripcion contiene comillas
  // literales (ej. 15.6"), el DOMParser trunca el documento y el root <cfdi:Impuestos>
  // nunca llega a construirse → iva/retenciones=0 aunque el XML SÍ los declara.
  // Si rootImp no existe O iva quedó en 0 con totalCFDI>subtotal, recurrimos a regex.
  const hasRootImp = !!rootImp
  const totalCFDIQuick = parseFloat(ga(comp, 'Total', 'total') || '0') || 0
  const subTotalQuick  = parseFloat(ga(comp, 'SubTotal', 'subtotal') || '0') || 0
  if ((!hasRootImp || iva === 0) && totalCFDIQuick > subTotalQuick) {
    const mTras = xmlText.match(/TotalImpuestosTrasladados="([0-9.]+)"/i)
    if (mTras) {
      const v = parseFloat(mTras[1]) || 0
      if (v > 0) iva = v
    }
    const mRet = xmlText.match(/TotalImpuestosRetenidos="([0-9.]+)"/i)
    if (mRet && retenciones === 0) {
      const v = parseFloat(mRet[1]) || 0
      if (v > 0) {
        // Buscar el desglose por tipo (001=ISR, 002=IVA) en los hijos <Retencion>.
        // Capturamos SOLO el bloque <cfdi:Retenciones>...</cfdi:Retenciones> del root.
        const retBlock = xmlText.match(/<\w*:?Retenciones>([\s\S]*?)<\/\w*:?Retenciones>/i)
        if (retBlock) {
          let isrSum = 0, ivaSum = 0
          const reEl = /<\w*:?Retencion\b[^>]*>/gi
          let m
          while ((m = reEl.exec(retBlock[1])) !== null) {
            const tag = m[0]
            const tipo = (tag.match(/Impuesto="([0-9]+)"/i) || [])[1] || ''
            const imp  = parseFloat((tag.match(/Importe="([0-9.]+)"/i) || [])[1] || '0') || 0
            if (tipo === '001') isrSum += imp
            else if (tipo === '002') ivaSum += imp
          }
          if (isrSum > 0 || ivaSum > 0) {
            retencionISR = isrSum
            retencionIVA = ivaSum
            retenciones  = isrSum + ivaSum
          } else {
            retenciones = v
          }
        } else {
          retenciones = v
        }
      }
    }
  }

  // Per-RFC override: some providers' <Retencion Impuesto="001"> is actually
  // a trasladado ISR (the invoice line-item tax), not a withholding.
  if (RFC_ISR_COMO_TRASLADO.includes(rfc)) {
    isrTrasladado += retencionISR
    retencionISR   = 0
    retenciones    = retencionIVA
  }

  // Hoteles con complemento ISH que generan un artefacto de "retención mínima IVA"
  // de $1.00 en cfdi:Retenciones aunque no procede. Se zeroan por RFC.
  // Patrón: CFDI tiene TrasladosLocales ISH + retencionIVA <= 2.00 (claramente incorrecto).
  const RFC_RETENCION_IVA_CERO = ['DBM121023M10']  // FIDEICOMISO IRREVOCABLE DB/1616 (Hampton Inn)
  if (RFC_RETENCION_IVA_CERO.includes(rfc) && retencionIVA > 0 && retencionIVA <= 2) {
    retenciones   -= retencionIVA
    retencionIVA   = 0
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
    tipo: autoDetectTipo(proveedor, descripcionFirstLine, colaborador?.categoria, claveProdServ),
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
    // The raw XML text gets stashed here so the ZIP export can write the
    // file without re-reading the original File handle (which is unreliable
    // long-term in some browsers).
    xmlContent: xmlText,
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

      {/* Estado: PDF + IMG (image OCR) + Banco — pill badges */}
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
          {g.imageDataURL && (
            <button
              type="button"
              className="badge-img"
              title="Ver imagen original"
              onClick={(e) => {
                e.stopPropagation()
                const w = window.open('', '_blank')
                if (w) {
                  w.document.write(`
                    <html>
                      <head>
                        <title>${g.proveedor || 'Imagen'}</title>
                        <style>
                          body { margin: 0; background: #0a0a0a; display: flex; align-items: center; justify-content: center; min-height: 100vh; }
                          img { max-width: 100%; max-height: 100vh; object-fit: contain; }
                        </style>
                      </head>
                      <body>
                        <img src="${g.imageDataURL}" alt="${g.proveedor || ''}" />
                      </body>
                    </html>
                  `)
                  w.document.close()
                }
              }}
            >
              IMG
            </button>
          )}
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

      {/* Monto USD — editing it auto-derives Tipo de Cambio from totalCFDI.
          When Pass 0 of validarBanco detects a foreign-currency tip on a
          matched ticket, it surfaces here as a small green subtitle. */}
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
        {g.propinaExtranjero > 0 && (
          <div className="propina-extranjera-hint" title="Propina detectada en moneda extranjera">
            💵 +${g.propinaExtranjero.toFixed(2)} propina
          </div>
        )}
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

function ConciliacionModal({ data, onClose, onConfirm, onCancel, onAgregarManual }) {
  const total = Math.max(1, data.bancoRows || 0)
  const pct = Math.min(100, Math.round((data.matches / total) * 100))
  const cBanco    = useCountUp(data.bancoRows)
  const cMatches  = useCountUp(data.matches)
  const cSin      = useCountUp(data.sinFactura.length)
  const cPropinas = useCountUp(data.propinas)
  const cPct      = useCountUp(pct)

  const matchedRows = data.matchedRows || []
  const revisionRows = matchedRows.filter(m => m.confidence < 80)
  const matchCount = data.matches || 0

  // Normalize close/cancel: an explicit Cancel is the same as closing the
  // modal without confirming — matches don't get applied either way.
  const handleCancel = onCancel || onClose

  const [tab, setTab] = useState('matches')
  const [query, setQuery] = useState('')
  const [filterKey, setFilterKey] = useState('all')
  const [expandedId, setExpandedId] = useState(null)
  const [loading, setLoading] = useState(true)
  useEffect(() => {
    const t = setTimeout(() => setLoading(false), 800)
    return () => clearTimeout(t)
  }, [])

  // Esc → cancel, Enter → confirm. Skip Enter when the user is typing in a
  // form control (the search input is right there).
  useEffect(() => {
    const onKey = (e) => {
      if (e.key === 'Escape') {
        e.preventDefault()
        handleCancel && handleCancel()
      } else if (e.key === 'Enter') {
        const tag = e.target && e.target.tagName
        if (tag === 'INPUT' || tag === 'TEXTAREA' || tag === 'SELECT') return
        e.preventDefault()
        onConfirm && onConfirm()
      }
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [handleCancel, onConfirm])

  const fmtMoney = (n, currency = 'MXN') => {
    const num = Number(n) || 0
    return `${currency === 'USD' ? 'US$' : '$'}${num.toLocaleString('es-MX', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`
  }

  const matchesPassesFilter = (m) => {
    if (filterKey === 'usd') return m.csvMoneda === 'USD' || m.invoiceMoneda === 'USD'
    if (filterKey === 'mxn') return (m.csvMoneda || 'MXN') === 'MXN' && (m.invoiceMoneda || 'MXN') === 'MXN'
    if (filterKey === 'tickets') return m.isTicket
    if (filterKey === 'facturas') return !m.isTicket
    return true
  }
  const sinFactPassesFilter = (s) => {
    if (filterKey === 'usd') return (s.moneda || 'MXN') === 'USD'
    if (filterKey === 'mxn') return (s.moneda || 'MXN') === 'MXN'
    return true   // tickets/facturas filter is only meaningful for matched rows
  }
  const matchesQuery = (txt) => {
    if (!query.trim()) return true
    const q = query.trim().toLowerCase()
    return (txt || '').toLowerCase().includes(q)
  }

  const visibleMatches = matchedRows
    .filter(m => matchesPassesFilter(m) && (matchesQuery(m.invoiceName) || matchesQuery(m.csvDescripcion)))
    .sort((a, b) => b.confidence - a.confidence)
  const visibleSin      = data.sinFactura.filter(s => sinFactPassesFilter(s) && matchesQuery(s.descripcion))
  const visibleRevision = revisionRows
    .filter(m => matchesPassesFilter(m) && (matchesQuery(m.invoiceName) || matchesQuery(m.csvDescripcion)))
    .sort((a, b) => b.confidence - a.confidence)

  // Confidence tier → chip class + human label. Anything < 50 shouldn't
  // surface here (Pass 2 rejects below 50, Pass 3 floors at 60), but we
  // still bucket it as 'Revisar' defensively.
  const confidenceTier = c => {
    if (c >= 95) return { cls: 'cm-conf-green',  label: 'Alta confianza' }
    if (c >= 75) return { cls: 'cm-conf-blue',   label: 'Confianza media' }
    return            { cls: 'cm-conf-yellow', label: 'Revisar' }
  }

  return (
    <motion.div
      className="cm-fs-overlay"
      initial={{ opacity: 0 }}
      animate={{ opacity: 1 }}
      exit={{ opacity: 0 }}
      transition={{ duration: 0.2 }}
      onClick={onClose}
    >
      <motion.div
        className="cm-fs-shell"
        initial={{ opacity: 0, y: 24, scale: 0.985 }}
        animate={{ opacity: 1, y: 0, scale: 1 }}
        exit={{ opacity: 0, y: 12, scale: 0.99 }}
        transition={{ type: 'spring', stiffness: 240, damping: 26 }}
        onClick={e => e.stopPropagation()}
      >
        <button className="cm-fs-close" onClick={onClose} aria-label="Cerrar">
          <X size={18} />
        </button>

        {/* ───── TOP SUMMARY BAR ───── */}
        <header className="cm-fs-summary">
          <div className="cm-fs-summary-top">
            <div className="cm-fs-summary-titles">
              <h2 className="cm-fs-title">Conciliación Terminada</h2>
              <p className="cm-fs-subtitle">
                {cPct}% conciliado · {data.bancoRows} {data.bancoRows === 1 ? 'cargo' : 'cargos'} procesados
              </p>
            </div>
            <div className="cm-fs-summary-toolbar">
              <label className="cm-fs-search">
                <Search size={14} strokeWidth={2.2} />
                <input
                  type="text"
                  placeholder="Buscar por proveedor o descripción…"
                  value={query}
                  onChange={e => setQuery(e.target.value)}
                />
              </label>
              <select
                className="cm-fs-filter"
                value={filterKey}
                onChange={e => setFilterKey(e.target.value)}
              >
                <option value="all">Todos</option>
                <option value="usd">USD</option>
                <option value="mxn">MXN</option>
                <option value="tickets">Tickets</option>
                <option value="facturas">Facturas</option>
              </select>
              <button
                className="cm-fs-export"
                disabled
                title="Próximamente"
                onClick={e => e.preventDefault()}
              >
                <Download size={14} /> Exportar
              </button>
            </div>
          </div>

          <div className="cm-fs-progress">
            <motion.div
              className="cm-fs-progress-fill"
              initial={{ width: 0 }}
              animate={{ width: `${pct}%` }}
              transition={{ duration: 1.0, ease: 'easeOut', delay: 0.15 }}
            />
          </div>

          <div className="cm-fs-chips">
            <div className="cm-fs-chip cm-chip-blue">
              <CreditCard size={14} />
              <div>
                <div className="cm-fs-chip-value">{cBanco}</div>
                <div className="cm-fs-chip-label">Total cargos</div>
              </div>
            </div>
            <div className="cm-fs-chip cm-chip-green">
              <Target size={14} />
              <div>
                <div className="cm-fs-chip-value">{cMatches}</div>
                <div className="cm-fs-chip-label">Matches exitosos</div>
              </div>
            </div>
            <div className="cm-fs-chip cm-chip-red">
              <XCircle size={14} />
              <div>
                <div className="cm-fs-chip-value">{cSin}</div>
                <div className="cm-fs-chip-label">Sin factura</div>
              </div>
            </div>
            <div className="cm-fs-chip cm-chip-purple">
              <Sparkles size={14} />
              <div>
                <div className="cm-fs-chip-value">{cPropinas}</div>
                <div className="cm-fs-chip-label">Propinas detectadas</div>
              </div>
            </div>
          </div>

          {(data.totalsMatched || data.totalsPending) && (
            <div className="cm-fs-totals">
              <div className="cm-fs-totals-block">
                <span className="cm-fs-totals-label">Conciliado</span>
                <span className="cm-fs-totals-value cm-fs-totals-green">
                  {fmtMoney(data.totalsMatched?.mxn)} MXN
                  {data.totalsMatched?.usd > 0 && <> · {fmtMoney(data.totalsMatched.usd, 'USD')} USD</>}
                </span>
              </div>
              <div className="cm-fs-totals-block">
                <span className="cm-fs-totals-label">Pendiente</span>
                <span className="cm-fs-totals-value cm-fs-totals-red">
                  {fmtMoney(data.totalsPending?.mxn)} MXN
                  {data.totalsPending?.usd > 0 && <> · {fmtMoney(data.totalsPending.usd, 'USD')} USD</>}
                </span>
              </div>
            </div>
          )}
        </header>

        {/* ───── TABS ───── */}
        <div className="cm-fs-tabs" role="tablist">
          {[
            { key: 'matches',  label: 'Matches',     icon: <Check size={14} />,        count: visibleMatches.length,  badgeClass: 'cm-fs-badge-green' },
            { key: 'sin',      label: 'Sin Factura', icon: <XCircle size={14} />,      count: visibleSin.length,      badgeClass: 'cm-fs-badge-red' },
            { key: 'revision', label: 'Revisión',    icon: <AlertCircle size={14} />,  count: visibleRevision.length, badgeClass: 'cm-fs-badge-yellow' },
          ].map(t => (
            <button
              key={t.key}
              role="tab"
              aria-selected={tab === t.key}
              className={`cm-fs-tab ${tab === t.key ? 'is-active' : ''}`}
              onClick={() => setTab(t.key)}
            >
              {t.icon}
              <span>{t.label}</span>
              <span className={`cm-fs-tab-badge ${t.badgeClass}`}>{t.count}</span>
              {tab === t.key && (
                <motion.span className="cm-fs-tab-underline" layoutId="cm-fs-tab-underline" />
              )}
            </button>
          ))}
        </div>

        {/* ───── BODY ───── */}
        <div className="cm-fs-body">
          {loading ? (
            <div className="cm-fs-skel-list">
              {Array.from({ length: 5 }).map((_, i) => (
                <div key={i} className="cm-fs-skel-card" />
              ))}
            </div>
          ) : (
            <AnimatePresence mode="wait">
              {tab === 'matches' && (
                <motion.div
                  key="matches"
                  className="cm-fs-pane"
                  initial={{ opacity: 0, x: 16 }}
                  animate={{ opacity: 1, x: 0 }}
                  exit={{ opacity: 0, x: -16 }}
                  transition={{ duration: 0.25, ease: 'easeOut' }}
                >
                  {visibleMatches.length === 0 ? (
                    <div className="cm-fs-empty">Sin matches que cumplan los filtros.</div>
                  ) : visibleMatches.map((m, i) => {
                    const id = `${m.invoiceId}-${i}`
                    const isOpen = expandedId === id
                    return (
                      <div
                        key={id}
                        className={`cm-fs-card cm-fs-match-card ${isOpen ? 'is-open' : ''}`}
                        onClick={() => setExpandedId(isOpen ? null : id)}
                      >
                        <div className="cm-fs-match-head">
                          <div className="cm-fs-match-left">
                            <CheckCircle2 size={18} className="cm-fs-match-check" />
                            <div>
                              <div className="cm-fs-match-name">{m.csvDescripcion || '(sin descripción)'}</div>
                              <div className="cm-fs-match-sub">{formatDateDisplay(m.csvDate)}</div>
                            </div>
                          </div>
                          <div className="cm-fs-match-center">
                            <ArrowRight size={14} className="cm-fs-match-arrow" />
                            <div>
                              <div className="cm-fs-match-name">{m.invoiceName}</div>
                              <div className="cm-fs-match-sub">{m.method}</div>
                            </div>
                          </div>
                          <div className="cm-fs-match-right">
                            {(() => {
                              const tier = confidenceTier(m.confidence)
                              return (
                                <span
                                  className={`cm-fs-conf ${tier.cls}`}
                                  title={`${m.confidence}%`}
                                >
                                  {tier.label}
                                </span>
                              )
                            })()}
                            <div className="cm-fs-match-amount">
                              {m.csvAmountMXN > 0 && <span>{fmtMoney(m.csvAmountMXN)} MXN</span>}
                              {m.csvAmountUSD > 0 && <span>{fmtMoney(m.csvAmountUSD, 'USD')} USD</span>}
                            </div>
                            <ChevronDown size={14} className={`cm-fs-chev ${isOpen ? 'is-open' : ''}`} />
                          </div>
                        </div>
                        <div className="cm-fs-match-detail" style={{ maxHeight: isOpen ? 220 : 0 }}>
                          <div className="cm-fs-detail-grid">
                            <div><span>Pass</span><strong>{m.pass} — {m.method}</strong></div>
                            <div><span>Factura</span><strong>{m.invoiceNumber || '—'}</strong></div>
                            <div><span>Total factura</span><strong>{fmtMoney(m.invoiceTotal, m.invoiceMoneda)}</strong></div>
                            {m.csvAuth && <div><span>Código autorización</span><strong>{m.csvAuth}</strong></div>}
                            <div><span>Confianza</span><strong>{m.confidence}%</strong></div>
                          </div>
                        </div>
                      </div>
                    )
                  })}
                </motion.div>
              )}

              {tab === 'sin' && (
                <motion.div
                  key="sin"
                  className="cm-fs-pane"
                  initial={{ opacity: 0, x: 16 }}
                  animate={{ opacity: 1, x: 0 }}
                  exit={{ opacity: 0, x: -16 }}
                  transition={{ duration: 0.25, ease: 'easeOut' }}
                >
                  {visibleSin.length === 0 ? (
                    <div className="cm-fs-empty">Sin cargos pendientes que cumplan los filtros.</div>
                  ) : visibleSin.map((s, i) => (
                    <div key={i} className="cm-fs-card cm-fs-sin-card">
                      <div className="cm-fs-sin-head">
                        <div className="cm-fs-sin-left">
                          <span className="cm-fs-badge cm-fs-badge-red">Sin Match</span>
                          <div>
                            <div className="cm-fs-match-name">{s.descripcion || 'Sin descripción'}</div>
                            <div className="cm-fs-match-sub">{formatDateDisplay(s.fecha)}</div>
                          </div>
                        </div>
                        <div className="cm-fs-sin-right">
                          <div className="cm-fs-match-amount">
                            <span>{fmtMoney(s.monto, s.moneda || 'MXN')} {s.moneda || 'MXN'}</span>
                          </div>
                          <button
                            className="cm-fs-add-manual"
                            onClick={() => onAgregarManual && onAgregarManual(s)}
                          >
                            <Plus size={12} /> Agregar manualmente
                          </button>
                        </div>
                      </div>
                      {s.sugerencias && s.sugerencias.length > 0 && (
                        <div className="cm-fs-suggest">
                          <span className="cm-fs-suggest-label">¿Quisiste decir…?</span>
                          {s.sugerencias.map(sg => (
                            <span key={sg.id} className="cm-fs-suggest-chip">
                              {sg.proveedor} <em>{sg.score}%</em>
                            </span>
                          ))}
                        </div>
                      )}
                    </div>
                  ))}
                </motion.div>
              )}

              {tab === 'revision' && (
                <motion.div
                  key="revision"
                  className="cm-fs-pane"
                  initial={{ opacity: 0, x: 16 }}
                  animate={{ opacity: 1, x: 0 }}
                  exit={{ opacity: 0, x: -16 }}
                  transition={{ duration: 0.25, ease: 'easeOut' }}
                >
                  {visibleRevision.length === 0 ? (
                    <div className="cm-fs-empty">Todo conciliado con alta confianza.</div>
                  ) : visibleRevision.map((m, i) => (
                    <div key={`${m.invoiceId}-${i}`} className="cm-fs-card cm-fs-rev-card">
                      <div className="cm-fs-match-head">
                        <div className="cm-fs-match-left">
                          <AlertCircle size={18} className="cm-fs-rev-icon" />
                          <div>
                            <div className="cm-fs-match-name">{m.csvDescripcion || '(sin descripción)'}</div>
                            <div className="cm-fs-match-sub">{formatDateDisplay(m.csvDate)}</div>
                          </div>
                        </div>
                        <div className="cm-fs-match-center">
                          <ArrowRight size={14} className="cm-fs-match-arrow" />
                          <div>
                            <div className="cm-fs-match-name">{m.invoiceName}</div>
                            <div className="cm-fs-match-sub">Pase {m.pass} — {m.method}</div>
                          </div>
                        </div>
                        <div className="cm-fs-match-right">
                          {(() => {
                            const tier = confidenceTier(m.confidence)
                            return (
                              <span className={`cm-fs-conf ${tier.cls}`} title={`${m.confidence}%`}>
                                {tier.label}
                              </span>
                            )
                          })()}
                          <div className="cm-fs-match-amount">
                            {m.csvAmountMXN > 0 && <span>{fmtMoney(m.csvAmountMXN)} MXN</span>}
                            {m.csvAmountUSD > 0 && <span>{fmtMoney(m.csvAmountUSD, 'USD')} USD</span>}
                          </div>
                        </div>
                      </div>
                      <div className="cm-fs-rev-note">
                        Confianza baja. Revisa el monto, la fecha y el proveedor antes de aceptar.
                      </div>
                    </div>
                  ))}
                </motion.div>
              )}
            </AnimatePresence>
          )}
        </div>

        {/* ───── STICKY FOOTER ───── */}
        <footer className="cm-fs-footer">
          <button
            type="button"
            className="cm-fs-btn cm-fs-btn-secondary"
            onClick={handleCancel}
          >
            Cancelar
          </button>
          <button
            type="button"
            className="cm-fs-btn cm-fs-btn-primary"
            onClick={onConfirm}
          >
            Aceptar conciliación · {matchCount} {matchCount === 1 ? 'match' : 'matches'}
          </button>
        </footer>
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
   COMPONENTE: PREMIUM MODAL (reusable info / confirm)
═══════════════════════════════════════════════════ */

// Reusable success / warning / error modal. Replaces the old plain-text
// `setAlerta` modal plus native window.confirm. Specific premium modals
// (Conciliacion / Export / Import / Drop) are richer experiences with
// count-up animations and stat grids; this is the workhorse for everything
// else (errors, confirmations, simple successes).
function PremiumModal({ open, type, title, subtitle, stats, primaryLabel, secondaryLabel, onPrimary, onSecondary, children }) {
  if (!open) return null

  const isSuccess = type === 'success'
  const isWarning = type === 'warning'
  const isError   = type === 'error'

  const accentColor = isSuccess ? '#59D39B'
    : isWarning ? '#f59e0b'
    : isError   ? '#ef4444'
    : '#3b82f6'
  const accentBg = isSuccess ? 'rgba(89,211,155,0.12)'
    : isWarning ? 'rgba(245,158,11,0.12)'
    : isError   ? 'rgba(239,68,68,0.12)'
    : 'rgba(59,130,246,0.12)'

  return (
    <div className="premium-overlay" onClick={onSecondary}>
      <div className="premium-modal" onClick={e => e.stopPropagation()}>
        <div
          className="premium-icon-wrap"
          style={{ background: accentBg, boxShadow: `0 0 40px ${accentColor}30` }}
        >
          {isSuccess && (
            <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke={accentColor} strokeWidth="2.5" className="premium-check">
              <polyline points="20 6 9 17 4 12"/>
            </svg>
          )}
          {isWarning && (
            <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke={accentColor} strokeWidth="2.5">
              <path d="M10.29 3.86L1.82 18a2 2 0 001.71 3h16.94a2 2 0 001.71-3L13.71 3.86a2 2 0 00-3.42 0z"/>
              <line x1="12" y1="9" x2="12" y2="13"/>
              <line x1="12" y1="17" x2="12.01" y2="17"/>
            </svg>
          )}
          {isError && (
            <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke={accentColor} strokeWidth="2.5">
              <circle cx="12" cy="12" r="10"/>
              <line x1="15" y1="9" x2="9" y2="15"/>
              <line x1="9" y1="9" x2="15" y2="15"/>
            </svg>
          )}
        </div>

        <h2 className="premium-title">{title}</h2>
        {subtitle && <p className="premium-subtitle">{subtitle}</p>}

        {stats && stats.length > 0 && (
          <div className="premium-stats">
            {stats.map((s, i) => (
              <div key={i} className="premium-stat-card" style={{ animationDelay: `${i * 80}ms` }}>
                <div className="premium-stat-value" style={{ color: s.color || accentColor }}>{s.value}</div>
                <div className="premium-stat-label">{s.label}</div>
              </div>
            ))}
          </div>
        )}

        {children}

        <div className="premium-actions">
          {secondaryLabel && (
            <button className="premium-btn-secondary" onClick={onSecondary}>
              {secondaryLabel}
            </button>
          )}
          {primaryLabel && (
            <button
              className="premium-btn-primary"
              style={{
                background: `linear-gradient(135deg, ${accentColor}, ${accentColor}dd)`,
                boxShadow: `0 8px 24px ${accentColor}40`,
              }}
              onClick={onPrimary}
            >
              {primaryLabel}
            </button>
          )}
        </div>
      </div>
    </div>
  )
}

/* ═══════════════════════════════════════════════════
   APP PRINCIPAL
═══════════════════════════════════════════════════ */

export default function App() {
  const [lista,         setLista]         = useState([])
  const [carpetaNombre, setCarpetaNombre] = useState('Ninguna carpeta seleccionada')
  // Reusable premium modal for info / warning / error / confirm flows.
  // Replaces the old setAlerta plain-text modal + window.confirm dialogs.
  // Specific success modals (Conciliacion / Export / Import / Drop) keep
  // their own state because they need richer affordances (count-up
  // animations, sin-factura lists, KPI cards, etc.).
  const [modal,         setModal]         = useState(null)
  const [conciliacion,  setConciliacion]  = useState(null)
  const [exportExito,   setExportExito]   = useState(null)
  const [colaborador,   setColaborador]   = useState(null)
  const [showColabModal, setShowColabModal] = useState(true)
  const [colabSearch,   setColabSearch]   = useState('')
  const [importSuccess, setImportSuccess] = useState(false)
  const [importSummary, setImportSummary] = useState(null)
  const [dropSummary,   setDropSummary]   = useState(null)
  const [ocrLoading,    setOcrLoading]    = useState(false)
  const [carpetaSuccess, setCarpetaSuccess] = useState(false)
  const [loading,       setLoading]       = useState(false)
  const [isDraggingOver, setIsDraggingOver] = useState(false)
  const [toast, setToast] = useState(null)
  // Auto-dismiss the inline toast after 2.5s. Re-arm if a new toast arrives
  // before the previous one finishes (clearTimeout in the cleanup).
  useEffect(() => {
    if (!toast) return
    const t = setTimeout(() => setToast(null), 2500)
    return () => clearTimeout(t)
  }, [toast])
  // Index-based fixed pixel widths — order matches COLUMNS positions:
  // [0] checkbox, [1] estado, [2] fecha factura, [3] fecha cobro,
  // [4] factura, [5] proveedor, [6] concepto, [7] tipo, [8] subtotal,
  // [9] iva, [10] isr/ish/ieps, [11] ret.isr, [12] ret.iva, [13] reten,
  // [14] total fac, [15] forma pago, [16] prop%, [17] prop$, [18] total final
  const [colWidths, setColWidths] = useState([40, 110, 115, 120, 120, 260, 140, 100, 120, 110, 135, 110, 110, 110, 125, 160, 95, 105, 130, 110, 80])
  const [sort,          setSort]          = useState({ field: null, dir: 'asc' })

  const folderRef = useRef(null)
  const bancoRef  = useRef(null)
  const photoRef  = useRef(null)

  // ── PremiumModal helpers ──
  const showModal  = (config) => setModal(config)
  const closeModal = () => setModal(null)
  // Promise-wrapped confirm: await askConfirm({...}) returns true on
  // primary click, false on secondary / overlay click.
  const askConfirm = (config) => new Promise(resolve => {
    setModal({
      ...config,
      onPrimary:   () => { resolve(true);  setModal(null) },
      onSecondary: () => { resolve(false); setModal(null) },
    })
  })

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
    // Persist PDF bytes as data URLs so the ZIP export survives even if
    // the original File reference goes stale. XMLs already carry their
    // text via gasto.xmlContent from parseCFDI.
    for (const g of nueva) {
      if (g.pdfFile && !g.pdfDataURL) {
        try { g.pdfDataURL = await fileToDataURL(g.pdfFile) }
        catch (err) { console.warn('PDF read failed:', g.pdfFile.name, err) }
      }
    }

    setLista(nueva)
    setLoading(false)

    // Surface a premium success modal + flash the toolbar button green for
    // 2.5s. Skipped on a zero-result load (empty folder / all parses failed)
    // — the empty-state onboarding card already explains what's needed.
    if (nueva.length > 0) {
      const linkedPDFs = nueva.filter(g => g.tienePDF).length
      const ocrCount = 0  // processFiles does not run OCR; only handleDrop does.
      showModal({
        type: 'success',
        title: 'Carpeta cargada',
        subtitle: `Procesamos ${xmls.length} archivo${xmls.length === 1 ? '' : 's'} correctamente.`,
        stats: [
          { value: `+${nueva.length}`,                 label: 'Facturas nuevas', color: '#59D39B' },
          { value: xmls.length,                        label: 'XMLs leídos',     color: 'rgba(255,255,255,0.85)' },
          ...(linkedPDFs > 0 ? [{ value: linkedPDFs,   label: 'PDFs vinculados', color: 'rgba(255,255,255,0.85)' }] : []),
          ...(ocrCount   > 0 ? [{ value: ocrCount,     label: 'OCR IA',          color: '#f59e0b' }] : []),
        ],
        primaryLabel: 'Continuar',
      })
      setCarpetaSuccess(true)
      setTimeout(() => setCarpetaSuccess(false), 2500)
    }
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
    const monedaCode = (parsed.moneda || 'MXN').toString().toUpperCase()
    const isExtranjera = !!monedaCode && monedaCode !== 'MXN'
    const today = new Date().toISOString().slice(0, 10)
    const subtotal = Number(parsed.subtotal) || 0
    const iva = Number(parsed.iva) || 0
    const total = Number(parsed.total) || 0
    const propina = Number(parsed.propina) || 0
    const propinaSugerida18 = Number(parsed.propinaSugerida18) || 0
    const propinaSugerida20 = Number(parsed.propinaSugerida20) || 0
    const propinaSugerida22 = Number(parsed.propinaSugerida22) || 0

    return {
      id: genId(),
      rfc: '',
      proveedor: parsed.proveedor || '',
      // Prefer the merchant's own folio when present, fall back to the card
      // approval code so it lines up with the bank statement's "Código de
      // autorización" column in Pass 0 of validarBanco.
      noFactura: (parsed.folio || parsed.approval_code)
        ? String(parsed.folio || parsed.approval_code)
        : ('TKT-' + uuid.slice(0, 6).toUpperCase()),
      fechaFac: parsed.fecha || today,
      concepto: parsed.concepto || '',
      tipo: autoDetectTipo(parsed.proveedor || '', parsed.concepto || '', colaborador?.categoria),
      // For foreign-currency tickets leave the MXN side at 0 — Pass 0 of
      // validarBanco will fill it from the bank statement's "Monto en MXN"
      // column once the authorization code matches.
      importe:        isExtranjera ? 0 : subtotal,
      iva:            isExtranjera ? 0 : iva,
      isrTrasladado:  0,
      retencionISR:   0,
      retencionIVA:   0,
      retenciones:    0,
      totalCFDI:      isExtranjera ? 0 : total,
      propinaPorcentaje: 0,
      montoPropina:   isExtranjera ? 0 : propina,
      fechaCobro: parsed.fecha || today,
      formaPago: parsed.formaPago || '04',
      uuid,
      tienePDF: false,
      pdfFile: null,
      xmlFile: null,
      hizoMatch: false,
      validado: false,
      // Foreign-currency-aware fields. montoExtranjero / propinaExtranjero
      // generalize the older USD-only fields; we keep montoUSD/moneda in
      // sync so the existing MONTO USD table column, USD KPI card, badge
      // and Excel export keep working without a wider refactor.
      montoExtranjero:    isExtranjera ? total   : 0,
      propinaExtranjero:  isExtranjera ? propina : 0,
      monedaCodigo:       monedaCode,
      esMonedaExtranjera: isExtranjera,
      montoUSD:           isExtranjera ? total : 0,
      tipoCambio:         0,
      moneda:             monedaCode,
      // Flag OCR-derived rows so validarBanco's Pass 0 can match them by
      // authorization code against the Clara bank statement.
      esTicket: true,
      // Subtotal + Suggested Gratuity amounts from the OCR — fuel
      // smartAmountMatch so the bank charge (which usually equals subtotal +
      // tip) lines up against the right ticket even when no propina row was
      // OCR'd separately.
      subtotal,
      propinaSugerida18,
      propinaSugerida20,
      propinaSugerida22,
    }
  }

  // First-class photo upload for foreign expenses. Mobile-camera-aware
  // (capture="environment" on the <input>) and compresses to ≤2000px wide
  // JPEG before sending to OCR so 5–10 MB phone shots don't blow Vercel's
  // 4.5 MB serverless body cap.
  const cargarFoto = async (e) => {
    const files = Array.from(e.target.files || [])
    if (!files.length) return

    const confirmed = await askConfirm({
      type: 'confirm',
      title: 'Procesar con OCR',
      subtitle: `Se procesarán ${files.length} ${files.length === 1 ? 'archivo' : 'archivos'} con IA. Esto consume créditos de Claude API. ¿Continuar?`,
      primaryLabel: 'Continuar',
      secondaryLabel: 'Cancelar',
    })
    if (!confirmed) { e.target.value = ''; return }

    setOcrLoading(true)
    const newGastos = []
    for (const file of files) {
      try {
        const fileForOCR = await compressImage(file)
        const base64 = await fileToBase64(fileForOCR)
        const mediaType = fileForOCR.type || 'image/jpeg'
        const gasto = await extractReceiptData(base64, mediaType, file.name)
        if (gasto) {
          gasto.esMonedaExtranjera = !!(gasto.moneda && gasto.moneda !== 'MXN')
          // Stash the (compressed) image data URL so the ZIP export can
          // wrap it into a single-page PDF named via buildFileName.
          if (mediaType.startsWith('image/')) {
            gasto.imageDataURL = `data:${mediaType};base64,${base64}`
            gasto.originalFileName = file.name
          }
          newGastos.push(gasto)
        }
      } catch (err) {
        console.warn('OCR error for', file.name, err)
      }
    }
    setOcrLoading(false)
    if (newGastos.length) {
      setLista(prev => [...prev, ...newGastos])
      setToast(`✓ ${newGastos.length} ${newGastos.length === 1 ? 'foto procesada' : 'fotos procesadas'} con OCR`)
    } else {
      setToast('No se pudo extraer información de las fotos')
    }
    e.target.value = ''
  }

  const handleDrop = async (e) => {
    e.preventDefault()
    e.stopPropagation()
    setIsDraggingOver(false)

    const files = Array.from(e.dataTransfer.files)
    if (!files.length) return

    const xmlFiles   = files.filter(f => f.name.toLowerCase().endsWith('.xml'))
    const pdfFiles   = files.filter(f => f.name.toLowerCase().endsWith('.pdf'))
    // Extended image set: phones drop .heic by default on iOS and Android
    // exports vary across .webp/.bmp/.gif. compressImage re-encodes them
    // to JPEG via canvas before OCR, so all of these land at the endpoint
    // as image/jpeg regardless of source format.
    const imageFiles = files.filter(f => /\.(jpe?g|png|webp|heic|heif|bmp|gif)$/i.test(f.name.toLowerCase()))

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
      const userConfirmed = await askConfirm({
        type: 'warning',
        title: 'Procesar archivos sin XML',
        subtitle: ocrFiles.length === 1
          ? `Se detectó 1 archivo sin XML: ${ocrFiles[0].name}`
          : `Se detectaron ${ocrFiles.length} archivos sin XML.`,
        stats: [
          { value: ocrFiles.length, label: 'Archivos OCR' },
          { value: `~$${(ocrFiles.length * 0.01).toFixed(2)}`, label: 'Costo estimado' },
        ],
        primaryLabel: 'Procesar con IA',
        secondaryLabel: 'Cancelar',
      })

      if (userConfirmed) {
        setOcrLoading(true)
        for (const file of ocrFiles) {
          try {
            // compressImage re-encodes images to a ≤2000px-wide JPEG so
            // phone shots don't blow Vercel's 4.5 MB body cap; PDFs pass
            // through untouched. mediaType derives from the compressed
            // File's type when present, else from the original extension.
            const fileForOCR = await compressImage(file)
            const base64 = await fileToBase64(fileForOCR)
            const ext = file.name.split('.').pop().toLowerCase()
            const mediaType = fileForOCR.type
              || (ext === 'pdf' ? 'application/pdf' : `image/${ext === 'jpg' ? 'jpeg' : ext}`)

            const gasto = await extractReceiptData(base64, mediaType, file.name)
            if (gasto) {
              // Keep the source PDF attached so it rides the ZIP export.
              // Reuse the base64 we already computed for the OCR call
              // instead of re-reading the file.
              if (mediaType === 'application/pdf') {
                gasto.pdfFile = file
                gasto.pdfDataURL = `data:application/pdf;base64,${base64}`
                gasto.tienePDF = true
              }
              // For OCR'd images, stash the data URL so exportar can wrap
              // it into a single-page PDF named via buildFileName.
              if (mediaType.startsWith('image/')) {
                gasto.imageDataURL = `data:${mediaType};base64,${base64}`
                gasto.originalFileName = file.name
              }
              gasto.isNew = true
              allNewGastos.push(gasto)
            }
          } catch (err) {
            console.warn('OCR error:', file.name, err)
            showModal({
              type: 'error',
              title: 'Error de OCR',
              subtitle: `${file.name}\n\n${err.message}`,
              primaryLabel: 'Entendido',
            })
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
        showModal({
          type: 'warning',
          title: 'Archivo no soportado',
          subtitle: 'Solo se aceptan archivos XML, PDF, JPG, PNG o WEBP.',
          primaryLabel: 'Entendido',
        })
      }
      return
    }

    // Persist PDF bytes as data URLs so the ZIP export doesn't depend on
    // the original File reference. OCR rows already set pdfDataURL above;
    // XML-paired PDFs (parseCFDI auto-match or the manual second pass)
    // still hold pdfFile only, so convert them here.
    for (const g of allNewGastos) {
      if (g.pdfFile && !g.pdfDataURL) {
        try { g.pdfDataURL = await fileToDataURL(g.pdfFile) }
        catch (err) { console.warn('PDF read failed:', g.pdfFile.name, err) }
      }
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
    // Read as bytes and pick the encoding: try strict UTF-8 first, fall back
    // to ISO-8859-1 (latin-1) which is what Clara USA exports. Without this,
    // accented chars (Código, Autorización, etc.) come back as mojibake.
    const buffer = await file.arrayBuffer()
    let raw
    try {
      raw = new TextDecoder('utf-8', { fatal: true }).decode(buffer)
    } catch {
      raw = new TextDecoder('iso-8859-1').decode(buffer)
    }
    const content = raw.replace(/\r\n/g, '\n').replace(/\r/g, '\n')
    const lines = content.split('\n')
    const sample = content.slice(0, 500)
    let sep = ','
    if (!sample.includes(',')) sep = sample.includes(';') ? ';' : '\t'

    // Clara-platform CSVs have a fixed column layout. Detect by any of the
    // header markers we know about — the MX and USA variants share most
    // columns, plus the USA file carries "Código de autorización" (col 12)
    // and "Moneda original" (col 4) which power the ticket Pass 0 match.
    // "Transacci"/"autorizaci" intentionally drop the accented chars so we
    // survive encoding mishaps regardless of UTF-8 vs latin-1.
    const headerLine = lines[0] || ''
    const isClara = headerLine.includes('Fecha de Transacci')
                 || headerLine.includes('digo de autorizaci')
                 || headerLine.includes('Moneda original')

    let matches = 0, propinas = 0
    const sinFactura = []
    const matchedRows = []
    // Reset hizoMatch for the new run, but seed fechaCobro from fechaFac
    // (the new default). Matched rows will overwrite this with the CSV's
    // dCSV; unmatched rows keep the invoice date so the FECHA DE COBRO
    // column never shows a blank.
    const nl = lista.map(g => ({ ...g, hizoMatch: false, fechaCobro: g.fechaFac || '' }))

    // Normalize for fuzzy-suggestion comparison: strip accents, lowercase,
    // squash non-alphanumerics. Sørensen-Dice on character bigrams gives a
    // reasonable 0..1 similarity that's tolerant of word-order swaps.
    const normalizeText = s => (s || '').toString().toLowerCase()
      .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
      .replace(/[^a-z0-9]+/g, ' ').trim()
    const bigramSet = s => {
      const t = normalizeText(s).replace(/\s/g, '')
      const set = new Set()
      for (let i = 0; i < t.length - 1; i++) set.add(t.slice(i, i + 2))
      return set
    }
    const nameOverlap = (a, b) => {
      const A = bigramSet(a), B = bigramSet(b)
      if (!A.size || !B.size) return 0
      let common = 0
      for (const g of A) if (B.has(g)) common++
      return common / Math.min(A.size, B.size)
    }

    // Snapshot the bank ↔ invoice pairing at apply time so the modal can
    // render the match list without re-deriving anything from `nl` (which
    // continues to mutate as later passes run).
    const snapshotMatch = (idx, row, pass, method, confidence) => {
      const g = nl[idx]
      matchedRows.push({
        pass,
        method,
        confidence,
        invoiceId: g.id,
        invoiceName: g.proveedor || '(sin proveedor)',
        invoiceNumber: g.noFactura || '',
        invoiceTotal: g.totalCFDI || 0,
        invoiceMoneda: g.moneda || 'MXN',
        isTicket: !!g.esTicket,
        csvDate: formatCobro(row.dCSV),
        csvDescripcion: row.descripcion || '',
        csvAmountMXN: row.montoMXN || 0,
        csvAmountUSD: row.montoUSD || 0,
        csvMoneda: row.moneda || 'MXN',
        csvAuth: row.autorizacion || '',
      })
    }
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

      // Will be filled for Clara rows (used by Pass 0 + USD auto-fill +
      // tip-eligibility gating).
      let autorizacion = ''
      let moneda = 'MXN'
      let montoUSD = 0
      let montoMXN = 0
      let categoria = ''

      if (isClara) {
        dCSV = parseDateRobusto(cols[0] || '')
        montoMXN = cleanNum(cols[5]) || 0
        montoUSD = cleanNum(cols[3]) || 0
        moneda = (cols[4] || 'MXN').trim().toUpperCase() || 'MXN'
        autorizacion = String(cols[12] || '').trim()
        categoria = (cols[13] || '').toString().trim()
        const amount = montoMXN || montoUSD || 0
        if (!dCSV || !amount) continue
        // Carry BOTH the MXN and USD figures into the amount candidates so
        // smartAmountMatch can find ticket receipts whose OCR was captured in
        // either currency without depending on which column is non-zero.
        amounts = [montoMXN, montoUSD].filter(v => v > 0)
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
      csvRows.push({ dCSV, amounts, descripcion, matched: false, autorizacion, moneda, montoUSD, montoMXN, categoria })
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
            if (predicate(nl[i], monto, row)) candidates.push(i)
          }
          if (!candidates.length) continue
          const idx = closestByDate(candidates, row.dCSV)
          apply(idx, monto, row.dCSV, row)
          row.matched = true
          break
        }
      }
    }

    // Pass 0 — TICKET (OCR-derived) rows match the bank statement's
    // "Código de autorización" against their noFactura. Generalized to any
    // foreign currency (not just USD): when bank moneda != 'MXN' we fill
    // importe/totalCFDI from the MXN side, derive tipoCambio = MXN÷foreign,
    // and detect a foreign-currency tip from the delta between the bank's
    // recorded foreign total and the OCR'd ticket subtotal (if the delta
    // sits in the typical 5–35% tip window).
    let ticketsMatched  = 0
    let foreignMatched  = 0
    for (const row of csvRows) {
      if (row.matched) continue
      if (!row.autorizacion) continue
      const idx = nl.findIndex(g => {
        if (g.esTicket && !g.hizoMatch) {
          console.log('Pass0 checking auth:', row.autorizacion, 'vs ticket noFactura:', g.noFactura)
        }
        return !g.hizoMatch &&
          g.esTicket &&
          String(g.noFactura || '').trim() === row.autorizacion
      })
      if (idx === -1) continue

      nl[idx].hizoMatch  = true
      nl[idx].fechaCobro = formatCobro(row.dCSV)
      nl[idx].formaPago  = '04'
      // Pin the FACTURA cell to the bank statement's authorization code so
      // the row visually traces back to the exact line on the CSV, even if
      // the OCR extracted a slightly different format.
      nl[idx].noFactura  = row.autorizacion

      const isExtranjera = row.moneda && row.moneda !== 'MXN' && row.montoUSD > 0
      if (isExtranjera) {
        nl[idx].monedaCodigo       = row.moneda
        nl[idx].moneda             = row.moneda  // keep legacy field in sync
        nl[idx].esMonedaExtranjera = true
        nl[idx].importe   = row.montoMXN
        nl[idx].totalCFDI = row.montoMXN
        nl[idx].tipoCambio = row.montoUSD > 0
          ? +(row.montoMXN / row.montoUSD).toFixed(4)
          : 0

        // Propina detection in foreign currency: bank's foreign total > OCR'd
        // ticket total → the delta IS the tip (sanity-checked 5–35%).
        const ticketTotal = nl[idx].montoExtranjero || 0
        if (row.montoUSD > ticketTotal && ticketTotal > 0) {
          const tipDetected = +(row.montoUSD - ticketTotal).toFixed(2)
          const tipPct = (tipDetected / ticketTotal) * 100
          if (tipPct >= 5 && tipPct <= 35) {
            nl[idx].propinaExtranjero = tipDetected
            // Mirror into the MXN propina field via tipoCambio so the
            // existing Prop $ / Prop % cells light up too.
            nl[idx].montoPropina = +(tipDetected * nl[idx].tipoCambio).toFixed(2)
            nl[idx].propinaPorcentaje = +tipPct.toFixed(2)
          }
        }
        // After-tip foreign total goes into both the new and legacy fields.
        nl[idx].montoExtranjero = row.montoUSD
        nl[idx].montoUSD        = row.montoUSD
        foreignMatched++
      } else if (row.montoMXN > 0) {
        nl[idx].totalCFDI = row.montoMXN
        nl[idx].importe   = row.montoMXN
      }
      snapshotMatch(idx, row, 0, 'Auth Code', 99)
      row.matched = true
      ticketsMatched++
      matches++
    }

    // Pass 1 — smart match, gated by tip eligibility.
    //   OCR tickets carry a real subtotal + Suggested Gratuity from the
    //     receipt; we probe subtotal alone, and — only when the (invoice,
    //     bank row) pair is tip-eligible — subtotal + each suggested tip.
    //   CFDI invoices only have totalCFDI; we probe totalCFDI alone, plus —
    //     when eligible — the common Mexican tip ladder (10..25%).
    // Eligibility uses Clara's "Categoría de Compra" first; if absent (non-
    // Clara CSV), falls back to keyword sniffing on proveedor + concepto.
    // The point is: a gas/uber/hotel CSV line can never bind to an invoice
    // via a phantom 18% tip — only direct totalCFDI matches survive.
    const RESTAURANT_KEYWORDS = [
      'restaurant', 'restaurante', 'bar', 'cafe', 'café', 'cafeteria', 'cafetería',
      'comida', 'aerocomidas', 'taqueria', 'taquería', 'pizzeria', 'pizzería',
      'grill', 'steakhouse', 'bistro', 'parrilla', 'cantina', 'fonda',
      'sushi', 'cocina', 'kitchen', 'diner', 'eatery', 'jumper',
    ]
    const TIP_ELIGIBLE_CATEGORIES = [
      'alimentos', 'bares y bebidas alcoholicas', 'bares y bebidas alcohólicas',
      'food', 'restaurants', 'restaurantes', 'comida', 'consumo', 'viaticos', 'viáticos',
    ]
    const isEligibleForTip = (inv, csvRow) => {
      const csvCat = (csvRow?.categoria || '').toLowerCase().trim()
      if (csvCat) {
        // CSV has a category column. Trust it definitively in both directions:
        // a food category → eligible; any other category → never tip.
        return TIP_ELIGIBLE_CATEGORIES.some(c => csvCat.includes(c))
      }
      // No CSV category (non-Clara source). Sniff invoice fields for
      // restaurant keywords instead.
      const haystack = ((inv?.proveedor || '') + ' ' + (inv?.concepto || '')).toLowerCase()
      return RESTAURANT_KEYWORDS.some(kw => haystack.includes(kw))
    }
    const asReceipt = inv => ({
      _raw: inv,  // kept so smartAmountMatch can re-check eligibility off proveedor/concepto
      isOCR: (inv.subtotal || 0) > 0 || (inv.propinaSugerida18 || inv.propinaSugerida20 || inv.propinaSugerida22) > 0,
      subtotalOCR: inv.subtotal || 0,
      totalCFDI: inv.totalCFDI || 0,
      importe: inv.importe || 0,
      propinaSugerida18: inv.propinaSugerida18 || 0,
      propinaSugerida20: inv.propinaSugerida20 || 0,
      propinaSugerida22: inv.propinaSugerida22 || 0,
      montoPropina: inv.montoPropina || 0,
    })
    // smartAmountMatch now returns the matched candidate itself (or null)
    // so callers can persist the exact propina the candidate represents
    // instead of computing it from a noisy delta. Tolerance is one cent —
    // pure floating-point safety. We no longer accept "close-enough" matches;
    // if Pass 1 can't bind a candidate within 1¢, the row stays unmatched.
    const smartAmountMatch = (receipt, csvAmount, csvRow, tolerance = 0.01) => {
      const eligible = isEligibleForTip(receipt._raw || receipt, csvRow)
      const candidates = []
      candidates.push({
        amount: receipt.isOCR ? receipt.subtotalOCR : receipt.totalCFDI,
        propina: 0, pct: 0, label: 'exact',
      })
      if (eligible) {
        if (receipt.isOCR) {
          const sub = receipt.subtotalOCR
          if (receipt.propinaSugerida18 > 0) candidates.push({ amount: sub + receipt.propinaSugerida18, propina: receipt.propinaSugerida18, pct: 18, label: 'ocr18' })
          if (receipt.propinaSugerida20 > 0) candidates.push({ amount: sub + receipt.propinaSugerida20, propina: receipt.propinaSugerida20, pct: 20, label: 'ocr20' })
          if (receipt.propinaSugerida22 > 0) candidates.push({ amount: sub + receipt.propinaSugerida22, propina: receipt.propinaSugerida22, pct: 22, label: 'ocr22' })
        } else {
          const total = receipt.totalCFDI
          ;[0.10, 0.12, 0.13, 0.15, 0.18, 0.20, 0.22, 0.25].forEach(p => {
            candidates.push({ amount: total * (1 + p), propina: total * p, pct: p * 100, label: `pct${p * 100}` })
          })
        }
      }
      for (const c of candidates) {
        if (c.amount > 0 && Math.abs(c.amount - csvAmount) <= tolerance) return c
      }
      return null
    }

    // Label → user-facing method + confidence. With a 1¢ tolerance an
    // "exact" hit is genuinely exact (95 → Alta confianza); an OCR
    // suggested-gratuity hit is also rock solid since the bank charge
    // matches a number printed on the receipt (95). The CFDI tip ladder
    // is a heuristic guess at which percentage the user added → 85.
    const methodFromMatch = (matchResult) => {
      if (!matchResult) return { method: 'Smart Amount', confidence: 80 }
      const { label } = matchResult
      if (label === 'exact') return { method: 'Monto exacto', confidence: 95 }
      if (label.startsWith('ocr')) return { method: `Propina sugerida +${label.slice(3)}%`, confidence: 95 }
      if (label.startsWith('pct')) return { method: `Propina +${label.slice(3)}%`, confidence: 85 }
      return { method: 'Smart Amount', confidence: 80 }
    }
    tryPass(
      (inv, m, row) => {
        if (smartAmountMatch(asReceipt(inv), m, row)) return true
        // USD secondary: for invoices flagged as USD / foreign currency, also
        // probe the bank row's montoUSD directly so a Clara USA line whose
        // MXN figure missed (FX drift, rounding) still binds via its USD
        // figure.
        const isUSDInv = inv.moneda === 'USD' || inv.esMonedaExtranjera
        if (isUSDInv && row.montoUSD > 0 && smartAmountMatch(asReceipt(inv), row.montoUSD, row)) return true
        return false
      },
      (idx, m, dCSV, row) => {
        const inv = nl[idx]
        const receipt = asReceipt(inv)
        // Re-run smartAmountMatch to retrieve which candidate fired (primary
        // monto first, then USD secondary). The match must succeed because
        // the predicate just said so.
        let matchResult = smartAmountMatch(receipt, m, row)
        if (!matchResult && row.montoUSD > 0) {
          matchResult = smartAmountMatch(receipt, row.montoUSD, row)
        }
        inv.hizoMatch = true
        inv.fechaCobro = formatCobro(dCSV)
        inv.formaPago = '04'  // bank-matched → card transaction
        // Persist the EXACT propina the matched candidate represents — no
        // longer a computed delta. OCR candidates carry the literal Suggested
        // Gratuity amount; CFDI tip-ladder candidates carry totalCFDI × pct.
        if (matchResult && matchResult.propina > 0) {
          inv.montoPropina = parseFloat(matchResult.propina.toFixed(2))
          inv.propinaPorcentaje = parseFloat(matchResult.pct.toFixed(2))
          propinas++
        }
        const { method, confidence } = methodFromMatch(matchResult)
        snapshotMatch(idx, row, 1, method, confidence)
        matches++
      }
    )

    // Collect unmatched rows for the result modal, attaching the top-2
    // fuzzy-name suggestions from existing gastos for the "¿Quisiste decir...?"
    // hint. Pool is all gastos (matched or not) — the user might have already
    // reconciled a similarly-named row earlier and want to merge.
    for (const row of csvRows) {
      if (row.matched) continue
      const desc = row.descripcion || ''
      const sugerencias = nl
        .map(g => ({ id: g.id, proveedor: g.proveedor || '', score: nameOverlap(desc, g.proveedor) }))
        .filter(s => s.proveedor && s.score >= 0.4)
        .sort((a, b) => b.score - a.score)
        .slice(0, 2)
        .map(s => ({ id: s.id, proveedor: s.proveedor, score: Math.round(s.score * 100) }))
      sinFactura.push({
        fecha: formatCobro(row.dCSV),
        monto: Math.max(...row.amounts),
        descripcion: row.descripcion,
        montoMXN: row.montoMXN || 0,
        montoUSD: row.montoUSD || 0,
        moneda: row.moneda || 'MXN',
        sugerencias,
      })
    }

    // Surface the total foreign-currency matches at the end so the modal can
    // show "X en moneda extranjera vinculadas" — covers Pass 0 hits plus any
    // existing row already flagged esMonedaExtranjera that got matched.
    const foreignMatchesTotal = nl.filter(g => g.esMonedaExtranjera && g.fechaCobro).length

    // Aggregate currency totals (Matched vs Pending) for the top summary
    // bar — split MXN vs USD so the chips stay readable across mixed sheets.
    const totalsMatched = matchedRows.reduce((acc, m) => {
      acc.mxn += m.csvAmountMXN || 0
      acc.usd += m.csvAmountUSD || 0
      return acc
    }, { mxn: 0, usd: 0 })
    const totalsPending = sinFactura.reduce((acc, s) => {
      acc.mxn += s.montoMXN || (s.moneda === 'MXN' ? s.monto : 0)
      acc.usd += s.montoUSD || (s.moneda === 'USD' ? s.monto : 0)
      return acc
    }, { mxn: 0, usd: 0 })

    // Defer setLista — the modal now confirms/cancels the apply step.
    // pendingLista travels with the result object; onConfirm flushes it,
    // onCancel discards it. matchedRows/sinFactura are pre-snapshotted so
    // the modal renders correctly regardless of whether lista was applied.
    setConciliacion({
      bancoRows,
      matches,
      propinas,
      sinFactura,
      matchedRows,
      totalsMatched,
      totalsPending,
      facturasSinCargo: nl.length - matches,
      ticketsMatched,
      foreignMatched: foreignMatchesTotal,
      pendingLista: nl,
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
      const cobro = formatDateDisplay(g.fechaCobro)
      const r  = `${g.rfc}\t${g.proveedor.replace(/\t/g,' ')}\t${g.noFactura}\t${fac}\t${g.concepto.replace(/\t/g,' ')}\t${g.importe.toFixed(2)}\t${g.iva.toFixed(2)}\t${(g.isrTrasladado||0).toFixed(2)}\t${(g.retencionISR||0).toFixed(2)}\t${(g.retencionIVA||0).toFixed(2)}\t${g.retenciones.toFixed(2)}\t${g.totalCFDI.toFixed(2)}\t${(g.montoUSD||0).toFixed(2)}\t${(g.tipoCambio||0).toFixed(2)}\t\t${g.formaPago}\t${cobro}\n`
      const p  = g.montoPropina > 0
        ? `\t${g.proveedor} - PROPINA\t\t${fac}\tPROPINA\t${g.montoPropina.toFixed(2)}\t0.00\t0.00\t0.00\t0.00\t0.00\t${g.montoPropina.toFixed(2)}\t\t\t\t${g.formaPago}\t${cobro}\n`
        : ''
      return [r, p]
    })
    navigator.clipboard.writeText(hdr + rows.join(''))
      .then(() => showModal({
        type: 'success',
        title: '¡Copiado al portapapeles!',
        subtitle: 'Ve a tu Excel, haz clic en la celda donde quieres los datos y presiona Ctrl+V (o Cmd+V en Mac) para pegar.',
        primaryLabel: 'OK',
      }))
      .catch(() => showModal({
        type: 'error',
        title: 'No se pudo copiar',
        subtitle: 'Verifica los permisos del navegador.',
        primaryLabel: 'Entendido',
      }))
  }

  // Canonical filename builder used by the ZIP export. Output shape:
  //   Proveedor Folio Tipo MM-DD-YY
  //   e.g. Aerocomidas 66901114763782 Consumo 03-20-26
  //        Fideicomiso Irrevocable DB1616 CUUMXA110440 Hotel 03-19-26
  //        Grupo Ferreteria Calzada FAC102026491 Herramienta 03-20-26
  // Tipo is the gasto category (Vuelo / Hotel / Transporte / Herramienta /
  // Consumo / \u2026) \u2014 a stable taxonomy that reads better than the raw CFDI
  // concepto string. Falls back to g.concepto if tipo is missing.
  const toTitleCase = (str) => str
    .toLowerCase()
    .split(/\s+/)
    .map(w => w ? w.charAt(0).toUpperCase() + w.slice(1) : '')
    .join(' ')
    .trim()

  const buildFileName = (g) => {
    // Proveedor \u2014 Title Case, max 40 chars, brackets stripped
    const prov = toTitleCase(
      (g.proveedor || 'Proveedor')
        .replace(/[\/\\:*?"<>|()\[\]{}]/g, '')
        .replace(/[-_]+/g, ' ')
        .replace(/\s+/g, ' ')
        .trim()
    ).slice(0, 40)

    // Folio \u2014 preserved exactly as detected
    const folio = (g.noFactura || 'SN')
      .replace(/[\/\\:*?"<>|()\[\]{}]/g, '')
      .trim()

    // Concepto \u2014 primera l\u00ednea de la descripci\u00f3n del CFDI (g.concepto ya viene
    // recortado a la primera l\u00ednea desde parseCFDI). Title Case, m\u00e1x 25 chars
    // para que el renombre no se haga gigante. Reemplaza al antiguo "tipo"
    // (categor\u00eda) en el nombre del archivo.
    const rawConcepto = (g.concepto || g.tipo || 'Gasto')
      .replace(/[\/\\:*?"<>|()\[\]{}]/g, '')
      .replace(/\s+/g, ' ')
      .trim()
    let concepto = toTitleCase(rawConcepto).slice(0, 25).trim()
    // Si el corte deja un guion / coma final, lo limpiamos para que el nombre
    // no termine con caracter colgante.
    concepto = concepto.replace(/[\s\-,.;:]+$/, '') || 'Gasto'

    // Fecha \u2014 MM-DD-YY (formato del nombre de archivo, distinto al de pantalla)
    const f = (g.fechaFac || '').split('-')
    const fecha = f.length === 3
      ? `${f[1].padStart(2,'0')}-${f[2].padStart(2,'0')}-${f[0].slice(-2)}`
      : 'SN'

    return `${prov} ${folio} ${concepto} ${fecha}`
  }

  // ── Exportar ZIP con Excel + carpeta Facturas/ ──
  // Output: SMTO_Gastos_<Colab>_<YYYYMMDD>.zip containing
  //   • Reporte_<Colab>_<YYYYMMDD>.xlsx  (fetched from /api/export-excel)
  //   • Facturas/Proveedor Folio Tipo MM-DD-YY .xml + .pdf for every row
  // PDFs are read from gasto.pdfDataURL (intake-time data URL) so the
  // export doesn't depend on the original File handle, which can be
  // unreliable later. Falls back to pdfFile.arrayBuffer() for any gasto
  // that pre-dates the data-URL change in this session.
  const exportar = async () => {
    const zip = new JSZip()
    const facturas = zip.folder('Facturas')

    const colabSlug = (colaborador?.nombre || 'SMTO')
      .replace(/[^a-zA-Z0-9]/g, '_')
      .replace(/_+/g, '_')
      .replace(/^_|_$/g, '')
    const today = new Date().toISOString().slice(0, 10).replace(/-/g, '')

    // 1) Pull the Excel report from the server-side renderer. Payload slim
    //    para no rebasar el límite de 4.5 MB de Vercel con muchas facturas.
    const gastosSlim = lista.map(g => ({
      rfc:              g.rfc || '',
      proveedor:        g.proveedor || '',
      tipo:             g.tipo || '',
      noFactura:        g.noFactura || '',
      fechaFac:         g.fechaFac || '',
      fechaCobro:       g.fechaCobro || '',
      concepto:         g.concepto || '',
      importe:          Number(g.importe) || 0,
      iva:              Number(g.iva) || 0,
      retenciones:      Number(g.retenciones) || 0,
      totalCFDI:        Number(g.totalCFDI) || 0,
      formaPago:        g.formaPago || '',
      montoUSD:         Number(g.montoUSD) || 0,
      montoExtranjero:  Number(g.montoExtranjero) || 0,
      tipoCambio:       Number(g.tipoCambio) || 0,
      montoPropina:     Number(g.montoPropina) || 0,
      propinaExtranjero:Number(g.propinaExtranjero) || 0,
      moneda:           g.moneda || '',
      monedaCodigo:     g.monedaCodigo || '',
    }))
    try {
      const response = await fetch('/api/export-excel', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ gastos: gastosSlim, colaborador: colaborador?.nombre || '' }),
      })
      if (response.ok) {
        const excelBlob = await response.blob()
        zip.file(`Reporte_${colabSlug}_${today}.xlsx`, excelBlob)
      }
    } catch (err) {
      console.warn('Excel fetch failed for ZIP:', err)
    }

    // 2) Rename each gasto's XML + PDF through buildFileName.
    let r = 0
    for (const g of lista) {
      if (g.imageDataURL && !g.pdfDataURL) {
        try {
          g.pdfDataURL = await imageToPDF(g.imageDataURL)
        } catch (err) {
          console.warn('Could not convert image to PDF for', g.proveedor, err)
        }
      }
      if (g.xmlContent || g.pdfDataURL) {
        const nom = buildFileName(g)
        if (g.xmlContent) {
          facturas.file(`${nom}.xml`, g.xmlContent)
          r++
        }
        if (g.pdfDataURL) {
          const base64Data = g.pdfDataURL.split(',')[1]
          const binaryStr = atob(base64Data)
          const bytes = new Uint8Array(binaryStr.length)
          for (let i = 0; i < binaryStr.length; i++) bytes[i] = binaryStr.charCodeAt(i)
          facturas.file(`${nom}.pdf`, bytes, { binary: true })
          if (!g.xmlContent) r++
        }
        if (g.xmlFile && !g.xmlContent) {
          facturas.file(`${nom}.xml`, await g.xmlFile.arrayBuffer())
          r++
        }
        if (g.pdfFile && !g.pdfDataURL) {
          facturas.file(`${nom}.pdf`, await g.pdfFile.arrayBuffer())
          if (!g.xmlFile && !g.xmlContent) r++
        }
      }
    }
    void r

    // 3) Generate + trigger download.
    const zipBlob = await zip.generateAsync({ type: 'blob' })
    const zipName = `SMTO_Gastos_${colabSlug}_${today}.zip`
    const a = Object.assign(document.createElement('a'), {
      href: URL.createObjectURL(zipBlob),
      download: zipName,
    })
    a.click()
    URL.revokeObjectURL(a.href)

    // 4) Success modal — generic PremiumModal with three stats.
    const xmlCount = lista.filter(g => g.xmlContent || g.xmlFile).length
    const pdfCount = lista.filter(g => g.pdfDataURL || g.pdfFile).length
    showModal({
      type: 'success',
      title: '¡ZIP Generado!',
      subtitle: 'Paquete descargado con Excel + facturas renombradas.',
      stats: [
        { value: xmlCount, label: 'XMLs',   color: '#59D39B' },
        { value: pdfCount, label: 'PDFs',   color: '#60a5fa' },
        { value: formatBytes(zipBlob.size), label: 'Tamaño', color: 'rgba(255,255,255,0.6)' },
      ],
      primaryLabel: 'Listo',
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
        showModal({
          type: 'warning',
          title: 'Archivo vacío',
          subtitle: 'No se encontraron registros válidos en el archivo Excel.',
          primaryLabel: 'Entendido',
        })
        return
      }

      // Ask user: append or replace. Primary = append, secondary = replace
      // (matches the old window.confirm pattern where Cancel meant replace).
      const action = await askConfirm({
        type: 'warning',
        title: 'Importar al reporte',
        subtitle: `Se encontraron ${gastos.length} registros en el archivo.`,
        stats: [
          { value: gastos.length, label: 'Encontrados' },
          { value: lista.length,  label: 'Actuales' },
        ],
        primaryLabel: 'Agregar al actual',
        secondaryLabel: 'Reemplazar todo',
      })

      // Tag every imported row as isNew so each one picks up the green-flash
      // row-fade-in animation; flag is cleared after the animation settles.
      const tagged = gastos.map(g => ({ ...g, isNew: true }))

      if (action) {
        setLista(prev => [...prev, ...tagged])
      } else {
        setLista(tagged)
      }

      // Premium success modal (count-up + progress bar + glow). Errors
      // still route through the generic PremiumModal via showModal().
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
      showModal({
        type: 'error',
        title: 'Error al importar Excel',
        subtitle: err.message,
        primaryLabel: 'Entendido',
      })
    }
  }

  // ── Exportar a Excel ──
  // Defers to /api/export-excel (Python serverless function on Vercel) so the
  // .xls file gets full template formatting via xlrd + xlutils + xlwt —
  // something community SheetJS can't preserve in the browser.
  const exportarExcel = async () => {
    if (!lista.length) return
    try {
      closeModal()
      // SLIM el payload: el endpoint sólo consume 19 campos numéricos/texto
      // por gasto. Si mandamos `lista` tal cual, cada gasto puede traer
      // imageDataURL (base64, ~150KB), pdfDataURL, XML crudo, etc. Con 100+
      // facturas eso supera el límite de 4.5MB de Vercel y devuelve 413.
      const gastosSlim = lista.map(g => ({
        rfc:              g.rfc || '',
        proveedor:        g.proveedor || '',
        tipo:             g.tipo || '',
        noFactura:        g.noFactura || '',
        fechaFac:         g.fechaFac || '',
        fechaCobro:       g.fechaCobro || '',
        concepto:         g.concepto || '',
        importe:          Number(g.importe) || 0,
        iva:              Number(g.iva) || 0,
        retenciones:      Number(g.retenciones) || 0,
        totalCFDI:        Number(g.totalCFDI) || 0,
        formaPago:        g.formaPago || '',
        montoUSD:         Number(g.montoUSD) || 0,
        montoExtranjero:  Number(g.montoExtranjero) || 0,
        tipoCambio:       Number(g.tipoCambio) || 0,
        montoPropina:     Number(g.montoPropina) || 0,
        propinaExtranjero:Number(g.propinaExtranjero) || 0,
        moneda:           g.moneda || '',
        monedaCodigo:     g.monedaCodigo || '',
      }))
      const response = await fetch('/api/export-excel', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ gastos: gastosSlim, colaborador: colaborador?.nombre || '' }),
      })
      if (!response.ok) {
        try {
          const errData = await response.json()
          showModal({
            type: 'error',
            title: 'Error al generar Excel',
            subtitle: (errData.error || `API error ${response.status}`) + (errData.trace ? `\n\n${(errData.trace || '').slice(0, 300)}` : ''),
            primaryLabel: 'Entendido',
          })
        } catch {
          showModal({
            type: 'error',
            title: 'Error al generar Excel',
            subtitle: `API error ${response.status}`,
            primaryLabel: 'Entendido',
          })
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
      showModal({
        type: 'error',
        title: 'Error al generar Excel',
        subtitle: err.message,
        primaryLabel: 'Entendido',
      })
    }
  }

  // ── Abrir PDF en nueva pestaña ──
  const openPDF = pdfFile => {
    if (!pdfFile) return
    try {
      const url = URL.createObjectURL(pdfFile)
      const win = window.open(url, '_blank')
      if (!win) showModal({
        type: 'warning',
        title: 'No se pudo abrir el PDF',
        subtitle: 'Verifica que el navegador permita ventanas emergentes para este sitio.',
        primaryLabel: 'Entendido',
      })
    } catch (err) {
      showModal({
        type: 'error',
        title: 'Error al abrir el PDF',
        subtitle: err && err.message ? err.message : String(err),
        primaryLabel: 'Entendido',
      })
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
          <h1 className="header-title">Reporte de Gastos SMTO<span className="version-badge">v7.46</span></h1>
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
          <PremiumButton
            title={carpetaSuccess ? 'Cargado' : 'Cargar Carpeta'}
            variant={carpetaSuccess ? 'success' : 'primary'}
            onClick={() => folderRef.current?.click()}
            icon={carpetaSuccess ? (
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
                <polyline points="20 6 9 17 4 12"/>
              </svg>
            ) : (
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <path d="M22 19a2 2 0 01-2 2H4a2 2 0 01-2-2V5a2 2 0 012-2h5l2 3h9a2 2 0 012 2z"/>
              </svg>
            )}
          />
          <PremiumButton title="Cargar Foto"    icon="📸" variant="secondary" onClick={() => photoRef.current?.click()} />
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
      <input
        ref={photoRef}
        type="file"
        accept="image/*,application/pdf"
        capture="environment"
        multiple
        style={{ display: 'none' }}
        onChange={cargarFoto}
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

      {/* ─── PREMIUM MODAL (info / warning / error / confirm) ─── */}
      {modal && (
        <PremiumModal
          open={true}
          type={modal.type || 'success'}
          title={modal.title}
          subtitle={modal.subtitle}
          stats={modal.stats}
          primaryLabel={modal.primaryLabel || 'Continuar'}
          secondaryLabel={modal.secondaryLabel}
          onPrimary={() => {
            if (modal.onPrimary) modal.onPrimary()
            else closeModal()
          }}
          onSecondary={() => {
            if (modal.onSecondary) modal.onSecondary()
            else closeModal()
          }}
        />
      )}

      {/* ─── MODAL CONCILIACIÓN BANCARIA (premium glass) ─── */}
      <AnimatePresence>
        {conciliacion && (
          <ConciliacionModal
            data={conciliacion}
            onClose={() => setConciliacion(null)}
            onConfirm={() => {
              const count = conciliacion.matches || 0
              if (conciliacion.pendingLista) setLista(conciliacion.pendingLista)
              setConciliacion(null)
              setToast(`✓ Conciliación aplicada · ${count} ${count === 1 ? 'cargo vinculado' : 'cargos vinculados'}`)
            }}
            onCancel={() => setConciliacion(null)}
            onAgregarManual={() => {
              // Apply pending matches first so they aren't wiped when
              // agregarManual mutates lista. Then close + add a manual row.
              if (conciliacion.pendingLista) setLista(conciliacion.pendingLista)
              setConciliacion(null)
              agregarManual()
            }}
          />
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

      {/* ─── INLINE TOAST ─── */}
      <AnimatePresence>
        {toast && (
          <motion.div
            className="smto-toast"
            initial={{ opacity: 0, y: 24, scale: 0.96 }}
            animate={{ opacity: 1, y: 0, scale: 1 }}
            exit={{ opacity: 0, y: 12, scale: 0.98 }}
            transition={{ type: 'spring', stiffness: 280, damping: 24 }}
          >
            {toast}
          </motion.div>
        )}
      </AnimatePresence>
    </div>
  )
}
