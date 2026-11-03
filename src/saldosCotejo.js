// src/saldosCotejo.js
// Cotejo de los gastos de los colaboradores especiales (Alejandro Olivar,
// Victor Aceves, …) contra su archivo "Saldos AAAA". Vive aparte de App.jsx
// para poder probarlo con archivos reales sin levantar la UI.
import * as XLSX from 'xlsx'

/* Parser del Excel "Saldos AAAA" usado por los colaboradores especiales.
   Cada pestaña suele traer headers Fecha|Tipo|Folio|Factura|Concepto|
   Ingreso|Egreso en alguna fila (varía por hoja: r3 a r177). Algunas
   pestañas legacy (Banamex, Intercam) tienen columna 'Desc' en vez de
   'Tipo' y sin 'Folio'; las parseamos igual aunque su folio quede
   vacío.

   Devuelve [{ sheet, fecha, tipo, folio, factura, concepto, egreso }]. */
/* Convierte la fecha cruda del Saldos (Date | número serial | string) al
   formato interno YYYY-MM-DD que usa la app para fechaCobro. */
export function saldosFechaToIso(d) {
  if (!d && d !== 0) return ''
  let dt
  if (d instanceof Date) dt = d
  else if (typeof d === 'number') dt = new Date(Date.UTC(1899, 11, 30) + d * 86400000)
  else {
    const s = String(d).trim()
    const iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/)
    if (iso) return `${iso[1]}-${iso[2]}-${iso[3]}`
    const dmy = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/)
    if (dmy) {
      let [, dd, mm, yy] = dmy
      if (yy.length === 2) yy = '20' + yy
      return `${yy}-${mm.padStart(2, '0')}-${dd.padStart(2, '0')}`
    }
    return ''
  }
  if (isNaN(dt)) return ''
  const yyyy = dt.getUTCFullYear()
  const mm = String(dt.getUTCMonth() + 1).padStart(2, '0')
  const dd = String(dt.getUTCDate()).padStart(2, '0')
  return `${yyyy}-${mm}-${dd}`
}

export const leerSaldosWorkbook = (arrayBuffer) =>
  XLSX.read(arrayBuffer, { type: 'array', cellDates: true })

// Acepta el ArrayBuffer del archivo o el workbook ya leído (leerSaldosWorkbook),
// para no parsear dos veces un Saldos de varios MB.
export function parseSaldosXLSX(arrayBufferOrWb) {
  const wb = arrayBufferOrWb?.SheetNames ? arrayBufferOrWb : leerSaldosWorkbook(arrayBufferOrWb)
  const rows = []
  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName]
    if (!ws || !ws['!ref']) continue
    const range = XLSX.utils.decode_range(ws['!ref'])
    const getCell = (r, c) => ws[XLSX.utils.encode_cell({ r, c })]?.v
    // Texto FORMATEADO de la celda (.w), con fallback al valor crudo (.v). Se
    // usa para el folio de la póliza: Excel guarda "0036234032" como número
    // 36234032 pero lo MUESTRA con ceros a la izquierda vía formato; .v los
    // perdería, .w los conserva tal como se ven.
    const getCellText = (r, c) => {
      const cell = ws[XLSX.utils.encode_cell({ r, c })]
      if (!cell) return ''
      return String(cell.w ?? cell.v ?? '')
    }

    // Encuentra TODAS las filas de header en la pestaña. Algunos archivos
    // de Saldos apilan varias cuentas verticalmente en una sola pestaña
    // (ej. "BBVA MXN Cheques" trae Concentradora, Cheques SMTO Pesos, etc.).
    // Cada header inicia un bloque de datos que corre hasta el siguiente
    // header (o hasta el final de la pestaña).
    const headers = []
    for (let r = range.s.r; r <= range.e.r; r++) {
      const v = (c) => String(getCell(r, c) || '').trim()
      if (v(0) === 'Fecha' && v(1) === 'Tipo' && v(2) === 'Folio' && v(3) === 'Factura') {
        headers.push({ row: r, cols: { fecha: 0, tipo: 1, folio: 2, factura: 3, concepto: 4, ingreso: 5, egreso: 6 } })
      } else if (v(0) === 'Fecha' && (v(1) === 'Desc' || v(1) === 'Descripción') && v(3) === 'Factura') {
        headers.push({ row: r, cols: { fecha: 0, tipo: 1, folio: null, factura: 3, concepto: 4, ingreso: 5, egreso: 6 } })
      }
    }
    if (headers.length === 0) continue

    for (let h = 0; h < headers.length; h++) {
      const { row: hRow, cols } = headers[h]
      const endRow = h + 1 < headers.length ? headers[h + 1].row - 1 : range.e.r
      for (let r = hRow + 1; r <= endRow; r++) {
        const factura = String(getCell(r, cols.factura) || '').trim()
        const tipo    = String(getCell(r, cols.tipo) || '').trim()
        const concepto = String(getCell(r, cols.concepto) || '').trim()
        // Skip filas completamente vacías (sin tipo, factura, ni concepto)
        if (!factura && !tipo && !concepto) continue
        // Las filas con factura 'NA' SI se conservan — el cotejo principal
        // (por factura) las ignora, pero el fallback de cotejo (por
        // proveedor + monto + año-mes) las usa para asignar banco a gastos
        // cuyo registro en Saldos no trae folio explícito.
        const folio = cols.folio !== null ? getCellText(r, cols.folio).trim() : ''
        // Detección de USD: barre TODAS las columnas del renglón en busca del
        // acrónimo USD o "dólar/dolar". Algunas pestañas marcan moneda en una
        // columna fuera del set principal (Tipo / Concepto / Factura). La
        // pestaña misma ("Monex USD Cheques", "BBVA USD Cheques") también
        // cuenta: sus renglones casi nunca repiten "USD" en las celdas.
        let moneda = /\bUSD\b/i.test(sheetName) ? 'USD' : 'MXN'
        const colMax = Math.min(range.e.c, 20)
        for (let c = range.s.c; c <= colMax; c++) {
          const v = String(getCell(r, c) || '').toUpperCase()
          if (/\bUSD\b|\bD[ÓO]LAR/.test(v)) { moneda = 'USD'; break }
        }
        rows.push({
          sheet:    sheetName,
          fecha:    getCell(r, cols.fecha) || '',
          tipo,
          folio,
          factura,
          concepto,
          egreso:   Number(getCell(r, cols.egreso)) || 0,
          moneda,
        })
      }
    }
  }
  return rows
}


const MESES = ['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO', 'JULIO',
  'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE']

/* Tipo de cambio USD de cada mes que contabilidad captura en el Saldos: en la
   pestaña de BBVA MXN Cheques, al pie, una fila "Tipo Cambio" con 12 valores
   (promedio del mes) alineados bajo un encabezado Enero…Diciembre. Es el T/C
   con el que contabilidad convierte el reporte (oct-2026 = 18.4), así que el
   Excel lo usa en vez del T/C de la primera factura en dólares.
   Devuelve { '01': 17.486, …, '10': 18.4 } (sólo los meses con valor). */
export function parseTipoCambioSaldos(wb) {
  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName]
    if (!ws || !ws['!ref']) continue
    const range = XLSX.utils.decode_range(ws['!ref'])
    const v = (r, c) => ws[XLSX.utils.encode_cell({ r, c })]?.v
    const txt = (r, c) => String(v(r, c) ?? '').trim().toUpperCase()
    // El encabezado "Enero, Febrero, …" más ABAJO de la pestaña es la tabla del
    // año en curso (el Saldos arrastra las tablas de años anteriores arriba).
    let hdrRow = -1, colEnero = -1
    for (let r = range.e.r; r >= range.s.r && hdrRow < 0; r--) {
      for (let c = range.s.c; c < Math.min(range.e.c, 60); c++) {
        if (txt(r, c) === 'ENERO' && txt(r, c + 1) === 'FEBRERO') { hdrRow = r; colEnero = c; break }
      }
    }
    if (hdrRow < 0) continue
    // La fila "Tipo Cambio" más cercana a ese encabezado (a ±100 renglones).
    let tcRow = -1
    for (let dist = 0; dist <= 100 && tcRow < 0; dist++) {
      for (const r of [hdrRow - dist, hdrRow + dist]) {
        if (r < range.s.r || r > range.e.r) continue
        for (let c = range.s.c; c < colEnero; c++) {
          if (/^TIPO\s*(DE\s*)?CAMBIO$/.test(txt(r, c))) { tcRow = r; break }
        }
        if (tcRow >= 0) break
      }
    }
    if (tcRow < 0) continue
    const out = {}
    MESES.forEach((_, i) => {
      const n = Number(v(tcRow, colEnero + i))
      if (n > 5 && n < 50) out[String(i + 1).padStart(2, '0')] = Math.round(n * 10000) / 10000
    })
    if (Object.keys(out).length) return out
  }
  return {}
}

// Guiones Unicode que Excel/Word meten al autocorregir. "SGG‐5020" del Saldos
// trae U+2010 y sin esto no coincidía con el folio SGG5020 del CFDI.
const GUIONES_UNICODE = /[‐-―−﹘﹣－]/g

/* Normaliza una factura para comparación: sin espacios, guiones, slashes,
   puntos; uppercase. "48 3993 64000" y "48399364000" matchearán. */
export const normFactura = (s) => String(s || '')
  .replace(GUIONES_UNICODE, '-')
  .replace(/[\s\-/.]/g, '')
  .toUpperCase()

/* Normaliza un proveedor/concepto para comparación laxa: sin acentos, sólo
   A-Z0-9, uppercase. Permite buscar el proveedor del gasto dentro del
   concepto del Saldos (que suele traer el nombre embebido en texto libre). */
const normProv = (s) => String(s || '')
  .normalize('NFD').replace(/\p{Diacritic}/gu, '')
  .toUpperCase().replace(/[^A-Z0-9]/g, '')

/* Las facturas que trae el campo Factura de un renglón del Saldos. Un mismo
   pago puede cubrir varias ("A8472, CAGW000250903, 26 48 1731 6009172": agente
   aduanal + almacenaje + pedimento), y antes se normalizaba todo junto en una
   sola cadena que no coincidía con ninguna. Cada factura lleva sus claves: la
   completa y, si termina en "-N", la clave sin ese sufijo (IBG registra
   "A6446-1" el pago de la factura A6446). La completa se conserva porque hay
   folios reales que terminan así ("F-12"). */
function facturasDelRenglon(raw) {
  return String(raw || '')
    .replace(GUIONES_UNICODE, '-')
    .split(/[,;]+/)
    .map(t => t.trim())
    .filter(Boolean)
    .map(t => {
      const keys = [normFactura(t)]
      if (/-\d{1,2}$/.test(t)) {
        const sinSufijo = normFactura(t.replace(/-\d{1,2}$/, ''))
        if (sinSufijo && sinSufijo !== keys[0]) keys.push(sinSufijo)
      }
      return { raw: t, keys }
    })
    .filter(t => t.keys[0] && t.keys[0] !== 'NA')
}

/* Últimos dígitos de una factura normalizada (mínimo 3). Cubre los folios que
   el Saldos captura sin la serie o con otra: "1127" ↔ "I-1127", "371" ↔ "F371",
   "FA868" ↔ "FAC868", "FactAgr-1021049" ↔ "EFactAgr-1021049". Como coincidir
   sólo por dígitos es débil, esos candidatos exigen además monto o proveedor. */
const digitosFinales = (k) => (String(k).match(/(\d{3,})$/) || [])[1] || ''

// Cómo aparece en el concepto del Saldos un proveedor cuyo nombre fiscal no
// viene escrito (el SAT se registra como "SAT ISR Sep", "SAT ISR Ret IVA Sep").
const ALIAS_POR_RFC = {
  'SAT970701NN3': ['SAT'],
}

// Palabras que no identifican a un proveedor por sí solas.
const PALABRAS_GENERICAS = new Set([
  'GRUPO', 'SERVICIOS', 'SERVICIO', 'COMERCIALIZADORA', 'CORPORATIVO', 'OPERADORA',
  'DISTRIBUIDORA', 'INMOBILIARIA', 'CONSTRUCTORA', 'TRANSPORTES', 'PROMOTORA',
  'ADMINISTRADORA', 'ARRENDADORA', 'INGENIERIA', 'INDUSTRIAL', 'INTEGRADORA',
  'SECRETARIA', 'COORDINADORA', 'CONSULTORIA', 'SOLUCIONES', 'EMPRESA',
  'LIMITED', 'HOLDING', 'HOLDINGS', 'INTERNATIONAL', 'PRODUCTS', 'TECHNOLOGY',
  'TECHNOLOGIES', 'ELECTRONICS', 'ELECTRONIC', 'CORPORATION', 'COMPANY',
  'MEXICO', 'INDUSTRIES', 'SOLUTIONS', 'SYSTEMS', 'TRADING',
])

/* ¿Alguna palabra distintiva del proveedor aparece en el concepto? Para las
   facturas extranjeras la razón social casi nunca coincide con como la
   escribe el Saldos: "SCIENTIFIC TECHNOLOGY … DBA SCIENSCOPE INTERNATIONAL" es
   "ScienScope, …"; "ANA TRADING CORP., U.S.A." es "ANA Trading, …" (la
   primera palabra es corta, así que vale junto con la segunda). */
function marcaEnConcepto(proveedor, concepto) {
  const conc = normProv(concepto)
  const palabras = String(proveedor || '')
    .normalize('NFD').replace(/\p{Diacritic}/gu, '')
    .toUpperCase().split(/[^A-Z0-9]+/).filter(Boolean)
  if (palabras.length >= 2 && palabras[0].length < 5 && conc.includes(palabras[0] + palabras[1])) return true
  return palabras.some(w => w.length >= 5 && !PALABRAS_GENERICAS.has(w) && conc.includes(w))
}

// Folios que la app inventa cuando el documento no trae uno legible.
const esFolioProvisional = (k) => /^(TKT|SN|PED)/.test(k)

/* ¿El proveedor del gasto aparece en el concepto del Saldos? El Saldos casi
   nunca trae la razón social completa ("Aerovías, Envíos", "Despachos Cano,
   Ped exp…"), así que también vale su primera palabra distintiva. */
function proveedorEnConcepto(proveedor, concepto) {
  const conc = normProv(concepto)
  const prov = normProv(proveedor)
  if (prov.length >= 4 && conc.includes(prov)) return true
  const primera = String(proveedor || '')
    .normalize('NFD').replace(/\p{Diacritic}/gu, '')
    .toUpperCase().split(/[^A-Z0-9]+/).find(Boolean) || ''
  return primera.length >= 5 && !PALABRAS_GENERICAS.has(primera) && conc.includes(primera)
}

/* Total del gasto en su propia divisa: el Saldos registra el egreso en la
   moneda de la cuenta, y los CFDI en USD traen totalCFDI = 0 (el MXN se
   calcula en el Excel con el T/C), así que se compara contra el nativo. Si el
   gasto ya se prorrateó en un cotejo anterior, cuenta la factura completa. */
export const totalNativo = (g) => {
  const f = g.montosOriginales ? { ...g, ...g.montosOriginales } : g
  return Number(f.esMonedaExtranjera ? (f.montoExtranjero || f.montoUSD) : f.totalCFDI) || 0
}

/* Mes (AAAA-MM) que se está reportando: el más frecuente entre las fechas de
   factura. Cuando una factura se pagó en varias exhibiciones (anticipo en
   septiembre, liquidación en octubre; las 12 domiciliaciones de un seguro), el
   cotejo prefiere el pago que cae en este mes. */
export function mesDelReporte(lista) {
  const cuenta = new Map()
  for (const g of lista) {
    const ym = String(g.fechaFac || '').slice(0, 7)
    if (/^\d{4}-\d{2}$/.test(ym)) cuenta.set(ym, (cuenta.get(ym) || 0) + 1)
  }
  let mejor = '', n = 0
  for (const [ym, c] of cuenta) if (c > n || (c === n && ym > mejor)) { mejor = ym; n = c }
  return mejor
}

/* Puntúa un renglón del Saldos como candidato para un gasto. Decide entre
   varios renglones que traen la misma factura: folios repetidos entre
   proveedores, o la misma factura pagada en varias exhibiciones. Señales de
   más a menos decisiva:
     1. monto exacto del egreso ≈ total del gasto (±$0.01)
     2. el pago cae en el mes del reporte
     3. proveedor del gasto presente en el concepto del Saldos
     4. misma moneda (USD↔USD, MXN↔MXN)
     5. el renglón agrupa otras facturas que también están en el reporte
     6. cercanía de fecha factura↔cobro (desempate suave)
   Devuelve { score, amountMatch, provMatch } para poder marcar después si la
   colisión se resolvió por una señal fuerte (monto/proveedor) o sólo débil. */
function scoreSaldosCandidate(g, sRow, ctx = {}) {
  const gMoneda = (g.monedaCodigo || g.moneda || 'MXN').toString().toUpperCase()
  const sMoneda = (sRow.moneda || 'MXN').toString().toUpperCase()
  const gTotal  = totalNativo(g)
  const sEgreso = Number(sRow.egreso) || 0
  const amountMatch = gTotal > 0 && sEgreso > 0 && Math.abs(sEgreso - gTotal) <= 0.01
  const provMatch = proveedorEnConcepto(g.proveedor, sRow.concepto)
  const sIso = saldosFechaToIso(sRow.fecha) || ''

  const enMes = !!ctx.mes && sIso.slice(0, 7) === ctx.mes
  let score = 0
  if (amountMatch)         score += 10000
  if (enMes)               score += 6000
  if (provMatch)           score += 4000
  if (sMoneda === gMoneda) score += 1000
  score += 300 * (ctx.companeras || 0)
  const gIso = String(g.fechaFac || '').slice(0, 10)
  let dias = Infinity
  if (/^\d{4}-\d{2}-\d{2}$/.test(gIso) && /^\d{4}-\d{2}-\d{2}$/.test(sIso)) {
    const days = Math.abs(new Date(gIso + 'T00:00:00Z') - new Date(sIso + 'T00:00:00Z')) / 86400000
    if (!Number.isNaN(days)) { dias = days; score += Math.max(0, 200 - days) }
  }
  return { score, amountMatch, provMatch, enMes, dias }
}

/* Cotejo: por cada gasto busca el renglón del Saldos con la misma factura
   (o, si no, con los mismos dígitos finales + monto/proveedor), MISMO AÑO.
   Un renglón que paga varias facturas puede vincular a un gasto por cada una.
   Devuelve los matched, sin-match por ambos lados, y marca las discrepancias
   de tipo para que el usuario decida. */
export function cotejarConSaldos(saldosRows, lista) {
  // 1) Conjunto de años presentes en los gastos
  const gastoYears = new Set()
  lista.forEach(g => {
    const y = String(g.fechaFac || '').slice(0, 4)
    if (/^\d{4}$/.test(y)) gastoYears.add(y)
  })

  // 2) Filtrar Saldos: solo renglones cuyo año esté en el set de gastos.
  //    Si los gastos no tienen año detectable, cae al comportamiento previo.
  const filterByYear = gastoYears.size > 0
  const filteredSaldos = filterByYear
    ? saldosRows.filter(row => {
        const yearStr = (saldosFechaToIso(row.fecha) || '').slice(0, 4)
        return /^\d{4}$/.test(yearStr) && gastoYears.has(yearStr)
      })
    : saldosRows

  // 3) Índices por factura normalizada y por dígitos finales → { idx, f }
  //    (renglón + posición de la factura dentro del renglón). Las filas con
  //    factura 'NA' (nóminas/comisiones) no entran: el fallback de abajo las
  //    usa por proveedor + monto.
  const facturasPorRenglon = filteredSaldos.map(row => facturasDelRenglon(row.factura))
  const byFactura = new Map()
  const byDigitos = new Map()
  const largas = []   // claves de 10+ caracteres, para folios truncados

  const push = (map, k, v) => { if (!map.has(k)) map.set(k, []); map.get(k).push(v) }
  facturasPorRenglon.forEach((facturas, idx) => {
    facturas.forEach((fac, f) => {
      for (const k of fac.keys) push(byFactura, k, { idx, f })
      const dig = digitosFinales(fac.keys[fac.keys.length - 1])
      if (dig) push(byDigitos, dig, { idx, f })
      if (fac.keys[0].length >= 10) largas.push({ key: fac.keys[0], idx, f })
    })
  })

  const facturasDelReporte = new Set(lista.map(g => normFactura(g.noFactura)).filter(Boolean))
  const mes = mesDelReporte(lista)
  // Cada factura de cada renglón se consume una sola vez (`idx#f`); un renglón
  // con 3 facturas puede vincular hasta 3 gastos.
  const usados = new Set()
  const renglonUsado = new Set()

  const matched = []
  const sinMatchGastos = []

  // 4) Match eligiendo el MEJOR candidato libre. Primero por factura exacta;
  //    si no hay, por dígitos finales, y al final por folio truncado (el
  //    Saldos cortó "26 48 3450 6002815" a "26 48 3450 600"). Los dos últimos
  //    sólo cuentan con monto o proveedor (o, el truncado, con otra factura
  //    del mismo pago en el reporte).
  lista.forEach((g) => {
    const k = normFactura(g.noFactura)
    const exactos = k ? (byFactura.get(k) || []) : []
    // Estos sólo se evalúan si la factura exacta no dio (o sus candidatos se
    // descartaron: A6467 exacto estaba en un pago de agosto de otro emisor y
    // el pago real viene como "6467-1").
    const porDigitos = digitosFinales(k) ? (byDigitos.get(digitosFinales(k)) || []) : []
    const truncados = k.length >= 10
      ? largas.filter(x => x.key.length < k.length && k.startsWith(x.key)) : []

    let elegido, best = null, candidatosUsados = [], modoElegido = ''
    const evaluar = (cands, modo) => {
      const vistos = new Set()
      for (const c of cands) {
        const key = `${c.idx}#${c.f}`
        if (usados.has(key) || vistos.has(key)) continue
        vistos.add(key)
        const facturas = facturasPorRenglon[c.idx]
        const companeras = facturas.filter((fac, j) =>
          j !== c.f && fac.keys.some(x => facturasDelReporte.has(x))).length
        const cand = scoreSaldosCandidate(g, filteredSaldos[c.idx], { mes, companeras })
        const fuerte = cand.amountMatch || cand.provMatch
        if (modo === 'digitos' && !fuerte) continue
        if (modo === 'truncado' && !fuerte && !companeras) continue
        if (modo === 'concepto') {
          // El folio del gasto (a menudo el PO que leyó el OCR) viene escrito
          // en el concepto del pago ("ICT, P00761 Stencil Inspection…"), o el
          // monto es exacto en la misma moneda y el concepto nombra al proveedor.
          const sRow = filteredSaldos[c.idx]
          const folioEnConcepto = k.length >= 5 && !esFolioProvisional(k) && normProv(sRow.concepto).includes(k)
          const mismaMoneda = (sRow.moneda || 'MXN') === (g.monedaCodigo || g.moneda || 'MXN').toString().toUpperCase()
          if (!folioEnConcepto && !(cand.amountMatch && mismaMoneda && marcaEnConcepto(g.proveedor, sRow.concepto))) continue
        }
        // Un renglón que agrupa varias facturas puede traer el MISMO folio de
        // otro emisor (los folios son por emisor): A6467 de IBG aparecía en un
        // pago de agosto a otro agente aduanal, dos meses antes de la factura.
        // Se acepta si además cae en el mes del reporte, cerca de la fecha de
        // la factura, o con monto/proveedor.
        if (facturas.length > 1 && !fuerte && !cand.enMes && !(cand.dias <= 45)) continue
        if (!best || cand.score > best.score) { best = cand; elegido = c }
      }
      if (elegido && !modoElegido) { candidatosUsados = cands; modoElegido = modo }
    }
    evaluar(exactos, 'exacto')
    if (!elegido) evaluar(porDigitos, 'digitos')
    if (!elegido) evaluar(truncados, 'truncado')
    if (!elegido) {
      // Último intento: renglones de UNA factura cuyo concepto delata al gasto.
      const porConcepto = []
      facturasPorRenglon.forEach((facturas, idx) => {
        if (facturas.length === 1) porConcepto.push({ idx, f: 0 })
      })
      evaluar(porConcepto, 'concepto')
      // Para exhibiciones sólo cuentan los renglones que pasaron el filtro.
      if (elegido) {
        const fac = facturasPorRenglon[elegido.idx][0].keys[0]
        candidatosUsados = porConcepto.filter(c => facturasPorRenglon[c.idx][0].keys[0] === fac)
      }
    }

    if (!elegido) {
      sinMatchGastos.push(g)
      return
    }
    usados.add(`${elegido.idx}#${elegido.f}`)
    renglonUsado.add(elegido.idx)
    const sRow = filteredSaldos[elegido.idx]
    const tipoDiffers = !!sRow.tipo && !!g.tipo &&
      sRow.tipo.toLowerCase() !== g.tipo.toLowerCase()
    // Colisión = había más de un renglón del Saldos con este mismo número de
    // factura. La marcamos (y si se ancló por una señal fuerte) para que el
    // usuario pueda verificar en el modal que se vinculó la correcta.
    const renglonesCandidatos = new Set(candidatosUsados.map(c => c.idx))
    const facturaColision = renglonesCandidatos.size > 1
    const colisionResuelta = best.amountMatch || best.provMatch
    matched.push({
      gastoId: g.id, gasto: g, saldosRow: sRow, tipoDiffers, facturaColision, colisionResuelta,
      // El egreso del renglón sólo equivale al pago de ESTA factura cuando el
      // renglón no agrupa otras; si agrupa, el monto es de todo el grupo.
      multiFactura: facturasPorRenglon[elegido.idx].length > 1,
      // Cuántos pagos de SÓLO esta factura hay en el Saldos. Con 2+ es un pago
      // en exhibiciones (anticipo + liquidación, 12 domiciliaciones) y el
      // reporte lleva la parte de este pago; con uno solo, una diferencia de
      // monto es sólo eso (descuento, redondeo) y no se prorratea.
      exhibiciones: [...renglonesCandidatos].filter(i => facturasPorRenglon[i].length === 1).length,
      // Si se encontró por el concepto, el folio que leyó el OCR no era el de
      // la factura (era el PO o un provisional TKT-): se usa el del Saldos.
      facturaSaldos: modoElegido === 'concepto' ? facturasPorRenglon[elegido.idx][0].raw : null,
    })
  })

  // 5) FALLBACK match: para gastos que no encontraron factura en el Saldos,
  //    intenta matchear contra los renglones con factura='NA' (o vacía) por
  //    PROVEEDOR + MONTO + AÑO-MES. Cubre el caso donde el Saldos no tiene
  //    el folio de la factura pero sí registra el movimiento del banco.
  const stillSinMatch = []
  for (const g of sinMatchGastos) {
    const gProv = normProv(g.proveedor)
    const gTotal = totalNativo(g)
    const alias = ALIAS_POR_RFC[String(g.rfc || '').toUpperCase()] || []
    if (gProv.length < 8 && !(gTotal > 0 && alias.length)) { stillSinMatch.push(g); continue }
    const gYM = String(g.fechaFac || '').slice(0, 7)

    let foundIdx = -1
    for (let i = 0; i < filteredSaldos.length; i++) {
      if (renglonUsado.has(i)) continue
      const row = filteredSaldos[i]
      if (facturasPorRenglon[i].length) continue  // solo filas sin factura clara
      const rowYM = (saldosFechaToIso(row.fecha) || '').slice(0, 7)
      if (rowYM && gYM && rowYM !== gYM) continue
      const concNorm = normProv(row.concepto)
      const montoIgual = gTotal > 0 && Math.abs((Number(row.egreso) || 0) - gTotal) <= 0.01
      // Proveedor: la razón social completa en el concepto o, sólo si el monto
      // es exacto, un alias conocido ("SAT ISR Sep"). No vale la primera
      // palabra: en nóminas es el nombre de pila y hay varios "David".
      const provOk = (gProv.length >= 8 && concNorm.includes(gProv)) || (montoIgual &&
        alias.some(a => new RegExp(`\\b${a}\\b`, 'i').test(row.concepto || '')))
      if (!provOk) continue
      // Tolera diferencia de ±$0.01 para evitar problemas de redondeo.
      if (gTotal > 0 && !montoIgual) continue
      foundIdx = i
      break
    }

    if (foundIdx === -1) { stillSinMatch.push(g); continue }
    renglonUsado.add(foundIdx)
    const sRow = filteredSaldos[foundIdx]
    const tipoDiffers = !!sRow.tipo && !!g.tipo &&
      sRow.tipo.toLowerCase() !== g.tipo.toLowerCase()
    matched.push({ gastoId: g.id, gasto: g, saldosRow: sRow, tipoDiffers, multiFactura: false, exhibiciones: 1 })
  }
  sinMatchGastos.length = 0
  sinMatchGastos.push(...stillSinMatch)

  const sinMatchSaldos = filteredSaldos.filter((_, idx) => !renglonUsado.has(idx))
  return { matched, sinMatchGastos, sinMatchSaldos }
}

// Campos de monto que se prorratean cuando el banco pagó sólo una parte.
const CAMPOS_MXN = ['importe', 'iva', 'isrTrasladado', 'ishIeps', 'retencionISR',
  'retencionIVA', 'retenciones', 'totalCFDI', 'subtotal']
const CAMPOS_EXT = ['importeUSD', 'ivaUSD', 'retencionesUSD', 'montoExtranjero', 'montoUSD']
const r2 = (n) => Math.round((Number(n) || 0) * 100) / 100

/* Ajusta los montos de un gasto al pago que registró el Saldos.

   Si el renglón paga SÓLO esta factura y el egreso no coincide con su total,
   `cobroSaldos` guarda cobrado vs facturado para las columnas COBRADO /
   FACTURADO / DIFERENCIA del Excel. Si además la factura se pagó en varias
   exhibiciones (2+ pagos en el Saldos: anticipo 30% + liquidación, 12
   domiciliaciones de un seguro), el reporte lleva lo que se pagó en ESTE
   pago, así que los montos se prorratean al egreso. Con un solo pago (Total
   Play cobró 1,298 de 1,378.65) no se prorratea: sólo se marca la diferencia.

   Los montos originales se guardan en `montosOriginales`, así que volver a
   cotejar parte siempre de la factura completa (no prorratea dos veces). */
export function aplicarCobroSaldos(g, m) {
  const base = g.montosOriginales || null
  const original = { ...g, ...(base || {}) }
  const limpio = { ...original, montosOriginales: undefined, cobroSaldos: null }
  const egreso = r2(m?.saldosRow?.egreso)
  const facturado = r2(totalNativo(original))
  if (!m || m.multiFactura || egreso <= 0 || facturado <= 0 || Math.abs(egreso - facturado) <= 0.01) {
    return limpio
  }
  const ext = !!original.esMonedaExtranjera
  const moneda = (original.monedaCodigo || original.moneda || 'MXN').toString().toUpperCase()
  const exhibiciones = m.exhibiciones ?? 1
  const parcial = egreso < facturado && exhibiciones >= 2
  const cobroSaldos = { cobrado: egreso, facturado, moneda, parcial }
  if (!parcial) return { ...limpio, cobroSaldos }

  const ratio = egreso / facturado
  const campos = ext ? CAMPOS_EXT : CAMPOS_MXN
  const snapshot = {}
  const out = { ...limpio, cobroSaldos }
  for (const c of [...CAMPOS_MXN, ...CAMPOS_EXT]) {
    if (original[c] !== undefined) snapshot[c] = original[c]
  }
  for (const c of campos) {
    if (typeof original[c] === 'number') out[c] = r2(original[c] * ratio)
  }
  // El residuo del redondeo va al importe para que el TOTAL del Excel
  // (importe + IVA + IEPS − retenciones) sea exactamente el egreso.
  if (ext) {
    out.montoExtranjero = egreso
    if (original.montoUSD) out.montoUSD = egreso
    out.importeUSD = r2(egreso - (out.ivaUSD || 0) + (out.retencionesUSD || 0))
  } else {
    out.totalCFDI = egreso
    out.importe = r2(egreso - (out.iva || 0) - (out.ishIeps || 0) + (out.retenciones || 0))
  }
  out.montosOriginales = snapshot
  return out
}
