// src/presupuestoDinamico.js
// Genera el "SMTo - Presupuesto Dinámico" del mes (perfil de Alejandro Olivar).
//
// Fuentes, las mismas con que se llenaba a mano:
//   1) Hoja de trabajo del mes (la que exporta esta app, con su hoja de
//      egresos —"Reporte SMTO" o "Egresos"— y su hoja "Ingresos"):
//      llena Ingresos (filas 15-44, por F. PAGO) y Egresos (55-107, por F. COBRO).
//   2) Archivo de Saldos (precargado): llena el Consolidado bancario (121-229).
//   3) T/C de la Hoja de trabajo: va a la fila 203 del mes.
// Se escribe directo en el XML del Presupuesto base (conserva logo, estilos,
// paneles e impresión), se recalculan todas sus fórmulas aquí mismo (sólo usa
// SUM y NOW) para que el archivo abra con los totales correctos en cualquier
// visor, y se coteja contra los saldos del archivo de Saldos.
//
// Las reglas se validaron contra junio y septiembre 2026 (capturados a mano) y
// reproducen octubre 2026 sin diferencias. Ver "Como se llena el Presupuesto
// Dinamico.md" en ~/Documents/SMTO/Presupuesto Dinamico/.
import JSZip from 'jszip'
import * as XLSX from 'xlsx'
import { parseSaldosXLSX, leerSaldosWorkbook, saldosFechaToIso } from './saldosCotejo'

const rango = (a, b) => Array.from({ length: b - a + 1 }, (_, i) => a + i)
export const colANum = (c) => c.split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0)
export const numACol = (n) => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26) } return s }

// ── Estructura del Presupuesto (hoja "Real") ──
const MESES_COLS = { 1: ['N', 'R'], 2: ['S', 'V'], 3: ['W', 'Z'], 4: ['AA', 'AD'], 5: ['AE', 'AI'], 6: ['AJ', 'AM'],
  7: ['AN', 'AR'], 8: ['AS', 'AV'], 9: ['AW', 'AZ'], 10: ['BA', 'BE'], 11: ['BF', 'BI'], 12: ['BJ', 'BM'] }
export const MESES_CORTOS = ['', 'Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic']
export const MESES_NOMBRE = ['', 'enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre']
const FILA_FECHAS = 12
const FILA_TC = 203
const FILAS_INGRESOS = rango(15, 44)
const FILAS_EGRESOS = [...rango(55, 66), ...rango(71, 90), 95, 96, 97, ...rango(102, 107)]
// Consolidado bancario: pestaña del Saldos -> primera fila del bloque
// (Ingresos, Traspasos+, Egresos, Traspasos-; la quinta es el saldo).
export const BLOQUES_BANCO = {
  'BBVA MXN Cheques': 121, 'BBVA MXN Concent': 127, 'Monex MXN Cheques': 133,
  'Clara MXN Credito': 139, 'Kapital MXN Cheques': 145, 'Kapital MXN Factoraje': 151,
  'Kapital MXN Flex': 157, 'Xepelin MXN Crédito': 163, 'Monex MXN Credito': 169,
  'Konfio Crédito MXN': 175, 'BBVA MXN Crédito': 181, 'Konfio Tarjeta Crédito MXN': 187,
  'Inbursa MXN Cheques': 193,
  'BBVA USD Cheques': 205, 'Monex USD Cheques': 212, 'Monex USA USD SMTo INC': 219, 'Texas Bank USD': 226,
}
export const CUENTAS_USD = new Set(['BBVA USD Cheques', 'Monex USD Cheques', 'Monex USA USD SMTo INC', 'Texas Bank USD'])
const ING = 0, TMAS = 1, EGR = 2, TMENOS = 3

export const norm = (s) => String(s ?? '').normalize('NFD').replace(/[̀-ͯ]/g, '')
  .toLowerCase().replace(/\s+/g, ' ').trim()
const columnasMes = (mes) => { const [a, b] = MESES_COLS[mes]; return rango(colANum(a), colANum(b)).map(numACol) }

// ── Reglas de clasificación ──
const EGRESOS_TIPO = {
  // Gastos estratégicos
  'nomina soc': 55, 'nomina ser': 56, 'nomina ven': 57, 'nomina adm': 58,
  'fondo de ahorro': 59, 'regalias': 60, 'seguros': 61, 'celular': 62, 'renta oficina': 63,
  'imss': 64, 'infonavit': 64, 'isr': 65, 'iva': 65, 'isn': 66, '3% isn': 66,
  // Gastos tácticos
  'gasolina': 71, 'gasolina servicio': 72, 'consumo': 73,
  // Viáticos Servicio = todo lo de viaje SIN "Ventas"
  'consumo viaticos': 74, 'hotel': 74, 'avion': 74, 'taxi': 74, 'estacionamiento': 74,
  'casetas': 74, 'caseta': 74, 'autobus': 74, 'gasolina viaticos': 74,
  'gasolina ventas': 75, 'gastos rep': 76,
  // Viáticos Ventas = todo lo de viaje CON "Ventas" + Gastos Rep Viáticos
  'avion ventas': 77, 'hotel ventas': 77, 'casetas ventas': 77, 'estacionamiento ventas': 77,
  'taxi ventas': 77, 'autobus ventas': 77, 'gasolina ventas viaticos': 77, 'gastos rep viaticos': 77,
  'consumo ventas': 77, 'consumo ventas viaticos': 77,
  'marketing': 78, 'automovil': 79, 'renta auto': 79, 'curso': 80, 'cursos': 80,
  'pasaporte o visa': 81, 'pasaporte y visa': 81, 'permiso': 82, 'permisos': 82, 'pc': 83,
  'contador': 84, 'comercio exterior': 85, 'consultoria': 86, 'envios': 87, 'aduana': 88,
  'transporte': 89, 'transportes': 89, 'cogs': 90,
  // Cuellos de botella
  'manto auto': 95, 'it & sw': 96, 'herramientas': 97, 'herramientas ventas': 97,
  // Generales
  'papeleria': 102, 'uniformes': 103, 'comisiones banco': 104, 'prestamo nomina': 105,
  'pago credito': 106, 'no comprobado': 107, 'no coprobado': 107,
}
const EGRESOS_IGNORAR = new Set(['rechazada', 'traspaso'])
// Principales: [palabra clave en el cliente, fila Comisiones, fila Consultoría]
const PRINCIPALES = [
  ['VITROX', 15, 17],
  ['VIE TECHNOLOGIES', 15, null],          // decisión 10-nov-2026: va con ViTrox Ventas
  ['UNIVERSAL INSTRUMENTS', 18, 19],
  ['FANCORT', 20, 21],
  ['ANA TRADING', 22, 23],
  ['LPKF', 24, 25],
  ['DIGITALTEST', 26, 27], ['DIGITAL TEST', 26, 27],
  ['ESSEMTEC', 28, 29],
  ['SCIENSCOPE', 30, 31],
  ['ADMAT', 32, null],
  ['ICT HOLDING', null, 33],
  ['SMT THERMAL', 34, null],
  ['SYSTRONIC', 35, null],
]
// Clientes cuya "Consultoría" se registraba en Distribución (marcas que también distribuimos)
const CONSULTORIA_A_DISTRIBUCION = ['TEKTRONIX', 'PREEN POWER']
const TIPOS_DISTRIBUCION = new Set(['partes', 'equipo', 'soldadura', 'sodadura', 'refacciones', 'nota de credito', 'material', 'etiquetas'])
const TIPOS_CONSULTORIA = new Set(['consultoria', 'consulting', 'servicio', 'soporte', 'renta'])
const esPersonaFisica = (rfc) => String(rfc || '').toUpperCase().replace(/[^A-Z0-9&Ñ]/g, '').length === 13

export function clasificaIngreso(tipo, cliente, rfc) {
  const t = norm(tipo)
  const c = String(cliente || '').toUpperCase()
  const princ = PRINCIPALES.find(p => c.includes(p[0]))
  if (t.startsWith('comision')) {
    if (princ && princ[1]) return { fila: princ[1], regla: `Comisiones ${princ[0]}` }
    return { fila: null, regla: `Comisión de un cliente que no tiene renglón en el Presupuesto: ${cliente}` }
  }
  if (TIPOS_CONSULTORIA.has(t)) {
    if (princ && princ[2]) return { fila: princ[2], regla: `Consultoría ${princ[0]}` }
    if (CONSULTORIA_A_DISTRIBUCION.some(k => c.includes(k))) return { fila: 36, regla: 'Consultoría de marca distribuida → Distribución' }
    return { fila: 37, regla: 'Consultoría de cliente → Clientes Diversos Soporte' }
  }
  if (['equipo', 'pc', 'laptop', 'venta activos', 'venta activo', 'activo'].includes(t) && esPersonaFisica(rfc)) {
    return { fila: 38, regla: 'Equipo vendido a persona física → Venta Activos' }
  }
  if (['venta activos', 'venta activo', 'activo'].includes(t)) return { fila: 38, regla: 'Venta Activos' }
  if (TIPOS_DISTRIBUCION.has(t)) return { fila: 36, regla: `${tipo} → Clientes Diversos Distribución` }
  if (t.startsWith('devoluci')) return { fila: 40, regla: 'Devoluciones' }
  if (t.startsWith('interes')) return { fila: 41, regla: 'Intereses' }
  if (t === 'fondo de ahorro') return { fila: 39, regla: 'Fondo de Ahorro' }
  if (t.startsWith('pago prestamo')) return { fila: 44, regla: 'Pago Préstamos' }
  if (t.startsWith('capital')) return { fila: 43, regla: 'Capital Variable' }
  if (t.startsWith('prestamo') || t.startsWith('credito') || t.includes('factoraje')) {
    return { fila: 42, regla: 'Crédito / Préstamo / Factoraje' }
  }
  return { fila: null, regla: `Tipo de ingreso sin regla: "${tipo || '(vacío)'}"` }
}

// ── Evaluador de fórmulas (sólo aritmética, SUM, ROUND e IFERROR) ──
const REF = "(?:(?:'([^']+)'|([A-Za-z_][\\w.]*))!)?(\\$?[A-Z]{1,3}\\$?\\d+)"
const RE_RANGO = new RegExp(`${REF}:(\\$?[A-Z]{1,3}\\$?\\d+)`, 'g')
const RE_REF = new RegExp(REF, 'g')
const FUNCIONES = { SUM: '__s', ROUND: '__r', IFERROR: '__i' }
const __s = (...a) => a.reduce((x, y) => x + y, 0)
const __r = (x, n = 0) => { const p = 10 ** n; return Math.sign(x) * Math.round(Math.abs(x) * p + 1e-9) / p }
const __i = (a, b) => (Number.isFinite(a) ? a : b)

// valorDe(hoja|null, 'A1') -> número | texto | undefined
export function evaluarFormula(formula, valorDe) {
  let f = String(formula || '').replace(/^=/, '').trim()
  if (!f || /#[A-Z/0!?]+/.test(f) || f.includes('"')) return undefined
  const fns = f.match(/[A-Z][A-Z0-9.]*(?=\()/g) || []
  if (fns.some(x => !(x in FUNCIONES))) return undefined
  let error = false
  const num = (hoja, ref) => {
    const v = valorDe(hoja || null, ref.replace(/\$/g, ''))
    if (typeof v === 'number') return Number.isFinite(v) ? v : (error = true, 0)
    if (v === undefined || v === null || v === '') return 0
    if (typeof v === 'boolean') return v ? 1 : 0
    if (typeof v === 'string' && v.startsWith('#')) error = true
    return 0
  }
  const lit = (n) => `(${String(n)})`
  f = f.replace(RE_RANGO, (_, q, h, a, b) => {
    const [ca, ra] = a.replace(/\$/g, '').match(/([A-Z]+)(\d+)/).slice(1)
    const [cb, rb] = b.replace(/\$/g, '').match(/([A-Z]+)(\d+)/).slice(1)
    let s = 0
    for (let c = Math.min(colANum(ca), colANum(cb)); c <= Math.max(colANum(ca), colANum(cb)); c++) {
      for (let r = Math.min(+ra, +rb); r <= Math.max(+ra, +rb); r++) s += num(q || h, `${numACol(c)}${r}`)
    }
    return lit(s)
  })
  f = f.replace(RE_REF, (_, q, h, a) => lit(num(q || h, a)))
  f = f.replace(/[A-Z][A-Z0-9.]*(?=\()/g, (fn) => FUNCIONES[fn])
  if (error || !/^[0-9.eE+\-*/(),\s_sri]*$/.test(f)) return undefined
  try {
    const r = Function('__s', '__r', '__i', `"use strict"; return (${f});`)(__s, __r, __i)
    return Number.isFinite(r) ? r : undefined
  } catch { return undefined }
}

// Lector con memoria: valor en caché o, si la celda no lo trae (archivos
// generados por código), el resultado de su fórmula.
function crearLector(wb) {
  const memo = new Map()
  const valor = (hoja, ref, prof = 0) => {
    const k = `${hoja}!${ref}`
    if (memo.has(k)) return memo.get(k)
    const c = wb.Sheets[hoja]?.[ref]
    let v
    if (!c) v = undefined
    else if (c.t === 'e') v = c.w || '#ERROR'
    else if (c.v !== undefined && c.t !== 'z') v = c.v
    else if (c.f && prof < 40) v = evaluarFormula(c.f, (h, r) => valor(h || hoja, r, prof + 1))
    memo.set(k, v)
    return v
  }
  return valor
}

const SERIAL_0 = Date.UTC(1899, 11, 30)
const fechaIso = (v) => {
  if (typeof v === 'number' && v > 20000 && v < 80000) {
    const d = new Date(SERIAL_0 + Math.round(v) * 86400000)
    return d.toISOString().slice(0, 10)
  }
  return saldosFechaToIso(v)
}

function buscaEncabezados(ws, claves) {
  if (!ws?.['!ref']) return null
  const rg = XLSX.utils.decode_range(ws['!ref'])
  for (let r = rg.s.r; r <= Math.min(rg.e.r, 25); r++) {
    const cols = {}
    for (let c = rg.s.c; c <= Math.min(rg.e.c, 40); c++) {
      const t = norm(ws[XLSX.utils.encode_cell({ r, c })]?.v)
      if (t && !(t in cols)) cols[t] = c
    }
    if (claves.every(k => k in cols)) return { fila: r, cols }
  }
  return null
}

// Celda del T/C: la primera numérica a la derecha de "TIPO DE CAMBIO …"
// (prefiere la de USD si hay varias divisas).
function buscaTC(ws, hoja, valor, maxFila) {
  if (!ws?.['!ref']) return null
  const rg = XLSX.utils.decode_range(ws['!ref'])
  const encontrados = []
  for (let r = rg.s.r; r <= Math.min(rg.e.r, maxFila); r++) {
    for (let c = rg.s.c; c <= Math.min(rg.e.c, 40); c++) {
      const t = norm(ws[XLSX.utils.encode_cell({ r, c })]?.v)
      if (!t.includes('tipo de cambio')) continue
      for (let cc = c + 1; cc <= c + 6; cc++) {
        const ref = XLSX.utils.encode_cell({ r, c: cc })
        const v = valor(hoja, ref)
        if (typeof v === 'number' && v > 0) { encontrados.push({ tc: v, ref, usd: t.includes('usd') }); break }
      }
    }
  }
  return encontrados.find(x => x.usd) || encontrados[0] || null
}

const usaTC = (formula, ref) => {
  if (!formula || !ref) return false
  const [, col, fila] = ref.match(/([A-Z]+)(\d+)/)
  return new RegExp(`(^|[^A-Z$])\\$?${col}\\$?${fila}(?!\\d)`).test(String(formula))
}

// ── 1) Hoja de trabajo ──
export function leerHojaTrabajo(buffer) {
  const wb = XLSX.read(buffer, { type: 'array', cellFormula: true, sheetStubs: true })
  const valor = crearLector(wb)
  let egr = null, ing = null
  for (const nombre of wb.SheetNames) {
    const ws = wb.Sheets[nombre]
    if (!egr) {
      const h = buscaEncabezados(ws, ['tipo', 'total', 'f. cobro', 'banco', 'importe'])
      if (h) { egr = { nombre, ws, ...h }; continue }
    }
    if (!ing) {
      const h = buscaEncabezados(ws, ['tipo', 'total', 'f. pago', 'banco', 'importe'])
      if (h) ing = { nombre, ws, ...h }
    }
  }
  if (!egr) throw new Error('No encontré la hoja de egresos (encabezados TIPO, IMPORTE, TOTAL, F. COBRO y BANCO). ¿Es la Hoja de trabajo que exporta la app?')
  const tcEg = buscaTC(egr.ws, egr.nombre, valor, egr.fila)
  if (!tcEg) throw new Error('La Hoja de trabajo no trae el "TIPO DE CAMBIO USD".')
  const celda = (h, r, col) => h.ws[XLSX.utils.encode_cell({ r, c: col })]
  const val = (h, r, col) => (col === undefined ? undefined : valor(h.nombre, XLSX.utils.encode_cell({ r, c: col })))
  const texto = (h, r, col) => { const v = val(h, r, col); return v === undefined || v === null ? '' : String(v).trim() }

  const egresos = []
  const ce = egr.cols
  const cMe = ce['monto m.e.'] ?? ce['monto usd']
  const finE = XLSX.utils.decode_range(egr.ws['!ref']).e.r
  for (let r = egr.fila + 1; r <= finE; r++) {
    const fecha = fechaIso(val(egr, r, ce['f. cobro']))
    if (!/^\d{4}-\d{2}-\d{2}$/.test(fecha)) continue
    const total = Number(val(egr, r, ce.total)) || 0
    egresos.push({
      fila: r + 1, fecha, tipo: texto(egr, r, ce.tipo), proveedor: texto(egr, r, ce.proveedor),
      concepto: texto(egr, r, ce.concepto), banco: texto(egr, r, ce.banco), total,
      usd: usaTC(celda(egr, r, ce.importe)?.f, tcEg.ref),
      me: cMe !== undefined ? Number(val(egr, r, cMe)) || 0 : null,
    })
  }
  const ingresos = []
  if (ing) {
    const ci = ing.cols
    const tcIn = buscaTC(ing.ws, ing.nombre, valor, ing.fila) || tcEg
    const cCli = ci['nombre del cliente'] ?? ci.cliente
    const cRfc = ci['rfc cliente'] ?? ci.rfc
    const finI = XLSX.utils.decode_range(ing.ws['!ref']).e.r
    for (let r = ing.fila + 1; r <= finI; r++) {
      const fecha = fechaIso(val(ing, r, ci['f. pago']))
      if (!/^\d{4}-\d{2}-\d{2}$/.test(fecha)) continue
      ingresos.push({
        fila: r + 1, fecha, tipo: texto(ing, r, ci.tipo), cliente: texto(ing, r, cCli), rfc: texto(ing, r, cRfc),
        factura: texto(ing, r, ci.factura), concepto: texto(ing, r, ci.concepto), banco: texto(ing, r, ci.banco),
        total: Number(val(ing, r, ci.total)) || 0, usd: usaTC(celda(ing, r, ci.importe)?.f, tcIn.ref),
      })
    }
  }
  // Mes de la hoja: el que más renglones tiene.
  const cuenta = {}
  for (const x of [...egresos, ...ingresos]) cuenta[x.fecha.slice(0, 7)] = (cuenta[x.fecha.slice(0, 7)] || 0) + 1
  const ym = Object.entries(cuenta).sort((a, b) => b[1] - a[1])[0]?.[0]
  if (!ym) throw new Error('La Hoja de trabajo no tiene renglones con fecha.')
  return {
    tc: tcEg.tc, anio: +ym.slice(0, 4), mes: +ym.slice(5, 7), egresos, ingresos,
    hojaEgresos: egr.nombre, hojaIngresos: ing?.nombre || null,
  }
}

// ── 2) Semanas del mes según la fila 12 del Presupuesto ──
function semanas(ws, valorReal, anio, mes) {
  const cols = columnasMes(mes)
  const finMes = new Date(Date.UTC(anio, mes, 0)).getUTCDate()
  let prev = null
  return cols.map((c, i) => {
    let dia
    if (i === cols.length - 1) dia = finMes
    else {
      const v = valorReal(`${c}${FILA_FECHAS}`)
      if (typeof v === 'number' && v > 20000) dia = new Date(SERIAL_0 + Math.round(v) * 86400000).getUTCDate()
      else if (prev !== null) dia = prev + 7
      else throw new Error(`No entiendo la fecha de la semana en ${c}${FILA_FECHAS} del Presupuesto.`)
    }
    prev = dia
    return [c, dia]
  })
}
const columnaDe = (cortes, iso) => { const d = +iso.slice(8, 10); return (cortes.find(([, dia]) => d <= dia) || cortes[cortes.length - 1])[0] }

// ── 3) Cálculo del mes ──
export function calcularMes({ hoja, saldosRows, cortes, anio, mes }) {
  const ym = `${anio}-${String(mes).padStart(2, '0')}`
  const cols = cortes.map(([c]) => c)
  const celdas = new Map()       // `${fila}|${col}` -> { mxn, usd }
  const bancos = new Map()       // `${fila}|${col}` -> monto
  const avisos = [], errores = [], detalle = []
  const celda = (fila, col) => {
    const k = `${fila}|${col}`
    if (!celdas.has(k)) celdas.set(k, { mxn: 0, usd: 0 })
    return celdas.get(k)
  }
  const tc = hoja.tc
  const usdDe = (total) => Math.round(total / tc * 100) / 100

  // Egresos (Hoja de trabajo, por F. COBRO)
  let ultimo = null
  const meDif = new Map()
  for (const e of hoja.egresos) {
    const t = norm(e.tipo)
    if (e.fecha.slice(0, 7) !== ym) {
      avisos.push(`Egresos renglón ${e.fila}: F. COBRO ${e.fecha} es de otro mes; no se incluyó (${fmt(e.total)}).`)
      continue
    }
    if (EGRESOS_IGNORAR.has(t)) {
      avisos.push(`Egresos renglón ${e.fila}: tipo "${e.tipo}" no es gasto; no se incluyó (${fmt(e.total)}).`)
      continue
    }
    let fila, regla
    if (t.startsWith('propina')) { fila = ultimo; regla = 'Propina → mismo renglón que su gasto' }
    else { fila = EGRESOS_TIPO[t] ?? null; regla = `tipo "${e.tipo}"`; ultimo = fila }
    if (!fila) {
      errores.push(`Egresos renglón ${e.fila}: el tipo "${e.tipo || '(vacío)'}" no tiene renglón en el Presupuesto (${e.proveedor}, ${fmt(e.total)}).`)
      continue
    }
    const col = columnaDe(cortes, e.fecha)
    const c = celda(fila, col)
    if (e.usd) {
      const usd = usdDe(e.total)
      c.usd += usd
      if (e.me !== null && Math.abs(Math.abs(e.me) - Math.abs(usd)) > 0.5) {
        const k = `${e.proveedor}|${usd}|${e.me}`
        if (!meDif.has(k)) meDif.set(k, { proveedor: e.proveedor, usd, me: e.me, filas: [] })
        meDif.get(k).filas.push(e.fila)
      }
    } else c.mxn += e.total
    detalle.push({ hoja: 'Egresos', renglon: e.fila, fecha: e.fecha, tipo: e.tipo, nombre: e.proveedor, banco: e.banco,
      total: e.total, moneda: e.usd ? 'USD' : 'MXN', fila, col, regla })
  }
  for (const d of meDif.values()) {
    avisos.push(`Egresos ${d.filas.length > 1 ? 'renglones' : 'renglón'} ${d.filas.join(', ')}: ${d.proveedor} — el TOTAL equivale a ${fmt(d.usd)} USD pero MONTO M.E. dice ${fmt(d.me)} USD. Se usó el TOTAL; revisa la captura.`)
  }

  // Ingresos (Hoja de trabajo, por F. PAGO)
  for (const i of hoja.ingresos) {
    if (i.fecha.slice(0, 7) !== ym) {
      avisos.push(`Ingresos renglón ${i.fila}: F. PAGO ${i.fecha} es de otro mes; no se incluyó (${fmt(i.total)}).`)
      continue
    }
    const { fila, regla } = clasificaIngreso(i.tipo, i.cliente, i.rfc)
    if (!fila) { errores.push(`Ingresos renglón ${i.fila}: ${regla} (${fmt(i.total)}).`); continue }
    const col = columnaDe(cortes, i.fecha)
    const c = celda(fila, col)
    if (i.usd) c.usd += usdDe(i.total)
    else {
      c.mxn += i.total
      if (/USD/i.test(i.banco) && fila !== 40 && fila !== 41) {
        avisos.push(`Ingresos renglón ${i.fila}: ${i.cliente} ${fmt(i.total)} está en pesos pero el banco dice "${i.banco}". Se tomó como pesos; confirma la moneda.`)
      }
    }
    detalle.push({ hoja: 'Ingresos', renglon: i.fila, fecha: i.fecha, tipo: i.tipo, nombre: i.cliente, banco: i.banco,
      total: i.total, moneda: i.usd ? 'USD' : 'MXN', fila, col, regla })
  }

  // Consolidado bancario (archivo de Saldos)
  for (const base of Object.values(BLOQUES_BANCO)) for (let o = 0; o < 4; o++) for (const col of cols) bancos.set(`${base + o}|${col}`, 0)
  const suma = (fila, col, monto) => { if (monto) bancos.set(`${fila}|${col}`, bancos.get(`${fila}|${col}`) + monto) }
  const sinBloque = new Set()
  for (const m of saldosRows || []) {
    const iso = saldosFechaToIso(m.fecha)
    if (iso.slice(0, 7) !== ym) continue
    const ing = Number(m.ingreso) || 0, egr = Number(m.egreso) || 0
    if (!ing && !egr) continue
    const cuenta = String(m.sheet || '').trim().replace(/\s+/g, ' ')
    const base = BLOQUES_BANCO[cuenta]
    if (!base) { sinBloque.add(cuenta); continue }
    const col = columnaDe(cortes, iso)
    const cat = norm(m.tipo)
    if (cuenta === 'Monex MXN Cheques' && cat === 'pago credito') {
      // Convención de septiembre: sale como Traspaso- y entra al bloque Monex MXN Crédito.
      const cred = BLOQUES_BANCO['Monex MXN Credito']
      suma(base + TMENOS, col, egr); suma(cred + TMAS, col, egr); suma(cred + EGR, col, egr)
    } else if (cuenta === 'Konfio Crédito MXN') {
      if (cat === 'pago') suma(base + ING, col, ing || egr)
      else if (cat.startsWith('interes')) suma(base + TMENOS, col, egr || ing)
      else { suma(base + ING, col, ing); suma(base + EGR, col, egr) }
    } else if (cuenta === 'BBVA MXN Crédito') {
      // El pago baja la deuda (Traspasos+); intereses, comisiones y el crédito dispuesto la suben (Egresos).
      if (cat.startsWith('pago')) suma(base + TMAS, col, ing || egr)
      else suma(base + EGR, col, egr || ing)
    } else if (cat === 'traspaso') {
      suma(base + TMAS, col, ing); suma(base + TMENOS, col, egr)
    } else {
      suma(base + ING, col, ing); suma(base + EGR, col, egr)
    }
  }
  for (const c of sinBloque) avisos.push(`Saldos: "${c}" tiene movimientos en el mes pero no tiene renglón en el Presupuesto; no se incluyó.`)
  return { celdas, bancos, avisos, errores, detalle }
}

// ── 4) Escritura directa en el XML del Presupuesto ──
const numTxt = (x) => { const s = (Math.round(x * 100) / 100).toFixed(2).replace(/\.?0+$/, ''); return s === '-0' || s === '' ? '0' : s }
// JSZip agrega entradas de carpeta al reescribir un archivo; un .xlsx no las lleva.
const SIN_CARPETAS = { createFolders: false }
const CELDA_RE = /<c r="([A-Z]+\d+)"((?:\s+[a-zA-Z:]+="[^"]*")*)\s*(\/>|>([\s\S]*?)<\/c>)/g

async function rutaHoja(zip, nombre) {
  const wbx = await zip.file('xl/workbook.xml').async('string')
  const esc = nombre.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
  const rid = (wbx.match(new RegExp(`<sheet [^>]*name="${esc}"[^>]*r:id="([^"]+)"`)) || [])[1]
  if (!rid) return null
  const rels = await zip.file('xl/_rels/workbook.xml.rels').async('string')
  const rel = rels.match(new RegExp(`<Relationship [^>]*Id="${rid}"[^>]*/>`))?.[0]
  const tgt = rel?.match(/Target="([^"]+)"/)?.[1]
  if (!tgt) return null
  const t = tgt.replace(/^\//, '')
  return t.startsWith('xl/') ? t : `xl/${t}`
}

function celdaXml(ref, attrs, formula, valor) {
  const s = (attrs.match(/\ss="\d+"/) || [''])[0]
  return `<c r="${ref}"${s}>${formula ? `<f>${formula}</f>` : ''}<v>${String(valor)}</v></c>`
}

function insertaCelda(xml, ref, contenido) {
  const fila = ref.match(/\d+/)[0]
  const col = colANum(ref.match(/[A-Z]+/)[0])
  const m = xml.match(new RegExp(`<row r="${fila}"[^>]*?(?:/>|>([\\s\\S]*?)</row>)`))
  if (!m) throw new Error(`El Presupuesto no tiene la fila ${fila}.`)
  if (m[0].endsWith('/>')) return xml.replace(m[0], m[0].slice(0, -2) + '>' + contenido + '</row>')
  const cuerpo = m[1]
  let pos = cuerpo.length
  for (const mc of cuerpo.matchAll(/<c r="([A-Z]+)\d+"/g)) { if (colANum(mc[1]) > col) { pos = mc.index; break } }
  const ini = m.index + m[0].indexOf(cuerpo)
  return xml.slice(0, ini + pos) + contenido + xml.slice(ini + pos)
}

const fmt = (n) => Number(n || 0).toLocaleString('es-MX', { minimumFractionDigits: 2, maximumFractionDigits: 2 })

// Saldo base de cada pestaña del Saldos ("Saldo base histórico" del libro de
// Tesorería; el crudo no lo trae y arranca en 0).
export function basesDeSaldos(wb) {
  const bases = {}
  for (const nombre of wb.SheetNames) {
    const ws = wb.Sheets[nombre]
    if (!ws?.['!ref']) continue
    for (let r = 0; r < 10 && bases[nombre] === undefined; r++) {
      for (let c = 0; c < 15; c++) {
        if (!/saldo base hist/.test(norm(ws[XLSX.utils.encode_cell({ r, c })]?.v))) continue
        for (let cc = c + 1; cc < c + 6; cc++) {
          const v = ws[XLSX.utils.encode_cell({ r, c: cc })]?.v
          if (typeof v === 'number') { bases[nombre] = v; break }
        }
        break
      }
    }
  }
  return bases
}

// Lee el archivo de Saldos una sola vez: movimientos + saldo base por pestaña.
export function leerSaldosParaPresupuesto(buffer) {
  const wb = leerSaldosWorkbook(buffer)
  return { rows: parseSaldosXLSX(wb), bases: basesDeSaldos(wb) }
}

/* Genera el Presupuesto del mes de la Hoja de trabajo.
   baseBuffer: Presupuesto anterior (.xlsx). hojaBuffer: Hoja de trabajo.
   saldos: { rows, bases } de leerSaldosParaPresupuesto.
   Devuelve { ok, archivo (Uint8Array), nombreArchivo, anio, mes, tc, totales,
   cotejoBancos, avisos, errores, detalle }. Si hay partidas sin regla y no se
   pasa forzar, devuelve ok:false con los errores y no genera el archivo. */
export async function generarPresupuesto({ baseBuffer, hojaBuffer, saldos, forzar = false }) {
  const hoja = leerHojaTrabajo(hojaBuffer)
  const { anio, mes, tc } = hoja
  const wbBase = XLSX.read(baseBuffer, { type: 'array', cellFormula: true, sheetStubs: true })
  const ws = wbBase.Sheets.Real
  if (!ws) throw new Error('El Presupuesto base no tiene la hoja "Real".')

  // Valores actuales de la hoja Real (caché del archivo).
  const valores = new Map()
  const formulas = new Map()
  for (const ref of Object.keys(ws)) {
    if (ref[0] === '!') continue
    const c = ws[ref]
    if (c.t === 'e') valores.set(ref, c.w || '#ERROR')
    else if (c.v !== undefined && c.t !== 'z') valores.set(ref, c.v)
    if (c.f) formulas.set(ref, c.f)
  }
  const cortes = semanas(ws, (ref) => valores.get(ref), anio, mes)
  const cols = cortes.map(([c]) => c)
  const calc = calcularMes({ hoja, saldosRows: saldos?.rows || [], cortes, anio, mes })
  const base = { anio, mes, tc, hojaEgresos: hoja.hojaEgresos, hojaIngresos: hoja.hojaIngresos, avisos: calc.avisos, errores: calc.errores, detalle: calc.detalle }
  if (calc.errores.length && !forzar) return { ok: false, ...base }

  // Entradas del mes: concepto (Hoja de trabajo), bancos (Saldos) y T/C.
  const entradas = new Map()
  for (const fila of [...FILAS_INGRESOS, ...FILAS_EGRESOS]) {
    for (const col of cols) {
      const c = calc.celdas.get(`${fila}|${col}`)
      const mxn = c ? Math.round(c.mxn * 100) / 100 : 0
      const usd = c ? Math.round(c.usd * 100) / 100 : 0
      if (usd) {
        let f = `${numTxt(usd)}*${col}${FILA_TC}`
        if (mxn) f += (mxn > 0 ? '+' : '-') + numTxt(Math.abs(mxn))
        entradas.set(`${col}${fila}`, { f })
      } else entradas.set(`${col}${fila}`, { v: mxn })
    }
  }
  for (const [k, v] of calc.bancos) { const [fila, col] = k.split('|'); entradas.set(`${col}${fila}`, { v: Math.round(v * 100) / 100 }) }
  for (const col of cols) entradas.set(`${col}${FILA_TC}`, { v: tc })
  for (const [ref, e] of entradas) {
    if (e.f) formulas.set(ref, e.f); else { formulas.delete(ref); valores.set(ref, e.v) }
  }

  // Recalcula todas las fórmulas (en orden de filas; repite hasta que no cambie nada).
  const orden = [...formulas.entries()].sort((a, b) => {
    const [, ca, ra] = a[0].match(/([A-Z]+)(\d+)/), [, cb, rb] = b[0].match(/([A-Z]+)(\d+)/)
    return (+ra - +rb) || (colANum(ca) - colANum(cb))
  })
  const lee = (h, ref) => valores.get(ref)
  for (let pasada = 0, cambio = true; cambio && pasada < 200; pasada++) {
    cambio = false
    for (const [ref, f] of orden) {
      const v = evaluarFormula(f, lee)
      if (v === undefined) continue
      const prev = valores.get(ref)
      if (typeof prev !== 'number' || Math.abs(prev - v) > 1e-9 * Math.max(1, Math.abs(v))) { valores.set(ref, v); cambio = true }
    }
  }

  // XML: entradas + caché de cada fórmula.
  const zip = await JSZip.loadAsync(baseBuffer)
  const ruta = await rutaHoja(zip, 'Real')
  if (!ruta) throw new Error('No encontré la hoja "Real" dentro del Presupuesto base.')
  let xml = await zip.file(ruta).async('string')
  const escritas = new Set()
  xml = xml.replace(CELDA_RE, (todo, ref, attrs, cierre, inner) => {
    const e = entradas.get(ref)
    if (e) { escritas.add(ref); return celdaXml(ref, attrs, e.f, valores.get(ref) ?? 0) }
    if (inner === undefined || !inner.includes('<f')) return todo
    const v = valores.get(ref)
    if (typeof v !== 'number') return todo
    const fpart = inner.match(/<f[^>]*\/>|<f[^>]*>[\s\S]*?<\/f>/)[0]
    return `<c r="${ref}"${attrs.replace(/\st="[^"]*"/, '')}>${fpart}<v>${String(v)}</v></c>`
  })
  for (const [ref, e] of entradas) {
    if (!escritas.has(ref)) xml = insertaCelda(xml, ref, celdaXml(ref, '', e.f, valores.get(ref) ?? 0))
  }
  zip.file(ruta, xml, SIN_CARPETAS)
  let wbx = await zip.file('xl/workbook.xml').async('string')
  if (!/fullCalcOnLoad/.test(wbx)) wbx = wbx.replace(/<calcPr([^>]*?)\/>/, '<calcPr$1 fullCalcOnLoad="1"/>')
  zip.file('xl/workbook.xml', wbx, SIN_CARPETAS)
  if (zip.file('xl/calcChain.xml')) {
    zip.remove('xl/calcChain.xml')
    const ct = await zip.file('[Content_Types].xml').async('string')
    zip.file('[Content_Types].xml', ct.replace(/<Override[^>]*PartName="\/xl\/calcChain.xml"[^>]*\/>/, ''), SIN_CARPETAS)
    const rl = await zip.file('xl/_rels/workbook.xml.rels').async('string')
    zip.file('xl/_rels/workbook.xml.rels', rl.replace(/<Relationship[^>]*Target="calcChain.xml"[^>]*\/>/, ''), SIN_CARPETAS)
  }
  const archivo = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 6 } })

  // Cotejo: totales contra la Hoja y saldos al cierre contra el archivo de Saldos.
  const ym = `${anio}-${String(mes).padStart(2, '0')}`
  const sumaMes = (filas) => filas.reduce((s, f) => s + cols.reduce((t, c) => t + (Number(valores.get(`${c}${f}`)) || 0), 0), 0)
  const enMes = (x) => x.fecha.slice(0, 7) === ym
  const totales = {
    ingresosPresupuesto: sumaMes(FILAS_INGRESOS),
    ingresosHoja: hoja.ingresos.filter(enMes).reduce((s, i) => s + i.total, 0),
    egresosPresupuesto: sumaMes(FILAS_EGRESOS),
    egresosHoja: hoja.egresos.filter(e => enMes(e) && !EGRESOS_IGNORAR.has(norm(e.tipo))).reduce((s, e) => s + e.total, 0),
  }
  const finMes = `${ym}-${String(cortes[cortes.length - 1][1]).padStart(2, '0')}`
  const ultCol = cols[cols.length - 1]
  const cotejoBancos = []
  const porCuenta = new Map()
  for (const m of saldos?.rows || []) {
    const iso = saldosFechaToIso(m.fecha)
    if (!iso || iso > finMes) continue
    const cuenta = String(m.sheet || '').trim().replace(/\s+/g, ' ')
    porCuenta.set(cuenta, (porCuenta.get(cuenta) || 0) + (Number(m.ingreso) || 0) - (Number(m.egreso) || 0))
  }
  for (const [cuenta, fila] of Object.entries(BLOQUES_BANCO)) {
    if (!porCuenta.has(cuenta)) continue
    const presupuesto = Number(valores.get(`${ultCol}${fila + 4}`)) || 0
    const enSaldos = (saldos?.bases?.[cuenta] || 0) + porCuenta.get(cuenta)
    cotejoBancos.push({ cuenta, moneda: CUENTAS_USD.has(cuenta) ? 'USD' : 'MXN', presupuesto, saldos: enSaldos, diferencia: presupuesto - enSaldos })
  }
  return {
    ok: true, ...base, archivo, totales, cotejoBancos, celdasEscritas: entradas.size,
    nombreArchivo: `SMTo - Presupuesto Dinámico ${MESES_CORTOS[mes]} ${anio}.xlsx`,
  }
}

// Detalle de la generación en CSV (para auditar a qué renglón fue cada partida).
export function detalleCSV(detalle) {
  const esc = (v) => { const s = String(v ?? ''); return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s }
  const filas = [['Hoja', 'Renglón de la Hoja', 'Fecha', 'Tipo', 'Proveedor / Cliente', 'Banco', 'Total MXN', 'Moneda', 'Renglón del Presupuesto', 'Columna', 'Regla']]
  for (const d of detalle) filas.push([d.hoja, d.renglon, d.fecha, d.tipo, d.nombre, d.banco, d.total.toFixed(2), d.moneda, d.fila, d.col, d.regla])
  return '﻿' + filas.map(f => f.map(esc).join(',')).join('\n')
}
