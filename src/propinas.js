/* ═══════════════════════════════════════════════════
   ELEGIBILIDAD DE PROPINA (conciliación bancaria)
   Extraído de validarBanco() para poder probarlo aislado
   (scripts/test-propinas.mjs). Lo usan los Pass 1 (escalera de %) y
   Pass 2 (delta) — un par (factura, cargo) NO elegible solo puede ligarse
   por monto directo, nunca vía una propina fantasma.
═══════════════════════════════════════════════════ */

import { esConsumoAlimentos } from './tipoRules.js'

export const RESTAURANT_KEYWORDS = [
  'restaurant', 'restaurante', 'bar', 'cafe', 'café', 'cafeteria', 'cafetería',
  'comida', 'aerocomidas', 'taqueria', 'taquería', 'pizzeria', 'pizzería',
  'grill', 'steakhouse', 'bistro', 'parrilla', 'cantina', 'fonda',
  'sushi', 'cocina', 'kitchen', 'diner', 'eatery', 'jumper',
  'wings', 'mariscos', 'marisqueria', 'marisquería', 'asador', 'birria',
  'antojitos', 'burger', 'burguer', 'pub', 'brewery', 'cerveceria', 'cervecería',
]

export const TIP_ELIGIBLE_CATEGORIES = [
  'alimentos', 'bares y bebidas alcoholicas', 'bares y bebidas alcohólicas',
  'food', 'restaurants', 'restaurantes', 'comida', 'consumo', 'viaticos', 'viáticos',
]

// ── Veto por evidencia de la FACTURA o del cargo ──
// Antes la categoría de Clara mandaba "definitivamente": una caseta o
// gasolinera que Clara etiquetara como "Viáticos"/"Consumo" (ambas en la
// lista de arriba) quedaba elegible y podía ligarse con una propina inventada
// (caseta $355 + "propina" $95.80 = cargo $450.80; gasolina $1,030 +
// "propina" $220.05 = cargo $1,250.05). Con evidencia de que NO es comida,
// nunca hay propina.

// 1) Tipo del gasto. OJO: "Consumo Viáticos" / "Gastos Rep" son también el
//    FALLBACK de autoDetectTipo cuando no reconoce al proveedor, así que NO
//    prueban que sea comida — por eso no eximen de los vetos 2–4.
const TIPO_SIN_PROPINA = /gasolin|caseta|hotel|taxi|avi[oó]n|estacionamiento|renta auto|autom[oó]vil|manto auto|env[ií]os/i

// 2) Clave SAT del CFDI: 1510xxxx = combustibles (gasolina, diésel…).
const CLAVE_SAT_COMBUSTIBLE = /^1510/

// 3) Combustible: inequívoco en proveedor, concepto o descripción del banco.
//    ("premium"/"magna" solo cuando el concepto ES el producto, para no vetar
//    una "hamburguesa premium".)
const COMBUSTIBLE = ['gasolin', 'combustible', 'diesel', 'diésel', 'pemex', 'oxxo gas']
const CONCEPTO_COMBUSTIBLE = /^\s*(gasolina\s+)?(premium|magna|regular|diesel|diésel)\b/i

// 4) Otros giros sin propina, en proveedor/concepto. Pueden aparecer en el
//    nombre de un restaurante ("Restaurante La Carretera"), así que no vetan
//    si el proveedor/concepto ya dice que es comida.
const OTROS_SIN_PROPINA = ['caseta', 'peaje', 'autopista', 'carretera', 'capufe', 'hotel', 'hospedaje', 'uber', 'didi', 'taxi', 'estacionamiento']

// Categorías de Clara que SÍ son comida (no las genéricas "Viáticos"/"Consumo").
const CATEGORIA_COMIDA = /alimentos|restaurant|bares|food/i

export const facturaSinPropina = (inv, csvRow) => {
  // Combustible: veto absoluto, nada lo exime.
  if (CLAVE_SAT_COMBUSTIBLE.test(String(inv?.claveProdServ || '').trim())) return true
  const factura = ((inv?.proveedor || '') + ' ' + (inv?.concepto || '')).toLowerCase()
  const banco = (csvRow?.descripcion || '').toLowerCase()
  if (COMBUSTIBLE.some(kw => factura.includes(kw) || banco.includes(kw))) return true
  if (CONCEPTO_COMBUSTIBLE.test(inv?.concepto || '')) return true
  // Consumo de alimentos facturado por un hotel (restaurante/bar/room
  // service): el concepto, la clave SAT o la categoría "Alimentos" de Clara
  // dicen que es comida aunque el emisor (o un tipo viejo) diga Hotel.
  const esComida = esConsumoAlimentos(inv?.concepto || '', inv?.claveProdServ) ||
    CATEGORIA_COMIDA.test(csvRow?.categoria || '')
  if (TIPO_SIN_PROPINA.test(inv?.tipo || '') && !(/hotel/i.test(inv?.tipo || '') && esComida)) return true
  const pareceComida = esComida || RESTAURANT_KEYWORDS.some(kw => factura.includes(kw))
  return !pareceComida && OTROS_SIN_PROPINA.some(kw => factura.includes(kw))
}

export const isEligibleForTip = (inv, csvRow) => {
  if (facturaSinPropina(inv, csvRow)) return false
  const csvCat = (csvRow?.categoria || '').toLowerCase().trim()
  if (csvCat) {
    // CSV has a category column: a food category → eligible; any other
    // category → never tip (the vetoes above already ran).
    return TIP_ELIGIBLE_CATEGORIES.some(c => csvCat.includes(c))
  }
  // No CSV category (non-Clara source, o estado de cuenta Clara en Excel
  // que no trae la columna Categoría). Husmea palabras clave de restaurante
  // tanto en la factura (proveedor + concepto) como en la descripción del
  // propio movimiento bancario ("CASA MATRIA RESTAURANT", "REST WENDYS",
  // "BUFALO WILD WINGS"), que suele delatar el giro aunque el proveedor del
  // CFDI sea un nombre corporativo sin la palabra "restaurante".
  const haystack = (
    (inv?.proveedor || '') + ' ' + (inv?.concepto || '') + ' ' + (csvRow?.descripcion || '')
  ).toLowerCase()
  return RESTAURANT_KEYWORDS.some(kw => haystack.includes(kw))
}
