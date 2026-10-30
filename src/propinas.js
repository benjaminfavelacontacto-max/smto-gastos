/* ═══════════════════════════════════════════════════
   ELEGIBILIDAD DE PROPINA (conciliación bancaria)
   Extraído de validarBanco() para poder probarlo aislado
   (scripts/test-propinas.mjs). Lo usan los Pass 1 (escalera de %) y
   Pass 2 (delta) — un par (factura, cargo) NO elegible solo puede ligarse
   por monto directo, nunca vía una propina fantasma.
═══════════════════════════════════════════════════ */

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

// Veto por evidencia de la FACTURA. Antes la categoría de Clara mandaba
// "definitivamente": una caseta o gasolinera que Clara etiquetara como
// "Viáticos"/"Consumo" (ambas en la lista de arriba) quedaba elegible y podía
// ligarse con una propina inventada (p. ej. caseta $355 + "propina" $95.80 =
// cargo $450.80). Si el TIPO del gasto o el proveedor/concepto dicen que es
// gasolina, caseta, hotel, transporte, etc., nunca hay propina.
const TIPO_SIN_PROPINA = /gasolin|caseta|hotel|taxi|avi[oó]n|estacionamiento|renta auto|autom[oó]vil|manto auto|env[ií]os/i
const FACTURA_SIN_PROPINA = [
  'gasolina', 'gasolinera', 'combustible', 'diesel', 'diésel', 'pemex',
  'caseta', 'peaje', 'autopista', 'carretera', 'capufe',
  'hotel', 'hospedaje', 'uber', 'didi', 'taxi', 'estacionamiento',
]
// Si el tipo ya dice comida ("Consumo", "Gastos Rep"), gana el tipo: así un
// "Restaurante La Carretera" no queda vetado por la palabra "carretera".
const TIPO_COMIDA = /consumo|gastos rep|alimento/i
export const facturaSinPropina = inv => {
  const tipo = inv?.tipo || ''
  if (TIPO_SIN_PROPINA.test(tipo)) return true
  if (TIPO_COMIDA.test(tipo)) return false
  const texto = ((inv?.proveedor || '') + ' ' + (inv?.concepto || '')).toLowerCase()
  return FACTURA_SIN_PROPINA.some(kw => texto.includes(kw))
}

export const isEligibleForTip = (inv, csvRow) => {
  if (facturaSinPropina(inv)) return false
  const csvCat = (csvRow?.categoria || '').toLowerCase().trim()
  if (csvCat) {
    // CSV has a category column: a food category → eligible; any other
    // category → never tip (the invoice veto above already ran).
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
