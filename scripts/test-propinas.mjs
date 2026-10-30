// Prueba de regresión de elegibilidad de propina en la conciliación.
// Uso: npm run test:propinas
import assert from 'node:assert/strict'
import { isEligibleForTip } from '../src/propinas.js'

const casos = [
  // [descripción, factura, fila del CSV, esperado]
  ['Restaurante con categoría Clara "Alimentos"', { tipo: 'Consumo', proveedor: 'LA MESA SA', concepto: 'Consumo' }, { categoria: 'Alimentos' }, true],
  ['Restaurante sin categoría (keyword en descripción)', { tipo: 'Consumo', proveedor: 'OPERADORA XYZ', concepto: 'Consumo' }, { descripcion: 'BUFALO WILD WINGS' }, true],
  ['"Restaurante La Carretera" tipo Consumo sigue elegible', { tipo: 'Consumo', proveedor: 'RESTAURANTE LA CARRETERA', concepto: 'Consumo' }, { categoria: 'Restaurantes' }, true],
  // Regresión de la captura: caseta y gasolina NUNCA llevan propina, aunque
  // Clara las haya categorizado con una categoría elegible genérica.
  ['Caseta con categoría Clara "Viáticos"', { tipo: 'Casetas', proveedor: 'RED DE CARRETERAS DE OCCIDENTE', concepto: 'Peaje' }, { categoria: 'Viáticos' }, false],
  ['Gasolina con categoría Clara "Consumo"', { tipo: 'Gasolina', proveedor: 'MULTISERVICIO LOS PITS', concepto: 'Gasolina Premium' }, { categoria: 'Consumo' }, false],
  ['Gasolina sin tipo, detectada por concepto', { tipo: '', proveedor: 'MULTISERVICIO LOS PITS', concepto: 'GASOLINA PREMIUM' }, { categoria: 'Consumo' }, false],
  ['Caseta sin tipo, detectada por proveedor', { tipo: '', proveedor: 'RED DE CARRETERAS DE OCCIDENTE', concepto: '' }, {}, false],
  ['Hotel Ventas', { tipo: 'Hotel Ventas', proveedor: 'FIESTA INN', concepto: 'Hospedaje' }, { categoria: 'Viáticos' }, false],
  ['Taxi / Uber', { tipo: 'Taxi', proveedor: 'UBER', concepto: 'Viaje' }, { categoria: 'Consumo' }, false],
  // Regresión de la prueba real: gasolinera con tipo FALLBACK "Consumo
  // Viáticos" (autoDetectTipo no la reconoció) y Clara "Viáticos".
  ['Gasolinera con tipo fallback y clave SAT 15101514', { tipo: 'Consumo Viáticos', proveedor: 'MULTISERVICIO LOS PITS', concepto: 'PREMIUM', claveProdServ: '15101514' }, { categoria: 'Viáticos' }, false],
  ['Gasolinera con tipo fallback, concepto "MAGNA" (sin clave SAT)', { tipo: 'Consumo Viáticos', proveedor: 'MULTISERVICIO LOS PITS', concepto: 'MAGNA' }, { categoria: 'Viáticos' }, false],
  ['Gasolinera con tipo fallback, delatada por la descripción del banco', { tipo: 'Consumo Viáticos', proveedor: 'MULTISERVICIO LOS PITS', concepto: 'VENTA' }, { categoria: 'Viáticos', descripcion: 'GASOLINERA LOS PITS' }, false],
  ['Restaurante con tipo fallback sigue elegible', { tipo: 'Consumo Viáticos', proveedor: 'MARIA DEL MAR CAYON GONZALEZ', concepto: 'CONSUMO DE ALIMENTOS', claveProdServ: '90101501' }, { categoria: 'Restaurantes' }, true],
  ['"Hamburguesa premium" no se veta', { tipo: 'Consumo', proveedor: 'BURGER HOUSE', concepto: 'Hamburguesa premium' }, { categoria: 'Alimentos' }, true],
  ['Categoría Clara no alimenticia', { tipo: 'Consumo', proveedor: 'LA MESA SA', concepto: 'Consumo' }, { categoria: 'Gasolineras' }, false],
]

let fallas = 0
for (const [desc, inv, row, esperado] of casos) {
  try {
    assert.equal(isEligibleForTip(inv, row), esperado)
    console.log(`  ✓ ${desc}`)
  } catch {
    fallas++
    console.log(`  ✗ ${desc} — esperaba ${esperado}`)
  }
}
console.log(fallas ? `\n${fallas} prueba(s) fallaron` : `\n${casos.length} pruebas OK`)
process.exit(fallas ? 1 : 0)
