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
