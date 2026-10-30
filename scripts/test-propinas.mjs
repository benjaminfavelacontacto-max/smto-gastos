// Prueba de regresión de elegibilidad de propina en la conciliación.
// Uso: npm run test:propinas
import assert from 'node:assert/strict'
import { isEligibleForTip } from '../src/propinas.js'
import { autoDetectTipo } from '../src/tipoRules.js'

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
  // Consumo de alimentos facturado por un HOTEL: sí lleva propina.
  ['Restaurante del hotel (tipo Consumo Viáticos)', { tipo: 'Consumo Viáticos', proveedor: 'HOSPITALIDAD DE IRAPUATO', concepto: 'Restaurante' }, { categoria: 'Alimentos' }, true],
  ['Hotel con tipo viejo "Hotel" pero concepto de alimentos', { tipo: 'Hotel', proveedor: 'HOTEL FIESTA INN', concepto: 'Consumo de alimentos y bebidas' }, { categoria: 'Viáticos' }, true],
  ['Hotel tipo "Hotel", Clara dice "Alimentos"', { tipo: 'Hotel', proveedor: 'OPERADORA DE HOTELES', concepto: 'Servicio' }, { categoria: 'Alimentos' }, true],
  ['Hospedaje del hotel sigue sin propina', { tipo: 'Hotel', proveedor: 'HOSPITALIDAD DE IRAPUATO', concepto: 'Hospedaje' }, { categoria: 'Viajes' }, false],
  ['Hotel con categoría "Viajes" y concepto de habitación', { tipo: 'Consumo Viáticos', proveedor: 'HOTEL CITY EXPRESS', concepto: 'Habitación sencilla' }, { categoria: 'Viáticos' }, false],
  ['Categoría Clara no alimenticia', { tipo: 'Consumo', proveedor: 'LA MESA SA', concepto: 'Consumo' }, { categoria: 'Gasolineras' }, false],
]

// Tipo detectado: un hotel que factura comida es "Consumo Viáticos".
const casosTipo = [
  ['Hotel factura restaurante', ['HOTEL FIESTA INN', 'Consumo de alimentos y bebidas', 'Servicio'], 'Consumo Viáticos'],
  ['Hotel factura room service', ['OPERADORA DE HOTELES GRAND', 'Servicio a la habitación', 'Servicio'], 'Consumo Viáticos'],
  ['Hotel por clave SAT de restaurante', ['HOTEL MEDRANO', 'Servicio', 'Servicio', '90101501'], 'Consumo Viáticos'],
  ['Hotel factura comida (Ventas)', ['FIESTA INN', 'Alimentos', 'Ventas'], 'Gastos Rep Viáticos'],
  ['Hotel factura hospedaje', ['HOTEL FIESTA INN', 'Hospedaje', 'Servicio'], 'Hotel'],
  ['Hotel factura habitación + alimentos', ['HOTEL FIESTA INN', 'Habitación con desayuno', 'Servicio'], 'Hotel'],
  ['Hotel por nombre sin concepto de comida', ['CITY EXPRESS', 'Noche', 'Servicio'], 'Hotel'],
  ['Gasolinera desconocida por clave SAT', ['MULTISERVICIO LOS PITS', 'Premium', 'Servicio', '15101514'], 'Gasolina Viáticos'],
]

let fallas = 0
for (const [desc, args, esperado] of casosTipo) {
  const real = autoDetectTipo(...args)
  if (real === esperado) console.log(`  ✓ Tipo: ${desc}`)
  else { fallas++; console.log(`  ✗ Tipo: ${desc} — esperaba ${esperado}, salió ${real}`) }
}
for (const [desc, inv, row, esperado] of casos) {
  try {
    assert.equal(isEligibleForTip(inv, row), esperado)
    console.log(`  ✓ ${desc}`)
  } catch {
    fallas++
    console.log(`  ✗ ${desc} — esperaba ${esperado}`)
  }
}
console.log(fallas ? `\n${fallas} prueba(s) fallaron` : `\n${casos.length + casosTipo.length} pruebas OK`)
process.exit(fallas ? 1 : 0)
