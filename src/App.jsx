import { useState, useRef, useMemo, useEffect, useCallback, memo } from 'react'
import JSZip from 'jszip'
import * as XLSX from 'xlsx'
import { motion, AnimatePresence } from 'framer-motion'
import { CheckCircle2, X, CreditCard, Target, Sparkles, AlertTriangle, FileText, FileSpreadsheet, FileWarning, Package, Check, Plus, Link2, Search, Download, ArrowRight, ChevronDown, XCircle, AlertCircle, Info } from 'lucide-react'
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
  'Curso',
  'Envíos',
  'Estacionamiento Ventas',
  'Gasolina Ventas',
  'Gasolina Ventas Viáticos',
  'Gasolina Viáticos',
  'Gastos Rep',
  'Gastos Rep Viáticos',
  'Herramientas Ventas',
  'Hotel Ventas',
  'IT & SW',
  'Manto Auto',
  'Marketing',
  'Permiso',
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
  'Curso',
  'Envíos',
  'Estacionamiento',
  'Gasolina',
  'Gasolina Viáticos',
  'Herramientas',
  'Hotel',
  'IT & SW',
  'Manto Auto',
  'Marketing',
  'No Comprobado',
  'Papelería',
  'PC',
  'Permiso',
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
  'Gasolina Ventas Viáticos',
  'Gasolina Viáticos',
  'Herramienta',
  'Hotel',
  'IMSS',
  'ISR',
  'IT&SW',
  'IT & SW',
  'Manto Auto',
  'Marketing',
  'No Comprobado',
  'Nómina',
  'Nómina Adm',
  'Nómina Soc',
  'Nómina Ser',
  'Nómina Ven',
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

// Socios que se configuran COMO ingenieros de soporte: en toda la lógica que
// depende de la categoría (tipos, caseta, autodetección) se comportan como
// 'Servicio', aunque en la UI sigan mostrándose como Socio.
const SOCIOS_COMO_SOPORTE = ['Edie Haro', 'David Delgado', 'Isaias Valencia', 'Rosy Corral']

// Categoría efectiva para la lógica del app (no para mostrar): los socios de
// SOCIOS_COMO_SOPORTE se tratan como 'Servicio'.
const categoriaEfectiva = (colaborador) =>
  colaborador && SOCIOS_COMO_SOPORTE.includes(colaborador.nombre) ? 'Servicio' : colaborador?.categoria

// Socios que usan la lista EXTENDIDA de tipos (TIPOS_ESPECIALES) igual que
// Alejandro/Victor, PERO sin el perfil bancario especial: conservan su tarjeta
// Clara (banco default + conciliación con el CSV de Clara MXN) y NO ven la
// columna BANCO ni el cotejo con Saldos. Solo se les amplían los tipos de factura.
const SOCIOS_TIPOS_EXTENDIDOS = ['Rosy Corral', 'David Delgado']

// ¿El colaborador usa el vocabulario extendido de tipos? (lista TIPOS_ESPECIALES,
// caseta 'Caseta', autodetección sin sesgo de categoría). Aplica a los
// especiales y a los socios de SOCIOS_TIPOS_EXTENDIDOS.
const usaTiposExtendidos = (colaborador) =>
  COLABORADORES_ESPECIALES.includes(colaborador?.nombre) ||
  SOCIOS_TIPOS_EXTENDIDOS.includes(colaborador?.nombre)

// Tipo de peaje (FNI) según la lista de tipos del colaborador. Cada lista usa
// un nombre distinto para la caseta, así que devolvemos el string EXACTO que
// existe en la lista activa para que no salga "— Tipo —" en blanco:
//   Ventas/Socio → 'Casetas Ventas' (TIPOS_VENTAS)
//   Especiales   → 'Caseta'         (TIPOS_ESPECIALES)
//   Admin/Servicio → 'Casetas'      (TIPOS_NORMALES)
const tipoPeajeFNI = (colaborador) => {
  if (usaTiposExtendidos(colaborador)) return 'Caseta'
  const cat = categoriaEfectiva(colaborador)
  return (cat === 'Ventas' || cat === 'Socio') ? 'Casetas Ventas' : 'Casetas'
}

// Bancos disponibles para el dropdown de la columna BANCO (solo especiales).
// Coinciden con los nombres de las pestañas del archivo de Saldos.
const BANCOS_DISPONIBLES = [
  'BBVA MXN Cheques',
  'BBVA MXN Concent',
  'BBVA USD Cheques',
  'Clara MXN Credito',
  'Monex USD Cheques',
  'Monex MXN Cheques',
  'Monex Fondo Ahorro',
  'Kapital MXN Cheques',
  'Kapital MXN Flex',
]
// Banco por default para colaboradores NO especiales (tarjeta corporativa).
const DEFAULT_BANCO_NO_ESPECIAL = 'Clara MXN Credito'

// Helper: regresa el banco que se le debe asignar a un gasto nuevo, según
// si el colaborador es especial. Los especiales arrancan sin banco (se
// asigna después al cotejar con Saldos o manualmente). Los demás caen al
// default fijo.
const defaultBancoFor = (colaborador) =>
  COLABORADORES_ESPECIALES.includes(colaborador?.nombre) ? '' : DEFAULT_BANCO_NO_ESPECIAL

/* Folio Clara de 4 dígitos por colaborador. Aparece en la columna "PÓLIZA"
   del Excel exportado. Quien no tenga tarjeta Clara queda en 'N/A'. Los
   especiales (Olivar, Aceves, Miranda, Olivia) manejan múltiples cuentas y
   su mapeo vendrá de un Excel separado en una iteración posterior. */
const POLIZAS_CLARA = {
  'Daniel Covarrubias': '3789',
  'Edie Haro': '0610',
  'David Delgado': '0045',
  'Isaias Valencia': '1836',
  'Sigifredo Olivas': '6961',
  'Rosy Corral': '3136',
  'Heriberto Chacón': '3371',
  'Eduardo Carranco': '7507',
  'Daniel Gutierrez': '6839',
  'James Tisoto': '1295',
  'Misael Cruz': '1812',
  'Benjamin Favela': '4487',
  'Viviana Perez': '8571',
  'Juan Francisco Cuellar': '3269',
  'Juan Carlos Virgen': '4398',
  'David de Jesus Delgado': '4104',
  'Omar Monclova': '4454',
  'Antonio Uribe': '3588',
  'Natividad Garcia': '8847',
  'Raydel Baltazar': '3783',
  'Miguel Castillo': '2817',
  'David Lopez': '6985',
  'David Castillo': '8740',
  'Dario Lopez': '3891',
  'Moises Padilla': '7201',
  'Ricardo Pacheco Glez.': '1608',
  'Emmanuel Navarro': '9251',
  'Carlos Ponce': '1023',
  'Armando Torres': '0259',
  'Juan Carlos Santoyo': '9153',
  'Ricardo Pacheco': '1548',
  'Mariana Gonzalez': '1229',
  'Julio Torres': '9218',
  'Mauricio Rodriguez': '5168',
  'Hector Duarte': '5967',
  'Juan Sotomayor': '5277',
  'Marco Alvarado': '9481',
  'Marco Sanchez': '1159',
  'Juan Alfaro': '6660',
  // Soporte agregados 2026-07-23 (terminación tarjeta Clara)
  'Ernesto Rodriguez': '9164',
  'Luis Castillo': '4375',
  'Jorge Perales': '1294',
  'Marcos Ascencio': '3109',
}

// Roster fiscal de empleados — fuente: RFC Empleados.xlsx. Se usa para
// llenar RFC + nombre completo en los renglones de nómina (botón
// 'Agregar Nómina' de Alejandro Olivar). El concepto del Saldos trae solo
// el nombre corto (e.g. 'Rosalba Olivar', 'PTU Rosalba Olivar 25'); el
// matcher convierte a 'ROSALBA BEATRIZ OLIVAR CERVANTES' + RFC OICR420107MR5.
const EMPLEADOS_RFC = [
  { rfc: 'OICR420107MR5',  nombre: 'ROSALBA BEATRIZ OLIVAR CERVANTES' },
  { rfc: 'GILO690915CD5',  nombre: 'OLIVIA YOLANDA GIL LOPEZ' },
  { rfc: 'HAGE750515RA5',  nombre: 'EDIE HARO GUZMÁN' },
  { rfc: 'DECD7508312C4',  nombre: 'DAVID EDUARDO DELGADO CORONA' },
  { rfc: 'VALI8107064Z0',  nombre: 'JOSÉ ISAÍAS VALENCIA LUNA' },
  { rfc: 'OIGS750803UF5',  nombre: 'SIGIFREDO OLIVAS GONZALEZ' },
  { rfc: 'OICA721231DB6',  nombre: 'ALEJANDRO OLIVAR CERVANTES' },
  { rfc: 'COMR690829TL1',  nombre: 'ROSA MARÍA CORRAL MARTÍNEZ', aliases: ['Rosy Corral'] },
  { rfc: 'CAUH880306JH0',  nombre: 'HERIBERTO CHACÓN URTIZ' },
  { rfc: 'CASR870808CM3',  nombre: 'RAMÓN EDUARDO CARRANCO SALINAS' },
  { rfc: 'GURD880220KG8',  nombre: 'DANIEL FERNANDO GUTIERREZ RAMÍREZ' },
  { rfc: 'NARE7911283U5',  nombre: 'EMMANUEL NAVARRO RUVALCABA' },
  { rfc: 'TIRJ841005U2A',  nombre: 'JAMES TISOTO ROLDÁN' },
  { rfc: 'CUCM910507S91',  nombre: 'MISAEL CRUZ CASTAÑEDA' },
  { rfc: 'FAGA920705NUA',  nombre: 'ALEJANDRO BENJAMIN FAVELA GIL' },
  { rfc: 'PESG830329JZ2',  nombre: 'GUISELA VIVIANA PEREZ SEPULVEDA' },
  { rfc: 'PORC790703R46',  nombre: 'CARLOS ANDRES PONCE RAMIREZ' },
  { rfc: 'GOGG920529U36',  nombre: 'GEMMA GONZALEZ GOMEZ' },
  { rfc: 'CUGJ7909041M9',  nombre: 'JUAN FRANCISCO CUELLAR GAMEZ' },
  { rfc: 'VIIJ8504073E0',  nombre: 'JUAN CARLOS VIRGEN IBARRA' },
  { rfc: 'GUSP9109057E4',  nombre: 'PAOLA VIRIDIANA GUTIERREZ SANCHEZ' },
  { rfc: 'TOAA701020IP3',  nombre: 'ARMANDO TORRES ALFARO' },
  { rfc: 'SADJ8305029P2',  nombre: 'JUAN CARLOS SANTOYO DURON' },
  { rfc: 'DEMD000526AY7',  nombre: 'DAVID DE JESUS DELGADO MARQUEZ' },
  { rfc: 'MOOO880930NI9',  nombre: 'OMAR MONCLOVA ORENDAIN' },
  { rfc: 'PARR820325BP9',  nombre: 'RICARDO PACHECO ROMÁN' },
  { rfc: 'UIVA670201II0',  nombre: 'JOSÉ ANTONIO URIBE VASQUEZ' },
  { rfc: 'GAGN870908AA3',  nombre: 'NATIVIDAD GARCÍA GARCÍA' },
  { rfc: 'SOML720529STA',  nombre: 'LETICIA GUADALUPE SOLIS MENDOZA' },
  { rfc: 'BARA861202368',  nombre: 'ANGEL RAYDEL BALTAZAR RIOS' },
  { rfc: 'CARM830418FM6',  nombre: 'MIGUEL ANGEL CASTILLO RODRIGUEZ' },
  { rfc: 'LOCD9004018C9',  nombre: 'DAVID ALEJANDRO LOPEZ CAMARENA' },
  { rfc: 'CACD9203055Q6',  nombre: 'DAVID ELISEO CASTILLO CHAVEZ' },
  { rfc: 'LORD970609TR9',  nombre: 'DAVID DARIO LOPEZ ROSAS' },
  { rfc: 'GORM910221GD7',  nombre: 'MARIANA GONZALEZ RODRIGUEZ' },
  { rfc: 'AEVV840602128',  nombre: 'VICTOR MANUEL ACEVES VARGAS' },
  { rfc: 'DIVC900105330',  nombre: 'CYNTHIA YOALI DÍAZ VILLANUEVA' },
  { rfc: 'FAAL9306111TA', nombre: 'JOSE LUIS FALCON AMARO' },
  { rfc: 'TOIJ790819988',  nombre: 'JULIO CESAR TORRES IBARRA' },
  { rfc: 'ROHM9111232C7', nombre: 'MAURICIO RODRIGUEZ HERNANDEZ' },
  { rfc: 'MOBC910417CI0', nombre: 'CINDY MARITZA MONTAÑO BENITEZ' },
  { rfc: 'PAMM971124B39', nombre: 'MOISES ALEJANDRO PADILLA MORA' },
  { rfc: 'DUMH890817S35', nombre: 'HECTOR ALIM DUARTE MARTINEZ' },
  { rfc: 'SOBJ750419RW3', nombre: 'JUAN ANTONIO SOTOMAYOR BETANCOUT' },
  { rfc: 'VALM840128S14', nombre: 'MARCO ANTONIO VALENCIA LUNA' },
  { rfc: 'AASM8706127N1', nombre: 'MARCO ANTONIO ALVARADO SILVA' },
  { rfc: 'HACE030127BB1', nombre: 'EDIE EMMANUEL HARO CASTAÑEDA' },
  { rfc: 'NANM060324KG2', nombre: 'MIRANDA XIMENA NAVARRO NUÑO' },
  { rfc: 'SATM800303IN5', nombre: 'MARCO ANTONIO SANCHEZ TABAREZ' },
  { rfc: 'AANJ851122MM0', nombre: 'JUAN ALBERTO ALFARO NUÑEZ' },
  { rfc: 'PAGR051125AM7', nombre: 'RICARDO PACHECO GONZALEZ', aliases: ['Ricardo Pacheco Glez', 'Ricardo Pacheco Glez.'] },
  { rfc: 'COTL980719GD0', nombre: 'LUIS DANIEL COVARRUBIAS TORRES' },
  { rfc: 'RORJ9407238M1', nombre: 'JESUS ERNESTO RODRIGUEZ RODRIGUEZ' },
  { rfc: 'CAGL9208034V7', nombre: 'LUIS ENRIQUE CASTILLO GOMEZ' },
  { rfc: 'PEMJ99071253A', nombre: 'JORGE ANDRES PERALES MARTINEZ' },
  { rfc: 'AEPM820425FH9', nombre: 'MARCOS ASCENCIO PEREZ' },
]

// Matchea un nombre corto (extraído del concepto de Saldos, ej. "Rosalba
// Olivar") contra EMPLEADOS_RFC. Devuelve {rfc, nombre} si todas las
// palabras del nombre corto aparecen en el nombre completo del empleado
// O en alguno de sus aliases registrados. Sin acentos, mayúsculas,
// puntuación strippeada, en cualquier orden.
const matchEmpleadoByShortName = (shortName) => {
  if (!shortName) return null
  const norm = s => String(s || '')
    .toUpperCase()
    .normalize('NFD')
    .replace(/\p{Diacritic}/gu, '')
    .replace(/[.,;:'"]/g, '')
  const shortWords = norm(shortName).split(/\s+/).filter(Boolean)
  if (shortWords.length === 0) return null
  const wordsOf = (s) => new Set(norm(s).split(/\s+/).filter(Boolean))
  return EMPLEADOS_RFC.find(emp => {
    if (shortWords.every(w => wordsOf(emp.nombre).has(w))) return true
    if (emp.aliases) {
      for (const alias of emp.aliases) {
        if (shortWords.every(w => wordsOf(alias).has(w))) return true
      }
    }
    return false
  }) || null
}

const getTiposForColaborador = (colaborador) => {
  if (!colaborador) return TIPOS_NORMALES
  // Especiales (Alejandro/Victor) + socios con tipos extendidos (Rosy/David):
  // lista TIPOS_ESPECIALES (incluye 'Pasaporte o Visa').
  if (usaTiposExtendidos(colaborador)) return TIPOS_ESPECIALES
  // Edie/Isaias: lista de soporte + 'Pasaporte o Visa'.
  if (SOCIOS_COMO_SOPORTE.includes(colaborador.nombre)) return [...TIPOS_NORMALES, 'Pasaporte o Visa']
  const cat = categoriaEfectiva(colaborador)
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
  { categoria: 'Servicio', nombre: 'Luis Castillo' },
  { categoria: 'Servicio', nombre: 'Ernesto Rodriguez' },
  { categoria: 'Servicio', nombre: 'Jorge Perales' },
  { categoria: 'Servicio', nombre: 'Marcos Ascencio' },
  { categoria: 'Servicio', nombre: 'Noe Lua' },
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
  { categoria: 'Ventas', nombre: 'Juan Alfaro' },
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
  { key: 'isrTrasladado',     label: 'ISH/IEPS',     width: 120, sortable: true,  type: 'number' },
  { key: 'retencionISR',      label: 'Ret. ISR',     width: 110, sortable: true,  type: 'number' },
  { key: 'retencionIVA',      label: 'Ret. IVA',     width: 110, sortable: true,  type: 'number' },
  { key: 'retenciones',       label: 'Reten.',       width: 110, sortable: true,  type: 'number' },
  { key: 'totalCFDI',         label: 'Total Fac.',   width: 125, sortable: true,  type: 'number' },
  { key: 'formaPago',         label: 'Forma de Pago', width: 160, sortable: true,  type: 'string' },
  // BANCO sólo se muestra para colaboradores especiales (Alejandro, Victor,
  // Miranda, Olivia). El render del header y del row la salta cuando no aplica.
  { key: 'banco',             label: 'Banco',        width: 195, sortable: true,  type: 'string', specialOnly: true },
  { key: 'propinaPorcentaje', label: 'Prop. %',      width: 95,  sortable: true,  type: 'number' },
  { key: 'montoPropina',      label: 'Prop. $',      width: 105, sortable: true,  type: 'number' },
  { key: 'totalFinal',        label: 'Total Final',  width: 130, sortable: true,  type: 'number',
    // Fallback al total nativo para que los tickets extranjeros sin conciliar
    // (lado MXN en 0) no se apilen todos como "0" al ordenar.
    getValue: g => (g.totalCFDI + g.montoPropina)
      || ((Number(g.montoExtranjero) || 0) + (Number(g.propinaExtranjero) || 0)) },
  // Renombrada de "Monto USD": el monto va en la divisa original del ticket
  // (el Excel la formatea con su símbolo dentro de la columna MONTO M.E.).
  { key: 'montoUSD',          label: 'Monto M.E.',   width: 110, sortable: true,  type: 'number' },
  { key: 'tipoCambio',        label: 'T/C',          width: 80,  sortable: true,  type: 'number' },
]

/* Símbolo por divisa — espejo de CURRENCY_SYMBOLS en api/export-excel.py, para
   que la tabla muestre exactamente el mismo símbolo que el Excel. */
const MONEDA_SIMBOLOS = {
  MXN: '$',   USD: '$',   EUR: '€',   GBP: '£',   JPY: '¥',   CNY: '¥',
  CAD: 'C$',  AUD: 'A$',  CHF: 'Fr',  MYR: 'RM',  SGD: 'S$',  HKD: 'HK$',
  TWD: 'NT$', KRW: '₩',   THB: '฿',   VND: '₫',   INR: '₹',   PHP: '₱',
  IDR: 'Rp',  BRL: 'R$',
}

const simboloMoneda = (code) => MONEDA_SIMBOLOS[(code || 'MXN').toUpperCase()] || ''

/* Normaliza el código de divisa de un gasto ("RM", "yuan", "€" → ISO). Espejo
   de normaliza_moneda() en export-excel.py. */
const ALIAS_MONEDA = {
  'RM': 'MYR', 'RINGGIT': 'MYR', 'RMB': 'CNY', 'YUAN': 'CNY', 'YEN': 'JPY',
  'EURO': 'EUR', 'EUROS': 'EUR', 'US$': 'USD', 'DOLAR': 'USD', 'DOLLAR': 'USD',
  'LIBRA': 'GBP', 'POUND': 'GBP',
  '€': 'EUR', '£': 'GBP', '¥': 'CNY', '₩': 'KRW', '฿': 'THB',
  '₫': 'VND', '₹': 'INR', '₱': 'PHP',
}

const normalizaMoneda = (raw) => {
  const s = String(raw ?? '').trim().toUpperCase()
  if (!s) return 'MXN'
  if (/^[A-Z]{3}$/.test(s)) return s
  return ALIAS_MONEDA[s] || 'MXN'
}

/* Campos cuya edición recalcula el TOTAL de la fila en vivo — espejo EXACTO de
   la fórmula del Excel (=IMPORTE+IVA−ISR+ISH/IEPS−RETENCIÓN sin ISR). Sin esto,
   la tabla se quedaba con el total viejo (o $0.00 en filas manuales) y el
   número correcto solo aparecía al abrir el Excel exportado. */
const RECALC_TOTAL_FIELDS = new Set([
  'importe', 'iva', 'isrTrasladado', 'ishIeps',
  'retencionISR', 'retencionIVA', 'retenciones',
])

/* ═══════════════════════════════════════════════════
   UTILIDADES
═══════════════════════════════════════════════════ */

const genId = () => Math.random().toString(36).slice(2, 11)

/* Display dates as DD-MM-YYYY app-wide. Internal storage stays YYYY-MM-DD
   (HTML5 date input requirement). parseDateDisplay reverses for storage. */
const formatDateDisplay = (dateStr) => {
  if (!dateStr) return ''
  let d, m, y
  if (dateStr.includes('-') && dateStr.length === 10) {
    const parts = dateStr.split('-')
    if (parts[0].length === 4) {
      // YYYY-MM-DD (storage)
      [y, m, d] = parts
    } else {
      // Already DD-MM-YYYY (idempotent passthrough)
      return dateStr
    }
  } else if (dateStr.includes('/')) {
    // DD/MM/YYYY (export format or CSV)
    const parts = dateStr.split('/')
    d = parts[0]; m = parts[1]; y = parts[2] || ''
    if (y.length === 2) y = '20' + y
  } else if (dateStr.includes('-') && dateStr.length <= 8) {
    // Legacy MM-DD-YY (saved in old reports) — promote to DD-MM-YYYY
    const [mm, dd, yy] = dateStr.split('-')
    d = dd; m = mm; y = '20' + yy
  } else {
    return dateStr
  }
  return `${d.padStart(2, '0')}-${m.padStart(2, '0')}-${y}`
}

const parseDateDisplay = (s) => {
  if (!s) return ''
  const parts = s.split('-')
  if (parts.length !== 3) return s
  const [dd, mm, y] = parts
  if (y.length === 2) return `20${y}-${mm.padStart(2, '0')}-${dd.padStart(2, '0')}`
  if (y.length === 4) return `${y}-${mm.padStart(2, '0')}-${dd.padStart(2, '0')}`
  return s
}

// Re-encode an image File at most `maxWidth` px wide, JPEG `quality`,
// before shipping to OCR. Phone photos arrive at 5–10 MB which routinely
// breaks the 4.5 MB Vercel serverless body limit (HTTP 413). This brings
// them under the cap with negligible quality loss for OCR. Non-image
// files pass through untouched.
// Por encima de este tamaño, un PDF se rasteriza a JPEG en el navegador antes
// de subirlo al OCR (escaneos pesados). Debajo se manda tal cual para conservar
// la extracción de texto del server (CFDI normales, FNI ~45KB).
const LARGE_PDF_BYTES = 600_000

const sleep = (ms) => new Promise((r) => setTimeout(r, ms))

// Pausa entre facturas al procesar un LOTE con OCR. Groq (tier gratuito) limita
// a 8000 tokens/min; disparar 14 facturas instantáneas de golpe hace que casi
// todas choquen con el rate limit al mismo tiempo. Con una pausa breve entre
// cada una repartimos las peticiones y bajamos las colisiones. OJO: el mecanismo
// PRECISO sigue siendo el reintento del servidor (que respeta el Retry-After
// real de Groq); esta pausa es solo un suavizador para no hacer ráfaga. El tope
// duro es el TPM de Groq, así que un lote grande sigue tardando (ver Dev Tier).
const OCR_PACING_MS = 5000

// Costo aprox por factura con OCR de OpenAI (gpt-4o visión): prompt ~2,800 +
// imagen ~900 tokens de entrada y ~300 de salida, a $2.50/$10 por 1M in/out
// → ~$0.012 USD. ESTIMADO para mostrar ANTES de correr el lote; el server
// devuelve el costo real por llamada en `_ocrCostUsd`. Groq (fallback) = $0.
const OCR_COST_USD = 0.012
const fmtUsd = (n) => (n < 0.10 ? `$${n.toFixed(3)}` : `$${n.toFixed(2)}`)

// Clave de deduplicación al fusionar carpetas/lotes. Para CFDI, `rfc|folio` es
// único y correcto. Pero los tickets de OCR (esTicket) suelen traer rfc VACÍO y
// un "folio" (código de autorización) que a veces se REPITE entre recibos del
// MISMO proveedor (el modelo toma un número de tienda/soporte constante, o el
// recibo no trae auth code) → dos recibos distintos del mismo proveedor con
// montos diferentes se colapsaban en una sola fila y se perdían. Para tickets
// agregamos el nombre de archivo (único por documento): recibos distintos ya no
// se fusionan, pero re-subir el MISMO archivo sí deduplica. Los pedimentos
// (esTicket=false) conservan su folio único real y no se ven afectados.
const gastoDedupKey = (g) => {
  const base = `${g.rfc || ''}|${g.noFactura || ''}`
  if (!g.esTicket) return base
  const src = g.pdfFile?.name || g.originalFileName || g.uuid || ''
  const monto = g.montoExtranjero || g.totalCFDI || g.total || 0
  return `${base}|${src}|${monto}`
}

// fetch con timeout (AbortController). Sin esto, una llamada OCR que se cuelga
// (red lenta, Groq atorado) dejaría el spinner girando para siempre = sistema
// trabado. Con timeout falla limpio, reintenta y/o se reporta al usuario.
const fetchConTimeout = async (url, opts = {}, ms = 90_000) => {
  const ctrl = new AbortController()
  const id = setTimeout(() => ctrl.abort(), ms)
  try {
    return await fetch(url, { ...opts, signal: ctrl.signal })
  } finally {
    clearTimeout(id)
  }
}

// Instancia única del módulo wasm de libheif-js (ver index.html) — decodificar
// es costoso, así que la promesa de inicialización se reutiliza entre fotos.
let _libheifModPromise = null
const getLibheifMod = () => {
  if (!window.libheif) return null
  if (!_libheifModPromise) _libheifModPromise = window.libheif()
  return _libheifModPromise
}

// Decodifica un File HEIC/HEIF a un Blob JPEG usando libheif-js (wasm). Chrome
// y Firefox no traen códec HEIC nativo (solo Safari, vía <canvas>), así que sin
// esto las fotos de iPhone subidas desde esos navegadores fallaban en el OCR.
const heicToJpegBlob = async (file) => {
  const heifMod = await getLibheifMod()
  const buf = new Uint8Array(await file.arrayBuffer())
  const decoder = new heifMod.HeifDecoder()
  const images = decoder.decode(buf)
  if (!images || !images.length) throw new Error('HEIC sin imágenes decodificables')
  const image = images[0]
  const width = image.get_width()
  const height = image.get_height()
  const canvas = document.createElement('canvas')
  canvas.width = width
  canvas.height = height
  const ctx = canvas.getContext('2d')
  const imageData = ctx.createImageData(width, height)
  await new Promise((resolve, reject) => {
    image.display(imageData, (displayData) => {
      if (!displayData) return reject(new Error('HEIF processing error'))
      resolve()
    })
  })
  ctx.putImageData(imageData, 0, 0)
  return new Promise(resolve => canvas.toBlob(resolve, 'image/jpeg', 0.92))
}

const compressImage = async (file, maxWidth = 2000, quality = 0.85) => {
  const isHeic = /\.(heic|heif)$/i.test(file.name) || /^image\/hei[cf]/i.test(file.type || '')
  let sourceFile = file
  if (isHeic && window.libheif) {
    // Ruta principal (todos los navegadores): decodificamos el HEIC con
    // libheif-js a un JPEG antes del resize. Si falla, NO tronamos: dejamos el
    // archivo original y el <canvas> de abajo lo intenta (Safari sí lo lee) y,
    // si tampoco puede, el img.onerror reporta un mensaje claro (no "undefined").
    try {
      const jpegBlob = await heicToJpegBlob(file)
      if (!jpegBlob) throw new Error('conversión HEIC vacía')
      sourceFile = new File(
        [jpegBlob],
        file.name.replace(/\.(heic|heif)$/i, '.jpg'),
        { type: 'image/jpeg' }
      )
    } catch (err) {
      console.warn('Conversión HEIC falló:', file.name, err)
      sourceFile = file
    }
  } else if (!isHeic && !file.type.startsWith('image/')) {
    return file
  }
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
              sourceFile.name.replace(/\.(heic|heif|webp|png)$/i, '.jpg'),
              { type: 'image/jpeg' }
            )
            resolve(compressed)
          },
          'image/jpeg',
          quality
        )
      }
      // Rechazamos con un Error real (no el Event del handler) para que el
      // modal muestre un mensaje útil y no "undefined". Un HEIC que llega aquí
      // es porque libheif no cargó y el navegador no lo decodifica (Chrome/FF).
      img.onerror = () => reject(new Error(
        isHeic
          ? 'No se pudo leer la foto HEIC en este navegador. Ábrela en Safari o convierte la foto a JPG antes de subirla.'
          : 'No se pudo leer la imagen (formato no soportado o archivo dañado).'
      ))
      img.src = e.target.result
    }
    reader.onerror = () => reject(new Error('No se pudo leer el archivo de la foto.'))
    reader.readAsDataURL(sourceFile)
  })
}

// Rasteriza la 1ª página de un PDF a un File JPEG chico usando pdf.js (CDN).
// Se usa SOLO para PDFs grandes (escaneos) antes de mandarlos al OCR: subir un
// JPEG de ~150KB es mucho más rápido que un PDF de 1.8MB. Devuelve null si
// pdf.js no está disponible o falla, para caer al envío del PDF original.
const pdfFirstPageToJpeg = async (file, maxDim = 1500, quality = 0.8) => {
  const pdfjs = window.pdfjsLib
  if (!pdfjs) return null
  try {
    const buf = await file.arrayBuffer()
    const pdf = await pdfjs.getDocument({ data: buf }).promise
    const page = await pdf.getPage(1)
    let viewport = page.getViewport({ scale: 1 })
    const scale = Math.min(maxDim / viewport.width, maxDim / viewport.height, 3)
    viewport = page.getViewport({ scale: Math.max(scale, 0.1) })
    const canvas = document.createElement('canvas')
    canvas.width = Math.round(viewport.width)
    canvas.height = Math.round(viewport.height)
    const ctx = canvas.getContext('2d')
    // Fondo blanco: los PDFs escaneados pueden tener transparencia.
    ctx.fillStyle = '#ffffff'
    ctx.fillRect(0, 0, canvas.width, canvas.height)
    await page.render({ canvasContext: ctx, viewport }).promise
    const blob = await new Promise(res => canvas.toBlob(res, 'image/jpeg', quality))
    if (!blob) return null
    return new File([blob], file.name.replace(/\.pdf$/i, '.jpg'), { type: 'image/jpeg' })
  } catch (err) {
    console.warn('pdf.js rasterize falló:', file?.name, err)
    return null
  }
}

// Lee el TEXTO embebido de las primeras páginas de un PDF usando pdf.js (CDN),
// SIN red y SIN OCR. Sirve para enlazar un PDF a su XML cuando el NOMBRE del
// archivo no comparte ningún token (folio/UUID/RFC) — buscamos esos tokens
// dentro del contenido. Devuelve el texto en MAYÚSCULAS y sin espacios para
// matchear igual que las llaves normalizadas. Devuelve '' si pdf.js no está
// disponible, si el PDF es un escaneo sin capa de texto, o si falla.
const pdfTextContent = async (file, maxPages = 2) => {
  const pdfjs = window.pdfjsLib
  if (!pdfjs) return ''
  try {
    const buf = await file.arrayBuffer()
    const pdf = await pdfjs.getDocument({ data: buf }).promise
    const pages = Math.min(maxPages, pdf.numPages)
    let texto = ''
    for (let i = 1; i <= pages; i++) {
      const page = await pdf.getPage(i)
      const content = await page.getTextContent()
      texto += content.items.map(it => it.str).join(' ')
    }
    // Normaliza igual que las llaves: sin acentos, sin separadores, MAYÚSCULAS.
    return texto.normalize('NFD').replace(/\p{Diacritic}/gu, '')
      .toUpperCase().replace(/[^A-Z0-9]/g, '')
  } catch (err) {
    console.warn('pdf.js text extract falló:', file?.name, err)
    return ''
  }
}

// NIVEL 6 de matching (corre DESPUÉS de linkPdfsExclusive): para los gastos que
// quedaron SIN PDF y los PDFs huérfanos, lee el TEXTO embebido del PDF y enlaza
// si contiene el UUID, el folio (noFactura) o el RFC del gasto. Cierra el caso
// "el PDF fue renombrado a algo arbitrario (scan_001.pdf) pero adentro trae los
// datos correctos". Exclusivo: un PDF → un solo gasto. Muta gasto.pdfFile /
// gasto.tienePDF. Devuelve los PDFs que siguen sin pareja (candidatos a OCR).
const linkPdfsByContent = async (gastos, orphanPdfs) => {
  const norm = s => (s || '').normalize('NFD').replace(/\p{Diacritic}/gu, '')
    .toUpperCase().replace(/[^A-Z0-9]/g, '')
  const sinPdf = gastos.filter(g => !g.pdfFile)
  if (sinPdf.length === 0 || !orphanPdfs || orphanPdfs.length === 0) return orphanPdfs || []

  // Cache de texto por PDF (se lee una sola vez aunque varios gastos lo evalúen).
  const pool = [...orphanPdfs]
  const textCache = new Map()
  const getText = async (p) => {
    if (!textCache.has(p)) textCache.set(p, await pdfTextContent(p))
    return textCache.get(p)
  }

  for (const g of sinPdf) {
    if (g.pdfFile) continue
    const uuid  = norm(g.uuid)
    const folio = norm(g.noFactura)
    const rfc   = norm(g.rfc)
    for (let i = 0; i < pool.length; i++) {
      const p = pool[i]
      const txt = await getText(p)
      if (!txt) continue
      const hit =
        (uuid.length  >= 16 && txt.includes(uuid))  ||
        (folio.length >= 4  && txt.includes(folio)) ||
        (rfc.length   >= 10 && txt.includes(rfc))
      if (hit) {
        g.pdfFile = p
        g.tienePDF = true
        pool.splice(i, 1)
        break
      }
    }
  }
  return pool
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

// Extrae la fecha del NOMBRE de archivo de facturas de EE.UU. (Microsoft las
// nombra "Microsoft G160071537 Office 05-21-26.pdf", con la fecha en MM-DD-YY).
// Es una fuente determinista que evita que el OCR invierta día/mes y mande el
// gasto a un mes anterior equivocado. Desambiguación: si el primer número es
// > 12 es DD-MM; si el segundo es > 12 es MM-DD; si ambos son ≤ 12 asumimos el
// orden US (MM-DD) con el que Microsoft nombra sus archivos. Devuelve
// 'YYYY-MM-DD' (formato interno) o '' si el nombre no trae fecha reconocible.
function fechaDesdeNombreUS(name) {
  const m = (name || '').match(/(\d{1,2})[-_.](\d{1,2})[-_.](\d{2,4})/)
  if (!m) return ''
  const a = parseInt(m[1], 10), b = parseInt(m[2], 10)
  let y = parseInt(m[3], 10)
  if (y < 100) y += 2000
  let mes, dia
  if (a > 12)      { dia = a; mes = b }   // DD-MM
  else if (b > 12) { mes = a; dia = b }   // MM-DD
  else             { mes = a; dia = b }   // ambiguo → MM-DD (nombre US de Microsoft)
  if (mes < 1 || mes > 12 || dia < 1 || dia > 31) return ''
  return `${y}-${String(mes).padStart(2, '0')}-${String(dia).padStart(2, '0')}`
}

// Corrige el AÑO de una fecha ISO (YYYY-MM-DD) leída por OCR. El modelo a veces
// lee mal el año de un ticket (p.ej. "2023" en vez de "2026"); en un reporte de
// gastos las facturas son SIEMPRE del año en curso, salvo el cruce
// diciembre→enero (un ticket de nov/dic capturado en ene/feb del año siguiente).
// Conserva el día y el mes del OCR (que suelen venir bien) y solo ajusta el año.
// SOLO se usa en el path de OCR (tickets/pedimentos); los CFDI traen su fecha
// fiscal correcta del XML y no pasan por aquí.
function forzarAñoEnCurso(iso) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(iso || ''))
  if (!m) return iso
  const mes = m[2], dia = m[3]
  const hoy = new Date()
  let anio = hoy.getFullYear()
  // Cruce dic-ene: si HOY es ene/feb y el ticket es de nov/dic → año anterior.
  if ((hoy.getMonth() + 1) <= 2 && parseInt(mes, 10) >= 11) anio -= 1
  return `${anio}-${mes}-${dia}`
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

/* Parser del Excel "Saldos AAAA" usado por los colaboradores especiales.
   Cada pestaña suele traer headers Fecha|Tipo|Folio|Factura|Concepto|
   Ingreso|Egreso en alguna fila (varía por hoja: r3 a r177). Algunas
   pestañas legacy (Banamex, Intercam) tienen columna 'Desc' en vez de
   'Tipo' y sin 'Folio'; las parseamos igual aunque su folio quede
   vacío.

   Devuelve [{ sheet, fecha, tipo, folio, factura, concepto, egreso }]. */
/* Convierte la fecha cruda del Saldos (Date | número serial | string) al
   formato interno YYYY-MM-DD que usa la app para fechaCobro. */
function saldosFechaToIso(d) {
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

function parseSaldosXLSX(arrayBuffer) {
  const wb = XLSX.read(arrayBuffer, { type: 'array', cellDates: true })
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
        // columna fuera del set principal (Tipo / Concepto / Factura).
        let moneda = 'MXN'
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

/* Normaliza una factura para comparación: sin espacios, guiones, slashes,
   puntos; uppercase. "48 3993 64000" y "48399364000" matchearán. */
const normFactura = (s) => String(s || '').replace(/[\s\-/.]/g, '').toUpperCase()

/* Normaliza un proveedor/concepto para comparación laxa: sin acentos, sólo
   A-Z0-9, uppercase. Permite buscar el proveedor del gasto dentro del
   concepto del Saldos (que suele traer el nombre embebido en texto libre). */
const normProv = (s) => String(s || '')
  .normalize('NFD').replace(/\p{Diacritic}/gu, '')
  .toUpperCase().replace(/[^A-Z0-9]/g, '')

/* Puntúa un renglón del Saldos como candidato para un gasto. Se usa SÓLO
   para DESEMPATAR cuando varios renglones comparten el MISMO número de
   factura (folios repetidos entre proveedores o cuentas — el folio CFDI es
   por emisor, así que dos proveedores distintos pueden traer el mismo
   número). Señales de más a menos decisiva:
     1. monto exacto del egreso ≈ total del CFDI (±$0.01)  → la más fuerte
     2. proveedor del gasto presente en el concepto del Saldos
     3. misma moneda (USD↔USD, MXN↔MXN)
     4. cercanía de fecha factura↔cobro (desempate suave)
   Devuelve { score, amountMatch, provMatch } para poder marcar después si la
   colisión se resolvió por una señal fuerte (monto/proveedor) o sólo débil. */
function scoreSaldosCandidate(g, sRow) {
  const gMoneda = (g.monedaCodigo || g.moneda || 'MXN').toString().toUpperCase()
  const sMoneda = (sRow.moneda || 'MXN').toString().toUpperCase()
  const gTotal  = Number(g.totalCFDI) || 0
  const sEgreso = Number(sRow.egreso) || 0
  const amountMatch = gTotal > 0 && sEgreso > 0 && Math.abs(sEgreso - gTotal) <= 0.01
  const gProv = normProv(g.proveedor)
  const provMatch = gProv.length >= 4 && normProv(sRow.concepto).includes(gProv)

  let score = 0
  if (amountMatch)         score += 10000
  if (provMatch)           score += 4000
  if (sMoneda === gMoneda) score += 1000
  const gIso = String(g.fechaFac || '').slice(0, 10)
  const sIso = saldosFechaToIso(sRow.fecha) || ''
  if (/^\d{4}-\d{2}-\d{2}$/.test(gIso) && /^\d{4}-\d{2}-\d{2}$/.test(sIso)) {
    const days = Math.abs(new Date(gIso + 'T00:00:00Z') - new Date(sIso + 'T00:00:00Z')) / 86400000
    if (!Number.isNaN(days)) score += Math.max(0, 200 - days)
  }
  return { score, amountMatch, provMatch }
}

/* Cotejo: por cada gasto busca la fila del Saldos con la misma factura
   normalizada, MISMO AÑO y MISMA MONEDA. Esto evita matchear una factura
   2026 contra un renglón Saldos 2025, y evita matchear USD contra MXN.
   Devuelve los matched, sin-match por ambos lados, y marca las
   discrepancias de tipo para que el usuario decida. */
function cotejarConSaldos(saldosRows, lista) {
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

  // 3) Índice por factura normalizada → posiciones dentro de filteredSaldos.
  //    Skip filas con factura 'NA' (típicamente nóminas/comisiones) — no son
  //    matchables contra los gastos y meterlas crearía falsos positivos si
  //    algún gasto manual tuviera noFactura='NA'.
  const byFactura = new Map()
  filteredSaldos.forEach((row, idx) => {
    const k = normFactura(row.factura)
    if (!k || k === 'NA') return
    if (!byFactura.has(k)) byFactura.set(k, [])
    byFactura.get(k).push(idx)
  })

  const matched = []
  const sinMatchGastos = []
  const usedSaldos = new Set()

  // 4) Match eligiendo el MEJOR candidato entre los renglones libres con esa
  //    factura. Con un solo candidato es ese renglón (el folio basta). Con
  //    VARIOS (mismo número de factura repetido entre proveedores/cuentas),
  //    scoreSaldosCandidate desempata por monto/proveedor/moneda/fecha para
  //    no vincular el gasto con la factura equivocada y mezclar póliza,
  //    banco, fecha de cobro o tipo de otra factura que no es.
  //    Greedy de un paso (igual que el resto del cotejo vía usedSaldos): cada
  //    gasto toma su mejor renglón libre. En la colisión real (proveedores
  //    distintos con el mismo folio) cada gasto prefiere fuertemente SU propio
  //    renglón por monto+proveedor, así que el orden no importa.
  lista.forEach((g) => {
    const k = normFactura(g.noFactura)
    const candidates = k ? (byFactura.get(k) || []) : []

    let saldosIdx, best = null
    for (const idx of candidates) {
      if (usedSaldos.has(idx)) continue
      const cand = scoreSaldosCandidate(g, filteredSaldos[idx])
      if (!best || cand.score > best.score) { best = cand; saldosIdx = idx }
    }

    if (saldosIdx === undefined) {
      sinMatchGastos.push(g)
      return
    }
    usedSaldos.add(saldosIdx)
    const sRow = filteredSaldos[saldosIdx]
    const tipoDiffers = !!sRow.tipo && !!g.tipo &&
      sRow.tipo.toLowerCase() !== g.tipo.toLowerCase()
    // Colisión = había más de un renglón del Saldos con este mismo número de
    // factura. La marcamos (y si se ancló por una señal fuerte) para que el
    // usuario pueda verificar en el modal que se vinculó la correcta.
    const facturaColision = candidates.length > 1
    const colisionResuelta = best.amountMatch || best.provMatch
    matched.push({ gastoId: g.id, gasto: g, saldosRow: sRow, tipoDiffers, facturaColision, colisionResuelta })
  })

  // 5) FALLBACK match: para gastos que no encontraron factura en el Saldos,
  //    intenta matchear contra los renglones con factura='NA' (o vacía) por
  //    PROVEEDOR + MONTO + AÑO-MES. Cubre el caso donde el Saldos no tiene
  //    el folio de la factura pero sí registra el movimiento del banco.
  const stillSinMatch = []
  for (const g of sinMatchGastos) {
    const gProv = normProv(g.proveedor)
    if (gProv.length < 8) { stillSinMatch.push(g); continue }
    const gTotal = Number(g.totalCFDI) || 0
    const gYM = String(g.fechaFac || '').slice(0, 7)

    let foundIdx = -1
    for (let i = 0; i < filteredSaldos.length; i++) {
      if (usedSaldos.has(i)) continue
      const row = filteredSaldos[i]
      const facNorm = normFactura(row.factura)
      if (facNorm && facNorm !== 'NA') continue  // solo filas sin factura clara
      const rowYM = (saldosFechaToIso(row.fecha) || '').slice(0, 7)
      if (rowYM && gYM && rowYM !== gYM) continue
      const concNorm = normProv(row.concepto)
      if (!concNorm.includes(gProv)) continue
      // Tolera diferencia de ±$0.01 para evitar problemas de redondeo.
      if (gTotal > 0 && Math.abs((Number(row.egreso) || 0) - gTotal) > 0.01) continue
      foundIdx = i
      break
    }

    if (foundIdx === -1) { stillSinMatch.push(g); continue }
    usedSaldos.add(foundIdx)
    const sRow = filteredSaldos[foundIdx]
    const tipoDiffers = !!sRow.tipo && !!g.tipo &&
      sRow.tipo.toLowerCase() !== g.tipo.toLowerCase()
    matched.push({ gastoId: g.id, gasto: g, saldosRow: sRow, tipoDiffers })
  }
  sinMatchGastos.length = 0
  sinMatchGastos.push(...stillSinMatch)

  const sinMatchSaldos = filteredSaldos.filter((_, idx) => !usedSaldos.has(idx))
  return { matched, sinMatchGastos, sinMatchSaldos }
}

function parseCFDI(xmlText, xmlFile, pdfFiles, colaborador) {
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
  // NoIdentificacion del concepto: identificador único de la transacción (p.ej.
  // el ID de cruce/peaje del Fondo Nacional de Infraestructura). Se usa como
  // número de factura cuando el CFDI no trae Serie/Folio (ver noFactura abajo).
  let noIdentificacion = conceptoEl ? (ga(conceptoEl, 'NoIdentificacion', 'noidentificacion') || '') : ''
  if (!noIdentificacion) {
    const nm = xmlText.match(/NoIdentificacion="([^"]+)"/i); if (nm) noIdentificacion = nm[1]
  }
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
  // ishIeps = impuestos que SE SUMAN al total y no son IVA: ISH (locales de
  // hoteles), IEPS y el raro ISR trasladado (001 en Traslados).
  // El "ISR" del Excel es el ISR RETENIDO (retencionISR), que RESTA del total
  // y va en su propia columna. Se separan porque tienen signo distinto.
  //
  // ⚠️ IEPS federal viaja en <Traslado Impuesto="003"> (gasolinas, alimentos con
  // IEPS como refrescos/botanas en OXXO/Tiendas Extra, gasolineras). Antes solo
  // sumábamos '001' y el 003 se perdía → el usuario lo capturaba a mano. Ahora
  // el bucket ishIeps agrupa 001 (ISR trasladado) + 003 (IEPS); 002 (IVA) queda
  // aparte en su propio bucket.
  let   ishIeps       = sumByTipo(trasladosBox,   'traslado',  t => t === '001' || t === '003')
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

  // El ISR retenido (retencionISR) se queda como retención y se muestra en su
  // propia columna ISR del Excel (resta del total). Antes había un override que
  // lo movía a "traslado" para ciertos RFC, pero eso lo dejaba de restar y
  // descuadraba el total (p.ej. Volare: 494.12+79.06−6.18 = 567.00, no 579.36).

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
    if      (ln === 'trasladoslocales')  ishIeps      += parseFloat(ga(el, 'Importe', 'importe') || '0') || 0
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
  //
  // ⚠️ El importe SALE de GralConsumos, NO de GralImporte: GralImporte ya
  // trae el IEPS sumado (GralImporte = GralConsumos + GralIEPS). Si usáramos
  // GralImporte, el IEPS se contaría DOS veces — una dentro del importe y otra
  // en la columna ISH/IEPS — inflando el total. GralConsumos es el subtotal de
  // combustible sin IEPS ni IVA, de modo que
  //   importe(GralConsumos) + iva(GralImpuesto) + ieps(GralIEPS) = GralTotal.
  let isEDC = false, dispersion = null
  for (const el of doc.querySelectorAll('*')) {
    if (!el.localName) continue
    const ln = el.localName.toLowerCase()
    if      (ln === 'estadodecuentacombustible')   isEDC = true
    else if (ln === 'dispersion' && !dispersion)   dispersion = el
  }
  if (isEDC && dispersion) {
    importe         = parseFloat(ga(dispersion, 'GralConsumos', 'gralConsumos') || '0') || 0
    iva             = parseFloat(ga(dispersion, 'GralImpuesto', 'gralImpuesto') || '0') || 0
    ishIeps        += parseFloat(ga(dispersion, 'GralIEPS',     'gralIEPS')     || '0') || 0
    totalCFDI       = parseFloat(ga(dispersion, 'GralTotal',    'gralTotal')    || '0') || 0
    conceptoClasif  = 'Combustible'
  }

  // ── Regla especial: facturas de Total Play (RFC TPT890516JP5) ──
  // El CFDI se timbra por el total bruto de servicios (Total = 1418.64,
  // SubTotal + IVA), pero el cargo REAL del mes — lo que cobra el banco — es el
  // "Cargos del Mes", ya neto de descuentos/promociones. Total Play lo guarda
  // en su Addenda propia como <tp:Cuerpo SaldoDelMes="1378.65" Subtotal="...">.
  // Tomamos ese monto para que el cotejo cuadre con el banco y derivamos el IVA
  // del subtotal de la Addenda (importe + IVA = total, igual que el Excel).
  // Equivalente XML de la regla OCR de Total Play; aquí es exacto y sin OCR.
  const esTotalPlayXML = rfc === 'TPT890516JP5' || /TOTAL\s*PLAY/i.test(proveedor || '')
  if (esTotalPlayXML) {
    let cuerpoTP = null
    for (const el of doc.querySelectorAll('*')) {
      if (!el.localName || el.localName.toLowerCase() !== 'cuerpo') continue
      if ((parseFloat(ga(el, 'SaldoDelMes', 'saldodelmes') || '0') || 0) > 0) { cuerpoTP = el; break }
    }
    if (cuerpoTP) {
      const saldoMes = parseFloat(ga(cuerpoTP, 'SaldoDelMes', 'saldodelmes') || '0') || 0
      const subMes   = parseFloat(ga(cuerpoTP, 'Subtotal',    'subtotal')    || '0') || 0
      totalCFDI = saldoMes
      importe   = subMes > 0 ? subMes : importe
      iva       = parseFloat((totalCFDI - importe).toFixed(2))
    }
  }

  // Buscar PDF asociado — AQUÍ solo señales FUERTES y exclusivas:
  //   1. Coincidencia exacta de nombre base (XML.xml ↔ XML.pdf)
  //   2. UUID del CFDI dentro del nombre del PDF
  // Las heurísticas "fuzzy" (substring / prefijo / folio / RFC) se aplican
  // DESPUÉS en linkPdfsExclusive(), de forma exclusiva (un PDF → un gasto).
  //
  // ⚠️ Por qué NO se hace fuzzy aquí: parseCFDI corre por-XML y devuelve el
  // PRIMER PDF que pasa. El viejo match por "primeros 15 caracteres" tomaba
  // el prefijo del NOMBRE DEL PROVEEDOR ("aeroenlacesnaci…"), así que con
  // 2+ facturas del mismo proveedor en el mes (Aeroenlaces, Aerovías, …)
  // TODOS los XML robaban el PDF de la primera factura → un PDF duplicado en
  // varias filas y los PDFs reales perdidos. El nombre base exacto ya
  // distingue cada factura por su folio, que es lo correcto.
  //
  // stripDiacritics: NFD + lower + sin acentos para que XMLs/PDFs con
  // distinta codificación Unicode (PDF de Safari/WhatsApp en NFC vs XML en
  // NFD) matchéen igual ("Representación.pdf" ↔ "Representación.xml").
  const stripDiacritics = s => (s || '')
    .normalize('NFD')
    .replace(/\p{Diacritic}/gu, '')
  const base    = stripDiacritics(xmlFile.name.replace(/\.xml$/i, '').toLowerCase())
  const pdfFile = pdfFiles.find(f => {
    const pdfBase = stripDiacritics(f.name.replace(/\.pdf$/i, '').toLowerCase())
    if (pdfBase === base) return true
    if (uuid && stripDiacritics(f.name.toUpperCase()).includes(uuid.toUpperCase())) return true
    return false
  }) || null

  const fechaFac = (ga(comp, 'Fecha', 'fecha') || '').slice(0, 10)
  // Moneda + tipo de cambio del CFDI. Para CFDIs en USD / otra divisa:
  //   - importe / iva / retenciones / totalCFDI quedan en 0 en la UI hasta
  //     que el usuario aplique un T/C (se hace en el Excel, no aquí).
  //   - los montos originales en USD se preservan en importeUSD / ivaUSD /
  //     retencionesUSD / montoUSD para que el backend de Excel los pueda
  //     multiplicar por la celda editable de T/C (=USD * $N$7).
  const monedaXML      = (ga(comp, 'Moneda', 'moneda') || 'MXN').toUpperCase()
  const tipoCambioXML  = parseFloat(ga(comp, 'TipoCambio', 'tipocambio') || '0') || 0
  const esExtranjera   = monedaXML !== 'MXN' && monedaXML !== 'XXX'
  const importeUSD     = esExtranjera ? importe     : 0
  const ivaUSD         = esExtranjera ? iva         : 0
  const retencionesUSD = esExtranjera ? retenciones : 0
  // Combinado (ISH/IEPS) para la tabla de UI / portapapeles / compat. El ISR
  // retenido vive aparte en retencionISR (su propia columna en el Excel).
  const isrTrasladado  = ishIeps
  return {
    id: genId(),
    rfc,
    proveedor,
    noFactura: (() => {
      const serie = ga(comp, 'Serie', 'serie') || ''
      let   folio = ga(comp, 'Folio', 'folio') || ''
      // REFIEL pad: el CFDI viene con folio sin zero-padding (e.g. 1406)
      // pero el PDF y el archivo de Saldos lo usan padeado a 5 dígitos
      // (E01406). Sin este pad, el cotejo con Saldos no matchea.
      if (rfc === 'REF2208125E6' && /^\d{1,5}$/.test(folio)) {
        folio = folio.padStart(5, '0')
      }
      // CFDIs de Odoo (Vauxoo, MBGE, etc.): la Serie termina en "/"
      // ("INV/2026/") y el Folio es solo el consecutivo SIN ceros ("591"). El
      // número fiscal completo que imprime el PDF lo padea a 5 dígitos
      // ("INV/2026/00591"). Replicamos ese pad para que la FACTURA del Excel y
      // el nombre del archivo conserven los ceros a la izquierda.
      if (serie.endsWith('/') && /^\d{1,5}$/.test(folio)) {
        folio = folio.padStart(5, '0')
      }
      if (serie || folio) return serie + folio
      // Peaje del Fondo Nacional de Infraestructura (RFC FNI970829JR9): su CFDI
      // NO trae Serie/Folio. El Serie+Folio (p.ej. FNPE72984235) solo está
      // impreso en el PDF, pero MUY frecuentemente también en el NOMBRE del
      // archivo. Lo sacamos del nombre — instantáneo, SIN red — para no depender
      // del OCR del PDF (que bloqueaba la carga). Buscamos en el nombre del XML
      // y del PDF enlazado.
      if (rfc === 'FNI970829JR9') {
        const nombres = `${xmlFile?.name || ''} ${pdfFile?.name || ''}`
        const m = nombres.match(/\bFNPE\d{4,}\b/i)
        if (m) return m[0].toUpperCase()
      }
      // Si no vino en el nombre, usamos el NoIdentificacion del concepto (único
      // por cruce) en vez del placeholder SN-xxxx. Acotado a FNI a propósito:
      // en otros emisores el NoIdentificacion puede ser un SKU repetido entre
      // facturas y colisionaría en la dedup (rfc|noFactura).
      if (rfc === 'FNI970829JR9' && noIdentificacion) return noIdentificacion
      // Sin Serie/Folio ni NoIdentificacion: usar últimos 4 chars del UUID para
      // que cada factura tenga un identificador único y no colisione en la dedup.
      const uuidSuffix = (uuid || '').replace(/-/g, '').slice(-4).toUpperCase()
      return uuidSuffix ? `SN-${uuidSuffix}` : `SN-${genId().slice(0, 4).toUpperCase()}`
    })(),
    fechaFac,
    concepto:   conceptoClasif,
    tipo: rfc === 'FNI970829JR9'
      ? tipoPeajeFNI(colaborador)
      : autoDetectTipo(proveedor, descripcionFirstLine, usaTiposExtendidos(colaborador) ? undefined : categoriaEfectiva(colaborador), claveProdServ, rfc),
    importe:        esExtranjera ? 0 : importe,
    iva:            esExtranjera ? 0 : iva,
    isrTrasladado,
    ishIeps:        esExtranjera ? 0 : ishIeps,
    retencionISR:   esExtranjera ? 0 : retencionISR,
    retencionIVA,
    retenciones:    esExtranjera ? 0 : retenciones,
    totalCFDI:      esExtranjera ? 0 : totalCFDI,
    propinaPorcentaje: 0,
    montoPropina: 0,
    fechaCobro: fechaFac,
    // Telcel (Radiomóvil Dipsa) timbra sus CFDI con FormaPago="99" (por definir),
    // pero SMTO siempre los paga por transferencia → forzamos "03".
    formaPago:  rfc === 'RDI841003QJ4' ? '03' : (ga(comp, 'FormaPago') || '04'),
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
    montoUSD:           esExtranjera ? totalCFDI : 0,
    montoExtranjero:    esExtranjera ? totalCFDI : 0,
    importeUSD,
    ivaUSD,
    retencionesUSD,
    tipoCambio:         esExtranjera ? tipoCambioXML : 0,
    moneda:             monedaXML,
    monedaCodigo:       monedaXML,
    esMonedaExtranjera: esExtranjera,
    banco:              defaultBancoFor(colaborador),
    // Snapshot del total facturado al momento de leer el CFDI. validarBanco
    // sobreescribe totalCFDI con el monto real del banco; este campo se
    // queda intacto para exponer la diferencia 'cobrado vs facturado' en
    // la columna DIFERENCIA del Excel (solo Clara MXN Credito).
    montoFacturado:     totalCFDI,
    // CFDIs no traen propina, así que la propina "esperada" es 0. Si el
    // usuario agrega propina manualmente después, aparece como DIFERENCIA
    // (= montoPropina actual - montoPropinaOriginal = todo lo agregado).
    montoPropinaOriginal: 0,
  }
}

/* ═══════════════════════════════════════════════════
   VINCULACIÓN EXCLUSIVA PDF ↔ GASTO
═══════════════════════════════════════════════════ */
// Asigna cada PDF a UN SOLO gasto y cada gasto a UN SOLO PDF, por niveles
// de confianza (fuerte → débil). Un PDF reclamado en un nivel sale del pool,
// así una heurística débil NUNCA puede robar un PDF que una señal fuerte ya
// reclamó. Esto reemplaza las "segundas pasadas" por-PDF (no exclusivas) que
// duplicaban un PDF en varias filas cuando un proveedor tenía 2+ facturas en
// el mes. Muta gasto.pdfFile / gasto.tienePDF. Devuelve los PDFs sin pareja
// (candidatos a OCR).
function linkPdfsExclusive(gastos, pdfFiles) {
  const norm     = s => (s || '').normalize('NFD').replace(/\p{Diacritic}/gu, '')
    .toLowerCase().replace(/[^a-z0-9]/g, '')
  const baseName = name => norm(String(name || '').replace(/\.(pdf|xml)$/i, ''))

  // Pool de PDFs aún sin reclamar (respeta asignaciones previas por nombre).
  const pool = pdfFiles.filter(p => !gastos.some(g => g.pdfFile?.name === p.name))
  const claim = (g, p) => {
    g.pdfFile  = p
    g.tienePDF = true
    const i = pool.indexOf(p)
    if (i >= 0) pool.splice(i, 1)
  }
  // Un nivel: para cada gasto aún sin PDF, toma el PRIMER PDF del pool que
  // cumpla `ok`. Exclusivo — el PDF reclamado abandona el pool de inmediato.
  const tier = ok => {
    for (const g of gastos) {
      if (g.pdfFile) continue
      const p = pool.find(pf => ok(g, pf))
      if (p) claim(g, p)
    }
  }

  // 1. Nombre base idéntico (la señal más fuerte; distingue por folio).
  tier((g, p) => { const b = baseName(g.xmlFile?.name); return !!b && b === baseName(p.name) })
  // 2. UUID del CFDI dentro del nombre del PDF.
  tier((g, p) => { const u = norm(g.uuid); return u.length >= 16 && norm(p.name).includes(u) })
  // 3. Folio (noFactura) dentro del nombre del PDF — único por factura.
  tier((g, p) => { const f = norm(g.noFactura); return f.length >= 4 && norm(p.name).includes(f) })
  // 4. Un nombre base contiene al otro (PDF renombrado/truncado por el usuario).
  tier((g, p) => {
    const b = baseName(g.xmlFile?.name), pb = baseName(p.name)
    return b.length >= 8 && pb.length >= 8 && (b.includes(pb) || pb.includes(b))
  })
  // 5. RFC en el nombre del PDF — señal DÉBIL (igual para todo el proveedor),
  //    por eso va al final: solo enlaza un PDF huérfano a un gasto aún suelto.
  tier((g, p) => { const r = norm(g.rfc); return r.length >= 10 && norm(p.name).includes(r) })

  return pool
}

/* ═══════════════════════════════════════════════════
   COMPONENTE: BOTÓN PREMIUM
═══════════════════════════════════════════════════ */

function PremiumButton({ title, icon, variant = 'primary', isDisabled = false, onClick, id }) {
  return (
    <button
      id={id}
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
   RECORRIDO GUIADO — "Cómo utilizar"
   Spotlight sobre cada botón clave + popover con el paso.
═══════════════════════════════════════════════════ */

const TOUR_STEPS = [
  {
    targetId: 'tour-carpeta',
    icon: '📁',
    title: 'Paso 1 · Cargar Carpeta',
    body: 'Empieza aquí. Selecciona la carpeta que contiene tus archivos XML y PDF de las facturas. La app los empareja automáticamente y llena la tabla.',
  },
  {
    targetId: 'tour-banco',
    icon: '🏦',
    title: 'Paso 2 · Validar Banco',
    body: 'Antes, descarga el estado de cuenta de Clara del periodo de tu reporte. Luego súbelo aquí: la app concilia cada gasto con su cargo en el banco y detecta propinas.',
  },
  {
    targetId: 'tour-zip',
    icon: '📦',
    title: 'Paso 3 · Exportar a ZIP',
    body: 'Descarga el ZIP con tus facturas renombradas y el archivo de Excel. Esa carpeta completa es la que se le debe compartir al correo de Víctor: victor.aceves@smto.mx',
  },
]

function GuideTour({ step, setStep }) {
  const [rect, setRect] = useState(null)
  const total = TOUR_STEPS.length
  const current = step != null ? TOUR_STEPS[step] : null

  // Mide el botón objetivo (y re-mide en resize/scroll) para posicionar el
  // spotlight y el popover. getBoundingClientRect es relativo al viewport,
  // que es justo lo que necesita position:fixed.
  useEffect(() => {
    if (!current) return
    const measure = () => {
      const el = document.getElementById(current.targetId)
      if (!el) { setRect(null); return }
      el.scrollIntoView({ behavior: 'smooth', block: 'nearest', inline: 'center' })
      const r = el.getBoundingClientRect()
      setRect({ top: r.top, left: r.left, width: r.width, height: r.height })
    }
    measure()
    window.addEventListener('resize', measure)
    window.addEventListener('scroll', measure, true)
    return () => {
      window.removeEventListener('resize', measure)
      window.removeEventListener('scroll', measure, true)
    }
  }, [current])

  // Teclado: Esc cierra, flechas navegan.
  useEffect(() => {
    if (step == null) return
    const onKey = (e) => {
      if (e.key === 'Escape') setStep(null)
      else if (e.key === 'ArrowRight') setStep(s => Math.min((s ?? 0) + 1, total - 1))
      else if (e.key === 'ArrowLeft') setStep(s => Math.max((s ?? 0) - 1, 0))
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [step, setStep, total])

  if (step == null || !current) return null

  const PAD = 8
  const spot = rect ? {
    top: rect.top - PAD,
    left: rect.left - PAD,
    width: rect.width + PAD * 2,
    height: rect.height + PAD * 2,
  } : null

  // El popover se coloca debajo del spotlight (los botones viven en la barra
  // superior, así que siempre hay espacio abajo); si no, arriba. Clamp lateral.
  const POP_W = 340
  let popStyle = { top: '50%', left: '50%', transform: 'translate(-50%,-50%)' }
  let place = 'below'
  if (spot) {
    const belowSpace = window.innerHeight - (spot.top + spot.height)
    place = belowSpace > 230 ? 'below' : 'above'
    let left = spot.left + spot.width / 2 - POP_W / 2
    left = Math.max(16, Math.min(left, window.innerWidth - POP_W - 16))
    const top = place === 'below' ? spot.top + spot.height + 16 : spot.top - 16
    popStyle = place === 'below'
      ? { top, left, transform: 'none' }
      : { top, left, transform: 'translateY(-100%)' }
  }

  const isLast = step === total - 1
  const isFirst = step === 0

  return (
    <div className="tour-root">
      {/* Backdrop que bloquea la interacción con el resto de la app. */}
      <div className="tour-backdrop" />

      {/* Spotlight: la sombra gigante crea el oscurecido con un "hueco". */}
      {spot && (
        <div
          className="tour-spotlight"
          style={{ top: spot.top, left: spot.left, width: spot.width, height: spot.height }}
        />
      )}

      {/* Popover con la explicación del paso. */}
      <motion.div
        key={step}
        className={`tour-pop tour-pop-${place}`}
        style={popStyle}
        initial={{ opacity: 0, y: place === 'below' ? -8 : 8, scale: 0.97 }}
        animate={{ opacity: 1, y: 0, scale: 1 }}
        transition={{ type: 'spring', stiffness: 320, damping: 26 }}
      >
        <button className="tour-close" aria-label="Cerrar" onClick={() => setStep(null)}>
          <X size={16} />
        </button>
        <div className="tour-pop-head">
          <span className="tour-pop-ic">{current.icon}</span>
          <h3 className="tour-pop-title">{current.title}</h3>
        </div>
        <p className="tour-pop-body">{current.body}</p>

        <div className="tour-dots">
          {TOUR_STEPS.map((_, i) => (
            <span
              key={i}
              className={`tour-dot${i === step ? ' active' : ''}`}
              onClick={() => setStep(i)}
            />
          ))}
        </div>

        <div className="tour-actions">
          {!isFirst && (
            <button className="tour-btn-ghost" onClick={() => setStep(s => s - 1)}>
              Anterior
            </button>
          )}
          {isLast ? (
            <button className="tour-btn-primary" onClick={() => setStep(null)}>
              Entendido
            </button>
          ) : (
            <button className="tour-btn-primary" onClick={() => setStep(s => s + 1)}>
              Siguiente
            </button>
          )}
        </div>
      </motion.div>
    </div>
  )
}

/* ═══════════════════════════════════════════════════
   GUÍA "CÓMO DESCARGAR TU ESTADO DE CUENTA DE CLARA"
   Modal-carrusel paso a paso (imagen + texto) que abre un botón
   junto a "Validar Banco". Las capturas viven en public/clara-guia/
   y se cargan por ruta; si falta alguna, el paso muestra solo texto.
═══════════════════════════════════════════════════ */

const CLARA_GUIDE_STEPS = [
  {
    img: '/clara-guia/paso1.png',
    icon: '🌐',
    title: 'Paso 1 · Entra a Clara',
    body: (
      <>
        Abre tu navegador y entra a <b>app.clara.cc</b>. También puedes ir a{' '}
        <b>clara.com</b> y dar clic en <b>"Acceso Clientes"</b> (arriba a la derecha).
      </>
    ),
  },
  {
    img: '/clara-guia/paso2.png',
    icon: '🔑',
    title: 'Paso 2 · Inicia sesión',
    body: (
      <>
        En <b>"Ingresa a Clara"</b>, escribe tu <b>correo corporativo</b>{' '}
        (ej. nombre@smto.mx) y da clic en <b>Continuar</b>. También puedes usar{' '}
        <b>"Continuar con Google"</b> y elegir tu cuenta de Gmail, o entrar con tu{' '}
        <b>app de autenticador</b>.
      </>
    ),
  },
  {
    img: '/clara-guia/paso3.png',
    icon: '📄',
    title: 'Paso 3 · Abre "Movimientos"',
    body: (
      <>
        En el menú de la izquierda, ve a <b>Cuenta → Movimientos</b>. Ahí aparecen
        todos los cargos de las tarjetas.
      </>
    ),
  },
  {
    img: '/clara-guia/paso4.png',
    icon: '📅',
    title: 'Paso 4 · Filtra por el periodo del reporte',
    body: (
      <>
        Da clic en <b>"Añadir filtro" → "Fechas"</b> y elige{' '}
        <b>"Fechas personalizadas"</b>. Pon la <b>Fecha inicial</b> y la{' '}
        <b>Fecha final</b> del periodo de tu reporte (ej. 01/06/2026 → 30/06/2026).
        También puedes usar el atajo <b>"Estado de cuenta actual"</b>.
      </>
    ),
  },
  {
    img: '/clara-guia/paso5.png',
    icon: '⬇️',
    title: 'Paso 5 · Descarga el CSV',
    body: (
      <>
        Da clic en el botón azul <b>"Descargar"</b> (arriba a la derecha) y elige{' '}
        <b>".CSV — Solo datos de transacciones"</b>.{' '}
        <span className="clg-warn">
          Elige el <b>.CSV normal</b>, NO el ".CSV Enriquecido (Pro)".
        </span>{' '}
        Ese archivo es el que subes en <b>"Validar Banco"</b> aquí en la app.
      </>
    ),
  },
]

// Imagen de un paso: si la captura aún no existe en public/clara-guia/,
// se oculta sola (onError) y el paso queda solo con el texto.
function ClaraStepImage({ src, alt }) {
  const [broken, setBroken] = useState(false)
  if (broken) return null
  return (
    <img
      className="clg-shot"
      src={src}
      alt={alt}
      loading="lazy"
      onError={() => setBroken(true)}
    />
  )
}

function ClaraGuideModal({ open, onClose }) {
  const [step, setStep] = useState(0)
  const total = CLARA_GUIDE_STEPS.length

  // Reinicia al primer paso cada vez que se abre.
  useEffect(() => { if (open) setStep(0) }, [open])

  // Teclado: Esc cierra, flechas navegan.
  useEffect(() => {
    if (!open) return
    const onKey = (e) => {
      if (e.key === 'Escape') onClose()
      else if (e.key === 'ArrowRight') setStep(s => Math.min(s + 1, total - 1))
      else if (e.key === 'ArrowLeft') setStep(s => Math.max(s - 1, 0))
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [open, onClose, total])

  if (!open) return null

  const current = CLARA_GUIDE_STEPS[step]
  const isFirst = step === 0
  const isLast = step === total - 1

  return (
    <div className="clg-overlay" onClick={onClose}>
      <div className="clg-modal" onClick={e => e.stopPropagation()}>
        <button className="clg-close" aria-label="Cerrar" onClick={onClose}>
          <X size={18} />
        </button>

        <div className="clg-head">
          <span className="clg-badge">Guía · Clara</span>
          <h2 className="clg-title">Cómo descargar tu estado de cuenta</h2>
        </div>

        <div className="clg-figure">
          <motion.div
            key={step}
            initial={{ opacity: 0, x: 16 }}
            animate={{ opacity: 1, x: 0 }}
            transition={{ type: 'spring', stiffness: 320, damping: 30 }}
            className="clg-figure-inner"
          >
            <ClaraStepImage src={current.img} alt={current.title} />
            <div className="clg-step-head">
              <span className="clg-step-ic">{current.icon}</span>
              <h3 className="clg-step-title">{current.title}</h3>
            </div>
            <p className="clg-body">{current.body}</p>
          </motion.div>
        </div>

        <div className="clg-dots">
          {CLARA_GUIDE_STEPS.map((_, i) => (
            <span
              key={i}
              className={`clg-dot${i === step ? ' active' : ''}`}
              onClick={() => setStep(i)}
            />
          ))}
        </div>

        <div className="clg-actions">
          <span className="clg-counter">{step + 1} / {total}</span>
          <div className="clg-actions-btns">
            {!isFirst && (
              <button className="clg-btn-ghost" onClick={() => setStep(s => s - 1)}>
                Anterior
              </button>
            )}
            {isLast ? (
              <button className="clg-btn-primary" onClick={onClose}>
                Entendido
              </button>
            ) : (
              <button className="clg-btn-primary" onClick={() => setStep(s => s + 1)}>
                Siguiente
              </button>
            )}
          </div>
        </div>
      </div>
    </div>
  )
}

/* ═══════════════════════════════════════════════════
   COMPONENTE: CELDA NUMÉRICA EDITABLE
   Mantiene un "draft" (texto crudo) mientras el input tiene
   foco: muestra exactamente lo que se teclea, sin reformatear
   en cada tecla, así no salta el cursor ni se snapea a 0.
   Sólo al perder el foco vuelve al valor formateado. Vive a
   nivel de módulo (no dentro de GastoRow) para que su estado
   sobreviva los re-renders del padre.
═══════════════════════════════════════════════════ */

function NumCell({ g, upd, field, prefix, suffix, format = true, compact = false }) {
  const v = g[field]
  const [draft, setDraft] = useState(null)
  const formatted = v
    ? compact ? String(+Number(v).toFixed(2))
    : format  ? Number(v).toFixed(2)
    :           String(v)
    : ''
  const display = draft !== null ? draft : formatted
  return (
    <div className="num-cell">
      {prefix && <span className="sym">{prefix}</span>}
      <input
        type="text"
        inputMode="decimal"
        className="cell-in"
        value={display}
        placeholder="0"
        onFocus={e => {
          const el = e.target
          setDraft(v ? String(+Number(v).toFixed(2)) : '')
          requestAnimationFrame(() => el.select())
        }}
        onChange={e => {
          const raw = e.target.value
          // Permite vacío y números parciales ("1.", ".5", "10"); rechaza
          // cualquier otro carácter para que la edición fluya sin pelear.
          if (raw !== '' && !/^-?\d*\.?\d*$/.test(raw)) return
          setDraft(raw)
          upd(field, parseFloat(raw) || 0)
        }}
        onBlur={() => setDraft(null)}
      />
      {suffix && <span className="sym">{suffix}</span>}
    </div>
  )
}

/* ═══════════════════════════════════════════════════
   COMPONENTE: FILA DE LA TABLA
═══════════════════════════════════════════════════ */

const GastoRow = memo(function GastoRow({ g, update, remove, openPDF, tiposList, isSpecial }) {
  // Callbacks bound to this row's id. useCallback keeps their identity stable
  // across the parent's re-renders (update/remove are themselves stable), so
  // React.memo above can skip re-rendering every OTHER row on each keystroke —
  // only the edited row reconciles, which is what keeps typing/deleting fluido.
  const upd = useCallback((field, val) => update(g.id, field, val), [update, g.id])
  const onDelete = useCallback(() => remove(g.id), [remove, g.id])

  // Display ↔ storage: app-wide formatDateDisplay/parseDateDisplay handle
  // the DD-MM-YYYY ↔ YYYY-MM-DD round-trip.
  const dateDisplay  = formatDateDisplay(g.fechaFac)
  const onDateChange = v => upd('fechaFac', parseDateDisplay(v))

  // Divisa del renglón (monedaCodigo es el campo canónico; `moneda` queda por
  // compatibilidad con gastos viejos). Ticket extranjero SIN lado MXN todavía
  // (sin conciliar): la propina real vive en propinaExtranjero y el total real
  // en montoExtranjero — antes esas celdas mostraban $0 en la tabla aunque el
  // Excel sí traía los montos correctos.
  const monedaRow = normalizaMoneda(g.monedaCodigo || g.moneda || 'MXN')
  const propinaNativa = monedaRow !== 'MXN' && !(Number(g.montoPropina) > 0)
  const totalFinalMXN = (Number(g.totalCFDI) || 0) + (Number(g.montoPropina) || 0)
  const totalNativo   = (Number(g.montoExtranjero) || 0) + (Number(g.propinaExtranjero) || 0)
  const muestraNativo = monedaRow !== 'MXN' && totalFinalMXN === 0 && totalNativo > 0

  // Per-row toggle for the Fecha Cobro cell: span (DD-MM-YYYY) when blurred,
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

      {/* Fecha Cobro — span shows DD-MM-YYYY; click swaps to date picker
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
        {/* Badge con la divisa REAL del ticket (antes decía siempre "USD",
            aunque el gasto fuera en euros o ringgit). Fila manual con monto
            tecleado pero sin divisa detectada (moneda=MXN): conserva la
            suposición histórica de USD — mismo comportamiento que v8.92. */}
        {g.montoUSD > 0 && (
          <span className="badge-usd">{monedaRow === 'MXN' ? 'USD' : monedaRow}</span>
        )}
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
      <td><NumCell g={g} upd={upd} field="importe"     prefix="$" /></td>

      {/* IVA */}
      <td><NumCell g={g} upd={upd} field="iva"         prefix="$" /></td>

      {/* ISR trasladado */}
      <td><NumCell g={g} upd={upd} field="isrTrasladado" prefix="$" /></td>

      {/* Ret. ISR */}
      <td><NumCell g={g} upd={upd} field="retencionISR" prefix="$" /></td>

      {/* Ret. IVA */}
      <td><NumCell g={g} upd={upd} field="retencionIVA" prefix="$" /></td>

      {/* Retenciones (total) */}
      <td><NumCell g={g} upd={upd} field="retenciones" prefix="$" /></td>

      {/* Total Factura */}
      <td><NumCell g={g} upd={upd} field="totalCFDI"   prefix="$" /></td>

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
          <option value="99">99 - Por Definir</option>
        </select>
      </td>

      {/* Banco — solo visible para colaboradores especiales. Dropdown con
          las cuentas conocidas. Permite ajustar manualmente el banco
          después del cotejo con Saldos. */}
      {isSpecial && (
        <td>
          <select
            className="cell-select"
            value={g.banco || ''}
            onChange={e => upd('banco', e.target.value)}
          >
            <option value="">—</option>
            {BANCOS_DISPONIBLES.map(b => (
              <option key={b} value={b}>{b}</option>
            ))}
          </select>
        </td>
      )}

      {/* Propina % — compact: rounds to 2 decimals, drops trailing zeros for typing comfort */}
      <td><NumCell g={g} upd={upd} field="propinaPorcentaje" suffix="%" compact /></td>

      {/* Propina $ — en ticket extranjero sin conciliar se captura/muestra en
          la divisa del ticket (propinaExtranjero); ya conciliado (o en MXN),
          en pesos como siempre. key fuerza remount al cambiar el binding para
          que el draft interno del NumCell no arrastre el valor del otro campo. */}
      <td>
        {propinaNativa
          ? <NumCell key="prop-ext" g={g} upd={upd} field="propinaExtranjero" prefix={simboloMoneda(monedaRow) || monedaRow + ' '} compact />
          : <NumCell key="prop-mxn" g={g} upd={upd} field="montoPropina" prefix="$" compact />}
      </td>

      {/* Total Final — si el lado MXN sigue en 0 (ticket extranjero sin
          conciliar) muestra el total NATIVO con su símbolo en vez de un
          "$0.00" engañoso; el equivalente en pesos llega al Validar Banco. */}
      <td>
        <span
          className={`total-val${g.hizoMatch ? ' is-blue' : ''}`}
          title={muestraNativo ? `Total en ${monedaRow}. El equivalente en pesos se llena al validar con el banco.` : undefined}
        >
          {muestraNativo
            ? `${simboloMoneda(monedaRow) || ''}${totalNativo.toFixed(2)} ${monedaRow}`
            : `$${totalFinalMXN.toFixed(2)}`}
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
            💵 +{simboloMoneda(monedaRow) || ''}{g.propinaExtranjero.toFixed(2)} propina
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
})

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
  const sinCargoRows = data.facturasSinCargo || []
  const docsSubidos = data.docsSubidos || 0
  const cBanco    = useCountUp(data.bancoRows)
  const cMatches  = useCountUp(data.matches)
  const cSin      = useCountUp(data.sinFactura.length)
  const cSinCargo = useCountUp(sinCargoRows.length)
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
    const code = (currency || 'MXN').toUpperCase()
    // US$ para dólares (no confundirlo con pesos) y el símbolo nativo para el
    // resto: antes CUALQUIER divisa que no fuera USD se pintaba con "$", así
    // que un cargo en yuanes salía como "$244.11 CNY".
    const pre = code === 'USD' ? 'US$' : code === 'MXN' ? '$' : (simboloMoneda(code) || '')
    return `${pre}${num.toLocaleString('es-MX', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`
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
  const sinCargoPassesFilter = (f) => {
    if (filterKey === 'usd') return (f.moneda || 'MXN') === 'USD'
    if (filterKey === 'mxn') return (f.moneda || 'MXN') === 'MXN'
    if (filterKey === 'tickets') return f.isTicket
    if (filterKey === 'facturas') return !f.isTicket
    return true
  }
  // Clara llena "Monto original" incluso en cargos en pesos, así que un match
  // MXN traía csvAmountUSD == csvAmountMXN y la tarjeta mostraba el mismo
  // dinero dos veces (en pesos y en "USD"). Solo es divisa si la fila lo dice.
  const tieneDivisa = (m) => (m.csvMoneda || 'MXN') !== 'MXN' && m.csvAmountUSD > 0

  const matchesQuery = (txt) => {
    if (!query.trim()) return true
    const q = query.trim().toLowerCase()
    return (txt || '').toLowerCase().includes(q)
  }

  const visibleMatches = matchedRows
    .filter(m => matchesPassesFilter(m) && (matchesQuery(m.invoiceName) || matchesQuery(m.csvDescripcion)))
    .sort((a, b) => b.confidence - a.confidence)
  const visibleSin      = data.sinFactura.filter(s => sinFactPassesFilter(s) && matchesQuery(s.descripcion))
  const visibleSinCargo = sinCargoRows
    .filter(f => sinCargoPassesFilter(f) && (matchesQuery(f.proveedor) || matchesQuery(f.noFactura)))
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
              <h2 className="cm-fs-title">Listo — así quedó tu conciliación</h2>
              <p className="cm-fs-subtitle">
                Comparamos los <b>{data.bancoRows}</b> {data.bancoRows === 1 ? 'cargo' : 'cargos'} de tu estado de
                cuenta de Clara contra los <b>{docsSubidos}</b>{' '}
                {docsSubidos === 1 ? 'comprobante que subiste' : 'comprobantes que subiste'}.
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

          <p className="cm-fs-progress-caption">
            {Math.round(cPct)}% de los cargos ya tienen su comprobante
          </p>

          {/* Aviso anti-pánico: la gente leía "sin factura / pendiente" como
              un adeudo. Este banner aclara, antes que cualquier número, que
              esto es un checklist de documentos, no un estado de cuenta. */}
          <div className="cm-fs-explainer">
            <Info size={15} className="cm-fs-explainer-icon" />
            <div>
              <b>Esto no es un adeudo.</b> Estos cargos ya se pagaron con la tarjeta Clara. Aquí solo
              ves <b>cuáles ya tienen su factura o ticket</b> y cuáles todavía no.
            </div>
          </div>

          <div className="cm-fs-chips">
            <div className="cm-fs-chip cm-chip-blue">
              <CreditCard size={14} />
              <div>
                <div className="cm-fs-chip-value">{Math.round(cBanco)}</div>
                <div className="cm-fs-chip-label">Cargos en Clara</div>
                <div className="cm-fs-chip-hint">líneas del estado de cuenta</div>
              </div>
            </div>
            <div className="cm-fs-chip cm-chip-green">
              <Target size={14} />
              <div>
                <div className="cm-fs-chip-value">{Math.round(cMatches)}</div>
                <div className="cm-fs-chip-label">Ya tienen comprobante</div>
                <div className="cm-fs-chip-hint">cargo vinculado a tu factura</div>
              </div>
            </div>
            <div className="cm-fs-chip cm-chip-amber">
              <FileWarning size={14} />
              <div>
                <div className="cm-fs-chip-value">{Math.round(cSin)}</div>
                <div className="cm-fs-chip-label">Les falta comprobante</div>
                <div className="cm-fs-chip-hint">sube su factura o ticket</div>
              </div>
            </div>
            <div className="cm-fs-chip cm-chip-slate">
              <FileText size={14} />
              <div>
                <div className="cm-fs-chip-value">{Math.round(cSinCargo)}</div>
                <div className="cm-fs-chip-label">No están en el estado</div>
                <div className="cm-fs-chip-hint">comprobantes de otro periodo o pago</div>
              </div>
            </div>
            <div className="cm-fs-chip cm-chip-purple">
              <Sparkles size={14} />
              <div>
                <div className="cm-fs-chip-value">{Math.round(cPropinas)}</div>
                <div className="cm-fs-chip-label">Propinas estimadas</div>
                <div className="cm-fs-chip-hint">el banco cobró más que la factura</div>
              </div>
            </div>
          </div>

          {(data.totalsMatched || data.totalsPending) && (
            <div className="cm-fs-totals">
              <div className="cm-fs-totals-block">
                <span className="cm-fs-totals-label">Cargos ya comprobados</span>
                <span className="cm-fs-totals-value cm-fs-totals-green">
                  {fmtMoney(data.totalsMatched?.mxn)} MXN
                  {data.totalsMatched?.usd > 0 && <> · {fmtMoney(data.totalsMatched.usd, 'USD')} USD</>}
                </span>
              </div>
              <div className="cm-fs-totals-block">
                <span className="cm-fs-totals-label">Cargos por comprobar</span>
                <span className="cm-fs-totals-value cm-fs-totals-amber">
                  {fmtMoney(data.totalsPending?.mxn)} MXN
                  {data.totalsPending?.usd > 0 && <> · {fmtMoney(data.totalsPending.usd, 'USD')} USD</>}
                </span>
                <span className="cm-fs-totals-note">no es dinero que debas, es papeleo que falta</span>
              </div>
            </div>
          )}
        </header>

        {/* ───── TABS ───── */}
        <div className="cm-fs-tabs" role="tablist">
          {[
            { key: 'matches',  label: 'Conciliados',            icon: <Check size={14} />,       count: visibleMatches.length,  badgeClass: 'cm-fs-badge-green' },
            { key: 'sin',      label: 'Cargos sin comprobante', icon: <FileWarning size={14} />, count: visibleSin.length,      badgeClass: 'cm-fs-badge-amber' },
            { key: 'sincargo', label: 'No están en el estado',  icon: <FileText size={14} />,    count: visibleSinCargo.length, badgeClass: 'cm-fs-badge-slate' },
            { key: 'revision', label: 'Revisar',                icon: <AlertCircle size={14} />, count: visibleRevision.length, badgeClass: 'cm-fs-badge-yellow' },
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
                  <div className="cm-fs-pane-note cm-note-green">
                    Cada tarjeta es un <b>cargo de tu estado de cuenta</b> (izquierda) que quedó vinculado
                    con <b>el comprobante que subiste</b> (derecha). Haz clic en una para ver cómo cuadró.
                  </div>
                  {visibleMatches.length === 0 ? (
                    <div className="cm-fs-empty">Ningún cargo conciliado cumple los filtros.</div>
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
                              <div className="cm-fs-match-sub">
                                {m.propina > 0
                                  ? <span className="cm-fs-tip-pill"><Sparkles size={10} /> Propina estimada {fmtMoney(m.propina, m.invoiceMoneda)}</span>
                                  : m.method}
                              </div>
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
                              {tieneDivisa(m) && <span>{fmtMoney(m.csvAmountUSD, 'USD')} {m.csvMoneda}</span>}
                            </div>
                            <ChevronDown size={14} className={`cm-fs-chev ${isOpen ? 'is-open' : ''}`} />
                          </div>
                        </div>
                        <div className="cm-fs-match-detail" style={{ maxHeight: isOpen ? 260 : 0 }}>
                          {/* La aritmética explícita evita el "¿por qué el banco
                              me cobró más que la factura?" que dispara la duda. */}
                          <div className="cm-fs-math">
                            {m.propina > 0 ? (
                              <>Factura {fmtMoney(m.invoiceTotal, m.invoiceMoneda)} + propina{' '}
                                {fmtMoney(m.propina, m.invoiceMoneda)}
                                {m.propinaPct > 0 && <> ({m.propinaPct.toFixed(1)}%)</>} ={' '}
                                <b>{fmtMoney(tieneDivisa(m) && m.invoiceMoneda !== 'MXN' ? m.csvAmountUSD : m.csvAmountMXN, m.invoiceMoneda)} cobrados por Clara</b>
                              </>
                            ) : (
                              <>Factura {fmtMoney(m.invoiceTotal, m.invoiceMoneda)} ={' '}
                                <b>{fmtMoney(tieneDivisa(m) && m.invoiceMoneda !== 'MXN' ? m.csvAmountUSD : m.csvAmountMXN, m.invoiceMoneda)} cobrados por Clara</b>
                              </>
                            )}
                          </div>
                          <div className="cm-fs-detail-grid">
                            <div><span>Cómo se vinculó</span><strong>{m.method}</strong></div>
                            <div><span>Factura</span><strong>{m.invoiceNumber || '—'}</strong></div>
                            <div><span>Total factura</span><strong>{fmtMoney(m.invoiceTotal, m.invoiceMoneda)}</strong></div>
                            {m.propina > 0 && <div><span>Propina estimada</span><strong>{fmtMoney(m.propina, m.invoiceMoneda)}</strong></div>}
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
                  <div className="cm-fs-pane-note cm-note-amber">
                    Estos cargos <b>sí aparecen en tu estado de cuenta de Clara</b>, pero todavía no subiste
                    su factura o ticket. <b>No debes ese dinero</b> — ya se pagó con la tarjeta; lo único
                    pendiente es el comprobante para el reporte.
                  </div>
                  {visibleSin.length === 0 ? (
                    <div className="cm-fs-empty">Todos los cargos tienen su comprobante. 🎉</div>
                  ) : visibleSin.map((s, i) => (
                    <div key={i} className="cm-fs-card cm-fs-sin-card">
                      <div className="cm-fs-sin-head">
                        <div className="cm-fs-sin-left">
                          <span className="cm-fs-badge cm-fs-badge-amber">Falta comprobante</span>
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
                          <span className="cm-fs-suggest-label">¿Será alguno de estos que ya subiste?</span>
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

              {tab === 'sincargo' && (
                <motion.div
                  key="sincargo"
                  className="cm-fs-pane"
                  initial={{ opacity: 0, x: 16 }}
                  animate={{ opacity: 1, x: 0 }}
                  exit={{ opacity: 0, x: -16 }}
                  transition={{ duration: 0.25, ease: 'easeOut' }}
                >
                  <div className="cm-fs-pane-note cm-note-slate">
                    Estos comprobantes <b>sí los subiste</b>, pero no encontramos su cargo en este estado de
                    cuenta. Es normal: suele ser de <b>otro periodo</b>, pagado con <b>otra tarjeta</b>,
                    en <b>efectivo</b> o por <b>transferencia</b>. Se quedan en tu reporte tal cual.
                  </div>
                  {visibleSinCargo.length === 0 ? (
                    <div className="cm-fs-empty">Todos tus comprobantes encontraron su cargo en Clara. 🎉</div>
                  ) : visibleSinCargo.map((f, i) => (
                    <div key={f.id || i} className="cm-fs-card cm-fs-nocharge-card">
                      <div className="cm-fs-sin-head">
                        <div className="cm-fs-sin-left">
                          <span className="cm-fs-badge cm-fs-badge-slate">
                            {f.isTicket ? 'Ticket' : 'Factura'}
                          </span>
                          <div>
                            <div className="cm-fs-match-name">{f.proveedor}</div>
                            <div className="cm-fs-match-sub">
                              {formatDateDisplay(f.fecha)}{f.noFactura ? ` · ${f.noFactura}` : ''}
                            </div>
                          </div>
                        </div>
                        <div className="cm-fs-sin-right">
                          <div className="cm-fs-match-amount">
                            <span>{fmtMoney(f.total, f.moneda)} {f.moneda}</span>
                          </div>
                          <span className="cm-fs-nocharge-hint">no aparece en este estado de cuenta</span>
                        </div>
                      </div>
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
                  <div className="cm-fs-pane-note cm-note-yellow">
                    Estos cargos <b>sí quedaron vinculados</b>, pero el monto o la propina no cuadran al
                    centavo. Échales un ojo antes de continuar: si el proveedor y la fecha son correctos,
                    puedes dejarlos así.
                  </div>
                  {visibleRevision.length === 0 ? (
                    <div className="cm-fs-empty">Nada que revisar: todo cuadró con alta confianza. 🎉</div>
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
                            {tieneDivisa(m) && <span>{fmtMoney(m.csvAmountUSD, 'USD')} {m.csvMoneda}</span>}
                          </div>
                        </div>
                      </div>
                      <div className="cm-fs-rev-note">
                        {m.descuadre
                          ? 'La factura más la propina no suman exactamente lo que cobró Clara. Revisa el monto antes de continuar.'
                          : 'Coincidencia aproximada. Verifica proveedor, fecha y monto antes de continuar.'}
                      </div>
                    </div>
                  ))}
                </motion.div>
              )}
            </AnimatePresence>
          )}
        </div>

        {/* ───── LEYENDA + STICKY FOOTER ───── */}
        <div className="cm-fs-legend">
          <span className="cm-fs-legend-title">Cómo leer esto:</span>
          <span className="cm-fs-legend-item"><i className="cm-dot cm-dot-green" />Conciliado — el cargo y tu comprobante coinciden</span>
          <span className="cm-fs-legend-item"><i className="cm-dot cm-dot-amber" />Falta comprobante — el cargo ya se pagó, sube su factura</span>
          <span className="cm-fs-legend-item"><i className="cm-dot cm-dot-slate" />No está en el estado — comprobante de otro periodo o pago</span>
          <span className="cm-fs-legend-item"><i className="cm-dot cm-dot-yellow" />Revisar — cuadró, pero conviene verificarlo</span>
        </div>
        <footer className="cm-fs-footer">
          <span className="cm-fs-footer-hint">
            {matchCount} {matchCount === 1 ? 'cargo quedará vinculado' : 'cargos quedarán vinculados'} en tu reporte
          </span>
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
            Continuar
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
function PremiumModal({ open, type, title, subtitle, body, stats, primaryLabel, secondaryLabel, onPrimary, onSecondary, children }) {
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
        {body && (
          <pre style={{
            textAlign: 'left', background: 'rgba(255,255,255,0.05)',
            borderRadius: 8, padding: '10px 14px', fontSize: 12,
            color: 'rgba(255,255,255,0.75)', whiteSpace: 'pre-wrap',
            wordBreak: 'break-word', margin: '8px 0 0', maxHeight: 160,
            overflowY: 'auto', fontFamily: 'inherit',
          }}>{body}</pre>
        )}

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
  // Progreso del OCR por lote: null = no hay lote; {current, total} = "N de M".
  // Lo lee el overlay para que el usuario vea avance en vez de un spinner mudo
  // durante los minutos que tarda un lote grande (límite TPM de Groq gratuito).
  const [ocrProgress,   setOcrProgress]   = useState(null)
  // Small chooser shown by the "Cargar Foto" button: on móvil el input con
  // capture="environment" abre la cámara directo, así que ofrecemos elegir
  // entre tomar foto (cámara) o subir una existente (galería/archivos).
  const [showPhotoChoice, setShowPhotoChoice] = useState(false)
  // Recorrido guiado "Cómo utilizar": null = apagado, 0..n = paso activo.
  const [tourStep, setTourStep] = useState(null)
  // Guía "Cómo descargar tu estado de cuenta de Clara" (modal-carrusel).
  const [showClaraGuide, setShowClaraGuide] = useState(false)
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

  // Auto-fit del ancho de la columna FACTURA (colWidths[4]) al folio más
  // largo presente en la lista, para que folios como "F103069-141757" o
  // "FAC102026491" nunca se corten con elipsis. Se recalcula cuando cambia
  // la lista; respeta un mínimo de 120px.
  useEffect(() => {
    if (lista.length === 0) return
    const longest = lista.reduce((max, g) => {
      const len = String(g.noFactura || '').length
      return len > max ? len : max
    }, 0)
    // ~8.5px por char en Inter 13px + ~30px de padding/borders
    const fitted = Math.max(120, Math.ceil(longest * 8.5) + 30)
    setColWidths(prev => {
      if (prev[4] === fitted) return prev
      const next = [...prev]
      next[4] = fitted
      return next
    })
  }, [lista])

  // Cmd/Ctrl+K → enfoca el buscador universal. Escape → cierra/clear.
  useEffect(() => {
    const onKey = (e) => {
      if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === 'k') {
        e.preventDefault()
        searchInputRef.current?.focus()
        searchInputRef.current?.select()
      } else if (e.key === 'Escape' && document.activeElement === searchInputRef.current) {
        setBusqueda('')
        searchInputRef.current?.blur()
      }
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [])
  // Index-based fixed pixel widths — order matches COLUMNS positions:
  // [0] checkbox, [1] estado, [2] fecha factura, [3] fecha cobro,
  // [4] factura, [5] proveedor, [6] concepto, [7] tipo, [8] subtotal,
  // [9] iva, [10] isr/ish/ieps, [11] ret.isr, [12] ret.iva, [13] reten,
  // [14] total fac, [15] forma pago, [16] prop%, [17] prop$, [18] total final
  // 22 columnas: ...formaPago(15), banco(16, solo especiales), prop%(17),
  // prop$(18), totalFinal(19), montoUSD(20), tipoCambio(21)
  const [colWidths, setColWidths] = useState([40, 110, 115, 120, 120, 260, 140, 100, 120, 110, 135, 110, 110, 110, 125, 160, 195, 95, 105, 130, 110, 80])
  const [sort,          setSort]          = useState({ field: null, dir: 'asc' })
  const [busqueda,      setBusqueda]      = useState('')
  const searchInputRef = useRef(null)

  const folderRef = useRef(null)
  const bancoRef  = useRef(null)
  const photoRef  = useRef(null)       // subir foto existente (galería / archivos)
  const cameraRef = useRef(null)       // tomar foto (capture="environment")
  const saldosRef = useRef(null)
  const nominaRef = useRef(null)
  const [cotejoModal, setCotejoModal] = useState(null)
  const [duplicadosModal, setDuplicadosModal] = useState(null)
  const [ocrSelectionModal, setOcrSelectionModal] = useState(null)
  const [nominaPickerModal, setNominaPickerModal] = useState(null)

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

  // Aviso visible cuando hay PDFs muy grandes entre los que se mandarán al OCR.
  // Le explica al usuario que se comprimirán (rasterizar 1ª página a JPEG chico)
  // para procesarlos rápido, sin tocar el PDF original que se exporta. Devuelve
  // true si el usuario acepta continuar; false si cancela. Si no hay PDFs
  // grandes, no muestra nada y sigue de largo.
  const avisarPdfsGrandes = async (archivos) => {
    const grandes = (archivos || []).filter(f =>
      f && f.name && f.name.toLowerCase().endsWith('.pdf') && f.size > LARGE_PDF_BYTES
    )
    if (grandes.length === 0) return true
    const detalle = grandes
      .map(f => `• ${f.name}  (${(f.size / 1024 / 1024).toFixed(1)} MB)`)
      .join('\n')
    return await askConfirm({
      type: 'warning',
      title: grandes.length === 1 ? 'PDF demasiado grande' : `${grandes.length} PDFs demasiado grandes`,
      subtitle: `Para que el OCR sea rápido y no se trabe, ${grandes.length === 1 ? 'lo comprimiremos' : 'los comprimiremos'} automáticamente antes de procesar${grandes.length === 1 ? 'lo' : 'los'}. El archivo original NO se altera (se exporta igual en el ZIP).\n\n${detalle}`,
      stats: [
        { value: grandes.length, label: grandes.length === 1 ? 'PDF pesado' : 'PDFs pesados' },
        { value: 'Automático', label: 'Compresión' },
      ],
      primaryLabel: grandes.length === 1 ? 'Comprimir y continuar' : 'Comprimir todos y continuar',
      secondaryLabel: 'Cancelar',
    })
  }

  // ── Métricas (cards en el encabezado de la tabla) ──
  const metrics = useMemo(() => {
    const sum = field => lista.reduce((s, g) => s + (g[field] || 0), 0)
    // Totales POR DIVISA (incluyen la propina nativa, igual que la SUM(R) del
    // Excel). Antes era un solo `totalUSD` que sumaba todo monto extranjero
    // sin distinguir moneda: dos tickets en yuanes (¥244.11 + ¥5) se mostraban
    // como "TOTAL USD $249.11". Usa montoExtranjero con fallback a montoUSD
    // para mantener compat con rows viejos.
    const porMoneda = {}
    for (const g of lista) {
      const code = normalizaMoneda(g.monedaCodigo || g.moneda || 'MXN')
      if (code === 'MXN') continue
      const monto = (Number(g.montoExtranjero) || Number(g.montoUSD) || 0)
        + (Number(g.propinaExtranjero) || 0)
      if (!monto) continue
      porMoneda[code] = +((porMoneda[code] || 0) + monto).toFixed(2)
    }
    return {
      totalFacturado:   sum('totalCFDI'),
      ivaTotal:         sum('iva'),
      retencionesTotal: sum('retenciones'),
      sinCobrar:        lista.filter(g => !g.fechaCobro).length,
      count:            lista.length,
      porMoneda,
    }
  }, [lista])
  const fmtMoney = n => `$${n.toLocaleString('es-MX', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`
  const fmtDivisa = (code, n) =>
    `${simboloMoneda(code) || code + ' '}${n.toLocaleString('es-MX', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`

  // ── Búsqueda universal — substring (case + accent-insensitive) sobre
  //    todos los campos significativos del gasto. Se aplica ANTES del sort
  //    para que el orden actuál actúe sobre el subset filtrado.
  const filteredLista = useMemo(() => {
    const q = busqueda.trim()
    if (!q) return lista
    const nq = q.normalize('NFD').replace(/\p{Diacritic}/gu, '').toLowerCase()
    return lista.filter(g => {
      const haystack = [
        g.rfc, g.proveedor, g.noFactura, g.concepto, g.tipo,
        g.fechaFac, g.fechaCobro, g.formaPago, g.polizaNumero,
        g.importe, g.iva, g.retenciones, g.totalCFDI,
        g.montoUSD, g.tipoCambio, g.moneda, g.monedaCodigo,
        g.uuid,
      ]
        .filter(v => v !== null && v !== undefined && v !== '')
        .join(' ')
        .normalize('NFD')
        .replace(/\p{Diacritic}/gu, '')
        .toLowerCase()
      return haystack.includes(nq)
    })
  }, [lista, busqueda])

  // ── Sort — three-state cycle (asc → desc → unsorted), null-safe, with
  //          date detection that accepts both YYYY-MM-DD and DD/MM/YYYY.
  const sortedLista = useMemo(() => {
    if (!sort.field) return filteredLista
    const col = COLUMNS.find(c => c.key === sort.field)
    if (!col) return filteredLista
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
    return [...filteredLista].sort((a, b) => {
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
  }, [filteredLista, sort])

  // TIPO dropdown list — depends on the selected colaborador's categoría.
  // Ventas/Socio get the sales-flavored list; Admin/Servicio get the
  // operational list; null (modal still open) falls back to a merged
  // sorted union so existing rows can still render their saved tipo.
  const tiposList = useMemo(() => getTiposForColaborador(colaborador), [colaborador])
  const esColaboradorEspecial = COLABORADORES_ESPECIALES.includes(colaborador?.nombre)

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
  // useCallback con deps vacías → identidad estable entre renders. Sólo usa
  // setLista (con updater funcional, sin cerrar sobre `lista`), así que no hay
  // closure vieja. Estable permite que React.memo en GastoRow evite re-render
  // de las demás filas al teclear.
  const update = useCallback((id, field, value) =>
    setLista(prev => prev.map(g => {
      if (g.id !== id) return g
      const u = { ...g, [field]: value }

      // La celda "ISH/IEPS" edita isrTrasladado, pero el Excel lee ishIeps
      // (con fallback a isrTrasladado). En filas importadas de un Excel viejo
      // ishIeps ya venía poblado, así que editar la celda NO viajaba al
      // reporte. Se espejan siempre.
      if (field === 'isrTrasladado') u.ishIeps = value

      // Las retenciones parciales mantienen el total de retenciones. Van ANTES
      // del recálculo de TOTAL porque éste las lee.
      if (field === 'retencionISR') u.retenciones = value + (u.retencionIVA || 0)
      if (field === 'retencionIVA') u.retenciones = value + (u.retencionISR || 0)

      // TOTAL en vivo — espejo de la fórmula del Excel (=J+K−L+M−N). Antes la
      // app NO recalculaba totalCFDI al editar Subtotal/IVA/retenciones: la
      // fila manual mostraba $0.00 (y Total Final mal) aunque el Excel
      // exportado sí sacara el total correcto con su fórmula.
      if (RECALC_TOTAL_FIELDS.has(field)) {
        const ish = Number(u.ishIeps ?? u.isrTrasladado) || 0
        const retNoIsr = (Number(u.retenciones) || 0) - (Number(u.retencionISR) || 0)
        u.totalCFDI = Math.round(((Number(u.importe) || 0) + (Number(u.iva) || 0)
          - (Number(u.retencionISR) || 0) + ish - retNoIsr) * 100) / 100
      }

      // Propina encadenada al total (editado directo o recalculado arriba).
      const totalBase = field === 'totalCFDI' ? value : u.totalCFDI
      if ((field === 'totalCFDI' || RECALC_TOTAL_FIELDS.has(field)) && u.propinaPorcentaje > 0)
        u.montoPropina = Math.round(totalBase * u.propinaPorcentaje / 100 * 100) / 100
      if (field === 'propinaPorcentaje') {
        u.montoPropina = Math.round(g.totalCFDI * value / 100 * 100) / 100
        // Fila extranjera: el % también aplica sobre el monto nativo — antes
        // calculaba contra totalCFDI=0 (sin conciliar) y siempre daba $0.
        if ((Number(g.montoExtranjero) || 0) > 0)
          u.propinaExtranjero = Math.round(g.montoExtranjero * value / 100 * 100) / 100
      }
      if (field === 'montoPropina' && g.totalCFDI > 0)
        u.propinaPorcentaje = Math.round((value / g.totalCFDI) * 10000) / 100
      if (field === 'propinaExtranjero' && (Number(g.montoExtranjero) || 0) > 0)
        u.propinaPorcentaje = Math.round((value / g.montoExtranjero) * 10000) / 100
      return u
    })), [])

  // Eliminar un gasto por id — estable para no romper la memoización de filas.
  const removeGasto = useCallback(id =>
    setLista(prev => prev.filter(x => x.id !== id)), [])

  // ── Cargar carpeta (XMLs + PDFs) ──
  // Shared file-processing pipeline — used by folder picker and drag/drop.
  const processFiles = async (files, folderName) => {
    if (!files.length) return
    const xmls = files.filter(f => f.name.toLowerCase().endsWith('.xml'))
    const pdfs = files.filter(f => f.name.toLowerCase().endsWith('.pdf'))
    // Normaliza nombres (NFD + lower + sin acentos) para comparaciones de
    // nombre base; usado por detectOcrCandidates más abajo.
    const normName = s => (s || '').normalize('NFD').replace(/\p{Diacritic}/gu, '').toLowerCase()
    setCarpetaNombre(folderName)
    setLoading(true)
    // Detectar XMLs corruptos (0 KB) antes de procesar
    const xmlsCorruptos = xmls.filter(f => f.size === 0)
    const xmlsValidos   = xmls.filter(f => f.size > 0)
    // Nombres base (sin .xml, normalizados) de los XML corruptos. El PDF que
    // comparta nombre base con uno de estos NO tiene XML usable → se manda a
    // OCR para no perder la info, y el gasto resultante se marca xmlFaltante.
    const corruptBaseNames = new Set(
      xmlsCorruptos.map(f => normName(f.name.replace(/\.xml$/i, '')))
    )
    if (xmlsCorruptos.length > 0) {
      await askConfirm({
        type: 'warning',
        title: 'XMLs corruptos detectados',
        subtitle: `${xmlsCorruptos.length} archivo${xmlsCorruptos.length > 1 ? 's' : ''} XML ${xmlsCorruptos.length > 1 ? 'están vacíos' : 'está vacío'} (0 KB) y no ${xmlsCorruptos.length > 1 ? 'pudieron' : 'pudo'} procesarse.`,
        body: xmlsCorruptos.map(f => `• ${f.name}`).join('\n') +
          '\n\nDescarga nuevamente el XML y su PDF con el mismo nombre desde el portal del SAT o del proveedor y vuelve a subirlos.',
        primaryLabel: 'Entendido',
      })
    }

    const nueva = []
    for (const f of xmlsValidos) {
      try {
        const text = await f.text()
        const g = parseCFDI(text, f, pdfs, colaborador)
        if (g) nueva.push(g)
      } catch {}
    }
    // Segunda pasada de matching XML↔PDF: parseCFDI solo enlaza por nombre
    // base exacto o UUID. Aquí completamos los PDFs renombrados por el usuario
    // (folio/RFC/substring en el nombre) de forma EXCLUSIVA: cada PDF se asigna
    // a un solo gasto, con prioridad por confianza. Misma lógica que handleDrop.
    const huerfanosNombre = linkPdfsExclusive(nueva, pdfs)
    // NIVEL 6: para los gastos que siguen sin PDF, lee el TEXTO embebido de los
    // PDFs huérfanos (sin red) y enlaza por UUID/folio/RFC encontrado DENTRO del
    // contenido. Cierra el caso del PDF renombrado a algo arbitrario.
    await linkPdfsByContent(nueva, huerfanosNombre)

    // Persist PDF bytes as data URLs so the ZIP export survives even if
    // the original File reference goes stale. XMLs already carry their
    // text via gasto.xmlContent from parseCFDI.
    for (const g of nueva) {
      if (g.pdfFile && !g.pdfDataURL) {
        try { g.pdfDataURL = await fileToDataURL(g.pdfFile) }
        catch (err) { console.warn('PDF read failed:', g.pdfFile.name, err) }
      }
    }

    // FNI (peaje, RFC FNI970829JR9): su CFDI XML NO trae Serie/Folio — el
    // número de factura (p.ej. FNPE72984235) solo aparece impreso en el PDF.
    // Para las facturas FNI que entraron por XML con PDF adjunto, leemos el PDF
    // por OCR y sustituimos el folio (que de otro modo cae al NoIdentificacion)
    // por el Serie+Folio real. Solo para FNI y solo si aún no lo tenemos.
    await enrichFniFoliosDesdePDF(nueva)
    await enrichTelcelParcialidad(nueva)

    const applyBatch = (batch, ocrCount = 0) => {
      // Merge ACUMULATIVO: en vez de sobrescribir, agrega la carpeta nueva a los
      // registros existentes. Dedup por rfc|noFactura — una factura repetida
      // actualiza la fila previa en vez de duplicarse. Así el usuario puede
      // cargar varias carpetas seguidas y todas se acumulan en la lista.
      let added = 0, updated = 0
      setLista(prev => {
        const merged = [...prev]
        const keys = new Set(prev.map(g => gastoDedupKey(g)))
        for (const newG of batch) {
          const key = gastoDedupKey(newG)
          if (keys.has(key)) {
            const idx = merged.findIndex(g => gastoDedupKey(g) === key)
            if (idx !== -1) {
              // Si el entrante trae XML real (xmlContent), el usuario re-subió
              // el XML correcto → limpia la marca xmlFaltante.
              const clearFlag = !!newG.xmlContent
              merged[idx] = {
                ...merged[idx], ...newG, isNew: true,
                xmlFaltante: clearFlag ? false : (newG.xmlFaltante || merged[idx].xmlFaltante),
              }
              updated++
            }
          } else {
            merged.push(newG); keys.add(key); added++
          }
        }
        return merged
      })
      setTimeout(() => {
        setLista(l => l.map(g => g.isNew ? { ...g, isNew: false } : g))
      }, 1500)
      setLoading(false)
      if (batch.length > 0) {
        const linkedPDFs = batch.filter(g => g.tienePDF).length
        showModal({
          type: 'success',
          title: 'Carpeta cargada',
          subtitle: `Procesamos ${xmls.length} XML${xmls.length === 1 ? '' : 's'}${ocrCount > 0 ? ` + ${ocrCount} OCR` : ''} correctamente.`,
          stats: [
            { value: `+${added}`,                       label: 'Facturas nuevas', color: '#59D39B' },
            ...(updated > 0 ? [{ value: updated,        label: 'Actualizadas',    color: 'rgba(255,255,255,0.85)' }] : []),
            { value: xmls.length,                       label: 'XMLs leídos',     color: 'rgba(255,255,255,0.85)' },
            ...(linkedPDFs > 0 ? [{ value: linkedPDFs, label: 'PDFs vinculados', color: 'rgba(255,255,255,0.85)' }] : []),
            ...(ocrCount > 0  ? [{ value: ocrCount, label: `OCR (≈ ${fmtUsd(ocrCount * OCR_COST_USD)})`, color: '#f59e0b' }] : []),
          ],
          primaryLabel: 'Continuar',
        })
        setCarpetaSuccess(true)
        setTimeout(() => setCarpetaSuccess(false), 2500)
      }
    }

    // Helper: detecta orphan PDFs (sin XML pareja) + imágenes en la carpeta.
    // Son candidatos a OCR; el usuario decide uno por uno cuáles procesar.
    const detectOcrCandidates = (xmlGastos) => {
      // Fuente 1: gastos que parseCFDI logró linkear con su PDF
      const linkedPdfNames = new Set(
        xmlGastos.filter(g => g.pdfFile).map(g => g.pdfFile.name)
      )
      // Fuente 2 (defensa de raíz): la carpeta tiene un archivo .xml VÁLIDO con
      // el MISMO nombre base que el PDF. Aunque parseCFDI haya fallado al parsear
      // ese XML (namespace raro, sello inválido, RFC vacío, etc.) o aunque
      // la heurística de matching no haya tripeado, el PDF claramente tiene
      // XML pareja en el folder y NO debe mandarse a OCR. OJO: usamos solo los
      // XML válidos — un PDF cuyo único XML está corrupto (0 KB) SÍ debe ir a
      // OCR para rescatar la info, y se marcará xmlFaltante más abajo.
      const xmlBaseNames = new Set(
        xmlsValidos.map(x => normName(x.name.replace(/\.xml$/i, '')))
      )
      const orphanPDFs = pdfs.filter(p => {
        if (linkedPdfNames.has(p.name)) return false
        const pdfBase = normName(p.name.replace(/\.pdf$/i, ''))
        if (xmlBaseNames.has(pdfBase)) return false
        return true
      })
      const imageFiles = files.filter(f =>
        /\.(jpe?g|png|webp|heic|heif|bmp|gif)$/i.test(f.name)
      )
      return [...orphanPDFs, ...imageFiles]
    }

    // Helper: corre OCR sobre los archivos seleccionados, regresa los gastos
    // ya formateados (con pdfDataURL/imageDataURL para el export).
    const runOcrOnSelection = async (selectedFiles) => {
      if (!selectedFiles.length) return { results: [], failed: [] }
      setLoading(true)
      const results = []
      const failed = []
      for (let i = 0; i < selectedFiles.length; i++) {
        const file = selectedFiles[i]
        // Progreso "N de M" + pausa breve (salvo en la primera) para repartir las
        // peticiones y no chocar todas a la vez con el rate limit de Groq.
        setOcrProgress({ current: i + 1, total: selectedFiles.length })
        if (i > 0) await sleep(OCR_PACING_MS)
        // Reintento: el OCR (Groq, tier gratuito) puede fallar de forma
        // transitoria por límite de tasa cuando se mandan varias facturas
        // seguidas, o por un JSON mal formado puntual. Antes el archivo se
        // descartaba en silencio (solo console.warn) — por eso a veces de 3
        // facturas solo aparecían 2. Reintentamos una vez con backoff y, si
        // aun así falla, lo reportamos al usuario en vez de perderlo callado.
        let ok = false
        let lastErr = null
        for (let attempt = 0; attempt < 2 && !ok; attempt++) {
          if (attempt > 0) await new Promise(r => setTimeout(r, 1200))
          try {
            const isImage = (file.type || '').startsWith('image/') ||
              /\.(jpe?g|png|webp|heic|heif|bmp|gif)$/i.test(file.name)
            let fileForOCR = file
            let mediaType = isImage ? 'image/jpeg' : 'application/pdf'
            if (isImage) {
              fileForOCR = await compressImage(file)
            } else if (file.size > LARGE_PDF_BYTES) {
              // PDF grande (escaneo): rasteriza a JPEG chico en el navegador
              // para que la subida no se tarde. Los PDFs chicos (CFDI normales,
              // FNI) se mandan tal cual para no perder la extracción de texto.
              const jpg = await pdfFirstPageToJpeg(file)
              if (jpg) { fileForOCR = jpg; mediaType = 'image/jpeg' }
            }
            const base64 = await fileToBase64(fileForOCR)
            const g = await extractReceiptData(base64, mediaType, file.name)
            if (isImage) {
              try { g.imageDataURL = await fileToDataURL(fileForOCR) } catch {}
              g.originalFileName = file.name
            } else {
              g.pdfFile = file
              g.tienePDF = true
              try { g.pdfDataURL = await fileToDataURL(file) } catch {}
            }
            // Si este PDF/imagen es el rescate de un XML corrupto (mismo nombre
            // base que un .xml de 0 KB), marca el gasto para que el Excel señale
            // "Falta XML" en la fila. Se limpia si luego suben el XML correcto.
            const baseSinExt = normName(file.name.replace(/\.[^.]+$/i, ''))
            if (corruptBaseNames.has(baseSinExt)) g.xmlFaltante = true
            g.isNew = true
            results.push(g)
            ok = true
          } catch (err) {
            lastErr = err
          }
        }
        if (!ok) {
          console.warn('OCR failed for', file.name, lastErr)
          failed.push(file.name)
        }
      }
      setLoading(false)
      setOcrProgress(null)
      return { results, failed }
    }

    // Wraps applyBatch: si hay candidatos OCR abre el modal de selección;
    // si el usuario cancela u opta por no procesar nada, aplica sólo XMLs.
    const applyBatchWithOcrGate = (xmlBatch) => {
      const candidates = detectOcrCandidates(xmlBatch)
      if (candidates.length === 0) {
        applyBatch(xmlBatch)
        return
      }
      setLoading(false)
      setOcrSelectionModal({
        items: candidates.map(f => ({ file: f, selected: false })),
        onConfirm: async (selectedFiles) => {
          setOcrSelectionModal(null)
          // Si hay PDFs muy grandes, avisa que se comprimirán antes de procesar.
          const continuar = await avisarPdfsGrandes(selectedFiles)
          if (!continuar) { applyBatch(xmlBatch); return }
          const { results: ocrResults, failed } = await runOcrOnSelection(selectedFiles)
          applyBatch([...xmlBatch, ...ocrResults], ocrResults.length)
          if (failed.length > 0) {
            showModal({
              type: 'error',
              title: `${failed.length} factura${failed.length === 1 ? '' : 's'} no se pudo procesar`,
              subtitle: `El OCR falló en:\n\n${failed.join('\n')}\n\nVuelve a arrastrar solo ${failed.length === 1 ? 'ese archivo' : 'esos archivos'} para reintentar.`,
              primaryLabel: 'Entendido',
            })
          }
        },
        onCancel: () => {
          setOcrSelectionModal(null)
          applyBatch(xmlBatch)
        },
      })
    }

    // Detect duplicate RFC+noFactura keys within this batch
    const seenKeys = new Map()
    const batchDupes = []
    for (const g of nueva) {
      const key = `${g.rfc}|${g.noFactura}`
      if (seenKeys.has(key)) batchDupes.push(g)
      else seenKeys.set(key, true)
    }

    if (batchDupes.length > 0) {
      setLoading(false)
      setDuplicadosModal({
        items: batchDupes,
        incoming: nueva,
        onConfirm: (kept) => {
          setDuplicadosModal(null)
          applyBatchWithOcrGate(kept)
        },
      })
    } else {
      applyBatchWithOcrGate(nueva)
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

  // FNI (peaje): completa el Serie+Folio leyendo el PDF adjunto por OCR cuando
  // el CFDI XML no lo trae. Muta cada gasto FNI in-place. Acotado a FNI: en
  // otros emisores no aplica y evitamos llamadas OCR innecesarias.
  const enrichFniFoliosDesdePDF = async (gastos) => {
    // Serie+Folio de FNI: letras de serie + dígitos (p.ej. FNPE72984235). El
    // NoIdentificacion (p.ej. JKZY7DSB79M6XD9UHQ) mezcla letras y dígitos y NO
    // calza este patrón, así que no lo confundimos con un folio válido.
    const esSerieFolio = (s) => /^[A-Za-z]{2,6}\d{3,}$/.test((s || '').trim())
    const pendientes = gastos.filter(g =>
      g.rfc === 'FNI970829JR9' && !esSerieFolio(g.noFactura) && g.pdfFile
    )
    // En PARALELO y con folioOnly: el server lee el Serie+Folio del TEXTO del
    // PDF sin llamar a Groq (sin visión ni rate limit). Así una carpeta con
    // varias FNI ya no se traba — antes eran N llamadas OCR secuenciales, cada
    // una con su retry de rate limit (hasta 30s), = minutos de spinner.
    if (pendientes.length === 0) return
    const trabajo = Promise.all(pendientes.map(async (g) => {
      try {
        const base64 = await fileToBase64(g.pdfFile)
        const resp = await fetchConTimeout('/api/ocr-ticket', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ base64, mediaType: 'application/pdf', folioOnly: true }),
        }, 10_000)
        if (!resp.ok) return
        const j = await resp.json()
        if (j && esSerieFolio(j.folio)) g.noFactura = String(j.folio).trim()
      } catch (err) {
        console.warn('FNI folio falló:', g.pdfFile?.name, err)
      }
    }))
    // Tope DURO de tiempo: pase lo que pase con la red, seguimos en ≤12s para
    // que la carga NUNCA se quede trabada. Lo que no alcanzó a resolver queda
    // con su NoIdentificacion (el usuario lo puede editar).
    await Promise.race([trabajo, new Promise(r => setTimeout(r, 12_000))])
  }

  // Telcel (Radiomóvil Dipsa, RFC RDI841003QJ4): el CFDI cierra en "Cargos del
  // mes" pero el banco cobra el "Total a pagar", que suma una "Parcialidad de
  // Equipo" DESPUÉS del IVA — ese monto NO está en el XML, solo en el PDF. Lo
  // leemos del texto del PDF (telcelExtra, sin Groq) y lo sumamos al importe y
  // al total para que el gasto cuadre con el cargo del banco. Muta in-place.
  const enrichTelcelParcialidad = async (gastos) => {
    const pendientes = gastos.filter(g =>
      g.rfc === 'RDI841003QJ4' && g.pdfFile && !g.parcialidadEquipo
    )
    if (pendientes.length === 0) return
    const trabajo = Promise.all(pendientes.map(async (g) => {
      try {
        const base64 = await fileToBase64(g.pdfFile)
        const resp = await fetchConTimeout('/api/ocr-ticket', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ base64, mediaType: 'application/pdf', telcelExtra: true }),
        }, 10_000)
        if (!resp.ok) return
        const j = await resp.json()
        const p = Number(j?.parcialidad) || 0
        if (p > 0) {
          const r2 = (n) => parseFloat((Number(n) + p).toFixed(2))
          g.importe        = r2(g.importe)
          g.totalCFDI      = r2(g.totalCFDI)
          g.montoFacturado = r2(g.montoFacturado)
          g.parcialidadEquipo = p   // marca idempotente: no volver a sumar
        }
      } catch (err) {
        console.warn('Telcel parcialidad falló:', g.pdfFile?.name, err)
      }
    }))
    await Promise.race([trabajo, new Promise(r => setTimeout(r, 12_000))])
  }

  // Posts a base64 image/PDF to /api/ocr-ticket and shapes the parsed
  // receipt into a full gasto row. Reads `colaborador` from closure so
  // autoDetectTipo gets routed to the right category list. USD receipts
  // land montoUSD with importe/totalCFDI = 0 (the user fills the MXN side
  // from their card statement later).
  const extractReceiptData = async (base64, mediaType, fileName) => {
    const response = await fetchConTimeout('/api/ocr-ticket', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ base64, mediaType }),
    }, 90_000)
    if (!response.ok) {
      const err = await response.json().catch(() => ({}))
      throw new Error(err.error || `OCR ${response.status}`)
    }
    const parsed = await response.json()
    const uuid = crypto.randomUUID()
    const esPedimento  = (parsed.tipoDocumento || '').toString().toLowerCase() === 'pedimento'
    const monedaCode   = esPedimento ? 'MXN' : (parsed.moneda || 'MXN').toString().toUpperCase()
    const isExtranjera = !esPedimento && !!monedaCode && monedaCode !== 'MXN'
    const today = new Date().toISOString().slice(0, 10)
    let subtotal = Number(parsed.subtotal) || 0
    let iva = Number(parsed.iva) || 0
    let total = Number(parsed.total) || 0
    const propina = esPedimento ? 0 : Number(parsed.propina) || 0
    const propinaSugerida18 = esPedimento ? 0 : Number(parsed.propinaSugerida18) || 0
    const propinaSugerida20 = esPedimento ? 0 : Number(parsed.propinaSugerida20) || 0
    const propinaSugerida22 = esPedimento ? 0 : Number(parsed.propinaSugerida22) || 0

    // Pedimento branch: el importe pagado entra como total y como importe;
    // sin IVA desglosado, sin propina. proveedor/concepto fijos. esTicket=false
    // para que Pass 0 (auth code en CSV Clara) NO intente matchearlo.
    if (esPedimento) {
      const importePagado = total || subtotal
      return {
        id: genId(),
        // El pedimento se tramita a nombre de SMTO (el importador), así que el
        // RFC que va en el Excel es el propio de SMTO, no el de un proveedor.
        rfc: 'SEN1504132C9',
        proveedor: parsed.proveedor || 'SMTO ENGINEERING',
        noFactura: String(parsed.folio || ('PED-' + uuid.slice(0, 6).toUpperCase())),
        fechaFac:  forzarAñoEnCurso(parsed.fecha || today),
        concepto:  parsed.concepto || 'Pedimento de Importacion',
        tipo: 'Aduana',
        importe:    importePagado,
        iva:        0,
        isrTrasladado:  0,
        retencionISR:   0,
        retencionIVA:   0,
        retenciones:    0,
        totalCFDI:  importePagado,
        propinaPorcentaje: 0,
        montoPropina:      0,
        fechaCobro: forzarAñoEnCurso(parsed.fecha || today),
        formaPago: parsed.formaPago || '03',
        uuid,
        tienePDF: false,
        pdfFile: null,
        xmlFile: null,
        hizoMatch: false,
        validado: false,
        montoExtranjero:    0,
        propinaExtranjero:  0,
        monedaCodigo:       'MXN',
        esMonedaExtranjera: false,
        montoUSD:           0,
        tipoCambio:         0,
        moneda:             'MXN',
        esTicket:           false,
        esPedimento:        true,
        subtotal:           importePagado,
        propinaSugerida18:  0,
        propinaSugerida20:  0,
        propinaSugerida22:  0,
        banco:              defaultBancoFor(colaborador),
        montoFacturado:     importePagado,
        montoPropinaOriginal: 0,
      }
    }

    // ── Regla especial: recibos de ITESO (Universidad Jesuita de Guadalajara) ──
    // Estos recibos ponen al cliente (SMTO) en "Nombre:" y el monto real en la
    // LÍNEA del concepto, dejando el campo "Total:" en $0.00. Forzamos
    // proveedor/tipo de forma determinista y rescatamos el monto del subtotal
    // si el total vino en cero.
    const esITESO = /ITESO|universidad jesuita/i.test(parsed.proveedor || '') ||
                    /ADEUDO\s+DEP[OÓ]SITO\s+GARANT[IÍ]A/i.test(parsed.concepto || '')
    // Recibo oficial de ISN (Impuesto Sobre Nómina) de la Secretaría de la
    // Hacienda Pública (Jalisco). El recibo pone a SMTO en "Nombre", así que el
    // modelo suele devolver SMTO como proveedor; lo forzamos a Hacienda y el
    // tipo al existente '3% ISN'.
    const textoISN = `${parsed.proveedor || ''} ${parsed.concepto || ''}`
    const esISN = /IMPUESTO SOBRE N[OÓ]MINA|\(ISN\)|SECRETAR[IÍ]A DE (LA )?HACIENDA|HACIENDA P[UÚ]BLICA/i.test(textoISN)
    // Total Play: el estado de cuenta trae varios montos. El correcto (y el que
    // cobra el banco) es "Cargos del Mes"; las líneas "A PAGAR" traen redondeo
    // (precio de lista) o descuento (pronto pago) y NO cuadran con el banco.
    const esTotalPlay = /TOTAL\s*PLAY|TOTALPLAY/i.test(`${parsed.proveedor || ''} ${parsed.concepto || ''}`)
    // Microsoft (y facturas de EE.UU. en general): la fecha suele venir en orden
    // distinto y el OCR a veces invierte día/mes, mandando el gasto a un mes
    // anterior equivocado. Forzamos proveedor/folio/tipo y resolvemos la fecha
    // de forma determinista (ver fechaDesdeNombreUS más abajo).
    const esMicrosoft = /microsoft/i.test(`${parsed.proveedor || ''} ${parsed.concepto || ''}`)
    // GW Instek: el membrete dice "GW INSTEK" pero la razón social del proveedor
    // es "INSTEK AMERICA CORP". Forzamos el nombre correcto.
    const esInstek = /gw\s*instek|instek\s*america/i.test(`${parsed.proveedor || ''} ${parsed.concepto || ''}`)
    // FNI peaje: el server ya devuelve proveedor "Fondo Nacional de
    // Infraestructura" (y el folio Serie+Folio) cuando detecta el PDF de FNI.
    const esFNI = /fondo nacional de infraestructura|peaje y cruce/i.test(`${parsed.proveedor || ''} ${parsed.concepto || ''}`)
    // Vanstron (proforma invoice de equipo/maquinaria de importación): el
    // servidor ya devuelve proveedor "Vanstron Automation" y el folio
    // determinista. Su tipo siempre es COGS (costo de la mercancía).
    const esVanstron = /vanstron/i.test(`${parsed.proveedor || ''} ${parsed.concepto || ''}`)
    // ICT Holding: mismo caso que Vanstron (proforma de equipo PCBA importado
    // sin XML). El servidor devuelve proveedor "ICT Holding", folio "PI No." y
    // concepto de la columna Machine. Su tipo también es COGS.
    const esICT = /ict holding/i.test(`${parsed.proveedor || ''} ${parsed.concepto || ''}`)
    // Essemtec USA LLC: máquinas pick&place FOX2 importadas, invoice sin XML.
    // Servidor devuelve proveedor/folio/concepto deterministas. Tipo = COGS.
    const esEssemtec = /essemtec/i.test(`${parsed.proveedor || ''} ${parsed.concepto || ''}`)
    // SAT (acuse de declaración de impuestos federales): el servidor devuelve
    // proveedor "Servicio de Administración Tributaria", folio (Número de
    // operación) y total (suma de "Cantidad a pagar"). RFC del SAT y tipo ISR.
    const esSAT = /servicio de administraci[oó]n tributaria/i.test(`${parsed.proveedor || ''}`)
    let proveedorFinal = parsed.proveedor || ''
    let rfcFinal = ''
    let folioFinal = parsed.folio || parsed.approval_code || ''
    let fechaFinal = forzarAñoEnCurso(parsed.fecha || today)
    let tipoFinal = autoDetectTipo(parsed.proveedor || '', parsed.concepto || '', usaTiposExtendidos(colaborador) ? undefined : categoriaEfectiva(colaborador), '', parsed.rfc || '')
    if (esITESO) {
      proveedorFinal = 'ITESO'
      tipoFinal = 'Renta Oficina'
      // El recibo de ITESO no imprime su RFC; lo forzamos al oficial.
      rfcFinal = 'ITE570731PS9'
      if (!total && subtotal) total = subtotal
      if (!subtotal && total) subtotal = total
    } else if (esISN) {
      proveedorFinal = 'Secretaria de la Hacienda Pública'
      tipoFinal = '3% ISN'
      // RFC de la Secretaría de la Hacienda Pública (Jalisco); el recibo no lo
      // imprime, lo forzamos al oficial.
      rfcFinal = 'SPC130227L99'
      if (!total && subtotal) total = subtotal
      if (!subtotal && total) subtotal = total
    } else if (esMicrosoft) {
      proveedorFinal = 'Microsoft'
      tipoFinal = 'IT & SW'
      // RFC fiscal de Microsoft México (el Billing Summary lo imprime como
      // "RFC: MCO091123MR8"; el OCR no siempre lo captura).
      rfcFinal = 'MCO091123MR8'
      // Fecha determinista: Microsoft nombra sus PDF "...Office MM-DD-YY.pdf".
      // Preferimos esa fecha sobre la del OCR para que el gasto caiga en el mes
      // correcto y nunca en un mes anterior por un día/mes invertido.
      fechaFinal = forzarAñoEnCurso(fechaDesdeNombreUS(fileName) || parsed.fecha || today)
      // Billing Number determinista desde el nombre ("Microsoft G160071537 ...").
      // CRÍTICO: como forzamos un RFC constante para Microsoft, la dedup
      // (rfc|noFactura) colisiona si el OCR repite el folio entre dos facturas
      // casi idénticas — y una se descartaría como "duplicado". El número del
      // nombre del archivo es único por factura y evita esa pérdida.
      const billingDesdeNombre = (fileName || '').match(/\bG\d{6,}\b/i)
      if (billingDesdeNombre) folioFinal = billingDesdeNombre[0].toUpperCase()
    } else if (esTotalPlay) {
      proveedorFinal = 'Total Play'
      tipoFinal = 'IT & SW'
      // El OCR ya extrae "Cargos del Mes" en total. Recalculamos el IVA como
      // total − subtotal para que la fila quede consistente (importe + IVA =
      // total, igual que la fórmula TOTAL del Excel) aunque el IVA impreso no
      // cuadre con esa resta.
      if (total > 0 && subtotal > 0 && total >= subtotal) {
        iva = parseFloat((total - subtotal).toFixed(2))
      }
    } else if (esInstek) {
      proveedorFinal = 'INSTEK AMERICA CORP'
    } else if (esVanstron) {
      proveedorFinal = 'Vanstron Automation'
      tipoFinal = 'COGS'
    } else if (esICT) {
      proveedorFinal = 'ICT Holding'
      tipoFinal = 'COGS'
    } else if (esEssemtec) {
      proveedorFinal = 'Essemtec USA LLC'
      tipoFinal = 'COGS'
    } else if (esSAT) {
      proveedorFinal = 'Servicio de Administración Tributaria'
      rfcFinal = 'SAT970701NN3'
      tipoFinal = 'ISR'
      if (!total && subtotal) total = subtotal
      if (!subtotal && total) subtotal = total
    } else if (esFNI) {
      proveedorFinal = 'Fondo Nacional de Infraestructura'
      rfcFinal = 'FNI970829JR9'
      tipoFinal = tipoPeajeFNI(colaborador)
    }

    return {
      id: genId(),
      rfc: rfcFinal,
      proveedor: proveedorFinal,
      // Prefer the merchant's own folio when present, fall back to the card
      // approval code so it lines up with the bank statement's "Código de
      // autorización" column in Pass 0 of validarBanco.
      noFactura: folioFinal
        ? String(folioFinal)
        : ('TKT-' + uuid.slice(0, 6).toUpperCase()),
      fechaFac: fechaFinal,
      concepto: parsed.concepto || '',
      tipo: tipoFinal,
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
      fechaCobro: fechaFinal,
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
      banco:              defaultBancoFor(colaborador),
      // Snapshot del total al momento del OCR (sin tip).
      montoFacturado:     isExtranjera ? 0 : total,
      // Snapshot de la propina detectada por OCR — congelada al inicio.
      // Si el usuario después edita montoPropina manualmente, ese cambio
      // genera DIFERENCIA en el Excel. Para CFDIs/manual donde no hay
      // propina al crear, queda en 0 y cualquier propina agregada después
      // aparece como discrepancia.
      montoPropinaOriginal: isExtranjera ? 0 : propina,
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
      subtitle: `Se procesarán ${files.length} ${files.length === 1 ? 'archivo' : 'archivos'} con OCR (OpenAI). Costo estimado: ≈ ${fmtUsd(files.length * OCR_COST_USD)} USD (~${fmtUsd(OCR_COST_USD)}/factura). ¿Continuar?`,
      primaryLabel: 'Continuar',
      secondaryLabel: 'Cancelar',
    })
    if (!confirmed) { e.target.value = ''; return }

    // Si hay PDFs muy grandes, avisa que se comprimirán antes de procesar.
    const continuarGrandes = await avisarPdfsGrandes(files)
    if (!continuarGrandes) { e.target.value = ''; return }

    setOcrLoading(true)
    const newGastos = []
    for (let i = 0; i < files.length; i++) {
      const file = files[i]
      setOcrProgress({ current: i + 1, total: files.length })
      if (i > 0) await sleep(OCR_PACING_MS)
      try {
        const isPdfSource = file.name.toLowerCase().endsWith('.pdf')
        const isLargePdf = isPdfSource && file.size > LARGE_PDF_BYTES
        let fileForOCR, mediaType
        if (isLargePdf) {
          // PDF grande (escaneo): rasteriza a JPEG chico antes de subir.
          const jpg = await pdfFirstPageToJpeg(file)
          if (jpg) { fileForOCR = jpg; mediaType = 'image/jpeg' }
          else { fileForOCR = file; mediaType = 'application/pdf' }
        } else {
          fileForOCR = await compressImage(file)
          mediaType = fileForOCR.type || 'image/jpeg'
        }
        const base64 = await fileToBase64(fileForOCR)
        const gasto = await extractReceiptData(base64, mediaType, file.name)
        if (gasto) {
          gasto.esMonedaExtranjera = !!(gasto.moneda && gasto.moneda !== 'MXN')
          if (isPdfSource) {
            gasto.pdfFile = file
            gasto.pdfDataURL = isLargePdf
              ? await fileToDataURL(file)
              : `data:application/pdf;base64,${base64}`
            gasto.tienePDF = true
          } else if (mediaType.startsWith('image/')) {
            // Stash the (compressed) image data URL so the ZIP export can
            // wrap it into a single-page PDF named via buildFileName.
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
    setOcrProgress(null)
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

    // Detectar XMLs corruptos (0 KB)
    const xmlCorruptos = xmlFiles.filter(f => f.size === 0)
    const xmlValidos   = xmlFiles.filter(f => f.size > 0)
    // Nombres base (normalizados) de XMLs corruptos: el PDF que comparta nombre
    // base irá a OCR (ya cae en unmatchedPDFs porque su XML no produjo gasto) y
    // el gasto resultante se marca xmlFaltante para señalarlo en el Excel.
    const normDrop = s => (s || '').normalize('NFD').replace(/\p{Diacritic}/gu, '').toLowerCase()
    const corruptDropBases = new Set(
      xmlCorruptos.map(f => normDrop(f.name.replace(/\.xml$/i, '')))
    )
    if (xmlCorruptos.length > 0) {
      await askConfirm({
        type: 'warning',
        title: 'XMLs corruptos detectados',
        subtitle: `${xmlCorruptos.length} archivo${xmlCorruptos.length > 1 ? 's' : ''} XML ${xmlCorruptos.length > 1 ? 'están vacíos' : 'está vacío'} (0 KB) y no ${xmlCorruptos.length > 1 ? 'pudieron' : 'pudo'} procesarse.`,
        body: xmlCorruptos.map(f => `• ${f.name}`).join('\n') +
          '\n\nDescarga nuevamente el XML y su PDF con el mismo nombre desde el portal del SAT o del proveedor y vuelve a subirlos.',
        primaryLabel: 'Entendido',
      })
    }

    // STEP 1 — parse XMLs locally (always free). parseCFDI takes the full
    // signature so xmlFile + auto-PDF-link + auto-tipo all work.
    for (const file of xmlValidos) {
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

    // STEP 2 — vincular los PDFs sueltos a su gasto de forma EXCLUSIVA
    // (un PDF → un gasto, por prioridad de confianza). parseCFDI ya enlazó
    // los de nombre/UUID exacto; esto completa folio/RFC/substring sin
    // duplicar. Lo que quede sin pareja son candidatos a OCR.
    const unmatchedNombre = linkPdfsExclusive(allNewGastos, pdfFiles)
    // NIVEL 6: enlaza por TEXTO embebido del PDF (UUID/folio/RFC en el contenido)
    // los gastos que siguen sin PDF, para PDFs renombrados a algo arbitrario.
    const unmatchedPDFs = await linkPdfsByContent(allNewGastos, unmatchedNombre)

    // FNI (peaje): completa el Serie+Folio desde el PDF cuando el XML no lo trae
    // (ver enrichFniFoliosDesdePDF). Mismo arreglo que en la carga de carpeta.
    await enrichFniFoliosDesdePDF(allNewGastos)
    await enrichTelcelParcialidad(allNewGastos)

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
          { value: `≈ ${fmtUsd(ocrFiles.length * OCR_COST_USD)}`, label: 'Costo estim. (USD)' },
        ],
        primaryLabel: 'Procesar con IA',
        secondaryLabel: 'Cancelar',
      })

      // Si hay PDFs muy grandes, avisa que se comprimirán antes de procesar.
      const continuarGrandes = userConfirmed ? await avisarPdfsGrandes(ocrFiles) : false
      if (userConfirmed && continuarGrandes) {
        setOcrLoading(true)
        for (let i = 0; i < ocrFiles.length; i++) {
          const file = ocrFiles[i]
          setOcrProgress({ current: i + 1, total: ocrFiles.length })
          if (i > 0) await sleep(OCR_PACING_MS)
          try {
            // compressImage re-encodes images to a ≤2000px-wide JPEG so
            // phone shots don't blow Vercel's 4.5 MB body cap. PDFs chicos
            // pasan tal cual; los PDFs GRANDES (escaneos) se rasterizan a un
            // JPEG chico en el navegador para que la subida sea rápida.
            const isPdfSource = file.name.toLowerCase().endsWith('.pdf')
            const isLargePdf = isPdfSource && file.size > LARGE_PDF_BYTES
            let fileForOCR, mediaType
            if (isLargePdf) {
              const jpg = await pdfFirstPageToJpeg(file)
              if (jpg) { fileForOCR = jpg; mediaType = 'image/jpeg' }
              else { fileForOCR = file; mediaType = 'application/pdf' }
            } else {
              fileForOCR = await compressImage(file)
              const ext = file.name.split('.').pop().toLowerCase()
              mediaType = fileForOCR.type
                || (ext === 'pdf' ? 'application/pdf' : `image/${ext === 'jpg' ? 'jpeg' : ext}`)
            }
            const base64 = await fileToBase64(fileForOCR)

            const gasto = await extractReceiptData(base64, mediaType, file.name)
            if (gasto) {
              // El PDF ORIGINAL siempre se adjunta para el ZIP. Si rasterizamos
              // (PDF grande), el base64 es el JPEG, así que leemos el PDF aparte.
              if (isPdfSource) {
                gasto.pdfFile = file
                gasto.pdfDataURL = isLargePdf
                  ? await fileToDataURL(file)
                  : `data:application/pdf;base64,${base64}`
                gasto.tienePDF = true
              } else if (mediaType.startsWith('image/')) {
                // For OCR'd images, stash the data URL so exportar can wrap
                // it into a single-page PDF named via buildFileName.
                gasto.imageDataURL = `data:${mediaType};base64,${base64}`
                gasto.originalFileName = file.name
              }
              // Rescate de XML corrupto: si este archivo comparte nombre base
              // con un .xml de 0 KB, marca el gasto para señalar "Falta XML".
              if (corruptDropBases.has(normDrop(file.name.replace(/\.[^.]+$/i, '')))) {
                gasto.xmlFaltante = true
              }
              gasto.isNew = true
              allNewGastos.push(gasto)
            }
          } catch (err) {
            console.warn('OCR error:', file.name, err)
            showModal({
              type: 'error',
              title: 'Error de OCR',
              subtitle: `${file.name}\n\n${err && err.message ? err.message : String(err)}`,
              primaryLabel: 'Entendido',
            })
          }
        }
        setOcrLoading(false)
        setOcrProgress(null)
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

    // STEP 5 — detect duplicates, show modal if any, then merge.
    const applyDropMerge = (incoming) => {
      setLista(prev => {
        const merged = [...prev]
        const keys = new Set(prev.map(g => gastoDedupKey(g)))
        let added = 0
        let updated = 0
        for (const newG of incoming) {
          const key = gastoDedupKey(newG)
          if (keys.has(key)) {
            const idx = merged.findIndex(g => gastoDedupKey(g) === key)
            if (idx !== -1) {
              // Si el entrante trae XML real (xmlContent), el usuario re-subió
              // el XML correcto → limpia la marca xmlFaltante.
              const clearFlag = !!newG.xmlContent
              merged[idx] = {
                ...merged[idx], ...newG, isNew: true,
                xmlFaltante: clearFlag ? false : (newG.xmlFaltante || merged[idx].xmlFaltante),
              }
              updated++
            }
          } else {
            merged.push(newG); keys.add(key); added++
          }
        }
        if (added > 0 || updated > 0 || pdfFiles.length > 0) {
          setTimeout(() => setDropSummary({ added, updated, pdfs: pdfFiles.length }), 100)
        }
        return merged
      })
      setTimeout(() => {
        setLista(l => l.map(g => g.isNew ? { ...g, isNew: false } : g))
      }, 1500)
    }

    // Duplicates against existing lista AND within the incoming batch itself
    const dropDupes = []
    const seenInBatch = new Set()
    for (const g of allNewGastos) {
      const key = `${g.rfc}|${g.noFactura}`
      const inLista = lista.some(e => e.rfc === g.rfc && e.noFactura === g.noFactura)
      const inBatch = seenInBatch.has(key)
      if (inLista || inBatch) dropDupes.push(g)
      seenInBatch.add(key)
    }

    if (dropDupes.length > 0) {
      setDuplicadosModal({
        items: dropDupes,
        incoming: allNewGastos,
        onConfirm: (kept) => {
          setDuplicadosModal(null)
          applyDropMerge(kept)
        },
      })
    } else {
      applyDropMerge(allNewGastos)
    }
  }

  // ── Validar estado de cuenta bancario ──
  const validarBanco = async e => {
    const file = e.target.files[0]; if (!file) return
    const buffer = await file.arrayBuffer()
    const isExcel = /\.xlsx?$/i.test(file.name || '')

    // Normalizamos ambos formatos a `records` = filas [{ cols[], raw }] para
    // que el resto del parser sea agnóstico al origen. `cols` son las celdas;
    // `raw` es la línea original (para el escaneo genérico de bancos no-Clara).
    let records = []
    let headerLine = ''
    if (isExcel) {
      // Estado de cuenta (Clara u otro) exportado a Excel en vez de CSV
      // — p.ej. el formato de Natividad. Leemos la primera hoja tal cual;
      // las fechas llegan como Date (cellDates) y las pasamos a YYYY-MM-DD
      // para que parseDateRobusto las entienda igual que en el CSV.
      const wb = XLSX.read(buffer, { type: 'array', cellDates: true })
      const ws = wb.Sheets[wb.SheetNames[0]]
      const aoa = ws ? XLSX.utils.sheet_to_json(ws, { header: 1, blankrows: false, raw: true }) : []
      const cellToStr = v => {
        if (v == null) return ''
        if (v instanceof Date) {
          const yyyy = v.getFullYear()
          const mm   = String(v.getMonth() + 1).padStart(2, '0')
          const dd   = String(v.getDate()).padStart(2, '0')
          return `${yyyy}-${mm}-${dd}`
        }
        return String(v)
      }
      records = aoa.map(row => {
        const cols = (Array.isArray(row) ? row : []).map(cellToStr)
        return { cols, raw: cols.join(',') }
      })
      headerLine = (records[0]?.cols || []).join(',')
    } else {
      // Read as bytes and pick the encoding: try strict UTF-8 first, fall back
      // to ISO-8859-1 (latin-1) which is what Clara USA exports. Without this,
      // accented chars (Código, Autorización, etc.) come back as mojibake.
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
      records = lines.map(line => ({ cols: parseCSVLine(line, sep), raw: line }))
      headerLine = lines[0] || ''
    }

    // Clara-platform CSVs have a fixed column layout. Detect by any of the
    // header markers we know about — the MX and USA variants share most
    // columns, plus the USA file carries "Código de autorización" (col 12)
    // and "Moneda original" (col 4) which power the ticket Pass 0 match.
    // "Transacci"/"autorizaci" intentionally drop the accented chars so we
    // survive encoding mishaps regardless of UTF-8 vs latin-1.
    const isClara = headerLine.includes('Fecha de Transacci')
                 || headerLine.includes('digo de autorizaci')
                 || headerLine.includes('Moneda original')

    // Las posiciones de columna de Clara difieren entre el CSV (trae una
    // columna extra "Estado de Cuenta" + "Código de autorización" +
    // "Categoría") y el Excel (formato de Natividad: sin esas tres, todo
    // corrido una a la izquierda). Resolvemos cada columna por su encabezado
    // para que AMBOS layouts funcionen; si un encabezado no aparece, caemos a
    // los índices clásicos del CSV (columnas opcionales quedan en -1 → vacío).
    const headerCells = (records[0]?.cols || []).map(c =>
      String(c || '').toLowerCase().normalize('NFD').replace(/\p{Diacritic}/gu, '').trim())
    const findCol = (needle, fallback) => {
      const i = headerCells.findIndex(h => h.includes(needle))
      return i >= 0 ? i : fallback
    }
    const findColExact = (val, fallback) => {
      const i = headerCells.findIndex(h => h === val)
      return i >= 0 ? i : fallback
    }
    // "fecha de transaccion" también contiene "transaccion", por eso la
    // descripción se busca con match exacto del encabezado "Transacción".
    const colFecha    = findCol('fecha de transacc', 0)
    const colDesc     = findColExact('transaccion', 2)
    const colMontoOrg = findCol('monto original', 3)
    const colMoneda   = findCol('moneda original', 4)
    const colMontoMXN = findCol('monto en mxn', 5)
    const colAuth     = findCol('autorizacion', -1)
    const colCat      = findCol('categoria', -1)

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
      // COBRADO real: el monto que efectivamente cargó la tarjeta según el CSV
      // del banco (en pesos). Se guarda independiente de factura+propina para
      // que la columna DIFERENCIA del Excel pueda detectar un excedente real
      // (cargo > factura+propina), no sólo reconstruir factura+propina.
      g.montoCobrado = row.montoMXN || 0
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
        // Propina ya aplicada por el pass que acaba de correr (Pass 1 escalera
        // / Pass 2 delta). El modal la usa para explicar la aritmética
        // "factura + propina = cargo" en vez de solo decir "propina detectada".
        propina: g.montoPropina || 0,
        propinaPct: g.propinaPorcentaje || 0,
      })
    }
    const formatCobro = d => {
      // Store as YYYY-MM-DD so the native date picker on the row accepts it
      // directly and formatDateDisplay produces DD-MM-YYYY for the read view.
      const dd   = String(d.getDate()).padStart(2, '0')
      const mm   = String(d.getMonth() + 1).padStart(2, '0')
      const yyyy = d.getFullYear()
      return `${yyyy}-${mm}-${dd}`
    }
    const cleanNum = s => parseFloat(String(s || '').replace(/[$,\s]/g, ''))

    // Step 1 — parse every CSV row into a flat list { dCSV, amounts[], descripcion }.
    const csvRows = []
    for (let li = 0; li < records.length; li++) {
      const rec = records[li]
      const cols = rec?.cols || []
      const line = rec?.raw || ''
      if (!cols.some(c => String(c).trim())) continue   // fila vacía
      if (isClara && li === 0) continue   // skip Clara header row

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
        dCSV = parseDateRobusto(cols[colFecha] || '')
        montoMXN = cleanNum(cols[colMontoMXN]) || 0
        montoUSD = cleanNum(cols[colMontoOrg]) || 0
        moneda = (cols[colMoneda] || 'MXN').trim().toUpperCase() || 'MXN'
        autorizacion = colAuth >= 0 ? String(cols[colAuth] || '').trim() : ''
        categoria = colCat >= 0 ? (cols[colCat] || '').toString().trim() : ''
        const amount = montoMXN || montoUSD || 0
        if (!dCSV || !amount) continue
        // Carry BOTH the MXN and USD figures into the amount candidates so
        // smartAmountMatch can find ticket receipts whose OCR was captured in
        // either currency without depending on which column is non-zero.
        amounts = [montoMXN, montoUSD].filter(v => v > 0)
        descripcion = (cols[colDesc] || '').trim().slice(0, 60)
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
    // Rellena el lado peso/divisa de un ticket extranjero desde la fila del
    // banco. Antes SOLO Pass 0 (auth code) lo hacía; Pass 1 (monto) y Pass 2
    // (propina) dejaban importe/totalCFDI en 0, así que en el Excel FACTURADO
    // quedaba en 0 y DIFERENCIA pintaba TODA la divisa como faltante. Ahora los
    // tres passes usan este helper. Semántica del Excel (export-excel.py):
    // totalCFDI = cargo MXN SIN propina, montoPropina = tip → TOTAL = totalCFDI +
    // propina = COBRADO y DIFERENCIA = 0. La propina en divisa se calcula sobre
    // el TOTAL del ticket (no el subtotal) para NO contar el IVA como propina.
    const isForeignRow = row => !!(row.moneda && row.moneda !== 'MXN' && row.montoUSD > 0)
    const fillForeignFromBank = (inv, row) => {
      inv.monedaCodigo = row.moneda
      inv.moneda = row.moneda
      inv.esMonedaExtranjera = true
      inv.tipoCambio = row.montoUSD > 0 ? +(row.montoMXN / row.montoUSD).toFixed(4) : 0
      let tipMxn = 0, tipExt = 0
      const ticketTotal = inv.montoExtranjero || 0
      if (row.montoUSD > ticketTotal && ticketTotal > 0) {
        const delta = +(row.montoUSD - ticketTotal).toFixed(2)
        const tipPct = (delta / ticketTotal) * 100
        if (tipPct >= 5 && tipPct <= 35) {
          tipExt = delta
          inv.propinaExtranjero = delta
          tipMxn = +(delta * inv.tipoCambio).toFixed(2)
          inv.montoPropina = tipMxn
          inv.propinaPorcentaje = +tipPct.toFixed(2)
        }
      }
      // El renglón PADRE va NETO (sin propina) en AMBAS monedas, para que padre
      // + sub-fila de propina sumen el cargo completo. MXN neto = montoMXN −
      // tipMxn; divisa neto = montoUSD − tipExt (= el total del ticket). La
      // propina vive en montoPropina / propinaExtranjero. Antes el lado USD se
      // quedaba con el total completo → padre 84.52 + propina 10 se veía doble
      // e inflaba el KPI de USD.
      inv.importe   = +(row.montoMXN - tipMxn).toFixed(2)
      inv.totalCFDI = +(row.montoMXN - tipMxn).toFixed(2)
      inv.montoExtranjero = +(row.montoUSD - tipExt).toFixed(2)
      inv.montoUSD = +(row.montoUSD - tipExt).toFixed(2)
    }
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

      if (isForeignRow(row)) {
        fillForeignFromBank(nl[idx], row)
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
      'wings', 'mariscos', 'marisqueria', 'marisquería', 'asador', 'birria',
      'antojitos', 'burger', 'burguer', 'pub', 'brewery', 'cerveceria', 'cervecería',
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
    const asReceipt = inv => ({
      _raw: inv,  // kept so smartAmountMatch can re-check eligibility off proveedor/concepto
      isOCR: (inv.subtotal || 0) > 0 || (inv.propinaSugerida18 || inv.propinaSugerida20 || inv.propinaSugerida22) > 0,
      isTicket: !!inv.esTicket,
      subtotalOCR: inv.subtotal || 0,
      totalCFDI: inv.totalCFDI || 0,
      // Total OCR = lo que el banco realmente cobra (con impuestos). En moneda
      // extranjera vive en montoExtranjero; en MXN, en totalCFDI.
      totalOCR: inv.montoExtranjero || inv.totalCFDI || 0,
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
      // Tickets OCR: el banco cobra el TOTAL (con impuestos), no el subtotal.
      // Probar también el total rescata recibos con IVA (hoteles, comidas sin
      // propina) cuyo subtotal < total, y los Uber cuyo OCR no trajo subtotal.
      // El subtotal + propina de abajo sigue cubriendo las comidas con propina.
      if (receipt.isTicket && receipt.totalOCR > 0) {
        candidates.push({ amount: receipt.totalOCR, propina: 0, pct: 0, label: 'exact' })
      }
      if (eligible) {
        if (receipt.isOCR) {
          const sub = receipt.subtotalOCR
          if (receipt.propinaSugerida18 > 0) candidates.push({ amount: sub + receipt.propinaSugerida18, propina: receipt.propinaSugerida18, pct: 18, label: 'ocr18' })
          if (receipt.propinaSugerida20 > 0) candidates.push({ amount: sub + receipt.propinaSugerida20, propina: receipt.propinaSugerida20, pct: 20, label: 'ocr20' })
          if (receipt.propinaSugerida22 > 0) candidates.push({ amount: sub + receipt.propinaSugerida22, propina: receipt.propinaSugerida22, pct: 22, label: 'ocr22' })
        } else {
          const total = receipt.totalCFDI
          ;[0.10, 0.12, 0.13, 0.15, 0.16, 0.18, 0.20, 0.22, 0.25].forEach(p => {
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
        // Extranjera: rellena peso / T-C / propina desde el banco (igual que
        // Pass 0) para que FACTURADO = COBRADO y no aparezca diferencia falsa.
        if (isForeignRow(row)) fillForeignFromBank(inv, row)
        const { method, confidence } = methodFromMatch(matchResult)
        snapshotMatch(idx, row, 1, method, confidence)
        matches++
      }
    )

    // Pass 2 — propina por delta, SOLO para pares (factura, cargo) elegibles
    // a propina. La escalera de Pass 1 falla con propinas atípicas (16%) o
    // cuando el mesero redondea el monto a mano (15% cobrado como $88.79 en
    // vez de $88.80): cualquier centavo fuera de la escalera dejaba la fila
    // sin match y la propina se perdía del reporte. Aquí, si el cargo del
    // banco excede el total de la factura por un porcentaje plausible de
    // propina (5%–35%, misma ventana que ya usa Pass 0 para moneda
    // extranjera), el delta ES la propina — exacto al centavo porque viene
    // del cargo real, no de un porcentaje teórico. Corre DESPUÉS de los
    // passes exactos para que éstos reclamen sus filas primero, y el gate de
    // categoría (Alimentos/keywords) impide que un cargo de hotel/gasolina
    // se enganche a una factura vía propina fantasma.
    const TIP_DELTA_MIN = 0.05, TIP_DELTA_MAX = 0.35
    const tipBase = inv => {
      const r = asReceipt(inv)
      return r.isOCR ? r.subtotalOCR : r.totalCFDI
    }
    tryPass(
      (inv, m, row) => {
        if (!isEligibleForTip(inv, row)) return false
        // Candado de moneda: comparar SOLO contra el monto en la MISMA divisa
        // que la factura. Sin esto, un subtotal en divisa (GBP/EUR/USD) se
        // comparaba contra el "Monto en MXN" del cargo (o un ticket USD contra
        // un cargo EUR) e inventaba propinas fantasma cruzando monedas.
        const invForeign = inv.esMonedaExtranjera || (inv.moneda && inv.moneda !== 'MXN')
        if (invForeign) {
          if (!row.moneda || row.moneda !== inv.moneda) return false
          if (!(row.montoUSD > 0) || Math.abs(m - row.montoUSD) > 0.001) return false
        } else if (!(row.montoMXN > 0) || Math.abs(m - row.montoMXN) > 0.001) {
          return false
        }
        const base = tipBase(inv)
        if (!(base > 0) || m <= base) return false
        const pct = (m - base) / base
        return pct >= TIP_DELTA_MIN && pct <= TIP_DELTA_MAX
      },
      (idx, m, dCSV, row) => {
        const inv = nl[idx]
        inv.hizoMatch = true
        inv.fechaCobro = formatCobro(dCSV)
        inv.formaPago = '04'
        if (isForeignRow(row)) {
          // Extranjera: el helper fija peso + T-C y calcula la propina sobre el
          // TOTAL del ticket (no el subtotal), para no contar el IVA como tip.
          fillForeignFromBank(inv, row)
          if ((inv.montoPropina || 0) > 0) propinas++
          snapshotMatch(idx, row, 2, `Propina detectada +${(inv.propinaPorcentaje || 0).toFixed(1)}%`, 80)
        } else {
          const base = tipBase(inv)
          const tip = +(m - base).toFixed(2)
          const pct = +((tip / base) * 100).toFixed(2)
          inv.montoPropina = tip
          inv.propinaPorcentaje = pct
          propinas++
          snapshotMatch(idx, row, 2, `Propina detectada +${pct.toFixed(1)}%`, 80)
        }
        matches++
      }
    )

    // Pass 3 — casi exacto. Cargo y factura difieren por una cantidad pequeña
    // (≤1% Y ≤$20): típico de un monto tecleado a mano en la terminal del
    // comercio (hotel cobró 7,334 con factura de 7,344) o un ajuste menor.
    // SÍ vincula la fila — fechaCobro/montoCobrado reales — pero con
    // confianza 70, debajo del umbral de 80, así que cae en la pestaña
    // "Revisión" del modal y la columna DIFERENCIA del Excel exhibe el
    // faltante/excedente real en vez de perder el vínculo por completo.
    // Corre al final: nunca le roba filas a los passes exactos ni de propina.
    const NEAR_ABS_MAX = 20, NEAR_PCT_MAX = 0.01
    tryPass(
      (inv, m) => {
        const base = (inv.totalCFDI || 0) + (inv.montoPropina || 0)
        if (!(base > 0)) return false
        const diff = Math.abs(m - base)
        return diff > 0.01 && diff <= NEAR_ABS_MAX && diff / base <= NEAR_PCT_MAX
      },
      (idx, m, dCSV, row) => {
        const inv = nl[idx]
        const diff = m - ((inv.totalCFDI || 0) + (inv.montoPropina || 0))
        inv.hizoMatch = true
        inv.fechaCobro = formatCobro(dCSV)
        inv.formaPago = '04'
        const signo = diff > 0 ? '+' : '−'
        snapshotMatch(idx, row, 3, `Casi exacto (dif ${signo}$${Math.abs(diff).toFixed(2)})`, 70)
        matches++
      }
    )

    // ── GUARDIÁN: cuadre padre + propina vs cargo del banco ──────────────
    // Post-condición de la conciliación: en cada factura EXTRANJERA conciliada,
    // el renglón padre + su sub-fila de propina deben sumar el cargo del banco
    // en AMBAS monedas (MXN y divisa). fillForeignFromBank ya lo garantiza; esto
    // atrapa cualquier caso raro nuevo o un cambio futuro que rompa el split
    // ANTES de exportar: baja la confianza del match para que caiga en la
    // pestaña "Revisión" (chip amarillo) en vez de generar un Excel mal en
    // silencio. Vale para TODOS los colaboradores y reportes.
    let descuadres = 0
    for (const m of matchedRows) {
      const g = nl.find(x => x.id === m.invoiceId)
      if (!g || !g.esMonedaExtranjera) continue
      const mxnCuadra = Math.abs((g.totalCFDI || 0) + (g.montoPropina || 0) - (m.csvAmountMXN || 0)) <= 0.02
      const usdCuadra = Math.abs((g.montoExtranjero || 0) + (g.propinaExtranjero || 0) - (m.csvAmountUSD || 0)) <= 0.02
      if (!mxnCuadra || !usdCuadra) {
        m.confidence = Math.min(m.confidence, 40)
        m.method = '⚠ Descuadre padre+propina vs cargo — revisar'
        m.descuadre = true
        descuadres++
      }
    }
    if (descuadres > 0) console.warn(`[conciliación] ${descuadres} factura(s) extranjera(s) con descuadre → Revisión`)

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
        // amounts viene vacío para devoluciones (los montos negativos se
        // filtran al parsear) — Math.max() de un arreglo vacío da -Infinity
        // y el modal mostraba "$-∞". Cae al monto MXN/USD crudo con signo,
        // para que una devolución se vea como -109.96.
        monto: row.amounts.length ? Math.max(...row.amounts) : (row.montoMXN || row.montoUSD || 0),
        descripcion: row.descripcion,
        montoMXN: row.montoMXN || 0,
        montoUSD: row.montoUSD || 0,
        moneda: row.moneda || 'MXN',
        sugerencias,
      })
    }

    // Comprobantes que el usuario SÍ subió pero que no aparecen en este estado
    // de cuenta. No es un error ni un adeudo: normalmente son de otro periodo,
    // se pagaron con otra tarjeta o en efectivo. Se listan aparte para que
    // nadie lea "sin cargo" como "debo dinero".
    const facturasSinCargo = nl
      .filter(g => !g.hizoMatch)
      .map(g => ({
        id: g.id,
        proveedor: g.proveedor || '(sin proveedor)',
        noFactura: g.noFactura || '',
        fecha: g.fechaFac || '',
        total: g.totalCFDI || 0,
        moneda: g.moneda || 'MXN',
        isTicket: !!g.esTicket,
      }))

    // Surface the total foreign-currency matches at the end so the modal can
    // show "X en moneda extranjera vinculadas" — covers Pass 0 hits plus any
    // existing row already flagged esMonedaExtranjera that got matched.
    const foreignMatchesTotal = nl.filter(g => g.esMonedaExtranjera && g.fechaCobro).length

    // Aggregate currency totals (Matched vs Pending) for the top summary
    // bar — split MXN vs USD so the chips stay readable across mixed sheets.
    // Ojo: en el CSV de Clara "Monto original" viene lleno SIEMPRE, también en
    // cargos en pesos, así que montoUSD == montoMXN en un estado MXN puro.
    // Sumar eso a ciegas pintaba un "US$24,438.80 USD" fantasma al lado del
    // total en pesos y la gente leía el doble de dinero. Solo acumulamos la
    // divisa cuando la fila realmente NO es MXN.
    const totalsMatched = matchedRows.reduce((acc, m) => {
      acc.mxn += m.csvAmountMXN || 0
      if ((m.csvMoneda || 'MXN') !== 'MXN') acc.usd += m.csvAmountUSD || 0
      return acc
    }, { mxn: 0, usd: 0 })
    const totalsPending = sinFactura.reduce((acc, s) => {
      acc.mxn += s.montoMXN || (s.moneda === 'MXN' ? s.monto : 0)
      if ((s.moneda || 'MXN') !== 'MXN') acc.usd += s.montoUSD || s.monto || 0
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
      facturasSinCargo,
      docsSubidos: nl.length,
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
      montoUSD: 0, tipoCambio: 0, moneda: 'MXN', banco: defaultBancoFor(colaborador), montoFacturado: 0, montoPropinaOriginal: 0, montoCobrado: 0,
    }])
  }

  // ── Copiar a portapapeles (TSV para Excel) ──
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
    // Proveedor \u2014 Title Case, nombre completo (sin recorte), brackets stripped
    const prov = toTitleCase(
      (g.proveedor || 'Proveedor')
        .replace(/[\/\\:*?"<>|()\[\]{}]/g, '')
        .replace(/[-_]+/g, ' ')
        .replace(/\s+/g, ' ')
        .trim()
    )

    // Folio \u2014 preserved exactly as detected. Las diagonales / y \\ no son
    // v\u00e1lidas en nombres de archivo, pero forman parte del n\u00famero fiscal de los
    // CFDI de Odoo ("INV/2026/00591"); las convertimos a guion en vez de
    // borrarlas para conservar el n\u00famero completo ("INV-2026-00591").
    const folio = (g.noFactura || 'SN')
      .replace(/[\/\\]+/g, '-')
      .replace(/[:*?"<>|()\[\]{}]/g, '')
      .trim()

    // Concepto \u2014 primera l\u00ednea de la descripci\u00f3n del CFDI (g.concepto ya viene
    // recortado a la primera l\u00ednea desde parseCFDI). Title Case, m\u00e1x 30 chars
    // para que el renombre no se haga gigante. Reemplaza al antiguo "tipo"
    // (categor\u00eda) en el nombre del archivo.
    const rawConcepto = (g.concepto || g.tipo || 'Gasto')
      .replace(/[\/\\:*?"<>|()\[\]{}]/g, '')
      .replace(/\s+/g, ' ')
      .trim()
    let concepto = toTitleCase(rawConcepto).slice(0, 30).trim()
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
    // RFC — facturas extranjeras solo-PDF (OCR) no traen RFC mexicano, así que
    // en el Excel se muestran como "NA". Los CFDI (incluidas las extranjeras
    // con RFC genérico XEXX010101000) conservan su RFC real.
    const rfcParaExcel = (g) => {
      if (g.rfc) return g.rfc
      const m = (g.monedaCodigo || g.moneda || 'MXN').toString().toUpperCase()
      return (m !== 'MXN' && m !== 'XXX') ? 'N/A' : ''
    }
    const gastosSlim = lista.map(g => ({
      rfc:              rfcParaExcel(g),
      // PROVEEDOR siempre en MAYÚSCULAS en el Excel.
      proveedor:        (g.proveedor || '').toUpperCase(),
      tipo:             g.tipo || '',
      noFactura:        g.noFactura || '',
      fechaFac:         g.fechaFac || '',
      fechaCobro:       g.fechaCobro || '',
      concepto:         g.concepto || '',
      importe:          Number(g.importe) || 0,
      iva:              Number(g.iva) || 0,
      isrTrasladado:    Number(g.isrTrasladado) || 0,
      ishIeps:          Number(g.ishIeps ?? g.isrTrasladado) || 0,
      retencionISR:     Number(g.retencionISR) || 0,
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
      importeUSD:       Number(g.importeUSD) || 0,
      ivaUSD:           Number(g.ivaUSD) || 0,
      retencionesUSD:   Number(g.retencionesUSD) || 0,
      banco:            g.banco || '',
      montoFacturado:   Number(g.montoFacturado) || 0,
      montoPropinaOriginal: Number(g.montoPropinaOriginal) || 0,
      montoCobrado:     Number(g.montoCobrado) || 0,
      hizoMatch:        !!g.hizoMatch,
      polizaNumero:     g.polizaNumero || '',
      xmlFaltante:      !!g.xmlFaltante,
    }))
    try {
      const response = await fetch('/api/export-excel', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ gastos: gastosSlim, colaborador: colaborador?.nombre || '', polizaNumero: POLIZAS_CLARA[colaborador?.nombre] || 'N/A', polizasMap: POLIZAS_CLARA }),
      })
      if (response.ok) {
        const excelBlob = await response.blob()
        zip.file(`Reporte_${colabSlug}_${today}.xlsx`, excelBlob)
      }
    } catch (err) {
      console.warn('Excel fetch failed for ZIP:', err)
    }

    // 2) Rename each gasto's XML + PDF through buildFileName.
    // Dos blindajes contra pérdida/confusión silenciosa de archivos:
    //  (a) Colisión de nombres: si dos filas generan el MISMO buildFileName,
    //      JSZip escribiría dos entradas con el mismo nombre y al descomprimir
    //      una pisa a la otra → factura perdida. uniqueName añade " (2)", " (3)"…
    //      manteniendo el .xml y el .pdf de la fila bajo el mismo nombre.
    //  (b) Contenido duplicado: si facturas con folio DISTINTO terminan con el
    //      MISMO PDF byte a byte, casi siempre es data de origen corrupta (PDFs
    //      mal descargados o residuo de un export viejo). No lo podemos
    //      "reparar" —el PDF correcto no está en la carpeta— pero lo detectamos
    //      y avisamos al final para que el usuario revise el origen.
    const usedNames = new Set()
    const uniqueName = base => {
      let name = base, n = 2
      while (usedNames.has(name.toLowerCase())) name = `${base} (${n++})`
      usedNames.add(name.toLowerCase())
      return name
    }
    // djb2 sobre los bytes del PDF — contenido idéntico ⇒ misma huella.
    const djb2 = (readByte, len) => {
      let h = 5381
      for (let i = 0; i < len; i++) h = ((h << 5) + h + readByte(i)) | 0
      return `${h}:${len}`
    }
    const pdfHuellas = new Map()  // huella → Set<folio> (folios que comparten ese PDF)
    const registraPdf = (huella, g) => {
      if (!pdfHuellas.has(huella)) pdfHuellas.set(huella, new Set())
      pdfHuellas.get(huella).add(g.noFactura || '(sin folio)')
    }

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
        const nom = uniqueName(buildFileName(g))
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
          registraPdf(djb2(i => bytes[i], bytes.length), g)
        }
        if (g.xmlFile && !g.xmlContent) {
          facturas.file(`${nom}.xml`, await g.xmlFile.arrayBuffer())
          r++
        }
        if (g.pdfFile && !g.pdfDataURL) {
          const buf = new Uint8Array(await g.pdfFile.arrayBuffer())
          facturas.file(`${nom}.pdf`, buf)
          if (!g.xmlFile && !g.xmlContent) r++
          registraPdf(djb2(i => buf[i], buf.length), g)
        }
      }
    }
    void r

    // Grupos de folios DISTINTOS que comparten un PDF idéntico (data corrupta).
    const dupPdfGroups = [...pdfHuellas.values()].filter(set => set.size > 1)

    // 3) Generate + trigger download.
    const zipBlob = await zip.generateAsync({ type: 'blob' })
    const zipName = `SMTO_Gastos_${colabSlug}_${today}.zip`
    const a = Object.assign(document.createElement('a'), {
      href: URL.createObjectURL(zipBlob),
      download: zipName,
    })
    a.click()
    URL.revokeObjectURL(a.href)

    // 4) Success modal — generic PremiumModal with three stats. Si se detectó
    //    PDF idéntico entre folios distintos, el modal pasa a 'warning' y lista
    //    los folios afectados: el PDF de esas facturas está mal en el origen.
    const xmlCount = lista.filter(g => g.xmlContent || g.xmlFile).length
    const pdfCount = lista.filter(g => g.pdfDataURL || g.pdfFile).length
    const hayDupPdf = dupPdfGroups.length > 0
    const dupFolios = dupPdfGroups
      .map(set => [...set].join(', '))
      .slice(0, 6)
      .join('  ·  ')
    showModal({
      type: hayDupPdf ? 'warning' : 'success',
      title: hayDupPdf ? 'ZIP generado — revisa estos PDFs' : '¡ZIP Generado!',
      subtitle: hayDupPdf
        ? `Se generó el paquete, pero ${dupPdfGroups.length} grupo(s) de facturas con folio distinto comparten un PDF idéntico — el PDF de esas facturas está equivocado en la carpeta de origen, no en el renombrado. Vuelve a descargar el PDF correcto de: ${dupFolios}`
        : 'Paquete descargado con Excel + facturas renombradas.',
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
      const wb = XLSX.read(arrayBuffer, { type: 'array', cellStyles: true, cellDates: true })
      const ws = wb.Sheets[wb.SheetNames[0]]

      // El layout del export (api/export-excel.py) creció con el tiempo: se
      // insertaron PÓLIZA, ISR, ISH/IEPS, BANCO y MONTO USD entre las columnas
      // originales. Cuando este import usaba índices de columna FIJOS quedaron
      // desfasados y leía IMPORTE/IVA/TOTAL de columnas equivocadas (p.ej.
      // IMPORTE leía el texto de CONCEPTO → 0) y "no sumaba". Ahora mapeamos
      // cada columna por su ENCABEZADO (fila 9), así el import sobrevive
      // cualquier cambio de layout futuro.
      const HEADER_ROW = 9
      const normHdr = (s) => String(s ?? '').toUpperCase().normalize('NFD')
        .replace(/\p{Diacritic}/gu, '').replace(/[^A-Z0-9]/g, '')
      const colByHeader = {}
      for (let c = 0; c < 40; c++) {
        const h = normHdr(ws[XLSX.utils.encode_cell({ r: HEADER_ROW - 1, c })]?.v)
        if (h && !(h in colByHeader)) colByHeader[h] = c
      }
      const findCol = (...aliases) => {
        for (const a of aliases) { const k = normHdr(a); if (k in colByHeader) return colByHeader[k] }
        return -1
      }
      const C = {
        rfc:      findCol('RFC'),
        prov:     findCol('PROVEEDOR'),
        tipo:     findCol('TIPO'),
        factura:  findCol('FACTURA'),
        fFac:     findCol('F. FACTURA', 'F.FACTURA'),
        fCobro:   findCol('F. COBRO', 'F.COBRO'),
        concepto: findCol('CONCEPTO'),
        importe:  findCol('IMPORTE'),
        iva:      findCol('IVA'),
        isr:      findCol('ISR'),
        ish:      findCol('ISH/IEPS', 'ISH IEPS'),
        ret:      findCol('RETENCIÓN', 'RETENCION'),
        total:    findCol('TOTAL'),
        forma:    findCol('FORMA PAGO'),
        banco:    findCol('BANCO'),
        // 'MONTO M.E.' es el encabezado nuevo; 'MONTO USD' el histórico —
        // ambos deben importar (reportes viejos y nuevos).
        usd:      findCol('MONTO M.E.', 'MONTO ME', 'MONTO USD'),
        tc:       findCol('T/C', 'TC'),
      }
      // Si faltan columnas clave, el archivo no tiene el formato del reporte:
      // abortamos con aviso en vez de importar basura desalineada.
      if (C.rfc < 0 || C.importe < 0 || C.total < 0) {
        showModal({
          type: 'error',
          title: 'Formato de Excel no reconocido',
          subtitle: 'No se encontraron los encabezados (RFC, IMPORTE, TOTAL) en la fila 9. ¿Es un Excel exportado por esta app?',
          primaryLabel: 'Entendido',
        })
        return
      }

      // Fechas: con cellDates el datetime nativo del export llega como Date;
      // también toleramos strings MM-DD-YY y YYYY-MM-DD.
      const parseDate = (d) => {
        if (d == null || d === '') return ''
        if (d instanceof Date && !isNaN(d)) {
          const yyyy = d.getFullYear()
          const mm = String(d.getMonth() + 1).padStart(2, '0')
          const dd = String(d.getDate()).padStart(2, '0')
          return `${yyyy}-${mm}-${dd}`
        }
        const s = String(d)
        if (s.includes('-') && s.length === 8) {
          const [mm, dd, yy] = s.split('-')
          return `20${yy}-${mm}-${dd}`
        }
        return s
      }

      const gastos = []
      let row = 10

      while (true) {
        const rfc = ws[XLSX.utils.encode_cell({ r: row - 1, c: C.rfc })]?.v
        if (!rfc || String(rfc).trim() === '' || rfc === 'TOTAL CUENTA') break

        // .v = valor crudo; .w = texto ya formateado (útil para la etiqueta de
        // forma de pago "04 - Tarjeta de Crédito").
        const val = (c) => c >= 0 ? (ws[XLSX.utils.encode_cell({ r: row - 1, c })]?.v ?? '') : ''
        const txt = (c) => {
          if (c < 0) return ''
          const cell = ws[XLSX.utils.encode_cell({ r: row - 1, c })]
          return cell?.w ?? cell?.v ?? ''
        }

        // Forma de pago: la celda trae la etiqueta "04 - Tarjeta de Crédito",
        // la recortamos al código de dos dígitos.
        const formaLabel = String(txt(C.forma))
        const formaCode = formaLabel.startsWith('04') ? '04'
          : formaLabel.startsWith('03') ? '03'
          : formaLabel.startsWith('02') ? '02'
          : formaLabel.startsWith('01') ? '01' : '04'

        // Campos numéricos. Semántica idéntica al export (export-excel.py):
        //   L=ISR → retencionISR (RESTA), M=ISH/IEPS → ishIeps (SUMA),
        //   N=RETENCIÓN → retenciones − retencionISR (RESTA).
        // La columna TOTAL (O) es la fórmula viva =J+K-L+M-N SIN valor cacheado,
        // así que casi siempre llega vacía → la reconstruimos con esos valores.
        const importe      = Number(val(C.importe)) || 0
        const iva          = Number(val(C.iva))     || 0
        const retencionISR = Number(val(C.isr))     || 0
        const ishIeps      = Number(val(C.ish))     || 0
        const retNoIsr     = Number(val(C.ret))     || 0   // col N = retenciones − ISR
        const totalRaw     = Number(val(C.total))   || 0
        const totalCFDI    = totalRaw || (importe + iva - retencionISR + ishIeps - retNoIsr)
        const montoUSD     = Number(val(C.usd))     || 0

        gastos.push({
          // id is required by React (table key) and the delete handler;
          // tienePDF/pdfFile/xmlFile/hizoMatch keep the row shape identical to
          // parseCFDI-produced rows so every cell renders correctly.
          id: genId(),
          rfc: String(rfc || ''),
          proveedor: String(val(C.prov) || ''),
          // Normaliza tipos legacy renombrados: IT & SW antes era "IT & SW
          // (Software/Sistemas)" — un Excel viejo importado todavía trae el
          // nombre largo. Se reescribe al canónico para evitar discrepancias
          // falsas en cotejo con Saldos.
          tipo: (() => {
            const raw = String(val(C.tipo) || '').trim()
            if (raw === 'IT & SW (Software/Sistemas)') return 'IT & SW'
            return raw
          })(),
          noFactura: String(val(C.factura) || ''),
          fechaFac: parseDate(val(C.fFac)),
          fechaCobro: parseDate(val(C.fCobro)),
          concepto: String(val(C.concepto) || ''),
          importe,
          iva,
          isrTrasladado: ishIeps,
          ishIeps,
          retencionISR,
          retencionIVA: retNoIsr,
          retenciones: retencionISR + retNoIsr,
          totalCFDI,
          formaPago: formaCode,
          banco: String(val(C.banco) || ''),
          propinaPorcentaje: 0,
          montoPropina: 0,
          uuid: crypto.randomUUID(),
          tienePDF: false,
          pdfFile: null,
          xmlFile: null,
          hizoMatch: false,
          validado: false,
          montoUSD,
          tipoCambio: Number(val(C.tc)) || 0,
          moneda: montoUSD > 0 ? 'USD' : 'MXN',
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
  // ── Cotejo con Saldos (solo colaboradores especiales) ──
  // Aplica las pólizas que sí matchean automáticamente; abre el modal de
  // discrepancias para que el usuario decida tipo por tipo.
  const handleSaldosFile = async (e) => {
    const file = e.target.files?.[0]
    if (!file) return
    try {
      const arrayBuffer = await file.arrayBuffer()
      const saldosRows = parseSaldosXLSX(arrayBuffer)
      if (saldosRows.length === 0) {
        showModal({
          type: 'error',
          title: 'Saldos sin datos',
          subtitle: 'No se encontraron filas con Fecha|Tipo|Folio|Factura en ninguna pestaña.',
          primaryLabel: 'Entendido',
          onPrimary: closeModal,
        })
        return
      }
      const result = cotejarConSaldos(saldosRows, lista)
      // Aplica las pólizas y la fecha de cobro de TODOS los matches (la
      // fecha del Saldos es la fecha real de pago del banco, equivalente
      // a dCSV en validarBanco). Tipo NO se cambia aquí; el modal lo
      // resuelve después de que el usuario decida por discrepancia.
      setLista(prev => prev.map(g => {
        const m = result.matched.find(x => x.gastoId === g.id)
        if (!m) return g
        const fechaCobro = saldosFechaToIso(m.saldosRow.fecha) || g.fechaCobro
        // El "banco" es el nombre COMPLETO de la pestaña Saldos donde
        // apareció la fila (e.g. "BBVA MXN Cheques", "Monex USD Cheques",
        // "Kapital MXN Flex"). Esto identifica la cuenta exacta — distingue
        // BBVA Cheques de BBVA Concentradora, MXN de USD, etc. Se muestra
        // en la columna BANCO del Excel exportado.
        const banco = String(m.saldosRow.sheet || '').trim().replace(/\s+/g, ' ')
        return {
          ...g,
          polizaNumero: m.saldosRow.folio || 'N/A',
          fechaCobro,
          banco,
        }
      }))
      setCotejoModal({ result, decisions: {} })
    } catch (err) {
      console.warn('parseSaldosXLSX:', err)
      showModal({
        type: 'error',
        title: 'Error leyendo Saldos',
        subtitle: err.message || 'No se pudo leer el archivo.',
        primaryLabel: 'Entendido',
        onPrimary: closeModal,
      })
    } finally {
      e.target.value = ''
    }
  }

  // ── AGREGAR NÓMINA (solo Alejandro Olivar) ──
  // Sube el archivo de Saldos, busca todas las filas cuyo tipo empiece con
  // "Nómina" (Nómina Adm/Soc/Ser/Ven, etc.) y las inserta como gastos con
  // el mismo shape que los demás para que aparezcan en la tabla y en el
  // export a Excel. Aplica filtro por año (si ya hay gastos) y dedup por
  // clave sintética (folio + persona + fecha) para no duplicar en re-imports.
  const handleNominaFile = async (e) => {
    const file = e.target.files?.[0]
    if (!file) return
    try {
      const arrayBuffer = await file.arrayBuffer()
      const saldosRows = parseSaldosXLSX(arrayBuffer)
      const nominaRows = saldosRows.filter(r =>
        /^n[óo]mina\b/i.test(String(r.tipo || '').trim())
      )
      if (nominaRows.length === 0) {
        showModal({
          type: 'error',
          title: 'Sin nóminas',
          subtitle: 'No se encontraron filas tipo "Nómina" en el archivo.',
          primaryLabel: 'Entendido',
          onPrimary: closeModal,
        })
        return
      }

      // Agrupa nóminas por YYYY-MM para que el usuario elija el mes a importar.
      const MES_NOMBRES = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre']
      const byMonth = new Map()  // ym → { rows, total }
      for (const r of nominaRows) {
        const ym = (saldosFechaToIso(r.fecha) || '').slice(0, 7)
        if (!/^\d{4}-\d{2}$/.test(ym)) continue
        if (!byMonth.has(ym)) byMonth.set(ym, { rows: [], total: 0 })
        const bucket = byMonth.get(ym)
        bucket.rows.push(r)
        bucket.total += Number(r.egreso) || 0
      }

      const months = [...byMonth.entries()]
        .map(([ym, { rows, total }]) => {
          const [y, m] = ym.split('-')
          return {
            ym,
            label: `${MES_NOMBRES[parseInt(m, 10) - 1]} ${y}`,
            count: rows.length,
            total,
            rows,
          }
        })
        .sort((a, b) => b.ym.localeCompare(a.ym))  // más reciente primero

      if (months.length === 0) {
        showModal({
          type: 'error',
          title: 'Sin nóminas con fecha válida',
          subtitle: 'No se pudieron extraer fechas de las nóminas del archivo.',
          primaryLabel: 'Entendido',
          onPrimary: closeModal,
        })
        return
      }

      // Sugiere el mes que matchea con los gastos cargados (o el más reciente
      // del archivo si no hay gastos).
      const gastoYearMonths = new Set()
      lista.forEach(g => {
        const ym = String(g.fechaFac || '').slice(0, 7)
        if (/^\d{4}-\d{2}$/.test(ym)) gastoYearMonths.add(ym)
      })
      const sugerido = months.find(m => gastoYearMonths.has(m.ym))?.ym || months[0].ym

      // Convierte el bucket del mes seleccionado en gastos y los inserta.
      const applyNominasForMonth = (selectedYM) => {
        const bucket = months.find(m => m.ym === selectedYM)
        if (!bucket) return
        // Dedup compound key para nóminas (los gastos no-nómina usan
        // rfc|noFactura como antes). Necesario porque varios empleados
        // comparten el mismo cheque (folio) y noFactura ahora son solo
        // dígitos del folio.
        const nominaDedupKey = (g) => `NOM|${g.rfc || ''}|${g.polizaNumero || ''}|${g.proveedor || ''}|${g.fechaFac || ''}`
        const existingKeys = new Set(
          lista.map(g => g.esNomina ? nominaDedupKey(g) : `${g.rfc}|${g.noFactura}`)
        )
        const newGastos = []
        let dupesSkipped = 0

        for (const r of bucket.rows) {
          const fechaIso = saldosFechaToIso(r.fecha) || ''
          const conceptoStr = String(r.concepto || '').trim()
          // Cleanup del concepto: quita prefijos de beneficio (Nómina/PTU/
          // Aguinaldo/Prima) y sufijos de período. Soporta:
          //   "Nómina Rosalba Olivar 1a Ene 26" → "Rosalba Olivar"
          //   "PTU Rosalba Olivar 25"           → "Rosalba Olivar"
          //   "Aguinaldo Edie Haro 25"          → "Edie Haro"
          const shortName = conceptoStr
            .replace(/^(?:n[óo]mina|ptu|aguinaldo|prima\s+vacacional)\s+/i, '')
            .replace(/\s+\d+a?\s+\w+\s+\d+\s*$/i, '')
            .replace(/\s+\d{2,4}\s*$/, '')
            .trim() || conceptoStr
          // Lookup empleado: si matchea, usa RFC + nombre completo del
          // roster fiscal. Si no, deja short name como proveedor sin RFC.
          const emp = matchEmpleadoByShortName(shortName)
          const proveedor = emp?.nombre || shortName
          const empRfc    = emp?.rfc    || ''
          const folio = String(r.folio || '').trim()
          // El folio del cheque vive en el campo de folio (polizaNumero). En la
          // columna Factura las nóminas van como 'N/A' (no son CFDI cotejables;
          // normFactura quita la diagonal así que sigue omitiéndose en el cotejo).
          const noFactura = 'N/A'

          // Dedup compuesto para evitar colisión cuando varios empleados
          // comparten folio.
          const dedupKey = `NOM|${empRfc}|${folio}|${proveedor}|${fechaIso}`
          if (existingKeys.has(dedupKey)) { dupesSkipped++; continue }
          existingKeys.add(dedupKey)

          const importe = Number(r.egreso) || 0
          newGastos.push({
            id: genId(),
            rfc: empRfc,
            proveedor,
            tipo: String(r.tipo || '').trim(),
            noFactura,
            fechaFac:  fechaIso,
            fechaCobro: fechaIso,
            concepto:  conceptoStr,
            importe,
            iva: 0,
            isrTrasladado: 0,
            retencionISR:  0,
            retencionIVA:  0,
            retenciones:   0,
            totalCFDI:     importe,
            propinaPorcentaje: 0,
            montoPropina:      0,
            formaPago: '03',
            uuid:     crypto.randomUUID(),
            tienePDF: false,
            pdfFile:  null,
            xmlFile:  null,
            hizoMatch: false,
            validado:  false,
            montoUSD:           0,
            montoExtranjero:    0,
            importeUSD:         0,
            ivaUSD:             0,
            retencionesUSD:     0,
            tipoCambio:         0,
            moneda:             'MXN',
            monedaCodigo:       'MXN',
            esMonedaExtranjera: false,
            esTicket:    false,
            esNomina:    true,
            polizaNumero: folio,
            banco:       String(r.sheet || '').trim().replace(/\s+/g, ' '),
            montoFacturado: importe,
            montoPropinaOriginal: 0,
            isNew: true,
          })
        }

        if (newGastos.length === 0) {
          showModal({
            type: 'info',
            title: 'Sin nóminas nuevas',
            subtitle: dupesSkipped > 0
              ? `Las ${dupesSkipped} nóminas de ${bucket.label} ya están cargadas en el reporte.`
              : `No se encontraron nóminas nuevas en ${bucket.label}.`,
            primaryLabel: 'Entendido',
            onPrimary: closeModal,
          })
          return
        }

        setLista(prev => [...prev, ...newGastos])
        const totalNominas = newGastos.reduce((s, g) => s + g.importe, 0)
        showModal({
          type: 'success',
          title: `Nóminas ${bucket.label} agregadas`,
          subtitle: dupesSkipped > 0
            ? `+${newGastos.length} nuevas · ${dupesSkipped} duplicadas omitidas`
            : `${newGastos.length} nómina${newGastos.length === 1 ? '' : 's'} agregada${newGastos.length === 1 ? '' : 's'} al reporte.`,
          stats: [
            { value: `+${newGastos.length}`, label: 'Nóminas', color: '#59D39B' },
            { value: `$${totalNominas.toLocaleString('es-MX', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`, label: 'Total', color: 'rgba(255,255,255,0.85)' },
          ],
          primaryLabel: 'Continuar',
        })
        setTimeout(() => {
          setLista(l => l.map(g => g.isNew ? { ...g, isNew: false } : g))
        }, 1500)
      }

      // Abre el picker de mes
      setNominaPickerModal({
        months,
        selectedYM: sugerido,
        onConfirm: (ym) => {
          setNominaPickerModal(null)
          applyNominasForMonth(ym)
        },
        onCancel: () => setNominaPickerModal(null),
      })
    } catch (err) {
      console.warn('handleNominaFile:', err)
      showModal({
        type: 'error',
        title: 'Error leyendo archivo',
        subtitle: err.message || 'No se pudo leer el archivo de Saldos.',
        primaryLabel: 'Entendido',
        onPrimary: closeModal,
      })
    } finally {
      e.target.value = ''
    }
  }

  // Aplica las decisiones de tipo del modal al state y lo cierra.
  const aplicarDecisionesCotejo = () => {
    if (!cotejoModal) return
    const { result, decisions } = cotejoModal
    const matchedCount    = result.matched.length
    const discrepancias   = result.matched.filter(m => m.tipoDiffers)
    const tiposActualizados = discrepancias.filter(m => decisions[m.gasto.id] === 'saldos').length
    const tiposMantenidos   = discrepancias.length - tiposActualizados
    const colisionesRevisar = result.matched.filter(m => m.facturaColision && !m.colisionResuelta).length
    const sinMatchG       = result.sinMatchGastos.length
    const sinMatchS       = result.sinMatchSaldos.length
    setLista(prev => prev.map(g => {
      const m = result.matched.find(x => x.gastoId === g.id)
      if (!m || !m.tipoDiffers) return g
      const choice = decisions[g.id]
      if (choice === 'saldos') return { ...g, tipo: m.saldosRow.tipo }
      return g
    }))
    setCotejoModal(null)
    showModal({
      type: 'success',
      title: 'Cotejo aplicado',
      subtitle: `${matchedCount} factura${matchedCount === 1 ? '' : 's'} vinculada${matchedCount === 1 ? '' : 's'} al Saldos.`,
      stats: [
        { value: matchedCount,       label: 'Matches',           color: '#59D39B' },
        ...(tiposActualizados > 0 ? [{ value: tiposActualizados, label: 'Tipos actualizados', color: 'rgba(255,255,255,0.85)' }] : []),
        ...(tiposMantenidos   > 0 ? [{ value: tiposMantenidos,   label: 'Tipos mantenidos',   color: 'rgba(255,255,255,0.85)' }] : []),
        ...(colisionesRevisar > 0 ? [{ value: colisionesRevisar, label: 'Folios repetidos a revisar', color: '#FF9F0A' }] : []),
        ...(sinMatchG > 0 ? [{ value: sinMatchG, label: 'Gastos sin match', color: '#FF9F0A' }] : []),
        ...(sinMatchS > 0 ? [{ value: sinMatchS, label: 'Saldos sin match', color: 'rgba(255,255,255,0.65)' }] : []),
      ],
      primaryLabel: 'Continuar',
    })
  }

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
        isrTrasladado:    Number(g.isrTrasladado) || 0,
        ishIeps:          Number(g.ishIeps ?? g.isrTrasladado) || 0,
        retencionISR:     Number(g.retencionISR) || 0,
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
        importeUSD:       Number(g.importeUSD) || 0,
        ivaUSD:           Number(g.ivaUSD) || 0,
        retencionesUSD:   Number(g.retencionesUSD) || 0,
        banco:            g.banco || '',
        montoFacturado:   Number(g.montoFacturado) || 0,
        montoPropinaOriginal: Number(g.montoPropinaOriginal) || 0,
        montoCobrado:     Number(g.montoCobrado) || 0,
        hizoMatch:        !!g.hizoMatch,
        xmlFaltante:      !!g.xmlFaltante,
      }))
      const response = await fetch('/api/export-excel', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ gastos: gastosSlim, colaborador: colaborador?.nombre || '', polizaNumero: POLIZAS_CLARA[colaborador?.nombre] || 'N/A', polizasMap: POLIZAS_CLARA }),
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
  // Estable (deps vacías, usa setModal directo) para preservar la memoización
  // de GastoRow — así abrir un PDF no fuerza el re-render de toda la tabla.
  const openPDF = useCallback(pdfFile => {
    if (!pdfFile) return
    try {
      const url = URL.createObjectURL(pdfFile)
      const win = window.open(url, '_blank')
      if (!win) setModal({
        type: 'warning',
        title: 'No se pudo abrir el PDF',
        subtitle: 'Verifica que el navegador permita ventanas emergentes para este sitio.',
        primaryLabel: 'Entendido',
      })
    } catch (err) {
      setModal({
        type: 'error',
        title: 'Error al abrir el PDF',
        subtitle: err && err.message ? err.message : String(err),
        primaryLabel: 'Entendido',
      })
    }
  }, [])

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
          <h1 className="header-title">Reporte de Gastos SMTO<span className="version-badge">v8.99</span></h1>
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
        <div className="guide-btn-group">
          <button className="guide-btn guide-btn-clara" onClick={() => setShowClaraGuide(true)} title="Manual para descargar el estado de cuenta de Clara">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
              <path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/>
              <polyline points="14 2 14 8 20 8"/>
              <line x1="16" y1="13" x2="8" y2="13"/>
              <line x1="16" y1="17" x2="8" y2="17"/>
            </svg>
            ¿Cómo Descargar tu estado de cuenta CLARA?
          </button>
          <button className="guide-btn" onClick={() => setTourStep(0)} title="Recorrido paso a paso">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
              <circle cx="12" cy="12" r="10"/>
              <path d="M9.09 9a3 3 0 015.83 1c0 2-3 3-3 3"/>
              <line x1="12" y1="17" x2="12.01" y2="17"/>
            </svg>
            ¿Cómo utilizar la app?
          </button>
        </div>
      </div>

      <div className="divider" />

      {/* ─── BARRA DE ACCIONES ─── */}
      <div className="action-bar">
        <div className="action-group">
          <PremiumButton title="Manual"         icon="＋"  variant="ghost"     onClick={agregarManual} />
          <PremiumButton
            title={importSuccess ? 'Importado' : 'Importar reporte'}
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
            id="tour-carpeta"
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
          <PremiumButton title="Cargar Foto"    icon="📸" variant="secondary" onClick={() => setShowPhotoChoice(true)} />
          <PremiumButton id="tour-banco" title="Validar Banco"  icon="🏦" variant="secondary" onClick={() => bancoRef.current?.click()} />
          {COLABORADORES_ESPECIALES.includes(colaborador?.nombre) && (
            <PremiumButton
              title="Cotejar con Saldos"
              icon="🔍"
              variant="secondary"
              isDisabled={!lista.length}
              onClick={() => saldosRef.current?.click()}
            />
          )}
          {colaborador?.nombre === 'Alejandro Olivar' && (
            <PremiumButton
              title="Agregar Nómina"
              icon="👥"
              variant="secondary"
              onClick={() => nominaRef.current?.click()}
            />
          )}
          <input
            ref={saldosRef}
            type="file"
            accept=".xlsx,.xls"
            style={{ display: 'none' }}
            onChange={handleSaldosFile}
          />
          <input
            ref={nominaRef}
            type="file"
            accept=".xlsx,.xls"
            style={{ display: 'none' }}
            onChange={handleNominaFile}
          />
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
          <PremiumButton title="Exportar a Excel" icon="📊" variant="primary" isDisabled={!lista.length} onClick={exportarExcel} />
          <PremiumButton id="tour-zip" title="Exportar a ZIP"  icon="📦" variant="export"  isDisabled={!lista.length} onClick={exportar} />
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
        {/* Divisas — con UNA moneda extranjera muestra su total con su símbolo
            ("Total CNY ¥249.11"); con varias, cuántas hay (desglose en el
            tooltip) porque sumarlas entre sí no significa nada. Espejo del
            comportamiento aprobado de la tarjeta del Excel. */}
        {(() => {
          const divisas = Object.entries(metrics.porMoneda)
          const detalle = divisas.map(([c, v]) => `${c} ${fmtDivisa(c, v)}`).join('  ·  ')
          return (
            <div className="metric-card" style={{ '--accent': '#59D39B' }} title={detalle || undefined}>
              <div className="metric-label">
                {divisas.length === 1 ? `Total ${divisas[0][0]}` : divisas.length > 1 ? 'Moneda extranjera' : 'Total USD'}
              </div>
              <div className="metric-value">
                {divisas.length === 0 ? fmtMoney(0)
                  : divisas.length === 1 ? fmtDivisa(divisas[0][0], divisas[0][1])
                  : `${divisas.length} divisas`}
              </div>
            </div>
          )
        })()}
        <div className="metric-card" style={{ '--accent': '#30D158' }}>
          <div className="metric-label">Registros</div>
          <div className="metric-value">{metrics.count}</div>
        </div>
        <div className="metric-card" style={{ '--accent': '#FF453A' }}>
          <div className="metric-label">Por Corroborar</div>
          <div className="metric-value">{metrics.sinCobrar}</div>
        </div>
      </div>

      {/* ─── BUSCADOR UNIVERSAL ─── */}
      {lista.length > 0 && (
        <div className="search-bar">
          <svg className="search-icon" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
            <circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/>
          </svg>
          <input
            ref={searchInputRef}
            type="text"
            className="search-input"
            value={busqueda}
            onChange={e => setBusqueda(e.target.value)}
            placeholder="Buscar por proveedor, factura, fecha, RFC, concepto, póliza..."
          />
          {busqueda && (
            <span className="search-count">
              {filteredLista.length} de {lista.length}
            </span>
          )}
          {busqueda && (
            <button
              className="search-clear"
              onClick={() => { setBusqueda(''); searchInputRef.current?.focus() }}
              title="Limpiar búsqueda"
            >
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
            </button>
          )}
          <span className="search-kbd">⌘K</span>
        </div>
      )}

      {/* Inputs de archivo ocultos */}
      <input
        ref={el => { folderRef.current = el; if (el) { el.setAttribute('webkitdirectory',''); el.setAttribute('directory','') } }}
        type="file" multiple style={{ display: 'none' }}
        onChange={cargar}
      />
      <input
        ref={bancoRef}
        type="file" accept=".csv,.txt,.tsv,.xlsx,.xls" style={{ display: 'none' }}
        onChange={validarBanco}
      />
      {/* Subir foto/PDF existente — sin capture, abre galería/archivos */}
      <input
        ref={photoRef}
        type="file"
        accept="image/*,.heic,.heif,application/pdf"
        multiple
        style={{ display: 'none' }}
        onChange={cargarFoto}
      />
      {/* Tomar foto — capture="environment" abre la cámara trasera en móvil */}
      <input
        ref={cameraRef}
        type="file"
        accept="image/*,.heic,.heif"
        capture="environment"
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
            <div className="loading-text">
              {ocrProgress
                ? `Procesando factura ${ocrProgress.current} de ${ocrProgress.total}…`
                : 'Procesando facturas XML…'}
            </div>
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
                {COLUMNS.map((col, idx) => {
                  // Oculta columnas specialOnly (banco) para no-especiales.
                  if (col.specialOnly && !esColaboradorEspecial) return null
                  return (
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
                  )
                })}
                {/* Eliminar — sticky-right action column, no label, fixed 40px width */}
                <th className="th-delete" style={{ width: 40 }} />
              </tr>
            </thead>
            <tbody>
              {sortedLista.map(g => (
                <GastoRow
                  key={g.id}
                  g={g}
                  update={update}
                  remove={removeGasto}
                  openPDF={openPDF}
                  tiposList={tiposList}
                  isSpecial={esColaboradorEspecial}
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
          body={modal.body}
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

      {/* ─── MODAL COTEJO CON SALDOS (solo especiales) ─── */}
      {cotejoModal && (() => {
        const { result, decisions } = cotejoModal
        const discrepancias = result.matched.filter(m => m.tipoDiffers)
        const colisiones = result.matched.filter(m => m.facturaColision)
        const sinMatchG = result.sinMatchGastos
        const sinMatchS = result.sinMatchSaldos
        const setDecision = (gastoId, choice) => {
          setCotejoModal(prev => ({ ...prev, decisions: { ...prev.decisions, [gastoId]: choice } }))
        }
        const fmtMoney = (n) => Number(n || 0).toLocaleString('es-MX', { minimumFractionDigits: 2, maximumFractionDigits: 2 })
        const fmtFecha = (d) => {
          if (!d) return ''
          if (d instanceof Date) return d.toLocaleDateString('es-MX')
          if (typeof d === 'number') {
            const dt = new Date(Date.UTC(1899, 11, 30) + d * 86400000)
            return dt.toLocaleDateString('es-MX')
          }
          return String(d).slice(0, 10)
        }
        return (
          <div className="premium-overlay" onClick={() => setCotejoModal(null)}>
            <div className="premium-modal" onClick={e => e.stopPropagation()}
              style={{ maxWidth: 880, width: '92vw', maxHeight: '88vh', display: 'flex', flexDirection: 'column' }}>
              <div className="premium-icon-wrap" style={{ background: 'rgba(89,211,155,0.12)', boxShadow: '0 0 40px #59D39B30' }}>
                <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="#59D39B" strokeWidth="2.5">
                  <circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/>
                </svg>
              </div>
              <div className="premium-title">Cotejo con Saldos</div>
              <div className="premium-subtitle">
                {result.matched.length} pólizas aplicadas · {discrepancias.length} discrepancia(s) de tipo{colisiones.length > 0 ? ` · ${colisiones.length} factura(s) con número repetido` : ''} · {sinMatchG.length} gasto(s) sin match · {sinMatchS.length} fila(s) Saldos sobrantes
              </div>

              <div style={{ overflowY: 'auto', flex: 1, marginTop: 16, padding: '0 4px' }}>
                {discrepancias.length > 0 && (
                  <section style={{ marginBottom: 24 }}>
                    <div style={{ fontSize: 13, fontWeight: 700, color: '#f59e0b', marginBottom: 8, textTransform: 'uppercase', letterSpacing: 0.5 }}>
                      Discrepancias de tipo
                    </div>
                    <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
                      {discrepancias.map(m => {
                        const choice = decisions[m.gastoId]
                        return (
                          <div key={m.gastoId} style={{ background: 'rgba(245,158,11,0.06)', border: '1px solid rgba(245,158,11,0.3)', borderRadius: 10, padding: 12 }}>
                            <div style={{ fontSize: 13, fontWeight: 700, color: '#e2e8f0', marginBottom: 4 }}>
                              {m.gasto.proveedor} · {m.gasto.noFactura}
                            </div>
                            <div style={{ fontSize: 11, color: '#94a3b8', marginBottom: 8 }}>
                              {fmtFecha(m.gasto.fechaFac)} · ${fmtMoney(m.gasto.totalCFDI)}
                            </div>
                            <div style={{ display: 'flex', gap: 8, alignItems: 'center', flexWrap: 'wrap' }}>
                              <button
                                onClick={() => setDecision(m.gastoId, 'mantener')}
                                style={{
                                  flex: 1, minWidth: 200, padding: '8px 12px', borderRadius: 8,
                                  border: choice === 'mantener' ? '2px solid #59D39B' : '1px solid rgba(148,163,184,0.3)',
                                  background: choice === 'mantener' ? 'rgba(89,211,155,0.15)' : 'transparent',
                                  color: '#e2e8f0', cursor: 'pointer', fontSize: 12, fontWeight: 600, textAlign: 'left',
                                }}>
                                <div style={{ fontSize: 10, color: '#94a3b8', marginBottom: 2 }}>MANTENER (app)</div>
                                {m.gasto.tipo}
                              </button>
                              <button
                                onClick={() => setDecision(m.gastoId, 'saldos')}
                                style={{
                                  flex: 1, minWidth: 200, padding: '8px 12px', borderRadius: 8,
                                  border: choice === 'saldos' ? '2px solid #59D39B' : '1px solid rgba(148,163,184,0.3)',
                                  background: choice === 'saldos' ? 'rgba(89,211,155,0.15)' : 'transparent',
                                  color: '#e2e8f0', cursor: 'pointer', fontSize: 12, fontWeight: 600, textAlign: 'left',
                                }}>
                                <div style={{ fontSize: 10, color: '#94a3b8', marginBottom: 2 }}>USAR SALDOS</div>
                                {m.saldosRow.tipo}
                              </button>
                            </div>
                          </div>
                        )
                      })}
                    </div>
                  </section>
                )}

                {colisiones.length > 0 && (
                  <section style={{ marginBottom: 24 }}>
                    <div style={{ fontSize: 13, fontWeight: 700, color: '#38bdf8', marginBottom: 4, textTransform: 'uppercase', letterSpacing: 0.5 }}>
                      Facturas con número repetido
                    </div>
                    <div style={{ fontSize: 11, color: '#94a3b8', marginBottom: 8 }}>
                      Más de un renglón del Saldos comparte este número de factura. Se vinculó el de mayor coincidencia (monto, proveedor, moneda y fecha). Verifica los marcados con ⚠.
                    </div>
                    <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
                      {colisiones.map(m => (
                        <div key={m.gastoId} style={{
                          background: m.colisionResuelta ? 'rgba(56,189,248,0.06)' : 'rgba(245,158,11,0.08)',
                          border: `1px solid ${m.colisionResuelta ? 'rgba(56,189,248,0.3)' : 'rgba(245,158,11,0.4)'}`,
                          borderRadius: 8, padding: 10,
                        }}>
                          <div style={{ fontSize: 12, fontWeight: 700, color: '#e2e8f0', marginBottom: 2 }}>
                            {m.colisionResuelta ? '✓' : '⚠'} Factura {m.gasto.noFactura}
                          </div>
                          <div style={{ fontSize: 11, color: '#94a3b8' }}>
                            Gasto: {m.gasto.proveedor} · {fmtFecha(m.gasto.fechaFac)} · ${fmtMoney(m.gasto.totalCFDI)}
                          </div>
                          <div style={{ fontSize: 11, color: '#94a3b8' }}>
                            Vinculado a: {m.saldosRow.sheet} · {m.saldosRow.tipo || '—'} · {fmtFecha(m.saldosRow.fecha)} · ${fmtMoney(m.saldosRow.egreso)}
                          </div>
                          {!m.colisionResuelta && (
                            <div style={{ fontSize: 10, color: '#f59e0b', marginTop: 4 }}>
                              Sin coincidencia clara de monto ni proveedor — revisa que sea la factura correcta.
                            </div>
                          )}
                        </div>
                      ))}
                    </div>
                  </section>
                )}

                {sinMatchG.length > 0 && (
                  <section style={{ marginBottom: 24 }}>
                    <div style={{ fontSize: 13, fontWeight: 700, color: '#ef4444', marginBottom: 8, textTransform: 'uppercase', letterSpacing: 0.5 }}>
                      Gastos sin match en Saldos
                    </div>
                    <div style={{ display: 'flex', flexDirection: 'column', gap: 4, fontSize: 12 }}>
                      {sinMatchG.map(g => (
                        <div key={g.id} style={{ background: 'rgba(239,68,68,0.06)', borderLeft: '3px solid #ef4444', padding: '6px 10px', borderRadius: 4 }}>
                          <span style={{ color: '#e2e8f0', fontWeight: 600 }}>{g.proveedor}</span>
                          <span style={{ color: '#94a3b8' }}> · {g.noFactura} · ${fmtMoney(g.totalCFDI)}</span>
                        </div>
                      ))}
                    </div>
                  </section>
                )}

                {sinMatchS.length > 0 && (
                  <section style={{ marginBottom: 24 }}>
                    <div style={{ fontSize: 13, fontWeight: 700, color: '#94a3b8', marginBottom: 8, textTransform: 'uppercase', letterSpacing: 0.5 }}>
                      Filas del Saldos sin match (informativo)
                    </div>
                    <div style={{ display: 'flex', flexDirection: 'column', gap: 4, fontSize: 12, maxHeight: 180, overflowY: 'auto' }}>
                      {sinMatchS.map((s, i) => (
                        <div key={i} style={{ background: 'rgba(148,163,184,0.06)', borderLeft: '3px solid #64748b', padding: '6px 10px', borderRadius: 4 }}>
                          <span style={{ color: '#e2e8f0' }}>{s.sheet}</span>
                          <span style={{ color: '#94a3b8' }}> · {fmtFecha(s.fecha)} · {s.tipo || '—'} · {s.factura} · ${fmtMoney(s.egreso)}</span>
                        </div>
                      ))}
                    </div>
                  </section>
                )}
              </div>

              <div className="premium-actions" style={{ marginTop: 16 }}>
                <button className="premium-btn-secondary" onClick={() => setCotejoModal(null)}>Cancelar</button>
                <button className="premium-btn-primary" onClick={aplicarDecisionesCotejo}>Aplicar y cerrar</button>
              </div>
            </div>
          </div>
        )
      })()}

      {/* ─── MODAL DUPLICADOS ─── */}
      {duplicadosModal && (
        <div className="premium-overlay" onClick={() => setDuplicadosModal(null)}>
          <div className="premium-modal" style={{ maxWidth: 540 }} onClick={e => e.stopPropagation()}>
            <div className="premium-icon-wrap" style={{ background: 'rgba(255,159,10,0.15)', border: '1px solid rgba(255,159,10,0.25)' }}>
              <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="#FF9F0A" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                <path d="M10.29 3.86L1.82 18a2 2 0 001.71 3h16.94a2 2 0 001.71-3L13.71 3.86a2 2 0 00-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/>
              </svg>
            </div>
            <h2 className="premium-title">Duplicados detectados</h2>
            <p className="premium-subtitle">
              {duplicadosModal.items.length > 0
                ? `${duplicadosModal.items.length} factura${duplicadosModal.items.length === 1 ? '' : 's'} ya exist${duplicadosModal.items.length === 1 ? 'e' : 'en'} en el reporte. Elimina las que no quieras reemplazar.`
                : 'Sin duplicados pendientes. Puedes continuar.'}
            </p>
            {duplicadosModal.items.length > 0 && (
              <div className="dup-list">
                {duplicadosModal.items.map(g => (
                  <div key={g.uuid} className="dup-item">
                    <div className="dup-item-info">
                      <span className="dup-item-proveedor">{g.proveedor || '—'}</span>
                      <span className="dup-item-meta">
                        {g.noFactura && <> · {g.noFactura}</>}
                        {g.fechaFac && <> · {g.fechaFac}</>}
                        {' · '}${Number(g.totalCFDI || 0).toLocaleString('es-MX', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                      </span>
                    </div>
                    <button
                      className="dup-item-remove"
                      title="Eliminar este duplicado"
                      onClick={() => setDuplicadosModal(prev => ({
                        ...prev,
                        items:    prev.items.filter(x => x.uuid !== g.uuid),
                        incoming: prev.incoming.filter(x => x.uuid !== g.uuid),
                      }))}
                    >
                      <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
                    </button>
                  </div>
                ))}
              </div>
            )}
            <div className="premium-actions" style={{ marginTop: 20 }}>
              <button className="premium-btn-secondary" onClick={() => setDuplicadosModal(null)}>Cancelar</button>
              <button className="premium-btn-primary" onClick={() => duplicadosModal.onConfirm(duplicadosModal.incoming)}>
                Continuar
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ─── MODAL SELECCIÓN OCR (archivos sin XML en la carpeta) ─── */}
      {ocrSelectionModal && (() => {
        const { items, onConfirm, onCancel } = ocrSelectionModal
        const selectedCount = items.filter(i => i.selected).length
        const setAllSelected = (val) => setOcrSelectionModal(prev => ({
          ...prev,
          items: prev.items.map(i => ({ ...i, selected: val })),
        }))
        const toggle = (idx) => setOcrSelectionModal(prev => ({
          ...prev,
          items: prev.items.map((i, k) => k === idx ? { ...i, selected: !i.selected } : i),
        }))
        return (
          <div className="premium-overlay" onClick={onCancel}>
            <div className="premium-modal" style={{ maxWidth: 580 }} onClick={e => e.stopPropagation()}>
              <div className="premium-icon-wrap" style={{ background: 'rgba(245,158,11,0.15)', border: '1px solid rgba(245,158,11,0.25)' }}>
                <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="#F59E0B" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                  <circle cx="12" cy="12" r="9"/><path d="M9 12l2 2 4-4"/>
                </svg>
              </div>
              <h2 className="premium-title">Archivos sin XML detectados</h2>
              <p className="premium-subtitle">
                {items.length} archivo{items.length === 1 ? '' : 's'} en la carpeta no {items.length === 1 ? 'tiene' : 'tienen'} un XML pareja.
                Selecciona cuáles procesar con OCR (OpenAI, ~{fmtUsd(OCR_COST_USD)} USD/factura).
              </p>
              <div style={{ display: 'flex', gap: 10, justifyContent: 'center', margin: '12px 0 4px', fontSize: 12 }}>
                <button
                  onClick={() => setAllSelected(true)}
                  style={{ background: 'transparent', border: '1px solid rgba(255,255,255,0.15)', color: 'rgba(255,255,255,0.75)', padding: '4px 12px', borderRadius: 6, cursor: 'pointer' }}
                >Seleccionar todos</button>
                <button
                  onClick={() => setAllSelected(false)}
                  style={{ background: 'transparent', border: '1px solid rgba(255,255,255,0.15)', color: 'rgba(255,255,255,0.75)', padding: '4px 12px', borderRadius: 6, cursor: 'pointer' }}
                >Ninguno</button>
              </div>
              <div className="ocr-list">
                {items.map((it, idx) => {
                  const isImage = (it.file.type || '').startsWith('image/') ||
                    /\.(jpe?g|png|webp|heic|heif|bmp|gif)$/i.test(it.file.name)
                  return (
                    <label key={idx} className={`ocr-item${it.selected ? ' selected' : ''}`}>
                      <input type="checkbox" checked={it.selected} onChange={() => toggle(idx)} />
                      <span className={`ocr-item-badge ${isImage ? 'img' : 'pdf'}`}>{isImage ? 'IMG' : 'PDF'}</span>
                      <span className="ocr-item-name">{it.file.name}</span>
                    </label>
                  )
                })}
              </div>
              <div style={{ marginTop: 12, fontSize: 12, color: 'rgba(255,255,255,0.55)' }}>
                {selectedCount === 0
                  ? 'Sin archivos seleccionados — se cargarán solo los XMLs'
                  : `${selectedCount} seleccionado${selectedCount === 1 ? '' : 's'} · ≈ ${fmtUsd(selectedCount * OCR_COST_USD)} USD`}
              </div>
              <div className="premium-actions" style={{ marginTop: 18 }}>
                <button className="premium-btn-secondary" onClick={onCancel}>Cancelar</button>
                <button className="premium-btn-primary" onClick={() => onConfirm(items.filter(i => i.selected).map(i => i.file))}>
                  {selectedCount === 0 ? 'Continuar sin OCR' : `Procesar ${selectedCount}`}
                </button>
              </div>
            </div>
          </div>
        )
      })()}

      {/* ─── MODAL SELECCIÓN DE MES PARA NÓMINAS ─── */}
      {nominaPickerModal && (() => {
        const { months, selectedYM, onConfirm, onCancel } = nominaPickerModal
        const pick = (ym) => setNominaPickerModal(prev => ({ ...prev, selectedYM: ym }))
        return (
          <div className="premium-overlay" onClick={onCancel}>
            <div className="premium-modal" style={{ maxWidth: 480 }} onClick={e => e.stopPropagation()}>
              <div className="premium-icon-wrap" style={{ background: 'rgba(89,211,155,0.15)', border: '1px solid rgba(89,211,155,0.25)' }}>
                <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="#59D39B" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                  <rect x="3" y="4" width="18" height="18" rx="2" ry="2"/>
                  <line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/>
                </svg>
              </div>
              <h2 className="premium-title">Selecciona el mes</h2>
              <p className="premium-subtitle">
                Elige de qué mes quieres importar las nóminas. Se encontraron {months.length} {months.length === 1 ? 'mes' : 'meses'} con nóminas en el archivo.
              </p>
              <div className="month-list">
                {months.map(m => (
                  <label key={m.ym} className={`month-item${selectedYM === m.ym ? ' selected' : ''}`}>
                    <input
                      type="radio"
                      name="nomina-month"
                      checked={selectedYM === m.ym}
                      onChange={() => pick(m.ym)}
                    />
                    <div className="month-info">
                      <div className="month-name">{m.label}</div>
                      <div className="month-stats">
                        {m.count} nómina{m.count === 1 ? '' : 's'} · ${m.total.toLocaleString('es-MX', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                      </div>
                    </div>
                  </label>
                ))}
              </div>
              <div className="premium-actions" style={{ marginTop: 18 }}>
                <button className="premium-btn-secondary" onClick={onCancel}>Cancelar</button>
                <button className="premium-btn-primary" disabled={!selectedYM} onClick={() => onConfirm(selectedYM)}>
                  Agregar nóminas
                </button>
              </div>
            </div>
          </div>
        )
      })()}

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

      {/* ─── ELEGIR: TOMAR FOTO / SUBIR FOTO ─── */}
      <AnimatePresence>
        {showPhotoChoice && (
          <motion.div
            className="premium-overlay"
            onClick={() => setShowPhotoChoice(false)}
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
          >
            <motion.div
              className="premium-modal photo-choice-modal"
              onClick={e => e.stopPropagation()}
              initial={{ opacity: 0, y: 20, scale: 0.96 }}
              animate={{ opacity: 1, y: 0, scale: 1 }}
              exit={{ opacity: 0, y: 12, scale: 0.98 }}
              transition={{ type: 'spring', stiffness: 300, damping: 26 }}
            >
              <button
                className="cm-close"
                aria-label="Cerrar"
                onClick={() => setShowPhotoChoice(false)}
              >
                <X size={16} />
              </button>
              <h2 className="premium-title">Cargar foto</h2>
              <p className="premium-subtitle">Elige cómo quieres agregar el comprobante.</p>

              <div className="photo-choice-grid">
                <button
                  className="photo-choice-btn"
                  onClick={() => { setShowPhotoChoice(false); cameraRef.current?.click() }}
                >
                  <span className="photo-choice-ic">
                    <svg width="26" height="26" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round">
                      <path d="M23 19a2 2 0 0 1-2 2H3a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h4l2-3h6l2 3h4a2 2 0 0 1 2 2z"/>
                      <circle cx="12" cy="13" r="4"/>
                    </svg>
                  </span>
                  <span className="photo-choice-t">Tomar foto</span>
                  <span className="photo-choice-s">Usa la cámara</span>
                </button>

                <button
                  className="photo-choice-btn"
                  onClick={() => { setShowPhotoChoice(false); photoRef.current?.click() }}
                >
                  <span className="photo-choice-ic">
                    <svg width="26" height="26" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round">
                      <rect x="3" y="3" width="18" height="18" rx="2" ry="2"/>
                      <circle cx="8.5" cy="8.5" r="1.5"/>
                      <path d="M21 15l-5-5L5 21"/>
                    </svg>
                  </span>
                  <span className="photo-choice-t">Subir foto</span>
                  <span className="photo-choice-s">Galería o archivos</span>
                </button>
              </div>
            </motion.div>
          </motion.div>
        )}
      </AnimatePresence>

      {/* ─── OCR LOADING OVERLAY — animación de análisis ─── */}
      {ocrLoading && (
        <div className="cm-overlay" style={{ pointerEvents: 'auto' }}>
          <div className="ocr-scan">
            <div className="ocr-scan-frame">
              <span className="ocr-scan-corner tl" />
              <span className="ocr-scan-corner tr" />
              <span className="ocr-scan-corner bl" />
              <span className="ocr-scan-corner br" />
              <svg className="ocr-scan-doc" width="44" height="44" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round">
                <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                <polyline points="14 2 14 8 20 8"/>
                <line x1="8" y1="13" x2="16" y2="13"/>
                <line x1="8" y1="17" x2="13" y2="17"/>
              </svg>
              <span className="ocr-scan-line" />
            </div>
            <div className="ocr-scan-text">
              {ocrProgress && ocrProgress.total > 1
                ? `Analizando ${ocrProgress.current} de ${ocrProgress.total}…`
                : 'Analizando foto…'}
            </div>
            <div className="ocr-scan-sub">Extrayendo datos con OCR</div>
          </div>
        </div>
      )}

      {/* ─── RECORRIDO GUIADO "Cómo utilizar" ─── */}
      <GuideTour step={tourStep} setStep={setTourStep} />
      <ClaraGuideModal open={showClaraGuide} onClose={() => setShowClaraGuide(false)} />

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
