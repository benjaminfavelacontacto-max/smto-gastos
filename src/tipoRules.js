// src/tipoRules.js
// Reglas predictivas de Tipo de gasto — v5
// 18,846 transacciones históricas SMTo + razones sociales fiscales mexicanas + conceptos CFDI
// Keywords largos primero (greedy specificity). Sin keywords ≤3 chars ambiguos.

export const TIPO_RULES = [
  // ── Proveedores SMTO con asignación fija (lista 2026-05) ────────────────
  // Estas reglas viven ARRIBA de las genéricas para ganar prioridad. P.ej.
  // FEDEX abajo dice "Envíos" pero aquí lo forzamos a "Aduana" por convenio
  // interno SMTO.
  ['BISMAK',                   'Consultoría'],
  ['BISMAC',                   'Consultoría'],
  ['C Y C ',                   'Consultoría'],
  ['C Y C,',                   'Consultoría'],
  ['JUAN M ROB',               'Consultoría'],
  ['JUAN MANUEL ROB',          'Consultoría'],
  ['JUAN MIGUEL ROB',          'Consultoría'],
  ['ROBLES VARGAS',            'Consultoría'],
  ['KLS ',                     'Consultoría'],    // KLS Consultoría Aduanal
  ['CORREDURIA PUBLICA',       'Consultoría'],    // Corporativo Jurídico y Correduría
  ['DIOCELINE',                'Curso'],
  ['MARIANA MORALES',          'Marketing'],
  ['OMEGA EDGE',               'COGS'],
  ['ZUÑIGA',                   'Contador'],
  ['ZUNIGA',                   'Contador'],
  ['ZÚÑIGA',                   'Contador'],
  // El SAT registró el Ñ del despacho como '&' literal: 'ZU&IGA RODRIGUEZ Y CIA'.
  // Agregamos la variante exacta para que la auto-detección funcione tal cual
  // viene en el CFDI sin tocar el XML.
  ['ZU&IGA',                   'Contador'],
  // Taxista persona física — sin este override, el 'ADOLF' genérico de abajo
  // (pensado para gasolineras) lo atrapa como 'Gasolina' antes de llegar a la
  // regla de persona física + TARIFA que lo clasificaría bien como Taxi.
  ['ADOLFO REYNA PEÑA',        'Taxi'],
  ['DHL ',                     'Envíos'],
  ['DHL,',                     'Envíos'],
  ['DHL EXPRESS',              'Envíos'],
  ['DHL DE MEXICO',            'Envíos'],
  ['DHL MEXICO',               'Envíos'],
  ['AEROVIAS EMPRESA DE CARGO','Envíos'],          // override de AEROVIAS→Avión genérico
  ['FEDERAL EXPRESS',          'Envíos'],
  ['FEDEX',                    'Aduana'],          // override del FEDEX→Envíos genérico
  ['INTEGRADORA',              'Aduana'],
  ['DESP CANO',                'Aduana'],
  ['DESPACHO CANO',            'Aduana'],
  ['IBG ',                     'Comercio Exterior'],
  ['KLF ',                     'Comercio Exterior'],
  ['MBGE',                     'IT & SW'],
  ['SFERP',                    'IT & SW'],
  ['TOTALPLAY',                'IT & SW'],
  ['TOTAL PLAY',               'IT & SW'],
  ['PROVISOR',                 'COGS'],
  ['TEKTRONIX',                'COGS'],
  ['ROHDE',                    'COGS'],
  ['INSTEK',                   'COGS'],    // GW Instek / Instek America Corp (equipo de medición)
  ['GASNGO',                   'Gasolina'],
  ['ILIANA ALV',               'Renta Oficina'],
  ['ILIANA ALVAREZ',           'Renta Oficina'],
  ['LA VENTOLERA',             'Renta Oficina'],
  ['TOMAS VALLES',             'Renta Oficina'],
  ['ITESO',                    'Renta Oficina'],
  ['FIAT ',                    'Renta Oficina'],   // override de FIAT (marca de auto)
  ['VANRENTA',                 'Automóvil'],
  ['SALUD DIGNA',              'Permiso'],       // facturas de salud (RFC SDI121109B14) → siempre Permiso por convenio SMTO
  ['OPERADORA OMX',            'Papelería'],     // Office Max (razón social 'Operadora OMX', RFC OOM960429832). DEBE ir antes de 'OPERADORA OM'→Hotel, que hace prefix-match sobre esta razón social.
  ['CERVECERIA REGIOMONTANA',              'Consumo Viáticos'],  // Sierra Madre MTY (cargo banco: VELPAY *SIERRA MADRE). CFDI concepto = "CONSUMO ALIMENTOS".
  ['OPERADORA INDUSTRIAL Y COMERCIAL APODACA', 'Consumo Viáticos'],  // Restaurante aeropuerto MTY (cargo banco: THUNDER*PAPALOTE AEROP APODACA). CFDI concepto = "CONSUMO DE ALIMENTOS".
  ['LAS NUEVAS DELICIAS GASTRONOMICAS', 'Consumo Viáticos'],  // DEBE ir antes de 'GAS'→Gasolina: "GASTRONOMICAS" contiene "GAS" como substring y el genérico la atrapaba mal.
  // ── Fin proveedores SMTO ────────────────────────────────────────────────

  // ── Equipo de cómputo / PC ──────────────────────────────────────────────
  ['EQUIPO DE COMPUTO',    'PC (Equipo)'],
  ['EQUIPO DE CÓMPUTO',   'PC (Equipo)'],
  ['EQUIPO DE ESCRITORIO','PC (Equipo)'],
  ['COMPUTADORA PORTATIL','PC (Equipo)'],
  ['COMPUTADORA PORTÁTIL','PC (Equipo)'],
  ['COMPUTADORA PERSONAL', 'PC (Equipo)'],
  ['COMPUTADORA',          'PC (Equipo)'],
  ['ALIENWARE',            'PC (Equipo)'],
  ['THINKPAD',             'PC (Equipo)'],
  ['NOTEBOOK',             'PC (Equipo)'],
  ['SERVIDOR',             'PC (Equipo)'],
  ['IMPRESORA',            'PC (Equipo)'],
  ['MACBOOK',              'PC (Equipo)'],
  ['LAPTOP',               'PC (Equipo)'],
  ['TABLET',               'PC (Equipo)'],
  // ── Fin equipo de cómputo ───────────────────────────────────────────────
  ['COMERCIALIZADORA DE COMBUSTIBLES', 'Gasolina'],
  ['COMISION FEDERAL DE ELECTRICIDAD', 'Renta Oficina'],
  ['OPERADORA DE ESTACIONAMIENTOS', 'Estacionamiento'],
  ['CONCESIONARIA DE AUTOPISTA', 'Casetas'],
  ['MATERIALES DE CONSTRUCCION', 'Herramientas'],
  ['TARIFA SERVICIO TRANSPORTE', 'Taxi'],
  ['ARRENDADORA DE VEHICULOS', 'Renta Auto'],
  ['OPERADORA DE FRANQUICIAS', 'Consumo'],
  ['TRANSPORTADORA EJECUTIVA', 'Taxi'],
  ['ADMINISTRADORA HOTELERA', 'Hotel'],
  ['FACTURACION ELECTRONICA', 'IT & SW'],
  ['MENSAJERIA Y PAQUETERIA', 'Envíos'],
  ['REFACCIONES Y SERVICIOS', 'Manto Auto'],
  ['SERVICIO DE COMBUSTIBLE', 'Gasolina'],
  ['SERVICIOS DE TRANSPORTE', 'Taxi'],
  ['AEROENLACES NACIONALES', 'Avión'],
  ['INVERSIONISTA HOTELERO', 'Hotel'],
  ['SERVICIO DE MENSAJERIA', 'Envíos'],
  ['SERVICIO DE TRANSPORTE', 'Taxi'],
  ['SERVICIOS DE ALIMENTOS', 'Consumo'],
  ['SISTEMA INTERMUNICIPAL', 'Renta Oficina'],
  ['TRANSPORTES TURISTICOS', 'Taxi'],
  ['AMAZON MX MARKETPLACE', 'Marketing'],
  ['REST DENNYS CHIHUAHUA', 'Consumo Viáticos'],
  ['SORIANA981TLAQUEPAQUE', 'Consumo'],
  ['COMPANIA DE AVIACION', 'Avión'],
  ['COMPAÑIA DE AVIACION', 'Avión'],
  ['ESTACION DE SERVICIO', 'Gasolina'],
  ['MERCADOPAGO AEROCHAR', 'Envíos'],
  ['OPERADORA DE HOTELES', 'Hotel'],
  ['SERVICIOS LOGISTICOS', 'Envíos'],
  ['STARBUCKS SAN MARCOS', 'Consumo'],
  ['TRANSPORTE EJECUTIVO', 'Taxi'],
  ['ALIMENTOS Y BEBIDAS', 'Consumo'],
  ['AT&T COMUNICACIONES', 'Celular'],
  ['BURGER KING SENDERO', 'Consumo'],
  ['CONCESIONARIA VUELA', 'Avión'],
  ['HABITACION SENCILLA', 'Hotel'],
  ['LLANTAS Y SERVICIOS', 'Manto Auto'],
  ['OXXO GASCORDILLERAS', 'Gasolina'],
  ['OXXO GASTEQUEPEXPAN', 'Gasolina'],
  ['SERVICIO AUTOMOTRIZ', 'Manto Auto'],
  ['STARBUCKS SAN PEDRO', 'Gastos Rep'],
  ['SUSCRIPCION MENSUAL', 'IT & SW'],
  ['TRASLADO AEROPUERTO', 'Taxi'],
  ['AEROVIAS DE MEXICO', 'Avión'],
  ['COSTCO GUADALAJARA', 'Consumo'],
  ['FIESTA INN NOGALES', 'Consumo'],
  ['NOCHE DE HOSPEDAJE', 'Hotel'],
  ['OPERADORA HOTELERA', 'Hotel'],
  ['PROMOTORA HOTELERA', 'Hotel'],
  ['RENTA DE CAMIONETA', 'Renta Auto'],
  ['TRANSPORTE PRIVADO', 'Taxi'],
  ['TRANSPORTES AEREOS', 'Avión'],
  ['POSADAS DE MEXICO', 'Hotel'],
  ['RENTA DE VEHICULO', 'Renta Auto'],
  ['SOLUCION FACTIBLE', 'IT & SW'],
  ['SUSCRIPCION ANUAL', 'IT & SW'],
  ['AEROMEXICO CARGO', 'Envíos'],
  ['AEROVIAS EMPRESA', 'Avión'],
  ['ESTACIONAMIENTOS', 'Estacionamiento'],
  ['HABITACION DOBLE', 'Hotel'],
  ['HOLIDAYINNCIUDAD', 'Hotel'],
  ['RADIOMOVIL DIPSA', 'Celular'],
  ['SERVICIO DE AGUA', 'Renta Oficina'],
  ['TARIFA POR VIAJE', 'Taxi'],
  ['BOLETO DE AVION', 'Avión'],
  ['CADENA HOTELERA', 'Hotel'],
  ['GRUPO ALIMENTOS', 'Consumo'],
  ['LICENCIA DE USO', 'IT & SW'],
  ['OXXO GASCRUCERO', 'Gasolina'],
  ['REST LA COSECHA', 'Consumo Viáticos'],
  ['STARBUCKS CITAD', 'Gastos Rep'],
  ['TALLER MECANICO', 'Manto Auto'],
  ['TARIFA DE VIAJE', 'Taxi'],
  ['ABC AEROLINEAS', 'Avión'],
  ['AEROMEXICO WEB', 'Avión'],
  ['CARLS JR STIVA', 'Consumo'],
  ['ESTACIONAMIENT', 'Estacionamiento'],
  ['GASNECTARMARIA', 'Gasolina'],
  ['STARBUCKS CARR', 'Gastos Rep'],
  ['TACOS REYNALDO', 'Gastos Rep'],
  ['TAR AEROLINEAS', 'Avión'],
  ['VUELA COMPANIA', 'Avión'],
  ['VUELA COMPAÑIA', 'Avión'],
  ['CLIP MX ESTAC', 'Estacionamiento'],
  ['GRUPO POSADAS', 'Hotel'],
  ['HOTEL HAMTPON', 'Hotel'],
  ['HOTEL MEDRANO', 'Hotel'],
  ['MAGNICHARTERS', 'Avión'],
  ['OXXO TERMINAL', 'Consumo'],
  ['PAGANDO ESTAC', 'Estacionamiento'],
  ['REFACCIONARIA', 'Manto Auto'],
  ['RENTA DE AUTO', 'Renta Auto'],
  ['REST FORJADOR', 'Consumo Viáticos'],
  ['STARBUCKS EST', 'Gastos Rep'],
  ['STARBUCKS GDL', 'Gastos Rep'],
  ['STARBUCKS LAS', 'Gastos Rep'],
  ['WYNDHAMGARDEN', 'Hotel'],
  ['AGUA POTABLE', 'Renta Oficina'],
  ['BEST WESTERN', 'Hotel'],
  ['BOLETO AEREO', 'Avión'],
  ['CHILIS STIVA', 'Consumo'],
  ['CITY EXPRESS', 'Hotel'],
  ['COMBUSTIBLES', 'Gasolina'],
  ['CONFECCIONES', 'Uniformes'],
  ['GRUPO FLECHA', 'Envíos'],
  ['HOTEL ENCORE', 'Hotel'],
  ['HOTEL QUARTZ', 'Hotel'],
  ['LA SABROSITA', 'Consumo'],
  ['OFFICE DEPOT', 'Papelería'],
  ['OPERADORA OM', 'Hotel'],
  ['PANIFICADORA', 'Consumo'],
  ['PLAN MENSUAL', 'IT & SW'],
  ['RENTA CHIREY', 'Renta Auto'],
  ['RENTA NISSAN', 'Renta Auto'],
  ['RENTA TOYOTA', 'Renta Auto'],
  ['REST CASCADA', 'Consumo'],
  ['RESTAURANTES', 'Consumo Viáticos'],
  ['TARIFA AEREA', 'Avión'],
  ['A1088BMXN02', 'IT & SW'],
  ['AEROENLACES', 'Avión'],
  ['AEROLITORAL', 'Avión'],
  ['ANTIESTATIC', 'Uniformes'],
  ['DOMINIO WEB', 'IT & SW'],
  ['ELEVEN PQUE', 'Consumo'],
  ['GAS NATURAL', 'Renta Oficina'],
  ['GASOLINERIA', 'Gasolina'],
  ['HABITACION ', 'Hotel'],
  ['HOSPITALITY', 'Hotel'],
  ['OFFICEDEPOT', 'Papelería'],
  ['RCH CUCAPAH', 'Gasolina'],
  ['RECAUDADORA', 'Manto Auto'],
  ['REST CHILIS', 'Consumo'],
  ['SUPERCENTER', 'Consumo'],
  ['VIVAAEROBUS', 'Avión'],
  ['AEROLINEAS', 'Avión'],
  ['AEROPUERTO', 'Estacionamiento'],
  ['AUTOMOTRIZ', 'Manto Auto'],
  ['AUTOPISTAS', 'Casetas'],
  ['AUTOTRANSP', 'Taxi'],
  ['CAR RENTAL', 'Renta Auto'],
  ['FERRETERIA', 'Herramientas'],
  ['FIESTA INN', 'Hotel'],
  ['GASOLINERA', 'Gasolina'],
  ['HOME DEPOT', 'Herramientas'],
  ['MIT TRANSP', 'Avión'],
  ['OFFICE MAX', 'Papelería'],
  ['ONE MINUTE', 'Estacionamiento'],
  ['PAQUETERIA', 'Envíos'],
  ['PEGASO PCS', 'Celular'],
  ['RENT A CAR', 'Renta Auto'],
  ['RENTA AUTO', 'Renta Auto'],
  ['ROOM NIGHT', 'Hotel'],
  ['SUPER SERV', 'Gasolina'],
  ['TLAPALERIA', 'Herramientas'],
  ['TOTAL PLAY', 'IT & SW'],
  ['TRANS GUTS', 'Taxi'],
  ['VOLARISMXN', 'Avión'],
  ['AEROLINEA', 'Avión'],
  ['AERONAVES', 'Avión'],
  ['ANTHROPIC', 'IT & SW'],
  ['APPLE COM', 'IT & SW'],
  ['BODECATTA', 'Consumo'],
  ['CAFETERIA', 'Consumo'],
  ['CHIHUAHUA', 'Herramientas'],
  ['EBSSATAXI', 'Taxi'],
  ['FACTURAMA', 'IT & SW'],
  ['FACTURIFY', 'IT & SW'],
  ['FIESTAINN', 'Consumo'],
  ['HOSPEDAJE', 'Hotel'],
  ['INSTAGRAM', 'Marketing'],
  ['MCDONALDS', 'Consumo'],
  ['MEGACABLE', 'IT & SW'],
  ['MICROSOFT', 'IT & SW'],
  ['OFFICEMAX', 'Papelería'],
  ['PAPELERIA', 'Papelería'],
  ['SAMS CLUB', 'Consumo'],
  ['STARBUCKS', 'Consumo'],
  ['TACOS SAN', 'Gastos Rep'],
  ['TOTALPLAY', 'IT & SW'],
  ['TST BEACH', 'Consumo'],
  ['UBER EATS', 'Consumo'],
  ['UNIFORMES', 'Uniformes'],
  ['AUTOZONE', 'Manto Auto'],
  ['AVIACION', 'Avión'],
  ['CAMPOMAR', 'Gastos Rep'],
  ['CHEDRAUI', 'Consumo'],
  ['ESTACION', 'Gasolina'],
  ['ESTAFETA', 'Envíos'],
  ['FACEBOOK', 'Marketing'],
  ['HOTELERA', 'Hotel'],
  ['IMPRENTA', 'Papelería'],
  ['INTERJET', 'Avión'],
  ['IUSACELL', 'Celular'],
  ['LA COMER', 'Consumo'],
  ['LINKEDIN', 'Marketing'],
  ['MARRIOTT', 'Hotel'],
  ['MOVISTAR', 'Celular'],
  ['ONE PARK', 'Estacionamiento'],
  ['SANBORNS', 'Consumo'],
  ['SERV LAS', 'Gasolina'],
  ['SOFTWARE', 'IT & SW'],
  ['TEXTILES', 'Uniformes'],
  ['AEROMAR', 'Avión'],
  ['BOOKING', 'Hotel'],
  ['CASETAS', 'Casetas'],
  ['CHEVRON', 'Gasolina'],
  ['CONTPAQ', 'IT & SW'],
  ['DROPBOX', 'IT & SW'],
  ['HOLAFLY', 'Celular'],
  ['HOLIDAY', 'Hotel'],
  ['HOSTING', 'IT & SW'],
  ['NETFLIX', 'IT & SW'],
  ['NEUFELD', 'Consumo'],
  ['OXXO EL', 'Gastos Rep'],
  ['PARKING', 'Estacionamiento'],
  ['RECARGA', 'Casetas'],
  ['REDPACK', 'Envíos'],
  ['RENTING', 'Renta Auto'],
  ['SORIANA', 'Consumo'],
  ['SPOTIFY', 'IT & SW'],
  ['TORTAAS', 'Gastos Rep'],
  ['VOLARIS', 'Avión'],
  ['WALMART', 'Consumo'],
  ['AIRBNB', 'Hotel'],
  ['AMERIX', 'Consumo'],
  ['BP GAS', 'Gasolina'],
  ['BUDGET', 'Renta Auto'],
  ['CABIFY', 'Taxi'],
  ['CAPUFE', 'Casetas'],
  ['CF SAN', 'Consumo'],
  ['CHILIS', 'Consumo'],
  ['DENNYS', 'Gastos Rep'],
  ['DQ LAS', 'Consumo'],
  ['GITHUB', 'IT & SW'],
  ['GOOGLE', 'IT & SW'],
  ['GSUITE', 'IT & SW'],
  ['HILTON', 'Hotel'],
  ['IBERIA', 'Avión'],
  ['MARISA', 'Consumo'],
  ['NOTION', 'IT & SW'],
  ['OPENAI', 'IT & SW'],
  ['PETRO7', 'Gasolina'],
  ['REFACC', 'Manto Auto'],
  ['REPSOL', 'Gasolina'],
  ['RINCON', 'Gasolina'],
  ['STEREN', 'Herramientas'],
  ['SUBWAY', 'Consumo'],
  ['TELCEL', 'Celular'],
  ['TELMEX', 'IT & SW'],
  ['TIKTOK', 'Marketing'],
  ['TRUPER', 'Herramientas'],
  ['UNEFON', 'Celular'],
  ['VUELO ', 'Avión'],
  ['ABARR', 'Consumo'],
  ['ADOBE', 'IT & SW'],
  ['ADOLF', 'Gasolina'],
  ['CHILI', 'Consumo'],
  ['CLOUD', 'IT & SW'],
  ['FEDEX', 'Envíos'],
  ['FERRE', 'Herramientas'],
  ['FOGON', 'Consumo'],
  ['GASOL', 'Gasolina'],
  ['HERTZ', 'Renta Auto'],
  ['HYATT', 'Hotel'],
  ['MAZDA', 'Manto Auto'],
  ['MOBIL', 'Gasolina'],
  ['PARCO', 'Estacionamiento'],
  ['PEAJE', 'Casetas'],
  ['PEMEX', 'Gasolina'],
  ['PETRO', 'Gasolina'],
  ['PIZZA', 'Consumo'],
  ['PUERT', 'Estacionamiento'],
  ['SHELL', 'Gasolina'],
  ['SLACK', 'IT & SW'],
  ['SUSHI', 'Consumo'],
  ['TACOS', 'Consumo'],
  ['TAXIS', 'Taxi'],
  ['AT T', 'Celular'],
  ['ATTA', 'Taxi'],
  ['AVIS', 'Renta Auto'],
  ['BRIX', 'Consumo'],
  ['DIDI', 'Taxi'],
  ['GRAB', 'Taxi'],
  ['GULF', 'Gasolina'],
  ['IZZI', 'IT & SW'],
  ['JMAS', 'Renta Oficina'],
  ['LYFT', 'Taxi'],
  ['META', 'Marketing'],
  ['MSFT', 'IT & SW'],
  ['PAYU', 'IT & SW'],
  ['SIXT', 'Renta Auto'],
  ['TAXI', 'Taxi'],
  ['TOKS', 'Consumo'],
  ['UBER', 'Taxi'],
  ['VIPS', 'Consumo'],
  ['VIVA', 'Avión'],
  ['ZOOM', 'IT & SW'],
  ['AWS', 'IT & SW'],
  ['CFE', 'Renta Oficina'],
  ['DHL', 'Envíos'],
  ['GAS', 'Gasolina'],
  ['PZA', 'Casetas'],
  ['UPS', 'Envíos'],
]

const VENTAS_VARIANT = {
  'Avión': 'Avión Ventas',
  'Casetas': 'Casetas Ventas',
  'Estacionamiento': 'Estacionamiento Ventas',
  'Gasolina': 'Gasolina Ventas Viáticos',
  'Herramientas': 'Herramientas Ventas',
  'Hotel': 'Hotel Ventas',
  'Consumo': 'Gastos Rep',
  'Consumo Viáticos': 'Gastos Rep Viáticos',
  'Taxi': 'Gastos Rep',
}

// Variantes para colaboradores NO-Ventas (Admin/Servicio). La gasolina siempre
// es un viático en SMTO, así que se etiqueta como 'Gasolina Viáticos' en vez del
// 'Gasolina' genérico. Ventas/Socio usan VENTAS_VARIANT ('Gasolina Ventas Viáticos').
const NORMAL_VARIANT = {
  'Gasolina': 'Gasolina Viáticos',
}

// Reglas por RFC EXACTO del emisor. Ganan sobre TODO (el catch-all de Hotel y las
// reglas por nombre), porque el RFC identifica al proveedor sin ambigüedad aunque
// la razón social venga escrita distinto en el CFDI. Solo para proveedores donde
// el tipo es fijo por convenio SMTO, pase lo que pase en nombre/concepto.
export const TIPO_RULES_RFC = {
  'OOM960429832': 'Papelería',   // Operadora OMX (Office Max) → siempre Papelería
}

// Prefijos de ClaveProdServ SAT que corresponden a equipo de cómputo / IT.
// 4321xxxx = Hardware de cómputo, 4322xxxx = Periféricos, 4323xxxx = Redes.
const SAT_PC_PREFIXES = ['4321', '4322', '4323']

/**
 * @param {string} proveedor      Nombre del emisor del CFDI
 * @param {string} descripcion    Primera línea del concepto del CFDI
 * @param {string} categoria      'Admin'|'Servicio'|'Ventas'|'Socio'
 * @param {string} claveProdServ  Código SAT del producto (opcional)
 * @param {string} rfc            RFC del emisor (opcional) — activa TIPO_RULES_RFC
 */
export function autoDetectTipo(proveedor, descripcion, categoria, claveProdServ = '', rfc = '') {
  const texto = `${proveedor || ''} ${descripcion || ''}`.toUpperCase()
  const isVentas = categoria === 'Ventas' || categoria === 'Socio'

  // 0. RFC exacto — MÁXIMA prioridad. Gana sobre el catch-all de Hotel y sobre
  //    las reglas por nombre, para proveedores con tipo fijo por convenio SMTO.
  const rfcKey = (rfc || '').trim().toUpperCase()
  if (rfcKey && TIPO_RULES_RFC[rfcKey]) {
    const tipoBase = TIPO_RULES_RFC[rfcKey]
    return isVentas ? (VENTAS_VARIANT[tipoBase] || tipoBase) : (NORMAL_VARIANT[tipoBase] || tipoBase)
  }

  // FACTURIFY es plataforma de facturación usada por choferes/taxistas.
  // Cuando el concepto es "TARIFA" es la tarifa del viaje, no suscripción de software.
  if ((proveedor || '').toUpperCase().includes('FACTURIFY') &&
      (descripcion || '').toUpperCase().trim() === 'TARIFA') {
    return isVentas ? 'Gastos Rep' : 'Taxi'
  }

  // REFIEL → si la descripción empieza con "RENTA" es Automóvil; en cualquier
  // otro caso cae a las reglas genéricas / fallback.
  if ((proveedor || '').toUpperCase().includes('REFIEL') &&
      (descripcion || '').toUpperCase().trim().startsWith('RENTA')) {
    return 'Automóvil'
  }

  // Catch-all de HOTEL: cualquier factura cuyo proveedor o concepto mencione
  // hotel / hoteles / hotelera / hotelería / hospedaje / habitación / noche(s)
  // de hospedaje cae a Hotel (o Hotel Ventas para Ventas/Socio). Va ANTES de
  // las reglas genéricas para ganar prioridad y atrapar hoteles no listados
  // (p.ej. "Hotel Fiesta Americana", "Servicios de hotelería"). 'HOTEL' como
  // substring cubre HOTELES/HOTELERA/HOTELERO/HOTELERIA de un solo golpe.
  const HOTEL_KEYWORDS = ['HOTEL', 'HOSPEDAJE', 'HABITACION', 'HABITACIÓN', 'NOCHE DE HOSPEDAJE', 'ROOM NIGHT']
  if (HOTEL_KEYWORDS.some(kw => texto.includes(kw))) {
    return isVentas && VENTAS_VARIANT['Hotel'] ? VENTAS_VARIANT['Hotel'] : 'Hotel'
  }

  // 1. Primero: buscar en reglas conocidas (marcas, razones sociales, conceptos específicos)
  for (const [kw, tipoBase] of TIPO_RULES) {
    if (texto.includes(kw)) {
      if (isVentas) return VENTAS_VARIANT[tipoBase] || tipoBase
      return NORMAL_VARIANT[tipoBase] || tipoBase
    }
  }

  // 2. Código SAT — si el ClaveProdServ pertenece a hardware/cómputo, es PC (Equipo).
  //    Aplica cuando la descripción no tiene keywords reconocibles (ej. genérico "Equipo de trabajo").
  const cp = (claveProdServ || '').trim()
  if (cp && SAT_PC_PREFIXES.some(prefix => cp.startsWith(prefix))) {
    return 'PC (Equipo)'
  }

  // 3. Si ninguna regla matcheó y el concepto es exactamente "TARIFA":
  //    proveedor sin palabras corporativas = persona física (taxista/chofer) → Taxi
  const CORPORATE_WORDS = [
    'S.A', 'S.R.L', 'S.A.S', 'GRUPO ', 'EMPRESA', 'COMPAÑIA', 'COMPANIA',
    'SERVICIOS', 'RESTAURANTE', 'OPERADORA', 'DISTRIBUIDORA', 'COMERCIALIZADORA',
    'CONCESIONARIA', 'ASOCIADOS', 'CORPORATIVO', 'INMOBILIARIA', 'CONSTRUCTORA',
    'TRANSPORTES', 'PROMOTORA', 'ADMINISTRADORA', 'ARRENDADORA', 'HOTELERA',
  ]
  const provUpper = (proveedor || '').toUpperCase()
  const descUpper = (descripcion || '').toUpperCase().trim()
  const isCorporate = CORPORATE_WORDS.some(w => provUpper.includes(w))
  if (descUpper === 'TARIFA' && !isCorporate) {
    return isVentas ? 'Gastos Rep' : 'Taxi'
  }

  // 4. Fallback
  return isVentas ? 'Gastos Rep' : 'Consumo Viáticos'
}
