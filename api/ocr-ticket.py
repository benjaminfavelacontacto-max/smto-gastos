from http.server import BaseHTTPRequestHandler
import json, os, base64, io, traceback, re, time
import urllib.request, urllib.error

# Groq es 100% gratuito (capa free con límites de uso) y reemplaza a la API
# de Anthropic para no generar costo al leer tickets/PDFs sin XML.
# Modelo de visión vigente en GroqCloud (Maverick fue deprecado feb-2026).
GROQ_URL   = 'https://api.groq.com/openai/v1/chat/completions'
GROQ_MODEL = 'meta-llama/llama-4-scout-17b-16e-instruct'

# Límite de Groq para imágenes en base64: 4MB. Dejamos margen apuntando a ~3MB
# de imagen cruda (base64 expande ~1.33x) y reescalando si hace falta.
MAX_IMG_BYTES = 3_000_000

# Dependencias opcionales: se importan perezosamente para que un fallo de
# import no tumbe todo el handler (igual que el patrón previo con anthropic).
IMPORT_ERROR = None
try:
    import fitz  # PyMuPDF — rasteriza PDF → imagen (Groq no lee PDFs)
    from PIL import Image
except Exception:
    IMPORT_ERROR = traceback.format_exc()


OCR_PROMPT = '''Extract receipt/ticket/invoice data. Respond ONLY with a JSON object, no markdown, no explanation:
{
  "tipoDocumento": "ticket | pedimento",
  "proveedor": "business or restaurant name",
  "concepto": "brief description (food, gas, hotel, taxi, etc)",
  "fecha": "YYYY-MM-DD or null",
  "moneda": "3-letter ISO code: MXN, USD, EUR, GBP, JPY, CHF, CAD, etc.",
  "subtotal": 0.00,
  "iva": 0.00,
  "propina": 0.00,
  "total": 0.00,
  "propinaSugerida18": 0.00,
  "propinaSugerida20": 0.00,
  "propinaSugerida22": 0.00,
  "folio": "card authorization / approval code as plain digits",
  "formaPago": "04"
}

DETECCION DEL TIPO DE DOCUMENTO:
- Si el documento contiene los textos "Num Pedimento", "Pedimento", "Importe Pagado", "Fecha de Pago" y/o "Aduana" → es un PEDIMENTO de importacion mexicano. Marca tipoDocumento="pedimento".
- En cualquier otro caso → tipoDocumento="ticket" y aplica las reglas de ticket/recibo normales.

REGLAS PARA PEDIMENTO (tipoDocumento="pedimento"):
- folio: el numero que aparece junto a "Num Pedimento" o "Numero de Pedimento" — preservalo TAL CUAL (con guiones, espacios, etc, sin alterar). Este es el identificador unico del pedimento. Ejemplo: "26 48 3993 6406196"
- fecha: la fecha que aparece junto a "Fecha de Pago" (formato YYYY-MM-DD). NO uses la fecha de emision ni la fecha de operacion.
- total: el monto que aparece junto a "Importe Pagado" o "Total Pagado". Es la cifra final en pesos mexicanos. Quita las comas de miles antes de convertir a numero (1,449.00 → 1449.00).
- subtotal: usa el mismo valor que total (no aplica desglose de IVA en pedimentos).
- iva: 0 (los pedimentos ya incluyen los impuestos prorrateados en el importe pagado).
- propina: 0
- proveedor: "SMTO ENGINEERING"
- concepto: "Pedimento de Importacion"
- moneda: "MXN" (los pedimentos se pagan siempre en pesos mexicanos)
- formaPago: "03" (transferencia)
- propinaSugerida18/20/22: 0

REGLAS PARA TICKET (tipoDocumento="ticket"):
- proveedor: full business name as shown on receipt
- folio: look for Approval Code, Authorization Code, Auth, Approval, Code — return ONLY the digits as a plain string. If multiple codes exist prefer the one labeled "Approval" or "Authorization". This number is critical for bank matching, so do not strip leading zeros or insert dashes. If no auth code is shown fall back to a visible Check #, Ticket # or Receipt # in the same plain-digits form.
- formaPago: 04=card/mastercard/visa, 02=cash/efectivo
- moneda: the 3-letter ISO currency code of the amounts on the receipt. Detect it from the currency symbol AND the country/address on the ticket:
  - "€" or "EUR" -> EUR;  "£" or "GBP" -> GBP;  "¥" / "円" / "JPY" -> JPY (use CNY only if clearly from China);  "CHF" -> CHF;  "R$" -> BRL.
  - "$" is AMBIGUOUS: use USD if the receipt is from the USA (US city/state/ZIP, "US$", or prices say USD); CAD if from Canada; and MXN ONLY when there is positive evidence the receipt is Mexican (Mexican address, RFC, "IVA", "M.N.", "pesos", or amounts clearly in pesos).
  - If an explicit ISO code is printed on the receipt, trust it over the symbol.
  - CRITICAL: NEVER default a foreign-looking receipt (European/Asian/other address, or a non-"$" symbol like €, £, ¥) to MXN. Return MXN only with real evidence it is Mexican; if a European ticket shows "€", the answer is EUR, never MXN.
- subtotal: amount before tax and tip
- iva: tax amount (Tax, IVA, Impuesto)
- propina: tip amount (Tip, Gratuity, Propina), often HANDWRITTEN on a restaurant receipt — 0 if not present
- total: the PRE-TIP total ONLY — the consumption amount including tax but BEFORE the tip. This is the printed/authorized "Total" or "Amount" (e.g. the card pre-authorization), NOT the handwritten grand total at the bottom. NEVER add the tip into this field. When a diner handwrites a grand total (= base + tip), you MUST split it: the printed base goes in "total" and the handwritten tip goes in "propina" — never put the grand total here. Example: printed "Total $68.00", handwritten "Tip $12.24", handwritten grand "Total $80.24" → total=68.00, propina=12.24 (NEVER total=80.24)
- propinaSugerida18/20/22: look for "Suggested Gratuity" table on the receipt and extract the tip amounts for 18%, 20%, 22% — use 0 if not present
- If any field not found use null for strings and 0 for numbers

REGLA ESPECIAL — RECIBOS DE ITESO (Universidad Jesuita de Guadalajara):
- Si el documento tiene el logo o encabezado "ITESO" o "Universidad Jesuita de Guadalajara" → es un recibo de ITESO. Aplica estas reglas:
- proveedor: "ITESO" (NO uses el nombre que aparece junto a "Nombre:", ese es el cliente, no el proveedor).
- concepto: copia TAL CUAL la descripcion de la LINEA del concepto, p.ej. "ADEUDO DEPOSITO GARANTIA RENTA PARQUE TE 1DI".
- total: usa el importe que aparece en la LINEA del concepto (p.ej. 1370.00). NO uses el campo "Total :" porque suele venir en 0.00. Quita comas de miles.
- subtotal: mismo valor que total.
- iva: 0
- propina: 0
- moneda: "MXN"
- formaPago: "03"

REGLA ESPECIAL — RECIBOS DE ISN / SECRETARIA DE LA HACIENDA PUBLICA (Jalisco, "SET Jalisco", "RECIBO OFICIAL FORMA UNIVERSAL UNICA"):
- Si el documento es un recibo oficial de impuesto estatal con encabezado "Secretaria de la Hacienda Publica" o tramite "Impuesto Sobre Nomina (ISN)" → aplica estas reglas:
- proveedor: "Secretaria de la Hacienda Publica" (NO uses el nombre del contribuyente, p.ej. "SMTO ENGINEERING", ese es el cliente).
- concepto: la descripcion del concepto, p.ej. "Impuesto Sobre Nomina (ISN)".
- total: el importe que aparece en "TOTAL" / "Sub Total" / columna "IMPORTE" (p.ej. 105533.00). NUNCA uses la cifra de "Base gravable" (es la nomina base, no el impuesto a pagar). Quita comas de miles.
- subtotal: mismo valor que total.
- iva: 0
- folio: el numero grande de "Folio" (p.ej. 48297279), no el "Folio Banco".
- moneda: "MXN"
- formaPago: "03"

REGLA ESPECIAL — RECIBOS DE TOTAL PLAY / TOTALPLAY (estado de cuenta "Totalplay Empresarial"):
- Si el documento tiene el logo o encabezado "Totalplay", "Total Play" o "Totalplay Empresarial" → es un estado de cuenta de Total Play. Aplica estas reglas:
- proveedor: "Total Play".
- total: usa SIEMPRE el importe de la linea "Cargos del Mes" (p.ej. 1,378.65). NUNCA uses las lineas "A PAGAR" — ni "Precio de lista" (viene redondeado) ni "Precio pronto pago" (trae descuento). El banco cobra el monto de "Cargos del Mes". Quita comas de miles.
- subtotal: el importe de "SubTotal" (p.ej. 1,188.50).
- iva: la diferencia "Cargos del Mes" menos "SubTotal" (NO copies el IVA impreso si no cuadra con esa resta).
- propina: 0
- moneda: "MXN"
- formaPago: "03"

REGLA ESPECIAL — FACTURAS DE MICROSOFT (logo de Microsoft / "Billing Summary" / "Microsoft Corporation"):
- Si el documento tiene el logo de Microsoft o el encabezado "Billing Summary" y/o "Microsoft Corporation" → aplica estas reglas:
- proveedor: "Microsoft".
- folio: el valor que aparece junto a "Billing Number" (p.ej. G160071537). Preservalo EXACTAMENTE como aparece, INCLUYENDO la letra inicial. NO lo conviertas a solo digitos ni lo confundas con el "Document Date" ni con el RFC.
- fecha: la fecha junto a "Document Date". Microsoft la imprime en formato DD/MM/YYYY (el PRIMER numero es el DIA, el SEGUNDO es el MES). Conviertela a YYYY-MM-DD respetando ese orden. Ejemplos: "21/05/2026" -> "2026-05-21"; "02/05/2026" -> "2026-05-02" (NUNCA "2026-02-05").
- total: el "Total Amount" (p.ej. 121.24).
- moneda: "USD" si el monto aparece como "USD".
- subtotal: mismo valor que total.
- iva: 0 (el Billing Summary no desglosa IVA).
- propina: 0
- concepto: "Microsoft Office" o la suscripcion indicada.
- formaPago: "04"

REGLA ESPECIAL — FACTURAS DE GW INSTEK / INSTEK AMERICA CORP:
- Si el documento tiene el logo "GW INSTEK" o el encabezado "INSTEK AMERICA CORP" → proveedor: "INSTEK AMERICA CORP" (NO uses "GW Instek" del logo). El RFC "SEN..." que aparece en "Bill To" es de SMTO, NO del proveedor; ignoralo.

REGLA ESPECIAL — FACTURAS DE PEAJE / FONDO NACIONAL DE INFRAESTRUCTURA (FNI, "SERVICIO DE PEAJE Y CRUCE CARRETERO"):
- Si el documento es un CFDI con emisor "FONDO NACIONAL DE INFRAESTRUCTURA" (RFC FNI970829JR9) o el concepto dice "SERVICIO DE PEAJE Y CRUCE CARRETERO" → aplica estas reglas:
- proveedor: "Fondo Nacional de Infraestructura".
- folio: concatena la "Serie" y el "Folio" que aparecen arriba en el recuadro del CFDI, TAL CUAL, en ese orden. Ejemplo: Serie "FNPE" + Folio "72984235" -> "FNPE72984235". NO uses el "Folio fiscal" (UUID) ni el "No. de certificado".
- moneda: "MXN".
- formaPago: "01".

REGLA GENERAL DE FECHA (facturas/recibos de EE.UU.):
- Si un documento de EE.UU. imprime la fecha con numeros ambiguos, recuerda: si algun componente es > 12 ese es el DIA. Devuelve siempre "fecha" en formato YYYY-MM-DD ya desambiguado.'''


def _shrink_image(img):
    '''Reescala/recomprime una PIL.Image a JPEG <= MAX_IMG_BYTES.'''
    if img.mode not in ('RGB', 'L'):
        img = img.convert('RGB')
    quality = 82
    # Borde largo máx 1800px: legible para tickets y mantiene la imagen chica
    # para que Groq responda rápido y use menos tokens (menos rate limit).
    max_dim = 1800
    while True:
        w, h = img.size
        if max(w, h) > max_dim:
            scale = max_dim / float(max(w, h))
            img = img.resize((max(1, int(w * scale)), max(1, int(h * scale))))
        buf = io.BytesIO()
        img.save(buf, format='JPEG', quality=quality)
        out = buf.getvalue()
        if len(out) <= MAX_IMG_BYTES or (quality <= 45 and max_dim <= 1200):
            return out
        # Aprieta calidad primero, luego dimensiones.
        if quality > 45:
            quality -= 15
        else:
            max_dim = int(max_dim * 0.8)


def _extract_fni_folio(raw_pdf_bytes):
    '''Fondo Nacional de Infraestructura (peaje): su CFDI XML no trae Serie/Folio,
    pero el PDF SÍ los imprime en texto ("Folio: 73123541 Serie: FNPE"). Los
    extraemos de forma determinista del texto embebido (más confiable que la
    visión). Devuelve "SERIEFOLIO" (p.ej. "FNPE73123541") o None.'''
    try:
        doc = fitz.open(stream=raw_pdf_bytes, filetype='pdf')
        text = doc.load_page(0).get_text()
        doc.close()
    except Exception:
        return None
    up = text.upper()
    if 'FONDO NACIONAL DE INFRAESTRUCTURA' not in up and 'FNI970829JR9' not in up:
        return None
    # "Folio: 73123541 Serie: FNPE" — el colon distingue del "Folio fiscal" (UUID).
    folio_m = re.search(r'Folio:\s*([A-Za-z0-9\-]+)', text)
    serie_m = re.search(r'Serie:\s*([A-Za-z0-9]+)', text)
    if folio_m and serie_m:
        return (serie_m.group(1) + folio_m.group(1)).strip()
    if folio_m:
        return folio_m.group(1).strip()
    return None


def _to_image_data_url(base64_data, media_type):
    '''Devuelve un data URL de imagen JPEG listo para Groq.
    Convierte PDFs (página 1) a imagen y reescala lo que exceda 4MB.'''
    raw = base64.b64decode(base64_data)

    if media_type == 'application/pdf':
        doc = fitz.open(stream=raw, filetype='pdf')
        page = doc.load_page(0)
        # Renderizamos apuntando a un borde largo de ~1600px (suficiente para
        # leer texto de tickets) en vez de un zoom fijo. Así una página grande no
        # genera un pixmap enorme: imagen más chica → Groq responde más rápido y
        # consume menos tokens (menos rate limit). Cap a 2.0 para no agrandar
        # tickets chiquitos de más.
        long_pt = max(page.rect.width, page.rect.height) or 1
        zoom = min(2.0, max(1.0, 1600.0 / long_pt))
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
        img = Image.open(io.BytesIO(pix.tobytes('png')))
        doc.close()
        jpeg = _shrink_image(img)
        return 'data:image/jpeg;base64,' + base64.b64encode(jpeg).decode()

    # Imagen: si ya cabe en el límite la dejamos tal cual; si no, recomprimimos.
    if len(raw) <= MAX_IMG_BYTES:
        mt = media_type if media_type.startswith('image/') else 'image/jpeg'
        return f'data:{mt};base64,{base64_data}'
    img = Image.open(io.BytesIO(raw))
    jpeg = _shrink_image(img)
    return 'data:image/jpeg;base64,' + base64.b64encode(jpeg).decode()


class handler(BaseHTTPRequestHandler):
    def do_POST(self):
        try:
            if IMPORT_ERROR:
                raise RuntimeError(f'Import failed: {IMPORT_ERROR}')

            content_length = int(self.headers.get('Content-Length', 0))
            body = self.rfile.read(content_length)
            data = json.loads(body)

            base64_data = data.get('base64')
            media_type = data.get('mediaType', 'image/jpeg')
            # folioOnly: el cliente solo quiere el Serie+Folio de una factura FNI
            # (peaje) cuyo XML no lo trae. Lo sacamos del TEXTO del PDF y
            # respondemos al instante, SIN llamar a Groq (sin visión, sin rate
            # limit). Evita que cargar una carpeta con varias FNI se trabe.
            folio_only = bool(data.get('folioOnly', False))
            if folio_only:
                fni_folio = None
                if media_type == 'application/pdf':
                    try:
                        fni_folio = _extract_fni_folio(base64.b64decode(base64_data))
                    except Exception:
                        fni_folio = None
                self.send_response(200)
                self.send_header('Content-Type', 'application/json')
                self.send_header('Access-Control-Allow-Origin', '*')
                self.end_headers()
                self.wfile.write(json.dumps({
                    'folio': fni_folio or '',
                    'proveedor': 'Fondo Nacional de Infraestructura' if fni_folio else '',
                }).encode())
                return

            api_key = os.environ.get('SMTO_GROQ_API_KEY', '')
            if not api_key:
                raise RuntimeError('SMTO_GROQ_API_KEY not configured in Vercel environment variables')

            image_url = _to_image_data_url(base64_data, media_type)

            payload = {
                'model': GROQ_MODEL,
                'temperature': 0,
                'max_tokens': 1000,
                'response_format': {'type': 'json_object'},
                'messages': [{
                    'role': 'user',
                    'content': [
                        {'type': 'image_url', 'image_url': {'url': image_url}},
                        {'type': 'text', 'text': OCR_PROMPT},
                    ],
                }],
            }

            req = urllib.request.Request(
                GROQ_URL,
                data=json.dumps(payload).encode(),
                headers={
                    'Content-Type': 'application/json',
                    'Authorization': f'Bearer {api_key}',
                    # Cloudflare (delante de Groq) banea el User-Agent por
                    # defecto de urllib (Python-urllib/x.y) con error 1010.
                    # Mandamos uno normal para que no nos bloquee.
                    'User-Agent': 'smto-app/1.0 (+https://smto-app.vercel.app)',
                    'Accept': 'application/json',
                },
                method='POST',
            )
            # Groq (tier gratuito) limita por tokens/minuto. Al subir varias
            # facturas seguidas las últimas llamadas chocan con un 429 y antes
            # se perdía la factura. Reintentamos respetando el tiempo que Groq
            # indica (header Retry-After o el "try again in Xs" del mensaje).
            # Vercel permite hasta 300s, así que esperar unos segundos aquí es
            # seguro y deja que la factura entre en vez de desaparecer.
            groq_raw = None
            max_retries = 5
            for attempt in range(max_retries):
                try:
                    with urllib.request.urlopen(req, timeout=60) as resp:
                        groq_raw = resp.read().decode()
                    break
                except urllib.error.HTTPError as he:
                    detail = he.read().decode(errors='replace')
                    # Solo reintentamos en 429 (rate limit). El resto es fatal.
                    if he.code != 429 or attempt == max_retries - 1:
                        raise RuntimeError(f'Groq API {he.code}: {detail}')
                    wait = None
                    ra = he.headers.get('Retry-After') if he.headers else None
                    if ra:
                        try: wait = float(ra)
                        except ValueError: wait = None
                    if wait is None:
                        m = re.search(r'try again in\s*([0-9.]+)\s*s', detail, re.I)
                        if m: wait = float(m.group(1))
                    # Fallback con backoff si Groq no dio un tiempo explícito.
                    if wait is None:
                        wait = 3.0 * (attempt + 1)
                    # Margen + tope para no exceder el límite de la función.
                    wait = min(wait + 0.5, 30.0)
                    time.sleep(wait)
            if groq_raw is None:
                raise RuntimeError('Groq API: sin respuesta tras reintentos')

            groq_json = json.loads(groq_raw)
            text = groq_json['choices'][0]['message']['content']
            clean = text.replace('```json', '').replace('```', '').strip()
            result = json.loads(clean)

            # FNI (peaje): sobrescribe el folio con el Serie+Folio leído del texto
            # del PDF (determinista). La visión a veces devuelve solo el número o
            # lo confunde; el texto embebido siempre trae "Folio: N Serie: FNPE".
            if media_type == 'application/pdf':
                try:
                    fni_folio = _extract_fni_folio(base64.b64decode(base64_data))
                    if fni_folio:
                        result['folio'] = fni_folio
                        result['proveedor'] = 'Fondo Nacional de Infraestructura'
                except Exception:
                    pass

            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            self.wfile.write(json.dumps(result).encode())

        except Exception as e:
            self.send_response(500)
            self.send_header('Content-Type', 'application/json')
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            self.wfile.write(json.dumps({
                'error': str(e),
                'trace': traceback.format_exc()
            }).encode())

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()
