from http.server import BaseHTTPRequestHandler
import json, os, base64, io, traceback
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
  "moneda": "MXN or USD",
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
- moneda: if prices are in USD or receipt is from USA use USD, otherwise MXN
- subtotal: amount before tax and tip
- iva: tax amount (Tax, IVA, Impuesto)
- propina: tip amount (Tip, Gratuity, Propina) — 0 if not present
- total: final total including tax but before tip, or grand total if tip included
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

REGLA GENERAL DE FECHA (facturas/recibos de EE.UU.):
- Si un documento de EE.UU. imprime la fecha con numeros ambiguos, recuerda: si algun componente es > 12 ese es el DIA. Devuelve siempre "fecha" en formato YYYY-MM-DD ya desambiguado.'''


def _shrink_image(img):
    '''Reescala/recomprime una PIL.Image a JPEG <= MAX_IMG_BYTES.'''
    if img.mode not in ('RGB', 'L'):
        img = img.convert('RGB')
    quality = 85
    max_dim = 2400
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


def _to_image_data_url(base64_data, media_type):
    '''Devuelve un data URL de imagen JPEG listo para Groq.
    Convierte PDFs (página 1) a imagen y reescala lo que exceda 4MB.'''
    raw = base64.b64decode(base64_data)

    if media_type == 'application/pdf':
        doc = fitz.open(stream=raw, filetype='pdf')
        page = doc.load_page(0)
        # zoom 2.0 ≈ 144 DPI: buena lectura sin reventar el tamaño.
        pix = page.get_pixmap(matrix=fitz.Matrix(2.0, 2.0))
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
            try:
                with urllib.request.urlopen(req, timeout=60) as resp:
                    groq_raw = resp.read().decode()
            except urllib.error.HTTPError as he:
                detail = he.read().decode(errors='replace')
                raise RuntimeError(f'Groq API {he.code}: {detail}')

            groq_json = json.loads(groq_raw)
            text = groq_json['choices'][0]['message']['content']
            clean = text.replace('```json', '').replace('```', '').strip()
            result = json.loads(clean)

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
