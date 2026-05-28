from http.server import BaseHTTPRequestHandler
import json, os, traceback

IMPORT_ERROR = None
try:
    import anthropic
except Exception:
    IMPORT_ERROR = traceback.format_exc()

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

            api_key = os.environ.get('ANTHROPIC_API_KEY', '')
            if not api_key:
                raise RuntimeError('ANTHROPIC_API_KEY not configured in Vercel environment variables')

            client = anthropic.Anthropic(api_key=api_key)

            if media_type == 'application/pdf':
                content_block = {
                    'type': 'document',
                    'source': {'type': 'base64', 'media_type': 'application/pdf', 'data': base64_data}
                }
            else:
                content_block = {
                    'type': 'image',
                    'source': {'type': 'base64', 'media_type': media_type, 'data': base64_data}
                }

            message = client.messages.create(
                model='claude-sonnet-4-6',
                max_tokens=1000,
                messages=[{
                    'role': 'user',
                    'content': [
                        content_block,
                        {
                            'type': 'text',
                            'text': '''Extract receipt/ticket/invoice data. Respond ONLY with a JSON object, no markdown, no explanation:
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
- folio: el numero que aparece junto a "Num Pedimento" o "Numero de Pedimento" — preservalo TAL CUAL (con guiones, espacios, etc, sin alterar). Este es el identificador unico del pedimento.
- fecha: la fecha que aparece junto a "Fecha de Pago" (formato YYYY-MM-DD). NO uses la fecha de emision ni la fecha de operacion.
- total: el monto que aparece junto a "Importe Pagado" o "Total Pagado". Es la cifra final en pesos mexicanos.
- subtotal: usa el mismo valor que total (no aplica desglose de IVA en pedimentos).
- iva: 0 (los pedimentos ya incluyen los impuestos prorrateados en el importe pagado).
- propina: 0
- proveedor: "Aduana"
- concepto: "Tramite de Aduana"
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
- If any field not found use null for strings and 0 for numbers'''
                        }
                    ]
                }]
            )

            text = message.content[0].text
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
