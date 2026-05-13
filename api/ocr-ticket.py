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
                model='claude-sonnet-4-5-20251001',
                max_tokens=1000,
                messages=[{
                    'role': 'user',
                    'content': [
                        content_block,
                        {
                            'type': 'text',
                            'text': '''Extract receipt/ticket/invoice data. Respond ONLY with a JSON object, no markdown, no explanation:
{
  "proveedor": "business or restaurant name",
  "concepto": "brief description (food, gas, hotel, taxi, etc)",
  "fecha": "YYYY-MM-DD or null",
  "moneda": "MXN or USD",
  "subtotal": 0.00,
  "iva": 0.00,
  "propina": 0.00,
  "total": 0.00,
  "folio": "ticket check receipt or approval code number",
  "formaPago": "04"
}
Rules:
- proveedor: full business name as shown on receipt
- folio: look for Check #, Ticket #, Receipt #, Approval Code, Folio, Authorization — use the most specific reference number available
- formaPago: 04=card/mastercard/visa, 02=cash/efectivo
- moneda: if prices are in USD or receipt is from USA use USD, otherwise MXN
- subtotal: amount before tax and tip
- iva: tax amount (Tax, IVA, Impuesto)
- propina: tip amount (Tip, Gratuity, Propina) — 0 if not present
- total: final total including tax but before tip, or grand total if tip included
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
