from http.server import BaseHTTPRequestHandler
import json, os, sys, tempfile, traceback

# Guarded imports: a module-load failure used to surface as
# FUNCTION_INVOCATION_FAILED with no useful body. Now any import error
# is captured so do_POST/do_GET can surface it.
IMPORT_ERROR = None
try:
    from openpyxl import load_workbook
    from openpyxl.drawing.image import Image as XLImage
except Exception:
    IMPORT_ERROR = traceback.format_exc()


# CFDI FormaPago codes → human label.
FORMA_PAGO_MAP = {
    '01': '01 - Efectivo',
    '02': '02 - Efectivo',
    '03': '03 - Transferencia',
    '04': '04 - Tarjeta de Crédito',
}


def forma_pago_label(code):
    return FORMA_PAGO_MAP.get(code, code or '')


def find_template():
    candidates = [
        os.path.join(os.path.dirname(__file__), 'TEMPLATE.xlsx'),
        os.path.join(os.path.dirname(__file__), '..', 'public', 'TEMPLATE.xlsx'),
        '/var/task/public/TEMPLATE.xlsx',
        '/var/task/api/TEMPLATE.xlsx',
    ]
    for path in candidates:
        if os.path.exists(path):
            return path
    raise FileNotFoundError(f'TEMPLATE.xlsx not found. Tried: {candidates}')


def find_logo():
    candidates = [
        os.path.join(os.path.dirname(__file__), 'logo.png'),
        os.path.join(os.path.dirname(__file__), '..', 'public', 'logo.png'),
        '/var/task/public/logo.png',
        '/var/task/api/logo.png',
    ]
    for path in candidates:
        if os.path.exists(path):
            return path
    return None


class handler(BaseHTTPRequestHandler):
    def do_GET(self):
        diag = {
            'ok': IMPORT_ERROR is None,
            'python_version': sys.version,
            'cwd': os.getcwd(),
            'file_dir': os.path.dirname(__file__),
            'import_error': IMPORT_ERROR,
        }
        try:
            diag['template_path'] = find_template()
        except Exception as e:
            diag['template_error'] = str(e)
        diag['logo_path'] = find_logo()
        self.send_response(200)
        self.send_header('Content-Type', 'application/json')
        self.send_header('Access-Control-Allow-Origin', '*')
        self.end_headers()
        self.wfile.write(json.dumps(diag, indent=2).encode())

    def do_POST(self):
        try:
            if IMPORT_ERROR:
                raise RuntimeError(f'Module import failed at function load:\n{IMPORT_ERROR}')

            content_length = int(self.headers.get('Content-Length', 0))
            body = self.rfile.read(content_length)
            gastos = json.loads(body)

            wb = load_workbook(find_template())
            ws = wb.active   # first sheet — template has only "Gastos"

            # openpyxl uses 1-indexed rows + columns. Header is at row 5, data
            # starts at row 6.
            row = 6
            for g in gastos:
                f = g.get('fechaFac', '').split('-')
                fecha_mx = f'{f[2]}/{f[1]}/{f[0]}' if len(f) == 3 else g.get('fechaFac', '')

                ws.cell(row=row, column=1,  value=g.get('rfc', ''))
                ws.cell(row=row, column=2,  value=g.get('proveedor', ''))
                ws.cell(row=row, column=3,  value=g.get('noFactura', ''))
                ws.cell(row=row, column=4,  value=fecha_mx)
                ws.cell(row=row, column=5,  value=g.get('concepto', ''))
                ws.cell(row=row, column=6,  value=round(g.get('importe', 0), 2))
                ws.cell(row=row, column=7,  value=round(g.get('iva', 0), 2))
                ws.cell(row=row, column=8,  value=round(g.get('retenciones', 0), 2))
                ws.cell(row=row, column=9,  value=f'=F{row}+G{row}-H{row}')
                ws.cell(row=row, column=10, value=forma_pago_label(g.get('formaPago', '')))
                ws.cell(row=row, column=11, value=g.get('fechaCobro', 'Pendiente'))
                row += 1

                propina = round(g.get('montoPropina', 0), 2)
                if propina > 0:
                    ws.cell(row=row, column=2,  value=g.get('proveedor', '') + ' - PROPINA')
                    ws.cell(row=row, column=4,  value=fecha_mx)
                    ws.cell(row=row, column=5,  value='PROPINA')
                    ws.cell(row=row, column=6,  value=propina)
                    ws.cell(row=row, column=7,  value=0)
                    ws.cell(row=row, column=8,  value=0)
                    ws.cell(row=row, column=9,  value=f'=F{row}+G{row}-H{row}')
                    ws.cell(row=row, column=10, value=forma_pago_label(g.get('formaPago', '')))
                    ws.cell(row=row, column=11, value=g.get('fechaCobro', 'Pendiente'))
                    row += 1

            # Clear leftover template pre-fills in any data rows we didn't reach.
            # Stops at row 24 so the totals row stays in place.
            for clear_r in range(row, 24):
                for c in [6, 7, 8, 9, 12, 13]:
                    ws.cell(row=clear_r, column=c).value = None

            # Totals row (Excel row 24). openpyxl preserves the template's cell
            # styling when we set .value; we only need to write the formulas.
            ws.cell(row=24, column=5, value='Total Cuenta:')
            ws.cell(row=24, column=6, value='=SUM(F6:F23)')
            ws.cell(row=24, column=7, value='=SUM(G6:G23)')
            ws.cell(row=24, column=8, value='=SUM(H6:H23)')
            ws.cell(row=24, column=9, value='=SUM(I6:I23)')

            # Embed the logo (PNG, native — openpyxl uses Pillow under the hood
            # so transparency etc. is preserved). Anchored top-left.
            logo_path = find_logo()
            if logo_path:
                img = XLImage(logo_path)
                img.width = 180
                img.height = 80
                ws.add_image(img, 'A1')

            with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
                tmp_path = tmp.name
            wb.save(tmp_path)
            with open(tmp_path, 'rb') as f:
                file_data = f.read()
            os.unlink(tmp_path)

            self.send_response(200)
            self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            self.send_header('Content-Disposition', 'attachment; filename="Reporte_Gastos_SMTO.xlsx"')
            self.send_header('Content-Length', str(len(file_data)))
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            self.wfile.write(file_data)

        except Exception as e:
            error_msg = traceback.format_exc()
            self.send_response(500)
            self.send_header('Content-Type', 'application/json')
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            self.wfile.write(json.dumps({'error': str(e), 'trace': error_msg}).encode())

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()
