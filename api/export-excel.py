from http.server import BaseHTTPRequestHandler
import json, os, sys, tempfile, traceback

# Guarded imports: if these fail, FUNCTION_INVOCATION_FAILED used to swallow the
# error. Now we capture it and surface it via the response body on first call.
IMPORT_ERROR = None
try:
    import xlrd
    import xlwt
    from xlutils.copy import copy as xl_copy
except Exception:
    IMPORT_ERROR = traceback.format_exc()


# CFDI FormaPago codes → human label written to the Excel "FORMA DE PAGO" column.
# Unknown codes fall through to the raw code so SAT values like "28" don't get lost.
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
        os.path.join(os.path.dirname(__file__), 'TEMPLATE.xls'),
        os.path.join(os.path.dirname(__file__), '..', 'public', 'TEMPLATE.xls'),
        '/var/task/public/TEMPLATE.xls',
        '/var/task/api/TEMPLATE.xls',
    ]
    for path in candidates:
        if os.path.exists(path):
            return path
    raise FileNotFoundError(f'TEMPLATE.xls not found. Tried: {candidates}')


class handler(BaseHTTPRequestHandler):
    # Connectivity probe: GET returns 200 + diagnostic JSON so we can verify the
    # function is reachable even when POST is failing.
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

            template_path = find_template()
            rb = xlrd.open_workbook(template_path, formatting_info=True)
            wb = xl_copy(rb)
            ws = wb.get_sheet(0)

            row_idx = 5
            for g in gastos:
                f = g.get('fechaFac', '').split('-')
                fecha_mx = f'{f[2]}/{f[1]}/{f[0]}' if len(f) == 3 else g.get('fechaFac', '')
                total_final = round(g.get('totalCFDI', 0) + g.get('montoPropina', 0), 2)
                ws.write(row_idx, 0, g.get('rfc', ''))
                ws.write(row_idx, 1, g.get('proveedor', ''))
                ws.write(row_idx, 2, g.get('noFactura', ''))
                ws.write(row_idx, 3, fecha_mx)
                ws.write(row_idx, 4, g.get('concepto', ''))
                ws.write(row_idx, 5, round(g.get('importe', 0), 2))
                ws.write(row_idx, 6, round(g.get('iva', 0), 2))
                ws.write(row_idx, 7, round(g.get('retenciones', 0), 2))
                # Mirror the template's I<n> = F<n> + G<n> - H<n> formula so
                # Excel recomputes Total if the user later edits importe/IVA/ret.
                ws.write(row_idx, 8, xlwt.Formula(f'F{row_idx+1}+G{row_idx+1}-H{row_idx+1}'))
                ws.write(row_idx, 9, forma_pago_label(g.get('formaPago', '')))
                ws.write(row_idx, 10, g.get('fechaCobro', 'Pendiente'))
                row_idx += 1
                propina = round(g.get('montoPropina', 0), 2)
                if propina > 0:
                    ws.write(row_idx, 0, '')
                    ws.write(row_idx, 1, g.get('proveedor', '') + ' - PROPINA')
                    ws.write(row_idx, 2, '')
                    ws.write(row_idx, 3, fecha_mx)
                    ws.write(row_idx, 4, 'PROPINA')
                    ws.write(row_idx, 5, propina)
                    ws.write(row_idx, 6, 0)
                    ws.write(row_idx, 7, 0)
                    ws.write(row_idx, 8, xlwt.Formula(f'F{row_idx+1}+G{row_idx+1}-H{row_idx+1}'))
                    ws.write(row_idx, 9, forma_pago_label(g.get('formaPago', '')))
                    ws.write(row_idx, 10, g.get('fechaCobro', 'Pendiente'))
                    row_idx += 1

            # ── Clear leftover template pre-fills (zeros in cols 5,6,7,8,11,12)
            # in any data rows we didn't reach. Stops at row 23 so the templated
            # totals row (Excel row 24) stays in place.
            for clear_r in range(row_idx, 23):
                for c in [5, 6, 7, 8, 11, 12]:
                    ws.write(clear_r, c, '')

            # ── Totals row (Excel row 24, 0-indexed 23). Overwrite the template's
            # cells with explicit SUM formulas + label. Fixes the template's
            # I24 = SUM(N6:N22) bug (should sum column I, not N).
            ws.write(23, 4, 'Total Cuenta:')
            ws.write(23, 5, xlwt.Formula('SUM(F6:F23)'))
            ws.write(23, 6, xlwt.Formula('SUM(G6:G23)'))
            ws.write(23, 7, xlwt.Formula('SUM(H6:H23)'))
            ws.write(23, 8, xlwt.Formula('SUM(I6:I23)'))

            with tempfile.NamedTemporaryFile(suffix='.xls', delete=False) as tmp:
                tmp_path = tmp.name
            wb.save(tmp_path)
            with open(tmp_path, 'rb') as f:
                file_data = f.read()
            os.unlink(tmp_path)

            self.send_response(200)
            self.send_header('Content-Type', 'application/vnd.ms-excel')
            self.send_header('Content-Disposition', 'attachment; filename="Reporte_Gastos_SMTO.xls"')
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
