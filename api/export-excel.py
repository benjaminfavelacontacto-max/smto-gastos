from http.server import BaseHTTPRequestHandler
import json, os, sys, tempfile, traceback

# Guarded imports: if these fail, FUNCTION_INVOCATION_FAILED used to swallow the
# error. Now we capture it and surface it via the response body on first call.
IMPORT_ERROR = None
try:
    import xlrd
    from xlutils.copy import copy as xl_copy
except Exception:
    IMPORT_ERROR = traceback.format_exc()


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
                ws.write(row_idx, 8, total_final)
                ws.write(row_idx, 9, g.get('formaPago', ''))
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
                    ws.write(row_idx, 8, propina)
                    ws.write(row_idx, 9, g.get('formaPago', ''))
                    ws.write(row_idx, 10, g.get('fechaCobro', 'Pendiente'))
                    row_idx += 1

            # ── Totales (siempre justo después de la última fila de datos) ──
            totals_row = row_idx

            total_importe = round(sum(g.get('importe', 0) for g in gastos), 2)
            total_iva = round(sum(g.get('iva', 0) for g in gastos), 2)
            total_ret = round(sum(g.get('retenciones', 0) for g in gastos), 2)
            total_cfdi = round(sum(g.get('totalCFDI', 0) + g.get('montoPropina', 0) for g in gastos), 2)
            total_propinas_importe = round(sum(g.get('montoPropina', 0) for g in gastos), 2)

            # Re-open the template (read-only) so we can copy back its content
            # for the rows beyond `totals_row` (preserves any original layout
            # past where our data landed).
            rb_orig = xlrd.open_workbook(template_path, formatting_info=True)
            ws_orig = rb_orig.sheet_by_index(0)

            ws.write(totals_row, 4, 'Total Cuenta:')
            ws.write(totals_row, 5, total_importe)
            ws.write(totals_row, 6, total_iva)
            ws.write(totals_row, 7, total_ret)
            ws.write(totals_row, 8, total_cfdi)

            for clear_row in range(totals_row + 1, 24):
                for col in range(16):
                    orig_cell = ws_orig.cell(clear_row, col)
                    ws.write(clear_row, col, '' if orig_cell.ctype == 0 else orig_cell.value)

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
