from http.server import BaseHTTPRequestHandler
import json, os, sys, tempfile, traceback
import xlrd
from xlutils.copy import copy as xl_copy


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
    def do_POST(self):
        try:
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
        self.send_header('Access-Control-Allow-Methods', 'POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()
