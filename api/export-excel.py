from http.server import BaseHTTPRequestHandler
import json, os, sys, tempfile, traceback

IMPORT_ERROR = None
try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, NamedStyle
    from openpyxl.utils import get_column_letter
    from openpyxl.drawing.image import Image as XLImage
    from openpyxl.formatting.rule import CellIsRule
    from PIL import Image as PILImage
except Exception:
    IMPORT_ERROR = traceback.format_exc()

# Premium color palette
COLOR_BLUE = '2563EB'
COLOR_TEXT = '111827'
COLOR_TEXT_SECONDARY = '6B7280'
COLOR_BORDER = 'E5E7EB'
COLOR_ROW_ALT = 'F9FAFB'
COLOR_HEADER_BG = 'F3F4F6'
COLOR_WHITE = 'FFFFFF'
COLOR_KPI_BG = 'FAFBFC'


def find_logo():
    candidates = [
        os.path.join(os.path.dirname(__file__), 'logo.png'),
        os.path.join(os.path.dirname(__file__), '..', 'public', 'logo.png'),
        '/var/task/api/logo.png',
        '/var/task/public/logo.png',
    ]
    for p in candidates:
        if os.path.exists(p):
            return p
    return None


def build_workbook(gastos):
    wb = Workbook()
    ws = wb.active
    ws.title = 'Gastos'

    # Hide gridlines and headers for clean look
    ws.sheet_view.showGridLines = False
    ws.sheet_view.showRowColHeaders = False

    # Set column widths (more generous spacing)
    widths = {
        'A': 2, 'B': 16, 'C': 28, 'D': 18, 'E': 13, 'E': 13,
        'F': 32, 'G': 14, 'H': 14, 'I': 14, 'J': 14, 'K': 18,
        'L': 13, 'M': 13, 'N': 2
    }
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    # Thin border style
    thin_border = Border(
        bottom=Side(style='thin', color=COLOR_BORDER)
    )

    # ===== ROW 1-5: HEADER =====
    ws.row_dimensions[2].height = 50
    ws.row_dimensions[3].height = 8

    # Logo (will insert at end)

    # Title
    ws['C2'] = 'Reporte de Gastos'
    ws['C2'].font = Font(name='Inter', size=22, bold=True, color=COLOR_TEXT)
    ws['C2'].alignment = Alignment(horizontal='left', vertical='center')

    ws['C3'] = 'SMTO Engineering · Cuenta de Gastos y Viajes'
    ws['C3'].font = Font(name='Inter', size=11, color=COLOR_TEXT_SECONDARY)
    ws['C3'].alignment = Alignment(horizontal='left', vertical='top')

    # ===== ROW 5: KPI CARDS =====
    ws.row_dimensions[5].height = 70

    total_facturado = sum(g.get('totalCFDI', 0) + g.get('montoPropina', 0) for g in gastos)
    total_iva = sum(g.get('iva', 0) for g in gastos)
    total_ret = sum(g.get('retenciones', 0) for g in gastos)
    num_facturas = len(gastos)

    kpis = [
        ('B5:D5', 'TOTAL FACTURADO', f'${total_facturado:,.2f}', COLOR_BLUE),
        ('E5:G5', 'FACTURAS', f'{num_facturas}', COLOR_TEXT),
        ('H5:J5', 'IVA TOTAL', f'${total_iva:,.2f}', COLOR_TEXT),
        ('K5:M5', 'RETENCIONES', f'${total_ret:,.2f}', COLOR_TEXT),
    ]

    for cell_range, label, value, accent_color in kpis:
        ws.merge_cells(cell_range)
        first_cell = cell_range.split(':')[0]
        cell = ws[first_cell]
        cell.value = f'{label}\n{value}'
        cell.font = Font(name='Inter', size=10, color=COLOR_TEXT_SECONDARY)
        cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True, indent=1)
        cell.fill = PatternFill('solid', start_color=COLOR_KPI_BG)
        # Border with accent on left
        cell.border = Border(
            left=Side(style='thick', color=accent_color),
            top=Side(style='thin', color=COLOR_BORDER),
            right=Side(style='thin', color=COLOR_BORDER),
            bottom=Side(style='thin', color=COLOR_BORDER)
        )

    # Spacer
    ws.row_dimensions[6].height = 14

    # ===== ROW 7: TABLE HEADERS =====
    headers = ['#', 'RFC', 'PROVEEDOR', 'FACTURA', 'FECHA', 'CONCEPTO',
               'TIPO', 'IMPORTE', 'IVA', 'RET.', 'TOTAL', 'FORMA PAGO', 'F. COBRO']

    ws.row_dimensions[7].height = 32
    for i, h in enumerate(headers, start=1):
        cell = ws.cell(row=7, column=i, value=h)
        cell.font = Font(name='Inter', size=9, bold=True, color=COLOR_TEXT_SECONDARY)
        cell.alignment = Alignment(horizontal='left' if i in (2, 3, 4, 5, 6, 7, 12, 13) else 'right',
                                   vertical='center', indent=1)
        cell.fill = PatternFill('solid', start_color=COLOR_HEADER_BG)
        cell.border = Border(bottom=Side(style='medium', color=COLOR_TEXT))

    # ===== DATA ROWS =====
    FORMA_PAGO_MAP = {
        '01': '01 · Efectivo',
        '02': '02 · Efectivo',
        '03': '03 · Transferencia',
        '04': '04 · Tarjeta Crédito',
    }

    row = 8
    for idx, g in enumerate(gastos, start=1):
        ws.row_dimensions[row].height = 28
        is_alt = (idx % 2 == 0)
        bg = COLOR_ROW_ALT if is_alt else COLOR_WHITE

        # Format date
        f = g.get('fechaFac', '').split('-')
        fecha_mx = f'{f[2]}/{f[1]}/{f[0]}' if len(f) == 3 else g.get('fechaFac', '')
        fc = g.get('fechaCobro', '').split('-')
        fecha_cobro_mx = f'{fc[2]}/{fc[1]}/{fc[0]}' if len(fc) == 3 else g.get('fechaCobro', '')

        forma = FORMA_PAGO_MAP.get(g.get('formaPago', '04'), g.get('formaPago', ''))
        importe = round(g.get('importe', 0), 2)
        iva = round(g.get('iva', 0), 2)
        ret = round(g.get('retenciones', 0), 2)
        total = round(g.get('totalCFDI', 0) + g.get('montoPropina', 0), 2)
        tipo = g.get('tipo', 'Factura')

        cells_data = [
            (1, idx, 'right'),
            (2, g.get('rfc', ''), 'left'),
            (3, g.get('proveedor', ''), 'left'),
            (4, g.get('noFactura', ''), 'left'),
            (5, fecha_mx, 'left'),
            (6, g.get('concepto', ''), 'left'),
            (7, tipo, 'left'),
            (8, importe, 'right'),
            (9, iva, 'right'),
            (10, ret, 'right'),
            (11, total, 'right'),
            (12, forma, 'left'),
            (13, fecha_cobro_mx, 'left'),
        ]

        for col, val, align in cells_data:
            cell = ws.cell(row=row, column=col, value=val)
            cell.fill = PatternFill('solid', start_color=bg)
            cell.alignment = Alignment(horizontal=align, vertical='center', indent=1)
            cell.border = Border(bottom=Side(style='thin', color=COLOR_BORDER))

            if col == 1:  # # column
                cell.font = Font(name='Inter', size=10, color=COLOR_TEXT_SECONDARY)
            elif col in (8, 9, 10, 11):  # currency cols
                cell.number_format = '"$"#,##0.00'
                cell.font = Font(name='Inter', size=10,
                                 color=COLOR_TEXT if col != 11 else COLOR_TEXT,
                                 bold=(col == 11))
            else:
                cell.font = Font(name='Inter', size=10, color=COLOR_TEXT)

        row += 1

    # ===== TOTALS ROW =====
    ws.row_dimensions[row].height = 36
    totals_cells = [
        (6, 'TOTAL CUENTA', 'right'),
        (8, round(sum(g.get('importe', 0) for g in gastos), 2), 'right'),
        (9, round(sum(g.get('iva', 0) for g in gastos), 2), 'right'),
        (10, round(sum(g.get('retenciones', 0) for g in gastos), 2), 'right'),
        (11, round(sum(g.get('totalCFDI', 0) + g.get('montoPropina', 0) for g in gastos), 2), 'right'),
    ]

    # Empty cells for the totals row first
    for col in range(1, 14):
        cell = ws.cell(row=row, column=col)
        cell.fill = PatternFill('solid', start_color=COLOR_WHITE)
        cell.border = Border(top=Side(style='medium', color=COLOR_TEXT))

    for col, val, align in totals_cells:
        cell = ws.cell(row=row, column=col, value=val)
        cell.alignment = Alignment(horizontal=align, vertical='center', indent=1)
        cell.font = Font(name='Inter', size=11, bold=True, color=COLOR_TEXT)
        if col >= 8:
            cell.number_format = '"$"#,##0.00'

    # ===== LOGO INSERT =====
    logo_path = find_logo()
    if logo_path:
        try:
            # Resize logo preserving aspect ratio
            pil = PILImage.open(logo_path)
            target_h = 50
            ratio = target_h / pil.height
            target_w = int(pil.width * ratio)
            pil = pil.resize((target_w, target_h), PILImage.LANCZOS)
            with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmp:
                logo_tmp = tmp.name
            pil.save(logo_tmp, 'PNG')

            img = XLImage(logo_tmp)
            img.width = target_w
            img.height = target_h
            img.anchor = 'B2'
            ws.add_image(img)
        except Exception as e:
            print(f'Logo warning: {e}')

    # Footer note
    footer_row = row + 2
    ws.cell(row=footer_row, column=2, value=f'Generado por SMTO · {num_facturas} registros').font = Font(name='Inter', size=9, color=COLOR_TEXT_SECONDARY, italic=True)

    return wb


class handler(BaseHTTPRequestHandler):
    def do_GET(self):
        diag = {
            'ok': IMPORT_ERROR is None,
            'python_version': sys.version,
            'logo_path': find_logo(),
            'import_error': IMPORT_ERROR,
        }
        self.send_response(200)
        self.send_header('Content-Type', 'application/json')
        self.send_header('Access-Control-Allow-Origin', '*')
        self.end_headers()
        self.wfile.write(json.dumps(diag, indent=2).encode())

    def do_POST(self):
        try:
            if IMPORT_ERROR:
                raise RuntimeError(f'Import failed:\n{IMPORT_ERROR}')

            content_length = int(self.headers.get('Content-Length', 0))
            body = self.rfile.read(content_length)
            gastos = json.loads(body)

            wb = build_workbook(gastos)

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
            err = traceback.format_exc()
            self.send_response(500)
            self.send_header('Content-Type', 'application/json')
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            self.wfile.write(json.dumps({'error': str(e), 'trace': err}).encode())

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()
