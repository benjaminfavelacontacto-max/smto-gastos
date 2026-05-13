from http.server import BaseHTTPRequestHandler
import json, os, sys, tempfile, traceback
from datetime import datetime

IMPORT_ERROR = None
try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, numbers
    from openpyxl.utils import get_column_letter
    from openpyxl.drawing.image import Image as XLImage
    from PIL import Image as PILImage
except Exception:
    IMPORT_ERROR = traceback.format_exc()

# Ultra premium palette
BLUE = '2563EB'
BLUE_LIGHT = 'EFF6FF'
BLUE_ACCENT = 'DBEAFE'
TEXT_PRIMARY = '111827'
TEXT_SECONDARY = '6B7280'
TEXT_MUTED = '9CA3AF'
BORDER_LIGHT = 'E5E7EB'
BORDER_MEDIUM = 'D1D5DB'
ROW_ALT = 'F9FAFB'
WHITE = 'FFFFFF'
KPI_BG = 'F8FAFC'
HEADER_BG = 'F1F5F9'
GREEN_ACCENT = '059669'
GREEN_BG = 'ECFDF5'


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


def format_date(date_str):
    if not date_str:
        return ''
    if '-' in date_str and len(date_str) == 10:
        parts = date_str.split('-')
        return f'{parts[1]}-{parts[2]}-{parts[0][2:]}'
    if '/' in date_str:
        parts = date_str.split('/')
        return f'{parts[1]}-{parts[0]}-{parts[2][2:] if len(parts[2]) > 2 else parts[2]}'
    return date_str


FORMA_PAGO_MAP = {
    '01': 'Efectivo',
    '02': 'Efectivo',
    '03': 'Transferencia',
    '04': 'Tarjeta Crédito',
}


def s_border(bottom=None, top=None, left=None, right=None):
    sides = {}
    if bottom: sides['bottom'] = Side(style=bottom[0], color=bottom[1])
    if top:    sides['top']    = Side(style=top[0],    color=top[1])
    if left:   sides['left']   = Side(style=left[0],   color=left[1])
    if right:  sides['right']  = Side(style=right[0],  color=right[1])
    return Border(**sides)


def build_workbook(gastos):
    wb = Workbook()
    ws = wb.active
    ws.title = 'Reporte de Gastos'
    ws.sheet_view.showGridLines = False

    # Print settings
    ws.sheet_properties.pageSetUpPr = None

    # Column widths — generous premium spacing
    # A=spacer, B=RFC, C=Proveedor, D=Factura, E=F.Factura, F=F.Cobro, G=Concepto, H=Tipo, I=Importe, J=IVA, K=Ret, L=Total, M=FormaPago, N=spacer
    col_widths = {'A': 3, 'B': 17, 'C': 30, 'D': 20, 'E': 12, 'F': 12, 'G': 34, 'H': 12, 'I': 15, 'J': 13, 'K': 13, 'L': 16, 'M': 18, 'N': 3}
    for col, w in col_widths.items():
        ws.column_dimensions[col].width = w

    # ════════════════════════════════════════════
    # SECTION 1: PREMIUM HEADER (rows 1-4)
    # ════════════════════════════════════════════
    for r in range(1, 5):
        ws.row_dimensions[r].height = 6
        for c in range(1, 15):
            ws.cell(row=r, column=c).fill = PatternFill('solid', start_color=WHITE)

    # Row 2: Logo + Title
    ws.row_dimensions[2].height = 45
    ws.merge_cells('C2:F2')
    title = ws['C2']
    title.value = 'Reporte de Gastos'
    title.font = Font(name='Aptos', size=20, bold=True, color=TEXT_PRIMARY)
    title.alignment = Alignment(horizontal='left', vertical='center')

    # Row 3: Subtitle + metadata
    ws.row_dimensions[3].height = 20
    ws.merge_cells('C3:F3')
    sub = ws['C3']
    sub.value = 'SMTO Engineering'
    sub.font = Font(name='Aptos', size=10, color=TEXT_SECONDARY)
    sub.alignment = Alignment(horizontal='left', vertical='top')

    # Right side metadata
    ws.merge_cells('J2:M2')
    meta1 = ws['J2']
    now = datetime.now()
    meta1.value = now.strftime('%d %b %Y · %H:%M')
    meta1.font = Font(name='Aptos', size=9, color=TEXT_MUTED)
    meta1.alignment = Alignment(horizontal='right', vertical='center')

    ws.merge_cells('J3:M3')
    meta2 = ws['J3']
    meta2.value = f'{len(gastos)} registros procesados'
    meta2.font = Font(name='Aptos', size=9, color=TEXT_MUTED)
    meta2.alignment = Alignment(horizontal='right', vertical='top')

    # Row 4: Thin separator
    ws.row_dimensions[4].height = 2
    for c in range(2, 14):
        ws.cell(row=4, column=c).border = s_border(bottom=('thin', BORDER_LIGHT))

    # ════════════════════════════════════════════
    # SECTION 2: KPI CARDS (rows 5-7)
    # ════════════════════════════════════════════
    ws.row_dimensions[5].height = 12  # spacer
    ws.row_dimensions[6].height = 52
    ws.row_dimensions[7].height = 12  # spacer

    total_facturado = round(sum(g.get('totalCFDI', 0) + g.get('montoPropina', 0) for g in gastos), 2)
    total_iva = round(sum(g.get('iva', 0) for g in gastos), 2)
    total_ret = round(sum(g.get('retenciones', 0) for g in gastos), 2)
    num_facturas = len(gastos)

    # KPI card positions: each spans 3 columns
    kpi_data = [
        ('B6:D6', 'TOTAL FACTURADO', f'${total_facturado:,.2f}', BLUE, BLUE_LIGHT),
        ('E6:G6', 'IVA TOTAL', f'${total_iva:,.2f}', TEXT_PRIMARY, KPI_BG),
        ('H6:J6', 'RETENCIONES', f'${total_ret:,.2f}', TEXT_PRIMARY, KPI_BG),
        ('K6:M6', 'FACTURAS', str(num_facturas), GREEN_ACCENT, GREEN_BG),
    ]

    for rng, label, value, accent, bg_color in kpi_data:
        ws.merge_cells(rng)
        first = rng.split(':')[0]
        cell = ws[first]
        cell.value = f'  {value}\n  {label}'
        cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
        cell.fill = PatternFill('solid', start_color=bg_color)
        cell.font = Font(name='Aptos', size=11, color=accent)
        # Subtle card border with left accent
        cell.border = Border(
            left=Side(style='medium', color=accent),
            top=Side(style='thin', color=BORDER_LIGHT),
            right=Side(style='thin', color=BORDER_LIGHT),
            bottom=Side(style='thin', color=BORDER_LIGHT),
        )

    # ════════════════════════════════════════════
    # SECTION 3: TABLE HEADER (row 8)
    # ════════════════════════════════════════════
    ws.row_dimensions[8].height = 30

    headers = ['RFC', 'PROVEEDOR', 'FACTURA', 'F. FACTURA', 'F. COBRO', 'CONCEPTO', 'TIPO', 'IMPORTE', 'IVA', 'RETENCIÓN', 'TOTAL', 'FORMA PAGO']

    for i, h in enumerate(headers):
        col = i + 2  # start at column B
        cell = ws.cell(row=8, column=col, value=h)
        cell.font = Font(name='Aptos', size=8.5, bold=True, color=TEXT_SECONDARY)
        is_num = h in ('IMPORTE', 'IVA', 'RETENCIÓN', 'TOTAL')
        cell.alignment = Alignment(
            horizontal='right' if is_num else 'left',
            vertical='center',
            indent=1
        )
        cell.fill = PatternFill('solid', start_color=HEADER_BG)
        cell.border = Border(
            bottom=Side(style='thin', color=BORDER_MEDIUM),
            top=Side(style='thin', color=BORDER_LIGHT),
        )

    # ════════════════════════════════════════════
    # SECTION 4: DATA ROWS (row 9+)
    # ════════════════════════════════════════════
    row = 9
    for idx, g in enumerate(gastos):
        ws.row_dimensions[row].height = 30
        is_alt = (idx % 2 == 1)
        bg = ROW_ALT if is_alt else WHITE

        fecha_fac = format_date(g.get('fechaFac', ''))
        fecha_cobro = format_date(g.get('fechaCobro', ''))
        forma = FORMA_PAGO_MAP.get(g.get('formaPago', '04'), g.get('formaPago', ''))
        importe = round(g.get('importe', 0), 2)
        iva = round(g.get('iva', 0), 2)
        ret = round(g.get('retenciones', 0), 2)
        total = round(g.get('totalCFDI', 0) + g.get('montoPropina', 0), 2)
        tipo = g.get('tipo', 'Factura')

        data = [
            (2,  g.get('rfc', ''),       False),
            (3,  g.get('proveedor', ''), False),
            (4,  g.get('noFactura', ''), False),
            (5,  fecha_fac,              False),
            (6,  fecha_cobro,            False),
            (7,  g.get('concepto', ''),  False),
            (8,  tipo,                   False),
            (9,  importe,                True),
            (10, iva,                    True),
            (11, ret,                    True),
            (12, total,                  True),
            (13, forma,                  False),
        ]

        for col, val, is_currency in data:
            cell = ws.cell(row=row, column=col, value=val)
            cell.fill = PatternFill('solid', start_color=bg)
            cell.border = s_border(bottom=('hair', BORDER_LIGHT))

            if is_currency:
                cell.number_format = '"$"#,##0.00'
                cell.alignment = Alignment(horizontal='right', vertical='center', indent=1)
                if col == 12:  # TOTAL column bold
                    cell.font = Font(name='Aptos', size=10, bold=True, color=TEXT_PRIMARY)
                else:
                    cell.font = Font(name='Aptos', size=10, color=TEXT_PRIMARY)
            elif col == 2:  # RFC monospace feel
                cell.font = Font(name='Aptos', size=9, color=TEXT_SECONDARY)
                cell.alignment = Alignment(horizontal='left', vertical='center', indent=1)
            elif col == 3:  # Proveedor
                cell.font = Font(name='Aptos', size=10, color=TEXT_PRIMARY)
                cell.alignment = Alignment(horizontal='left', vertical='center', indent=1)
            elif col == 8:  # Tipo - small pill style
                cell.font = Font(name='Aptos', size=9, color=TEXT_SECONDARY)
                cell.alignment = Alignment(horizontal='center', vertical='center')
            else:
                cell.font = Font(name='Aptos', size=10, color=TEXT_PRIMARY)
                cell.alignment = Alignment(horizontal='left', vertical='center', indent=1)

        # Spacer columns
        for sc in [1, 14]:
            ws.cell(row=row, column=sc).fill = PatternFill('solid', start_color=WHITE)

        row += 1

    # ════════════════════════════════════════════
    # SECTION 5: TOTALS ROW (premium footer)
    # ════════════════════════════════════════════
    # Thin separator line above totals
    for c in range(2, 14):
        ws.cell(row=row, column=c).border = s_border(top=('medium', TEXT_PRIMARY))
        ws.cell(row=row, column=c).fill = PatternFill('solid', start_color=WHITE)

    ws.row_dimensions[row].height = 38

    # "TOTAL CUENTA" label
    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=8)
    total_label = ws.cell(row=row, column=2, value='TOTAL CUENTA')
    total_label.font = Font(name='Aptos', size=11, bold=True, color=TEXT_PRIMARY)
    total_label.alignment = Alignment(horizontal='right', vertical='center', indent=2)
    total_label.fill = PatternFill('solid', start_color=WHITE)

    # Total values
    totals = [
        (9,  round(sum(g.get('importe', 0) for g in gastos), 2)),
        (10, round(sum(g.get('iva', 0) for g in gastos), 2)),
        (11, round(sum(g.get('retenciones', 0) for g in gastos), 2)),
        (12, total_facturado),
    ]

    for col, val in totals:
        cell = ws.cell(row=row, column=col, value=val)
        cell.number_format = '"$"#,##0.00'
        cell.font = Font(name='Aptos', size=11, bold=True, color=BLUE if col == 12 else TEXT_PRIMARY)
        cell.alignment = Alignment(horizontal='right', vertical='center', indent=1)
        cell.fill = PatternFill('solid', start_color=WHITE)

    ws.cell(row=row, column=13).fill = PatternFill('solid', start_color=WHITE)

    # ════════════════════════════════════════════
    # SECTION 6: FOOTER
    # ════════════════════════════════════════════
    row += 1
    ws.row_dimensions[row].height = 8
    row += 1
    ws.row_dimensions[row].height = 4
    for c in range(2, 14):
        ws.cell(row=row, column=c).border = s_border(bottom=('hair', BORDER_LIGHT))

    row += 1
    ws.row_dimensions[row].height = 20
    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=7)
    footer = ws.cell(row=row, column=2)
    footer.value = f'Generado automáticamente · SMTO Engineering · {now.strftime("%d/%m/%Y %H:%M")}'
    footer.font = Font(name='Aptos', size=8, italic=True, color=TEXT_MUTED)
    footer.alignment = Alignment(horizontal='left', vertical='center')

    ws.merge_cells(start_row=row, start_column=10, end_row=row, end_column=13)
    footer_r = ws.cell(row=row, column=10)
    footer_r.value = f'{len(gastos)} registros · v4.1'
    footer_r.font = Font(name='Aptos', size=8, italic=True, color=TEXT_MUTED)
    footer_r.alignment = Alignment(horizontal='right', vertical='center')

    # ════════════════════════════════════════════
    # LOGO INSERT (small, elegant)
    # ════════════════════════════════════════════
    logo_path = find_logo()
    if logo_path:
        try:
            pil = PILImage.open(logo_path)
            target_h = 38
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
            print(f'Logo: {e}')

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
