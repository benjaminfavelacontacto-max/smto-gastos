from http.server import BaseHTTPRequestHandler
import json, os, sys, tempfile, traceback
from datetime import datetime

IMPORT_ERROR = None
try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.drawing.image import Image as XLImage
    from PIL import Image as PILImage
except Exception:
    IMPORT_ERROR = traceback.format_exc()

# SMTO Brand palette
SMTO_BLACK = '050505'
SMTO_GREEN = '59D39B'
SMTO_GREEN_DARK = '41C784'
SMTO_GREEN_SOFT = 'E8F8F0'

# Premium neutrals
BG_PAGE = 'F5F7FA'
WHITE = 'FFFFFF'
TEXT_PRIMARY = '0F172A'
TEXT_SECONDARY = '64748B'
TEXT_MUTED = '94A3B8'
BORDER_LIGHT = 'E2E8F0'
ROW_ALT = 'F8FAFC'
HEADER_BG = 'F1F5F9'

# Badge colors (soft transparent feel)
BADGE_BLUE_BG = 'EFF6FF'
BADGE_BLUE_FG = '1E40AF'
BADGE_GREEN_BG = SMTO_GREEN_SOFT
BADGE_GREEN_FG = '047857'
BADGE_GRAY_BG = 'F1F5F9'
BADGE_GRAY_FG = '475569'
BADGE_AMBER_BG = 'FEF3C7'
BADGE_AMBER_FG = '92400E'
BADGE_PURPLE_BG = 'F3E8FF'
BADGE_PURPLE_FG = '6B21A8'

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
    if not date_str: return ''
    if '-' in date_str and len(date_str) == 10:
        parts = date_str.split('-')
        return f'{parts[1]}-{parts[2]}-{parts[0][2:]}'
    if '/' in date_str:
        parts = date_str.split('/')
        return f'{parts[1]}-{parts[0]}-{parts[2][2:] if len(parts[2])>2 else parts[2]}'
    return date_str

FORMA_PAGO_MAP = {
    '01': 'Efectivo',
    '02': 'Efectivo',
    '03': 'Transferencia',
    '04': 'Tarjeta Crédito',
}

def get_tipo_badge_colors(tipo):
    """Return (bg, fg) for the tipo badge."""
    t = (tipo or '').lower()
    if any(k in t for k in ['hotel', 'avión', 'avion', 'taxi', 'consumo']):
        return BADGE_BLUE_BG, BADGE_BLUE_FG
    if any(k in t for k in ['gasolina', 'caseta', 'estacionamiento', 'manto']):
        return BADGE_AMBER_BG, BADGE_AMBER_FG
    if any(k in t for k in ['it & sw', 'marketing', 'celular']):
        return BADGE_PURPLE_BG, BADGE_PURPLE_FG
    if any(k in t for k in ['rechazada', 'devolución', 'no comprobado']):
        return BADGE_GRAY_BG, BADGE_GRAY_FG
    return BADGE_GREEN_BG, BADGE_GREEN_FG

def s_border(bottom=None, top=None, left=None, right=None):
    sides = {}
    if bottom: sides['bottom'] = Side(style=bottom[0], color=bottom[1])
    if top: sides['top'] = Side(style=top[0], color=top[1])
    if left: sides['left'] = Side(style=left[0], color=left[1])
    if right: sides['right'] = Side(style=right[0], color=right[1])
    return Border(**sides)

def fill_row_bg(ws, row, start_col, end_col, color):
    for c in range(start_col, end_col + 1):
        ws.cell(row=row, column=c).fill = PatternFill('solid', start_color=color)

def build_workbook(gastos):
    wb = Workbook()
    ws = wb.active
    ws.title = 'Reporte SMTO'
    ws.sheet_view.showGridLines = False

    # Column widths — wider for premium feel
    col_widths = {
        'A': 4, 'B': 20, 'C': 36, 'D': 22, 'E': 14, 'F': 14,
        'G': 38, 'H': 22, 'I': 17, 'J': 15, 'K': 15, 'L': 18, 'M': 22, 'N': 4
    }
    for col, w in col_widths.items():
        ws.column_dimensions[col].width = w

    # Paint full background light gray (extra rows so spacer columns past
    # the data table still inherit BG_PAGE).
    for r in range(1, 300):
        fill_row_bg(ws, r, 1, 14, BG_PAGE)

    # ═══ TOP MARGIN (rows 1-3) ═══
    for r in range(1, 4):
        ws.row_dimensions[r].height = 18

    # ═══ HEADER (rows 4-7) ═══
    ws.row_dimensions[4].height = 80  # title (sized for larger logo)
    ws.row_dimensions[5].height = 32  # subtitle
    ws.row_dimensions[6].height = 10  # spacer
    ws.row_dimensions[7].height = 1   # hairline separator

    # Title — starts at col D so the 72px logo can occupy cols B-C
    ws.merge_cells('D4:H4')
    title = ws['D4']
    title.value = 'Reporte de Gastos'
    title.font = Font(name='Aptos', size=30, bold=True, color=SMTO_BLACK)
    title.alignment = Alignment(horizontal='left', vertical='center')
    title.fill = PatternFill('solid', start_color=BG_PAGE)

    # Subtitle
    ws.merge_cells('D5:H5')
    sub = ws['D5']
    sub.value = 'SMTO Engineering · Análisis Financiero'
    sub.font = Font(name='Aptos', size=13, color=TEXT_SECONDARY)
    sub.alignment = Alignment(horizontal='left', vertical='top')
    sub.fill = PatternFill('solid', start_color=BG_PAGE)

    # Right metadata (rows 4-5)
    now = datetime.now()
    ws.merge_cells('J4:M4')
    meta1 = ws['J4']
    meta1.value = now.strftime('%d de %b, %Y · %H:%M')
    meta1.font = Font(name='Aptos', size=11, color=TEXT_SECONDARY)
    meta1.alignment = Alignment(horizontal='right', vertical='center')
    meta1.fill = PatternFill('solid', start_color=BG_PAGE)

    ws.merge_cells('J5:M5')
    meta2 = ws['J5']
    meta2.value = f'{len(gastos)} registros procesados'
    meta2.font = Font(name='Aptos', size=10, color=TEXT_MUTED)
    meta2.alignment = Alignment(horizontal='right', vertical='top')
    meta2.fill = PatternFill('solid', start_color=BG_PAGE)

    # Hairline separator across row 7 (B7:M7)
    for c in range(2, 14):
        cell = ws.cell(row=7, column=c)
        cell.fill = PatternFill('solid', start_color=BG_PAGE)
        cell.border = Border(bottom=Side(style='thin', color=BORDER_LIGHT))

    # ═══ KPI CARDS (rows 8-12) ═══
    ws.row_dimensions[8].height = 24   # spacer above
    ws.row_dimensions[9].height = 62   # KPI VALUES
    ws.row_dimensions[10].height = 26  # KPI LABELS
    ws.row_dimensions[11].height = 8   # thin accent line
    ws.row_dimensions[12].height = 24  # spacer below

    total_facturado = round(sum(g.get('totalCFDI', 0) + g.get('montoPropina', 0) for g in gastos), 2)
    total_iva = round(sum(g.get('iva', 0) for g in gastos), 2)
    total_ret = round(sum(g.get('retenciones', 0) for g in gastos), 2)
    num_facturas = len(gastos)

    # (col_start, col_end, label, value, accent)
    kpis = [
        ('B', 'D', 'TOTAL FACTURADO', f'${total_facturado:,.2f}', SMTO_GREEN),
        ('E', 'G', 'IVA TOTAL',       f'${total_iva:,.2f}',        SMTO_BLACK),
        ('H', 'J', 'RETENCIONES',     f'${total_ret:,.2f}',        SMTO_BLACK),
        ('K', 'M', 'REGISTROS',       str(num_facturas),           SMTO_BLACK),
    ]

    for col_start, col_end, label, value, accent in kpis:
        # Value row (row 9) — big bold accent-colored number, hugs bottom
        ws.merge_cells(f'{col_start}9:{col_end}9')
        v_cell = ws[f'{col_start}9']
        v_cell.value = value
        v_cell.alignment = Alignment(horizontal='left', vertical='bottom', indent=2)
        v_cell.fill = PatternFill('solid', start_color=WHITE)
        v_cell.font = Font(name='Aptos', size=28, bold=True, color=accent)
        v_cell.border = Border(
            left=Side(style='thick', color=accent),
            top=Side(style='thin', color=BORDER_LIGHT),
            right=Side(style='thin', color=BORDER_LIGHT),
            bottom=Side(style='thin', color=BORDER_LIGHT),
        )

        # Label row (row 10) — small uppercase muted label, hugs top
        ws.merge_cells(f'{col_start}10:{col_end}10')
        l_cell = ws[f'{col_start}10']
        l_cell.value = label
        l_cell.alignment = Alignment(horizontal='left', vertical='top', indent=2)
        l_cell.fill = PatternFill('solid', start_color=WHITE)
        l_cell.font = Font(name='Aptos', size=9, bold=True, color=TEXT_MUTED)
        l_cell.border = Border(
            left=Side(style='thick', color=accent),
            right=Side(style='thin', color=BORDER_LIGHT),
            bottom=Side(style='thin', color=BORDER_LIGHT),
        )

    # Full-width hairline under the KPI section (row 11)
    for c in range(2, 14):
        cell = ws.cell(row=11, column=c)
        cell.fill = PatternFill('solid', start_color=BG_PAGE)
        cell.border = Border(bottom=Side(style='thin', color=BORDER_LIGHT))

    # ═══ TABLE HEADER (row 13) ═══
    ws.row_dimensions[13].height = 44

    headers = ['RFC', 'PROVEEDOR', 'FACTURA', 'F. FACTURA', 'F. COBRO', 'CONCEPTO', 'TIPO', 'IMPORTE', 'IVA', 'RETENCIÓN', 'TOTAL', 'FORMA PAGO']

    for i, h in enumerate(headers):
        col = i + 2
        cell = ws.cell(row=13, column=col, value=h)
        cell.font = Font(name='Aptos', size=10, bold=True, color=TEXT_SECONDARY)
        is_num = h in ('IMPORTE', 'IVA', 'RETENCIÓN', 'TOTAL')
        cell.alignment = Alignment(
            horizontal='right' if is_num else 'left',
            vertical='center', indent=2
        )
        cell.fill = PatternFill('solid', start_color=HEADER_BG)
        cell.border = Border(bottom=Side(style='medium', color=SMTO_BLACK))

    # ═══ DATA ROWS (row 14+) ═══
    row = 14
    for idx, g in enumerate(gastos):
        ws.row_dimensions[row].height = 44
        is_alt = (idx % 2 == 1)
        bg = ROW_ALT if is_alt else WHITE

        fecha_fac = format_date(g.get('fechaFac', ''))
        fecha_cobro = format_date(g.get('fechaCobro', ''))
        forma = FORMA_PAGO_MAP.get(g.get('formaPago', '04'), g.get('formaPago', ''))
        importe = round(g.get('importe', 0), 2)
        iva = round(g.get('iva', 0), 2)
        ret = round(g.get('retenciones', 0), 2)
        total = round(g.get('totalCFDI', 0) + g.get('montoPropina', 0), 2)
        tipo = g.get('tipo', 'Consumo')

        cells = [
            (2, g.get('rfc', ''), 'left', 'mono'),
            (3, g.get('proveedor', ''), 'left', 'normal'),
            (4, g.get('noFactura', ''), 'left', 'normal'),
            (5, fecha_fac, 'left', 'normal'),
            (6, fecha_cobro, 'left', 'normal'),
            (7, g.get('concepto', ''), 'left', 'normal'),
            (8, tipo, 'center', 'badge_tipo'),
            (9, importe, 'right', 'currency'),
            (10, iva, 'right', 'currency'),
            (11, ret, 'right', 'currency'),
            (12, total, 'right', 'currency_bold'),
            (13, forma, 'center', 'badge_pago'),
        ]

        for col, val, align, style_type in cells:
            cell = ws.cell(row=row, column=col, value=val)
            cell.fill = PatternFill('solid', start_color=bg)
            cell.border = Border(bottom=Side(style='hair', color=BORDER_LIGHT))
            cell.alignment = Alignment(horizontal=align, vertical='center', indent=2 if align != 'center' else 0)

            if style_type == 'currency':
                cell.number_format = '"$"#,##0.00'
                cell.font = Font(name='Aptos', size=11, color=TEXT_PRIMARY)
            elif style_type == 'currency_bold':
                cell.number_format = '"$"#,##0.00'
                cell.font = Font(name='Aptos', size=12, bold=True, color=SMTO_BLACK)
            elif style_type == 'mono':
                cell.font = Font(name='Aptos', size=10, color=TEXT_SECONDARY)
            elif style_type == 'badge_tipo':
                bg_b, fg_b = get_tipo_badge_colors(tipo)
                cell.fill = PatternFill('solid', start_color=bg_b)
                cell.font = Font(name='Aptos', size=10, bold=True, color=fg_b)
            elif style_type == 'badge_pago':
                cell.fill = PatternFill('solid', start_color=BADGE_GRAY_BG)
                cell.font = Font(name='Aptos', size=10, color=BADGE_GRAY_FG)
            else:
                cell.font = Font(name='Aptos', size=11, color=TEXT_PRIMARY)

        # Side spacer cells keep page bg
        ws.cell(row=row, column=1).fill = PatternFill('solid', start_color=BG_PAGE)
        ws.cell(row=row, column=14).fill = PatternFill('solid', start_color=BG_PAGE)

        row += 1

    # ═══ SPACER BEFORE TOTALS — two empty rows of height 16 ═══
    ws.row_dimensions[row].height = 16
    row += 1
    ws.row_dimensions[row].height = 16
    row += 1

    # ═══ TOTALS ROW ═══
    ws.row_dimensions[row].height = 60
    for c in range(2, 14):
        cell = ws.cell(row=row, column=c)
        cell.fill = PatternFill('solid', start_color=SMTO_BLACK)
        cell.border = Border()

    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=8)
    lbl = ws.cell(row=row, column=2, value='TOTAL CUENTA')
    lbl.font = Font(name='Aptos', size=14, bold=True, color=WHITE)
    lbl.alignment = Alignment(horizontal='right', vertical='center', indent=2)
    lbl.fill = PatternFill('solid', start_color=SMTO_BLACK)

    totals = [
        (9, round(sum(g.get('importe', 0) for g in gastos), 2), False),
        (10, round(sum(g.get('iva', 0) for g in gastos), 2), False),
        (11, round(sum(g.get('retenciones', 0) for g in gastos), 2), False),
        (12, total_facturado, True),
    ]
    for col, val, is_main in totals:
        cell = ws.cell(row=row, column=col, value=val)
        cell.number_format = '"$"#,##0.00'
        cell.font = Font(name='Aptos', size=22 if is_main else 13, bold=is_main,
                          color=SMTO_GREEN if is_main else WHITE)
        cell.alignment = Alignment(horizontal='right', vertical='center', indent=2)
        cell.fill = PatternFill('solid', start_color=SMTO_BLACK)

    ws.cell(row=row, column=13).fill = PatternFill('solid', start_color=SMTO_BLACK)

    # ═══ FOOTER ═══
    # spacer (16) → hairline → spacer (10) → footer (28)
    row += 1
    ws.row_dimensions[row].height = 16  # spacer above hairline
    row += 1
    ws.row_dimensions[row].height = 1   # hairline separator
    for c in range(2, 14):
        cell = ws.cell(row=row, column=c)
        cell.fill = PatternFill('solid', start_color=BG_PAGE)
        cell.border = Border(bottom=Side(style='thin', color=BORDER_LIGHT))
    row += 1
    ws.row_dimensions[row].height = 10  # spacer below hairline
    row += 1
    ws.row_dimensions[row].height = 28  # footer
    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=7)
    f1 = ws.cell(row=row, column=2)
    f1.value = f'Generado automáticamente · {now.strftime("%d/%m/%Y %H:%M")}'
    f1.font = Font(name='Aptos', size=9, italic=True, color=TEXT_MUTED)
    f1.alignment = Alignment(horizontal='left', vertical='center')
    f1.fill = PatternFill('solid', start_color=BG_PAGE)

    ws.merge_cells(start_row=row, start_column=10, end_row=row, end_column=13)
    f2 = ws.cell(row=row, column=10)
    f2.value = f'SMTO Engineering · v4.7'
    f2.font = Font(name='Aptos', size=9, italic=True, color=TEXT_MUTED)
    f2.alignment = Alignment(horizontal='right', vertical='center')
    f2.fill = PatternFill('solid', start_color=BG_PAGE)

    # ═══ LOGO ═══
    logo_path = find_logo()
    if logo_path:
        try:
            pil = PILImage.open(logo_path)
            target_h = 72
            ratio = target_h / pil.height
            target_w = int(pil.width * ratio)
            pil = pil.resize((target_w, target_h), PILImage.LANCZOS)
            with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmp:
                logo_tmp = tmp.name
            pil.save(logo_tmp, 'PNG')
            img = XLImage(logo_tmp)
            img.width = target_w
            img.height = target_h
            img.anchor = 'B4'
            ws.add_image(img)
        except Exception as e:
            print(f'Logo: {e}')

    return wb

class handler(BaseHTTPRequestHandler):
    def do_GET(self):
        diag = {'ok': IMPORT_ERROR is None, 'python_version': sys.version, 'logo_path': find_logo(), 'import_error': IMPORT_ERROR}
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
