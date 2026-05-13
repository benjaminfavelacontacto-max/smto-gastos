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
BORDER_MEDIUM = '94A3B8'
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

def crop_logo(pil_img):
    """Remove excess whitespace/transparent padding around the logo so it
    visually aligns with surrounding type, regardless of the source PNG's
    bounding box."""
    try:
        import numpy as np
        rgba = pil_img.convert('RGBA')
        data = np.array(rgba)
        # Treat any non-transparent pixel as logo content.
        mask = data[:, :, 3] > 30
        rows = np.any(mask, axis=1)
        cols = np.any(mask, axis=0)
        if rows.any() and cols.any():
            rmin, rmax = np.where(rows)[0][[0, -1]]
            cmin, cmax = np.where(cols)[0][[0, -1]]
            pad = 4
            rmin = max(0, rmin - pad); rmax = min(data.shape[0], rmax + pad)
            cmin = max(0, cmin - pad); cmax = min(data.shape[1], cmax + pad)
            return pil_img.crop((cmin, rmin, cmax, rmax))
    except Exception:
        pass
    return pil_img

def compute_date_range(gastos):
    """Min/max of fechaCobro across gastos (falls back to fechaFac when a
    row has no cobro date). Returns ('', '') when no parseable dates exist.
    Output format: MM-DD-YY."""
    dates = []
    for g in gastos:
        for key in ('fechaCobro', 'fechaFac'):
            v = g.get(key, '')
            if v and '-' in v and len(v) == 10:
                dates.append(v)  # YYYY-MM-DD
                break
    if not dates:
        return '', ''
    dates.sort()
    fmt = lambda d: f'{d[5:7]}-{d[8:10]}-{d[2:4]}'
    return fmt(dates[0]), fmt(dates[-1])

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
    '01': '01 - Efectivo',
    '02': '02 - Efectivo',
    '03': '03 - Transferencia',
    '04': '04 - Tarjeta de Crédito',
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

    # Column widths — semantic (wide CONCEPTO + supplier, narrow dates) so
    # each column gets the width that fits its content, regardless of order.
    col_widths = {
        'A': 5, 'B': 22, 'C': 43, 'D': 22, 'E': 23, 'F': 16, 'G': 14.5,
        'H': 42, 'I': 18, 'J': 16, 'K': 16, 'L': 20, 'M': 24, 'N': 5
    }
    for col, w in col_widths.items():
        ws.column_dimensions[col].width = w

    # Paint full background light gray — scales with row count so the
    # outer spacer cols past the totals/footer still inherit BG_PAGE.
    nrows_painted = max(80, 40 + len(gastos))
    for r in range(1, nrows_painted):
        fill_row_bg(ws, r, 1, 14, BG_PAGE)

    # ═══ HEADER (rows 1-2) — title + colaborador labels + fields ═══
    ws.row_dimensions[1].height = 54
    ws.row_dimensions[2].height = 36

    # Title — centered across D1:G2; pairs with the 60px logo at B1
    ws.merge_cells('D1:G2')
    title = ws['D1']
    title.value = 'Reporte de Gastos'
    title.font = Font(name='Aptos', size=36, bold=True, color=SMTO_BLACK)
    title.alignment = Alignment(horizontal='center', vertical='center')
    title.fill = PatternFill('solid', start_color=BG_PAGE)

    # Right-side form labels (col H) — larger, non-bold per reference
    lbl1 = ws['H1']
    lbl1.value = 'Nombre de colaborador:'
    lbl1.font = Font(name='Aptos', size=18, color=TEXT_SECONDARY)
    lbl1.alignment = Alignment(horizontal='right', vertical='center')
    lbl1.fill = PatternFill('solid', start_color=BG_PAGE)

    lbl2 = ws['H2']
    lbl2.value = 'Fecha de viaje:'
    lbl2.font = Font(name='Aptos', size=18, color=TEXT_SECONDARY)
    lbl2.alignment = Alignment(horizontal='right', vertical='center')
    lbl2.fill = PatternFill('solid', start_color=BG_PAGE)

    # Right-side form fields (cols J:M) — white fill, font 14
    ws.merge_cells('J1:M1')
    f1 = ws['J1']
    f1.value = ''  # left blank for the user to fill in
    f1.font = Font(name='Aptos', size=14, color=TEXT_PRIMARY)
    f1.fill = PatternFill('solid', start_color=WHITE)
    f1.border = Border(bottom=Side(style='thin', color=BORDER_LIGHT))
    f1.alignment = Alignment(horizontal='left', vertical='center', indent=1)

    ws.merge_cells('J2:M2')
    f2 = ws['J2']
    fecha_min, fecha_max = compute_date_range(gastos)
    f2.value = f'DE: {fecha_min}   A: {fecha_max}' if fecha_min else ''
    f2.font = Font(name='Aptos', size=14, color=TEXT_PRIMARY)
    f2.fill = PatternFill('solid', start_color=WHITE)
    f2.alignment = Alignment(horizontal='left', vertical='center', indent=1)

    # Row 3 spacer + Row 4 hairline (1pt row with bottom border)
    ws.row_dimensions[3].height = 18
    ws.row_dimensions[4].height = 1
    for c in range(2, 14):
        cell = ws.cell(row=4, column=c)
        cell.fill = PatternFill('solid', start_color=BG_PAGE)
        cell.border = Border(bottom=Side(style='thin', color=BORDER_LIGHT))

    # ═══ KPI CARDS (rows 5-6) ═══
    ws.row_dimensions[5].height = 46  # KPI VALUES
    ws.row_dimensions[6].height = 23  # KPI LABELS

    total_facturado = round(sum(g.get('totalCFDI', 0) + g.get('montoPropina', 0) for g in gastos), 2)
    total_iva = round(sum(g.get('iva', 0) for g in gastos), 2)
    total_ret = round(sum(g.get('retenciones', 0) for g in gastos), 2)
    num_facturas = len(gastos)

    # (col_start, col_end, label, value, value_color)
    # All four KPIs share a thick SMTO_GREEN left border per the reference;
    # only the first value is colored green, the rest stay SMTO_BLACK.
    kpis = [
        ('B', 'D', 'TOTAL FACTURADO', f'${total_facturado:,.2f}', SMTO_GREEN),
        ('E', 'G', 'IVA TOTAL',       f'${total_iva:,.2f}',       SMTO_BLACK),
        ('H', 'J', 'RETENCIONES',     f'${total_ret:,.2f}',       SMTO_BLACK),
        ('K', 'M', 'REGISTROS',       str(num_facturas),          SMTO_BLACK),
    ]

    for col_start, col_end, label, value, value_color in kpis:
        # Value row (row 5)
        ws.merge_cells(f'{col_start}5:{col_end}5')
        v_cell = ws[f'{col_start}5']
        v_cell.value = value
        v_cell.font = Font(name='Aptos', size=26, bold=True, color=value_color)
        v_cell.alignment = Alignment(horizontal='left', vertical='center', indent=2)
        v_cell.fill = PatternFill('solid', start_color=WHITE)
        v_cell.border = Border(
            left=Side(style='thick', color=SMTO_GREEN),
            top=Side(style='thin', color=BORDER_LIGHT),
            right=Side(style='thin', color=BORDER_LIGHT),
            bottom=Side(style='thin', color=BORDER_LIGHT),
        )

        # Label row (row 6)
        ws.merge_cells(f'{col_start}6:{col_end}6')
        l_cell = ws[f'{col_start}6']
        l_cell.value = label
        l_cell.font = Font(name='Aptos', size=14, bold=True, color=TEXT_MUTED)
        l_cell.alignment = Alignment(horizontal='left', vertical='center', indent=2)
        l_cell.fill = PatternFill('solid', start_color=WHITE)
        l_cell.border = Border(
            left=Side(style='thick', color=SMTO_GREEN),
            right=Side(style='thin', color=BORDER_LIGHT),
            bottom=Side(style='thin', color=BORDER_LIGHT),
        )

    # Row 7 small gap + Row 8 larger pre-table spacer
    ws.row_dimensions[7].height = 4
    ws.row_dimensions[8].height = 28

    # ═══ TABLE HEADER (row 9) — green text, mostly centered ═══
    ws.row_dimensions[9].height = 48

    headers = ['RFC', 'PROVEEDOR', 'TIPO', 'FACTURA', 'F. FACTURA', 'F. COBRO', 'CONCEPTO', 'IMPORTE', 'IVA', 'RETENCIÓN', 'TOTAL', 'FORMA PAGO']
    # PROVEEDOR and CONCEPTO stay left-aligned; the rest center.
    left_align_headers = {'PROVEEDOR', 'CONCEPTO'}

    for i, h in enumerate(headers):
        col = i + 2
        cell = ws.cell(row=9, column=col, value=h)
        cell.font = Font(name='Aptos', size=14, bold=True, color=SMTO_GREEN)
        is_left = h in left_align_headers
        cell.alignment = Alignment(
            horizontal='left' if is_left else 'center',
            vertical='center',
            indent=2 if is_left else 0
        )
        cell.fill = PatternFill('solid', start_color=HEADER_BG)
        cell.border = Border(bottom=Side(style='medium', color=SMTO_BLACK))

    # ═══ DATA ROWS (row 10+) ═══
    row = 10
    for idx, g in enumerate(gastos):
        ws.row_dimensions[row].height = 48
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

        # Column order matches the headers. PROVEEDOR and CONCEPTO are the only
        # left-aligned cells; everything else centers per the reference.
        cells = [
            (2,  g.get('rfc', ''),       'left',   'normal_bold'),
            (3,  g.get('proveedor', ''), 'left',   'normal_bold'),
            (4,  tipo,                   'center', 'badge_tipo'),
            (5,  g.get('noFactura', ''), 'center', 'normal'),
            (6,  fecha_fac,              'center', 'normal'),
            (7,  fecha_cobro,            'center', 'normal'),
            (8,  g.get('concepto', ''),  'left',   'normal'),
            (9,  importe,                'center', 'currency'),
            (10, iva,                    'center', 'currency'),
            (11, ret,                    'center', 'currency'),
            (12, total,                  'center', 'currency_bold'),
            (13, forma,                  'center', 'badge_pago'),
        ]

        for col, val, align, style_type in cells:
            cell = ws.cell(row=row, column=col, value=val)
            cell.fill = PatternFill('solid', start_color=bg)
            cell.border = Border(bottom=Side(style='hair', color=BORDER_LIGHT))
            cell.alignment = Alignment(
                horizontal=align,
                vertical='center',
                indent=2 if align == 'left' else 0
            )

            if style_type == 'currency':
                cell.number_format = '"$"#,##0.00'
                cell.font = Font(name='Aptos', size=13, color=TEXT_PRIMARY)
            elif style_type == 'currency_bold':
                cell.number_format = '"$"#,##0.00'
                cell.font = Font(name='Aptos', size=14, bold=True, color=SMTO_BLACK)
            elif style_type == 'normal_bold':
                cell.font = Font(name='Aptos', size=11, bold=True, color=TEXT_PRIMARY)
            elif style_type == 'badge_tipo':
                bg_b, fg_b = get_tipo_badge_colors(tipo)
                cell.fill = PatternFill('solid', start_color=bg_b)
                cell.font = Font(name='Aptos', size=11, bold=True, color=fg_b)
            elif style_type == 'badge_pago':
                cell.fill = PatternFill('solid', start_color=BADGE_GRAY_BG)
                cell.font = Font(name='Aptos', size=11, bold=True, color=BADGE_GRAY_FG)
            else:  # 'normal'
                cell.font = Font(name='Aptos', size=13, color=TEXT_PRIMARY)

        # Side spacer cells keep page bg through the data band.
        ws.cell(row=row, column=1).fill = PatternFill('solid', start_color=BG_PAGE)
        ws.cell(row=row, column=14).fill = PatternFill('solid', start_color=BG_PAGE)

        row += 1

    # ═══ ADAPTIVE TOTALS — placed right after the last data row ═══
    ws.row_dimensions[row].height = 14  # spacer
    row += 1

    ws.row_dimensions[row].height = 44
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
        (9,  round(sum(g.get('importe', 0)     for g in gastos), 2), False),
        (10, round(sum(g.get('iva', 0)         for g in gastos), 2), False),
        (11, round(sum(g.get('retenciones', 0) for g in gastos), 2), False),
        (12, total_facturado,                                          True),
    ]
    for col, val, is_main in totals:
        cell = ws.cell(row=row, column=col, value=val)
        cell.number_format = '"$"#,##0.00'
        cell.font = Font(
            name='Aptos',
            size=20 if is_main else 13,
            bold=is_main,
            color=SMTO_GREEN if is_main else WHITE,
        )
        cell.alignment = Alignment(horizontal='right', vertical='center', indent=2)
        cell.fill = PatternFill('solid', start_color=SMTO_BLACK)

    ws.cell(row=row, column=13).fill = PatternFill('solid', start_color=SMTO_BLACK)

    # ═══ FOOTER — one spacer row + a right-aligned version line ═══
    row += 2  # blank spacer + footer row
    ws.row_dimensions[row].height = 24
    ws.merge_cells(start_row=row, start_column=10, end_row=row, end_column=13)
    ft = ws.cell(row=row, column=10)
    ft.value = 'SMTO Engineering · v4.10'
    ft.font = Font(name='Aptos', size=9, italic=True, color=TEXT_MUTED)
    ft.alignment = Alignment(horizontal='right', vertical='center')
    ft.fill = PatternFill('solid', start_color=BG_PAGE)

    # ═══ LOGO — cropped of padding, anchored at B1 next to the title ═══
    logo_path = find_logo()
    if logo_path:
        try:
            pil = PILImage.open(logo_path)
            pil = crop_logo(pil)
            target_h = 60
            ratio = target_h / pil.height
            target_w = int(pil.width * ratio)
            pil = pil.resize((target_w, target_h), PILImage.LANCZOS)
            with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmp:
                logo_tmp = tmp.name
            pil.save(logo_tmp, 'PNG')
            img = XLImage(logo_tmp)
            img.width = target_w
            img.height = target_h
            img.anchor = 'B1'
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
