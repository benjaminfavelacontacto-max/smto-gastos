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
EXCEL_GREEN = '00B050'
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

# Currency-symbol lookup for the propina sub-row label when the parent
# gasto is in a foreign currency.
CURRENCY_SYMBOLS = {
    'USD': '$', 'EUR': '€', 'GBP': '£', 'JPY': '¥', 'CNY': '¥',
    'MXN': '$', 'CAD': 'C$', 'MYR': 'RM', 'CHF': 'Fr', 'AUD': 'A$',
}
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
    Output format: MM-DD-YY (2-digit year)."""
    dates = []
    for g in gastos:
        for key in ('fechaCobro', 'fechaFac'):
            v = g.get(key, '')
            if v and len(v) >= 8:
                # Normalize to YYYY-MM-DD for sorting
                if '-' in v and len(v) == 10 and v[4] == '-':
                    dates.append(v)  # already YYYY-MM-DD
                break
    if not dates:
        return '', ''
    dates.sort()  # ascending — smallest date first
    def fmt(d):
        # YYYY-MM-DD → MM-DD-YY
        return f'{d[5:7]}-{d[8:10]}-{d[2:4]}'
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

def build_workbook(gastos, colaborador=''):
    wb = Workbook()
    ws = wb.active
    ws.title = 'Reporte SMTO'
    ws.sheet_view.showGridLines = False

    # Column widths — semantic (wide CONCEPTO + supplier, narrow dates) so
    # each column gets the width that fits its content, regardless of order.
    # N = MONTO USD, O = T/C, P = right spacer. N + O sized equal so the
    # new "USD" KPI card (N5:O5) reads as a tidy 2-column tile above them.
    col_widths = {
        'A': 3, 'B': 15, 'C': 30, 'D': 11, 'E': 14, 'F': 11, 'G': 11,
        'H': 28, 'I': 12, 'J': 11, 'K': 11, 'L': 18, 'M': 15,
        'N': 13, 'O': 13, 'P': 3,
    }
    for col, w in col_widths.items():
        ws.column_dimensions[col].width = w

    # Paint full background light gray — scales with row count so the
    # outer spacer cols past the totals/footer still inherit BG_PAGE.
    nrows_painted = max(80, 40 + len(gastos))
    for r in range(1, nrows_painted):
        fill_row_bg(ws, r, 1, 16, BG_PAGE)

    # ═══ HEADER (rows 1-2) — title + colaborador labels + fields ═══
    ws.row_dimensions[1].height = 50
    ws.row_dimensions[2].height = 36

    # Title — centered across D1:G2; pairs with the 64px logo at B1
    ws.merge_cells('D1:G2')
    title = ws['D1']
    title.value = 'Reporte de Gastos'
    title.font = Font(name='Aptos', size=26, bold=True, color=SMTO_BLACK)
    title.alignment = Alignment(horizontal='center', vertical='center')
    title.fill = PatternFill('solid', start_color=BG_PAGE)

    # Right-side form labels (col H) — compact 11pt bold
    lbl1 = ws['H1']
    lbl1.value = 'Nombre de colaborador:'
    lbl1.font = Font(name='Aptos', size=11, bold=True, color=TEXT_SECONDARY)
    lbl1.alignment = Alignment(horizontal='right', vertical='center')
    lbl1.fill = PatternFill('solid', start_color=BG_PAGE)

    lbl2 = ws['H2']
    lbl2.value = 'Fecha de viaje:'
    lbl2.font = Font(name='Aptos', size=11, bold=True, color=TEXT_SECONDARY)
    lbl2.alignment = Alignment(horizontal='right', vertical='center')
    lbl2.fill = PatternFill('solid', start_color=BG_PAGE)

    # Right-side form fields (cols J:M) — compact 11pt
    ws.merge_cells('J1:M1')
    f1 = ws['J1']
    f1.value = colaborador  # filled programmatically from the selected colaborador
    f1.font = Font(name='Aptos', size=11, color=TEXT_PRIMARY)
    f1.fill = PatternFill('solid', start_color=WHITE)
    f1.border = Border(bottom=Side(style='thin', color=BORDER_LIGHT))
    f1.alignment = Alignment(horizontal='left', vertical='center', indent=1)

    ws.merge_cells('J2:M2')
    f2 = ws['J2']
    fecha_min, fecha_max = compute_date_range(gastos)
    if fecha_min:
        f2.value = f'DE: {fecha_min}  A: {fecha_max}'
    else:
        f2.value = ''
    f2.font = Font(name='Aptos', size=11, color=TEXT_PRIMARY)
    f2.fill = PatternFill('solid', start_color=WHITE)
    f2.alignment = Alignment(horizontal='left', vertical='center', indent=1)

    # Row 3 spacer + Row 4 closing line under the header section.
    # H3 carries a thin EXCEL_GREEN underline below the "Nombre/Fecha de viaje"
    # labels; row 4 itself is a 1pt-tall band sandwiched between a thin top
    # green border and a medium bottom green border so it reads as one bar.
    ws.row_dimensions[3].height = 10
    ws.row_dimensions[4].height = 1
    ws['H3'].border = Border(bottom=Side(style='thin', color=EXCEL_GREEN))
    for c in range(2, 16):
        cell = ws.cell(row=4, column=c)
        cell.fill = PatternFill('solid', start_color=BG_PAGE)
        cell.border = Border(
            top=Side(style='thin', color=EXCEL_GREEN),
            bottom=Side(style='medium', color=EXCEL_GREEN),
        )

    # ═══ KPI CARDS (rows 5-6) ═══
    ws.row_dimensions[5].height = 32  # KPI VALUES
    ws.row_dimensions[6].height = 16  # KPI LABELS

    # Pre-compute the data range so the KPI cards AND the bottom totals row
    # share the same =SUM() references. Counts every main row + any propina
    # sub-row, clamps to data_first when no gastos so the SUM range stays
    # valid (I10:I10 → 0).
    data_first = 10
    has_propina = lambda g: (g.get('montoPropina', 0) or 0) > 0 or (g.get('propinaExtranjero', 0) or 0) > 0
    data_rows_total = len(gastos) + sum(1 for g in gastos if has_propina(g))
    data_last = max(data_first, data_first + data_rows_total - 1)
    num_facturas = len(gastos)

    # (col_start, col_end, label, value, number_format, value_color)
    # KPI cards are live Excel formulas referencing the data band so the
    # numbers always match the bottom totals row — edit any data cell and
    # both update together. REGISTROS stays a static count (not a sum).
    # All five cards share a medium EXCEL_GREEN frame; TOTAL FACTURADO and
    # USD are tinted brand-green, the rest stay SMTO_BLACK.
    kpis = [
        ('B', 'D', 'TOTAL FACTURADO', f'=SUM(L{data_first}:L{data_last})', '"$"#,##0.00', SMTO_GREEN),
        ('E', 'G', 'IVA TOTAL',       f'=SUM(J{data_first}:J{data_last})', '"$"#,##0.00', SMTO_BLACK),
        ('H', 'J', 'RETENCIONES',     f'=SUM(K{data_first}:K{data_last})', '"$"#,##0.00', SMTO_BLACK),
        ('K', 'M', 'REGISTROS',       num_facturas,                        '0',           SMTO_BLACK),
        ('N', 'O', 'USD',             f'=SUM(N{data_first}:N{data_last})', '"$"#,##0.00', SMTO_GREEN),
    ]

    for col_start, col_end, label, value, fmt, value_color in kpis:
        # Value row (row 5) — anchor cell carries the value/font/fill; the
        # right-edge border has to also be applied to the LAST cell of the
        # merge because openpyxl renders borders per cell, not per merge.
        ws.merge_cells(f'{col_start}5:{col_end}5')
        v_cell = ws[f'{col_start}5']
        v_cell.value = value
        v_cell.number_format = fmt
        v_cell.font = Font(name='Aptos', size=20, bold=True, color=value_color)
        v_cell.alignment = Alignment(horizontal='left', vertical='center', indent=2)
        v_cell.fill = PatternFill('solid', start_color=WHITE)
        v_cell.border = Border(
            left=Side(style='medium', color=EXCEL_GREEN),
            top=Side(style='medium', color=EXCEL_GREEN),
            right=Side(style='medium', color=EXCEL_GREEN),
            bottom=Side(style='thin', color=BORDER_LIGHT),
        )
        ws[f'{col_end}5'].border = Border(
            right=Side(style='medium', color=EXCEL_GREEN),
            top=Side(style='medium', color=EXCEL_GREEN),
            bottom=Side(style='thin', color=BORDER_LIGHT),
        )

        # Label row (row 6) — same trick: border the merge end-cell too.
        ws.merge_cells(f'{col_start}6:{col_end}6')
        l_cell = ws[f'{col_start}6']
        l_cell.value = label
        l_cell.font = Font(name='Aptos', size=9, bold=True, color=TEXT_MUTED)
        l_cell.alignment = Alignment(horizontal='left', vertical='center', indent=2)
        l_cell.fill = PatternFill('solid', start_color=WHITE)
        l_cell.border = Border(
            left=Side(style='medium', color=EXCEL_GREEN),
            top=Side(style='thin', color=BORDER_LIGHT),
            right=Side(style='medium', color=EXCEL_GREEN),
            bottom=Side(style='medium', color=EXCEL_GREEN),
        )
        ws[f'{col_end}6'].border = Border(
            right=Side(style='medium', color=EXCEL_GREEN),
            top=Side(style='thin', color=BORDER_LIGHT),
            bottom=Side(style='medium', color=EXCEL_GREEN),
        )

    # Middle cells of each merged KPI range need explicit borders too —
    # without them, some Excel/LibreOffice versions render visible gaps
    # inside the cards even though the anchor + end cells have borders.
    MED_GREEN  = Side(style='medium', color=EXCEL_GREEN)
    THIN_LIGHT = Side(style='thin',   color=BORDER_LIGHT)
    for col_letter in ('C', 'F', 'I', 'L'):
        ws[f'{col_letter}5'].border = Border(
            left=MED_GREEN, right=MED_GREEN, top=MED_GREEN, bottom=THIN_LIGHT,
        )
        ws[f'{col_letter}6'].border = Border(
            left=MED_GREEN, right=MED_GREEN, top=THIN_LIGHT, bottom=MED_GREEN,
        )

    # Row 7 small gap + Row 8 pre-table spacer
    ws.row_dimensions[7].height = 8
    ws.row_dimensions[8].height = 14

    # ═══ TABLE HEADER (row 9) — green text, mostly centered ═══
    ws.row_dimensions[9].height = 28

    headers = ['RFC', 'PROVEEDOR', 'TIPO', 'FACTURA', 'F. FACTURA', 'F. COBRO', 'CONCEPTO', 'IMPORTE', 'IVA', 'RETENCIÓN', 'TOTAL', 'FORMA PAGO', 'MONTO USD', 'T/C']
    # PROVEEDOR and CONCEPTO stay left-aligned; the rest center.
    left_align_headers = {'PROVEEDOR', 'CONCEPTO'}

    for i, h in enumerate(headers):
        col = i + 2
        cell = ws.cell(row=9, column=col, value=h)
        cell.font = Font(name='Aptos', size=11, bold=True, color=EXCEL_GREEN)
        is_left = h in left_align_headers
        cell.alignment = Alignment(
            horizontal='left' if is_left else 'center',
            vertical='center',
            indent=2 if is_left else 0
        )
        cell.fill = PatternFill(fill_type=None)
        # Top + bottom medium green on every header; left edge on B9 (first),
        # right edge on M9 (last) so the band reads as one bordered strip.
        cell.border = Border(
            top=Side(style='medium', color=EXCEL_GREEN),
            bottom=Side(style='medium', color=EXCEL_GREEN),
            left=Side(style='medium', color=EXCEL_GREEN) if col == 2 else None,
            right=Side(style='medium', color=EXCEL_GREEN) if col == 15 else None,
        )

    # ═══ DATA ROWS (row 10+) ═══
    row = 10
    for idx, g in enumerate(gastos):
        ws.row_dimensions[row].height = 24
        is_alt = (idx % 2 == 1)
        bg = ROW_ALT if is_alt else WHITE

        fecha_fac = format_date(g.get('fechaFac', ''))
        fecha_cobro = format_date(g.get('fechaCobro', ''))
        forma = FORMA_PAGO_MAP.get(g.get('formaPago', '04'), g.get('formaPago', ''))
        importe_raw = round(g.get('importe', 0), 2)
        iva = round(g.get('iva', 0), 2)
        ret = round(g.get('retenciones', 0), 2)
        total = round(g.get('totalCFDI', 0) + g.get('montoPropina', 0), 2)
        tipo = g.get('tipo', 'Consumo')
        monto_usd_raw = round(g.get('montoUSD', 0) or 0, 2)
        tipo_cambio = round(g.get('tipoCambio', 0) or 0, 2)

        # Propina-row fields. OCR tickets matched by validarBanco Pass 0 have
        # importe/montoUSD that already INCLUDE the tip — so for those rows
        # we subtract the propina from the main-row display so the new
        # propina sub-row does not double-count the tip in SUM(I) / SUM(N).
        # CFDI rows with manual propina keep importe as the pre-tip subtotal
        # (the SAT XML stores SubTotal pre-tip), so no adjustment is needed.
        propina_mxn = round(g.get('montoPropina', 0) or 0, 2)
        propina_ext = round(g.get('propinaExtranjero', 0) or 0, 2)
        moneda_code = g.get('monedaCodigo') or g.get('moneda') or 'MXN'
        es_ticket = bool(g.get('esTicket'))
        if es_ticket and propina_mxn > 0:
            importe = max(0, round(importe_raw - propina_mxn, 2))
            monto_usd = max(0, round(monto_usd_raw - propina_ext, 2))
        else:
            importe = importe_raw
            monto_usd = monto_usd_raw

        # Column order matches the headers. PROVEEDOR and CONCEPTO are the only
        # left-aligned cells; everything else centers per the reference.
        cells = [
            (2,  g.get('rfc', ''),       'left',   'rfc'),
            (3,  g.get('proveedor', ''), 'left',   'normal_bold'),
            (4,  tipo,                   'center', 'badge_tipo'),
            (5,  g.get('noFactura', ''), 'center', 'normal'),
            (6,  fecha_fac,              'center', 'normal'),
            (7,  fecha_cobro,            'center', 'normal'),
            (8,  g.get('concepto', ''),  'left',   'normal'),
            (9,  importe,                  'center', 'currency'),
            (10, iva,                      'center', 'currency'),
            (11, ret,                      'center', 'currency'),
            # TOTAL is a true Excel formula so the column re-sums correctly
            # if the user edits Importe/IVA/Retención manually — also kills
            # the green "inconsistent formula" triangle Excel shows when a
            # column has static derived values next to formula candidates.
            (12, f'=I{row}+J{row}-K{row}', 'center', 'currency_bold'),
            (13, forma,                  'center', 'badge_pago'),
            (14, monto_usd,              'center', 'currency'),
            (15, tipo_cambio,            'center', 'tipocambio'),
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
                cell.font = Font(name='Aptos', size=10, color=TEXT_PRIMARY)
            elif style_type == 'currency_bold':
                cell.number_format = '"$"#,##0.00'
                cell.font = Font(name='Aptos', size=11, bold=True, color=SMTO_BLACK)
            elif style_type == 'rfc':
                # RFC sits one notch smaller than the supplier name so the
                # supplier reads as the row's primary identifier.
                cell.font = Font(name='Aptos', size=9, bold=True, color=TEXT_PRIMARY)
            elif style_type == 'normal_bold':
                cell.font = Font(name='Aptos', size=10, bold=True, color=TEXT_PRIMARY)
            elif style_type == 'badge_tipo':
                bg_b, fg_b = get_tipo_badge_colors(tipo)
                cell.fill = PatternFill('solid', start_color=bg_b)
                cell.font = Font(name='Aptos', size=9, bold=True, color=fg_b)
            elif style_type == 'badge_pago':
                cell.fill = PatternFill('solid', start_color=BADGE_GRAY_BG)
                cell.font = Font(name='Aptos', size=9, bold=True, color=BADGE_GRAY_FG)
            elif style_type == 'tipocambio':
                cell.number_format = '#,##0.00'
                cell.font = Font(name='Aptos', size=10, color=TEXT_PRIMARY)
            else:  # 'normal'
                cell.font = Font(name='Aptos', size=10, color=TEXT_PRIMARY)

        # Side spacer cells keep page bg through the data band.
        ws.cell(row=row, column=1).fill = PatternFill('solid', start_color=BG_PAGE)
        ws.cell(row=row, column=16).fill = PatternFill('solid', start_color=BG_PAGE)

        row += 1

        # ── Optional propina sub-row ──
        # When the parent gasto has any propina, append a slim italic sub-row
        # immediately below it so the tip reads as a separate ledger line.
        # SUM(I) / SUM(L) / SUM(N) at the totals row include it automatically.
        if propina_mxn > 0 or propina_ext > 0:
            ws.row_dimensions[row].height = 20
            propina_bg = 'F0FDF4'  # very subtle green tint

            # Paint the whole band first so per-cell font/border calls below
            # only need to touch the cells that carry content.
            for c in range(1, 16):
                pcell = ws.cell(row=row, column=c)
                pcell.fill = PatternFill('solid', start_color=propina_bg)
                pcell.border = Border(bottom=Side(style='hair', color=BORDER_LIGHT))

            # Col C — "↳ Propina" label, italic SMTO_GREEN.
            label = ws.cell(row=row, column=3)
            label.value = '  ↳  Propina'
            label.font = Font(name='Aptos', size=9, italic=True, color=SMTO_GREEN)
            label.alignment = Alignment(horizontal='left', vertical='center', indent=2)

            # Col H — concepto detail. For foreign currency, include the
            # native symbol + amount so the row reads as the original tip.
            concepto_p = ws.cell(row=row, column=8)
            if propina_ext > 0 and moneda_code != 'MXN':
                symbol = CURRENCY_SYMBOLS.get(moneda_code, moneda_code + ' ')
                concepto_p.value = f'Propina {symbol}{propina_ext:,.2f} {moneda_code}'
            else:
                concepto_p.value = 'Propina'
            concepto_p.font = Font(name='Aptos', size=9, italic=True, color=TEXT_SECONDARY)
            concepto_p.alignment = Alignment(horizontal='left', vertical='center', indent=2)

            # Col I (IMPORTE), col L (TOTAL) — the tip in MXN, brand-green
            # italic. Col J / K stay at 0 so the TOTAL formula on adjacent
            # rows isn't affected (this row writes a literal, not a formula).
            for col, val, bold in [(9, propina_mxn, False), (12, propina_mxn, True)]:
                c = ws.cell(row=row, column=col, value=val)
                c.number_format = '"$"#,##0.00'
                c.font = Font(name='Aptos', size=9, italic=True, bold=bold, color=SMTO_GREEN)
                c.alignment = Alignment(horizontal='right', vertical='center', indent=1)

            for col in (10, 11):  # IVA / RETENCIÓN — explicit 0 so SUM works
                c = ws.cell(row=row, column=col, value=0)
                c.number_format = '"$"#,##0.00'
                c.font = Font(name='Aptos', size=9, italic=True, color=TEXT_SECONDARY)
                c.alignment = Alignment(horizontal='right', vertical='center', indent=1)

            # Col M — FORMA PAGO mirror from the parent, lighter style.
            forma_p = ws.cell(row=row, column=13)
            forma_p.value = forma
            forma_p.font = Font(name='Aptos', size=9, italic=True, color=BADGE_GRAY_FG)
            forma_p.alignment = Alignment(horizontal='center', vertical='center')

            # Col N — foreign tip amount (only when the parent is foreign).
            if propina_ext > 0 and moneda_code != 'MXN':
                ext_c = ws.cell(row=row, column=14, value=propina_ext)
                ext_c.number_format = '#,##0.00'
                ext_c.font = Font(name='Aptos', size=9, italic=True, color=SMTO_GREEN)
                ext_c.alignment = Alignment(horizontal='right', vertical='center', indent=1)

            # Outer spacers stay on page bg so the propina band fits inside
            # the table boundary like every other data row.
            ws.cell(row=row, column=1).fill = PatternFill('solid', start_color=BG_PAGE)
            ws.cell(row=row, column=16).fill = PatternFill('solid', start_color=BG_PAGE)

            row += 1

    # ═══ ADAPTIVE TOTALS — placed right after the last data row ═══
    ws.row_dimensions[row].height = 14  # spacer
    row += 1

    ws.row_dimensions[row].height = 32
    for c in range(2, 16):
        cell = ws.cell(row=row, column=c)
        cell.fill = PatternFill('solid', start_color=SMTO_BLACK)
        cell.border = Border()

    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=8)
    lbl = ws.cell(row=row, column=2, value='TOTAL CUENTA')
    lbl.font = Font(name='Aptos', size=12, bold=True, color=WHITE)
    lbl.alignment = Alignment(horizontal='right', vertical='center', indent=2)
    lbl.fill = PatternFill('solid', start_color=SMTO_BLACK)

    # Same SUM formulas as the KPI cards above so the two views always agree.
    # `data_first` / `data_last` are pre-computed once at the top of the KPI
    # section and cover both main rows and any propina sub-rows that the data
    # loop inserted. T/C (col O) intentionally has no total — a sum of
    # exchange rates is not meaningful.
    totals = [
        (9,  f'=SUM(I{data_first}:I{data_last})', False),
        (10, f'=SUM(J{data_first}:J{data_last})', False),
        (11, f'=SUM(K{data_first}:K{data_last})', False),
        (12, f'=SUM(L{data_first}:L{data_last})', True),
        (14, f'=SUM(N{data_first}:N{data_last})', False),
    ]
    for col, val, is_main in totals:
        cell = ws.cell(row=row, column=col, value=val)
        cell.number_format = '"$"#,##0.00'
        cell.font = Font(
            name='Aptos',
            size=16 if is_main else 11,
            bold=is_main,
            color=SMTO_GREEN if is_main else WHITE,
        )
        cell.alignment = Alignment(horizontal='right', vertical='center', indent=2)
        cell.fill = PatternFill('solid', start_color=SMTO_BLACK)

    # FORMA PAGO (M) and T/C (O) inside the totals band get the black fill
    # but no value (FORMA is non-aggregatable text; T/C is per-row).
    ws.cell(row=row, column=13).fill = PatternFill('solid', start_color=SMTO_BLACK)
    ws.cell(row=row, column=15).fill = PatternFill('solid', start_color=SMTO_BLACK)

    # ═══ FOOTER — one spacer row + a right-aligned version line ═══
    row += 2  # blank spacer + footer row
    ws.row_dimensions[row].height = 18
    ws.merge_cells(start_row=row, start_column=10, end_row=row, end_column=15)
    ft = ws.cell(row=row, column=10)
    ft.value = 'SMTO Engineering · v7.16'
    ft.font = Font(name='Aptos', size=8, italic=True, color=TEXT_MUTED)
    ft.alignment = Alignment(horizontal='right', vertical='center')
    ft.fill = PatternFill('solid', start_color=BG_PAGE)

    # ═══ LOGO — cropped of padding, anchored at B1 next to the title ═══
    logo_path = find_logo()
    if logo_path:
        try:
            pil = PILImage.open(logo_path)
            pil = crop_logo(pil)
            target_h = 64
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
            data = json.loads(body)
            # Backward compatible: old clients send a bare array, new clients
            # send { gastos: [...], colaborador: '...' }.
            if isinstance(data, dict):
                gastos = data.get('gastos', [])
                colaborador = data.get('colaborador', '')
            else:
                gastos = data
                colaborador = ''
            wb = build_workbook(gastos, colaborador)
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
