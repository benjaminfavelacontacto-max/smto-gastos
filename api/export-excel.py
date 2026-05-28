from http.server import BaseHTTPRequestHandler
import json, os, sys, tempfile, traceback, re
from datetime import datetime

IMPORT_ERROR = None
try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
    from openpyxl.utils import get_column_letter
    from openpyxl.drawing.image import Image as XLImage
    from openpyxl.workbook.protection import WorkbookProtection
    from PIL import Image as PILImage
except Exception:
    IMPORT_ERROR = traceback.format_exc()

# openpyxl rejects XML 1.0 illegal control chars (vertical tab, form feed, etc.).
# CFDI fields and OCR output occasionally smuggle one in (PDF text extraction
# is notorious for U+000B between syllables). Strip them recursively from any
# string value in the incoming payload before it reaches ws.cell.
_ILLEGAL_XLSX_CHARS = re.compile(r'[\x00-\x08\x0b-\x0c\x0e-\x1f]')

def sanitize(v):
    if isinstance(v, str):
        return _ILLEGAL_XLSX_CHARS.sub('', v)
    if isinstance(v, dict):
        return {k: sanitize(x) for k, x in v.items()}
    if isinstance(v, list):
        return [sanitize(x) for x in v]
    return v

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

# Colaboradores especiales — replica de COLABORADORES_ESPECIALES en App.jsx.
# Estos colaboradores reciben una celda editable de Tipo de Cambio en su Excel
# para convertir filas en USD a MXN sobre la marcha.
COLABORADORES_ESPECIALES = {
    'Alejandro Olivar',
    'Victor Aceves',
    'Miranda Navarro',
    'Olivia Gil',
}

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
        # YYYY-MM-DD → DD/MM/AAAA
        return f'{d[8:10]}/{d[5:7]}/{d[0:4]}'
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
    """Formato de fecha SOLO para el Excel exportado: DD/MM/AAAA.
    Acepta YYYY-MM-DD (interno) o DD/MM/YYYY (ya formateado)."""
    if not date_str: return ''
    if '-' in date_str and len(date_str) == 10:
        parts = date_str.split('-')
        # YYYY-MM-DD → DD/MM/YYYY
        return f'{parts[2]}/{parts[1]}/{parts[0]}'
    if '/' in date_str:
        parts = date_str.split('/')
        # Asume DD/MM/YYYY ya correcto; sólo normaliza el año a 4 dígitos.
        yr = parts[2] if len(parts[2]) == 4 else ('20' + parts[2])
        return f'{parts[0]}/{parts[1]}/{yr}'
    return date_str

FORMA_PAGO_MAP = {
    '01': '01 - Efectivo',
    '02': '02 - Efectivo',
    '03': '03 - Transferencia',
    '04': '04 - Tarjeta de Crédito',
    '99': '99 - Por Definir',
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

def build_workbook(gastos, colaborador='', poliza_numero='N/A'):
    wb = Workbook()
    # Protección a nivel libro EXPLÍCITAMENTE deshabilitada. Sin esto, algunos
    # usuarios de Windows ven el diálogo "es un archivo de solo lectura" al
    # intentar guardar — Excel interpreta cualquier elemento <fileSharing> o
    # WorkbookProtection con flags ambiguos como "read-only recommended".
    wb.security = WorkbookProtection(
        lockStructure=False,
        lockWindows=False,
        lockRevision=False,
        workbookPassword=None,
        revisionsPassword=None,
    )
    ws = wb.active
    ws.title = 'Reporte SMTO'
    ws.sheet_view.showGridLines = False
    ws.protection.sheet = False  # make the sheet fully editable after export
    ws.protection.enabled = False  # belt-and-suspenders: also drop the element

    # Column widths — semantic (wide CONCEPTO + supplier, narrow dates).
    # Layout: A spacer, B-N data, O BANCO (nuevo), P MONTO USD, Q T/C, R spacer.
    col_widths = {
        'A': 3, 'B': 15, 'C': 30, 'D': 11, 'E': 10, 'F': 14, 'G': 11, 'H': 11,
        'I': 28, 'J': 12, 'K': 11, 'L': 11, 'M': 18, 'N': 15,
        'O': 12, 'P': 13, 'Q': 13, 'R': 3,
    }
    for col, w in col_widths.items():
        ws.column_dimensions[col].width = w

    # Paint full background light gray — scales with row count so the
    # outer spacer cols past the totals/footer still inherit BG_PAGE.
    nrows_painted = max(80, 40 + len(gastos))
    for r in range(1, nrows_painted):
        fill_row_bg(ws, r, 1, 18, BG_PAGE)

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
    for c in range(2, 18):
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
        ('B', 'D', 'TOTAL FACTURADO', f'=SUM(M{data_first}:M{data_last})', '"$"#,##0.00', SMTO_GREEN),
        ('E', 'G', 'IVA TOTAL',       f'=SUM(K{data_first}:K{data_last})', '"$"#,##0.00', SMTO_BLACK),
        ('H', 'J', 'RETENCIONES',     f'=SUM(L{data_first}:L{data_last})', '"$"#,##0.00', SMTO_BLACK),
        ('K', 'M', 'REGISTROS',       num_facturas,                        '0',           SMTO_BLACK),
        ('N', 'P', 'USD',             f'=SUM(P{data_first}:P{data_last})', '"$"#,##0.00', SMTO_GREEN),
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
    for col_letter in ('C', 'F', 'I', 'L', 'O'):
        ws[f'{col_letter}5'].border = Border(
            left=MED_GREEN, right=MED_GREEN, top=MED_GREEN, bottom=THIN_LIGHT,
        )
        ws[f'{col_letter}6'].border = Border(
            left=MED_GREEN, right=MED_GREEN, top=THIN_LIGHT, bottom=MED_GREEN,
        )

    # ═══ TIPO DE CAMBIO EDITABLE (solo colaboradores especiales) ═══
    # Una sola celda en N7 que el usuario puede editar libremente. Las filas
    # de gastos en USD usan formulas =(monto USD)*$N$7 en IMPORTE / IVA /
    # RETENCIÓN, y =$N$7 en la columna T/C. Cambiar este valor recalcula
    # automáticamente todas las filas USD y los KPI / totales que las suman.
    is_especial = (colaborador or '').strip() in COLABORADORES_ESPECIALES
    tc_ref = None

    if is_especial:
        ws.row_dimensions[7].height = 30

        # Label en K7:M7 — alineado a la derecha, fondo ámbar suave.
        ws.merge_cells('K7:M7')
        lbl_tc = ws['K7']
        lbl_tc.value = 'TIPO DE CAMBIO USD →'
        lbl_tc.font = Font(name='Aptos', size=11, bold=True, color=BADGE_AMBER_FG)
        lbl_tc.alignment = Alignment(horizontal='right', vertical='center', indent=1)
        lbl_tc.fill = PatternFill('solid', start_color=BADGE_AMBER_BG)
        # Pintar el resto de la celda combinada para que se vea uniforme
        for col_letter in ('L', 'M'):
            ws[f'{col_letter}7'].fill = PatternFill('solid', start_color=BADGE_AMBER_BG)

        # Input editable en N7 — fondo ámbar más vivo, borde grueso naranja.
        tc_input = ws['N7']
        default_tc = 17.50
        for g in gastos:
            m = (g.get('monedaCodigo') or g.get('moneda') or 'MXN').upper()
            rate = g.get('tipoCambio') or 0
            if m != 'MXN' and rate > 0:
                default_tc = round(float(rate), 4)
                break
        tc_input.value = default_tc
        tc_input.number_format = '#,##0.0000'
        tc_input.font = Font(name='Aptos', size=16, bold=True, color=BADGE_AMBER_FG)
        tc_input.alignment = Alignment(horizontal='center', vertical='center')
        tc_input.fill = PatternFill('solid', start_color='FDE68A')  # amber-200
        amber_side = Side(style='medium', color='F59E0B')
        tc_input.border = Border(
            left=amber_side, right=amber_side, top=amber_side, bottom=amber_side,
        )

        tc_ref = '$N$7'

    # Row 7 (cuando NO es especial) small gap + Row 8 pre-table spacer
    if not is_especial:
        ws.row_dimensions[7].height = 8
    ws.row_dimensions[8].height = 14

    # ═══ TABLE HEADER (row 9) — green text, mostly centered ═══
    ws.row_dimensions[9].height = 28

    headers = ['RFC', 'PROVEEDOR', 'TIPO', 'PÓLIZA', 'FACTURA', 'F. FACTURA', 'F. COBRO', 'CONCEPTO', 'IMPORTE', 'IVA', 'RETENCIÓN', 'TOTAL', 'FORMA PAGO', 'BANCO', 'MONTO USD', 'T/C']
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
            right=Side(style='medium', color=EXCEL_GREEN) if col == 17 else None,
        )

    # ═══ DATA ROWS (row 10+) ═══
    row = 10
    for idx, g in enumerate(gastos):
        ws.row_dimensions[row].height = 15
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
        # Valores en USD preservados desde parseCFDI (importe/iva/ret en la
        # gasto vienen en 0 hasta que se aplique el T/C en el Excel). El
        # bloque is_usd_row de abajo los multiplica por la celda editable.
        importe_usd_raw = round(g.get('importeUSD', 0) or 0, 2)
        iva_usd_raw     = round(g.get('ivaUSD', 0) or 0, 2)
        ret_usd_raw     = round(g.get('retencionesUSD', 0) or 0, 2)
        # Per-row póliza (sólo se llena cuando el colaborador es especial
        # y corrió el cotejo con su Saldos); para todos los demás cae al
        # folio Clara global del colaborador.
        poliza_row = g.get('polizaNumero') or poliza_numero

        # Propina-row fields. When the gasto has any propina, the main row
        # shows the NET amount (total − tip) so main_row + propina_sub_row
        # always sums to the full charge in SUM(I) / SUM(L) / SUM(N).
        # Basis is totalCFDI − montoPropina (and montoExtranjero −
        # propinaExtranjero on the foreign side). Both fields are clamped
        # to ≥ 0.
        propina_mxn = round(g.get('montoPropina', 0) or 0, 2)
        propina_ext = round(g.get('propinaExtranjero', 0) or 0, 2)
        moneda_code = g.get('monedaCodigo') or g.get('moneda') or 'MXN'
        total_mxn   = round(g.get('totalCFDI', 0) or 0, 2)
        monto_ext   = round(g.get('montoExtranjero', 0) or g.get('montoUSD', 0) or 0, 2)

        if propina_mxn > 0 or propina_ext > 0:
            # El renglón principal mantiene los montos REALES del CFDI:
            # IMPORTE = SubTotal (sin IVA), de modo que L = I+J-K = totalCFDI
            # tal y como aparece en la factura ($410, $842, etc).
            # La propina vive en su propio sub-renglón con su L propio.
            importe   = importe_raw
            monto_usd = monto_usd_raw
        else:
            importe   = importe_raw
            monto_usd = monto_usd_raw

        # USD passthrough: si el colaborador es especial y la factura viene
        # en moneda extranjera, IMPORTE / IVA / RETENCIÓN / T/C usan formulas
        # que multiplican el valor en USD por la celda editable $N$7. Cambiar
        # esa celda recalcula toda la fila al instante. importe/iva/ret en
        # la gasto vienen en 0 (la UI los muestra así hasta tener T/C); los
        # valores reales en USD viajan en importeUSD / ivaUSD / retencionesUSD.
        is_usd_row = (
            is_especial
            and tc_ref is not None
            and (moneda_code or 'MXN').upper() not in ('MXN', '', 'XXX')
            and (importe_usd_raw > 0 or iva_usd_raw > 0 or ret_usd_raw > 0 or monto_usd_raw > 0)
        )
        if is_usd_row:
            # Fallback: si por alguna razón no llegaron los desgloses USD,
            # asumimos que el total USD ES el subtotal (IVA/Ret = 0) — el
            # MXN-importe siempre coincidirá con el MONTO USD visible.
            usd_imp = importe_usd_raw if importe_usd_raw > 0 else monto_usd_raw
            usd_iva = iva_usd_raw
            usd_ret = ret_usd_raw
            importe_val = f'={usd_imp}*{tc_ref}'
            iva_val     = f'={usd_iva}*{tc_ref}' if usd_iva > 0 else 0
            ret_val     = f'={usd_ret}*{tc_ref}' if usd_ret > 0 else 0
            tc_val      = f'={tc_ref}'
        else:
            importe_val = importe
            iva_val     = iva
            ret_val     = ret
            tc_val      = tipo_cambio

        # Column order matches the headers. PROVEEDOR and CONCEPTO are the only
        # left-aligned cells; everything else centers per the reference.
        banco = (g.get('banco') or '').strip()
        cells = [
            (2,  g.get('rfc', ''),       'left',   'rfc'),
            (3,  g.get('proveedor', ''), 'left',   'normal_bold'),
            (4,  tipo,                   'center', 'badge_tipo'),
            (5,  poliza_row,             'center', 'normal'),
            (6,  g.get('noFactura', ''), 'center', 'normal'),
            (7,  fecha_fac,              'center', 'normal'),
            (8,  fecha_cobro,            'center', 'normal'),
            (9,  g.get('concepto', ''),  'left',   'normal'),
            (10, importe_val,              'center', 'currency'),
            (11, iva_val,                  'center', 'currency'),
            (12, ret_val,                  'center', 'currency'),
            # TOTAL is a true Excel formula so the column re-sums correctly
            # if the user edits Importe/IVA/Retención manually — también
            # recalcula automáticamente cuando IMPORTE/IVA/RET son fórmulas
            # de USD apuntando a la celda editable de T/C.
            (13, f'=J{row}+K{row}-L{row}', 'center', 'currency_bold'),
            (14, forma,                  'center', 'badge_pago'),
            (15, banco,                  'center', 'normal'),
            (16, monto_usd,              'center', 'currency'),
            (17, tc_val,                 'center', 'tipocambio'),
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
                cell.font = Font(name='Calibri', size=10, color=TEXT_PRIMARY)
            elif style_type == 'currency_bold':
                cell.number_format = '"$"#,##0.00'
                cell.font = Font(name='Calibri', size=10, bold=True, color=SMTO_BLACK)
            elif style_type == 'rfc':
                cell.font = Font(name='Calibri', size=10, bold=True, color=TEXT_PRIMARY)
            elif style_type == 'normal_bold':
                cell.font = Font(name='Calibri', size=10, bold=True, color=TEXT_PRIMARY)
            elif style_type == 'badge_tipo':
                bg_b, fg_b = get_tipo_badge_colors(tipo)
                cell.fill = PatternFill('solid', start_color=bg_b)
                cell.font = Font(name='Calibri', size=9, bold=True, color=fg_b)
            elif style_type == 'badge_pago':
                cell.fill = PatternFill('solid', start_color=BADGE_GRAY_BG)
                cell.font = Font(name='Calibri', size=9, bold=True, color=BADGE_GRAY_FG)
            elif style_type == 'tipocambio':
                cell.number_format = '#,##0.00'
                cell.font = Font(name='Calibri', size=10, color=TEXT_PRIMARY)
            else:  # 'normal'
                cell.font = Font(name='Calibri', size=10, color=TEXT_PRIMARY)

        # Side spacer cells keep page bg through the data band.
        ws.cell(row=row, column=1).fill = PatternFill('solid', start_color=BG_PAGE)
        ws.cell(row=row, column=18).fill = PatternFill('solid', start_color=BG_PAGE)

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
            for c in range(1, 18):
                pcell = ws.cell(row=row, column=c)
                pcell.fill = PatternFill('solid', start_color=propina_bg)
                pcell.border = Border(bottom=Side(style='hair', color=BORDER_LIGHT))

            # Col C — "↳ Propina" label, italic SMTO_GREEN.
            label = ws.cell(row=row, column=3)
            label.value = '  ↳  Propina'
            label.font = Font(name='Calibri', size=9, italic=True, color=SMTO_GREEN)
            label.alignment = Alignment(horizontal='left', vertical='center', indent=2)

            # Col I — concepto detail. For foreign currency, include the
            # native symbol + amount so the row reads as the original tip.
            concepto_p = ws.cell(row=row, column=9)
            if propina_ext > 0 and moneda_code != 'MXN':
                symbol = CURRENCY_SYMBOLS.get(moneda_code, moneda_code + ' ')
                concepto_p.value = f'Propina {symbol}{propina_ext:,.2f} {moneda_code}'
            else:
                concepto_p.value = 'Propina'
            concepto_p.font = Font(name='Calibri', size=9, italic=True, color=TEXT_SECONDARY)
            concepto_p.alignment = Alignment(horizontal='left', vertical='center', indent=2)

            # Col J (IMPORTE), col M (TOTAL) — the tip in MXN, brand-green
            # italic. Col K / L stay at 0 so the TOTAL formula on adjacent
            # rows isn't affected (this row writes a literal, not a formula).
            for col, val, bold in [(10, propina_mxn, False), (13, propina_mxn, True)]:
                c = ws.cell(row=row, column=col, value=val)
                c.number_format = '"$"#,##0.00'
                c.font = Font(name='Calibri', size=9, italic=True, bold=bold, color=SMTO_GREEN)
                c.alignment = Alignment(horizontal='right', vertical='center', indent=1)

            for col in (11, 12):  # IVA / RETENCIÓN — explicit 0 so SUM works
                c = ws.cell(row=row, column=col, value=0)
                c.number_format = '"$"#,##0.00'
                c.font = Font(name='Calibri', size=9, italic=True, color=TEXT_SECONDARY)
                c.alignment = Alignment(horizontal='right', vertical='center', indent=1)

            # Col N — FORMA PAGO mirror from the parent, lighter style.
            forma_p = ws.cell(row=row, column=14)
            forma_p.value = forma
            forma_p.font = Font(name='Calibri', size=9, italic=True, color=BADGE_GRAY_FG)
            forma_p.alignment = Alignment(horizontal='center', vertical='center')

            # Col P — foreign tip amount (only when the parent is foreign).
            # BANCO en col 15 queda vacío para propinas (heredan banco del padre).
            if propina_ext > 0 and moneda_code != 'MXN':
                ext_c = ws.cell(row=row, column=16, value=propina_ext)
                ext_c.number_format = '#,##0.00'
                ext_c.font = Font(name='Calibri', size=9, italic=True, color=SMTO_GREEN)
                ext_c.alignment = Alignment(horizontal='right', vertical='center', indent=1)

            # Outer spacers stay on page bg so the propina band fits inside
            # the table boundary like every other data row.
            ws.cell(row=row, column=1).fill = PatternFill('solid', start_color=BG_PAGE)
            ws.cell(row=row, column=18).fill = PatternFill('solid', start_color=BG_PAGE)

            row += 1

    # ═══ ADAPTIVE TOTALS — placed right after the last data row ═══
    ws.row_dimensions[row].height = 14  # spacer
    row += 1

    ws.row_dimensions[row].height = 32
    for c in range(2, 18):
        cell = ws.cell(row=row, column=c)
        cell.fill = PatternFill('solid', start_color=SMTO_BLACK)
        cell.border = Border()

    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=9)
    lbl = ws.cell(row=row, column=2, value='TOTAL CUENTA')
    lbl.font = Font(name='Aptos', size=12, bold=True, color=WHITE)
    lbl.alignment = Alignment(horizontal='right', vertical='center', indent=2)
    lbl.fill = PatternFill('solid', start_color=SMTO_BLACK)

    # Same SUM formulas as the KPI cards above so the two views always agree.
    # `data_first` / `data_last` are pre-computed once at the top de la sección
    # de KPIs y cubren las filas principales más las sub-filas de propina.
    # MONTO USD ahora vive en col P; T/C (col Q) no tiene totalizado — sumar
    # tipos de cambio no es significativo. BANCO (col O) tampoco totaliza.
    totals = [
        (10, f'=SUM(J{data_first}:J{data_last})', False),
        (11, f'=SUM(K{data_first}:K{data_last})', False),
        (12, f'=SUM(L{data_first}:L{data_last})', False),
        (13, f'=SUM(M{data_first}:M{data_last})', True),
        (16, f'=SUM(P{data_first}:P{data_last})', False),
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

    # FORMA PAGO (N), BANCO (O) y T/C (Q) en la banda de totales sólo llevan
    # el fill negro (FORMA y BANCO son no-agregables; T/C es per-fila).
    ws.cell(row=row, column=14).fill = PatternFill('solid', start_color=SMTO_BLACK)
    ws.cell(row=row, column=15).fill = PatternFill('solid', start_color=SMTO_BLACK)
    ws.cell(row=row, column=17).fill = PatternFill('solid', start_color=SMTO_BLACK)

    # ═══ FOOTER — one spacer row + a right-aligned version line ═══
    row += 2  # blank spacer + footer row
    ws.row_dimensions[row].height = 18
    ws.merge_cells(start_row=row, start_column=11, end_row=row, end_column=17)
    ft = ws.cell(row=row, column=11)
    ft.value = 'SMTO Engineering · v7.78'
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

    # Explicitly unlock every cell so openpyxl's default locked=True
    # doesn't prevent editing even when sheet protection is off.
    _unlocked = Protection(locked=False)
    for row in ws.iter_rows():
        for cell in row:
            cell.protection = _unlocked

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
                poliza_numero = data.get('polizaNumero', 'N/A')
            else:
                gastos = data
                colaborador = ''
                poliza_numero = 'N/A'
            gastos = sanitize(gastos)
            colaborador = sanitize(colaborador)
            poliza_numero = sanitize(poliza_numero)
            wb = build_workbook(gastos, colaborador, poliza_numero)
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
