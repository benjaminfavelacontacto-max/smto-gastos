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

def parse_date_obj(date_str):
    """Convierte la fecha a un datetime REAL para que Excel la trate como
    fecha (celda con number_format DD/MM/AAAA: se ve como fecha, se ordena y
    se filtra como fecha — no como texto 'General'). Acepta YYYY-MM-DD
    (interno), DD/MM/YYYY o DD/MM/YY. Devuelve None si no se puede parsear,
    para que el caller caiga al string de format_date sin perder el dato."""
    if not date_str:
        return None
    s = str(date_str).strip()
    # YYYY-MM-DD (formato interno de la app)
    if '-' in s and len(s) == 10 and s[4] == '-':
        try:
            return datetime(int(s[0:4]), int(s[5:7]), int(s[8:10]))
        except ValueError:
            return None
    # DD/MM/YYYY o DD/MM/YY (ya formateado)
    if '/' in s:
        parts = s.split('/')
        if len(parts) == 3:
            d, m, y = parts
            if len(y) == 2:
                y = '20' + y
            try:
                return datetime(int(y), int(m), int(d))
            except ValueError:
                return None
    return None

FORMA_PAGO_MAP = {
    '01': '01 - Efectivo',
    '02': '02 - Efectivo',
    '03': '03 - Transferencia',
    '04': '04 - Tarjeta de Crédito',
    '99': '99 - Por Definir',
}

def get_tipo_badge_colors(tipo):
    """Return (bg, fg) for the tipo badge.

    Paleta acordada (archivo 'Colores para tipos.xlsx'):
      VERDE    → nómina, impuestos y costos fijos recurrentes
      AZUL     → operación, viáticos y servicios (default)
      MORADO   → mantenimiento, herramientas y tecnología
      AMARILLO → papelería, uniformes, banca, préstamos y no comprobado
      GRIS     → traspasos / rechazadas (neutro)
    """
    t = (tipo or '').lower().strip()

    # ── MORADO: mantenimiento, herramientas, tecnología ──
    if any(k in t for k in ['manto', 'herramienta', 'it & sw', 'it&sw']):
        return BADGE_PURPLE_BG, BADGE_PURPLE_FG

    # ── AMARILLO: papelería, uniformes, banca, préstamos, no comprobado ──
    if any(k in t for k in ['papelería', 'papeleria', 'uniforme', 'comisión banco',
                            'comision banco', 'préstamo', 'prestamo', 'crédito',
                            'credito', 'no comprobado']):
        return BADGE_AMBER_BG, BADGE_AMBER_FG

    # ── GRIS: neutros (traspasos, rechazadas) ──
    if any(k in t for k in ['rechazada', 'devolución', 'devolucion', 'traspaso']):
        return BADGE_GRAY_BG, BADGE_GRAY_FG

    # ── VERDE: nómina, impuestos y costos fijos ──
    if any(k in t for k in ['nómina', 'nomina', 'fondo de ahorro', 'regalía', 'regalia',
                            'seguro', 'celular', 'renta oficina', 'imss', 'infonavit',
                            'isr', 'iva', 'isn']):
        return BADGE_GREEN_BG, BADGE_GREEN_FG

    # ── AZUL: operación, viáticos y servicios (default) ──
    return BADGE_BLUE_BG, BADGE_BLUE_FG

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

def style_data_cell(cell, style_type, tipo='', diff_num=None):
    """Aplica number_format + font (y fill en badges) a una celda de datos.
    Lo comparten el renglón principal y la sub-fila de propina para que ambos
    se vean idénticos."""
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
    elif style_type == 'date':
        cell.number_format = 'DD/MM/YYYY'
        cell.font = Font(name='Calibri', size=10, color=TEXT_PRIMARY)
    elif style_type == 'diff':
        cell.number_format = '"$"#,##0.00'
        if isinstance(diff_num, (int, float)):
            if diff_num < 0:
                cell.font = Font(name='Calibri', size=10, bold=True, color='B91C1C')
            elif diff_num > 0:
                cell.font = Font(name='Calibri', size=10, bold=True, color='15803D')
            else:
                cell.font = Font(name='Calibri', size=10, color=TEXT_PRIMARY)
        else:
            cell.font = Font(name='Calibri', size=10, color=TEXT_PRIMARY)
    else:  # 'normal'
        cell.font = Font(name='Calibri', size=10, color=TEXT_PRIMARY)

def build_workbook(gastos, colaborador='', poliza_numero='N/A', polizas_map=None):
    # Invierte POLIZAS_CLARA (nombre → folio) a (folio → nombre) para que la
    # columna USUARIO pueda resolver cada póliza al colaborador dueño.
    polizas_map = polizas_map or {}
    usuario_por_poliza = {}
    for nombre, folio in polizas_map.items():
        if folio:
            usuario_por_poliza[str(folio).strip()] = nombre
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
    # Forzar recálculo de TODAS las fórmulas al abrir (KPIs, TOTAL, SUM). openpyxl
    # no precalcula valores; sin esto, un visor que respeta el caché muestra 0.
    wb.calculation.fullCalcOnLoad = True
    ws = wb.active
    ws.title = 'Reporte SMTO'
    ws.sheet_view.showGridLines = False
    ws.protection.sheet = False  # make the sheet fully editable after export
    ws.protection.enabled = False  # belt-and-suspenders: also drop the element

    # Column widths — semantic (wide CONCEPTO + supplier, narrow dates).
    # Layout: A spacer, B RFC, C PROVEEDOR, D TIPO, E PÓLIZA, F FACTURA,
    # G F.FACTURA, H F.COBRO, I CONCEPTO, J IMPORTE, K IVA, L ISH/IEPS (nueva),
    # M RETENCIÓN, N TOTAL, O FORMA PAGO, P BANCO, Q MONTO USD, R T/C,
    # S USUARIO, T COBRADO, U FACTURADO, V DIFERENCIA, W spacer.
    col_widths = {
        'A': 3, 'B': 15, 'C': 30, 'D': 11, 'E': 10, 'F': 14, 'G': 11, 'H': 11,
        'I': 28, 'J': 12, 'K': 11, 'L': 11, 'M': 11, 'N': 18,
        'O': 15, 'P': 22, 'Q': 13, 'R': 13, 'S': 24,
        'T': 14, 'U': 14, 'V': 14, 'W': 3,
    }
    for col, w in col_widths.items():
        ws.column_dimensions[col].width = w

    # Paint full background light gray — scales with row count so the
    # outer spacer cols past the totals/footer still inherit BG_PAGE.
    nrows_painted = max(80, 40 + len(gastos))
    for r in range(1, nrows_painted):
        fill_row_bg(ws, r, 1, 23, BG_PAGE)

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
    for c in range(2, 23):
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
    # NOTA: tras agregar la columna ISH/IEPS (L), las columnas de datos se
    # recorrieron: TOTAL ahora vive en N, RETENCIÓN en M, MONTO USD en Q y
    # DIFERENCIA en V. Los SUM de los KPIs apuntan a esas letras nuevas.
    kpis = [
        ('B', 'D', 'TOTAL FACTURADO', f'=SUM(N{data_first}:N{data_last})', '"$"#,##0.00', SMTO_GREEN),
        ('E', 'G', 'IVA TOTAL',       f'=SUM(K{data_first}:K{data_last})', '"$"#,##0.00', SMTO_BLACK),
        ('H', 'J', 'RETENCIONES',     f'=SUM(M{data_first}:M{data_last})', '"$"#,##0.00', SMTO_BLACK),
        ('K', 'M', 'REGISTROS',       num_facturas,                        '0',           SMTO_BLACK),
        ('N', 'P', 'USD',             f'=SUM(Q{data_first}:Q{data_last})', '"$"#,##0.00', SMTO_GREEN),
        # DIFERENCIA cobrado vs facturado — span Q:V (6 cols) para que la banda
        # de tarjetas llegue al borde derecho de la tabla (ahora B:V).
        # La suma SOLO incluye renglones Clara MXN Credito (los demás
        # escriben '' en col V, ignorados por SUM).
        ('Q', 'V', 'DIFERENCIA',      f'=SUM(V{data_first}:V{data_last})', '"$"#,##0.00', SMTO_GREEN),
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
    # Middle cells por tarjeta: B-D→C, E-G→F, H-J→I, K-M→L, N-P→O,
    # Q-V (DIFERENCIA, 6 cols) → R, S, T, U
    for col_letter in ('C', 'F', 'I', 'L', 'O', 'R', 'S', 'T', 'U'):
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

    headers = ['RFC', 'PROVEEDOR', 'TIPO', 'PÓLIZA', 'FACTURA', 'F. FACTURA', 'F. COBRO', 'CONCEPTO', 'IMPORTE', 'IVA', 'ISH/IEPS', 'RETENCIÓN', 'TOTAL', 'FORMA PAGO', 'BANCO', 'MONTO USD', 'T/C', 'USUARIO', 'COBRADO', 'FACTURADO', 'DIFERENCIA']
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
        # right edge on V9 (last, col 22) so the band reads as one bordered strip.
        cell.border = Border(
            top=Side(style='medium', color=EXCEL_GREEN),
            bottom=Side(style='medium', color=EXCEL_GREEN),
            left=Side(style='medium', color=EXCEL_GREEN) if col == 2 else None,
            right=Side(style='medium', color=EXCEL_GREEN) if col == 22 else None,
        )

    # Autofiltro sobre el encabezado (fila 9) + todas las filas de datos
    # (B9:U{data_last}), para que el usuario filtre cualquier columna —
    # CONCEPTO, TIPO, PROVEEDOR, fechas, etc.— con los dropdowns nativos.
    # data_last se calculó arriba (cubre filas principales + sub-filas de
    # propina) y NO incluye la banda de TOTAL CUENTA, que queda fuera del
    # filtro. Nota: TOTAL usa SUM (no SUBTOTAL), así que la fila de totales
    # no cambia al filtrar — es el comportamiento previo, intacto.
    ws.auto_filter.ref = f'B9:V{data_last}'

    # ═══ DATA ROWS (row 10+) ═══
    row = 10
    for idx, g in enumerate(gastos):
        ws.row_dimensions[row].height = 15
        is_alt = (idx % 2 == 1)
        bg = ROW_ALT if is_alt else WHITE

        # Fechas como datetime REAL (Excel las trata como fecha, no como texto
        # 'General'): se ordenan y filtran como fecha. Si no parsean, cae al
        # string DD/MM/AAAA de format_date para no perder el dato.
        fecha_fac_obj   = parse_date_obj(g.get('fechaFac', ''))
        fecha_cobro_obj = parse_date_obj(g.get('fechaCobro', ''))
        fecha_fac   = fecha_fac_obj   if fecha_fac_obj   else format_date(g.get('fechaFac', ''))
        fecha_cobro = fecha_cobro_obj if fecha_cobro_obj else format_date(g.get('fechaCobro', ''))
        forma = FORMA_PAGO_MAP.get(g.get('formaPago', '04'), g.get('formaPago', ''))
        importe_raw = round(g.get('importe', 0), 2)
        iva = round(g.get('iva', 0), 2)
        # ISH/IEPS: impuestos locales (TrasladosLocales de hoteles, IEPS de
        # combustible). Viven en isrTrasladado. Van en su propia columna y se
        # suman al TOTAL (=importe+IVA+ISH−retención) para que cuadre exacto.
        ish = round(g.get('isrTrasladado', 0) or 0, 2)
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
        # Fallback de banco: si el gasto NO tiene banco asignado (típico de
        # data legacy creada antes de v7.81 o cuando el colaborador no es
        # especial y nunca cotejó con Saldos) y el colaborador NO es especial,
        # defaultea a Clara MXN Credito (misma lógica que el frontend en
        # defaultBancoFor). Sin esto, la reconciliación cobrado vs facturado
        # quedaba vacía aunque el colaborador usa Clara por default.
        if not banco and colaborador and colaborador not in COLABORADORES_ESPECIALES:
            banco = 'Clara MXN Credito'
        # USUARIO se resuelve por el folio (poliza_row) usando el mapa
        # invertido de POLIZAS_CLARA. Si el folio del Saldos no matchea
        # ningún colaborador del mapa, cae al nombre del colaborador dueño
        # del reporte (e.g., 'Alejandro Olivar' para sus propias facturas).
        usuario = usuario_por_poliza.get(str(poliza_row).strip()) or colaborador or ''
        # Reconciliación cobrado vs facturado — solo aplica para renglones
        # cuya tarjeta es Clara MXN Credito. Para los demás (BBVA, Monex,
        # Kapital, USD), las 3 celdas quedan vacías (no se reconcilian con
        # Clara CSV en este flujo).
        #
        # IMPORTANTE: FACTURADO muestra el monto de la factura SIN propina
        # (el CFDI no incluye el tip). La propina se contempla en la fórmula
        # de DIFERENCIA: cobrado − (facturado + propina). Así un tip legítimo
        # nunca aparece como discrepancia, y sólo se marca en rojo cuando el
        # banco cargó más que factura + propina.
        es_clara = banco.lower() == 'clara mxn credito'
        if es_clara:
            propina_actual = round(g.get('montoPropina', 0) or 0, 2)
            # FACTURADO = monto de la factura SIN propina (el CFDI no incluye el tip).
            facturado = round(g.get('montoFacturado', 0) or g.get('totalCFDI', 0) or 0, 2)
            # COBRADO = monto REAL que cargó la tarjeta (del CSV de Clara, guardado
            # al validar banco). Es INDEPENDIENTE de factura+propina, por eso
            # DIFERENCIA puede detectar un excedente real. Si la fila no se validó
            # contra el banco (montoCobrado=0), cae al estimado factura+propina,
            # que cuadra → DIFERENCIA 0 hasta que se concilie con el banco.
            monto_cobrado = round(g.get('montoCobrado', 0) or 0, 2)
            cobrado = monto_cobrado if monto_cobrado > 0 else round((g.get('totalCFDI', 0) or 0) + propina_actual, 2)
            # DIFERENCIA = cobrado − (facturado + propina). Se escribe el VALOR
            # calculado (NO una fórmula viva) para que la celda muestre el número
            # correcto en CUALQUIER visor — Excel, Numbers, Quick Look, Google
            # Sheets — sin depender de que recalcule fórmulas. openpyxl no
            # precalcula fórmulas, así que un visor que no recalcula mostraba 0
            # o vacío. Resta la propina sólo cuando existe (no todas la tienen).
            diff_num = round(cobrado - facturado - propina_actual, 2)
            diferencia = diff_num
        else:
            facturado = ''
            cobrado = ''
            diff_num = None
            diferencia = ''
        # RFC de facturas extranjeras (sin RFC mexicano): se muestra "N/A"
        # centrado. Los RFC reales van alineados a la izquierda como siempre.
        rfc_raw = g.get('rfc', '')
        rfc_is_na = str(rfc_raw).strip().upper() in ('NA', 'N/A')
        rfc_disp = 'N/A' if rfc_is_na else rfc_raw
        rfc_align = 'center' if rfc_is_na else 'left'
        # Layout tras agregar ISH/IEPS en L (col 12): RETENCIÓN→M(13),
        # TOTAL→N(14), FORMA→O(15), BANCO→P(16), MONTO USD→Q(17), T/C→R(18),
        # USUARIO→S(19), COBRADO→T(20), FACTURADO→U(21), DIFERENCIA→V(22).
        cells = [
            (2,  rfc_disp,               rfc_align, 'rfc'),
            # PROVEEDOR: SIEMPRE en MAYÚSCULAS en el Excel, sin importar la fuente
            # (XML emisor, OCR, manual) ni el usuario de la plantilla.
            (3,  str(g.get('proveedor', '') or '').upper(), 'left', 'normal_bold'),
            (4,  tipo,                   'center', 'badge_tipo'),
            (5,  poliza_row,             'center', 'normal'),
            (6,  g.get('noFactura', ''), 'center', 'normal'),
            (7,  fecha_fac,              'center', 'date'),
            (8,  fecha_cobro,            'center', 'date'),
            (9,  g.get('concepto', ''),  'left',   'normal'),
            (10, importe_val,            'center', 'currency'),
            (11, iva_val,                'center', 'currency'),
            (12, ish,                    'center', 'currency'),   # ISH/IEPS (nueva)
            (13, ret_val,                'center', 'currency'),
            # TOTAL = IMPORTE + IVA + ISH − RETENCIÓN (fórmula viva). Al incluir
            # ISH, hoteles/combustible cuadran exacto con el total del CFDI.
            # Recalcula solo si el usuario edita J/K/L/M o cuando J/K/M son
            # fórmulas de USD apuntando a la celda editable de T/C.
            (14, f'=J{row}+K{row}+L{row}-M{row}', 'center', 'currency_bold'),
            (15, forma,                  'center', 'badge_pago'),
            (16, banco,                  'center', 'normal'),
            (17, monto_usd,              'center', 'currency'),
            (18, tc_val,                 'center', 'tipocambio'),
            (19, usuario,                'center', 'normal'),
            (20, cobrado,                'center', 'currency'),
            (21, facturado,              'center', 'currency'),
            (22, diferencia,             'center', 'diff'),
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
            style_data_cell(cell, style_type, tipo=tipo, diff_num=diff_num)

        # Side spacer cells keep page bg through the data band.
        ws.cell(row=row, column=1).fill = PatternFill('solid', start_color=BG_PAGE)
        ws.cell(row=row, column=23).fill = PatternFill('solid', start_color=BG_PAGE)

        row += 1

        # ── Optional propina sub-row ──
        # La propina es un renglón COMPLETO que HEREDA del padre: mismo RFC,
        # mismo PROVEEDOR (no dice "Propina" en esa columna), mismo TIPO, misma
        # PÓLIZA, misma FORMA PAGO, mismo BANCO, mismo USUARIO y T/C. Cambia
        # sólo: FACTURA = "N/A", CONCEPTO = "Propina", IMPORTE = monto de la
        # propina, IVA = 0, ISH/IEPS = 0, RETENCIÓN = 0, TOTAL = la propina.
        # COBRADO/FACTURADO/DIFERENCIA quedan vacías (el renglón principal ya
        # contempla la propina en su fórmula de DIFERENCIA — no duplicamos).
        if propina_mxn > 0 or propina_ext > 0:
            ws.row_dimensions[row].height = 15
            propina_bg = 'F0FDF4'  # tinte verde muy sutil para agrupar con el padre

            # Pintar toda la banda primero (mismo fondo) para que las columnas
            # sin contenido (COBRADO/FACTURADO/DIFERENCIA) mantengan el tinte.
            for c in range(2, 23):
                pcell = ws.cell(row=row, column=c)
                pcell.fill = PatternFill('solid', start_color=propina_bg)
                pcell.border = Border(bottom=Side(style='hair', color=BORDER_LIGHT))

            # Valores de la propina, en paralelo al renglón principal (incluye
            # el caso USD: IMPORTE/T/C usan la celda editable de T/C).
            if is_usd_row and propina_ext > 0:
                p_importe_val = f'={propina_ext}*{tc_ref}'
                p_monto_usd   = propina_ext
                p_tc_val      = f'={tc_ref}'
            else:
                p_importe_val = propina_mxn
                p_monto_usd   = propina_ext if propina_ext > 0 else 0
                p_tc_val      = tipo_cambio

            # Concepto: para moneda extranjera incluye símbolo + monto nativo.
            if propina_ext > 0 and moneda_code != 'MXN':
                symbol = CURRENCY_SYMBOLS.get(moneda_code, moneda_code + ' ')
                concepto_prop = f'Propina {symbol}{propina_ext:,.2f} {moneda_code}'
            else:
                concepto_prop = 'Propina'

            pcells = [
                (2,  rfc_disp,               rfc_align, 'rfc'),
                (3,  str(g.get('proveedor', '') or '').upper(), 'left', 'normal_bold'),
                (4,  tipo,                   'center', 'badge_tipo'),
                (5,  poliza_row,             'center', 'normal'),
                (6,  'N/A',                  'center', 'normal'),
                (7,  fecha_fac,              'center', 'date'),
                (8,  fecha_cobro,            'center', 'date'),
                (9,  concepto_prop,          'left',   'normal'),
                (10, p_importe_val,          'center', 'currency'),
                (11, 0,                      'center', 'currency'),
                (12, 0,                      'center', 'currency'),
                (13, 0,                      'center', 'currency'),
                (14, f'=J{row}+K{row}+L{row}-M{row}', 'center', 'currency_bold'),
                (15, forma,                  'center', 'badge_pago'),
                (16, banco,                  'center', 'normal'),
                (17, p_monto_usd,            'center', 'currency'),
                (18, p_tc_val,               'center', 'tipocambio'),
                (19, usuario,                'center', 'normal'),
            ]
            for col, val, align, style_type in pcells:
                cell = ws.cell(row=row, column=col, value=val)
                cell.fill = PatternFill('solid', start_color=propina_bg)
                cell.border = Border(bottom=Side(style='hair', color=BORDER_LIGHT))
                cell.alignment = Alignment(
                    horizontal=align,
                    vertical='center',
                    indent=2 if align == 'left' else 0
                )
                style_data_cell(cell, style_type, tipo=tipo, diff_num=None)

            # Outer spacers stay on page bg so the propina band fits inside
            # the table boundary like every other data row.
            ws.cell(row=row, column=1).fill = PatternFill('solid', start_color=BG_PAGE)
            ws.cell(row=row, column=23).fill = PatternFill('solid', start_color=BG_PAGE)

            row += 1

    # ═══ ADAPTIVE TOTALS — placed right after the last data row ═══
    ws.row_dimensions[row].height = 14  # spacer
    row += 1

    ws.row_dimensions[row].height = 32
    for c in range(2, 23):
        cell = ws.cell(row=row, column=c)
        cell.fill = PatternFill('solid', start_color=SMTO_BLACK)
        cell.border = Border()

    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=9)
    lbl = ws.cell(row=row, column=2, value='TOTAL CUENTA')
    lbl.font = Font(name='Aptos', size=12, bold=True, color=WHITE)
    lbl.alignment = Alignment(horizontal='right', vertical='center', indent=2)
    lbl.fill = PatternFill('solid', start_color=SMTO_BLACK)

    # Same SUM formulas as the KPI cards above so the two views always agree.
    # `data_first` / `data_last` cubren filas principales + sub-filas de propina.
    # Tras agregar ISH/IEPS (L): IMPORTE=J, IVA=K, ISH=L, RETENCIÓN=M, TOTAL=N,
    # MONTO USD=Q. T/C (R), BANCO (P), USUARIO (S) no se totalizan.
    totals = [
        (10, f'=SUM(J{data_first}:J{data_last})', False),  # IMPORTE
        (11, f'=SUM(K{data_first}:K{data_last})', False),  # IVA
        (12, f'=SUM(L{data_first}:L{data_last})', False),  # ISH/IEPS (nueva)
        (13, f'=SUM(M{data_first}:M{data_last})', False),  # RETENCIÓN
        (14, f'=SUM(N{data_first}:N{data_last})', True),   # TOTAL
        (17, f'=SUM(Q{data_first}:Q{data_last})', False),  # MONTO USD
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

    # FORMA PAGO (O), BANCO (P), T/C (R), USUARIO (S), COBRADO (T),
    # FACTURADO (U) y DIFERENCIA (V) en la banda de totales sólo llevan el
    # fill negro (no son agregables o sólo aplican a Clara MXN Credito).
    for c in (15, 16, 18, 19, 20, 21, 22):
        ws.cell(row=row, column=c).fill = PatternFill('solid', start_color=SMTO_BLACK)

    # ═══ FOOTER — one spacer row + a right-aligned version line ═══
    row += 2  # blank spacer + footer row
    ws.row_dimensions[row].height = 18
    ws.merge_cells(start_row=row, start_column=11, end_row=row, end_column=22)
    ft = ws.cell(row=row, column=11)
    ft.value = 'SMTO Engineering · v8.26'
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

    # Unlock cells so they're editable, PERO dejamos las celdas con fórmula
    # en locked=True (el default). Razón: Excel marca con triángulo verde
    # ("fórmula desprotegida") cualquier celda que tenga fórmula y esté
    # desbloqueada. Como la hoja NO está protegida, el estado locked no afecta
    # la edición (todo es editable igual), así que dejar las fórmulas locked
    # solo sirve para quitar esa advertencia y dejar el archivo más limpio.
    _unlocked = Protection(locked=False)
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and cell.value.startswith('='):
                continue  # fórmula → se queda locked=True, sin triángulo
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
                polizas_map = data.get('polizasMap', {}) or {}
            else:
                gastos = data
                colaborador = ''
                poliza_numero = 'N/A'
                polizas_map = {}
            gastos = sanitize(gastos)
            colaborador = sanitize(colaborador)
            poliza_numero = sanitize(poliza_numero)
            polizas_map = sanitize(polizas_map) if isinstance(polizas_map, dict) else {}
            wb = build_workbook(gastos, colaborador, poliza_numero, polizas_map)
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
