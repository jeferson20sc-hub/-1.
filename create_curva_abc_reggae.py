#!/usr/bin/env python3
"""
Curva ABC Reggae - Professional Dashboard
Colors: Green #009B3A, Gold #FFD700, Red #CC0000, Black #000000
Charts use static pre-computed data for reliable rendering.
"""

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.chart import BarChart, LineChart, Reference, Series
from openpyxl.formatting.rule import FormulaRule

# ===================== REGGAE COLORS =====================
GREEN       = "009B3A"
GOLD        = "FFD700"
RED         = "CC0000"
BLACK       = "000000"
WHITE       = "FFFFFF"
GREEN_DARK  = "006B28"
GREEN_LIGHT = "C8F5DB"
GOLD_DARK   = "C8A400"
GOLD_LIGHT  = "FFF5B0"
RED_DARK    = "990000"
RED_LIGHT   = "FFD0D0"
GRAY_DARK   = "B0B0B0"
DARK_BG     = "1A1A1A"


def fill(hex_color):
    return PatternFill(fill_type="solid", fgColor=hex_color)

def font(bold=False, color=WHITE, size=11, italic=False):
    return Font(bold=bold, color=color, size=size, italic=italic, name="Calibri")

def border(style="thin", color=BLACK):
    s = Side(style=style, color=color)
    return Border(left=s, right=s, top=s, bottom=s)

def align(h="center", wrap=False):
    return Alignment(horizontal=h, vertical="center", wrap_text=wrap)


def outer_border(ws, min_row, min_col, max_row, max_col, color=BLACK, style="medium"):
    thick = Side(style=style, color=color)
    thin  = Side(style="thin", color=color)
    for row in ws.iter_rows(min_row=min_row, max_row=max_row,
                             min_col=min_col, max_col=max_col):
        for cell in row:
            r, c = cell.row, cell.column
            cell.border = Border(
                top    = thick if r == min_row else thin,
                bottom = thick if r == max_row else thin,
                left   = thick if c == min_col else thin,
                right  = thick if c == max_col else thin,
            )


# ===================== SHEET ITEM DATA =====================
SHEET_CONFIGS = [
    {
        "name": "Primeira",
        "title": "CURVA ABC — PRIMEIRA",
        "items": [
            ("Servidor Backup",          3,   2, 25000),
            ('Monitor 27" 4K',           5,  10,  3000),
            ("Notebook Administrativo",  9,   6,  4500),
            ("Nobreak Profissional",      1,   5,  2000),
            ("Switch Gerenciavel",        7,   4,  1500),
            ("Mouse sem fio",             4,  40,    50),
            ("Patch Cord 1m",             8, 100,    15),
            ("Teclado USB",               6,  80,    35),
            ("Cabo Rede Cat6 (m)",        2, 200,     8),
            ("Mousepad Slim",            10, 120,    12),
        ],
    },
    {
        "name": "01",
        "title": "CURVA ABC — ABA 01",
        "items": [
            ("Notebook Executivo i7",      10,  35,  6500),
            ("Servidor de Dados Pro",       2,  12, 15000),
            ("Impressora Multifuncional",  17,  55,  1100),
            ("Toner Impressora Laser",     12,  68,   680),
            ("Memoria RAM 16GB",           19,  45,   380),
            ('Monitor LED 24"',             4,  20,   850),
            ("Webcam Full HD",             14,  82,   180),
            ("Roteador Wi-Fi 6",           11,  15,   650),
            ("SSD 480GB Sata",             15,  55,   190),
            ("Nobreak 1500VA",              6,   8,  1200),
            ("HD Externo 2TB",             20,  18,   420),
            ("Teclado Mecanico RGB",        5,  30,   250),
            ("Pen Drive 64GB",              9, 160,    45),
            ("Headset com Microfone",      13,  64,   110),
            ("Pacote de Papel A4",          7, 237,    28),
            ("Filtro de Linha 5 Tom.",     18, 110,    40),
            ("Cabo HDMI 2m",               1, 150,    25),
            ("Adaptador USB-C",           16,  90,    35),
            ("Mouse Optico Simples",        3, 120,    15),
            ("Patch Cord RJ45 1m",          8, 200,    12),
        ],
    },
    {
        "name": "02",
        "title": "CURVA ABC — ABA 02",
        "items": [
            ("Servidor Rack Dell Pro",     2,  15, 22500),
            ("Workstation Grafica",       17,  19, 12000),
            ("Notebook i7 32GB RAM",      10,  28,  7200),
            ("Switch Gerenciavel 48p",     8,  10,  4800),
            ("Placa de Video RTX",        22,   8,  5500),
            ("Storage NAS 40TB",          21,   2, 18500),
            ("Memoria RAM 16GB",          19,  45,   380),
            ("Toner Impressora Laser",    12,  40,   320),
            ('Monitor LED 24"',            4,  15,   850),
            ("Roteador Wi-Fi 6",          11,  25,   450),
            ("SSD 480GB Sata",            15,  55,   190),
            ("Nobreak 1500VA",             6,   8,  1200),
            ("HD Externo 2TB",            20,  18,   420),
            ("Teclado Mecanico RGB",       5,  30,   250),
            ("Pen Drive 64GB",             9, 150,    45),
            ("Headset com Microfone",     13,  60,   110),
            ("Pacote de Papel A4",         7, 200,    28),
            ("Teclado de Entrada",        29, 100,    45),
            ("Filtro de Linha",           18, 110,    40),
            ("Webcam Full HD",            14,  22,   180),
            ("Pilhas AA (Pacote)",        23, 300,    12),
            ("Adaptador USB-C",           16,  90,    35),
            ("Suporte para Monitor",      28,  20,   140),
            ("Cabo HDMI 2m",               1,  85,    25),
            ("Mousepad Simples",          24, 250,     8),
            ("Mouse Optico Simples",       3, 120,    15),
            ("Organizador de Cabos",      25, 180,    10),
            ("Ar comprimido (Lata)",      27,  35,    45),
            ("Pasta Termica",             26,  45,    35),
            ("Conector RJ45 (Cento)",     30,  50,    30),
        ],
    },
]


def classify_abc(items):
    """Pre-compute ABC classification, sorted by VT descending."""
    rows = []
    for desc, cod, qty, cost in items:
        rows.append({"desc": desc, "cod": cod, "qty": qty, "cost": cost, "vt": qty * cost})
    rows.sort(key=lambda x: x["vt"], reverse=True)
    total = sum(r["vt"] for r in rows)
    cumsum = 0
    for r in rows:
        cumsum += r["vt"]
        r["pct_ind"] = r["vt"] / total if total else 0
        r["pct_acc"] = cumsum / total if total else 0
        r["abc"] = "A" if r["pct_acc"] <= 0.80 else ("B" if r["pct_acc"] <= 0.95 else "C")
    return rows, total


def write_kpi_label(ws, row, col, label, label_color, bg=DARK_BG, span=2):
    ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col + span - 1)
    c = ws.cell(row=row, column=col, value=label)
    c.fill = fill(bg); c.font = font(bold=True, color=label_color, size=9)
    c.alignment = align()

def write_kpi_value(ws, row, col, value, bg, fmt=None, span=2):
    ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col + span - 1)
    c = ws.cell(row=row, column=col, value=value)
    c.fill = fill(bg); c.font = font(bold=True, color=WHITE, size=14)
    c.alignment = align()
    if fmt: c.number_format = fmt


# ===================== DATA SHEET =====================
def create_sheet(wb, cfg):
    ws = wb.create_sheet(title=cfg["name"])
    items = cfg["items"]
    DS = 10   # data_start row
    n  = len(items)
    DE = DS + n - 1   # data_end row

    sorted_data, total_vt = classify_abc(items)

    # Build lookup: original desc -> abc class (for row coloring)
    abc_lookup = {(d["desc"], d["cod"]): d["abc"] for d in sorted_data}

    # ---- Column widths ----
    col_w = {
        "A": 7,  "B": 32, "C": 10, "D": 10, "E": 14,
        "F": 16, "G": 13, "H": 13,
        # spacer
        "I": 2,
        # summary panel
        "J": 10, "K": 16, "L": 10, "M": 10, "N": 22,
        # hidden chart cols
        "P": 30, "Q": 14, "R": 14, "S": 14,
    }
    for c, w in col_w.items():
        ws.column_dimensions[c].width = w

    # ---- Row heights ----
    ws.row_dimensions[1].height = 48
    for r in range(2, 6):
        ws.row_dimensions[r].height = 26
    ws.row_dimensions[6].height = 16
    ws.row_dimensions[7].height = 20
    ws.row_dimensions[8].height = 24
    ws.row_dimensions[9].height = 6
    for r in range(DS, DE + 2):
        ws.row_dimensions[r].height = 19

    # ---- Row 1: Title bar ----
    ws.merge_cells("A1:H1")
    t = ws["A1"]
    t.value = cfg["title"]
    t.fill = fill(BLACK); t.font = font(bold=True, color=GREEN, size=22)
    t.alignment = align()

    ws.merge_cells("J1:N1")
    leg = ws["J1"]
    leg.value = "A = ate 80%   |   B = 80–95%   |   C = 95–100%"
    leg.fill = fill(BLACK); leg.font = font(color=GOLD, size=10)
    leg.alignment = align()

    # Reggae stripe (I1)
    for col_letter, c in [("I", BLACK)]:
        ws[f"{col_letter}1"].fill = fill(c)

    # ---- Rows 2–5: KPI panel ----
    for r in range(2, 6):
        for c in range(1, 9):
            ws.cell(row=r, column=c).fill = fill(DARK_BG)

    # Row 2–3 KPIs
    write_kpi_label(ws, 2, 1, "V.T. TOTAL",   GREEN)
    write_kpi_value(ws, 3, 1, f"=SUM($F${DS}:$F${DE})", GREEN_DARK, fmt='R$ #,##0.00')

    write_kpi_label(ws, 2, 3, "V.T. CLASSE A", GREEN)
    write_kpi_value(ws, 3, 3, f'=SUMIF($A${DS}:$A${DE},"A",$F${DS}:$F${DE})', GREEN, fmt='R$ #,##0.00')

    write_kpi_label(ws, 2, 5, "% ITENS A",  GOLD)
    write_kpi_value(ws, 3, 5,
        f'=IFERROR(COUNTIF($A${DS}:$A${DE},"A")/COUNTA($B${DS}:$B${DE}),0)',
        GOLD_DARK, fmt='0.0%')

    write_kpi_label(ws, 2, 7, "MAIOR V.T.", RED_LIGHT)
    write_kpi_value(ws, 3, 7, f'=MAX($F${DS}:$F${DE})', RED_DARK, fmt='R$ #,##0.00')

    # Row 4–5 KPIs
    write_kpi_label(ws, 4, 1, "MENOR V.T.", GREEN)
    write_kpi_value(ws, 5, 1, f'=MIN($F${DS}:$F${DE})', GREEN_DARK, fmt='R$ #,##0.00')

    write_kpi_label(ws, 4, 3, "STATUS", GREEN)
    write_kpi_value(ws, 5, 3,
        f'=IF(IFERROR(MAX($H${DS}:$H${DE}),0)>1.0001,"VERIFICAR","OK")',
        GREEN)

    write_kpi_label(ws, 4, 5, "LIMITE A", GOLD)
    write_kpi_value(ws, 5, 5, 0.80, GOLD_DARK, fmt='0%')

    write_kpi_label(ws, 4, 7, "LIMITE B", RED_LIGHT)
    write_kpi_value(ws, 5, 7, 0.95, RED_DARK, fmt='0%')

    outer_border(ws, 2, 1, 5, 8, color=GREEN, style="medium")

    # ---- Row 6: subtitle ----
    ws.merge_cells("A6:H6")
    s = ws["A6"]
    s.value = "  Edite somente Quant. e Custo Unit. — ABC, V.T. e % recalculam automaticamente."
    s.fill = fill(GREEN_DARK); s.font = font(italic=True, color=GOLD_LIGHT, size=8)
    s.alignment = align(h="left")

    # ---- Row 7: right-side summary header ----
    for c, label in [(10,"Classe"),(11,"Total V.T."),(12,"% Total"),(13,"Qtd. Itens"),(14,"Barra")]:
        cell = ws.cell(row=7, column=c, value=label)
        cell.fill = fill(BLACK); cell.font = font(bold=True, size=9)
        cell.alignment = align(); cell.border = border(color=GRAY_DARK)

    # ---- Rows 8–10: class summary ----
    for cls, bg, r in [("A", GREEN, 8), ("B", GOLD_DARK, 9), ("C", RED_DARK, 10)]:
        fg = BLACK if cls == "B" else WHITE
        cells_data = [
            (10, cls),
            (11, f'=SUMIF($A${DS}:$A${DE},"{cls}",$F${DS}:$F${DE})'),
            (12, f'=IFERROR(K{r}/SUM($F${DS}:$F${DE}),0)'),
            (13, f'=COUNTIF($A${DS}:$A${DE},"{cls}")'),
            (14, f'=REPT(CHAR(9608),ROUND(L{r}*20,0))'),
        ]
        fmts = [None, 'R$ #,##0', '0.0%', None, None]
        for (col, val), fmt in zip(cells_data, fmts):
            c = ws.cell(row=r, column=col, value=val)
            c.fill = fill(bg); c.font = font(bold=True, color=fg, size=11 if col != 14 else 8)
            c.alignment = align(h="left" if col == 14 else "center")
            c.border = border(color=BLACK)
            if fmt: c.number_format = fmt

    # ---- Row 8: column headers ----
    ws.row_dimensions[8].height = 26
    headers = [
        (1, "ABC",             BLACK, GREEN),
        (2, "Descricao",       BLACK, WHITE),
        (3, "Cod.",            BLACK, WHITE),
        (4, "Quant.",          BLACK, GOLD),
        (5, "Custo Unit.(R$)", BLACK, GOLD),
        (6, "V.T. (R$)",       BLACK, GREEN),
        (7, "% Individual",    BLACK, WHITE),
        (8, "% Acumulado",     BLACK, WHITE),
    ]
    for col, label, bg, fg in headers:
        c = ws.cell(row=8, column=col, value=label)
        c.fill = fill(bg); c.font = font(bold=True, color=fg, size=10)
        c.alignment = align(); c.border = border(color=GREEN_DARK)

    # Hidden chart-data header row
    for col, label in [(16,"Label"),(17,"VT Ord."),(18,"% Acum."),(19,"VT A"),(20,"VT B"),(21,"VT C")]:
        c = ws.cell(row=8, column=col, value=label)
        c.fill = fill(DARK_BG); c.font = font(bold=True, color=GRAY_DARK, size=8)
        c.alignment = align()

    # ---- Data rows (formulas for main table, STATIC values for chart cols) ----
    from openpyxl.styles.protection import Protection
    unlock = Protection(locked=False)

    for i, (desc, cod, qty, cost) in enumerate(items):
        row = DS + i
        item_abc = abc_lookup.get((desc, cod), "C")

        if item_abc == "A":
            rf = fill(GREEN_LIGHT); abc_bg = GREEN
        elif item_abc == "B":
            rf = fill(GOLD_LIGHT);  abc_bg = GOLD_DARK
        else:
            rf = fill(RED_LIGHT);   abc_bg = RED_DARK

        # Col A: ABC formula (recalculates)
        c = ws.cell(row=row, column=1,
            value=f'=IF($H{row}<=$E$5,"A",IF($H{row}<=$G$5,"B","C"))')
        c.fill = fill(abc_bg); c.font = font(bold=True, size=11)
        c.alignment = align(); c.border = border(color=BLACK)

        # Col B: Description
        c = ws.cell(row=row, column=2, value=desc)
        c.fill = rf; c.font = font(color=BLACK, size=10)
        c.alignment = align(h="left"); c.border = border(color=GRAY_DARK)

        # Col C: Code
        c = ws.cell(row=row, column=3, value=cod)
        c.fill = rf; c.font = font(color=BLACK, size=10)
        c.alignment = align(); c.border = border(color=GRAY_DARK)

        # Col D: Quantity (editable)
        c = ws.cell(row=row, column=4, value=qty)
        c.fill = fill(WHITE); c.font = font(bold=True, color=BLACK, size=10)
        c.alignment = align()
        c.border = Border(
            left=Side(style="medium", color=GREEN_DARK),
            right=Side(style="medium", color=GREEN_DARK),
            top=Side(style="thin", color=GRAY_DARK),
            bottom=Side(style="thin", color=GRAY_DARK))
        c.number_format = '#,##0'
        c.protection = unlock

        # Col E: Unit cost (editable)
        c = ws.cell(row=row, column=5, value=cost)
        c.fill = fill(WHITE); c.font = font(bold=True, color=BLACK, size=10)
        c.alignment = align(h="right")
        c.border = Border(
            left=Side(style="medium", color=GREEN_DARK),
            right=Side(style="medium", color=GREEN_DARK),
            top=Side(style="thin", color=GRAY_DARK),
            bottom=Side(style="thin", color=GRAY_DARK))
        c.number_format = 'R$ #,##0.00'
        c.protection = unlock

        # Col F: VT formula
        c = ws.cell(row=row, column=6, value=f'=IFERROR(D{row}*E{row},0)')
        c.fill = rf; c.font = font(bold=True, color=BLACK, size=10)
        c.alignment = align(h="right"); c.border = border(color=GRAY_DARK)
        c.number_format = 'R$ #,##0.00'

        # Col G: % Individual
        c = ws.cell(row=row, column=7,
            value=f'=IFERROR(F{row}/SUM($F${DS}:$F${DE}),0)')
        c.fill = rf; c.font = font(color=BLACK, size=10)
        c.alignment = align(); c.border = border(color=GRAY_DARK)
        c.number_format = '0.00%'

        # Col H: % Accumulated (SUMIF approach - proper Pareto)
        c = ws.cell(row=row, column=8,
            value=f'=IFERROR(SUMIF($F${DS}:$F${DE},">="&F{row},$F${DS}:$F${DE})/SUM($F${DS}:$F${DE}),0)')
        c.fill = rf; c.font = font(color=BLACK, size=10)
        c.alignment = align(); c.border = border(color=GRAY_DARK)
        c.number_format = '0.00%'

    # ---- STATIC chart data (pre-computed Python values) ----
    # Col P(16): labels, Q(17): VT sorted, R(18): % acc, S(19): VT_A, T(20): VT_B, U(21): VT_C
    for i, d in enumerate(sorted_data):
        row = DS + i
        label = f"{d['desc'][:20]} ({d['abc']})"

        c = ws.cell(row=row, column=16, value=label)
        c.fill = fill(DARK_BG); c.font = font(color=GRAY_DARK, size=8)
        c.alignment = align(h="left")

        c = ws.cell(row=row, column=17, value=d["vt"])
        c.fill = fill(DARK_BG); c.font = font(color=GRAY_DARK, size=8)
        c.number_format = '#,##0'

        c = ws.cell(row=row, column=18, value=round(d["pct_acc"], 4))
        c.fill = fill(DARK_BG); c.font = font(color=GRAY_DARK, size=8)
        c.number_format = '0.0%'

        # Separate VT by class (0 for non-matching, not NA)
        c = ws.cell(row=row, column=19, value=d["vt"] if d["abc"] == "A" else 0)
        c.fill = fill(GREEN_DARK); c.font = font(color=GRAY_DARK, size=8)
        c.number_format = '#,##0'

        c = ws.cell(row=row, column=20, value=d["vt"] if d["abc"] == "B" else 0)
        c.fill = fill(GOLD_DARK); c.font = font(color=GRAY_DARK, size=8)
        c.number_format = '#,##0'

        c = ws.cell(row=row, column=21, value=d["vt"] if d["abc"] == "C" else 0)
        c.fill = fill(RED_DARK); c.font = font(color=GRAY_DARK, size=8)
        c.number_format = '#,##0'

    # ---- Total row ----
    tr = DE + 1
    ws.row_dimensions[tr].height = 22
    ws.merge_cells(start_row=tr, start_column=1, end_row=tr, end_column=3)
    c = ws.cell(row=tr, column=1, value="TOTAL GERAL")
    c.fill = fill(BLACK); c.font = font(bold=True, color=GOLD, size=11)
    c.alignment = align()

    c = ws.cell(row=tr, column=4, value=f'=SUM(D{DS}:D{DE})')
    c.fill = fill(BLACK); c.font = font(bold=True, size=11)
    c.alignment = align(); c.number_format = '#,##0'

    ws.cell(row=tr, column=5).fill = fill(BLACK)

    c = ws.cell(row=tr, column=6, value=f'=SUM(F{DS}:F{DE})')
    c.fill = fill(GREEN); c.font = font(bold=True, size=11)
    c.alignment = align(h="right"); c.number_format = 'R$ #,##0.00'

    for col in [7, 8]:
        c = ws.cell(row=tr, column=col, value=1.0)
        c.fill = fill(BLACK); c.font = font(bold=True, size=11)
        c.alignment = align(); c.number_format = '0%'

    for col in range(1, 9):
        c = ws.cell(row=tr, column=col)
        c.border = Border(
            top=Side(style="medium", color=GREEN),
            bottom=Side(style="medium", color=GREEN),
            left=Side(style="thin", color=GREEN),
            right=Side(style="thin", color=GREEN))

    # ---- Charts (using STATIC pre-computed columns P–U) ----
    chart_row = tr + 3

    # Chart 1: Bar chart — VT por classe (A=green, B=gold, C=red)
    bar = BarChart()
    bar.type      = "col"
    bar.grouping  = "stacked"
    bar.title     = f"Curva ABC — Valor Total por Item  ({cfg['name']})"
    bar.y_axis.title = "Valor Total (R$)"
    bar.x_axis.title = "Itens (ordenados por V.T. decrescente)"
    bar.style     = 10
    bar.width     = 26
    bar.height    = 14

    ser_a = Series(Reference(ws, min_col=19, min_row=DS, max_row=DE), title="Classe A")
    ser_a.graphicalProperties.solidFill = GREEN
    ser_a.graphicalProperties.line.solidFill = GREEN_DARK

    ser_b = Series(Reference(ws, min_col=20, min_row=DS, max_row=DE), title="Classe B")
    ser_b.graphicalProperties.solidFill = GOLD
    ser_b.graphicalProperties.line.solidFill = GOLD_DARK

    ser_c = Series(Reference(ws, min_col=21, min_row=DS, max_row=DE), title="Classe C")
    ser_c.graphicalProperties.solidFill = RED
    ser_c.graphicalProperties.line.solidFill = RED_DARK

    bar.series.append(ser_a)
    bar.series.append(ser_b)
    bar.series.append(ser_c)
    bar.set_categories(Reference(ws, min_col=16, min_row=DS, max_row=DE))
    ws.add_chart(bar, f"A{chart_row}")

    # Chart 2: Line chart — Curva de Pareto (% acumulado)
    line = LineChart()
    line.title        = f"Curva de Pareto — % Acumulado  ({cfg['name']})"
    line.y_axis.title = "% Acumulado"
    line.x_axis.title = "Itens (ordem decrescente de V.T.)"
    line.y_axis.numFmt = '0%'
    line.y_axis.scaling.min = 0
    line.y_axis.scaling.max = 1
    line.style     = 10
    line.width     = 26
    line.height    = 14

    ser_p = Series(Reference(ws, min_col=18, min_row=DS, max_row=DE), title="% Acumulado")
    ser_p.graphicalProperties.line.solidFill = GREEN
    ser_p.graphicalProperties.line.width     = 28000
    ser_p.smooth = True
    line.series.append(ser_p)

    # 80% reference
    ser_80 = Series(Reference(ws, min_col=17, min_row=DS - 1, max_row=DS - 1), title="Limite A (80%)")
    line.series.append(ser_80)

    line.set_categories(Reference(ws, min_col=16, min_row=DS, max_row=DE))
    # Place line chart to the right
    right_col = get_column_letter(14)  # col N
    ws.add_chart(line, f"{right_col}{chart_row}")

    # ---- Footer banner ----
    footer_row = chart_row + 30
    ws.row_dimensions[footer_row].height = 16
    ws.merge_cells(f"A{footer_row}:H{footer_row}")
    fc = ws.cell(row=footer_row, column=1,
                 value=f"  CURVA ABC REGGAE  |  {cfg['title']}  |  One Love, One Data  |  2026")
    fc.fill = fill(BLACK); fc.font = font(bold=True, color=GREEN, size=9)
    fc.alignment = align()

    # ---- Conditional formatting ----
    rng = f"A{DS}:H{DE}"
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f'$A{DS}="A"'], fill=PatternFill(fill_type="solid", fgColor=GREEN_LIGHT),
        font=Font(color=BLACK)))
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f'$A{DS}="B"'], fill=PatternFill(fill_type="solid", fgColor=GOLD_LIGHT),
        font=Font(color=BLACK)))
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f'$A{DS}="C"'], fill=PatternFill(fill_type="solid", fgColor=RED_LIGHT),
        font=Font(color=BLACK)))

    # ---- Sheet protection ----
    ws.protection.sheet    = True
    ws.protection.password = "abc2026."
    ws.protection.enable()
    ws.protection.selectLockedCells   = False
    ws.protection.selectUnlockedCells = False

    # ---- Freeze panes ----
    ws.freeze_panes = f"A{DS}"

    # ---- Tab color ----
    tab_colors = {"Primeira": GREEN, "01": GOLD_DARK, "02": RED_DARK}
    ws.sheet_properties.tabColor = tab_colors.get(cfg["name"], GREEN)

    return ws


# ===================== INDEX SHEET =====================
def create_index(wb):
    ws = wb.create_sheet(title="INICIO", index=0)

    for col_letter, w in {"A": 3, "B": 28, "C": 28, "D": 28, "E": 28, "F": 3}.items():
        ws.column_dimensions[col_letter].width = w
    for r in range(1, 40):
        ws.row_dimensions[r].height = 20
    ws.row_dimensions[2].height = 64
    ws.row_dimensions[3].height = 28

    for r in range(1, 40):
        for c in range(1, 7):
            ws.cell(row=r, column=c).fill = fill(DARK_BG)

    # Reggae stripes row 1
    for col, clr in [(2, GREEN), (3, GOLD), (4, RED), (5, BLACK)]:
        ws.cell(row=1, column=col).fill = fill(clr)

    ws.merge_cells("B2:E2")
    t = ws["B2"]
    t.value = "CURVA ABC REGGAE"
    t.fill = fill(BLACK); t.font = Font(bold=True, color=GREEN, size=30, name="Calibri")
    t.alignment = align()

    ws.merge_cells("B3:E3")
    s = ws["B3"]
    s.value = "Dashboard Profissional de Analise de Estoque"
    s.fill = fill(BLACK); s.font = font(italic=True, color=GOLD, size=13)
    s.alignment = align()

    ws.merge_cells("B4:E4")
    ws["B4"].fill = fill(BLACK)
    for col, clr in [(2, GREEN), (3, GOLD), (4, RED), (5, GREEN)]:
        ws.cell(row=4, column=col).fill = fill(clr)

    info = [
        (6,  "Classificacao ABC",
             ["A = ate 80% do valor total", "B = 80% a 95% do valor", "C = 95% a 100% do valor"], GREEN),
        (11, "Abas Disponiveis",
             ["Primeira: 10 itens", "Aba 01: 20 itens", "Aba 02: 30 itens"], GOLD_DARK),
        (16, "Como Usar",
             ["1. Va para a aba desejada", "2. Edite Quant. e Custo Unit.", "3. ABC recalcula automaticamente"], RED_DARK),
        (21, "Protecao da Planilha",
             ["Senha: abc2026.", "Colunas D e E liberadas para edicao", "Demais celulas protegidas"], GREEN_DARK),
    ]

    for start_row, title_txt, lines, clr in info:
        ws.merge_cells(start_row=start_row, start_column=2, end_row=start_row, end_column=5)
        c = ws.cell(row=start_row, column=2, value=title_txt)
        c.fill = fill(clr); c.font = font(bold=True, size=12)
        c.alignment = align(); c.border = border(color=BLACK)

        for j, line in enumerate(lines):
            r = start_row + 1 + j
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
            c = ws.cell(row=r, column=2, value=line)
            c.fill = fill(DARK_BG)
            c.font = font(color=GOLD_LIGHT if clr == GOLD_DARK else WHITE, size=10)
            c.alignment = align(h="left"); c.border = border(color=clr)

    ws.merge_cells("B36:E36")
    f = ws["B36"]
    f.value = "One Love, One Data  |  CURVA ABC REGGAE  |  2026"
    f.fill = fill(BLACK); f.font = font(bold=True, color=GREEN, size=9)
    f.alignment = align()

    ws.sheet_properties.tabColor = BLACK
    return ws


# ===================== MAIN =====================
def main():
    output = "/home/user/Zyth/curva_abc_reggae.xlsx"
    wb = Workbook()
    wb.remove(wb.active)

    create_index(wb)

    for cfg in SHEET_CONFIGS:
        print(f"  Building sheet: {cfg['name']} ({len(cfg['items'])} items)...")
        create_sheet(wb, cfg)

    wb.properties.title   = "Curva ABC Reggae"
    wb.properties.subject = "Analise de Estoque ABC - Reggae"
    wb.properties.creator = "CURVA ABC REGGAE"

    wb.save(output)
    import os
    size = os.path.getsize(output)
    print(f"Saved: {output}  ({size/1024:.1f} KB)")


if __name__ == "__main__":
    main()
