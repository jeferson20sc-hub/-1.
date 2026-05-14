#!/usr/bin/env python3
"""
🇯🇲 CURVA ABC REGGAE — One Love, One Excel! 🇯🇲
3 abas: INICIO  |  CURVA ABC  |  SCRIPTS
Gráfico: bandas de fundo (A=verde, B=dourado, C=vermelho) + barras + curva Pareto
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.chart import AreaChart, BarChart, LineChart, Reference, Series
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.styles.protection import Protection

# ── Paleta Reggae ──────────────────────────────────────────────────────────────
GREEN_DARK  = "006B28"
GREEN       = "009B3A"
GREEN_LIGHT = "C8F5DB"
GREEN_CHART = "ADEBAD"   # zona A no gráfico
GOLD_DARK   = "C8A400"
GOLD        = "FFD700"
GOLD_LIGHT  = "FFF5B0"
GOLD_CHART  = "FFE87C"   # zona B no gráfico
RED_DARK    = "990000"
RED         = "CC0000"
RED_LIGHT   = "FFD0D0"
RED_CHART   = "FFAD99"   # zona C no gráfico (laranja-salmão como na imagem)
BLACK       = "000000"
WHITE       = "FFFFFF"
DARK_BG     = "1A1A1A"
BLUE_BAR    = "2E75B6"   # azul das barras (igual imagem de referência)
GRAY_DARK   = "B0B0B0"
GRAY_MID    = "E0E0E0"

# ── Helpers ────────────────────────────────────────────────────────────────────
def F(hex_color):
    return PatternFill(fill_type="solid", fgColor=hex_color)

def Ft(bold=False, color=WHITE, size=11, italic=False, name="Calibri"):
    return Font(bold=bold, color=color, size=size, italic=italic, name=name)

def B(style="thin", color=BLACK):
    s = Side(style=style, color=color)
    return Border(left=s, right=s, top=s, bottom=s)

def A(h="center", wrap=False):
    return Alignment(horizontal=h, vertical="center", wrap_text=wrap)

def mc(ws, r1, c1, r2, c2):
    ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)

def cell(ws, row, col, value="", fill_c=None, font=None, border=None, align=None, fmt=None):
    c = ws.cell(row=row, column=col, value=value)
    if fill_c:  c.fill      = F(fill_c)
    if font:    c.font      = font
    if border:  c.border    = border
    if align:   c.alignment = align
    if fmt:     c.number_format = fmt
    return c

def outer_border(ws, r1, c1, r2, c2, color=GREEN, thick="medium"):
    t = Side(style=thick, color=color)
    n = Side(style="thin", color=color)
    for row in ws.iter_rows(r1, r2, c1, c2):
        for c in row:
            r, cc = c.row, c.column
            c.border = Border(
                top    = t if r  == r1 else n,
                bottom = t if r  == r2 else n,
                left   = t if cc == c1 else n,
                right  = t if cc == c2 else n,
            )

# ── Items ──────────────────────────────────────────────────────────────────────
ITEMS = [
    ("SKU001", "Servidor Rack Dell Pro",        15, 22500),
    ("SKU002", "Storage NAS 40TB",               3, 18500),
    ("SKU003", "Notebook i7 32GB RAM",           28,  7200),
    ("SKU004", "Workstation Grafica RTX",         8,  9500),
    ("SKU005", "Placa de Video RTX 4070",         6,  5500),
    ("SKU006", "Switch Gerenciavel 48p",         10,  4800),
    ("SKU007", "Impressora Multifuncional",       55,  1100),
    ("SKU008", "Nobreak 1500VA",                  8,  1200),
    ("SKU009", "Monitor LED 24\"",               20,    850),
    ("SKU010", "Roteador Wi-Fi 6",               25,    650),
    ("SKU011", "Toner Impressora Laser",          68,    680),
    ("SKU012", "HD Externo 2TB",                 18,    420),
    ("SKU013", "Memoria RAM 16GB",               45,    380),
    ("SKU014", "Webcam Full HD",                  82,    180),
    ("SKU015", "SSD 480GB Sata",                 55,    190),
    ("SKU016", "Teclado Mecanico RGB",            30,    250),
    ("SKU017", "Pen Drive 64GB",                 160,     45),
    ("SKU018", "Headset com Microfone",            64,    110),
    ("SKU019", "Mouse Optico Simples",            120,     15),
    ("SKU020", "Patch Cord RJ45 1m",             200,     12),
]

def classify(items, lim_a=0.80, lim_b=0.95):
    rows = [{"sku": s, "desc": d, "qty": q, "cost": c, "vt": q*c}
            for s, d, q, c in items]
    rows.sort(key=lambda x: x["vt"], reverse=True)
    total = sum(r["vt"] for r in rows)
    cumsum = 0
    for r in rows:
        cumsum += r["vt"]
        r["pct_ind"] = r["vt"] / total if total else 0
        r["pct_acc"] = cumsum / total if total else 0
        r["abc"] = "A" if r["pct_acc"] <= lim_a else ("B" if r["pct_acc"] <= lim_b else "C")
    return rows, total


# ═══════════════════════════════════════════════════════════════════════════════
#  ABA 1 — 🇯🇲 INICIO
# ═══════════════════════════════════════════════════════════════════════════════
def create_inicio(wb):
    ws = wb.create_sheet(title="\U0001f1ef\U0001f1f2 INICIO", index=0)

    for c, w in {"A":3,"B":22,"C":22,"D":22,"E":22,"F":3}.items():
        ws.column_dimensions[c].width = w
    for r in range(1, 55):
        ws.row_dimensions[r].height = 20
        for col in range(1, 7):
            ws.cell(row=r, column=col).fill = F(DARK_BG)

    # Faixa reggae topo
    ws.row_dimensions[1].height = 8
    for col, clr in [(2,GREEN),(3,GOLD),(4,RED),(5,BLACK)]:
        ws.cell(row=1, column=col).fill = F(clr)

    # Título principal
    ws.row_dimensions[2].height = 80
    mc(ws, 2, 2, 2, 5)
    c = ws.cell(row=2, column=2, value="\U0001f1ef\U0001f1f2 CURVA ABC REGGAE")
    c.fill = F(BLACK)
    c.font = Font(bold=True, color=GREEN, size=32, name="Calibri")
    c.alignment = A()

    ws.row_dimensions[3].height = 28
    mc(ws, 3, 2, 3, 5)
    c = ws.cell(row=3, column=2, value="One Love, One Excel! ✌️  —  Dashboard Profissional de Curva ABC")
    c.fill = F(BLACK)
    c.font = Ft(italic=True, color=GOLD, size=14)
    c.alignment = A()

    # Faixas reggae fundo do título
    ws.row_dimensions[4].height = 8
    for col, clr in [(2,RED),(3,GOLD),(4,GREEN),(5,BLACK)]:
        ws.cell(row=4, column=col).fill = F(clr)

    # Tagline
    ws.row_dimensions[5].height = 30
    mc(ws, 5, 2, 5, 5)
    c = ws.cell(row=5, column=2,
                value="\U0001f981  Classifique seus produtos como um verdadeiro Rasta do Excel  \U0001f33f")
    c.fill = F(BLACK)
    c.font = Ft(italic=True, color=GOLD_LIGHT, size=11)
    c.alignment = A()

    # Emojis decorativos
    ws.row_dimensions[6].height = 24
    mc(ws, 6, 2, 6, 5)
    c = ws.cell(row=6, column=2,
                value="   \U0001f3b5  \U0001f3b8  \U0001f941  ☀️  \U0001f30a  ✈️  \U0001f33f  \U0001f981  ❤️  \U0001f3b5")
    c.fill = F(GREEN_DARK)
    c.font = Ft(size=14, color=WHITE)
    c.alignment = A()

    ws.row_dimensions[7].height = 8

    # Cards de informação
    def info_card(row, icon, title, lines, bg_color):
        ws.row_dimensions[row].height = 26
        mc(ws, row, 2, row, 5)
        c = ws.cell(row=row, column=2, value=f"{icon}  {title}")
        c.fill = F(bg_color)
        c.font = Ft(bold=True, size=12)
        c.alignment = A()
        c.border = B(color=BLACK)
        for j, line in enumerate(lines):
            r = row + 1 + j
            ws.row_dimensions[r].height = 20
            mc(ws, r, 2, r, 5)
            cx = ws.cell(row=r, column=2, value=line)
            cx.fill = F(DARK_BG)
            cx.font = Ft(color=WHITE, size=10)
            cx.alignment = A(h="left")
            cx.border = B(color=bg_color)
        return row + len(lines) + 2

    next_row = 8
    next_row = info_card(next_row, "\U0001f7e2", "CLASSE A — Os Campeões de Valor",
        ["  •  80% do valor total em poucos itens",
         "  •  Gestão rigorosa, estoque de segurança alto",
         "  •  Fórmula: % acumulado ≤ 80%"], GREEN_DARK)

    next_row = info_card(next_row, "\U0001f7e1", "CLASSE B — Os Médios Necessários",
        ["  •  Entre 80% e 95% do valor acumulado",
         "  •  Gestão intermediária, revise trimestralmente",
         "  •  Fórmula: % acumulado de 80% a 95%"], GOLD_DARK)

    next_row = info_card(next_row, "\U0001f534", "CLASSE C — A Longa Cauda",
        ["  •  Os últimos 5% do valor acumulado",
         "  •  Muitos itens, pouco impacto financeiro",
         "  •  Fórmula: % acumulado acima de 95%"], RED_DARK)

    next_row = info_card(next_row, "⚙️", "Como Usar Este Dashboard",
        ["  1️⃣  Vá para \U0001f4ca CURVA ABC e edite Qtd. e Custo Unit.",
         "  2️⃣  O gráfico com as faixas coloridas atualiza automaticamente",
         "  3️⃣  Para colorir via Excel Web: use o script na aba ⚙️ SCRIPTS",
         "  4️⃣  Proteção ativa — senha: abc2026."], "3A3A3A")

    ws.row_dimensions[next_row].height = 8
    next_row += 1

    # Footer
    ws.row_dimensions[next_row].height = 22
    mc(ws, next_row, 2, next_row, 5)
    c = ws.cell(row=next_row, column=2,
                value="\U0001f1ef\U0001f1f2  Jah bless your spreadsheet!  |  One Love, One Excel!  |  2026  \U0001f33f")
    c.fill = F(BLACK)
    c.font = Ft(bold=True, color=GREEN, size=10)
    c.alignment = A()

    ws.sheet_properties.tabColor = GREEN_DARK
    return ws


# ═══════════════════════════════════════════════════════════════════════════════
#  ABA 2 — 📊 CURVA ABC
# ═══════════════════════════════════════════════════════════════════════════════
def create_abc_sheet(wb):
    ws = wb.create_sheet(title="\U0001f4ca CURVA ABC")

    DS = 10   # data start row
    sorted_data, total_vt = classify(ITEMS)
    n  = len(sorted_data)
    DE = DS + n - 1

    # ── Larguras ──
    col_w = {
        "A":  7,  "B": 10,  "C": 30,  "D": 10,
        "E": 14,  "F": 16,  "G": 12,  "H": 12,
        "I":  2,  "J": 11,  "K": 16,  "L": 10,
        "M": 10,  "N":  2,
        # Células de limite (editáveis)
        "O": 12,  "P": 12,
        # Dados do gráfico (ocultos)
        "Q":  2,  "R": 28,  "S": 14,  "T": 14,
        "U": 10,  "V": 10,  "W": 10,
    }
    for c, w in col_w.items():
        ws.column_dimensions[c].width = w

    # ── Alturas ──
    ws.row_dimensions[1].height = 52
    for r in range(2, 8):
        ws.row_dimensions[r].height = 24
    ws.row_dimensions[8].height = 10
    ws.row_dimensions[9].height = 26
    for r in range(DS, DE + 3):
        ws.row_dimensions[r].height = 18

    # ── Linha 1: Título ──
    mc(ws, 1, 1, 1, 8)
    t = ws.cell(row=1, column=1,
                value="\U0001f1ef\U0001f1f2  CURVA ABC REGGAE  —  One Love, One Excel! ✌️")
    t.fill = F(BLACK)
    t.font = Font(bold=True, color=GREEN, size=22, name="Calibri")
    t.alignment = A()

    mc(ws, 1, 10, 1, 13)
    leg = ws.cell(row=1, column=10,
                  value="\U0001f7e2 A=até 80%  |\U0001f7e1 B=80-95%  |\U0001f534 C=95-100%")
    leg.fill = F(BLACK)
    leg.font = Ft(color=GOLD, size=10)
    leg.alignment = A()
    for c in range(9, 14):
        ws.cell(row=1, column=c).fill = F([GREEN,GOLD,RED,GREEN,BLACK][(c-9)%5])

    # ── KPI ──────────────────────────────────────────────────────────────────
    for r in range(2, 8):
        for c in range(1, 9):
            ws.cell(row=r, column=c).fill = F(DARK_BG)

    def kpi(lrow, vrow, col, label, formula, bg, lfg=WHITE, span=2, fmt=None):
        mc(ws, lrow, col, lrow, col+span-1)
        lc = ws.cell(row=lrow, column=col, value=label)
        lc.fill = F(DARK_BG); lc.font = Ft(bold=True, color=lfg, size=9)
        lc.alignment = A()
        mc(ws, vrow, col, vrow, col+span-1)
        vc = ws.cell(row=vrow, column=col, value=formula)
        vc.fill = F(bg); vc.font = Ft(bold=True, size=14)
        vc.alignment = A()
        if fmt: vc.number_format = fmt

    kpi(2,3, 1,"V.T. TOTAL",    f"=SUM($F${DS}:$F${DE})",      GREEN_DARK, GREEN, fmt='R$ #,##0.00')
    kpi(2,3, 3,"\U0001f7e2 V.T. CLASSE A",
              f'=SUMIF($A${DS}:$A${DE},"A",$F${DS}:$F${DE})',   GREEN,      WHITE, fmt='R$ #,##0.00')
    kpi(2,3, 5,"% ITENS CLASSE A",
              f'=IFERROR(COUNTIF($A${DS}:$A${DE},"A")/COUNTA($B${DS}:$B${DE}),0)',
              GOLD_DARK, GOLD, fmt='0.0%')
    kpi(2,3, 7,"\U0001f534 MAIOR V.T.", f"=MAX($F${DS}:$F${DE})", RED_DARK, RED_LIGHT, fmt='R$ #,##0.00')

    kpi(4,5, 1,"MENOR V.T.",    f"=MIN($F${DS}:$F${DE})",       GREEN_DARK, GREEN,     fmt='R$ #,##0.00')
    kpi(4,5, 3,"\U0001f7e1 V.T. CLASSE B",
              f'=SUMIF($A${DS}:$A${DE},"B",$F${DS}:$F${DE})',   GOLD_DARK,  BLACK,     fmt='R$ #,##0.00')
    kpi(4,5, 5,"STATUS",
              f'=IF(IFERROR(MAX($H${DS}:$H${DE}),0)>1.0001,"⚠️ VERIFICAR","✅ OK")',
              GREEN, WHITE)
    kpi(4,5, 7,"ITENS TOTAIS",  f"=COUNTA($B${DS}:$B${DE})",    "3A3A3A",   WHITE)

    # Linha 6: controles de limite
    mc(ws, 6, 1, 7, 2)
    ctrl_lbl = ws.cell(row=6, column=1, value="⚙️  LIMITES\nA / B")
    ctrl_lbl.fill = F(DARK_BG); ctrl_lbl.font = Ft(bold=True, color=GOLD, size=9)
    ctrl_lbl.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Limite A (editável)
    la_lbl = ws.cell(row=6, column=3, value="Limite A (ex: 80%):")
    la_lbl.fill = F(GREEN_DARK); la_lbl.font = Ft(bold=True, size=9); la_lbl.alignment = A(h="left")
    mc(ws, 6, 3, 6, 4)

    la_val = ws.cell(row=6, column=5, value=0.80)
    la_val.fill = F(GREEN); la_val.font = Ft(bold=True, color=WHITE, size=14)
    la_val.alignment = A(); la_val.number_format = '0%'
    la_val.protection = Protection(locked=False)  # editável

    # Limite B (editável)
    lb_lbl = ws.cell(row=7, column=3, value="Limite B (ex: 95%):")
    lb_lbl.fill = F(GOLD_DARK); lb_lbl.font = Ft(bold=True, color=BLACK, size=9); lb_lbl.alignment = A(h="left")
    mc(ws, 7, 3, 7, 4)

    lb_val = ws.cell(row=7, column=5, value=0.95)
    lb_val.fill = F(GOLD); lb_val.font = Ft(bold=True, color=BLACK, size=14)
    lb_val.alignment = A(); lb_val.number_format = '0%'
    lb_val.protection = Protection(locked=False)  # editável

    # Data validation nos limites
    dv_a = DataValidation(type="decimal", operator="between", formula1="0.5", formula2="0.89",
                          error="Limite A deve ser entre 50% e 89%",
                          errorTitle="❌ Valor Inválido", showErrorMessage=True)
    dv_b = DataValidation(type="decimal", operator="between", formula1="0.5", formula2="0.99",
                          error="Limite B deve ser entre 50% e 99%",
                          errorTitle="❌ Valor Inválido", showErrorMessage=True)
    ws.add_data_validation(dv_a); dv_a.add(ws["E6"])
    ws.add_data_validation(dv_b); dv_b.add(ws["E7"])

    outer_border(ws, 2, 1, 7, 8)

    # ── Painel de resumo (cols J-M) ──
    mc(ws, 2, 10, 2, 13)
    sh = ws.cell(row=2, column=10, value="RESUMO POR CLASSE")
    sh.fill = F(BLACK); sh.font = Ft(bold=True, color=GOLD, size=10); sh.alignment = A()
    for c, label in [(10,"Classe"),(11,"Total V.T."),(12,"% Total"),(13,"Qtd.")]:
        cx = ws.cell(row=3, column=c, value=label)
        cx.fill = F("2A2A2A"); cx.font = Ft(bold=True, size=9); cx.alignment = A()
        cx.border = B(color=GRAY_DARK)

    for cls, bg, r in [("A",GREEN,4),("B",GOLD_DARK,5),("C",RED_DARK,6)]:
        fg = BLACK if cls == "B" else WHITE
        for col, val, fmt in [
            (10, cls,  None),
            (11, f'=SUMIF($A${DS}:$A${DE},"{cls}",$F${DS}:$F${DE})', 'R$ #,##0'),
            (12, f'=IFERROR(K{r}/SUM($F${DS}:$F${DE}),0)', '0.0%'),
            (13, f'=COUNTIF($A${DS}:$A${DE},"{cls}")', None),
        ]:
            cx = ws.cell(row=r, column=col, value=val)
            cx.fill = F(bg); cx.font = Ft(bold=True, color=fg, size=11 if col==10 else 10)
            cx.alignment = A(); cx.border = B(color=BLACK)
            if fmt: cx.number_format = fmt
        ws.row_dimensions[r].height = 22

    mc(ws, 7, 10, 7, 13)
    ok = ws.cell(row=7, column=10,
                 value=f'=IF(IFERROR(MAX($H${DS}:$H${DE}),0)>1.0001,"⚠️ VERIFICAR","✅ Tudo consistente")')
    ok.fill = F(DARK_BG); ok.font = Ft(color=GREEN, size=9); ok.alignment = A()

    # ── Instrução ──
    mc(ws, 8, 1, 8, 13)
    ins = ws.cell(row=8, column=1,
                  value="  \U0001f33f  Edite somente Qtd. (col D) e Custo Unit. (col E)  "
                        "| Limites A/B nas células E6 e E7  "
                        "| Gráfico atualiza automaticamente  \U0001f3b5")
    ins.fill = F(GREEN_DARK); ins.font = Ft(italic=True, color="C8F5DB", size=8)
    ins.alignment = A(h="left")
    ws.row_dimensions[8].height = 14

    # ── Cabeçalho de dados (row 9) ──
    headers = [
        (1,"ABC",BLACK,GREEN),(2,"SKU",BLACK,WHITE),(3,"Descrição",BLACK,WHITE),
        (4,"Qtd.",BLACK,GOLD),(5,"Custo Unit.(R$)",BLACK,GOLD),
        (6,"V.T. (R$)",BLACK,GREEN),(7,"% Individual",BLACK,WHITE),(8,"% Acumulado",BLACK,WHITE),
    ]
    for col, label, bg, fg in headers:
        cx = ws.cell(row=9, column=col, value=label)
        cx.fill = F(bg); cx.font = Ft(bold=True, color=fg, size=10)
        cx.alignment = A(); cx.border = B(color=GREEN_DARK)

    # Cabeçalhos auxiliares do gráfico (cols R-W, row 9)
    for col, label in [(18,"Label"),(19,"% Ind."),(20,"% Acum."),(21,"Zona A"),(22,"Zona B"),(23,"Zona C")]:
        cx = ws.cell(row=9, column=col, value=label)
        cx.fill = F(DARK_BG); cx.font = Ft(bold=True, color=GRAY_DARK, size=8); cx.alignment = A()

    unlock = Protection(locked=False)
    abc_lookup = {(d["sku"], d["desc"]): d["abc"] for d in sorted_data}

    # ── Linhas de dados ──
    for i, (sku, desc, qty, cost) in enumerate(ITEMS):
        row = DS + i
        abc_class = abc_lookup.get((sku, desc), "C")

        if abc_class == "A":
            rf = GREEN_LIGHT; abc_bg = GREEN
        elif abc_class == "B":
            rf = GOLD_LIGHT;  abc_bg = GOLD_DARK
        else:
            rf = RED_LIGHT;   abc_bg = RED_DARK

        fg_abc = BLACK if abc_class == "B" else WHITE

        # Col A: fórmula ABC (referencia E6 e E7 como limites dinâmicos)
        cx = ws.cell(row=row, column=1,
                     value=f'=IF($H{row}<=$E$6,"A",IF($H{row}<=$E$7,"B","C"))')
        cx.fill = F(abc_bg); cx.font = Ft(bold=True, color=fg_abc, size=11)
        cx.alignment = A(); cx.border = B(color=BLACK)

        # Col B: SKU
        cx = ws.cell(row=row, column=2, value=sku)
        cx.fill = F(rf); cx.font = Ft(color=BLACK, size=10); cx.alignment = A(); cx.border = B(color=GRAY_DARK)

        # Col C: Descrição
        cx = ws.cell(row=row, column=3, value=desc)
        cx.fill = F(rf); cx.font = Ft(color=BLACK, size=10); cx.alignment = A(h="left"); cx.border = B(color=GRAY_DARK)

        # Col D: Qtd (editável)
        cx = ws.cell(row=row, column=4, value=qty)
        cx.fill = F(WHITE); cx.font = Ft(bold=True, color=BLACK, size=10)
        cx.alignment = A(); cx.number_format = '#,##0'; cx.protection = unlock
        cx.border = Border(
            left=Side(style="medium", color=GREEN_DARK), right=Side(style="medium", color=GREEN_DARK),
            top=Side(style="thin", color=GRAY_MID),     bottom=Side(style="thin", color=GRAY_MID))

        # Col E: Custo Unit (editável)
        cx = ws.cell(row=row, column=5, value=cost)
        cx.fill = F(WHITE); cx.font = Ft(bold=True, color=BLACK, size=10)
        cx.alignment = A(h="right"); cx.number_format = 'R$ #,##0.00'; cx.protection = unlock
        cx.border = Border(
            left=Side(style="medium", color=GREEN_DARK), right=Side(style="medium", color=GREEN_DARK),
            top=Side(style="thin", color=GRAY_MID),     bottom=Side(style="thin", color=GRAY_MID))

        # Col F: VT
        cx = ws.cell(row=row, column=6, value=f'=IFERROR(D{row}*E{row},0)')
        cx.fill = F(rf); cx.font = Ft(bold=True, color=BLACK, size=10)
        cx.alignment = A(h="right"); cx.border = B(color=GRAY_DARK); cx.number_format = 'R$ #,##0.00'

        # Col G: % Individual
        cx = ws.cell(row=row, column=7,
                     value=f'=IFERROR(F{row}/SUM($F${DS}:$F${DE}),0)')
        cx.fill = F(rf); cx.font = Ft(color=BLACK, size=10)
        cx.alignment = A(); cx.border = B(color=GRAY_DARK); cx.number_format = '0.00%'

        # Col H: % Acumulado
        cx = ws.cell(row=row, column=8,
                     value=f'=IFERROR(SUMIF($F${DS}:$F${DE},">="&F{row},$F${DS}:$F${DE})/SUM($F${DS}:$F${DE}),0)')
        cx.fill = F(rf); cx.font = Ft(color=BLACK, size=10)
        cx.alignment = A(); cx.border = B(color=GRAY_DARK); cx.number_format = '0.00%'

        # ── Dados estáticos para o gráfico (pré-calculados em Python) ──
        d = sorted_data[i]  # sorted_data está na mesma ordem que ITEMS após sort? NÃO.
        # Preciso encontrar o item correto em sorted_data

    # Rewrite chart data correctly — use sorted_data order
    for i, d in enumerate(sorted_data):
        row = DS + i
        lbl = f"{d['sku']} ({d['abc']})"

        # Col R(18): label
        cx = ws.cell(row=row, column=18, value=lbl)
        cx.fill = F(DARK_BG); cx.font = Ft(color=GRAY_DARK, size=8); cx.alignment = A(h="left")

        # Col S(19): % individual
        cx = ws.cell(row=row, column=19, value=round(d["pct_ind"], 4))
        cx.fill = F(DARK_BG); cx.font = Ft(color=GRAY_DARK, size=8)
        cx.number_format = '0.00%'

        # Col T(20): % acumulado
        cx = ws.cell(row=row, column=20, value=round(d["pct_acc"], 4))
        cx.fill = F(DARK_BG); cx.font = Ft(color=GRAY_DARK, size=8)
        cx.number_format = '0.00%'

        # Cols U/V/W (21/22/23): zona A / zona B / zona C
        # Cada item tem exatamente 1.0 em sua zona e 0.0 nas demais
        cx = ws.cell(row=row, column=21, value=1.0 if d["abc"] == "A" else 0.0)
        cx.fill = F(GREEN_DARK); cx.font = Ft(color=GRAY_DARK, size=8)

        cx = ws.cell(row=row, column=22, value=1.0 if d["abc"] == "B" else 0.0)
        cx.fill = F(GOLD_DARK); cx.font = Ft(color=GRAY_DARK, size=8)

        cx = ws.cell(row=row, column=23, value=1.0 if d["abc"] == "C" else 0.0)
        cx.fill = F(RED_DARK); cx.font = Ft(color=GRAY_DARK, size=8)

    # ── Linha de total ──
    tr = DE + 1
    ws.row_dimensions[tr].height = 22
    mc(ws, tr, 1, tr, 3)
    cx = ws.cell(row=tr, column=1, value="TOTAL GERAL")
    cx.fill = F(BLACK); cx.font = Ft(bold=True, color=GOLD, size=11); cx.alignment = A()

    cx = ws.cell(row=tr, column=4, value=f'=SUM(D{DS}:D{DE})')
    cx.fill = F(BLACK); cx.font = Ft(bold=True, size=11); cx.alignment = A(); cx.number_format = '#,##0'
    ws.cell(row=tr, column=5).fill = F(BLACK)
    cx = ws.cell(row=tr, column=6, value=f'=SUM(F{DS}:F{DE})')
    cx.fill = F(GREEN); cx.font = Ft(bold=True, size=11); cx.alignment = A(h="right")
    cx.number_format = 'R$ #,##0.00'
    for col in [7, 8]:
        cx = ws.cell(row=tr, column=col, value=1.0)
        cx.fill = F(BLACK); cx.font = Ft(bold=True, size=11); cx.alignment = A(); cx.number_format = '0%'
    for col in range(1, 9):
        ws.cell(row=tr, column=col).border = Border(
            top=Side(style="medium",color=GREEN), bottom=Side(style="medium",color=GREEN),
            left=Side(style="thin",color=GREEN),  right=Side(style="thin",color=GREEN))

    # ── GRÁFICO — Curva ABC com faixas coloridas ────────────────────────────
    # AreaChart (stacked) para as bandas de fundo: verde(A) | dourado(B) | vermelho(C)
    # + BarChart para % individual (barras azuis)
    # + LineChart para % acumulado (linha vermelha)

    chart_row = tr + 3

    # ─ AreaChart: bandas de fundo ─
    area = AreaChart()
    area.grouping = "stacked"
    area.title    = "\U0001f1ef\U0001f1f2 CURVA ABC REGGAE — One Love, One Excel!"
    area.y_axis.title = "% do Total"
    area.x_axis.title = "Itens (ordenados por V.T. decrescente)"
    area.style    = 10
    area.width    = 30
    area.height   = 16

    # Série Zona A — verde
    ser_za = Series(Reference(ws, min_col=21, min_row=DS, max_row=DE), title="Zona A \U0001f7e2")
    ser_za.graphicalProperties.solidFill = GREEN_CHART
    area.series.append(ser_za)

    # Série Zona B — dourado
    ser_zb = Series(Reference(ws, min_col=22, min_row=DS, max_row=DE), title="Zona B \U0001f7e1")
    ser_zb.graphicalProperties.solidFill = GOLD_CHART
    area.series.append(ser_zb)

    # Série Zona C — vermelho/laranja (como na imagem de referência)
    ser_zc = Series(Reference(ws, min_col=23, min_row=DS, max_row=DE), title="Zona C \U0001f534")
    ser_zc.graphicalProperties.solidFill = RED_CHART
    area.series.append(ser_zc)

    area.set_categories(Reference(ws, min_col=18, min_row=DS, max_row=DE))
    area.y_axis.scaling.min = 0
    area.y_axis.scaling.max = 1
    area.y_axis.numFmt = '0%'

    # ─ BarChart: % individual ─
    bar = BarChart()
    bar.type = "col"
    ser_bar = Series(Reference(ws, min_col=19, min_row=DS, max_row=DE), title="% Individual \U0001f4ca")
    ser_bar.graphicalProperties.solidFill = BLUE_BAR
    ser_bar.graphicalProperties.line.solidFill = "1A4E8A"
    bar.series.append(ser_bar)
    bar.set_categories(Reference(ws, min_col=18, min_row=DS, max_row=DE))

    # ─ LineChart: % acumulado (curva de Pareto) ─
    line = LineChart()
    ser_line = Series(Reference(ws, min_col=20, min_row=DS, max_row=DE), title="% Acumulado \U0001f4c8")
    ser_line.graphicalProperties.line.solidFill = RED_DARK
    ser_line.graphicalProperties.line.width      = 28000  # linha grossa
    ser_line.smooth = True  # curva suave como na imagem
    line.series.append(ser_line)

    # ─ Combina os três chart types ─
    area += bar
    area += line

    ws.add_chart(area, f"A{chart_row}")

    # ── Formatação Condicional ──
    rng = f"A{DS}:H{DE}"
    for cls, clr in [("A", GREEN_LIGHT), ("B", GOLD_LIGHT), ("C", RED_LIGHT)]:
        ws.conditional_formatting.add(rng, FormulaRule(
            formula=[f'$A{DS}="{cls}"'],
            fill=PatternFill(fill_type="solid", fgColor=clr),
            font=Font(color=BLACK)))

    # ── Proteção ──
    ws.protection.sheet    = True
    ws.protection.password = "abc2026."
    ws.protection.enable()
    ws.protection.selectLockedCells   = False
    ws.protection.selectUnlockedCells = False
    for r in range(DS, DE + 1):
        for col in [4, 5]:  # Qtd. e Custo Unit. — editáveis
            ws.cell(row=r, column=col).protection = unlock
    ws.cell(row=6, column=5).protection = unlock  # Limite A
    ws.cell(row=7, column=5).protection = unlock  # Limite B

    ws.freeze_panes = f"A{DS}"
    ws.sheet_properties.tabColor = GREEN

    return ws


# ═══════════════════════════════════════════════════════════════════════════════
#  ABA 3 — ⚙️ SCRIPTS
# ═══════════════════════════════════════════════════════════════════════════════

OFFICE_SCRIPT_TS = '''\
/**
 * ================================================================
 *  🇯🇲 ONE LOVE, ONE EXCEL! — Curva ABC Reggae Automator
 *  Office Script para Excel Web (TypeScript)
 * ================================================================
 *  COMO USAR:
 *  1. Abra o arquivo no Excel Web (office.com)
 *  2. Clique em "Automatizar" → "Novo Script"
 *  3. Apague o código padrão e cole ESTE script inteiro
 *  4. Clique em "Executar" 🎵
 *
 *  Jah bless your spreadsheet! 🌿
 * ================================================================
 */
function main(workbook: ExcelScript.Workbook) {

  // 🎵 CONFIGURAÇÕES — ajuste conforme sua planilha
  const NOME_ABA      = "📊 CURVA ABC";  // Nome exato da aba
  const LINHA_INICIO  = 10;              // Linha onde começam os dados (1-based)
  const COL_ABC       = 1;              // Coluna da classificação ABC (A=1)
  const COL_INICIO    = 1;              // 1ª coluna dos dados
  const COL_FIM       = 8;             // Última coluna dos dados
  const LIMITE_A      = 0.80;           // 80% para classe A
  const LIMITE_B      = 0.95;           // 95% para classe B

  // 🦁 PALETA REGGAE — Jah Colors
  const COR_A_LINHA  = "#C8F5DB";  // Verde claro — fundo linha A
  const COR_A_BADGE  = "#006B28";  // Verde escuro — célula ABC
  const COR_B_LINHA  = "#FFF5B0";  // Dourado claro — fundo linha B
  const COR_B_BADGE  = "#C8A400";  // Dourado escuro — célula ABC
  const COR_C_LINHA  = "#FFD0D0";  // Vermelho claro — fundo linha C
  const COR_C_BADGE  = "#990000";  // Vermelho escuro — célula ABC
  const BRANCO       = "#FFFFFF";
  const PRETO        = "#000000";

  // 🌿 Pega a planilha
  const sheet = workbook.getWorksheet(NOME_ABA);
  if (!sheet) {
    console.log(`❌ Aba "${NOME_ABA}" não encontrada! Verifique o nome exato.`);
    return;
  }

  // 📊 Descobre a última linha com dados
  const usedRange  = sheet.getUsedRange();
  if (!usedRange) { console.log("❌ Planilha vazia, mon!"); return; }
  const ultimaLinha = usedRange.getLastRow().getRowIndex() + 1; // 1-based

  console.log(`🎵 Iniciando coloração reggae... Verificando linhas ${LINHA_INICIO} até ${ultimaLinha}`);

  let countA = 0, countB = 0, countC = 0, skip = 0;

  // ☀️ Loop linha por linha — Bob Marley style: um passo de cada vez
  for (let row = LINHA_INICIO; row <= ultimaLinha; row++) {
    const cellABC  = sheet.getCell(row - 1, COL_ABC - 1);         // 0-based
    const valorABC = String(cellABC.getValue() ?? "").trim().toUpperCase();

    if (valorABC !== "A" && valorABC !== "B" && valorABC !== "C") {
      skip++;
      continue;
    }

    // 🌊 Range da linha inteira (colunas de dados)
    const rowRange = sheet.getRangeByIndexes(
      row - 1, COL_INICIO - 1, 1, COL_FIM
    );

    let corLinha: string;
    let corBadge: string;

    if (valorABC === "A") {
      corLinha = COR_A_LINHA; corBadge = COR_A_BADGE; countA++;
    } else if (valorABC === "B") {
      corLinha = COR_B_LINHA; corBadge = COR_B_BADGE; countB++;
    } else {
      corLinha = COR_C_LINHA; corBadge = COR_C_BADGE; countC++;
    }

    // 🥁 Pinta a linha toda
    rowRange.getFormat().getFill().setColor(corLinha);
    rowRange.getFormat().getFont().setColor(PRETO);

    // 🦁 Destaca o badge ABC com cor sólida e negrito
    cellABC.getFormat().getFill().setColor(corBadge);
    cellABC.getFormat().getFont().setColor(BRANCO);
    cellABC.getFormat().getFont().setBold(true);
    cellABC.getFormat().getFont().setSize(12);
    cellABC.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  }

  // 🇯🇲 Relatório final — One Love!
  const total = countA + countB + countC;
  console.log("\\n╔══════════════════════════════════════════╗");
  console.log("║   🇯🇲 ONE LOVE, ONE EXCEL! 🇯🇲           ║");
  console.log("╠══════════════════════════════════════════╣");
  console.log(`║  🟢 Classe A: ${String(countA).padEnd(3)} itens — campeões de valor  ║`);
  console.log(`║  🟡 Classe B: ${String(countB).padEnd(3)} itens — médios necessários ║`);
  console.log(`║  🔴 Classe C: ${String(countC).padEnd(3)} itens — a longa cauda      ║`);
  console.log(`║  📊 Total   : ${String(total).padEnd(3)} itens classificados       ║`);
  console.log("╚══════════════════════════════════════════╝");
  console.log("✅ Curva ABC colorida com reggae! Jah bless your data! 🌿");
}
'''

def create_scripts_sheet(wb):
    ws = wb.create_sheet(title="⚙️ SCRIPTS")

    for c, w in {"A":4,"B":6,"C":100,"D":6}.items():
        ws.column_dimensions[c].width = w

    # Background
    for r in range(1, 120):
        ws.row_dimensions[r].height = 16
        for c in range(1, 5):
            ws.cell(row=r, column=c).fill = F(DARK_BG)

    # Faixa reggae
    ws.row_dimensions[1].height = 8
    for col, clr in [(1,GREEN),(2,GOLD),(3,RED),(4,BLACK)]:
        ws.cell(row=1, column=col).fill = F(clr)

    # Título
    ws.row_dimensions[2].height = 56
    ws.cell(row=2, column=2, value="⚙️  OFFICE SCRIPTS  —  TypeScript para Excel Web").fill = F(BLACK)
    c = ws.cell(row=2, column=2)
    c.value = "⚙️  OFFICE SCRIPTS — One Love, One Excel! \U0001f1ef\U0001f1f2"
    c.fill = F(BLACK); c.font = Font(bold=True, color=GREEN, size=22, name="Calibri"); c.alignment = A()

    # Instruções
    def instrucao(row, text, bg, fg=WHITE, bold=False, size=10):
        ws.row_dimensions[row].height = 18
        cx = ws.cell(row=row, column=2, value=text)
        cx.fill = F(bg); cx.font = Ft(bold=bold, color=fg, size=size); cx.alignment = A(h="left")
        cx.border = B(color=bg)
        return cx

    instrucao(3,  "\U0001f4cb  COMO USAR O SCRIPT ABAIXO:", GREEN_DARK, WHITE, True, 11)
    instrucao(4,  "  1️⃣  Abra seu arquivo .xlsx no Excel Web (office.com ou Microsoft 365)",
              "2A2A2A")
    instrucao(5,  "  2️⃣  Clique na aba \"Automatizar\" no menu superior do Excel Web",
              "2A2A2A")
    instrucao(6,  "  3️⃣  Clique em \"Novo Script\" e apague o código padrão",
              "2A2A2A")
    instrucao(7,  "  4️⃣  Copie TODO o código abaixo e cole no editor de scripts",
              "2A2A2A")
    instrucao(8,  "  5️⃣  Ajuste as constantes no topo (NOME_ABA, LINHA_INICIO, etc.) se necessário",
              "2A2A2A")
    instrucao(9,  "  6️⃣  Clique em \"Executar\" \U0001f3b5  —  As cores reggae serão aplicadas!",
              "2A2A2A")
    instrucao(10, "  ✅  O script colore as linhas: Verde (A) | Dourado (B) | Vermelho (C)",
              GREEN_DARK, WHITE, True)
    instrucao(11, "  ⚠️  VBA não roda no Excel Web — use este Office Script (TypeScript) no lugar",
              GOLD_DARK, BLACK, True)

    ws.row_dimensions[12].height = 10

    # Cabeçalho do código
    instrucao(13, "  \U0001f4dc  CÓDIGO TYPESCRIPT — Copie tudo abaixo:", BLACK, GOLD, True, 11)
    ws.row_dimensions[13].height = 22

    # Código TypeScript — uma linha por célula
    code_start = 14
    lines = OFFICE_SCRIPT_TS.split('\n')
    for i, line in enumerate(lines):
        r = code_start + i
        ws.row_dimensions[r].height = 15
        cx = ws.cell(row=r, column=2, value=line if line else " ")
        cx.fill = F("0D0D0D")
        # Coloração sintática simples
        stripped = line.strip()
        if stripped.startswith("//") or stripped.startswith("*") or stripped.startswith("/**") or stripped.startswith("*/"):
            cx.font = Font(color="6A9955", size=9, name="Courier New", italic=True)  # verde = comentário
        elif any(kw in line for kw in ["function","const","let","if","else","for","return","string","number"]):
            cx.font = Font(color="569CD6", size=9, name="Courier New", bold=False)  # azul = keyword
        elif stripped.startswith("console."):
            cx.font = Font(color="FFD700", size=9, name="Courier New")  # dourado = console
        else:
            cx.font = Font(color="D4D4D4", size=9, name="Courier New")  # branco/cinza = código normal
        cx.alignment = Alignment(horizontal="left", vertical="center")
        cx.border = Border(left=Side(style="thin", color="1A1A2E"), right=Side(style="thin", color="1A1A2E"),
                           top=Side(style="hair", color="1A1A2E"), bottom=Side(style="hair", color="1A1A2E"))

    # Pós código
    after_row = code_start + len(lines) + 1
    ws.row_dimensions[after_row].height = 10
    instrucao(after_row + 1,
              "\U0001f1ef\U0001f1f2  Jah bless your automation! One Love, One Excel!  \U0001f33f",
              BLACK, GREEN, True, 10)

    ws.sheet_properties.tabColor = RED_DARK
    return ws


# ═══════════════════════════════════════════════════════════════════════════════
#  MAIN
# ═══════════════════════════════════════════════════════════════════════════════
def main():
    OUTPUT = "/home/user/Zyth/curva_abc_reggae.xlsx"

    wb = Workbook()
    wb.remove(wb.active)

    print("  🇯🇲 Criando aba INICIO...")
    create_inicio(wb)

    print("  📊 Criando aba CURVA ABC (com gráfico de faixas)...")
    create_abc_sheet(wb)

    print("  ⚙️ Criando aba SCRIPTS (TypeScript Office Scripts)...")
    create_scripts_sheet(wb)

    wb.properties.title   = "Curva ABC Reggae — One Love, One Excel!"
    wb.properties.subject = "Dashboard ABC — Reggae Colors — Office Scripts"
    wb.properties.creator = "🇯🇲 Curva ABC Reggae"

    wb.save(OUTPUT)

    import os
    size = os.path.getsize(OUTPUT)
    print(f"\n✅  Salvo: {OUTPUT}")
    print(f"📦  Tamanho: {size/1024:.1f} KB")
    print("🇯🇲  One Love, One Excel!")


if __name__ == "__main__":
    main()
