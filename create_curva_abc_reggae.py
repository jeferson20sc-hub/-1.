#!/usr/bin/env python3
"""
🇯🇲 CURVA ABC REGGAE — One Love, One Excel!
3 abas de dados: Primeira (10 itens) | 01 (20 itens) | 02 (30 itens)
Gráfico: bandas coloridas (A=verde | B=dourado | C=vermelho) + barras + curva Pareto
Nome do produto em cada coluna do gráfico
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.chart import AreaChart, BarChart, LineChart, Reference, Series
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.styles.protection import Protection

# ── Paleta Reggae ──────────────────────────────────────────────────────────────
G_DARK   = "006B28";  G      = "009B3A";  G_LITE  = "C8F5DB";  G_CHART = "ADEBAD"
Y_DARK   = "C8A400";  Y      = "FFD700";  Y_LITE  = "FFF5B0";  Y_CHART = "FFE87C"
R_DARK   = "990000";  R      = "CC0000";  R_LITE  = "FFD0D0";  R_CHART = "FFAD99"
BLK      = "000000";  WHT    = "FFFFFF";  DARK    = "1A1A1A"
BLUE_BAR = "2E75B6";  GDARK2 = "B0B0B0";  GMID    = "E8E8E8"

# ── Helpers ────────────────────────────────────────────────────────────────────
def F(c): return PatternFill(fill_type="solid", fgColor=c)
def Ft(bold=False, color=WHT, size=11, italic=False, name="Calibri"):
    return Font(bold=bold, color=color, size=size, italic=italic, name=name)
def B(style="thin", color=BLK):
    s = Side(style=style, color=color); return Border(left=s, right=s, top=s, bottom=s)
def Al(h="center", wrap=False): return Alignment(horizontal=h, vertical="center", wrap_text=wrap)
def mc(ws, r1, c1, r2, c2): ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)

def thick_border(ws, r1, c1, r2, c2, color=G, thick="medium"):
    t = Side(style=thick, color=color); n = Side(style="thin", color=color)
    for row in ws.iter_rows(r1, r2, c1, c2):
        for cx in row:
            r, cc = cx.row, cx.column
            cx.border = Border(
                top   =t if r==r1 else n, bottom=t if r==r2 else n,
                left  =t if cc==c1 else n, right =t if cc==c2 else n)

# ── Dados das 3 Abas ───────────────────────────────────────────────────────────
SHEETS = [
    {
        "name": "🟢 Primeira",
        "title": "CURVA ABC REGGAE — PRIMEIRA",
        "tab_color": G_DARK,
        "items": [
            # (SKU,  Descrição,                  Cod, Qtd,  Custo)
            ("SKU01","Servidor Backup",            3,   2,  25000),
            ("SKU02",'Monitor 27" 4K',             5,  10,   3000),
            ("SKU03","Notebook Administrativo",     9,   6,   4500),
            ("SKU04","Nobreak Profissional",         1,   5,   2000),
            ("SKU05","Switch Gerenciavel",           7,   4,   1500),
            ("SKU06","Mouse sem fio",                4,  40,     50),
            ("SKU07","Patch Cord 1m",                8, 100,     15),
            ("SKU08","Teclado USB",                  6,  80,     35),
            ("SKU09","Cabo Rede Cat6 (m)",            2, 200,      8),
            ("SKU10","Mousepad Slim",               10, 120,     12),
        ],
    },
    {
        "name": "🟡 01",
        "title": "CURVA ABC REGGAE — ABA 01",
        "tab_color": Y_DARK,
        "items": [
            ("SKU01","Notebook Executivo i7",       10,  35,  6500),
            ("SKU02","Servidor de Dados Pro",        2,  12, 15000),
            ("SKU03","Impressora Multifuncional",   17,  55,  1100),
            ("SKU04","Toner Impressora Laser",      12,  68,   680),
            ("SKU05","Memoria RAM 16GB",            19,  45,   380),
            ("SKU06",'Monitor LED 24"',              4,  20,   850),
            ("SKU07","Webcam Full HD",              14,  82,   180),
            ("SKU08","Roteador Wi-Fi 6",            11,  15,   650),
            ("SKU09","SSD 480GB Sata",              15,  55,   190),
            ("SKU10","Nobreak 1500VA",               6,   8,  1200),
            ("SKU11","HD Externo 2TB",              20,  18,   420),
            ("SKU12","Teclado Mecanico RGB",         5,  30,   250),
            ("SKU13","Pen Drive 64GB",               9, 160,    45),
            ("SKU14","Headset com Microfone",       13,  64,   110),
            ("SKU15","Pacote de Papel A4",           7, 237,    28),
            ("SKU16","Filtro de Linha 5 Tom.",      18, 110,    40),
            ("SKU17","Cabo HDMI 2m",                 1, 150,    25),
            ("SKU18","Adaptador USB-C",             16,  90,    35),
            ("SKU19","Mouse Optico Simples",          3, 120,    15),
            ("SKU20","Patch Cord RJ45 1m",           8, 200,    12),
        ],
    },
    {
        "name": "🔴 02",
        "title": "CURVA ABC REGGAE — ABA 02",
        "tab_color": R_DARK,
        "items": [
            ("SKU01","Servidor Rack Dell Pro",       2,  15, 22500),
            ("SKU02","Workstation Grafica",         17,  19, 12000),
            ("SKU03","Notebook i7 32GB RAM",        10,  28,  7200),
            ("SKU04","Switch Gerenciavel 48p",       8,  10,  4800),
            ("SKU05","Placa de Video RTX",          22,   8,  5500),
            ("SKU06","Storage NAS 40TB",            21,   2, 18500),
            ("SKU07","Memoria RAM 16GB",            19,  45,   380),
            ("SKU08","Toner Impressora Laser",      12,  40,   320),
            ("SKU09",'Monitor LED 24"',              4,  15,   850),
            ("SKU10","Roteador Wi-Fi 6",            11,  25,   450),
            ("SKU11","SSD 480GB Sata",              15,  55,   190),
            ("SKU12","Nobreak 1500VA",               6,   8,  1200),
            ("SKU13","HD Externo 2TB",              20,  18,   420),
            ("SKU14","Teclado Mecanico RGB",         5,  30,   250),
            ("SKU15","Pen Drive 64GB",               9, 150,    45),
            ("SKU16","Headset com Microfone",       13,  60,   110),
            ("SKU17","Pacote de Papel A4",           7, 200,    28),
            ("SKU18","Teclado de Entrada",          29, 100,    45),
            ("SKU19","Filtro de Linha",             18, 110,    40),
            ("SKU20","Webcam Full HD",              14,  22,   180),
            ("SKU21","Pilhas AA (Pacote)",          23, 300,    12),
            ("SKU22","Adaptador USB-C",             16,  90,    35),
            ("SKU23","Suporte para Monitor",        28,  20,   140),
            ("SKU24","Cabo HDMI 2m",                 1,  85,    25),
            ("SKU25","Mousepad Simples",            24, 250,      8),
            ("SKU26","Mouse Optico Simples",          3, 120,    15),
            ("SKU27","Organizador de Cabos",        25, 180,    10),
            ("SKU28","Ar comprimido (Lata)",        27,  35,    45),
            ("SKU29","Pasta Termica",               26,  45,    35),
            ("SKU30","Conector RJ45 (Cento)",       30,  50,    30),
        ],
    },
]

# ── Classificação ABC ──────────────────────────────────────────────────────────
def classify(items, lim_a=0.80, lim_b=0.95):
    rows = [{"sku": s, "desc": d, "cod": c, "qty": q, "cost": cu, "vt": q*cu}
            for s, d, c, q, cu in items]
    rows.sort(key=lambda x: x["vt"], reverse=True)
    total = sum(r["vt"] for r in rows)
    cumsum = 0
    for r in rows:
        cumsum += r["vt"]
        r["pct_ind"] = r["vt"] / total if total else 0
        r["pct_acc"] = cumsum / total if total else 0
        r["abc"] = "A" if r["pct_acc"] <= lim_a else ("B" if r["pct_acc"] <= lim_b else "C")
    return rows, total

# ── Cria cada aba de dados ─────────────────────────────────────────────────────
def create_data_sheet(wb, cfg):
    ws   = wb.create_sheet(title=cfg["name"])
    raw  = cfg["items"]
    DS   = 12     # data start row
    n    = len(raw)
    DE   = DS + n - 1

    # Pré-calcula ABC em Python (para estilização e gráfico)
    sorted_data, total_vt = classify(raw)
    # Lookup desc -> abc para colorir as linhas na ordem original
    abc_lk = {(d["sku"], d["desc"]): d["abc"] for d in sorted_data}

    # ── Larguras das colunas ─────────────────────────────────────────────────
    widths = {
        "A": 7,  "B": 9,  "C": 8,  "D": 30, "E": 10,
        "F": 14, "G": 16, "H": 12, "I": 12,
        "J": 2,                                          # espaçador
        "K": 11, "L": 16, "M": 10, "N": 10,             # painel resumo
        "O": 2,                                          # espaçador
        # Dados do gráfico (ocultos visualmente)
        "P": 28, "Q": 13, "R": 13,                      # label | % ind | % acc
        "S": 10, "T": 10, "U": 10,                      # zona A | zona B | zona C
    }
    for c, w in widths.items():
        ws.column_dimensions[c].width = w

    # ── Alturas das linhas ───────────────────────────────────────────────────
    ws.row_dimensions[1].height = 52   # título
    ws.row_dimensions[2].height = 8    # faixa reggae
    ws.row_dimensions[3].height = 26   # kpi label 1
    ws.row_dimensions[4].height = 30   # kpi value 1
    ws.row_dimensions[5].height = 26   # kpi label 2
    ws.row_dimensions[6].height = 30   # kpi value 2
    ws.row_dimensions[7].height = 8    # faixa
    ws.row_dimensions[8].height = 22   # controles + sumário header
    ws.row_dimensions[9].height = 22   # sumário A
    ws.row_dimensions[10].height = 22  # sumário B
    ws.row_dimensions[11].height = 22  # cabeçalho dados
    for r in range(DS, DE + 2):
        ws.row_dimensions[r].height = 18

    # ── Row 1: Banner título ─────────────────────────────────────────────────
    mc(ws, 1, 1, 1, 9)
    t = ws.cell(row=1, column=1, value=f"🇯🇲  {cfg['title']}  —  One Love, One Excel! ✌️")
    t.fill = F(BLK); t.font = Font(bold=True, color=G, size=20, name="Calibri"); t.alignment = Al()

    # Faixa lateral colorida
    for ci, clr in enumerate([G, Y, R, G, BLK], start=10):
        ws.cell(row=1, column=ci).fill = F(clr)
    mc(ws, 1, 10, 1, 14)
    lg = ws.cell(row=1, column=10, value="🟢 A=até 80%  |  🟡 B=80-95%  |  🔴 C=95-100%")
    lg.fill = F(BLK); lg.font = Ft(color=Y, size=10); lg.alignment = Al()

    # ── Row 2: faixa reggae ──────────────────────────────────────────────────
    for ci, clr in enumerate([G, Y, R, BLK, G, Y, R, BLK, G], start=1):
        ws.cell(row=2, column=ci).fill = F(clr)

    # ── Rows 3-6: KPI cards ──────────────────────────────────────────────────
    for r in range(3, 7):
        for c in range(1, 10):
            ws.cell(row=r, column=c).fill = F(DARK)

    def kpi(lr, vr, c, label, formula, bg, span=2, fmt=None, lfg=WHT):
        mc(ws, lr, c, lr, c+span-1)
        lc = ws.cell(row=lr, column=c, value=label)
        lc.fill = F(DARK); lc.font = Ft(bold=True, color=lfg, size=9); lc.alignment = Al()
        mc(ws, vr, c, vr, c+span-1)
        vc = ws.cell(row=vr, column=c, value=formula)
        vc.fill = F(bg); vc.font = Ft(bold=True, size=14); vc.alignment = Al()
        if fmt: vc.number_format = fmt

    kpi(3,4, 1, "💰 V.T. TOTAL",       f"=SUM($G${DS}:$G${DE})",        G_DARK, fmt='R$ #,##0.00')
    kpi(3,4, 3, "🟢 V.T. CLASSE A",
               f'=SUMIF($A${DS}:$A${DE},"A",$G${DS}:$G${DE})',            G,     fmt='R$ #,##0.00')
    kpi(3,4, 5, "🟡 V.T. CLASSE B",
               f'=SUMIF($A${DS}:$A${DE},"B",$G${DS}:$G${DE})',            Y_DARK,fmt='R$ #,##0.00', lfg=BLK)
    kpi(3,4, 7, "🔴 MAIOR V.T.",        f"=MAX($G${DS}:$G${DE})",         R_DARK,fmt='R$ #,##0.00')

    kpi(5,6, 1, "📦 TOTAL ITENS",       f"=COUNTA($D${DS}:$D${DE})",      G_DARK)
    kpi(5,6, 3, "% ITENS CLASSE A",
               f'=IFERROR(COUNTIF($A${DS}:$A${DE},"A")/COUNTA($D${DS}:$D${DE}),0)',
               G, fmt='0.0%')
    kpi(5,6, 5, "🔴 V.T. CLASSE C",
               f'=SUMIF($A${DS}:$A${DE},"C",$G${DS}:$G${DE})',            R_DARK,fmt='R$ #,##0.00')
    kpi(5,6, 7, "✅ STATUS",
               f'=IF(IFERROR(MAX($I${DS}:$I${DE}),0)>1.001,"⚠️ VERIFICAR","✅ OK")',
               G_DARK)

    thick_border(ws, 3, 1, 6, 9)

    # ── Row 7: faixa fina ────────────────────────────────────────────────────
    for ci, clr in enumerate([R, Y, G, BLK, R, Y, G, BLK, R], start=1):
        ws.cell(row=7, column=ci).fill = F(clr)

    # ── Rows 8-10: Painel de resumo A/B/C (cols K-N) ─────────────────────────
    for ci, label in [(11,"Classe"),(12,"Total V.T."),(13,"% Total"),(14,"Qtd.")]:
        cx = ws.cell(row=8, column=ci, value=label)
        cx.fill = F(BLK); cx.font = Ft(bold=True, size=9); cx.alignment = Al()
        cx.border = B(color=GDARK2)

    for cls, bg, row in [("A", G, 9), ("B", Y_DARK, 10)]:
        fg = BLK if cls == "B" else WHT
        for ci, val, fmt in [
            (11, cls, None),
            (12, f'=SUMIF($A${DS}:$A${DE},"{cls}",$G${DS}:$G${DE})', 'R$ #,##0'),
            (13, f'=IFERROR(L{row}/SUM($G${DS}:$G${DE}),0)', '0.0%'),
            (14, f'=COUNTIF($A${DS}:$A${DE},"{cls}")', None),
        ]:
            cx = ws.cell(row=row, column=ci, value=val)
            cx.fill = F(bg); cx.font = Ft(bold=True, color=fg, size=10 if ci>11 else 14)
            cx.alignment = Al(); cx.border = B(color=BLK)
            if fmt: cx.number_format = fmt

    # Classe C em row 8 também (shift: 8=header, 9=A, 10=B, 11=header dados → Não tem lugar para C)
    # Vou colocar classe C inline abaixo da B:
    for ci, val, fmt in [
        (11, "C", None),
        (12, f'=SUMIF($A${DS}:$A${DE},"C",$G${DS}:$G${DE})', 'R$ #,##0'),
        (13, f'=IFERROR(L11/SUM($G${DS}:$G${DE}),0)', '0.0%'),
        (14, f'=COUNTIF($A${DS}:$A${DE},"C")', None),
    ]:
        cx = ws.cell(row=11, column=ci, value=val)
        cx.fill = F(R_DARK); cx.font = Ft(bold=True, color=WHT, size=10 if ci>11 else 14)
        cx.alignment = Al(); cx.border = B(color=BLK)
        if fmt: cx.number_format = fmt

    # ── Controles de limite (cols A-I, row 8-10) ─────────────────────────────
    # Instrução
    mc(ws, 8, 1, 8, 9)
    ins = ws.cell(row=8, column=1,
                  value="  🌿  Edite somente Qtd. (col F) e Custo Unit. (col G)  |  Limites: E9=Classe A (80%)  E10=Classe B (95%)")
    ins.fill = F(G_DARK); ins.font = Ft(italic=True, color="C8F5DB", size=8)
    ins.alignment = Al(h="left")

    # Células de limite
    ws.cell(row=9, column=1, value="Limite A:").fill  = F(G_DARK)
    ws.cell(row=9, column=1).font = Ft(bold=True, color=WHT, size=9)
    ws.cell(row=9, column=1).alignment = Al(h="right")

    lim_a = ws.cell(row=9, column=2, value=0.80)
    lim_a.fill = F(G); lim_a.font = Ft(bold=True, color=WHT, size=13)
    lim_a.alignment = Al(); lim_a.number_format = '0%'
    lim_a.protection = Protection(locked=False)

    ws.cell(row=10, column=1, value="Limite B:").fill = F(Y_DARK)
    ws.cell(row=10, column=1).font = Ft(bold=True, color=BLK, size=9)
    ws.cell(row=10, column=1).alignment = Al(h="right")

    lim_b = ws.cell(row=10, column=2, value=0.95)
    lim_b.fill = F(Y); lim_b.font = Ft(bold=True, color=BLK, size=13)
    lim_b.alignment = Al(); lim_b.number_format = '0%'
    lim_b.protection = Protection(locked=False)

    # Data validation
    dv_a = DataValidation(type="decimal", operator="between", formula1="0.5", formula2="0.89",
                          error="Limite A: entre 50% e 89%", errorTitle="❌ Valor Inválido",
                          showErrorMessage=True)
    dv_b = DataValidation(type="decimal", operator="between", formula1="0.5", formula2="0.99",
                          error="Limite B: entre 50% e 99%", errorTitle="❌ Valor Inválido",
                          showErrorMessage=True)
    ws.add_data_validation(dv_a); dv_a.add(ws.cell(row=9,  column=2))
    ws.add_data_validation(dv_b); dv_b.add(ws.cell(row=10, column=2))

    # ── Row 11: Cabeçalho das colunas de dados ────────────────────────────────
    headers = [
        (1,"ABC",BLK,G), (2,"SKU",BLK,WHT), (3,"Cod.",BLK,WHT),
        (4,"Descrição do Item",BLK,WHT), (5,"Qtd.",BLK,Y),
        (6,"Custo Unit.(R$)",BLK,Y), (7,"V.T. (R$)",BLK,G),
        (8,"% Individual",BLK,WHT), (9,"% Acumulado",BLK,WHT),
    ]
    for col, label, bg, fg in headers:
        cx = ws.cell(row=11, column=col, value=label)
        cx.fill = F(bg); cx.font = Ft(bold=True, color=fg, size=10)
        cx.alignment = Al(wrap=(col==4)); cx.border = B(color=G_DARK)

    # Cabeçalhos auxiliares do gráfico
    for ci, label in [(16,"Rótulo X"),(17,"% Ind"),(18,"% Acc"),(19,"ZonaA"),(20,"ZonaB"),(21,"ZonaC")]:
        cx = ws.cell(row=11, column=ci, value=label)
        cx.fill = F(DARK); cx.font = Ft(bold=True, color=GDARK2, size=7); cx.alignment = Al()

    # ── Linhas de dados ───────────────────────────────────────────────────────
    unlock = Protection(locked=False)

    for i, (sku, desc, cod, qty, cost) in enumerate(raw):
        row = DS + i
        abc_class = abc_lk.get((sku, desc), "C")

        rf     = G_LITE if abc_class=="A" else (Y_LITE if abc_class=="B" else R_LITE)
        abc_bg = G      if abc_class=="A" else (Y_DARK if abc_class=="B" else R_DARK)
        abc_fg = WHT    if abc_class!="B" else BLK

        # Col A: fórmula ABC (limites dinâmicos em B9 e B10)
        cx = ws.cell(row=row, column=1,
                     value=f'=IF($I{row}<=$B$9,"A",IF($I{row}<=$B$10,"B","C"))')
        cx.fill = F(abc_bg); cx.font = Ft(bold=True, color=abc_fg, size=12)
        cx.alignment = Al(); cx.border = B(color=BLK)

        # Col B: SKU
        cx = ws.cell(row=row, column=2, value=sku)
        cx.fill = F(rf); cx.font = Ft(color=BLK, size=9)
        cx.alignment = Al(); cx.border = B(color=GDARK2)

        # Col C: Código
        cx = ws.cell(row=row, column=3, value=cod)
        cx.fill = F(rf); cx.font = Ft(color=BLK, size=9)
        cx.alignment = Al(); cx.border = B(color=GDARK2)

        # Col D: Descrição
        cx = ws.cell(row=row, column=4, value=desc)
        cx.fill = F(rf); cx.font = Ft(color=BLK, size=10)
        cx.alignment = Al(h="left"); cx.border = B(color=GDARK2)

        # Col E: Qtd (editável)
        cx = ws.cell(row=row, column=5, value=qty)
        cx.fill = F(WHT); cx.font = Ft(bold=True, color=BLK, size=10)
        cx.alignment = Al(); cx.number_format = '#,##0'; cx.protection = unlock
        cx.border = Border(
            left=Side(style="medium", color=G_DARK), right=Side(style="medium", color=G_DARK),
            top=Side(style="thin",   color=GMID),   bottom=Side(style="thin",  color=GMID))

        # Col F: Custo Unit (editável)
        cx = ws.cell(row=row, column=6, value=cost)
        cx.fill = F(WHT); cx.font = Ft(bold=True, color=BLK, size=10)
        cx.alignment = Al(h="right"); cx.number_format = 'R$ #,##0.00'; cx.protection = unlock
        cx.border = Border(
            left=Side(style="medium", color=G_DARK), right=Side(style="medium", color=G_DARK),
            top=Side(style="thin",   color=GMID),   bottom=Side(style="thin",  color=GMID))

        # Col G: VT
        cx = ws.cell(row=row, column=7, value=f'=IFERROR(E{row}*F{row},0)')
        cx.fill = F(rf); cx.font = Ft(bold=True, color=BLK, size=10)
        cx.alignment = Al(h="right"); cx.border = B(color=GDARK2); cx.number_format = 'R$ #,##0.00'

        # Col H: % Individual
        cx = ws.cell(row=row, column=8,
                     value=f'=IFERROR(G{row}/SUM($G${DS}:$G${DE}),0)')
        cx.fill = F(rf); cx.font = Ft(color=BLK, size=10)
        cx.alignment = Al(); cx.border = B(color=GDARK2); cx.number_format = '0.00%'

        # Col I: % Acumulado
        cx = ws.cell(row=row, column=9,
                     value=f'=IFERROR(SUMIF($G${DS}:$G${DE},">="&G{row},$G${DS}:$G${DE})/SUM($G${DS}:$G${DE}),0)')
        cx.fill = F(rf); cx.font = Ft(color=BLK, size=10)
        cx.alignment = Al(); cx.border = B(color=GDARK2); cx.number_format = '0.00%'

    # ── Dados estáticos do gráfico (pré-calculados em Python, ordenados por VT) ──
    # X axis: "A — SKU01 — Servidor Backup" (nome completo do produto)
    for i, d in enumerate(sorted_data):
        row = DS + i
        # Rótulo do X com classe + SKU + nome do produto (truncado 22 chars)
        desc_short = d["desc"][:22] if len(d["desc"]) > 22 else d["desc"]
        lbl = f"{d['abc']} — {d['sku']} — {desc_short}"

        # Col P(16): rótulo X para o gráfico
        cx = ws.cell(row=row, column=16, value=lbl)
        cx.fill = F(DARK); cx.font = Ft(color=GDARK2, size=7); cx.alignment = Al(h="left")

        # Col Q(17): % Individual estática
        cx = ws.cell(row=row, column=17, value=round(d["pct_ind"], 5))
        cx.fill = F(DARK); cx.font = Ft(color=GDARK2, size=7); cx.number_format = '0.00%'

        # Col R(18): % Acumulado estática
        cx = ws.cell(row=row, column=18, value=round(d["pct_acc"], 5))
        cx.fill = F(DARK); cx.font = Ft(color=GDARK2, size=7); cx.number_format = '0.00%'

        # Cols S/T/U (19/20/21): bandas de zona (1.0 na zona correta, 0 nas demais)
        for ci_idx, cls in enumerate(["A","B","C"]):
            cx = ws.cell(row=row, column=19+ci_idx,
                         value=1.0 if d["abc"]==cls else 0.0)
            cx.fill = F([G_DARK, Y_DARK, R_DARK][ci_idx])
            cx.font = Ft(color=GDARK2, size=7)

    # ── Linha TOTAL ───────────────────────────────────────────────────────────
    tr = DE + 1
    ws.row_dimensions[tr].height = 22
    mc(ws, tr, 1, tr, 4)
    cx = ws.cell(row=tr, column=1, value="TOTAL GERAL")
    cx.fill = F(BLK); cx.font = Ft(bold=True, color=Y, size=11); cx.alignment = Al()

    cx = ws.cell(row=tr, column=5, value=f'=SUM(E{DS}:E{DE})')
    cx.fill = F(BLK); cx.font = Ft(bold=True, size=11); cx.alignment = Al(); cx.number_format = '#,##0'

    ws.cell(row=tr, column=6).fill = F(BLK)

    cx = ws.cell(row=tr, column=7, value=f'=SUM(G{DS}:G{DE})')
    cx.fill = F(G); cx.font = Ft(bold=True, size=11); cx.alignment = Al(h="right")
    cx.number_format = 'R$ #,##0.00'

    for col in [8, 9]:
        cx = ws.cell(row=tr, column=col, value=1.0)
        cx.fill = F(BLK); cx.font = Ft(bold=True, size=11); cx.alignment = Al(); cx.number_format = '0%'

    for col in range(1, 10):
        ws.cell(row=tr, column=col).border = Border(
            top=Side(style="medium", color=G), bottom=Side(style="medium", color=G),
            left=Side(style="thin",  color=G), right=Side(style="thin",  color=G))

    # ── GRÁFICO CURVA ABC ─────────────────────────────────────────────────────
    chart_row = tr + 3

    # ─ AreaChart: bandas de fundo (base do combo) ─
    area = AreaChart()
    area.grouping = "stacked"
    area.title    = f"🇯🇲 {cfg['title']} — Curva de Pareto & Classificação ABC"
    area.style    = 10
    area.width    = 32   # gráfico largo para caber os nomes dos produtos
    area.height   = 16

    area.y_axis.title  = "% do Valor Total"
    area.y_axis.numFmt = '0%'
    area.y_axis.scaling.min = 0
    area.y_axis.scaling.max = 1

    area.x_axis.title   = "Produtos (ordenados por V.T. decrescente)"
    area.x_axis.numFmt  = 'General'
    # Rotaciona os rótulos do eixo X para caber os nomes
    area.x_axis.txPr = None

    cats = Reference(ws, min_col=16, min_row=DS, max_row=DE)

    # Zona A — verde
    ser_a = Series(Reference(ws, min_col=19, min_row=DS, max_row=DE), title="🟢 Zona A (≤80%)")
    ser_a.graphicalProperties.solidFill = G_CHART
    ser_a.graphicalProperties.line.solidFill = G
    ser_a.graphicalProperties.line.width = 5000
    area.series.append(ser_a)

    # Zona B — dourado
    ser_b = Series(Reference(ws, min_col=20, min_row=DS, max_row=DE), title="🟡 Zona B (80-95%)")
    ser_b.graphicalProperties.solidFill = Y_CHART
    ser_b.graphicalProperties.line.solidFill = Y_DARK
    ser_b.graphicalProperties.line.width = 5000
    area.series.append(ser_b)

    # Zona C — vermelho/salmão
    ser_c = Series(Reference(ws, min_col=21, min_row=DS, max_row=DE), title="🔴 Zona C (95-100%)")
    ser_c.graphicalProperties.solidFill = R_CHART
    ser_c.graphicalProperties.line.solidFill = R_DARK
    ser_c.graphicalProperties.line.width = 5000
    area.series.append(ser_c)

    area.set_categories(cats)

    # ─ BarChart: % individual por produto ─
    bar = BarChart()
    bar.type = "col"

    ser_bar = Series(Reference(ws, min_col=17, min_row=DS, max_row=DE),
                     title="📊 % Individual por Produto")
    ser_bar.graphicalProperties.solidFill = BLUE_BAR
    ser_bar.graphicalProperties.line.solidFill = "1A4E8A"
    ser_bar.graphicalProperties.line.width = 8000
    bar.series.append(ser_bar)
    bar.set_categories(cats)

    # Data labels nas barras mostrando a % de cada produto
    bar.dLbls = DataLabelList()
    bar.dLbls.showVal     = True
    bar.dLbls.showSerName = False
    bar.dLbls.showCatName = False
    bar.dLbls.numFmt      = '0.0%'

    # ─ LineChart: curva % acumulado (Pareto) ─
    line = LineChart()

    ser_line = Series(Reference(ws, min_col=18, min_row=DS, max_row=DE),
                      title="📈 % Acumulado (Curva Pareto)")
    ser_line.graphicalProperties.line.solidFill = R_DARK
    ser_line.graphicalProperties.line.width      = 30000  # linha grossa e vermelha
    ser_line.smooth = True   # curva suave
    # Marcadores na linha
    ser_line.marker.symbol  = "circle"
    ser_line.marker.size    = 5
    ser_line.marker.graphicalProperties.solidFill      = R
    ser_line.marker.graphicalProperties.line.solidFill = R_DARK
    line.series.append(ser_line)
    line.set_categories(cats)

    # ─ Combina os 3 tipos em um único gráfico ─
    area += bar
    area += line

    ws.add_chart(area, f"A{chart_row}")

    # ── Formatação Condicional na tabela ──────────────────────────────────────
    rng = f"A{DS}:I{DE}"
    for cls, clr in [("A", G_LITE), ("B", Y_LITE), ("C", R_LITE)]:
        ws.conditional_formatting.add(rng, FormulaRule(
            formula=[f'$A{DS}="{cls}"'],
            fill=PatternFill(fill_type="solid", fgColor=clr),
            font=Font(color=BLK)))

    # ── Proteção ──────────────────────────────────────────────────────────────
    ws.protection.sheet    = True
    ws.protection.password = "abc2026."
    ws.protection.enable()
    ws.protection.selectLockedCells   = False
    ws.protection.selectUnlockedCells = False
    # Desbloqueia: Qtd (E), Custo (F), e células de limite (B9, B10)
    for r in range(DS, DE + 1):
        ws.cell(row=r, column=5).protection = unlock
        ws.cell(row=r, column=6).protection = unlock
    ws.cell(row=9,  column=2).protection = unlock
    ws.cell(row=10, column=2).protection = unlock

    ws.freeze_panes = f"A{DS}"
    ws.sheet_properties.tabColor = cfg["tab_color"]

    return ws


# ── MAIN ───────────────────────────────────────────────────────────────────────
def main():
    OUTPUT = "/home/user/Zyth/curva_abc_reggae.xlsx"
    wb = Workbook()
    wb.remove(wb.active)

    for cfg in SHEETS:
        print(f"  Criando aba: {cfg['name']}  ({len(cfg['items'])} itens)...")
        create_data_sheet(wb, cfg)

    wb.properties.title   = "Curva ABC Reggae — One Love, One Excel!"
    wb.properties.subject = "Dashboard ABC 3 abas | Reggae Colors"
    wb.properties.creator = "🇯🇲 Curva ABC Reggae"

    wb.save(OUTPUT)
    import os
    size = os.path.getsize(OUTPUT)
    print(f"\n✅  Salvo: {OUTPUT}")
    print(f"📦  {size/1024:.1f} KB   |   {len(SHEETS)} abas")
    print("🇯🇲  One Love, One Excel!")

if __name__ == "__main__":
    main()
