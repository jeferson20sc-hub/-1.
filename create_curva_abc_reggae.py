#!/usr/bin/env python3
"""
🇯🇲 CURVA ABC REGGAE — NOTA 10 EDITION
Fixes: colunas corretas | formatação condicional | validação | barras coloridas por classe
Bônus: Calendário 2026
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.chart import AreaChart, BarChart, LineChart, Reference, Series
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.styles.protection import Protection
import calendar as cal_module

# ── Paleta Reggae ──────────────────────────────────────────────────────────────
G_DARK   = "006B28";  G       = "009B3A";  G_LITE  = "D6F5E3";  G_CHART = "A8E6C0"
Y_DARK   = "C8A400";  Y       = "FFD700";  Y_LITE  = "FFF8D0";  Y_CHART = "FFE87C"
R_DARK   = "990000";  R       = "CC0000";  R_LITE  = "FFEDED";  R_CHART = "FFAD99"
BLK      = "000000";  WHT     = "FFFFFF";  DARK    = "141414";  DARK2   = "1E1E1E"
GDARK2   = "999999";  GMID    = "DDDDDD";  GBORDER = "CCCCCC"
LINE_CLR = "111111"   # linha Pareto quase preta — elegante

# ── Helpers ────────────────────────────────────────────────────────────────────
def F(c):  return PatternFill(fill_type="solid", fgColor=c)
def Ft(bold=False, color=WHT, size=11, italic=False, name="Calibri"):
    return Font(bold=bold, color=color, size=size, italic=italic, name=name)
def B(style="thin", color=BLK):
    s = Side(style=style, color=color)
    return Border(left=s, right=s, top=s, bottom=s)
def Al(h="center", wrap=False):
    return Alignment(horizontal=h, vertical="center", wrap_text=wrap)
def mc(ws, r1, c1, r2, c2):
    ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)

def outer_border(ws, r1, c1, r2, c2, color=G, thick="medium"):
    t = Side(style=thick, color=color); n = Side(style="thin", color=color)
    for row in ws.iter_rows(r1, r2, c1, c2):
        for cx in row:
            r, cc = cx.row, cx.column
            cx.border = Border(
                top   =t if r ==r1 else n, bottom=t if r ==r2 else n,
                left  =t if cc==c1 else n, right =t if cc==c2 else n)

# ── Dados das 3 Abas ───────────────────────────────────────────────────────────
SHEETS = [
    {
        "name": "🟢 Primeira",
        "title": "CURVA ABC REGGAE — PRIMEIRA",
        "tab_color": G_DARK,
        "items": [
            #  SKU     Descrição                   Cod  Qtd  Custo
            ("SKU01", "Servidor Backup",             3,   2, 25000),
            ("SKU02", 'Monitor 27" 4K',              5,  10,  3000),
            ("SKU03", "Notebook Administrativo",      9,   6,  4500),
            ("SKU04", "Nobreak Profissional",          1,   5,  2000),
            ("SKU05", "Switch Gerenciavel",            7,   4,  1500),
            ("SKU06", "Mouse sem fio",                 4,  40,    50),
            ("SKU07", "Patch Cord 1m",                 8, 100,    15),
            ("SKU08", "Teclado USB",                   6,  80,    35),
            ("SKU09", "Cabo Rede Cat6 (m)",             2, 200,     8),
            ("SKU10", "Mousepad Slim",                10, 120,    12),
        ],
    },
    {
        "name": "🟡 01",
        "title": "CURVA ABC REGGAE — ABA 01",
        "tab_color": Y_DARK,
        "items": [
            ("SKU01","Notebook Executivo i7",        10,  35,  6500),
            ("SKU02","Servidor de Dados Pro",          2,  12, 15000),
            ("SKU03","Impressora Multifuncional",     17,  55,  1100),
            ("SKU04","Toner Impressora Laser",        12,  68,   680),
            ("SKU05","Memoria RAM 16GB",              19,  45,   380),
            ("SKU06",'Monitor LED 24"',                4,  20,   850),
            ("SKU07","Webcam Full HD",                14,  82,   180),
            ("SKU08","Roteador Wi-Fi 6",              11,  15,   650),
            ("SKU09","SSD 480GB Sata",                15,  55,   190),
            ("SKU10","Nobreak 1500VA",                  6,   8,  1200),
            ("SKU11","HD Externo 2TB",                20,  18,   420),
            ("SKU12","Teclado Mecanico RGB",            5,  30,   250),
            ("SKU13","Pen Drive 64GB",                  9, 160,    45),
            ("SKU14","Headset com Microfone",          13,  64,   110),
            ("SKU15","Pacote de Papel A4",              7, 237,    28),
            ("SKU16","Filtro de Linha 5 Tom.",         18, 110,    40),
            ("SKU17","Cabo HDMI 2m",                    1, 150,    25),
            ("SKU18","Adaptador USB-C",                16,  90,    35),
            ("SKU19","Mouse Optico Simples",             3, 120,    15),
            ("SKU20","Patch Cord RJ45 1m",               8, 200,    12),
        ],
    },
    {
        "name": "🔴 02",
        "title": "CURVA ABC REGGAE — ABA 02",
        "tab_color": R_DARK,
        "items": [
            ("SKU01","Servidor Rack Dell Pro",          2,  15, 22500),
            ("SKU02","Workstation Grafica",            17,  19, 12000),
            ("SKU03","Notebook i7 32GB RAM",           10,  28,  7200),
            ("SKU04","Switch Gerenciavel 48p",           8,  10,  4800),
            ("SKU05","Placa de Video RTX",             22,   8,  5500),
            ("SKU06","Storage NAS 40TB",               21,   2, 18500),
            ("SKU07","Memoria RAM 16GB",               19,  45,   380),
            ("SKU08","Toner Impressora Laser",         12,  40,   320),
            ("SKU09",'Monitor LED 24"',                  4,  15,   850),
            ("SKU10","Roteador Wi-Fi 6",               11,  25,   450),
            ("SKU11","SSD 480GB Sata",                 15,  55,   190),
            ("SKU12","Nobreak 1500VA",                   6,   8,  1200),
            ("SKU13","HD Externo 2TB",                 20,  18,   420),
            ("SKU14","Teclado Mecanico RGB",             5,  30,   250),
            ("SKU15","Pen Drive 64GB",                   9, 150,    45),
            ("SKU16","Headset com Microfone",           13,  60,   110),
            ("SKU17","Pacote de Papel A4",               7, 200,    28),
            ("SKU18","Teclado de Entrada",             29, 100,    45),
            ("SKU19","Filtro de Linha",                18, 110,    40),
            ("SKU20","Webcam Full HD",                 14,  22,   180),
            ("SKU21","Pilhas AA (Pacote)",             23, 300,    12),
            ("SKU22","Adaptador USB-C",                16,  90,    35),
            ("SKU23","Suporte para Monitor",           28,  20,   140),
            ("SKU24","Cabo HDMI 2m",                     1,  85,    25),
            ("SKU25","Mousepad Simples",               24, 250,     8),
            ("SKU26","Mouse Optico Simples",             3, 120,    15),
            ("SKU27","Organizador de Cabos",           25, 180,    10),
            ("SKU28","Ar comprimido (Lata)",           27,  35,    45),
            ("SKU29","Pasta Termica",                  26,  45,    35),
            ("SKU30","Conector RJ45 (Cento)",          30,  50,    30),
        ],
    },
]

# ── Classificação ABC ──────────────────────────────────────────────────────────
def classify(items, lim_a=0.80, lim_b=0.95):
    rows = [{"sku":s,"desc":d,"cod":c,"qty":q,"cost":cu,"vt":q*cu}
            for s,d,c,q,cu in items]
    rows.sort(key=lambda x: x["vt"], reverse=True)
    total = sum(r["vt"] for r in rows); cumsum = 0
    for r in rows:
        cumsum += r["vt"]
        r["pct_ind"] = r["vt"]/total if total else 0
        r["pct_acc"] = cumsum/total  if total else 0
        r["abc"] = "A" if r["pct_acc"]<=lim_a else ("B" if r["pct_acc"]<=lim_b else "C")
    return rows, total

# ═══════════════════════════════════════════════════════════════════════════════
#  ABA DE DADOS
# ═══════════════════════════════════════════════════════════════════════════════
def create_data_sheet(wb, cfg):
    ws  = wb.create_sheet(title=cfg["name"])
    raw = cfg["items"]
    DS  = 12          # data start row (linha onde começam os dados)
    n   = len(raw)
    DE  = DS + n - 1  # data end row

    sorted_data, _ = classify(raw)
    abc_lk = {(d["sku"], d["desc"]): d["abc"] for d in sorted_data}

    # ── Larguras ──────────────────────────────────────────────────────────────
    widths = {
        "A":7,  "B":9,  "C":8,  "D":30, "E":10,    # ABC | SKU | Cod | Desc | Qtd
        "F":14, "G":16, "H":12, "I":12,             # Custo | VT | %Ind | %Acc
        "J":2,  "K":11, "L":16, "M":10, "N":10,    # spacer | resumo A/B/C
        "O":2,  "P":28, "Q":13, "R":13, "S":13,    # spacer | chart label | %indA | %indB | %indC
        "T":13, "U":10, "V":10, "W":10,             # %acc | zonaA | zonaB | zonaC
    }
    for c, w in widths.items():
        ws.column_dimensions[c].width = w

    # ── Alturas ───────────────────────────────────────────────────────────────
    ws.row_dimensions[1].height  = 54
    ws.row_dimensions[2].height  = 8
    ws.row_dimensions[3].height  = 24
    ws.row_dimensions[4].height  = 32
    ws.row_dimensions[5].height  = 24
    ws.row_dimensions[6].height  = 32
    ws.row_dimensions[7].height  = 8
    ws.row_dimensions[8].height  = 20
    ws.row_dimensions[9].height  = 24
    ws.row_dimensions[10].height = 24
    ws.row_dimensions[11].height = 24
    for r in range(DS, DE + 2):
        ws.row_dimensions[r].height = 18

    # ── Row 1: Banner ─────────────────────────────────────────────────────────
    mc(ws, 1, 1, 1, 9)
    t = ws.cell(row=1, column=1, value=f"🇯🇲  {cfg['title']}  —  One Love, One Excel! ✌️")
    t.fill = F(BLK); t.font = Font(bold=True, color=G, size=20, name="Calibri"); t.alignment = Al()
    for ci, clr in enumerate([G,Y,R,BLK,G,Y], start=10):
        ws.cell(row=1, column=ci).fill = F(clr)
    mc(ws, 1, 10, 1, 14)
    lg = ws.cell(row=1, column=10, value="🟢 A = até 80%  |  🟡 B = 80–95%  |  🔴 C = 95–100%")
    lg.fill = F(BLK); lg.font = Ft(color=Y, size=10); lg.alignment = Al()

    # ── Row 2: faixa reggae ───────────────────────────────────────────────────
    for ci in range(1, 15):
        ws.cell(row=2, column=ci).fill = F([G,Y,R,BLK][(ci-1)%4])

    # ── Rows 3-6: KPI cards ───────────────────────────────────────────────────
    for r in range(3, 7):
        for c in range(1, 10):
            ws.cell(row=r, column=c).fill = F(DARK2)

    def kpi(lr, vr, c, label, formula, bg, span=2, fmt=None, lfg=WHT):
        mc(ws, lr, c, lr, c+span-1)
        lc = ws.cell(row=lr, column=c, value=label)
        lc.fill = F(DARK2); lc.font = Ft(bold=True, color=lfg, size=9); lc.alignment = Al()
        mc(ws, vr, c, vr, c+span-1)
        vc = ws.cell(row=vr, column=c, value=formula)
        vc.fill = F(bg); vc.font = Ft(bold=True, size=14); vc.alignment = Al()
        if fmt: vc.number_format = fmt

    kpi(3,4,1,"💰 V.T. TOTAL",       f"=SUM($G${DS}:$G${DE})",              G_DARK, fmt='R$ #,##0.00')
    kpi(3,4,3,"🟢 V.T. CLASSE A",    f'=SUMIF($A${DS}:$A${DE},"A",$G${DS}:$G${DE})', G, fmt='R$ #,##0.00')
    kpi(3,4,5,"🟡 V.T. CLASSE B",    f'=SUMIF($A${DS}:$A${DE},"B",$G${DS}:$G${DE})', Y_DARK, fmt='R$ #,##0.00', lfg=BLK)
    kpi(3,4,7,"🔴 MAIOR V.T.",        f"=MAX($G${DS}:$G${DE})",              R_DARK, fmt='R$ #,##0.00')

    kpi(5,6,1,"📦 TOTAL DE ITENS",    f"=COUNTA($D${DS}:$D${DE})",           G_DARK)
    kpi(5,6,3,"% ITENS CLASSE A",
              f'=IFERROR(COUNTIF($A${DS}:$A${DE},"A")/COUNTA($D${DS}:$D${DE}),0)',
              G, fmt='0.0%')
    kpi(5,6,5,"🔴 V.T. CLASSE C",    f'=SUMIF($A${DS}:$A${DE},"C",$G${DS}:$G${DE})', R_DARK, fmt='R$ #,##0.00')
    kpi(5,6,7,"✅ STATUS",
              f'=IF(IFERROR(MAX($I${DS}:$I${DE}),0)>1.001,"⚠️ VERIFICAR","✅ OK")',
              G_DARK)

    outer_border(ws, 3, 1, 6, 9)

    # ── Row 7: faixa divider ──────────────────────────────────────────────────
    for ci in range(1, 15):
        ws.cell(row=7, column=ci).fill = F([R,Y,G,BLK][(ci-1)%4])

    # ── Rows 8-11: Painel resumo + controles de limite ────────────────────────
    # Instrução — cols A:I row 8 (CORRIGIDO: col E=Qtd, col F=Custo)
    mc(ws, 8, 1, 8, 9)
    ins = ws.cell(row=8, column=1,
                  value="  ✏️  Edite somente: Qtd. (col E) e Custo Unit. (col F)  "
                        "|  Limite A → célula B9  |  Limite B → célula B10  "
                        "|  As fórmulas V.T., % e ABC recalculam automaticamente")
    ins.fill = F(G_DARK); ins.font = Ft(italic=True, color="C8F5DB", size=8)
    ins.alignment = Al(h="left")

    # Limite A — col A label, col B valor editável
    for row, label, val, bg, fg, err_msg in [
        (9,  "Limite A (Classe A):", 0.80, G,      BLK, "Digite um valor entre 0,50 e 0,89 (ex: 0,80 para 80%)"),
        (10, "Limite B (Classe B):", 0.95, Y_DARK, WHT, "Digite um valor entre 0,50 e 0,99 (ex: 0,95 para 95%)"),
    ]:
        lbl = ws.cell(row=row, column=1, value=label)
        lbl.fill = F(DARK2); lbl.font = Ft(bold=True, color=WHT, size=9); lbl.alignment = Al(h="right")

        val_cell = ws.cell(row=row, column=2, value=val)
        val_cell.fill = F(bg); val_cell.font = Ft(bold=True, color=fg, size=14)
        val_cell.alignment = Al(); val_cell.number_format = '0%'
        val_cell.protection = Protection(locked=False)

        # ── Validação de dados (FIX: corrigido para aceitar apenas 0-1) ──────
        dv = DataValidation(
            type="decimal", operator="between", formula1="0.5", formula2="0.99",
            error=err_msg,
            errorTitle="❌ Valor Inválido — Tente Novamente",
            prompt=f"Digite um percentual decimal (ex: 0,80 para 80%)",
            promptTitle="ℹ️ Dica de Preenchimento",
            showErrorMessage=True, showInputMessage=True)
        ws.add_data_validation(dv)
        dv.add(val_cell)

    # Painel de resumo A/B/C — cols K-N
    mc(ws, 8, 11, 8, 14)
    sh = ws.cell(row=8, column=11, value="RESUMO POR CLASSE")
    sh.fill = F(BLK); sh.font = Ft(bold=True, color=Y, size=10); sh.alignment = Al()
    # Header das colunas do resumo (rows 8 are now merged, use row=8 only for col 11 above)
    # Add sub-headers in a separate logic after unmerging conceptually
    for ci, lbl in [(11,"Classe"),(12,"V.T. Total"),(13,"% do Total"),(14,"Qtd. Itens")]:
        if ci == 11: continue  # already written above as merged header
        cx = ws.cell(row=8, column=ci)
        cx.fill = F("222222"); cx.font = Ft(bold=True, size=9)
        cx.alignment = Al(); cx.border = B(color=GDARK2)

    for cls, bg, row in [("A",G,9),("B",Y_DARK,10),("C",R_DARK,11)]:
        fg = BLK if cls=="B" else WHT
        for ci, val, fmt in [
            (11, cls, None),
            (12, f'=SUMIF($A${DS}:$A${DE},"{cls}",$G${DS}:$G${DE})', 'R$ #,##0'),
            (13, f'=IFERROR(L{row}/SUM($G${DS}:$G${DE}),0)', '0.0%'),
            (14, f'=COUNTIF($A${DS}:$A${DE},"{cls}")', None),
        ]:
            cx = ws.cell(row=row, column=ci, value=val)
            cx.fill = F(bg); cx.font = Ft(bold=True, color=fg, size=11 if ci==11 else 10)
            cx.alignment = Al(); cx.border = B(color=BLK)
            if fmt: cx.number_format = fmt

    # ── Row 11: Cabeçalho das colunas de dados ────────────────────────────────
    ws.row_dimensions[11].height = 26
    headers = [
        (1,"ABC",BLK,G),(2,"SKU",BLK,WHT),(3,"Cod.",BLK,WHT),
        (4,"Descrição do Produto",BLK,WHT),
        (5,"Qtd.",BLK,Y),          # ← col E = Qtd (CORRIGIDO)
        (6,"Custo Unit.(R$)",BLK,Y), # ← col F = Custo (CORRIGIDO)
        (7,"V.T. (R$)",BLK,G),
        (8,"% Individual",BLK,WHT),
        (9,"% Acumulado",BLK,WHT),
    ]
    for col, label, bg, fg in headers:
        cx = ws.cell(row=11, column=col, value=label)
        cx.fill = F(bg); cx.font = Ft(bold=True, color=fg, size=10)
        cx.alignment = Al(wrap=(col==4)); cx.border = B(color=G_DARK)

    # Cabeçalhos cols de dados do gráfico
    for ci, lbl in [(16,"Rótulo X"),(17,"%IndA"),(18,"%IndB"),(19,"%IndC"),
                    (20,"%Acum."),(21,"ZonaA"),(22,"ZonaB"),(23,"ZonaC")]:
        cx = ws.cell(row=11, column=ci, value=lbl)
        cx.fill = F(DARK2); cx.font = Ft(bold=True, color=GDARK2, size=7); cx.alignment = Al()

    # ── Linhas de dados ───────────────────────────────────────────────────────
    unlock = Protection(locked=False)

    for i, (sku, desc, cod, qty, cost) in enumerate(raw):
        row = DS + i
        abc_class = abc_lk.get((sku, desc), "C")
        rf     = G_LITE if abc_class=="A" else (Y_LITE if abc_class=="B" else R_LITE)
        abc_bg = G      if abc_class=="A" else (Y_DARK if abc_class=="B" else R_DARK)
        abc_fg = WHT    if abc_class!="B" else BLK

        # Col A: fórmula ABC (usa B9=LimA e B10=LimB)
        cx = ws.cell(row=row, column=1,
                     value=f'=IF($I{row}<=$B$9,"A",IF($I{row}<=$B$10,"B","C"))')
        cx.fill = F(abc_bg); cx.font = Ft(bold=True, color=abc_fg, size=12)
        cx.alignment = Al(); cx.border = B(color=BLK)

        # Col B: SKU
        cx = ws.cell(row=row, column=2, value=sku)
        cx.fill = F(rf); cx.font = Ft(color=BLK, size=9); cx.alignment = Al(); cx.border = B(color=GBORDER)

        # Col C: Código
        cx = ws.cell(row=row, column=3, value=cod)
        cx.fill = F(rf); cx.font = Ft(color=BLK, size=9); cx.alignment = Al(); cx.border = B(color=GBORDER)

        # Col D: Descrição
        cx = ws.cell(row=row, column=4, value=desc)
        cx.fill = F(rf); cx.font = Ft(color=BLK, size=10)
        cx.alignment = Al(h="left"); cx.border = B(color=GBORDER)

        # Col E: Qtd (editável — col E CORRIGIDO)
        cx = ws.cell(row=row, column=5, value=qty)
        cx.fill = F(WHT); cx.font = Ft(bold=True, color=BLK, size=10)
        cx.alignment = Al(); cx.number_format = '#,##0'; cx.protection = unlock
        cx.border = Border(
            left=Side(style="medium",color=G_DARK), right=Side(style="medium",color=G_DARK),
            top=Side(style="thin",  color=GMID),   bottom=Side(style="thin", color=GMID))

        # Col F: Custo Unit (editável — col F CORRIGIDO)
        cx = ws.cell(row=row, column=6, value=cost)
        cx.fill = F(WHT); cx.font = Ft(bold=True, color=BLK, size=10)
        cx.alignment = Al(h="right"); cx.number_format = 'R$ #,##0.00'; cx.protection = unlock
        cx.border = Border(
            left=Side(style="medium",color=G_DARK), right=Side(style="medium",color=G_DARK),
            top=Side(style="thin",  color=GMID),   bottom=Side(style="thin", color=GMID))

        # Col G: VT (fórmula = E*F — col G CORRIGIDO)
        cx = ws.cell(row=row, column=7, value=f'=IFERROR(E{row}*F{row},0)')
        cx.fill = F(rf); cx.font = Ft(bold=True, color=BLK, size=10)
        cx.alignment = Al(h="right"); cx.border = B(color=GBORDER); cx.number_format = 'R$ #,##0.00'

        # Col H: % Individual
        cx = ws.cell(row=row, column=8,
                     value=f'=IFERROR(G{row}/SUM($G${DS}:$G${DE}),0)')
        cx.fill = F(rf); cx.font = Ft(color=BLK, size=10)
        cx.alignment = Al(); cx.border = B(color=GBORDER); cx.number_format = '0.00%'

        # Col I: % Acumulado (SUMIF decrescente — não precisa de ordenação manual)
        cx = ws.cell(row=row, column=9,
                     value=f'=IFERROR(SUMIF($G${DS}:$G${DE},">="&G{row},$G${DS}:$G${DE})/SUM($G${DS}:$G${DE}),0)')
        cx.fill = F(rf); cx.font = Ft(color=BLK, size=10)
        cx.alignment = Al(); cx.border = B(color=GBORDER); cx.number_format = '0.00%'

    # ── Dados estáticos do gráfico (ordenados por VT decrescente) ─────────────
    for i, d in enumerate(sorted_data):
        row  = DS + i
        desc_short = d["desc"][:20] if len(d["desc"])>20 else d["desc"]
        lbl  = f"{d['abc']} — {d['sku']} — {desc_short}"

        # Col P(16): rótulo do eixo X
        cx = ws.cell(row=row, column=16, value=lbl)
        cx.fill = F(DARK2); cx.font = Ft(color=GDARK2, size=7); cx.alignment = Al(h="left")

        # Cols Q/R/S (17/18/19): % individual por classe (só a classe correta tem valor)
        for ci_off, cls in enumerate(["A","B","C"]):
            cx = ws.cell(row=row, column=17+ci_off,
                         value=round(d["pct_ind"],5) if d["abc"]==cls else 0.0)
            cx.fill = F([G_DARK,Y_DARK,R_DARK][ci_off])
            cx.font = Ft(color=GDARK2, size=7); cx.number_format = '0.00%'

        # Col T(20): % acumulado
        cx = ws.cell(row=row, column=20, value=round(d["pct_acc"],5))
        cx.fill = F(DARK2); cx.font = Ft(color=GDARK2, size=7); cx.number_format = '0.00%'

        # Cols U/V/W (21/22/23): bandas de fundo (1.0 na zona correta)
        for ci_off, cls in enumerate(["A","B","C"]):
            cx = ws.cell(row=row, column=21+ci_off,
                         value=1.0 if d["abc"]==cls else 0.0)
            cx.fill = F([G_DARK,Y_DARK,R_DARK][ci_off])
            cx.font = Ft(color=GDARK2, size=7)

    # ── Linha TOTAL ───────────────────────────────────────────────────────────
    tr = DE + 1
    ws.row_dimensions[tr].height = 24
    mc(ws, tr, 1, tr, 4)
    cx = ws.cell(row=tr, column=1, value="TOTAL GERAL")
    cx.fill = F(BLK); cx.font = Ft(bold=True, color=Y, size=12); cx.alignment = Al()

    cx = ws.cell(row=tr, column=5, value=f'=SUM(E{DS}:E{DE})')
    cx.fill = F(BLK); cx.font = Ft(bold=True, size=12); cx.alignment = Al(); cx.number_format = '#,##0'

    ws.cell(row=tr, column=6).fill = F(BLK)

    cx = ws.cell(row=tr, column=7, value=f'=SUM(G{DS}:G{DE})')
    cx.fill = F(G); cx.font = Ft(bold=True, size=12); cx.alignment = Al(h="right")
    cx.number_format = 'R$ #,##0.00'

    for col in [8, 9]:
        cx = ws.cell(row=tr, column=col, value=1.0)
        cx.fill = F(BLK); cx.font = Ft(bold=True, size=12); cx.alignment = Al(); cx.number_format = '0%'

    for col in range(1, 10):
        ws.cell(row=tr, column=col).border = Border(
            top=Side(style="medium",color=G), bottom=Side(style="medium",color=G),
            left=Side(style="thin",color=G),  right=Side(style="thin",color=G))

    # ── GRÁFICO CURVA ABC ─────────────────────────────────────────────────────
    chart_row = tr + 3
    cats = Reference(ws, min_col=16, min_row=DS, max_row=DE)

    # ─ Base: AreaChart com as bandas de fundo (A/B/C) ─
    area = AreaChart()
    area.grouping = "stacked"
    area.title  = f"🇯🇲  {cfg['title']}  —  Curva ABC de Pareto"
    area.style  = 10
    area.width  = 34
    area.height = 16

    area.y_axis.title     = "% do Valor Total"
    area.y_axis.numFmt    = '0%'
    area.y_axis.scaling.min = 0
    area.y_axis.scaling.max = 1
    area.x_axis.title = "Produtos (ordenados por V.T. decrescente — nome do produto no eixo X)"

    for col_idx, (cls_lbl, clr_fill, clr_line) in enumerate([
        ("🟢 Faixa A (≤80%)",  G_CHART, G),
        ("🟡 Faixa B (80-95%)", Y_CHART, Y_DARK),
        ("🔴 Faixa C (>95%)",   R_CHART, R_DARK),
    ]):
        ser = Series(Reference(ws, min_col=21+col_idx, min_row=DS, max_row=DE), title=cls_lbl)
        ser.graphicalProperties.solidFill      = clr_fill
        ser.graphicalProperties.line.solidFill = clr_line
        ser.graphicalProperties.line.width     = 6000
        area.series.append(ser)

    area.set_categories(cats)

    # ─ BarChart: barras coloridas por classe (A=verde, B=dourado, C=vermelho) ─
    bar = BarChart()
    bar.type     = "col"
    bar.grouping = "stacked"

    for col_idx, (cls_lbl, clr_fill, clr_line) in enumerate([
        ("🟢 Classe A",  G,      G_DARK),
        ("🟡 Classe B",  Y,      Y_DARK),
        ("🔴 Classe C",  R,      R_DARK),
    ]):
        ser = Series(Reference(ws, min_col=17+col_idx, min_row=DS, max_row=DE), title=cls_lbl)
        ser.graphicalProperties.solidFill      = clr_fill
        ser.graphicalProperties.line.solidFill = clr_line
        ser.graphicalProperties.line.width     = 8000
        bar.series.append(ser)

    bar.set_categories(cats)

    # Data labels nas barras (mostra % de cada produto)
    bar.dLbls            = DataLabelList()
    bar.dLbls.showVal    = True
    bar.dLbls.showSerName = False
    bar.dLbls.showCatName = False
    bar.dLbls.numFmt     = '0.0%'

    # ─ LineChart: curva Pareto — preta (como recomendaram os jurados) ─────────
    line = LineChart()
    ser_line = Series(Reference(ws, min_col=20, min_row=DS, max_row=DE),
                      title="📈 % Acumulado (Curva Pareto)")
    ser_line.graphicalProperties.line.solidFill = LINE_CLR   # preto — elegante
    ser_line.graphicalProperties.line.width     = 28000      # ~2.8pt
    ser_line.smooth = True
    ser_line.marker.symbol  = "circle"
    ser_line.marker.size    = 5
    ser_line.marker.graphicalProperties.solidFill      = BLK
    ser_line.marker.graphicalProperties.line.solidFill = WHT
    line.series.append(ser_line)
    line.set_categories(cats)

    # ─ Combina os 3 tipos num único gráfico ──────────────────────────────────
    area += bar
    area += line
    ws.add_chart(area, f"A{chart_row}")

    # ── Formatação Condicional (CORRIGIDO — linhas inteiras coloridas) ─────────
    # Aplica em A:I (todas as 9 colunas de dados visíveis)
    rng = f"A{DS}:I{DE}"
    for cls, clr in [("A", G_LITE), ("B", Y_LITE), ("C", R_LITE)]:
        # IMPORTANTE: $A{DS} → $ na coluna A, linha relativa → cada linha avalia sua própria col A
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
    # Apenas col E (Qtd), col F (Custo Unit), B9 (LimA) e B10 (LimB) editáveis
    for r in range(DS, DE+1):
        ws.cell(row=r, column=5).protection = unlock   # col E
        ws.cell(row=r, column=6).protection = unlock   # col F
    ws.cell(row=9,  column=2).protection = unlock      # B9
    ws.cell(row=10, column=2).protection = unlock      # B10

    ws.freeze_panes = f"A{DS}"
    ws.sheet_properties.tabColor = cfg["tab_color"]
    return ws


# ═══════════════════════════════════════════════════════════════════════════════
#  ABA CALENDÁRIO 2026
# ═══════════════════════════════════════════════════════════════════════════════
def create_calendar_sheet(wb, year=2026):
    ws = wb.create_sheet(title="📅 CALENDÁRIO")

    MONTHS_PT = ["Janeiro","Fevereiro","Março","Abril","Maio","Junho",
                 "Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"]
    DAYS_PT   = ["Seg","Ter","Qua","Qui","Sex","Sáb","Dom"]

    # Layout: 4 meses por linha, 3 linhas = 12 meses
    # Cada mês: 7 colunas de dias + 1 coluna espaçador = 8 colunas
    DAY_W  = 5.2    # largura de cada coluna-dia
    SEP_W  = 1.5    # largura do espaçador
    MONTH_ROWS = 11 # linhas por bloco de mês (título + header + até 6 semanas + espaçador)

    # Larguras das colunas
    for ci in range(1, 34):
        col_letter = get_column_letter(ci)
        # Coluna espaçador: a cada 8 colunas (0-indexed: 7, 15, 23)
        if (ci - 1) % 8 == 7:
            ws.column_dimensions[col_letter].width = SEP_W
        else:
            ws.column_dimensions[col_letter].width = DAY_W

    # Background geral
    for r in range(1, 48):
        ws.row_dimensions[r].height = 17
        for c in range(1, 34):
            ws.cell(row=r, column=c).fill = F(DARK)

    # ── Título ────────────────────────────────────────────────────────────────
    ws.row_dimensions[1].height = 10
    for ci in range(1, 34):
        ws.cell(row=1, column=ci).fill = F([G,Y,R,BLK][(ci-1)%4])

    ws.row_dimensions[2].height = 54
    mc(ws, 2, 1, 2, 32)
    t = ws.cell(row=2, column=1,
                value=f"🇯🇲  CALENDÁRIO {year}  —  One Love, One Excel! ✌️")
    t.fill = F(BLK); t.font = Font(bold=True, color=G, size=24, name="Calibri")
    t.alignment = Al()

    ws.row_dimensions[3].height = 10
    for ci in range(1, 34):
        ws.cell(row=3, column=ci).fill = F([R,Y,G,BLK][(ci-1)%4])

    ws.row_dimensions[4].height = 8

    # Legenda
    ws.row_dimensions[5].height = 18
    mc(ws, 5, 1, 5, 16)
    leg = ws.cell(row=5, column=1,
                  value="🟢 Itens Classe A — revisar diariamente   |   🟡 Itens Classe B — revisar semanalmente   |   🔴 Itens Classe C — revisar mensalmente")
    leg.fill = F(BLK); leg.font = Ft(color=Y, size=9, italic=True); leg.alignment = Al(h="left")

    ws.row_dimensions[6].height = 6

    START_ROW = 7
    # Cores dos cabeçalhos de mês (A/B/C em ciclo reggae)
    month_hdr_colors = [G_DARK,G_DARK,G_DARK, Y_DARK,Y_DARK,Y_DARK, R_DARK,R_DARK,R_DARK, G_DARK,Y_DARK,R_DARK]

    for month in range(1, 13):
        row_idx = (month - 1) // 4   # 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2
        col_idx = (month - 1) % 4    # 0, 1, 2, 3 repetido

        mr    = START_ROW + row_idx * MONTH_ROWS  # row inicial do bloco
        mc_c  = 1     + col_idx * 8               # col inicial do bloco (1,9,17,25)

        # ─ Título do mês ─
        ws.row_dimensions[mr].height = 22
        mc(ws, mr, mc_c, mr, mc_c+6)
        mh = ws.cell(row=mr, column=mc_c, value=f"{MONTHS_PT[month-1].upper()} {year}")
        mh.fill = F(month_hdr_colors[month-1])
        mh.font = Ft(bold=True, size=11)
        mh.alignment = Al()
        mh.border = Border(bottom=Side(style="medium",color=WHT))

        # ─ Header dos dias da semana ─
        ws.row_dimensions[mr+1].height = 18
        for d_idx, day_name in enumerate(DAYS_PT):
            is_weekend = d_idx >= 5
            cx = ws.cell(row=mr+1, column=mc_c+d_idx, value=day_name)
            cx.fill = F(R_DARK if is_weekend else G_DARK)
            cx.font = Ft(bold=True, color=WHT, size=9)
            cx.alignment = Al()
            cx.border = B(style="thin", color=DARK)

        # ─ Datas ─
        weeks = cal_module.monthcalendar(year, month)  # semanas com seg=0
        for w_idx, week in enumerate(weeks):
            ws.row_dimensions[mr+2+w_idx].height = 16
            for d_idx, day_num in enumerate(week):
                cx = ws.cell(row=mr+2+w_idx, column=mc_c+d_idx)
                if day_num == 0:
                    cx.fill = F("0A0A0A"); cx.value = ""
                else:
                    is_weekend = d_idx >= 5
                    # Verifica se é dia especial de revisão (1°, 15°, último)
                    last_day = cal_module.monthrange(year, month)[1]
                    is_review = day_num in [1, 8, 15, 22]
                    if is_weekend:
                        bg, fg = "1A0000", R_LITE
                    elif is_review:
                        bg, fg = G_DARK,   Y_LITE
                    else:
                        bg, fg = "0D1A0D",  G_LITE
                    cx.fill = F(bg); cx.font = Ft(color=fg, size=9, bold=is_review)
                    cx.value = day_num; cx.alignment = Al()
                    cx.border = B(style="hair", color="1A1A1A")

        # Preenche linhas de semanas não usadas (para alinhar os blocos)
        for w_idx in range(len(weeks), 6):
            ws.row_dimensions[mr+2+w_idx].height = 16
            for d_idx in range(7):
                ws.cell(row=mr+2+w_idx, column=mc_c+d_idx).fill = F("0A0A0A")

        # Espaçador entre meses (última coluna do bloco)
        if col_idx < 3:
            for r_off in range(MONTH_ROWS):
                ws.cell(row=mr+r_off, column=mc_c+7).fill = F(DARK)

    # Linha de legenda final
    footer_row = START_ROW + 3 * MONTH_ROWS
    ws.row_dimensions[footer_row].height = 20
    mc(ws, footer_row, 1, footer_row, 32)
    f = ws.cell(row=footer_row, column=1,
                value="🟢 = Dias de revisão ABC marcados   |   🔴 = Finais de semana   |   🇯🇲 Jah bless your schedule! — One Love, One Excel! 2026")
    f.fill = F(BLK); f.font = Ft(bold=True, color=G, size=10); f.alignment = Al()

    ws.sheet_properties.tabColor = Y_DARK
    return ws


# ── MAIN ───────────────────────────────────────────────────────────────────────
def main():
    OUTPUT = "/home/user/Zyth/curva_abc_reggae.xlsx"
    wb = Workbook()
    wb.remove(wb.active)

    for cfg in SHEETS:
        print(f"  Criando aba {cfg['name']} ({len(cfg['items'])} itens)...")
        create_data_sheet(wb, cfg)

    print("  Criando aba Calendário 2026...")
    create_calendar_sheet(wb, year=2026)

    wb.properties.title   = "Curva ABC Reggae — Nota 10 Edition"
    wb.properties.subject = "Dashboard ABC 3 abas + Calendário 2026 | Reggae Colors"
    wb.properties.creator = "🇯🇲 Curva ABC Reggae"

    wb.save(OUTPUT)
    import os
    size = os.path.getsize(OUTPUT)
    print(f"\n  Salvo: {OUTPUT}")
    print(f"  Tamanho: {size/1024:.1f} KB  |  {len(wb.sheetnames)} abas")
    print("  One Love, One Excel!")

if __name__ == "__main__":
    main()
