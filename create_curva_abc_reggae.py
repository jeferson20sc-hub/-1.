#!/usr/bin/env python3
"""
🇯🇲 CURVA ABC REGGAE — SUPER VERSION (NOTA 10 COMPLETO)
Fixes: CF funcional | cols auxiliares ocultas | Pareto VT + eixo secundário | comentários | visual limpo
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, LineChart, Reference, Series
from openpyxl.chart.label import DataLabelList
from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.styles.protection import Protection
import calendar as cal_module

# ── Cores ──────────────────────────────────────────────────────────────────────
G_DARK   = "006B28";  G      = "009B3A";  G_LITE  = "D6F5E3"
Y_DARK   = "C8A400";  Y      = "FFD700";  Y_LITE  = "FFFBD0"
R_DARK   = "8B0000";  R      = "CC0000";  R_LITE  = "FFECEC"
BLK      = "000000";  WHT    = "FFFFFF";  DARK    = "141414"
GDARK2   = "888888";  GMID   = "DDDDDD";  GBORDER = "CCCCCC"
LINE_CLR = "0A0A0A"   # preto-quase para a curva Pareto

def F(c):
    # ARGB: "FF" prefix = fully opaque (required; openpyxl defaults to "00" alpha
    # which is transparent — fine for cell xf fills but BREAKS CF dxf fills)
    return PatternFill(fill_type="solid", fgColor="FF"+c if len(c)==6 else c)
def Ft(bold=False, color=WHT, size=11, italic=False, name="Calibri"):
    return Font(bold=bold, color=color, size=size, italic=italic, name=name)
def B(style="thin", color=BLK):
    s = Side(style=style, color=color)
    return Border(left=s, right=s, top=s, bottom=s)
def Al(h="center", wrap=False):
    return Alignment(horizontal=h, vertical="center", wrap_text=wrap)
def mc(ws, r1, c1, r2, c2):
    ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)

def frame(ws, r1, c1, r2, c2, color=G, w="medium"):
    t = Side(style=w, color=color); n = Side(style="thin", color=color)
    for row in ws.iter_rows(r1, r2, c1, c2):
        for cx in row:
            r, cc = cx.row, cx.column
            cx.border = Border(
                top   = t if r==r1  else n, bottom = t if r==r2  else n,
                left  = t if cc==c1 else n, right  = t if cc==c2 else n)

# ── Dados ──────────────────────────────────────────────────────────────────────
SHEETS = [
    {
        "name": "🟢 Primeira",  "title": "CURVA ABC — PRIMEIRA",
        "tab_color": G_DARK,
        "items": [
            #  SKU      Descrição                   Cod  Qtd    Custo
            ("SKU01",  "Servidor Backup",             3,   2,  25000),
            ("SKU02",  'Monitor 27" 4K',              5,  10,   3000),
            ("SKU03",  "Notebook Administrativo",      9,   6,   4500),
            ("SKU04",  "Nobreak Profissional",          1,   5,   2000),
            ("SKU05",  "Switch Gerenciavel",            7,   4,   1500),
            ("SKU06",  "Mouse sem fio",                 4,  40,     50),
            ("SKU07",  "Patch Cord 1m",                 8, 100,     15),
            ("SKU08",  "Teclado USB",                   6,  80,     35),
            ("SKU09",  "Cabo Rede Cat6 (m)",             2, 200,      8),
            ("SKU10",  "Mousepad Slim",                10, 120,     12),
        ],
    },
    {
        "name": "🟡 01",  "title": "CURVA ABC — ABA 01",
        "tab_color": Y_DARK,
        "items": [
            ("SKU01", "Notebook Executivo i7",        10,  35,  6500),
            ("SKU02", "Servidor de Dados Pro",          2,  12, 15000),
            ("SKU03", "Impressora Multifuncional",     17,  55,  1100),
            ("SKU04", "Toner Impressora Laser",        12,  68,   680),
            ("SKU05", "Memoria RAM 16GB",              19,  45,   380),
            ("SKU06", 'Monitor LED 24"',                4,  20,   850),
            ("SKU07", "Webcam Full HD",                14,  82,   180),
            ("SKU08", "Roteador Wi-Fi 6",              11,  15,   650),
            ("SKU09", "SSD 480GB Sata",                15,  55,   190),
            ("SKU10", "Nobreak 1500VA",                  6,   8,  1200),
            ("SKU11", "HD Externo 2TB",                20,  18,   420),
            ("SKU12", "Teclado Mecanico RGB",            5,  30,   250),
            ("SKU13", "Pen Drive 64GB",                  9, 160,    45),
            ("SKU14", "Headset com Microfone",          13,  64,   110),
            ("SKU15", "Pacote de Papel A4",              7, 237,    28),
            ("SKU16", "Filtro de Linha 5 Tom.",         18, 110,    40),
            ("SKU17", "Cabo HDMI 2m",                    1, 150,    25),
            ("SKU18", "Adaptador USB-C",                16,  90,    35),
            ("SKU19", "Mouse Optico Simples",             3, 120,    15),
            ("SKU20", "Patch Cord RJ45 1m",               8, 200,    12),
        ],
    },
    {
        "name": "🔴 02",  "title": "CURVA ABC — ABA 02",
        "tab_color": R_DARK,
        "items": [
            ("SKU01", "Servidor Rack Dell Pro",          2,  15, 22500),
            ("SKU02", "Workstation Grafica",            17,  19, 12000),
            ("SKU03", "Notebook i7 32GB RAM",           10,  28,  7200),
            ("SKU04", "Switch Gerenciavel 48p",           8,  10,  4800),
            ("SKU05", "Placa de Video RTX",             22,   8,  5500),
            ("SKU06", "Storage NAS 40TB",               21,   2, 18500),
            ("SKU07", "Memoria RAM 16GB",               19,  45,   380),
            ("SKU08", "Toner Impressora Laser",         12,  40,   320),
            ("SKU09", 'Monitor LED 24"',                  4,  15,   850),
            ("SKU10", "Roteador Wi-Fi 6",               11,  25,   450),
            ("SKU11", "SSD 480GB Sata",                 15,  55,   190),
            ("SKU12", "Nobreak 1500VA",                   6,   8,  1200),
            ("SKU13", "HD Externo 2TB",                 20,  18,   420),
            ("SKU14", "Teclado Mecanico RGB",             5,  30,   250),
            ("SKU15", "Pen Drive 64GB",                   9, 150,    45),
            ("SKU16", "Headset com Microfone",           13,  60,   110),
            ("SKU17", "Pacote de Papel A4",               7, 200,    28),
            ("SKU18", "Teclado de Entrada",             29, 100,    45),
            ("SKU19", "Filtro de Linha",                18, 110,    40),
            ("SKU20", "Webcam Full HD",                 14,  22,   180),
            ("SKU21", "Pilhas AA (Pacote)",             23, 300,    12),
            ("SKU22", "Adaptador USB-C",                16,  90,    35),
            ("SKU23", "Suporte para Monitor",           28,  20,   140),
            ("SKU24", "Cabo HDMI 2m",                     1,  85,    25),
            ("SKU25", "Mousepad Simples",               24, 250,     8),
            ("SKU26", "Mouse Optico Simples",             3, 120,    15),
            ("SKU27", "Organizador de Cabos",           25, 180,    10),
            ("SKU28", "Ar comprimido (Lata)",           27,  35,    45),
            ("SKU29", "Pasta Termica",                  26,  45,    35),
            ("SKU30", "Conector RJ45 (Cento)",          30,  50,    30),
        ],
    },
]

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
def create_data_sheet(wb, cfg):
    ws  = wb.create_sheet(title=cfg["name"])
    raw = cfg["items"]
    DS  = 12          # data start row
    n   = len(raw)
    DE  = DS + n - 1

    sorted_data, _ = classify(raw)
    abc_lk = {(d["sku"], d["desc"]): d["abc"] for d in sorted_data}

    # ── Larguras ──────────────────────────────────────────────────────────────
    widths = {
        "A":7, "B":9, "C":7, "D":28,
        "E":10,"F":14,"G":16,"H":12,"I":12,
        "J":2, "K":10,"L":16,"M":10,"N":10,
    }
    for c, w in widths.items():
        ws.column_dimensions[c].width = w

    # Colunas auxiliares do gráfico (P-T) — OCULTAS
    for c_letter, w in [("P",22),("Q",12),("R",12),("S",12),("T",12)]:
        ws.column_dimensions[c_letter].width = w
        ws.column_dimensions[c_letter].hidden = True  # ← OCULTAS

    # ── Alturas ───────────────────────────────────────────────────────────────
    ws.row_dimensions[1].height  = 52
    ws.row_dimensions[2].height  = 8
    for r in [3,5]: ws.row_dimensions[r].height = 24
    for r in [4,6]: ws.row_dimensions[r].height = 30
    ws.row_dimensions[7].height  = 8
    ws.row_dimensions[8].height  = 18
    ws.row_dimensions[9].height  = 24
    ws.row_dimensions[10].height = 24
    ws.row_dimensions[11].height = 24
    for r in range(DS, DE+3): ws.row_dimensions[r].height = 18

    # ── Row 1: Banner ─────────────────────────────────────────────────────────
    mc(ws, 1, 1, 1, 9)
    t = ws.cell(row=1, column=1, value=f"  {cfg['title']}  —  Curva ABC de Pareto | Reggae Dashboard")
    t.fill = F(BLK); t.font = Font(bold=True, color=G, size=18, name="Calibri"); t.alignment = Al()
    for ci, clr in enumerate([G,Y,R,BLK,G,Y], start=10):
        ws.cell(row=1, column=ci).fill = F(clr)
    mc(ws, 1, 10, 1, 14)
    lg = ws.cell(row=1, column=10, value="Classe A = ate 80%   |   Classe B = 80-95%   |   Classe C = 95-100%")
    lg.fill = F(BLK); lg.font = Ft(color=Y, size=10); lg.alignment = Al()

    # ── Row 2: faixa reggae ───────────────────────────────────────────────────
    for ci in range(1, 15): ws.cell(row=2, column=ci).fill = F([G,Y,R,BLK][(ci-1)%4])

    # ── Rows 3-6: KPI cards ───────────────────────────────────────────────────
    for r in range(3, 7):
        for c in range(1, 10): ws.cell(row=r, column=c).fill = F(DARK)

    def kpi(lr, vr, c, label, formula, bg, span=2, fmt=None, lfg=WHT):
        mc(ws, lr, c, lr, c+span-1)
        lc = ws.cell(row=lr, column=c, value=label)
        lc.fill = F(DARK); lc.font = Ft(bold=True, color=lfg, size=9); lc.alignment = Al()
        mc(ws, vr, c, vr, c+span-1)
        vc = ws.cell(row=vr, column=c, value=formula)
        vc.fill = F(bg); vc.font = Ft(bold=True, size=14); vc.alignment = Al()
        if fmt: vc.number_format = fmt

    kpi(3,4,1,"VALOR TOTAL",         f"=SUM($G${DS}:$G${DE})",                      G_DARK, fmt='R$ #,##0.00')
    kpi(3,4,3,"CLASSE A (VT)",        f'=SUMIF($A${DS}:$A${DE},"A",$G${DS}:$G${DE})', G,     fmt='R$ #,##0.00')
    kpi(3,4,5,"CLASSE B (VT)",        f'=SUMIF($A${DS}:$A${DE},"B",$G${DS}:$G${DE})', Y_DARK,fmt='R$ #,##0.00', lfg=BLK)
    kpi(3,4,7,"MAIOR V.T.",           f"=MAX($G${DS}:$G${DE})",                       R_DARK,fmt='R$ #,##0.00')
    kpi(5,6,1,"TOTAL DE ITENS",       f"=COUNTA($D${DS}:$D${DE})",                    G_DARK)
    kpi(5,6,3,"% ITENS CLASSE A",     f'=IFERROR(COUNTIF($A${DS}:$A${DE},"A")/COUNTA($D${DS}:$D${DE}),0)', G,fmt='0.0%')
    kpi(5,6,5,"CLASSE C (VT)",        f'=SUMIF($A${DS}:$A${DE},"C",$G${DS}:$G${DE})', R_DARK,fmt='R$ #,##0.00')
    kpi(5,6,7,"STATUS",               f'=IF(IFERROR(MAX($I${DS}:$I${DE}),0)>1.001,"VERIFICAR","OK")', G_DARK)

    frame(ws, 3, 1, 6, 9)

    # ── Row 7: divider ────────────────────────────────────────────────────────
    for ci in range(1, 15): ws.cell(row=7, column=ci).fill = F([R,Y,G,BLK][(ci-1)%4])

    # ── Row 8: instrução (CORRIGIDA) ──────────────────────────────────────────
    mc(ws, 8, 1, 8, 9)
    ins = ws.cell(row=8, column=1,
                  value="  Edite somente: Qtd. (col E) e Custo Unit. (col F) "
                        "| Limite A → celula B9  |  Limite B → celula B10 "
                        "| V.T. e ABC recalculam automaticamente")
    ins.fill = F(G_DARK); ins.font = Ft(italic=True, color="C8F5DB", size=8); ins.alignment = Al(h="left")

    # ── Rows 9-10: controles de limite (com comentários) ──────────────────────
    for row, label, val, bg, fg in [
        (9,  "Limite A (Classe A):", 0.80, G,      BLK),
        (10, "Limite B (Classe B):", 0.95, Y_DARK, WHT),
    ]:
        lbl = ws.cell(row=row, column=1, value=label)
        lbl.fill = F(DARK); lbl.font = Ft(bold=True, color=WHT, size=9); lbl.alignment = Al(h="right")

        vc = ws.cell(row=row, column=2, value=val)
        vc.fill = F(bg); vc.font = Ft(bold=True, color=fg, size=14)
        vc.alignment = Al(); vc.number_format = '0%'
        vc.protection = Protection(locked=False)

        # Comentário na célula de limite
        tip = (f"EDITAVEL: Limite para Classe {'A' if row==9 else 'B'}\n"
               f"Padrao: {int(val*100)}% | Formato: decimal (ex: 0,80)\n"
               "Intervalo valido: entre 0,50 e 0,99")
        vc.comment = Comment(tip, "Curva ABC Reggae")

        # Validacao de dados robusta
        f2 = "0.89" if row == 9 else "0.99"
        dv = DataValidation(
            type="decimal", operator="between", formula1="0.5", formula2=f2,
            error=f"Digite um decimal entre 0,50 e {f2} (ex: 0,80 para 80%)",
            errorTitle="Valor invalido",
            prompt="Formato decimal. Ex: 0,80 = 80%",
            promptTitle="Limite da Curva ABC",
            showErrorMessage=True, showInputMessage=True)
        ws.add_data_validation(dv); dv.add(vc)

    # ── Painel resumo A/B/C (cols K-N, rows 8-11) ─────────────────────────────
    mc(ws, 8, 11, 8, 14)
    ws.cell(row=8, column=11, value="RESUMO POR CLASSE").fill = F(BLK)
    ws.cell(row=8, column=11).font = Ft(bold=True, color=Y, size=10); ws.cell(row=8,column=11).alignment = Al()

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
        ws.row_dimensions[row].height = 24

    # ── Row 11: cabeçalho dos dados ───────────────────────────────────────────
    ws.row_dimensions[11].height = 26
    for col, label, bg, fg in [
        (1,"ABC",BLK,G),(2,"SKU",BLK,WHT),(3,"Cod.",BLK,WHT),
        (4,"Descricao do Produto",BLK,WHT),
        (5,"Qtd.",BLK,Y),(6,"Custo Unit.(R$)",BLK,Y),
        (7,"V.T. (R$)",BLK,G),(8,"% Individual",BLK,WHT),(9,"% Acumulado",BLK,WHT),
    ]:
        cx = ws.cell(row=11, column=col, value=label)
        cx.fill = F(bg); cx.font = Ft(bold=True, color=fg, size=10)
        cx.alignment = Al(wrap=(col==4)); cx.border = B(color=G_DARK)

    # ── Linhas de dados ───────────────────────────────────────────────────────
    # Linhas escritas em ordem DECRESCENTE de V.T. (sorted_data) — padrão Pareto.
    # SEM fills estáticos nas linhas; CF é a ÚNICA fonte de cor das linhas.
    unlock = Protection(locked=False)

    for i, d in enumerate(sorted_data):
        sku, desc, cod, qty, cost = d["sku"], d["desc"], d["cod"], d["qty"], d["cost"]
        row = DS + i

        # ── Col A: fórmula ABC (sem fill estático — CF cuida disso) ──────────
        cx = ws.cell(row=row, column=1,
                     value=f'=IF($I{row}<=$B$9,"A",IF($I{row}<=$B$10,"B","C"))')
        cx.font = Ft(bold=True, size=12); cx.alignment = Al(); cx.border = B(color=BLK)

        # ── Col B: SKU ────────────────────────────────────────────────────────
        cx = ws.cell(row=row, column=2, value=sku)
        cx.font = Ft(color=BLK, size=9); cx.alignment = Al(); cx.border = B(color=GBORDER)

        # ── Col C: Código ─────────────────────────────────────────────────────
        cx = ws.cell(row=row, column=3, value=cod)
        cx.font = Ft(color=BLK, size=9); cx.alignment = Al(); cx.border = B(color=GBORDER)

        # ── Col D: Descrição ──────────────────────────────────────────────────
        cx = ws.cell(row=row, column=4, value=desc)
        cx.font = Ft(color=BLK, size=10); cx.alignment = Al(h="left"); cx.border = B(color=GBORDER)

        # ── Col E: Qtd — EDITÁVEL — com comentário na primeira linha ──────────
        cx = ws.cell(row=row, column=5, value=qty)
        cx.font = Ft(bold=True, color=BLK, size=10); cx.alignment = Al()
        cx.number_format = '#,##0'; cx.protection = unlock
        cx.border = Border(
            left=Side(style="medium",color=G_DARK), right=Side(style="medium",color=G_DARK),
            top=Side(style="thin",  color=GMID),   bottom=Side(style="thin", color=GMID))
        if i == 0:  # comentário só na primeira linha
            cx.comment = Comment(
                "EDITAVEL — Quantidade do produto\n"
                "Altere este valor e pressione Enter.\n"
                "V.T. (col G) e a classificacao ABC atualizam automaticamente.",
                "Curva ABC Reggae")

        # ── Col F: Custo Unit — EDITÁVEL — com comentário na primeira linha ───
        cx = ws.cell(row=row, column=6, value=cost)
        cx.font = Ft(bold=True, color=BLK, size=10); cx.alignment = Al(h="right")
        cx.number_format = 'R$ #,##0.00'; cx.protection = unlock
        cx.border = Border(
            left=Side(style="medium",color=G_DARK), right=Side(style="medium",color=G_DARK),
            top=Side(style="thin",  color=GMID),   bottom=Side(style="thin", color=GMID))
        if i == 0:
            cx.comment = Comment(
                "EDITAVEL — Custo Unitario do produto\n"
                "Altere este valor e pressione Enter.\n"
                "V.T. (col G) e a classificacao ABC atualizam automaticamente.",
                "Curva ABC Reggae")

        # ── Col G: VT = E * F ─────────────────────────────────────────────────
        cx = ws.cell(row=row, column=7, value=f'=IFERROR(E{row}*F{row},0)')
        cx.font = Ft(bold=True, color=BLK, size=10)
        cx.alignment = Al(h="right"); cx.border = B(color=GBORDER); cx.number_format = 'R$ #,##0.00'

        # ── Col H: % Individual ───────────────────────────────────────────────
        cx = ws.cell(row=row, column=8,
                     value=f'=IFERROR(G{row}/SUM($G${DS}:$G${DE}),0)')
        cx.font = Ft(color=BLK, size=10)
        cx.alignment = Al(); cx.border = B(color=GBORDER); cx.number_format = '0.00%'

        # ── Col I: % Acumulado — soma cumulativa progressiva (tabela ordenada por VT) ─
        # =SUM($G$12:G12)/SUM($G$12:$G$21): para row 12 = G12/Total; row 13 = (G12+G13)/Total
        # Mais robusto que SUMIF: não falha com V.T. duplicados
        cx = ws.cell(row=row, column=9,
                     value=f'=IFERROR(SUM($G${DS}:G{row})/SUM($G${DS}:$G${DE}),0)')
        cx.font = Ft(color=BLK, size=10)
        cx.alignment = Al(); cx.border = B(color=GBORDER); cx.number_format = '0.00%'

    # ── Dados do gráfico (cols P-T, OCULTAS) ─────────────────────────────────
    # Pareto clássico: V.T. por produto (3 séries A/B/C) + % acumulado
    for i, d in enumerate(sorted_data):
        row = DS + i
        desc_s = d["desc"][:20] if len(d["desc"])>20 else d["desc"]
        ws.cell(row=row, column=16, value=f"{d['abc']} | {d['sku']} | {desc_s}")
        ws.cell(row=row, column=17, value=d["vt"] if d["abc"]=="A" else 0)
        ws.cell(row=row, column=18, value=d["vt"] if d["abc"]=="B" else 0)
        ws.cell(row=row, column=19, value=d["vt"] if d["abc"]=="C" else 0)
        ws.cell(row=row, column=20, value=round(d["pct_acc"], 5))

    # ── Linha TOTAL ───────────────────────────────────────────────────────────
    tr = DE + 1
    ws.row_dimensions[tr].height = 24
    mc(ws, tr, 1, tr, 4)
    cx = ws.cell(row=tr, column=1, value="TOTAL GERAL")
    cx.fill = F(BLK); cx.font = Ft(bold=True, color=Y, size=12); cx.alignment = Al()
    ws.cell(row=tr,column=5, value=f'=SUM(E{DS}:E{DE})'  ).font = Ft(bold=True, color=WHT,size=12)
    ws.cell(row=tr,column=5).fill=F(BLK); ws.cell(row=tr,column=5).alignment=Al()
    ws.cell(row=tr,column=5).number_format='#,##0'
    ws.cell(row=tr,column=6).fill=F(BLK)
    cx=ws.cell(row=tr,column=7,value=f'=SUM(G{DS}:G{DE})')
    cx.fill=F(G); cx.font=Ft(bold=True,size=12); cx.alignment=Al(h="right"); cx.number_format='R$ #,##0.00'
    for col in [8,9]:
        cx=ws.cell(row=tr,column=col,value=1.0)
        cx.fill=F(BLK); cx.font=Ft(bold=True,size=12); cx.alignment=Al(); cx.number_format='0%'
    for col in range(1,10):
        ws.cell(row=tr,column=col).border=Border(
            top=Side(style="medium",color=G), bottom=Side(style="medium",color=G),
            left=Side(style="thin",color=G),  right=Side(style="thin",color=G))

    # ── FORMATAÇÃO CONDICIONAL ────────────────────────────────────────────────
    # SOLUÇÃO DEFINITIVA: células sem fill estático = CF é a única fonte de cor
    # Aplicamos em duas camadas:
    #   1. Linha toda (A:I) com cor clara (background reggae)
    #   2. Col A (badge ABC) com cor sólida + texto branco em negrito
    rng_row   = f"A{DS}:I{DE}"   # linha toda — cor de fundo suave
    rng_badge = f"A{DS}:A{DE}"   # só col A — badge com cor forte

    for cls, lite, dark, font_clr in [("A",G_LITE,G,WHT),("B",Y_LITE,Y_DARK,BLK),("C",R_LITE,R_DARK,WHT)]:
        # FF prefix = fully opaque alpha in dxf — critical for CF fills to render
        lite_ff = "FF"+lite if len(lite)==6 else lite
        dark_ff = "FF"+dark if len(dark)==6 else dark
        # 1. Fundo suave na linha toda
        ws.conditional_formatting.add(rng_row, FormulaRule(
            formula=[f'$A{DS}="{cls}"'],
            fill=PatternFill(fill_type="solid", fgColor=lite_ff),
            font=Font(color=BLK)))
        # 2. Badge escuro na col A (prioridade maior — adicionado depois)
        ws.conditional_formatting.add(rng_badge, FormulaRule(
            formula=[f'$A{DS}="{cls}"'],
            fill=PatternFill(fill_type="solid", fgColor=dark_ff),
            font=Font(color=font_clr, bold=True, size=12)))

    # ── GRÁFICO — Pareto Clássico ─────────────────────────────────────────────
    # Barras de V.T. por produto (A=verde, B=dourado, C=vermelho) + curva % acumulado
    chart_row = tr + 3
    cats = Reference(ws, min_col=16, min_row=DS, max_row=DE)

    # ─ BarChart: V.T. por produto colorido por classe ─
    bar = BarChart()
    bar.type     = "col"
    bar.grouping = "stacked"
    bar.title    = f"{cfg['title']} — Pareto: V.T. por Produto & Curva % Acumulado"
    bar.style    = 10
    bar.width    = 34
    bar.height   = 16
    bar.y_axis.title  = "Valor Total (R$)"
    bar.y_axis.numFmt = 'R$ #,##0'
    bar.x_axis.title  = "Produtos ordenados por V.T. decrescente (nome + classe no rotulo)"

    for col_off, (cls_lbl, clr, clr_line) in enumerate([
        ("Classe A", G,      G_DARK),
        ("Classe B", Y,      Y_DARK),
        ("Classe C", R,      R_DARK),
    ]):
        ser = Series(Reference(ws, min_col=17+col_off, min_row=DS, max_row=DE), title=cls_lbl)
        ser.graphicalProperties.solidFill      = clr
        ser.graphicalProperties.line.solidFill = clr_line
        ser.graphicalProperties.line.width     = 6000
        bar.series.append(ser)

    bar.set_categories(cats)

    # Data labels nas barras
    bar.dLbls            = DataLabelList()
    bar.dLbls.showVal    = True
    bar.dLbls.showSerName = False
    bar.dLbls.showCatName = False
    bar.dLbls.numFmt     = 'R$ #,##0'

    # ─ LineChart: % acumulado no EIXO SECUNDÁRIO (preto, curva clássica Pareto) ─
    line = LineChart()
    line.y_axis.axId    = 200          # eixo Y secundário
    line.y_axis.axPos   = "r"          # DIREITA — sem isso a linha fica sobre o eixo primário
    line.y_axis.crosses = "max"        # cruza no máximo do eixo primário
    line.y_axis.scaling.min = 0
    line.y_axis.scaling.max = 1
    line.y_axis.numFmt  = '0%'
    line.y_axis.title   = "% Acumulado (Curva Pareto)"

    ser_l = Series(Reference(ws, min_col=20, min_row=DS, max_row=DE), title="% Acumulado")
    ser_l.graphicalProperties.line.solidFill = LINE_CLR
    ser_l.graphicalProperties.line.width     = 28000
    ser_l.smooth = True
    ser_l.marker.symbol  = "circle"
    ser_l.marker.size    = 5
    ser_l.marker.graphicalProperties.solidFill      = BLK
    ser_l.marker.graphicalProperties.line.solidFill = WHT
    line.series.append(ser_l)
    line.set_categories(cats)

    bar += line
    ws.add_chart(bar, f"A{chart_row}")

    # ── Proteção da planilha ──────────────────────────────────────────────────
    ws.protection.sheet    = True
    ws.protection.password = "abc2026."
    ws.protection.enable()
    ws.protection.selectLockedCells   = False
    ws.protection.selectUnlockedCells = False
    for r in range(DS, DE+1):
        ws.cell(row=r, column=5).protection = unlock
        ws.cell(row=r, column=6).protection = unlock
    ws.cell(row=9,  column=2).protection = unlock
    ws.cell(row=10, column=2).protection = unlock

    ws.freeze_panes = f"A{DS}"
    ws.sheet_properties.tabColor = cfg["tab_color"]
    return ws


# ═══════════════════════════════════════════════════════════════════════════════
#  CALENDÁRIO 2026
# ═══════════════════════════════════════════════════════════════════════════════
def create_calendar_sheet(wb, year=2026):
    ws = wb.create_sheet(title="📅 CALENDARIO")

    MONTHS_PT = ["Janeiro","Fevereiro","Marco","Abril","Maio","Junho",
                 "Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"]
    DAYS_PT   = ["Seg","Ter","Qua","Qui","Sex","Sab","Dom"]
    MONTH_ROWS = 11
    MONTH_CLR  = [G_DARK,G_DARK,G_DARK, Y_DARK,Y_DARK,Y_DARK,
                  R_DARK,R_DARK,R_DARK, G_DARK,Y_DARK,R_DARK]

    for ci in range(1, 34):
        cl = get_column_letter(ci)
        ws.column_dimensions[cl].width = 1.2 if (ci-1)%8==7 else 4.8

    for r in range(1, 48):
        ws.row_dimensions[r].height = 16
        for c in range(1, 34):
            ws.cell(row=r, column=c).fill = F(DARK)

    # Faixa reggae topo
    ws.row_dimensions[1].height = 8
    for ci in range(1,34): ws.cell(row=1,column=ci).fill=F([G,Y,R,BLK][(ci-1)%4])

    # Título
    ws.row_dimensions[2].height = 52
    mc(ws, 2, 1, 2, 32)
    t = ws.cell(row=2, column=1, value=f"CALENDARIO {year}  —  One Love, One Excel!  |  Curva ABC Reggae")
    t.fill=F(BLK); t.font=Font(bold=True,color=G,size=22,name="Calibri"); t.alignment=Al()

    ws.row_dimensions[3].height = 8
    for ci in range(1,34): ws.cell(row=3,column=ci).fill=F([R,Y,G,BLK][(ci-1)%4])

    # Legenda de revisão
    ws.row_dimensions[4].height = 18
    mc(ws, 4, 1, 4, 32)
    leg = ws.cell(row=4, column=1,
        value="  Classe A = revisar diariamente   |   "
              "Classe B = revisar semanalmente   |   "
              "Classe C = revisar mensalmente   |   "
              "Dias marcados = sugestao de inventario rotativo")
    leg.fill=F("0D2010"); leg.font=Ft(italic=True,color=G_LITE,size=9); leg.alignment=Al(h="left")

    ws.row_dimensions[5].height = 8
    START_ROW = 6

    for month in range(1, 13):
        ri  = (month-1) // 4
        ci_ = (month-1) % 4
        mr  = START_ROW + ri * MONTH_ROWS
        mc_ = 1 + ci_ * 8

        ws.row_dimensions[mr].height = 22
        mc(ws, mr, mc_, mr, mc_+6)
        mh = ws.cell(row=mr, column=mc_, value=f"{MONTHS_PT[month-1].upper()} {year}")
        mh.fill=F(MONTH_CLR[month-1]); mh.font=Ft(bold=True,size=10); mh.alignment=Al()
        mh.border=Border(bottom=Side(style="medium",color=WHT))

        ws.row_dimensions[mr+1].height = 16
        for d_idx, day_name in enumerate(DAYS_PT):
            cx=ws.cell(row=mr+1, column=mc_+d_idx, value=day_name)
            cx.fill=F(R_DARK if d_idx>=5 else G_DARK)
            cx.font=Ft(bold=True,size=8); cx.alignment=Al()
            cx.border=B(style="thin",color=DARK)

        weeks = cal_module.monthcalendar(year, month)
        for w_idx, week in enumerate(weeks):
            ws.row_dimensions[mr+2+w_idx].height = 15
            for d_idx, day_num in enumerate(week):
                cx=ws.cell(row=mr+2+w_idx, column=mc_+d_idx)
                if day_num==0:
                    cx.fill=F("080808"); cx.value=""
                else:
                    is_we      = d_idx >= 5
                    is_review  = day_num in [1, 8, 15, 22]
                    if is_we:
                        bg, fg = "1A0000", R_LITE
                    elif is_review:
                        bg, fg = G_DARK, Y_LITE
                    else:
                        bg, fg = "0D1A0D", G_LITE
                    cx.fill=F(bg); cx.font=Ft(color=fg,size=9,bold=is_review)
                    cx.value=day_num; cx.alignment=Al()
                    cx.border=B(style="hair",color="111111")

        for w_off in range(len(weeks), 6):
            ws.row_dimensions[mr+2+w_off].height = 15
            for d_idx in range(7):
                ws.cell(row=mr+2+w_off,column=mc_+d_idx).fill=F("080808")

    # footer
    footer = START_ROW + 3*MONTH_ROWS
    ws.row_dimensions[footer].height = 20
    mc(ws, footer, 1, footer, 32)
    f = ws.cell(row=footer, column=1,
        value="  Dias marcados (1, 8, 15, 22) = sugestao de revisao/inventario  "
              "|  Fim de semana em vermelho  "
              "|  Jah bless your schedule — One Love, One Excel! 2026")
    f.fill=F(BLK); f.font=Ft(bold=True,color=G,size=10); f.alignment=Al(h="left")

    ws.sheet_properties.tabColor = Y_DARK
    return ws


# ── MAIN ───────────────────────────────────────────────────────────────────────
def main():
    OUTPUT = "/home/user/Zyth/curva_abc_reggae.xlsx"
    wb = Workbook()
    wb.remove(wb.active)

    for cfg in SHEETS:
        print(f"  {cfg['name']}  ({len(cfg['items'])} itens)...")
        create_data_sheet(wb, cfg)

    print("  Calendario 2026...")
    create_calendar_sheet(wb, 2026)

    wb.properties.title   = "Curva ABC Reggae — Super Version"
    wb.properties.subject = "Dashboard ABC Pareto | 3 abas + Calendario | Reggae"
    wb.properties.creator = "Curva ABC Reggae"
    wb.save(OUTPUT)

    import os
    size = os.path.getsize(OUTPUT)
    print(f"\n  Salvo: {OUTPUT}")
    print(f"  {size/1024:.1f} KB  |  {len(wb.sheetnames)} abas")

if __name__ == "__main__":
    main()
