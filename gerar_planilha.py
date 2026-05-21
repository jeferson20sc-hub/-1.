"""Gera Gestao_Estoque_LIS.xlsx: 3 abas com gráficos automáticos.
- Lancamentos: Tabela 'Lancamentos' (autofilter) p/ o Power Automate gravar.
- Dashboard: KPIs (fórmulas) + gráfico de barras de entradas/saidas dos últimos 14 dias
  e gráfico de pizza com saídas por item.
- Itens: catálogo dos 10 produtos base com unidade e perecibilidade."""
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList

COLS = [
    ("ID", 22), ("DataHora", 22), ("Data", 20), ("Operador", 16),
    ("Codigo", 10), ("Descricao", 40), ("Tipo", 12), ("Quantidade", 12),
    ("Saldo", 12), ("EstoqueSeg", 13), ("PontoPedido", 13),
    ("Status", 18), ("Origem", 10),
]
EXEMPLO = [
    "LIS-EXEMPLO-001", "2026-05-19T12:00:00.000Z", "19/05/2026 12:00:00",
    "Jeferson", "001", "Filtro de óleo (loja de peças)", "Entrada", 10,
    130, 40, 140, "🟢 OK", "Site",
]
ITENS = [
    ("001", "Filtro de óleo (loja de peças)", "un", "Não", 40, 140),
    ("002", "Arroz (restaurante)", "kg", "Sim", 10, 55),
    ("003", "Chapas de MDF (fábrica de móveis)", "un", "Não", 100, 300),
    ("004", "Polivitamínico (farmácia)", "un", "Sim", 30, 100),
    ("005", "Resmas de papel A4 (papelaria)", "un", "Não", 150, 350),
    ("006", "Essência cosmética (cosméticos)", "L", "Sim", 50, 150),
    ("007", "Componente crítico (peças críticas)", "un", "Não", 150, 450),
    ("008", "Laptops (varejo eletrônicos)", "un", "Não", 50, 200),
    ("009", "Kits de luvas hospitalar (NR-32)", "cx", "Sim", 2000, 7000),
    ("010", "Insumo industrial", "un", "Não", 200, 600),
]

azul = PatternFill("solid", fgColor="00529B")
verde = PatternFill("solid", fgColor="107C41")
cinza = PatternFill("solid", fgColor="F1F5F9")
branco = Font(color="FFFFFF", bold=True, size=11)
titulo = Font(color="00529B", bold=True, size=14)
borda = Border(*[Side(style="thin", color="D9E2EC")] * 4)

wb = Workbook()

# === ABA 1: Lancamentos (Tabela para o Power Automate) ===
ws = wb.active
ws.title = "Lancamentos"
ws.sheet_view.showGridLines = False
ws.append([c[0] for c in COLS])
ws.append(EXEMPLO)
for ci, (nome, larg) in enumerate(COLS, 1):
    L = get_column_letter(ci)
    ws.column_dimensions[L].width = larg
    h = ws.cell(row=1, column=ci)
    h.fill = azul; h.font = branco
    h.alignment = Alignment(horizontal="center", vertical="center")
    h.border = borda
    d = ws.cell(row=2, column=ci)
    d.border = borda
    d.alignment = Alignment(horizontal="center" if nome != "Descricao" else "left")
ws.row_dimensions[1].height = 26
ws.freeze_panes = "A2"
ref = f"A1:{get_column_letter(len(COLS))}2"
tab = Table(displayName="Lancamentos", ref=ref)
tab.tableStyleInfo = TableStyleInfo(
    name="TableStyleMedium2", showFirstColumn=False,
    showLastColumn=False, showRowStripes=True, showColumnStripes=False,
)
ws.add_table(tab)
ws.auto_filter.ref = ref

# === ABA 2: Itens (catálogo) ===
wi = wb.create_sheet("Itens")
wi.sheet_view.showGridLines = False
itens_cols = [("Codigo", 10), ("Descricao", 40), ("Unidade", 12),
              ("VenceRapido", 13), ("EstoqueSeg", 12), ("PontoPedido", 13)]
wi.append([c[0] for c in itens_cols])
for it in ITENS:
    wi.append(list(it))
for ci, (nome, larg) in enumerate(itens_cols, 1):
    L = get_column_letter(ci)
    wi.column_dimensions[L].width = larg
    h = wi.cell(row=1, column=ci)
    h.fill = verde; h.font = branco
    h.alignment = Alignment(horizontal="center"); h.border = borda
wi.row_dimensions[1].height = 24
wi.freeze_panes = "A2"
itab = Table(displayName="Itens", ref=f"A1:F{1+len(ITENS)}")
itab.tableStyleInfo = TableStyleInfo(
    name="TableStyleMedium4", showRowStripes=True,
)
wi.add_table(itab)
wi.auto_filter.ref = f"A1:F{1+len(ITENS)}"

# === ABA 3: Dashboard ===
wd = wb.create_sheet("Dashboard")
wd.sheet_view.showGridLines = False
wd["A1"] = "Dashboard — Gestão de Estoque LIS"
wd["A1"].font = titulo
wd.merge_cells("A1:F1")

kpis = [
    ("Total de movimentações", "=COUNTA(Lancamentos!E:E)-1"),
    ("Quantidade total ENTRADAS", "=SUMIFS(Lancamentos!H:H,Lancamentos!G:G,\"Entrada\")"),
    ("Quantidade total SAÍDAS", "=SUMIFS(Lancamentos!H:H,Lancamentos!G:G,\"Saída\")"),
    ("Operadores únicos", "=SUMPRODUCT((Lancamentos!D2:D5000<>\"\")/COUNTIF(Lancamentos!D2:D5000,Lancamentos!D2:D5000&\"\"))"),
    ("Itens em status CRÍTICO 🔴", "=COUNTIF(Lancamentos!L:L,\"*CRÍTICO*\")"),
    ("Itens em status PEDIR 🟡", "=COUNTIF(Lancamentos!L:L,\"*PEDIR*\")"),
]
for i, (rot, frm) in enumerate(kpis):
    r = 3 + i
    wd.cell(row=r, column=1, value=rot).font = Font(bold=True, color="00529B")
    wd.cell(row=r, column=2, value=frm)
    wd.cell(row=r, column=2).font = Font(bold=True, size=14, color="107C41")
    wd.cell(row=r, column=1).fill = cinza
    wd.cell(row=r, column=2).fill = cinza
wd.column_dimensions["A"].width = 34
wd.column_dimensions["B"].width = 22

# Helper area: últimos 14 dias (linhas 3..16, colunas F..H)
wd["F2"] = "Data"; wd["G2"] = "Entradas"; wd["H2"] = "Saídas"
for c in ("F2", "G2", "H2"):
    wd[c].font = branco; wd[c].fill = azul; wd[c].alignment = Alignment(horizontal="center")
for i in range(14):
    r = 3 + i
    wd.cell(row=r, column=6, value=f"=TODAY()-{13-i}").number_format = "dd/mm/yyyy"
    wd.cell(row=r, column=7,
            value=f'=SUMPRODUCT((LEFT(Lancamentos!C$2:C$5000,10)=TEXT(F{r},"dd/mm/yyyy"))*(Lancamentos!G$2:G$5000="Entrada")*Lancamentos!H$2:H$5000)')
    wd.cell(row=r, column=8,
            value=f'=SUMPRODUCT((LEFT(Lancamentos!C$2:C$5000,10)=TEXT(F{r},"dd/mm/yyyy"))*(Lancamentos!G$2:G$5000="Saída")*Lancamentos!H$2:H$5000)')
for col in ("F", "G", "H"):
    wd.column_dimensions[col].width = 14

# Helper: saídas por item (linhas 3..12, colunas J..K) para gráfico de pizza
wd["J2"] = "Item"; wd["K2"] = "Saídas"
for c in ("J2", "K2"):
    wd[c].font = branco; wd[c].fill = azul; wd[c].alignment = Alignment(horizontal="center")
for i, it in enumerate(ITENS):
    r = 3 + i
    wd.cell(row=r, column=10, value=it[0])
    wd.cell(row=r, column=11,
            value=f'=SUMIFS(Lancamentos!H:H,Lancamentos!E:E,J{r},Lancamentos!G:G,"Saída")')
wd.column_dimensions["J"].width = 10
wd.column_dimensions["K"].width = 12

# Gráfico de barras: entradas vs saídas por dia
bar = BarChart()
bar.type = "col"; bar.style = 11; bar.grouping = "clustered"
bar.title = "Entradas vs Saídas — últimos 14 dias"
bar.y_axis.title = "Quantidade"; bar.x_axis.title = "Dia"
data = Reference(wd, min_col=7, max_col=8, min_row=2, max_row=16)
cats = Reference(wd, min_col=6, min_row=3, max_row=16)
bar.add_data(data, titles_from_data=True)
bar.set_categories(cats)
bar.height = 9; bar.width = 20
wd.add_chart(bar, "A12")

# Gráfico de pizza: saídas por item
pie = PieChart()
pie.title = "Saídas por item (catálogo base)"
pdata = Reference(wd, min_col=11, min_row=2, max_row=12)
pcats = Reference(wd, min_col=10, min_row=3, max_row=12)
pie.add_data(pdata, titles_from_data=True)
pie.set_categories(pcats)
pie.dataLabels = DataLabelList(showPercent=True)
pie.height = 9; pie.width = 14
wd.add_chart(pie, "A32")

wb.save("/home/user/senai-cba/Gestao_Estoque_LIS.xlsx")
print("OK: workbook gerado com abas:", wb.sheetnames)
