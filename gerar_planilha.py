"""Gera Gestao_Estoque_LIS.xlsx: 1 aba 'Lancamentos' com Tabela + filtros,
pronta para o Power Automate (acao 'Adicionar uma linha a uma tabela')."""
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

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

wb = Workbook()
ws = wb.active
ws.title = "Lancamentos"
ws.sheet_view.showGridLines = False

azul = PatternFill("solid", fgColor="00529B")
branco = Font(color="FFFFFF", bold=True, size=11)
borda = Border(*[Side(style="thin", color="D9E2EC")] * 4)

headers = [c[0] for c in COLS]
ws.append(headers)
ws.append(EXEMPLO)

for ci, (nome, larg) in enumerate(COLS, 1):
    L = get_column_letter(ci)
    ws.column_dimensions[L].width = larg
    h = ws.cell(row=1, column=ci)
    h.fill = azul
    h.font = branco
    h.alignment = Alignment(horizontal="center", vertical="center")
    h.border = borda
    d = ws.cell(row=2, column=ci)
    d.border = borda
    d.alignment = Alignment(horizontal="center" if nome not in ("Descricao",) else "left")

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

wb.save("/home/user/senai-cba/Gestao_Estoque_LIS.xlsx")
print("OK: Gestao_Estoque_LIS.xlsx -> aba 'Lancamentos', tabela 'Lancamentos'", ref)
