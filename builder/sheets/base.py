"""BASE sheet: tbEventos master event log."""

from ..data import gerar_eventos_amostra
from ..theme import FMT_BRL, FMT_DATA, FMT_HORAS
from ..util import (
    freeze_header,
    hide_gridlines,
    page_setup_landscape,
    section_header,
    set_column_widths,
    title_bar,
    write_table,
)


HEADERS = [
    "ID", "Data", "Hora", "Operador", "Forno", "TipoEvento",
    "Categoria", "Componente", "Criticidade", "DuracaoHoras",
    "CustoEstimado", "Observacao", "Status",
]


def build(wb):
    ws = wb.create_sheet("BASE")
    hide_gridlines(ws)
    set_column_widths(ws, {
        "A": 2, "B": 6, "C": 12, "D": 8, "E": 18, "F": 8, "G": 18,
        "H": 14, "I": 22, "J": 12, "K": 14, "L": 16, "M": 32, "N": 14,
    })
    title_bar(
        ws,
        "EXAUSTAO 360 - BASE DE EVENTOS",
        "tbEventos - base mestra de registros operacionais. Nao editar manualmente; usar a aba REGISTRO.",
        last_col="N",
    )
    section_header(ws, "B4:N4", "tbEventos - Registros Operacionais")

    eventos = gerar_eventos_amostra(60)
    ref = write_table(
        ws, "B5",
        HEADERS,
        eventos,
        "tbEventos",
        style="TableStyleMedium2",
    )

    # Format Data, Hora, Duracao, Custo columns inside table body
    n = len(eventos)
    for r in range(6, 6 + n):
        ws.cell(row=r, column=3).number_format = FMT_DATA          # Data (col C)
        ws.cell(row=r, column=4).number_format = "hh:mm"          # Hora (col D)
        ws.cell(row=r, column=11).number_format = FMT_HORAS       # Duracao (col K)
        ws.cell(row=r, column=12).number_format = FMT_BRL         # Custo (col L)

    freeze_header(ws, "A6")
    page_setup_landscape(ws)
    return ws
