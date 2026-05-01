"""LOG sheet (very hidden) - audit log of system actions."""

from ..util import set_column_widths, write_table


HEADERS = ["DataHora", "Usuario", "Acao", "Aba", "Registro", "Resultado"]


def build(wb):
    ws = wb.create_sheet("LOG")
    set_column_widths(ws, {"A": 22, "B": 20, "C": 28, "D": 18, "E": 16, "F": 28})

    ws["A1"] = "EXAUSTAO 360 - LOG DE AUDITORIA"
    ws["A1"].style = "exa_title"
    ws.row_dimensions[1].height = 28

    write_table(
        ws, "A3",
        HEADERS,
        [],  # empty - VBA appends
        "tbLogEdicao",
        style="TableStyleMedium11",
    )

    # Hide from end users
    ws.sheet_state = "veryHidden"
    return ws
