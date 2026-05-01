"""CONFIG sheet: business parameters + reference tables."""

from openpyxl.styles import Font

from ..data import (
    CATEGORIAS,
    COMPONENTES,
    CRITICIDADES,
    FORNOS,
    OPERADORES,
    PARAMETROS,
    REGRAS_RISCO,
    TIPOS_EVENTO,
    TURNOS,
)
from ..theme import (
    BORDER_ALL_THIN,
    BRANCO,
    CINZA_BORDA,
    CINZA_CLARO,
    FMT_BRL,
    FMT_INT,
    FMT_PCT,
    GRAFITE,
    PETROLEO,
    fill,
)
from ..util import (
    freeze_header,
    hide_gridlines,
    page_setup_landscape,
    section_header,
    set_column_widths,
    set_row_heights,
    title_bar,
    write_table,
)


def build(wb):
    ws = wb.create_sheet("CONFIG")
    hide_gridlines(ws)
    set_column_widths(ws, {
        "A": 2, "B": 28, "C": 22, "D": 22, "E": 22,
        "F": 22, "G": 22, "H": 22, "I": 22, "J": 22, "K": 22,
    })
    title_bar(
        ws,
        "EXAUSTAO 360 - CONFIGURACOES DO SISTEMA",
        "Parametros de negocio, fornos, componentes e tabelas auxiliares - editavel apenas em modo Supervisor",
        last_col="K",
    )

    # Parametros (tbParametros)
    section_header(ws, "B4:D4", "Parametros de Negocio (tbParametros)")
    write_table(
        ws, "B5",
        ["Parametro", "Valor", "Descricao"],
        [(p, v, d) for p, v, d in PARAMETROS],
        "tbParametros",
        style="TableStyleMedium2",
    )
    # Format the value column for currency / pct rows when applicable
    fmt_map = {
        "CustoHoraParada": FMT_BRL,
        "MetaDisponibilidade": FMT_PCT,
        "TaxaDescontoAnual": FMT_PCT,
        "PercentualReducaoEsperada": FMT_PCT,
        "CustoSolucaoPro": FMT_BRL,
        "HorasOperacionalDia": FMT_INT,
        "DiasOperacionalAno": FMT_INT,
        "MetaMTBF": FMT_INT,
        "MetaMTTR": FMT_INT,
        "LimiteCriticidadeAlta": FMT_INT,
    }
    for i, (param, _, _) in enumerate(PARAMETROS, start=6):
        if param in fmt_map:
            ws.cell(row=i, column=3).number_format = fmt_map[param]

    # Fornos
    section_header(ws, "F4:J4", "Fornos (tbFornos)")
    write_table(
        ws, "F5",
        ["ID", "Nome", "Area", "Capacidade (kg/h)", "Status"],
        FORNOS,
        "tbFornos",
        style="TableStyleMedium3",
    )

    # Tipos de Evento
    last_param_row = 5 + len(PARAMETROS) + 2
    section_header(ws, f"B{last_param_row}:D{last_param_row}", "Tipos de Evento (tbTiposEvento)")
    write_table(
        ws, f"B{last_param_row + 1}",
        ["Tipo", "Categoria", "ImpactoBase"],
        TIPOS_EVENTO,
        "tbTiposEvento",
        style="TableStyleMedium4",
    )

    # Categorias
    last_forno_row = 5 + len(FORNOS) + 2
    section_header(ws, f"F{last_forno_row}:G{last_forno_row}", "Categorias (tbCategorias)")
    write_table(
        ws, f"F{last_forno_row + 1}",
        ["Categoria", "Descricao"],
        CATEGORIAS,
        "tbCategorias",
        style="TableStyleMedium5",
    )

    # Componentes
    next_row_b = last_param_row + 1 + len(TIPOS_EVENTO) + 3
    section_header(ws, f"B{next_row_b}:E{next_row_b}", "Componentes (tbComponentes)")
    write_table(
        ws, f"B{next_row_b + 1}",
        ["ID", "Nome", "Tipo", "VidaUtilMeses"],
        COMPONENTES,
        "tbComponentes",
        style="TableStyleMedium6",
    )

    # Operadores
    next_row_f = last_forno_row + 1 + len(CATEGORIAS) + 3
    section_header(ws, f"F{next_row_f}:I{next_row_f}", "Operadores (tbOperadores)")
    write_table(
        ws, f"F{next_row_f + 1}",
        ["ID", "Nome", "Turno", "Area"],
        OPERADORES,
        "tbOperadores",
        style="TableStyleMedium7",
    )

    # Turnos
    next_row_b2 = next_row_b + 1 + len(COMPONENTES) + 3
    section_header(ws, f"B{next_row_b2}:D{next_row_b2}", "Turnos (tbTurnos)")
    write_table(
        ws, f"B{next_row_b2 + 1}",
        ["Turno", "Inicio", "Fim"],
        TURNOS,
        "tbTurnos",
        style="TableStyleMedium8",
    )

    # Criticidades
    next_row_f2 = next_row_f + 1 + len(OPERADORES) + 3
    section_header(ws, f"F{next_row_f2}:H{next_row_f2}", "Criticidades (tbCriticidades)")
    write_table(
        ws, f"F{next_row_f2 + 1}",
        ["Nivel", "Pontuacao", "Descricao"],
        CRITICIDADES,
        "tbCriticidades",
        style="TableStyleMedium9",
    )

    # Regras de Risco
    next_row_b3 = next_row_b2 + 1 + len(TURNOS) + 3
    section_header(ws, f"B{next_row_b3}:E{next_row_b3}", "Regras de Risco (tbRegrasRisco)")
    write_table(
        ws, f"B{next_row_b3 + 1}",
        ["Regra", "Condicao", "NivelRisco", "AcaoRecomendada"],
        REGRAS_RISCO,
        "tbRegrasRisco",
        style="TableStyleMedium10",
    )

    freeze_header(ws, "A4")
    page_setup_landscape(ws)
    return ws
