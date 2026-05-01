"""EXAUSTAO 360 ENTERPRISE PRO - workbook builder entry point.

Usage:
    python build.py [output_path]

Generates `dist/EXAUSTAO_360_ENTERPRISE_PRO.xlsx` with:
    - 7 user-facing sheets (DASHBOARD, REGISTRO, BASE, ANALISE, SIMULADOR,
      MODELO_COMERCIAL, CONFIG) + LOG (very hidden).
    - Structured tables (tbEventos, tbParametros, tbFornos, tbFMEA, tbPlanoAcao, ...).
    - Charts (Pareto, trend, ranking, MTBF/MTTR, scenarios).
    - Defined names used by formulas and VBA.

To produce the final `.xlsm` (with VBA macros), run `assemble.ps1` on a
Windows machine with Excel installed. That script imports the modules from
`vba/` into the workbook's VBA project and saves it as `.xlsm`.
"""

from __future__ import annotations

import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.workbook.defined_name import DefinedName

from builder.sheets import analise, base, comercial, config, dashboard, log, registro, simulador
from builder.sheets.simulador import INPUT_CELLS as SIM_INPUTS
from builder.theme import register_named_styles


OUTPUT_DEFAULT = Path("dist") / "EXAUSTAO_360_ENTERPRISE_PRO.xlsx"


def register_defined_names(wb) -> None:
    """Register all global defined names referenced by formulas and VBA."""

    # Constants and parameter lookups
    name_formulas: dict[str, str] = {
        # Constant: ~4 months of 24/7 operation, gives stable MTBF baseline
        "K_HORAS_PERIODO": "24*120",

        # Look up business parameters from tbParametros
        "K_META_DISPONIBILIDADE":
            'INDEX(tbParametros[Valor],MATCH("MetaDisponibilidade",tbParametros[Parametro],0))',
        "K_META_MTBF":
            'INDEX(tbParametros[Valor],MATCH("MetaMTBF",tbParametros[Parametro],0))',
        "K_META_MTTR":
            'INDEX(tbParametros[Valor],MATCH("MetaMTTR",tbParametros[Parametro],0))',
        "K_CUSTO_HORA":
            'INDEX(tbParametros[Valor],MATCH("CustoHoraParada",tbParametros[Parametro],0))',
        "K_PCT_REDUCAO":
            'INDEX(tbParametros[Valor],MATCH("PercentualReducaoEsperada",tbParametros[Parametro],0))',
        "K_CUSTO_SOLUCAO":
            'INDEX(tbParametros[Valor],MATCH("CustoSolucaoPro",tbParametros[Parametro],0))',
        "K_LIMIAR_CUSTO": "500000",

        # Aggregated KPIs used on DASHBOARD
        "ECONOMIA_ANUAL":
            "SUMPRODUCT(tbEventos[CustoEstimado])*K_PCT_REDUCAO*3",
        "ROI_ESTIMADO":
            "(SUMPRODUCT(tbEventos[CustoEstimado])*K_PCT_REDUCAO*3-K_CUSTO_SOLUCAO)"
            "/MAX(K_CUSTO_SOLUCAO,1)",
        "VALOR_PROTEGIDO":
            "SUMPRODUCT(tbEventos[CustoEstimado])*0.5",
        "SEMAFORO":
            'IF(SUMPRODUCT(tbEventos[CustoEstimado])>K_LIMIAR_CUSTO,"VERMELHO",'
            'IF(SUMPRODUCT(tbEventos[CustoEstimado])>K_LIMIAR_CUSTO/2,"AMARELO","VERDE"))',
    }

    for name, formula in name_formulas.items():
        dn = DefinedName(name=name, attr_text=formula)
        wb.defined_names[name] = dn

    # SIMULADOR INPUT_* cells (sheet-qualified absolute references)
    for name, cell in SIM_INPUTS.items():
        dn = DefinedName(name=name, attr_text=f"SIMULADOR!${cell[0]}${cell[1:]}")
        wb.defined_names[name] = dn


def main(argv: list[str]) -> int:
    out_path = Path(argv[1]) if len(argv) > 1 else OUTPUT_DEFAULT
    out_path.parent.mkdir(parents=True, exist_ok=True)

    wb = Workbook()
    # Remove default Sheet
    wb.remove(wb.active)

    register_named_styles(wb)

    # Build in dependency order: CONFIG and BASE first (defined names depend
    # on tables), then everything else.
    config.build(wb)
    base.build(wb)
    registro.build(wb)
    dashboard.build(wb)
    analise.build(wb)
    simulador.build(wb)
    comercial.build(wb)
    log.build(wb)

    register_defined_names(wb)

    # Reorder sheets: DASHBOARD first, CONFIG last among visible, LOG hidden
    desired = ["DASHBOARD", "REGISTRO", "BASE", "ANALISE", "SIMULADOR",
               "MODELO_COMERCIAL", "CONFIG", "LOG"]
    wb._sheets = [wb[name] for name in desired]

    # Workbook-level properties
    wb.properties.title = "EXAUSTAO 360 ENTERPRISE PRO"
    wb.properties.subject = "Confiabilidade Industrial - Exaustao"
    wb.properties.creator = "EXAUSTAO 360 PRO"
    wb.properties.description = (
        "Solucao B2B de gestao de confiabilidade para sistemas industriais "
        "de exaustao - dashboard executivo, analise FMEA, simulador de ROI."
    )
    wb.properties.keywords = "confiabilidade,manutencao,FMEA,MTBF,MTTR,ROI,industrial"

    wb.save(out_path)
    print(f"[OK] Workbook saved to: {out_path}")
    print(f"[INFO] Sheets: {[s.title for s in wb.worksheets if s.sheet_state != 'veryHidden']}")
    print(f"[INFO] Hidden:  {[s.title for s in wb.worksheets if s.sheet_state == 'veryHidden']}")
    print()
    print("Next step: on a Windows machine with Excel, run:")
    print("    powershell -ExecutionPolicy Bypass -File assemble.ps1")
    print("to inject VBA from vba/ and save the final .xlsm.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
