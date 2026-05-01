# EXAUSTAO 360 ENTERPRISE PRO

Kit gerador da solucao industrial **EXAUSTAO_360_ENTERPRISE_PRO.xlsm** para
gestao de confiabilidade de sistemas de exaustao em ambientes industriais.

> Reengenharia do prototipo `EXAUSTAO_360_pilm.xlsm` (20 abas, ~R$ 3.500,00)
> em produto B2B premium auditavel com 7 abas, ROI demonstravel e camada
> VBA profissional.

## Estrutura

```
.
├── build.py              # Gera dist/EXAUSTAO_360_ENTERPRISE_PRO.xlsx (cross-platform)
├── assemble.ps1          # Windows: injeta VBA e salva como .xlsm
├── build_all.bat         # Atalho Windows: build + assemble
├── requirements.txt      # openpyxl
├── vba/                  # Modulos VBA (.bas / .cls) com Option Explicit
├── docs/                 # Relatorio de entrega, manual, modelo comercial
└── dist/                 # Saida do build
```

## Como compilar (Windows com Excel instalado)

```powershell
pip install -r requirements.txt
python build.py
powershell -ExecutionPolicy Bypass -File assemble.ps1
```

Saida: `dist/EXAUSTAO_360_ENTERPRISE_PRO.xlsm`.

## Como compilar so o esqueleto (qualquer SO)

```bash
pip install -r requirements.txt
python build.py
```

Saida: `dist/EXAUSTAO_360_ENTERPRISE_PRO.xlsx` (sem macros). Para obter o
`.xlsm` final e necessario rodar `assemble.ps1` em uma maquina Windows com
Excel instalado.

## 7 abas finais

1. **DASHBOARD** - KPIs executivos, Pareto, tendencia, ranking de fornos.
2. **REGISTRO** - Formulario validado de eventos.
3. **BASE** - tbEventos (base mestra).
4. **ANALISE** - MTBF/MTTR, FMEA, plano de acao.
5. **SIMULADOR** - Cenarios conservador / realista / agressivo, ROI, payback.
6. **MODELO_COMERCIAL** - Pacotes Diagnostico / Pro / Enterprise, pricing.
7. **CONFIG** - Parametros, fornos, componentes, tipos, operadores.

A aba **LOG** existe como `xlSheetVeryHidden` para auditoria.

## Documentacao

- [`docs/RELATORIO_ENTREGA.md`](docs/RELATORIO_ENTREGA.md) - O que foi corrigido / redesenhado.
- [`docs/MANUAL_USUARIO.md`](docs/MANUAL_USUARIO.md) - Operacao do sistema.
- [`docs/MODELO_COMERCIAL.md`](docs/MODELO_COMERCIAL.md) - Pricing e ROI.
