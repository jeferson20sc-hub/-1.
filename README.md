# Controle de Fornos

Planilha rapida e enxuta para registro, monitoramento e analise de fornos.

## Arquivos

| Arquivo | Para que serve |
| --- | --- |
| `Controle_Fornos.xlsm` | Planilha com macros (REGISTRAR, LIMPAR, LIMPAR FILTROS) |
| `Controle_Fornos.xlsx` | Mesma planilha sem macros (fallback) |
| `modFornos.bas` | Codigo VBA do modulo de macros (importavel manualmente) |
| `Sheet2_Registro.txt` | Codigo do modulo da aba Registro (event handlers) |
| `Sheet3_Filtros.txt` | Codigo do modulo da aba Filtros (event handlers) |
| `build_xlsm.py` | Script Python que regenera tudo |

## Estrutura (3 abas)

### 1. Dashboard
- 5 KPIs: total de registros, fornos distintos, falhas, operacoes, mes mais ativo.
- 3 graficos: registros por mes, top 10 dias com mais registros, distribuicao por tipo.
- Tabela resumo Mes x Tipo (12 meses x 4 tipos + total) com escala de cor.

### 2. Registro
- Form de input rapido: **Data, Forno, Tipo, Lado**.
- Botao **REGISTRAR** (clique na celula G4) - adiciona o registro na tabela.
- Botao **LIMPAR** (clique na celula J4) - reseta o form.
- 4 tabelas lado a lado (uma por tipo: Operacao, Manutencao, Falha, Inspecao) com forno + qtd e total no rodape.
- Tabela "Resumo por Forno" com contagem por tipo (O/M/F/I) e percentual de falhas.

### 3. Filtros
- Filtros: Data inicial, Data final, Forno, Tipo, Lado, Dia, Mes, Ano.
- Busca livre por texto (forno/tipo/lado/data).
- Botao **LIMPAR** (clique na celula J4).
- Tabela de resultados com 150 linhas, formatacao zebra e cor por tipo.

## Performance

Decisoes para deixar o arquivo **rapido**:

- **3 abas visiveis apenas** (`_Dados` fica oculta).
- **Tabela Excel nativa** (`tblFornos`) - referencia estruturada, otimizada pelo motor.
- **Formulas COUNTIFS/SUMIFS** (nao volateis). Sem `OFFSET`, `INDIRECT`, `NOW`.
- Filtros usam **AGGREGATE(15,6,...)** que suporta arrays sem CSE.
- Limites: 150 linhas no resultado de filtro, 200 linhas de auxiliares - suficiente sem peso.
- VBA so durante clique: `Application.Calculation = xlCalculationManual` durante o registro evita recalculo intermediario.

## Como usar

**Com macros (xlsm):**
1. Abra `Controle_Fornos.xlsm`.
2. Quando o Excel pedir, **habilite as macros** (barra amarela "Habilitar Conteudo").
3. Na aba **Registro**: preencha Data, Forno, Tipo, Lado e clique em **REGISTRAR**.
4. Na aba **Filtros**: preencha qualquer combinacao de filtros - resultados atualizam ao vivo.

**Se as macros nao funcionarem:** veja "Importar VBA manualmente" abaixo.

**Sem macros (xlsx ou xlsm com macros desabilitadas):**
1. Os graficos, KPIs e formulas continuam funcionando.
2. Para registrar manualmente: clique-direito na barra de abas > Re-exibir > selecione `_Dados` > digite na proxima linha vazia da tabela. As outras abas atualizam sozinhas.

## Importar VBA manualmente

Caso o Excel rejeite a VBA embarcada:

1. Abra `Controle_Fornos.xlsm` no Excel.
2. Pressione **Alt+F11** para abrir o Editor VBA.
3. Menu **Arquivo > Importar Arquivo** > escolha `modFornos.bas`.
4. No painel esquerdo, abra o modulo da planilha **Sheet2 (Registro)** com duplo clique.
   - Apague qualquer codigo existente.
   - Cole todo o conteudo de `Sheet2_Registro.txt`.
5. Faca o mesmo com **Sheet3 (Filtros)** e `Sheet3_Filtros.txt`.
6. Salve (Ctrl+S) e feche o editor. Pronto.

## Regenerar do zero

```bash
python3 build_xlsm.py
```

Gera `Controle_Fornos.xlsx`, `Controle_Fornos.xlsm`, `modFornos.bas`, `Sheet2_Registro.txt`, `Sheet3_Filtros.txt`.
