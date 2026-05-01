# Manual do Usuario - EXAUSTAO 360 ENTERPRISE PRO

## 1. Abertura segura

1. Abra `EXAUSTAO_360_ENTERPRISE_PRO.xlsm`.
2. Quando o Excel pedir, **habilite as macros** (clique em "Habilitar
   conteudo").
3. O arquivo abre automaticamente em `DASHBOARD`, em **Modo Usuario**
   (CONFIG fica oculta, abas operacionais protegidas contra edicao
   acidental).

---

## 2. Fluxo diario

```
        +----------------+        +----------------+        +-----------+
opera-> | REGISTRO       | --> +  | BASE/tbEventos | -----> | DASHBOARD |
   dor  | (formulario)   |     ^  | (auditavel)    |   KPIs |  & ANALISE|
        +----------------+     |  +----------------+        +-----------+
                               |
               botao SALVAR/REGISTRAR (modRegistro)
```

### 2.1 Registrar um evento

1. Clique em **REGISTRO** (ou botao "Registro" no menu).
2. Preencha:
   - **Data, Hora, Operador, Forno, Tipo de Evento, Criticidade** (obrigatorios).
   - **Componente, Categoria, Duracao, Custo, Status, Observacao** (opcionais).
3. Clique em **VALIDAR**.
4. Se OK, clique em **SALVAR / REGISTRAR**. O evento e gravado em
   `BASE > tbEventos`, e o dashboard recalcula automaticamente.
5. Use **LIMPAR** para iniciar novo registro.

### 2.2 Consultar o dashboard

- Clique em **AtualizarDashboard** (ou execute via `Atualizar Tudo`).
- Veja KPIs no topo: Disponibilidade, MTBF, MTTR, Custo Total, ROI.
- Pareto, Ranking de Fornos e Tendencia mensal sao recalculados.
- A area "Recomendacoes Automaticas" sugere acoes baseadas em metas
  definidas em CONFIG.

### 2.3 Analise tecnica

Aba **ANALISE**:
- Tabela MTBF/MTTR por forno com status colorido (OK/Atencao/Revisar).
- FMEA com calculo de RPN (Severidade x Ocorrencia x Deteccao).
- Plano de Acao com Status condicional (Planejado/Em andamento/Concluido/Atrasado).

### 2.4 Simulacao financeira

Aba **SIMULADOR**:
1. Edite as **Premissas** (custo hora parada, perda anual, investimento, etc).
2. Edite os **% de reducao** dos cenarios.
3. Clique em **SIMULAR CENARIOS** para recalcular.
4. Para analise probabilistica, clique em **SIMULAR MONTE CARLO**:
   5.000 simulacoes em segundos, retornando media, desvio e probabilidade
   de payback positivo.

### 2.5 Modelo comercial

Aba **MODELO_COMERCIAL**:
- 3 pacotes (Diagnostico, Pro, Enterprise).
- Calculadora de ROI personalizada (lado direito).
- Argumento comercial ja escrito.

---

## 3. Modo Supervisor

Para editar parametros, ativar:

1. Clique em **AlternarModoSupervisor** (ou via menu Desenvolvedor > Macros).
2. Confirme no aviso.
3. Aba **CONFIG** fica visivel, todas as abas destravam.
4. Edite `tbParametros` (custo hora parada, metas, percentuais).
5. Clique novamente em **AlternarModoSupervisor** para voltar ao
   Modo Usuario.

---

## 4. Backup e auditoria

- **CriarBackup360**: gera copia em `<pasta>\backup\` com timestamp.
- **LOG (oculto)**: registra cada acao com data/hora, usuario, aba,
  registro afetado e resultado. Acessivel apenas em Modo Supervisor.

---

## 5. Exportar

- **ExportarRelatorioPDF360**: gera PDF de DASHBOARD + ANALISE em
  paisagem, ja formatado para diretoria.
- **ExportarConsulta**: gera CSV completo de `tbEventos`.

---

## 6. Solucao de problemas

| Sintoma | Diagnostico | Acao |
|---|---|---|
| "Macros desabilitadas" | Habilitar conteudo | Clique em "Habilitar conteudo" |
| Botao nao executa | Macro nao vinculada | Botao direito no shape > "Atribuir macro" |
| `Run-time error 1004` | Tabela renomeada | Restaurar nome em CONFIG; conferir constantes em `modGlobais` |
| Dashboard congelado | Application travado | Fechar e reabrir; o `Workbook_Open` restaura runtime |
| Aba CONFIG sumiu | Modo Usuario ativo | Ativar Modo Supervisor |
| `#NOME?` em formula | Defined name removido | Reabrir o arquivo gerado por `build.py` (defined names sao recriados) |

---

## 7. Padroes obrigatorios

- **Nao** edite diretamente `BASE > tbEventos` linha a linha. Use sempre
  REGISTRO.
- **Nao** apague linhas do LOG.
- **Nao** renomeie planilhas ou tabelas - as macros buscam por nome.
- **Nao** salve como .xlsx (perde as macros). Mantenha sempre .xlsm.
