# Modelo Comercial - EXAUSTAO 360 ENTERPRISE PRO

## 1. Proposta de valor (1 frase)

EXAUSTAO 360 reduz perdas operacionais em sistemas industriais de exaustao
combinando **registro auditavel de eventos**, **analise FMEA**,
**MTBF/MTTR por equipamento** e **simulador de ROI** - entregando ao
gestor industrial uma decisao baseada em numero, e ao financeiro uma
evidencia mensuravel de retorno.

---

## 2. Problema industrial resolvido

Plantas que operam fornos, tubulacoes, bujoes refratarios, marmitas e
mantas isolantes enfrentam:

- Falhas recorrentes que custam **R$ 10 mil a R$ 80 mil por hora parada**.
- Perda anual tipica entre **R$ 1 mi e R$ 5 mi** em plantas medias.
- Decisoes baseadas em planilhas isoladas, sem rastreabilidade.
- Falta de evidencia de ROI para aprovar investimentos em confiabilidade.
- Equipes reativas - reagem a falha em vez de prevenir.

EXAUSTAO 360 ataca esses 5 pontos com infraestrutura digital.

---

## 3. Como o sistema reduz perdas

| Mecanismo | Impacto direto |
|---|---|
| Registro estruturado de eventos | Visibilidade do que esta acontecendo (sem ela, nao ha gestao) |
| MTBF / MTTR por forno | Identifica equipamento mais critico - foco do plano de acao |
| FMEA com RPN | Prioriza modos de falha por risco (Severidade x Ocorrencia x Deteccao) |
| Plano de acao com status | Acompanha execucao e cobra responsaveis |
| Recomendacoes automaticas no DASHBOARD | Quando KPI cai abaixo da meta, sistema sugere acao |
| Simulador de cenarios | Conecta acao tecnica a impacto financeiro - habilita pitch para diretoria |

---

## 4. Como calcular ROI

```
Economia Anual  =  Perda Anual Atual  x  % Reducao Esperada
ROI Anual       =  (Economia Anual  -  Investimento)  /  Investimento
Payback (anos)  =  Investimento  /  Economia Anual
Contrato Sugerido = Economia Anual x 0.30   (30% como margem do cliente)
```

**Exemplo (cliente piloto medio):**

| Variavel | Valor |
|---|---|
| Perda anual atual | R$ 2.000.000 |
| Reducao esperada | 10% |
| Economia anual | **R$ 200.000** |
| Investimento (pacote Pro) | R$ 80.000 |
| ROI anual | **150%** |
| Payback | **5 meses** |

---

## 5. Pacotes

### 5.1 Diagnostico - R$ 24.000

- Implantacao em **1 planta** durante **30 dias**.
- Configuracao do EXAUSTAO 360 com parametros da planta.
- Treinamento basico (1 turma, 4 horas).
- Coleta de eventos por 30 dias para baseline.
- Suporte por 30 dias.
- Entrega: **Relatorio de diagnostico** com primeiros KPIs e plano de
  acao priorizado.

### 5.2 Pro (recomendado) - R$ 80.000

- Implantacao em **1 planta** durante **60 dias**.
- Configuracao + customizacao do dashboard ate 20%.
- Treinamento completo (3 turmas: operacao, manutencao, gestao).
- Suporte por **12 meses**.
- Atualizacoes e patches inclusos.
- Auditoria mensal e relatorio executivo.
- Entrega: produto plenamente operacional + 1 ciclo de melhoria.

### 5.3 Enterprise - R$ 240.000

- Multiunidade (ate 5 plantas).
- Customizacao total, modos de visualizacao por unidade.
- RBAC (Role-Based Access Control) sobre o Modo Supervisor.
- Suporte por **24 meses** com SLA.
- Trilha completa de treinamento + certificacao.
- Relatorios quinzenais para diretoria corporativa.
- Atualizacoes inclusas por 24 meses.
- Entrega: padronizacao corporativa de confiabilidade de exaustao.

> **Pricing dinamico:** o preco de cada pacote esta em
> `CONFIG > tbParametros > CustoSolucaoPro`. Diagnostico = 30% deste valor,
> Enterprise = 3x este valor. Edite uma vez para reprecificar todos.

### 5.4 Recorrencia (renovacao)

18% ao ano sobre o pacote contratado, cobrindo:
- Suporte tecnico em horario comercial.
- Patches e atualizacoes.
- Auditoria mensal.
- Acesso a novos casos de uso e templates de relatorio.

---

## 6. Argumentos de venda

### Para gestor de manutencao
> "Voce passa a ter visibilidade do que cada forno esta consumindo em
> manutencao corretiva e MTBF real - hoje voce so tem 'achismo'."

### Para diretor industrial
> "A cada R$ 80 mil investidos, a planta protege R$ 200 mil/ano em
> margem operacional - payback em 5 meses, ROI 150%."

### Para financeiro
> "Solucao auditavel, com log de cada acao, exportacao PDF para
> apresentacao em comites, e modelo de calculo replicavel para outras
> plantas."

### Para diretor de TI
> "Excel + VBA - infraestrutura ja aprovada, sem servidor, sem nuvem
> obrigatoria, sem licenciamento adicional. Operacional em 30 dias."

---

## 7. Objecoes comuns

| Objecao | Resposta |
|---|---|
| "Excel nao e seguro" | Modo Supervisor + LOG auditavel + protecao por aba. Para empresas que querem mais, oferecemos exportacao para SQL Server na versao Enterprise. |
| "Ja temos SAP/PIMS" | EXAUSTAO 360 e camada de **decisao**, nao de captura. Importa dados do PIMS via CSV/REST. |
| "Falta de dados historicos" | O Diagnostico de 30 dias coleta o baseline. Em 90 dias temos dados suficientes para validar economia. |
| "Sem orcamento agora" | Diagnostico (R$ 24 mil) cabe em capex de manutencao e gera o estudo que aprova o Pro. |

---

## 8. Pipeline comercial sugerido

1. **Lead** -> Demonstracao do arquivo .xlsm pronto (SIMULADOR + DASHBOARD).
2. **Discovery** -> Coleta da perda anual atual e do parque de fornos.
3. **Proposta** -> Diagnostico (R$ 24 mil, 30 dias).
4. **Diagnostico entregue** -> Sustentacao do Pro com numeros do cliente.
5. **Pro** -> R$ 80 mil, 60 dias, fecha em ~30 dias apos diagnostico.
6. **Recorrencia** -> 18% a.a. + upsell para Enterprise se houver outras plantas.

LTV tipico de cliente que evolui Diagnostico -> Pro -> Enterprise:
**R$ 350.000 - R$ 600.000 em 36 meses.**

---

## 9. Limites a comunicar com honestidade

- **Nao prometemos R$ X de economia sem dado real da planta.** O
  simulador trabalha com premissas editaveis, e a economia real depende
  de:
  - Disciplina operacional do registro (lixo entra = lixo sai).
  - Qualidade da execucao do plano de acao.
  - Maturidade da equipe de manutencao.
- **EXAUSTAO 360 nao substitui CMMS, ERP ou PIMS.** E uma camada
  executiva de confiabilidade que **conecta** a operacao ao financeiro.
- **Nao ha garantia de ROI especifico.** Ha um framework para mensurar
  ROI - quem entrega o ROI e a empresa cliente, nao a planilha.
