"""Seed data and reference lists for EXAUSTAO 360 ENTERPRISE PRO.

Sample data is illustrative; it lets the dashboard come alive on first open
without requiring real plant data. Replace via REGISTRO once deployed.
"""

from datetime import datetime, timedelta
import random

# ---------------------------------------------------------------------------
# Reference (CONFIG) data
# ---------------------------------------------------------------------------
FORNOS = [
    ("F-101", "Forno Rotativo 01", "Area A", 1200, "Operando"),
    ("F-102", "Forno Rotativo 02", "Area A", 1200, "Operando"),
    ("F-201", "Forno Tunel 01", "Area B", 800, "Operando"),
    ("F-202", "Forno Tunel 02", "Area B", 800, "Em manutencao"),
    ("F-301", "Forno Reverbero 01", "Area C", 1500, "Operando"),
    ("F-302", "Forno Reverbero 02", "Area C", 1500, "Operando"),
]

TIPOS_EVENTO = [
    ("Falha", "Corretiva", 5),
    ("Parada Programada", "Preventiva", 1),
    ("Inspecao", "Preditiva", 1),
    ("Troca de Componente", "Preventiva", 3),
    ("Reparo", "Corretiva", 4),
    ("Limpeza", "Operacional", 1),
    ("Calibracao", "Preditiva", 2),
]

CATEGORIAS = [
    ("Mecanica", "Falha mecanica de equipamento"),
    ("Eletrica", "Falha eletrica ou eletronica"),
    ("Operacional", "Erro ou desvio operacional"),
    ("Refratario", "Desgaste de revestimento refratario"),
    ("Termica", "Problema termico ou de combustao"),
    ("Instrumentacao", "Problema em sensores e instrumentacao"),
]

COMPONENTES = [
    ("C-TUB", "Tubulacao", "Tubulacao", 60),
    ("C-BUJ", "Bujao Refratario", "Bujao", 36),
    ("C-MAR", "Marmita", "Marmita", 48),
    ("C-MAN", "Manta Isolante", "Manta", 24),
    ("C-QUE", "Queimador", "Queimador", 72),
    ("C-VEN", "Ventilador de Exaustao", "Ventilador", 84),
    ("C-DAM", "Damper", "Damper", 60),
    ("C-SEN", "Sensor de Temperatura", "Instrumentacao", 36),
]

OPERADORES = [
    ("OP-01", "Carlos Silva", "A", "Manutencao"),
    ("OP-02", "Ana Costa", "B", "Operacao"),
    ("OP-03", "Joao Santos", "C", "Manutencao"),
    ("OP-04", "Maria Oliveira", "A", "Engenharia"),
    ("OP-05", "Pedro Almeida", "B", "Operacao"),
    ("OP-06", "Lucia Ferreira", "C", "Engenharia"),
]

TURNOS = [
    ("A", "06:00", "14:00"),
    ("B", "14:00", "22:00"),
    ("C", "22:00", "06:00"),
]

CRITICIDADES = [
    ("Baixa", 1, "Sem impacto significativo"),
    ("Media", 2, "Impacto controlado em produtividade"),
    ("Alta", 3, "Impacto relevante - acao em 24h"),
    ("Critica", 4, "Parada nao programada - acao imediata"),
]

REGRAS_RISCO = [
    ("MTBF baixo", "MTBF < 200 h", "Alta", "Investigar componente"),
    ("MTTR alto", "MTTR > 6 h", "Alta", "Revisar plano de manutencao"),
    ("Disponibilidade baixa", "Disponibilidade < 90%", "Critica", "Plano de acao imediato"),
    ("Recorrencia", "Mesmo componente > 3x/mes", "Critica", "Substituicao programada"),
    ("Custo alto", "Custo evento > R$ 50.000", "Alta", "Analise causa raiz"),
]

# ---------------------------------------------------------------------------
# Business parameters (CONFIG -> tbParametros)
# ---------------------------------------------------------------------------
PARAMETROS = [
    ("CustoHoraParada", 25000, "Custo medio por hora parada (R$)"),
    ("MetaDisponibilidade", 0.95, "Meta de disponibilidade da planta (%)"),
    ("LimiteCriticidadeAlta", 3, "Pontuacao a partir da qual evento e critico"),
    ("Moeda", "BRL", "Moeda padrao do sistema"),
    ("Planta", "Planta Industrial 01", "Identificacao da planta"),
    ("AreaResponsavel", "Engenharia de Confiabilidade", "Area responsavel"),
    ("VersaoSistema", "1.0.0 ENTERPRISE PRO", "Versao do sistema"),
    ("HorasOperacionalDia", 24, "Horas de operacao por dia"),
    ("DiasOperacionalAno", 330, "Dias de operacao por ano"),
    ("TaxaDescontoAnual", 0.12, "Taxa de desconto anual para VPL"),
    ("MetaMTBF", 720, "Meta de MTBF em horas"),
    ("MetaMTTR", 4, "Meta de MTTR em horas"),
    ("CustoSolucaoPro", 80000, "Custo do pacote PRO (R$)"),
    ("PercentualReducaoEsperada", 0.10, "Reducao esperada de perdas (%)"),
]

# ---------------------------------------------------------------------------
# Plano de Acao (ANALISE -> tbPlanoAcao)
# ---------------------------------------------------------------------------
PLANO_ACAO = [
    (1, "Substituicao programada de bujao no F-101", "Carlos Silva",
     datetime(2026, 5, 15), "Em andamento", "Alta", 35000, "Critica",
     "Compra do componente"),
    (2, "Inspecao termografica nos fornos da Area B", "Maria Oliveira",
     datetime(2026, 5, 30), "Planejado", "Alta", 12000, "Alta",
     "Agendar empresa especializada"),
    (3, "Revisao do plano preventivo de mantas", "Pedro Almeida",
     datetime(2026, 6, 10), "Planejado", "Media", 8000, "Media",
     "Analisar historico"),
    (4, "Calibracao geral de sensores de temperatura", "Lucia Ferreira",
     datetime(2026, 5, 20), "Concluido", "Media", 5000, "Baixa",
     "Concluido em 28/04"),
    (5, "Treinamento operacional - turno C", "Joao Santos",
     datetime(2026, 6, 25), "Planejado", "Baixa", 3000, "Media",
     "Definir conteudo programatico"),
    (6, "Estudo de criticidade FMEA - Forno F-202", "Ana Costa",
     datetime(2026, 5, 10), "Atrasado", "Critica", 50000, "Critica",
     "Reuniao de alinhamento pendente"),
]

# ---------------------------------------------------------------------------
# FMEA (ANALISE -> tbFMEA)
# ---------------------------------------------------------------------------
FMEA = [
    ("F-101", "Bujao Refratario", "Trinca por choque termico",
     "Parada do forno", "Inspecao visual semanal", 4, 3, 4, 48, "Alta"),
    ("F-101", "Tubulacao", "Vazamento por corrosao",
     "Perda de eficiencia", "Inspecao trimestral", 3, 4, 3, 36, "Media"),
    ("F-202", "Manta Isolante", "Degradacao por temperatura",
     "Perda termica", "Substituicao programada", 5, 5, 3, 75, "Critica"),
    ("F-301", "Queimador", "Entupimento de bicos",
     "Combustao incompleta", "Limpeza mensal", 4, 4, 4, 64, "Critica"),
    ("F-201", "Ventilador", "Desbalanceamento",
     "Vibracao excessiva", "Analise de vibracao", 3, 3, 4, 36, "Media"),
    ("F-302", "Damper", "Travamento mecanico",
     "Falha de regulagem", "Lubrificacao programada", 2, 3, 5, 30, "Baixa"),
    ("F-102", "Sensor Temperatura", "Drift de leitura",
     "Controle impreciso", "Calibracao semestral", 3, 2, 4, 24, "Baixa"),
]

# ---------------------------------------------------------------------------
# Sample events (BASE -> tbEventos)
# ---------------------------------------------------------------------------
def gerar_eventos_amostra(n: int = 60, seed: int = 42):
    """Generate n synthetic events spread over the last ~120 days."""
    random.seed(seed)
    eventos = []
    base = datetime(2026, 1, 1, 6, 0, 0)
    for i in range(1, n + 1):
        dias = random.randint(0, 119)
        horas = random.randint(0, 23)
        minutos = random.choice([0, 15, 30, 45])
        data = base + timedelta(days=dias, hours=horas, minutes=minutos)
        forno = random.choice(FORNOS)[0]
        tipo = random.choices(
            TIPOS_EVENTO,
            weights=[35, 20, 15, 10, 8, 8, 4],
            k=1,
        )[0]
        comp = random.choice(COMPONENTES)
        cat = random.choice(CATEGORIAS)[0]
        op = random.choice(OPERADORES)[1]
        crit = random.choices(
            ["Baixa", "Media", "Alta", "Critica"],
            weights=[40, 35, 18, 7],
            k=1,
        )[0]
        # Duration depends on type
        if tipo[0] in ("Falha", "Reparo"):
            dur = round(random.uniform(0.5, 12), 1)
        elif tipo[0] == "Parada Programada":
            dur = round(random.uniform(2, 8), 1)
        else:
            dur = round(random.uniform(0.2, 2), 1)
        # Cost loosely proportional to duration + criticality
        crit_mult = {"Baixa": 1, "Media": 2, "Alta": 4, "Critica": 8}[crit]
        custo = round(dur * 25000 * crit_mult / 4, 2)
        status = random.choice(["Resolvido", "Resolvido", "Resolvido", "Em analise", "Aberto"])
        obs = f"Evento {i:03d} registrado pelo turno"
        eventos.append((
            i,                       # ID
            data.date(),             # Data
            data.time(),             # Hora
            op,                      # Operador
            forno,                   # Forno
            tipo[0],                 # TipoEvento
            cat,                     # Categoria
            comp[1],                 # Componente
            crit,                    # Criticidade
            dur,                     # DuracaoHoras
            custo,                   # CustoEstimado
            obs,                     # Observacao
            status,                  # Status
        ))
    eventos.sort(key=lambda e: (e[1], e[2]))
    return eventos
