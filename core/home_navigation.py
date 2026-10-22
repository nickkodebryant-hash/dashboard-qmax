PAINEIS = [
    (
        "Dashboard Operacional",
        "📋",
        "Dashboard Operacional",
        "Acompanhe metas, andamento mensal e ações da qualidade.",
    ),
    (
        "Visão Executiva Q-MAX",
        "🎯",
        "Visão Executiva Q-MAX",
        "Indicadores essenciais para uma leitura executiva rápida.",
    ),
    (
        "Relatórios",
        "📊",
        "Relatórios",
        "Consulte relatórios de ações abertas e vencidas.",
    ),
    (
        "Target Gerência",
        "🎯",
        "Target Gerência",
        "Acompanhe a distribuição ponderada da meta e o realizado semanal.",
    ),
    (
        "Custos da não Qualidade",
        "💰",
        "Custos da não Qualidade",
        "Explore lançamentos históricos, despesas por projeto e indicadores de custo.",
    ),
    (
        "Organograma Semanal",
        "🗓️",
        "Organograma Semanal",
        "Visualize a capacidade semanal, atividades e distribuição das horas.",
    ),
    (
        "Certificados de Calibração",
        "🔬",
        "Certificados de Calibração",
        "Analise certificados PDF com os critérios internos de calibração.",
    ),
    (
        "Gestão de Calibração",
        "📏",
        "Gestão de Calibração",
        "Acompanhe vencimentos, status, equipamentos e relatórios de calibração.",
    ),
    (
        "Gestão de CN's",
        "📋",
        "Gestão de CN's",
        "Gerencie CNs, etapas, prazos, responsáveis e confirmações em um só lugar.",
    ),
    (
        "⚙ Administração",
        "⚙️",
        "Administração",
        "Mantenha os cadastros mestres compartilhados de CNs.",
    ),
]

PAINEIS_POR_NOME = {painel[0]: painel for painel in PAINEIS}
GRUPOS_FERRAMENTAS = {
    "grupo_qualidade": {
        "titulo": "Conectados com a qualidade",
        "icone": "✅",
        "descricao": "Ações, acompanhamento e melhoria contínua da qualidade.",
        "paineis": (
            "Dashboard Operacional",
            "Visão Executiva Q-MAX",
            "Relatórios",
            "Target Gerência",
        ),
    },
    "grupo_confiabilidade": {
        "titulo": "Confiabilidade",
        "icone": "🏭",
        "descricao": "Ferramentas independentes para capacidade e gestão de CNs.",
        "paineis": ("Organograma Semanal", "Gestão de CN's"),
    },
    "grupo_calibracao": {
        "titulo": "Calibração",
        "icone": "📏",
        "descricao": "Gestão dos equipamentos e análise de certificados.",
        "paineis": ("Gestão de Calibração", "Certificados de Calibração"),
    },
}
CARDS_CENTRAL = (
    (
        GRUPOS_FERRAMENTAS["grupo_qualidade"]["titulo"],
        "✅",
        GRUPOS_FERRAMENTAS["grupo_qualidade"]["descricao"],
        "grupo_qualidade",
    ),
    (
        "Custos da não Qualidade",
        "💰",
        "Explore lançamentos históricos, despesas por projeto e indicadores de custo.",
        "Custos da não Qualidade",
    ),
    (
        "Confiabilidade",
        "🏭",
        GRUPOS_FERRAMENTAS["grupo_confiabilidade"]["descricao"],
        "grupo_confiabilidade",
    ),
    (
        "Calibração",
        "📏",
        GRUPOS_FERRAMENTAS["grupo_calibracao"]["descricao"],
        "grupo_calibracao",
    ),
    (
        "Administração",
        "⚙️",
        "Gerencie cadastros mestres compartilhados e responsáveis.",
        "⚙ Administração",
    ),
)

DESTINOS_VALIDOS = {
    *(painel[0] for painel in PAINEIS),
    *GRUPOS_FERRAMENTAS,
}
