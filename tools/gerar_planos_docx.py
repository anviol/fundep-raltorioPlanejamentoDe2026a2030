"""Gera os planos de ação em .docx a partir do conteúdo dos arquivos .tex.

Cada função monta um plano com o mesmo conteúdo e o mesmo visual do PDF
compilado pelo LaTeX. Os arquivos são gravados na raiz do repositório,
com o mesmo nome do .tex de origem.

Uso:
    python3 tools/gerar_planos_docx.py
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from fundep_docx import (  # noqa: E402
    RAIZ,
    celula,
    escreve,
    hiperlink,
    item_lista,
    linha,
    mantem_tabelas_inteiras,
    nova_secao_paisagem,
    nova_tabela,
    novo_documento,
    paragrafo,
    quebra_pagina,
    titulo_atividade,
    titulo_documento,
    titulo_secao,
    vazias,
)

LINK_PLANILHA = (
    "https://fundep.sharepoint.com/:x:/s/msteams_d9be15_927448/"
    "IQDlWOOVcV0ySa6pQ8AKcar6AXgZc2DPSjLTqQ-oyDIG_2Y?e=EOus7n"
)
NOME_PLANILHA = "PE 26-30 - PPA - Modelo cronograma fisico-financeiro.xlsx"

# Larguras úteis: 17 cm em retrato e 25,7 cm em paisagem
COL_IDENTIFICACAO = [4.8, 12.2]
COL_RESUMO_PAISAGEM = [1.22, 11.26, 5.63, 1.39, 1.57, 4.63]


# ------------------------------------------------------------------
# Blocos comuns aos três planos
# ------------------------------------------------------------------

def tabela_identificacao(doc, campos):
    """Quadro de identificação do plano, com rótulos em azul claro."""
    tabela = nova_tabela(
        doc, COL_IDENTIFICACAO, borda_cor="bordaCinza", borda_pt=1
    )
    for rotulo, valor in campos:
        linha(
            tabela,
            [
                celula(rotulo, fundo="azulClaro", negrito=True, tamanho=10),
                celula(valor, tamanho=10),
            ],
        )
    return tabela


def lista_fofa(doc, categoria, itens, tamanho=12):
    """Categoria do diagnóstico FOFA seguida dos itens em marcadores."""
    paragrafo(
        doc,
        categoria,
        negrito=True,
        alinhamento="esquerda",
        espaco_antes=8,
        espaco_depois=4,
        entrelinha=1.0,
    )
    for texto in itens:
        item_lista(doc, "\u2022", texto, tamanho=tamanho, entrelinha=1.0)


def rodape_link(doc, tamanho=12):
    """Parágrafo com o link para o modelo de cronograma."""
    par = paragrafo(doc, espaco_antes=8, alinhamento="justificado")
    escreve(
        par,
        "Link para planilha com modelo de orçamento (baixar e editar na sua "
        "máquina, não editar o arquivo online): ",
        tamanho=tamanho,
        cor="azulVivo",
    )
    hiperlink(par, NOME_PLANILHA, LINK_PLANILHA, tamanho=tamanho)


def cabecalho_resumo(tabela, tamanho=9.5):
    linha(
        tabela,
        [
            celula("Item", fundo="azulCabecalho", negrito=True, tamanho=tamanho,
                   alinhamento="centro"),
            celula("Descrição", fundo="azulCabecalho", negrito=True, tamanho=tamanho,
                   alinhamento="centro"),
            celula("Resultados Esperados", fundo="azulCabecalho", negrito=True,
                   tamanho=tamanho, alinhamento="centro"),
            celula("Início", fundo="azulCabecalho", negrito=True, tamanho=tamanho,
                   alinhamento="centro"),
            celula("Término", fundo="azulCabecalho", negrito=True, tamanho=tamanho,
                   alinhamento="centro"),
            celula("Responsável", fundo="azulCabecalho", negrito=True, tamanho=tamanho,
                   alinhamento="centro"),
        ],
        alinhamento_vertical="centro",
    )


def linha_atividade_resumo(tabela, item, descricao, inicio, termino, responsavel):
    """Linha de atividade do resumo, em azul escuro com texto branco.

    Descrição e Resultados Esperados são mescladas, como no documento
    de origem.
    """
    comum = {"fundo": "azulTitulo", "negrito": True, "cor": "branco", "tamanho": 10}
    linha(
        tabela,
        [
            celula(item, alinhamento="centro", **comum),
            celula(descricao, span=2, **comum),
            celula(inicio, alinhamento="centro", **comum),
            celula(termino, alinhamento="centro", **comum),
            celula(responsavel, alinhamento="centro", **comum),
        ],
        alinhamento_vertical="centro",
    )


def linha_subatividade_resumo(
    tabela, item, descricao, resultado, inicio, termino, responsavel, zebra
):
    fundo = "cinzaClaro" if zebra else None
    return linha(
        tabela,
        [
            celula(item, fundo="azulClaro", alinhamento="centro", tamanho=10),
            celula(descricao, fundo=fundo, tamanho=10),
            celula(resultado, fundo=fundo, tamanho=10),
            celula(inicio, fundo=fundo, alinhamento="centro", tamanho=10),
            celula(termino, fundo=fundo, alinhamento="centro", tamanho=10),
            celula(responsavel, fundo=fundo, alinhamento="centro", italico=True,
                   tamanho=10),
        ],
    )


def cabecalho_cronograma(tabela, anos, tamanho_ano=10, tamanho_tri=9):
    """Duas primeiras linhas do cronograma: anos e trimestres.

    O parâmetro anos é uma lista de pares (ano, trimestres exibidos).
    """
    celulas = [celula(span=2)]
    for ano, trimestres in anos:
        celulas.append(
            celula(ano, span=len(trimestres), fundo="azulVivo", negrito=True,
                   cor="branco", tamanho=tamanho_ano, alinhamento="centro")
        )
    linha(tabela, celulas, alinhamento_vertical="centro")

    celulas = [
        celula("Atividade", span=2, fundo="azulCelula", negrito=True,
               tamanho=tamanho_ano)
    ]
    for _, trimestres in anos:
        for trimestre in trimestres:
            celulas.append(
                celula(trimestre, fundo="azulCelula", tamanho=tamanho_tri,
                       alinhamento="centro")
            )
    linha(tabela, celulas, alinhamento_vertical="centro")


def linha_total(tabela, colunas_valor, texto, tamanho=12):
    linha(
        tabela,
        [
            celula("TOTAL", span=2, fundo="azulVivo", negrito=True, cor="branco",
                   tamanho=tamanho),
            celula(texto, span=colunas_valor, fundo="azulVivo", negrito=True,
                   cor="branco", tamanho=tamanho, alinhamento="centro"),
        ],
        alinhamento_vertical="centro",
    )


# ------------------------------------------------------------------
# Plano 1: Implementar o GPF 2.0
# ------------------------------------------------------------------

def plano_gpf():
    doc = novo_documento(espaco_paragrafo_pt=8, entrelinha=1.08)

    titulo_documento(
        doc, "PLANEJAMENTO ESTRATÉGICO", "Gestão 2026 a 2030: Plano de Ação"
    )

    titulo_secao(doc, "IDENTIFICAÇÃO DO PLANO")
    tabela_identificacao(
        doc,
        [
            ("DIRETRIZ", "Transformação Digital"),
            ("AÇÃO", "Implementar o GPF 2.0"),
            ("RESPONSÁVEL", "Gerente de TI - André Vinícius de Oliveira"),
            (
                "OBJETIVO ESTRATÉGICO VINCULADO",
                "Otimização das Rotinas Operacionais por meio da Automação e "
                "Simplificação de Processos",
            ),
            (
                "META(S) OPERACIONAL(IS) VINCULADA(S)",
                "Desenvolver e implementar o GPF 2.0; integrar os agentes de "
                "inteligência artificial aos sistemas GPF 1.0 e GPF 2.0; garantir "
                "a máxima disponibilidade dos sistemas; e eliminar o uso de "
                "controles paralelos no fluxo de gestão de projetos.",
            ),
            ("PERÍODO DE ACOMPANHAMENTO", "Julho de 2026 a dezembro de 2029"),
            (
                "ÁREAS ENVOLVIDAS",
                "Conselho Diretor, Tecnologia da Informação, CIAs, Prestação de "
                "Contas, Compras, Importação, Logística, Gestão de Documentos, "
                "Administração de Pessoal, Gestão de Pessoas, Negócios e Parcerias, "
                "Financeiro e Contabilidade, Compliance, Jurídico.",
            ),
            ("VERSÃO", "0.5"),
        ],
    )

    titulo_secao(doc, "JUSTIFICATIVA")
    paragrafo(
        doc,
        "O crescimento da dependência tecnológica e da complexidade das demandas "
        "expõe fragilidades como fragmentação de sistemas, controles manuais "
        "(planilhas, e-mails) e variações informais entre áreas, comprometendo a "
        "rastreabilidade e a eficiência. A modernização tecnológica é essencial "
        "para sustentar a operação, integrar os sistemas corporativos e assegurar "
        "qualidade e previsibilidade.",
    )
    paragrafo(
        doc,
        "Este plano de ação organiza, ao longo do quadriênio 2026-2029, a "
        "consolidação de uma plataforma corporativa única, a ampliação da "
        "automação e do uso de inteligência artificial, a garantia de alta "
        "disponibilidade dos sistemas e a eliminação progressiva de controles "
        "paralelos, com acompanhamento sistemático de prazos, responsáveis e "
        "indicadores.",
    )
    paragrafo(
        doc,
        "Esta ação responde ao diagnóstico institucional elaborado pelos gerentes "
        "da Fundep. Relacionam-se a ela os seguintes itens da matriz FOFA, das "
        "fortalezas e dos problemas mapeados, identificados pela numeração de "
        "origem:",
    )

    lista_fofa(doc, "Forças", [
        "(101) Excelência na gestão administrativa e financeira de projetos.",
        "(103) Benchmarking nacional em excelência operacional no setor de fundações.",
        "(109) Capital humano altamente qualificado e multidisciplinar.",
        "(111) Vantagem competitiva em soluções de TI proprietárias.",
        "(114) Incentivo da Alta Direção ao uso de IA e automação.",
        "(110) Gestão tecnológica e de informações integrada e adaptável.",
        "(122) Orientação a dados (Business Intelligence).",
    ])
    lista_fofa(doc, "Fraquezas", [
        "(311) Confiabilidade dos dados financeiros e extratos.",
        "(313) Ausência de gestão de processos.",
        "(320) Mudanças de processos de forma setorizada.",
        "(332) Deficiência de processos e controles manuais.",
        "(319) Ausência de sistemas para suporte às atividades de algumas áreas.",
        "(333) Sistema ERP com deficiências.",
        "(328) Processos de trabalho fragmentados.",
        "(321) Falta de gestão de riscos.",
        "(303) Falta de gestão de contingência.",
        "(304) Manutenção da qualidade dos serviços diante de turnovers.",
        "(325) Falta de integração de indicadores financeiros de ponta a ponta.",
        "(340) Falta de regramento formal para as tomadas de decisão.",
        "(315) Dependência de pessoas específicas para processos críticos.",
        "(317) Segurança da informação e proteção de dados deficientes.",
        "(327) Decisões não baseadas em dados.",
        "(316) Falta de atendimento integral à LGPD.",
        "(301) Fragilidade na comunicação interna: institucional e interáreas.",
        "(323) Falta de cultura de gestão do conhecimento.",
        "(331) Falta de integração dos fluxos de documentos.",
        "(312) Instrumentos de governança desatualizados ou inexistentes.",
        "(342) Falta de direcionamento para as mudanças e inovações.",
        "(338) Avanço tecnológico mais rápido que a capacidade de adaptação interna.",
    ])
    lista_fofa(doc, "Oportunidades", [
        "(221) Novo Sistema Financiado Externamente.",
        "(206) Tecnologias digitais de Inteligência Artificial emergentes aplicadas "
        "a processos.",
        "(220) Parcerias estratégicas com outras fundações de apoio.",
    ])
    lista_fofa(doc, "Ameaças", [
        "(409) Escassez de mão de obra qualificada.",
        "(407) Exigências crescentes de órgãos de controle, aumentando a "
        "complexidade operacional.",
    ])
    lista_fofa(doc, "Fortalezas", [
        "(2102) Capacidade técnica operacional para atender demandas da UFMG de "
        "diferentes naturezas.",
        "(6101) Referência e pioneirismo técnico em tecnologias digitais em gestão "
        "e atividades entre as Fundações de Apoio.",
        "(6102) Constante evolução da fronteira tecnológica para tecnologias "
        "digitais, possibilitando aumento da eficiência operacional.",
    ])
    lista_fofa(doc, "Problemas", [
        "(2201) Fraca integração entre equipes na gestão ponta a ponta de projetos.",
        "(2202) Despadronização das operações gerais na Fundep.",
        "(2203) Gestão de risco não estruturada para ampliar a segurança da tomada "
        "de decisão.",
        "(4203) Dependência de pessoas de referência.",
        "(6201) Fragilidade no atendimento legal em segurança de dados.",
        "(6202) As necessidades técnicas digitais e de segurança são subestimadas "
        "comparadas às vontades dos usuários.",
    ])

    titulo_secao(doc, "ATIVIDADES")

    titulo_atividade(doc, "01. Desenvolver e implementar o GPF 2.0")
    paragrafo(
        doc,
        "Desenvolvimento e implantação de plataforma corporativa única, modular e "
        "escalável, cobrindo o ciclo de vida do projeto, de modo que, ao fim do "
        "ciclo 2026-2029, 100% dos projetos ativos operem integralmente no sistema. "
        "Desenvolvimento previsto até 2028 e implantação até 2029.",
    )
    paragrafo(
        doc,
        "Investimento previsto de R$ 2.832.967,52, dividido em três parcelas "
        "conforme a etapa:",
    )
    item_lista(
        doc,
        "a)",
        "Manutenção do espaço de trabalho, R$ 30.000,00 em 2027: adequação e "
        "conservação do ambiente ocupado pela equipe durante a construção da "
        "plataforma.",
    )
    item_lista(
        doc,
        "b)",
        "Lanche e utilidades da equipe terceirizada, R$ 202.967,52 entre o 3º "
        "trimestre de 2026 e o 2º trimestre de 2028: R$ 8.456,98 por mês, ao longo "
        "de 24 meses, cobrindo pães e frutas, auxiliar de limpeza, energia elétrica "
        "e água. Por ano, o gasto fica em R$ 50.741,88 em 2026 (seis meses), "
        "R$ 101.483,76 em 2027 e R$ 50.741,88 em 2028 (seis meses).",
    )
    item_lista(
        doc,
        "c)",
        "Implantação, R$ 2.600.000,00 em 2029: entrada em produção da plataforma e "
        "migração dos projetos ativos. A composição desta parcela (equipe, licenças, "
        "infraestrutura e serviços de terceiros) está a preencher.",
        espaco_depois=8,
    )
    paragrafo(
        doc,
        "Indicadores: percentual de projetos ativos operando na plataforma (meta: "
        "100% até 2029); percentual de escopo desenvolvido (meta: 100% até 2028), "
        "medido trimestralmente.",
    )

    titulo_atividade(
        doc,
        "02. Integrar os agentes de inteligência artificial aos sistemas GPF 1.0 "
        "e GPF 2.0.",
    )
    par = paragrafo(doc)
    escreve(
        par,
        "Integração dos agentes de inteligência artificial desenvolvidos no plano "
        "de ação ",
    )
    escreve(par, "Implementar Agentes de IA na Gestão de Projetos Ponta a Ponta",
            italico=True)
    escreve(
        par,
        " aos sistemas corporativos, de modo que operem tanto sobre o GPF 1.0, "
        "ainda em uso durante a transição, quanto sobre o GPF 2.0, à medida que "
        "seus módulos entrarem em produção. Período de julho/2026 a março/2028.",
    )
    paragrafo(
        doc,
        "Investimento previsto de R$ 0,00. A atividade não tem custo próprio porque "
        "os agentes já são custeados no plano de ação específico de agentes de IA, "
        "que responde pela aquisição de equipamentos, pelas licenças de uso e pela "
        "operação dos modelos. Aqui resta o trabalho de integração, executado pela "
        "equipe própria de TI com as ferramentas e a infraestrutura já contratadas, "
        "sem aquisição adicional prevista.",
    )
    paragrafo(
        doc,
        "Indicador: percentual de agentes de inteligência artificial integrados aos "
        "sistemas corporativos (meta: 100% dos agentes em produção até 2028), "
        "medido trimestralmente.",
    )

    titulo_atividade(doc, "03. Garantir a máxima disponibilidade dos sistemas.")
    paragrafo(
        doc,
        "Estruturação de processos e infraestrutura para assegurar a disponibilidade "
        "dos sistemas críticos. Período de junho/2026 a março/2029.",
    )
    paragrafo(
        doc,
        "Investimento previsto de R$ 4.550.000,00, o maior desta ação, por se "
        "tratar do custeio contínuo do ambiente que sustenta a operação: "
        "infraestrutura, licenças, contratos de suporte e monitoramento dos "
        "sistemas críticos. Diferente das demais atividades, o gasto se repete todo "
        "ano enquanto os sistemas estiverem em operação, e não se concentra em uma "
        "entrega. A composição detalhada de cada ano está a preencher, assim como o "
        "valor de 2026, ano de início da atividade.",
    )
    paragrafo(doc, "Gasto anual previsto para a atividade:")

    tabela = nova_tabela(
        doc, [2.6, 5.2], borda_cor="bordaEscura", borda_pt=0.5
    )
    linha(
        tabela,
        [
            celula("Ano", fundo="azulCabecalho", negrito=True, alinhamento="centro"),
            celula("Gasto anual", fundo="azulCabecalho", negrito=True,
                   alinhamento="centro"),
        ],
        alinhamento_vertical="centro",
    )
    for ano, valor, zebra in (
        ("2027", "R$ 1.511.500,00", False),
        ("2028", "R$ 1.570.000,00", True),
        ("2029", "R$ 1.468.500,00", False),
    ):
        fundo = "cinzaClaro" if zebra else None
        linha(
            tabela,
            [
                celula(ano, fundo=fundo, alinhamento="centro"),
                celula(valor, fundo=fundo, alinhamento="direita"),
            ],
            alinhamento_vertical="centro",
        )
    linha(
        tabela,
        [
            celula("Total", fundo="azulVivo", negrito=True, cor="branco",
                   alinhamento="centro"),
            celula("R$ 4.550.000,00", fundo="azulVivo", negrito=True, cor="branco",
                   alinhamento="direita"),
        ],
        alinhamento_vertical="centro",
    )
    paragrafo(doc, espaco_depois=8)

    paragrafo(
        doc,
        "Indicador: disponibilidade igual ou superior a 99% por sistema crítico, "
        "medida mensalmente e consolidada anualmente (excluídas manutenções "
        "programadas e comunicadas).",
    )

    titulo_atividade(
        doc,
        "04. Eliminar o uso de controles paralelos no fluxo de gestão de projetos "
        "ponta a ponta.",
    )
    paragrafo(
        doc,
        "Substituição de planilhas, e-mails e similares por rotinas integradas ao "
        "sistema corporativo, de ponta a ponta no fluxo de gestão de projetos ponta "
        "a ponta. Período de maio/2027 a dezembro/2029.",
    )
    paragrafo(
        doc,
        "Investimento previsto de R$ 0,00. As rotinas que substituem as planilhas e "
        "os e-mails são as próprias funcionalidades do GPF 2.0, já custeadas na "
        "atividade 01, de modo que nada é adquirido especificamente para esta "
        "atividade. O esforço é de mapeamento, transição e descontinuação dos "
        "controles paralelos, conduzido pela equipe de TI em conjunto com os "
        "gestores das áreas, com pessoal já alocado.",
    )
    paragrafo(
        doc,
        "Indicador: zero rotinas operacionais que ainda utilizem planilhas, e-mail "
        "ou similares até 2029.",
    )

    # --- Resumo das atividades, em retrato, como no PDF deste plano
    titulo_secao(doc, "RESUMO DAS ATIVIDADES")
    tabela = nova_tabela(
        doc,
        [0.86, 5.92, 5.16, 1.15, 1.43, 2.48],
        borda_cor="bordaEscura",
        borda_pt=0.5,
        margem_celula_cm=0.1,
    )
    cabecalho_resumo(tabela, tamanho=9)
    resumo = [
        (
            "01",
            "Desenvolver e implementar o GPF 2.0 como plataforma corporativa "
            "única, modular e escalável.",
            "100% dos projetos ativos operando na plataforma e 100% do escopo "
            "desenvolvido.",
            "jul/26",
            "dez/29",
            "André Oliveira",
            False,
        ),
        (
            "02",
            "Integrar os agentes de inteligência artificial aos sistemas GPF 1.0 "
            "e GPF 2.0.",
            "100% dos agentes de IA em produção integrados aos sistemas "
            "corporativos.",
            "jul/26",
            "mar/28",
            "André Oliveira",
            True,
        ),
        (
            "03",
            "Garantir a máxima disponibilidade dos sistemas.",
            "Disponibilidade \u2265 99% por sistema crítico.",
            "jun/26",
            "dez/29",
            "André Oliveira",
            False,
        ),
        (
            "04",
            "Eliminar o uso de controles paralelos no fluxo de gestão de projetos "
            "ponta a ponta.",
            "Zero rotinas operacionais utilizando planilhas, e-mail ou similares.",
            "mai/27",
            "dez/29",
            "André Oliveira e Gestores das áreas",
            True,
        ),
    ]
    for item, descricao, resultado, inicio, termino, responsavel, zebra in resumo:
        fundo = "cinzaClaro" if zebra else None
        linha(
            tabela,
            [
                celula(item, fundo=fundo, alinhamento="centro", tamanho=9),
                celula(descricao, fundo=fundo, tamanho=9),
                celula(resultado, fundo=fundo, tamanho=9),
                celula(inicio, fundo=fundo, alinhamento="centro", tamanho=9),
                celula(termino, fundo=fundo, alinhamento="centro", tamanho=9),
                celula(responsavel, fundo=fundo, alinhamento="centro", italico=True,
                       tamanho=9),
            ],
        )

    # --- Cronograma físico-financeiro, em retrato, com régua de 16 trimestres
    titulo_secao(doc, "CRONOGRAMA FÍSICO-FINANCEIRO")
    larguras = [0.5, 4.3] + [0.7625] * 16
    tabela = nova_tabela(
        doc, larguras, borda_cor="bordaCinza", borda_pt=0.4, margem_celula_cm=0.04
    )
    cabecalho_cronograma(
        tabela,
        [
            ("2026", ["2º tri", "3º tri", "4º tri"]),
            ("2027", ["1º tri", "2º tri", "3º tri", "4º tri"]),
            ("2028", ["1º tri", "2º tri", "3º tri", "4º tri"]),
            ("2029", ["1º tri", "2º tri", "3º tri", "4º tri"]),
            ("2030", ["1º tri"]),
        ],
        tamanho_ano=6,
        tamanho_tri=6,
    )

    def rotulo(texto):
        return celula(texto, span=2, fundo="cinzaMedio", negrito=True, tamanho=6)

    def valor(texto, span, fundo, negrito=False):
        return celula(texto, span=span, fundo=fundo, negrito=negrito, tamanho=6,
                      alinhamento="centro")

    def alinea(texto):
        return [celula(), celula(texto, tamanho=6)]

    linha(tabela, [
        rotulo("01 - Desenvolver e implementar o GPF 2.0"),
        celula(),
        valor("R$ 2.832.967,52", 14, "cinzaMedio", negrito=True),
        celula(),
    ], alinhamento_vertical="centro")
    linha(tabela, alinea("a) Desenvolvimento: manutenção do espaço de trabalho") + [
        celula(span=3),
        valor("R$ 30.000,00", 4, "cinzaE8"),
        celula(span=9),
    ], alinhamento_vertical="centro")
    linha(tabela, alinea(
        "b) Desenvolvimento: lanche e utilidades da equipe terceirizada") + [
        celula(),
        valor("R$ 202.967,52", 8, "cinzaE8"),
        celula(span=7),
    ], alinhamento_vertical="centro")
    linha(tabela, alinea("c) Implantação") + [
        celula(span=11),
        valor("R$ 2.600.000,00", 4, "cinzaE8"),
        celula(),
    ], alinhamento_vertical="centro")
    linha(tabela, [
        rotulo("02 - Integrar agentes de IA ao GPF 1.0 e 2.0"),
        celula(),
        valor("R$ 0,00", 7, "cinzaMedio", negrito=True),
    ] + vazias(8), alinhamento_vertical="centro")
    linha(tabela, [
        rotulo("03 - Garantir máxima disponibilidade dos sistemas"),
        celula(span=3, fundo="cinzaMedio"),
        valor("R$ 1.511.500,00", 4, "cinzaMedio", negrito=True),
        valor("R$ 1.570.000,00", 4, "cinzaMedio", negrito=True),
        valor("R$ 1.468.500,00", 4, "cinzaMedio", negrito=True),
        celula(fundo="cinzaMedio"),
    ], alinhamento_vertical="centro")
    linha(tabela, [
        rotulo("04 - Eliminar uso de controles paralelos"),
    ] + vazias(4) + [
        valor("R$ 0,00", 12, "cinzaMedio", negrito=True),
    ], alinhamento_vertical="centro")
    linha_total(
        tabela,
        16,
        "R$ 7.382.967,52 (parcial: trechos sem valor ainda a definir)",
        tamanho=6,
    )

    return doc, "PE 26-30 - PPA - Plano de Acao - Implementar o GPF 2.0.docx"


# ------------------------------------------------------------------
# Plano 2: Implementar Agentes de IA
# ------------------------------------------------------------------

def plano_agentes_ia():
    doc = novo_documento(espaco_paragrafo_pt=12, entrelinha=1.5)

    titulo_documento(
        doc, "PLANO DE AÇÃO", "Planejamento Estratégico Fundep - Gestão 2026 a 2030"
    )

    titulo_secao(doc, "IDENTIFICAÇÃO DO PLANO")
    tabela_identificacao(
        doc,
        [
            ("DIRETRIZ", "Transformação Digital"),
            ("AÇÃO", "Implementar Agentes de IA na Gestão de Projetos Ponta a Ponta."),
            ("RESPONSÁVEL", "Gerente de TI - André Vinícius de Oliveira"),
            (
                "OBJETIVO ESTRATÉGICO VINCULADO",
                "Otimização das Rotinas Operacionais por meio da Automação e "
                "Simplificação de Processos",
            ),
            (
                "META(S) OPERACIONAL(IS) VINCULADA(S)",
                "Reduzir o tempo operacional com agentes de inteligência artificial.",
            ),
            ("PERÍODO DE ACOMPANHAMENTO", "Julho de 2026 a março de 2028"),
            (
                "ÁREAS ENVOLVIDAS",
                "Conselho Diretor, Tecnologia da Informação, CIAs, Prestação de "
                "Contas, Compras, Importação, Logística, Gestão de Documentos, "
                "Administração de Pessoal, Gestão de Pessoas, Negócios e Parcerias, "
                "Financeiro e Contabilidade, Compliance, Jurídico.",
            ),
            ("VERSÃO", "0.5"),
        ],
    )

    titulo_secao(doc, "JUSTIFICATIVA")
    paragrafo(
        doc,
        "O crescimento da dependência tecnológica e da complexidade das demandas "
        "expõe fragilidades como fragmentação de sistemas, controles manuais "
        "(planilhas, e-mails) e variações informais entre áreas, comprometendo a "
        "rastreabilidade e a eficiência. A modernização tecnológica é essencial "
        "para sustentar a operação, padronizar processos críticos e assegurar "
        "qualidade e previsibilidade.",
    )
    paragrafo(
        doc,
        "Este plano de ação organiza, ao longo do quadriênio 2026-2029, a "
        "consolidação de uma plataforma corporativa única, a ampliação da "
        "automação e do uso de inteligência artificial, a garantia de alta "
        "disponibilidade dos sistemas e a eliminação progressiva de controles "
        "paralelos, com acompanhamento sistemático de prazos, responsáveis e "
        "indicadores.",
    )
    paragrafo(
        doc,
        "Esta ação responde ao diagnóstico institucional elaborado pelos gerentes "
        "da Fundep. Relacionam-se a ela os seguintes itens da matriz FOFA, das "
        "fortalezas e dos problemas mapeados, identificados pela numeração de "
        "origem:",
    )

    lista_fofa(doc, "Forças", [
        "(101) Excelência na gestão administrativa e financeira de projetos.",
        "(103) Benchmarking nacional em excelência operacional no setor de fundações.",
        "(109) Capital humano altamente qualificado e multidisciplinar.",
        "(111) Vantagem competitiva em soluções de TI proprietárias.",
        "(114) Incentivo da Alta Direção ao uso de IA e automação.",
        "(110) Gestão tecnológica e de informações integrada e adaptável.",
        "(122) Orientação a dados (Business Intelligence).",
    ])
    lista_fofa(doc, "Fraquezas", [
        "(311) Confiabilidade dos dados financeiros e extratos.",
        "(313) Ausência de gestão de processos.",
        "(320) Mudanças de processos de forma setorizada.",
        "(332) Deficiência de processos e controles manuais.",
        "(319) Ausência de sistemas para suporte às atividades de algumas áreas.",
        "(333) Sistema ERP com deficiências.",
        "(328) Processos de trabalho fragmentados.",
        "(325) Falta de integração de indicadores financeiros de ponta a ponta.",
        "(304) Manutenção da qualidade dos serviços diante de turnovers.",
        "(315) Dependência de pessoas específicas para processos críticos.",
        "(317) Segurança da informação e proteção de dados deficientes.",
        "(327) Decisões não baseadas em dados.",
        "(316) Falta de atendimento integral à LGPD.",
        "(323) Falta de cultura de gestão do conhecimento.",
        "(331) Falta de integração dos fluxos de documentos.",
        "(342) Falta de direcionamento para as mudanças e inovações.",
        "(338) Avanço tecnológico mais rápido que a capacidade de adaptação interna.",
    ])
    lista_fofa(doc, "Oportunidades", [
        "(221) Novo Sistema Financiado Externamente.",
        "(206) Tecnologias digitais de Inteligência Artificial emergentes aplicadas "
        "a processos.",
        "(220) Parcerias estratégicas com outras fundações de apoio.",
    ])
    lista_fofa(doc, "Ameaças", [
        "(409) Escassez de mão de obra qualificada.",
        "(407) Exigências crescentes de órgãos de controle, aumentando a "
        "complexidade operacional.",
    ])
    lista_fofa(doc, "Fortalezas", [
        "(2102) Capacidade técnica operacional para atender demandas da UFMG de "
        "diferentes naturezas.",
        "(6101) Referência e pioneirismo técnico em tecnologias digitais em gestão "
        "e atividades entre as Fundações de Apoio.",
        "(6102) Constante evolução da fronteira tecnológica para tecnologias "
        "digitais, possibilitando aumento da eficiência operacional.",
    ])
    lista_fofa(doc, "Problemas", [
        "(2201) Fraca integração entre equipes na gestão ponta a ponta de projetos.",
        "(2202) Despadronização das operações gerais na Fundep.",
        "(2203) Gestão de risco não estruturada para ampliar a segurança da tomada "
        "de decisão.",
        "(4203) Dependência de pessoas de referência.",
        "(6201) Fragilidade no atendimento legal em segurança de dados.",
        "(6202) As necessidades técnicas digitais e de segurança são subestimadas "
        "comparadas às vontades dos usuários.",
    ])

    titulo_secao(doc, "ATIVIDADES")
    titulo_atividade(
        doc, "01. Reduzir o tempo operacional com agentes de inteligência artificial"
    )
    paragrafo(
        doc,
        "Adoção de agentes de inteligência artificial para apoiar processos "
        "operacionais, reduzindo o tempo médio de execução. Período de julho de "
        "2026 a março de 2028.",
    )
    paragrafo(
        doc,
        "Indicador: redução mínima de 20% no tempo médio de execução dos processos "
        "apoiados por inteligência artificial até 2029, tomando como base o período "
        "anterior à implementação.",
    )
    paragrafo(
        doc,
        "Investimento total de R$ 315.000,00 no período de julho de 2026 a março de "
        "2028, integralmente alocado nos custos de desenvolvimento detalhados na "
        "alínea b). A alínea a) não tem custo próprio.",
    )
    item_lista(
        doc,
        "a)",
        "Investigar e implementar agentes de inteligência artificial nos processos "
        "operacionais.",
        espaco_depois=6,
    )
    paragrafo(
        doc,
        "Levantamento dos processos operacionais da Fundep para identificar aqueles "
        "em que um agente de inteligência artificial é a solução adequada, seguido "
        "da implementação dos agentes nos processos selecionados, com apuração do "
        "tempo de execução antes e depois da adoção.",
        recuo_esquerda=1.25,
        espaco_depois=12,
    )
    item_lista(doc, "b)", "Desenvolvimento dos agentes de inteligência artificial.",
               espaco_depois=6)
    paragrafo(
        doc,
        "Construção dos agentes selecionados na alínea anterior, com equipe "
        "contratada, ambiente de execução em nuvem e ferramentas de inteligência "
        "artificial de terceiros. Os custos previstos são:",
        recuo_esquerda=1.25,
        espaco_depois=6,
    )
    for texto in (
        "Equipamentos para a equipe contratada: R$ 42.000,00, em aquisição única no "
        "início do projeto.",
        "Uso de ambiente em nuvem para executar as aplicações: R$ 5.000,00 por mês, "
        "do início ao fim do projeto.",
        "Uso de inteligência artificial de terceiros: R$ 5.000,00 por mês, do início "
        "ao fim do projeto.",
        "Agente de inteligência artificial de terceiro para auxiliar no "
        "desenvolvimento: R$ 3.000,00 por mês, do início ao fim do projeto.",
    ):
        item_lista(doc, "\u2022", texto, recuo=2.0, espaco_depois=6)
    paragrafo(
        doc,
        "Os três custos mensais somam R$ 13.000,00 por mês. Nos 21 meses do período "
        "(julho de 2026 a março de 2028), o custeio mensal chega a R$ 273.000,00 "
        "que, acrescidos dos R$ 42.000,00 de equipamentos da equipe contratada, "
        "resultam em R$ 315.000,00.",
        recuo_esquerda=1.25,
    )

    # --- Seção em paisagem: os quadros largos cabem em tamanho natural
    nova_secao_paisagem(doc)

    titulo_secao(doc, "RESUMO DAS ATIVIDADES")
    tabela = nova_tabela(
        doc,
        COL_RESUMO_PAISAGEM,
        borda_cor="bordaCinza",
        borda_pt=0.5,
        borda_externa_cor="bordaEscura",
        margem_celula_cm=0.1,
    )
    cabecalho_resumo(tabela)
    linha(
        tabela,
        [
            celula("01", fundo="azulTitulo", negrito=True, cor="branco", tamanho=10,
                   alinhamento="centro"),
            celula("REDUZIR O TEMPO OPERACIONAL COM AGENTES DE INTELIGÊNCIA "
                   "ARTIFICIAL", fundo="azulTitulo", negrito=True, cor="branco",
                   tamanho=10),
            celula("Redução mínima de 20% no tempo médio de execução dos processos "
                   "apoiados por IA", fundo="azulTitulo", negrito=True, cor="branco",
                   tamanho=10),
            celula("jul/26", fundo="azulTitulo", negrito=True, cor="branco",
                   tamanho=10, alinhamento="centro"),
            celula("mar/28", fundo="azulTitulo", negrito=True, cor="branco",
                   tamanho=10, alinhamento="centro"),
            celula("Walmir Caminhas - Diretor de TI", fundo="azulTitulo",
                   negrito=True, cor="branco", tamanho=10, alinhamento="centro"),
        ],
    )
    linha_subatividade_resumo(
        tabela,
        "01.a",
        "Investigar e implementar agentes de inteligência artificial nos processos "
        "operacionais em que essa é a solução adequada.",
        "Processos identificados e agentes implementados, com redução do tempo "
        "médio de execução.",
        "jul/26",
        "mar/28",
        "Walmir Caminhas - Diretor de TI",
        zebra=True,
    )
    linha_subatividade_resumo(
        tabela,
        "01.b",
        "Desenvolver os agentes de inteligência artificial, com equipe contratada, "
        "ambiente de execução em nuvem e ferramentas de inteligência artificial de "
        "terceiros.",
        "Agentes desenvolvidos e em operação nos processos selecionados.",
        "jul/26",
        "mar/28",
        "Walmir Caminhas - Diretor de TI",
        zebra=True,
    )

    titulo_secao(doc, "CRONOGRAMA FÍSICO-FINANCEIRO")
    paragrafo(
        doc,
        "Atividade posicionada nos trimestres correspondentes ao seu início e ao seu "
        "término, com os valores agrupados por ano. A alínea a) não tem custo "
        "próprio. A alínea b) concentra todo o investimento: R$ 42.000,00 de "
        "equipamentos para a equipe contratada, em 2026, mais R$ 13.000,00 por mês "
        "de ambiente em nuvem, uso de inteligência artificial de terceiros e agente "
        "de terceiro de apoio ao desenvolvimento, distribuídos pelos 21 meses do "
        "período. O valor de 2028 corresponde ao 1º trimestre, quando a atividade se "
        "encerra.",
    )

    larguras = [0.81, 7.12] + [1.269] * 14
    tabela = nova_tabela(
        doc, larguras, borda_cor="bordaCinza", borda_pt=0.4, margem_celula_cm=0.08
    )
    cabecalho_cronograma(
        tabela,
        [
            ("2026", ["3º tri", "4º tri"]),
            ("2027", ["1º tri", "2º tri", "3º tri", "4º tri"]),
            ("2028", ["1º tri", "2º tri", "3º tri", "4º tri"]),
            ("2029", ["1º tri", "2º tri", "3º tri", "4º tri"]),
        ],
        tamanho_ano=12,
        tamanho_tri=11,
    )
    linha(tabela, [
        celula("01 - Reduzir o tempo operacional com agentes de IA", span=2,
               fundo="cinzaMedio", negrito=True, tamanho=12),
        celula("R$ 120.000,00", span=2, fundo="cinzaMedio", negrito=True,
               tamanho=10, alinhamento="centro"),
        celula("R$ 156.000,00", span=4, fundo="cinzaMedio", negrito=True,
               tamanho=10, alinhamento="centro"),
        celula("R$ 39.000,00", span=4, fundo="cinzaMedio", negrito=True,
               tamanho=10, alinhamento="centro"),
    ] + vazias(4), alinhamento_vertical="centro")
    for descricao, valores in (
        ("a) Investigar e implementar agentes de IA nos processos",
         ("R$ 0,00", "R$ 0,00", "R$ 0,00")),
        ("b) Desenvolver os agentes de IA (equipe, nuvem e IA de terceiros)",
         ("R$ 120.000,00", "R$ 156.000,00", "R$ 39.000,00")),
    ):
        linha(tabela, [
            celula(),
            celula(descricao, tamanho=10),
            celula(valores[0], span=2, fundo="cinzaE8", tamanho=10,
                   alinhamento="centro"),
            celula(valores[1], span=4, fundo="cinzaE8", tamanho=10,
                   alinhamento="centro"),
            celula(valores[2], span=4, fundo="cinzaE8", tamanho=10,
                   alinhamento="centro"),
        ] + vazias(4), alinhamento_vertical="centro")
    linha_total(tabela, 14, "R$ 315.000,00")

    rodape_link(doc)

    return doc, "PE 26-30 - PPA - Plano de Acao - Implementar Agentes de IA.docx"


# ------------------------------------------------------------------
# Plano 3: Aprimorar a Segurança da Informação
# ------------------------------------------------------------------

def plano_seguranca():
    doc = novo_documento(espaco_paragrafo_pt=12, entrelinha=1.5)

    titulo_documento(
        doc, "PLANO DE AÇÃO", "Planejamento Estratégico Fundep - Gestão 2026 a 2030"
    )

    titulo_secao(doc, "IDENTIFICAÇÃO DO PLANO")
    tabela_identificacao(
        doc,
        [
            ("DIRETRIZ", "Transformação Digital"),
            ("AÇÃO", "Aprimorar a Segurança da Informação"),
            ("RESPONSÁVEL", "DPO - Flavia Carolina de Oliveira Andrade"),
            (
                "FORTALEZA/PROBLEMA",
                "Fragilidade no atendimento legal em segurança de dados",
            ),
            (
                "OBJETIVO(S) ESTRATÉGICO(S) VINCULADO(S)",
                "Fortalecimento da Segurança da Informação com Tecnologias Digitais",
            ),
            (
                "META(S) OPERACIONAL(IS) VINCULADA(S)",
                "Não registrar nenhum incidente de segurança da informação a partir "
                "de 2028",
            ),
            ("PERÍODO DE ACOMPANHAMENTO", "Agosto de 2026 a dezembro de 2029"),
            (
                "ÁREAS ENVOLVIDAS",
                "Conselho Diretor, Tecnologia da Informação, CIAs, Prestação de "
                "Contas, Compras, Importação, Infraestrutura e Logística, Gestão de "
                "Documentos, Administração de Pessoal, Gestão de Pessoas, Negócios e "
                "Parcerias, Financeiro e Contabilidade, Compliance, Jurídico, "
                "Comunicação e Marketing, Inovação, Engenharia, Gestão de Programas, "
                "Gestão de Concursos.",
            ),
        ],
    )

    titulo_secao(doc, "JUSTIFICATIVA")
    paragrafo(
        doc,
        "A Fundep realiza milhares de transações com uso de dados por dia. Além "
        "disso, os avanços tecnológicos acelerados com o uso de ferramentas de "
        "inteligência artificial aumentam exponencialmente as ameaças cibernéticas, "
        "reforçando a importância de manter os mecanismos de segurança da informação "
        "da instituição alinhados a padrões adequados de conformidade. Nesse "
        "sentido, a contratação de uma auditoria externa, integrante do grupo das "
        "seis maiores empresas do setor, justifica-se pela necessidade de uma "
        "avaliação independente sobre a maturidade dos controles internos, processos "
        "e ativos tecnológicos da instituição.",
    )
    paragrafo(
        doc,
        "Além de identificar vulnerabilidades e apontar melhorias para a mitigação "
        "de riscos operacionais e financeiros, a avaliação por uma entidade externa "
        "contribui para a governança corporativa, a transparência perante parceiros "
        "e órgãos reguladores, e para a reputação institucional, ao evidenciar o "
        "compromisso da instituição com a proteção de dados e a continuidade dos "
        "negócios.",
    )
    paragrafo(
        doc,
        "Esta ação está associada aos itens do diagnóstico FOFA da Fundep "
        "relacionados a seguir, mapeados pelos gerentes da instituição. O número que "
        "antecede cada item é o seu identificador no diagnóstico.",
    )

    lista_fofa(doc, "Forças", [
        "102. Reputação institucional consolidada e elevada credibilidade perante "
        "stakeholders.",
        "110. Gestão tecnológica e de informações integrada e adaptável.",
        "111. Vantagem competitiva em soluções de TI proprietárias.",
        "112. Programa de Integridade constituído.",
    ])
    lista_fofa(doc, "Oportunidades", [
        "206. Tecnologias digitais de Inteligência Artificial emergentes aplicadas a "
        "processos.",
    ])
    lista_fofa(doc, "Fraquezas", [
        "303. Falta de gestão de contingência.",
        "311. Confiabilidade dos dados financeiros e extratos.",
        "312. Instrumentos de governança desatualizados ou inexistentes.",
        "316. Falta de atendimento integral à LGPD.",
        "317. Segurança da informação e proteção de dados deficientes.",
        "321. Falta de gestão de riscos.",
        "331. Falta de integração dos fluxos de documentos.",
        "338. Avanço tecnológico mais rápido que a capacidade de adaptação interna.",
    ])
    lista_fofa(doc, "Ameaças", [
        "405. Mudanças em outras legislações que impactam as Fundações de Apoio.",
        "407. Exigências crescentes de órgãos de controle, aumentando a complexidade "
        "operacional.",
    ])
    lista_fofa(doc, "Fortalezas", [
        "6101. Referência e pioneirismo técnico em tecnologias digitais em gestão e "
        "atividades entre as Fundações de Apoio.",
        "6102. Constante evolução da fronteira tecnológica para tecnologias "
        "digitais, possibilitando aumento da eficiência operacional.",
    ])
    lista_fofa(doc, "Problemas", [
        "6201. Fragilidade no atendimento legal em segurança de dados.",
        "6202. As necessidades técnicas digitais e de segurança são subestimadas "
        "comparadas às vontades dos usuários.",
    ])

    titulo_secao(doc, "ATIVIDADES")

    titulo_atividade(
        doc, "01. Contratar consultoria de uma das seis maiores empresas de auditoria"
    )
    paragrafo(
        doc,
        "Antes de passar por auditoria, será necessário compreender a situação atual "
        "da fundação, por meio da contratação de uma consultoria especializada, que "
        "elaborará relatório detalhado com os pontos fracos e fortes.",
    )
    item_lista(doc, "a)", "Definir empresa por meio de pesquisa.", espaco_depois=6)
    paragrafo(
        doc,
        "Realizar pesquisa para escolha da melhor empresa cujo escopo de atuação se "
        "adeque à realidade da fundação.",
        recuo_esquerda=1.25,
        espaco_depois=12,
    )
    item_lista(doc, "b)", "Acompanhar atuação da empresa na Fundep.", espaco_depois=6)
    paragrafo(
        doc,
        "Formar equipe interna composta pelas áreas interessadas (TI, DPO e "
        "Compliance) para prestar informações à consultoria e acompanhar as "
        "entregas.",
        recuo_esquerda=1.25,
    )

    titulo_atividade(
        doc, "02. Contratar, instalar e formar equipe de Segurança da Informação"
    )
    paragrafo(
        doc,
        "Para acompanhamento das atividades da consultoria, e para manutenção das "
        "atividades de cibersegurança na Fundep, será contratada, instalada e "
        "formada uma equipe de Segurança da Informação, que irá atuar preventiva e "
        "ativamente.",
    )
    paragrafo(
        doc,
        "Custo estimado de R$ 30.000,00 para a instalação da equipe, referente a "
        "equipamentos, mobiliário e preparação da sala de trabalho.",
    )
    item_lista(doc, "a)", "Contratar e instalar os profissionais.", espaco_depois=6)
    paragrafo(
        doc,
        "Formar equipe de Segurança da Informação, vinculada à Gerência de TI, por "
        "meio da abertura de vagas e seleção de candidatos, e instalá-la em "
        "condições de trabalho, com as estações de trabalho e a sala de trabalho "
        "preparadas para o início das atividades.",
        recuo_esquerda=1.25,
        espaco_depois=12,
    )
    item_lista(doc, "b)", "Estruturar as atividades e delegar funções.",
               espaco_depois=6)
    paragrafo(
        doc,
        "Definição das funções da equipe de Segurança da Informação, para realização "
        "de monitoramento contínuo, prevenção e atuação em incidentes, gestão de "
        "vulnerabilidades, e outras atividades referentes à segurança da informação.",
        recuo_esquerda=1.25,
        espaco_depois=12,
    )
    item_lista(doc, "c)", "Custo salarial.", espaco_depois=6)
    paragrafo(
        doc,
        "Custo salarial da equipe de Segurança da Informação ao longo do período, "
        "conforme a composição prevista para cada ano:",
        recuo_esquerda=1.25,
        espaco_depois=6,
    )

    tabela = nova_tabela(
        doc, [1.9, 8.0, 3.6], borda_cor="bordaCinza", borda_pt=0.5
    )
    linha(
        tabela,
        [
            celula("Ano", fundo="azulCabecalho", negrito=True, alinhamento="centro"),
            celula("Composição da equipe", fundo="azulCabecalho", negrito=True),
            celula("Custo salarial", fundo="azulCabecalho", negrito=True,
                   alinhamento="direita"),
        ],
        alinhamento_vertical="centro",
    )
    for ano, composicao, custo, zebra in (
        ("2026", "1 Especialista + 1 Pleno", "R$ 375.023,46", True),
        ("2027", "1 Especialista + 1 Pleno (consolidação)", "R$ 375.023,46", False),
        ("2028", "1 Especialista + 2 Plenos", "R$ 518.534,40", True),
        ("2029", "1 Especialista + 2 Plenos + 1 Júnior", "R$ 623.875,98", False),
    ):
        fundo = "cinzaClaro" if zebra else None
        linha(
            tabela,
            [
                celula(ano, fundo=fundo, alinhamento="centro"),
                celula(composicao, fundo=fundo),
                celula(custo, fundo=fundo, alinhamento="direita"),
            ],
            alinhamento_vertical="centro",
        )
    linha(
        tabela,
        [
            celula("Total de 2026 a 2029", span=2, fundo="azulVivo", negrito=True,
                   cor="branco"),
            celula("R$ 1.892.457,30", fundo="azulVivo", negrito=True, cor="branco",
                   alinhamento="direita"),
        ],
        alinhamento_vertical="centro",
    )
    paragrafo(doc, espaco_depois=6)
    paragrafo(
        doc,
        "(a preencher: custo salarial de 2030, com a composição da equipe prevista "
        "para o ano.)",
        recuo_esquerda=1.25,
    )

    titulo_atividade(doc, "03. Implementar melhorias")
    paragrafo(
        doc,
        "O relatório realizado pela consultoria embasará as tratativas a serem "
        "implementadas nessa atividade. Cada ponto levantado pela consultoria "
        "contratada poderá gerar novos planos de ação, a depender da complexidade da "
        "atividade.",
    )
    item_lista(doc, "a)", "Atuação no relatório da consultoria.", espaco_depois=6)
    paragrafo(
        doc,
        "Atendimento aos pontos levantados pela consultoria, com a equipe de "
        "segurança da informação atuando tecnicamente, e DPO, Compliance e TI "
        "atuando estrategicamente na revisão de processos, elaboração de políticas e "
        "revisão de procedimentos que se fizerem necessários.",
        recuo_esquerda=1.25,
        espaco_depois=12,
    )
    item_lista(doc, "b)", "Validação com Conselho Diretor para auditoria.",
               espaco_depois=6)
    paragrafo(
        doc,
        "Levar relatório da consultoria ao Conselho Diretor, com a implementação das "
        "melhorias realizadas, para preparação para auditoria.",
        recuo_esquerda=1.25,
    )

    titulo_atividade(doc, "04. Realizar auditoria")
    paragrafo(
        doc,
        "Após implementar as correções apontadas pela consultoria, a fundação estará "
        "habilitada a ser auditada por uma das seis grandes empresas. O foco da "
        "auditoria será a segurança da informação e a conformidade com a Lei Geral "
        "de Proteção de Dados Pessoais (LGPD).",
    )
    for marcador, titulo in (
        ("a)", "Orçar e realizar contratação."),
        ("b)", "Apoiar o trabalho da auditoria."),
        ("c)", "Receber parecer da auditoria."),
    ):
        item_lista(doc, marcador, titulo, espaco_depois=6)
        paragrafo(
            doc,
            "(a preencher: detalhamento da subatividade, com as áreas envolvidas.)",
            recuo_esquerda=1.25,
            espaco_depois=12,
        )

    titulo_atividade(
        doc,
        "05. Controle e monitoramento das comunicações institucionais com os "
        "públicos externos",
    )
    paragrafo(
        doc,
        "A Fundep mantém comunicação diária com coordenadores, fornecedores, alunos, "
        "candidatos e demais públicos externos por meio de aplicativos de mensagens "
        "instalados em mais de 200 dispositivos. Hoje essas comunicações não são "
        "registradas nem monitoradas, o que representa risco à segurança dos dados, "
        "à conformidade com a LGPD e à continuidade do atendimento. Será implantada "
        "solução corporativa de controle e monitoramento dessas conversas, com foco "
        "inicial na comunicação com os coordenadores.",
    )
    paragrafo(
        doc,
        "Além da segurança dos dados, objetivo principal da atividade, a solução "
        "gera ganhos colaterais relevantes: controle gerencial do atendimento, "
        "levantamento e histórico de dados das interações, painel de controle, "
        "acompanhamento de mensagens sem resposta e repasse estruturado de carteira "
        "entre os analistas das CIAs.",
    )
    paragrafo(
        doc,
        "Custo estimado de R$ 10.000,00 por mês, do segundo semestre de 2026 ao "
        "primeiro trimestre de 2030, totalizando R$ 440.000,00 no período. Caso o "
        "projeto seja firmado, esse valor passa a constituir despesa fixa da Fundep.",
    )
    item_lista(doc, "a)", "Contratar e implantar a solução de monitoramento.",
               espaco_depois=6)
    paragrafo(
        doc,
        "Pesquisa e contratação de plataforma de controle e monitoramento das "
        "comunicações institucionais com os públicos externos, com implantação nos "
        "dispositivos utilizados na comunicação com os coordenadores.",
        recuo_esquerda=1.25,
        espaco_depois=12,
    )
    item_lista(doc, "b)", "Acompanhar, evoluir e ampliar o controle.", espaco_depois=6)
    paragrafo(
        doc,
        "Monitoramento contínuo das comunicações, operação do painel de controle e "
        "dos indicadores de atendimento, e ampliação gradual do escopo para os "
        "demais públicos atendidos pela fundação.",
        recuo_esquerda=1.25,
    )

    # --- Seção em paisagem: resumo das atividades e cronograma
    nova_secao_paisagem(doc)

    titulo_secao(doc, "RESUMO DAS ATIVIDADES")
    tabela = nova_tabela(
        doc,
        COL_RESUMO_PAISAGEM,
        borda_cor="bordaCinza",
        borda_pt=0.5,
        borda_externa_cor="bordaEscura",
        margem_celula_cm=0.1,
    )
    cabecalho_resumo(tabela)

    linha_atividade_resumo(
        tabela,
        "01",
        "CONTRATAR CONSULTORIA DE UMA DAS SEIS MAIORES EMPRESAS DE AUDITORIA",
        "08/26",
        "02/27",
        "James Mota",
    )
    linha_subatividade_resumo(
        tabela, "01.a", "Definir empresa por meio de pesquisa", "Empresa escolhida",
        "08/26", "08/26", "James", zebra=True,
    )
    linha_subatividade_resumo(
        tabela, "01.b", "Acompanhar atuação da empresa na fundação",
        "Entrega de relatório de consultoria", "09/26", "02/27",
        "James, André, Flávia", zebra=False,
    )
    linha_atividade_resumo(
        tabela,
        "02",
        "CONTRATAR, INSTALAR E FORMAR EQUIPE DE SEGURANÇA DA INFORMAÇÃO",
        "08/26",
        "12/29",
        "André Oliveira",
    )
    linha_subatividade_resumo(
        tabela, "02.a", "Contratar e instalar os profissionais",
        "Equipe formada e instalada, com estações de trabalho e sala prontas",
        "08/26", "10/26", "André, Débora", zebra=True,
    )
    linha_subatividade_resumo(
        tabela, "02.b", "Estruturar atividades e delegar funções",
        "Plano de atividades desenvolvido.", "10/26", "10/26", "André", zebra=False,
    )
    linha_subatividade_resumo(
        tabela, "02.c", "Custo salarial da equipe",
        "Equipe mantida e ampliada, de 2 profissionais em 2026 a 4 em 2029",
        "08/26", "12/29", "André, Débora", zebra=True,
    )
    linha_atividade_resumo(
        tabela, "03", "IMPLEMENTAR MELHORIAS", "02/27", "05/29", "André Oliveira",
    )
    linha_subatividade_resumo(
        tabela, "03.a", "Atuação no relatório da consultoria",
        "Relatório da consultoria com melhorias realizadas", "02/27", "04/29",
        "André, Flávia", zebra=True,
    )
    linha_subatividade_resumo(
        tabela, "03.b", "Validação com Conselho Diretor para auditoria",
        "Relatório validado", "04/29", "05/29", "James, Flávia, André", zebra=False,
    )
    linha_atividade_resumo(
        tabela, "04", "REALIZAR AUDITORIA", "06/29", "01/30", "James Mota",
    )
    linha_subatividade_resumo(
        tabela, "04.a", "Orçar e realizar contratação", "Empresa contratada.",
        "06/29", "07/29", "James, Flávia", zebra=True,
    )
    linha_subatividade_resumo(
        tabela, "04.b", "Apoiar o trabalho da auditoria", "", "07/29", "12/29",
        "André, James, Flávia", zebra=False,
    )
    linha_subatividade_resumo(
        tabela, "04.c", "Receber parecer da auditoria", "Relatório recebido.",
        "01/30", "01/30", "James, Flávia", zebra=True,
    )
    linha_atividade_resumo(
        tabela,
        "05",
        "CONTROLAR E MONITORAR AS COMUNICAÇÕES INSTITUCIONAIS COM OS PÚBLICOS "
        "EXTERNOS",
        "08/26",
        "03/30",
        "André Oliveira",
    )
    linha_subatividade_resumo(
        tabela, "05.a",
        "Contratar e implantar a solução de monitoramento, com foco inicial na "
        "comunicação com os coordenadores.",
        "Solução contratada e implantada em 2026.", "08/26", "12/26",
        "André, Flávia", zebra=True,
    )
    linha_subatividade_resumo(
        tabela, "05.b",
        "Acompanhar, evoluir e ampliar o controle das comunicações.",
        "Acompanhamento contínuo das conversas, com painel de controle, "
        "monitoramento de falta de resposta e repasse de carteira entre analistas "
        "das CIAs.",
        "01/27", "03/30", "André, Flávia", zebra=False,
    )

    # O cronograma começa em página própria, como no PDF
    quebra_pagina(doc)
    titulo_secao(doc, "CRONOGRAMA FÍSICO-FINANCEIRO")
    paragrafo(
        doc,
        "Atividades e subatividades posicionadas nos trimestres correspondentes ao "
        "seu início e ao seu término, conforme o resumo acima. A régua vai do 3º "
        "trimestre de 2026 ao 4º trimestre de 2029, acompanhando o período de "
        "acompanhamento do plano, de modo que as subatividades 04.c e 05.b, cujo "
        "término se dá em 2030, aparecem lançadas no último trimestre exibido. Os "
        "R$ 120.000,00 da atividade 01 referem-se à contratação da consultoria e "
        "estão lançados na subatividade 01.b, que é a que acompanha a execução do "
        "contrato. Na atividade 02, os R$ 30.000,00 da subatividade 02.a cobrem os "
        "equipamentos, o mobiliário e a preparação da sala de trabalho da equipe de "
        "Segurança da Informação, e os R$ 1.892.457,30 da subatividade 02.c "
        "correspondem ao custo salarial da equipe de 2026 a 2029, lançado ano a ano "
        "conforme a composição prevista para cada exercício. Os R$ 120.000,00 da "
        "atividade 04 referem-se à contratação da auditoria e estão lançados na "
        "subatividade 04.a. Os R$ 440.000,00 da atividade 05 correspondem a "
        "R$ 10.000,00 por mês, do segundo semestre de 2026 ao 1º trimestre de 2030. "
        "O valor da atividade 03 ainda será definido, de modo que o total "
        "apresentado é parcial. (a preencher: valor da atividade 03, com a origem e "
        "o critério de composição.)",
        espaco_depois=8,
    )

    larguras = [0.85, 6.16] + [1.335] * 14
    tabela = nova_tabela(
        doc, larguras, borda_cor="bordaCinza", borda_pt=0.4, margem_celula_cm=0.08
    )
    cabecalho_cronograma(
        tabela,
        [
            ("2026", ["3º tri", "4º tri"]),
            ("2027", ["1º tri", "2º tri", "3º tri", "4º tri"]),
            ("2028", ["1º tri", "2º tri", "3º tri", "4º tri"]),
            ("2029", ["1º tri", "2º tri", "3º tri", "4º tri"]),
        ],
        tamanho_ano=12,
        tamanho_tri=11,
    )

    def rotulo(texto):
        return celula(texto, span=2, fundo="cinzaMedio", negrito=True, tamanho=12)

    def faixa(texto, span, fundo, negrito=False, tamanho=10):
        return celula(texto, span=span, fundo=fundo, negrito=negrito,
                      tamanho=tamanho, alinhamento="centro")

    def alinea(texto):
        return [celula(), celula(texto, tamanho=10)]

    linha(tabela, [
        rotulo("01 - Contratar consultoria de auditoria"),
        faixa("R$ 120.000,00", 3, "cinzaMedio", negrito=True, tamanho=12),
    ] + vazias(11), alinhamento_vertical="centro")
    linha(tabela, alinea("a) Definir empresa por meio de pesquisa") + [
        celula(fundo="cinzaE8"),
    ] + vazias(13), alinhamento_vertical="centro")
    linha(tabela, alinea("b) Acompanhar atuação da empresa na fundação") + [
        faixa("R$ 120.000,00", 3, "cinzaE8"),
    ] + vazias(11), alinhamento_vertical="centro")

    linha(tabela, [
        rotulo("02 - Contratar e instalar equipe de Segurança da Informação"),
        faixa("R$ 1.922.457,30", 14, "cinzaMedio", negrito=True, tamanho=12),
    ], alinhamento_vertical="centro")
    linha(tabela, alinea("a) Contratar e instalar os profissionais") + [
        faixa("R$ 30.000,00", 2, "cinzaE8"),
    ] + vazias(12), alinhamento_vertical="centro")
    linha(tabela, alinea("b) Estruturar atividades e delegar funções") + [
        celula(),
        celula(fundo="cinzaE8"),
    ] + vazias(12), alinhamento_vertical="centro")
    linha(tabela, alinea("c) Custo salarial da equipe") + [
        faixa("R$ 375.023,46", 2, "cinzaE8"),
        faixa("R$ 375.023,46", 4, "cinzaE8"),
        faixa("R$ 518.534,40", 4, "cinzaE8"),
        faixa("R$ 623.875,98", 4, "cinzaE8"),
    ], alinhamento_vertical="centro")

    linha(tabela, [
        rotulo("03 - Implementar melhorias"),
    ] + vazias(2) + [
        celula(span=10, fundo="cinzaMedio"),
    ] + vazias(2), alinhamento_vertical="centro")
    linha(tabela, alinea("a) Atuação no relatório da consultoria") + vazias(2) + [
        celula(span=10, fundo="cinzaE8"),
    ] + vazias(2), alinhamento_vertical="centro")
    linha(tabela, alinea(
        "b) Validação com Conselho Diretor para auditoria") + vazias(11) + [
        celula(fundo="cinzaE8"),
    ] + vazias(2), alinhamento_vertical="centro")

    linha(tabela, [
        rotulo("04 - Realizar auditoria"),
    ] + vazias(11) + [
        faixa("R$ 120.000,00", 3, "cinzaMedio", negrito=True, tamanho=12),
    ], alinhamento_vertical="centro")
    linha(tabela, alinea("a) Orçar e realizar contratação") + vazias(11) + [
        faixa("R$ 120.000,00", 2, "cinzaE8"),
        celula(),
    ], alinhamento_vertical="centro")
    linha(tabela, alinea("b) Apoiar o trabalho da auditoria") + vazias(12) + [
        celula(span=2, fundo="cinzaE8"),
    ], alinhamento_vertical="centro")
    linha(tabela, alinea("c) Receber parecer da auditoria") + vazias(13) + [
        celula(fundo="cinzaE8"),
    ], alinhamento_vertical="centro")

    linha(tabela, [
        rotulo("05 - Monitorar as comunicações com os públicos externos"),
        faixa("R$ 440.000,00 (R$ 10.000,00/mês)", 14, "cinzaMedio", negrito=True,
              tamanho=12),
    ], alinhamento_vertical="centro")
    linha(tabela, alinea("a) Contratar e implantar a solução de monitoramento") + [
        faixa("R$ 50.000,00", 2, "cinzaE8"),
    ] + vazias(12), alinhamento_vertical="centro")
    linha(tabela, alinea(
        "b) Acompanhar, evoluir e ampliar o controle das comunicações") + vazias(2) + [
        faixa("R$ 390.000,00", 12, "cinzaE8"),
    ], alinhamento_vertical="centro")

    linha_total(
        tabela, 14, "R$ 2.602.457,30 (parcial: valor da atividade 03 a definir)"
    )

    rodape_link(doc)

    nome = (
        "PE 26-30 - PPA - Plano de Acao - Aprimorar a Seguranca da informacao.docx"
    )
    return doc, nome


# ------------------------------------------------------------------

def main():
    for gerador in (plano_gpf, plano_agentes_ia, plano_seguranca):
        doc, nome = gerador()
        mantem_tabelas_inteiras(doc)
        destino = RAIZ / nome
        doc.save(str(destino))
        print("gerado: %s" % nome)


if __name__ == "__main__":
    main()
