# -*- coding: utf-8 -*-
"""Laminas de setembro/2026 -- fechamento de agosto/2026.

ATENCAO: as direcoes ("Alta em agosto" / "Queda em agosto") vem das abas
now_last_* da planilha, nao dos numeros citados pelos analistas -- ver o
relatorio de entrega. Os analistas usam janela de 30 dias corridos; a lamina
usa o mes calendario. Em 16 dos 38 nomes o sinal difere.
"""

# --- comentarios reaproveitados entre carteiras -----------------------------
PETR4 = ("Alta em agosto, acima do Ibovespa. A Petrobras segue entre nossas preferências no setor: em "
         "nossa visão, a companhia oferece uma das melhores relações entre risco e retorno diante da "
         "volatilidade associada ao conflito entre Estados Unidos e Irã, e os retornos permanecem atrativos "
         "caso os preços do petróleo se mantenham em patamares elevados.")

PRIO3 = ("Alta em agosto, acima do Ibovespa. A PRIO segue como nossa principal preferência em óleo e gás: "
         "os preços do petróleo permanecem em patamares relativamente elevados e esperamos forte geração de "
         "caixa em 2027. A companhia mantém níveis elevados de produção, impulsionados sobretudo pela entrada "
         "em regime do campo de Wahoo e por Peregrino.")

VALE3 = ("Alta em agosto, acima do Ibovespa, apesar de fundamentos ainda frágeis para o minério de ferro, "
         "com demanda fraca na China e oferta confortável. Ao longo do mês, as discussões apontaram um mercado "
         "de minério relativamente equilibrado, enquanto o bom desempenho de Metais Básicos trouxe uma leitura "
         "mais favorável para essa parte da tese.")

ITUB4 = ("Queda em agosto, abaixo do Ibovespa e em linha com os pares, refletindo o cenário macroeconômico "
         "desafiador, que pode pesar sobre o crescimento da carteira e a qualidade do crédito. Seguimos vendo o "
         "banco bem posicionado para manter retornos sobre o patrimônio acima dos pares e do próprio custo de "
         "capital, com riscos altistas relevantes para a tese.")

AXIA3 = ("Alta em agosto, acima do Ibovespa. As chuvas recentes no Sul recuperaram os níveis dos "
         "reservatórios e reduziram a pressão sobre os preços de energia, movimento que pode se intensificar "
         "caso o El Niño traga maior precipitação na região. A companhia segue entre as elétricas brasileiras "
         "preferidas dos investidores estrangeiros.")

ORVR3 = ("Queda em agosto, abaixo do Ibovespa. O resultado do 2T26 veio, de maneira geral, em linha com as "
         "expectativas, impulsionado pelo desempenho sólido dos aterros e pelo avanço da planta de biometano de "
         "Jaboatão, mas contrabalançado pela Vital, pressionada por custos da conversão da frota para biometano.")

# --- Top Ações --------------------------------------------------------------
ACOES = {
    'PETR4': PETR4,
    'PRIO3': PRIO3,
    'GGBR4': ("Queda em agosto, abaixo do Ibovespa, refletindo as preocupações com uma possível "
              "flexibilização das tarifas sobre o aço canadense nos Estados Unidos, que elevaria a competição de "
              "importados e aceleraria a normalização das margens na América do Norte. O movimento ocorreu apesar "
              "dos resultados sólidos do 2T26 e das revisões positivas de lucro."),
    'VALE3': VALE3,
    'EMBJ3': ("Alta em agosto, acima do Ibovespa, acompanhada de resultados robustos no 2T26, com "
              "crescimento disseminado das receitas, melhora das margens recorrentes e forte geração de caixa, que "
              "levaram a companhia a elevar a projeção de margem operacional. Segue nossa preferência em Bens de "
              "Capital, apoiada pela melhora da qualidade da carteira de pedidos."),
    'TOTS3': ("Alta em agosto, acima do Ibovespa, após um mês de forte volatilidade. O papel chegou a recuar "
              "depois da divulgação dos resultados do 2T26, pressionado pelas saídas de capital estrangeiro e pela "
              "deterioração do sentimento em relação ao mercado brasileiro, mas recuperou integralmente as perdas "
              "iniciais na segunda metade do mês."),
    'IGTI11': ("Queda em agosto, abaixo do Ibovespa. Mantemos visão positiva, apoiada na qualidade e "
               "produtividade dos ativos, na maior exposição ao segmento de alta renda e na perspectiva de "
               "crescimento acima da média do setor. Apesar do resultado financeiro do 2T26 abaixo do esperado e da "
               "maior alavancagem, os indicadores operacionais seguiram saudáveis."),
    'CURY3': ("Incluímos a Cury na carteira em setembro, nossa principal escolha entre as construtoras de "
              "baixa renda. A tese segue apoiada por execução consistente, elevada rentabilidade, forte geração de "
              "caixa e perfil atrativo de dividendos."),
    'RDOR3': ("Alta em agosto, acima do Ibovespa, refletindo um resultado sólido no 2T26, marcado pela "
              "melhora da sinistralidade na SulAmérica e pelo bom desempenho hospitalar. Reiteramos nossa visão "
              "positiva para a companhia."),
    'AXIA3': ("Incluímos a Axia Energia na carteira em setembro. A forte visibilidade da companhia entre os "
              "investidores estrangeiros e um potencial aumento da alocação em Brasil podem dar suporte às ações. À "
              "frente, o El Niño pode elevar os riscos de ondas de calor e a volatilidade da geração eólica, "
              "sustentando os preços de energia no 4T26."),
    'SBSP3': ("Queda em agosto, abaixo do Ibovespa. Por ser um papel com grande participação de investidores "
              "institucionais estrangeiros, sofreu com a troca de alocação em emergentes por posições ligadas à "
              "inteligência artificial. O resultado do 2T26 também ficou abaixo do esperado, com temperaturas mais "
              "baixas reduzindo volumes e receita, além de maior pressão de custos."),
    'ITUB4': ITUB4,
    'BPAC11': ("Queda em agosto, abaixo do Ibovespa, refletindo um cenário mais desafiador para o mercado de "
               "capitais e a preocupação dos investidores com o avanço do crédito. Seguimos confiantes na "
               "sustentabilidade de retornos sobre o patrimônio acima dos pares e na capacidade de crescimento "
               "saudável do banco, o que sustenta a manutenção do papel na carteira."),
    'ROXO34': ("Alta em agosto, acima do Ibovespa, após um 2T26 acima das expectativas do mercado, com forte "
               "recuperação da margem financeira ajustada ao risco. Seguimos vendo uma combinação de crescimento "
               "estrutural, melhor visibilidade de lucros e múltiplo atrativo."),
}

# --- Top Dividendos ---------------------------------------------------------
DIVS = {
    'PETR4': PETR4,
    'PRIO3': PRIO3,
    'VALE3': VALE3,
    'VIVT3': ("Queda em agosto, abaixo do Ibovespa. Apesar dos resultados sólidos no 2T26, a reação do "
              "mercado foi negativa, já que as expectativas dos investidores estavam acima do que foi entregue. O "
              "movimento foi agravado pela percepção de competição mais intensa no setor e pela continuidade das "
              "saídas de capital estrangeiro da Bolsa."),
    'ALOS3': ("Desempenho praticamente estável em agosto, em linha com o Ibovespa. O resultado do 2T26 "
              "reforçou nossa visão positiva, com crescimento de receitas e EBITDA e fundamentos operacionais "
              "saudáveis. Mantemos a posição pela combinação de geração de caixa, potencial de retorno ao acionista "
              "e oportunidades de reciclagem do portfólio."),
    'CSMG3': ("Queda em agosto, abaixo do Ibovespa. O resultado do 2T26 veio abaixo das expectativas, "
              "pressionado por despesas operacionais e custos variáveis elevados, apesar da receita em linha. A "
              "companhia avançou na repactuação dos contratos municipais e formalizou acordo para expansão dos "
              "serviços de esgoto, reforçando a perspectiva de crescimento de ativos."),
    'AXIA3': AXIA3,
    'CPLE3': ("Alta em agosto, acima do Ibovespa. O resultado do 2T26 veio em linha com as expectativas do "
              "mercado: em geração e transmissão, os efeitos favoráveis da modulação de energia foram parcialmente "
              "neutralizados pelo maior custo de compra associado à restrição de geração, enquanto na distribuição "
              "o maior volume foi compensado pelo avanço das despesas."),
    'ITUB4': ITUB4,
    'B3SA3': ("Alta em agosto, acima do Ibovespa, apesar do cenário macroeconômico mais desafiador. O carrego "
              "segue interessante, o que motiva a permanência do papel na carteira de dividendos."),
    'CXSE3': ("Queda em agosto, abaixo do Ibovespa, corrigindo após um desempenho forte nos meses anteriores. "
              "Com um múltiplo que exige menos, uma operação que segue tracionada e retorno em dividendos próximo "
              "de 8%, preferimos manter o nome na carteira."),
}

# --- Top Small Caps (serve tambem a Small Caps Clear) -----------------------
SMALL = {
    'BMOB3': ("Alta em agosto, acima do Ibovespa, em um mês de forte oscilação. O mercado reagiu "
              "positivamente ao conjunto sólido de resultados do 2T26, com o papel avançando no pregão seguinte à "
              "divulgação, ainda que parte do movimento tenha sido devolvida nos dias posteriores, à medida que os "
              "investidores reavaliaram os números."),
    'VIVA3': ("Alta em agosto, acima do Ibovespa, refletindo resultados em linha com as expectativas no 2T26. "
              "Apesar do cenário competitivo mais desafiador para a Life, dos impactos da Copa do Mundo e da base "
              "de comparação elevada, a marca Vivara apresentou crescimento saudável, e a companhia espera melhora "
              "nas vendas mesmas lojas da Life no 2S26."),
    'CEAB3': ("Queda em agosto, abaixo do Ibovespa. O movimento reflete o ambiente mais negativo para o setor "
              "de vestuário após os resultados do 2T26, afetados pelo cenário macroeconômico desafiador e pelo "
              "menor fluxo nos shoppings durante a Copa do Mundo. Para o 3T, um inverno mais quente pode pressionar "
              "as vendas da coleção e exigir remarcações."),
    'TUPY3': ("Queda em agosto, abaixo do Ibovespa, em meio a resultados ainda pressionados no 2T26, com "
              "efeitos do câmbio e demanda mais fraca no mercado doméstico. Por outro lado, a melhora sequencial "
              "das margens a partir do 3T26, a redução da dívida líquida e o fortalecimento da carteira de pedidos "
              "na América do Norte reforçam a expectativa de inflexão."),
    'PRNR3': "Queda em agosto, abaixo do Ibovespa.",
    'ECOR3': "Alta em agosto, acima do Ibovespa.",
    'LWSA3': ("Alta em agosto, acima do Ibovespa e um dos destaques positivos da carteira no mês. O movimento "
              "foi impulsionado pelos resultados do 2T26, que superaram o consenso de mercado e reforçaram a "
              "percepção de melhora operacional, com o papel mantendo trajetória favorável ao longo do período."),
    'CURY3': ("Alta em agosto, acima do Ibovespa. Segue nossa principal escolha entre as construtoras de "
              "baixa renda, apoiada por execução consistente, elevada rentabilidade, forte velocidade de vendas e "
              "geração de caixa. O resultado do 2T26 reforçou essa leitura, com crescimento sólido de receita e "
              "margens preservadas em patamares elevados."),
    'DIRR3': ("Incluímos a Direcional na carteira em setembro, apoiados sobretudo pelo múltiplo atrativo e "
              "pela exposição à melhora das perspectivas do MCMV. A companhia negocia a 4,8 vezes o lucro projetado "
              "para 2027, mantendo execução sólida e elevada rentabilidade."),
    'TTEN3': ("Queda em agosto, abaixo do Ibovespa, refletindo o sentimento ainda pressionado após um "
              "trimestre de resultados abaixo do esperado. Persistem riscos de revisões negativas de lucro e ruídos "
              "na comunicação com o mercado, mas esperamos melhora estrutural ao longo do 3T e do 4T. Segue nossa "
              "principal preferência no segmento."),
    'SLCE3': ("Alta em agosto, acima do Ibovespa e o principal destaque positivo da carteira, refletindo a "
              "alta dos preços das commodities agrícolas. A alavancagem da companhia e os riscos associados ao El "
              "Niño ainda limitam uma valorização mais expressiva no curto prazo, mas o cenário mais construtivo "
              "para as commodities sustenta a trajetória."),
    'ORVR3': ORVR3,
    'ALUP11': ("Alta em agosto, acima do Ibovespa. Por ser um papel mais defensivo, com receitas previsíveis, "
               "tende a apresentar desempenho mais resiliente em períodos adversos. Os resultados do 2T26 vieram em "
               "linha com as expectativas, com geração ligeiramente acima do esperado e avanço das receitas de "
               "transmissão, refletindo a entrada em operação de novos projetos."),
    'BRBI11': ("Queda em agosto, abaixo do Ibovespa, acompanhando o cenário macroeconômico mais desafiador, "
               "que pode limitar a conversão de operações e a atividade de mercado de capitais em 2026. Seguimos "
               "construtivos com a tese, ainda que o cenário de curto prazo permaneça desafiador."),
}

# --- caixas cinzas ----------------------------------------------------------
MUDANCAS_ACOES = [
    ("Estamos adicionando AXIA3 com 7,5%.", True),
    (" A forte visibilidade da companhia entre os estrangeiros e um potencial aumento na alocação em Brasil podem "
     "dar suporte às ações. À frente, o El Niño pode elevar os riscos de ondas de calor e a volatilidade da "
     "geração eólica, sustentando os preços de energia no 4T26. ", False),
    ("Estamos adicionando CURY3 com 5%, ", True),
    ("nossa principal escolha entre as construtoras de baixa renda. A tese segue apoiada por execução "
     "consistente, elevada rentabilidade, forte geração de caixa e perfil atrativo de dividendos. ", False),
    ("Estamos aumentando nossa exposição a RDOR3 de 7,5% para 10%", True),
    (", líder de mercado com posicionamento forte em hospitais e em planos de saúde, o que sustenta ganhos "
     "contínuos de participação de mercado. ", False),
    ("Estamos reduzindo o peso de ITUB4 de 10,0% para 7,5% e ROXO34 de 7,5% para 5,0%", True),
    (", diante de um cenário macroeconômico e de crédito mais desafiador. ", False),
    ("Estamos retirando ORVR3 (de 5%).", True),
    (" Seguimos construtivos no longo prazo, mas as incertezas de curto prazo devem continuar pressionando os "
     "próximos resultados. ", False),
    ("Estamos zerando nossa posição em LREN3.", True),
    (" Entre os riscos, destacamos a revogação definitiva da taxação das encomendas internacionais de baixo "
     "valor, um cenário macroeconômico ainda desafiador e um 3T pressionado.", False),
]

MUDANCAS_DIVS = [
    ("Estamos aumentando o peso de PETR4 de 10% para 12,5%, ", True),
    ("principalmente para elevar nossa exposição a Óleo & Gás e reduzir a subalocação do portfólio frente ao "
     "Ibovespa. A Petrobras segue entre nossas principais preferências no setor e, em nossa visão, oferece uma "
     "das melhores relações entre risco e retorno no ambiente atual. ", False),
    ("Estamos aumentando ALOS3 de 7,5% para 10%, ", True),
    ("apoiados pelo múltiplo atrativo e pela maior visibilidade de retorno ao acionista. A companhia oferece "
     "retorno em dividendos de 13,1% em 2026E, o maior entre as operadoras de shopping da nossa cobertura, e "
     "negocia a 7,2x EV/EBITDA em 2027E. ", False),
    ("Estamos reduzindo o peso de ITUB4 de 15,0% para 10,0%.", True),
    (" Seguimos vendo o Itaú como um dos bancos mais bem posicionados para sustentar retornos acima dos pares e "
     "do próprio custo de capital. No entanto, diante de um cenário macroeconômico e de crédito mais desafiador, "
     "optamos por reduzir marginalmente a exposição ao papel.", False),
]

MUDANCAS_SMALL = [
    ("Estamos adicionando DIRR3 com 5% de peso, ", True),
    ("apoiados sobretudo pelo múltiplo atrativo e pela exposição à melhora das perspectivas do MCMV. A companhia "
     "negocia a 4,8x P/L em 2027E, mantendo execução sólida e elevada rentabilidade. ", False),
    ("Estamos aumentando SLCE3 de 5% para 7,5%, ", True),
    ("refletindo uma visão mais construtiva para as commodities agrícolas, especialmente algodão, soja e milho, o "
     "que resulta em um balanço de risco e retorno mais favorável. ", False),
    ("Estamos reduzindo nossa posição em CEAB3 de 10% para 5%", True),
    (", em função da decisão de diminuir a exposição ao setor de consumo em um cenário macroeconômico ainda "
     "desafiador e de um 3T de vestuário pressionado pelo inverno mais quente e pelo menor fluxo durante a Copa "
     "do Mundo. ", False),
    ("Estamos reduzindo CURY3 de 12,5% para 10%", True),
    (" apenas para abrir espaço para a Direcional e manter exposição equilibrada ao setor, sem alteração em nossa "
     "visão positiva para a Cury.", False),
]

SUB_ACOES = "Adicionando CURY3 e AXIA3; aumentando RDOR3; removendo LREN3 e ORVR3; reduzindo ITUB4 e ROXO34"
SUB_DIVS = "Aumentando PETR4 e ALOS3; reduzindo ITUB4"
SUB_SMALL = "Adicionando DIRR3; aumentando SLCE3; reduzindo CEAB3 e CURY3"
