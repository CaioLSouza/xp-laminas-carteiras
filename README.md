# Lâminas das Carteiras Recomendadas XP — Setembro 2026

Fechamento de **agosto/2026**. Quatro lâminas geradas a partir do Raio-XP do mês,
da planilha `Charts Lâmina Carteiras.xlsm` e dos comentários de desempenho dos
analistas.

| Lâmina | Nomes | Retorno em agosto | Benchmark |
|---|---|---|---|
| Top Ações XP | 14 | −0,4% | Ibovespa −0,3% |
| Top Dividendos | 11 | −0,9% | Ibovespa −0,3% |
| Top Small Caps | 14 | +1,2% | Ibovespa −0,3% / SMLL −1,3% |
| Small Caps Clear | 14 | +1,2% | espelho da Top Small Caps |

Os arquivos finais estão em [`output/`](output).

## Mudanças de composição

- **Top Ações** — entram CURY3 (5,0%) e AXIA3 (7,5%); saem LREN3 e ORVR3.
  RDOR3 7,5% → 10,0%; ITUB4 10,0% → 7,5%; ROXO34 7,5% → 5,0%.
- **Top Dividendos** — sem entradas nem saídas. PETR4 10,0% → 12,5%;
  ALOS3 7,5% → 10,0%; ITUB4 15,0% → 10,0%.
- **Top Small Caps / Clear** — entra DIRR3 (5,0%). SLCE3 5,0% → 7,5%;
  CEAB3 10,0% → 5,0%; CURY3 12,5% → 10,0%.

## Nota sobre a fonte dos dados

As abas `_last` da planilha chegaram com o fechamento de **julho**: a composição
e os pesos estavam atualizados para agosto, mas as colunas de retorno não. O
sintoma é silencioso — a soma ponderada Σ(peso × desempenho no mês) não fechava
com a série `base 100` por 5,0 p.p. (Top Ações), 3,8 p.p. (Dividendos) e
6,4 p.p. (Small Caps), e os números reproduziam célula a célula a lâmina do mês
anterior.

As abas `now_last_*` reconciliam com a série na quarta casa decimal nas três
carteiras, e é delas que estes decks foram gerados. Em 16 dos 38 nomes o sinal
do mês difere entre as duas fontes.

Pelo mesmo motivo, a direção de cada comentário ("Alta em agosto" / "Queda em
agosto") vem da planilha, não dos percentuais citados pelos analistas — que
usam janela de 30 dias corridos, enquanto a lâmina usa o mês calendário.

## Estrutura

```
build.py              monta as quatro lâminas
verify.py             19 checagens por lâmina
render.py             .pptx -> PNG/PDF para o QA visual
content.py            comentários por ticker, caixas cinzas e subtítulos do mês
fix_tabela.py         ajuste das tabelas de composição das capas de Small Caps
fix_coment.py         ajuste das tabelas de comentários de Small Caps
output/               as quatro lâminas
```

`build.py`, `verify.py` e `render.py` são cópias dos scripts da skill
`xp-laminas-carteiras`, com as abas de desempenho por ativo apontadas para
`now_last_*` neste mês.

Reprodução:

```bash
python build.py --fechamento 2026-08 --xlsx "Charts Lâmina Carteiras.xlsm" --templates <pasta-dos-templates>
python verify.py --fechamento 2026-08 --xlsx "Charts Lâmina Carteiras.xlsm"
python fix_tabela.py && python fix_coment.py
```

## Ajustes de layout

A Small Caps foi de 13 para 14 nomes e duas tabelas passaram do rodapé. O
`reposiciona()` do `build.py` dimensiona pela `shape.height`, mas quem manda no
render é a soma das alturas de linha — as duas divergem no template (2,58 vs
2,73 pol).

- Capa Small Caps XP: altura de linha 0,1612 → 0,1475 pol, tabela devolvida ao
  topo original do template.
- Capa Clear: altura de linha 0,1612 → 0,1560 pol, topo 4,249 → 4,181 pol.
- Tabelas de comentários: alturas redimensionadas pelo número de linhas de cada
  comentário; na XP a tabela subiu de 0,950 para 0,850 pol.

## Pendências

- **Priner (PRNR3)** e **Ecorodovias (ECOR3)** ficaram sem comentário dos
  analistas. Receberam apenas a direção verificável do mês, sem tese.
- **Figura 6 da Top Dividendos** (fluxo de dividendos pagos) é preenchida à mão
  e ainda termina em jun/26 — faltam jul/26 e ago/26.
- O texto fixo dos templates traz "ações com *valuation* atraentes" (Top Ações e
  Small Caps) e "*dividend yield* atrativo" (Dividendos), herdados.

---

Material do XP Research. Os decks trazem o disclaimer da XP Investimentos.
