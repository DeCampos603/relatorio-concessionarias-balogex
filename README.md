# Relatório Consolidado de Concessionárias — BA Ap Log Ex e OMDS

**No ar:** https://decampos603.github.io/relatorio-concessionarias-balogex/


Site estático (HTML autocontido) que consolida os relatórios do **SAG — Sistema de Acompanhamento
da Gestão** de **água e esgoto** e **energia elétrica** das seis UG apoiadas pela BA Ap Log Ex na
1ª RM, de 2023 a 2026, e responde à pergunta: **por que a despesa está mais cara em 2026?**

- Entrada: `G:\Meu Drive\REPO\RELATORIO CONCESSIONARIAS`
  - 56 arquivos `.xls` do **SAG** (consumo e faturamento, 2023–2026)
  - 3 crosstabs `.xlsx` do **Tesouro Gerencial** (execução orçamentária, 2022–2025 encerrados
    e 2026 até agosto)
- Saída: `site/index.html` (autocontido) + `site/dados.json`
- Publicação: **GitHub Pages**
- Dependência: `openpyxl` (só para os crosstabs; sem ela o bloco orçamentário é omitido)

---

## Unidades Gestoras cobertas

| UG | Sigla | Denominação |
|---|---|---|
| 160238 | Ba Ap Log Ex | Base de Apoio Logístico |
| 160246 | DC Mun | Depósito Central de Munição |
| 160304 | BMSA | Batalhão de Manutenção e Suprimento de Armamento |
| 160307 | 1º D Sup | 1º Depósito de Suprimentos |
| 160321 | ECT | Estabelecimento Central de Transportes |
| 160329 | BCMS | Batalhão Central de Manutenção e Suprimento |

Todas em **1ª RM / CML / 1º CGCFEx / RJ**.

> **Procedência dos identificadores.** UG, sigla, RM, comando, CGCFEx e UF vêm dos próprios
> arquivos do SAG. A **denominação por extenso** foi informada pela BA Ap Log Ex — o sistema
> não a exporta. O **nome da concessionária** continua sem fonte em todo o acervo, e por isso a
> organização "por concessionária" é feita por **serviço**.
>
> Para alterar as denominações, edite a lista `UGS` no topo de `gerar_site.py`.

---

## O que o relatório conclui

1. **Água — a conta subiu porque o preço subiu.** A tarifa média efetiva do conjunto passou de
   R$ 27,20/m³ (Jan–Jun/2025) para R$ 28,21/m³ (Jan–Jun/2026), +3,7%. Desde 2023 a alta acumulada
   é de 28,8%. O conjunto consumiu 8,6% menos água e ainda assim só reduziu 5,2% da despesa:
   o **efeito preço de R$ 81,6 mil** no semestre anulou 43% da economia de consumo.
   Três UG gastaram **mais reais consumindo menos água**.
2. **Energia — o consumo subiu em 4 das 6 UG**, mas o valor lançado no SAG **não fecha**.
   São quatro defeitos de lançamento distintos, documentados como achados A1 a A4.
3. **Orçamento — 2026 herdou R$ 5,64 milhões de restos a pagar de 2025.** O exercício de 2025
   empenhou R$ 12,15 mi em concessionárias e liquidou apenas 53,6%; a diferença virou resto a
   pagar — **2,6× o que 2024 havia deixado**. O crosstab de 2026 confirma o valor **ao centavo**.
   Até agosto de 2026, **88,1%** do desembolso com concessionárias foi para quitar exercícios
   anteriores, e a **água roda inteiramente sobre o resto a pagar de 2025** (zero empenhado, zero
   liquidado em 2026). O custo próprio de 2026 ainda não apareceu (achado A8).
4. **O caixa valida a água e desmente a energia.** A razão entre o pago no ano e o faturado no
   SAG é **1,01–1,05 na água** (confere) e **1,19 na energia** nos três exercícios — ou seja,
   o SAG registra só ~84% do que a energia custa. Faltam demanda contratada, energia
   reativa, bandeiras e iluminação pública (achado A7).
5. **Acervo — o SAG não identifica a concessionária.** A organização "por concessionária"
   é por **serviço** (energia elétrica / água e esgoto). Não há faturas originais no acervo.

### Planos internos das concessionárias

| PI | Serviço | Pago em 2025 |
|---|---|---|
| `I3DACSPAGES` | Água e esgoto | R$ 4.390.789 |
| `I3DACSPENEL` | Energia elétrica e iluminação pública | R$ 4.209.322 |
| `I3DACSPTELM` | Telefonia móvel | R$ 25.175 |
| `I3DACSPTELF` | Telefonia fixa | R$ 0 (plano extinto após 2023) |
| `I3DACSPCORR` | Serviços postais | R$ 12.009 |

Telefonia e correios somados são **0,43%** do desembolso com concessionárias.

Os achados A1 a A4 são inferidos da consistência interna dos dados. São fortes o bastante para
suspender o uso dos números, mas **só a fatura original confirma** o que aconteceu.

### Regra da meta (apurada nos arquivos, não presumida)

A meta lançada no SAG é uma fração fixa da média dos três anos anteriores, mês a mês.
Conferido em todos os registros do acervo:

| Série | Meta |
|---|---|
| Água (m³) | **85%** da média trienal |
| Energia fora ponta (kWh) | **85%** da média trienal |
| Energia ponta (kWh) | **90%** da média trienal |

Somar as duas metas de energia e comparar com o consumo total mistura dois critérios — por isso
o relatório apresenta ponta e fora ponta separadamente.

---

## Como funciona

```
RELATORIO CONCESSIONARIAS/
  00-CONSOLIDADO-1a-RM/             (2 relatórios de meta)
  01-ENERGIA-ELETRICA/<UG>-<SIGLA>/ (6 séries por UG)
  02-AGUA-E-ESGOTO/<UG>-<SIGLA>/    (3 séries por UG)
  CRÉDITO DISP *.xlsx               (crosstabs do Tesouro Gerencial)
                    │
        gerar_site.py  ─►  lê o HTML dentro dos .xls e os crosstabs do TG,
                    │      calcula indicadores, decompõe preço × volume,
                    │      cruza faturamento × caixa, monta os achados
                    ▼
        site/index.html + site/dados.json  ─►  GitHub Pages
```

**Duas pegadinhas de leitura**, ambas tratadas no script:

- Os arquivos do SAG têm extensão `.xls` mas o conteúdo é **HTML de tabela do Highcharts** —
  são lidos direto do HTML, sem Excel.
- Os crosstabs do Tesouro Gerencial trazem **linhas de subtotal por Plano Interno**, que são
  descartadas na leitura (sem isso o valor dobraria), e a natureza `339000` carrega apenas
  provisão e crédito disponível, nunca despesa. O **layout do cabeçalho muda entre extrações**
  (métrica antes ou depois do período), então as duas linhas são localizadas pelo conteúdo, não
  pela posição. O arquivo de 2026 usa `Mês Lançamento` em vez de `NC — Dia Emissão`, e por isso
  a posição parcial do exercício é apresentada em bloco separado dos anos fechados.

---

## Rodar localmente

```bash
py -3 -m pip install -r requirements.txt
```

```bash
py -3 gerar_site.py
```

Abra `site\index.html` no navegador, ou sirva a pasta:

```bash
py -3 -m http.server 8142 --directory site
```

Para apontar para outra pasta de origem:

```bash
py -3 gerar_site.py --fonte "D:\outro\caminho"
```

---

## Publicação

O repositório é **público** e o site está no ar em
https://decampos603.github.io/relatorio-concessionarias-balogex/

> ⚠️ **O GitHub Pages é público.** O consumo e a despesa das seis UG estão visíveis para qualquer
> pessoa com o link e podem ser indexados por buscadores. Para restringir, torne o repositório
> **Private** — nesse caso o Pages só funciona em planos pagos, e a alternativa é distribuir o
> `site/index.html` como arquivo.

O workflow `.github/workflows/pages.yml` republica a pasta `site/` a cada push na `main`, e usa
`enablement: true` no `configure-pages` — ou seja, ativa o Pages sozinho num clone limpo, sem
passo manual em Settings.

### Republicar depois de atualizar os dados

```bash
py -3 gerar_site.py && git add -A && git commit -m "Atualiza dados" && git push
```

---

## Atualizar com dados novos

1. Exporte os novos gráficos do SAG.
2. Coloque cada arquivo na pasta da sua UG e serviço, seguindo o padrão de nome:
   `SERVICO_UG_SIGLA_METRICA_PERIODO.xls`
   (ex.: `ENERGIA_160238_BA-AP-LOG-EX_CONSUMO-KWH-PONTA_2023-2026.xls`).
   O script lê pelo `<caption>` de dentro do arquivo, então o nome serve para você, não para ele.
3. Rode `py -3 gerar_site.py` e faça o push.

O arquivo `_INVENTARIO.csv` na raiz da pasta de origem mantém o de/para entre os nomes originais
do SAG e os nomes atuais.

---

## Estrutura

```
Relatorio-Concessionarias-BALogEx/
├── gerar_site.py                    # gerador (parser + indicadores + HTML)
├── README.md
├── .gitignore
├── .github/workflows/pages.yml      # publicação no GitHub Pages
└── site/
    ├── index.html                   # relatório (autocontido, tema claro/escuro, imprimível)
    ├── dados.json                   # dados calculados, para reuso
    ├── .nojekyll
    └── assets/logos/                # brasões das seis UG
```

---

*Fonte: SAG — relatórios de consumo e meta de água e energia, UG 160238, 160246, 160304,
160307, 160321 e 160329. Séries de Jan/2023 a Jul/2026. Documento de trabalho: os achados A1 a A4
devem ser conferidos contra as faturas originais antes de qualquer providência externa.*
