# Conformidade com PROCONVE L8: Análise de Emissões Veiculares no Brasil

**Trabalho de Conclusão de Curso — Sistemas de Informação | UFU Monte Carmelo**  
**Autor:** Felipe Leme Dias | **Orientador:** Igor da Penha Natal

---

## Objetivo

Identificar quais modelos de veículos leves comercializados no Brasil em 2026
**violam os limites de emissão do PROCONVE Fase L8**, quantificar o excesso e
categorizar os resultados por marca, tipo de combustível e categoria de veículo.

**Pergunta central:**
> *Quais modelos violam os limites de CO e NMOG+NOx estabelecidos pelo PROCONVE L8,
> e qual é o percentual de excesso em relação ao limite vigente?*

---

## Contexto Regulatório — PROCONVE L8

O **PROCONVE** (Programa de Controle da Poluição do Ar por Veículos Automotores)
é a principal norma brasileira de controle de emissões veiculares. A **Fase L8**
entrou em vigor em 1º de janeiro de 2025 e é a versão mais restritiva já adotada
no país.

### Limites por tipo de combustível

| Poluente       | L8 Gasolina / Flex | L8 Diesel    | Fase L7 (Gasolina) | Redução |
|----------------|--------------------|--------------|--------------------|---------|
| CO             | 700 mg/km          | 500 mg/km    | 1 300 mg/km        | −46 %   |
| NMOG + NOx     | 60 mg/km           | 460 mg/km    | 80 mg/km           | −25 %   |
| NOx            | 0,06 g/km          | 0,08 g/km    | 0,08 g/km          | −25 %   |
| THC            | 60 mg/km           | —            | 80 mg/km           | −25 %   |
| Material Part. | 4,5 mg/km          | 4,5 mg/km    | 6 mg/km            | −25 %   |

> **Atenção:** os limites de NMOG+NOx para diesel (460 mg/km) são
> completamente distintos dos limites para gasolina/flex (60 mg/km). Toda a
> análise de conformidade é segmentada por combustível.

---

## Fonte de Dados

| Fonte | Descrição | Acesso |
|-------|-----------|--------|
| **PBEV / INMETRO** | Programa Brasileiro de Etiquetagem Veicular — tabela anual de emissões de automóveis e comerciais leves | [inmetro.gov.br](https://www.inmetro.gov.br) |

**Arquivo utilizado:** `Tabela PBEV 2026_20_JAN-REV04.pdf`  
**Localização:** `data/pbve/`  
**Cobertura:** Janeiro/2026, Revisão 04  
**Conteúdo:** 769 modelos/versões certificados, 39 marcas, 16 categorias

> Dados de vendas (FENABRAVE / ANFAVEA) foram avaliados mas descartados:
> FENABRAVE não disponibiliza rankings por modelo em formato aberto;
> ANFAVEA fornece dados apenas por marca, sem granularidade por modelo.

---

## Estrutura do Projeto

```
.
├── data/
│   ├── pbve/                       # PDFs originais PBEV
│   │   └── Tabela PBEV 2026_20_JAN-REV04.pdf
│   └── pbev_2026.csv               # Dataset extraído e processado
├── extract_pbev.py                 # Pipeline de extração (pdfplumber)
├── analyze.ipynb                   # Notebook de análise exploratória (14 seções)
├── dicionario_pbev_2026.xlsx       # Dicionário de variáveis (3 abas)
├── make_dict.py                    # Script que gera o dicionário Excel
├── fig_01_composicao.png           # Composição da frota
├── fig_02_dispersao_l8.png         # Dispersão CO × NMOG+NOx
├── fig_03_co_distribuicao.png      # Distribuição de CO
├── fig_04_nmog_distribuicao.png    # Distribuição de NMOG+NOx
├── fig_05_top20_co.png             # Top 20 maiores emissores de CO
├── fig_06_marca_conformidade.png   # Conformidade por marca
├── fig_07_categoria.png            # Conformidade por categoria
├── fig_08_co2.png                  # CO₂ e consumo por tecnologia
├── fig_09_nota_verde.png           # Nota Verde INMETRO vs análise própria
├── fig_10_pbe.png                  # Classificação PBE vs emissões
├── fig_11_correlacao.png           # Correlações entre variáveis
├── fig_12_sintese.png              # Card de síntese dos achados
├── .env                            # Chave da API (não versionado)
├── .gitignore
└── README.md
```

---

## Pipeline de Extração

### Por que pdfplumber e não Claude API?

A opção inicial era usar `claude-opus-4-7` via Anthropic API para extrair os
dados do PDF. Dois obstáculos inviabilizaram essa abordagem no tier gratuito:

1. **Rate limit de input:** 10 000 tokens/minuto — o PDF de 9 páginas excede
   esse limite em um único request.
2. **Rate limit de output:** 4 000 tokens/minuto — o JSON com 769 veículos
   (~60 000 tokens) levaria ~15 minutos para ser gerado, causando timeout
   inevitável (SDK padrão: 10 min). O custo seria cobrado mesmo sem retorno.

A solução foi usar **`pdfplumber`**, biblioteca Python que extrai tabelas de
PDFs gerados digitalmente (não-escaneados) **sem custo de API**.

### Como o mapeamento de colunas foi feito

O PBEV 2026 tem um cabeçalho multi-linha com células mescladas que confunde
extratores automáticos. A abordagem foi:

1. Inspecionar a linha do **CHEVROLET SPIN 1.8L** (cujo CO=800 mg/km era
   conhecido de antemão) para confirmar cada índice de coluna.
2. Hard-codar os índices verificados empiricamente:

| Índice | Coluna CSV          | Exemplo (SPIN 1.8L AT) |
|--------|---------------------|------------------------|
| 0      | `categoria`         | Médio                  |
| 1      | `marca`             | CHEVROLET              |
| 2      | `modelo`            | SPIN                   |
| 3      | `versao`            | AT                     |
| 4      | `motor`             | 1.8L - 8V              |
| 5      | `tipo_propulsao`    | Combustão              |
| 6      | `transmissao`       | A-6                    |
| 9      | `combustivel`       | F                      |
| 10     | `nmog_nox_mg_km`    | 46                     |
| **11** | **`co_mg_km`**      | **800** ← confirmado   |
| 13     | `nota_verde_l8`     | C                      |
| 15     | `co2_fossil_g_km`   | 119                    |
| 23     | `consumo_mj_km`     | 1.92                   |
| 25     | `classificacao_pbe` | E                      |

### Filtro de linhas válidas

Em vez de uma lista de marcas hardcoded (quebraria com marcas novas como
ZEEKR, LEAPMOTOR, OMODA), o script filtra por:

- Linha com pelo menos 12 colunas
- Posição `[1]` (marca) **não vazia** e diferente de palavras-chave de cabeçalho
- Posição `[2]` (modelo) **não vazia**

### Colunas calculadas no `extract_pbev.py`

Após a extração, o script calcula automaticamente:

| Coluna              | Fórmula                                        |
|---------------------|------------------------------------------------|
| `co_excesso_pct`    | `max(0, (co_mg_km − 700) / 700 × 100)`        |
| `nmog_excesso_pct`  | `max(0, (nmog_nox_mg_km − 60) / 60 × 100)`   |
| `viola_l8`          | `co_mg_km > 700 OR nmog_nox_mg_km > 60`       |

> Esses limites são para gasolina/flex. O `analyze.ipynb` recalcula `viola_l8`
> corretamente separando diesel (CO≤500, NMOG+NOx≤460) de gasolina/flex.

---

## Schema do Dataset (`data/pbev_2026.csv`)

| Coluna              | Tipo    | Unidade  | Descrição                                           |
|---------------------|---------|----------|-----------------------------------------------------|
| `categoria`         | texto   | —        | Segmento de mercado (Sub Compacto, Médio, Picape…)  |
| `marca`             | texto   | —        | Fabricante (normalizado para maiúsculas)             |
| `modelo`            | texto   | —        | Nome comercial                                      |
| `versao`            | texto   | —        | Trim level (LX, EXL, PREMIER…) — 19 nulos          |
| `motor`             | texto   | —        | Ex.: 1.0-6V, 1.8L-8V, Elétrico                    |
| `tipo_propulsao`    | texto   | —        | Combustão · Elétrico · Híbrido · Plug-in            |
| `transmissao`       | texto   | —        | M-5, A-6, CVT, DCT-7, N.A.…                        |
| `combustivel`       | texto   | —        | F=Flex · G=Gasolina · D=Diesel · E=Elétrico         |
| `nmog_nox_mg_km`    | real    | mg/km    | NMOG+NOx — 71 nulos (EVs)                          |
| `co_mg_km`          | real    | mg/km    | Monóxido de carbono — 45 nulos (EVs)               |
| `nota_verde_l8`     | texto   | —        | Avaliação INMETRO de conformidade L8: A→C           |
| `co2_fossil_g_km`   | real    | g/km     | CO₂ fóssil — 158 nulos (EVs e PHEVs)              |
| `consumo_mj_km`     | real    | MJ/km    | Consumo energético total — sem nulos                |
| `classificacao_pbe` | texto   | —        | Eficiência energética PBE: A (melhor) → E (pior)   |
| `co_excesso_pct`    | real    | %        | Excesso de CO sobre limite L8 (calculado)           |
| `nmog_excesso_pct`  | real    | %        | Excesso de NMOG+NOx sobre limite L8 (calculado)    |
| `viola_l8`          | bool    | —        | True se qualquer limite L8 foi violado (calculado)  |

**Dicionário completo:** `dicionario_pbev_2026.xlsx` (3 abas: Variáveis, Legendas do PDF, Limites L8)

---

## Como Reproduzir

### 1. Dependências

```bash
pip install pdfplumber pandas matplotlib seaborn openpyxl jupyter
```

### 2. Extração dos dados

```bash
python extract_pbev.py
# Saída: data/pbev_2026.csv
# Log: nº de veículos por página + total extraído + contagem de violadores
```

### 3. Dicionário de variáveis

```bash
python make_dict.py
# Saída: dicionario_pbev_2026.xlsx
```

### 4. Análise exploratória

```bash
jupyter notebook analyze.ipynb
# Executa as 14 seções e salva fig_01 a fig_12
```

---

## Análise Exploratória — Resultados

### Estatísticas do dataset

| Métrica                  | Valor          |
|--------------------------|----------------|
| Total de registros       | 751            |
| Marcas                   | 40             |
| Categorias               | 16             |
| Colunas                  | 17             |
| Período de referência    | Janeiro/2026   |

### Composição da frota (Seção 2)

| Tipo de Propulsão | Modelos | %     |
|-------------------|---------|-------|
| Combustão interna | 429     | 57,1% |
| Elétrico puro     | 142     | 18,9% |
| Híbrido (HEV)     | 89      | 11,9% |
| Plug-in (PHEV)    | 89      | 11,9% |

| Combustível | Modelos | %     |
|-------------|---------|-------|
| Gasolina    | 261     | 34,8% |
| Flex        | 230     | 30,6% |
| Diesel      | 118     | 15,7% |
| Elétrico    | 142     | 18,9% |

### Conformidade PROCONVE L8 (Seção 3)

#### Gasolina + Flex (limites: CO ≤ 700 mg/km · NMOG+NOx ≤ 60 mg/km)

| Resultado      | Modelos | % do grupo |
|----------------|---------|------------|
| Conformes      | ~479    | ~96%       |
| Violam L8      | ~11     | ~4%        |

#### Diesel (limites: CO ≤ 500 mg/km · NMOG+NOx ≤ 460 mg/km)

A análise inicial marcou 63 veículos diesel como violadores por aplicar
equivocadamente os limites de gasolina (60 mg/km para NMOG+NOx). Após a
correção com os limites diesel corretos, esses veículos são conformes. A
separação por combustível é etapa obrigatória nesta análise.

### Distribuição de CO — Gasolina e Flex (Seção 4)

| Estatística | Flex    | Gasolina |
|-------------|---------|----------|
| Mínimo      | 0 mg/km | 0 mg/km  |
| Mediana     | ~200    | ~105     |
| Média       | ~230    | ~120     |
| Máximo      | 800     | ~265     |

- A mediana de CO do Flex é significativamente maior que a da Gasolina, pois
  veículos Flex mais antigos (ex.: SPIN 1.8L) ainda utilizam plataformas de
  motor com tecnologia anterior ao L8.
- A distribuição é assimétrica à direita: a grande maioria dos modelos
  modernos está abaixo de 400 mg/km.

### Distribuição de NMOG+NOx (Seção 5)

- Média: 30 mg/km (bem abaixo do limite de 60 mg/km)
- Modelos que superam 60 mg/km são principalmente motorizações antigas Flex
  (ex.: SPIN 1.8L 6V manual: NMOG+NOx = 62 mg/km)

### Top 20 maiores emissores de CO (Seção 6)

| Pos. | Modelo                        | Comb. | CO (mg/km) | Excesso L8 |
|------|-------------------------------|-------|------------|------------|
| 1    | CHEVROLET SPIN 1.8L (4 versões) | Flex  | 800        | +14,3%     |
| 5    | RENAULT DUSTER 1.6L (3 versões) | Flex  | 654        | −6,6% (ok) |
| 8    | CHEVROLET SPIN LT Manual       | Flex  | 611        | −12,7% (ok)|
| 9    | RENAULT KANGOO (2 versões)     | Flex  | 595        | −15% (ok)  |

> Apenas o **CHEVROLET SPIN 1.8L** (versões AT e LT automático) supera o
> limite de 700 mg/km de CO para Flex/Gasolina.

### Análise por Marca (Seção 7)

Marcas com maior **taxa percentual** de modelos violadores (todos combustíveis
combinados, antes da segmentação diesel):

| Marca      | Total | Violam | %     | Observação principal          |
|------------|-------|--------|-------|-------------------------------|
| RAM        | 5     | 5      | 100%  | Picapes diesel                |
| CITROEN    | ~10   | 8      | ~80%  | Comerciais leves diesel       |
| FORD       | ~12   | 8      | ~67%  | Nova Ranger diesel            |
| CHEVROLET  | ~30   | 10     | ~33%  | Inclui SPIN + S10 diesel      |
| FIAT       | ~40   | 14     | ~35%  | Ducato/Doblò comercial diesel |
| TOYOTA     | ~35   | 14     | ~40%  | Hilux e SW4 diesel            |

Após segmentação correta (limites diesel vs gasolina/flex), o único modelo
**Flex/Gasolina com CO acima do limite** pertence à **CHEVROLET** (SPIN).

### Análise por Categoria (Seção 8)

| Categoria           | Total | Violam (bruto) | Nota                             |
|---------------------|-------|----------------|----------------------------------|
| Picape              | 76    | 38             | Quase todos diesel (limites distintos) |
| Comercial           | 78    | 25             | Quase todos diesel               |
| Fora de Estrada Grande | 38 | 6              | Diesel                           |
| Médio               | 92    | 3              | SPIN 1.8L (Flex)                |
| Minivan             | 5     | 2              | SPIN (Flex)                     |

### CO₂ Fóssil e Consumo Energético (Seção 9)

| Tecnologia | CO₂ Fóssil (mediana) | Consumo (mediana) |
|------------|----------------------|-------------------|
| Combustão  | 111 g/km             | 1,95 MJ/km        |
| Híbrido    | 103 g/km             | 1,84 MJ/km        |
| Plug-in    | 34 g/km              | 0,88 MJ/km        |
| Elétrico   | 0 g/km               | 0,59 MJ/km        |

- Plug-ins reduzem o CO₂ fóssil em ~**70%** em relação à combustão convencional
- Elétricos puros eliminam totalmente o CO₂ fóssil direto

### Nota Verde INMETRO × Nossa Análise (Seção 10)

A **Nota Verde** do INMETRO é atribuída globalmente pelo fabricante durante a
certificação e usa critérios que incluem todos os poluentes (CO, NMOG+NOx, NOx)
com ponderações internas não públicas. Nossa análise avalia apenas CO e
NMOG+NOx com os limites tabelados.

| Nota Verde | Conformes (nossa análise) | Violam (nossa análise) |
|------------|--------------------------|------------------------|
| A          | 515                      | 63 (todos diesel*)     |
| B          | 143                      | 6                      |
| C          | 19                       | 5                      |

*Os 63 veículos com Nota Verde A marcados como violadores na análise bruta
são todos diesel — o INMETRO os classifica corretamente com A porque
aplicou os limites diesel corretos.

**Achado:** nenhum veículo com Nota Verde A ou B viola os limites de
Gasolina/Flex. A classificação INMETRO é consistente após a segmentação.

### Classificação PBE × Conformidade L8 (Seção 11)

A **Classificação PBE** (eficiência energética) e a **conformidade L8**
(emissões de poluentes) são critérios **completamente independentes**:

- Um veículo pode ter PBE = E (menos eficiente energeticamente) e ainda
  assim ser 100% conforme com os limites de poluentes do L8.
- A SPIN 1.8L tem classificação PBE = E e viola o L8 — ambos os critérios
  apontam para o mesmo veículo, mas por razões diferentes.
- Modelos elétricos têm PBE = A e Nota Verde = A, sendo os únicos que
  maximizam ambas as métricas simultaneamente.

### Correlações entre Variáveis (Seção 12)

| Par de variáveis             | Correlação (r) | Interpretação              |
|------------------------------|---------------|---------------------------|
| CO₂ fóssil × Consumo MJ/km  | ~0,90         | Forte positiva — esperada |
| CO × NMOG+NOx                | ~0,55         | Moderada positiva         |
| CO × CO₂ fóssil              | ~0,20         | Fraca — poluentes ≠ GEE   |

A correlação fraca entre CO (poluente) e CO₂ (gás de efeito estufa) é um
resultado relevante: reduzir CO₂ não implica necessariamente reduzir CO, e
vice-versa. São problemas distintos que exigem soluções distintas.

### Ranking final — Violadores Flex/Gasolina (Seção 13)

| # | Marca     | Modelo  | Versão    | CO (mg/km) | Excesso CO | NMOG+NOx | Nota Verde |
|---|-----------|---------|-----------|------------|------------|----------|------------|
| 1 | CHEVROLET | SPIN    | AT        | 800        | +14,3%     | 46       | C          |
| 2 | CHEVROLET | SPIN    | LT (aut.) | 800        | +14,3%     | 46       | C          |
| 3 | CHEVROLET | SPIN    | LTZ       | 800        | +14,3%     | 46       | C          |
| 4 | CHEVROLET | SPIN    | PREMIER   | 800        | +14,3%     | 46       | C          |
| 5 | CHEVROLET | SPIN    | LT (man.) | 611        | — (ok)     | 62       | C          |

> O SPIN LT manual não viola o CO (611 < 700) mas viola o NMOG+NOx (62 > 60,
> excesso de +3,3%).

---

## Principais Achados

1. **751 modelos/versões** de **40 marcas** foram certificados no PBEV 2026,
   representando toda a frota de veículos leves novos comercializados no Brasil.

2. A frota é majoritariamente composta por **combustão interna (57%)**, mas
   elétricos e plug-ins já somam **31%** dos modelos certificados.

3. **Limites L8 são combustível-específicos.** Analisar diesel com limites de
   gasolina gera falsos positivos. A segmentação é obrigatória.

4. Entre veículos **Flex e Gasolina**, apenas os modelos com motorização
   **CHEVROLET SPIN 1.8L** violam o limite de CO do L8 (+14,3% acima de 700 mg/km).

5. **NMOG+NOx:** o SPIN LT manual (motorização 6V) também viola levemente o
   limite de 60 mg/km com 62 mg/km (+3,3%).

6. **Eficiência energética (PBE) e conformidade de poluentes (L8) são métricas
   independentes.** Um veículo eficiente não é necessariamente "limpo" em
   termos de poluentes locais, e vice-versa.

7. **Plug-ins** reduzem o CO₂ fóssil em ~70% em relação à combustão
   convencional. **Elétricos puros** eliminam totalmente o CO₂ fóssil direto.

8. CO₂ e CO são fracamente correlacionados (r ≈ 0,20): são problemas
   distintos que exigem políticas distintas.

---

## Limitações

| Limitação | Impacto |
|-----------|---------|
| Dados de vendas por modelo indisponíveis em formato aberto | Impossível ranquear violadores por popularidade |
| Apenas ano 2026 analisado | Sem análise de tendência temporal |
| Valores de NOx não extraídos (coluna ambígua no PDF) | Análise restrita a CO e NMOG+NOx |
| pdfplumber perde ~18 linhas em quebras de página | ~2% de subcobertura no dataset |
| Teste realizado no ciclo brasileiro (não WLTP) | Valores podem diferir de normas internacionais |

---

## Arquivos Gerados

| Arquivo | Descrição |
|---------|-----------|
| `data/pbev_2026.csv` | Dataset principal (751 linhas × 17 colunas) |
| `dicionario_pbev_2026.xlsx` | Dicionário de variáveis com 3 abas |
| `fig_01_composicao.png` | Composição da frota por propulsão, combustível e categoria |
| `fig_02_dispersao_l8.png` | Dispersão CO × NMOG+NOx com limites L8 por combustível |
| `fig_03_co_distribuicao.png` | Histograma + boxplot de CO (Gasolina/Flex) |
| `fig_04_nmog_distribuicao.png` | Histograma de NMOG+NOx com zona de violação |
| `fig_05_top20_co.png` | Top 20 maiores emissores de CO |
| `fig_06_marca_conformidade.png` | Taxa de violação e CO médio por marca |
| `fig_07_categoria.png` | Conformidade e CO mediano por categoria |
| `fig_08_co2.png` | CO₂ fóssil e consumo energético por tecnologia |
| `fig_09_nota_verde.png` | Nota Verde INMETRO × violação calculada |
| `fig_10_pbe.png` | Classificação PBE versus emissões e consumo |
| `fig_11_correlacao.png` | Mapa de correlação entre variáveis numéricas |
| `fig_12_sintese.png` | Card visual com os 9 principais achados |
