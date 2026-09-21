# 🎨 Especificação Visual - Dashboard Power BI: Álbum da Copa Definitivo 2010-2026

**Versão:** 1.0 | **Data:** 2026 | **Dataset:** 4.048 jogadores | 344 partidas | 4.846 eventos | 5 Copas

---

## 📋 Visão Geral

Dashboard profissional de **8 páginas** com navegação tipo álbum de figurinhas, drillthrough contextual, bookmarks para storytelling, what-if parameters para projeções, decomposition tree para análise de gols, key influencers para descobrir drivers de performance, e RLS (Row Level Security) por confederação.

**Público-alvo:** Recrutadores freelance (Upwork/Fiverr), gestores de futebol, analistas de performance, fãs de estatísticas.

---

## 🗂️ Modelo de Dados (Star Schema)

```
dim_world_cups          dim_countries          dim_teams
─────────────────       ─────────────────       ─────────────────
Year (PK)               CountryCode(PK)        TeamCode,Year(PK)
Host                    CountryName            CountryCode(FK)
Winner                  Confederation          WorldCupYear
RunnerUp                                       
Third                   
Fourth                  
Teams                   
Matches                 
Goals                   
────────┬────────       ────────┬────────       ────────┬────────
         │                     │                     │
         ▼                     ▼                     ▼
┌─────────────────────────────────────────────────────────────────┐
│                        fact_matches                              │
│─────────────────────────────────────────────────────────────────│
│ MatchID (PK)  │ WorldCupYear(FK) │ Stage     │ Date             │
│ Stadium       │ HomeTeam(FK)     │ AwayTeam  │ HomeGoals        │
│ AwayGoals     │ Winner           │ Attendance│ xG_Home          │
│ xG_Away       │                                                              │
└─────────────────────────────────────────────────────────────────┘
         │
         ▼
┌─────────────────────────────────────────────────────────────────┐
│                        fact_events                               │
│─────────────────────────────────────────────────────────────────│
│ EventID (PK)  │ MatchID (FK)     │ PlayerID (FK) │ TeamCode     │
│ EventType     │ Minute           │ Detail        │ xG           │
└─────────────────────────────────────────────────────────────────┘
         │
         ▼
┌─────────────────┐
│  dim_players    │
│─────────────────│
│ PlayerID (PK)   │
│ Name            │
│ CountryCode(FK) │
│ WorldCupYear(FK)│
│ Position        │
│ Age             │
│ Height/Weight   │
│ PreferredFoot   │
│ MarketValue     │
│ OverallRating   │
│ Matches/Min/G/A │
│ xG/Shots/SoT    │
│ PassAcc/Tkl/Int │
│ Clr/Drib/Fouls  │
│ YC/RC           │
│ CleanSheets/Saves│
└─────────────────┘
```

### Relacionamentos Principais (1:N)

| Tabela Origem | Coluna | → | Tabela Destino | Coluna |
|---------------|--------|---|----------------|--------|
| dim_world_cups | Year | → | fact_matches | WorldCupYear |
| dim_world_cups | Year | → | dim_players | WorldCupYear |
| dim_countries | CountryCode | → | dim_teams | CountryCode |
| dim_teams | TeamCode, WorldCupYear | → | fact_matches | HomeTeam, WorldCupYear |
| dim_teams | TeamCode, WorldCupYear | → | fact_matches | AwayTeam, WorldCupYear |
| dim_players | PlayerID | → | fact_events | PlayerID |
| fact_matches | MatchID | → | fact_events | MatchID |

---

## 📐 Medidas DAX Principais (Copie e Cole no Power BI)

### KPIs Básicos
```dax
-- Total de Gols (baseado em xG dos eventos Goal)
Total Goals = 
SUMX(
    FILTER(fact_events, fact_events[EventType] = "Goal"),
    fact_events[xG]
)

-- Partidas Distintas
Matches Played = DISTINCTCOUNT(fact_matches[MatchID])

-- Média Gols/Partida
Avg Goals Per Match = DIVIDE([Total Goals], [Matches Played])

-- Artilheiro da Copa Selecionada
Top Scorer = 
MAXX(
    TOPN(1,
        SUMMARIZE(
            dim_players,
            dim_players[Name],
            dim_players[CountryCode],
            "TotalGoals", CALCULATE(
                SUMX(FILTER(fact_events, fact_events[EventType]="Goal"), fact_events[xG]),
                ALLEXCEPT(dim_players, dim_players[PlayerID])
            )
        ),
        [TotalGoals], DESC
    ),
    dim_players[Name] & " (" & dim_players[CountryCode] & ") - " & FORMAT([TotalGoals], "0.0")
)

-- Melhor Goleiro (Mais Defesas)
Best GK = 
MAXX(
    TOPN(1,
        FILTER(
            SUMMARIZE(
                dim_players,
                dim_players[Name],
                dim_players[CountryCode],
                "Saves", CALCULATE(
                    SUM(fact_events[Saves]),
                    fact_events[EventType]="Save",
                    ALLEXCEPT(dim_players, dim_players[PlayerID])
                )
            ),
            dim_players[Position] = "GK"
        ),
        [Saves], DESC
    ),
    dim_players[Name] & " (" & dim_players[CountryCode] & ") - " & FORMAT([Saves], "0") & " defesas"
)

-- Valor de Mercado Total do Elenco
Total Market Value = SUMX(dim_players, dim_players[MarketValue])

-- Rating Médio Ponderado
Avg Rating = AVERAGE(dim_players[OverallRating])

-- Idade Média do Elenco
Avg Squad Age = AVERAGE(dim_players[Age])

-- % Jogadores Sub-23
Pct U23 = 
DIVIDE(
    COUNTROWS(FILTER(dim_players, dim_players[Age] < 23)),
    COUNTROWS(dim_players)
)

-- Gols por 90 Minutos
Goals Per 90 = 
DIVIDE(
    SUMX(FILTER(fact_events, fact_events[EventType]="Goal"), fact_events[xG]) * 90,
    SUM(dim_players[Minutes])
)

-- xG por 90 Minutos
xG Per 90 = 
DIVIDE(
    SUM(dim_players[xG]) * 90,
    SUM(dim_players[Minutes])
)

-- Clean Sheets (Jogos sem sofrer gols)
Clean Sheets = 
CALCULATE(
    DISTINCTCOUNT(fact_matches[MatchID]),
    OR(fact_matches[HomeGoals] = 0, fact_matches[AwayGoals] = 0)
)

-- Precisão de Passes Média
Pass Accuracy = AVERAGE(dim_players[PassAccuracy])

-- Cartões por Partida
Cards Per Match = 
DIVIDE(
    SUM(dim_players[YellowCards]) + SUM(dim_players[RedCards]) * 3,
    [Matches Played]
)

-- Evolução Valor de Mercado (YoY 4 anos)
Market Value YoY = 
VAR CurrYear = SELECTEDVALUE(dim_world_cups[Year])
VAR PrevYear = CurrYear - 4
RETURN
    DIVIDE(
        CALCULATE([Total Market Value], dim_world_cups[Year] = CurrYear),
        CALCULATE([Total Market Value], dim_world_cups[Year] = PrevYear)
    ) - 1

-- Impacto no Torneio (Performance Média)
Tournament Impact = 
DIVIDE(
    SUMX(dim_players, dim_players[Goals] + dim_players[Assists] + dim_players[xG]),
    COUNTROWS(dim_players)
)

-- Idade do Pico de Performance
Peak Age = 
CALCULATE(
    AVERAGE(dim_players[Age]),
    TOPN(10, dim_players, dim_players[OverallRating], DESC)
)

-- Vantagem de Casa
Home Win % = 
DIVIDE(
    CALCULATE(COUNTROWS(fact_matches), fact_matches[Winner] = fact_matches[HomeTeam]),
    [Matches Played]
)

-- Vitórias de Virada
Comeback Wins = 
COUNTROWS(
    FILTER(
        fact_matches,
        OR(
            AND(fact_matches[Winner] <> fact_matches[HomeTeam], fact_matches[HomeGoals] > fact_matches[AwayGoals]),
            AND(fact_matches[Winner] <> fact_matches[AwayTeam], fact_matches[AwayGoals] > fact_matches[HomeGoals])
        )
    )
)

-- Conversão de Pênaltis
Penalty Conversion = 
DIVIDE(
    CALCULATE(COUNTROWS(fact_events), fact_events[EventType]="Goal", fact_events[Detail]="Penalty"),
    CALCULATE(COUNTROWS(fact_events), fact_events[Detail]="Penalty")
)

-- Top 3 Artilheiros (Golden Boot)
Golden Boot = 
TOPN(3,
    SUMMARIZE(
        dim_players,
        dim_players[Name],
        dim_players[CountryCode],
        "Goals", CALCULATE(
            SUMX(FILTER(fact_events, fact_events[EventType]="Goal"), fact_events[xG]),
            ALLEXCEPT(dim_players, dim_players[PlayerID])
        )
    ),
    [Goals], DESC
)
```

### Calculation Groups (Time Intelligence)
```dax
-- Calculation Group: Time Intelligence
-- Crie um Calculation Group chamado 'Time Intelligence' com os itens:

-- Current Period (Base)
-- YoY (Year over Year - 4 anos)
CALCULATE(SELECTEDMEASURE(), SAMEPERIODLASTYEAR(dim_world_cups[Year]))

-- YTD (Year to Date - não aplicável direto, use torneio atual)
CALCULATE(SELECTEDMEASURE(), dim_world_cups[Year] = MAX(dim_world_cups[Year]))

-- Previous Tournament
CALCULATE(SELECTEDMEASURE(), dim_world_cups[Year] = MAX(dim_world_cups[Year]) - 4)

-- Tournament Rank
RANKX(ALL(dim_world_cups[Year]), CALCULATE(SELECTEDMEASURE()), , DESC, Dense)
```

---

## 🎯 Especificação das 8 Páginas

### PÁGINA 1: Executive Overview
**Objetivo:** Visão executiva instantânea do torneio selecionado.

| Elemento | Tipo | Posição | Configuração |
|----------|------|---------|--------------|
| Slicer Copa | Dropdown | Topo Esq | dim_world_cups[Year] - Single Select, Default 2026 |
| Slicer Confederação | Dropdown | Topo Dir | dim_countries[Confederation] - Multi Select |
| KPI: Gols Totais | Card Grande | Linha 1, Col 1 | Valor: [Total Goals], Sparkline: Gols por Copa (últimas 5) |
| KPI: Partidas | Card Grande | Linha 1, Col 2 | Valor: [Matches Played], Target: 64/104 |
| KPI: Média Gols/Jogo | Card Grande | Linha 1, Col 3 | Valor: [Avg Goals Per Match], Formato: 0.00 |
| KPI: Valor Mercado Total | Card Grande | Linha 1, Col 4 | Valor: [Total Market Value], Formato: Bi€ |
| KPI: Artilheiro | Card Médio | Linha 2, Col 1 | Valor: [Top Scorer] |
| KPI: Melhor GK | Card Médio | Linha 2, Col 2 | Valor: [Best GK] |
| KPI: Idade Média | Card Médio | Linha 2, Col 3 | Valor: [Avg Squad Age], Formato: 0.0 |
| KPI: % Sub-23 | Card Médio | Linha 2, Col 4 | Valor: [Pct U23], Formato: 0.0% |
| Gráfico: Evolução Gols | Line Chart | Linha 3-4, Col 1-2 | Eixo X: Copa, Eixo Y: Gols, Legenda: Fase (Grupo/Knockout) |
| Gráfico: Valor Mercado | Clustered Column | Linha 3-4, Col 3-4 | Eixo X: Copa, Eixo Y: Valor (Bi€), Cor: Confederação |
| Matriz: Top 10 Países | Table | Linha 5-8, Col 1-4 | Colunas: País, Gols, xG, Partidas, Vitórias, % Vitórias, Rating Médio, Valor Mercado |

**Bookmarks:** Executive_2010, Executive_2014, Executive_2018, Executive_2022, Executive_2026 (botões de navegação no topo)

---

### PÁGINA 2: Tournament Deep Dive
**Objetivo:** Análise fase a fase com bracket interativo.

| Elemento | Tipo | Configuração |
|----------|------|--------------|
| Slicer Fase | Buttons/Bookmark | Grupo, Oitavas, Quartas, Semi, Final, 3º Lugar |
| Matriz Resultados | Matrix | Linhas: Mandante, Colunas: Visitante, Valores: Gols (Home/Away), Cores condicionais: Verde (Vitória), Vermelho (Derrota), Amarelo (Empate) |
| Bracket Visual | Custom Visual / Image | Imagem SVG do bracket com hyperlinks para drillthrough Partida |
| Gráfico Gols por Fase | Stacked Column | Eixo X: Fase, Eixo Y: Gols, Legenda: Tipo (Open Play, Penalty, Own Goal, Cabeça, Fora da Área) |
| Timeline Partidas | Timeline Custom | Data no eixo X, Partidas como bolhas (tamanho = gols, cor = fase) |
| Tabela Partidas | Table | Colunas: Data, Fase, Estádio, Mandante, Placar, Visitante, Vencedor, Público, xG H, xG A |
| Drillthrough → Página 6 (Match Timeline) | Right-click | Contexto: MatchID |

**Bookmarks por Fase:** Tournament_Group, Tournament_R16, Tournament_QF, Tournament_SF, Tournament_Final, Tournament_3rd

---

### PÁGINA 3: Player Explorer (Figurinha)
**Objetivo:** Perfil completo do jogador estilo "figurinha de álbum".

| Elemento | Tipo | Configuração |
|----------|------|--------------|
| Slicer Jogador | Dropdown Search | dim_players[Name] - Busca incremental, ordenado por Rating |
| Slicer Copa | Dropdown | dim_players[WorldCupYear] |
| Slicer Posição | Dropdown | dim_players[Position] |
| Slicer País | Dropdown | dim_players[CountryCode] |
| **Header Figurinha** | Card + Image | Nome grande, País (bandeira), Copa, Posição, Idade, Pé, Altura/Peso |
| **Rating Overall** | Gauge | 0-99, Zonas: <70/80/90/95, Valor: [OverallRating] |
| **Valor Mercado** | Card | Formato: "€ #,##0.0M" |
| **Radar Chart (6 Eixos)** | Radar Chart Custom | **Por Posição:**<br>• GK: Defesas, Clean Sheets, Jogo com Pés, Saídas, Pênaltis, Consistência<br>• DEF: Desarmes, Interceptações, Cortes, Jogo Aéreo, Passes, Disciplina<br>• MID: Passes, Progressão, Criação, Defesa, Resistência, Versatilidade<br>• FWD: Finalização, xG, Movimento, Jogo Aéreo, Criação, Clutch |
| **Stats Principais** | Multi-row Card | Partidas, Minutos, Gols, Assist, xG, Finaliz, No Alvo, Dribles, Passes%, Desarmes, Intercept, Faltas, Amarelos, Vermelhos |
| **Histórico Copas** | Table | Colunas: Copa, País, Idade, Partidas, Min, Gols, Assist, xG, Rating, Fase Alcançada |
| **Comparação Pares** | Scatter Plot | Eixo X: xG/90, Eixo Y: Gols/90, Tamanho: Minutos, Cor: Posição, Tooltip: Nome/País. Destacar jogador selecionado. |
| **Percentis** | Bullet Chart | Para cada métrica chave: mostrar P10, P50, P90, Valor Jogador |
| Drillthrough ← Página 1/2/4 | Contexto | PlayerID |

**Tooltips Personalizados:**
- Hover no Radar → Valor exato de cada eixo + percentil
- Hover no Scatter → Mini figurinha (nome, país, foto placeholder)

---

### PÁGINA 4: Team Profile
**Objetivo:** Perfil completo da seleção com formação tática.

| Elemento | Tipo | Configuração |
|----------|------|--------------|
| Slicer País | Dropdown | dim_countries[CountryName] com bandeira |
| Slicer Copa | Dropdown | dim_teams[WorldCupYear] |
| **Header Seleção** | Card | Nome, Confederação, Copa, Fase Alcançada, Técnico (placeholder) |
| **KPIs Seleção** | 4 Cards | Partidas, V-E-D, Gols Marcados, Gols Sofridos, Saldo |
| **Artilheiro / Garçom** | 2 Cards | Nome + Gols/Assist + xG |
| **Formação Tática** | Custom Visual / SVG | Campo 2D com 11 posições. Clicável → Drillthrough Player (Página 3). Cores por posição. |
| **Elenco Completo** | Table (23 linhas) | Colunas: #, Nome, Pos, Idade, Clube, Partidas, Min, Gols, Assist, xG, Rating, Valor (M€) |
| **Distribuição Idade** | Histogram | Bins: 18-20, 21-23, 24-26, 27-29, 30-32, 33-35, 36+ |
| **Valor Mercado por Posição** | Stacked Bar | Eixo Y: Posição, Eixo X: Valor Mercado (M€), Cor: Rating |
| **Minutos por Posição** | Donut Chart | Legenda: GK, DEF, MID, FWD |
| **Disciplina** | Table | Jogador, Amarelos, Vermelhos, Faltas Cometidas, Faltas Sofridas |
| **Evolução Histórica** | Line Chart | Eixo X: Copas (2010-2026), Eixo Y: Fase Alcançada (numérica), Marcador: Artilheiro da Copa |

**Drillthrough → Página 3 (Player Explorer)** via clique no elenco ou formação.

---

### PÁGINA 5: Stats Laboratory
**Objetivo:** Análise estatística profunda: Per 90, Percentis, Aging Curves, Projeções.

| Seção | Visual | Configuração |
|-------|--------|--------------|
| **Per 90 por Posição** | Table + Heatmap | Linhas: Posição, Colunas: Gols/90, Assist/90, xG/90, Finaliz/90, Passes/90, Prec%, Dribles/90, Desarmes/90, Intercept/90, Faltas/90. Formatação condicional: Verde (Alto) → Vermelho (Baixo) |
| **Percentis (Box Plot)** | Box & Whisker | Métricas: Gols, Assist, xG, Finaliz, Passes%, Desarmes, Dribles, Rating, Valor Mercado. Por Posição (Small Multiples) |
| **Aging Curves** | Line Chart | Eixo X: Idade (18-40), Eixo Y: Métrica (Gols/90, xG/90, Rating, Valor Mercado). Linhas separadas por Posição. Sombreamento: Pico Físico (25-29), Pico Técnico (27-32) |
| **Market Value vs Performance** | Scatter Plot | Eixo X: OverallRating, Eixo Y: MarketValue, Tamanho: Minutos, Cor: Idade, Forma: Posição. Linha de tendência. Outliers = Oportunidades de mercado |
| **What-If: Projeção 2030** | What-If Parameter + Card | Parâmetro: GrowthRate (-20% a +50%, step 5%). Medida: Projected Value = [Total Market Value] * (1 + GrowthRate). Slider visível. |
| **Decomposition Tree: Gols** | Decomposition Tree | Métrica: [Total Goals]. Explicar por: Copa → Fase → Tipo de Gol → Posição → País → Jogador |
| **Key Influencers: Mais Gols** | Key Influencers | Analisar: Goals (binário: >0). Explicar por: Posição, Idade, Minutos, xG, Finaliz, Rating, Valor Mercado, Copa |
| **Correlation Matrix** | Heatmap | Variáveis: Gols, Assist, xG, Finaliz, Passes%, Desarmes, Dribles, Idade, Valor, Rating. Cores: Azul (+1) → Vermelho (-1) |

---

### PÁGINA 6: Match Timeline
**Objetivo:** Linha do tempo interativa de todas as partidas com eventos.

| Elemento | Tipo | Configuração |
|----------|------|--------------|
| Slicer Copa | Dropdown | fact_matches[WorldCupYear] |
| Slicer Fase | Buttons | Grupo, Oitavas, Quartas, Semi, Final, 3º Lugar |
| Slicer Time | Dropdown | fact_matches[HomeTeam] + fact_matches[AwayTeam] (union via calculated table) |
| **Timeline Visual** | Timeline Custom | Eixo X: Data. Bolhas: Partidas. Tamanho = Total Goals. Cor = Fase. Hover = Tooltip customizado. |
| **Tooltip Report Page** | Report Page Tooltip | Page Size: Tooltip. Conteúdo: Mandante/Visitante, Placar, Estádio, Gols (minuto, jogador, tipo, xG), Cartões, Substituições, xG total. |
| **Tabela Partidas** | Table | Data, Fase, Mandante, Gols, Visitante, Vencedor, Público, xG H, xG A. Ordenável. |
| **Decomposition Tree: Gols por Partida** | Decomposition Tree | Raiz: Partida Selecionada. Níveis: Time → Jogador → Tipo de Gol → Minuto (1T/2T/Prorrogação/Pênaltis) |
| **Estatísticas da Partida** | Cards | Posse estimada, Finalizações, Escanteios, Faltas, Cartões, Impedimentos, xG Diff |

**Drillthrough ← Página 2 (Tournament)** via MatchID.

---

### PÁGINA 7: Comparative Analytics
**Objetivo:** Comparar 2+ Copas lado a lado.

| Elemento | Tipo | Configuração |
|----------|------|--------------|
| **Multi-Select Copas** | Slicer (Checkbox) | dim_world_cups[Year] - Permitir múltipla seleção |
| **Small Multiples: Gols** | Line Chart (Small Multiples) | Eixo X: Fase, Eixo Y: Gols/Partida, Painel: Copa. Uma linha por Copa selecionada. |
| **Small Multiples: xG** | Line Chart (Small Multiples) | Eixo X: Fase, Eixo Y: xG/Partida, Painel: Copa |
| **Small Multiples: Disciplina** | Clustered Bar (Small Multiples) | Eixo X: Copa, Eixo Y: Cartões/Partida, Legenda: Amarelo/Vermelho |
| **Small Multiples: Idade** | Line Chart (Small Multiples) | Eixo X: Copa, Eixo Y: Idade Média, Linhas: Por Posição |
| **Small Multiples: Valor Mercado** | Area Chart (Small Multiples) | Eixo X: Copa, Eixo Y: Valor Total (Bi€), Painel: Confederação |
| **Tabela Comparativa** | Matrix | Linhas: Métrica (Gols, xG, Amarelos, Vermelhos, Idade Média, Valor Mercado, Rating, %Sub23, CleanSheets, HomeWin%), Colunas: Copas Selecionadas, Valores: Medida. Formatação: Heatmap por coluna. |
| **Key Influencers: O que explica Gols?** | Key Influencers | Analisar: Goals (contínuo). Explicar por: Copa, Fase, Mandante/Visitante, xG, Posse, Finalizações, Escanteios. |
| **Q&A Visual** | Q&A | Perguntas naturais: "Quem marcou mais gols em 2022?", "Qual país tem melhor xG/90?", "Evolução de cartões vermelhos" |

**Bookmarks:** Compare_2010_2014, Compare_2018_2022, Compare_All, Compare_Champions

---

### PÁGINA 8: Scouting & Projections
**Objetivo:** Identificar talentos, projetar 2030, necessidades por posição.

| Seção | Visual | Configuração |
|-------|--------|--------------|
| **Jovens Talentos (Sub-23)** | Table + Conditional Formatting | Filtro: Age < 23. Colunas: Nome, País, Copa, Pos, Idade, Min, Gols/90, xG/90, Rating, Valor, Potencial (Rating + (25-Idade)*0.5). Cores: Verde (Alto potencial) → Vermelho. Ordenar por Potencial DESC. |
| **Projeção 2030** | Line Chart + Forecast | Eixo X: Copa (2010-2026 + 2030 projetado). Eixo Y: Métrica (Gols Totais, Valor Mercado, Idade Média, %Sub23). Linha histórica + Forecast (ETS). Parâmetro What-If: Taxa Crescimento. |
| **Transfer Value Estimator** | Scatter + What-If | Eixo X: Performance Score (Gols+Assist+xG/90 * Rating). Eixo Y: MarketValue. Linha Regressão. Parâmetro: Multiplier (0.5-3.0). Card: Estimated Value = PerformanceScore * Multiplier. |
| **Positional Need by Team** | Matrix + Heatmap | Linhas: País. Colunas: Posição. Valores: Rating Médio dos jogadores da posição. Cores: Vermelho (Fraco) → Verde (Forte). Identificar lacunas. |
| **Sucessão de Estrelas** | Table | Jogadores 30+ com alto Rating → Jovens mesma posição/país com potencial. Colunas: Estrela (Nome, Idade, Rating), Sucessor (Nome, Idade, Potencial, Gap Anos). |
| **RLS: Confederação** | Roles | Criar Roles: CONMEBOL, UEFA, CAF, AFC, CONCACAF, OFC. Filtro: dim_countries[Confederation] = USERPRINCIPALNAME() mapeado via tabela de usuários. Testar com View as Roles. |

---

## 🎨 Tema Visual e Branding

### Paleta de Cores Oficial
```json
{
  "primary": "#1B4F72",
  "secondary": "#2E86C1",
  "accent": "#E74C3C",
  "success": "#27AE60",
  "warning": "#F39C12",
  "info": "#3498DB",
  "background": "#F8F9FA",
  "surface": "#FFFFFF",
  "text_primary": "#2C3E50",
  "text_secondary": "#7F8C8D",
  "grid": "#E0E0E0",
  "confederations": {
    "UEFA": "#3498DB",
    "CONMEBOL": "#F39C12",
    "CAF": "#27AE60",
    "AFC": "#9B59B6",
    "CONCACAF": "#E67E22",
    "OFC": "#1ABC9C"
  },
  "position_colors": {
    "GK": "#E74C3C",
    "CB": "#3498DB",
    "RB": "#2980B9",
    "LB": "#2980B9",
    "CDM": "#27AE60",
    "CM": "#2ECC71",
    "CAM": "#F39C12",
    "RW": "#E67E22",
    "LW": "#E67E22",
    "ST": "#E74C3C"
  }
}
```

### Tipografia
- **Títulos:** Segoe UI Bold, 24pt, #1B4F72
- **Subtítulos:** Segoe UI SemiBold, 16pt, #2E86C1
- **Corpo:** Segoe UI Regular, 11pt, #2C3E50
- **KPIs:** Segoe UI Bold, 28pt, Cores por tipo
- **Tabelas:** Segoe UI, 10pt, Header Bold

### Ícones e Imagens
- Bandeiras: Pasta assets/flags/{CountryCode}.svg (200x130px)
- Fotos jogadores: assets/players/{PlayerID}.jpg (placeholder se não existir)
- Estádio: assets/stadiums/{StadiumID}.jpg
- Logo Copa: assets/logos/wc{Year}.png

---

## 🔧 Configurações Técnicas

### Performance
- **Incremental Refresh:** fact_events (por WorldCupYear), fact_matches (por WorldCupYear)
- **Agregações:** Tabela agregada agg_matches_by_team_cup (TeamCode, WorldCupYear, Goals, xG, Matches, W-D-L)
- **DirectQuery:** Apenas para dados em tempo real (não aplicável aqui)
- **VertiPaq:** Otimizar colunas de alta cardinalidade (PlayerID, MatchID) com dicionários

### Segurança (RLS)
```dax
-- Tabela de usuários (carregar via Excel/CSV)
-- Colunas: UserPrincipalName, Confederation, Role

-- Role: CONMEBOL
[CountryCode] IN SELECTCOLUMNS(
    FILTER(Users, Users[Confederation] = "CONMEBOL" && Users[UserPrincipalName] = USERPRINCIPALNAME()),
    "Code", Users[CountryCode]
)

-- Role: UEFA
[CountryCode] IN SELECTCOLUMNS(
    FILTER(Users, Users[Confederation] = "UEFA" && Users[UserPrincipalName] = USERPRINCIPALNAME()),
    "Code", Users[CountryCode]
)
-- Repetir para CAF, AFC, CONCACAF, OFC
-- Role: Admin (sem filtro)
```

### Parâmetros What-If
| Nome | Tipo | Mín | Máx | Passo | Padrão | Uso |
|------|------|-----|-----|-------|--------|-----|
| GrowthRate | Decimal | -0.2 | 0.5 | 0.05 | 0.1 | Projeção Valor Mercado 2030 |
| Multiplier | Decimal | 0.5 | 3.0 | 0.1 | 1.5 | Transfer Value Estimator |
| MinMinutes | Whole | 0 | 540 | 30 | 90 | Filtro Per 90 |
| TopN | Whole | 3 | 20 | 1 | 10 | Top N Jogadores/Equipes |

---

## 📱 Navegação e UX

### Bookmarks Principais
| Bookmark | Página | Filtros | Descrição |
|----------|--------|---------|-----------|
| Home | 1 | Copa=2026 | Landing page |
| Executive_2010 | 1 | Copa=2010 | Visão 2010 |
| Executive_2014 | 1 | Copa=2014 | Visão 2014 |
| Executive_2018 | 1 | Copa=2018 | Visão 2018 |
| Executive_2022 | 1 | Copa=2022 | Visão 2022 |
| Executive_2026 | 1 | Copa=2026 | Visão 2026 |
| Tournament_Group | 2 | Fase=Grupo | Fase de Grupos |
| Tournament_Knockout | 2 | Fase≠Grupo | Mata-mata-mata |
| Player_Messi | 3 | Player=Messi | Demo figurinha |
| Team_Brazil | 4 | Country=Brazil | Demo seleção |
| Compare_Champions | 7 | Copas=Campeãs | Comparar campeãs |
| Scouting_U23 | 8 | Age<23 | Jovens talentos |

### Botões de Navegação (Action Buttons)
- **Topo fixo (todas páginas):** Botões ícone: Home | Executive | Tournament | Players | Teams | Stats | Matches | Compare | Scouting
- **Botão Voltar (drillthrough):** "← Voltar" no canto superior direito das páginas 3, 4, 6
- **Botão Reset Filtros:** "Limpar Filtros" em cada página

### Tooltips Personalizados
| Tooltip Page | Gatilho | Conteúdo |
|--------------|---------|----------|
| Tooltip_Player | Hover Player (Qualquer gráfico) | Mini figurinha: Foto, Nome, País, Pos, Idade, Gols, Assist, xG, Rating, Valor, Sparkline Histórico |
| Tooltip_Match | Hover Partida (Timeline) | Placar, Data, Estádio, Gols (min/jogador/tipo), xG, Cartões, Público |
| Tooltip_Team | Hover País (Mapa/Scatter) | Bandeira, Copa Atual, Fase, Artilheiro, Rating Médio, Valor Total, Últimas 3 Copas |

---

## 📦 Entregáveis para Publicação

### Arquivos Necessários
```
Album_Copa_Definitivo/
├── Album_Copa_Definitivo.pbix          # Arquivo principal Power BI
├── data/
│   ├── dim_world_cups.csv
│   ├── dim_countries.csv
│   ├── dim_teams.csv
│   ├── dim_players.csv
│   ├── fact_matches.csv
│   └── fact_events.csv
├── assets/
│   ├── flags/                          # 63 SVGs de bandeiras
│   ├── players/                        # Placeholders ou fotos reais
│   ├── stadiums/                       # Fotos estádios
│   └── logos/                          # Logos das 5 Copas
├── docs/
│   ├── README.md                       # Este arquivo
│   ├── DAX_MEASURES.md                 # Medidas DAX documentadas
│   ├── DATA_DICTIONARY.md              # Dicionário de dados
│   └── DEPLOYMENT_GUIDE.md             # Guia publicação
└── Album_Copa_Definitivo_2010_2026.xlsx # Excel interativo (referência)
```

### Checklist Pré-Publicação
- [ ] Todos os relacionamentos ativos (1:N, Single Cross-filter)
- [ ] Medidas DAX testadas com DAX Studio / Performance Analyzer
- [ ] RLS testado com View as Roles para cada confederação
- [ ] Bookmarks funcionando (navegação, filtros, visibilidade)
- [ ] Drillthrough configurado (Página 3←1,2,4 | Página 6←2)
- [ ] Tooltips personalizados vinculados corretamente
- [ ] What-If parameters funcionando nos cards/gráficos
- [ ] Decomposition Tree e Key Influencers renderizando
- [ ] Formatação condicional nas tabelas (heatmap, data bars, icons)
- [ ] Responsividade testada: Desktop, Tablet, Mobile (Power BI App)
- [ ] Tema JSON importado e aplicado
- [ ] Dicionário de dados documentado
- [ ] README com instruções de uso

---

## 🚀 Publicação no Power BI Service

### Workspace
1. Criar workspace: Album Copa Definitivo - Portfolio
2. Publicar .pbix
3. Configurar Scheduled Refresh: Diário 06:00 UTC (dados estáticos, mas boa prática)
4. Configurar Row Level Security: Adicionar usuários às roles por confederação
5. Criar App para distribuição: Album Copa Definitivo - Public
6. Configurar Embed para portfólio web (iframe seguro)

### Métricas de Sucesso (Portfólio Freelance)
- **Complexidade:** 8 páginas, 50+ visuais, 20+ medidas DAX, RLS, Bookmarks, Drillthrough, What-If, Decomposition Tree, Key Influencers
- **Técnicas Avançadas:** Calculation Groups, Aggregations, Incremental Refresh, Custom Tooltips, Small Multiples, Forecasting
- **Domínio:** Futebol/Esportes - alto apelo visual, dados reais, storytelling claro
- **Diferencial:** Álbum interativo + Excel VBA + Dataset completo + Documentação técnica

---

## 📚 Referências e Recursos

- Power BI Documentation - Bookmarks
- Power BI Documentation - Drillthrough
- Power BI Documentation - What-If Parameters
- DAX Guide - Time Intelligence
- SQLBI - Calculation Groups
- Power BI Tips - Decomposition Tree
- FIFA World Cup Statistics
- FBref - World Cup Data

---

**Desenvolvido por:** Investigation Team AI (Rise Kujikawa - Orquestradora)  
**Stack:** Python (pandas/numpy/xlsxwriter) → CSV → Power BI Desktop → DAX → Excel VBA  
**Dados:** 5 Copas (2010-2026) | 4.048 jogadores | 344 partidas | 4.846 eventos  
**Licença:** Portfolio Demonstration - Dados simulados baseados em estatísticas reais