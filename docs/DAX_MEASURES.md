# 📐 Medidas DAX - Álbum da Copa Definitivo 2010-2026

Todas as medidas prontas para copiar/colar no Power BI Desktop. Organizadas por categoria.

---

## 1️⃣ KPIs BÁSICOS

### Total de Gols (baseado em xG dos eventos Goal)
```dax
Total Goals = 
SUMX(
    FILTER(fact_events, fact_events[EventType] = "Goal"),
    fact_events[xG]
)
```

### Partidas Distintas
```dax
Matches Played = DISTINCTCOUNT(fact_matches[MatchID])
```

### Média de Gols por Partida
```dax
Avg Goals Per Match = DIVIDE([Total Goals], [Matches Played])
```

### Artilheiro da Copa Selecionada
```dax
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
```

### Melhor Goleiro (Mais Defesas)
```dax
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
```

### Valor de Mercado Total do Elenco
```dax
Total Market Value = SUMX(dim_players, dim_players[MarketValue])
```

### Rating Médio do Elenco
```dax
Avg Rating = AVERAGE(dim_players[OverallRating])
```

### Idade Média do Elenco
```dax
Avg Squad Age = AVERAGE(dim_players[Age])
```

### % Jogadores Sub-23
```dax
Pct U23 = 
DIVIDE(
    COUNTROWS(FILTER(dim_players, dim_players[Age] < 23)),
    COUNTROWS(dim_players)
)
```

### Gols por 90 Minutos
```dax
Goals Per 90 = 
DIVIDE(
    SUMX(FILTER(fact_events, fact_events[EventType]="Goal"), fact_events[xG]) * 90,
    SUM(dim_players[Minutes])
)
```

### xG por 90 Minutos
```dax
xG Per 90 = 
DIVIDE(
    SUM(dim_players[xG]) * 90,
    SUM(dim_players[Minutes])
)
```

### Clean Sheets (Jogos sem sofrer gols)
```dax
Clean Sheets = 
CALCULATE(
    DISTINCTCOUNT(fact_matches[MatchID]),
    OR(fact_matches[HomeGoals] = 0, fact_matches[AwayGoals] = 0)
)
```

### Precisão de Passes Média
```dax
Pass Accuracy = AVERAGE(dim_players[PassAccuracy])
```

### Cartões por Partida
```dax
Cards Per Match = 
DIVIDE(
    SUM(dim_players[YellowCards]) + SUM(dim_players[RedCards]) * 3,
    [Matches Played]
)
```

---

## 2️⃣ EVOLUÇÃO TEMPORAL (Time Intelligence)

### Evolução Valor de Mercado (YoY - 4 anos)
```dax
Market Value YoY = 
VAR CurrYear = SELECTEDVALUE(dim_world_cups[Year])
VAR PrevYear = CurrYear - 4
RETURN
    DIVIDE(
        CALCULATE([Total Market Value], dim_world_cups[Year] = CurrYear),
        CALCULATE([Total Market Value], dim_world_cups[Year] = PrevYear)
    ) - 1
```

### Impacto no Torneio (Performance Média)
```dax
Tournament Impact = 
DIVIDE(
    SUMX(dim_players, dim_players[Goals] + dim_players[Assists] + dim_players[xG]),
    COUNTROWS(dim_players)
)
```

### Idade do Pico de Performance
```dax
Peak Age = 
CALCULATE(
    AVERAGE(dim_players[Age]),
    TOPN(10, dim_players, dim_players[OverallRating], DESC)
)
```

### Vantagem de Casa
```dax
Home Win % = 
DIVIDE(
    CALCULATE(COUNTROWS(fact_matches), fact_matches[Winner] = fact_matches[HomeTeam]),
    [Matches Played]
)
```

### Vitórias de Virada
```dax
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
```

### Conversão de Pênaltis
```dax
Penalty Conversion = 
DIVIDE(
    CALCULATE(COUNTROWS(fact_events), fact_events[EventType]="Goal", fact_events[Detail]="Penalty"),
    CALCULATE(COUNTROWS(fact_events), fact_events[Detail]="Penalty")
)
```

### Top 3 Artilheiros (Golden Boot)
```dax
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

---

## 3️⃣ POR JOGADOR (Player Level)

### Gols do Jogador
```dax
Player Goals = 
CALCULATE(
    SUMX(FILTER(fact_events, fact_events[EventType]="Goal"), fact_events[xG]),
    ALLEXCEPT(dim_players, dim_players[PlayerID])
)
```

### Assistências do Jogador
```dax
Player Assists = 
CALCULATE(
    SUM(dim_players[Assists]),
    ALLEXCEPT(dim_players, dim_players[PlayerID])
)
```

### xG do Jogador
```dax
Player xG = 
CALCULATE(
    SUM(dim_players[xG]),
    ALLEXCEPT(dim_players, dim_players[PlayerID])
)
```

### Minutos do Jogador
```dax
Player Minutes = 
CALCULATE(
    SUM(dim_players[Minutes]),
    ALLEXCEPT(dim_players, dim_players[PlayerID])
)
```

### Gols/90 do Jogador
```dax
Player Goals Per 90 = 
DIVIDE([Player Goals] * 90, [Player Minutes])
```

### xG/90 do Jogador
```dax
Player xG Per 90 = 
DIVIDE([Player xG] * 90, [Player Minutes])
```

### Partidas do Jogador
```dax
Player Matches = 
CALCULATE(
    SUM(dim_players[Matches]),
    ALLEXCEPT(dim_players, dim_players[PlayerID])
)
```

### Rating do Jogador
```dax
Player Rating = 
CALCULATE(
    AVERAGE(dim_players[OverallRating]),
    ALLEXCEPT(dim_players, dim_players[PlayerID])
)
```

### Valor de Mercado do Jogador
```dax
Player Market Value = 
CALCULATE(
    MAX(dim_players[MarketValue]),
    ALLEXCEPT(dim_players, dim_players[PlayerID])
)
```

---

## 4️⃣ POR SELEÇÃO (Team Level)

### Gols da Seleção
```dax
Team Goals = 
CALCULATE(
    SUMX(
        FILTER(fact_events, fact_events[EventType]="Goal" && fact_events[TeamCode]=SELECTEDVALUE(dim_teams[TeamCode])),
        fact_events[xG]
    )
)
```

### Valor de Mercado da Seleção
```dax
Team Market Value = 
SUMX(
    FILTER(
        dim_players, 
        dim_players[CountryCode]=SELECTEDVALUE(dim_teams[TeamCode]) && 
        dim_players[WorldCupYear]=SELECTEDVALUE(dim_teams[WorldCupYear])
    ),
    dim_players[MarketValue]
)
```

### Rating Médio da Seleção
```dax
Team Avg Rating = 
AVERAGEX(
    FILTER(
        dim_players, 
        dim_players[CountryCode]=SELECTEDVALUE(dim_teams[TeamCode]) && 
        dim_players[WorldCupYear]=SELECTEDVALUE(dim_teams[WorldCupYear])
    ),
    dim_players[OverallRating]
)
```

### Idade Média da Seleção
```dax
Team Avg Age = 
AVERAGEX(
    FILTER(
        dim_players, 
        dim_players[CountryCode]=SELECTEDVALUE(dim_teams[TeamCode]) && 
        dim_players[WorldCupYear]=SELECTEDVALUE(dim_teams[WorldCupYear])
    ),
    dim_players[Age]
)
```

### Artilheiro da Seleção
```dax
Team Top Scorer = 
MAXX(
    TOPN(1,
        SUMMARIZE(
            FILTER(
                dim_players, 
                dim_players[CountryCode]=SELECTEDVALUE(dim_teams[TeamCode]) && 
                dim_players[WorldCupYear]=SELECTEDVALUE(dim_teams[WorldCupYear])
            ),
            dim_players[Name],
            "Goals", CALCULATE(
                SUMX(FILTER(fact_events, fact_events[EventType]="Goal"), fact_events[xG]),
                ALLEXCEPT(dim_players, dim_players[PlayerID])
            )
        ),
        [Goals], DESC
    ),
    [Name] & " - " & FORMAT([Goals], "0.0")
)
```

### Melhor Jogador da Seleção
```dax
Team Best Player = 
MAXX(
    TOPN(1,
        FILTER(
            dim_players, 
            dim_players[CountryCode]=SELECTEDVALUE(dim_teams[TeamCode]) && 
            dim_players[WorldCupYear]=SELECTEDVALUE(dim_teams[WorldCupYear])
        ),
        dim_players[OverallRating], DESC
    ),
    dim_players[Name] & " (★" & dim_players[OverallRating] & ")"
)
```

### Distribuição de Minutos por Posição
```dax
Team Minutes By Position = 
SUMMARIZE(
    FILTER(
        dim_players, 
        dim_players[CountryCode]=SELECTEDVALUE(dim_teams[TeamCode]) && 
        dim_players[WorldCupYear]=SELECTEDVALUE(dim_teams[WorldCupYear])
    ),
    dim_players[Position],
    "TotalMinutes", SUM(dim_players[Minutes]),
    "PctMinutes", DIVIDE(SUM(dim_players[Minutes]), CALCULATE(SUM(dim_players[Minutes]), ALL(dim_players[Position])))
)
```

---

## 5️⃣ CALCULATION GROUPS (Time Intelligence)

Criar um **Calculation Group** chamado "Time Intelligence" com os seguintes itens:

### Current Period (Base)
```dax
-- Sem modificação, retorna a medida original
SELECTEDMEASURE()
```

### Previous Tournament (4 anos atrás)
```dax
CALCULATE(
    SELECTEDMEASURE(),
    dim_world_cups[Year] = MAX(dim_world_cups[Year]) - 4
)
```

### Tournament Rank (Ranking entre Copas)
```dax
RANKX(
    ALL(dim_world_cups[Year]),
    CALCULATE(SELECTEDMEASURE()),
    , DESC, Dense
)
```

### Tournament Growth (Crescimento vs Copa anterior)
```dax
VAR CurrYear = SELECTEDVALUE(dim_world_cups[Year])
VAR PrevYear = CurrYear - 4
VAR CurrVal = CALCULATE(SELECTEDMEASURE(), dim_world_cups[Year] = CurrYear)
VAR PrevVal = CALCULATE(SELECTEDMEASURE(), dim_world_cups[Year] = PrevYear)
RETURN
    DIVIDE(CurrVal - PrevVal, PrevVal)
```

### Best Tournament (Melhor Copa para a métrica)
```dax
MAXX(
    ALL(dim_world_cups[Year]),
    CALCULATE(SELECTEDMEASURE())
)
```

---

## 6️⃣ PARÂMETROS WHAT-IF

Criar 4 parâmetros what-if no Power BI:

| Nome | Tipo | Mín | Máx | Passo | Padrão |
|------|------|-----|-----|-------|--------|
| GrowthRate | Decimal | -0.2 | 0.5 | 0.05 | 0.10 |
| Multiplier | Decimal | 0.5 | 3.0 | 0.1 | 1.5 |
| MinMinutes | Inteiro | 0 | 540 | 30 | 90 |
| TopN | Inteiro | 3 | 20 | 1 | 10 |

### Uso nas Medidas

#### Valor Projetado 2030
```dax
Projected Market Value 2030 = [Total Market Value] * (1 + GrowthRate[GrowthRate Value])
```

#### Transfer Value Estimator
```dax
Estimated Transfer Value = [Player Performance Score] * Multiplier[Multiplier Value]
```

#### Filtro Per 90 (Mínimo Minutos)
```dax
Filtered Goals Per 90 = 
CALCULATE(
    [Goals Per 90],
    FILTER(dim_players, dim_players[Minutes] >= MinMinutes[MinMinutes Value])
)
```

#### Top N Dinâmico
```dax
Dynamic Top N = 
TOPN(
    TopN[TopN Value],
    SUMMARIZE(dim_players, dim_players[Name], dim_players[CountryCode], "Metric", [Selected Metric]),
    [Metric], DESC
)
```

---

## 7️⃣ RLS (Row Level Security)

### Tabela de Usuários (carregar via Excel/CSV)
```dax
Users = 
DATATABLE(
    "UserPrincipalName", STRING,
    "Confederation", STRING,
    "Role", STRING,
    {
        {"analyst@uefa.com", "UEFA", "Analyst"},
        {"analyst@conmebol.com", "CONMEBOL", "Analyst"},
        {"analyst@caf.com", "CAF", "Analyst"},
        {"analyst@afc.com", "AFC", "Analyst"},
        {"analyst@concacaf.com", "CONCACAF", "Analyst"},
        {"analyst@ofc.com", "OFC", "Analyst"},
        {"admin@fifa.com", "ALL", "Admin"}
    }
)
```

### Roles DAX

#### Role: UEFA
```dax
dim_countries[Confederation] = "UEFA"
```

#### Role: CONMEBOL
```dax
dim_countries[Confederation] = "CONMEBOL"
```

#### Role: CAF
```dax
dim_countries[Confederation] = "CAF"
```

#### Role: AFC
```dax
dim_countries[Confederation] = "AFC"
```

#### Role: CONCACAF
```dax
dim_countries[Confederation] = "CONCACAF"
```

#### Role: OFC
```dax
dim_countries[Confederation] = "OFC"
```

#### Role: Admin (sem filtro)
```dax
-- Nenhum filtro, vê tudo
TRUE()
```

---

## 8️⃣ MEDIDAS PARA VISUAIS ESPECÍFICOS

### Decomposition Tree - Gols
```dax
Goals Decomposition = [Total Goals]
-- Explicar por: WorldCupYear → Stage → EventType (Detail) → Position → CountryCode → Player Name
```

### Key Influencers - O que explica Gols
```dax
Goals Influencers = [Total Goals]
-- Explicar por: Position, Age, Minutes, xG, Shots, Rating, MarketValue, WorldCupYear
```

### Forecast - Projeção 2030
```dax
Forecast Goals 2030 = 
-- Usar visual de previsão nativo do Power BI na linha de tendência de Goals por Copa
-- Configurar: Temporada = dim_world_cups[Year], Valor = [Total Goals], Previsão = 1 passo (2030)
```

### Small Multiples - Comparação Copas
```dax
Goals By Stage By Cup = 
SUMMARIZE(
    fact_matches,
    dim_world_cups[Year],
    fact_matches[Stage],
    "AvgGoals", AVERAGEX(fact_matches, fact_matches[TotalGoals])
)
```

---

## 📋 CHECKLIST DE IMPLEMENTAÇÃO

- [ ] Criar todas as medidas em **Modeling → New Measure**
- [ ] Organizar em **Display Folders**: "KPIs", "Player", "Team", "Time Intelligence", "Advanced"
- [ ] Criar **Calculation Group** "Time Intelligence"
- [ ] Criar **4 What-If Parameters**
- [ ] Configurar **RLS** com 7 roles
- [ ] Testar cada medida com **DAX Studio** / Performance Analyzer
- [ ] Documentar dependências entre medidas
- [ ] Validar resultados com dados conhecidos (ex: Messi 2022 = 7 gols, Mbappe 2022 = 8 gols)

---

**Versão:** 1.0 | **Data:** 2026 | **Autor:** Investigation Team AI (Rise Kujikawa)