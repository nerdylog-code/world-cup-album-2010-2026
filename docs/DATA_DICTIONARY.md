# 📚 Dicionário de Dados - Álbum da Copa Definitivo 2010-2026

## Visão Geral
Modelo dimensional em estrela (Star Schema) com 5 dimensões e 2 tabelas de fato, cobrindo 5 Copas do Mundo (2010-2026) com 4.048 jogadores, 344 partidas e 4.846 eventos.

---

## Tabelas de Dimensão

### 1. dim_world_cups (5 linhas)
| Coluna | Tipo | PK | Descrição | Valores Exemplo |
|--------|------|----|-----------|-----------------|
| Year | Integer | ✓ | Ano da Copa | 2010, 2014, 2018, 2022, 2026 |
| Host | String | | País sede | "Africa do Sul", "Brasil", "Russia", "Catar", "EUA/Mexico/Canada" |
| Winner | String | | Campeão (FK País) | "ESP", "GER", "FRA", "ARG", "ARG" |
| RunnerUp | String | | Vice-campeão | "NED", "ARG", "CRO", "FRA", "BRA" |
| Third | String | | 3º lugar | "GER", "NED", "BEL", "CRO", "FRA" |
| Fourth | String | | 4º lugar | "URU", "BRA", "ENG", "MAR", "ESP" |
| Teams | Integer | | Seleções participantes | 32 (2010-2022), 48 (2026) |
| Matches | Integer | | Partidas totais | 64 (2010-2022), 104 (2026) |
| Goals | Integer | | Gols totais | 145, 171, 169, 172, 287 |

### 2. dim_countries (63 linhas)
| Coluna | Tipo | PK | Descrição | Valores |
|--------|------|----|-----------|---------|
| CountryCode | String | ✓ | Código FIFA (3 letras) | "BRA", "ARG", "FRA", "ESP", "GER", etc. |
| CountryName | String | | Nome completo | "Brasil", "Argentina", "França", "Espanha", "Alemanha" |
| Confederation | String | | Confederação | "UEFA", "CONMEBOL", "CAF", "AFC", "CONCACAF", "OFC" |

**Confederações por país:**
- UEFA (Europa): GER, ESP, FRA, ENG, POR, NED, ITA, BEL, CRO, etc.
- CONMEBOL (América do Sul): BRA, ARG, URU, COL, CHI, PAR, ECU, PER, BOL, VEN
- CAF (África): SEN, MAR, TUN, EGY, ALG, NGA, CIV, GHA, CMR, MLI, etc.
- AFC (Ásia): JPN, KOR, IRN, AUS, QAT, KSA, IRQ, UZB, JOR, OMA, BHR
- CONCACAF (América do Norte/Central): MEX, USA, CAN, CRC, PAN, JAM, HON, TRI
- OFC (Oceania): NZL

### 3. dim_teams (176 linhas = 32×4 + 48)
| Coluna | Tipo | PK | Descrição |
|--------|------|----|-----------|
| TeamCode | String | ✓* | Código do time (igual a CountryCode) |
| WorldCupYear | Integer | ✓* | Ano da Copa |

*PK composta: (TeamCode, WorldCupYear)

Cada linha representa uma seleção participando de uma Copa específica.

### 4. dim_players (4.048 linhas = 23 jogadores × 176 seleções)
| Coluna | Tipo | PK/FK | Descrição | Range/Valores |
|--------|------|-------|-----------|---------------|
| PlayerID | Integer | ✓ | ID único sequencial | 1-4048 |
| Name | String | | Nome do jogador | "Messi", "Mbappe", "Neymar", etc. |
| CountryCode | String | FK | País (→ dim_countries) | "BRA", "ARG", "FRA", etc. |
| WorldCupYear | Integer | FK | Ano da Copa (→ dim_world_cups) | 2010, 2014, 2018, 2022, 2026 |
| Position | String | | Posição principal | GK, CB, RB, LB, CDM, CM, CAM, RW, LW, ST |
| Age | Integer | | Idade na Copa | 17-42 |
| BirthYear | Integer | | Ano de nascimento | 1968-2009 |
| Height | Integer | | Altura (cm) | 160-202 |
| Weight | Integer | | Peso (kg) | 55-100 |
| PreferredFoot | String | | Pé preferido | "Right", "Left", "Both" |
| MarketValue | Decimal | | Valor de mercado (milhões €) | 0.5-250.0 |
| OverallRating | Integer | | Rating geral 60-99 | 60-99 |
| Matches | Integer | | Partidas jogadas | 0-7 |
| Minutes | Integer | | Minutos em campo | 0-630 |
| Goals | Decimal | | Gols (baseado em xG) | 0.00-8.00 |
| Assists | Decimal | | Assistências | 0.00-5.00 |
| Saves | Decimal | | Defesas (GK) | 0.00-30.00 |
| CleanSheets | Decimal | | Jogos sem sofrer gols | 0.00-7.00 |
| GoalsConceded | Decimal | | Gols sofridos (GK) | 0.00-10.00 |
| PassAccuracy | Decimal | | % passes certos | 60.0-95.0 |
| Tackles | Decimal | | Desarmes | 0.00-25.00 |
| Interceptions | Decimal | | Interceptações | 0.00-20.00 |
| Clearances | Decimal | | Cortes | 0.00-40.00 |
| Errors | Decimal | | Erros que levaram a gol | 0.00-5.00 |
| PenaltiesSaved | Decimal | | Pênaltis defendidos | 0.00-3.00 |
| xG | Decimal | | Expected Goals | 0.00-10.00 |
| Shots | Decimal | | Finalizações | 0.00-40.00 |
| ShotsOnTarget | Decimal | | Finalizações no alvo | 0.00-20.00 |
| Dribbles | Decimal | | Dribles completados | 0.00-40.00 |
| Fouls | Decimal | | Faltas cometidas | 0.00-30.00 |
| YellowCards | Integer | | Cartões amarelos | 0-4 |
| RedCards | Integer | | Cartões vermelhos | 0-1 |

**Distribuição por posição (por seleção - 23 jogadores):**
| Posição | Quantidade | Função Principal |
|---------|------------|------------------|
| GK | 3 | Goleiros |
| CB | 4 | Zagueiros centrais |
| RB | 2 | Laterais direitos |
| LB | 2 | Laterais esquerdos |
| CDM | 2 | Volantes |
| CM | 3 | Meias centrais |
| CAM | 1 | Meia atacante |
| RW | 2 | Pontas direitos |
| LW | 2 | Pontas esquerdos |
| ST | 2 | Centroavantes |

---

## Tabelas de Fato

### 5. fact_matches (344 linhas)
| Coluna | Tipo | PK/FK | Descrição |
|--------|------|-------|-----------|
| MatchID | Integer | ✓ | ID único da partida |
| WorldCupYear | Integer | FK | Ano (→ dim_world_cups) |
| Stage | String | | Fase | Group, Round16, Quarter, Semi, Final, ThirdPlace |
| Date | Date | | Data da partida | 2010-06-11 a 2026-07-19 |
| Stadium | String | | Estádio | "Soccer City", "Maracanã", "Luzhniki", etc. |
| HomeTeam | String | FK | Time mandante (→ dim_teams) | "BRA", "ARG", etc. |
| AwayTeam | String | FK | Time visitante (→ dim_teams) | "GER", "FRA", etc. |
| HomeGoals | Integer | | Gols mandante | 0-7 |
| AwayGoals | Integer | | Gols visitante | 0-7 |
| TotalGoals | Integer | | Total gols | 0-10 |
| Winner | String | | Vencedor | HomeTeam, AwayTeam, "Draw" |
| Attendance | Integer | | Público | 30.000-80.000 |
| xG_Home | Decimal | | xG mandante | 0.0-4.0 |
| xG_Away | Decimal | | xG visitante | 0.0-4.0 |

**Distribuição por Copa:**
| Copa | Fase Grupos | Oitavas | Quartas | Semi | Final | 3º Lugar | Total |
|------|-------------|---------|---------|------|-------|----------|-------|
| 2010 | 48 | 8 | 4 | 2 | 1 | 1 | 64 |
| 2014 | 48 | 8 | 4 | 2 | 1 | 1 | 64 |
| 2018 | 48 | 8 | 4 | 2 | 1 | 1 | 64 |
| 2022 | 48 | 8 | 4 | 2 | 1 | 1 | 64 |
| 2026 | 72 | 16 | 8 | 4 | 2 | 1 | 103* |

*2026: 12 grupos × 4 = 48 times → 24 classificados + 8 melhores 3º = 32 → mata-mata tradicional

### 6. fact_events (4.846 linhas)
| Coluna | Tipo | PK/FK | Descrição |
|--------|------|-------|-----------|
| EventID | Integer | ✓ | ID único do evento |
| MatchID | Integer | FK | Partida (→ fact_matches) |
| PlayerID | Integer | FK | Jogador (→ dim_players) |
| TeamCode | String | | Time do jogador | "BRA", "ARG", etc. |
| EventType | String | | Tipo do evento | "Goal", "YellowCard", "RedCard", "Substitution" |
| Minute | Integer | | Minuto do evento | 1-120 |
| Detail | String | | Detalhe do evento | Ver abaixo |
| xG | Decimal | | Expected Goals (apenas gols) | 0.00-1.00 |

**EventType + Detail:**
| EventType | Detail | Descrição |
|-----------|--------|-----------|
| Goal | "Open Play" | Gol em jogo normal |
| Goal | "Penalty" | Gol de pênalti |
| Goal | "Own Goal" | Gol contra |
| Goal | "Free Kick" | Gol de falta |
| Goal | "Header" | Gol de cabeça |
| YellowCard | "Foul" | Cartão amarelo por falta |
| YellowCard | "Dissent" | Cartão amarelo por reclamação |
| YellowCard | "Time Wasting" | Cartão amarelo por cera |
| RedCard | "Serious Foul" | Cartão vermelho por falta grave |
| RedCard | "Violent Conduct" | Cartão vermelho por conduta violenta |
| RedCard | "DOGSO" | Cartão vermelho por negar gol claro |
| Substitution | "Sub On" | Jogador entra |
| Substitution | "Sub Off" | Jogador sai |

---

## Relacionamentos (Star Schema)

```
dim_world_cups[Year] ────── 1:N ────── fact_matches[WorldCupYear]
dim_world_cups[Year] ────── 1:N ────── dim_players[WorldCupYear]

dim_countries[CountryCode] ── 1:N ──── dim_teams[TeamCode]

dim_teams[TeamCode, WorldCupYear] ── 1:N ── fact_matches[HomeTeam, WorldCupYear]
dim_teams[TeamCode, WorldCupYear] ── 1:N ── fact_matches[AwayTeam, WorldCupYear]

dim_players[PlayerID] ────── 1:N ────── fact_events[PlayerID]

fact_matches[MatchID] ────── 1:N ────── fact_events[MatchID]
```

**Cardinalidades:**
- Todas as relações são 1:N (um para muitos)
- Cross-filter direction: Single (da dimensão para o fato)
- Relacionamentos ativos: Todos

---

## Medidas Calculadas (DAX)

### KPIs Principais
```dax
Total Goals = SUMX(FILTER(fact_events, fact_events[EventType]="Goal"), fact_events[xG])
Matches Played = DISTINCTCOUNT(fact_matches[MatchID])
Avg Goals Per Match = DIVIDE([Total Goals], [Matches Played])
```

### Por Jogador
```dax
Player Goals = CALCULATE(SUMX(FILTER(fact_events, fact_events[EventType]="Goal"), fact_events[xG]), ALLEXCEPT(dim_players, dim_players[PlayerID]))
Player Assists = CALCULATE(SUM(dim_players[Assists]), ALLEXCEPT(dim_players, dim_players[PlayerID]))
Player xG = CALCULATE(SUM(dim_players[xG]), ALLEXCEPT(dim_players, dim_players[PlayerID]))
Player Minutes = CALCULATE(SUM(dim_players[Minutes]), ALLEXCEPT(dim_players, dim_players[PlayerID]))
Goals Per 90 = DIVIDE([Player Goals] * 90, [Player Minutes])
```

### Por Seleção
```dax
Team Goals = CALCULATE(SUMX(FILTER(fact_events, fact_events[EventType]="Goal" && fact_events[TeamCode]=SELECTEDVALUE(dim_teams[TeamCode])), fact_events[xG]))
Team Market Value = SUMX(FILTER(dim_players, dim_players[CountryCode]=SELECTEDVALUE(dim_teams[TeamCode]) && dim_players[WorldCupYear]=SELECTEDVALUE(dim_teams[WorldCupYear])), dim_players[MarketValue])
Team Avg Rating = AVERAGEX(FILTER(dim_players, dim_players[CountryCode]=SELECTEDVALUE(dim_teams[TeamCode]) && dim_players[WorldCupYear]=SELECTEDVALUE(dim_teams[WorldCupYear])), dim_players[OverallRating])
```

### Time Intelligence
```dax
Market Value YoY = 
VAR CurrYear = SELECTEDVALUE(dim_world_cups[Year])
VAR PrevYear = CurrYear - 4
RETURN DIVIDE(
    CALCULATE([Total Market Value], dim_world_cups[Year]=CurrYear),
    CALCULATE([Total Market Value], dim_world_cups[Year]=PrevYear)
) - 1
```

### Avançadas
```dax
Top Scorer = MAXX(TOPN(1, SUMMARIZE(dim_players, dim_players[Name], dim_players[CountryCode], "Goals", CALCULATE(SUMX(FILTER(fact_events, fact_events[EventType]="Goal"), fact_events[xG]), ALLEXCEPT(dim_players, dim_players[PlayerID]))), [Goals], DESC), dim_players[Name] & " (" & dim_players[CountryCode] & ") - " & FORMAT([Goals], "0.0"))

Best GK = MAXX(TOPN(1, FILTER(SUMMARIZE(dim_players, dim_players[Name], dim_players[CountryCode], "Saves", CALCULATE(SUM(fact_events[Saves]), fact_events[EventType]="Save", ALLEXCEPT(dim_players, dim_players[PlayerID]))), dim_players[Position]="GK"), [Saves], DESC), dim_players[Name] & " (" & dim_players[CountryCode] & ") - " & FORMAT([Saves], "0") & " defesas")
```

---

## Validações de Qualidade

### Contagens Esperadas
| Tabela | Linhas Esperadas | Verificação |
|--------|------------------|-------------|
| dim_world_cups | 5 | Anos 2010, 2014, 2018, 2022, 2026 |
| dim_countries | 63 | Códigos FIFA únicos |
| dim_teams | 176 | 32×4 + 48 |
| dim_players | 4.048 | 23 × 176 |
| fact_matches | 344 | 64×4 + 104 |
| fact_events | ~4.846 | ~14 eventos/partida |

### Integridade Referencial
- [ ] Todos PlayerID em fact_events existem em dim_players
- [ ] Todos MatchID em fact_events existem em fact_matches
- [ ] Todos CountryCode em dim_players existem em dim_countries
- [ ] Todos TeamCode em fact_matches existem em dim_teams (para WorldCupYear correspondente)
- [ ] Todos WorldCupYear em dim_players existem em dim_world_cups

### Validações de Negócio
- [ ] Cada seleção por Copa tem exatamente 23 jogadores
- [ ] Posições por seleção: 3 GK, 4 CB, 2 RB, 2 LB, 2 CDM, 3 CM, 1 CAM, 2 RW, 2 LW, 2 ST
- [ ] Gols em fact_events = HomeGoals + AwayGoals em fact_matches (por partida)
- [ ] Minutos ≤ Partidas × 90 (aprox)
- [ ] Idade entre 16-45
- [ ] MarketValue > 0
- [ ] OverallRating entre 60-99

---

## Arquivos de Origem
```
data/
├── dim_world_cups.csv          # 5 linhas
├── dim_countries.csv           # 63 linhas
├── dim_teams.csv               # 176 linhas
├── dim_players.csv             # 4.048 linhas
├── fact_matches.csv            # 344 linhas
└── fact_events.csv             # 4.846 linhas
```

---

**Gerado em:** 2026  
**Versão:** 1.0  
**Responsável:** Investigation Team AI (Rise Kujikawa)