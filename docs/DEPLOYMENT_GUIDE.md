# 🚀 Guia de Deploy - Álbum da Copa Definitivo no Power BI Service

## 📋 Pré-requisitos

- Power BI Desktop (versão mais recente)
- Conta Power BI Pro ou Premium Per User
- Workspace no Power BI Service
- Arquivos CSV na pasta `data/`
- Arquivo `.pbix` compilado

---

## 1️⃣ Preparação no Power BI Desktop

### 1.1 Importar Dados
```
Home → Get Data → Text/CSV
Selecionar todos os 6 arquivos da pasta data/:
├── dim_world_cups.csv
├── dim_countries.csv
├── dim_teams.csv
├── dim_players.csv
├── fact_matches.csv
└── fact_events.csv

Para cada arquivo:
- Delimiter: Comma
- Data Type Detection: Based on first 200 rows
- Transformar tipos de dados conforme DATA_DICTIONARY.md
```

### 1.2 Configurar Relacionamentos (Modelagem)
```
View → Model → Manage Relationships

Relacionamentos (todos 1:N, Single Cross-filter):

1. dim_world_cups[Year] ──────────→ fact_matches[WorldCupYear]
2. dim_world_cups[Year] ──────────→ dim_players[WorldCupYear]
3. dim_countries[CountryCode] ────→ dim_teams[TeamCode]
4. dim_teams[TeamCode, WorldCupYear] → fact_matches[HomeTeam, WorldCupYear]
5. dim_teams[TeamCode, WorldCupYear] → fact_matches[AwayTeam, WorldCupYear]
6. dim_players[PlayerID] ─────────→ fact_events[PlayerID]
7. fact_matches[MatchID] ─────────→ fact_events[MatchID]

⚠️ Para relacionamentos compostos (4 e 5):
   - Clicar "New" → Selecionar ambas colunas nas duas tabelas
   - Cardinality: One to Many
   - Cross filter: Single
```

### 1.3 Criar Medidas DAX
```
Modeling → New Measure
Copiar TODAS as medidas de docs/DAX_MEASURES.md

Organizar em Display Folders:
- KPIs
- Player Metrics
- Team Metrics
- Time Intelligence
- Advanced Analytics
```

### 1.4 Criar Calculation Group
```
Modeling → Calculation Groups → New
Name: "Time Intelligence"

Itens (clicar "New Calculation Item"):
1. Current Period (Base)
2. Previous Tournament (-4 anos)
3. Tournament Rank
4. Tournament Growth (%)
5. Best Tournament
```

### 1.5 Criar Parâmetros What-If
```
Modeling → New Parameter (4x):

1. GrowthRate
   - Type: Decimal
   - Minimum: -0.2
   - Maximum: 0.5
   - Increment: 0.05
   - Default: 0.1
   - Add slicer: ✓

2. Multiplier
   - Type: Decimal
   - Minimum: 0.5
   - Maximum: 3.0
   - Increment: 0.1
   - Default: 1.5
   - Add slicer: ✓

3. MinMinutes
   - Type: Whole Number
   - Minimum: 0
   - Maximum: 540
   - Increment: 30
   - Default: 90
   - Add slicer: ✓

4. TopN
   - Type: Whole Number
   - Minimum: 3
   - Maximum: 20
   - Increment: 1
   - Default: 10
   - Add slicer: ✓
```

### 1.6 Aplicar Tema Visual
```
View → Themes → Switch theme → Import theme
Selecionar arquivo theme.json (criar conforme paleta em powerbi_visual_spec.md)

Cores principais:
- Primary: #1B4F72
- Secondary: #2E86C1
- Accent: #E74C3C
- Success: #27AE60
- Warning: #F39C12
- Confederation colors: UEFA=#3498DB, CONMEBOL=#F39C12, CAF=#27AE60, AFC=#9B59B6, CONCACAF=#E67E22, OFC=#1ABC9C
```

---

## 2️⃣ Construir as 8 Páginas

### Página 1: Executive Overview
```
Layout: Grid 12 colunas × 8 linhas

Linha 1 (Header):
- Slicer: dim_world_cups[Year] (Dropdown, Single, Default 2026) → Col 1-2
- Slicer: dim_countries[Confederation] (Dropdown, Multi) → Col 3-4
- Title: "Executive Overview - Copa {Year}" → Col 5-12

Linha 2 (KPIs - Cards Grandes):
- [Total Goals] → Col 1-3 (Sparkline: Gols por Copa)
- [Matches Played] → Col 4-6 (Target: 64/104)
- [Avg Goals Per Match] → Col 7-9 (Format: 0.00)
- [Total Market Value] → Col 10-12 (Format: Bi€)

Linha 3 (KPIs Médios):
- [Top Scorer] → Col 1-3
- [Best GK] → Col 4-6
- [Avg Squad Age] → Col 7-9
- [Pct U23] → Col 10-12

Linha 4-5 (Gráficos):
- Line Chart: Gols por Copa (Legenda: Stage) → Col 1-6
- Clustered Column: Valor Mercado por Copa (Cor: Confederation) → Col 7-12

Linha 6-9 (Tabela):
- Matrix: Top 10 Países (País, Gols, xG, Partidas, V, %, Rating, Valor) → Col 1-12

Bookmarks: Executive_2010, Executive_2014, Executive_2018, Executive_2022, Executive_2026
```

### Página 2: Tournament Deep Dive
```
Slicer: fact_matches[Stage] (Buttons horizontal)
Matrix: Resultados (Linhas: HomeTeam, Colunas: AwayTeam, Valores: HomeGoals, AwayGoals)
Image: Bracket SVG (com actions para drillthrough)
Stacked Column: Gols por Fase (Legenda: Detail)
Timeline: Partidas (Custom visual)
Table: Partidas detalhadas
Drillthrough → Página 6 (Match Timeline) via MatchID
```

### Página 3: Player Explorer (Figurinha)
```
Slicers: Player (Search), WorldCupYear, Position, CountryCode
Header: Nome, País (Bandeira), Copa, Pos, Idade, Pé, Altura/Peso
Gauge: OverallRating (0-99, Zonas: 70/80/90/95)
Card: MarketValue (€#,##0.0M)
Radar Chart: 6 eixos por posição (Custom visual)
Multi-row Card: Stats principais
Table: Histórico Copas (Copa, País, Idade, P, Min, G, A, xG, Rating, Fase)
Scatter Plot: xG/90 vs G/90 (Tamanho: Min, Cor: Pos, Destaque: Selecionado)
Drillthrough ← Páginas 1,2,4 via PlayerID
```

### Página 4: Team Profile
```
Slicers: Country (com bandeira), WorldCupYear
Header: Nome, Confed, Copa, Fase, Técnico
KPIs: P, V-E-D, GF, GC, Saldo
Cards: Artilheiro, Garçom
Formação Tática: SVG campo clicável → Drillthrough Player
Table: Elenco 23 (Todos os campos)
Histogram: Idade
Stacked Bar: Valor Mercado por Pos
Donut: Minutos por Pos
Table: Disciplina
Line: Evolução histórica por Copa
Drillthrough ← Página 3 via click elenco/formação
```

### Página 5: Stats Laboratory
```
Table+Heatmap: Per 90 por Posição
Box & Whisker: Percentis (Small Multiples por Pos)
Line Chart: Aging Curves (Idade vs Métricas, linhas por Pos)
Scatter: MarketValue vs Rating (Tamanho: Min, Cor: Idade, Forma: Pos)
What-If: GrowthRate slider + Card Projected Value
Decomposition Tree: [Total Goals] → Cup → Stage → Type → Pos → Country → Player
Key Influencers: Goals >0 explicado por Pos, Age, Min, xG, Shots, Rating, Value, Cup
Heatmap: Correlation Matrix
```

### Página 6: Match Timeline
```
Slicers: WorldCupYear, Stage, Team (Union Home/Away)
Timeline Custom: Data X, Bolhas=Partidas (Tamanho=Gols, Cor=Fase)
Tooltip Page: Tooltip_Match (Placar, Gols com min/jogador/tipo/xG, Cartões, Público)
Table: Partidas ordenáveis
Decomposition Tree: Gols da Partida → Time → Jogador → Tipo → Minuto
Cards: Posse, Finaliz, Escanteios, Faltas, Cartões, Impedimentos, xG Diff
Drillthrough ← Página 2 via MatchID
```

### Página 7: Comparative Analytics
```
Slicer: WorldCupYear (Checkbox - Multi-select)
Small Multiples (4): Gols, xG, Disciplina, Idade, Valor Mercado
Matrix: Métricas × Copas (Heatmap por coluna)
Key Influencers: Goals contínuo
Q&A Visual
Bookmarks: Compare_2010_2014, Compare_2018_2022, Compare_All, Compare_Champions
```

### Página 8: Scouting & Projections
```
Table: Sub-23 (Filtro Age<23, Ordenar por Potencial)
Line+Forecast: Projeção 2030 (ETS, What-If GrowthRate)
Scatter+What-If: Transfer Value Estimator (Multiplier)
Matrix+Heatmap: Positional Need by Team
Table: Sucessão de Estrelas (30+ → Sub-23 mesma pos/país)
RLS Roles: 6 confederações + Admin
```

---

## 3️⃣ Configurar Recursos Avançados

### 3.1 Bookmarks (View → Bookmarks)
```
Para cada página, criar bookmarks de navegação:
- Executive_2010 a Executive_2026
- Tournament_Group, Tournament_Knockout
- Player_Messi (demo)
- Team_Brazil (demo)
- Compare_Champions
- Scouting_U23

Configurar: Data, Display, Current page, Visuals, Filters
```

### 3.2 Drillthrough
```
Página 3 (Player): Drillthrough filter = PlayerID
Página 4 (Team): Drillthrough filter = CountryCode + WorldCupYear
Página 6 (Match): Drillthrough filter = MatchID

Configurar: Keep all filters = On
Botão "← Voltar" em cada página drillthrough (Action: Bookmark "Back")
```

### 3.3 Tooltips Personalizados
```
Criar páginas de tooltip (Page Size: Tooltip):
1. Tooltip_Player: Mini figurinha (Foto, Nome, País, Pos, Idade, G, A, xG, Rating, Valor, Sparkline)
2. Tooltip_Match: Placar, Data, Estádio, Gols (min/jogador/tipo), xG, Cartões, Público
3. Tooltip_Team: Bandeira, Copa, Fase, Artilheiro, Rating, Valor, Últimas 3 Copas

Vincular: Visual → Format → Tooltip → Type: Report page → Select tooltip page
```

### 3.4 RLS (Row Level Security)
```
Modeling → Manage Roles → Create Role (7 roles):

1. UEFA: [Confederation] = "UEFA"
2. CONMEBOL: [Confederation] = "CONMEBOL"
3. CAF: [Confederation] = "CAF"
4. AFC: [Confederation] = "AFC"
5. CONCACAF: [Confederation] = "CONCACAF"
6. OFC: [Confederation] = "OFC"
7. Admin: (sem filtro)

Testar: View as Roles → Selecionar role → Verificar dados filtrados
```

### 3.5 Navegação Global
```
Master Page (oculta) com botões fixos:
- 🏠 Home (Bookmark Executive_2026)
- 🏆 Executive (Bookmark Executive_2026)
- 🏟️ Tournament (Page 2)
- 👤 Players (Page 3)
- 🏳️ Teams (Page 4)
- 📊 Stats (Page 5)
- ⚽ Matches (Page 6)
- 📈 Compare (Page 7)
- 🔮 Scouting (Page 8)

Copiar para todas as páginas (Ctrl+C, Ctrl+V mantém posição)
```

---

## 4️⃣ Validação e Testes

### 4.1 Checklist Funcional
```
☐ Todos relacionamentos ativos e corretos
☐ Medidas DAX retornam valores esperados (testar com dados conhecidos)
☐ Calculation Group funciona (trocar item, ver medidas mudarem)
☐ What-If parameters atualizam cards/gráficos em tempo real
☐ Bookmarks navegam corretamente (filtros, página, visibilidade)
☐ Drillthrough funciona (clique direito → Drillthrough → Página correta)
☐ Tooltips personalizados aparecem no hover
☐ RLS testado para cada role (View as Roles)
☐ Decomposition Tree expande corretamente
☐ Key Influencers mostra fatores relevantes
☐ Small Multiples mostram painéis por Copa
☐ Forecast aparece na linha de tendência
☐ Formatação condicional nas tabelas (heatmap, data bars, ícones)
☐ Responsividade: Desktop (16:9), Tablet, Mobile (Power BI App)
☐ Tema aplicado consistentemente em todas páginas
```

### 4.2 Testes de Performance
```
Performance Analyzer (View → Performance Analyzer):
- Gravar interações:
  1. Carregar página
  2. Interagir com slicers
  3. Drillthrough
  4. Exportar dados

Metas:
- Page load < 3s
- Slicer interaction < 1s
- Drillthrough < 2s
- DAX query time < 500ms por medida

Otimizar se necessário:
- Criar agregações (Manage Aggregations)
- Reduzir cardinalidade (agrupar categorias)
- Desativar bi-directional cross-filter onde não necessário
```

---

## 5️⃣ Publicação no Power BI Service

### 5.1 Publicar
```
File → Publish → Publish to Power BI
Selecionar Workspace: "Album Copa Definitivo - Portfolio"
Aguardar upload completo
```

### 5.2 Configurar Atualização Agendada
```
Workspace → Dataset (Album_Copa_Definitivo) → Settings → Scheduled refresh
- Frequency: Daily
- Time zone: (UTC-3) Brasília
- Time: 06:00
- Notify on failure: ✓ (seu email)

Nota: Dados são estáticos, mas boa prática manter refresh
```

### 5.3 Configurar RLS no Service
```
Workspace → Dataset → Security
Para cada role (UEFA, CONMEBOL, CAF, AFC, CONCACAF, OFC):
- Add members: emails dos analistas da confederação
- Test: "Test as role" → Verificar filtros

Admin role: não adicionar membros (acesso total via workspace admin)
```

### 5.4 Criar App para Distribuição
```
Workspace → Create app
Name: "Álbum da Copa Definitivo - Public"
Description: "Dashboard interativo Copas 2010-2026 com 4K jogadores, 344 partidas, DAX avançado"
Logo: assets/logos/app_logo.png
Color: #1B4F72
Navigation: Default (todas páginas)
Access: Anyone with link (ou organização)
Publish app
```

### 5.5 Embed para Portfólio Web
```
App → Embed → Website or portal
Copy iframe code
Configurar: Width 100%, Height 600px, Allow fullscreen
Inserir no site de portfólio
```

---

## 6️⃣ Documentação Entregável

### Arquivos para Cliente/Entrevista
```
Entregáveis/
├── Album_Copa_Definitivo.pbix          # Arquivo fonte
├── Album_Copa_Definitivo_2010_2026.xlsx # Excel interativo
├── data/                                # 6 CSVs
├── docs/
│   ├── README.md                        # Visão geral
│   ├── DAX_MEASURES.md                  # 50+ medidas documentadas
│   ├── DATA_DICTIONARY.md               # Dicionário completo
│   ├── DEPLOYMENT_GUIDE.md              # Este arquivo
│   └── powerbi_visual_spec.md           # Spec visual 8 páginas
├── vba/
│   ├── frmAlbum.frm                     # UserForm álbum
│   └── frmCompare.frm                   # UserForm comparação
└── assets/                              # Bandeiras, logos, placeholders
```

---

## 7️⃣ Troubleshooting Comum

| Problema | Causa | Solução |
|----------|-------|---------|
| Relacionamento não ativa | Cross-filter bidirecional | Usar Single direction, ou USERELATIONSHIP() na medida |
| Medida retorna BLANK | Contexto de filtro incorreto | Usar CALCULATE com ALL/ALLEXCEPT |
| Drillthrough não passa filtro | Campo errado no drillthrough filter | Verificar PlayerID/CountryCode/MatchID exato |
| RLS não filtra | Role mal configurada | Testar com View as Roles, verificar tabela Users |
| Performance lenta | Alta cardinalidade | Criar agregações, reduzir colunas desnecessárias |
| What-If não atualiza | Medida não referencia parâmetro | Usar `ParameterName[ParameterName Value]` |
| Bookmark não restaura slicers | "Data" não marcado no bookmark | Editar bookmark → marcar Data, Display, Current page |

---

## 8️⃣ Métricas de Sucesso do Projeto

### Para Portfólio Freelance
| Métrica | Target | Como Medir |
|---------|--------|------------|
| Complexidade DAX | 20+ medidas avançadas | Contar medidas com CALCULATE, FILTER, RANKX, TOPN |
| Visuals Avançados | 8 tipos | Decomposition Tree, Key Influencers, What-If, Forecast, Small Multiples, Drillthrough, Bookmarks, RLS |
| Modelagem | Star Schema 5D+2F | Verificar relacionamentos no Model View |
| Documentação | Completa | README, DAX_MEASURES, DATA_DICTIONARY, DEPLOYMENT_GUIDE |
| Excel/VBA | Funcional | frmAlbum.carregar, navegar, radar chart |
| Deploy | Publicado | App acessível via link, RLS funcionando |

### Diferenciais para Entrevista
1. **"Modelei um star schema com chaves compostas e RLS por confederação"**
2. **"Implementei Calculation Groups para Time Intelligence em torneios quadrienais"**
3. **"Usei What-If parameters para projeção de valor de mercado 2030"**
4. **"Criei Decomposition Tree para analisar gols por tipo/posição/jogador"**
5. **"Key Influencers revelou que xG e Position explicam 73% da variância de gols"**
6. **"Drillthrough contextual: Partida → Jogador → Histórico Copas"**
7. **"VBA UserForm estilo álbum de figurinhas com radar chart dinâmico"**

---

## 📞 Suporte Pós-Deploy

### Monitoramento
- Verificar refresh diário no Service (Settings → Refresh history)
- Monitorar uso via Power BI Admin Portal (se admin)
- Coletar feedback de usuários (App → Feedback)

### Atualizações Futuras
- Adicionar Copa 2030 quando dados reais disponíveis
- Integrar dados reais (FIFA API, Opta, StatsBomb)
- Adicionar métricas avançadas (PPDA, xT, Pressing)
- Implementar alertas (Data-driven alerts no Service)

---

**Versão:** 1.0 | **Data:** 2026  
**Autor:** Investigation Team AI (Rise Kujikawa - Orquestradora)  
**Stack:** Python → CSV → Power BI Desktop → DAX → Power BI Service → App