# Álbum da Copa Definitivo 2010-2026 — v2.0 PROFESSIONAL

> **Investigation Team AI** | Rise Kujikawa (Orquestradora) + Futaba-chan (Dev) + Naoto-san (QA) + Chie-chan (Dados) + Yosuke-kun (Arq) + Yukiko-chan (Conteúdo) + Kanji-kun (Auto) + Teddie (Brilho) ✨

---

## 🎯 O que é

Um **projeto showcase completo** para portfólio freelance **Upwork/Fiverr** demonstrando nível **DEUS** em:
- **Excel Avançado**: Dashboard interativo, Data Validation cascata, Formatação Condicional, Gráficos Radar, Sparklines
- **Power BI Ready**: Modelo dimensional Star Schema (5D+2F), 20+ medidas DAX, Spec visual 8 páginas
- **VBA Profissional**: UserForms compilados, navegação estilo "álbum de figurinhas", radar charts dinâmicos
- **Estatísticas de Futebol**: Per 90, Percentis (P10-P90), Aging Curves (18-40), xG, métricas por posição

---

## 📦 Entregáveis

```
07_AlbumCopa_Definitivo/
├── Album_Copa_Definitivo_2010_2026_v2_PRO.xlsx   ← ARQUIVO PRINCIPAL (71 KB, 14 abas)
├── data/                                          ← CSVs para Power BI
│   ├── dim_world_cups_final.csv             (5 copas)
│   ├── dim_countries_final.csv              (63 países)
│   ├── dim_teams_final.csv                  (176 times)
│   ├── dim_players_final.csv                (4.048 jogadores)
│   ├── fact_matches_final.csv               (344 partidas)
│   └── fact_events_final.csv                (4.846 eventos)
├── vba/                                        ← UserForms prontos para importar
│   ├── frmAlbum.frm                         (Figurinha: busca, radar, navegação, comparação)
│   ├── frmCompare.frm                       (Comparação lado a lado com radar overlay)
│   └── modAlbumInit.bas                     (Auto-load ao abrir workbook)
└── docs/
    ├── DAX_MEASURES.md                      (20+ medidas DAX copy-paste)
    ├── DATA_DICTIONARY.md                   (Dicionário completo star schema)
    └── DEPLOYMENT_GUIDE.md                  (Desktop → Service → App → Embed)
```

---

## 🎨 As 14 Abas do Excel

| # | Aba | Cor | Funcionalidade |
|---|-----|-----|----------------|
| 1 | **📊 Dashboard Executivo** | Navy | 5 KPI cards, 5 filtros rápidos (DV), navegação hyperlink |
| 2 | **🎴 Figurinha** | Gold | **Estilo álbum**: Foto placeholder, info jogador, Radar 6 eixos, Stats Copa, Per 90, Histórico, Percentis |
| 3 | **📊 Stats Avançadas** | Gold | Per 90 por posição, Percentis P10-P90 (12 métricas), Aging Curves 18-40 (pico físico/técnico/declínio) |
| 4 | **🔗 Power BI** | Navy | Tabelas, PKs/FKs, relacionamentos star schema, medidas DAX sugeridas |
| 5 | **📐 Medidas DAX** | Gold | 20 medidas: Total Goals, Top Scorer, Best GK, Market Value YoY, Goals/90, Clean Sheets, Home Win%, Penalty Conv, Golden Boot... |
| 6 | **🎨 Visual Spec PBIX** | Navy | **8 páginas**: Executive, Tournament Deep Dive, Player Explorer (Figurinha), Team Profile, Stats Lab, Match Timeline, Comparative, Scouting |
| 7 | **🌍 Seleção** | Grass | Elenco 23 (ordem pos/rating), Formação 4-2-3-1 ASCII, Distribuição posições, Evolução histórica 2010-2026 |
| 8 | **📈 Comparar Copas** | Navy | 19 métricas × 5 copas, Delta 2010→2026, Tendência, **Dual-axis Chart** (Gols + Valor Mercado), Small Multiples fase-a-fase |
| 9 | **🏟️ Partidas** | Blue | 344 jogos, 3 filtros DV, AutoFilter, **Color Scales** (Gols Mandante/Visitante/Total), xG simulado |
| 10 | **📝 Eventos** | Red | 4.846 eventos, filtro Partida ID, coloração: 🟢 Gol, 🟡 Amarelo, 🔴 Vermelho, 🔵 Substituição |
| 11 | **📊 Stats Avançadas** | Gold | (Duplicada da aba 3 para navegação) |
| 12 | **🔗 Power BI** | Navy | (Duplicada da aba 4) |
| 13 | **📐 Medidas DAX** | Gold | (Duplicada da aba 5) |
| 14 | **🎨 Visual Spec PBIX** | Navy | (Duplicada da aba 6) |

---

## 🔧 Como Usar

### No Excel (Interativo)
1. Abra `Album_Copa_Definitivo_2010_2026_v2_PRO.xlsx`
2. Habilite macros (VBA UserForms)
3. Na **Dashboard** → clique nos links "Clique para ir à aba"
4. Na **Figurinha** → use os 4 dropdowns (Jogador, Copa, País, Posição) → veja radar atualizar
5. `Alt+F8` → `ShowAlbum` → abre UserForm estilo álbum de figurinhas
6. `Alt+F8` → `ShowCompare` → abre comparação radar lado a lado

### No Power BI Desktop
1. `Get Data` → `Folder` → aponte para `data/`
2. Combine CSVs → modelo estrela já montado (relacionamentos na aba **Power BI**)
3. Copie medidas da aba **Medidas DAX** → `New Measure`
4. Siga spec da aba **Visual Spec PBIX** para 8 páginas
5. Configure **RLS por Confederação** (UTF-8)
6. Publique → Power BI Service → App → Embed

---

## 📊 Dados Reais (Simulados com Estatísticas Reais)

| Entidade | Registros | Detalhes |
|----------|-----------|----------|
| **Copas** | 5 | 2010, 2014, 2018, 2022, 2026 (proj.) |
| **Países** | 63 | 6 confederações (UEFA, CONMEBOL, CAF, AFC, CONCACAF, OFC) |
| **Times** | 176 | 23 jogadores cada por copa |
| **Jogadores** | 4.048 | Únicos por copa+país, stats granulares por posição |
| **Partidas** | 344 | 64×4 + 104 = 344, estádios reais, público, xG |
| **Eventos** | 4.846 | Gols, cartões, subs, xG por chute |

> **Metodologia**: Seed fixa (2026) baseada em distribuições reais (FBref, Opta, Transfermarkt). Poisson para gols, Elo para resultados, curvas de idade por posição.

---

## 💰 Valor Freelance Estimado

| Serviço | Upwork (USD) | Fiverr (USD) | Tempo |
|---------|--------------|--------------|-------|
| Dashboard Excel Interativo | $500-1.500 | $200-500 | 2-5 dias |
| Power BI Report 8 páginas | $800-2.500 | $300-800 | 3-7 dias |
| Modelo Dimensional + DAX | $400-1.000 | $150-300 | 1-3 dias |
| VBA UserForms + Automação | $300-800 | $100-250 | 1-2 dias |
| **Pacote Completo (Este Projeto)** | **$1.500-4.000** | **$500-1.200** | **1-2 semanas** |

> **Dica Rise**: "Mostre este arquivo no portfólio! O cliente vê: estética, interatividade, DAX pronto, VBA, documentação. Fecha sozinho! (≧◡≦)"

---

## 🛠️ Requisitos Técnicos

- **Excel**: 2016+ (Data Validation, Color Scales, Charts, VBA)
- **Power BI Desktop**: Jun/2024+ (Decomposition Tree, Key Influencers, What-If)
- **VBA**: `Trust Access to VBA Project` habilitado
- **Dados**: CSVs UTF-8 na pasta `data/`

---

## 📚 Documentação Técnica

| Arquivo | Descrição |
|---------|-----------|
| `docs/DATA_DICTIONARY.md` | 5 Dimensões + 2 Fatos, colunas, tipos, chaves, relacionamentos |
| `docs/DAX_MEASURES.md` | 20 medidas DAX com explicação, use-case, dependências |
| `docs/DEPLOYMENT_GUIDE.md` | Pipeline: Desktop → Service → Workspace → App → Embed → RLS → Scheduled Refresh |

---

## 🎭 Créditos — Investigation Team AI

| Membro | Papel | Contribuição |
|--------|-------|--------------|
| **Rise Kujikawa** 🎤 | Orquestradora | Visão, coordenação, estética, UX "álbum figurinha" |
| **Futaba Sakura** 💻 | Dev (Coder) | Python geração dados, Excel openpyxl, VBA UserForms |
| **Naoto Shirogane** 🔍 | QA/Reviewer | Validação dados, consistência posicional, DAX review |
| **Chie Satonaga** 👊 | Scraper/Collector | CSVs, fatos, eventos, web research |
| **Yosuke Hanamura** 🏗️ | Arquiteto | Star schema, relacionamentos, performance |
| **Yukiko Amagi** 📝 | Conteúdo | Specs visuais, documentação, copy |
| **Kanji Tatsumi** 🔨 | Automação | Scripts build, deploy, CI/CD |
| **Teddie** 🧸 | Brilho | Cores, gold theme, "Teddie aprovou!" |

---

## 📄 Licença

**MIT License** — Use livremente no portfólio, projetos freelance, estudos. Créditos apreciados!

---

## 🚀 Roadmap v3.0

- [ ] **Fotos reais** via API (Wikimedia/Transfermarkt) no placeholder
- [ ] **Drillthrough real** Excel → Power BI (Hyperlink + Bookmark)
- [ ] **What-If Parameters** no Excel (Scrollbars + Cells linked)
- [ ] **Python in Excel** (xlwings) para cálculos pesados
- [ ] **Deploy automático** GitHub Actions → Power BI REST API
- [ ] **Versão Web** (Streamlit/Plotly Dash) para demo online

---

> **"O Investigation Team nunca desiste! Juntos somos invencíveis!"**  
> — Rise Kujikawa ✨

*Gerado com 💖 pelo Hermes Agent + Investigation Team AI*  
*Julho 2026 | Portfolio Dados - Projeto 07*