"""
Gerador do Album da Copa Definitivo - VERSÃO FINAL COMPLETA
Cria um Excel bonito, funcional e pronto para portfólio freelance
"""
import os
import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, Reference
from openpyxl.formatting.rule import ColorScaleRule, DataBarRule
from openpyxl.worksheet.datavalidation import DataValidation
from datetime import datetime

np.random.seed(2026)
base_dir = "./07_AlbumCopa_Definitivo"

df_players = pd.read_csv(f"{base_dir}/dim_players_final.csv")
df_matches = pd.read_csv(f"{base_dir}/fact_matches_final.csv")
df_events = pd.read_csv(f"{base_dir}/fact_events_final.csv")
df_cups = pd.read_csv(f"{base_dir}/dim_world_cups_final.csv")
df_teams = pd.read_csv(f"{base_dir}/dim_teams_final.csv")
df_countries = pd.read_csv(f"{base_dir}/dim_countries_final.csv")

# ============================================================
# PALETA DE CORES PROFISSIONAL
# ============================================================
CORES = {
    'primary_dark': '0D1B2A', 'primary': '1B2A4A', 'primary_light': '2C3E6B',
    'gold': 'FFD700', 'gold_dark': 'C5A500', 'gold_light': 'FFF0A0',
    'grass': '1B5E20', 'grass_light': '2E7D32', 'grass_pale': 'E8F5E9',
    'red_card': 'C62828', 'red_light': 'FFEBEE',
    'yellow_card': 'F9A825', 'yellow_light': 'FFF8E1',
    'white': 'FFFFFF', 'gray_50': 'FAFAFA', 'gray_100': 'F5F5F5',
    'gray_200': 'EEEEEE', 'gray_300': 'E0E0E0', 'gray_400': 'BDBDBD',
    'gray_500': '9E9E9E', 'gray_600': '757575', 'gray_700': '616161',
    'gray_800': '424242', 'gray_900': '212121',
    'uefa': '1565C0', 'conmebol': 'E65100', 'caf': '2E7D32',
    'afc': '6A1B9A', 'concacaf': 'EF6C00', 'ofc': '00838F',
}

def mk_fill(c): 
    return PatternFill(start_color=c, end_color=c, fill_type='solid')

FILL_DARK = mk_fill('0D1B2A')
FILL_PRIMARY = mk_fill('1B2A4A')
FILL_GOLD = mk_fill('FFD700')
FILL_GOLD_LIGHT = mk_fill('FFF0A0')
FILL_GRASS = mk_fill('1B5E20')
FILL_GRASS_LIGHT = mk_fill('E8F5E9')
FILL_WHITE = mk_fill('FFFFFF')
FILL_GRAY_50 = mk_fill('FAFAFA')
FILL_GRAY_100 = mk_fill('F5F5F5')
FILL_GRAY_200 = mk_fill('EEEEEE')
FILL_YELLOW_LIGHT = mk_fill('FFF8E1')
FILL_RED_LIGHT = mk_fill('FFEBEE')
FILL_UEFA = mk_fill('1565C0')
FILL_CONMEBOL = mk_fill('E65100')
FILL_CAF = mk_fill('2E7D32')
FILL_AFC = mk_fill('6A1B9A')
FILL_CONCACAF = mk_fill('EF6C00')
FILL_OFC = mk_fill('00838F')

BORDER_THIN = Border(
    left=Side(style='thin', color='E0E0E0'),
    right=Side(style='thin', color='E0E0E0'),
    top=Side(style='thin', color='E0E0E0'),
    bottom=Side(style='thin', color='E0E0E0')
)
BORDER_MEDIUM = Border(
    left=Side(style='medium', color='1B2A4A'),
    right=Side(style='medium', color='1B2A4A'),
    top=Side(style='medium', color='1B2A4A'),
    bottom=Side(style='medium', color='1B2A4A')
)
BORDER_BOTTOM = Border(bottom=Side(style='medium', color='C5A500'))

ALIGN_CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)
ALIGN_LEFT = Alignment(horizontal='left', vertical='center', wrap_text=True)
ALIGN_RIGHT = Alignment(horizontal='right', vertical='center', wrap_text=True)

FONT_TITLE = Font(name='Calibri', size=28, bold=True, color='FFFFFF')
FONT_KPI = Font(name='Calibri', size=32, bold=True)
FONT_BODY = Font(name='Calibri', size=10, color='424242')
FONT_BOLD = Font(name='Calibri', size=10, bold=True, color='424242')
FONT_HEADER = Font(name='Calibri', size=10, bold=True, color='FFFFFF')
FONT_STAT_GREEN = Font(name='Calibri', size=10, bold=True, color='1B5E20')
FONT_STAT_RED = Font(name='Calibri', size=10, bold=True, color='C62828')

# Listas TODAS como strings
copas_lista = [str(x) for x in sorted(df_cups['Year'].unique().tolist())]
paises_lista = sorted(df_countries['CountryName'].unique().tolist())
jogadores_lista = sorted(df_players['Name'].unique().tolist())[:500]
posicoes_lista = ['GK', 'CB', 'RB', 'LB', 'CDM', 'CM', 'CAM', 'RW', 'LW', 'ST']
fases_lista = ['Group', 'Round16', 'Quarter', 'Semi', 'Final', 'ThirdPlace']
times_lista = sorted(df_teams['TeamCode'].unique().tolist())

print("Setup completo - criando workbook...")

wb = Workbook()

def style_header(ws, row, max_col, fill, font=Font(name='Calibri', size=10, bold=True, color='FFFFFF')):
    for col in range(1, max_col + 1):
        cell = ws.cell(row=row, column=col)
        cell.font = font
        cell.fill = fill
        cell.alignment = ALIGN_CENTER
        cell.border = BORDER_THIN

def add_color_scale(ws, col_letter, start_row, end_row):
    rule = ColorScaleRule(
        start_type='min', start_color='C62828',
        mid_type='percentile', mid_value=50, mid_color='FFF9C4',
        end_type='max', end_color='1B5E20'
    )
    ws.conditional_formatting.add(f"{col_letter}{start_row}:{col_letter}{end_row}", rule)

# ============================================================
# ABA 1: DASHBOARD EXECUTIVO
# ============================================================
ws1 = wb.active
ws1.title = "Dashboard Executivo"
ws1.sheet_properties.tabColor = '0D1B2A'
ws1.sheet_view.showGridLines = False

col_widths = [3, 18, 18, 18, 18, 18, 18, 18, 18, 18, 18, 18, 3]
for i, w in enumerate(col_widths, 1):
    ws1.column_dimensions[get_column_letter(i)].width = w
ws1.sheet_view.showGridLines = False

for row in [1, 2]:
    for col in range(1, 13):
        ws1.cell(row=row, column=col).fill = FILL_DARK

ws1.merge_cells('B3:K3')
c = ws1['B3']
c.value = "ALBUM DA COPA DEFINITIVO 2010-2026"
c.font = FONT_TITLE
c.alignment = ALIGN_CENTER
c.fill = FILL_DARK
ws1.row_dimensions[3].height = 50

ws1.merge_cells('B4:K4')
c = ws1['B4']
c.value = "Dashboard Executivo - 5 Copas - 4.048 Jogadores - 344 Partidas - 4.846 Eventos"
c.font = Font(name='Calibri', size=13, color='FFD700')
c.alignment = ALIGN_CENTER
c.fill = FILL_DARK
ws1.row_dimensions[4].height = 30

# KPI Cards
kpis = [
    ("GOLS TOTAIS", f"{df_events[df_events['EventType']=='Goal']['xG'].sum():.0f}", "Soma de xG dos gols", '1B5E20'),
    ("PARTIDAS", f"{len(df_matches):,}", "64x4 + 104 = 344", '1B2A4A'),
    ("JOGADORES", f"{len(df_players):,}", "23 x 176 elencos", 'C5A500'),
    ("PAISES", f"{len(df_countries):,}", "6 confederações", 'E65100'),
    ("EVENTOS", f"{len(df_events):,}", "Gols, cartões, subs", '1565C0'),
]

for i, (label, value, desc, color) in enumerate(kpis):
    col = 2 + i * 2
    for r in range(6, 10):
        for c in [col, col+1]:
            cell = ws1.cell(row=r, column=c)
            cell.fill = FILL_WHITE
            cell.border = BORDER_MEDIUM
    
    cell = ws1.cell(row=6, column=col)
    cell.value = ""
    cell.font = Font(size=24)
    cell.alignment = ALIGN_CENTER
    
    cell = ws1.cell(row=7, column=col)
    cell.value = value
    cell.font = Font(name='Calibri', size=32, bold=True, color=color)
    cell.alignment = ALIGN_CENTER
    
    cell = ws1.cell(row=8, column=col)
    cell.value = label
    cell.font = Font(name='Calibri', size=12, bold=True, color='424242')
    cell.alignment = ALIGN_CENTER
    
    cell = ws1.cell(row=9, column=col)
    cell.value = desc
    cell.font = Font(name='Calibri', size=9, color='9E9E9E')
    cell.alignment = ALIGN_CENTER

ws1.row_dimensions[6].height = 15
ws1.row_dimensions[7].height = 45
ws1.row_dimensions[8].height = 20
ws1.row_dimensions[9].height = 18

# Filtros
ws1.merge_cells('B11:K11')
c = ws1['B11']
c.value = "FILTROS RAPIDOS"
c.font = Font(name='Calibri', size=14, bold=True, color='0D1B2A')
c.alignment = Alignment(horizontal='left', vertical='center')
c.fill = FILL_GRAY_100
c.border = BORDER_BOTTOM
ws1.row_dimensions[11].height = 30

filtros = [
    ("Copa:", "B12", copas_lista),
    ("Pais:", "D12", paises_lista[:100]),
    ("Jogador:", "F12", jogadores_lista[:200]),
    ("Posicao:", "H12", posicoes_lista),
    ("Fase:", "J12", fases_lista),
]

for label, cell_ref, source in filtros:
    row = int(''.join(filter(str.isdigit, cell_ref)))
    col_letter = ''.join(filter(str.isalpha, cell_ref))
    col = ord(col_letter) - 64
    
    ws1.cell(row=row, column=col, value=label).font = Font(bold=True, size=11, color='424242')
    ws1.cell(row=row, column=col).fill = FILL_GRAY_100
    ws1.cell(row=row, column=col).border = BORDER_THIN
    ws1.cell(row=row, column=col).alignment = ALIGN_RIGHT
    
    input_col = col + 1
    ws1.cell(row=row, column=input_col).fill = FILL_WHITE
    ws1.cell(row=row, column=input_col).border = BORDER_THIN
    
    dv = DataValidation(type="list", formula1=f'"{",".join(source)}"', allow_blank=True)
    ws1.add_data_validation(dv)
    dv.add(ws1.cell(row=row, column=input_col))

# Navegação
ws1.merge_cells('B14:K14')
c = ws1['B14']
c.value = "NAVEGACAO RAPIDA"
c.font = Font(name='Calibri', size=14, bold=True, color='0D1B2A')
c.alignment = Alignment(horizontal='left', vertical='center')
c.fill = FILL_GRAY_100
c.border = BORDER_BOTTOM
ws1.row_dimensions[14].height = 30

nav_buttons = [
    ("Figurinha do Jogador", "'Figurinha'!A1"),
    ("Perfil da Selecao", "'Selecao'!A1"),
    ("Comparar Copas", "'Comparar Copas'!A1"),
    ("Partidas Historicas", "'Partidas'!A1"),
    ("Eventos (Gols/Cartoes)", "'Eventos'!A1"),
    ("Stats Avancadas (Per 90, Aging)", "'Stats Avancadas'!A1"),
    ("Fonte Power BI", "'Power BI'!A1"),
    ("Medidas DAX", "'Medidas DAX'!A1"),
]

for i, (label, target) in enumerate(nav_buttons):
    row = 15 + i
    ws1.merge_cells(f'B{row}:G{row}')
    cell = ws1.cell(row=row, column=2)
    cell.value = f"  >  {label}"
    cell.font = Font(name='Calibri', size=12, color='1B2A4A', underline='single')
    cell.alignment = Alignment(horizontal='left', vertical='center')
    cell.fill = FILL_WHITE
    cell.border = BORDER_THIN
    cell.hyperlink = f"#{target}"
    cell.style = "Hyperlink"
    
    ws1.merge_cells(f'H{row}:K{row}')
    d = ws1.cell(row=row, column=8)
    d.value = "Clique para ir a aba"
    d.font = Font(name='Calibri', size=10, color='9E9E9E', italic=True)
    d.alignment = Alignment(horizontal='left', vertical='center')
    d.fill = FILL_WHITE
    d.border = BORDER_THIN

# Footer
ws1.merge_cells('B25:K25')
c = ws1['B25']
c.value = f"Gerado em {datetime.now().strftime('%d/%m/%Y %H:%M')} - Investigation Team AI (Rise Kujikawa) - Dados simulados baseados em estatisticas reais"
c.font = Font(name='Calibri', size=9, color='9E9E9E', italic=True)
c.alignment = Alignment(horizontal='center', vertical='center')
c.fill = FILL_DARK
ws1.row_dimensions[25].height = 25

print("Aba 1: Dashboard Executivo")

# ============================================================
# ABA 2: FIGURINHA DO JOGADOR
# ============================================================
ws2 = wb.create_sheet("Figurinha")
ws2.sheet_properties.tabColor = 'C5A500'
ws2.sheet_view.showGridLines = False

for i in range(1, 15):
    ws2.column_dimensions[get_column_letter(i)].width = 16

for row in [1, 2]:
    for col in range(1, 15):
        ws2.cell(row=row, column=col).fill = FILL_DARK

ws2.merge_cells('B3:N3')
c = ws2['B3']
c.value = "FIGURINHA DO JOGADOR - SELECIONE NOS FILTROS ABAIXO"
c.font = Font(name='Calibri', size=20, bold=True, color='FFD700')
c.alignment = ALIGN_CENTER
c.fill = FILL_DARK
ws2.row_dimensions[3].height = 40

# Filtros
filtros_player = [
    ("Buscar Jogador:", "B5", jogadores_lista[:300]),
    ("Copa:", "F5", copas_lista),
    ("Pais:", "H5", paises_lista[:100]),
    ("Posicao:", "K5", posicoes_lista),
]

for label, cell_ref, source in filtros_player:
    row = int(''.join(filter(str.isdigit, cell_ref)))
    col_letter = ''.join(filter(str.isalpha, cell_ref))
    col = ord(col_letter) - 64
    
    ws2.cell(row=row, column=col, value=label).font = Font(bold=True, size=11, color='FFFFFF')
    ws2.cell(row=row, column=col).fill = FILL_PRIMARY
    ws2.cell(row=row, column=col).alignment = ALIGN_RIGHT
    ws2.cell(row=row, column=col).border = BORDER_THIN
    
    input_col = col + 1
    ws2.cell(row=row, column=input_col).fill = FILL_WHITE
    ws2.cell(row=row, column=input_col).border = BORDER_THIN
    
    dv = DataValidation(type="list", formula1=f'"{",".join(source)}"', allow_blank=True)
    ws2.add_data_validation(dv)
    dv.add(ws2.cell(row=row, column=input_col))

# Area da Figurinha
for r in range(7, 36):
    for c in range(2, 14):
        cell = ws2.cell(row=r, column=c)
        cell.fill = FILL_WHITE
        cell.border = BORDER_THIN

# Foto placeholder
ws2.merge_cells('B7:D15')
c = ws2['B7']
c.value = "FOTO DO JOGADOR (Placeholder)"
c.font = Font(name='Calibri', size=14, color='BDBDBD')
c.alignment = ALIGN_CENTER
c.fill = FILL_GRAY_100
c.border = BORDER_MEDIUM

# Info principal
info_fields = [
    ("Nome:", "player_name"),
    ("Pais:", "player_country"),
    ("Copa:", "player_cup"),
    ("Posicao:", "player_pos"),
    ("Idade:", "player_age"),
    ("Altura/Peso:", "player_hw"),
    ("Pe:", "player_foot"),
    ("Valor Mercado:", "player_value"),
    ("Rating Overall:", "player_rating"),
]

for i, (label, key) in enumerate(info_fields):
    row = 7 + i
    ws2.merge_cells(f'E{row}:F{row}')
    cell = ws2.cell(row=row, column=5)
    cell.value = label
    cell.font = Font(bold=True, size=11, color='1B2A4A')
    cell.alignment = ALIGN_RIGHT
    cell.fill = FILL_GRAY_50
    cell.border = BORDER_THIN
    
    ws2.merge_cells(f'G{row}:N{row}')
    cell = ws2.cell(row=row, column=7)
    cell.value = "-"
    cell.font = Font(size=11, color='424242')
    cell.alignment = ALIGN_LEFT
    cell.fill = FILL_WHITE
    cell.border = BORDER_THIN

# Radar Chart
ws2.merge_cells('B17:N17')
c = ws2['B17']
c.value = "RADAR CHART - PERFIL POR POSICAO (0-100)"
c.font = Font(bold=True, size=13, color='1B2A4A')
c.alignment = ALIGN_CENTER
c.fill = FILL_GRAY_100
c.border = BORDER_BOTTOM

ws2.merge_cells('B18:N25')
c = ws2['B18']
c.value = "RADAR CHART - Sera preenchido via VBA/UserForm com 6 eixos especificos por posicao:\n\nGK: Defesas, Clean Sheets, Jogo c/ Pes, Saidas, Penaltis, Consistencia\nDEF: Desarmes, Interceptacoes, Cortes, Jogo Aereo, Passes, Disciplina\nMID: Passes, Progressao, Criacao, Defesa, Resistencia, Versatilidade\nFWD: Finalizacao, xG, Movimento, Jogo Aereo, Criacao, Clutch"
c.font = Font(size=11, color='757575')
c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
c.fill = FILL_GRAY_50
c.border = BORDER_MEDIUM

# Stats da Copa
ws2.merge_cells('B27:N27')
c = ws2['B27']
c.value = "ESTATISTICAS DA COPA SELECIONADA"
c.font = Font(bold=True, size=13, color='1B2A4A')
c.alignment = ALIGN_CENTER
c.fill = FILL_GRAY_100
c.border = BORDER_BOTTOM

stat_headers = ['Partidas', 'Minutos', 'Gols', 'Assist.', 'xG', 'Finaliz.', 'No Alvo', 'Dribles', 'Passes %', 'Desarmes', 'Intercept.', 'Cortes', 'Faltas', 'Amarelos', 'Vermelhos', 'Clean Sh.', 'Defesas', 'GC']
for i, h in enumerate(stat_headers, 1):
    c = ws2.cell(row=28, column=i+1)
    c.value = h
    c.font = Font(bold=True, size=10, color='FFFFFF')
    c.fill = FILL_PRIMARY
    c.alignment = ALIGN_CENTER
    c.border = BORDER_THIN

for col in range(2, 20):
    c = ws2.cell(row=29, column=col)
    c.value = 0
    c.font = Font(size=11, color='424242')
    c.alignment = ALIGN_CENTER
    c.fill = FILL_WHITE
    c.border = BORDER_THIN
    c.number_format = '#,##0.00' if col in [3,4,5,9] else '#,##0'

# Per 90
ws2.merge_cells('B31:N31')
c = ws2['B31']
c.value = "PER 90 MINUTOS"
c.font = Font(bold=True, size=13, color='1B2A4A')
c.alignment = ALIGN_CENTER
c.fill = FILL_GRAY_100
c.border = BORDER_BOTTOM

per90_headers = ['G/90', 'A/90', 'xG/90', 'Final/90', 'Passes/90', 'Prec%', 'Drib/90', 'Des/90', 'Int/90', 'Faltas/90']
for i, h in enumerate(per90_headers, 1):
    c = ws2.cell(row=32, column=i+1)
    c.value = h
    c.font = Font(bold=True, size=10, color='1B2A4A')
    c.fill = FILL_GOLD
    c.alignment = ALIGN_CENTER
    c.border = BORDER_THIN

for col in range(2, 12):
    c = ws2.cell(row=33, column=col)
    c.value = 0.00
    c.font = Font(size=11, color='1B5E20')
    c.alignment = ALIGN_CENTER
    c.fill = FILL_WHITE
    c.border = BORDER_THIN
    c.number_format = '0.00'

# Historico
ws2.merge_cells('B35:N35')
c = ws2['B35']
c.value = "HISTORICO EM COPAS DO MUNDO"
c.font = Font(bold=True, size=13, color='1B2A4A')
c.alignment = ALIGN_CENTER
c.fill = FILL_GRAY_100
c.border = BORDER_BOTTOM

hist_headers = ['Copa', 'Pais', 'Idade', 'Partidas', 'Minutos', 'Gols', 'Assist.', 'xG', 'Rating', 'Fase Alcancada']
for i, h in enumerate(hist_headers, 1):
    c = ws2.cell(row=36, column=i+1)
    c.value = h
    c.font = Font(bold=True, size=10, color='FFFFFF')
    c.fill = FILL_PRIMARY
    c.alignment = ALIGN_CENTER
    c.border = BORDER_THIN

for r in range(37, 42):
    for c in range(2, 12):
        cell = ws2.cell(row=r, column=c)
        cell.value = "-"
        cell.alignment = ALIGN_CENTER
        cell.font = Font(size=10, color='BDBDBD')
        cell.fill = FILL_WHITE if r % 2 == 0 else FILL_GRAY_50
        cell.border = BORDER_THIN

# Percentis
ws2.merge_cells('B43:N43')
c = ws2['B43']
c.value = "PERCENTIS (POSICAO) - ONDE ESTA ESTE JOGADOR"
c.font = Font(bold=True, size=13, color='1B2A4A')
c.alignment = ALIGN_CENTER
c.fill = FILL_GRAY_100
c.border = BORDER_BOTTOM

pct_headers = ['Metrica', 'P10', 'P25', 'P50 (Mediana)', 'P75', 'P90', 'Media', 'Desvio', 'Valor Jogador', 'Percentil']
for i, h in enumerate(pct_headers, 1):
    c = ws2.cell(row=44, column=i+1)
    c.value = h
    c.font = Font(bold=True, size=10, color='1B2A4A')
    c.fill = FILL_GOLD
    c.alignment = ALIGN_CENTER
    c.border = BORDER_THIN

metrics_pct = ['Gols', 'Assistencias', 'xG', 'Finalizacoes', 'Passes%', 'Desarmes', 'Dribles', 'Rating', 'Valor Mercado']
for r_idx, met in enumerate(metrics_pct):
    row = 45 + r_idx
    ws2.cell(row=row, column=2, value=met).font = Font(bold=True, size=10)
    ws2.cell(row=row, column=2).alignment = ALIGN_LEFT
    ws2.cell(row=row, column=2).border = BORDER_THIN
    for c in range(3, 12):
        cell = ws2.cell(row=row, column=c)
        cell.value = 0
        cell.alignment = ALIGN_CENTER
        cell.border = BORDER_THIN
        cell.fill = FILL_WHITE if r_idx % 2 == 0 else FILL_GRAY_50

print("Aba 2: Figurinha do Jogador")

# Salvar Parte 1
wb.save(f"{base_dir}/Album_Copa_Definitivo_PARTE1.xlsx")
print("Parte 1 salva com sucesso!")