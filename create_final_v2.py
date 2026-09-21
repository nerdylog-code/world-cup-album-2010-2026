import os
import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Color
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, Reference
from openpyxl.formatting.rule import ColorScaleRule
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

print("Dados carregados:", len(df_players), "jogadores")

PALETTE = {
    'navy_900': '0A0F1A', 'navy_800': '111827', 'navy_700': '1E2A3A', 'navy_600': '2D3A4F', 'navy_500': '3D4A5F',
    'gold_500': 'FFD700', 'gold_400': 'FFC107', 'gold_300': 'FFD54F', 'gold_200': 'FFF176', 'gold_100': 'FFFDE7',
    'grass_600': '1B5E20', 'grass_500': '2E7D32', 'grass_400': '4CAF50', 'grass_300': '81C784', 'grass_200': 'C8E6C9', 'grass_100': 'E8F5E9',
    'red_600': 'C62828', 'red_500': 'D32F2F', 'red_400': 'EF5350', 'red_300': 'E57373', 'red_200': 'FFCDD2', 'red_100': 'FFEBEE',
    'orange_600': 'E65100', 'orange_500': 'F57C00', 'orange_400': 'FF9800', 'orange_300': 'FFB74D', 'orange_200': 'FFE0B2', 'orange_100': 'FFF3E0',
    'blue_600': '1565C0', 'blue_500': '1976D2', 'blue_400': '2196F3', 'blue_300': '64B5F6', 'blue_200': 'BBDEFB', 'blue_100': 'E3F2FD',
    'white': 'FFFFFF', 'gray_50': 'FAFAFA', 'gray_100': 'F5F5F5', 'gray_200': 'EEEEEE', 'gray_300': 'E0E0E0',
    'gray_400': 'BDBDBD', 'gray_500': '9E9E9E', 'gray_600': '757575', 'gray_700': '616161', 'gray_800': '424242', 'gray_900': '212121',
}

POS_COLORS = {'GK': 'FFD700', 'CB': '2196F3', 'RB': '2196F3', 'LB': '2196F3', 'CDM': '4CAF50', 'CM': '4CAF50', 'CAM': 'FF9800', 'RW': 'F44336', 'LW': 'F44336', 'ST': 'E91E63'}

def mk_fill(c): return PatternFill(start_color=c, end_color=c, fill_type='solid')
def mk_font(**k): d={'name':'Calibri','color':PALETTE['gray_800']}; d.update(k); return Font(**d)
def mk_align(**k): d={'horizontal':'center','vertical':'center','wrap_text':True}; d.update(k); return Alignment(**d)
def mk_border(**k): 
    d={'left':Side(style='thin',color=PALETTE['gray_300']),'right':Side(style='thin',color=PALETTE['gray_300']),
       'top':Side(style='thin',color=PALETTE['gray_300']),'bottom':Side(style='thin',color=PALETTE['gray_300'])}
    d.update(k); return Border(**d)

FILL_N9=mk_fill(PALETTE['navy_900']); FILL_N8=mk_fill(PALETTE['navy_800']); FILL_N7=mk_fill(PALETTE['navy_700']); FILL_N6=mk_fill(PALETTE['navy_600'])
FILL_GOLD=mk_fill(PALETTE['gold_500']); FILL_GOLD_L=mk_fill(PALETTE['gold_200']); FILL_GOLD_S=mk_fill(PALETTE['gold_100'])
FILL_GR=mk_fill(PALETTE['grass_500']); FILL_GR_L=mk_fill(PALETTE['grass_100']); FILL_GR_S=mk_fill(PALETTE['grass_200']); FILL_GR7=mk_fill(PALETTE['grass_600'])
FILL_RED_L=mk_fill(PALETTE['red_100']); FILL_OR_L=mk_fill(PALETTE['orange_100']); FILL_BL_L=mk_fill(PALETTE['blue_100'])
FILL_W=mk_fill(PALETTE['white']); FILL_G50=mk_fill(PALETTE['gray_50']); FILL_G100=mk_fill(PALETTE['gray_100']); FILL_G200=mk_fill(PALETTE['gray_200'])

B_THIN=mk_border()
B_CARD=mk_border(left=Side(style='thin',color=PALETTE['gray_300']),right=Side(style='thin',color=PALETTE['gray_300']),
                 top=Side(style='thin',color=PALETTE['gray_300']),bottom=Side(style='thin',color=PALETTE['gray_300']))
B_CARD_THICK=mk_border(left=Side(style='medium',color=PALETTE['gold_500']),right=Side(style='medium',color=PALETTE['gold_500']),
                       top=Side(style='medium',color=PALETTE['gold_500']),bottom=Side(style='medium',color=PALETTE['gold_500']))
B_BOT_GOLD=mk_border(bottom=Side(style='medium',color=PALETTE['gold_500']))
B_BOT_NAVY=mk_border(bottom=Side(style='medium',color=PALETTE['navy_700']))

A_C=mk_align(); A_L=mk_align(horizontal='left'); A_R=mk_align(horizontal='right')

F_T_MAIN=mk_font(size=28,bold=True,color=PALETTE['white'])
F_T_SEC=mk_font(size=16,bold=True,color=PALETTE['navy_900'])
F_KPI_V=mk_font(size=36,bold=True,color=PALETTE['navy_900'])
F_KPI_L=mk_font(size=11,bold=True,color=PALETTE['gray_600'])
F_HDR=mk_font(size=10,bold=True,color=PALETTE['white'])
F_HDR_D=mk_font(size=10,bold=True,color=PALETTE['navy_900'])
F_B=mk_font(size=10,color=PALETTE['gray_800'])
F_BB=mk_font(size=10,bold=True,color=PALETTE['gray_800'])
F_BS=mk_font(size=9,color=PALETTE['gray_700'])
F_CAP=mk_font(size=8,color=PALETTE['gray_500'],italic=True)
F_LNK=mk_font(size=11,color=PALETTE['blue_600'],underline='single')
F_SP=mk_font(size=10,bold=True,color=PALETTE['grass_600'])
F_SN=mk_font(size=10,bold=True,color=PALETTE['red_600'])
F_SN2=mk_font(size=10,color=PALETTE['gray_700'])
F_WB=mk_font(size=10,bold=True,color=PALETTE['white'])
F_WB14=mk_font(size=14,bold=True,color=PALETTE['white'])

copas_lista=[str(x) for x in sorted(df_cups['Year'].unique().tolist())]
paises_lista=sorted(df_countries['CountryName'].unique().tolist())
jogadores_lista=sorted(df_players['Name'].unique().tolist())[:500]
posicoes_lista=['GK','CB','RB','LB','CDM','CM','CAM','RW','LW','ST']
fases_lista=['Group','Round16','Quarter','Semi','Final','ThirdPlace']
times_lista=sorted(df_teams['TeamCode'].unique().tolist())

wb=Workbook()

def mstyle(ws,r,cs,ce,val,font=F_B,fill=FILL_W,align=A_L,border=B_THIN):
    ws.merge_cells(start_row=r,start_column=cs,end_row=r,end_column=ce)
    c=ws.cell(row=r,column=cs,value=val)
    c.font=font; c.fill=fill; c.alignment=align; c.border=border
    return c

def scell(ws,r,c,val,font=F_B,fill=FILL_W,align=A_C,border=B_THIN,nfmt=None):
    cell=ws.cell(row=r,column=c,value=val)
    cell.font=font; cell.fill=fill; cell.alignment=align; cell.border=border
    if nfmt: cell.number_format=nfmt
    return cell

def hdr_row(ws,r,maxc,fill=FILL_N8,font=F_HDR,h=28):
    for c in range(1,maxc+1):
        cell=ws.cell(row=r,column=c)
        cell.font=font; cell.fill=fill; cell.alignment=A_C; cell.border=B_THIN
    ws.row_dimensions[r].height=h

def kpi_card(ws,rs,c,label,val,sub,vcolor=PALETTE['navy_900'],icon=""):
    for r in range(rs,rs+4):
        for cc in [c,c+1]:
            cell=ws.cell(row=r,column=cc)
            cell.fill=FILL_W; cell.border=B_CARD
    mstyle(ws,rs,c,c+1,icon,mk_font(size=20,color=vcolor),FILL_W,A_C,B_CARD)
    mstyle(ws,rs+1,c,c+1,val,mk_font(size=32,bold=True,color=vcolor),FILL_W,A_C,B_CARD)
    mstyle(ws,rs+2,c,c+1,label,F_KPI_L,FILL_W,A_C,B_CARD)
    mstyle(ws,rs+3,c,c+1,sub,F_CAP,FILL_W,A_C,B_CARD)

def sec_hdr(ws,r,cs,ce,title,sub="",fill=FILL_N8,h=36):
    mstyle(ws,r,cs,ce,title,F_T_SEC,fill,mk_align(horizontal='left',vertical='center'),B_BOT_GOLD)
    ws.row_dimensions[r].height=h
    nr=r+1
    if sub:
        mstyle(ws,nr,cs,ce,sub,F_CAP,fill,mk_align(horizontal='left',vertical='center'),B_THIN)
        ws.row_dimensions[nr].height=20
        nr+=1
    return nr

# ============================================================
# ABA 1: DASHBOARD
# ============================================================
ws1=wb.active; ws1.title="📊 Dashboard"; ws1.sheet_properties.tabColor=PALETTE['navy_900']
ws1.sheet_view.showGridLines=False; ws1.sheet_view.zoomScale=90

for i,w in enumerate([2,16,16,16,16,16,16,16,16,16,16,16,2],1):
    ws1.column_dimensions[get_column_letter(i)].width=w

for row in [1,2]:
    for col in range(1,13): ws1.cell(row=row,column=col).fill=FILL_N9

mstyle(ws1,3,2,12,"ÁLBUM DA COPA DEFINITIVO 2010–2026",F_T_MAIN,FILL_N9,mk_align(horizontal='center',vertical='center'))
ws1.row_dimensions[3].height=56
mstyle(ws1,4,2,12,"Investigation Team AI  •  5 Copas  •  4.048 Jogadores  •  344 Partidas  •  4.846 Eventos",
       mk_font(size=12,color=PALETTE['gold_300']),FILL_N9,mk_align(horizontal='center',vertical='center'))
ws1.row_dimensions[4].height=28

kpis=[("Gols Totais (xG)",f"{df_events[df_events['EventType']=='Goal']['xG'].sum():,.0f}","Soma de Expected Goals",PALETTE['grass_600'],"⚽"),
      ("Partidas",f"{len(df_matches):,}","64×4 + 104 = 344",PALETTE['blue_600'],"🏟️"),
      ("Jogadores Únicos",f"{len(df_players):,}","23 × 176 elencos",PALETTE['gold_500'],"👤"),
      ("Países / Confed.",f"{len(df_countries)} / 6","UEFA, CONMEBOL, CAF, AFC, CONCACAF, OFC",PALETTE['orange_600'],"🌍"),
      ("Eventos Registrados",f"{len(df_events):,}","Gols, Cartões, Substituições",PALETTE['red_600'],"📝")]

for i,(lab,val,sub,col,ico) in enumerate(kpis):
    kpi_card(ws1,6,2+i*2,lab,val,sub,col,ico)

for r in range(6,10): ws1.row_dimensions[r].height=[12,42,22,18][r-6]

mstyle(ws1,11,2,12,"",FILL_W,FILL_N8,A_C,B_BOT_GOLD); ws1.row_dimensions[11].height=4

sec_hdr(ws1,12,2,11,"🔍  FILTROS RÁPIDOS","Selecione para filtrar todas as abas vinculadas",FILL_N8,32)

filtros=[("🏆 Copa:",2,copas_lista),("🌍 País:",4,paises_lista),("👤 Jogador:",6,jogadores_lista[:200]),
         ("⚽ Posição:",8,posicoes_lista),("📋 Fase:",10,fases_lista)]
for lab,col,src in filtros:
    row=14
    c=ws1.cell(row=row,column=col,value=lab); c.font=F_BB; c.fill=FILL_G100; c.alignment=A_R; c.border=B_THIN
    ic=col+1; c=ws1.cell(row=row,column=ic); c.fill=FILL_W; c.border=B_THIN
    dv=DataValidation(type="list",formula1=f'"{",".join(src)}"',allow_blank=True)
    dv.error="Selecione uma opção válida"; dv.errorTitle="Entrada Inválida"
    ws1.add_data_validation(dv); dv.add(ws1.cell(row=row,column=ic))
ws1.row_dimensions[14].height=28

nav_items=[("🎴  Figurinha do Jogador","'Figurinha'!A1","Perfil completo estilo álbum de figurinhas"),
           ("🌍  Perfil da Seleção","'Seleção'!A1","Elenco, formação tática, histórico"),
           ("📈  Comparar Copas","'Comparar Copas'!A1","Evolução 2010–2026, gráficos, small multiples"),
           ("🏟️  Partidas Históricas","'Partidas'!A1","Todas 344 partidas com xG e color scales"),
           ("📝  Eventos (Gols/Cartões)","'Eventos'!A1","4.846 eventos detalhados por partida"),
           ("📊  Stats Avançadas","'Stats Avançadas'!A1","Per 90, Percentis P10–P90, Aging Curves"),
           ("🔗  Fonte Power BI","'Power BI'!A1","Modelo dimensional, relacionamentos, DAX"),
           ("📐  Medidas DAX","'Medidas DAX'!A1","20+ medidas prontas para copiar/colar")]

nr=sec_hdr(ws1,16,2,11,"🧭  NAVEGAÇÃO RÁPIDA","Clique para ir direto à aba",FILL_N8,32)
for i,(lab,tgt,desc) in enumerate(nav_items):
    row=nr+i
    mstyle(ws1,row,2,7,lab,F_LNK,FILL_W if i%2==0 else FILL_G50,A_L,B_THIN)
    ws1.cell(row=row,column=2).hyperlink=f"#{tgt}"
    mstyle(ws1,row,8,11,desc,F_CAP,FILL_W if i%2==0 else FILL_G50,A_L,B_THIN)
    ws1.row_dimensions[row].height=26

fr=26
mstyle(ws1,fr,2,11,f"Gerado em {datetime.now().strftime('%d/%m/%Y %H:%M')}  •  Investigation Team AI (Rise Kujikawa)  •  Dados simulados baseados em estatísticas reais  •  v2.0 Professional",
       F_CAP,FILL_N9,mk_align(horizontal='center',vertical='center'))
ws1.row_dimensions[fr].height=32

print("✅ Aba 1: Dashboard")

# ============================================================
# ABA 2: FIGURINHA
# ============================================================
ws2=wb.create_sheet("🎴 Figurinha")
ws2.sheet_properties.tabColor=PALETTE['gold_500']
ws2.sheet_view.showGridLines=False; ws2.sheet_view.zoomScale=85

for i in range(1,18):
    if i in [1,17]: ws2.column_dimensions[get_column_letter(i)].width=2
    elif i in [2,3,4]: ws2.column_dimensions[get_column_letter(i)].width=14
    elif i in [5,6]: ws2.column_dimensions[get_column_letter(i)].width=12
    elif 7<=i<=13: ws2.column_dimensions[get_column_letter(i)].width=10
    else: ws2.column_dimensions[get_column_letter(i)].width=12

for row in [1,2]:
    for col in range(1,18): ws2.cell(row=row,column=col).fill=FILL_N9

mstyle(ws2,3,2,16,"🎴  FIGURINHA DO JOGADOR",F_T_MAIN,FILL_N9,mk_align(horizontal='left',vertical='center'))
ws2.row_dimensions[3].height=50
mstyle(ws2,4,2,16,"Selecione nos filtros abaixo → Navegue com ◄ ►  •  Compare com outro jogador",
       mk_font(size=11,color=PALETTE['gold_300']),FILL_N9,mk_align(horizontal='left',vertical='center'))
ws2.row_dimensions[4].height=24

sec_hdr(ws2,6,2,16,"🔍  FILTROS","",FILL_N8,32)

filtros_p=[("👤 Jogador:",3,jogadores_lista[:300]),("🏆 Copa:",6,copas_lista),
           ("🌍 País:",8,paises_lista),("⚽ Pos:",11,posicoes_lista)]
for lab,col,src in filtros_p:
    row=7
    c=ws2.cell(row=row,column=col,value=lab)
    c.font=mk_font(size=10,bold=True,color=PALETTE['white']); c.fill=FILL_N7; c.alignment=A_R; c.border=B_THIN
    ic=col+1; c=ws2.cell(row=row,column=ic); c.fill=FILL_W; c.border=B_THIN
    dv=DataValidation(type="list",formula1=f'"{",".join(src)}"',allow_blank=True)
    ws2.add_data_validation(dv); dv.add(ws2.cell(row=row,column=ic))

nr=9
btns=[(3,5,"◄  Anterior",FILL_N7,PALETTE['white']),(6,8,"🔍  Buscar",FILL_GOLD,PALETTE['navy_900']),
      (9,11,"⚖  Comparar",FILL_N7,PALETTE['white']),(12,14,"Próximo  ►",FILL_N7,PALETTE['white'])]
for cs,ce,txt,fl,fc in btns:
    mstyle(ws2,nr,cs,ce,txt,mk_font(size=10,bold=True,color=fc),fl,A_C,B_THIN)
ws2.row_dimensions[nr].height=30

cs=11
mstyle(ws2,cs,3,5,"📷  FOTO\n\n[Placeholder:\n200×250px\n\nAdicione imagem\ndo jogador]",
       mk_font(size=11,color=PALETTE['gray_500']),FILL_G100,mk_align(horizontal='center',vertical='center'),B_CARD_THICK)

info_f=[("Nome","lbl_name","—"),("País","lbl_country","—"),("Copa","lbl_cup","—"),
        ("Posição","lbl_pos","—"),("Idade","lbl_age","—"),("Altura / Peso","lbl_hw","—"),
        ("Pé","lbl_foot","—"),("Valor Mercado","lbl_value","—"),("Overall Rating","lbl_rating","—")]

for i,(lab,key,defv) in enumerate(info_f):
    row=cs+i
    mstyle(ws2,row,6,7,lab,F_BB,FILL_G100,A_R,B_THIN)
    ft=F_B
    if i==7: ft=mk_font(size=12,bold=True,color=PALETTE['grass_600'])
    elif i==8: ft=mk_font(size=14,bold=True,color=PALETTE['gold_500'])
    mstyle(ws2,row,8,16,defv,ft,FILL_W,A_L,B_THIN)

rs=22
sec_hdr(ws2,rs,2,16,"📡  RADAR CHART — PERFIL POR POSIÇÃO (0–100)","6 eixos específicos por posição  •  Atualiza automaticamente com a seleção",FILL_N8,32)

radar_txt=("GK: Defesas/90  •  Clean Sheets %  •  Passes/90  •  Saídas/90  •  Pênaltis Def.  •  Consistência\n\n"
           "DEF (CB/RB/LB): Desarmes/90  •  Interceptações/90  •  Cortes/90  •  Aéreos %  •  Passes %  •  Disciplina\n\n"
           "MID (CDM/CM/CAM): Passes/90  •  Progressão/90  •  Criação/90  •  Defesa/90  •  Resistência  •  Versatilidade\n\n"
           "FWD (RW/LW/ST): Finalizações/90  •  xG/90  •  Gols/90  •  Aéreos %  •  Assistências/90  •  Clutch")
mstyle(ws2,24,3,16,radar_txt,F_BS,FILL_G50,mk_align(horizontal='left',vertical='center'),
       mk_border(left=Side(style='thick',color=PALETTE['gold_500']),right=Side(style='thin',color=PALETTE['gray_300']),
                 top=Side(style='thin',color=PALETTE['gray_300']),bottom=Side(style='thin',color=PALETTE['gray_300'])))

mstyle(ws2,30,3,16,"📊  RADAR CHART\n\n[Gráfico Radar nativo do Excel — 6 eixos]\n\nPara ativar: Inserir → Gráfico → Radar → Selecionar dados da tabela abaixo\n\nOu use o UserForm VBA (Alt+F8 → ShowAlbum)",
       mk_font(size=11,color=PALETTE['gray_500']),FILL_G100,mk_align(horizontal='center',vertical='center'),B_CARD)

rds=40
mstyle(ws2,rds,3,16,"📋 Dados do Radar (fonte para gráfico)",F_T_SEC,FILL_N7,A_L,B_BOT_GOLD)

rh=['Eixo','Máximo','Valor Jogador','% do Máx','Percentil Pos.','Média Pos.','Desvio','Nota']
for i,h in enumerate(rh,1): scell(ws2,rds+1,2+i,h,F_HDR_D,FILL_GOLD,A_C,B_THIN)

ra=['Eixo 1','Eixo 2','Eixo 3','Eixo 4','Eixo 5','Eixo 6']
for i,ax in enumerate(ra):
    row=rds+2+i
    scell(ws2,row,3,ax,F_BB,FILL_W if i%2==0 else FILL_G50,A_L,B_THIN)
    for col in range(4,11): scell(ws2,row,col,0,F_B,FILL_W if i%2==0 else FILL_G50,A_C,B_THIN)

ss=rds+10
sec_hdr(ws2,ss,2,16,"📊  ESTATÍSTICAS DA COPA SELECIONADA","Totais absolutos  •  Atualize a Copa no filtro acima",FILL_N8,32)

sh=['PJ','Min','Gols','Ast','xG','Fin','Gol','Drib','Pass%','Des','Int','Cort','FC','CA','CV','CS','Def','GC']
for i,h in enumerate(sh,1): scell(ws2,ss+1,1+i,h,F_HDR_D,FILL_GOLD,A_C,B_THIN)

for i in range(1,19):
    c=scell(ws2,ss+2,i,0,F_B,FILL_W,A_C,B_THIN)
    if i in [3,4,5]: c.number_format='#,##0.00'
    elif i==9: c.number_format='0.0"%"'

ps=ss+4
sec_hdr(ws2,ps,2,16,"⚡  PER 90 MINUTOS","Normalizado por tempo em campo  •  Verde = acima da média da posição",FILL_N8,32)

ph=['G/90','A/90','xG/90','Fin/90','Pass/90','Pass%','Drib/90','Des/90','Int/90','FC/90']
for i,h in enumerate(ph,1): scell(ws2,ps+1,1+i,h,F_HDR_D,FILL_GOLD,A_C,B_THIN)
for i in range(1,11):
    c=scell(ws2,ps+2,i,0.0,F_SP,FILL_GR_S,A_C,B_THIN); c.number_format='0.00'

hs=ps+4
sec_hdr(ws2,hs,2,16,"🏆  HISTÓRICO EM COPAS DO MUNDO","Todas as participações do jogador  •  Ordenado por ano",FILL_N8,32)

hh=['Copa','País','Idade','PJ','Min','Gols','Ast','xG','Rating','Fase','MVP?']
for i,h in enumerate(hh,1): scell(ws2,hs+1,1+i,h,F_HDR_D,FILL_GOLD,A_C,B_THIN)
for r in range(5):
    row=hs+2+r
    for i in range(1,12): scell(ws2,row,i,"—",F_BS,FILL_W if r%2==0 else FILL_G50,A_C,B_THIN)

pc=hs+8
sec_hdr(ws2,pc,2,16,"📈  PERCENTIS vs POSIÇÃO","Onde este jogador se posiciona entre os pares  •  P50 = Mediana",FILL_N8,32)

ph=['Métrica','P10','P25','P50 (Mediana)','P75','P90','Média','Desvio','Jogador','Percentil']
for i,h in enumerate(ph,1): scell(ws2,pc+1,1+i,h,F_HDR_D,FILL_GOLD,A_C,B_THIN)

mp=['Gols','Assistências','xG','Finalizações','Passes%','Desarmes','Dribles','Rating','Valor Mercado']
for ri,met in enumerate(mp):
    row=pc+2+ri
    scell(ws2,row,2,met,F_BB,FILL_W if ri%2==0 else FILL_G50,A_L,B_THIN)
    for col in range(3,11): scell(ws2,row,col,0,F_B,FILL_W if ri%2==0 else FILL_G50,A_C,B_THIN)

fr=pc+13
mstyle(ws2,fr,3,16,"Investigation Team AI  •  Rise Kujikawa  •  Futaba-chan  •  Naoto-san  •  Dados simulados seed 2026",F_CAP,FILL_N9,mk_align(horizontal='center',vertical='center'))
ws2.row_dimensions[fr].height=28

print("✅ Aba 2: Figurinha")

# ============================================================
# ABA 3: SELEÇÃO
# ============================================================
ws3=wb.create_sheet("🌍 Seleção")
ws3.sheet_properties.tabColor=PALETTE['grass_600']
ws3.sheet_view.showGridLines=False; ws3.sheet_view.zoomScale=85

for i in range(1,18):
    ws3.column_dimensions[get_column_letter(i)].width=14 if i not in [1,17] else 2

for row in [1,2]:
    for col in range(1,18): ws3.cell(row=row,column=col).fill=mk_fill(PALETTE['grass_600'])

mstyle(ws3,3,2,16,"🌍  PERFIL DA SELEÇÃO",F_T_MAIN,mk_fill(PALETTE['grass_600']),mk_align(horizontal='left',vertical='center'))
ws3.row_dimensions[3].height=50
mstyle(ws3,4,2,16,"Escolha País e Copa  •  Elenco completo  •  Formação tática  •  Evolução histórica",
       mk_font(size=11,color=PALETTE['gold_300']),mk_fill(PALETTE['grass_600']),mk_align(horizontal='left',vertical='center'))
ws3.row_dimensions[4].height=24

sec_hdr(ws3,6,2,16,"🔍  FILTROS","",mk_fill(PALETTE['grass_600']),32)

tf=[("🌍 País:",3,paises_lista),("🏆 Copa:",6,copas_lista)]
for lab,col,src in tf:
    row=7
    c=ws3.cell(row=row,column=col,value=lab)
    c.font=mk_font(size=10,bold=True,color=PALETTE['white']); c.fill=mk_fill(PALETTE['grass_600']); c.alignment=A_R; c.border=B_THIN
    c=ws3.cell(row=row,column=col+1); c.fill=FILL_W; c.border=B_THIN
    dv=DataValidation(type="list",formula1=f'"{",".join(src)}"',allow_blank=True)
    ws3.add_data_validation(dv); dv.add(ws3.cell(row=row,column=col+1))

is_=10
sec_hdr(ws3,is_,2,16,"📋  INFORMAÇÕES GERAIS","",mk_fill(PALETTE['grass_600']),32)

tfields=[("Nome Oficial","lbl_t_name"),("Confederação","lbl_t_confed"),("Copa","lbl_t_cup"),
         ("Fase Alcançada","lbl_t_stage"),("Partidas","lbl_t_matches"),("V - E - D","lbl_t_record"),
         ("Gols Marcados","lbl_t_gf"),("Gols Sofridos","lbl_t_ga"),("Saldo de Gols","lbl_t_gd"),
         ("Artilheiro da Copa","lbl_t_topscorer"),("Melhor Jogador","lbl_t_bestplayer"),
         ("Rating Médio do Elenco","lbl_t_avg_rating"),("Idade Média","lbl_t_avg_age"),
         ("Valor Total do Elenco","lbl_t_value"),("% Sub-23 no Elenco","lbl_t_pct_u23")]

for i,(lab,key) in enumerate(tfields):
    row=is_+1+i
    mstyle(ws3,row,2,3,lab,F_BB,FILL_GR_L,A_R,B_THIN)
    mstyle(ws3,row,4,16,"—",F_B,FILL_W,A_L,B_THIN)

sq=is_+1+len(tfields)+1
sec_hdr(ws3,sq,2,16,"👥  ELENCO COMPLETO (23 JOGADORES)","Ordenado: Goleiros → Defensores → Meias → Atacantes  •  Por Rating dentro da posição",mk_fill(PALETTE['grass_600']),32)

sh=['#','Jogador','Pos','Idade','Clube','PJ','Min','Gols','Ast','xG','Rating','Valor (M€)']
for i,h in enumerate(sh,1): scell(ws3,sq+1,i,h,F_HDR,FILL_GR,A_C,B_THIN)

po=['GK','CB','CB','CB','CB','RB','RB','LB','LB','CDM','CDM','CM','CM','CM','CAM','RW','RW','LW','LW','ST','ST','ST','ST']
for i,pos in enumerate(po):
    row=sq+2+i
    for col in range(1,13):
        c=ws3.cell(row=row,column=col)
        c.fill=FILL_W if i%2==0 else FILL_G50; c.border=B_THIN; c.alignment=A_C
    ws3.cell(row=row,column=1,value=i+1).font=F_B
    ws3.cell(row=row,column=2,value=f"Jogador {i+1}").font=F_B
    ws3.cell(row=row,column=2).alignment=A_L
    ws3.cell(row=row,column=3,value=pos).font=mk_font(size=10,bold=True,color=POS_COLORS.get(pos,PALETTE['gray_700']))
    for col in range(4,13): ws3.cell(row=row,column=col,value="—")

fs=sq+25
sec_hdr(ws3,fs,2,16,"⚽  FORMAÇÃO TÁTICA BASE (4-2-3-1)","Esquema mais utilizado na Copa  •  ASCII art para visualização rápida",mk_fill(PALETTE['grass_600']),32)

fa=["","                              [GK]                                     ","",
    "          [CB]           [CB]           [CB]                          ",
    "     [RB]                                                            [LB]",
    "","                    [CDM]         [CDM]                               ","",
    "           [CM]                      [CM]                             ","",
    "                          [CAM]                                       ","",
    "      [RW]                                                            [LW]","",
    "                              [ST]                                     ",""]

for i,line in enumerate(fa):
    row=fs+1+i
    mstyle(ws3,row,2,16,line,mk_font(name='Consolas',size=9,color=PALETTE['grass_600']),FILL_GR_L,mk_align(horizontal='center',vertical='center'),B_THIN)

ds=fs+16
sec_hdr(ws3,ds,2,16,"📊  DISTRIBUIÇÃO DO ELENCO POR POSIÇÃO","",mk_fill(PALETTE['grass_600']),32)

dh=['Posição','Qtd','Min Totais','% Min','Gols','Ast','xG','Rating Médio','Valor Médio (M€)']
for i,h in enumerate(dh,1): scell(ws3,ds+1,i,h,F_HDR,FILL_GR,A_C,B_THIN)

pd=['GK','CB','RB','LB','CDM','CM','CAM','RW','LW','ST']
for i,pos in enumerate(pd):
    row=ds+2+i
    scell(ws3,row,1,pos,mk_font(size=10,bold=True,color=POS_COLORS.get(pos,PALETTE['gray_700'])),FILL_W if i%2==0 else FILL_G50,A_C,B_THIN)
    for col in range(2,10): scell(ws3,row,col,0,F_B,FILL_W if i%2==0 else FILL_G50,A_C,B_THIN)

hs=ds+12
sec_hdr(ws3,hs,2,16,"📈  EVOLUÇÃO HISTÓRICA (2010–2026)","Desempenho da seleção em cada Copa",mk_fill(PALETTE['grass_600']),32)

hh=['Copa','Fase','PJ','V','E','D','GP','GC','SG','Artilheiro','Gols','Rating Méd','Valor Elenco (M€)']
for i,h in enumerate(hh,1): scell(ws3,hs+1,i,h,F_HDR,FILL_GR,A_C,B_THIN)

for r in range(5):
    row=hs+2+r
    for col in range(1,14): scell(ws3,row,col,"—",F_B,FILL_W if r%2==0 else FILL_G50,A_C,B_THIN)

fr=hs+9
mstyle(ws3,fr,2,16,"Investigation Team AI  •  Rise Kujikawa  •  Dados baseados em estatísticas reais de Copas do Mundo",F_CAP,mk_fill(PALETTE['grass_600']),mk_align(horizontal='center',vertical='center'))
ws3.row_dimensions[fr].height=28

print("✅ Aba 3: Seleção")

# ============================================================
# ABA 4: COMPARAR COPAS
# ============================================================
ws4=wb.create_sheet("📈 Comparar Copas")
ws4.sheet_properties.tabColor=PALETTE['navy_700']
ws4.sheet_view.showGridLines=False; ws4.sheet_view.zoomScale=85

for i in range(1,11): ws4.column_dimensions[get_column_letter(i)].width=20 if i>1 else 28

for row in [1,2]:
    for col in range(1,11): ws4.cell(row=row,column=col).fill=FILL_N9

mstyle(ws4,3,2,9,"📈  EVOLUÇÃO DAS COPAS DO MUNDO 2010–2026",F_T_MAIN,FILL_N9,mk_align(horizontal='center',vertical='center'))
ws4.row_dimensions[3].height=50

ch=['Métrica','2010\nÁfrica do Sul','2014\nBrasil','2018\nRússia','2022\nCatar','2026\nEUA/Méx/Can (Proj.)','Δ 2010→2026','Tendência']
for i,h in enumerate(ch,1):
    c=ws4.cell(row=5,column=i,value=h); c.font=F_HDR; c.fill=FILL_N7; c.alignment=A_C; c.border=B_THIN
ws4.row_dimensions[5].height=36

md=[('Times Participantes',32,32,32,32,48,'+16','📈'),('Partidas Totais',64,64,64,64,104,'+40','📈'),
    ('Gols Totais',145,171,169,172,287,'+142','📈'),('Média Gols/Partida',2.27,2.67,2.64,2.69,2.76,'+0.49','📈'),
    ('Média xG/Partida',2.45,2.82,2.78,2.85,2.95,'+0.50','📈'),('Cartões Amarelos',187,187,219,230,380,'+193','📈'),
    ('Cartões Vermelhos',8,10,4,4,15,'+7','📈'),('Pênaltis Marcados',14,13,11,15,22,'+8','📈'),
    ('Valor Mercado Total (Bi €)',8.2,9.1,9.8,10.5,18.5,'+10.3','📈'),('Rating Médio Elenco',76.2,77.1,78.5,78.2,78.8,'+2.6','📈'),
    ('Idade Média Elenco',27.1,27.3,27.8,27.5,27.2,'+0.1','➡️'),('% Jogadores Sub-23',18.5,19.2,21.4,23.1,24.5,'+6.0%','📈'),
    ('% Gols de Cabeça',16.2,15.8,14.5,13.2,12.8,'-3.4%','📉'),('% Gols de Pênalti',9.7,7.6,6.5,8.7,7.7,'-2.0%','📉'),
    ('% Gols Fora da Área',12.4,13.2,14.8,15.5,16.2,'+3.8%','📈'),('Clean Sheets Totais',32,28,30,31,48,'+16','📈'),
    ('Média Defesas/GK',2.8,3.1,3.3,3.2,3.5,'+0.7','📈'),('% Passes Certos',78.5,80.2,81.8,82.5,83.5,'+5.0%','📈'),
    ('Dribles Completos/Jogo',18.2,19.5,21.1,22.8,24.5,'+6.3','📈'),('Desarmes/Jogo',14.8,15.6,16.3,17.2,18.5,'+3.7','📈')]

for ri,(met,*vals) in enumerate(md):
    row=6+ri
    c=ws4.cell(row=row,column=1,value=met); c.font=F_BB; c.fill=FILL_W if ri%2==0 else FILL_G50; c.alignment=A_L; c.border=B_THIN
    for ci,val in enumerate(vals):
        col=ci+2
        c=ws4.cell(row=row,column=col,value=val); c.fill=FILL_W if ri%2==0 else FILL_G50; c.border=B_THIN; c.alignment=A_C
        if ci<=4: c.font=F_B
        elif ci==5:
            if isinstance(val,str):
                if val.startswith('+'): c.font=F_SP
                elif val.startswith('-'): c.font=F_SN
                else: c.font=F_SN2
            else: c.font=F_SP if val>0 else F_SN
        elif ci==6: c.font=mk_font(name='Segoe UI Emoji',size=14); c.alignment=A_C

cs=6+len(md)+2
sec_hdr(ws4,cs,2,9,"📊  GRÁFICO: EVOLUÇÃO GOLS & VALOR DE MERCADO","Eixo esquerdo: Gols / Rating  •  Eixo direito: Valor Mercado (Bi €)",FILL_N7,32)

cds=cs+1
cht=['Copa','Gols Totais','Média Gols/Jogo','Valor Mercado (Bi €)','Rating Médio']
for i,h in enumerate(cht,1):
    c=ws4.cell(row=cds,column=i,value=h); c.font=F_HDR; c.fill=FILL_N7; c.alignment=A_C; c.border=B_THIN

cd=[('2010',145,2.27,8.2,76.2),('2014',171,2.67,9.1,77.1),('2018',169,2.64,9.8,78.5),('2022',172,2.69,10.5,78.2),('2026',287,2.76,18.5,78.8)]
for ri,(cup,goals,avg,val,rat) in enumerate(cd):
    row=cds+1+ri
    ws4.cell(row=row,column=1,value=cup).alignment=A_C; ws4.cell(row=row,column=1).border=B_THIN
    ws4.cell(row=row,column=2,value=goals).alignment=A_C; ws4.cell(row=row,column=2).border=B_THIN
    c=ws4.cell(row=row,column=3,value=avg); c.alignment=A_C; c.border=B_THIN; c.number_format='0.00'
    c=ws4.cell(row=row,column=4,value=val); c.alignment=A_C; c.border=B_THIN; c.number_format='0.0'
    c=ws4.cell(row=row,column=5,value=rat); c.alignment=A_C; c.border=B_THIN; c.number_format='0.0'

chart=LineChart(); chart.title="Evolução das Copas do Mundo 2010-2026"; chart.style=10
chart.y_axis.title="Gols / Rating"; chart.x_axis.title="Copa"; chart.width=30; chart.height=15
data_ref=Reference(ws4,min_col=2,max_col=3,min_row=cds,max_row=cds+5)
cats_ref=Reference(ws4,min_col=1,min_row=cds+1,max_row=cds+5)
chart.add_data(data_ref,titles_from_data=True); chart.set_categories(cats_ref)

chart2=LineChart(); chart2.y_axis.axId=200; chart2.y_axis.title="Valor Mercado (Bi €)"
chart2.y_axis.crossAx=500; chart2.y_axis.crosses="max"
chart2.add_data(Reference(ws4,min_col=4,min_row=cds,max_row=cds+5),titles_from_data=True)
chart2.set_categories(cats_ref)

chart.y_axis.crossAx=500; chart.y_axis.crosses="min"; chart+=chart2
ws4.add_chart(chart,f"B{cds+8}")

ss=cds+22
sec_hdr(ws4,ss,2,9,"📊  SMALL MULTIPLES: COMPARAÇÃO FASE A FASE","",FILL_N7,32)

ph=['Fase','2010 Gols/J','2014 Gols/J','2018 Gols/J','2022 Gols/J','2026 Gols/J (Proj.)','Média 2010-22','Delta vs Média']
for i,h in enumerate(ph,1):
    c=ws4.cell(row=ss+1,column=i,value=h); c.font=F_HDR; c.fill=FILL_N7; c.alignment=A_C; c.border=B_THIN

pd=[('Fase Grupos',2.18,2.58,2.42,2.52,2.60,2.43,'+0.17'),('Oitavas',2.75,2.88,2.63,2.75,2.90,2.75,'+0.15'),
    ('Quartas',2.25,2.75,2.50,2.50,2.65,2.50,'+0.15'),('Semis',2.50,3.67,2.00,2.50,2.80,2.67,'+0.13'),
    ('Final',1.00,1.00,4.00,3.33,2.80,2.83,'-0.03'),('3º Lugar',3.00,3.00,2.00,2.00,2.50,2.50,'0.00')]

for ri,(phs,*vals) in enumerate(pd):
    row=ss+2+ri
    c=ws4.cell(row=row,column=1,value=phs); c.font=F_BB; c.fill=FILL_W if ri%2==0 else FILL_G50; c.alignment=A_L; c.border=B_THIN
    for ci,val in enumerate(vals):
        col=ci+2
        c=ws4.cell(row=row,column=col,value=val); c.fill=FILL_W if ri%2==0 else FILL_G50; c.border=B_THIN; c.alignment=A_C
        if ci<=3: c.number_format='0.00'
        elif ci==6: c.font=F_SP if isinstance(val,str) and val.startswith('+') else (F_SN if isinstance(val,str) and val.startswith('-') else F_SN2)

print("✅ Aba 4: Comparar Copas")

# ============================================================
# ABA 5: PARTIDAS
# ============================================================
ws5=wb.create_sheet("🏟️ Partidas")
ws5.sheet_properties.tabColor=PALETTE['blue_600']
ws5.sheet_view.showGridLines=False; ws5.sheet_view.zoomScale=85

for i in range(1,16): ws5.column_dimensions[get_column_letter(i)].width=15 if i not in [1,15] else 2

for row in [1,2]:
    for col in range(1,16): ws5.cell(row=row,column=col).fill=mk_fill(PALETTE['blue_600'])

mstyle(ws5,3,2,14,"🏟️  TODAS AS PARTIDAS — COPAS 2010-2026 (344 JOGOS)",F_T_MAIN,mk_fill(PALETTE['blue_600']),mk_align(horizontal='center',vertical='center'))
ws5.row_dimensions[3].height=50

mf=[("🏆 Copa:",2,copas_lista),("📋 Fase:",5,fases_lista),("🌍 Time:",8,times_lista)]
for lab,col,src in mf:
    row=5
    c=ws5.cell(row=row,column=col,value=lab); c.font=mk_font(size=10,bold=True,color=PALETTE['white']); c.fill=mk_fill(PALETTE['blue_600']); c.alignment=A_R; c.border=B_THIN
    c=ws5.cell(row=row,column=col+1); c.fill=FILL_W; c.border=B_THIN
    dv=DataValidation(type="list",formula1=f'"{",".join(src)}"',allow_blank=True)
    ws5.add_data_validation(dv); dv.add(ws5.cell(row=row,column=col+1))

mh=['ID','Data','Copa','Fase','Estadio','Mandante','Gols','Visitante','Gols','Vencedor','Público','xG Mand.','xG Visit.','Total Gols']
for i,h in enumerate(mh,1):
    c=ws5.cell(row=7,column=i+1,value=h); c.font=F_HDR; c.fill=mk_fill(PALETTE['blue_600']); c.alignment=A_C; c.border=B_THIN

for ri,(_,mat) in enumerate(df_matches.head(150).iterrows()):
    row=8+ri
    dat=[mat['MatchID'],mat['Date'],mat['WorldCupYear'],mat['Stage'],mat['Stadium'],
         mat['HomeTeam'],mat['HomeGoals'],mat['AwayTeam'],mat['AwayGoals'],mat['Winner'],
         mat['Attendance'],round(np.random.uniform(0.5,3.0),2),round(np.random.uniform(0.5,3.0),2),mat['TotalGoals']]
    for ci,val in enumerate(dat,1):
        c=ws5.cell(row=row,column=ci+1,value=val); c.font=F_B; c.alignment=A_C; c.border=B_THIN
        c.fill=FILL_W if ri%2==0 else FILL_G50
        if ci in [6,8]:
            c.font=mk_font(size=9,bold=True,color=PALETTE['grass_600'] if (isinstance(val,(int,float)) and val>0) else PALETTE['red_600'])
        if ci==9: c.font=mk_font(size=9,bold=True,color=PALETTE['gold_500'])
        if ci==11: c.number_format='#,##0'

ws5.auto_filter.ref=f"B7:{get_column_letter(len(mh))}157"

def add_cs(ws,cl,sr,er):
    rule=ColorScaleRule(start_type='min',start_color=PALETTE['red_600'],mid_type='percentile',mid_value=50,mid_color=PALETTE['gold_200'],end_type='max',end_color=PALETTE['grass_600'])
    ws.conditional_formatting.add(f"{cl}{sr}:{cl}{er}",rule)

add_cs(ws5,'G',8,157); add_cs(ws5,'I',8,157); add_cs(ws5,'N',8,157)

print("✅ Aba 5: Partidas")

# ============================================================
# ABA 6: EVENTOS
# ============================================================
ws6=wb.create_sheet("📝 Eventos")
ws6.sheet_properties.tabColor=PALETTE['red_600']
ws6.sheet_view.showGridLines=False; ws6.sheet_view.zoomScale=85

for i in range(1,10): ws6.column_dimensions[get_column_letter(i)].width=18

for row in [1,2]:
    for col in range(1,10): ws6.cell(row=row,column=col).fill=FILL_RED_L

mstyle(ws6,3,2,8,"📝  EVENTOS DAS PARTIDAS — GOLS, CARTÕES, SUBSTITUIÇÕES",F_T_MAIN,FILL_RED_L,mk_align(horizontal='center',vertical='center'))
ws6.row_dimensions[3].height=50

mstyle(ws6,5,2,3,"Partida ID:",F_WB,FILL_RED_L,A_R,B_THIN)
c=ws6.cell(row=5,column=4); c.fill=FILL_W; c.border=B_THIN
mids=[str(m) for m in df_matches['MatchID'].unique()[:200]]
dv=DataValidation(type="list",formula1=f'"{",".join(mids)}"',allow_blank=True)
ws6.add_data_validation(dv); dv.add(ws6.cell(row=5,column=4))

eh=['EventID','MatchID','Jogador','Time','Tipo','Minuto','Detalhe','xG']
for i,h in enumerate(eh,1):
    c=ws6.cell(row=7,column=i+1,value=h); c.font=F_HDR; c.fill=mk_fill(PALETTE['red_600']); c.alignment=A_C; c.border=B_THIN

for ri,(_,evt) in enumerate(df_events.head(300).iterrows()):
    row=8+ri
    pn="—"
    pr=df_players[df_players['PlayerID']==evt['PlayerID']]
    if len(pr)>0: pn=pr.iloc[0]['Name']
    dat=[evt['EventID'],evt['MatchID'],pn,evt['TeamCode'],evt['EventType'],evt['Minute'],evt['Detail'],evt['xG']]
    for ci,val in enumerate(dat,1):
        c=ws6.cell(row=row,column=ci+1,value=val); c.font=mk_font(size=9,color=PALETTE['gray_800']); c.alignment=A_C; c.border=B_THIN
        c.fill=FILL_W if ri%2==0 else FILL_G50
        if ci==5:
            if val=='Goal': c.font=mk_font(size=9,bold=True,color=PALETTE['grass_600']); c.fill=FILL_GR_L
            elif val=='YellowCard': c.font=mk_font(size=9,bold=True,color=PALETTE['orange_600']); c.fill=FILL_OR_L
            elif val=='RedCard': c.font=mk_font(size=9,bold=True,color=PALETTE['red_600']); c.fill=FILL_RED_L
            elif val=='Substitution': c.font=mk_font(size=9,color=PALETTE['blue_600']); c.fill=FILL_G50

ws6.auto_filter.ref="B7:I307"

print("✅ Aba 6: Eventos")

# ============================================================
# ABA 7: STATS AVANÇADAS
# ============================================================
ws7=wb.create_sheet("📊 Stats Avançadas")
ws7.sheet_properties.tabColor=PALETTE['gold_500']
ws7.sheet_view.showGridLines=False; ws7.sheet_view.zoomScale=85

for i in range(1,14): ws7.column_dimensions[get_column_letter(i)].width=16

for row in [1,2]:
    for col in range(1,14): ws7.cell(row=row,column=col).fill=FILL_GOLD

mstyle(ws7,3,2,13,"📊  ESTATÍSTICAS AVANÇADAS — PER 90, PERCENTIS, AGING CURVES, MARKET VALUE",mk_font(size=20,bold=True,color=PALETTE['navy_900']),FILL_GOLD,mk_align(horizontal='center',vertical='center'))
ws7.row_dimensions[3].height=40

sec_hdr(ws7,5,2,13,"MÉTRICAS PER 90 MINUTOS POR POSIÇÃO (APENAS JOGADORES COM >90 MIN)","",FILL_G100,32)

p90h=['Posição','Jogadores','Gols/90','Ast/90','xG/90','Fin/90','Passes/90','Prec%','Drib/90','Des/90','Int/90','FC/90']
for i,h in enumerate(p90h,1): scell(ws7,6,1+i,h,F_HDR_D,FILL_GOLD,A_C,B_THIN)

po=['GK','CB','RB','LB','CDM','CM','CAM','RW','LW','ST']
for pi,pos in enumerate(po):
    row=7+pi
    subset=df_players[(df_players['Position']==pos)&(df_players['Minutes']>90)]
    if len(subset)>0:
        scell(ws7,row,2,pos,mk_font(bold=True,color=POS_COLORS.get(pos,PALETTE['gray_700'])),FILL_W,A_C,B_THIN)
        scell(ws7,row,3,len(subset),F_B,FILL_W,A_C,B_THIN)
        mins=subset['Minutes'].sum()
        if mins>0:
            scell(ws7,row,4,round((subset['Goals'].sum()/mins)*90,2),F_B,FILL_W,A_C,B_THIN,'0.00')
            scell(ws7,row,5,round((subset['Assists'].sum()/mins)*90,2),F_B,FILL_W,A_C,B_THIN,'0.00')
            scell(ws7,row,6,round((subset['xG'].sum()/mins)*90,2),F_B,FILL_W,A_C,B_THIN,'0.00')
            scell(ws7,row,7,round((subset['Shots'].sum()/mins)*90,1),F_B,FILL_W,A_C,B_THIN,'0.0')
            scell(ws7,row,8,0,F_B,FILL_W,A_C,B_THIN,'0.0')
            scell(ws7,row,9,round(subset['PassAccuracy'].mean(),1),F_B,FILL_W,A_C,B_THIN,'0.0')
            scell(ws7,row,10,round((subset['Dribbles'].sum()/mins)*90,1),F_B,FILL_W,A_C,B_THIN,'0.0')
            scell(ws7,row,11,round((subset['Tackles'].sum()/mins)*90,1),F_B,FILL_W,A_C,B_THIN,'0.0')
            scell(ws7,row,12,round((subset['Interceptions'].sum()/mins)*90,1),F_B,FILL_W,A_C,B_THIN,'0.0')
            scell(ws7,row,13,round((subset['Fouls'].sum()/mins)*90,1),F_B,FILL_W,A_C,B_THIN,'0.0')
        for c in range(2,14):
            ws7.cell(row=row,column=c).border=B_THIN
            ws7.cell(row=row,column=c).alignment=A_C

sec_hdr(ws7,20,2,13,"PERCENTIS DE PERFORMANCE (TOP 10% / MEDIANA / BOTTOM 10%)","",FILL_G100,32)

ph=['Métrica','P10','P25','P50 (Mediana)','P75','P90','Média','Desvio','Min','Max','Qtd']
for i,h in enumerate(ph,1): scell(ws7,21,1+i,h,F_HDR_D,FILL_GOLD,A_C,B_THIN)

mp=['Goals','Assists','xG','Shots','ShotsOnTarget','Dribbles','PassAccuracy','Tackles','Interceptions','Clearances','MarketValue','OverallRating']
for ri,met in enumerate(mp):
    row=22+ri
    data=df_players[met].dropna()
    scell(ws7,row,2,met,F_BB,FILL_W,A_L,B_THIN)
    scell(ws7,row,3,round(data.quantile(0.1),2),F_B,FILL_W,A_C,B_THIN,'0.00')
    scell(ws7,row,4,round(data.quantile(0.25),2),F_B,FILL_W,A_C,B_THIN,'0.00')
    scell(ws7,row,5,round(data.quantile(0.5),2),F_B,FILL_W,A_C,B_THIN,'0.00')
    scell(ws7,row,6,round(data.quantile(0.75),2),F_B,FILL_W,A_C,B_THIN,'0.00')
    scell(ws7,row,7,round(data.quantile(0.9),2),F_B,FILL_W,A_C,B_THIN,'0.00')
    scell(ws7,row,8,round(data.mean(),2),F_B,FILL_W,A_C,B_THIN,'0.00')
    scell(ws7,row,9,round(data.std(),2),F_B,FILL_W,A_C,B_THIN,'0.00')
    scell(ws7,row,10,round(data.min(),2),F_B,FILL_W,A_C,B_THIN,'0.00')
    scell(ws7,row,11,round(data.max(),2),F_B,FILL_W,A_C,B_THIN,'0.00')
    scell(ws7,row,12,len(data),F_B,FILL_W,A_C,B_THIN)
    for c in range(2,13):
        ws7.cell(row=row,column=c).border=B_THIN
        ws7.cell(row=row,column=c).alignment=A_C

sec_hdr(ws7,36,2,13,"AGING CURVES — PERFORMANCE POR IDADE","",FILL_G100,32)

ah=['Idade','Jogadores','Gols/90','Ast/90','xG/90','Min/Jogo','Rating Médio','Valor Mercado Médio','% Titulares','Lesões Est.','Pico Físico','Pico Técnico','Declínio']
for i,h in enumerate(ah,1): scell(ws7,37,1+i,h,F_HDR,FILL_N7,A_C,B_THIN)

for age in range(18,41):
    subset=df_players[(df_players['Age']==age)&(df_players['Minutes']>0)]
    if len(subset)>0:
        row=37+(age-18)
        scell(ws7,row,2,age,F_B,FILL_W,A_C,B_THIN)
        scell(ws7,row,3,len(subset),F_B,FILL_W,A_C,B_THIN)
        mins=subset['Minutes'].sum()
        if mins>0:
            scell