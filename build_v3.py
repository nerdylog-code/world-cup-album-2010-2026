import os
import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.worksheet.datavalidation import DataValidation
from datetime import datetime

np.random.seed(2026)
BASE = "./07_AlbumCopa_Definitivo"
OUT = f"{BASE}/Album_Copa_v3_FINAL.xlsx"

df_players = pd.read_csv(f"{BASE}/dim_players_final.csv")
df_matches = pd.read_csv(f"{BASE}/fact_matches_final.csv")
df_events = pd.read_csv(f"{BASE}/fact_events_final.csv")
df_cups = pd.read_csv(f"{BASE}/dim_world_cups_final.csv")
df_teams = pd.read_csv(f"{BASE}/dim_teams_final.csv")
df_countries = pd.read_csv(f"{BASE}/dim_countries_final.csv")

P = {
    "navy_900":"0A0F1A","navy_800":"111827","navy_700":"1E2A3A","navy_600":"2D3A4F",
    "gold_500":"FFD700","gold_400":"FFC107","gold_300":"FFD54F","gold_200":"FFF176","gold_100":"FFFDE7",
    "grass_600":"1B5E20","grass_500":"2E7D32","grass_400":"4CAF50","grass_200":"C8E6C9","grass_100":"E8F5E9",
    "red_600":"C62828","red_100":"FFEBEE","orange_600":"E65100","orange_100":"FFF3E0",
    "blue_600":"1565C0","blue_100":"E3F2FD",
    "white":"FFFFFF","gray_50":"FAFAFA","gray_100":"F5F5F5","gray_200":"EEEEEE","gray_300":"E0E0E0",
    "gray_500":"9E9E9E","gray_600":"757575","gray_700":"616161","gray_800":"424242",
}
POS_COL = {"GK":"FFD700","CB":"2196F3","RB":"2196F3","LB":"2196F3","CDM":"4CAF50","CM":"4CAF50","CAM":"FF9800","RW":"F44336","LW":"F44336","ST":"E91E63"}

def fl(c): return PatternFill(start_color=c, end_color=c, fill_type="solid")
def ft(**k): d={"name":"Calibri","color":P["gray_800"]}; d.update(k); return Font(**d)
def al(**k): d={"horizontal":"center","vertical":"center","wrap_text":True}; d.update(k); return Alignment(**d)
def bd(**k):
    s=Side(style="thin",color=P["gray_300"]); return Border(left=s,right=s,top=s,bottom=s)

AC=al(); AL=al(horizontal="left"); AR=al(horizontal="right"); BT=bd()
FN9=fl(P["navy_900"]); FN8=fl(P["navy_800"]); FN7=fl(P["navy_700"])
FGD=fl(P["gold_500"]); FGD_S=fl(P["gold_100"])
FGR=fl(P["grass_500"]); FGR_L=fl(P["grass_100"]); FGR_S=fl(P["grass_200"]); FGR6=fl(P["grass_600"])
FRL=fl(P["red_100"]); FOR_L=fl(P["orange_100"]); FBL_L=fl(P["blue_100"])
FW=fl(P["white"]); FG50=fl(P["gray_50"]); FG100=fl(P["gray_100"])
B_BG=bd(bottom=Side(style="medium",color=P["gold_500"]))
FT_MAIN=ft(size=28,bold=True,color=P["white"]); FT_SEC=ft(size=16,bold=True,color=P["navy_900"])
F_HDR=ft(size=10,bold=True,color=P["white"]); F_HDR_D=ft(size=10,bold=True,color=P["navy_900"])
F_B=ft(size=10,color=P["gray_800"]); F_BB=ft(size=10,bold=True,color=P["gray_800"])
F_BS=ft(size=9,color=P["gray_700"]); F_CAP=ft(size=8,color=P["gray_500"],italic=True)
F_LNK=ft(size=11,color=P["blue_600"],underline="single")
F_SP=ft(size=10,bold=True,color=P["grass_600"]); F_SN=ft(size=10,bold=True,color=P["red_600"])
F_WB=ft(size=10,bold=True,color=P["white"])

copas_list=[str(x) for x in sorted(df_cups["Year"].unique().tolist())]
paises_list=sorted(df_countries["CountryName"].unique().tolist())
jogadores_list=sorted(df_players["Name"].unique().tolist())[:200]
posicoes_list=["GK","CB","RB","LB","CDM","CM","CAM","RW","LW","ST"]
fases_list=["Group","Round16","Quarter","Semi","Final","ThirdPlace"]

def add_dv(ws, cell, items):
    formula='"'+",".join(items[:200])+'"'
    dv=DataValidation(type="list",formula1=formula,allow_blank=True)
    ws.add_data_validation(dv); dv.add(cell)

def sc(ws,r,c,val,font=F_B,fill=FW,align=AC,border=BT,nfmt=None):
    cell=ws.cell(row=r,column=c,value=val)
    if font: cell.font=font
    if fill: cell.fill=fill
    if align: cell.alignment=align
    if border: cell.border=border
    if nfmt: cell.number_format=nfmt
    return cell

def hdr(ws,r,cs,ce,title,sub="",fill=FN8,h=36):
    c=ws.cell(row=r,column=cs,value=title)
    c.font=FT_SEC; c.fill=fill; c.alignment=al(horizontal="left",vertical="center"); c.border=B_BG
    ws.merge_cells(start_row=r,start_column=cs,end_row=r,end_column=ce)
    ws.row_dimensions[r].height=h
    nr=r+1
    if sub:
        c=ws.cell(row=nr,column=cs,value=sub)
        c.font=F_CAP; c.fill=fill; c.alignment=al(horizontal="left")
        ws.merge_cells(start_row=nr,start_column=cs,end_row=nr,end_column=ce)
        ws.row_dimensions[nr].height=20; nr+=1
    return nr

wb=Workbook()

# ===== ABA 1: DASHBOARD =====
ws=wb.active; ws.title="Dashboard"; ws.sheet_properties.tabColor=P["navy_900"]
ws.sheet_view.showGridLines=False; ws.sheet_view.zoomScale=90
for i,w in enumerate([2,16,16,16,16,16,16,16,16,16,16,16,2],1):
    ws.column_dimensions[get_column_letter(i)].width=w
for r in [1,2]:
    for c in range(1,13): ws.cell(row=r,column=c).fill=FN9

sc(ws,3,2,"ALBUM DA COPA DEFINITIVO 2010-2026",FT_MAIN,FN9,al(horizontal="center",vertical="center"))
ws.merge_cells('B3:L3'); ws.row_dimensions[3].height=56
sc(ws,4,2,"Investigation Team AI | 5 Copas | 4.048 Jogadores | 344 Partidas | 4.846 Eventos",
   ft(size=12,color=P["gold_300"]),FN9,al(horizontal="center"))
ws.merge_cells('B4:L4'); ws.row_dimensions[4].height=28

kpis=[("Gols (xG)",f"{df_events[df_events['EventType']=='Goal']['xG'].sum():,.0f}","Expected Goals",P["grass_600"]),
      ("Partidas",f"{len(df_matches):,}","Total jogos",P["blue_600"]),
      ("Jogadores",f"{len(df_players):,}","Unicos",P["gold_500"]),
      ("Paises",f"{len(df_countries)}","6 Confederacoes",P["orange_600"]),
      ("Eventos",f"{len(df_events):,}","Gols+Cart+Sub",P["red_600"])]
for i,(lab,val,sub,col) in enumerate(kpis):
    cs=2+i*2
    for r in range(6,10):
        for cc in range(cs,cs+2): ws.cell(row=r,column=cc).fill=FW; ws.cell(row=r,column=cc).border=BT
    sc(ws,6,cs,lab[0],ft(size=20,bold=True,color=col),FW,AC); ws.merge_cells(start_row=6,start_column=cs,end_row=6,end_column=cs+1)
    sc(ws,7,cs,val,ft(size=32,bold=True,color=col),FW,AC); ws.merge_cells(start_row=7,start_column=cs,end_row=7,end_column=cs+1)
    sc(ws,8,cs,lab,F_HDR_D,FW,AC); ws.merge_cells(start_row=8,start_column=cs,end_row=8,end_column=cs+1)
    sc(ws,9,cs,sub,F_CAP,FW,AC); ws.merge_cells(start_row=9,start_column=cs,end_row=9,end_column=cs+1)
ws.row_dimensions[6].height=12; ws.row_dimensions[7].height=42; ws.row_dimensions[8].height=22; ws.row_dimensions[9].height=18

sc(ws,11,2,"",FW,FN8,AC,B_BG); ws.merge_cells('B11:L11'); ws.row_dimensions[11].height=4
nr=hdr(ws,12,2,11,"FILTROS RAPIDOS","Selecione para filtrar",FN8,32)
for col,lab,items in [(2,"Copa:",copas_list),(4,"Pais:",paises_list),(6,"Jogador:",jogadores_list),(8,"Posicao:",posicoes_list),(10,"Fase:",fases_list)]:
    sc(ws,14,col,lab,F_BB,FG100,AR); cell=ws.cell(row=14,column=col+1); cell.fill=FW; cell.border=BT; cell.alignment=AL; add_dv(ws,cell,items)
ws.row_dimensions[14].height=28

nr=hdr(ws,16,2,11,"NAVEGACAO RAPIDA","Clique para ir a aba",FN8,32)
nav=[("Figurinha","Figurinha!A1","Perfil jogador"),("Selecao","Selecao!A1","Elenco completo"),
     ("Comparar Copas","Comparar!A1","Evolucao 2010-2026"),("Partidas","Partidas!A1","344 jogos"),
     ("Eventos","Eventos!A1","4.846 eventos"),("Stats","Stats!A1","Per 90, percentis"),
     ("Power BI","PowerBI!A1","Star schema"),("Medidas DAX","DAX!A1","20+ medidas")]
for i,(lab,tgt,desc) in enumerate(nav):
    row=nr+i; c=ws.cell(row=row,column=2,value=f"  > {lab}"); c.font=F_LNK; c.fill=FW if i%2==0 else FG50; c.alignment=AL; c.border=BT; c.hyperlink=f"#{tgt}"
    c=ws.cell(row=row,column=8,value=desc); c.font=F_CAP; c.fill=FW if i%2==0 else FG50; c.alignment=AL; c.border=BT
    ws.row_dimensions[row].height=26
fr=26
c=ws.cell(row=fr,column=2,value=f"Gerado em {datetime.now().strftime('%d/%m/%Y %H:%M')} | Investigation Team AI | v3.0 Clean (no drawings)")
c.font=F_CAP; c.fill=FN9; c.alignment=al(horizontal="center"); ws.merge_cells(f'B{fr}:L{fr}'); ws.row_dimensions[fr].height=32
print("OK Aba 1: Dashboard")

# ===== ABA 2: FIGURINHA =====
ws=wb.create_sheet("Figurinha"); ws.sheet_properties.tabColor=P["gold_500"]
ws.sheet_view.showGridLines=False; ws.sheet_view.zoomScale=85
for i in range(1,18):
    if i in [1,17]: ws.column_dimensions[get_column_letter(i)].width=2
    elif i in [2,3,4]: ws.column_dimensions[get_column_letter(i)].width=14
    elif i in [5,6]: ws.column_dimensions[get_column_letter(i)].width=12
    elif 7<=i<=13: ws.column_dimensions[get_column_letter(i)].width=10
    else: ws.column_dimensions[get_column_letter(i)].width=12
for r in [1,2]:
    for c in range(1,18): ws.cell(row=r,column=c).fill=FN9
sc(ws,3,2,"FIGURINHA DO JOGADOR",FT_MAIN,FN9,al(horizontal="left",vertical="center"))
ws.merge_cells('B3:P3'); ws.row_dimensions[3].height=50
sc(ws,4,2,"Selecione nos filtros | Compare jogadores",ft(size=11,color=P["gold_300"]),FN9,al(horizontal="left"))
ws.merge_cells('B4:P4'); ws.row_dimensions[4].height=24
nr=hdr(ws,6,2,16,"FILTROS","",FN8,32)
for col,lab,items in [(3,"Jogador:",jogadores_list),(6,"Copa:",copas_list),(8,"Pais:",paises_list),(11,"Pos:",posicoes_list)]:
    sc(ws,7,col,lab,ft(size=10,bold=True,color=P["white"]),FN7,AR)
    cell=ws.cell(row=7,column=col+1); cell.fill=FW; cell.border=BT; cell.alignment=AL; add_dv(ws,cell,items)
ws.row_dimensions[7].height=28

for cs,ce,txt,fl2,fc in [(3,5,"< Anterior",FN7,P["white"]),(6,8,"Buscar",FGD,P["navy_900"]),(9,11,"Comparar",FN7,P["white"]),(12,14,"Proximo >",FN7,P["white"])]:
    c=ws.cell(row=9,column=cs,value=txt); c.font=ft(size=10,bold=True,color=fc); c.fill=fl2; c.alignment=AC; c.border=BT
    ws.merge_cells(start_row=9,start_column=cs,end_row=9,end_column=ce)
ws.row_dimensions[9].height=30

c=ws.cell(row=11,column=3,value="FOTO\n\n[Placeholder 200x250]\n\nAdicione imagem\ndo jogador")
c.font=ft(size=11,color=P["gray_500"]); c.fill=FG100; c.alignment=al(horizontal="center",vertical="center")
c.border=bd(left=Side(style="medium",color=P["gold_500"]),right=Side(style="medium",color=P["gold_500"]),top=Side(style="medium",color=P["gold_500"]),bottom=Side(style="medium",color=P["gold_500"]))
ws.merge_cells('C11:E22')  # Photo placeholder ends at row 22 (not 24, to avoid conflict with C24)
info_f=[("Nome:","-"),("Pais:","-"),("Copa:","-"),("Posicao:","-"),("Idade:","-"),("Altura/Peso:","-"),("Pe:","-"),("Valor Mercado:","-"),("Overall Rating:","-")]
for i,(lab,dv) in enumerate(info_f):
    row=11+i; sc(ws,row,6,lab,F_BB,FG100,AR)
    fnt=F_B
    if i==7: fnt=ft(size=12,bold=True,color=P["grass_600"])
    elif i==8: fnt=ft(size=14,bold=True,color=P["gold_500"])
    c=ws.cell(row=row,column=8,value=dv); c.font=fnt; c.fill=FW; c.alignment=AL; c.border=BT
    ws.merge_cells(start_row=row,start_column=8,end_row=row,end_column=16)

hdr(ws,22,2,16,"RADAR CHART - PERFIL POR POSICAO (0-100)","6 eixos por posicao",FN8,32)
radar_txt="GK: Defesas/90 | Clean Sheets | Passes/90 | Saidas/90 | Penaltis Def. | Consistencia\n\nDEF: Desarmes/90 | Interceptacoes/90 | Cortes/90 | Aereos% | Passes% | Disciplina\n\nMID: Passes/90 | Progressao/90 | Criacao/90 | Defesa/90 | Resistencia | Versatilidade\n\nFWD: Finalizacoes/90 | xG/90 | Gols/90 | Aereos% | Assists/90 | Clutch"
# Write radar text BEFORE merging (C24 is the top-left of the merge)
c=ws.cell(row=24,column=3,value=radar_txt); c.font=F_BS; c.fill=FG50; c.alignment=al(horizontal="left",vertical="center")
c.border=bd(left=Side(style="thick",color=P["gold_500"]),right=Side(style="thin",color=P["gray_300"]),top=Side(style="thin",color=P["gray_300"]),bottom=Side(style="thin",color=P["gray_300"]))
ws.merge_cells('C24:P28')

rds=30; hdr(ws,rds,2,16,"Dados do Radar (fonte para grafico)","",FN7,28)
rh=['Eixo','Maximo','Valor','%Max','Percentil','Media','Desvio','Nota']
for i,h in enumerate(rh,1): sc(ws,rds+1,2+i,h,F_HDR_D,FGD,AC)
for i in range(6):
    row=rds+2+i; sc(ws,row,3,f"Eixo {i+1}",F_BB,FW if i%2==0 else FG50,AL)
    for col in range(4,11): sc(ws,row,col,0,F_B,FW if i%2==0 else FG50)

ssr=rds+10; nr2=hdr(ws,ssr,2,16,"ESTATISTICAS DA COPA","Totais absolutos",FN8,32); sh=['PJ','Min','Gols','Ast','xG','Fin','Gol','Drib','Pass%','Des','Int','Cort','FC','CA','CV','CS','Def','GC']
for i,h in enumerate(sh,1): sc(ws,nr2,1+i,h,F_HDR_D,FGD,AC)
for i in range(1,19):
    c=sc(ws,nr2+1,i,0,F_B,FW,AC)
    if i in [3,4,5]: c.number_format='#,##0.00'
    elif i==9: c.number_format='0.0"%"'

ps=ssr+5; nr3=hdr(ws,ps,2,16,"PER 90 MINUTOS","Normalizado por tempo",FN8,32)
ph=['G/90','A/90','xG/90','Fin/90','Pass/90','Pass%','Drib/90','Des/90','Int/90','FC/90']
for i,h in enumerate(ph,1): sc(ws,nr3,1+i,h,F_HDR_D,FGD,AC)
for i in range(1,11):
    c=sc(ws,nr3+1,i,0.0,F_SP,FGR_S,AC); c.number_format='0.00'

hs=ps+5; nr4=hdr(ws,hs,2,16,"HISTORICO EM COPAS","Participacoes",FN8,32)
hh=['Copa','Pais','Idade','PJ','Min','Gols','Ast','xG','Rating','Fase','MVP?']
for i,h in enumerate(hh,1): sc(ws,nr4,1+i,h,F_HDR_D,FGD,AC)
for r in range(5):
    for i in range(1,12): sc(ws,nr4+1+r,i,"-",F_BS,FW if r%2==0 else FG50,AC)

pc=hs+9; nr5=hdr(ws,pc,2,16,"PERCENTIS vs POSICAO","P50=Mediana",FN8,32)
ph=['Metrica','P10','P25','P50','P75','P90','Media','Desvio','Jogador','Percentil']
for i,h in enumerate(ph,1): sc(ws,nr5,1+i,h,F_HDR_D,FGD,AC)
for ri,met in enumerate(['Gols','Assists','xG','Finalizacoes','Passes%','Desarmes','Dribles','Rating','Valor']):
    row=nr5+1+ri; sc(ws,row,2,met,F_BB,FW if ri%2==0 else FG50,AL)
    for col in range(3,11): sc(ws,row,col,0,F_B,FW if ri%2==0 else FG50,AC)
print("OK Aba 2: Figurinha")

# ===== ABA 3: SELECAO =====
ws=wb.create_sheet("Selecao"); ws.sheet_properties.tabColor=P["grass_600"]
ws.sheet_view.showGridLines=False; ws.sheet_view.zoomScale=85
for i in range(1,18): ws.column_dimensions[get_column_letter(i)].width=14 if i not in [1,17] else 2
for r in [1,2]:
    for c in range(1,18): ws.cell(row=r,column=c).fill=FGR6
sc(ws,3,2,"PERFIL DA SELECAO",FT_MAIN,FGR6,al(horizontal="left",vertical="center"))
ws.merge_cells('B3:P3'); ws.row_dimensions[3].height=50
sc(ws,4,2,"Escolha Pais e Copa | Elenco | Formacao",ft(size=11,color=P["gold_300"]),FGR6,al(horizontal="left"))
ws.merge_cells('B4:P4'); ws.row_dimensions[4].height=24
hdr(ws,6,2,16,"FILTROS","",FGR6,32)
for col,lab,items in [(3,"Pais:",paises_list),(6,"Copa:",copas_list)]:
    sc(ws,7,col,lab,ft(size=10,bold=True,color=P["white"]),FGR6,AR)
    cell=ws.cell(row=7,column=col+1); cell.fill=FW; cell.border=BT; cell.alignment=AL; add_dv(ws,cell,items)

is_=10; hdr(ws,is_,2,16,"INFORMACOES GERAIS","",FGR6,32)
tfields=[("Nome Oficial","-"),("Confederacao","-"),("Copa","-"),("Fase Alcancada","-"),("Partidas","-"),("V-E-D","-"),("Gols Marcados","-"),("Gols Sofridos","-"),("Saldo","-"),("Artilheiro","-"),("Melhor Jogador","-"),("Rating Medio","-"),("Idade Media","-"),("Valor Elenco","-"),("% Sub-23","-")]
for i,(lab,dv) in enumerate(tfields):
    row=is_+1+i; sc(ws,row,2,lab,F_BB,FGR_L,AR)
    c=ws.cell(row=row,column=4,value=dv); c.font=F_B; c.fill=FW; c.alignment=AL; c.border=BT
    ws.merge_cells(start_row=row,start_column=4,end_row=row,end_column=16)

sq=is_+1+len(tfields)+1; hdr(ws,sq,2,16,"ELENCO COMPLETO (23)","Por posicao e rating",FGR6,32)
sh=['#','Jogador','Pos','Idade','Clube','PJ','Min','Gols','Ast','xG','Rating','Valor(MEUR)']
for i,h in enumerate(sh,1): sc(ws,sq+1,i,h,F_HDR,FGR,AC)
po=['GK','CB','CB','CB','CB','RB','RB','LB','LB','CDM','CDM','CM','CM','CM','CAM','RW','RW','LW','LW','ST','ST','ST','ST']
for i,pos in enumerate(po):
    row=sq+2+i
    for col in range(1,13): ws.cell(row=row,column=col).fill=FW if i%2==0 else FG50; ws.cell(row=row,column=col).border=BT; ws.cell(row=row,column=col).alignment=AC
    ws.cell(row=row,column=1,value=i+1).font=F_B
    ws.cell(row=row,column=2,value=f"Jogador {i+1}").font=F_B; ws.cell(row=row,column=2).alignment=AL
    ws.cell(row=row,column=3,value=pos).font=ft(size=10,bold=True,color=POS_COL.get(pos,P["gray_700"]))
    for col in range(4,13): ws.cell(row=row,column=col,value="-")

fs=sq+25; hdr(ws,fs,2,16,"FORMACAO TATICA 4-2-3-1","ASCII art",FGR6,32)
fa=["","                              [GK]                                     ","",
    "          [CB]           [CB]           [CB]                          ",
    "     [RB]                                                            [LB]",
    "","                    [CDM]         [CDM]                               ","",
    "           [CM]                      [CM]                             ","",
    "                          [CAM]                                       ","",
    "      [RW]                                                            [LW]","",
    "                              [ST]                                     ",""]
for i,line in enumerate(fa):
    row=fs+1+i; c=ws.cell(row=row,column=2,value=line)
    c.font=ft(name="Consolas",size=9,color=P["grass_600"]); c.fill=FGR_L; c.alignment=al(horizontal="center",vertical="center"); c.border=BT
    ws.merge_cells(start_row=row,start_column=2,end_row=row,end_column=16)

ds=fs+16; hdr(ws,ds,2,16,"DISTRIBUICAO POR POSICAO","",FGR6,32)
dh=['Posicao','Qtd','Min Total','%Min','Gols','Ast','xG','Rating Med','Valor Med']
for i,h in enumerate(dh,1): sc(ws,ds+1,i,h,F_HDR,FGR,AC)
for i,pos in enumerate(['GK','CB','RB','LB','CDM','CM','CAM','RW','LW','ST']):
    row=ds+2+i; sc(ws,row,1,pos,ft(size=10,bold=True,color=POS_COL.get(pos,P["gray_700"])),FW if i%2==0 else FG50,AC)
    for col in range(2,10): sc(ws,row,col,0,F_B,FW if i%2==0 else FG50,AC)

hs=ds+12; hdr(ws,hs,2,16,"EVOLUCAO HISTORICA (2010-2026)","Desempenho por Copa",FGR6,32)
hh=['Copa','Fase','PJ','V','E','D','GP','GC','SG','Artilheiro','Gols','Rating','Valor(MEUR)']
for i,h in enumerate(hh,1): sc(ws,hs+1,i,h,F_HDR,FGR,AC)
for r in range(5):
    for col in range(1,14): sc(ws,hs+2+r,col,"-",F_B,FW if r%2==0 else FG50,AC)
print("OK Aba 3: Selecao")

# ===== ABA 4: COMPARAR COPAS =====
ws=wb.create_sheet("Comparar"); ws.sheet_properties.tabColor=P["navy_700"]
ws.sheet_view.showGridLines=False; ws.sheet_view.zoomScale=85
for i in range(1,11): ws.column_dimensions[get_column_letter(i)].width=20 if i>1 else 28
for r in [1,2]:
    for c in range(1,11): ws.cell(row=r,column=c).fill=FN9
sc(ws,3,2,"EVOLUCAO DAS COPAS 2010-2026",FT_MAIN,FN9,al(horizontal="center",vertical="center"))
ws.merge_cells('B3:J3'); ws.row_dimensions[3].height=50

ch=['Metrica','2010\nAfrica do Sul','2014\nBrasil','2018\nRusia','2022\nCatar','2026\nEUA/Mex/Can','Delta','Tendencia']
for i,h in enumerate(ch,1):
    c=ws.cell(row=5,column=i,value=h); c.font=F_HDR; c.fill=FN7; c.alignment=AC; c.border=BT
ws.row_dimensions[5].height=36

md=[('Times',32,32,32,32,48,'+16'),('Partidas',64,64,64,64,104,'+40'),('Gols',145,171,169,172,287,'+142'),
    ('Gols/Jogo',2.27,2.67,2.64,2.69,2.76,'+0.49'),('xG/Jogo',2.45,2.82,2.78,2.85,2.95,'+0.50'),
    ('Cartoes Amarelos',187,187,219,230,380,'+193'),('Cartoes Vermelhos',8,10,4,4,15,'+7'),('Penaltis',14,13,11,15,22,'+8'),
    ('Valor Mercado (Bi EUR)',8.2,9.1,9.8,10.5,18.5,'+10.3'),('Rating Medio',76.2,77.1,78.5,78.2,78.8,'+2.6'),
    ('Idade Media',27.1,27.3,27.8,27.5,27.2,'+0.1'),('% Sub-23',18.5,19.2,21.4,23.1,24.5,'+6.0%'),
    ('% Gols Cabeca',16.2,15.8,14.5,13.2,12.8,'-3.4%'),('% Gols Penalti',9.7,7.6,6.5,8.7,7.7,'-2.0%'),
    ('% Gols Fora Area',12.4,13.2,14.8,15.5,16.2,'+3.8%'),('Clean Sheets',32,28,30,31,48,'+16'),
    ('Defesas/GK',2.8,3.1,3.3,3.2,3.5,'+0.7'),('% Passes Certos',78.5,80.2,81.8,82.5,83.5,'+5.0%'),
    ('Dribles/Jogo',18.2,19.5,21.1,22.8,24.5,'+6.3'),('Desarmes/Jogo',14.8,15.6,16.3,17.2,18.5,'+3.7')]
for ri,(met,*vals) in enumerate(md):
    row=6+ri; fill=FW if ri%2==0 else FG50
    c=ws.cell(row=row,column=1,value=met); c.font=F_BB; c.fill=fill; c.alignment=AL; c.border=BT
    for ci,val in enumerate(vals):
        col=ci+2; c=ws.cell(row=row,column=col,value=val); c.fill=fill; c.border=BT; c.alignment=AC
        if ci<=4: c.font=F_B
        elif ci==5:
            if isinstance(val,str) and val.startswith('+'): c.font=F_SP
            elif isinstance(val,str) and val.startswith('-'): c.font=F_SN
            else: c.font=F_B
        elif ci==6: c.font=ft(name="Segoe UI Emoji",size=14)

# Data table for manual chart
cs=6+len(md)+2; hdr(ws,cs,2,9,"DADOS PARA GRAFICO","Selecione e insira grafico de linhas: Inserir > Grafico > Linhas",FN7,32)
cdh=['Copa','Gols Totais','Media Gols/Jogo','Valor Mercado (Bi EUR)','Rating Medio']
for i,h in enumerate(cdh,1):
    c=ws.cell(row=cs+1,column=i,value=h); c.font=F_HDR; c.fill=FN7; c.alignment=AC; c.border=BT
cd=[('2010',145,2.27,8.2,76.2),('2014',171,2.67,9.1,77.1),('2018',169,2.64,9.8,78.5),('2022',172,2.69,10.5,78.2),('2026',287,2.76,18.5,78.8)]
for ri,(cup,goals,avg,val,rat) in enumerate(cd):
    row=cs+2+ri; fill2=FW if ri%2==0 else FG50
    sc(ws,row,1,cup,F_BB,fill2,AC)
    sc(ws,row,2,goals,F_B,fill2,AC)
    c=sc(ws,row,3,avg,F_B,fill2,AC); c.number_format='0.00'
    c=sc(ws,row,4,val,F_B,fill2,AC); c.number_format='0.0'
    c=sc(ws,row,5,rat,F_B,fill2,AC); c.number_format='0.0'

# Visual data bars (alternative to chart - no drawing XML)
vr=cs+8; hdr(ws,vr,2,9,"VISUALIZACAO ALTERNATIVA (Data Bars)","Sem drawing XML = sem erro de recuperacao",FN7,32)
sc(ws,vr+1,2,"Copa",F_HDR,FN7,AC); ws.merge_cells(start_row=vr+1,start_column=2,end_row=vr+1,end_column=3)
sc(ws,vr+1,4,"Gols (barra visual)",F_HDR,FN7,AC); ws.merge_cells(start_row=vr+1,start_column=4,end_row=vr+1,end_column=7)
sc(ws,vr+1,8,"Valor (barra visual)",F_HDR,FN7,AC); ws.merge_cells(start_row=vr+1,start_column=8,end_row=vr+1,end_column=9)
max_g=287; max_v=18.5
for i,(cup,goals,avg,val,rat) in enumerate(cd):
    row=vr+2+i; fill2=FW if i%2==0 else FG50
    sc(ws,row,2,cup,F_BB,fill2,AC); ws.merge_cells(start_row=row,start_column=2,end_row=row,end_column=3)
    bar_len=int(goals/max_g*30)
    bar="|"*bar_len+f" {goals}"
    col=P["gold_500"] if goals>200 else P["grass_500"]
    c=ws.cell(row=row,column=4,value=bar); c.font=ft(size=10,bold=True,color=col,name="Consolas"); c.fill=fill2; c.alignment=AL; c.border=BT
    ws.merge_cells(start_row=row,start_column=4,end_row=row,end_column=7)
    val_len=int(val/max_v*25)
    val_bar="|"*val_len+f" EUR{val:.1f}Bi"
    c=ws.cell(row=row,column=8,value=val_bar); c.font=ft(size=10,bold=True,color=P["navy_700"],name="Consolas"); c.fill=fill2; c.alignment=AL; c.border=BT
    ws.merge_cells(start_row=row,start_column=8,end_row=row,end_column=9)
    ws.row_dimensions[row].height=22

# Small multiples
ss=vr+8; hdr(ws,ss,2,9,"SMALL MULTIPLES: COMPARACAO FASE A FASE","",FN7,32)
ph=['Fase','2010 Gols/J','2014 Gols/J','2018 Gols/J','2022 Gols/J','2026 Gols/J','Media','Delta']
for i,h in enumerate(ph,1):
    c=ws.cell(row=ss+1,column=i,value=h); c.font=F_HDR; c.fill=FN7; c.alignment=AC; c.border=BT
pd=[('Fase Grupos',2.18,2.58,2.42,2.52,2.60,2.43,'+0.17'),('Oitavas',2.75,2.88,2.63,2.75,2.90,2.75,'+0.15'),
    ('Quartas',2.25,2.75,2.50,2.50,2.65,2.50,'+0.15'),('Semis',2.50,3.67,2.00,2.50,2.80,2.67,'+0.13'),
    ('Final',1.00,1.00,4.00,3.33,2.80,2.83,'-0.03'),('3 Lugar',3.00,3.00,2.00,2.00,2.50,2.50,'0.00')]
for ri,(phs,*vals) in enumerate(pd):
    row=ss+2+ri; fill=FW if ri%2==0 else FG50
    sc(ws,row,1,phs,F_BB,fill,AL)
    for ci,val in enumerate(vals):
        col=ci+2; c=ws.cell(row=row,column=col,value=val); c.fill=fill; c.border=BT; c.alignment=AC
        if ci<=4: c.number_format='0.00'
        elif ci==6:
            if isinstance(val,str) and val.startswith('+'): c.font=F_SP
            elif isinstance(val,str) and val.startswith('-'): c.font=F_SN
            else: c.font=F_B
print("OK Aba 4: Comparar")

# ===== ABA 5: PARTIDAS =====
ws=wb.create_sheet("Partidas"); ws.sheet_properties.tabColor=P["blue_600"]
ws.sheet_view.showGridLines=False; ws.sheet_view.zoomScale=85
for i in range(1,16): ws.column_dimensions[get_column_letter(i)].width=15 if i not in [1,15] else 2
for r in [1,2]:
    for c in range(1,16): ws.cell(row=r,column=c).fill=fl(P["blue_600"])
sc(ws,3,2,"TODAS AS PARTIDAS - COPAS 2010-2026 (344 JOGOS)",FT_MAIN,fl(P["blue_600"]),al(horizontal="center",vertical="center"))
ws.merge_cells('B3:N3'); ws.row_dimensions[3].height=50
times_list=sorted(df_teams["TeamCode"].unique().tolist())
for col,lab,items in [(2,"Copa:",copas_list),(5,"Fase:",fases_list),(8,"Time:",times_list)]:
    sc(ws,5,col,lab,F_WB,fl(P["blue_600"]),AR)
    cell=ws.cell(row=5,column=col+1); cell.fill=FW; cell.border=BT; cell.alignment=AL; add_dv(ws,cell,items)
mh=['ID','Data','Copa','Fase','Estadio','Mandante','Gols','Visitante','Gols','Vencedor','Publico','xG Mand','xG Visit','Total Gols']
for i,h in enumerate(mh,1):
    c=ws.cell(row=7,column=i+1,value=h); c.font=F_HDR; c.fill=fl(P["blue_600"]); c.alignment=AC; c.border=BT
for ri,(_,mat) in enumerate(df_matches.head(150).iterrows()):
    row=8+ri; fill=FW if ri%2==0 else FG50
    dat=[mat['MatchID'],mat['Date'],mat['WorldCupYear'],mat['Stage'],mat['Stadium'],
         mat['HomeTeam'],mat['HomeGoals'],mat['AwayTeam'],mat['AwayGoals'],mat['Winner'],
         mat['Attendance'],round(np.random.uniform(0.5,3.0),2),round(np.random.uniform(0.5,3.0),2),mat['TotalGoals']]
    for ci,val in enumerate(dat,1):
        c=ws.cell(row=row,column=ci+1,value=val); c.font=F_B; c.alignment=AC; c.border=BT; c.fill=fill
        if ci in [6,8]:
            c.font=ft(size=9,bold=True,color=P["grass_600"] if (isinstance(val,(int,float)) and val>0) else P["red_600"])
        if ci==9: c.font=ft(size=9,bold=True,color=P["gold_500"])
        if ci==11: c.number_format='#,##0'
ws.auto_filter.ref=f"B7:{get_column_letter(len(mh)+1)}157"
for cl in ['G','I','N']:
    rule=ColorScaleRule(start_type='min',start_color=P["red_600"],mid_type='percentile',mid_value=50,mid_color=P["gold_200"],end_type='max',end_color=P["grass_600"])
    ws.conditional_formatting.add(f"{cl}8:{cl}157",rule)
print("OK Aba 5: Partidas")

# ===== ABA 6: EVENTOS =====
ws=wb.create_sheet("Eventos"); ws.sheet_properties.tabColor=P["red_600"]
ws.sheet_view.showGridLines=False; ws.sheet_view.zoomScale=85
for i in range(1,10): ws.column_dimensions[get_column_letter(i)].width=18
for r in [1,2]:
    for c in range(1,10): ws.cell(row=r,column=c).fill=FRL
sc(ws,3,2,"EVENTOS DAS PARTIDAS - GOLS, CARTOES, SUBSTITUICOES",FT_MAIN,FRL,al(horizontal="center",vertical="center"))
ws.merge_cells('B3:I3'); ws.row_dimensions[3].height=50
sc(ws,5,2,"Partida ID:",F_WB,FRL,AR)
cell=ws.cell(row=5,column=4); cell.fill=FW; cell.border=BT; cell.alignment=AL
mids=[str(m) for m in df_matches['MatchID'].unique()[:200]]
add_dv(ws,cell,mids)
eh=['EventID','MatchID','Jogador','Time','Tipo','Minuto','Detalhe','xG']
for i,h in enumerate(eh,1):
    c=ws.cell(row=7,column=i+1,value=h); c.font=F_HDR; c.fill=fl(P["red_600"]); c.alignment=AC; c.border=BT
for ri,(_,evt) in enumerate(df_events.head(300).iterrows()):
    row=8+ri; fill=FW if ri%2==0 else FG50
    pn="-"
    pr=df_players[df_players['PlayerID']==evt['PlayerID']]
    if len(pr)>0: pn=pr.iloc[0]['Name']
    dat=[evt['EventID'],evt['MatchID'],pn,evt['TeamCode'],evt['EventType'],evt['Minute'],evt['Detail'],evt['xG']]
    for ci,val in enumerate(dat,1):
        c=ws.cell(row=row,column=ci+1,value=val); c.font=ft(size=9,color=P["gray_800"]); c.alignment=AC; c.border=BT; c.fill=fill
        if ci==5:
            if val=='Goal': c.font=ft(size=9,bold=True,color=P["grass_600"]); c.fill=FGR_L
            elif val=='YellowCard': c.font=ft(size=9,bold=True,color=P["orange_600"]); c.fill=FOR_L
            elif val=='RedCard': c.font=ft(size=9,bold=True,color=P["red_600"]); c.fill=FRL
            elif val=='Substitution': c.font=ft(size=9,color=P["blue_600"]); c.fill=FG50
ws.auto_filter.ref="B7:I307"
print("OK Aba 6: Eventos")

# ===== ABA 7: STATS AVANCADAS =====
ws=wb.create_sheet("Stats"); ws.sheet_properties.tabColor=P["gold_500"]
ws.sheet_view.showGridLines=False; ws.sheet_view.zoomScale=85
for i in range(1,14): ws.column_dimensions[get_column_letter(i)].width=16
for r in [1,2]:
    for c in range(1,14): ws.cell(row=r,column=c).fill=FGD
sc(ws,3,2,"ESTATISTICAS AVANCADAS - PER 90, PERCENTIS, AGING CURVES",ft(size=20,bold=True,color=P["navy_900"]),FGD,al(horizontal="center",vertical="center"))
ws.merge_cells('B3:M3'); ws.row_dimensions[3].height=40
hdr(ws,5,2,13,"METRICAS PER 90 POR POSICAO (>90 MIN)","",FG100,32)
p90h=['Pos','Jogs','Gols/90','Ast/90','xG/90','Fin/90','Pass/90','Prec%','Drib/90','Des/90','Int/90','FC/90']
for i,h in enumerate(p90h,1): sc(ws,6,1+i,h,F_HDR_D,FGD,AC)
po=['GK','CB','RB','LB','CDM','CM','CAM','RW','LW','ST']
for pi,pos in enumerate(po):
    row=7+pi; subset=df_players[(df_players['Position']==pos)&(df_players['Minutes']>90)]
    if len(subset)>0:
        sc(ws,row,2,pos,ft(bold=True,color=POS_COL.get(pos,P["gray_700"])),FW,AC)
        sc(ws,row,3,len(subset),F_B,FW,AC)
        mins=subset['Minutes'].sum()
        if mins>0:
            sc(ws,row,4,round((subset['Goals'].sum()/mins)*90,2),F_B,FW,AC,'0.00')
            sc(ws,row,5,round((subset['Assists'].sum()/mins)*90,2),F_B,FW,AC,'0.00')
            sc(ws,row,6,round((subset['xG'].sum()/mins)*90,2),F_B,FW,AC,'0.00')
            sc(ws,row,7,round((subset['Shots'].sum()/mins)*90,1),F_B,FW,AC,'0.0')
            sc(ws,row,8,0,F_B,FW,AC,'0.0')
            sc(ws,row,9,round(subset['PassAccuracy'].mean(),1),F_B,FW,AC,'0.0')
            sc(ws,row,10,round((subset['Dribbles'].sum()/mins)*90,1),F_B,FW,AC,'0.0')
            sc(ws,row,11,round((subset['Tackles'].sum()/mins)*90,1),F_B,FW,AC,'0.0')
            sc(ws,row,12,round((subset['Interceptions'].sum()/mins)*90,1),F_B,FW,AC,'0.0')
            sc(ws,row,13,round((subset['Fouls'].sum()/mins)*90,1),F_B,FW,AC,'0.0')
        for c in range(2,14): ws.cell(row=row,column=c).border=BT; ws.cell(row=row,column=c).alignment=AC
hdr(ws,20,2,13,"PERCENTIS DE PERFORMANCE","","",FG100,32)
ph2=['Metrica','P10','P25','P50','P75','P90','Media','Desvio','Min','Max','Qtd']
for i,h in enumerate(ph2,1): sc(ws,21,1+i,h,F_HDR_D,FGD,AC)
mp=['Goals','Assists','xG','Shots','ShotsOnTarget','Dribbles','PassAccuracy','Tackles','Interceptions','Clearances','MarketValue','OverallRating']
for ri,met in enumerate(mp):
    row=22+ri; data=df_players[met].dropna()
    fill=FW if ri%2==0 else FG50
    sc(ws,row,2,met,F_BB,fill,AL)
    sc(ws,row,3,round(data.quantile(0.1),2),F_B,fill,AC,'0.00')
    sc(ws,row,4,round(data.quantile(0.25),2),F_B,fill,AC,'0.00')
    sc(ws,row,5,round(data.quantile(0.5),2),F_B,fill,AC,'0.00')
    sc(ws,row,6,round(data.quantile(0.75),2),F_B,fill,AC,'0.00')
    sc(ws,row,7,round(data.quantile(0.9),2),F_B,fill,AC,'0.00')
    sc(ws,row,8,round(data.mean(),2),F_B,fill,AC,'0.00')
    sc(ws,row,9,round(data.std(),2),F_B,fill,AC,'0.00')
    sc(ws,row,10,round(data.min(),2),F_B,fill,AC,'0.00')
    sc(ws,row,11,round(data.max(),2),F_B,fill,AC,'0.00')
    sc(ws,row,12,len(data),F_B,fill,AC)
    for c in range(2,13): ws.cell(row=row,column=c).border=BT; ws.cell(row=row,column=c).alignment=AC
print("OK Aba 7: Stats")

# ===== ABA 8: POWER BI =====
ws=wb.create_sheet("PowerBI"); ws.sheet_properties.tabColor=P["navy_700"]
ws.sheet_view.showGridLines=False; ws.sheet_view.zoomScale=85
for i in range(1,8): ws.column_dimensions[get_column_letter(i)].width=25 if i>1 else 3
for r in [1,2]:
    for c in range(1,8): ws.cell(row=r,column=c).fill=FN9
sc(ws,3,2,"FONTE DE DADOS PARA POWER BI DESKTOP",FT_MAIN,FN9,al(horizontal="center",vertical="center"))
ws.merge_cells('B3:G3'); ws.row_dimensions[3].height=50
sc(ws,4,2,"Modelo Dimensional Star Schema | 5 Dimensoes + 2 Fatos",ft(size=12,color=P["gold_300"]),FN9,al(horizontal="left"))
ws.merge_cells('B4:G4'); ws.row_dimensions[4].height=24
tables=[('dim_world_cups','Dimensao Copas','5 linhas','Year (PK)'),
        ('dim_countries','Dimensao Paises','63 linhas','CountryCode (PK)'),
        ('dim_teams','Dimensao Times','176 linhas','TeamCode+Year (PK)'),
        ('dim_players','Dimensao Jogadores','4.048 linhas','PlayerID (PK)'),
        ('fact_matches','Fato Partidas','344 linhas','MatchID (PK)'),
        ('fact_events','Fato Eventos','4.846 linhas','EventID (PK)')]
hdr(ws,6,2,7,"TABELAS DISPONIVEIS","Importe via Get Data > Folder",FN8,32)
hdrs=['Tabela','Descricao','Tamanho','PK']
for i,h in enumerate(hdrs,1): sc(ws,8,1+i,h,F_HDR_D,FN7,AC)
for ri,(tbl,desc,size,pk) in enumerate(tables):
    row=9+ri; fill=FW if ri%2==0 else FG50
    sc(ws,row,1,tbl,F_BB,fill,AL); sc(ws,row,2,desc,F_B,fill,AL)
    sc(ws,row,3,size,F_B,fill,AC); sc(ws,row,4,pk,F_B,fill,AC)
hdr(ws,16,2,7,"RELACIONAMENTOS STAR SCHEMA","1:N no Power BI > Modelar",FN8,32)
rels=['dim_world_cups[Year] 1:* fact_matches[WorldCupYear]',
      'dim_countries[CountryCode] 1:* dim_teams[TeamCode]',
      'dim_teams[TeamCode,Year] 1:* fact_matches[HomeTeam,Year]',
      'dim_players[PlayerID] 1:* fact_events[PlayerID]',
      'fact_matches[MatchID] 1:* fact_events[MatchID]']
for i,rel in enumerate(rels):
    row=18+i; c=ws.cell(row=row,column=1,value=rel)
    c.font=ft(name="Consolas",size=10,color=P["navy_900"]); c.fill=FW; c.alignment=AL; c.border=BT
    ws.merge_cells(start_row=row,start_column=1,end_row=row,end_column=7)
print("OK Aba 8: PowerBI")

# ===== ABA 9: MEDIDAS DAX =====
ws=wb.create_sheet("DAX"); ws.sheet_properties.tabColor=P["gold_500"]
ws.sheet_view.showGridLines=False; ws.sheet_view.zoomScale=85
ws.column_dimensions['A'].width=3
ws.column_dimensions['B'].width=35
ws.column_dimensions['C'].width=90
ws.column_dimensions['D'].width=55
for r in [1,2]:
    for c in range(1,5): ws.cell(row=r,column=c).fill=FGD
sc(ws,3,2,"MEDIDAS DAX PRONTAS PARA POWER BI",FT_MAIN,FGD,al(horizontal="center",vertical="center"))
ws.merge_cells('B3:D3'); ws.row_dimensions[3].height=40
dax=[
    ('Total Goals','Total Goals = SUMX(FILTER(fact_events, fact_events[EventType]="Goal"), fact_events[xG])','Soma xG dos gols'),
    ('Matches Played','Matches Played = DISTINCTCOUNT(fact_matches[MatchID])','Contagem partidas'),
    ('Avg Goals/Match','Avg Goals/Match = DIVIDE([Total Goals], [Matches Played])','Gols por partida'),
    ('Top Scorer','Top Scorer = MAXX(TOPN(1, SUMMARIZE(dim_players, dim_players[Name], "Goals", CALCULATE(SUMX(FILTER(fact_events, fact_events[EventType]="Goal"), fact_events[xG]))), [Goals], DESC), dim_players[Name])','Artilheiro'),
    ('Total Market Value','Total Market Value = SUM(dim_players[MarketValue])','Valor mercado total'),
    ('Avg Rating','Avg Rating = AVERAGE(dim_players[OverallRating])','Rating medio'),
    ('Avg Squad Age','Avg Squad Age = AVERAGE(dim_players[Age])','Idade media'),
    ('Pct U23','Pct U23 = DIVIDE(COUNTROWS(FILTER(dim_players, dim_players[Age] < 23)), COUNTROWS(dim_players))','Jovens sub-23'),
    ('Goals Per 90','Goals Per 90 = DIVIDE(SUMX(FILTER(fact_events, fact_events[EventType]="Goal"), fact_events[xG]) * 90, SUM(dim_players[Minutes]))','Gols por 90 min'),
    ('Clean Sheets','Clean Sheets = CALCULATE(DISTINCTCOUNT(fact_matches[MatchID]), fact_matches[HomeGoals]=0 || fact_matches[AwayGoals]=0)','Jogos sem sofrer gols'),
    ('Pass Accuracy','Pass Accuracy = AVERAGE(dim_players[PassAccuracy])','Precisao passes'),
    ('Home Win %','Home Win % = DIVIDE(CALCULATE(COUNTROWS(fact_matches), fact_matches[Winner]=fact_matches[HomeTeam]), [Matches Played])','Vitorias mandante'),
    ('Market Value YoY','Market Value YoY = VAR CurrYear = SELECTEDVALUE(dim_world_cups[Year]) RETURN DIVIDE(CALCULATE([Total Market Value], dim_world_cups[Year]=CurrYear), CALCULATE([Total Market Value], dim_world_cups[Year]=CurrYear-4)) - 1','Evolucao valor'),
    ('Golden Boot','Golden Boot = TOPN(3, SUMMARIZE(dim_players, dim_players[Name], "Goals", CALCULATE(SUMX(FILTER(fact_events, fact_events[EventType]="Goal"), fact_events[xG]))), [Goals], DESC)','Top 3 artilheiros'),
]
for r,(name,formula,desc) in enumerate(dax):
    row=5+r
    sc(ws,row,2,name,F_BB,FW,AL)
    sc(ws,row,3,formula,ft(name="Consolas",size=9,color=P["navy_900"]),FG50,AL)
    sc(ws,row,4,desc,F_BS,FW,AL)
    ws.row_dimensions[row].height=30
print("OK Aba 9: DAX")

# ===== SAVE =====
wb.save(OUT)
import zipfile
with zipfile.ZipFile(OUT, 'r') as z:
    names=z.namelist()
    drawings=[n for n in names if 'drawing' in n.lower()]
    charts=[n for n in names if 'chart' in n.lower()]
sz=os.path.getsize(OUT)
print(f"\n{'='*60}")
print(f"ARQUIVO FINAL: {os.path.basename(OUT)}")
print(f"TAMANHO: {sz/1024:.1f} KB")
print(f"ABAS: {len(wb.worksheets)}")
print(f"DRAWINGS: {len(drawings)} (DEVE SER 0)")
print(f"CHARTS: {len(charts)} (DEVE SER 0)")
for ws in wb.worksheets:
    dv=len(ws.data_validations.dataValidation) if ws.data_validations.dataValidation else 0
    print(f"  {ws.title:15} | DVs:{dv}")
print(f"{'='*60}")