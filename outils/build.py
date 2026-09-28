# -*- coding: utf-8 -*-
import json, datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, NamedStyle
from openpyxl.utils import get_column_letter as CL
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.properties import PageSetupProperties
from openpyxl.comments import Comment

D=json.load(open('rows.json',encoding='utf-8'))
M=json.load(open('model.json',encoding='utf-8'))
rows=D['rows']; WEEKS=D['weeks']; CYCLES=D['cycles']; CLASS_ORDER=D['class_order']
DAYS=['Lundi','Mardi','Mercredi','Jeudi','Vendredi','Samedi','Dimanche']
CYC={c:cy for cy,lst in CYCLES for c in lst}

# ============ PALETTE (sobre, professionnelle) ============
INK      = '1B2A41'   # bleu nuit  - bandeaux principaux
INK2     = '2E4763'   # bleu ardoise - sous-bandeaux
ACC      = '2F6F7E'   # sarcelle - accent
GREY_H   = 'E8EDF2'   # gris bleuté - en-têtes secondaires
GREY_L   = 'F5F7FA'   # fond très clair
WHITE    = 'FFFFFF'
LINE     = 'B9C4D0'   # bordures
LINE_L   = 'DDE4EB'
CYC_FILL = {'Primaire':'EAF2F8','Collège':'E8F3EF','Secondaire':'FBF3E6','Bac':'EEEBF5'}
CYC_BAND = {'Primaire':'3C6E91','Collège':'2F7A63','Secondaire':'9A6B1F','Bac':'5B4E8C'}
WEEKEND  = 'FDFAF4'   # colonnes samedi/dimanche (séances de jour)
ALERT_BG = 'FDEDEC'; ALERT_FG='B03A2E'
OK_FG    = '1E7B4D'
QUINZ_FG = '8A5A00'

# --- Familles de matières : 8 couleurs pastel, texte foncé lisible sur chacune ---
FAMILLES=[('Arabe',              'FBE3DC', ['Arabe','عربية']),
          ('Français',           'DCE8F7', ['Français']),
          ('Langues étrangères', 'DFF0E4', ['Anglais','Espagnol','Italien']),
          ('Mathématiques',      'FDF0D5', ['Math','رياضيات']),
          ('Sciences',           'E2E6F5', ['Physique','فيزياء','SVT','Éveil scientifique','إيقاظ علمي']),
          ('Sciences humaines',  'F3E1EE', ['Histoire','Géographie','Philo']),
          ('Économie / Gestion', 'FAE6CF', ['Eco','Gestion']),
          ('Informatique',       'DCEFF2', ['Informatique','ALGO','STI'])]
FAM_OF={m:(n,c) for n,c,lst in FAMILLES for m in lst}

F='Arial'
def font(sz=10,b=False,c='1B2A41',i=False): return Font(name=F,size=sz,bold=b,color=c,italic=i)
def fill(c): return PatternFill('solid',fgColor=c)
def thin(c=LINE_L): return Side(style='thin',color=c)
BOX   = Border(left=thin(),right=thin(),top=thin(),bottom=thin())
BOX_M = Border(left=thin(LINE),right=thin(LINE),top=thin(LINE),bottom=thin(LINE))
CTR   = Alignment(horizontal='center',vertical='center',wrap_text=True)
LFT   = Alignment(horizontal='left',vertical='center',wrap_text=True)
LTOP  = Alignment(horizontal='left',vertical='top',wrap_text=True)

wb = Workbook()
wb.calculation.fullCalcOnLoad = True
wb.remove(wb.active)

def setup_print(ws, orient='landscape', paper=9, fitw=1, fith=0, titles=None, area=None,
                margins=(0.35,0.35,0.5,0.45), header=None, footer=None, grid=False, center=True):
    ws.page_setup.orientation = orient
    ws.page_setup.paperSize = paper            # 9 = A4
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.fitToWidth = fitw
    ws.page_setup.fitToHeight = fith
    ws.page_margins.left, ws.page_margins.right = margins[0], margins[1]
    ws.page_margins.top, ws.page_margins.bottom = margins[2], margins[3]
    ws.page_margins.header, ws.page_margins.footer = 0.2, 0.2
    if titles: ws.print_title_rows = titles
    if area: ws.print_area = area
    ws.print_options.horizontalCentered = center
    ws.print_options.gridLines = grid
    ws.oddHeader.right.text = header or ''
    ws.oddHeader.right.size, ws.oddHeader.right.font = 8, 'Arial,Italic'
    ws.oddFooter.left.text = footer or 'RS 2026-2027'
    ws.oddFooter.left.size = 8
    ws.oddFooter.right.text = 'Page &P / &N'
    ws.oddFooter.right.size = 8
    ws.sheet_view.showGridLines = False

def title_block(ws, last_col, t1, t2, h1=34, h2=20):
    ws.merge_cells(f'A1:{last_col}1'); ws.merge_cells(f'A2:{last_col}2')
    c=ws['A1']; c.value=t1; c.font=font(15,True,WHITE); c.fill=fill(INK); c.alignment=Alignment(horizontal='left',vertical='center',indent=1)
    c=ws['A2']; c.value=t2; c.font=font(9,False,INK2); c.fill=fill(GREY_H); c.alignment=Alignment(horizontal='left',vertical='center',indent=1)
    for col in range(1,ws.max_column+1):
        ws.cell(1,col).fill=fill(INK); ws.cell(2,col).fill=fill(GREY_H)
    ws.row_dimensions[1].height=h1; ws.row_dimensions[2].height=h2

# ==========================================================
# 1) PLANNING  (table maître = source de vérité)
# ==========================================================
P = wb.create_sheet('Planning')
HDR = ['Semaine','Période','Parité','Date','Jour','Début','Fin','Durée (h)','Cycle','Classe / Groupe',
       'Matière','Matière (normalisée)','Enseignant','Salle','Mode','Rythme','Origine','Notes',
       '⚠ Conflit enseignant','⚠ Conflit classe','⚠ Doublon','Rang','Clé','Libellé grille',
       'Texte d\'origine','Réf. cellule source']
WID = [9,30,7,11,11,7.5,7.5,9,12,18,20,19,17,10,12,15,19,30,19,17,12,7,34,44,52,26]
NB = len(rows); R0 = 4; R1 = R0+NB-1
title_block(P, CL(len(HDR)),
  'PLANNING DES SÉANCES — Année scolaire 2026-2027',
  "TABLE DE RÉFÉRENCE : toute modification se fait ici. Les grilles hebdomadaires et tous les indicateurs se recalculent automatiquement. "
  "Colonnes S à Z = zone technique (contrôles + formules), ne pas saisir.")
for j,(h,w) in enumerate(zip(HDR,WID),1):
    c=P.cell(3,j,h); c.font=font(9,True,WHITE); c.fill=fill(INK2 if j<19 else (ACC if j<22 else '6B7A8C'))
    c.alignment=CTR; c.border=BOX_M
    P.column_dimensions[CL(j)].width=w
P.row_dimensions[3].height=32

for i,r in enumerate(rows):
    x=R0+i
    P.cell(x,1,r['semaine']).alignment=CTR
    P.cell(x,2,r['periode']).alignment=LFT
    P.cell(x,3,r['parite']).alignment=CTR
    c=P.cell(x,4,datetime.date.fromisoformat(r['date'])); c.number_format='ddd dd/mm/yyyy'; c.alignment=CTR
    P.cell(x,5,r['jour']).alignment=CTR
    for col,val in ((6,r['debut']),(7,r['fin'])):
        h,mi=val.split(':'); c=P.cell(x,col,datetime.time(int(h),int(mi))); c.number_format='hh:mm'; c.alignment=CTR
    c=P.cell(x,8,f'=IF(OR($F{x}="",$G{x}=""),"",($G{x}-$F{x})*24)'); c.number_format='0.00'; c.alignment=CTR
    P.cell(x,9,r['cycle']).alignment=CTR
    P.cell(x,10,r['classe']).alignment=LFT
    P.cell(x,11,r['matiere']).alignment=LFT
    P.cell(x,12,r['matiere_n']).alignment=LFT
    P.cell(x,13,r['prof']).alignment=LFT
    P.cell(x,14,None).alignment=CTR          # Salle : absente de la source
    P.cell(x,15,None).alignment=CTR          # Mode  : absent de la source
    P.cell(x,16,r['rythme']).alignment=CTR
    P.cell(x,17,r['origine']).alignment=LFT
    P.cell(x,18,r['notes']).alignment=LFT
    # --- contrôles vivants ---
    P.cell(x,19,f'=IF($M{x}="","",IF(SUMPRODUCT(($D${R0}:$D${R1}=$D{x})*($M${R0}:$M${R1}=$M{x})*($F${R0}:$F${R1}<$G{x})*($G${R0}:$G${R1}>$F{x}))>1,"⚠ CONFLIT",""))').alignment=CTR
    P.cell(x,20,f'=IF($J{x}="","",IF(SUMPRODUCT(($D${R0}:$D${R1}=$D{x})*($J${R0}:$J${R1}=$J{x})*($F${R0}:$F${R1}<$G{x})*($G${R0}:$G${R1}>$F{x}))>1,"⚠ CONFLIT",""))').alignment=CTR
    P.cell(x,21,f'=IF(COUNTIFS($D${R0}:$D${R1},$D{x},$J${R0}:$J${R1},$J{x},$F${R0}:$F${R1},$F{x},$K${R0}:$K${R1},$K{x})>1,"⚠ DOUBLON","")').alignment=CTR
    P.cell(x,22,f'=COUNTIFS($A${R0}:$A{x},$A{x},$J${R0}:$J{x},$J{x},$E${R0}:$E{x},$E{x})').alignment=CTR
    P.cell(x,23,f'=$A{x}&"|"&$J{x}&"|"&$E{x}&"|"&$V{x}').alignment=LFT
    # 1 ligne = horaire (+ repère quinzaine) / 2 = matière / 3 = enseignant / puis salle, mode, note
    P.cell(x,24,(f'=TEXT($F{x},"hh:mm")&" - "&TEXT($G{x},"hh:mm")'
                 f'&IF($P{x}="Quinzaine A","   ◆ Q-A",IF($P{x}="Quinzaine B","   ◆ Q-B",IF($P{x}="Exception","   ◆ PONCTUEL","")))'
                 f'&CHAR(10)&$K{x}&CHAR(10)&$M{x}'
                 f'&IF($N{x}<>"",CHAR(10)&"Salle "&$N{x},"")&IF($O{x}="En ligne",CHAR(10)&"◆ EN LIGNE","")'
                 f'&IF($R{x}<>"",CHAR(10)&$R{x},"")')).alignment=LTOP
    P.cell(x,25,r['brut']).alignment=LFT
    P.cell(x,26,r['src']).alignment=LFT
    for j in range(1,len(HDR)+1):
        cc=P.cell(x,j); cc.font=font(9); cc.border=BOX
        if j>=19: cc.font=font(8,c='6B7A8C')
    P.row_dimensions[x].height=15

P.freeze_panes='F4'
P.auto_filter.ref=f'A3:{CL(len(HDR))}{R1}'
rng=f'A{R0}:R{R1}'
for cy,f_ in CYC_FILL.items():
    P.conditional_formatting.add(rng, FormulaRule(formula=[f'$I{R0}="{cy}"'], fill=fill(f_), stopIfTrue=False))
for _nom,_col,_mats in FAMILLES:   # repère couleur de la famille de matières
    _test='+'.join(f'($K{R0}="{m}")' for m in _mats)
    P.conditional_formatting.add(f'K{R0}:L{R1}', FormulaRule(formula=[f'({_test})>0'], fill=fill(_col), stopIfTrue=True))
P.conditional_formatting.add(f'P{R0}:P{R1}', FormulaRule(formula=[f'LEFT($P{R0},9)="Quinzaine"'], font=font(9,True,QUINZ_FG)))
P.conditional_formatting.add(f'P{R0}:P{R1}', FormulaRule(formula=[f'$P{R0}="Exception"'], font=font(9,True,ALERT_FG)))
P.conditional_formatting.add(f'Q{R0}:Q{R1}', FormulaRule(formula=[f'$Q{R0}="Généré (octobre)"'], font=font(9,False,ACC)))
for col in ('S','T','U'):
    P.conditional_formatting.add(f'{col}{R0}:{col}{R1}', FormulaRule(formula=[f'LEN(TRIM(${col}{R0}))>0'], fill=fill(ALERT_BG), font=font(8,True,ALERT_FG)))
setup_print(P, titles='1:3', area=f'A1:R{R1}',
            header='Planning des séances — RS 2026-2027', footer='RS 2026-2027 · Planning (table de référence)')
P.sheet_properties.tabColor=INK

# ==========================================================
# 2) RÉFÉRENTIELS
# ==========================================================
Rf = wb.create_sheet('Référentiels')
matieres = sorted({r['matiere'] for r in rows})
profs    = sorted({r['prof'] for r in rows if r['prof']})
modes    = ['Présentiel','En ligne','Hybride']
rythmes  = ['Hebdomadaire','Quinzaine A','Quinzaine B','Exception','Ponctuelle / modifiée']
Rf.merge_cells('A1:T1'); Rf.merge_cells('A2:T2')
title_block(Rf,'T','RÉFÉRENTIELS — listes de référence du classeur',
  "Ces listes alimentent les menus déroulants de la feuille Planning. Ajouter une valeur ici la rend immédiatement sélectionnable. "
  "La colonne Salle est volontairement vide : cette information est absente du fichier d'origine.")
blocks=[('CLASSES / GROUPES',1,['Ordre','Classe / Groupe','Cycle','Séances sem. A','Séances sem. B'],
         [[i+1,c,CYC[c],None,None] for i,c in enumerate(CLASS_ORDER)],[7,18,12,13,13]),
        ('MATIÈRES',7,['Matière (libellé affiché)','Normalisée','Séances (total)'],
         [[m_,{'عربية':'Arabe','رياضيات':'Math','إيقاظ علمي':'Éveil scientifique','فيزياء':'Physique'}.get(m_,m_),None] for m_ in matieres],[22,19,14]),
        ('ENSEIGNANTS',11,['Enseignant','Séances / sem. A','Heures / sem. A'],
         [[p,None,None] for p in profs],[18,15,15]),
        ('SALLES',15,['Salle (à compléter)'],[[None] for _ in range(12)],[18]),
        ('MODES',17,['Mode'],[[x] for x in modes],[14]),
        ('RYTHMES',19,['Rythme'],[[x] for x in rythmes],[21])]
for name,c0,hdrs,data,widths in blocks:
    Rf.cell(3,c0,name).font=font(9,True,WHITE)
    lastc=c0+len(hdrs)-1
    if lastc>c0: Rf.merge_cells(start_row=3,start_column=c0,end_row=3,end_column=lastc)
    for k in range(c0,lastc+1): Rf.cell(3,k).fill=fill(ACC); Rf.cell(3,k).alignment=CTR; Rf.cell(3,k).border=BOX_M
    for k,h in enumerate(hdrs):
        c=Rf.cell(4,c0+k,h); c.font=font(9,True); c.fill=fill(GREY_H); c.alignment=CTR; c.border=BOX_M
        Rf.column_dimensions[CL(c0+k)].width=widths[k]
    for i,d_ in enumerate(data):
        for k,v in enumerate(d_):
            c=Rf.cell(5+i,c0+k,v); c.font=font(9); c.border=BOX; c.alignment=CTR if k else LFT
Rf.row_dimensions[3].height=20; Rf.row_dimensions[4].height=28
# formules de comptage
for i,c_ in enumerate(CLASS_ORDER):
    x=5+i
    Rf.cell(x,4,f'=COUNTIFS(Planning!$C${R0}:$C${R1},"A",Planning!$A${R0}:$A${R1},"S5",Planning!$J${R0}:$J${R1},$B{x})')
    Rf.cell(x,5,f'=COUNTIFS(Planning!$C${R0}:$C${R1},"B",Planning!$A${R0}:$A${R1},"S4",Planning!$J${R0}:$J${R1},$B{x})')
for i,m_ in enumerate(matieres):
    Rf.cell(5+i,9,f'=COUNTIF(Planning!$K${R0}:$K${R1},$G{5+i})')
for i,p in enumerate(profs):
    x=5+i
    Rf.cell(x,12,f'=COUNTIFS(Planning!$A${R0}:$A${R1},"S5",Planning!$M${R0}:$M${R1},$K{x})')
    Rf.cell(x,13,f'=SUMIFS(Planning!$H${R0}:$H${R1},Planning!$A${R0}:$A${R1},"S5",Planning!$M${R0}:$M${R1},$K{x})')
    Rf.cell(x,13).number_format='0.0'
# semaines
Rf.cell(3,21,'CALENDRIER DES SEMAINES').font=font(9,True,WHITE)
Rf.merge_cells(start_row=3,start_column=21,end_row=3,end_column=26)
for k in range(21,27): Rf.cell(3,k).fill=fill(ACC); Rf.cell(3,k).border=BOX_M; Rf.cell(3,k).alignment=CTR
for k,h in enumerate(['Code','Début','Fin','Parité','Période','Mois']):
    c=Rf.cell(4,21+k,h); c.font=font(9,True); c.fill=fill(GREY_H); c.alignment=CTR; c.border=BOX_M
    Rf.column_dimensions[CL(21+k)].width=[9,12,12,8,34,12][k]
for i,(code,start,par,per,shname,mois) in enumerate(WEEKS):
    x=5+i; sd=datetime.date.fromisoformat(start)
    for k,v in enumerate([code,sd,sd+datetime.timedelta(days=6),par,per,mois]):
        c=Rf.cell(x,21+k,v); c.font=font(9); c.border=BOX; c.alignment=CTR
        if k in (1,2): c.number_format='ddd dd/mm/yyyy'
Rf.freeze_panes='A5'
setup_print(Rf, area=f'A1:Z{4+len(CLASS_ORDER)}', titles='1:4', footer='RS 2026-2027 · Référentiels')
Rf.sheet_properties.tabColor='6B7A8C'

# noms définis + validations
defn = {'LST_Classes':f"Référentiels!$B$5:$B${4+len(CLASS_ORDER)}",
        'LST_Matieres':f"Référentiels!$G$5:$G${4+len(matieres)}",
        'LST_Profs':f"Référentiels!$K$5:$K${4+len(profs)}",
        'LST_Salles':"Référentiels!$O$5:$O$16",
        'LST_Modes':f"Référentiels!$Q$5:$Q${4+len(modes)}",
        'LST_Rythmes':f"Référentiels!$S$5:$S${4+len(rythmes)}"}
from openpyxl.workbook.defined_name import DefinedName
for n,ref in defn.items(): wb.defined_names.add(DefinedName(n, attr_text=ref))
for col,nm,strict in (('J','LST_Classes',False),('K','LST_Matieres',False),('M','LST_Profs',False),
                      ('N','LST_Salles',False),('O','LST_Modes',True),('P','LST_Rythmes',True)):
    dv=DataValidation(type='list',formula1=f'={nm}',allow_blank=True,
                      errorStyle='stop' if strict else 'warning',
                      error="Valeur absente du référentiel.", errorTitle='Vérification',
                      prompt=f"Choisir dans la liste ({nm}).", promptTitle='Saisie guidée')
    dv.showErrorMessage=True; dv.showInputMessage=False
    P.add_data_validation(dv); dv.add(f'{col}{R0}:{col}{R1+200}')

# ==========================================================
# 3) GRILLES HEBDOMADAIRES — une séance par ligne
# ==========================================================
from openpyxl.worksheet.pagebreak import Break
import math
CPL=21          # caractères tenant sur une ligne dans une colonne « jour »
FS_GRID=11      # taille de police des cases

_byday={}
for _r in rows:
    _byday.setdefault((_r['semaine'],_r['classe'],_r['jour']),[]).append(_r)
for _k in _byday: _byday[_k].sort(key=lambda r:r['debut'])

def nsub(code,cls):
    """Nombre de sous-lignes nécessaires : le maximum de séances sur un même jour."""
    return max([len(_byday.get((code,cls,d),[])) for d in DAYS]+[1])

def label(r):
    tag={'Quinzaine A':'   ◆ Q-A','Quinzaine B':'   ◆ Q-B','Exception':'   ◆ PONCTUEL'}.get(r['rythme'],'')
    p=[f"{r['debut']} - {r['fin']}"+tag, r['matiere'], r['prof']]
    if r['notes']: p.append(r['notes'])
    return '\n'.join(p)

def subrow_h(code,cls,n):
    """Hauteur d'une sous-ligne = nb de lignes de texte de la case la plus chargée ce jour-là."""
    mx=3
    for d in DAYS:
        lst=_byday.get((code,cls,d),[])
        if len(lst)<n: continue
        r=lst[n-1]
        L=1+max(1,math.ceil(len(r['matiere'])/CPL))+max(1,math.ceil(len(r['prof'])/CPL))
        if r['notes']: L+=max(1,math.ceil(len(r['notes'])/CPL))
        mx=max(mx,L)
    return mx*15.2+7

def build_grid(wbk, week, static, with_note=True):
    """static=True -> texte figé (fichier autonome) ; sinon formules liées à Planning."""
    code,start_iso,par,per,shname,mois = week
    sd=datetime.date.fromisoformat(start_iso)
    G=wbk.create_sheet(shname)
    G.column_dimensions['A'].width=20
    for k in range(2,9): G.column_dimensions[CL(k)].width=22
    ref=' — SEMAINE DE RÉFÉRENCE' if code=='S3' else ''
    src=('Contenu figé.' if static else
         ('Contenu généré automatiquement depuis la feuille Planning.' if mois=='Octobre'
          else 'Contenu issu du fichier d’origine, repris depuis la feuille Planning.'))
    title_block(G,'H', f"{code} · {per.upper()}   |   SEMAINE {par}{ref}",
      "Chaque case = une séance : 1re ligne l’horaire, 2e la matière, 3e l’enseignant.   "
      "La couleur indique la famille de matières.   ◆ Q-A / ◆ Q-B = séance par quinzaine.   "
      f"Case grise = aucune séance.   {src}", h2=26)
    c=G.cell(3,1,'Classe / Groupe'); c.font=font(11,True,WHITE); c.fill=fill(INK2); c.alignment=CTR; c.border=BOX_M
    for k,day in enumerate(DAYS):
        dd=sd+datetime.timedelta(days=k)
        c=G.cell(3,2+k,f"{day}\n{dd.strftime('%d/%m')}")
        c.font=font(11,True,WHITE); c.fill=fill(ACC if k>=5 else INK2); c.alignment=CTR; c.border=BOX_M
    G.row_dimensions[3].height=36
    G.cell(4,1,'__jour')                                   # ligne technique masquée
    for k,day in enumerate(DAYS): G.cell(4,2+k,day)
    G.cell(4,10,code)
    G.row_dimensions[4].hidden=True; G.row_dimensions[4].height=3
    x=5; breaks=[]
    for cy,classes in CYCLES:
        if x>5: breaks.append(x-1)
        G.merge_cells(start_row=x,start_column=1,end_row=x,end_column=8)
        c=G.cell(x,1,f"CYCLE {cy.upper()}   ({len(classes)} classes)")
        c.font=font(10,True,WHITE); c.alignment=Alignment(horizontal='left',vertical='center',indent=1)
        for k in range(1,9): G.cell(x,k).fill=fill(CYC_BAND[cy]); G.cell(x,k).border=BOX_M
        G.row_dimensions[x].height=21; x+=1
        for cls in classes:
            nb=nsub(code,cls); x0=x
            for n in range(1,nb+1):
                for k,day in enumerate(DAYS):
                    col=CL(2+k)
                    if static:
                        lst=_byday.get((code,cls,day),[])
                        v=label(lst[n-1]) if len(lst)>=n else None
                    else:
                        key=f'$J$4&"|"&$A${x0}&"|"&{col}$4&"|{n}"'
                        v=f'=IFERROR(VLOOKUP({key},Planning!$W:$X,2,0),"")'
                    cc=G.cell(x,2+k,v)
                    cc.font=font(FS_GRID); cc.alignment=LTOP; cc.border=BOX
                    cc.fill=fill(WEEKEND if k>=5 else WHITE)
                G.row_dimensions[x].height=subrow_h(code,cls,n); x+=1
            if nb>1: G.merge_cells(start_row=x0,start_column=1,end_row=x-1,end_column=1)
            c=G.cell(x0,1,cls); c.font=font(11,True,INK); c.fill=fill(CYC_FILL[cy]); c.alignment=CTR; c.border=BOX_M
            for r2 in range(x0,x):
                G.cell(r2,1).fill=fill(CYC_FILL[cy]); G.cell(r2,1).border=BOX_M
    last=x-1
    rg=f'B5:H{last}'
    # Une case = une seule séance : la couleur de famille est donc toujours exacte.
    # Les règles « famille » passent en premier ; le repère quinzaine suit et s'y ajoute
    # dans Excel (dans LibreOffice, seule la couleur de fond s'applique, le repère ◆ Q-A
    # restant lisible dans le texte).
    for _nom,_col,_mats in FAMILLES:
        t='+'.join(f'ISNUMBER(SEARCH(CHAR(10)&"{m}"&CHAR(10),B5))' for m in _mats)
        G.conditional_formatting.add(rg, FormulaRule(formula=[f'({t})>0'], fill=fill(_col)))
    G.conditional_formatting.add(rg, FormulaRule(formula=['ISNUMBER(SEARCH("◆ PONCTUEL",B5))'], font=font(FS_GRID,True,ALERT_FG)))
    G.conditional_formatting.add(rg, FormulaRule(formula=['ISNUMBER(SEARCH("◆ Q-",B5))'], font=font(FS_GRID,True,QUINZ_FG)))
    G.conditional_formatting.add(rg, FormulaRule(formula=['LEN(B5)=0'], fill=fill(GREY_L), stopIfTrue=True))
    G.freeze_panes='B5'
    for br in breaks: G.row_breaks.append(Break(id=br))
    setup_print(G, titles='1:3', area=f'A1:H{last}',
                header=f"{per}", footer=f'RS 2026-2027 · {code} (semaine {par})')
    G.sheet_properties.tabColor = INK2 if code=='S3' else ('8494A7' if mois=='Septembre' else ACC)
    return shname

grid_sheets=[]
for _w in WEEKS:
    build_grid(wb,_w,static=False)
    grid_sheets.append((_w[4],_w[0],_w[3],_w[2],_w[5]))

PL=f'Planning!'
A_=f'{PL}$A${R0}:$A${R1}'; C_=f'{PL}$C${R0}:$C${R1}'; D_=f'{PL}$D${R0}:$D${R1}'
E_=f'{PL}$E${R0}:$E${R1}'; F_=f'{PL}$F${R0}:$F${R1}'; G_=f'{PL}$G${R0}:$G${R1}'
H_=f'{PL}$H${R0}:$H${R1}'; I_=f'{PL}$I${R0}:$I${R1}'; J_=f'{PL}$J${R0}:$J${R1}'
K_=f'{PL}$K${R0}:$K${R1}'; M_=f'{PL}$M${R0}:$M${R1}'; P_=f'{PL}$P${R0}:$P${R1}'
Q_=f'{PL}$Q${R0}:$Q${R1}'; S_=f'{PL}$S${R0}:$S${R1}'; T_=f'{PL}$T${R0}:$T${R1}'; U_=f'{PL}$U${R0}:$U${R1}'
CODES=[w[0] for w in WEEKS]

# ==========================================================
# 4) VUE ENSEIGNANTS
# ==========================================================
subj_by_prof={}
for r in rows:
    subj_by_prof.setdefault(r['prof'],set()).add(r['matiere_n'])
V=wb.create_sheet('Vue Enseignants')
ncol=2+7+1+7+1+1
title_block(V,CL(ncol),'CHARGE PAR ENSEIGNANT — séances et heures par semaine',
  "Tableau entièrement calculé depuis la feuille Planning. « Heures » = somme des durées réelles. "
  "La colonne de droite signale tout chevauchement horaire détecté pour l’enseignant.")
V.cell(3,1,'Enseignant'); V.cell(3,2,'Matière(s) — indicatif')
V.merge_cells(start_row=3,start_column=3,end_row=3,end_column=9); V.cell(3,3,'NOMBRE DE SÉANCES')
V.cell(3,10,'Total\nséances')
V.merge_cells(start_row=3,start_column=11,end_row=3,end_column=17); V.cell(3,11,'HEURES D’ENSEIGNEMENT')
V.cell(3,18,'Total\nheures'); V.cell(3,19,'⚠ Chevauchements')
for k in range(1,ncol+1):
    c=V.cell(3,k); c.font=font(9,True,WHITE); c.fill=fill(INK2); c.alignment=CTR; c.border=BOX_M
for k,code in enumerate(CODES):
    for base in (3,11):
        c=V.cell(4,base+k,code); c.font=font(9,True); c.fill=fill(GREY_H); c.alignment=CTR; c.border=BOX_M
for k in (1,2,10,18,19):
    V.cell(4,k,'—').font=font(9,True); V.cell(4,k).fill=fill(GREY_H); V.cell(4,k).alignment=CTR; V.cell(4,k).border=BOX_M
V.row_dimensions[3].height=30; V.row_dimensions[4].height=16
for k,w in enumerate([17,24]+[6.5]*7+[9]+[7.5]*7+[9,15],1): V.column_dimensions[CL(k)].width=w
for i,p in enumerate(profs):
    x=5+i
    V.cell(x,1,p).alignment=LFT
    V.cell(x,2,', '.join(sorted(subj_by_prof.get(p,[])))).alignment=LFT
    for k,code in enumerate(CODES):
        V.cell(x,3+k,f'=COUNTIFS({A_},"{code}",{M_},$A{x})').alignment=CTR
        c=V.cell(x,11+k,f'=SUMIFS({H_},{A_},"{code}",{M_},$A{x})'); c.number_format='0.0'; c.alignment=CTR
    V.cell(x,10,f'=SUM(C{x}:I{x})').alignment=CTR
    c=V.cell(x,18,f'=SUM(K{x}:Q{x})'); c.number_format='0.0'; c.alignment=CTR
    V.cell(x,19,f'=COUNTIFS({M_},$A{x},{S_},"⚠ CONFLIT")').alignment=CTR
    for k in range(1,ncol+1):
        c=V.cell(x,k); c.font=font(9,True if k in (10,18) else False); c.border=BOX
        if k in (10,18): c.fill=fill(GREY_H)
    V.row_dimensions[x].height=15
lastV=4+len(profs)
V.cell(lastV+1,1,'TOTAL').font=font(9,True,WHITE)
for k in range(1,ncol+1):
    c=V.cell(lastV+1,k); c.fill=fill(INK2); c.border=BOX_M; c.alignment=CTR; c.font=font(9,True,WHITE)
for k in list(range(3,10))+[10]+list(range(11,18))+[18,19]:
    c=V.cell(lastV+1,k,f'=SUM({CL(k)}5:{CL(k)}{lastV})')
    c.font=font(9,True,WHITE); c.fill=fill(INK2); c.border=BOX_M; c.alignment=CTR
    if 11<=k<=18: c.number_format='0.0'
V.conditional_formatting.add(f'S5:S{lastV}', FormulaRule(formula=[f'$S5>0'], fill=fill(ALERT_BG), font=font(9,True,ALERT_FG)))
V.conditional_formatting.add(f'C5:I{lastV}', FormulaRule(formula=['C5=0'], font=font(9,False,'B9C4D0')))
V.freeze_panes='C5'
V.auto_filter.ref=f'A4:{CL(ncol)}{lastV}'
setup_print(V, titles='1:4', area=f'A1:{CL(ncol)}{lastV+1}', footer='RS 2026-2027 · Charge enseignants')
V.sheet_properties.tabColor=ACC

# ==========================================================
# 5) CONTRÔLES
# ==========================================================
Ck=wb.create_sheet('Contrôles')
title_block(Ck,'F','CONTRÔLES AUTOMATIQUES — fiabilité des données',
  "Tous les contrôles sont des formules vivantes : ils se recalculent dès qu’une ligne est modifiée dans Planning. "
  "Objectif : 0 anomalie dans la colonne Statut.")
for k,w in enumerate([5,58,30,13,13,46],1): Ck.column_dimensions[CL(k)].width=w
for k,h in enumerate(['#','Contrôle','Règle appliquée','Attendu','Résultat','Statut / action'],1):
    c=Ck.cell(3,k,h); c.font=font(9,True,WHITE); c.fill=fill(INK2); c.alignment=CTR; c.border=BOX_M
Ck.row_dimensions[3].height=26
CHECKS=[
 ("Conflits d’enseignant","Même enseignant, même date, horaires qui se chevauchent","0",f'=COUNTIF({S_},"⚠ CONFLIT")'),
 ("Conflits de classe","Même classe, même date, horaires qui se chevauchent","0",f'=COUNTIF({T_},"⚠ CONFLIT")'),
 ("Doublons de séance","Même date + classe + heure de début + matière","0",f'=COUNTIF({U_},"⚠ DOUBLON")'),
 ("Séances sans enseignant","Colonne Enseignant vide","0",f'=SUMPRODUCT(({A_}<>"")*({M_}=""))'),
 ("Séances sans horaire","Heure de début ou de fin vide","0",f'=SUMPRODUCT(({A_}<>"")*(({F_}="")+({G_}="")>0))'),
 ("Séances sans matière","Colonne Matière vide","0",f'=SUMPRODUCT(({A_}<>"")*({K_}=""))'),
 ("Horaires hors plage 08:00–21:30","Début < 08:00 ou fin > 21:30","0",f'=SUMPRODUCT(({A_}<>"")*({F_}<>"")*(({F_}<TIME(8,0,0))+({G_}>TIME(21,30,0))>0))'),
 ("Durées anormales","Durée < 1,0 h ou > 3,0 h","0",f'=SUMPRODUCT(({A_}<>"")*({F_}<>"")*(({H_}<1)+({H_}>3)>0))'),
 ("Heure de fin ≤ heure de début","Incohérence d’horaire","0",f'=SUMPRODUCT(({A_}<>"")*({F_}<>"")*({G_}<={F_}))'),
 ("Matières hors référentiel","Valeur absente de la liste Référentiels","0",f'=SUMPRODUCT(({A_}<>"")*({K_}<>"")*(COUNTIF(LST_Matieres,{K_})=0))'),
 ("Enseignants hors référentiel","Valeur absente de la liste Référentiels","0",f'=SUMPRODUCT(({A_}<>"")*({M_}<>"")*(COUNTIF(LST_Profs,{M_})=0))'),
 ("Classes hors référentiel","Valeur absente de la liste Référentiels","0",f'=SUMPRODUCT(({A_}<>"")*({J_}<>"")*(COUNTIF(LST_Classes,{J_})=0))'),
 ("Quinzaine A placée en semaine B","Rythme « Quinzaine A » sur une semaine de parité B","0",f'=SUMPRODUCT(({P_}="Quinzaine A")*({C_}="B"))'),
 ("Quinzaine B placée en semaine A","Rythme « Quinzaine B » sur une semaine de parité A","0",f'=SUMPRODUCT(({P_}="Quinzaine B")*({C_}="A"))'),
 ("Dates hors période couverte","Date < 14/09/2026 ou > 01/11/2026","0",f'=SUMPRODUCT(({A_}<>"")*(({D_}<DATE(2026,9,14))+({D_}>DATE(2026,11,1))>0))'),
 ("Jour ≠ jour réel de la date","Libellé du jour incohérent avec la date","0",f'=SUMPRODUCT(({A_}<>"")*({D_}<>"")*({E_}<>CHOOSE(WEEKDAY({D_},2),"Lundi","Mardi","Mercredi","Jeudi","Vendredi","Samedi","Dimanche")))'),
 ("Classes sans séance (semaine A type S5)","Chaque classe doit avoir au moins 1 séance","0",f'=SUMPRODUCT(--(COUNTIFS({A_},"S5",{J_},LST_Classes)=0))'),
 ("Classes sans séance (semaine B type S4)","Chaque classe doit avoir au moins 1 séance","0",f'=SUMPRODUCT(--(COUNTIFS({A_},"S4",{J_},LST_Classes)=0))'),
 ("Intégrité septembre — séances conservées","Les 301 séances d’origine doivent être toutes présentes","301",f'=COUNTIF({Q_},"Original (septembre)")'),
 ("Volume octobre généré","4 semaines : 99 + 101 + 99 + 101","400",f'=COUNTIF({Q_},"Généré (octobre)")'),
 ("Séances tombant un jour férié signalé","Jeudi 15/10/2026 (Fête de l’Évacuation) — à confirmer",'—',f'=COUNTIFS({D_},DATE(2026,10,15))'),
]
x=4
for i,(lib,regle,att,formule) in enumerate(CHECKS,1):
    Ck.cell(x,1,i).alignment=CTR
    Ck.cell(x,2,lib).alignment=LFT
    Ck.cell(x,3,regle).alignment=LFT
    Ck.cell(x,4,att).alignment=CTR
    Ck.cell(x,5,formule).alignment=CTR
    if att=='—':
        Ck.cell(x,6,f'=IF($E{x}=0,"Aucune séance ce jour-là","À CONFIRMER : "&$E{x}&" séance(s) planifiée(s) le 15/10")')
    else:
        Ck.cell(x,6,f'=IF($E{x}={att},"OK","À VÉRIFIER : "&$E{x}&" au lieu de {att}")')
    Ck.cell(x,6).alignment=LFT
    for k in range(1,7):
        c=Ck.cell(x,k); c.font=font(9); c.border=BOX
        if k==2: c.font=font(9,True)
        if i%2==0: c.fill=fill(GREY_L)
    Ck.row_dimensions[x].height=17; x+=1
lastC=x-1
Ck.conditional_formatting.add(f'F4:F{lastC}', FormulaRule(formula=['$F4="OK"'], font=font(9,True,OK_FG)))
Ck.conditional_formatting.add(f'A4:F{lastC}', FormulaRule(formula=['LEFT($F4,3)="À V"'], fill=fill(ALERT_BG), font=font(9,True,ALERT_FG)))
Ck.conditional_formatting.add(f'A4:F{lastC}', FormulaRule(formula=['LEFT($F4,3)="À C"'], fill=fill('FEF5E7'), font=font(9,True,QUINZ_FG)))

# --- décisions & points signalés ---
x+=1
Ck.merge_cells(start_row=x,start_column=1,end_row=x,end_column=6)
c=Ck.cell(x,1,'DÉCISIONS DE CONSTRUCTION DU PLANNING D’OCTOBRE ET POINTS SIGNALÉS')
c.font=font(10,True,WHITE); c.alignment=Alignment(horizontal='left',vertical='center',indent=1)
for k in range(1,7): Ck.cell(x,k).fill=fill(INK); Ck.cell(x,k).border=BOX_M
Ck.row_dimensions[x].height=24; x+=1
for k,h in enumerate(['#','Élément concerné','Constat','Type','Décision','Justification'],1):
    c=Ck.cell(x,k,h); c.font=font(9,True,WHITE); c.fill=fill(INK2); c.alignment=CTR; c.border=BOX_M
x+=1
NOTES=[
 ("7éme (A) mercredi · 8éme (A) vendredi · 9éme (A) lundi — Physique (Melek)",
  "Cellule porte la mention « /par quinzaine » mais la séance est présente dans les 3 semaines de septembre.",
  "Étiquette contradictoire","Traitée comme HEBDOMADAIRE",
  "Le fait observé (3 semaines sur 3) prime sur l’étiquette. Si le rythme est réellement quinzaine, passer la colonne Rythme à « Quinzaine A »."),
 ("3éme INFO — STI (Aymen)",
  "Dimanche 15:30–17:30 en semaine du 21/09 uniquement, puis lundi 15:30–17:30 dans la semaine de référence.",
  "AMBIGU","Traitée comme HEBDOMADAIRE le lundi",
  "La semaine de référence est la source prioritaire et ne porte aucune mention de quinzaine. À arbitrer : si la séance est en quinzaine, la retirer des semaines B (S4 et S6)."),
 ("2éme INFO — Informatique (Aymen)",
  "Lundi 19:00–21:00 en S1 et S2, déplacée samedi 17:00–19:00 dans la semaine de référence.",
  "Déplacement","Reconduite le SAMEDI 17:00–19:00",
  "Séance hebdomadaire relocalisée ; c’est la position de la semaine de référence qui est reconduite."),
 ("3éme SCE — Math (Lotfi)",
  "Dimanche 13:30–15:00 en S1 et S2, déplacée lundi 19:00–20:30 dans la semaine de référence.",
  "Déplacement","Reconduite le LUNDI 19:00–20:30",
  "Idem : position de la semaine de référence."),
 ("Bac INFO — STI (Aymen)",
  "Absente en S1 (elle était le dimanche), présente mercredi 19:00–21:00 en S2 et en S3.",
  "Nouvelle séance","Traitée comme HEBDOMADAIRE le mercredi",
  "Deux semaines consécutives confirment l’installation du créneau."),
 ("6éme Pilote (A) — Français (Hend), dimanche 09:30–11:00, mention « Rattrapage »",
  "Présente uniquement dans la semaine du 21/09.",
  "Séance ponctuelle","NON reconduite en octobre",
  "Séance de rattrapage explicitement ponctuelle : la reconduire créerait des séances inexistantes."),
 ("Histoire / Géographie (Abdelwaheb) — 2éme Eco, 3éme Eco, Bac Eco",
  "Le même créneau porte « Histoire » en S1 et S3, « Géographie » en S2. Aucune mention de quinzaine.",
  "Alternance non étiquetée","Alternance CONSERVÉE (Histoire en semaine A, Géographie en semaine B)",
  "Rythme déduit de l’observation sur 3 semaines. Ajouter la mention de quinzaine dans la source éviterait toute ambiguïté."),
 ("Espagnol (Imen) et Italien (Amel)",
  "Le créneau du mardi 19:15–20:45 alterne entre Bac Eco / 3éme Eco (Espagnol) et Bac SCE / 3éme INFO (Italien), sans mention de quinzaine.",
  "Alternance non étiquetée","Alternance CONSERVÉE",
  "Même logique : alternance stricte constatée sur les 3 semaines."),
 ("2éme Eco — Français (Nabil) vendredi 17:30–19:00",
  "Présente en S1 et S3, absente en S2 ; la mention « /par quinzaine » est absente alors que le créneau alterne avec Bac Eco.",
  "Mention manquante","Traitée comme QUINZAINE A",
  "Alternance stricte avec Bac Eco sur le même créneau et le même enseignant."),
 ("Enseignante « Imen Ghoumem » / « Imen Ghoumeme »",
  "Deux orthographes pour la même personne (1ére (B) et 1ére (A), SVT dimanche).",
  "Orthographe","Septembre laissé INTACT ; octobre unifié en « Imen Ghoumeme »",
  "Les données d’origine ne sont pas modifiées. Corriger la source si « Ghoumeme » est la bonne orthographe."),
 ("Mentions « /par quizaine » et « /Par quizaine »",
  "4 cellules comportent une faute de frappe sur « quinzaine ».",
  "Orthographe","Rythme normalisé dans la colonne Rythme ; texte d’origine conservé en colonne Y",
  "Aucune perte : le libellé brut d’origine reste consultable."),
 ("Salle et Mode (présentiel / en ligne)",
  "Ces deux informations sont totalement absentes du fichier d’origine.",
  "Donnée manquante","Colonnes créées et VIDES, avec menus déroulants",
  "Aucune valeur n’a été inventée. À compléter dans Planning ; les grilles les afficheront automatiquement."),
 ("Jeudi 15/10/2026",
  "Correspond à la Fête de l’Évacuation, jour férié en Tunisie.",
  "Calendrier","Séances MAINTENUES, signalées en colonne Notes",
  "À arbitrer par la direction : si l’établissement ferme, supprimer ou déplacer les séances de cette date."),
 ("Semaine S7 (26/10 → 01/11/2026)",
  "Cette semaine déborde sur le dimanche 1er novembre.",
  "Calendrier","Semaine COMPLÈTE conservée",
  "Cohérent avec le fichier d’origine, où la semaine de référence déborde déjà sur octobre (28/09 → 04/10)."),
]
for i,(el,cons,typ,dec,just) in enumerate(NOTES,1):
    Ck.cell(x,1,i).alignment=CTR
    for k,v in enumerate([el,cons,typ,dec,just],2): Ck.cell(x,k,v).alignment=LTOP
    for k in range(1,7):
        c=Ck.cell(x,k); c.font=font(9); c.border=BOX
        if k==2: c.font=font(9,True)
        if i%2==0: c.fill=fill(GREY_L)
    Ck.row_dimensions[x].height=34; x+=1
Ck.freeze_panes='A4'
setup_print(Ck, area=f'A1:F{x-1}', titles='1:3', footer='RS 2026-2027 · Contrôles et décisions')
Ck.sheet_properties.tabColor=ALERT_FG

# ==========================================================
# 6) RÉCAPITULATIF
# ==========================================================
Rc=wb.create_sheet('Récapitulatif')
title_block(Rc,'K','RÉCAPITULATIF ET INDICATEURS',
  "Tous les chiffres sont calculés depuis la feuille Planning : ils se mettent à jour automatiquement.")
for k,w in enumerate([26,34,9,12,12,12,12,12,12,12,12],1): Rc.column_dimensions[CL(k)].width=w
x=4
def band(x,txt,last='K'):
    Rc.merge_cells(start_row=x,start_column=1,end_row=x,end_column=11)
    c=Rc.cell(x,1,txt); c.font=font(10,True,WHITE); c.alignment=Alignment(horizontal='left',vertical='center',indent=1)
    for k in range(1,12): Rc.cell(x,k).fill=fill(INK); Rc.cell(x,k).border=BOX_M
    Rc.row_dimensions[x].height=23
    return x+1
def hrow(x,vals,widthsfrom=1):
    for k,v in enumerate(vals,widthsfrom):
        c=Rc.cell(x,k,v); c.font=font(9,True,WHITE); c.fill=fill(INK2); c.alignment=CTR; c.border=BOX_M
    Rc.row_dimensions[x].height=26
    return x+1
def drow(x,vals,nf=None,bold=False,shade=None,start=1):
    for k,v in enumerate(vals,start):
        c=Rc.cell(x,k,v); c.font=font(9,bold); c.border=BOX
        c.alignment=LFT if k<=2 and isinstance(v,str) else CTR
        if nf and k in nf: c.number_format=nf[k]
        if shade: c.fill=fill(shade)
    Rc.row_dimensions[x].height=16
    return x+1

x=band(x,'INDICATEURS CLÉS (sur l’ensemble de la période 14/09 → 01/11/2026)')
KPI=[('Séances planifiées (total)',f'=COUNTA({A_})','0'),
     ('Heures d’enseignement (total)',f'=SUM({H_})','0.0'),
     ('Séances de septembre (d’origine)',f'=COUNTIF({Q_},"Original (septembre)")','0'),
     ('Séances d’octobre (générées)',f'=COUNTIF({Q_},"Généré (octobre)")','0'),
     ('Classes / groupes',f'=COUNTA(LST_Classes)','0'),
     ('Enseignants',f'=COUNTA(LST_Profs)','0'),
     ('Matières distinctes',f'=COUNTA(LST_Matieres)','0'),
     ('Séances par quinzaine',f'=COUNTIF({P_},"Quinzaine A")+COUNTIF({P_},"Quinzaine B")','0'),
     ('⚠ Anomalies détectées',f'=COUNTIF({S_},"⚠ CONFLIT")+COUNTIF({T_},"⚠ CONFLIT")+COUNTIF({U_},"⚠ DOUBLON")','0')]
x=hrow(x,['Indicateur','Valeur'])
k0=x
for lib,f_,nf in KPI:
    Rc.cell(x,1,lib).font=font(9,True); Rc.cell(x,1).border=BOX; Rc.cell(x,1).alignment=LFT
    c=Rc.cell(x,2,f_); c.font=font(11,True,ACC); c.border=BOX; c.alignment=CTR; c.number_format=nf
    Rc.row_dimensions[x].height=18; x+=1
Rc.conditional_formatting.add(f'B{k0}:B{x-1}', FormulaRule(formula=[f'AND(ROW()={x-1},B{x-1}>0)'], fill=fill(ALERT_BG), font=font(11,True,ALERT_FG)))
x+=1

x=band(x,'VOLUME PAR SEMAINE')
x=hrow(x,['Semaine','Période','Parité','Séances','Heures','dont Hebdo','dont Quinz.','dont Ponct.','Classes actives','Enseignants','Moy. h / séance'])
for code,start,par,per,shname,mois in WEEKS:
    x=drow(x,[code,per,par,
      f'=COUNTIF({A_},"{code}")', f'=SUMIF({A_},"{code}",{H_})',
      f'=COUNTIFS({A_},"{code}",{P_},"Hebdomadaire")',
      f'=COUNTIFS({A_},"{code}",{P_},"Quinzaine A")+COUNTIFS({A_},"{code}",{P_},"Quinzaine B")',
      f'=COUNTIF({A_},"{code}")-D{x}+0*0' if False else f'=COUNTIFS({A_},"{code}",{P_},"Exception")+COUNTIFS({A_},"{code}",{P_},"Ponctuelle / modifiée")',
      f'=SUMPRODUCT(--(COUNTIFS({A_},"{code}",{J_},LST_Classes)>0))',
      f'=SUMPRODUCT(--(COUNTIFS({A_},"{code}",{M_},LST_Profs)>0))',
      f'=IFERROR(E{x}/D{x},"")'],
      nf={5:'0.0',11:'0.00'}, shade=GREY_L if mois=='Octobre' else None)
xt=x
# I et J = effectifs DISTINCTS : les additionner n'aurait aucun sens, on recompte sur la période
x=drow(x,['TOTAL','—','—']+[f'=SUM({CL(k)}{xt-7}:{CL(k)}{xt-1})' for k in range(4,9)]
        +[f'=SUMPRODUCT(--(COUNTIF({J_},LST_Classes)>0))', f'=SUMPRODUCT(--(COUNTIF({M_},LST_Profs)>0))']
        +[f'=IFERROR(E{xt}/D{xt},"")'],
       nf={5:'0.0',11:'0.00'}, bold=True, shade=GREY_H)
x+=1

x=band(x,'SÉANCES PAR CYCLE ET PAR SEMAINE')
x=hrow(x,['Cycle','Nb classes']+CODES+['Total'])
for cy,classes in CYCLES:
    x=drow(x,[cy,len(classes)]+[f'=COUNTIFS({A_},"{c}",{I_},$A{x})' for c in CODES]+[f'=SUM(C{x}:I{x})'],bold=False)
x+=1

x=band(x,'SÉANCES ET HEURES PAR MATIÈRE (période complète)')
x=hrow(x,['Matière','Normalisée','Séances','Heures','Enseignants','Classes','Part des séances'])
for m_ in matieres:
    x=drow(x,[m_,
      f'=IFERROR(INDEX(Référentiels!$H$5:$H${4+len(matieres)},MATCH($A{x},Référentiels!$G$5:$G${4+len(matieres)},0)),"")',
      f'=COUNTIF({K_},$A{x})', f'=SUMIF({K_},$A{x},{H_})',
      f'=SUMPRODUCT(--(COUNTIFS({K_},$A{x},{M_},LST_Profs)>0))',
      f'=SUMPRODUCT(--(COUNTIFS({K_},$A{x},{J_},LST_Classes)>0))',
      f'=IFERROR(C{x}/COUNTA({A_}),"")'], nf={4:'0.0',7:'0.0%'})
x+=1

x=band(x,'RÉPARTITION PAR JOUR (semaine A type = S5 · semaine B type = S4)')
x=hrow(x,['Jour','Séances semaine A','Heures semaine A','Séances semaine B','Heures semaine B','1re séance (A)','Dernière fin (A)'])
for day in DAYS:
    x=drow(x,[day,
      f'=COUNTIFS({A_},"S5",{E_},$A{x})', f'=SUMIFS({H_},{A_},"S5",{E_},$A{x})',
      f'=COUNTIFS({A_},"S4",{E_},$A{x})', f'=SUMIFS({H_},{A_},"S4",{E_},$A{x})',
      f'=IFERROR(MINIFS({F_},{A_},"S5",{E_},$A{x}),"")',
      f'=IFERROR(MAXIFS({G_},{A_},"S5",{E_},$A{x}),"")'], nf={3:'0.0',5:'0.0',6:'hh:mm',7:'hh:mm'})
Rc.freeze_panes='A4'
setup_print(Rc, area=f'A1:K{x-1}', titles='1:3', footer='RS 2026-2027 · Récapitulatif')
Rc.sheet_properties.tabColor='9A6B1F'

# ==========================================================
# 7) SOMMAIRE
# ==========================================================
So=wb.create_sheet('Sommaire',0)
for k,w in enumerate([3,40,62,20,20],1): So.column_dimensions[CL(k)].width=w
So.merge_cells('B2:E2'); So.merge_cells('B3:E3')
c=So['B2']; c.value='EMPLOI DU TEMPS — ANNÉE SCOLAIRE 2026-2027'; c.font=Font(name=F,size=20,bold=True,color=WHITE)
c.alignment=Alignment(horizontal='left',vertical='center',indent=1); c.fill=fill(INK)
c=So['B3']; c.value='Septembre et octobre 2026  ·  23 classes  ·  7 semaines  ·  du 14 septembre au 1er novembre 2026'
c.font=font(10,False,WHITE); c.alignment=Alignment(horizontal='left',vertical='center',indent=1); c.fill=fill(INK2)
So.row_dimensions[2].height=46; So.row_dimensions[3].height=24
for col in 'BCDE': So[f'{col}2'].fill=fill(INK); So[f'{col}3'].fill=fill(INK2)

x=5
def sband(x,t):
    So.merge_cells(start_row=x,start_column=2,end_row=x,end_column=5)
    c=So.cell(x,2,t); c.font=font(10,True,WHITE); c.alignment=Alignment(horizontal='left',vertical='center',indent=1)
    for k in range(2,6): So.cell(x,k).fill=fill(ACC); So.cell(x,k).border=BOX_M
    So.row_dimensions[x].height=22
    return x+1

x=sband(x,'NAVIGATION — cliquer sur un nom de feuille')
for k,h in enumerate(['Feuille','Contenu','Séances','Type'],2):
    c=So.cell(x,k,h); c.font=font(9,True,WHITE); c.fill=fill(INK2); c.alignment=CTR; c.border=BOX_M
x+=1
NAV=[('Planning','Table de référence : 1 ligne = 1 séance. C’est ICI que l’on modifie les données.',f'=COUNTA({A_})','Saisie'),
     ('Référentiels','Listes des classes, matières, enseignants, salles, modes, rythmes et semaines.','—','Paramètres'),
     ('Contrôles','21 contrôles automatiques + toutes les décisions prises pour construire octobre.','—','Fiabilité'),
     ('Récapitulatif','Indicateurs : volumes par semaine, cycle, matière et jour.','—','Pilotage'),
     ('Vue Enseignants','Charge de chaque enseignant, semaine par semaine (séances et heures).','—','Pilotage')]
for shname,code,per,par,mois in grid_sheets:
    NAV.append((shname,f'Grille imprimable — {per} (semaine {par}).',f'=COUNTIF({A_},"{code}")',
                'Référence' if code=='S3' else ('Septembre' if mois=='Septembre' else 'Octobre')))
for name,desc,cnt,typ in NAV:
    c=So.cell(x,2,name); c.font=Font(name=F,size=10,bold=True,color='1F5C86',underline='single')
    c.hyperlink=f"#'{name}'!A1"; c.alignment=LFT
    So.cell(x,3,desc).alignment=LFT
    So.cell(x,4,cnt).alignment=CTR
    So.cell(x,5,typ).alignment=CTR
    for k in range(2,6):
        cc=So.cell(x,k); cc.border=BOX
        if k>2: cc.font=font(9)
        if typ=='Octobre': cc.fill=fill('EAF4F5')
        elif typ=='Référence': cc.fill=fill('FEF5E7')
    So.row_dimensions[x].height=17; x+=1
x+=1

x=sband(x,'LÉGENDE')
LEG=([('__TITRE__','COULEUR DES CASES — famille de matières','','','')]
   + [(_n,'Matières : '+', '.join(_m),'','','') for _n,_c,_m in FAMILLES]
   + [('__TITRE__','COULEUR DE LA COLONNE CLASSE — cycle','','','')]
   + [('Primaire','4éme, 5éme, 6éme Pilote (A), 6éme (B)','','',''),
      ('Collège','7éme, 8éme et 9éme (8 groupes)','','',''),
      ('Secondaire','1ére, 2éme et 3éme (8 groupes)','','',''),
      ('Bac','Bac Eco, Bac SCE, Bac INFO','','','')]
   + [('__TITRE__','REPÈRES DANS LE TEXTE','','','')]
   + [('◆ Q-A','Séance par quinzaine, semaines A : 14/09 · 28/09 · 12/10 · 26/10','','',''),
      ('◆ Q-B','Séance par quinzaine, semaines B : 21/09 · 05/10 · 19/10','','',''),
      ('◆ EN LIGNE','Séance à distance (à renseigner dans la colonne Mode de Planning)','','',''),
      ('◆ PONCTUEL','Séance exceptionnelle, non récurrente (ex. rattrapage)','','',''),
      ('Case grisée','Aucune séance sur ce créneau','','',''),
      ('⚠ en rouge','Conflit horaire, doublon ou anomalie détectée automatiquement','','','')])
FAM_COL={_n:_c for _n,_c,_m in FAMILLES}
for i,row_ in enumerate(LEG):
    lab=row_[0]
    if lab=='__TITRE__':                      # sous-titre de bloc dans la légende
        So.merge_cells(start_row=x,start_column=2,end_row=x,end_column=5)
        c=So.cell(x,2,row_[1]); c.font=font(9,True,WHITE); c.fill=fill(INK2)
        c.alignment=Alignment(horizontal='left',vertical='center',indent=1); c.border=BOX_M
        So.row_dimensions[x].height=16; x+=1; continue
    c=So.cell(x,2,lab); c.alignment=LFT; c.border=BOX
    if lab in FAM_COL:                        # pastille réelle de la couleur de la famille
        c.font=font(9,True,INK); c.fill=fill(FAM_COL[lab])
    elif lab in CYC_FILL:                     # pastille réelle de la couleur du cycle
        c.font=font(9,True,INK); c.fill=fill(CYC_FILL[lab])
    else:
        c.font=font(9,True,QUINZ_FG if 'Q-' in lab else (ALERT_FG if '⚠' in lab or 'PONCTUEL' in lab else INK))
    So.merge_cells(start_row=x,start_column=3,end_row=x,end_column=5)
    So.cell(x,3,row_[1]).font=font(9); So.cell(x,3).alignment=LFT; So.cell(x,3).border=BOX
    So.row_dimensions[x].height=16; x+=1
# pastilles de cycle
xl=x-len(LEG)
for k,(cy,_) in enumerate(CYCLES):
    pass
x+=1

x=sband(x,'MODE D’EMPLOI')
HOW=[('1','Modifier, ajouter ou supprimer une séance','Tout se passe dans la feuille Planning : 1 ligne = 1 séance. Les grilles hebdomadaires et tous les indicateurs se recalculent seuls.'),
     ('2','Ajouter une séance','Insérer une ligne dans Planning, puis recopier les formules des colonnes H, S, T, U, V, W et X depuis la ligne du dessus (double-clic sur la poignée de recopie).'),
     ('3','Renseigner une salle ou une séance en ligne','Colonnes Salle et Mode de Planning (menus déroulants). L’indication apparaît aussitôt dans la grille de la semaine.'),
     ('4','Vérifier la fiabilité','Ouvrir la feuille Contrôles : la colonne Statut doit afficher « OK » partout.'),
     ('5','Imprimer ou envoyer en PDF','Chaque grille est déjà paramétrée : A4 paysage, 1 page de large, un cycle par page, en-têtes répétés. Fichier ▸ Exporter au format PDF.'),
     ('6','Préparer novembre','Dupliquer une grille d’octobre, renommer l’onglet, mettre à jour la cellule J4 (masquée) avec le nouveau code de semaine, ajouter la semaine dans Référentiels, puis ajouter ses lignes dans Planning (filtrer une semaine existante, copier, coller et changer Semaine / Période / Date).')]
for k,h in enumerate(['#','Action','Comment faire'],2):
    c=So.cell(x,k,h); c.font=font(9,True,WHITE); c.fill=fill(INK2); c.alignment=CTR; c.border=BOX_M
So.merge_cells(start_row=x,start_column=4,end_row=x,end_column=5); x+=1
for n,act,how in HOW:
    So.cell(x,2,f'{n}.  {act}').font=font(9,True); So.cell(x,2).alignment=LFT; So.cell(x,2).border=BOX
    So.merge_cells(start_row=x,start_column=3,end_row=x,end_column=5)
    So.cell(x,3,how).font=font(9); So.cell(x,3).alignment=LTOP; So.cell(x,3).border=BOX
    So.row_dimensions[x].height=30; x+=1
x+=1

x=sband(x,'À LIRE AVANT UTILISATION')
WARN=[("Salle et Mode (présentiel / en ligne) sont absents du fichier d’origine : les colonnes existent mais sont vides. Aucune valeur n’a été inventée."),
      ("Le planning d’octobre a été construit à partir de la semaine du 28/09 au 04/10, désignée comme référence. Les séances par quinzaine ont été identifiées en comparant les 3 semaines de septembre."),
      ("Les 301 séances de septembre sont conservées à l’identique, y compris les fautes de frappe d’origine (colonne « Texte d’origine » de Planning)."),
      ("14 décisions et points ambigus sont détaillés en bas de la feuille Contrôles. Trois méritent votre arbitrage : les Physique « /par quinzaine » de Melek, le STI du lundi en 3éme INFO, et le jeudi 15/10 (jour férié)."),
      ("Les grilles hebdomadaires sont calculées par formules : à la première ouverture, laisser Excel recalculer (ou appuyer sur F9).")]
for w_ in WARN:
    So.merge_cells(start_row=x,start_column=2,end_row=x,end_column=5)
    c=So.cell(x,2,'•   '+w_); c.font=font(9); c.alignment=LTOP; c.border=BOX; c.fill=fill('FEF5E7')
    So.row_dimensions[x].height=28; x+=1

So.sheet_view.showGridLines=False
setup_print(So, orient='portrait', area=f'B1:E{x-1}', fitw=1, fith=0, footer='RS 2026-2027 · Sommaire')
So.sheet_properties.tabColor='F1C232'
So.sheet_view.zoomScale=100
wb.active=0

wb.save('RS_2026-2027_Planning_Septembre-Octobre.xlsx')
print("OK -> RS_2026-2027_Planning_Septembre-Octobre.xlsx")
print("  feuilles:", wb.sheetnames)

# ==========================================================
# 8) SECOND CLASSEUR : les 4 semaines d'octobre, rien d'autre
#    Contenu figé : fichier autonome, aucun recalcul nécessaire.
# ==========================================================
wo = Workbook(); wo.remove(wo.active)
for _w in WEEKS:
    if _w[5]=='Octobre': build_grid(wo,_w,static=True)
wo.active=0
wo.save('RS_2026-2027_Octobre_seul.xlsx')
print("OK -> RS_2026-2027_Octobre_seul.xlsx")
print("  feuilles:", wo.sheetnames)
