# -*- coding: utf-8 -*-
import json, datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, NamedStyle
from openpyxl.utils import get_column_letter as CL
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.properties import PageSetupProperties
from openpyxl.comments import Comment

D=json.load(open('rows2.json',encoding='utf-8'))
M=json.load(open('model2.json',encoding='utf-8'))
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

# --- Une couleur par MATIÈRE. Teintes réparties sur la roue chromatique et
#     vérifiées : écart Lab minimum 13.9 (nettement distinguables), contraste
#     minimum 9.2:1 avec le texte gras noir. Math = bleu, Physique = jaune.
MATIERES=[
          ('Math', '95B5EC', ['Math', 'رياضيات']),
          ('Physique', 'ECDC95', ['Physique', 'فيزياء']),
          ('Français', 'CBF3E2', ['Français']),
          ('Arabe', 'ECF3CB', ['Arabe', 'عربية']),
          ('Anglais', 'AAF6BB', ['Anglais']),
          ('SVT', '95ECE6', ['SVT']),
          ('Éveil scientifique', 'CFF6AA', ['Éveil scientifique', 'إيقاظ علمي']),
          ('Histoire', 'F6AAD8', ['Histoire']),
          ('Géographie', 'D195EC', ['Géographie']),
          ('Philo', 'CBCCF3', ['Philo']),
          ('Eco', 'F6CDAA', ['Eco']),
          ('Gestion', 'AAE0F6', ['Gestion']),
          ('Informatique', 'C4AAF6', ['Informatique']),
          ('ALGO', 'F6AABD', ['ALGO']),
          ('STI', 'F3CBF1', ['STI']),
          ('Espagnol', 'A0EC95', ['Espagnol']),
          ('Italien', 'F3CFCB', ['Italien']),
]
MAT_COLOR={lab:c for _n,c,labs in MATIERES for lab in labs}

# --- Statuts : posés par l'utilisateur dans Planning, PRIORITAIRES sur la
#     couleur de matière. Ce sont les seules règles conditionnelles des grilles,
#     pour que toute autre cellule reste colorable à la main.
STATUTS=[
         ('Annulé', 'F5A7A0'),
         ('Examen / Devoir', 'F7D154'),
         ('Rattrapage', 'C9A0DC'),
         ('À confirmer', 'F9C784'),
         ('Changement de salle', '8FC7E8'),
         ('Séance en ligne', '9FE0C8'),
]


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
       'Matière','Matière (normalisée)','Enseignant','Salle','Mode','Rythme',
       'Séance partagée avec','Nb groupes','Ligne principale','ID séance','Origine','Notes',
       'Statut','Cours en option','⚠ Conflit enseignant','⚠ Conflit classe','⚠ Doublon','⚠ Partage incohérent',
       'Rang','Clé','Libellé grille','Texte d\'origine','Réf. cellule source']
WID = [9,30,7,11,11,7.5,7.5,9,12,18,20,19,17,10,12,24,34,9,12,11,19,30,19,14,19,17,12,17,7,34,46,52,26]
NB = len(rows); R0 = 4; R1 = R0+NB-1
title_block(P, CL(len(HDR)),
  'PLANNING DES SÉANCES — Année scolaire 2026-2027',
  "TABLE DE RÉFÉRENCE : toute modification se fait ici. 1 ligne = 1 groupe dans une séance. "
  "Une séance partagée entre plusieurs groupes occupe autant de lignes que de groupes, reliées par le même ID séance "
  "(seule la 1re porte « Ligne principale = Oui », pour ne compter l'heure d'enseignement qu'une fois). "
  "Colonnes Y à AG = zone technique, ne pas saisir.")
for j,(h,w) in enumerate(zip(HDR,WID),1):
    c=P.cell(3,j,h); c.font=font(9,True,WHITE)
    c.fill=fill(INK2 if j<17 else (ACC if j<21 else (INK2 if j<23 else ('9A6B1F' if j<25 else ('B03A2E' if j<29 else '6B7A8C')))))
    c.alignment=CTR; c.border=BOX_M
    P.column_dimensions[CL(j)].width=w
P.row_dimensions[3].height=34

for i,r in enumerate(rows):
    x=R0+i
    P.cell(x,1,r['semaine']).alignment=CTR
    P.cell(x,2,r['periode']).alignment=LFT
    P.cell(x,3,r['parite']).alignment=CTR
    c=P.cell(x,4,datetime.date.fromisoformat(r['date'])); c.number_format='ddd dd/mm/yyyy'; c.alignment=CTR
    P.cell(x,5,r['jour']).alignment=CTR
    for col,val in ((6,r['debut']),(7,r['fin'])):
        hh,mi=val.split(':'); c=P.cell(x,col,datetime.time(int(hh),int(mi))); c.number_format='hh:mm'; c.alignment=CTR
    c=P.cell(x,8,f'=IF(OR($F{x}="",$G{x}=""),"",($G{x}-$F{x})*24)'); c.number_format='0.00'; c.alignment=CTR
    P.cell(x,9,r['cycle']).alignment=CTR
    P.cell(x,10,r['classe']).alignment=LFT
    P.cell(x,11,r['matiere']).alignment=LFT
    P.cell(x,12,r['matiere_n']).alignment=LFT
    P.cell(x,13,r['prof']).alignment=LFT
    P.cell(x,14,None).alignment=CTR          # Salle : absente de la source
    P.cell(x,15,None).alignment=CTR          # Mode  : absent de la source
    P.cell(x,16,r['rythme']).alignment=CTR
    P.cell(x,17,r['groupes'] if r['nbgr']>1 else '').alignment=LFT
    P.cell(x,18,r['nbgr']).alignment=CTR
    P.cell(x,19,r['principale']).alignment=CTR
    P.cell(x,20,r['seance']).alignment=CTR
    P.cell(x,21,r['origine']).alignment=LFT
    P.cell(x,22,r['notes']).alignment=LFT
    P.cell(x,23,None).alignment=CTR                 # Statut : à renseigner par l'utilisateur
    P.cell(x,24,r.get('option','')).alignment=CTR
    # --- contrôles vivants ---
    # Une séance partagée occupe plusieurs lignes au même horaire : on ne compte comme
    # conflit que les chevauchements portant un ID SÉANCE DIFFÉRENT. Et deux cours en
    # option sur le même créneau (LV3 Espagnol / Italien) ne sont pas un conflit.
    P.cell(x,25,f'=IF($M{x}="","",IF(SUMPRODUCT(($D${R0}:$D${R1}=$D{x})*($M${R0}:$M${R1}=$M{x})*($T${R0}:$T${R1}<>$T{x})*($F${R0}:$F${R1}<$G{x})*($G${R0}:$G${R1}>$F{x}))>0,"⚠ CONFLIT",""))').alignment=CTR
    P.cell(x,26,(f'=IF($J{x}="","",IF(SUMPRODUCT(($D${R0}:$D${R1}=$D{x})*($J${R0}:$J${R1}=$J{x})*($T${R0}:$T${R1}<>$T{x})'
                 f'*($F${R0}:$F${R1}<$G{x})*($G${R0}:$G${R1}>$F{x})*(1-($X${R0}:$X${R1}="Oui")*($X{x}="Oui")))>0,"⚠ CONFLIT",""))')).alignment=CTR
    P.cell(x,27,f'=IF(COUNTIFS($D${R0}:$D${R1},$D{x},$J${R0}:$J${R1},$J{x},$F${R0}:$F${R1},$F{x},$K${R0}:$K${R1},$K{x})>1,"⚠ DOUBLON","")').alignment=CTR
    P.cell(x,28,(f'=IF(COUNTIF($T${R0}:$T${R1},$T{x})<>COUNTIFS($T${R0}:$T${R1},$T{x},$D${R0}:$D${R1},$D{x},'
                 f'$F${R0}:$F${R1},$F{x},$G${R0}:$G${R1},$G{x},$K${R0}:$K${R1},$K{x},$M${R0}:$M${R1},$M{x}),"⚠ INCOHÉRENT","")')).alignment=CTR
    P.cell(x,29,f'=COUNTIFS($A${R0}:$A{x},$A{x},$J${R0}:$J{x},$J{x},$E${R0}:$E{x},$E{x})').alignment=CTR
    P.cell(x,30,f'=$A{x}&"|"&$J{x}&"|"&$E{x}&"|"&$AC{x}').alignment=LFT
    P.cell(x,31,(f'=IF($W{x}<>"","⚑ "&$W{x}&CHAR(10),"")'
                 f'&TEXT($F{x},"hh:mm")&"-"&TEXT($G{x},"hh:mm")'
                 f'&IF(OR($P{x}="Quinzaine A",AND($P{x}="Hebdo · groupes alternés",$C{x}="A"))," Q-A",'
                 f'IF(OR($P{x}="Quinzaine B",AND($P{x}="Hebdo · groupes alternés",$C{x}="B"))," Q-B",'
                 f'IF($P{x}="Exception"," PONCT.","")))'
                 f'&IF($X{x}="Oui"," OPT","")'
                 f'&CHAR(10)&$K{x}&CHAR(10)&$M{x}'
                 f'&IF($N{x}<>"",CHAR(10)&"Salle "&$N{x},"")&IF($O{x}="En ligne",CHAR(10)&"EN LIGNE","")'
                 f'&IF($V{x}<>"",CHAR(10)&$V{x},"")')).alignment=LTOP
    P.cell(x,32,r['brut']).alignment=LFT
    P.cell(x,33,r['src']).alignment=LFT
    for j in range(1,len(HDR)+1):
        cc=P.cell(x,j); cc.font=font(9); cc.border=BOX
        if j<=22: cc.fill=fill(CYC_FILL[r['cycle']])          # teinte de cycle, posée en dur
        if j in (11,12): cc.fill=fill(MAT_COLOR.get(r['matiere'], CYC_FILL[r['cycle']]))
        if j>=25: cc.font=font(8,c='6B7A8C')
    P.row_dimensions[x].height=15

P.freeze_panes='F4'
P.auto_filter.ref=f'A3:{CL(len(HDR))}{R1}'
# Fonds posés en dur ci-dessus : ils restent donc modifiables à la main.
# Seules les ALERTES et les STATUTS gardent une mise en forme conditionnelle,
# parce qu'ils doivent justement l'emporter sur toute couleur manuelle.
for _nom,_col in STATUTS:
    P.conditional_formatting.add(f'W{R0}:W{R1}', FormulaRule(formula=[f'$W{R0}="{_nom}"'], fill=fill(_col), stopIfTrue=True))
P.conditional_formatting.add(f'P{R0}:P{R1}', FormulaRule(formula=[f'LEFT($P{R0},9)="Quinzaine"'], font=font(9,True,QUINZ_FG)))
P.conditional_formatting.add(f'P{R0}:P{R1}', FormulaRule(formula=[f'$P{R0}="Hebdo · groupes alternés"'], font=font(9,True,ACC)))
P.conditional_formatting.add(f'U{R0}:U{R1}', FormulaRule(formula=[f'$U{R0}="Généré (octobre)"'], font=font(9,False,ACC)))
for col in ('Y','Z','AA','AB'):
    P.conditional_formatting.add(f'{col}{R0}:{col}{R1}', FormulaRule(formula=[f'LEN(TRIM(${col}{R0}))>0'], fill=fill(ALERT_BG), font=font(8,True,ALERT_FG)))
setup_print(P, titles='1:3', area=f'A1:X{R1}',
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
        ('MATIÈRES',7,['Matière (libellé affiché)','Normalisée','Séances (total)','Couleur'],
         [[m_,{'عربية':'Arabe','رياضيات':'Math','إيقاظ علمي':'Éveil scientifique','فيزياء':'Physique'}.get(m_,m_),None,None] for m_ in matieres],[22,19,14,11]),
        ('ENSEIGNANTS',11,['Enseignant','Séances / sem. A','Heures / sem. A'],
         [[p,None,None] for p in profs],[18,15,15]),
        ('SALLES',15,['Salle (à compléter)'],[[None] for _ in range(12)],[18]),
        ('MODES',17,['Mode'],[[x] for x in modes],[14]),
        ('RYTHMES',19,['Rythme'],[[x] for x in rythmes],[21]),
        ('STATUTS',28,['Statut (couleur prioritaire)'],[[n] for n,_c in STATUTS],[26])]
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
    cc=Rf.cell(5+i,10,'')                       # pastille de la couleur de la matière
    if m_ in MAT_COLOR: cc.fill=fill(MAT_COLOR[m_])
    cc.border=BOX
for i,p in enumerate(profs):
    x=5+i
    Rf.cell(x,12,f'=COUNTIFS(Planning!$A${R0}:$A${R1},"S5",Planning!$M${R0}:$M${R1},$K{x})')
    Rf.cell(x,13,f'=SUMIFS(Planning!$H${R0}:$H${R1},Planning!$A${R0}:$A${R1},"S5",Planning!$M${R0}:$M${R1},$K{x})')
    Rf.cell(x,13).number_format='0.0'
# semaines
for _i,(_n,_c) in enumerate(STATUTS):
    Rf.cell(5+_i,28).fill=fill(_c)
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
        'LST_Rythmes':f"Référentiels!$S$5:$S${4+len(rythmes)}",
        'LST_Statuts':f"Référentiels!$AB$5:$AB${4+len(STATUTS)}"}
from openpyxl.workbook.defined_name import DefinedName
for n,ref in defn.items(): wb.defined_names.add(DefinedName(n, attr_text=ref))
for col,nm,strict in (('J','LST_Classes',False),('K','LST_Matieres',False),('M','LST_Profs',False),
                      ('N','LST_Salles',False),('O','LST_Modes',True),('P','LST_Rythmes',True),
                      ('W','LST_Statuts',True)):
    dv=DataValidation(type='list',formula1=f'={nm}',allow_blank=True,
                      errorStyle='stop' if strict else 'warning',
                      error="Valeur absente du référentiel.", errorTitle='Vérification',
                      prompt=f"Choisir dans la liste ({nm}).", promptTitle='Saisie guidée')
    dv.showErrorMessage=True; dv.showInputMessage=False
    P.add_data_validation(dv); dv.add(f'{col}{R0}:{col}{R1+200}')

# ==========================================================
# 3) GRILLES HEBDOMADAIRES — une séance écrite UNE SEULE FOIS
#    Une séance partagée entre plusieurs groupes occupe une cellule FUSIONNÉE
#    verticalement sur les lignes de ces groupes, comme dans le fichier d'origine.
# ==========================================================
from openpyxl.worksheet.pagebreak import Break
import math
CPL=19          # caractères tenant sur une ligne (colonne « jour », gras 12 pt)
FS_GRID=12      # taille de police des cases
BLACK='000000'

_byday={}
for _r in rows:
    _byday.setdefault((_r['semaine'],_r['classe'],_r['jour']),[]).append(_r)
for _k in _byday: _byday[_k].sort(key=lambda r:r['debut'])

# séance -> liste ordonnée de ses groupes
_groups={}
for _r in rows:
    _groups.setdefault((_r['semaine'],_r['seance']),set()).add(_r['classe'])
_groups={k:sorted(v,key=CLASS_ORDER.index) for k,v in _groups.items()}

def role_of(r):
    """Place d'un groupe dans une séance partagée : haut / milieu / bas / seul."""
    g=_groups[(r['semaine'],r['seance'])]
    if len(g)==1: return 'seul'
    i=g.index(r['classe'])
    return 'haut' if i==0 else ('bas' if i==len(g)-1 else 'milieu')

def tag_of(r):
    """Repère vu du GROUPE : une séance à groupes alternés revient une semaine sur deux."""
    ry=r['rythme']; t=''
    if ry=='Quinzaine A' or (ry=='Hebdo · groupes alternés' and r['parite']=='A'): t=' Q-A'
    elif ry=='Quinzaine B' or (ry=='Hebdo · groupes alternés' and r['parite']=='B'): t=' Q-B'
    elif ry=='Exception': t=' PONCT.'
    if r.get('option')=='Oui': t+=' OPT'
    return t

def label(r):
    # Le partage n'est plus écrit : la fusion de la cellule le montre déjà.
    p=[]
    if r.get('statut'): p.append('⚑ '+r['statut'])
    p+=[f"{r['debut']}-{r['fin']}"+tag_of(r), r['matiere'], r['prof']]
    if r['notes']: p.append(r['notes'])
    return '\n'.join(p)

def _nlines(r):
    n =max(1,math.ceil(len(f"{r['debut']}-{r['fin']}"+tag_of(r))/CPL))
    n+=max(1,math.ceil(len(r['matiere'])/CPL))+max(1,math.ceil(len(r['prof'])/CPL))
    if r['notes']: n+=max(1,math.ceil(len(r['notes'])/CPL))
    return n

def nsub(code,cls):
    return max([len(_byday.get((code,cls,d),[])) for d in DAYS]+[1])

def slots(code,cls,day,h):
    """Attribue à chaque séance du jour sa sous-ligne.
    Une séance partagée doit toucher le bord qui la relie au groupe voisin :
    dernière sous-ligne si le groupe est en haut du partage, première s'il est en bas."""
    lst=_byday.get((code,cls,day),[]); out={}; taken=set()
    for r in lst:
        ro=role_of(r)
        if ro=='haut': sl=h-1
        elif ro=='bas': sl=0
        elif ro=='milieu': sl=0
        else: continue
        out[id(r)]=sl; taken.add(sl)
    for r in lst:
        if id(r) in out: continue
        sl=next(k for k in range(h) if k not in taken)
        out[id(r)]=sl; taken.add(sl)
    return out

def build_grid(wbk, week, static):
    code,start_iso,par,per,shname,mois = week
    sd=datetime.date.fromisoformat(start_iso)
    G=wbk.create_sheet(shname)
    G.column_dimensions['A'].width=17
    for k in range(2,9): G.column_dimensions[CL(k)].width=23
    ref=' — SEMAINE DE RÉFÉRENCE' if code=='S3' else ''
    title_block(G,'H', f"{code} · {per.upper()}   |   SEMAINE {par}{ref}",
      "Une case = une séance : horaire / matière / enseignant.   Chaque matière a sa couleur.   "
      "Une case fusionnée sur plusieurs classes = séance commune à ces groupes.   "
      "Q-A ou Q-B = une semaine sur deux.   OPT = cours au choix.   Les couleurs sont modifiables à la main ; seul un statut ⚑ les remplace.", h2=26)
    c=G.cell(3,1,'Classe / Groupe'); c.font=font(11,True,WHITE); c.fill=fill(INK2); c.alignment=CTR; c.border=BOX_M
    for k,day in enumerate(DAYS):
        dd=sd+datetime.timedelta(days=k)
        c=G.cell(3,2+k,f"{day}\n{dd.strftime('%d/%m')}")
        c.font=font(11,True,WHITE); c.fill=fill(ACC if k>=5 else INK2); c.alignment=CTR; c.border=BOX_M
    G.row_dimensions[3].height=36

    # --- répartition des lignes ---
    H={}; TOP={}; x=5; breaks=[]; bands=[]
    for cy,classes in CYCLES:
        if x>5: breaks.append(x-1)
        bands.append((x,cy,len(classes))); x+=1
        for cls in classes:
            H[cls]=nsub(code,cls); TOP[cls]=x; x+=H[cls]
    last=x-1
    MID=Alignment(horizontal='left',vertical='center',wrap_text=True)

    # --- fond, bordures, hauteurs ---
    for cy,classes in CYCLES:
        for cls in classes:
            for sl in range(H[cls]):
                r=TOP[cls]+sl
                for k in range(7):
                    cc=G.cell(r,2+k); cc.font=font(FS_GRID,True,BLACK); cc.alignment=MID
                    cc.border=BOX; cc.fill=fill('F2F0EA' if k>=5 else GREY_L)   # case vide
                mx=3
                for day in DAYS:
                    for rr in _byday.get((code,cls,day),[]):
                        if slots(code,cls,day,H[cls])[id(rr)]==sl and role_of(rr)=='seul':
                            mx=max(mx,_nlines(rr))
                G.row_dimensions[r].height=mx*18+9

    # --- bandeaux de cycle + noms de classes ---
    for rb,cy,n in bands:
        G.merge_cells(start_row=rb,start_column=1,end_row=rb,end_column=8)
        c=G.cell(rb,1,f"CYCLE {cy.upper()}   ({n} classes)")
        c.font=font(10,True,WHITE); c.alignment=Alignment(horizontal='left',vertical='center',indent=1)
        for k in range(1,9): G.cell(rb,k).fill=fill(CYC_BAND[cy]); G.cell(rb,k).border=BOX_M
        G.row_dimensions[rb].height=21
    for cy,classes in CYCLES:
        for cls in classes:
            for sl in range(H[cls]):
                cc=G.cell(TOP[cls]+sl,1); cc.fill=fill(CYC_FILL[cy]); cc.border=BOX_M
            c=G.cell(TOP[cls],1,cls); c.font=font(12,True,BLACK); c.alignment=CTR
            if H[cls]>1:   # le nom de classe couvre toutes ses sous-lignes
                G.merge_cells(start_row=TOP[cls],start_column=1,end_row=TOP[cls]+H[cls]-1,end_column=1)

    # --- contenu : une seule écriture par séance, fusion si partagée ---
    merges=[]
    seen=set()
    for cy,classes in CYCLES:
        for cls in classes:
            for k,day in enumerate(DAYS):
                col=2+k
                sl=slots(code,cls,day,H[cls])
                for r in _byday.get((code,cls,day),[]):
                    ro=role_of(r)
                    if ro in ('milieu','bas'): continue      # écrite par le groupe du haut
                    key=(r['semaine'],r['seance'],day)
                    if key in seen: continue
                    seen.add(key)
                    r0=TOP[cls]+sl[id(r)]
                    if ro=='haut':
                        g=_groups[(r['semaine'],r['seance'])]
                        r1=TOP[g[-1]]                        # 1re sous-ligne du dernier groupe
                        # on étend la fusion aux sous-lignes libres ce jour-là, pour que le bloc
                        # couvre franchement la bande de chaque classe concernée
                        occ={sl[id(o)] for o in _byday.get((code,cls,day),[]) if o is not r}
                        while r0-1>=TOP[cls] and (r0-1-TOP[cls]) not in occ: r0-=1
                        lastc=g[-1]; sll=slots(code,lastc,day,H[lastc])
                        occl={sll[id(o)] for o in _byday.get((code,lastc,day),[]) if o['seance']!=r['seance']}
                        while r1+1<=TOP[lastc]+H[lastc]-1 and (r1+1-TOP[lastc]) not in occl: r1+=1
                        merges.append((r0,r1,col))
                    else:
                        r1=r0
                    v = label(r) if static else f'=IFERROR(VLOOKUP("{r["seance"]}",Planning!$T:$AE,12,0),"")'
                    cc=G.cell(r0,col,v); cc.font=font(FS_GRID,True,BLACK); cc.alignment=MID
                    # couleur de la matière, posée EN DUR (donc modifiable à la main)
                    mc=MAT_COLOR.get(r['matiere'])
                    if mc:
                        for rr in range(r0,r1+1):
                            G.cell(rr,col).fill=fill(mc)
    for r0,r1,col in merges:
        if r1>r0: G.merge_cells(start_row=r0,start_column=col,end_row=r1,end_column=col)

    rg=f'B5:H{last}'
    # Les couleurs de matière sont posées en dur : on peut donc recolorer n'importe
    # quelle case à la main. Les seules règles conditionnelles sont les STATUTS,
    # qui doivent justement l'emporter sur la couleur manuelle.
    for _nom,_col in STATUTS:
        G.conditional_formatting.add(rg, FormulaRule(formula=[f'ISNUMBER(SEARCH("⚑ {_nom}",B5))'],
                                                     fill=fill(_col), stopIfTrue=True))
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
Q_=f'{PL}$U${R0}:$U${R1}'   # Origine
S_=f'{PL}$Y${R0}:$Y${R1}'   # conflit enseignant
T_=f'{PL}$Z${R0}:$Z${R1}'   # conflit classe
U_=f'{PL}$AA${R0}:$AA${R1}' # doublon
Z_=f'{PL}$AB${R0}:$AB${R1}' # partage incohérent
OPT_=f'{PL}$X${R0}:$X${R1}' # cours en option
STA_=f'{PL}$W${R0}:$W${R1}' # statut
PR_=f'{PL}$S${R0}:$S${R1}'  # ligne principale
ID_=f'{PL}$T${R0}:$T${R1}'  # ID séance
NB_=f'{PL}$R${R0}:$R${R1}'  # nb groupes
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
  "Une séance partagée entre plusieurs groupes n'est comptée QU'UNE FOIS : c'est la charge réelle de l'enseignant. "
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
        V.cell(x,3+k,f'=COUNTIFS({A_},"{code}",{M_},$A{x},{PR_},"Oui")').alignment=CTR
        c=V.cell(x,11+k,f'=SUMIFS({H_},{A_},"{code}",{M_},$A{x},{PR_},"Oui")'); c.number_format='0.0'; c.alignment=CTR
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
 ("Conflits d’enseignant","Même enseignant, même date, horaires qui se chevauchent, séances différentes","0",f'=COUNTIF({S_},"⚠ CONFLIT")'),
 ("Conflits de classe","Même groupe, même date, horaires qui se chevauchent, séances différentes","0",f'=COUNTIF({T_},"⚠ CONFLIT")'),
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
 ("Intégrité septembre — lignes conservées","370 lignes groupe-séance (124 + 121 + 125)","370",f'=COUNTIF({Q_},"Original (septembre)")'),
 ("Intégrité septembre — séances distinctes","301 séances (100 + 100 + 101)","301",f'=COUNTIFS({Q_},"Original (septembre)",{PR_},"Oui")'),
 ("Volume octobre généré — lignes","490 lignes (120 + 125 + 120 + 125)","490",f'=COUNTIF({Q_},"Généré (octobre)")'),
 ("Volume octobre généré — séances","400 séances (99 + 101 + 99 + 101)","400",f'=COUNTIFS({Q_},"Généré (octobre)",{PR_},"Oui")'),
 ("Séances partagées — total","106 séances réunissant plusieurs groupes","106",f'=COUNTIFS({PR_},"Oui",{NB_},">1")'),
 ("Séances partagées cohérentes","Toutes les lignes d’un même ID séance ont mêmes date, horaires, matière et enseignant","0",f'=COUNTIF({Z_},"⚠ INCOHÉRENT")'),
 ("Lignes principales = séances","Exactement une ligne principale par ID séance","0",f'=SUMPRODUCT(--(COUNTIFS({ID_},{ID_},{PR_},"Oui")<>1))'),
 ("Lignes sans ID séance","Chaque ligne doit porter un ID","0",f'=SUMPRODUCT(({A_}<>"")*({ID_}=""))'),
 ("Chevauchements admis (cours en option)","Espagnol / Italien sur le même créneau : le groupe se scinde, ce n’est pas un conflit",'—',f'=COUNTIF({OPT_},"Oui")'),
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
        Ck.cell(x,6,f'=IF($B{x}="Séances tombant un jour férié signalé",IF($E{x}=0,"Aucune séance ce jour-là","À CONFIRMER : "&$E{x}&" ligne(s) planifiée(s) le 15/10"),"INFORMATIF : "&$E{x}&" ligne(s)")')
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
 ("SÉANCES PARTAGÉES ENTRE PLUSIEURS GROUPES — 16 en semaine A, 14 en semaine B",
  "Dans le fichier d’origine, une séance réunissant plusieurs groupes est écrite une seule fois, dans une cellule fusionnée qui déborde sur les blocs de toutes les classes concernées.",
  "Structure","CONSERVÉES et rendues explicites",
  "Chaque groupe a désormais sa propre ligne dans Planning, les lignes d’une même séance étant reliées par un ID séance commun. Dans la grille, la séance n'est écrite qu'une fois, dans une cellule fusionnée qui couvre les lignes de tous les groupes réunis. La charge enseignant ne compte la séance qu’une fois."),
 ("Physique (Melek) — 7éme (A)+(B) mercredi, 8éme (A)+(B)+(C) vendredi, 9éme (A)+(B)+(C) lundi",
  "Séances partagées sur tout un niveau, présentes les 3 semaines, mais portant la mention « /par quinzaine ».",
  "Étiquette contradictoire","Traitées comme HEBDOMADAIRES et PARTAGÉES",
  "Le fait observé (3 semaines sur 3, mêmes groupes) prime sur l’étiquette. Si le rythme est réellement une quinzaine, passer la colonne Rythme à « Quinzaine A »."),
 ("13 séances à GROUPES ALTERNÉS (Espagnol, Italien, Philo, Arabe Najet, Français Nabil, Anglais Khaled, Anglais Imen Ben Lazrak, Anglais Imen Bennour, SVT Sabrine, Éveil Bassma, Anglais Samar)",
  "La séance a lieu CHAQUE semaine, mais les groupes qui y assistent changent d’une semaine à l’autre.",
  "Rythme","ALTERNANCE CONSERVÉE à l’identique",
  "Vue de l’enseignant la séance est hebdomadaire ; vue d’un groupe elle revient une semaine sur deux, d’où le repère ◆ Q-A / ◆ Q-B dans la grille. Le détail des deux jeux de groupes figure dans Planning (colonne Séance partagée avec)."),
 ("Samedi 17:30–19:00 — SVT (Sabrine) et Anglais (Imen Ben Lazrak) en 8éme",
  "Semaine A : SVT pour 8éme (A), Anglais pour 8éme (B)+(C). Semaine B : l’inverse.",
  "Rythme","ALTERNANCE CONSERVÉE, 8éme (C) incluse",
  "8éme (C) fait bien partie du groupe partagé ; l’oublier la priverait de sa séance du samedi."),
 ("COURS EN OPTION — Espagnol (Imen) et Italien (Amel), mardi 19:15–20:45",
  "Les deux séances se chevauchent pour un même groupe : Bac SCE en semaine A, 3éme INFO en semaine B. Présent tel quel dans les 3 semaines d’origine.",
  "Chevauchement volontaire","Marqué « Cours en option » ; EXCLU de la détection de conflit de groupe",
  "Lecture retenue : le groupe se scinde, chaque élève suit l’une OU l’autre langue. Sans ce marquage, le classeur signalerait 14 faux conflits en permanence. Si ce n’est PAS une option, videz la colonne « Cours en option » : le conflit sera alors signalé."),
 ("3éme INFO — STI (Aymen)",
  "Dimanche 15:30–17:30 dans la seule semaine du 21/09, puis lundi 15:30–17:30 dans la semaine de référence.",
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
  "Le même créneau porte « Histoire » en semaine A et « Géographie » en semaine B. Aucune mention de quinzaine. Ces trois séances ne sont PAS partagées : chaque groupe a son propre horaire.",
  "Alternance de matière","Alternance CONSERVÉE",
  "Rythme déduit de l’observation sur 3 semaines. Ajouter la mention de quinzaine dans la source éviterait toute ambiguïté."),
 ("Mardi 19:15–20:45 en 2éme — Anglais (Imen Bennour) et Arabe (Najet)",
  "Le créneau partagé par 2éme Eco + 2éme SCE + 2éme INFO porte l’Anglais en semaine A et l’Arabe en semaine B.",
  "Alternance de matière sur créneau partagé","CONSERVÉE",
  "Deux séances distinctes en quinzaine, l’une en semaine A, l’autre en semaine B, sur le même groupe de trois classes."),
 ("Enseignante « Imen Ghoumem » / « Imen Ghoumeme »",
  "Deux orthographes pour la même personne (1ére (B) et 1ére (A), SVT dimanche).",
  "Orthographe","Septembre laissé INTACT ; octobre unifié en « Imen Ghoumeme »",
  "Les données d’origine ne sont pas modifiées. Corriger la source si « Ghoumeme » est la bonne orthographe."),
 ("Mentions « /par quizaine » et « /Par quizaine »",
  "4 cellules comportent une faute de frappe sur « quinzaine ».",
  "Orthographe","Rythme normalisé dans la colonne Rythme ; texte d’origine conservé en colonne AD",
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
KPI=[('Séances planifiées (total)',f'=COUNTIF({PR_},"Oui")','0'),
     ('dont séances partagées',f'=COUNTIFS({PR_},"Oui",{NB_},">1")','0'),
     ('Lignes groupe-séance',f'=COUNTA({A_})','0'),
     ('Heures d’enseignement (total)',f'=SUMIF({PR_},"Oui",{H_})','0.0'),
     ('Séances de septembre (d’origine)',f'=COUNTIFS({Q_},"Original (septembre)",{PR_},"Oui")','0'),
     ('Séances d’octobre (générées)',f'=COUNTIFS({Q_},"Généré (octobre)",{PR_},"Oui")','0'),
     ('Classes / groupes',f'=COUNTA(LST_Classes)','0'),
     ('Enseignants',f'=COUNTA(LST_Profs)','0'),
     ('Matières distinctes',f'=COUNTA(LST_Matieres)','0'),
     ('Séances à rythme quinzaine',f'=COUNTIFS({P_},"Quinzaine A",{PR_},"Oui")+COUNTIFS({P_},"Quinzaine B",{PR_},"Oui")','0'),
     ('Séances à groupes alternés',f'=COUNTIFS({P_},"Hebdo · groupes alternés",{PR_},"Oui")','0'),
     ('⚠ Anomalies détectées',f'=COUNTIF({S_},"⚠ CONFLIT")+COUNTIF({T_},"⚠ CONFLIT")+COUNTIF({U_},"⚠ DOUBLON")+COUNTIF({Z_},"⚠ INCOHÉRENT")','0')]
x=hrow(x,['Indicateur','Valeur'])
k0=x
for lib,f_,nf in KPI:
    Rc.cell(x,1,lib).font=font(9,True); Rc.cell(x,1).border=BOX; Rc.cell(x,1).alignment=LFT
    c=Rc.cell(x,2,f_); c.font=font(11,True,ACC); c.border=BOX; c.alignment=CTR; c.number_format=nf
    Rc.row_dimensions[x].height=18; x+=1
Rc.conditional_formatting.add(f'B{k0}:B{x-1}', FormulaRule(formula=[f'AND(ROW()={x-1},B{x-1}>0)'], fill=fill(ALERT_BG), font=font(11,True,ALERT_FG)))
x+=1

x=band(x,'VOLUME PAR SEMAINE')
x=hrow(x,['Semaine','Période','Parité','Séances','Heures','dont partagées','dont Quinz. / alternées','dont Ponct.','Classes actives','Enseignants','Moy. h / séance'])
for code,start,par,per,shname,mois in WEEKS:
    x=drow(x,[code,per,par,
      f'=COUNTIFS({A_},"{code}",{PR_},"Oui")', f'=SUMIFS({H_},{A_},"{code}",{PR_},"Oui")',
      f'=COUNTIFS({A_},"{code}",{PR_},"Oui",{NB_},">1")',
      f'=COUNTIFS({A_},"{code}",{PR_},"Oui",{P_},"Quinzaine A")+COUNTIFS({A_},"{code}",{PR_},"Oui",{P_},"Quinzaine B")+COUNTIFS({A_},"{code}",{PR_},"Oui",{P_},"Hebdo · groupes alternés")',
      f'=COUNTIFS({A_},"{code}",{PR_},"Oui",{P_},"Exception")+COUNTIFS({A_},"{code}",{PR_},"Oui",{P_},"Ponctuelle / modifiée")',
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

x=band(x,'LIGNES GROUPE-SÉANCE PAR CYCLE ET PAR SEMAINE (une séance partagée compte pour chacun de ses groupes)')
x=hrow(x,['Cycle','Nb classes']+CODES+['Total'])
for cy,classes in CYCLES:
    x=drow(x,[cy,len(classes)]+[f'=COUNTIFS({A_},"{c}",{I_},$A{x})' for c in CODES]+[f'=SUM(C{x}:I{x})'],bold=False)
x+=1

x=band(x,'PAR MATIÈRE (période complète) — en lignes groupe-séance')
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
      f'=COUNTIFS({A_},"S5",{E_},$A{x},{PR_},"Oui")', f'=SUMIFS({H_},{A_},"S5",{E_},$A{x},{PR_},"Oui")',
      f'=COUNTIFS({A_},"S4",{E_},$A{x},{PR_},"Oui")', f'=SUMIFS({H_},{A_},"S4",{E_},$A{x},{PR_},"Oui")',
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
NAV=[('Planning','Table de référence : 1 ligne = 1 groupe dans une séance. C’est ICI que l’on modifie les données.',f'=COUNTIF({PR_},"Oui")','Saisie'),
     ('Référentiels','Listes des classes, matières, enseignants, salles, modes, rythmes et semaines.','—','Paramètres'),
     ('Contrôles','21 contrôles automatiques + toutes les décisions prises pour construire octobre.','—','Fiabilité'),
     ('Récapitulatif','Indicateurs : volumes par semaine, cycle, matière et jour.','—','Pilotage'),
     ('Vue Enseignants','Charge de chaque enseignant, semaine par semaine (séances et heures).','—','Pilotage')]
for shname,code,per,par,mois in grid_sheets:
    NAV.append((shname,f'Grille imprimable — {per} (semaine {par}).',f'=COUNTIFS({A_},"{code}",{PR_},"Oui")',
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
LEG=([('__TITRE__','COULEUR DES CASES — une couleur par matière','','','')]
   + [(_n,('Écrit dans la case : '+' ou '.join(_m)) if len(_m)>1 else '','','','') for _n,_c,_m in MATIERES]
   + [('__TITRE__','COULEUR DE LA COLONNE CLASSE — cycle','','','')]
   + [('Primaire','4éme, 5éme, 6éme Pilote (A), 6éme (B)','','',''),
      ('Collège','7éme, 8éme et 9éme (8 groupes)','','',''),
      ('Secondaire','1ére, 2éme et 3éme (8 groupes)','','',''),
      ('Bac','Bac Eco, Bac SCE, Bac INFO','','','')]
   + [('__TITRE__','STATUT — couleur PRIORITAIRE, posée dans la colonne Statut de Planning','','','')]
   + [(_n,'Remplace la couleur de la matière ; apparaît dans la case précédé de ⚑','','','') for _n,_c in STATUTS]
   + [('__TITRE__','REPÈRES DANS LE TEXTE','','','')]
   + [('Case fusionnée','SÉANCE PARTAGÉE : elle n’est écrite qu’une fois et sa case couvre toutes les classes réunies sur ce créneau','','',''),
      ('Q-A','Ce groupe a la séance en semaines A : 14/09 · 28/09 · 12/10 · 26/10','','',''),
      ('Q-B','Ce groupe a la séance en semaines B : 21/09 · 05/10 · 19/10','','',''),
      ('OPT','Cours au choix : le groupe se scinde sur ce créneau (Espagnol OU Italien)','','',''),
      ('PONCT.','Séance exceptionnelle, non récurrente (ex. rattrapage)','','',''),
      ('EN LIGNE','Séance à distance (à renseigner dans la colonne Mode de Planning)','','',''),
      ('Case grisée','Aucune séance sur ce créneau','','',''),
      ('Couleur libre','Toute case est recolorable à la main : les couleurs de matière sont posées en dur, pas en mise en forme conditionnelle','','',''),
      ('⚠ en rouge','Conflit horaire, doublon ou anomalie détectée automatiquement','','','')])
FAM_COL={_n:_c for _n,_c,_m in MATIERES}
FAM_COL.update({_n:_c for _n,_c in STATUTS})
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
HOW=[('1','Modifier une séance existante','Dans la feuille Planning : changer l’horaire, la matière, l’enseignant, la salle ou le mode. La grille de la semaine et tous les indicateurs se mettent à jour SEULS (chaque case va chercher sa séance par son ID).'),
     ('2','Ajouter ou supprimer une séance','Ajouter la ligne dans Planning (recopier les formules des colonnes H et X à AD depuis la ligne du dessus). ATTENTION : la STRUCTURE des grilles (nombre de lignes par classe et cellules fusionnées) est figée à la construction — une séance ajoutée n’apparaîtra pas toute seule dans la grille. Demandez-moi de régénérer le classeur, ou reportez-la à la main.'),
     ('3','Renseigner une salle ou une séance en ligne','Colonnes Salle et Mode de Planning (menus déroulants). L’indication apparaît aussitôt dans la grille de la semaine.'),
     ('3 bis','Mettre une séance en évidence','Deux moyens. (a) Colonne Statut de Planning (Annulé, Examen, Rattrapage…) : la case prend la couleur du statut, précédée de ⚑, et cela survit à une régénération du classeur. (b) Colorer la case directement dans la grille : c’est possible partout, les couleurs de matière étant posées en dur. Un statut l’emporte toujours sur une couleur posée à la main.'),
     ('4','Vérifier la fiabilité','Ouvrir la feuille Contrôles : la colonne Statut doit afficher « OK » partout.'),
     ('5','Imprimer ou envoyer en PDF','Chaque grille est déjà paramétrée : A4 paysage, 1 page de large, saut de page à chaque changement de cycle, en-têtes répétés. Fichier ▸ Exporter au format PDF.'),
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
      ("Les 301 séances de septembre — soit 370 lignes groupe-séance — sont conservées à l’identique, y compris les fautes de frappe d’origine (colonne « Texte d’origine » de Planning)."),
      ("Les séances partagées entre plusieurs groupes (Physique de Melek, Espagnol, Italien, Français de Nassima…) sont conservées comme telles : un groupe par ligne dans Planning, reliés par un même ID séance, et une seule cellule fusionnée dans la grille."),
      ("17 décisions et points ambigus sont détaillés en bas de la feuille Contrôles. Trois méritent votre arbitrage : les Physique « /par quinzaine » de Melek, le STI du lundi en 3éme INFO, et le jeudi 15/10 (jour férié)."),
      ("Chaque matière a sa propre couleur (17 teintes, liste en légende), posée EN DUR dans les cellules : vous pouvez donc recolorer n’importe quelle case à la main. Revers de la médaille : si vous changez la matière d’une séance dans Planning, le texte suit mais PAS la couleur — régénérez le classeur, ou corrigez la couleur vous-même."),
      ("Les grilles sont calculées par formules : à la première ouverture, laisser Excel recalculer (ou appuyer sur F9). Leur structure (lignes et fusions) est en revanche figée : voir le point 2 du mode d’emploi.")]
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
