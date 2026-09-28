# -*- coding: utf-8 -*-
import json, datetime
from collections import defaultdict
d=json.load(open('parsed2.json',encoding='utf-8'))
m=json.load(open('model2.json',encoding='utf-8'))
sheets=list(d.keys())
DAYS=['Lundi','Mardi','Mercredi','Jeudi','Vendredi','Samedi','Dimanche']
CYCLES=[('Primaire',['4éme','5éme','6éme Pilote (A)','6éme (B)']),
 ('Collège',['7éme (A)','7éme (B)','8éme (A)','8éme (B)','8éme (C)','9éme (A)','9éme (B)','9éme (C)']),
 ('Secondaire',['1ére (A)','1ére (B)','2éme Eco','2éme SCE','2éme INFO','3éme Eco','3éme INFO','3éme SCE']),
 ('Bac',['Bac Eco','Bac SCE','Bac INFO'])]
CYC={c:cy for cy,l in CYCLES for c in l}
ORDER=[c for _,l in CYCLES for c in l]
WEEKS=[('S1',datetime.date(2026,9,14),'A','du 14 au 20 septembre 2026','S1 14-20 sept','Septembre'),
 ('S2',datetime.date(2026,9,21),'B','du 21 au 27 septembre 2026','S2 21-27 sept','Septembre'),
 ('S3',datetime.date(2026,9,28),'A','du 28 septembre au 4 octobre 2026','S3 28 sept-04 oct','Référence'),
 ('S4',datetime.date(2026,10,5),'B','du 5 au 11 octobre 2026','S4 05-11 oct','Octobre'),
 ('S5',datetime.date(2026,10,12),'A','du 12 au 18 octobre 2026','S5 12-18 oct','Octobre'),
 ('S6',datetime.date(2026,10,19),'B','du 19 au 25 octobre 2026','S6 19-25 oct','Octobre'),
 ('S7',datetime.date(2026,10,26),'A','du 26 octobre au 1er novembre 2026','S7 26 oct-01 nov','Octobre')]
AR={'عربية':'Arabe','رياضيات':'Math','إيقاظ علمي':'Éveil scientifique','فيزياء':'Physique'}
TFIX={'Imen Ghoumem':'Imen Ghoumeme'}
FERIES={datetime.date(2026,10,15):"Jour férié (Évacuation) — à confirmer"}

def srt(cl): return sorted(cl,key=ORDER.index)
def k4(day,subj,teach,cl): return (day,subj,teach,tuple(srt(cl)))

# rythme de chaque combinaison (jour, matière, prof, groupes)
RYT={}
for s in m['FIXE']:
    if 'classesB' in s:
        RYT[k4(s['day'],s['subject'],s['teacher'],s['classes'])]='Hebdo · groupes alternés'
        RYT[k4(s['day'],s['subject'],s['teacher'],s['classesB'])]='Hebdo · groupes alternés'
    else: RYT[k4(s['day'],s['subject'],s['teacher'],s['classes'])]='Hebdomadaire'
for s in m['QA']: RYT[k4(s['day'],s['subject'],s['teacher'],s['classes'])]='Quinzaine A'
for s in m['QB']: RYT[k4(s['day'],s['subject'],s['teacher'],s['classes'])]='Quinzaine B'

rows=[]; sid=defaultdict(int)
def add(wk, s, classes, origine, rythme, note, src, brut):
    code,start,par,per,shname,mois=wk
    date=start+datetime.timedelta(days=DAYS.index(s['day']))
    teach=s['teacher'] or ''
    if origine.startswith('Généré'): teach=TFIX.get(teach,teach)
    sid[code]+=1; sident=f"{code}-{sid[code]:03d}"
    cl=srt(classes)
    notes=[n for n in [note] if n]
    if origine.startswith('Généré') and date in FERIES: notes.append(FERIES[date])
    for i,c in enumerate(cl):
        rows.append(dict(semaine=code,periode=per,parite=par,date=date.isoformat(),jour=s['day'],
            debut=f"{s['start']//60:02d}:{s['start']%60:02d}", fin=f"{s['end']//60:02d}:{s['end']%60:02d}",
            cycle=CYC[c], classe=c, groupes=' + '.join(cl), nbgr=len(cl),
            partagee='Oui' if len(cl)>1 else 'Non',
            principale='Oui' if i==0 else '', seance=sident,
            matiere=s['subject_raw'], matiere_n=AR.get(s['subject_raw'],s['subject_raw']),
            prof=teach, rythme=rythme, origine=origine, notes=' | '.join(notes),
            src=src, brut=brut))

# ---- septembre : verbatim, groupes d'origine conservés ----
for wi,sh in enumerate(sheets):
    for s in d[sh]['sessions']:
        k=k4(s['day'],s['subject'],s['teacher'],s['classes'])
        if 'rattrapage' in (s['note'] or '').lower(): r='Exception'
        else: r=RYT.get(k,'Ponctuelle / modifiée')
        add(WEEKS[wi], s, s['classes'], 'Original (septembre)', r, s['note'], f"{sh.strip()}!{s['cell']}", s['raw'])

# ---- octobre : généré ----
for wk in WEEKS[3:]:
    A = wk[2]=='A'
    for s in m['FIXE']:
        cl = s['classes'] if (A or 'classesB' not in s) else s['classesB']
        r = 'Hebdo · groupes alternés' if 'classesB' in s else 'Hebdomadaire'
        add(wk, s, cl, 'Généré (octobre)', r, '', '— généré —', '')
    for s in (m['QA'] if A else m['QB']):
        add(wk, s, s['classes'], 'Généré (octobre)', 'Quinzaine A' if A else 'Quinzaine B', '', '— généré —', '')

# --- détection automatique des COURS EN OPTION ---------------------------------
# Deux séances qui se chevauchent pour un MÊME groupe, déjà présentes dans les données
# d'origine de septembre, ne sont pas un conflit : le groupe se scinde (ex. LV3 Espagnol
# OU Italien). On repère ces couples sur septembre, puis on marque les mêmes séances
# partout, octobre compris.
import itertools as _it
_bycd=defaultdict(list)
for r_ in rows:
    if r_['origine'].startswith('Original'):
        _bycd[(r_['semaine'],r_['classe'],r_['jour'])].append(r_)
def _mn(t): h,m=t.split(':'); return int(h)*60+int(m)
OPT=set()
for _k,_l in _bycd.items():
    for a,b in _it.combinations(_l,2):
        if a['seance']==b['seance']: continue
        if _mn(a['debut'])<_mn(b['fin']) and _mn(b['debut'])<_mn(a['fin']):
            for z in (a,b): OPT.add((z['jour'],z['matiere'],z['prof'],z['debut']))
for r_ in rows:
    r_['option']='Oui' if (r_['jour'],r_['matiere'],r_['prof'],r_['debut']) in OPT else ''
print("séances en option détectées :", len(OPT), "->", sorted({(o[1],o[2]) for o in OPT}))
print("lignes marquées 'option' :", sum(1 for r_ in rows if r_['option']=='Oui'))

rows.sort(key=lambda r:(r['semaine'], r['date'], ORDER.index(r['classe']), r['debut']))
print(f"TOTAL lignes classe-séance : {len(rows)}")
for code,start,par,per,sn,mo in WEEKS:
    rr=[r for r in rows if r['semaine']==code]
    ns=len({r['seance'] for r in rr}); sh=len({r['seance'] for r in rr if r['nbgr']>1})
    print(f"  {code} parité {par} : {ns:>3} séances | {len(rr):>3} lignes | {sh:>2} partagées")
json.dump({'rows':rows,'weeks':[[c,s.isoformat(),p,pe,sn,mo] for c,s,p,pe,sn,mo in WEEKS],
           'cycles':CYCLES,'class_order':ORDER},open('rows2.json','w',encoding='utf-8'),ensure_ascii=False,indent=1)
