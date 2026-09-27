# -*- coding: utf-8 -*-
import json, datetime
from collections import defaultdict
d=json.load(open('parsed.json',encoding='utf-8'))
m=json.load(open('model.json',encoding='utf-8'))
sheets=list(d.keys())
DAYS=['Lundi','Mardi','Mercredi','Jeudi','Vendredi','Samedi','Dimanche']

CYCLES=[('Primaire',['4éme','5éme','6éme Pilote (A)','6éme (B)']),
 ('Collège',['7éme (A)','7éme (B)','8éme (A)','8éme (B)','8éme (C)','9éme (A)','9éme (B)','9éme (C)']),
 ('Secondaire',['1ére (A)','1ére (B)','2éme Eco','2éme SCE','2éme INFO','3éme Eco','3éme INFO','3éme SCE']),
 ('Bac',['Bac Eco','Bac SCE','Bac INFO'])]
CYC={c:cy for cy,lst in CYCLES for c in lst}
CLASS_ORDER=[c for cy,lst in CYCLES for c in lst]

WEEKS=[('S1',datetime.date(2026,9,14),'A','du 14 au 20 septembre 2026','S1 14-20 sept','Septembre'),
 ('S2',datetime.date(2026,9,21),'B','du 21 au 27 septembre 2026','S2 21-27 sept','Septembre'),
 ('S3',datetime.date(2026,9,28),'A','du 28 septembre au 4 octobre 2026','S3 28 sept-04 oct','Référence'),
 ('S4',datetime.date(2026,10,5),'B','du 5 au 11 octobre 2026','S4 05-11 oct','Octobre'),
 ('S5',datetime.date(2026,10,12),'A','du 12 au 18 octobre 2026','S5 12-18 oct','Octobre'),
 ('S6',datetime.date(2026,10,19),'B','du 19 au 25 octobre 2026','S6 19-25 oct','Octobre'),
 ('S7',datetime.date(2026,10,26),'A','du 26 octobre au 1er novembre 2026','S7 26 oct-01 nov','Octobre')]

AR={'عربية':'Arabe','رياضيات':'Math','إيقاظ علمي':'Éveil scientifique','فيزياء':'Physique'}
TEACHER_FIX={'Imen Ghoumem':'Imen Ghoumeme'}
FERIES={datetime.date(2026,10,15):"Jour férié (Évacuation) — à confirmer"}

# rythme des lignes de septembre
pres=defaultdict(dict)
for i,sh in enumerate(sheets):
    for s in d[sh]['sessions']: pres[(s['classe'],s['day'],s['subject'])][i]=s
EXCL_KEYS={tuple(k) for _p,k,_r in m['EXCL']}
def rythme_sept(s, wi):
    k=(s['classe'],s['day'],s['subject'])
    if 'rattrapage' in (s['note'] or '').lower(): return 'Exception'
    # séance déplacée ou restructurée dans la semaine de référence : ce n'est pas une quinzaine
    if k in EXCL_KEYS: return 'Ponctuelle / modifiée'
    p=''.join('X' if i in pres[k] else '.' for i in range(3))
    if 'rattrapage' in (s['note'] or '').lower(): return 'Exception'
    return {'XXX':'Hebdomadaire','X.X':'Quinzaine A','.X.':'Quinzaine B',
            '.XX':'Hebdomadaire','..X':'Hebdomadaire'}.get(p,'Ponctuelle / modifiée')

rows=[]
def add(wk, s, origine, rythme, note, src, brut):
    code,start,par,per,shname,mois = wk
    date = start + datetime.timedelta(days=DAYS.index(s['day']))
    teach = s['teacher'] or ''
    if origine.startswith('Généré'): teach = TEACHER_FIX.get(teach, teach)
    notes=[n for n in [note] if n]
    if origine.startswith('Généré') and date in FERIES: notes.append(FERIES[date])
    rows.append(dict(semaine=code, periode=per, parite=par, date=date, jour=s['day'],
        debut=datetime.time(s['start']//60, s['start']%60), fin=datetime.time(s['end']//60, s['end']%60),
        cycle=CYC[s['classe']], classe=s['classe'], matiere=s['subject_raw'], matiere_n=AR.get(s['subject_raw'], s['subject_raw']),
        prof=teach, rythme=rythme, origine=origine, notes=' | '.join(notes), src=src, brut=brut))

# --- septembre : verbatim ---
for wi,sh in enumerate(sheets):
    for s in d[sh]['sessions']:
        add(WEEKS[wi], s, 'Original (septembre)', rythme_sept(s,wi), s['note'], f"{sh.strip()}!{s['cell']}", s['raw'])

# --- octobre : généré ---
for wk in WEEKS[3:]:
    pool = m['FIXE'] + (m['QA'] if wk[2]=='A' else m['QB'])
    for s in pool:
        r = 'Hebdomadaire' if s in m['FIXE'] else ('Quinzaine A' if wk[2]=='A' else 'Quinzaine B')
        add(wk, s, 'Généré (octobre)', r, '', '— généré —', '')

rows.sort(key=lambda r:(r['semaine'], r['date'], CLASS_ORDER.index(r['classe']), r['debut']))
print("TOTAL lignes:", len(rows))
for code,start,par,per,shname,mois in WEEKS:
    n=sum(1 for r in rows if r['semaine']==code)
    print(f"  {code} {per:<40} parité {par}  -> {n} séances")
json.dump({'rows':[{**r,'date':r['date'].isoformat(),'debut':r['debut'].strftime('%H:%M'),'fin':r['fin'].strftime('%H:%M')} for r in rows],
           'weeks':[[c,s.isoformat(),p,pe,sn,mo] for c,s,p,pe,sn,mo in WEEKS],
           'cycles':CYCLES,'class_order':CLASS_ORDER},
          open('rows.json','w',encoding='utf-8'),ensure_ascii=False,indent=1)
