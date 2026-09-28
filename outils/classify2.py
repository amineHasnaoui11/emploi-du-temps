# -*- coding: utf-8 -*-
import json, itertools
from collections import defaultdict
d=json.load(open('parsed2.json',encoding='utf-8'))
sheets=list(d.keys()); W=['S1','S2','S3']
DAYS=['Lundi','Mardi','Mercredi','Jeudi','Vendredi','Samedi','Dimanche']
CYCLES=[('Primaire',['4éme','5éme','6éme Pilote (A)','6éme (B)']),
 ('Collège',['7éme (A)','7éme (B)','8éme (A)','8éme (B)','8éme (C)','9éme (A)','9éme (B)','9éme (C)']),
 ('Secondaire',['1ére (A)','1ére (B)','2éme Eco','2éme SCE','2éme INFO','3éme Eco','3éme INFO','3éme SCE']),
 ('Bac',['Bac Eco','Bac SCE','Bac INFO'])]
ORDER=[c for _,l in CYCLES for c in l]
def fmt(m): return f"{m//60:02d}:{m%60:02d}"
def key1(s): return (s['day'],s['subject'],s['teacher'],tuple(sorted(s['classes'],key=ORDER.index)))

pres=defaultdict(dict)
for i,sh in enumerate(sheets):
    for s in d[sh]['sessions']: pres[key1(s)][i]=s

# --- fusion des séances à GROUPES ALTERNÉS (même jour/matière/prof/horaire, groupes disjoints, présences complémentaires) ---
bysubj=defaultdict(list)
for k in pres: bysubj[(k[0],k[1],k[2])].append(k)
ALT={}          # keyA -> keyB
used=set()
for kk,lst in bysubj.items():
    for a,b in itertools.combinations(lst,2):
        if a in used or b in used: continue
        pa,pb=set(pres[a]),set(pres[b])
        sa,sb=set(a[3]),set(b[3])
        ta=pres[a][max(pa)]['start']; tb=pres[b][max(pb)]['start']
        if ta==tb and not (sa & sb) and pa=={0,2} and pb=={1}:
            ALT[a]=b; used|={a,b}
        elif ta==tb and not (sa & sb) and pb=={0,2} and pa=={1}:
            ALT[b]=a; used|={a,b}

FIXE=[]; QA=[]; QB=[]; EXCL=[]; FLAGS=[]
cs=defaultdict(lambda: defaultdict(set))     # (classe,matière) -> semaine -> jours
for i,sh in enumerate(sheets):
    for s in d[sh]['sessions']:
        for c in s['classes']: cs[(c,s['subject'])][i].add(s['day'])

def entry(s, classes, classesB=None):
    e=dict(s); e['classes']=list(classes)
    if classesB is not None: e['classesB']=list(classesB)
    return e

for k in list(pres):
    if k in used and k not in ALT: continue          # c'est le volet B d'une alternance
    p=set(pres[k]); pat=''.join('X' if i in p else '.' for i in range(3))
    if k in ALT:
        kb=ALT[k]; sa=pres[k][2]; sb=pres[kb][1]
        FIXE.append(entry(sa, k[3], kb[3]))
        FLAGS.append(('GROUPES_ALTERNES',k[:3],
            f"semaine A -> {', '.join(k[3])} | semaine B -> {', '.join(kb[3])} ({fmt(sa['start'])}-{fmt(sa['end'])})"))
        continue
    day,subj,teach,classes=k
    if 2 in p:
        s=pres[k][2]
        if pat=='XXX':
            FIXE.append(entry(s,classes))
            if s['quinzaine']: FLAGS.append(('LABEL_QUINZAINE_MAIS_HEBDO',k[:3],s['raw']))
        elif pat=='X.X': QA.append(entry(s,classes))
        elif pat=='.XX': FIXE.append(entry(s,classes)); FLAGS.append(('NOUVEAU_DEPUIS_S2',k[:3],s['raw']))
        elif pat=='..X':
            old=set()
            for c in classes:
                old |= (cs[(c,subj)].get(0,set())|cs[(c,subj)].get(1,set()))-{day}
            if old:
                FIXE.append(entry(s,classes)); FLAGS.append(('RELOCALISATION',k[:3],f"{sorted(old)} -> {day}"))
            else:
                FIXE.append(entry(s,classes)); FLAGS.append(('AMBIGU_S3_SEUL',k[:3],s['raw']))
        else: FIXE.append(entry(s,classes))
    else:
        s=pres[k].get(1) or pres[k].get(0)
        if 'rattrapage' in (s['note'] or '').lower(): EXCL.append((pat,k[:3],'exception ponctuelle (Rattrapage)')); continue
        if pat=='.X.':
            later=set()
            for c in classes: later |= cs[(c,subj)].get(2,set())
            if later:
                EXCL.append((pat,k[:3],f"restructuré en S3 -> {sorted(later)}"))
                FLAGS.append(('RESTRUCTURE_S3',k[:3],f"S2 {day} -> S3 {sorted(later)}")); continue
            QB.append(entry(s,classes))
        elif pat in ('XX.','X..'):
            later=set()
            for c in classes: later |= cs[(c,subj)].get(2,set())
            EXCL.append((pat,k[:3],f"absent de la semaine de référence ; en S3 -> {sorted(later) or 'NULLE PART'}"))
        else: EXCL.append((pat,k[:3],'autre'))

def lines(lst): return sum(len(s['classes']) for s in lst)
print(f"FIXE={len(FIXE)} (dont {sum(1 for s in FIXE if 'classesB' in s)} à groupes alternés)  QA={len(QA)}  QB={len(QB)}  EXCL={len(EXCL)}")
nA=len(FIXE)+len(QA); nB=len(FIXE)+len(QB)
lA=sum(len(s['classes']) for s in FIXE)+lines(QA)
lB=sum(len(s.get('classesB',s['classes'])) for s in FIXE)+lines(QB)
print(f"semaine A : {nA} séances / {lA} lignes classe-séance   (réf. S3 = 101 / 125)")
print(f"semaine B : {nB} séances / {lB} lignes classe-séance   (S2 = 100 / 121, moins rattrapage)")
print("\n--- SÉANCES PARTAGÉES RECONDUITES ---")
for s in sorted([x for x in FIXE+QA+QB if len(x['classes'])>1],key=lambda s:(DAYS.index(s['day']),s['start'])):
    alt=f"   [sem. B -> {', '.join(s['classesB'])}]" if 'classesB' in s else ''
    print(f"  {s['day']:<9} {fmt(s['start'])}-{fmt(s['end'])} {s['subject']:<20} {s['teacher']:<16} -> {', '.join(s['classes'])}{alt}")
print("\n--- EXCLUES ---")
for p,k,r in EXCL: print(f"  [{p}] {k} :: {r}")
print("\n--- POINTS SIGNALÉS ---")
for t,k,r in FLAGS: print(f"  [{t}] {k} :: {r}")
json.dump({'FIXE':FIXE,'QA':QA,'QB':QB,'EXCL':[[p,list(k),r] for p,k,r in EXCL],
           'FLAGS':[[t,list(k),r] for t,k,r in FLAGS]},open('model2.json','w',encoding='utf-8'),ensure_ascii=False,indent=1)
