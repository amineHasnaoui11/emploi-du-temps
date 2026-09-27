# -*- coding: utf-8 -*-
import json, datetime
from collections import defaultdict
d=json.load(open('parsed.json',encoding='utf-8'))
sheets=list(d.keys())
DAYS=['Lundi','Mardi','Mercredi','Jeudi','Vendredi','Samedi','Dimanche']
def fmt(m): return f"{m//60:02d}:{m%60:02d}"

# max séances par (classe, jour) dans une semaine
mx=0
for sh in sheets:
    c=defaultdict(int)
    for s in d[sh]['sessions']: c[(s['classe'],s['day'])]+=1
    mx=max(mx,max(c.values()))
print("Max séances par (classe,jour,semaine) =", mx)

pres=defaultdict(dict)
for i,sh in enumerate(sheets):
    for s in d[sh]['sessions']:
        pres[(s['classe'],s['day'],s['subject'])][i]=s
def pat(k): return ''.join('X' if i in pres[k] else '.' for i in range(3))

# index (classe, matière) -> jours par semaine
cs=defaultdict(lambda: defaultdict(set))
for i,sh in enumerate(sheets):
    for s in d[sh]['sessions']:
        cs[(s['classe'],s['subject'])][i].add(s['day'])

FIXE=[]; QA=[]; QB=[]; EXCL=[]; FLAGS=[]
S1,S2,S3=0,1,2
for k in pres:
    p=pat(k); classe,day,subj=k
    if S3 in pres[k]:
        s=pres[k][S3]
        if p=='XXX':
            FIXE.append(s)
            if s['quinzaine']:
                FLAGS.append(('LABEL_QUINZAINE_MAIS_HEBDO',k,s['raw']))
        elif p=='X.X': QA.append(s)
        elif p=='.XX': FIXE.append(s); FLAGS.append(('NOUVEAU_DEPUIS_S2_TRAITE_HEBDO',k,s['raw']))
        elif p=='..X':
            # relocalisation ? même (classe,matière) présente un autre jour en S1 et S2
            prev = cs[(classe,subj)]
            old_days = (prev.get(S1,set()) | prev.get(S2,set())) - {day}
            if prev.get(S1) and prev.get(S2) and old_days:
                FIXE.append(s); FLAGS.append(('RELOCALISATION_HEBDO',k,f"ancien(s) jour(s)={sorted(old_days)} -> {day}"))
            elif prev.get(S2) and old_days:
                FIXE.append(s); FLAGS.append(('AMBIGU_S3_SEUL_TRAITE_HEBDO',k,f"vu en S2 le {sorted(old_days)}, en S3 le {day}"))
            else:
                FIXE.append(s); FLAGS.append(('AMBIGU_S3_SEUL_TRAITE_HEBDO',k,s['raw']))
    else:
        s = pres[k].get(S2) or pres[k].get(S1)
        if p=='.X.':
            if 'rattrapage' in (s['note'] or '').lower():
                EXCL.append((p,k,'exception ponctuelle (Rattrapage)')); continue
            # restructuré ? la (classe,matière) est devenue hebdo en S3 un autre jour
            if cs[(classe,subj)].get(S3):
                EXCL.append((p,k,f"restructuré en S3 -> {sorted(cs[(classe,subj)][S3])}"))
                FLAGS.append(('RESTRUCTURE_S3',k,f"S2 {day} {fmt(s['start'])} -> S3 {sorted(cs[(classe,subj)][S3])}"))
                continue
            QB.append(s)
        elif p=='XX.':
            EXCL.append((p,k,f"hebdo S1+S2 mais absent S3; (classe,matière) en S3 -> {sorted(cs[(classe,subj)].get(S3,set())) or 'NULLE PART'}"))
        elif p=='X..':
            EXCL.append((p,k,f"S1 seulement; en S3 -> {sorted(cs[(classe,subj)].get(S3,set())) or 'NULLE PART'}"))
        elif p=='X.X': QA.append(s)  # ne devrait pas arriver
        else: EXCL.append((p,k,'autre'))

print(f"\nFIXE={len(FIXE)}  QUINZAINE_A={len(QA)}  QUINZAINE_B={len(QB)}  EXCLUES={len(EXCL)}")
print(f"=> semaine A = {len(FIXE)+len(QA)} séances ; semaine B = {len(FIXE)+len(QB)} séances")
print("\n--- QUINZAINE A (semaines A : 28/09, 12/10, 26/10) ---")
for s in sorted(QA,key=lambda s:(DAYS.index(s['day']),s['start'])): print(f"  {s['day']:<9} {fmt(s['start'])}-{fmt(s['end'])} {s['classe']:<16} {s['subject']:<20} {s['teacher']}")
print("\n--- QUINZAINE B (semaines B : 21/09, 05/10, 19/10) ---")
for s in sorted(QB,key=lambda s:(DAYS.index(s['day']),s['start'])): print(f"  {s['day']:<9} {fmt(s['start'])}-{fmt(s['end'])} {s['classe']:<16} {s['subject']:<20} {s['teacher']}")
print("\n--- EXCLUES (non reconduites en octobre) ---")
for p,k,r in EXCL: print(f"  [{p}] {k[0]} {k[1]} {k[2]} :: {r}")
print("\n--- POINTS A SIGNALER ---")
for t,k,r in FLAGS: print(f"  [{t}] {k[0]} {k[1]} {k[2]} :: {r}")
json.dump({'FIXE':FIXE,'QA':QA,'QB':QB,
           'EXCL':[[p,list(k),r] for p,k,r in EXCL],
           'FLAGS':[[t,list(k),r] for t,k,r in FLAGS]},
          open('model.json','w',encoding='utf-8'),ensure_ascii=False,indent=1)
