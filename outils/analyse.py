# -*- coding: utf-8 -*-
import json, itertools
from collections import defaultdict
d=json.load(open('parsed.json',encoding='utf-8'))
sheets=list(d.keys())
W=['S1 (14-20/09)','S2 (21-27/09)','S3 (28/09-04/10)']
def fmt(m): return f"{m//60:02d}:{m%60:02d}" if m is not None else "??"

# --- 1. conflits enseignants (même prof, même jour, chevauchement) ---
print("### CONFLITS ENSEIGNANTS (chevauchement horaire)")
for si,sh in enumerate(sheets):
    byt=defaultdict(list)
    for s in d[sh]['sessions']:
        if s['teacher'] and s['start'] is not None:
            byt[(s['teacher'],s['day'])].append(s)
    for (t,day),lst in sorted(byt.items()):
        for a,b in itertools.combinations(lst,2):
            if a['start']<b['end'] and b['start']<a['end']:
                q = ' [les 2 par quinzaine]' if (a['quinzaine'] and b['quinzaine']) else (' [1 par quinzaine]' if (a['quinzaine'] or b['quinzaine']) else ' *** LES 2 FIXES ***')
                print(f" {W[si]} {t} {day}: {a['classe']} {fmt(a['start'])}-{fmt(a['end'])} {a['subject']} ({a['cell']}) VS {b['classe']} {fmt(b['start'])}-{fmt(b['end'])} {b['subject']} ({b['cell']}){q}")

# --- 2. conflits classe (même classe, chevauchement) ---
print("\n### CONFLITS CLASSE (2 séances superposées pour la même classe)")
for si,sh in enumerate(sheets):
    byc=defaultdict(list)
    for s in d[sh]['sessions']:
        if s['start'] is not None: byc[(s['classe'],s['day'])].append(s)
    for (c,day),lst in sorted(byc.items()):
        for a,b in itertools.combinations(lst,2):
            if a['start']<b['end'] and b['start']<a['end']:
                print(f" {W[si]} {c} {day}: {fmt(a['start'])}-{fmt(a['end'])} {a['subject']} VS {fmt(b['start'])}-{fmt(b['end'])} {b['subject']}")

# --- 3. présence par (classe, jour, matière) sur les 3 semaines ---
print("\n### RYTHME PAR (CLASSE, JOUR, MATIERE)  [S1 S2 S3] + marquage quinzaine")
keys=set()
pres={}
for si,sh in enumerate(sheets):
    for s in d[sh]['sessions']:
        k=(s['classe'],s['day'],s['subject'])
        keys.add(k)
        pres.setdefault(k,{})[si]=s
order={c['classe']:i for i,c in enumerate(d[sheets[0]]['blocks'])}
DAYO={x:i for i,x in enumerate(['Lundi','Mardi','Mercredi','Jeudi','Vendredi','Samedi','Dimanche'])}
for k in sorted(keys, key=lambda k:(order.get(k[0],99), DAYO[k[1]], k[2])):
    p=pres[k]
    pat=''.join('X' if i in p else '.' for i in range(3))
    qs=''.join('Q' if (i in p and p[i]['quinzaine']) else ('-' if i in p else ' ') for i in range(3))
    times=' / '.join(f"{fmt(p[i]['start'])}-{fmt(p[i]['end'])}({p[i]['teacher']})" if i in p else '—' for i in range(3))
    flag=''
    if pat!='XXX': flag=' <<<'
    print(f" {pat} q={qs} | {k[0]:<16} {k[1]:<9} {k[2]:<20} | {times}{flag}")
