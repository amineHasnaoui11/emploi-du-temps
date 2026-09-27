# -*- coding: utf-8 -*-
import json
from collections import defaultdict, Counter
d=json.load(open('parsed.json',encoding='utf-8'))
sheets=list(d.keys()); W=['S1','S2','S3']
def fmt(m): return f"{m//60:02d}:{m%60:02d}" if m is not None else "??"

print("### DOUBLONS exacts (classe/jour/matiere/heure identiques dans la meme semaine)")
for si,sh in enumerate(sheets):
    c=Counter((s['classe'],s['day'],s['subject'],s['start'],s['end']) for s in d[sh]['sessions'])
    for k,v in c.items():
        if v>1: print(" ",W[si],k,v)
print(" (aucun si vide)")

print("\n### CELLULES NON PARSEES (heure ou prof manquant)")
for si,sh in enumerate(sheets):
    for s in d[sh]['sessions']:
        if s['start'] is None or s['end'] is None or not s['teacher']:
            print(f"  {W[si]} {s['cell']} {s['classe']}: {s['raw']!r}")

print("\n### NOTES / MENTIONS LIBRES")
for si,sh in enumerate(sheets):
    for s in d[sh]['sessions']:
        if s['note']: print(f"  {W[si]} {s['cell']} {s['classe']} {s['day']}: note={s['note']!r} | {s['raw']!r}")

print("\n### VARIANTES D'ECRITURE DU MARQUEUR QUINZAINE")
import re
pats=Counter()
for sh in sheets:
    for s in d[sh]['sessions']:
        m=re.search(r'/\s*[Pp]ar\s*\S+', s['raw'])
        if m: pats[m.group(0)]+=1
for k,v in pats.most_common(): print(f"  {k!r}: {v}")

print("\n### MATIERES (libelles bruts)")
for k,v in sorted(Counter(s['subject_raw'] for sh in sheets for s in d[sh]['sessions']).items()): print(f"  {k!r}: {v}")

print("\n### ENSEIGNANTS")
for k,v in sorted(Counter(s['teacher'] for sh in sheets for s in d[sh]['sessions'] if s['teacher']).items()): print(f"  {k!r}: {v}")

print("\n### CLASSES (ordre + libelles par feuille)")
for si,sh in enumerate(sheets):
    print(" ",W[si], [b['classe'] for b in d[sh]['blocks']])

print("\n### PLAGES HORAIRES distinctes")
print(sorted({(fmt(s['start']),fmt(s['end'])) for sh in sheets for s in d[sh]['sessions'] if s['start']}))
print("\n### DUREES")
print(sorted(Counter((s['end']-s['start']) for sh in sheets for s in d[sh]['sessions'] if s['start'] and s['end']).items()))
