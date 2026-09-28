# -*- coding: utf-8 -*-
"""Une couleur par matière. Contraintes : Math = bleu, Physique = jaune (demande),
fond clair pour du texte gras NOIR (contraste >= 8:1), teintes réellement distinguables."""
import colorsys, itertools, json, random

SUBJ=[('Math',['Math','رياضيات']), ('Physique',['Physique','فيزياء']),
      ('Français',['Français']), ('Arabe',['Arabe','عربية']), ('Anglais',['Anglais']),
      ('SVT',['SVT']), ('Éveil scientifique',['Éveil scientifique','إيقاظ علمي']),
      ('Histoire',['Histoire']), ('Géographie',['Géographie']), ('Philo',['Philo']),
      ('Eco',['Eco']), ('Gestion',['Gestion']), ('Informatique',['Informatique']),
      ('ALGO',['ALGO']), ('STI',['STI']), ('Espagnol',['Espagnol']), ('Italien',['Italien'])]
N=len(SUBJ); STEP=360.0/N
HUES=[(218.0-k*STEP)%360 for k in range(N)]      # slot 0 = bleu (Math), slot 8 ≈ 48° = jaune
TIERS=[(0.875,0.62),(0.815,0.80),(0.755,0.70)]

def hexa(h,l,s):
    r,g,b=colorsys.hls_to_rgb(h/360.0,l,s)
    return '%02X%02X%02X'%(round(r*255),round(g*255),round(b*255))
def _lin(u): return u/12.92 if u<=0.04045 else ((u+0.055)/1.055)**2.4
def lab(hx):
    r,g,b=[_lin(int(hx[i:i+2],16)/255) for i in (0,2,4)]
    X=(0.4124*r+0.3576*g+0.1805*b)/0.95047; Y=0.2126*r+0.7152*g+0.0722*b; Z=(0.0193*r+0.1192*g+0.9505*b)/1.08883
    def k(t): return t**(1/3) if t>0.008856 else 7.787*t+16/116
    X,Y,Z=k(X),k(Y),k(Z); return (116*Y-16,500*(X-Y),200*(Y-Z))
def contrast(hx):
    r,g,b=[_lin(int(hx[i:i+2],16)/255) for i in (0,2,4)]
    return (0.2126*r+0.7152*g+0.0722*b+0.05)/0.05

HEX=[[hexa(HUES[s],*TIERS[t]) for t in range(3)] for s in range(N)]   # pré-calcul
LAB=[[lab(HEX[s][t]) for t in range(3)] for s in range(N)]
CON=[[contrast(HEX[s][t]) for t in range(3)] for s in range(N)]
OK =[[CON[s][t]>=8.0 for t in range(3)] for s in range(N)]

def mind(sl,ti):
    m=1e9
    for i,j in itertools.combinations(range(N),2):
        a,b=LAB[sl[i]][ti[i]],LAB[sl[j]][ti[j]]
        d=(a[0]-b[0])**2+(a[1]-b[1])**2+(a[2]-b[2])**2
        if d<m: m=d
    return m**0.5

FIXED={0:0, 1:8}                                  # Math -> slot bleu, Physique -> slot jaune
movable=[i for i in range(N) if i not in FIXED]
freeslots=[s for s in range(N) if s not in FIXED.values()]
random.seed(11); best=(-1,None,None)
for _ in range(60):
    sl=[0]*N; sl[0],sl[1]=0,8
    sh=freeslots[:]; random.shuffle(sh)
    for k,i in enumerate(movable): sl[i]=sh[k]
    ti=[random.choice([t for t in range(3) if OK[sl[i]][t]] or [0]) for i in range(N)]
    ti[0]=2; ti[1]=2          # Math et Physique en teinte franche (bleu net, jaune net)
    cur=mind(sl,ti)
    for _ in range(4000):
        if random.random()<0.5:
            a,b=random.sample(movable,2); sl[a],sl[b]=sl[b],sl[a]
            if not(OK[sl[a]][ti[a]] and OK[sl[b]][ti[b]]): sl[a],sl[b]=sl[b],sl[a]; continue
            n=mind(sl,ti)
            if n>=cur: cur=n
            else: sl[a],sl[b]=sl[b],sl[a]
        else:
            a=random.randrange(N)
            if a in (0,1): continue          # teintes de Math et Physique verrouillées
            old=ti[a]; t=random.randrange(3)
            if not OK[sl[a]][t]: continue
            ti[a]=t; n=mind(sl,ti)
            if n>=cur: cur=n
            else: ti[a]=old
    if cur>best[0]: best=(cur,sl[:],ti[:])

sc,sl,ti=best
cols=[HEX[sl[i]][ti[i]] for i in range(N)]
print(f"{'matière':<20} {'couleur':<9} {'teinte':>7}  contraste")
for i,(n,lb) in enumerate(SUBJ):
    print(f"  {n:<18} #{cols[i]}   {HUES[sl[i]]:5.0f}°   {contrast(cols[i]):5.1f}:1")
print(f"\ncontraste minimum avec le texte noir : {min(contrast(c) for c in cols):.1f}:1")
L=[lab(c) for c in cols]
pairs=sorted((sum((a-b)**2 for a,b in zip(L[i],L[j]))**0.5,SUBJ[i][0],SUBJ[j][0])
             for i,j in itertools.combinations(range(N),2))
print("\n5 paires les plus proches :")
for d,a,b in pairs[:5]: print(f"  ΔE={d:5.1f}  {a} / {b}")
print(f"\nΔE minimum = {pairs[0][0]:.1f}   (seuil : 10 = nettement distinguable)")
json.dump({SUBJ[i][0]:[cols[i],SUBJ[i][1]] for i in range(N)},
          open('palette.json','w',encoding='utf-8'),ensure_ascii=False,indent=1)
