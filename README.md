# Emploi du temps — Année scolaire 2026-2027

Refonte du planning de septembre 2026 et construction du planning complet d'octobre 2026.

## Contenu du dépôt

| Dossier | Fichier | Rôle |
|---|---|---|
| `source/` | `RS_Septembre_2026-2027_ORIGINAL.xlsx` | Fichier d'origine, **strictement intact** (md5 identique à l'envoi) |
| `livrables/` | `RS_2026-2027_Planning_Septembre-Octobre.xlsx` | **Le classeur complet** (septembre + octobre, 13 feuilles) |
| `livrables/` | `RS_2026-2027_Planning_Septembre-Octobre.pdf` | Son export PDF (44 pages, A4) |
| `livrables/` | `RS_2026-2027_Octobre_seul.xlsx` | **Octobre seul** : 4 semaines, rien d'autre, contenu figé |
| `livrables/` | `RS_2026-2027_Octobre_seul.pdf` | Son export PDF (16 pages, A4) |
| `outils/` | `parse.py`, `analyse.py`, `analyse2.py` | Lecture et diagnostic du fichier d'origine |
| `outils/` | `classify.py` | Détection des rythmes fixe / quinzaine A / quinzaine B |
| `outils/` | `rows.py`, `build.py` | Construction du classeur final |

Les scripts sont reproductibles : `pip install openpyxl` puis
`python3 parse.py && python3 analyse.py && python3 classify.py && python3 rows.py && python3 build.py`.

## Structure du classeur (13 feuilles)

1. **Sommaire** — navigation cliquable, légende, mode d'emploi, points d'attention.
2. **Planning** — *table de référence*. 1 ligne = 1 séance (701 lignes), 26 colonnes,
   filtres, volets figés, menus déroulants, contrôles de conflit par ligne.
   **C'est la seule feuille où l'on saisit.**
3. **Référentiels** — classes, matières, enseignants, salles, modes, rythmes, calendrier des semaines.
4. à 10. **S1 … S7** — 7 grilles hebdomadaires imprimables (classes × jours).
   Leur contenu est **calculé par formules** depuis `Planning` : aucune double saisie.
   Une case = **une seule séance** : 1re ligne l'horaire, 2e la matière, 3e l'enseignant.
   Une classe ayant deux séances le même jour occupe deux sous-lignes.
5. **Vue Enseignants** — séances et heures par enseignant et par semaine.
6. **Contrôles** — 21 contrôles automatiques + les 14 décisions de construction d'octobre.
7. **Récapitulatif** — indicateurs par semaine, cycle, matière et jour.

## Séances partagées entre plusieurs groupes

Dans le fichier d'origine, une séance réunissant plusieurs groupes est écrite **une seule fois**,
dans une cellule fusionnée qui déborde sur les blocs de toutes les classes concernées.
Il y en a **16 par semaine A et 14 par semaine B** (Physique de Melek sur tout un niveau,
Espagnol, Italien, Philo, Français de Nassima en 9éme (B)+(C), etc.).

Elles sont conservées et rendues explicites :

- dans `Planning`, **un groupe par ligne**, les lignes d'une même séance étant reliées par un **ID séance** commun ;
- une seule ligne porte `Ligne principale = Oui`, pour ne compter l'heure d'enseignement qu'une fois ;
- dans les grilles, la séance n'est **écrite qu'une seule fois**, dans une **cellule fusionnée**
  qui couvre les lignes de tous les groupes concernés — comme dans le fichier d'origine.

Deux rythmes coexistent :

| Rythme | Signification |
|---|---|
| `Hebdomadaire` | même séance, mêmes groupes, toutes les semaines |
| `Hebdo · groupes alternés` | la séance a lieu chaque semaine, mais **les groupes alternent** (13 cas) |
| `Quinzaine A` / `Quinzaine B` | la séance elle-même n'a lieu qu'une semaine sur deux |

Vu d'un groupe, les deux derniers reviennent une semaine sur deux : d'où le repère `◆ Q-A` / `◆ Q-B`.

## Cours en option

Espagnol (Imen) et Italien (Amel) se chevauchent le mardi 19:15–20:45 pour un même groupe
(Bac SCE en semaine A, 3éme INFO en semaine B). Lecture retenue : **le groupe se scinde**,
chaque élève suivant l'une OU l'autre langue. Ces séances portent `Cours en option = Oui`
et sont exclues de la détection de conflit. Si ce n'est pas une option, videz cette colonne.

## Règle des semaines A / B

Les séances par quinzaine alternent sur deux semaines. La parité a été déduite
en comparant les 3 semaines de septembre :

| Parité | Semaines |
|---|---|
| **A** | 14/09 · **28/09 (référence)** · 12/10 · 26/10 |
| **B** | 21/09 · 05/10 · 19/10 |

Construction d'octobre : semaine A = 81 séances fixes + 20 quinzaine A = **101**,
semaine B = 81 séances fixes + 18 quinzaine B = **99**.

## Garanties vérifiées

- 301 / 301 séances d'origine présentes (370 lignes groupe-séance), **0 altérée, groupes compris** (le texte brut de chaque cellule
  d'origine est conservé en colonne Y de `Planning`, avec sa référence en colonne Z).
- 0 conflit d'enseignant, 0 conflit de groupe, 0 doublon, 0 partage incohérent, sur les 7 semaines.
- 0 erreur de formule dans l'ensemble du classeur (recalcul complet vérifié).
- Dates et jours conformes au calendrier réel d'octobre 2026.
- Impression : A4 paysage, 1 page de large, saut de page à chaque cycle, en-têtes répétés.
  57 pages pour le fichier complet, 24 pour celui d'octobre.

## Points nécessitant un arbitrage

Détaillés en bas de la feuille **Contrôles**. Les trois principaux :

1. **Physique (Melek)** — 7éme (A) mercredi, 8éme (A) vendredi, 9éme (A) lundi :
   portent la mention « /par quinzaine » mais sont présentes les 3 semaines.
   Traitées comme **hebdomadaires**.
2. **STI (Aymen) en 3éme INFO** — dimanche en S2, lundi dans la semaine de référence.
   Traitée comme **hebdomadaire le lundi**.
3. **Jeudi 15/10/2026** — Fête de l'Évacuation. Les 12 séances sont **maintenues** et signalées.

Par ailleurs, **Salle** et **Mode (présentiel / en ligne)** sont absents du fichier d'origine :
les colonnes existent mais restent vides. Aucune valeur n'a été inventée.

## Préparer novembre

1. Dupliquer une grille d'octobre, renommer l'onglet.
2. Mettre à jour la cellule **J4** (ligne masquée) avec le nouveau code de semaine.
3. Ajouter la semaine dans **Référentiels** (code, dates, parité).
4. Dans **Planning**, filtrer une semaine de même parité, copier ses lignes, les coller
   et changer *Semaine* / *Période* / *Date*.

La grille se remplit alors toute seule.

## Code couleur des grilles

Le fond de chaque case indique la **famille de matières** :

| Famille | Matières |
|---|---|
| Arabe | Arabe, عربية |
| Français | Français |
| Langues étrangères | Anglais, Espagnol, Italien |
| Mathématiques | Math, رياضيات |
| Sciences | Physique, فيزياء, SVT, Éveil scientifique, إيقاظ علمي |
| Sciences humaines | Histoire, Géographie, Philo |
| Économie / Gestion | Eco, Gestion |
| Informatique | Informatique, ALGO, STI |

La colonne *Classe* garde la couleur du **cycle**. Les repères `◆ Q-A` / `◆ Q-B`
marquent les séances par quinzaine, `◆ PONCTUEL` les séances non récurrentes.

Le texte des cases est en **gras noir**, 12 pt. Le code couleur est posé par
**mise en forme conditionnelle** : si vous changez la matière d'une séance dans
`Planning`, la couleur suit toute seule.

## Ce qui est vivant, ce qui est figé

Chaque case de grille va chercher sa séance dans `Planning` **par son ID séance**.

- **Vivant** : horaire, matière, enseignant, salle, mode, notes. Modifiez-les dans
  `Planning`, la grille et tous les indicateurs suivent.
- **Figé** : la *structure* des grilles — nombre de lignes par classe et cellules
  fusionnées. Une séance **ajoutée** dans `Planning` n'apparaîtra pas d'elle-même
  dans la grille ; il faut régénérer le classeur (`python3 outils/build.py`).

## Le fichier « Octobre seul »

`RS_2026-2027_Octobre_seul.xlsx` ne contient que les 4 grilles d'octobre, avec un
**contenu figé** (aucune formule, aucun lien externe). Il est autonome : rien à
recalculer, rien à activer. C'est celui à diffuser à l'équipe.
Son contenu a été comparé cellule à cellule avec le fichier complet : identique.
