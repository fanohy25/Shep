# Shep

Deck de présentation « **Shep — Dossier stratégique 2026** », généré par code
et entièrement modifiable dans PowerPoint / Keynote / Google Slides.

## Livrable

`Shep_Presentation_Strategique_2026.pptx` — 18 diapositives, format 16:9
(33,87 × 19,05 cm), notes de présentateur incluses sur chaque slide clé.

| # | Slide | Type |
|---|-------|------|
| 01 | Couverture | Visuel plein écran |
| 02 | Sommaire | Grille 2 × 5 |
| 03 | Séparateur — Contexte & problème | Visuel |
| 04 | Un marché qui bascule vers l'orchestration | 4 KPI + encart lecture |
| 05 | Le coût caché de la fragmentation | 4 cartes + bandeau |
| 06 | Séparateur — La plateforme Shep | Visuel |
| 07 | Une architecture en trois couches | 3 colonnes + bandeau résultats |
| 08 | Ce que les alternatives ne couvrent pas | Matrice comparative 5 × 6 |
| 09 | Séparateur — Traction & performance | Visuel |
| 10 | Une adoption qui s'accélère | Histogramme + 3 KPI |
| 11 | Un modèle à fort effet de levier | Grille tarifaire + unit economics |
| 12 | Trajectoire 2026 – 2029 | Courbes + anneau + hypothèses |
| 13 | Séquence de conquête par segment | 3 moteurs + priorités |
| 14 | Quatre phases, des jalons vérifiables | Timeline + jalons bloquants |
| 15 | L'équipe qui exécute | 4 profils + effectifs + gouvernance |
| 16 | Risques et plans de mitigation | Matrice 2 × 2 |
| 17 | Ce que nous demandons | Allocation + étapes + termes |
| 18 | Clôture | Visuel + contacts |

Les graphiques (slides 10 et 12) sont des **graphiques Office natifs** : un
double-clic ouvre la feuille de données, les chiffres se modifient directement
dans PowerPoint.

## Système de design

| Élément | Valeur |
|---|---|
| Fond clair | `#F7F9FC` · texte `#142132` |
| Fond sombre | `#081424` · carte `#102136` |
| Accent primaire | `#2E7DFF` |
| Accent secondaire | `#00C2A8` |
| Alerte / nuance | `#F5A623` · `#F0625B` |
| Typographie | Segoe UI (Light / Regular / Semibold) |
| Grille | marge 2,16 cm, filet d'accent latéral, pied de page numéroté |

Les slides alternent fonds clairs et sombres pour rythmer la lecture ; les trois
séparateurs de section utilisent des visuels générés (`assets/`).

## Régénérer ou modifier le deck

```bash
pip install python-pptx
cd deck
python3 build_deck.py                    # -> ../Shep_Presentation_Strategique_2026.pptx
python3 build_deck.py /chemin/sortie.pptx
```

| Fichier | Rôle |
|---|---|
| `deck/theme.py` | palette, typographie, grille, helpers de rendu |
| `deck/components.py` | masters de slides, cartes KPI, timeline, tableaux, graphiques |
| `deck/build_deck.py` | contenu et mise en page des 18 slides |
| `deck/preview.py` | rendu PNG de contrôle (outil interne, hors livrable) |
| `assets/` | visuels de fond des slides pleine page |

**Adapter le contenu** : tout le texte vit dans `build_deck.py`, slide par slide,
dans l'ordre de la présentation. **Adapter la charte** : modifier les constantes
de couleur et `FONT` dans `theme.py` — la mise à jour se propage à l'ensemble du
deck.

## Contrôle visuel

```bash
cd deck && python3 preview.py            # PNG approximatifs dans /tmp/preview
```
Rendu indicatif (Pillow, sans moteur Office) : utile pour repérer les
débordements, pas pour juger du rendu typographique final.
