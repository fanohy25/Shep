#!/usr/bin/env python3
"""
Génération du deck « Shep — Dossier stratégique 2026 ».
Usage : python3 build_deck.py [chemin_de_sortie.pptx]
"""

import os
import sys

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

from theme import *
from components import *

HERE = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.join(os.path.dirname(HERE), "assets")
OUT = sys.argv[1] if len(sys.argv) > 1 else os.path.join(
    os.path.dirname(HERE), "Shep_Presentation_Strategique_2026.pptx")

COVER = os.path.join(ASSETS, "cover_bg.png")
DIVIDER = os.path.join(ASSETS, "divider_bg.png")
CLOSING = os.path.join(ASSETS, "closing_bg.png")

prs = Presentation()
prs.slide_width, prs.slide_height = SW, SH
prs.core_properties.title = "Shep — Dossier stratégique 2026"
prs.core_properties.author = "Shep"
prs.core_properties.subject = "Plateforme d'orchestration des opérations"
prs.core_properties.comments = "Document confidentiel"


def notes(slide, txt):
    slide.notes_slide.notes_text_frame.text = txt


# ==========================================================================
# 01 — COUVERTURE
# ==========================================================================
s = blank(prs)
rect(s, 0, 0, SW, SH, INK)
picture_bg(s, COVER, prs)
scrim(s, prs, INK, 38000)
# bande verticale d'accent
rect(s, 0, 0, Inches(0.13), SH, ACCENT)

text(s, MARGIN, Inches(1.5), Inches(3), Inches(0.3),
     "DOSSIER STRATÉGIQUE  ·  2026", size=11, color=TEAL, bold=True,
     font=FONT_SEMI, spacing=3.2)

text(s, MARGIN, Inches(2.05), Inches(8.6), Inches(1.0),
     "SHEP", size=72, color=WHITE, bold=True, font=FONT_SEMI, spacing=6, line=1.0)

rect(s, MARGIN, Inches(3.18), Inches(1.6), Pt(4), ACCENT)

text(s, MARGIN, Inches(3.55), Inches(7.6), Inches(1.4),
     [("Orchestrer les opérations critiques\n", {"size": 30, "color": WHITE,
                                                 "font": FONT_LIGHT}),
      ("d'un bout à l'autre de la chaîne de valeur.", {"size": 30,
                                                       "color": WHITE,
                                                       "bold": True,
                                                       "font": FONT_SEMI})],
     line=1.18)

text(s, MARGIN, Inches(5.25), Inches(7.2), Inches(0.7),
     "Vision produit, trajectoire de croissance et plan d'exécution "
     "présentés au comité stratégique et aux partenaires financiers.",
     size=13, color=SLATE_DK, line=1.45)

rect(s, MARGIN, Inches(6.28), Inches(6.4), Pt(0.75), RGBColor(0x24, 0x3C, 0x58))
text(s, MARGIN, Inches(6.52), Inches(7), Inches(0.3),
     "Comité stratégique  ·  Antananarivo  ·  Septembre 2026",
     size=10.5, color=SLATE_DK, spacing=1.4)
text(s, SW - MARGIN - Inches(3), Inches(6.52), Inches(3), Inches(0.3),
     "CONFIDENTIEL", size=10.5, color=TEAL, bold=True, font=FONT_SEMI,
     align=PP_ALIGN.RIGHT, spacing=2.4)
notes(s, "Ouverture : 30 secondes. Poser le cadre — où nous en sommes, "
         "ce que nous demandons, et pourquoi maintenant.")

# ==========================================================================
# 02 — SOMMAIRE
# ==========================================================================
s = light_slide(prs, "Sommaire", kicker="Déroulé de la séance",
                subtitle="Quarante minutes de présentation, vingt minutes "
                         "d'échange.")
items = [
    ("01", "Contexte de marché", "Une fenêtre d'opportunité de 18 mois"),
    ("02", "Problème adressé", "Le coût caché de la fragmentation"),
    ("03", "Notre solution", "La plateforme Shep en trois couches"),
    ("04", "Différenciation", "Ce que les alternatives ne font pas"),
    ("05", "Traction", "Preuves d'adoption et signaux de marché"),
    ("06", "Modèle économique", "Unit economics et effet de levier"),
    ("07", "Trajectoire financière", "Du plan 2026 à l'horizon 2029"),
    ("08", "Go-to-market", "Séquence de conquête par segment"),
    ("09", "Roadmap", "Quatre phases, jalons vérifiables"),
    ("10", "Organisation & risques", "Équipe, gouvernance, mitigation"),
]
col_x = [MARGIN, MARGIN + Inches(6.05)]
for i, (num, title, sub) in enumerate(items):
    cx = col_x[i // 5]
    cy = Inches(2.35) + (i % 5) * Inches(0.88)
    text(s, cx, cy + Inches(0.02), Inches(0.6), Inches(0.36), num,
         size=19, color=ACCENT, bold=True, font=FONT_SEMI)
    text(s, cx + Inches(0.72), cy, Inches(4.7), Inches(0.3), title,
         size=14.5, color=GRAPHITE, bold=True, font=FONT_SEMI)
    text(s, cx + Inches(0.72), cy + Inches(0.29), Inches(4.7), Inches(0.3), sub,
         size=10.5, color=SLATE)
    rect(s, cx, cy + Inches(0.68), Inches(5.35), Pt(0.75), LINE_LIGHT)
notes(s, "Annoncer la structure. Préciser que les annexes financières "
         "détaillées sont fournies en document séparé.")

# ==========================================================================
# 03 — SÉPARATEUR : LE CONTEXTE
# ==========================================================================
def divider(prs, num, title, subtitle, bg=DIVIDER):
    s = blank(prs)
    rect(s, 0, 0, SW, SH, INK)
    picture_bg(s, bg, prs)
    scrim(s, prs, INK, 52000)
    rect(s, 0, 0, Inches(0.13), SH, TEAL)
    text(s, MARGIN, Inches(2.55), Inches(2), Inches(0.9), num,
         size=88, color=RGBColor(0x1C, 0x3B, 0x5E), bold=True, font=FONT_SEMI,
         line=1.0)
    text(s, MARGIN, Inches(3.55), Inches(9), Inches(0.7), title,
         size=40, color=WHITE, bold=True, font=FONT_SEMI, line=1.05)
    rect(s, MARGIN, Inches(4.35), Inches(1.3), Pt(4), TEAL)
    text(s, MARGIN, Inches(4.7), Inches(7.6), Inches(0.6), subtitle,
         size=14, color=SLATE_DK, line=1.4)
    return s


divider(prs, "01", "Contexte & problème",
        "Pourquoi le statu quo n'est plus tenable pour les directions "
        "opérationnelles.")

# ==========================================================================
# 04 — CONTEXTE DE MARCHÉ (KPI)
# ==========================================================================
s = light_slide(prs, "Un marché qui bascule vers l'orchestration",
                kicker="Contexte de marché",
                subtitle="Trois dynamiques convergent et créent une fenêtre "
                         "d'opportunité estimée à 18 mois.")
kpis = [
    ("42 Md$", "Marché adressable 2029", "TAM des plateformes d'orchestration "
     "opérationnelle, TCAC de 21 %.", ACCENT),
    ("68 %", "Directions en refonte", "Des DSI/DOP déclarent un projet de "
     "rationalisation d'outillage en cours.", TEAL),
    ("9,4", "Outils par équipe", "Nombre médian d'applications utilisées sur "
     "un même processus terrain.", AMBER),
    ("14 M$", "Coût de l'inaction", "Perte annuelle moyenne liée aux ruptures "
     "de coordination sur un ETI de 2 000 personnes.", CORAL),
]
cw = (CONTENT_W - Inches(0.36) * 3) / 4
for i, (v, l, n, c) in enumerate(kpis):
    kpi_card(s, MARGIN + i * (cw + Inches(0.36)), Inches(2.3), cw,
             Inches(2.3), v, l, n, accent=c)

card(s, MARGIN, Inches(4.92), CONTENT_W, Inches(1.5), WHITE, corner=0.06,
     border=LINE_LIGHT)
rect(s, MARGIN, Inches(4.92), Pt(4), Inches(1.5), ACCENT)
text(s, MARGIN + Inches(0.42), Inches(5.18), Inches(1.9), Inches(0.3),
     "LECTURE", size=10, color=ACCENT, bold=True, font=FONT_SEMI, spacing=2)
text(s, MARGIN + Inches(0.42), Inches(5.5), CONTENT_W - Inches(0.9), Inches(0.9),
     "La valeur ne se déplace plus vers de nouveaux outils métiers, mais vers "
     "la couche qui les relie. Les acheteurs consolident : ils remplacent "
     "quatre à six briques par une plateforme unique capable de porter la "
     "donnée, la règle métier et l'exécution dans le même flux.",
     size=12.5, color=SLATE, line=1.4)
notes(s, "Insister sur la consolidation : le budget existe déjà, il est "
         "simplement dispersé. Nous ne créons pas une ligne budgétaire, "
         "nous en agrégeons plusieurs.")

# ==========================================================================
# 05 — LE PROBLÈME
# ==========================================================================
s = dark_slide(prs, "Le coût caché de la fragmentation",
               kicker="Problème adressé",
               subtitle="Quatre ruptures structurelles qui dégradent la marge "
                        "opérationnelle, trimestre après trimestre.")
probs = [
    ("01", "Données dispersées",
     "Les référentiels vivent dans des silos non réconciliés. Chaque décision "
     "démarre par une reconstruction manuelle de la vérité terrain."),
    ("02", "Processus non traçables",
     "Aucune piste d'audit continue entre la demande, l'arbitrage et "
     "l'exécution. Les écarts se découvrent en fin de cycle."),
    ("03", "Coordination manuelle",
     "31 % du temps des managers de terrain part en relances, ressaisies et "
     "réconciliations entre systèmes."),
    ("04", "Pilotage en retard",
     "Le reporting arrive avec 12 à 20 jours de décalage : on constate les "
     "dérives au lieu de les prévenir."),
]
cw = (CONTENT_W - Inches(0.32) * 3) / 4
for i, (n, t, b) in enumerate(probs):
    feature_card(s, MARGIN + i * (cw + Inches(0.32)), Inches(2.32), cw,
                 Inches(2.72), n, t, b,
                 accent=[ACCENT, TEAL, AMBER, CORAL][i], dark=True)

card(s, MARGIN, Inches(5.36), CONTENT_W, Inches(1.1), RGBColor(0x14, 0x2B, 0x47),
     corner=0.08, shadow=False)
rect(s, MARGIN, Inches(5.36), Pt(4), Inches(1.1), TEAL)
text(s, MARGIN + Inches(0.42), Inches(5.62), CONTENT_W - Inches(0.9), Inches(0.6),
     [("Conséquence mesurée : ", {"bold": True, "font": FONT_SEMI,
                                  "color": WHITE, "size": 13}),
      ("entre 6 et 11 points de marge opérationnelle perdus chaque année sur "
       "les processus multi-équipes — sans qu'aucun responsable unique ne "
       "puisse en rendre compte.", {"color": SLATE_DK, "size": 13})],
     line=1.35)
notes(s, "Ancrer sur le chiffre de fin : c'est celui que le CFO retiendra.")

# ==========================================================================
# 06 — SÉPARATEUR : LA SOLUTION
# ==========================================================================
divider(prs, "02", "La plateforme Shep",
        "Une couche d'orchestration unique : connecter, décider, exécuter, "
        "prouver.")

# ==========================================================================
# 07 — LA SOLUTION (3 couches)
# ==========================================================================
s = light_slide(prs, "Une architecture en trois couches",
                kicker="Notre solution",
                subtitle="Shep s'installe au-dessus de l'existant. Aucune "
                         "migration lourde, une valeur mesurable en 30 jours.")
layers = [
    ("COUCHE 1", "Connexion", ACCENT,
     ["Connecteurs natifs ERP, CRM, MES, RH", "Réconciliation des "
      "référentiels en temps réel", "Historisation immuable et horodatée"]),
    ("COUCHE 2", "Orchestration", TEAL,
     ["Moteur de règles métier sans code", "Arbitrages automatisés et "
      "escalades", "Workflows conditionnels multi-équipes"]),
    ("COUCHE 3", "Pilotage", AMBER,
     ["Tableaux de bord temps réel par rôle", "Détection d'anomalies "
      "prédictive", "Piste d'audit exportable et conforme"]),
]
cw = (CONTENT_W - Inches(0.4) * 2) / 3
for i, (tag, title, col, feats) in enumerate(layers):
    x = MARGIN + i * (cw + Inches(0.4))
    card(s, x, Inches(2.3), cw, Inches(3.4), WHITE, corner=0.06,
         border=LINE_LIGHT)
    rect(s, x, Inches(2.3), cw, Inches(0.09), col)
    text(s, x + Inches(0.34), Inches(2.62), cw - Inches(0.6), Inches(0.26),
         tag, size=9.5, color=col, bold=True, font=FONT_SEMI, spacing=2)
    text(s, x + Inches(0.34), Inches(2.95), cw - Inches(0.6), Inches(0.42),
         title, size=22, color=GRAPHITE, bold=True, font=FONT_SEMI)
    rect(s, x + Inches(0.34), Inches(3.48), Inches(0.7), Pt(2.5), col)
    bullets(s, x + Inches(0.34), Inches(3.78), cw - Inches(0.68), Inches(1.7),
            feats, size=11.5, color=SLATE, bullet_color=col, gap=9)

# bandeau résultat
card(s, MARGIN, Inches(5.92), CONTENT_W, Inches(0.72), NAVY, corner=0.1,
     shadow=False)
res = [("30 jours", "première valeur mesurée"),
       ("−31 %", "temps de coordination"),
       ("×3,2", "vitesse de traitement des exceptions"),
       ("100 %", "traçabilité des décisions")]
seg = CONTENT_W / 4
for i, (v, l) in enumerate(res):
    x = MARGIN + i * seg
    text(s, x, Inches(6.08), seg, Inches(0.28),
         [(v + "   ", {"size": 15, "bold": True, "font": FONT_SEMI,
                       "color": TEAL}),
          (l, {"size": 11, "color": RGBColor(0xC7, 0xD6, 0xE8)})],
         align=PP_ALIGN.CENTER)
notes(s, "Message clé : Shep ne remplace pas les systèmes en place, "
         "il les rend cohérents. C'est l'argument qui désamorce l'objection IT.")

# ==========================================================================
# 08 — DIFFÉRENCIATION (tableau)
# ==========================================================================
s = light_slide(prs, "Ce que les alternatives ne couvrent pas",
                kicker="Différenciation",
                subtitle="Positionnement face aux trois options réellement "
                         "évaluées par nos prospects.")
headers = ["Critère de décision", "Shep", "Suite ERP étendue",
           "Outils best-of-breed", "Développement interne"]
rows = [
    ["Délai de mise en production", "4 à 6 semaines", "9 à 14 mois",
     "2 à 3 mois", "12 à 24 mois"],
    ["Couverture inter-systèmes", "Native et complète", "Partielle",
     "Point à point", "Sur mesure, fragile"],
    ["Coût total sur 3 ans", "€", "€€€€", "€€€", "€€€€"],
    ["Autonomie des équipes métier", "Sans code", "Dépendance éditeur",
     "Variable", "Dépendance DSI"],
    ["Traçabilité & conformité", "Bout en bout", "Par module", "Absente",
     "À construire"],
    ["Évolutivité du modèle", "Continue", "Par version", "Limitée", "Coûteuse"],
]
comparison_table(s, MARGIN, Inches(2.3), CONTENT_W, headers, rows,
                 col_w=[Inches(3.5), Inches(2.1), Inches(2.1), Inches(2.1),
                        Inches(1.83)],
                 highlight_col=1)
text(s, MARGIN, Inches(6.05), CONTENT_W, Inches(0.5),
     [("Angle mort du marché : ", {"bold": True, "font": FONT_SEMI,
                                   "color": GRAPHITE, "size": 12}),
      ("aucun acteur ne combine aujourd'hui le temps de déploiement d'un "
       "outil léger avec la profondeur fonctionnelle d'une suite. "
       "C'est exactement l'espace que Shep occupe.",
       {"color": SLATE, "size": 12})], line=1.4)
notes(s, "Ne pas dénigrer les concurrents : présenter la matrice comme une "
         "aide à la décision. Le prospect se positionne seul.")

# ==========================================================================
# 09 — SÉPARATEUR : TRACTION
# ==========================================================================
divider(prs, "03", "Traction & performance",
        "Les preuves d'adoption, la mécanique de revenus et la trajectoire "
        "financière.")

# ==========================================================================
# 10 — TRACTION (graphique + KPI)
# ==========================================================================
s = light_slide(prs, "Une adoption qui s'accélère",
                kicker="Traction",
                subtitle="Revenu récurrent annuel et base clients, "
                         "T1 2025 – T3 2026.")
card(s, MARGIN, Inches(2.3), Inches(8.0), Inches(4.0), WHITE, corner=0.05,
     border=LINE_LIGHT)
text(s, MARGIN + Inches(0.35), Inches(2.52), Inches(6), Inches(0.3),
     "ARR (k€) et clients actifs", size=11.5, color=GRAPHITE, bold=True,
     font=FONT_SEMI)
bar_chart(s, MARGIN + Inches(0.15), Inches(2.85), Inches(7.7), Inches(3.3),
          ["T1-25", "T2-25", "T3-25", "T4-25", "T1-26", "T2-26", "T3-26"],
          [("ARR (k€)", [180, 310, 470, 720, 1050, 1480, 2060])],
          labels=True)

rx = MARGIN + Inches(8.4)
rw = CONTENT_W - Inches(8.4)
stats = [
    ("2,06 M€", "ARR au T3 2026", "+186 % en glissement annuel", ACCENT),
    ("134 %", "Net revenue retention", "Expansion nette sur base installée", TEAL),
    ("11 mois", "Payback CAC", "Sous la barre des 12 mois visée", AMBER),
]
for i, (v, l, n, c) in enumerate(stats):
    kpi_card(s, rx, Inches(2.3) + i * Inches(1.4), rw, Inches(1.28),
             v, l, n, accent=c, compact=True)
text(s, rx, Inches(6.5), rw, Inches(0.24),
     "Source : reporting interne, clôture 30/09/2026.",
     size=8, color=SLATE, align=PP_ALIGN.RIGHT)
notes(s, "Le NRR à 134 % est le chiffre le plus important de la slide : "
         "il prouve que la valeur se confirme après la vente.")

# ==========================================================================
# 11 — MODÈLE ÉCONOMIQUE
# ==========================================================================
s = dark_slide(prs, "Un modèle à fort effet de levier",
               kicker="Modèle économique",
               subtitle="Abonnement par plateforme, extension par volume "
                        "d'opérations orchestrées.")
tiers = [
    ("ESSENTIEL", "1 900 €", "/ mois", ACCENT,
     ["Jusqu'à 150 utilisateurs", "8 connecteurs standards",
      "Support 8h/5j", "Onboarding guidé"]),
    ("BUSINESS", "4 800 €", "/ mois", TEAL,
     ["Utilisateurs illimités", "Connecteurs illimités",
      "Support 24/5 et SLA 99,9 %", "Customer success dédié"]),
    ("ENTREPRISE", "Sur devis", "", AMBER,
     ["Déploiement multi-entités", "Hébergement souverain au choix",
      "Conformité et audit renforcés", "Accompagnement stratégique"]),
]
cw = (Inches(7.9) - Inches(0.3) * 2) / 3
for i, (name, price, per, col, feats) in enumerate(tiers):
    x = MARGIN + i * (cw + Inches(0.3))
    highlight = (i == 1)
    card(s, x, Inches(2.3), cw, Inches(3.15),
         RGBColor(0x14, 0x2C, 0x4A) if highlight else INK_SOFT,
         corner=0.07, shadow=False)
    rect(s, x, Inches(2.3), cw, Inches(0.07), col)
    if highlight:
        pill(s, x + cw - Inches(1.32), Inches(2.5), Inches(1.1), Inches(0.26),
             "RECOMMANDÉ", TEAL, INK, size=7.5)
    text(s, x + Inches(0.3), Inches(2.56), cw - Inches(0.6), Inches(0.26),
         name, size=9.5, color=col, bold=True, font=FONT_SEMI, spacing=2)
    text(s, x + Inches(0.3), Inches(2.9), cw - Inches(0.6), Inches(0.45),
         [(price, {"size": 25, "bold": True, "font": FONT_SEMI, "color": WHITE}),
          ("  " + per, {"size": 11, "color": SLATE_DK})], line=1.0)
    rect(s, x + Inches(0.3), Inches(3.52), cw - Inches(0.6), Pt(0.75),
         RGBColor(0x24, 0x3C, 0x58))
    bullets(s, x + Inches(0.3), Inches(3.75), cw - Inches(0.6), Inches(1.5),
            feats, size=10.5, color=SLATE_DK, bullet_color=col, gap=7)

rx = MARGIN + Inches(8.2)
rw = CONTENT_W - Inches(8.2)
text(s, rx, Inches(2.3), rw, Inches(0.3), "UNIT ECONOMICS", size=10,
     color=TEAL, bold=True, font=FONT_SEMI, spacing=2)
ue = [("LTV / CAC", "4,7×"), ("Marge brute", "82 %"),
      ("ACV moyen", "46 k€"), ("Churn logo annuel", "6,1 %"),
      ("Cycle de vente", "74 jours")]
for i, (k, v) in enumerate(ue):
    y = Inches(2.72) + i * Inches(0.58)
    text(s, rx, y, rw - Inches(1.1), Inches(0.28), k, size=11.5,
         color=SLATE_DK)
    text(s, rx, y, rw, Inches(0.28), v, size=13, color=WHITE, bold=True,
         font=FONT_SEMI, align=PP_ALIGN.RIGHT)
    rect(s, rx, y + Inches(0.38), rw, Pt(0.75), RGBColor(0x1E, 0x33, 0x4A))
notes(s, "Le tier Business concentre 61 % des signatures : c'est notre "
         "point d'équilibre prix/valeur.")

# ==========================================================================
# 12 — TRAJECTOIRE FINANCIÈRE
# ==========================================================================
s = light_slide(prs, "Trajectoire 2026 – 2029",
                kicker="Plan financier",
                subtitle="Scénario central : rentabilité opérationnelle "
                         "atteinte au second semestre 2028.")
card(s, MARGIN, Inches(2.3), Inches(7.3), Inches(3.05), WHITE, corner=0.05,
     border=LINE_LIGHT)
text(s, MARGIN + Inches(0.32), Inches(2.5), Inches(5), Inches(0.28),
     "Revenus et résultat opérationnel (M€)", size=11, color=GRAPHITE,
     bold=True, font=FONT_SEMI)
line_chart(s, MARGIN + Inches(0.1), Inches(2.8), Inches(7.05), Inches(2.4),
           ["2026", "2027", "2028", "2029"],
           [("Revenus", [2.6, 6.8, 15.4, 31.0]),
            ("Résultat opérationnel", [-1.9, -2.4, 0.9, 7.6])],
           colors=[ACCENT, TEAL])

card(s, MARGIN + Inches(7.62), Inches(2.3), CONTENT_W - Inches(7.62),
     Inches(3.05), WHITE, corner=0.05, border=LINE_LIGHT)
text(s, MARGIN + Inches(7.94), Inches(2.5), Inches(4), Inches(0.28),
     "Structure de coûts 2027", size=11, color=GRAPHITE, bold=True,
     font=FONT_SEMI)
donut_chart(s, MARGIN + Inches(7.74), Inches(2.76), Inches(3.78), Inches(2.42),
            ["R&D", "Go-to-market", "Delivery & support", "G&A"],
            [38, 34, 18, 10],
            colors=[ACCENT, TEAL, AMBER, SLATE])

# hypothèses
card(s, MARGIN, Inches(5.55), CONTENT_W, Inches(1.0), WHITE, corner=0.05,
     border=LINE_LIGHT)
hyp = [("×2,6", "croissance annuelle moyenne des revenus"),
       ("82 %", "marge brute maintenue à l'échelle"),
       ("S2 2028", "point mort opérationnel"),
       ("22 mois", "runway au closing du tour")]
seg = CONTENT_W / 4
for i, (v, l) in enumerate(hyp):
    x = MARGIN + i * seg
    if i:
        rect(s, x, Inches(5.72), Pt(0.75), Inches(0.66), LINE_LIGHT)
    text(s, x + Inches(0.3), Inches(5.76), seg - Inches(0.5), Inches(0.32), v,
         size=18, color=ACCENT, bold=True, font=FONT_SEMI)
    text(s, x + Inches(0.3), Inches(6.12), seg - Inches(0.5), Inches(0.3), l,
         size=10, color=SLATE, line=1.2)
notes(s, "Préciser que le scénario prudent (−30 % sur le pipeline) reporte "
         "le point mort au T2 2029 sans besoin de financement additionnel.")

# ==========================================================================
# 13 — GO-TO-MARKET
# ==========================================================================
s = dark_slide(prs, "Séquence de conquête par segment",
               kicker="Go-to-market",
               subtitle="Trois moteurs d'acquisition activés dans un ordre "
                        "précis pour préserver l'efficacité du capital.")
motors = [
    ("MOTEUR 1", "Vente directe grands comptes", ACCENT,
     "Équipe de 6 AE seniors ciblant l'industrie et la logistique. "
     "ACV cible 80 k€, cycle 90 jours.", "54 % du pipeline"),
    ("MOTEUR 2", "Partenariats intégrateurs", TEAL,
     "Cinq cabinets référencés, marge revendeur 25 %. Effet de levier sur "
     "les appels d'offres publics.", "29 % du pipeline"),
    ("MOTEUR 3", "Product-led expansion", AMBER,
     "Essai encadré 21 jours puis expansion par équipe. Alimente le segment "
     "mid-market sans coût commercial.", "17 % du pipeline"),
]
for i, (tag, title, col, body, share) in enumerate(motors):
    y = Inches(2.3) + i * Inches(1.35)
    card(s, MARGIN, y, Inches(8.2), Inches(1.15), INK_SOFT, corner=0.08,
         shadow=False)
    rect(s, MARGIN, y, Pt(4), Inches(1.15), col)
    text(s, MARGIN + Inches(0.4), y + Inches(0.2), Inches(1.8), Inches(0.24),
         tag, size=9, color=col, bold=True, font=FONT_SEMI, spacing=1.8)
    text(s, MARGIN + Inches(0.4), y + Inches(0.46), Inches(3.4), Inches(0.3),
         title, size=14, color=WHITE, bold=True, font=FONT_SEMI)
    text(s, MARGIN + Inches(4.0), y + Inches(0.26), Inches(2.9), Inches(0.72),
         body, size=10.5, color=SLATE_DK, line=1.3)
    pill(s, MARGIN + Inches(6.95), y + Inches(0.4), Inches(1.12), Inches(0.3),
         share.split(" ")[0] + " %", col, INK, size=10)
    text(s, MARGIN + Inches(6.95), y + Inches(0.76), Inches(1.12), Inches(0.22),
         "du pipeline", size=8, color=SLATE_DK, align=PP_ALIGN.CENTER)

rx = MARGIN + Inches(8.6)
rw = CONTENT_W - Inches(8.6)
card(s, rx, Inches(2.3), rw, Inches(4.05), INK_SOFT, corner=0.06, shadow=False)
text(s, rx + Inches(0.32), Inches(2.55), rw - Inches(0.6), Inches(0.3),
     "PRIORITÉS 12 MOIS", size=9.5, color=TEAL, bold=True, font=FONT_SEMI,
     spacing=2)
prio = [
    ("Verticaliser", "industrie et logistique d'abord, santé en 2027."),
    ("Industrialiser", "playbook de vente et démonstrateur sectoriel."),
    ("Internationaliser", "Benelux et péninsule ibérique au S2 2027."),
    ("Prouver", "trois études de cas chiffrées et auditées."),
]
for i, (k, v) in enumerate(prio):
    y = Inches(3.0) + i * Inches(0.82)
    rect(s, rx + Inches(0.32), y + Inches(0.07), Inches(0.1), Inches(0.1),
         TEAL, MSO_SHAPE.OVAL)
    text(s, rx + Inches(0.6), y, rw - Inches(0.95), Inches(0.62),
         [(k + " — ", {"bold": True, "font": FONT_SEMI, "color": WHITE,
                       "size": 11.5}),
          (v, {"color": SLATE_DK, "size": 11.5})], line=1.3)
notes(s, "Ordre volontaire : la vente directe finance les deux autres "
         "moteurs. Ne pas les lancer en parallèle.")

# ==========================================================================
# 14 — ROADMAP
# ==========================================================================
s = light_slide(prs, "Quatre phases, des jalons vérifiables",
                kicker="Roadmap d'exécution",
                subtitle="Chaque phase est conditionnée par l'atteinte des "
                         "jalons de la précédente.")
timeline(s, MARGIN, Inches(2.4), CONTENT_W, [
    ("T4 2026", "Consolidation", [
        "Certification ISO 27001", "Connecteurs SAP et Dynamics",
        "3 études de cas publiées", "ARR 2,9 M€"]),
    ("S1 2027", "Industrialisation", [
        "Moteur de règles v2 sans code", "Portail partenaires intégrateurs",
        "SLA 99,95 %", "ARR 5,1 M€"]),
    ("S2 2027", "Expansion", [
        "Ouverture Benelux et Ibérie", "Modules sectoriels santé",
        "Marketplace de connecteurs", "ARR 6,8 M€"]),
    ("2028 – 2029", "Passage à l'échelle", [
        "Copilote décisionnel prédictif", "Souveraineté multi-région",
        "Point mort opérationnel", "ARR 15 M€ puis 31 M€"]),
])
card(s, MARGIN, Inches(5.55), CONTENT_W, Inches(1.0), NAVY, corner=0.1,
     shadow=False)
gates = [("Jalon bloquant", "Certification ISO 27001"),
         ("Décision comité", "Fin T4 2026"),
         ("Indicateur de passage", "ARR et NRR engagés"),
         ("Responsable", "Comité produit + DG")]
seg = CONTENT_W / 4
for i, (k, v) in enumerate(gates):
    x = MARGIN + i * seg
    if i:
        rect(s, x, Inches(5.72), Pt(0.75), Inches(0.66),
             RGBColor(0x2C, 0x4A, 0x6B))
    text(s, x + Inches(0.34), Inches(5.76), seg - Inches(0.55), Inches(0.26),
         k.upper(), size=8.5, color=TEAL, bold=True, font=FONT_SEMI, spacing=1.6)
    text(s, x + Inches(0.34), Inches(6.04), seg - Inches(0.55), Inches(0.3), v,
         size=12, color=WHITE, bold=True, font=FONT_SEMI)
notes(s, "Insister sur le caractère conditionnel : pas de phase suivante "
         "sans jalon atteint. C'est notre discipline d'allocation.")

# ==========================================================================
# 15 — ÉQUIPE & GOUVERNANCE
# ==========================================================================
s = light_slide(prs, "L'équipe qui exécute",
                kicker="Organisation",
                subtitle="42 collaborateurs, dont 58 % en R&D et delivery.")
team = [
    ("Direction générale", "Stratégie, financement, partenariats clés",
     "18 ans en édition logicielle B2B", ACCENT),
    ("Direction technique", "Plateforme, sécurité, fiabilité",
     "Ex-architecte de systèmes à forte charge", TEAL),
    ("Direction produit", "Vision, discovery, design system",
     "3 produits menés du zéro à l'échelle", AMBER),
    ("Direction commerciale", "Vente directe, partenariats, expansion",
     "Construction de deux équipes à 10 M€ d'ARR", CORAL),
]
cw = (CONTENT_W - Inches(0.34) * 3) / 4
for i, (role, scope, bio, col) in enumerate(team):
    x = MARGIN + i * (cw + Inches(0.34))
    card(s, x, Inches(2.3), cw, Inches(2.18), WHITE, corner=0.06,
         border=LINE_LIGHT)
    rect(s, x + Inches(0.3), Inches(2.58), Inches(0.5), Inches(0.5), col,
         MSO_SHAPE.ROUNDED_RECTANGLE, 0.25)
    text(s, x + Inches(0.3), Inches(2.72), Inches(0.5), Inches(0.28),
         ["DG", "CTO", "CPO", "CRO"][i], size=11, color=WHITE, bold=True,
         font=FONT_SEMI, align=PP_ALIGN.CENTER)
    text(s, x + Inches(0.3), Inches(3.24), cw - Inches(0.6), Inches(0.5),
         role, size=12.5, color=GRAPHITE, bold=True, font=FONT_SEMI, line=1.1)
    text(s, x + Inches(0.3), Inches(3.62), cw - Inches(0.6), Inches(0.42),
         scope, size=10, color=SLATE, line=1.25)
    text(s, x + Inches(0.3), Inches(4.04), cw - Inches(0.6), Inches(0.36),
         bio, size=9.5, color=ACCENT, line=1.2)

# répartition des effectifs
card(s, MARGIN, Inches(4.62), Inches(6.2), Inches(1.85), WHITE, corner=0.05,
     border=LINE_LIGHT)
text(s, MARGIN + Inches(0.32), Inches(4.8), Inches(4), Inches(0.26),
     "Répartition des effectifs", size=10.5, color=GRAPHITE, bold=True,
     font=FONT_SEMI)
dist = [("R&D et produit", 24, ACCENT), ("Commercial et marketing", 10, TEAL),
        ("Delivery et support", 5, AMBER), ("Support corporate", 3, SLATE)]
total = sum(d[1] for d in dist)
for i, (lbl, n, col) in enumerate(dist):
    y = Inches(5.14) + i * Inches(0.32)
    text(s, MARGIN + Inches(0.32), y, Inches(2.1), Inches(0.24), lbl,
         size=9.5, color=SLATE)
    bw = Inches(2.6) * (n / total)
    rect(s, MARGIN + Inches(2.5), y + Inches(0.05), Inches(2.6), Inches(0.13),
         LINE_LIGHT, MSO_SHAPE.ROUNDED_RECTANGLE, 0.5)
    rect(s, MARGIN + Inches(2.5), y + Inches(0.05), bw, Inches(0.13), col,
         MSO_SHAPE.ROUNDED_RECTANGLE, 0.5)
    text(s, MARGIN + Inches(5.25), y, Inches(0.6), Inches(0.24), str(n),
         size=9.5, color=GRAPHITE, bold=True, font=FONT_SEMI)

# gouvernance
card(s, MARGIN + Inches(6.5), Inches(4.62), CONTENT_W - Inches(6.5),
     Inches(1.85), NAVY, corner=0.05, shadow=False)
text(s, MARGIN + Inches(6.82), Inches(4.8), Inches(4), Inches(0.26),
     "GOUVERNANCE", size=9.5, color=TEAL, bold=True, font=FONT_SEMI, spacing=2)
gov = [
    "Conseil stratégique trimestriel avec investisseurs et deux indépendants",
    "Revue de performance mensuelle sur 11 indicateurs engagés",
    "Comité produit bimensuel arbitrant la roadmap à 90 jours",
]
bullets(s, MARGIN + Inches(6.82), Inches(5.14), CONTENT_W - Inches(7.14),
        Inches(1.2), gov, size=10.5, color=RGBColor(0xC7, 0xD6, 0xE8),
        bullet_color=TEAL, gap=7)
notes(s, "Prévoir 20 secondes par profil maximum. La gouvernance est ce qui "
         "rassure le plus les investisseurs institutionnels.")

# ==========================================================================
# 16 — RISQUES & MITIGATION
# ==========================================================================
s = dark_slide(prs, "Risques identifiés et plans de mitigation",
               kicker="Maîtrise des risques",
               subtitle="Quatre risques suivis en comité, avec un responsable "
                        "et un indicateur d'alerte chacun.")
risks = [
    ("Élevé", CORAL, "Concentration client",
     "Les trois premiers clients représentent 34 % de l'ARR.",
     "Objectif < 20 % fin 2027 via 18 nouvelles signatures mid-market."),
    ("Moyen", AMBER, "Pression concurrentielle",
     "Entrée probable d'un éditeur généraliste sur le segment.",
     "Avance produit de 14 mois maintenue, verrou par la profondeur des "
     "connecteurs."),
    ("Moyen", AMBER, "Recrutement technique",
     "Tension sur les profils plateforme et données.",
     "Marque employeur, rémunération au 75e centile, 30 % de postes "
     "distribués."),
    ("Faible", TEAL, "Conformité et souveraineté",
     "Exigences réglementaires croissantes sur les données.",
     "ISO 27001 en cours, hébergement régional au choix dès le T4 2026."),
]
cw = (CONTENT_W - Inches(0.3)) / 2
for i, (level, col, title, desc, mit) in enumerate(risks):
    x = MARGIN + (i % 2) * (cw + Inches(0.3))
    y = Inches(2.32) + (i // 2) * Inches(2.05)
    card(s, x, y, cw, Inches(1.82), INK_SOFT, corner=0.06, shadow=False)
    rect(s, x, y, Pt(4), Inches(1.82), col)
    pill(s, x + Inches(0.36), y + Inches(0.26), Inches(0.82), Inches(0.26),
         level.upper(), col, INK, size=8)
    text(s, x + Inches(1.3), y + Inches(0.26), cw - Inches(1.7), Inches(0.3),
         title, size=13.5, color=WHITE, bold=True, font=FONT_SEMI)
    text(s, x + Inches(0.36), y + Inches(0.68), cw - Inches(0.72), Inches(0.36),
         desc, size=10.5, color=SLATE_DK, line=1.3)
    text(s, x + Inches(0.36), y + Inches(1.16), cw - Inches(0.72), Inches(0.5),
         [("Mitigation — ", {"bold": True, "font": FONT_SEMI, "color": col,
                             "size": 10.5}),
          (mit, {"color": RGBColor(0xC7, 0xD6, 0xE8), "size": 10.5})],
         line=1.3)
notes(s, "Montrer que les risques sont nommés, quantifiés et pilotés. "
         "Un risque non mentionné est un risque suspect.")

# ==========================================================================
# 17 — LA DEMANDE / NEXT STEPS
# ==========================================================================
s = light_slide(prs, "Ce que nous demandons",
                kicker="Décision attendue",
                subtitle="Un tour de série A de 8 M€ pour financer 24 mois "
                         "d'exécution jusqu'au point mort.")
alloc = [
    ("45 %", "Produit et R&D", "Moteur de règles v2, copilote décisionnel, "
     "souveraineté multi-région.", ACCENT),
    ("35 %", "Go-to-market", "Doublement de l'équipe commerciale, ouverture "
     "de deux marchés européens.", TEAL),
    ("20 %", "Delivery et conformité", "Certifications, support 24/5, "
     "industrialisation de l'onboarding.", AMBER),
]
cw = (Inches(8.1) - Inches(0.32) * 2) / 3
for i, (pct, title, body, col) in enumerate(alloc):
    x = MARGIN + i * (cw + Inches(0.32))
    card(s, x, Inches(2.3), cw, Inches(2.25), WHITE, corner=0.06,
         border=LINE_LIGHT)
    text(s, x + Inches(0.3), Inches(2.56), cw - Inches(0.6), Inches(0.5),
         pct, size=34, color=col, bold=True, font=FONT_SEMI, line=1.0)
    text(s, x + Inches(0.3), Inches(3.12), cw - Inches(0.6), Inches(0.46),
         title, size=12.5, color=GRAPHITE, bold=True, font=FONT_SEMI, line=1.1)
    rect(s, x + Inches(0.3), Inches(3.62), Inches(0.6), Pt(2.5), col)
    text(s, x + Inches(0.3), Inches(3.82), cw - Inches(0.6), Inches(0.66),
         body, size=10.5, color=SLATE, line=1.35)

card(s, MARGIN, Inches(4.78), Inches(8.1), Inches(1.65), NAVY, corner=0.06,
     shadow=False)
text(s, MARGIN + Inches(0.4), Inches(5.02), Inches(4), Inches(0.28),
     "PROCHAINES ÉTAPES", size=9.5, color=TEAL, bold=True, font=FONT_SEMI,
     spacing=2)
steps = [("1", "Accès data room", "sous 48 heures"),
         ("2", "Due diligence", "2 semaines"),
         ("3", "Term sheet", "mi-novembre"),
         ("4", "Closing", "31 décembre 2026")]
seg = (Inches(8.1) - Inches(0.8)) / 4
for i, (n, label, when) in enumerate(steps):
    x = MARGIN + Inches(0.4) + i * seg
    rect(s, x, Inches(5.45), Inches(0.26), Inches(0.26), TEAL,
         MSO_SHAPE.OVAL)
    text(s, x, Inches(5.49), Inches(0.26), Inches(0.2), n, size=9,
         color=INK, bold=True, font=FONT_SEMI, align=PP_ALIGN.CENTER)
    if i < 3:
        rect(s, x + Inches(0.3), Inches(5.57), seg - Inches(0.42), Pt(1),
             RGBColor(0x2C, 0x4A, 0x6B))
    text(s, x, Inches(5.8), seg - Inches(0.18), Inches(0.26), label,
         size=10.5, color=WHITE, bold=True, font=FONT_SEMI)
    text(s, x, Inches(6.08), seg - Inches(0.18), Inches(0.24), when,
         size=9.5, color=SLATE_DK)

rx = MARGIN + Inches(8.4)
rw = CONTENT_W - Inches(8.4)
card(s, rx, Inches(2.3), rw, Inches(4.13), WHITE, corner=0.06,
     border=LINE_LIGHT)
rect(s, rx, Inches(2.3), rw, Inches(0.09), ACCENT)
text(s, rx + Inches(0.34), Inches(2.62), rw - Inches(0.68), Inches(0.3),
     "TERMES INDICATIFS", size=9.5, color=ACCENT, bold=True, font=FONT_SEMI,
     spacing=2)
terms = [("Montant recherché", "8 M€"), ("Instrument", "Actions de préférence"),
         ("Tour précédent", "2,2 M€ (2024)"), ("Engagement existant", "2,5 M€"),
         ("Usage", "24 mois de runway"), ("Horizon de sortie", "2031 – 2032")]
for i, (k, v) in enumerate(terms):
    y = Inches(3.05) + i * Inches(0.55)
    text(s, rx + Inches(0.34), y, rw - Inches(0.68), Inches(0.24), k,
         size=10, color=SLATE)
    text(s, rx + Inches(0.34), y + Inches(0.21), rw - Inches(0.68), Inches(0.26),
         v, size=12.5, color=GRAPHITE, bold=True, font=FONT_SEMI)
notes(s, "Terminer sur la date de closing. Demander un engagement de "
         "calendrier, pas un avis.")

# ==========================================================================
# 18 — CLÔTURE
# ==========================================================================
s = blank(prs)
rect(s, 0, 0, SW, SH, INK)
picture_bg(s, CLOSING, prs)
scrim(s, prs, INK, 46000)
rect(s, 0, 0, Inches(0.13), SH, TEAL)

text(s, MARGIN, Inches(2.15), Inches(3), Inches(0.3),
     "MERCI DE VOTRE ATTENTION", size=11, color=TEAL, bold=True,
     font=FONT_SEMI, spacing=3.2)
text(s, MARGIN, Inches(2.6), Inches(10.4), Inches(1.7),
     [("Relier les systèmes.\n", {"size": 40, "color": WHITE,
                                  "font": FONT_LIGHT}),
      ("Rendre l'exécution évidente.", {"size": 40, "color": WHITE,
                                        "bold": True, "font": FONT_SEMI})],
     line=1.2)
rect(s, MARGIN, Inches(4.72), Inches(1.6), Pt(4), ACCENT)

text(s, MARGIN, Inches(5.06), Inches(8), Inches(0.5),
     "Nous restons à votre disposition pour approfondir tout volet de ce "
     "dossier : accès à la data room sous 48 heures.",
     size=13, color=SLATE_DK, line=1.4)

contacts = [("Contact investisseurs", "invest@shep.io"),
            ("Direction générale", "direction@shep.io"),
            ("Data room", "shep.io/dataroom")]
for i, (k, v) in enumerate(contacts):
    x = MARGIN + i * Inches(3.35)
    rect(s, x, Inches(5.85), Inches(2.9), Pt(0.75), RGBColor(0x24, 0x3C, 0x58))
    text(s, x, Inches(6.02), Inches(2.9), Inches(0.24), k, size=9.5,
         color=SLATE_DK, spacing=1.2)
    text(s, x, Inches(6.26), Inches(2.9), Inches(0.26), v, size=12.5,
         color=WHITE, bold=True, font=FONT_SEMI)
notes(s, "Laisser cette slide affichée pendant la session de questions.")

# ==========================================================================
prs.save(OUT)
print(f"✔ Deck généré : {OUT}")
print(f"  {len(prs.slides.__iter__.__self__._sldIdLst)} diapositives")
