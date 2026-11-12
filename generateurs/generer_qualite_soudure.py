# SPDX-License-Identifier: MIT
# -*- coding: utf-8 -*-
"""Jeu simulé de qualité de soudures : MecaIA-Qualite_soudure.xlsx.

Ce script fabrique, de bout en bout, le jeu que le livre Méca-Py emploie au chapitre 18 : une campagne d'essais de
900 soudures, 19 colonnes, une cible binaire (conforme ou non-conforme). Aucune donnée existante n'entre dans le calcul :
chaque valeur vient d'un tirage pseudo-aléatoire de graine fixée et des relations écrites plus bas.

Principe. Une soudure a un soudeur (qualification, expérience), une tôle (épaisseur, matériau), un procédé, des réglages
(courant, vitesse, préchauffage, nombre de passes), un contexte de fabrication (site, assemblage, robotisation, protection
gazeuse), un lot et un poste. La conformité se tire d'un score qui additionne des effets de qualification, de procédé, de
matériau, de protection gazeuse, de robotisation, de préchauffage, de nombre de passes, de lot, de poste et d'écart des
réglages à leur valeur de référence. Seule une soudure sur quatre a sa traçabilité (lot, poste) renseignée : le lot et le
poste agissent sur toutes les soudures, le fichier ne les montre que pour ces lignes.

Rôle pédagogique. Le jeu porte de quoi exercer les opérations du chapitre 18 : une colonne d'identifiant, une colonne
entièrement vide, deux colonnes de traçabilité très lacunaires, des cases vides dans l'expérience du soudeur (plus
fréquentes chez les débutants), quelques cases vides isolées dans deux colonnes de texte, des colonnes de texte et des
booléens à coder, des nombres d'échelles très différentes, et un contrôle visuel posé après la soudure qui reprend la
cible (variable à exclure).

Lecture du jeu. Les relations sont plausibles, sans être calées sur des essais réels. Les coefficients sont volontairement
forts (amplitude 2) pour que les relations ressortent dans 900 lignes : ils construisent un signal pédagogique et ne
mesurent aucun effet industriel. Les taux de conformité par qualification, procédé, matériau, protection gazeuse, lot ou
poste sont ceux du simulateur, pas des résultats d'essais ; les colonnes sont liées entre elles (qualification et épaisseur,
matériau et procédé), si bien qu'un taux d'ensemble ne dit pas à lui seul l'effet d'une colonne. gaz_protecteur_utilise vaut
False quand la protection gazeuse a manqué (débit nul ou insuffisant : bouteille vide, fuite, courant d'air). La colonne
controle_visuel_resultat recode la cible (OK pour conforme, KO pour non-conforme) : ce n'est pas une entrée, le contrôle est
posé après la soudure. type_gaz_protection est vide (rien à imputer). Le jeu sert à apprendre une méthode de préparation
des données, pas à régler un poste de soudage.

Reproductibilité. Les nombres aléatoires sont les uniformes du générateur PCG64 (graine 42), bornées à
[1e-12, 1 - 1e-12] pour éviter des quantiles infinis (troncature négligeable à cette échelle). Les lois (normale, gamma,
Poisson) s'obtiennent par leurs fonctions quantiles de SciPy, ce qui rend le résultat indépendant des évolutions des lois
de NumPy. Les valeurs, les types et les cases vides se retrouvent à l'identique avec les versions de requirements.txt
(numpy 2.4.6, pandas 3.0.5, scipy 1.18.0, openpyxl 3.1.5). Le classeur est écrit avec des propriétés fixes : ses octets
sont identiques d'une exécution à l'autre sur un même poste ; d'une installation à l'autre, l'égalité des octets dépend
aussi de la bibliothèque XML et de zlib, c'est pourquoi la commande verifier compare d'abord les contenus.

Usage, depuis un terminal :
    python generer_qualite_soudure.py                  écrit MecaIA-Qualite_soudure.xlsx dans le dossier courant
    python generer_qualite_soudure.py ecrire <dossier> écrit le classeur dans le dossier indiqué
    python generer_qualite_soudure.py verifier         régénère le jeu sans écrire de classeur, contrôle ses propriétés et,
                                                       si le classeur existe dans le dossier courant, compare son contenu
    python generer_qualite_soudure.py verifier <dossier>
"""
import argparse
import datetime
import hashlib
import io
import sys
import zipfile
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.packaging.core import DocumentProperties
from openpyxl.xml.functions import tostring
from scipy.optimize import brentq
from scipy.special import expit, gammaincinv, ndtri
from scipy.stats import poisson

GRAINE = 42
N_LIGNES = 900
NOM_FICHIER = "MecaIA-Qualite_soudure.xlsx"
FEUILLE = "soudage"

# --------------------------------------------------------------------------------------------------------------------
# Modalités et probabilités (chaque ligne d'un tableau suit l'ordre des modalités)
# --------------------------------------------------------------------------------------------------------------------
NIVEAUX = ("N1-Expert", "N2-Intermédiaire", "N3-Junior")
EPAISSEURS = (5, 8, 12)                                   # mm
MATERIAUX = ("acier carbone", "acier inox", "aluminium")
PROCEDES = ("MIG/MAG", "TIG")
ASSEMBLAGES = ("recouvrement", "bout_a_bout", "angle")
SITES = ("Le Havre", "Nantes", "Brest")
LOTS = ("LOT-A", "LOT-B", "LOT-C", "LOT-D", "LOT-E", "LOT-F", "LOT-G")
POSTES = ("Poste_A", "Poste_B", "Poste_C", "Poste_D", "Poste_E", "Poste_F")

P_NIVEAU = (0.22, 0.31, 0.47)
# Épaisseur selon la qualification (hypothèse du scénario) : un soudeur reçoit les tôles de la plage où il est qualifié.
P_EPAISSEUR = ((0.04, 0.14, 0.82), (0.12, 0.76, 0.12), (0.90, 0.08, 0.02))
# Matériau selon l'épaisseur : l'acier carbone domine dans les fortes épaisseurs.
P_MATERIAU = ((0.50, 0.32, 0.18), (0.62, 0.26, 0.12), (0.80, 0.14, 0.06))
# Probabilité du TIG selon le matériau : l'inox et l'aluminium se soudent plus souvent en TIG.
P_TIG = (0.18, 0.55, 0.78)
# Assemblage selon l'épaisseur.
P_ASSEMBLAGE = ((0.75, 0.12, 0.13), (0.55, 0.30, 0.15), (0.30, 0.50, 0.20))
P_SITE = (0.55, 0.30, 0.15)
P_ROBOT = (0.45, 0.30)                                    # selon le procédé : MIG/MAG, TIG
P_GAZ = 0.85                                              # protection gazeuse effective, quel que soit le procédé
# Expérience en années : loi gamma (forme, échelle) selon la qualification.
GAMMA_EXPERIENCE = ((12.0, 1.4), (10.0, 0.8), (4.0, 0.65))
P_EXPERIENCE_ABSENTE = (0.06, 0.14, 0.26)                 # case vide, selon la qualification
# Traçabilité : lot et poste ne sont renseignés que pour certaines soudures, surtout au site principal.
P_TRACEE = (0.36, 0.12, 0.05)
P_LOT = (0.10, 0.22, 0.28, 0.14, 0.12, 0.09, 0.05)
P_POSTE = (0.15, 0.25, 0.20, 0.18, 0.12, 0.10)
NB_VIDES_ASSEMBLAGE, NB_VIDES_SITE = 4, 3                 # cases vides isolées dans deux colonnes de texte

# Nombre de passes : 1 + Poisson(moyenne selon l'épaisseur), plus une demi-passe en moyenne de reprise au TIG.
MOYENNE_PASSES = (0.45, 1.5, 3.0)
# Préchauffage : probabilité d'en faire un, selon le matériau (lignes) et l'épaisseur (colonnes), et température de base.
P_PRECHAUFFAGE = ((0.10, 0.45, 0.80), (0.05, 0.05, 0.05), (0.15, 0.40, 0.70))
T_BASE = (90.0, 60.0, 70.0)                               # degrés Celsius, selon le matériau
# Réglages de référence. Courant (A) : (a + b * épaisseur) * facteur du matériau. Vitesse (mm/min) : (c + d * épaisseur),
# majorée de 10 % pour une soudure robotisée. Les réglages réels s'écartent de ces références, d'autant plus que le soudeur
# est moins qualifié (dispersion relative ECART_REGLAGES, réduite de 30 % en robotisé).
COURANT_REF = ((145.0, 5.0), (90.0, 4.5))                 # MIG/MAG, TIG : (a, b)
VITESSE_REF = ((450.0, -4.0), (200.0, -2.0))              # MIG/MAG, TIG : (c, d)
FACTEUR_COURANT_MATERIAU = (1.00, 0.95, 1.10)
ECART_REGLAGES = (0.05, 0.08, 0.13)

# Score de conformité, additif (échelle logistique), multiplié par AMPLITUDE. L'ordonnée à l'origine se règle pour que la
# probabilité moyenne de conformité vaille TAUX_CONFORME : la part de soudures conformes du jeu est un tirage autour de
# cette valeur.
TAUX_CONFORME = 0.40
EFFET_NIVEAU = (1.20, 0.20, -1.00)
EFFET_PROCEDE = (0.00, 1.40)                              # TIG mieux maîtrisé en essais
EFFET_MATERIAU = (0.00, -0.30, -1.20)
EFFET_SANS_GAZ = -2.00                                    # porosités : effet arbitrairement très fort
EFFET_ROBOT = 0.70
EFFET_ASSEMBLAGE = (0.00, -0.60, -0.20)
EFFET_SITE = (0.00, 0.30, -0.50)
EFFET_EXPERIENCE_PAR_AN = 0.03                            # écart d'expérience à la moyenne de la qualification
PENALITE_COURANT, PENALITE_VITESSE = 12.5, 5.0            # coefficients des écarts relatifs, au carré
EFFET_FISSURATION = (-1.40, 0.30)                         # acier carbone épais (8 et 12 mm) : sans, avec préchauffage
PENALITE_PASSE_MANQUANTE = 0.25
PASSES_REQUISES = (1, 2, 4)
EFFET_LOT = (0.0, 0.0, -0.70, 0.0, 0.0, 0.0, 0.50)
EFFET_POSTE = (0.0, 0.0, 0.0, -0.50, 0.0, 0.0)
AMPLITUDE = 2.0                                           # netteté du signal : multiplie la somme des effets

COLONNES = ("id_soudure", "qualite_soudure", "niveau_qualification_soudeur", "type_soudage_procede",
            "experience_soudeur_ans", "nb_passes_soudure", "temperature_prechauffage_C", "intensite_courant_A",
            "type_assemblage", "epaisseur_tole_mm", "materiau", "gaz_protecteur_utilise", "lot_production",
            "poste_soudage", "type_gaz_protection", "site_fabrication", "controle_visuel_resultat",
            "soudage_robotise", "vitesse_soudage_mm_min")


# --------------------------------------------------------------------------------------------------------------------
# Outils de tirage
# --------------------------------------------------------------------------------------------------------------------
class Tirages:
    """Uniformes du générateur PCG64, transformées en lois par leurs fonctions quantiles."""

    def __init__(self, graine):
        self.flux = np.random.Generator(np.random.PCG64(graine))

    def uniforme(self, n):
        return np.clip(self.flux.random(n), 1e-12, 1 - 1e-12)

    def normale(self, n):
        return ndtri(self.uniforme(n))

    def poisson(self, moyenne, n):
        return poisson.ppf(self.uniforme(n), moyenne)

    def indice(self, probabilites, n):
        """Indice de modalité de chaque ligne. `probabilites` : (k,) pour toutes les lignes, ou (n, k) ligne par ligne."""
        p = np.asarray(probabilites, dtype=float)
        if (p < 0).any() or (p.sum(axis=-1) <= 0).any():
            raise ValueError("probabilités négatives ou de somme nulle")
        cumul = np.cumsum(p, axis=-1)
        cumul = cumul / cumul[..., -1:]
        cumul = np.broadcast_to(cumul, (n, cumul.shape[-1]))
        return np.minimum((self.uniforme(n)[:, None] > cumul).sum(axis=1), cumul.shape[1] - 1)

    def lignes_au_hasard(self, k, n):
        """Indices de k lignes distinctes, tirées au hasard."""
        if not 0 <= k <= n:
            raise ValueError("%d lignes demandées parmi %d" % (k, n))
        return np.sort(np.argsort(self.uniforme(n))[:k])


# --------------------------------------------------------------------------------------------------------------------
# Simulation
# --------------------------------------------------------------------------------------------------------------------
def generer(n=N_LIGNES, graine=GRAINE):
    """Tableau de n soudures, colonnes dans l'ordre de COLONNES. Le résultat ne dépend que de n et de la graine."""
    if n < NB_VIDES_ASSEMBLAGE + NB_VIDES_SITE:
        raise ValueError("au moins %d lignes sont nécessaires" % (NB_VIDES_ASSEMBLAGE + NB_VIDES_SITE))
    t = Tirages(graine)

    # Qualification, expérience et épaisseur de tôle
    niveau = t.indice(P_NIVEAU, n)
    forme, echelle = np.array(GAMMA_EXPERIENCE)[niveau, 0], np.array(GAMMA_EXPERIENCE)[niveau, 1]
    experience_vraie = np.clip(echelle * gammaincinv(forme, t.uniforme(n)), 0.0, 38.0)
    i_epaisseur = t.indice(np.array(P_EPAISSEUR)[niveau], n)
    epaisseur = np.array(EPAISSEURS)[i_epaisseur]

    # Matériau, procédé, assemblage, site
    materiau = t.indice(np.array(P_MATERIAU)[i_epaisseur], n)
    procede = (t.uniforme(n) < np.array(P_TIG)[materiau]).astype(int)          # 0 : MIG/MAG, 1 : TIG
    assemblage = t.indice(np.array(P_ASSEMBLAGE)[i_epaisseur], n)
    site = t.indice(P_SITE, n)

    # Contexte de fabrication : robotisation, protection gazeuse
    robotise = t.uniforme(n) < np.array(P_ROBOT)[procede]
    gaz = t.uniforme(n) < P_GAZ

    # Réglages : passes, préchauffage, courant, vitesse
    passes = np.clip(1 + t.poisson(np.array(MOYENNE_PASSES)[i_epaisseur] + 0.5 * procede, n), 1, 8).astype(int)
    fait = t.uniforme(n) < np.array(P_PRECHAUFFAGE)[materiau, i_epaisseur]
    temperature = np.array(T_BASE)[materiau] + 4.0 * (epaisseur - 5) + 15.0 * t.normale(n)
    prechauffage = np.where(fait, np.clip(np.round(temperature), 40, 200), 0).astype(int)
    dispersion = np.array(ECART_REGLAGES)[niveau] * np.where(robotise, 0.7, 1.0)
    a, b = np.array(COURANT_REF)[procede].T
    courant_ref = (a + b * epaisseur) * np.array(FACTEUR_COURANT_MATERIAU)[materiau]
    courant = np.clip(np.round(courant_ref * (1 + dispersion * t.normale(n))), 80, 260).astype(int)
    c, d = np.array(VITESSE_REF)[procede].T
    vitesse_ref = (c + d * epaisseur) * np.where(robotise, 1.10, 1.0)
    vitesse = np.clip(np.round(vitesse_ref * np.exp(1.2 * dispersion * t.normale(n)), 1), 100.0, 600.0)

    # Lot et poste : tirés pour toutes les soudures ; la traçabilité n'en renseigne qu'une partie
    tracee = t.uniforme(n) < np.array(P_TRACEE)[site]
    lot = t.indice(P_LOT, n)
    poste = t.indice(P_POSTE, n)

    # Conformité : score additif, puis tirage de Bernoulli
    ecart_courant = (courant - courant_ref) / courant_ref
    ecart_vitesse = np.log(vitesse / vitesse_ref)
    carbone_epais = (materiau == 0) & (epaisseur >= 8)
    experience_moyenne = np.array([g[0] * g[1] for g in GAMMA_EXPERIENCE])[niveau]
    score = AMPLITUDE * (
        np.array(EFFET_NIVEAU)[niveau] + np.array(EFFET_PROCEDE)[procede]
        + np.array(EFFET_MATERIAU)[materiau] + np.where(gaz, 0.0, EFFET_SANS_GAZ)
        + np.where(robotise, EFFET_ROBOT, 0.0) + np.array(EFFET_ASSEMBLAGE)[assemblage]
        + np.array(EFFET_SITE)[site]
        + EFFET_EXPERIENCE_PAR_AN * (experience_vraie - experience_moyenne)
        - PENALITE_COURANT * ecart_courant ** 2 - PENALITE_VITESSE * ecart_vitesse ** 2
        + np.where(carbone_epais, np.where(prechauffage > 0, EFFET_FISSURATION[1], EFFET_FISSURATION[0]), 0.0)
        - PENALITE_PASSE_MANQUANTE * np.maximum(0, np.array(PASSES_REQUISES)[i_epaisseur] - passes)
        + np.array(EFFET_LOT)[lot] + np.array(EFFET_POSTE)[poste])

    def ecart_au_taux(origine):
        return expit(score + origine).mean() - TAUX_CONFORME

    if ecart_au_taux(-40.0) > 0 or ecart_au_taux(40.0) < 0:
        raise ValueError("TAUX_CONFORME hors d'atteinte pour ces effets")
    origine = brentq(ecart_au_taux, -40.0, 40.0)
    conforme = t.uniforme(n) < expit(score + origine)

    # Cases vides : expérience (selon la qualification), deux colonnes de texte (quelques lignes isolées)
    experience = np.round(experience_vraie, 1)
    experience[t.uniforme(n) < np.array(P_EXPERIENCE_ABSENTE)[niveau]] = np.nan
    vide_assemblage = t.lignes_au_hasard(NB_VIDES_ASSEMBLAGE, n)
    vide_site = t.lignes_au_hasard(NB_VIDES_SITE, n)

    def texte(valeurs, indices):
        return pd.Series(np.array(valeurs, dtype=object)[indices], dtype=object)

    df = pd.DataFrame({
        "id_soudure": np.arange(1, n + 1),
        "qualite_soudure": texte(("non-conforme", "conforme"), conforme.astype(int)),
        "niveau_qualification_soudeur": texte(NIVEAUX, niveau),
        "type_soudage_procede": texte(PROCEDES, procede),
        "experience_soudeur_ans": experience,
        "nb_passes_soudure": passes,
        "temperature_prechauffage_C": prechauffage,
        "intensite_courant_A": courant,
        "type_assemblage": texte(ASSEMBLAGES, assemblage),
        "epaisseur_tole_mm": epaisseur.astype(int),
        "materiau": texte(MATERIAUX, materiau),
        "gaz_protecteur_utilise": gaz,
        "lot_production": texte(LOTS, lot),
        "poste_soudage": texte(POSTES, poste),
        "type_gaz_protection": np.full(n, np.nan),
        "site_fabrication": texte(SITES, site),
        "controle_visuel_resultat": texte(("KO", "OK"), conforme.astype(int)),
        "soudage_robotise": robotise,
        "vitesse_soudage_mm_min": vitesse,
    })
    df.loc[~tracee, ["lot_production", "poste_soudage"]] = np.nan
    df.loc[vide_assemblage, "type_assemblage"] = np.nan
    df.loc[vide_site, "site_fabrication"] = np.nan
    assert tuple(df.columns) == COLONNES
    return df


# --------------------------------------------------------------------------------------------------------------------
# Écriture du classeur (octets identiques d'une exécution à l'autre, sur un même poste)
# --------------------------------------------------------------------------------------------------------------------
DATE_FIXE = datetime.datetime(2026, 10, 1, 0, 0, 0)
DATE_ZIP = (2026, 10, 1, 0, 0, 0)


def classeur_en_octets(df):
    """Classeur Excel d'une feuille `soudage` : en-têtes, puis une ligne par soudure ; case vide pour chaque manquant."""
    wb = Workbook()
    ws = wb.active
    ws.title = FEUILLE
    ws.append(list(df.columns))
    for ligne in df.itertuples(index=False, name=None):
        ws.append([None if (isinstance(v, float) and np.isnan(v)) else (v.item() if hasattr(v, "item") else v)
                   for v in ligne])
    for k, nom in enumerate(df.columns, start=1):
        ws.column_dimensions[ws.cell(row=1, column=k).column_letter].width = max(12, min(30, len(nom) + 5))
    tampon = io.BytesIO()
    wb.save(tampon)                      # openpyxl date la sauvegarde : on remplace donc les propriétés après coup
    tampon.seek(0)
    proprietes = DocumentProperties(creator="Meca-IA", lastModifiedBy="Meca-IA", title="MecaIA-Qualite_soudure",
                                    created=DATE_FIXE, modified=DATE_FIXE)
    sortie = io.BytesIO()
    with zipfile.ZipFile(tampon) as entree, zipfile.ZipFile(sortie, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as arc:
        for info in entree.infolist():
            octets = entree.read(info.filename)
            if info.filename == "docProps/core.xml":
                octets = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + tostring(proprietes.to_tree())
            zi = zipfile.ZipInfo(info.filename, date_time=DATE_ZIP)
            zi.compress_type, zi.create_system, zi.external_attr = zipfile.ZIP_DEFLATED, 0, 0
            arc.writestr(zi, octets, compress_type=zipfile.ZIP_DEFLATED, compresslevel=9)
    return sortie.getvalue()


def sha256(octets):
    return hashlib.sha256(octets).hexdigest()


def colonnes_differentes(df, relu):
    """Colonnes dont les valeurs, les cases vides ou le genre de type (entier, décimal, booléen, texte) diffèrent."""
    if list(df.columns) != list(relu.columns) or len(df) != len(relu):
        return ["forme du tableau"]
    genre = lambda s: "booléen" if s.dtype == bool else ("entier" if s.dtype.kind == "i" else ("décimal" if s.dtype.kind == "f" else "texte"))
    diff = []
    for c in df.columns:
        a, b = df[c], relu[c]
        if genre(a) != genre(b):
            diff.append(c)
        elif genre(a) in ("décimal", "entier", "booléen"):
            if not np.array_equal(a.to_numpy(float), b.to_numpy(float), equal_nan=True):
                diff.append(c)
        elif a.astype(object).where(a.notna(), None).tolist() != b.astype(object).where(b.notna(), None).tolist():
            diff.append(c)
    return diff


# --------------------------------------------------------------------------------------------------------------------
# Propriétés que le jeu doit garder (structure et rôle pédagogique)
# --------------------------------------------------------------------------------------------------------------------
def part_conforme_par(df, colonne):
    return (df.groupby(colonne)["qualite_soudure"].apply(lambda s: (s == "conforme").mean())).to_dict()


def proprietes(df):
    """Liste de (propriété, vrai ou faux, mesure) : ce que le jeu doit garder pour servir au chapitre 18."""
    n = len(df)
    res = []

    def noter(libelle, ok, mesure):
        res.append((libelle, bool(ok), mesure))

    noter("19 colonnes, dans l'ordre prévu", tuple(df.columns) == COLONNES, "%d colonnes" % df.shape[1])
    noter("900 lignes", n == N_LIGNES, "%d lignes" % n)
    noter("identifiant : 1 à n, sans doublon", df["id_soudure"].tolist() == list(range(1, n + 1)), "id_soudure")
    noter("colonne type_gaz_protection entièrement vide", df["type_gaz_protection"].isna().all(), "100 % de cases vides")
    controle = (df["controle_visuel_resultat"] == "OK") == (df["qualite_soudure"] == "conforme")
    noter("contrôle visuel : OK si et seulement si conforme", controle.all(), "%d lignes concordantes" % controle.sum())
    part = (df["qualite_soudure"] == "conforme").mean()
    noter("classe conforme minoritaire, entre 36 et 44 %", 0.36 <= part <= 0.44, "%.1f %%" % (100 * part))
    vides = df.isna().mean()
    noter("expérience : 12 à 25 % de cases vides", 0.12 <= vides["experience_soudeur_ans"] <= 0.25,
          "%.1f %%" % (100 * vides["experience_soudeur_ans"]))
    noter("lot et poste vides ensemble, 65 à 85 % des lignes",
          (df["lot_production"].isna() == df["poste_soudage"].isna()).all() and 0.65 <= vides["lot_production"] <= 0.85,
          "%.1f %%" % (100 * vides["lot_production"]))
    noter("quelques cases vides isolées (assemblage, site)",
          0 < df["type_assemblage"].isna().sum() <= 10 and 0 < df["site_fabrication"].isna().sum() <= 10,
          "%d et %d" % (df["type_assemblage"].isna().sum(), df["site_fabrication"].isna().sum()))
    autres = [c for c in ("nb_passes_soudure", "temperature_prechauffage_C", "intensite_courant_A", "epaisseur_tole_mm",
                          "vitesse_soudage_mm_min") if df[c].isna().any()]
    noter("aucune case vide dans les autres colonnes numériques", not autres, "colonnes concernées : %s" % (autres or "aucune"))
    noter("modalités de texte prévues",
          set(df["niveau_qualification_soudeur"]) == set(NIVEAUX) and set(df["type_soudage_procede"]) == set(PROCEDES)
          and set(df["materiau"]) == set(MATERIAUX) and set(df["type_assemblage"].dropna()) == set(ASSEMBLAGES)
          and set(df["site_fabrication"].dropna()) == set(SITES), "5 colonnes comparées")
    noter("plages physiques", df["intensite_courant_A"].between(80, 260).all()
          and df["vitesse_soudage_mm_min"].between(100, 600).all() and df["experience_soudeur_ans"].dropna().between(0, 38).all()
          and df["temperature_prechauffage_C"].between(0, 200).all() and df["nb_passes_soudure"].between(1, 8).all()
          and set(df["epaisseur_tole_mm"]) == set(EPAISSEURS), "courant, vitesse, expérience, préchauffage, passes")
    niveaux = part_conforme_par(df, "niveau_qualification_soudeur")
    noter("conformité décroissante de N1 à N3", niveaux["N1-Expert"] > niveaux["N2-Intermédiaire"] > niveaux["N3-Junior"],
          ", ".join("%s %.0f %%" % (k, 100 * v) for k, v in niveaux.items()))
    procedes = part_conforme_par(df, "type_soudage_procede")
    noter("TIG plus conforme que MIG/MAG", procedes["TIG"] > procedes["MIG/MAG"],
          ", ".join("%s %.0f %%" % (k, 100 * v) for k, v in procedes.items()))
    gaz = part_conforme_par(df, "gaz_protecteur_utilise")
    noter("soudure sans protection gazeuse moins conforme que sous protection", gaz[False] < gaz[True],
          "sans %.0f %%, avec %.0f %%" % (100 * gaz[False], 100 * gaz[True]))
    robot = part_conforme_par(df, "soudage_robotise")
    noter("soudure robotisée plus conforme que manuelle", robot[True] > robot[False],
          "robotisée %.0f %%, manuelle %.0f %%" % (100 * robot[True], 100 * robot[False]))
    noter("nombres d'échelles très différentes", df["vitesse_soudage_mm_min"].std() / df["nb_passes_soudure"].std() > 20,
          "écarts types %.0f et %.1f" % (df["vitesse_soudage_mm_min"].std(), df["nb_passes_soudure"].std()))
    return res


# --------------------------------------------------------------------------------------------------------------------
# Ligne de commande
# --------------------------------------------------------------------------------------------------------------------
def main(argv=None):
    p = argparse.ArgumentParser(description="Jeu simulé de qualité de soudures (MecaIA-Qualite_soudure.xlsx).")
    p.add_argument("action", nargs="?", choices=("ecrire", "verifier"), default="ecrire")
    p.add_argument("dossier", nargs="?", type=Path, default=Path("."), help="dossier du classeur (défaut : dossier courant)")
    args = p.parse_args(argv)
    df = generer()
    octets = classeur_en_octets(df)
    chemin = args.dossier / NOM_FICHIER
    if args.action == "ecrire":
        chemin.write_bytes(octets)
        print("%s : %d lignes, %d colonnes, %d octets, sha256 %s" % (chemin, len(df), df.shape[1], len(octets), sha256(octets)))
        return 0
    echecs = 0
    for libelle, ok, mesure in proprietes(df):
        print("%-4s %-60s %s" % ("OK" if ok else "KO", libelle, mesure))
        echecs += not ok
    rejeu = classeur_en_octets(generer()) == octets
    print("%-4s %-60s %s" % ("OK" if rejeu else "KO", "deux générations donnent les mêmes octets", sha256(octets)))
    echecs += not rejeu
    if chemin.is_file():
        diff = colonnes_differentes(df, pd.read_excel(chemin))
        print("%-4s %-60s %s" % ("OK" if not diff else "KO", "classeur présent : valeurs, types et cases vides identiques",
                                 chemin if not diff else "colonnes différentes : %s" % diff))
        echecs += bool(diff)
        print("%-4s %-60s %s" % ("INFO", "classeur présent : octets identiques à la génération",
                                 "oui" if chemin.read_bytes() == octets else "non (installation différente ?)"))
    print("%d propriété(s) en échec" % echecs)
    return 1 if echecs else 0


if __name__ == "__main__":
    sys.exit(main())
