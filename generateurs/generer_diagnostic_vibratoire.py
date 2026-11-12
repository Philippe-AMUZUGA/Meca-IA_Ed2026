# SPDX-License-Identifier: MIT
# -*- coding: utf-8 -*-
"""Jeu simulé de diagnostic vibratoire : MecaIA-Diagnostic_Vibratoire.xlsx.

Ce script fabrique, de bout en bout, le jeu que le livre Méca-IA emploie au chapitre 6 : 2 400 machines tournantes,
12 colonnes, une cible à six états (machine saine ou l'un de cinq défauts). Aucune donnée existante n'entre dans le calcul :
chaque valeur vient d'un tirage pseudo-aléatoire de graine fixée et des relations écrites plus bas.

Principe. Une machine a un facteur de charge (de 0,6 à 1,7), qui multiplie toutes ses amplitudes, et un niveau de base propre
à chaque indicateur, très resserré pour une machine saine. Un défaut ajoute à certaines raies une amplitude proportionnelle à
une sévérité log-normale de médiane 1 et très dispersée (un défaut naissant reste proche de l'état sain, un défaut avancé
domine le spectre) ; l'ajout se fait en quadrature, comme s'additionnent des signaux indépendants, avec une saturation douce
à 22 mm/s par raie. Six raies (0,5, 1 et 2 fois la fréquence de rotation, haute fréquence des roulements, fréquence
d'engrènement, bandes latérales) et le bruit large bande sont des valeurs efficaces en mm/s. Le niveau global est la valeur
efficace du signal entier : Vib_RMS_globale_mms au carré vaut la somme des carrés des sept indicateurs, plus un résidu qui
représente les autres raies non listées (harmoniques supérieures, bruit du capteur). Une raie ne dépasse donc jamais le niveau
global, ce que le script contrôle ligne par ligne. Kurtosis et facteur de crête se déduisent des mêmes composantes : des
sinusoïdes (kurtosis 1,5), un bruit gaussien (kurtosis 3) et, pour les défauts de roulement, de denture et d'accouplement, une
part de l'énergie portée par des chocs (kurtosis propre, grand). La formule du kurtosis d'une somme de composantes
indépendantes est exacte, celle du facteur de crête est une approximation de la somme des crêtes.

Rôle pédagogique. Le jeu sert à classer des états de machines tournantes : un indicateur levé par état (le déséquilibrage
lève la raie à une fois la rotation, le désalignement celle à deux fois, le défaut de roulement l'amplitude haute fréquence
et le kurtosis, le défaut de denture l'engrènement et ses bandes latérales), des états de transmission dont les signatures
se recouvrent (denture et jeu d'accouplement), des dispersions très différentes d'un état à l'autre (très peu pour une
machine saine, beaucoup pour chaque défaut), des amplitudes à queue lourde, des indicateurs d'ordres de
grandeur différents. Un identifiant sans contenu (Piece_ID) sert à montrer qu'une colonne se justifie ou s'écarte.

Lecture du jeu. Les relations sont plausibles, sans être calées sur des mesures : coefficients, sévérités et parts de chocs
sont des hypothèses pédagogiques, pas des résultats d'essais. Le jeu sert à apprendre une méthode de classification, pas à
régler un seuil d'alarme. Les amplitudes extrêmes (jusqu'à plusieurs dizaines de mm/s) représentent des défauts très avancés,
qu'une machine réelle ne laisserait pas atteindre. Le jeu diffusé est celui de la graine 42 ; les graines 0 à 39 servent à
contrôler que les propriétés ci-dessous ne tiennent pas à un tirage particulier (commande essais).

Reproductibilité. Les nombres aléatoires sont les uniformes du générateur PCG64 (graine 42), bornées à [1e-12, 1 - 1e-12]
pour éviter des quantiles infinis ; la loi normale s'obtient par sa fonction quantile de SciPy, ce qui rend le résultat
indépendant des évolutions des lois de NumPy. Les valeurs et les types se retrouvent à l'identique avec les versions de
requirements.txt (numpy 2.4.6, pandas 3.0.5, scipy 1.18.0, openpyxl 3.1.5). Le classeur est écrit avec des propriétés
fixes : ses octets sont identiques d'une exécution à l'autre sur un même poste ; d'une installation à l'autre, l'égalité
des octets dépend aussi de la bibliothèque XML et de zlib, c'est pourquoi la commande verifier compare d'abord les contenus.

Usage, depuis un terminal :
    python generer_diagnostic_vibratoire.py                  écrit MecaIA-Diagnostic_Vibratoire.xlsx dans le dossier courant
    python generer_diagnostic_vibratoire.py ecrire <dossier> écrit le classeur dans le dossier indiqué
    python generer_diagnostic_vibratoire.py verifier         régénère le jeu sans écrire de classeur, contrôle ses propriétés
                                                             et, si le classeur existe dans le dossier courant, compare son contenu
    python generer_diagnostic_vibratoire.py verifier <dossier>
    python generer_diagnostic_vibratoire.py essais           contrôle les propriétés sur les graines 0 à 39 (le jeu reste celui de 42)
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
from scipy.special import ndtri

GRAINE = 42
N_LIGNES = 2400
NOM_FICHIER = "MecaIA-Diagnostic_Vibratoire.xlsx"
FEUILLE = "Sheet1"
PREMIER_IDENTIFIANT = 1001

# --------------------------------------------------------------------------------------------------------------------
# États, indicateurs et paramètres
# --------------------------------------------------------------------------------------------------------------------
ETATS = ("Sain", "Jeu accouplement", "Déséquilibrage", "Défaut denture", "Défaut roulement", "Désalignement")
P_ETAT = (0.20, 0.16, 0.16, 0.16, 0.16, 0.16)

# Les sept indicateurs d'amplitude, dans l'ordre des colonnes : raies à 0,5, 1 et 2 fois la rotation, haute fréquence des
# roulements, fréquence d'engrènement, bandes latérales, bruit large bande.
RAIES = ("Amp_05x_tr_mms", "Amp_1x_tr_mms", "Amp_2x_tr_mms", "Amp_HF_roulement_mms", "Amp_freq_maillage_mms",
         "Amp_sidebands_mms", "Bruit_large_bande_mms")
COLONNES = ("Piece_ID", "Vib_RMS_globale_mms") + RAIES + ("Kurtosis", "Crest_factor", "Defaut")
SINUS = (True, True, True, False, True, True, False)      # raie sinusoïdale (kurtosis 1,5) ou aléatoire (kurtosis 3)

# Machine saine : valeur efficace de base (mm/s) et dispersion log-normale de chaque indicateur ; résidu des autres raies.
BASE = (0.10, 0.35, 0.18, 0.20, 0.32, 0.15, 0.55)
DISPERSION = (0.18, 0.10, 0.10, 0.12, 0.08, 0.14, 0.10)
RESIDU_BASE, RESIDU_DISPERSION, RESIDU_MAX = 0.15, 0.50, 3.0
DISPERSION_CHARGE = 0.30                                   # facteur de charge commun à tous les indicateurs
CHARGE_MIN, CHARGE_MAX = 0.6, 1.7                          # bornes du facteur de charge
AMPLITUDE_SATURATION = 22.0                                # ajout d'un défaut à une raie : saturation douce, en mm/s
DISPERSION_FORME = 0.25                                    # écart de chaque raie à la signature moyenne du défaut
SEVERITE_MAX = 15.0                                        # sévérité plafonnée

# Amplitude ajoutée (mm/s, en quadrature) par unité de sévérité, par état et par indicateur.
AJOUT = {
    "Sain":             (0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0),
    "Jeu accouplement": (1.35, 0.75, 1.20, 0.0, 1.50, 1.20, 0.0),
    "Déséquilibrage":   (0.0, 3.60, 0.45, 0.0, 0.0, 0.0, 0.0),
    "Défaut denture":   (0.0, 0.0, 0.0, 0.0, 4.50, 2.70, 0.0),
    "Défaut roulement": (0.0, 0.0, 0.0, 3.30, 0.0, 0.0, 1.35),
    "Désalignement":    (0.0, 1.50, 3.00, 0.0, 0.0, 0.0, 0.0),
}
# Dispersion log-normale de la sévérité (médiane 1) : un défaut se déclare de très faible à très fort, et reste proche de
# l'état sain tant qu'il est naissant.
SEVERITE = {"Sain": 0.0, "Jeu accouplement": 1.9, "Déséquilibrage": 1.9, "Défaut denture": 1.9,
            "Défaut roulement": 1.9, "Désalignement": 1.9}
# Part de l'énergie de chaque indicateur portée par des chocs, et kurtosis des chocs (médiane, dispersion log-normale).
PART_CHOCS = {
    "Jeu accouplement": (0.0, 0.0, 0.0, 0.0, 0.35, 0.35, 0.0),
    "Défaut denture":   (0.0, 0.0, 0.0, 0.0, 0.55, 0.55, 0.0),
    "Défaut roulement": (0.0, 0.0, 0.0, 0.90, 0.0, 0.0, 0.30),
}
KURTOSIS_CHOCS = {"Jeu accouplement": 9.0, "Défaut denture": 18.0, "Défaut roulement": 60.0}
DISPERSION_CHOCS = 0.25
PIC_GAUSSIEN = 3.8                                         # crête d'un bruit gaussien, en écarts types
DISPERSION_MESURE = 0.06                                   # dispersion relative de l'estimation du kurtosis et de la crête
KURTOSIS_MIN, CRETE_MIN = 1.5, float(np.sqrt(2.0))         # sinusoïde pure : kurtosis 1,5, facteur de crête racine de 2


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

    def indice(self, probabilites, n):
        """Indice de modalité de chaque ligne, pour des probabilités (k,) communes à toutes les lignes."""
        p = np.asarray(probabilites, dtype=float)
        if (p < 0).any() or p.sum() <= 0:
            raise ValueError("probabilités négatives ou de somme nulle")
        cumul = np.cumsum(p) / p.sum()
        return np.minimum((self.uniforme(n)[:, None] > cumul[None, :]).sum(axis=1), len(p) - 1)


# --------------------------------------------------------------------------------------------------------------------
# Simulation
# --------------------------------------------------------------------------------------------------------------------
def generer(n=N_LIGNES, graine=GRAINE):
    """Tableau de n machines, colonnes dans l'ordre de COLONNES. Le résultat ne dépend que de n et de la graine."""
    t = Tirages(graine)
    etat = t.indice(P_ETAT, n)
    charge = np.clip(np.exp(DISPERSION_CHARGE * t.normale(n)), CHARGE_MIN, CHARGE_MAX)
    severite = np.minimum(np.exp(np.array([SEVERITE[e] for e in ETATS])[etat] * t.normale(n)), SEVERITE_MAX)
    # Amplitude de chaque indicateur : base de la machine saine, puis ajout du défaut, en quadrature, par raie
    base = np.array(BASE)[None, :] * np.exp(np.array(DISPERSION)[None, :] * np.stack([t.normale(n) for _ in RAIES], axis=1))
    forme = np.exp(DISPERSION_FORME * np.stack([t.normale(n) for _ in RAIES], axis=1))
    ajout = np.array([AJOUT[e] for e in ETATS])[etat] * severite[:, None] * forme
    ajout = AMPLITUDE_SATURATION * np.tanh(ajout / AMPLITUDE_SATURATION)
    amplitudes = charge[:, None] * np.sqrt(base ** 2 + ajout ** 2)
    residu = charge * RESIDU_BASE * np.minimum(np.exp(RESIDU_DISPERSION * t.normale(n)), RESIDU_MAX)
    amplitudes = np.round(amplitudes, 3)
    # Niveau global : valeur efficace de tout le signal, jamais inférieure à celle d'une raie (arrondi par excès)
    energie_raies = (amplitudes ** 2).sum(axis=1)
    niveau = np.ceil(np.sqrt(energie_raies + residu ** 2) * 1e4) / 1e4

    # Kurtosis et facteur de crête, d'après les composantes : sinusoïdes, bruit gaussien, chocs
    part = np.array([PART_CHOCS.get(e, (0.0,) * 7) for e in ETATS])[etat]
    kurtosis_chocs = np.array([KURTOSIS_CHOCS.get(e, 3.0) for e in ETATS])[etat] * np.exp(DISPERSION_CHOCS * t.normale(n))
    kurtosis_chocs = np.maximum(kurtosis_chocs, 5.0)
    v = amplitudes ** 2
    v_chocs = (v * part).sum(axis=1)
    v_stationnaire = v * (1 - part)
    sinus = np.array(SINUS)[None, :]
    v_sinus = (v_stationnaire * sinus).sum(axis=1)
    v_gauss = (v_stationnaire * ~sinus).sum(axis=1) + residu ** 2
    total = v_sinus + v_gauss + v_chocs
    # composantes indépendantes : kurtosis total = 3 + somme de (kurtosis_i - 3) * v_i^2 / total^2, v_i variance de la composante i
    ecart = ((v_stationnaire * sinus) ** 2).sum(axis=1) * (KURTOSIS_MIN - 3.0) + (v_chocs ** 2) * (kurtosis_chocs - 3.0)
    kurtosis = 3.0 + ecart / total ** 2
    crete = (np.sqrt(2.0 * v_sinus) + 0.85 * (PIC_GAUSSIEN * np.sqrt(v_gauss) + np.sqrt(kurtosis_chocs * v_chocs))) / np.sqrt(total)
    kurtosis = np.maximum(KURTOSIS_MIN, np.round(kurtosis * np.exp(DISPERSION_MESURE * t.normale(n)), 2))
    crete = np.maximum(round(CRETE_MIN, 2), np.round(crete * np.exp(DISPERSION_MESURE * t.normale(n)), 2))

    df = pd.DataFrame({"Piece_ID": np.arange(PREMIER_IDENTIFIANT, PREMIER_IDENTIFIANT + n), "Vib_RMS_globale_mms": niveau})
    for k, nom in enumerate(RAIES):
        df[nom] = amplitudes[:, k]
    df["Kurtosis"] = kurtosis
    df["Crest_factor"] = crete
    df["Defaut"] = pd.Series(np.array(ETATS, dtype=object)[etat], dtype=object)
    assert tuple(df.columns) == COLONNES
    return df


# --------------------------------------------------------------------------------------------------------------------
# Écriture du classeur (octets identiques d'une exécution à l'autre, sur un même poste)
# --------------------------------------------------------------------------------------------------------------------
DATE_FIXE = datetime.datetime(2026, 10, 1, 0, 0, 0)
DATE_ZIP = (2026, 10, 1, 0, 0, 0)


def classeur_en_octets(df):
    """Classeur Excel d'une feuille `Sheet1` : en-têtes, puis une ligne par machine."""
    wb = Workbook()
    ws = wb.active
    ws.title = FEUILLE
    ws.append(list(df.columns))
    for ligne in df.itertuples(index=False, name=None):
        ws.append([v.item() if hasattr(v, "item") else v for v in ligne])
    for k, nom in enumerate(df.columns, start=1):
        ws.column_dimensions[ws.cell(row=1, column=k).column_letter].width = max(12, min(30, len(nom) + 5))
    tampon = io.BytesIO()
    wb.save(tampon)                      # openpyxl date la sauvegarde : on remplace donc les propriétés après coup
    tampon.seek(0)
    proprietes = DocumentProperties(creator="Meca-IA", lastModifiedBy="Meca-IA", title="MecaIA-Diagnostic_Vibratoire",
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
    """Colonnes dont les valeurs ou le genre de type (entier, décimal, texte) diffèrent."""
    if list(df.columns) != list(relu.columns) or len(df) != len(relu):
        return ["forme du tableau"]
    genre = lambda s: "entier" if s.dtype.kind == "i" else ("décimal" if s.dtype.kind == "f" else "texte")
    diff = []
    for c in df.columns:
        a, b = df[c], relu[c]
        if genre(a) != genre(b):
            diff.append(c)
        elif genre(a) in ("décimal", "entier"):
            if not np.array_equal(a.to_numpy(float), b.to_numpy(float), equal_nan=True):
                diff.append(c)
        elif a.astype(object).tolist() != b.astype(object).tolist():
            diff.append(c)
    return diff


# --------------------------------------------------------------------------------------------------------------------
# Propriétés que le jeu doit garder (cohérence physique, structure et rôle pédagogique)
# --------------------------------------------------------------------------------------------------------------------
SIGNATURE = {"Jeu accouplement": "Amp_05x_tr_mms", "Déséquilibrage": "Amp_1x_tr_mms", "Défaut denture": "Amp_freq_maillage_mms",
             "Défaut roulement": "Amp_HF_roulement_mms", "Désalignement": "Amp_2x_tr_mms"}


def proprietes(df, n=N_LIGNES):
    """Liste de (propriété, vrai ou faux, mesure) : ce que le jeu doit garder pour servir au chapitre 6."""
    res = []

    def noter(libelle, ok, mesure):
        res.append((libelle, bool(ok), mesure))

    amp = df[list(RAIES)]
    niveau = df["Vib_RMS_globale_mms"]
    noter("12 colonnes, dans l'ordre prévu", tuple(df.columns) == COLONNES, "%d colonnes" % df.shape[1])
    noter("%d lignes" % n, len(df) == n, "%d lignes" % len(df))
    noter("identifiant : %d à %d, sans doublon" % (PREMIER_IDENTIFIANT, PREMIER_IDENTIFIANT + n - 1),
          df["Piece_ID"].tolist() == list(range(PREMIER_IDENTIFIANT, PREMIER_IDENTIFIANT + n)), "Piece_ID")
    noter("aucune case vide", not df.isna().any().any(), "%d cases vides" % int(df.isna().sum().sum()))
    noter("types : un entier, dix décimaux, un texte",
          df["Piece_ID"].dtype.kind == "i" and all(df[c].dtype.kind == "f" for c in COLONNES[1:11])
          and df["Defaut"].map(type).eq(str).all(), "Piece_ID, 10 colonnes numériques, Defaut")
    parts = df["Defaut"].value_counts(normalize=True)
    noter("six états, Sain le plus fréquent (17 à 23 %), les autres de 13 à 19 %",
          set(parts.index) == set(ETATS) and 0.17 <= parts["Sain"] <= 0.23
          and all(0.13 <= parts[e] <= 0.19 for e in ETATS[1:]) and parts.idxmax() == "Sain",
          ", ".join("%s %.1f %%" % (e, 100 * parts[e]) for e in ETATS))
    # Cohérence énergétique
    depasse = (amp.to_numpy() > niveau.to_numpy()[:, None]).any(axis=1).sum()
    noter("aucune raie ni bruit au-dessus du niveau global", depasse == 0, "%d ligne(s) en défaut" % depasse)
    somme = (amp ** 2).sum(axis=1)
    noter("somme des carrés des sept indicateurs au plus égale au carré du niveau global",
          (somme <= niveau ** 2 + 1e-12).all(), "%d ligne(s) en défaut" % int((somme > niveau ** 2 + 1e-12).sum()))
    part_residu = 1 - somme / niveau ** 2
    noter("résidu d'énergie (autres raies) : médiane au-dessous de 5 %, jamais au-dessus de 60 %",
          part_residu.median() < 0.05 and part_residu.max() < 0.60,
          "médiane %.1f %%, maximum %.1f %%" % (100 * part_residu.median(), 100 * part_residu.max()))
    noter("niveau global : 0,3 à 120 mm/s", niveau.between(0.3, 120).all(), "%.3f à %.3f" % (niveau.min(), niveau.max()))
    noter("amplitudes positives, à trois décimales", (amp > 0).all().all() and np.array_equal(amp.to_numpy(), amp.round(3).to_numpy()),
          "minimum %.3f" % amp.min().min())
    noter("kurtosis au moins 1,5, facteur de crête au moins racine de 2",
          (df["Kurtosis"] >= 1.5).all() and (df["Crest_factor"] >= 1.41).all(),
          "kurtosis %.2f à %.2f, crête %.2f à %.2f" % (df["Kurtosis"].min(), df["Kurtosis"].max(),
                                                         df["Crest_factor"].min(), df["Crest_factor"].max()))
    sain = df[df["Defaut"] == "Sain"]
    noter("machine saine : kurtosis médian 2,4 à 3,6, facteur de crête médian 3 à 4,5",
          2.4 <= sain["Kurtosis"].median() <= 3.6 and 3.0 <= sain["Crest_factor"].median() <= 4.5,
          "kurtosis %.2f, crête %.2f" % (sain["Kurtosis"].median(), sain["Crest_factor"].median()))
    # Signatures par état
    med = df.groupby("Defaut")[list(RAIES) + ["Kurtosis", "Crest_factor"]].median()
    for etat, raie in SIGNATURE.items():
        lignes = df[df["Defaut"] == etat]
        seuil = sain[raie].quantile(0.95)
        part_etat = (lignes[raie] > seuil).mean()
        noter("%s : %s de médiane au moins 3 fois celle de l'état sain et au-dessus du 95e centile sain pour 55 à 99 %% des machines" % (etat, raie),
              med.loc[etat, raie] >= 3 * med.loc["Sain", raie] and 0.55 <= part_etat <= 0.99,
              "médiane %.3f contre %.3f, part %.0f %%" % (med.loc[etat, raie], med.loc["Sain", raie], 100 * part_etat))
    noter("défaut de roulement : kurtosis et facteur de crête médians au moins 2 fois ceux de l'état sain",
          med.loc["Défaut roulement", "Kurtosis"] >= 2 * med.loc["Sain", "Kurtosis"]
          and med.loc["Défaut roulement", "Crest_factor"] >= 1.6 * med.loc["Sain", "Crest_factor"],
          "kurtosis %.2f, crête %.2f" % (med.loc["Défaut roulement", "Kurtosis"], med.loc["Défaut roulement", "Crest_factor"]))
    noter("déséquilibrage : signal presque sinusoïdal, kurtosis médian inférieur à celui de l'état sain",
          med.loc["Déséquilibrage", "Kurtosis"] < med.loc["Sain", "Kurtosis"],
          "%.2f contre %.2f" % (med.loc["Déséquilibrage", "Kurtosis"], med.loc["Sain", "Kurtosis"]))
    # Recouvrement des deux défauts de transmission
    jeu, dent = df[df["Defaut"] == "Jeu accouplement"], df[df["Defaut"] == "Défaut denture"]
    haut_jeu = jeu["Amp_freq_maillage_mms"].quantile(0.90)
    part_dent = (dent["Amp_freq_maillage_mms"] <= haut_jeu).mean()
    noter("denture et jeu d'accouplement : engrènement qui se recouvre (40 à 90 % des dentures sous le 90e centile du jeu)",
          0.40 <= part_dent <= 0.90 and jeu["Amp_freq_maillage_mms"].median() >= 3 * med.loc["Sain", "Amp_freq_maillage_mms"],
          "%.0f %%" % (100 * part_dent))
    # Dispersions très différentes d'un état à l'autre
    ecarts = df.groupby("Defaut")["Amp_freq_maillage_mms"].std()
    noter("écart-type de l'engrènement : au moins 20 fois plus grand en défaut de denture qu'à l'état sain",
          ecarts["Défaut denture"] >= 20 * ecarts["Sain"], "%.3f contre %.3f" % (ecarts["Défaut denture"], ecarts["Sain"]))
    asym = df.loc[df["Defaut"] == "Défaut roulement", "Amp_HF_roulement_mms"]
    noter("amplitude haute fréquence des roulements en défaut : queue à droite (asymétrie au-dessus de 1)",
          asym.skew() > 1.0, "asymétrie %.2f" % asym.skew())
    noter("aucune ligne en double", not df.drop(columns="Piece_ID").duplicated().any(), "%d doublon(s)" % int(df.drop(columns="Piece_ID").duplicated().sum()))
    return res


# --------------------------------------------------------------------------------------------------------------------
# Ligne de commande
# --------------------------------------------------------------------------------------------------------------------
def main(argv=None):
    p = argparse.ArgumentParser(description="Jeu simulé de diagnostic vibratoire (MecaIA-Diagnostic_Vibratoire.xlsx).")
    p.add_argument("action", nargs="?", choices=("ecrire", "verifier", "essais"), default="ecrire")
    p.add_argument("dossier", nargs="?", type=Path, default=Path("."), help="dossier du classeur (défaut : dossier courant)")
    args = p.parse_args(argv)
    if args.action == "essais":
        echecs = 0
        for graine in range(40):
            df = generer(graine=graine)
            ko = [libelle for libelle, ok, _ in proprietes(df) if not ok]
            echecs += len(ko)
            print("%-4s graine %2d : %d propriété(s) en échec%s" % ("OK" if not ko else "KO", graine, len(ko),
                                                                   (" : " + " ; ".join(ko)) if ko else ""))
        print("%d échec(s) au total sur 40 graines" % echecs)
        return 1 if echecs else 0
    df = generer()
    octets = classeur_en_octets(df)
    chemin = args.dossier / NOM_FICHIER
    if args.action == "ecrire":
        chemin.write_bytes(octets)
        print("%s : %d lignes, %d colonnes, %d octets, sha256 %s" % (chemin, len(df), df.shape[1], len(octets), sha256(octets)))
        return 0
    echecs = 0
    for libelle, ok, mesure in proprietes(df):
        print("%-4s %-60s %s" % ("OK" if ok else "KO", libelle[:60], mesure))
        echecs += not ok
    rejeu = classeur_en_octets(generer()) == octets
    print("%-4s %-60s %s" % ("OK" if rejeu else "KO", "deux générations donnent les mêmes octets", sha256(octets)))
    echecs += not rejeu
    if chemin.is_file():
        diff = colonnes_differentes(df, pd.read_excel(chemin))
        print("%-4s %-60s %s" % ("OK" if not diff else "KO", "classeur présent : valeurs et types identiques",
                                 chemin if not diff else "colonnes différentes : %s" % diff))
        echecs += bool(diff)
        print("%-4s %-60s %s" % ("INFO", "classeur présent : octets identiques à la génération",
                                 "oui" if chemin.read_bytes() == octets else "non (installation différente ?)"))
    print("%d propriété(s) en échec" % echecs)
    return 1 if echecs else 0


if __name__ == "__main__":
    sys.exit(main())
