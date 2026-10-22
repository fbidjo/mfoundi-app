"""Données saisies (état de l'application, version 2 : base B, taux engagés, simulateur) et passerelle vers le moteur de calcul.

L'état `S` reprend la structure de la démo HTML :
  params   : tiao, frais, maj, base, ref (AAAA-MM-JJ), auto, dev, parking,
             txeng, actu, kb, charges, effort_max (v2)
  groupes  : liste des groupes
  produits : id, cat, niv, sup, px (prix total, None pour un parking au prix par défaut)
  acq      : id, nom, g, vi, dur, d1 (AAAA-MM-JJ), ts (taux spécifique ou None)
  ven      : id (acquéreur), adr (produit), aj (ajustement de prix)
  pai      : id, n (échéance), d (date), m (montant)
  rem      : id, d (date), m (montant)
"""
import json
import os
from datetime import date

import numpy as np
import pandas as pd

from moteur import xround

CHEMIN_DEMO = os.path.join(os.path.dirname(os.path.abspath(__file__)), "donnees_demo.json")


PARAMS_DEFAUT = dict(
    txeng=0.0,        # taux d'intérêt des acquéreurs engagés : 0 % (prix contractuels, v3)
    actu=0.045,       # taux d'actualisation SIC = taux de neutralisation (VAN nulle)
    kb=0.728857,      # coefficient de la base de prix B (CA HT 16 346,6 M / grille 26 745,1 M)
    charges=99976.0,  # charges mensuelles de référence (simulateur)
    effort_max=0.45,  # taux d'effort éliminatoire
)


def completer(S):
    """Ajoute les paramètres introduits en version 2 aux anciennes données ou sauvegardes."""
    for k, v in PARAMS_DEFAUT.items():
        S["params"].setdefault(k, v)
    return S


def charger_demo():
    with open(CHEMIN_DEMO, encoding="utf-8") as f:
        return completer(json.load(f))


def _iso(x):
    return pd.Timestamp(x).strftime("%Y-%m-%d")


def _num(x):
    return None if x is None or pd.isna(x) else float(x)


def depuis_classeur(d, anonymiser=False):
    """Convertit le résultat de `moteur.lire_classeur` en état S."""
    P = d["params"]
    S = dict(
        params=dict(tiao=P["tiao"], frais=P["frais"], maj=P["majoration"], base=int(P["base_jours"]),
                    ref=_iso(P["date_ref"]), auto=False, dev=P["devise"], parking=P["prix_parking"]),
        groupes=list(P["groupes"]),
        produits=[], acq=[], ven=[], pai=[], rem=[],
    )
    completer(S)
    for r in d["produits"].itertuples():
        if pd.notna(r.prix_total):
            px = float(r.prix_total)
        elif pd.notna(r.prix_m2) and pd.notna(r.superficie):
            px = float(xround(r.superficie * r.prix_m2))
        else:
            px = None
        S["produits"].append(dict(id=str(r.adressage), cat=r.categorie, niv=r.niveau,
                                  sup=_num(r.superficie), px=px))
    for a in d["acquereurs"].itertuples():
        S["acq"].append(dict(id=str(a.id), nom="" if anonymiser else str(a.nom or ""), g=a.groupe,
                             vi=float(a.versement_initial), dur=int(a.duree), d1=_iso(a.date_1ere),
                             ts=_num(a.taux_specifique)))
    for v in d["ventes"].itertuples():
        S["ven"].append(dict(id=str(v.id), adr=str(v.adressage), aj=float(v.ajustement)))
    for p in d["paiements"].itertuples():
        S["pai"].append(dict(id=str(p.id), n=int(p.echeance), d=_iso(p.date), m=float(p.montant)))
    for r in d["remb"].itertuples():
        S["rem"].append(dict(id=str(r.id), d=_iso(r.date), m=float(r.montant)))
    return S


def reference(S):
    """Date de référence effective (aujourd'hui si le suivi en temps réel est activé)."""
    P = S["params"]
    return pd.Timestamp(date.today()) if P.get("auto") else pd.Timestamp(P["ref"])


def vers_moteur(S):
    """Construit le dictionnaire attendu par `moteur.calculer`."""
    P = S["params"]
    # v2 : intérêts au taux des acquéreurs engagés (0 % par défaut) ; pénalités à TIAO + frais + majoration
    params = dict(tiao=P["tiao"], frais=P["frais"], majoration=P["maj"], base_jours=int(P["base"]),
                  date_ref=reference(S), devise=P["dev"], prix_parking=float(P["parking"]),
                  taux_base=P.get("txeng", 0.0), taux_pen=P["tiao"] + P["frais"] + P["maj"],
                  groupes=list(S["groupes"]))

    prod = pd.DataFrame([dict(adressage=p["id"], categorie=p["cat"], niveau=p.get("niv"),
                              superficie=np.nan if p["sup"] is None else p["sup"],
                              prix_ht=p["px"] if p["px"] is not None
                              else (float(P["parking"]) if p["cat"] == "Parking" else 0.0))
                         for p in S["produits"]],
                        columns=["adressage", "categorie", "niveau", "superficie", "prix_ht"])
    acq = pd.DataFrame([dict(id=a["id"], nom=a.get("nom", ""), groupe=a["g"],
                             versement_initial=float(a["vi"]), duree=int(a["dur"]),
                             date_1ere=pd.Timestamp(a["d1"]),
                             taux_specifique=np.nan if a.get("ts") is None else float(a["ts"]))
                        for a in S["acq"]],
                       columns=["id", "nom", "groupe", "versement_initial", "duree", "date_1ere", "taux_specifique"])
    acq["date_1ere"] = pd.to_datetime(acq["date_1ere"])
    ven = pd.DataFrame([dict(id=v["id"], adressage=v["adr"], ajustement=float(v.get("aj", 0.0))) for v in S["ven"]],
                       columns=["id", "adressage", "ajustement"])
    pai = pd.DataFrame([dict(id=p["id"], echeance=int(p["n"]), date=pd.Timestamp(p["d"]), montant=float(p["m"]))
                        for p in S["pai"]], columns=["id", "echeance", "date", "montant"])
    pai["date"] = pd.to_datetime(pai["date"])
    rem = pd.DataFrame([dict(id=r["id"], date=pd.Timestamp(r["d"]), montant=float(r["m"])) for r in S["rem"]],
                       columns=["id", "date", "montant"])
    rem["date"] = pd.to_datetime(rem["date"])
    return dict(params=params, produits=prod, acquereurs=acq, ventes=ven, paiements=pai, remb=rem)


def sauvegarder(S):
    return json.dumps(S, ensure_ascii=False, indent=1).encode("utf-8")
