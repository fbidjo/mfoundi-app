"""Scoring des postulants acquéreurs et simulation du plan d'amortissement
Résidence Le Mfoundi (SIC) — mêmes règles que le classeur Excel Simulateur_Scoring_Acquereurs.xlsx

Module utilisé par l'onglet « Simulateur » de l'application (app.py).
Utilisation autonome :  python simulateur.py   (démonstration)
"""
from dataclasses import dataclass
from datetime import date
import calendar

# ------------------------------------------------------------------ paramètres
P = dict(
    taux_actualisation=0.045,   # taux d'actualisation SIC = taux de neutralisation (VAN nulle)
    charges=99_976,             # charges mensuelles de référence (F CFA)
    effort_elim=0.45,           # taux d'effort éliminatoire
    taux_scoring=0.055,         # taux utilisé pour estimer le taux d'effort lors du scoring
    k_accel=1.0,                # exposant d'accélération (1 = réduction proportionnelle)
    premiere_echeance=date(2027, 1, 1),
)
EFFORT = [(0.25, 35), (0.33, 28), (0.40, 15)]           # sinon 0
STABILITE = {"Public ou grande entreprise, >= 3 ans": 20, "Privé ou indépendant avec bilans, 1-3 ans": 14,
             "Revenus variables ou étrangers documentés": 10, "Autres": 4}
COMPORTEMENT = {"Aucun incident, solde moyen >= 3 mensualités": 15, "Incidents mineurs régularisés": 8, "Incident grave": 0}
GARANTIES = {"Toutes les garanties de la classe": 10, "Une garantie facultative manquante": 6, "Une garantie obligatoire manquante": 0}
PATRIMOINE = [(2, 10), (1, 7), (1e-6, 3)]               # ratio patrimoine net / capital financé
APPORT = [(0.50, 10), (0.40, 7), (0.35, 3)]
CLASSES = {  # score min, durée max (mois), taux de classe, apport min, garanties
    "A": (75, 120, 0.045, 0.25, "Garanties de base"),
    "B": (60, 84, 0.055, 0.35, "Billets à ordre obligatoires"),
    "C": (45, 60, 0.070, 0.50, "Caution solidaire + billets à ordre"),
    "D": (0, 0, None, None, "Comptant (remise de 2 %) ou refus"),
}
PERIODICITE = {"Mensuelle": 12, "Trimestrielle": 4}


@dataclass
class Postulant:
    id: str; nom: str; prix: float; revenu: float; mensualites_en_cours: float
    stabilite: str; comportement: str; patrimoine: float; apport: float; garanties: str
    duree: int = 84; periodicite: str = "Mensuelle"
    piece_fausse: bool = False; origine_fonds: bool = False; partie_liee: bool = False


def pmt(r, n, c):
    return c / n if r == 0 else c * r / (1 - (1 + r) ** -n)


def _pts(v, grille, ge=True):
    for seuil, pts in grille:
        if v >= seuil: return pts
    return 0


def taux_periodique(taux_annuel, m):
    """Neutralisation de la périodicité : taux annuel effectif identique à celui du plan mensuel."""
    return (1 + taux_annuel / 12) ** (12 / m) - 1


def taux_ajuste(classe, duree):
    """Taux de remboursement accéléré : de la classe (durée max) vers le taux de neutralisation (plancher)."""
    _, dmax, tc, _, _ = CLASSES[classe]
    neu = P["taux_actualisation"]
    return neu + (tc - neu) * min(1.0, duree / dmax) ** P["k_accel"]


def evaluer(p: Postulant):
    cap = p.prix * (1 - p.apport)
    mens_est = pmt(P["taux_scoring"] / 12, p.duree, cap)
    effort = (mens_est + P["charges"] + p.mensualites_en_cours) / p.revenu
    s_eff = 35 if effort <= .25 else 28 if effort <= .33 else 15 if effort <= .40 else 0
    ratio = p.patrimoine / cap
    pts = dict(effort=s_eff, stabilite=STABILITE[p.stabilite], comportement=COMPORTEMENT[p.comportement],
               patrimoine=_pts(ratio, PATRIMOINE), apport=_pts(p.apport, APPORT), garanties=GARANTIES[p.garanties])
    score = sum(pts.values())
    motifs = [m for m, c in (("Pièce fausse", p.piece_fausse), ("Origine des fonds", p.origine_fonds),
              ("Partie liée", p.partie_liee), ("Incident grave", p.comportement == "Incident grave"),
              ("Taux d'effort > 45 %", effort > P["effort_elim"])) if c]
    classe = "Refus" if motifs else "A" if score >= 75 else "B" if score >= 60 else "C" if score >= 45 else "D"
    r = dict(id=p.id, nom=p.nom, capital=cap, taux_effort=effort, points=pts, score=score, eliminatoires=motifs, classe=classe)
    if classe in "ABC" and classe != "Refus":
        sc, dmax, tc, amin, gar = CLASSES[classe]
        d = min(p.duree, dmax)
        if p.periodicite == "Trimestrielle": d = d // 3 * 3
        m = PERIODICITE[p.periodicite]
        t = taux_ajuste(classe, d)
        e = pmt(taux_periodique(t, m), d * m // 12, cap)
        alertes = []
        if p.apport < amin: alertes.append(f"apport inférieur à {amin:.0%}")
        if (e * m / 12 + P["charges"] + p.mensualites_en_cours) / p.revenu > P["effort_elim"]: alertes.append("taux d'effort final excessif")
        if d < p.duree: alertes.append("durée plafonnée")
        r.update(duree=d, taux_classe=tc, taux=t, echeance=e, periodicite=p.periodicite, garanties_exigees=gar,
                 taux_effort_final=(e * m / 12 + P["charges"] + p.mensualites_en_cours) / p.revenu, alertes=alertes,
                 decision="Accord sous conditions" if alertes else "Accord")
    else:
        r["decision"] = "Refus" if classe == "Refus" else CLASSES["D"][4]
    return r


def _edate(d, mois):
    y, m = divmod(d.month - 1 + mois, 12)
    y += d.year; m += 1
    return date(y, m, min(d.day, calendar.monthrange(y, m)[1]))


def plan(capital, taux, duree, periodicite="Mensuelle", depart=None):
    """Tableau d'amortissement + VAN de la marge financière actualisée au taux de neutralisation."""
    m = PERIODICITE[periodicite]; n = duree * m // 12
    rp = taux_periodique(taux, m); e = pmt(rp, n, capital)
    depart = depart or P["premiere_echeance"]; bal = capital; lignes = []; van = 0
    for k in range(1, n + 1):
        i = bal * rp; c = e - i
        van += e * (1 + P["taux_actualisation"]) ** (-k / m)
        lignes.append((k, _edate(depart, (k - 1) * 12 // m), bal, e, i, c, bal - c)); bal -= c
    return dict(lignes=lignes, echeance=e, total=e * n, interets=e * n - capital, van_marge=van - capital)


def comparatif(classe, capital, durees=(36, 48, 60, 72, 84, 96, 108, 120)):
    dmax, tc = CLASSES[classe][1], CLASSES[classe][2]
    out = []
    for d in durees:
        if d > dmax: continue
        t = taux_ajuste(classe, d)
        m_ = plan(capital, t, d, "Mensuelle"); q_ = plan(capital, t, d, "Trimestrielle")
        ref = plan(capital, tc, d, "Mensuelle")
        out.append(dict(duree=d, taux=t, mensualite=m_["echeance"], trimestrialite=q_["echeance"], interets_m=m_["interets"],
                        interets_q=q_["interets"], van_m=m_["van_marge"], van_q=q_["van_marge"], economie=ref["interets"] - m_["interets"]))
    return out


EXEMPLES = [  # postulants fictifs
    Postulant("P01", "Exemple 1", 140.4e6, 4_800_000, 0, list(STABILITE)[0], list(COMPORTEMENT)[0], 250e6, 0.35, list(GARANTIES)[0], 120),
    Postulant("P02", "Exemple 2", 140.4e6, 4_500_000, 150_000, list(STABILITE)[1], list(COMPORTEMENT)[1], 60e6, 0.35, list(GARANTIES)[1], 84),
    Postulant("P03", "Exemple 3", 166.5e6, 5_500_000, 0, list(STABILITE)[1], list(COMPORTEMENT)[0], 120e6, 0.40, list(GARANTIES)[0], 60, "Trimestrielle"),
    Postulant("P04", "Exemple 4", 234.7e6, 7_000_000, 400_000, list(STABILITE)[2], list(COMPORTEMENT)[1], 40e6, 0.35, list(GARANTIES)[1], 120),
    Postulant("P05", "Exemple 5", 45.9e6, 2_800_000, 50_000, list(STABILITE)[3], list(COMPORTEMENT)[0], 5e6, 0.50, list(GARANTIES)[0], 36),
    Postulant("P06", "Exemple 6", 140.4e6, 6_000_000, 0, list(STABILITE)[0], list(COMPORTEMENT)[2], 300e6, 0.50, list(GARANTIES)[0], 84),
]

if __name__ == "__main__":
    for p in EXEMPLES:
        r = evaluer(p)
        print(r["id"], r["classe"], r["score"], f"effort {r['taux_effort']:.1%}", r["decision"],
              (f"{r['duree']} mois {r['periodicite']} à {r['taux']:.3%} : {r['echeance']:,.0f} F" if "taux" in r else ""))
