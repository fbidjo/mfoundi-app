"""Application Streamlit – Plans d'amortissement · Résidence Mfoundi (SIC).

Reprend les fonctionnalités de la démo HTML : vue globale, par groupe, plan individuel, acquéreurs,
paiements, remboursements anticipés et paramètres. Les données de départ viennent du classeur Excel
(donnees_demo.json) ; les noms d'acquéreurs sont laissés vides.

Version 2 (base B) : acquéreurs engagés à taux 0 % (prix contractuels), pénalités inchangées, et nouvel onglet
« Simulateur » (scoring du postulant, taux automatique selon la durée, périodicité mensuelle ou trimestrielle).
"""
import io
import json
from datetime import date

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

import etat
import simulateur as sim
from moteur import calculer, controles, lire_classeur

st.set_page_config(page_title="Plans d'amortissement – Mfoundi", page_icon="🏢", layout="wide")

COULEUR_STATUT = {"À jour": "#dff2e8", "Payé": "#dff2e8", "En retard": "#fbe3e0", "Impayé": "#fbe3e0",
                  "Partiel": "#fbefd5"}
TEAL, AMBRE = "#0b6e6e", "#cf8a2b"


# ------------------------------------------------------------------ outils
def nombre(x):
    return f"{x:,.0f}".replace(",", " ")


def argent(x):
    return nombre(x) + " " + S["params"]["dev"]


def pct(x):
    return "–" if x is None or pd.isna(x) else f"{x * 100:.1f}".replace(".", ",") + " %"


def dfr(t):
    return "" if t is None or pd.isna(t) else pd.Timestamp(t).strftime("%d/%m/%Y")


def lab(id_, nom):
    return f"{id_} · {nom}" if nom else id_


def montrer(df, formats=None, col_statut=None, lignes_rouges=None):
    """Affiche un tableau (sans index), nombres formatés, statuts colorés."""
    s = df.style.format(formats or {}, na_rep="–")
    if lignes_rouges is not None:
        s = s.apply(lambda r: ["background-color:#fbe3e0;color:#14212b" if lignes_rouges(r) else "" for _ in r], axis=1)
    if col_statut:
        f = lambda v: f"background-color:{COULEUR_STATUT[v]};color:#14212b" if v in COULEUR_STATUT else ""
        s = (s.map if hasattr(s, "map") else s.applymap)(f, subset=[col_statut])
    st.dataframe(s, hide_index=True)


def remplacer(nouvel_etat):
    """Remplace toutes les données et réinitialise les champs de saisie."""
    st.session_state.S = nouvel_etat
    for k in list(st.session_state):
        if k.startswith(("par_", "acq_", "pay_", "rem_")):
            del st.session_state[k]
    st.rerun()


def encaissements(plan_df, ref, titre):
    """Mensualités attendues par mois sur les 24 mois suivant la date de référence."""
    debut = pd.Timestamp(ref.year, ref.month, 1) + pd.DateOffset(months=1)
    mois = [debut + pd.DateOffset(months=i) for i in range(24)]
    f = plan_df[(plan_df.date_echeance > ref) & (plan_df.date_echeance <= ref + pd.DateOffset(months=24))
                & (plan_df.mensualite_due > 0)] if len(plan_df) else plan_df
    tot = f.groupby(f.date_echeance.dt.to_period("M")).mensualite_due.sum() if len(f) else pd.Series(dtype=float)
    y = [float(tot.get(m.to_period("M"), 0.0)) for m in mois]
    fig = go.Figure(go.Bar(x=[m.strftime("%m/%y") for m in mois], y=y, marker_color=TEAL, name="Encaissement prévu"))
    fig.update_layout(title=titre, height=330, margin=dict(t=50, b=30, l=10, r=10))
    st.plotly_chart(fig)
    st.caption(f"Mensualités attendues par mois, en {S['params']['dev']}.")


# ------------------------------------------------------------------ état et calcul
if "S" not in st.session_state:
    st.session_state.S = etat.charger_demo()
S = st.session_state.S


@st.cache_data(show_spinner="Calcul des plans d'amortissement…")
def calc(s_json, jour):
    d = etat.vers_moteur(json.loads(s_json))
    return (*calculer(d), controles(d))


@st.cache_data(show_spinner=False)
def classeur_resultats(s_json, jour):
    synth, plan, ventes, _ = calc(s_json, jour)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter", datetime_format="dd/mm/yyyy") as xw:
        synth.to_excel(xw, sheet_name="Synthese_Acquereurs", index=False)
        plan.to_excel(xw, sheet_name="Echeanciers", index=False)
        ventes.to_excel(xw, sheet_name="Ventes_Detail", index=False)
        for ws in xw.sheets.values():
            ws.set_column(0, 20, 18)
    return buf.getvalue()


# ------------------------------------------------------------------ barre latérale
st.sidebar.header("Données")
st.sidebar.caption("Données de démonstration issues du classeur Excel. Les modifications restent dans votre "
                   "session : utilisez la sauvegarde pour les conserver.")
st.sidebar.download_button("⬇️ Sauvegarder les données (JSON)", data=etat.sauvegarder(S),
                           file_name="sauvegarde_mfoundi.json", mime="application/json")
with st.sidebar.expander("Charger une sauvegarde ou un classeur"):
    up_json = st.file_uploader("Sauvegarde (.json)", type=["json"], key="up_json")
    if up_json is not None and st.button("Charger la sauvegarde"):
        try:
            data = etat.completer(json.loads(up_json.getvalue().decode("utf-8")))
            etat.vers_moteur(data)  # validation de la structure
        except Exception as exc:
            st.error(f"Sauvegarde illisible : {exc}")
        else:
            remplacer(data)
    up_xlsx = st.file_uploader("Classeur Excel (Plan_Amortissement_Acquereurs)", type=["xlsx"], key="up_xlsx")
    if up_xlsx is not None and st.button("Importer le classeur"):
        try:
            nouveau = etat.depuis_classeur(lire_classeur(io.BytesIO(up_xlsx.getvalue())))
        except Exception as exc:
            st.error(f"Impossible de lire le classeur : {exc}")
        else:
            remplacer(nouveau)

st.title("🏢 Plans d'amortissement · Résidence Mfoundi")
st.caption("Vente de logements à tempérament, suivi des acquéreurs · version 2 · base B")
if st.session_state.get("flash"):
    st.success(st.session_state.pop("flash"))

tabs = st.tabs(["Vue globale", "Par groupe", "Plan individuel", "Acquéreurs", "Paiements", "Remboursements",
                "Simulateur", "Paramètres"])
t_glob, t_grp, t_plan, t_acq, t_pai, t_rem, t_sim, t_par = tabs

# ------------------------------------------------------------------ Paramètres (rendus en premier : le calcul en dépend)
with t_par:
    P = S["params"]
    st.subheader("Paramètres de calcul")
    st.caption("Chaque changement recalcule immédiatement tous les plans.")
    c1, c2, c3 = st.columns(3)
    P["tiao"] = c1.number_input("Taux directeur BEAC (%)", 0.0, 50.0, round(P["tiao"] * 100, 3), 0.025,
                                format="%.3f", key="par_tiao") / 100
    P["frais"] = c2.number_input("Frais de gestion (%)", 0.0, 50.0, round(P["frais"] * 100, 3), 0.025,
                                 format="%.3f", key="par_frais") / 100
    P["maj"] = c3.number_input("Majoration de pénalité (%)", 0.0, 50.0, round(P["maj"] * 100, 3), 0.025,
                               format="%.3f", key="par_maj") / 100
    c4, c5, c6 = st.columns(3)
    P["base"] = int(c4.number_input("Base de calcul des jours", 1, 400, int(P["base"]), 5, key="par_base"))
    P["parking"] = float(c5.number_input("Prix d'un parking par défaut", 0, 1_000_000_000, int(P["parking"]),
                                         500_000, key="par_pk"))
    P["auto"] = c6.checkbox("Date du jour (suivi en temps réel)", P["auto"], key="par_auto")
    ref_saisie = c6.date_input("Date de référence", date.fromisoformat(P["ref"]), disabled=P["auto"],
                               format="DD/MM/YYYY", key="par_ref")
    if not P["auto"]:
        P["ref"] = ref_saisie.isoformat()
    st.markdown("**Acquéreurs engagés et simulateur (version 2)**")
    d1, d2, d3 = st.columns(3)
    P["txeng"] = d1.number_input("Taux d'intérêt des acquéreurs engagés (%)", 0.0, 50.0,
                                 round(P["txeng"] * 100, 3), 0.025, format="%.3f", key="par_txeng") / 100
    P["actu"] = d2.number_input("Taux d'actualisation SIC (%)", 0.0, 50.0, round(P["actu"] * 100, 3), 0.025,
                                format="%.3f", key="par_actu") / 100
    P["kb"] = d3.number_input("Coefficient base B", 0.1, 1.5, float(P["kb"]), 0.001, format="%.6f", key="par_kb")
    d4, d5, _ = st.columns(3)
    P["charges"] = float(d4.number_input("Charges mensuelles de référence", 0, 10_000_000, int(P["charges"]), 1000,
                                         key="par_chg"))
    P["effort_max"] = d5.number_input("Taux d'effort éliminatoire (%)", 10.0, 100.0, round(P["effort_max"] * 100, 1),
                                      1.0, key="par_eff") / 100
    tb = P["txeng"]
    tpen = P["tiao"] + P["frais"] + P["maj"]
    st.info(f"Taux d'intérêt des acquéreurs engagés : {pct(tb)} (prix contractuels conservés) · taux de pénalité : "
            f"{pct(tpen)} · pénalité = impayé × taux × jours de retard ÷ {P['base']}.")
    if st.button("Réinitialiser la démonstration"):
        remplacer(etat.charger_demo())

# ------------------------------------------------------------------ calcul
ref = etat.reference(S)
tb = S["params"]["txeng"]
tp = S["params"]["tiao"] + S["params"]["frais"] + S["params"]["maj"]
synth, plan, ventes, alertes = calc(json.dumps(S, sort_keys=True), date.today().isoformat())
a_des_acq = len(synth) > 0
noms = {a["id"]: a.get("nom", "") for a in S["acq"]}
libelle = lambda i: lab(i, noms.get(i, ""))

st.sidebar.download_button("⬇️ Résultats complets (Excel)",
                           data=classeur_resultats(json.dumps(S, sort_keys=True), date.today().isoformat()),
                           file_name="Resultats_Amortissement.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
if alertes:
    with st.sidebar.expander(f"⚠️ {len(alertes)} alerte(s)", expanded=True):
        for a in alertes:
            st.write("•", a)
else:
    st.sidebar.success("Aucun contrôle en erreur")

VIDE = "Aucun acquéreur. Ajoutez-en un dans l'onglet Acquéreurs."

# ------------------------------------------------------------------ Vue globale
with t_glob:
    if not a_des_acq:
        st.info(VIDE)
    else:
        echu, imp = synth.echu_a_date.sum(), synth.impayes.sum()
        late = synth[synth.statut == "En retard"]
        k = st.columns(5)
        k[0].metric("Acquéreurs", len(synth))
        k[1].metric("Encours (capital restant dû)", argent(synth.capital_restant_du.sum()))
        k[2].metric("Impayés et pénalités", argent(synth.total_du.sum()))
        k[3].metric("Taux de recouvrement", pct(1 - imp / echu) if echu else "–")
        k[4].metric("Acquéreurs en retard", len(late))
        st.caption(f"Situation au {dfr(ref)} · taux appliqué {pct(tb)} · taux de pénalité {pct(tp)}")

        st.subheader("À relancer")
        if len(late):
            t = pd.DataFrame({"Acquéreur": [libelle(i) for i in late.id], "Groupe": late.groupe.values,
                              "Mensualité": late.mensualite.values, "Impayé": late.impayes.values,
                              "Pénalité": late.penalites.values, "Total dû": late.total_du.values})
            montrer(t, {c: nombre for c in ("Mensualité", "Impayé", "Pénalité", "Total dû")})
        else:
            st.info("Aucun impayé à la date de référence.")

        g1, g2 = st.columns(2)
        with g1:
            st.subheader("Par groupe")
            g = synth.groupby("groupe").agg(a=("id", "count"), s=("surface_m2", "sum"),
                                            e=("capital_restant_du", "sum"), d=("total_du", "sum")).reset_index()
            g.columns = ["Groupe", "Acq.", "Surface m²", "Encours", "Dû"]
            montrer(g, {c: nombre for c in g.columns[1:]})
        with g2:
            st.subheader("Par catégorie de produit")
            c = ventes.groupby("categorie").agg(n=("adressage", "count"), p=("prix_net", "sum")).reset_index()
            c.columns = ["Catégorie", "Produits", "Prix net HT"]
            montrer(c, {"Produits": nombre, "Prix net HT": nombre})
        st.subheader("Encaissements prévus, 24 mois")
        encaissements(plan, ref, "Encaissements prévus du portefeuille")

# ------------------------------------------------------------------ Par groupe
with t_grp:
    groupes = sorted(set(synth.groupe)) if a_des_acq else []
    if not groupes:
        st.info("Aucun groupe.")
    else:
        grp = st.selectbox("Groupe", groupes)
        s = synth[synth.groupe == grp]
        k = st.columns(4)
        k[0].metric("Acquéreurs", len(s))
        k[1].metric("Surface", f"{s.surface_m2.sum():,.0f} m²".replace(",", " "))
        k[2].metric("Encours", argent(s.capital_restant_du.sum()))
        k[3].metric("Impayés et pénalités", argent(s.total_du.sum()))
        t = pd.DataFrame({"Acquéreur": [libelle(i) for i in s.id], "Surface m²": s.surface_m2.values,
                          "Prix": s.prix_bien.values, "Mensualité": s.mensualite.values,
                          "Capital restant": s.capital_restant_du.values, "Dû": s.total_du.values,
                          "Statut": s.statut.values})
        montrer(t, {"Surface m²": "{:,.1f}", "Prix": nombre, "Mensualité": nombre, "Capital restant": nombre,
                    "Dû": nombre}, col_statut="Statut")
        st.subheader("Encaissements prévus, 24 mois")
        encaissements(plan[plan.id.isin(s.id)], ref, f"Encaissements prévus – {grp}")

# ------------------------------------------------------------------ Plan individuel
with t_plan:
    if not a_des_acq:
        st.info(VIDE)
    else:
        id_ = st.selectbox("Acquéreur", list(synth.id), format_func=libelle, key="sel_plan")
        s = synth[synth.id == id_].iloc[0]
        e = plan[(plan.id == id_) & (plan.mensualite_due > 0)].copy()
        st.markdown(f"**Statut : {s.statut}**")
        k = st.columns(4)
        k[0].metric("Prix du bien", argent(s.prix_bien))
        k[1].metric("Capital financé", argent(s.capital_finance))
        k[2].metric("Mensualité", argent(s.mensualite))
        k[3].metric("Capital restant dû", argent(s.capital_restant_du))
        k = st.columns(4)
        k[0].metric("Taux annuel", pct(s.taux_annuel))
        k[1].metric("Versement initial", nombre(s.versement_initial))
        k[2].metric("Échéances", int(s.nb_echeances))
        k[3].metric("Fin prévue", dfr(s.date_fin) or "–")
        k = st.columns(4)
        k[0].metric("Total payé", nombre(s.total_paye))
        k[1].metric("Impayés", nombre(s.impayes))
        k[2].metric("Pénalités", nombre(s.penalites))
        k[3].metric("Intérêts totaux", nombre(s.interets_totaux))
        prods = ", ".join(ventes.loc[ventes.id == id_, "adressage"].astype(str))
        st.caption("Produits acquis : " + (prods or "aucun"))

        if len(e):
            ca, cb = st.columns(2)
            lbl = e.date_echeance.dt.strftime("%m/%y")
            f1 = go.Figure()
            f1.add_bar(x=lbl, y=e.capital_amorti, name="Capital", marker_color=TEAL)
            f1.add_bar(x=lbl, y=e.interets, name="Intérêts", marker_color=AMBRE)
            f1.update_layout(barmode="stack", title="Composition des mensualités", height=330,
                             margin=dict(t=50, b=30, l=10, r=10), legend=dict(orientation="h", y=-0.2))
            ca.plotly_chart(f1)
            f2 = go.Figure(go.Scatter(x=lbl, y=e.capital_fin, fill="tozeroy", line=dict(color=TEAL), name="Capital"))
            f2.update_layout(title="Capital restant dû", height=330, margin=dict(t=50, b=30, l=10, r=10))
            cb.plotly_chart(f2)

            st.subheader("Échéancier")
            t = pd.DataFrame({"N°": e.n.values, "Date": e.date_echeance.map(dfr).values,
                              "Capital début": e.capital_debut.values, "Intérêts": e.interets.values,
                              "Mensualité": e.mensualite_due.values, "Capital amorti": e.capital_amorti.values,
                              "Remb. anticipé": e.remb_anticipe.values, "Capital fin": e.capital_fin.values,
                              "Payé": e.paye.values, "Impayé": e.impaye.values, "Pénalité": e.penalite.values,
                              "Statut": e.statut.values})
            fm = {c: nombre for c in t.columns if c not in ("N°", "Date", "Statut")}
            montrer(t, fm, col_statut="Statut")
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine="xlsxwriter") as xw:
                t.to_excel(xw, sheet_name="Echeancier", index=False)
                xw.sheets["Echeancier"].set_column(0, 12, 16)
            st.download_button("⬇️ Télécharger cet échéancier (Excel)", data=buf.getvalue(),
                               file_name=f"Echeancier_{id_}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ------------------------------------------------------------------ Acquéreurs
with t_acq:
    st.subheader(f"Acquéreurs ({len(S['acq'])})")
    if a_des_acq:
        t = pd.DataFrame({"ID": synth.id.values, "Nom": synth.nom.values, "Groupe": synth.groupe.values,
                          "Prix": synth.prix_bien.values, "Capital financé": synth.capital_finance.values,
                          "Mensualité": synth.mensualite.values, "Fin prévue": synth.date_fin.map(dfr).values,
                          "Capital restant": synth.capital_restant_du.values, "Dû": synth.total_du.values,
                          "Statut": synth.statut.values})
        montrer(t, {c: nombre for c in ("Prix", "Capital financé", "Mensualité", "Capital restant", "Dû")},
                col_statut="Statut")
    else:
        st.info(VIDE)

    action = st.radio("Action", ["Ajouter", "Modifier", "Supprimer"], horizontal=True, key="acq_action")
    prod_par_id = {p["id"]: p for p in S["produits"]}
    prix_defaut = lambda p: p["px"] if p["px"] is not None else (S["params"]["parking"] if p["cat"] == "Parking" else 0)

    if action == "Supprimer":
        if not S["acq"]:
            st.info("Rien à supprimer.")
        else:
            cible = st.selectbox("Acquéreur à supprimer", [a["id"] for a in S["acq"]], format_func=libelle, key="acq_del")
            ok = st.checkbox("Je confirme la suppression (acquéreur, ventes, paiements et remboursements liés)",
                             key=f"acq_delok_{cible}")
            if st.button("Supprimer", disabled=not ok):
                for cle in ("acq", "ven", "pai", "rem"):
                    S[cle] = [x for x in S[cle] if x["id"] != cible]
                st.session_state["flash"] = f"{cible} supprimé."
                st.rerun()
    else:
        modif = action == "Modifier"
        a = None
        if modif:
            if not S["acq"]:
                st.info("Aucun acquéreur à modifier : le formulaire ci-dessous crée un nouvel acquéreur.")
                modif = False
            else:
                cible = st.selectbox("Acquéreur à modifier", [x["id"] for x in S["acq"]], format_func=libelle,
                                     key="acq_mod")
                a = next(x for x in S["acq"] if x["id"] == cible)
        suffixe = a["id"] if a else "new"
        vendus_ailleurs = {v["adr"] for v in S["ven"] if not a or v["id"] != a["id"]}
        miens = [v["adr"] for v in S["ven"] if a and v["id"] == a["id"]]
        dispo = [p["id"] for p in S["produits"] if p["id"] not in vendus_ailleurs]

        with st.form(f"acq_form_{suffixe}"):
            st.markdown(f"**{'Modifier ' + a['id'] if a else 'Nouvel acquéreur'}**")
            f1, f2, f3 = st.columns(3)
            nom = f1.text_input("Nom ou raison sociale (optionnel)", a["nom"] if a else "")
            grp_opts = S["groupes"] or ["-"]
            grp = f2.selectbox("Groupe", grp_opts, index=grp_opts.index(a["g"]) if a and a["g"] in grp_opts else 0)
            vi = f3.number_input("Versement initial", 0, 10_000_000_000, int(a["vi"]) if a else 0, 1000)
            f4, f5, f6 = st.columns(3)
            dur = f4.number_input("Durée en mois (1 à 180)", 1, 180, int(a["dur"]) if a else 120, 1)
            d1 = f5.date_input("Date de la 1ère échéance",
                               date.fromisoformat(a["d1"]) if a else (ref + pd.DateOffset(months=1)).date(),
                               format="DD/MM/YYYY")
            ts = f6.number_input("Taux annuel spécifique en % (0 = taux standard)", 0.0, 50.0,
                                 round(a["ts"] * 100, 2) if a and a["ts"] is not None else 0.0, 0.01)
            prods = st.multiselect(
                "Produits achetés", dispo, default=[p for p in miens if p in dispo],
                format_func=lambda i: f"{i} · {prod_par_id[i]['cat']}"
                                      + (f" · {prod_par_id[i]['sup']} m²" if prod_par_id[i]["sup"] else "")
                                      + f" · {nombre(prix_defaut(prod_par_id[i]))}")
            envoi = st.form_submit_button("Enregistrer")
        if envoi:
            if not prods:
                st.error("Choisissez au moins un produit.")
            else:
                id_new = a["id"] if a else "ACQ-%03d" % (1 + max([0] + [int(x["id"][4:]) for x in S["acq"]
                                                                       if x["id"][4:].isdigit()]))
                rec = dict(id=id_new, nom=nom.strip(), g=grp, vi=float(vi), dur=int(dur), d1=d1.isoformat(),
                           ts=(ts / 100 if ts > 0 else None))
                if a:
                    S["acq"][S["acq"].index(a)] = rec
                else:
                    S["acq"].append(rec)
                S["ven"] = [v for v in S["ven"] if v["id"] != id_new] + [dict(id=id_new, adr=p, aj=0.0) for p in prods]
                st.session_state["flash"] = f"{id_new} enregistré."
                st.rerun()

# ------------------------------------------------------------------ Paiements
with t_pai:
    if not a_des_acq:
        st.info(VIDE)
    else:
        retard = synth[synth.statut == "En retard"]
        defaut = list(synth.id).index(retard.id.iloc[0]) if len(retard) else 0
        st.subheader("Enregistrer un paiement")
        pid = st.selectbox("Acquéreur", list(synth.id), index=defaut, format_func=libelle, key="pay_id")
        ech = plan[(plan.id == pid) & (plan.mensualite_due > 0) & (plan.date_echeance <= ref)]
        non_soldees = ech[ech.paye < ech.mensualite_due]
        n_def = int(non_soldees.n.iloc[0]) if len(non_soldees) else len(ech) + 1
        cle = f"{pid}_{len(S['pai'])}"
        c1, c2, c3 = st.columns(3)
        n = int(c1.number_input("N° d'échéance", 1, 1000, n_def, 1, key=f"pay_n_{cle}"))
        ligne = plan[(plan.id == pid) & (plan.n == n)]
        du = float(ligne.mensualite_due.iloc[0]) if len(ligne) else 0.0
        dpay = c2.date_input("Date de paiement", ref.date(), format="DD/MM/YYYY", key=f"pay_d_{cle}")
        m = c3.number_input(f"Montant (dû : {nombre(du)})", 0, 10_000_000_000, int(round(du)), 1000,
                            key=f"pay_m_{cle}_{n}")
        if st.button("Enregistrer le paiement"):
            if m <= 0:
                st.error("Le montant doit être positif.")
            else:
                S["pai"].append(dict(id=pid, n=n, d=dpay.isoformat(), m=float(m)))
                st.session_state["flash"] = f"Paiement de {nombre(m)} enregistré pour {libelle(pid)}, échéance {n}."
                st.rerun()

        st.subheader(f"Derniers paiements ({len(S['pai'])})")
        liste = sorted(enumerate(S["pai"]), key=lambda x: x[1]["d"], reverse=True)[:40]
        if liste:
            montrer(pd.DataFrame({"Date": [dfr(p["d"]) for _, p in liste],
                                  "Acquéreur": [libelle(p["id"]) for _, p in liste],
                                  "Échéance": [p["n"] for _, p in liste],
                                  "Montant": [p["m"] for _, p in liste]}), {"Montant": nombre})
            choix = st.selectbox("Supprimer un paiement", [None] + [i for i, _ in liste], key=f"pay_del_{len(S['pai'])}",
                                 format_func=lambda i: "–" if i is None else
                                 f"{dfr(S['pai'][i]['d'])} · {libelle(S['pai'][i]['id'])} · éch. {S['pai'][i]['n']} · {nombre(S['pai'][i]['m'])}")
            if choix is not None and st.button("Supprimer ce paiement"):
                S["pai"].pop(choix)
                st.session_state["flash"] = "Paiement supprimé."
                st.rerun()

# ------------------------------------------------------------------ Remboursements anticipés
with t_rem:
    if not a_des_acq:
        st.info(VIDE)
    else:
        st.subheader("Remboursement anticipé")
        st.caption("Le montant réduit le capital à l'échéance suivante. La mensualité reste la même et la durée raccourcit.")
        cle = len(S["rem"])
        c1, c2, c3 = st.columns(3)
        rid = c1.selectbox("Acquéreur", list(synth.id), format_func=libelle, key="rem_id")
        rd = c2.date_input("Date", ref.date(), format="DD/MM/YYYY", key=f"rem_d_{cle}")
        rm = c3.number_input("Montant", 0, 10_000_000_000, 5_000_000, 1000, key=f"rem_m_{cle}")
        if st.button("Enregistrer le remboursement"):
            if rm <= 0:
                st.error("Le montant doit être positif.")
            else:
                S["rem"].append(dict(id=rid, d=rd.isoformat(), m=float(rm)))
                st.session_state["flash"] = f"Remboursement anticipé de {nombre(rm)} enregistré."
                st.rerun()
        st.subheader(f"Remboursements enregistrés ({len(S['rem'])})")
        if S["rem"]:
            montrer(pd.DataFrame({"Date": [dfr(r["d"]) for r in S["rem"]],
                                  "Acquéreur": [libelle(r["id"]) for r in S["rem"]],
                                  "Montant": [r["m"] for r in S["rem"]]}), {"Montant": nombre})
            sup = st.selectbox("Supprimer un remboursement", [None] + list(range(len(S["rem"]))), key=f"rem_del_{cle}",
                               format_func=lambda i: "–" if i is None else
                               f"{dfr(S['rem'][i]['d'])} · {libelle(S['rem'][i]['id'])} · {nombre(S['rem'][i]['m'])}")
            if sup is not None and st.button("Supprimer ce remboursement"):
                S["rem"].pop(sup)
                st.session_state["flash"] = "Remboursement supprimé."
                st.rerun()
        else:
            st.info("Aucun remboursement anticipé.")

# ------------------------------------------------------------------ Simulateur (nouveaux acquéreurs)
with t_sim:
    P = S["params"]
    sim.P.update(taux_actualisation=P["actu"], charges=P["charges"], effort_elim=P["effort_max"])
    st.subheader("Simulateur de mensualité et scoring d'un postulant")
    st.caption("Le taux baisse automatiquement quand la durée raccourcit, jusqu'au taux d'actualisation SIC "
               f"({pct(P['actu'])}, VAN nulle). Périodicité trimestrielle : même taux annuel effectif que le mensuel.")
    vendus = {v["adr"] for v in S["ven"]}
    libres = [p for p in S["produits"] if p["id"] not in vendus and p["px"] is not None]
    s1, s2 = st.columns(2)
    choix = s1.selectbox("Produit disponible", ["Prix libre"] + [p["id"] for p in libres], key="sim_prod",
                         format_func=lambda i: i if i == "Prix libre" else
                         next(f"{p['id']} · {p['cat']} · {nombre(p['px'])}" for p in libres if p["id"] == i))
    base = s2.radio("Base de prix", ["Base B", "Prix affichés"], horizontal=True, key="sim_base")
    px0 = next((p["px"] for p in libres if p["id"] == choix), 140_000_000.0)
    px0 = px0 * (P["kb"] if base == "Base B" else 1.0)
    prix = float(st.number_input("Prix du lot", 0, 5_000_000_000, int(round(px0)), 100_000, key=f"sim_prix_{choix}_{base}"))

    c1, c2, c3, c4 = st.columns(4)
    apport = c1.slider("Apport (%)", 20, 60, 35, key="sim_apport") / 100
    duree = int(c2.slider("Durée (mois)", 12, 120, 84, 3, key="sim_duree"))
    per = c3.radio("Périodicité", list(sim.PERIODICITE), horizontal=True, key="sim_per")
    depart = c4.date_input("Première échéance", date(2027, 1, 5), format="DD/MM/YYYY", key="sim_dep")

    st.markdown("**Profil du postulant** (exemple modifiable)")
    q1, q2, q3 = st.columns(3)
    revenu = float(q1.number_input("Revenu net mensuel", 0, 1_000_000_000, 5_800_000, 50_000, key="sim_rev"))
    enc = float(q2.number_input("Mensualités en cours", 0, 100_000_000, 0, 10_000, key="sim_enc"))
    patr = float(q3.number_input("Patrimoine net", 0, 10_000_000_000, 120_000_000, 5_000_000, key="sim_pat"))
    q4, q5, q6 = st.columns(3)
    stab = q4.selectbox("Stabilité des revenus", list(sim.STABILITE), key="sim_stab")
    comp = q5.selectbox("Comportement bancaire", list(sim.COMPORTEMENT), key="sim_comp")
    gar = q6.selectbox("Garanties offertes", list(sim.GARANTIES), key="sim_gar")
    e1, e2, e3 = st.columns(3)
    pf = e1.checkbox("Pièce fausse", key="sim_pf")
    of = e2.checkbox("Origine des fonds non justifiée", key="sim_of")
    pl = e3.checkbox("Partie liée non déclarée", key="sim_pl")
    m1, m2 = st.columns(2)
    imposee = m1.selectbox("Classe retenue", ["Selon le scoring", "A", "B", "C"], key="sim_cls")
    mode = m2.radio("Taux", ["Automatique", "Manuel"], horizontal=True, key="sim_mode")
    tman = (m2.number_input("Taux annuel (%)", 0.0, 30.0, 5.5, 0.1, key="sim_tman") / 100) if mode == "Manuel" else None

    post = sim.Postulant("SIM", "Simulation", prix, max(revenu, 1.0), enc, stab, comp, patr, apport, gar,
                         duree, per, pf, of, pl)
    ev = sim.evaluer(post)
    classe = ev["classe"] if imposee == "Selon le scoring" else imposee
    k = st.columns(4)
    k[0].metric("Score", f"{ev['score']} / 100")
    k[1].metric("Classe du scoring", ev["classe"])
    k[2].metric("Taux d'effort estimé", pct(ev["taux_effort"]))
    k[3].metric("Capital financé", argent(ev["capital"]))
    if ev["eliminatoires"]:
        st.error("Critère éliminatoire : " + ", ".join(ev["eliminatoires"]))
    st.caption("Points : " + " · ".join(f"{n} {v}" for n, v in ev["points"].items()))

    if classe not in ("A", "B", "C"):
        st.warning("Classe D ou refus : paiement comptant avec remise de 2 %, ou refus. "
                   "Choisissez une classe dans « Classe retenue » pour simuler tout de même un plan.")
    else:
        sc, dmax, tc, amin, garex = sim.CLASSES[classe]
        d = min(duree, dmax) if mode == "Automatique" else duree
        if per == "Trimestrielle":
            d = d // 3 * 3
        t = sim.taux_ajuste(classe, d) if mode == "Automatique" else tman
        m = sim.PERIODICITE[per]
        cap = ev["capital"]
        res = sim.plan(cap, t, d, per, depart)
        eff = (res["echeance"] * m / 12 + P["charges"] + enc) / max(revenu, 1.0)
        if d < duree:
            st.info(f"Durée ramenée à {d} mois (maximum de la classe {classe}).")
        if apport < amin - 1e-9 and mode == "Automatique":
            st.warning(f"Apport inférieur au minimum de la classe {classe} ({amin:.0%}).")
        if eff > P["effort_max"]:
            st.error(f"Taux d'effort final de {pct(eff)} : au-dessus du seuil éliminatoire.")
        st.caption(f"Garanties exigées en classe {classe} : {garex}.")
        k = st.columns(4)
        k[0].metric("Échéance " + ("mensuelle" if m == 12 else "trimestrielle"), argent(res["echeance"]))
        k[1].metric("Taux appliqué", pct(t), f"classe {classe} : {pct(tc)}", delta_color="off")
        k[2].metric("Total payé", argent(res["total"]), f"dont intérêts {nombre(res['interets'])}", delta_color="off")
        k[3].metric("VAN de la marge SIC", argent(res["van_marge"]))
        k = st.columns(3)
        k[0].metric("Taux d'effort final", pct(eff))
        k[1].metric("Revenu net requis (33 %)", argent((res["echeance"] * m / 12 + P["charges"] + enc) / 0.33))
        k[2].metric("Échéances", len(res["lignes"]))

        L = pd.DataFrame(res["lignes"], columns=["N°", "Date", "Capital début", "Échéance", "Intérêts",
                                                 "Capital remboursé", "Capital fin"])
        g1, g2 = st.columns(2)
        lbl = [x.strftime("%m/%y") for x in L["Date"]]
        f1 = go.Figure()
        f1.add_bar(x=lbl, y=L["Capital remboursé"], name="Capital", marker_color=TEAL)
        f1.add_bar(x=lbl, y=L["Intérêts"], name="Intérêts", marker_color=AMBRE)
        f1.update_layout(barmode="stack", title="Composition des échéances", height=330,
                         margin=dict(t=50, b=30, l=10, r=10), legend=dict(orientation="h", y=-0.2))
        g1.plotly_chart(f1)
        f2 = go.Figure(go.Scatter(x=lbl, y=L["Capital fin"], fill="tozeroy", line=dict(color=TEAL)))
        f2.update_layout(title="Capital restant dû", height=330, margin=dict(t=50, b=30, l=10, r=10))
        g2.plotly_chart(f2)

        st.subheader("Comparatif des durées")
        cp = pd.DataFrame(sim.comparatif(classe, cap))
        if len(cp):
            cp["rev"] = (cp.mensualite + P["charges"] + enc) / 0.33
            tcp = pd.DataFrame({"Durée (mois)": cp.duree, "Taux": cp.taux.map(pct),
                                "Mensualité": cp.mensualite, "Trimestrialité": cp.trimestrialite,
                                "Intérêts (mensuel)": cp.interets_m, "VAN marge (mensuel)": cp.van_m,
                                "Économie d'intérêts": cp.economie, "Revenu net requis": cp.rev})
            montrer(tcp, {c: nombre for c in tcp.columns if c not in ("Durée (mois)", "Taux")})
        st.subheader("Tableau d'amortissement")
        T = L.copy(); T["Date"] = T["Date"].map(lambda x: x.strftime("%d/%m/%Y"))
        montrer(T, {c: nombre for c in T.columns if c not in ("N°", "Date")})
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="xlsxwriter") as xw:
            T.to_excel(xw, sheet_name="Simulation", index=False)
            xw.sheets["Simulation"].set_column(0, 7, 18)
        st.download_button("⬇️ Télécharger cette simulation (Excel)", data=buf.getvalue(),
                           file_name="Simulation_echeancier.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
