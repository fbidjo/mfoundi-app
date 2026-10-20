"""Application Streamlit – Plan d'amortissement (vente à tempérament) · SIC · Résidence Mfoundi."""
import io
import os

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from moteur import calculer, controles, lire_classeur

st.set_page_config(page_title="Plan d'amortissement – Mfoundi", page_icon="🏢", layout="wide")

DEFAUT = os.path.join("data", "Plan_Amortissement_Acquereurs.xlsx")


@st.cache_data(show_spinner="Calcul des plans d'amortissement…")
def traiter(contenu: bytes):
    d = lire_classeur(io.BytesIO(contenu))
    alertes = controles(d)
    synth, plan, ventes = calculer(d)
    return d, alertes, synth, plan, ventes


def nombre(x, devise):
    return f"{x:,.0f}".replace(",", " ") + " " + devise


# ---------- Source des données ----------
st.title("🏢 Plan d'amortissement – Résidence Mfoundi (SIC)")
st.sidebar.header("Données")
fichier = st.sidebar.file_uploader("Classeur Excel (Plan_Amortissement_Acquereurs)", type=["xlsx"])

if fichier is not None:
    contenu = fichier.getvalue()
elif os.path.exists(DEFAUT):
    with open(DEFAUT, "rb") as f:
        contenu = f.read()
    st.sidebar.info(f"Fichier par défaut : {DEFAUT}")
else:
    st.info("Chargez le classeur Excel dans la barre latérale pour démarrer.")
    st.stop()

try:
    d, alertes, synth, plan, ventes = traiter(contenu)
except Exception as exc:
    st.error(f"Impossible de lire le classeur : {exc}")
    st.caption("Vérifiez que les onglets Parametres, Produits, Acquereurs, Ventes, Paiements et Remb_Anticipes sont présents.")
    st.stop()

P = d["params"]
dev = P["devise"]
ref = P["date_ref"]

st.sidebar.markdown(
    f"**Date de référence** : {ref:%d/%m/%Y}  \n"
    f"**Taux appliqué** : {P['taux_base']:.2%}  \n"
    f"**Taux pénalité** : {P['taux_pen']:.2%}"
)
if alertes:
    with st.sidebar.expander(f"⚠️ {len(alertes)} alerte(s)", expanded=True):
        for a in alertes:
            st.write("•", a)
else:
    st.sidebar.success("Aucun contrôle en erreur")

tab_glob, tab_grp, tab_acq, tab_exp = st.tabs(["Portefeuille", "Par groupe", "Par acquéreur", "Export"])

# ---------- Portefeuille ----------
with tab_glob:
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Acquéreurs", len(synth))
    c2.metric("Encours", nombre(synth.capital_restant_du.sum(), dev))
    c3.metric("Impayés + pénalités", nombre(synth.total_du.sum(), dev))
    echu = synth.echu_a_date.sum()
    c4.metric("Taux de recouvrement", f"{(1 - synth.impayes.sum() / echu):.1%}" if echu else "n/a")

    g = synth.groupby("groupe").agg(
        acquereurs=("id", "count"), encours=("capital_restant_du", "sum"),
        impayes=("total_du", "sum"), surface_m2=("surface_m2", "sum"),
    )
    st.subheader("Synthèse par groupe")
    st.dataframe(g.style.format("{:,.0f}"), use_container_width=True)

    cat = ventes.groupby("categorie").agg(produits=("adressage", "count"), prix_net=("prix_net", "sum"))
    st.subheader("Ventes par catégorie")
    st.dataframe(cat.style.format({"prix_net": "{:,.0f}"}), use_container_width=True)

    f = plan[(plan.date_echeance > ref) & (plan.date_echeance <= ref + pd.DateOffset(months=24))]
    m = f.groupby(f.date_echeance.dt.to_period("M").astype(str)).mensualite_due.sum()
    col_a, col_b = st.columns([2, 1])
    col_a.plotly_chart(
        go.Figure(go.Bar(x=m.index, y=m.values)).update_layout(title="Encaissements prévus – 24 prochains mois", height=350),
        use_container_width=True,
    )
    vc = synth.statut.value_counts()
    col_b.plotly_chart(
        go.Figure(go.Pie(labels=vc.index, values=vc.values)).update_layout(title="Statut des acquéreurs", height=350),
        use_container_width=True,
    )

# ---------- Par groupe ----------
with tab_grp:
    if not P["groupes"]:
        st.info("Aucun groupe défini dans l'onglet Parametres.")
    else:
        grp = st.selectbox("Groupe", P["groupes"])
        s = synth[synth.groupe == grp]
        st.write(
            f"**{grp}** : {len(s)} acquéreurs | surface {s.surface_m2.sum():,.0f} m² | "
            f"encours {nombre(s.capital_restant_du.sum(), dev)} | impayés {nombre(s.total_du.sum(), dev)}"
        )
        cols = ["id", "nom", "surface_m2", "prix_bien", "mensualite", "capital_restant_du", "impayes", "penalites", "statut"]
        st.dataframe(
            s[cols].style.format({"surface_m2": "{:,.1f}", "prix_bien": "{:,.0f}", "mensualite": "{:,.0f}",
                                  "capital_restant_du": "{:,.0f}", "impayes": "{:,.0f}", "penalites": "{:,.0f}"}),
            use_container_width=True,
        )
        fg = plan[plan.id.isin(s.id) & (plan.date_echeance > ref) & (plan.date_echeance <= ref + pd.DateOffset(months=24))]
        mg = fg.groupby(fg.date_echeance.dt.to_period("M").astype(str)).mensualite_due.sum()
        st.plotly_chart(
            go.Figure(go.Bar(x=mg.index, y=mg.values)).update_layout(title="Encaissements prévus – 24 prochains mois", height=330),
            use_container_width=True,
        )

# ---------- Par acquéreur ----------
with tab_acq:
    libelles = {r.id: f"{r.id} – {r.nom}" for r in synth.itertuples()}
    id_ = st.selectbox("Acquéreur", list(libelles), format_func=lambda k: libelles[k])
    s = synth[synth.id == id_].iloc[0]
    e = plan[(plan.id == id_) & (plan.mensualite_due > 0)].copy()

    st.markdown(f"### {s.nom} ({s.id}) – groupe {s.groupe} – statut : **{s.statut}**")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Prix du bien", nombre(s.prix_bien, dev))
    c2.metric("Capital financé", nombre(s.capital_finance, dev))
    c3.metric("Mensualité", nombre(s.mensualite, dev))
    c4.metric("Capital restant dû", nombre(s.capital_restant_du, dev))
    c5, c6, c7, c8 = st.columns(4)
    c5.metric("Versement initial", nombre(s.versement_initial, dev))
    c6.metric("Total payé", nombre(s.total_paye, dev))
    c7.metric("Impayés", nombre(s.impayes, dev))
    c8.metric("Pénalités", nombre(s.penalites, dev))
    fin = f"{s.date_fin:%m/%Y}" if pd.notna(s.date_fin) else "n/a"
    st.caption(f"Taux {s.taux_annuel:.2%} | {s.nb_echeances} échéances | fin prévue {fin}")
    st.caption("Produits : " + ", ".join(ventes.loc[ventes.id == id_, "adressage"].astype(str)))

    fig = go.Figure()
    fig.add_bar(x=e.date_echeance, y=e.capital_amorti, name="Capital amorti")
    fig.add_bar(x=e.date_echeance, y=e.interets, name="Intérêts")
    fig.add_scatter(x=e.date_echeance, y=e.capital_fin, name="Capital restant", yaxis="y2")
    fig.update_layout(barmode="stack", height=380, title="Décomposition des mensualités",
                      yaxis2=dict(overlaying="y", side="right"))
    st.plotly_chart(fig, use_container_width=True)

    t = e[["n", "date_echeance", "capital_debut", "interets", "mensualite_due", "capital_amorti",
           "remb_anticipe", "capital_fin", "paye", "impaye", "penalite", "statut"]]
    fmt = {c: "{:,.0f}" for c in t.columns if t[c].dtype.kind == "f"}
    fmt["date_echeance"] = "{:%d/%m/%Y}"
    st.dataframe(t.style.format(fmt), use_container_width=True, hide_index=True)

# ---------- Export ----------
with tab_exp:
    st.write("Télécharger les résultats complets (synthèse, échéanciers, détail des ventes).")
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter", datetime_format="dd/mm/yyyy") as xw:
        synth.to_excel(xw, sheet_name="Synthese_Acquereurs", index=False)
        plan.to_excel(xw, sheet_name="Echeanciers", index=False)
        ventes.to_excel(xw, sheet_name="Ventes_Detail", index=False)
        for ws in xw.sheets.values():
            ws.set_column(0, 20, 18)
    st.download_button(
        "⬇️ Télécharger Resultats_Amortissement.xlsx", data=buf.getvalue(),
        file_name="Resultats_Amortissement.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
