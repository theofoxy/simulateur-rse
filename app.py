import streamlit as st
import pandas as pd
import altair as alt
from fpdf import FPDF
from datetime import datetime

# --- CONFIGURATION DE LA PAGE ---
st.set_page_config(
    page_title="Simulateur RSE - Protection des sols",
    page_icon="⚖️",
    layout="wide"
)

# --- BLOC D'ACCUEIL OFFICIEL ---
st.title("⚖️ Votre simulateur RSE : comparatif écologique et économique des surchausses sur un événement")

st.info(
    "👋 **Bienvenue sur votre simulateur d'arbitrage RSE & logistique !**\n\n"
    "Cet outil vous permet de mesurer avec précision l'impact économique et environnemental de la protection des sols sur vos chantiers et événements. "
    "**Ajustez les paramètres de votre événement dans le volet latéral à gauche** (durée, effectif, rotations quotidiennes et pertes prévisionnelles) "
    "pour comparer instantanément le modèle jetable face aux modèles réutilisables en boucle fermée."
)

# --- 1. PARAMÈTRES (BARRE LATÉRALE) ---
st.sidebar.header("🎯 Paramètres de l'Événement")

jours = st.sidebar.number_input(
    "Durée du chantier (jours)",
    min_value=1,
    max_value=60,
    value=7,
    help="Nombre de journées consécutives où le sol doit être protégé."
)

personnes = st.sidebar.number_input(
    "Effectif moyen par jour",
    min_value=10,
    max_value=2000,
    value=100,
    step=10,
    help="Nombre de personnes présentes simultanément sur le site."
)

rotations_jour = st.sidebar.number_input(
    "Paires jetables consommées / pers. / jour",
    min_value=1,
    max_value=12,
    value=4,
    help="Fréquence de remplacement des housses en plastique (déchirures, pauses, sorties de zone propre)."
)

perte_pct_input = st.sidebar.slider(
    "Taux de perte estimé du réutilisable (%)",
    min_value=0,
    max_value=100,
    value=10,
    step=1,
    help="Part du stock textile non restituée (<5% : 12 €/p, entre 5% et 10% : 18 €/p, >10% : 26 €/p)."
)
perte_pct = perte_pct_input / 100.0

usure_pct_input = st.sidebar.slider(
    "Taux d'usure des paires perdues (%)",
    min_value=0,
    max_value=100,
    value=0,
    step=1,
    help="Degré d'usure avant égarement (la précaution ADEME recommande 0 % : paire neuve perdue dès le début)."
)
usure_pct = usure_pct_input / 100.0

reassorts = st.sidebar.number_input(
    "Passages de coursier / réassort (facultatif)",
    min_value=0,
    max_value=20,
    value=0,
    step=1,
    help="Nombre de passages d'un coursier pour échanger le linge sale contre du propre en cours d'événement (hors livraison/reprise initiale incluse à 100 €)."
)

# --- 2. CONSTANTES MÉTIER & ADEME ---
PRIX_JETABLE = 0.19        # Coût unitaire d'une housse plastique (€ HT)
PRIX_NEUF_PAIRE = 26.00    # Valeur de référence d'une paire textile neuve hors barème perte (€ HT)
LIVRAISON_FIXE = 100.00    # Forfait transport aller-retour standard (€ HT)

POIDS_JETABLE = 0.009      # 9 g par paire jetable
FE_PLASTIQUE_FAB = 6.90    # kg CO2 / kg PE vierge (ADEME)
FE_INCINERATION = 2.27     # kg CO2 / kg incinéré (ADEME)
EAU_PLASTIQUE_L_KG = 50.0  # Litres d'eau / kg plastique vierge (ADEME)

POIDS_TEXTILE = 0.090      # 90 g par paire textile tissée
FE_PA_FAB = 13.20          # kg CO2 / kg Polyamide (ADEME)
FE_RECYCLAGE_PA = 0.401    # kg CO2 / kg recyclé (ADEME)
EAU_FAB_PA_L_KG = 150.0    # Litres d'eau / kg fabrication textile
EAU_LAVAGE_L_KG = 15.0     # Litres d'eau / kg textile lavé en blanchisserie
FE_EAU_M3 = 0.132          # kg CO2 / m3 d'eau traitée

# --- MATRICE TARIFAIRE OFFICIELLE EXCEL ---
def get_forfait_loc(stock, nb_jours):
    paliers_stock = [50, 100, 150, 200, 250, 300, 400, 500]
    seuils_jours = [1, 4, 10, 17, 24]
    
    grille = {
        50:  [150.0, 195.0, 240.0, 285.0, 330.0],
        100: [280.0, 364.0, 448.0, 532.0, 616.0],
        150: [405.0, 526.5, 648.0, 769.5, 891.0],
        200: [520.0, 676.0, 832.0, 988.0, 1144.0],
        250: [625.0, 812.5, 1000.0, 1187.5, 1375.0],
        300: [720.0, 936.0, 1152.0, 1368.0, 1584.0],
        400: [880.0, 1144.0, 1408.0, 1672.0, 1936.0],
        500: [1050.0, 1365.0, 1680.0, 1995.0, 2310.0]
    }
    
    # Matching jours (MATCH avec 1 dans Excel)
    j_idx = 0
    for idx, s in enumerate(seuils_jours):
        if nb_jours >= s:
            j_idx = idx
            
    # Matching jauge stock
    s_matched = 50
    for s in paliers_stock:
        if stock >= s:
            s_matched = s
            
    base = grille[s_matched][j_idx]
    if stock > 500:
        base = grille[500][j_idx] * (stock / 500.0)
    elif stock < 50:
        base = grille[50][j_idx]
        
    return base

def get_frais_logistiques(stock, nb_reassorts):
    if stock <= 100:
        prix_reassort = 150.0
    elif stock <= 200:
        prix_reassort = 200.0
    elif stock <= 300:
        prix_reassort = 250.0
    else:
        prix_reassort = 350.0
    return LIVRAISON_FIXE + (nb_reassorts * prix_reassort)

def get_tarif_unitaire_perte(taux_perte_ratio):
    if taux_perte_ratio < 0.05:
        return 12.00
    elif taux_perte_ratio <= 0.10:
        return 18.00
    else:
        return 26.00

def get_seuil_max_pertes(budget_max_penalite, stock_total):
    if budget_max_penalite <= 0 or stock_total <= 0:
        return 0, 0.0
    max_paires = 0
    max_pct = 0.0
    for p in range(1, 101):
        p_ratio = p / 100.0
        nb_p = round(stock_total * p_ratio)
        tarif = get_tarif_unitaire_perte(p_ratio)
        cout = nb_p * tarif
        if cout <= budget_max_penalite:
            max_paires = nb_p
            max_pct = p
        else:
            break
    return max_paires, float(max_pct)

# --- 3. MOTEUR DE CALCULS ---
vol_jetable = int(jours * personnes * rotations_jour)
stock_circulaire = int(personnes)
paires_perdues = int(round(stock_circulaire * perte_pct))
paires_rendues = int(stock_circulaire - paires_perdues)

# Volet financier
budget_jetable = vol_jetable * PRIX_JETABLE
tarif_unitaire_perte = get_tarif_unitaire_perte(perte_pct)
amende_pertes = paires_perdues * tarif_unitaire_perte
forfait_loc = get_forfait_loc(stock_circulaire, jours)
frais_log = get_frais_logistiques(stock_circulaire, reassorts)
budget_location = forfait_loc + frais_log + amende_pertes
cout_usage_achat = ((stock_circulaire * PRIX_NEUF_PAIRE) / 100.0) * jours
budget_achat = cout_usage_achat + amende_pertes

# Volet déchets plastiques
dechets_plastique_kg = vol_jetable * POIDS_JETABLE

# Volet eau (Méthode officielle Excel)
eau_jetable_litres = vol_jetable * POIDS_JETABLE * EAU_PLASTIQUE_L_KG
eau_perdue_unit = POIDS_TEXTILE * EAU_FAB_PA_L_KG * (1.0 - usure_pct)
eau_rendue_unit = (POIDS_TEXTILE * EAU_FAB_PA_L_KG / 100.0 * jours) + (POIDS_TEXTILE * EAU_LAVAGE_L_KG * jours)
eau_reutilisable_litres = (paires_perdues * eau_perdue_unit) + (paires_rendues * eau_rendue_unit)

# Volet carbone (Méthode officielle Excel)
co2_jetable = vol_jetable * POIDS_JETABLE * (FE_PLASTIQUE_FAB + FE_INCINERATION)
co2_rendue_unit = (
    (POIDS_TEXTILE * FE_PA_FAB / 100.0 * jours) +
    (POIDS_TEXTILE * jours * EAU_LAVAGE_L_KG / 1000.0 * FE_EAU_M3) +
    (POIDS_TEXTILE / 100.0 * jours * FE_RECYCLAGE_PA)
)
co2_perdue_unit = POIDS_TEXTILE * ((FE_PA_FAB * (1.0 - usure_pct)) + FE_INCINERATION)
co2_circulaire_total = (paires_perdues * co2_perdue_unit) + (paires_rendues * co2_rendue_unit)

def afficher_recap_parametres():
    with st.expander("📋 Paramètres de l'événement associés (pour export / capture d'écran)"):
        r1, r2, r3, r4 = st.columns(4)
        r1.write(f"⏱️ **Durée :** {jours} jours")
        r2.write(f"👥 **Effectif :** {personnes} pers./jour")
        r3.write(f"🔄 **Rotations :** {rotations_jour} paires/j")
        r4.write(f"⚠️ **Pertes estimées :** {perte_pct_input} % (Tarif : {tarif_unitaire_perte:.2f} €/p)")

# ==============================================================================
# SECTION A : FLUX PHYSIQUES & VOLUMÉTRIE
# ==============================================================================
st.markdown("### 📦 Volumétrie et Flux Physiques sur l'Événement")
f1, f2, f3, f4 = st.columns(4)
f1.metric("Paires JETABLES consommées", f"{vol_jetable:,} paires".replace(",", " "))
f2.metric("Stock RÉUTILISABLE mobilisé", f"{stock_circulaire:,} paires".replace(",", " "))
f3.metric("Paires perdues / non restituées", f"{paires_perdues:,} paires".replace(",", " "))
f4.metric("Paires restituées & relavées", f"{paires_rendues:,} paires".replace(",", " "))

st.divider()

# ==============================================================================
# SECTION B : SYNTHÈSE BUDGÉTAIRE
# ==============================================================================
st.subheader("💰 Synthèse Budgétaire (TCO - Coût Total de Possession)")

b1, b2, b3 = st.columns(3)
b1.metric("Option Jetable", f"{budget_jetable:,.0f} €".replace(",", " "))
b2.metric(
    "Option Location",
    f"{budget_location:,.0f} €".replace(",", " "),
    delta=f"{budget_jetable - budget_location:,.0f} € vs Jetable"
)
b3.metric(
    "Achat d'un stock réutilisable",
    f"{budget_achat:,.0f} €".replace(",", " "),
    delta=f"{budget_jetable - budget_achat:,.0f} € vs Jetable"
)

if paires_perdues == 0:
    st.success("💡 **Pertes nulles (0 %) :** Aucune pénalité appliquée. Le modèle réutilisable est exploité à son efficience économique maximale.")
elif perte_pct < 0.05:
    st.info(f"📊 **Taux de perte < 5 % ({perte_pct_input} %) :** Chaque paire égarée est facturée **12,00 € HT**. Facturation totale des pertes : {paires_perdues} paires × 12,00 € = **{amende_pertes:,.2f} € HT**.")
elif perte_pct <= 0.10:
    st.warning(f"⚠️ **Taux de perte entre 5 % et 10 % ({perte_pct_input} %) :** Le tarif monte à **18,00 € HT** pour chaque paire égarée. Facturation totale des pertes : {paires_perdues} paires × 18,00 € = **{amende_pertes:,.2f} € HT**.")
else:
    st.error(f"🚨 **Taux de perte > 10 % ({perte_pct_input} %) :** Le tarif maximal s'applique : **26,00 € HT** pour chaque paire égarée. Facturation totale des pertes : {paires_perdues} paires × 26,00 € = **{amende_pertes:,.2f} € HT**.")

with st.expander("🔍 Méthodologie détaillée : comment sont calculés les budgets et les pertes ?"):
    st.markdown(f"""
    * **Règle de tarification des pertes textiles :** Le coût unitaire appliqué à l'ensemble des paires égarées dépend du taux global :
      * **< 5 % :** **12,00 € HT** par paire.
      * **5 % à 10 % :** **18,00 € HT** par paire.
      * **> 10 % :** **26,00 € HT** par paire.
    * **Scénario Jetable :** Achat direct du volume consommé ({vol_jetable:,} paires × {PRIX_JETABLE:.2f} € HT = **{budget_jetable:,.2f} € HT**).
    * **Scénario Location :** Forfait selon la grille officielle ({forfait_loc:,.0f} €) + transport initial et reprises ({frais_log:,.0f} €) + pénalités ({amende_pertes:,.2f} € HT).
    * **Scénario Achat d'un stock réutilisable :** Amortissement d'usage sur 100 cycles ({cout_usage_achat:,.2f} €) + rachat des {paires_perdues} paires perdues au tarif de {tarif_unitaire_perte:.2f} € HT (**{amende_pertes:,.2f} € HT**).
    """)

df_budget = pd.DataFrame({
    "Option": ["Jetable", "Location", "Achat d'un stock réutilisable"],
    "Coût (€)": [round(budget_jetable), round(budget_location), round(budget_achat)]
})

chart_b = alt.Chart(df_budget).mark_bar(size=60).encode(
    x=alt.X(
        "Option:N",
        title=None,
        sort=["Jetable", "Location", "Achat d'un stock réutilisable"],
        axis=alt.Axis(labelAngle=0, labelFontSize=12)
    ),
    y=alt.Y("Coût (€):Q", title="Budget total (€ HT)"),
    color=alt.Color(
        "Option:N",
        scale=alt.Scale(
            domain=["Jetable", "Location", "Achat d'un stock réutilisable"],
            range=["#E74C3C", "#3498DB", "#2ECC71"]
        ),
        legend=None
    )
).properties(height=280)

text_b = chart_b.mark_text(dy=-10, fontSize=13, fontWeight="bold").encode(
    text=alt.Text("Coût (€):Q", format=",.0f")
)
st.altair_chart((chart_b + text_b), use_container_width=True)
afficher_recap_parametres()

st.divider()

# ==============================================================================
# SECTION C : DÉCHETS PLASTIQUES GÉNÉRÉS
# ==============================================================================
st.subheader("🗑️ Déchets Plastiques Générés")

with st.expander("🔍 Méthodologie détaillée : comment est calculée la masse de déchets ?"):
    st.markdown(f"""
    * Chaque paire plastique à usage unique pèse **{POIDS_JETABLE * 1000:.0f} grammes**.
    * Le jetable engendre **{vol_jetable:,} paires jetées**, soit **{dechets_plastique_kg:,.1f} kg de plastique**.
    * Les options réutilisables génèrent **0 kg de déchet à usage unique**.
    """)

d1, d2 = st.columns(2)
d1.metric("Déchets plastiques : Jetable", f"{dechets_plastique_kg:,.1f} kg".replace(",", " "))
d2.metric("Déchets plastiques : Réutilisable", "0 kg", delta=f"-{dechets_plastique_kg:,.1f} kg évités")

st.divider()

# ==============================================================================
# SECTION D : BILAN CARBONE (ACV)
# ==============================================================================
st.subheader("🌿 Bilan Carbone (ACV - Analyse de Cycle de Vie)")

with st.expander("🔍 Méthodologie détaillée : comment sont calculées les émissions CO2 ?"):
    st.markdown(f"""
    Les facteurs d'émissions sont issus des bases officielles de l'**ADEME** :
    * **Jetable ({co2_jetable:,.1f} kg CO2) :** Polyéthylène ({FE_PLASTIQUE_FAB} kg CO2/kg) + incinération ({FE_INCINERATION} kg CO2/kg).
    * **Paire réutilisable restituée ({co2_rendue_unit:.3f} kg CO2 / paire) :** Fabrication amortie sur 100 cycles + lavages quotidiens + recyclage ({FE_RECYCLAGE_PA} kg CO2/kg).
    * **Paire égarée ({co2_perdue_unit:.3f} kg CO2 / paire) :** Selon l'usure choisie ({usure_pct_input} %), sa fabrication est imputée à l'événement et complétée par son incinération.
    """)

c1, c2 = st.columns(2)
reduction_co2 = co2_jetable - co2_circulaire_total
c1.metric("Émissions : Jetable", f"{co2_jetable:,.1f} kg CO2".replace(",", " "))
c2.metric(
    "Émissions : Réutilisable",
    f"{co2_circulaire_total:,.1f} kg CO2".replace(",", " "),
    delta=f"-{reduction_co2:,.1f} kg CO2 évités" if reduction_co2 >= 0 else f"+{abs(reduction_co2):,.1f} kg CO2 d'excès"
)

df_co2 = pd.DataFrame({
    "Option": ["Jetable", "Réutilisable (Loc / Achat)"],
    "Émissions (kg CO2)": [round(co2_jetable, 1), round(co2_circulaire_total, 1)]
})

chart_c = alt.Chart(df_co2).mark_bar(size=60).encode(
    x=alt.X(
        "Option:N",
        title=None,
        sort=["Jetable", "Réutilisable (Loc / Achat)"],
        axis=alt.Axis(labelAngle=0, labelFontSize=12)
    ),
    y=alt.Y("Émissions (kg CO2):Q", title="Émissions (kg CO2-eq)"),
    color=alt.Color(
        "Option:N",
        scale=alt.Scale(domain=["Jetable", "Réutilisable (Loc / Achat)"], range=["#E74C3C", "#2ECC71"]),
        legend=None
    )
).properties(height=280)

text_c = chart_c.mark_text(dy=-10, fontSize=13, fontWeight="bold").encode(
    text=alt.Text("Émissions (kg CO2):Q", format=",.1f")
)
st.altair_chart((chart_c + text_c), use_container_width=True)
afficher_recap_parametres()

st.divider()

# ==============================================================================
# SECTION E : CONSOMMATION D'EAU & LAVAGE
# ==============================================================================
st.subheader("💧 Consommation d'Eau & Cycles de Lavage")

with st.expander("🔍 Méthodologie détaillée : comment est mesuré l'impact sur la ressource en eau ?"):
    st.markdown(f"""
    * **Réutilisable :** Lavage en blanchisserie industrielle standard ({EAU_LAVAGE_L_KG} L/kg) + part de fabrication amortie.
    * **Jetable :** Raffinage du polyéthylène ({EAU_PLASTIQUE_L_KG} L/kg de plastique).
    """)

e1, e2 = st.columns(2)
diff_eau = eau_jetable_litres - eau_reutilisable_litres
e1.metric("Eau consommée : Jetable", f"{eau_jetable_litres:,.0f} L".replace(",", " "))
e2.metric(
    "Eau consommée : Réutilisable",
    f"{eau_reutilisable_litres:,.0f} L".replace(",", " "),
    delta=f"{diff_eau:,.0f} L vs Jetable"
)

df_eau = pd.DataFrame({
    "Option": ["Jetable (Pétrochimie)", "Réutilisable (Lavage & Fab)"],
    "Consommation (Litres)": [round(eau_jetable_litres), round(eau_reutilisable_litres)]
})

chart_e = alt.Chart(df_eau).mark_bar(size=60).encode(
    x=alt.X(
        "Option:N",
        title=None,
        sort=["Jetable (Pétrochimie)", "Réutilisable (Lavage & Fab)"],
        axis=alt.Axis(labelAngle=0, labelFontSize=12)
    ),
    y=alt.Y("Consommation (Litres):Q", title="Volume d'eau (Litres)"),
    color=alt.Color(
        "Option:N",
        scale=alt.Scale(domain=["Jetable (Pétrochimie)", "Réutilisable (Lavage & Fab)"], range=["#E67E22", "#3498DB"]),
        legend=None
    )
).properties(height=280)

text_e = chart_e.mark_text(dy=-10, fontSize=13, fontWeight="bold").encode(
    text=alt.Text("Consommation (Litres):Q", format=",.0f")
)
st.altair_chart((chart_e + text_e), use_container_width=True)
afficher_recap_parametres()

st.divider()

# ==============================================================================
# SECTION F : BREAK-EVEN & SENSIBILITÉ AUX PERTES
# ==============================================================================
st.subheader("🎯 Seuils de Rupture & Sensibilité aux Pertes (Break-Even)")

pertes_gamme = list(range(0, 101, 1))
data_courbes = []

for p in pertes_gamme:
    p_ratio = p / 100.0
    nb_p = round(stock_circulaire * p_ratio)
    tarif_p = get_tarif_unitaire_perte(p_ratio)
    penalite = nb_p * tarif_p
    cout_loc = forfait_loc + frais_log + penalite
    cout_ach = cout_usage_achat + penalite
    
    data_courbes.append({"Taux de perte (%)": p, "Budget (€ HT)": budget_jetable, "Solution": "Jetable (Fixe)"})
    data_courbes.append({"Taux de perte (%)": p, "Budget (€ HT)": cout_loc, "Solution": "Location"})
    data_courbes.append({"Taux de perte (%)": p, "Budget (€ HT)": cout_ach, "Solution": "Achat d'un stock réutilisable"})

df_sensib = pd.DataFrame(data_courbes)

courbe_chart = alt.Chart(df_sensib).mark_line(strokeWidth=3).encode(
    x=alt.X("Taux de perte (%):Q", title="Taux de perte textile (%)", scale=alt.Scale(domain=[0, 100])),
    y=alt.Y("Budget (€ HT):Q", title="Budget global (€ HT)"),
    color=alt.Color(
        "Solution:N",
        scale=alt.Scale(
            domain=["Jetable (Fixe)", "Location", "Achat d'un stock réutilisable"],
            range=["#E74C3C", "#3498DB", "#2ECC71"]
        ),
        legend=alt.Legend(title="Scénarios", orient="top")
    )
).properties(height=360)

point_actuel = pd.DataFrame([{"x": perte_pct_input}])
regle_actuelle = alt.Chart(point_actuel).mark_rule(
    color="#7F8C8D",
    strokeDash=[4, 4],
    size=2
).encode(x="x:Q")

st.altair_chart(courbe_chart + regle_actuelle, use_container_width=True)

marge_loc_brute = budget_jetable - (forfait_loc + frais_log)
pertes_max_loc, seuil_loc_pct = get_seuil_max_pertes(marge_loc_brute, stock_circulaire)
marge_achat_brute = budget_jetable - cout_usage_achat
pertes_max_achat, seuil_achat_pct = get_seuil_max_pertes(marge_achat_brute, stock_circulaire)

denom_eco = co2_perdue_unit - co2_rendue_unit
bep_eco_pct = ((co2_jetable / stock_circulaire) - co2_rendue_unit) / denom_eco if denom_eco > 0 else 1.0

k1, k2 = st.columns(2)
with k1:
    st.markdown("#### 💶 Arbitrage Économique")
    if marge_loc_brute < 0:
        surcout_base = (forfait_loc + frais_log) - budget_jetable
        st.error(
            f"🚨 **Location structurellement non rentable :** Pour ce format d'événement ({jours} jours, {personnes} personnes), "
            f"le forfait de mise à disposition ({forfait_loc:,.0f} €) et la logistique ({frais_log:,.0f} €) rendent la location plus chère que le plastique **dès le départ, même avec 0 % de perte** "
            f"(surcoût initial de {surcout_base:,.0f} € HT)."
        )
        if marge_achat_brute > 0:
            st.info(f"💡 **Alternative économique :** L'**Achat d'un stock réutilisable** reste rentable jusqu'à **{pertes_max_achat} paires égarées** ({seuil_achat_pct:.0f} % du stock).")
    else:
        if paires_perdues > pertes_max_loc:
            st.error(
                f"🚨 **Déficit budgétaire :** Avec {paires_perdues} paires perdues ({amende_pertes:,.0f} € facturés à {tarif_unitaire_perte:.0f} €/p), "
                f"la location est plus chère que le jetable de {budget_location - budget_jetable:,.0f} €. "
                f"Le seuil de rentabilité se situe à **{seuil_loc_pct:.0f} %** (soit max **{pertes_max_loc} paires**)."
            )
        else:
            st.success(
                f"✅ **Rentabilité préservée :** Vos {paires_perdues} paires perdues restent sous la limite de **{pertes_max_loc} paires** "
                f"({seuil_loc_pct:.0f} % du parc). Gain net : **{budget_jetable - budget_location:,.0f} €** par rapport au jetable."
            )

with k2:
    st.markdown("#### 🌿 Arbitrage Écologique")
    if bep_eco_pct >= 1.0:
        st.success("✅ **Bénéfice écologique inconditionnel :** Même en cas de perte de 100 % du parc neuf, le volume massif de déchets plastiques évités rend le textile toujours moins émetteur en CO2.")
    elif bep_eco_pct <= 0.0:
        st.error("🚨 **Volume plastique insuffisant :** Sur un chantier très court ou avec un renouvellement minime, le cycle de vie textile émet plus de CO2 que le peu de plastique consommé.")
    else:
        max_paires_eco = int(bep_eco_pct * stock_circulaire)
        if paires_perdues > max_paires_eco:
            st.error(
                f"🚨 **Déficit carbone :** Perdre {paires_perdues} paires dépasse la limite écologique de **{max_paires_eco} paires** "
                f"({bep_eco_pct * 100:.1f} %). La refabrication du textile annule l'intérêt environnemental."
            )
        else:
            st.success(
                f"✅ **Bénéfice carbone maîtrisé :** Le réutilisable reste écologiquement gagnant tant que vous ne dépassez pas "
                f"**{bep_eco_pct * 100:.1f} % de perte** (soit **{max_paires_eco} paires maximum**)."
            )

st.divider()

# ==============================================================================
# SECTION G : RAPPORT CLIENT PDF
# ==============================================================================
st.subheader("📄 Synthèse Client & Export PDF")

class PDFExecutive(FPDF):
    def header(self):
        self.set_fill_color(31, 78, 121)
        self.rect(0, 0, 210, 24, 'F')
        self.set_font("helvetica", "B", 13)
        self.set_text_color(255, 255, 255)
        self.cell(0, 14, "RAPPORT D'ARBITRAGE RSE - SURCHAUSSURES EVENEMENTIELLES", border=0, align="C", new_x="LMARGIN", new_y="NEXT")
        self.set_font("helvetica", "", 9)
        self.cell(0, 5, f"Synthese executive generee le {datetime.now().strftime('%d/%m/%Y')}", border=0, align="C", new_x="LMARGIN", new_y="NEXT")
        self.ln(10)

    def footer(self):
        self.set_y(-15)
        self.set_font("helvetica", "I", 8)
        self.set_text_color(130, 130, 130)
        self.cell(0, 10, "Document d'aide a la decision RSE & Logistique - Confidentialite client", border=0, align="C")

def generer_pdf_client():
    pdf = PDFExecutive(orientation="P", unit="mm", format="A4")
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    
    pdf.set_font("helvetica", "B", 11)
    pdf.set_text_color(31, 78, 121)
    pdf.cell(0, 8, "1. PARAMETRES DE L'EVENEMENT", border="B", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(2)
    
    pdf.set_font("helvetica", "", 9)
    pdf.set_text_color(40, 40, 40)
    pdf.cell(95, 6, f"- Duree du chantier : {jours} jours", new_x="RIGHT")
    pdf.cell(95, 6, f"- Effectif equipe : {personnes} personnes / jour", new_x="LMARGIN", new_y="NEXT")
    pdf.cell(95, 6, f"- Rotations jetables : {rotations_jour} paires / pers. / jour", new_x="RIGHT")
    pdf.cell(95, 6, f"- Pertes prevues : {perte_pct_input} % ({paires_perdues} p. a {tarif_unitaire_perte:.2f} EUR/p)", new_x="LMARGIN", new_y="NEXT")
    pdf.cell(95, 6, f"- Usure paires perdues : {usure_pct_input} %", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(4)
    
    pdf.set_font("helvetica", "B", 11)
    pdf.set_text_color(31, 78, 121)
    pdf.cell(0, 8, "2. FLUX PHYSIQUES & INVENTAIRE", border="B", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(2)
    
    pdf.set_font("helvetica", "", 9)
    pdf.set_text_color(40, 40, 40)
    pdf.cell(95, 6, f"* Housses jetables consommees : {vol_jetable:,} paires".replace(",", " "), new_x="RIGHT")
    pdf.cell(95, 6, f"* Stock reutilisable mobilise : {stock_circulaire:,} paires".replace(",", " "), new_x="LMARGIN", new_y="NEXT")
    pdf.cell(95, 6, f"* Paires perdues / degradees : {paires_perdues:,} paires", new_x="RIGHT")
    pdf.cell(95, 6, f"* Paires lavees et remises en stock : {paires_rendues:,} paires", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(4)
    
    pdf.set_font("helvetica", "B", 11)
    pdf.set_text_color(31, 78, 121)
    pdf.cell(0, 8, "3. SYNTHESE BUDGETAIRE GLOBALE (TCO)", border="B", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(2)
    
    pdf.set_fill_color(240, 240, 240)
    pdf.set_font("helvetica", "B", 9)
    pdf.cell(70, 7, "Option", border=1, fill=True, new_x="RIGHT")
    pdf.cell(50, 7, "Cout Global (HT)", border=1, align="C", fill=True, new_x="RIGHT")
    pdf.cell(70, 7, "Ecart vs Jetable", border=1, align="C", fill=True, new_x="LMARGIN", new_y="NEXT")
    
    pdf.set_font("helvetica", "", 9)
    pdf.cell(70, 7, "Jetable (Polyethylene)", border=1, new_x="RIGHT")
    pdf.cell(50, 7, f"{budget_jetable:,.0f} EUR".replace(",", " "), border=1, align="C", new_x="RIGHT")
    pdf.cell(70, 7, "Reference", border=1, align="C", new_x="LMARGIN", new_y="NEXT")
    
    pdf.cell(70, 7, "Location (Circulaire)", border=1, new_x="RIGHT")
    pdf.cell(50, 7, f"{budget_location:,.0f} EUR".replace(",", " "), border=1, align="C", new_x="RIGHT")
    pdf.cell(70, 7, f"{budget_jetable - budget_location:+,.0f} EUR".replace(",", " "), border=1, align="C", new_x="LMARGIN", new_y="NEXT")
    
    pdf.cell(70, 7, "Achat d'un stock reutilisable", border=1, new_x="RIGHT")
    pdf.cell(50, 7, f"{budget_achat:,.0f} EUR".replace(",", " "), border=1, align="C", new_x="RIGHT")
    pdf.cell(70, 7, f"{budget_jetable - budget_achat:+,.0f} EUR".replace(",", " "), border=1, align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(4)
    
    pdf.set_font("helvetica", "B", 11)
    pdf.set_text_color(31, 78, 121)
    pdf.cell(0, 8, "4. BILAN ENVIRONNEMENTAL & RESSOURCES (ADEME)", border="B", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(2)
    
    pdf.set_fill_color(240, 240, 240)
    pdf.set_font("helvetica", "B", 9)
    pdf.cell(55, 7, "Indicateur", border=1, fill=True, new_x="RIGHT")
    pdf.cell(45, 7, "Jetable", border=1, align="C", fill=True, new_x="RIGHT")
    pdf.cell(45, 7, "Reutilisable", border=1, align="C", fill=True, new_x="RIGHT")
    pdf.cell(45, 7, "Bilan Net", border=1, align="C", fill=True, new_x="LMARGIN", new_y="NEXT")
    
    pdf.set_font("helvetica", "", 9)
    pdf.cell(55, 7, "Dechets plastiques", border=1, new_x="RIGHT")
    pdf.cell(45, 7, f"{dechets_plastique_kg:,.1f} kg", border=1, align="C", new_x="RIGHT")
    pdf.cell(45, 7, "0.0 kg", border=1, align="C", new_x="RIGHT")
    pdf.cell(45, 7, f"-{dechets_plastique_kg:,.1f} kg evites", border=1, align="C", new_x="LMARGIN", new_y="NEXT")
    
    pdf.cell(55, 7, "Emissions CO2-eq", border=1, new_x="RIGHT")
    pdf.cell(45, 7, f"{co2_jetable:,.1f} kg", border=1, align="C", new_x="RIGHT")
    pdf.cell(45, 7, f"{co2_circulaire_total:,.1f} kg", border=1, align="C", new_x="RIGHT")
    pdf.cell(45, 7, f"-{reduction_co2:,.1f} kg evites", border=1, align="C", new_x="LMARGIN", new_y="NEXT")
    
    pdf.cell(55, 7, "Consommation d'eau", border=1, new_x="RIGHT")
    pdf.cell(45, 7, f"{eau_jetable_litres:,.0f} L (Petro)", border=1, align="C", new_x="RIGHT")
    pdf.cell(45, 7, f"{eau_reutilisable_litres:,.0f} L (Lavage)", border=1, align="C", new_x="RIGHT")
    pdf.cell(45, 7, f"{diff_eau:+,.0f} L", border=1, align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(4)
    
    pdf.set_font("helvetica", "B", 11)
    pdf.set_text_color(31, 78, 121)
    pdf.cell(0, 8, "5. CONCLUSION DECISIONNELLE (BREAK-EVEN)", border="B", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(2)
    
    pdf.set_font("helvetica", "", 9)
    pdf.set_text_color(40, 40, 40)
    if marge_loc_brute < 0:
        verdict = f"La location est defavorable face au jetable (+{(forfait_loc + frais_log) - budget_jetable:,.0f} EUR des 0% de perte). Privilegier l'achat amorti ou une mutualisation."
    else:
        verdict = f"La location est recommandee et rentable jusqu'a {pertes_max_loc} paires egarees ({seuil_loc_pct:.0f}% de perte autorisee)."
    
    pdf.multi_cell(190, 5, verdict)
    return bytes(pdf.output())

col_pdf1, col_pdf2 = st.columns([2, 1])
with col_pdf1:
    st.write("Téléchargez un rapport au format PDF haute définition, prêt à être transmis à votre direction des achats ou à vos clients :")
with col_pdf2:
    pdf_bytes = generer_pdf_client()
    st.download_button(
        label="📥 Télécharger le Rapport Client (PDF)",
        data=pdf_bytes,
        file_name=f"Rapport_RSE_Surchausses_{jours}j_{personnes}p.pdf",
        mime="application/pdf"
    )
