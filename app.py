import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
import requests
from datetime import datetime
from io import BytesIO
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.pdfgen import canvas
import json

# Configuration de la page
st.set_page_config(
    layout="wide",
    page_title="Rapport RSE CEE P5",
    page_icon="🌱",
    initial_sidebar_state="expanded"
)

# Style CSS personnalisé
st.markdown("""
    <style>
    .big-font { font-size:20px !important; font-weight: bold; color: #1e3d59; }
    .medium-font { font-size:16px !important; color: #2e5266; }
    .stMetric { background-color: #f8f9fa; padding: 15px; border-radius: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
    </style>
""", unsafe_allow_html=True)

# Palette
COLOR_ENERGY = "#2a9d8f"  # économies d'énergie (GWh)
COLOR_ENERGY_ACCENT = "#1f776b"
COLOR_CO2 = "#e76f51"  # CO2 évité
COLOR_CO2_ACCENT = "#c4563d"
COLOR_CEE = "#457b9d"  # GWhc
COLOR_ECONOMY = "#f4a261"  # Couleur pour l'économie
COLOR_ECONOMY_ACCENT = "#e5934a"
COLOR_PRIME = "#457b9d"
COLOR_PRIME_ACCENT = "#3a698a"

# =========================
# CONSTANTES & HYPOTHÈSES
# =========================
FACTEUR_CUMAC_TO_KWH = {
    'BAR-TH': 1 / 12.16,
    'BAR-EN': 1 / 17.29,
    'BAR-EQ': 1 / 11.12,
    'BAT-TH': 1 / 12.16,
    'AGRI-TH': 1 / 12.16,
    'BAT-EN': 1 / 17.29,
    'TRA': 1 / 0.9615,  # Modifié pour une durée de vie de 1 an
    'DEFAULT': 1 / 8.11
}

DUREE_VIE_EQUIPEMENT = {
    'BAR-TH': 17,
    'AGRI-TH': 17,
    'BAR-EN': 30,
    'BAR-EQ': 15,
    'BAT-TH': 17,
    'BAT-EN': 30,
    'TRA': 1,  # Modifié de 7 à 1 an
    'DEFAULT': 10
}

# Hypothèses pour la France
EMISSION_CO2_KWH = 0.057
CO2_PAR_VOITURE_AN = 2.8
CO2_PAR_KM_VOITURE = 0.12

# Nouveaux coûts simplifiés
COUT_ELECTRICITE_KWH = 0.22  # Gardé pour référence si besoin
COUT_CHAUFFAGE_KWH = 0.10

CONSO_MOYENNE_FOYER_KWH = 15312  # kWh/an (FR  chauffage élec +elec)
CIRCONFERENCE_TERRE_KM = 40075
TAUX_ACTUALISATION = 0.04
TAUX_EFFICACITE_DEFAULT = 0.45

VILLES_REFERENCE = {
    10000: "Luxeuil-les-Bains (10k hab)",
    25000: "Saintes (25k hab)",
    32175: "Aix-les-Bains-Rhône (32k hab)",
    50000: "Niort (50k hab)",
    100000: "Nancy (100k hab)",
    250000: "Montpellier (250k hab)",
    500000: "Lyon (500k hab)",
    1000000: "Marseille (1M hab)",
    2000000: "Paris (2.2M hab)"
}


# =========================
# UTILITAIRES
# =========================
def get_ville_equivalente(nb_habitants):
    for seuil, ville in sorted(VILLES_REFERENCE.items()):
        if nb_habitants <= seuil:
            return ville
    return VILLES_REFERENCE[2000000]


def format_number(num, decimals=0):
    if pd.isna(num):
        return "N/A"
    if decimals == 0:
        return f"{int(round(num)):,}".replace(',', ' ')
    return f"{num:,.{decimals}f}".replace(',', ' ')


@st.cache_data
def load_and_process_data(file, taux_efficacite):
    try:
        df = pd.read_excel(file, engine='openpyxl')
        df.columns = df.columns.str.strip()

        # Dates
        date_cols = ['Date Validation', 'Date depot', 'Date de début', 'Date de fin', 'Date de la facture']
        for col in date_cols:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors='coerce')

        # Code postal -> département
        if 'code postal' in df.columns:
            df['code postal'] = df['code postal'].astype(str).str.strip().str.zfill(5)
            df['departement'] = df['code postal'].str[:2]

        # Période
        if 'PERIODE' in df.columns:
            df['Période'] = df['PERIODE'].astype(str).str.strip().str.upper()
        elif 'Depot' in df.columns:
            df['Période'] = df['Depot'].astype(str).str.extract(r'(P\d)', expand=False).fillna('P5')
        else:
            df['Période'] = 'P5'

        # Mandataire
        if 'Mandataire' in df.columns:
            df['Mandataire'] = df['Mandataire'].astype(str).str.strip()
            df.loc[df['Mandataire'].isin(['nan', 'NaN']), 'Mandataire'] = 'Non renseigné'
        else:
            df['Mandataire'] = 'Non renseigné'

        # Numériques primaires
        for col in ['Total', 'Total précarité', 'Total classique', 'Tableau Recapitulatif champ 23']:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
            else:
                df[col] = 0

        # Mapping Code équipement (MAJUSCULE)
        if 'Code équipement' in df.columns:
            code = df['Code équipement'].astype(str).str.strip().str.upper()
            df['CodeEquip_prefix'] = code.str.split('-').str[0]
            df['CodeEquip_sub'] = code.str.split('-').str[1].fillna('')

            df['FacteurKey'] = np.where(
                df['CodeEquip_sub'].isin(['TH', 'EN', 'EQ']),
                df['CodeEquip_prefix'] + '-' + df['CodeEquip_sub'],
                df['CodeEquip_prefix']  # ex: TRA, IND, AGRI…
            )
            df['Facteur_Conversion'] = df['FacteurKey'].map(FACTEUR_CUMAC_TO_KWH).fillna(
                FACTEUR_CUMAC_TO_KWH['DEFAULT'])

            df['Duree_Vie'] = df['FacteurKey'].map(DUREE_VIE_EQUIPEMENT).fillna(DUREE_VIE_EQUIPEMENT['DEFAULT'])

            df['Secteur'] = df['CodeEquip_prefix'].map({
                'BAR': 'Bât. Résidentiel',
                'BAT': 'Bât. Tertiaire',
                'TRA': 'Transport',
                'AGRI': 'Agriculture',
                'IND': 'Industrie'
            }).fillna('Autre')
            df['Sous_Categorie'] = df['CodeEquip_sub'].replace('', 'N/A')
        else:
            df['Facteur_Conversion'] = FACTEUR_CUMAC_TO_KWH['DEFAULT']
            df['Duree_Vie'] = DUREE_VIE_EQUIPEMENT['DEFAULT']
            df['Secteur'] = 'Autre'
            df['Sous_Categorie'] = 'N/A'
            df['CodeEquip_prefix'] = 'Autre'

        # Type bénéficiaire
        siren_col_8 = 'Tableau Recapitulatif champ 8'
        siren_col_9 = 'Tableau Recapitulatif champ 9'
        if siren_col_8 not in df.columns: df[siren_col_8] = ''
        if siren_col_9 not in df.columns: df[siren_col_9] = ''
        df[siren_col_8] = df[siren_col_8].astype(str).str.strip().replace('nan', '')
        df[siren_col_9] = df[siren_col_9].astype(str).str.strip().replace('nan', '')

        conditions = [
            df['Total précarité'] > 0,
            (df[siren_col_8] != '') | (df[siren_col_9] != '')
        ]
        choices = ['Précarité énergétique', 'Personne Morale']
        df['Type_Beneficiaire'] = np.select(conditions, choices, default='Ménage Classique')

        df['Statut'] = df['Date Validation'].apply(lambda x: 'Validé' if pd.notna(x) else 'En cours')
        df['Annee_Depot'] = df['Date depot'].dt.year

        # Calculs énergie / CO2 / €
        df['kWh_cumac'] = pd.to_numeric(df.get('Total', 0), errors='coerce').fillna(0)
        df['GWh_cumac'] = df['kWh_cumac'] / 1_000_000
        df['kWh_reels_annuels'] = df['kWh_cumac'] * df['Facteur_Conversion'] * taux_efficacite
        df['GWh_reels_annuels'] = df['kWh_reels_annuels'] / 1_000_000

        df['CO2_evite_tonnes_an'] = (df['kWh_reels_annuels'] * EMISSION_CO2_KWH) / 1000
        df['Nb_foyers_equivalents'] = df['kWh_reels_annuels'] / CONSO_MOYENNE_FOYER_KWH

        # Logique de coût ultra-simplifiée : on considère que toute économie impacte le chauffage
        df['Economies_euros_an'] = df['kWh_reels_annuels'] * COUT_CHAUFFAGE_KWH
        df['Prime_versee'] = df['Tableau Recapitulatif champ 23']

        return df

    except Exception as e:
        st.error(f"Erreur lors du chargement du fichier : {str(e)}")
        return None


# =========================
# UI
# =========================
st.title("🌱 RAPPORT RSE - ACTIVITÉ CEE")
st.markdown("<p class='medium-font'>Tableau de bord de suivi et d'impact de la transition énergétique</p>",
            unsafe_allow_html=True)

# Sidebar Hypothèses
st.sidebar.markdown("## 🌍 Hypothèses")
taux_efficacite = st.sidebar.slider(
    "Taux d'efficacité réelle des économies d'énergie (%)",
    min_value=10, max_value=100, value=int(TAUX_EFFICACITE_DEFAULT * 100), step=5
) / 100

uploaded_file = st.file_uploader("📁 Charger votre fichier Excel CEE", type=['xlsx', 'xls'])

if uploaded_file is not None:
    with st.spinner('Traitement des données en cours...'):
        df = load_and_process_data(uploaded_file, taux_efficacite)

    if df is not None and not df.empty:
        st.sidebar.markdown("## 🎯 Filtres")

        # Filtres
        periodes_disponibles = sorted(df['Période'].dropna().unique())
        periode_filter = st.sidebar.multiselect("📅 Période", options=periodes_disponibles,
                                                default=[
                                                    'P5'] if 'P5' in periodes_disponibles else periodes_disponibles)

        mandataires_disponibles = sorted(df['Mandataire'].dropna().unique())
        mandataire_filter = st.sidebar.multiselect("🏢 Mandataire", options=mandataires_disponibles,
                                                   default=mandataires_disponibles)

        if 'Annee_Depot' in df.columns and not df['Annee_Depot'].isnull().all():
            annees = sorted(df['Annee_Depot'].dropna().unique().astype(int))
            annee_filter = st.sidebar.multiselect("📆 Année de dépôt", options=annees, default=annees)
        else:
            annee_filter = []

        type_benef_filter = st.sidebar.multiselect("👥 Type de bénéficiaire",
                                                   options=df['Type_Beneficiaire'].unique(),
                                                   default=df['Type_Beneficiaire'].unique())

        # Application des filtres
        df_filtered = df.copy()
        if periode_filter: df_filtered = df_filtered[df_filtered['Période'].isin(periode_filter)]
        if mandataire_filter: df_filtered = df_filtered[df_filtered['Mandataire'].isin(mandataire_filter)]
        if annee_filter: df_filtered = df_filtered[df_filtered['Annee_Depot'].isin(annee_filter)]
        if type_benef_filter: df_filtered = df_filtered[df_filtered['Type_Beneficiaire'].isin(type_benef_filter)]

        # KPIs
        st.markdown("## 📊 Indicateurs Clés de Performance")
        total_dossiers = len(df_filtered)
        total_gwh_reels = df_filtered['GWh_reels_annuels'].sum()
        total_foyers = df_filtered['Nb_foyers_equivalents'].sum()
        total_primes = df_filtered['Prime_versee'].sum()
        total_couts_evites = df_filtered['Economies_euros_an'].sum()
        nb_operations_uniques = df_filtered[
            'Code équipement'].nunique() if 'Code équipement' in df_filtered.columns else 0

        col1, col2, col3, col4, col5, col6 = st.columns(6)
        with col1:
            st.metric("📋 Dossiers traités", format_number(total_dossiers), f"Période {', '.join(periode_filter)}")
        with col2:
            st.metric("⚡ Économies réelles/an", f"{format_number(total_gwh_reels, 1)} GWh",
                      f"Efficacité {taux_efficacite * 100:.0f}%")
        with col3:
            st.metric("🏠 Foyers équivalents-- chauffage et elec", format_number(total_foyers),
                      f"≈ {get_ville_equivalente(total_foyers * 2.2)}")
        with col4:
            st.metric("💰 Primes versées", f"{format_number(total_primes / 1_000_000, 1)} M€",
                      f"{format_number(total_primes / total_dossiers if total_dossiers > 0 else 0)} €/dossier")
        with col5:
            st.metric("💸 Coûts évités/an", f"{format_number(total_couts_evites / 1_000_000, 1)} M€", "sur factures")
        with col6:
            st.metric("🔬 Opérations Uniques", format_number(nb_operations_uniques))

        # TABS
        tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = st.tabs([
            "🌍 Impact Environnemental", "👥 Impact Social", "🗺️ Impact Géographique",
            "💼 Impact Économique", "📈 Analyses Détaillées", "📈 Évolution CEE (GWhc)", "📝 Hypothèses",
            "📈 Projections Futures"
        ])

        # ... (Content of tabs 1 to 7 remains the same)

        # ---------- TAB 1 : IMPACT ENVIRONNEMENTAL ----------
        with tab1:
            st.markdown("### 🌱 Contribution à la Transition Écologique (par an)")
            total_co2_evite = df_filtered['CO2_evite_tonnes_an'].sum()
            tours_terre = (total_co2_evite * 1000) / (CIRCONFERENCE_TERRE_KM * CO2_PAR_KM_VOITURE)
            arbres_equivalent = (total_co2_evite * 1000) / 25
            voitures_retirees = total_co2_evite / CO2_PAR_VOITURE_AN

            c1, c2, c3 = st.columns(3)
            with c1:
                st.metric("🌡️ CO₂ évité/an", f"{format_number(total_co2_evite)} tonnes",
                          f"≈ {format_number(voitures_retirees)} voitures retirées")
            with c2:
                st.metric("🚗 Tours de la Terre", f"{format_number(tours_terre, 1)} tours/an", "en voiture économisés")
            with c3:
                st.metric("🌳 Arbres équivalents", format_number(arbres_equivalent), "arbres plantés")

            # Groupes annuels
            if 'Annee_Depot' in df_filtered.columns and not df_filtered['Annee_Depot'].dropna().empty:
                g = df_filtered.groupby('Annee_Depot').agg(
                    GWh=('GWh_reels_annuels', 'sum'),
                    CO2=('CO2_evite_tonnes_an', 'sum')
                ).reset_index().sort_values('Annee_Depot')

                # Cumul
                g['GWh_cumul'] = g['GWh'].cumsum()
                g['CO2_cumul'] = g['CO2'].cumsum()

                st.markdown("### ⚡ Évolution (4 cadrans séparés)")
                colA, colB = st.columns(2)
                with colA:
                    fig_energy_year = go.Figure(
                        go.Bar(x=g['Annee_Depot'], y=g['GWh'], marker_color=COLOR_ENERGY, name="Énergie (GWh/an)"))
                    fig_energy_year.update_layout(title="ÉNERGIE — Annuel (GWh/an)", xaxis_title="Année",
                                                  yaxis_title="GWh/an", height=380, showlegend=False)
                    st.plotly_chart(fig_energy_year, use_container_width=True)
                with colB:
                    fig_co2_year = go.Figure(
                        go.Bar(x=g['Annee_Depot'], y=g['CO2'], marker_color=COLOR_CO2, name="CO₂ évité (t/an)"))
                    fig_co2_year.update_layout(title="CO₂ — Annuel (t/an)", xaxis_title="Année", yaxis_title="t/an",
                                               height=380, showlegend=False)
                    st.plotly_chart(fig_co2_year, use_container_width=True)
                colC, colD = st.columns(2)
                with colC:
                    fig_energy_cum = go.Figure(go.Scatter(x=g['Annee_Depot'], y=g['GWh_cumul'], mode='lines+markers',
                                                          line=dict(color=COLOR_ENERGY_ACCENT, width=3),
                                                          name="Énergie cumulée (GWh)"))
                    fig_energy_cum.update_layout(title="ÉNERGIE — Cumul (GWh)", xaxis_title="Année",
                                                 yaxis_title="GWh cumulés", height=380, showlegend=False)
                    st.plotly_chart(fig_energy_cum, use_container_width=True)
                with colD:
                    fig_co2_cum = go.Figure(go.Scatter(x=g['Annee_Depot'], y=g['CO2_cumul'], mode='lines+markers',
                                                       line=dict(color=COLOR_CO2_ACCENT, width=3, dash='dot'),
                                                       name="CO₂ cumulé (t)"))
                    fig_co2_cum.update_layout(title="CO₂ — Cumul (t)", xaxis_title="Année", yaxis_title="t cumulées",
                                              height=380, showlegend=False)
                    st.plotly_chart(fig_co2_cum, use_container_width=True)

        # ---------- TAB 2 : IMPACT SOCIAL ----------
        with tab2:
            st.markdown("### 🫂 Soutien aux Bénéficiaires")
            df_plot = df_filtered.copy()
            df_plot['Secteur'] = df_plot['Secteur'].fillna('Autre')
            df_plot['Type_Beneficiaire'] = df_plot['Type_Beneficiaire'].fillna('Non renseigné')

            benef_volume = df_plot.groupby('Type_Beneficiaire')['GWh_cumac'].sum()
            benef_counts = df_plot['Type_Beneficiaire'].value_counts()

            c1, c2 = st.columns(2)
            with c1:
                st.markdown("#### Volume CEE par type de bénéficiaire")
                fig_benef_pie = px.pie(names=benef_volume.index, values=benef_volume.values,
                                       title="Répartition du volume CEE (GWhc)")
                st.plotly_chart(fig_benef_pie, use_container_width=True)
            with c2:
                st.markdown("#### Nombre de dossiers par type de bénéficiaire")
                st.metric("🤝 Ménages en précarité", format_number(benef_counts.get('Précarité énergétique', 0)))
                st.metric("👤 Ménages Classiques", format_number(benef_counts.get('Ménage Classique', 0)))
                st.metric("🏢 Personnes Morales", format_number(benef_counts.get('Personne Morale', 0)))

            st.markdown("---")
            st.markdown("#### Répartition par secteur d'activité")
            sector_volume = df_plot.groupby('Secteur')['GWh_cumac'].sum()
            fig_sector_pie = px.pie(names=sector_volume.index, values=sector_volume.values,
                                    title="Répartition du volume CEE par secteur", hole=0.4)
            st.plotly_chart(fig_sector_pie, use_container_width=True)

            st.markdown("---")
            st.markdown("#### Évolution annuelle de la répartition (%)")
            if 'Annee_Depot' in df_filtered.columns and not df_filtered['Annee_Depot'].dropna().empty:
                # Évolution par type de bénéficiaire
                evolution_benef = df_filtered.groupby(['Annee_Depot', 'Type_Beneficiaire'])['GWh_cumac'].sum().unstack(
                    fill_value=0).apply(lambda x: 100 * x / x.sum(), axis=1).reset_index()
                evolution_benef = evolution_benef.melt(id_vars='Annee_Depot', var_name='Type_Beneficiaire',
                                                       value_name='Percentage')

                # Évolution par secteur
                evolution_sector = df_filtered.groupby(['Annee_Depot', 'Secteur'])['GWh_cumac'].sum().unstack(
                    fill_value=0).apply(lambda x: 100 * x / x.sum(), axis=1).reset_index()
                evolution_sector = evolution_sector.melt(id_vars='Annee_Depot', var_name='Secteur',
                                                         value_name='Percentage')

                col_evo1, col_evo2 = st.columns(2)
                with col_evo1:
                    fig_evol_benef = px.area(
                        evolution_benef, x='Annee_Depot', y='Percentage', color='Type_Beneficiaire',
                        title="Évolution de la part du volume par bénéficiaire (%)",
                        labels={'Percentage': '% du Volume (GWhc)', 'Annee_Depot': 'Année'}
                    )
                    st.plotly_chart(fig_evol_benef, use_container_width=True)
                with col_evo2:
                    fig_evol_sector = px.area(
                        evolution_sector, x='Annee_Depot', y='Percentage', color='Secteur',
                        title="Évolution de la part du volume par secteur (%)",
                        labels={'Percentage': '% du Volume (GWhc)', 'Annee_Depot': 'Année'}
                    )
                    st.plotly_chart(fig_evol_sector, use_container_width=True)

        # ---------- TAB 3 : CARTE GÉOGRAPHIQUE ----------
        with tab3:
            st.markdown("### 🗺️ Impact Géographique (Sélection de métrique)")
            metric_choice = st.selectbox("Métrique à cartographier", [
                "Économies d'énergie (GWh réels/an)",
                "CO₂ évité (tonnes/an)"
            ], index=0)

            if 'departement' in df_filtered.columns:
                if metric_choice.startswith("Économies"):
                    map_df = df_filtered.groupby('departement')['GWh_reels_annuels'].sum().reset_index()
                    color_col = 'GWh_reels_annuels'
                    color_title = "GWh réels/an"
                    color_scale = "Viridis"
                else:
                    map_df = df_filtered.groupby('departement')['CO2_evite_tonnes_an'].sum().reset_index()
                    color_col = 'CO2_evite_tonnes_an'
                    color_title = "Tonnes CO₂ évitées/an"
                    color_scale = "Turbo"

                fig_map = px.choropleth(
                    map_df,
                    geojson="https://raw.githubusercontent.com/gregoiredavid/france-geojson/master/departements.geojson",
                    locations='departement',
                    featureidkey="properties.code",
                    color=color_col,
                    color_continuous_scale=color_scale,
                    scope="europe",
                    title=f"Impact par département ({color_title})",
                    hover_data={'departement': True, color_col: ':.3f'}
                )
                fig_map.update_geos(fitbounds="locations", visible=False)
                fig_map.update_traces(marker_line_width=0.6, marker_line_color="white")
                fig_map.update_layout(height=620, coloraxis_colorbar_title=color_title)
                st.plotly_chart(fig_map, use_container_width=True)
            else:
                st.info("Aucun code postal / département détecté dans vos données.")

        # ---------- TAB 4 : ECONOMIQUE ----------
        with tab4:
            st.markdown("### 💰 Valorisation Économique")
            st.markdown("Analyse des flux financiers : Primes versées et économies générées pour les bénéficiaires.")

            if 'Annee_Depot' in df_filtered.columns and not df_filtered['Annee_Depot'].dropna().empty:
                eco_g = df_filtered.groupby('Annee_Depot').agg(
                    Primes=('Prime_versee', 'sum'),
                    Couts_Evites=('Economies_euros_an', 'sum')
                ).reset_index().sort_values('Annee_Depot')

                # Add cumulative calculations for both metrics
                eco_g['Primes_Cumul'] = eco_g['Primes'].cumsum()
                eco_g['Couts_Evites_Cumul'] = eco_g['Couts_Evites'].cumsum()

                # Convert to k€
                eco_g['Primes_k'] = eco_g['Primes'] / 1000
                eco_g['Primes_Cumul_k'] = eco_g['Primes_Cumul'] / 1000
                eco_g['Couts_Evites_k'] = eco_g['Couts_Evites'] / 1000
                eco_g['Couts_Evites_Cumul_k'] = eco_g['Couts_Evites_Cumul'] / 1000

                # Create a 2x2 grid for the charts
                col1, col2 = st.columns(2)

                with col1:
                    # Graph 1: Primes annuelles
                    fig_primes = px.bar(eco_g, x='Annee_Depot', y='Primes_k',
                                        title="Primes versées annuelles",
                                        text=eco_g['Primes_k'].apply(lambda x: f'{x:,.0f}'))
                    fig_primes.update_layout(yaxis_title="Primes Versées (k€)", xaxis_title="Année", showlegend=False)
                    fig_primes.update_traces(marker_color=COLOR_PRIME, texttemplate='%{text} k', textposition='outside')
                    st.plotly_chart(fig_primes, use_container_width=True)

                    # Graph 2: Coûts évités annuels
                    fig_couts_evites = px.bar(eco_g, x='Annee_Depot', y='Couts_Evites_k',
                                              title="Coûts évités annuels sur factures",
                                              text=eco_g['Couts_Evites_k'].apply(lambda x: f'{x:,.0f}'))
                    fig_couts_evites.update_layout(yaxis_title="Coûts Évités (k€)", xaxis_title="Année",
                                                   showlegend=False)
                    fig_couts_evites.update_traces(marker_color=COLOR_ECONOMY, texttemplate='%{text} k',
                                                   textposition='outside')
                    st.plotly_chart(fig_couts_evites, use_container_width=True)

                with col2:
                    # Graph 3: Primes cumulées
                    fig_primes_cumul = go.Figure(
                        go.Scatter(x=eco_g['Annee_Depot'], y=eco_g['Primes_Cumul_k'], mode='lines+markers',
                                   line=dict(color=COLOR_PRIME_ACCENT, width=3),
                                   name="Primes cumulées (k€)"))
                    fig_primes_cumul.update_layout(title="Primes versées cumulées", xaxis_title="Année",
                                                   yaxis_title="Primes Cumulées (k€)", showlegend=False)
                    st.plotly_chart(fig_primes_cumul, use_container_width=True)

                    # Graph 4: Coûts évités cumulés
                    fig_couts_cumul = go.Figure(
                        go.Scatter(x=eco_g['Annee_Depot'], y=eco_g['Couts_Evites_Cumul_k'], mode='lines+markers',
                                   line=dict(color=COLOR_ECONOMY_ACCENT, width=3),
                                   name="Coûts évités cumulés (k€)"))
                    fig_couts_cumul.update_layout(title="Coûts évités cumulés sur factures", xaxis_title="Année",
                                                  yaxis_title="Coûts Évités Cumulés (k€)", showlegend=False)
                    st.plotly_chart(fig_couts_cumul, use_container_width=True)

        # ---------- TAB 5 : ANALYSES DÉTAILLÉES ----------
        with tab5:
            st.markdown("### 🔬 Analyses Détaillées par Opération")
            if 'Code équipement' in df_filtered.columns:
                agg = df_filtered.groupby('Code équipement').agg(
                    GWh_cumac=('GWh_cumac', 'sum'),
                    Nb_Dossiers=('GWh_cumac', 'size'),
                    GWh_reels_annuels=('GWh_reels_annuels', 'sum'),
                    CO2_evite_tonnes_an=('CO2_evite_tonnes_an', 'sum')
                ).reset_index()
                indicateur = st.radio("Indicateur pour le classement et le graphique",
                                      ["Nombre de dossiers", "GWh cumac", "GWh réels/an", "CO₂ évité (t/an)"],
                                      horizontal=True)

                if indicateur == "Nombre de dossiers":
                    value_col, title_bar, color_bar = 'Nb_Dossiers', "Top opérations – Nombre de dossiers", "#fca311"
                elif indicateur == "GWh cumac":
                    value_col, title_bar, color_bar = 'GWh_cumac', "Top opérations – Volume GWh cumac", "#4c78a8"
                elif indicateur == "GWh réels/an":
                    value_col, title_bar, color_bar = 'GWh_reels_annuels', "Top opérations – GWh réels/an", COLOR_ENERGY
                else:  # CO2
                    value_col, title_bar, color_bar = 'CO2_evite_tonnes_an', "Top opérations – CO₂ évité (t/an)", COLOR_CO2

                top = agg.nlargest(10, value_col).sort_values(value_col)
                fig_ops_bar = go.Figure(
                    go.Bar(x=top[value_col], y=top['Code équipement'], orientation='h', marker_color=color_bar))
                fig_ops_bar.update_layout(title=title_bar, xaxis_title=indicateur, yaxis_title="Code équipement",
                                          height=520, margin=dict(l=120, r=40, t=60, b=40))
                st.plotly_chart(fig_ops_bar, use_container_width=True)

                st.markdown("---")
                st.markdown(f"#### Évolution annuelle par {indicateur} (Top 10 opérations)")
                top_10_codes = agg.nlargest(10, value_col)['Code équipement'].tolist()
                df_top_ops = df_filtered[df_filtered['Code équipement'].isin(top_10_codes)]

                if value_col == 'Nb_Dossiers':
                    evolution_ops = df_top_ops.groupby(['Annee_Depot', 'Code équipement']).size().reset_index(
                        name=value_col)
                else:
                    evolution_ops = df_top_ops.groupby(['Annee_Depot', 'Code équipement'])[
                        value_col].sum().reset_index()

                fig_evol_ops = px.bar(
                    evolution_ops, x='Annee_Depot', y=value_col, color='Code équipement',
                    title=f"Évolution annuelle par {indicateur} (Top 10 opérations)",
                    labels={value_col: indicateur, 'Annee_Depot': 'Année'},
                    barmode='stack'
                )
                st.plotly_chart(fig_evol_ops, use_container_width=True)

        # ---------- TAB 6 : ÉVOLUTION CEE (GWHC) ----------
        with tab6:
            st.markdown("### 📈 Évolution du Volume CEE (GWh cumac) par an")
            if 'Annee_Depot' in df_filtered.columns and not df_filtered['Annee_Depot'].dropna().empty:
                gwhc_yearly = df_filtered.groupby('Annee_Depot')['GWh_cumac'].sum().reset_index()

                fig_gwhc_yearly = px.bar(
                    gwhc_yearly,
                    x='Annee_Depot',
                    y='GWh_cumac',
                    title="Volume CEE (GWhc) par Année de Dépôt",
                    labels={'GWh_cumac': 'GWh cumac', 'Annee_Depot': 'Année'},
                    color='GWh_cumac',
                    color_continuous_scale='Blues',
                    text_auto='.2s'
                )
                fig_gwhc_yearly.update_traces(textposition='outside')
                fig_gwhc_yearly.update_layout(
                    xaxis_title="Année",
                    yaxis_title="GWh cumac",
                    font=dict(size=14)
                )
                st.plotly_chart(fig_gwhc_yearly, use_container_width=True)
            else:
                st.info("Les données d'année de dépôt sont nécessaires pour afficher cette évolution.")

        # ---------- TAB 7 : HYPOTHÈSES ----------
        with tab7:
            st.markdown("### 📝 Hypothèses de Travail")
            st.json({
                "Périmètre": "France",
                "Hypothèses d'équivalence": {
                    "Voitures retirées": f"Basé sur {CO2_PAR_VOITURE_AN} tCO2/an/voiture.",
                    "Tours de la Terre": f"Basé sur {CO2_PAR_KM_VOITURE} kgCO2/km et une circonférence de {CIRCONFERENCE_TERRE_KM} km.",
                    "Arbres équivalents": "Basé sur 25 kgCO2/an/arbre (valeur indicative)."
                },
                "Facteurs de Conversion (Cumac -> kWh/an)": {k: round(v, 4) for k, v in FACTEUR_CUMAC_TO_KWH.items()},
                "Durées de vie des équipements (années)": DUREE_VIE_EQUIPEMENT,
                "Constantes d'Impact": {
                    "Consommation moyenne d'un foyer (kWh/an)": CONSO_MOYENNE_FOYER_KWH,
                    "Émissions CO2 (kg/kWh)": EMISSION_CO2_KWH,
                    "Coût du chauffage (€/kWh)": COUT_CHAUFFAGE_KWH
                }
            })

        # ---------- TAB 8 : PROJECTIONS FUTURES ----------
        with tab8:
            st.markdown("### 📈 Projections Futures des Économies d'Énergie")
            st.info(
                "Cette section modélise l'évolution du flux d'économies d'énergie annuelles en tenant compte de la durée de vie des équipements.")

            horizon = st.slider("Horizon de projection (années)", 10, 40, 20)

            if 'Annee_Depot' in df_filtered.columns and not df_filtered.dropna(
                    subset=['Annee_Depot', 'Duree_Vie']).empty:

                start_year = int(df_filtered['Annee_Depot'].min())
                current_year = datetime.now().year
                end_year = current_year + horizon

                projection_years = list(range(start_year, end_year + 1))
                projection_breakdown_list = []

                # Determine top 5 op types
                top_ops = df_filtered.groupby('FacteurKey')['GWh_reels_annuels'].sum().nlargest(5).index.tolist()

                df_proj = df_filtered.copy()
                df_proj['Type Opération'] = df_proj['FacteurKey'].apply(lambda x: x if x in top_ops else 'Autres')

                for year in projection_years:
                    # An operation is active if the projection year is between its start and end of life
                    active_ops = df_proj[
                        (df_proj['Annee_Depot'] <= year) &
                        (df_proj['Annee_Depot'] + df_proj['Duree_Vie'] > year)
                        ]

                    # Breakdown by operation type (with 'Autres')
                    breakdown = active_ops.groupby('Type Opération')['GWh_reels_annuels'].sum()
                    for op_type, saving in breakdown.items():
                        projection_breakdown_list.append({
                            'Année': year,
                            'Type Opération': op_type,
                            'Économies GWh/an': saving
                        })

                # --- Graph 1: Total Projection ---
                if projection_breakdown_list:
                    projection_df = pd.DataFrame(projection_breakdown_list)
                    total_projection_df = projection_df.groupby('Année')['Économies GWh/an'].sum().reset_index()

                    st.markdown("#### Évolution du total des économies annuelles")
                    fig_projection = px.area(
                        total_projection_df,
                        x='Année',
                        y='Économies GWh/an',
                        title=f"Projection du flux d'économies sur {horizon} ans",
                    )
                    fig_projection.add_vline(x=current_year, line_width=2, line_dash="dash", line_color="red",
                                             annotation_text="Aujourd'hui")
                    fig_projection.update_layout(
                        xaxis_title="Année",
                        yaxis_title="GWh réels / an",
                        font=dict(size=14)
                    )
                    st.plotly_chart(fig_projection, use_container_width=True)

                    # --- Graph 2: Breakdown Projection ---
                    st.markdown("#### Composition des économies projetées")
                    st.info(
                        "Ce graphique décompose la projection totale pour montrer la contribution des 5 principaux types d'opérations. Les autres sont regroupés pour plus de lisibilité.")

                    fig_projection_breakdown = px.area(
                        projection_df,
                        x='Année',
                        y='Économies GWh/an',
                        color='Type Opération',
                        title=f"Composition des économies annuelles projetées",
                        labels={'Économies GWh/an': 'GWh réels / an'},
                        # Ensure 'Autres' is at the bottom for clarity
                        category_orders={"Type Opération": top_ops + ['Autres']}
                    )
                    fig_projection_breakdown.add_vline(x=current_year, line_width=2, line_dash="dash",
                                                       line_color="white", annotation_text="Aujourd'hui")
                    fig_projection_breakdown.update_layout(
                        xaxis_title="Année",
                        yaxis_title="GWh réels / an",
                        font=dict(size=14)
                    )
                    st.plotly_chart(fig_projection_breakdown, use_container_width=True)


    else:
        st.warning("Le fichier a été chargé mais ne contient aucune ligne exploitable.")
else:
    st.info("👋 Bienvenue ! Veuillez charger votre fichier de données pour commencer l'analyse.")



