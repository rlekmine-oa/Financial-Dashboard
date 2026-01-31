import streamlit as st
import concurrent.futures 
import pandas as pd
import yfinance as yf
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import os
import numpy as np

# 1. CONFIGURATION
st.set_page_config(page_title="Mon Patrimoine", layout="wide", page_icon="🏦")

# --- BOUTON RESET CACHE 
if st.sidebar.button("🗑️ FORCER RECHARGEMENT DONNÉES"):
    st.cache_data.clear()
    st.rerun()

# ---------------------------------------------------------
# 2. FICHIERS ET PARAMÈTRES
# ---------------------------------------------------------
FILE_BOURSE = "Stratégie et Feuille de Route_Bourse.xlsm"
BASE_NAME_BANQUE = "Comptes" 
SHEET_BOURSE = "Historique_Ordres"

# --- FONCTION : CHARGEMENT BOURSE 
@st.cache_data
def load_stock_data(file_path):
    if not os.path.exists(file_path): return None
    try:
        temp_df = pd.read_excel(file_path, sheet_name=SHEET_BOURSE, header=None)
    except: return "ERROR_SHEET"

    header_row = -1
    for i, row in temp_df.iterrows():
        row_str = [str(val).strip() for val in row.values]
        if "Date exécution Ordres" in row_str:
            header_row = i
            break
    if header_row == -1: return "ERROR_HEADER"

    df = pd.read_excel(file_path, sheet_name=SHEET_BOURSE, header=header_row)
    df.columns = df.columns.str.strip()
    
    ticker_col = next((c for c in df.columns if c.lower() in ['ticker', 'symbole', 'code']), None)
    if not ticker_col: return "ERROR_NO_TICKER"

    type_asset_col = None
    candidates = ['type d\'actif', 'type actif', 'support', 'catégorie', 'nature', 'type instrument']
    
    # 1. Recherche exacte dans les candidats
    for c in df.columns:
        if c.lower() in candidates:
            type_asset_col = c
            break
    
    # 2. Si pas trouvé, recherche une colonne "Type" qui n'est PAS "Type d'ordre"
    if not type_asset_col:
        for c in df.columns:
            if "type" in c.lower() and "ordre" not in c.lower():
                type_asset_col = c
                break

    # Dictionnaire de renommage
    rename_map = {
        "Date exécution Ordres": "Date",
        "Actif": "Nom",
        "Type d'ordre": "Sens",
        "Nombre d'ordre": "Qte",
        "Valeur total Ordres + frais": "Montant",
        ticker_col: "Ticker"
    }
    
    # Si on a trouvé la colonne Type, on l'ajoute au renommage
    if type_asset_col:
        rename_map[type_asset_col] = "Type_Actif"

    df = df.rename(columns=rename_map)
    
    # Si la colonne n'existait pas, on met "Action" par défaut pour éviter les bugs
    if "Type_Actif" not in df.columns:
        df["Type_Actif"] = "Action" 
    
    df['Date'] = pd.to_datetime(df['Date'])
    df['Ticker'] = df['Ticker'].astype(str).str.strip().str.upper()
    df = df.dropna(subset=['Ticker'])
    df = df[df['Ticker'] != 'NAN']

    for col in ['Qte', 'Montant']:
        if col in df.columns:
            if df[col].dtype == 'object':
                df[col] = df[col].astype(str).str.replace('€', '').str.replace(' ', '').str.replace(',', '.')
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0)

    return df

# --- FONCTION : CHARGEMENT BANQUE ---
@st.cache_data
def load_bank_data(base_name):
    file_path = None
    if os.path.exists(base_name + ".xlsx"): file_path = base_name + ".xlsx"
    elif os.path.exists(base_name + ".xls"): file_path = base_name + ".xls"
    
    if file_path is None: return "NOT_FOUND"
            
    try:
        temp_df = pd.read_excel(file_path, header=None)
        header_row = -1
        
        for i, row in temp_df.iterrows():
            row_str = [str(val).strip().lower() for val in row.values]
            if "compte" in row_str and "montant" in row_str:
                header_row = i
                break
        
        if header_row == -1: return "ERROR_COLUMNS"

        df = pd.read_excel(file_path, header=header_row)
        df.columns = df.columns.str.strip()
        
        col_date = next((c for c in df.columns if c.lower() == 'date'), None)
        col_compte = next((c for c in df.columns if c.lower() == 'compte'), None)
        col_montant = next((c for c in df.columns if c.lower() == 'montant'), None)
        
        if not col_compte or not col_montant: return "ERROR_COLUMNS"
            
        rename_dict = {col_compte: "Compte", col_montant: "Montant"}
        if col_date: rename_dict[col_date] = "Date"
        df = df.rename(columns=rename_dict)
        
        df = df.dropna(subset=['Compte'])
        if "Unnamed: 0" in df.columns: del df["Unnamed: 0"]
        
        if df['Montant'].dtype == 'object':
            df['Montant'] = df['Montant'].astype(str).str.replace('€', '').str.replace(' ', '').str.replace(',', '.')
        df['Montant'] = pd.to_numeric(df['Montant'], errors='coerce').fillna(0)
        
        if "Date" in df.columns:
            df['Date'] = pd.to_datetime(df['Date'], dayfirst=True, errors='coerce')
            if df['Date'].isna().all(): df['Date'] = datetime.now()
            else: df['Date'] = df['Date'].ffill().fillna(datetime.now())
        else:
            df['Date'] = datetime.now()
            
        return df
    except Exception as e:
        return f"ERROR_READ: {str(e)}"

# --- FONCTION YFINANCE ---
def get_market_data(tickers, start_date, end_date_limit):
    if not tickers: return None, None, None
    final_date = max(datetime.now(), end_date_limit) + timedelta(days=1)
    
    all_days = pd.date_range(start=start_date, end=final_date, freq='D')
    
    try:
        download = yf.download(tickers, start=start_date, end=final_date, progress=False)['Close']
        if download.empty: return None, None, None
        
        if isinstance(download, pd.Series):
            download = download.to_frame()
            download.columns = tickers
            
        # 1. Données brutes
        raw_prices = download.copy()
        
        # 2. Données lissées
        prices = download.reindex(all_days).ffill().bfill()
        
        return prices, all_days, raw_prices
    except: return None, None, None


# ---  FONCTION DONNÉES FONDAMENTALES ---
def fetch_single_fundamental(t):
    """Fonction helper pour récupérer les infos d'un seul ticker"""
    try:
        ticker = yf.Ticker(t)
        # On force la récupération rapide si possible, sinon info standard
        info = ticker.info
        
        # Si info est vide ou None, on sort
        if not info:
            return None

        # Prix & EPS
        # On cherche le prix dans plusieurs clés possibles par sécurité
        price = info.get('currentPrice') or info.get('regularMarketPrice') or info.get('previousClose')
        eps = info.get('trailingEps') or info.get('forwardEps')
        
        # Calcul PER Manuel (plus fiable que le trailingPE de Yahoo parfois manquant)
        per = None
        if price and eps and eps != 0:
            per = price / eps
            
        # Autres indicateurs
        roe = info.get('returnOnEquity')
        margin = info.get('profitMargins') # Marge Nette
        debt_eq = info.get('debtToEquity') # Souvent en %
        div = info.get('dividendYield') # Rendement du dividende
        currency = info.get('currency', '???') # Récupération de la devise

        return {
            "Ticker": t,
            "Prix Actuel": price,
            "Devise": currency,  # <--- Ajout important pour vérification
            "BPA (EPS)": eps,
            "PER": per,
            "ROE": roe * 100 if roe else None,
            "Marge Nette": margin*100,
            "Dette/Eq": debt_eq / 100 if debt_eq else None, 
            "Dividende %": div
        }
    except Exception as e:
        # En production, on pourrait logger l'erreur : print(f"Erreur sur {t}: {e}")
        return None

@st.cache_data(ttl=3600*12) # Cache augmenté à 12h (inutile de rafraîchir ça souvent)
def get_fundamentals_data(tickers):
    """Récupère PER, ROE, Marges, Dette en parallèle."""
    if not tickers: return pd.DataFrame()
    
    data = []
    total = len(tickers)
    
    # ThreadPool pour lancer plusieurs requêtes Yahoo en même temps
    # max_workers=10 requêtes simultanées
    with st.spinner(f"Analyse fondamentale de {total} actifs en parallèle..."):
        with concurrent.futures.ThreadPoolExecutor(max_workers=10) as executor:
            # On lance toutes les tâches
            future_to_ticker = {executor.submit(fetch_single_fundamental, t): t for t in tickers}
            
            # On récolte les résultats au fur et à mesure
            for future in concurrent.futures.as_completed(future_to_ticker):
                res = future.result()
                if res:
                    data.append(res)
    
    if not data:
        return pd.DataFrame()

    return pd.DataFrame(data)
# ---------------------------------------------------------
# 3. MOTEUR DE CALCUL
# ---------------------------------------------------------
st.title("🏛️ Mon Tableau de Bord Patrimonial")

with st.spinner('Actualisation des marchés et calculs...'):
    # 1. Chargement
    df_ops = load_stock_data(FILE_BOURSE)
    if isinstance(df_ops, str) or df_ops is None:
        st.error(f"Erreur Bourse: {df_ops}")
        st.stop()
    
    df_banque_raw = load_bank_data(BASE_NAME_BANQUE)
    
    start_date_bourse = df_ops['Date'].min()
    if not isinstance(df_banque_raw, str) and not df_banque_raw.empty:
        max_date_banque = df_banque_raw['Date'].max()
        start_date_banque = df_banque_raw['Date'].min()
    else:
        max_date_banque = datetime.now()
        start_date_banque = start_date_bourse

    # 2. Marché
    tickers = df_ops['Ticker'].unique().tolist()
    
    # --- APPEL NOUVELLE FONCTION FONDAMENTAUX ---
    df_fundamentals = get_fundamentals_data(tickers)
    
    prices, all_days, raw_prices = get_market_data(tickers, min(start_date_bourse, start_date_banque), max_date_banque)

    if prices is None:
        st.error("Erreur Yahoo Finance (Aucune donnée récupérée).")
        st.stop()

    # --- HISTORIQUE BOURSE ---
    history_records = []
    portfolio = {} 
    ticker_names = {}
    somme_brute_achats_excel = 0.0
    ops_group = df_ops.groupby('Date')
    
    for day in all_days:
        if day in ops_group.groups:
            todays_ops = ops_group.get_group(day)
            for _, row in todays_ops.iterrows():
                t = row['Ticker']
                q = row['Qte']; m = row['Montant']
                if pd.isna(q): q = 0.0
                if pd.isna(m): m = 0.0
                
                sens = str(row['Sens']).strip().lower()
                ticker_names[t] = row['Nom']
                
                if t not in portfolio: portfolio[t] = {'qty': 0.0, 'investi': 0.0}
                
                if 'vente' in sens:
                    if portfolio[t]['qty'] > 0:
                        ratio = q / portfolio[t]['qty']
                        if ratio > 1: ratio = 1
                        portfolio[t]['investi'] -= (portfolio[t]['investi'] * ratio)
                        portfolio[t]['qty'] -= q
                elif 'ost' in sens:
                    if q > 0: portfolio[t]['qty'] += q
                else: # Achat
                    somme_brute_achats_excel += m
                    portfolio[t]['qty'] += q
                    portfolio[t]['investi'] += m
                
                if portfolio[t]['qty'] < 0.00001: 
                    portfolio[t]['qty'] = 0; portfolio[t]['investi'] = 0

        valeur_jour = 0.0
        investi_jour = 0.0
        for t, data in portfolio.items():
            qty = data['qty']
            if qty > 0:
                if t in prices.columns:
                    try:
                        price = prices.at[day, t]
                        if pd.isna(price): price = prices[t].iloc[-1]
                        if pd.isna(price): price = 0.0
                    except: price = 0.0
                else: price = 0.0
                
                valeur_jour += qty * price
                investi_jour += data['investi']
        
        history_records.append({'Date': day, 'Valeur_Bourse': valeur_jour, 'Investi_Bourse': investi_jour})

    df_hist_bourse = pd.DataFrame(history_records).set_index('Date')
    df_hist_bourse['PV_Bourse'] = df_hist_bourse['Valeur_Bourse'] - df_hist_bourse['Investi_Bourse']
    df_hist_bourse = df_hist_bourse.fillna(0.0)

    # --- HISTORIQUE BANQUE ---
    df_hist_banque = pd.DataFrame(index=all_days)
    last_banque_state = pd.DataFrame()
    
    if not isinstance(df_banque_raw, str) and not df_banque_raw.empty:
        pivot_banque = df_banque_raw.pivot_table(index='Date', columns='Compte', values='Montant', aggfunc='last')
        df_hist_banque_detail = pivot_banque.reindex(all_days).ffill().fillna(0)
        df_hist_banque['Valeur_Banque'] = df_hist_banque_detail.sum(axis=1)
        
        last_banque_state = df_hist_banque_detail.iloc[-1].reset_index()
        last_banque_state.columns = ['Compte', 'Montant']
        last_banque_state = last_banque_state[last_banque_state['Montant'] > 0]
    else:
        df_hist_banque['Valeur_Banque'] = 0

    # --- FUSION GLOBALE ---
    df_global = pd.concat([df_hist_bourse, df_hist_banque], axis=1).ffill().fillna(0)
    df_global['Patrimoine_Total'] = df_global['Valeur_Bourse'] + df_global['Valeur_Banque']
    df_global_monthly = df_global.resample('M').last()

    # --- DÉTAILS BOURSE + VARIATION ROBUSTE ---
    details_data = []
    last_prices_smooth = prices.iloc[-1]
    
    if len(raw_prices) >= 2:
        last_prices_raw = raw_prices.iloc[-1]
        prev_prices_raw = raw_prices.iloc[-2]
    else:
        last_prices_raw = raw_prices.iloc[-1]
        prev_prices_raw = last_prices_raw

    total_var_24h_euro = 0.0
    sim_assets = {} 
    
    # Liste pour filtrer le tableau fondamental
    active_tickers_list = []

    for t, data in portfolio.items():
        qty = data['qty']
        if qty > 0.001:
            active_tickers_list.append(t)
            curr_price = last_prices_smooth[t] if t in last_prices_smooth else 0.0
            if pd.isna(curr_price): curr_price = 0.0
            
            p_now_raw = last_prices_raw[t] if t in last_prices_raw else 0
            p_prev_raw = prev_prices_raw[t] if t in prev_prices_raw else 0
            
            if pd.isna(p_now_raw): p_now_raw = 0.0
            if pd.isna(p_prev_raw): p_prev_raw = 0.0
            
            var_24h_asset_pct = 0.0
            var_24h_asset_euro = 0.0
            
            if t in raw_prices.columns:
                asset_history = raw_prices[t].dropna()
                if len(asset_history) >= 2:
                    p_now = asset_history.iloc[-1]
                    p_prev = asset_history.iloc[-2]
                    
                    if p_prev > 0:
                        var_24h_asset_pct = (p_now - p_prev) / p_prev * 100
                        var_24h_asset_euro = qty * (p_now - p_prev)
            
            total_var_24h_euro += var_24h_asset_euro

            val_actuelle = qty * curr_price
            
            if val_actuelle > 0:
                sim_assets[t] = val_actuelle
            
            details_data.append({
                "Actif": ticker_names.get(t, t), "Ticker": t, "Qté": qty, 
                "PRU": data['investi'] / qty, "Cours": curr_price, 
                "Var 24h %": var_24h_asset_pct,
                "Var 24h €": var_24h_asset_euro,
                "Valeur": val_actuelle, "Plus-Value": val_actuelle - data['investi'],
                "Rendement %": (val_actuelle - data['investi']) / data['investi'] * 100 if data['investi'] > 0 else 0
            })
    
    df_details = pd.DataFrame(details_data)

    val_bourse_now = df_hist_bourse.iloc[-1]['Valeur_Bourse']
    valeur_pf_impliquee_hier = val_bourse_now - total_var_24h_euro
    
    if valeur_pf_impliquee_hier > 0:
        var_24h_bourse_pct = (total_var_24h_euro / valeur_pf_impliquee_hier) * 100
    else:
        var_24h_bourse_pct = 0.0

# ---------------------------------------------------------
# 4. AFFICHAGE
# ---------------------------------------------------------
last_vals = df_global.iloc[-1]
val_total_now = last_vals['Patrimoine_Total']
val_banque_now = last_vals['Valeur_Banque']
if pd.isna(val_bourse_now): val_bourse_now = 0.0

pv_bourse_now = df_hist_bourse.iloc[-1]['PV_Bourse']
invest_bourse_now = df_hist_bourse.iloc[-1]['Investi_Bourse']

if pd.isna(val_total_now): val_total_now = 0.0
if pd.isna(val_banque_now): val_banque_now = 0.0
if pd.isna(pv_bourse_now): pv_bourse_now = 0.0
if pd.isna(invest_bourse_now): invest_bourse_now = 0.0
if pd.isna(total_var_24h_euro): total_var_24h_euro = 0.0

var_24h_glob_euro = total_var_24h_euro 
if (val_total_now - var_24h_glob_euro) > 0:
    var_24h_glob_pct = (var_24h_glob_euro / (val_total_now - var_24h_glob_euro)) * 100
else:
    var_24h_glob_pct = 0.0


tab_patrimoine, tab_bourse, tab_suivi, tab_simu = st.tabs(["🌍 Patrimoine Global", "📈 Portefeuille", "🎯 Suivi Objectif (DCA)", "🔮 Monte Carlo"])

# === ONGLET 1 : PATRIMOINE GLOBAL ===
with tab_patrimoine:
    st.markdown("### Vue d'ensemble")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("🌍 Patrimoine Total", f"{val_total_now:,.2f} €")
    c2.metric("📅 Variation 24h", f"{var_24h_glob_euro:+.2f} €", f"{var_24h_glob_pct:+.2f} %")
    c3.metric("🏦 Banque (Cash)", f"{val_banque_now:,.2f} €")
    c4.metric("⚡ Bourse (Actions)", f"{val_bourse_now:,.2f} €")
    
    st.markdown("---")
    
    st.subheader("⏳ Évolution du Patrimoine Total (Mensuel)")
    fig_global = go.Figure()
    fig_global.add_trace(go.Bar(
        x=df_global_monthly.index, y=df_global_monthly['Valeur_Banque'],
        name='Banque (Sécurisé)', marker_color='#2ca02c'
    ))
    fig_global.add_trace(go.Bar(
        x=df_global_monthly.index, y=df_global_monthly['Valeur_Bourse'],
        name='Bourse (Risque)', marker_color='#1f77b4'
    ))
    fig_global.update_layout(
        height=500, hovermode="x unified", title="Accumulation Richesse (Fin de mois)",
        barmode='stack', xaxis_title="Mois", yaxis_title="Montant (€)"
    )
    fig_global.update_xaxes(dtick="M1", tickformat="%b %Y") 
    st.plotly_chart(fig_global, use_container_width=True)

    st.markdown("---")
    
    if not last_banque_state.empty:
        st.subheader("📊 Répartition de vos actifs")
        
        SEUIL_BAS = 10000; SEUIL_HAUT = 11000
        surplus_fixe = val_banque_now - 10000
        surplus_variable = val_banque_now - (val_bourse_now * 0.2)
        reste_a_investir = min(surplus_fixe, surplus_variable)
        
        sous_titre = f"<br><span style='font-size:13px; color: grey'>(Reste à investir : {reste_a_investir:,.2f} €)</span>"
        
        if val_banque_now < SEUIL_BAS or val_banque_now > SEUIL_HAUT:
            prefix_securite = "⚠️ 🛡️"
            label_banque = f"<span style='color:#dc3545'>{prefix_securite} Matelas de Sécurité : {val_banque_now:,.2f} €</span>{sous_titre}"
            color_banque = '#dc3545'
        else:
            prefix_securite = "🛡️"
            label_banque = f"{prefix_securite} Matelas de Sécurité : {val_banque_now:,.2f} €{sous_titre}"
            color_banque = '#2ca02c'
            
        color_bourse = '#1f77b4'
        label_bourse = f"⚡ Portefeuille Boursier : {val_bourse_now:,.2f} €"
        
        df_macro = pd.DataFrame([
            {"Label_Legende": label_banque, "Montant": val_banque_now, "Couleur": color_banque},
            {"Label_Legende": label_bourse, "Montant": val_bourse_now, "Couleur": color_bourse}
        ])
        
        row_bourse = pd.DataFrame([{"Compte": "Actions (Global)", "Montant": val_bourse_now}])
        df_global_alloc = pd.concat([last_banque_state, row_bourse], ignore_index=True)
        
        col_macro, col_micro = st.columns(2)
        
        with col_macro:
            st.markdown("**1. Allocation Stratégique (Macro)**")
            fig_macro = px.pie(df_macro, values='Montant', names='Label_Legende', hole=0.5,
                               color_discrete_sequence=df_macro['Couleur'].tolist())
            fig_macro.update_traces(textposition='inside', textinfo='percent', textfont_size=16)
            fig_macro.update_layout(legend=dict(font=dict(size=16)))
            st.plotly_chart(fig_macro, use_container_width=True)
            
        with col_micro:
            st.markdown("**2. Allocation Détaillée (Micro)**")
            fig_micro = px.pie(df_global_alloc, values='Montant', names='Compte', hole=0.5,
                               color_discrete_sequence=px.colors.qualitative.Pastel)
            fig_micro.update_traces(textposition='inside', textinfo='percent+label')
            st.plotly_chart(fig_micro, use_container_width=True)
            
        st.markdown("---")
        st.subheader("Détail par Compte")
        st.dataframe(
            df_global_alloc.sort_values("Montant", ascending=False),
            column_config={"Montant": st.column_config.NumberColumn(format="%.2f €")},
            hide_index=True, use_container_width=True
        )
    else:
        if isinstance(df_banque_raw, str): st.warning(f"Info Banque : {df_banque_raw}")
        else: st.info("Aucune donnée bancaire trouvée.")

# === ONGLET 2 : PORTEFEUILLE BOURSIER ===
with tab_bourse:
    st.markdown("### Analyse détaillée")
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("💰 Valeur Actions", f"{val_bourse_now:,.2f} €")
    k2.metric("📅 Variation 24h", f"{total_var_24h_euro:+.2f} €", f"{var_24h_bourse_pct:+.2f} %")
    k3.metric("🚀 Plus-Value Latente", f"{pv_bourse_now:+.2f} €", f"{(pv_bourse_now/invest_bourse_now*100 if invest_bourse_now>0 else 0):+.2f} %")
    k4.metric("🏦 Capital Investi", f"{invest_bourse_now:,.2f} €")
    
    st.markdown("---")
    
    col_graphs, col_repart = st.columns([2, 1])
    with col_graphs:
        st.subheader("Performance Historique")
        
        st.caption("Évolution de la valeur totale")
        fig_hist = px.line(df_hist_bourse, x=df_hist_bourse.index, y=['Valeur_Bourse', 'Investi_Bourse'], 
                      color_discrete_map={'Valeur_Bourse': '#1F4E78', 'Investi_Bourse': '#B0B0B0'})
        st.plotly_chart(fig_hist, use_container_width=True)
        
        st.caption("Évolution des Gains/Pertes Latents")
        if 'PV_Bourse' in df_hist_bourse.columns:
            df_hist_bourse['Color_PV'] = df_hist_bourse['PV_Bourse'].apply(lambda x: 'Gain' if x>=0 else 'Perte')
            fig_pv = px.area(df_hist_bourse, x=df_hist_bourse.index, y='PV_Bourse', color='Color_PV',
                             color_discrete_map={'Gain': '#38761D', 'Perte': '#9C0006'})
            st.plotly_chart(fig_pv, use_container_width=True)
    
    with col_repart:
        st.subheader("Allocation & Cible")
        
        # --- 1. Graphique par Actif ---
        if not df_details.empty:
            map_types = df_ops.drop_duplicates('Ticker').set_index('Ticker')['Type_Actif'].to_dict()
            df_details['Type_Actif'] = df_details['Ticker'].map(map_types).fillna('Action')

            st.caption("Répartition par Actif")
            fig_pie_stock = px.pie(df_details, values='Valeur', names='Actif', hole=0.4)
            fig_pie_stock.update_layout(showlegend=False, margin=dict(t=0, b=0, l=0, r=0), height=250)
            st.plotly_chart(fig_pie_stock, use_container_width=True)

            st.markdown("---")

            # --- 2. Graphique par Type ---
            st.caption("Répartition par Type")
            df_par_type = df_details.groupby('Type_Actif')['Valeur'].sum().reset_index()
            
            if not df_par_type.empty:
                fig_pie_type = px.pie(df_par_type, values='Valeur', names='Type_Actif', hole=0.4,
                                     color_discrete_sequence=px.colors.qualitative.Pastel)
                fig_pie_type.update_traces(textposition='inside', textinfo='percent+label')
                fig_pie_type.update_layout(showlegend=False, margin=dict(t=0, b=0, l=0, r=0), height=250)
                st.plotly_chart(fig_pie_type, use_container_width=True)
            
            st.markdown("---")

            # --- 3. Calculateur Smart DCA (Pas de Vente) ---
            st.write("**⚖️ Prochain Investissement**")
            
            TARGETS = {"ETF": 0.70, "Action": 0.28, "ETC": 0.02}
            
            total_pf_value = df_details['Valeur'].sum()
            current_totals = {k: 0.0 for k in TARGETS.keys()}
            
            # Classification
            for _, row in df_par_type.iterrows():
                type_name = str(row['Type_Actif']).upper()
                val = row['Valeur']
                if "ETF" in type_name or "TRACKER" in type_name:
                    current_totals["ETF"] += val
                elif "ETC" in type_name or "OR" in type_name or "GOLD" in type_name or "MATIÈRE" in type_name:
                    current_totals["ETC"] += val
                else:
                    current_totals["Action"] += val

            # Calcul des écarts pour trouver la priorité
            gaps = {}
            for cat, target in TARGETS.items():
                theorique = total_pf_value * target
                gap = theorique - current_totals[cat]
                gaps[cat] = gap

            # La priorité est celle qui a le plus grand manque (gap le plus positif)
            # On ignore les gaps négatifs (surplus) pour la priorité d'achat
            positive_gaps = {k: v for k, v in gaps.items() if v > 0}
            priority_cat = max(positive_gaps, key=positive_gaps.get) if positive_gaps else None

            # Affichage
            for category, target_pct in TARGETS.items():
                gap = gaps[category]
                val_reelle = current_totals[category]
                val_theorique = total_pf_value * target_pct
                pct_actuel = (val_reelle / total_pf_value * 100) if total_pf_value > 0 else 0
                
                nom_cat = f"{category} (Cible {target_pct*100:.0f}%)"
                
                # Logique d'affichage NO-SELL
                if gap > 1:
                    # C'est un manque -> À INVESTIR
                    is_priority = (category == priority_cat)
                    prefix = "🌟 " if is_priority else ""
                    delta_txt = f"Investir {gap:,.0f} €"
                    
                    st.metric(
                        label=f"{prefix}{nom_cat}",
                        value=f"{pct_actuel:.1f} %",
                        delta=delta_txt,
                        delta_color="normal" # Vert
                    )
                    if is_priority:
                        st.caption("👆 **C'est ici qu'il faut investir en priorité.**")
                        
                else:
                    # C'est un surplus -> ON GARDE (Pas de vente)
                    st.metric(
                        label=nom_cat,
                        value=f"{pct_actuel:.1f} %",
                        delta=f"Surplus (+{abs(gap):,.0f} €) - Garder",
                        delta_color="off" # Gris (Neutre)
                    )
                
                # Barre de progression (Ratio Actuel / Cible)
                # On cape à 100% visuellement si ça dépasse
                ratio = val_reelle / val_theorique if val_theorique > 0 else 0
                st.progress(min(ratio, 1.0))
                st.write("")

    if len(tickers) > 1 and not df_details.empty:
        st.markdown("---")
        returns = raw_prices.pct_change().dropna() 
        if not returns.empty:
            corr_matrix = returns.corr()
            current_vals = df_details.set_index('Ticker')['Valeur']
            total_pf_val = current_vals.sum()
            weights = (current_vals / total_pf_val).reindex(corr_matrix.index).fillna(0)
            
            weighted_sum = 0; weights_prod_sum = 0
            for i in range(len(corr_matrix)):
                for j in range(len(corr_matrix)):
                    if i != j:
                        w_i = weights.iloc[i]; w_j = weights.iloc[j]; rho = corr_matrix.iloc[i, j]
                        weighted_sum += w_i * w_j * rho; weights_prod_sum += w_i * w_j
            
            avg_corr_weighted = weighted_sum / weights_prod_sum if weights_prod_sum > 0 else 0
            
            sc1, sc2 = st.columns([1, 2])
            with sc1:
                st.metric("Score Diversification", f"{avg_corr_weighted:.2f}")
                if avg_corr_weighted > 0.7: st.error("Faible ⚠️")
                elif avg_corr_weighted > 0.4: st.warning("Moyenne 😐")
                else: st.success("Excellente ✅")
            with sc2:
                with st.expander("Voir Matrice Corrélation"):
                    fig_corr = px.imshow(corr_matrix, text_auto=".2f", color_continuous_scale="RdBu_r", zmin=-1, zmax=1)
                    st.plotly_chart(fig_corr, use_container_width=True)

    st.markdown("---")
    st.subheader("Détail des positions")
    if not df_details.empty:
        st.dataframe(
            df_details.sort_values("Valeur", ascending=False).style
            .format({
                "PRU": "{:.2f} €", "Cours": "{:.2f} €", 
                "Var 24h %": "{:+.2f} %", "Var 24h €": "{:+.2f} €",
                "Valeur": "{:.2f} €", "Plus-Value": "{:+.2f} €", "Rendement %": "{:+.2f} %", "Qté": "{:.2f}"
            })
            .map(lambda x: f'color: {"#28a745" if x >= 0 else "#dc3545"}', subset=['Plus-Value', 'Rendement %', 'Var 24h %', 'Var 24h €']),
            use_container_width=True, hide_index=True
        )
        
    st.markdown("---")
    st.subheader("🔬 Santé Financière & Analyse Fondamentale")
    
    # --- PRÉPARATION TABLEAU 2 (Ajouté) ---
    if not df_fundamentals.empty and not df_details.empty:
        # Filtrer pour ne garder que les lignes actives
        df_fund_active = df_fundamentals[df_fundamentals['Ticker'].isin(active_tickers_list)].copy()
        
        if not df_fund_active.empty:
            # Mapper les noms
            df_fund_active["Actif"] = df_fund_active["Ticker"].map(lambda t: ticker_names.get(t, t))
            
            # Réordonner les colonnes pour l'affichage
            cols = ["Actif", "Ticker", "Prix Actuel", "BPA (EPS)", "PER", "ROE", "Marge Nette", "Dette/Eq", "Dividende %"]
            df_fund_active = df_fund_active[cols]
            
            st.dataframe(
                df_fund_active,
                column_config={
                    "Prix Actuel": st.column_config.NumberColumn(format="%.2f €"),
                    "BPA (EPS)": st.column_config.NumberColumn(format="%.2f"),
                    "PER": st.column_config.NumberColumn(format="%.2f", help="Cours / BPA"),
                    "ROE": st.column_config.NumberColumn(format="%.2f %%"),
                    "Marge Nette": st.column_config.NumberColumn("Marge Nette", format="%.2f %%"),
                    "Dette/Eq": st.column_config.NumberColumn(format="%.2f"),
                    "Dividende %": st.column_config.NumberColumn(format="%.2f %%"),
                },
                hide_index=True,
                use_container_width=True
            )
        else:
            st.info("Pas de données fondamentales disponibles pour les positions actuelles.")
    else:
        st.info("Chargement des fondamentaux...")

# === ONGLET 3 : SUIVI INVESTISSEMENT (DCA DÉCOMPOSÉ) ===
with tab_suivi:
    OBJECTIF_MENSUEL = 600  # Votre cible

    st.markdown("### 🎯 Discipline & Répartition (Flux Net)")
    st.caption(f"Décomposition du flux net mensuel (Achats - Ventes) par type d'actif. Objectif : **{OBJECTIF_MENSUEL} €**.")

    # 1. Traitement des données
    if not df_ops.empty and 'Sens' in df_ops.columns:
        df_flux = df_ops.copy()
        
        # --- A. Sécurisation de la colonne Type_Actif ---
        if 'Type_Actif' not in df_flux.columns:
            df_flux['Type_Actif'] = 'Action' # Valeur par défaut si colonne manquante

        # --- B. Calcul du montant signé (+ Achat / - Vente) ---
        def calculer_flux_signe(row):
            sens = str(row['Sens']).lower()
            montant = row['Montant']
            if 'achat' in sens:
                return montant
            elif 'vente' in sens:
                return -montant
            return 0.0
        
        df_flux['Flux_Net'] = df_flux.apply(calculer_flux_signe, axis=1)

        # --- C. Catégorisation Simplifiée (2 Groupes) ---
        def simplifier_categorie(type_raw):
            t = str(type_raw).upper()
            if any(x in t for x in ['ETF', 'TRACKER', 'ETC', 'OR', 'GOLD', 'MATIÈRE']):
                return 'ETF / ETC'
            return 'Actions' # Tout le reste est considéré comme Action
        
        df_flux['Categorie_Simple'] = df_flux['Type_Actif'].apply(simplifier_categorie)

        # --- D. Agrégation Double : Par Mois ET par Catégorie ---
        # On groupe par Mois (MS) et par Categorie
        df_grouped = df_flux.set_index('Date').groupby([pd.Grouper(freq='MS'), 'Categorie_Simple'])['Flux_Net'].sum().reset_index()

        # --- E. Calcul des Totaux par Mois (pour les KPIs et le texte) ---
        df_totals = df_grouped.groupby('Date')['Flux_Net'].sum().reset_index()
        
        # --- CALCULS KPI DU MOIS EN COURS ---
        now = datetime.now()
        mask_now = (df_totals['Date'].dt.year == now.year) & (df_totals['Date'].dt.month == now.month)
        
        investi_total_ce_mois = 0.0
        if any(mask_now):
            investi_total_ce_mois = df_totals.loc[mask_now, 'Flux_Net'].iloc[0]

        reste_a_investir = OBJECTIF_MENSUEL - investi_total_ce_mois
        pct_completion = min(max(investi_total_ce_mois / OBJECTIF_MENSUEL, 0.0), 1.0)

        # --- AFFICHAGE KPI ---
        c1, c2, c3 = st.columns(3)
        c1.metric("📅 Flux Total ce mois", f"{investi_total_ce_mois:,.2f} €")
        
        if reste_a_investir <= 0:
            c2.metric("✅ Statut", "Objectif Atteint !", f"+{abs(reste_a_investir):.2f} €")
        else:
            c2.metric("⚠️ Reste à investir", f"{reste_a_investir:,.2f} €", delta_color="inverse")
            
        c3.write(f"Progression : {pct_completion*100:.0f}%")
        c3.progress(pct_completion)

        st.markdown("---")

        # --- F. GRAPHIQUE EMPILÉ (STACKED BAR) ---
        st.subheader("📊 Historique par type d'actif")

        fig_dca = go.Figure()

        # 1. Ajout de la couche "Actions"
        df_actions = df_grouped[df_grouped['Categorie_Simple'] == 'Actions']
        fig_dca.add_trace(go.Bar(
            x=df_actions['Date'],
            y=df_actions['Flux_Net'],
            name='Actions',
            marker_color='#ff7f0e', # Orange
            hovertemplate='Actions: %{y:.2f} €<extra></extra>'
        ))

        # 2. Ajout de la couche "ETF / ETC"
        df_etfs = df_grouped[df_grouped['Categorie_Simple'] == 'ETF / ETC']
        fig_dca.add_trace(go.Bar(
            x=df_etfs['Date'],
            y=df_etfs['Flux_Net'],
            name='ETF / ETC',
            marker_color='#1f77b4', # Bleu
            hovertemplate='ETF/ETC: %{y:.2f} €<extra></extra>'
        ))

        # 3. Ajout des étiquettes de texte (Surplus / Manque)
        
        text_labels = []
        text_colors = []
        for val in df_totals['Flux_Net']:
            if val >= OBJECTIF_MENSUEL:
                surplus = val - OBJECTIF_MENSUEL
                text_labels.append(f"+{surplus:.0f}€")
                text_colors.append("#2ca02c") # Vert
            else:
                manque = OBJECTIF_MENSUEL - val
                text_labels.append(f"Reste {manque:.0f}€")
                text_colors.append("#d62728") # Rouge (ou gris sombre)

        fig_dca.add_trace(go.Scatter(
            x=df_totals['Date'],
            y=df_totals['Flux_Net'], # Position verticale = Sommet de la barre
            mode='text',
            text=text_labels,
            textposition='top center',
            textfont=dict(size=11, color=text_colors), # Couleur dynamique du texte
            showlegend=False,
            hoverinfo='skip'
        ))

        # 4. Ligne d'objectif
        fig_dca.add_shape(
            type="line",
            x0=df_totals['Date'].min(), y0=OBJECTIF_MENSUEL,
            x1=df_totals['Date'].max(), y1=OBJECTIF_MENSUEL,
            line=dict(color="gray", width=1, dash="dot"), opacity=0.5
        )

        # Configuration du Layout
        fig_dca.update_layout(
            barmode='relative', # 'relative' gère bien si un mois tu as vendu des actions (négatif) et acheté des ETF (positif)
            title="Répartition de l'effort d'épargne",
            xaxis_title="Mois",
            yaxis_title="Flux Net (€)",
            hovermode="x unified",
            height=500,
            yaxis=dict(tickmode='linear', dtick=200),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
        )
        
        fig_dca.update_xaxes(dtick="M1", tickformat="%b %y")
        st.plotly_chart(fig_dca, use_container_width=True)

    else:
        st.info("Aucune donnée disponible.")
# === ONGLET 4 : PRÉDICTION (MONTE CARLO ROBUSTE) ===
with tab_simu:
    st.markdown("### 🔮 Simulation Monte Carlo")
    st.caption("Projection des scénarios futurs basée sur la volatilité historique de votre portefeuille actuel.")
    
    sim_tickers = list(sim_assets.keys())
    
    if not sim_tickers:
        st.warning("Il faut détenir des actifs valorisés (>0€) pour lancer une simulation.")
    else:
        with st.form("sim_form"):
            c_hor, c_sim = st.columns(2)
            years = c_hor.slider("Horizon (Années)", 1, 40, 10)
            # LIMITE AUGMENTÉE JUSQU'À 100 000
            n_sims = c_sim.number_input("Nombre de simulations", 100, 100000, 1000, step=100)
            run_sim = st.form_submit_button("Lancer la simulation")
        
        if run_sim:
            with st.spinner("Calcul robuste en cours..."):
                # 1. DOWNLOAD
                hist_data = yf.download(sim_tickers, period="2y", progress=False)['Close']
                if isinstance(hist_data, pd.Series):
                    hist_data = hist_data.to_frame()
                    hist_data.columns = sim_tickers
                
                # 2. NETTOYAGE ABSOLU
                hist_data = hist_data.ffill().bfill()
                
                # 3. CALCUL RENDEMENTS
                returns = hist_data.pct_change().dropna()
                returns = returns.dropna(axis=1)
                
                valid_tickers = returns.columns.tolist()
                
                if not valid_tickers:
                    st.error("Impossible de récupérer un historique commun fiable pour vos actifs.")
                else:
                    valid_weights = []
                    for t in valid_tickers:
                        valid_weights.append(sim_assets.get(t, 0))
                    
                    weights_arr = np.array(valid_weights)
                    if weights_arr.sum() > 0:
                        weights_arr = weights_arr / weights_arr.sum()
                    
                    mean_returns = returns.mean()
                    cov_matrix = returns.cov()
                    
                    # --- ASTUCE MATHÉMATIQUE : REGULARISATION ---
                    cov_matrix = cov_matrix + np.eye(len(cov_matrix)) * 1e-6
                    
                    # 4. SIMULATION
                    days = years * 252
                    
                    try:
                        L = np.linalg.cholesky(cov_matrix)
                    except np.linalg.LinAlgError:
                        st.error("Erreur mathématique persistante : Vos actifs sont trop corrélés ou l'historique est insuffisant.")
                        st.stop()

                    portfolio_sims = np.zeros((days, n_sims))
                    initial_portfolio_value = val_bourse_now 
                    
                    for i in range(n_sims):
                        daily_shocks = np.random.normal(size=(days, len(valid_tickers)))
                        correlated_shocks = daily_shocks @ L.T
                        
                        daily_factors = (1 + mean_returns.values + correlated_shocks)
                        
                        # Calcul vectorisé
                        weighted_returns = (daily_factors * weights_arr).sum(axis=1)
                        portfolio_path = initial_portfolio_value * np.cumprod(weighted_returns)
                        
                        portfolio_sims[:, i] = portfolio_path

                    # 5. GRAPHIQUE
                    median_path = np.median(portfolio_sims, axis=1)
                    p1_path = np.percentile(portfolio_sims, 1, axis=1)   # VaR 1%
                    p10_path = np.percentile(portfolio_sims, 10, axis=1) # Pessimiste
                    p25_path = np.percentile(portfolio_sims, 25, axis=1) # Q1
                    p75_path = np.percentile(portfolio_sims, 75, axis=1) # Q3
                    p90_path = np.percentile(portfolio_sims, 90, axis=1) # Optimiste
                    p99_path = np.percentile(portfolio_sims, 99, axis=1) # Top 1%
                    
                    dates_future = pd.date_range(start=datetime.now(), periods=days, freq='B')
                    
                    fig_mc = go.Figure()
                    
                    # Zone 80% (P10-P90)
                    fig_mc.add_trace(go.Scatter(
                        x=np.concatenate([dates_future, dates_future[::-1]]),
                        y=np.concatenate([p90_path, p10_path[::-1]]),
                        fill='toself', fillcolor='rgba(0,100,80,0.2)',
                        line=dict(color='rgba(255,255,255,0)'),
                        name='Intervalle 80%'
                    ))
                    # Médiane
                    fig_mc.add_trace(go.Scatter(
                        x=dates_future, y=median_path,
                        mode='lines', name='Scénario Médian',
                        line=dict(color='#1f77b4', width=2)
                    ))
                    
                    fig_mc.update_layout(title=f"Projection {years} ans (Méthode Standard Corrigée)", yaxis_title="Valeur Portefeuille (€)")
                    st.plotly_chart(fig_mc, use_container_width=True)
                    
                    # 6. MÉTRIQUES ÉTENDUES
                    final_median = median_path[-1]
                    final_p1 = p1_path[-1]
                    final_p10 = p10_path[-1]
                    final_p25 = p25_path[-1]
                    final_p75 = p75_path[-1]
                    final_p90 = p90_path[-1]
                    final_p99 = p99_path[-1]
                    
                    st.markdown("### 📊 Statistiques Détaillées (Fin de période)")
                    
                    c1, c2, c3, c4 = st.columns(4)
                    c1.metric("VaR 1% (Crash)", f"{final_p1:,.0f} €", f"{(final_p1/initial_portfolio_value-1)*100:+.0f}%")
                    c2.metric("Pessimiste (10%)", f"{final_p10:,.0f} €", f"{(final_p10/initial_portfolio_value-1)*100:+.0f}%")
                    c3.metric("1er Quartile (25%)", f"{final_p25:,.0f} €", f"{(final_p25/initial_portfolio_value-1)*100:+.0f}%")
                    c4.metric("Médiane (50%)", f"{final_median:,.0f} €", f"{(final_median/initial_portfolio_value-1)*100:+.0f}%")
                    
                    c5, c6, c7, c8 = st.columns(4)
                    c5.metric("3ème Quartile (75%)", f"{final_p75:,.0f} €", f"{(final_p75/initial_portfolio_value-1)*100:+.0f}%")
                    c6.metric("Optimiste (90%)", f"{final_p90:,.0f} €", f"{(final_p90/initial_portfolio_value-1)*100:+.0f}%")
                    c7.metric("Top 1% (99%)", f"{final_p99:,.0f} €", f"{(final_p99/initial_portfolio_value-1)*100:+.0f}%")
                    c8.write("") 

if st.button('🔄 Actualiser'):
    st.cache_data.clear()
    st.rerun()