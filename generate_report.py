"""
generate_report.py
-------------------
Génère rapport_patrimoine.html à partir des fichiers Excel (Bourse + Banque).
Intègre la fiscalité par enveloppe (CTO / PEA / PEA-PME, taux LFSS 2026), l'attribution de perf, le look-through sectoriel et un système d'alerte.
"""
import os
import re
import argparse
import math
import warnings
import json
import email.utils
import unicodedata
import html as html_lib
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
import numpy as np
import pandas as pd
import yfinance as yf
from datetime import datetime, timedelta


import report_builder as rb

try:
    from arch import arch_model
    HAS_ARCH = True
except ImportError:
    HAS_ARCH = False

# ---------------------------------------------------------------
# CONFIG 
# ---------------------------------------------------------------
# Dossier des fichiers d'entrée (option --data-dir). Par défaut, relatif au SCRIPT et non au dossier
# courant (lancement depuis Spyder ou un autre dossier) : data/demo s'il existe, sinon le dossier du script.
_DOSSIER_SCRIPT = os.path.dirname(os.path.abspath(__file__))
_DOSSIER_DEMO = os.path.join(_DOSSIER_SCRIPT, "data", "demo")
DATA_DIR = _DOSSIER_DEMO if os.path.isdir(_DOSSIER_DEMO) else _DOSSIER_SCRIPT
FILE_BOURSE = "Stratégie et Feuille de Route_Bourse"   # .xlsm ou .xlsx
BASE_NAME_BANQUE = "Comptes"
SHEET_BOURSE = "Historique_Ordres"
CIBLE_MENSUELLE = 500
TAUX_SANS_RISQUE = 0.039
PFU_RATE = 0.314              # PFU LFSS 2026 : 12,8 % IR + 18,6 % PS (CTO)
PEA_RATE = 0.186              # Prélèvements sociaux seuls (PEA / PEA-PME de plus de 5 ans)
PEA_DATE_OUVERTURE = None     # "AAAA-MM-JJ" ; None = inconnue -> PFU appliqué au PEA par prudence
PEA_PME_DATE_OUVERTURE = None # Idem pour le PEA-PME (date d'ouverture propre à ce plan)
MC_YEARS = 10
MC_N_SIMS = 2000
MC_SEED = 42                  
MC_BLOCK_SIZE = 5             
MC_MIN_OBS_GARCH = 500        
OUTPUT_FILE = "rapport_patrimoine.html"
DEVISE_BASE = "EUR"
TICKER_BENCHMARK = "CW8.PA"
MIN_HISTORY_DAYS = 365 * 5 + 30

# --- CONFIGURATION : ANALYSE MARCHÉ ---
FILE_WATCHLIST = "Stock_list"
# Suffixe de place ajouté aux tickers sans suffixe, UNIQUEMENT si la colonne « Suffixe » est
# absente du fichier watchlist. None = aucun ajout (tickers pris tels quels).
WATCHLIST_DEFAULT_SUFFIX = ".PA"
CACHE_FUNDAMENTALS = "cache_fundamentals.json"
CACHE_FUNDAMENTALS_TTL_H = 24
VIX_TICKER = "^VIX"
VIX_TERM_TICKERS = ["^VIX9D", "^VIX3M"]
VIX_SEUILS = {"complaisance": 15, "normal": 20, "eleve": 30}

# --- CONFIGURATION : SURFACE DE VOLATILITÉ IMPLICITE (options S&P 500, sous-jacent du VIX) ---
VOL_SURFACE_TICKER = "^SPX"
VOL_SURFACE_MATURITES_CIBLES = [7, 14, 30, 60, 90, 120, 180, 270, 365, 540, 730]   # en jours
VOL_SURFACE_GRILLE_MONEYNESS = (0.80, 1.20, 0.025)   # K / Forward affiché en 3D : min, max, pas
VOL_SURFACE_TAUX = 0.04       # taux USD d'actualisation des options (méthode CBOE), à ajuster
VOL_SURFACE_CARRY = 0.03      # repli : r - q (forward = spot x exp(carry x T))
CACHE_VOL_SURFACE = "cache_vol_surface.json"
CACHE_VOL_SURFACE_TTL_H = 6

# --- CONFIGURATION : ACTUALITÉS DU PORTEFEUILLE ---
NEWS_JOURS = 14                 # profondeur du fil
NEWS_CACHE = "cache_news.json"
NEWS_CACHE_TTL_MIN = 60
NEWS_MAX_PAR_THEME = 15
NEWS_SEUIL_SIGMA = 2.0          # réaction de marché jugée significative au-delà de 2 écarts-types
NEWS_AGENDA_JOURS = 60
NEWS_FACTEUR_MENTION_INDIRECTE = 0.35   # article du flux d'une ligne qui ne cite pas la société (Boeing dans le flux Airbus…)
# Thèmes suivis via Google Actualités pour l'exposition qu'aucun flux par ticker ne couvre (ETF).
# "tickers" = lignes dont le poids donne l'exposition du thème ; "*" = toute la poche actions.
NEWS_THEMES = {
    "Actions mondiales": {"requete": '"MSCI World" OR "marchés mondiaux" OR "Wall Street"',
                          "tickers": ["WPEA.PA", "EWLD.PA", "VWCE.AS", "CW8.PA"]},
    "Marchés émergents": {"requete": '"marchés émergents" OR "pays émergents" bourse', "tickers": ["AEME.PA"]},
    "Or": {"requete": '"cours de l\'or" OR "prix de l\'or"', "tickers": ["GOLD-EUR.PA"]},
    "Banques centrales": {"requete": 'BCE OR "Réserve fédérale" taux directeurs', "tickers": "*"},
}

# --- CONFIGURATION : SIMULATEUR DE COUVERTURE (CFD vendeur sur indice, aucune position réelle) ---
# dividende : rendement annuel versé par le vendeur d'un CFD sur indice de PRIX (0 pour un ETF
# capitalisant, dont le cours inclut déjà les dividendes). Valeurs approximatives, à ajuster.
HEDGE_INSTRUMENTS = {
    "^STOXX50E": {"label": "Euro Stoxx 50", "devise": "EUR", "dividende": 0.030},
    "^FCHI":     {"label": "CAC 40", "devise": "EUR", "dividende": 0.030},
    "^GSPC":     {"label": "S&P 500", "devise": "USD", "dividende": 0.013},
    "CW8.PA":    {"label": "MSCI World (ETF CW8, capitalisant)", "devise": "EUR", "dividende": 0.0},
}
HEDGE_LEVIERS = [5, 10, 20]        # ESMA : 20:1 maximum sur les grands indices pour un particulier
HEDGE_TAUX_REFERENCE = 0.02        # taux overnight de référence (€STR / SOFR), à ajuster
HEDGE_SPREAD_FINANCEMENT = 0.025   # marge du broker sur le financement overnight, à ajuster


# =================================================================
# 1. CHARGEMENT DES DONNÉES
# =================================================================

def trouver_fichier(base_name, extensions):
    """Premier fichier existant parmi base_name + extension, dans l'ordre donné ; None sinon."""
    return next((base_name + ext for ext in extensions if os.path.exists(base_name + ext)), None)


def load_stock_data(base_name):
    file_path = trouver_fichier(base_name, (".xlsm", ".xlsx"))
    if file_path is None:
        return None
    try:
        temp_df = pd.read_excel(file_path, sheet_name=SHEET_BOURSE, header=None)
    except Exception:
        return "ERROR_SHEET"

    header_row = -1
    for i, row in temp_df.iterrows():
        row_str = [str(val).strip() for val in row.values]
        if "Date exécution Ordres" in row_str:
            header_row = i
            break
    if header_row == -1:
        return "ERROR_HEADER"

    df = pd.read_excel(file_path, sheet_name=SHEET_BOURSE, header=header_row)
    df.columns = df.columns.str.strip()

    ticker_col = next((c for c in df.columns if c.lower() in ['ticker', 'symbole', 'code']), None)
    compte_col = next((c for c in df.columns if 'compte' in str(c).lower()), None)
    
    if not ticker_col:
        return "ERROR_NO_TICKER"

    rename_dict = {
        "Date exécution Ordres": "Date", "Actif": "Nom", "Type d'ordre": "Sens",
        "Nombre d'ordre": "Qte", "Valeur total Ordres + frais": "Montant", ticker_col: "Ticker"
    }
    if compte_col:
        rename_dict[compte_col] = "Compte"

    df = df.rename(columns=rename_dict)
    
    if "Compte" not in df.columns:
        df["Compte"] = "CTO"

    df['Date'] = pd.to_datetime(df['Date'], errors='coerce').dt.normalize()
    df['Ticker'] = df['Ticker'].fillna('').astype(str).str.strip().str.upper()
    mask_empty = df['Ticker'].isin(['NAN', 'NONE', '', 'NATYPE', '<NA>'])
    df.loc[mask_empty, 'Ticker'] = 'CASH_OP'

    for col in ['Qte', 'Montant']:
        if col in df.columns:
            if df[col].dtype == 'object':
                df[col] = (df[col].astype(str).str.replace('€', '', regex=False)
                           .str.replace(r'\s+', '', regex=True).str.replace(',', '.', regex=False))
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0)

    return df


def load_bank_data(base_name):
    file_path = trouver_fichier(base_name, (".xlsx", ".xls"))
    if file_path is None:
        return "NOT_FOUND"

    try:
        temp_df = pd.read_excel(file_path, header=None)
        header_row = -1
        for i, row in temp_df.iterrows():
            row_str = [str(val).strip().lower() for val in row.values]
            if "compte" in row_str and "montant" in row_str:
                header_row = i
                break
        if header_row == -1:
            return "ERROR_COLUMNS"

        df = pd.read_excel(file_path, header=header_row)
        df.columns = df.columns.str.strip()

        col_date = next((c for c in df.columns if c.lower() == 'date'), None)
        col_compte = next((c for c in df.columns if c.lower() == 'compte'), None)
        col_montant = next((c for c in df.columns if c.lower() == 'montant'), None)
        if not col_compte or not col_montant:
            return "ERROR_COLUMNS"

        rename_dict = {col_compte: "Compte", col_montant: "Montant"}
        if col_date:
            rename_dict[col_date] = "Date"
        df = df.rename(columns=rename_dict)
        df = df.dropna(subset=['Compte'])
        if "Unnamed: 0" in df.columns:
            del df["Unnamed: 0"]

        if df['Montant'].dtype == 'object':
            df['Montant'] = df['Montant'].astype(str).str.replace('€', '').str.replace(' ', '').str.replace(',', '.')
        df['Montant'] = pd.to_numeric(df['Montant'], errors='coerce').fillna(0)

        # Toutes les dates à minuit (y compris la date du jour de repli) : sinon elles ne
        # correspondent pas au calendrier journalier `all_days` et les soldes sont perdus au reindex.
        today = pd.Timestamp.now().normalize()
        if "Date" in df.columns:
            df['Date'] = pd.to_datetime(df['Date'], dayfirst=True, errors='coerce').dt.normalize()
            df['Date'] = df['Date'].ffill().fillna(today) if not df['Date'].isna().all() else today
        else:
            df['Date'] = today
        return df
    except Exception as e:
        return f"ERROR_READ: {str(e)}"


def load_watchlist(base_name):
    file_path = trouver_fichier(base_name, (".xlsx", ".csv"))

    if not file_path:
        return []
        
    try:
        if file_path.endswith('.csv'):
            df = pd.read_csv(file_path, sep=None, engine='python') 
        else:
            df = pd.read_excel(file_path)
            
        df.columns = [str(c).strip() for c in df.columns]
        if 'Symbol' not in df.columns:
            return []

        col_suffixe = next((c for c in df.columns if c.lower() == 'suffixe'), None)
        df = df.dropna(subset=['Symbol'])

        tickers = []
        for _, row in df.iterrows():
            t = str(row['Symbol']).strip().upper()
            if not t or t == 'NAN':
                continue
            if col_suffixe:
                # Colonne Suffixe présente : elle fait foi, ligne par ligne (cellule vide = aucun suffixe)
                suffixe = row[col_suffixe]
                suffixe = '' if pd.isna(suffixe) else str(suffixe).strip().upper()
                if suffixe and not suffixe.startswith('.'):
                    suffixe = '.' + suffixe
                if suffixe and not t.endswith(suffixe):
                    t = t + suffixe
            elif WATCHLIST_DEFAULT_SUFFIX and "." not in t:
                # Comportement historique, désactivable via WATCHLIST_DEFAULT_SUFFIX = None
                t = t + WATCHLIST_DEFAULT_SUFFIX
            tickers.append(t)

        return list(set(tickers))
    except Exception:
        return []


def _charger_cache_fondamentaux(path=CACHE_FUNDAMENTALS):
    """Cache JSON {ticker: {"timestamp": ISO, "info": dict yfinance}} ; vide si absent ou illisible."""
    if not os.path.exists(path):
        return {}
    try:
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except (OSError, ValueError):
        return {}


def _sauver_cache_fondamentaux(cache, path=CACHE_FUNDAMENTALS):
    """Écriture atomique (fichier temporaire puis remplacement) pour ne jamais corrompre le cache."""
    tmp = path + ".tmp"
    try:
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(cache, f, ensure_ascii=False, default=str)
        os.replace(tmp, path)
    except OSError:
        pass


def _info_valide(info):
    """yfinance renvoie parfois un dict quasi vide au lieu de lever une exception."""
    return isinstance(info, dict) and any(
        info.get(k) is not None for k in ('currency', 'quoteType', 'regularMarketPrice', 'previousClose'))


def _info_ticker(t, cache, now):
    """Fondamentaux d'un ticker : cache s'il a moins de CACHE_FUNDAMENTALS_TTL_H heures, sinon réseau ;
    en cas d'échec réseau, cache même périmé plutôt que rien.
    Retourne (info | None, source) avec source parmi 'reseau', 'cache', 'cache_perime', None."""
    entry = cache.get(t)
    age_h = None
    if isinstance(entry, dict) and 'timestamp' in entry and isinstance(entry.get('info'), dict):
        try:
            age_h = (now - datetime.fromisoformat(entry['timestamp'])).total_seconds() / 3600
        except (TypeError, ValueError):
            entry = None
    else:
        entry = None

    if entry is not None and 0 <= age_h < CACHE_FUNDAMENTALS_TTL_H:
        return entry['info'], 'cache'
    try:
        info = yf.Ticker(t).info
    except Exception:
        info = None
    if _info_valide(info):
        cache[t] = {"timestamp": now.isoformat(timespec='seconds'), "info": info}
        return info, 'reseau'
    if entry is not None:
        return entry['info'], 'cache_perime'
    return None, None


def get_fundamentals_and_currency(tickers):
    data = []
    currency_map = {}
    api_tickers = [t for t in tickers if t != 'CASH_OP' and not str(t).startswith('^')]
    if not api_tickers:
        return pd.DataFrame(), {}

    cache = _charger_cache_fondamentaux()
    now = datetime.now()
    sources = {}
    for t in api_tickers:
        info, source = _info_ticker(t, cache, now)
        sources[source] = sources.get(source, 0) + 1
        if info is None:
            currency_map[t] = DEVISE_BASE
            continue
        currency_map[t] = info.get('currency') or DEVISE_BASE
        try:
            price = info.get('currentPrice') or info.get('regularMarketPrice') or info.get('previousClose')
            prev_close = info.get('previousClose')

            # Variation « live » seulement si l'info vient d'être téléchargée : depuis le cache elle
            # peut dater de la veille. À None, run_engine retombe sur les deux derniers cours de
            # yf.download (rafraîchis à chaque exécution).
            var_live_pct = None
            if source == 'reseau' and price and prev_close and prev_close > 0:
                var_live_pct = (price / prev_close - 1) * 100

            eps = info.get('trailingEps')
            per = price / eps if price and eps else None
            roe = info.get('returnOnEquity')
            margin = info.get('profitMargins')
            debt_eq = info.get('debtToEquity')
            div = info.get('dividendYield')
            rev_growth = info.get('revenueGrowth')
            sector = info.get('sector', 'Inconnu')
            
            data.append({
                "Ticker": t, "Prix Actuel": price, "BPA (EPS)": eps, "PER": per,
                "ROE": roe * 100 if roe else None, "Marge Nette": margin,
                "Dette/Eq": debt_eq / 100 if debt_eq else None, 
                "Croissance CA": rev_growth * 100 if rev_growth else None,
                "Dividende %": div,
                "Var_Live_Pct": var_live_pct,
                "Secteur": sector,
                # Événements (timestamps Unix Yahoo) pour l'agenda de l'onglet Actualités
                "Type": info.get('quoteType'),
                "Nom_Long": info.get('longName') or info.get('shortName'),
                "Prochains_Resultats": info.get('earningsTimestampStart') or info.get('earningsTimestamp'),
                "Ex_Dividende": info.get('exDividendDate'),
            })
        except Exception:
            pass

    if sources.get('reseau'):
        _sauver_cache_fondamentaux(cache)
    print(f"Fondamentaux : {sources.get('reseau', 0)} téléchargés, {sources.get('cache', 0)} depuis le cache, "
          f"{sources.get('cache_perime', 0)} depuis le cache périmé (échec réseau), {sources.get(None, 0)} indisponibles.")
    return pd.DataFrame(data), currency_map


def get_fx_rates(currencies, start_date, end_date, all_days):
    fx_series = {DEVISE_BASE: pd.Series(1.0, index=all_days)}
    other_currencies = sorted({c for c in currencies if c and c != DEVISE_BASE})
    if not other_currencies:
        return fx_series
    pairs = {c: f"{c}{DEVISE_BASE}=X" for c in other_currencies}
    
    try:
        raw_fx = yf.download(list(pairs.values()), start=start_date,
                              end=end_date + timedelta(days=1), progress=False)['Close']
        if isinstance(raw_fx, pd.Series):
            raw_fx = raw_fx.to_frame(); raw_fx.columns = list(pairs.values())
    except Exception:
        raw_fx = pd.DataFrame()

    for c, pair in pairs.items():
        if pair in raw_fx.columns and not raw_fx[pair].dropna().empty:
            s = raw_fx[pair].reindex(all_days).ffill().bfill()
        else:
            s = pd.Series(1.0, index=all_days)
        fx_series[c] = s
    return fx_series


def get_market_data(tickers, start_date, end_date_limit, currency_map):
    if not tickers:
        return None, None, None
    api_tickers = [t for t in tickers if t != 'CASH_OP']
    
    wide_start = pd.Timestamp(min(start_date, datetime.now() - timedelta(days=MIN_HISTORY_DAYS))).normalize()
    final_date = pd.Timestamp(max(datetime.now(), end_date_limit) + timedelta(days=1)).normalize()
    all_days = pd.date_range(start=wide_start, end=final_date, freq='D')

    try:
        raw = yf.download(api_tickers, start=wide_start, end=final_date, progress=False)
        if raw.empty:
            return None, None, None
        download = raw['Close']
        if isinstance(download, pd.Series):
            download = download.to_frame(); download.columns = api_tickers
    except Exception:
        return None, None, None

    fx_series = get_fx_rates(set(currency_map.values()), wide_start, final_date, all_days)
    
    converted_dict = {}
    for t in download.columns:
        # Indices sans devise connue (VIX…) : niveau brut. Les indices de couverture ont une
        # devise déclarée dans HEDGE_INSTRUMENTS et sont convertis en EUR comme les titres.
        if str(t).startswith('^') and t not in currency_map:
            converted_dict[t] = download[t]
        else:
            cur = currency_map.get(t, DEVISE_BASE)
            rate = fx_series.get(cur, fx_series[DEVISE_BASE]).reindex(download.index).ffill().bfill()
            converted_dict[t] = download[t] * rate

    raw_converted = pd.DataFrame(converted_dict, index=download.index)
    raw_prices = raw_converted.dropna(how='all')
    # `prices` (calendrier continu, ffill + bfill) ne sert QU'À VALORISER des titres déjà détenus :
    # le bfill évite une valeur nulle si un ordre est daté avant la première cotation disponible.
    # Il NE DOIT PAS être réutilisé pour calculer des rendements : le bfill recopie le premier cours
    # connu vers le passé et fabrique des rendements nuls artificiels (volatilité et corrélations
    # sous-estimées). Pour les rendements, partir de `raw_prices` via `rendements_reels`.
    prices = raw_prices.reindex(all_days).ffill().bfill()
    
    return prices, all_days, raw_prices


# =================================================================
# 1bis. FISCALITÉ PAR ENVELOPPE
# =================================================================

def classer_enveloppe(compte):
    """Ramène le libellé de compte à une enveloppe fiscale : 'PEA-PME', 'PEA' ou 'CTO'.
    Le test PEA-PME passe en premier car un simple contains('PEA') le confondrait avec le PEA.
    Tout compte non reconnu est traité comme un CTO (PFU), comme auparavant."""
    c = str(compte).strip().upper().replace(' ', '').replace('_', '-')
    if 'PEA-PME' in c or 'PEAPME' in c:
        return 'PEA-PME'
    if 'PEA' in c:
        return 'PEA'
    return 'CTO'


def taux_plan_epargne(date_ouverture, date_rapport):
    """Taux applicable à un PEA / PEA-PME : PS seuls si le plan a 5 ans ou plus,
    sinon PFU (retrait avant 5 ans). Date inconnue ou illisible -> PFU par prudence.
    Retourne (taux, plus_de_5_ans: bool|None, date_ouverture: Timestamp|None)."""
    if not date_ouverture:
        return PFU_RATE, None, None
    try:
        d_ouv = pd.Timestamp(date_ouverture).normalize()
    except (ValueError, TypeError):
        return PFU_RATE, None, None
    plus_de_5_ans = d_ouv + pd.DateOffset(years=5) <= pd.Timestamp(date_rapport).normalize()
    return (PEA_RATE if plus_de_5_ans else PFU_RATE), plus_de_5_ans, d_ouv


def calcul_fiscalite_enveloppes(df_details, date_rapport):
    """Impôt latent par enveloppe : max(0, somme des plus-values latentes) x taux.
    Les moins-values d'une enveloppe compensent ses plus-values. L'impôt de l'enveloppe est
    ensuite réparti sur les lignes au prorata de leurs plus-values positives, de sorte que la
    somme de la colonne 'PFU' égale exactement l'impôt total des enveloppes.
    Retourne (df_details enrichi, total_pfu, dict d'infos par enveloppe pour l'affichage)."""
    df = df_details.copy()
    df['Enveloppe'] = df['Compte'].apply(classer_enveloppe)
    df['PFU'] = 0.0

    dates_ouverture = {'PEA': PEA_DATE_OUVERTURE, 'PEA-PME': PEA_PME_DATE_OUVERTURE}
    infos = {}
    for env, grp in df.groupby('Enveloppe'):
        if env == 'CTO':
            taux, plus_5_ans, d_ouv = PFU_RATE, None, None
        else:
            taux, plus_5_ans, d_ouv = taux_plan_epargne(dates_ouverture[env], date_rapport)
        pv_nette = float(grp['Plus-Value'].sum())
        impot = max(0.0, pv_nette) * taux
        pv_pos = grp['Plus-Value'].clip(lower=0)
        if impot > 0 and pv_pos.sum() > 0:
            df.loc[grp.index, 'PFU'] = impot * pv_pos / pv_pos.sum()
        infos[env] = {
            "taux": taux, "pv_nette": pv_nette, "impot": impot,
            "plus_de_5_ans": plus_5_ans, "date_ouverture": d_ouv,
            "pfu_applique": env != 'CTO' and taux == PFU_RATE,
        }

    df['Taux_Fiscal'] = df['Enveloppe'].map({e: i['taux'] for e, i in infos.items()})
    df['Valeur_Nette'] = df['Valeur'] - df['PFU']
    total_pfu = float(sum(i['impot'] for i in infos.values()))
    return df, total_pfu, infos


# =================================================================
# 1ter. RENDEMENTS SUR HISTORIQUE RÉEL (sans bfill)
# =================================================================

def rendements_reels(price_series, ffill_limit=5):
    """Rendements journaliers d'un actif calculés uniquement sur sa période de cotation réelle.
    - Avant la première cotation et après la dernière : aucun rendement (pas de bfill, donc
      pas de rendement nul inventé avant l'introduction en bourse).
    - À l'intérieur de la période : ffill limité à `ffill_limit` jours (jours fériés d'une
      place quand une autre cote) ; None = ffill sans limite à l'intérieur de la période.
    - Un trou plus long que la limite ne produit pas de rendement (NaN supprimé)."""
    s = price_series.dropna()
    if s.empty:
        return pd.Series(dtype=float)
    s = price_series.loc[s.index[0]:s.index[-1]].ffill(limit=ffill_limit)
    return s.pct_change(fill_method=None).dropna()


def compute_twr(df_hist_bourse):
    """Rendement journalier time-weighted du compte bourse (neutralise dépôts et retraits).
    Base du jour = valeur de la veille + flux du jour ; rendement = valeur / base - 1.
    Colonnes ajoutées : CF, Prev_Val, Base, Daily_Ret, Gross_Ret, TWR_Index (indice cumulé, base 1),
    Year, Month."""
    df_perf = df_hist_bourse[['Valeur_Bourse', 'Investi_Bourse']].copy()
    if df_perf.empty:
        for col in ['CF', 'Prev_Val', 'Base', 'Daily_Ret', 'Gross_Ret', 'TWR_Index', 'Year', 'Month']:
            df_perf[col] = pd.Series(dtype=float)
        return df_perf
    df_perf['CF'] = df_perf['Investi_Bourse'].diff().fillna(df_perf['Investi_Bourse'].iloc[0])
    df_perf['Prev_Val'] = df_perf['Valeur_Bourse'].shift(1).fillna(0)
    df_perf['Base'] = df_perf['Prev_Val'] + df_perf['CF']

    df_perf['Daily_Ret'] = 0.0
    mask = df_perf['Base'] > 0
    df_perf.loc[mask, 'Daily_Ret'] = (df_perf.loc[mask, 'Valeur_Bourse'] - df_perf.loc[mask, 'Base']) / df_perf.loc[mask, 'Base']
    df_perf['Gross_Ret'] = 1 + df_perf['Daily_Ret']
    df_perf['TWR_Index'] = df_perf['Gross_Ret'].cumprod()
    df_perf['Year'] = df_perf.index.year
    df_perf['Month'] = df_perf.index.month
    return df_perf


def drawdown_courant_twr(df_twr):
    """Drawdown actuel (en %) : indice TWR cumulé / son plus-haut historique - 1."""
    twr_index = df_twr['TWR_Index']
    return float((twr_index.iloc[-1] / twr_index.cummax().iloc[-1] - 1) * 100)


def note_lignes_exclues(exclues, sim_assets, nb_rendements, minimum):
    """Texte expliquant les lignes écartées d'un calcul faute d'historique de cours suffisant."""
    if not exclues:
        return ""
    total = sum(sim_assets.values())
    detail = ", ".join(
        f"{t} ({nb_rendements.get(t, 0)} j, {sim_assets.get(t, 0) / total * 100:.1f} % du portefeuille)"
        for t in exclues) if total > 0 else ", ".join(exclues)
    return (f"Lignes exclues faute d'historique de cours exploitable (minimum {minimum} jours) : {detail}. "
            f"Les poids des autres lignes sont renormalisés.")


def periode_str(index):
    """Période couverte par un index de dates, au format lisible (JJ/MM/AAAA → JJ/MM/AAAA)."""
    if len(index) == 0:
        return "période vide"
    return f"du {index[0]:%d/%m/%Y} au {index[-1]:%d/%m/%Y}"


def valider_retraits(df_ops):
    """Convention : un retrait est saisi avec un montant NÉGATIF (le montant des virements est
    ajouté tel quel aux espèces et aux apports). Une ligne libellée « Retrait » avec un montant
    positif serait comptée comme un dépôt : on la signale sans la corriger.
    Retourne la liste des messages d'alerte (vide si tout est cohérent)."""
    sens = df_ops['Sens'].astype(str).str.strip().str.lower()
    suspectes = df_ops[sens.str.contains('retrait', na=False) & (df_ops['Montant'] > 0)]
    messages = []
    for _, r in suspectes.iterrows():
        date_txt = f"{r['Date']:%d/%m/%Y}" if pd.notna(r['Date']) else "date inconnue"
        compte = str(r.get('Compte', '')).strip()
        compte_txt = f" ({compte})" if compte and compte.upper() not in ('NAN', 'NONE') else ""
        messages.append(
            f"⚠️ Saisie à vérifier : ligne « Retrait » du {date_txt}{compte_txt} avec un montant positif "
            f"({rb.eur(r['Montant'])}). Elle est comptée comme un dépôt ; saisissez les retraits en négatif.")
    return messages


# =================================================================
# 1quater. SURFACE DE VOLATILITÉ IMPLICITE (options S&P 500)
# =================================================================
# Les IV publiées par Yahoo sont inexploitables hors séance US (bid/ask à 0, IV ≈ 0,4 %) : on les
# recalcule en Black-76 à partir des prix, avec un forward et une actualisation implicites
# (parité call-put), sur les seules options hors de la monnaie (les plus liquides).

_ERF = np.frompyfunc(math.erf, 1, 1)


def _norm_cdf(x):
    return 0.5 * (1.0 + _ERF(np.asarray(x, dtype=float) / math.sqrt(2.0)).astype(float))


def prix_black76(F, K, T, DF, sigma, is_call):
    """Prix Black-76 d'une option européenne sur forward (vectorisé)."""
    F, K, T, DF, sigma = (np.asarray(a, dtype=float) for a in (F, K, T, DF, sigma))
    sq = sigma * np.sqrt(T)
    d1 = (np.log(F / K) + 0.5 * sq ** 2) / sq
    d2 = d1 - sq
    call = DF * (F * _norm_cdf(d1) - K * _norm_cdf(d2))
    put = DF * (K * _norm_cdf(-d2) - F * _norm_cdf(-d1))
    return np.where(is_call, call, put)


def vol_implicite_black76(prix, F, K, T, DF, is_call, n_iter=60):
    """Volatilité implicite par dichotomie (le prix croît avec sigma). NaN si le prix viole les
    bornes de non-arbitrage (sous la valeur intrinsèque ou au-dessus du maximum théorique)."""
    prix = np.asarray(prix, dtype=float)
    K = np.asarray(K, dtype=float)
    is_call = np.asarray(is_call, dtype=bool)
    intrinseque = DF * np.where(is_call, np.maximum(F - K, 0.0), np.maximum(K - F, 0.0))
    borne_sup = DF * np.where(is_call, F, K)
    valide = np.isfinite(prix) & (prix > intrinseque + 1e-8) & (prix < borne_sup)
    lo = np.full(prix.shape, 1e-4)
    hi = np.full(prix.shape, 5.0)
    for _ in range(n_iter):
        mid = 0.5 * (lo + hi)
        trop_haut = prix_black76(F, K, T, DF, mid, is_call) > prix
        hi = np.where(trop_haut, mid, hi)
        lo = np.where(trop_haut, lo, mid)
    iv = 0.5 * (lo + hi)
    iv[~valide | (iv > 4.9) | (iv < 2e-4)] = np.nan
    return iv


def _prix_options(df, session):
    """Prix retenu : milieu bid/ask s'il existe ; sinon dernier prix, seulement s'il a été échangé
    pendant la dernière séance (un prix plus ancien n'est plus cohérent avec le sous-jacent)."""
    bid, ask = df['bid'].fillna(0), df['ask'].fillna(0)
    a_mid = (bid > 0) & (ask >= bid)
    derniere = pd.to_datetime(df['lastTradeDate'], utc=True).dt.tz_convert(None).dt.normalize()
    a_last = (~a_mid) & (derniere == session) & (df['lastPrice'] > 0)
    prix = pd.Series(np.nan, index=df.index)
    prix[a_mid] = (bid + ask)[a_mid] / 2
    prix[a_last] = df.loc[a_last, 'lastPrice']
    source = pd.Series(None, index=df.index, dtype=object)
    source[a_mid], source[a_last] = 'mid', 'last'
    return prix, source


def estimer_forward(calls, puts, spot, T):
    """Forward implicite, méthode du CBOE (calcul du VIX) : actualisation au taux de config, puis
    F = K + (C - P) / DF aux strikes où |C - P| est minimal (au voisinage du forward, F y est peu
    sensible au taux). Médiane sur les 5 meilleurs strikes pour résister à un prix aberrant.
    Estimer aussi l'actualisation par régression n'est pas fiable avec des derniers prix non
    synchrones (taux implicites de plusieurs centaines de % observés sur les échéances courtes)."""
    DF = math.exp(-VOL_SURFACE_TAUX * T)
    paires = calls[['strike', 'prix']].merge(puts[['strike', 'prix']], on='strike', suffixes=('_c', '_p')).dropna()
    paires = paires[(paires['strike'] / spot - 1).abs() < 0.10]
    if len(paires) >= 3:
        paires = paires.assign(ecart=(paires['prix_c'] - paires['prix_p']).abs()).nsmallest(5, 'ecart')
        F = float(np.median(paires['strike'] + (paires['prix_c'] - paires['prix_p']) / DF))
        if abs(F / spot - 1) < 0.20:
            return F, float(DF), "parité call-put"
    return float(spot * math.exp(VOL_SURFACE_CARRY * T)), float(DF), "taux de repli (config)"


def smile_echeance(calls, puts, spot, T, session):
    """Smile d'une échéance : IV des options hors de la monnaie (puts sous le forward, calls
    au-dessus), points aberrants retirés. Retourne (DataFrame m/iv/source, F, DF, méthode)."""
    calls, puts = calls.copy(), puts.copy()
    calls['prix'], calls['source'] = _prix_options(calls, session)
    puts['prix'], puts['source'] = _prix_options(puts, session)
    F, DF, methode = estimer_forward(calls, puts, spot, T)

    otm = pd.concat([calls[calls['strike'] >= F].assign(is_call=True),
                     puts[puts['strike'] < F].assign(is_call=False)])
    otm = otm.dropna(subset=['prix'])
    otm = otm[(otm['strike'] / F > 0.5) & (otm['strike'] / F < 1.5)]
    if otm.empty:
        return pd.DataFrame(columns=['m', 'iv', 'source']), F, DF, methode
    otm['iv'] = vol_implicite_black76(otm['prix'].values, F, otm['strike'].values, T, DF, otm['is_call'].values)
    otm['m'] = otm['strike'] / F
    smile = otm.dropna(subset=['iv']).sort_values('m')[['m', 'iv', 'source']].reset_index(drop=True)
    # Prix asynchrones ou périmés : on retire les points trop éloignés de la médiane locale. Le seuil
    # s'adapte au bruit local (MAD des écarts) : strict sur un smile propre, tolérant sur un smile bruité.
    if len(smile) >= 5:
        med = smile['iv'].rolling(7, center=True, min_periods=3).median()
        ecart = (smile['iv'] - med).abs()
        mad_local = ecart.rolling(7, center=True, min_periods=3).median()
        smile = smile[ecart <= np.maximum(0.01, 4 * 1.4826 * mad_local)].reset_index(drop=True)
    return smile, F, DF, methode


def surface_depuis_chaines(chaines, spot, session, min_points=8):
    """Assemble la surface à partir de chaînes [(échéance, calls, puts)]. La grille n'est remplie
    qu'à l'intérieur de la plage de strikes observée pour chaque échéance (aucune extrapolation)."""
    m_min, m_max, pas = VOL_SURFACE_GRILLE_MONEYNESS
    grille = np.round(np.arange(m_min, m_max + pas / 2, pas), 4)
    lignes = []
    for echeance, calls, puts in chaines:
        jours = (pd.Timestamp(echeance).normalize() - session).days
        if jours < 1:
            continue
        T = jours / 365.0
        smile, F, DF, methode = smile_echeance(calls, puts, spot, T, session)
        if len(smile) < min_points:
            continue
        m, iv = smile['m'].values, smile['iv'].values
        dans_plage = (grille >= m.min()) & (grille <= m.max())
        ligne_iv = np.where(dans_plage, np.interp(grille, m, iv), np.nan)
        lignes.append({
            "echeance": f"{pd.Timestamp(echeance):%Y-%m-%d}", "jours": int(jours), "forward": F,
            "actualisation": DF, "methode_forward": methode, "nb_points": int(len(smile)),
            "part_mid": float((smile['source'] == 'mid').mean()),
            "grille_iv": [None if np.isnan(v) else float(v) for v in ligne_iv],
            "smile_m": [float(v) for v in m], "smile_iv": [float(v) for v in iv],
        })
    lignes.sort(key=lambda r: r["jours"])
    return {"spot": float(spot), "session": f"{session:%Y-%m-%d}", "moneyness": grille.tolist(), "echeances": lignes}


def iv_surface(surface, m, jours):
    """IV interpolée en (moneyness, maturité) : linéaire en strike dans chaque smile observé, puis
    linéaire en variance totale (sigma² x T) entre les deux échéances encadrantes. None hors plage."""
    pts = []
    for r in surface.get("echeances", []):
        sm, si = r["smile_m"], r["smile_iv"]
        if sm and sm[0] <= m <= sm[-1]:
            pts.append((r["jours"], float(np.interp(m, sm, si))))
    for (d_a, v_a), (d_b, v_b) in zip(pts, pts[1:]):
        if d_a <= jours <= d_b:
            w = v_a ** 2 * d_a + (v_b ** 2 * d_b - v_a ** 2 * d_a) * (jours - d_a) / (d_b - d_a) if d_b > d_a else v_a ** 2 * d_a
            return math.sqrt(w / jours) if w > 0 else None
    return None


def _choisir_echeances(echeances, date_ref):
    """Une échéance par maturité cible (la plus proche, à ±35 %), sans doublon."""
    dispo = [(e, (pd.Timestamp(e) - date_ref).days) for e in echeances]
    dispo = [(e, d) for e, d in dispo if d >= 5]
    choix = []
    for cible in VOL_SURFACE_MATURITES_CIBLES:
        if not dispo:
            break
        e, d = min(dispo, key=lambda x: abs(x[1] - cible))
        if abs(d - cible) <= max(5, 0.35 * cible) and e not in choix:
            choix.append(e)
    return choix


def get_vol_surface():
    """Télécharge les chaînes d'options et calcule la surface. Cache CACHE_VOL_SURFACE_TTL_H heures ;
    en cas d'échec réseau, cache même périmé (signalé). Retourne None si rien n'est disponible."""
    cache = None
    if os.path.exists(CACHE_VOL_SURFACE):
        try:
            with open(CACHE_VOL_SURFACE, encoding="utf-8") as f:
                cache = json.load(f)
            age_h = (datetime.now() - datetime.fromisoformat(cache["timestamp"])).total_seconds() / 3600
            if 0 <= age_h < CACHE_VOL_SURFACE_TTL_H and cache["surface"].get("echeances"):
                return dict(cache["surface"], source="cache", timestamp=cache["timestamp"])
        except (OSError, ValueError, KeyError, TypeError):
            cache = None

    surface = None
    try:
        tk = yf.Ticker(VOL_SURFACE_TICKER)
        choix = _choisir_echeances(tk.options, pd.Timestamp.now().normalize())
        chaines, spot = [], None
        for e in choix:
            ch = tk.option_chain(e)
            chaines.append((e, ch.calls, ch.puts))
            if spot is None:
                spot = (getattr(ch, 'underlying', None) or {}).get('regularMarketPrice')
        if chaines:
            if not spot:
                spot = float(tk.history(period="5d")['Close'].iloc[-1])
            session = max(pd.to_datetime(df['lastTradeDate'], utc=True).dt.tz_convert(None).max()
                          for _, c, p in chaines for df in (c, p)).normalize()
            surface = surface_depuis_chaines(chaines, float(spot), session)
    except Exception as exc:
        print(f"Surface de volatilité : échec du téléchargement ({exc}).")
        surface = None

    if surface and surface["echeances"]:
        horodatage = datetime.now().isoformat(timespec='seconds')
        tmp = CACHE_VOL_SURFACE + ".tmp"
        try:
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump({"timestamp": horodatage, "surface": surface}, f)
            os.replace(tmp, CACHE_VOL_SURFACE)
        except OSError:
            pass
        return dict(surface, source="reseau", timestamp=horodatage)
    if cache and cache.get("surface", {}).get("echeances"):
        return dict(cache["surface"], source="cache_perime", timestamp=cache["timestamp"])
    return None


# =================================================================
# 1quinquies. ACTUALITÉS DU PORTEFEUILLE (veille)
# =================================================================
# Sources : flux RSS Yahoo Finance par ticker (actions détenues) et Google Actualités par thème
# (exposition des ETF, que les flux par ticker ne couvrent pas). Chaque article est classé, noté
# et rapproché de la réaction du cours le jour de séance qui suit sa publication.

# (type, importance, mots-clés FR/EN en minuscules)
NEWS_TYPES = [
    ("Résultats", 1.0, ["résultat", "chiffre d'affaires", "bénéfice", "trimestre", "semestre", "earnings",
                        "revenue", "quarterly", "profit"]),
    ("Fusions & acquisitions", 1.0, ["rachat de", "acquisition", "acquiert", "fusion", "opa ", "offre publique",
                                     "cession", "cède ", "merger", "takeover", "to acquire"]),
    ("Prévisions", 0.9, ["prévision", "objectifs annuels", "perspectives", "guidance", "outlook",
                         "avertissement sur", "profit warning"]),
    ("Activisme & actionnariat", 0.8, ["activiste", "entre au capital", "franchissement de seuil",
                                       "participation de", "activist", "stake in"]),
    ("Juridique & régulation", 0.8, ["enquête", "amende", "procès", "plainte", "régulateur", "sanction",
                                     "fda", "demande de licence", "autorisation de mise", "lawsuit", "investigation"]),
    ("Analystes", 0.6, ["objectif de cours", "recommandation", "analyste", "surpondérer", "sous-pondérer",
                        "upgrade", "downgrade", "price target"]),
    ("Contrats & commandes", 0.6, ["commande", "contrat", "partenariat", "s'allie", "accord avec", "order for",
                                   "contract", "partnership"]),
    ("Dividendes & rachats d'actions", 0.5, ["dividende", "rachat d'actions", "actions propres", "buyback",
                                            "dividend"]),
]
NEWS_MOTS_POSITIFS = ["hausse", "bondit", "grimpe", "s'envole", "progresse", "record", "relève", "dépasse",
                      "supérieur aux attentes", "succès", "rebond", "surperform", "beat", "upgrade", "surge",
                      "soar", "rally", "raises", "accélère"]
NEWS_MOTS_NEGATIFS = ["baisse", "chute", "recule", "plonge", "dégringole", "abaisse", "avertissement", "perte",
                      "déçoit", "décevant", "inférieur aux attentes", "enquête", "amende", "retard",
                      "licenciement", "suppression", "sanction", "downgrade", "plunge", "falls", "cuts",
                      "lawsuit", "warning", "perturbation"]
_NEWS_BRUIT = re.compile(r"cours action|cotation bourse|cours de bourse en temps r", re.I)


def _nettoyer_texte(s):
    """Texte brut : entités HTML décodées, balises retirées, espaces normalisés."""
    s = html_lib.unescape(s or "")
    s = re.sub(r"<[^>]+>", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def parser_rss(contenu):
    """Articles d'un flux RSS 2.0. Seuls les liens http(s) et les dates lisibles sont gardés."""
    articles = []
    for it in ET.fromstring(contenu).iter("item"):
        titre = _nettoyer_texte(it.findtext("title"))
        lien = (it.findtext("link") or "").strip()
        if not titre or not lien.lower().startswith(("http://", "https://")):
            continue
        try:
            date = pd.Timestamp(email.utils.parsedate_to_datetime(it.findtext("pubDate") or ""))
        except (TypeError, ValueError):
            continue
        date = date.tz_localize("UTC") if date.tzinfo is None else date.tz_convert("UTC")
        source = _nettoyer_texte(it.findtext("source")) or urllib.parse.urlparse(lien).netloc.replace("www.", "")
        if source and titre.endswith(" - " + source):           # suffixe ajouté par Google Actualités
            titre = titre[: -len(" - " + source)]
        resume = _nettoyer_texte(it.findtext("description"))
        if resume.startswith(titre[:40]):                          # Google : la description répète le titre
            resume = ""
        articles.append({"titre": titre, "lien": lien, "resume": resume[:400], "source": source,
                         "date": date.isoformat()})
    return articles


def _url_yahoo_rss(ticker):
    return (f"https://feeds.finance.yahoo.com/rss/2.0/headline?s={urllib.parse.quote(ticker)}"
            f"&region=FR&lang=fr-FR")


def _url_google_news(requete):
    q = urllib.parse.urlencode({"q": f"{requete} when:{NEWS_JOURS}d", "hl": "fr", "gl": "FR", "ceid": "FR:fr"})
    return f"https://news.google.com/rss/search?{q}"


def _lire_rss(url, timeout=10):
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 (dashboard patrimonial)"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return parser_rss(r.read())


def get_news(tickers_actions, themes):
    """Télécharge les flux (cache NEWS_CACHE_TTL_MIN minutes par flux ; flux en échec -> cache même
    périmé). Retourne {clé: {"articles", "timestamp", "etat"}} avec clé "T:<ticker>" ou "H:<thème>"."""
    try:
        with open(NEWS_CACHE, encoding="utf-8") as f:
            cache = json.load(f)
        if not isinstance(cache, dict):
            cache = {}
    except (OSError, ValueError):
        cache = {}

    now = datetime.now()
    requetes = {f"T:{t}": _url_yahoo_rss(t) for t in tickers_actions}
    requetes.update({f"H:{nom}": _url_google_news(cfg["requete"]) for nom, cfg in themes.items()})
    sortie, maj = {}, False
    for cle, url in requetes.items():
        ent = cache.get(cle) if isinstance(cache.get(cle), dict) else None
        age_min = None
        if ent:
            try:
                age_min = (now - datetime.fromisoformat(ent["timestamp"])).total_seconds() / 60
            except (KeyError, TypeError, ValueError):
                ent = None
        if ent and 0 <= age_min < NEWS_CACHE_TTL_MIN:
            sortie[cle] = dict(ent, etat="cache")
            continue
        try:
            articles = _lire_rss(url)
            if cle.startswith("H:"):
                articles = [a for a in articles if not _NEWS_BRUIT.search(a["titre"])]
                articles = sorted(articles, key=lambda a: a["date"], reverse=True)[:NEWS_MAX_PAR_THEME]
            ent = {"timestamp": now.isoformat(timespec="seconds"), "articles": articles}
            cache[cle], maj = ent, True
            sortie[cle] = dict(ent, etat="reseau")
        except Exception:
            sortie[cle] = dict(ent, etat="cache_perime") if ent else {"articles": [], "etat": "indisponible"}

    if maj:
        tmp = NEWS_CACHE + ".tmp"
        try:
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump(cache, f, ensure_ascii=False)
            os.replace(tmp, NEWS_CACHE)
        except OSError:
            pass
    etats = pd.Series([v["etat"] for v in sortie.values()]).value_counts().to_dict()
    print(f"Actualités : {len(sortie)} flux ({', '.join(f'{n} {e}' for e, n in etats.items())}).")
    return sortie


def classer_article(texte):
    """Types détectés (mots-clés), importance maximale et sentiment lexical indicatif dans [-1, 1]."""
    t = texte.lower()
    types = [(nom, imp) for nom, imp, mots in NEWS_TYPES if any(m in t for m in mots)]
    pos = sum(t.count(m) for m in NEWS_MOTS_POSITIFS)
    neg = sum(t.count(m) for m in NEWS_MOTS_NEGATIFS)
    sentiment = (pos - neg) / (pos + neg) if pos + neg else 0.0
    return [n for n, _ in types], max((i for _, i in types), default=None), sentiment


def date_seance_reaction(date_pub, ticker, index_seances):
    """Première séance où le marché peut réagir : le jour même si la publication précède la clôture
    de la place du titre, sinon la séance suivante (week-ends et fériés inclus). None si à venir."""
    tz, cloture = ("America/New_York", 16.0) if "." not in ticker else ("Europe/Paris", 17.5)
    local = pd.Timestamp(date_pub).tz_convert(tz)
    jour = local.normalize().tz_localize(None)
    if local.hour + local.minute / 60 >= cloture:
        jour += pd.Timedelta(days=1)
    suivantes = index_seances[index_seances >= jour]
    return suivantes[0] if len(suivantes) else None


def reaction_marche(ticker, date_pub, raw_prices, bench_returns):
    """Rendement du titre à la séance de réaction, écart avec le benchmark et significativité (z-score
    de l'écart, écart-type mesuré sur les 250 séances précédentes)."""
    if ticker not in raw_prices.columns:
        return None
    r = rendements_reels(raw_prices[ticker], ffill_limit=5)
    if r.empty:
        return None
    jour = date_seance_reaction(date_pub, ticker, r.index)
    if jour is None:
        return {"en_attente": True}
    marche = float(bench_returns.get(jour, 0.0)) if len(bench_returns) else 0.0
    ecart = float(r[jour]) - marche
    hist = (r - bench_returns.reindex(r.index).fillna(0.0)) if len(bench_returns) else r
    hist = hist[hist.index < jour].tail(250)
    sigma = float(hist.std()) if len(hist) >= 60 else None
    z = ecart / sigma if sigma else None
    return {"en_attente": False, "ticker": ticker, "date": f"{jour:%Y-%m-%d}", "rendement": float(r[jour]),
            "marche": marche, "ecart": ecart, "z": z,
            "significative": bool(z is not None and abs(z) >= NEWS_SEUIL_SIGMA)}


def _cle_titre(titre):
    return re.sub(r"[^a-z0-9]", "", titre.lower())[:70]


def _normaliser(s):
    """Minuscules, sans accents, ponctuation remplacée par des espaces."""
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode().lower()
    return re.sub(r"\s+", " ", re.sub(r"[^a-z0-9]", " ", s)).strip()


_ARTICLES_TETE = {"l", "le", "la", "les", "the"}
_SUFFIXES_SOCIETE = {"se", "sa", "s", "a", "sas", "sca", "nv", "n", "v", "ag", "plc", "inc", "incorporated",
                     "corp", "corporation", "group", "groupe", "ltd", "co", "spa"}
_MOTS_GENERIQUES = {"credit", "banque", "groupe", "societe", "compagnie", "amundi", "ishares", "vanguard",
                    "lyxor", "technologies", "international", "holding", "general", "france"}


def alias_societe(noms):
    """Variantes de nom servant à reconnaître une mention de la société : nom sans article ni
    suffixe juridique (« L'Air Liquide S.A. » -> « air liquide ») et premier mot s'il est distinctif."""
    alias = set()
    for nom in noms:
        mots = _normaliser(nom).split()
        while mots and mots[0] in _ARTICLES_TETE:
            mots = mots[1:]
        while len(mots) > 1 and mots[-1] in _SUFFIXES_SOCIETE:
            mots = mots[:-1]
        if mots:
            alias.add(" ".join(mots))
            if len(mots[0]) >= 4 and mots[0] not in _MOTS_GENERIQUES:
                alias.add(mots[0])
    return alias


def mention_directe(texte, alias, ticker):
    """Vrai si le texte cite la société (une de ses variantes de nom, en mot entier) ou son code
    boursier en majuscules (« DBV »), pour les codes d'au moins 3 lettres."""
    norm = f" {_normaliser(texte)} "
    if any(f" {a} " in norm for a in alias):
        return True
    racine = ticker.split(".")[0]
    return len(racine) >= 3 and re.search(rf"\b{re.escape(racine)}\b", texte) is not None


def analyser_actualites(flux, poids, noms, raw_prices, date_ref, themes=None, alias=None):
    """Fusionne les flux, dédoublonne, classe, rapproche du marché et note chaque article.
    poids : {ticker: part de la poche actions (0-1)} ; noms : {ticker: libellé} ;
    alias : {ticker: [noms possibles]} pour reconnaître une mention directe (défaut : noms).
    Score = 100 x importance x racine(exposition / plus grosse ligne) x fraîcheur (demi-vie ~3 j)
    x NEWS_FACTEUR_MENTION_INDIRECTE si la société n'est pas citée, x 1,5 si le cours a réagi
    significativement à un article qui la cite."""
    themes = NEWS_THEMES if themes is None else themes
    alias = {t: alias_societe((alias or {}).get(t, []) + [noms.get(t, t)]) for t in poids}
    date_ref = pd.Timestamp(date_ref)
    date_ref = date_ref.tz_localize("UTC") if date_ref.tzinfo is None else date_ref.tz_convert("UTC")
    bench = rendements_reels(raw_prices[TICKER_BENCHMARK], ffill_limit=5) \
        if TICKER_BENCHMARK in raw_prices.columns else pd.Series(dtype=float)
    poids_max = max(poids.values()) if poids else 1.0

    fusion = {}
    for cle, ent in flux.items():
        genre, code = cle.split(":", 1)
        if genre == "T":
            cible = {"code": code, "nom": noms.get(code, code), "poids": poids.get(code, 0.0), "theme": False}
        else:
            cfg = themes.get(code, {})
            lignes = cfg.get("tickers", [])
            expo = sum(poids.values()) if lignes == "*" else sum(poids.get(t, 0.0) for t in lignes)
            if expo <= 0:
                continue                        # thème sans ligne détenue : ne touche pas le portefeuille
            cible = {"code": code, "nom": code, "poids": expo, "theme": True}
        for a in ent.get("articles", []):
            d = pd.Timestamp(a["date"])
            age = (date_ref - d).total_seconds() / 86400
            if age > NEWS_JOURS or age < -1:
                continue
            k = _cle_titre(a["titre"])
            if k in fusion:
                if cible["code"] not in [c["code"] for c in fusion[k]["cibles"]]:
                    fusion[k]["cibles"].append(cible)
                continue
            fusion[k] = dict(a, cibles=[cible], age_jours=max(age, 0.0))

    articles = []
    for a in fusion.values():
        types, importance, sentiment = classer_article(f"{a['titre']} {a['resume']}")
        lignes = [c for c in a["cibles"] if not c["theme"]]
        mention, facteur, reaction = None, 1.0, None
        if lignes:
            importance = importance if importance is not None else 0.3
            texte = f"{a['titre']} {a['resume']}"
            citees = [c for c in lignes if mention_directe(texte, alias.get(c["code"], set()), c["code"])]
            mention = "directe" if citees else "indirecte"
            facteur = 1.0 if citees else NEWS_FACTEUR_MENTION_INDIRECTE
            principale = max(citees or lignes, key=lambda c: c["poids"])["code"]
            reaction = reaction_marche(principale, a["date"], raw_prices, bench)
        else:
            # Article de thème : les mots-clés « résultats », « fusion »… y visent des sociétés non
            # détenues (Tesla…) ; type et importance fixes.
            types, importance = ["Macro & marché"], 0.4
        expo = sum(c["poids"] for c in a["cibles"])
        # Un mouvement de cours n'est rattaché qu'à un article qui cite la société
        signif = bool(reaction and reaction.get("significative") and mention == "directe")
        fraicheur = math.exp(-a["age_jours"] / 4.0)
        score = (100 * importance * math.sqrt(min(expo / poids_max, 1.0)) * fraicheur * facteur
                 * (1.5 if signif else 1.0))
        articles.append({
            "titre": a["titre"], "lien": a["lien"], "resume": a["resume"], "source": a["source"],
            "date": a["date"], "age_jours": round(a["age_jours"], 2),
            "cibles": [{"code": c["code"], "nom": c["nom"], "poids": round(c["poids"] * 100, 2),
                        "theme": c["theme"]} for c in a["cibles"]],
            "types": types, "importance": importance,
            "sentiment": round(sentiment, 2),
            "sentiment_label": "positif" if sentiment > 0.2 else ("négatif" if sentiment < -0.2 else "neutre"),
            "reaction": reaction, "significative": signif, "mention": mention, "score": round(score, 1),
        })
    articles.sort(key=lambda x: x["score"], reverse=True)
    return articles


def agenda_evenements(df_fundamentals, poids, noms, date_ref, horizon_jours=NEWS_AGENDA_JOURS):
    """Prochains résultats et détachements de dividende des lignes détenues (données Yahoo)."""
    if df_fundamentals is None or df_fundamentals.empty or "Ticker" not in df_fundamentals:
        return []
    debut = pd.Timestamp(date_ref).normalize()
    fin = debut + pd.Timedelta(days=horizon_jours)
    evenements = []
    for _, row in df_fundamentals[df_fundamentals["Ticker"].isin(list(poids))].iterrows():
        for col, libelle in (("Prochains_Resultats", "Publication de résultats"), ("Ex_Dividende", "Détachement du dividende")):
            ts = row.get(col)
            if ts is None or pd.isna(ts):
                continue
            d = pd.Timestamp(int(ts), unit="s", tz="UTC").tz_convert("Europe/Paris").tz_localize(None)
            if debut <= d.normalize() <= fin:
                evenements.append({"date": d, "evenement": libelle, "ticker": row["Ticker"],
                                   "nom": noms.get(row["Ticker"], row["Ticker"]),
                                   "poids": poids.get(row["Ticker"], 0.0) * 100,
                                   "jours": (d.normalize() - debut).days})
    return sorted(evenements, key=lambda e: e["date"])


# =================================================================
# 2. MOTEUR DE CALCUL
# =================================================================

def run_engine():
    df_ops = load_stock_data(os.path.join(DATA_DIR, FILE_BOURSE))
    if df_ops is None:
        raise SystemExit(f"Fichier introuvable : {os.path.join(DATA_DIR, FILE_BOURSE)}.xlsm / .xlsx "
                         f"(dossier modifiable avec --data-dir).")
    if isinstance(df_ops, str):
        raise SystemExit(f"Erreur Bourse : {df_ops}")
    alertes_saisie = valider_retraits(df_ops)
    for msg in alertes_saisie:
        print(msg)

    df_banque_raw = load_bank_data(os.path.join(DATA_DIR, BASE_NAME_BANQUE))

    start_date_bourse = df_ops['Date'].min()
    if not isinstance(df_banque_raw, str) and not df_banque_raw.empty:
        max_date_banque = df_banque_raw['Date'].max()
        start_date_banque = df_banque_raw['Date'].min()
    else:
        max_date_banque = datetime.now()
        start_date_banque = start_date_bourse

    base_tickers = df_ops['Ticker'].unique().tolist()
    base_tickers = [t for t in base_tickers if isinstance(t, str) and t.strip() != '']
    
    watchlist_tickers = load_watchlist(os.path.join(DATA_DIR, FILE_WATCHLIST))
    
    all_api_tickers = list(set(base_tickers + watchlist_tickers + VIX_TERM_TICKERS + [VIX_TICKER, TICKER_BENCHMARK]
                               + list(HEDGE_INSTRUMENTS)))

    df_fundamentals, currency_map = get_fundamentals_and_currency(all_api_tickers)
    for t, cfg in HEDGE_INSTRUMENTS.items():
        currency_map.setdefault(t, cfg['devise'])

    prices, all_days, raw_prices = get_market_data(
        all_api_tickers, min(start_date_bourse, start_date_banque), max_date_banque, currency_map)
    
    if prices is None:
        raise SystemExit("Erreur Yahoo Finance (aucune donnée récupérée).")

    history_records = []
    portfolio, ticker_names = {}, {}
    solde_especes = 0.0
    total_virements = 0.0
    total_dividendes = 0.0
    ops_group = df_ops.groupby('Date')

    for day in all_days:
        if day in ops_group.groups:
            for _, row in ops_group.get_group(day).iterrows():
                t = row['Ticker']
                q = row['Qte'] if pd.notna(row['Qte']) else 0.0
                m = row['Montant'] if pd.notna(row['Montant']) else 0.0
                sens = str(row['Sens']).strip().lower()
                ticker_names[t] = row['Nom'] if pd.notna(row['Nom']) else t
                q_abs, m_abs = abs(q), abs(m)

                compte_val = str(row.get('Compte', 'CTO')).strip().upper()
                if compte_val in ['NAN', 'NONE', '']: compte_val = 'CTO'
                
                port_key = (t, compte_val)

                if any(k in sens for k in ['virement', 'depot', 'dépôt', 'apport', 'retrait']):
                    solde_especes += m; total_virements += m; continue
                if 'dividende' in sens or 'coupon' in sens:
                    solde_especes += m_abs; total_dividendes += m_abs
                    if t != 'CASH_OP' and q_abs > 0:
                        portfolio.setdefault(port_key, {'qty': 0.0, 'investi': 0.0})['qty'] += q_abs
                    continue
                if any(k in sens for k in ['frais', 'taxe', 'impot', 'impôt']):
                    solde_especes -= m_abs; continue
                if t == 'CASH_OP':
                    continue

                portfolio.setdefault(port_key, {'qty': 0.0, 'investi': 0.0})
                if 'vente' in sens:
                    if portfolio[port_key]['qty'] > 0:
                        ratio = min(q_abs / portfolio[port_key]['qty'], 1)
                        portfolio[port_key]['investi'] -= portfolio[port_key]['investi'] * ratio
                        portfolio[port_key]['qty'] -= q_abs
                    solde_especes += m_abs
                elif 'achat' in sens or 'souscription' in sens:
                    portfolio[port_key]['qty'] += q_abs; portfolio[port_key]['investi'] += m_abs
                    solde_especes -= m_abs
                elif 'ost' in sens:
                    if q != 0: portfolio[port_key]['qty'] += q
                    if m != 0: solde_especes += m

                if portfolio[port_key]['qty'] < 0.00001:
                    portfolio[port_key]['qty'] = 0; portfolio[port_key]['investi'] = 0

        valeur_actions_jour, investi_actions_jour = 0.0, 0.0
        for (t, compte_val), data in portfolio.items():
            qty = data['qty']
            if qty > 0:
                price = 0.0
                if t in prices.columns:
                    try:
                        price = prices.at[day, t]
                        if pd.isna(price): price = prices[t].iloc[-1]
                        if pd.isna(price): price = 0.0
                    except Exception:
                        price = 0.0
                valeur_actions_jour += qty * price
                investi_actions_jour += data['investi']

        valeur_bourse_totale = valeur_actions_jour + solde_especes
        history_records.append({
            'Date': day, 'Valeur_Actions': valeur_actions_jour, 'Especes': solde_especes,
            'Valeur_Bourse': valeur_bourse_totale, 'Investi_Actions': investi_actions_jour,
            'Investi_Bourse': total_virements, 'Total_Dividendes': total_dividendes
        })

    df_hist_bourse = pd.DataFrame(history_records).set_index('Date')
    df_hist_bourse['PV_Latente'] = df_hist_bourse['Valeur_Actions'] - df_hist_bourse['Investi_Actions']
    df_hist_bourse['PV_Totale'] = df_hist_bourse['Valeur_Bourse'] - df_hist_bourse['Investi_Bourse']
    df_hist_bourse = df_hist_bourse.fillna(0.0)

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

    df_global = pd.concat([df_hist_bourse, df_hist_banque], axis=1).ffill().fillna(0)
    
    real_start_date = min(start_date_bourse, start_date_banque)
    df_global = df_global[df_global.index >= real_start_date]
    df_hist_bourse = df_hist_bourse[df_hist_bourse.index >= real_start_date]
    df_hist_banque = df_hist_banque[df_hist_banque.index >= real_start_date]

    # Indice de performance time-weighted (dépôts/retraits neutralisés), partagé par les onglets
    df_twr = compute_twr(df_hist_bourse)

    df_global['Patrimoine_Total'] = df_global['Valeur_Bourse'] + df_global['Valeur_Banque']
    df_global_monthly = df_global.resample('ME').last()

    details_data, sim_assets, active_tickers_list = [], {}, []
    last_prices_smooth = prices.iloc[-1]

    total_var_24h_euro = 0.0
    for (t, compte_val), data in portfolio.items():
        qty = data['qty']
        if qty > 0.001:
            if t not in active_tickers_list:
                active_tickers_list.append(t)
            curr_price = last_prices_smooth.get(t, 0.0)
            if pd.isna(curr_price): curr_price = 0.0
            
            var_pct, var_eur = 0.0, 0.0
            
            fund = df_fundamentals[df_fundamentals["Ticker"] == t]
            secteur = "Inconnu"
            if not fund.empty:
                secteur = fund.iloc[0].get("Secteur", "Inconnu")
                live_pct = fund.iloc[0].get("Var_Live_Pct")
                if pd.notna(live_pct):
                    var_pct = float(live_pct)
                    if curr_price > 0:
                        p_prev_est = curr_price / (1 + var_pct / 100)
                        var_eur = qty * (curr_price - p_prev_est)

            if var_pct == 0.0 and t in raw_prices.columns:
                hist = raw_prices[t].dropna()
                if len(hist) >= 2:
                    p_now, p_prev = hist.iloc[-1], hist.iloc[-2]
                    if p_prev > 0:
                        var_pct = (p_now - p_prev) / p_prev * 100
                        var_eur = qty * (p_now - p_prev)

            total_var_24h_euro += var_eur
            val = qty * curr_price
            if val > 0:
                sim_assets[t] = sim_assets.get(t, 0.0) + val
                
            details_data.append({
                "Actif": ticker_names.get(t, t), "Ticker": t, "Compte": compte_val, "Secteur": secteur, "Qté": qty,
                "Devise": currency_map.get(t, DEVISE_BASE),
                "PRU": data['investi'] / qty, "Cours": curr_price,
                "Var 24h %": var_pct, "Var 24h €": var_eur, "Valeur": val,
                "Plus-Value": val - data['investi'],
                "Rendement %": (val - data['investi']) / data['investi'] * 100 if data['investi'] > 0 else 0
            })
            
    df_details = pd.DataFrame(details_data)
    
    # Enrichissement mathématique (Poids, Contrib, fiscalité par enveloppe)
    fiscalite = {}
    if not df_details.empty:
        total_valeur = df_details['Valeur'].sum()
        total_pv = df_details['Plus-Value'].sum()
        df_details['Poids %'] = (df_details['Valeur'] / total_valeur) * 100 if total_valeur > 0 else 0
        df_details['Contrib P&L %'] = (df_details['Plus-Value'] / total_pv) * 100 if total_pv != 0 else 0
        df_details, total_pfu, fiscalite = calcul_fiscalite_enveloppes(df_details, datetime.now())
    else:
        total_pfu = 0.0

    # Système d'alertes
    alerts = list(alertes_saisie)
    vix_data, vix_regime = {}, {}
    if VIX_TICKER in raw_prices.columns:
        vix_series = raw_prices[VIX_TICKER].dropna()
        if len(vix_series) > 0:
            current_vix = vix_series.iloc[-1]
            vix_mm20 = vix_series.rolling(20).mean()
            mean_1y = vix_series.tail(252).mean()
            std_1y = vix_series.tail(252).std()
            z_score = (current_vix - mean_1y) / std_1y if std_1y > 0 else 0
            pct_2y = (vix_series.tail(504) < current_vix).mean() * 100

            vix_data = {
                "history": vix_series.tail(504), "mm20": vix_mm20.tail(504),
                "current": current_vix, "z_score": z_score, "percentile": pct_2y,
            }
            if current_vix < VIX_SEUILS["complaisance"]: label = "Complaisance"
            elif current_vix < VIX_SEUILS["normal"]: label = "Normal"
            elif current_vix < VIX_SEUILS["eleve"]: label = "Élevé"
            else: label = "Stress"
            vix_regime = {"label": label, "comment": f"VIX à {current_vix:.1f}"}

            if label in ["Élevé", "Stress"]:
                alerts.append(f"⚠️ Alerte Risque : Le marché est en régime de {label} (VIX > {VIX_SEUILS['normal']}). Couverture ou patience recommandée.")

    # Surface de volatilité implicite des options S&P 500 (sous-jacent du VIX)
    vol_surface = get_vol_surface()

    # Actualités : flux par ticker pour les actions détenues (les ETF n'en ont pas) + thèmes
    # d'exposition (Google Actualités), puis notation et rapprochement avec les cours.
    total_sim = sum(sim_assets.values())
    poids_lignes = {t: v / total_sim for t, v in sim_assets.items()} if total_sim > 0 else {}
    noms_lignes = {t: str(ticker_names.get(t, t)) for t in poids_lignes}
    types_titres = (dict(zip(df_fundamentals['Ticker'], df_fundamentals['Type']))
                    if not df_fundamentals.empty and 'Type' in df_fundamentals.columns else {})
    flux_news = get_news([t for t in poids_lignes if types_titres.get(t) != 'ETF'], NEWS_THEMES)
    alias_lignes = ({t: [n] for t, n in zip(df_fundamentals['Ticker'], df_fundamentals['Nom_Long']) if isinstance(n, str)}
                    if not df_fundamentals.empty and 'Nom_Long' in df_fundamentals.columns else {})
    actualites = analyser_actualites(flux_news, poids_lignes, noms_lignes, raw_prices, pd.Timestamp.now(tz="UTC"),
                                     alias=alias_lignes)
    agenda = agenda_evenements(df_fundamentals, poids_lignes, noms_lignes, datetime.now())
    news_meta = {
        "nb_flux": len(flux_news),
        "etats": pd.Series([v["etat"] for v in flux_news.values()]).value_counts().to_dict() if flux_news else {},
        "maj": max((v["timestamp"] for v in flux_news.values() if v.get("timestamp")), default=None),
    }
    # Alertes : actualité récente ayant fait réagir le cours, événement dans les 5 jours.
    # Titres et noms viennent de sources externes : échappés car le bandeau est rendu en HTML.
    for a in [x for x in actualites if x['significative'] and x['age_jours'] <= 3][:3]:
        r = a['reaction']
        nom = next(c['nom'] for c in a['cibles'] if c['code'] == r['ticker'])
        alerts.append(
            f"📰 Actualité à impact sur {html_lib.escape(nom)} : {r['rendement'] * 100:+.1f} % le "
            f"{pd.Timestamp(r['date']):%d/%m} (écart {r['ecart'] * 100:+.1f} % vs marché, {r['z']:+.1f} σ) — "
            f"« {html_lib.escape(a['titre'])} »")
    for e in agenda:
        if e['jours'] <= 5:
            quand = "aujourd'hui" if e['jours'] == 0 else f"le {e['date']:%d/%m}"
            alerts.append(f"📅 {e['evenement']} — {html_lib.escape(e['nom'])} ({e['poids']:.1f} % de la poche actions) {quand}.")

    def clean_nan(val):
        return None if pd.isna(val) else float(val)

    watchlist_data = []
    for t in set(watchlist_tickers):
        if t in raw_prices.columns:
            p_series = raw_prices[t].dropna()
            if len(p_series) > 130:
                curr_p = p_series.iloc[-1]
                p_6m = p_series.iloc[-130]
                mom_6m = (curr_p / p_6m - 1) * 100
                dist_mm200 = (curr_p / p_series.rolling(200).mean().iloc[-1] - 1) * 100
                fund = df_fundamentals[df_fundamentals["Ticker"] == t].iloc[0] if not df_fundamentals[df_fundamentals["Ticker"] == t].empty else pd.Series()
                per = fund.get("PER", np.nan)
                roe = fund.get("ROE", np.nan)
                debt_eq = fund.get("Dette/Eq", np.nan)          
                croiss_ca = fund.get("Croissance CA", np.nan)   

                secteur = fund.get("Secteur", "Inconnu")
                if pd.isna(secteur) or not str(secteur).strip():
                    secteur = "Inconnu"

                watchlist_data.append({
                    "Ticker": t, "Secteur": str(secteur), "PER": clean_nan(per), "ROE": clean_nan(roe), "Dette_Eq": clean_nan(debt_eq),
                    "Croiss_CA": clean_nan(croiss_ca), "Mom_6M": clean_nan(mom_6m), "Dist_MM200": clean_nan(dist_mm200)
                })

    # Alerte Drawdown : sur l'indice TWR (un dépôt ne masque plus une baisse, un retrait n'en crée pas)
    if not df_twr.empty:
        current_dd = drawdown_courant_twr(df_twr)
        if current_dd < -15:
            alerts.append(f"🚨 Alerte Portefeuille : Vous êtes actuellement dans un Drawdown de {current_dd:.1f}% par rapport au plus-haut historique (performance time-weighted, hors effet des versements).")

    return dict(
        df_ops=df_ops, df_hist_bourse=df_hist_bourse, df_twr=df_twr, df_global=df_global,
        df_global_monthly=df_global_monthly, df_details=df_details, df_fundamentals=df_fundamentals,
        last_banque_state=last_banque_state, tickers=base_tickers, active_tickers_list=active_tickers_list,
        sim_assets=sim_assets, raw_prices=raw_prices, total_var_24h_euro=total_var_24h_euro,
        currency_map=currency_map, vix_data=vix_data, vix_regime=vix_regime, vix_seuils=VIX_SEUILS,
        vol_surface=vol_surface, actualites=actualites, agenda=agenda, news_meta=news_meta,
        watchlist_data=watchlist_data, alerts=list(set(alerts)), total_pfu=total_pfu,
        fiscalite=fiscalite
    )


# =================================================================
# 3. ASSEMBLAGE DES ONGLETS HTML
# =================================================================

def build_tab_patrimoine(ctx):
    df_global = ctx['df_global']; df_global_monthly = ctx['df_global_monthly']
    last_banque_state = ctx['last_banque_state']
    total_now = df_global.iloc[-1]['Patrimoine_Total']
    banque_now = df_global.iloc[-1]['Valeur_Banque']
    bourse_now = ctx['df_hist_bourse'].iloc[-1]['Valeur_Bourse']
    
    total_pfu = ctx['total_pfu']
    patrimoine_net = total_now - total_pfu
    delta_eur = ctx['total_var_24h_euro']
    delta_pct = (delta_eur / (total_now - delta_eur) * 100) if (total_now - delta_eur) > 0 else 0

    months = df_global_monthly.index.strftime('%Y-%m').tolist()
    spark = rb.fig_sparkline(months, df_global_monthly['Patrimoine_Total'].tolist())
    monthly_chart = rb.fig_monthly_stack(months, df_global_monthly['Valeur_Banque'].tolist(),
                                          df_global_monthly['Valeur_Bourse'].tolist())

    macro = rb.fig_donut([banque_now, bourse_now], ["Banque (sécurisé)", "Bourse (risque)"],
                          [rb.SAGE, rb.GOLD], rb.eur(total_now))

    micro_rows = []
    if not last_banque_state.empty:
        micro_vals = last_banque_state['Montant'].tolist() + [bourse_now]
        micro_labels = last_banque_state['Compte'].tolist() + ["Bourse (global)"]
        palette = ["#3E5A80", "#5A7CA6", "#8B96A8", "#C9A15E", "#E3C384", "#6E8CB0"]
        micro = rb.fig_donut(micro_vals, micro_labels, (palette * 3)[:len(micro_vals)])
        for _, r in last_banque_state.sort_values('Montant', ascending=False).iterrows():
            micro_rows.append([r['Compte'], rb.eur(r['Montant']), rb.pct(r['Montant'] / total_now * 100, signed=False)])
        micro_rows.append(["Bourse (global)", rb.eur(bourse_now), rb.pct(bourse_now / total_now * 100, signed=False)])
    else:
        micro = "<div class='note'>Aucune donnée bancaire trouvée.</div>"

    body = f"""
    <div class="hero">
      {spark}
      <div class="hero-label">Patrimoine Brut Actuel</div>
      <div class="hero-number">
        <span><span class="cur">€</span>{rb.eur(total_now)[:-2]}</span>
        <span class="delta-badge {'neg' if delta_eur<0 else ''}">{rb.eur(delta_eur, signed=True)} · {rb.pct(delta_pct)} (24h)</span>
      </div>
      <div class="stat-row">
        {rb.stat_chip("Valeur Réelle (Après Impôts)", rb.eur(patrimoine_net), f"Prélèvements (PEA/CTO) : {rb.eur(total_pfu)}")}
        {rb.stat_chip("Banque · sécurisé", rb.eur(banque_now), rb.pct(banque_now/total_now*100, signed=False))}
        {rb.stat_chip("Bourse · actions + espèces", rb.eur(bourse_now), rb.pct(bourse_now/total_now*100, signed=False))}
      </div>
    </div>
    {rb.section("Évolution mensuelle", f'<div class="card">{monthly_chart}</div>')}
    {rb.section("Allocation", f'''<div class="grid-2">
        <div class="card"><div class="card-title">Stratégique (macro)</div>{macro}</div>
        <div class="card"><div class="card-title">Détaillée (micro)</div>{micro}</div>
      </div>''')}
    {rb.section("Détail par compte", f'<div class="card">{rb.ledger_table(["Compte","Montant","Part"], micro_rows, num_cols=[1,2])}</div>' if micro_rows else "")}
    """
    return body


def build_tab_bourse(ctx):
    df_hist = ctx['df_hist_bourse']
    df_details = ctx['df_details'].copy()
    
    val_now = df_hist.iloc[-1]['Valeur_Bourse']
    invest_now = df_hist.iloc[-1]['Investi_Bourse']
    pv_totale = df_hist.iloc[-1]['PV_Totale']
    especes = df_hist.iloc[-1]['Especes']
    dividendes = df_hist.iloc[-1]['Total_Dividendes']
    pct_total = (pv_totale / invest_now * 100) if invest_now > 0 else 0
    
    raw_prices = ctx['raw_prices']
    has_bench = TICKER_BENCHMARK in raw_prices.columns
    if has_bench:
        bench_prices = raw_prices[TICKER_BENCHMARK].reindex(df_hist.index).ffill().bfill()
        bench_returns = bench_prices.pct_change().fillna(0)
        cf_series = df_hist['Investi_Bourse'].diff().fillna(df_hist['Investi_Bourse'].iloc[0])

        bench_val = [cf_series.iloc[0]]
        for i in range(1, len(df_hist)):
            new_val = bench_val[-1] * (1 + bench_returns.iloc[i]) + cf_series.iloc[i]
            bench_val.append(new_val)

        df_hist['Valeur_Benchmark'] = bench_val
        val_bench_now = df_hist['Valeur_Benchmark'].iloc[-1]
        surperformance_eur = val_now - val_bench_now
    else:
        val_bench_now = invest_now
        surperformance_eur = 0.0

    surperformance_pct = (surperformance_eur / invest_now * 100) if invest_now > 0 else 0.0
    # Écart de valeur avec un placement des mêmes flux dans le benchmark (pas un alpha au sens CAPM)
    if has_bench:
        surperf_val = rb.eur(surperformance_eur, signed=True)
        surperf_sub = f"{rb.pct(surperformance_pct, signed=True)} des apports · vs {TICKER_BENCHMARK} à flux identiques"
    else:
        surperf_val, surperf_sub = "n/d", f"Benchmark {TICKER_BENCHMARK} indisponible"
    delta_eur = ctx['total_var_24h_euro']
    val_prev = val_now - delta_eur
    delta_pct = (delta_eur / val_prev * 100) if val_prev > 0 else 0
    
    top_movers_html = ""
    if not df_details.empty:
        df_details['Abs_Var'] = df_details['Var 24h %'].abs()
        top_3 = df_details.sort_values(by='Abs_Var', ascending=False).head(3)
        movers_chips = [rb.stat_chip(r['Actif'], rb.colorize(r['Var 24h %'], rb.pct(r['Var 24h %'])), rb.colorize(r['Var 24h €'], rb.eur(r['Var 24h €'], signed=True))) for _, r in top_3.iterrows()]
        if movers_chips:
            top_movers_html = f"""
            <div style="margin-top: 24px; margin-bottom: 8px;"><span class="eyebrow">Les 3 plus forts mouvements (Top/Flop 24h)</span></div>
            <div class="stat-row">{''.join(movers_chips)}</div>
            """

    banner_html = f"""
    <div class="hero" style="padding-bottom: 24px; margin-bottom: 32px; border-bottom: 1px solid var(--line);">
      <div class="hero-label">Variation du portefeuille (24h)</div>
      <div class="hero-number">
        <span><span class="cur">€</span>{rb.eur(delta_eur, signed=True)[:-2]}</span>
        <span class="delta-badge {'neg' if delta_eur<0 else ''}">{rb.pct(delta_pct)}</span>
      </div>
      {top_movers_html}
    </div>
    """

    metrics_html = f"""
    <div class="grid-3" style="margin-bottom: 20px;">
      {rb.metric_card("Valeur totale compte", rb.eur(val_now))}
      {rb.metric_card("Apports nets", rb.eur(invest_now))}
      {rb.metric_card("Plus-value totale", rb.eur(pv_totale, signed=True), rb.pct(pct_total), pv_totale>=0)}
    </div>
    <div class="grid-3">
      {rb.metric_card("Surperformance vs benchmark", surperf_val, surperf_sub, surperformance_eur>=0)}
      {rb.metric_card("Espèces", rb.eur(especes))}
      {rb.metric_card("Dividendes perçus", rb.eur(dividendes))}
    </div>
    """

    # Rendements time-weighted calculés une seule fois dans run_engine (compute_twr)
    df_perf = ctx['df_twr']

    monthly_grouped = df_perf.groupby(['Year', 'Month'])['Gross_Ret'].prod() - 1
    pivot_monthly = monthly_grouped.unstack(level=1)
    for m in range(1, 13):
        if m not in pivot_monthly.columns:
            pivot_monthly[m] = np.nan
    pivot_monthly = pivot_monthly[list(range(1, 13))]
    pivot_monthly.columns = ['Jan', 'Fév', 'Mar', 'Avr', 'Mai', 'Juin', 'Juil', 'Aoû', 'Sep', 'Oct', 'Nov', 'Déc']
    heatmap_chart = rb.fig_returns_heatmap(pivot_monthly)
    
    df_perf_all = df_perf.copy() 
    if not df_perf_all.empty:
        yearly_perf = df_perf_all.groupby('Year')['Gross_Ret'].prod() - 1
        nb_years = (df_perf_all.index.max() - df_perf_all.index.min()).days / 365.25
        total_twr = df_perf_all['Gross_Ret'].prod()
        ann_return = (total_twr ** (1 / nb_years)) - 1 if nb_years >= 1.0 else total_twr - 1
    else:
        yearly_perf = pd.Series(dtype=float)
        ann_return = 0.0
    
    ann_return_html = f"""<div class="stat-row" style="margin-bottom: 12px; background: rgba(201,161,94,0.06); border-color: rgba(201,161,94,0.25);">
        {rb.stat_chip("Moyenne Annualisée (Depuis l'origine)", rb.colorize(ann_return, rb.pct(ann_return * 100, signed=True)))}
    </div>"""
    perf_chips = [rb.stat_chip(str(y), rb.colorize(perf, rb.pct(perf * 100, signed=True))) for y, perf in yearly_perf.items()]
    perf_html = ann_return_html + (f'<div class="stat-row" style="margin-bottom: 24px;">{"".join(perf_chips)}</div>' if perf_chips else "")

    dates = df_hist.index.strftime('%Y-%m-%d').tolist()
    line_chart = rb.fig_line_vs_invested(dates, df_hist['Valeur_Bourse'].tolist(), df_hist['Investi_Bourse'].tolist())
    pv_chart = rb.fig_pv_area(dates, df_hist['PV_Latente'].tolist(), df_hist['PV_Totale'].tolist())

    positions_html = ""
    alloc_pie, sector_pie, currency_pie = "", "", ""
    if not df_details.empty:
        rows = []
        for _, r in df_details.sort_values('Valeur', ascending=False).iterrows():
            rows.append([
                rb.colorize(r['Contrib P&L %'], rb.pct(r['Contrib P&L %'])),
                f"{r['Actif']} <span class='tag'>{r['Compte']}</span> <span class='tag'>{r['Devise']}</span>",
                r['Secteur'],
                rb.pct(r['Poids %'], signed=False),
                rb.eur(r['PRU']),
                rb.eur(r['Cours']),
                rb.colorize(r['Var 24h %'], rb.pct(r['Var 24h %'])),
                rb.colorize(r['Rendement %'], rb.pct(r['Rendement %'])),
                rb.colorize(r['Plus-Value'], rb.eur(r['Plus-Value'], signed=True)),
                rb.eur(r['PFU']),
                rb.eur(r['Valeur_Nette']),
            ])
        positions_html = rb.ledger_table(
            ["Contrib. P&L", "Actif", "Secteur", "Poids", "PRU", "Cours", "Var 24h", "Rendement", "PV Brute", "Fiscalité latente", "Val. Nette"],
            rows, num_cols=[0, 3, 4, 5, 6, 7, 8, 9, 10])
            
        alloc_pie = rb.fig_donut(df_details['Valeur'].tolist(), df_details['Actif'].tolist())
        df_sec = df_details.groupby('Secteur')['Valeur'].sum().reset_index()
        sector_pie = rb.fig_donut(df_sec['Valeur'].tolist(), df_sec['Secteur'].tolist())
        df_dev = df_details.groupby('Devise')['Valeur'].sum().reset_index()
        currency_pie = rb.fig_donut(df_dev['Valeur'].tolist(), df_dev['Devise'].tolist())
    else:
        alloc_pie = "<div class='note'>Aucune position valorisée.</div>"

    # Note fiscale : détail par enveloppe + avertissement si le PFU est appliqué à un PEA
    fiscal_note = ""
    fiscalite = ctx.get('fiscalite', {})
    if fiscalite:
        const_dates = {'PEA': 'PEA_DATE_OUVERTURE', 'PEA-PME': 'PEA_PME_DATE_OUVERTURE'}
        lignes = []
        for env in ['CTO', 'PEA', 'PEA-PME']:
            if env not in fiscalite:
                continue
            i = fiscalite[env]
            lignes.append(f"<b>{env}</b> : plus-value latente nette {rb.eur(i['pv_nette'], signed=True)} "
                          f"→ impôt latent {rb.eur(i['impot'])} (taux {rb.pct(i['taux'] * 100, signed=False)})")
        avertissements = []
        for env in ['PEA', 'PEA-PME']:
            i = fiscalite.get(env)
            if not i or not i['pfu_applique']:
                continue
            if i['date_ouverture'] is None:
                raison = f"date d'ouverture inconnue (renseignez <code>{const_dates[env]}</code> dans la config)"
            else:
                raison = (f"plan ouvert le {i['date_ouverture']:%d/%m/%Y}, moins de 5 ans "
                          f"au {datetime.now():%d/%m/%Y}")
            avertissements.append(
                f"⚠️ <b>{env}</b> : PFU de {rb.pct(PFU_RATE * 100, signed=False)} appliqué au lieu des "
                f"prélèvements sociaux seuls ({rb.pct(PEA_RATE * 100, signed=False)}) — {raison}. "
                f"Un retrait avant 5 ans clôture le plan et le gain est imposé au PFU.")
        fiscal_note = (f"<div class='note'>Fiscalité latente calculée <b>par enveloppe</b> : les moins-values "
                       f"compensent les plus-values au sein d'une même enveloppe, puis l'impôt est réparti sur "
                       f"les lignes au prorata de leurs plus-values positives.<br>{'<br>'.join(lignes)}"
                       + (f"<br>{'<br>'.join(avertissements)}" if avertissements else "") + "</div>")

    return f"""
    {banner_html}
    {rb.section("Résultats du portefeuille", metrics_html)}
    {rb.section("Performance historique", f'''
      {perf_html}
      <div class="card" style="margin-bottom: 24px; padding: 12px 22px;">
        <div class="card-title">Rendements mensuels (Time-Weighted)</div>
        {heatmap_chart}
      </div>
      <div class="grid-2">
        <div class="card"><div class="card-title">Valeur vs Apports nets</div>{line_chart}</div>
        <div class="card"><div class="card-title">Gains absolus (latents vs totaux)</div>{pv_chart}</div>
      </div>''')}
    {rb.section("Transparence des Risques (Look-through)", f'''
      <div class="grid-3">
        <div class="card"><div class="card-title">Allocation Actifs</div>{alloc_pie}</div>
        <div class="card"><div class="card-title">Exposition Secteurs</div>{sector_pie}</div>
        <div class="card"><div class="card-title">Exposition Devises</div>{currency_pie}</div>
      </div>
    ''')}
    {rb.section("Positions en cours (Impact Fiscal et Contribution)", (f'<div class="card table-container">{positions_html}</div>' if positions_html else "<div class='note'>Aucune position.</div>") + fiscal_note)}
    """
   
def build_tab_epargne(ctx, cible_mensuelle=CIBLE_MENSUELLE):
    df_ops = ctx['df_ops']
    mask = df_ops['Sens'].astype(str).str.lower().str.contains('virement|depot|dépôt|apport')
    df_mvt = df_ops[mask].copy()
    if df_mvt.empty:
        return "<div class='note'>Aucun virement détecté dans l'historique d'ordres.</div>"

    df_mvt['Mois_Str'] = df_mvt['Date'].dt.strftime('%Y-%m')
    apports = df_mvt.groupby('Mois_Str')['Montant'].sum().reset_index()
    apports.columns = ['Mois', 'Apport_Reel']
    start = df_mvt['Date'].min()
    all_months = pd.date_range(start=start.replace(day=1), end=datetime.now(), freq='MS')
    df_suivi = pd.DataFrame({'Mois': [d.strftime('%Y-%m') for d in all_months]})
    df_suivi = df_suivi.merge(apports, on='Mois', how='left').fillna(0)
    df_suivi['Cible'] = cible_mensuelle
    df_suivi['Ecart'] = df_suivi['Cible'] - df_suivi['Apport_Reel']
    df_suivi['Retard_Cumule'] = df_suivi['Ecart'].cumsum()

    retard = df_suivi.iloc[-1]['Retard_Cumule']
    mois_courant = df_suivi.iloc[-1]['Apport_Reel']

    metrics = f"""<div class="metric-row">
      {rb.metric_card("Cible mensuelle", rb.eur(cible_mensuelle))}
      {rb.metric_card("Investi ce mois-ci", rb.eur(mois_courant), rb.eur(mois_courant-cible_mensuelle, signed=True), mois_courant>=cible_mensuelle)}
      {rb.metric_card("Retard cumulé" if retard>0 else "Avance cumulée", rb.eur(abs(retard)), None, retard<=0)}
    </div>"""

    chart = rb.fig_savings_bar_line(df_suivi['Mois'].tolist(), df_suivi['Apport_Reel'].tolist(), df_suivi['Cible'].tolist())

    rows = []
    for _, r in df_suivi.sort_values('Mois', ascending=False).iterrows():
        rows.append([r['Mois'], rb.eur(r['Apport_Reel']), rb.eur(r['Cible']),
                     rb.colorize(-r['Ecart'], rb.eur(r['Ecart'], signed=True)),
                     rb.colorize(-r['Retard_Cumule'], rb.eur(r['Retard_Cumule']))])
    table = rb.ledger_table(["Mois", "Apport réel", "Cible", "Écart", "Retard cumulé"], rows, num_cols=[1, 2, 3, 4])

    return f"""
    {metrics}
    {rb.section("Historique des apports vs cible", f'<div class="card">{chart}</div>')}
    {rb.section("Discipline d'épargne", f'<div class="card">{table}</div>')}
    """
    
def _beta(y, x, min_obs=30):
    """Bêta de la série y sur la série x (dates communes uniquement) ; None si non calculable."""
    a = pd.concat([y, x], axis=1).dropna()
    if len(a) < min_obs:
        return None
    var_x = a.iloc[:, 1].var()
    if not var_x > 0:
        return None
    return float(a.iloc[:, 0].cov(a.iloc[:, 1]) / var_x)


def build_hedge_payload(pf_returns, raw_prices, bench_returns, val_actions, especes, beta_stress,
                        scenarios, taux_sans_risque):
    """Données du simulateur de couverture (CFD vendeur sur indice), recalculé dans le navigateur.
    Pour chaque instrument : rendements journaliers en EUR sur la fenêtre du portefeuille (null hors
    cotation, jamais de bfill), bêta du portefeuille sur l'instrument (sert à dimensionner le
    notionnel), corrélation (part du risque réellement couvrable) et bêta de l'instrument vs le
    benchmark (sert aux stress tests). Les coûts de portage sont des hypothèses de config."""
    instruments = {}
    for t, cfg in HEDGE_INSTRUMENTS.items():
        if t not in raw_prices.columns:
            continue
        h = rendements_reels(raw_prices[t], ffill_limit=5).reindex(pf_returns.index)
        if h.notna().sum() < 60:
            continue
        beta_pf = _beta(pf_returns, h)
        if beta_pf is None:
            continue
        corr = float(pd.concat([pf_returns, h], axis=1).dropna().corr().iloc[0, 1])
        beta_bench = _beta(h, bench_returns) if len(bench_returns) else None
        instruments[t] = {
            "label": cfg['label'],
            "dividende": cfg['dividende'],
            "returns": [None if pd.isna(v) else float(v) for v in h],
            "beta_pf": beta_pf,
            "corr": corr,
            "beta_bench": beta_bench if beta_bench is not None else 1.0,
        }

    has_bench = len(bench_returns) > 0
    return {
        "dates": pf_returns.index.strftime('%Y-%m-%d').tolist(),
        "pf_returns": [float(v) for v in pf_returns],
        "bench_returns": bench_returns.reindex(pf_returns.index).fillna(0.0).tolist() if has_bench else [],
        "has_bench": has_bench,
        "val_actions": float(val_actions),
        "especes": float(especes),
        "beta_stress": float(beta_stress),
        # Scénarios de baisse + un rallye : c'est en hausse qu'un short coûte et appelle de la marge
        "scenarios": [[name, shock] for name, shock in scenarios] + [["Rallye haussier", 0.20]],
        "taux_sans_risque": taux_sans_risque,
        "taux_reference": HEDGE_TAUX_REFERENCE,
        "spread": HEDGE_SPREAD_FINANCEMENT,
        "leviers": HEDGE_LEVIERS,
        "instruments": instruments,
    }


def build_tab_risque(ctx, taux_sans_risque=TAUX_SANS_RISQUE):
    sim_assets = ctx['sim_assets']
    if not sim_assets:
        return "<div class='note'>Aucun actif valorisé pour calculer les indicateurs de risque.</div>"

    raw_prices = ctx['raw_prices']
    active = [t for t in sim_assets if t in raw_prices.columns]
    if not active:
        return "<div class='note'>Aucune donnée de prix disponible pour ces actifs.</div>"

    cutoff = raw_prices.index.max() - timedelta(days=365)
    window = raw_prices.loc[raw_prices.index >= cutoff]
    # Rendements de chaque ligne sur sa cotation réelle (ffill limité, jamais de bfill), puis
    # restriction à la période commune à toutes les lignes : aucune valeur inventée.
    own_returns = {t: rendements_reels(window[t], ffill_limit=5) for t in active}
    # Une ligne quasi sans cours viderait la période commune : on l'écarte en le signalant
    # (avant, le bfill en faisait un actif à rendement nul, donc artificiellement sans risque).
    exclues = [t for t in active if len(own_returns[t]) < 30]
    exclues_note = note_lignes_exclues(exclues, sim_assets, {t: len(own_returns[t]) for t in exclues}, 30)
    active = [t for t in active if t not in exclues]
    if not active:
        return f"<div class='note'>Historique insuffisant pour calculer un risque pertinent. {exclues_note}</div>"
    daily_returns = pd.concat({t: own_returns[t] for t in active}, axis=1)[active].dropna(how='any')
    if len(daily_returns) < 30:
        return f"<div class='note'>Historique insuffisant pour calculer un risque pertinent. {exclues_note}</div>"

    # Fenêtre effective : plus courte que 12 mois si une ligne est cotée depuis peu
    fenetre_note = f"Fenêtre de calcul effective : {periode_str(daily_returns.index)} ({len(daily_returns)} séances)"
    if daily_returns.index[0] > cutoff + timedelta(days=10):
        limitant = max(active, key=lambda t: window[t].first_valid_index())
        fenetre_note += f", raccourcie car {limitant} n'est coté que depuis le {window[limitant].first_valid_index():%d/%m/%Y}"
    fenetre_note += ". " + exclues_note

    weights = pd.Series(sim_assets)[active] / pd.Series(sim_assets)[active].sum()
    pf_returns = daily_returns.dot(weights)
    annual_return = pf_returns.mean() * 252
    annual_vol = pf_returns.std() * np.sqrt(252)
    sharpe = (annual_return - taux_sans_risque) / annual_vol if annual_vol > 0 else 0.0
    cum = (1 + pf_returns).cumprod()
    dd = (cum - cum.cummax()) / cum.cummax()
    max_dd = dd.min() * 100

    # Downside deviation standard (cible 0) calculée sur TOUS les jours : sqrt(mean(min(r, 0)^2))
    downside_vol = float(np.sqrt(np.mean(np.minimum(pf_returns.values, 0.0) ** 2)) * np.sqrt(252))
    sortino = (annual_return - taux_sans_risque) / downside_vol if downside_vol > 0 else 0.0
    calmar = annual_return / (abs(max_dd) / 100) if max_dd != 0 else 0.0
    
    # Bêta vs benchmark : None (affiché « n/d ») si le benchmark est absent ou inexploitable
    has_bench = TICKER_BENCHMARK in raw_prices.columns and raw_prices[TICKER_BENCHMARK].notna().sum() > 1
    bench_returns = rendements_reels(raw_prices[TICKER_BENCHMARK], ffill_limit=5) if has_bench else pd.Series(dtype=float)
    beta = None
    aligned_returns = pd.concat([pf_returns, bench_returns], axis=1).dropna()
    if len(aligned_returns) > 1:
        cov_matrix = np.cov(aligned_returns.iloc[:, 0], aligned_returns.iloc[:, 1])
        if cov_matrix[1, 1] > 0:
            beta = float(cov_matrix[0, 1] / cov_matrix[1, 1])
    beta_txt = f"{beta:.2f}" if beta is not None else "n/d"
    
    corr_matrix = daily_returns.corr()
    corr_chart = rb.fig_correlation(corr_matrix.columns.tolist(), corr_matrix.values)
    
    control = rb.risk_control_panel("risque-sharpe", annual_return, annual_vol, taux_sans_risque, max_dd)

    # ---- Simulateur "ajout d'une position" (what-if, calculé côté navigateur) ----
    watchlist_source = ctx.get('watchlist_data', [])
    candidate_tickers = [w['Ticker'] for w in watchlist_source if w['Ticker'] not in active]
    candidates_returns = {}
    for t in sorted(set(candidate_tickers)):
        if t not in raw_prices.columns:
            continue
        raw_window = window[t]
        # Exclu s'il a moins de 60 cotations réelles sur la fenêtre utilisée par le simulateur
        if raw_window.loc[raw_window.index >= daily_returns.index[0]].notna().sum() < 60:
            continue
        # Rendement nul seulement À L'INTÉRIEUR de la période de cotation du candidat (jour sans
        # cours) ; avant son introduction en bourse : null (le JS ignore ces dates).
        c_ret = rendements_reels(raw_window, ffill_limit=None).reindex(daily_returns.index)
        first, last = c_ret.first_valid_index(), c_ret.last_valid_index()
        if first is None:
            continue
        inside = (c_ret.index >= first) & (c_ret.index <= last)
        c_ret[inside] = c_ret[inside].fillna(0.0)
        candidates_returns[t] = [None if pd.isna(v) else float(v) for v in c_ret]

    bench_ret_full = bench_returns.reindex(daily_returns.index).fillna(0.0).tolist() \
        if has_bench else [0.0] * len(daily_returns.index)

    current_value = float(pd.Series(sim_assets)[active].sum())
    whatif_payload = {
        "dates": daily_returns.index.strftime('%Y-%m-%d').tolist(),
        "returns_matrix": daily_returns.values.tolist(),
        "weights": weights.tolist(),
        "current_value": current_value,
        "bench_returns": bench_ret_full,
        "has_bench": bool(has_bench),
        "candidates": candidates_returns,
        "taux_sans_risque": taux_sans_risque,
    }
    whatif_html = rb.whatif_control_panel(
        "risque-whatif", whatif_payload, sorted(candidates_returns.keys()),
        taux_sans_risque, current_value
    )
    
    ratios_html = f"""
    <div class="metric-row" style="margin-top:16px; margin-bottom:16px;">
        {rb.metric_card("Ratio de Sortino", f"{sortino:.2f}", "Cible > 1.0", sortino>1)}
        {rb.metric_card("Ratio de Calmar", f"{calmar:.2f}", "Rendement / Max DD", calmar>0.5)}
        {rb.metric_card("Beta vs Benchmark", beta_txt, "Sensibilité marché" if beta is not None else f"Benchmark {TICKER_BENCHMARK} indisponible", True)}
    </div>
    """
    
    chart = rb.fig_drawdown(dd.index.strftime('%Y-%m-%d').tolist(), (dd * 100).tolist())

    val_actions = ctx['df_hist_bourse'].iloc[-1]['Valeur_Actions']
    especes = ctx['df_hist_bourse'].iloc[-1]['Especes']
    val_totale = val_actions + especes
    
    scenarios = [
        ("Crise des Subprimes (2008)", -0.50),
        ("Éclatement Bulle Dotcom (2000)", -0.40),
        ("Krach COVID-19 (Mars 2020)", -0.30),
        ("Correction Tech/Taux (2022)", -0.25)
    ]
    
    # Choc de marché transmis au portefeuille via son bêta, borné à [0.5, 2] pour éviter
    # qu'un bêta mal estimé (fenêtre courte) n'écrase ou n'exagère le scénario. Bêta n/d -> 1.
    beta_stress = float(np.clip(beta, 0.5, 2.0)) if beta is not None else 1.0
    beta_stress_txt = f"β = {beta_stress:.2f}" + ("" if beta is not None else " (bêta n/d)")

    stress_rows = []
    for name, shock in scenarios:
        shock_adj = max(shock * beta_stress, -1.0)   # une action ne perd pas plus de 100 %
        impact_eur = val_actions * shock_adj
        new_val = val_totale + impact_eur
        drop_pct = (impact_eur / val_totale) * 100 if val_totale > 0 else 0

        stress_rows.append([
            name,
            rb.colorize(shock * 100, rb.pct(shock * 100, signed=True)),
            rb.colorize(shock_adj * 100, rb.pct(shock_adj * 100, signed=True)),
            rb.colorize(impact_eur, rb.eur(impact_eur, signed=True)),
            rb.colorize(drop_pct, rb.pct(drop_pct, signed=True)),
            rb.eur(new_val)
        ])

    stress_table = rb.ledger_table(
        ["Scénario Historique", "Choc brut (marché)", f"Choc ajusté ({beta_stress_txt})", "Impact P&L",
         "Drawdown Portefeuille", "Valeur Résiduelle"],
        stress_rows, num_cols=[1, 2, 3, 4, 5]
    )

    # ---- Simulateur de couverture (CFD vendeur sur indice, recalculé côté navigateur) ----
    hedge_payload = build_hedge_payload(pf_returns, raw_prices, bench_returns, val_actions, especes,
                                        beta_stress, scenarios, taux_sans_risque)
    hedge_html = rb.hedge_control_panel("risque-hedge", hedge_payload)

    return f"""
    {control}
    {ratios_html}
    {rb.section("Matrice de corrélation", f'<div class="card">{corr_chart}</div>')}
    {rb.section("Simulateur — Ajout d'une position", f'<div class="card">{whatif_html}</div>')}
    {rb.section("Drawdown historique", f'<div class="card">{chart}</div>')}
    {rb.section("Stress Tests (Chocs Déterministes)", f'<div class="card">{stress_table}</div>')}
    {rb.section("Simulateur — Couverture (CFD vendeur sur indice)", f'<div class="card">{hedge_html}</div>')}
    <div class="note">Le ratio de Sharpe mesure la rentabilité obtenue par unité de risque pris —
    déplacez le curseur pour tester différentes hypothèses de taux sans risque, le ratio se
    recalcule instantanément. Il est calculé en appliquant l'allocation <b>actuelle</b> du
    portefeuille à l'historique des 12 derniers mois : c'est donc la performance qu'aurait eue
    cette allocation si elle avait été détenue toute l'année, pas la performance réellement vécue
    (qui dépend des dates d'achat/vente passées). Le max drawdown est la plus forte chute
    historique subie par cette allocation exacte sur un an. Les Stress Tests appliquent une décote
    immédiate sur la poche actions actuelle, en considérant la poche espèces comme valeur refuge :
    le choc brut du marché est multiplié par le bêta du portefeuille (borné entre 0,5 et 2).
    Le ratio de Sortino utilise la downside deviation (écart quadratique des seuls rendements
    négatifs, moyenné sur tous les jours).
    Le simulateur d'ajout de position recalcule vol/Sharpe/Sortino/Beta sur cette même fenêtre de
    12 mois en mélangeant vos rendements actuels avec ceux du titre choisi dans votre watchlist
    (pondéré par le montant simulé) — tout se recalcule dans le navigateur, aucune donnée n'est renvoyée
    au serveur. Si le titre choisi est coté depuis moins longtemps, la comparaison avant/après est faite
    uniquement sur ses jours de cotation réels. {fenetre_note}</div>
    """


def _fit_gjr_garch_t(returns_pct_series):
    r = returns_pct_series.dropna() * 100
    if len(r) < MC_MIN_OBS_GARCH:
        return None
    try:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            am = arch_model(r, mean='Constant', vol='GARCH', p=1, o=1, q=1, dist='t')
            res = am.fit(disp='off', show_warning=False)
        p = res.params
        nu = float(p['nu'])
        if not np.isfinite(nu) or nu <= 2.05:
            return None
        return dict(
            mu=p['mu'] / 100.0, omega=p['omega'], alpha=p['alpha[1]'],
            gamma=p['gamma[1]'], beta=p['beta[1]'], nu=nu,
            last_h=float(res.conditional_volatility.iloc[-1] ** 2),
            last_eps=float(res.resid.iloc[-1]),
            std_resid=(res.resid / res.conditional_volatility).dropna(),
        )
    except Exception:
        return None


def build_tab_montecarlo(ctx, years=MC_YEARS, n_sims=MC_N_SIMS, seed=MC_SEED, block_size=MC_BLOCK_SIZE):
    sim_assets = ctx['sim_assets']
    if not sim_assets:
        return "<div class='note'>Aucun actif valorisé pour lancer une simulation.</div>"

    raw_prices = ctx['raw_prices']
    tickers = [t for t in sim_assets if t in raw_prices.columns]

    # Rendements de chaque actif sur son PROPRE historique réel : ffill limité aux jours fériés
    # (limit=5), jamais de bfill -> un titre coté récemment n'hérite pas de rendements nuls inventés.
    own_returns = {t: rendements_reels(raw_prices[t], ffill_limit=5) for t in tickers}
    # Une ligne quasi sans cours viderait la période commune : on l'écarte en le signalant.
    exclues = [t for t in tickers if len(own_returns[t]) < 60]
    exclues_note = note_lignes_exclues(exclues, sim_assets, {t: len(own_returns[t]) for t in exclues}, 60)
    exclues_html = f"<div class='note'>{exclues_note}</div>" if exclues_note else ""
    tickers = [t for t in tickers if t not in exclues]
    if not tickers:
        return f"<div class='note'>Aucune donnée de prix exploitable pour ces actifs. {exclues_note}</div>"

    # Période commune à tous les actifs (corrélations + bootstrap) : uniquement des dates réelles.
    common_returns = pd.concat({t: own_returns[t] for t in tickers}, axis=1)[tickers].dropna(how='any')
    n_common = len(common_returns)
    limitant = max(tickers, key=lambda t: own_returns[t].index[0])
    limitant_txt = f"{limitant}, premier rendement disponible le {own_returns[limitant].index[0]:%d/%m/%Y}"

    weights = np.array([sim_assets.get(t, 0) for t in tickers])
    weights = weights / weights.sum() if weights.sum() > 0 else weights
    initial_value = float(ctx['df_hist_bourse'].iloc[-1]['Valeur_Actions'])

    fallback_reason = ""
    if not HAS_ARCH:
        fallback_reason = "librairie 'arch' indisponible"
    else:
        courts = [t for t in tickers if len(own_returns[t]) < MC_MIN_OBS_GARCH]
        if courts:
            detail = ", ".join(f"{t} : {len(own_returns[t])} j" for t in courts)
            fallback_reason = f"historique insuffisant pour un GARCH fiable ({detail} ; {MC_MIN_OBS_GARCH} requis)"

    garch_params = {}
    if not fallback_reason:
        for t in tickers:
            fitted = _fit_gjr_garch_t(own_returns[t])
            if fitted is None:
                fallback_reason = f"le GARCH n'a pas convergé sur {t}"
                break
            garch_params[t] = fitted

    std_resid_df = None
    if not fallback_reason:
        # Corrélation des résidus standardisés : période commune uniquement (dropna).
        std_resid_df = pd.concat({t: garch_params[t]['std_resid'] for t in tickers}, axis=1)[tickers].dropna()
        if len(std_resid_df) < 60:
            fallback_reason = (f"période commune trop courte pour estimer les corrélations "
                               f"({len(std_resid_df)} jours, limitée par {limitant_txt})")

    short_note = ""
    if fallback_reason:
        if n_common < 60 or n_common <= block_size:
            return (f"<div class='note'>Historique commun insuffisant pour une simulation fiable : "
                    f"{n_common} jours de rendements communs à toutes les lignes ({periode_str(common_returns.index)}), "
                    f"minimum ~60 jours. Période limitée par {limitant_txt}.</div>")
        mc_payload = {
            "method": "bootstrap",
            "weights": weights.tolist(),
            "initial_value": initial_value,
            "seed": seed,
            "block_size": block_size,
            "returns_matrix": common_returns.values.tolist(),
        }
        method_note = (f"<b>Repli sur bootstrap par blocs</b> ({fallback_reason}) — "
                        f"blocs de {block_size} jours sur {n_common} jours réels communs à toutes les lignes. "
                        f"Période de calibration : {periode_str(common_returns.index)}.")
        if n_common < block_size * 10:
            short_note = (f"<div class='note'>⚠️ La période commune à toutes les lignes ne compte que "
                          f"{n_common} jours (moins de {block_size * 10}, soit 10 blocs de {block_size} jours) : "
                          f"le bootstrap ré-échantillonne très peu de blocs distincts et les percentiles "
                          f"sont peu fiables. Période limitée par {limitant_txt}.</div>")
    else:
        corr = std_resid_df.corr().values
        corr = np.nan_to_num(corr, nan=0.0)
        np.fill_diagonal(corr, 1.0)
        try:
            L = np.linalg.cholesky(corr)
        except np.linalg.LinAlgError:
            corr = corr + np.eye(len(tickers)) * 1e-6
            L = np.linalg.cholesky(corr)

        nu_joint = float(np.clip(np.mean([garch_params[t]['nu'] for t in tickers]), 3.0, 30.0))

        mc_payload = {
            "method": "garch",
            "tickers": tickers,
            "weights": weights.tolist(),
            "initial_value": initial_value,
            "seed": seed,
            "nu_joint": nu_joint,
            "L": L.tolist(),
            "omega": [garch_params[t]['omega'] for t in tickers],
            "alpha": [garch_params[t]['alpha'] for t in tickers],
            "gamma": [garch_params[t]['gamma'] for t in tickers],
            "beta": [garch_params[t]['beta'] for t in tickers],
            "mu": [garch_params[t]['mu'] for t in tickers],
            "h0": [garch_params[t]['last_h'] for t in tickers],
            "eps0": [garch_params[t]['last_eps'] for t in tickers],
        }
        periodes = "; ".join(f"{t} {periode_str(own_returns[t].index)} ({len(own_returns[t])} j)" for t in tickers)
        method_note = (f"<b>GJR-GARCH(1,1)-t</b> par actif, calibré sur l'historique réel de chaque actif "
                        f"({periodes}), corrélation constante entre actifs (CCC-GARCH) estimée sur la période "
                        f"commune {periode_str(std_resid_df.index)} ({len(std_resid_df)} j), chocs simulés "
                        f"conjointement via une Student-t multivariée (ddl moyen ≈ {nu_joint:.1f}).")

    control = rb.mc_control_panel(
        "mc-projection", mc_payload, default_years=years, default_n_sims=n_sims,
        min_years=1, max_years=15, min_sims=200, max_sims=3000, step_sims=100,
    )

    return f"""
    <div class="note">Horizon et nombre de tirages ajustables ci-dessous — le calcul (quelques
    secondes pour les valeurs élevées) tourne directement dans le navigateur, rien n'est renvoyé
    au serveur. Méthode de simulation : {method_note}</div>
    {exclues_html}
    {short_note}
    {rb.section(f"Projection", control)}
    """

def build_vol_surface_html(ctx):
    """Section « surface de volatilité » : métriques de lecture + surface 3D, smiles et structure
    par terme. Tous les chiffres viennent de la surface calculée dans get_vol_surface."""
    s = ctx.get('vol_surface')
    if not s or not s.get('echeances'):
        return "<div class='note'>Surface de volatilité indisponible (chaînes d'options S&P 500 non récupérées).</div>"

    ech = s['echeances']
    vix = ctx.get('vix_data', {}).get('current')
    atm_1m = iv_surface(s, 1.0, 30)
    iv_90, iv_110 = iv_surface(s, 0.90, 30), iv_surface(s, 1.10, 30)
    skew_1m = iv_90 - iv_110 if iv_90 is not None and iv_110 is not None else None
    atm_1y = iv_surface(s, 1.0, 365)
    pente = atm_1y - atm_1m if atm_1y is not None and atm_1m is not None else None

    def pct_ou_nd(v):
        return rb.pct(v * 100, signed=False) if v is not None else "n/d"

    def pts_ou_nd(v):
        """Écart entre deux volatilités, en points de vol."""
        return f"{v * 100:+.2f} pts" if v is not None else "n/d"

    ecart_vix = (vix / 100 - atm_1m) if (vix is not None and atm_1m is not None) else None
    metrics = f"""<div class="metric-row">
      {rb.metric_card("Vol implicite ATM 1 mois", pct_ou_nd(atm_1m))}
      {rb.metric_card("VIX", f"{vix:.2f}" if vix is not None else "n/d",
                      f"{pts_ou_nd(ecart_vix)} vs ATM (prime de skew)" if ecart_vix is not None else None, True)}
      {rb.metric_card("Skew 1 mois (90 % − 110 %)", pts_ou_nd(skew_1m),
                      "Prime payée pour la protection baissière", True)}
      {rb.metric_card("Pente 1 an − 1 mois (ATM)", pts_ou_nd(pente),
                      ("Contango : marché calme" if pente >= 0 else "Inversion : stress à court terme") if pente is not None else None,
                      pente is None or pente >= 0)}
    </div>"""

    surface_fig = rb.fig_vol_surface(s['moneyness'], [r['jours'] for r in ech], [r['grille_iv'] for r in ech])

    # Smiles : échéances les plus proches de 1, 3, 6 et 12 mois
    retenues = []
    for cible in (30, 90, 180, 365):
        r = min(ech, key=lambda x: abs(x['jours'] - cible))
        if r not in retenues:
            retenues.append(r)
    smiles_fig = rb.fig_vol_smiles([(f"{r['jours']} j ({r['echeance']})", r['smile_m'], r['smile_iv']) for r in retenues])

    # Structure par terme ATM (interpolée dans chaque smile, sans extrapolation)
    terme = [(r['jours'], float(np.interp(1.0, r['smile_m'], r['smile_iv'])))
             for r in ech if r['smile_m'][0] <= 1.0 <= r['smile_m'][-1]]
    term_fig = rb.fig_vol_term([t[0] for t in terme], [t[1] for t in terme], vix / 100 if vix is not None else None)

    part_mid = np.average([r['part_mid'] for r in ech], weights=[r['nb_points'] for r in ech])
    repli = [r['echeance'] for r in ech if r['methode_forward'] != "parité call-put"]
    source_txt = {"reseau": "téléchargées à l'instant", "cache": "issues du cache",
                  "cache_perime": "issues du cache PÉRIMÉ (échec du téléchargement)"}.get(s.get('source'), "")
    note = (f"<div class='note'>Options sur l'indice S&amp;P 500 ({VOL_SURFACE_TICKER}, sous-jacent du VIX), "
            f"séance du {pd.Timestamp(s['session']):%d/%m/%Y}, données {source_txt} "
            f"({pd.Timestamp(s['timestamp']):%d/%m/%Y %H:%M}). {len(ech)} échéances, "
            f"{sum(r['nb_points'] for r in ech)} options hors de la monnaie. Les volatilités sont recalculées "
            f"(Black-76) : forward et actualisation déduits de la parité call-put"
            + (f" (taux de repli pour {', '.join(repli)})" if repli else "")
            + f". Prix utilisés : {part_mid * 100:.0f} % au milieu bid/ask, le reste au dernier prix échangé "
            f"pendant la séance — hors séance américaine, ces derniers prix ne sont pas synchrones et la surface "
            f"est plus bruitée. La grille n'est remplie que dans la plage de strikes réellement cotée (aucune "
            f"extrapolation) ; le VIX est une variance « sans modèle » sur 30 jours qui intègre le skew, d'où "
            f"son écart habituel avec la volatilité ATM.</div>")

    return f"""
    {metrics}
    <div class="card" style="margin-top:16px;"><div class="card-title">Surface de volatilité implicite</div>{surface_fig}</div>
    <div class="grid-2" style="margin-top:20px;">
      <div class="card"><div class="card-title">Smiles par échéance</div>{smiles_fig}</div>
      <div class="card"><div class="card-title">Structure par terme (ATM) et VIX</div>{term_fig}</div>
    </div>
    {note}
    """


def build_tab_actualites(ctx):
    articles = ctx.get('actualites', [])
    agenda = ctx.get('agenda', [])
    meta = ctx.get('news_meta', {})

    sur_lignes = [a for a in articles if any(not c['theme'] for c in a['cibles'])]
    signif = [a for a in articles if a['significative']]
    prochain = agenda[0] if agenda else None
    metrics = f"""<div class="metric-row">
      {rb.metric_card(f"Articles ({NEWS_JOURS} jours)", str(len(articles)),
                      f"{len(sur_lignes)} sur vos lignes · {len(articles) - len(sur_lignes)} thématiques")}
      {rb.metric_card("Réactions de cours significatives", str(len(signif)),
                      f"Écart au marché ≥ {NEWS_SEUIL_SIGMA:.0f} σ le jour de la publication", len(signif) == 0)}
      {rb.metric_card("Prochain événement", f"{prochain['date']:%d/%m}" if prochain else "—",
                      f"{html_lib.escape(prochain['nom'])} · {prochain['evenement'].lower()}" if prochain else
                      f"Aucun dans les {NEWS_AGENDA_JOURS} jours")}
    </div>"""

    if agenda:
        lignes = [[f"{e['date']:%d/%m/%Y}", e['evenement'],
                   html_lib.escape(e['nom']) + f" <span class='tag'>{html_lib.escape(e['ticker'])}</span>",
                   rb.pct(e['poids'], signed=False), "aujourd'hui" if e['jours'] == 0 else f"J-{e['jours']}"]
                  for e in agenda]
        agenda_html = rb.ledger_table(["Date", "Événement", "Ligne", "Poids", "Échéance"], lignes, num_cols=[3, 4])
    else:
        agenda_html = f"<div class='note'>Aucun résultat ni détachement de dividende annoncé dans les {NEWS_AGENDA_JOURS} jours pour vos lignes.</div>"

    # Deux fils séparés, comme sur les terminaux professionnels : nouvelles des sociétés détenues
    # d'un côté, marché et thèmes (exposition des ETF) de l'autre — leurs scores ne se comparent pas.
    def panneau(container_id, selection):
        cibles = {}
        for a in selection:
            for c in a['cibles']:
                cibles.setdefault(c['code'], (c['nom'], c['poids'], c['theme']))
        options = [{"code": k, "label": (v[0] if v[2] else f"{v[0]} ({k})") + f" — {v[1]:.1f} %"}
                   for k, v in sorted(cibles.items(), key=lambda kv: -kv[1][1])]
        types = sorted({t for a in selection for t in a['types']})
        return rb.news_panel(container_id, {"articles": selection, "cibles": options, "types": types,
                                            "seuil": NEWS_SEUIL_SIGMA})

    fil_lignes = panneau("news-lignes", sur_lignes)
    fil_themes = panneau("news-themes", [a for a in articles if a not in sur_lignes])

    etats = meta.get('etats', {})
    etat_txt = ", ".join(f"{n} {lib}" for cle, lib in (("reseau", "à jour"), ("cache", "en cache"),
                                                          ("cache_perime", "en cache PÉRIMÉ"), ("indisponible", "indisponibles"))
                         if (n := etats.get(cle)))
    maj = f", dernière mise à jour {pd.Timestamp(meta['maj']):%d/%m/%Y %H:%M}" if meta.get('maj') else ""
    note = (f"<div class='note'>Sources : flux RSS Yahoo Finance de chaque action détenue (les ETF n'en ont pas) et "
            f"Google Actualités pour les thèmes liés à l'exposition de vos ETF ({meta.get('nb_flux', 0)} flux : {etat_txt}{maj}). "
            f"<b>Pertinence</b> = importance du type d'information × racine du poids de la ligne (rapporté à votre plus "
            f"grosse ligne) × fraîcheur (demi-vie d'environ 3 jours), majorée de 50 % si le cours a réagi. Un article "
            f"du flux d'une ligne qui ne cite pas la société (Yahoo y range aussi ses concurrents) est marqué "
            f"« mention indirecte » et pondéré à {NEWS_FACTEUR_MENTION_INDIRECTE * 100:.0f} % ; aucun mouvement de "
            f"cours ne lui est rattaché. "
            f"<b>Réaction</b> : rendement du titre à la première séance où le marché pouvait réagir (le jour même si "
            f"l'article précède la clôture de la place, sinon la séance suivante), comparé au benchmark {TICKER_BENCHMARK} ; "
            f"significative au-delà de {NEWS_SEUIL_SIGMA:.0f} écarts-types de l'écart quotidien habituel (250 séances). "
            f"Une réaction le même jour ne prouve pas que l'article en est la cause. Types et sentiment sont détectés par "
            f"mots-clés : indicatifs, à vérifier en lisant l'article.</div>")

    return f"""
    {metrics}
    {rb.section("Agenda des lignes détenues", f'<div class="card table-container">{agenda_html}</div>')}
    {rb.section("Fil — vos lignes", f'<div class="card">{fil_lignes}</div>')}
    {rb.section("Fil — marché et thèmes liés à vos ETF", f'<div class="card">{fil_themes}</div>')}
    {note}
    """


def build_tab_marche(ctx):
    vix_data = ctx.get('vix_data', {})
    vix_regime = ctx.get('vix_regime', {})
    watchlist_data = ctx.get('watchlist_data', [])
    vix_seuils = ctx.get('vix_seuils', {"complaisance": 15, "normal": 20, "eleve": 30})

    if not vix_data:
        return "<div class='note'>Données de marché (VIX) indisponibles.</div>"

    badge = rb.regime_badge(vix_regime.get('label', 'Inconnu'), vix_regime.get('comment', ''))
    gauge = rb.fig_vix_gauge(vix_data['current'], vix_seuils)
    dates_vix = vix_data['history'].index.strftime('%Y-%m-%d').tolist()
    hist_vix = rb.fig_vix_history(dates_vix, vix_data['history'].tolist(), vix_data['mm20'].tolist(), vix_seuils)

    vix_html = f'''
    <div class="grid-12">
        <div class="card"><div class="card-title">Niveau Actuel (VIX)</div>{gauge}</div>
        <div class="card"><div class="card-title">Historique & Tendances (2 ans)</div>{hist_vix}</div>
    </div>
    '''

    screener_html = f"""
    <div class="control-panel" style="margin-bottom: 20px;">
      <div class="control">
        <label>PER Maximum</label>
        <input type="range" id="screener-per" min="5" max="30" step="1" value="12">
        <span class="control-value" id="screener-per-val">12</span>
      </div>
      <div class="control">
        <label>ROE Minimum</label>
        <input type="range" id="screener-roe" min="0" max="30" step="1" value="10">
        <span class="control-value" id="screener-roe-val">10%</span>
      </div>
      <div class="control">
        <label>Dette/Equity Maximum</label>
        <input type="range" id="screener-debt" min="0.5" max="5" step="0.5" value="3">
        <span class="control-value" id="screener-debt-val">3</span>
      </div>
      <div class="control">
        <label>Secteur</label>
        <select id="screener-sector">
          <option value="">Tous les secteurs</option>
        </select>
      </div>
    </div>
    <div style="margin-bottom: 12px; font-family: var(--font-mono); font-size: 12px; color: var(--gold);">
        Cibles éligibles dans la watchlist : <span id="screener-count" style="font-size: 14px; font-weight: 600; color: var(--ivory);">0</span>
    </div>
    <div class="table-container">
    <table class="ledger">
        <thead>
            <tr>
                <th>Cible</th>
                <th>Secteur</th>
                <th class="num">PER</th>
                <th class="num">ROE</th>
                <th class="num">Croiss. CA</th>
                <th class="num">Dette/Eq</th>
                <th class="num">Momentum 6M</th>
                <th class="num">Dist. MM200</th>
            </tr>
        </thead>
        <tbody id="screener-tbody">
        </tbody>
    </table>
    </div>
    <script>
        window.__screenerData = {json.dumps(watchlist_data)};
        if(typeof initScreener === 'function') initScreener('screener');
    </script>
    """

    reg_label = vix_regime.get('label', '')
    if reg_label in ["Complaisance", "Normal"]:
        lecture = "Le marché est dans une zone de volatilité gérable. C'est le moment d'utiliser le screener ci-dessus pour chercher des points d'entrée."
    elif reg_label in ["Élevé", "Stress"]:
        lecture = "Le marché traverse une phase de stress ou d'incertitude forte. Même les cibles qui passent vos filtres risquent de subir des secousses corrélées aux indices. Privilégiez des entrées très fractionnées ou patientez."
    else:
        lecture = "Configuration neutre."

    combo_html = f"""
    <div class="card" style="margin-bottom:24px; background:var(--panel-2);">
        <div class="card-title" style="color:var(--ivory);">Lecture du Régime (VIX)</div>
        <div style="line-height:1.6; font-size:13.5px;">
            Régime actuel : <b>{reg_label}</b>.<br><br>
            {lecture}<br><br>
            <i style="color:var(--slate); font-size:11px;">⚠️ Utilisez les curseurs du screener pour filtrer dynamiquement votre univers d'investissement selon vos critères stricts. La croissance du CA doit être positive pour apparaître dans le tableau. Le filtre Secteur permet d'isoler des comparables (peers) pour analyser une cible par rapport à son industrie.</i>
        </div>
    </div>
    """

    return f"""
    {badge}
    {rb.section("Analyse du Régime (Volatilité)", vix_html)}
    {rb.section("Surface de volatilité implicite (S&amp;P 500)", build_vol_surface_html(ctx))}
    {rb.section("Screener Fondamental Interactif", f'<div class="card">{screener_html}</div>')}
    {combo_html}
    """

# =================================================================
# 4. MAIN
# =================================================================

MOIS_FR = ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre",
           "octobre", "novembre", "décembre"]


def main(argv=None):
    global DATA_DIR, OUTPUT_FILE
    parser = argparse.ArgumentParser(description="Génère le rapport HTML du dashboard patrimonial.")
    parser.add_argument("--data-dir", default=DATA_DIR, help="dossier des fichiers Excel d'entrée (défaut : data/demo à côté du script "
                             "s'il existe, sinon le dossier du script)")
    parser.add_argument("--output", default=OUTPUT_FILE, help=f"fichier HTML produit (défaut : {OUTPUT_FILE})")
    args = parser.parse_args(argv)
    DATA_DIR, OUTPUT_FILE = args.data_dir, args.output

    ctx = run_engine()
    tabs = {
        "tab-patrimoine": {"label": "Patrimoine global", "html": build_tab_patrimoine(ctx)},
        "tab-bourse": {"label": "Portefeuille boursier", "html": build_tab_bourse(ctx)},
        "tab-epargne": {"label": "Suivi d'épargne", "html": build_tab_epargne(ctx)},
        "tab-risque": {"label": "Gestion des risques", "html": build_tab_risque(ctx)},
        "tab-simu": {"label": "Prédiction", "html": build_tab_montecarlo(ctx)},
        "tab-marche": {"label": "Analyse des marchés", "html": build_tab_marche(ctx)},
        "tab-actualites": {"label": "Actualités", "html": build_tab_actualites(ctx)},
    }
    meta = {
        "title": "Patrimoine — Relevé privé",
        # Mois en français quelle que soit la locale du système (strftime %B dépend de la locale)
        "date_str": f"{datetime.now():%d} {MOIS_FR[datetime.now().month - 1]} {datetime.now():%Y}",
        "footer": ("Dashboard de Gestion de portefeuille - Propriété de Rayan LEKMINE"),
    }
    alerts = ctx.get('alerts', [])
    html = rb.build_report(tabs, meta, alerts)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"Rapport généré : {OUTPUT_FILE}")

if __name__ == "__main__":
    main()