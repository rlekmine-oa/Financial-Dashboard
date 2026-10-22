"""
test_synthetique.py
-------------------
Tests sans réseau ni fichiers Excel, sur données synthétiques :
1. impôt par enveloppe : compensation des moins-values sur le CTO, PEA / PEA-PME distincts ;
2. aucun rendement nul artificiel avant la cotation d'un titre coté plus récemment
   (Monte Carlo bootstrap et GARCH, onglet risque, simulateur what-if) ;
3. drawdown de l'indice TWR insensible à un dépôt ou à un retrait.

Lancement : python tests/test_synthetique.py   (ou : python -m pytest)
"""
import json
import re
import sys
from pathlib import Path

import numpy as np
import pandas as pd

# Le module testé est à la racine du dépôt
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import generate_report as g  # noqa: E402


def _refuser_constante(c):
    raise ValueError(f"constante JSON non standard dans un payload : {c}")


def extraire_payload(html, var):
    """Récupère le JSON injecté via window.<var>["id"] = {...}; et refuse NaN / Infinity."""
    m = re.search(r'window\.%s\["[^"]+"\] = (\{.*\});\s*$' % re.escape(var), html, re.M)
    assert m, f"payload {var} introuvable"
    return json.loads(m.group(1), parse_constant=_refuser_constante)


# ---------------------------------------------------------------
# Données synthétiques : AAA coté depuis 800 séances, BBB depuis 200, CCC (watchlist) depuis 100
# ---------------------------------------------------------------
rng = np.random.default_rng(0)
N = 800
IDX = pd.bdate_range(end=pd.Timestamp.today().normalize(), periods=N)
DEBUT_B, DEBUT_C, DEBUT_D = IDX[N - 200], IDX[N - 100], IDX[N - 40]


def serie(debut):
    s = pd.Series(100 * np.exp(np.cumsum(rng.normal(0.0003, 0.012, N))), index=IDX)
    s[s.index < debut] = np.nan
    return s


RAW = pd.DataFrame({
    "AAA.PA": serie(IDX[0]),
    "BBB.PA": serie(DEBUT_B),
    "CCC.PA": serie(DEBUT_C),
    "DDD.PA": serie(DEBUT_D),      # candidat trop récent (< 60 jours réels) -> exclu
    "EEE.PA": serie(IDX[N - 10]),  # ligne détenue quasi sans cours -> exclue et signalée
})
JOUR_TROU_C = IDX[N - 50]
RAW.loc[JOUR_TROU_C, "CCC.PA"] = np.nan   # jour sans cours À L'INTÉRIEUR de la cotation de CCC

CTX = {
    "sim_assets": {"AAA.PA": 6000.0, "BBB.PA": 4000.0, "EEE.PA": 500.0},
    "raw_prices": RAW,                       # pas de benchmark : le bêta doit valoir « n/d »
    "watchlist_data": [{"Ticker": "CCC.PA"}, {"Ticker": "DDD.PA"}],
    "df_hist_bourse": pd.DataFrame({"Valeur_Actions": [10000.0], "Especes": [500.0]}),
}


def test_fiscalite():
    df = pd.DataFrame({
        "Compte": ["CTO", "CTO", "PEA", "PEA-PME", "PEA PME"],
        "Plus-Value": [1000.0, -400.0, 500.0, 200.0, -50.0],
        "Valeur": [5000.0, 3000.0, 4000.0, 1000.0, 800.0],
    })
    g.PEA_DATE_OUVERTURE, g.PEA_PME_DATE_OUVERTURE = None, None
    out, total, infos = g.calcul_fiscalite_enveloppes(df, "2026-10-22")
    assert abs(infos["CTO"]["impot"] - 600 * 0.314) < 1e-9, "CTO : la moins-value doit compenser"
    assert out.loc[1, "PFU"] == 0.0, "ligne CTO en perte : aucune fiscalité"
    assert abs(out.loc[0, "PFU"] - 600 * 0.314) < 1e-9, "impôt CTO réparti sur la seule ligne en gain"
    assert set(out["Enveloppe"]) == {"CTO", "PEA", "PEA-PME"}, "PEA-PME distinct du PEA"
    assert abs(infos["PEA-PME"]["impot"] - 150 * 0.314) < 1e-9
    assert infos["PEA"]["pfu_applique"], "date d'ouverture inconnue -> PFU au PEA"
    assert abs(out["PFU"].sum() - total) < 1e-9, "somme des lignes = impôt des enveloppes"

    g.PEA_DATE_OUVERTURE = "2015-01-01"
    _, _, infos = g.calcul_fiscalite_enveloppes(df, "2026-10-22")
    assert infos["PEA"]["taux"] == 0.186 and not infos["PEA"]["pfu_applique"], "PEA > 5 ans -> PS seuls"
    g.PEA_DATE_OUVERTURE = "2023-01-01"
    _, _, infos = g.calcul_fiscalite_enveloppes(df, "2026-10-22")
    assert infos["PEA"]["taux"] == 0.314, "PEA < 5 ans -> PFU"
    g.PEA_DATE_OUVERTURE = None
    print("OK 1. fiscalité par enveloppe (CTO : (1000 - 400) x 31,4 % = 188,40 €)")


def test_montecarlo_bootstrap():
    g.HAS_ARCH = False
    html = g.build_tab_montecarlo(CTX)
    p = extraire_payload(html, "__mcData")
    assert p["method"] == "bootstrap"
    mat = np.array(p["returns_matrix"])
    assert len(mat) == 199, f"période commune attendue : 199 rendements de BBB, obtenu {len(mat)}"
    assert not np.any(mat[:, 1] == 0.0), "rendement nul artificiel dans BBB"
    assert f"du {IDX[N - 199]:%d/%m/%Y} au {IDX[-1]:%d/%m/%Y}" in html, "période de calibration absente"
    assert len(p["weights"]) == 2 and "Lignes exclues" in html and "EEE.PA (9 j" in html, \
        "ligne sans historique : exclue et signalée"

    # Ancienne méthode (ffill + bfill) : ~600 rendements nuls inventés pour BBB
    ancien = RAW[["AAA.PA", "BBB.PA"]].ffill().bfill().pct_change().dropna()
    assert (ancien["BBB.PA"] == 0).sum() >= 599

    html_court = g.build_tab_montecarlo(CTX, block_size=25)    # 199 < 25 x 10 -> note
    assert "La période commune à toutes les lignes ne compte que" in html_court
    print(f"OK 2a. Monte Carlo bootstrap : {len(mat)} jours communs, aucun zéro inventé "
          f"(l'ancienne méthode en créait {(ancien['BBB.PA'] == 0).sum()})")


def test_montecarlo_garch_plomberie():
    """Faux ajustement GARCH (arch n'est pas requis) : vérifie que chaque actif est calibré
    sur son propre historique et que la corrélation n'utilise que la période commune."""
    appels = {}

    def faux_fit(r):
        appels[len(appels)] = (r.index[0], len(r))
        z = (r - r.mean()) / r.std()
        return dict(mu=float(r.mean()), omega=0.01, alpha=0.05, gamma=0.05, beta=0.9, nu=6.0,
                    last_h=1.0, last_eps=0.0, std_resid=z)

    vrai_fit, vrai_has, vrai_min = g._fit_gjr_garch_t, g.HAS_ARCH, g.MC_MIN_OBS_GARCH
    g._fit_gjr_garch_t, g.HAS_ARCH, g.MC_MIN_OBS_GARCH = faux_fit, True, 150
    try:
        html = g.build_tab_montecarlo(CTX)
    finally:
        g._fit_gjr_garch_t, g.HAS_ARCH, g.MC_MIN_OBS_GARCH = vrai_fit, vrai_has, vrai_min
    p = extraire_payload(html, "__mcData")
    assert p["method"] == "garch"
    assert appels[0] == (IDX[1], N - 1), "AAA calibré sur tout son historique réel"
    assert appels[1] == (IDX[N - 199], 199), "BBB calibré uniquement depuis sa cotation"
    assert f"commune du {IDX[N - 199]:%d/%m/%Y}" in html, "corrélation : période commune absente"
    print("OK 2b. GARCH : calibration par actif (799 j / 199 j), corrélation sur la période commune")


def test_risque_whatif():
    html = g.build_tab_risque(CTX)
    assert "n/d" in html, "bêta « n/d » attendu sans benchmark"
    p = extraire_payload(html, "__whatifData")
    assert p["has_bench"] is False
    assert len(p["weights"]) == 2 and "EEE.PA (9 j" in html, "ligne sans historique : exclue et signalée"
    dates = pd.to_datetime(p["dates"])
    assert dates[0] == IDX[N - 199], "matrice de risque restreinte à la période commune"
    assert "DDD.PA" not in p["candidates"], "candidat < 60 jours réels doit être exclu"
    c = pd.Series(p["candidates"]["CCC.PA"], index=dates, dtype=object)
    avant = c[dates < DEBUT_C]
    pendant = c[dates > DEBUT_C]
    assert avant.isna().all() and len(avant) > 0, "avant cotation : null, pas de zéro"
    assert pendant.notna().all(), "pendant la cotation : une valeur chaque jour"
    assert c[JOUR_TROU_C] == 0.0, "jour sans cours à l'intérieur de la cotation : rendement nul"
    print(f"OK 2c. What-if : {len(avant)} dates null avant la cotation de CCC, "
          f"0 seulement le jour sans cours ; DDD exclu ; bêta n/d sans benchmark")


def test_drawdown_twr():
    jours = pd.date_range("2026-01-01", periods=30, freq="D")
    marche = np.ones(30)
    marche[10:16] = np.linspace(1.0, 0.8, 6)       # baisse de marché de 20 %
    marche[16:] = 0.8

    def hist(flux_j20):
        investi = np.full(30, 1000.0)
        investi[20:] += flux_j20
        valeur = 1000.0 * marche
        valeur[20:] += flux_j20                     # le flux arrive en espèces, sans rendement
        return pd.DataFrame({"Valeur_Bourse": valeur, "Investi_Bourse": investi}, index=jours)

    res = {}
    for nom, flux in [("sans flux", 0.0), ("dépôt +1000", 1000.0), ("retrait -500", -500.0)]:
        h = hist(flux)
        twr_dd = g.drawdown_courant_twr(g.compute_twr(h))
        v = h["Valeur_Bourse"]
        ancien_dd = (v.iloc[-1] / v.cummax().iloc[-1] - 1) * 100
        res[nom] = (twr_dd, ancien_dd)
        assert abs(twr_dd - (-20.0)) < 1e-9, f"{nom} : drawdown TWR attendu -20 %, obtenu {twr_dd:.2f}"
    print("OK 3. drawdown TWR = -20 % dans les trois cas (ancienne méthode sur la valeur : "
          + ", ".join(f"{k} {v[1]:+.1f} %" for k, v in res.items()) + ")")


def test_validation_retraits():
    ops = pd.DataFrame({
        "Date": pd.to_datetime(["2026-01-05", "2026-02-05", "2026-03-05", "2026-04-05"]),
        "Sens": ["Virement", "Retrait", "Retrait", "Achat"],
        "Compte": ["PEA", "CTO", "PEA", "CTO"],
        "Montant": [500.0, -200.0, 300.0, 100.0],
    })
    msgs = g.valider_retraits(ops)
    assert len(msgs) == 1, "seul le retrait positif doit être signalé"
    assert "05/03/2026" in msgs[0] and "(PEA)" in msgs[0]
    print("OK 4. validation des retraits : seul le « Retrait » à montant positif est signalé")


def test_couverture():
    """Portefeuille = 0,8 x indice + bruit : bêta ≈ 0,8 ; un indice coté en cours de fenêtre
    reçoit null avant sa cotation (pas de zéro inventé)."""
    idx = IDX[-260:]
    r_ind = pd.Series(rng.normal(0, 0.01, len(idx)), index=idx)
    pf = 0.8 * r_ind + pd.Series(rng.normal(0, 0.002, len(idx)), index=idx)
    prix = pd.DataFrame({
        "^IND": 100 * (1 + r_ind).cumprod(),
        "^RECENT": 100 * (1 + r_ind).cumprod().where(idx >= idx[100]),
    })
    vrai = g.HEDGE_INSTRUMENTS
    g.HEDGE_INSTRUMENTS = {
        "^IND": {"label": "Indice", "devise": "EUR", "dividende": 0.03},
        "^RECENT": {"label": "Récent", "devise": "EUR", "dividende": 0.0},
        "^ABSENT": {"label": "Absent", "devise": "EUR", "dividende": 0.0},
    }
    try:
        p = g.build_hedge_payload(pf.iloc[1:], prix, pd.Series(dtype=float), 10000.0, 1000.0, 1.0,
                                  [("Krach", -0.30)], 0.039)
    finally:
        g.HEDGE_INSTRUMENTS = vrai
    json.loads(json.dumps(p), parse_constant=_refuser_constante)
    ind = p["instruments"]["^IND"]
    assert abs(ind["beta_pf"] - 0.8) < 0.05 and ind["corr"] > 0.9, ind["beta_pf"]
    assert "^ABSENT" not in p["instruments"], "instrument sans cours : ignoré"
    rec = pd.Series(p["instruments"]["^RECENT"]["returns"], dtype=object)
    assert rec.iloc[:99].isna().all() and rec.iloc[100:].notna().all(), "null avant cotation"
    assert p["scenarios"][-1][1] > 0, "un scénario de hausse doit tester le coût du short"
    assert p["has_bench"] is False and p["bench_returns"] == []
    print(f"OK 5. couverture : bêta portefeuille/indice {ind['beta_pf']:.2f} (attendu 0,80), "
          f"indice récent null avant cotation, instrument absent ignoré")


def test_surface_volatilite():
    """Chaînes générées avec un smile connu (skew + pente) : les IV recalculées doivent le
    retrouver, le forward et l'actualisation être déduits de la parité call-put, un prix aberrant
    et un prix périmé être écartés, et la grille rester vide hors de la plage cotée."""
    import math
    session, S = pd.Timestamp("2026-10-21"), 5000.0

    def vol_vraie(m, T):
        return 0.18 - 0.25 * np.log(m) + 0.5 * np.log(m) ** 2 + 0.02 * T

    chaines, params = [], {}
    for jours in (30, 91, 365):
        T = jours / 365
        DF, F = math.exp(-0.04 * T), S * math.exp(0.025 * T)
        params[jours] = (F, DF)
        K = np.arange(3800.0, 6201.0, 25.0)
        sig = vol_vraie(K / F, T)
        frames = []
        for is_call in (True, False):
            prix = g.prix_black76(F, K, T, DF, sig, is_call)
            frames.append(pd.DataFrame({
                "strike": K, "bid": prix * 0.99, "ask": prix * 1.01, "lastPrice": prix,
                "lastTradeDate": pd.Timestamp("2026-10-21 19:00", tz="UTC")}))
        calls, puts = frames
        if jours == 30:
            i_aberrant = int(np.argmin(np.abs(K - 4600)))          # put hors de la monnaie
            puts.loc[i_aberrant, ["bid", "ask"]] *= 1.6
            i_perime = int(np.argmin(np.abs(K - 4500)))            # dernier prix d'une ancienne séance
            puts.loc[i_perime, ["bid", "ask"]] = 0.0
            puts.loc[i_perime, "lastPrice"] *= 3
            puts.loc[i_perime, "lastTradeDate"] = pd.Timestamp("2026-10-15 19:00", tz="UTC")
            i_last = int(np.argmin(np.abs(K - 4400)))              # dernier prix de la séance : gardé
            puts.loc[i_last, ["bid", "ask"]] = 0.0
        chaines.append((session + pd.Timedelta(days=jours), calls, puts))

    grille_config = g.VOL_SURFACE_GRILLE_MONEYNESS
    g.VOL_SURFACE_GRILLE_MONEYNESS = (0.70, 1.30, 0.025)   # plus large que les strikes générés
    try:
        s = g.surface_depuis_chaines(chaines, S, session)
    finally:
        g.VOL_SURFACE_GRILLE_MONEYNESS = grille_config
    json.loads(json.dumps(s), parse_constant=_refuser_constante)
    assert [r["jours"] for r in s["echeances"]] == [30, 91, 365]
    erreur_max = 0.0
    for r in s["echeances"]:
        F, DF = params[r["jours"]]
        assert r["methode_forward"] == "parité call-put"
        assert abs(r["forward"] / F - 1) < 1e-6 and abs(r["actualisation"] - DF) < 1e-6
        m = np.array(r["smile_m"])
        erreur_max = max(erreur_max, np.max(np.abs(np.array(r["smile_iv"]) - vol_vraie(m, r["jours"] / 365))))
        assert r["grille_iv"][0] is None and r["grille_iv"][-1] is None, "pas d'extrapolation hors plage"
    assert erreur_max < 1e-4, f"IV mal reconstituée (erreur max {erreur_max:.2e})"

    r30 = s["echeances"][0]
    F30 = params[30][0]
    ms = np.round(np.array(r30["smile_m"]) * F30)
    assert 4600 not in ms, "prix aberrant non filtré"
    assert 4500 not in ms, "prix d'une ancienne séance utilisé"
    assert 4400 in ms and r30["part_mid"] < 1, "dernier prix de la séance non utilisé"

    v = g.iv_surface(s, 1.0, 60)
    v30, v91 = g.iv_surface(s, 1.0, 30), g.iv_surface(s, 1.0, 91)
    assert v is not None and min(v30, v91) <= v <= max(v30, v91)
    assert g.iv_surface(s, 1.0, 10) is None, "pas d'extrapolation avant la première échéance"
    print(f"OK 6. surface de volatilité : IV retrouvées à {erreur_max:.1e} près, forward et "
          f"actualisation exacts, prix aberrant et prix périmé écartés")


def test_actualites():
    import report_builder as rb

    # Parsing RSS : lien javascript rejeté, suffixe Google retiré, HTML du résumé nettoyé
    xml = b"""<?xml version="1.0"?><rss><channel>
      <item><title>Airbus rel\xc3\xa8ve ses objectifs - Les Echos</title><link>https://ex.com/a</link>
        <pubDate>Fri, 16 Oct 2026 16:30:00 +0000</pubDate><source>Les Echos</source>
        <description>&lt;b&gt;R\xc3\xa9sultats&lt;/b&gt; sup\xc3\xa9rieurs aux attentes</description></item>
      <item><title>Pi\xc3\xa8ge</title><link>javascript:alert(1)</link><pubDate>Fri, 16 Oct 2026 10:00:00 +0000</pubDate></item>
      <item><title>Sans date</title><link>https://ex.com/b</link></item>
    </channel></rss>"""
    arts = g.parser_rss(xml)
    assert len(arts) == 1 and arts[0]["titre"] == "Airbus relève ses objectifs" and arts[0]["source"] == "Les Echos"
    assert arts[0]["resume"] == "Résultats supérieurs aux attentes"

    types, imp, sent = g.classer_article("Airbus relève ses objectifs annuels après des résultats supérieurs aux attentes")
    assert "Résultats" in types and "Prévisions" in types and imp == 1.0 and sent > 0.2
    _, _, sent = g.classer_article("Le titre chute après une enquête et un avertissement sur résultats")
    assert sent < -0.2

    # Séance de réaction : vendredi 18 h à Paris -> lundi ; mardi 10 h -> mardi ; New York selon 16 h
    seances = pd.bdate_range("2026-10-12", "2026-10-23")
    assert g.date_seance_reaction("2026-10-16T16:30:00+00:00", "AIR.PA", seances) == pd.Timestamp("2026-10-19")
    assert g.date_seance_reaction("2026-10-13T08:00:00+00:00", "AIR.PA", seances) == pd.Timestamp("2026-10-13")
    assert g.date_seance_reaction("2026-10-13T19:00:00+00:00", "MA", seances) == pd.Timestamp("2026-10-13")
    assert g.date_seance_reaction("2026-10-13T21:00:00+00:00", "MA", seances) == pd.Timestamp("2026-10-14")
    assert g.date_seance_reaction("2026-10-23T20:00:00+00:00", "AIR.PA", seances) is None, "séance à venir"

    # Réaction significative : AAA.PA prend +6 % un jour où le benchmark est plat
    idx = pd.bdate_range(end="2026-10-20", periods=300)
    r_b = pd.Series(rng.normal(0, 0.008, len(idx)), index=idx)
    r_a = r_b + pd.Series(rng.normal(0, 0.008, len(idx)), index=idx)
    jour_choc = idx[-3]
    r_b[jour_choc], r_a[jour_choc] = 0.0, 0.06
    prix = pd.DataFrame({g.TICKER_BENCHMARK: 100 * (1 + r_b).cumprod(), "AAA.PA": 100 * (1 + r_a).cumprod(),
                         "BBB.PA": 100 * (1 + r_b).cumprod()})
    pub = (jour_choc + pd.Timedelta(hours=8)).tz_localize("UTC").isoformat()
    ancien = (idx[-6] + pd.Timedelta(hours=8)).tz_localize("UTC").isoformat()   # ~10 jours, dans la fenêtre
    flux = {
        "T:AAA.PA": {"articles": [
            {"titre": "AAA bondit après une acquisition </script><script>alert(1)</script>", "lien": "https://ex.com/1",
             "resume": "", "source": "S", "date": pub},
            {"titre": "Note sans importance", "lien": "https://ex.com/2", "resume": "", "source": "S", "date": ancien},
            # Concurrent rangé dans le flux de AAA le jour du choc : ni cité, ni rattaché au mouvement
            {"titre": "Boeing obtient une certification", "lien": "https://ex.com/5", "resume": "",
             "source": "S", "date": pub}]},
        "T:BBB.PA": {"articles": [   # même article, autre ligne -> fusionné
            {"titre": "AAA bondit après une acquisition </script><script>alert(1)</script>", "lien": "https://ex.com/1",
             "resume": "", "source": "S", "date": pub}]},
        "H:Monde": {"articles": [{"titre": "Wall Street termine en hausse", "lien": "https://ex.com/3",
                                  "resume": "", "source": "S", "date": pub}]},
        "H:Or": {"articles": [{"titre": "Le prix de l'or recule", "lien": "https://ex.com/4",
                               "resume": "", "source": "S", "date": pub}]},
    }
    themes = {"Monde": {"requete": "", "tickers": "*"}, "Or": {"requete": "", "tickers": ["GOLD.PA"]}}
    res = g.analyser_actualites(flux, {"AAA.PA": 0.6, "BBB.PA": 0.4}, {"AAA.PA": "AAA", "BBB.PA": "BBB"},
                                prix, "2026-10-21T12:00:00+00:00", themes=themes)
    titres = [a["titre"] for a in res]
    assert len(res) == 4 and "Le prix de l'or recule" not in titres, "thème sans ligne détenue exclu"
    boeing = next(a for a in res if a["titre"].startswith("Boeing"))
    assert boeing["mention"] == "indirecte" and not boeing["significative"], "mention indirecte non rattachée au choc"
    assert boeing["reaction"]["significative"], "le cours a pourtant bougé ce jour-là"

    # Variantes de nom : article, suffixe juridique, accents, code en majuscules
    assert g.mention_directe("L'Air Liquide relève ses objectifs", g.alias_societe(["L'Air Liquide S.A."]), "AI.PA")
    assert g.mention_directe("CREDIT AGRICOLE SA : résultats", g.alias_societe(["Crédit agricole SA"]), "ACA.PA")
    assert g.mention_directe("DBV dépose sa demande", g.alias_societe(["DBV Technologies"]), "DBV.PA")
    assert not g.mention_directe("Le crédit immobilier recule", g.alias_societe(["Crédit agricole SA"]), "ACA.PA")
    choc = res[0]
    assert choc["titre"].startswith("AAA bondit") and len(choc["cibles"]) == 2, "dédoublonnage multi-lignes"
    assert choc["significative"] and choc["reaction"]["z"] > 4 and abs(choc["reaction"]["rendement"] - 0.06) < 1e-9
    macro = next(a for a in res if a["titre"].startswith("Wall Street"))
    assert macro["importance"] <= 0.5 and macro["reaction"] is None, "article de thème : impact indirect"

    html = rb.news_panel("t", {"articles": res, "cibles": [], "types": [], "seuil": 2})
    assert "</script><script>alert" not in html, "titre externe non échappé dans le <script>"
    print(f"OK 7. actualités : parsing sûr, classification, séance de réaction, choc détecté "
          f"({choc['reaction']['z']:.1f} σ), dédoublonnage, thème sans exposition exclu, titre piégé neutralisé")


if __name__ == "__main__":
    test_validation_retraits()
    test_actualites()
    test_surface_volatilite()
    test_couverture()
    test_fiscalite()
    test_montecarlo_bootstrap()
    test_montecarlo_garch_plomberie()
    test_risque_whatif()
    test_drawdown_twr()
    print("Tous les tests passent.")
