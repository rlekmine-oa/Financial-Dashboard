"""
report_builder.py
------------------
Moteur de rendu HTML pour le dashboard patrimonial.
Ne contient AUCUNE logique de calcul financier : uniquement la mise en forme.
"""

import json
import html as html_lib
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from datetime import datetime

# ---------------------------------------------------------------
# 1. PALETTE / TYPO
# ---------------------------------------------------------------
INK    = "#0E1A2B"
PANEL  = "#15243B"
LINE   = "#2B3D59"
IVORY  = "#EFE9DB"
SLATE  = "#8B96A8"
GOLD   = "#C9A15E"
SAGE   = "#7BAE95"
BRICK  = "#B15B45"
FONT_SANS = "IBM Plex Sans, sans-serif"
FONT_MONO = "IBM Plex Mono, monospace"

NAVY_TEMPLATE = go.layout.Template(
    layout=go.Layout(
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family=FONT_SANS, color=SLATE, size=11),
        colorway=[GOLD, SAGE, "#5A7CA6", BRICK, SLATE, "#E3C384", "#9B6A6C", "#4A6C6F"],
        xaxis=dict(gridcolor=LINE, zeroline=False, tickfont=dict(family=FONT_MONO, size=10)),
        yaxis=dict(gridcolor=LINE, zeroline=False, tickfont=dict(family=FONT_MONO, size=10)),
        legend=dict(font=dict(color=IVORY, size=10), orientation="h", y=1.14, x=0),
        margin=dict(l=50, r=20, t=30, b=40),
        hoverlabel=dict(bgcolor=PANEL, bordercolor=LINE, font=dict(family=FONT_MONO, color=IVORY)),
    )
)

CHART_CONFIG = {"displayModeBar": False, "responsive": True}

def _fig_to_div(fig: go.Figure, height: int = 360) -> str:
    fig.update_layout(template=NAVY_TEMPLATE, height=height)
    return fig.to_html(full_html=False, include_plotlyjs=False, config=CHART_CONFIG)

# ---------------------------------------------------------------
# 2. FIGURES
# ---------------------------------------------------------------

def fig_monthly_stack(months, banque, bourse):
    fig = go.Figure()
    fig.add_bar(x=months, y=banque, name="Banque", marker_color=SAGE, opacity=.85)
    fig.add_bar(x=months, y=bourse, name="Bourse", marker_color=GOLD, opacity=.9)
    fig.update_layout(barmode="stack", yaxis_tickprefix="€", xaxis_tickangle=-45)
    return _fig_to_div(fig, 380)

def fig_donut(values, labels, colors=None, center_text=""):
    if colors is None:
        colors = [GOLD, SAGE, "#5A7CA6", BRICK, SLATE, "#E3C384", "#9B6A6C", "#4A6C6F", "#627B5D", "#B38C4F"] * 3
    fig = go.Figure(go.Pie(
        values=values, labels=labels, hole=.62,
        marker=dict(colors=colors),
        textinfo="percent", textposition="inside",
        textfont=dict(family=FONT_MONO, color=INK, size=12),
    ))
    fig.update_layout(
        showlegend=True, legend=dict(orientation="h", y=-.12, font=dict(color=SLATE, size=10)),
        annotations=[dict(text=center_text, showarrow=False, font=dict(family=FONT_MONO, size=15, color=IVORY))],
    )
    return _fig_to_div(fig, 280)

def fig_line_vs_invested(dates, valeur, investi):
    fig = go.Figure()
    fig.add_scatter(x=dates, y=valeur, name="Valeur du compte", mode="lines", line=dict(color=GOLD, width=2))
    fig.add_scatter(x=dates, y=investi, name="Apports nets", mode="lines", line=dict(color=SLATE, width=1.5, dash="dot"))
    fig.update_layout(yaxis_tickprefix="€")
    return _fig_to_div(fig, 320)

def fig_pv_area(dates, pv_latente, pv_totale):
    fig = go.Figure()
    fig.add_scatter(x=dates, y=pv_latente, name="PV latente (en cours)", fill="tozeroy", line=dict(color="#5A7CA6", width=1))
    fig.add_scatter(x=dates, y=pv_totale, name="PV totale (ventes + dividendes)", fill="tonexty", line=dict(color=SAGE, width=1))
    fig.update_layout(yaxis_tickprefix="€")
    return _fig_to_div(fig, 320)

def fig_correlation(tickers, matrix):
    fig = px.imshow(
        matrix, x=tickers, y=tickers, text_auto=".2f", zmin=-1, zmax=1,
        color_continuous_scale=[[0, BRICK], [0.5, INK], [1, SAGE]],
    )
    fig.update_layout(coloraxis_showscale=False)
    fig.update_traces(textfont=dict(family=FONT_MONO, size=11))
    dynamic_height = max(450, len(tickers) * 35 + 150)
    return _fig_to_div(fig, dynamic_height)

def fig_returns_heatmap(pivot_df):
    text_matrix = [[f"{v*100:+.1f}%" if pd.notna(v) else "" for v in row] for row in pivot_df.values]
    fig = go.Figure(data=go.Heatmap(
        z=pivot_df.values, x=pivot_df.columns, y=pivot_df.index,
        colorscale=[[0, BRICK], [0.5, INK], [1, SAGE]], zmid=0, 
        text=text_matrix, texttemplate="%{text}", hoverinfo="skip", showscale=False, xgap=2, ygap=2
    ))
    fig.update_layout(
        yaxis=dict(autorange="reversed", type="category", gridcolor="rgba(0,0,0,0)"),
        xaxis=dict(type="category", gridcolor="rgba(0,0,0,0)", side="top"), margin=dict(l=40, r=20, t=30, b=10)
    )
    return _fig_to_div(fig, max(180, len(pivot_df)*35 + 80))

def fig_savings_bar_line(months, apport, cible):
    fig = go.Figure()
    fig.add_bar(x=months, y=apport, name="Apport réel", marker_color=GOLD)
    fig.add_scatter(x=months, y=cible, name="Cible mensuelle", mode="lines", line=dict(color=BRICK, width=2, dash="dash"))
    fig.update_layout(yaxis_tickprefix="€", xaxis_tickangle=-45)
    return _fig_to_div(fig, 340)

def fig_drawdown(dates, drawdown_pct):
    fig = go.Figure()
    fig.add_scatter(x=dates, y=drawdown_pct, fill="tozeroy", line=dict(color=BRICK, width=1))
    fig.update_layout(yaxis_ticksuffix="%")
    return _fig_to_div(fig, 300)

def fig_montecarlo_fan(dates, median, p10, p90, p25=None, p75=None):
    fig = go.Figure()
    fig.add_scatter(x=list(dates) + list(dates[::-1]), y=list(p90) + list(p10[::-1]),
                     fill="toself", fillcolor="rgba(201,161,94,0.15)", line=dict(color="rgba(0,0,0,0)"), name="Intervalle 80%")
    if p25 is not None and p75 is not None:
        fig.add_scatter(x=list(dates) + list(dates[::-1]), y=list(p75) + list(p25[::-1]),
                         fill="toself", fillcolor="rgba(201,161,94,0.28)", line=dict(color="rgba(0,0,0,0)"), name="Intervalle 50%")
    fig.add_scatter(x=dates, y=median, mode="lines", name="Scénario médian", line=dict(color=IVORY, width=2))
    fig.update_layout(yaxis_tickprefix="€")
    return _fig_to_div(fig, 380)

def fig_sparkline(months, series):
    fig = go.Figure(go.Scatter(
        x=months, y=series, mode="lines", line=dict(color=GOLD, width=1.5, shape="spline"),
        fill="tozeroy", fillcolor="rgba(201,161,94,0.06)", hoverinfo="skip",
    ))
    fig.update_layout(margin=dict(l=0, r=0, t=0, b=0), xaxis_visible=False, yaxis_visible=False, showlegend=False)
    return _fig_to_div(fig, 120)

def fig_vix_gauge(val, seuils):
    fig = go.Figure(go.Indicator(
        mode="gauge+number", value=val, domain={'x': [0, 1], 'y': [0, 1]},
        gauge={
            'axis': {'range': [None, max(50, val + 5)], 'tickwidth': 1, 'tickcolor': SLATE, 'tickfont': dict(family=FONT_MONO, size=10, color=SLATE)},
            'bar': {'color': IVORY, 'thickness': 0.3}, 'bgcolor': PANEL, 'borderwidth': 0,
            'steps': [
                {'range': [0, seuils['complaisance']], 'color': "rgba(123,174,149,0.3)"},
                {'range': [seuils['complaisance'], seuils['normal']], 'color': "rgba(201,161,94,0.3)"},
                {'range': [seuils['normal'], seuils['eleve']], 'color': "rgba(177,91,69,0.3)"},
                {'range': [seuils['eleve'], 100], 'color': "rgba(177,91,69,0.7)"}],
        }
    ))
    fig.update_layout(height=280, margin=dict(l=20, r=20, t=30, b=20), paper_bgcolor="rgba(0,0,0,0)", font=dict(color=IVORY, family=FONT_MONO))
    return fig.to_html(full_html=False, include_plotlyjs=False, config=CHART_CONFIG)

def fig_vix_history(dates, vix, mm20, seuils):
    fig = go.Figure()
    fig.add_hrect(y0=0, y1=seuils['complaisance'], fillcolor=SAGE, opacity=0.05, line_width=0)
    fig.add_hrect(y0=seuils['complaisance'], y1=seuils['normal'], fillcolor=GOLD, opacity=0.05, line_width=0)
    fig.add_hrect(y0=seuils['normal'], y1=seuils['eleve'], fillcolor=BRICK, opacity=0.05, line_width=0)
    fig.add_hrect(y0=seuils['eleve'], y1=100, fillcolor=BRICK, opacity=0.15, line_width=0)
    fig.add_scatter(x=dates, y=vix, name="VIX", line=dict(color=IVORY, width=1.5))
    fig.add_scatter(x=dates, y=mm20, name="MM20", line=dict(color=GOLD, width=1.5, dash="dot"))
    fig.update_layout(yaxis_title="Niveau VIX", height=280, margin=dict(l=40, r=20, t=30, b=30))
    fig.update_yaxes(range=[min(10, min(vix)*0.9), max(max(vix)*1.1, seuils['eleve']+5)])
    return _fig_to_div(fig, 280)

VOL_COLORSCALE = [[0, SAGE], [0.5, GOLD], [1, BRICK]]

def fig_vol_surface(moneyness, jours, iv_rows):
    z = np.array([[np.nan if v is None else v * 100 for v in row] for row in iv_rows], dtype=float)
    fig = go.Figure(go.Surface(
        x=[m * 100 for m in moneyness], y=jours, z=z, colorscale=VOL_COLORSCALE,
        connectgaps=False, opacity=0.95,
        hovertemplate="K/F %{x:.1f} %<br>%{y} jours<br>IV %{z:.1f} %<extra></extra>",
        colorbar=dict(thickness=10, len=0.6, ticksuffix=" %", tickfont=dict(family=FONT_MONO, size=10, color=SLATE)),
    ))
    axe = dict(gridcolor=LINE, backgroundcolor=PANEL, showbackground=True, color=SLATE,
               tickfont=dict(family=FONT_MONO, size=10))
    fig.update_layout(
        scene=dict(
            xaxis=dict(axe, title=dict(text="Moneyness K/F (%)", font=dict(size=11))),
            yaxis=dict(axe, title=dict(text="Maturité (jours)", font=dict(size=11))),
            zaxis=dict(axe, title=dict(text="Vol implicite (%)", font=dict(size=11)), ticksuffix=" %"),
            camera=dict(eye=dict(x=1.55, y=-1.65, z=0.75)),
        ),
        margin=dict(l=0, r=0, t=10, b=0),
    )
    return _fig_to_div(fig, 540)

def fig_vol_smiles(series):
    """series : liste de (libellé, moneyness, iv)."""
    fig = go.Figure()
    couleurs = [SAGE, GOLD, "#5A7CA6", BRICK]
    for i, (label, m, iv) in enumerate(series):
        fig.add_scatter(x=[v * 100 for v in m], y=[v * 100 for v in iv], name=label, mode="lines+markers",
                        line=dict(color=couleurs[i % len(couleurs)], width=1.5), marker=dict(size=3))
    fig.update_layout(xaxis_title="Moneyness K/F (%)", xaxis_ticksuffix=" %", yaxis_ticksuffix=" %",
                      xaxis_range=[70, 130])
    return _fig_to_div(fig, 340)

def fig_vol_term(jours, atm_iv, vix=None):
    fig = go.Figure()
    fig.add_scatter(x=jours, y=[v * 100 for v in atm_iv], name="Vol ATM", mode="lines+markers",
                    line=dict(color=GOLD, width=2), marker=dict(size=6))
    if vix is not None:
        fig.add_scatter(x=[30], y=[vix * 100], name="VIX (30 j)", mode="markers",
                        marker=dict(color=BRICK, size=11, symbol="diamond"))
    fig.update_layout(xaxis_title="Maturité (jours)", yaxis_ticksuffix=" %")
    return _fig_to_div(fig, 340)

# ---------------------------------------------------------------
# 3. PETITS COMPOSANTS HTML
# ---------------------------------------------------------------

def eur(x, signed=False):
    s = f"{x:,.2f}".replace(",", " ").replace(".", ",")
    if signed and x >= 0: s = "+" + s
    return f"{s} €"

def pct(x, signed=True):
    s = f"{x:+.2f}" if signed else f"{x:.2f}"
    return f"{s} %"

def stat_chip(label, value, sub=None):
    sub_html = f"<small>{sub}</small>" if sub else ""
    return f"""<div class="stat-chip"><div class="k">{label}</div><div class="v">{value}{sub_html}</div></div>"""

def metric_card(label, value, delta=None, positive=True):
    cls = "pos" if positive else "neg"
    delta_html = f'<div class="metric-delta {cls}">{delta}</div>' if delta else ""
    return f"""<div class="metric-card"><div class="k">{label}</div><div class="v">{value}</div>{delta_html}</div>"""

def regime_badge(label, comment):
    color_bg = "rgba(123,174,149,.15)" if label == "Complaisance" else ("rgba(177,91,69,.15)" if label in ["Élevé", "Stress"] else "rgba(201,161,94,.15)")
    color_fg = SAGE if label == "Complaisance" else (BRICK if label in ["Élevé", "Stress"] else GOLD)
    return f"""
    <div style="background:{color_bg}; padding:16px 22px; border-radius:10px; border:1px solid {color_fg}; margin-bottom:32px;">
        <div style="font-family:'IBM Plex Mono',monospace; font-size:12px; letter-spacing:0.08em; text-transform:uppercase; color:{color_fg}; margin-bottom:6px;">Régime de Marché : {label}</div>
        <div style="font-family:'IBM Plex Sans',sans-serif; font-size:14px; color:var(--ivory); line-height:1.5;">{comment}</div>
    </div>
    """

def ledger_table(headers, rows, num_cols=None):
    num_cols = num_cols or []
    th = "".join(f'<th class="num">{h}</th>' if i in num_cols else f"<th>{h}</th>" for i, h in enumerate(headers))
    trs = []
    for row in rows:
        tds = "".join(f'<td class="num">{c}</td>' if i in num_cols else f"<td>{c}</td>" for i, c in enumerate(row))
        trs.append(f"<tr>{tds}</tr>")
    return f"""<table class="ledger"><thead><tr>{th}</tr></thead><tbody>{''.join(trs)}</tbody></table>"""

def colorize(value, text):
    cls = "pos" if value >= 0 else "neg"
    return f'<span class="{cls}">{text}</span>'

def risk_control_panel(container_id, annual_return, annual_vol, default_taux, max_dd_pct):
    payload = {"annual_return": annual_return, "annual_vol": annual_vol}
    sharpe0 = (annual_return - default_taux) / annual_vol if annual_vol > 0 else 0.0
    status0 = "pill-good" if sharpe0 > 1 else ("pill-mid" if sharpe0 > 0.5 else "pill-bad")
    label0 = "Bon" if sharpe0 > 1 else ("Moyen" if sharpe0 > 0.5 else "Faible")
    return f"""
    <div class="control-panel">
      <div class="control">
        <label>Taux sans risque</label>
        <input type="range" id="{container_id}-slider" min="0" max="10" step="0.1" value="{default_taux*100:.1f}">
        <span class="control-value" id="{container_id}-slider-value">{pct(default_taux*100, signed=False)}</span>
      </div>
    </div>
    <div class="metric-row">
      <div class="metric-card"><div class="k">Ratio de Sharpe</div><div class="v" id="{container_id}-sharpe-value">{sharpe0:.2f}</div></div>
      <div class="metric-card"><div class="k">Volatilité annualisée</div><div class="v">{pct(annual_vol*100, signed=False)}</div></div>
      <div class="metric-card"><div class="k">Perte maximale (drawdown)</div><div class="v">{pct(max_dd_pct, signed=False)}</div></div>
    </div>
    <span class="pill-status {status0}" id="{container_id}-pill">{label0}</span>
    <script>
      window.__riskData = window.__riskData || {{}};
      window.__riskData["{container_id}"] = {json.dumps(payload)};
      initRiskControl("{container_id}");
    </script>
    """

def mc_control_panel(container_id, mc_payload, default_years, default_n_sims, min_years=1, max_years=15, min_sims=200, max_sims=3000, step_sims=100):
    def stat_placeholder(key, label):
        return f"""<div class="metric-card" id="{container_id}-stat-{key}"><div class="k">{label}</div><div class="v">—</div><div class="metric-delta"></div></div>"""

    stats_html = f"""<div class="metric-row">
      {stat_placeholder('p1', 'VaR 1% (crash)')}{stat_placeholder('p10', 'Pessimiste (10%)')}{stat_placeholder('p25', '1er quartile (25%)')}{stat_placeholder('median', 'Médiane (50%)')}
    </div><div class="metric-row">
      {stat_placeholder('p75', '3e quartile (75%)')}{stat_placeholder('p90', 'Optimiste (90%)')}{stat_placeholder('p99', 'Top 1% (99%)')}
    </div>"""

    return f"""
    <div class="control-panel">
      <div class="control">
        <label>Horizon</label>
        <input type="range" id="{container_id}-years" min="{min_years}" max="{max_years}" step="1" value="{default_years}">
        <span class="control-value" id="{container_id}-years-value">{default_years} ans</span>
      </div>
      <div class="control">
        <label>Nombre de tirages</label>
        <input type="range" id="{container_id}-sims" min="{min_sims}" max="{max_sims}" step="{step_sims}" value="{default_n_sims}">
        <span class="control-value" id="{container_id}-sims-value">{default_n_sims:,} tirages</span>
      </div>
      <div class="control">
        <label>Versement mensuel</label>
        <input type="range" id="{container_id}-apport" min="0" max="2000" step="50" value="0">
        <span class="control-value" id="{container_id}-apport-value">0 €</span>
      </div>
      <div class="control">
        <button class="btn-recalc" id="{container_id}-recalc">Recalculer</button>
        <span class="calc-status" id="{container_id}-status"></span>
      </div>
    </div>
    <div class="card"><div id="{container_id}-chart" style="height:380px;"></div></div>
    <div style="margin-top:16px;">{stats_html}</div>
    <script>
      window.__mcData = window.__mcData || {{}};
      window.__mcData["{container_id}"] = {json.dumps(mc_payload)};
      initMonteCarlo("{container_id}");
    </script>
    """

def whatif_control_panel(container_id, payload, candidate_labels, default_taux, current_value):
    if not candidate_labels:
        return "<div class='note'>Aucun candidat disponible (watchlist vide ou historique de prix insuffisant) pour simuler l'ajout d'une position.</div>"
    options_html = "".join(f'<option value="{t}">{t}</option>' for t in candidate_labels)
    max_montant = max(5000, int(round(current_value * 0.5, -2))) if current_value > 0 else 5000
    default_montant = min(5000, max_montant)
    return f"""
    <div class="control-panel">
      <div class="control">
        <label>Nouvelle position (watchlist)</label>
        <select id="{container_id}-ticker">
          <option value="">— Sélectionner un titre —</option>
          {options_html}
        </select>
      </div>
      <div class="control">
        <label>Montant investi</label>
        <input type="range" id="{container_id}-montant" min="0" max="{max_montant}" step="100" value="{default_montant}">
        <span class="control-value" id="{container_id}-montant-value">{eur(default_montant)}</span>
      </div>
    </div>
    <div id="{container_id}-results">
      <div class="note">Sélectionnez un titre ci-dessus pour évaluer l'impact sur la volatilité et les ratios de risque du portefeuille.</div>
    </div>
    <div class="control-panel" style="margin-top:16px;">
      <div class="control">
        <label>Horizon simulation</label>
        <input type="range" id="{container_id}-mc-years" min="1" max="15" step="1" value="5">
        <span class="control-value" id="{container_id}-mc-years-value">5 ans</span>
      </div>
      <div class="control">
        <label>Nombre de tirages</label>
        <input type="range" id="{container_id}-mc-sims" min="200" max="3000" step="100" value="1000">
        <span class="control-value" id="{container_id}-mc-sims-value">1 000 tirages</span>
      </div>
      <div class="control">
        <button class="btn-recalc" id="{container_id}-mc-recalc">Simuler (Monte Carlo)</button>
        <span class="calc-status" id="{container_id}-mc-status"></span>
      </div>
    </div>
    <div class="note" id="{container_id}-mc-note">Projette la valeur de la <b>nouvelle position seule</b> (montant investi ci-dessus, à partir de son propre historique de rendements) sur l'horizon choisi, par bootstrap par blocs. Sélectionnez d'abord un titre ci-dessus.</div>
    <div class="card" id="{container_id}-mc-card" style="margin-top:12px; display:none;">
      <div id="{container_id}-mc-chart" style="height:340px;"></div>
      <div style="margin-top:16px;" id="{container_id}-mc-stats"></div>
    </div>
    <script>
      window.__whatifData = window.__whatifData || {{}};
      window.__whatifData["{container_id}"] = {json.dumps(payload)};
      initWhatIf("{container_id}");
    </script>
    """

def hedge_control_panel(container_id, payload):
    instruments = payload.get("instruments", {})
    if not instruments:
        return "<div class='note'>Aucun indice de couverture disponible (cours indisponibles sur la fenêtre de calcul).</div>"
    # Instrument le plus corrélé au portefeuille proposé en premier
    ordre = sorted(instruments.items(), key=lambda kv: kv[1]["corr"], reverse=True)
    options_html = "".join(
        f'<option value="{t}">{i["label"]} — corrélation {i["corr"]:.2f}</option>' for t, i in ordre)
    leviers = payload.get("leviers", [20])
    levier_html = "".join(
        f'<option value="{l}"{" selected" if l == max(leviers) else ""}>{l}:1</option>' for l in leviers)
    taux_ref, spread = payload["taux_reference"], payload["spread"]
    return f"""
    <div class="control-panel">
      <div class="control">
        <label>Instrument de couverture</label>
        <select id="{container_id}-instr">{options_html}</select>
      </div>
      <div class="control">
        <label>Part du bêta couverte</label>
        <input type="range" id="{container_id}-ratio" min="0" max="100" step="5" value="50">
        <span class="control-value" id="{container_id}-ratio-value">50 %</span>
      </div>
      <div class="control">
        <label>Levier du CFD</label>
        <select id="{container_id}-levier">{levier_html}</select>
      </div>
    </div>
    <div id="{container_id}-results"></div>
    <div class="card" style="margin-top:16px;">
      <div class="card-title">Valeur de la poche actions sur 12 mois — sans / avec couverture (coûts inclus)</div>
      <div id="{container_id}-chart" style="height:320px;"></div>
    </div>
    <div style="margin-top:16px;" class="table-container" id="{container_id}-stress"></div>
    <div class="note">Simulation uniquement, aucune position n'est prise. Le notionnel vendu vaut
    <b>part couverte × bêta du portefeuille sur l'instrument × valeur des actions</b>. Coût annuel d'un
    CFD vendeur retenu : spread de financement ({pct(spread * 100, signed=False)}) − taux de référence
    reçu ({pct(taux_ref * 100, signed=False)}) + dividendes de l'indice reversés par le vendeur (indices
    de prix uniquement) — hypothèses à ajuster dans la config selon votre broker. La corrélation indique
    la part du risque réellement couvrable : un indice peu corrélé laisse un risque de base important.
    Le seuil de liquidation suppose que toutes les espèces du compte servent de dépôt et une clôture
    automatique à 50 % de la marge (règle ESMA). Les gains de la couverture sont imposés au PFU,
    alors que les pertes d'un PEA ne les compensent pas.</div>
    <script>
      window.__hedgeData = window.__hedgeData || {{}};
      window.__hedgeData["{container_id}"] = {json.dumps(payload)};
      initHedge("{container_id}");
    </script>
    """

def json_pour_script(payload):
    """JSON injectable dans une balise <script> même s'il contient du texte externe (titres
    d'articles) : '<' est échappé pour qu'un « </script> » ne puisse pas fermer la balise."""
    return json.dumps(payload).replace("<", "\\u003c")

def news_panel(container_id, payload):
    if not payload.get("articles"):
        return "<div class='note'>Aucune actualité récente trouvée pour vos lignes et thèmes suivis.</div>"
    esc = lambda s: html_lib.escape(str(s), quote=True)
    cibles_html = "".join(f'<option value="{esc(c["code"])}">{esc(c["label"])}</option>' for c in payload["cibles"])
    types_html = "".join(f'<option value="{esc(t)}">{esc(t)}</option>' for t in payload["types"])
    return f"""
    <div class="control-panel">
      <div class="control">
        <label>Ligne ou thème</label>
        <select id="{container_id}-cible"><option value="">Tout le portefeuille</option>{cibles_html}</select>
      </div>
      <div class="control">
        <label>Type d'information</label>
        <select id="{container_id}-type"><option value="">Tous les types</option>{types_html}</select>
      </div>
      <div class="control">
        <label>Tri</label>
        <select id="{container_id}-tri"><option value="score">Pertinence</option><option value="date">Date</option></select>
      </div>
      <div class="control" style="min-width:auto;">
        <label>Filtre</label>
        <label class="news-check"><input type="checkbox" id="{container_id}-signif"> Réactions significatives</label>
      </div>
    </div>
    <div class="calc-status" id="{container_id}-count" style="margin-left:0;"></div>
    <div id="{container_id}-liste"></div>
    <script>
      window.__newsData = window.__newsData || {{}};
      window.__newsData["{container_id}"] = {json_pour_script(payload)};
      initNews("{container_id}");
    </script>
    """

def section(eyebrow, body_html):
    return f"""<div class="section"><div class="section-head"><span class="eyebrow">{eyebrow}</span><span class="rule"></span></div>{body_html}</div>"""

# ---------------------------------------------------------------
# 4. SHELL HTML GLOBAL (masthead, CSS, nav à onglets, footer)
# ---------------------------------------------------------------
CSS = """
:root{
  --ink:#0E1A2B; --panel:#15243B; --panel-2:#1C2E48; --line:#2B3D59;
  --gold:#C9A15E; --gold-soft:rgba(201,161,94,.14); --ivory:#EFE9DB;
  --slate:#8B96A8; --sage:#7BAE95; --brick:#B15B45;
}
*{box-sizing:border-box;}
html,body{margin:0;padding:0;}
body{
  background:radial-gradient(1200px 500px at 15% -10%, rgba(201,161,94,.06), transparent), var(--ink);
  color:var(--ivory); font-family:'IBM Plex Sans',sans-serif; -webkit-font-smoothing:antialiased;
  padding:48px 24px 96px;
}
.sheet{max-width:1400px; margin:0 auto;}
.eyebrow{font-family:'IBM Plex Mono',monospace; font-size:14px; font-weight:600; letter-spacing:.12em; text-transform:uppercase; color:var(--gold);}
.masthead{display:flex; justify-content:space-between; align-items:flex-end; border-bottom:1px solid var(--line); padding-bottom:20px; margin-bottom:28px;}
.masthead h1{font-family:'Fraunces',serif; font-weight:500; font-size:22px; margin:6px 0 0; letter-spacing:.01em;}
.masthead .right{text-align:right; color:var(--slate); font-family:'IBM Plex Mono',monospace; font-size:12px;}
.status-dot{display:inline-block;width:6px;height:6px;border-radius:50%;background:var(--sage);margin-right:6px;box-shadow:0 0 0 3px rgba(123,174,149,.15);}
.tabnav{display:flex; gap:2px; border-bottom:1px solid var(--line); margin-bottom:40px; overflow-x:auto;
  scrollbar-width:thin; scrollbar-color:var(--line) transparent;}
.tabnav::-webkit-scrollbar{height:4px;} .tabnav::-webkit-scrollbar-thumb{background:var(--line); border-radius:2px;}
.tabnav::-webkit-scrollbar-track{background:transparent;}
.tabnav button{
  background:none; border:none; color:var(--slate); font-family:'IBM Plex Mono',monospace;
  font-size:11px; letter-spacing:.06em; text-transform:uppercase; padding:12px 13px; cursor:pointer;
  border-bottom:2px solid transparent; white-space:nowrap; transition:color .15s, border-color .15s;
}
.tabnav button:hover{color:var(--ivory);}
.tabnav button.active{color:var(--gold); border-bottom-color:var(--gold);}
.tabpanel{display:none;} .tabpanel.active{display:block;}
.hero{position:relative; padding:8px 0 36px;}
.hero .spark{position:absolute; left:0; right:0; bottom:0; height:120px; opacity:.35; pointer-events:none;}
.hero-label{font-family:'Fraunces',serif; font-size:15px; color:var(--slate); margin-bottom:10px;}
.hero-number{font-family:'IBM Plex Mono',monospace; font-weight:600; font-size:clamp(36px,7vw,64px);
  letter-spacing:-.01em; display:flex; align-items:baseline; gap:16px; position:relative; z-index:2; flex-wrap:wrap;}
.hero-number .cur{font-size:.5em; color:var(--gold); font-weight:500;}
.delta-badge{font-family:'IBM Plex Mono',monospace; font-size:14px; padding:5px 10px; border-radius:20px;
  background:var(--gold-soft); color:var(--sage); border:1px solid rgba(123,174,149,.25);}
.delta-badge.neg{color:var(--brick); background:rgba(177,91,69,.12); border-color:rgba(177,91,69,.25);}
.alert-banner { background: rgba(177,91,69,.15); border: 1px solid var(--brick); border-radius: 8px; padding: 12px 20px; margin-bottom: 24px; }
.alert-banner ul { margin: 0; padding-left: 20px; color: var(--ivory); font-family: 'IBM Plex Sans', sans-serif; font-size: 13px; line-height: 1.6; }
.alert-banner li { margin-bottom: 4px; }
.alert-banner li:last-child { margin-bottom: 0; }
.stat-row{display:grid; grid-template-columns:repeat(auto-fit,minmax(180px,1fr)); gap:1px;
  background:var(--line); border:1px solid var(--line); border-radius:10px; overflow:hidden; margin-top:28px;}
.stat-chip{background:var(--panel); padding:18px 22px;}
.stat-chip .k{font-family:'IBM Plex Mono',monospace; font-size:10px; letter-spacing:.1em; text-transform:uppercase; color:var(--slate); margin-bottom:8px;}
.stat-chip .v{font-family:'IBM Plex Mono',monospace; font-size:20px; font-weight:500;}
.stat-chip .v small{color:var(--slate); font-size:12px; font-weight:400; margin-left:6px;}
.metric-row{display:grid; grid-template-columns:repeat(auto-fit,minmax(160px,1fr)); gap:16px; margin-bottom:8px;}
.metric-card{background:var(--panel); border:1px solid var(--line); border-radius:10px; padding:16px 18px;}
.metric-card .k{font-family:'IBM Plex Mono',monospace; font-size:12px; font-weight:600; letter-spacing:.05em; text-transform:uppercase; color:var(--slate); margin-bottom:8px;}
.metric-card .v{font-family:'IBM Plex Mono',monospace; font-size:19px; font-weight:500;}
.metric-delta{font-family:'IBM Plex Mono',monospace; font-size:11px; margin-top:4px;}
.section{margin-top:48px;}
.section-head{display:flex; align-items:center; gap:14px; margin-bottom:20px;}
.section-head .rule{height:1px; flex:1; background:linear-gradient(90deg,var(--line),transparent);}
.card{background:var(--panel); border:1px solid var(--line); border-radius:12px; padding:22px;}
.card-title{font-family:'Fraunces',serif; font-size:14px; color:var(--slate); margin-bottom:10px;}
.grid-2{display:grid; grid-template-columns:1fr 1fr; gap:20px;}
.grid-3{display:grid; grid-template-columns:1fr 1fr 1fr; gap:20px;}
.grid-12{display:grid; grid-template-columns:1fr 2fr; gap:20px;}
.grid-2 > *, .grid-3 > *, .grid-12 > *{min-width:0;}   /* un graphique Plotly n'élargit plus sa colonne */
@media (max-width:820px){ .grid-2,.grid-3,.grid-12{grid-template-columns:1fr;} }
.table-container { width: 100%; overflow-x: auto; }
table.ledger{width:100%; border-collapse:collapse; font-family:'IBM Plex Mono',monospace; font-size:12.5px;}
table.ledger th, table.ledger td { white-space: nowrap; }
table.ledger thead th{text-align:left; font-family:'IBM Plex Sans',sans-serif; font-size:10px; letter-spacing:.08em;
  text-transform:uppercase; color:var(--slate); font-weight:500; padding:0 10px 10px; border-bottom:1px solid var(--line);}
table.ledger thead th.num{text-align:right;}
table.ledger thead th{cursor:pointer; transition:color 0.15s;}
table.ledger thead th:hover{color:var(--gold);}
table.ledger thead th.asc::after{content:' ▲'; font-size:8px; color:var(--gold);}
table.ledger thead th.desc::after{content:' ▼'; font-size:8px; color:var(--gold);}
table.ledger tbody td{padding:10px; border-bottom:1px solid rgba(43,61,89,.5);}
table.ledger tbody tr:last-child td{border-bottom:none;}
table.ledger tbody tr:hover td{background:var(--panel-2);}
table.ledger td.num{text-align:right;}
.pos{color:var(--sage);} .neg{color:var(--brick);}
.tag{display:inline-block; font-family:'IBM Plex Sans',sans-serif; font-size:10px; padding:2px 8px; border-radius:10px; background:var(--gold-soft); color:var(--gold); margin-left:8px;}
.pill-status{display:inline-block; padding:3px 10px; border-radius:12px; font-family:'IBM Plex Mono',monospace; font-size:11px;}
.pill-good{background:rgba(123,174,149,.15); color:var(--sage);}
.pill-mid{background:var(--gold-soft); color:var(--gold);}
.pill-bad{background:rgba(177,91,69,.15); color:var(--brick);}
.control-panel{display:flex; flex-wrap:wrap; gap:28px; align-items:flex-end; margin-bottom:22px;}
.control{display:flex; flex-direction:column; gap:8px; min-width:200px;}
.control label{font-family:'IBM Plex Mono',monospace; font-size:10px; letter-spacing:.08em; text-transform:uppercase; color:var(--slate);}
.control .control-value{font-family:'IBM Plex Mono',monospace; font-size:13px; color:var(--gold);}
.control input[type=range]{-webkit-appearance:none; appearance:none; width:100%; height:4px; border-radius:2px; background:var(--line); outline:none;}
.control input[type=range]::-webkit-slider-thumb{-webkit-appearance:none; width:14px; height:14px; border-radius:50%; background:var(--gold); cursor:pointer; box-shadow:0 0 0 3px rgba(201,161,94,.18);}
.control input[type=range]::-moz-range-thumb{width:14px; height:14px; border-radius:50%; background:var(--gold); border:none; cursor:pointer;}
.control select{background:var(--panel-2); color:var(--ivory); border:1px solid var(--line); border-radius:6px;
  padding:9px 10px; font-family:'IBM Plex Mono',monospace; font-size:12px; outline:none; min-width:200px;}
.control select:focus{border-color:var(--gold);}
.metric-card .v small{font-family:'IBM Plex Mono',monospace;}
.btn-recalc{font-family:'IBM Plex Mono',monospace; font-size:11px; letter-spacing:.06em; text-transform:uppercase;
  background:var(--gold-soft); color:var(--gold); border:1px solid rgba(201,161,94,.35); border-radius:8px;
  padding:9px 16px; cursor:pointer; transition:background .15s;}
.btn-recalc:hover{background:rgba(201,161,94,.24);}
.btn-recalc:disabled{opacity:.5; cursor:wait;}
.calc-status{font-family:'IBM Plex Mono',monospace; font-size:11px; color:var(--slate); margin-left:12px;}
.note{color:var(--slate); font-size:12.5px; line-height:1.7; margin-top:14px;}
.news-item{padding:14px 2px; border-bottom:1px solid rgba(43,61,89,.5);}
.news-item:last-child{border-bottom:none;}
.news-impact{border-left:2px solid var(--brick); padding-left:12px;}
.news-meta{display:flex; flex-wrap:wrap; gap:6px; align-items:center; font-family:'IBM Plex Mono',monospace; font-size:11px; color:var(--slate);}
.news-meta .tag{margin-left:0;}
.tag-theme{background:rgba(123,174,149,.14); color:var(--sage);}
.tag-type{background:rgba(139,150,168,.14); color:var(--slate);}
.tag-impact{background:rgba(177,91,69,.15); color:var(--brick); margin-left:8px;}
.news-score{margin-left:auto; color:var(--gold);}
.news-title{display:block; margin:8px 0 4px; color:var(--ivory); font-size:14.5px; font-weight:500; line-height:1.4; text-decoration:none;}
.news-title:hover{color:var(--gold);}
.news-resume{margin:0; color:var(--slate); font-size:12.5px; line-height:1.55;}
.news-reaction{margin-top:6px; font-family:'IBM Plex Mono',monospace; font-size:11.5px; color:var(--slate);}
.control label.news-check{display:flex; align-items:center; gap:8px; font-family:'IBM Plex Sans',sans-serif; font-size:12.5px;
  color:var(--ivory); text-transform:none; letter-spacing:0; padding:9px 0; cursor:pointer;}
footer{margin-top:64px; padding-top:20px; border-top:1px solid var(--line); color:var(--slate);
  font-family:'IBM Plex Mono',monospace; font-size:11px; line-height:1.7;}
footer strong{color:var(--gold);}
"""

TAB_JS = """
function showTab(id, btn){
  document.querySelectorAll('.tabpanel').forEach(p=>p.classList.remove('active'));
  document.querySelectorAll('.tabnav button').forEach(b=>b.classList.remove('active'));
  const panel = document.getElementById(id);
  panel.classList.add('active');
  btn.classList.add('active');
  // Les graphiques tracés dans un onglet masqué ont une largeur fausse : on les redimensionne
  if (window.Plotly && Plotly.Plots){
    panel.querySelectorAll('.js-plotly-plot').forEach(function(el){
      try { Plotly.Plots.resize(el); } catch (e) { console.error(e); }
    });
  }
  // Monte Carlo différé : premier calcul à la première ouverture de l'onglet qui le contient
  if (window.__mcLazy){
    Object.keys(window.__mcLazy).forEach(function(k){
      const job = window.__mcLazy[k];
      if (!job.done && job.panelId === id){ job.done = true; job.run(); }
    });
  }
}
document.addEventListener('DOMContentLoaded', () => {
  document.querySelectorAll('table.ledger').forEach(table => {
    const headers = table.querySelectorAll('th');
    headers.forEach((th, i) => {
      th.addEventListener('click', () => {
        const tbody = table.querySelector('tbody');
        const rows = Array.from(tbody.querySelectorAll('tr'));
        const isAsc = th.classList.contains('asc');
        headers.forEach(h => h.classList.remove('asc', 'desc'));
        th.classList.toggle('desc', isAsc);
        th.classList.toggle('asc', !isAsc);
        rows.sort((a, b) => {
          let v1 = a.children[i].textContent.trim();
          let v2 = b.children[i].textContent.trim();
          const parseVal = (v) => {
            let clean = v.replace(/[€%\\s]/g, '').replace(',', '.');
            let num = Number(clean);
            return isNaN(num) ? v.toLowerCase() : num;
          };
          let p1 = parseVal(v1);
          let p2 = parseVal(v2);
          if (p1 < p2) return isAsc ? 1 : -1;
          if (p1 > p2) return isAsc ? -1 : 1;
          return 0;
        });
        tbody.append(...rows);
      });
    });
  });
});
"""

MC_ENGINE_JS = """
function mulberry32(seed){
  let a = seed >>> 0;
  return function(){
    a |= 0; a = (a + 0x6D2B79F5) | 0;
    let t = Math.imul(a ^ (a >>> 15), 1 | a);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}
function makeNormalSampler(rng){
  let spare = null;
  return function(){
    if (spare !== null){ const v = spare; spare = null; return v; }
    let u1 = 0, u2;
    while (u1 <= 1e-12) u1 = rng();
    u2 = rng();
    const r = Math.sqrt(-2*Math.log(u1));
    const theta = 2*Math.PI*u2;
    spare = r*Math.sin(theta);
    return r*Math.cos(theta);
  };
}
function randGamma(shape, rng, randNormal){
  if (shape < 1){
    const u = rng();
    return randGamma(1+shape, rng, randNormal) * Math.pow(u, 1/shape);
  }
  const d = shape - 1/3;
  const c = 1/Math.sqrt(9*d);
  while (true){
    let x, v;
    do { x = randNormal(); v = 1 + c*x; } while (v <= 0);
    v = v*v*v;
    const u = rng();
    if (u < 1 - 0.0331*x*x*x*x) return d*v;
    if (Math.log(u) < 0.5*x*x + d*(1-v+Math.log(v))) return d*v;
  }
}
function randChi2(nu, rng, randNormal){ return 2*randGamma(nu/2, rng, randNormal); }

function percentile(sortedArr, q){
  const n = sortedArr.length;
  const pos = (q/100)*(n-1);
  const lo = Math.floor(pos), hi = Math.ceil(pos);
  if (lo === hi) return sortedArr[lo];
  const frac = pos - lo;
  return sortedArr[lo]*(1-frac) + sortedArr[hi]*frac;
}

function runGarchSim(data, days, nSims, apportMensuel){
  const { tickers, weights, initial_value, seed, nu_joint, L, omega, alpha, gamma, beta, mu, h0, eps0 } = data;
  const nAssets = tickers.length;
  const rng = mulberry32(seed);
  const randNormal = makeNormalSampler(rng);
  const tScale = Math.sqrt((nu_joint-2)/nu_joint);

  const h = new Float64Array(nSims*nAssets);
  const eps = new Float64Array(nSims*nAssets);
  for (let s=0; s<nSims; s++){
    for (let a=0; a<nAssets; a++){ h[s*nAssets+a] = h0[a]; eps[s*nAssets+a] = eps0[a]; }
  }
  const val = new Float64Array(nSims).fill(initial_value);
  const out = { median:[], p10:[], p25:[], p75:[], p90:[], p1:[], p99:[] };
  const zNorm = new Float64Array(nAssets);
  const zCorr = new Float64Array(nAssets);
  const dayVals = new Float64Array(nSims);
  const dailyApport = (apportMensuel || 0) / 21.0;

  for (let t=0; t<days; t++){
    for (let s=0; s<nSims; s++){
      const base = s*nAssets;
      for (let a=0; a<nAssets; a++){
        const e = eps[base+a];
        const indicator = e < 0 ? 1 : 0;
        let hv = omega[a] + alpha[a]*e*e + gamma[a]*e*e*indicator + beta[a]*h[base+a];
        if (hv < 1e-8) hv = 1e-8;
        h[base+a] = hv;
      }
      for (let a=0; a<nAssets; a++) zNorm[a] = randNormal();
      for (let a=0; a<nAssets; a++){
        let sum = 0;
        for (let b=0; b<=a; b++) sum += L[a][b]*zNorm[b];
        zCorr[a] = sum;
      }
      const chi2 = randChi2(nu_joint, rng, randNormal);
      const scale = Math.sqrt(nu_joint/chi2)*tScale;
      let pfReturn = 0;
      for (let a=0; a<nAssets; a++){
        const z = zCorr[a]*scale;
        const e = Math.sqrt(h[base+a])*z;
        eps[base+a] = e;
        pfReturn += (mu[a] + e/100) * weights[a];
      }
      if (pfReturn < -0.99) pfReturn = -0.99;
      val[s] = val[s]*(1+pfReturn) + dailyApport;
      dayVals[s] = val[s];
    }
    const sorted = Array.from(dayVals).sort((a,b)=>a-b);
    out.p1.push(percentile(sorted,1)); out.p10.push(percentile(sorted,10));
    out.p25.push(percentile(sorted,25)); out.median.push(percentile(sorted,50));
    out.p75.push(percentile(sorted,75)); out.p90.push(percentile(sorted,90));
    out.p99.push(percentile(sorted,99));
  }
  return out;
}

function runBootstrapSim(data, days, nSims, apportMensuel){
  const { weights, initial_value, seed, block_size, returns_matrix } = data;
  const nHist = returns_matrix.length;
  const nAssets = weights.length;
  const rng = mulberry32(seed);
  const val = new Float64Array(nSims).fill(initial_value);
  const out = { median:[], p10:[], p25:[], p75:[], p90:[], p1:[], p99:[] };
  const dayVals = new Float64Array(nSims);
  const nBlocks = Math.ceil(days/block_size);
  const blockStartsPerSim = [];
  const dailyApport = (apportMensuel || 0) / 21.0;

  for (let s=0; s<nSims; s++){
    const starts = new Array(nBlocks);
    for (let b=0; b<nBlocks; b++) starts[b] = Math.floor(rng()*(nHist-block_size));
    blockStartsPerSim.push(starts);
  }
  for (let t=0; t<days; t++){
    for (let s=0; s<nSims; s++){
      const blockIdx = Math.floor(t/block_size);
      const offset = t % block_size;
      const row = returns_matrix[blockStartsPerSim[s][blockIdx] + offset];
      let pfReturn = 0;
      for (let a=0; a<nAssets; a++) pfReturn += row[a]*weights[a];
      if (pfReturn < -0.99) pfReturn = -0.99;
      val[s] = val[s]*(1+pfReturn) + dailyApport;
      dayVals[s] = val[s];
    }
    const sorted = Array.from(dayVals).sort((a,b)=>a-b);
    out.p1.push(percentile(sorted,1)); out.p10.push(percentile(sorted,10));
    out.p25.push(percentile(sorted,25)); out.median.push(percentile(sorted,50));
    out.p75.push(percentile(sorted,75)); out.p90.push(percentile(sorted,90));
    out.p99.push(percentile(sorted,99));
  }
  return out;
}
"""

INTERACTIVE_JS = """
function fmtEur(x){
  return x.toLocaleString('fr-FR', {minimumFractionDigits:2, maximumFractionDigits:2}) + ' €';
}
function fmtPct(x){
  const sign = x>=0 ? '+' : '';
  return sign + x.toFixed(2) + ' %';
}

function initRiskControl(id){
  const data = window.__riskData[id];
  const slider = document.getElementById(id+'-slider');
  const valueLabel = document.getElementById(id+'-slider-value');
  const sharpeEl = document.getElementById(id+'-sharpe-value');
  const pillEl = document.getElementById(id+'-pill');
  function update(){
    const taux = parseFloat(slider.value)/100;
    valueLabel.textContent = taux.toLocaleString('fr-FR', {style:'percent', minimumFractionDigits:1});
    const sharpe = data.annual_vol > 0 ? (data.annual_return - taux)/data.annual_vol : 0;
    sharpeEl.textContent = sharpe.toFixed(2);
    let cls = 'pill-bad', label = 'Faible';
    if (sharpe > 1){ cls='pill-good'; label='Bon'; }
    else if (sharpe > 0.5){ cls='pill-mid'; label='Moyen'; }
    pillEl.className = 'pill-status ' + cls;
    pillEl.textContent = label;
  }
  slider.addEventListener('input', update);
  update();
}

function whatifStats(returns, taux){
  const n = returns.length;
  const mean = returns.reduce((a,b)=>a+b,0)/n;
  const variance = n>1 ? returns.reduce((a,b)=>a+(b-mean)*(b-mean),0)/(n-1) : 0;
  const annualReturn = mean*252;
  const annualVol = Math.sqrt(Math.max(variance,0)*252);
  const sharpe = annualVol>0 ? (annualReturn-taux)/annualVol : 0;

  // Downside deviation standard (cible 0) sur TOUS les jours : sqrt(mean(min(r,0)^2)) * sqrt(252)
  const downsideVol = n>0
    ? Math.sqrt(returns.reduce((a,r)=>a+Math.min(r,0)*Math.min(r,0),0)/n * 252)
    : 0;
  const sortino = downsideVol>0 ? (annualReturn-taux)/downsideVol : 0;

  let cum=1, peak=1, maxDD=0;
  for (let i=0;i<n;i++){
    cum *= (1+returns[i]);
    if (cum>peak) peak=cum;
    const dd = (cum-peak)/peak;
    if (dd<maxDD) maxDD = dd;
  }
  return {annualReturn, annualVol, sharpe, sortino, maxDD: maxDD*100};
}
function whatifCov(a,b){
  const n = a.length;
  const ma = a.reduce((x,y)=>x+y,0)/n;
  const mb = b.reduce((x,y)=>x+y,0)/n;
  let s = 0;
  for (let i=0;i<n;i++) s += (a[i]-ma)*(b[i]-mb);
  return n>1 ? s/(n-1) : 0;
}
function whatifBeta(pf, bench){
  const covPB = whatifCov(pf, bench);
  const varB = whatifCov(bench, bench);
  return varB>0 ? covPB/varB : 1.0;
}
function whatifCorr(a,b){
  const covAB = whatifCov(a,b);
  const sA = Math.sqrt(whatifCov(a,a));
  const sB = Math.sqrt(whatifCov(b,b));
  return (sA>0 && sB>0) ? covAB/(sA*sB) : 0;
}
function fmtEurWhatif(x){
  return Math.round(x).toLocaleString('fr-FR') + ' €';
}

function initWhatIf(id){
  const data = window.__whatifData[id];
  const tickerSel = document.getElementById(id+'-ticker');
  const montantSlider = document.getElementById(id+'-montant');
  const montantVal = document.getElementById(id+'-montant-value');
  const resultsEl = document.getElementById(id+'-results');
  if (!tickerSel || !montantSlider) return;

  const mcYears = document.getElementById(id+'-mc-years');
  const mcSims = document.getElementById(id+'-mc-sims');
  const mcYearsVal = document.getElementById(id+'-mc-years-value');
  const mcSimsVal = document.getElementById(id+'-mc-sims-value');
  const mcBtn = document.getElementById(id+'-mc-recalc');
  const mcStatus = document.getElementById(id+'-mc-status');
  const mcCard = document.getElementById(id+'-mc-card');
  const mcStats = document.getElementById(id+'-mc-stats');
  let lastSim = { returns: null, total: 0 };

  function syncMcLabels(){
    if (mcYearsVal) mcYearsVal.textContent = mcYears.value + ' ans';
    if (mcSimsVal) mcSimsVal.textContent = parseInt(mcSims.value).toLocaleString('fr-FR') + ' tirages';
  }
  if (mcYears) mcYears.addEventListener('input', syncMcLabels);
  if (mcSims) mcSims.addEventListener('input', syncMcLabels);
  syncMcLabels();

  function runWhatifMonteCarlo(){
    if (!mcBtn) return;
    if (!lastSim.returns){
      mcStatus.textContent = 'Sélectionnez d\\u2019abord un titre ci-dessus.';
      return;
    }
    if (!lastSim.total || lastSim.total <= 0){
      mcStatus.textContent = 'Augmentez le montant investi pour lancer la simulation.';
      return;
    }
    mcBtn.disabled = true;
    mcStatus.textContent = 'Calcul en cours…';
    setTimeout(function(){
      try {
        const years = parseInt(mcYears.value);
        const nSims = parseInt(mcSims.value);
        const days = years*252;
        const payload = {
          weights: [1],
          initial_value: lastSim.total,
          seed: 42,
          block_size: 5,
          returns_matrix: lastSim.returns.map(function(r){ return [r]; })
        };
        const out = runBootstrapSim(payload, days, nSims, 0);

        const dates = [];
        const d = new Date();
        let added = 0;
        while (added < days){
          d.setDate(d.getDate()+1);
          const day = d.getDay();
          if (day !== 0 && day !== 6){ dates.push(d.toISOString().slice(0,10)); added++; }
        }

        const traces = [
          {x: dates.concat(dates.slice().reverse()), y: out.p90.concat(out.p10.slice().reverse()),
           fill:'toself', fillcolor:'rgba(201,161,94,0.15)', line:{color:'rgba(0,0,0,0)'},
           name:'Intervalle 80%', hoverinfo:'skip'},
          {x: dates.concat(dates.slice().reverse()), y: out.p75.concat(out.p25.slice().reverse()),
           fill:'toself', fillcolor:'rgba(201,161,94,0.28)', line:{color:'rgba(0,0,0,0)'},
           name:'Intervalle 50%', hoverinfo:'skip'},
          {x: dates, y: out.median, mode:'lines', name:'Scénario médian', line:{color:'#EFE9DB', width:2}},
        ];
        const layout = {
          paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)',
          font:{family:'IBM Plex Sans, sans-serif', color:'#8B96A8', size:11},
          xaxis:{gridcolor:'#2B3D59', zeroline:false, tickfont:{family:'IBM Plex Mono, monospace', size:10}},
          yaxis:{gridcolor:'#2B3D59', zeroline:false, tickfont:{family:'IBM Plex Mono, monospace', size:10}, tickprefix:'€'},
          legend:{font:{color:'#EFE9DB', size:10}, orientation:'h', y:1.14, x:0},
          margin:{l:50,r:20,t:30,b:40}, height:340,
        };
        mcCard.style.display = 'block';
        Plotly.react(id+'-mc-chart', traces, layout, {displayModeBar:false, responsive:true});

        const initial = lastSim.total;
        function statCard(label, val){
          const chg = initial>0 ? (val/initial-1)*100 : 0;
          return '<div class="metric-card"><div class="k">'+label+'</div><div class="v">'+fmtEur(val)+'</div>' +
            '<div class="metric-delta '+(chg>=0?'pos':'neg')+'">'+fmtPct(chg)+'</div></div>';
        }
        const last = (arr)=>arr[arr.length-1];
        mcStats.innerHTML =
          '<div class="metric-row">' +
            statCard('VaR 1% (crash)', last(out.p1)) + statCard('Pessimiste (10%)', last(out.p10)) +
            statCard('Médiane (50%)', last(out.median)) + statCard('Optimiste (90%)', last(out.p90)) +
          '</div>';

        mcStatus.textContent = '';
        mcBtn.disabled = false;
      } catch (e) {
        mcStatus.textContent = 'Erreur: ' + e.message;
        console.error(e);
        mcBtn.disabled = false;
      }
    }, 30);
  }
  if (mcBtn) mcBtn.addEventListener('click', runWhatifMonteCarlo);

  function deltaHtml(oldV, newV, higherIsBetter, decimals, suffix){
    const diff = newV - oldV;
    const good = higherIsBetter ? diff >= 0 : diff <= 0;
    const cls = Math.abs(diff) < 1e-9 ? '' : (good ? 'pos' : 'neg');
    const sign = diff >= 0 ? '+' : '';
    return '<span class="' + cls + '">' + sign + diff.toFixed(decimals) + suffix + '</span>';
  }

  function metricCard(label, oldV, newV, higherIsBetter, decimals, suffix){
    return '<div class="metric-card"><div class="k">' + label + '</div>' +
      '<div class="v">' + newV.toFixed(decimals) + suffix +
      ' <small style="color:var(--slate); font-size:11px;">(' + oldV.toFixed(decimals) + suffix + ' avant)</small></div>' +
      '<div class="metric-delta">' + deltaHtml(oldV, newV, higherIsBetter, decimals, suffix) + '</div></div>';
  }

  function render(){
    montantVal.textContent = fmtEurWhatif(parseFloat(montantSlider.value));
    const ticker = tickerSel.value;
    if (!ticker){
      resultsEl.innerHTML = '<div class="note">Sélectionnez un titre ci-dessus pour évaluer l\\u2019impact sur la volatilité et les ratios de risque du portefeuille.</div>';
      lastSim = { returns: null, total: 0 };
      if (mcCard) mcCard.style.display = 'none';
      if (mcStats) mcStats.innerHTML = '';
      if (mcStatus) mcStatus.textContent = '';
      return;
    }
    const montant = parseFloat(montantSlider.value);
    const V = data.current_value;
    const newTotal = V + montant;
    const wNew = newTotal > 0 ? montant/newTotal : 0;
    const scaleOld = newTotal > 0 ? V/newTotal : 0;

    const nDates = data.returns_matrix.length;
    const nAssets = data.weights.length;
    const candAll = data.candidates[ticker] || [];

    // Seules les dates où le candidat cote réellement sont retenues (null = avant son
    // introduction en bourse) : avant/après sont comparés sur exactement la même période.
    const baseReturns = [], newReturns = [], candReturns = [], benchReturns = [], usedDates = [];
    for (let t=0; t<nDates; t++){
      const c = candAll[t];
      if (c === null || c === undefined) continue;
      let rBase = 0;
      const row = data.returns_matrix[t];
      for (let a=0; a<nAssets; a++) rBase += row[a]*data.weights[a];
      baseReturns.push(rBase);
      newReturns.push(rBase*scaleOld + c*wNew);
      candReturns.push(c);
      benchReturns.push(data.bench_returns[t]);
      if (data.dates) usedDates.push(data.dates[t]);
    }

    // Le Monte Carlo ne doit projeter QUE la nouvelle position isolée (son propre historique
    // de rendements et son propre montant investi), pas le portefeuille combiné.
    lastSim = { returns: candReturns, total: montant };
    if (mcCard) mcCard.style.display = 'none';
    if (mcStats) mcStats.innerHTML = '';
    if (mcStatus) mcStatus.textContent = '';

    if (candReturns.length < 30){
      resultsEl.innerHTML = '<div class="note">Historique de cotation trop court pour ce titre sur la fenêtre de calcul.</div>';
      lastSim = { returns: null, total: 0 };
      return;
    }

    const taux = data.taux_sans_risque;
    const before = whatifStats(baseReturns, taux);
    const after = whatifStats(newReturns, taux);
    const beta = data.has_bench === false ? null : whatifBeta(newReturns, benchReturns);
    const corr = whatifCorr(baseReturns, candReturns);
    const periodNote = (candReturns.length < nDates && usedDates.length)
      ? '<div class="note">Comparaison calculée sur les ' + candReturns.length + ' séances où ' + ticker +
        ' cote réellement (du ' + usedDates[0] + ' au ' + usedDates[usedDates.length-1] + '), au lieu de ' + nDates + '.</div>'
      : '';

    resultsEl.innerHTML =
      '<div class="metric-row" style="margin-top:16px;">' +
        metricCard('Volatilité annualisée', before.annualVol*100, after.annualVol*100, false, 2, ' %') +
        metricCard('Ratio de Sharpe', before.sharpe, after.sharpe, true, 2, '') +
        metricCard('Ratio de Sortino', before.sortino, after.sortino, true, 2, '') +
        metricCard('Max Drawdown (1 an)', before.maxDD, after.maxDD, true, 2, ' %') +
      '</div>' +
      '<div class="metric-row">' +
        '<div class="metric-card"><div class="k">Beta vs Benchmark (après ajout)</div><div class="v">' + (beta === null ? 'n/d' : beta.toFixed(2)) + '</div></div>' +
        '<div class="metric-card"><div class="k">Corrélation position / portefeuille actuel</div><div class="v">' + corr.toFixed(2) + '</div></div>' +
        '<div class="metric-card"><div class="k">Poids de la nouvelle position</div><div class="v">' + (wNew*100).toFixed(1) + ' %</div></div>' +
      '</div>' + periodNote;
  }

  tickerSel.addEventListener('change', render);
  montantSlider.addEventListener('input', render);
  render();
}

function initNews(id){
  const data = window.__newsData[id];
  const selCible = document.getElementById(id+'-cible');
  const selType = document.getElementById(id+'-type');
  const selTri = document.getElementById(id+'-tri');
  const chkSignif = document.getElementById(id+'-signif');
  const liste = document.getElementById(id+'-liste');
  const countEl = document.getElementById(id+'-count');
  if (!data || !liste || !selCible) return;

  // Contenu externe (titres, résumés, sources) : toujours inséré via textContent, jamais en HTML
  function el(tag, cls, text){
    const e = document.createElement(tag);
    if (cls) e.className = cls;
    if (text !== undefined && text !== null) e.textContent = text;
    return e;
  }
  function fmtDate(iso){
    const d = new Date(iso);
    return d.toLocaleDateString('fr-FR', {day:'2-digit', month:'2-digit'}) + ' ' +
      d.toLocaleTimeString('fr-FR', {hour:'2-digit', minute:'2-digit'});
  }
  function pctTxt(x){ return (x >= 0 ? '+' : '') + (x*100).toFixed(2) + ' %'; }

  function render(){
    const cible = selCible.value, type = selType.value, signif = chkSignif.checked;
    let arts = data.articles.filter(function(a){
      return (!cible || a.cibles.some(function(c){ return c.code === cible; })) &&
             (!type || a.types.indexOf(type) >= 0) && (!signif || a.significative);
    });
    if (selTri.value === 'date') arts = arts.slice().sort(function(a, b){ return a.date < b.date ? 1 : -1; });
    countEl.textContent = arts.length + (arts.length > 1 ? ' articles affichés' : ' article affiché');
    liste.replaceChildren();
    if (!arts.length){ liste.appendChild(el('div', 'note', 'Aucun article ne correspond à ces filtres.')); return; }

    arts.forEach(function(a){
      const item = el('div', 'news-item' + (a.significative ? ' news-impact' : ''));
      const meta = el('div', 'news-meta');
      meta.appendChild(el('span', null, fmtDate(a.date)));
      meta.appendChild(el('span', null, '·'));
      meta.appendChild(el('span', null, a.source || ''));
      a.cibles.forEach(function(c){
        meta.appendChild(el('span', 'tag' + (c.theme ? ' tag-theme' : ''),
          (c.theme ? c.nom : c.code) + ' · ' + c.poids.toFixed(1) + ' %'));
      });
      a.types.forEach(function(t){ meta.appendChild(el('span', 'tag tag-type', t)); });
      if (a.mention === 'indirecte') meta.appendChild(el('span', 'tag tag-type', 'mention indirecte'));
      const cls = a.sentiment_label === 'positif' ? 'pill-good' : (a.sentiment_label === 'négatif' ? 'pill-bad' : 'pill-mid');
      meta.appendChild(el('span', 'pill-status ' + cls, a.sentiment_label));
      meta.appendChild(el('span', 'news-score', 'pertinence ' + Math.round(a.score)));
      item.appendChild(meta);

      const lien = el('a', 'news-title', a.titre);
      if (/^https?:\\/\\//i.test(a.lien)){ lien.href = a.lien; lien.target = '_blank'; lien.rel = 'noopener noreferrer'; }
      item.appendChild(lien);
      if (a.resume) item.appendChild(el('p', 'news-resume', a.resume));

      const r = a.reaction;
      if (r){
        const ligne = el('div', 'news-reaction');
        if (r.en_attente){
          ligne.textContent = 'Réaction du cours : séance à venir';
        } else {
          ligne.appendChild(el('span', null, 'Réaction ' + r.ticker + ' le ' + r.date.slice(8, 10) + '/' + r.date.slice(5, 7) + ' : '));
          ligne.appendChild(el('span', r.rendement >= 0 ? 'pos' : 'neg', pctTxt(r.rendement)));
          ligne.appendChild(el('span', null, ' (marché ' + pctTxt(r.marche) + ', écart '));
          ligne.appendChild(el('span', r.ecart >= 0 ? 'pos' : 'neg', pctTxt(r.ecart)));
          ligne.appendChild(el('span', null, r.z === null ? ')' : ', ' + (r.z >= 0 ? '+' : '') + r.z.toFixed(1) + ' σ)'));
          if (r.significative) ligne.appendChild(el('span', 'tag tag-impact', 'mouvement significatif'));
        }
        item.appendChild(ligne);
      }
      liste.appendChild(item);
    });
  }
  [selCible, selType, selTri, chkSignif].forEach(function(s){ s.addEventListener('change', render); });
  render();
}

function initHedge(id){
  const data = window.__hedgeData[id];
  const sel = document.getElementById(id+'-instr');
  const ratio = document.getElementById(id+'-ratio');
  const ratioVal = document.getElementById(id+'-ratio-value');
  const lev = document.getElementById(id+'-levier');
  const resultsEl = document.getElementById(id+'-results');
  const stressEl = document.getElementById(id+'-stress');
  if (!data || !sel || !ratio || !lev) return;

  function card(label, value, sub, cls){
    return '<div class="metric-card"><div class="k">' + label + '</div><div class="v">' + value + '</div>' +
      (sub ? '<div class="metric-delta ' + (cls||'') + '">' + sub + '</div>' : '') + '</div>';
  }
  function compare(label, oldV, newV, higherIsBetter, decimals, suffix){
    const diff = newV - oldV;
    const cls = Math.abs(diff) < 1e-9 ? '' : ((higherIsBetter ? diff >= 0 : diff <= 0) ? 'pos' : 'neg');
    return card(label, newV.toFixed(decimals) + suffix +
      ' <small style="color:var(--slate); font-size:11px;">(' + oldV.toFixed(decimals) + suffix + ' sans)</small>',
      (diff >= 0 ? '+' : '') + diff.toFixed(decimals) + suffix, cls);
  }

  function render(){
    const ins = data.instruments[sel.value];
    const taux = parseFloat(ratio.value)/100;
    ratioVal.textContent = ratio.value + ' %';
    const levier = parseFloat(lev.value);
    const V = data.val_actions;

    // Notionnel vendu : neutralise la part choisie du bêta du portefeuille sur l'instrument
    const notional = Math.max(0, taux * ins.beta_pf * V);
    const k = V > 0 ? notional / V : 0;
    const marge = notional / levier;
    // Coût annuel d'un CFD vendeur : spread - taux de référence reçu + dividendes reversés (indice de prix)
    const carryRate = data.spread - data.taux_reference + ins.dividende;
    const carryEur = notional * carryRate;
    // Liquidation (ESMA) quand le dépôt restant passe sous 50 % de la marge ; dépôt = espèces du compte
    const depot = data.especes;
    const ouvrable = depot >= marge;
    const seuilLiq = (notional > 0 && ouvrable) ? (depot - 0.5*marge) / notional : null;

    // Rejeu historique sur 12 mois, uniquement aux dates où l'instrument cote
    const base = [], hedged = [], bench = [], dates = [];
    for (let i=0; i<data.pf_returns.length; i++){
      const h = ins.returns[i];
      if (h === null || h === undefined) continue;
      const r = data.pf_returns[i];
      base.push(r);
      hedged.push(r - k*h - k*carryRate/252);
      if (data.has_bench) bench.push(data.bench_returns[i]);
      dates.push(data.dates[i]);
    }
    const tx = data.taux_sans_risque;
    const before = whatifStats(base, tx);
    const after = whatifStats(hedged, tx);
    const betaAvant = data.has_bench ? whatifBeta(base, bench) : null;
    const betaApres = data.has_bench ? whatifBeta(hedged, bench) : null;

    let alerte = '';
    if (notional > 0 && !ouvrable){
      alerte = '<div class="note" style="color:var(--brick);">La marge requise (' + fmtEur(marge) +
        ') dépasse les espèces du compte (' + fmtEur(depot) + ') : la position ne pourrait pas être ouverte sans nouveau versement.</div>';
    }
    resultsEl.innerHTML =
      '<div class="metric-row" style="margin-top:8px;">' +
        card('Notionnel vendu', fmtEur(notional), (k*100).toFixed(1) + ' % de la poche actions') +
        card('Marge requise (' + levier + ':1)', fmtEur(marge), 'Espèces disponibles : ' + fmtEur(depot)) +
        card('Coût de portage annuel', fmtEur(carryEur), (carryRate*100).toFixed(2) + ' % du notionnel', carryEur > 0 ? 'neg' : 'pos') +
        card('Hausse déclenchant la liquidation', seuilLiq === null ? '—' : '+' + (seuilLiq*100).toFixed(1) + ' %',
             (notional > 0 && !ouvrable) ? 'Marge supérieure aux espèces' : 'de l\\u2019indice, sans nouveau versement',
             (notional > 0 && !ouvrable) || (seuilLiq !== null && seuilLiq < 0.10) ? 'neg' : '') +
      '</div>' +
      '<div class="metric-row">' +
        compare('Volatilité annualisée', before.annualVol*100, after.annualVol*100, false, 2, ' %') +
        compare('Max drawdown (1 an)', before.maxDD, after.maxDD, true, 2, ' %') +
        compare('Rendement annualisé', before.annualReturn*100, after.annualReturn*100, true, 2, ' %') +
        (data.has_bench ? compare('Bêta vs benchmark', betaAvant, betaApres, false, 2, '') : card('Bêta vs benchmark', 'n/d')) +
      '</div>' +
      '<div class="metric-row">' +
        card('Corrélation portefeuille / instrument', ins.corr.toFixed(2),
             'Part de la variance couvrable : ' + (ins.corr*ins.corr*100).toFixed(0) + ' %') +
        card('Bêta du portefeuille sur l\\u2019instrument', ins.beta_pf.toFixed(2)) +
      '</div>' + alerte;

    // Valeur cumulée de la poche actions, sans / avec couverture
    const vBase = [], vHedge = [];
    let a = V, b = V;
    for (let i=0; i<base.length; i++){ a *= (1+base[i]); b *= (1+hedged[i]); vBase.push(a); vHedge.push(b); }
    const traces = [
      {x: dates, y: vBase, mode:'lines', name:'Sans couverture', line:{color:'#C9A15E', width:2}},
      {x: dates, y: vHedge, mode:'lines', name:'Avec couverture', line:{color:'#7BAE95', width:2}},
    ];
    const layout = {
      paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)',
      font:{family:'IBM Plex Sans, sans-serif', color:'#8B96A8', size:11},
      xaxis:{gridcolor:'#2B3D59', zeroline:false, tickfont:{family:'IBM Plex Mono, monospace', size:10}},
      yaxis:{gridcolor:'#2B3D59', zeroline:false, tickfont:{family:'IBM Plex Mono, monospace', size:10}, tickprefix:'€'},
      legend:{font:{color:'#EFE9DB', size:10}, orientation:'h', y:1.14, x:0},
      margin:{l:50,r:20,t:30,b:40}, height:320,
    };
    if (window.Plotly) Plotly.react(id+'-chart', traces, layout, {displayModeBar:false, responsive:true});

    // Stress tests avec / sans couverture (choc indice = choc marché x bêta de l'indice vs benchmark)
    const total = V + depot;
    let rows = '';
    data.scenarios.forEach(function(sc){
      const s = sc[1];
      const actions = V * Math.max(s * data.beta_stress, -1);
      const couverture = -notional * Math.max(s * ins.beta_bench, -1);
      const net = actions + couverture;
      const liquide = notional > 0 && couverture < 0 && (!ouvrable || (depot + couverture) < 0.5*marge);
      const col = (v, txt) => '<span class="' + (v >= 0 ? 'pos' : 'neg') + '">' + txt + '</span>';
      const eurS = (v) => (v >= 0 ? '+' : '') + fmtEur(v);
      rows += '<tr><td>' + sc[0] + '</td>' +
        '<td class="num">' + col(s, fmtPct(s*100)) + '</td>' +
        '<td class="num">' + col(actions, eurS(actions)) + '</td>' +
        '<td class="num">' + col(couverture, eurS(couverture)) + '</td>' +
        '<td class="num">' + col(net, eurS(net)) + '</td>' +
        '<td class="num">' + col(net, total > 0 ? fmtPct(net/total*100) : '—') + '</td>' +
        '<td class="num">' + fmtEur(total + net) + '</td>' +
        '<td>' + (liquide ? '<span class="neg">Liquidation probable</span>' : '—') + '</td></tr>';
    });
    stressEl.innerHTML = '<table class="ledger"><thead><tr>' +
      '<th>Scénario</th><th class="num">Choc marché</th><th class="num">Actions</th>' +
      '<th class="num">Couverture</th><th class="num">Impact net</th><th class="num">Impact net %</th>' +
      '<th class="num">Valeur résiduelle</th><th>Marge</th></tr></thead><tbody>' + rows + '</tbody></table>';
  }

  sel.addEventListener('change', render);
  ratio.addEventListener('input', render);
  lev.addEventListener('change', render);
  render();
}

function initMonteCarlo(id){
  const data = window.__mcData[id];
  const yearsSlider = document.getElementById(id+'-years');
  const simsSlider = document.getElementById(id+'-sims');
  const yearsVal = document.getElementById(id+'-years-value');
  const simsVal = document.getElementById(id+'-sims-value');
  const btn = document.getElementById(id+'-recalc');
  const status = document.getElementById(id+'-status');
  const apportSlider = document.getElementById(id+'-apport');
  const apportVal = document.getElementById(id+'-apport-value');

  function syncLabels(){
    yearsVal.textContent = yearsSlider.value + ' ans';
    simsVal.textContent = parseInt(simsSlider.value).toLocaleString('fr-FR') + ' tirages';
    if(apportSlider) apportVal.textContent = apportSlider.value + ' €';
  }
  yearsSlider.addEventListener('input', syncLabels);
  simsSlider.addEventListener('input', syncLabels);
  if(apportSlider) apportSlider.addEventListener('input', syncLabels);
  syncLabels();

  function run(){
    btn.disabled = true;
    status.textContent = 'Calcul en cours…';
    
    setTimeout(function(){
      try {
        const years = parseInt(yearsSlider.value);
        const nSims = parseInt(simsSlider.value);
        const apportMensuel = apportSlider ? parseFloat(apportSlider.value) : 0;
        const days = years*252;
        const out = data.method === 'garch'
          ? runGarchSim(data, days, nSims, apportMensuel)
          : runBootstrapSim(data, days, nSims, apportMensuel);

        const dates = [];
        const d = new Date();
        let added = 0;
        while (added < days){
          d.setDate(d.getDate()+1);
          const day = d.getDay();
          if (day !== 0 && day !== 6){ dates.push(d.toISOString().slice(0,10)); added++; }
        }

        const traces = [
          {x: dates.concat(dates.slice().reverse()), y: out.p90.concat(out.p10.slice().reverse()),
           fill:'toself', fillcolor:'rgba(201,161,94,0.15)', line:{color:'rgba(0,0,0,0)'},
           name:'Intervalle 80%', hoverinfo:'skip'},
          {x: dates.concat(dates.slice().reverse()), y: out.p75.concat(out.p25.slice().reverse()),
           fill:'toself', fillcolor:'rgba(201,161,94,0.28)', line:{color:'rgba(0,0,0,0)'},
           name:'Intervalle 50%', hoverinfo:'skip'},
          {x: dates, y: out.median, mode:'lines', name:'Scénario médian', line:{color:'#EFE9DB', width:2}},
        ];
        const layout = {
          paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)',
          font:{family:'IBM Plex Sans, sans-serif', color:'#8B96A8', size:11},
          xaxis:{gridcolor:'#2B3D59', zeroline:false, tickfont:{family:'IBM Plex Mono, monospace', size:10}},
          yaxis:{gridcolor:'#2B3D59', zeroline:false, tickfont:{family:'IBM Plex Mono, monospace', size:10}, tickprefix:'€'},
          legend:{font:{color:'#EFE9DB', size:10}, orientation:'h', y:1.14, x:0},
          margin:{l:50,r:20,t:30,b:40}, height:380,
        };
        Plotly.react(id+'-chart', traces, layout, {displayModeBar:false, responsive:true});

        const initial = data.initial_value;
        function setStat(key, val){
          const el = document.getElementById(id+'-stat-'+key);
          if (!el) return;
          const chg = (val/initial - 1)*100;
          el.querySelector('.v').textContent = fmtEur(val);
          const deltaEl = el.querySelector('.metric-delta');
          if (deltaEl){
            deltaEl.textContent = fmtPct(chg);
            deltaEl.className = 'metric-delta ' + (chg>=0 ? 'pos' : 'neg');
          }
        }
        const last = (arr)=>arr[arr.length-1];
        setStat('p1', last(out.p1)); setStat('p10', last(out.p10));
        setStat('p25', last(out.p25)); setStat('median', last(out.median));
        setStat('p75', last(out.p75)); setStat('p90', last(out.p90));
        setStat('p99', last(out.p99));

        status.textContent = '';
        btn.disabled = false;
      } catch (e) {
        status.textContent = 'Erreur: ' + e.message;
        console.error(e);
        btn.disabled = false;
      }
    }, 30);
  }
  btn.addEventListener('click', run);

  // Pas de calcul au chargement de la page : le premier run() est déclenché par showTab à la
  // première ouverture de l'onglet qui contient la projection (une seule fois), puis via le bouton.
  const chartEl = document.getElementById(id+'-chart');
  const panel = chartEl ? chartEl.closest('.tabpanel') : null;
  const job = { panelId: panel ? panel.id : null, run: run, done: false };
  window.__mcLazy = window.__mcLazy || {};
  window.__mcLazy[id] = job;
  status.textContent = 'Calcul lancé à l\\u2019ouverture de l\\u2019onglet.';
  // Projection déjà visible au chargement (onglet actif ou hors onglet) : lancement immédiat
  if (!panel || panel.classList.contains('active')){ job.done = true; run(); }
}

function initScreener(id) {
    const data = window.__screenerData;
    if (!data) return;
    const perMaxInput = document.getElementById(id+'-per');
    const roeMinInput = document.getElementById(id+'-roe');
    const debtMaxInput = document.getElementById(id+'-debt');
    const sectorInput = document.getElementById(id+'-sector');
    const tbody = document.getElementById(id+'-tbody');
    const countSpan = document.getElementById(id+'-count');

    if (sectorInput) {
        const sectors = Array.from(new Set(data.map(r => r.Secteur || 'Inconnu'))).sort();
        sectors.forEach(s => {
            const opt = document.createElement('option');
            opt.value = s; opt.textContent = s;
            sectorInput.appendChild(opt);
        });
    }

    function render() {
        const perMax = parseFloat(perMaxInput.value);
        const roeMin = parseFloat(roeMinInput.value);
        const debtMax = parseFloat(debtMaxInput.value);
        const sectorFilter = sectorInput ? sectorInput.value : '';
        
        document.getElementById(id+'-per-val').textContent = perMax;
        document.getElementById(id+'-roe-val').textContent = roeMin + '%';
        document.getElementById(id+'-debt-val').textContent = debtMax;

        let html = '';
        let count = 0;

        const sortedData = data.slice().sort((a, b) => (b.Mom_6M || 0) - (a.Mom_6M || 0));

        sortedData.forEach(r => {
            const per = r.PER;
            const roe = r.ROE;
            const debt = r.Dette_Eq;
            const croiss = r.Croiss_CA;
            const secteur = r.Secteur || 'Inconnu';

            if (sectorFilter && secteur !== sectorFilter) return;
            if (per === null || per > perMax) return;
            if (roe === null || roe < roeMin) return;
            if (debt === null || debt > debtMax) return;
            if (croiss === null || croiss <= 0) return; 

            count++;
            
            const fmtDec = (v, dec) => v === null ? '-' : v.toFixed(dec);
            const fmtPctCol = (v) => v === null ? '-' : `<span class="${v>=0?'pos':'neg'}">${v>0?'+':''}${v.toFixed(1)} %</span>`;

            html += `<tr>
                <td><b>${r.Ticker}</b></td>
                <td>${secteur}</td>
                <td class="num">${fmtDec(per, 1)}</td>
                <td class="num">${fmtDec(roe, 1)} %</td>
                <td class="num">${fmtPctCol(croiss)}</td>
                <td class="num">${fmtDec(debt, 2)}</td>
                <td class="num">${fmtPctCol(r.Mom_6M)}</td>
                <td class="num">${fmtPctCol(r.Dist_MM200)}</td>
            </tr>`;
        });

        if(count === 0) {
            html = `<tr><td colspan="8" style="text-align:center; padding:20px; color:var(--slate);">Aucune entreprise ne respecte ces critères stricts.</td></tr>`;
        }
        tbody.innerHTML = html;
        countSpan.textContent = count;
    }

    perMaxInput.addEventListener('input', render);
    roeMinInput.addEventListener('input', render);
    debtMaxInput.addEventListener('input', render);
    if (sectorInput) sectorInput.addEventListener('change', render);
    render();
}
"""

DYNAMIC_JS = MC_ENGINE_JS + INTERACTIVE_JS

HEAD = """<!DOCTYPE html>
<html lang="fr"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>{title}</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,400;9..144,500;9..144,600&family=IBM+Plex+Mono:wght@400;500;600&family=IBM+Plex+Sans:wght@400;500;600&display=swap" rel="stylesheet">
<script src="https://cdn.plot.ly/plotly-2.32.0.min.js"></script>
<script>{dynamic_js}</script>
<style>{css}</style></head><body>"""

def build_report(tabs: dict, meta: dict, alerts: list = None) -> str:
    nav_buttons = []
    panels = []
    for i, (tab_id, t) in enumerate(tabs.items()):
        active = "active" if i == 0 else ""
        nav_buttons.append(f'<button class="{active}" onclick="showTab(\'{tab_id}\',this)">{t["label"]}</button>')
        panels.append(f'<div id="{tab_id}" class="tabpanel {active}">{t["html"]}</div>')

    alerts_html = ""
    if alerts:
        lis = "".join(f"<li>{a}</li>" for a in alerts)
        alerts_html = f'<div class="alert-banner"><ul>{lis}</ul></div>'

    html = HEAD.format(title=meta.get("title", "Patrimoine"), css=CSS, dynamic_js=DYNAMIC_JS)
    html += f"""
    <div class="sheet">
      <div class="masthead">
        <div><div class="eyebrow">Patrimoine — Relevé privé</div><h1>Tableau de bord patrimonial</h1></div>
        <div class="right"><div><span class="status-dot"></span>Marchés — données actualisées</div>
        <div>Relevé au {meta.get('date_str', datetime.now().strftime('%d %B %Y'))}</div></div>
      </div>
      {alerts_html}
      <div class="tabnav">{''.join(nav_buttons)}</div>
      {''.join(panels)}
      <footer>{meta.get('footer', '')}</footer>
    </div>
    <script>{TAB_JS}</script>
    </body></html>"""
    return html