"""Noon Academy Monthly Financials Dashboard
Run:  streamlit run app.py
"""

import json
from pathlib import Path

import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st

# ── Brand tokens ──────────────────────────────────────────────────────────────
DARK_BLUE   = "#11203A"
GREEN       = "#17E4A1"
GREEN_D     = "#0FB07F"
ORANGE      = "#FE814F"
PURPLE      = "#7C5CE1"
CREAM_LIGHT = "#FAF6EC"
GRAY_DARK   = "#5A6878"
GRAY_MID    = "#9AA4B0"
GRAY_LIGHT  = "#D6D6D6"

BU_COLORS = {
    "Tracks":            GREEN,
    "Government Schools": ORANGE,
    "B2B":               PURPLE,
    "KSA B2C":           "#FE4F81",
    "Global Non-Saudi":  "#4FC8FE",
    "Out of School":     "#FEC84F",
}
BU_ORDER = list(BU_COLORS.keys())

SAR_RATE = 3.75   # 1 USD = 3.75 SAR

# ── Weekly CF data (hardcoded from "Noon Weekly CF Update May 3, 2026" PDF)
CF_OPENING_SAR = 194_000
CF_INFLOWS_SAR = {
    "Tracks Collections": 800_000,
    "B2B Collections":    4_100_000,
    "Financing":          7_700_000,
    "Gov. Grants":        0,
}
CF_OUTFLOWS_SAR = {
    "Payroll & Benefits":    5_800_000,
    "AP Payments":           3_100_000,
    "Taxes & Gov. Charges":  3_000_000,
    "Financing Repayments":  2_600_000,
}

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Noon Academy | Monthly Financials",
    page_icon="🌙",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@300;400;500;600;700&display=swap');
html, body, [class*="css"], .stApp {{
    font-family: 'IBM Plex Sans', 'Segoe UI', sans-serif;
    background-color: {CREAM_LIGHT};
}}
.block-container {{ padding-top: 1rem; padding-bottom: 1rem; max-width: 1400px; }}
.kpi-card {{
    background: white;
    border-radius: 10px;
    padding: 14px 18px 12px 18px;
    border-left: 4px solid {GREEN};
    box-shadow: 0 1px 4px rgba(0,0,0,0.07);
}}
.kpi-card-dark {{
    background: {DARK_BLUE};
    border-radius: 10px;
    padding: 14px 18px 12px 18px;
    border-left: 4px solid {GREEN};
    box-shadow: 0 1px 4px rgba(0,0,0,0.15);
}}
.kpi-label {{
    font-size: 10px; font-weight: 600;
    text-transform: uppercase; letter-spacing: 0.07em;
    color: {GRAY_DARK}; margin-bottom: 4px;
}}
.kpi-label-dark {{
    font-size: 10px; font-weight: 600;
    text-transform: uppercase; letter-spacing: 0.07em;
    color: {GRAY_MID}; margin-bottom: 4px;
}}
.kpi-value {{
    font-size: 22px; font-weight: 700;
    color: {DARK_BLUE}; line-height: 1.1;
}}
.kpi-value-dark {{
    font-size: 22px; font-weight: 700;
    color: white; line-height: 1.1;
}}
.kpi-sub {{
    font-size: 11px; margin-top: 4px; color: {GRAY_MID};
}}
.section-hdr {{
    font-size: 16px; font-weight: 700;
    color: {DARK_BLUE}; margin: 20px 0 8px 0;
    padding-bottom: 6px;
    border-bottom: 2px solid {GREEN};
}}
.chart-note {{
    font-size: 10px; color: {GRAY_MID}; margin-top: -10px; margin-bottom: 4px;
}}
</style>
""", unsafe_allow_html=True)


# ── Data ──────────────────────────────────────────────────────────────────────
@st.cache_data
def load_data():
    p = Path("fy_data.json")
    if not p.exists():
        st.error("fy_data.json not found — run `python etl.py` first.")
        st.stop()
    with open(p) as f:
        return json.load(f)


DATA   = load_data()
ALL_BUS = DATA["buses"]
FY2526  = DATA["fy2526"]
CASH    = DATA["cash"]


# ── Helpers ───────────────────────────────────────────────────────────────────
def fmt(val, unit_mul=1, symbol="$"):
    v = val * unit_mul
    if abs(v) >= 1_000_000:
        return f"{symbol}{v/1_000_000:.1f}M"
    if abs(v) >= 1_000:
        return f"{symbol}{v/1_000:.0f}K"
    return f"{symbol}{v:.0f}"


def group_total(field, n):
    return sum(
        m["by_bu"].get(bu, {}).get(field, 0)
        for m in FY2526[:n]
        for bu in ALL_BUS
    )


def monthly_by_bu(field):
    return {
        bu: [m["by_bu"].get(bu, {}).get(field, 0) for m in FY2526]
        for bu in BU_ORDER if bu in ALL_BUS
    }


def monthly_total(field):
    return [
        sum(m["by_bu"].get(bu, {}).get(field, 0) for bu in ALL_BUS)
        for m in FY2526
    ]


def pct_series(num_field, den_field):
    result = []
    for m in FY2526:
        num = sum(m["by_bu"].get(bu, {}).get(num_field, 0) for bu in ALL_BUS)
        den = sum(m["by_bu"].get(bu, {}).get(den_field, 0) for bu in ALL_BUS)
        result.append(num / den * 100 if den else 0)
    return result


def chart_layout(height=300):
    return dict(
        height=height,
        margin=dict(l=8, r=8, t=10, b=30),
        paper_bgcolor="white",
        plot_bgcolor="white",
        font=dict(family="IBM Plex Sans", color=GRAY_DARK, size=11),
        legend=dict(orientation="h", x=0, y=-0.18, font=dict(size=10)),
        xaxis=dict(gridcolor=GRAY_LIGHT, showgrid=False, zeroline=False, tickfont=dict(size=10)),
        yaxis=dict(gridcolor=GRAY_LIGHT, showgrid=True, zeroline=False, tickfont=dict(size=10)),
        showlegend=True,
    )


# ── Header ────────────────────────────────────────────────────────────────────
st.markdown(f"""
<div style="background:{DARK_BLUE}; border-radius:12px; padding:14px 24px 12px 24px;
            margin-bottom:16px; display:flex; align-items:center; justify-content:space-between;">
  <div>
    <span style="font-size:20px; font-weight:700; color:white;">
      Noon Academy — Monthly Financials
    </span>
    <span style="font-size:11px; color:{GRAY_MID}; display:block; margin-top:3px;">
      FY 2025/26 · Internal Management Dashboard · Confidential
    </span>
  </div>
  <span style="font-size:30px; font-weight:800; color:{GREEN}; letter-spacing:-0.03em;">noon</span>
</div>
""", unsafe_allow_html=True)

# Currency toggle
col_toggle, col_month = st.columns([2, 5])
with col_toggle:
    currency = st.radio("Currency", ["USD", "SAR"], horizontal=True, label_visibility="collapsed")
mul = SAR_RATE if currency == "SAR" else 1.0
sym = "SAR " if currency == "SAR" else "$"

with col_month:
    n = st.slider("Months elapsed (FY 2025/26)", 1, 12, 9,
                  help="Jul-25 = 1, Mar-26 = 9, Jun-26 = 12")
    st.caption(f"Jul-25 → {FY2526[n-1]['label']}")

month_labels = [m["label"] for m in FY2526]

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 1 — FINANCIAL PERFORMANCE REVIEW
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-hdr">Financial Performance Review</div>', unsafe_allow_html=True)

# ── 7 KPI boxes ───────────────────────────────────────────────────────────────
rev_ytd   = group_total("revenue", n) * mul
gp_ytd    = group_total("gross_profit", n) * mul
cm_ytd    = group_total("contribution_margin", n) * mul
ebi_ytd   = group_total("ebitda", n) * mul
gm_pct    = (group_total("gross_profit", n) / group_total("revenue", n) * 100) if group_total("revenue", n) else 0
cm_pct_v  = (group_total("contribution_margin", n) / group_total("revenue", n) * 100) if group_total("revenue", n) else 0
ebi_pct_v = (group_total("ebitda", n) / group_total("revenue", n) * 100) if group_total("revenue", n) else 0

k1, k2, k3, k4, k5, k6, k7 = st.columns(7)

def kpi_html(label, value, sub=None, accent=GREEN, dark=False):
    card_cls = "kpi-card-dark" if dark else "kpi-card"
    lbl_cls  = "kpi-label-dark" if dark else "kpi-label"
    val_cls  = "kpi-value-dark" if dark else "kpi-value"
    sub_html = f'<div class="kpi-sub">{sub}</div>' if sub else ""
    return (
        f'<div class="{card_cls}" style="border-left-color:{accent}">'
        f'<div class="{lbl_cls}">{label}</div>'
        f'<div class="{val_cls}">{value}</div>'
        f'{sub_html}</div>'
    )

k1.markdown(kpi_html("Revenue YTD",          fmt(rev_ytd, symbol=sym),  f"as of {FY2526[n-1]['label']}"), unsafe_allow_html=True)
k2.markdown(kpi_html("% Budget Achieved",    "—",                        "Budget TBD",        accent=GRAY_MID), unsafe_allow_html=True)
k3.markdown(kpi_html("Gross Profit YTD",     fmt(gp_ytd, symbol=sym),   f"GM {gm_pct:.1f}%", accent=GREEN_D), unsafe_allow_html=True)
k4.markdown(kpi_html("Contribution Profit",  fmt(cm_ytd, symbol=sym),   f"CM {cm_pct_v:.1f}%", accent=PURPLE), unsafe_allow_html=True)
k5.markdown(kpi_html("EBITDA YTD",           fmt(ebi_ytd, symbol=sym),  None,                accent=ORANGE), unsafe_allow_html=True)
k6.markdown(kpi_html("EBITDA %",             f"{ebi_pct_v:.1f}%",       None,                accent=ORANGE), unsafe_allow_html=True)
k7.markdown(kpi_html("Headcount",            "—",                        "Data pending",      accent=GRAY_MID), unsafe_allow_html=True)

st.markdown("<div style='height:8px'/>", unsafe_allow_html=True)

# ── Chart 1: Revenue by BU — stacked bar ─────────────────────────────────────
active_bus = [bu for bu in BU_ORDER if bu in ALL_BUS]
rev_by_bu  = monthly_by_bu("revenue")
rev_totals = [sum(rev_by_bu.get(bu, [0]*12)[i] * mul for bu in active_bus) for i in range(12)]

fig1 = go.Figure()
for bu in active_bus:
    vals = [v * mul for v in rev_by_bu.get(bu, [0]*12)]
    fig1.add_trace(go.Bar(
        name=bu, x=month_labels, y=vals,
        marker_color=BU_COLORS.get(bu, GRAY_MID),
        text=[f"{sym}{v/1e6:.1f}M" if v > 0 else "" for v in vals],
        textposition="inside",
        textfont=dict(size=9, color="white"),
        hovertemplate=f"{bu}: {sym}%{{y:,.0f}}<extra></extra>",
    ))

# Total annotations at top of each bar
for i, (label, total) in enumerate(zip(month_labels, rev_totals)):
    fig1.add_annotation(
        x=label, y=total,
        text=f"<b>{sym}{total/1e6:.1f}M</b>",
        showarrow=False, yanchor="bottom",
        font=dict(size=9, color=DARK_BLUE, family="IBM Plex Sans"),
        yshift=2,
    )

lo1 = chart_layout(height=310)
lo1["barmode"] = "stack"
lo1["yaxis"]["tickprefix"] = sym
lo1["yaxis"]["showgrid"] = True
lo1["xaxis"]["showgrid"] = False
lo1["legend"] = dict(orientation="h", x=0, y=-0.15, font=dict(size=10))
fig1.update_layout(**lo1)
st.plotly_chart(fig1, use_container_width=True, config={"displayModeBar": False})
st.markdown('<div class="chart-note">Revenue by Business Unit — monthly stacked bars with data labels</div>', unsafe_allow_html=True)


# ── Chart 2: Gross Profit bar + GM% line ─────────────────────────────────────
gp_monthly = [v * mul for v in monthly_total("gross_profit")]
gm_pct_mo  = pct_series("gross_profit", "revenue")

fig2 = make_subplots(specs=[[{"secondary_y": True}]])
fig2.add_trace(go.Bar(
    name="Gross Profit", x=month_labels, y=gp_monthly,
    marker_color=GREEN_D, opacity=0.85,
    hovertemplate=f"Gross Profit: {sym}%{{y:,.0f}}<extra></extra>",
), secondary_y=False)
fig2.add_trace(go.Scatter(
    name="GM %", x=month_labels, y=gm_pct_mo,
    mode="lines+markers",
    line=dict(color=DARK_BLUE, width=2),
    marker=dict(size=5, color=DARK_BLUE),
    hovertemplate="GM%%: %{y:.1f}%%<extra></extra>",
), secondary_y=True)

lo2 = chart_layout(height=280)
lo2.pop("xaxis", None); lo2.pop("yaxis", None)
fig2.update_layout(**lo2)
fig2.update_yaxes(title_text=f"Gross Profit ({currency})", tickprefix=sym, gridcolor=GRAY_LIGHT, secondary_y=False)
fig2.update_yaxes(title_text="GM %", ticksuffix="%", showgrid=False, secondary_y=True)
fig2.update_xaxes(showgrid=False, tickfont=dict(size=10))
st.plotly_chart(fig2, use_container_width=True, config={"displayModeBar": False})
st.markdown('<div class="chart-note">Gross Profit (bar) and Gross Margin % (line, right axis)</div>', unsafe_allow_html=True)


# ── Chart 3: Contribution Profit bar + CM% line ───────────────────────────────
cm_monthly = [v * mul for v in monthly_total("contribution_margin")]
cm_pct_mo  = pct_series("contribution_margin", "revenue")

fig3 = make_subplots(specs=[[{"secondary_y": True}]])
fig3.add_trace(go.Bar(
    name="Contribution Profit", x=month_labels, y=cm_monthly,
    marker_color=PURPLE, opacity=0.85,
    hovertemplate=f"Contribution Profit: {sym}%{{y:,.0f}}<extra></extra>",
), secondary_y=False)
fig3.add_trace(go.Scatter(
    name="CM %", x=month_labels, y=cm_pct_mo,
    mode="lines+markers",
    line=dict(color=DARK_BLUE, width=2),
    marker=dict(size=5, color=DARK_BLUE),
    hovertemplate="CM%%: %{y:.1f}%%<extra></extra>",
), secondary_y=True)

lo3 = chart_layout(height=280)
lo3.pop("xaxis", None); lo3.pop("yaxis", None)
fig3.update_layout(**lo3)
fig3.update_yaxes(title_text=f"Contribution Profit ({currency})", tickprefix=sym, gridcolor=GRAY_LIGHT, secondary_y=False)
fig3.update_yaxes(title_text="CM %", ticksuffix="%", showgrid=False, secondary_y=True)
fig3.update_xaxes(showgrid=False, tickfont=dict(size=10))
st.plotly_chart(fig3, use_container_width=True, config={"displayModeBar": False})
st.markdown('<div class="chart-note">Contribution Profit (bar) and Contribution Margin % (line, right axis)</div>', unsafe_allow_html=True)


# ── Chart 4: EBITDA bar (green/orange conditional) + EBITDA% line ─────────────
ebi_monthly = [v * mul for v in monthly_total("ebitda")]
ebi_pct_mo  = pct_series("ebitda", "revenue")
ebi_colors  = [GREEN if v >= 0 else ORANGE for v in ebi_monthly]

fig4 = make_subplots(specs=[[{"secondary_y": True}]])
fig4.add_trace(go.Bar(
    name="EBITDA", x=month_labels, y=ebi_monthly,
    marker_color=ebi_colors, opacity=0.88,
    hovertemplate=f"EBITDA: {sym}%{{y:,.0f}}<extra></extra>",
), secondary_y=False)
fig4.add_trace(go.Scatter(
    name="EBITDA %", x=month_labels, y=ebi_pct_mo,
    mode="lines+markers",
    line=dict(color=DARK_BLUE, width=2),
    marker=dict(size=5, color=DARK_BLUE),
    hovertemplate="EBITDA%%: %{y:.1f}%%<extra></extra>",
), secondary_y=True)
# Zero reference line
fig4.add_shape(type="line", x0=month_labels[0], x1=month_labels[-1], y0=0, y1=0,
               line=dict(color=GRAY_MID, width=1, dash="dot"), layer="below")

lo4 = chart_layout(height=280)
lo4.pop("xaxis", None); lo4.pop("yaxis", None)
fig4.update_layout(**lo4)
fig4.update_yaxes(title_text=f"EBITDA ({currency})", tickprefix=sym, gridcolor=GRAY_LIGHT, secondary_y=False)
fig4.update_yaxes(title_text="EBITDA %", ticksuffix="%", showgrid=False, secondary_y=True)
fig4.update_xaxes(showgrid=False, tickfont=dict(size=10))
st.plotly_chart(fig4, use_container_width=True, config={"displayModeBar": False})
st.markdown('<div class="chart-note">EBITDA (green = positive, orange = negative) and EBITDA Margin % (line, right axis)</div>', unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 2 — CASH POSITION
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-hdr">Cash Position</div>', unsafe_allow_html=True)

# ── 3 KPI boxes ───────────────────────────────────────────────────────────────
cash_slice  = CASH[:n]
last_cash   = cash_slice[-1]["eom_cash"] if cash_slice else 0

if len(cash_slice) >= 2:
    recent = cash_slice[max(0, len(cash_slice) - 3):]
    burns  = [recent[i]["eom_cash"] - recent[i-1]["eom_cash"] for i in range(1, len(recent))]
    avg_burn = sum(burns) / len(burns) if burns else 0
    runway_mo = round(-last_cash / avg_burn, 1) if avg_burn < 0 else None
else:
    avg_burn  = 0
    runway_mo = None

cash_display = last_cash * mul
burn_display = avg_burn * mul
runway_str   = f"{runway_mo} months" if runway_mo else "—"

ck1, ck2, ck3, _sp = st.columns([2, 2, 2, 1])
ck1.markdown(kpi_html("Current Cash Balance", fmt(cash_display, symbol=sym), f"as of {FY2526[n-1]['label']}", accent=GREEN, dark=True), unsafe_allow_html=True)
ck2.markdown(kpi_html("Avg Monthly Cash Burn", fmt(burn_display, symbol=sym), "trailing 3 months", accent=ORANGE, dark=True), unsafe_allow_html=True)
ck3.markdown(kpi_html("Runway", runway_str, "at current burn rate", accent="#4FC8FE", dark=True), unsafe_allow_html=True)

st.markdown("<div style='height:12px'/>", unsafe_allow_html=True)

# ── Next 4 weeks: Inflows vs Outflows (from CF weekly PDF) ───────────────────
cf_mul = SAR_RATE if currency == "USD" else 1.0
# CF data is in SAR → convert to USD (÷3.75) if USD selected, keep SAR if SAR selected
inflow_vals  = [v / SAR_RATE * mul for v in CF_INFLOWS_SAR.values()]
outflow_vals = [v / SAR_RATE * mul for v in CF_OUTFLOWS_SAR.values()]
inflow_cats  = list(CF_INFLOWS_SAR.keys())
outflow_cats = list(CF_OUTFLOWS_SAR.keys())

st.markdown(f'<div style="font-size:13px; font-weight:600; color:{DARK_BLUE}; margin-bottom:6px;">Next 4 Weeks Cash Flow &nbsp;<span style="font-size:10px; color:{GRAY_MID}; font-weight:400;">As of 03 May 2026 · Source: Noon Weekly CF Update</span></div>', unsafe_allow_html=True)

col_in, col_out = st.columns(2)

with col_in:
    fig_in = go.Figure(go.Bar(
        y=inflow_cats, x=inflow_vals,
        orientation="h",
        marker_color=GREEN, opacity=0.85,
        text=[fmt(v, symbol=sym) for v in inflow_vals],
        textposition="outside",
        textfont=dict(size=10, color=DARK_BLUE),
        hovertemplate="%{y}: " + sym + "%{x:,.0f}<extra></extra>",
    ))
    lo_in = chart_layout(height=220)
    lo_in["showlegend"] = False
    lo_in["xaxis"]["tickprefix"] = sym
    lo_in["margin"] = dict(l=8, r=60, t=30, b=20)
    lo_in["title"] = dict(text="<b>Inflows</b>", font=dict(size=12, color=GREEN_D), x=0)
    fig_in.update_layout(**lo_in)
    st.plotly_chart(fig_in, use_container_width=True, config={"displayModeBar": False})

with col_out:
    fig_out = go.Figure(go.Bar(
        y=outflow_cats, x=outflow_vals,
        orientation="h",
        marker_color=ORANGE, opacity=0.85,
        text=[fmt(v, symbol=sym) for v in outflow_vals],
        textposition="outside",
        textfont=dict(size=10, color=DARK_BLUE),
        hovertemplate="%{y}: " + sym + "%{x:,.0f}<extra></extra>",
    ))
    lo_out = chart_layout(height=220)
    lo_out["showlegend"] = False
    lo_out["xaxis"]["tickprefix"] = sym
    lo_out["margin"] = dict(l=8, r=60, t=30, b=20)
    lo_out["title"] = dict(text="<b>Outflows</b>", font=dict(size=12, color=ORANGE), x=0)
    fig_out.update_layout(**lo_out)
    st.plotly_chart(fig_out, use_container_width=True, config={"displayModeBar": False})

# ── Cash Collection vs Invoiced ───────────────────────────────────────────────
st.markdown(f'<div style="font-size:13px; font-weight:600; color:{DARK_BLUE}; margin-bottom:6px;">Cash Collection vs. Invoiced &nbsp;<span style="font-size:10px; color:{GRAY_MID}; font-weight:400;">Schools + B2B billing files</span></div>', unsafe_allow_html=True)

cash_labels   = [m["label"]     for m in cash_slice]
invoiced_vals = [m["invoiced"]  * mul for m in cash_slice]
collected_vals= [m["collected"] * mul for m in cash_slice]
coll_rate     = [
    (c / i * 100) if i > 0 else 0
    for c, i in zip(collected_vals, invoiced_vals)
]

fig5 = make_subplots(specs=[[{"secondary_y": True}]])
fig5.add_trace(go.Bar(
    name="Invoiced", x=cash_labels, y=invoiced_vals,
    marker_color=PURPLE, opacity=0.80,
    hovertemplate=f"Invoiced: {sym}%{{y:,.0f}}<extra></extra>",
), secondary_y=False)
fig5.add_trace(go.Bar(
    name="Collected", x=cash_labels, y=collected_vals,
    marker_color=GREEN, opacity=0.88,
    hovertemplate=f"Collected: {sym}%{{y:,.0f}}<extra></extra>",
), secondary_y=False)
fig5.add_trace(go.Scatter(
    name="Collection Rate %", x=cash_labels, y=coll_rate,
    mode="lines+markers",
    line=dict(color=ORANGE, width=2),
    marker=dict(size=5, color=ORANGE),
    hovertemplate="Collection Rate: %{y:.1f}%%<extra></extra>",
), secondary_y=True)

lo5 = chart_layout(height=300)
lo5["barmode"] = "group"
lo5.pop("xaxis", None); lo5.pop("yaxis", None)
fig5.update_layout(**lo5)
fig5.update_yaxes(title_text=f"Amount ({currency})", tickprefix=sym, gridcolor=GRAY_LIGHT, secondary_y=False)
fig5.update_yaxes(title_text="Collection Rate %", ticksuffix="%", showgrid=False, range=[0, 120], secondary_y=True)
fig5.update_xaxes(showgrid=False, tickfont=dict(size=10))
st.plotly_chart(fig5, use_container_width=True, config={"displayModeBar": False})

if DATA.get("meta", {}).get("billing_note"):
    st.caption(DATA["meta"]["billing_note"])


# ── Footer ────────────────────────────────────────────────────────────────────
st.markdown(f"""
<hr style="border:none; border-top:1px solid {GRAY_LIGHT}; margin:16px 0 4px 0;">
<div style="font-size:10px; color:{GRAY_MID}; text-align:center; padding-bottom:6px;">
  Noon Academy · Internal Management Dashboard · Confidential
  &nbsp;|&nbsp; Data as of {DATA['meta']['generated_at']}
</div>
""", unsafe_allow_html=True)
