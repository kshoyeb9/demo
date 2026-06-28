"""Noon Academy Monthly Financials Dashboard
Run:  streamlit run app.py
"""

import json
from datetime import date
from pathlib import Path

import plotly.graph_objects as go
import streamlit as st

# ── Brand tokens ──────────────────────────────────────────────────────────────
DARK_BLUE   = "#11203A"
DARK_BLUE_D = "#0B1729"
GREEN       = "#17E4A1"
GREEN_D     = "#0FB07F"
ORANGE      = "#FE814F"
PURPLE      = "#7C5CE1"
CREAM       = "#F4EEE0"
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

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Noon Academy | Monthly Financials",
    page_icon="🌙",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:ital,wght@0,300;0,400;0,500;0,600;0,700;1,400&display=swap');
html, body, [class*="css"], .stApp {{
    font-family: 'IBM Plex Sans', 'Segoe UI', sans-serif;
    background-color: {CREAM_LIGHT};
}}
.block-container {{ padding-top: 0.75rem; padding-bottom: 1rem; }}
section[data-testid="stSidebar"] > div {{
    background-color: {DARK_BLUE};
}}
section[data-testid="stSidebar"] label,
section[data-testid="stSidebar"] p,
section[data-testid="stSidebar"] span,
section[data-testid="stSidebar"] div {{
    color: rgba(255,255,255,0.85) !important;
}}
.kpi-card {{
    background: white;
    border-radius: 10px;
    padding: 14px 18px 12px 18px;
    border-left: 4px solid {GREEN};
    box-shadow: 0 1px 4px rgba(0,0,0,0.07);
    height: 100%;
}}
.kpi-label {{
    font-size: 10px;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.07em;
    color: {GRAY_DARK};
    margin-bottom: 5px;
}}
.kpi-value {{
    font-size: 21px;
    font-weight: 700;
    color: {DARK_BLUE};
    line-height: 1.15;
}}
.kpi-delta {{ font-size: 11px; font-weight: 500; margin-top: 5px; }}
.delta-pos {{ color: {GREEN_D}; }}
.delta-neg {{ color: {ORANGE}; }}
.section-hdr {{
    font-size: 11px;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.08em;
    color: {GRAY_MID};
    margin: 10px 0 4px 0;
    padding-bottom: 4px;
    border-bottom: 1px solid {GRAY_LIGHT};
}}
</style>
""", unsafe_allow_html=True)


# ── Data ──────────────────────────────────────────────────────────────────────
@st.cache_data
def load_data():
    p = Path("fy_data.json")
    if not p.exists():
        st.error("fy_data.json not found. Run `python etl.py` to generate it, or place the sample file here.")
        st.stop()
    with open(p) as f:
        return json.load(f)


DATA    = load_data()
ALL_BUS = DATA["buses"]
FY2425  = DATA["fy2425"]
FY2526  = DATA["fy2526"]
CASH    = DATA["cash"]


# ── Helpers ───────────────────────────────────────────────────────────────────
def fmt_m(val):
    if abs(val) >= 1_000_000:
        return f"${val/1_000_000:.1f}M"
    if abs(val) >= 1_000:
        return f"${val/1_000:.0f}K"
    return f"${val:.0f}"


def group_total(fy, field, n, buses):
    return sum(
        m["by_bu"].get(bu, {}).get(field, 0)
        for m in fy[:n]
        for bu in buses
    )


def monthly_group(fy, field, buses):
    return [
        sum(m["by_bu"].get(bu, {}).get(field, 0) for bu in buses)
        for m in fy
    ]


def fytd_by_bu(fy, field, n, buses):
    return {
        bu: sum(m["by_bu"].get(bu, {}).get(field, 0) for m in fy[:n])
        for bu in buses
    }


def base_layout(title="", height=290):
    return dict(
        title=dict(
            text=title,
            font=dict(size=12, color=GRAY_DARK, family="IBM Plex Sans"),
            x=0, pad=dict(l=0, t=0),
        ),
        height=height,
        margin=dict(l=4, r=12, t=32 if title else 8, b=30),
        paper_bgcolor="white",
        plot_bgcolor="white",
        font=dict(family="IBM Plex Sans", color=GRAY_DARK, size=11),
        legend=dict(orientation="h", x=0, y=-0.14, font=dict(size=10)),
        xaxis=dict(gridcolor=GRAY_LIGHT, showgrid=True, zeroline=False, tickfont=dict(size=10)),
        yaxis=dict(gridcolor=GRAY_LIGHT, showgrid=True, zeroline=False, tickfont=dict(size=10)),
    )


# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown(f"""
    <div style="padding:6px 0 18px 0; border-bottom:1px solid rgba(255,255,255,0.12); margin-bottom:18px;">
      <span style="font-size:24px; font-weight:800; color:{GREEN}; letter-spacing:-0.02em;">noon</span>
      <span style="font-size:11px; color:{GRAY_MID}; display:block; margin-top:2px;">
        Academy · Monthly Financials
      </span>
    </div>
    """, unsafe_allow_html=True)

    today = date.today()
    if date(2025, 7, 1) <= today <= date(2026, 6, 30):
        default_n = (today.month - 7) % 12 + 1
    else:
        default_n = 12

    months_elapsed = st.slider(
        "Months elapsed · FY 2025/26",
        min_value=1, max_value=12, value=default_n,
        help="Slide to set how many FY25/26 months are shown (1 = July 2025 only)",
    )
    month_end_label = FY2526[months_elapsed - 1]["label"]
    st.caption(f"Jul-25 → {month_end_label}")

    st.markdown("<div style='height:12px'/>", unsafe_allow_html=True)

    selected_bus = st.multiselect(
        "Business Units",
        options=ALL_BUS,
        default=ALL_BUS,
    )
    if not selected_bus:
        selected_bus = ALL_BUS

    st.markdown(f"""
    <div style="margin-top:24px; font-size:10px; color:{GRAY_MID}; line-height:1.7;">
      Generated: {DATA['meta']['generated_at']}<br>
      <em>{DATA['meta']['source']}</em>
    </div>
    """, unsafe_allow_html=True)


# ── Derived KPIs ──────────────────────────────────────────────────────────────
n   = months_elapsed
bus = selected_bus

rev_26  = group_total(FY2526, "revenue",               n, bus)
gp_26   = group_total(FY2526, "gross_profit",          n, bus)
cm_26   = group_total(FY2526, "contribution_margin",   n, bus)
ebi_26  = group_total(FY2526, "ebitda",                n, bus)
rev_25  = group_total(FY2425, "revenue",               n, bus)
ebi_25  = group_total(FY2425, "ebitda",                n, bus)

gm_pct  = gp_26  / rev_26 * 100 if rev_26 else 0
cm_pct  = cm_26  / rev_26 * 100 if rev_26 else 0
ebi_pct = ebi_26 / rev_26 * 100 if rev_26 else 0

rev_yoy  = (rev_26 - rev_25) / rev_25 * 100 if rev_25 else 0
ebi_yoy  = (ebi_26 - ebi_25) / abs(ebi_25) * 100 if ebi_25 else 0

cash_slice = CASH[:n]
last_cash  = cash_slice[-1]["eom_cash"] if cash_slice else 0
if len(cash_slice) >= 2:
    recent = cash_slice[max(0, len(cash_slice) - 3):]
    burns  = [recent[i]["eom_cash"] - recent[i-1]["eom_cash"] for i in range(1, len(recent))]
    avg_burn = sum(burns) / len(burns) if burns else 0
    runway = round(-last_cash / avg_burn, 1) if avg_burn < 0 else None
else:
    runway = None


# ── Header ────────────────────────────────────────────────────────────────────
st.markdown(f"""
<div style="background:{DARK_BLUE}; border-radius:12px; padding:14px 22px 12px 22px;
            margin-bottom:12px; display:flex; align-items:center; justify-content:space-between;">
  <div>
    <span style="font-size:19px; font-weight:700; color:white; letter-spacing:-0.01em;">
      Noon Academy — Monthly Financials
    </span>
    <span style="font-size:11px; color:{GRAY_MID}; display:block; margin-top:3px;">
      FY 2025/26 · {n} month{"s" if n != 1 else ""} elapsed
      &nbsp;|&nbsp; YoY vs FY 2024/25 &nbsp;|&nbsp; USD
    </span>
  </div>
  <span style="font-size:28px; font-weight:800; color:{GREEN}; letter-spacing:-0.03em;">noon</span>
</div>
""", unsafe_allow_html=True)


# ── KPI Strip ─────────────────────────────────────────────────────────────────
def kpi(label, value, delta=None, higher_is_better=True, accent=GREEN):
    if delta is not None:
        good = (delta >= 0) == higher_is_better
        cls  = "delta-pos" if good else "delta-neg"
        sym  = "▲" if delta > 0 else "▼"
        dhtml = f'<div class="kpi-delta {cls}">{sym} {abs(delta):.1f}% YoY</div>'
    else:
        dhtml = ""
    return (
        f'<div class="kpi-card" style="border-left-color:{accent}">'
        f'<div class="kpi-label">{label}</div>'
        f'<div class="kpi-value">{value}</div>'
        f'{dhtml}</div>'
    )


runway_str = f"{runway} mo" if runway else "—"

c1, c2, c3, c4, c5, c6 = st.columns(6)
c1.markdown(kpi("Revenue · FYTD",       fmt_m(rev_26),         rev_yoy,  True,  GREEN),   unsafe_allow_html=True)
c2.markdown(kpi("Gross Margin",          f"{gm_pct:.1f}%",      None,     True,  GREEN_D), unsafe_allow_html=True)
c3.markdown(kpi("Contribution Margin",   f"{cm_pct:.1f}%",      None,     True,  PURPLE),  unsafe_allow_html=True)
c4.markdown(kpi("EBITDA Margin",         f"{ebi_pct:.1f}%",     None,     True,  ORANGE),  unsafe_allow_html=True)
c5.markdown(kpi("EBITDA · FYTD",         fmt_m(ebi_26),         ebi_yoy,  False, ORANGE),  unsafe_allow_html=True)
c6.markdown(kpi("Cash · Runway",         f"{fmt_m(last_cash)} · {runway_str}", None, True, "#4FC8FE"), unsafe_allow_html=True)

st.markdown("<div style='height:6px'/>", unsafe_allow_html=True)


# ── Row 1 · Revenue ───────────────────────────────────────────────────────────
st.markdown('<div class="section-hdr">Revenue</div>', unsafe_allow_html=True)
col_line, col_bar = st.columns([3, 2])

month_labels = [m["label"] for m in FY2425]  # Jul–Jun axis

with col_line:
    rev25_mo = monthly_group(FY2425, "revenue", bus)
    rev26_mo = monthly_group(FY2526, "revenue", bus)

    cum25 = [sum(rev25_mo[:i+1]) / 1e6 for i in range(12)]
    cum26 = [sum(rev26_mo[:i+1]) / 1e6 if i < n else None for i in range(12)]
    cum26_x = month_labels[:n]
    cum26_y = [v for v in cum26 if v is not None]

    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=month_labels, y=cum25,
        name="FY 2024/25", mode="lines",
        line=dict(color=GRAY_MID, width=2, dash="dot"),
        hovertemplate="%{x}: $%{y:.2f}M<extra>FY24/25</extra>",
    ))
    fig.add_trace(go.Scatter(
        x=cum26_x, y=cum26_y,
        name="FY 2025/26", mode="lines+markers",
        line=dict(color=GREEN, width=3),
        marker=dict(size=6, color=GREEN),
        hovertemplate="%{x}: $%{y:.2f}M<extra>FY25/26</extra>",
    ))
    lo = base_layout("Cumulative Revenue · FYTD ($M)")
    lo["yaxis"]["tickprefix"] = "$"
    lo["yaxis"]["ticksuffix"] = "M"
    fig.update_layout(**lo)
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

with col_bar:
    bu_rev = fytd_by_bu(FY2526, "revenue", n, bus)
    sorted_bus = sorted(bu, key=lambda b: bu_rev.get(b, 0))

    fig2 = go.Figure(go.Bar(
        y=sorted_bus,
        x=[bu_rev.get(b, 0) / 1e6 for b in sorted_bus],
        orientation="h",
        marker_color=[BU_COLORS.get(b, GRAY_MID) for b in sorted_bus],
        text=[fmt_m(bu_rev.get(b, 0)) for b in sorted_bus],
        textposition="outside",
        textfont=dict(size=10, color=GRAY_DARK),
        hovertemplate="%{y}: $%{x:.2f}M<extra></extra>",
    ))
    lo2 = base_layout("Revenue by BU · FYTD ($M)")
    lo2["xaxis"]["tickprefix"] = "$"
    lo2["xaxis"]["ticksuffix"] = "M"
    lo2["margin"]["r"] = 55
    lo2["showlegend"] = False
    fig2.update_layout(**lo2)
    st.plotly_chart(fig2, use_container_width=True, config={"displayModeBar": False})


# ── Row 2 · Margins ───────────────────────────────────────────────────────────
st.markdown('<div class="section-hdr">Margins</div>', unsafe_allow_html=True)
col_tbl, col_trend = st.columns([2, 3])

with col_tbl:
    bus_rows = [b for b in BU_ORDER if b in bus]
    gm_vals, cm_vals, eb_vals = [], [], []
    for b in bus_rows:
        rv = sum(m["by_bu"].get(b, {}).get("revenue", 0)             for m in FY2526[:n])
        gp = sum(m["by_bu"].get(b, {}).get("gross_profit", 0)        for m in FY2526[:n])
        cm = sum(m["by_bu"].get(b, {}).get("contribution_margin", 0) for m in FY2526[:n])
        eb = sum(m["by_bu"].get(b, {}).get("ebitda", 0)              for m in FY2526[:n])
        gm_vals.append(gp / rv * 100 if rv else 0)
        cm_vals.append(cm / rv * 100 if rv else 0)
        eb_vals.append(eb / rv * 100 if rv else 0)

    def heat(val, lo=-50, hi=70):
        t = max(0.0, min(1.0, (val - lo) / (hi - lo)))
        r = int(255 * (1 - t) + 23  * t)
        g = int(80  * (1 - t) + 228 * t)
        bv = int(80  * (1 - t) + 161 * t)
        return f"rgba({r},{g},{bv},0.55)"

    fig3 = go.Figure(go.Table(
        header=dict(
            values=["<b>Business Unit</b>", "<b>GM %</b>", "<b>CM %</b>", "<b>EBITDA %</b>"],
            fill_color=DARK_BLUE,
            font=dict(color="white", size=11, family="IBM Plex Sans"),
            align=["left", "center", "center", "center"],
            height=30,
        ),
        cells=dict(
            values=[
                bus_rows,
                [f"{v:.1f}%" for v in gm_vals],
                [f"{v:.1f}%" for v in cm_vals],
                [f"{v:.1f}%" for v in eb_vals],
            ],
            fill_color=[
                [CREAM_LIGHT] * len(bus_rows),
                [heat(v) for v in gm_vals],
                [heat(v) for v in cm_vals],
                [heat(v) for v in eb_vals],
            ],
            font=dict(color=DARK_BLUE, size=11, family="IBM Plex Sans"),
            align=["left", "center", "center", "center"],
            height=27,
        ),
    ))
    fig3.update_layout(
        height=260, margin=dict(l=0, r=0, t=0, b=0),
        paper_bgcolor="white",
    )
    st.plotly_chart(fig3, use_container_width=True, config={"displayModeBar": False})

with col_trend:
    labels_26n = [m["label"] for m in FY2526[:n]]
    gm_t, cm_t, eb_t = [], [], []
    for m in FY2526[:n]:
        rv = sum(m["by_bu"].get(b, {}).get("revenue", 0)             for b in bus)
        gp = sum(m["by_bu"].get(b, {}).get("gross_profit", 0)        for b in bus)
        cm = sum(m["by_bu"].get(b, {}).get("contribution_margin", 0) for b in bus)
        eb = sum(m["by_bu"].get(b, {}).get("ebitda", 0)              for b in bus)
        gm_t.append(gp / rv * 100 if rv else 0)
        cm_t.append(cm / rv * 100 if rv else 0)
        eb_t.append(eb / rv * 100 if rv else 0)

    fig4 = go.Figure()
    for vals, name, color in [(gm_t, "GM %", GREEN), (cm_t, "CM %", PURPLE), (eb_t, "EBITDA %", ORANGE)]:
        fig4.add_trace(go.Scatter(
            x=labels_26n, y=vals, name=name, mode="lines+markers",
            line=dict(color=color, width=2),
            marker=dict(size=5, color=color),
            hovertemplate=f"%{{x}}: %{{y:.1f}}%<extra>{name}</extra>",
        ))
    if labels_26n:
        fig4.add_shape(type="line",
            x0=labels_26n[0], x1=labels_26n[-1], y0=0, y1=0,
            line=dict(color=GRAY_LIGHT, width=1, dash="dot"),
        )
    lo4 = base_layout("Group Margin Trends · FY 2025/26", height=260)
    lo4["yaxis"]["ticksuffix"] = "%"
    fig4.update_layout(**lo4)
    st.plotly_chart(fig4, use_container_width=True, config={"displayModeBar": False})


# ── Row 3 · Cash ──────────────────────────────────────────────────────────────
st.markdown('<div class="section-hdr">Cash</div>', unsafe_allow_html=True)
col_bill, col_cash = st.columns([3, 2])

cash_labels    = [m["label"]     for m in cash_slice]
invoiced_vals  = [m["invoiced"]  / 1e6 for m in cash_slice]
collected_vals = [m["collected"] / 1e6 for m in cash_slice]
eom_vals       = [m["eom_cash"] / 1e6 for m in cash_slice]

with col_bill:
    fig5 = go.Figure()
    fig5.add_trace(go.Bar(name="Invoiced",  x=cash_labels, y=invoiced_vals,
                          marker_color=PURPLE, opacity=0.82,
                          hovertemplate="%{x}: $%{y:.2f}M<extra>Invoiced</extra>"))
    fig5.add_trace(go.Bar(name="Collected", x=cash_labels, y=collected_vals,
                          marker_color=GREEN, opacity=0.90,
                          hovertemplate="%{x}: $%{y:.2f}M<extra>Collected</extra>"))
    lo5 = base_layout("Invoiced vs Collected ($M)", height=240)
    lo5["barmode"] = "group"
    lo5["yaxis"]["tickprefix"] = "$"
    lo5["yaxis"]["ticksuffix"] = "M"
    fig5.update_layout(**lo5)
    if DATA.get("meta", {}).get("billing_note"):
        st.caption(DATA["meta"]["billing_note"])
    st.plotly_chart(fig5, use_container_width=True, config={"displayModeBar": False})

with col_cash:
    fig6 = go.Figure()
    fig6.add_trace(go.Scatter(
        x=cash_labels, y=eom_vals,
        name="EoM Cash", mode="lines+markers",
        line=dict(color="#4FC8FE", width=3),
        marker=dict(size=6, color="#4FC8FE"),
        fill="tozeroy", fillcolor="rgba(79,200,254,0.10)",
        hovertemplate="%{x}: $%{y:.2f}M<extra>EoM Cash</extra>",
    ))
    if runway and cash_labels:
        fig6.add_annotation(
            x=cash_labels[-1], y=eom_vals[-1],
            text=f"  {runway} mo runway",
            showarrow=False, xanchor="left",
            font=dict(color=ORANGE, size=10, family="IBM Plex Sans"),
        )
    lo6 = base_layout("Cash Position ($M)", height=240)
    lo6["yaxis"]["tickprefix"] = "$"
    lo6["yaxis"]["ticksuffix"] = "M"
    lo6["showlegend"] = False
    fig6.update_layout(**lo6)
    st.plotly_chart(fig6, use_container_width=True, config={"displayModeBar": False})


# ── Footer ────────────────────────────────────────────────────────────────────
st.markdown(f"""
<hr style="border:none; border-top:1px solid {GRAY_LIGHT}; margin:8px 0 4px 0;">
<div style="font-size:10px; color:{GRAY_MID}; text-align:center; padding-bottom:6px;">
  Noon Academy · Internal Management Dashboard · Confidential
  &nbsp;|&nbsp; Data as of {DATA['meta']['generated_at']}
</div>
""", unsafe_allow_html=True)
