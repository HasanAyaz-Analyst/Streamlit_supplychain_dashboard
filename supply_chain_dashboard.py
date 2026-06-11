import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np

st.set_page_config(page_title="Supply Chain Performance", layout="wide", page_icon="📦", initial_sidebar_state="expanded")

C = {
    "bg": "#f0eef8", "card": "#ffffff", "panel": "#ede9f6",
    "dark": "#1e1245", "mid": "#3d2e8c", "light": "#7c6bbf",
    "lighter": "#b3a8dc", "lightest": "#dbd6f2",
    "gold": "#e8b84b", "green": "#2db87a", "red": "#e05555",
    "text": "#2a1f5e", "subtext": "#7c6bbf",
}
CH_COLORS = [C["lighter"], C["light"], C["mid"], C["dark"]]

st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800&display=swap');
html, body, [class*="css"] {{ font-family: 'Inter', sans-serif; }}
.stApp {{ background: {C['bg']} !important; }}

section[data-testid="stSidebar"] > div:first-child {{
    background: {C['card']} !important; padding: 0 !important; width: 220px !important;
}}
section[data-testid="stSidebar"] {{ width: 220px !important; min-width: 220px !important; }}
[data-testid="stSidebarContent"] {{ padding: 16px 14px !important; }}

.stRadio > label {{ display: none; }}
.stRadio > div {{ flex-direction: column !important; gap: 4px !important; }}
.stRadio > div > label {{
    background: transparent; border-radius: 6px; padding: 7px 10px !important;
    font-size: 13px !important; font-weight: 600; color: {C['text']} !important; cursor: pointer;
}}
.stRadio > div > label:has(input:checked) {{
    background: {C['lightest']} !important; color: {C['dark']} !important;
}}

.kpi-box {{ padding: 10px 0; border-bottom: 1px solid {C['lightest']}; margin-bottom: 2px; }}
.kpi-val {{ font-size: 24px; font-weight: 800; color: {C['text']}; line-height: 1.1; }}
.kpi-lbl {{ font-size: 11px; color: #999; margin-top: 1px; }}
.kpi-sub {{ font-size: 11px; margin-top: 3px; font-weight: 500; }}

.sc-card {{
    background: {C['card']}; border-radius: 10px; padding: 14px 16px;
    margin-bottom: 10px; box-shadow: 0 1px 4px rgba(0,0,0,0.04);
}}
.sc-title {{ font-size: 14px; font-weight: 700; color: {C['text']}; margin-bottom: 1px; }}
.sc-sub {{ font-size: 11px; color: {C['subtext']}; font-style: italic; margin-bottom: 6px; line-height: 1.4; }}

.metric-card {{
    background: {C['card']}; border-radius: 10px; padding: 14px 16px;
    box-shadow: 0 1px 4px rgba(0,0,0,0.04);
}}
.metric-val {{ font-size: 26px; font-weight: 800; color: {C['text']}; }}
.metric-lbl {{ font-size: 12px; color: #999; margin-top: 2px; }}
.metric-sub {{ font-size: 11px; margin-top: 4px; }}

.block-container {{ padding-top: 0rem !important; padding-bottom: 1rem !important; max-width: 100% !important; }}

/* Hide Streamlit default top header/toolbar */
header[data-testid="stHeader"] {{ display: none !important; height: 0 !important; }}
div[data-testid="stToolbar"] {{ display: none !important; }}
#MainMenu {{ visibility: hidden !important; }}
footer {{ visibility: hidden !important; }}
.stDeployButton {{ display: none !important; }}

/* Push sidebar down less */
section[data-testid="stSidebar"] {{ top: 0 !important; padding-top: 0 !important; }}
div[data-testid="column"] {{ padding: 0 5px !important; }}

div[data-testid="stPlotlyChart"] > div {{ border-radius: 10px !important; overflow: hidden !important; }}

div[data-testid="stSelectbox"] > label {{ font-size: 11px; font-weight: 600; color: {C['subtext']}; margin-bottom: 2px; }}
div[data-testid="stSelectbox"] > div > div {{
    background: white !important; border: 1px solid {C['lighter']} !important;
    border-radius: 6px !important; font-size: 12px !important; min-height: 32px !important; padding: 2px 8px !important;
}}

button[kind="secondary"] {{
    background: {C['dark']} !important; color: white !important; border: none !important;
    border-radius: 6px !important; font-size: 12px !important; font-weight: 600 !important; padding: 6px 14px !important;
}}
</style>
""", unsafe_allow_html=True)

def PL(h=260, ml=55, mr=55, mt=15, mb=45):
    return dict(
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter,sans-serif", size=11, color=C["text"]),
        height=h, margin=dict(t=mt, r=mr, b=mb, l=ml),
    )

CFG = dict(displayModeBar=False, responsive=True)

@st.cache_data
def load():
    path = "Dataset_food_beverage_supply_chain_4u_reports_challenge.xlsx"
    orders = pd.read_excel(path, sheet_name="Fact_Orders")
    inv    = pd.read_excel(path, sheet_name="Inventory_Snapshots")
    prod   = pd.read_excel(path, sheet_name="Dim_Product")
    sup    = pd.read_excel(path, sheet_name="Dim_Supplier")
    cust   = pd.read_excel(path, sheet_name="Dim_Customer")
    wh     = pd.read_excel(path, sheet_name="Dim_Warehouse")
    orders["OrderDate"]    = pd.to_datetime(orders["OrderDate"])
    orders["Year"]         = orders["OrderDate"].dt.year
    orders["MonthName"]    = orders["OrderDate"].dt.strftime("%b")
    orders["MonthNo"]      = orders["OrderDate"].dt.month
    inv["SnapshotMonth"]   = pd.to_datetime(inv["SnapshotMonth"])
    inv["Year"]            = inv["SnapshotMonth"].dt.year
    inv["MonthName"]       = inv["SnapshotMonth"].dt.strftime("%b")
    inv["MonthNo"]         = inv["SnapshotMonth"].dt.month
    orders = orders.merge(sup[["SupplierID","SupplierName","RiskTier","SupplierRegion"]], on="SupplierID", how="left")
    orders = orders.merge(prod[["ProductID","Category","ProductName"]], on="ProductID", how="left")
    orders = orders.merge(cust[["CustomerID","SalesRegion","Country"]], on="CustomerID", how="left")
    orders = orders.merge(wh[["WarehouseID","WarehouseName"]], on="WarehouseID", how="left")
    inv    = inv.merge(wh[["WarehouseID","WarehouseName"]], on="WarehouseID", how="left")
    inv    = inv.merge(prod[["ProductID","Category","ProductName"]], on="ProductID", how="left")
    return orders, inv

orders_raw, inv_raw = load()

# ── Sidebar ────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown(f"""
    <div style="display:flex;align-items:center;gap:10px;margin-bottom:18px;padding-bottom:12px;border-bottom:1px solid {C['lightest']}">
      <div style="width:38px;height:38px;background:{C['mid']};border-radius:8px;display:flex;align-items:center;justify-content:center;flex-shrink:0">
        <svg width="22" height="22" viewBox="0 0 20 20" fill="none">
          <rect x="2" y="2" width="7" height="7" rx="1.5" fill="white" opacity="0.95"/>
          <rect x="11" y="2" width="7" height="7" rx="1.5" fill="white" opacity="0.65"/>
          <rect x="2" y="11" width="7" height="7" rx="1.5" fill="white" opacity="0.65"/>
          <rect x="11" y="11" width="7" height="7" rx="1.5" fill="white" opacity="0.35"/>
        </svg>
      </div>
      <div>
        <div style="font-size:13px;font-weight:800;color:{C['text']};line-height:1.2">SUPPLY CHAIN</div>
        <div style="font-size:9px;color:{C['subtext']};letter-spacing:2px;font-weight:600">PERFORMANCE</div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    tab = st.radio("nav", ["📊  Overview", "🚚  Fulfilment Risk", "📦  Inventory"], label_visibility="collapsed")

    st.markdown(f"<div style='margin:10px 0 6px;font-size:10px;font-weight:700;color:{C['subtext']};letter-spacing:1px'>KEY METRICS</div>", unsafe_allow_html=True)

    o    = orders_raw
    prev = o[o["Year"]==2024]
    curr = o[o["Year"]==2025]

    def yoy(a, b): return (b - a) / a * 100 if a else 0

    rev_tot   = o["Revenue"].sum()
    gm_pct    = o["GrossProfit"].sum() / o["Revenue"].sum() * 100
    otif_pct  = o["OTIF_Flag"].mean() * 100
    so_rate   = o["StockoutFlag"].mean() * 100
    waste_n   = int(o["WasteQty"].sum())
    rev_yoy   = yoy(prev["Revenue"].sum(), curr["Revenue"].sum())
    otif_yoy  = curr["OTIF_Flag"].mean()*100 - prev["OTIF_Flag"].mean()*100
    so_yoy    = yoy(prev["StockoutFlag"].mean()*100, curr["StockoutFlag"].mean()*100)
    waste_yoy = yoy(int(prev["WasteQty"].sum()), int(curr["WasteQty"].sum()))

    def kpi_card(val, lbl, sub_lbl, pct, good_pos=True):
        up   = pct > 0
        good = up == good_pos
        col  = C["green"] if good else C["red"]
        arr  = "▲" if up else "▼"
        st.markdown(f"""
        <div class="kpi-box">
          <div class="kpi-val">{val}</div>
          <div class="kpi-lbl">{lbl}</div>
          <div class="kpi-sub" style="color:{col}">{sub_lbl} &nbsp;<b>{arr} {abs(pct):.2f}%</b></div>
        </div>""", unsafe_allow_html=True)

    kpi_card(f"${rev_tot/1000:.0f}K",  "Revenue",       "Revenue YoY %",   rev_yoy,   True)
    kpi_card(f"{gm_pct:.2f}%",         "Gross Margin %","Gross Margin %",  gm_pct,    True)
    kpi_card(f"{otif_pct:.2f}%",       "OTIF %",        "OTIF YoY %",      otif_yoy,  True)
    kpi_card(f"{so_rate:.2f}%",        "Stockout Rate", "Stockout YoY %",  so_yoy,    False)
    kpi_card(f"{waste_n}",             "Waste Qty",     "Waste Qty YoY %", waste_yoy, False)

# ── Options helper ─────────────────────────────────────────────────────────
def opt(s): return ["All"] + sorted(s.dropna().unique().tolist())

# ── Page header ─────────────────────────────────────────────────────────────
TITLES = {
    "📊  Overview":        ("Performance Overview",    "Supply chain visibility at a glance."),
    "🚚  Fulfilment Risk": ("Fulfilment Risk",          "Where are orders failing and why?"),
    "📦  Inventory":       ("Inventory Health",         "Where are we losing product — and can we stop it?"),
}
ttl, sub = TITLES[tab]

hc1, hc2 = st.columns([6, 1])
with hc1:
    st.markdown(f"""
    <div style="padding:2px 0 8px">
      <div style="font-size:22px;font-weight:800;color:{C['text']};line-height:1.2">{ttl}</div>
      <div style="font-size:12px;color:{C['subtext']};font-style:italic;margin-top:2px">{sub}</div>
    </div>""", unsafe_allow_html=True)
with hc2:
    if st.button("↺  Reset", key="reset"):
        st.rerun()

# ── Filters ────────────────────────────────────────────────────────────────
if tab == "📊  Overview":
    fc1, fc2, fc3, fc4 = st.columns(4)
    f_risk  = fc1.selectbox("RiskTier", opt(orders_raw["RiskTier"]), key="ov_risk")
    f_ch    = fc2.selectbox("Channel",  opt(orders_raw["Channel"]),  key="ov_ch")
    f_cat   = fc3.selectbox("Category", opt(orders_raw["Category"]), key="ov_cat")
    f_year  = fc4.selectbox("Year",     ["2024","2025"],             key="ov_yr")
    f_yr_ff = "All"; f_yr_inv = "All"; f_sup = "All"; f_country = "All"

elif tab == "🚚  Fulfilment Risk":
    fc1, fc2, fc3 = st.columns(3)
    f_yr_ff = fc1.selectbox("Year",     ["All","2024","2025"],       key="ff_yr")
    f_ch    = fc2.selectbox("Channel",  opt(orders_raw["Channel"]),  key="ff_ch")
    f_cat   = fc3.selectbox("Category", opt(orders_raw["Category"]), key="ff_cat")
    f_risk = "All"; f_year = "2024"; f_yr_inv = "All"; f_sup = "All"; f_country = "All"

else:
    fc1, fc2, fc3, fc4, fc5, fc6 = st.columns(6)
    f_yr_inv  = fc1.selectbox("Year",     ["All","2024","2025"],             key="inv_yr")
    f_country = fc2.selectbox("Country",  opt(orders_raw["Country"]),        key="inv_co")
    f_sup     = fc3.selectbox("Supplier", opt(orders_raw["SupplierName"]),   key="inv_sup")
    f_risk    = fc4.selectbox("RiskTier", opt(orders_raw["RiskTier"]),       key="inv_risk")
    f_ch      = fc5.selectbox("Channel",  opt(orders_raw["Channel"]),        key="inv_ch")
    f_cat     = fc6.selectbox("Category", opt(orders_raw["Category"]),       key="inv_cat")
    f_year = "2024"; f_yr_ff = "All"

# ── Apply filters ──────────────────────────────────────────────────────────
def filt_orders(df):
    d = df.copy()
    if tab == "📊  Overview":
        d = d[d["Year"] == int(f_year)]
        if f_risk != "All":    d = d[d["RiskTier"] == f_risk]
        if f_ch   != "All":    d = d[d["Channel"]  == f_ch]
        if f_cat  != "All":    d = d[d["Category"] == f_cat]
    elif tab == "🚚  Fulfilment Risk":
        if f_yr_ff != "All":   d = d[d["Year"]     == int(f_yr_ff)]
        if f_ch    != "All":   d = d[d["Channel"]  == f_ch]
        if f_cat   != "All":   d = d[d["Category"] == f_cat]
    else:
        if f_yr_inv  != "All": d = d[d["Year"]        == int(f_yr_inv)]
        if f_ch      != "All": d = d[d["Channel"]     == f_ch]
        if f_cat     != "All": d = d[d["Category"]    == f_cat]
        if f_risk    != "All": d = d[d["RiskTier"]    == f_risk]
        if f_sup     != "All": d = d[d["SupplierName"]== f_sup]
        if f_country != "All": d = d[d["Country"]     == f_country]
    return d

def filt_inv(df):
    d = df.copy()
    if tab == "📦  Inventory":
        if f_yr_inv != "All": d = d[d["Year"]     == int(f_yr_inv)]
        if f_cat    != "All": d = d[d["Category"] == f_cat]
    return d

orders = filt_orders(orders_raw)
inv    = filt_inv(inv_raw)

# ── Helpers ────────────────────────────────────────────────────────────────
def card(title, subtitle=""):
    st.markdown(f"""
    <div class="sc-card">
      <div class="sc-title">{title}</div>
      {"<div class='sc-sub'>" + subtitle + "</div>" if subtitle else ""}
    </div>""", unsafe_allow_html=True)

def pc(fig):  # plotly chart wrapper with rounded corners
    st.markdown('<div style="border-radius:10px;overflow:hidden;">', unsafe_allow_html=True)
    st.plotly_chart(fig, use_container_width=True, config=CFG)
    st.markdown('</div>', unsafe_allow_html=True)

def stacked_bar(data_dict, is_rev=True):
    channels = ["Distributor", "E-commerce", "HoReCa", "Retail"]
    fig = go.Figure()
    for i, ch in enumerate(channels):
        v = data_dict.get(ch, 0)
        lbl = f"${int(v):,}" if is_rev else str(int(v))
        fig.add_trace(go.Bar(
            x=[v], y=[""], orientation="h", name=ch,
            marker_color=CH_COLORS[i],
            text=[lbl], textposition="inside", insidetextanchor="middle",
            textfont=dict(color="white" if i >= 2 else C["text"], size=12),
            hovertemplate=f"<b>{ch}</b><br>{'Revenue' if is_rev else 'Returns'}: %{{x:,}}<extra></extra>",
            width=0.7,
        ))
    fig.update_layout(
        **PL(h=95, ml=5, mr=10, mt=5, mb=40),
        barmode="stack",
        xaxis=dict(visible=False), yaxis=dict(visible=False),
        showlegend=True,
        legend=dict(orientation="h", y=-0.7, x=0.02, font=dict(size=11), traceorder="normal"),
    )
    pc(fig)

# ════════════════════════════════════════════════════════════════
# OVERVIEW
# ════════════════════════════════════════════════════════════════
if tab == "📊  Overview":

    card("Revenue By Channel")
    stacked_bar(orders.groupby("Channel")["Revenue"].sum().to_dict(), is_rev=True)

    col1, col2 = st.columns(2)

    with col1:
        card("Revenue by Supplier", "Who contributes the most?")
        sr = orders.groupby("SupplierName")["Revenue"].sum().sort_values(ascending=True).reset_index()
        fig = go.Figure(go.Bar(
            x=sr["Revenue"], y=sr["SupplierName"], orientation="h",
            marker_color=[C["dark"] if v == sr["Revenue"].max() else
                          C["mid"]  if v > sr["Revenue"].quantile(0.65) else C["light"]
                          for v in sr["Revenue"]],
            text=[f"${v/1000:.0f}K" for v in sr["Revenue"]], textposition="outside",
            hovertemplate="<b>%{y}</b><br>Revenue: $%{x:,.0f}<extra></extra>",
        ))
        fig.update_layout(**PL(h=290, ml=160, mr=75, mt=10, mb=30),
            xaxis=dict(tickprefix="$", tickformat=",.0f", gridcolor=C["lightest"]),
            yaxis=dict(tickfont=dict(size=11)))
        pc(fig)

    with col2:
        card("Revenue and OTIF % by Month", "Is our delivery reliability keeping up with demand?")
        rm = orders.groupby(["MonthName","MonthNo"]).agg(
            Revenue=("Revenue","sum"), OTIF=("OTIF_Flag","mean")
        ).reset_index().sort_values("MonthNo")
        rm["OTIF"] *= 100
        fig = make_subplots(specs=[[{"secondary_y": True}]])
        fig.add_trace(go.Bar(x=rm["MonthName"], y=rm["Revenue"], name="Revenue",
            marker_color=C["mid"],
            hovertemplate="<b>%{x}</b><br>Revenue: $%{y:,.0f}<extra></extra>"), secondary_y=False)
        fig.add_trace(go.Scatter(x=rm["MonthName"], y=rm["OTIF"], name="OTIF %",
            mode="lines+markers",
            line=dict(color=C["gold"], width=2.5, shape="spline", smoothing=1.3),
            marker=dict(color=C["gold"], size=6),
            hovertemplate="<b>%{x}</b><br>OTIF: %{y:.1f}%<extra></extra>"), secondary_y=True)
        fig.update_layout(**PL(h=290, ml=55, mr=60, mt=10, mb=50),
            bargap=0.3, showlegend=True,
            legend=dict(orientation="h", y=-0.2, x=0, font=dict(size=11)))
        fig.update_yaxes(tickprefix="$", tickformat=",.0f", gridcolor=C["lightest"], secondary_y=False)
        fig.update_yaxes(ticksuffix="%", showgrid=False, tickfont=dict(color=C["gold"]), secondary_y=True)
        pc(fig)

    col3, col4 = st.columns(2)

    with col3:
        card("Revenue by Region and Country", "Where is our growth coming from?")
        rr = orders.groupby("SalesRegion")["Revenue"].sum().sort_values(ascending=False).reset_index()
        fig = go.Figure(go.Bar(
            x=rr["SalesRegion"], y=rr["Revenue"],
            marker_color=[C["dark"],C["mid"],C["light"],C["lighter"],C["lightest"]][:len(rr)],
            text=[f"${v/1000:.0f}K" for v in rr["Revenue"]], textposition="outside",
            hovertemplate="<b>%{x}</b><br>Revenue: $%{y:,.0f}<extra></extra>",
        ))
        fig.update_layout(**PL(h=260, ml=60, mr=20, mt=10, mb=45),
            yaxis=dict(tickprefix="$", tickformat=",.0f", gridcolor=C["lightest"]),
            xaxis=dict(tickfont=dict(size=11)))
        pc(fig)

    with col4:
        card("Gross Margin % by Category", "Where are margins strongest?")
        cg = orders.groupby("Category").agg(Rev=("Revenue","sum"), GP=("GrossProfit","sum")).reset_index()
        cg["GM"] = cg["GP"] / cg["Rev"] * 100
        fig = go.Figure(go.Bar(
            x=cg["Category"], y=cg["GM"],
            marker_color=[C["dark"],C["mid"],C["light"],C["lighter"],C["lightest"]][:len(cg)],
            text=[f"{v:.1f}%" for v in cg["GM"]], textposition="outside",
            hovertemplate="<b>%{x}</b><br>GM%: %{y:.2f}%<extra></extra>",
        ))
        fig.update_layout(**PL(h=260, ml=60, mr=20, mt=10, mb=45),
            yaxis=dict(ticksuffix="%", gridcolor=C["lightest"]),
            xaxis=dict(tickfont=dict(size=11)))
        pc(fig)

# ════════════════════════════════════════════════════════════════
# FULFILMENT RISK
# ════════════════════════════════════════════════════════════════
elif tab == "🚚  Fulfilment Risk":

    worst_sup     = orders.groupby("SupplierName")["OTIF_Flag"].mean().idxmin()
    worst_val     = orders.groupby("SupplierName")["OTIF_Flag"].mean().min() * 100
    qual_n        = int(orders["QualityIssueFlag"].sum())
    qual_sup      = orders.groupby("SupplierName")["QualityIssueFlag"].sum().idxmax()
    qual_rate     = orders.groupby("SupplierName")["QualityIssueFlag"].mean().max() * 100
    so_r          = orders["StockoutFlag"].mean() * 100
    so_sup        = orders.groupby("SupplierName")["StockoutFlag"].sum().idxmax()
    so_sup_r      = orders.groupby("SupplierName")["StockoutFlag"].mean().max() * 100

    m1, m2, m3, m4 = st.columns(4)
    for col, v, l, s, sc in [
        (m1, f"{orders['OTIF_Flag'].mean()*100:.2f}%", "OTIF %",
         f"Worst: {worst_sup} ({worst_val:.2f}%)", C["red"]),
        (m2, str(qual_n), "Quality Issue Orders",
         f"Most: {qual_sup} ({qual_rate:.2f}%)", C["red"]),
        (m3, f"{so_r:.2f}%", "Stockout Rate",
         f"Highest: {so_sup} ({so_sup_r:.2f}%)", C["red"]),
        (m4, str(int(orders["ReturnQty"].sum())), "Total Returns", "By Channel →", C["subtext"]),
    ]:
        col.markdown(f"""
        <div class="metric-card">
          <div class="metric-val">{v}</div>
          <div class="metric-lbl">{l}</div>
          <div class="metric-sub" style="color:{sc}">{s}</div>
        </div>""", unsafe_allow_html=True)

    st.markdown("<div style='margin-top:10px'></div>", unsafe_allow_html=True)

    card("Return By Channel")
    stacked_bar(orders.groupby("Channel")["ReturnQty"].sum().to_dict(), is_rev=False)

    col1, col2 = st.columns(2)

    with col1:
        card("Stockout Orders and Late Orders Rate by Channel",
             "Which channel struggles most with availability and on-time delivery?")
        so_ch = orders.groupby("Channel").agg(
            Stockout=("StockoutFlag","sum"),
            Late=("OTIF_Flag", lambda x: (1 - x.mean()) * 100)
        ).reindex(["Distributor","E-commerce","HoReCa","Retail"], fill_value=0).reset_index()
        fig = make_subplots(specs=[[{"secondary_y": True}]])
        fig.add_trace(go.Bar(
            y=so_ch["Channel"], x=so_ch["Stockout"], orientation="h",
            name="Stockout Orders",
            marker_color=[C["lighter"] if ch != "Retail" else C["dark"] for ch in so_ch["Channel"]],
            text=so_ch["Stockout"].astype(int), textposition="inside",
            insidetextanchor="middle", textfont=dict(color="white", size=12),
            hovertemplate="<b>%{y}</b><br>Stockout Orders: %{x}<extra></extra>",
            width=0.5,
        ), secondary_y=False)
        fig.add_trace(go.Scatter(
            y=so_ch["Channel"], x=so_ch["Late"],
            mode="markers+text", name="Late Rate %",
            marker=dict(color=C["gold"], size=14),
            text=[f"{v:.2f}%" for v in so_ch["Late"]],
            textposition="middle right", textfont=dict(color=C["text"], size=11),
            hovertemplate="<b>%{y}</b><br>Late Rate: %{x:.2f}%<extra></extra>",
        ), secondary_y=True)
        fig.update_layout(**PL(h=300, ml=95, mr=110, mt=10, mb=60),
            yaxis=dict(autorange="reversed", tickfont=dict(size=12)),
            xaxis=dict(gridcolor=C["lightest"]),
            xaxis2=dict(visible=False),
            showlegend=True, legend=dict(orientation="h", y=-0.22, x=0, font=dict(size=11)))
        pc(fig)

    with col2:
        card("Stockout Orders and Late Orders Rate by Category",
             "Which categories are creating the most fulfilment risk?")
        so_cat = orders.groupby("Category").agg(
            Stockout=("StockoutFlag","sum"),
            Late=("OTIF_Flag", lambda x: (1 - x.mean()) * 100)
        ).reset_index()
        fig = make_subplots(specs=[[{"secondary_y": True}]])
        fig.add_trace(go.Bar(x=so_cat["Category"], y=so_cat["Stockout"], name="Stockout Orders",
            marker_color=C["dark"],
            hovertemplate="<b>%{x}</b><br>Stockout: %{y}<extra></extra>"), secondary_y=False)
        fig.add_trace(go.Scatter(x=so_cat["Category"], y=so_cat["Late"], name="Late Rate %",
            mode="lines+markers",
            line=dict(color=C["gold"], width=2.5, shape="spline", smoothing=1.3),
            marker=dict(color=C["gold"], size=7),
            hovertemplate="<b>%{x}</b><br>Late Rate: %{y:.2f}%<extra></extra>"), secondary_y=True)
        fig.update_layout(**PL(h=300, ml=45, mr=60, mt=10, mb=55),
            bargap=0.35,
            yaxis=dict(gridcolor=C["lightest"]),
            yaxis2=dict(ticksuffix="%", showgrid=False, tickfont=dict(color=C["gold"])),
            showlegend=True, legend=dict(orientation="h", y=-0.22, x=0, font=dict(size=11)))
        pc(fig)

    card("OTIF %, Gross Margin % and Revenue by Supplier and Risk Tier",
         "Are our high-risk suppliers costing us on delivery — and are they worth it?")
    sa = orders.groupby(["SupplierName","RiskTier"]).agg(
        OTIF=("OTIF_Flag","mean"), GP=("GrossProfit","sum"), Rev=("Revenue","sum")
    ).reset_index()
    sa["OTIF"] *= 100
    sa["GM"]    = sa["GP"] / sa["Rev"] * 100
    rc = {"Low": C["lighter"], "Medium": C["light"], "High": C["dark"]}
    fig = go.Figure()
    for risk in ["Low","Medium","High"]:
        d = sa[sa["RiskTier"] == risk]
        fig.add_trace(go.Scatter(
            x=d["GM"], y=d["OTIF"], mode="markers", name=f"{risk} Risk",
            text=d["SupplierName"],
            marker=dict(size=np.sqrt(d["Rev"]/1000)*2.2, color=rc.get(risk, C["light"]), opacity=0.85),
            hovertemplate="<b>%{text}</b><br>GM%: %{x:.1f}%<br>OTIF: %{y:.1f}%<extra></extra>",
        ))
    fig.update_layout(**PL(h=280, ml=65, mr=130, mt=10, mb=55),
        showlegend=True,
        xaxis=dict(title="Gross Margin %", ticksuffix="%", gridcolor=C["lightest"]),
        yaxis=dict(title="OTIF %", ticksuffix="%", gridcolor=C["lightest"]),
        legend=dict(font=dict(size=12)))
    pc(fig)

# ════════════════════════════════════════════════════════════════
# INVENTORY
# ════════════════════════════════════════════════════════════════
elif tab == "📦  Inventory":

    exp_qty  = int(inv["ExpiredQty"].sum())
    top_wh   = inv.groupby("WarehouseName")["ExpiredQty"].sum().idxmax() if exp_qty > 0 else "N/A"
    below_rp = int((inv_raw.groupby("ProductID").apply(
        lambda x: (x["ClosingStock"] < x["ReorderPoint"]).any())).sum())
    so_inv_n = int(inv["StockoutFlag"].sum())
    so_inv_r = inv["StockoutFlag"].mean() * 100

    m1, m2, m3, m4 = st.columns(4)
    for col, v, l, s, sc in [
        (m1, f"{exp_qty/1000:.1f}K",          "Expired Qty",               f"Top Warehouse: {top_wh}", C["red"]),
        (m2, str(below_rp),                    "Products Below Reorder Pt.","Monitor closely",          C["red"]),
        (m3, f"{so_inv_n/1000:.1f}K",          "Stockout Orders INV",       f"Rate: {so_inv_r:.2f}%",   C["red"]),
        (m4, str(int(orders["WasteQty"].sum())),"Total Waste Qty",           "All suppliers",            C["subtext"]),
    ]:
        col.markdown(f"""
        <div class="metric-card">
          <div class="metric-val">{v}</div>
          <div class="metric-lbl">{l}</div>
          <div class="metric-sub" style="color:{sc}">{s}</div>
        </div>""", unsafe_allow_html=True)

    st.markdown("<div style='margin-top:10px'></div>", unsafe_allow_html=True)

    col1, col2 = st.columns(2)

    with col1:
        card("Waste Qty by Supplier", "Which supplier is responsible for the most wasted product?")
        ws = orders.groupby("SupplierName")["WasteQty"].sum().sort_values(ascending=True).reset_index()
        fig = go.Figure(go.Bar(
            x=ws["WasteQty"], y=ws["SupplierName"], orientation="h",
            marker_color=[C["dark"] if v == ws["WasteQty"].max() else
                          C["mid"]  if v > ws["WasteQty"].quantile(0.65) else C["light"]
                          for v in ws["WasteQty"]],
            text=ws["WasteQty"].astype(int), textposition="outside",
            hovertemplate="<b>%{y}</b><br>Waste Qty: %{x:,}<extra></extra>",
        ))
        fig.update_layout(**PL(h=280, ml=165, mr=65, mt=10, mb=30),
            xaxis=dict(gridcolor=C["lightest"]),
            yaxis=dict(tickfont=dict(size=11)))
        pc(fig)

    with col2:
        lbl = f_yr_inv if f_yr_inv != "All" else "All Years"
        card(f"Expired Qty and Stockout Orders INV by Month ({lbl})",
             "Are expiry spikes and stockouts happening at the same time?")
        im = inv.groupby(["MonthName","MonthNo"]).agg(
            Expired=("ExpiredQty","sum"), SO=("StockoutFlag","sum")
        ).reset_index().sort_values("MonthNo")
        fig = make_subplots(specs=[[{"secondary_y": True}]])
        fig.add_trace(go.Bar(x=im["MonthName"], y=im["Expired"], name="Expired Qty",
            marker_color=C["dark"],
            hovertemplate="<b>%{x}</b><br>Expired Qty: %{y:,}<extra></extra>"), secondary_y=False)
        fig.add_trace(go.Scatter(x=im["MonthName"], y=im["SO"], name="Stockout INV",
            mode="lines+markers",
            line=dict(color=C["gold"], width=2.5, shape="spline", smoothing=1.3),
            marker=dict(color=C["gold"], size=6),
            hovertemplate="<b>%{x}</b><br>Stockout Orders: %{y:,}<extra></extra>"), secondary_y=True)
        fig.update_layout(**PL(h=280, ml=55, mr=60, mt=10, mb=50),
            bargap=0.3,
            yaxis=dict(gridcolor=C["lightest"]),
            yaxis2=dict(showgrid=False, tickfont=dict(color=C["gold"])),
            showlegend=True, legend=dict(orientation="h", y=-0.22, x=0, font=dict(size=11)))
        pc(fig)

    col3, col4 = st.columns(2)

    with col3:
        card("Waste Qty by Category and Product",
             "Is our waste concentrated in short shelf-life categories?")
        wc = orders.groupby("Category")["WasteQty"].sum().reset_index()
        fig = go.Figure(go.Pie(
            labels=wc["Category"], values=wc["WasteQty"], hole=0.5,
            marker=dict(colors=[C["dark"],C["mid"],C["light"],C["lighter"],C["lightest"]]),
            textinfo="label+percent", textfont=dict(size=12),
            hovertemplate="<b>%{label}</b><br>Waste Qty: %{value:,}<br>%{percent}<extra></extra>",
        ))
        fig.update_layout(**PL(h=280, ml=20, mr=20, mt=10, mb=20))
        pc(fig)

    with col4:
        card("Expired Qty and Stockout Orders INV by Warehouse",
             "Which warehouses have the worst expiry and availability problem?")
        iw = inv.groupby("WarehouseName").agg(
            Expired=("ExpiredQty","sum"), SO=("StockoutFlag","sum")
        ).reset_index().sort_values("Expired", ascending=False)
        fig = make_subplots(specs=[[{"secondary_y": True}]])
        fig.add_trace(go.Bar(x=iw["WarehouseName"], y=iw["Expired"], name="Expired Qty",
            marker_color=C["dark"],
            hovertemplate="<b>%{x}</b><br>Expired Qty: %{y:,}<extra></extra>"), secondary_y=False)
        fig.add_trace(go.Scatter(x=iw["WarehouseName"], y=iw["SO"], name="Stockout INV",
            mode="lines+markers+text",
            line=dict(color=C["gold"], width=2.5, shape="spline", smoothing=1.3),
            marker=dict(color=C["gold"], size=8),
            text=iw["SO"], textposition="top center", textfont=dict(size=11),
            hovertemplate="<b>%{x}</b><br>Stockout Orders: %{y:,}<extra></extra>"), secondary_y=True)
        fig.update_layout(**PL(h=280, ml=55, mr=60, mt=20, mb=55),
            bargap=0.3,
            xaxis=dict(tickfont=dict(size=10)),
            yaxis=dict(gridcolor=C["lightest"]),
            yaxis2=dict(showgrid=False, tickfont=dict(color=C["gold"])),
            showlegend=True, legend=dict(orientation="h", y=-0.22, x=0, font=dict(size=11)))
        pc(fig)

st.markdown(f"""
<div style='text-align:center;color:{C["subtext"]};font-size:10px;padding:16px 0 4px'>
  Supply Chain Performance Dashboard &nbsp;•&nbsp; 2024–2025
</div>""", unsafe_allow_html=True)