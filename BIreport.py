import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
import json

st.set_page_config(layout="wide", page_title="Sales Dashboard", page_icon="=")

# ---- SHARED CSS (identical to KPI report) ----
st.markdown(
    """
    <style>
    [data-testid="stToolbar"] { display: none !important; }
    [data-testid="stDecoration"] { display: none !important; }
    header[data-testid="stHeader"] { display: none !important; }

    html, body, [class*="css"], .stApp {
        background-color: #111827 !important;
        color: #e2e8f0 !important;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif !important;
    }
    [data-testid="stSidebar"] {
        background: #0f172a !important;
        border-right: 1px solid #1e293b !important;
    }
    [data-testid="stSidebar"] * { color: #cbd5e1 !important; }
    [data-testid="stSidebar"] .block-container { padding: 1.25rem 1rem !important; max-width: 100% !important; }
    .block-container { max-width: 1400px !important; padding: 1.5rem 2rem 3rem !important; margin: 0 auto !important; }
    h1 { font-size: 2.4rem !important; font-weight: 700 !important; color: #f1f5f9 !important; letter-spacing: -.02em !important; margin-bottom: .15rem !important; }
    .sidebar-section { font-size: .7rem; font-weight: 600; text-transform: uppercase; letter-spacing: .1em; color: #64748b; margin: 1.2rem 0 .4rem; display: block; }
    [data-testid="stFileUploader"] { border: 1px dashed #334155 !important; border-radius: 8px !important; background: transparent !important; }
    [data-testid="stPlotlyChart"] { overflow: visible !important; }
    .stTextInput input { background: #1e293b !important; border: 1px solid #334155 !important; border-radius: 6px !important; color: #e2e8f0 !important; font-size: .85rem !important; }

    .kpi-card { background: #1e293b; border: 1px solid #334155; border-radius: 8px; padding: 1rem 1.25rem; }
    .kpi-card-label { font-size: .65rem; text-transform: uppercase; letter-spacing: .1em; color: #64748b; margin-bottom: .3rem; }
    .kpi-card-value { font-size: 1.8rem; font-weight: 700; line-height: 1; color: #f1f5f9; }
    .kpi-card-sub { font-size: .7rem; color: #475569; margin-top: .25rem; }
    .kpi-card-delta-pos { color: #4ade80; font-size: .72rem; margin-top: .2rem; }
    .kpi-card-delta-neg { color: #f87171; font-size: .72rem; margin-top: .2rem; }

    .section-label { font-size: .68rem; font-weight: 600; text-transform: uppercase; letter-spacing: .1em; color: #475569; margin: 2rem 0 .75rem; border-bottom: 1px solid #1e293b; padding-bottom: .4rem; }

    [data-testid="stDataFrame"] td, [data-testid="stDataFrame"] th { text-align: left !important; }
    [data-testid="stDataFrame"] [class*="cell"] { text-align: left !important; justify-content: flex-start !important; }

    /* Chat UI */
    .chat-container { background: #0f172a; border: 1px solid #1e293b; border-radius: 10px; padding: 1rem 1.25rem; margin-bottom: 1rem; max-height: 420px; overflow-y: auto; }
    .chat-msg-user { display: flex; justify-content: flex-end; margin-bottom: .75rem; }
    .chat-msg-ai { display: flex; justify-content: flex-start; margin-bottom: .75rem; }
    .chat-bubble-user { background: #1d4ed8; color: #fff; border-radius: 14px 14px 2px 14px; padding: .55rem .9rem; font-size: .82rem; max-width: 75%; line-height: 1.5; }
    .chat-bubble-ai { background: #1e293b; color: #e2e8f0; border-radius: 14px 14px 14px 2px; padding: .55rem .9rem; font-size: .82rem; max-width: 85%; line-height: 1.5; border: 1px solid #334155; }
    .chat-label { font-size: .62rem; color: #475569; margin-bottom: .2rem; text-transform: uppercase; letter-spacing: .08em; }
    .chat-thinking { color: #475569; font-style: italic; font-size: .78rem; }
    </style>
    """,
    unsafe_allow_html=True,
)

# ---- HELPERS ----
def fmt_currency(v):
    if pd.isna(v): return "$0"
    if abs(v) >= 1_000_000: return f"${v/1_000_000:.2f}M"
    if abs(v) >= 1_000: return f"${v/1_000:.1f}K"
    return f"${v:.2f}"

def fmt_units(v):
    if pd.isna(v): return "0"
    if abs(v) >= 1_000_000: return f"{v/1_000_000:.2f}M"
    if abs(v) >= 1_000: return f"{v/1_000:.1f}K"
    return f"{int(v):,}"

def delta_html(curr, prev, fmt_fn, label="vs prev period"):
    if prev == 0 or pd.isna(prev): return ""
    pct = (curr - prev) / abs(prev) * 100
    arrow = "+" if pct >= 0 else ""
    cls = "kpi-card-delta-pos" if pct >= 0 else "kpi-card-delta-neg"
    return f'<div class="{cls}">{arrow}{pct:.1f}% {label}</div>'

PLOT_LAYOUT = dict(
    paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(15,23,42,.6)',
    template='plotly_dark', margin=dict(l=0, r=0, t=30, b=0),
    xaxis=dict(gridcolor='rgba(255,255,255,.05)', linecolor='rgba(255,255,255,.08)', tickfont=dict(size=11, color='#64748b')),
    yaxis=dict(gridcolor='rgba(255,255,255,.05)', linecolor='rgba(255,255,255,.08)', tickfont=dict(size=11, color='#64748b')),
    legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1, font=dict(size=11, color='#94a3b8'), bgcolor='rgba(0,0,0,0)'),
    hovermode='x unified',
    hoverlabel=dict(bgcolor='#1e293b', bordercolor='#334155', font=dict(size=12, color='#e2e8f0')),
)
PAL = ['#60a5fa', '#34d399', '#a78bfa', '#fb923c', '#f472b6', '#facc15', '#38bdf8']

# ---- SIDEBAR ----
with st.sidebar:
    st.markdown("<div style='font-size:1rem;font-weight:700;color:#f1f5f9;margin-bottom:.25rem;'>Sales Dashboard</div>", unsafe_allow_html=True)
    st.markdown("<div style='font-size:.65rem;color:#475569;text-transform:uppercase;letter-spacing:.1em;margin-bottom:1rem;'>Revenue & Units Analytics</div>", unsafe_allow_html=True)
    st.markdown('<span class="sidebar-section">Data Source</span>', unsafe_allow_html=True)
    uploaded_file = st.file_uploader("", type=["csv", "xlsx", "gz"], label_visibility="collapsed")
    if uploaded_file is None:
        st.info("Upload a CSV, XLSX, or CSV.GZ to begin.")
        st.stop()

# ---- LOAD DATA ----
try:
    name = uploaded_file.name.lower()
    if name.endswith(".csv.gz") or name.endswith(".gz"):
        raw = pd.read_csv(uploaded_file, compression="gzip")
    elif name.endswith(".csv"):
        raw = pd.read_csv(uploaded_file)
    else:
        raw = pd.read_excel(uploaded_file)
except Exception as e:
    st.sidebar.error(f"Error reading file: {e}")
    st.stop()

if raw is None or raw.empty:
    st.warning("No data found.")
    st.stop()

df = raw.copy()
df['date'] = pd.to_datetime(df['date'], errors='coerce')
df['week'] = df['date'].dt.to_period('W').dt.start_time
df['month'] = df['date'].dt.to_period('M').dt.start_time
df['sale_type'] = df['bundle_name'].fillna('Platform Sale').replace('', 'Platform Sale')

# ---- SIDEBAR FILTERS ----
with st.sidebar:
    if 'franchise' in df.columns:
        opts = sorted(df['franchise'].dropna().unique())
        st.markdown('<span class="sidebar-section">Franchise</span>', unsafe_allow_html=True)
        sel = st.multiselect("", options=opts, default=opts, label_visibility="collapsed")
        df = df[df['franchise'].isin(sel)]

    if 'game' in df.columns:
        opts = sorted(df['game'].dropna().unique())
        st.markdown('<span class="sidebar-section">Game</span>', unsafe_allow_html=True)
        sel = st.multiselect("", options=opts, default=opts, label_visibility="collapsed")
        df = df[df['game'].isin(sel)]

    if 'platform' in df.columns:
        opts = sorted(df['platform'].dropna().unique())
        st.markdown('<span class="sidebar-section">Platform</span>', unsafe_allow_html=True)
        sel = st.multiselect("", options=opts, default=opts, label_visibility="collapsed")
        df = df[df['platform'].isin(sel)]

    if 'bp_region' in df.columns:
        opts = sorted(df['bp_region'].dropna().unique())
        st.markdown('<span class="sidebar-section">Region</span>', unsafe_allow_html=True)
        sel = st.multiselect("", options=opts, default=opts, label_visibility="collapsed")
        df = df[df['bp_region'].isin(sel)]

    if 'product_type' in df.columns:
        opts = sorted(df['product_type'].dropna().unique())
        st.markdown('<span class="sidebar-section">Product Type</span>', unsafe_allow_html=True)
        sel = st.multiselect("", options=opts, default=opts, label_visibility="collapsed")
        df = df[df['product_type'].isin(sel)]

    if 'date' in df.columns and df['date'].notna().any():
        min_d, max_d = df['date'].min().date(), df['date'].max().date()
        st.markdown('<span class="sidebar-section">Date Range</span>', unsafe_allow_html=True)
        date_range = st.date_input("", value=(min_d, max_d), min_value=min_d, max_value=max_d, label_visibility="collapsed")
        if isinstance(date_range, (list, tuple)) and len(date_range) == 2:
            df = df[(df['date'].dt.date >= date_range[0]) & (df['date'].dt.date <= date_range[1])]

# ---- PAGE HEADER ----
st.markdown(
    f"<h1>Sales Dashboard</h1>"
    f"<div style='font-size:.78rem;color:#475569;margin-bottom:1.5rem;'>"
    f"{len(df):,} transactions &nbsp;&middot;&nbsp; "
    f"{df['game'].nunique() if 'game' in df.columns else 0} games &nbsp;&middot;&nbsp; "
    f"{df['date'].min().strftime('%b %d, %Y') if df['date'].notna().any() else ''} &ndash; "
    f"{df['date'].max().strftime('%b %d, %Y') if df['date'].notna().any() else ''}"
    f"</div>",
    unsafe_allow_html=True,
)

# ---- COMPUTE WEEKLY WoW FOR DELTA ----
weeks = sorted(df['week'].dropna().unique())
curr_rev = prev_rev = curr_units = prev_units = 0
if len(weeks) >= 2:
    curr_w, prev_w = weeks[-1], weeks[-2]
    curr_rev = df[df['week'] == curr_w]['revenue'].sum()
    prev_rev = df[df['week'] == prev_w]['revenue'].sum()
    curr_units = df[df['week'] == curr_w]['quantity'].sum()
    prev_units = df[df['week'] == prev_w]['quantity'].sum()

total_rev = df['revenue'].sum()
total_units = df['quantity'].sum()
total_net = df['net_revenue_usd'].sum() if 'net_revenue_usd' in df.columns else 0
direct_rev = df[df['sale_type'] == 'Direct Package Sale']['revenue'].sum()
platform_rev = df[df['sale_type'] == 'Platform Sale']['revenue'].sum()
direct_pct = (direct_rev / total_rev * 100) if total_rev > 0 else 0

# ---- SUMMARY CARDS ----
st.markdown('<div class="section-label">Summary</div>', unsafe_allow_html=True)
c1, c2, c3, c4, c5 = st.columns(5)

with c1:
    st.markdown(
        f'<div class="kpi-card"><div class="kpi-card-label">Total Revenue</div>'
        f'<div class="kpi-card-value">{fmt_currency(total_rev)}</div>'
        f'{delta_html(curr_rev, prev_rev, fmt_currency, "WoW")}'
        f'<div class="kpi-card-sub">gross USD</div></div>', unsafe_allow_html=True)

with c2:
    st.markdown(
        f'<div class="kpi-card"><div class="kpi-card-label">Total Units</div>'
        f'<div class="kpi-card-value">{fmt_units(total_units)}</div>'
        f'{delta_html(curr_units, prev_units, fmt_units, "WoW")}'
        f'<div class="kpi-card-sub">quantity sold</div></div>', unsafe_allow_html=True)

with c3:
    st.markdown(
        f'<div class="kpi-card"><div class="kpi-card-label">Net Revenue</div>'
        f'<div class="kpi-card-value">{fmt_currency(total_net)}</div>'
        f'<div class="kpi-card-sub">after royalties</div></div>', unsafe_allow_html=True)

with c4:
    avg_price = total_rev / total_units if total_units > 0 else 0
    st.markdown(
        f'<div class="kpi-card"><div class="kpi-card-label">Avg Sale Price</div>'
        f'<div class="kpi-card-value">{fmt_currency(avg_price)}</div>'
        f'<div class="kpi-card-sub">revenue / unit</div></div>', unsafe_allow_html=True)

with c5:
    st.markdown(
        f'<div class="kpi-card"><div class="kpi-card-label">Direct Sale %</div>'
        f'<div class="kpi-card-value">{direct_pct:.1f}%</div>'
        f'<div class="kpi-card-sub">vs platform sales</div></div>', unsafe_allow_html=True)

# ---- CHARTS ROW 1: Revenue over time + Platform split ----
st.markdown('<div class="section-label">Revenue Trends</div>', unsafe_allow_html=True)
ch1, ch2 = st.columns([3, 2])

with ch1:
    weekly = df.groupby(['week', 'game'])['revenue'].sum().reset_index()
    games = sorted(weekly['game'].unique())
    fig = go.Figure()
    for i, g in enumerate(games):
        gd = weekly[weekly['game'] == g].sort_values('week')
        fig.add_trace(go.Scatter(
            x=gd['week'], y=gd['revenue'], mode='lines+markers', name=g,
            line=dict(color=PAL[i % len(PAL)], width=2),
            marker=dict(size=5, line=dict(width=1.5, color='#111827')),
        ))
    fig.update_layout(height=320, title=dict(text="Weekly Revenue by Game", font=dict(size=13, color='#94a3b8')), **PLOT_LAYOUT)
    fig.update_layout(yaxis=dict(tickprefix='$', **PLOT_LAYOUT['yaxis']))
    st.plotly_chart(fig, use_container_width=True)

with ch2:
    plat = df.groupby('platform')[['revenue', 'quantity']].sum().reset_index().sort_values('revenue', ascending=False)
    fig2 = go.Figure(go.Bar(
        x=plat['platform'], y=plat['revenue'],
        marker_color=PAL[:len(plat)],
        text=[fmt_currency(v) for v in plat['revenue']],
        textposition='outside', textfont=dict(size=11, color='#94a3b8'),
    ))
    fig2.update_layout(height=320, title=dict(text="Revenue by Platform", font=dict(size=13, color='#94a3b8')), showlegend=False, **PLOT_LAYOUT)
    fig2.update_layout(yaxis=dict(tickprefix='$', **PLOT_LAYOUT['yaxis']))
    st.plotly_chart(fig2, use_container_width=True)

# ---- CHARTS ROW 2: Region + Direct vs Bundle ----
ch3, ch4 = st.columns([2, 2])

with ch3:
    region = df.groupby('bp_region')['revenue'].sum().reset_index().sort_values('revenue', ascending=False)
    fig3 = go.Figure(go.Bar(
        x=region['bp_region'], y=region['revenue'],
        marker_color=PAL[2:2+len(region)],
        text=[fmt_currency(v) for v in region['revenue']],
        textposition='outside', textfont=dict(size=11, color='#94a3b8'),
    ))
    fig3.update_layout(height=300, title=dict(text="Revenue by Region", font=dict(size=13, color='#94a3b8')), showlegend=False, **PLOT_LAYOUT)
    fig3.update_layout(yaxis=dict(tickprefix='$', **PLOT_LAYOUT['yaxis']))
    st.plotly_chart(fig3, use_container_width=True)

with ch4:
    sale_split = df.groupby('sale_type')['revenue'].sum().reset_index()
    fig4 = go.Figure(go.Pie(
        labels=sale_split['sale_type'], values=sale_split['revenue'],
        marker=dict(colors=PAL[:len(sale_split)]),
        hole=0.55,
        textinfo='label+percent',
        textfont=dict(size=11, color='#e2e8f0'),
        hovertemplate='%{label}<br>$%{value:,.2f}<extra></extra>',
    ))
    fig4.update_layout(
        height=300,
        title=dict(text="Direct vs Platform Sales", font=dict(size=13, color='#94a3b8')),
        paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(15,23,42,.6)',
        legend=dict(font=dict(size=11, color='#94a3b8'), bgcolor='rgba(0,0,0,0)'),
        margin=dict(l=0, r=0, t=30, b=0),
    )
    st.plotly_chart(fig4, use_container_width=True)

# ---- CHARTS ROW 3: Units over time + Top countries ----
ch5, ch6 = st.columns([3, 2])

with ch5:
    weekly_units = df.groupby(['week', 'platform'])['quantity'].sum().reset_index()
    plats = sorted(weekly_units['platform'].unique())
    fig5 = go.Figure()
    for i, p in enumerate(plats):
        pd_ = weekly_units[weekly_units['platform'] == p].sort_values('week')
        fig5.add_trace(go.Bar(x=pd_['week'], y=pd_['quantity'], name=p, marker_color=PAL[i % len(PAL)]))
    fig5.update_layout(barmode='stack', height=300, title=dict(text="Weekly Units by Platform", font=dict(size=13, color='#94a3b8')), **PLOT_LAYOUT)
    st.plotly_chart(fig5, use_container_width=True)

with ch6:
    top_countries = df.groupby('country')['revenue'].sum().sort_values(ascending=False).head(10).reset_index()
    fig6 = go.Figure(go.Bar(
        x=top_countries['revenue'], y=top_countries['country'],
        orientation='h',
        marker_color=PAL[1],
        text=[fmt_currency(v) for v in top_countries['revenue']],
        textposition='outside', textfont=dict(size=10, color='#94a3b8'),
    ))
    fig6.update_layout(
        height=300, title=dict(text="Top 10 Countries by Revenue", font=dict(size=13, color='#94a3b8')),
        showlegend=False, **PLOT_LAYOUT,
        yaxis=dict(autorange='reversed', tickfont=dict(size=10, color='#94a3b8')),
        xaxis=dict(tickprefix='$', tickfont=dict(size=10, color='#64748b')),
    )
    st.plotly_chart(fig6, use_container_width=True)

# ---- DATA TABLE ----
st.markdown('<div class="section-label">Data Table</div>', unsafe_allow_html=True)

display_cols = ['date', 'game', 'platform', 'country', 'bp_region', 'product_type',
                'sale_type', 'quantity', 'revenue', 'net_revenue_usd', 'pre_order']
display_cols = [c for c in display_cols if c in df.columns]
table_df = df[display_cols].copy().reset_index(drop=True)
table_df['date'] = table_df['date'].dt.strftime('%Y-%m-%d')

search_col, _, count_col = st.columns([3, 3, 1])
with search_col:
    search_term = st.text_input("", placeholder="Search rows...", label_visibility="collapsed")
with count_col:
    st.markdown(
        f"<div style='text-align:right;font-size:.72rem;color:#475569;padding-top:.6rem;'>{len(table_df):,} rows</div>",
        unsafe_allow_html=True,
    )

if search_term:
    mask = table_df.astype(str).apply(lambda c: c.str.contains(search_term, case=False, na=False)).any(axis=1)
    table_df = table_df[mask].reset_index(drop=True)

for c in ['revenue', 'net_revenue_usd']:
    if c in table_df.columns:
        table_df[c] = table_df[c].apply(lambda v: f"${v:,.2f}" if pd.notna(v) else "")

col_cfg = {}
if 'revenue' in table_df.columns:
    col_cfg['revenue'] = st.column_config.TextColumn("Revenue (USD)")
if 'net_revenue_usd' in table_df.columns:
    col_cfg['net_revenue_usd'] = st.column_config.TextColumn("Net Revenue (USD)")
if 'quantity' in table_df.columns:
    col_cfg['quantity'] = st.column_config.NumberColumn("Units", format="%d")
if 'sale_type' in table_df.columns:
    col_cfg['sale_type'] = st.column_config.TextColumn("Sale Type")

st.dataframe(table_df, use_container_width=True, height=460, column_config=col_cfg, hide_index=True)

# ---- AI CHATBOT ----
st.markdown('<div class="section-label">AI Analyst</div>', unsafe_allow_html=True)

if "sales_chat_history" not in st.session_state:
    st.session_state.sales_chat_history = []

# Build a compact data summary to send with every prompt
def build_data_context(df):
    weekly = df.groupby('week')[['revenue', 'quantity']].sum()
    by_game = df.groupby('game')[['revenue', 'quantity']].sum().to_dict()
    by_platform = df.groupby('platform')[['revenue', 'quantity']].sum().to_dict()
    by_region = df.groupby('bp_region')['revenue'].sum().to_dict()
    by_sale_type = df.groupby('sale_type')['revenue'].sum().to_dict()
    top_countries = df.groupby('country')['revenue'].sum().sort_values(ascending=False).head(15).to_dict()
    weekly_str = weekly.to_string()

    base_game_units = df[df['product_type'].str.lower().str.contains('base|game|standard', na=False)]['quantity'].sum() if 'product_type' in df.columns else 0
    dlc_units = df[df['product_type'].str.lower().str.contains('dlc|add', na=False)]['quantity'].sum() if 'product_type' in df.columns else 0

    ctx = f"""
You are an expert sales data analyst. Answer questions about this sales dataset concisely and accurately.

DATASET OVERVIEW:
- Date range: {df['date'].min().strftime('%Y-%m-%d')} to {df['date'].max().strftime('%Y-%m-%d')}
- Total transactions: {len(df):,}
- Total units sold: {df['quantity'].sum():,}
- Total gross revenue (USD): ${df['revenue'].sum():,.2f}
- Total net revenue (USD): ${df['net_revenue_usd'].sum():,.2f}
- Games: {', '.join(sorted(df['game'].unique()))}
- Platforms: {', '.join(sorted(df['platform'].unique()))}
- Regions: {', '.join(sorted(df['bp_region'].unique()))}
- Product types: {', '.join(sorted(df['product_type'].unique()))}
- Sale types: {', '.join(sorted(df['sale_type'].unique()))}

WEEKLY REVENUE & UNITS:
{weekly_str}

REVENUE BY GAME: {json.dumps({k: round(v,2) for k,v in by_game['revenue'].items()}, default=str)}
UNITS BY GAME: {json.dumps({k: int(v) for k,v in by_game['quantity'].items()}, default=str)}

REVENUE BY PLATFORM: {json.dumps({k: round(v,2) for k,v in by_platform['revenue'].items()}, default=str)}
UNITS BY PLATFORM: {json.dumps({k: int(v) for k,v in by_platform['quantity'].items()}, default=str)}

REVENUE BY REGION: {json.dumps({k: round(v,2) for k,v in by_region.items()}, default=str)}
REVENUE BY SALE TYPE: {json.dumps({k: round(v,2) for k,v in by_sale_type.items()}, default=str)}
TOP COUNTRIES BY REVENUE: {json.dumps({k: round(v,2) for k,v in top_countries.items()}, default=str)}

BASE GAME UNITS: {int(base_game_units):,}
DLC UNITS: {int(dlc_units):,}
DLC ATTACH RATE: {(dlc_units/base_game_units*100):.1f}% (if base game units > 0, otherwise N/A from this data slice)

Note: This may be a filtered slice of the full dataset. Answer based only on what is available.
If a game/platform/metric is not in the data, say so clearly.
Format numbers with commas and dollar signs where appropriate.
Keep answers concise — 2-5 sentences unless detail is needed.
"""
    return ctx

# Render chat history
chat_html = '<div class="chat-container" id="chat-scroll">'
if not st.session_state.sales_chat_history:
    chat_html += '<div class="chat-thinking">Ask me anything about this sales data — WoW changes, platform comparisons, DLC attach rates, top countries, and more.</div>'
for msg in st.session_state.sales_chat_history:
    if msg['role'] == 'user':
        chat_html += f'<div class="chat-msg-user"><div class="chat-bubble-user">{msg["content"]}</div></div>'
    else:
        content = msg['content'].replace('\n', '<br>')
        chat_html += f'<div class="chat-msg-ai"><div class="chat-bubble-ai">{content}</div></div>'
chat_html += '</div>'
st.markdown(chat_html, unsafe_allow_html=True)

# Input row
input_col, btn_col, clear_col = st.columns([6, 1, 1])
with input_col:
    user_input = st.text_input("", placeholder="e.g. What is the WoW revenue change? Which platform drives the most units?", label_visibility="collapsed", key="sales_chat_input")
with btn_col:
    send = st.button("Ask", use_container_width=True, type="primary")
with clear_col:
    if st.button("Clear", use_container_width=True):
        st.session_state.sales_chat_history = []
        st.rerun()

if send and user_input.strip():
    st.session_state.sales_chat_history.append({"role": "user", "content": user_input.strip()})

    data_ctx = build_data_context(df)
    messages = [{"role": "user" if m["role"] == "user" else "assistant", "content": m["content"]}
                for m in st.session_state.sales_chat_history[:-1]]
    messages.append({"role": "user", "content": f"{data_ctx}\n\nUSER QUESTION: {user_input.strip()}"})

    try:
        import requests
        resp = requests.post(
            "https://api.anthropic.com/v1/messages",
            headers={"Content-Type": "application/json", "x-api-key": st.secrets["ANTHROPIC_API_KEY"], "anthropic-version": "2023-06-01"},
            json={
                "model": "claude-sonnet-4-20250514",
                "max_tokens": 1000,
                "system": "You are a concise, accurate sales data analyst. Answer questions about the provided sales data. Be direct and data-driven.",
                "messages": messages,
            },
            timeout=30,
        )
        resp.raise_for_status()
        answer = resp.json()["content"][0]["text"]
    except Exception as e:
        answer = f"Sorry, I couldn't reach the AI API: {e}"

    st.session_state.sales_chat_history.append({"role": "assistant", "content": answer})
    st.rerun()