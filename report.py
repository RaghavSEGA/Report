import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go

st.set_page_config(layout="wide", page_title="KPI Report", page_icon="=")

st.markdown(
    """
    <style>
    /* Hide the deploy toolbar */
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

    /* Metric cards — default and active states */
    .kpi-card {
        background: #1e293b;
        border: 1px solid #334155;
        border-radius: 8px;
        padding: 1rem 1.25rem;
        cursor: pointer;
        transition: border-color .15s, background .15s;
        user-select: none;
    }
    .kpi-card:hover { border-color: #64748b; background: #273548; }
    .kpi-card.active-green  { border-color: #22c55e !important; background: #052e16 !important; }
    .kpi-card.active-amber  { border-color: #f97316 !important; background: #7c2d12 !important; }
    .kpi-card.active-red    { border-color: #ef4444 !important; background: #2d0a0f !important; }
    .kpi-card-label { font-size: .65rem; text-transform: uppercase; letter-spacing: .1em; color: #64748b; margin-bottom: .3rem; }
    .kpi-card-value { font-size: 1.8rem; font-weight: 700; line-height: 1; }
    .kpi-card-sub { font-size: .7rem; color: #475569; margin-top: .25rem; }

    /* Legend pills */
    .legend { display: flex; gap: .75rem; flex-wrap: wrap; margin-bottom: 1rem; align-items: center; }
    .pill { display: inline-flex; align-items: center; gap: .35rem; font-size: .72rem; color: #94a3b8; background: #1e293b; border: 1px solid #334155; border-radius: 20px; padding: .25rem .7rem; }
    .dot { width: 8px; height: 8px; border-radius: 50%; flex-shrink: 0; }

    /* Section divider */
    .section-label { font-size: .68rem; font-weight: 600; text-transform: uppercase; letter-spacing: .1em; color: #475569; margin: 2rem 0 .75rem; border-bottom: 1px solid #1e293b; padding-bottom: .4rem; }

    /* Left-align all dataframe cell content */
    [data-testid="stDataFrame"] td, [data-testid="stDataFrame"] th { text-align: left !important; }
    [data-testid="stDataFrame"] [class*="cell"] { text-align: left !important; justify-content: flex-start !important; }

    /* Active filter banner */
    .filter-banner {
        display: flex; align-items: center; justify-content: space-between;
        background: #1e293b; border: 1px solid #334155; border-radius: 6px;
        padding: .5rem .9rem; margin-bottom: .75rem; font-size: .78rem; color: #94a3b8;
    }
    .filter-banner b { color: #e2e8f0; }
    </style>
    """,
    unsafe_allow_html=True,
)

# ---- HELPERS ----
def _to_num(v):
    if pd.isna(v): return np.nan
    s = str(v).strip().replace(",", "")
    if "%" in s:
        try: return float(s.replace("%", "")) / 100.0
        except: return np.nan
    try: return float(s)
    except: return np.nan

# ---- SIDEBAR ----
with st.sidebar:
    st.markdown("<div style='font-size:1rem;font-weight:700;color:#f1f5f9;margin-bottom:.25rem;'>KPI Report</div>", unsafe_allow_html=True)
    st.markdown("<div style='font-size:.65rem;color:#475569;text-transform:uppercase;letter-spacing:.1em;margin-bottom:1rem;'>Performance Dashboard</div>", unsafe_allow_html=True)
    st.markdown('<span class="sidebar-section">Data Source</span>', unsafe_allow_html=True)
    uploaded_file = st.file_uploader("", type=["csv", "xlsx"], label_visibility="collapsed")
    if uploaded_file is None:
        st.info("Upload a CSV or XLSX to begin.")
        st.stop()

# ---- LOAD DATA ----
try:
    data = pd.read_csv(uploaded_file) if uploaded_file.name.lower().endswith(".csv") else pd.read_excel(uploaded_file)
except Exception as e:
    st.sidebar.error(f"Error reading file: {e}")
    st.stop()

if data is None or data.empty:
    st.warning("No data found in the uploaded file.")
    st.stop()

df = data.copy()

# ---- SIDEBAR FILTERS ----
with st.sidebar:
    if "Title" in df.columns:
        titles = sorted(df["Title"].dropna().unique())
        st.markdown('<span class="sidebar-section">Title</span>', unsafe_allow_html=True)
        sel_titles = st.multiselect("", options=titles, default=titles, label_visibility="collapsed")
        df = df[df["Title"].isin(sel_titles)]

    if "KPI" in df.columns:
        kpis_orig = sorted(df["KPI"].dropna().unique())
        display_map, display_options = {}, []
        for orig in kpis_orig:
            d = orig.replace("Channel ", "")
            k = d if d not in display_map else orig
            display_map[k] = orig
            display_options.append(k)
        st.markdown('<span class="sidebar-section">KPI Type</span>', unsafe_allow_html=True)
        sel_disp = st.multiselect("", options=display_options, default=display_options, label_visibility="collapsed")
        df = df[df["KPI"].isin([display_map[d] for d in sel_disp])]

    if "Beat" in df.columns:
        beats = sorted(df["Beat"].dropna().unique())
        st.markdown('<span class="sidebar-section">Beat</span>', unsafe_allow_html=True)
        sel_beats = st.multiselect("", options=beats, default=beats, label_visibility="collapsed")
        df = df[df["Beat"].isin(sel_beats)]

# ---- IDENTIFY KEY COLUMNS ----
p45_col = next((c for c in df.columns if "45" in c or "45th" in c.lower()), None)
p60_col = next((c for c in df.columns if "60" in c or "60th" in c.lower()), None)
actuals_cols = [c for c in df.columns if "actual" in c.lower()]

# ---- PRE-COMPUTE STATUS FOR ALL ROWS ----
STATUS_COL_PREFIX = "_status_"
dft = df.reset_index(drop=True).copy()

if actuals_cols and p45_col and p60_col:
    p45t = dft[p45_col].apply(_to_num)
    p60t = dft[p60_col].apply(_to_num)
    for col in actuals_cols:
        nums = dft[col].apply(_to_num)
        statuses = []
        for i in range(len(dft)):
            av, p4, p6 = nums.iloc[i], p45t.iloc[i], p60t.iloc[i]
            if any(pd.isna(x) for x in [av, p4, p6]):
                statuses.append("")
            elif av >= p6:
                statuses.append("Above target")
            elif av >= p4:
                statuses.append("On target")
            else:
                statuses.append("Below target")
        dft[STATUS_COL_PREFIX + col] = statuses

# ---- PAGE HEADER ----
kpi_bit = f" &nbsp;·&nbsp; {df['KPI'].nunique()} KPI types" if "KPI" in df.columns else ""
st.markdown(
    f"<h1>KPI Report</h1>"
    f"<div style='font-size:.78rem;color:#475569;margin-bottom:1.5rem;'>"
    f"{len(df):,} rows{kpi_bit}</div>",
    unsafe_allow_html=True,
)

# ---- SUMMARY CARDS (clickable filters) ----
# active_filter is stored in session_state as None | ("Above target"|"On target"|"Below target", col)
if "active_filter" not in st.session_state:
    st.session_state.active_filter = None

if actuals_cols and p45_col and p60_col:
    p45_s = dft[p45_col].apply(_to_num)
    p60_s = dft[p60_col].apply(_to_num)
    cards = []
    for col in actuals_cols:
        sc = STATUS_COL_PREFIX + col
        if sc not in dft.columns: continue
        cards.append((
            col,
            int((dft[sc] == "Above target").sum()),
            int((dft[sc] == "On target").sum()),
            int((dft[sc] == "Below target").sum()),
            int(dft[sc].ne("").sum()),
        ))

    if cards:
        st.markdown('<div class="section-label">Summary — click a card to filter the table</div>', unsafe_allow_html=True)
        col_widgets = st.columns(min(len(cards) * 3, 9))
        for i, (col, g, a, r, total) in enumerate(cards):
            b = i * 3

            def make_card(label, value, color, status, col_name, css_class, sub):
                af = st.session_state.active_filter
                is_active = af == (status, col_name)
                active_cls = f" {css_class}" if is_active else ""
                return (
                    f'<div class="kpi-card{active_cls}" style="margin-bottom:.5rem;">'
                    f'<div class="kpi-card-label">{label}</div>'
                    f'<div class="kpi-card-value" style="color:{color};">{value}</div>'
                    f'<div class="kpi-card-sub">{sub}</div>'
                    f'</div>'
                )

            with col_widgets[b]:
                st.markdown(make_card(f"{col} — Above 60th", g, "#4ade80", "Above target", col, "active-green", f"of {total} rows"), unsafe_allow_html=True)
                if st.button("Filter", key=f"btn_g_{col}_{i}",
                             type="primary" if st.session_state.active_filter == ("Above target", col) else "secondary",
                             use_container_width=True):
                    st.session_state.active_filter = None if st.session_state.active_filter == ("Above target", col) else ("Above target", col)
                    st.rerun()

            with col_widgets[b+1]:
                st.markdown(make_card(f"{col} — In Range", a, "#fb923c", "On target", col, "active-amber", "45th–60th pct"), unsafe_allow_html=True)
                if st.button("Filter", key=f"btn_a_{col}_{i}",
                             type="primary" if st.session_state.active_filter == ("On target", col) else "secondary",
                             use_container_width=True):
                    st.session_state.active_filter = None if st.session_state.active_filter == ("On target", col) else ("On target", col)
                    st.rerun()

            with col_widgets[b+2]:
                st.markdown(make_card(f"{col} — Below 45th", r, "#f87171", "Below target", col, "active-red", "needs attention"), unsafe_allow_html=True)
                if st.button("Filter", key=f"btn_r_{col}_{i}",
                             type="primary" if st.session_state.active_filter == ("Below target", col) else "secondary",
                             use_container_width=True):
                    st.session_state.active_filter = None if st.session_state.active_filter == ("Below target", col) else ("Below target", col)
                    st.rerun()

# ---- DATA TABLE ----
st.markdown('<div class="section-label">Data Table</div>', unsafe_allow_html=True)

# Apply card filter
table_df = dft.copy()
af = st.session_state.active_filter
if af is not None:
    status_val, filter_col = af
    sc = STATUS_COL_PREFIX + filter_col
    if sc in table_df.columns:
        table_df = table_df[table_df[sc] == status_val].reset_index(drop=True)

# Search bar + active filter banner + row count
search_col, _, count_col = st.columns([3, 3, 1])
with search_col:
    search_term = st.text_input("", placeholder="Search rows...", label_visibility="collapsed")
with count_col:
    st.markdown(
        f"<div style='text-align:right;font-size:.72rem;color:#475569;padding-top:.6rem;'>{len(table_df):,} rows</div>",
        unsafe_allow_html=True,
    )

# Active filter banner
if af is not None:
    status_val, filter_col = af
    status_colors = {"Above target": "#4ade80", "On target": "#fb923c", "Below target": "#f87171"}
    color = status_colors.get(status_val, "#94a3b8")
    st.markdown(
        f'<div class="filter-banner">'
        f'<span>Filtered by: <b style="color:{color};">{status_val}</b> on <b>{filter_col}</b> &nbsp;·&nbsp; {len(table_df):,} matching rows</span>'
        f'</div>',
        unsafe_allow_html=True,
    )

# Apply text search
if search_term:
    mask = table_df.astype(str).apply(lambda c: c.str.contains(search_term, case=False, na=False)).any(axis=1)
    table_df = table_df[mask].reset_index(drop=True)

# Legend pills
status_cols_present = [c for c in table_df.columns if c.startswith(STATUS_COL_PREFIX)]
if status_cols_present:
    st.markdown(
        '<div class="legend">'
        '<span style="font-size:.68rem;color:#475569;">Color key:</span>'
        '<span class="pill"><span class="dot" style="background:#4ade80;"></span> Above target (&gt;60th)</span>'
        '<span class="pill"><span class="dot" style="background:#fb923c;"></span> On target (45–60th)</span>'
        '<span class="pill"><span class="dot" style="background:#f87171;"></span> Below target (&lt;45th)</span>'
        '</div>',
        unsafe_allow_html=True,
    )

# Row styling
def _style_table(row):
    styles = [""] * len(row)
    # Default text color for all cells so nothing inherits a dark color
    base = "color:#e2e8f0;"
    styles = [base] * len(row)
    for col in actuals_cols:
        sc = STATUS_COL_PREFIX + col
        if sc not in row.index: continue
        status = row[sc]
        if status == "Above target":
            bg = "background-color:#14532d; color:#bbf7d0;"
            sc_style = "background-color:#14532d; color:#4ade80; font-weight:600;"
        elif status == "On target":
            bg = "background-color:#7c2d12; color:#fed7aa;"
            sc_style = "background-color:#7c2d12; color:#fb923c; font-weight:600;"
        elif status == "Below target":
            bg = "background-color:#4c0519; color:#fecdd3;"
            sc_style = "background-color:#4c0519; color:#fca5a5; font-weight:600;"
        else:
            continue
        if col in row.index:
            styles[list(row.index).index(col)] = bg
        styles[list(row.index).index(sc)] = sc_style
    return styles

# Column config
# Format numeric columns as strings so we control alignment fully
def _fmt(v):
    n = _to_num(v)
    if pd.isna(n): return ""
    if abs(n) < 10: return f"{n:.2f}"
    return f"{n:,.0f}"

for col in actuals_cols + ([p45_col] if p45_col else []) + ([p60_col] if p60_col else []):
    if col in table_df.columns:
        table_df[col] = table_df[col].apply(_fmt)

col_config = {}
for col in actuals_cols:
    sc = STATUS_COL_PREFIX + col
    if sc in table_df.columns:
        label = col.replace("Actual", "").replace("actual", "").strip() or col
        col_config[sc] = st.column_config.TextColumn(label=f"{label} Status", help="Relative to 45th/60th benchmarks", width="medium")
    col_config[col] = st.column_config.TextColumn(label=col)
if p45_col:
    col_config[p45_col] = st.column_config.TextColumn(label=p45_col)
if p60_col:
    col_config[p60_col] = st.column_config.TextColumn(label=p60_col)

styled_df = table_df.style.apply(_style_table, axis=1)

st.dataframe(styled_df, use_container_width=True, height=520, column_config=col_config, hide_index=True)

# ---- TREND ANALYSIS (no expander) ----
st.markdown('<div class="section-label">Trend Analysis</div>', unsafe_allow_html=True)

if "KPI" in df.columns:
    kpi_options = sorted(df["KPI"].dropna().unique())
    if kpi_options:
        pc, _ = st.columns([2, 5])
        with pc:
            sel_kpi = st.selectbox("Select KPI", options=kpi_options)
        dk = df[df["KPI"] == sel_kpi].copy()
        if not dk.empty and "Beat" in dk.columns:
            p45 = next((c for c in dk.columns if "45" in c or "45th" in c.lower()), None)
            p60 = next((c for c in dk.columns if "60" in c or "60th" in c.lower()), None)
            ac = [c for c in dk.columns if "actual" in c.lower()]
            bo = list(dk["Beat"].dropna().unique())
            for c in ([p45] if p45 else []) + ([p60] if p60 else []) + ac:
                dk[c] = dk[c].apply(_to_num)
            am = {**({p45: 'mean'} if p45 else {}), **({p60: 'mean'} if p60 else {}), **{c: 'mean' for c in ac}}
            agg = dk.groupby("Beat").agg(am).reindex(bo)
            fig = go.Figure()
            if p45 and p60:
                fig.add_trace(go.Scatter(x=agg.index, y=agg[p60], mode='lines',
                    line=dict(color='rgba(0,0,0,0)'), showlegend=False, hoverinfo='skip'))
                fig.add_trace(go.Scatter(x=agg.index, y=agg[p45], mode='lines',
                    fill='tonexty', fillcolor='rgba(148,163,184,.1)',
                    line=dict(color='rgba(148,163,184,.3)', width=1, dash='dot'),
                    name='Target Range (45th–60th)'))
            pal = ['#60a5fa', '#34d399', '#a78bfa', '#fb923c', '#f472b6']
            for i, c in enumerate(ac):
                fig.add_trace(go.Scatter(x=agg.index, y=agg[c], mode='lines+markers',
                    name=c, line=dict(color=pal[i % len(pal)], width=2),
                    marker=dict(size=6, line=dict(width=1.5, color='#111827'))))
            fig.update_layout(
                paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(15,23,42,.6)',
                template='plotly_dark', height=380, margin=dict(l=0, r=0, t=24, b=0),
                xaxis=dict(gridcolor='rgba(255,255,255,.05)', linecolor='rgba(255,255,255,.08)',
                    tickfont=dict(size=11, color='#64748b')),
                yaxis=dict(gridcolor='rgba(255,255,255,.05)', linecolor='rgba(255,255,255,.08)',
                    tickfont=dict(size=11, color='#64748b')),
                legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1,
                    font=dict(size=11, color='#94a3b8'), bgcolor='rgba(0,0,0,0)'),
                hovermode='x unified',
                hoverlabel=dict(bgcolor='#1e293b', bordercolor='#334155',
                    font=dict(size=12, color='#e2e8f0')),
            )
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("No data for selected KPI or no Beat column found.")
    else:
        st.info("No KPI options available.")
else:
    st.info("No KPI column found.")

# ---- AI CHATBOT ----
st.markdown('<div class="section-label">AI Analyst</div>', unsafe_allow_html=True)

if "kpi_chat_history" not in st.session_state:
    st.session_state.kpi_chat_history = []

def build_kpi_context(df, actuals_cols, p45_col, p60_col):
    lines = [
        "You are an expert KPI analyst. Answer questions about this KPI report data concisely and accurately.",
        f"DATASET OVERVIEW:",
        f"- Total rows: {len(df):,}",
        f"- Columns: {', '.join(df.columns.tolist())}",
    ]
    if "KPI" in df.columns:
        lines.append(f"- KPI types: {', '.join(sorted(df['KPI'].dropna().unique()))}")
    if "Title" in df.columns:
        lines.append(f"- Titles: {', '.join(sorted(df['Title'].dropna().unique()))}")
    if "Beat" in df.columns:
        lines.append(f"- Beats: {', '.join(sorted(df['Beat'].dropna().unique()))}")
    if p45_col:
        lines.append(f"- 45th percentile column: {p45_col}")
    if p60_col:
        lines.append(f"- 60th percentile column: {p60_col}")
    if actuals_cols:
        lines.append(f"- Actuals columns: {', '.join(actuals_cols)}")
        for col in actuals_cols:
            nums = df[col].apply(lambda v: float(str(v).replace(',','').replace('%','')) if pd.notna(v) else float('nan'))
            lines.append(f"  {col}: min={nums.min():.2f}, max={nums.max():.2f}, mean={nums.mean():.2f}, count={nums.notna().sum()}")
    lines.append("\nSAMPLE DATA (first 30 rows):")
    lines.append(df.head(30).to_string())
    lines.append("\nAnswer based only on available data. Keep answers to 2-5 sentences unless detail is needed.")
    return "\n".join(lines)

# Render chat history
chat_html = '<div class="chat-container">'
if not st.session_state.kpi_chat_history:
    chat_html += '<div class="chat-thinking">Ask me anything about this KPI data — performance vs benchmarks, which titles are above/below target, trends across beats, and more.</div>'
for msg in st.session_state.kpi_chat_history:
    if msg['role'] == 'user':
        chat_html += f'<div class="chat-msg-user"><div class="chat-bubble-user">{msg["content"]}</div></div>'
    else:
        content = msg['content'].replace('\n', '<br>')
        chat_html += f'<div class="chat-msg-ai"><div class="chat-bubble-ai">{content}</div></div>'
chat_html += '</div>'
st.markdown(chat_html, unsafe_allow_html=True)

st.markdown("""
<style>
.chat-container { background: #0f172a; border: 1px solid #1e293b; border-radius: 10px; padding: 1rem 1.25rem; margin-bottom: 1rem; max-height: 420px; overflow-y: auto; }
.chat-msg-user { display: flex; justify-content: flex-end; margin-bottom: .75rem; }
.chat-msg-ai { display: flex; justify-content: flex-start; margin-bottom: .75rem; }
.chat-bubble-user { background: #1d4ed8; color: #fff; border-radius: 14px 14px 2px 14px; padding: .55rem .9rem; font-size: .82rem; max-width: 75%; line-height: 1.5; }
.chat-bubble-ai { background: #1e293b; color: #e2e8f0; border-radius: 14px 14px 14px 2px; padding: .55rem .9rem; font-size: .82rem; max-width: 85%; line-height: 1.5; border: 1px solid #334155; }
.chat-thinking { color: #475569; font-style: italic; font-size: .78rem; }
</style>
""", unsafe_allow_html=True)

input_col, btn_col, clear_col = st.columns([6, 1, 1])
with input_col:
    user_input = st.text_input("", placeholder="e.g. Which titles are above target? How does Beat 1 compare to Beat 2?", label_visibility="collapsed", key="kpi_chat_input")
with btn_col:
    send = st.button("Ask", use_container_width=True, type="primary")
with clear_col:
    if st.button("Clear", use_container_width=True):
        st.session_state.kpi_chat_history = []
        st.rerun()

if send and user_input.strip():
    st.session_state.kpi_chat_history.append({"role": "user", "content": user_input.strip()})
    ctx = build_kpi_context(df, actuals_cols, p45_col, p60_col)
    messages = [{"role": "user" if m["role"] == "user" else "assistant", "content": m["content"]}
                for m in st.session_state.kpi_chat_history[:-1]]
    messages.append({"role": "user", "content": f"{ctx}\n\nUSER QUESTION: {user_input.strip()}"})
    try:
        import requests
        resp = requests.post(
            "https://api.anthropic.com/v1/messages",
            headers={"Content-Type": "application/json", "x-api-key": st.secrets["ANTHROPIC_API_KEY"], "anthropic-version": "2023-06-01"},
            json={
                "model": "claude-sonnet-4-20250514",
                "max_tokens": 1000,
                "system": "You are a concise, accurate KPI analyst. Answer questions about the provided KPI data. Be direct and data-driven.",
                "messages": messages,
            },
            timeout=30,
        )
        resp.raise_for_status()
        answer = resp.json()["content"][0]["text"]
    except Exception as e:
        answer = f"Sorry, I couldn't reach the AI API: {e}"
    st.session_state.kpi_chat_history.append({"role": "assistant", "content": answer})
    st.rerun()