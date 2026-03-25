import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import json
import random
import time
import hashlib
import hmac
import base64

st.set_page_config(layout="wide", page_title="Sales Dashboard", page_icon="=")

# ─────────────────────────────────────────────────────────────
# OTP AUTHENTICATION
# ─────────────────────────────────────────────────────────────

ALLOWED_DOMAIN     = "@segaamerica.com"
OTP_EXPIRY_SECS    = 600   # 10 minutes
COOKIE_EXPIRY_DAYS = 7
COOKIE_NAME        = "sega_sales_auth"

def _send_otp(email: str, code: str) -> bool:
    """Send OTP via AWS SES. Returns True on success."""
    try:
        import boto3
        from botocore.exceptions import ClientError
        ses = boto3.client(
            "ses",
            region_name=st.secrets.get("AWS_SES_REGION", "us-east-1"),
            aws_access_key_id=st.secrets.get("AWS_ACCESS_KEY_ID", ""),
            aws_secret_access_key=st.secrets.get("AWS_SECRET_ACCESS_KEY", ""),
        )
        ses.send_email(
            Source=st.secrets.get("EMAIL_FROM", "noreply@segaamerica.com"),
            Destination={"ToAddresses": [email]},
            Message={
                "Subject": {"Data": "SEGA Sales Dashboard — Your verification code", "Charset": "UTF-8"},
                "Body": {
                    "Text": {
                        "Data": f"Your SEGA Sales Dashboard verification code is: {code}\n\nThis code expires in 10 minutes.\nIf you didn't request this, you can safely ignore this email.",
                        "Charset": "UTF-8",
                    },
                    "Html": {
                        "Data": f"""
                        <div style="font-family:Arial,sans-serif;max-width:480px;margin:0 auto;padding:32px 24px;">
                          <div style="font-size:22px;font-weight:900;letter-spacing:0.1em;color:#1A6BFF;margin-bottom:4px;">SEGA</div>
                          <div style="font-size:14px;color:#444;margin-bottom:28px;">Sales Dashboard</div>
                          <div style="font-size:14px;color:#222;margin-bottom:16px;">Your verification code is:</div>
                          <div style="font-size:42px;font-weight:900;letter-spacing:0.18em;color:#1a1a2e;
                                      background:#f0f4ff;border-radius:8px;padding:18px 24px;
                                      display:inline-block;margin-bottom:24px;">{code}</div>
                          <div style="font-size:12px;color:#888;">
                            This code expires in 10 minutes.<br>
                            If you didn't request this, you can safely ignore this email.
                          </div>
                        </div>""",
                        "Charset": "UTF-8",
                    },
                },
            },
        )
        return True
    except ClientError as e:
        st.error(f"SES error: {e.response['Error']['Message']}")
        return False
    except Exception as e:
        st.error(f"Failed to send email: {e}")
        return False

def _sign_cookie(email: str) -> str:
    """Create an HMAC-signed token: email|expiry|signature."""
    secret = st.secrets.get("COOKIE_SIGNING_KEY", "fallback-change-this")
    expiry = int(time.time()) + (COOKIE_EXPIRY_DAYS * 86400)
    payload = f"{email}|{expiry}"
    sig = hmac.new(secret.encode(), payload.encode(), hashlib.sha256).hexdigest()
    return base64.urlsafe_b64encode(f"{payload}|{sig}".encode()).decode()

def _verify_cookie(token: str) -> str | None:
    """Verify signed cookie. Returns email if valid, None otherwise."""
    try:
        secret = st.secrets.get("COOKIE_SIGNING_KEY", "fallback-change-this")
        decoded = base64.urlsafe_b64decode(token.encode()).decode()
        email, expiry_str, sig = decoded.rsplit("|", 2)
        payload = f"{email}|{expiry_str}"
        expected = hmac.new(secret.encode(), payload.encode(), hashlib.sha256).hexdigest()
        if not hmac.compare_digest(sig, expected):
            return None
        if int(time.time()) > int(expiry_str):
            return None
        return email
    except Exception:
        return None

# ── Check for existing valid cookie ──────────────────────────
try:
    import extra_streamlit_components as stx
    _cookie_manager = stx.CookieManager(key="auth_cookies")
    _existing_cookie = _cookie_manager.get(COOKIE_NAME)
except Exception:
    _cookie_manager = None
    _existing_cookie = None

_cookie_email = _verify_cookie(_existing_cookie) if _existing_cookie else None

# ── Auth state init ───────────────────────────────────────────
for _k, _v in [
    ("auth_verified",   False),
    ("auth_email",      ""),
    ("otp_code",        ""),
    ("otp_email",       ""),
    ("otp_expiry",      0),
    ("otp_sent",        False),
    ("otp_attempts",    0),
]:
    if _k not in st.session_state:
        st.session_state[_k] = _v

# If valid cookie found, mark as verified
if _cookie_email and not st.session_state.auth_verified:
    st.session_state.auth_verified = True
    st.session_state.auth_email    = _cookie_email

# ── Render login gate if not verified ────────────────────────
if not st.session_state.auth_verified:
    st.markdown("""
    <style>
    .auth-wrap {
        max-width: 420px; margin: 5rem auto; padding: 2.5rem 2.5rem 2rem;
        background: var(--surface); border: 1px solid var(--border);
        border-top: 3px solid var(--blue); border-radius: 0 0 10px 10px;
    }
    .auth-logo { font-family:'Inter Tight',sans-serif; font-size:1.6rem; font-weight:900;
                 letter-spacing:0.12em; color:var(--blue) !important; margin-bottom:0.2rem; }
    .auth-title { font-family:'Inter Tight',sans-serif; font-size:1rem; font-weight:700;
                  color:var(--text) !important; margin-bottom:0.25rem; }
    .auth-sub { font-size:0.8rem; color:var(--muted) !important; margin-bottom:1.5rem; }
    .auth-note { font-size:0.72rem; color:var(--muted) !important; margin-top:1rem;
                 text-align:center; line-height:1.5; }
    :root {
        --bg: #0a0c1a; --surface: #0f1120; --border: #232640;
        --blue: #4080ff; --text: #eef0fa; --muted: #5a5f82;
    }
    html, body, .stApp { background: var(--bg) !important; }
    </style>
    """, unsafe_allow_html=True)

    _lc, _mc, _rc = st.columns([1, 2, 1])
    with _mc:
        st.markdown("""
        <div class="auth-wrap">
          <div class="auth-logo">SEGA</div>
          <div class="auth-title">Sales Dashboard</div>
          <div class="auth-sub">Sign in with your SEGA America email</div>
        </div>
        """, unsafe_allow_html=True)

        if not st.session_state.otp_sent:
            _email_input = st.text_input(
                "Email address",
                placeholder="you@segaamerica.com",
                label_visibility="hidden",
                key="auth_email_input",
            )
            _send_btn = st.button("Send verification code", width="stretch")

            if _send_btn and _email_input:
                if not _email_input.strip().lower().endswith(ALLOWED_DOMAIN):
                    st.error(f"Access restricted to {ALLOWED_DOMAIN} addresses.")
                else:
                    _code = str(random.randint(100000, 999999))
                    if _send_otp(_email_input.strip().lower(), _code):
                        st.session_state.otp_code     = _code
                        st.session_state.otp_email    = _email_input.strip().lower()
                        st.session_state.otp_expiry   = time.time() + OTP_EXPIRY_SECS
                        st.session_state.otp_sent     = True
                        st.session_state.otp_attempts = 0
                        st.rerun()

        else:
            st.info(f"Code sent to **{st.session_state.otp_email}** — check your inbox.")
            _code_input = st.text_input(
                "6-digit code",
                placeholder="123456",
                label_visibility="hidden",
                max_chars=6,
                key="auth_code_input",
            )
            _verify_btn = st.button("Verify code", width="stretch")

            if _verify_btn and _code_input:
                if st.session_state.otp_attempts >= 5:
                    st.error("Too many attempts. Please request a new code.")
                    st.session_state.otp_sent = False
                elif time.time() > st.session_state.otp_expiry:
                    st.error("Code has expired. Please request a new one.")
                    st.session_state.otp_sent = False
                elif _code_input.strip() != st.session_state.otp_code:
                    st.session_state.otp_attempts += 1
                    _remaining = 5 - st.session_state.otp_attempts
                    st.error(f"Incorrect code. {_remaining} attempt{'s' if _remaining != 1 else ''} remaining.")
                else:
                    st.session_state.auth_verified = True
                    st.session_state.auth_email    = st.session_state.otp_email
                    st.session_state.otp_code      = ""
                    if _cookie_manager:
                        _token = _sign_cookie(st.session_state.auth_email)
                        _cookie_manager.set(
                            COOKIE_NAME, _token,
                            expires_at=None,
                            key="set_auth_cookie",
                        )
                    st.rerun()

            if st.button("← Use a different email", key="auth_back"):
                st.session_state.otp_sent = False
                st.session_state.otp_code = ""
                st.rerun()

        st.markdown(
            '<div class="auth-note">Restricted to @segaamerica.com addresses only.<br>'
            'Codes expire after 10 minutes.</div>',
            unsafe_allow_html=True,
        )

    st.stop()

# ── Signed-in user + sign out (sidebar) ──────────────────────
with st.sidebar:
    st.markdown(
        f'<div style="font-size:.6rem;font-weight:700;letter-spacing:.12em;text-transform:uppercase;'
        f'color:#5a5f82;margin-bottom:.3rem;">SEGA America</div>'
        f'<div style="font-size:.7rem;font-weight:600;color:#b8bcd4;margin-bottom:.75rem;">'
        f'{st.session_state.auth_email}</div>',
        unsafe_allow_html=True,
    )
    if st.button("Sign out", key="sign_out_btn"):
        if _cookie_manager:
            _cookie_manager.delete(COOKIE_NAME, key="delete_auth_cookie")
        for _k in ["auth_verified", "auth_email", "otp_sent", "otp_code",
                   "otp_email", "otp_expiry", "otp_attempts"]:
            st.session_state[_k] = False if _k == "auth_verified" else ""
        st.rerun()
    st.markdown("<hr style='border:none;border-top:1px solid #232640;margin:.5rem 0 1rem;'>", unsafe_allow_html=True)

st.markdown(
    """
    <style>
    [data-testid="stToolbar"] { display: none !important; }
    [data-testid="stDecoration"] { display: none !important; }
    header[data-testid="stHeader"] { display: none !important; }
    html, body, [class*="css"], .stApp { background-color: #111827 !important; color: #e2e8f0 !important; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif !important; }
    [data-testid="stSidebar"] { background: #0f172a !important; border-right: 1px solid #1e293b !important; }
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
    .chat-container { background: #0f172a; border: 1px solid #1e293b; border-radius: 10px; padding: 1rem 1.25rem; margin-bottom: 1rem; max-height: 420px; overflow-y: auto; }
    .chat-msg-user { display: flex; justify-content: flex-end; margin-bottom: .75rem; }
    .chat-msg-ai { display: flex; justify-content: flex-start; margin-bottom: .75rem; }
    .chat-bubble-user { background: #1d4ed8; color: #fff; border-radius: 14px 14px 2px 14px; padding: .55rem .9rem; font-size: .82rem; max-width: 75%; line-height: 1.5; }
    .chat-bubble-ai { background: #1e293b; color: #e2e8f0; border-radius: 14px 14px 14px 2px; padding: .55rem .9rem; font-size: .82rem; max-width: 85%; line-height: 1.5; border: 1px solid #334155; }
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

def delta_html(curr, prev, label="vs prev period"):
    if prev == 0 or pd.isna(prev): return ""
    pct = (curr - prev) / abs(prev) * 100
    sign = "+" if pct >= 0 else ""
    cls = "kpi-card-delta-pos" if pct >= 0 else "kpi-card-delta-neg"
    return f'<div class="{cls}">{sign}{pct:.1f}% {label}</div>'

# Base plot style — NO xaxis/yaxis here to avoid double-kwarg conflicts
LEGEND_DEFAULT = dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1,
                      font=dict(size=11, color='#94a3b8'), bgcolor='rgba(0,0,0,0)')

def base_layout(height=300, title="", margin=None, legend=None, showlegend=True, hovermode="x unified"):
    return dict(
        height=height,
        title=dict(text=title, font=dict(size=13, color='#94a3b8')),
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(15,23,42,.6)',
        template='plotly_dark',
        margin=margin or dict(l=0, r=0, t=40, b=10),
        hovermode=hovermode,
        hoverlabel=dict(bgcolor='#1e293b', bordercolor='#334155', font=dict(size=12, color='#e2e8f0')),
        legend=legend if legend is not None else LEGEND_DEFAULT,
        showlegend=showlegend,
    )

# Reusable axis styles
AX = dict(gridcolor="rgba(255,255,255,.05)", linecolor="rgba(255,255,255,.08)")
AX_DEFAULT = dict(**AX, tickfont=dict(size=11, color="#64748b"))
AX_MONEY = dict(**AX, tickprefix="$", showticklabels=False, tickfont=dict(size=11, color="#64748b"))
AX_LABEL = dict(**AX, tickfont=dict(size=10, color="#94a3b8"))

PAL = ['#60a5fa', '#34d399', '#a78bfa', '#fb923c', '#f472b6', '#facc15', '#38bdf8']

# ---- SIDEBAR ----
with st.sidebar:
    st.markdown("<div style='font-size:1rem;font-weight:700;color:#f1f5f9;margin-bottom:.25rem;'>Sales Dashboard</div>", unsafe_allow_html=True)
    st.markdown("<div style='font-size:.65rem;color:#475569;text-transform:uppercase;letter-spacing:.1em;margin-bottom:1rem;'>Revenue & Units Analytics</div>", unsafe_allow_html=True)
    st.markdown('<span class="sidebar-section">Data Source</span>', unsafe_allow_html=True)
    uploaded_file = st.file_uploader("Upload file", type=["csv", "xlsx", "gz"], label_visibility="hidden")
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
df['sale_type'] = df['bundle_name'].fillna('Platform Sale').replace('', 'Platform Sale')

# ---- SIDEBAR FILTERS ----
with st.sidebar:
    if 'franchise' in df.columns:
        opts = sorted(df['franchise'].dropna().unique())
        st.markdown('<span class="sidebar-section">Franchise</span>', unsafe_allow_html=True)
        sel = st.multiselect("Franchise", options=opts, default=opts, label_visibility="hidden")
        df = df[df['franchise'].isin(sel)]

    if 'game' in df.columns:
        opts = sorted(df['game'].dropna().unique())
        st.markdown('<span class="sidebar-section">Game</span>', unsafe_allow_html=True)
        sel = st.multiselect("Game", options=opts, default=opts, label_visibility="hidden")
        df = df[df['game'].isin(sel)]

    if 'platform' in df.columns:
        opts = sorted(df['platform'].dropna().unique())
        st.markdown('<span class="sidebar-section">Platform</span>', unsafe_allow_html=True)
        sel = st.multiselect("Platform", options=opts, default=opts, label_visibility="hidden")
        df = df[df['platform'].isin(sel)]

    if 'bp_region' in df.columns:
        opts = sorted(df['bp_region'].dropna().unique())
        st.markdown('<span class="sidebar-section">Region</span>', unsafe_allow_html=True)
        sel = st.multiselect("Region", options=opts, default=opts, label_visibility="hidden")
        df = df[df['bp_region'].isin(sel)]

    if 'product_type' in df.columns:
        opts = sorted(df['product_type'].dropna().unique())
        st.markdown('<span class="sidebar-section">Product Type</span>', unsafe_allow_html=True)
        sel = st.multiselect("Product Type", options=opts, default=opts, label_visibility="hidden")
        df = df[df['product_type'].isin(sel)]

    if 'date' in df.columns and df['date'].notna().any():
        min_d, max_d = df['date'].min().date(), df['date'].max().date()
        st.markdown('<span class="sidebar-section">Date Range</span>', unsafe_allow_html=True)
        date_range = st.date_input("Date range", value=(min_d, max_d), min_value=min_d, max_value=max_d, label_visibility="hidden")
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

# ---- INLINE DATE RANGE PICKER ----
if 'date' in df.columns and df['date'].notna().any():
    all_min = raw['date'].pipe(pd.to_datetime).min().date()
    all_max = raw['date'].pipe(pd.to_datetime).max().date()
    dr1, dr2, _ = st.columns([2, 2, 5])
    with dr1:
        st.markdown('<span style="font-size:.65rem;text-transform:uppercase;letter-spacing:.1em;color:#64748b;">From</span>', unsafe_allow_html=True)
        start_date = st.date_input("From date", value=df['date'].min().date(), min_value=all_min, max_value=all_max, label_visibility="hidden", key="inline_start")
    with dr2:
        st.markdown('<span style="font-size:.65rem;text-transform:uppercase;letter-spacing:.1em;color:#64748b;">To</span>', unsafe_allow_html=True)
        end_date = st.date_input("To date", value=df['date'].max().date(), min_value=all_min, max_value=all_max, label_visibility="hidden", key="inline_end")
    df = df[(df['date'].dt.date >= start_date) & (df['date'].dt.date <= end_date)]

# ---- WoW DELTAS ----
weeks = sorted(df['week'].dropna().unique())
curr_rev = prev_rev = curr_units = prev_units = 0
if len(weeks) >= 2:
    curr_w, prev_w = weeks[-1], weeks[-2]
    curr_rev   = df[df['week'] == curr_w]['revenue'].sum()
    prev_rev   = df[df['week'] == prev_w]['revenue'].sum()
    curr_units = df[df['week'] == curr_w]['quantity'].sum()
    prev_units = df[df['week'] == prev_w]['quantity'].sum()

total_rev   = df['revenue'].sum()
total_units = df['quantity'].sum()
total_net   = df['net_revenue_usd'].sum() if 'net_revenue_usd' in df.columns else 0
direct_rev  = df[df['sale_type'] == 'Direct Package Sale']['revenue'].sum()
direct_pct  = (direct_rev / total_rev * 100) if total_rev > 0 else 0
avg_price   = total_rev / total_units if total_units > 0 else 0

# ---- SUMMARY CARDS ----
st.markdown('<div class="section-label">Summary</div>', unsafe_allow_html=True)
c1, c2, c3, c4, c5 = st.columns(5)
with c1:
    st.markdown(f'<div class="kpi-card"><div class="kpi-card-label">Total Revenue</div><div class="kpi-card-value">{fmt_currency(total_rev)}</div>{delta_html(curr_rev, prev_rev, "WoW")}<div class="kpi-card-sub">gross USD</div></div>', unsafe_allow_html=True)
with c2:
    st.markdown(f'<div class="kpi-card"><div class="kpi-card-label">Total Units</div><div class="kpi-card-value">{fmt_units(total_units)}</div>{delta_html(curr_units, prev_units, "WoW")}<div class="kpi-card-sub">quantity sold</div></div>', unsafe_allow_html=True)
with c3:
    st.markdown(f'<div class="kpi-card"><div class="kpi-card-label">Net Revenue</div><div class="kpi-card-value">{fmt_currency(total_net)}</div><div class="kpi-card-sub">after royalties</div></div>', unsafe_allow_html=True)
with c4:
    st.markdown(f'<div class="kpi-card"><div class="kpi-card-label">Avg Sale Price</div><div class="kpi-card-value">{fmt_currency(avg_price)}</div><div class="kpi-card-sub">revenue / unit</div></div>', unsafe_allow_html=True)
with c5:
    st.markdown(f'<div class="kpi-card"><div class="kpi-card-label">Direct Sale %</div><div class="kpi-card-value">{direct_pct:.1f}%</div><div class="kpi-card-sub">vs platform sales</div></div>', unsafe_allow_html=True)

# ---- CHART 1: Weekly revenue by game (full width) ----
st.markdown('<div class="section-label">Revenue Trends</div>', unsafe_allow_html=True)

weekly = df.groupby(['week', 'game'])['revenue'].sum().reset_index()
games  = sorted(weekly['game'].unique())
fig1   = go.Figure()
for i, g in enumerate(games):
    gd    = weekly[weekly['game'] == g].sort_values('week')
    short = g if len(g) <= 35 else g[:33] + '...'
    fig1.add_trace(go.Scatter(
        x=gd['week'], y=gd['revenue'], mode='lines', name=short,
        line=dict(color=PAL[i % len(PAL)], width=2),
        hovertemplate=f'<b>{g}</b><br>%{{x|%b %d, %Y}}<br>$%{{y:,.0f}}<extra></extra>',
    ))
fig1.update_layout(
    **base_layout(
        height=340, title="Weekly Revenue by Game", margin=dict(l=0, r=0, t=40, b=60),
        legend=dict(orientation='h', yanchor='top', y=-0.15, xanchor='left', x=0,
                    font=dict(size=10, color='#94a3b8'), bgcolor='rgba(0,0,0,0)'),
    ),
    xaxis=AX_DEFAULT,
    yaxis=dict(**AX, tickprefix='$'),
)
st.plotly_chart(fig1, width="stretch")

# ---- CHARTS ROW 2: Platform + Region ----
ch2, ch3 = st.columns(2)

with ch2:
    plat = df.groupby('platform')['revenue'].sum().reset_index().sort_values('revenue', ascending=False).head(8)
    fig2 = go.Figure(go.Bar(
        x=plat['platform'], y=plat['revenue'],
        marker_color=PAL[:len(plat)],
        text=[fmt_currency(v) for v in plat['revenue']],
        textposition='outside', textfont=dict(size=10, color='#94a3b8'),
    ))
    fig2.update_layout(
        **base_layout(height=300, title="Revenue by Platform", showlegend=False),
        xaxis=AX_LABEL,
        yaxis=dict(**AX, tickprefix='$', showticklabels=False),
    )
    st.plotly_chart(fig2, width="stretch")

with ch3:
    region = df.groupby('bp_region')['revenue'].sum().reset_index().sort_values('revenue', ascending=False)
    fig3   = go.Figure(go.Bar(
        x=region['bp_region'], y=region['revenue'],
        marker_color=PAL[2:2+len(region)],
        text=[fmt_currency(v) for v in region['revenue']],
        textposition='outside', textfont=dict(size=11, color='#94a3b8'),
    ))
    fig3.update_layout(
        **base_layout(height=300, title="Revenue by Region", showlegend=False),
        xaxis=AX_LABEL,
        yaxis=dict(**AX, tickprefix='$', showticklabels=False),
    )
    st.plotly_chart(fig3, width="stretch")

# ---- CHARTS ROW 3: Sale type + Top countries ----
ch4, ch5 = st.columns(2)

with ch4:
    sale_rev = df.groupby('sale_type')['revenue'].sum().sort_values(ascending=False).reset_index()
    if len(sale_rev) > 5:
        top5      = sale_rev.head(5)
        other_val = sale_rev.iloc[5:]['revenue'].sum()
        sale_rev  = pd.concat([top5, pd.DataFrame([{'sale_type': 'Other', 'revenue': other_val}])], ignore_index=True)
    sale_rev['label'] = sale_rev['sale_type'].apply(lambda s: s if len(s) <= 30 else s[:28] + '...')
    fig4 = go.Figure(go.Bar(
        x=sale_rev['revenue'], y=sale_rev['label'],
        orientation='h',
        marker_color=PAL[:len(sale_rev)],
        text=[fmt_currency(v) for v in sale_rev['revenue']],
        textposition='outside', textfont=dict(size=10, color='#94a3b8'),
    ))
    fig4.update_layout(
        **base_layout(height=300, title="Revenue by Sale Type", margin=dict(l=0, r=70, t=40, b=10), showlegend=False, hovermode='y unified'),
        xaxis=dict(**AX, tickprefix='$', showticklabels=False),
        yaxis=dict(**AX_LABEL, autorange='reversed'),
    )
    st.plotly_chart(fig4, width="stretch")

with ch5:
    top_c = df.groupby('country')['revenue'].sum().sort_values(ascending=False).head(10).reset_index()
    fig5  = go.Figure(go.Bar(
        x=top_c['revenue'], y=top_c['country'],
        orientation='h',
        marker_color=PAL[1],
        text=[fmt_currency(v) for v in top_c['revenue']],
        textposition='outside', textfont=dict(size=10, color='#94a3b8'),
    ))
    fig5.update_layout(
        **base_layout(height=300, title="Top 10 Countries by Revenue", margin=dict(l=0, r=70, t=40, b=10), showlegend=False, hovermode='y unified'),
        xaxis=dict(**AX, tickprefix='$', showticklabels=False),
        yaxis=dict(**AX_LABEL, autorange='reversed'),
    )
    st.plotly_chart(fig5, width="stretch")

# ---- CHART ROW 4: Units by platform (top 5) ----
st.markdown('<div class="section-label">Unit Volume</div>', unsafe_allow_html=True)

top_plats    = df.groupby('platform')['quantity'].sum().sort_values(ascending=False).head(5).index.tolist()
weekly_units = df[df['platform'].isin(top_plats)].groupby(['week', 'platform'])['quantity'].sum().reset_index()
fig6         = go.Figure()
for i, p in enumerate(top_plats):
    pd_ = weekly_units[weekly_units['platform'] == p].sort_values('week')
    fig6.add_trace(go.Bar(x=pd_['week'], y=pd_['quantity'], name=p, marker_color=PAL[i % len(PAL)]))
fig6.update_layout(
    **base_layout(height=300, title="Weekly Units by Platform (top 5)"),
    barmode='stack',
    xaxis=AX_DEFAULT,
    yaxis=AX_DEFAULT,
)
st.plotly_chart(fig6, width="stretch")

# ---- DATA TABLE ----
st.markdown('<div class="section-label">Data Table</div>', unsafe_allow_html=True)

display_cols = ['date', 'game', 'platform', 'country', 'bp_region', 'product_type',
                'sale_type', 'quantity', 'revenue', 'net_revenue_usd', 'pre_order']
display_cols = [c for c in display_cols if c in df.columns]
table_df     = df[display_cols].copy().reset_index(drop=True)
table_df['date'] = table_df['date'].dt.strftime('%Y-%m-%d')

sc, _, cc = st.columns([3, 3, 1])
with sc:
    search_term = st.text_input("Search", placeholder="Search rows...", label_visibility="hidden")
with cc:
    st.markdown(f"<div style='text-align:right;font-size:.72rem;color:#475569;padding-top:.6rem;'>{len(table_df):,} rows</div>", unsafe_allow_html=True)

if search_term:
    mask     = table_df.astype(str).apply(lambda c: c.str.contains(search_term, case=False, na=False)).any(axis=1)
    table_df = table_df[mask].reset_index(drop=True)

for c in ['revenue', 'net_revenue_usd']:
    if c in table_df.columns:
        table_df[c] = table_df[c].apply(lambda v: f"${v:,.2f}" if pd.notna(v) else "")

col_cfg = {}
if 'revenue'        in table_df.columns: col_cfg['revenue']        = st.column_config.TextColumn("Revenue (USD)")
if 'net_revenue_usd' in table_df.columns: col_cfg['net_revenue_usd'] = st.column_config.TextColumn("Net Revenue (USD)")
if 'quantity'       in table_df.columns: col_cfg['quantity']       = st.column_config.NumberColumn("Units", format="%d")
if 'sale_type'      in table_df.columns: col_cfg['sale_type']      = st.column_config.TextColumn("Sale Type")

st.dataframe(table_df, width="stretch", height=460, column_config=col_cfg, hide_index=True)

# ---- AI CHATBOT ----
st.markdown('<div class="section-label">AI Analyst</div>', unsafe_allow_html=True)

if "sales_chat_history" not in st.session_state:
    st.session_state.sales_chat_history = []

def build_data_context(df):
    df2 = df.copy()
    df2['week_str'] = df2['week'].dt.strftime('%Y-%m-%d')

    # WoW
    weekly = df2.groupby('week_str')[['revenue', 'quantity']].sum().sort_index()
    weeks_sorted = sorted(weekly.index.tolist())
    if len(weeks_sorted) >= 2:
        cur_w, prev_w = weeks_sorted[-1], weeks_sorted[-2]
        cr, pr = weekly.loc[cur_w, 'revenue'], weekly.loc[prev_w, 'revenue']
        cu, pu = weekly.loc[cur_w, 'quantity'], weekly.loc[prev_w, 'quantity']
        wow_text = (
            f"Latest week ({cur_w}): revenue=${cr:,.2f}, units={int(cu):,}. "
            f"Prior week ({prev_w}): revenue=${pr:,.2f}, units={int(pu):,}. "
            f"WoW revenue change: {(cr-pr)/abs(pr)*100:+.1f}%. "
            f"WoW unit change: {(cu-pu)/abs(pu)*100:+.1f}%."
        )
    else:
        wow_text = "Only one week of data — WoW not possible."

    # Pre-computed tables
    game_plat     = df2.groupby(['game','platform'])[['revenue','quantity']].sum().reset_index().to_string(index=False)
    game_sale     = df2.groupby(['game','sale_type'])[['revenue','quantity']].sum().reset_index().to_string(index=False)
    game_plat_sale= df2.groupby(['game','platform','sale_type'])[['revenue','quantity']].sum().reset_index().to_string(index=False)
    game_week     = df2.groupby(['game','week_str'])[['revenue','quantity']].sum().reset_index().to_string(index=False)
    sale_totals   = df2.groupby('sale_type')[['revenue','quantity']].sum().reset_index().to_string(index=False)
    top_countries = df2.groupby('country')[['revenue','quantity']].sum().sort_values('revenue',ascending=False).head(20).to_string()
    prod_type     = df2.groupby('product_type')[['revenue','quantity']].sum().reset_index().to_string(index=False)

    dlc_units  = df2[df2['product_type'].str.lower().str.contains('dlc|add', na=False)]['quantity'].sum() if 'product_type' in df2.columns else 0
    base_units = df2[df2['product_type'].str.lower().str.contains('base|game|standard', na=False)]['quantity'].sum() if 'product_type' in df2.columns else 0
    attach_str = (f"{dlc_units/base_units*100:.1f}% ({int(dlc_units):,} DLC / {int(base_units):,} base units)"
                  if base_units > 0 else f"No base game rows found. DLC/Add-on units: {int(dlc_units):,}")

    # Raw CSV (capped at 3000 rows)
    MAX_ROWS = 3000
    keep_cols = [c for c in ['date','game','platform','country','bp_region',
                              'product_type','sale_type','quantity','revenue','net_revenue_usd']
                 if c in df2.columns]
    raw = df2[keep_cols].copy()
    raw['date'] = raw['date'].dt.strftime('%Y-%m-%d')
    truncated = len(raw) > MAX_ROWS
    csv_str = raw.head(MAX_ROWS).to_csv(index=False)
    trunc_note = f"(first {MAX_ROWS} of {len(raw):,} rows shown)" if truncated else "(full dataset)"

    overview = (
        f"Dates: {df2['date'].min().strftime('%Y-%m-%d')} to {df2['date'].max().strftime('%Y-%m-%d')}\n"
        f"Rows: {len(df2):,} {trunc_note}\n"
        f"Gross revenue: {df2['revenue'].sum():,.2f} USD\n"
        f"Net revenue: {df2['net_revenue_usd'].sum():,.2f} USD\n"
        f"Units: {df2['quantity'].sum():,}\n"
        f"Games: {', '.join(sorted(df2['game'].unique()))}\n"
        f"Platforms: {', '.join(sorted(df2['platform'].unique()))}\n"
        f"Regions: {', '.join(sorted(df2['bp_region'].unique()))}"
    )

    return (
        "You are an expert sales data analyst with full transaction-level data.\n"
        "NEVER say a breakdown is unavailable. All dimensions (game, platform, sale_type, country) are present.\n"
        "Use the pre-computed tables for speed. Use the raw CSV for cross-dimensional queries.\n\n"
        "=== OVERVIEW ===\n" + overview + "\n\n"
        "=== WEEK-OVER-WEEK ===\n" + wow_text + "\n\n"
        "=== WEEKLY TOTALS ===\n" + weekly.to_string() + "\n\n"
        "=== BY GAME + PLATFORM ===\n" + game_plat + "\n\n"
        "=== BY GAME + SALE TYPE ===\n" + game_sale + "\n\n"
        "=== BY GAME + PLATFORM + SALE TYPE ===\n" + game_plat_sale + "\n\n"
        "=== BY GAME + WEEK ===\n" + game_week + "\n\n"
        "=== SALE TYPE TOTALS ===\n" + sale_totals + "\n\n"
        "=== TOP 20 COUNTRIES ===\n" + top_countries + "\n\n"
        "=== PRODUCT TYPE ===\n" + prod_type + "\n\n"
        "=== DLC ATTACH RATE ===\n" + attach_str + "\n\n"
        "=== RAW TRANSACTIONS (CSV) ===\n" + csv_str + "\n\n"
        "COLUMN NOTES:\n"
        "- sale_type: 'Direct Package Sale'=direct; 'Platform Sale'=storefront; other values=bundle/edition names.\n"
        "- platform: Steam, PS5, PS4, XBSX, XB1, NSW, NSW2, epic, etc.\n"
        "- For Steam-specific or platform-specific breakdowns, filter the raw CSV by platform.\n"
        "- Always give exact dollar amounts and percentages."
    )

if "chat_thinking" not in st.session_state:
    st.session_state.chat_thinking = False

chat_html = '<div class="chat-container">'
if not st.session_state.sales_chat_history:
    chat_html += '<div class="chat-thinking">Ask me anything about this sales data — WoW changes, platform comparisons, DLC attach rates, top countries, and more.</div>'
for msg in st.session_state.sales_chat_history:
    if msg['role'] == 'user':
        chat_html += f'<div class="chat-msg-user"><div class="chat-bubble-user">{msg["content"]}</div></div>'
    else:
        chat_html += f'<div class="chat-msg-ai"><div class="chat-bubble-ai">{msg["content"].replace(chr(10), "<br>")}</div></div>'
if st.session_state.chat_thinking:
    chat_html += '<div class="chat-msg-ai"><div class="chat-bubble-ai chat-thinking">Analysing your data...</div></div>'
chat_html += '</div>'
st.markdown(chat_html, unsafe_allow_html=True)

ic, bc, clc = st.columns([6, 1, 1])
with ic:
    user_input = st.text_input("Ask the AI", placeholder="e.g. What is the WoW revenue change? Which platform drives the most units?", label_visibility="hidden", key="sales_chat_input", disabled=st.session_state.chat_thinking)
with bc:
    send = st.button("Ask", width="stretch", type="primary", disabled=st.session_state.chat_thinking)
with clc:
    if st.button("Clear", width="stretch", disabled=st.session_state.chat_thinking):
        st.session_state.sales_chat_history = []
        st.rerun()

if send and user_input.strip() and not st.session_state.chat_thinking:
    st.session_state.sales_chat_history.append({"role": "user", "content": user_input.strip()})
    st.session_state.chat_thinking = True
    st.rerun()

if st.session_state.chat_thinking and st.session_state.sales_chat_history and st.session_state.sales_chat_history[-1]["role"] == "user":
    ctx      = build_data_context(df)
    messages = [{"role": "user" if m["role"] == "user" else "assistant", "content": m["content"]}
                for m in st.session_state.sales_chat_history[:-1]]
    messages.append({"role": "user", "content": f"{ctx}\n\nUSER QUESTION: {st.session_state.sales_chat_history[-1]['content']}"})
    with st.spinner("Analysing your data..."):
        try:
            import requests
            resp = requests.post(
                "https://api.anthropic.com/v1/messages",
                headers={"Content-Type": "application/json", "x-api-key": st.secrets["ANTHROPIC_API_KEY"], "anthropic-version": "2023-06-01"},
                json={"model": "claude-sonnet-4-20250514", "max_tokens": 1000,
                      "system": "You are a concise, accurate sales data analyst. Be direct and data-driven.",
                      "messages": messages},
                timeout=60,
            )
            resp.raise_for_status()
            answer = resp.json()["content"][0]["text"]
        except Exception as e:
            answer = f"Sorry, I couldn't reach the AI API: {e}"
    st.session_state.chat_thinking = False
    st.session_state.sales_chat_history.append({"role": "assistant", "content": answer})
    st.rerun()