"""
Submission Tracker — SEGA Platform Ops
=======================================
Run with:  streamlit run submission_tracker.py

Secrets (.streamlit/secrets.toml):
    AWS_ACCESS_KEY_ID          = "AKIA..."   # SES (OTP email)
    AWS_SECRET_ACCESS_KEY      = "..."
    AWS_SES_REGION             = "us-east-1"
    EMAIL_FROM                 = "noreply@segaamerica.com"
    AWS_ACCESS_KEY_ID_API      = "AKIA..."   # Bedrock (Claude AI)
    AWS_SECRET_ACCESS_KEY_API  = "..."
    AWS_BEDROCK_REGION         = "us-east-1"
    COOKIE_SIGNING_KEY         = "some-long-random-string"
"""

import base64
import hashlib
import hmac
import io
import random
import time

import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import streamlit as st

# ─────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="Submission Tracker",
    page_icon="🎮",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────────────────────
# SEGA BRAND STYLES
# ─────────────────────────────────────────────────────────────

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter+Tight:wght@400;700;800;900&family=Poppins:wght@300;400;500;600&display=swap');

:root {
    --bg:        #0a0c1a;  --surface:   #0f1120;  --surface2: #141728;
    --surface3:  #1a1e30;  --border:    #232640;  --border-hi:#323760;
    --blue:      #4080ff;  --blue-glow: rgba(64,128,255,0.16);
    --text:      #eef0fa;  --text-dim:  #b8bcd4;  --muted:    #5a5f82;
    --pos:       #20c65a;  --neg:       #ff3d52;  --amber:    #f0a500;
}
html, body { background: var(--bg) !important; color: var(--text) !important; }
.stApp, .stApp > div,
section[data-testid="stAppViewContainer"],
section[data-testid="stAppViewContainer"] > div,
div[data-testid="stMain"], div[data-testid="stVerticalBlock"],
div[data-testid="stHorizontalBlock"],
.main .block-container, .block-container {
    background-color: var(--bg) !important; color: var(--text) !important;
}
*, *::before, *::after { font-family: 'Poppins', sans-serif; box-sizing: border-box; }
p, span, div, li, td, th, label, h1, h2, h3, h4, h5, h6,
.stMarkdown, .stMarkdown p, .stMarkdown span,
[data-testid="stMarkdownContainer"], [data-testid="stMarkdownContainer"] p,
[data-testid="stMarkdownContainer"] span, [data-testid="stMarkdownContainer"] li,
[data-testid="stMarkdownContainer"] strong, [class*="css"] { color: var(--text) !important; }
.stCaption, [data-testid="stCaptionContainer"] p { color: var(--muted) !important; }
code { background: var(--surface3) !important; color: var(--blue) !important; padding:.1em .4em; border-radius:3px; }
#MainMenu, footer, header { visibility: hidden; }
.block-container { padding: 0 2.5rem 4rem !important; max-width: 1600px !important; }
::-webkit-scrollbar { width:5px; height:5px; }
::-webkit-scrollbar-thumb { background: var(--border-hi); border-radius:4px; }

/* TOP NAV */
.topbar {
    background:var(--surface); border-bottom:1px solid var(--border);
    padding:.8rem 2.5rem; margin:0 -2.5rem 1.75rem;
    display:flex; align-items:center; gap:1.25rem; position:relative;
}
.topbar::after {
    content:''; position:absolute; bottom:-1px; left:0; right:0; height:1px;
    background:linear-gradient(90deg, var(--blue) 0%, rgba(64,128,255,0) 55%);
}
.topbar-logo { font-family:'Inter Tight',sans-serif; font-size:.95rem; font-weight:900;
               color:var(--text) !important; letter-spacing:.12em; text-transform:uppercase; }
.topbar-logo .seg { color:var(--blue); }
.topbar-divider { width:1px; height:18px; background:var(--border-hi); flex-shrink:0; }
.topbar-label { font-size:.6rem; font-weight:600; color:var(--muted) !important;
                letter-spacing:.2em; text-transform:uppercase; }
.topbar-pill { margin-left:auto; background:var(--blue-glow); border:1px solid rgba(64,128,255,.28);
               border-radius:20px; padding:.18rem .7rem; font-size:.58rem; font-weight:700;
               letter-spacing:.14em; text-transform:uppercase; color:var(--blue) !important; }

/* AUTH */
.auth-wrap { max-width:420px; margin:5rem auto; padding:2.5rem 2.5rem 2rem;
             background:var(--surface); border:1px solid var(--border);
             border-top:3px solid var(--blue); border-radius:0 0 10px 10px; }
.auth-logo { font-family:'Inter Tight',sans-serif; font-size:1.6rem; font-weight:900;
             letter-spacing:.12em; color:var(--blue) !important; margin-bottom:.2rem; }
.auth-title { font-family:'Inter Tight',sans-serif; font-size:1rem; font-weight:700;
              color:var(--text) !important; margin-bottom:.25rem; }
.auth-sub { font-size:.8rem; color:var(--muted) !important; margin-bottom:1.5rem; }
.auth-note { font-size:.72rem; color:var(--muted) !important; margin-top:1rem;
             text-align:center; line-height:1.5; }

/* METRIC CARDS */
.metric-card { background:var(--surface); border:1px solid var(--border);
               border-radius:8px; padding:1rem 1.25rem; text-align:center; }
.metric-value { font-family:'Inter Tight',sans-serif; font-size:2rem; font-weight:800; }
.metric-label { font-size:.72rem; color:var(--muted); margin-top:4px;
                text-transform:uppercase; letter-spacing:.1em; }

/* SECTION HEADER */
.section-header { display:flex; align-items:center; gap:.55rem;
    font-family:'Inter Tight',sans-serif; font-size:.72rem; font-weight:800;
    letter-spacing:.18em; text-transform:uppercase; color:var(--text-dim) !important;
    margin:1.6rem 0 .9rem; }
.section-header .dot { width:6px; height:6px; border-radius:50%; background:var(--blue); flex-shrink:0; }

/* BUTTONS */
.stButton > button {
    background:var(--blue) !important; color:#fff !important; border:none !important;
    border-radius:6px !important; font-family:'Inter Tight',sans-serif !important;
    font-size:.78rem !important; font-weight:800 !important; letter-spacing:.12em !important;
    text-transform:uppercase !important; padding:.5rem 1.5rem !important;
    transition:background .15s, box-shadow .15s, transform .1s !important;
    box-shadow:0 2px 10px rgba(64,128,255,.3) !important;
}
.stButton > button:hover { background:#2d6aee !important; transform:translateY(-1px) !important; }
.stButton > button:disabled { background:var(--surface3) !important; color:var(--muted) !important; box-shadow:none !important; }

/* FILE UPLOADER */
[data-testid="stFileUploader"] { background:var(--surface2) !important;
    border:1px dashed var(--border-hi) !important; border-radius:8px !important; padding:.5rem !important; }
[data-testid="stFileUploader"]:hover { border-color:var(--blue) !important; }

/* SIDEBAR */
section[data-testid="stSidebar"], section[data-testid="stSidebar"] > div {
    background:var(--surface) !important; border-right:1px solid var(--border) !important; }

/* SELECT */
div[data-baseweb="select"] > div { background:var(--bg) !important; border-color:var(--border) !important; color:var(--text) !important; }
div[data-baseweb="menu"], div[data-baseweb="popover"] { background:var(--surface2) !important; border:1px solid var(--border-hi) !important; }
div[data-baseweb="menu"] li { color:var(--text) !important; }
div[data-baseweb="menu"] li:hover { background:var(--surface3) !important; }

/* TABS */
button[data-baseweb="tab"] { font-family:'Inter Tight',sans-serif !important; font-size:.72rem !important;
    font-weight:800 !important; letter-spacing:.14em !important; text-transform:uppercase !important;
    color:var(--muted) !important; background:transparent !important; }
button[data-baseweb="tab"][aria-selected="true"] { color:var(--blue) !important; }
div[data-baseweb="tab-highlight"] { background:var(--blue) !important; }
div[data-baseweb="tab-border"] { background:var(--border) !important; }

/* DATAFRAME */
[data-testid="stDataFrame"] { border:1px solid var(--border) !important; border-radius:8px !important; }

/* CHAT */
.stChatMessage { background:var(--surface2) !important; border:1px solid var(--border) !important; border-radius:8px !important; }

/* SUGGESTED CHIPS */
.chip-wrap { display:flex; flex-wrap:wrap; gap:.5rem; margin-bottom:1.25rem; }
.chip {
    background: var(--surface2);
    border: 1px solid var(--border-hi);
    border-radius: 20px;
    padding: .3rem .85rem;
    font-size: .72rem;
    color: var(--text-dim) !important;
    cursor: pointer;
    transition: all .15s;
    white-space: nowrap;
}
.chip:hover { border-color: var(--blue); color: var(--blue) !important; background: rgba(64,128,255,.08); }

/* FOOTER */
.footer { margin-top:4rem; padding:1.5rem 0; border-top:1px solid var(--border);
          display:flex; align-items:center; justify-content:space-between; }
.footer-brand { font-family:'Inter Tight',sans-serif; font-size:.72rem; font-weight:900;
                letter-spacing:.18em; color:var(--muted) !important; }
.footer-note { font-size:.65rem; color:var(--muted) !important; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────
# SECRETS
# ─────────────────────────────────────────────────────────────

claude_key = ""  # unused — auth is via AWS Bedrock IAM credentials

# ─────────────────────────────────────────────────────────────
# OTP / SES AUTH
# ─────────────────────────────────────────────────────────────

ALLOWED_DOMAIN    = "@segaamerica.com"
OTP_EXPIRY_SECS   = 600
TOKEN_EXPIRY_DAYS = 7


def _send_otp(email: str, code: str) -> bool:
    """Send a 6-digit OTP via AWS SES."""
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
                "Subject": {"Data": "Submission Tracker — Verification Code", "Charset": "UTF-8"},
                "Body": {
                    "Text": {
                        "Data": f"Your verification code is: {code}\n\nExpires in 10 minutes.",
                        "Charset": "UTF-8",
                    },
                    "Html": {
                        "Data": f"""
                        <div style="font-family:Arial,sans-serif;max-width:480px;margin:0 auto;padding:32px 24px;background:#f8f9ff;">
                          <div style="font-size:22px;font-weight:900;letter-spacing:.1em;color:#4080ff;margin-bottom:4px;">SEGA</div>
                          <div style="font-size:14px;color:#444;margin-bottom:28px;">Submission Tracker</div>
                          <div style="font-size:14px;color:#222;margin-bottom:16px;">Your verification code is:</div>
                          <div style="font-size:42px;font-weight:900;letter-spacing:.18em;color:#1a1a2e;
                                      background:#e8eeff;border-radius:8px;padding:18px 24px;
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


# ── Token helpers (URL query param approach — no cookie library) ──
def _make_token(email: str) -> str:
    """Create a signed URL token: base64(email|expiry|hmac)."""
    secret  = st.secrets.get("COOKIE_SIGNING_KEY", "fallback-change-this")
    expiry  = int(time.time()) + (TOKEN_EXPIRY_DAYS * 86400)
    payload = f"{email}|{expiry}"
    sig     = hmac.new(secret.encode(), payload.encode(), hashlib.sha256).hexdigest()
    return base64.urlsafe_b64encode(f"{payload}|{sig}".encode()).decode()


def _verify_token(token: str) -> str | None:
    """Verify token. Returns email if valid, None otherwise."""
    try:
        secret   = st.secrets.get("COOKIE_SIGNING_KEY", "fallback-change-this")
        decoded  = base64.urlsafe_b64decode(token.encode()).decode()
        email, expiry_str, sig = decoded.rsplit("|", 2)
        payload  = f"{email}|{expiry_str}"
        expected = hmac.new(secret.encode(), payload.encode(), hashlib.sha256).hexdigest()
        if not hmac.compare_digest(sig, expected):
            return None
        if int(time.time()) > int(expiry_str):
            return None
        return email
    except Exception:
        return None


# ── Read token from URL query param ──────────────────────────
_url_token   = st.query_params.get("t", "")
_token_email = _verify_token(_url_token) if _url_token else None

# ── Session state defaults ────────────────────────────────────
for k, v in {
    "auth_verified": False, "auth_email": "", "auth_token": "",
    "otp_code": "", "otp_email": "", "otp_expiry": 0,
    "otp_sent": False, "otp_attempts": 0,
    "chat_history": [], "chat_pending": False,
}.items():
    if k not in st.session_state:
        st.session_state[k] = v

# If valid token in URL, mark as verified
if _token_email and not st.session_state.auth_verified:
    st.session_state.auth_verified = True
    st.session_state.auth_email    = _token_email
    st.session_state.auth_token    = _url_token

# ─────────────────────────────────────────────────────────────
# LOGIN GATE
# ─────────────────────────────────────────────────────────────

if not st.session_state.auth_verified:
    # Hide sidebar on login screen
    st.markdown("<style>section[data-testid='stSidebar']{display:none!important;}</style>",
                unsafe_allow_html=True)

    _, mid, _ = st.columns([1, 2, 1])
    with mid:
        st.markdown("""
        <div class="auth-wrap">
          <div class="auth-logo">SEGA</div>
          <div class="auth-title">Submission Tracker</div>
          <div class="auth-sub">Sign in with your SEGA America email to continue</div>
        </div>""", unsafe_allow_html=True)

        if not st.session_state.otp_sent:
            email_in = st.text_input("Email address", placeholder="you@segaamerica.com",
                                     label_visibility="collapsed", key="auth_email_input")
            if st.button("Send verification code", key="btn_send"):
                addr = email_in.strip().lower()
                if not addr.endswith(ALLOWED_DOMAIN):
                    st.error(f"Access restricted to {ALLOWED_DOMAIN} addresses.")
                else:
                    code = str(random.randint(100000, 999999))
                    if _send_otp(addr, code):
                        st.session_state.otp_code     = code
                        st.session_state.otp_email    = addr
                        st.session_state.otp_expiry   = time.time() + OTP_EXPIRY_SECS
                        st.session_state.otp_sent     = True
                        st.session_state.otp_attempts = 0
                        st.rerun()
        else:
            st.info(f"Code sent to **{st.session_state.otp_email}** — check your inbox.")
            code_in = st.text_input("6-digit code", placeholder="123456",
                                    label_visibility="collapsed", max_chars=6, key="auth_code_input")
            if st.button("Verify code", key="btn_verify"):
                if st.session_state.otp_attempts >= 5:
                    st.error("Too many attempts. Please request a new code.")
                    st.session_state.otp_sent = False
                elif time.time() > st.session_state.otp_expiry:
                    st.error("Code has expired. Please request a new one.")
                    st.session_state.otp_sent = False
                elif code_in.strip() != st.session_state.otp_code:
                    st.session_state.otp_attempts += 1
                    rem = 5 - st.session_state.otp_attempts
                    st.error(f"Incorrect code. {rem} attempt{'s' if rem != 1 else ''} remaining.")
                else:
                    # Success — generate token and inject into URL
                    st.session_state.auth_verified = True
                    st.session_state.auth_email    = st.session_state.otp_email
                    st.session_state.otp_code      = ""
                    _token = _make_token(st.session_state.auth_email)
                    st.session_state.auth_token    = _token
                    st.query_params["t"] = _token
                    st.rerun()

            if st.button("← Use a different email", key="btn_back"):
                st.session_state.otp_sent = False
                st.session_state.otp_code = ""
                st.rerun()

        st.markdown(
            f'<div class="auth-note">Restricted to {ALLOWED_DOMAIN} addresses only.<br>'
            f'Codes expire after 10 minutes.</div>',
            unsafe_allow_html=True,
        )
    st.stop()

# ─────────────────────────────────────────────────────────────
# SIDEBAR — user info, sign out, file upload, filters
# ─────────────────────────────────────────────────────────────

with st.sidebar:
    st.markdown(
        f'<div style="font-size:.7rem;font-weight:600;color:var(--muted);margin-bottom:.25rem;">'
        f'Signed in as<br>'
        f'<span style="color:var(--text);font-weight:700;">{st.session_state.auth_email}</span>'
        f'</div>', unsafe_allow_html=True,
    )
    if st.button("Sign out", key="sign_out"):
        st.query_params.clear()
        for k in ["auth_verified", "auth_email", "auth_token", "otp_sent",
                  "otp_code", "otp_email", "otp_expiry", "otp_attempts"]:
            st.session_state[k] = False if k == "auth_verified" else ""
        st.rerun()

    st.markdown("---")
    st.markdown("### 📁 Fiscal Year")
    fy_choice = st.radio(
        "Select dataset",
        ["FY26", "FY27"],
        horizontal=True,
        label_visibility="collapsed",
        key="fy_choice",
    )
    st.markdown("---")

# ─────────────────────────────────────────────────────────────
# TOP NAV BAR
# ─────────────────────────────────────────────────────────────

st.markdown("""
<div class="topbar">
  <div class="topbar-logo"><span class="seg">SEGA</span> SUBMISSION TRACKER</div>
  <div class="topbar-divider"></div>
  <div class="topbar-label">Platform Operations</div>
  <div class="topbar-pill">Internal</div>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────
# CONSTANTS
# ─────────────────────────────────────────────────────────────

RESULT_COLORS  = {"PASS": "#20c65a", "FAIL": "#ff3d52", "PRE": "#aaaaaa", "PENDING": "#f0a500"}
PRODUCT_COLORS = {"Demo": "#4c9be8", "Dlc": "#6a0dad", "Game": "#e07030", "Patch": "#3333cc"}
QUARTER_MAP    = {
    "Q1 (Apr–Jun)": [4,5,6],   "Q2 (Jul–Sep)": [7,8,9],
    "Q3 (Oct–Dec)": [10,11,12], "Q4 (Jan–Mar)": [1,2,3],
    "All": list(range(1,13)),
}

# ─────────────────────────────────────────────────────────────
# DATA LOADING
# ─────────────────────────────────────────────────────────────

@st.cache_data(show_spinner=False)
def load_data(file_bytes: bytes, fname: str) -> pd.DataFrame:
    if fname.endswith(".csv"):
        # Try common encodings in order — Excel exports are often latin-1 / cp1252
        for enc in ("utf-8", "utf-8-sig", "latin-1", "cp1252"):
            try:
                df = pd.read_csv(io.BytesIO(file_bytes), encoding=enc)
                break
            except (UnicodeDecodeError, Exception):
                continue
        else:
            # Last-resort: replace undecodable bytes
            df = pd.read_csv(io.BytesIO(file_bytes), encoding="utf-8", errors="replace")
    else:
        df = pd.read_excel(io.BytesIO(file_bytes))
    df.columns = df.columns.str.strip()
    if "Result" in df.columns:
        df["Result"] = (df["Result"].astype(str).str.strip().str.upper()
                         .replace({"NAN": "PENDING", "": "PENDING"}))
    for col in ["Sub Date", "Result Date"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")
    if "Sub Date" in df.columns:
        df["Month"]    = df["Sub Date"].dt.strftime("%B")
        df["MonthNum"] = df["Sub Date"].dt.month
    if "Product" in df.columns:
        df["Product"] = df["Product"].astype(str).str.strip().str.title()
    return df


def opts(series) -> list:
    return ["All"] + sorted(series.dropna().astype(str).unique())


# ── Local CSV paths (relative to repo root) ───────────────────
DATA_FILES = {
    "FY26": "data/fy26.csv",
    "FY27": "data/fy27.csv",
}

@st.cache_data(show_spinner=False)
def load_local(path: str) -> pd.DataFrame:
    """Read a bundled repo CSV and process it through the standard pipeline."""
    with open(path, "rb") as f:
        raw = f.read()
    return load_data(raw, path)


# ─────────────────────────────────────────────────────────────
# TABS
# ─────────────────────────────────────────────────────────────

tab_dash, tab_chat = st.tabs(["📊  Dashboard", "💬  Ask Claude"])

# ══════════════════════════════════════════════════════════════
# TAB 1 — DASHBOARD
# ══════════════════════════════════════════════════════════════

with tab_dash:
    # fy_choice is set by the sidebar radio above; default FY26
    fy        = fy_choice          # already a string: "FY26" or "FY27"
    data_path = DATA_FILES[fy]

    try:
        df_raw = load_local(data_path)
    except FileNotFoundError:
        st.error(
            f"Data file not found: `{data_path}`\n\n"
            f"Make sure `{data_path}` exists in your repository."
        )
        st.stop()
    except Exception as e:
        st.error(f"Failed to load {data_path}: {e}")
        st.stop()

    # ── Sidebar filters ───────────────────────────────────────
    with st.sidebar:
        st.markdown("### Filters")
        def sb(label, col):
            o = opts(df_raw[col]) if col in df_raw.columns else ["All"]
            return st.selectbox(label, o, key=f"{fy}_f_{col.replace(' ','_')}")

        f_codename  = sb("Codename",  "Codename")
        f_party     = sb("1st Party", "1st Party")
        f_platform  = sb("Platform",  "Platform")
        f_product   = sb("Product",   "Product")
        f_submitter = sb("Submitter", "Submitter")
        f_month     = sb("Month",     "Month")
        f_quarter   = st.selectbox("Quarter", list(QUARTER_MAP.keys()), key=f"{fy}_f_quarter")

    # ── Apply filters ─────────────────────────────────────────
    df = df_raw.copy()
    for col, val in [
        ("Codename", f_codename), ("1st Party", f_party), ("Platform", f_platform),
        ("Product", f_product), ("Submitter", f_submitter), ("Month", f_month),
    ]:
        if val != "All" and col in df.columns:
            df = df[df[col].astype(str) == val]
    if "MonthNum" in df.columns:
        df = df[df["MonthNum"].isin(QUARTER_MAP[f_quarter])]

    # ── Metric cards ──────────────────────────────────────────
    st.markdown(f'<div class="section-header"><span class="dot"></span>OVERVIEW · {fy}</div>',
                unsafe_allow_html=True)
    total   = len(df)
    passed  = int((df["Result"] == "PASS").sum())  if "Result" in df.columns else 0
    failed  = int((df["Result"] == "FAIL").sum())  if "Result" in df.columns else 0
    pending = int((df["Result"] == "PENDING").sum()) if "Result" in df.columns else 0
    avg_d   = df["Days in Test"].mean() if "Days in Test" in df.columns else None
    pass_rt = f"{passed/total*100:.1f}%" if total else "—"

    for col, (label, val, color) in zip(st.columns(6), [
        ("Total",     total,   "#eef0fa"),
        ("Pass",      passed,  RESULT_COLORS["PASS"]),
        ("Fail",      failed,  RESULT_COLORS["FAIL"]),
        ("Pending",   pending, RESULT_COLORS["PENDING"]),
        ("Pass Rate", pass_rt, RESULT_COLORS["PASS"]),
        ("Avg Days",  f"{avg_d:.1f}" if avg_d else "—", "#4080ff"),
    ]):
        with col:
            st.markdown(f"""<div class="metric-card">
              <div class="metric-value" style="color:{color}">{val}</div>
              <div class="metric-label">{label}</div></div>""", unsafe_allow_html=True)

    # ── Donut + Monthly stacked bar ───────────────────────────
    st.markdown('<div class="section-header"><span class="dot"></span>PASS / FAIL BREAKDOWN</div>',
                unsafe_allow_html=True)
    c1, c2 = st.columns(2)

    with c1:
        if "Result" in df.columns and not df.empty:
            counts = df["Result"].value_counts()
            labels = counts.index.tolist()
            values = counts.values.tolist()
            colors = [RESULT_COLORS.get(r, "#888") for r in labels]
            fig = go.Figure(go.Pie(
                labels=labels, values=values, hole=.56, marker_colors=colors,
                textposition="outside",
                texttemplate="%{value} (%{percent:.2f})",
                outsidetextfont=dict(size=11, color="#eef0fa"),
                hovertemplate="%{label}: %{value}<extra></extra>",
            ))
            fig.update_layout(
                paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                legend=dict(orientation="h", x=0, y=1.12, font=dict(color="#eef0fa", size=12)),
                margin=dict(t=50, b=20, l=20, r=20), height=360,
                annotations=[dict(text=f"<b>{total}</b><br><span style='font-size:11px'>Total</span>",
                                  x=.5, y=.5, font=dict(size=16, color="#eef0fa"), showarrow=False)],
            )
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("No result data to display.")

    with c2:
        if "Month" in df.columns and "Product" in df.columns and not df.empty:
            pivot = (df.groupby(["Month","MonthNum","Product"])
                       .size().reset_index(name="Count").sort_values("MonthNum"))
            months   = (pivot[["Month","MonthNum"]].drop_duplicates()
                        .sort_values("MonthNum")["Month"].tolist())
            products = sorted(df["Product"].dropna().unique())
            fig2 = go.Figure()
            for prod in products:
                sub = pivot[pivot["Product"] == prod]
                m2c = dict(zip(sub["Month"], sub["Count"]))
                ys  = [m2c.get(m, 0) for m in months]
                fig2.add_trace(go.Bar(
                    name=prod, x=months, y=ys,
                    marker_color=PRODUCT_COLORS.get(prod, "#888"),
                    text=[v if v else "" for v in ys],
                    textposition="inside", insidetextanchor="middle",
                    textfont=dict(color="white", size=11),
                ))
            totals = pivot.groupby("Month")["Count"].sum()
            for m in months:
                fig2.add_annotation(x=m, y=int(totals.get(m, 0)),
                                    text=f"<b>{int(totals.get(m,0))}</b>",
                                    showarrow=False, yshift=10,
                                    font=dict(size=12, color="#eef0fa"))
            fig2.update_layout(
                barmode="stack",
                paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                legend=dict(orientation="h", x=0, y=1.12, font=dict(color="#eef0fa", size=12)),
                xaxis=dict(title="Month", tickfont=dict(color="#b8bcd4"), gridcolor="#232640"),
                yaxis=dict(title="Submissions", tickfont=dict(color="#b8bcd4"), gridcolor="#232640"),
                margin=dict(t=50, b=40, l=20, r=20), height=360,
            )
            st.plotly_chart(fig2, use_container_width=True)
        else:
            st.info("No monthly data to display.")

    # ── Fail reasons + Submission log ─────────────────────────
    st.markdown('<div class="section-header"><span class="dot"></span>DETAIL</div>',
                unsafe_allow_html=True)
    c3, c4 = st.columns([1, 2])

    with c3:
        st.markdown("**Top Fail Reasons**")
        if "Fail Reason" in df.columns:
            fails = df[df["Result"] == "FAIL"]["Fail Reason"].dropna()
            if not fails.empty:
                fc   = fails.value_counts().head(10)
                fig3 = px.bar(x=fc.values, y=fc.index, orientation="h",
                              color_discrete_sequence=[RESULT_COLORS["FAIL"]])
                fig3.update_layout(
                    paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                    xaxis=dict(tickfont=dict(color="#b8bcd4"), gridcolor="#232640"),
                    yaxis=dict(tickfont=dict(color="#b8bcd4"), autorange="reversed"),
                    margin=dict(t=10, b=20, l=10, r=10), height=300,
                )
                st.plotly_chart(fig3, use_container_width=True)
            else:
                st.info("No failures in current filter.")
        else:
            st.info("No 'Fail Reason' column found.")

    with c4:
        st.markdown("**Submission Log**")
        display_cols = [c for c in [
            "Codename","Product","Sub #","Ver #","1st Party","Platform",
            "Submitter","Sub Date","Result","Days in Test","Fail Reason",
        ] if c in df.columns]
        sort_col = "Sub Date" if "Sub Date" in df.columns else display_cols[0]
        st.dataframe(
            df[display_cols].sort_values(sort_col, ascending=False).reset_index(drop=True),
            use_container_width=True, height=300,
        )

    # ── Download ──────────────────────────────────────────────
    st.markdown("---")
    buf = io.BytesIO()
    df.to_excel(buf, index=False)
    st.download_button(
        "⬇️ Download filtered data (.xlsx)",
        data=buf.getvalue(),
        file_name=f"{fy.lower()}_filtered_submissions.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# ══════════════════════════════════════════════════════════════
# TAB 2 — CLAUDE CHAT
# ══════════════════════════════════════════════════════════════

with tab_chat:
    if not st.secrets.get("AWS_ACCESS_KEY_ID_API", ""):
        st.warning("Add `AWS_ACCESS_KEY_ID_API`, `AWS_SECRET_ACCESS_KEY_API`, and `AWS_BEDROCK_REGION` to `.streamlit/secrets.toml` to enable the AI assistant.")
        st.stop()

    st.markdown("""
    <div style="margin-bottom:1rem;">
      <div style="font-family:'Inter Tight',sans-serif;font-size:1.1rem;font-weight:800;color:var(--text);">
        Ask Claude about your submissions
      </div>
      <div style="font-size:.82rem;color:var(--muted);margin-top:.25rem;">
        Claude can answer questions, spot trends, surface risks, or summarise results
        based on the filtered data in the Dashboard tab.
      </div>
    </div>""", unsafe_allow_html=True)

    SUGGESTED_PROMPTS = [
        f"What is the overall pass rate for {fy}?",
        "Which titles have the most failures?",
        "Show me all FAILs and their reasons",
        "Which submitters have the highest pass rate?",
        "Which platforms have the most submissions?",
        "List titles that are still PENDING",
        "What are the most common fail reasons?",
        "How many submissions were there per month?",
    ]

    st.markdown("**Suggested questions**")
    chip_cols = st.columns(4)
    for i, prompt in enumerate(SUGGESTED_PROMPTS):
        with chip_cols[i % 4]:
            if st.button(prompt, key=f"chip_{i}"):
                st.session_state.chat_history.append({"role": "user", "content": prompt})
                st.session_state.chat_pending = True
                st.rerun()

    st.markdown("<div style='margin-top:.5rem'></div>", unsafe_allow_html=True)

    # Build data context for Claude
    def _data_context() -> str:
        try:
            snap = df.head(300).to_csv(index=False)
            return (
                f"{fy} Submission tracker — {len(df)} rows after current filters.\n\n"
                f"CSV sample (up to 300 rows):\n{snap}"
            )
        except Exception:
            return "Data available but could not be serialised."

    SYSTEM = f"""You are an internal data analyst for SEGA America's Platform Operations team.
You help the team understand their {fy} 1st-party submission tracker.

{_data_context()}

Answer questions about pass/fail rates, trends, specific titles, submitters, platforms,
fail reasons, turnaround times, or anything else the user asks.
Be concise. Use markdown tables or bullet lists where helpful.
If you cannot answer from the available data, say so clearly."""

    # Render history
    for msg in st.session_state.chat_history:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    # Stream response for pending message
    if st.session_state.chat_pending:
        st.session_state.chat_pending = False
        try:
            import anthropic
            client = anthropic.AnthropicBedrock(
                aws_access_key=st.secrets.get("AWS_ACCESS_KEY_ID_API", ""),
                aws_secret_key=st.secrets.get("AWS_SECRET_ACCESS_KEY_API", ""),
                aws_region=st.secrets.get("AWS_BEDROCK_REGION", "us-east-1"),
            )
            api_msgs = [{"role": m["role"], "content": m["content"]}
                        for m in st.session_state.chat_history]
            with st.chat_message("assistant"):
                reply = ""
                ph    = st.empty()
                with client.messages.stream(
                    model="us.anthropic.claude-sonnet-4-6",
                    max_tokens=1024,
                    system=SYSTEM,
                    messages=api_msgs,
                ) as stream:
                    for delta in stream.text_stream:
                        reply += delta
                        ph.markdown(reply + "▌")
                ph.markdown(reply)
            st.session_state.chat_history.append({"role": "assistant", "content": reply})
        except Exception as e:
            st.error(f"Claude error: {type(e).__name__}: {e}")

    user_msg = st.chat_input("Ask about your submission data…")
    if user_msg:
        st.session_state.chat_history.append({"role": "user", "content": user_msg})
        st.session_state.chat_pending = True
        st.rerun()

    if st.session_state.chat_history:
        if st.button("Clear conversation", key="clear_chat"):
            st.session_state.chat_history = []
            st.session_state.chat_pending = False
            st.rerun()

# ─────────────────────────────────────────────────────────────
# FOOTER
# ─────────────────────────────────────────────────────────────

st.markdown("""
<div class="footer">
  <div class="footer-brand">SEGA SUBMISSION TRACKER</div>
  <div class="footer-note">Powered by Claude · Data processed locally · Internal use only</div>
</div>
""", unsafe_allow_html=True)