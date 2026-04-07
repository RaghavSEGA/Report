"""
Ratings Tracker — AI Assistant
================================
Run with:  streamlit run ratings_ai.py

Secrets (.streamlit/secrets.toml):
    ANTHROPIC_API_KEY      = "sk-ant-..."
    AWS_ACCESS_KEY_ID      = "AKIA..."
    AWS_SECRET_ACCESS_KEY  = "..."
    AWS_SES_REGION         = "us-east-1"
    EMAIL_FROM             = "noreply@segaamerica.com"
    COOKIE_SIGNING_KEY     = "some-long-random-string"
"""

import base64
import hashlib
import hmac
import io
import random
import time

import pandas as pd
import streamlit as st

# ─────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="Ratings Tracker AI",
    page_icon="🎮",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────────────────────
# STYLES
# ─────────────────────────────────────────────────────────────

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter+Tight:wght@400;600;700;800;900&family=Poppins:wght@300;400;500&display=swap');

:root {
    --bg:        #08091a;
    --surface:   #0d0f22;
    --surface2:  #111328;
    --surface3:  #171a2f;
    --border:    #1e2140;
    --border-hi: #2a2e52;
    --blue:      #4080ff;
    --blue-dim:  rgba(64,128,255,0.12);
    --blue-glow: rgba(64,128,255,0.25);
    --text:      #eef0fa;
    --text-dim:  #9ba3c4;
    --muted:     #4a5070;
    --pass:      #20c65a;
    --fail:      #ff3d52;
    --amber:     #f0a500;
    --purple:    #9b59f5;
}

html, body { background: var(--bg) !important; color: var(--text) !important; }
.stApp, .stApp > div,
section[data-testid="stAppViewContainer"],
section[data-testid="stAppViewContainer"] > div,
div[data-testid="stMain"],
div[data-testid="stVerticalBlock"],
div[data-testid="stHorizontalBlock"],
.main .block-container, .block-container {
    background-color: var(--bg) !important;
    color: var(--text) !important;
}

*, *::before, *::after { font-family: 'Poppins', sans-serif; box-sizing: border-box; }

p, span, div, li, td, th, label, h1, h2, h3, h4,
.stMarkdown, .stMarkdown p, .stMarkdown span,
[data-testid="stMarkdownContainer"],
[data-testid="stMarkdownContainer"] p,
[data-testid="stMarkdownContainer"] span,
[data-testid="stMarkdownContainer"] li,
[data-testid="stMarkdownContainer"] strong,
[class*="css"] { color: var(--text) !important; }

.stCaption, [data-testid="stCaptionContainer"] p { color: var(--muted) !important; }

code {
    background: var(--surface3) !important;
    color: var(--blue) !important;
    padding: .15em .45em;
    border-radius: 4px;
    font-size: .85em;
}

#MainMenu, footer, header { visibility: hidden; }
.block-container { padding: 0 2rem 4rem !important; max-width: 1400px !important; }

::-webkit-scrollbar { width: 4px; height: 4px; }
::-webkit-scrollbar-thumb { background: var(--border-hi); border-radius: 4px; }

/* ── AUTH ────────────────────────────────────── */
.auth-wrap {
    max-width: 420px; margin: 5rem auto; padding: 2.5rem 2.5rem 2rem;
    background: var(--surface); border: 1px solid var(--border);
    border-top: 3px solid var(--blue); border-radius: 0 0 10px 10px;
}
.auth-logo { font-family:'Inter Tight',sans-serif; font-size:1.6rem; font-weight:900;
             letter-spacing:.12em; color:var(--blue) !important; margin-bottom:.2rem; }
.auth-title { font-family:'Inter Tight',sans-serif; font-size:1rem; font-weight:700;
              color:var(--text) !important; margin-bottom:.25rem; }
.auth-sub { font-size:.8rem; color:var(--muted) !important; margin-bottom:1.5rem; }
.auth-note { font-size:.72rem; color:var(--muted) !important; margin-top:1rem;
             text-align:center; line-height:1.5; }

/* ── TOP NAV ─────────────────────────────── */
.topbar {
    background: var(--surface);
    border-bottom: 1px solid var(--border);
    padding: .75rem 2rem;
    margin: 0 -2rem 2rem;
    display: flex;
    align-items: center;
    gap: 1rem;
    position: relative;
}
.topbar::after {
    content: '';
    position: absolute;
    bottom: -1px; left: 0; right: 0; height: 1px;
    background: linear-gradient(90deg, var(--blue) 0%, transparent 60%);
}
.topbar-brand {
    font-family: 'Inter Tight', sans-serif;
    font-weight: 900;
    font-size: .9rem;
    letter-spacing: .14em;
    text-transform: uppercase;
    color: var(--text) !important;
}
.topbar-brand .hi { color: var(--blue); }
.topbar-sep { width: 1px; height: 16px; background: var(--border-hi); }
.topbar-sub {
    font-size: .58rem;
    font-weight: 600;
    letter-spacing: .2em;
    text-transform: uppercase;
    color: var(--muted) !important;
}
.topbar-badge {
    margin-left: auto;
    background: var(--blue-dim);
    border: 1px solid rgba(64,128,255,.3);
    border-radius: 20px;
    padding: .15rem .65rem;
    font-size: .56rem;
    font-weight: 700;
    letter-spacing: .14em;
    text-transform: uppercase;
    color: var(--blue) !important;
}

/* ── SIDEBAR ─────────────────────────────── */
section[data-testid="stSidebar"],
section[data-testid="stSidebar"] > div {
    background: var(--surface) !important;
    border-right: 1px solid var(--border) !important;
}

/* ── FILE UPLOADER ───────────────────────── */
[data-testid="stFileUploader"] {
    background: var(--surface2) !important;
    border: 1px dashed var(--border-hi) !important;
    border-radius: 10px !important;
    transition: border-color .2s;
}
[data-testid="stFileUploader"]:hover { border-color: var(--blue) !important; }
[data-testid="stFileUploaderDropzoneInstructions"] span,
[data-testid="stFileUploaderDropzoneInstructions"] p { color: var(--muted) !important; }

/* ── BUTTONS ─────────────────────────────── */
.stButton > button {
    background: var(--blue) !important;
    color: #fff !important;
    border: none !important;
    border-radius: 6px !important;
    font-family: 'Inter Tight', sans-serif !important;
    font-size: .75rem !important;
    font-weight: 800 !important;
    letter-spacing: .12em !important;
    text-transform: uppercase !important;
    padding: .45rem 1.25rem !important;
    box-shadow: 0 2px 12px var(--blue-glow) !important;
    transition: all .15s !important;
}
.stButton > button:hover {
    background: #2d6aee !important;
    box-shadow: 0 4px 20px var(--blue-glow) !important;
    transform: translateY(-1px) !important;
}
.stButton > button:disabled {
    background: var(--surface3) !important;
    color: var(--muted) !important;
    box-shadow: none !important;
    transform: none !important;
}

/* ── SELECT ──────────────────────────────── */
div[data-baseweb="select"] > div {
    background: var(--surface2) !important;
    border-color: var(--border) !important;
    color: var(--text) !important;
}
div[data-baseweb="menu"], div[data-baseweb="popover"] {
    background: var(--surface2) !important;
    border: 1px solid var(--border-hi) !important;
    box-shadow: 0 8px 32px rgba(0,0,0,.6) !important;
}
div[data-baseweb="menu"] li { color: var(--text) !important; }
div[data-baseweb="menu"] li:hover { background: var(--surface3) !important; }

/* ── DATAFRAME ───────────────────────────── */
[data-testid="stDataFrame"] {
    border: 1px solid var(--border) !important;
    border-radius: 10px !important;
    overflow: hidden;
}

/* ── CHAT ────────────────────────────────── */
.stChatMessage {
    background: var(--surface2) !important;
    border: 1px solid var(--border) !important;
    border-radius: 10px !important;
    margin-bottom: .5rem !important;
}
[data-testid="stChatMessageContent"] p { color: var(--text) !important; }

/* ── SECTION LABEL ───────────────────────── */
.sec {
    display: flex;
    align-items: center;
    gap: .5rem;
    font-family: 'Inter Tight', sans-serif;
    font-size: .68rem;
    font-weight: 800;
    letter-spacing: .2em;
    text-transform: uppercase;
    color: var(--text-dim) !important;
    margin: 1.5rem 0 .75rem;
}
.sec::before {
    content: '';
    width: 6px; height: 6px;
    border-radius: 50%;
    background: var(--blue);
    flex-shrink: 0;
}

/* ── STAT CARDS ──────────────────────────── */
.stat-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(130px, 1fr));
    gap: .75rem;
    margin-bottom: 1.5rem;
}
.stat-card {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 10px;
    padding: .9rem 1rem;
    text-align: center;
    position: relative;
    overflow: hidden;
}
.stat-card::before {
    content: '';
    position: absolute;
    top: 0; left: 0; right: 0;
    height: 2px;
    background: var(--accent-color, var(--blue));
}
.stat-val {
    font-family: 'Inter Tight', sans-serif;
    font-size: 1.8rem;
    font-weight: 900;
    color: var(--accent-color, var(--blue)) !important;
    line-height: 1;
}
.stat-lbl {
    font-size: .65rem;
    font-weight: 600;
    letter-spacing: .1em;
    text-transform: uppercase;
    color: var(--muted) !important;
    margin-top: .35rem;
}

/* ── RATING PILL ─────────────────────────── */
.pill {
    display: inline-block;
    padding: .15rem .55rem;
    border-radius: 4px;
    font-size: .7rem;
    font-weight: 700;
    letter-spacing: .05em;
}
.pill-rated   { background: rgba(32,198,90,.15);  color: #20c65a !important; border: 1px solid rgba(32,198,90,.3); }
.pill-missing { background: rgba(255,61,82,.15);   color: #ff3d52 !important; border: 1px solid rgba(255,61,82,.3); }
.pill-pending { background: rgba(240,165,0,.15);   color: #f0a500 !important; border: 1px solid rgba(240,165,0,.3); }
.pill-na      { background: rgba(74,80,112,.2);    color: #9ba3c4 !important; border: 1px solid var(--border-hi); }

/* ── PROMPT CHIPS ────────────────────────── */
.chip-row { display: flex; flex-wrap: wrap; gap: .5rem; margin-bottom: 1rem; }
.chip {
    background: var(--surface2);
    border: 1px solid var(--border-hi);
    border-radius: 20px;
    padding: .3rem .8rem;
    font-size: .72rem;
    color: var(--text-dim) !important;
    cursor: pointer;
    transition: all .15s;
    white-space: nowrap;
}
.chip:hover {
    border-color: var(--blue);
    color: var(--blue) !important;
    background: var(--blue-dim);
}

/* ── EMPTY STATE ─────────────────────────── */
.empty {
    text-align: center;
    padding: 4rem 2rem;
}
.empty-title {
    font-family: 'Inter Tight', sans-serif;
    font-size: 1.4rem;
    font-weight: 900;
    color: var(--border-hi) !important;
    letter-spacing: -.01em;
    margin-bottom: .5rem;
}
.empty-sub { font-size: .85rem; color: var(--muted) !important; line-height: 1.7; }

/* ── FOOTER ──────────────────────────────── */
.footer {
    margin-top: 3rem;
    padding: 1.25rem 0;
    border-top: 1px solid var(--border);
    display: flex;
    justify-content: space-between;
    align-items: center;
}
.footer-l {
    font-family: 'Inter Tight', sans-serif;
    font-size: .65rem;
    font-weight: 900;
    letter-spacing: .18em;
    color: var(--muted) !important;
}
.footer-r { font-size: .6rem; color: var(--muted) !important; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────
# SECRETS
# ─────────────────────────────────────────────────────────────

claude_key = st.secrets.get("ANTHROPIC_API_KEY", "")

# ─────────────────────────────────────────────────────────────
# OTP / SES AUTH
# ─────────────────────────────────────────────────────────────

ALLOWED_DOMAIN     = "@segaamerica.com"
OTP_EXPIRY_SECS    = 600   # 10 minutes
COOKIE_EXPIRY_DAYS = 7
COOKIE_NAME        = "sega_ratings_auth"


def _send_otp(email: str, code: str) -> bool:
    """Send a 6-digit OTP via AWS SES. Returns True on success."""
    try:
        import boto3
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
                "Subject": {"Data": "Ratings Tracker — Verification Code", "Charset": "UTF-8"},
                "Body": {
                    "Text": {
                        "Data": f"Your verification code is: {code}\n\nExpires in 10 minutes.",
                        "Charset": "UTF-8",
                    },
                    "Html": {
                        "Data": f"""
                        <div style="font-family:Arial,sans-serif;max-width:480px;margin:0 auto;padding:32px 24px;background:#f8f9ff;">
                          <div style="font-size:22px;font-weight:900;letter-spacing:.1em;color:#4080ff;margin-bottom:4px;">SEGA</div>
                          <div style="font-size:14px;color:#444;margin-bottom:28px;">Ratings Tracker · AI Assistant</div>
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
    except Exception as e:
        st.error(f"Failed to send email: {e}")
        return False


def _sign_cookie(email: str) -> str:
    """Return a base64-encoded HMAC-signed token: email|expiry|signature."""
    secret  = st.secrets.get("COOKIE_SIGNING_KEY", "change-this-secret")
    expiry  = int(time.time()) + (COOKIE_EXPIRY_DAYS * 86400)
    payload = f"{email}|{expiry}"
    sig     = hmac.new(secret.encode(), payload.encode(), hashlib.sha256).hexdigest()
    return base64.urlsafe_b64encode(f"{payload}|{sig}".encode()).decode()


def _verify_cookie(token: str) -> str | None:
    """Verify signed cookie. Returns email if valid, else None."""
    try:
        secret  = st.secrets.get("COOKIE_SIGNING_KEY", "change-this-secret")
        decoded = base64.urlsafe_b64decode(token.encode()).decode()
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


# ── Cookie manager ────────────────────────────────────────────
try:
    import extra_streamlit_components as stx
    _cookie_mgr = stx.CookieManager(key="ratings_cookies")
    _existing   = _cookie_mgr.get(COOKIE_NAME)
except Exception:
    _cookie_mgr = None
    _existing   = None

_cookie_email = _verify_cookie(_existing) if _existing else None

# ─────────────────────────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────────────────────────

for k, v in [
    ("auth_verified", False), ("auth_email", ""),
    ("otp_code", ""), ("otp_email", ""), ("otp_expiry", 0),
    ("otp_sent", False), ("otp_attempts", 0),
    ("chat_history", []), ("chat_pending", False), ("pending_prompt", ""),
]:
    if k not in st.session_state:
        st.session_state[k] = v

if _cookie_email and not st.session_state.auth_verified:
    st.session_state.auth_verified = True
    st.session_state.auth_email    = _cookie_email

# ─────────────────────────────────────────────────────────────
# LOGIN GATE
# ─────────────────────────────────────────────────────────────

if not st.session_state.auth_verified:
    st.markdown("<style>section[data-testid='stSidebar']{display:none!important;}</style>",
                unsafe_allow_html=True)

    _, mid, _ = st.columns([1, 2, 1])
    with mid:
        st.markdown("""
        <div class="auth-wrap">
          <div class="auth-logo">SEGA</div>
          <div class="auth-title">Ratings Tracker · AI Assistant</div>
          <div class="auth-sub">Sign in with your SEGA America email to continue</div>
        </div>""", unsafe_allow_html=True)

        if not st.session_state.otp_sent:
            email_in = st.text_input("Email", placeholder="you@segaamerica.com",
                                     label_visibility="collapsed", key="login_email")
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
                                    label_visibility="collapsed", max_chars=6, key="login_code")
            if st.button("Verify code", key="btn_verify"):
                if st.session_state.otp_attempts >= 5:
                    st.error("Too many attempts. Please request a new code.")
                    st.session_state.otp_sent = False
                elif time.time() > st.session_state.otp_expiry:
                    st.error("Code expired. Please request a new one.")
                    st.session_state.otp_sent = False
                elif code_in.strip() != st.session_state.otp_code:
                    st.session_state.otp_attempts += 1
                    rem = 5 - st.session_state.otp_attempts
                    st.error(f"Incorrect code. {rem} attempt{'s' if rem != 1 else ''} remaining.")
                else:
                    st.session_state.auth_verified = True
                    st.session_state.auth_email    = st.session_state.otp_email
                    st.session_state.otp_code      = ""
                    if _cookie_mgr:
                        tok = _sign_cookie(st.session_state.auth_email)
                        _cookie_mgr.set(COOKIE_NAME, tok, expires_at=None, key="set_ck")
                    st.rerun()

            if st.button("← Use a different email", key="btn_back"):
                st.session_state.otp_sent = False
                st.session_state.otp_code = ""
                st.rerun()

        st.markdown(
            f'<div class="auth-note">Restricted to {ALLOWED_DOMAIN} only. '
            f'Codes expire in 10 minutes.</div>',
            unsafe_allow_html=True,
        )
    st.stop()

# ─────────────────────────────────────────────────────────────
# CONSTANTS — expected columns
# ─────────────────────────────────────────────────────────────

RATING_COLS = ["ESRB\u200e \u200e \u200b(Americas)", "PEGI\u200e \u200e (Europe)",
               "USK \u200e \u200e (Germany)", "ACB \u200e \u200e (Australia)",
               "CERO \u200e (Japan)", "IARC"]

RATING_ALIASES = {
    "esrb": "ESRB", "pegi": "PEGI", "usk": "USK",
    "acb": "ACB", "cero": "CERO", "iarc": "IARC",
}

SUGGESTED_PROMPTS = [
    "Which titles are missing ESRB ratings?",
    "Show me everything releasing this month",
    "Which titles still need rating certificates uploaded?",
    "List all titles with PEGI ratings and their values",
    "Which titles are under embargo?",
    "Summarise the overall ratings status across all regions",
    "Which titles have CERO ratings?",
    "Show titles where Status is not yet complete",
]

# ─────────────────────────────────────────────────────────────
# DATA LOADING
# ─────────────────────────────────────────────────────────────

@st.cache_data(show_spinner=False)
def load_csv(file_bytes: bytes, fname: str) -> pd.DataFrame:
    for enc in ("utf-8", "utf-8-sig", "latin-1", "cp1252"):
        try:
            df = pd.read_csv(io.BytesIO(file_bytes), encoding=enc)
            break
        except (UnicodeDecodeError, Exception):
            continue
    else:
        df = pd.read_csv(io.BytesIO(file_bytes), encoding="utf-8", errors="replace")

    df.columns = df.columns.str.strip()

    for col in ["Global Release Date", "Embargo Date"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    return df


def find_col(df: pd.DataFrame, keyword: str) -> str | None:
    """Case-insensitive partial column match."""
    kw = keyword.lower()
    for c in df.columns:
        if kw in c.lower():
            return c
    return None


def df_to_context(df: pd.DataFrame, max_rows: int = 400) -> str:
    """Serialise the dataframe for Claude's system prompt."""
    total = len(df)
    sample = df.head(max_rows)
    csv_str = sample.to_csv(index=False)
    note = f"\n[Showing {min(total, max_rows)} of {total} total rows]" if total > max_rows else ""
    return csv_str + note

# ─────────────────────────────────────────────────────────────
# TOP NAV
# ─────────────────────────────────────────────────────────────

st.markdown("""
<div class="topbar">
  <div class="topbar-brand"><span class="hi">SEGA</span> RATINGS TRACKER</div>
  <div class="topbar-sep"></div>
  <div class="topbar-sub">AI Assistant · Platform Ops</div>
  <div class="topbar-badge">Internal</div>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────
# SIDEBAR — file upload + filters
# ─────────────────────────────────────────────────────────────

with st.sidebar:
    st.markdown(
        f'<div style="font-size:.7rem;font-weight:600;color:var(--muted);margin-bottom:.25rem;">'
        f'Signed in as<br>'
        f'<span style="color:var(--text);font-weight:700;">{st.session_state.auth_email}</span>'
        f'</div>', unsafe_allow_html=True,
    )
    if st.button("Sign out", key="sign_out"):
        if _cookie_mgr:
            _cookie_mgr.delete(COOKIE_NAME, key="del_ck")
        for k in ["auth_verified", "auth_email", "otp_sent", "otp_code",
                  "otp_email", "otp_expiry", "otp_attempts"]:
            st.session_state[k] = False if k == "auth_verified" else ""
        st.rerun()

    st.markdown("---")
    st.markdown("### 📁 Data Source")
    uploaded = st.file_uploader(
        "Upload Ratings CSV",
        type=["csv"],
        help="Export your ratings tracker as CSV and upload here.",
    )

    df = None
    if uploaded:
        df = load_csv(uploaded.getvalue(), uploaded.name)
        st.markdown(f"<div style='font-size:.7rem;color:var(--muted);margin-top:.5rem;'>"
                    f"✓ {len(df):,} rows · {len(df.columns)} columns</div>",
                    unsafe_allow_html=True)

    st.markdown("---")

    if df is not None:
        st.markdown("### Filters")

        # Status filter
        status_col = find_col(df, "status")
        f_status = "All"
        if status_col:
            f_status = st.selectbox("Status", ["All"] + sorted(df[status_col].dropna().astype(str).unique()))

        # Format filter
        format_col = find_col(df, "format")
        f_format = "All"
        if format_col:
            f_format = st.selectbox("Format", ["All"] + sorted(df[format_col].dropna().astype(str).unique()))

        # Needs certs filter
        cert_col = find_col(df, "needs rating")
        f_certs = "All"
        if cert_col:
            f_certs = st.selectbox("Needs Certs", ["All"] + sorted(df[cert_col].dropna().astype(str).unique()))

        # Apply filters
        filtered_df = df.copy()
        if f_status != "All" and status_col:
            filtered_df = filtered_df[filtered_df[status_col].astype(str) == f_status]
        if f_format != "All" and format_col:
            filtered_df = filtered_df[filtered_df[format_col].astype(str) == f_format]
        if f_certs != "All" and cert_col:
            filtered_df = filtered_df[filtered_df[cert_col].astype(str) == f_certs]

        st.markdown(f"<div style='font-size:.7rem;color:var(--muted);'>"
                    f"{len(filtered_df):,} titles after filters</div>",
                    unsafe_allow_html=True)
        st.markdown("---")

    st.markdown("### 🤖 Model")
    model = st.selectbox("Claude model", [
        "claude-sonnet-4-20250514",
        "claude-opus-4-20250514",
        "claude-haiku-4-5-20251001",
    ], label_visibility="collapsed")

# ─────────────────────────────────────────────────────────────
# NO FILE UPLOADED — empty state
# ─────────────────────────────────────────────────────────────

if df is None:
    st.markdown("""
    <div class="empty">
      <div class="empty-title">UPLOAD YOUR RATINGS CSV</div>
      <div class="empty-sub">
        Use the sidebar to upload your ratings tracker export.<br><br>
        <strong>Expected columns:</strong><br>
        Title · Global Release Date · Format · Status<br>
        ESRB · PEGI · USK · ACB · CERO · IARC<br>
        Tracker URL · Embargo Date · Needs Rating Certificates Uploaded
      </div>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

# Use filtered df from here on
display_df = filtered_df

# ─────────────────────────────────────────────────────────────
# STAT CARDS
# ─────────────────────────────────────────────────────────────

st.markdown('<div class="sec">OVERVIEW</div>', unsafe_allow_html=True)

total_titles = len(display_df)

# Count titles with at least one rating filled
rating_cols_present = [c for c in display_df.columns
                       if any(r.lower() in c.lower()
                              for r in ["esrb","pegi","usk","acb","cero","iarc"])]
has_any_rating = display_df[rating_cols_present].notna().any(axis=1).sum() \
    if rating_cols_present else 0

# Needs certs
cert_col_found = find_col(display_df, "needs rating")
needs_certs = 0
if cert_col_found:
    needs_certs = int(display_df[cert_col_found]
                      .astype(str).str.lower()
                      .isin(["yes","true","1","y"]).sum())

# Embargo count
embargo_col = find_col(display_df, "embargo")
under_embargo = 0
if embargo_col:
    under_embargo = int(display_df[embargo_col].notna().sum())

# Missing ESRB
esrb_col = find_col(display_df, "esrb")
missing_esrb = 0
if esrb_col:
    missing_esrb = int(display_df[esrb_col].isna().sum())

stats = [
    ("Total Titles",    total_titles,   "#4080ff"),
    ("Any Rating",      has_any_rating, "#20c65a"),
    ("Missing ESRB",    missing_esrb,   "#ff3d52"),
    ("Needs Certs",     needs_certs,    "#f0a500"),
    ("Under Embargo",   under_embargo,  "#9b59f5"),
]

# Render stat cards using columns
cols = st.columns(len(stats))
for col, (label, val, color) in zip(cols, stats):
    with col:
        st.markdown(f"""
        <div class="stat-card" style="--accent-color:{color}">
          <div class="stat-val">{val}</div>
          <div class="stat-lbl">{label}</div>
        </div>""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────
# DATA TABLE
# ─────────────────────────────────────────────────────────────

st.markdown('<div class="sec">TITLES</div>', unsafe_allow_html=True)

# Choose columns to display — title first, then key cols
title_col = find_col(display_df, "title")
priority_cols = (
    ([title_col] if title_col else []) +
    [c for c in display_df.columns if find_col(pd.DataFrame(columns=[c]), "release") or "release" in c.lower()][:1] +
    [c for c in display_df.columns if "format" in c.lower() or "status" in c.lower() or
     any(r in c.lower() for r in ["esrb","pegi","usk","acb","cero","iarc"])]
)
# Deduplicate while preserving order
seen = set()
show_cols = [c for c in priority_cols + list(display_df.columns)
             if c not in seen and not seen.add(c)][:14]

st.dataframe(
    display_df[show_cols].reset_index(drop=True),
    use_container_width=True,
    height=340,
)

# ─────────────────────────────────────────────────────────────
# AI CHAT
# ─────────────────────────────────────────────────────────────

st.markdown('<div class="sec">ASK CLAUDE</div>', unsafe_allow_html=True)

if not claude_key:
    st.warning("Add `ANTHROPIC_API_KEY` to `.streamlit/secrets.toml` to enable the AI assistant.")
    st.code('ANTHROPIC_API_KEY = "sk-ant-..."', language="toml")
    st.stop()

# Suggested prompt chips (rendered as buttons in a row)
st.markdown("**Suggested questions**")
chip_cols = st.columns(4)
for i, prompt in enumerate(SUGGESTED_PROMPTS):
    with chip_cols[i % 4]:
        if st.button(prompt, key=f"chip_{i}"):
            st.session_state.chat_history.append({"role": "user", "content": prompt})
            st.session_state.chat_pending = True
            st.rerun()

st.markdown("<div style='margin-top:.5rem'></div>", unsafe_allow_html=True)

# Build Claude system prompt
SYSTEM = f"""You are an expert ratings and compliance analyst for SEGA America's Platform Operations team.
You help the team manage video game content ratings across global territories.

COLUMN GUIDE:
- Title / タイトル: Game title (may include Japanese)
- Global Release Date: Planned release date
- Format: Platform/format (e.g. PS5, Xbox, Switch, PC)
- Status: Current workflow status
- ESRB (Americas): Rating from the Entertainment Software Rating Board
- PEGI (Europe): Pan European Game Information rating
- USK (Germany): Unterhaltungssoftware Selbstkontrolle rating
- ACB (Australia): Australian Classification Board rating
- CERO (Japan): Computer Entertainment Rating Organization rating
- IARC: International Age Rating Coalition rating
- Tracker URL: Link to the ratings tracker
- Embargo Date: Date until which information is embargoed
- Needs Rating Certificates Uploaded: Whether certificates still need to be uploaded

CURRENT DATA ({len(display_df)} titles after active filters):
{df_to_context(display_df)}

INSTRUCTIONS:
- Answer questions about ratings status, missing ratings, upcoming releases, embargoes, or certificates
- When listing titles, use markdown tables for clarity
- Be concise and specific — reference actual titles and data from the CSV
- If a question cannot be answered from the data, say so clearly
- Dates should be formatted as Month DD, YYYY for readability
- Treat empty/NaN cells as "not yet rated" or "not applicable" depending on context"""

# Render chat history
for msg in st.session_state.chat_history:
    with st.chat_message(msg["role"]):
        st.markdown(msg["content"])

# Stream pending reply
if st.session_state.chat_pending:
    st.session_state.chat_pending = False
    try:
        import anthropic
        client   = anthropic.Anthropic(api_key=claude_key)
        api_msgs = [{"role": m["role"], "content": m["content"]}
                    for m in st.session_state.chat_history]
        with st.chat_message("assistant"):
            reply = ""
            ph    = st.empty()
            with client.messages.stream(
                model=model,
                max_tokens=1500,
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

# Chat input
user_msg = st.chat_input("Ask about ratings, embargoes, missing certs…")
if user_msg:
    st.session_state.chat_history.append({"role": "user", "content": user_msg})
    st.session_state.chat_pending = True
    st.rerun()

# Clear button
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
  <div class="footer-l">SEGA RATINGS TRACKER · AI ASSISTANT</div>
  <div class="footer-r">Powered by Claude · Data processed locally · Internal use only</div>
</div>
""", unsafe_allow_html=True)