import streamlit as st
from concurrent.futures import ThreadPoolExecutor, as_completed
import json
import os
import tempfile
import io
import time
import hashlib
import hmac
import random
import base64
import requests
import pandas as pd
import pypdf
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from lxml import etree

# ─────────────────────────────────────────────────────────────
# THEME & FILL HELPERS — prevent white-on-white text
# ─────────────────────────────────────────────────────────────

# Dark SEGA theme: dk1=white, lt1=dark-navy.
# This means any text that falls back to theme colour is white,
# which is correct on our dark backgrounds.  Replaces the default
# Office theme (dk1=windowText which can resolve to white on some
# systems/dark-mode setups, making text invisible).
_SEGA_THEME_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="SEGA Dark">'
    '<a:themeElements>'
    '<a:clrScheme name="SEGA Dark">'
    '<a:dk1><a:srgbClr val="FFFFFF"/></a:dk1>'
    '<a:lt1><a:srgbClr val="040A1C"/></a:lt1>'
    '<a:dk2><a:srgbClr val="0055AA"/></a:dk2>'
    '<a:lt2><a:srgbClr val="1A2A4A"/></a:lt2>'
    '<a:accent1><a:srgbClr val="00AADD"/></a:accent1>'
    '<a:accent2><a:srgbClr val="F5C218"/></a:accent2>'
    '<a:accent3><a:srgbClr val="00BB66"/></a:accent3>'
    '<a:accent4><a:srgbClr val="CC2244"/></a:accent4>'
    '<a:accent5><a:srgbClr val="8899BB"/></a:accent5>'
    '<a:accent6><a:srgbClr val="AABBCC"/></a:accent6>'
    '<a:hlink><a:srgbClr val="00AADD"/></a:hlink>'
    '<a:folHlink><a:srgbClr val="0055AA"/></a:folHlink>'
    '</a:clrScheme>'
    '<a:fontScheme name="SEGA">'
    '<a:majorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>'
    '<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>'
    '</a:fontScheme>'
    '<a:fmtScheme name="Office">'
    '<a:fillStyleLst>'
    '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
    '<a:gradFill rotWithShape="1"><a:gsLst>'
    '<a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/></a:schemeClr></a:gs>'
    '<a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="15000"/></a:schemeClr></a:gs>'
    '</a:gsLst><a:lin ang="16200000" scaled="1"/></a:gradFill>'
    '<a:gradFill rotWithShape="1"><a:gsLst>'
    '<a:gs pos="0"><a:schemeClr val="phClr"><a:shade val="51000"/></a:schemeClr></a:gs>'
    '<a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="94000"/></a:schemeClr></a:gs>'
    '</a:gsLst><a:lin ang="16200000" scaled="0"/></a:gradFill>'
    '</a:fillStyleLst>'
    '<a:lnStyleLst>'
    '<a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>'
    '<a:ln w="25400"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>'
    '<a:ln w="38100"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>'
    '</a:lnStyleLst>'
    '<a:effectStyleLst>'
    '<a:effectStyle><a:effectLst/></a:effectStyle>'
    '<a:effectStyle><a:effectLst/></a:effectStyle>'
    '<a:effectStyle><a:effectLst/></a:effectStyle>'
    '</a:effectStyleLst>'
    '<a:bgFillStyleLst>'
    '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
    '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
    '<a:gradFill rotWithShape="1"><a:gsLst>'
    '<a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/></a:schemeClr></a:gs>'
    '<a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="98000"/></a:schemeClr></a:gs>'
    '</a:gsLst><a:lin ang="16200000" scaled="1"/></a:gradFill>'
    '</a:bgFillStyleLst>'
    '</a:fmtScheme>'
    '</a:themeElements>'
    '</a:theme>'
)


def _patch_theme(pptx_bytes: bytes) -> bytes:
    """Swap every theme file in the PPTX zip for the SEGA dark theme."""
    import zipfile, re as _re
    buf_in  = io.BytesIO(pptx_bytes)
    buf_out = io.BytesIO()
    theme_bytes = _SEGA_THEME_XML.encode("utf-8")
    with zipfile.ZipFile(buf_in, "r") as zin,          zipfile.ZipFile(buf_out, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if _re.match(r"ppt/theme/theme\d+\.xml", item.filename):
                data = theme_bytes
            zout.writestr(item, data)
    return buf_out.getvalue()


def _lock_txb(txb):
    """
    Add explicit noFill + noLine to a textbox shape element so it can never
    inherit a white background from the theme or slide layout.
    Must be called after add_textbox() but before saving.
    """
    sp   = txb._element
    spPr = sp.find(qn("p:spPr"))
    if spPr is None:
        return
    # Remove any existing fill child
    for tag in ("a:solidFill", "a:gradFill", "a:pattFill", "a:blipFill", "a:grpFill", "a:noFill"):
        el = spPr.find(qn(tag))
        if el is not None:
            spPr.remove(el)
    # Insert <a:noFill/> after the <a:xfrm> and <a:prstGeom> elements
    no_fill = etree.SubElement(spPr, qn("a:noFill"))
    # Also kill any line (border) that might show as white
    ln = spPr.find(qn("a:ln"))
    if ln is None:
        ln = etree.SubElement(spPr, qn("a:ln"))
    # Clear the line's fill children and set noFill
    for child in list(ln):
        ln.remove(child)
    etree.SubElement(ln, qn("a:noFill"))


st.set_page_config(
    page_title="SEGA Intelligence Analyzer",
    page_icon="🎮",
    layout="wide",
)

# ─────────────────────────────────────────────────────────────
# OTP AUTHENTICATION  (mirrors BIreport.py exactly)
# ─────────────────────────────────────────────────────────────

ALLOWED_DOMAIN     = "@segaamerica.com"
OTP_EXPIRY_SECS    = 600   # 10 minutes
COOKIE_EXPIRY_DAYS = 7
COOKIE_NAME        = "sega_analyzer_auth"


def _send_otp(email: str, code: str) -> bool:
    """Send OTP via AWS SES using credentials from st.secrets."""
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
                "Subject": {
                    "Data": "SEGA Intelligence Analyzer — Your verification code",
                    "Charset": "UTF-8",
                },
                "Body": {
                    "Text": {
                        "Data": (
                            f"Your SEGA Intelligence Analyzer verification code is: {code}\n\n"
                            "This code expires in 10 minutes.\n"
                            "If you didn't request this, you can safely ignore this email."
                        ),
                        "Charset": "UTF-8",
                    },
                    "Html": {
                        "Data": f"""
                        <div style="font-family:Arial,sans-serif;max-width:480px;margin:0 auto;padding:32px 24px;">
                          <div style="font-size:22px;font-weight:900;letter-spacing:0.1em;color:#1A6BFF;margin-bottom:4px;">SEGA</div>
                          <div style="font-size:14px;color:#444;margin-bottom:28px;">Intelligence Analyzer</div>
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
    .auth-logo  { font-family:'Arial Black',sans-serif; font-size:1.6rem; font-weight:900;
                  letter-spacing:0.12em; color:var(--blue) !important; margin-bottom:0.2rem; }
    .auth-title { font-size:1rem; font-weight:700; color:var(--text) !important; margin-bottom:0.25rem; }
    .auth-sub   { font-size:0.8rem; color:var(--muted) !important; margin-bottom:1.5rem; }
    .auth-note  { font-size:0.72rem; color:var(--muted) !important; margin-top:1rem;
                  text-align:center; line-height:1.5; }
    :root { --bg:#0a0c1a; --surface:#0f1120; --border:#232640; --blue:#4080ff; --text:#eef0fa; --muted:#5a5f82; }
    html, body, .stApp { background: var(--bg) !important; }
    </style>
    """, unsafe_allow_html=True)

    _lc, _mc, _rc = st.columns([1, 2, 1])
    with _mc:
        st.markdown("""
        <div class="auth-wrap">
          <div class="auth-logo">SEGA</div>
          <div class="auth-title">Intelligence Analyzer</div>
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
            _send_btn = st.button("Send verification code", use_container_width=True)

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
            _verify_btn = st.button("Verify code", use_container_width=True)

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
                        _cookie_manager.set(COOKIE_NAME, _token, expires_at=None, key="set_auth_cookie")
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

# ─────────────────────────────────────────────────────────────
# GLOBAL STYLES
# ─────────────────────────────────────────────────────────────
st.markdown("""
<style>
[data-testid="stToolbar"]      { display: none !important; }
[data-testid="stDecoration"]   { display: none !important; }
header[data-testid="stHeader"] { display: none !important; }

html, body, [class*="css"], .stApp {
    background-color: #0a0c1a !important;
    color: #e2e8f0 !important;
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif !important;
}
[data-testid="stSidebar"] { background: #0f172a !important; border-right: 1px solid #1e293b !important; }
[data-testid="stSidebar"] * { color: #cbd5e1 !important; }
[data-testid="stSidebar"] .block-container { padding: 1.25rem 1rem !important; max-width: 100% !important; }
.block-container { max-width: 1400px !important; padding: 1.5rem 2rem 3rem !important; margin: 0 auto !important; }

h1 { font-size: 2rem !important; font-weight: 700 !important; color: #f1f5f9 !important;
     letter-spacing: -.02em !important; margin-bottom: .15rem !important; }

.sidebar-section { font-size:.7rem; font-weight:600; text-transform:uppercase;
                   letter-spacing:.1em; color:#64748b; margin:1.2rem 0 .4rem; display:block; }

[data-testid="stFileUploader"] {
    border: 1px dashed #334155 !important;
    border-radius: 8px !important;
    background: transparent !important;
}

.stTextInput input, .stTextArea textarea {
    background: #1e293b !important;
    border: 1px solid #334155 !important;
    border-radius: 6px !important;
    color: #e2e8f0 !important;
    font-size: .85rem !important;
}

.section-label {
    font-size: .68rem; font-weight: 600; text-transform: uppercase;
    letter-spacing: .1em; color: #475569;
    margin: 2rem 0 .75rem;
    border-bottom: 1px solid #1e293b;
    padding-bottom: .4rem;
}

.status-card {
    background: #1e293b; border: 1px solid #334155; border-radius: 8px; padding: 1rem 1.25rem;
}
.status-card-label { font-size:.65rem; text-transform:uppercase; letter-spacing:.1em; color:#64748b; margin-bottom:.3rem; }
.status-card-value { font-size:.95rem; font-weight:600; color:#f1f5f9; line-height:1.4; }

.step-row { display:flex; align-items:center; gap:.6rem; padding:.45rem .75rem;
            margin:.2rem 0; border-radius:5px; font-size:.8rem; }
.step-done    { background:rgba(52,211,153,.1); color:#34d399; }
.step-active  { background:rgba(96,165,250,.12); color:#60a5fa; }
.step-pending { color:#475569; }

.result-log {
    background: #0f172a; border: 1px solid #1e293b; border-radius: 8px;
    padding: 1rem 1.25rem; font-size: .82rem;
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
    color: #94a3b8; max-height: 480px; overflow-y: auto; line-height: 1.7;
}
.result-log b  { color: #e2e8f0; font-weight: 600; }
.result-log i  { color: #60a5fa; }
.log-detail    { color: #64748b; font-size: .76rem; line-height: 1.55; display: block;
                 margin-top: .15rem; margin-bottom: .35rem; }
.log-entry     { border-left: 2px solid #1e3a5f; padding-left: .6rem;
                 margin-bottom: .5rem; }

/* ── Active spinner entry ── */
@keyframes spin { to { transform: rotate(360deg); } }
.log-active {
    border-left: 2px solid #00AADD;
    padding-left: .6rem;
    margin-bottom: .5rem;
    display: flex;
    align-items: flex-start;
    gap: .55rem;
}
.log-spinner {
    flex-shrink: 0;
    width: 14px; height: 14px;
    margin-top: 3px;
    border: 2px solid rgba(0,170,221,0.25);
    border-top-color: #00AADD;
    border-radius: 50%;
    animation: spin 0.8s linear infinite;
    display: inline-block;
}
.log-active-text { flex: 1; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────
# SIDEBAR — settings (post-auth)
# ─────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("<div style='font-size:1rem;font-weight:700;color:#f1f5f9;margin-bottom:.25rem;'>Intelligence Analyzer</div>", unsafe_allow_html=True)
    st.markdown("<div style='font-size:.65rem;color:#475569;text-transform:uppercase;letter-spacing:.1em;margin-bottom:1rem;'>Game Competitive Analysis</div>", unsafe_allow_html=True)

    st.markdown('<span class="sidebar-section">Model</span>', unsafe_allow_html=True)
    model = st.selectbox(
        "Model", ["claude-sonnet-4-5", "claude-opus-4-5", "claude-haiku-4-5-20251001"],
        label_visibility="hidden",
    )

    st.markdown('<span class="sidebar-section">Options</span>', unsafe_allow_html=True)
    web_search_enabled = st.checkbox("Web search for reference game", value=True)
    slide_count        = st.slider("Target slides", 6, 20, 10)

    st.markdown('<span class="sidebar-section">Theme</span>', unsafe_allow_html=True)
    theme_preset = st.selectbox(
        "Theme",
        ["SEGA Blue — Corporate Executive", "SEGA Dark — Game Reveal Style", "SEGA Sonic — High Energy"],
        label_visibility="hidden",
    )

    template_file = st.file_uploader(
        "Upload .pptx template (optional)",
        type=["pptx"],
        label_visibility="hidden",
        help="Upload a branded .pptx file to extract its colour palette and fonts. "
             "Overrides the theme preset above.",
        key="template_upload",
    )
    if template_file:
        st.caption(f"✓ Template: {template_file.name}")

    st.markdown('<span class="sidebar-section">Pipeline</span>', unsafe_allow_html=True)
    pipeline_steps = st.session_state.get("pipeline_steps", {
        "upload": False, "extract": False, "research": False, "analyze": False, "generate": False,
    })
    for key, label in {"upload":"Document upload","extract":"Content extraction",
                        "research":"Web research","analyze":"AI analysis","generate":"PPTX generation"}.items():
        done = pipeline_steps.get(key, False)
        st.markdown(
            f'<div class="step-row {"step-done" if done else "step-pending"}">'
            f'{"✓" if done else "○"}&nbsp; {label}</div>',
            unsafe_allow_html=True,
        )

# ─────────────────────────────────────────────────────────────
# PAGE HEADER
# ─────────────────────────────────────────────────────────────
st.markdown(
    "<h1>Game Intelligence Analyzer</h1>"
    "<div style='font-size:.78rem;color:#475569;margin-bottom:1.5rem;'>"
    "Upload internal documents &nbsp;·&nbsp; Benchmark against a released title &nbsp;·&nbsp; Generate a SEGA-branded PPTX"
    "</div>",
    unsafe_allow_html=True,
)

# ─────────────────────────────────────────────────────────────
# MAIN LAYOUT
# ─────────────────────────────────────────────────────────────
col_left, col_right = st.columns([1.1, 0.9], gap="large")

with col_left:
    st.markdown('<div class="section-label">Documents</div>', unsafe_allow_html=True)
    uploaded_files = st.file_uploader(
        "Upload internal game documents",
        type=["pdf", "xlsx", "xls", "csv", "txt", "docx"],
        accept_multiple_files=True,
        label_visibility="hidden",
    )
    if uploaded_files:
        st.caption(f"{len(uploaded_files)} file(s): " + ", ".join(f.name for f in uploaded_files))

    st.markdown('<div class="section-label">Analysis inputs</div>', unsafe_allow_html=True)

    game_title = st.text_input(
        "Reference / competitor game title",
        placeholder="e.g. Sonic Frontiers, Metaphor: ReFantazio, Persona 5…",
    )
    business_question = st.text_area(
        "Business question",
        placeholder=(
            "e.g. Compare our internal game's combat mechanics and scope against Sonic Frontiers, "
            "highlighting gaps and opportunities for the executive team…"
        ),
        height=130,
    )
    audience = st.text_input(
        "Presentation audience", value="Executive team",
        placeholder="e.g. Executive team, Product leads, Marketing…",
    )

    run_btn = st.button("⚡ Run analysis", use_container_width=True, type="primary")

with col_right:
    st.markdown('<div class="section-label">Output</div>', unsafe_allow_html=True)
    output_area   = st.empty()
    download_area = st.empty()

    if not run_btn and "pptx_bytes" not in st.session_state:
        output_area.markdown("""
<div class="status-card">
<div class="status-card-label">Ready</div>
<div class="status-card-value" style="color:#475569;font-size:.82rem;line-height:1.8;">
Fill in the inputs on the left and click <strong style="color:#e2e8f0;">Run analysis</strong>.<br><br>
The pipeline will:<br>
&nbsp;1. Extract your uploaded documents<br>
&nbsp;2. Search the web for the reference game<br>
&nbsp;3. Run a Claude-powered comparative analysis<br>
&nbsp;4. Render a SEGA-branded PPTX for download
</div>
</div>
""", unsafe_allow_html=True)

    if "pptx_bytes" in st.session_state and not run_btn:
        download_area.download_button(
            label="⬇️ Download previous PPTX",
            data=st.session_state["pptx_bytes"],
            file_name=st.session_state.get("pptx_filename", "SEGA_Analysis.pptx"),
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True,
        )

# ─────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────

def extract_text_from_file(uploaded_file) -> str:
    name = uploaded_file.name.lower()
    content = ""
    try:
        if name.endswith(".pdf"):
            reader = pypdf.PdfReader(io.BytesIO(uploaded_file.read()))
            content = "\n\n".join(p.extract_text() or "" for p in reader.pages)
        elif name.endswith((".xlsx", ".xls")):
            sheets = pd.read_excel(io.BytesIO(uploaded_file.read()), sheet_name=None)
            content = "\n\n".join(f"=== Sheet: {s} ===\n{df.to_string(max_rows=200)}" for s, df in sheets.items())
        elif name.endswith(".csv"):
            content = pd.read_csv(io.BytesIO(uploaded_file.read())).to_string(max_rows=300)
        elif name.endswith(".txt"):
            content = uploaded_file.read().decode("utf-8", errors="replace")
        elif name.endswith(".docx"):
            import zipfile, xml.etree.ElementTree as ET
            zf   = zipfile.ZipFile(io.BytesIO(uploaded_file.read()))
            root = ET.fromstring(zf.read("word/document.xml"))
            ns   = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
            content = "\n".join(
                "".join(r.text for r in p.iter(f"{ns}t") if r.text)
                for p in root.iter(f"{ns}p")
            )
        else:
            content = f"[Unsupported type: {name}]"
    except Exception as e:
        content = f"[Extraction error for {uploaded_file.name}: {e}]"
    if len(content) > 15000:
        content = content[:15000] + "\n\n[... truncated ...]"
    return content



# ─────────────────────────────────────────────────────────────
# PPTX GENERATION — pure python-pptx (no Node.js required)
# ─────────────────────────────────────────────────────────────

def _rgb(hex_str: str) -> RGBColor:
    h = hex_str.lstrip("#")
    return RGBColor(int(h[0:2],16), int(h[2:4],16), int(h[4:6],16))

# Slide canvas: 13.3 × 7.5 inches (LAYOUT_WIDE)
W_IN, H_IN = 13.3, 7.5

def _in(v):  return Inches(v)
def _pt(v):  return Pt(v)

def _rect(slide, x, y, w, h, fill_hex, alpha=None):
    shape = slide.shapes.add_shape(
        1,  # MSO_SHAPE_TYPE.RECTANGLE
        _in(x), _in(y), _in(w), _in(h)
    )
    shape.line.fill.background()
    if fill_hex:
        shape.fill.solid()
        shape.fill.fore_color.rgb = _rgb(fill_hex)
    if alpha is not None:
        # alpha 0-100 (0=opaque, 100=transparent in pptx land)
        shape.fill.fore_color._element.getparent().getparent()  # noop touch
    shape.line.fill.background()
    return shape

def _add_text(slide, text, x, y, w, h,
              size=12, bold=False, italic=False, color="FFFFFF",
              align=PP_ALIGN.LEFT, wrap=True, font_name="Calibri"):
    txb = slide.shapes.add_textbox(_in(x), _in(y), _in(w), _in(h))
    _lock_txb(txb)  # prevent white-box inheritance
    tf  = txb.text_frame
    tf.word_wrap = wrap
    p   = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = str(text)
    run.font.size  = _pt(size)
    run.font.bold  = bold
    run.font.italic = italic
    run.font.color.rgb = _rgb(color)
    run.font.name  = font_name
    return txb

def _chrome(slide, idx, total, C):
    """Top stripe, bottom bar, page number, right edge accent."""
    _rect(slide, 0, 0,       W_IN, 0.1,  C["accent"])
    _rect(slide, 0, H_IN-0.36, W_IN, 0.36, C["header_bg"])
    _add_text(slide, "SEGA INTELLIGENCE ANALYZER",
              0.3, H_IN-0.30, 5, 0.24,
              size=7, bold=True, color=C["midgray"],
              font_name=C["body_font"])
    _add_text(slide, f"{idx} / {total}",
              W_IN-1.4, H_IN-0.30, 1.1, 0.24,
              size=8, color=C["neutral"], align=PP_ALIGN.RIGHT,
              font_name=C["body_font"])
    _rect(slide, W_IN-0.1, 0, 0.1, H_IN, C["primary"])

def _set_bg(slide, hex_color):
    bg   = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = _rgb(hex_color)

def _add_bullets(slide, bullets, x, y, w, h, bullet_color, text_color,
                 size=13, font_name=C["body_font"]):
    if not bullets:
        return
    txb = slide.shapes.add_textbox(_in(x), _in(y), _in(w), _in(h))
    _lock_txb(txb)  # prevent white-box inheritance
    tf  = txb.text_frame
    tf.word_wrap = True
    first = True
    for b in bullets:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        # bullet marker
        r1 = p.add_run()
        r1.text = "▸ "
        r1.font.color.rgb = _rgb(bullet_color)
        r1.font.bold  = True
        r1.font.size  = _pt(size)
        r1.font.name  = font_name
        # bullet text
        r2 = p.add_run()
        r2.text = b
        r2.font.color.rgb = _rgb(text_color)
        r2.font.size  = _pt(size)
        r2.font.name  = font_name

# ── Slide type renderers ──────────────────────────────────────

def _slide_title(slide, s, idx, total, C):
    _set_bg(slide, C["bg"])
    _rect(slide, 7.4, 0, 5.9, H_IN, C["primary"])
    # SEGA wordmark
    _add_text(slide, "SEGA", 8.0, 0.55, 4.8, 0.9,
              size=52, bold=True, color=C["white"], font_name="Arial Black",
              align=PP_ALIGN.CENTER)
    _rect(slide, 8.3, 1.52, 4.2, 0.05, C["accent"])
    _add_text(slide, "INTELLIGENCE ANALYZER", 8.0, 1.58, 4.8, 0.38,
              size=9, color=C["accent"], align=PP_ALIGN.CENTER, font_name=C["body_font"])
    # Main title
    _add_text(slide, s.get("title",""), 0.6, 1.8, 6.8, 2.2,
              size=34, bold=True, color=C["white"], font_name=C["body_font"])
    if s.get("subtitle"):
        _add_text(slide, s["subtitle"], 0.6, 4.2, 6.6, 0.6,
                  size=16, italic=True, color=C["accent"], font_name=C["body_font"])
    if s.get("body"):
        _add_text(slide, s["body"], 0.6, 4.95, 6.6, 1.4,
                  size=12, color=C["midgray"], font_name=C["body_font"])
    _chrome(slide, idx, total, C)

def _slide_section(slide, s, idx, total, C):
    _set_bg(slide, C["bg"])
    _rect(slide, 0, 0, 0.28, H_IN, C["accent"])
    _add_text(slide, s.get("title",""), 0.55, 2.3, 12.0, 1.6,
              size=40, bold=True, color=C["white"], font_name=C["heading_font"])
    if s.get("subtitle"):
        _add_text(slide, s["subtitle"], 0.55, 4.1, 11.0, 0.65,
                  size=19, italic=True, color=C["accent"], font_name=C["body_font"])
    _chrome(slide, idx, total, C)

def _slide_bullets(slide, s, idx, total, C):
    _set_bg(slide, C["bg"])
    _rect(slide, 0, 0.1, W_IN, 0.95, C["subtle"])
    _rect(slide, 0, 0.1, 0.17, 0.95, C["accent"])
    _add_text(slide, s.get("title",""), 0.38, 0.14, W_IN-0.6, 0.84,
              size=22, bold=True, color=C["white"], font_name=C["body_font"])
    _add_bullets(slide, s.get("bullets",[]),
                 0.45, 1.2, W_IN-1.0, H_IN-2.0,
                 bullet_color=C["accent"], text_color=C["white"],
                 size=13)
    if s.get("body"):
        _add_text(slide, s["body"], 0.45, H_IN-1.45, W_IN-1.0, 1.0,
                  size=11, italic=True, color=C["midgray"], font_name=C["body_font"])
    _chrome(slide, idx, total, C)

def _slide_stats(slide, s, idx, total, C):
    _set_bg(slide, C["bg"])
    _rect(slide, 0, 0.1, W_IN, 0.95, C["subtle"])
    _rect(slide, 0, 0.1, 0.17, 0.95, C["gold"])
    _add_text(slide, s.get("title",""), 0.38, 0.14, W_IN-0.6, 0.84,
              size=22, bold=True, color=C["white"], font_name=C["body_font"])
    stats = s.get("stats", [])[:4]
    if stats:
        box_w = (W_IN - 1.2) / len(stats)
        for i, stat in enumerate(stats):
            x = 0.6 + i * (box_w + 0.1)
            _rect(slide, x, 1.3, box_w, 2.9, C["subtle"])
            _add_text(slide, stat.get("value","—"),
                      x+0.1, 1.55, box_w-0.2, 1.2,
                      size=36, bold=True, color=C["accent"],
                      align=PP_ALIGN.CENTER, font_name=C["body_font"])
            _add_text(slide, stat.get("label",""),
                      x+0.1, 2.85, box_w-0.2, 0.7,
                      size=12, bold=True, color=C["white"],
                      align=PP_ALIGN.CENTER, font_name=C["body_font"])
            if stat.get("note"):
                _add_text(slide, stat["note"],
                          x+0.1, 3.6, box_w-0.2, 0.45,
                          size=9, color=C["midgray"],
                          align=PP_ALIGN.CENTER, font_name=C["body_font"])
    if s.get("body"):
        _add_text(slide, s["body"], 0.45, H_IN-1.5, W_IN-1.0, 1.0,
                  size=11, italic=True, color=C["midgray"], font_name=C["body_font"])
    _chrome(slide, idx, total, C)

def _slide_comparison(slide, s, idx, total, C):
    _set_bg(slide, C["bg"])
    _rect(slide, 0, 0.1, W_IN, 0.95, C["subtle"])
    _rect(slide, 0, 0.1, 0.17, 0.95, C["accent"])
    _add_text(slide, s.get("title",""), 0.38, 0.14, W_IN-0.6, 0.84,
              size=20, bold=True, color=C["white"], font_name=C["body_font"])
    cmp    = s.get("comparison", {})
    rows   = cmp.get("rows", [])
    label_w = 2.8
    col_w   = (W_IN - 1.0 - label_w) / 2
    left_x  = 0.5
    mid_x   = left_x + label_w + 0.05
    right_x = mid_x + col_w + 0.05
    start_y = 1.22
    row_h   = 0.42
    # Column headers
    _rect(slide, mid_x,   start_y, col_w, row_h, C["primary"])
    _rect(slide, right_x, start_y, col_w, row_h, C["gold"])
    _add_text(slide, cmp.get("left_title","Internal"),
              mid_x+0.05, start_y, col_w-0.1, row_h,
              size=10, bold=True, color=C["white"], align=PP_ALIGN.CENTER)
    _add_text(slide, cmp.get("right_title","Reference"),
              right_x+0.05, start_y, col_w-0.1, row_h,
              size=10, bold=True, color=C["white"], align=PP_ALIGN.CENTER)
    for ri, row in enumerate(rows[:10]):
        y   = start_y + row_h + ri * row_h
        row_bg = "0D1530" if ri % 2 == 0 else "111E3A"
        total_w = label_w + col_w * 2 + 0.1
        _rect(slide, left_x, y, total_w, row_h - 0.04, row_bg)
        _add_text(slide, row.get("label",""),
                  left_x+0.08, y, label_w-0.12, row_h,
                  size=9, bold=True, color=C["midgray"], font_name=C["body_font"])
        _add_text(slide, row.get("left","—"),
                  mid_x+0.05, y, col_w-0.1, row_h,
                  size=9, color=C["light"], align=PP_ALIGN.CENTER, font_name=C["body_font"])
        delta = (row.get("delta","") or "").lower()
        dc = C["green"] if delta == "positive" else C["red"] if delta == "negative" else C["neutral"]
        _add_text(slide, row.get("right","—"),
                  right_x+0.05, y, col_w-0.1, row_h,
                  size=9, color=dc, align=PP_ALIGN.CENTER, font_name=C["body_font"])
    _chrome(slide, idx, total, C)

def _slide_recommendation(slide, s, idx, total, C):
    _set_bg(slide, C["bg"])
    _rect(slide, 0, 0.1, W_IN, 0.95, C["subtle"])
    _rect(slide, 0, 0.1, 0.17, 0.95, C["green"])
    _add_text(slide, s.get("title",""), 0.38, 0.14, W_IN-0.6, 0.84,
              size=22, bold=True, color=C["white"], font_name=C["body_font"])
    for i, b in enumerate((s.get("bullets") or [])[:6]):
        y = 1.3 + i * 0.88
        _rect(slide, 0.45, y, W_IN-1.0, 0.76, C["subtle"])
        # Numbered circle (just a coloured rect behind the number)
        _rect(slide, 0.55, y+0.1, 0.5, 0.5, C["green"])
        _add_text(slide, str(i+1), 0.55, y+0.1, 0.5, 0.5,
                  size=14, bold=True, color="000000",
                  align=PP_ALIGN.CENTER, font_name=C["body_font"])
        _add_text(slide, b, 1.18, y+0.1, W_IN-1.8, 0.56,
                  size=12, color=C["white"], font_name=C["body_font"])
    _chrome(slide, idx, total, C)

def _slide_closing(slide, s, idx, total, C):
    _set_bg(slide, C["bg"])
    _rect(slide, 0, H_IN/2 - 0.05, W_IN, 0.1,  C["accent"])
    _rect(slide, 0, 0,              0.28, H_IN,  C["accent"])
    _add_text(slide, s.get("title",""), 0.55, 1.8, 12.0, 1.6,
              size=38, bold=True, color=C["white"], font_name=C["body_font"])
    if s.get("subtitle"):
        _add_text(slide, s["subtitle"], 0.55, 3.6, 11.0, 0.65,
                  size=17, italic=True, color=C["accent"], font_name=C["body_font"])
    if s.get("body"):
        _add_text(slide, s["body"], 0.55, 4.4, 11.0, 1.5,
                  size=12, color=C["midgray"], font_name=C["body_font"])
    _add_text(slide, "SEGA  •  CONFIDENTIAL",
              0.55, H_IN-0.75, 8, 0.38,
              size=9, bold=True, color=C["midgray"], font_name=C["body_font"])
    _chrome(slide, idx, total, C)


def _darken(hex6: str, amount: float = 0.15) -> str:
    h = hex6.lstrip('#').upper().ljust(6,'0')[:6]
    r,g,b = int(h[0:2],16), int(h[2:4],16), int(h[4:6],16)
    return f"{max(0,int(r*(1-amount))):02X}{max(0,int(g*(1-amount))):02X}{max(0,int(b*(1-amount))):02X}"

def _lighten(hex6: str, amount: float = 0.1) -> str:
    h = hex6.lstrip('#').upper().ljust(6,'0')[:6]
    r,g,b = int(h[0:2],16), int(h[2:4],16), int(h[4:6],16)
    return f"{min(255,int(r+(255-r)*amount)):02X}{min(255,int(g+(255-g)*amount)):02X}{min(255,int(b+(255-b)*amount)):02X}"

def extract_template_palette(pptx_bytes: bytes) -> dict:
    """
    Read a .pptx template and return a palette dict compatible with
    the generate_pptx() C dict.  Pulls colours and fonts from the
    theme XML and slide master, then derives text colours based on
    whether the background is dark or light.
    """
    import zipfile as _zf, xml.etree.ElementTree as ET

    ns = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}

    def _lum(h):
        h = (h or '').lstrip('#').upper()
        if len(h) != 6: return 0.5
        r,g,b = int(h[0:2],16), int(h[2:4],16), int(h[4:6],16)
        return (0.299*r + 0.587*g + 0.114*b) / 255

    def _safe(h):
        if not h: return None
        h = h.lstrip('#').upper()
        return h if len(h) == 6 and all(c in '0123456789ABCDEF' for c in h) else None

    zf = _zf.ZipFile(io.BytesIO(pptx_bytes))
    colors, fonts = {}, {}

    # ── Theme XML ────────────────────────────────────────────────────────
    theme_files = [n for n in zf.namelist() if 'ppt/theme/theme' in n and n.endswith('.xml')]
    if theme_files:
        tree = ET.fromstring(zf.read(theme_files[0]))
        clr  = tree.find('.//a:clrScheme', ns)
        if clr is not None:
            for child in clr:
                tag = child.tag.split('}')[-1]
                for cel in child:
                    ctag = cel.tag.split('}')[-1]
                    val  = cel.get('val') or cel.get('lastClr')
                    if val:
                        if ctag == 'srgbClr':
                            colors[tag] = _safe(val)
                        elif ctag == 'sysClr':
                            colors[tag] = 'FFFFFF' if 'window' in val.lower() else '000000'

        fscheme = tree.find('.//a:fontScheme', ns)
        if fscheme is not None:
            for child in fscheme:
                tag   = child.tag.split('}')[-1]
                latin = child.find('a:latin', ns)
                if latin is not None:
                    tf = (latin.get('typeface') or '').strip()
                    if tf and not tf.startswith('+'):
                        fonts['major' if 'major' in tag.lower() else 'minor'] = tf

    # ── Slide master background ──────────────────────────────────────────
    master_bg = None
    master_files = [n for n in zf.namelist() if 'slideMasters/slideMaster' in n and n.endswith('.xml')]
    if master_files:
        mtree = ET.fromstring(zf.read(master_files[0]))
        for solidFill in mtree.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}solidFill'):
            for srgb in solidFill:
                val = _safe(srgb.get('val') or '')
                if val:
                    master_bg = val
                    break
            if master_bg:
                break

    # ── Derive palette ───────────────────────────────────────────────────
    bg      = _safe(master_bg or colors.get('lt1') or colors.get('dk1') or '') or '040A1C'
    is_dark = _lum(bg) < 0.4

    primary = _safe(colors.get('dk2') or colors.get('accent1') or '') or '0055AA'
    # Pick accent that differs from primary
    accent  = _safe(colors.get('accent1') if colors.get('accent1') != primary else
                    colors.get('accent5') or '') or '00AADD'

    if is_dark:
        white, light = 'FFFFFF', 'D0E8FF'
        midgray      = '8899BB'
        subtle       = _darken(bg, 0.12) if _lum(_darken(bg, 0.12)) > 0.05 else _lighten(bg, 0.12)
        header_bg    = _darken(primary, 0.3)
    else:
        white, light = '111122', '334466'
        midgray      = '556677'
        subtle       = _lighten(bg, 0.06)
        header_bg    = primary

    heading_font = fonts.get('major', 'Calibri')
    body_font    = fonts.get('minor', 'Calibri')

    return {
        'bg': bg, 'subtle': subtle, 'header_bg': header_bg,
        'white': white, 'light': light, 'midgray': midgray, 'neutral': midgray,
        'gold':  _safe(colors.get('accent2') or '') or 'F5C242',
        'green': '22DD88', 'red': 'EE3355', 'dark': '111122',
        'primary': primary, 'accent': accent,
        'heading_font': heading_font, 'body_font': body_font,
        'is_dark': is_dark,
    }


def generate_pptx(slide_data: dict, template_palette: dict | None = None) -> bytes:
    """Build a PPTX in memory with python-pptx. Returns raw bytes."""
    from pptx.util import Inches, Emu

    theme = slide_data.get("theme", {})

    def _safe_dark_hex(val, fallback):
        """Accept Claude's hex only if it is a mid/dark vivid colour (lum 0.04–0.82)."""
        if not val or not isinstance(val, str): return fallback
        h = val.lstrip("#").upper()
        if len(h) != 6: return fallback
        try: int(h, 16)
        except ValueError: return fallback
        r, g, b = int(h[0:2],16), int(h[2:4],16), int(h[4:6],16)
        lum = (0.299*r + 0.587*g + 0.114*b) / 255
        return h if 0.04 < lum < 0.82 else fallback

    # Base SEGA defaults
    C = {
        "bg":        "040A1C",
        "subtle":    "1A2A4A",
        "header_bg": "0033AA",
        "white":     "FFFFFF",
        "light":     "D0E4FF",
        "midgray":   "8899BB",
        "neutral":   "AABBCC",
        "gold":      "F5C242",
        "green":     "22DD88",
        "red":       "EE3355",
        "dark":      "040A1C",
        "primary":   _safe_dark_hex(theme.get("primary"), "0055AA"),
        "accent":    _safe_dark_hex(theme.get("accent"),  "00AADD"),
        "heading_font": "Calibri",
        "body_font":    "Calibri",
    }

    # If a template was uploaded, override with its extracted palette
    if template_palette:
        for key in ("bg","subtle","header_bg","white","light","midgray","neutral",
                    "gold","green","red","dark","primary","accent",
                    "heading_font","body_font"):
            if template_palette.get(key):
                C[key] = template_palette[key]

    prs = Presentation()
    # 13.3 × 7.5 inches  (LAYOUT_WIDE equivalent)
    prs.slide_width  = Inches(W_IN)
    prs.slide_height = Inches(H_IN)

    # Blank layout (index 6 is always blank in the default template)
    blank_layout = prs.slide_layouts[6]

    slides_data = slide_data.get("slides", [])
    total       = len(slides_data)

    RENDERERS = {
        "title":          _slide_title,
        "section":        _slide_section,
        "bullets":        _slide_bullets,
        "stats":          _slide_stats,
        "comparison":     _slide_comparison,
        "recommendation": _slide_recommendation,
        "closing":        _slide_closing,
    }

    for idx, s in enumerate(slides_data, start=1):
        slide    = prs.slides.add_slide(blank_layout)
        stype    = (s.get("type") or "bullets").lower()
        renderer = RENDERERS.get(stype, _slide_bullets)
        renderer(slide, s, idx, total, C)

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    # Replace the Office theme with SEGA dark theme so that any text
    # or shape that falls back to theme colours uses white-on-dark,
    # never black-on-dark or white-on-white.
    return _patch_theme(buf.read())



# ─────────────────────────────────────────────────────────────
# API HELPERS  (with rate-limit retry + back-off)
# ─────────────────────────────────────────────────────────────

_RL_WAIT_BASE = 20   # seconds to wait on first 429
_RL_MAX_TRIES = 5    # max retries before giving up


def _api_post(headers: dict, payload: dict, timeout: int = 90,
              on_wait=None) -> dict:
    """
    POST to /v1/messages with exponential back-off on 429.
    on_wait(seconds, attempt) is called before each sleep so callers can
    surface a countdown message to the UI.
    """
    for attempt in range(_RL_MAX_TRIES):
        resp = requests.post(
            "https://api.anthropic.com/v1/messages",
            headers=headers, json=payload, timeout=timeout,
        )
        if resp.status_code == 429:
            if attempt == _RL_MAX_TRIES - 1:
                raise RuntimeError(
                    f"Rate limited after {_RL_MAX_TRIES} retries. "
                    "Try switching to Haiku in the sidebar (much higher rate limit) "
                    "or wait a minute and run again."
                )
            # Honour retry-after header when present; otherwise double each time
            try:
                wait = max(float(
                    resp.headers.get("retry-after") or
                    resp.headers.get("x-ratelimit-reset-requests") or
                    _RL_WAIT_BASE * (2 ** attempt)
                ), 1.0)
            except (TypeError, ValueError):
                wait = _RL_WAIT_BASE * (2 ** attempt)
            if on_wait:
                on_wait(wait, attempt + 1)
            time.sleep(wait)
            continue
        resp.raise_for_status()
        return resp.json()
    raise RuntimeError("Unexpected exit from retry loop.")


def _api_stream(headers: dict, payload: dict, timeout: int = 240,
                on_rate_limit=None):
    """
    POST with stream=True using httpx for a true wall-clock timeout.
    Yields text delta chunks as they arrive via SSE.
    Retries on 429; raises RuntimeError on other errors.
    """
    import httpx

    for attempt in range(_RL_MAX_TRIES):
        stream_payload = {**payload, "stream": True}
        stream_headers = {**headers, "Accept": "text/event-stream"}

        try:
            with httpx.stream(
                "POST",
                "https://api.anthropic.com/v1/messages",
                headers=stream_headers,
                json=stream_payload,
                timeout=httpx.Timeout(timeout, connect=10),
            ) as resp:
                if resp.status_code == 429:
                    if attempt == _RL_MAX_TRIES - 1:
                        raise RuntimeError(
                            "Rate limited after max retries. "
                            "Switch to Haiku in the sidebar or wait a minute."
                        )
                    try:
                        wait = max(float(
                            resp.headers.get("retry-after") or
                            _RL_WAIT_BASE * (2 ** attempt)
                        ), 1.0)
                    except (TypeError, ValueError):
                        wait = _RL_WAIT_BASE * (2 ** attempt)
                    if on_rate_limit:
                        on_rate_limit(wait, attempt + 1)
                    time.sleep(wait)
                    continue

                if resp.status_code != 200:
                    raise RuntimeError(
                        f"API error {resp.status_code}: {resp.text[:400]}"
                    )

                for raw_line in resp.iter_lines():
                    line = raw_line.strip()
                    if not line or not line.startswith("data: "):
                        continue
                    payload_str = line[6:].strip()
                    if payload_str == "[DONE]":
                        return
                    try:
                        evt = json.loads(payload_str)
                    except json.JSONDecodeError:
                        continue
                    if evt.get("type") == "content_block_delta":
                        delta = evt.get("delta", {})
                        if delta.get("type") == "text_delta":
                            yield delta.get("text", "")
                return  # clean exit

        except httpx.TimeoutException:
            raise RuntimeError(
                f"Analysis API timed out after {timeout}s. "
                "Try reducing the slide count or switching to Haiku in the sidebar."
            )



# ─────────────────────────────────────────────────────────────
# PIPELINE  (parallel extraction + research, streaming analysis)
# ─────────────────────────────────────────────────────────────

# Characters of document context to send.  Keeping this under ~8 000 chars
# (~2 000 tokens) leaves plenty of headroom on the 30 000 input-token/min limit.
_MAX_DOC_CHARS = 8_000


def run_pipeline(model, uploaded_files, game_title, business_question, audience,
                 theme_preset, web_search_en, slide_count, template_palette=None):
    """
    Generator yielding (event_type, payload) tuples consumed by the run-button
    handler.  Produces rich narrative log messages so users understand exactly
    what is happening at each stage.
    """
    api_key = st.secrets.get("ANTHROPIC_API_KEY", "")
    if not api_key:
        yield ("error", "ANTHROPIC_API_KEY not found in st.secrets. "
                        "Add it to .streamlit/secrets.toml.")
        return

    headers = {
        "Content-Type": "application/json",
        "x-api-key": api_key,
        "anthropic-version": "2023-06-01",
    }

    # ── STAGE 1 + 2: document extraction & web research in parallel ──────────
    yield ("spinner", (
        "📂 <b>Stage 1 of 4 — Loading your documents</b><br>"
        "<span class='log-detail'>Reading uploaded files and extracting their text content. "
        "PDFs are parsed page-by-page; Excel/CSV sheets are converted to plain text tables; "
        "Word documents are unzipped and their paragraph text pulled out. "
        "Content is capped at {:,} characters to stay within API token limits.</span>"
    ).format(_MAX_DOC_CHARS))

    if web_search_en and game_title:
        yield ("spinner", (
            "🔍 <b>Stage 2 of 4 — Web research running in parallel</b><br>"
            "<span class='log-detail'>While your documents are being extracted, Claude is "
            "simultaneously searching the web for <i>{}</i>. "
            "Two focused searches run in sequence: first critical reception &amp; sales data, "
            "then gameplay mechanics &amp; market context. "
            "Single focused search with a 110s hard timeout — falls back to model knowledge if it times out.</span>"
        ).format(game_title))

    combined_docs  = "[No documents uploaded]"
    research_text  = "[No reference game specified]"

    def _extract_docs():
        if not uploaded_files:
            return "[No documents uploaded]"
        texts = []
        for f in uploaded_files:
            f.seek(0)
            texts.append(f"=== {f.name} ===\n{extract_text_from_file(f)}")
        full = "\n\n".join(texts)
        # Trim to token-safe length
        if len(full) > _MAX_DOC_CHARS:
            full = full[:_MAX_DOC_CHARS] + "\n\n[... document content trimmed to fit token budget ...]"
        return full

    def _web_research():
        """
        Fetch competitive intel via Anthropic web-search tool.
        Uses httpx with a true wall-clock timeout so the call CANNOT
        run longer than HARD_TIMEOUT seconds regardless of socket behaviour.
        """
        if not (web_search_en and game_title):
            if game_title:
                return f"[Web search disabled — using model knowledge for '{game_title}']"
            return "[No reference game specified]"

        import httpx

        HARD_TIMEOUT = 110  # httpx enforces this as a true wall-clock deadline

        prompt = (
            f"Research the video game \"{game_title}\" for an executive competitive analysis presentation. "
            "Search the web for current information and write a thorough structured report with these sections:\n\n"
            "OVERVIEW: Developer, publisher, release date, platforms, genre, ESRB/PEGI rating, launch price.\n\n"
            "CRITICAL RECEPTION: Metacritic score (critic + user), OpenCritic score, "
            "scores from at least 3-4 named outlets (e.g. IGN, Eurogamer, GameSpot). "
            "List 4 specific things reviewers praised and 4 specific things they criticised — "
            "be precise (e.g. 'combat depth', 'open world size', 'performance issues on Switch').\n\n"
            "COMMERCIAL PERFORMANCE: Launch window sales figures, lifetime sales if available, "
            "any sales milestones or statements from the publisher, chart positions.\n\n"
            "GAMEPLAY & FEATURES: 6-7 core mechanics explained in 1-2 sentences each. "
            "Main story length and completionist length. Multiplayer or co-op features. "
            "Accessibility options.\n\n"
            "POST-LAUNCH: Each DLC pack — name, price, release date, brief description, reception. "
            "Major patches or updates. Current player activity signals.\n\n"
            "MARKET CONTEXT: The 3-4 biggest competitor titles released in the same window. "
            "How this game compares to the previous entry in its franchise. "
            "Any notable controversies, marketing moments, or cultural impact.\n\n"
            "Use real numbers throughout. Aim for 700-900 words total."
        )

        payload = {
            "model": model,
            "max_tokens": 2500,
            "tools": [{"type": "web_search_20250305", "name": "web_search"}],
            "messages": [{"role": "user", "content": prompt}],
        }

        try:
            resp = httpx.post(
                "https://api.anthropic.com/v1/messages",
                headers=headers,
                json=payload,
                timeout=httpx.Timeout(HARD_TIMEOUT, connect=10),
            )
            resp.raise_for_status()
            blocks = resp.json().get("content", [])
            text = "\n".join(b.get("text", "") for b in blocks if b.get("type") == "text")
            return text.strip() or (
                f"[No web results — analysis will use model knowledge for '{game_title}']"
            )
        except httpx.TimeoutException:
            return (
                f"[Web search timed out after {HARD_TIMEOUT}s. "
                f"The analysis will use Claude's training knowledge for '{game_title}' instead. "
                "This is normal for popular titles with many search results.]"
            )
        except Exception as e:
            return f"[Web search error: {e}]"

    # ── Poll futures with heartbeat so Streamlit never blocks ────────────────
    # as_completed() would block the generator for up to 120 s with no yields,
    # freezing the UI.  Instead we poll every 3 s and emit a spinner tick so
    # Streamlit keeps receiving events and the log stays alive.
    with ThreadPoolExecutor(max_workers=2) as pool:
        fut_docs     = pool.submit(_extract_docs)
        fut_research = pool.submit(_web_research)

        pending      = {fut_docs: "docs", fut_research: "research"}
        docs_done    = False
        research_done = False
        elapsed       = 0
        TICK          = 3   # seconds between heartbeat yields

        while pending:
            # Check each pending future without blocking
            resolved = []
            for fut, label in list(pending.items()):
                if fut.done():
                    resolved.append((fut, label))

            for fut, label in resolved:
                del pending[fut]
                if label == "docs":
                    docs_done     = True
                    combined_docs = fut.result()
                    n_files = len(uploaded_files)
                    n_chars = len(combined_docs)
                    yield ("log", (
                        "✅ <b>Documents extracted</b> — {} file{}, {:,} characters of content<br>"
                        "<span class='log-detail'>Text successfully pulled from your uploads. "
                        "This content will be the primary source of facts about your internal game.</span>"
                    ).format(n_files, "s" if n_files != 1 else "", n_chars))
                    yield ("step_done", "extract")
                else:
                    research_done = True
                    res = fut.result()
                    is_error   = "[Web search error" in res
                    is_timeout = "timed out after" in res
                    is_fallback = any(x in res for x in (
                        "[Web search disabled", "[No reference", "[No web results",
                        "timed out after", "[No results"
                    ))
                    if is_error and not is_timeout:
                        research_text = res
                        yield ("log", f"⚠️ <b>Web research error</b> — {res}<br>"
                               "<span class='log-detail'>The analysis will proceed using "
                               "Claude's training knowledge for this title.</span>")
                    elif is_fallback:
                        research_text = res
                        yield ("log", (
                            "⚠️ <b>Web search timed out</b> — falling back to model training knowledge "
                            f"for <i>{game_title}</i><br>"
                            "<span class='log-detail'>Claude has extensive knowledge of most "
                            "released games from training data. The analysis will still be "
                            "data-rich; it just won't include the very latest web sources.</span>"
                        ))
                    else:
                        research_text = res
                        word_count    = len(res.split())
                        yield ("log", (
                            "✅ <b>Web research complete</b> — ~{} words on <i>{}</i><br>"
                            "<span class='log-detail'>Claude searched the web and compiled: "
                            "review scores, sales data, gameplay mechanics, and player reception. "
                            "This will be used as the benchmark in the comparison slides.</span>"
                        ).format(word_count, game_title))
                    yield ("step_done", "research")

            if pending:
                # Sleep briefly then emit a heartbeat so the UI stays live
                time.sleep(TICK)
                elapsed += TICK
                still_doing = []
                if not docs_done:
                    still_doing.append("extracting documents")
                if not research_done:
                    still_doing.append(
                        f"searching web for <i>{game_title}</i> ({elapsed}s elapsed)"
                    )
                if still_doing:
                    yield ("spinner", (
                        "⏳ <b>Still working…</b> " + " &amp; ".join(still_doing) + "<br>"
                        "<span class='log-detail'>Web search has a 110s hard limit — if it times out, "
                        "the analysis will automatically use Claude's training knowledge instead. "
                        "The pipeline will not hang.</span>"
                    ))


    # ── STAGE 3: streaming analysis ───────────────────────────────────────────

    # Estimate input token count roughly (1 token ≈ 4 chars)
    prompt_chars = len(combined_docs) + len(research_text) + len(business_question) + 800
    est_input_tokens = prompt_chars // 4
    yield ("spinner", (
        "🤖 <b>Stage 3 of 4 — Claude is writing your presentation</b><br>"
        "<span class='log-detail'>Sending ~{:,} input tokens to {}. Claude will read all your "
        "document content, cross-reference the research on <i>{}</i>, interpret your business "
        "question, and produce a structured {}-slide JSON outline with titles, bullet points, "
        "comparison tables, stat callouts, and speaker notes. "
        "You will see progress updates as each section is written.</span>"
    ).format(est_input_tokens, model, game_title or "the reference game", slide_count))

    theme_desc = {
        "SEGA Blue — Corporate Executive": "Professional SEGA corporate blue (#0055AA), boardroom-ready.",
        "SEGA Dark — Game Reveal Style":   "Dark dramatic (#040A1C) with electric blue accents.",
        "SEGA Sonic — High Energy":        "Vibrant SEGA blue with gold accents, high energy.",
    }.get(theme_preset, "SEGA corporate blue")

    analysis_prompt = f"""You are a senior game industry analyst at SEGA.
Analyse the following and produce a JSON object for a {slide_count}-slide executive presentation.

## INTERNAL GAME DOCUMENTS:
{combined_docs}

## REFERENCE GAME RESEARCH — {game_title}:
{research_text}

## BUSINESS QUESTION:
{business_question}

## AUDIENCE: {audience}

Output a single JSON object. Schema:
{{
  "title":"...", "subtitle":"...",
  "theme":{{"primary":"hex (dark-to-mid blue for SEGA branding, e.g. 0055AA)","accent":"hex (vivid cyan or teal, e.g. 00AADD)"}},
  "slides":[
    {{
      "type":"title|section|comparison|bullets|stats|recommendation|closing",
      "title":"...","subtitle":"...","body":"...",
      "bullets":["..."],
      "stats":[{{"label":"...","value":"...","note":"..."}}],
      "comparison":{{
        "left_title":"Internal Game","right_title":"{game_title}",
        "rows":[{{"label":"...","left":"...","right":"...","delta":"positive|negative|neutral"}}]
      }},
      "speaker_notes":"..."
    }}
  ]
}}

Rules:
- Use REAL data from the documents and research — no generic placeholders
- Be specific and data-driven for {audience}
- theme.primary and theme.accent must be dark-to-mid vivid hex colours (6 digits, no #).
  Never use white, near-white, or light pastels (no values above DDDDDD).
  Good: "0055AA", "003380", "00AADD". Bad: "FFFFFF", "F0F0F0", "C8D8EE".
- Keep speaker_notes to 1-2 sentences maximum — they are brief presenter cues, not essays
- Bullets: max 6 per slide, each under 15 words
- Comparison rows: max 8 per slide
- Return ONLY valid JSON — no markdown fences, no explanation"""

    raw_chunks   = []
    char_count   = 0
    last_tick_at = 0

    # Slide detection: watch for "title": patterns in the accumulating JSON
    # so we can announce each slide as it is written
    slides_announced = 0

    def _count_slides_so_far(text):
        """Count how many slide title fields have appeared so far in the stream."""
        import re
        return len(re.findall(r'"type"\s*:\s*"(?:title|section|bullets|stats|comparison|recommendation|closing)"', text))

    def _on_rate_limit(wait_secs, attempt):
        # This runs in the main thread during the stream retry sleep;
        # we can't yield from here, so we just log via a mutable flag
        pass  # handled below via RuntimeError propagation

    try:
        for chunk in _api_stream(
            headers=headers,
            payload={
                "model": model,
                "max_tokens": 8000,
                "system": "You are a precise game industry analyst. Return valid JSON only.",
                "messages": [{"role": "user", "content": analysis_prompt}],
            },
        ):
            raw_chunks.append(chunk)
            char_count += len(chunk)

            # Announce new slides as they appear in the stream
            current_text  = "".join(raw_chunks)
            slides_so_far = _count_slides_so_far(current_text)
            if slides_so_far > slides_announced:
                for _ in range(slides_so_far - slides_announced):
                    slides_announced += 1
                    yield ("spinner", (
                        f"  ✏️ Writing slide {slides_announced} of {slide_count}…"
                    ))

            # Periodic byte-count tick every ~600 chars
            if char_count - last_tick_at >= 600:
                last_tick_at = char_count
                pct = min(int(char_count / (slide_count * 220) * 100), 95)
                yield ("spinner", (
                    f"  📝 Generating JSON… {char_count:,} chars written (~{pct}% complete)"
                ))

    except RuntimeError as e:
        err = str(e)
        if "rate limit" in err.lower() or "429" in err:
            yield ("error",
                "⏱️ <b>Rate limit hit</b> — your organisation has exceeded 30,000 input tokens/min.<br><br>"
                "Quick fixes:<br>"
                "• Switch to <b>Haiku</b> in the sidebar (much higher rate limits)<br>"
                "• Wait 60 seconds then click Run again<br>"
                "• Reduce the number of uploaded documents to lower input token usage"
            )
        else:
            yield ("error", err)
        return
    except Exception as e:
        yield ("error", f"Streaming error: {e}")
        return

    raw_json = "".join(raw_chunks).strip()
    if raw_json.startswith("```"):
        raw_json = raw_json.split("\n", 1)[1]
        raw_json = raw_json.rsplit("```", 1)[0]

    # ── Truncation recovery ───────────────────────────────────────────────
    # If the stream was cut off mid-JSON (max_tokens hit), try to salvage
    # whatever complete slides were already written before the truncation.
    def _recover_truncated_json(s: str) -> dict | None:
        """
        Try progressively shorter substrings until we get valid JSON,
        or attempt to close open brackets to recover partial output.
        Returns parsed dict or None.
        """
        import re
        # First: try as-is
        try: return json.loads(s)
        except json.JSONDecodeError: pass

        # Second: strip everything after the last complete slide object
        # by finding the last `}` before the slides array closes
        # Look for the last complete slide: ends with either },\n  { or }]\n}
        last_close = s.rfind('}\n    ]')  # end of slides array
        if last_close == -1:
            last_close = s.rfind('}, {')  # between slides
        if last_close == -1:
            last_close = s.rfind('"}')    # end of any string field

        if last_close > 100:
            # Try to close the JSON properly
            truncated = s[:last_close + 1]
            # Count unclosed brackets
            opens  = truncated.count('{') - truncated.count('}')
            opens2 = truncated.count('[') - truncated.count(']')
            candidate = truncated + (']' * opens2) + ('}' * opens)
            try: return json.loads(candidate)
            except json.JSONDecodeError: pass

        return None

    try:
        slide_data = json.loads(raw_json)
    except json.JSONDecodeError as e:
        # Try recovery before giving up
        recovered = _recover_truncated_json(raw_json)
        if recovered and recovered.get("slides"):
            n_recovered = len(recovered["slides"])
            yield ("log", (
                f"⚠️ <b>Output was truncated</b> — recovered {n_recovered} of {slide_count} slides.<br>"
                f"<span class='log-detail'>The model hit the token limit mid-stream. "
                f"Recovered {n_recovered} complete slides. Consider reducing slide count or switching "
                f"to a higher-capacity model.</span>"
            ))
            slide_data = recovered
        else:
            yield ("error",
                f"JSON parse error: {e}<br>"
                f"The model returned malformed JSON and recovery failed. Try running again, "
                f"or reduce the slide count in the sidebar.<br><br>"
                f"First 400 chars received:<br><code>{raw_json[:400]}</code>"
            )
            return

    n_slides = len(slide_data.get("slides", []))
    yield ("step_done", "analyze")
    yield ("log", (
        "✅ <b>Analysis complete</b> — {} slides generated, {:,} characters of content<br>"
        "<span class='log-detail'>Claude has finished writing all slide content including "
        "titles, bullet points, comparison rows, stat callouts, and speaker notes. "
        "Slide types used: {}.</span>"
    ).format(
        n_slides,
        char_count,
        ", ".join(sorted(set(s.get("type","?") for s in slide_data.get("slides",[]))))
    ))
    yield ("slide_data", slide_data)

    # ── STAGE 4: PPTX rendering ───────────────────────────────────────────────
    yield ("spinner", (
        "🖥️ <b>Stage 4 of 4 — Building the PowerPoint file</b><br>"
        "<span class='log-detail'>Rendering {} slides with python-pptx{}. "
        "Setting backgrounds, drawing shape layers, placing text with correct fonts, "
        "adding chrome footer with page numbers, writing to memory.</span>"
    ).format(
        n_slides,
        " using your uploaded template's colour palette and fonts" if template_palette else " with SEGA dark theme",
    ))

    try:
        pptx_bytes_out = generate_pptx(slide_data, template_palette=template_palette)
    except Exception as e:
        yield ("error", f"PPTX generation error: {e}")
        return

    yield ("step_done", "generate")
    yield ("log", (
        "🎉 <b>All done!</b> — PPTX is {:.0f} KB across {} slides<br>"
        "<span class='log-detail'>Your presentation is ready to download. "
        "Open it in PowerPoint or Google Slides. Speaker notes are included on each slide. "
        "The file uses standard OOXML format and is compatible with all modern presentation apps.</span>"
    ).format(len(pptx_bytes_out) / 1024, n_slides))
    yield ("pptx_bytes_out", pptx_bytes_out)


# ─────────────────────────────────────────────────────────────
# RUN BUTTON HANDLER
# ─────────────────────────────────────────────────────────────
if run_btn:
    if not business_question.strip():
        st.error("Please enter a business question.")
    else:
        # Extract template palette if a file was uploaded
        _template_palette = None
        _template_file = st.session_state.get("template_upload")
        if _template_file is not None:
            try:
                _template_file.seek(0)
                _template_palette = extract_template_palette(_template_file.read())
            except Exception as _e:
                st.warning(f"Could not read template palette: {_e}. Using default theme.")

        pipeline_steps = {
            "upload": bool(uploaded_files), "extract": False,
            "research": False, "analyze": False, "generate": False,
        }
        st.session_state["pipeline_steps"] = pipeline_steps
        log_lines = []

        with output_area.container():
            st.markdown('<div class="section-label">Pipeline log</div>', unsafe_allow_html=True)
            log_area = st.empty()

            try:
                for event in run_pipeline(
                    model, uploaded_files, game_title, business_question,
                    audience, theme_preset, web_search_enabled, slide_count,
                    template_palette=_template_palette,
                ):
                    etype = event[0]

                    if etype in ("log", "spinner"):
                        # "spinner" = active working entry (animated ring at bottom)
                        # "log"     = completed entry (static, left-border style)
                        # When a new event arrives, any previous spinner is promoted
                        # to a plain log entry so only the latest one animates.
                        if etype == "spinner":
                            # Promote the last spinner (if any) to a plain log entry
                            if log_lines and log_lines[-1][0] == "spinner":
                                log_lines[-1] = ("log", log_lines[-1][1])
                            log_lines.append(("spinner", event[1]))
                        else:
                            # Promote any trailing spinner before adding the completed msg
                            if log_lines and log_lines[-1][0] == "spinner":
                                log_lines[-1] = ("log", log_lines[-1][1])
                            log_lines.append(("log", event[1]))

                        # Build self-contained HTML with inline CSS.
                        # Streamlit does NOT share CSS between separate markdown() calls,
                        # so all styles must be embedded in every render.
                        _CSS = (
                            "<style>"
                            "@keyframes _sp{to{transform:rotate(360deg)}}"
                            "._lw{background:#0f172a;border:1px solid #1e293b;border-radius:8px;"
                            "padding:1rem 1.25rem;font-size:.82rem;"
                            "font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;"
                            "color:#94a3b8;max-height:480px;overflow-y:auto;line-height:1.7}"
                            "._lw b{color:#e2e8f0;font-weight:600}"
                            "._lw i{color:#60a5fa}"
                            "._ld{color:#64748b;font-size:.76rem;line-height:1.55;display:block;"
                            "margin-top:.15rem;margin-bottom:.35rem}"
                            "._le{border-left:2px solid #1e3a5f;padding-left:.6rem;margin-bottom:.5rem}"
                            "._la{border-left:2px solid #00AADD;padding-left:.6rem;"
                            "margin-bottom:.5rem;display:flex;align-items:flex-start;gap:.55rem}"
                            "._lr{flex-shrink:0;width:14px;height:14px;margin-top:3px;"
                            "border:2px solid rgba(0,170,221,.25);border-top-color:#00AADD;"
                            "border-radius:50%;animation:_sp .8s linear infinite}"
                            "._lt{flex:1}"
                            "</style>"
                        )
                        html_parts = [_CSS, '<div class="_lw">']
                        for _kind, _text in log_lines[-14:]:
                            _t = _text.replace("class='log-detail'", "class='_ld'")
                            if _kind == "spinner":
                                html_parts.append(
                                    f'<div class="_la"><div class="_lr"></div>'
                                    f'<div class="_lt">{_t}</div></div>'
                                )
                            else:
                                html_parts.append(f'<div class="_le">{_t}</div>')
                        html_parts.append("</div>")
                        log_area.markdown("".join(html_parts), unsafe_allow_html=True)
                    elif etype == "step_done":
                        pipeline_steps[event[1]] = True
                        st.session_state["pipeline_steps"] = pipeline_steps

                    elif etype == "slide_data":
                        st.session_state["slide_data"] = event[1]

                    elif etype == "pptx_bytes_out":
                        fname = f"SEGA_Analysis_{(game_title or 'Report').replace(' ','_')}.pptx"
                        st.session_state["pptx_bytes"]    = event[1]
                        st.session_state["pptx_filename"] = fname

                    elif etype == "error":
                        st.error(event[1])
                        break

            except Exception as ex:
                st.error(f"Unexpected error: {ex}")
                import traceback
                st.code(traceback.format_exc())

        if "pptx_bytes" in st.session_state:
            with download_area.container():
                st.success("Analysis complete.")
                st.download_button(
                    label="⬇️ Download PPTX",
                    data=st.session_state["pptx_bytes"],
                    file_name=st.session_state["pptx_filename"],
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    use_container_width=True,
                )
                if "slide_data" in st.session_state:
                    with st.expander("Slide outline", expanded=False):
                        for i, sl in enumerate(st.session_state["slide_data"].get("slides", []), 1):
                            st.markdown(f"**{i}.** `{sl.get('type','?').upper()}` — {sl.get('title','')}")