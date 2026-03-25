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
    padding: 1rem 1.25rem; font-size: .78rem;
    font-family: "SF Mono","Fira Code",monospace;
    color: #94a3b8; max-height: 320px; overflow-y: auto; line-height: 1.8;
}
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
              font_name="Calibri")
    _add_text(slide, f"{idx} / {total}",
              W_IN-1.4, H_IN-0.30, 1.1, 0.24,
              size=8, color=C["neutral"], align=PP_ALIGN.RIGHT,
              font_name="Calibri")
    _rect(slide, W_IN-0.1, 0, 0.1, H_IN, C["primary"])

def _set_bg(slide, hex_color):
    bg   = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = _rgb(hex_color)

def _add_bullets(slide, bullets, x, y, w, h, bullet_color, text_color,
                 size=13, font_name="Calibri"):
    if not bullets:
        return
    txb = slide.shapes.add_textbox(_in(x), _in(y), _in(w), _in(h))
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
              size=9, color=C["accent"], align=PP_ALIGN.CENTER, font_name="Calibri")
    # Main title
    _add_text(slide, s.get("title",""), 0.6, 1.8, 6.8, 2.2,
              size=34, bold=True, color=C["white"], font_name="Calibri")
    if s.get("subtitle"):
        _add_text(slide, s["subtitle"], 0.6, 4.2, 6.6, 0.6,
                  size=16, italic=True, color=C["accent"], font_name="Calibri")
    if s.get("body"):
        _add_text(slide, s["body"], 0.6, 4.95, 6.6, 1.4,
                  size=12, color=C["midgray"], font_name="Calibri")
    _chrome(slide, idx, total, C)

def _slide_section(slide, s, idx, total, C):
    _set_bg(slide, C["bg"])
    _rect(slide, 0, 0, 0.28, H_IN, C["accent"])
    _add_text(slide, s.get("title",""), 0.55, 2.3, 12.0, 1.6,
              size=40, bold=True, color=C["white"], font_name="Calibri")
    if s.get("subtitle"):
        _add_text(slide, s["subtitle"], 0.55, 4.1, 11.0, 0.65,
                  size=19, italic=True, color=C["accent"], font_name="Calibri")
    _chrome(slide, idx, total, C)

def _slide_bullets(slide, s, idx, total, C):
    _set_bg(slide, C["bg"])
    _rect(slide, 0, 0.1, W_IN, 0.95, C["subtle"])
    _rect(slide, 0, 0.1, 0.17, 0.95, C["accent"])
    _add_text(slide, s.get("title",""), 0.38, 0.14, W_IN-0.6, 0.84,
              size=22, bold=True, color=C["white"], font_name="Calibri")
    _add_bullets(slide, s.get("bullets",[]),
                 0.45, 1.2, W_IN-1.0, H_IN-2.0,
                 bullet_color=C["accent"], text_color=C["white"],
                 size=13)
    if s.get("body"):
        _add_text(slide, s["body"], 0.45, H_IN-1.45, W_IN-1.0, 1.0,
                  size=11, italic=True, color=C["midgray"], font_name="Calibri")
    _chrome(slide, idx, total, C)

def _slide_stats(slide, s, idx, total, C):
    _set_bg(slide, C["bg"])
    _rect(slide, 0, 0.1, W_IN, 0.95, C["subtle"])
    _rect(slide, 0, 0.1, 0.17, 0.95, C["gold"])
    _add_text(slide, s.get("title",""), 0.38, 0.14, W_IN-0.6, 0.84,
              size=22, bold=True, color=C["white"], font_name="Calibri")
    stats = s.get("stats", [])[:4]
    if stats:
        box_w = (W_IN - 1.2) / len(stats)
        for i, stat in enumerate(stats):
            x = 0.6 + i * (box_w + 0.1)
            _rect(slide, x, 1.3, box_w, 2.9, C["subtle"])
            _add_text(slide, stat.get("value","—"),
                      x+0.1, 1.55, box_w-0.2, 1.2,
                      size=36, bold=True, color=C["accent"],
                      align=PP_ALIGN.CENTER, font_name="Calibri")
            _add_text(slide, stat.get("label",""),
                      x+0.1, 2.85, box_w-0.2, 0.7,
                      size=12, bold=True, color=C["white"],
                      align=PP_ALIGN.CENTER, font_name="Calibri")
            if stat.get("note"):
                _add_text(slide, stat["note"],
                          x+0.1, 3.6, box_w-0.2, 0.45,
                          size=9, color=C["midgray"],
                          align=PP_ALIGN.CENTER, font_name="Calibri")
    if s.get("body"):
        _add_text(slide, s["body"], 0.45, H_IN-1.5, W_IN-1.0, 1.0,
                  size=11, italic=True, color=C["midgray"], font_name="Calibri")
    _chrome(slide, idx, total, C)

def _slide_comparison(slide, s, idx, total, C):
    _set_bg(slide, C["bg"])
    _rect(slide, 0, 0.1, W_IN, 0.95, C["subtle"])
    _rect(slide, 0, 0.1, 0.17, 0.95, C["accent"])
    _add_text(slide, s.get("title",""), 0.38, 0.14, W_IN-0.6, 0.84,
              size=20, bold=True, color=C["white"], font_name="Calibri")
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
                  size=9, bold=True, color=C["midgray"], font_name="Calibri")
        _add_text(slide, row.get("left","—"),
                  mid_x+0.05, y, col_w-0.1, row_h,
                  size=9, color=C["white"], align=PP_ALIGN.CENTER, font_name="Calibri")
        delta = (row.get("delta","") or "").lower()
        dc = C["green"] if delta == "positive" else C["red"] if delta == "negative" else C["neutral"]
        _add_text(slide, row.get("right","—"),
                  right_x+0.05, y, col_w-0.1, row_h,
                  size=9, color=dc, align=PP_ALIGN.CENTER, font_name="Calibri")
    _chrome(slide, idx, total, C)

def _slide_recommendation(slide, s, idx, total, C):
    _set_bg(slide, C["bg"])
    _rect(slide, 0, 0.1, W_IN, 0.95, C["subtle"])
    _rect(slide, 0, 0.1, 0.17, 0.95, C["green"])
    _add_text(slide, s.get("title",""), 0.38, 0.14, W_IN-0.6, 0.84,
              size=22, bold=True, color=C["white"], font_name="Calibri")
    for i, b in enumerate((s.get("bullets") or [])[:6]):
        y = 1.3 + i * 0.88
        _rect(slide, 0.45, y, W_IN-1.0, 0.76, C["subtle"])
        # Numbered circle (just a coloured rect behind the number)
        _rect(slide, 0.55, y+0.1, 0.5, 0.5, C["green"])
        _add_text(slide, str(i+1), 0.55, y+0.1, 0.5, 0.5,
                  size=14, bold=True, color="000000",
                  align=PP_ALIGN.CENTER, font_name="Calibri")
        _add_text(slide, b, 1.18, y+0.1, W_IN-1.8, 0.56,
                  size=12, color=C["white"], font_name="Calibri")
    _chrome(slide, idx, total, C)

def _slide_closing(slide, s, idx, total, C):
    _set_bg(slide, C["bg"])
    _rect(slide, 0, H_IN/2 - 0.05, W_IN, 0.1,  C["accent"])
    _rect(slide, 0, 0,              0.28, H_IN,  C["accent"])
    _add_text(slide, s.get("title",""), 0.55, 1.8, 12.0, 1.6,
              size=38, bold=True, color=C["white"], font_name="Calibri")
    if s.get("subtitle"):
        _add_text(slide, s["subtitle"], 0.55, 3.6, 11.0, 0.65,
                  size=17, italic=True, color=C["accent"], font_name="Calibri")
    if s.get("body"):
        _add_text(slide, s["body"], 0.55, 4.4, 11.0, 1.5,
                  size=12, color=C["midgray"], font_name="Calibri")
    _add_text(slide, "SEGA  •  CONFIDENTIAL",
              0.55, H_IN-0.75, 8, 0.38,
              size=9, bold=True, color=C["midgray"], font_name="Calibri")
    _chrome(slide, idx, total, C)


def generate_pptx(slide_data: dict) -> bytes:
    """Build a PPTX in memory with python-pptx. Returns raw bytes."""
    from pptx.util import Inches, Emu

    theme = slide_data.get("theme", {})
    C = {
        "bg":        theme.get("background", "040A1C"),
        "primary":   theme.get("primary",    "0055AA"),
        "accent":    theme.get("accent",     "00AADD"),
        "white":     theme.get("text_light", "FFFFFF"),
        "dark":      theme.get("text_dark",  "0A0A1A"),
        "gold":      "F5C842",
        "subtle":    "1A2A4A",
        "midgray":   "8899BB",
        "green":     "00BB66",
        "red":       "CC2244",
        "neutral":   "AABBCC",
        "header_bg": "0033AA",
    }

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
    return buf.read()



# ─────────────────────────────────────────────────────────────
# API HELPERS
# ─────────────────────────────────────────────────────────────

def _api_post(headers: dict, payload: dict, timeout: int = 90) -> dict:
    """Single POST with rate-limit retry (max 3 attempts, exponential back-off)."""
    for attempt in range(3):
        resp = requests.post(
            "https://api.anthropic.com/v1/messages",
            headers=headers, json=payload, timeout=timeout,
        )
        if resp.status_code == 429:
            wait = float(resp.headers.get("retry-after") or (10 * 2 ** attempt))
            time.sleep(wait)
            continue
        resp.raise_for_status()
        return resp.json()
    raise RuntimeError("Rate limited after 3 retries. Try again shortly.")


def _api_stream(headers: dict, payload: dict, timeout: int = 120):
    """
    POST with stream=True. Yields text chunks as they arrive (SSE parsing).
    Caller collects chunks; raises RuntimeError on HTTP errors.
    """
    stream_headers = {**headers, "Accept": "text/event-stream"}
    payload = {**payload, "stream": True}
    with requests.post(
        "https://api.anthropic.com/v1/messages",
        headers=stream_headers, json=payload,
        timeout=timeout, stream=True,
    ) as resp:
        if resp.status_code != 200:
            raise RuntimeError(f"API error {resp.status_code}: {resp.text[:300]}")
        for raw_line in resp.iter_lines():
            if not raw_line:
                continue
            line = raw_line.decode("utf-8") if isinstance(raw_line, bytes) else raw_line
            if not line.startswith("data: "):
                continue
            payload_str = line[6:].strip()
            if payload_str == "[DONE]":
                break
            try:
                evt = json.loads(payload_str)
            except json.JSONDecodeError:
                continue
            if evt.get("type") == "content_block_delta":
                delta = evt.get("delta", {})
                if delta.get("type") == "text_delta":
                    yield delta.get("text", "")


# ─────────────────────────────────────────────────────────────
# PIPELINE  (parallel extraction + research, streaming analysis)
# ─────────────────────────────────────────────────────────────

def run_pipeline(model, uploaded_files, game_title, business_question, audience,
                 theme_preset, web_search_en, slide_count):
    """
    Generator yielding (event_type, payload) tuples.

    Speed optimisations vs. the previous version:
      1. Document extraction and web research run in PARALLEL via ThreadPoolExecutor.
      2. The analysis call uses SSE streaming so text appears token-by-token.
      3. max_tokens reduced to 4 000 (enough for 10-slide JSON; was 8 000).
      4. Web-search summary capped at 1 500 tokens (was 3 000) — we only need facts.
    """
    api_key = st.secrets.get("ANTHROPIC_API_KEY", "")
    if not api_key:
        yield ("error", "ANTHROPIC_API_KEY not found in st.secrets.")
        return

    headers = {
        "Content-Type": "application/json",
        "x-api-key": api_key,
        "anthropic-version": "2023-06-01",
    }

    # ── Step 1 + 2 in parallel ────────────────────────────────
    yield ("log", "📄 Extracting documents & searching web in parallel…")

    combined_docs   = "[No documents uploaded]"
    research_text   = "[No reference game specified]"

    def _extract_docs():
        texts = []
        for f in uploaded_files:
            f.seek(0)
            texts.append(f"=== {f.name} ===\n{extract_text_from_file(f)}")
        return "\n\n".join(texts) if texts else "[No documents uploaded]"

    def _web_research():
        if not (web_search_en and game_title):
            if game_title:
                return f"[Web search disabled — using model knowledge for '{game_title}']"
            return "[No reference game specified]"
        try:
            data = _api_post(
                headers=headers,
                payload={
                    "model": model,
                    "max_tokens": 1500,       # ← was 3 000; halved
                    "tools": [{"type": "web_search_20250305", "name": "web_search"}],
                    "messages": [{"role": "user", "content": (
                        f"Research '{game_title}' for a game business analysis. "
                        "Give a concise structured summary covering: genre, platforms, "
                        "developer, release date, Metacritic score, sales estimates, "
                        "key mechanics, player reception highlights, and any DLC/post-launch. "
                        "Be brief and factual — 400 words max."
                    )}],
                },
                timeout=60,
            )
            blocks = data.get("content", [])
            return "\n".join(b.get("text","") for b in blocks if b.get("type")=="text")
        except Exception as e:
            return f"[Web search error: {e}]"

    with ThreadPoolExecutor(max_workers=2) as pool:
        fut_docs     = pool.submit(_extract_docs)
        fut_research = pool.submit(_web_research)

        # Surface completions as they finish
        for fut in as_completed([fut_docs, fut_research]):
            if fut is fut_docs:
                combined_docs = fut.result()
                yield ("log", f"✅ Documents extracted ({len(combined_docs):,} chars)")
                yield ("step_done", "extract")
            else:
                res = fut.result()
                if res.startswith("[Web search error"):
                    yield ("log", f"⚠️ {res}")
                else:
                    research_text = res
                    yield ("log", "✅ Web research complete")
                yield ("step_done", "research")

    # ── Step 3: streaming analysis ────────────────────────────
    yield ("log", "🤖 Analysing with Claude (streaming)…")

    theme_desc = {
        "SEGA Blue — Corporate Executive": "Professional SEGA corporate blue (#0055AA) — boardroom-ready.",
        "SEGA Dark — Game Reveal Style":   "Dark dramatic (#040A1C) with electric blue accents.",
        "SEGA Sonic — High Energy":        "Vibrant SEGA blue + gold accents, dynamic energy.",
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
  "theme":{{"primary":"hex","accent":"hex","background":"hex","text_light":"hex","text_dark":"hex"}},
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
- Use REAL data from the documents and research
- Be specific and data-driven for {audience}
- Return ONLY valid JSON — no markdown fences, no explanation"""

    raw_chunks = []
    token_count = [0]

    try:
        for chunk in _api_stream(
            headers=headers,
            payload={
                "model": model,
                "max_tokens": 4000,     # ← was 8 000; halved
                "system": "You are a precise game industry analyst. Return valid JSON only.",
                "messages": [{"role": "user", "content": analysis_prompt}],
            },
        ):
            raw_chunks.append(chunk)
            token_count[0] += 1
            # Surface a progress tick every 40 tokens so the UI feels live
            if token_count[0] % 40 == 0:
                yield ("log", f"🤖 Generating slide content… ({token_count[0]} tokens)")

    except RuntimeError as e:
        yield ("error", str(e))
        return
    except Exception as e:
        yield ("error", f"API streaming error: {e}")
        return

    raw_json = "".join(raw_chunks).strip()
    # Strip accidental markdown fences
    if raw_json.startswith("```"):
        raw_json = raw_json.split("\n", 1)[1]
        raw_json = raw_json.rsplit("```", 1)[0]

    try:
        slide_data = json.loads(raw_json)
    except json.JSONDecodeError as e:
        yield ("error", f"JSON parse error: {e}\n\nFirst 400 chars:\n{raw_json[:400]}")
        return

    n_slides = len(slide_data.get("slides", []))
    yield ("step_done", "analyze")
    yield ("log", f"✅ Analysis complete — {n_slides} slides, {token_count[0]} tokens")
    yield ("slide_data", slide_data)

    # ── Step 4: PPTX generation ───────────────────────────────
    yield ("log", "🖥️ Building PPTX…")
    try:
        pptx_bytes_out = generate_pptx(slide_data)
    except Exception as e:
        yield ("error", f"PPTX generation error: {e}")
        return

    yield ("step_done", "generate")
    yield ("log", f"✅ Done — {len(pptx_bytes_out):,} bytes")
    yield ("pptx_bytes_out", pptx_bytes_out)


# ─────────────────────────────────────────────────────────────
# RUN BUTTON HANDLER
# ─────────────────────────────────────────────────────────────
if run_btn:
    if not business_question.strip():
        st.error("Please enter a business question.")
    else:
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
                ):
                    etype = event[0]

                    if etype == "log":
                        log_lines.append(event[1])
                        log_area.markdown(
                            '<div class="result-log">' + "<br>".join(log_lines[-12:]) + "</div>",
                            unsafe_allow_html=True,
                        )

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