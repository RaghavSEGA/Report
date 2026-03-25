import streamlit as st
import anthropic
import json
import os
import tempfile
import subprocess
import base64
import pandas as pd
import pypdf
import io
import time
import hashlib
import hmac
import random
import requests

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


def build_js_script(data_file: str, output_file: str) -> str:
    return f"""
const pptxgen = require("pptxgenjs");
const fs = require("fs");

const raw = fs.readFileSync("{data_file}", "utf8");
const data = JSON.parse(raw);
const theme = data.theme || {{}};

const C = {{
  bg:       theme.background || "040A1C",
  primary:  theme.primary    || "0055AA",
  accent:   theme.accent     || "00AADD",
  white:    theme.text_light || "FFFFFF",
  dark:     theme.text_dark  || "0A0A1A",
  gold:     "F5C842",  subtle:"1A2A4A",  midgray:"8899BB",
  green:    "00BB66",  red:"CC2244",     neutral:"AABBCC",
  header_bg:"0033AA",
}};

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.title  = data.title || "SEGA Analysis";
const W = 13.3, H = 7.5;

function addChrome(slide, n, total) {{
  slide.addShape(pres.shapes.RECTANGLE,{{x:0,y:0,w:W,h:0.12,fill:{{color:C.accent}},line:{{type:"none"}}}});
  slide.addShape(pres.shapes.RECTANGLE,{{x:0,y:H-0.38,w:W,h:0.38,fill:{{color:C.header_bg}},line:{{type:"none"}}}});
  slide.addText("SEGA INTELLIGENCE ANALYZER",{{x:0.3,y:H-0.32,w:5,h:0.28,fontSize:7,color:"8899BB",fontFace:"Calibri",align:"left",valign:"middle",bold:true,charSpacing:3}});
  slide.addText(`${{n}} / ${{total}}`,{{x:W-1.5,y:H-0.32,w:1.2,h:0.28,fontSize:8,color:C.neutral,fontFace:"Calibri",align:"right",valign:"middle"}});
  slide.addShape(pres.shapes.RECTANGLE,{{x:W-0.12,y:0,w:0.12,h:H,fill:{{color:C.primary}},line:{{type:"none"}}}});
}}

function makeTitleSlide(sl,s,i,t){{
  sl.background={{color:C.bg}};
  sl.addShape(pres.shapes.RECTANGLE,{{x:7.5,y:0,w:5.8,h:H,fill:{{color:C.primary}},line:{{type:"none"}}}});
  sl.addShape(pres.shapes.RECTANGLE,{{x:6.8,y:0,w:1.0,h:H,fill:{{color:C.accent,transparency:60}},line:{{type:"none"}}}});
  sl.addText("SEGA",{{x:8,y:0.6,w:4.8,h:0.9,fontSize:52,bold:true,color:C.white,fontFace:"Arial Black",align:"center",charSpacing:12}});
  sl.addShape(pres.shapes.RECTANGLE,{{x:8.3,y:1.55,w:4.2,h:0.04,fill:{{color:C.accent}},line:{{type:"none"}}}});
  sl.addText("INTELLIGENCE ANALYZER",{{x:8,y:1.6,w:4.8,h:0.4,fontSize:9,color:C.accent,fontFace:"Calibri",align:"center",charSpacing:4}});
  sl.addText(s.title||pres.title,{{x:0.6,y:1.8,w:6.8,h:2.4,fontSize:36,bold:true,color:C.white,fontFace:"Calibri",align:"left",valign:"middle",wrap:true}});
  if(s.subtitle) sl.addText(s.subtitle,{{x:0.6,y:4.3,w:6.6,h:0.6,fontSize:16,color:C.accent,fontFace:"Calibri",italic:true}});
  if(s.body)     sl.addText(s.body,{{x:0.6,y:5.0,w:6.6,h:1.5,fontSize:13,color:C.midgray,fontFace:"Calibri",wrap:true}});
  addChrome(sl,i+1,t);
}}

function makeSectionSlide(sl,s,i,t){{
  sl.background={{color:C.bg}};
  sl.addShape(pres.shapes.RECTANGLE,{{x:0,y:0,w:0.3,h:H,fill:{{color:C.accent}},line:{{type:"none"}}}});
  sl.addText(s.title,{{x:0.9,y:2.5,w:11,h:1.5,fontSize:42,bold:true,color:C.white,fontFace:"Calibri"}});
  if(s.subtitle) sl.addText(s.subtitle,{{x:0.9,y:4.2,w:10,h:0.6,fontSize:20,color:C.accent,fontFace:"Calibri",italic:true}});
  addChrome(sl,i+1,t);
}}

function makeBulletsSlide(sl,s,i,t){{
  sl.background={{color:C.bg}};
  sl.addShape(pres.shapes.RECTANGLE,{{x:0,y:0.12,w:W,h:1.0,fill:{{color:C.subtle}},line:{{type:"none"}}}});
  sl.addShape(pres.shapes.RECTANGLE,{{x:0,y:0.12,w:0.18,h:1.0,fill:{{color:C.accent}},line:{{type:"none"}}}});
  sl.addText(s.title,{{x:0.4,y:0.18,w:W-0.6,h:0.88,fontSize:24,bold:true,color:C.white,fontFace:"Calibri",valign:"middle"}});
  const bullets=(s.bullets||[]);
  const items=bullets.map((b,j)=>[
    {{text:"▸ ",options:{{color:C.accent,bold:true,fontSize:15}}}},
    {{text:b,   options:{{color:C.white, fontSize:15,breakLine:j<bullets.length-1}}}}
  ]).flat();
  if(items.length) sl.addText(items,{{x:0.5,y:1.4,w:W-1.2,h:H-2.1,fontFace:"Calibri",valign:"top",wrap:true,paraSpaceAfter:10}});
  if(s.body) sl.addText(s.body,{{x:0.5,y:H-1.5,w:W-1.0,h:1.0,fontSize:12,color:C.midgray,fontFace:"Calibri",italic:true,wrap:true}});
  addChrome(sl,i+1,t);
}}

function makeStatsSlide(sl,s,i,t){{
  sl.background={{color:C.bg}};
  sl.addShape(pres.shapes.RECTANGLE,{{x:0,y:0.12,w:W,h:1.0,fill:{{color:C.subtle}},line:{{type:"none"}}}});
  sl.addShape(pres.shapes.RECTANGLE,{{x:0,y:0.12,w:0.18,h:1.0,fill:{{color:C.gold}},line:{{type:"none"}}}});
  sl.addText(s.title,{{x:0.4,y:0.18,w:W-0.6,h:0.88,fontSize:24,bold:true,color:C.white,fontFace:"Calibri",valign:"middle"}});
  const stats=s.stats||[];
  const cols=Math.min(stats.length,4), boxW=(W-1.2)/cols;
  stats.forEach((stat,j)=>{{
    if(j>=4)return;
    const x=0.6+j*(boxW+0.12);
    sl.addShape(pres.shapes.RECTANGLE,{{x,y:1.4,w:boxW,h:2.8,fill:{{color:C.subtle}},line:{{color:C.primary,pt:1}}}});
    sl.addText(stat.value||"—",{{x:x+0.1,y:1.7, w:boxW-0.2,h:1.2,fontSize:38,bold:true,color:C.accent, fontFace:"Calibri",align:"center",valign:"middle"}});
    sl.addText(stat.label||"",{{x:x+0.1,y:2.95,w:boxW-0.2,h:0.7,fontSize:13,color:C.white,  fontFace:"Calibri",align:"center",wrap:true,bold:true}});
    if(stat.note) sl.addText(stat.note,{{x:x+0.1,y:3.65,w:boxW-0.2,h:0.45,fontSize:9,color:C.midgray,fontFace:"Calibri",align:"center",wrap:true}});
  }});
  if(s.body) sl.addText(s.body,{{x:0.5,y:H-1.6,w:W-1.0,h:1.0,fontSize:12,color:C.midgray,fontFace:"Calibri",italic:true,wrap:true}});
  addChrome(sl,i+1,t);
}}

function makeComparisonSlide(sl,s,i,t){{
  sl.background={{color:C.bg}};
  sl.addShape(pres.shapes.RECTANGLE,{{x:0,y:0.12,w:W,h:1.0,fill:{{color:C.subtle}},line:{{type:"none"}}}});
  sl.addShape(pres.shapes.RECTANGLE,{{x:0,y:0.12,w:0.18,h:1.0,fill:{{color:C.accent}},line:{{type:"none"}}}});
  sl.addText(s.title,{{x:0.4,y:0.18,w:W-0.6,h:0.88,fontSize:22,bold:true,color:C.white,fontFace:"Calibri",valign:"middle"}});
  const cmp=s.comparison||{{}},rows=cmp.rows||[];
  const colW=3.8,labelW=2.8,leftX=0.5,midX=leftX+labelW+0.1,rightX=midX+colW+0.1,startY=1.35,rowH=0.48;
  [[cmp.left_title||"Internal",C.primary],[cmp.right_title||"Reference",C.gold]].forEach(([title,col],ci)=>{{
    const x=ci===0?midX:rightX;
    sl.addShape(pres.shapes.RECTANGLE,{{x,y:startY,w:colW,h:0.42,fill:{{color:col}},line:{{type:"none"}}}});
    sl.addText(title,{{x:x+0.1,y:startY,w:colW-0.2,h:0.42,fontSize:11,bold:true,color:C.white,fontFace:"Calibri",align:"center",valign:"middle"}});
  }});
  rows.slice(0,10).forEach((row,ri)=>{{
    const y=startY+0.44+ri*rowH;
    sl.addShape(pres.shapes.RECTANGLE,{{x:leftX,y,w:labelW+colW*2+0.22,h:rowH-0.04,fill:{{color:ri%2===0?"0D1530":"111E3A"}},line:{{type:"none"}}}});
    sl.addText(row.label||"",{{x:leftX+0.1,y,w:labelW-0.2,h:rowH,fontSize:10,color:C.midgray,fontFace:"Calibri",valign:"middle",bold:true}});
    sl.addText(row.left||"—",{{x:midX+0.1,y,w:colW-0.2,h:rowH,fontSize:10,color:C.white,fontFace:"Calibri",align:"center",valign:"middle",wrap:true}});
    const dc=(row.delta||"").toLowerCase()==="positive"?C.green:(row.delta||"").toLowerCase()==="negative"?C.red:C.neutral;
    sl.addText(row.right||"—",{{x:rightX+0.1,y,w:colW-0.2,h:rowH,fontSize:10,color:dc,fontFace:"Calibri",align:"center",valign:"middle",wrap:true}});
  }});
  addChrome(sl,i+1,t);
}}

function makeRecommendationSlide(sl,s,i,t){{
  sl.background={{color:C.bg}};
  sl.addShape(pres.shapes.RECTANGLE,{{x:0,y:0.12,w:W,h:1.0,fill:{{color:C.subtle}},line:{{type:"none"}}}});
  sl.addShape(pres.shapes.RECTANGLE,{{x:0,y:0.12,w:0.18,h:1.0,fill:{{color:C.green}},line:{{type:"none"}}}});
  sl.addText(s.title,{{x:0.4,y:0.18,w:W-0.6,h:0.88,fontSize:24,bold:true,color:C.white,fontFace:"Calibri",valign:"middle"}});
  (s.bullets||[]).slice(0,6).forEach((b,j)=>{{
    const y=1.5+j*0.9;
    sl.addShape(pres.shapes.RECTANGLE,{{x:0.5,y,w:W-1.2,h:0.78,fill:{{color:C.subtle}},line:{{color:C.green,pt:1}}}});
    sl.addShape(pres.shapes.OVAL,{{x:0.6,y:y+0.14,w:0.48,h:0.48,fill:{{color:C.green}},line:{{type:"none"}}}});
    sl.addText(String(j+1),{{x:0.6,y:y+0.14,w:0.48,h:0.48,fontSize:14,bold:true,color:C.dark,fontFace:"Calibri",align:"center",valign:"middle"}});
    sl.addText(b,{{x:1.22,y:y+0.08,w:W-2.0,h:0.6,fontSize:13,color:C.white,fontFace:"Calibri",valign:"middle",wrap:true}});
  }});
  addChrome(sl,i+1,t);
}}

function makeClosingSlide(sl,s,i,t){{
  sl.background={{color:C.bg}};
  sl.addShape(pres.shapes.RECTANGLE,{{x:0,y:H/2-0.05,w:W,h:0.1,fill:{{color:C.accent}},line:{{type:"none"}}}});
  sl.addShape(pres.shapes.RECTANGLE,{{x:0,y:0,w:0.3,h:H,fill:{{color:C.accent}},line:{{type:"none"}}}});
  sl.addText(s.title,{{x:0.9,y:2.0,w:11,h:1.6,fontSize:40,bold:true,color:C.white,fontFace:"Calibri"}});
  if(s.subtitle) sl.addText(s.subtitle,{{x:0.9,y:3.8,w:10,h:0.6,fontSize:18,color:C.accent,fontFace:"Calibri",italic:true}});
  if(s.body)     sl.addText(s.body,{{x:0.9,y:4.5,w:10,h:1.5,fontSize:13,color:C.midgray,fontFace:"Calibri",wrap:true}});
  sl.addText("SEGA  •  CONFIDENTIAL",{{x:0.9,y:H-0.8,w:6,h:0.4,fontSize:9,color:C.midgray,fontFace:"Calibri",bold:true,charSpacing:3}});
  addChrome(sl,i+1,t);
}}

const slides=data.slides||[], total=slides.length;
slides.forEach((s,i)=>{{
  const sl=pres.addSlide(), type=(s.type||"bullets").toLowerCase();
  if     (type==="title")          makeTitleSlide(sl,s,i,total);
  else if(type==="section")        makeSectionSlide(sl,s,i,total);
  else if(type==="comparison")     makeComparisonSlide(sl,s,i,total);
  else if(type==="stats")          makeStatsSlide(sl,s,i,total);
  else if(type==="recommendation") makeRecommendationSlide(sl,s,i,total);
  else if(type==="closing")        makeClosingSlide(sl,s,i,total);
  else                             makeBulletsSlide(sl,s,i,total);
  if(s.speaker_notes) sl.addNotes(s.speaker_notes);
}});

pres.writeFile({{fileName:"{output_file}"}})
  .then(()=>console.log("done"))
  .catch(e=>{{console.error(e);process.exit(1);}});
"""


def generate_pptx(slide_data: dict) -> str:
    tmp_dir     = tempfile.mkdtemp()
    data_file   = os.path.join(tmp_dir, "slide_data.json")
    output_file = os.path.join(tmp_dir, "output.pptx")
    script_file = os.path.join(tmp_dir, "gen_pptx.js")
    with open(data_file, "w") as f:
        json.dump(slide_data, f, indent=2)
    with open(script_file, "w") as f:
        f.write(build_js_script(data_file, output_file))
    result = subprocess.run(["node", script_file], capture_output=True, text=True, timeout=60)
    if result.returncode != 0:
        raise RuntimeError(f"pptxgenjs error:\n{result.stderr}\n{result.stdout}")
    if not os.path.exists(output_file):
        raise FileNotFoundError("PPTX was not created")
    return output_file


def _api_post(headers: dict, payload: dict, timeout: int = 120, max_retries: int = 4) -> dict:
    """
    POST to the Anthropic messages endpoint with exponential-backoff retry on 429.
    Reads the retry-after header when present; otherwise backs off at 10s, 20s, 40s, 80s.
    Raises RuntimeError on non-retryable errors or exhausted retries.
    """
    base_wait = 10
    for attempt in range(max_retries + 1):
        resp = requests.post(
            "https://api.anthropic.com/v1/messages",
            headers=headers,
            json=payload,
            timeout=timeout,
        )
        if resp.status_code == 429:
            if attempt == max_retries:
                raise RuntimeError(
                    f"Rate limited after {max_retries} retries. "
                    "Try again in a minute, or switch to a lower-tier model (Haiku) in the sidebar."
                )
            # Honour the retry-after header if the API sends one
            retry_after = resp.headers.get("retry-after") or resp.headers.get("x-ratelimit-reset-requests")
            try:
                wait = max(float(retry_after), 1)
            except (TypeError, ValueError):
                wait = base_wait * (2 ** attempt)
            time.sleep(wait)
            continue

        resp.raise_for_status()
        return resp.json()

    raise RuntimeError("Unexpected exit from retry loop")


def run_pipeline(model, uploaded_files, game_title, business_question, audience,
                 theme_preset, web_search_en, slide_count):
    """Generator — yields (type, ...) tuples for incremental UI updates."""

    # API key from st.secrets — same pattern as BIreport.py
    api_key = st.secrets.get("ANTHROPIC_API_KEY", "")
    if not api_key:
        yield ("error", "ANTHROPIC_API_KEY not found in st.secrets. Add it to .streamlit/secrets.toml.")
        return

    headers = {
        "Content-Type": "application/json",
        "x-api-key": api_key,
        "anthropic-version": "2023-06-01",
    }

    # Step 1 — extract
    yield ("log", "📄 Extracting document content…")
    doc_texts = []
    for f in uploaded_files:
        f.seek(0)
        doc_texts.append(f"=== {f.name} ===\n{extract_text_from_file(f)}")
    combined_docs = "\n\n".join(doc_texts) if doc_texts else "[No documents uploaded]"
    yield ("step_done", "extract")
    yield ("log", f"✅ Extracted {len(uploaded_files)} document(s)")

    # Step 2 — web research
    yield ("log", f"🔍 Searching web for: {game_title}…")
    research_text = ""
    if web_search_en and game_title:
        try:
            data = _api_post(
                headers=headers,
                payload={
                    "model": model,
                    "max_tokens": 3000,
                    "tools": [{"type": "web_search_20250305", "name": "web_search"}],
                    "messages": [{"role": "user", "content": (
                        f"Research the game '{game_title}' for a business analysis. Cover: "
                        "genre, platform, developer, publisher, release date, Metacritic/OpenCritic scores, "
                        "user sentiment, sales figures, key gameplay mechanics, scope, player reception "
                        "(what worked/didn't), post-launch updates/DLC, and market context. "
                        "Provide a thorough structured summary."
                    )}],
                },
                timeout=60,
            )
            blocks = data.get("content", [])
            research_text = "\n".join(b.get("text", "") for b in blocks if b.get("type") == "text")
        except RuntimeError as e:
            yield ("error", str(e))
            return
        except Exception as e:
            research_text = f"[Web search error: {e}]"
    elif game_title:
        research_text = f"[Web search disabled — using model knowledge for '{game_title}']"
    else:
        research_text = "[No reference game specified]"

    yield ("step_done", "research")
    yield ("log", "✅ Web research complete")

    # Step 3 — analysis
    yield ("log", "🤖 Running Claude analysis…")

    theme_desc = {
        "SEGA Blue — Corporate Executive": "Professional SEGA corporate blue (#0055AA) — boardroom-ready.",
        "SEGA Dark — Game Reveal Style":   "Dark dramatic (#040A1C) with electric blue accents.",
        "SEGA Sonic — High Energy":        "Vibrant SEGA blue + gold accents, dynamic energy.",
    }.get(theme_preset, "SEGA corporate blue")

    analysis_prompt = f"""You are a senior game industry analyst at SEGA.
Analyse the following and produce structured content for a {slide_count}-slide executive presentation.

## INTERNAL GAME DOCUMENTS:
{combined_docs}

## REFERENCE GAME RESEARCH — {game_title}:
{research_text}

## BUSINESS QUESTION:
{business_question}

## AUDIENCE:
{audience}

Produce a JSON object for {slide_count} slides in SEGA style ({theme_desc}).
Schema:

{{
  "title":"...", "subtitle":"...",
  "theme":{{"primary":"hex","secondary":"hex","accent":"hex","background":"hex","text_light":"hex","text_dark":"hex"}},
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

Make content specific, data-driven, and genuinely useful for {audience}.
Use REAL data from the documents and research.
Return ONLY the JSON object — no markdown, no explanation."""

    try:
        data = _api_post(
            headers=headers,
            payload={
                "model": model,
                "max_tokens": 8000,
                "system": "You are a precise game industry analyst. Always return valid JSON only.",
                "messages": [{"role": "user", "content": analysis_prompt}],
            },
            timeout=120,
        )
        raw_json = data["content"][0]["text"].strip()
    except RuntimeError as e:
        yield ("error", str(e))
        return
    except Exception as e:
        yield ("error", f"API error: {e}")
        return

    if raw_json.startswith("```"):
        raw_json = raw_json.split("\n", 1)[1]
        if raw_json.endswith("```"):
            raw_json = raw_json.rsplit("```", 1)[0]

    try:
        slide_data = json.loads(raw_json)
    except json.JSONDecodeError as e:
        yield ("error", f"JSON parse error: {e}\n\nFirst 400 chars:\n{raw_json[:400]}")
        return

    yield ("step_done", "analyze")
    yield ("log", f"✅ Analysis complete — {len(slide_data.get('slides', []))} slides planned")
    yield ("slide_data", slide_data)

    # Step 4 — generate PPTX
    yield ("log", "🖥️ Generating PPTX…")
    try:
        pptx_path = generate_pptx(slide_data)
    except Exception as e:
        yield ("error", f"PPTX generation error: {e}")
        return

    yield ("step_done", "generate")
    yield ("log", "✅ PPTX ready!")
    yield ("pptx_path", pptx_path)


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

                    elif etype == "pptx_path":
                        with open(event[1], "rb") as f:
                            pptx_bytes = f.read()
                        fname = f"SEGA_Analysis_{(game_title or 'Report').replace(' ','_')}.pptx"
                        st.session_state["pptx_bytes"]    = pptx_bytes
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