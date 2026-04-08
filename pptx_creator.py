"""
pptx_creator.py — General-purpose AI PowerPoint Creator
A standalone Streamlit app for generating polished presentations on any topic.

Reuses the PPTX rendering engine from pptxruns.py (generate_pptx,
_render_plan_modal, extract_text_from_file etc.) — no duplication.

Auth: configurable domain via st.secrets["ALLOWED_DOMAIN"] or open to all.
"""

import streamlit as st
import io, time, hashlib, hmac, base64, json, re, os
import pypdf
import pandas as pd
from datetime import datetime, timedelta, timezone
from concurrent.futures import ThreadPoolExecutor, as_completed

# ── Reuse rendering engine from pptxruns ─────────────────────────────────────
import sys, importlib
sys.path.insert(0, os.path.dirname(__file__))

# We import just the pure functions — no Streamlit side-effects at import time
from pptxruns import (
    generate_pptx,
    extract_text_from_file,
    _make_anthropic_client,
    _render_plan_modal,
    _sign_cookie, _verify_cookie,
    _send_otp,
)

# ── Storage ───────────────────────────────────────────────────────────────────
from storage_pptx import (
    init_db as _init_db,
    get_projects, project_exists, create_project,
    rename_project, delete_project, load_project, save_project,
)
_init_db()

# ── Constants ─────────────────────────────────────────────────────────────────
_MAX_DOC_CHARS     = 60_000
COOKIE_EXPIRY_DAYS = 7
COOKIE_NAME        = "pptx_creator_auth"
OTP_EXPIRY_SECS    = 600

# Auth: set ALLOWED_DOMAIN in secrets to restrict (e.g. "@mycompany.com")
# Leave blank or omit to allow any email address.
ALLOWED_DOMAIN = st.secrets.get("ALLOWED_DOMAIN", "")

# ── Purpose presets ───────────────────────────────────────────────────────────
PURPOSE_PRESETS = {
    "Executive briefing":   "Concise, data-driven. Lead with key takeaways. Use stats and comparison slides.",
    "Market analysis":      "Comprehensive research. Include market sizing, competitors, trends, opportunities.",
    "Project proposal":     "Problem → solution → plan → ask. Include timeline, budget, risks, success metrics.",
    "Sales deck":           "Hook early. Focus on value, outcomes, social proof. Close with clear CTA.",
    "Training material":    "Step-by-step. Clear objectives, examples, exercises, summaries per section.",
    "Research summary":     "Findings-first. Methodology, evidence, implications, limitations.",
    "Board update":         "High-level. KPIs, risks, decisions needed. Minimal detail.",
    "Product roadmap":      "Vision → now/next/later. Features, timeline, dependencies, success metrics.",
    "Investor pitch":       "Problem, solution, market, traction, team, ask. Narrative arc.",
    "General / Other":      "Balanced professional presentation suited to the audience and topic.",
}

# ── Theme presets (not SEGA-specific) ─────────────────────────────────────────
THEME_PRESETS = {
    "Corporate Blue":    {"primary": "1A3A6B", "accent": "0099CC"},
    "Executive Dark":    {"primary": "1E2761", "accent": "CADCFC"},
    "Forest & Moss":     {"primary": "2C5F2D", "accent": "97BC62"},
    "Charcoal Minimal":  {"primary": "36454F", "accent": "A8C0CC"},
    "Coral Energy":      {"primary": "F96167", "accent": "2F3C7E"},
    "Warm Terracotta":   {"primary": "B85042", "accent": "A7BEAE"},
    "Teal Trust":        {"primary": "028090", "accent": "02C39A"},
    "Berry & Cream":     {"primary": "6D2E46", "accent": "ECE2D0"},
}

# ─────────────────────────────────────────────────────────────────────────────
# Auth helpers (cookie rolling, OTP — same pattern as pptxruns)
# ─────────────────────────────────────────────────────────────────────────────

def _auth_gate():
    """
    Show login UI if not authenticated. Returns True if authenticated.
    Handles: cookie check, OTP send, OTP verify.
    """
    # Init cookie manager
    _cm = None
    _existing = None
    try:
        import extra_streamlit_components as stx
        _cm = stx.CookieManager(key="pptx_creator_cookies")
        _existing = _cm.get(COOKIE_NAME)
    except Exception:
        pass

    _cookie_email = _verify_cookie(_existing) if _existing else None

    for k, v in [
        ("auth_verified", False), ("auth_email", ""),
        ("otp_code", ""), ("otp_email", ""), ("otp_expiry", 0),
        ("otp_sent", False), ("otp_attempts", 0),
    ]:
        if k not in st.session_state:
            st.session_state[k] = v

    if _cookie_email and not st.session_state.auth_verified:
        st.session_state.auth_verified = True
        st.session_state.auth_email    = _cookie_email
        # Roll cookie
        if _cm:
            _cm.set(COOKIE_NAME, _sign_cookie(_cookie_email),
                    expires_at=datetime.now(timezone.utc) + timedelta(days=COOKIE_EXPIRY_DAYS),
                    key="roll_cookie")

    if st.session_state.auth_verified:
        return True

    # ── Login form ────────────────────────────────────────────────────────────
    st.markdown("""
    <style>
    .login-wrap{max-width:400px;margin:3rem auto;padding:2rem;
        background:#1e293b;border-radius:12px;border:1px solid #334155}
    .login-title{font-size:1.5rem;font-weight:700;color:#f1f5f9;margin-bottom:.25rem}
    .login-sub{font-size:.85rem;color:#94a3b8;margin-bottom:1.5rem}
    </style>
    <div class="login-wrap">
      <div class="login-title">✨ PowerPoint Creator</div>
      <div class="login-sub">Sign in to continue</div>
    </div>
    """, unsafe_allow_html=True)

    with st.form("login_form"):
        email = st.text_input("Email address", placeholder="you@yourcompany.com")
        send  = st.form_submit_button("Send verification code", use_container_width=True)

    if send and email:
        if ALLOWED_DOMAIN and not email.lower().endswith(ALLOWED_DOMAIN.lower()):
            st.error(f"Only {ALLOWED_DOMAIN} addresses are allowed.")
            return False
        code = str(hash(time.time()) % 900000 + 100000)
        if _send_otp(email, code):
            st.session_state.otp_code    = code
            st.session_state.otp_email   = email
            st.session_state.otp_expiry  = time.time() + OTP_EXPIRY_SECS
            st.session_state.otp_sent    = True
            st.session_state.otp_attempts = 0
            st.rerun()
        else:
            st.error("Failed to send email. Check AWS SES config.")

    if st.session_state.otp_sent:
        with st.form("otp_form"):
            otp_in = st.text_input("Enter 6-digit code", max_chars=6)
            verify = st.form_submit_button("Verify", use_container_width=True, type="primary")

        if verify:
            if time.time() > st.session_state.otp_expiry:
                st.error("Code expired. Please request a new one.")
                st.session_state.otp_sent = False
            elif st.session_state.otp_attempts >= 5:
                st.error("Too many attempts. Please request a new code.")
                st.session_state.otp_sent = False
            elif otp_in.strip() == st.session_state.otp_code:
                st.session_state.auth_verified = True
                st.session_state.auth_email    = st.session_state.otp_email
                if _cm:
                    _cm.set(COOKIE_NAME, _sign_cookie(st.session_state.otp_email),
                            expires_at=datetime.now(timezone.utc) + timedelta(days=COOKIE_EXPIRY_DAYS),
                            key="set_cookie")
                st.rerun()
            else:
                st.session_state.otp_attempts += 1
                st.error(f"Incorrect code ({5 - st.session_state.otp_attempts} attempts left).")

    return False


# ─────────────────────────────────────────────────────────────────────────────
# Web research — generic, topic-aware
# ─────────────────────────────────────────────────────────────────────────────

def _web_research(topic: str, purpose: str, industry: str, question: str) -> str:
    """Run a web search and return a structured research brief on the topic."""
    if not topic.strip():
        return "[No research topic specified]"

    purpose_hint = PURPOSE_PRESETS.get(purpose, purpose)
    industry_hint = f" in the {industry} industry" if industry.strip() else ""

    prompt = (
        f'Research "{topic}"{industry_hint} for a presentation. '
        f'Presentation purpose: {purpose}. {purpose_hint}\n\n'
        f'Business question to address: {question or "(general overview)"}\n\n'
        "Write a structured research brief covering:\n"
        "OVERVIEW: What is this topic? Current state, key facts, relevant numbers.\n"
        "KEY PLAYERS / STAKEHOLDERS: Main organisations, people, or forces involved.\n"
        "DATA & EVIDENCE: Specific statistics, survey results, market figures, benchmarks.\n"
        "TRENDS: 3-4 notable recent developments or directions.\n"
        "CHALLENGES & OPPORTUNITIES: Main risks and areas for growth or improvement.\n"
        "CONTEXT: How does this compare to alternatives, competitors, or prior state?\n\n"
        "Use real, specific numbers wherever possible. Aim for 600-800 words."
    )

    import concurrent.futures as _cf
    TIMEOUT = 90

    def _fetch():
        try:
            client = _make_anthropic_client()
            msg = client.messages.create(
                model="claude-haiku-4-5-20251001",
                max_tokens=2000,
                tools=[{"type": "web_search_20250305", "name": "web_search"}],
                messages=[{"role": "user", "content": prompt}],
            )
            return "\n".join(
                b.text for b in msg.content
                if hasattr(b, "type") and b.type == "text" and b.text
            ) or "[No results returned]"
        except Exception as e:
            return f"[Web research error: {e}]"

    with _cf.ThreadPoolExecutor(max_workers=1) as ex:
        fut = ex.submit(_fetch)
        try:
            return fut.result(timeout=TIMEOUT)
        except _cf.TimeoutError:
            return f"[Web research timed out after {TIMEOUT}s — using model knowledge]"


# ─────────────────────────────────────────────────────────────────────────────
# Analysis prompt — fully generic
# ─────────────────────────────────────────────────────────────────────────────

def _build_analysis_prompt(
    topic: str,
    purpose: str,
    industry: str,
    audience: str,
    question: str,
    slide_count: int,
    doc_text: str,
    research_text: str,
    data_summary: str,
    theme: dict,
) -> str:
    purpose_hint = PURPOSE_PRESETS.get(purpose, purpose)
    industry_ctx = f" in the {industry} industry" if industry.strip() else ""
    primary  = theme.get("primary", "1A3A6B")
    accent   = theme.get("accent",  "0099CC")

    return f"""You are an expert presentation strategist and analyst.
Create a JSON outline for a {slide_count}-slide professional presentation.

## TOPIC: {topic}{industry_ctx}
## PURPOSE: {purpose} — {purpose_hint}
## AUDIENCE: {audience}
## BUSINESS QUESTION / GOAL:
{question or "(No specific question — create a comprehensive overview)"}

## UPLOADED DOCUMENTS:
{doc_text}

## RESEARCH BRIEF:
{research_text}

## DATA FOR CHARTS:
{data_summary if data_summary else "(none uploaded)"}

Output a single JSON object. Schema:
{{
  "title":"...", "subtitle":"...",
  "theme":{{"primary":"{primary}","accent":"{accent}"}},
  "slides":[
    {{
      "type":"title|section|bullets|stats|comparison|recommendation|chart|closing",
      "title":"...","subtitle":"...","body":"...",
      "bullets":["..."],
      "stats":[{{"label":"...","value":"...","note":"..."}}],
      "comparison":{{
        "left_title":"...","right_title":"...",
        "rows":[{{"label":"...","left":"...","right":"...","delta":"positive|negative|neutral"}}]
      }},
      "chart":{{
        "chart_type":"bar|line|scatter|pie|horizontal_bar",
        "title":"...","x_label":"...","y_label":"...",
        "categories":["..."],
        "series":[{{"label":"...","values":[0.0]}}],
        "colors":["hex"]
      }},
      "speaker_notes":"..."
    }}
  ]
}}

Rules:
- Use REAL data from documents and research — no placeholders
- Tailor depth, tone, and structure to: {purpose} for {audience}
- Purpose hint: {purpose_hint}
- theme.primary and theme.accent must be vivid hex (6 digits, no #)
- speaker_notes: 1-2 sentences max — brief presenter cues only
- Bullets: max 6 per slide, each under 15 words
- Comparison rows: max 8 per slide
- Use "chart" type when data supports it
- Return ONLY valid JSON — no markdown fences, no explanation
"""


# ─────────────────────────────────────────────────────────────────────────────
# Pipeline — yields (event_type, payload) events
# ─────────────────────────────────────────────────────────────────────────────

def run_pipeline(
    model: str,
    uploaded_files: list,
    topic: str,
    purpose: str,
    industry: str,
    audience: str,
    question: str,
    web_search_en: bool,
    slide_count: int,
    theme: dict,
    template_bytes: bytes | None = None,
    data_files: list | None = None,
    plan_mode: bool = True,
):
    if not st.secrets.get("ANTHROPIC_API_KEY", ""):
        yield ("error", "ANTHROPIC_API_KEY not found in st.secrets.")
        return

    yield ("spinner", (
        "📂 <b>Stage 1 of 4 — Extracting your documents</b><br>"
        "<span class='log-detail'>Reading uploaded files and pulling out text content. "
        f"Content capped at {_MAX_DOC_CHARS:,} chars to stay within token limits.</span>"
    ))

    if web_search_en and topic.strip():
        yield ("spinner", (
            f"🔍 <b>Stage 2 of 4 — Web research on \"{topic}\" running in parallel</b><br>"
            "<span class='log-detail'>Searching for current data, stats, and context. "
            "90s deadline — falls back to model knowledge if it times out.</span>"
        ))

    doc_text     = "[No documents uploaded]"
    research_txt = "[No web research — using model knowledge]"
    _file_stats  = []

    def _extract_docs():
        if not uploaded_files:
            return "[No documents uploaded]"
        parts = []
        for f in uploaded_files:
            f.seek(0)
            txt = extract_text_from_file(f)
            _file_stats.append((f.name, len(txt)))
            parts.append(f"=== {f.name} ===\n{txt}")
        full = "\n\n".join(parts)
        if len(full) > _MAX_DOC_CHARS:
            full = full[:_MAX_DOC_CHARS] + "\n\n[... trimmed ...]"
        return full

    def _do_research():
        if not web_search_en or not topic.strip():
            return "[Web search disabled]"
        return _web_research(topic, purpose, industry, question)

    # ── Parallel: extract docs + web research ─────────────────────────────────
    with ThreadPoolExecutor(max_workers=2) as pool:
        fut_docs     = pool.submit(_extract_docs)
        fut_research = pool.submit(_do_research)
        pending = {fut_docs: "docs", fut_research: "research"}
        elapsed = 0
        TICK    = 2
        CAP     = 110

        while pending:
            done = [f for f in list(pending) if f.done()]
            for f in done:
                label = pending.pop(f)
                if label == "docs":
                    doc_text = f.result()
                    nf = len(uploaded_files) if uploaded_files else 0
                    yield ("log", (
                        f"✅ <b>Documents extracted</b> — {nf} file(s), "
                        f"{len(doc_text):,} chars<br>"
                        "<span class='log-detail'>" +
                        "".join(f"<br>&nbsp;&nbsp;· {n}: {c:,} chars"
                                for n, c in _file_stats) +
                        "</span>"
                    ))
                    yield ("step_done", "extract")
                else:
                    research_txt = f.result()
                    ok = not research_txt.startswith("[")
                    yield ("log", (
                        f"{'✅' if ok else '⚠️'} <b>Web research</b> — "
                        f"{'~' + str(len(research_txt.split())) + ' words' if ok else 'used model knowledge'}<br>"
                        f"<span class='log-detail'>Topic: {topic}</span>"
                    ))
                    yield ("step_done", "research")

            if pending:
                elapsed += TICK
                if elapsed >= CAP:
                    for f2 in list(pending):
                        pending.pop(f2)
                    research_txt = "[Research timed out]"
                    yield ("step_done", "research")
                    break
                still = list(pending.values())
                yield ("spinner", f"⏳ Still working: {', '.join(still)} ({elapsed}s)…")
                time.sleep(TICK)

    # ── Data files ─────────────────────────────────────────────────────────────
    _DATA_CAP   = 40_000
    data_summary = ""
    if data_files:
        parts = []
        total = 0
        for df_file in data_files:
            if total >= _DATA_CAP:
                parts.append(f"{df_file.name}: skipped (cap reached)")
                continue
            try:
                df_file.seek(0)
                name = df_file.name.lower()
                if name.endswith(".csv"):
                    df = pd.read_csv(df_file)
                    csv = df.to_csv(index=False)
                    if len(csv) > _DATA_CAP - total:
                        csv = df.head(200).to_csv(index=False) + "\n[truncated]"
                    parts.append(f"FILE: {df_file.name}\n{csv}")
                    total += len(csv)
                else:
                    sheets = pd.read_excel(df_file, sheet_name=None)
                    for sname, sdf in sheets.items():
                        csv = sdf.to_csv(index=False)
                        if len(csv) > (_DATA_CAP - total) // max(len(sheets), 1):
                            csv = sdf.head(200).to_csv(index=False) + "\n[truncated]"
                        parts.append(f"FILE: {df_file.name} / Sheet: {sname}\n{csv}")
                        total += len(csv)
            except Exception as e:
                parts.append(f"{df_file.name}: error — {e}")
        data_summary = "\n\n".join(parts)
        yield ("log", f"📊 <b>Data loaded</b> — {len(data_files)} file(s), {len(data_summary):,} chars")

    # ── Stage 3: streaming analysis ────────────────────────────────────────────
    prompt = _build_analysis_prompt(
        topic, purpose, industry, audience, question,
        slide_count, doc_text, research_txt, data_summary, theme,
    )
    est_tokens = len(prompt) // 4

    yield ("spinner", (
        f"🤖 <b>Stage 3 of 4 — Generating {slide_count}-slide outline</b><br>"
        f"<span class='log-detail'>Sending ~{est_tokens:,} tokens to {model}. "
        "Claude will synthesise your documents, research, and goal into a "
        "structured JSON slide plan with titles, bullets, stats, and speaker notes.</span>"
    ))

    import queue as _q, threading as _th

    chunk_q   = _q.Queue()
    SENTINEL  = object()
    err_box   = [None]

    def _stream():
        try:
            client = _make_anthropic_client()
            with client.messages.stream(
                model=model, max_tokens=8000,
                system="You are a precise presentation strategist. Return valid JSON only.",
                messages=[{"role": "user", "content": prompt}],
            ) as s:
                for chunk in s.text_stream:
                    chunk_q.put(chunk)
        except Exception as exc:
            err_box[0] = exc
        finally:
            chunk_q.put(SENTINEL)

    th = _th.Thread(target=_stream, daemon=True)
    th.start()

    raw_chunks   = []
    char_count   = 0
    last_tick    = 0
    first_chunk  = True
    stream_sec   = 0
    STREAM_LIMIT = 300
    slides_seen  = 0

    def _count_slides(txt):
        return len(re.findall(
            r'"type"\s*:\s*"(?:title|section|bullets|stats|comparison|recommendation|closing|chart)"',
            txt))

    done = False
    while not done:
        try:
            item = chunk_q.get(timeout=2)
        except _q.Empty:
            stream_sec += 2
            if stream_sec >= STREAM_LIMIT:
                yield ("error", f"Generation timed out after {STREAM_LIMIT}s.")
                return
            if first_chunk:
                yield ("spinner", f"⏳ Waiting for first token… {stream_sec}s")
            else:
                pct = min(int(char_count / (slide_count * 220) * 100), 95)
                yield ("spinner", f"  📝 Generating… {char_count:,} chars (~{pct}%)")
            continue

        if item is SENTINEL:
            if err_box[0]:
                raise err_box[0]
            done = True
            break

        first_chunk = False
        raw_chunks.append(item)
        char_count += len(item)

        cur = "".join(raw_chunks)
        n_sl = _count_slides(cur)
        if n_sl > slides_seen:
            for _ in range(n_sl - slides_seen):
                slides_seen += 1
                yield ("spinner", f"  ✏️ Writing slide {slides_seen} of {slide_count}…")

        if char_count - last_tick >= 600:
            last_tick = char_count
            pct = min(int(char_count / (slide_count * 220) * 100), 95)
            yield ("spinner", f"  📝 Generating… {char_count:,} chars (~{pct}%)")

    # ── Parse JSON ─────────────────────────────────────────────────────────────
    raw_json = "".join(raw_chunks).strip()
    if raw_json.startswith("```"):
        raw_json = raw_json.split("\n", 1)[1].rsplit("```", 1)[0]

    def _recover(s):
        try: return json.loads(s)
        except Exception: pass
        for pos in [s.rfind("}\n    ]"), s.rfind("}, {"), s.rfind('"}')]:
            if pos > 100:
                frag = s[:pos+1]
                opens  = frag.count("{") - frag.count("}")
                opens2 = frag.count("[") - frag.count("]")
                try: return json.loads(frag + "]"*opens2 + "}"*opens)
                except Exception: pass
        return None

    try:
        slide_data = json.loads(raw_json)
    except json.JSONDecodeError as e:
        recovered = _recover(raw_json)
        if recovered and recovered.get("slides"):
            n_rec = len(recovered["slides"])
            yield ("log", f"⚠️ Output truncated — recovered {n_rec} slides.")
            slide_data = recovered
        else:
            yield ("error", (
                f"JSON parse error: {e}<br>"
                f"First 400 chars: <code>{raw_json[:400]}</code>"
            ))
            return

    n_slides = len(slide_data.get("slides", []))
    yield ("step_done", "analyze")
    yield ("log", (
        f"✅ <b>Outline complete</b> — {n_slides} slides, "
        f"{char_count:,} chars<br>"
        f"<span class='log-detail'>Slide types: "
        f"{', '.join(sorted(set(s.get('type','?') for s in slide_data.get('slides',[]))))}"
        f"</span>"
    ))
    yield ("slide_data", slide_data)

    if plan_mode:
        yield ("plan_ready", slide_data)
        return

    # ── Stage 4: PPTX rendering ────────────────────────────────────────────────
    yield ("spinner", (
        f"🖥️ <b>Stage 4 of 4 — Building PPTX</b><br>"
        f"<span class='log-detail'>Rendering {n_slides} slides with python-pptx. "
        "Placing text, shapes, native charts, speaker notes.</span>"
    ))
    try:
        pptx_bytes = generate_pptx(slide_data, template_bytes=template_bytes)
    except Exception as e:
        yield ("error", f"PPTX render error: {e}")
        return

    yield ("step_done", "generate")
    yield ("log", (
        f"🎉 <b>Done!</b> — {len(pptx_bytes)//1024} KB, {n_slides} slides<br>"
        "<span class='log-detail'>Download below. Open in PowerPoint or Google Slides. "
        "Speaker notes included on every slide.</span>"
    ))
    yield ("pptx_bytes_out", pptx_bytes)


# ─────────────────────────────────────────────────────────────────────────────
# UI helpers
# ─────────────────────────────────────────────────────────────────────────────

def _render_log(log_lines: list) -> str:
    CSS = (
        "<style>"
        "@keyframes _sp{to{transform:rotate(360deg)}}"
        "._lw{background:#0f172a;border:1px solid #1e293b;border-radius:8px;"
        "padding:1rem 1.25rem;font-size:.82rem;"
        "font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;"
        "color:#94a3b8;max-height:480px;overflow-y:auto;line-height:1.7}"
        "._lw b{color:#e2e8f0}._lw i{color:#60a5fa}"
        "._ld{color:#64748b;font-size:.76rem;display:block;margin-top:.15rem}"
        "._le{border-left:2px solid #1e3a5f;padding-left:.6rem;margin-bottom:.5rem}"
        "._la{border-left:2px solid #6366f1;padding-left:.6rem;"
        "margin-bottom:.5rem;display:flex;align-items:flex-start;gap:.55rem}"
        "._lr{flex-shrink:0;width:14px;height:14px;margin-top:3px;"
        "border:2px solid rgba(99,102,241,.25);border-top-color:#6366f1;"
        "border-radius:50%;animation:_sp .8s linear infinite}"
        "._lt{flex:1}"
        "</style>"
    )
    parts = [CSS, '<div class="_lw">']
    for kind, text in log_lines[-14:]:
        t = text.replace("class='log-detail'", "class='_ld'")
        if kind == "spinner":
            parts.append(f'<div class="_la"><div class="_lr"></div><div class="_lt">{t}</div></div>')
        else:
            parts.append(f'<div class="_le">{t}</div>')
    parts.append("</div>")
    return "".join(parts)


# ─────────────────────────────────────────────────────────────────────────────
# Project management helpers
# ─────────────────────────────────────────────────────────────────────────────

def _load_project(owner: str, name: str):
    data = load_project(owner, name)
    if not data:
        return
    st.session_state["active_project"]      = name
    st.session_state["proj_topic"]          = data.get("business_question", "")
    st.session_state["proj_purpose"]        = data.get("audience", "General / Other")
    st.session_state["proj_industry"]       = data.get("game_title", "")
    st.session_state["proj_audience"]       = data.get("audience", "")
    st.session_state["project_doc_names"]   = data.get("doc_names", [])
    st.session_state["plan_slide_data"]     = data.get("slide_json") or {}
    st.session_state["plan_chat"]           = data.get("plan_chat", [])
    st.session_state["pptx_bytes"]          = data.get("pptx_bytes") or None
    st.session_state["saved_template_bytes"] = data.get("template_bytes") or None
    if st.session_state["plan_slide_data"]:
        st.session_state["plan_mode_active"] = True


def _save_project(owner: str, name: str):
    save_project(
        owner, name,
        business_question = st.session_state.get("proj_topic", ""),
        game_title        = st.session_state.get("proj_industry", ""),
        audience          = st.session_state.get("proj_audience", ""),
        doc_names         = st.session_state.get("project_doc_names", []),
        slide_json        = st.session_state.get("plan_slide_data") or {},
        plan_chat         = st.session_state.get("plan_chat", []),
        pptx_bytes        = st.session_state.get("pptx_bytes"),
        template_bytes    = st.session_state.get("saved_template_bytes"),
    )


def _clear_project():
    for k in [
        "active_project", "proj_topic", "proj_purpose", "proj_industry",
        "proj_audience", "project_doc_names", "plan_slide_data", "plan_chat",
        "pptx_bytes", "pptx_filename", "saved_template_bytes",
        "plan_mode_active", "plan_slide_history", "pipeline_steps",
    ]:
        st.session_state.pop(k, None)


# ─────────────────────────────────────────────────────────────────────────────
# Main app
# ─────────────────────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="PowerPoint Creator",
    page_icon="✨",
    layout="wide",
)

st.markdown("""
<style>
.section-label{font-size:.72rem;font-weight:700;letter-spacing:.08em;
    text-transform:uppercase;color:#6366f1;margin:.9rem 0 .35rem}
.sidebar-section{font-size:.68rem;font-weight:700;letter-spacing:.07em;
    text-transform:uppercase;color:#6366f1;display:block;margin:.8rem 0 .3rem}
.step-row{font-size:.78rem;padding:.18rem 0;color:#64748b}
.step-done{color:#22c55e}
.step-pending{color:#475569}
.status-card{background:#1e293b;border:1px solid #334155;border-radius:8px;
    padding:1.25rem 1.5rem;margin-top:.5rem}
.status-card-label{font-size:.68rem;font-weight:700;letter-spacing:.08em;
    text-transform:uppercase;color:#6366f1;margin-bottom:.5rem}
</style>
""", unsafe_allow_html=True)

# ── Auth gate ─────────────────────────────────────────────────────────────────
if not _auth_gate():
    st.stop()

OWNER = st.session_state.auth_email

# ── Sidebar ────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown(
        f"<div style='font-size:.78rem;color:#94a3b8;margin-bottom:1rem'>"
        f"Signed in as <b style='color:#e2e8f0'>{OWNER}</b></div>",
        unsafe_allow_html=True,
    )
    if st.button("Sign out", use_container_width=True):
        try:
            import extra_streamlit_components as stx
            cm = stx.CookieManager(key="signout_cookies")
            cm.delete(COOKIE_NAME, key="del_cookie")
        except Exception:
            pass
        for k in ["auth_verified", "auth_email", "otp_sent", "otp_code"]:
            st.session_state.pop(k, None)
        st.rerun()

    st.divider()

    # ── Project management ────────────────────────────────────────────────────
    st.markdown('<span class="sidebar-section">Project</span>', unsafe_allow_html=True)

    _projects = get_projects(OWNER)
    _proj_names = [p["name"] for p in _projects]
    _active = st.session_state.get("active_project", "")

    _select_opts = ["— select a project —"] + _proj_names
    _cur_idx = (_select_opts.index(_active) if _active in _select_opts else 0)
    _selected = st.selectbox("Project", _select_opts, index=_cur_idx,
                             label_visibility="hidden")

    if _selected != "— select a project —" and _selected != _active:
        _load_project(OWNER, _selected)
        st.query_params["project"] = _selected
        st.rerun()

    _new_name = st.text_input("New project name", placeholder="e.g. Q3 Investor Update",
                               label_visibility="hidden")
    if st.button("＋ Create project", use_container_width=True):
        if _new_name.strip():
            if project_exists(OWNER, _new_name.strip()):
                st.error("Name already exists.")
            else:
                create_project(OWNER, _new_name.strip())
                _clear_project()
                st.session_state["active_project"] = _new_name.strip()
                st.query_params["project"] = _new_name.strip()
                st.rerun()

    if _active:
        _c_save, _c_rename, _c_del = st.columns([3, 2, 1])
        with _c_save:
            if st.button("💾 Save", use_container_width=True):
                _save_project(OWNER, _active)
                st.toast(f'Saved "{_active}"', icon="✅")
        with _c_rename:
            if st.button("✏️", help="Rename", use_container_width=True):
                st.session_state["_renaming"] = True
        with _c_del:
            if st.button("🗑", help="Delete", use_container_width=True):
                st.session_state["_confirm_del"] = True

        if st.session_state.get("_renaming"):
            _rn = st.text_input("New name", value=_active, key="rename_input")
            if st.button("Confirm rename", key="confirm_rename"):
                if _rn.strip() and _rn.strip() != _active:
                    rename_project(OWNER, _active, _rn.strip())
                    st.session_state["active_project"] = _rn.strip()
                    st.session_state.pop("_renaming", None)
                    st.rerun()

        if st.session_state.get("_confirm_del"):
            st.warning(f'Delete "{_active}"?')
            _dc, _dk = st.columns(2)
            with _dc:
                if st.button("Yes, delete", key="del_yes"):
                    delete_project(OWNER, _active)
                    _clear_project()
                    st.session_state.pop("_confirm_del", None)
                    st.query_params.clear()
                    st.rerun()
            with _dk:
                if st.button("Cancel", key="del_no"):
                    st.session_state.pop("_confirm_del", None)
                    st.rerun()

        if st.session_state.get("pptx_bytes") or st.session_state.get("plan_slide_data"):
            if st.button("🔄 Reset output", use_container_width=True):
                for k in ["pptx_bytes","pptx_filename","plan_slide_data",
                          "plan_mode_active","plan_chat","plan_slide_history","pipeline_steps"]:
                    st.session_state.pop(k, None)
                save_project(OWNER, _active, slide_json={}, plan_chat=[], clear_pptx=True)
                st.rerun()

    st.divider()

    # ── Model ────────────────────────────────────────────────────────────────
    st.markdown('<span class="sidebar-section">Model</span>', unsafe_allow_html=True)
    model = st.selectbox(
        "Model", ["claude-sonnet-4-5", "claude-opus-4-5", "claude-haiku-4-5-20251001"],
        label_visibility="hidden",
    )
    st.session_state["_selected_model"] = model

    # ── Options ───────────────────────────────────────────────────────────────
    st.markdown('<span class="sidebar-section">Options</span>', unsafe_allow_html=True)
    web_search_enabled = st.checkbox("Web research", value=True,
        help="Search the web for current data on your topic")
    slide_count = st.slider("Target slides", 6, 25, 12)

    # ── Theme ────────────────────────────────────────────────────────────────
    st.markdown('<span class="sidebar-section">Theme</span>', unsafe_allow_html=True)
    theme_name = st.selectbox("Color theme", list(THEME_PRESETS.keys()),
                               label_visibility="hidden")
    theme = THEME_PRESETS[theme_name]

    # ── Template ────────────────────────────────────────────────────────────
    st.markdown('<span class="sidebar-section">Template (optional)</span>', unsafe_allow_html=True)
    template_file = st.file_uploader(
        "Upload .pptx template", type=["pptx"],
        label_visibility="hidden", key="template_upload",
        help="Upload a branded .pptx — layouts, fonts and colors will be preserved.",
    )

    # ── Pipeline status ─────────────────────────────────────────────────────
    st.divider()
    st.markdown('<span class="sidebar-section">Pipeline</span>', unsafe_allow_html=True)
    _pipeline = st.session_state.get("pipeline_steps", {
        "upload":False,"extract":False,"research":False,"analyze":False,"generate":False
    })
    for key, label in {
        "upload":"Document upload","extract":"Content extraction",
        "research":"Web research","analyze":"AI analysis","generate":"PPTX generation"
    }.items():
        done = _pipeline.get(key, False)
        st.markdown(
            f'<div class="step-row {"step-done" if done else "step-pending"}">'
            f'{"✓" if done else "○"}&nbsp; {label}</div>',
            unsafe_allow_html=True,
        )

# ── No-project gate ───────────────────────────────────────────────────────────
if not _active:
    # Check URL param
    _qp = st.query_params.get("project", "")
    if _qp and project_exists(OWNER, _qp):
        _load_project(OWNER, _qp)
        st.rerun()

    st.markdown("""
    <div style='max-width:520px;margin:4rem auto;text-align:center'>
      <div style='font-size:2.5rem;margin-bottom:1rem'>✨</div>
      <h2 style='color:#e2e8f0;margin-bottom:.5rem'>PowerPoint Creator</h2>
      <p style='color:#64748b;margin-bottom:2rem'>
        AI-powered presentations for any topic — market analysis, project proposals,
        executive briefings, research summaries, and more.
      </p>
      <p style='color:#94a3b8;font-size:.85rem'>
        Select an existing project or create a new one in the sidebar to get started.
      </p>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

# ── Page header ────────────────────────────────────────────────────────────────
st.markdown(
    f"<h1>✨ PowerPoint Creator</h1>"
    f"<div style='font-size:.78rem;color:#475569;margin-bottom:1.5rem'>"
    f"Project: <b>{_active}</b> &nbsp;·&nbsp; "
    f"Upload documents &nbsp;·&nbsp; Describe your goal &nbsp;·&nbsp; Generate a polished PPTX"
    f"</div>",
    unsafe_allow_html=True,
)

_tab_main, _tab_pdf = st.tabs(["📊 Create Presentation", "📄 PDF → Editable PPTX"])

with _tab_pdf:
    st.markdown(
        "<div style='font-size:.85rem;color:#94a3b8;margin-bottom:1rem'>"
        "Upload any PDF of slides and convert it to a fully editable PPTX — "
        "works even on image-only exports from Canva, Google Slides, or NotebookLM."
        "</div>",
        unsafe_allow_html=True,
    )
    _pdf_col, _pdf_out = st.columns([1, 1], gap="large")
    with _pdf_col:
        _pdf_file  = st.file_uploader("PDF", type=["pdf"],
                                       label_visibility="hidden", key="pdf_upload")
        _pdf_model = st.selectbox("Vision model",
                                   ["claude-opus-4-5", "claude-sonnet-4-5"],
                                   key="pdf_vision_model")
        _pdf_dpi   = st.select_slider("Quality (DPI)",
                                       options=[96,120,150,200], value=150, key="pdf_dpi")
        _pdf_btn   = st.button("⚡ Convert", use_container_width=True,
                                type="primary", disabled=not _pdf_file)
    with _pdf_out:
        _pdf_area = st.empty()
        if "pdf_pptx_bytes" in st.session_state:
            _pdf_area.download_button(
                "⬇️ Download editable PPTX",
                data=st.session_state["pdf_pptx_bytes"],
                file_name=st.session_state.get("pdf_pptx_filename","converted.pptx"),
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True,
            )
    if _pdf_btn and _pdf_file:
        from pdf_to_pptx import pdf_to_editable_pptx as _pdf_convert
        _pdf_file.seek(0)
        _raw = _pdf_file.read()
        _pb  = _pdf_area.progress(0, text="Starting…")
        _la  = st.empty()
        _ll  = []
        def _plog(m):
            _ll.append(m)
            _la.markdown(
                "<div style='background:#0f172a;border-radius:6px;padding:.75rem;"
                "font-size:.8rem;color:#94a3b8;font-family:monospace'>"
                + "".join(f"<div>{x}</div>" for x in _ll[-12:]) + "</div>",
                unsafe_allow_html=True)
        def _pprog(f):
            _pb.progress(min(f,0.99), text=f"Page {max(1,round(f*10))}…")
        _plog("🔍 Extracting and converting…")
        try:
            _out, _errs = _pdf_convert(
                _raw, api_key=st.secrets.get("ANTHROPIC_API_KEY",""),
                model=_pdf_model, dpi=_pdf_dpi, progress_cb=_pprog)
            for e in _errs: _plog(f"⚠️ {e}")
            _pb.progress(1.0, text="Done!")
            _fname = _pdf_file.name.rsplit(".",1)[0] + "_editable.pptx"
            st.session_state["pdf_pptx_bytes"]    = _out
            st.session_state["pdf_pptx_filename"] = _fname
            _pdf_area.download_button(
                "⬇️ Download editable PPTX", data=_out, file_name=_fname,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True)
            st.rerun()
        except Exception as _ex:
            st.error(f"Conversion failed: {_ex}")

with _tab_main:
    col_left, col_right = st.columns([1.1, 0.9], gap="large")

    with col_left:
        st.markdown('<div class="section-label">Documents (optional)</div>', unsafe_allow_html=True)
        uploaded_files = st.file_uploader(
            "Upload supporting documents",
            type=["pdf","xlsx","xls","csv","txt","docx"],
            accept_multiple_files=True, label_visibility="hidden",
        )
        if uploaded_files:
            st.caption(f"{len(uploaded_files)} file(s): " + ", ".join(f.name for f in uploaded_files))
            st.session_state["project_doc_names"] = [f.name for f in uploaded_files]

        st.markdown('<div class="section-label">Data for charts (optional)</div>', unsafe_allow_html=True)
        data_files = st.file_uploader(
            "Upload data files", type=["xlsx","xls","csv"],
            accept_multiple_files=True, label_visibility="hidden", key="data_upload",
            help="Excel or CSV files — mention charts in the goal field below.",
        )

        st.markdown('<div class="section-label">Presentation details</div>', unsafe_allow_html=True)

        topic = st.text_input(
            "Topic / subject",
            value=st.session_state.get("proj_topic", ""),
            placeholder="e.g. EV market landscape, Q3 earnings review, office relocation proposal",
            key="topic_input",
        )
        st.session_state["proj_topic"] = topic

        purpose = st.selectbox(
            "Purpose",
            list(PURPOSE_PRESETS.keys()),
            index=list(PURPOSE_PRESETS.keys()).index(
                st.session_state.get("proj_purpose", "General / Other")
                if st.session_state.get("proj_purpose", "General / Other")
                   in PURPOSE_PRESETS else "General / Other"
            ),
            help="Shapes the structure, depth, and tone of the presentation.",
        )
        st.session_state["proj_purpose"] = purpose

        industry = st.text_input(
            "Industry / context (optional)",
            value=st.session_state.get("proj_industry", ""),
            placeholder="e.g. SaaS, healthcare, real estate, fintech",
            key="industry_input",
        )
        st.session_state["proj_industry"] = industry

        audience = st.text_input(
            "Audience",
            value=st.session_state.get("proj_audience", "Executive team"),
            placeholder="e.g. Board of directors, Sales team, Engineering leads",
            key="audience_input",
        )
        st.session_state["proj_audience"] = audience

        question = st.text_area(
            "Goal / business question",
            value=st.session_state.get("proj_question", ""),
            placeholder=(
                "e.g. Make the case for expanding into the EU market by Q2, "
                "highlighting opportunity size, competitive gaps, and required investment…"
            ),
            height=120,
            key="question_input",
        )
        st.session_state["proj_question"] = question

        run_btn = st.button("⚡ Generate presentation", use_container_width=True, type="primary")

    with col_right:
        st.markdown('<div class="section-label">Output</div>', unsafe_allow_html=True)
        output_area   = st.empty()
        download_area = st.empty()

        if st.session_state.get("plan_mode_active") and not run_btn:
            with output_area.container():
                _render_plan_modal(st.session_state.get("template_upload"))
        elif not run_btn and "pptx_bytes" not in st.session_state:
            output_area.markdown("""
<div class="status-card">
<div class="status-card-label">Ready</div>
<div class="status-card-value" style="color:#475569;font-size:.82rem;line-height:1.9">
Fill in the details on the left and click <strong style="color:#e2e8f0">Generate presentation</strong>.<br><br>
The pipeline will:<br>
&nbsp;1. Extract your uploaded documents<br>
&nbsp;2. Search the web for current data on your topic<br>
&nbsp;3. Generate a structured slide outline with Claude<br>
&nbsp;4. Let you review and edit the outline<br>
&nbsp;5. Export to a polished, editable PPTX
</div>
</div>
""", unsafe_allow_html=True)

        if "pptx_bytes" in st.session_state and not run_btn:
            download_area.download_button(
                "⬇️ Download previous PPTX",
                data=st.session_state["pptx_bytes"],
                file_name=st.session_state.get("pptx_filename", "presentation.pptx"),
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True,
            )

    # ── Run button handler ─────────────────────────────────────────────────────
    if run_btn:
        if not topic.strip():
            st.error("Please enter a topic.")
        else:
            data_files_list = st.session_state.get("data_upload") or []
            if hasattr(data_files_list, "read"):
                data_files_list = [data_files_list]

            _template_bytes = None
            _tf = st.session_state.get("template_upload")
            if _tf:
                try:
                    _tf.seek(0); _template_bytes = _tf.read()
                    st.session_state["saved_template_bytes"] = _template_bytes
                except Exception:
                    pass

            _pipeline_steps = {
                "upload": bool(uploaded_files), "extract": False,
                "research": False, "analyze": False, "generate": False,
            }
            st.session_state["pipeline_steps"] = _pipeline_steps
            log_lines = []

            with output_area.container():
                st.markdown('<div class="section-label">Pipeline log</div>', unsafe_allow_html=True)
                log_area = st.empty()

                try:
                    for event in run_pipeline(
                        model=model,
                        uploaded_files=uploaded_files,
                        topic=topic,
                        purpose=purpose,
                        industry=industry,
                        audience=audience,
                        question=question,
                        web_search_en=web_search_enabled,
                        slide_count=slide_count,
                        theme=theme,
                        template_bytes=_template_bytes,
                        data_files=data_files_list,
                        plan_mode=True,
                    ):
                        etype = event[0]

                        if etype in ("log", "spinner"):
                            if etype == "spinner":
                                if log_lines and log_lines[-1][0] == "spinner":
                                    log_lines[-1] = ("log", log_lines[-1][1])
                                log_lines.append(("spinner", event[1]))
                            else:
                                if log_lines and log_lines[-1][0] == "spinner":
                                    log_lines[-1] = ("log", log_lines[-1][1])
                                log_lines.append(("log", event[1]))
                            log_area.markdown(_render_log(log_lines), unsafe_allow_html=True)

                        elif etype == "step_done":
                            _pipeline_steps[event[1]] = True
                            st.session_state["pipeline_steps"] = _pipeline_steps

                        elif etype == "slide_data":
                            st.session_state["slide_data"] = event[1]

                        elif etype == "plan_ready":
                            st.session_state["plan_slide_data"]  = event[1]
                            st.session_state["plan_mode_active"] = True
                            if _active:
                                _save_project(OWNER, _active)

                        elif etype == "pptx_bytes_out":
                            _slug = re.sub(r"[^a-zA-Z0-9]+", "_", topic)[:50]
                            fname = f"Presentation_{_slug}.pptx"
                            st.session_state["pptx_bytes"]    = event[1]
                            st.session_state["pptx_filename"] = fname
                            if _active:
                                _save_project(OWNER, _active)

                        elif etype == "error":
                            st.error(event[1], icon="🚨")
                            break

                except Exception as ex:
                    st.error(f"Unexpected error: {ex}")
                    import traceback; st.code(traceback.format_exc())

            if st.session_state.get("plan_mode_active"):
                with output_area.container():
                    _render_plan_modal(st.session_state.get("template_upload"))

            if "pptx_bytes" in st.session_state:
                with download_area.container():
                    st.success("Presentation ready!")
                    st.download_button(
                        "⬇️ Download PPTX",
                        data=st.session_state["pptx_bytes"],
                        file_name=st.session_state.get("pptx_filename","presentation.pptx"),
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        use_container_width=True,
                    )