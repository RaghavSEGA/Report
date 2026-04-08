"""
storage_pptx.py — Supabase/Postgres-backed persistence for the SEGA PowerPoint Creator.

Connection is via DATABASE_URL in st.secrets (standard Postgres DSN).
Example secrets.toml entry:
    DATABASE_URL = "postgresql://user:password@host:5432/dbname"

On Supabase: use the "Session mode" connection string from
Settings → Database → Connection string → URI (port 5432, NOT 6543).

Tables created on first run:
  - pptx_projects  : project metadata + slide JSON + chat history
  - (pptx_bytes and template_bytes are stored as base64 TEXT to avoid
     bytea encoding complexity — they are typically <5 MB each)
"""

import json
import base64
import streamlit as st
import psycopg2
import psycopg2.extras
from contextlib import contextmanager


# ── Connection ────────────────────────────────────────────────────────────────

@contextmanager
def _get_conn():
    """Yield a psycopg2 connection from the DATABASE_URL secret."""
    url = st.secrets.get("DATABASE_URL", "")
    if not url:
        raise RuntimeError(
            "DATABASE_URL not set in secrets.toml.\n"
            "Add: DATABASE_URL = \"postgresql://user:pass@host:5432/dbname\""
        )
    conn = psycopg2.connect(
        url,
        sslmode="require",
        options="-c statement_timeout=30000",
    )
    conn.autocommit = False
    psycopg2.extras.register_default_jsonb(conn)
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


# ── Schema ────────────────────────────────────────────────────────────────────

def init_db():
    """Create tables if they don't exist. Safe to call on every startup."""
    with _get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute("""
                CREATE TABLE IF NOT EXISTS pptx_projects (
                    owner           TEXT    NOT NULL,
                    name            TEXT    NOT NULL,
                    business_question TEXT  DEFAULT '',
                    game_title      TEXT    DEFAULT '',
                    audience        TEXT    DEFAULT '',
                    doc_names       JSONB   DEFAULT '[]',
                    slide_json      JSONB   DEFAULT '{}',
                    plan_chat       JSONB   DEFAULT '[]',
                    pptx_b64        TEXT    DEFAULT NULL,
                    template_b64    TEXT    DEFAULT NULL,
                    updated_at      TIMESTAMPTZ DEFAULT NOW(),
                    PRIMARY KEY (owner, name)
                )
            """)
            cur.execute("""
                CREATE INDEX IF NOT EXISTS idx_pptx_projects_owner
                ON pptx_projects (owner)
            """)


# ── Helpers ───────────────────────────────────────────────────────────────────

def _b64enc(data: bytes | None) -> str | None:
    if not data:
        return None
    return base64.b64encode(data).decode("ascii")


def _b64dec(s: str | None) -> bytes | None:
    if not s:
        return None
    return base64.b64decode(s.encode("ascii"))


# ── Projects ──────────────────────────────────────────────────────────────────

def get_projects(owner: str) -> list[dict]:
    """Return list of {name, ...metadata} for this owner, sorted by name."""
    with _get_conn() as conn:
        with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
            cur.execute("""
                SELECT name, business_question, game_title, audience,
                       doc_names, updated_at
                FROM pptx_projects
                WHERE owner = %s
                ORDER BY name
            """, (owner,))
            rows = cur.fetchall()
    return [dict(r) for r in rows]


def project_exists(owner: str, name: str) -> bool:
    with _get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute(
                "SELECT 1 FROM pptx_projects WHERE owner = %s AND name = %s",
                (owner, name)
            )
            return cur.fetchone() is not None


def create_project(owner: str, name: str):
    """Insert a new blank project row."""
    with _get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute("""
                INSERT INTO pptx_projects (owner, name)
                VALUES (%s, %s)
                ON CONFLICT (owner, name) DO NOTHING
            """, (owner, name))


def rename_project(owner: str, old_name: str, new_name: str):
    with _get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute("""
                UPDATE pptx_projects SET name = %s, updated_at = NOW()
                WHERE owner = %s AND name = %s
            """, (new_name, owner, old_name))


def delete_project(owner: str, name: str):
    with _get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute(
                "DELETE FROM pptx_projects WHERE owner = %s AND name = %s",
                (owner, name)
            )


def load_project(owner: str, name: str) -> dict | None:
    """
    Return all project data including binary blobs (decoded from base64).
    Returns None if project doesn't exist.
    """
    with _get_conn() as conn:
        with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
            cur.execute("""
                SELECT business_question, game_title, audience,
                       doc_names, slide_json, plan_chat,
                       pptx_b64, template_b64
                FROM pptx_projects
                WHERE owner = %s AND name = %s
            """, (owner, name))
            row = cur.fetchone()
    if not row:
        return None
    return {
        "business_question": row["business_question"] or "",
        "game_title":        row["game_title"] or "",
        "audience":          row["audience"] or "",
        "doc_names":         row["doc_names"] or [],
        "slide_json":        row["slide_json"] or {},
        "plan_chat":         row["plan_chat"] or [],
        "pptx_bytes":        _b64dec(row["pptx_b64"]),
        "template_bytes":    _b64dec(row["template_b64"]),
    }


def save_project(
    owner: str,
    name: str,
    business_question: str = "",
    game_title: str = "",
    audience: str = "",
    doc_names: list = None,
    slide_json: dict = None,
    plan_chat: list = None,
    pptx_bytes: bytes | None = None,
    template_bytes: bytes | None = None,
    clear_pptx: bool = False,
):
    """
    Upsert all project fields.
    Pass clear_pptx=True to null out the stored PPTX without providing new bytes.
    """
    pptx_b64     = None if clear_pptx else _b64enc(pptx_bytes)
    template_b64 = _b64enc(template_bytes)

    with _get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute("""
                INSERT INTO pptx_projects
                    (owner, name, business_question, game_title, audience,
                     doc_names, slide_json, plan_chat,
                     pptx_b64, template_b64, updated_at)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, NOW())
                ON CONFLICT (owner, name) DO UPDATE SET
                    business_question = EXCLUDED.business_question,
                    game_title        = EXCLUDED.game_title,
                    audience          = EXCLUDED.audience,
                    doc_names         = EXCLUDED.doc_names,
                    slide_json        = EXCLUDED.slide_json,
                    plan_chat         = EXCLUDED.plan_chat,
                    pptx_b64          = CASE
                                          WHEN %s THEN NULL
                                          ELSE COALESCE(EXCLUDED.pptx_b64,
                                                        pptx_projects.pptx_b64)
                                        END,
                    template_b64      = COALESCE(EXCLUDED.template_b64,
                                                 pptx_projects.template_b64),
                    updated_at        = NOW()
            """, (
                owner, name,
                business_question, game_title, audience,
                json.dumps(doc_names or []),
                json.dumps(slide_json or {}),
                json.dumps(plan_chat or []),
                pptx_b64,
                template_b64,
                # second pass of clear_pptx for the CASE expression
                clear_pptx,
            ))