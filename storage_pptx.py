"""
storage_pptx.py — Postgres-backed persistence for the SEGA PowerPoint Creator.

Connection is via DATABASE_URL in st.secrets (standard Postgres DSN).
Example secrets.toml entry:
    DATABASE_URL = "postgresql://user:password@host:5432/dbname"

On Supabase: use the Session mode connection string from
Settings -> Database -> Connection string -> URI (port 5432, NOT 6543).

Run this SQL once in Supabase SQL Editor before first use:

    CREATE TABLE IF NOT EXISTS pptx_projects (
        owner             TEXT        NOT NULL,
        name              TEXT        NOT NULL,
        business_question TEXT        DEFAULT '',
        game_title        TEXT        DEFAULT '',
        audience          TEXT        DEFAULT '',
        doc_names         JSONB       DEFAULT '[]',
        slide_json        JSONB       DEFAULT '{}',
        plan_chat         JSONB       DEFAULT '[]',
        pptx_b64          TEXT        DEFAULT NULL,
        template_b64      TEXT        DEFAULT NULL,
        updated_at        TIMESTAMPTZ DEFAULT NOW(),
        PRIMARY KEY (owner, name)
    );
    CREATE INDEX IF NOT EXISTS idx_pptx_projects_owner ON pptx_projects (owner);
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


# ── Binary helpers ────────────────────────────────────────────────────────────

def _enc(data):
    if not data:
        return None
    return base64.b64encode(data).decode("ascii")


def _dec(s):
    if not s:
        return None
    try:
        return base64.b64decode(s.encode("ascii"))
    except Exception:
        return None


# ── Schema ────────────────────────────────────────────────────────────────────

def init_db():
    """Probe that the table exists. Create it if missing."""
    with _get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute("""
                CREATE TABLE IF NOT EXISTS pptx_projects (
                    owner             TEXT        NOT NULL,
                    name              TEXT        NOT NULL,
                    business_question TEXT        DEFAULT '',
                    game_title        TEXT        DEFAULT '',
                    audience          TEXT        DEFAULT '',
                    doc_names         JSONB       DEFAULT '[]',
                    slide_json        JSONB       DEFAULT '{}',
                    plan_chat         JSONB       DEFAULT '[]',
                    pptx_b64          TEXT        DEFAULT NULL,
                    template_b64      TEXT        DEFAULT NULL,
                    updated_at        TIMESTAMPTZ DEFAULT NOW(),
                    PRIMARY KEY (owner, name)
                )
            """)
            cur.execute("""
                CREATE INDEX IF NOT EXISTS idx_pptx_projects_owner
                ON pptx_projects (owner)
            """)


# ── Projects ──────────────────────────────────────────────────────────────────

def get_projects(owner: str) -> list:
    with _get_conn() as conn:
        with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
            cur.execute(
                """SELECT name, business_question, game_title, audience,
                          doc_names, updated_at
                   FROM pptx_projects WHERE owner = %s ORDER BY name""",
                (owner,)
            )
            return [
                {**dict(r), "updated_at": r["updated_at"].isoformat() if r["updated_at"] else ""}
                for r in cur.fetchall()
            ]


def project_exists(owner: str, name: str) -> bool:
    with _get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute(
                "SELECT 1 FROM pptx_projects WHERE owner = %s AND name = %s",
                (owner, name)
            )
            return cur.fetchone() is not None


def create_project(owner: str, name: str):
    with _get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """INSERT INTO pptx_projects (owner, name)
                   VALUES (%s, %s) ON CONFLICT DO NOTHING""",
                (owner, name)
            )


def rename_project(owner: str, old_name: str, new_name: str):
    with _get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """UPDATE pptx_projects SET name = %s, updated_at = NOW()
                   WHERE owner = %s AND name = %s""",
                (new_name, owner, old_name)
            )


def delete_project(owner: str, name: str):
    with _get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute(
                "DELETE FROM pptx_projects WHERE owner = %s AND name = %s",
                (owner, name)
            )


def load_project(owner: str, name: str):
    with _get_conn() as conn:
        with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
            cur.execute(
                """SELECT business_question, game_title, audience,
                          doc_names, slide_json, plan_chat,
                          pptx_b64, template_b64
                   FROM pptx_projects WHERE owner = %s AND name = %s""",
                (owner, name)
            )
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
        "pptx_bytes":        _dec(row["pptx_b64"]),
        "template_bytes":    _dec(row["template_b64"]),
    }


def save_project(
    owner, name,
    business_question="", game_title="", audience="",
    doc_names=None, slide_json=None, plan_chat=None,
    pptx_bytes=None, template_bytes=None, clear_pptx=False,
):
    """
    Upsert all project fields.
    clear_pptx=True nulls the stored PPTX.
    If pptx_bytes/template_bytes are None (and clear_pptx is False),
    existing blobs are preserved via COALESCE.
    """
    pptx_b64 = None if clear_pptx else _enc(pptx_bytes)
    tmpl_b64 = _enc(template_bytes)

    with _get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute("""
                INSERT INTO pptx_projects
                    (owner, name, business_question, game_title, audience,
                     doc_names, slide_json, plan_chat,
                     pptx_b64, template_b64, updated_at)
                VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,NOW())
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
                pptx_b64, tmpl_b64,
                clear_pptx,   # second use in CASE expression
            ))