"""
storage_pptx.py — Postgres-backed persistence for the SEGA PowerPoint Creator.
New columns: industry, guided_chat, sources_json, web_research
"""

import json
import base64
import streamlit as st
import psycopg2
import psycopg2.extras
from contextlib import contextmanager


@contextmanager
def _get_conn():
    url = st.secrets.get("DATABASE_URL", "")
    if not url:
        raise RuntimeError("DATABASE_URL not set in secrets.toml.")
    conn = psycopg2.connect(url, sslmode="require", options="-c statement_timeout=30000")
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


def _enc(data):
    return base64.b64encode(data).decode("ascii") if data else None

def _dec(s):
    try:
        return base64.b64decode(s.encode("ascii")) if s else None
    except Exception:
        return None


def init_db():
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
            cur.execute("CREATE INDEX IF NOT EXISTS idx_pptx_projects_owner ON pptx_projects (owner)")
            # Migrate: add new columns to existing tables
            for stmt in [
                "ALTER TABLE pptx_projects ADD COLUMN IF NOT EXISTS industry     TEXT    DEFAULT ''",
                "ALTER TABLE pptx_projects ADD COLUMN IF NOT EXISTS guided_chat  JSONB   DEFAULT '[]'",
                "ALTER TABLE pptx_projects ADD COLUMN IF NOT EXISTS sources_json JSONB   DEFAULT '[]'",
                "ALTER TABLE pptx_projects ADD COLUMN IF NOT EXISTS web_research BOOLEAN DEFAULT TRUE",
            ]:
                cur.execute(stmt)


def get_projects(owner: str) -> list:
    with _get_conn() as conn:
        with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
            cur.execute(
                """SELECT name, business_question, game_title, industry,
                          audience, doc_names, updated_at
                   FROM pptx_projects WHERE owner = %s
                   ORDER BY updated_at DESC NULLS LAST""",
                (owner,)
            )
            return [
                {**dict(r), "updated_at": r["updated_at"].isoformat() if r["updated_at"] else ""}
                for r in cur.fetchall()
            ]


def project_exists(owner: str, name: str) -> bool:
    with _get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute("SELECT 1 FROM pptx_projects WHERE owner=%s AND name=%s", (owner, name))
            return cur.fetchone() is not None


def create_project(owner: str, name: str):
    with _get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute("INSERT INTO pptx_projects (owner,name) VALUES (%s,%s) ON CONFLICT DO NOTHING", (owner, name))


def rename_project(owner: str, old_name: str, new_name: str):
    with _get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute("UPDATE pptx_projects SET name=%s,updated_at=NOW() WHERE owner=%s AND name=%s",
                        (new_name, owner, old_name))


def delete_project(owner: str, name: str):
    with _get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute("DELETE FROM pptx_projects WHERE owner=%s AND name=%s", (owner, name))


def load_project(owner: str, name: str) -> dict | None:
    with _get_conn() as conn:
        with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
            cur.execute(
                """SELECT business_question, game_title, industry, audience,
                          doc_names, slide_json, plan_chat, guided_chat,
                          sources_json, web_research, pptx_b64, template_b64
                   FROM pptx_projects WHERE owner=%s AND name=%s""",
                (owner, name)
            )
            row = cur.fetchone()
    if not row:
        return None
    return {
        "business_question": row["business_question"] or "",
        "game_title":        row["game_title"] or "",
        "industry":          row.get("industry") or row.get("game_title") or "",
        "audience":          row["audience"] or "",
        "doc_names":         row["doc_names"] or [],
        "slide_json":        row["slide_json"] or {},
        "plan_chat":         row["plan_chat"] or [],
        "guided_chat":       row.get("guided_chat") or [],
        "sources":           row.get("sources_json") or [],
        "web_research":      row["web_research"] if row.get("web_research") is not None else True,
        "pptx_bytes":        _dec(row["pptx_b64"]),
        "template_bytes":    _dec(row["template_b64"]),
    }


def save_project(
    owner: str, name: str,
    business_question: str = "", game_title: str = "", industry: str = "",
    audience: str = "", doc_names: list | None = None,
    slide_json: dict | None = None, plan_chat: list | None = None,
    guided_chat: list | None = None, sources: list | None = None,
    web_research: bool = True,
    pptx_bytes: bytes | None = None, template_bytes: bytes | None = None,
    clear_pptx: bool = False,
):
    pptx_b64 = None if clear_pptx else _enc(pptx_bytes)
    tmpl_b64 = _enc(template_bytes)
    with _get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute("""
                INSERT INTO pptx_projects
                    (owner, name, business_question, game_title, industry,
                     audience, doc_names, slide_json, plan_chat, guided_chat,
                     sources_json, web_research, pptx_b64, template_b64, updated_at)
                VALUES (%s,%s,%s,%s,%s, %s,%s,%s,%s,%s, %s,%s, %s,%s, NOW())
                ON CONFLICT (owner, name) DO UPDATE SET
                    business_question = EXCLUDED.business_question,
                    game_title        = EXCLUDED.game_title,
                    industry          = EXCLUDED.industry,
                    audience          = EXCLUDED.audience,
                    doc_names         = EXCLUDED.doc_names,
                    slide_json        = EXCLUDED.slide_json,
                    plan_chat         = EXCLUDED.plan_chat,
                    guided_chat       = EXCLUDED.guided_chat,
                    sources_json      = EXCLUDED.sources_json,
                    web_research      = EXCLUDED.web_research,
                    pptx_b64          = CASE WHEN %s THEN NULL
                                        ELSE COALESCE(EXCLUDED.pptx_b64, pptx_projects.pptx_b64) END,
                    template_b64      = COALESCE(EXCLUDED.template_b64, pptx_projects.template_b64),
                    updated_at        = NOW()
            """, (
                owner, name,
                business_question, game_title or industry, industry,
                audience,
                json.dumps(doc_names or []),
                json.dumps(slide_json or {}),
                json.dumps(plan_chat or []),
                json.dumps(guided_chat or []),
                json.dumps(sources or []),
                web_research,
                pptx_b64, tmpl_b64,
                clear_pptx,
            ))