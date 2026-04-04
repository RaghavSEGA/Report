"""
storage_pptx.py
===============
SQLite-backed project storage for the SEGA PowerPoint Creator.
Mirrors the pattern from the Twitter Sentiment tool's storage.py.

Each project is scoped to an owner (email address) and stores:
  - name            : project display name
  - business_question: the question used to drive the analysis
  - doc_names       : JSON list of uploaded document filenames
  - slide_json      : JSON string of the slide plan (plan_slide_data)
  - pptx_bytes      : raw bytes of the generated PPTX (nullable)
  - template_bytes  : raw bytes of the custom template (nullable)
  - created_at      : ISO timestamp
  - updated_at      : ISO timestamp
"""

import sqlite3
import json
from pathlib import Path
from datetime import datetime, timezone

DB_PATH = Path("sega_pptx.db")


def _now() -> str:
    return datetime.now(timezone.utc).isoformat()


def get_conn() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db() -> None:
    with get_conn() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS projects (
                owner             TEXT    NOT NULL,
                name              TEXT    NOT NULL,
                business_question TEXT    NOT NULL DEFAULT '',
                doc_names         TEXT    NOT NULL DEFAULT '[]',
                slide_json        TEXT    NOT NULL DEFAULT '{}',
                pptx_bytes        BLOB,
                template_bytes    BLOB,
                created_at        TEXT    NOT NULL,
                updated_at        TEXT    NOT NULL,
                PRIMARY KEY (owner, name)
            )
        """)
        conn.commit()

    # Run migrations in a fresh connection so PRAGMA table_info reflects
    # the committed schema (avoids "duplicate column" on partial-migrated DBs)
    with get_conn() as conn:
        _existing_cols = {r[1] for r in conn.execute("PRAGMA table_info(projects)").fetchall()}
        migrations = [
            ("game_title", "ALTER TABLE projects ADD COLUMN game_title TEXT NOT NULL DEFAULT ''"),
            ("audience",   "ALTER TABLE projects ADD COLUMN audience TEXT NOT NULL DEFAULT 'Executive team'"),
            ("plan_chat",  "ALTER TABLE projects ADD COLUMN plan_chat TEXT NOT NULL DEFAULT '[]'"),
        ]
        for col, sql in migrations:
            if col not in _existing_cols:
                conn.execute(sql)
        conn.commit()


# ── CRUD ──────────────────────────────────────────────────────

def get_projects(owner: str) -> list[dict]:
    """Return all projects for this owner, ordered by updated_at desc."""
    with get_conn() as conn:
        rows = conn.execute(
            """SELECT name, business_question, game_title, audience, doc_names, slide_json,
                      pptx_bytes, template_bytes, plan_chat, created_at, updated_at
               FROM projects
               WHERE owner = ?
               ORDER BY updated_at DESC""",
            (owner,),
        ).fetchall()
    return [dict(r) for r in rows]


def project_exists(owner: str, name: str) -> bool:
    with get_conn() as conn:
        row = conn.execute(
            "SELECT 1 FROM projects WHERE owner = ? AND name = ?",
            (owner, name),
        ).fetchone()
    return row is not None


def create_project(owner: str, name: str) -> None:
    """Create a blank project. Raises ValueError if name already taken."""
    if project_exists(owner, name):
        raise ValueError(f"A project named '{name}' already exists.")
    ts = _now()
    with get_conn() as conn:
        conn.execute(
            """INSERT INTO projects
               (owner, name, business_question, game_title, audience, doc_names, slide_json,
                pptx_bytes, template_bytes, created_at, updated_at)
               VALUES (?, ?, '', '', 'Executive team', '[]', '{}', NULL, NULL, ?, ?)""",
            (owner, name, ts, ts),
        )
        conn.commit()


def rename_project(owner: str, old_name: str, new_name: str) -> None:
    """Rename a project. Raises ValueError if new_name already taken."""
    if old_name == new_name:
        return
    if project_exists(owner, new_name):
        raise ValueError(f"A project named '{new_name}' already exists.")
    with get_conn() as conn:
        conn.execute(
            "UPDATE projects SET name = ?, updated_at = ? WHERE owner = ? AND name = ?",
            (new_name, _now(), owner, old_name),
        )
        conn.commit()


def delete_project(owner: str, name: str) -> None:
    with get_conn() as conn:
        conn.execute(
            "DELETE FROM projects WHERE owner = ? AND name = ?",
            (owner, name),
        )
        conn.commit()


def load_project(owner: str, name: str) -> dict | None:
    """Load a project's full state. Returns None if not found."""
    with get_conn() as conn:
        row = conn.execute(
            """SELECT name, business_question, game_title, audience, doc_names, slide_json,
                      pptx_bytes, template_bytes, plan_chat, created_at, updated_at
               FROM projects WHERE owner = ? AND name = ?""",
            (owner, name),
        ).fetchone()
    if not row:
        return None
    d = dict(row)
    # Parse JSON fields
    d["doc_names"]  = json.loads(d["doc_names"]  or "[]")
    d["slide_json"] = json.loads(d["slide_json"] or "{}")
    d["plan_chat"]  = json.loads(d["plan_chat"]  or "[]")
    return d


def save_project(
    owner: str,
    name: str,
    *,
    business_question: str | None = None,
    game_title: str | None = None,
    audience: str | None = None,
    doc_names: list[str] | None = None,
    slide_json: dict | None = None,
    plan_chat: list | None = None,
    pptx_bytes: bytes | None = None,
    template_bytes: bytes | None = None,
    clear_pptx: bool = False,
    clear_template: bool = False,
) -> None:
    """
    Upsert a project's state. Only provided (non-None) fields are updated.
    Pass clear_pptx=True to NULL out saved PPTX bytes.
    Pass clear_template=True to NULL out saved template bytes.
    """
    # Load existing row to merge
    existing = load_project(owner, name) or {}

    bq  = business_question if business_question is not None else existing.get("business_question", "")
    gt  = game_title if game_title is not None else existing.get("game_title", "")
    aud = audience if audience is not None else existing.get("audience", "Executive team")
    dns = json.dumps(doc_names if doc_names is not None else existing.get("doc_names", []))
    sj  = json.dumps(slide_json if slide_json is not None else existing.get("slide_json", {}))
    pc  = json.dumps(plan_chat if plan_chat is not None else existing.get("plan_chat", []))
    pb  = None if clear_pptx else (pptx_bytes if pptx_bytes is not None else existing.get("pptx_bytes"))
    tb  = None if clear_template else (template_bytes if template_bytes is not None else existing.get("template_bytes"))

    ts  = _now()
    created = existing.get("created_at", ts)

    with get_conn() as conn:
        conn.execute(
            """INSERT INTO projects
               (owner, name, business_question, game_title, audience, doc_names, slide_json,
                pptx_bytes, template_bytes, plan_chat, created_at, updated_at)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
               ON CONFLICT(owner, name) DO UPDATE SET
                   business_question = excluded.business_question,
                   game_title        = excluded.game_title,
                   audience          = excluded.audience,
                   doc_names         = excluded.doc_names,
                   slide_json        = excluded.slide_json,
                   pptx_bytes        = excluded.pptx_bytes,
                   template_bytes    = excluded.template_bytes,
                   plan_chat         = excluded.plan_chat,
                   updated_at        = excluded.updated_at""",
            (owner, name, bq, gt, aud, dns, sj, pb, tb, pc, created, ts),
        )
        conn.commit()