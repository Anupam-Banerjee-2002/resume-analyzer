import json
import os
import sqlite3
from contextlib import contextmanager
from datetime import datetime
from typing import Any, Dict, Iterable, Optional, Tuple


DB_PATH = os.environ.get("APP_DB_PATH", os.path.join(os.path.dirname(__file__), "app.db"))


def _dict_factory(cursor, row):
    return {col[0]: row[idx] for idx, col in enumerate(cursor.description)}


@contextmanager
def get_conn():
    conn = sqlite3.connect(DB_PATH, detect_types=sqlite3.PARSE_DECLTYPES)
    conn.row_factory = _dict_factory
    try:
        yield conn
        conn.commit()
    finally:
        conn.close()


def init_db():
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    with get_conn() as conn:
        conn.execute("PRAGMA foreign_keys = ON")

        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS users (
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              name TEXT NOT NULL,
              email TEXT NOT NULL UNIQUE,
              password_hash TEXT NOT NULL,
              role TEXT NOT NULL CHECK (role IN ('HR','HOD','ADMIN')),
              department TEXT NULL,
              is_active INTEGER NOT NULL DEFAULT 1,
              created_at TEXT NOT NULL
            )
            """
        )

        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS resumes (
              id TEXT PRIMARY KEY,
              original_filename TEXT NOT NULL,
              stored_path TEXT NOT NULL,
              mime TEXT NULL,
              size INTEGER NULL,
              uploaded_by INTEGER NULL,
              candidate_name TEXT NULL,
              predicted_department TEXT NOT NULL,
              probabilities_json TEXT NOT NULL,
              confidence REAL NOT NULL,
              kw_score REAL NOT NULL,
              relevance REAL NOT NULL,
              shortlisted INTEGER NOT NULL,
              skills_json TEXT NOT NULL,
              education_json TEXT NOT NULL,
              experience TEXT NOT NULL,
              created_at TEXT NOT NULL,
              FOREIGN KEY (uploaded_by) REFERENCES users(id) ON DELETE SET NULL
            )
            """
        )

        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS keywords (
              department TEXT NOT NULL,
              keyword TEXT NOT NULL,
              created_at TEXT NOT NULL,
              PRIMARY KEY (department, keyword)
            )
            """
        )


def utcnow_iso() -> str:
    return datetime.utcnow().replace(microsecond=0).isoformat() + "Z"


def users_count(conn) -> int:
    row = conn.execute("SELECT COUNT(*) AS c FROM users").fetchone()
    return int(row["c"])


def insert_user(conn, *, name: str, email: str, password_hash: str, role: str, department: Optional[str]) -> int:
    cur = conn.execute(
        """
        INSERT INTO users(name, email, password_hash, role, department, is_active, created_at)
        VALUES(?,?,?,?,?,1,?)
        """,
        (name, email.lower().strip(), password_hash, role, department, utcnow_iso()),
    )
    return int(cur.lastrowid)


def get_user_by_email(conn, email: str) -> Optional[Dict[str, Any]]:
    return conn.execute("SELECT * FROM users WHERE email = ?", (email.lower().strip(),)).fetchone()


def get_user_by_id(conn, user_id: int) -> Optional[Dict[str, Any]]:
    return conn.execute("SELECT * FROM users WHERE id = ?", (user_id,)).fetchone()


def list_users(conn, role: Optional[str] = None) -> Iterable[Dict[str, Any]]:
    if role:
        return conn.execute("SELECT * FROM users WHERE role = ? ORDER BY created_at DESC", (role,)).fetchall()
    return conn.execute("SELECT * FROM users ORDER BY created_at DESC").fetchall()


def delete_user(conn, user_id: int) -> int:
    cur = conn.execute("DELETE FROM users WHERE id = ?", (user_id,))
    return int(cur.rowcount)


def upsert_keywords(conn, department: str, keywords: Iterable[str]):
    department = (department or "").strip()
    now = utcnow_iso()
    for kw in keywords:
        kw = (kw or "").strip().lower()
        if not kw:
            continue
        conn.execute(
            "INSERT OR IGNORE INTO keywords(department, keyword, created_at) VALUES(?,?,?)",
            (department, kw, now),
        )


def set_keywords_for_department(conn, department: str, keywords: Iterable[str]):
    department = (department or "").strip()
    conn.execute("DELETE FROM keywords WHERE department = ?", (department,))
    upsert_keywords(conn, department, keywords)


def get_keywords_by_department(conn, department: str) -> list[str]:
    department = (department or "").strip()
    rows = conn.execute(
        "SELECT keyword FROM keywords WHERE department = ? ORDER BY keyword ASC",
        (department,),
    ).fetchall()
    return [r["keyword"] for r in rows]


def get_all_keywords(conn) -> Dict[str, list[str]]:
    rows = conn.execute(
        "SELECT department, keyword FROM keywords ORDER BY department ASC, keyword ASC"
    ).fetchall()
    out: Dict[str, list[str]] = {}
    for r in rows:
        out.setdefault(r["department"], []).append(r["keyword"])
    return out


def insert_resume(conn, resume: Dict[str, Any]):
    conn.execute(
        """
        INSERT INTO resumes(
          id, original_filename, stored_path, mime, size, uploaded_by,
          candidate_name, predicted_department, probabilities_json, confidence,
          kw_score, relevance, shortlisted,
          skills_json, education_json, experience, created_at
        ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
        """,
        (
            resume["id"],
            resume["original_filename"],
            resume["stored_path"],
            resume.get("mime"),
            resume.get("size"),
            resume.get("uploaded_by"),
            resume.get("candidate_name"),
            resume["predicted_department"],
            json.dumps(resume["probabilities"], ensure_ascii=False),
            float(resume["confidence"]),
            float(resume["kw_score"]),
            float(resume["relevance"]),
            1 if resume["shortlisted"] else 0,
            json.dumps(resume.get("skills") or [], ensure_ascii=False),
            json.dumps(resume.get("education") or [], ensure_ascii=False),
            resume.get("experience") or "Not specified",
            resume.get("created_at") or utcnow_iso(),
        ),
    )


def get_resume_by_id(conn, resume_id: str) -> Optional[Dict[str, Any]]:
    return conn.execute("SELECT * FROM resumes WHERE id = ?", (resume_id,)).fetchone()


def delete_resume(conn, resume_id: str) -> int:
    cur = conn.execute("DELETE FROM resumes WHERE id = ?", (resume_id,))
    return int(cur.rowcount)


def recalculate_shortlisted_by_threshold(conn, threshold: float) -> int:
    """Set shortlisted = 1 when relevance >= threshold (matches predict-time logic)."""
    cur = conn.execute(
        "UPDATE resumes SET shortlisted = CASE WHEN relevance >= ? THEN 1 ELSE 0 END",
        (float(threshold),),
    )
    return int(cur.rowcount)


def query_resumes(
    conn,
    *,
    department: Optional[str],
    shortlisted: Optional[bool],
    search: Optional[str],
    sort: str,
    page: int,
    page_size: int,
    forced_department: Optional[str] = None,
) -> Tuple[int, list[Dict[str, Any]]]:
    where = []
    params: list[Any] = []

    effective_dept = forced_department if forced_department else department
    if effective_dept:
        where.append("predicted_department = ?")
        params.append(effective_dept)

    if shortlisted is True:
        where.append("shortlisted = 1")
    elif shortlisted is False:
        where.append("shortlisted = 0")

    if search:
        s = f"%{search.strip()}%"
        where.append("(candidate_name LIKE ? OR original_filename LIKE ?)")
        params.extend([s, s])

    where_sql = ("WHERE " + " AND ".join(where)) if where else ""

    sort_map = {
        "relevance": "relevance DESC",
        "confidence": "confidence DESC",
        "kw_score": "kw_score DESC",
        "timestamp": "created_at DESC",
        "created_at": "created_at DESC",
    }
    order_sql = sort_map.get((sort or "").strip(), "created_at DESC")

    row = conn.execute(f"SELECT COUNT(*) AS c FROM resumes {where_sql}", params).fetchone()
    total = int(row["c"])

    offset = max(0, (page - 1) * page_size)
    rows = conn.execute(
        f"SELECT * FROM resumes {where_sql} ORDER BY {order_sql} LIMIT ? OFFSET ?",
        params + [page_size, offset],
    ).fetchall()
    return total, rows

