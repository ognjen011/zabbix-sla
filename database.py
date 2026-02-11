"""
Database module for Zabbix SLA Report app.
Handles user authentication, roles, and report history using SQLite.
"""

import hashlib
import json
import secrets
import sqlite3
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path

DB_PATH = Path(__file__).parent / "sla_app.db"


@contextmanager
def get_db():
    """Context manager for database connections."""
    conn = sqlite3.connect(str(DB_PATH))
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def init_db():
    """Create tables and default admin user if they don't exist."""
    with get_db() as conn:
        conn.executescript("""
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT UNIQUE NOT NULL,
                password_hash TEXT NOT NULL,
                salt TEXT NOT NULL,
                role TEXT NOT NULL DEFAULT 'user',
                display_name TEXT NOT NULL DEFAULT '',
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            );

            CREATE TABLE IF NOT EXISTS report_history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                generated_by TEXT NOT NULL,
                report_name TEXT NOT NULL,
                period TEXT NOT NULL,
                groups_list TEXT NOT NULL,
                host_count INTEGER DEFAULT 0,
                generated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                summary_data TEXT,
                detail_data TEXT,
                excel_data BLOB
            );
        """)

        # Create default admin if no users exist
        row = conn.execute("SELECT COUNT(*) as cnt FROM users").fetchone()
        if row["cnt"] == 0:
            pw_hash, salt = hash_password("admin")
            conn.execute(
                "INSERT INTO users (username, password_hash, salt, role, display_name) "
                "VALUES (?, ?, ?, ?, ?)",
                ("admin", pw_hash, salt, "admin", "Administrator"),
            )


# --- Password utilities ---

def hash_password(password: str, salt: str = None) -> tuple[str, str]:
    if salt is None:
        salt = secrets.token_hex(16)
    pw_hash = hashlib.pbkdf2_hmac("sha256", password.encode(), salt.encode(), 260000)
    return pw_hash.hex(), salt


def verify_password(password: str, pw_hash: str, salt: str) -> bool:
    test_hash = hashlib.pbkdf2_hmac("sha256", password.encode(), salt.encode(), 260000)
    return test_hash.hex() == pw_hash


# --- User operations ---

def authenticate(username: str, password: str) -> dict | None:
    """Authenticate user. Returns user dict or None."""
    with get_db() as conn:
        row = conn.execute(
            "SELECT * FROM users WHERE username = ?", (username,)
        ).fetchone()
        if row and verify_password(password, row["password_hash"], row["salt"]):
            return dict(row)
    return None


def get_all_users() -> list[dict]:
    with get_db() as conn:
        rows = conn.execute(
            "SELECT id, username, role, display_name, created_at FROM users ORDER BY id"
        ).fetchall()
        return [dict(r) for r in rows]


def get_user(user_id: int) -> dict | None:
    with get_db() as conn:
        row = conn.execute(
            "SELECT id, username, role, display_name, created_at FROM users WHERE id = ?",
            (user_id,),
        ).fetchone()
        return dict(row) if row else None


def create_user(username: str, password: str, role: str, display_name: str) -> int | None:
    """Create a user. Returns user id or None if username taken."""
    pw_hash, salt = hash_password(password)
    try:
        with get_db() as conn:
            cur = conn.execute(
                "INSERT INTO users (username, password_hash, salt, role, display_name) "
                "VALUES (?, ?, ?, ?, ?)",
                (username, pw_hash, salt, role, display_name),
            )
            return cur.lastrowid
    except sqlite3.IntegrityError:
        return None


def update_user(user_id: int, display_name: str = None, role: str = None, password: str = None) -> bool:
    """Update user fields. Only provided fields are changed."""
    fields = []
    values = []
    if display_name is not None:
        fields.append("display_name = ?")
        values.append(display_name)
    if role is not None:
        fields.append("role = ?")
        values.append(role)
    if password is not None:
        pw_hash, salt = hash_password(password)
        fields.append("password_hash = ?")
        values.append(pw_hash)
        fields.append("salt = ?")
        values.append(salt)
    if not fields:
        return False
    values.append(user_id)
    with get_db() as conn:
        conn.execute(f"UPDATE users SET {', '.join(fields)} WHERE id = ?", values)
        return True


def delete_user(user_id: int) -> bool:
    with get_db() as conn:
        conn.execute("DELETE FROM users WHERE id = ?", (user_id,))
        return True


def change_password(user_id: int, old_password: str, new_password: str) -> bool:
    """Change password after verifying the old one. Returns True on success."""
    with get_db() as conn:
        row = conn.execute("SELECT * FROM users WHERE id = ?", (user_id,)).fetchone()
        if not row:
            return False
        if not verify_password(old_password, row["password_hash"], row["salt"]):
            return False
        pw_hash, salt = hash_password(new_password)
        conn.execute(
            "UPDATE users SET password_hash = ?, salt = ? WHERE id = ?",
            (pw_hash, salt, user_id),
        )
        return True


# --- Report history operations ---

def save_report(
    generated_by: str,
    report_name: str,
    period: str,
    groups_list: list[str],
    host_count: int,
    summary_data: list[dict],
    detail_data: dict,
    excel_data: bytes,
) -> int:
    """Save a generated report to history. Returns report id."""
    with get_db() as conn:
        cur = conn.execute(
            "INSERT INTO report_history "
            "(generated_by, report_name, period, groups_list, host_count, summary_data, detail_data, excel_data) "
            "VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
            (
                generated_by,
                report_name,
                period,
                json.dumps(groups_list),
                host_count,
                json.dumps(summary_data),
                json.dumps(detail_data),
                excel_data,
            ),
        )
        return cur.lastrowid


def get_reports(limit: int = 50, offset: int = 0) -> list[dict]:
    """Get report history (metadata only, no blobs)."""
    with get_db() as conn:
        rows = conn.execute(
            "SELECT id, generated_by, report_name, period, groups_list, host_count, generated_at, summary_data "
            "FROM report_history ORDER BY generated_at DESC LIMIT ? OFFSET ?",
            (limit, offset),
        ).fetchall()
        results = []
        for r in rows:
            d = dict(r)
            d["groups_list"] = json.loads(d["groups_list"])
            d["summary_data"] = json.loads(d["summary_data"]) if d["summary_data"] else []
            results.append(d)
        return results


def get_report(report_id: int) -> dict | None:
    """Get a single report with all data including Excel blob."""
    with get_db() as conn:
        row = conn.execute(
            "SELECT * FROM report_history WHERE id = ?", (report_id,)
        ).fetchone()
        if not row:
            return None
        d = dict(row)
        d["groups_list"] = json.loads(d["groups_list"])
        d["summary_data"] = json.loads(d["summary_data"]) if d["summary_data"] else []
        d["detail_data"] = json.loads(d["detail_data"]) if d["detail_data"] else {}
        return d


def get_report_excel(report_id: int) -> bytes | None:
    """Get just the Excel data for a report."""
    with get_db() as conn:
        row = conn.execute(
            "SELECT excel_data FROM report_history WHERE id = ?", (report_id,)
        ).fetchone()
        return row["excel_data"] if row else None


def delete_report(report_id: int) -> bool:
    with get_db() as conn:
        conn.execute("DELETE FROM report_history WHERE id = ?", (report_id,))
        return True


def get_report_count() -> int:
    with get_db() as conn:
        row = conn.execute("SELECT COUNT(*) as cnt FROM report_history").fetchone()
        return row["cnt"]
