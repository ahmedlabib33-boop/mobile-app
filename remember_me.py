"""
Persistent "Keep me signed in" support for Project Intelligence Hub.

Works across:
- Web browser
- iPhone PWA / Add to Home Screen
- Android WebView

Security model:
- Raw token is stored only in browser cookie.
- Only SHA-256 token hash is stored in SQLite.
- User remains signed in until manual sign out, token expiry, admin revoke, or account disable.
"""

from __future__ import annotations

import hashlib
import secrets
import sqlite3
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Optional, Dict, Any, Tuple

import streamlit as st

try:
    from streamlit_cookies_controller import CookieController
except Exception:
    CookieController = None


REMEMBER_ME_DAYS = 365
REMEMBER_COOKIE_NAME = "project_intelligence_hub_remember_token"


def _utc_now() -> str:
    return datetime.now(timezone.utc).isoformat()


def _utc_plus_days(days: int) -> str:
    return (datetime.now(timezone.utc) + timedelta(days=days)).isoformat()


def _hash_token(raw_token: str) -> str:
    return hashlib.sha256(raw_token.encode("utf-8")).hexdigest()


def _db_path() -> Path:
    """
    Default database path.
    Adjust this if your auth system uses a different database name.
    """
    root = Path(__file__).resolve().parent
    preferred = root / "app_auth.db"
    fallback = root / "users.db"

    if preferred.exists():
        return preferred

    if fallback.exists():
        return fallback

    return preferred


def get_connection() -> sqlite3.Connection:
    db = _db_path()
    conn = sqlite3.connect(db, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn


def init_remember_sessions_table() -> None:
    conn = get_connection()
    cur = conn.cursor()

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS remember_sessions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL,
            selector TEXT NOT NULL UNIQUE,
            token_hash TEXT NOT NULL,
            device_label TEXT,
            user_agent TEXT,
            created_at TEXT NOT NULL,
            expires_at TEXT NOT NULL,
            last_used_at TEXT,
            revoked_at TEXT,
            is_active INTEGER NOT NULL DEFAULT 1
        )
        """
    )

    conn.commit()
    conn.close()


def _get_cookie_controller():
    if CookieController is None:
        return None

    if "_cookie_controller" not in st.session_state:
        st.session_state["_cookie_controller"] = CookieController()

    return st.session_state["_cookie_controller"]


def get_remember_cookie() -> Optional[str]:
    controller = _get_cookie_controller()

    if controller is None:
        return None

    try:
        return controller.get(REMEMBER_COOKIE_NAME)
    except Exception:
        return None


def set_remember_cookie(token_value: str) -> None:
    controller = _get_cookie_controller()

    if controller is None:
        st.warning("Keep me signed in requires streamlit-cookies-controller. Please install requirements.txt.")
        return

    try:
        max_age = REMEMBER_ME_DAYS * 24 * 60 * 60
        controller.set(
            REMEMBER_COOKIE_NAME,
            token_value,
            max_age=max_age,
            secure=True,
            same_site="Lax",
        )
    except TypeError:
        # Compatibility fallback for older versions of the cookie package
        controller.set(REMEMBER_COOKIE_NAME, token_value)
    except Exception:
        pass


def clear_remember_cookie() -> None:
    controller = _get_cookie_controller()

    if controller is None:
        return

    try:
        controller.remove(REMEMBER_COOKIE_NAME)
    except Exception:
        try:
            controller.set(REMEMBER_COOKIE_NAME, "", max_age=0)
        except Exception:
            pass


def create_remember_session(
    user_id: int,
    user_agent: str = "",
    device_label: str = "Trusted device",
) -> str:
    init_remember_sessions_table()

    selector = secrets.token_urlsafe(12)
    raw_token = secrets.token_urlsafe(48)
    token_hash = _hash_token(raw_token)

    conn = get_connection()
    cur = conn.cursor()

    cur.execute(
        """
        INSERT INTO remember_sessions
        (user_id, selector, token_hash, device_label, user_agent, created_at, expires_at, last_used_at, revoked_at, is_active)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, NULL, 1)
        """,
        (
            user_id,
            selector,
            token_hash,
            device_label,
            user_agent,
            _utc_now(),
            _utc_plus_days(REMEMBER_ME_DAYS),
            _utc_now(),
        ),
    )

    conn.commit()
    conn.close()

    return f"{selector}:{raw_token}"


def _parse_token(token_value: str) -> Optional[Tuple[str, str]]:
    if not token_value or ":" not in token_value:
        return None

    selector, raw_token = token_value.split(":", 1)

    if not selector or not raw_token:
        return None

    return selector, raw_token


def validate_remember_token(token_value: str) -> Optional[Dict[str, Any]]:
    """
    Returns user dict if token is valid.
    Returns None if invalid.
    """
    init_remember_sessions_table()

    parsed = _parse_token(token_value)

    if parsed is None:
        return None

    selector, raw_token = parsed
    token_hash = _hash_token(raw_token)

    conn = get_connection()
    cur = conn.cursor()

    cur.execute(
        """
        SELECT rs.*, u.username, u.email, u.full_name, u.role, u.is_active AS user_is_active
        FROM remember_sessions rs
        JOIN users u ON u.id = rs.user_id
        WHERE rs.selector = ?
        LIMIT 1
        """,
        (selector,),
    )

    row = cur.fetchone()

    if row is None:
        conn.close()
        return None

    if int(row["is_active"]) != 1:
        conn.close()
        return None

    if row["revoked_at"] is not None:
        conn.close()
        return None

    if int(row["user_is_active"]) != 1:
        conn.close()
        return None

    if row["token_hash"] != token_hash:
        conn.close()
        return None

    try:
        expires_at = datetime.fromisoformat(row["expires_at"])
        if expires_at < datetime.now(timezone.utc):
            conn.close()
            return None
    except Exception:
        conn.close()
        return None

    cur.execute(
        """
        UPDATE remember_sessions
        SET last_used_at = ?
        WHERE id = ?
        """,
        (_utc_now(), row["id"]),
    )

    conn.commit()

    user = {
        "id": row["user_id"],
        "username": row["username"],
        "email": row["email"],
        "full_name": row["full_name"],
        "role": row["role"],
        "is_active": row["user_is_active"],
        "remember_session_id": row["id"],
    }

    conn.close()
    return user


def restore_session_from_remember_cookie() -> bool:
    """
    Run this at app startup before showing login page.
    """
    if st.session_state.get("authenticated"):
        return True

    token_value = get_remember_cookie()

    if not token_value:
        return False

    user = validate_remember_token(token_value)

    if not user:
        clear_remember_cookie()
        return False

    st.session_state["authenticated"] = True
    st.session_state["user_id"] = user["id"]
    st.session_state["username"] = user["username"]
    st.session_state["email"] = user.get("email")
    st.session_state["full_name"] = user.get("full_name")
    st.session_state["role"] = user["role"]
    st.session_state["remember_session_id"] = user.get("remember_session_id")

    return True


def issue_remember_cookie_for_user(
    user_id: int,
    user_agent: str = "",
    device_label: str = "Trusted device",
) -> None:
    token_value = create_remember_session(
        user_id=user_id,
        user_agent=user_agent,
        device_label=device_label,
    )
    set_remember_cookie(token_value)


def revoke_remember_token(token_value: Optional[str] = None) -> None:
    init_remember_sessions_table()

    if token_value is None:
        token_value = get_remember_cookie()

    parsed = _parse_token(token_value or "")

    if parsed is None:
        clear_remember_cookie()
        return

    selector, _ = parsed

    conn = get_connection()
    cur = conn.cursor()

    cur.execute(
        """
        UPDATE remember_sessions
        SET is_active = 0, revoked_at = ?
        WHERE selector = ?
        """,
        (_utc_now(), selector),
    )

    conn.commit()
    conn.close()

    clear_remember_cookie()


def revoke_all_user_tokens(user_id: int) -> None:
    init_remember_sessions_table()

    conn = get_connection()
    cur = conn.cursor()

    cur.execute(
        """
        UPDATE remember_sessions
        SET is_active = 0, revoked_at = ?
        WHERE user_id = ?
        """,
        (_utc_now(), user_id),
    )

    conn.commit()
    conn.close()


def revoke_all_tokens() -> None:
    init_remember_sessions_table()

    conn = get_connection()
    cur = conn.cursor()

    cur.execute(
        """
        UPDATE remember_sessions
        SET is_active = 0, revoked_at = ?
        WHERE is_active = 1
        """,
        (_utc_now(),),
    )

    conn.commit()
    conn.close()


def sign_out_and_clear_session() -> None:
    revoke_remember_token()

    for key in [
        "authenticated",
        "user_id",
        "username",
        "email",
        "full_name",
        "role",
        "remember_session_id",
    ]:
        st.session_state.pop(key, None)


def list_remember_sessions() -> list[dict]:
    init_remember_sessions_table()

    conn = get_connection()
    cur = conn.cursor()

    cur.execute(
        """
        SELECT
            rs.id,
            rs.user_id,
            u.username,
            u.role,
            rs.device_label,
            rs.user_agent,
            rs.created_at,
            rs.expires_at,
            rs.last_used_at,
            rs.revoked_at,
            rs.is_active
        FROM remember_sessions rs
        LEFT JOIN users u ON u.id = rs.user_id
        ORDER BY rs.last_used_at DESC
        """
    )

    rows = [dict(r) for r in cur.fetchall()]
    conn.close()
    return rows


def revoke_session_by_id(session_id: int) -> None:
    init_remember_sessions_table()

    conn = get_connection()
    cur = conn.cursor()

    cur.execute(
        """
        UPDATE remember_sessions
        SET is_active = 0, revoked_at = ?
        WHERE id = ?
        """,
        (_utc_now(), session_id),
    )

    conn.commit()
    conn.close()