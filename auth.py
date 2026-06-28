from __future__ import annotations

import base64
import hashlib
import hmac
import os
import secrets
import sqlite3
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

import pandas as pd
import streamlit as st

ROLES = ("admin", "director", "viewer")
PBKDF2_ITERATIONS = 260_000


def utc_now() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def db_path(app_dir: Path) -> Path:
    runtime_dir = app_dir / "runtime"
    runtime_dir.mkdir(parents=True, exist_ok=True)
    return runtime_dir / "project_intelligence_hub_mobile_auth.sqlite"


def connect(app_dir: Path) -> sqlite3.Connection:
    conn = sqlite3.connect(db_path(app_dir))
    conn.row_factory = sqlite3.Row
    return conn


def init_auth_db(app_dir: Path) -> None:
    with connect(app_dir) as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT NOT NULL UNIQUE,
                email TEXT NOT NULL UNIQUE,
                full_name TEXT NOT NULL,
                password_hash TEXT NOT NULL,
                role TEXT NOT NULL CHECK(role IN ('admin', 'director', 'viewer')),
                is_active INTEGER NOT NULL DEFAULT 1,
                must_change_password INTEGER NOT NULL DEFAULT 0,
                created_at TEXT NOT NULL,
                last_login_at TEXT
            )
            """
        )
        conn.execute("CREATE INDEX IF NOT EXISTS idx_users_username ON users(username)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_users_email ON users(email)")


def user_count(app_dir: Path) -> int:
    init_auth_db(app_dir)
    with connect(app_dir) as conn:
        return int(conn.execute("SELECT COUNT(*) FROM users").fetchone()[0])


def hash_password(password: str) -> str:
    salt = secrets.token_bytes(16)
    digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, PBKDF2_ITERATIONS)
    return "pbkdf2_sha256${}${}${}".format(
        PBKDF2_ITERATIONS,
        base64.b64encode(salt).decode("ascii"),
        base64.b64encode(digest).decode("ascii"),
    )


def verify_password(password: str, stored_hash: str) -> bool:
    try:
        algorithm, iterations, salt_b64, digest_b64 = stored_hash.split("$", 3)
        if algorithm != "pbkdf2_sha256":
            return False
        salt = base64.b64decode(salt_b64.encode("ascii"))
        expected = base64.b64decode(digest_b64.encode("ascii"))
        actual = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, int(iterations))
        return hmac.compare_digest(actual, expected)
    except (ValueError, TypeError, OSError):
        return False


def password_policy_errors(password: str, confirmation: str | None = None) -> list[str]:
    errors: list[str] = []
    if len(password) < 10:
        errors.append("Password must be at least 10 characters.")
    if not any(char.isupper() for char in password):
        errors.append("Password must include an uppercase letter.")
    if not any(char.islower() for char in password):
        errors.append("Password must include a lowercase letter.")
    if not any(char.isdigit() for char in password):
        errors.append("Password must include a number.")
    if confirmation is not None and password != confirmation:
        errors.append("Password confirmation does not match.")
    return errors


def create_user(
    app_dir: Path,
    *,
    username: str,
    email: str,
    full_name: str,
    password: str,
    role: str,
    is_active: bool = True,
    must_change_password: bool = False,
) -> tuple[bool, str]:
    init_auth_db(app_dir)
    username = username.strip().lower()
    email = email.strip().lower()
    full_name = full_name.strip()
    role = role.strip().lower()
    if not username or not email or not full_name:
        return False, "Username, email, and full name are required."
    if role not in ROLES:
        return False, "Invalid role."
    errors = password_policy_errors(password)
    if errors:
        return False, " ".join(errors)
    try:
        with connect(app_dir) as conn:
            conn.execute(
                """
                INSERT INTO users (
                    username, email, full_name, password_hash, role,
                    is_active, must_change_password, created_at
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    username,
                    email,
                    full_name,
                    hash_password(password),
                    role,
                    1 if is_active else 0,
                    1 if must_change_password else 0,
                    utc_now(),
                ),
            )
        return True, "User account created."
    except sqlite3.IntegrityError:
        return False, "Username or email already exists."


def authenticate(app_dir: Path, username_or_email: str, password: str) -> tuple[bool, str, dict[str, Any] | None]:
    init_auth_db(app_dir)
    lookup = username_or_email.strip().lower()
    with connect(app_dir) as conn:
        row = conn.execute(
            """
            SELECT id, username, email, full_name, password_hash, role, is_active, must_change_password
            FROM users
            WHERE lower(username) = ? OR lower(email) = ?
            """,
            (lookup, lookup),
        ).fetchone()
        if row is None or not verify_password(password, row["password_hash"]):
            return False, "Invalid username/email or password.", None
        if not int(row["is_active"]):
            return False, "This account is inactive. Contact the administrator.", None
        conn.execute("UPDATE users SET last_login_at = ? WHERE id = ?", (utc_now(), row["id"]))
    user = {key: row[key] for key in row.keys() if key != "password_hash"}
    return True, "Login successful.", user


def set_password(app_dir: Path, user_id: int, password: str, must_change_password: bool = False) -> tuple[bool, str]:
    errors = password_policy_errors(password)
    if errors:
        return False, " ".join(errors)
    with connect(app_dir) as conn:
        conn.execute(
            "UPDATE users SET password_hash = ?, must_change_password = ? WHERE id = ?",
            (hash_password(password), 1 if must_change_password else 0, int(user_id)),
        )
    return True, "Password updated."


def update_user(app_dir: Path, user_id: int, role: str, is_active: bool) -> None:
    if role not in ROLES:
        raise ValueError("Invalid role")
    with connect(app_dir) as conn:
        conn.execute(
            "UPDATE users SET role = ?, is_active = ? WHERE id = ?",
            (role, 1 if is_active else 0, int(user_id)),
        )


def users_dataframe(app_dir: Path) -> pd.DataFrame:
    init_auth_db(app_dir)
    with connect(app_dir) as conn:
        rows = conn.execute(
            """
            SELECT id, username, email, full_name, role, is_active, must_change_password, created_at, last_login_at
            FROM users
            ORDER BY id
            """
        ).fetchall()
    return pd.DataFrame([dict(row) for row in rows])


def render_auth_css() -> None:
    st.markdown(
        """
        <style>
        .auth-shell{max-width:980px;margin:0 auto;padding:22px 0 32px}
        .auth-hero{background:linear-gradient(135deg,#07111f,#0f766e);color:#fff;border-radius:14px;padding:24px;margin:8px 0 18px;box-shadow:0 18px 42px rgba(15,23,42,.18)}
        .auth-hero h1{color:#fff;margin:0;font-size:30px}
        .auth-hero p{color:#dbeafe;margin:8px 0 0}
        .auth-footer{text-align:center;color:#64748b;font-size:12px;margin-top:20px}
        .role-pill{display:inline-block;background:#e0f2fe;color:#075985;border-radius:999px;padding:4px 9px;font-weight:800;font-size:12px}
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_footer() -> None:
    st.markdown("<div class='auth-footer'>Created by Engr. Ahmed Labib</div>", unsafe_allow_html=True)


def render_first_admin_setup(app_dir: Path) -> None:
    render_auth_css()
    st.markdown(
        """
        <div class='auth-shell'>
          <div class='auth-hero'>
            <h1>Project Intelligence Hub</h1>
            <p>First-run administrator setup. Create the owner account before the dashboard is opened.</p>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    with st.form("first_admin_setup"):
        full_name = st.text_input("Full name", value="Engr. Ahmed Labib")
        username = st.text_input("Admin username", value="admin")
        email = st.text_input("Admin email")
        password = st.text_input("Password", type="password")
        confirm = st.text_input("Confirm password", type="password")
        submitted = st.form_submit_button("Create Admin Account", type="primary", width="stretch")
    if submitted:
        errors = password_policy_errors(password, confirm)
        if errors:
            st.error(" ".join(errors))
        else:
            ok, message = create_user(
                app_dir,
                username=username,
                email=email,
                full_name=full_name,
                password=password,
                role="admin",
                is_active=True,
            )
            if ok:
                st.success("Admin account created. Sign in to continue.")
                st.session_state["auth_view"] = "login"
                st.rerun()
            else:
                st.error(message)
    render_footer()


def render_login(app_dir: Path) -> None:
    render_auth_css()
    st.markdown(
        """
        <div class='auth-shell'>
          <div class='auth-hero'>
            <h1>Project Intelligence Hub</h1>
            <p>Secure mobile-ready project controls intelligence platform.</p>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    tabs = st.tabs(["Login", "Sign Up"])
    with tabs[0]:
        with st.form("login_form"):
            username = st.text_input("Username or email")
            password = st.text_input("Password", type="password")
            submitted = st.form_submit_button("Login", type="primary", width="stretch")
        if submitted:
            ok, message, user = authenticate(app_dir, username, password)
            if ok and user:
                st.session_state["auth_user"] = user
                st.success(message)
                st.rerun()
            else:
                st.error(message)
    with tabs[1]:
        with st.form("signup_form"):
            full_name = st.text_input("Full name")
            username = st.text_input("Username")
            email = st.text_input("Email")
            password = st.text_input("Password", type="password")
            confirm = st.text_input("Confirm password", type="password")
            submitted = st.form_submit_button("Create Viewer Account", width="stretch")
        if submitted:
            errors = password_policy_errors(password, confirm)
            if errors:
                st.error(" ".join(errors))
            else:
                ok, message = create_user(
                    app_dir,
                    username=username,
                    email=email,
                    full_name=full_name,
                    password=password,
                    role="viewer",
                    is_active=True,
                )
                if ok:
                    st.success("Account created. You can now log in.")
                else:
                    st.error(message)
    render_footer()


def render_admin_account_management(app_dir: Path) -> None:
    with st.sidebar.expander("Admin Account Management", expanded=False):
        users_df = users_dataframe(app_dir)
        if users_df.empty:
            st.info("No users found.")
        else:
            st.dataframe(users_df, width="stretch", hide_index=True)

        st.markdown("**Create user**")
        with st.form("admin_create_user"):
            full_name = st.text_input("Full name", key="admin_create_full_name")
            username = st.text_input("Username", key="admin_create_username")
            email = st.text_input("Email", key="admin_create_email")
            role = st.selectbox("Role", ROLES, index=1, key="admin_create_role")
            password = st.text_input("Temporary password", type="password", key="admin_create_password")
            submitted = st.form_submit_button("Create Account")
        if submitted:
            ok, message = create_user(
                app_dir,
                username=username,
                email=email,
                full_name=full_name,
                password=password,
                role=role,
                is_active=True,
                must_change_password=True,
            )
            if ok:
                st.success("Account created. Share credentials through a secure channel.")
                st.rerun()
            else:
                st.error(message)

        if not users_df.empty:
            st.markdown("**Edit user**")
            user_label = st.selectbox(
                "User",
                users_df["username"].tolist(),
                key="admin_edit_user",
            )
            selected = users_df[users_df["username"] == user_label].iloc[0]
            new_role = st.selectbox(
                "Role",
                ROLES,
                index=ROLES.index(str(selected["role"])),
                key="admin_edit_role",
            )
            active = st.checkbox("Active", value=bool(selected["is_active"]), key="admin_edit_active")
            if st.button("Save User Changes", key="admin_save_user", width="stretch"):
                update_user(app_dir, int(selected["id"]), new_role, active)
                st.success("User updated.")
                st.rerun()
            reset_password = st.text_input("New password", type="password", key="admin_reset_password")
            if st.button("Reset Password", key="admin_reset_password_button", width="stretch"):
                ok, message = set_password(app_dir, int(selected["id"]), reset_password, must_change_password=True)
                if ok:
                    st.success("Password reset. Share it securely and ask the user to change it.")
                else:
                    st.error(message)


def current_user() -> dict[str, Any] | None:
    user = st.session_state.get("auth_user")
    return user if isinstance(user, dict) else None


def is_authenticated() -> bool:
    return current_user() is not None


def logout() -> None:
    st.session_state.pop("auth_user", None)


def require_authentication(app_dir: Path) -> dict[str, Any]:
    init_auth_db(app_dir)
    if user_count(app_dir) == 0:
        render_first_admin_setup(app_dir)
        st.stop()
    user = current_user()
    if not user:
        render_login(app_dir)
        st.stop()

    with st.sidebar:
        st.markdown(f"**{user.get('full_name', user.get('username', 'User'))}**")
        st.markdown(f"<span class='role-pill'>{str(user.get('role', 'viewer')).upper()}</span>", unsafe_allow_html=True)
        if st.button("Logout", width="stretch"):
            logout()
            st.rerun()
        if user.get("role") == "admin":
            render_admin_account_management(app_dir)
        elif user.get("role") == "viewer":
            st.info("Viewer mode: read-only dashboard access.")
        elif user.get("role") == "director":
            st.info("Director mode: executive dashboard access.")
    return user


def apply_role_ui(user: dict[str, Any]) -> None:
    role = str(user.get("role", "viewer")).lower()
    if role == "viewer":
        st.markdown(
            """
            <style>
            .stFileUploader, .stDownloadButton {display:none!important;}
            </style>
            """,
            unsafe_allow_html=True,
        )
    if role == "director":
        st.session_state.setdefault("active_project_slide_name", "Output Studio")
