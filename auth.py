from __future__ import annotations

import base64
import hashlib
import hmac
import json
import os
import secrets
import sqlite3
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any

import pandas as pd
import streamlit as st

ROLES = ("admin", "director", "viewer")
NON_ADMIN_ROLES = ("director", "viewer")
PBKDF2_ITERATIONS = 260_000
OWNER_USERNAME = "ahmed_labib"
OWNER_EMAIL = "ahmedlabib33@gmail.com"
OWNER_FULL_NAME = "Engr. Ahmed Labib"
OWNER_DEFAULT_PASSWORD = os.getenv("PIH_MOBILE_APP_OWNER_PASSWORD", "Ahmed731988")
REMEMBER_DAYS = 45
DASHBOARD_SECTIONS = [
    "Overview",
    "WBS",
    "Activities",
    "Milestones",
    "S-Curve",
    "EVM Analysis",
    "Contracts",
    "Letters Intelligence",
    "Risks",
    "Delay Analysis - Time Impact Analysis",
    "Contract & Claims Intelligence Center",
    "Output Studio",
    "Data Quality & Export Center",
]
DIRECTOR_DEFAULT_SECTIONS = ["Overview", "EVM Analysis", "Risks", "Output Studio", "Data Quality & Export Center"]
VIEWER_DEFAULT_SECTIONS = ["Overview", "WBS", "Activities", "Milestones", "S-Curve", "EVM Analysis", "Contracts", "Risks"]


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


def normalize_username(username: str) -> str:
    return username.strip().lower()


def normalize_email(email: str) -> str:
    return email.strip().lower()


def default_access_sections(role: str) -> list[str]:
    role = str(role).strip().lower()
    if role == "admin":
        return DASHBOARD_SECTIONS.copy()
    if role == "director":
        return DIRECTOR_DEFAULT_SECTIONS.copy()
    return VIEWER_DEFAULT_SECTIONS.copy()


def encode_sections(sections: list[str] | tuple[str, ...] | None, role: str = "viewer") -> str:
    valid = [section for section in (sections or default_access_sections(role)) if section in DASHBOARD_SECTIONS]
    if not valid:
        valid = default_access_sections(role)
    return json.dumps(valid)


def decode_sections(value: str | None, role: str = "viewer") -> list[str]:
    if not value:
        return default_access_sections(role)
    try:
        loaded = json.loads(value)
    except (TypeError, json.JSONDecodeError):
        return default_access_sections(role)
    if not isinstance(loaded, list):
        return default_access_sections(role)
    sections = [str(item) for item in loaded if str(item) in DASHBOARD_SECTIONS]
    return sections or default_access_sections(role)


def is_owner_identity(username: str = "", email: str = "") -> bool:
    return normalize_username(username) == OWNER_USERNAME or normalize_email(email) == OWNER_EMAIL


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
        existing_columns = {row["name"] for row in conn.execute("PRAGMA table_info(users)").fetchall()}
        migrations = {
            "approved_by": "ALTER TABLE users ADD COLUMN approved_by TEXT",
            "approved_at": "ALTER TABLE users ADD COLUMN approved_at TEXT",
            "access_sections": "ALTER TABLE users ADD COLUMN access_sections TEXT",
            "remember_token_hash": "ALTER TABLE users ADD COLUMN remember_token_hash TEXT",
            "remember_expires_at": "ALTER TABLE users ADD COLUMN remember_expires_at TEXT",
        }
        for column, statement in migrations.items():
            if column not in existing_columns:
                conn.execute(statement)
        conn.execute(
            "UPDATE users SET access_sections = ? WHERE access_sections IS NULL OR access_sections = ''",
            (encode_sections(VIEWER_DEFAULT_SECTIONS, "viewer"),),
        )


def ensure_owner_account(app_dir: Path) -> None:
    init_auth_db(app_dir)
    with connect(app_dir) as conn:
        owner = conn.execute(
            "SELECT id FROM users WHERE lower(username) = ? OR lower(email) = ?",
            (OWNER_USERNAME, OWNER_EMAIL),
        ).fetchone()
        if owner is None:
            conn.execute(
                """
                INSERT INTO users (
                    username, email, full_name, password_hash, role, is_active,
                    must_change_password, created_at, approved_by, approved_at, access_sections
                )
                VALUES (?, ?, ?, ?, 'admin', 1, 0, ?, 'system', ?, ?)
                """,
                (
                    OWNER_USERNAME,
                    OWNER_EMAIL,
                    OWNER_FULL_NAME,
                    hash_password(OWNER_DEFAULT_PASSWORD),
                    utc_now(),
                    utc_now(),
                    encode_sections(DASHBOARD_SECTIONS, "admin"),
                ),
            )
        else:
            conn.execute(
                """
                UPDATE users
                SET username = ?, email = ?, full_name = ?, role = 'admin', is_active = 1,
                    access_sections = ?
                WHERE id = ?
                """,
                (OWNER_USERNAME, OWNER_EMAIL, OWNER_FULL_NAME, encode_sections(DASHBOARD_SECTIONS, "admin"), owner["id"]),
            )
        conn.execute(
            """
            UPDATE users
            SET role = 'viewer'
            WHERE role = 'admin' AND lower(username) <> ? AND lower(email) <> ?
            """,
            (OWNER_USERNAME, OWNER_EMAIL),
        )


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
    is_active: bool = False,
    must_change_password: bool = False,
    access_sections: list[str] | None = None,
    approved_by: str | None = None,
) -> tuple[bool, str]:
    init_auth_db(app_dir)
    username = normalize_username(username)
    email = normalize_email(email)
    full_name = full_name.strip()
    role = role.strip().lower()
    if not username or not email or not full_name:
        return False, "Username, email, and full name are required."
    if role == "admin" and not is_owner_identity(username, email):
        return False, "Only Ahmed Labib can be administrator."
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
                    is_active, must_change_password, created_at, approved_by, approved_at, access_sections
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
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
                    approved_by,
                    utc_now() if is_active else None,
                    encode_sections(access_sections, role),
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
            SELECT id, username, email, full_name, password_hash, role, is_active, must_change_password, access_sections
            FROM users
            WHERE lower(username) = ? OR lower(email) = ?
            """,
            (lookup, lookup),
        ).fetchone()
        if row is None or not verify_password(password, row["password_hash"]):
            return False, "Invalid username/email or password.", None
        if not int(row["is_active"]):
            return False, "Your account is pending administrator approval.", None
        conn.execute("UPDATE users SET last_login_at = ? WHERE id = ?", (utc_now(), row["id"]))
    user = {key: row[key] for key in row.keys() if key != "password_hash"}
    user["access_sections"] = decode_sections(user.get("access_sections"), str(user.get("role", "viewer")))
    return True, "Login successful.", user


def hash_remember_token(token: str) -> str:
    return hashlib.sha256(token.encode("utf-8")).hexdigest()


def create_remember_token(app_dir: Path, user_id: int) -> str:
    token = secrets.token_urlsafe(32)
    expires = datetime.now(timezone.utc) + timedelta(days=REMEMBER_DAYS)
    with connect(app_dir) as conn:
        conn.execute(
            "UPDATE users SET remember_token_hash = ?, remember_expires_at = ? WHERE id = ?",
            (hash_remember_token(token), expires.replace(microsecond=0).isoformat(), int(user_id)),
        )
    return token


def revoke_remember_token(app_dir: Path, user_id: int | None) -> None:
    if not user_id:
        return
    with connect(app_dir) as conn:
        conn.execute("UPDATE users SET remember_token_hash = NULL, remember_expires_at = NULL WHERE id = ?", (int(user_id),))


def user_from_remember_token(app_dir: Path, token: str) -> dict[str, Any] | None:
    token = str(token or "").strip()
    if not token:
        return None
    now = utc_now()
    with connect(app_dir) as conn:
        row = conn.execute(
            """
            SELECT id, username, email, full_name, role, is_active, must_change_password, access_sections, remember_expires_at
            FROM users
            WHERE remember_token_hash = ?
            """,
            (hash_remember_token(token),),
        ).fetchone()
        if row is None or not int(row["is_active"]):
            return None
        if str(row["remember_expires_at"] or "") < now:
            conn.execute("UPDATE users SET remember_token_hash = NULL, remember_expires_at = NULL WHERE id = ?", (row["id"],))
            return None
    user = dict(row)
    user["access_sections"] = decode_sections(user.get("access_sections"), str(user.get("role", "viewer")))
    return user


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


def update_user(app_dir: Path, user_id: int, role: str, is_active: bool, access_sections: list[str] | None = None, approved_by: str | None = None) -> None:
    if role not in ROLES:
        raise ValueError("Invalid role")
    with connect(app_dir) as conn:
        existing = conn.execute("SELECT username, email, role FROM users WHERE id = ?", (int(user_id),)).fetchone()
        if existing is None:
            raise ValueError("User not found")
        owner = is_owner_identity(existing["username"], existing["email"])
        if owner:
            role = "admin"
            is_active = True
            access_sections = DASHBOARD_SECTIONS
        elif role == "admin":
            raise ValueError("Only Ahmed Labib can be administrator.")
        conn.execute(
            """
            UPDATE users
            SET role = ?, is_active = ?, access_sections = ?,
                approved_by = CASE WHEN ? = 1 THEN COALESCE(approved_by, ?) ELSE approved_by END,
                approved_at = CASE WHEN ? = 1 THEN COALESCE(approved_at, ?) ELSE approved_at END
            WHERE id = ?
            """,
            (
                role,
                1 if is_active else 0,
                encode_sections(access_sections, role),
                1 if is_active else 0,
                approved_by,
                1 if is_active else 0,
                utc_now(),
                int(user_id),
            ),
        )


def delete_user(app_dir: Path, user_id: int) -> tuple[bool, str]:
    with connect(app_dir) as conn:
        existing = conn.execute("SELECT username, email FROM users WHERE id = ?", (int(user_id),)).fetchone()
        if existing is None:
            return False, "User not found."
        if is_owner_identity(existing["username"], existing["email"]):
            return False, "The owner administrator account cannot be removed."
        conn.execute("DELETE FROM users WHERE id = ?", (int(user_id),))
    return True, "User removed."


def users_dataframe(app_dir: Path) -> pd.DataFrame:
    init_auth_db(app_dir)
    with connect(app_dir) as conn:
        rows = conn.execute(
            """
            SELECT id, username, email, full_name, role, is_active, must_change_password, access_sections, approved_by, approved_at, created_at, last_login_at
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
        .auth-account-bar{position:sticky;top:0;z-index:999;background:rgba(255,255,255,.96);backdrop-filter:blur(12px);border:1px solid #d9e4ef;border-radius:12px;padding:10px 12px;margin:0 0 14px;box-shadow:0 10px 28px rgba(15,23,42,.08)}
        .auth-account-name{font-weight:900;color:#0b2a4a;margin:0}
        .auth-account-meta{font-size:12px;color:#64748b;margin:2px 0 0}
        .auth-mode-note{background:#f4f7fa;border:1px solid #d9e4ef;border-radius:10px;padding:10px 12px;color:#334155;margin:8px 0 14px}
        @media(max-width:768px){.auth-account-bar{border-radius:0;margin:-1rem -1rem 14px;padding:10px 14px}.auth-account-name{font-size:14px}}
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
            <h1>Projects Intelligence Hub</h1>
            <p>Integrated Project Controls System - owner setup.</p>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    with st.form("first_admin_setup"):
        full_name = st.text_input("Full name", value="Engr. Ahmed Labib")
        username = st.text_input("Admin username", value="Ahmed_Labib", disabled=True)
        email = st.text_input("Admin email", value=OWNER_EMAIL, disabled=True)
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
            <h1>Projects Intelligence Hub</h1>
            <p>Integrated Project Controls System</p>
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
            remember_me = st.checkbox("Remember me on this device", value=True)
            submitted = st.form_submit_button("Login", type="primary", width="stretch")
        if submitted:
            ok, message, user = authenticate(app_dir, username, password)
            if ok and user:
                st.session_state["auth_user"] = user
                if remember_me:
                    st.query_params["remember"] = create_remember_token(app_dir, int(user["id"]))
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
            submitted = st.form_submit_button("Request Access", width="stretch")
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
                    is_active=False,
                )
                if ok:
                    st.success("Access request submitted. Wait for administrator approval before logging in.")
                else:
                    st.error(message)
    render_footer()


def render_admin_account_management(app_dir: Path, expanded: bool = False) -> None:
    with st.expander("Admin Control Center", expanded=expanded):
        users_df = users_dataframe(app_dir)
        if users_df.empty:
            st.info("No users found.")
        else:
            st.dataframe(users_df, width="stretch", hide_index=True)

        st.markdown("**Create approved user**")
        with st.form("admin_create_user"):
            full_name = st.text_input("Full name", key="admin_create_full_name")
            username = st.text_input("Username", key="admin_create_username")
            email = st.text_input("Email", key="admin_create_email")
            role = st.selectbox("Role", NON_ADMIN_ROLES, index=0, key="admin_create_role")
            access_sections = st.multiselect(
                "Allowed dashboard sections",
                DASHBOARD_SECTIONS,
                default=default_access_sections(role),
                key="admin_create_access_sections",
            )
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
                access_sections=access_sections,
                approved_by=OWNER_USERNAME,
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
            selected_is_owner = is_owner_identity(str(selected["username"]), str(selected["email"]))
            role_options = ROLES if selected_is_owner else NON_ADMIN_ROLES
            selected_role = str(selected["role"])
            if selected_role not in role_options:
                selected_role = role_options[0]
            new_role = st.selectbox("Role", role_options, index=role_options.index(selected_role), key="admin_edit_role", disabled=selected_is_owner)
            current_sections = decode_sections(str(selected.get("access_sections", "")), selected_role)
            access_sections = st.multiselect(
                "Allowed dashboard sections",
                DASHBOARD_SECTIONS,
                default=DASHBOARD_SECTIONS if selected_is_owner else current_sections,
                key="admin_edit_access_sections",
                disabled=selected_is_owner,
            )
            active = st.checkbox("Approved / Active", value=bool(selected["is_active"]), key="admin_edit_active", disabled=selected_is_owner)
            if st.button("Save User Changes", key="admin_save_user", width="stretch"):
                try:
                    update_user(app_dir, int(selected["id"]), new_role, active, access_sections, approved_by=OWNER_USERNAME)
                except ValueError as exc:
                    st.error(str(exc))
                    st.stop()
                st.success("User updated.")
                st.rerun()
            reset_password = st.text_input("New password", type="password", key="admin_reset_password")
            if st.button("Reset Password", key="admin_reset_password_button", width="stretch"):
                ok, message = set_password(app_dir, int(selected["id"]), reset_password, must_change_password=True)
                if ok:
                    st.success("Password reset. Share it securely and ask the user to change it.")
                else:
                    st.error(message)
            if st.button("Remove User", key="admin_remove_user", width="stretch", disabled=selected_is_owner):
                ok, message = delete_user(app_dir, int(selected["id"]))
                if ok:
                    st.success(message)
                    st.rerun()
                else:
                    st.error(message)


def current_user() -> dict[str, Any] | None:
    user = st.session_state.get("auth_user")
    return user if isinstance(user, dict) else None


def is_authenticated() -> bool:
    return current_user() is not None


def logout(app_dir: Path | None = None) -> None:
    if app_dir is not None:
        user = current_user()
        revoke_remember_token(app_dir, int(user["id"]) if user and user.get("id") else None)
    st.session_state.pop("auth_user", None)
    if "remember" in st.query_params:
        del st.query_params["remember"]


def require_authentication(app_dir: Path) -> dict[str, Any]:
    init_auth_db(app_dir)
    ensure_owner_account(app_dir)
    user = current_user()
    if not user:
        token = str(st.query_params.get("remember", "") or "").strip()
        remembered = user_from_remember_token(app_dir, token)
        if remembered:
            st.session_state["auth_user"] = remembered
            user = remembered
    if not user:
        render_login(app_dir)
        st.stop()

    render_auth_css()
    account_col, role_col, logout_col = st.columns([0.62, 0.18, 0.20])
    with account_col:
        st.markdown(
            "<div class='auth-account-bar'>"
            f"<p class='auth-account-name'>{user.get('full_name', user.get('username', 'User'))}</p>"
            f"<p class='auth-account-meta'>{user.get('email', '')}</p>"
            "</div>",
            unsafe_allow_html=True,
        )
    with role_col:
        st.markdown(
            f"<div class='auth-account-bar'><span class='role-pill'>{str(user.get('role', 'viewer')).upper()}</span></div>",
            unsafe_allow_html=True,
        )
    with logout_col:
        if st.button("Logout", width="stretch"):
            logout(app_dir)
            st.rerun()
    if user.get("role") == "viewer":
        st.markdown("<div class='auth-mode-note'>Viewer mode: read-only dashboard access.</div>", unsafe_allow_html=True)
    elif user.get("role") == "director":
        st.markdown("<div class='auth-mode-note'>Director mode: executive dashboard access.</div>", unsafe_allow_html=True)
    return user


def require_admin_authentication(app_dir: Path) -> dict[str, Any]:
    user = require_authentication(app_dir)
    if str(user.get("role", "")).lower() != "admin" or not is_owner_identity(str(user.get("username", "")), str(user.get("email", ""))):
        st.error("Admin Console access is restricted to Engr. Ahmed Labib.")
        if st.button("Sign out", width="stretch"):
            logout(app_dir)
            st.rerun()
        st.stop()
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
