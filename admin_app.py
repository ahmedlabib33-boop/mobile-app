from __future__ import annotations

from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

import auth
from src.construction_system import premium_platform as premium


APP_DIR = Path(__file__).resolve().parent
MAIN_APP_URL = "https://samco-mob-intelligence-dashboard.streamlit.app/"


st.set_page_config(
    page_title="Project Intelligence Hub Admin Console",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="collapsed",
)

premium.apply_premium_shell_css()
auth.render_auth_css()

st.markdown(
    """
    <style>
    .admin-hero{background:linear-gradient(135deg,#061a2f,#0b5260);border-radius:16px;padding:24px;margin:4px 0 18px;color:#fff;box-shadow:0 18px 42px rgba(11,42,74,.18)}
    .admin-hero h1{margin:0;color:#fff;font-size:32px}
    .admin-hero p{margin:8px 0 0;color:#dbeafe}
    .admin-warning{background:#fff7ed;border:1px solid #fed7aa;color:#7c2d12;border-radius:12px;padding:12px;margin:10px 0}
    .admin-link-card{background:#fff;border:1px solid #d9e4ef;border-radius:12px;padding:16px;box-shadow:0 10px 24px rgba(15,23,42,.07)}
    @media(max-width:768px){.admin-hero{border-radius:0;margin:0 -.75rem 14px}.admin-hero h1{font-size:24px}}
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="admin-hero">
      <h1>Project Intelligence Hub Admin Console</h1>
      <p>Owner-only access control, user approval, role assignment, and dashboard-section permissions.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

admin_user = auth.require_admin_authentication(APP_DIR)

st.markdown(
    """
    <div class="admin-warning">
      This admin console is intentionally separate from the client-facing dashboard. Deploy it as a second Streamlit Cloud app using <b>admin_app.py</b> as the entry point.
    </div>
    """,
    unsafe_allow_html=True,
)

users_df = auth.users_dataframe(APP_DIR)
if users_df.empty:
    users_df = pd.DataFrame(columns=["id", "username", "email", "full_name", "role", "is_active", "access_sections", "created_at", "last_login_at"])

pending_count = int((users_df.get("is_active", pd.Series(dtype=int)).astype(int) == 0).sum()) if not users_df.empty else 0
active_count = int((users_df.get("is_active", pd.Series(dtype=int)).astype(int) == 1).sum()) if not users_df.empty else 0
director_count = int((users_df.get("role", pd.Series(dtype=str)).astype(str) == "director").sum()) if not users_df.empty else 0
viewer_count = int((users_df.get("role", pd.Series(dtype=str)).astype(str) == "viewer").sum()) if not users_df.empty else 0

metric_cols = st.columns(5)
metric_data = [
    ("Active Users", active_count, "Approved accounts"),
    ("Pending Signups", pending_count, "Need owner action"),
    ("Directors", director_count, "Executive access"),
    ("Viewers", viewer_count, "Read-only access"),
    ("Sections", len(auth.DASHBOARD_SECTIONS), "Assignable dashboard areas"),
]
for col, (label, value, note) in zip(metric_cols, metric_data):
    col.markdown(
        f"<div class='premium-kpi'><span>{label}</span><strong>{value}</strong><small>{note}</small></div>",
        unsafe_allow_html=True,
    )

tabs = st.tabs(["User Control", "Access Matrix", "Admin Audit", "Deployment"])

with tabs[0]:
    auth.render_admin_account_management(APP_DIR, expanded=True)

with tabs[1]:
    st.markdown("<div class='premium-card'><h3>Access Matrix</h3><p>Review exactly which dashboard sections each approved user can open.</p></div>", unsafe_allow_html=True)
    matrix = users_df.copy()
    if "access_sections" in matrix.columns:
        matrix["access_sections"] = matrix["access_sections"].apply(
            lambda value: ", ".join(auth.decode_sections(str(value), "viewer"))
        )
    st.dataframe(
        matrix[["username", "email", "full_name", "role", "is_active", "access_sections", "created_at", "last_login_at"]]
        if not matrix.empty
        else matrix,
        width="stretch",
        hide_index=True,
    )
    st.download_button(
        "Download access matrix CSV",
        matrix.to_csv(index=False).encode("utf-8"),
        "admin_access_matrix.csv",
        "text/csv",
        width="stretch",
    )

with tabs[2]:
    st.markdown("<div class='premium-card'><h3>Admin Audit View</h3><p>Use this view before handover or after approving users.</p></div>", unsafe_allow_html=True)
    if not users_df.empty:
        role_df = users_df.groupby(["role", "is_active"], dropna=False).size().reset_index(name="count")
        fig = px.bar(role_df, x="role", y="count", color=role_df["is_active"].map({1: "Active", 0: "Pending"}), title="Users by Role and Approval Status")
        fig.update_layout(height=360, margin=dict(l=20, r=20, t=60, b=20), showlegend=True)
        st.plotly_chart(fig, width="stretch", config={"displaylogo": False, "responsive": True})
    st.info("Only the owner identity can hold administrator role. Any non-owner admin role is automatically downgraded by the auth initializer.")

with tabs[3]:
    st.markdown(
        f"""
        <div class="admin-link-card">
          <h3>Main Dashboard Host</h3>
          <p><a href="{MAIN_APP_URL}" target="_blank">{MAIN_APP_URL}</a></p>
          <p>Local admin host: <code>http://localhost:18756</code></p>
          <p>Local main host: <code>http://localhost:18755</code></p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.markdown(
        """
        <div class="admin-warning">
          Free Streamlit Cloud note: a second Streamlit app deployed from <code>admin_app.py</code> is a separate host. Local SQLite user data is shared automatically only when both hosts run from the same local project folder. For cloud-wide shared persistence, add a shared database service or another explicit storage backend.
        </div>
        """,
        unsafe_allow_html=True,
    )

auth.render_footer()
