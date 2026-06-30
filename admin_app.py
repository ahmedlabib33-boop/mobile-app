from __future__ import annotations

from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

import auth


APP_DIR = Path(__file__).resolve().parent
MAIN_APP_URL = "https://samco-mob-intelligence-dashboard.streamlit.app/"
ADMIN_STATUS_COLORS = {"Active": "#0F8492", "Pending": "#D1A329"}


st.set_page_config(
    page_title="Projects Intelligence Hub Admin Console",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="collapsed",
)

auth.render_auth_css()

st.markdown(
    """
    <style>
    .admin-hero{background:linear-gradient(135deg,#061a2f,#0b5260);border-radius:16px;padding:24px;margin:4px 0 18px;color:#fff;box-shadow:0 18px 42px rgba(11,42,74,.18)}
    .admin-hero h1{margin:0;color:#fff;font-size:32px}
    .admin-hero p{margin:8px 0 0;color:#dbeafe}
    .admin-warning{background:#fff7ed;border:1px solid #fed7aa;color:#7c2d12;border-radius:12px;padding:12px;margin:10px 0}
    .admin-link-card{background:#fff;border:1px solid #d9e4ef;border-radius:12px;padding:16px;box-shadow:0 10px 24px rgba(15,23,42,.07)}
    .premium-kpi{background:#fff;border:1px solid #d9e4ef;border-left:5px solid #0f8492;border-radius:12px;padding:16px;min-height:112px;box-shadow:0 10px 22px rgba(15,23,42,.07)}
    .premium-kpi span{display:block;color:#64748b;font-size:12px;font-weight:800;text-transform:uppercase}
    .premium-kpi strong{display:block;color:#0b2a4a;font-size:30px;line-height:1.1;margin-top:8px}
    .premium-kpi small{display:block;color:#64748b;margin-top:6px}
    .premium-card{background:#fff;border:1px solid #d9e4ef;border-radius:12px;padding:18px;box-shadow:0 12px 28px rgba(15,23,42,.08);margin-bottom:16px}
    @media(max-width:768px){.admin-hero{border-radius:0;margin:0 -.75rem 14px}.admin-hero h1{font-size:24px}}
    </style>
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
        fig = px.bar(
            role_df,
            x="role",
            y="count",
            color=role_df["is_active"].map({1: "Active", 0: "Pending"}),
            title="Users by Role and Approval Status",
            color_discrete_map=ADMIN_STATUS_COLORS,
        )
        fig.update_layout(
            height=360,
            autosize=True,
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="#ffffff",
            font=dict(color="#172033", family="Arial, sans-serif", size=12),
            margin=dict(l=18, r=18, t=64, b=38),
            title=dict(x=0.02, xanchor="left", font=dict(size=17, color="#0f172a")),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            hoverlabel=dict(bgcolor="#0f172a", bordercolor="#0f172a", font=dict(color="#ffffff", size=12)),
        )
        fig.update_xaxes(showline=False, ticks="outside", gridcolor="#edf3f6", zerolinecolor="#dde7ef", automargin=True)
        fig.update_yaxes(showline=False, ticks="outside", gridcolor="#edf3f6", zerolinecolor="#dde7ef", automargin=True)
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
