from __future__ import annotations

import io
import json
import math
import re
import zipfile
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib.units import mm
    from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

    REPORTLAB_READY = True
except ModuleNotFoundError:
    REPORTLAB_READY = False


SOURCE_DIRS = ("projects", "data", "exports", "reports")
SUPPORTED_SUFFIXES = {".csv", ".xlsx", ".xls", ".json"}
PLACEHOLDERS = {"", "na", "n/a", "none", "null", "tbd", "to be verified", "required", "-", "--"}
ID_HINTS = ("id", "code", "activity_id", "task_id", "wbs_id", "record_id", "event_id", "claim_id")
DATE_HINTS = ("date", "start", "finish", "baseline", "actual", "planned", "period")
STATUS_HINTS = ("status", "state", "approval")
OWNER_HINTS = ("owner", "responsible", "party", "contractor", "engineer", "discipline")
CATEGORY_HINTS = ("category", "type", "class", "risk", "claim", "delay", "package")


def apply_premium_shell_css() -> None:
    st.markdown(
        """
        <style>
        :root{--navy:#0B2A4A;--blue:#16496F;--steel:#1D5C83;--teal:#0F8492;--gold:#D1A329;--red:#B94642;--bg:#F4F7FA;--border:#D9E4EF;--text:#1F2937}
        html,body,.stApp{font-family:Inter,"Segoe UI",Arial,sans-serif;background:var(--bg);color:var(--text)}
        section[data-testid="stSidebar"], [data-testid="stSidebar"], [data-testid="collapsedControl"]{display:none!important}
        .block-container{max-width:1440px;padding:1.2rem 1.5rem 3rem}
        header[data-testid="stHeader"]{background:rgba(244,247,250,.84);backdrop-filter:blur(14px)}
        #MainMenu, footer, [data-testid="stToolbar"], [data-testid="stDecoration"], [data-testid="stStatusWidget"],
        [data-testid="manage-app-button"], [data-testid="stDeployButton"], .viewerBadge_container__1QSob,
        a[href*="github.com"][target="_blank"], a[href*="streamlit.io"]{display:none!important;visibility:hidden!important}
        .pih-topnav{position:sticky;top:0;z-index:990;background:rgba(255,255,255,.96);border:1px solid var(--border);border-radius:14px;padding:12px;margin:0 0 18px;box-shadow:0 14px 34px rgba(11,42,74,.10);backdrop-filter:blur(14px)}
        .pih-brand{display:flex;gap:10px;align-items:center;font-weight:950;color:var(--navy);letter-spacing:.01em}
        .pih-brand-dot{width:13px;height:13px;border-radius:99px;background:linear-gradient(135deg,var(--gold),var(--teal))}
        .premium-card{background:#fff;border:1px solid var(--border);border-radius:12px;padding:18px;box-shadow:0 12px 28px rgba(15,23,42,.08);margin-bottom:16px}
        .premium-card h3{margin:0 0 6px;color:var(--navy)}
        .premium-kpi{background:#fff;border:1px solid var(--border);border-left:5px solid var(--teal);border-radius:12px;padding:16px;min-height:112px;box-shadow:0 10px 22px rgba(15,23,42,.07)}
        .premium-kpi span{display:block;color:#64748b;font-size:12px;font-weight:800;text-transform:uppercase}
        .premium-kpi strong{display:block;color:var(--navy);font-size:30px;line-height:1.1;margin-top:8px}
        .premium-kpi small{display:block;color:#64748b;margin-top:6px}
        .status-badge{display:inline-flex;border-radius:999px;padding:5px 10px;font-size:12px;font-weight:900;border:1px solid var(--border);background:#fff}
        .sev-high{color:#fff;background:var(--red);border-color:var(--red)} .sev-medium{color:#4d3700;background:#fff3c4;border-color:#ead27a} .sev-low{color:#075985;background:#e0f2fe;border-color:#bae6fd}
        .quality-note{background:#eef9fb;border:1px solid #bde7ed;color:#0b5260;border-radius:10px;padding:12px;margin:12px 0}
        div[data-testid="stDataFrame"], div[data-testid="stDataEditor"]{border:1px solid var(--border);border-radius:12px;overflow:auto}
        .stButton button,.stDownloadButton button{min-height:42px;border-radius:10px;font-weight:800}
        @media(max-width:1200px){.block-container{padding:1rem 1rem 3rem}.premium-kpi strong{font-size:25px}}
        @media(max-width:768px){.block-container{padding:.75rem .75rem 5rem}.pih-topnav{top:0;border-radius:0;margin:0 -.75rem 14px}.premium-card,.premium-kpi{border-radius:10px;padding:14px}.premium-kpi strong{font-size:23px}.stTabs [data-baseweb="tab-list"]{overflow-x:auto;white-space:nowrap}.st-emotion-cache-ocqkz7{gap:.55rem}}
        @media(max-width:480px){.premium-kpi strong{font-size:21px}.premium-card{padding:12px}.stButton button,.stDownloadButton button{width:100%;min-height:46px}}
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_top_navigation(slides: list[str], key: str = "active_project_slide_name") -> str:
    if not slides:
        return ""
    if st.session_state.get(key) not in slides:
        st.session_state[key] = slides[0]
    st.markdown(
        "<div class='pih-topnav'><div class='pih-brand'><span class='pih-brand-dot'></span>Project Intelligence Hub</div></div>",
        unsafe_allow_html=True,
    )
    try:
        selected = st.pills("Navigation", slides, key=key, label_visibility="collapsed")
        if selected:
            st.session_state[key] = selected
    except Exception:
        st.selectbox("Navigation", slides, key=key, label_visibility="collapsed")
    mobile_key = f"{key}_mobile_dropdown"
    if st.session_state.get(mobile_key) not in slides:
        st.session_state[mobile_key] = st.session_state.get(key, slides[0])
    mobile_selected = st.selectbox("Mobile navigation", slides, key=mobile_key, label_visibility="collapsed")
    if mobile_selected != st.session_state.get(key):
        st.session_state[key] = mobile_selected
        return str(mobile_selected)
    return str(st.session_state.get(key, slides[0]))


def _safe_key(value: str) -> str:
    return re.sub(r"[^A-Za-z0-9_]+", "_", value).strip("_")[:90]


def _standardize_columns(columns: list[Any]) -> list[str]:
    seen: dict[str, int] = {}
    result: list[str] = []
    for column in columns:
        name = re.sub(r"[^a-z0-9]+", "_", str(column).strip().lower()).strip("_") or "unnamed"
        seen[name] = seen.get(name, 0) + 1
        result.append(name if seen[name] == 1 else f"{name}_{seen[name]}")
    return result


def _read_file(path: Path) -> dict[str, pd.DataFrame]:
    try:
        if path.suffix.lower() == ".csv":
            return {path.stem: pd.read_csv(path)}
        if path.suffix.lower() in {".xlsx", ".xls"}:
            return {f"{path.stem}__{sheet}": df for sheet, df in pd.read_excel(path, sheet_name=None).items()}
        if path.suffix.lower() == ".json":
            payload = json.loads(path.read_text(encoding="utf-8"))
            if isinstance(payload, list):
                return {path.stem: pd.DataFrame(payload)}
            if isinstance(payload, dict):
                tables = {k: pd.DataFrame(v) for k, v in payload.items() if isinstance(v, list)}
                return tables or {path.stem: pd.json_normalize(payload)}
    except Exception as exc:
        return {f"LOAD_ERROR__{path.stem}": pd.DataFrame([{"file": str(path), "error": str(exc)}])}
    return {}


@st.cache_data(show_spinner=False)
def discover_source_tables(root: str) -> dict[str, pd.DataFrame]:
    base = Path(root)
    tables: dict[str, pd.DataFrame] = {}
    roots = [base / name for name in SOURCE_DIRS if (base / name).exists()]
    if not roots:
        roots = [base]
    for folder in roots:
        for path in folder.rglob("*"):
            if path.is_file() and path.suffix.lower() in SUPPORTED_SUFFIXES:
                if any(skip in path.parts for skip in ("_RETURN_POINTS", ".venv", "__pycache__", "dist", ".pih_mobile_app_sync_state")):
                    continue
                for name, df in _read_file(path).items():
                    tables[_safe_key(f"{path.relative_to(base)}__{name}")] = df
    return tables


def clean_table(name: str, df: pd.DataFrame) -> tuple[pd.DataFrame, list[dict[str, Any]]]:
    audit: list[dict[str, Any]] = []
    cleaned = df.copy()
    before_cols = list(cleaned.columns)
    cleaned.columns = _standardize_columns(before_cols)
    if list(cleaned.columns) != before_cols:
        audit.append(_audit(name, "columns", str(before_cols), str(list(cleaned.columns)), "Standardized column names", True, "Low", "Resolved"))
    before_rows = len(cleaned)
    cleaned = cleaned.dropna(how="all")
    if len(cleaned) != before_rows:
        audit.append(_audit(name, "rows", str(before_rows), str(len(cleaned)), "Removed fully empty rows", True, "Low", "Resolved"))
    for col in cleaned.select_dtypes(include="object").columns:
        original = cleaned[col].copy()
        cleaned[col] = cleaned[col].astype(str).str.strip()
        if not original.equals(cleaned[col]):
            audit.append(_audit(name, col, "Untrimmed text", "Trimmed text", "Trimmed leading/trailing spaces", True, "Low", "Resolved"))
    dupes = int(cleaned.duplicated().sum())
    if dupes:
        cleaned = cleaned.drop_duplicates()
        audit.append(_audit(name, "rows", str(dupes), "0", "Removed exact duplicate rows", True, "Medium", "Resolved"))
    return cleaned.reset_index(drop=True), audit


def _audit(source: str, field: str, original: str, corrected: str, reason: str, auto: bool, severity: str, status: str) -> dict[str, Any]:
    return {
        "issue_id": f"AUD-{abs(hash((source, field, original, corrected))) % 100000:05d}",
        "source": source,
        "field": field,
        "original_value": original,
        "corrected_value": corrected,
        "correction_reason": reason,
        "auto_corrected": "Yes" if auto else "No",
        "requires_user_approval": "No" if auto else "Yes",
        "severity": severity,
        "status": status,
        "reviewer_note": "",
    }


def _contains_any(columns: list[str], hints: tuple[str, ...]) -> str | None:
    for hint in hints:
        for col in columns:
            if hint in col:
                return col
    return None


def validate_tables(tables: dict[str, pd.DataFrame]) -> pd.DataFrame:
    issues: list[dict[str, Any]] = []
    for source, df in tables.items():
        cols = list(df.columns)
        if df.empty:
            issues.append(_issue("High", source, "", "", "", "Load verified source data", "Source table is empty.", "Open"))
            continue
        id_col = _contains_any(cols, ID_HINTS)
        status_col = _contains_any(cols, STATUS_HINTS)
        owner_col = _contains_any(cols, OWNER_HINTS)
        if id_col and df[id_col].duplicated().any():
            issues.append(_issue("High", source, "Multiple", id_col, "Duplicate IDs", "Make identifiers unique", "Duplicate identifiers block reliable reporting.", "Open"))
        for col in cols:
            missing = df[col].isna() | df[col].astype(str).str.strip().str.lower().isin(PLACEHOLDERS)
            if int(missing.sum()):
                sev = "High" if any(h in col for h in ID_HINTS) else "Medium"
                issues.append(_issue(sev, source, str(int(missing.sum())), col, "Required / To be verified", "Complete or approve missing values", f"{int(missing.sum())} missing or placeholder values detected.", "Open"))
        if not status_col:
            issues.append(_issue("Medium", source, "", "status", "Missing", "Add status field where applicable", "No status-like column was detected.", "Open"))
        if not owner_col and len(df) > 5:
            issues.append(_issue("Low", source, "", "owner", "Missing", "Add responsible party where applicable", "No owner/responsibility-like column was detected.", "Open"))
        date_cols = [c for c in cols if any(h in c for h in DATE_HINTS)]
        parsed_dates = {c: pd.to_datetime(df[c], errors="coerce") for c in date_cols}
        for col, parsed in parsed_dates.items():
            invalid = df[col].notna() & parsed.isna()
            if int(invalid.sum()):
                issues.append(_issue("Medium", source, str(int(invalid.sum())), col, "Invalid date", "Correct date format", "Date values could not be parsed.", "Open"))
        start_col = _contains_any(cols, ("start",))
        finish_col = _contains_any(cols, ("finish", "end"))
        if start_col and finish_col:
            start = pd.to_datetime(df[start_col], errors="coerce")
            finish = pd.to_datetime(df[finish_col], errors="coerce")
            bad = start.notna() & finish.notna() & (start > finish)
            if int(bad.sum()):
                issues.append(_issue("High", source, str(int(bad.sum())), f"{start_col}/{finish_col}", "Start after finish", "Correct chronology", "Records have start dates later than finish dates.", "Open"))
        parent_col = _contains_any(cols, ("parent", "predecessor", "successor"))
        if id_col and parent_col:
            ids = set(df[id_col].dropna().astype(str))
            refs = set()
            for value in df[parent_col].dropna().astype(str):
                refs.update([x.strip() for x in re.split(r"[,;|]", value) if x.strip()])
            orphans = refs - ids
            if orphans:
                issues.append(_issue("High", source, str(len(orphans)), parent_col, ", ".join(sorted(list(orphans))[:8]), "Map missing related records", "Broken relationship/orphan references detected.", "Open"))
    if not issues:
        issues.append(_issue("Low", "All sources", "", "validation", "No blocking issue detected", "Continue review", "No unresolved validation issues were generated by the generic checker.", "Resolved"))
    return pd.DataFrame(issues)


def _issue(severity: str, source: str, record_id: str, field: str, current: str, recommendation: str, explanation: str, status: str) -> dict[str, Any]:
    return {
        "issue_id": f"VAL-{abs(hash((severity, source, record_id, field, current))) % 100000:05d}",
        "severity": severity,
        "source": source,
        "record_id": record_id,
        "field": field,
        "current_value": current,
        "recommended_correction": recommendation,
        "explanation": explanation,
        "status": status,
        "reviewer_comment": "",
    }


def readiness_score(issues: pd.DataFrame, tables: dict[str, pd.DataFrame]) -> tuple[int, dict[str, int], list[str]]:
    open_issues = issues[issues["status"].astype(str).str.lower() != "resolved"] if not issues.empty else issues
    high = int((open_issues["severity"] == "High").sum()) if not open_issues.empty else 0
    medium = int((open_issues["severity"] == "Medium").sum()) if not open_issues.empty else 0
    low = int((open_issues["severity"] == "Low").sum()) if not open_issues.empty else 0
    empty = sum(1 for df in tables.values() if df.empty)
    total_rows = sum(len(df) for df in tables.values())
    score = max(0, min(100, 100 - high * 12 - medium * 5 - low * 2 - empty * 10 + (5 if total_rows else -20)))
    categories = {
        "Data completeness": max(0, 100 - medium * 6 - empty * 20),
        "Logic correctness": max(0, 100 - high * 18 - medium * 4),
        "Relationship completeness": max(0, 100 - high * 10),
        "Chart completeness": 90 if tables else 0,
        "Export readiness": max(0, 100 - high * 15),
    }
    actions = []
    if high:
        actions.append("Resolve high-severity validation issues before client submission.")
    if medium:
        actions.append("Complete missing required values or approve them as To be verified.")
    if not tables:
        actions.append("Add verified source data files before exporting a professional package.")
    if not actions:
        actions.append("Ready for professional review with routine management qualifications.")
    return int(score), categories, actions


def score_band(score: int) -> str:
    if score >= 90:
        return "Ready for professional submission / presentation."
    if score >= 75:
        return "Ready with minor qualifications."
    if score >= 50:
        return "Needs technical review."
    return "Not ready."


def chart_pack(tables: dict[str, pd.DataFrame], issues: pd.DataFrame, score: int, categories: dict[str, int]) -> dict[str, tuple[go.Figure, pd.DataFrame, str]]:
    charts: dict[str, tuple[go.Figure, pd.DataFrame, str]] = {}
    if not issues.empty:
        sev_df = issues.groupby(["severity", "status"], dropna=False).size().reset_index(name="count")
        fig = px.bar(sev_df, x="severity", y="count", color="status", title="Validation Issues by Severity", color_discrete_sequence=["#B94642", "#D1A329", "#0F8492"])
        fig.update_layout(height=360, margin=dict(l=20, r=20, t=60, b=20))
        charts["validation_issues_by_severity"] = (fig, sev_df, "Shows unresolved and resolved issues by severity.")
    score_df = pd.DataFrame({"category": list(categories.keys()) + ["Overall"], "score": list(categories.values()) + [score]})
    fig = px.bar(score_df, x="category", y="score", title="Readiness Score by Category", color="score", color_continuous_scale=["#B94642", "#D1A329", "#0F8492"])
    fig.update_yaxes(range=[0, 100])
    fig.update_layout(height=380, margin=dict(l=20, r=20, t=60, b=80))
    charts["readiness_score_by_category"] = (fig, score_df, "Measures readiness across completeness, logic, relationships, charts, and exports.")
    for name, df in list(tables.items())[:8]:
        if df.empty:
            continue
        cols = list(df.columns)
        for hints, label in ((STATUS_HINTS, "status"), (CATEGORY_HINTS, "category"), (OWNER_HINTS, "owner")):
            col = _contains_any(cols, hints)
            if col:
                data = df[col].fillna("Required / To be verified").astype(str).replace("", "Required / To be verified").value_counts().head(12).reset_index()
                data.columns = [col, "count"]
                fig = px.bar(data, x=col, y="count", title=f"{name}: Summary by {label.title()}", color=col)
                fig.update_layout(showlegend=False, height=360, margin=dict(l=20, r=20, t=60, b=90))
                charts[f"{_safe_key(name)}_by_{label}"] = (fig, data, f"Distribution of records by detected {label} field.")
                break
        date_col = _contains_any(cols, DATE_HINTS)
        if date_col:
            series = pd.to_datetime(df[date_col], errors="coerce").dropna()
            if not series.empty:
                trend = series.dt.to_period("M").astype(str).value_counts().sort_index().reset_index()
                trend.columns = ["period", "count"]
                fig = px.line(trend, x="period", y="count", markers=True, title=f"{name}: Records by Date Period")
                fig.update_layout(height=360, margin=dict(l=20, r=20, t=60, b=70))
                charts[f"{_safe_key(name)}_date_trend"] = (fig, trend, "Monthly trend based on detected date field.")
    return charts


def dataframe_to_excel_bytes(tables: dict[str, pd.DataFrame]) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for name, df in tables.items():
            df.to_excel(writer, sheet_name=_safe_key(name)[:31] or "data", index=False)
    return output.getvalue()


def dataframe_to_json_bytes(tables: dict[str, pd.DataFrame]) -> bytes:
    return json.dumps({name: df.to_dict(orient="records") for name, df in tables.items()}, indent=2, default=str).encode("utf-8")


def chart_html_bytes(fig: go.Figure) -> bytes:
    return fig.to_html(full_html=True, include_plotlyjs="cdn").encode("utf-8")


def chart_png_bytes(fig: go.Figure) -> bytes | None:
    try:
        return fig.to_image(format="png", scale=2)
    except Exception:
        return None


def pdf_report_bytes(title: str, tables: dict[str, pd.DataFrame], issues: pd.DataFrame, score: int, actions: list[str], charts: dict[str, tuple[go.Figure, pd.DataFrame, str]], technical: bool = False) -> bytes:
    if not REPORTLAB_READY:
        return b"ReportLab is not installed. Install reportlab to enable PDF generation."
    output = io.BytesIO()
    doc = SimpleDocTemplate(output, pagesize=landscape(A4), rightMargin=12 * mm, leftMargin=12 * mm, topMargin=12 * mm, bottomMargin=12 * mm)
    styles = getSampleStyleSheet()
    story: list[Any] = []
    story.append(Paragraph(title, styles["Title"]))
    story.append(Paragraph(f"Export date: {datetime.now().strftime('%Y-%m-%d %H:%M')}", styles["Normal"]))
    story.append(Paragraph(f"Readiness score: {score}/100 - {score_band(score)}", styles["Heading2"]))
    story.append(Spacer(1, 8))
    for action in actions:
        story.append(Paragraph(f"- {action}", styles["Normal"]))
    if charts:
        story.append(PageBreak())
        story.append(Paragraph("Chart Pack", styles["Heading1"]))
        for name, (_, data, note) in list(charts.items())[:8]:
            story.append(Paragraph(name.replace("_", " ").title(), styles["Heading2"]))
            story.append(Paragraph(note, styles["Normal"]))
            story.append(_pdf_table(data.head(12)))
            story.append(Spacer(1, 8))
    story.append(PageBreak())
    story.append(Paragraph("Validation Issue Register", styles["Heading1"]))
    story.append(_pdf_table(issues.head(40)))
    if technical:
        story.append(PageBreak())
        story.append(Paragraph("Technical Data Appendix", styles["Heading1"]))
        for name, df in tables.items():
            story.append(Paragraph(name, styles["Heading2"]))
            story.append(_pdf_table(df.head(25)))
            story.append(Spacer(1, 8))
    doc.build(story, onFirstPage=_page_footer, onLaterPages=_page_footer)
    return output.getvalue()


def _page_footer(canvas: Any, doc: Any) -> None:
    canvas.saveState()
    canvas.setFont("Helvetica", 8)
    canvas.drawString(12 * mm, 8 * mm, "Project Intelligence Hub - Validation note: exports reflect edited session data and unresolved issues at generation time.")
    canvas.drawRightString(285 * mm, 8 * mm, f"Page {doc.page}")
    canvas.restoreState()


def _pdf_table(df: pd.DataFrame) -> Any:
    if df.empty:
        return Paragraph("No data available.", getSampleStyleSheet()["Normal"])
    clipped = df.copy().astype(str)
    clipped = clipped.iloc[:, :8]
    rows = [list(clipped.columns)] + clipped.head(30).values.tolist()
    table = Table(rows, repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0B2A4A")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#D9E4EF")),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 7),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ]))
    return table


def zip_package_bytes(tables: dict[str, pd.DataFrame], issues: pd.DataFrame, audit: pd.DataFrame, score: int, actions: list[str], charts: dict[str, tuple[go.Figure, pd.DataFrame, str]]) -> bytes:
    output = io.BytesIO()
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("full_report.pdf", pdf_report_bytes("Project Intelligence Hub - Full Report", tables, issues, score, actions, charts, technical=True))
        zf.writestr("executive_summary.pdf", pdf_report_bytes("Project Intelligence Hub - Executive Summary", tables, issues, score, actions, charts, technical=False))
        zf.writestr("technical_appendix.pdf", pdf_report_bytes("Project Intelligence Hub - Technical Appendix", tables, issues, score, actions, charts, technical=True))
        zf.writestr("chart_pack.pdf", pdf_report_bytes("Project Intelligence Hub - Chart Pack", {}, issues, score, actions, charts, technical=False))
        zf.writestr("data_workbook.xlsx", dataframe_to_excel_bytes(tables))
        zf.writestr("edited_data.json", dataframe_to_json_bytes(tables))
        zf.writestr("validation_report.csv", issues.to_csv(index=False).encode("utf-8"))
        zf.writestr("audit_log.csv", audit.to_csv(index=False).encode("utf-8"))
        for name, df in tables.items():
            zf.writestr(f"csv_exports/{_safe_key(name)}.csv", df.to_csv(index=False).encode("utf-8"))
        for name, (fig, data, _) in charts.items():
            zf.writestr(f"chart_html/{name}.html", chart_html_bytes(fig))
            zf.writestr(f"chart_source_csv/{name}.csv", data.to_csv(index=False).encode("utf-8"))
            png = chart_png_bytes(fig)
            if png:
                zf.writestr(f"chart_png/{name}.png", png)
        zf.writestr("README.txt", f"Project Intelligence Hub export package\nGenerated: {datetime.now().isoformat(timespec='minutes')}\nReadiness score: {score}/100\n{score_band(score)}\n")
    return output.getvalue()


def render_quality_export_center(app_dir: Path) -> None:
    st.markdown("<div class='section-header'><h3>Data Quality & Export Center</h3></div>", unsafe_allow_html=True)
    st.markdown("<div class='quality-note'>This center reviews available source files, applies only safe formatting corrections, tracks unresolved issues, and exports the current edited session data. Missing data is marked as Required / To be verified.</div>", unsafe_allow_html=True)
    raw_tables = discover_source_tables(str(app_dir))
    cleaned: dict[str, pd.DataFrame] = {}
    audit_rows: list[dict[str, Any]] = []
    for name, df in raw_tables.items():
        cdf, audits = clean_table(name, df)
        cleaned[name] = cdf
        audit_rows.extend(audits)
    if "premium_edited_tables" not in st.session_state:
        st.session_state["premium_edited_tables"] = {name: df.copy() for name, df in cleaned.items()}
    tables: dict[str, pd.DataFrame] = st.session_state["premium_edited_tables"]
    if set(tables.keys()) != set(cleaned.keys()):
        for name, df in cleaned.items():
            tables.setdefault(name, df.copy())
        for name in list(tables.keys()):
            if name not in cleaned:
                tables.pop(name, None)
    audit_df = pd.DataFrame(audit_rows)
    issues = validate_tables(tables)
    score, categories, actions = readiness_score(issues, tables)
    charts = chart_pack(tables, issues, score, categories)
    _render_quality_kpis(score, issues, audit_df, tables)
    tabs = st.tabs(["Executive Readiness", "Editable Data", "Charts", "Validation Register", "Export Center", "Audit Log"])
    with tabs[0]:
        st.markdown(f"<div class='premium-card'><h3>Readiness Score: {score}/100</h3><p>{score_band(score)}</p></div>", unsafe_allow_html=True)
        for action in actions:
            st.warning(action)
        score_df = pd.DataFrame({"Category": list(categories.keys()), "Score": list(categories.values())})
        st.dataframe(score_df, width="stretch", hide_index=True)
    with tabs[1]:
        _render_editable_tables(tables, cleaned)
    with tabs[2]:
        _render_charts(charts)
    with tabs[3]:
        edited_issues = st.data_editor(issues, width="stretch", hide_index=True, num_rows="dynamic", key="premium_validation_editor")
        st.download_button("Download validation CSV", edited_issues.to_csv(index=False).encode("utf-8"), "validation_report.csv", "text/csv", width="stretch")
    with tabs[4]:
        _render_export_center(tables, issues, audit_df, score, actions, charts)
    with tabs[5]:
        if audit_df.empty:
            st.success("No automatic correction was required beyond the current loaded structure.")
        else:
            st.dataframe(audit_df, width="stretch", hide_index=True)
            st.download_button("Download audit log CSV", audit_df.to_csv(index=False).encode("utf-8"), "source_data_audit_log.csv", "text/csv", width="stretch")


def _render_quality_kpis(score: int, issues: pd.DataFrame, audit_df: pd.DataFrame, tables: dict[str, pd.DataFrame]) -> None:
    open_issues = issues[issues["status"].astype(str).str.lower() != "resolved"] if not issues.empty else issues
    high = int((open_issues["severity"] == "High").sum()) if not open_issues.empty else 0
    medium = int((open_issues["severity"] == "Medium").sum()) if not open_issues.empty else 0
    rows = sum(len(df) for df in tables.values())
    values = [
        ("Readiness", f"{score}/100", score_band(score)),
        ("Open High", str(high), "Must be cleared before final submission"),
        ("Open Medium", str(medium), "Needs review or approval"),
        ("Source Tables", str(len(tables)), f"{rows:,} loaded records"),
        ("Auto Corrections", str(len(audit_df)), "Logged formatting/data hygiene actions"),
    ]
    cols = st.columns(len(values))
    for col, (label, value, note) in zip(cols, values):
        col.markdown(f"<div class='premium-kpi'><span>{label}</span><strong>{value}</strong><small>{note}</small></div>", unsafe_allow_html=True)


def _render_editable_tables(tables: dict[str, pd.DataFrame], original: dict[str, pd.DataFrame]) -> None:
    uploaded = st.file_uploader("Load edited JSON", type=["json"], key="premium_load_json")
    if uploaded is not None:
        try:
            payload = json.loads(uploaded.read().decode("utf-8"))
            st.session_state["premium_edited_tables"] = {k: pd.DataFrame(v) for k, v in payload.items() if isinstance(v, list)}
            st.success("Edited JSON loaded into the current session.")
            st.rerun()
        except Exception as exc:
            st.error(f"JSON import failed: {exc}")
    reset_col, json_col = st.columns(2)
    if reset_col.button("Reset to original loaded data", width="stretch"):
        st.session_state["premium_edited_tables"] = {name: df.copy() for name, df in original.items()}
        st.rerun()
    json_col.download_button("Download current edited JSON", dataframe_to_json_bytes(tables), "edited_data.json", "application/json", width="stretch")
    for name, df in tables.items():
        with st.expander(f"{name} ({len(df):,} rows)", expanded=False):
            search = st.text_input("Search table", key=f"search_{_safe_key(name)}")
            view = df
            if search:
                mask = df.astype(str).apply(lambda row: row.str.contains(search, case=False, na=False).any(), axis=1)
                view = df[mask]
            edited = st.data_editor(view, width="stretch", hide_index=True, num_rows="dynamic", key=f"editor_{_safe_key(name)}")
            if not search:
                tables[name] = edited
            st.download_button("CSV", view.to_csv(index=False).encode("utf-8"), f"{_safe_key(name)}.csv", "text/csv", key=f"csv_{_safe_key(name)}")
            st.download_button("Excel", dataframe_to_excel_bytes({name: view}), f"{_safe_key(name)}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"xlsx_{_safe_key(name)}")


def _render_charts(charts: dict[str, tuple[go.Figure, pd.DataFrame, str]]) -> None:
    if not charts:
        st.info("No chartable fields were detected in the current source data.")
        return
    for name, (fig, data, note) in charts.items():
        with st.container():
            st.markdown(f"<div class='premium-card'><h3>{name.replace('_', ' ').title()}</h3><p>{note}</p></div>", unsafe_allow_html=True)
            st.plotly_chart(fig, width="stretch", config={"displaylogo": False, "responsive": True})
            c1, c2, c3 = st.columns(3)
            c1.download_button("Download HTML", chart_html_bytes(fig), f"{name}.html", "text/html", key=f"html_{name}", width="stretch")
            c2.download_button("Download source CSV", data.to_csv(index=False).encode("utf-8"), f"{name}.csv", "text/csv", key=f"chart_csv_{name}", width="stretch")
            png = chart_png_bytes(fig)
            c3.download_button("Download PNG" if png else "PNG unavailable", png or b"Kaleido is not installed.", f"{name}.png" if png else f"{name}_png_unavailable.txt", "image/png" if png else "text/plain", key=f"png_{name}", width="stretch")


def _render_export_center(tables: dict[str, pd.DataFrame], issues: pd.DataFrame, audit_df: pd.DataFrame, score: int, actions: list[str], charts: dict[str, tuple[go.Figure, pd.DataFrame, str]]) -> None:
    st.markdown("<div class='premium-card'><h3>Professional Export Center</h3><p>All downloads reflect the current edited data, validation register, and chart pack.</p></div>", unsafe_allow_html=True)
    include_technical = st.checkbox("Include technical appendix and raw data", value=True)
    c1, c2, c3 = st.columns(3)
    c1.download_button("Full PDF Report", pdf_report_bytes("Project Intelligence Hub - Full Report", tables, issues, score, actions, charts, technical=include_technical), "full_report.pdf", "application/pdf", width="stretch")
    c2.download_button("Excel Workbook", dataframe_to_excel_bytes(tables), "data_workbook.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", width="stretch")
    c3.download_button("JSON Data Model", dataframe_to_json_bytes(tables), "edited_data.json", "application/json", width="stretch")
    z1, z2 = st.columns(2)
    z1.download_button("Complete ZIP Package", zip_package_bytes(tables, issues, audit_df, score, actions, charts), "project_intelligence_hub_export_package.zip", "application/zip", width="stretch")
    z2.download_button("Validation CSV", issues.to_csv(index=False).encode("utf-8"), "validation_report.csv", "text/csv", width="stretch")
