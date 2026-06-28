from __future__ import annotations

from html import escape
from pathlib import Path

from .analytics import (
    get_contract_summary,
    get_delay_analysis,
    get_project_control_summary,
    get_risk_analysis,
)
from .database import DEFAULT_DB_PATH, ROOT


def _format_value(value):
    if value is None:
        return "-"
    if isinstance(value, float):
        return f"{value:,.2f}"
    return str(value)


def _table_from_rows(rows: list[dict], columns: list[str]) -> str:
    head = "".join(f"<th>{escape(col)}</th>" for col in columns)
    body_rows = []
    for row in rows:
        cells = "".join(f"<td>{escape(_format_value(row.get(col)))}</td>" for col in columns)
        body_rows.append(f"<tr>{cells}</tr>")
    body = "".join(body_rows) if body_rows else "<tr><td colspan='99'>No data</td></tr>"
    return f"<table><thead><tr>{head}</tr></thead><tbody>{body}</tbody></table>"


def generate_html_report(db_path=DEFAULT_DB_PATH, output_path: str | Path | None = None) -> Path:
    output_path = Path(output_path) if output_path else ROOT / "reports" / "project_control_report.html"
    output_path.parent.mkdir(parents=True, exist_ok=True)

    project_data = get_project_control_summary(db_path)
    contracts = get_contract_summary(db_path)
    delays = get_delay_analysis(db_path)
    risks = get_risk_analysis(db_path)

    project = project_data["project"] or {}
    kpis = project_data["kpis"]

    html = f"""
    <html>
    <head>
        <meta charset="utf-8">
        <title>Project Control Report</title>
        <style>
            body {{ font-family: Arial, sans-serif; margin: 24px; background: #f7f7f7; color: #222; }}
            .card {{ background: #fff; padding: 16px; margin-bottom: 18px; border-radius: 10px; box-shadow: 0 1px 6px rgba(0,0,0,.08); }}
            .grid {{ display: grid; grid-template-columns: repeat(4, 1fr); gap: 12px; }}
            .kpi {{ background: #f0f4f8; padding: 12px; border-radius: 8px; }}
            .kpi .label {{ font-size: 12px; color: #666; }}
            .kpi .value {{ font-size: 20px; font-weight: bold; margin-top: 6px; }}
            table {{ width: 100%; border-collapse: collapse; margin-top: 10px; }}
            th, td {{ border: 1px solid #ddd; padding: 8px; font-size: 13px; text-align: left; }}
            th {{ background: #e9eef5; }}
            h1, h2 {{ margin: 0 0 12px 0; }}
        </style>
    </head>
    <body>
        <div class="card">
            <h1>Construction Project Control and Contract Management Report</h1>
            <p><strong>Project:</strong> {escape(str(project.get("project_name", "N/A")))}</p>
            <p><strong>Code:</strong> {escape(str(project.get("project_code", "N/A")))}</p>
        </div>

        <div class="card">
            <h2>Core KPIs</h2>
            <div class="grid">
                <div class="kpi"><div class="label">BAC</div><div class="value">{_format_value(kpis.get("bac"))}</div></div>
                <div class="kpi"><div class="label">AC</div><div class="value">{_format_value(kpis.get("ac"))}</div></div>
                <div class="kpi"><div class="label">EV</div><div class="value">{_format_value(kpis.get("ev"))}</div></div>
                <div class="kpi"><div class="label">PV</div><div class="value">{_format_value(kpis.get("pv"))}</div></div>
                <div class="kpi"><div class="label">SPI</div><div class="value">{_format_value(kpis.get("spi"))}</div></div>
                <div class="kpi"><div class="label">CPI</div><div class="value">{_format_value(kpis.get("cpi"))}</div></div>
                <div class="kpi"><div class="label">EAC</div><div class="value">{_format_value(kpis.get("eac"))}</div></div>
                <div class="kpi"><div class="label">Critical Activities</div><div class="value">{_format_value(kpis.get("critical_activities"))}</div></div>
            </div>
        </div>

        <div class="card">
            <h2>Contract Summary</h2>
            {_table_from_rows(contracts, ["contract_no", "contractor_name", "original_value", "approved_variations", "pending_variations", "certified_amount", "paid_amount", "unpaid_certified_balance", "status"])}
        </div>

        <div class="card">
            <h2>Delay Analysis</h2>
            {_table_from_rows(delays, ["event_title", "activity_code", "activity_name", "delay_days", "responsible_party", "critical_impact", "eot_days", "eot_potential", "concurrent_delay_warning", "status"])}
        </div>

        <div class="card">
            <h2>Risk Analysis</h2>
            {_table_from_rows(risks, ["risk_code", "risk_title", "category", "probability", "time_impact_days", "cost_impact", "weighted_time_exposure", "severity_band", "owner", "mitigation_status"])}
        </div>
    </body>
    </html>
    """

    output_path.write_text(html, encoding="utf-8")
    return output_path