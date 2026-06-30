from __future__ import annotations

from datetime import datetime, timezone
import importlib
import shutil
import site
import sqlite3
import sys
from pathlib import Path
from typing import Any

import pandas as pd

from reports.word_template_exporter import (
    find_table_by_headers,
    replace_text_preserve_format,
    sanitize_filename,
    update_keyed_table_preserve_format,
    update_table_preserve_format,
    write_cell_preserve_format,
)

try:
    Document = importlib.import_module("docx").Document
except ModuleNotFoundError:
    try:
        user_site = site.getusersitepackages()
        if user_site and user_site not in sys.path:
            sys.path.append(user_site)
        Document = importlib.import_module("docx").Document
    except ModuleNotFoundError:
        Document = None


APP_DIR = Path(__file__).resolve().parents[1]
DEFAULT_TEMPLATE_PATH = APP_DIR / "reports" / "templates" / "time_impact_analysis_report_director_pack.docx"
DEFAULT_OUTPUT_DIR = APP_DIR / "11-outputs"
DEFAULT_DB_PATH = APP_DIR / "construction_system.db"
REPORT_TYPE_TIA_DIRECTOR_PACK = "Time Impact Analysis Report | Director Pack"
CONTRACTOR_DISPLAY_NAME = "SAMCO - NATIONAL"


def _format_date(value: Any) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)) or pd.isna(value):
        return ""
    dt = pd.to_datetime(value, errors="coerce")
    if pd.isna(dt):
        return str(value)
    return dt.strftime("%d-%b-%Y")


def _text(value: Any, fallback: str = "") -> str:
    if value is None or pd.isna(value):
        return fallback
    rendered = str(value).strip()
    return rendered if rendered else fallback


def _int_text(value: Any, fallback: str = "0") -> str:
    try:
        return str(int(round(float(value))))
    except Exception:
        return fallback


def ensure_generated_outputs_table(db_path: Path) -> None:
    db_path.parent.mkdir(parents=True, exist_ok=True)
    with sqlite3.connect(db_path) as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS generated_outputs (
                output_id INTEGER PRIMARY KEY AUTOINCREMENT,
                report_type TEXT NOT NULL,
                file_name TEXT NOT NULL,
                file_path TEXT NOT NULL,
                generated_at TEXT NOT NULL,
                generated_by TEXT,
                project_name TEXT,
                data_date TEXT,
                revision TEXT,
                source_template TEXT,
                status TEXT,
                notes TEXT
            )
            """
        )
        conn.commit()


def fetch_last_generated_report(db_path: Path, report_type: str = REPORT_TYPE_TIA_DIRECTOR_PACK) -> dict[str, Any] | None:
    if not db_path.exists():
        return None
    ensure_generated_outputs_table(db_path)
    with sqlite3.connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        row = conn.execute(
            """
            SELECT output_id, report_type, file_name, file_path, generated_at, generated_by,
                   project_name, data_date, revision, source_template, status, notes
            FROM generated_outputs
            WHERE report_type = ?
            ORDER BY output_id DESC
            LIMIT 1
            """,
            (report_type,),
        ).fetchone()
    return dict(row) if row else None


def _log_generated_output(db_path: Path, payload: dict[str, Any]) -> None:
    ensure_generated_outputs_table(db_path)
    with sqlite3.connect(db_path) as conn:
        conn.execute(
            """
            INSERT INTO generated_outputs (
                report_type, file_name, file_path, generated_at, generated_by,
                project_name, data_date, revision, source_template, status, notes
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                payload.get("report_type", REPORT_TYPE_TIA_DIRECTOR_PACK),
                payload.get("file_name", ""),
                payload.get("file_path", ""),
                payload.get("generated_at", ""),
                payload.get("generated_by", ""),
                payload.get("project_name", ""),
                payload.get("data_date", ""),
                payload.get("revision", ""),
                payload.get("source_template", ""),
                payload.get("status", ""),
                payload.get("notes", ""),
            ),
        )
        conn.commit()


def build_replacements(context: dict) -> dict[str, str]:
    return {
        "[Project Name]": _text(context.get("project_name"), "Project Intelligence Hub"),
        "[Contract No.]": _text(context.get("contract_no"), "N/A"),
        "[DD-MMM-YYYY]": _format_date(context.get("data_date")),
        "[Rev. 00]": _text(context.get("revision"), "Rev. 00"),
        "[Insert project name]": _text(context.get("project_name"), "Project Intelligence Hub"),
        "[Insert employer name]": _text(context.get("employer"), "N/A"),
        "[Insert contractor name]": _text(context.get("contractor"), CONTRACTOR_DISPLAY_NAME),
        "[Insert contract clause]": _text(context.get("contract_form_clause"), "N/A"),
        "[BL Rev / Date]": _text(context.get("accepted_baseline_programme"), "N/A"),
        "[Update Rev / Data Date]": _text(context.get("impacted_update_programme"), "N/A"),
        "[5-day / 6-day / 7-day / project calendar]": _text(context.get("calendar_basis"), "N/A"),
        "[XER / P6 file name]": _text(context.get("schedule_file_name"), "N/A"),
        "[Record settings]": _text(context.get("schedule_options"), "N/A"),
        "[Longest Path / Total Float threshold]": _text(context.get("critical_path_basis"), "N/A"),
        "[Setting used]": _text(context.get("retained_logic_setting"), "N/A"),
        "[Treatment]": _text(context.get("out_of_sequence_treatment"), "N/A"),
        "[List constraints]": _text(context.get("constraints"), "None recorded"),
        "[List calendars]": _text(context.get("calendars"), "Not recorded"),
    }


def _build_project_controls_df(context: dict) -> pd.DataFrame:
    rows = [
        ("Project Name", context.get("project_name", "")),
        ("Employer / Client", context.get("employer", "")),
        ("Contractor", context.get("contractor", CONTRACTOR_DISPLAY_NAME)),
        ("Contract Form / Clause Basis", context.get("contract_form_clause", "")),
        ("Accepted Baseline Programme", context.get("accepted_baseline_programme", "")),
        ("Impacted Update Programme", context.get("impacted_update_programme", "")),
        ("Calendar Basis", context.get("calendar_basis", "")),
        ("Data Date", _format_date(context.get("data_date"))),
    ]
    return pd.DataFrame(rows, columns=["Control Item", "Input Value"])


def _build_impact_calculation_df(context: dict) -> pd.DataFrame:
    return context.get("impact_calculation_df", pd.DataFrame(columns=["Calculation Item", "Value", "Interpretation"])).copy()


def _build_eot_claim_position_df(context: dict) -> pd.DataFrame:
    if "eot_claim_position_df" in context and isinstance(context["eot_claim_position_df"], pd.DataFrame):
        return context["eot_claim_position_df"].copy()

    delay_events_df = context.get("delay_event_register_df", pd.DataFrame()).copy()
    if delay_events_df.empty:
        return pd.DataFrame(columns=["Decision Bucket", "Events", "Delay Days Treatment", "Management Action"])

    working_days_col = "Claimed Delay Duration (days)" if "Claimed Delay Duration (days)" in delay_events_df.columns else "Working Days"
    rows = []
    for decision, group in delay_events_df.groupby("Claim Decision", dropna=False):
        events = ", ".join(group["Event ID"].astype(str).tolist())
        if working_days_col in group.columns:
            total_days = pd.to_numeric(group[working_days_col], errors="coerce").fillna(0).sum()
        else:
            total_days = 0
        if str(decision).strip().lower() == "valid for tia":
            action = "Proceed to formal EOT narrative and substantiation pack."
        else:
            action = "Keep for management visibility until CPM linkage and evidence are closed."
        rows.append(
            {
                "Decision Bucket": _text(decision, "Unclassified"),
                "Events": events,
                "Delay Days Treatment": f"{_int_text(total_days)} working days",
                "Management Action": action,
            }
        )
    return pd.DataFrame(rows)


def _build_evidence_checklist_df(context: dict) -> pd.DataFrame:
    if "evidence_checklist_df" in context and isinstance(context["evidence_checklist_df"], pd.DataFrame):
        return context["evidence_checklist_df"].copy()

    evidence_df = context.get("evidence_register_df", pd.DataFrame()).copy()
    if evidence_df.empty:
        return pd.DataFrame(columns=["Checklist Item", "Status", "Required Evidence"])

    rows = []
    for _, row in evidence_df.iterrows():
        rows.append(
            {
                "Checklist Item": _text(row.get("Key Fact Proven") or row.get("Related Event"), "Evidence item"),
                "Status": _text(row.get("Strength"), "Pending"),
                "Required Evidence": _text(row.get("Missing Item"), "Already linked in current register"),
            }
        )
    return pd.DataFrame(rows)


def _update_kpi_cards(doc, context: dict) -> None:
    kpis = context.get("kpis", {})
    label_map = {
        "Total Recorded Working Delay Days": _int_text(kpis.get("total_recorded_working_delay_days"), "0"),
        "Potential EOT Days": _int_text(kpis.get("potential_eot_days"), "0"),
        "Critical Events": _int_text(kpis.get("critical_delay_events_count"), "0"),
        "Missing Evidence Items": _int_text(kpis.get("missing_evidence_items_count"), "0"),
    }
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                cell_text = cell.text.strip()
                for label, value in label_map.items():
                    if label in cell_text:
                        lines = [line.strip() for line in cell_text.splitlines() if line.strip()]
                        detail_line = ""
                        for line in lines:
                            if line != label and not line.isdigit():
                                detail_line = line
                                break
                        replacement_lines = [value, label]
                        if detail_line:
                            replacement_lines.append(detail_line)
                        write_cell_preserve_format(cell, "\n".join(replacement_lines))
                        break


def _update_executive_conclusion(doc, context: dict) -> None:
    conclusion = _text(context.get("kpis", {}).get("executive_conclusion"))
    if not conclusion:
        return
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if "Executive conclusion" in cell.text:
                    write_cell_preserve_format(cell, f"Executive conclusion\n• {conclusion}")
                    return


class TIADirectorPackGenerator:
    def __init__(self, template_path: Path, output_dir: Path, db_path: Path | None = None):
        self.template_path = Path(template_path)
        self.output_dir = Path(output_dir)
        self.db_path = Path(db_path) if db_path else DEFAULT_DB_PATH

    def validate_template(self) -> None:
        if Document is None:
            raise RuntimeError("python-docx is not installed.")
        if not self.template_path.exists():
            raise FileNotFoundError(f"Missing template: {self.template_path}")
        self.output_dir.mkdir(parents=True, exist_ok=True)

    def generate(self, context: dict) -> Path:
        self.validate_template()
        output_path = self.save_output(context)
        doc = Document(output_path)
        self.apply_replacements(doc, context)
        self.update_tables(doc, context)
        if not bool(context.get("preserve_original_charts_images", True)):
            pass
        doc.save(output_path)
        self._log_generation(context, output_path)
        return output_path

    def apply_replacements(self, doc, context: dict) -> None:
        replace_text_preserve_format(doc, build_replacements(context))
        _update_kpi_cards(doc, context)
        _update_executive_conclusion(doc, context)

    def update_tables(self, doc, context: dict) -> None:
        control_table = find_table_by_headers(doc, ["Control Item", "Input Value", "Why It Matters"])
        if control_table is not None:
            updates = {
                "Project Name": [_text(context.get("project_name"))],
                "Employer / Client": [_text(context.get("employer"))],
                "Contractor": [_text(context.get("contractor"), CONTRACTOR_DISPLAY_NAME)],
                "Contract Form / Clause Basis": [_text(context.get("contract_form_clause"))],
                "Accepted Baseline Programme": [_text(context.get("accepted_baseline_programme"))],
                "Impacted Update Programme": [_text(context.get("impacted_update_programme"))],
                "Calendar Basis": [_text(context.get("calendar_basis"))],
                "Data Date": [_format_date(context.get("data_date"))],
            }
            update_keyed_table_preserve_format(control_table, 0, updates)

        readiness_table = find_table_by_headers(doc, ["Data Input", "Priority", "Reason for TIA", "Status"])
        if readiness_table is not None:
            update_table_preserve_format(readiness_table, context.get("readiness_matrix_df", pd.DataFrame()))

        delay_events_table = find_table_by_headers(
            doc,
            ["Event", "Description", "Party", "Affected Activity", "Working Days", "Critical / LP?", "Evidence", "Decision"],
        )
        if delay_events_table is not None:
            update_table_preserve_format(
                delay_events_table,
                context.get("delay_event_register_df", pd.DataFrame()),
                column_mapping={
                    "Event": "Event ID",
                    "Description": "Event Description",
                    "Party": "Responsible Party",
                    "Affected Activity": "Affected Activity ID",
                    "Working Days": "Claimed Delay Duration (days)",
                    "Critical / LP?": "Critical / Longest Path?",
                    "Evidence": "Evidence Status",
                    "Decision": "Claim Decision",
                },
            )

        fragnet_table = find_table_by_headers(doc, ["Fragnet", "Frag Act", "Description", "Dur.", "Pred.", "Rel.", "Succ.", "Evidence"])
        if fragnet_table is not None:
            update_table_preserve_format(
                fragnet_table,
                context.get("fragnet_register_df", pd.DataFrame()),
                column_mapping={
                    "Fragnet": "Fragnet ID",
                    "Frag Act": "Fragnet Activity ID",
                    "Description": "Fragnet Activity Description",
                    "Dur.": "Duration (days)",
                    "Pred.": "Predecessor Activity",
                    "Rel.": "Relationship / Lag",
                    "Succ.": "Successor Activity",
                    "Evidence": "Record Reference",
                },
            )

        impact_calc_table = find_table_by_headers(doc, ["Calculation Item", "Value", "Interpretation"])
        if impact_calc_table is not None:
            update_table_preserve_format(impact_calc_table, _build_impact_calculation_df(context))

        activity_table = find_table_by_headers(doc, ["Activity", "Name", "Delta", "TF Before", "TF After", "Critical?", "Downstream Milestone"])
        if activity_table is not None:
            update_table_preserve_format(
                activity_table,
                context.get("activity_impact_df", pd.DataFrame()),
                column_mapping={
                    "Activity": "Activity ID",
                    "Name": "Activity Name",
                    "Delta": "Claimed Delay Duration (days)",
                    "TF Before": "Total Float Before",
                    "TF After": "Total Float After",
                    "Critical?": "Critical / LP?",
                    "Downstream Milestone": "Downstream Milestone",
                },
            )

        causation_table = find_table_by_headers(doc, ["Event", "Cause", "Critical Impact Proven?", "Concurrency / Risk Test", "Claim Treatment", "Required Action"])
        if causation_table is not None:
            update_table_preserve_format(causation_table, context.get("causation_matrix_df", pd.DataFrame()))

        p6_table = find_table_by_headers(doc, ["Control Item", "Recorded Value", "Why It Must Be Shown"])
        if p6_table is not None:
            update_table_preserve_format(p6_table, context.get("p6_controls_df", pd.DataFrame()))

        decision_table = find_table_by_headers(doc, ["Decision Bucket", "Events", "Delay Days Treatment", "Management Action"])
        if decision_table is not None:
            update_table_preserve_format(decision_table, _build_eot_claim_position_df(context))

        evidence_table = find_table_by_headers(doc, ["Checklist Item", "Status", "Required Evidence"])
        if evidence_table is not None:
            update_table_preserve_format(evidence_table, _build_evidence_checklist_df(context))

    def save_output(self, context: dict) -> Path:
        project_name = sanitize_filename(_text(context.get("project_name"), "Project"))
        data_date = sanitize_filename(_format_date(context.get("data_date")) or datetime.now().strftime("%d-%b-%Y"))
        revision = sanitize_filename(_text(context.get("revision"), "Rev_00"))
        file_name = f"TIA_Director_Pack_{project_name}_{data_date}_{revision}.docx"
        output_path = self.output_dir / file_name
        shutil.copy2(self.template_path, output_path)
        return output_path

    def _log_generation(self, context: dict, output_path: Path) -> None:
        payload = {
            "report_type": REPORT_TYPE_TIA_DIRECTOR_PACK,
            "file_name": output_path.name,
            "file_path": str(output_path),
            "generated_at": datetime.now(timezone.utc).isoformat(),
            "generated_by": _text(context.get("generated_by"), "Planning Department"),
            "project_name": _text(context.get("project_name")),
            "data_date": _format_date(context.get("data_date")),
            "revision": _text(context.get("revision")),
            "source_template": str(self.template_path),
            "status": "Generated with missing placeholders" if context.get("missing_required_fields") else "Generated",
            "notes": _text(context.get("generation_notes"), ""),
        }
        _log_generated_output(self.db_path, payload)


def generate_tia_director_pack_report(context: dict, output_dir: Path) -> Path:
    generator = TIADirectorPackGenerator(DEFAULT_TEMPLATE_PATH, output_dir, context.get("db_path", DEFAULT_DB_PATH))
    return generator.generate(context)
