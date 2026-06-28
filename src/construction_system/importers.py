from __future__ import annotations

import csv
import re
from pathlib import Path

from .database import get_connection, init_db


# ── Column aliases: maps CSV header → DB column name ─────────────────────────
COLUMN_ALIASES: dict[str, dict[str, str]] = {
    "contracts": {
        "package":           "contract_no",
        "contractor":        "contractor_name",
        "paid_to_date":      "paid_amount",
        "retention_percent": "retention_amount",
        "contract_value":    "original_value",   # fallback if original_value absent
    },
    "delay_events": {
        "delay_title":          "event_title",
        "start_date":           "event_date",
        "responsibility":       "responsible_party",
        "estimated_delay_days": "delay_days",
        "approved_eot_days":    "eot_days",
        "primary_event_id":     "delay_event_id",
        "activity_id":          "activity_id",
        "start":                "event_date",
        "overlap_start":        "event_date",
        "delayed_duration_after_overlap": "delay_days",
        "delayed_duration":     "delay_days",
        # cause_category / notice_ref are ignored (no matching DB column)
    },
    "payments": {
        "payment":        "payment_id",
        "invoice_no":     "invoice_no",
        "invoice_date":   "payment_date",
        "paid_amount":    "amount",
        "certified_amount": "amount",
        "payment_status": "status",
        "project":        "project_id",
        "contract":       "contract_id",
    },
    "milestones": {
        "activity_name": "milestone_name",
        "planned_date":  "target_date",
        "forecast_date": "actual_date",          # best available proxy
        # milestone_contractual_type ignored
    },
    "projects": {
        "name": "project_name",
    },
    "cost_items": {
        "cost_code":     "cost_item_id",   # activity_id stored as cost_item_id for lookup
        "activity_name": "cost_category",  # activity description stored in cost_category
    },
    "wbs": {
        "WBS Code": "wbs_id",
        "WBS Name": "wbs_name",
    },
}

# DB columns that must be numeric (strip currency symbols, commas, %)
NUMERIC_COLUMNS = {
    "original_value", "approved_variations", "pending_variations",
    "certified_amount", "paid_amount", "retention_amount",
    "contract_value", "budget", "actual_cost",
    "planned_weight", "planned_progress", "actual_progress",
    "total_float_days", "delay_days", "eot_days",
    "probability", "time_impact_days", "cost_impact",
    "claim_value", "amount",
}

# DB columns that must be 0/1 integers
BOOLEAN_COLUMNS = {
    "is_critical",
    "critical_impact",
}

# Known DB columns per table (used to drop unknown CSV columns)
TABLE_COLUMNS: dict[str, set[str]] = {
    "projects":         {"project_id","project_name","project_code","client_name","contractor",
                         "planned_start","planned_finish","forecast_finish","contract_value",
                         "currency","status"},
    "wbs":              {"wbs_id","project_id","wbs_code","wbs_name","parent_wbs_id"},
    "activities": {
        "project_id",
            "wbs_id",
            "activity_id",
            "activity_name",
            "planned_start",
            "planned_finish",
            "actual_start",
            "actual_finish",
            "forecast_start",
            "forecast_finish",
            "planned_weight",
            "planned_progress",
            "actual_progress",
            "total_float_days",
            "is_critical",
            "responsible_party"
            },
    "progress_updates": {"update_id","project_id","activity_id","update_date","planned_progress",
                         "actual_progress","planned_value","earned_value","actual_cost","notes"},
    "cost_items":       {"cost_item_id","project_id","wbs_id","cost_category","budget_cost",
                         "actual_cost","forecast_cost"},
    "contracts":        {"contract_id","project_id","contract_no","contractor_name","contract_type",
                         "original_value","approved_variations","pending_variations",
                         "certified_amount","paid_amount","retention_amount",
                         "start_date","finish_date","status"},
    "change_orders":    {"change_order_id","project_id","contract_id","title","issue_date",
                         "cost_impact","time_impact_days","status"},
    "claims":           {"claim_id","project_id","claim_title","claim_type","claim_date",
                         "claim_value","status","responsible_party","eot_days"},
    "payments":         {"payment_id","project_id","contract_id","invoice_no","payment_date",
                         "amount","status"},
    "delay_events":     {"delay_event_id","project_id","activity_id","event_title","event_date",
                         "delay_days","responsible_party","critical_impact","eot_days","status"},
    "risks":            {"risk_id","project_id","risk_code","risk_title","category",
                         "probability","time_impact_days","cost_impact","response_strategy",
                         "owner","mitigation_status","due_date"},
    "milestones":       {"milestone_id","project_id","milestone_name","target_date",
                         "actual_date","status"},
}


def _derive_parent_wbs_id(wbs_id: str | None) -> str | None:
    if not wbs_id:
        return None
    value = str(wbs_id).strip()
    if not value or "." not in value or value.count(".") <= 1:
        return None
    return value.rsplit(".", 1)[0]


def _normalize_wbs_row(row: dict) -> dict:
    """Support Primavera-style WBS exports that use WBS Code / WBS Name and omit hierarchy fields."""
    wbs_id = row.get("wbs_id") or row.get("wbs_code")
    wbs_name = row.get("wbs_name")
    if wbs_id:
        row["wbs_id"] = wbs_id
    if not row.get("wbs_code") and wbs_id:
        row["wbs_code"] = wbs_id
    if wbs_name:
        row["wbs_name"] = wbs_name
    if not row.get("parent_wbs_id"):
        row["parent_wbs_id"] = _derive_parent_wbs_id(wbs_id)
    return row

TABLE_FILES = {
    "projects.csv":         "projects",
    "wbs.csv":              "wbs",
    "activities.csv":       "activities",
    "progress_updates.csv": "progress_updates",
    "cost_items.csv":       "cost_items",
    "contracts.csv":        "contracts",
    "change_orders.csv":    "change_orders",
    "claims.csv":           "claims",
    "payments.csv":         "payments",
    "delay_events.csv":     "delay_events",
    "risks.csv":            "risks",
    "milestones.csv":       "milestones",
}


# ── Value cleaners ─────────────────────────────────────────────────────────────

def _clean_str(value) -> str | None:
    if value is None:
        return None
    s = str(value).strip()
    return s if s else None


def _clean_numeric(value) -> float | None:
    """Strip currency symbols, commas, %, spaces; return float or None."""
    if value is None:
        return None
    s = str(value).strip()
    if not s:
        return None
    # Remove currency symbols, commas, spaces, trailing %
    s = re.sub(r"[$,\s]", "", s)
    s = s.rstrip("%")
    try:
        return float(s)
    except ValueError:
        return None


def _clean_boolean(value) -> int | None:
    """Convert Yes/No/True/False/1/0 → 1/0."""
    if value is None:
        return 0
    s = str(value).strip().lower()
    if s in ("yes", "true", "1", "y"):
        return 1
    if s in ("no", "false", "0", "n", ""):
        return 0
    return 0


def _clean_date(value) -> str | None:
    """Normalise dates like '11-Sep-25 A' → '11-Sep-25', keep ISO dates as-is."""
    if value is None:
        return None
    s = str(value).strip()
    if not s:
        return None
    # Strip trailing ' A' (Actual marker from Primavera exports)
    s = re.sub(r"\s+A$", "", s).strip()
    return s if s else None


DATE_COLUMNS = {
    "planned_start", "planned_finish", "actual_start", "actual_finish",
    "forecast_start", "forecast_finish", "event_date", "start_date",
    "finish_date", "issue_date", "claim_date", "payment_date",
    "update_date", "target_date", "actual_date", "due_date",
}


def _apply_aliases(row: dict, table_name: str) -> dict:
    """Rename CSV columns to DB column names using COLUMN_ALIASES."""
    aliases = COLUMN_ALIASES.get(table_name, {})
    normalized_aliases = {
        re.sub(r"[^a-z0-9]+", "_", key.strip().lower()).strip("_"): value
        for key, value in aliases.items()
    }
    result = {}
    for k, v in row.items():
        normalized_key = re.sub(r"[^a-z0-9]+", "_", str(k).strip().lower()).strip("_")
        mapped = aliases.get(k, normalized_aliases.get(normalized_key, normalized_key))
        # Don't overwrite a key that already exists with the canonical name
        if mapped not in result:
            result[mapped] = v
    return result


def _normalize_modified_template_row(row: dict, table_name: str) -> dict:
    if table_name == "delay_events":
        if not row.get("event_title"):
            row["event_title"] = row.get("delay_event_id") or row.get("activity_name") or "Delay Event"
        if not row.get("delay_event_id"):
            row["delay_event_id"] = row.get("event_title") or row.get("activity_id")
        if row.get("delay_event_id") and row.get("activity_id"):
            event_id = str(row.get("delay_event_id")).strip()
            activity_id = str(row.get("activity_id")).strip()
            if event_id and activity_id and activity_id.upper() not in event_id.upper():
                row["delay_event_id"] = f"{event_id}-{activity_id}"
        if not row.get("responsible_party"):
            row["responsible_party"] = "Employer / Client"
        if not row.get("status"):
            row["status"] = "Open"
        if not row.get("eot_days"):
            row["eot_days"] = 0
        if not row.get("critical_impact"):
            critical_value = str(row.get("current_critical_path") or row.get("bl_critical_path") or "").strip().lower()
            row["critical_impact"] = 1 if critical_value in {"yes", "true", "1", "y"} else 0
    elif table_name == "payments":
        if not row.get("payment_id") and row.get("invoice_no"):
            row["payment_id"] = row.get("invoice_no")
        if not row.get("status") and row.get("payment_status"):
            row["status"] = row.get("payment_status")
    return row


def _clean_row(row: dict, table_name: str) -> dict:
    """Apply type coercions and drop columns not in the DB schema."""
    row = _normalize_modified_template_row(row, table_name)
    if table_name == "wbs":
        row = _normalize_wbs_row(row)
    valid_cols = TABLE_COLUMNS.get(table_name, set())
    cleaned = {}
    for col, val in row.items():
        if valid_cols and col not in valid_cols:
            continue  # drop unknown columns
        if col in BOOLEAN_COLUMNS:
            cleaned[col] = _clean_boolean(val)
        elif col in NUMERIC_COLUMNS:
            cleaned[col] = _clean_numeric(val)
        elif col in DATE_COLUMNS:
            cleaned[col] = _clean_date(val)
        else:
            cleaned[col] = _clean_str(val)
    return cleaned


# ── CSV reader ─────────────────────────────────────────────────────────────────

def _read_csv(file_path: Path, table_name: str) -> list[dict]:
    with file_path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        rows = []
        for raw in reader:
            row = _apply_aliases(dict(raw), table_name)
            row = _clean_row(row, table_name)
            if any(v is not None for v in row.values()):
                rows.append(row)
    return rows


# ── DB insert ──────────────────────────────────────────────────────────────────

def _insert_rows(conn, table_name: str, rows: list[dict]) -> int:
    if not rows:
        return 0

    # Collect all columns present across all rows
    all_cols: list[str] = list(dict.fromkeys(col for row in rows for col in row))
    placeholders = ", ".join(["?"] * len(all_cols))
    column_sql = ", ".join(all_cols)
    sql = f"INSERT OR IGNORE INTO {table_name} ({column_sql}) VALUES ({placeholders})"

    values = [[row.get(col) for col in all_cols] for row in rows]
    conn.executemany(sql, values)
    return len(rows)


# ── Public API ─────────────────────────────────────────────────────────────────

def import_csv_folder(folder_path, db_path, reset: bool = False) -> dict[str, int]:
    folder = Path(folder_path)
    if not folder.exists():
        raise FileNotFoundError(f"Folder not found: {folder}")

    init_db(db_path, reset=reset)

    results: dict[str, int] = {}

    with get_connection(db_path) as conn:
        for file_name, table_name in TABLE_FILES.items():
            file_path = folder / file_name
            if not file_path.exists():
                continue

            rows = _read_csv(file_path, table_name)
            count = _insert_rows(conn, table_name, rows)
            results[table_name] = count

    return results
