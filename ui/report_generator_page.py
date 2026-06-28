from __future__ import annotations

from pathlib import Path
from typing import Any

import pandas as pd


DIRECTOR_PACK_REQUIRED_FIELDS = [
    "project_name",
    "data_date",
    "revision",
]


def compute_required_fields_completion(inputs: dict[str, Any]) -> tuple[float, list[str]]:
    missing = [field for field in DIRECTOR_PACK_REQUIRED_FIELDS if not str(inputs.get(field, "")).strip()]
    completion = 100.0 * (len(DIRECTOR_PACK_REQUIRED_FIELDS) - len(missing)) / max(len(DIRECTOR_PACK_REQUIRED_FIELDS), 1)
    return completion, missing


def build_replacement_preview_df(inputs: dict[str, Any]) -> pd.DataFrame:
    preview_rows = [
        ("Project Name", inputs.get("project_name", "")),
        ("Contract No.", inputs.get("contract_no", "")),
        ("Data Date", inputs.get("data_date", "")),
        ("Revision", inputs.get("revision", "")),
        ("Employer / Client", inputs.get("employer", "")),
        ("Contractor", inputs.get("contractor", "")),
        ("Contract Clause / Form", inputs.get("contract_form_clause", "")),
        ("Accepted Baseline Programme", inputs.get("accepted_baseline_programme", "")),
        ("Impacted Update Programme", inputs.get("impacted_update_programme", "")),
        ("Calendar Basis", inputs.get("calendar_basis", "")),
        ("Schedule File Name", inputs.get("schedule_file_name", "")),
        ("Schedule Options", inputs.get("schedule_options", "")),
        ("Critical / Longest Path Basis", inputs.get("critical_path_basis", "")),
        ("Retained Logic / Progress Override", inputs.get("retained_logic_setting", "")),
        ("Out-of-sequence Progress Treatment", inputs.get("out_of_sequence_treatment", "")),
        ("Constraints", inputs.get("constraints", "")),
        ("Calendars", inputs.get("calendars", "")),
        ("Open Ends", inputs.get("open_ends", "")),
        ("Negative Float", inputs.get("negative_float", "")),
    ]
    return pd.DataFrame(preview_rows, columns=["Replacement Field", "Value"])


def build_data_source_status_df(dataframes: dict[str, pd.DataFrame]) -> pd.DataFrame:
    rows = []
    for label, df in dataframes.items():
        rows.append(
            {
                "Data Source": label,
                "Rows": int(len(df)) if isinstance(df, pd.DataFrame) else 0,
                "Status": "Ready" if isinstance(df, pd.DataFrame) and not df.empty else "Missing / Empty",
            }
        )
    return pd.DataFrame(rows)


def resolve_existing_output_path(raw_path: str | None) -> Path | None:
    if not raw_path:
        return None
    path = Path(raw_path)
    return path if path.exists() else None
