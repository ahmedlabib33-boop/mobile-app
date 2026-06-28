from __future__ import annotations

from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
from typing import Any
import re

import pandas as pd


CANONICAL_P6_FIELDS = [
    "Activity ID",
    "Activity Name",
    "WBS",
    "Activity Code",
    "Baseline Start",
    "Baseline Finish",
    "Start",
    "Finish",
    "Actual Start",
    "Actual Finish",
    "Remaining Duration",
    "Total Float",
    "Free Float",
    "Critical",
    "Longest Path",
    "Predecessors",
    "Successors",
    "Remaining Units",
    "Budgeted Units",
    "Physical % Complete",
]

CANONICAL_STEEL_FIELDS = [
    "Date",
    "Steel Type / Diameter",
    "Delivered Qty",
    "Consumed Qty",
    "Balance Qty",
    "Delivery Ref",
    "Remarks",
    "Building / Zone Allocation",
    "Activity ID consuming the steel",
    "Remaining Required Qty",
    "Balance After Allocation",
]

CANONICAL_REQ_FIELDS = [
    "Activity ID",
    "Activity Name",
    "Building",
    "Steel Type / Diameter",
    "Required Steel Qty",
    "Required Date",
    "Priority",
    "Remarks",
]

CANONICAL_REL_FIELDS = [
    "Activity ID",
    "Predecessor ID",
    "Relationship Type",
    "Lag",
    "Successor ID",
]

CANONICAL_CLIENT_ACTIVITY_SUPPLY_FIELDS = [
    "Resource ID",
    "Activity ID",
    "Activity Name",
    "Original Duration",
    "Actual Duration",
    "Remaining Duration",
    "Deviated Duration",
    "Planned Start",
    "Planned Finish",
    "BL Start",
    "BL Finish",
    "Start",
    "Finish",
    "Actual Finish",
    "Budgeted Units",
    "Actual Units",
    "Date of receiving units supplied by client",
    "Available units at site received by client",
    "Available units supplied by client vs Actual units of activities",
    "Remaining Units",
    "Units % Complete",
    "Budgeted Units / Time",
    "Drive Activity Dates",
    "Reason for Delay",
    "Responsibility",
]

P6_ALIASES = {
    "Activity ID": ["Activity ID", "activity_id", "activity id"],
    "Activity Name": ["Activity Name", "activity_name", "activity name"],
    "WBS": ["WBS", "wbs_id", "wbs", "WBS ID"],
    "Activity Code": ["Activity Code", "activity_code", "resource id"],
    "Baseline Start": ["Baseline Start", "bl_project_start", "planned start", "Planned Start"],
    "Baseline Finish": ["Baseline Finish", "bl_project_finish", "planned finish", "Planned Finish"],
    "Start": ["Start", "start"],
    "Finish": ["Finish", "finish"],
    "Actual Start": ["Actual Start", "actual_start"],
    "Actual Finish": ["Actual Finish", "actual_finish"],
    "Remaining Duration": ["Remaining Duration", "remaining_duration", "original_duration", "Original Duration"],
    "Total Float": ["Total Float", "total_float", "total_float_days", "total float"],
    "Free Float": ["Free Float", "free_float"],
    "Critical": ["Critical", "critical", "is_critical"],
    "Longest Path": ["Longest Path", "longest_path", "Drive Activity Dates"],
    "Predecessors": ["Predecessors", "predecessors", "predecessor_details"],
    "Successors": ["Successors", "successors", "successor_details"],
    "Remaining Units": ["Remaining Units", "remaining_units"],
    "Budgeted Units": ["Budgeted Units", "budgeted_units"],
    "Physical % Complete": ["Physical % Complete", "activity_%_complete", "Units % Complete", "actual_progress"],
}

STEEL_ALIASES = {
    "Date": ["Date", "Date of delivery", "delivery date", "date", "Date of receiving units supplied by client"],
    "Steel Type / Diameter": ["Steel Type / Diameter", "steel_type", "diameter", "Steel Type", "steel type / diameter"],
    "Delivered Qty": ["Delivered Qty", "Total Quantity", "delivered_qty", "qty delivered", "Available units at site received by client"],
    "Consumed Qty": ["Consumed Qty", "consumed_qty"],
    "Balance Qty": ["Balance Qty", "balance_qty"],
    "Delivery Ref": ["Delivery Ref", "Reference", "Coding No.", "delivery_ref"],
    "Remarks": ["Remarks", "remarks"],
    "Building / Zone Allocation": ["Building / Zone Allocation", "building", "zone", "allocation"],
    "Activity ID consuming the steel": ["Activity ID consuming the steel", "Activity ID", "activity_id"],
    "Remaining Required Qty": ["Remaining Required Qty", "remaining required qty"],
    "Balance After Allocation": ["Balance After Allocation", "balance after allocation"],
}

REQ_ALIASES = {
    "Activity ID": ["Activity ID", "activity_id"],
    "Activity Name": ["Activity Name", "activity_name"],
    "Building": ["Building", "building"],
    "Steel Type / Diameter": ["Steel Type / Diameter", "Resource ID", "steel_type", "diameter"],
    "Required Steel Qty": ["Required Steel Qty", "Remaining Units", "remaining_units", "Budgeted Units", "budgeted_units"],
    "Required Date": ["Required Date", "Start", "Planned Start", "start", "planned_start"],
    "Priority": ["Priority", "priority"],
    "Remarks": ["Remarks", "remarks"],
}

REL_ALIASES = {
    "Activity ID": ["Activity ID", "activity_id"],
    "Predecessor ID": ["Predecessor ID", "predecessor_id", "predecessor_details"],
    "Relationship Type": ["Relationship Type", "relationship_type"],
    "Lag": ["Lag", "lag"],
    "Successor ID": ["Successor ID", "successor_id", "successor_details"],
}

CLIENT_ACTIVITY_SUPPLY_ALIASES = {
    "Resource ID": ["Resource ID", "resource_id", "resource id"],
    "Activity ID": ["Activity ID", "activity_id", "activity id"],
    "Activity Name": ["Activity Name", "activity_name", "activity name"],
    "Original Duration": ["Original Duration", "original_duration"],
    "Actual Duration": ["Actual Duration", "actual_duration"],
    "Remaining Duration": ["Remaining Duration", "remaining_duration"],
    "Deviated Duration": ["Deviated Duration", "deviated_duration"],
    "Planned Start": ["Planned Start", "planned_start"],
    "Planned Finish": ["Planned Finish", "planned_finish"],
    "BL Start": ["BL Start", "Baseline Start", "baseline start", "bl start"],
    "BL Finish": ["BL Finish", "Baseline Finish", "baseline finish", "bl finish"],
    "Start": ["Start", "start"],
    "Finish": ["Finish", "finish"],
    "Actual Finish": ["Actual Finish", "actual_finish"],
    "Budgeted Units": ["Budgeted Units", "budgeted_units"],
    "Actual Units": ["Actual Units", "actual_units"],
    "Date of receiving units supplied by client": ["Date of receiving units supplied by client"],
    "Available units at site received by client": ["Available units at site received by client"],
    "Available units supplied by client vs Actual units of activities": ["Available units supplied by client vs Actual units of activities"],
    "Remaining Units": ["Remaining Units", "remaining_units"],
    "Units % Complete": ["Units % Complete", "units % complete", "physical % complete"],
    "Budgeted Units / Time": ["Budgeted Units / Time", "budgeted units / time"],
    "Drive Activity Dates": ["Drive Activity Dates", "drive activity dates", "Longest Path"],
    "Reason for Delay": ["Reason for Delay", "reason for delay"],
    "Responsibility": ["Responsibility", "responsibility"],
}

CONTRACT_TOPIC_KEYWORDS = {
    "steel": ["steel", "rebar", "reinforcement", "rft", "free issue", "employer supplied", "client supplied", "material shortage"],
    "delay": ["delay", "eot", "extension", "time", "critical path", "float", "tia"],
    "payment": ["payment", "invoice", "cashflow", "paid", "certified"],
    "design": ["ifc", "shop drawing", "design", "rfi", "approval", "engineer response"],
    "access": ["access", "possession", "work front", "handover"],
}

CIVIL_KEYWORD_GROUPS = {
    "Reinforcement / Steel": ["rebar", "reinforcement", "rft", "steel fixing", "fix reinforcement", "bbs", "bar bending", "steel delivery", "steel shortage", "client supplied steel", "employer supplied steel", "free issue material"],
    "Concrete Works": ["concrete", "casting", "pouring", "pour", "slab casting", "column casting", "beam casting", "wall casting", "raft casting"],
    "Formwork": ["formwork", "shuttering", "striking", "de-shuttering", "table form", "climbing form"],
    "Civil / Structural Elements": ["foundation", "footing", "raft", "pile cap", "basement", "retaining wall", "column", "wall", "core wall", "shear wall", "beam", "slab", "staircase", "ramp", "roof", "transfer slab"],
    "Design / Engineering": ["ifc", "shop drawing", "drawing", "rfi", "design clarification", "approval", "material approval", "mir", "inspection"],
    "Access / Handover": ["site access", "handover", "possession", "obstruction", "area not available", "work front", "predecessor not completed"],
}


@dataclass
class SteelTiaSettings:
    usability_lag_days: int = 2
    near_critical_float_threshold: int = 10
    data_date: pd.Timestamp | None = None


def normalize_name(name: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(name).strip().lower())


def standardize_building(value: Any) -> str:
    text = str(value or "").upper()
    match = re.search(r"B0?[1-4]", text)
    return match.group(0).replace("B0", "B0") if match else "UNALLOCATED"


def standardize_steel_type(value: Any) -> str:
    text = str(value or "").strip()
    if not text:
        return "UNSPECIFIED STEEL TYPE"
    upper = text.upper()
    diameter_match = re.search(r"(\d{1,2}\s*MM)", upper)
    if diameter_match:
        return diameter_match.group(1).replace(" ", "")
    if "RFT" in upper:
        return "RFT"
    if "REBAR" in upper or "REINFORC" in upper:
        return "REINFORCEMENT"
    return upper


def infer_steel_type_from_context(*parts: Any) -> str:
    text = " ".join(str(part or "") for part in parts)
    inferred = standardize_steel_type(text)
    if inferred != "UNSPECIFIED STEEL TYPE":
        return inferred
    lower = text.lower()
    if any(token in lower for token in ["rft", "rebar", "reinforcement", "steel fixing", "bar bending"]):
        return "REINFORCEMENT"
    return "UNSPECIFIED STEEL TYPE"


def safe_numeric(series: pd.Series) -> pd.Series:
    return pd.to_numeric(
        series.astype(str).str.replace(",", "", regex=False).str.replace("$", "", regex=False).str.replace("%", "", regex=False).str.strip(),
        errors="coerce",
    ).fillna(0.0)


def parse_mixed_date(value: Any) -> Any:
    if value is None or pd.isna(value):
        return pd.NaT
    if isinstance(value, pd.Timestamp):
        return value

    text = str(value).strip()
    if not text:
        return pd.NaT

    text = re.sub(r"[​‎‏﻿]", "", text)
    text = re.sub(r"\*+$", "", text).strip()
    text = re.sub(r"\s+[A-Za-z]$", "", text).strip()
    text = re.sub(r"\s+", " ", text)
    if not text:
        return pd.NaT

    if re.fullmatch(r"\d+(\.\d+)?", text):
        number = float(text)
        if 20000 < number < 80000:
            return pd.to_datetime("1899-12-30") + pd.to_timedelta(number, unit="D")

    iso_match = re.fullmatch(r"\d{4}-\d{1,2}-\d{1,2}(?: \d{1,2}:\d{2}(?::\d{2})?)?", text)
    if iso_match:
        return pd.to_datetime(text, errors="coerce", dayfirst=False)

    slash_match = re.fullmatch(r"(\d{1,2})/(\d{1,2})/(\d{2,4})", text)
    if slash_match:
        first = int(slash_match.group(1))
        second = int(slash_match.group(2))
        if first > 12 and second <= 12:
            return pd.to_datetime(text, errors="coerce", dayfirst=True)
        return pd.to_datetime(text, errors="coerce", dayfirst=False)

    for fmt in ("%d-%b-%y", "%d-%b-%Y", "%d/%m/%Y", "%m/%d/%Y"):
        parsed = pd.to_datetime(text, format=fmt, errors="coerce")
        if pd.notna(parsed):
            return parsed

    parsed = pd.to_datetime(text, errors="coerce", dayfirst=True)
    if pd.notna(parsed):
        return parsed
    return pd.to_datetime(text, errors="coerce", dayfirst=False)


def load_table(source: Any) -> pd.DataFrame:
    if source is None:
        return pd.DataFrame()
    if hasattr(source, "name"):
        name = str(source.name).lower()
        if name.endswith(".xlsx") or name.endswith(".xls"):
            return pd.read_excel(source).fillna("")
        return pd.read_csv(source).fillna("")
    path = Path(source)
    if not path.exists():
        return pd.DataFrame()
    if path.suffix.lower() in {".xlsx", ".xls"}:
        return pd.read_excel(path).fillna("")
    return pd.read_csv(path).fillna("")


def suggest_mapping(df: pd.DataFrame, aliases: dict[str, list[str]]) -> dict[str, str]:
    normalized = {normalize_name(col): col for col in df.columns}
    mapping: dict[str, str] = {}
    for target, candidates in aliases.items():
        for candidate in candidates:
            match = normalized.get(normalize_name(candidate))
            if match:
                mapping[target] = match
                break
    return mapping


def apply_mapping(df: pd.DataFrame, mapping: dict[str, str], required_fields: list[str]) -> tuple[pd.DataFrame, list[str], list[str]]:
    out = pd.DataFrame()
    missing: list[str] = []
    for field in required_fields:
        src = mapping.get(field)
        if src and src in df.columns:
            out[field] = df[src]
        else:
            out[field] = ""
            missing.append(field)
    duplicates = []
    if "Activity ID" in out.columns:
        dup_mask = out["Activity ID"].astype(str).str.strip().duplicated(keep=False)
        duplicates = sorted(out.loc[dup_mask, "Activity ID"].astype(str).str.strip().unique().tolist())
    out = out.fillna("")
    return out, missing, duplicates


def infer_civil_package(activity_name: str) -> str:
    text = str(activity_name or "").lower()
    for group, keywords in CIVIL_KEYWORD_GROUPS.items():
        if any(keyword in text for keyword in keywords):
            return group
    return "General Construction"


def _steel_activity_match(text: str) -> list[str]:
    lower = str(text or "").lower()
    matches: list[str] = []
    for keyword in CIVIL_KEYWORD_GROUPS["Reinforcement / Steel"]:
        if keyword in lower:
            matches.append(keyword)
    return sorted(set(matches))


def _parse_pct_value(value: Any) -> float:
    text = str(value or "").replace("%", "").replace(",", "").strip()
    try:
        return float(text)
    except ValueError:
        return 0.0


def build_auto_requirement_df(p6_df: pd.DataFrame) -> pd.DataFrame:
    if p6_df.empty:
        return pd.DataFrame(columns=CANONICAL_REQ_FIELDS + ["Requirement Basis", "Selection Explanation", "Steel Logic Match"])

    df = p6_df.copy()
    for col in ["Activity ID", "Activity Name", "WBS", "Activity Code", "Start", "Baseline Start", "Actual Finish", "Remaining Units", "Budgeted Units", "Physical % Complete", "Critical", "Longest Path", "Total Float", "Required Steel Qty", "Required Date", "Steel Type / Diameter"]:
        if col not in df.columns:
            df[col] = ""

    df["Activity ID"] = df["Activity ID"].astype(str).str.strip().str.upper()
    df["Activity Name"] = df["Activity Name"].astype(str).str.strip()
    df["WBS"] = df["WBS"].astype(str).str.strip()
    df["Activity Code"] = df["Activity Code"].astype(str).str.strip()
    df["Remaining Units Numeric"] = safe_numeric(df["Remaining Units"])
    df["Budgeted Units Numeric"] = safe_numeric(df["Budgeted Units"])
    df["Linked Required Steel Qty Numeric"] = safe_numeric(df["Required Steel Qty"])
    df["Physical % Complete Numeric"] = df["Physical % Complete"].apply(_parse_pct_value)
    df["Actual Finish Parsed"] = df["Actual Finish"].apply(parse_mixed_date)
    df["Required Date Parsed"] = df["Required Date"].replace("", pd.NA).fillna(df["Start"]).replace("", pd.NA).fillna(df["Baseline Start"]).apply(parse_mixed_date)
    df["Building Derived"] = df["Activity ID"].apply(standardize_building)

    steel_logic_matches: list[list[str]] = []
    for _, row in df.iterrows():
        text = " ".join([str(row["Activity Name"]), str(row["Activity Code"]), str(row["WBS"])])
        steel_logic_matches.append(_steel_activity_match(text))
    df["Steel Logic Match"] = steel_logic_matches
    df["Steel Relevant"] = df["Steel Logic Match"].apply(bool)
    df["Not Completed"] = (df["Physical % Complete Numeric"] < 100) | df["Actual Finish Parsed"].isna()

    remaining_positive = df["Remaining Units Numeric"] > 0
    budget_based_qty = df["Budgeted Units Numeric"] * (1 - (df["Physical % Complete Numeric"] / 100.0))
    df["Auto Required Steel Qty"] = df["Remaining Units Numeric"].where(remaining_positive, budget_based_qty)
    df.loc[df["Auto Required Steel Qty"] <= 0, "Auto Required Steel Qty"] = df["Linked Required Steel Qty Numeric"]
    df["Auto Required Steel Qty"] = df["Auto Required Steel Qty"].clip(lower=0.0)

    df["Requirement Basis"] = "No quantity basis"
    df.loc[remaining_positive, "Requirement Basis"] = "Derived from Remaining Units"
    df.loc[~remaining_positive & (budget_based_qty > 0), "Requirement Basis"] = "Derived from Budgeted Units x (1 - Physical % Complete)"
    df.loc[(df["Auto Required Steel Qty"] > 0) & df["Requirement Basis"].eq("No quantity basis"), "Requirement Basis"] = "Derived from linked steel quantity register"

    df["Priority"] = df.apply(
        lambda row: (
            "Critical"
            if _critical_flag(row["Critical"]) or _longest_path_flag(row["Longest Path"])
            else ("Near-Critical" if _float_value(row["Total Float"]) <= 10 else "Normal")
        ),
        axis=1,
    )

    selected = df[
        df["Steel Relevant"]
        & df["Not Completed"]
        & (df["Auto Required Steel Qty"] > 0)
        & df["Required Date Parsed"].notna()
    ].copy()
    if selected.empty:
        return pd.DataFrame(columns=CANONICAL_REQ_FIELDS + ["Requirement Basis", "Selection Explanation", "Steel Logic Match"])

    selected["Steel Type / Diameter"] = selected.apply(
        lambda row: infer_steel_type_from_context(
            row.get("Steel Type / Diameter", ""),
            row.get("Activity Code", ""),
            row.get("Activity Name", ""),
            row.get("WBS", ""),
        ),
        axis=1,
    )
    selected.loc[selected["Steel Type / Diameter"].eq("GENERAL REINFORCEMENT STEEL"), "Steel Type / Diameter"] = "RFT STEEL"
    selected["Selection Explanation"] = selected.apply(
        lambda row: "; ".join(
            [
                "Selected because activity is steel-related",
                f"matched keywords: {', '.join(row['Steel Logic Match']) or 'activity code / WBS steel signal'}",
                "activity is not completed",
                f"required quantity {row['Auto Required Steel Qty']:.2f} computed from {row['Requirement Basis'].lower()}",
                f"required date taken from {display_date_for_requirement(row['Required Date Parsed'])}",
                f"priority classified as {row['Priority']}",
            ]
        ),
        axis=1,
    )
    selected["Remarks"] = selected["Selection Explanation"]

    return selected[
        [
            "Activity ID",
            "Activity Name",
            "Building Derived",
            "Steel Type / Diameter",
            "Auto Required Steel Qty",
            "Required Date Parsed",
            "Priority",
            "Remarks",
            "Requirement Basis",
            "Selection Explanation",
            "Steel Logic Match",
        ]
    ].rename(
        columns={
            "Building Derived": "Building",
            "Auto Required Steel Qty": "Required Steel Qty",
            "Required Date Parsed": "Required Date",
        }
    ).reset_index(drop=True)


def build_requirement_df_from_client_supply_sheet(client_supply_df: pd.DataFrame) -> pd.DataFrame:
    if client_supply_df.empty:
        return pd.DataFrame(columns=CANONICAL_REQ_FIELDS + ["Requirement Basis", "Selection Explanation", "Steel Logic Match"])

    df = client_supply_df.copy()
    for col in CANONICAL_CLIENT_ACTIVITY_SUPPLY_FIELDS:
        if col not in df.columns:
            df[col] = ""

    df["Activity ID"] = df["Activity ID"].astype(str).str.strip().str.upper()
    df["Activity Name"] = df["Activity Name"].astype(str).str.strip()
    df["Resource ID"] = df["Resource ID"].astype(str).str.strip()
    df["Remaining Units Numeric"] = safe_numeric(df["Remaining Units"])
    df["Budgeted Units Numeric"] = safe_numeric(df["Budgeted Units"])
    df["Actual Units Numeric"] = safe_numeric(df["Actual Units"])
    df["Client Available Qty Numeric"] = safe_numeric(df["Available units at site received by client"])
    df["Client Supply vs Actual Gap Numeric"] = safe_numeric(df["Available units supplied by client vs Actual units of activities"])
    df["Units % Complete Numeric"] = df["Units % Complete"].apply(_parse_pct_value)
    df["Required Date Parsed"] = (
        df["Start"]
        .replace("", pd.NA)
        .fillna(df["Planned Start"])
        .replace("", pd.NA)
        .fillna(df["BL Start"])
        .replace("", pd.NA)
        .fillna(df["Date of receiving units supplied by client"])
        .apply(parse_mixed_date)
    )
    df["Client Supply Receipt Date Parsed"] = df["Date of receiving units supplied by client"].apply(parse_mixed_date)
    df["Planned Start Parsed"] = df["Planned Start"].apply(parse_mixed_date)
    df["BL Start Parsed"] = df["BL Start"].apply(parse_mixed_date)
    df["Building Derived"] = df["Activity ID"].apply(standardize_building)

    df["Steel Logic Match"] = df.apply(
        lambda row: _steel_activity_match(
            " ".join(
                [
                    str(row.get("Activity Name", "")),
                    str(row.get("Resource ID", "")),
                    str(row.get("Responsibility", "")),
                ]
            )
        ),
        axis=1,
    )
    df["Steel Relevant"] = df["Steel Logic Match"].apply(bool)
    reason_series = df["Reason for Delay"].astype(str).str.strip().str.lower()
    responsibility_series = df["Responsibility"].astype(str).str.strip().str.lower()
    df["Historical Employer Delay Flag"] = reason_series.str.contains("no rft", na=False) & responsibility_series.str.contains("employer", na=False)
    df["Not Completed"] = (df["Units % Complete Numeric"] < 100) & (df["Remaining Units Numeric"] > 0)

    budget_minus_actual = (df["Budgeted Units Numeric"] - df["Actual Units Numeric"]).clip(lower=0.0)
    df["Auto Required Steel Qty"] = df["Remaining Units Numeric"].where(df["Remaining Units Numeric"] > 0, budget_minus_actual)
    df.loc[df["Historical Employer Delay Flag"] & (df["Auto Required Steel Qty"] <= 0), "Auto Required Steel Qty"] = df.loc[df["Historical Employer Delay Flag"] & (df["Auto Required Steel Qty"] <= 0), "Actual Units Numeric"].clip(lower=0.0)
    df["Requirement Basis"] = "Derived from client supply vs actual activity sheet"
    df["Steel Type / Diameter"] = df.apply(
        lambda row: infer_steel_type_from_context(
            row.get("Resource ID", ""),
            row.get("Activity Name", ""),
        ),
        axis=1,
    )
    df["Priority"] = df["Drive Activity Dates"].astype(str).str.strip().str.lower().apply(
        lambda value: "Critical" if "yes" in value else "Normal"
    )
    df["Selection Explanation"] = df.apply(
        lambda row: "; ".join(
            [
                "Selected from uploaded client available units vs actual units sheet",
                "activity is steel-related and monitored at activity level",
                f"remaining steel requirement {row['Auto Required Steel Qty']:.2f} derived from remaining units / budget minus actual",
                f"client available quantity recorded as {row['Client Available Qty Numeric']:.2f}",
                f"client supply vs actual gap recorded as {row['Client Supply vs Actual Gap Numeric']:.2f}",
                f"required date taken from {display_date_for_requirement(row['Required Date Parsed'])}",
                f"drive activity dates flag is {str(row.get('Drive Activity Dates', '')).strip() or 'No'}",
            ]
        ),
        axis=1,
    )
    df["Remarks"] = df["Selection Explanation"]

    selected = df[df["Steel Relevant"] & (df["Not Completed"] | df["Historical Employer Delay Flag"])].copy()
    selected = selected[(selected["Auto Required Steel Qty"] > 0) & selected["Required Date Parsed"].notna()]
    if selected.empty:
        return pd.DataFrame(columns=CANONICAL_REQ_FIELDS + ["Requirement Basis", "Selection Explanation", "Steel Logic Match"])

    return selected[
        [
            "Activity ID",
            "Activity Name",
            "Building Derived",
            "Steel Type / Diameter",
            "Auto Required Steel Qty",
            "Required Date Parsed",
            "Priority",
            "Remarks",
            "Requirement Basis",
            "Selection Explanation",
            "Steel Logic Match",
            "Client Available Qty Numeric",
            "Client Supply vs Actual Gap Numeric",
            "Client Supply Receipt Date Parsed",
            "Planned Start Parsed",
            "BL Start Parsed",
            "Drive Activity Dates",
            "Units % Complete Numeric",
            "Reason for Delay",
            "Responsibility",
            "Historical Employer Delay Flag",
        ]
    ].rename(
        columns={
            "Building Derived": "Building",
            "Auto Required Steel Qty": "Required Steel Qty",
            "Required Date Parsed": "Required Date",
            "Client Available Qty Numeric": "Client Available Qty",
            "Client Supply vs Actual Gap Numeric": "Client Supply vs Actual Gap",
            "Client Supply Receipt Date Parsed": "Client Supply Receipt Date",
            "Planned Start Parsed": "Planned Start Date",
            "BL Start Parsed": "BL Start Date",
            "Units % Complete Numeric": "Activity Units % Complete",
            "Historical Employer Delay Flag": "Historical Employer Delay Flag",
        }
    ).reset_index(drop=True)


def display_date_for_requirement(value: Any) -> str:
    ts = pd.to_datetime(value, errors="coerce")
    return ts.strftime("%d-%b-%Y") if not pd.isna(ts) else "N/A"


def prepare_default_p6_export(activities_path: Path, relationships_path: Path) -> pd.DataFrame:
    activities_df = load_table(activities_path)
    relationships_df = load_table(relationships_path)

    rel_map = suggest_mapping(relationships_df, P6_ALIASES)
    rel_df, _, _ = apply_mapping(relationships_df, rel_map, CANONICAL_P6_FIELDS)

    rel_df["Activity ID"] = rel_df["Activity ID"].astype(str).str.strip().str.lower()
    merged = rel_df.copy()

    if not activities_df.empty and "activity_id" in activities_df.columns:
        act = activities_df.copy()
        act["activity_id"] = act["activity_id"].astype(str).str.strip().str.lower()
        merged = merged.merge(
            act[["activity_id", "wbs_id", "actual_finish", "actual_start", "planned_start", "planned_finish", "forecast_start", "forecast_finish", "is_critical", "total_float_days", "responsible_party"]],
            left_on="Activity ID",
            right_on="activity_id",
            how="left",
        )
        merged["WBS"] = merged["WBS"].replace("", pd.NA).fillna(merged.get("wbs_id", ""))
        merged["Actual Finish"] = merged["Actual Finish"].replace("", pd.NA).fillna(merged.get("actual_finish", ""))
        merged["Actual Start"] = merged["Actual Start"].replace("", pd.NA).fillna(merged.get("actual_start", ""))
        merged["Baseline Start"] = merged["Baseline Start"].replace("", pd.NA).fillna(merged.get("planned_start", ""))
        merged["Baseline Finish"] = merged["Baseline Finish"].replace("", pd.NA).fillna(merged.get("planned_finish", ""))
        merged["Critical"] = merged["Critical"].replace("", pd.NA).fillna(merged.get("is_critical", ""))
        merged["Total Float"] = merged["Total Float"].replace("", pd.NA).fillna(merged.get("total_float_days", ""))

    merged["Activity ID"] = merged["Activity ID"].astype(str).str.upper()
    merged["Building"] = merged["Activity ID"].apply(standardize_building)
    if "Steel Type / Diameter" not in merged.columns:
        merged["Steel Type / Diameter"] = ""
    merged["Steel Type / Diameter"] = merged.apply(
        lambda row: infer_steel_type_from_context(
            row.get("Steel Type / Diameter", ""),
            row.get("Activity Code", ""),
            row.get("Activity Name", ""),
            row.get("WBS", ""),
        ),
        axis=1,
    )
    if "Required Steel Qty" not in merged.columns:
        merged["Required Steel Qty"] = ""
    if "Required Date" not in merged.columns:
        merged["Required Date"] = ""
    merged["Required Date"] = merged["Required Date"].replace("", pd.NA).fillna(merged["Start"]).replace("", pd.NA).fillna(merged["Baseline Start"])
    merged["Physical % Complete"] = merged["Physical % Complete"].replace("", pd.NA).fillna(merged.get("activity_%_complete", ""))
    return merged


def contract_clauses_to_df(clauses: list[Any]) -> pd.DataFrame:
    rows = []
    for clause in clauses:
        rows.append(
            {
                "Clause / Topic": getattr(clause, "topic", ""),
                "Location": getattr(clause, "location", ""),
                "Plain English Meaning": getattr(clause, "plain_english", ""),
                "Research the Lines": getattr(clause, "beneath_lines", ""),
                "Who Holds Leverage": getattr(clause, "leverage_holder", ""),
                "Notice / Time Bar": getattr(clause, "notice_requirement", ""),
                "Money Impact": getattr(clause, "money_impact", ""),
                "Schedule Impact": getattr(clause, "schedule_impact", ""),
                "Practical Action / Evidence": getattr(clause, "practical_action", ""),
            }
        )
    return pd.DataFrame(rows)


def clean_contract_library(df: pd.DataFrame) -> pd.DataFrame:
    aliases = {
        "Clause / Topic": ["Clause / Topic", "clause topic"],
        "Location": ["Location"],
        "Plain English Meaning": ["Plain English Meaning"],
        "Research the Lines": ["Research the Lines", "Beneath the Lines"],
        "Who Holds Leverage": ["Who Holds Leverage"],
        "Notice / Time Bar": ["Notice / Time Bar"],
        "Money Impact": ["Money Impact"],
        "Schedule Impact": ["Schedule Impact"],
        "Practical Action / Evidence": ["Practical Action / Evidence"],
    }
    mapping = suggest_mapping(df, aliases)
    canonical, _, _ = apply_mapping(df, mapping, list(aliases.keys()))
    return canonical.replace({pd.NA: "", None: ""}).fillna("")


def keyword_match_score(text: str, keyword_groups: list[str]) -> int:
    lower = text.lower()
    return sum(1 for key in keyword_groups if key.lower() in lower)


def score_steel_supply_clause(searchable: str, event_text: str) -> tuple[int, str]:
    lower = searchable.lower()
    event_lower = event_text.lower()
    score = 0
    rationale = []

    steel_event = any(term in event_lower for term in ["steel", "rft", "reinforcement", "free issue", "client free issue"])
    if not steel_event:
        return 0, ""

    if "steel" in lower and any(term in lower for term in ["employer-supplied", "employer supplied", "client supplied", "free-issue", "free issue"]):
        score += 18
        rationale.append("steel/free-issue supply")
    if any(term in lower for term in ["approved programme", "approved program"]) and any(term in lower for term in ["requisition", "delivery note", "delivery notes"]):
        score += 12
        rationale.append("programme/requisition evidence")
    if any(term in lower for term in ["possible eot", "late steel may support delay", "critical and properly requested", "critical and not contractor-caused"]):
        score += 10
        rationale.append("critical EOT language")
    if any(term in lower for term in ["stock record", "stock records", "activity impact", "activity impacts", "impact analysis"]):
        score += 8
        rationale.append("activity-impact evidence")
    if "clause 4.20" in lower:
        score += 6
        rationale.append("specific free-issue clause")

    if any(term in lower for term in ["prequalification", "supplier prequalification", "supplier approval"]):
        score -= 10
        rationale.append("penalized supplier-prequalification row")
    if any(term in lower for term in ["capmas", "price adjustment", "inflation", "price movement"]):
        score -= 8
        rationale.append("penalized price-only row")

    return max(score, 0), "; ".join(rationale)


def build_daily_balance(requirements_df: pd.DataFrame, steel_df: pd.DataFrame, settings: SteelTiaSettings) -> pd.DataFrame:
    req = requirements_df.copy()
    deliveries = steel_df.copy()

    req["Required Date"] = pd.to_datetime(req["Required Date"].apply(parse_mixed_date), errors="coerce")
    req["Required Steel Qty"] = safe_numeric(req["Required Steel Qty"])
    req = req[req["Required Date"].notna()].copy()
    req["Building"] = req["Building"].apply(standardize_building)
    req["Steel Type / Diameter"] = req["Steel Type / Diameter"].apply(standardize_steel_type)

    deliveries["Date"] = pd.to_datetime(deliveries["Date"].apply(parse_mixed_date), errors="coerce")
    deliveries["Delivered Qty"] = safe_numeric(deliveries["Delivered Qty"])
    deliveries["Usable Steel Date"] = deliveries["Date"] + pd.to_timedelta(int(settings.usability_lag_days), unit="D")
    deliveries["Steel Type / Diameter"] = deliveries["Steel Type / Diameter"].apply(standardize_steel_type)
    deliveries = deliveries[deliveries["Usable Steel Date"].notna()].copy()

    all_dates = sorted(set(req["Required Date"].dropna().dt.normalize().tolist()) | set(deliveries["Usable Steel Date"].dropna().dt.normalize().tolist()))
    if not all_dates:
        return pd.DataFrame(columns=[
            "Date", "Building", "Steel Type / Diameter", "Activity ID", "Activity Name", "Required Steel Qty",
            "Client Delivered Qty", "Cumulative Delivered", "Cumulative Required", "Steel Balance", "Stock-Out?", "Shortage Qty"
        ])

    req_by_date = req.groupby(req["Required Date"].dt.normalize())["Required Steel Qty"].sum()
    del_by_date = deliveries.groupby(deliveries["Usable Steel Date"].dt.normalize())["Delivered Qty"].sum()

    cumulative_delivered = 0.0
    cumulative_required = 0.0
    rows: list[dict[str, Any]] = []
    req_lookup = req.groupby(req["Required Date"].dt.normalize())

    for date in all_dates:
        daily_delivered = float(del_by_date.get(date, 0.0))
        cumulative_delivered += daily_delivered
        day_rows = req_lookup.get_group(date) if date in req_lookup.groups else pd.DataFrame(columns=req.columns)
        day_required_total = float(day_rows["Required Steel Qty"].sum()) if not day_rows.empty else 0.0
        cumulative_required += day_required_total
        steel_balance = cumulative_delivered - cumulative_required
        stock_out = "YES" if steel_balance <= 0 else "NO"
        shortage_qty = abs(steel_balance) if steel_balance < 0 else 0.0

        if day_rows.empty:
            rows.append({
                "Date": date,
                "Building": "ALL",
                "Steel Type / Diameter": "ALL STEEL TYPES",
                "Activity ID": "",
                "Activity Name": "Delivery / balance checkpoint",
                "Required Steel Qty": 0.0,
                "Client Delivered Qty": daily_delivered,
                "Cumulative Delivered": cumulative_delivered,
                "Cumulative Required": cumulative_required,
                "Steel Balance": steel_balance,
                "Stock-Out?": stock_out,
                "Shortage Qty": shortage_qty,
            })
        else:
            for _, row in day_rows.iterrows():
                rows.append({
                    "Date": date,
                    "Building": row["Building"],
                    "Steel Type / Diameter": row["Steel Type / Diameter"],
                    "Activity ID": row["Activity ID"],
                    "Activity Name": row["Activity Name"],
                    "Required Steel Qty": float(row["Required Steel Qty"]),
                    "Client Delivered Qty": daily_delivered,
                    "Cumulative Delivered": cumulative_delivered,
                    "Cumulative Required": cumulative_required,
                    "Steel Balance": steel_balance,
                    "Stock-Out?": stock_out,
                    "Shortage Qty": shortage_qty,
                })
    return pd.DataFrame(rows)


def build_stock_out_events(balance_df: pd.DataFrame) -> pd.DataFrame:
    if balance_df.empty:
        return pd.DataFrame(columns=["Stock-Out Date", "Steel Type", "Building", "Quantity Shortage", "Steel Recovered Later?", "Recovery Date"])
    summary = balance_df.sort_values("Date").groupby("Date", as_index=False).agg({
        "Client Delivered Qty": "sum",
        "Cumulative Delivered": "max",
        "Cumulative Required": "max",
        "Steel Balance": "min",
        "Shortage Qty": "max",
    })
    summary["prior_balance"] = summary["Steel Balance"].shift(1).fillna(1)
    event_rows = []
    later_positive_dates = summary.loc[summary["Steel Balance"] > 0, "Date"]
    for _, row in summary[(summary["Steel Balance"] <= 0) & (summary["prior_balance"] > 0)].iterrows():
        later = later_positive_dates[later_positive_dates > row["Date"]]
        event_rows.append({
            "Stock-Out Date": row["Date"],
            "Steel Type": "ALL STEEL TYPES",
            "Building": "PROJECT-WIDE",
            "Quantity Shortage": abs(float(row["Steel Balance"])) if float(row["Steel Balance"]) < 0 else 0.0,
            "Steel Recovered Later?": "Yes" if not later.empty else "No",
            "Recovery Date": later.iloc[0] if not later.empty else pd.NaT,
        })
    return pd.DataFrame(event_rows)


def build_stock_out_events_from_activity_supply(requirements_df: pd.DataFrame) -> pd.DataFrame:
    if requirements_df.empty or "Client Supply vs Actual Gap" not in requirements_df.columns:
        return pd.DataFrame(columns=["Stock-Out Date", "Steel Type", "Building", "Quantity Shortage", "Steel Recovered Later?", "Recovery Date"])

    df = requirements_df.copy()
    df["Required Date"] = pd.to_datetime(df["Required Date"].apply(parse_mixed_date), errors="coerce")
    df["Client Supply vs Actual Gap"] = safe_numeric(df["Client Supply vs Actual Gap"])
    df["Steel Type / Diameter"] = df["Steel Type / Diameter"].apply(standardize_steel_type)
    df["Building"] = df["Building"].apply(standardize_building)
    shortage_rows = df[(df["Required Date"].notna()) & (df["Client Supply vs Actual Gap"] < 0)].copy()
    if shortage_rows.empty:
        return pd.DataFrame(columns=["Stock-Out Date", "Steel Type", "Building", "Quantity Shortage", "Steel Recovered Later?", "Recovery Date"])

    grouped = (
        shortage_rows.groupby(
            [shortage_rows["Required Date"].dt.normalize(), "Building", "Steel Type / Diameter"],
            as_index=False,
        )["Client Supply vs Actual Gap"]
        .min()
    )
    grouped.columns = ["Stock-Out Date", "Building", "Steel Type", "Client Supply vs Actual Gap"]
    grouped["Quantity Shortage"] = grouped["Client Supply vs Actual Gap"].abs()
    grouped["Steel Recovered Later?"] = "Unknown"
    grouped["Recovery Date"] = pd.NaT
    return grouped[
        ["Stock-Out Date", "Steel Type", "Building", "Quantity Shortage", "Steel Recovered Later?", "Recovery Date"]
    ].sort_values(["Stock-Out Date", "Building", "Steel Type"]).reset_index(drop=True)


def _critical_flag(value: Any) -> bool:
    text = str(value or "").strip().lower()
    return text in {"yes", "true", "1", "critical"} or "yes" in text


def _longest_path_flag(value: Any) -> bool:
    text = str(value or "").strip().lower()
    return text in {"yes", "true", "1"} or "drive activity dates" in text or "longest" in text


def _float_value(value: Any) -> float:
    text = str(value or "").replace(",", "").strip()
    try:
        return float(text)
    except ValueError:
        return 0.0


def build_affected_candidates(p6_df: pd.DataFrame, balance_df: pd.DataFrame, settings: SteelTiaSettings, stock_events_df: pd.DataFrame | None = None) -> pd.DataFrame:
    if p6_df.empty:
        return pd.DataFrame()
    stock_events = stock_events_df.copy() if stock_events_df is not None else build_stock_out_events(balance_df)
    if stock_events.empty:
        return pd.DataFrame()
    stock_events["Stock-Out Date"] = stock_events["Stock-Out Date"].apply(parse_mixed_date)
    stock_events["Building"] = stock_events["Building"].apply(standardize_building)
    stock_events["Steel Type"] = stock_events["Steel Type"].apply(standardize_steel_type)
    stock_out_date = stock_events["Stock-Out Date"].min() if not stock_events.empty else pd.NaT
    if pd.isna(stock_out_date):
        return pd.DataFrame()

    df = p6_df.copy()
    df["Required Date"] = df["Required Date"].apply(parse_mixed_date)
    df["Required Steel Qty"] = safe_numeric(df["Required Steel Qty"])
    df["Total Float"] = df["Total Float"].apply(_float_value)
    df["Physical % Complete"] = safe_numeric(df["Physical % Complete"])
    if "Client Supply vs Actual Gap" not in df.columns:
        df["Client Supply vs Actual Gap"] = 0
    if "Client Available Qty" not in df.columns:
        df["Client Available Qty"] = 0
    if "Activity Units % Complete" not in df.columns:
        df["Activity Units % Complete"] = df["Physical % Complete"]
    if "Client Supply Receipt Date" not in df.columns:
        df["Client Supply Receipt Date"] = pd.NaT
    if "Planned Start Date" not in df.columns:
        df["Planned Start Date"] = pd.NaT
    if "BL Start Date" not in df.columns:
        df["BL Start Date"] = pd.NaT
    if "Historical Employer Delay Flag" not in df.columns:
        df["Historical Employer Delay Flag"] = False
    df["Client Supply vs Actual Gap"] = safe_numeric(df["Client Supply vs Actual Gap"])
    df["Client Available Qty"] = safe_numeric(df["Client Available Qty"])
    df["Client Supply Receipt Date"] = df["Client Supply Receipt Date"].apply(parse_mixed_date)
    df["Planned Start Date"] = df["Planned Start Date"].apply(parse_mixed_date)
    df["BL Start Date"] = df["BL Start Date"].apply(parse_mixed_date)
    df["Historical Employer Delay Flag"] = df["Historical Employer Delay Flag"].astype(str).str.strip().str.lower().isin(["true", "1", "yes"])
    activity_status_series = df["Activity Status"] if "Activity Status" in df.columns else (df["activity_status"] if "activity_status" in df.columns else pd.Series([""] * len(df), index=df.index))
    df["Activity Status Normalized"] = activity_status_series.astype(str).str.strip().str.lower()
    df["Actual Finish Parsed"] = df["Actual Finish"].apply(parse_mixed_date)
    df["Building"] = df["Building"].apply(standardize_building)
    df["Steel Type / Diameter"] = df["Steel Type / Diameter"].apply(standardize_steel_type)
    df["Civil / Structural Work Package"] = df["Activity Name"].apply(infer_civil_package)
    df["Construction Sequence Impact"] = df["Activity Name"].apply(
        lambda x: "Structural sequence impacted" if any(k in str(x).lower() for k in ["rft", "reinforcement", "beam", "slab", "column", "wall", "raft", "footing"]) else "General sequence impact"
    )

    candidates = []
    for _, row in df.iterrows():
        required_qty = float(row["Required Steel Qty"])
        required_date = row["Required Date"]
        if required_qty <= 0 or pd.isna(required_date):
            continue
        test1 = required_qty > 0
        historical_delay_flag = bool(row.get("Historical Employer Delay Flag", False))
        planned_start_date = row.get("Planned Start Date")
        planned_start_date = pd.to_datetime(planned_start_date, errors="coerce")
        client_supply_receipt_date = pd.to_datetime(row.get("Client Supply Receipt Date"), errors="coerce")
        related_events = stock_events[(stock_events["Building"] == row["Building"]) & (stock_events["Steel Type"] == row["Steel Type / Diameter"])]
        candidate_stock_out_date = related_events["Stock-Out Date"].min() if not related_events.empty else pd.NaT
        if historical_delay_flag:
            due_anchor = planned_start_date if not pd.isna(planned_start_date) else required_date
            if not pd.isna(due_anchor):
                candidate_stock_out_date = due_anchor if pd.isna(candidate_stock_out_date) else min(candidate_stock_out_date, due_anchor)
        if pd.isna(candidate_stock_out_date):
            candidate_stock_out_date = stock_out_date
        test2 = (pd.isna(row["Actual Finish Parsed"]) and row["Physical % Complete"] < 100) or historical_delay_flag
        test3 = (
            candidate_stock_out_date is not pd.NaT
            and required_date.normalize() >= candidate_stock_out_date.normalize()
            and required_date.normalize() <= candidate_stock_out_date.normalize() + pd.Timedelta(days=7)
        ) or (historical_delay_flag and pd.notna(client_supply_receipt_date) and pd.notna(planned_start_date) and client_supply_receipt_date.normalize() > planned_start_date.normalize())
        balance_match = balance_df[balance_df["Date"] == required_date.normalize()] if not balance_df.empty else pd.DataFrame()
        current_balance = float(balance_match["Steel Balance"].min()) if not balance_match.empty else 0.0
        activity_gap = float(row.get("Client Supply vs Actual Gap", 0.0))
        activity_available_qty = float(row.get("Client Available Qty", 0.0))
        receipt_after_need = historical_delay_flag and pd.notna(client_supply_receipt_date) and ((pd.notna(planned_start_date) and client_supply_receipt_date.normalize() > planned_start_date.normalize()) or client_supply_receipt_date.normalize() > required_date.normalize())
        test4 = current_balance <= 0 or activity_gap < 0 or receipt_after_need
        test5 = _critical_flag(row["Critical"]) or _longest_path_flag(row["Longest Path"]) or row["Total Float"] <= float(settings.near_critical_float_threshold)
        successors = str(row.get("Successors", "")).strip()
        test6 = bool(successors) or any(k in str(row["Activity Name"]).lower() for k in ["column", "slab", "beam", "wall", "raft", "next floor", "milestone", "core"])
        shortage_qty = max(abs(current_balance) if current_balance < 0 else 0.0, abs(activity_gap) if activity_gap < 0 else 0.0)
        if shortage_qty <= 0 and receipt_after_need:
            shortage_qty = required_qty
        available_qty = max(activity_available_qty, current_balance, 0.0)

        score = (
            (20 if test1 else 0) +
            (15 if test2 else 0) +
            (20 if test3 else 0) +
            (20 if test4 else 0) +
            (15 if test5 else 0) +
            (10 if test6 else 0)
        )
        classification = "Strong TIA Candidate" if score >= 85 else ("Potential TIA Candidate" if score >= 65 else "Weak / Non-driving Candidate")
        candidates.append(
            {
                "Stock-Out Date": candidate_stock_out_date,
                "Steel Type": row["Steel Type / Diameter"],
                "Building": row["Building"],
                "Activity ID": row["Activity ID"],
                "Activity Name": row["Activity Name"],
                "Required Qty": required_qty,
                "Available Qty": available_qty,
                "Shortage Qty": shortage_qty,
                "Total Float": row["Total Float"],
                "Longest Path": "Yes" if _longest_path_flag(row["Longest Path"]) else "No",
                "Critical": "Yes" if _critical_flag(row["Critical"]) else "No",
                "TIA Candidate Score": score,
                "Candidate Classification": classification,
                "Due / Ready Test": "Pass" if test3 else "Fail",
                "Stock Unavailable Test": "Pass" if test4 else "Fail",
                "Not Completed Test": "Pass" if test2 else "Fail",
                "Downstream Impact Test": "Pass" if test6 else "Fail",
                "Civil / Structural Work Package": row["Civil / Structural Work Package"],
                "Construction Sequence Impact": row["Construction Sequence Impact"],
                "Predecessors": row.get("Predecessors", ""),
                "Successors": successors,
                "Required Date": required_date,
                "Selection Explanation": row.get("Selection Explanation", ""),
                "Client Supply vs Actual Gap": activity_gap,
                "Client Available Qty": activity_available_qty,
                "Client Supply Receipt Date": client_supply_receipt_date,
                "Planned Start Date": planned_start_date,
                "Historical Employer Delay Flag": historical_delay_flag,
                "Affected Activity Explanation": "; ".join(
                    [
                        f"steel required {required_qty:.2f}",
                        "not completed" if test2 else "completion status weak",
                        "ready / due to start around stock-out" if test3 else "not due at stock-out window",
                        "stock unavailable on required date" if test4 else "stock not proven unavailable on required date",
                        f"activity-level supply gap {activity_gap:.2f}",
                        "critical / longest path / near-critical" if test5 else "not critical / near-critical",
                        "downstream impact exists" if test6 else "limited downstream impact",
                    ]
                ),
            }
        )

    result = pd.DataFrame(candidates)
    if result.empty:
        return result
    return result.sort_values(["TIA Candidate Score", "Required Date"], ascending=[False, True]).reset_index(drop=True)


def build_fragnet_recommendation(candidates_df: pd.DataFrame, balance_df: pd.DataFrame, settings: SteelTiaSettings) -> pd.DataFrame:
    if candidates_df.empty:
        return pd.DataFrame()

    qualified = candidates_df[candidates_df["Candidate Classification"].isin(["Strong TIA Candidate", "Potential TIA Candidate"])].copy()
    if qualified.empty:
        qualified = candidates_df.copy()

    balance_summary = balance_df.sort_values("Date").groupby("Date", as_index=False)["Steel Balance"].min() if not balance_df.empty else pd.DataFrame(columns=["Date", "Steel Balance"])
    rows = []
    for _, candidate in qualified.iterrows():
        fragment_start = pd.to_datetime(candidate["Stock-Out Date"], errors="coerce")
        if pd.isna(fragment_start):
            continue

        client_receipt = pd.to_datetime(candidate.get("Client Supply Receipt Date"), errors="coerce")
        if pd.notna(client_receipt) and client_receipt >= fragment_start:
            fragment_finish = client_receipt + pd.Timedelta(days=int(settings.usability_lag_days))
            open_delay = False
        else:
            later_positive = balance_summary[(balance_summary["Date"] > fragment_start) & (balance_summary["Steel Balance"] > 0)] if not balance_summary.empty else pd.DataFrame()
            fragment_finish = later_positive.iloc[0]["Date"] if not later_positive.empty else pd.Timestamp(settings.data_date or fragment_start)
            open_delay = later_positive.empty

        fragment_duration = max((fragment_finish - fragment_start).days, 0)
        predecessor_text = str(candidate.get("Predecessors", "")).split(",")[0].strip() if str(candidate.get("Predecessors", "")).strip() else ""
        successor_text = str(candidate.get("Successors", "")).split(",")[0].strip() if str(candidate.get("Successors", "")).strip() else ""
        recommended_logic = "Finish-to-finish (FF) with lag" if predecessor_text and successor_text else "Finish-to-start (FS) placeholder pending P6 logic review"
        insertion_sequence = "Stepped insertion by chronological stock-out window" if len(qualified) > 1 else "Global insertion acceptable for single event"
        rows.append(
            {
                "Fragment Activity Name": f"Client Steel Shortage - {candidate['Building']} - {candidate['Steel Type']} - Before {candidate['Activity ID']}",
                "Last completed / available predecessor": predecessor_text,
                "Affected Activity": f"{candidate['Activity ID']} - {candidate['Activity Name']}",
                "Insert Fragment Before": candidate["Activity ID"],
                "Fragment Start": fragment_start,
                "Fragment Finish": fragment_finish,
                "Fragment Duration": fragment_duration,
                "TIA Window Basis": "Statused update impact window from stock-out date to recovery date",
                "Blindsight Basis": "Use only information known at the start of this window; later events require a separate window",
                "Recommended Logic Tie": recommended_logic,
                "Recommended Lag Days": int(settings.usability_lag_days),
                "Insertion Sequence": insertion_sequence,
                "Recovery Point": "Steel balance returns positive or client receipt plus usability lag",
                "Embedded Contractor Delay Review": "Check readiness, predecessor completion, submittals, and site constraints before entitlement",
                "Concurrency / Compensability Note": "EOT may be supportable if critical; compensation requires no concurrent contractor-caused delay",
                "Traceability Assumptions": "Do not use simple addition or double counting; retain P6, steel register, notice, and relationship evidence",
                "Delay Status": "Open Delay - Steel Not Yet Fully Recovered" if open_delay else "Recovered Delay Window",
                "Affected Building": candidate["Building"],
                "Steel Type": candidate["Steel Type"],
                "Shortage Qty": candidate["Shortage Qty"],
                "TIA Candidate Score": candidate["TIA Candidate Score"],
                "Candidate Classification": candidate["Candidate Classification"],
            }
        )

    result = pd.DataFrame(rows)
    if result.empty:
        return result
    return result.sort_values(["Fragment Start", "Affected Building", "Insert Fragment Before"]).reset_index(drop=True)


def build_contract_matches(candidates_df: pd.DataFrame, contract_df: pd.DataFrame, delay_events_df: pd.DataFrame) -> pd.DataFrame:
    if candidates_df.empty or contract_df.empty:
        return pd.DataFrame()
    top_event = candidates_df.iloc[0]
    event_text = " ".join(
        [
            str(top_event.get("Activity Name", "")),
            str(top_event.get("Steel Type", "")),
            "client free issue steel delay reinforcement rft eot notice time bar mitigation schedule cost",
        ]
    ).lower()

    rows = []
    for _, clause in contract_df.iterrows():
        searchable = " ".join(
            [
                str(clause.get("Clause / Topic", "")),
                str(clause.get("Location", "")),
                str(clause.get("Plain English Meaning", "")),
                str(clause.get("Research the Lines", "")),
                str(clause.get("Money Impact", "")),
                str(clause.get("Schedule Impact", "")),
                str(clause.get("Practical Action / Evidence", "")),
            ]
        ).lower()
        topic_score = sum(keyword_match_score(searchable, keywords) for keywords in CONTRACT_TOPIC_KEYWORDS.values())
        steel_score = keyword_match_score(searchable, ["steel", "free issue", "employer supplied", "client supplied", "reinforcement", "rft"])
        steel_supply_score, ranking_rationale = score_steel_supply_clause(searchable, event_text)
        event_score = keyword_match_score(event_text, searchable.split())
        total_match = topic_score + (steel_score * 2) + steel_supply_score + (5 if event_score else 0)
        if total_match <= 0:
            continue
        rows.append(
            {
                **clause.to_dict(),
                "Matched Delay Event": top_event["Activity Name"],
                "Matched Activity ID": top_event["Activity ID"],
                "Keyword Match Score": total_match,
                "Steel Supply Clause Score": steel_supply_score,
                "Ranking Rationale": ranking_rationale,
            }
        )
    result = pd.DataFrame(rows)
    if result.empty:
        return result
    return result.sort_values(["Keyword Match Score", "Steel Supply Clause Score"], ascending=[False, False]).reset_index(drop=True)


def build_contract_support_scoring(matches_df: pd.DataFrame, top_candidate: pd.Series | None, delay_events_df: pd.DataFrame) -> pd.DataFrame:
    if matches_df.empty or top_candidate is None:
        return pd.DataFrame()
    row = matches_df.iloc[0]
    notice_ref = ""
    if not delay_events_df.empty:
        target_activity = str(top_candidate["Activity ID"]).strip().upper()
        normalized_columns = {normalize_name(col): col for col in delay_events_df.columns}
        activity_col = normalized_columns.get("activityid")
        notice_col = normalized_columns.get("noticeref")
        if activity_col:
            activity_tokens = delay_events_df[activity_col].astype(str).str.upper().apply(
                lambda value: [item.strip() for item in re.split(r"[,;|]", value) if item.strip()]
            )
            steel_delay_rows = delay_events_df[activity_tokens.apply(lambda items: target_activity in items)]
            if not steel_delay_rows.empty and notice_col:
                notice_ref = str(steel_delay_rows.iloc[0].get(notice_col, "")).strip()
    score_lines = [
        ("Relevant Clause / Topic matched", 20 if str(row.get("Clause / Topic", "")).strip() else 0),
        ("Location / contract reference available", 10 if str(row.get("Location", "")).strip() else 0),
        ("Plain English Meaning supports Contractor position", 15 if str(row.get("Plain English Meaning", "")).strip() else 0),
        ("Research the Lines supports the delay logic", 15 if str(row.get("Research the Lines", "")).strip() else 0),
        ("Who Holds Leverage favors Contractor or depends on evidence", 10 if any(x in str(row.get("Who Holds Leverage", "")).lower() for x in ["contractor", "depends", "shared"]) else 0),
        ("Notice / Time Bar is identified and notice appears compliant", 10 if str(row.get("Notice / Time Bar", "")).strip() and notice_ref else 0),
        ("Money Impact supports cost entitlement", 5 if str(row.get("Money Impact", "")).strip() else 0),
        ("Schedule Impact supports EOT / TIA / critical path requirement", 10 if str(row.get("Schedule Impact", "")).strip() else 0),
        ("Practical Action / Evidence is available and evidence is linked", 5 if str(row.get("Practical Action / Evidence", "")).strip() else 0),
    ]
    total = sum(score for _, score in score_lines)
    if total >= 85:
        strength = "Strong Contract-Supported Delay Position"
    elif total >= 65:
        strength = "Moderate Contract Support - Needs Additional Evidence"
    elif total >= 40:
        strength = "Weak Contract Support - High Substantiation Risk"
    else:
        strength = "Not Contractually Proven"
    return pd.DataFrame(
        [{"Score Line": label, "Points": score} for label, score in score_lines]
    ).assign(**{"Contract Support Score": total, "Contractual Strength": strength, "Notice Reference": notice_ref or "Not linked"})


def classify_delay_position(top_candidate: pd.Series | None, support_strength: str) -> str:
    if top_candidate is None:
        return "Not Proven"
    ready = top_candidate.get("Due / Ready Test") == "Pass"
    stock = top_candidate.get("Stock Unavailable Test") == "Pass"
    critical = str(top_candidate.get("Critical", "")).lower() == "yes" or str(top_candidate.get("Longest Path", "")).lower() == "yes" or float(top_candidate.get("Total Float", 0)) <= 10
    if ready and stock and critical and support_strength.startswith("Strong"):
        return "Employer / Client Risk"
    if ready and stock and support_strength.startswith("Moderate"):
        return "Concurrent Risk"
    if not ready:
        return "Concurrent / Non-Driving Risk"
    return "Not Proven"


def build_contractual_assessment(candidates_df: pd.DataFrame, fragnet_df: pd.DataFrame, contract_matches_df: pd.DataFrame, support_df: pd.DataFrame, delay_events_df: pd.DataFrame) -> pd.DataFrame:
    if candidates_df.empty:
        return pd.DataFrame()
    top = candidates_df.iloc[0]
    matching_fragnet = pd.DataFrame()
    if not fragnet_df.empty and "Insert Fragment Before" in fragnet_df.columns:
        matching_fragnet = fragnet_df[fragnet_df["Insert Fragment Before"].astype(str).str.strip() == str(top["Activity ID"]).strip()]
    fragnet_row = matching_fragnet.iloc[0] if not matching_fragnet.empty else (fragnet_df.iloc[0] if not fragnet_df.empty else pd.Series(dtype=object))
    support_strength = support_df["Contractual Strength"].iloc[0] if not support_df.empty else "Not Contractually Proven"
    delay_classification = classify_delay_position(top, support_strength)
    notice_ref = support_df["Notice Reference"].iloc[0] if not support_df.empty else "Not linked"
    location = contract_matches_df["Location"].iloc[0] if not contract_matches_df.empty else ""
    practical = contract_matches_df["Practical Action / Evidence"].iloc[0] if not contract_matches_df.empty else ""
    return pd.DataFrame(
        [
            {
                "Stock-Out Date": top["Stock-Out Date"],
                "Steel Type": top["Steel Type"],
                "Building": top["Building"],
                "Activity ID": top["Activity ID"],
                "Activity Name": top["Activity Name"],
                "Required Qty": top["Required Qty"],
                "Available Qty": top["Available Qty"],
                "Shortage Qty": top["Shortage Qty"],
                "Total Float": top["Total Float"],
                "Longest Path": top["Longest Path"],
                "Critical": top["Critical"],
                "Fragment Start": fragnet_row.get("Fragment Start", pd.NaT),
                "Fragment Finish": fragnet_row.get("Fragment Finish", pd.NaT),
                "Fragment Duration": fragnet_row.get("Fragment Duration", 0),
                "Insert Fragment Before": top["Activity ID"],
                "TIA Candidate Score": top["TIA Candidate Score"],
                "Delay Classification": delay_classification,
                "Notice Reference": notice_ref,
                "Evidence Reference": location,
                "User Comment": "",
                "Final Assessment": support_strength,
                "Money Impact Assessment": contract_matches_df["Money Impact"].iloc[0] if not contract_matches_df.empty else "",
                "Schedule Impact and EOT Assessment": contract_matches_df["Schedule Impact"].iloc[0] if not contract_matches_df.empty else "",
                "Practical Action / Evidence": practical,
            }
        ]
    )


def build_rebuttal_matrix(assessment_df: pd.DataFrame, contract_matches_df: pd.DataFrame) -> pd.DataFrame:
    if assessment_df.empty:
        return pd.DataFrame()
    row = assessment_df.iloc[0]
    clause = contract_matches_df.iloc[0] if not contract_matches_df.empty else pd.Series(dtype=object)
    rebuttals = [
        "The activity was not ready.",
        "The activity was not critical.",
        "The Contractor failed to mitigate.",
        "Steel was delivered but not properly managed by Contractor.",
        "The delay was concurrent.",
    ]
    return pd.DataFrame(
        [
            {
                "Delay Event": "Client / Employer free-issue steel shortage",
                "Affected Activity": row["Activity Name"],
                "Contractor Position": row["Delay Classification"],
                "Contract Support from Clause / Topic": clause.get("Clause / Topic", ""),
                "Location": clause.get("Location", ""),
                "Engineer / Employer Likely Rebuttal": rebuttal,
                "Counterargument Using Plain English Meaning": clause.get("Plain English Meaning", ""),
                "Counterargument Using Research the Lines": clause.get("Research the Lines", ""),
                "Leverage Position": clause.get("Who Holds Leverage", ""),
                "Notice / Time Bar Risk": clause.get("Notice / Time Bar", ""),
                "Money Impact Risk": clause.get("Money Impact", ""),
                "Schedule Impact Risk": clause.get("Schedule Impact", ""),
                "Practical Action / Evidence Required": clause.get("Practical Action / Evidence", ""),
                "Final Risk Rating": "High" if "not ready" in rebuttal.lower() or "concurrent" in rebuttal.lower() else "Medium",
            }
            for rebuttal in rebuttals
        ]
    )


def build_action_tracker(assessment_df: pd.DataFrame, contract_matches_df: pd.DataFrame) -> pd.DataFrame:
    if assessment_df.empty:
        return pd.DataFrame()
    clause = contract_matches_df.iloc[0] if not contract_matches_df.empty else pd.Series(dtype=object)
    row = assessment_df.iloc[0]
    actions = [item.strip() for item in re.split(r"[.;]\s*", str(clause.get("Practical Action / Evidence", ""))) if item.strip()]
    if not actions:
        actions = ["Issue protective notice and maintain stock, requisition, and critical path records."]
    return pd.DataFrame(
        [
            {
                "Delay Event": "Client / Employer free-issue steel shortage",
                "Activity ID": row["Activity ID"],
                "Clause / Topic": clause.get("Clause / Topic", ""),
                "Location": clause.get("Location", ""),
                "Required Action": action,
                "Required Evidence": clause.get("Research the Lines", ""),
                "Responsible Party": "Planning / Commercial / Site Team",
                "Due Date": row["Fragment Start"],
                "Status": "Open",
                "User Comment": "",
            }
            for action in actions
        ]
    )


def build_professional_narrative(assessment_df: pd.DataFrame, contract_matches_df: pd.DataFrame, support_df: pd.DataFrame) -> pd.DataFrame:
    if assessment_df.empty:
        return pd.DataFrame()
    row = assessment_df.iloc[0]
    clause = contract_matches_df.iloc[0] if not contract_matches_df.empty else pd.Series(dtype=object)
    support_strength = support_df["Contractual Strength"].iloc[0] if not support_df.empty else "Not Contractually Proven"
    narrative = (
        f"Based on the uploaded P6 schedule, steel delivery register, and contract library, the analysis indicates that the Client / Employer supplied reinforcement steel became insufficient on {pd.to_datetime(row['Stock-Out Date']).strftime('%d-%b-%Y')} "
        f"for {row['Building']} / {row['Steel Type']}. The first affected activity identified by the logic engine is {row['Activity ID']} - {row['Activity Name']}, "
        f"which relates to {row.get('Activity Name', '')} and forms part of the structural sequence leading to downstream successor activities.\n\n"
        f"The TIA should be inserted in the statused schedule update window covering {pd.to_datetime(row['Fragment Start']).strftime('%d-%b-%Y')} to {pd.to_datetime(row['Fragment Finish']).strftime('%d-%b-%Y')}, "
        f"using a fragnet before the affected activity rather than a simple addition of all alleged delay days. The recommended scheduling check is a stepped insertion by chronological window where multiple stock-outs exist; "
        f"a global insertion may be used only as a sensitivity run. The logic review should test finish-to-finish links with the required lag where the affected work was already in progress, and should retain the predecessor, successor, recovery point, and float movement created by the inserted fragnet.\n\n"
        f"Blindsight should be controlled by using only information known at the start of the analysed window. Embedded contractor-caused delay must be reviewed before final entitlement, including readiness, predecessor completion, approved submittals, access, manpower, and mitigation records. "
        f"The result can support EOT where the event delays the critical path or consumes available float; compensation should remain subject to a separate concurrency review because concurrent contractor-caused delay can defeat or reduce compensability even when time entitlement remains arguable.\n\n"
        f"The contract library row under {clause.get('Clause / Topic', 'the matched clause')}, located at {clause.get('Location', 'the contract reference')}, indicates that {clause.get('Plain English Meaning', '')}. "
        f"The related contractual interpretation states {clause.get('Research the Lines', '')}. The leverage assessment is {clause.get('Who Holds Leverage', '')}.\n\n"
        f"Subject to notice compliance under {clause.get('Notice / Time Bar', '')}, and subject to the availability of the practical evidence listed under {clause.get('Practical Action / Evidence', '')}, "
        f"this event is assessed as {support_strength} for schedule entitlement and {clause.get('Money Impact', '')} for money impact."
    )
    return pd.DataFrame(
        [
            {
                "Event Description": "Client / Employer free-issue reinforcement steel shortage",
                "Civil / Structural Work Package Affected": row["Activity Name"],
                "Construction Sequence Impact": "Steel shortage -> RFT blocked -> inspection/casting sequence delayed -> successors exposed",
                "Schedule Impact": row["Schedule Impact and EOT Assessment"],
                "Contract Library Support": clause.get("Clause / Topic", ""),
                "Notice / Time-Bar Review": clause.get("Notice / Time Bar", ""),
                "Money Impact": row["Money Impact Assessment"],
                "Required Evidence": clause.get("Practical Action / Evidence", ""),
                "Employer / Engineer Possible Rebuttal": "Activity not ready, not critical, concurrent constraints, or lack of records.",
                "Contractor Response": clause.get("Research the Lines", ""),
                "Final Professional Opinion": narrative,
            }
        ]
    )


def build_validation_report(mapped_df: pd.DataFrame, missing_fields: list[str], duplicate_ids: list[str], dataset_name: str) -> pd.DataFrame:
    issues = []
    if missing_fields:
        issues.append({"Dataset": dataset_name, "Issue Type": "Missing Fields", "Detail": ", ".join(missing_fields), "Severity": "Warning"})
    if duplicate_ids:
        issues.append({"Dataset": dataset_name, "Issue Type": "Duplicated Activity IDs", "Detail": ", ".join(duplicate_ids[:20]), "Severity": "Critical"})
    if mapped_df.empty:
        issues.append({"Dataset": dataset_name, "Issue Type": "No Data", "Detail": "The table is empty after mapping.", "Severity": "Critical"})
    return pd.DataFrame(issues)


def run_steel_delay_tia_analysis(
    p6_df: pd.DataFrame,
    steel_df: pd.DataFrame,
    requirement_df: pd.DataFrame,
    relationship_df: pd.DataFrame,
    contract_library_df: pd.DataFrame,
    delay_events_df: pd.DataFrame,
    settings: SteelTiaSettings,
) -> dict[str, pd.DataFrame | dict[str, Any]]:
    if "Required Steel Qty" not in requirement_df.columns and not requirement_df.empty:
        requirement_df, _, _ = apply_mapping(requirement_df, suggest_mapping(requirement_df, REQ_ALIASES), CANONICAL_REQ_FIELDS)
    if "Date" not in steel_df.columns and not steel_df.empty:
        steel_df, _, _ = apply_mapping(steel_df, suggest_mapping(steel_df, STEEL_ALIASES), CANONICAL_STEEL_FIELDS)
    if "Activity ID" not in p6_df.columns and not p6_df.empty:
        p6_df, _, _ = apply_mapping(p6_df, suggest_mapping(p6_df, P6_ALIASES), CANONICAL_P6_FIELDS)

    requirement_df = requirement_df.copy()
    auto_requirements_used = False
    if requirement_df.empty:
        requirement_df = build_auto_requirement_df(p6_df)
        auto_requirements_used = True
    quantity_basis_valid = False
    if not requirement_df.empty:
        req_check = requirement_df.copy()
        req_check["Required Steel Qty"] = safe_numeric(req_check.get("Required Steel Qty", 0))
        req_check["Required Date"] = pd.to_datetime(req_check.get("Required Date", pd.Series(dtype=object)).apply(parse_mixed_date), errors="coerce")
        quantity_basis_valid = bool(
            (req_check["Required Steel Qty"] > 0).any()
            and req_check["Required Date"].notna().any()
        )

    if "Building" in p6_df.columns:
        building_series = p6_df["Building"]
    else:
        building_series = pd.Series([""] * len(p6_df), index=p6_df.index)
    p6_df["Building"] = building_series.replace("", pd.NA).fillna(p6_df["Activity ID"].apply(standardize_building))
    if "Steel Type / Diameter" not in p6_df.columns:
        p6_df["Steel Type / Diameter"] = ""
    p6_df["Steel Type / Diameter"] = p6_df.apply(
        lambda row: infer_steel_type_from_context(
            row.get("Steel Type / Diameter", ""),
            row.get("Activity Code", ""),
            row.get("Activity Name", ""),
            row.get("WBS", ""),
        ),
        axis=1,
    )
    if not requirement_df.empty:
        req_lookup = requirement_df.copy()
        req_lookup["Activity ID"] = req_lookup["Activity ID"].astype(str).str.strip().str.upper()
        for optional_col in [
            "Building",
            "Steel Type / Diameter",
            "Required Steel Qty",
            "Required Date",
            "Priority",
            "Remarks",
            "Requirement Basis",
            "Selection Explanation",
            "Steel Logic Match",
            "Client Available Qty",
            "Client Supply vs Actual Gap",
            "Client Supply Receipt Date",
            "Drive Activity Dates",
            "Activity Units % Complete",
            "Reason for Delay",
            "Responsibility",
        ]:
            if optional_col not in req_lookup.columns:
                req_lookup[optional_col] = ""
        req_lookup = req_lookup.drop_duplicates(subset=["Activity ID"], keep="first")
        p6_merge = p6_df.merge(
            req_lookup[
                [
                    "Activity ID",
                    "Building",
                    "Steel Type / Diameter",
                    "Required Steel Qty",
                    "Required Date",
                    "Priority",
                    "Remarks",
                    "Requirement Basis",
                    "Selection Explanation",
                    "Steel Logic Match",
                    "Client Available Qty",
                    "Client Supply vs Actual Gap",
                    "Client Supply Receipt Date",
                    "Planned Start Date",
                    "BL Start Date",
                    "Drive Activity Dates",
                    "Activity Units % Complete",
                    "Reason for Delay",
                    "Responsibility",
                    "Historical Employer Delay Flag",
                ]
            ],
            on="Activity ID",
            how="left",
            suffixes=("", "_req"),
        )
        for merged_col in [
            "Building",
            "Steel Type / Diameter",
            "Required Steel Qty",
            "Required Date",
            "Priority",
            "Remarks",
            "Requirement Basis",
            "Selection Explanation",
            "Steel Logic Match",
            "Client Available Qty",
            "Client Supply vs Actual Gap",
            "Client Supply Receipt Date",
            "Planned Start Date",
            "BL Start Date",
            "Drive Activity Dates",
            "Activity Units % Complete",
            "Reason for Delay",
            "Responsibility",
            "Historical Employer Delay Flag",
        ]:
            req_col = f"{merged_col}_req"
            if req_col in p6_merge.columns:
                if merged_col not in p6_merge.columns:
                    p6_merge[merged_col] = p6_merge[req_col]
                else:
                    p6_merge[merged_col] = p6_merge[merged_col].replace("", pd.NA).fillna(p6_merge[req_col])
                p6_merge = p6_merge.drop(columns=[req_col])
        p6_df = p6_merge
        if "Longest Path" in p6_df.columns and "Drive Activity Dates" in p6_df.columns:
            p6_df["Longest Path"] = p6_df["Longest Path"].replace("", pd.NA).fillna(p6_df["Drive Activity Dates"])
    p6_df["Required Steel Qty"] = safe_numeric(p6_df.get("Required Steel Qty", p6_df.get("Remaining Units", 0)))
    p6_df["Required Date"] = p6_df.get("Required Date", p6_df.get("Start", p6_df.get("Baseline Start", "")))

    if quantity_basis_valid:
        balance_df = build_daily_balance(requirement_df, steel_df, settings)
        stock_out_df = build_stock_out_events(balance_df)
        if stock_out_df.empty:
            stock_out_df = build_stock_out_events_from_activity_supply(requirement_df)
        candidates_df = build_affected_candidates(p6_df, balance_df, settings, stock_out_df)
        fragnet_df = build_fragnet_recommendation(candidates_df, balance_df, settings)
        contract_matches_df = build_contract_matches(candidates_df, contract_library_df, delay_events_df)
        support_df = build_contract_support_scoring(contract_matches_df, candidates_df.iloc[0] if not candidates_df.empty else None, delay_events_df)
        assessment_df = build_contractual_assessment(candidates_df, fragnet_df, contract_matches_df, support_df, delay_events_df)
        rebuttal_df = build_rebuttal_matrix(assessment_df, contract_matches_df)
        action_tracker_df = build_action_tracker(assessment_df, contract_matches_df)
        narrative_df = build_professional_narrative(assessment_df, contract_matches_df, support_df)
    else:
        balance_df = pd.DataFrame()
        stock_out_df = pd.DataFrame()
        candidates_df = pd.DataFrame()
        fragnet_df = pd.DataFrame()
        contract_matches_df = pd.DataFrame()
        support_df = pd.DataFrame()
        assessment_df = pd.DataFrame()
        rebuttal_df = pd.DataFrame()
        action_tracker_df = pd.DataFrame()
        narrative_df = pd.DataFrame()

    current_balance = float(balance_df["Steel Balance"].iloc[-1]) if not balance_df.empty else 0.0
    total_delivered = float(balance_df["Client Delivered Qty"].sum()) if not balance_df.empty else 0.0
    total_required = float(balance_df["Required Steel Qty"].sum()) if not balance_df.empty else 0.0
    total_shortage = float(stock_out_df["Quantity Shortage"].sum()) if not stock_out_df.empty else 0.0
    open_delay_duration = int(assessment_df["Fragment Duration"].iloc[0]) if not assessment_df.empty else 0
    employer_risk_events = int(assessment_df["Delay Classification"].astype(str).str.contains("Employer", case=False).sum()) if not assessment_df.empty else 0
    concurrent_risk_events = int(assessment_df["Delay Classification"].astype(str).str.contains("Concurrent", case=False).sum()) if not assessment_df.empty else 0

    executive_rows = []
    if not quantity_basis_valid:
        executive_rows.append(
            {
                "Question": "Why is the analysis blocked?",
                "Answer": "No valid steel quantity basis is available. Upload a requirement file or a P6 export with usable quantity fields before stock-out and fragnet analysis can run.",
            }
        )
    executive_rows.extend(
        [
            {"Question": "When did steel quantity become zero?", "Answer": stock_out_df["Stock-Out Date"].min() if not stock_out_df.empty else pd.NaT},
            {"Question": "Which steel type / diameter was affected?", "Answer": assessment_df["Steel Type"].iloc[0] if not assessment_df.empty else "Not identified"},
            {"Question": "Which building was affected?", "Answer": assessment_df["Building"].iloc[0] if not assessment_df.empty else "Not identified"},
            {"Question": "Which activity was first impacted?", "Answer": assessment_df["Activity ID"].iloc[0] + " - " + assessment_df["Activity Name"].iloc[0] if not assessment_df.empty else "Not identified"},
            {"Question": "Where should the TIA fragnet be inserted?", "Answer": fragnet_df["Insert Fragment Before"].iloc[0] if not fragnet_df.empty else "Not identified"},
            {"Question": "What is the fragment start date?", "Answer": fragnet_df["Fragment Start"].iloc[0] if not fragnet_df.empty else pd.NaT},
            {"Question": "What is the fragment finish date?", "Answer": fragnet_df["Fragment Finish"].iloc[0] if not fragnet_df.empty else pd.NaT},
            {"Question": "What is the fragment duration?", "Answer": fragnet_df["Fragment Duration"].iloc[0] if not fragnet_df.empty else 0},
            {"Question": "Is the delay critical, near-critical, concurrent, or non-driving?", "Answer": assessment_df["Delay Classification"].iloc[0] if not assessment_df.empty else "Not identified"},
            {"Question": "Is the delay likely Employer / Client risk or Contractor risk?", "Answer": assessment_df["Final Assessment"].iloc[0] if not assessment_df.empty else "Not identified"},
        ]
    )
    executive_summary = pd.DataFrame(executive_rows)

    data_quality_rows = []
    if not quantity_basis_valid:
        data_quality_rows.append(
            {
                "Dataset": "Activity Steel Requirement",
                "Issue Type": "No Derivable Quantities",
                "Detail": "No valid steel quantity basis is available. Upload a requirement file or a P6 export containing usable quantity fields before stock-out and fragnet analysis can run.",
                "Severity": "Critical",
            }
        )
    elif auto_requirements_used:
        no_qty_count = int((requirement_df["Required Steel Qty"] <= 0).sum()) if not requirement_df.empty else 0
        data_quality_rows.append(
            {
                "Dataset": "Activity Steel Requirement",
                "Issue Type": "Auto-Derived Requirements",
                "Detail": f"{len(requirement_df)} activity requirement rows were derived automatically from the uploaded / loaded project records.",
                "Severity": "Info",
            }
        )
        if no_qty_count:
            data_quality_rows.append(
                {
                    "Dataset": "Activity Steel Requirement",
                    "Issue Type": "Zero Quantity Rows",
                    "Detail": f"{no_qty_count} derived rows have zero quantity and may need explicit quantity support from P6 units or a separate requirement file.",
                    "Severity": "Warning",
                }
            )
    data_quality_df = pd.DataFrame(data_quality_rows)

    notice_review_df = assessment_df[["Activity ID", "Notice Reference", "Delay Classification"]].rename(columns={"Delay Classification": "Notice Position"}) if not assessment_df.empty else pd.DataFrame()
    money_impact_df = assessment_df[["Activity ID", "Money Impact Assessment", "Shortage Qty"]].copy() if not assessment_df.empty else pd.DataFrame()
    schedule_impact_df = assessment_df[["Activity ID", "Schedule Impact and EOT Assessment", "Fragment Duration", "Total Float"]].copy() if not assessment_df.empty else pd.DataFrame()

    return {
        "balance_df": balance_df,
        "stock_out_df": stock_out_df,
        "candidates_df": candidates_df,
        "fragnet_df": fragnet_df,
        "contract_matches_df": contract_matches_df,
        "support_df": support_df,
        "assessment_df": assessment_df,
        "rebuttal_df": rebuttal_df,
        "action_tracker_df": action_tracker_df,
        "narrative_df": narrative_df,
        "executive_summary_df": executive_summary,
        "notice_review_df": notice_review_df,
        "money_impact_df": money_impact_df,
        "schedule_impact_df": schedule_impact_df,
        "data_quality_df": data_quality_df,
        "kpis": {
            "Total Delivered Steel": total_delivered,
            "Total Required Steel": total_required,
            "Current Steel Balance": current_balance,
            "Total Shortage Qty": total_shortage,
            "First Stock-Out Date": stock_out_df["Stock-Out Date"].min() if not stock_out_df.empty else pd.NaT,
            "Number of Stock-Out Events": int(len(stock_out_df)),
            "Number of Strong TIA Candidates": int(candidates_df["Candidate Classification"].eq("Strong TIA Candidate").sum()) if not candidates_df.empty else 0,
            "Critical Activities Affected": int(candidates_df["Critical"].eq("Yes").sum()) if not candidates_df.empty else 0,
            "Longest Path Activities Affected": int(candidates_df["Longest Path"].eq("Yes").sum()) if not candidates_df.empty else 0,
            "Open Delay Duration": open_delay_duration,
            "Employer Risk Events": employer_risk_events,
            "Concurrent Risk Events": concurrent_risk_events,
            "Auto Requirements Used": auto_requirements_used,
            "Quantity Basis Valid": quantity_basis_valid,
        },
        "requirement_df": requirement_df,
    }
