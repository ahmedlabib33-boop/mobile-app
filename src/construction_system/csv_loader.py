"""CSV Data Loader - Load all dashboard data from CSV files."""

import pandas as pd
import re
from pathlib import Path
from typing import Any
from typing import Dict, Optional

from .project_catalog import discover_projects, project_data_path, projects_frame


APP_DIR = Path(__file__).parent.parent.parent
PROJECTS_DIR = APP_DIR / "projects"

CSV_FILES = {
    "projects": "projects.csv",
    "activities": "activities.csv",
    "contracts": "contracts.csv",
    "cost_items": "cost_items.csv",
    "delay_events": "delay_events.csv",
    "risks": "risks.csv",
    "milestones": "milestones.csv",
    "payments": "payments.csv",
    "progress_updates": "progress_updates.csv",
    "change_orders": "change_orders.csv",
    "claims": "claims.csv",
    "wbs": "wbs.csv",
}


def _normalized_column_lookup(df: pd.DataFrame) -> dict[str, str]:
    return {
        re.sub(r"[^a-z0-9]+", "", str(col).strip().lower()): col
        for col in df.columns
    }


def _first_existing_column(df: pd.DataFrame, candidates: list[str]) -> str | None:
    lookup = _normalized_column_lookup(df)
    for candidate in candidates:
        if candidate in df.columns:
            return candidate
        match = lookup.get(re.sub(r"[^a-z0-9]+", "", str(candidate).strip().lower()))
        if match is not None:
            return match
    return None


def _copy_column_if_missing(df: pd.DataFrame, target: str, candidates: list[str], default: Any = "") -> None:
    if target in df.columns:
        return
    source = _first_existing_column(df, candidates)
    df[target] = df[source] if source is not None else default


def normalize_import_template_frame(file_key: str, df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    working = df.loc[:, [col for col in df.columns if str(col).strip()]].copy()
    if file_key == "payments":
        _copy_column_if_missing(working, "payment_id", ["payment", "payment id", "payment_id"])
        _copy_column_if_missing(working, "contract_id", ["contract", "contract id", "contract_id"])
        _copy_column_if_missing(working, "project_id", ["project", "project id", "project_id"])
        _copy_column_if_missing(working, "invoice_no", ["invoice no", "invoice_no", "invoice"])
        _copy_column_if_missing(working, "invoice_date", ["invoice date", "invoice_date", "payment_date"])
        _copy_column_if_missing(working, "payment_date", ["Date of Cash Cheque Receipt", "payment date", "payment_date", "invoice date"])
        _copy_column_if_missing(working, "certified_amount", ["certified amount", "certified_amount", "certified"])
        _copy_column_if_missing(working, "paid_amount", ["paid amount", "paid_amount", "paid"])
        _copy_column_if_missing(working, "payment_status", ["payment status", "payment_status", "status"])
    elif file_key == "delay_events":
        _copy_column_if_missing(working, "delay_id", ["delay_id", "delay event id", "event id", "Primary Event ID"])
        _copy_column_if_missing(working, "delay_title", ["delay_title", "event_title", "Primary Event ID", "Activity Name"])
        _copy_column_if_missing(working, "project_id", ["project_id", "project"])
        _copy_column_if_missing(working, "activity_id", ["activity_id", "Activity ID"])
        _copy_column_if_missing(working, "activity_name", ["activity_name", "Activity Name"])
        _copy_column_if_missing(working, "start_date", ["start_date", "Start", "Overlap Start", "BL Start"])
        _copy_column_if_missing(working, "end_date", ["end_date", "Finish", "Overlap Finish", "BL Finish"])
        _copy_column_if_missing(working, "estimated_delay_days", ["estimated_delay_days", "Delayed duration after overlap", "Delayed duration", "Concurrent delay"])
        _copy_column_if_missing(working, "approved_eot_days", ["approved_eot_days", "approved eot days"], 0)
        _copy_column_if_missing(working, "responsibility", ["responsibility", "responsible_party"], "Employer / Client")
        _copy_column_if_missing(working, "cause_category", ["cause_category", "Primary Event ID"], "Delay")
        _copy_column_if_missing(working, "notice_ref", ["notice_ref", "notice ref"], "")
        _copy_column_if_missing(working, "status", ["status"], "Open")
    return working


class CSVDataLoader:
    """Load and cache CSV data for dashboard."""
    
    def __init__(self, project_id: str = ""):
        """Initialize data loader with cache."""
        self.project_id = str(project_id or "").strip()
        self.cache = {}
        self.load_status = {}
    
    def load_csv(self, file_key: str, required: bool = False) -> Optional[pd.DataFrame]:
        """Load CSV file with caching.
        
        Args:
            file_key: Key in CSV_FILES dict
            required: If True, raise error if file not found
            
        Returns:
            DataFrame or None if file not found
        """
        # Return cached data if available
        if file_key in self.cache:
            return self.cache[file_key]
        
        # Get file path
        if file_key not in CSV_FILES:
            if required:
                raise ValueError(f"Unknown CSV file: {file_key}")
            return None
        
        try:
            records = discover_projects(PROJECTS_DIR)
            project_ids = [self.project_id] if self.project_id else [row["project_id"] for row in records]
            frames = []
            for project_id in project_ids:
                file_path = project_data_path(PROJECTS_DIR, project_id, "core", CSV_FILES[file_key])
                if not file_path.exists():
                    continue
                frame = normalize_import_template_frame(file_key, pd.read_csv(file_path))
                if "project_id" not in frame.columns:
                    frame.insert(0, "project_id", project_id)
                else:
                    source_ids = frame["project_id"].astype(str).str.strip()
                    mismatched = source_ids.ne("") & source_ids.ne(project_id)
                    if mismatched.any() and "source_project_id" not in frame.columns:
                        frame.insert(1, "source_project_id", source_ids)
                    frame["project_id"] = project_id
                frames.append(frame)
            df = pd.concat(frames, ignore_index=True, sort=False) if frames else pd.DataFrame()
            if df.empty and required:
                raise FileNotFoundError(f"Required CSV file not found for project scope: {file_key}")
            self.cache[file_key] = df
            self.load_status[file_key] = f"Loaded: {len(df)} rows"
            return df
        except Exception as e:
            self.load_status[file_key] = f"Error loading: {str(e)}"
            if required:
                raise
            return None
    
    def get_projects(self) -> pd.DataFrame:
        """Get projects data."""
        return projects_frame(PROJECTS_DIR)
    
    def get_activities(self) -> pd.DataFrame:
        """Get activities data."""
        df = self.load_csv("activities", required=False)
        if df is None:
            return pd.DataFrame()
        return df
    
    def get_contracts(self) -> pd.DataFrame:
        """Get contracts data."""
        df = self.load_csv("contracts", required=False)
        if df is None:
            return pd.DataFrame()
        return df
    
    def get_cost_items(self) -> pd.DataFrame:
        """Get cost items data."""
        df = self.load_csv("cost_items", required=False)
        if df is None:
            return pd.DataFrame()
        return df
    
    def get_delay_events(self) -> pd.DataFrame:
        """Get delay events data."""
        df = self.load_csv("delay_events", required=False)
        if df is None:
            return pd.DataFrame()
        return df
    
    def get_risks(self) -> pd.DataFrame:
        """Get risks data."""
        df = self.load_csv("risks", required=False)
        if df is None:
            return pd.DataFrame()
        return df
    
    def get_milestones(self) -> pd.DataFrame:
        """Get milestones data."""
        df = self.load_csv("milestones", required=False)
        if df is None:
            return pd.DataFrame()
        return df
    
    def get_payments(self) -> pd.DataFrame:
        """Get payments data."""
        df = self.load_csv("payments", required=False)
        if df is None:
            return pd.DataFrame()
        return df
    
    def get_progress_updates(self) -> pd.DataFrame:
        """Get progress updates data."""
        df = self.load_csv("progress_updates", required=False)
        if df is None:
            return pd.DataFrame()
        return df
    
    def get_change_orders(self) -> pd.DataFrame:
        """Get change orders data."""
        df = self.load_csv("change_orders", required=False)
        if df is None:
            return pd.DataFrame()
        return df
    
    def get_claims(self) -> pd.DataFrame:
        """Get claims data."""
        df = self.load_csv("claims", required=False)
        if df is None:
            return pd.DataFrame()
        return df
    
    def get_wbs(self) -> pd.DataFrame:
        """Get WBS data."""
        df = self.load_csv("wbs", required=False)
        if df is None:
            return pd.DataFrame()
        return df
    
    def get_load_status(self) -> Dict[str, str]:
        """Get status of all loaded files."""
        return self.load_status
    
    def reload_all(self):
        """Clear cache and reload all files."""
        self.cache.clear()
        self.load_status.clear()
        for key in CSV_FILES.keys():
            self.load_csv(key, required=False)


# Global loader instance
_loader = None


def get_loader() -> CSVDataLoader:
    """Get or create global loader instance."""
    global _loader
    if _loader is None:
        _loader = CSVDataLoader()
    return _loader


def reload_data():
    """Reload all CSV data."""
    get_loader().reload_all()
