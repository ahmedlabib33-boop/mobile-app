"""CSV Data Integration - Bridge between CSV files and dashboard."""

import pandas as pd
from typing import Dict, List, Optional, Tuple
from datetime import datetime
from src.construction_system.csv_loader import get_loader as get_csv_loader
from src.construction_system.evm_loader import get_evm_loader


class CSVDataIntegration:
    """Integrate CSV data for dashboard display."""
    
    def __init__(self):
        """Initialize integration layer."""
        self.csv_loader = get_csv_loader()
        self.evm_loader = get_evm_loader()
    
    def get_project_summary(self) -> Dict:
        """Get project summary from CSV data.
        
        Returns:
            Dictionary with project summary
        """
        projects = self.csv_loader.get_projects()
        
        if projects is None or len(projects) == 0:
            return {
                "project_name": "No Project Data",
                "contractor": "N/A",
                "employer": "N/A",
                "contract_value": 0,
                "start_date": None,
                "end_date": None,
            }
        
        # Get first project
        project = projects.iloc[0]
        
        return {
            "project_name": project.get("project_name", "Unknown"),
            "contractor": project.get("contractor", "N/A"),
            "employer": project.get("employer", "N/A"),
            "contract_value": float(project.get("contract_value", 0)),
            "start_date": project.get("start_date"),
            "end_date": project.get("end_date"),
        }
    
    def get_activities_data(self) -> pd.DataFrame:
        """Get activities data from CSV.
        
        Returns:
            DataFrame with activities
        """
        df = self.csv_loader.get_activities()
        if df is None or len(df) == 0:
            return pd.DataFrame()
        
        # Ensure required columns exist
        required_cols = ["activity_id", "activity_name", "percent_complete", "planned_cost", "actual_cost"]
        for col in required_cols:
            if col not in df.columns:
                df[col] = 0
        
        return df
    
    def get_evm_kpis(self) -> Dict:
        """Calculate EVM KPIs from CSV data.
        
        Returns:
            Dictionary with EVM metrics
        """
        activities = self.get_activities_data()
        
        if activities is None or len(activities) == 0:
            return {
                "bac": 0,
                "ac": 0,
                "ev": 0,
                "pv": 0,
                "cpi": None,
                "spi": None,
                "cv": 0,
                "sv": 0,
                "eac": 0,
                "etc": 0,
                "tcpi": None,
                "percent_complete": 0,
                "total_activities": 0,
                "critical_activities": 0,
            }
        
        # Calculate metrics
        bac = activities["planned_cost"].sum() if "planned_cost" in activities.columns else 0
        ac = activities["actual_cost"].sum() if "actual_cost" in activities.columns else 0
        
        # Calculate EV based on percent complete
        if "percent_complete" in activities.columns:
            ev = (activities["planned_cost"] * activities["percent_complete"] / 100).sum()
        else:
            ev = 0
        
        # Calculate PV (assume linear progress)
        pv = bac * 0.5  # Placeholder - should be based on schedule
        
        # Calculate indices
        cpi = ev / ac if ac > 0 else None
        spi = ev / pv if pv > 0 else None
        
        # Calculate variances
        cv = ev - ac
        sv = ev - pv
        
        # Calculate EAC and ETC
        eac = bac / cpi if cpi and cpi > 0 else bac
        etc = eac - ac
        
        # Calculate TCPI
        tcpi = (bac - ev) / (eac - ac) if (eac - ac) > 0 else None
        
        # Count activities
        total_activities = len(activities)
        critical_activities = len(activities[activities.get("critical_flag", "No") == "Yes"]) if "critical_flag" in activities.columns else 0
        
        # Overall progress
        avg_progress = activities["percent_complete"].mean() if "percent_complete" in activities.columns else 0
        
        return {
            "bac": bac,
            "ac": ac,
            "ev": ev,
            "pv": pv,
            "cpi": cpi,
            "spi": spi,
            "cv": cv,
            "sv": sv,
            "eac": eac,
            "etc": etc,
            "tcpi": tcpi,
            "percent_complete": avg_progress,
            "total_activities": total_activities,
            "critical_activities": critical_activities,
        }
    
    def get_contracts_data(self) -> pd.DataFrame:
        """Get contracts data from CSV.
        
        Returns:
            DataFrame with contracts
        """
        df = self.csv_loader.get_contracts()
        if df is None or len(df) == 0:
            return pd.DataFrame()
        
        return df
    
    def get_risks_data(self) -> pd.DataFrame:
        """Get risks data from CSV.
        
        Returns:
            DataFrame with risks
        """
        df = self.csv_loader.get_risks()
        if df is None or len(df) == 0:
            return pd.DataFrame()
        
        return df
    
    def get_delays_data(self) -> pd.DataFrame:
        """Get delays data from CSV.
        
        Returns:
            DataFrame with delays
        """
        df = self.csv_loader.get_delay_events()
        if df is None or len(df) == 0:
            return pd.DataFrame()
        
        return df
    
    def get_milestones_data(self) -> pd.DataFrame:
        """Get milestones data from CSV.
        
        Returns:
            DataFrame with milestones
        """
        df = self.csv_loader.get_milestones()
        if df is None or len(df) == 0:
            return pd.DataFrame()
        
        return df
    
    def get_change_orders_data(self) -> pd.DataFrame:
        """Get change orders data from CSV.
        
        Returns:
            DataFrame with change orders
        """
        df = self.csv_loader.get_change_orders()
        if df is None or len(df) == 0:
            return pd.DataFrame()
        
        return df
    
    def get_payments_data(self) -> pd.DataFrame:
        """Get payments data from CSV.
        
        Returns:
            DataFrame with payments
        """
        df = self.csv_loader.get_payments()
        if df is None or len(df) == 0:
            return pd.DataFrame()
        
        return df
    
    def get_progress_updates_data(self) -> pd.DataFrame:
        """Get progress updates data from CSV.
        
        Returns:
            DataFrame with progress updates
        """
        df = self.csv_loader.get_progress_updates()
        if df is None or len(df) == 0:
            return pd.DataFrame()
        
        return df
    
    def get_wbs_data(self) -> pd.DataFrame:
        """Get WBS data from CSV.
        
        Returns:
            DataFrame with WBS
        """
        df = self.csv_loader.get_wbs()
        if df is None or len(df) == 0:
            return pd.DataFrame()
        
        return df
    
    def get_cost_items_data(self) -> pd.DataFrame:
        """Get cost items data from CSV.
        
        Returns:
            DataFrame with cost items
        """
        df = self.csv_loader.get_cost_items()
        if df is None or len(df) == 0:
            return pd.DataFrame()
        
        return df
    
    def get_claims_data(self) -> pd.DataFrame:
        """Get claims data from CSV.
        
        Returns:
            DataFrame with claims
        """
        df = self.csv_loader.get_claims()
        if df is None or len(df) == 0:
            return pd.DataFrame()
        
        return df
    
    def get_data_status(self) -> Dict[str, str]:
        """Get status of all loaded data.
        
        Returns:
            Dictionary with load status for each data source
        """
        status = {
            "csv_files": self.csv_loader.get_load_status(),
            "evm": self.evm_loader.get_load_status(),
        }
        return status
    
    def reload_all_data(self):
        """Reload all data from CSV files."""
        self.csv_loader.reload_all()
        self.evm_loader.reload()


# Global integration instance
_integration = None


def get_integration() -> CSVDataIntegration:
    """Get or create global integration instance.
    
    Returns:
        CSVDataIntegration instance
    """
    global _integration
    if _integration is None:
        _integration = CSVDataIntegration()
    return _integration


def reload_all_data():
    """Reload all data from CSV files."""
    get_integration().reload_all_data()
