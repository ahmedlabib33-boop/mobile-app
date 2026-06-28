"""EVM Data Loader - Load EVM data from CSV file."""

import pandas as pd
from pathlib import Path
from typing import Optional, Dict
import numpy as np

# Define EVM CSV file path
DATA_DIR = Path(__file__).parent.parent.parent / "data" / "import_templates"
EVM_FILE = DATA_DIR / "evm.csv"


class EVMDataLoader:
    """Load and process EVM data from CSV."""
    
    def __init__(self):
        """Initialize EVM loader."""
        self.df = None
        self.load_status = "Not loaded"
    
    def load_evm_data(self) -> bool:
        """Load EVM data from CSV file.
        
        Returns:
            True if loaded successfully, False otherwise
        """
        if not EVM_FILE.exists():
            self.load_status = f"EVM file not found: {EVM_FILE}"
            return False
        
        try:
            self.df = pd.read_csv(EVM_FILE)
            self.load_status = f"Loaded: {len(self.df)} rows"
            return True
        except Exception as e:
            self.load_status = f"Error loading EVM: {str(e)}"
            return False
    
    def get_evm_data(self) -> Optional[pd.DataFrame]:
        """Get EVM data, loading if necessary.
        
        Returns:
            DataFrame with EVM data or None if not available
        """
        if self.df is None:
            self.load_evm_data()
        return self.df
    
    def get_evm_summary(self) -> Dict:
        """Get EVM summary metrics.
        
        Returns:
            Dictionary with EVM metrics
        """
        df = self.get_evm_data()
        if df is None or len(df) == 0:
            return {}
        
        summary = {}
        
        # Calculate metrics from EVM data
        # Expected columns: BAC, AC, EV, PV, etc.
        
        if "BAC" in df.columns:
            summary["BAC"] = df["BAC"].sum()
        if "AC" in df.columns:
            summary["AC"] = df["AC"].sum()
        if "EV" in df.columns:
            summary["EV"] = df["EV"].sum()
        if "PV" in df.columns:
            summary["PV"] = df["PV"].sum()
        
        # Calculate indices
        if "AC" in summary and summary["AC"] > 0 and "EV" in summary:
            summary["CPI"] = summary["EV"] / summary["AC"]
        if "PV" in summary and summary["PV"] > 0 and "EV" in summary:
            summary["SPI"] = summary["EV"] / summary["PV"]
        
        # Calculate variances
        if "EV" in summary and "AC" in summary:
            summary["CV"] = summary["EV"] - summary["AC"]
        if "EV" in summary and "PV" in summary:
            summary["SV"] = summary["EV"] - summary["PV"]
        
        return summary
    
    def get_evm_by_activity(self) -> Optional[pd.DataFrame]:
        """Get EVM data by activity.
        
        Returns:
            DataFrame with EVM data by activity
        """
        df = self.get_evm_data()
        if df is None or len(df) == 0:
            return None
        
        # Group by activity if activity column exists
        if "activity_id" in df.columns or "activity" in df.columns:
            activity_col = "activity_id" if "activity_id" in df.columns else "activity"
            return df.groupby(activity_col).agg({
                "BAC": "sum",
                "AC": "sum",
                "EV": "sum",
                "PV": "sum"
            }).reset_index()
        
        return df
    
    def get_evm_by_period(self) -> Optional[pd.DataFrame]:
        """Get EVM data by period (month/week).
        
        Returns:
            DataFrame with EVM data by period
        """
        df = self.get_evm_data()
        if df is None or len(df) == 0:
            return None
        
        # Group by period if period column exists
        if "period" in df.columns or "month" in df.columns or "week" in df.columns:
            period_col = None
            for col in ["period", "month", "week", "date"]:
                if col in df.columns:
                    period_col = col
                    break
            
            if period_col:
                return df.groupby(period_col).agg({
                    "BAC": "sum",
                    "AC": "sum",
                    "EV": "sum",
                    "PV": "sum"
                }).reset_index()
        
        return df
    
    def get_load_status(self) -> str:
        """Get load status message.
        
        Returns:
            Status message
        """
        return self.load_status
    
    def reload(self):
        """Reload EVM data from file."""
        self.df = None
        self.load_evm_data()


# Global loader instance
_evm_loader = None


def get_evm_loader() -> EVMDataLoader:
    """Get or create global EVM loader instance.
    
    Returns:
        EVMDataLoader instance
    """
    global _evm_loader
    if _evm_loader is None:
        _evm_loader = EVMDataLoader()
    return _evm_loader


def get_evm_data() -> Optional[pd.DataFrame]:
    """Get EVM data.
    
    Returns:
        DataFrame with EVM data or None
    """
    return get_evm_loader().get_evm_data()


def get_evm_summary() -> Dict:
    """Get EVM summary metrics.
    
    Returns:
        Dictionary with EVM metrics
    """
    return get_evm_loader().get_evm_summary()


def reload_evm():
    """Reload EVM data from file."""
    get_evm_loader().reload()
