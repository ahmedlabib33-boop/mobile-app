from pathlib import Path
import sys

# Ensure the src package is importable when running pytest from the project root
ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from construction_system.database import init_db
from construction_system.seed import seed_demo_data
from construction_system.analytics import (
    get_project_control_summary,
    get_contract_summary,
    get_delay_analysis,
    get_risk_analysis,
)


def test_project_control_summary(tmp_path):
    db_path = tmp_path / "test.db"
    seed_demo_data(db_path)

    summary = get_project_control_summary(db_path)

    assert summary["project"] is not None
    assert summary["kpis"]["total_activities"] > 0
    assert summary["kpis"]["bac"] > 0


def test_contract_summary(tmp_path):
    db_path = tmp_path / "test.db"
    seed_demo_data(db_path)

    contracts = get_contract_summary(db_path)

    assert len(contracts) > 0
    assert "original_value" in contracts[0]
    assert "unpaid_certified_balance" in contracts[0]


def test_delay_analysis(tmp_path):
    db_path = tmp_path / "test.db"
    seed_demo_data(db_path)

    delays = get_delay_analysis(db_path)

    assert len(delays) > 0
    assert "eot_potential" in delays[0]
    assert "concurrent_delay_warning" in delays[0]


def test_risk_analysis(tmp_path):
    db_path = tmp_path / "test.db"
    seed_demo_data(db_path)

    risks = get_risk_analysis(db_path)

    assert len(risks) > 0
    assert "severity_band" in risks[0]
    # Highest-scored risk should be first
    scores = [r["weighted_time_exposure"] for r in risks]
    assert scores == sorted(scores, reverse=True)
