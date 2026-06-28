from pathlib import Path
import sqlite3
from contextlib import contextmanager

ROOT = Path(__file__).resolve().parent.parent.parent

DEFAULT_DB_PATH = ROOT / "construction_system.db"


@contextmanager
def get_connection(db_path=DEFAULT_DB_PATH):
    """Yield a sqlite3 connection and close it when done."""
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def init_db(db_path=DEFAULT_DB_PATH, reset=False):
    if reset:
        Path(db_path).unlink(missing_ok=True)

    # Use a plain connection here (not the context manager) so we can
    # run executescript which issues an implicit COMMIT.
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()

    cur.executescript("""

CREATE TABLE IF NOT EXISTS projects (
    project_id      TEXT PRIMARY KEY,
    project_name    TEXT,
    project_code    TEXT,
    client_name     TEXT,
    contractor      TEXT,
    planned_start   TEXT,
    planned_finish  TEXT,
    forecast_finish TEXT,
    contract_value  REAL,
    currency        TEXT,
    status          TEXT,
    planned_progress_percent REAL,
    actual_progress_percent REAL,
    planned_value REAL,
    earned_value REAL
);

CREATE TABLE IF NOT EXISTS wbs (
    wbs_id       TEXT PRIMARY KEY,
    project_id   TEXT,
    wbs_code     TEXT,
    wbs_name     TEXT,
    parent_wbs_id TEXT,
    FOREIGN KEY (project_id) REFERENCES projects(project_id)
);

CREATE TABLE IF NOT EXISTS activities (
    activity_id       TEXT PRIMARY KEY,
    project_id        TEXT,
    wbs_id            TEXT,
    activity_code     TEXT,
    activity_name     TEXT,
    planned_start     TEXT,
    planned_finish    TEXT,
    actual_start      TEXT,
    actual_finish     TEXT,
    forecast_start    TEXT,
    forecast_finish   TEXT,
    planned_weight    REAL,
    planned_progress  REAL,
    actual_progress   REAL,
    budget            REAL DEFAULT 0,
    actual_cost       REAL DEFAULT 0,
    total_float_days  REAL,
    is_critical       INTEGER DEFAULT 0,
    responsible_party TEXT,
    FOREIGN KEY (project_id) REFERENCES projects(project_id),
    FOREIGN KEY (wbs_id) REFERENCES wbs(wbs_id)
);

CREATE TABLE IF NOT EXISTS progress_updates (
    update_id        TEXT PRIMARY KEY,
    project_id       TEXT,
    activity_id      TEXT,
    update_date      TEXT,
    planned_progress REAL,
    actual_progress  REAL,
    planned_value    REAL,
    earned_value     REAL,
    actual_cost      REAL,
    notes            TEXT,
    FOREIGN KEY (project_id)  REFERENCES projects(project_id),
    FOREIGN KEY (activity_id) REFERENCES activities(activity_id)
);

CREATE TABLE IF NOT EXISTS cost_items (
    cost_item_id  TEXT PRIMARY KEY,
    project_id    TEXT,
    wbs_id        TEXT,
    cost_category TEXT,
    budget_cost   REAL,
    actual_cost   REAL,
    forecast_cost REAL,
    FOREIGN KEY (project_id) REFERENCES projects(project_id),
    FOREIGN KEY (wbs_id)     REFERENCES wbs(wbs_id)
);

CREATE TABLE IF NOT EXISTS contracts (
    contract_id          TEXT PRIMARY KEY,
    project_id           TEXT,
    contract_no          TEXT,
    contractor_name      TEXT,
    contract_type        TEXT,
    original_value       REAL,
    approved_variations  REAL DEFAULT 0,
    pending_variations   REAL DEFAULT 0,
    certified_amount     REAL DEFAULT 0,
    paid_amount          REAL DEFAULT 0,
    retention_amount     REAL DEFAULT 0,
    start_date           TEXT,
    finish_date          TEXT,
    status               TEXT,
    FOREIGN KEY (project_id) REFERENCES projects(project_id)
);

CREATE TABLE IF NOT EXISTS change_orders (
    change_order_id TEXT PRIMARY KEY,
    project_id      TEXT,
    contract_id     TEXT,
    title           TEXT,
    issue_date      TEXT,
    cost_impact     REAL,
    time_impact_days REAL,
    status          TEXT,
    FOREIGN KEY (project_id) REFERENCES projects(project_id)
);

CREATE TABLE IF NOT EXISTS claims (
    claim_id         TEXT PRIMARY KEY,
    project_id       TEXT,
    claim_title      TEXT,
    claim_type       TEXT,
    claim_date       TEXT,
    claim_value      REAL,
    status           TEXT,
    responsible_party TEXT,
    eot_days         REAL,
    FOREIGN KEY (project_id) REFERENCES projects(project_id)
);

CREATE TABLE IF NOT EXISTS payments (
    payment_id   TEXT PRIMARY KEY,
    project_id   TEXT,
    contract_id  TEXT,
    invoice_no   TEXT,
    payment_date TEXT,
    amount       REAL,
    status       TEXT,
    FOREIGN KEY (project_id) REFERENCES projects(project_id)
);

CREATE TABLE IF NOT EXISTS delay_events (
    delay_event_id    TEXT PRIMARY KEY,
    project_id        TEXT,
    activity_id       TEXT,
    event_title       TEXT,
    event_date        TEXT,
    delay_days        REAL,
    responsible_party TEXT,
    critical_impact   INTEGER DEFAULT 0,
    eot_days          REAL DEFAULT 0,
    status            TEXT,
    FOREIGN KEY (project_id)  REFERENCES projects(project_id),
    FOREIGN KEY (activity_id) REFERENCES activities(activity_id)
);

CREATE TABLE IF NOT EXISTS risks (
    risk_id           TEXT PRIMARY KEY,
    project_id        TEXT,
    risk_code         TEXT,
    risk_title        TEXT,
    category          TEXT,
    probability       REAL,
    time_impact_days  REAL,
    cost_impact       REAL,
    response_strategy TEXT,
    owner             TEXT,
    mitigation_status TEXT,
    due_date          TEXT,
    FOREIGN KEY (project_id) REFERENCES projects(project_id)
);

CREATE TABLE IF NOT EXISTS milestones (
    milestone_id   TEXT PRIMARY KEY,
    project_id     TEXT,
    milestone_name TEXT,
    target_date    TEXT,
    actual_date    TEXT,
    status         TEXT,
    FOREIGN KEY (project_id) REFERENCES projects(project_id)
);

""")

    conn.commit()
    conn.close()

    print("Database initialized successfully.")
