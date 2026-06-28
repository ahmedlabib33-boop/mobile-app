from __future__ import annotations

from .database import init_db, get_connection


def seed_demo_data(db_path="construction_system.db"):
    """Initialise the database and insert a realistic demo dataset."""

    init_db(db_path, reset=True)

    with get_connection(db_path) as conn:

        # ── Projects ──────────────────────────────────────────────────────────
        conn.execute(
            """
            INSERT INTO projects
                (project_id, project_name, project_code, client_name, contractor,
                 planned_start, planned_finish, forecast_finish,
                 contract_value, currency, status)
            VALUES (?,?,?,?,?,?,?,?,?,?,?)
            """,
            (
                "PRJ-001",
                "Riverside Mixed-Use Development",
                "RMD-2024",
                "Riverside Holdings Ltd",
                "BuildCo International",
                "2024-01-15",
                "2025-12-31",
                "2026-02-28",
                45_000_000.0,
                "USD",
                "In Progress",
            ),
        )

        # ── WBS ───────────────────────────────────────────────────────────────
        wbs_rows = [
            ("WBS-001", "PRJ-001", "1",   "Civil Works",          None),
            ("WBS-002", "PRJ-001", "1.1", "Foundations",          "WBS-001"),
            ("WBS-003", "PRJ-001", "1.2", "Superstructure",       "WBS-001"),
            ("WBS-004", "PRJ-001", "2",   "MEP Works",            None),
            ("WBS-005", "PRJ-001", "2.1", "Mechanical",           "WBS-004"),
            ("WBS-006", "PRJ-001", "2.2", "Electrical",           "WBS-004"),
            ("WBS-007", "PRJ-001", "3",   "Finishing Works",      None),
        ]
        conn.executemany(
            "INSERT INTO wbs (wbs_id, project_id, wbs_code, wbs_name, parent_wbs_id) VALUES (?,?,?,?,?)",
            wbs_rows,
        )

        # ── Activities ────────────────────────────────────────────────────────
        activity_rows = [
            # id, proj, wbs, code, name, pl_start, pl_fin, ac_start, ac_fin, fc_start, fc_fin,
            # pl_weight, pl_prog, ac_prog, budget, actual_cost, float, is_critical, resp_party
            ("ACT-001","PRJ-001","WBS-002","A1010","Site Preparation",
             "2024-01-15","2024-02-28","2024-01-15","2024-03-05",None,None,
             5.0, 100.0, 100.0, 500_000.0, 520_000.0, 0.0, 1, "BuildCo International"),
            ("ACT-002","PRJ-001","WBS-002","A1020","Piling Works",
             "2024-03-01","2024-05-31","2024-03-06",None,None,"2024-06-15",
             10.0, 100.0, 95.0, 2_000_000.0, 1_950_000.0, 0.0, 1, "BuildCo International"),
            ("ACT-003","PRJ-001","WBS-003","A2010","Ground Floor Slab",
             "2024-06-01","2024-07-31","2024-06-16",None,None,"2024-08-15",
             8.0, 80.0, 70.0, 1_500_000.0, 1_100_000.0, 0.0, 1, "BuildCo International"),
            ("ACT-004","PRJ-001","WBS-003","A2020","Structural Frame – Levels 1-5",
             "2024-08-01","2024-11-30",None,None,None,"2024-12-31",
             15.0, 40.0, 25.0, 5_000_000.0, 1_200_000.0, 5.0, 1, "BuildCo International"),
            ("ACT-005","PRJ-001","WBS-005","M1010","HVAC Rough-In",
             "2024-10-01","2025-02-28",None,None,None,"2025-03-31",
             10.0, 20.0, 10.0, 3_000_000.0, 300_000.0, 10.0, 0, "MechSub Ltd"),
            ("ACT-006","PRJ-001","WBS-006","E1010","LV Distribution",
             "2024-10-01","2025-03-31",None,None,None,"2025-04-30",
             10.0, 15.0, 8.0, 2_500_000.0, 200_000.0, 15.0, 0, "ElecSub Ltd"),
            ("ACT-007","PRJ-001","WBS-007","F1010","Internal Plastering",
             "2025-01-01","2025-06-30",None,None,None,"2025-07-31",
             12.0, 0.0, 0.0, 2_000_000.0, 0.0, 20.0, 0, "FinishCo"),
            ("ACT-008","PRJ-001","WBS-007","F1020","Tiling & Flooring",
             "2025-04-01","2025-09-30",None,None,None,"2025-10-31",
             10.0, 0.0, 0.0, 1_800_000.0, 0.0, 25.0, 0, "FinishCo"),
        ]
        conn.executemany(
            """
            INSERT INTO activities
                (activity_id, project_id, wbs_id, activity_code, activity_name,
                 planned_start, planned_finish, actual_start, actual_finish,
                 forecast_start, forecast_finish,
                 planned_weight, planned_progress, actual_progress,
                 budget, actual_cost, total_float_days, is_critical, responsible_party)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            """,
            activity_rows,
        )

        # ── Contracts ─────────────────────────────────────────────────────────
        contract_rows = [
            ("CON-001","PRJ-001","C-2024-001","BuildCo International","Lump Sum",
             38_000_000.0, 1_200_000.0, 350_000.0, 18_500_000.0, 17_000_000.0,
             1_850_000.0, "2024-01-15","2025-12-31","Active"),
            ("CON-002","PRJ-001","C-2024-002","MechSub Ltd","Remeasurement",
             4_500_000.0, 0.0, 0.0, 800_000.0, 750_000.0,
             80_000.0, "2024-10-01","2025-06-30","Active"),
            ("CON-003","PRJ-001","C-2024-003","ElecSub Ltd","Remeasurement",
             3_800_000.0, 0.0, 0.0, 500_000.0, 450_000.0,
             50_000.0, "2024-10-01","2025-07-31","Active"),
        ]
        conn.executemany(
            """
            INSERT INTO contracts
                (contract_id, project_id, contract_no, contractor_name, contract_type,
                 original_value, approved_variations, pending_variations,
                 certified_amount, paid_amount, retention_amount,
                 start_date, finish_date, status)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            """,
            contract_rows,
        )

        # ── Delay Events ──────────────────────────────────────────────────────
        delay_rows = [
            ("DEL-001","PRJ-001","ACT-001",
             "Unforeseen Ground Conditions","2024-02-10",
             5.0,"Employer",1,5.0,"Closed"),
            ("DEL-002","PRJ-001","ACT-002",
             "Piling Equipment Breakdown","2024-03-20",
             15.0,"Main Contractor",1,0.0,"Open"),
            ("DEL-003","PRJ-001","ACT-003",
             "Design Change – Slab Thickness","2024-06-25",
             10.0,"Employer",1,10.0,"Open"),
            ("DEL-004","PRJ-001","ACT-004",
             "Steel Delivery Delay","2024-09-01",
             20.0,"Main Contractor",1,0.0,"Open"),
        ]
        conn.executemany(
            """
            INSERT INTO delay_events
                (delay_event_id, project_id, activity_id,
                 event_title, event_date, delay_days,
                 responsible_party, critical_impact, eot_days, status)
            VALUES (?,?,?,?,?,?,?,?,?,?)
            """,
            delay_rows,
        )

        # ── Risks ─────────────────────────────────────────────────────────────
        risk_rows = [
            ("RSK-001","PRJ-001","R-001","Labour Shortage","Resource",
             0.6, 30.0, 500_000.0,"Contingency Plan","BuildCo International","Active","2025-03-31"),
            ("RSK-002","PRJ-001","R-002","Material Price Escalation","Cost",
             0.5, 0.0, 800_000.0,"Fixed-Price Sub-Contracts","Project Manager","Active","2025-06-30"),
            ("RSK-003","PRJ-001","R-003","Permit Delay – Phase 2","Regulatory",
             0.4, 45.0, 200_000.0,"Early Submission","Client","Active","2024-12-31"),
            ("RSK-004","PRJ-001","R-004","Subcontractor Insolvency","Commercial",
             0.2, 60.0, 1_200_000.0,"Performance Bond","Project Manager","Active","2025-09-30"),
            ("RSK-005","PRJ-001","R-005","Adverse Weather","Environmental",
             0.3, 10.0, 50_000.0,"Float Buffer","Site Manager","Mitigated","2025-01-31"),
        ]
        conn.executemany(
            """
            INSERT INTO risks
                (risk_id, project_id, risk_code, risk_title, category,
                 probability, time_impact_days, cost_impact,
                 response_strategy, owner, mitigation_status, due_date)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?)
            """,
            risk_rows,
        )

        # ── Milestones ────────────────────────────────────────────────────────
        milestone_rows = [
            ("MS-001","PRJ-001","Piling Complete",       "2024-05-31","2024-06-15","Delayed"),
            ("MS-002","PRJ-001","Ground Floor Slab",     "2024-07-31",None,        "In Progress"),
            ("MS-003","PRJ-001","Structure Topping Out", "2024-11-30",None,        "Pending"),
            ("MS-004","PRJ-001","MEP Rough-In Complete", "2025-02-28",None,        "Pending"),
            ("MS-005","PRJ-001","Practical Completion",  "2025-12-31",None,        "Pending"),
        ]
        conn.executemany(
            "INSERT INTO milestones (milestone_id, project_id, milestone_name, target_date, actual_date, status) VALUES (?,?,?,?,?,?)",
            milestone_rows,
        )

    print("Database seeded successfully.")


if __name__ == "__main__":
    seed_demo_data()
