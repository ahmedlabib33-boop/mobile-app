from __future__ import annotations

from typing import Any
from datetime import datetime

from .database import get_connection


def _safe_div(a, b):
    if a is None or b in (None, 0):
        return None
    return a / b


def _parse_date(value):
    for fmt in ("%d-%b-%y", "%Y-%m-%d", "%d-%b-%Y"):
        try:
            return datetime.strptime(value, fmt)
        except (TypeError, ValueError):
            continue
    return None


# ── Project Control Summary (EVM from cost_items) ─────────────────────────────

def get_project_control_summary(db_path) -> dict[str, Any]:
    with get_connection(db_path) as conn:
        project = conn.execute(
            "SELECT * FROM projects ORDER BY project_id LIMIT 1"
        ).fetchone()

        if project is None:
            return {"project": None, "kpis": {}, "activities": [], "wbs_costs": []}

        pid = project["project_id"]

        evm = conn.execute(
            """
            SELECT
                COUNT(DISTINCT a.activity_id)                                           AS total_activities,
                AVG(COALESCE(a.planned_progress, 0))                                    AS avg_planned,
                AVG(COALESCE(a.actual_progress,  0))                                    AS avg_actual,
                SUM(COALESCE(a.planned_weight, 0))                                      AS total_weight,
                SUM(COALESCE(a.planned_weight, 0) * COALESCE(a.planned_progress, 0))    AS weighted_planned,
                SUM(COALESCE(a.planned_weight, 0) * COALESCE(a.actual_progress, 0))     AS weighted_actual,
                SUM(CASE WHEN a.is_critical = 1 THEN 1 ELSE 0 END)                     AS critical_activities
            FROM activities a
            WHERE a.project_id = ?
            """,
            (pid,),
        ).fetchone()

        bac = float(project["contract_value"] or 0)

        ac_data = conn.execute(
            """
            SELECT SUM(actual_cost) AS total_ac
            FROM cost_items
            WHERE project_id = ?
            """,
            (project["project_id"],),
        ).fetchone()

        ac = float(ac_data["total_ac"] or 0)

        total_weight = float(evm["total_weight"] or 0)
        weighted_planned_pct = (
            float(evm["weighted_planned"] or 0) / total_weight if total_weight else float(evm["avg_planned"] or 0)
        )
        weighted_actual_pct = (
            float(evm["weighted_actual"] or 0) / total_weight if total_weight else float(evm["avg_actual"] or 0)
        )

        ev = float(project["earned_value"] or 0)
        if ev == 0 and bac > 0:
            ev = bac * weighted_actual_pct / 100.0

        pv = float(project["planned_value"] or 0)
        if pv == 0 and bac > 0:
            pv = bac * weighted_planned_pct / 100.0

        spi = _safe_div(ev, pv)

        cpi = _safe_div(ev, ac)

        eac = _safe_div(bac, cpi) if cpi else 0

        sv = ev - pv

        cv = ev - ac

        tcpi = _safe_div((bac - ev), (bac - ac))

        planned_start = _parse_date(project["planned_start"])

        planned_finish = _parse_date(project["planned_finish"])

        today = datetime.today()

        total_duration = (planned_finish - planned_start).days if planned_start and planned_finish else 0

        elapsed_duration = max((today - planned_start).days, 0) if planned_start else 0

        time_elapsed = 0

        if total_duration > 0:
            time_elapsed = min((elapsed_duration / total_duration) * 100, 100)

        remaining_duration = 100 - time_elapsed
          
        # Activities list
        activities = conn.execute(
            """
           SELECT
           a.activity_id,
            a.activity_name,
            a.wbs_id,
            a.planned_start,
            a.planned_finish,
            a.actual_start,
            a.actual_finish,
            a.forecast_start,
            a.forecast_finish,
            a.planned_weight,
            a.planned_progress,
            a.actual_progress,
            a.total_float_days,
            a.is_critical,
            a.responsible_party
            FROM activities a
            WHERE a.project_id = ?
            ORDER BY a.activity_id
            """,
            (pid,),
        ).fetchall()

        # WBS cost rollup
        wbs_costs = conn.execute(
            """
            SELECT w.wbs_name,
                   SUM(c.budget_cost)       AS budget,
                   SUM(c.actual_cost)       AS actual,
                   SUM(c.forecast_cost)     AS forecast
            FROM cost_items c
            JOIN wbs w ON w.wbs_id = c.wbs_id
            WHERE c.project_id = ?
            GROUP BY w.wbs_name
            ORDER BY budget DESC
            LIMIT 15
            """,
            (pid,),
        ).fetchall()

        # Monthly spend (by activity actual_finish month)
        monthly = conn.execute(
            """
            SELECT substr(update_date, 1, 7) AS month,
                   SUM(actual_cost)          AS actual_cost
            FROM progress_updates
            WHERE project_id = ? AND actual_cost > 0
            GROUP BY month
            ORDER BY month
            """,
            (pid,),
        ).fetchall()

        # Critical vs non-critical progress
        crit_prog = conn.execute(
            """
            SELECT is_critical,
                   AVG(planned_progress) AS avg_planned,
                   AVG(actual_progress)  AS avg_actual,
                   COUNT(*)              AS cnt
            FROM activities
            WHERE project_id = ?
            GROUP BY is_critical
            """,
            (pid,),
        ).fetchall()

        return {
            "project": dict(project),
            "kpis": {
                "total_activities":   int(evm["total_activities"] or 0),
                "bac":  bac,  "ac":  ac,  "ev":  ev,  "pv":  pv,
                "spi": spi,
                "cpi": cpi,
                "eac": eac,
                "sv":   sv,   "cv":  cv,  "tcpi": tcpi,
                "planned_progress": float(project["planned_progress_percent"] or weighted_planned_pct or evm["avg_planned"] or 0),
                "overall_progress": float(project["actual_progress_percent"] or weighted_actual_pct or evm["avg_actual"] or 0),
                "time_elapsed": round(time_elapsed, 2),
                "remaining_duration": round(remaining_duration, 2),
                "critical_activities": int(evm["critical_activities"] or 0),
            },
            
            "activities":  [dict(r) for r in activities],
            "wbs_costs":   [dict(r) for r in wbs_costs],
            "monthly_spend": [dict(r) for r in monthly],
            "crit_progress": [dict(r) for r in crit_prog],
        }


# ── Contract Summary ──────────────────────────────────────────────────────────

def get_contract_summary(db_path) -> list[dict[str, Any]]:
    with get_connection(db_path) as conn:
        rows = conn.execute(
            """
            SELECT
                c.contract_id, c.contract_no, c.contractor_name, c.contract_type,
                c.original_value, c.approved_variations, c.pending_variations,
                c.certified_amount, c.paid_amount, c.retention_amount,
                c.start_date, c.finish_date, c.status,
                (COALESCE(c.original_value,0) + COALESCE(c.approved_variations,0))
                    AS revised_contract_value,
                (COALESCE(c.original_value,0) + COALESCE(c.approved_variations,0)
                    - COALESCE(c.certified_amount,0))
                    AS uncertified_balance,
                (COALESCE(c.certified_amount,0) - COALESCE(c.paid_amount,0))
                    AS unpaid_certified_balance
            FROM contracts c
            ORDER BY c.contract_id
            """
        ).fetchall()
        return [dict(r) for r in rows]


# ── Delay Analysis ────────────────────────────────────────────────────────────

def get_delay_analysis(db_path) -> list[dict[str, Any]]:
    with get_connection(db_path) as conn:
        rows = conn.execute(
            """
            SELECT d.delay_event_id, d.event_title, d.event_date,
                   d.delay_days, d.responsible_party, d.critical_impact,
                   d.eot_days, d.status,
                   a.activity_id AS activity_code, a.activity_name
            FROM delay_events d
            LEFT JOIN activities a ON a.activity_id = d.activity_id
            ORDER BY d.delay_event_id
            """
        ).fetchall()

    results = []
    for row in rows:
        delay_days = int(row["delay_days"] or 0)
        critical   = int(row["critical_impact"] or 0)
        party      = (row["responsible_party"] or "").lower()
        eot_potential = "Yes" if delay_days > 0 and critical == 1 and "employer" in party else "Review"
        concurrent    = "Yes" if delay_days > 0 and critical == 1 and "main contractor" in party else "No"
        results.append({
            **dict(row),
            "delay_days":               delay_days,
            "eot_potential":            eot_potential,
            "concurrent_delay_warning": concurrent,
            "pending_exposure_days":    max(delay_days - int(row["eot_days"] or 0), 0),
        })
    return results


# ── Risk Analysis ─────────────────────────────────────────────────────────────

def get_risk_analysis(db_path) -> list[dict[str, Any]]:
    with get_connection(db_path) as conn:
        rows = conn.execute(
            """
            SELECT r.risk_id, r.risk_code, r.risk_title, r.category,
                   r.probability, r.time_impact_days, r.cost_impact,
                   r.response_strategy, r.owner, r.mitigation_status, r.due_date,
                   (COALESCE(r.probability,0) * COALESCE(r.time_impact_days,0)) AS weighted_time_exposure,
                   (COALESCE(r.probability,0) * COALESCE(r.cost_impact,0))       AS weighted_cost_exposure
            FROM risks r
            ORDER BY (COALESCE(r.probability,0) * COALESCE(r.time_impact_days,0)) DESC
            """
        ).fetchall()

    results = []
    for row in rows:
        score = float(row["weighted_time_exposure"] or 0)
        band  = "High" if score >= 15 else ("Medium" if score >= 8 else "Low")
        results.append({**dict(row), "severity_band": band})
    return results


# ── Milestone Status ──────────────────────────────────────────────────────────

def get_milestones(db_path) -> list[dict[str, Any]]:
    with get_connection(db_path) as conn:
        rows = conn.execute(
            """
            SELECT milestone_id, milestone_name, target_date, actual_date, status
            FROM milestones
            ORDER BY target_date
            """
        ).fetchall()
        return [dict(r) for r in rows]


# ── Change Orders ─────────────────────────────────────────────────────────────

def get_change_orders(db_path) -> list[dict[str, Any]]:
    with get_connection(db_path) as conn:
        rows = conn.execute(
            """
            SELECT change_order_id, title, issue_date,
                   cost_impact, time_impact_days, status
            FROM change_orders
            ORDER BY issue_date DESC
            """
        ).fetchall()
        return [dict(r) for r in rows]
