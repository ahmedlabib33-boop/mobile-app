from __future__ import annotations

from datetime import date, datetime
from typing import Any

from agents import function_tool


def _split_terms(text: str) -> list[str]:
    separators = ["\n", ";", ","]
    terms = [text]
    for separator in separators:
        expanded: list[str] = []
        for item in terms:
            expanded.extend(item.split(separator))
        terms = expanded
    return [item.strip(" -\t") for item in terms if item.strip(" -\t")]


def extract_tasks_from_brief_impl(product_brief: str, launch_date: str, constraints: str) -> dict[str, Any]:
    text = f"{product_brief}\n{constraints}".lower()
    tasks = [
        {"task": "Confirm release scope and success criteria", "priority": "P0", "owner": "Product / Engineering"},
        {"task": "Freeze input data sources and required formats", "priority": "P0", "owner": "Data Analytics"},
        {"task": "Build launch readiness review pack", "priority": "P1", "owner": "Project Controls"},
    ]
    if any(term in text for term in ["delay", "tia", "critical path", "schedule", "primavera"]):
        tasks.insert(1, {"task": "Run delay analysis and critical path impact review", "priority": "P0", "owner": "Planning Engineer"})
    if any(term in text for term in ["csv", "excel", "pdf", "drawing", "xer", "xml"]):
        tasks.append({"task": "Normalize mixed-format assets into a single evidence register", "priority": "P1", "owner": "Data Engineering"})
    if any(term in text for term in ["client", "contract", "claim", "eot"]):
        tasks.append({"task": "Map release risks to contract, claim, and EOT evidence needs", "priority": "P1", "owner": "Contracts"})

    return {
        "launch_date": launch_date,
        "task_count": len(tasks),
        "tasks": tasks,
    }


@function_tool
def extract_tasks_from_brief(product_brief: str, launch_date: str, constraints: str) -> dict[str, Any]:
    """Extract actionable launch and delay-analysis tasks from the brief."""
    return extract_tasks_from_brief_impl(product_brief, launch_date, constraints)


def check_launch_readiness_impl(product_brief: str, audience: str, launch_date: str, available_assets: str) -> dict[str, Any]:
    today = date.today()
    try:
        parsed_launch = datetime.fromisoformat(launch_date[:10]).date()
        days_to_launch = (parsed_launch - today).days
    except ValueError:
        days_to_launch = None

    assets = _split_terms(available_assets)
    rubric = {
        "scope_clarity": 25 if len(product_brief) >= 120 else 12,
        "audience_fit": 20 if len(audience) >= 12 else 10,
        "asset_readiness": min(25, len(assets) * 5),
        "schedule_feasibility": 20 if days_to_launch is None or days_to_launch >= 14 else 8,
        "risk_visibility": 10 if any(term in product_brief.lower() for term in ["risk", "delay", "constraint", "dependency"]) else 4,
    }
    total = sum(rubric.values())
    status = "Ready with controls" if total >= 75 else "Needs closure before launch" if total >= 50 else "Not launch ready"
    return {
        "score": total,
        "status": status,
        "rubric": rubric,
        "days_to_launch": days_to_launch,
        "asset_count": len(assets),
    }


@function_tool
def check_launch_readiness(product_brief: str, audience: str, launch_date: str, available_assets: str) -> dict[str, Any]:
    """Score launch readiness against a practical engineering analytics rubric."""
    return check_launch_readiness_impl(product_brief, audience, launch_date, available_assets)


def generate_owner_checklist_impl(available_assets: str, constraints: str, delay_analysis_context: str) -> dict[str, Any]:
    constraints_list = _split_terms(constraints)
    assets = _split_terms(available_assets)
    checklist = [
        {"owner": "Engineering Lead", "item": "Approve release scope, dependencies, rollback owner, and launch window."},
        {"owner": "Data Analytics", "item": "Validate source data formats, lineage, assumptions, and transformation logic."},
        {"owner": "Planning Engineer", "item": "Confirm baseline, update, critical path, float movement, and delay-event windows."},
        {"owner": "Commercial / Contracts", "item": "Confirm notices, claim language, entitlement basis, and evidence gaps."},
        {"owner": "Communications", "item": "Prepare internal launch note, client-facing summary, and Q&A script."},
    ]
    if constraints_list:
        checklist.append({"owner": "Risk Owner", "item": f"Close constraint register: {constraints_list[0]}."})
    if assets:
        checklist.append({"owner": "Release Coordinator", "item": f"Package available assets and verify access to {len(assets)} source items."})
    if delay_analysis_context.strip():
        checklist.append({"owner": "Delay Analyst", "item": "Convert delay-analysis context into a dated action register and decision log."})
    return {"checklist": checklist}


@function_tool
def generate_owner_checklist(available_assets: str, constraints: str, delay_analysis_context: str) -> dict[str, Any]:
    """Generate owner-specific launch checklist items for engineering and project controls."""
    return generate_owner_checklist_impl(available_assets, constraints, delay_analysis_context)


def draft_channel_launch_copy_impl(product_brief: str, audience: str) -> dict[str, Any]:
    one_line = product_brief.strip().split(".")[0][:180]
    return {
        "internal_email": f"Team, we are preparing to launch: {one_line}. The focus is on {audience}. Please review scope, risks, and owner actions before release.",
        "executive_update": f"Launch planning is underway for {one_line}. The plan converts analytics and delay-analysis inputs into prioritized engineering actions, risk controls, and accountable owner checklists.",
        "client_note": f"We are aligning the release plan around validated data, schedule-impact logic, and clear follow-up decisions for {audience}.",
        "standup_prompt": "What changed since yesterday, which launch risk moved, and which evidence or owner decision is blocking release readiness?",
    }


@function_tool
def draft_channel_launch_copy(product_brief: str, audience: str) -> dict[str, Any]:
    """Draft channel-specific launch copy suggestions."""
    return draft_channel_launch_copy_impl(product_brief, audience)
