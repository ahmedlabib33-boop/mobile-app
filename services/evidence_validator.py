from __future__ import annotations

import re
from typing import Any


SCHEDULE_TERMS = {
    "delay", "eot", "extension of time", "critical path", "float", "baseline", "update",
    "fragnet", "concurrency", "schedule", "programme", "program", "activity", "predecessor", "successor",
}
LEGAL_TERMS = {"fidic", "law", "clause", "contract", "claim", "notice", "entitlement", "authority", "regulation"}


def yes(value: Any) -> bool:
    return str(value or "").strip().casefold() in {"yes", "true", "1", "y", "required"}


def classify_question(question: str, row: dict[str, Any] | None = None) -> dict[str, str]:
    text = f"{question} {row.get('department', '') if row else ''} {row.get('section_or_level', '') if row else ''}".casefold()
    categories = [
        ("Delay / Schedule", ["delay", "eot", "schedule", "programme", "critical", "float", "milestone"]),
        ("Cost / Payment", ["cost", "payment", "cash", "variation", "budget", "invoice"]),
        ("Quality", ["quality", "ncr", "inspection", "wir", "mir", "defect"]),
        ("HSE", ["hse", "safety", "incident", "risk", "fire", "civil defense"]),
        ("Procurement", ["procurement", "material", "supplier", "purchase", "submittal"]),
        ("Design / Engineering", ["design", "drawing", "ifc", "rfi", "engineering", "shop drawing"]),
        ("Claim / Contract", ["claim", "contract", "fidic", "clause", "notice", "entitlement"]),
        ("Handover / Maintenance", ["handover", "snag", "maintenance", "warranty", "asset"]),
    ]
    for label, terms in categories:
        if any(term in text for term in terms):
            return {"department_group": label, "issue_type": label}
    return {"department_group": "General Project Controls", "issue_type": "General"}


def requires_schedule_evidence(question: str) -> bool:
    lowered = question.casefold()
    return any(term in lowered for term in SCHEDULE_TERMS)


def requires_legal_basis(question: str, fidic_enabled: bool = False, egypt_law_enabled: bool = False) -> bool:
    lowered = question.casefold()
    return fidic_enabled or egypt_law_enabled or any(term in lowered for term in LEGAL_TERMS)


def has_schedule_evidence(evidence: list[dict[str, Any]]) -> bool:
    pattern = re.compile(r"activity|baseline|finish|start|float|critical|predecessor|successor|delay|milestone", re.I)
    for item in evidence:
        blob = " ".join(str(item.get(key, "")) for key in ("source_file", "matched_fields", "excerpt", "sheet_name"))
        if pattern.search(blob):
            return True
    return False


def validate_answer(
    *,
    question: str,
    answer: str,
    evidence: list[dict[str, Any]],
    legal_refs: list[dict[str, Any]],
    require_project_data: bool,
    fidic_enabled: bool,
    egypt_law_enabled: bool,
) -> dict[str, Any]:
    missing: list[str] = []
    checks: list[dict[str, str]] = []
    score = 35

    if require_project_data and not evidence:
        missing.append("Project evidence is required but no matching project records were found.")
        checks.append({"check": "Project evidence", "status": "Missing"})
    elif evidence:
        score += min(25, 8 + len(evidence) * 3)
        checks.append({"check": "Project evidence", "status": f"{len(evidence)} item(s) found"})

    if requires_schedule_evidence(question):
        if has_schedule_evidence(evidence):
            score += 15
            checks.append({"check": "Schedule evidence", "status": "Found"})
        else:
            missing.append("Schedule questions need activity, baseline/update, float, milestone, or critical-path evidence.")
            checks.append({"check": "Schedule evidence", "status": "Missing"})

    if requires_legal_basis(question, fidic_enabled, egypt_law_enabled):
        if legal_refs:
            score += 15
            checks.append({"check": "Legal/FIDIC basis", "status": f"{len(legal_refs)} public reference(s)"})
        else:
            missing.append("Legal/FIDIC basis is enabled or implied, but no cited public reference was found.")
            checks.append({"check": "Legal/FIDIC basis", "status": "Missing"})

    if answer.strip():
        score += 10
        checks.append({"check": "Answer draft", "status": "Generated"})
    else:
        missing.append("No answer was generated.")
        checks.append({"check": "Answer draft", "status": "Missing"})

    score = max(0, min(100, score - len(missing) * 8))
    confidence = "High" if score >= 80 and not missing else "Medium" if score >= 55 else "Low"
    return {
        "confidence": confidence,
        "score": score,
        "missing_evidence": missing,
        "checks": checks,
        "risks": missing[:],
        "export_blocking": confidence == "Low",
    }
