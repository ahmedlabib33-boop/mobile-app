"""
Contract Clause Matching Engine

Matches project events (delays, claims, variations) to applicable contract clauses
and provides entitlement analysis, notice requirements, and financial/schedule impacts.
"""

from pathlib import Path
from typing import Any, Dict, List
from dataclasses import dataclass
import re

import pandas as pd


@dataclass
class ContractClause:
    """Represents a contract clause with its properties."""
    clause_id: str
    topic: str
    location: str
    plain_english: str
    beneath_lines: str
    leverage_holder: str
    notice_requirement: str
    money_impact: str
    schedule_impact: str
    practical_action: str


PROJECT_ROOT = Path(__file__).resolve().parents[2]
CLAUSE_LIBRARY_CANDIDATES = [
    PROJECT_ROOT / "projects" / "_PROJECT_TEMPLATE" / "contracts" / "source" / "Overall_Contract_clause_library.xlsx",
]
CLAUSE_LIBRARY_PATH = next((path for path in CLAUSE_LIBRARY_CANDIDATES if path.exists()), CLAUSE_LIBRARY_CANDIDATES[0])


def set_clause_library_path(path: Path) -> None:
    """Set the active project's clause library without cross-project fallback."""
    global CLAUSE_LIBRARY_PATH, CONTRACT_CLAUSES
    CLAUSE_LIBRARY_PATH = Path(path)
    if "CONTRACT_CLAUSES" in globals():
        CONTRACT_CLAUSES = _load_clause_library()


def _seed_clauses() -> List[ContractClause]:
    return [
    ContractClause(
        clause_id="1.0",
        topic="Contract basics",
        location="Agreement / Contract Data",
        plain_english="The signed agreement and contract data define the parties, project, effective date, price, time, and risk allocation.",
        beneath_lines="All later letters should be read against the project-specific signed agreement and contract data.",
        leverage_holder="Neutral",
        notice_requirement="Use exact contract references in all notices.",
        money_impact="Defines the starting commercial baseline.",
        schedule_impact="Defines the baseline programme period.",
        practical_action="Keep contract date, parties, amount, and commencement records in every claim file."
    ),
    ContractClause(
        clause_id="1.1",
        topic="Accepted Contract Amount",
        location="Agreement / Appendix",
        plain_english="The accepted contract amount and included taxes are governed by the project-specific signed agreement and appendix.",
        beneath_lines="The price is intended to be comprehensive. The Contractor must prove why any extra is outside the agreed scope or qualifies under variation/adjustment clauses.",
        leverage_holder="Employer",
        notice_requirement="Any extra must follow Variation or Claim procedure.",
        money_impact="High barrier to extra payment unless clear instruction/change exists.",
        schedule_impact="No direct time right unless linked to variation or delay event.",
        practical_action="For each extra cost, attach instruction, BOQ gap analysis, and timely notice."
    ),
    ContractClause(
        clause_id="1.2",
        topic="Re-measured contract",
        location="Agreement / Appendix",
        plain_english="Contract is on a re-measured basis and payment is based on actual executed quantities.",
        beneath_lines="Quantities can increase or decrease, but rates and risk wording still matter. More quantity is not automatically a new scope claim.",
        leverage_holder="Shared",
        notice_requirement="Measurement records must be timely and supported.",
        money_impact="Payment follows measured work, not tender BOQ quantity warranty.",
        schedule_impact="Quantity change may affect time only if it impacts critical path and is claimed.",
        practical_action="Keep survey sheets, IR approvals, measurement sheets, and BOQ mapping."
    ),
    ContractClause(
        clause_id="4.20",
        topic="Employer equipment and free-issue materials",
        location="Clause 4.20 / Appendix",
        plain_english="Employer-supplied items, especially steel as part of advance, depend on approved programme and requisition process.",
        beneath_lines="A steel delay claim needs proof that Contractor requested correctly, was ready to use it, and the delay hit critical path.",
        leverage_holder="Shared",
        notice_requirement="Notice immediately if free issue material is late, defective, or short.",
        money_impact="May affect payment/advance and possible prolongation.",
        schedule_impact="Possible EOT if critical and not Contractor-caused.",
        practical_action="Keep approved programme, requisitions, delivery notes, stock records, and activity impacts."
    ),
    ContractClause(
        clause_id="8.2",
        topic="Time for Completion",
        location="Appendix / Clause 8.2",
        plain_english="Time for Completion is 19 months from Commencement Date.",
        beneath_lines="All delay analysis should measure impact against this contractual completion window.",
        leverage_holder="Employer unless EOT proved",
        notice_requirement="EOT claim required for any excusable delay.",
        money_impact="Delay damages risk after expiry without EOT.",
        schedule_impact="Critical path must be shown against 19-month baseline.",
        practical_action="Maintain approved baseline programme and monthly updates."
    ),
    ContractClause(
        clause_id="8.4",
        topic="Extension of time",
        location="Clause 8.4 / 20.1",
        plain_english="EOT is available only for qualifying events and compliant claims.",
        beneath_lines="Even a real Employer-caused delay can be lost if notice, particulars, monthly updates, or critical path proof are missing.",
        leverage_holder="Shared",
        notice_requirement="Strict claim notice and monthly account requirements.",
        money_impact="Prolongation cost depends on entitlement and concurrency rules.",
        schedule_impact="EOT protects against delay damages if proven.",
        practical_action="Maintain delay event register, notices, particulars, updates, and delay analysis."
    ),
    ContractClause(
        clause_id="8.7",
        topic="Delay damages",
        location="Appendix / Clause 8.7",
        plain_english="Delay damages are 1 percent of outstanding works per week or part week, capped at 10 percent of Contract Price.",
        beneath_lines="Even part of a week counts. Employer can use this as strong pressure once completion is late without EOT.",
        leverage_holder="Employer",
        notice_requirement="EOT must be claimed before delay damages crystallize.",
        money_impact="High exposure, capped at 10 percent.",
        schedule_impact="Critical if no approved EOT.",
        practical_action="Update EOT register and forecast exposure weekly."
    ),
    ContractClause(
        clause_id="13.1",
        topic="Right to vary",
        location="Clause 13.1",
        plain_english="Employer/Engineer may instruct variations within the contractual mechanism.",
        beneath_lines="Contractor generally must proceed with instructed variation and argue valuation/time through the procedure, not refuse execution.",
        leverage_holder="Employer",
        notice_requirement="Reservation and proposal must follow Clause 13.3/20.1.",
        money_impact="Potential payment if instructed and valued.",
        schedule_impact="Possible EOT if variation delays critical path.",
        practical_action="Get written instruction, submit proposal, maintain records, and update programme."
    ),
    ContractClause(
        clause_id="13.3",
        topic="Variation procedure",
        location="Clause 13.3",
        plain_english="Variations require notice/proposal procedure, including time and cost impact.",
        beneath_lines="A variation without a timely proposal and monthly claim account may become hard to recover later.",
        leverage_holder="Shared",
        notice_requirement="Very important: submit notice/proposal within required period and include monthly account.",
        money_impact="Preserves extra payment.",
        schedule_impact="Preserves EOT if critical.",
        practical_action="Use a variation register with instruction date, notice date, proposal date, status, and monthly update."
    ),
    ContractClause(
        clause_id="14.7",
        topic="Payment period",
        location="Clause 14.7",
        plain_english="Employer pays amounts due within 45 days after Engineer receives a complete statement and supporting documents.",
        beneath_lines="The 45 days likely starts only after a compliant submission, not after an incomplete invoice.",
        leverage_holder="Employer until submission complete",
        notice_requirement="Late payment arguments need proof of complete submission date.",
        money_impact="Interest or delayed payment arguments depend on compliance.",
        schedule_impact="Suspension rights require strict clause compliance.",
        practical_action="Keep transmittal, Engineer receipt date, checklist of supporting documents, and payment certificate."
    ),
    ContractClause(
        clause_id="20.1",
        topic="Claims time bar",
        location="Clause 20.1",
        plain_english="Failure to comply with claim notice requirements is an irrevocable waiver and release of Employer/Engineer from claims.",
        beneath_lines="This is one of the strongest clauses. A good claim can die if the notice procedure is missed.",
        leverage_holder="Employer",
        notice_requirement="Strict notice, particulars, and monthly claim account.",
        money_impact="Missed notice can erase payment claim.",
        schedule_impact="Missed notice can erase EOT claim.",
        practical_action="Maintain a live claims register and send protective notices early."
    ),
    ]


def _load_clause_library() -> List[ContractClause]:
    if not CLAUSE_LIBRARY_PATH.exists():
        return []

    try:
        if CLAUSE_LIBRARY_PATH.suffix.lower() == ".csv":
            df = pd.read_csv(CLAUSE_LIBRARY_PATH).fillna("")
        else:
            df = pd.read_excel(CLAUSE_LIBRARY_PATH, sheet_name="Clause Library").fillna("")
    except ImportError:
        return []
    except Exception:
        return []

    clauses: List[ContractClause] = []
    for index, row in df.iterrows():
        topic = str(row.get("Clause / Topic", "")).strip()
        if not topic:
            continue
        clauses.append(
            ContractClause(
                clause_id=str(index + 1),
                topic=topic,
                location=str(row.get("Location", "")).strip(),
                plain_english=str(row.get("Plain English Meaning", "")).strip(),
                beneath_lines=str(row.get("Research the Lines", row.get("Beneath the Lines", ""))).strip(),
                leverage_holder=str(row.get("Who Holds Leverage", "")).strip(),
                notice_requirement=str(row.get("Notice / Time Bar", "")).strip(),
                money_impact=str(row.get("Money Impact", "")).strip(),
                schedule_impact=str(row.get("Schedule Impact", "")).strip(),
                practical_action=str(row.get("Practical Action / Evidence", "")).strip(),
            )
        )
    return clauses


CONTRACT_CLAUSES = _load_clause_library()


EVENT_KEYWORDS = {
    "material": ["material", "steel", "delivery", "free-issue", "requisition", "equipment"],
    "variation": ["variation", "change", "instruction", "scope", "extra", "proposal"],
    "payment": ["payment", "invoice", "ipc", "statement", "certified", "cash", "paid"],
    "remeasurement": ["remeasure", "quantity", "measurement", "boq", "survey"],
    "delay": ["delay", "eot", "extension", "time", "completion", "critical", "damages"],
    "notice": ["notice", "claim", "time bar", "particulars", "monthly", "waiver"],
}


def _extract_first_percent(text: str) -> float | None:
    match = re.search(r"(\d+(?:\.\d+)?)\s*percent|(\d+(?:\.\d+)?)\s*%", str(text), re.IGNORECASE)
    if not match:
        return None
    value = match.group(1) or match.group(2)
    return float(value) if value is not None else None


def _extract_first_days(text: str) -> int | None:
    match = re.search(r"(\d+)\s*days?", str(text), re.IGNORECASE)
    return int(match.group(1)) if match else None


def _clause_bundle_text(clause: ContractClause) -> str:
    return " ".join(
        [
            clause.topic,
            clause.location,
            clause.plain_english,
            clause.beneath_lines,
            clause.notice_requirement,
            clause.money_impact,
            clause.schedule_impact,
            clause.practical_action,
        ]
    )


def _score_clause(clause: ContractClause, terms: List[str]) -> int:
    haystack = " ".join(
        [
            clause.topic,
            clause.location,
            clause.plain_english,
            clause.beneath_lines,
            clause.notice_requirement,
            clause.money_impact,
            clause.schedule_impact,
            clause.practical_action,
        ]
    ).lower()
    return sum(3 if term in clause.topic.lower() else 1 for term in terms if term and term in haystack)


def _make_match(clause: ContractClause, score: int) -> Dict[str, Any]:
    relevance = "CRITICAL" if score >= 6 or "time bar" in clause.topic.lower() else ("HIGH" if score >= 3 else "MEDIUM")
    return {
        "clause": clause,
        "relevance": relevance,
        "reason": f"Matched to contract subject: {clause.topic}",
        "action_required": clause.practical_action or clause.notice_requirement,
        "entitlement_risk": "HIGH" if relevance in {"CRITICAL", "HIGH"} else "MEDIUM",
        "time_impact": clause.schedule_impact,
        "cost_impact": clause.money_impact,
    }


def match_event_to_clauses(event_type: str, event_description: str) -> List[Dict[str, Any]]:
    """
    Match a project event to applicable contract clauses.
    
    Args:
        event_type: Type of event (e.g., 'material_delay', 'variation', 'payment_delay')
        event_description: Description of the event
    
    Returns:
        List of applicable clauses with analysis
    """
    combined = f"{event_type} {event_description}".lower()
    terms = [word for word in combined.replace("/", " ").replace("-", " ").split() if len(word) > 2]
    for subject, keywords in EVENT_KEYWORDS.items():
        if subject in combined or any(keyword in combined for keyword in keywords):
            terms.extend(keywords)

    scored = []
    for clause in CONTRACT_CLAUSES:
        score = _score_clause(clause, terms)
        if score:
            scored.append((score, clause))

    scored.sort(key=lambda item: item[0], reverse=True)
    return [_make_match(clause, score) for score, clause in scored[:12]]


def analyze_delay_event(
    event_id: str,
    event_name: str,
    delay_days: int,
    is_contractor_caused: bool,
    is_critical_path: bool,
    contract_value: float = 0.0,
) -> Dict[str, Any]:
    """
    Analyze a delay event against contract clauses.
    
    Args:
        event_id: Event identifier
        event_name: Event name
        delay_days: Number of delay days
        is_contractor_caused: Whether delay is contractor-caused
        is_critical_path: Whether delay affects critical path
    
    Returns:
        Comprehensive delay analysis with entitlements and risks
    """
    
    applicable_clauses = match_event_to_clauses("material_delay", event_name)
    contract_terms = get_contract_terms()
    active_contract_value = float(contract_value or contract_terms.get("contract_value") or 0.0)
    weekly_ld_rate_pct = float(contract_terms.get("weekly_ld_rate_pct") or 0.0)
    delay_damages_cap_pct = float(contract_terms.get("delay_damages_cap_pct") or 0.0)

    weeks = delay_days / 7
    delay_damages_per_week = active_contract_value * (weekly_ld_rate_pct / 100.0) if active_contract_value and weekly_ld_rate_pct else 0.0
    total_delay_damages = min(
        delay_damages_per_week * weeks,
        active_contract_value * (delay_damages_cap_pct / 100.0) if active_contract_value and delay_damages_cap_pct else delay_damages_per_week * weeks,
    )

    if is_contractor_caused:
        entitlement = "NO EOT"
    elif is_critical_path:
        entitlement = "POSSIBLE EOT"
    else:
        entitlement = "TIME IMPACT TO BE PROVEN"

    risk_level = "HIGH" if is_contractor_caused or is_critical_path else "MEDIUM"
    if not applicable_clauses:
        risk_level = "LOW"
    
    return {
        'event_id': event_id,
        'event_name': event_name,
        'delay_days': delay_days,
        'is_contractor_caused': is_contractor_caused,
        'is_critical_path': is_critical_path,
        'entitlement': entitlement,
        'risk_level': risk_level,
        'delay_damages_exposure': total_delay_damages,
        'delay_damages_per_week': delay_damages_per_week,
        'applicable_clauses': applicable_clauses,
        'critical_actions': [
            'Send protective notice immediately (Clause 20.1)',
            'Maintain approved programme and requisitions (Clause 4.20)',
            'Submit monthly claim account (Clause 20.1)',
            'Document critical path impact (Clause 8.2)',
            'Secure EOT before damages crystallize (Clause 8.7)'
        ] if not is_contractor_caused else [
            'Implement recovery plan (Clause 8.6)',
            'Accelerate non-critical activities',
            'Reallocate resources'
        ]
    }


def get_contract_terms(contract_value: float = 0.0) -> Dict[str, Any]:
    delay_clause = next((clause for clause in CONTRACT_CLAUSES if "delay damages" in clause.topic.lower()), None)
    payment_clause = next((clause for clause in CONTRACT_CLAUSES if "payment period" in clause.topic.lower()), None)
    accepted_amount_clause = next((clause for clause in CONTRACT_CLAUSES if "accepted contract amount" in clause.topic.lower()), None)

    weekly_ld_rate_pct = _extract_first_percent(_clause_bundle_text(delay_clause)) if delay_clause else None
    delay_damages_cap_pct = None
    if delay_clause:
        percentages = re.findall(r"(\d+(?:\.\d+)?)\s*percent|(\d+(?:\.\d+)?)\s*%", _clause_bundle_text(delay_clause), re.IGNORECASE)
        extracted = [float(a or b) for a, b in percentages if (a or b)]
        if len(extracted) >= 2:
            weekly_ld_rate_pct = extracted[0]
            delay_damages_cap_pct = extracted[1]
        elif extracted:
            delay_damages_cap_pct = extracted[0]

    payment_period_days = _extract_first_days(_clause_bundle_text(payment_clause)) if payment_clause else None
    resolved_contract_value = float(contract_value or 0.0)
    if not resolved_contract_value and accepted_amount_clause:
        match = re.search(r"EGP\s*([\d,]+(?:\.\d+)?)", _clause_bundle_text(accepted_amount_clause), re.IGNORECASE)
        if match:
            resolved_contract_value = float(match.group(1).replace(",", ""))

    return {
        "contract_value": resolved_contract_value,
        "weekly_ld_rate_pct": weekly_ld_rate_pct,
        "delay_damages_cap_pct": delay_damages_cap_pct,
        "payment_period_days": payment_period_days,
    }


def generate_ai_clause_brief(event_type: str, event_description: str, matches: List[Dict[str, Any]]) -> Dict[str, Any]:
    if not matches:
        return {
            "headline": "No direct clause match found.",
            "risk_level": "LOW",
            "key_notice": "Review the event wording and search the contract library manually.",
            "cost_position": "No clear money position identified from the current input.",
            "time_position": "No clear schedule position identified from the current input.",
            "next_actions": ["Refine the event description with dates, responsibility, and affected activity IDs."],
        }

    top_matches = matches[:3]
    top_clause = top_matches[0]["clause"]
    relevance_levels = [match["relevance"] for match in top_matches]
    risk_level = "CRITICAL" if "CRITICAL" in relevance_levels else ("HIGH" if "HIGH" in relevance_levels else "MEDIUM")

    next_actions: List[str] = []
    seen = set()
    for match in top_matches:
        action = match.get("action_required") or ""
        if action and action not in seen:
            next_actions.append(action)
            seen.add(action)

    return {
        "headline": f"Primary contract anchor: {top_clause.topic}",
        "risk_level": risk_level,
        "key_notice": top_clause.notice_requirement or "Check notice and time-bar requirements in the matched clauses.",
        "cost_position": top_clause.money_impact or "Assess commercial effect from the matched clauses.",
        "time_position": top_clause.schedule_impact or "Assess schedule effect from the matched clauses.",
        "next_actions": next_actions[:5] or ["Review the top matched clauses and issue a protective notice."],
    }


def get_clause_by_id(clause_id: str) -> ContractClause:
    """Get a specific clause by ID."""
    for clause in CONTRACT_CLAUSES:
        if clause.clause_id == clause_id:
            return clause
    return None


def get_all_clauses() -> List[ContractClause]:
    """Get all contract clauses."""
    return CONTRACT_CLAUSES


def search_clauses(keyword: str) -> List[ContractClause]:
    """Search clauses by keyword."""
    keyword_lower = keyword.lower()
    results = []
    for clause in CONTRACT_CLAUSES:
        if (keyword_lower in clause.topic.lower() or
            keyword_lower in clause.plain_english.lower() or
            keyword_lower in clause.beneath_lines.lower()):
            results.append(clause)
    return results
