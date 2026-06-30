from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd

from services.evidence_validator import classify_question, validate_answer, yes
from services.legal_research import DISCLAIMER, public_legal_search
from services.ollama_client import OllamaClient
from services.problem_solving_rag import find_similar_questions, search_project_evidence


ANSWER_HEADINGS = [
    "Direct Answer", "Technical Context", "Project Evidence Found", "Legal / FIDIC / Contract Basis",
    "Root Cause", "Impact on Time / Cost / Quality / HSE / Contract / Maintenance", "Recommended Decision",
    "Owner and Deadline", "Missing Evidence / Risks", "Confidence Score", "Next Action",
]

LEARNING_HEADINGS = [
    "Core Concept", "Why It Matters", "Required Records", "Step-by-Step Workflow", "Checks and Validation",
    "Common Mistakes", "Dashboard Output", "Practical Example", "Confidence / Evidence Need",
]


def _evidence_summary(evidence: list[dict[str, Any]], limit: int = 8) -> str:
    if not evidence:
        return "No project evidence found."
    lines = []
    for item in evidence[:limit]:
        lines.append(
            f"- {Path(str(item.get('source_file', ''))).name} / {item.get('sheet_name', '')} row {item.get('row_number', '')}: "
            f"{item.get('excerpt', '')[:300]}"
        )
    return "\n".join(lines)


def _legal_summary(legal_refs: list[dict[str, Any]]) -> str:
    if not legal_refs:
        return "No legal or FIDIC public reference was retrieved."
    return "\n".join(f"- {item.get('title', '')}: {item.get('url', '')} ({item.get('retrieved_at', '')})" for item in legal_refs[:6])


def _fallback_answer(question: str, evidence: list[dict[str, Any]], legal_refs: list[dict[str, Any]], validation: dict[str, Any], learning: bool) -> str:
    headings = LEARNING_HEADINGS if learning else ANSWER_HEADINGS
    parts: list[str] = []
    for heading in headings:
        if heading in {"Direct Answer", "Core Concept"}:
            text = "A defensible answer requires the listed evidence and validation checks. Based on the current search, use the evidence below and treat missing items as open actions."
        elif heading in {"Project Evidence Found", "Required Records"}:
            text = _evidence_summary(evidence)
        elif heading in {"Legal / FIDIC / Contract Basis"}:
            text = _legal_summary(legal_refs) + f"\n\n{DISCLAIMER}"
        elif heading in {"Missing Evidence / Risks", "Confidence / Evidence Need"}:
            missing = validation.get("missing_evidence", []) or ["No critical missing evidence detected by the current validator."]
            text = "\n".join(f"- {item}" for item in missing)
        elif heading == "Confidence Score":
            text = f"{validation.get('confidence', 'Low')} ({validation.get('score', 0)}/100)"
        elif heading in {"Recommended Decision", "Next Action"}:
            text = "Assign an owner to close missing evidence, review the cited records, and approve the answer only after unresolved high-risk gaps are cleared."
        else:
            text = "To be verified against the selected project records and approved by the responsible reviewer."
        parts.append(f"### {heading}\n{text}")
    return "\n\n".join(parts)


def _ollama_answer(
    *,
    question: str,
    evidence: list[dict[str, Any]],
    legal_refs: list[dict[str, Any]],
    validation: dict[str, Any],
    learning: bool,
    enabled: bool,
) -> tuple[str, str]:
    if not enabled:
        return _fallback_answer(question, evidence, legal_refs, validation, learning), "Local Ollama disabled; deterministic governed answer generated."
    client = OllamaClient()
    ok, health = client.health()
    if not ok:
        return _fallback_answer(question, evidence, legal_refs, validation, learning), health
    headings = LEARNING_HEADINGS if learning else ANSWER_HEADINGS
    prompt = f"""
Question:
{question}

Use exactly these headings:
{json.dumps(headings)}

Project evidence available:
{_evidence_summary(evidence, limit=10)}

Legal/FIDIC public references:
{_legal_summary(legal_refs)}

Validation status:
{json.dumps(validation, ensure_ascii=False)}

Rules:
- Do not invent project evidence.
- Every project-specific statement must refer to the listed evidence or say "Required / To be verified".
- Every legal or FIDIC point must cite the public reference list or say "Unverified".
- End with a practical owner/deadline action.
"""
    result = client.generate(
        prompt,
        system="You are a construction project controls and claims analyst. Be concise, evidence-based, and management-ready.",
        temperature=0.15,
    )
    if not result.ok or not result.text:
        return _fallback_answer(question, evidence, legal_refs, validation, learning), result.warning or "Ollama returned no answer."
    return result.text, f"Ollama generated with {result.model}."


def run_iterative_solver(
    *,
    question: str,
    question_row: dict[str, Any] | None,
    question_bank: pd.DataFrame,
    project_root: Path,
    project_id: str,
    project_name: str,
    portfolio: bool,
    local_ollama: bool,
    web_research: bool,
    fidic_check: bool,
    egypt_law_check: bool,
    max_iterations: int = 3,
) -> dict[str, Any]:
    max_iterations = max(1, min(5, int(max_iterations or 3)))
    question_row = question_row or {}
    require_project_data = yes(question_row.get("project_data_required")) or not question_row
    learning = "learning" in str(question_row.get("source_layer", "")).casefold() or "know" in question.casefold()
    classification = classify_question(question, question_row)
    similar = find_similar_questions(question_bank, question)
    extra_keywords = str(question_row.get("search_keywords", ""))
    legal_refs: list[dict[str, Any]] = []
    if web_research and (fidic_check or egypt_law_check):
        legal_refs = public_legal_search(question, fidic=fidic_check, egypt_law=egypt_law_check)

    evidence: list[dict[str, Any]] = []
    iteration_log: list[dict[str, Any]] = []
    answer = ""
    validation: dict[str, Any] = {}
    warning = ""
    confidence_before = "Low"
    for iteration in range(1, max_iterations + 1):
        needed = " ".join(validation.get("missing_evidence", [])) if validation else ""
        evidence = search_project_evidence(project_root, question, f"{extra_keywords} {needed}", max_results=30, portfolio=portfolio)
        validation = validate_answer(
            question=question,
            answer=answer,
            evidence=evidence,
            legal_refs=legal_refs,
            require_project_data=require_project_data,
            fidic_enabled=fidic_check,
            egypt_law_enabled=egypt_law_check,
        )
        answer, warning = _ollama_answer(
            question=question,
            evidence=evidence,
            legal_refs=legal_refs,
            validation=validation,
            learning=learning,
            enabled=local_ollama,
        )
        validation = validate_answer(
            question=question,
            answer=answer,
            evidence=evidence,
            legal_refs=legal_refs,
            require_project_data=require_project_data,
            fidic_enabled=fidic_check,
            egypt_law_enabled=egypt_law_check,
        )
        iteration_log.append({
            "iteration": iteration,
            "evidence_found": len(evidence),
            "missing_evidence": "; ".join(validation.get("missing_evidence", [])),
            "confidence_before": confidence_before,
            "confidence_after": validation.get("confidence", "Low"),
            "answer_changes": "Generated or refreshed answer against latest evidence.",
            "stop_reason": "High confidence reached." if validation.get("confidence") == "High" else "Continue until max iterations or high confidence.",
        })
        confidence_before = validation.get("confidence", "Low")
        if validation.get("confidence") == "High":
            break

    created_at = datetime.utcnow().isoformat(timespec="seconds") + "Z"
    return {
        "answer_id": f"PSE-{datetime.utcnow().strftime('%Y%m%d%H%M%S')}",
        "question_id": str(question_row.get("question_id", "CUSTOM") or "CUSTOM"),
        "project_id": project_id,
        "project_name": project_name,
        "question_text": question,
        "classification": classification,
        "similar_questions": similar.to_dict("records"),
        "evidence": evidence,
        "legal_refs": legal_refs,
        "generated_answer": answer,
        "confidence": validation.get("confidence", "Low"),
        "score": validation.get("score", 0),
        "validation": validation,
        "iteration_log": iteration_log,
        "iteration_count": len(iteration_log),
        "warning": warning,
        "created_at": created_at,
    }
