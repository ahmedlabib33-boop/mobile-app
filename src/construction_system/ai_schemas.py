from __future__ import annotations

from typing import Any


CLAIM_ANSWER_SCHEMA: dict[str, Any] = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "short_answer": {"type": "string"},
        "contractor_friendly_interpretation": {"type": "string"},
        "required_evidence": {"type": "string"},
        "missing_evidence": {"type": "string"},
        "likely_client_rejection": {"type": "string"},
        "contractor_rebuttal": {"type": "string"},
        "recommended_next_action": {"type": "string"},
        "claim_strategy": {"type": "string"},
        "confidence": {"type": "string", "enum": ["Low", "Medium", "High"]},
        "source_limits": {"type": "string"},
    },
    "required": [
        "short_answer",
        "contractor_friendly_interpretation",
        "required_evidence",
        "missing_evidence",
        "likely_client_rejection",
        "contractor_rebuttal",
        "recommended_next_action",
        "claim_strategy",
        "confidence",
        "source_limits",
    ],
}


REBUTTAL_SCHEMA: dict[str, Any] = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "client_argument_summary": {"type": "string"},
        "contractor_counterargument": {"type": "string"},
        "evidence_needed": {"type": "string"},
        "recommended_response_wording": {"type": "string"},
        "probability_of_success": {"type": "string", "enum": ["Low", "Medium", "High"]},
        "confidence": {"type": "string", "enum": ["Low", "Medium", "High"]},
        "source_limits": {"type": "string"},
    },
    "required": [
        "client_argument_summary",
        "contractor_counterargument",
        "evidence_needed",
        "recommended_response_wording",
        "probability_of_success",
        "confidence",
        "source_limits",
    ],
}


CLAIM_DRAFT_SCHEMA: dict[str, Any] = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "narrative_text": {"type": "string"},
        "factual_background": {"type": "string"},
        "cause_effect": {"type": "string"},
        "entitlement_statement": {"type": "string"},
        "time_impact_statement": {"type": "string"},
        "cost_impact_statement": {"type": "string"},
        "rebuttal_section": {"type": "string"},
        "attachment_checklist": {"type": "string"},
        "confidence": {"type": "string", "enum": ["Low", "Medium", "High"]},
        "source_limits": {"type": "string"},
    },
    "required": [
        "narrative_text",
        "factual_background",
        "cause_effect",
        "entitlement_statement",
        "time_impact_statement",
        "cost_impact_statement",
        "rebuttal_section",
        "attachment_checklist",
        "confidence",
        "source_limits",
    ],
}


def required_schema_keys(schema: dict[str, Any]) -> set[str]:
    return {str(key) for key in schema.get("required", [])}
