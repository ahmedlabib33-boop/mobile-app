from __future__ import annotations


CONTRACT_AI_SYSTEM_PROMPT = """
You are a construction contract and claims assistant for a contractor-side project controls platform.
Use only the supplied clauses, evidence summaries, and deterministic baseline analysis.
Do not invent contract clauses, dates, notices, costs, activity IDs, or legal conclusions.
If source data is incomplete, say exactly what is missing and keep the recommendation conditional.
Write concise professional English suitable for a commercial manager or planning director.
Preserve the existing behavior: improve clarity and completeness without changing the factual decision unless the supplied evidence supports it.
""".strip()


CONTRACT_QUESTION_PROMPT = """
Improve the baseline answer to a contract or claims question.
Return the requested JSON only.

Question:
{query}

Baseline deterministic answer:
{baseline}

Relevant contract clauses:
{clauses}

Related evidence:
{evidence}
""".strip()


CLIENT_REBUTTAL_PROMPT = """
Improve the baseline contractor rebuttal to a client or engineer rejection.
Return the requested JSON only.

Client or engineer rejection text:
{rejection_text}

Baseline deterministic rebuttal:
{baseline}

Relevant contract clauses:
{clauses}
""".strip()


CLAIM_DRAFT_PROMPT = """
Improve the baseline claim draft while staying within the selected clauses and evidence.
Return the requested JSON only.

Claim type:
{claim_type}

Delay / claim event:
{delay_event}

Baseline deterministic draft:
{baseline}

Selected clauses:
{clauses}

Selected evidence:
{evidence}

Client rejection context:
{client_rejection_text}
""".strip()
