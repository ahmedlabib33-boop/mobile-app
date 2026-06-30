# Problem Solving Engine Status

## Implementation Summary

The Streamlit application now includes a third main phase:

1. Decision Making Dashboard
2. Projects
3. Problem Solving Engine

The new phase is routed from `dashboard.py` to `modules/problem_solving_engine.py` and does not replace the existing dashboard or project modules.

## Created Files

- `modules/problem_solving_engine.py`
- `modules/__init__.py`
- `services/ollama_client.py`
- `services/problem_solving_rag.py`
- `services/legal_research.py`
- `services/evidence_validator.py`
- `services/iterative_solver.py`
- `services/__init__.py`
- `data/problem_solving_engine_question_bank.csv`
- `data/problem_solving_answer_register.csv`
- `storage/problem_solving_history.sqlite`

## Data Rules

- The master question bank is loaded from `data/problem_solving_engine_question_bank.csv`.
- Approved answers are saved separately to `data/problem_solving_answer_register.csv`.
- Solver audit/history records are stored in `storage/problem_solving_history.sqlite`.
- Project evidence search is scoped to the selected project by default.
- Portfolio-wide search is available only through an explicit toggle.
- Public legal/FIDIC search sends search terms only; private project files are not sent online.

## Ollama

The local Ollama endpoint is:

`http://localhost:11434`

The model is read from:

`OLLAMA_MODEL`

Default:

`llama3.1:8b`

If Ollama is offline or the model is unavailable, the app shows a warning and generates a governed fallback answer instead of crashing.

## Verification

- Python compile check passed for `dashboard.py`, `modules`, and `services`.
- Solver smoke test passed with Ollama disabled.
- Local Streamlit HTTP smoke test returned `HTTP 200`.

## Notes

The generated answers are evidence-assisted drafts. Project-specific statements must remain supported by the evidence cards shown in the engine. Legal/FIDIC references are technical/commercial support only and are not legal advice.
