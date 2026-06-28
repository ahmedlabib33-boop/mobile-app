from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "src"))

import contract_claims_center as ccc
from construction_system.openai_gateway import create_structured_completion
from construction_system.ai_schemas import CLAIM_ANSWER_SCHEMA, required_schema_keys


def test_contract_answer_uses_local_engine_when_openai_disabled(tmp_path, monkeypatch):
    monkeypatch.delenv("PROJECT_HUB_OPENAI_ENABLED", raising=False)
    monkeypatch.delenv("OPENAI_API_KEY", raising=False)
    db_path = tmp_path / "claims.db"
    ccc.init_contract_claims_db(db_path)

    answer = ccc.answer_contract_question(db_path, "Can we claim EOT for late IFC drawings?")

    assert answer["entitlement_decision"] == "NOT ENOUGH DATA"
    assert answer["ai_status"]["status"] == "skipped"
    assert answer["ai_status"]["source"] == "disabled"


def test_gateway_does_not_call_openai_without_enable_flag(tmp_path, monkeypatch):
    monkeypatch.delenv("PROJECT_HUB_OPENAI_ENABLED", raising=False)
    monkeypatch.setenv("OPENAI_API_KEY", "not-used")
    result = create_structured_completion(
        db_path=tmp_path / "ai.db",
        feature_name="unit_test",
        system_prompt="Return JSON.",
        user_prompt="Return JSON.",
        schema_name="contract_question_answer",
        schema=CLAIM_ANSWER_SCHEMA,
        input_payload={"sample": "value"},
        required_keys=required_schema_keys(CLAIM_ANSWER_SCHEMA),
    )

    assert result.status == "skipped"
    assert result.source == "disabled"
