from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "src"))

from construction_system.ai_schemas import CLAIM_ANSWER_SCHEMA, required_schema_keys
from construction_system.openai_gateway import create_structured_completion


def main() -> int:
    parser = argparse.ArgumentParser(description="Smoke-check the Project Intelligence Hub OpenAI gateway.")
    parser.add_argument("--live", action="store_true", help="Run a real OpenAI request. Requires OPENAI_API_KEY.")
    args = parser.parse_args()

    if not args.live:
        os.environ.pop("PROJECT_HUB_OPENAI_ENABLED", None)
        os.environ.pop("OPENAI_API_KEY", None)
    elif not os.getenv("OPENAI_API_KEY"):
        print("OPENAI_API_KEY is required for --live.")
        return 2
    else:
        os.environ["PROJECT_HUB_OPENAI_ENABLED"] = "1"

    result = create_structured_completion(
        db_path=ROOT / "construction_system.db",
        feature_name="smoke_gateway",
        system_prompt="You are a concise JSON-only assistant.",
        user_prompt="Return a conservative construction claim answer for a missing notice scenario.",
        schema_name="contract_question_answer",
        schema=CLAIM_ANSWER_SCHEMA,
        input_payload={"scenario": "missing notice"},
        required_keys=required_schema_keys(CLAIM_ANSWER_SCHEMA),
    )
    print(
        {
            "status": result.status,
            "source": result.source,
            "model": result.model,
            "latency_ms": result.latency_ms,
            "error": result.error,
        }
    )
    return 0 if result.status in {"ok", "skipped"} else 1


if __name__ == "__main__":
    raise SystemExit(main())
