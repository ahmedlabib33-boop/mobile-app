from __future__ import annotations

import json
import os
from collections.abc import AsyncIterator
from typing import Any

from agents import Agent, ModelSettings, Runner, trace

from .schemas import LaunchPlanningRequest
from .tools import (
    check_launch_readiness,
    draft_channel_launch_copy,
    extract_tasks_from_brief,
    generate_owner_checklist,
)


MODEL = os.getenv("ANALYTICS_AGENT_MODEL", os.getenv("OPENAI_MODEL", "gpt-5.5"))

AGENT_INSTRUCTIONS = """
You are analytics, an engineering release-planning agent specialized in data analytics, delay analysis, TIA, CPM schedule impact, claims evidence, and launch readiness.

You help engineering and project-controls teams convert mixed-format inputs into an actionable release plan.
For every run:
1. Use the available tools before the final response.
2. Produce a prioritized plan, risk register, owner checklist, channel-specific launch copy, and follow-up questions.
3. Treat missing dates, missing owners, missing assets, incomplete schedule evidence, missing baseline/update data, and unclear audience as explicit gaps.
4. Preserve uncertainty. Do not invent unavailable facts, activity IDs, costs, dates, or contractual entitlement.
5. Keep the output structured with clear headings and concise bullets.
""".strip()

analytics_agent = Agent(
    name="analytics",
    instructions=AGENT_INSTRUCTIONS,
    model=MODEL,
    model_settings=ModelSettings(max_tokens=1400, verbosity="medium"),
    tools=[
        extract_tasks_from_brief,
        check_launch_readiness,
        generate_owner_checklist,
        draft_channel_launch_copy,
    ],
)


def build_agent_input(payload: LaunchPlanningRequest) -> str:
    return json.dumps(
        {
            "product_brief": payload.product_brief,
            "audience": payload.audience,
            "launch_date": payload.launch_date,
            "constraints": payload.constraints,
            "available_assets": payload.available_assets,
            "delay_analysis_context": payload.delay_analysis_context,
            "required_output": [
                "prioritized plan",
                "risk register",
                "owner checklist",
                "launch copy suggestions",
                "follow-up questions for missing details",
            ],
        },
        ensure_ascii=True,
        indent=2,
    )


def _event_payload(event_type: str, **payload: Any) -> dict[str, Any]:
    return {"type": event_type, **payload}


async def stream_launch_plan(payload: LaunchPlanningRequest) -> AsyncIterator[dict[str, Any]]:
    agent_input = build_agent_input(payload)
    yield _event_payload("status", message="Starting analytics agent", model=MODEL)
    with trace(
        "analytics_launch_planning",
        metadata={
            "audience": payload.audience[:120],
            "launch_date": payload.launch_date,
            "model": MODEL,
        },
    ):
        result = Runner.run_streamed(analytics_agent, input=agent_input, max_turns=8)
        async for event in result.stream_events():
            if event.type == "agent_updated_stream_event":
                yield _event_payload("status", message=f"Agent active: {event.new_agent.name}")
                continue

            if event.type == "run_item_stream_event":
                name = getattr(event, "name", "")
                if name in {"tool_called", "tool_output"}:
                    item = getattr(event, "item", None)
                    yield _event_payload(
                        "tool_progress",
                        name=name,
                        item_type=getattr(item, "type", item.__class__.__name__ if item is not None else ""),
                        message="Tool call completed" if name == "tool_output" else "Tool call started",
                    )
                continue

            if event.type == "raw_response_event":
                data = getattr(event, "data", None)
                data_type = getattr(data, "type", "")
                if data_type == "response.output_text.delta":
                    delta = getattr(data, "delta", "")
                    if delta:
                        yield _event_payload("text_delta", delta=delta)
                elif data_type == "response.completed":
                    yield _event_payload("status", message="Model response completed")

    final_output = getattr(result, "final_output", None)
    if final_output:
        yield _event_payload("final", output=str(final_output))
    yield _event_payload("done", message="Stream complete")
