from __future__ import annotations

import json
import sys
from pathlib import Path

import httpx


SAMPLE_PAYLOAD = {
    "product_brief": (
        "Launch an engineering analytics release that converts XER schedules, Excel delay registers, "
        "claim PDFs, and meeting notes into a prioritized release plan for delay analysis and management decisions."
    ),
    "audience": "Engineering directors, planning engineers, project controls, and contracts team",
    "launch_date": "2026-07-15",
    "constraints": "Need to preserve contractual uncertainty, no invented dates, limited validated cost records, critical path must be checked.",
    "available_assets": "Primavera XER baseline, monthly update XML, delay event Excel register, claim PDF, owner correspondence notes",
    "delay_analysis_context": "Potential steel delivery delay, late IFC drawings, and concurrent MEP constraints need TIA screening.",
}


def main() -> int:
    url = sys.argv[1] if len(sys.argv) > 1 else "http://127.0.0.1:8788/api/agent/stream"
    saw_tool = False
    saw_delta = False
    events: list[dict] = []
    with httpx.Client(timeout=90) as client:
        with client.stream("POST", url, json=SAMPLE_PAYLOAD) as response:
            response.raise_for_status()
            for line in response.iter_lines():
                if not line.startswith("data: "):
                    continue
                event = json.loads(line[6:])
                events.append(event)
                if event.get("type") == "tool_progress":
                    saw_tool = True
                if event.get("type") == "text_delta" and event.get("delta"):
                    saw_delta = True
                if event.get("type") == "error":
                    print(json.dumps({"status": "error", "event": event}, indent=2))
                    return 1
                if saw_tool and saw_delta:
                    print(json.dumps({"status": "ok", "saw_tool_progress": True, "saw_text_delta": True, "events_seen": len(events)}, indent=2))
                    return 0
    print(json.dumps({"status": "failed", "saw_tool_progress": saw_tool, "saw_text_delta": saw_delta, "events_seen": len(events)}, indent=2))
    Path("last_stream_events.json").write_text(json.dumps(events, indent=2), encoding="utf-8")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
