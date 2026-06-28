from __future__ import annotations

import json
import os
from collections.abc import AsyncIterator
from pathlib import Path

from dotenv import load_dotenv
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse

from .schemas import HealthResponse, LaunchPlanningRequest


ROOT = Path(__file__).resolve().parents[3]
load_dotenv(ROOT / ".env.local")
load_dotenv(ROOT / ".env")

from .agent import MODEL, stream_launch_plan

app = FastAPI(title="analytics agent API", version="1.0.0")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://127.0.0.1:5177", "http://localhost:5177"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.get("/api/health", response_model=HealthResponse)
async def health() -> HealthResponse:
    return HealthResponse(
        status="ok",
        openai_key_present=bool(os.getenv("OPENAI_API_KEY")),
        model=MODEL,
    )


def _sse(data: dict) -> str:
    return f"data: {json.dumps(data, ensure_ascii=True)}\n\n"


@app.post("/api/agent/stream")
async def stream_agent(payload: LaunchPlanningRequest) -> StreamingResponse:
    async def generate() -> AsyncIterator[str]:
        try:
            async for event in stream_launch_plan(payload):
                yield _sse(event)
        except Exception as exc:
            yield _sse({"type": "error", "message": f"{exc.__class__.__name__}: {exc}"})

    return StreamingResponse(generate(), media_type="text/event-stream")
