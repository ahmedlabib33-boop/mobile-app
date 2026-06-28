from __future__ import annotations

from pydantic import BaseModel, Field


class LaunchPlanningRequest(BaseModel):
    product_brief: str = Field(..., min_length=20)
    audience: str = Field(..., min_length=2)
    launch_date: str = Field(..., min_length=4)
    constraints: str = ""
    available_assets: str = ""
    delay_analysis_context: str = ""


class HealthResponse(BaseModel):
    status: str
    openai_key_present: bool
    model: str
