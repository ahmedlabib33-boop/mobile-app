from __future__ import annotations

import hashlib
import json
import os
import sqlite3
import time
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


DEFAULT_MODEL = "gpt-5.5"


@dataclass(frozen=True)
class AISettings:
    enabled: bool
    api_key: str
    model: str
    reasoning_effort: str
    timeout_seconds: float
    max_retries: int
    max_output_tokens: int
    prompt_version: str


@dataclass(frozen=True)
class AIResult:
    status: str
    data: dict[str, Any] | None = None
    source: str = "none"
    error: str = ""
    cache_key: str = ""
    model: str = ""
    latency_ms: int = 0

    @property
    def ok(self) -> bool:
        return self.status == "ok" and isinstance(self.data, dict)


def now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def load_ai_settings() -> AISettings:
    enabled_value = os.getenv("PROJECT_HUB_OPENAI_ENABLED", "").strip().lower()
    enabled = enabled_value in {"1", "true", "yes", "on"}
    return AISettings(
        enabled=enabled,
        api_key=os.getenv("OPENAI_API_KEY", "").strip(),
        model=os.getenv("PROJECT_HUB_OPENAI_MODEL", DEFAULT_MODEL).strip() or DEFAULT_MODEL,
        reasoning_effort=os.getenv("PROJECT_HUB_OPENAI_REASONING_EFFORT", "low").strip() or "low",
        timeout_seconds=float(os.getenv("PROJECT_HUB_OPENAI_TIMEOUT", "30")),
        max_retries=int(os.getenv("PROJECT_HUB_OPENAI_MAX_RETRIES", "2")),
        max_output_tokens=int(os.getenv("PROJECT_HUB_OPENAI_MAX_OUTPUT_TOKENS", "900")),
        prompt_version=os.getenv("PROJECT_HUB_OPENAI_PROMPT_VERSION", "contract-ai-v1").strip() or "contract-ai-v1",
    )


def is_openai_ready() -> bool:
    settings = load_ai_settings()
    return settings.enabled and bool(settings.api_key)


def stable_json_hash(payload: dict[str, Any]) -> str:
    normalized = json.dumps(payload, sort_keys=True, ensure_ascii=True, default=str, separators=(",", ":"))
    return hashlib.sha256(normalized.encode("utf-8")).hexdigest()


def init_ai_runs_table(db_path: Path) -> None:
    conn = sqlite3.connect(str(db_path))
    try:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS ai_runs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                cache_key TEXT UNIQUE,
                feature_name TEXT NOT NULL,
                prompt_version TEXT NOT NULL,
                model TEXT,
                status TEXT NOT NULL,
                latency_ms INTEGER,
                input_hash TEXT,
                output_json TEXT,
                error_message TEXT,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            )
            """
        )
        conn.commit()
    finally:
        conn.close()


def _read_cached_result(db_path: Path, cache_key: str) -> dict[str, Any] | None:
    init_ai_runs_table(db_path)
    conn = sqlite3.connect(str(db_path))
    try:
        row = conn.execute(
            "SELECT output_json FROM ai_runs WHERE cache_key = ? AND status = 'ok' AND output_json IS NOT NULL",
            (cache_key,),
        ).fetchone()
    finally:
        conn.close()
    if not row:
        return None
    try:
        payload = json.loads(row[0])
    except (TypeError, json.JSONDecodeError):
        return None
    return payload if isinstance(payload, dict) else None


def _write_run(
    db_path: Path,
    cache_key: str,
    feature_name: str,
    settings: AISettings,
    status: str,
    input_hash: str,
    output: dict[str, Any] | None,
    error: str,
    latency_ms: int,
) -> None:
    init_ai_runs_table(db_path)
    stamp = now_iso()
    conn = sqlite3.connect(str(db_path))
    try:
        conn.execute(
            """
            INSERT INTO ai_runs (
                cache_key, feature_name, prompt_version, model, status, latency_ms,
                input_hash, output_json, error_message, created_at, updated_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(cache_key) DO UPDATE SET
                status = excluded.status,
                latency_ms = excluded.latency_ms,
                output_json = excluded.output_json,
                error_message = excluded.error_message,
                updated_at = excluded.updated_at
            """,
            (
                cache_key,
                feature_name,
                settings.prompt_version,
                settings.model,
                status,
                latency_ms,
                input_hash,
                json.dumps(output, ensure_ascii=True, default=str) if output is not None else None,
                error[:1000],
                stamp,
                stamp,
            ),
        )
        conn.commit()
    finally:
        conn.close()


def _validate_payload(payload: Any, required_keys: set[str]) -> dict[str, Any] | None:
    if not isinstance(payload, dict):
        return None
    if not required_keys.issubset(set(payload.keys())):
        return None
    return payload


def _call_openai(
    settings: AISettings,
    system_prompt: str,
    user_prompt: str,
    schema_name: str,
    schema: dict[str, Any],
) -> dict[str, Any]:
    try:
        from openai import OpenAI
    except ImportError as exc:
        raise RuntimeError("OpenAI SDK is not installed. Add 'openai' to requirements and install dependencies.") from exc

    client = OpenAI(api_key=settings.api_key, timeout=settings.timeout_seconds, max_retries=settings.max_retries)
    response = client.responses.create(
        model=settings.model,
        input=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
        reasoning={"effort": settings.reasoning_effort},
        max_output_tokens=settings.max_output_tokens,
        text={
            "format": {
                "type": "json_schema",
                "name": schema_name,
                "schema": schema,
                "strict": True,
            }
        },
    )
    raw_text = getattr(response, "output_text", "") or ""
    if not raw_text:
        raise RuntimeError("OpenAI returned an empty response.")
    return json.loads(raw_text)


def create_structured_completion(
    db_path: Path,
    feature_name: str,
    system_prompt: str,
    user_prompt: str,
    schema_name: str,
    schema: dict[str, Any],
    input_payload: dict[str, Any],
    required_keys: set[str],
) -> AIResult:
    settings = load_ai_settings()
    input_hash = stable_json_hash(
        {
            "feature_name": feature_name,
            "prompt_version": settings.prompt_version,
            "model": settings.model,
            "input": input_payload,
        }
    )
    cache_key = f"{feature_name}:{input_hash}"

    if not settings.enabled:
        return AIResult(status="skipped", source="disabled", cache_key=cache_key, model=settings.model)
    if not settings.api_key:
        return AIResult(status="skipped", source="missing_api_key", cache_key=cache_key, model=settings.model)

    cached = _read_cached_result(db_path, cache_key)
    if cached is not None:
        return AIResult(status="ok", data=cached, source="cache", cache_key=cache_key, model=settings.model)

    started = time.perf_counter()
    try:
        payload = _call_openai(settings, system_prompt, user_prompt, schema_name, schema)
        validated = _validate_payload(payload, required_keys)
        if validated is None:
            raise RuntimeError("OpenAI response did not match the required schema keys.")
        latency_ms = int((time.perf_counter() - started) * 1000)
        _write_run(db_path, cache_key, feature_name, settings, "ok", input_hash, validated, "", latency_ms)
        return AIResult(status="ok", data=validated, source="openai", cache_key=cache_key, model=settings.model, latency_ms=latency_ms)
    except Exception as exc:
        latency_ms = int((time.perf_counter() - started) * 1000)
        error = f"{exc.__class__.__name__}: {exc}"
        _write_run(db_path, cache_key, feature_name, settings, "error", input_hash, None, error, latency_ms)
        return AIResult(status="error", source="openai", error=error, cache_key=cache_key, model=settings.model, latency_ms=latency_ms)
