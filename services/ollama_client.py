from __future__ import annotations

import json
import os
import urllib.error
import urllib.request
from dataclasses import dataclass
from typing import Any


@dataclass
class OllamaResult:
    ok: bool
    text: str = ""
    warning: str = ""
    model: str = ""


class OllamaClient:
    """Small dependency-free Ollama client for local-only reasoning."""

    def __init__(self, base_url: str = "http://localhost:11434", model: str | None = None, timeout: int = 45) -> None:
        self.base_url = base_url.rstrip("/")
        self.model = (model or os.getenv("OLLAMA_MODEL") or "llama3.1:8b").strip()
        self.timeout = timeout

    def _request_json(self, path: str, payload: dict[str, Any] | None = None) -> dict[str, Any]:
        data = None
        headers = {"Accept": "application/json"}
        if payload is not None:
            data = json.dumps(payload).encode("utf-8")
            headers["Content-Type"] = "application/json"
        request = urllib.request.Request(f"{self.base_url}{path}", data=data, headers=headers, method="POST" if payload else "GET")
        with urllib.request.urlopen(request, timeout=self.timeout) as response:
            return json.loads(response.read().decode("utf-8", errors="replace"))

    def available_models(self) -> list[str]:
        try:
            payload = self._request_json("/api/tags")
        except (OSError, urllib.error.URLError, TimeoutError, json.JSONDecodeError):
            return []
        return [str(model.get("name", "")).strip() for model in payload.get("models", []) if model.get("name")]

    def health(self) -> tuple[bool, str]:
        models = self.available_models()
        if not models:
            return False, "Ollama is not reachable at http://localhost:11434."
        if self.model not in models:
            return False, f"Ollama is running, but model '{self.model}' is not installed. Available: {', '.join(models[:8])}."
        return True, f"Ollama model '{self.model}' is available."

    def generate(self, prompt: str, system: str = "", temperature: float = 0.2) -> OllamaResult:
        payload = {
            "model": self.model,
            "prompt": prompt,
            "system": system,
            "stream": False,
            "options": {"temperature": temperature},
        }
        try:
            response = self._request_json("/api/generate", payload)
        except (OSError, urllib.error.URLError, TimeoutError, json.JSONDecodeError) as exc:
            return OllamaResult(ok=False, warning=f"Ollama request failed: {exc}", model=self.model)
        return OllamaResult(ok=True, text=str(response.get("response", "")).strip(), model=self.model)
