from __future__ import annotations

import html
import re
import urllib.parse
import urllib.request
from datetime import datetime
from typing import Any


DISCLAIMER = "Technical/commercial support only. This is not legal advice and must be reviewed by qualified counsel."


def _strip_tags(value: str) -> str:
    return re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", html.unescape(value))).strip()


def public_legal_search(question: str, *, fidic: bool, egypt_law: bool, max_results: int = 5) -> list[dict[str, Any]]:
    """Search only public terms. No project documents or private file content are sent online."""
    prefixes: list[str] = []
    if fidic:
        prefixes.append("FIDIC construction contract clause guidance")
    if egypt_law:
        prefixes.append("Egypt construction law public authority requirements")
    if not prefixes:
        return []

    query = " OR ".join(prefixes) + " " + " ".join(re.findall(r"[A-Za-z0-9]{4,}", question)[:8])
    url = "https://duckduckgo.com/html/?" + urllib.parse.urlencode({"q": query})
    request = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 ProjectIntelligenceHub/1.0"})
    retrieved_at = datetime.utcnow().isoformat(timespec="seconds") + "Z"
    try:
        with urllib.request.urlopen(request, timeout=12) as response:
            body = response.read().decode("utf-8", errors="replace")
    except Exception as exc:
        return [{
            "title": "Public legal research unavailable",
            "url": "",
            "snippet": f"Search failed: {exc}",
            "retrieved_at": retrieved_at,
            "source_type": "warning",
        }]

    results: list[dict[str, Any]] = []
    blocks = re.findall(r'<a rel="nofollow" class="result__a" href="(?P<url>.*?)".*?>(?P<title>.*?)</a>.*?<a class="result__snippet".*?>(?P<snippet>.*?)</a>', body, re.S)
    for raw_url, title, snippet in blocks[:max_results]:
        results.append({
            "title": _strip_tags(title),
            "url": html.unescape(raw_url),
            "snippet": _strip_tags(snippet),
            "retrieved_at": retrieved_at,
            "source_type": "FIDIC/Egyptian Law public search",
        })
    if not results:
        results.append({
            "title": "No public reference found",
            "url": "",
            "snippet": "No searchable public reference was returned for the selected terms.",
            "retrieved_at": retrieved_at,
            "source_type": "warning",
        })
    return results
