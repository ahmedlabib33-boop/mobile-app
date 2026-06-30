from __future__ import annotations

import csv
import json
import re
import sqlite3
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd


QUESTION_BANK_COLUMNS = [
    "question_id", "source_layer", "department", "section_or_level", "question_text", "answer_status",
    "answer_column", "dropdown_module", "destination_view", "local_ollama_enabled", "web_research_enabled",
    "fidic_check_enabled", "egypt_law_check_enabled", "project_data_required", "evidence_fields_required",
    "response_template", "search_keywords", "source_file",
]

ANSWER_REGISTER_COLUMNS = [
    "answer_id", "question_id", "project_id", "project_name", "question_text", "generated_answer",
    "approved_answer", "evidence_refs", "legal_refs", "confidence", "status", "owner", "deadline",
    "created_at", "updated_at", "iteration_count", "user_notes",
]

EXCLUDED_DIRS = {
    ".venv", "__pycache__", ".pytest_cache", ".pih_mobile_app_sync_state", "_RETURN_POINTS",
    "dist", "android_app", "mobile_app", "ios_app_source", ".streamlit",
}
SEARCH_EXTENSIONS = {".csv", ".xlsx", ".xls", ".json", ".txt", ".md"}


def ensure_csv(path: Path, columns: list[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    if not path.exists():
        path.write_text(",".join(columns) + "\n", encoding="utf-8")


def load_question_bank(path: Path) -> pd.DataFrame:
    ensure_csv(path, QUESTION_BANK_COLUMNS)
    frame = pd.read_csv(path, dtype=str).fillna("")
    for column in QUESTION_BANK_COLUMNS:
        if column not in frame.columns:
            frame[column] = ""
    return frame[QUESTION_BANK_COLUMNS]


def load_answer_register(path: Path) -> pd.DataFrame:
    ensure_csv(path, ANSWER_REGISTER_COLUMNS)
    frame = pd.read_csv(path, dtype=str).fillna("")
    for column in ANSWER_REGISTER_COLUMNS:
        if column not in frame.columns:
            frame[column] = ""
    return frame[ANSWER_REGISTER_COLUMNS]


def normalize_columns(frame: pd.DataFrame) -> pd.DataFrame:
    updated = frame.copy()
    updated.columns = [re.sub(r"[^a-z0-9]+", "_", str(column).strip().lower()).strip("_") or "field" for column in updated.columns]
    return updated.dropna(how="all").drop_duplicates()


def keywords_from_text(text: str, extra: str = "") -> list[str]:
    tokens = re.findall(r"[A-Za-z0-9][A-Za-z0-9_-]{2,}", f"{text} {extra}".casefold())
    stop = {"the", "and", "for", "with", "from", "this", "that", "what", "which", "when", "where", "shall", "project"}
    seen: set[str] = set()
    keywords: list[str] = []
    for token in tokens:
        if token in stop or token in seen:
            continue
        seen.add(token)
        keywords.append(token)
    return keywords[:24]


def find_similar_questions(question_bank: pd.DataFrame, question: str, limit: int = 8) -> pd.DataFrame:
    if question_bank.empty:
        return question_bank
    keywords = set(keywords_from_text(question))
    if not keywords:
        return question_bank.head(limit)
    scored: list[tuple[int, int]] = []
    for idx, row in question_bank.iterrows():
        haystack = " ".join(str(row.get(column, "")) for column in ("question_text", "department", "section_or_level", "search_keywords")).casefold()
        score = sum(1 for keyword in keywords if keyword in haystack)
        if score:
            scored.append((score, idx))
    indexes = [idx for _, idx in sorted(scored, reverse=True)[:limit]]
    return question_bank.loc[indexes] if indexes else question_bank.head(limit)


def _safe_read_csv(path: Path) -> pd.DataFrame:
    for encoding in ("utf-8-sig", "utf-8", "cp1256", "latin1"):
        try:
            return pd.read_csv(path, dtype=str, encoding=encoding).fillna("")
        except Exception:
            continue
    return pd.DataFrame()


def _iter_search_files(root: Path, portfolio: bool = False) -> list[Path]:
    if not root.exists():
        return []
    files: list[Path] = []
    for candidate in root.rglob("*"):
        if any(part in EXCLUDED_DIRS for part in candidate.parts):
            continue
        if candidate.is_file() and candidate.suffix.lower() in SEARCH_EXTENSIONS:
            files.append(candidate)
    return files if portfolio else files[:700]


def _record_from_row(path: Path, sheet: str, row_number: int, row: pd.Series, keywords: list[str]) -> dict[str, Any] | None:
    row_text = " | ".join(str(value) for value in row.dropna().astype(str).tolist())
    lowered = row_text.casefold()
    matches = [keyword for keyword in keywords if keyword in lowered]
    if not matches:
        return None
    columns = {str(key).casefold(): value for key, value in row.to_dict().items()}
    def first(names: list[str]) -> str:
        for name in names:
            for key, value in columns.items():
                if name in key and str(value).strip():
                    return str(value).strip()
        return ""
    return {
        "source_file": str(path),
        "sheet_name": sheet,
        "row_number": row_number,
        "matched_fields": ", ".join(matches[:8]),
        "date": first(["date", "start", "finish"]),
        "activity_id": first(["activity_id", "activity id", "act_id", "id"]),
        "owner": first(["owner", "responsible", "engineer", "contractor"]),
        "status": first(["status", "state"]),
        "excerpt": row_text[:900],
        "confidence": min(95, 45 + len(matches) * 8),
    }


def search_project_evidence(root: Path, question: str, extra_keywords: str = "", max_results: int = 30, portfolio: bool = False) -> list[dict[str, Any]]:
    keywords = keywords_from_text(question, extra_keywords)
    if not keywords:
        return []
    results: list[dict[str, Any]] = []
    for path in _iter_search_files(root, portfolio=portfolio):
        try:
            if path.suffix.lower() == ".csv":
                sheets = {"CSV": normalize_columns(_safe_read_csv(path))}
            elif path.suffix.lower() in {".xlsx", ".xls"}:
                sheets = {name: normalize_columns(frame.fillna("")) for name, frame in pd.read_excel(path, sheet_name=None, dtype=str).items()}
            elif path.suffix.lower() == ".json":
                payload = json.loads(path.read_text(encoding="utf-8", errors="replace"))
                sheets = {"JSON": normalize_columns(pd.json_normalize(payload if isinstance(payload, list) else [payload]).fillna(""))}
            else:
                lines = path.read_text(encoding="utf-8", errors="replace").splitlines()
                for row_number, line in enumerate(lines, start=1):
                    lowered = line.casefold()
                    matches = [keyword for keyword in keywords if keyword in lowered]
                    if matches:
                        results.append({
                            "source_file": str(path), "sheet_name": "Text", "row_number": row_number,
                            "matched_fields": ", ".join(matches[:8]), "date": "", "activity_id": "",
                            "owner": "", "status": "", "excerpt": line[:900], "confidence": min(90, 45 + len(matches) * 8),
                        })
                        if len(results) >= max_results:
                            return results
                continue
            for sheet_name, frame in sheets.items():
                if frame.empty:
                    continue
                for idx, row in frame.head(5000).iterrows():
                    record = _record_from_row(path, sheet_name, int(idx) + 2, row, keywords)
                    if record:
                        results.append(record)
                        if len(results) >= max_results:
                            return results
        except Exception as exc:
            results.append({
                "source_file": str(path), "sheet_name": "Read warning", "row_number": 0,
                "matched_fields": "", "date": "", "activity_id": "", "owner": "", "status": "warning",
                "excerpt": f"Could not search file: {exc}", "confidence": 0,
            })
    return sorted(results, key=lambda item: int(item.get("confidence", 0)), reverse=True)[:max_results]


def append_answer(register_path: Path, record: dict[str, Any]) -> None:
    ensure_csv(register_path, ANSWER_REGISTER_COLUMNS)
    row = {column: str(record.get(column, "") or "").replace("\r", " ").replace("\n", "\\n") for column in ANSWER_REGISTER_COLUMNS}
    with register_path.open("a", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=ANSWER_REGISTER_COLUMNS)
        writer.writerow(row)


def init_history_db(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with sqlite3.connect(path) as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS problem_solving_history (
                answer_id TEXT PRIMARY KEY,
                project_id TEXT,
                question_id TEXT,
                question_text TEXT,
                confidence TEXT,
                score INTEGER,
                payload_json TEXT,
                created_at TEXT
            )
            """
        )


def append_history(path: Path, record: dict[str, Any]) -> None:
    init_history_db(path)
    payload = json.dumps(record, ensure_ascii=False, default=str)
    with sqlite3.connect(path) as conn:
        conn.execute(
            """
            INSERT OR REPLACE INTO problem_solving_history
            (answer_id, project_id, question_id, question_text, confidence, score, payload_json, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                record.get("answer_id", ""),
                record.get("project_id", ""),
                record.get("question_id", ""),
                record.get("question_text", ""),
                record.get("confidence", ""),
                int(record.get("score", 0) or 0),
                payload,
                record.get("created_at", datetime.utcnow().isoformat(timespec="seconds") + "Z"),
            ),
        )
