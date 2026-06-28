from __future__ import annotations

import csv
import hashlib
import importlib
import io
import json
import re
import shutil
import sqlite3
import subprocess
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

import pandas as pd

from src.construction_system.ai_prompts import (
    CLAIM_DRAFT_PROMPT,
    CLIENT_REBUTTAL_PROMPT,
    CONTRACT_AI_SYSTEM_PROMPT,
    CONTRACT_QUESTION_PROMPT,
)
from src.construction_system.ai_schemas import (
    CLAIM_ANSWER_SCHEMA,
    CLAIM_DRAFT_SCHEMA,
    REBUTTAL_SCHEMA,
    required_schema_keys,
)
from src.construction_system.openai_gateway import create_structured_completion


SUPPORTED_CONTRACT_EXTENSIONS = {".pdf", ".docx", ".txt", ".csv", ".xls", ".xlsx"}
SUPPORTED_EVIDENCE_EXTENSIONS = SUPPORTED_CONTRACT_EXTENSIONS | {".eml", ".msg", ".html"}


CLAIM_SECTION_RULES = [
    {
        "section_name": "Contractor Rights",
        "keywords": ["contractor shall be entitled", "contractor may", "entitled to", "right to", "the contractor shall have the right"],
        "claim_type": "General Contractor Entitlement",
        "required_evidence": "Relevant contract notice, event chronology, supporting project records.",
        "possible_client_rejection": "No entitlement because the event does not fall under the cited clause.",
        "counterargument": "Demonstrate the event wording matches the entitlement trigger and support it with contemporaneous records.",
    },
    {
        "section_name": "Employer Obligations",
        "keywords": ["employer shall", "the employer shall", "provided by the employer", "employer must"],
        "claim_type": "Employer Obligation Breach",
        "required_evidence": "Employer instruction / omission record, correspondence, and resulting project impact evidence.",
        "possible_client_rejection": "The obligation was fulfilled or was outside the employer's contractual scope.",
        "counterargument": "Point to the exact contractual wording and the contemporaneous records showing the obligation was not met.",
    },
    {
        "section_name": "Engineer Obligations",
        "keywords": ["engineer shall", "the engineer shall", "engineer may instruct", "approval by the engineer"],
        "claim_type": "Engineer Instruction / Approval Delay",
        "required_evidence": "Submittals, RFIs, engineer replies, meeting minutes, and schedule effect evidence.",
        "possible_client_rejection": "Engineer actions do not create entitlement without further proof of impact.",
        "counterargument": "Show the delayed engineer action blocked critical or near-critical work and triggered contractual relief.",
    },
    {
        "section_name": "Claim Triggers",
        "keywords": ["if the contractor suffers", "in the event of", "if the contractor incurs", "as a result of"],
        "claim_type": "Claim Trigger",
        "required_evidence": "Trigger event record, cause-and-effect logic, notice, and schedule/commercial substantiation.",
        "possible_client_rejection": "The factual trigger did not occur or was not properly evidenced.",
        "counterargument": "Use the trigger language, event chronology, and project records to prove the trigger did occur.",
    },
    {
        "section_name": "EOT Entitlements",
        "keywords": ["extension of time", "eot", "extension to the time for completion", "time for completion shall be extended"],
        "claim_type": "Extension of Time",
        "required_evidence": "Accepted programme basis, event window, contemporaneous records, CPM logic, notices.",
        "possible_client_rejection": "Delay is not critical, not on longest path, or not proven to affect completion.",
        "counterargument": "Support the claim with TIA / window logic, path analysis, and event-specific schedule evidence.",
    },
    {
        "section_name": "Variation Entitlements",
        "keywords": ["variation", "change in the works", "instruction", "valuation of variations", "additional work"],
        "claim_type": "Variation / Change",
        "required_evidence": "Instruction, design change, measured quantities, pricing basis, time/cost impact.",
        "possible_client_rejection": "No written instruction, no prior approval, or the change is within original scope.",
        "counterargument": "Show constructive change, changed scope, or instruction trail even if formal variation wording is absent.",
    },
    {
        "section_name": "Payment Entitlements",
        "keywords": ["payment", "certificate", "interim payment", "statement", "amount due", "invoice"],
        "claim_type": "Payment / Certification",
        "required_evidence": "Payment certificate, invoice, valuation, payment status records, and resulting disruption evidence.",
        "possible_client_rejection": "No due amount, no approval, or payment delay does not justify time/cost relief.",
        "counterargument": "Show the certified / due amounts, delayed release, and the resulting impact on progress or cost.",
    },
    {
        "section_name": "Suspension Rights",
        "keywords": ["suspend", "suspension", "reduce the rate of work", "termination", "right to suspend"],
        "claim_type": "Suspension",
        "required_evidence": "Notice trail, unpaid amounts or blocking event proof, mitigation and consequence records.",
        "possible_client_rejection": "Suspension conditions were not met or notice was defective.",
        "counterargument": "Prove the contractual preconditions and the notice chain were satisfied.",
    },
    {
        "section_name": "Delay Damages Defense",
        "keywords": ["delay damages", "liquidated damages", "deduction for delay", "damages for delay"],
        "claim_type": "Delay Damages Defense",
        "required_evidence": "Criticality analysis, entitlement basis, excusable delay record, concurrency review.",
        "possible_client_rejection": "Damages remain payable because excusable delay is not proven.",
        "counterargument": "Use excusable delay and concurrency analysis to defeat or reduce liquidated damages exposure.",
    },
    {
        "section_name": "Material Escalation / Price Adjustment",
        "keywords": ["price adjustment", "escalation", "fluctuation", "change in cost", "rise or fall"],
        "claim_type": "Escalation / Price Adjustment",
        "required_evidence": "Market data, quotations, supplier notices, procurement chronology, contractual formula or basis.",
        "possible_client_rejection": "Fixed-price contract with no escalation mechanism.",
        "counterargument": "Identify contractual exceptions, statutory changes, instructed changes, or specific escalation wording.",
    },
    {
        "section_name": "Notice Requirements",
        "keywords": ["notice", "within", "days after", "shall give notice", "notify the engineer", "submit notice"],
        "claim_type": "Notice Compliance",
        "required_evidence": "Notice letters, emails, meeting minutes, proof of submission and receipt.",
        "possible_client_rejection": "Notice was late, missing, or defective.",
        "counterargument": "Demonstrate timely notice, continuing notice trail, or waiver / actual knowledge by the receiving party.",
    },
    {
        "section_name": "Time-Bar Risks",
        "keywords": ["time bar", "barred", "shall not be entitled unless", "failure to give notice", "forfeit"],
        "claim_type": "Time-Bar Risk",
        "required_evidence": "Notice dates, acknowledgement, waiver conduct, actual knowledge, continuing event trail.",
        "possible_client_rejection": "Claim is time-barred.",
        "counterargument": "Argue compliance, waiver, estoppel, continuing breach, or actual knowledge depending on the facts.",
    },
    {
        "section_name": "Client Defense Clauses",
        "keywords": ["sole remedy", "no additional payment", "fixed price", "contractor deemed to have inspected", "contractor shall be deemed"],
        "claim_type": "Client Defense",
        "required_evidence": "Full clause context, factual distinction, exception wording, event-specific records.",
        "possible_client_rejection": "Client relies on express limiting language to reject entitlement.",
        "counterargument": "Narrow the clause to its true scope and identify express exceptions or inconsistent employer conduct.",
    },
    {
        "section_name": "Contractor Counterarguments",
        "keywords": ["provided that", "except", "unless", "subject to", "without prejudice"],
        "claim_type": "Counterargument / Exception",
        "required_evidence": "Exception facts, supporting letters, notices, approvals, and event records.",
        "possible_client_rejection": "The exception does not apply on the facts.",
        "counterargument": "Tie the facts directly to the contractual exception or saving wording.",
    },
    {
        "section_name": "Evidence Requirements",
        "keywords": ["records", "particulars", "contemporary", "substantiation", "evidence", "maintain records"],
        "claim_type": "Evidence / Substantiation",
        "required_evidence": "All records explicitly required by the clause plus supporting schedule and commercial evidence.",
        "possible_client_rejection": "Claim is not substantiated.",
        "counterargument": "Show that the required records exist and directly support the claimed event and impact.",
    },
    {
        "section_name": "Dispute Resolution",
        "keywords": ["dispute", "arbitration", "amicable settlement", "adjudication", "dissatisfaction"],
        "claim_type": "Dispute Resolution",
        "required_evidence": "Notices of dispute, decision dates, referrals, and the underlying claim record.",
        "possible_client_rejection": "Procedure not followed.",
        "counterargument": "Follow the contractual sequence strictly and preserve all procedural milestones.",
    },
    {
        "section_name": "Risk Clauses",
        "keywords": ["risk", "liability", "indemnity", "responsibility", "loss", "damage"],
        "claim_type": "Risk Allocation",
        "required_evidence": "Event facts, allocation wording, notices, mitigation evidence, and actual impact proof.",
        "possible_client_rejection": "Risk lies with the contractor under the contract.",
        "counterargument": "Show the risk was expressly allocated elsewhere or shifted by employer / engineer conduct.",
    },
]


DEFENSE_RULES = [
    {
        "code": "NO_NOTICE",
        "title": "No notice",
        "keywords": ["no notice", "notice not submitted", "no formal notice", "no valid notice"],
        "danger": "A missing notice can be used to defeat entitlement entirely where the contract contains notice conditions or time-bar wording.",
    },
    {
        "code": "LATE_NOTICE",
        "title": "Late notice",
        "keywords": ["late notice", "notice was late", "notice submitted late", "out of time notice", "time bar"],
        "danger": "Late notice creates immediate time-bar risk and gives the employer a procedural rejection before the merits are even tested.",
    },
    {
        "code": "NO_WRITTEN_INSTRUCTION",
        "title": "No written instruction",
        "keywords": ["no written instruction", "verbal instruction", "no formal instruction", "no instruction in writing"],
        "danger": "The client can argue the alleged change was never instructed and is therefore not a compensable variation or compensable delay event.",
    },
    {
        "code": "NOT_CRITICAL_PATH",
        "title": "Not on critical path",
        "keywords": ["not on critical path", "not critical", "float available", "not on longest path", "non critical"],
        "danger": "If the impacted activity was not driving completion or a key milestone, time entitlement can be rejected even if disruption occurred.",
    },
    {
        "code": "CONCURRENT_DELAY",
        "title": "Concurrent delay",
        "keywords": ["concurrent delay", "overlapping delay", "parallel delay", "same period delay"],
        "danger": "Concurrency can reduce or defeat time and cost recovery unless the contractor isolates the client-caused impact and the overlap period.",
    },
    {
        "code": "CONTRACTOR_DELAY",
        "title": "Contractor delay",
        "keywords": ["contractor delay", "contractor caused", "late by contractor", "contractor responsible delay"],
        "danger": "A contractor-responsible delay allegation can break the causal chain and shift liability away from the employer event.",
    },
    {
        "code": "FIXED_PRICE",
        "title": "Fixed price",
        "keywords": ["fixed price", "lump sum", "no escalation", "firm price"],
        "danger": "Fixed-price wording is commonly used to reject escalation, disruption, and some variation valuation arguments unless a contractual exception exists.",
    },
    {
        "code": "SCOPE_ALREADY_INCLUDED",
        "title": "Scope already included",
        "keywords": ["already included", "within original scope", "included in contract", "contract scope already covers"],
        "danger": "If the event is treated as original scope, the contractor may lose both valuation and time entitlement for the claimed change.",
    },
    {
        "code": "INSUFFICIENT_SUBSTANTIATION",
        "title": "Insufficient substantiation",
        "keywords": ["insufficient substantiation", "not substantiated", "insufficient evidence", "unsupported claim"],
        "danger": "Even a strong contractual entitlement will fail if the factual record does not prove event, impact, and quantum.",
    },
    {
        "code": "NO_CONTEMPORANEOUS_RECORDS",
        "title": "No contemporaneous records",
        "keywords": ["no contemporaneous records", "no records at the time", "no site records", "no contemporaneous evidence"],
        "danger": "Without contemporaneous records, the client can challenge credibility, causation, and the reliability of later reconstructions.",
    },
    {
        "code": "ENGINEER_APPROVAL_NOT_ENTITLEMENT",
        "title": "Engineer approval not entitlement",
        "keywords": ["engineer approval not entitlement", "approval does not mean entitlement", "approval is not claim", "approval only"],
        "danger": "Approvals may support the facts but usually do not by themselves establish contractual entitlement or valuation.",
    },
    {
        "code": "FAILURE_TO_MITIGATE",
        "title": "Failure to mitigate",
        "keywords": ["failed to mitigate", "failure to mitigate", "no mitigation", "insufficient mitigation"],
        "danger": "Mitigation failures can reduce recoverable time or cost and weaken the contractor's commercial credibility.",
    },
    {
        "code": "NO_CAUSATION",
        "title": "No causation",
        "keywords": ["no causation", "no causal link", "did not cause delay", "no cause and effect"],
        "danger": "If cause-and-effect is not shown clearly, entitlement can fail even where the event itself is undisputed.",
    },
    {
        "code": "NO_COST_BREAKDOWN",
        "title": "No cost breakdown",
        "keywords": ["no cost breakdown", "cost not broken down", "no quantum support", "no supporting cost detail"],
        "danger": "Cost heads can be rejected where valuation is not transparently tied to records, quantities, and contractual valuation rules.",
    },
    {
        "code": "NO_PROGRAMME_IMPACT",
        "title": "No programme impact",
        "keywords": ["no programme impact", "no schedule impact", "no impact on programme", "no impact on completion"],
        "danger": "Without a measured programme effect, time entitlement arguments remain incomplete even if a disruptive event occurred.",
    },
]


CLAIM_CATEGORIES = [
    "EOT / Delay Claim",
    "Variation Claim",
    "Payment Claim",
    "Suspension / Slowdown",
    "Material Escalation",
    "Late Drawings / Late Approvals",
    "Employer Free-Issue Material",
    "Site Access / Handover",
    "Engineer Instruction",
    "Defects / DLP",
    "Delay Damages Defense",
    "Dispute / Notice / Time-Bar",
    "General Contract Risk",
]


QUESTION_CATEGORY_RULES = [
    ("EOT / Delay Claim", ["eot", "extension of time", "delay claim", "time impact", "critical path", "delay event", "longest path"]),
    ("Variation Claim", ["variation", "change order", "additional work", "scope change", "constructive variation", "quantity change"]),
    ("Payment Claim", ["payment", "invoice", "certificate", "certified amount", "ipc", "certification", "unpaid"]),
    ("Suspension / Slowdown", ["suspension", "slowdown", "reduced rate", "demobil", "stop work"]),
    ("Material Escalation", ["escalation", "price adjustment", "steel increased", "material increase", "fluctuation"]),
    ("Late Drawings / Late Approvals", ["ifc", "late drawing", "late approval", "drawing log", "rfi", "submittal", "approval"]),
    ("Employer Free-Issue Material", ["free-issue material", "free issue", "employer supplied material", "delayed steel", "material shortage", "steel supply"]),
    ("Site Access / Handover", ["site access", "handover", "possession", "workfront", "access restriction"]),
    ("Engineer Instruction", ["engineer instruction", "instruction", "site instruction", "verbal instruction", "directed by engineer"]),
    ("Defects / DLP", ["defect", "dlp", "defects liability", "snag", "rectification"]),
    ("Delay Damages Defense", ["delay damages", "liquidated damages", "ld", "damages for delay"]),
    ("Dispute / Notice / Time-Bar", ["notice", "time bar", "dispute", "adjudication", "arbitration", "waiver"]),
]


EVIDENCE_SCORE_RULES = [
    ("contract_clause_exists", "Contract clause exists", 20),
    ("written_notice_exists", "Written notice exists", 15),
    ("correspondence_exists", "Engineer/client correspondence exists", 15),
    ("approved_instruction_exists", "Approved drawing or instruction exists", 15),
    ("programme_impact_exists", "Programme impact exists", 15),
    ("cost_records_exist", "Cost records exist", 10),
    ("daily_records_exist", "Daily/site records exist", 10),
]


def now_iso() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def normalize_text(value: Any) -> str:
    text = str(value or "")
    text = re.sub(r"\s+", " ", text).strip()
    return text


def slugify(text: str) -> str:
    return re.sub(r"[^a-z0-9]+", "_", text.lower()).strip("_")


def compute_file_hash_from_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def ensure_contract_claims_dirs(app_dir: Path) -> dict[str, Path]:
    base_dir = app_dir / "05-contracts"
    contracts_dir = base_dir / "source"
    evidence_dir = app_dir / "06-evidence"
    exports_dir = app_dir / "11-outputs"
    for path in [base_dir, contracts_dir, evidence_dir, exports_dir]:
        path.mkdir(parents=True, exist_ok=True)
    return {
        "base_dir": base_dir,
        "contracts_dir": contracts_dir,
        "evidence_dir": evidence_dir,
        "exports_dir": exports_dir,
    }


def connect_db(db_path: Path) -> sqlite3.Connection:
    conn = sqlite3.connect(str(db_path))
    conn.row_factory = sqlite3.Row
    return conn


def init_contract_claims_db(db_path: Path) -> None:
    conn = connect_db(db_path)
    cur = conn.cursor()
    ddl = [
        """
        CREATE TABLE IF NOT EXISTS contract_documents (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            file_name TEXT NOT NULL,
            file_path TEXT NOT NULL,
            file_hash TEXT NOT NULL,
            file_type TEXT,
            upload_date TEXT,
            analysis_date TEXT,
            contract_version TEXT,
            processing_status TEXT,
            knowledge_base_status TEXT,
            content_text TEXT,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        )
        """,
        """
        CREATE TABLE IF NOT EXISTS contract_clauses (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            document_id INTEGER,
            clause_number TEXT,
            clause_title TEXT,
            exact_clause_text TEXT,
            plain_english_meaning TEXT,
            section_name TEXT,
            contractor_right TEXT,
            employer_obligation TEXT,
            engineer_obligation TEXT,
            claim_type TEXT,
            required_evidence TEXT,
            possible_client_rejection TEXT,
            contractor_counterargument TEXT,
            risk_level TEXT,
            claim_strength TEXT,
            time_impact TEXT,
            cost_impact TEXT,
            notice_required TEXT,
            recommended_action TEXT,
            related_project_records_needed TEXT,
            keywords_json TEXT,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        )
        """,
        """
        CREATE TABLE IF NOT EXISTS claim_categories (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            category_name TEXT UNIQUE,
            description TEXT,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        )
        """,
        """
        CREATE TABLE IF NOT EXISTS claim_triggers (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            category_name TEXT,
            trigger_name TEXT,
            trigger_keywords TEXT,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        )
        """,
        """
        CREATE TABLE IF NOT EXISTS evidence_documents (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            file_name TEXT NOT NULL,
            file_path TEXT NOT NULL,
            file_hash TEXT NOT NULL,
            file_type TEXT,
            source_stream TEXT,
            upload_date TEXT,
            extracted_text TEXT,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        )
        """,
        """
        CREATE TABLE IF NOT EXISTS evidence_mappings (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            evidence_document_id INTEGER,
            clause_id INTEGER,
            claim_category TEXT,
            mapping_score REAL,
            mapping_basis TEXT,
            missing_evidence_items TEXT,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        )
        """,
        """
        CREATE TABLE IF NOT EXISTS client_defenses (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            defense_code TEXT UNIQUE,
            defense_title TEXT,
            defense_keywords TEXT,
            standard_risk TEXT,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        )
        """,
        """
        CREATE TABLE IF NOT EXISTS contractor_rebuttals (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            defense_id INTEGER,
            clause_id INTEGER,
            rebuttal_text TEXT,
            evidence_needed TEXT,
            probability_of_success TEXT,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        )
        """,
        """
        CREATE TABLE IF NOT EXISTS claim_drafts (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            draft_name TEXT,
            claim_type TEXT,
            delay_event TEXT,
            relevant_clause_ids TEXT,
            evidence_ids TEXT,
            client_rejection_text TEXT,
            narrative_text TEXT,
            contractual_basis TEXT,
            factual_background TEXT,
            cause_effect TEXT,
            evidence_list TEXT,
            entitlement_statement TEXT,
            time_impact_statement TEXT,
            cost_impact_statement TEXT,
            rebuttal_section TEXT,
            attachment_checklist TEXT,
            status TEXT,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        )
        """,
        """
        CREATE TABLE IF NOT EXISTS contract_risk_register (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            clause_id INTEGER,
            risk_title TEXT,
            risk_level TEXT,
            risk_type TEXT,
            exposure TEXT,
            mitigation_action TEXT,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        )
        """,
        """
        CREATE TABLE IF NOT EXISTS contract_metadata (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            meta_key TEXT UNIQUE,
            meta_value TEXT,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        )
        """,
        """
        CREATE TABLE IF NOT EXISTS contract_versions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            document_id INTEGER,
            file_hash TEXT,
            contract_version TEXT,
            analysis_date TEXT,
            is_active INTEGER,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        )
        """,
        """
        CREATE TABLE IF NOT EXISTS contract_analysis_status (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            document_id INTEGER,
            file_hash TEXT,
            last_analysis_date TEXT,
            processing_status TEXT,
            knowledge_base_status TEXT,
            total_clauses_extracted INTEGER,
            error_message TEXT,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        )
        """,
        """
        CREATE TABLE IF NOT EXISTS contract_knowledge_base (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            document_id INTEGER,
            knowledge_json TEXT,
            summary_text TEXT,
            knowledge_base_status TEXT,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        )
        """,
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
        """,
    ]
    for statement in ddl:
        cur.execute(statement)
    _seed_reference_data(cur)
    conn.commit()
    conn.close()


def _seed_reference_data(cur: sqlite3.Cursor) -> None:
    stamp = now_iso()
    for category in CLAIM_CATEGORIES:
        cur.execute(
            """
            INSERT INTO claim_categories (category_name, description, created_at, updated_at)
            VALUES (?, ?, ?, ?)
            ON CONFLICT(category_name) DO UPDATE SET updated_at=excluded.updated_at
            """,
            (category, f"Auto-seeded claim category for {category}.", stamp, stamp),
        )
    for rule in DEFENSE_RULES:
        cur.execute(
            """
            INSERT INTO client_defenses (defense_code, defense_title, defense_keywords, standard_risk, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?, ?)
            ON CONFLICT(defense_code) DO UPDATE SET defense_title=excluded.defense_title, defense_keywords=excluded.defense_keywords, updated_at=excluded.updated_at
            """,
            (rule["code"], rule["title"], json.dumps(rule["keywords"]), "High", stamp, stamp),
        )
    for rule in CLAIM_SECTION_RULES:
        trigger_name = rule["section_name"]
        cur.execute(
            "SELECT id FROM claim_triggers WHERE category_name = ? AND trigger_name = ?",
            (rule["claim_type"], trigger_name),
        )
        if cur.fetchone() is None:
            cur.execute(
                """
                INSERT INTO claim_triggers (category_name, trigger_name, trigger_keywords, created_at, updated_at)
                VALUES (?, ?, ?, ?, ?)
                """,
                (rule["claim_type"], trigger_name, json.dumps(rule["keywords"]), stamp, stamp),
            )


def _is_path_inside(child: Path, parent: Path) -> bool:
    try:
        child.resolve().relative_to(parent.resolve())
        return True
    except (OSError, ValueError):
        return False


def prune_contract_documents_outside_scope(db_path: Path, contracts_dir: Path) -> int:
    """Remove stale analysis rows that point outside the active project's contract source folder."""
    init_contract_claims_db(db_path)
    conn = connect_db(db_path)
    cur = conn.cursor()
    cur.execute("SELECT id, file_path FROM contract_documents")
    stale_document_ids: list[int] = []
    for row in cur.fetchall():
        stored_path = Path(str(row["file_path"] or ""))
        if not stored_path.exists() or not _is_path_inside(stored_path, contracts_dir):
            stale_document_ids.append(int(row["id"]))
    if stale_document_ids:
        placeholders = ",".join("?" for _ in stale_document_ids)
        cur.execute(f"DELETE FROM contract_clauses WHERE document_id IN ({placeholders})", stale_document_ids)
        cur.execute(f"DELETE FROM contract_knowledge_base WHERE document_id IN ({placeholders})", stale_document_ids)
        cur.execute(f"DELETE FROM contract_versions WHERE document_id IN ({placeholders})", stale_document_ids)
        cur.execute(f"DELETE FROM contract_analysis_status WHERE document_id IN ({placeholders})", stale_document_ids)
        cur.execute("DELETE FROM contractor_rebuttals WHERE clause_id NOT IN (SELECT id FROM contract_clauses)")
        cur.execute("DELETE FROM contract_risk_register WHERE clause_id NOT IN (SELECT id FROM contract_clauses)")
        cur.execute("DELETE FROM evidence_mappings WHERE clause_id NOT IN (SELECT id FROM contract_clauses)")
        cur.execute(f"DELETE FROM contract_documents WHERE id IN ({placeholders})", stale_document_ids)
        stamp = now_iso()
        cur.execute(
            """
            INSERT INTO contract_metadata (meta_key, meta_value, created_at, updated_at)
            VALUES (?, ?, ?, ?)
            ON CONFLICT(meta_key) DO UPDATE SET meta_value=excluded.meta_value, updated_at=excluded.updated_at
            """,
            ("scope_cleanup", f"Removed {len(stale_document_ids)} stale document(s) outside {contracts_dir}", stamp, stamp),
        )
    cur.execute(
        """
        SELECT file_hash, GROUP_CONCAT(id) AS ids, COUNT(*) AS duplicate_count
        FROM contract_documents
        GROUP BY file_hash
        HAVING duplicate_count > 1
        """
    )
    duplicate_document_ids: list[int] = []
    for row in cur.fetchall():
        ids = [int(value) for value in str(row["ids"] or "").split(",") if str(value).strip().isdigit()]
        duplicate_document_ids.extend(sorted(ids)[:-1])
    if duplicate_document_ids:
        placeholders = ",".join("?" for _ in duplicate_document_ids)
        cur.execute(f"DELETE FROM contract_clauses WHERE document_id IN ({placeholders})", duplicate_document_ids)
        cur.execute(f"DELETE FROM contract_knowledge_base WHERE document_id IN ({placeholders})", duplicate_document_ids)
        cur.execute(f"DELETE FROM contract_versions WHERE document_id IN ({placeholders})", duplicate_document_ids)
        cur.execute(f"DELETE FROM contract_analysis_status WHERE document_id IN ({placeholders})", duplicate_document_ids)
        cur.execute(f"DELETE FROM contract_documents WHERE id IN ({placeholders})", duplicate_document_ids)
    conn.commit()
    conn.close()
    return len(stale_document_ids) + len(duplicate_document_ids)


def contract_file_list(contracts_dir: Path) -> list[Path]:
    if not contracts_dir.exists():
        return []
    return sorted([path for path in contracts_dir.iterdir() if path.is_file() and path.suffix.lower() in SUPPORTED_CONTRACT_EXTENSIONS])


def is_curated_contract_library_file(path: Path) -> bool:
    name = path.name.lower()
    return path.suffix.lower() in {".csv", ".xls", ".xlsx"} and ("clause_library" in name or "contract_library" in name)


def contract_source_files_for_library(contracts_dir: Path) -> list[Path]:
    return [
        path
        for path in contract_file_list(contracts_dir)
        if not is_curated_contract_library_file(path)
    ]


def build_curated_clause_library_rows(source_files: list[Path]) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for source_file in source_files:
        content_text = extract_text_for_clause_library_generation(source_file)
        clause_records = extract_clauses_from_text(content_text, document_id=0)
        for sequence, clause in enumerate(clause_records, start=1):
            rows.append(
                {
                    "Clause / Topic": clause.get("clause_title") or clause.get("clause_number") or f"{source_file.stem} Clause {sequence}",
                    "Location": clause.get("clause_number") or f"{source_file.name} / Extract {sequence}",
                    "Plain English Meaning": clause.get("plain_english_meaning") or normalize_text(clause.get("exact_clause_text", ""))[:500],
                    "Beneath the Lines": clause.get("contractor_counterargument") or "",
                    "Who Holds Leverage": "Contractor"
                    if str(clause.get("contractor_right", "")).strip()
                    else "Employer"
                    if str(clause.get("employer_obligation", "")).strip()
                    else "Engineer"
                    if str(clause.get("engineer_obligation", "")).strip()
                    else "Shared",
                    "Notice / Time Bar": clause.get("notice_required") or "",
                    "Money Impact": clause.get("cost_impact") or "",
                    "Schedule Impact": clause.get("time_impact") or "",
                    "Practical Action / Evidence": clause.get("required_evidence") or clause.get("recommended_action") or "",
                    "Source File": source_file.name,
                    "Claim Type": clause.get("claim_type") or "",
                    "Risk Level": clause.get("risk_level") or "",
                    "Claim Strength": clause.get("claim_strength") or "",
                    "Exact Clause Text": clause.get("exact_clause_text") or "",
                }
            )
    return rows


def extract_text_for_clause_library_generation(path: Path) -> str:
    """Use bounded extraction for automatic library generation so the UI never hangs."""
    suffix = path.suffix.lower()
    if suffix == ".pdf":
        for module_name in ["pypdf", "PyPDF2"]:
            try:
                module = importlib.import_module(module_name)
                reader_cls = getattr(module, "PdfReader", None)
                if reader_cls is None:
                    continue
                reader = reader_cls(str(path))
                pages = []
                for page in list(reader.pages)[:120]:
                    try:
                        pages.append(page.extract_text() or "")
                    except Exception:
                        pages.append("")
                return "\n".join(pages)
            except Exception:
                continue
        return ""
    return extract_text_from_path(path)


def ensure_auto_contract_clause_library(contracts_dir: Path, rebuild: bool = False) -> dict[str, Any]:
    contracts_dir.mkdir(parents=True, exist_ok=True)
    library_path = contracts_dir / "Overall_Contract_clause_library.xlsx"
    source_files = contract_source_files_for_library(contracts_dir)
    status = {
        "library_path": library_path,
        "generated": False,
        "source_files": [path.name for path in source_files],
        "row_count": 0,
        "message": "",
    }
    if not source_files:
        status["message"] = "No source contract file found for automatic clause-library generation."
        return status
    source_newer_than_library = (
        not library_path.exists()
        or any(path.stat().st_mtime > library_path.stat().st_mtime for path in source_files)
    )
    if library_path.exists() and not rebuild and not source_newer_than_library:
        status["message"] = "Existing clause library is current."
        return status
    rows = build_curated_clause_library_rows(source_files)
    if not rows:
        status["message"] = "No extractable contract clauses were found. Existing library was preserved." if library_path.exists() else "No extractable contract clauses were found."
        return status
    output_df = pd.DataFrame(rows)
    with pd.ExcelWriter(library_path, engine="openpyxl") as writer:
        output_df.to_excel(writer, sheet_name="Contract Clause Library", index=False)
        workbook = writer.book
        worksheet = writer.sheets["Contract Clause Library"]
        for column_cells in worksheet.columns:
            header = str(column_cells[0].value or "")
            max_length = max([len(str(cell.value or "")) for cell in column_cells[:200]] + [len(header)])
            worksheet.column_dimensions[column_cells[0].column_letter].width = min(max(max_length + 2, 14), 48)
        worksheet.freeze_panes = "A2"
    status.update(
        {
            "generated": True,
            "row_count": int(len(output_df)),
            "message": f"Generated {library_path.name} from {len(source_files)} source contract file(s).",
        }
    )
    return status


def evidence_file_list(evidence_dir: Path) -> list[Path]:
    if not evidence_dir.exists():
        return []
    return sorted([path for path in evidence_dir.iterdir() if path.is_file() and path.suffix.lower() in SUPPORTED_EVIDENCE_EXTENSIONS])


def _read_csv_or_excel_text(path: Path) -> str:
    if path.suffix.lower() == ".csv":
        df = pd.read_csv(path).fillna("")
        return df.to_csv(index=False)
    frames: list[str] = []
    xls = pd.ExcelFile(path)
    for sheet in xls.sheet_names:
        sheet_df = pd.read_excel(path, sheet_name=sheet).fillna("")
        frames.append(f"[Sheet: {sheet}]\n{sheet_df.to_csv(index=False)}")
    return "\n\n".join(frames)


def _read_docx_text(path: Path) -> str:
    try:
        Document = importlib.import_module("docx").Document
    except ModuleNotFoundError:
        return ""
    doc = Document(path)
    return "\n".join(p.text for p in doc.paragraphs if p.text.strip())


def _pdf_has_extractable_text(path: Path, max_pages: int = 5) -> bool:
    for module_name in ["pypdf", "PyPDF2"]:
        try:
            module = importlib.import_module(module_name)
            reader_cls = getattr(module, "PdfReader", None)
            if reader_cls is None:
                continue
            reader = reader_cls(str(path))
            pages = []
            for page in list(reader.pages)[:max_pages]:
                try:
                    pages.append(page.extract_text() or "")
                except Exception:
                    pages.append("")
            return bool(normalize_text("\n".join(pages)))
        except ModuleNotFoundError:
            continue
        except Exception:
            continue
    return False


def _ocr_engine_status() -> dict[str, Any]:
    tesseract_path = shutil.which("tesseract")
    try:
        importlib.import_module("pytesseract")
        pytesseract_available = True
    except ModuleNotFoundError:
        pytesseract_available = False
    return {
        "available": bool(tesseract_path and pytesseract_available),
        "tesseract_path": tesseract_path or "",
        "pytesseract_available": pytesseract_available,
    }


def _pdf_render_status() -> dict[str, Any]:
    try:
        importlib.import_module("pypdfium2")
        importlib.import_module("PIL")
        return {"available": True, "engine": "pypdfium2"}
    except ModuleNotFoundError:
        return {"available": False, "engine": ""}


def _read_pdf_text_via_ocr(path: Path, max_pages: int | None = None) -> str:
    ocr_status = _ocr_engine_status()
    render_status = _pdf_render_status()
    if not ocr_status["available"] or not render_status["available"]:
        return ""
    try:
        pdfium = importlib.import_module("pypdfium2")
        pytesseract = importlib.import_module("pytesseract")
        pdf = pdfium.PdfDocument(str(path))
        page_count = len(pdf)
        limit = page_count if max_pages is None else min(page_count, max_pages)
        pages = []
        for index in range(limit):
            page = pdf[index]
            image = page.render(scale=2).to_pil()
            pages.append(pytesseract.image_to_string(image) or "")
        return "\n".join(pages)
    except Exception:
        return ""


def build_pdf_source_diagnostics(path: Path) -> dict[str, Any]:
    pages = 0
    extractable_text = False
    try:
        module = importlib.import_module("pypdf")
        reader = module.PdfReader(str(path))
        pages = len(reader.pages)
        extractable_text = _pdf_has_extractable_text(path)
    except Exception:
        pass
    ocr_status = _ocr_engine_status()
    render_status = _pdf_render_status()
    return {
        "file_name": path.name,
        "pages": pages,
        "extractable_text": extractable_text,
        "is_scanned_or_image_based": not extractable_text,
        "render_preview_available": render_status["available"],
        "render_engine": render_status["engine"],
        "ocr_available": ocr_status["available"],
        "tesseract_path": ocr_status["tesseract_path"],
        "pytesseract_available": ocr_status["pytesseract_available"],
    }


def _read_pdf_text(path: Path) -> str:
    for module_name in ["pypdf", "PyPDF2"]:
        try:
            module = importlib.import_module(module_name)
            reader_cls = getattr(module, "PdfReader", None)
            if reader_cls is None:
                continue
            reader = reader_cls(str(path))
            pages = []
            for page in reader.pages:
                try:
                    pages.append(page.extract_text() or "")
                except Exception:
                    pages.append("")
            text = "\n".join(pages)
            if normalize_text(text):
                return text
        except ModuleNotFoundError:
            continue
        except Exception:
            continue
    word_text = _read_pdf_text_via_word(path)
    if normalize_text(word_text):
        return word_text
    return _read_pdf_text_via_ocr(path)


def _pdf_source_warning(path: Path) -> str:
    diagnostics = build_pdf_source_diagnostics(path)
    if not diagnostics["is_scanned_or_image_based"]:
        return ""
    if diagnostics["ocr_available"]:
        return f"{path.name} appears scanned or image-based; OCR fallback is available if this PDF is selected as the active analysis source."
    if diagnostics["render_preview_available"]:
        return f"{path.name} appears scanned or image-based with no extractable text; page rendering is available, but OCR is not installed. Install Tesseract and pytesseract for direct scanned-PDF clause extraction."
    return f"{path.name} appears scanned or image-based with no extractable text; install OCR support or provide a searchable PDF for direct PDF clause extraction."


def _read_pdf_text_via_word(path: Path) -> str:
    escaped_path = str(path).replace("'", "''")
    command = [
        "powershell",
        "-NoProfile",
        "-Command",
        (
            f"$path = '{escaped_path}'; "
            "$word = $null; "
            "try { "
            "$word = New-Object -ComObject Word.Application; "
            "$word.Visible = $false; "
            "$doc = $word.Documents.Open($path, $false, $true); "
            "$text = $doc.Content.Text; "
            "$doc.Close(); "
            "[Console]::OutputEncoding = [System.Text.Encoding]::UTF8; "
            "Write-Output $text "
            "} finally { "
            "if ($word) { $word.Quit() } "
            "}"
        ),
    ]
    try:
        result = subprocess.run(
            command,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="ignore",
            timeout=180,
            check=False,
        )
        text = result.stdout or ""
        return text if normalize_text(text) else ""
    except Exception:
        return ""


def extract_text_from_path(path: Path) -> str:
    suffix = path.suffix.lower()
    try:
        if suffix == ".txt":
            return path.read_text(encoding="utf-8", errors="ignore")
        if suffix in {".csv", ".xls", ".xlsx"}:
            return _read_csv_or_excel_text(path)
        if suffix == ".docx":
            return _read_docx_text(path)
        if suffix == ".pdf":
            return _read_pdf_text(path)
    except Exception:
        return ""
    return ""


def extract_text_from_bytes(file_name: str, data: bytes) -> str:
    suffix = Path(file_name).suffix.lower()
    if suffix == ".txt":
        return data.decode("utf-8", errors="ignore")
    if suffix == ".csv":
        try:
            df = pd.read_csv(io.BytesIO(data)).fillna("")
            return df.to_csv(index=False)
        except Exception:
            return data.decode("utf-8", errors="ignore")
    if suffix in {".xls", ".xlsx"}:
        try:
            xls = pd.ExcelFile(io.BytesIO(data))
            frames = []
            for sheet in xls.sheet_names:
                sheet_df = pd.read_excel(io.BytesIO(data), sheet_name=sheet).fillna("")
                frames.append(f"[Sheet: {sheet}]\n{sheet_df.to_csv(index=False)}")
            return "\n\n".join(frames)
        except Exception:
            return ""
    if suffix == ".docx":
        try:
            Document = importlib.import_module("docx").Document
            doc = Document(io.BytesIO(data))
            return "\n".join(p.text for p in doc.paragraphs if p.text.strip())
        except Exception:
            return ""
    if suffix == ".pdf":
        for module_name in ["pypdf", "PyPDF2"]:
            try:
                module = importlib.import_module(module_name)
                reader_cls = getattr(module, "PdfReader", None)
                if reader_cls is None:
                    continue
                reader = reader_cls(io.BytesIO(data))
                pages = []
                for page in reader.pages:
                    pages.append(page.extract_text() or "")
                text = "\n".join(pages)
                if normalize_text(text):
                    return text
            except Exception:
                continue
    return ""


def parse_spreadsheet_contract_library(path: Path, document_id: int) -> list[dict[str, Any]]:
    try:
        if path.suffix.lower() == ".csv":
            df = pd.read_csv(path).fillna("")
        else:
            df = pd.read_excel(path).fillna("")
    except Exception:
        return []
    if df.empty:
        return []
    normalized = {re.sub(r"[^a-z0-9]+", "", col.lower()): col for col in df.columns}
    curated_cols = {
        "topic": normalized.get("clausetopic"),
        "location": normalized.get("location"),
        "plain": normalized.get("plainenglishmeaning"),
        "beneath": normalized.get("beneaththelines"),
        "leverage": normalized.get("whoholdsleverage"),
        "notice": normalized.get("noticetimebar"),
        "money": normalized.get("moneyimpact"),
        "schedule": normalized.get("scheduleimpact"),
        "action": normalized.get("practicalactionevidence"),
    }
    if curated_cols["topic"] and curated_cols["plain"]:
        return parse_curated_contract_library(df, document_id, curated_cols)
    text_col = normalized.get("exactclausetext") or normalized.get("clausetext") or normalized.get("text")
    if not text_col:
        return []
    rows: list[dict[str, Any]] = []
    for _, row in df.iterrows():
        exact_text = normalize_text(row.get(text_col, ""))
        if not exact_text:
            continue
        clause_number = row.get(normalized.get("clausenumber", ""), "") if normalized.get("clausenumber") else ""
        clause_title = row.get(normalized.get("clausetitle", ""), "") if normalized.get("clausetitle") else ""
        meta = classify_clause_text(exact_text)
        rows.append(build_clause_record(document_id, clause_number, clause_title, exact_text, meta))
    return rows


def parse_curated_contract_library(
    df: pd.DataFrame,
    document_id: int,
    cols: dict[str, str | None],
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for _, row in df.iterrows():
        clause_title = normalize_text(row.get(cols["topic"] or "", ""))
        plain = normalize_text(row.get(cols["plain"] or "", ""))
        if not clause_title and not plain:
            continue
        location = normalize_text(row.get(cols["location"] or "", ""))
        beneath = normalize_text(row.get(cols["beneath"] or "", ""))
        leverage = normalize_text(row.get(cols["leverage"] or "", ""))
        notice = normalize_text(row.get(cols["notice"] or "", ""))
        money = normalize_text(row.get(cols["money"] or "", ""))
        schedule = normalize_text(row.get(cols["schedule"] or "", ""))
        action = normalize_text(row.get(cols["action"] or "", ""))
        combined_text = " ".join(
            part
            for part in [
                f"Topic: {clause_title}" if clause_title else "",
                f"Location: {location}" if location else "",
                plain,
                beneath,
                leverage,
                notice,
                money,
                schedule,
                action,
            ]
            if part
        )
        meta = classify_clause_text(combined_text)
        meta["plain_english_meaning"] = plain or meta["plain_english_meaning"]
        meta["contractor_counterargument"] = beneath or meta["contractor_counterargument"]
        meta["recommended_action"] = action or meta["recommended_action"]
        meta["required_evidence"] = action or meta["required_evidence"]
        meta["notice_required"] = "Yes" if notice else meta["notice_required"]
        meta["cost_impact"] = "Yes" if money else meta["cost_impact"]
        meta["time_impact"] = "Yes" if schedule else meta["time_impact"]
        meta["related_project_records_needed"] = action or meta["related_project_records_needed"]
        if leverage:
            leverage_lower = leverage.lower()
            if leverage_lower in {"employer", "engineer", "shared"}:
                meta["employer_obligation"] = "Yes" if leverage_lower in {"employer", "shared"} else meta["employer_obligation"]
                meta["engineer_obligation"] = "Yes" if leverage_lower == "engineer" else meta["engineer_obligation"]
            if leverage_lower == "contractor":
                meta["risk_level"] = "High"
        rows.append(
            build_clause_record(
                document_id=document_id,
                clause_number="",
                clause_title=clause_title,
                exact_text=combined_text,
                meta=meta,
            )
        )
    return rows


def extract_clauses_from_text(text: str, document_id: int) -> list[dict[str, Any]]:
    clean = normalize_text(text)
    if not clean:
        return []
    raw_lines = [line.strip() for line in text.splitlines() if line.strip()]
    clause_indices: list[tuple[int, str, str]] = []
    for idx, line in enumerate(raw_lines):
        match = re.match(r"^(?:clause\s*)?(\d+(?:\.\d+){0,3})\s*[-:.)]?\s*(.{0,120})$", line, flags=re.IGNORECASE)
        if match:
            clause_number = match.group(1)
            clause_title = match.group(2).strip()
            clause_indices.append((idx, clause_number, clause_title))
    records: list[dict[str, Any]] = []
    if len(clause_indices) >= 3:
        for index, (start_idx, clause_number, clause_title) in enumerate(clause_indices):
            end_idx = clause_indices[index + 1][0] if index + 1 < len(clause_indices) else len(raw_lines)
            chunk = " ".join(raw_lines[start_idx:end_idx])
            meta = classify_clause_text(chunk)
            records.append(build_clause_record(document_id, clause_number, clause_title, chunk, meta))
        return records
    paragraphs = [normalize_text(chunk) for chunk in re.split(r"\n\s*\n", text) if normalize_text(chunk)]
    for idx, paragraph in enumerate(paragraphs[:300], start=1):
        meta = classify_clause_text(paragraph)
        title = paragraph[:90]
        records.append(build_clause_record(document_id, f"P{idx:03d}", title, paragraph, meta))
    return records


def classify_clause_text(text: str) -> dict[str, Any]:
    lower = text.lower()
    best_rule = CLAIM_SECTION_RULES[0]
    best_score = -1
    for rule in CLAIM_SECTION_RULES:
        score = sum(1 for keyword in rule["keywords"] if keyword in lower)
        if score > best_score:
            best_rule = rule
            best_score = score
    section_name = best_rule["section_name"] if best_score > 0 else "Risk Clauses"
    claim_type = best_rule["claim_type"] if best_score > 0 else "General Contract Risk"
    contractor_right = "Possible" if any(term in lower for term in ["contractor shall be entitled", "contractor may", "right to", "extension of time", "payment", "variation"]) else ""
    employer_obligation = "Yes" if any(term in lower for term in ["employer shall", "provided by the employer", "the employer shall"]) else ""
    engineer_obligation = "Yes" if any(term in lower for term in ["engineer shall", "approval by the engineer", "the engineer shall"]) else ""
    risk_level = "Critical" if any(term in lower for term in ["time bar", "barred", "delay damages", "liquidated damages", "forfeit"]) else "High" if any(term in lower for term in ["notice", "suspend", "termination", "risk", "liability"]) else "Medium" if best_score > 0 else "Low"
    claim_strength = "Very Strong" if any(term in lower for term in ["shall be entitled", "the employer shall", "time for completion shall be extended"]) else "Strong" if any(term in lower for term in ["may", "extension of time", "variation", "payment"]) else "Medium" if best_score > 0 else "Weak"
    time_impact = "Yes" if any(term in lower for term in ["extension of time", "delay", "time for completion", "critical path", "suspend"]) else "Possible"
    cost_impact = "Yes" if any(term in lower for term in ["payment", "cost", "expense", "variation", "price", "damages"]) else "Possible"
    notice_required = "Yes" if "notice" in lower else "Unclear"
    plain_english = build_plain_english_meaning(text, section_name)
    recommended_action = build_recommended_action(section_name, claim_type, notice_required)
    related_records = build_related_records(section_name, claim_type)
    keywords = sorted({kw for rule in CLAIM_SECTION_RULES for kw in rule["keywords"] if kw in lower})
    return {
        "section_name": section_name,
        "claim_type": claim_type,
        "contractor_right": contractor_right,
        "employer_obligation": employer_obligation,
        "engineer_obligation": engineer_obligation,
        "required_evidence": best_rule["required_evidence"] if best_score > 0 else "Full clause context, event chronology, and project records.",
        "possible_client_rejection": best_rule["possible_client_rejection"] if best_score > 0 else "The clause does not create the entitlement being claimed.",
        "contractor_counterargument": best_rule["counterargument"] if best_score > 0 else "Tie the clause wording to the factual event with contemporaneous evidence.",
        "risk_level": risk_level,
        "claim_strength": claim_strength,
        "time_impact": time_impact,
        "cost_impact": cost_impact,
        "notice_required": notice_required,
        "plain_english_meaning": plain_english,
        "recommended_action": recommended_action,
        "related_project_records_needed": related_records,
        "keywords_json": json.dumps(keywords),
    }


def build_plain_english_meaning(text: str, section_name: str) -> str:
    short = normalize_text(text)[:320]
    return f"This clause falls under {section_name}. In practical contractor terms, it means: {short}"


def build_recommended_action(section_name: str, claim_type: str, notice_required: str) -> str:
    action = f"Review the clause under {section_name}, align the event facts to the claim type '{claim_type}', and prepare a contemporaneous evidence pack."
    if notice_required == "Yes":
        action += " Confirm that notice was issued on time and preserve proof of submission and receipt."
    return action


def build_related_records(section_name: str, claim_type: str) -> str:
    mapping = {
        "EOT Entitlements": "Accepted programme, fragnet / window analysis, notices, RFIs, instructions, daily reports.",
        "Variation Entitlements": "Instructions, revised drawings, quotations, quantity take-off, approvals, meeting minutes.",
        "Payment Entitlements": "Payment certificates, invoices, valuation sheets, payment status records, correspondence.",
        "Notice Requirements": "Notice letters, emails, transmittals, meeting minutes, acknowledgements.",
        "Engineer Obligations": "Submittal logs, RFI logs, drawing logs, approval records.",
        "Employer Obligations": "Site access records, free-issue material logs, handover records, employer letters.",
    }
    return mapping.get(section_name, f"Project records supporting {claim_type}, correspondence, and contemporaneous site / commercial records.")


def build_clause_record(document_id: int, clause_number: Any, clause_title: Any, exact_text: str, meta: dict[str, Any]) -> dict[str, Any]:
    stamp = now_iso()
    return {
        "document_id": document_id,
        "clause_number": normalize_text(clause_number),
        "clause_title": normalize_text(clause_title)[:220],
        "exact_clause_text": exact_text,
        "plain_english_meaning": meta["plain_english_meaning"],
        "section_name": meta["section_name"],
        "contractor_right": meta["contractor_right"],
        "employer_obligation": meta["employer_obligation"],
        "engineer_obligation": meta["engineer_obligation"],
        "claim_type": meta["claim_type"],
        "required_evidence": meta["required_evidence"],
        "possible_client_rejection": meta["possible_client_rejection"],
        "contractor_counterargument": meta["contractor_counterargument"],
        "risk_level": meta["risk_level"],
        "claim_strength": meta["claim_strength"],
        "time_impact": meta["time_impact"],
        "cost_impact": meta["cost_impact"],
        "notice_required": meta["notice_required"],
        "recommended_action": meta["recommended_action"],
        "related_project_records_needed": meta["related_project_records_needed"],
        "keywords_json": meta["keywords_json"],
        "created_at": stamp,
        "updated_at": stamp,
    }


def upsert_contract_document(cur: sqlite3.Cursor, path: Path, file_hash: str, content_text: str, contract_version: str) -> int:
    stamp = now_iso()
    clean_path = path.resolve()
    cur.execute("SELECT id FROM contract_documents WHERE file_path = ? OR file_hash = ? ORDER BY id LIMIT 1", (str(clean_path), file_hash))
    existing = cur.fetchone()
    if existing:
        cur.execute(
            """
            UPDATE contract_documents
            SET file_name=?, file_path=?, file_hash=?, file_type=?, upload_date=?, analysis_date=?, contract_version=?, processing_status=?, knowledge_base_status=?, content_text=?, updated_at=?
            WHERE id=?
            """,
            (clean_path.name, str(clean_path), file_hash, clean_path.suffix.lower(), stamp, stamp, contract_version, "Processed", "Ready", content_text, stamp, existing["id"]),
        )
        return int(existing["id"])
    cur.execute(
        """
        INSERT INTO contract_documents (file_name, file_path, file_hash, file_type, upload_date, analysis_date, contract_version, processing_status, knowledge_base_status, content_text, created_at, updated_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (clean_path.name, str(clean_path), file_hash, clean_path.suffix.lower(), stamp, stamp, contract_version, "Processed", "Ready", content_text, stamp, stamp),
    )
    return int(cur.lastrowid)


def persist_contract_analysis(db_path: Path, contracts_dir: Path, rebuild: bool = False) -> dict[str, Any]:
    init_contract_claims_db(db_path)
    contracts_dir = contracts_dir.resolve()
    auto_library_status = ensure_auto_contract_clause_library(contracts_dir, rebuild=rebuild)
    stale_documents_removed = prune_contract_documents_outside_scope(db_path, contracts_dir)
    files = contract_file_list(contracts_dir)
    analysis_files = [path for path in files if is_curated_contract_library_file(path)] or files
    supporting_files = [path for path in files if path not in analysis_files]
    source_warnings = []
    if analysis_files != files:
        source_warnings.append(
            "Curated contract clause library is the active analysis source; non-library contract files are retained as source references."
        )
    for path in supporting_files:
        if path.suffix.lower() == ".pdf":
            warning = _pdf_source_warning(path)
            if warning:
                source_warnings.append(warning)
    conn = connect_db(db_path)
    cur = conn.cursor()
    stamp = now_iso()
    contract_status = {
        "contract_loaded": False,
        "last_analysis_date": None,
        "total_clauses": 0,
        "contract_version": "N/A",
        "knowledge_base_status": "No contract detected",
        "reprocessed": False,
        "detected_files": [path.name for path in files],
        "analysis_files": [path.name for path in analysis_files],
        "supporting_files": [path.name for path in supporting_files],
        "source_warnings": source_warnings,
        "extraction_issue": "",
        "stale_documents_removed": stale_documents_removed,
        "auto_library_status": auto_library_status,
    }
    if not files:
        cur.execute(
            """
            INSERT INTO contract_metadata (meta_key, meta_value, created_at, updated_at)
            VALUES (?, ?, ?, ?)
            ON CONFLICT(meta_key) DO UPDATE SET meta_value=excluded.meta_value, updated_at=excluded.updated_at
            """,
            ("knowledge_base_status", "No contract detected", stamp, stamp),
        )
        conn.commit()
        conn.close()
        return contract_status

    file_fingerprints = []
    needs_rebuild = rebuild
    cur.execute("SELECT COUNT(*) FROM contract_clauses")
    existing_clause_count = int(cur.fetchone()[0] or 0)
    if existing_clause_count == 0:
        needs_rebuild = True
    for path in analysis_files:
        clean_path = path.resolve()
        file_hash = compute_file_hash_from_bytes(clean_path.read_bytes())
        file_fingerprints.append((clean_path, file_hash))
        cur.execute("SELECT file_hash FROM contract_documents WHERE file_path = ?", (str(clean_path),))
        existing = cur.fetchone()
        if existing is None or existing["file_hash"] != file_hash:
            needs_rebuild = True

    if needs_rebuild:
        cur.execute("DELETE FROM contract_clauses")
        cur.execute("DELETE FROM contract_knowledge_base")
        cur.execute("DELETE FROM contract_versions")
        cur.execute("DELETE FROM contract_analysis_status")
        total_clauses = 0
        version_counter = 1
        extraction_issue = ""
        for path, file_hash in file_fingerprints:
            content_text = extract_text_from_path(path)
            contract_version = f"v{version_counter}"
            document_id = upsert_contract_document(cur, path, file_hash, content_text, contract_version)
            if path.suffix.lower() in {".csv", ".xls", ".xlsx"}:
                clauses = parse_spreadsheet_contract_library(path, document_id)
                if not clauses:
                    clauses = extract_clauses_from_text(content_text, document_id)
            else:
                clauses = extract_clauses_from_text(content_text, document_id)
            if not normalize_text(content_text):
                extraction_issue = f"No extractable text was obtained from {path.name}. Use a searchable PDF, DOCX, TXT, or spreadsheet contract library."
            for clause in clauses:
                cur.execute(
                    """
                    INSERT INTO contract_clauses (
                        document_id, clause_number, clause_title, exact_clause_text, plain_english_meaning, section_name,
                        contractor_right, employer_obligation, engineer_obligation, claim_type, required_evidence,
                        possible_client_rejection, contractor_counterargument, risk_level, claim_strength, time_impact,
                        cost_impact, notice_required, recommended_action, related_project_records_needed, keywords_json,
                        created_at, updated_at
                    )
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        clause["document_id"], clause["clause_number"], clause["clause_title"], clause["exact_clause_text"],
                        clause["plain_english_meaning"], clause["section_name"], clause["contractor_right"], clause["employer_obligation"],
                        clause["engineer_obligation"], clause["claim_type"], clause["required_evidence"], clause["possible_client_rejection"],
                        clause["contractor_counterargument"], clause["risk_level"], clause["claim_strength"], clause["time_impact"],
                        clause["cost_impact"], clause["notice_required"], clause["recommended_action"], clause["related_project_records_needed"],
                        clause["keywords_json"], clause["created_at"], clause["updated_at"]
                    ),
                )
            total_clauses += len(clauses)
            cur.execute(
                """
                INSERT INTO contract_versions (document_id, file_hash, contract_version, analysis_date, is_active, created_at, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?)
                """,
                (document_id, file_hash, contract_version, stamp, 1, stamp, stamp),
            )
            cur.execute(
                """
                INSERT INTO contract_analysis_status (document_id, file_hash, last_analysis_date, processing_status, knowledge_base_status, total_clauses_extracted, error_message, created_at, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (document_id, file_hash, stamp, "Processed", "Ready", len(clauses), "", stamp, stamp),
            )
            summary = {
                "document_id": document_id,
                "file_name": path.name,
                "contract_version": contract_version,
                "total_clauses": len(clauses),
                "sections": list(pd.DataFrame(clauses)["section_name"].value_counts().index) if clauses else [],
            }
            cur.execute(
                """
                INSERT INTO contract_knowledge_base (document_id, knowledge_json, summary_text, knowledge_base_status, created_at, updated_at)
                VALUES (?, ?, ?, ?, ?, ?)
                """,
                (document_id, json.dumps(summary, ensure_ascii=True), f"{path.name}: {len(clauses)} clauses extracted", "Ready", stamp, stamp),
            )
            version_counter += 1
        kb_status_value = "Ready" if total_clauses > 0 else "Contract detected but no extractable text"
        cur.execute(
            """
            INSERT INTO contract_metadata (meta_key, meta_value, created_at, updated_at)
            VALUES (?, ?, ?, ?)
            ON CONFLICT(meta_key) DO UPDATE SET meta_value=excluded.meta_value, updated_at=excluded.updated_at
            """,
            ("knowledge_base_status", kb_status_value, stamp, stamp),
        )
        contract_status["reprocessed"] = True
        contract_status["extraction_issue"] = extraction_issue
    conn.commit()
    clauses_df = load_contract_library(db_path)
    analysis_status_df = load_contract_analysis_status(db_path)
    versions_df = load_contract_versions(db_path)
    knowledge_base_status = "Ready" if not clauses_df.empty else ("Contract detected but no extractable text" if files else "No contract detected")
    contract_status.update(
        {
            "contract_loaded": bool(files),
            "last_analysis_date": analysis_status_df["last_analysis_date"].max() if not analysis_status_df.empty else None,
            "total_clauses": int(len(clauses_df)),
            "contract_version": " / ".join(sorted(versions_df["contract_version"].dropna().astype(str).unique())) if not versions_df.empty else "N/A",
            "knowledge_base_status": knowledge_base_status,
            "source_warnings": source_warnings,
            "supporting_files": [path.name for path in supporting_files],
            "extraction_issue": contract_status.get("extraction_issue") or "; ".join(source_warnings),
            "auto_library_status": auto_library_status,
        }
    )
    conn.close()
    return contract_status


def dataframe_from_query(db_path: Path, query: str, params: tuple = ()) -> pd.DataFrame:
    conn = connect_db(db_path)
    df = pd.read_sql_query(query, conn, params=params)
    conn.close()
    return df


def load_contract_library(db_path: Path) -> pd.DataFrame:
    return dataframe_from_query(
        db_path,
        """
        SELECT c.*, d.file_name, d.contract_version
        FROM contract_clauses c
        LEFT JOIN contract_documents d ON d.id = c.document_id
        ORDER BY c.section_name, c.clause_number, c.id
        """,
    )


def load_contract_analysis_status(db_path: Path) -> pd.DataFrame:
    return dataframe_from_query(db_path, "SELECT * FROM contract_analysis_status ORDER BY last_analysis_date DESC")


def load_contract_versions(db_path: Path) -> pd.DataFrame:
    return dataframe_from_query(db_path, "SELECT * FROM contract_versions ORDER BY analysis_date DESC")


def load_evidence_documents(db_path: Path) -> pd.DataFrame:
    return dataframe_from_query(db_path, "SELECT * FROM evidence_documents ORDER BY upload_date DESC, id DESC")


def load_evidence_mappings(db_path: Path) -> pd.DataFrame:
    return dataframe_from_query(
        db_path,
        """
        SELECT em.*, ed.file_name, cc.clause_number, cc.clause_title, cc.section_name
        FROM evidence_mappings em
        LEFT JOIN evidence_documents ed ON ed.id = em.evidence_document_id
        LEFT JOIN contract_clauses cc ON cc.id = em.clause_id
        ORDER BY em.mapping_score DESC, em.id DESC
        """,
    )


def keyword_score(text: str, keywords: list[str]) -> int:
    lower = text.lower()
    return sum(1 for keyword in keywords if keyword in lower)


def clause_search_dataframe(clauses_df: pd.DataFrame, query: str, limit: int = 12) -> pd.DataFrame:
    if clauses_df.empty:
        return clauses_df
    query_lower = query.lower()
    tokens = [token for token in re.findall(r"[a-z0-9]+", query_lower) if len(token) > 2]
    late_design_query = any(term in query_lower for term in ["late ifc", "late drawing", "delayed drawing", "drawing", "rfi", "engineer response", "late approval"])
    eot_query = any(term in query_lower for term in ["eot", "extension of time", "delay", "time impact", "critical path"])
    notice_query = any(term in query_lower for term in ["notice", "time bar", "claim"])

    def score_row(row: pd.Series) -> int:
        haystack = " ".join(
            [
                normalize_text(row.get("clause_title", "")),
                normalize_text(row.get("exact_clause_text", "")),
                normalize_text(row.get("section_name", "")),
                normalize_text(row.get("claim_type", "")),
                normalize_text(row.get("possible_client_rejection", "")),
                normalize_text(row.get("contractor_counterargument", "")),
                normalize_text(row.get("required_evidence", "")),
                normalize_text(row.get("recommended_action", "")),
                normalize_text(row.get("related_project_records_needed", "")),
            ]
        ).lower()
        score = sum(haystack.count(token) for token in tokens)
        title = normalize_text(row.get("clause_title", "")).lower()
        section = normalize_text(row.get("section_name", "")).lower()
        claim_type = normalize_text(row.get("claim_type", "")).lower()

        if late_design_query:
            if any(term in title for term in ["delayed drawings", "drawing", "rfi", "engineer response", "shop drawings"]):
                score += 18
            if any(term in haystack for term in ["request dates", "late response proof", "critical path analysis", "submittal register", "approval delay"]):
                score += 10
            if any(term in title for term in ["design verification", "revit", "cad deliverables"]):
                score -= 8
        if eot_query:
            if "extension of time" in claim_type or "eot entitlements" in section:
                score += 12
            if any(term in haystack for term in ["critical path", "programme impact", "program impact", "time impact"]):
                score += 8
        if notice_query:
            if any(term in title for term in ["claims time bar", "monthly claim", "notice"]):
                score += 12
            if "notice requirements" in section or "evidence requirements" in section:
                score += 6
        return max(score, 0)

    scored = clauses_df.copy()
    scored["search_score"] = scored.apply(score_row, axis=1)
    strength_rank = {"Very Strong": 4, "Strong": 3, "Medium": 2, "Weak": 1}
    risk_rank = {"Critical": 4, "High": 3, "Medium": 2, "Low": 1}
    scored["claim_strength_rank"] = scored["claim_strength"].map(strength_rank).fillna(0)
    scored["risk_rank"] = scored["risk_level"].map(risk_rank).fillna(0)
    scored = scored[scored["search_score"] > 0].sort_values(["search_score", "claim_strength_rank", "risk_rank"], ascending=[False, False, False])
    return scored.head(limit)


def _unique_texts(values: list[Any]) -> list[str]:
    seen: set[str] = set()
    items: list[str] = []
    for value in values:
        text = normalize_text(value)
        if not text or text.lower() in seen:
            continue
        seen.add(text.lower())
        items.append(text)
    return items


def _truncate_text(value: Any, limit: int = 900) -> str:
    text = normalize_text(value)
    return text if len(text) <= limit else text[: limit - 3].rstrip() + "..."


def _compact_records(df: pd.DataFrame, columns: list[str], limit: int = 8, text_limit: int = 700) -> list[dict[str, str]]:
    if df.empty:
        return []
    available = [col for col in columns if col in df.columns]
    records: list[dict[str, str]] = []
    for _, row in df.head(limit).iterrows():
        record: dict[str, str] = {}
        for col in available:
            record[col] = _truncate_text(row.get(col, ""), text_limit)
        records.append(record)
    return records


def _baseline_for_ai(payload: dict[str, Any]) -> dict[str, Any]:
    compact: dict[str, Any] = {}
    for key, value in payload.items():
        if isinstance(value, pd.DataFrame):
            continue
        if isinstance(value, dict):
            compact[key] = {str(k): _truncate_text(v, 500) for k, v in value.items()}
        elif isinstance(value, list):
            compact[key] = [_truncate_text(item, 500) for item in value[:12]]
        else:
            compact[key] = _truncate_text(value, 900)
    return compact


def _json_block(value: Any) -> str:
    return json.dumps(value, ensure_ascii=True, default=str, indent=2)


def _merge_ai_text_fields(base: dict[str, Any], ai_data: dict[str, Any], allowed_fields: list[str]) -> dict[str, Any]:
    merged = dict(base)
    for field in allowed_fields:
        value = normalize_text(ai_data.get(field, ""))
        if value:
            merged[field] = value
    return merged


def _attach_ai_status(payload: dict[str, Any], result: Any) -> dict[str, Any]:
    enriched = dict(payload)
    enriched["ai_status"] = {
        "status": getattr(result, "status", "unknown"),
        "source": getattr(result, "source", "none"),
        "model": getattr(result, "model", ""),
        "latency_ms": getattr(result, "latency_ms", 0),
        "error": getattr(result, "error", ""),
    }
    return enriched


def classify_contract_question(query: str) -> list[str]:
    lower = normalize_text(query).lower()
    categories = [
        category
        for category, keywords in QUESTION_CATEGORY_RULES
        if any(keyword in lower for keyword in keywords)
    ]
    return categories or ["General Contract Risk"]


def _format_clause_reference(row: pd.Series) -> str:
    clause_number = normalize_text(row.get("clause_number", ""))
    clause_title = normalize_text(row.get("clause_title", ""))
    meaning = normalize_text(row.get("plain_english_meaning", ""))
    if clause_number and clause_title:
        label = f"Clause {clause_number} - {clause_title}"
    elif clause_number:
        label = f"Clause {clause_number}"
    elif clause_title:
        label = clause_title
    else:
        label = "Unnumbered clause extract"
    return f"{label}: {meaning}" if meaning else label


def _build_contract_basis_rows(matches: pd.DataFrame) -> list[str]:
    if matches.empty:
        return ["No matching clause found in the uploaded contract library."]
    return [_format_clause_reference(row) for _, row in matches.head(5).iterrows()]


def _category_defense_codes(categories: list[str]) -> list[str]:
    mapping = {
        "EOT / Delay Claim": ["NO_NOTICE", "NOT_CRITICAL_PATH", "CONCURRENT_DELAY", "NO_PROGRAMME_IMPACT", "NO_CAUSATION"],
        "Variation Claim": ["NO_WRITTEN_INSTRUCTION", "SCOPE_ALREADY_INCLUDED", "INSUFFICIENT_SUBSTANTIATION"],
        "Payment Claim": ["INSUFFICIENT_SUBSTANTIATION", "NO_COST_BREAKDOWN", "NO_CAUSATION"],
        "Suspension / Slowdown": ["NO_NOTICE", "FAILURE_TO_MITIGATE", "NO_CAUSATION"],
        "Material Escalation": ["FIXED_PRICE", "NO_COST_BREAKDOWN", "INSUFFICIENT_SUBSTANTIATION"],
        "Late Drawings / Late Approvals": ["NO_NOTICE", "NOT_CRITICAL_PATH", "NO_PROGRAMME_IMPACT"],
        "Employer Free-Issue Material": ["NO_NOTICE", "NOT_CRITICAL_PATH", "NO_CAUSATION"],
        "Site Access / Handover": ["NO_NOTICE", "CONTRACTOR_DELAY", "NO_CAUSATION"],
        "Engineer Instruction": ["NO_WRITTEN_INSTRUCTION", "ENGINEER_APPROVAL_NOT_ENTITLEMENT", "SCOPE_ALREADY_INCLUDED"],
        "Defects / DLP": ["CONTRACTOR_DELAY", "INSUFFICIENT_SUBSTANTIATION"],
        "Delay Damages Defense": ["CONCURRENT_DELAY", "CONTRACTOR_DELAY", "NOT_CRITICAL_PATH"],
        "Dispute / Notice / Time-Bar": ["NO_NOTICE", "LATE_NOTICE", "INSUFFICIENT_SUBSTANTIATION"],
        "General Contract Risk": ["INSUFFICIENT_SUBSTANTIATION", "NO_CAUSATION"],
    }
    codes: list[str] = []
    for category in categories:
        codes.extend(mapping.get(category, []))
    return list(dict.fromkeys(codes))


def _related_evidence_subset(
    matches: pd.DataFrame,
    evidence_df: pd.DataFrame,
    evidence_mappings_df: pd.DataFrame,
) -> pd.DataFrame:
    if evidence_df.empty:
        return pd.DataFrame()
    if matches.empty or evidence_mappings_df.empty or "clause_id" not in evidence_mappings_df.columns:
        return evidence_df.copy()
    clause_ids = set(matches["id"].dropna().astype(int).tolist())
    linked_ids = evidence_mappings_df[evidence_mappings_df["clause_id"].isin(clause_ids)]["evidence_document_id"].dropna().astype(int).tolist()
    if not linked_ids:
        return evidence_df.copy()
    return evidence_df[evidence_df["id"].isin(linked_ids)].copy()


def _evidence_signal_map(evidence_df: pd.DataFrame, query: str, matches: pd.DataFrame) -> dict[str, bool]:
    lower_query = normalize_text(query).lower()
    text_blob = " ".join(evidence_df.get("extracted_text", pd.Series(dtype=str)).fillna("").astype(str).str.lower().tolist())
    name_blob = " ".join(evidence_df.get("file_name", pd.Series(dtype=str)).fillna("").astype(str).str.lower().tolist())
    stream_blob = " ".join(evidence_df.get("source_stream", pd.Series(dtype=str)).fillna("").astype(str).str.lower().tolist())
    combined = f"{lower_query} {text_blob} {name_blob} {stream_blob}"

    def has_any(keywords: list[str]) -> bool:
        return any(keyword in combined for keyword in keywords)

    return {
        "contract_clause_exists": not matches.empty,
        "written_notice_exists": has_any(["notice", "letter", "email notice", "notify", "notification"]),
        "correspondence_exists": has_any(["email", "correspondence", "meeting minutes", "client rejection", "engineer reply", "letter"]),
        "approved_instruction_exists": has_any(["approved", "approval", "instruction", "ifc", "drawing", "site instruction", "engineer instruction"]),
        "programme_impact_exists": has_any(["programme impact", "schedule impact", "critical path", "longest path", "float", "primavera", "delay event"]),
        "cost_records_exist": has_any(["payment", "invoice", "cost", "quotation", "purchase order", "ipc", "certificate"]),
        "daily_records_exist": has_any(["daily report", "site record", "site diary", "weekly report", "progress report"]),
    }


def score_evidence_strength(matches: pd.DataFrame, evidence_df: pd.DataFrame, query: str) -> dict[str, Any]:
    signals = _evidence_signal_map(evidence_df, query, matches)
    score = sum(weight for code, _, weight in EVIDENCE_SCORE_RULES if signals.get(code, False))
    if score >= 80:
        label = "Very Strong"
    elif score >= 60:
        label = "Strong"
    elif score >= 40:
        label = "Medium"
    else:
        label = "Weak"
    missing_components = [label_text for code, label_text, _ in EVIDENCE_SCORE_RULES if not signals.get(code, False)]
    return {
        "score": int(score),
        "label": label,
        "signals": signals,
        "missing_components": missing_components,
    }


def _answer_decision(matches: pd.DataFrame, evidence_score: int) -> str:
    if matches.empty:
        return "NOT ENOUGH DATA"
    strengths = matches["claim_strength"].astype(str) if "claim_strength" in matches.columns else pd.Series(dtype=str)
    if evidence_score >= 80 and strengths.isin(["Very Strong", "Strong"]).any():
        return "YES"
    if evidence_score >= 40:
        return "POSSIBLE"
    if strengths.isin(["Very Strong", "Strong", "Medium"]).any():
        return "POSSIBLE"
    return "NO"


def _risk_assessment(matches: pd.DataFrame, evidence_profile: dict[str, Any]) -> str:
    if matches.empty:
        return "Critical"
    risk_levels = set(matches.get("risk_level", pd.Series(dtype=str)).fillna("").astype(str).str.lower().tolist())
    evidence_score = int(evidence_profile["score"])
    if "critical" in risk_levels or evidence_score < 40:
        return "Critical"
    if "high" in risk_levels or evidence_score < 60:
        return "High"
    if "medium" in risk_levels or evidence_score < 80:
        return "Medium"
    return "Low"


def detect_client_defenses(text: str) -> list[dict[str, Any]]:
    lower = normalize_text(text).lower()
    results: list[dict[str, Any]] = []
    for rule in DEFENSE_RULES:
        matched_keywords = [keyword for keyword in rule["keywords"] if keyword in lower]
        if matched_keywords:
            entry = dict(rule)
            entry["matched_keywords"] = matched_keywords
            entry["score"] = len(matched_keywords)
            results.append(entry)
    results.sort(key=lambda item: item["score"], reverse=True)
    return results


def detect_client_defense(text: str) -> dict[str, Any]:
    detected = detect_client_defenses(text)
    if detected:
        first = detected[0]
        return {
            "defense_code": first["code"],
            "defense_title": first["title"],
            "keywords": first["keywords"],
            "score": first["score"],
        }
    return {"defense_code": "UNKNOWN", "defense_title": "General Client Rejection", "keywords": [], "score": 0}


def _build_defense_rows(db_path: Path, text: str, defense_codes: list[str]) -> tuple[pd.DataFrame, pd.DataFrame]:
    clauses_df = load_contract_library(db_path)
    rows: list[dict[str, Any]] = []
    clause_frames: list[pd.DataFrame] = []
    rule_map = {rule["code"]: rule for rule in DEFENSE_RULES}
    for code in defense_codes:
        rule = rule_map.get(code)
        if not rule:
            continue
        matches = clause_search_dataframe(clauses_df, f"{text} {rule['title']} {' '.join(rule['keywords'])}", limit=5)
        clause_frames.append(matches)
        clause_ref = _build_contract_basis_rows(matches)[0] if not matches.empty else "No matching clause found in the uploaded contract library."
        evidence_needed = _unique_texts(matches.get("required_evidence", pd.Series(dtype=str)).tolist()) if not matches.empty else []
        counterarguments = _unique_texts(matches.get("contractor_counterargument", pd.Series(dtype=str)).tolist()) if not matches.empty else []
        suggested_wording = (
            f"The rejection based on '{rule['title']}' is not conclusive. "
            f"The contractor position is to rely on the contract wording, contemporaneous records, and the event chronology to defeat that defense."
        )
        if counterarguments:
            suggested_wording += " " + " ".join(counterarguments[:2])
        rows.append(
            {
                "Defense detected": rule["title"],
                "Why it is dangerous": rule["danger"],
                "Contract clause involved": clause_ref,
                "Evidence needed to defeat it": "; ".join(evidence_needed) if evidence_needed else "Potential entitlement exists, but evidence is incomplete.",
                "Contractor counterargument": "; ".join(counterarguments) if counterarguments else "Potential rebuttal exists, but the contractual and factual support must be assembled clearly.",
                "Suggested response wording": suggested_wording,
            }
        )
    combined_clauses = pd.concat(clause_frames, ignore_index=True).drop_duplicates(subset=["id"]) if clause_frames else pd.DataFrame()
    return pd.DataFrame(rows), combined_clauses


def _enhance_contract_answer_with_openai(
    db_path: Path,
    query: str,
    baseline: dict[str, Any],
    clauses_df: pd.DataFrame,
    evidence_df: pd.DataFrame,
) -> dict[str, Any]:
    clause_records = _compact_records(
        clauses_df,
        [
            "clause_number",
            "clause_title",
            "section_name",
            "claim_type",
            "risk_level",
            "claim_strength",
            "plain_english_meaning",
            "required_evidence",
            "possible_client_rejection",
            "contractor_counterargument",
            "recommended_action",
        ],
        limit=8,
    )
    evidence_records = _compact_records(
        evidence_df,
        ["file_name", "source_stream", "mapping_score", "mapping_basis", "missing_evidence_items"],
        limit=8,
    )
    input_payload = {
        "query": query,
        "baseline": _baseline_for_ai(baseline),
        "clauses": clause_records,
        "evidence": evidence_records,
    }
    result = create_structured_completion(
        db_path=db_path,
        feature_name="contract_question",
        system_prompt=CONTRACT_AI_SYSTEM_PROMPT,
        user_prompt=CONTRACT_QUESTION_PROMPT.format(
            query=query,
            baseline=_json_block(input_payload["baseline"]),
            clauses=_json_block(clause_records),
            evidence=_json_block(evidence_records),
        ),
        schema_name="contract_question_answer",
        schema=CLAIM_ANSWER_SCHEMA,
        input_payload=input_payload,
        required_keys=required_schema_keys(CLAIM_ANSWER_SCHEMA),
    )
    if result.ok and result.data:
        enhanced = _merge_ai_text_fields(
            baseline,
            result.data,
            [
                "short_answer",
                "contractor_friendly_interpretation",
                "required_evidence",
                "missing_evidence",
                "likely_client_rejection",
                "contractor_rebuttal",
                "recommended_next_action",
                "claim_strategy",
            ],
        )
        return _attach_ai_status(enhanced, result)
    return _attach_ai_status(baseline, result)


def _enhance_rebuttal_with_openai(
    db_path: Path,
    rejection_text: str,
    baseline: dict[str, Any],
    clauses_df: pd.DataFrame,
) -> dict[str, Any]:
    clause_records = _compact_records(
        clauses_df,
        [
            "clause_number",
            "clause_title",
            "section_name",
            "claim_type",
            "risk_level",
            "plain_english_meaning",
            "possible_client_rejection",
            "contractor_counterargument",
        ],
        limit=8,
    )
    input_payload = {
        "rejection_text": rejection_text,
        "baseline": _baseline_for_ai(baseline),
        "clauses": clause_records,
    }
    result = create_structured_completion(
        db_path=db_path,
        feature_name="client_rebuttal",
        system_prompt=CONTRACT_AI_SYSTEM_PROMPT,
        user_prompt=CLIENT_REBUTTAL_PROMPT.format(
            rejection_text=rejection_text,
            baseline=_json_block(input_payload["baseline"]),
            clauses=_json_block(clause_records),
        ),
        schema_name="client_rebuttal",
        schema=REBUTTAL_SCHEMA,
        input_payload=input_payload,
        required_keys=required_schema_keys(REBUTTAL_SCHEMA),
    )
    if result.ok and result.data:
        enhanced = _merge_ai_text_fields(
            baseline,
            result.data,
            [
                "client_argument_summary",
                "contractor_counterargument",
                "evidence_needed",
                "recommended_response_wording",
                "probability_of_success",
            ],
        )
        return _attach_ai_status(enhanced, result)
    return _attach_ai_status(baseline, result)


def _enhance_claim_draft_with_openai(
    db_path: Path,
    baseline: dict[str, Any],
    claim_type: str,
    delay_event: str,
    client_rejection_text: str,
    clause_subset: pd.DataFrame,
    evidence_subset: pd.DataFrame,
) -> dict[str, Any]:
    clause_records = _compact_records(
        clause_subset,
        ["clause_number", "clause_title", "section_name", "claim_type", "plain_english_meaning", "required_evidence"],
        limit=10,
    )
    evidence_records = _compact_records(evidence_subset, ["file_name", "source_stream", "extracted_text"], limit=8, text_limit=500)
    input_payload = {
        "claim_type": claim_type,
        "delay_event": delay_event,
        "client_rejection_text": client_rejection_text,
        "baseline": _baseline_for_ai(baseline),
        "clauses": clause_records,
        "evidence": evidence_records,
    }
    result = create_structured_completion(
        db_path=db_path,
        feature_name="claim_draft",
        system_prompt=CONTRACT_AI_SYSTEM_PROMPT,
        user_prompt=CLAIM_DRAFT_PROMPT.format(
            claim_type=claim_type,
            delay_event=delay_event,
            baseline=_json_block(input_payload["baseline"]),
            clauses=_json_block(clause_records),
            evidence=_json_block(evidence_records),
            client_rejection_text=client_rejection_text,
        ),
        schema_name="claim_draft",
        schema=CLAIM_DRAFT_SCHEMA,
        input_payload=input_payload,
        required_keys=required_schema_keys(CLAIM_DRAFT_SCHEMA),
    )
    if result.ok and result.data:
        enhanced = _merge_ai_text_fields(
            baseline,
            result.data,
            [
                "narrative_text",
                "factual_background",
                "cause_effect",
                "entitlement_statement",
                "time_impact_statement",
                "cost_impact_statement",
                "rebuttal_section",
                "attachment_checklist",
            ],
        )
        return _attach_ai_status(enhanced, result)
    return _attach_ai_status(baseline, result)


def answer_contract_question(db_path: Path, query: str) -> dict[str, Any]:
    clauses_df = load_contract_library(db_path)
    evidence_df = load_evidence_documents(db_path)
    mappings_df = load_evidence_mappings(db_path)
    categories = classify_contract_question(query)
    matches = clause_search_dataframe(clauses_df, query + " " + " ".join(categories), limit=10)
    related_evidence_df = _related_evidence_subset(matches, evidence_df, mappings_df)
    evidence_profile = score_evidence_strength(matches, related_evidence_df, query)
    decision = _answer_decision(matches, evidence_profile["score"])

    if matches.empty:
        baseline = {
            "question_categories": categories,
            "entitlement_decision": "NOT ENOUGH DATA",
            "short_answer": "No matching clause found in the uploaded contract library. Potential entitlement may exist, but the contract basis is not yet identifiable from the stored knowledge base.",
            "contract_basis_rows": ["No matching clause found in the uploaded contract library."],
            "relevant_clauses_df": pd.DataFrame(),
            "contractor_friendly_interpretation": "The first issue is contractual identification. Before making a claim decision, the exact entitlement wording must be located and verified in the stored contract library.",
            "required_evidence": "Relevant contract clauses, notices, event chronology, schedule impact records, and supporting project records.",
            "missing_evidence": "Potential entitlement exists, but evidence is incomplete.",
            "likely_client_rejection": "The client can reject the claim on the basis that no contractual entitlement has been identified.",
            "contractor_rebuttal": "Rebuild or expand the contract knowledge base, then identify the exact clause path before issuing a strong contractual position.",
            "risk_assessment": "Critical",
            "recommended_next_action": "Review the contract repository, rebuild the library if needed, and restate the question with the event type and contract context.",
            "claim_strategy": "Distinguish contractual entitlement first, then prove the facts, then prove programme impact and cost substantiation.",
            "evidence_strength_score": 0,
            "evidence_strength_label": "Weak",
            "evidence_missing_components": [label for _, label, _ in EVIDENCE_SCORE_RULES],
            "analysis_dimensions": {
                "Entitlement": "Not yet established because no matching clause has been identified.",
                "Proof": "No contract-linked proof path has been assembled yet.",
                "Delay Causation": "Cannot be assessed reliably until the entitlement path and supporting records are identified.",
                "Cost Substantiation": "Cannot be assessed reliably until the claim type and supporting valuation records are identified.",
                "Notice Compliance": "Must be checked once the relevant notice clause is identified.",
                "Rebuttal Strategy": "Start by identifying the clause basis and the evidence gap before drafting a rebuttal.",
            },
            "predicted_rebuttals_df": pd.DataFrame(),
        }
        return _enhance_contract_answer_with_openai(db_path, query, baseline, pd.DataFrame(), related_evidence_df)

    top = matches.head(5).copy()
    contract_basis_rows = _build_contract_basis_rows(top)
    required_evidence_items = _unique_texts(top.get("required_evidence", pd.Series(dtype=str)).tolist())
    missing_evidence_items = list(evidence_profile["missing_components"])
    if missing_evidence_items:
        missing_evidence_items.insert(0, "Potential entitlement exists, but evidence is incomplete.")
    likely_rejection_items = _unique_texts(top.get("possible_client_rejection", pd.Series(dtype=str)).tolist())
    defense_codes = _category_defense_codes(categories)
    predicted_rebuttals_df, predicted_clause_df = _build_defense_rows(db_path, query, defense_codes)
    predicted_clauses = predicted_clause_df if not predicted_clause_df.empty else top
    interpretation = " ".join(_unique_texts(top.get("plain_english_meaning", pd.Series(dtype=str)).tolist())[:4])
    contractor_rebuttal_text = "; ".join(
        _unique_texts(top.get("contractor_counterargument", pd.Series(dtype=str)).tolist())
        or ["The contractor case should be framed around the exact clause trigger, contemporaneous records, and the measured impact on programme and cost."]
    )
    risk_assessment = _risk_assessment(top, evidence_profile)
    short_answer = (
        f"{decision}. The current contract library contains relevant clause support under {', '.join(_unique_texts(top.get('section_name', pd.Series(dtype=str)).tolist())[:3])}. "
        f"The present evidence strength is {evidence_profile['label']} ({evidence_profile['score']}/100), so entitlement, proof, programme impact, and notice compliance must be addressed together."
    )
    next_action = "Validate notice compliance, assemble the missing evidence items, and tie the event to critical or near-critical activities before finalizing the claim position."
    claim_strategy = (
        "Separate the analysis into five tracks: contractual entitlement, factual proof, delay causation, cost substantiation, and rebuttal strategy. "
        "Use the matched clauses first, then connect the event chronology, correspondence, and programme impact records."
    )
    baseline = {
        "question_categories": categories,
        "entitlement_decision": decision,
        "short_answer": short_answer,
        "contract_basis_rows": contract_basis_rows,
        "relevant_clauses_df": predicted_clauses,
        "contractor_friendly_interpretation": interpretation or "The matched clauses support a contractor-led argument, but the final position depends on proof of trigger, impact, and compliance.",
        "required_evidence": "; ".join(required_evidence_items) if required_evidence_items else "Not enough evidence yet. Required documents are: notices, correspondence, programme impact records, and supporting factual documents.",
        "missing_evidence": "; ".join(missing_evidence_items) if missing_evidence_items else "No major evidence gap identified from the currently stored records.",
        "likely_client_rejection": "; ".join(_unique_texts(likely_rejection_items + predicted_rebuttals_df.get("Defense detected", pd.Series(dtype=str)).tolist())),
        "contractor_rebuttal": contractor_rebuttal_text,
        "risk_assessment": risk_assessment,
        "recommended_next_action": next_action,
        "claim_strategy": claim_strategy,
        "evidence_strength_score": evidence_profile["score"],
        "evidence_strength_label": evidence_profile["label"],
        "evidence_missing_components": evidence_profile["missing_components"],
        "analysis_dimensions": {
            "Entitlement": "Driven by the matched clauses and their claim-strength wording.",
            "Proof": "Driven by notices, correspondence, instructions, and contemporaneous site records mapped to the clauses.",
            "Delay Causation": "Must be proven through programme impact, critical path or longest path logic, and the event chronology.",
            "Cost Substantiation": "Requires valuation records, payment records, procurement records, or measured quantities depending on claim type.",
            "Notice Compliance": "Must be checked against the notice and time-bar clauses before relying on the substantive entitlement.",
            "Rebuttal Strategy": "Focus first on the likely client defenses, then defeat them with clause wording and evidence completeness.",
        },
        "predicted_rebuttals_df": predicted_rebuttals_df,
    }
    return _enhance_contract_answer_with_openai(db_path, query, baseline, predicted_clauses, related_evidence_df)


def persist_uploaded_evidence(db_path: Path, evidence_dir: Path, uploaded_files: list[Any], source_stream: str = "Manual Upload") -> list[str]:
    init_contract_claims_db(db_path)
    saved_names: list[str] = []
    conn = connect_db(db_path)
    cur = conn.cursor()
    clauses_df = load_contract_library(db_path)
    stamp = now_iso()
    for uploaded in uploaded_files:
        data = uploaded.getvalue()
        file_name = uploaded.name
        file_hash = compute_file_hash_from_bytes(data)
        target = evidence_dir / file_name
        target.write_bytes(data)
        text = extract_text_from_bytes(file_name, data)
        cur.execute(
            """
            INSERT INTO evidence_documents (file_name, file_path, file_hash, file_type, source_stream, upload_date, extracted_text, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (file_name, str(target), file_hash, Path(file_name).suffix.lower(), source_stream, stamp, text, stamp, stamp),
        )
        evidence_id = int(cur.lastrowid)
        mapping_rows = map_evidence_text_to_clauses(clauses_df, text, file_name)
        for mapping in mapping_rows:
            cur.execute(
                """
                INSERT INTO evidence_mappings (evidence_document_id, clause_id, claim_category, mapping_score, mapping_basis, missing_evidence_items, created_at, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (evidence_id, mapping["clause_id"], mapping["claim_category"], mapping["mapping_score"], mapping["mapping_basis"], mapping["missing_evidence_items"], stamp, stamp),
            )
        saved_names.append(file_name)
    conn.commit()
    conn.close()
    return saved_names


def map_evidence_text_to_clauses(clauses_df: pd.DataFrame, text: str, file_name: str) -> list[dict[str, Any]]:
    if clauses_df.empty:
        return []
    lower = text.lower()
    evidence_signals = {
        "notice": any(term in lower for term in ["notice", "reservation of rights", "time bar", "claim notice"]),
        "correspondence": any(term in lower for term in ["letter", "correspondence", "engineer response", "client response", "rfi"]),
        "programme": any(term in lower for term in ["programme", "program", "schedule", "critical path", "longest path", "time impact", "tia", "window"]),
        "cost": any(term in lower for term in ["cost", "payment", "valuation", "quantum", "invoice", "prolongation"]),
        "records": any(term in lower for term in ["contemporaneous", "records", "daily report", "site record", "mitigation"]),
        "late_design": any(term in lower for term in ["late ifc", "ifc", "late drawing", "delayed drawing", "late approval", "rfi response"]),
        "defense": any(term in lower for term in ["rejected", "no notice", "not on the critical path", "concurrent", "contractor delay"]),
    }
    rows: list[dict[str, Any]] = []
    for _, clause in clauses_df.iterrows():
        keywords = []
        try:
            keywords = json.loads(clause.get("keywords_json", "[]"))
        except Exception:
            keywords = []
        clause_title = normalize_text(clause.get("clause_title", ""))
        section_name = normalize_text(clause.get("section_name", ""))
        claim_type = normalize_text(clause.get("claim_type", ""))
        required_evidence = normalize_text(clause.get("required_evidence", ""))
        recommended_action = normalize_text(clause.get("recommended_action", ""))
        counterargument = normalize_text(clause.get("contractor_counterargument", ""))
        clause_context = " ".join([clause_title, section_name, claim_type, required_evidence, recommended_action, counterargument]).lower()
        score = keyword_score(lower, keywords)
        title_terms = [token for token in re.findall(r"[a-z0-9]+", clause_title.lower()) if len(token) > 3]
        score += keyword_score(lower, title_terms) * 2
        basis_parts = []
        if score > 0:
            basis_parts.append("keyword/title overlap")

        if evidence_signals["late_design"] and any(term in clause_context for term in ["delayed drawings", "late response proof", "drawing register", "submittal register", "rfi", "engineer response"]):
            score += 8
            basis_parts.append("late design/RFI signal")
        if evidence_signals["notice"] and any(term in clause_context for term in ["notice", "time bar", "monthly claim", "claims register"]):
            score += 7
            basis_parts.append("notice/time-bar signal")
        if evidence_signals["programme"] and any(term in clause_context for term in ["critical path", "programme", "program", "time impact", "extension of time", "eot"]):
            score += 7
            basis_parts.append("programme impact signal")
        if evidence_signals["cost"] and any(term in clause_context for term in ["cost", "payment", "valuation", "quantum", "prolongation"]):
            score += 5
            basis_parts.append("cost/valuation signal")
        if evidence_signals["records"] and any(term in clause_context for term in ["records", "contemporaneous", "daily", "mitigation", "particulars"]):
            score += 5
            basis_parts.append("records/mitigation signal")
        if evidence_signals["defense"] and any(term in clause_context for term in ["rebuttal", "counterargument", "concurrent", "contractor delay", "delay damages"]):
            score += 6
            basis_parts.append("client-defense signal")

        if score <= 0:
            continue
        required_components = []
        if "notice" in required_evidence.lower() and not evidence_signals["notice"]:
            required_components.append("notice")
        if any(term in required_evidence.lower() for term in ["programme", "program", "critical path", "delay analysis"]) and not evidence_signals["programme"]:
            required_components.append("programme impact")
        if any(term in required_evidence.lower() for term in ["cost", "payment", "valuation", "invoice"]) and not evidence_signals["cost"]:
            required_components.append("cost records")
        if any(term in required_evidence.lower() for term in ["records", "daily", "contemporaneous"]) and not evidence_signals["records"]:
            required_components.append("contemporaneous records")
        rows.append(
            {
                "clause_id": int(clause["id"]),
                "claim_category": clause.get("claim_type", "General Contract Risk"),
                "mapping_score": float(score),
                "mapping_basis": f"Evidence file '{file_name}' mapped by {', '.join(basis_parts) if basis_parts else 'context signal'} to clause {clause.get('clause_number', '') or clause_title}.",
                "missing_evidence_items": "; ".join(required_components),
            }
        )
    rows.sort(key=lambda item: item["mapping_score"], reverse=True)
    return rows[:25]


def build_client_rebuttal(db_path: Path, rejection_text: str) -> dict[str, Any]:
    detected_defenses = detect_client_defenses(rejection_text)
    detected_codes = [item["code"] for item in detected_defenses]
    defense_rows_df, matches = _build_defense_rows(db_path, rejection_text, detected_codes)
    clauses_df = load_contract_library(db_path)
    evidence_df = load_evidence_documents(db_path)
    mappings_df = load_evidence_mappings(db_path)
    if matches.empty:
        matches = clause_search_dataframe(clauses_df, rejection_text, limit=8)
    related_evidence_df = _related_evidence_subset(matches, evidence_df, mappings_df)
    evidence_profile = score_evidence_strength(matches, related_evidence_df, rejection_text)
    risk = _risk_assessment(matches, evidence_profile)
    probability = "Low"
    if evidence_profile["score"] >= 80 and not matches.empty:
        probability = "High"
    elif evidence_profile["score"] >= 40 and not matches.empty:
        probability = "Medium"
    client_argument_summary = ", ".join([item["title"] for item in detected_defenses]) if detected_defenses else "General Client Rejection"
    counterarguments = "; ".join(_unique_texts(defense_rows_df.get("Contractor counterargument", pd.Series(dtype=str)).tolist()))
    evidence_needed = "; ".join(_unique_texts(defense_rows_df.get("Evidence needed to defeat it", pd.Series(dtype=str)).tolist()))
    suggested_wording = "\n\n".join(_unique_texts(defense_rows_df.get("Suggested response wording", pd.Series(dtype=str)).tolist()))
    baseline = {
        "client_argument_summary": client_argument_summary,
        "contractual_risk": risk,
        "contractor_counterargument": counterarguments or "Potential rebuttal exists, but evidence is incomplete.",
        "evidence_needed": evidence_needed or "Potential entitlement exists, but evidence is incomplete.",
        "recommended_response_wording": suggested_wording or "Potential entitlement exists, but evidence is incomplete.",
        "probability_of_success": probability,
        "relevant_clauses_df": matches,
        "detected_defenses_df": defense_rows_df,
        "evidence_strength_score": evidence_profile["score"],
        "evidence_strength_label": evidence_profile["label"],
    }
    return _enhance_rebuttal_with_openai(db_path, rejection_text, baseline, matches)


def build_claim_draft_payload(
    db_path: Path,
    claim_type: str,
    delay_event: str,
    selected_clause_ids: list[int],
    selected_evidence_ids: list[int],
    client_rejection_text: str = "",
) -> dict[str, Any]:
    clauses_df = load_contract_library(db_path)
    evidence_df = load_evidence_documents(db_path)
    clause_subset = clauses_df[clauses_df["id"].isin(selected_clause_ids)].copy() if not clauses_df.empty else pd.DataFrame()
    evidence_subset = evidence_df[evidence_df["id"].isin(selected_evidence_ids)].copy() if not evidence_df.empty else pd.DataFrame()
    contractual_basis = "\n".join(
        f"- Clause {row.get('clause_number', '')}: {row.get('clause_title', '')} -> {row.get('plain_english_meaning', '')}"
        for _, row in clause_subset.iterrows()
    ) or "No clause selected."
    evidence_list = "\n".join(
        f"- {row.get('file_name', '')} ({row.get('source_stream', '')})"
        for _, row in evidence_subset.iterrows()
    ) or "No evidence selected."
    cause_effect = (
        f"The event '{delay_event}' is presented as a {claim_type} matter. The selected clauses provide the contractual path, "
        "and the selected records are the factual basis required to prove cause, effect, and entitlement."
    )
    rebuttal_section = ""
    if client_rejection_text.strip():
        rebuttal = build_client_rebuttal(db_path, client_rejection_text)
        rebuttal_section = (
            f"Client rejection summary: {rebuttal['client_argument_summary']}\n"
            f"Contractor counterargument: {rebuttal['contractor_counterargument']}\n"
            f"Evidence needed: {rebuttal['evidence_needed']}\n"
            f"Recommended response: {rebuttal['recommended_response_wording']}"
        )
    attachment_checklist = "\n".join(
        [
            "- Relevant contract clauses",
            "- Notices and correspondence",
            "- Event chronology",
            "- Schedule / delay analysis support",
            "- Commercial / cost support where applicable",
            "- Contemporaneous records and approvals",
        ]
    )
    payload = {
        "draft_name": f"{claim_type} - {delay_event}"[:180],
        "claim_type": claim_type,
        "delay_event": delay_event,
        "relevant_clause_ids": json.dumps(selected_clause_ids),
        "evidence_ids": json.dumps(selected_evidence_ids),
        "client_rejection_text": client_rejection_text,
        "narrative_text": (
            f"This draft claim addresses {delay_event}. It relies on the selected contractual clauses and the uploaded project evidence "
            "to establish entitlement, factual basis, and the required response strategy."
        ),
        "contractual_basis": contractual_basis,
        "factual_background": f"The claim arises from the event described as: {delay_event}. Supporting documents are listed below.",
        "cause_effect": cause_effect,
        "evidence_list": evidence_list,
        "entitlement_statement": f"The contractor position is that entitlement under {claim_type} is at least potentially available subject to final substantiation.",
        "time_impact_statement": "Time impact must be supported by CPM logic, TIA / window analysis, or other accepted schedule methodology where relevant.",
        "cost_impact_statement": "Cost impact must be supported by contractual basis, valuation records, payment / procurement records, or measured variation data where relevant.",
        "rebuttal_section": rebuttal_section or "No client rejection input was provided for rebuttal drafting.",
        "attachment_checklist": attachment_checklist,
    }
    payload = _enhance_claim_draft_with_openai(
        db_path,
        payload,
        claim_type,
        delay_event,
        client_rejection_text,
        clause_subset,
        evidence_subset,
    )
    save_claim_draft(db_path, payload)
    return payload


def get_contract_ai_test_cases() -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "Test Case": "Client rejected our variation because there is no written instruction. Can we still claim?",
                "Expected Decision": "POSSIBLE",
                "Expected Focus": "Constructive variation, revised drawings, engineer approval, quantity change, and missing written instruction risk.",
            },
            {
                "Test Case": "Can we claim EOT for late IFC drawings?",
                "Expected Decision": "POSSIBLE / YES",
                "Expected Focus": "Drawing log, baseline, updated programme, notice, and critical path impact.",
            },
            {
                "Test Case": "Steel increased from 35,000 to 52,000 EGP/ton. Can we claim?",
                "Expected Decision": "POSSIBLE / NO / YES depending on clause support",
                "Expected Focus": "Escalation clauses, fixed-price risks, invoices, index data, purchase orders, and timing.",
            },
            {
                "Test Case": "Your claim is rejected because no notice was submitted and the delay is not on the critical path.",
                "Expected Decision": "Multi-defense rebuttal required",
                "Expected Focus": "Detect No notice and Not on critical path, then produce separate rebuttals and an evidence checklist.",
            },
        ]
    )


def save_claim_draft(db_path: Path, payload: dict[str, Any]) -> None:
    conn = connect_db(db_path)
    cur = conn.cursor()
    stamp = now_iso()
    cur.execute(
        """
        INSERT INTO claim_drafts (
            draft_name, claim_type, delay_event, relevant_clause_ids, evidence_ids, client_rejection_text, narrative_text,
            contractual_basis, factual_background, cause_effect, evidence_list, entitlement_statement,
            time_impact_statement, cost_impact_statement, rebuttal_section, attachment_checklist, status, created_at, updated_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            payload.get("draft_name", ""),
            payload.get("claim_type", ""),
            payload.get("delay_event", ""),
            payload.get("relevant_clause_ids", "[]"),
            payload.get("evidence_ids", "[]"),
            payload.get("client_rejection_text", ""),
            payload.get("narrative_text", ""),
            payload.get("contractual_basis", ""),
            payload.get("factual_background", ""),
            payload.get("cause_effect", ""),
            payload.get("evidence_list", ""),
            payload.get("entitlement_statement", ""),
            payload.get("time_impact_statement", ""),
            payload.get("cost_impact_statement", ""),
            payload.get("rebuttal_section", ""),
            payload.get("attachment_checklist", ""),
            "Draft",
            stamp,
            stamp,
        ),
    )
    conn.commit()
    conn.close()


def load_claim_drafts(db_path: Path) -> pd.DataFrame:
    return dataframe_from_query(db_path, "SELECT * FROM claim_drafts ORDER BY created_at DESC, id DESC")


def build_contract_center_kpis(db_path: Path) -> dict[str, Any]:
    clauses_df = load_contract_library(db_path)
    mappings_df = load_evidence_mappings(db_path)
    status_df = load_contract_analysis_status(db_path)
    return {
        "total_clauses": int(len(clauses_df)),
        "contractor_rights": int(clauses_df["contractor_right"].astype(str).str.strip().ne("").sum()) if not clauses_df.empty else 0,
        "employer_obligations": int(clauses_df["employer_obligation"].astype(str).str.strip().eq("Yes").sum()) if not clauses_df.empty else 0,
        "high_risk_clauses": int(clauses_df["risk_level"].astype(str).isin(["High", "Critical"]).sum()) if not clauses_df.empty else 0,
        "strong_claim_opportunities": int(clauses_df["claim_strength"].astype(str).isin(["Strong", "Very Strong"]).sum()) if not clauses_df.empty else 0,
        "missing_evidence_items": int(mappings_df["missing_evidence_items"].astype(str).str.strip().ne("").sum()) if not mappings_df.empty else 0,
        "contract_status": "Contract Loaded" if not clauses_df.empty else "No Contract Loaded",
        "last_analysis_date": status_df["last_analysis_date"].max() if not status_df.empty else None,
    }


def build_contract_library_export_excel(clauses_df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        clauses_df.to_excel(writer, sheet_name="Contract Library", index=False)
    output.seek(0)
    return output.getvalue()


def build_contract_library_json_bytes(clauses_df: pd.DataFrame) -> bytes:
    return clauses_df.to_json(orient="records", force_ascii=True, indent=2).encode("utf-8")


def build_contract_library_html(clauses_df: pd.DataFrame) -> str:
    table_html = clauses_df.to_html(index=False, border=0) if not clauses_df.empty else "<p>No contract clauses are stored yet.</p>"
    return (
        "<html><head><meta charset='utf-8'><title>Contract & Claims Intelligence Center</title>"
        "<style>body{font-family:Arial,sans-serif;padding:24px;color:#172033;}table{border-collapse:collapse;width:100%;}th,td{border:1px solid #d1d5db;padding:8px;text-align:left;}th{background:#e5eef8;}</style>"
        "</head><body><h1>Contract & Claims Intelligence Center</h1>"
        "<h2>Stored Contract Library</h2>"
        + table_html
        + "</body></html>"
    )


def build_claim_draft_docx_bytes(payload: dict[str, Any]) -> bytes:
    Document = None
    try:
        Document = importlib.import_module("docx").Document
    except ModuleNotFoundError:
        return b""
    doc = Document()
    doc.add_heading("Contract Claim Draft", level=0)
    doc.add_paragraph(payload.get("draft_name", ""))
    sections = [
        ("Claim Narrative", payload.get("narrative_text", "")),
        ("Contractual Basis", payload.get("contractual_basis", "")),
        ("Factual Background", payload.get("factual_background", "")),
        ("Cause and Effect", payload.get("cause_effect", "")),
        ("Evidence List", payload.get("evidence_list", "")),
        ("Entitlement Statement", payload.get("entitlement_statement", "")),
        ("Time Impact Statement", payload.get("time_impact_statement", "")),
        ("Cost Impact Statement", payload.get("cost_impact_statement", "")),
        ("Rebuttal Section", payload.get("rebuttal_section", "")),
        ("Required Attachments Checklist", payload.get("attachment_checklist", "")),
    ]
    for title, body in sections:
        doc.add_heading(title, level=1)
        doc.add_paragraph(body)
    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output.getvalue()


def build_claim_draft_pdf_bytes(payload: dict[str, Any]) -> bytes:
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.pdfgen import canvas
    except ModuleNotFoundError:
        return b""
    output = io.BytesIO()
    pdf = canvas.Canvas(output, pagesize=A4)
    width, height = A4
    y = height - 50
    pdf.setFont("Helvetica-Bold", 15)
    pdf.drawString(40, y, "Contract Claim Draft")
    y -= 25
    pdf.setFont("Helvetica", 10)
    for title, body in [
        ("Draft", payload.get("draft_name", "")),
        ("Narrative", payload.get("narrative_text", "")),
        ("Contractual Basis", payload.get("contractual_basis", "")),
        ("Factual Background", payload.get("factual_background", "")),
        ("Cause and Effect", payload.get("cause_effect", "")),
        ("Evidence List", payload.get("evidence_list", "")),
        ("Entitlement Statement", payload.get("entitlement_statement", "")),
        ("Time Impact Statement", payload.get("time_impact_statement", "")),
        ("Cost Impact Statement", payload.get("cost_impact_statement", "")),
        ("Rebuttal", payload.get("rebuttal_section", "")),
    ]:
        pdf.setFont("Helvetica-Bold", 11)
        pdf.drawString(40, y, title)
        y -= 14
        pdf.setFont("Helvetica", 9)
        for line in str(body).splitlines():
            wrapped = re.findall(r".{1,110}(?:\s+|$)", line) or [line]
            for segment in wrapped:
                if y < 50:
                    pdf.showPage()
                    y = height - 50
                    pdf.setFont("Helvetica", 9)
                pdf.drawString(45, y, segment.strip())
                y -= 12
        y -= 10
    pdf.save()
    output.seek(0)
    return output.getvalue()
