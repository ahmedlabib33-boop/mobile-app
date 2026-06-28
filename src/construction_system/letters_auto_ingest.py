from __future__ import annotations

import re
from pathlib import Path
from typing import Callable

import pandas as pd


SAMCO_SHEET = "From Contractor"
ACE_SHEET = "From Consultant"
SAMCO_LINKS_SHEET = "Contractor Links"
ACE_LINKS_SHEET = "Consultant Links"
THREADS_SHEET = "Issue Threads"
AUTO_REGISTER_SHEET = "Auto Ingest Register"

SUPPORTED_EXTENSIONS = {".pdf", ".docx", ".txt", ".eml", ".csv", ".xlsx", ".xls"}

LETTER_COLUMNS = [
    "Ref No", "Date", "Type", "Subject", "Main Purpose", "Key Requests",
    "Scope Impact", "Responsibility", "Affected Activities", "Start Dependency",
    "Sequence Impact", "Commercial Impact", "Risk Type", "Risk Owner", "Delay Risk",
    "EOT Potential", "Claim Strength", "Required Actions",
]
SAMCO_LINK_COLUMNS = [
    "Thread", "SAMCO Ref No", "SAMCO Sheet Row", "SAMCO Date", "SAMCO Type",
    "SAMCO Subject", "Related ACE Ref No(s)", "ACE Sheet Row(s)", "ACE Date(s)",
    "ACE Subject(s)", "Relationship", "Confidence", "SAMCO Claim Strength",
    "ACE Claim Strength", "SAMCO Delay Risk", "ACE Delay Risk", "Recommended Follow-up",
]
ACE_LINK_COLUMNS = [
    "ACE Ref No", "ACE Sheet Row", "ACE Date", "ACE Type", "ACE Subject",
    "Related SAMCO Ref No(s)", "SAMCO Sheet Row(s)", "SAMCO Subject(s)", "Thread(s)",
    "ACE Delay Risk", "ACE EOT Potential", "ACE Claim Strength", "ACE Required Actions",
]
THREAD_COLUMNS = ["Thread", "SAMCO Ref(s)", "6", "Main Link", "Priority", "Next Action"]
REGISTER_COLUMNS = [
    "Source File", "Direction", "Reference", "Date", "Subject", "Classification",
    "Delay Risk", "Issue Thread", "Status", "Message",
]

FULL_REFERENCE_PATTERN = re.compile(
    r"\b(?:BD[-\s]?CW[-\s]?SAMCO[-\s]?ACE(?:PM)?[-\s]?LET[-\s]?STR[-\s]?\d{1,4}|"
    r"BD[-\s]?ACEPM[-\s]?SAMCO[-\s]?LET[-\s]?\d{1,4})\b",
    re.IGNORECASE,
)
DATE_PATTERN = re.compile(
    r"\b(?:\d{1,2}[\-/\.](?:\d{1,2}|[A-Za-z]{3,9})[\-/\.]\d{2,4}|"
    r"\d{1,2}\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[A-Za-z]*\s+\d{2,4})\b",
    re.IGNORECASE,
)


def folder_fingerprint(folder: Path) -> tuple[tuple[str, int, int], ...]:
    if not folder.exists():
        return ()
    rows = []
    for path in sorted(folder.rglob("*")):
        if path.is_file() and path.suffix.lower() in SUPPORTED_EXTENSIONS:
            stat = path.stat()
            rows.append((path.relative_to(folder).as_posix(), stat.st_mtime_ns, stat.st_size))
    return tuple(rows)


def _clean_reference(value: str) -> str:
    reference = re.sub(r"[\s_]+", "-", value.upper().strip())
    reference = re.sub(r"-+", "-", reference)
    reference = reference.replace("ACEPM-LET-STR", "ACE-LET-STR")
    return reference


def _direction_from_path(path: Path, text: str) -> str:
    folder_text = " ".join(part.lower() for part in path.parts)
    if "from samco" in folder_text or "from contractor" in folder_text:
        return "Contractor to Consultant"
    if "from acepm" in folder_text or "from ace" in folder_text or "from consultant" in folder_text or "from engineer" in folder_text:
        return "Consultant to Contractor"
    upper_text = text.upper()
    if "BD-CW-SAMCO-ACE" in upper_text:
        return "Contractor to Consultant"
    if "BD-ACEPM-SAMCO" in upper_text:
        return "Consultant to Contractor"
    return "Needs Review"


def _extract_reference(path: Path, text: str, direction: str) -> str:
    references = [_clean_reference(match.group(0)) for match in FULL_REFERENCE_PATTERN.finditer(f"{path.stem}\n{text}")]
    if direction == "Contractor to Consultant":
        own_reference = next((ref for ref in references if ref.startswith("BD-CW-SAMCO-ACE")), "")
        if own_reference:
            return own_reference
    elif direction == "Consultant to Contractor":
        own_reference = next((ref for ref in references if ref.startswith("BD-ACEPM-SAMCO")), "")
        if own_reference:
            return own_reference
    elif references:
        return references[0]
    leading_number = re.match(r"^\s*(\d{1,4})\b", path.stem)
    if not leading_number or direction == "Needs Review":
        return ""
    number = int(leading_number.group(1))
    folder_text = " ".join(part.lower() for part in path.parts)
    if "from samco" in folder_text:
        return f"BD-CW-SAMCO-ACE-LET-STR-{number:03d}"
    if "from acepm" in folder_text or "from ace" in folder_text:
        return f"BD-ACEPM-SAMCO-LET-{number:03d}"
    if direction == "Contractor to Consultant":
        return f"LTR-CTR-{number:03d}"
    return f"LTR-CNS-{number:03d}"


def _extract_date(text: str) -> str:
    match = DATE_PATTERN.search(text)
    if not match:
        return ""
    parsed = pd.to_datetime(match.group(0), errors="coerce", dayfirst=True)
    return parsed.strftime("%d-%b-%Y") if not pd.isna(parsed) else match.group(0)


def _extract_subject(path: Path, text: str) -> str:
    for line in text.splitlines()[:100]:
        match = re.match(r"\s*(?:subject|subj\.?|re)\s*[:\-]\s*(.+)", line, re.IGNORECASE)
        if match and match.group(1).strip():
            return match.group(1).strip()[:500]
    stem = re.sub(r"^\s*\d{1,4}\s*[-_.]*\s*", "", path.stem).strip(" -_")
    return stem or path.stem


def _classify(text: str) -> dict[str, str]:
    lowered = text.lower()
    if any(term in lowered for term in ["reinforcement steel", "free issue steel", "steel supply", "rft"]):
        return {
            "type": "Delay Notice",
            "risk": "Steel supply delay",
            "thread": "Reinforcement steel supply",
            "activities": "Reinforcement steel and affected structural activities",
            "action": "Link delivery evidence to affected activities and assess time impact.",
        }
    if any(term in lowered for term in ["ifc", "issued for construction", "drawing", "design change"]):
        return {
            "type": "Design / IFC Notice",
            "risk": "IFC / design delay",
            "thread": "IFC / design coordination",
            "activities": "Engineering, shop drawings, and downstream construction activities",
            "action": "Confirm IFC chronology, affected activities, and critical-path linkage.",
        }
    if any(term in lowered for term in ["rfi", "request for information", "delayed reply", "response delay"]):
        return {
            "type": "RFI / Response Notice",
            "risk": "RFI / response delay",
            "thread": "RFI / delayed response",
            "activities": "RFI-linked civil, structural, or MEP activities",
            "action": "Record submission and response dates and link the RFI to programme activities.",
        }
    if any(term in lowered for term in ["payment", "ipc", "invoice", "certificate", "certification"]):
        return {
            "type": "Payment Correspondence",
            "risk": "Payment delay",
            "thread": "Payments / IPCs / breakdown",
            "activities": "Payment certification and related resource or progress exposure",
            "action": "Verify due dates, certification, payment, and demonstrated schedule impact.",
        }
    if any(term in lowered for term in ["variation", "change order", "additional work", "new scope"]):
        return {
            "type": "Variation Notice",
            "risk": "Variation / scope change",
            "thread": "Variations / change control",
            "activities": "Changed or additional work activities",
            "action": "Confirm instruction, valuation, notice compliance, and programme effect.",
        }
    if any(term in lowered for term in ["delay", "extension of time", "eot", "late", "impact"]):
        return {
            "type": "Delay Notice",
            "risk": "General delay / EOT",
            "thread": "Delay notices / EOT",
            "activities": "Programme activities identified in the correspondence",
            "action": "Map the event to the programme, evidence, notices, and concurrency assessment.",
        }
    return {
        "type": "General Correspondence",
        "risk": "General correspondence",
        "thread": "General correspondence",
        "activities": "Programme linkage to be confirmed",
        "action": "Review the correspondence and confirm required action and activity linkage.",
    }


def _risk_values(text: str, classification: dict[str, str]) -> tuple[str, str, str]:
    lowered = text.lower()
    high = any(term in lowered for term in ["critical path", "extension of time", "eot", "work stoppage", "urgent", "material delay"])
    medium = any(term in lowered for term in ["delay", "late", "approval", "variation", "payment", "rfi", "ifc"])
    delay_risk = "High" if high else ("Medium" if medium else "Low")
    eot = "High" if high and classification["risk"] != "Payment delay" else ("Medium" if medium else "Low")
    strength = "Review Required" if delay_risk == "High" else "Preliminary"
    return delay_risk, eot, strength


def _related_references(text: str, current_reference: str, opposite_direction: str) -> list[str]:
    refs = []
    for match in FULL_REFERENCE_PATTERN.finditer(text):
        ref = _clean_reference(match.group(0))
        if not ref or ref == current_reference:
            continue
        is_ace = ref.startswith("BD-ACEPM-SAMCO")
        if (opposite_direction == "ACEPM" and is_ace) or (opposite_direction == "SAMCO" and not is_ace):
            if ref not in refs:
                refs.append(ref)
    return refs


def _ensure_frame(sheets: dict[str, pd.DataFrame], name: str, columns: list[str]) -> pd.DataFrame:
    legacy_aliases = {
        SAMCO_SHEET: "From SAMCO to ACE",
        ACE_SHEET: "From ACE to SAMCO",
        SAMCO_LINKS_SHEET: "SAMCO to ACE Links",
        ACE_LINKS_SHEET: "ACE to SAMCO Links",
    }
    source_name = name if name in sheets else legacy_aliases.get(name, name)
    frame = sheets.get(source_name, pd.DataFrame(columns=columns)).copy().fillna("")
    for column in columns:
        if column not in frame.columns:
            frame[column] = ""
    return frame


def merge_inbox_letters(
    source_sheets: dict[str, pd.DataFrame],
    inbox_dir: Path,
    extract_text: Callable[[Path], str],
) -> dict[str, pd.DataFrame]:
    sheets = {name: frame.copy().fillna("") for name, frame in source_sheets.items()}
    samco_df = _ensure_frame(sheets, SAMCO_SHEET, LETTER_COLUMNS)
    ace_df = _ensure_frame(sheets, ACE_SHEET, LETTER_COLUMNS)
    samco_links_df = _ensure_frame(sheets, SAMCO_LINKS_SHEET, SAMCO_LINK_COLUMNS)
    ace_links_df = _ensure_frame(sheets, ACE_LINKS_SHEET, ACE_LINK_COLUMNS)
    threads_df = _ensure_frame(sheets, THREADS_SHEET, THREAD_COLUMNS)
    register_rows = []

    files = [] if not inbox_dir.exists() else [
        path for path in sorted(inbox_dir.rglob("*"))
        if path.is_file() and path.suffix.lower() in SUPPORTED_EXTENSIONS
    ]
    for path in files:
        direction_hint = _direction_from_path(path, "")
        reference_hint = _extract_reference(path, "", direction_hint)
        hint_target = samco_df if direction_hint == "Contractor to Consultant" else ace_df
        if reference_hint and hint_target["Ref No"].astype(str).str.strip().eq(reference_hint).any():
            register_rows.append({
                "Source File": path.relative_to(inbox_dir).as_posix(), "Direction": direction_hint,
                "Reference": reference_hint, "Date": "", "Subject": path.stem,
                "Classification": "", "Delay Risk": "", "Issue Thread": "",
                "Status": "Already Registered", "Message": "The reference already exists in the intelligence workbook.",
            })
            continue
        try:
            text = str(extract_text(path) or "").strip()
        except Exception as exc:
            register_rows.append({
                "Source File": path.relative_to(inbox_dir).as_posix(), "Direction": "Needs Review",
                "Reference": "", "Date": "", "Subject": path.stem, "Classification": "",
                "Delay Risk": "", "Issue Thread": "", "Status": "Read Error", "Message": str(exc),
            })
            continue

        direction = _direction_from_path(path, text)
        reference = _extract_reference(path, text, direction)
        date = _extract_date(text)
        subject = _extract_subject(path, text)
        classification = _classify(f"{subject}\n{text}")
        delay_risk, eot_potential, claim_strength = _risk_values(text, classification)

        if direction == "Needs Review" or not reference:
            register_rows.append({
                "Source File": path.relative_to(inbox_dir).as_posix(), "Direction": direction,
                "Reference": reference, "Date": date, "Subject": subject,
                "Classification": classification["type"], "Delay Risk": delay_risk,
                "Issue Thread": classification["thread"], "Status": "Needs Review",
                "Message": "Place the file in a direction folder and include a reference number in the file or filename.",
            })
            continue

        target_df = samco_df if direction == "Contractor to Consultant" else ace_df
        if target_df["Ref No"].astype(str).str.strip().eq(reference).any():
            register_rows.append({
                "Source File": path.relative_to(inbox_dir).as_posix(), "Direction": direction,
                "Reference": reference, "Date": date, "Subject": subject,
                "Classification": classification["type"], "Delay Risk": delay_risk,
                "Issue Thread": classification["thread"], "Status": "Already Registered",
                "Message": "The reference already exists in the intelligence workbook.",
            })
            continue

        responsibility = "Engineer / Employer" if direction == "Contractor to Consultant" else "Contractor"
        row = {
            "Ref No": reference,
            "Date": date,
            "Type": classification["type"],
            "Subject": subject,
            "Main Purpose": subject,
            "Key Requests": classification["action"],
            "Scope Impact": "Review required from source correspondence",
            "Responsibility": responsibility,
            "Affected Activities": classification["activities"],
            "Start Dependency": "To be confirmed against programme logic",
            "Sequence Impact": delay_risk,
            "Commercial Impact": "Review required" if classification["risk"] != "General correspondence" else "Not identified",
            "Risk Type": classification["risk"],
            "Risk Owner": responsibility,
            "Delay Risk": delay_risk,
            "EOT Potential": eot_potential,
            "Claim Strength": claim_strength,
            "Required Actions": classification["action"],
        }
        target_df = pd.concat([target_df, pd.DataFrame([row])], ignore_index=True)
        if direction == "Contractor to Consultant":
            samco_df = target_df
            related = _related_references(text, reference, "ACEPM")
            link_row = {
                "Thread": classification["thread"], "SAMCO Ref No": reference,
                "SAMCO Sheet Row": len(samco_df) + 1, "SAMCO Date": date,
                "SAMCO Type": classification["type"], "SAMCO Subject": subject,
                "Related ACE Ref No(s)": "; ".join(related), "ACE Sheet Row(s)": "",
                "ACE Date(s)": "", "ACE Subject(s)": "",
                "Relationship": "Auto-linked from references found in the source letter" if related else "No related ACE reference detected",
                "Confidence": "High" if related else "Medium", "SAMCO Claim Strength": claim_strength,
                "ACE Claim Strength": "", "SAMCO Delay Risk": delay_risk, "ACE Delay Risk": "",
                "Recommended Follow-up": classification["action"],
            }
            samco_links_df = pd.concat([samco_links_df, pd.DataFrame([link_row])], ignore_index=True)
        else:
            ace_df = target_df
            related = _related_references(text, reference, "SAMCO")
            link_row = {
                "ACE Ref No": reference, "ACE Sheet Row": len(ace_df) + 1, "ACE Date": date,
                "ACE Type": classification["type"], "ACE Subject": subject,
                "Related SAMCO Ref No(s)": "; ".join(related), "SAMCO Sheet Row(s)": "",
                "SAMCO Subject(s)": "", "Thread(s)": classification["thread"],
                "ACE Delay Risk": delay_risk, "ACE EOT Potential": eot_potential,
                "ACE Claim Strength": claim_strength, "ACE Required Actions": classification["action"],
            }
            ace_links_df = pd.concat([ace_links_df, pd.DataFrame([link_row])], ignore_index=True)

        thread_match = threads_df["Thread"].astype(str).str.strip().eq(classification["thread"])
        short_ref = reference.split("-")[-1]
        if thread_match.any():
            index = thread_match[thread_match].index[0]
            ref_column = "SAMCO Ref(s)" if direction == "Contractor to Consultant" else "6"
            existing = [item.strip() for item in str(threads_df.at[index, ref_column]).split(";") if item.strip()]
            if short_ref not in existing:
                existing.append(short_ref)
            threads_df.at[index, ref_column] = "; ".join(existing)
            threads_df.at[index, "Priority"] = delay_risk
            threads_df.at[index, "Next Action"] = classification["action"]
        else:
            threads_df = pd.concat([threads_df, pd.DataFrame([{
                "Thread": classification["thread"],
                "SAMCO Ref(s)": short_ref if direction == "Contractor to Consultant" else "",
                "6": short_ref if direction == "Consultant to Contractor" else "",
                "Main Link": f"Automatically classified from {path.name}",
                "Priority": delay_risk,
                "Next Action": classification["action"],
            }])], ignore_index=True)

        register_rows.append({
            "Source File": path.relative_to(inbox_dir).as_posix(), "Direction": direction,
            "Reference": reference, "Date": date, "Subject": subject,
            "Classification": classification["type"], "Delay Risk": delay_risk,
            "Issue Thread": classification["thread"], "Status": "Added Automatically",
            "Message": "Added to the letter, link, and issue-thread tables.",
        })

    sheets[SAMCO_SHEET] = samco_df.fillna("")
    sheets[ACE_SHEET] = ace_df.fillna("")
    sheets[SAMCO_LINKS_SHEET] = samco_links_df.fillna("")
    sheets[ACE_LINKS_SHEET] = ace_links_df.fillna("")
    sheets[THREADS_SHEET] = threads_df.fillna("")
    sheets[AUTO_REGISTER_SHEET] = pd.DataFrame(register_rows, columns=REGISTER_COLUMNS).fillna("")
    return sheets
