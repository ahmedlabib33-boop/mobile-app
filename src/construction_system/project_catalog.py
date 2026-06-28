"""Folder-based project discovery and project-owned data paths."""

from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable

import pandas as pd


PROJECT_METADATA_FILE = "project.json"
PROJECT_MANIFEST_FILE = "project_manifest.json"
PROJECT_TEMPLATE_DIRNAME = "_PROJECT_TEMPLATE"
DEFAULT_SECTOR_NAME = "Unassigned"
PROJECT_SUBDIRECTORIES = (
    "01-data/import_templates",
    "02-delay_analysis/steel_delay_tia_templates",
    "02-delay_analysis/methodology",
    "03-schedule",
    "04-source_excel",
    "05-contracts/source",
    "05-contracts/clauses",
    "06-evidence",
    "07-letters_intelligence/inbox/From Contractor",
    "07-letters_intelligence/inbox/From Consultant",
    "08-branding",
    "09-notes",
    "10-deliverables",
    "11-outputs",
    "12-logs",
)


def safe_project_id(value: Any) -> str:
    """Return a filesystem-safe project id while preserving readable names."""
    text = str(value or "").strip()
    return "".join(ch for ch in text if ch not in '<>:"/\\|?*').strip(" .")


def safe_sector_name(value: Any) -> str:
    text = str(value or "").strip()
    return text or DEFAULT_SECTOR_NAME


def _read_project_csv_identity(project_dir: Path) -> dict[str, str]:
    """Read the user-maintained project identity from projects.csv when present."""
    for project_csv in (
        project_dir / "01-data" / "import_templates" / "projects.csv",
        project_dir / "data" / "import_templates" / "projects.csv",
    ):
        if not project_csv.exists():
            continue
        try:
            project_rows = pd.read_csv(project_csv, nrows=1).fillna("")
        except (OSError, UnicodeDecodeError, pd.errors.ParserError):
            continue
        if project_rows.empty:
            continue
        row = project_rows.iloc[0]
        return {
            "project_id": safe_project_id(row.get("project_id")),
            "project_name": str(row.get("project_name", "") or "").strip(),
            "client_name": str(row.get("client_name", "") or "").strip(),
            "contractor": str(row.get("contractor", "") or "").strip(),
            "currency": str(row.get("currency", "") or "").strip(),
            "status": str(row.get("status", "") or "").strip(),
            "sector_name": str(row.get("sector_name", "") or "").strip(),
        }
    return {}


def read_project_metadata(project_dir: Path, sector_dir: Path | None = None) -> dict[str, Any]:
    manifest_path = project_dir / PROJECT_MANIFEST_FILE
    metadata_path = project_dir / PROJECT_METADATA_FILE
    metadata: dict[str, Any] = {}
    if manifest_path.exists():
        try:
            loaded = json.loads(manifest_path.read_text(encoding="utf-8"))
            if isinstance(loaded, dict):
                metadata.update(loaded)
        except (OSError, UnicodeDecodeError, json.JSONDecodeError):
            pass
    if metadata_path.exists():
        try:
            loaded = json.loads(metadata_path.read_text(encoding="utf-8"))
            if isinstance(loaded, dict):
                for key, value in loaded.items():
                    metadata.setdefault(key, value)
        except (OSError, UnicodeDecodeError, json.JSONDecodeError):
            metadata = {}

    csv_identity = _read_project_csv_identity(project_dir)
    project_id = safe_project_id(csv_identity.get("project_id") or metadata.get("project_id")) or project_dir.name
    project_name = str(
        csv_identity.get("project_name")
        or metadata.get("project_display_name")
        or metadata.get("project_name")
        or project_dir.name
    ).strip()
    sector_name = safe_sector_name(csv_identity.get("sector_name") or metadata.get("sector_name") or (sector_dir.name if sector_dir else DEFAULT_SECTOR_NAME))
    sector_id = safe_project_id(metadata.get("sector_id") or sector_name) or DEFAULT_SECTOR_NAME
    relative_label = f"{sector_dir.name}/{project_dir.name}" if sector_dir else project_dir.name
    return {
        **metadata,
        "project_id": project_id,
        "project_name": project_name,
        "project_display_name": project_name,
        "project_folder_name": project_dir.name,
        "project_dir": str(project_dir),
        "project_relative_path": relative_label,
        "sector_id": sector_id,
        "sector_name": sector_name,
        "sector_folder_name": sector_dir.name if sector_dir else "",
        "sector_dir": str(sector_dir) if sector_dir else "",
    }


def _is_ignored_discovery_folder(path: Path) -> bool:
    return not path.is_dir() or path.name.startswith(("_", "."))


def _looks_like_project_folder(path: Path) -> bool:
    indicators = (
        PROJECT_MANIFEST_FILE,
        PROJECT_METADATA_FILE,
        "01-data",
        "02-delay_analysis",
        "08-branding",
        "05-contracts",
        "data",
        "delay_analysis",
        "1-branding",
        "2-contracts",
        "project.json",
    )
    return any((path / indicator).exists() for indicator in indicators)


def _iter_project_folders(projects_root: Path) -> Iterable[tuple[Path, Path | None]]:
    """Yield root projects and sector-contained projects without hardcoded registration."""
    for path in sorted(projects_root.iterdir(), key=lambda item: item.name.casefold()):
        if _is_ignored_discovery_folder(path):
            continue
        child_dirs = [child for child in path.iterdir() if child.is_dir() and not child.name.startswith(("_", "."))]
        if _looks_like_project_folder(path):
            yield path, None
            continue
        for child in sorted(child_dirs, key=lambda item: item.name.casefold()):
            yield child, path


def discover_projects(projects_root: Path) -> list[dict[str, Any]]:
    """Discover root projects and projects nested under sector folders."""
    if not projects_root.exists():
        return []
    records = []
    for path, sector_dir in _iter_project_folders(projects_root):
        ensure_project_manifest(path)
        ensure_project_structure(path)
        records.append(read_project_metadata(path, sector_dir=sector_dir))
    return sorted(records, key=lambda row: (str(row.get("sector_name", "")).casefold(), str(row["project_name"]).casefold(), str(row["project_id"]).casefold()))


def projects_frame(projects_root: Path) -> pd.DataFrame:
    records = discover_projects(projects_root)
    if not records:
        return pd.DataFrame(columns=["project_id", "project_name", "project_dir", "sector_id", "sector_name"])
    return pd.DataFrame(records)


def project_directory(projects_root: Path, project_id: str) -> Path:
    clean_id = safe_project_id(project_id)
    for record in discover_projects(projects_root):
        if str(record.get("project_id", "")).strip().casefold() == clean_id.casefold():
            return Path(str(record["project_dir"]))
    return projects_root / clean_id


def project_data_path(projects_root: Path, project_id: str, family: str, relative_path: Path | str) -> Path:
    family_roots = {
        "core": Path("01-data/import_templates"),
        "delay_analysis": Path("02-delay_analysis/steel_delay_tia_templates"),
        "bl": Path("03-schedule"),
        "fixed": Path("03-schedule"),
        "letters": Path("07-letters_intelligence"),
        "exports": Path("11-outputs"),
        "outputs": Path("11-outputs"),
        "reports": Path("10-deliverables"),
        "slides": Path("10-deliverables"),
        "branding": Path("08-branding"),
        "contracts": Path("05-contracts"),
        "evidence": Path("06-evidence"),
        "notes": Path("09-notes"),
    }
    if family not in family_roots:
        raise ValueError(f"Unknown project data family: {family}")
    return project_directory(projects_root, project_id) / family_roots[family] / Path(relative_path)


def ensure_project_structure(project_dir: Path) -> None:
    project_dir.mkdir(parents=True, exist_ok=True)
    for relative_dir in PROJECT_SUBDIRECTORIES:
        folder = project_dir / relative_dir
        folder.mkdir(parents=True, exist_ok=True)
        if not any(folder.iterdir()):
            (folder / ".gitkeep").touch(exist_ok=True)
    ensure_project_samples(project_dir)


def ensure_project_samples(project_dir: Path) -> None:
    """Create non-destructive examples for the standard user-managed folders."""
    project_id = project_dir.name
    manifest_path = project_dir / PROJECT_MANIFEST_FILE
    if manifest_path.exists():
        try:
            project_id = safe_project_id(json.loads(manifest_path.read_text(encoding="utf-8")).get("project_id")) or project_id
        except (OSError, UnicodeDecodeError, json.JSONDecodeError, AttributeError):
            pass
    samples = {
        "08-branding/README.md": "# Branding\n\nReplace example files with project-approved brand assets. Existing files are never overwritten.\n",
        "08-branding/project_identity_template.json": json.dumps({"project_name": "Example Project", "client": "Example Client", "contractor": "Example Contractor", "consultant": "Example Consultant"}, indent=2) + "\n",
        "08-branding/color_palette_template.csv": f"project_id,role,color_hex\n{project_id},primary,#123B5D\n{project_id},secondary,#5B6870\n{project_id},accent,#C7A34B\n{project_id},background,#FFFFFF\n",
        "08-branding/report_branding_template.json": json.dumps({"header_text": "Example Project", "footer_text": "Confidential", "logo_file": "logo.png"}, indent=2) + "\n",
        "08-branding/logo_placeholder.txt": "Place the approved project logo here as logo.png.\n",
        "06-evidence/contract_evidence_register_template.csv": "project_id,evidence_id,contract_clause,document_reference,event_id,description,status\n",
        "06-evidence/clause_reference_template.csv": "project_id,clause_id,clause_title,source_file,source_page,entitlement_topic,notes\n",
        "06-evidence/correspondence_reference_template.csv": "project_id,reference,date,from_party,to_party,subject,related_event_id,evidence_status\n",
        "06-evidence/evidence_register_template.csv": "project_id,evidence_id,evidence_type,source_file,source_date,event_id,activity_id,description,verified\n",
        "06-evidence/photo_log_template.csv": "project_id,photo_id,date,location,activity_id,event_id,file_name,caption,taken_by\n",
        "06-evidence/document_reference_template.csv": "project_id,document_id,document_type,reference,date,source_file,related_event_id,notes\n",
        "09-notes/meeting_notes_template.md": "# Meeting Notes\n\n- Date:\n- Attendees:\n- Decisions:\n- Actions:\n- Related event or activity:\n",
        "09-notes/engineering_notes_template.md": "# Engineering Notes\n\n- Date:\n- Discipline:\n- Drawing / RFI reference:\n- Constraint:\n- Required action:\n",
        "09-notes/claims_notes_template.md": "# Claims Notes\n\n- Date:\n- Event ID:\n- Clause reference:\n- Notice status:\n- Cause and effect:\n- Missing evidence:\n",
    }
    for relative_path, content in samples.items():
        path = project_dir / relative_path
        if not path.exists():
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(content, encoding="utf-8")
    readme_path = project_dir / "README.md"
    if not readme_path.exists():
        readme_path.write_text(
            f"# {project_dir.name}\n\nThis folder is automatically discovered by Project Intelligence Hub. Edit project_manifest.json and project-owned data only.\n",
            encoding="utf-8",
        )


def ensure_project_manifest(project_dir: Path) -> Path:
    """Create a stable manifest or refresh only rename-sensitive path fields."""
    manifest_path = project_dir / PROJECT_MANIFEST_FILE
    now = datetime.now(timezone.utc).isoformat()
    existing: dict[str, Any] = {}
    if manifest_path.exists():
        try:
            loaded = json.loads(manifest_path.read_text(encoding="utf-8"))
            if isinstance(loaded, dict):
                existing = loaded
        except (OSError, UnicodeDecodeError, json.JSONDecodeError):
            existing = {}
    legacy: dict[str, Any] = {}
    legacy_path = project_dir / PROJECT_METADATA_FILE
    if legacy_path.exists():
        try:
            loaded = json.loads(legacy_path.read_text(encoding="utf-8"))
            if isinstance(loaded, dict):
                legacy = loaded
        except (OSError, UnicodeDecodeError, json.JSONDecodeError):
            legacy = {}
    csv_identity = _read_project_csv_identity(project_dir)
    project_id = safe_project_id(csv_identity.get("project_id") or existing.get("project_id") or legacy.get("project_id") or project_dir.name)
    folder_changed = (
        str(existing.get("project_folder_name", "")) != project_dir.name
        or str(existing.get("project_folder_path", "")) != str(project_dir.resolve())
    )
    display_name = str(
        csv_identity.get("project_name")
        or existing.get("project_display_name")
        or legacy.get("project_name")
        or project_dir.name
    )
    status = str(csv_identity.get("status") or existing.get("status") or legacy.get("status") or "Setup")
    identity_changed = (
        str(existing.get("project_id", "")) != project_id
        or str(existing.get("project_display_name", "")) != display_name
        or str(existing.get("status", "")) != status
    )
    manifest = {
        **existing,
        "project_id": project_id,
        "project_display_name": display_name,
        "project_folder_name": project_dir.name,
        "project_folder_path": str(project_dir.resolve()),
        "created_at": existing.get("created_at") or now,
        "updated_at": now if folder_changed or identity_changed or not existing else existing.get("updated_at", now),
        "status": status,
    }
    if manifest != existing:
        manifest_path.write_text(json.dumps(manifest, indent=2, ensure_ascii=True) + "\n", encoding="utf-8")
    return manifest_path


def aggregate_project_csvs(paths: Iterable[tuple[str, Path]]) -> pd.DataFrame:
    frames: list[pd.DataFrame] = []
    for project_id, path in paths:
        if not path.exists():
            continue
        try:
            frame = pd.read_csv(path).fillna("")
        except (OSError, UnicodeDecodeError, pd.errors.ParserError):
            continue
        if "project_id" not in frame.columns:
            frame.insert(0, "project_id", project_id)
        else:
            source_ids = frame["project_id"].astype(str).str.strip()
            if "source_project_id" not in frame.columns:
                frame.insert(1, "source_project_id", source_ids)
            frame["project_id"] = project_id
        frames.append(frame)
    return pd.concat(frames, ignore_index=True, sort=False) if frames else pd.DataFrame()
