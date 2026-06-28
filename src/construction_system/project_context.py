"""Central selected-project context used by every project-specific workflow."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any


@dataclass(frozen=True)
class ProjectContext:
    project_id: str
    project_key: str
    project_display_name: str
    project_folder_name: str
    project_folder_path: Path
    is_all_projects: bool
    sector_id: str
    sector_name: str
    sector_folder_name: str
    sector_folder_path: Path
    data_path: Path
    outputs_path: Path
    reports_path: Path
    slides_path: Path
    exports_path: Path
    logs_path: Path
    branding_path: Path
    contracts_path: Path
    contracts_evidence_path: Path
    evidence_path: Path
    notes_path: Path

    @property
    def cache_key(self) -> str:
        try:
            modified_ns = self.project_folder_path.stat().st_mtime_ns
        except OSError:
            modified_ns = 0
        return f"{self.project_id}:{self.project_folder_name}:{modified_ns}"

    def require_project(self, feature_name: str) -> None:
        if self.is_all_projects:
            raise ValueError(f"{feature_name} requires a selected project.")


def build_project_context(record: dict[str, Any] | None, projects_root: Path) -> ProjectContext:
    if not record:
        portfolio_root = projects_root.parent / "11-outputs" / "portfolio_scope"
        return ProjectContext(
            project_id="",
            project_key="portfolio",
            project_display_name="All Projects",
            project_folder_name="",
            project_folder_path=portfolio_root,
            is_all_projects=True,
            sector_id="portfolio",
            sector_name="Decision Making dashboard",
            sector_folder_name="",
            sector_folder_path=projects_root,
            data_path=portfolio_root / "01-data",
            outputs_path=portfolio_root / "11-outputs",
            reports_path=portfolio_root / "10-deliverables",
            slides_path=portfolio_root / "10-deliverables",
            exports_path=portfolio_root / "11-outputs",
            logs_path=portfolio_root / "12-logs",
            branding_path=portfolio_root / "08-branding",
            contracts_path=portfolio_root / "05-contracts",
            contracts_evidence_path=portfolio_root / "06-evidence",
            evidence_path=portfolio_root / "06-evidence",
            notes_path=portfolio_root / "09-notes",
        )

    folder_path = Path(str(record["project_dir"])).resolve()
    project_id = str(record["project_id"]).strip()
    display_name = str(record.get("project_name") or record.get("project_display_name") or folder_path.name).strip()
    sector_folder = Path(str(record.get("sector_dir") or folder_path.parent)).resolve()
    return ProjectContext(
        project_id=project_id,
        project_key=project_id.casefold(),
        project_display_name=display_name,
        project_folder_name=folder_path.name,
        project_folder_path=folder_path,
        is_all_projects=False,
        sector_id=str(record.get("sector_id") or "").strip(),
        sector_name=str(record.get("sector_name") or "").strip(),
        sector_folder_name=str(record.get("sector_folder_name") or "").strip(),
        sector_folder_path=sector_folder,
        data_path=folder_path / "01-data",
        outputs_path=folder_path / "11-outputs",
        reports_path=folder_path / "10-deliverables",
        slides_path=folder_path / "10-deliverables",
        exports_path=folder_path / "11-outputs",
        logs_path=folder_path / "12-logs",
        branding_path=folder_path / "08-branding",
        contracts_path=folder_path / "05-contracts",
        contracts_evidence_path=folder_path / "06-evidence",
        evidence_path=folder_path / "06-evidence",
        notes_path=folder_path / "09-notes",
    )
