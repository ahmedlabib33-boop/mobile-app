import json

import pandas as pd

from src.construction_system.project_catalog import (
    aggregate_project_csvs,
    discover_projects,
    ensure_project_structure,
    project_data_path,
)
from src.construction_system.project_context import build_project_context


def test_discovers_more_than_twenty_project_folders(tmp_path):
    for index in range(25):
        project_dir = tmp_path / f"PROJECT-{index:02d}"
        ensure_project_structure(project_dir)
        (project_dir / "project.json").write_text(
            json.dumps({"project_id": project_dir.name, "project_name": f"Project {index:02d}"}),
            encoding="utf-8",
        )
    ensure_project_structure(tmp_path / "_PROJECT_TEMPLATE")

    records = discover_projects(tmp_path)

    assert len(records) == 25
    assert all(not row["project_id"].startswith("_") for row in records)


def test_discovers_projects_nested_under_sector_folders(tmp_path):
    sector = tmp_path / "Infrastructure"
    project = sector / "Bridge Package"
    ensure_project_structure(project)
    (project / "project.json").write_text(
        json.dumps({"project_id": "BRIDGE-001", "project_name": "Bridge Package"}),
        encoding="utf-8",
    )
    ensure_project_structure(tmp_path / "_PROJECT_TEMPLATE")

    records = discover_projects(tmp_path)

    assert len(records) == 1
    assert records[0]["project_id"] == "BRIDGE-001"
    assert records[0]["sector_name"] == "Infrastructure"
    assert records[0]["project_relative_path"] == "Infrastructure/Bridge Package"


def test_project_data_paths_are_isolated(tmp_path):
    first = project_data_path(tmp_path, "P-01", "core", "activities.csv")
    second = project_data_path(tmp_path, "P-02", "core", "activities.csv")

    assert first != second
    assert first == tmp_path / "P-01" / "01-data" / "import_templates" / "activities.csv"


def test_aggregate_adds_project_id_when_source_omits_it(tmp_path):
    first = tmp_path / "p1.csv"
    second = tmp_path / "p2.csv"
    pd.DataFrame([{"activity_id": "A1"}]).to_csv(first, index=False)
    pd.DataFrame([{"activity_id": "A2", "project_id": ""}]).to_csv(second, index=False)

    result = aggregate_project_csvs([("P-01", first), ("P-02", second)])

    assert result["project_id"].tolist() == ["P-01", "P-02"]


def test_folder_rename_preserves_project_id_and_updates_context_path(tmp_path):
    original = tmp_path / "Original Folder"
    ensure_project_structure(original)
    (original / "project.json").write_text(
        json.dumps({"project_id": "STABLE-001", "project_name": "Stable Project"}),
        encoding="utf-8",
    )
    first = discover_projects(tmp_path)[0]
    renamed = tmp_path / "Renamed Folder"
    original.rename(renamed)

    second = discover_projects(tmp_path)[0]
    context = build_project_context(second, tmp_path)

    assert first["project_id"] == second["project_id"] == "STABLE-001"
    assert context.project_folder_name == "Renamed Folder"
    assert context.project_folder_path == renamed.resolve()


def test_standard_structure_and_samples_are_non_destructive(tmp_path):
    project = tmp_path / "P-01"
    ensure_project_structure(project)
    identity = project / "08-branding" / "project_identity_template.json"
    identity.write_text("user-owned", encoding="utf-8")

    ensure_project_structure(project)

    assert identity.read_text(encoding="utf-8") == "user-owned"
    assert (project / "06-evidence" / "contract_evidence_register_template.csv").exists()
    assert (project / "06-evidence" / "photo_log_template.csv").exists()
    assert (project / "09-notes" / "claims_notes_template.md").exists()
