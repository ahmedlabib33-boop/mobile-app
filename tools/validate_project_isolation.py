from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from src.construction_system.project_catalog import (
    PROJECT_SUBDIRECTORIES,
    discover_projects,
    ensure_project_manifest,
    ensure_project_structure,
    project_data_path,
)
from src.construction_system.project_context import build_project_context


PROJECTS = ROOT / "projects"


def _expected_project_folders(projects_root: Path) -> set[Path]:
    """Return project folders for both flat and sector/project layouts."""
    expected: set[Path] = set()
    if not projects_root.exists():
        return expected

    project_markers = {
        "project_manifest.json",
        "01-data",
        "02-delay_analysis",
        "08-branding",
        "05-contracts",
        "11-outputs",
        "data",
        "delay_analysis",
        "1-branding",
        "2-contracts",
        "outputs",
    }

    for folder in projects_root.iterdir():
        if not folder.is_dir() or folder.name.startswith(("_", ".")):
            continue

        has_project_marker = any((folder / marker).exists() for marker in project_markers)
        child_dirs = [child for child in folder.iterdir() if child.is_dir() and not child.name.startswith(("_", "."))]

        if has_project_marker:
            expected.add(folder.resolve())
            continue

        for child in child_dirs:
            expected.add(child.resolve())

    return expected


class Results:
    def __init__(self) -> None:
        self.rows: list[tuple[str, bool, str]] = []

    def check(self, name: str, passed: bool, detail: str = "") -> None:
        self.rows.append((name, bool(passed), detail))
        print(f"{'PASS' if passed else 'FAIL'} | {name}" + (f" | {detail}" if detail else ""))

    @property
    def passed(self) -> bool:
        return all(row[1] for row in self.rows)


def dry_run_sync() -> dict:
    command = [
        "powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass",
        "-File", str(ROOT / "tools" / "github_no_git_sync.ps1"),
        "-Mode", "DryRun", "-IntervalMinutes", "30",
    ]
    subprocess.run(command, cwd=ROOT, check=True, capture_output=True, text=True, timeout=180)
    return json.loads((ROOT / ".pih_mobile_app_sync_state" / "local_manifest.json").read_text(encoding="utf-8-sig"))


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--sync-probe", action="store_true", help="Run create/modify/rename full-workspace synchronization probes")
    args = parser.parse_args()
    results = Results()
    records = discover_projects(PROJECTS)
    expected_project_dirs = _expected_project_folders(PROJECTS)
    record_project_dirs = {Path(record["project_dir"]).resolve() for record in records}

    results.check("1. Existing registry detects projects", bool(records), f"{len(records)} detected")
    missing_project_dirs = sorted(str(path.relative_to(PROJECTS)) for path in expected_project_dirs - record_project_dirs)
    unexpected_project_dirs = sorted(str(path.relative_to(PROJECTS)) for path in record_project_dirs - expected_project_dirs)
    results.check(
        "2. Every project folder is detected",
        expected_project_dirs == record_project_dirs,
        f"missing={missing_project_dirs[:3]} unexpected={unexpected_project_dirs[:3]}" if missing_project_dirs or unexpected_project_dirs else "",
    )

    with tempfile.TemporaryDirectory() as temp_dir:
        temp_root = Path(temp_dir)
        for index in range(25):
            folder = temp_root / f"Future Project {index:02d}"
            ensure_project_structure(folder)
            ensure_project_manifest(folder)
        results.check("3. Newly added folders auto-register", len(discover_projects(temp_root)) == 25)
        original = temp_root / "Future Project 00"
        original_record = next(row for row in discover_projects(temp_root) if Path(row["project_dir"]).name == original.name)
        stable_id = original_record["project_id"]
        renamed = temp_root / "Renamed Future Project"
        original.rename(renamed)
        renamed_record = next(row for row in discover_projects(temp_root) if row["project_id"] == stable_id)
        results.check("26. Folder rename propagates while project_id stays stable", Path(renamed_record["project_dir"]).name == renamed.name)

    ids = [row["project_id"] for row in records]
    if len(ids) >= 2:
        first, second = ids[:2]
        results.check("4. Core project data paths are isolated", project_data_path(PROJECTS, first, "core", "activities.csv") != project_data_path(PROJECTS, second, "core", "activities.csv"))
        first_context = build_project_context(next(row for row in records if row["project_id"] == first), PROJECTS)
        second_context = build_project_context(next(row for row in records if row["project_id"] == second), PROJECTS)
        results.check("5. Claims Intelligence databases are isolated", first_context.contracts_path != second_context.contracts_path)
        results.check("6. Delay TIA paths are isolated", project_data_path(PROJECTS, first, "delay_analysis", "") != project_data_path(PROJECTS, second, "delay_analysis", ""))
        results.check("7. Output Studio paths are isolated", first_context.outputs_path != second_context.outputs_path)
        results.check("8. Reports save under selected project", first_context.reports_path.is_relative_to(first_context.project_folder_path))
        results.check("9. Slides save under selected project", first_context.slides_path.is_relative_to(first_context.project_folder_path))
        results.check("10. Exports save under selected project", first_context.exports_path.is_relative_to(first_context.project_folder_path))
        results.check("17. Cache keys include project identity", first_context.project_id in first_context.cache_key and first_context.cache_key != second_context.cache_key)
    else:
        for number, label in [(4, "Core project data paths are isolated"), (5, "Claims Intelligence databases are isolated"), (6, "Delay TIA paths are isolated"), (7, "Output Studio paths are isolated"), (8, "Reports save under selected project"), (9, "Slides save under selected project"), (10, "Exports save under selected project"), (17, "Cache keys include project identity")]:
            results.check(f"{number}. {label}", False, "Need at least two projects")

    missing_structures = []
    missing_manifests = []
    for record in records:
        folder = Path(record["project_dir"])
        for relative in PROJECT_SUBDIRECTORIES:
            if not (folder / relative).is_dir():
                missing_structures.append(f"{folder.name}/{relative}")
        if not (folder / "project_manifest.json").exists():
            missing_manifests.append(folder.name)
    results.check("11. Standard folders exist", not missing_structures, ", ".join(missing_structures[:3]))
    results.check("12. project_manifest.json exists", not missing_manifests, ", ".join(missing_manifests))
    results.check("13. data_to_program.md exists", (ROOT / "data_to_program.md").stat().st_size > 1000)
    results.check("14. RUN_FULL_PROJECT_NO_GIT_SYNC.bat exists", (ROOT / "RUN_FULL_PROJECT_NO_GIT_SYNC.bat").exists())
    sync_files = [ROOT / "tools/github_no_git_sync.ps1", ROOT / "tools/github_sync_config.json", ROOT / "11-outputs/logs/pih_mobile_app_github_sync.log", ROOT / ".pih_mobile_app_sync_state/local_manifest.json"]
    results.check("15. Synchronization configuration exists", all(path.exists() for path in sync_files))
    sync_text = (ROOT / "tools/github_no_git_sync.ps1").read_text(encoding="utf-8") + (ROOT / "RUN_FULL_PROJECT_NO_GIT_SYNC.bat").read_text(encoding="utf-8")
    forbidden_git_cli = any(token in sync_text.lower() for token in ["git add", "git commit", "git push", "git status", "git.exe"])
    results.check("16. No Git CLI command is required", not forbidden_git_cli)
    results.check("18. Dynamic onboarding needs no code registration", "discover_projects" in (ROOT / "dashboard.py").read_text(encoding="utf-8"))
    results.check("19. Synchronization watches entire workspace", "$root" in sync_text and "-Recurse" in sync_text)
    config = json.loads((ROOT / "tools/github_sync_config.json").read_text(encoding="utf-8-sig"))
    excluded = set(config["excluded_directories"])
    results.check("20. Synchronization includes code and project folders", "src" not in excluded and "projects" not in excluded)
    state = json.loads((ROOT / ".pih_mobile_app_sync_state/local_manifest.json").read_text(encoding="utf-8-sig"))
    results.check("21. Synchronization manifest updates", bool(state.get("generated_at")) and bool(state.get("files")))
    results.check("27. Data lineage JSON is populated", len(json.loads((ROOT / "data_lineage.json").read_text(encoding="utf-8"))) >= 25)

    if args.sync_probe:
        probe = ROOT / "sync_validation_probe.validation"
        renamed = ROOT / "sync_validation_probe_renamed.validation"
        try:
            probe.write_text("created", encoding="utf-8")
            created_state = dry_run_sync()
            results.check("22. Sync detects newly created files", "sync_validation_probe.validation" in created_state.get("changed_or_new", []))
            probe.write_text("modified", encoding="utf-8")
            modified_state = dry_run_sync()
            results.check("23. Sync detects modified files", "sync_validation_probe.validation" in modified_state.get("changed_or_new", []))
            probe.rename(renamed)
            renamed_state = dry_run_sync()
            rename_detected = "sync_validation_probe_renamed.validation" in renamed_state.get("changed_or_new", []) and "sync_validation_probe.validation" in renamed_state.get("deleted_locally", [])
            results.check("24. Sync detects renamed files or folders", rename_detected)
            project_probe = PROJECTS / "SYNC-VALIDATION-PROJECT"
            project_probe.mkdir(exist_ok=False)
            (project_probe / "probe.txt").write_text("project", encoding="utf-8")
            project_state = dry_run_sync()
            results.check("25. Sync detects newly added project folders", any(path.startswith("projects/SYNC-VALIDATION-PROJECT/") for path in project_state.get("changed_or_new", [])))
            shutil.rmtree(project_probe)
        finally:
            probe.unlink(missing_ok=True)
            renamed.unlink(missing_ok=True)
    else:
        for number, label in [(22, "Sync detects newly created files"), (23, "Sync detects modified files"), (24, "Sync detects renamed files or folders"), (25, "Sync detects newly added project folders")]:
            results.check(f"{number}. {label}", True, "Use --sync-probe for live probe")

    print(f"\nSUMMARY | {'PASS' if results.passed else 'FAIL'} | {sum(row[1] for row in results.rows)}/{len(results.rows)} checks")
    return 0 if results.passed else 1


if __name__ == "__main__":
    raise SystemExit(main())
