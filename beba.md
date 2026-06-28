# Project Intelligence Hub Guide

## Identity And Discovery

The existing `Dashboard project` selector is driven by folders under `projects/`. `project_id` is the stable identity used for filtering, caching, databases, reports, slides, exports, and lineage. The folder name is the current location and may be renamed. The portfolio option is named **Decision Making dashboard**.

Each project contains `project_manifest.json`:

- `project_id`: stable; do not change during a folder rename.
- `project_display_name`: selector/business label.
- `project_folder_name` and `project_folder_path`: refreshed automatically after a rename.
- `created_at`, `updated_at`, and `status`: lifecycle metadata.

New folders are discovered without code changes. Projects may stay directly under `projects/`, or may be grouped as `projects/<sector>/<project>`. Sector names are taken from the sector folder names. Missing manifests, standard folders, samples, README files, and empty-folder `.gitkeep` files are created without overwriting user files. The architecture is tested with 25 folders and has no fixed project-count limit.

## Project Structure

```text
projects/<current-folder-name>/
  project_manifest.json
  project.json                         # backward-compatible metadata
  1-branding/
  2-contracts/source/
  2-contracts/evidence/
  3-evidence/
  4-notes/
  data/import_templates/
  delay_analysis/steel_delay_tia_templates/
  delay_analysis/methodology/
  bl/
  fixed/
  letters_intelligence/inbox/
  source_excel/
  outputs/
  reports/
  slides/
  exports/
  logs/
```

Sector layout is also supported:

```text
projects/<sector-folder-name>/<project-folder-name>/
  project_manifest.json
  data/import_templates/
  delay_analysis/
  reports/
  slides/
  exports/
```

`1-branding`, `2-contracts/evidence`, `3-evidence`, and `4-notes` contain non-destructive sample templates. Existing files are never replaced.

## Project Context

`src/construction_system/project_context.py` produces one `ProjectContext` from the selected stable `project_id`. It exposes the actual current folder and all project-owned data/output paths.

- Core dashboard CSVs: `data/import_templates`.
- Delay TIA: `delay_analysis/steel_delay_tia_templates` and `bl`.
- Letters: `letters_intelligence`.
- Claims database: `2-contracts/contract_claims.db`.
- Contract sources/evidence: `2-contracts/source` and `2-contracts/evidence`.
- Reports/slides/exports/logs: same-named folders inside the selected project.
- Branding: `1-branding/logo.png`.

There is no fallback to another project. Missing project data produces an empty/setup state. `Decision Making dashboard` aggregates only supported portfolio core data with `project_id` retained and appears as a standalone Phase 1 portfolio command view. Project tabs, Claims Intelligence, Delay TIA, reports, slides, and exports appear only after opening a Phase 2 project workspace.

## Adding Or Renaming A Project

1. Copy `projects/_PROJECT_TEMPLATE` to any folder name under `projects`.
2. Optional: create a sector folder and place the project under `projects/<sector>/`.
3. Refresh Streamlit. The app creates a manifest and standard structure.
4. Fill project-owned CSVs and update display metadata.
5. Keep the generated `project_id` stable.

To rename, rename only the folder. The next discovery refresh updates manifest path fields and the selector folder label while keeping `project_id` unchanged.

## Data Lineage

- Human-readable mapping: `data_to_program.md`.
- Machine-readable mapping: `data_lineage.json`.
- Generator: `tools/generate_data_lineage.py`.

The mapping covers dashboard KPIs, charts, tables, every slide, Claims Intelligence, Delay TIA, Output Studio, reports, slides, exports, evidence, notes, branding, filtering, cache keys, transformations, validation, and missing-data behavior.

## Run

```powershell
.\RUN_APP.bat
```

Local URL: `http://127.0.0.1:18755/`.

## Full Workspace No-Git Sync

Use only this launcher for the full code-and-project repository synchronization:

```powershell
.\RUN_MOBILE_APP_SYNC.bat
```

Optional modes:

```powershell
.\RUN_MOBILE_APP_SYNC.bat Watch 30
.\RUN_MOBILE_APP_SYNC.bat Once 30
.\RUN_MOBILE_APP_SYNC.bat DryRun 30
```

The engine is `tools/github_no_git_sync.ps1`; configuration is `tools/github_sync_config.json`; log is `11-outputs/logs/pih_mobile_app_github_sync.log`; local state is `.pih_mobile_app_sync_state/local_manifest.json`. It scans the whole workspace recursively, includes code and all current/future project folders, does not use Git commands, reads credentials only from `PIH_MOBILE_APP_GITHUB_TOKEN` or `PIH_MOBILE_APP_GH_TOKEN`, and does not delete remote files unless `sync_deletions` is explicitly enabled.

`RUN_LIVE_EXCEL_SYNC.bat` is separate and is not the full-workspace synchronizer.

### GitHub Credential And Refresh Controls

The synchronizer accepts only `PIH_MOBILE_APP_GITHUB_TOKEN` or `PIH_MOBILE_APP_GH_TOKEN` from the process, Windows user, or machine environment. It validates access to `ahmedlabib33-boop/mobile-app` and requires repository push permission before scanning remote content. A Codespaces user secret is available inside Codespaces only and does not authenticate the local Windows app.

Create a new repository-scoped GitHub token with **Contents: Read and write**. Never place it in this repository, `github_sync_config.json`, Streamlit source code, chat, or logs. Set it from a private terminal, then start a new terminal:

```powershell
setx PIH_MOBILE_APP_GITHUB_TOKEN "<new-repository-token>"
setx PIH_MOBILE_APP_SYNC_ADMIN_PIN "<new-private-admin-pin>"
```

The Output Studio contains a protected **Repository synchronization** panel with:

- **Sync**: immediately synchronizes the complete local workspace to the configured repository, then clears Streamlit data caches so local changes are reloaded.
- **Start 30-minute auto sync**: starts the same watcher used by `RUN_MOBILE_APP_SYNC.bat Watch 30`.
- A recent log preview from `11-outputs/logs/pih_mobile_app_github_sync.log`.

`PIH_MOBILE_APP_SYNC_ADMIN_PIN` is optional. When configured, the Sync action requests it; when omitted, the button remains available. `PIH_MOBILE_APP_GITHUB_TOKEN` can be provided either as an operating-system environment variable or as a Streamlit secret and is passed only to the synchronization child process.

VS Code provides the same operations through **Terminal > Run Task** using the `PIH:` tasks in `.vscode/tasks.json`.

Use **PIH: Configure sync credentials securely** first. VS Code requests the replacement token and a new administrator PIN through hidden terminal prompts, validates repository access, and writes them only to the Windows user environment. The administrator PIN must contain at least eight characters.

The connected Streamlit Community Cloud application redeploys from repository branch `main` after GitHub receives the synchronization commit. Deployment timing is controlled by Streamlit; the local **Sync** action cannot force an already-running cloud container to pull repository files directly.

General repository deletion remains disabled. `prune_legacy_project_folders` permits deletion only under the obsolete project-relative folders `branding`, `contracts`, `evidence`, and `notes`; numbered folders and all other repository paths remain protected.

## Validation

```powershell
python -m compileall -q dashboard.py contract_claims_center.py src tools scripts tests
python -m pytest -q tests
python tools/generate_data_lineage.py
python tools/validate_project_isolation.py --sync-probe
python scripts/evaluate_delay_tia.py --label project_isolation
python scripts/evaluate_output_studio.py --label project_isolation
```

## Rules

- Filter and cache by stable `project_id`.
- Resolve paths through `ProjectContext`; never construct a project path from a display name.
- Never use another project's data as fallback.
- Never add project-specific IDs, WBS codes, parties, paths, values, or methodology registration to Python.
- Never overwrite user files while onboarding a project.
- Keep existing business logic and exports; change only their project binding when required.


