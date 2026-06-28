# Projects Pipeline Report

## Purpose

This report explains how Project Intelligence Hub moves information from project folders into the Decision Making Dashboard and then into each selected project workspace.

## Current Project Source

Root folder:

```text
C:\Users\pc\OneDrive\Documents\Project Intelligence Hub\projects
```

Currently detected live project folders:

| Project folder | Sector | Role |
|---|---|---|
| `Buildings/The BIG - PH01` | `Buildings` | Live project folder |
| `Tunnels/Pr02` | `Tunnels` | Live project folder |
| `Tunnels/Pr03` | `Tunnels` | Live project folder |
| `Bridges` | Sector folder, currently empty | Holds future bridge projects |
| `_PROJECT_TEMPLATE` | None | Template folder, ignored by discovery |

## Pipeline Overview

```text
Project folder
  -> project_manifest.json
  -> project-owned CSV / Excel / evidence folders
  -> project discovery registry
  -> ProjectContext
  -> selected-project data loaders
  -> project calculations
  -> sector aggregation
  -> Decision Making Dashboard
  -> project deep-dive workspace
  -> project-owned reports / slides / exports
```

## Stage 1 - Project Folder Discovery

Implemented in:

```text
src/construction_system/project_catalog.py
```

The app scans the `projects` folder dynamically. It supports two layouts:

```text
projects/<project-folder>
projects/<sector-folder>/<project-folder>
```

Rules:

- Folders starting with `_` or `.` are ignored.
- `_PROJECT_TEMPLATE` is ignored as a live project.
- Root-level project folders remain supported.
- Sector folders are supported when projects are placed inside them.
- If a project has no manifest, the app creates one safely.
- Missing standard folders are created without overwriting user files.

## Stage 2 - Project Identity

Every project is controlled by:

```text
project_manifest.json
```

Important fields:

| Field | Purpose |
|---|---|
| `project_id` | Stable identity used for filtering, cache, reports, claims, Delay TIA, and exports |
| `project_display_name` | Human-readable display name |
| `project_folder_name` | Current folder name |
| `project_folder_path` | Current folder path |
| `sector_id` / `sector_name` | Sector grouping when project is inside a sector folder |
| `status` | Setup / active lifecycle value |

Folder names can change. `project_id` should stay stable.

## Stage 3 - Standard Project Structure

Each project is expected to have:

```text
01-data/import_templates/
02-delay_analysis/steel_delay_tia_templates/
02-delay_analysis/methodology/
03-schedule/
04-source_excel/
05-contracts/source/
05-contracts/clauses/
06-evidence/
07-letters_intelligence/inbox/From Contractor/
07-letters_intelligence/inbox/From Consultant/
08-branding/
09-notes/
10-deliverables/
11-outputs/
12-logs/
```

The app creates missing folders non-destructively.

## Stage 4 - Project Context

Implemented in:

```text
src/construction_system/project_context.py
```

`ProjectContext` converts the selected project into exact paths:

| Context path | Points to |
|---|---|
| `data_path` | selected project `01-data/` |
| `branding_path` | selected project `08-branding/` |
| `contracts_path` | selected project `05-contracts/` |
| `contracts_evidence_path` | selected project `06-evidence/` |
| `evidence_path` | selected project `06-evidence/` |
| `notes_path` | selected project `09-notes/` |
| `reports_path` | selected project `10-deliverables/` |
| `slides_path` | selected project `10-deliverables/` |
| `exports_path` | selected project `11-outputs/` |
| `logs_path` | selected project `12-logs/` |

This is the isolation layer. Project-specific modules must use the active `ProjectContext`.

## Stage 5 - Data Loading

Main loader:

```text
dashboard.py -> load_core_csv()
```

Behavior:

- If a project is selected, it loads only that project files.
- If Decision Making Dashboard mode is selected, it aggregates supported project files from all discovered projects.
- Every aggregated row keeps or receives `project_id`.
- If a file is missing, it returns an empty dataframe instead of loading another project.

Project-scoped path resolver:

```text
dashboard.py -> project_scoped_file()
```

It maps shared placeholder paths into the selected project folders:

| Family | Selected project folder |
|---|---|
| Core data | `01-data/import_templates/` |
| Delay TIA | `02-delay_analysis/steel_delay_tia_templates/` |
| BL / fixed logic | `03-schedule/` |

## Stage 6 - Decision Making Dashboard Aggregation

Implemented in:

```text
dashboard.py -> build_decision_dashboard_registry()
dashboard.py -> render_decision_making_dashboard()
```

The dashboard is the portfolio aggregation layer:

```text
Project Data -> Project Calculations -> Sector Aggregation -> Portfolio Dashboard
```

It reads from project-owned files such as:

| Source | Used for |
|---|---|
| `projects.csv` | project name, dates, contract value, planned/actual progress |
| `activities.csv` | activity progress fallback |
| `evm.csv` | BAC, PV, EV, AC, SPI, CPI, EAC, VAC |
| `payments.csv` | paid amounts |
| `contracts.csv` | actual cost / certified fallback |
| `delay_events.csv` | delay days and EOT/claim signals |
| `risks.csv` | risk count and risk score |
| `milestones.csv` | milestone count/status indicators |

No manual dashboard dataset is created.

## Stage 7 - EVM Calculation Logic

The dashboard uses available fields first:

```text
PV, EV, AC, BAC, SPI, CPI
```

When PV/EV are missing but BAC and progress percentages exist:

```text
PV = BAC x planned progress %
EV = BAC x actual progress %
```

Then:

```text
SPI = EV / PV
CPI = EV / AC
SV = EV - PV
CV = EV - AC
EAC = BAC / CPI
ETC = EAC - AC
VAC = BAC - EAC
Progress Variance = Actual Progress % - Planned Progress %
```

All ratios use safe division. Missing denominators return unavailable values instead of crashing.

## Stage 8 - Sector Aggregation

Sector source:

```text
project_catalog.read_project_metadata()
```

Sector is determined as:

1. `sector_name` in manifest, if available.
2. Parent sector folder name, if project is inside `projects/<sector>/<project>`.
3. `Unassigned`, if root-level project has no sector.

Sector dashboard values are aggregated from projects belonging to that sector.

## Stage 9 - Project Deep Dive Transition

The Decision Making Dashboard transition dropdown:

```text
Transition to Project Deep Dive
```

When a project is selected:

```text
st.session_state["active_project_id"] = selected_project_id
st.rerun()
```

This uses the existing app navigation and project workspace logic. It does not create a new project system.

## Stage 10 - Project-Specific Workspaces

When a project is selected:

- Dashboard cards load selected-project core CSVs.
- Delay Analysis / TIA uses selected project delay files.
- Claims Intelligence uses selected project contract database and folders.
- Output Studio saves to selected project output folders.
- Reports save to selected project `10-deliverables/`.
- Slides save to selected project `10-deliverables/`.
- Exports save to selected project `11-outputs/`.

Portfolio mode is blocked for project-specific workflows that require one project.

## Stage 11 - Output Pipeline

Project outputs are saved under the selected project:

```text
projects/<sector>/<project>/10-deliverables/
projects/<sector>/<project>/11-outputs/
projects/<sector>/<project>/12-logs/
```

This prevents cross-project output mixing.

## Stage 12 - Sync Pipeline

Full workspace sync launcher:

```text
RUN_FULL_PROJECT_NO_GIT_SYNC.bat
```

Main sync engine:

```text
tools/github_no_git_sync.ps1
```

Behavior:

- Watches the full workspace.
- Includes code folders and project folders.
- Does not use Git commands.
- Reads credentials only from `PIH_MOBILE_APP_GITHUB_TOKEN` or `PIH_MOBILE_APP_GH_TOKEN`.
- Writes logs to `11-outputs/logs/pih_mobile_app_github_sync.log`.
- Tracks local state in `.pih_mobile_app_sync_state/local_manifest.json`.

## Data Isolation Rules

| Rule | Status |
|---|---|
| Each project has its own folder | Active |
| Each project has stable `project_id` | Active |
| Project-specific loaders use selected project paths | Active |
| Portfolio aggregation retains `project_id` | Active |
| Missing files do not fallback to another project | Active |
| Claims database is project-owned | Active |
| Delay TIA files are project-owned | Active |
| Reports/slides/exports are project-owned | Active |

## How To Add A New Project

Recommended workflow:

1. Create a sector folder if needed:

```text
projects/<Sector Name>/
```

2. Add the project folder:

```text
projects/<Sector Name>/<Project Name>/
```

3. Copy the standard template if desired:

```text
projects/_PROJECT_TEMPLATE
```

4. Put project files inside the project-owned folders.
5. Reload the Streamlit app.
6. The project appears automatically in the dropdown and Decision Making Dashboard.

## Key Risk Points

| Risk | Control |
|---|---|
| Missing `project_id` in CSV | Loader inserts or normalizes project_id during aggregation |
| Folder rename | Manifest keeps stable project_id and refreshes folder path |
| Missing data file | Empty state instead of fallback |
| Project values inconsistent with dashboard | Dashboard reads the same project data sources |
| New project not appearing | Ensure it is not `_` prefixed and is under `projects/` |

## Validation Commands

```powershell
python -m pytest -q tests -p no:cacheprovider
python tools\validate_project_isolation.py
```

Expected validation:

```text
tests pass
project isolation validator passes 27/27
```


