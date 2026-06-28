# Projects Folder

Use this folder as the source of truth for project workspaces.

## Recommended Layout

```text
projects/
  <Sector Name>/
    <Project Name>/
      01-data/
      02-delay_analysis/
      03-schedule/
      04-source_excel/
      05-contracts/
      06-evidence/
      07-letters_intelligence/
      08-branding/
      09-notes/
      10-deliverables/
      11-outputs/
      12-logs/
```

Combined folders are intentionally single folders:

- `03-schedule` holds both baseline and fixed schedule reference files.
- `06-evidence` holds both contract evidence and general evidence.
- `10-deliverables` holds both report and slide deliverables.
- `11-outputs` holds both working outputs and export/download packages.

## Current Layout

```text
projects/
  Buildings/
    The BIG - PH01/
  Tunnels/
    Pr02/
    Pr03/
  Bridges/
  _PROJECT_TEMPLATE/
```

`_PROJECT_TEMPLATE` is kept at the root and ignored as a live project.

## Adding A Project

Create a folder under the correct sector:

```text
projects/<Sector Name>/<Project Name>/
```

The app will create missing standard numbered folders and a `project_manifest.json` without overwriting existing files.

## Renaming

You can rename a sector folder or project folder. Existing `project_manifest.json` files preserve the stable `project_id`; the displayed folder name updates from the actual folder path.
