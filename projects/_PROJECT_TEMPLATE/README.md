# New Project Template

Copy this folder into `projects` using any folder name. On refresh, the app generates `project_manifest.json`; keep its `project_id` stable even if the folder is renamed later.

- `data/import_templates`: dashboard CSVs.
- `delay_analysis/steel_delay_tia_templates`: Delay TIA CSVs.
- `delay_analysis/methodology`: project methodology documents.
- `bl`: baseline, critical path, float, longest path, and MEP files.
- `letters_intelligence/inbox/From Contractor`: outgoing letters.
- `letters_intelligence/inbox/From Consultant`: incoming letters.
- `1-branding`: logo, identity, palette, and report branding.
- `2-contracts/source` and `2-contracts/evidence`: contract sources and claim evidence.
- `3-evidence`: general evidence and photo/document registers.
- `4-notes`: meeting, engineering, and claims notes.
- `fixed` and `source_excel`: controlled inputs.
- `outputs`, `reports`, `slides`, `exports`, and `logs`: project-owned outputs.

The app detects the copied folder automatically. Folders beginning with `_` are ignored.
