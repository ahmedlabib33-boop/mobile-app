# Power BI Report Layout Specification

## Page 1 - Decision Making Dashboard
- KPI cards: Total Projects, Total Contract Value, Files Loaded, Contract Clauses, High Risk Clauses, Strong Claim Items.
- Donut: Projects by sector_name.
- Bar: Clauses by risk_level.
- Matrix: project_name, sector_name, source file counts, clause counts, evidence counts.

## Page 2 - Project Deep Dive
- Slicer: project_name.
- Cards: contract documents, contract clauses, evidence documents, draft claims.
- Table: source files with relative_path, suffix, modified_utc.

## Page 3 - Contract & Claims Intelligence
- Slicer: project_name, claim_type, risk_level.
- Cards: High Risk Clauses, Strong Claim Items, Schedule Impact Clauses, Cost Impact Clauses.
- Table: clause_number, clause_title, section_name, claim_type, risk_level, claim_strength, file_name.

## Page 4 - Source Data Register
- Matrix of all files by project and folder.
- Use this page for audit and data lineage.
