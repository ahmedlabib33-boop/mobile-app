from __future__ import annotations

import json
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


def item(page, section, displayed, source_file, columns, transform, function, scope="project", missing="Show zero, empty table, or selected-project setup state", validation="project_id must equal active project_id", source_folder="01-data/import_templates", output="Streamlit and selected-project exports"):
    return {
        "ui_page": page,
        "section_name": section,
        "displayed_field_kpi_chart_or_table": displayed,
        "active_project_behavior": "Filter by stable project_id and resolve the current project folder through project_manifest.json",
        "source_project": "active project_context.project_id" if scope == "project" else "all discovered project_ids",
        "source_folder": f"projects/{{project_folder_name}}/{source_folder}",
        "source_file": source_file,
        "source_sheet_or_table": "CSV rows" if source_file.endswith(".csv") else "Project-owned document/workbook/database",
        "source_columns": columns,
        "transformation_logic": transform,
        "filtering_logic": "selected project_id only" if scope == "project" else "explicit All projects aggregation retaining project_id",
        "cache_key": "project_context.cache_key + source path + file mtime_ns + file size",
        "output_file_where_used": output,
        "scope": scope,
        "missing_data_behavior": missing,
        "validation_rule": validation,
        "responsible_module_function": function,
    }


LINEAGE = [
    item("Global", "Project selector", "Dropdown label and folder label", "project_manifest.json", "project_id, project_display_name, project_folder_name, project_folder_path, status", "Folder discovery; stable project_id; current folder name shown when different", "project_catalog.discover_projects / dashboard.project_filter_options", scope="portfolio", source_folder="", output="All tabs"),
    item("Decision Making dashboard", "Overall Portfolio", "Portfolio KPI cards, sector donut, budget bar, bubble scatter, gauges, timeline, risk heatmap, SPI/CPI quadrant, radar, EV comparison, registry table", "projects.csv, activities.csv, evm.csv, risks.csv, milestones.csv + project_manifest.json", "project_id, sector_name, project_name, contract_value, progress, BAC, PV, EV, AC, SPI, CPI, risk and milestone rows", "Build one project registry row per discovered project; derive portfolio charts from registry without cross-project fallback", "dashboard.build_decision_dashboard_registry / render_decision_making_dashboard", scope="portfolio", source_folder="{sector}/{project}/01-data/import_templates", output="Streamlit Decision Making dashboard"),
    item("Decision Making dashboard", "Phase transition", "Phase 1 portfolio command view and Phase 2 project workspace opener", "project_manifest.json + projects.csv", "project_id, sector_name, project_name, status, progress, SPI, CPI, risks", "Portfolio mode is standalone; select a project_id from preview cards/dropdown, then rerun selected project workspace", "dashboard.render_decision_making_dashboard", scope="portfolio", source_folder="{sector}/{project}", output="Streamlit Decision Making dashboard"),
    item("Decision Making dashboard", "Sector Analysis", "Sector filter, sector health, budget vs spent, progress gauges, benchmark scatter, resource proxy, milestones, EV metrics", "projects.csv, evm.csv, risks.csv, milestones.csv", "Sector, Contract Value, AC, Progress, SPI, CPI, Risks, Milestones", "Filter portfolio registry rows by sector folder name", "dashboard.render_decision_making_dashboard", scope="portfolio", source_folder="{sector}/{project}/01-data/import_templates", output="Streamlit Decision Making dashboard"),
    item("Decision Making dashboard", "Projects Analysis", "Multi-project filter, summary cards, budget comparison, progress/quality/safety trends, SPI/CPI scatter, risk stacked bars, EV comparison matrix", "projects.csv, evm.csv, risks.csv, milestones.csv", "Project, Contract Value, Progress, Quality, Safety, SPI, CPI, Risks, BAC, PV, EV, AC", "User-selected project comparison from registry rows", "dashboard.render_decision_making_dashboard", scope="portfolio", source_folder="{sector}/{project}/01-data/import_templates", output="Streamlit Decision Making dashboard"),
    item("Global", "Header", "Project, contractor, employer, currency, status", "projects.csv + project_manifest.json", "project_name, contractor, client_name, currency, status", "Manifest identity merged with selected project metadata row", "dashboard active_project_record / build_overview_metrics"),
    item("Overview", "Date cards", "Start, finish, duration, elapsed, remaining", "projects.csv", "planned_start, planned_finish", "Parse dates; duration=finish-start; elapsed bounded to 0-100%", "dashboard.build_overview_metrics"),
    item("Overview", "Progress cards", "Overall and planned progress", "projects.csv", "actual_progress_percent, planned_progress_percent", "Numeric normalization; portfolio is contract-value weighted", "dashboard.build_overview_metrics"),
    item("Overview", "Commercial card", "Contract value", "projects.csv", "contract_value, currency", "Numeric normalization and currency formatting", "dashboard.build_overview_metrics"),
    item("Overview", "Activity cards", "Total and critical activities", "activities.csv", "activity_id, is_critical", "Count rows and is_critical=Yes", "dashboard.build_overview_metrics"),
    item("Overview", "Live alerts", "Priority correspondence threads", "letters_intelligence.xlsx", "Thread, Priority, Next Action", "Auto-ingest project inbox then show highest-priority threads", "dashboard.load_letters_workbook", source_folder="07-letters_intelligence"),
    item("WBS", "WBS analysis", "WBS table and construction charts", "wbs.csv", "WBS Code, WBS Name, schedule_%_complete, performance_%_complete, budget_cost, actual_cost", "Select available construction branches; no fixed WBS codes", "dashboard.build_wbs_metrics"),
    item("Activities", "Activity KPIs and tables", "Total, critical, deviated, RFT, variance tables", "activities.csv", "activity_id, activity_name, planned_progress, actual_progress, baseline_finish, forecast_finish, is_critical, wbs_id", "Progress variance and finish slip derived per activity", "dashboard.build_activity_metrics"),
    item("Milestones", "Milestone status", "Milestone counts, dates, status tables", "milestones.csv", "milestone_id, milestone_name, baseline_date, forecast_date, actual_date, status", "Date normalization and forecast variance", "dashboard.build_milestone_metrics"),
    item("Milestones", "Change orders", "Change-order exposure", "change_orders.csv", "change_order_id, status, submitted_amount, approved_amount, time_impact_days", "Count and sum selected-project rows", "dashboard.build_milestone_metrics"),
    item("S-Curve", "Progress curve", "Planned and actual curves", "s_curve.csv", "period, planned_progress, actual_progress", "Sort by period and normalize percentages", "dashboard.build_s_curve_metrics"),
    item("EVM Analysis", "EVM cards and charts", "BAC, AC, EV, PV, SV, CV, EAC, TCPI, SPI", "evm.csv", "BAC, AC, EV, PV, SV, CV", "Sum selected rows; derive EAC, TCPI, SPI", "dashboard.build_evm_metrics / build_earned_value_analysis_data"),
    item("Contracts", "Contract KPIs", "Contract count, certified, paid, outstanding", "contracts.csv + payments.csv", "contract_id, contract_value, certified_amount, paid_amount, payment_status", "Selected-project sums; missing numeric series becomes zero", "dashboard.build_contract_metrics"),
    item("Letters Intelligence", "Letters and links", "Outbound, inbound, linked threads, risk subjects", "letters_intelligence.xlsx + inbox", "Ref No, Date, Subject, Risk Type, Claim Strength, Delay Risk", "Merge workbook with new files in selected project inbox only", "dashboard.load_letters_workbook / letters_auto_ingest.merge_inbox_letters", source_folder="07-letters_intelligence"),
    item("Risks", "Risk register", "Open, high, IFC, RFI risk metrics and tables", "risks.csv + ifc_conflict.csv + rfi_ status.csv", "risk_id, probability, impact, status, Delay Days, RFI No.", "Normalize risk rating and count selected rows", "dashboard.build_risk_metrics"),
    item("Delay Analysis - TIA", "Uploads", "Recognized file inventory and include/exclude controls", "01-13 Delay TIA CSV files", "all source columns + project_id + source_file + source_folder + source_row", "Inspect selected project's Delay TIA folder only", "dashboard.build_steel_delay_template_inventory_df", source_folder="02-delay_analysis/steel_delay_tia_templates"),
    item("Delay Analysis - TIA", "Tables & Conclusion", "Canonical source tables and conclusion", "01-13 Delay TIA CSV files", "metadata, activity, supply, P6, relationships, clauses, IFC, payments, RFI, concurrency", "Build event/activity/supply/concurrency context without cross-project fallback", "dashboard.build_delay_tia_analysis_context", source_folder="02-delay_analysis/steel_delay_tia_templates"),
    item("Delay Analysis - TIA", "MEP Activities", "MEP activities, schedule, civil logic", "MEP Activities.csv + MEP Schedule.csv + MEP Civil Logic.csv", "activity IDs, names, durations, predecessors, successors, dates", "Read selected project schedule folder and preserve civil logic", "dashboard.load_delay_tia_bl_sources", source_folder="03-schedule"),
    item("Delay Analysis - TIA", "AI - TIA", "Readiness, fragnet, causation, concurrency, EOT", "02,03,04,05,06,07,08,09,10,11 CSVs", "event ID, activity ID, float, critical, longest path, overlap, evidence", "Hybrid retrospective TIA rules and selected-project source trace", "steel_delay_tia.run_steel_delay_tia_analysis", source_folder="02-delay_analysis/steel_delay_tia_templates"),
    item("Delay Analysis - TIA", "Question", "Column inventory, management answer and tables", "selected project Delay TIA CSVs", "all recognized columns", "Question analysis reads scoped directory and calculates from available records", "dashboard.load_delay_tia_question_frames", source_folder="02-delay_analysis/steel_delay_tia_templates"),
    item("Delay Analysis - TIA", "Download Reports", "DOCX/HTML/CSV/Primavera outputs", "canonical TIA datasets", "project_id, source_file, source_row, event_id, activity_id", "Generate into selected project deliverables path", "TIADirectorPackGenerator", source_folder="02-delay_analysis", output="projects/{project_folder_name}/10-deliverables"),
    item("Contract & Claims Intelligence Center", "Contract clauses", "Clause search, entitlement and event matching", "06- contract_library.csv", "Clause / Topic, leverage, notice, money impact, schedule impact, evidence", "Load active project clause library only", "contract_matcher.set_clause_library_path / render_contract_clause_matching_engine", source_folder="02-delay_analysis/steel_delay_tia_templates"),
    item("Contract & Claims Intelligence Center", "Claims intelligence", "Documents, clauses, evidence mappings, drafts, rebuttals", "contract_claims.db + 05-contracts/source + 06-evidence", "project_id, document path/hash, clause, evidence, mapping score, draft", "Physically separate SQLite database per project; project trace columns added to UI frames", "contract_claims_center.* / dashboard claims tab", source_folder="05-contracts", output="projects/{project_folder_name}/11-outputs"),
    item("Output Studio", "Executive dashboard", "All management KPIs, alerts, discipline and cost views", "projects.csv, activities.csv, evm.csv, contracts.csv, payments.csv, risks.csv, milestones.csv, s_curve.csv", "overview/EVM/contract/risk/milestone/activity fields", "Use already scoped dashboard metric dictionaries", "dashboard.build_the_big_decision_dashboard_html", output="selected-project-prefixed HTML download"),
    item("Output Studio", "Linked executive dashboard", "HTML/PPTX/A3/EVM outputs", "same active-project metric dictionaries", "all linked dashboard KPI fields", "No fallback to another project; project-prefixed filenames", "dashboard.build_linked_executive_dashboard_html and export helpers", output="selected-project-prefixed HTML/PPTX downloads"),
    item("Output Studio", "Detailed Progress report", "XLSX, HTML, DOCX, Power BI style", "all selected-project operational CSVs and letters", "report control, project, activities, schedule, cost, risk, correspondence", "Build report package from active project metrics and tables", "dashboard.build_detailed_progress_report_package", output="selected-project-prefixed XLSX/HTML/DOCX downloads"),
    item("Branding", "Logo and report identity", "Header logo and identity", "logo.png + project identity templates", "logo_file, project_display_name, contractor, client", "Use 08-branding first for selected project identity", "dashboard project_logo_path", source_folder="08-branding"),
    item("Evidence", "Evidence registers", "Claim/TIA supporting evidence", "evidence register templates and uploaded evidence", "evidence_id, source_file, event_id, activity_id, verified", "Selected project only", "project_context.evidence_path / contract_claims_center", source_folder="06-evidence"),
    item("Notes", "Project notes", "Meeting, engineering, and claims notes", "*.md", "date, references, decisions, actions", "Selected project only; notes are not imported into another project", "project_context.notes_path", source_folder="09-notes"),
    item("Portfolio", "Decision Making dashboard", "Portfolio overview, core KPIs, letters", "all discovered project core CSVs", "project_id retained on every row", "Aggregate only explicitly supported portfolio datasets; claims/TIA project workflows are blocked", "dashboard.load_core_csv / load_letters_workbook", scope="portfolio", validation="every aggregated row has project_id"),
]


def main() -> None:
    json_path = ROOT / "data_lineage.json"
    json_path.write_text(json.dumps(LINEAGE, indent=2, ensure_ascii=True) + "\n", encoding="utf-8")
    headers = list(LINEAGE[0])
    lines = [
        "# Data To Program Lineage",
        "",
        "This document maps visible application outputs to project-owned sources. `project_id` is the stable identity; `project_folder_name` is the current folder location. Folder renames update the manifest path without changing project identity.",
        "",
        "## Binding Rules",
        "",
        "- The existing `Dashboard project` selector writes `active_project_id` to Streamlit session state.",
        "- `ProjectContext` resolves that stable ID to exactly one current folder through `project_manifest.json`.",
        "- Project-specific Claims Intelligence, Delay TIA, reports, slides, and exports are blocked in `Decision Making dashboard` mode.",
        "- Missing files produce empty/setup states. No loader falls back to another project.",
        "- Core portfolio aggregation is explicit and retains `project_id` on every row.",
        "- The portfolio option is named `Decision Making dashboard`; it supports root projects and `projects/{sector}/{project}` folders.",
        "",
        "## Complete Mapping",
        "",
        "| " + " | ".join(headers) + " |",
        "|" + "|".join(["---"] * len(headers)) + "|",
    ]
    for row in LINEAGE:
        values = [str(row[h]).replace("|", "/").replace("\n", " ") for h in headers]
        lines.append("| " + " | ".join(values) + " |")
    lines += ["", "## Machine-Readable Copy", "", "The same mapping is stored in `data_lineage.json` for validation and future automation."]
    (ROOT / "data_to_program.md").write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"Generated {len(LINEAGE)} lineage records")


if __name__ == "__main__":
    main()
