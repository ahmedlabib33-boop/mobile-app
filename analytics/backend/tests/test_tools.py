from app.tools import check_launch_readiness_impl, extract_tasks_from_brief_impl, generate_owner_checklist_impl


def test_extract_tasks_includes_delay_task():
    result = extract_tasks_from_brief_impl(
        "Launch an analytics release for Primavera delay analysis and critical path review.",
        "2026-07-01",
        "Need XER and PDF evidence normalized.",
    )
    tasks = [item["task"] for item in result["tasks"]]
    assert any("delay analysis" in task.lower() for task in tasks)


def test_readiness_scores_assets():
    result = check_launch_readiness_impl(
        "A detailed launch brief with risks, delay constraints, schedule impact, and release scope. " * 3,
        "Engineering directors and planning team",
        "2026-07-30",
        "XER baseline, update XML, claim PDF, risk register",
    )
    assert result["score"] >= 60
    assert result["asset_count"] == 4


def test_owner_checklist_has_core_roles():
    result = generate_owner_checklist_impl("XER, PDF", "limited schedule access", "TIA required")
    owners = {item["owner"] for item in result["checklist"]}
    assert "Planning Engineer" in owners
    assert "Delay Analyst" in owners
