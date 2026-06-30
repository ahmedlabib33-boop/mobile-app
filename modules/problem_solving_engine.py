from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd
import streamlit as st

from services.iterative_solver import run_iterative_solver
from services.legal_research import DISCLAIMER
from services.ollama_client import OllamaClient
from services.problem_solving_rag import (
    append_answer,
    append_history,
    find_similar_questions,
    init_history_db,
    load_answer_register,
    load_question_bank,
    search_project_evidence,
)


def _yes(value: Any) -> bool:
    return str(value or "").strip().casefold() in {"yes", "true", "1", "y", "required"}


def _css() -> None:
    st.markdown(
        """
        <style>
        .pse-hero{background:linear-gradient(135deg,#081f36,#0f8492);border:1px solid rgba(255,255,255,.16);border-radius:20px;padding:24px 26px;color:#fff;box-shadow:0 18px 45px rgba(11,42,74,.22);margin:10px 0 18px}
        .pse-title{font-size:32px;font-weight:850;letter-spacing:0;margin-bottom:4px}
        .pse-subtitle{font-size:15px;color:rgba(255,255,255,.78);margin-bottom:14px}
        .pse-badges{display:flex;flex-wrap:wrap;gap:8px}
        .pse-badge{display:inline-flex;align-items:center;border:1px solid rgba(255,255,255,.25);background:rgba(255,255,255,.12);border-radius:999px;padding:6px 10px;font-size:12px;font-weight:700}
        .pse-grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:12px;margin:12px 0}
        .pse-card{background:#fff;border:1px solid #D9E4EF;border-radius:16px;padding:15px;box-shadow:0 10px 26px rgba(11,42,74,.08);margin-bottom:12px}
        .pse-card h4{margin:0 0 8px;font-size:15px;color:#0B2A4A}
        .pse-muted{color:#64748b;font-size:12px}
        .pse-chip{display:inline-block;border-radius:999px;padding:4px 9px;background:#F4F7FA;border:1px solid #D9E4EF;font-size:11px;font-weight:800;color:#16496F;margin:2px}
        .pse-chip.gold{background:#fff7df;color:#7b5b05;border-color:#efd78d}
        .pse-chip.red{background:#fff1f0;color:#B94642;border-color:#f0b9b6}
        .pse-chip.green{background:#e9fbf5;color:#0f766e;border-color:#a8ead8}
        .pse-answer{background:#f8fbfd;border:1px solid #D9E4EF;border-radius:16px;padding:16px;line-height:1.55}
        .pse-meter{height:12px;border-radius:999px;background:#e5edf5;overflow:hidden}
        .pse-meter span{display:block;height:100%;background:linear-gradient(90deg,#B94642,#D1A329,#0F8492)}
        .pse-evidence{border-left:4px solid #0F8492;padding:10px 12px;background:#fff;border-radius:10px;border-top:1px solid #D9E4EF;border-right:1px solid #D9E4EF;border-bottom:1px solid #D9E4EF;margin-bottom:8px}
        .pse-tabs div[data-testid="stTabs"] button{font-weight:800}
        @media(max-width:900px){.pse-grid{grid-template-columns:1fr 1fr}.pse-title{font-size:25px}}
        @media(max-width:520px){.pse-grid{grid-template-columns:1fr}.pse-hero{padding:18px}.pse-title{font-size:22px}.pse-badge{font-size:11px}.stButton button,.stDownloadButton button{min-height:42px!important}}
        </style>
        """,
        unsafe_allow_html=True,
    )


def _project_options(projects_df: pd.DataFrame) -> list[dict[str, str]]:
    options: list[dict[str, str]] = []
    if projects_df is None or projects_df.empty:
        return options
    for _, row in projects_df.iterrows():
        project_id = str(row.get("project_id", "")).strip()
        name = str(row.get("project_name", "") or project_id).strip()
        folder = str(row.get("project_folder_name", "") or project_id or name).strip()
        sector = str(row.get("sector_name", "") or "").strip()
        if not project_id and not folder:
            continue
        label = f"{sector} / {name}" if sector and sector.casefold() != "unassigned" else name
        if project_id and project_id not in label:
            label = f"{label} ({project_id})"
        options.append({"label": label, "project_id": project_id, "project_name": name, "folder": folder})
    return options


def _project_root(projects_dir: Path, selected: dict[str, str], portfolio: bool) -> Path:
    if portfolio:
        return projects_dir
    folder = selected.get("folder") or selected.get("project_id") or selected.get("project_name") or ""
    return projects_dir / folder


def _filter_questions(bank: pd.DataFrame, department: str, source_layer: str, status: str, project_required: str, search: str) -> pd.DataFrame:
    frame = bank.copy()
    if department != "All":
        frame = frame[frame["department"].astype(str).eq(department)]
    if source_layer != "All":
        frame = frame[frame["source_layer"].astype(str).eq(source_layer)]
    if status != "All":
        frame = frame[frame["answer_status"].astype(str).eq(status)]
    if project_required != "All":
        want = project_required == "Project data required"
        frame = frame[frame["project_data_required"].apply(_yes).eq(want)]
    if search.strip():
        q = search.strip().casefold()
        mask = frame.apply(lambda row: q in " ".join(str(v) for v in row.values).casefold(), axis=1)
        frame = frame[mask]
    return frame


def _render_question_card(row: pd.Series, idx: int) -> None:
    st.markdown(
        f"""
        <div class='pse-card'>
          <h4>{row.get('question_id','')}</h4>
          <div style='font-weight:750;color:#1F2937;margin-bottom:8px'>{row.get('question_text','')}</div>
          <span class='pse-chip'>{row.get('department','')}</span>
          <span class='pse-chip'>{row.get('section_or_level','')}</span>
          <span class='pse-chip gold'>{row.get('answer_status','')}</span>
          <span class='pse-chip {'green' if _yes(row.get('project_data_required')) else ''}'>Project data: {row.get('project_data_required','NO')}</span>
          <span class='pse-chip'>Ollama: {row.get('local_ollama_enabled','')}</span>
          <span class='pse-chip'>FIDIC: {row.get('fidic_check_enabled','')}</span>
          <span class='pse-chip'>Egypt Law: {row.get('egypt_law_check_enabled','')}</span>
        </div>
        """,
        unsafe_allow_html=True,
    )
    if st.button("Open solver", key=f"pse_open_{row.get('question_id','')}_{idx}", width="stretch"):
        st.session_state["pse_selected_question"] = row.to_dict()
        st.session_state["pse_custom_question"] = str(row.get("question_text", ""))
        st.rerun()


def _solution_downloads(solution: dict[str, Any]) -> None:
    payload = json.dumps(solution, ensure_ascii=False, indent=2, default=str)
    st.download_button("Download JSON", payload, file_name=f"{solution.get('answer_id','pse_answer')}.json", mime="application/json", width="stretch")
    markdown = f"# {solution.get('question_text','Problem Solving Answer')}\n\n{solution.get('generated_answer','')}"
    st.download_button("Download Markdown", markdown, file_name=f"{solution.get('answer_id','pse_answer')}.md", mime="text/markdown", width="stretch")
    evidence_df = pd.DataFrame(solution.get("evidence", []))
    if not evidence_df.empty:
        st.download_button("Download Evidence CSV", evidence_df.to_csv(index=False), file_name=f"{solution.get('answer_id','pse_evidence')}_evidence.csv", mime="text/csv", width="stretch")


def render_problem_solving_engine(*, app_dir: Path, projects_dir: Path, projects_df: pd.DataFrame, auth_user: dict[str, Any] | None = None) -> None:
    _css()
    question_bank_path = app_dir / "data" / "problem_solving_engine_question_bank.csv"
    answer_register_path = app_dir / "data" / "problem_solving_answer_register.csv"
    history_db_path = app_dir / "storage" / "problem_solving_history.sqlite"
    init_history_db(history_db_path)

    question_bank = load_question_bank(question_bank_path)
    answer_register = load_answer_register(answer_register_path)
    projects = _project_options(projects_df)

    st.markdown(
        """
        <div class='pse-hero'>
          <div class='pse-title'>Problem Solving Engine</div>
          <div class='pse-subtitle'>AI Evidence-Based Answers for Project Decisions</div>
          <div class='pse-badges'>
            <span class='pse-badge'>Project Data</span><span class='pse-badge'>Question Bank</span>
            <span class='pse-badge'>Local Ollama</span><span class='pse-badge'>Legal Search</span>
            <span class='pse-badge'>FIDIC</span><span class='pse-badge'>Egyptian Law</span><span class='pse-badge'>Confidence</span>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    control_cols = st.columns([0.24, 0.18, 0.18, 0.16, 0.12, 0.12])
    selected_project = projects[0] if projects else {"label": "No project configured", "project_id": "", "project_name": "", "folder": ""}
    with control_cols[0]:
        project_label = st.selectbox("Select project", [p["label"] for p in projects] or ["No project configured"], key="pse_project_selector")
        selected_project = next((p for p in projects if p["label"] == project_label), selected_project)
    with control_cols[1]:
        department = st.selectbox("Department", ["All"] + sorted(question_bank["department"].dropna().astype(str).unique().tolist()), key="pse_department")
    with control_cols[2]:
        source_layer = st.selectbox("Source layer", ["All"] + sorted(question_bank["source_layer"].dropna().astype(str).unique().tolist()), key="pse_source_layer")
    with control_cols[3]:
        answer_status = st.selectbox("Answer status", ["All"] + sorted(question_bank["answer_status"].dropna().astype(str).unique().tolist()), key="pse_answer_status")
    with control_cols[4]:
        project_required = st.selectbox("Data", ["All", "Project data required", "No project data required"], key="pse_project_required")
    with control_cols[5]:
        max_iterations = st.number_input("Iterations", min_value=1, max_value=5, value=3, step=1, key="pse_max_iterations")

    toggle_cols = st.columns([0.18, 0.18, 0.18, 0.18, 0.28])
    with toggle_cols[0]:
        local_ollama = st.toggle("Local Ollama", value=True, key="pse_local_ollama")
    with toggle_cols[1]:
        web_research = st.toggle("Web research", value=False, key="pse_web_research")
    with toggle_cols[2]:
        fidic_check = st.toggle("FIDIC check", value=False, key="pse_fidic_check")
    with toggle_cols[3]:
        egypt_law_check = st.toggle("Egypt law check", value=False, key="pse_egypt_law_check")
    with toggle_cols[4]:
        portfolio = st.toggle("Portfolio-wide search", value=False, help="Off keeps evidence search inside the selected project only.", key="pse_portfolio")

    project_root = _project_root(projects_dir, selected_project, portfolio)
    ollama_ok, ollama_message = OllamaClient().health()
    if local_ollama and not ollama_ok:
        st.warning(ollama_message)

    filtered = _filter_questions(question_bank, department, source_layer, answer_status, project_required, st.session_state.get("pse_bank_search", ""))
    st.markdown(
        f"""
        <div class='pse-grid'>
          <div class='pse-card'><h4>Question Bank</h4><div class='pse-title' style='font-size:24px;color:#0B2A4A'>{len(question_bank):,}</div><div class='pse-muted'>Master questions loaded</div></div>
          <div class='pse-card'><h4>Filtered Questions</h4><div class='pse-title' style='font-size:24px;color:#0F8492'>{len(filtered):,}</div><div class='pse-muted'>Ready for solving</div></div>
          <div class='pse-card'><h4>Answer Register</h4><div class='pse-title' style='font-size:24px;color:#D1A329'>{len(answer_register):,}</div><div class='pse-muted'>Approved or saved answers</div></div>
          <div class='pse-card'><h4>Evidence Scope</h4><div style='font-weight:850;color:#16496F'>{'Portfolio' if portfolio else selected_project.get('project_name','Selected Project')}</div><div class='pse-muted'>{project_root}</div></div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    tabs = st.tabs([
        "Engine Home", "Question Bank Explorer", "Ask a New Question", "Project Evidence Search",
        "Legal / FIDIC Research", "AI Answer Studio", "Validation & Confidence",
        "Iterative Improvement Log", "Export Center",
    ])

    with tabs[0]:
        st.markdown("<div class='pse-card'><h4>Operating Rule</h4>Project evidence stays local to the selected project unless Portfolio-wide search is explicitly enabled. Public legal research sends search terms only, never private project files.</div>", unsafe_allow_html=True)
        st.info("Use the Question Bank Explorer for predefined questions, or Ask a New Question for issues not yet in the bank.")
        st.caption(DISCLAIMER)

    with tabs[1]:
        st.text_input("Search question bank", key="pse_bank_search", placeholder="Search by keyword, clause, issue, department, status...")
        filtered = _filter_questions(question_bank, department, source_layer, answer_status, project_required, st.session_state.get("pse_bank_search", ""))
        for idx, row in filtered.head(40).iterrows():
            _render_question_card(row, int(idx))
        if len(filtered) > 40:
            st.caption(f"Showing first 40 of {len(filtered):,} filtered questions. Narrow the search to show more specific cards.")

    with tabs[2]:
        selected_question = st.session_state.get("pse_selected_question", {})
        default_question = st.session_state.get("pse_custom_question", selected_question.get("question_text", ""))
        question = st.text_area("Question", value=default_question, height=120, key="pse_question_input")
        similar = find_similar_questions(question_bank, question, limit=8) if question.strip() else pd.DataFrame()
        if not similar.empty:
            st.markdown("#### Similar questions")
            st.dataframe(similar[["question_id", "department", "section_or_level", "question_text"]], width="stretch", hide_index=True)
        if st.button("Solve question", type="primary", width="stretch", disabled=not question.strip()):
            row = selected_question if selected_question and selected_question.get("question_text") == default_question else {}
            with st.spinner("Searching evidence, checking references, generating answer, and validating confidence..."):
                solution = run_iterative_solver(
                    question=question,
                    question_row=row,
                    question_bank=question_bank,
                    project_root=project_root,
                    project_id=selected_project.get("project_id", ""),
                    project_name=selected_project.get("project_name", ""),
                    portfolio=portfolio,
                    local_ollama=local_ollama,
                    web_research=web_research,
                    fidic_check=fidic_check,
                    egypt_law_check=egypt_law_check,
                    max_iterations=int(max_iterations),
                )
            st.session_state["pse_last_solution"] = solution
            st.success(f"Answer generated with {solution['confidence']} confidence ({solution['score']}/100).")
            st.rerun()

    solution = st.session_state.get("pse_last_solution", {})

    with tabs[3]:
        manual_query = st.text_input("Evidence search terms", value=solution.get("question_text", ""), key="pse_manual_evidence_query")
        if st.button("Search selected project evidence", width="stretch", disabled=not manual_query.strip()):
            st.session_state["pse_manual_evidence"] = search_project_evidence(project_root, manual_query, max_results=30, portfolio=portfolio)
        evidence = st.session_state.get("pse_manual_evidence", solution.get("evidence", []))
        if evidence:
            for item in evidence[:20]:
                st.markdown(f"<div class='pse-evidence'><b>{Path(str(item.get('source_file',''))).name}</b> / {item.get('sheet_name','')} row {item.get('row_number','')}<br><span class='pse-muted'>Matches: {item.get('matched_fields','')} | Confidence: {item.get('confidence','')}</span><br>{item.get('excerpt','')}</div>", unsafe_allow_html=True)
        else:
            st.info("No evidence searched yet, or no matching records found.")

    with tabs[4]:
        st.warning("Public search uses only search terms. Do not paste confidential project facts into the legal search query.")
        if solution.get("legal_refs"):
            for ref in solution["legal_refs"]:
                st.markdown(f"<div class='pse-card'><h4>{ref.get('title','')}</h4><div>{ref.get('snippet','')}</div><div class='pse-muted'>{ref.get('url','')} | Retrieved {ref.get('retrieved_at','')}</div></div>", unsafe_allow_html=True)
        else:
            st.info("Enable Web research with FIDIC or Egyptian law check, then solve a question to retrieve public references.")
        st.caption(DISCLAIMER)

    with tabs[5]:
        if solution:
            st.markdown(f"<div class='pse-answer'>{solution.get('generated_answer','').replace(chr(10), '<br>')}</div>", unsafe_allow_html=True)
            st.text_area("Approved answer / reviewer edit", value=solution.get("generated_answer", ""), height=260, key="pse_approved_answer")
            cols = st.columns([0.25, 0.25, 0.25, 0.25])
            owner = cols[0].text_input("Owner", key="pse_owner")
            deadline = cols[1].date_input("Deadline", value=None, key="pse_deadline")
            status = cols[2].selectbox("Status", ["Draft", "Approved", "Rejected", "Needs Evidence"], key="pse_status")
            notes = cols[3].text_input("Reviewer notes", key="pse_notes")
            if st.button("Save to Answer Register", type="primary", width="stretch"):
                now = datetime.utcnow().isoformat(timespec="seconds") + "Z"
                record = {
                    "answer_id": solution.get("answer_id"),
                    "question_id": solution.get("question_id"),
                    "project_id": solution.get("project_id"),
                    "project_name": solution.get("project_name"),
                    "question_text": solution.get("question_text"),
                    "generated_answer": solution.get("generated_answer"),
                    "approved_answer": st.session_state.get("pse_approved_answer", ""),
                    "evidence_refs": json.dumps(solution.get("evidence", []), ensure_ascii=False),
                    "legal_refs": json.dumps(solution.get("legal_refs", []), ensure_ascii=False),
                    "confidence": solution.get("confidence"),
                    "status": status,
                    "owner": owner,
                    "deadline": str(deadline or ""),
                    "created_at": solution.get("created_at", now),
                    "updated_at": now,
                    "iteration_count": solution.get("iteration_count", 0),
                    "user_notes": notes,
                    "score": solution.get("score", 0),
                }
                append_answer(answer_register_path, record)
                append_history(history_db_path, {**solution, **record})
                st.success("Saved to answer register and local history database.")
        else:
            st.info("Solve a question first to open the AI Answer Studio.")

    with tabs[6]:
        if solution:
            score = int(solution.get("score", 0) or 0)
            st.markdown(f"<div class='pse-card'><h4>Confidence: {solution.get('confidence','Low')} ({score}/100)</h4><div class='pse-meter'><span style='width:{score}%'></span></div></div>", unsafe_allow_html=True)
            checks = pd.DataFrame(solution.get("validation", {}).get("checks", []))
            if not checks.empty:
                st.dataframe(checks, width="stretch", hide_index=True)
            missing = solution.get("validation", {}).get("missing_evidence", [])
            if missing:
                st.error("Missing evidence / risks:\n\n" + "\n".join(f"- {item}" for item in missing))
            else:
                st.success("No unresolved evidence gap detected by the current validator.")
        else:
            st.info("No validation result yet.")

    with tabs[7]:
        log = pd.DataFrame(solution.get("iteration_log", [])) if solution else pd.DataFrame()
        if not log.empty:
            st.dataframe(log, width="stretch", hide_index=True)
        else:
            st.info("Run the solver to generate an iteration log.")

    with tabs[8]:
        if solution:
            _solution_downloads(solution)
        register = load_answer_register(answer_register_path)
        st.download_button("Download Answer Register CSV", register.to_csv(index=False), "problem_solving_answer_register.csv", "text/csv", width="stretch")
        st.download_button("Download Question Bank CSV", question_bank.to_csv(index=False), "problem_solving_engine_question_bank.csv", "text/csv", width="stretch")
        if not register.empty:
            st.dataframe(register.tail(25), width="stretch", hide_index=True)
