import json
import os
from html import escape
from pathlib import Path

import pandas as pd
import streamlit as st

from excel import export_research, parse_research_response, research_query
from datacleaner import apply_operations, clean_pipeline, interpret_request, summarize
from presentation import powerpoint, presentation, refine_presentation
from reader import extract_text_from_image, parse_ocr_to_dataframe, process_image_ocr
from research_assistant import research_verified
from research_planner import generate_research_plan, refine_research_plan


BASE_DIR = Path(__file__).resolve().parent


def load_key_env():
    """Load local key.env values without requiring an additional package."""
    for filename in ("key.env", ".env"):
        env_file = BASE_DIR / filename
        if not env_file.exists():
            continue

        for line in env_file.read_text().splitlines():
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, value = line.split("=", 1)
            os.environ[key.strip()] = value.strip().strip('"').strip("'")


load_key_env()


def friendly_error_message(error):
    """Translate a Gemini rate-limit error into an actionable message."""
    if "429" in str(error) or "RESOURCE_EXHAUSTED" in str(error):
        return (
            "Gemini rate limit reached. Wait a bit before trying again, or check your "
            "plan and usage at https://ai.dev/rate-limit."
        )
    return str(error)


def is_rate_limit_error(error):
    message = str(error)
    return "429" in message or "RESOURCE_EXHAUSTED" in message or "rate limit" in message.lower()


def demo_chat_result(mode, prompt):
    """Return clearly labeled local output so the chat UX can be tested without quota."""
    if mode == "Ask anything":
        return {
            "content": (
                "Demo response: I would research this question, compare reliable sources, "
                "and show citations here once Gemini is available. Your request was: "
                f"{prompt}"
            ),
            "sources": [],
        }
    if mode == "Create a presentation":
        return {
            "presentation": {
                "title": "Demo presentation",
                "slides": [
                    {"title": "The question", "bullets": [prompt, "Context and audience framing"]},
                    {"title": "Key insight", "bullets": ["Evidence-led point one", "Evidence-led point two"]},
                    {"title": "Recommended next step", "bullets": ["Turn the strongest insight into an action"]},
                ],
            },
            "content": "Demo presentation drafted. Open Presentation Studio to refine it.",
        }
    return {
        "plan": {
            "title": "Demo research plan",
            "thesis_statement": prompt,
            "objectives": [{"title": "Define the question", "description": "Clarify the scope and success criteria."}],
            "literature_review": {"scope": "Review the strongest current evidence.", "key_themes": ["Context", "Evidence"], "sources_to_explore": ["Primary sources"]},
            "methodology": {"research_design": "Mixed methods", "data_collection": ["Document review"], "analysis_approach": "Compare themes and evidence.", "sample_size_or_scope": "Define during scoping."},
            "timeline": [{"phase": "Phase 1", "description": "Scope and review", "duration_weeks": 2}],
            "resources_needed": [],
            "expected_outcomes": ["A defensible answer"],
            "potential_challenges": [],
        },
        "content": "Demo research plan drafted. Open Research Planning to refine it.",
    }


st.set_page_config(page_title="Multipanel", page_icon="📊", layout="wide")
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Space+Grotesk:wght@500;600;700&display=swap');

    :root {
        --ink: #182126;
        --muted: #66737a;
        --line: #dce5e7;
        --mint: #d8f3e9;
        --teal: #167c72;
        --amber: #f2b84b;
        --paper: #f7faf9;
    }
    .stApp { background: var(--paper); color: var(--ink); }
    .block-container { max-width: 1180px; padding-top: 2.4rem; padding-bottom: 4rem; }
    h1, h2, h3, [data-testid="stTab"] { font-family: 'Space Grotesk', sans-serif; color: var(--ink); }
    p, label, .stMarkdown, [data-testid="stMetricValue"] { font-family: 'DM Sans', sans-serif; }
    h1 { font-size: clamp(2.4rem, 5vw, 4.5rem); letter-spacing: -0.04em; line-height: 0.98; margin-bottom: 0.35rem; }
    h2 { letter-spacing: -0.025em; }
    [data-testid="stHeader"] { background: transparent; }
    [data-testid="stTabs"] { margin-top: 2rem; }
    [data-testid="stTab"] { font-weight: 600; }
    [data-testid="stTabHighlight"] { background: var(--teal); }
    .app-kicker { color: var(--teal); font: 700 0.76rem 'DM Sans', sans-serif; letter-spacing: 0.14em; text-transform: uppercase; margin-bottom: 0.65rem; }
    .app-subtitle { color: var(--muted); font-size: 1.05rem; max-width: 680px; margin-bottom: 0; }
    .tool-intro { border-left: 4px solid var(--amber); padding: 0.15rem 0 0.15rem 1rem; margin: 0.4rem 0 1.5rem; }
    .tool-intro p { color: var(--muted); margin: 0.25rem 0 0; }
    .result-card { background: white; border: 1px solid var(--line); border-radius: 14px; padding: 1rem 1.2rem; margin: 1.25rem 0 0.9rem; box-shadow: 0 8px 24px rgba(24, 33, 38, 0.05); }
    .result-label { color: var(--teal); font: 700 0.72rem 'DM Sans', sans-serif; letter-spacing: 0.12em; text-transform: uppercase; }
    .result-title { color: var(--ink); font: 700 1.25rem 'Space Grotesk', sans-serif; margin-top: 0.25rem; }
    .result-meta { color: var(--muted); font-size: 0.88rem; margin-top: 0.2rem; }
    .preview-kicker { color: var(--teal); font: 700 0.72rem 'DM Sans', sans-serif; letter-spacing: 0.13em; text-transform: uppercase; margin: 1.6rem 0 0.35rem; }
    .slide-card { background: white; border: 1px solid var(--line); border-radius: 12px; min-height: 205px; padding: 1rem; box-shadow: 0 5px 16px rgba(24, 33, 38, 0.04); }
    .slide-number { color: var(--amber); font: 700 0.72rem 'DM Sans', sans-serif; letter-spacing: 0.12em; text-transform: uppercase; }
    .slide-title { color: var(--ink); font: 700 1.08rem 'Space Grotesk', sans-serif; line-height: 1.15; margin: 0.5rem 0 0.7rem; }
    .slide-bullet { color: var(--muted); font-size: 0.88rem; line-height: 1.4; margin: 0.35rem 0; }
    .source-card { background: white; border: 1px solid var(--line); border-radius: 10px; padding: 0.75rem 0.9rem; margin: 0.45rem 0; }
    .source-card a { color: var(--teal); font-weight: 700; text-decoration: none; }
    .source-meta { color: var(--muted); font-size: 0.8rem; margin-top: 0.2rem; }
    .answer-panel { background: #eef8f4; border: 1px solid #cce8dc; border-radius: 14px; padding: 1rem 1.2rem; }
    .workspace-card { background: white; border: 1px solid var(--line); border-radius: 14px; padding: 1.1rem 1.2rem; min-height: 150px; box-shadow: 0 5px 16px rgba(24, 33, 38, 0.04); }
    .workspace-card h3 { margin: 0 0 0.45rem; font-size: 1.05rem; }
    .workspace-card p { color: var(--muted); font-size: 0.9rem; line-height: 1.45; }
    .workspace-tag { color: var(--teal); font: 700 0.68rem 'DM Sans', sans-serif; letter-spacing: 0.12em; text-transform: uppercase; }
    .chat-shell { background: white; border: 1px solid var(--line); border-radius: 16px; padding: 1rem 1.2rem; box-shadow: 0 8px 24px rgba(24, 33, 38, 0.05); }
    .chat-note { color: var(--muted); font-size: 0.9rem; margin: 0.25rem 0 1rem; }
    .ask-hero { max-width: 760px; margin: 2.6rem auto 1.4rem; text-align: center; }
    .ask-hero h2 { font-size: clamp(2rem, 4vw, 3.4rem); margin-bottom: 0.45rem; letter-spacing: -0.045em; }
    .ask-hero p { color: var(--muted); margin: 0; }
    .chat-composer { max-width: 820px; margin: 0 auto; }
    .chat-actions { max-width: 820px; margin: 0.8rem auto 0; }
    .stDownloadButton > button { border-radius: 9px; font-weight: 700; }
    div[data-testid="stForm"] { background: white; border: 1px solid var(--line); border-radius: 14px; padding: 1.25rem 1.35rem 0.75rem; }
    </style>
    """,
    unsafe_allow_html=True,
)


def tool_intro(title, description):
    st.markdown(
        f'<div class="tool-intro"><h2>{title}</h2><p>{description}</p></div>',
        unsafe_allow_html=True,
    )


def result_card(file_name, description):
    st.markdown(
        f'<div class="result-card"><div class="result-label">Ready to export</div>'
        f'<div class="result-title">{escape(str(file_name))}</div>'
        f'<div class="result-meta">{escape(str(description))}</div></div>',
        unsafe_allow_html=True,
    )


def preview_kicker(label):
    st.markdown(f'<div class="preview-kicker">{label}</div>', unsafe_allow_html=True)


def slide_card(index, slide_data):
    bullets = "".join(
        f'<div class="slide-bullet">{escape(str(bullet))}</div>'
        for bullet in slide_data.get("bullets", [])
    )
    st.markdown(
        f'<div class="slide-card"><div class="slide-number">Slide {index:02d}</div>'
        f'<div class="slide-title">{escape(str(slide_data.get("title", "Untitled slide")))}</div>{bullets}</div>',
        unsafe_allow_html=True,
    )


def workspace_card(tag, title, description):
    st.markdown(
        f'<div class="workspace-card"><div class="workspace-tag">{tag}</div>'
        f'<h3>{title}</h3><p>{description}</p></div>',
        unsafe_allow_html=True,
    )


def render_presentation_result(presentation_info):
    output_path = BASE_DIR / "report.pptx"
    presentation(presentation_info, output_path)
    result_card("report.pptx", f"{len(presentation_info.get('slides', []))} slides ready for review")
    preview_kicker("Presentation preview")
    st.subheader(presentation_info.get("title", "Presentation"))
    st.caption("A visual story pass before you export the deck.")
    slide_columns = st.columns(2)
    for index, slide_data in enumerate(presentation_info.get("slides", []), start=1):
        with slide_columns[(index - 1) % 2]:
            slide_card(index, slide_data)
    preview_kicker("Refine with chat")
    revision = st.text_input(
        "Tell Multipanel what to change",
        placeholder="Make the opening sharper, simplify slide 3, and add a stronger conclusion.",
        key="presentation_revision",
    )
    if st.button("Apply revision", key="apply_presentation_revision", type="secondary"):
        if not revision.strip():
            st.warning("Describe the change you want first.")
        else:
            try:
                with st.spinner("Reworking the presentation..."):
                    revised = refine_presentation(presentation_info, revision.strip())
                st.session_state["presentation_result"] = revised
                st.rerun()
            except Exception as error:
                st.error(f"Could not revise the presentation: {friendly_error_message(error)}")
    preview_kicker("Export")
    st.download_button(
        "Download PowerPoint",
        data=output_path.read_bytes(),
        file_name="report.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        key="download_presentation",
    )


def render_plan_result(plan):
    result_card("research_plan.json", f"{len(plan.get('objectives', []))} objectives and {len(plan.get('timeline', []))} timeline phases")
    preview_kicker("Plan preview")
    with st.container(border=True):
        st.subheader(plan.get("title", "Research Project Plan"))
        st.markdown("**Thesis Statement:**")
        st.markdown(plan.get("thesis_statement", ""))
        st.subheader("Research Objectives")
        for obj in plan.get("objectives", []):
            with st.expander(obj.get("title", "Objective")):
                st.write(obj.get("description", ""))
    lit_review = plan.get("literature_review", {})
    with st.container(border=True):
        st.subheader("Literature Review")
        st.write("**Scope:**", lit_review.get("scope", ""))
        st.write("**Key themes:**", ", ".join(lit_review.get("key_themes", [])))
        st.write("**Sources to explore:**", ", ".join(lit_review.get("sources_to_explore", [])))
    methodology = plan.get("methodology", {})
    with st.container(border=True):
        st.subheader("Methodology")
        st.write("**Research Design:**", methodology.get("research_design", ""))
        st.write("**Data Collection Methods:**")
        for method in methodology.get("data_collection", []):
            st.write(f"- {method}")
        st.write("**Analysis Approach:**", methodology.get("analysis_approach", ""))
        st.write("**Sample Size/Scope:**", methodology.get("sample_size_or_scope", ""))
    timeline_data = [
        {"Phase": phase.get("phase", ""), "Description": phase.get("description", ""), "Duration (weeks)": phase.get("duration_weeks", 0)}
        for phase in plan.get("timeline", [])
    ]
    if timeline_data:
        preview_kicker("Project timeline")
        with st.container(border=True):
            st.dataframe(pd.DataFrame(timeline_data), use_container_width=True, hide_index=True)
    with st.container(border=True):
        st.subheader("Resources and deliverables")
        for resource in plan.get("resources_needed", []):
            st.write(f"**{resource.get('resource', '')}:** {resource.get('justification', '')}")
        for outcome in plan.get("expected_outcomes", []):
            st.write(f"- {outcome}")
    if plan.get("potential_challenges"):
        with st.expander("Potential challenges and mitigation"):
            for challenge in plan["potential_challenges"]:
                st.write(f"**{challenge.get('challenge', '')}:** {challenge.get('mitigation', '')}")
    preview_kicker("Refine with chat")
    revision = st.text_input(
        "Tell Multipanel what to change",
        placeholder="Make the methodology more practical and shorten the timeline to 12 weeks.",
        key="plan_revision",
    )
    if st.button("Apply revision", key="apply_plan_revision", type="secondary"):
        if not revision.strip():
            st.warning("Describe the change you want first.")
        else:
            try:
                with st.spinner("Reworking the research plan..."):
                    revised = refine_research_plan(plan, revision.strip())
                st.session_state["plan_result"] = revised
                st.rerun()
            except Exception as error:
                st.error(f"Could not revise the research plan: {friendly_error_message(error)}")
    preview_kicker("Export")
    with st.expander("Inspect JSON export"):
        st.json(plan)
    st.download_button(
        "Download plan as JSON",
        data=json.dumps(plan, indent=2),
        file_name="research_plan.json",
        mime="application/json",
        key="download_plan",
    )


def render_excel_result(excel_result):
    output_path = Path(excel_result["path"])
    dataframe = excel_result["dataframe"]
    summary = excel_result.get("summary") or ""
    sources = excel_result.get("sources") or []
    result_card("ai_table.xlsx", f"{len(dataframe):,} research rows with overview and source sheets")
    preview_kicker("Workbook preview")
    overview_column, stats_column = st.columns([2.2, 1], gap="large")
    with overview_column:
        with st.container(border=True):
            st.subheader("Executive overview")
            st.write(summary or "No executive summary was returned.")
    with stats_column:
        with st.container(border=True):
            st.metric("Research rows", f"{len(dataframe):,}")
            st.metric("Columns", len(dataframe.columns))
            st.metric("Source links", len(sources))
    with st.container(border=True):
        st.subheader("Research data")
        st.dataframe(dataframe, use_container_width=True, hide_index=True)
    if output_path.exists():
        workbook_sheets = pd.read_excel(output_path, sheet_name=None)
        for sheet_name, sheet_dataframe in workbook_sheets.items():
            if sheet_name == "Research Data":
                continue
            with st.expander(f"Workbook sheet: {sheet_name}"):
                st.dataframe(sheet_dataframe, use_container_width=True, hide_index=True)
        preview_kicker("Export")
        st.download_button(
            "Download Excel workbook",
            data=output_path.read_bytes(),
            file_name="ai_table.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_excel",
        )
    else:
        st.error("The Excel workbook file is missing. Create the workbook again to export it.")


def render_verified_result(report):
    result_card(
        "verified_research_report.json",
        f"{len(report.get('claims', []))} claims checked against {len(report.get('sources', []))} sources",
    )
    preview_kicker("Research brief")
    with st.container(border=True):
        st.subheader("Answer")
        st.write(report.get("answer", "No answer returned."))
        st.caption(f"Checked at {report.get('verified_at', 'unknown')}")

    preview_kicker("Evidence ledger")
    source_map = {source["id"]: source for source in report.get("sources", [])}
    for claim in report.get("claims", []):
        with st.container(border=True):
            confidence = str(claim.get("confidence", "unknown")).upper()
            st.markdown(f"**{escape(confidence)} confidence · {escape(str(claim.get('kind', 'claim')))}**")
            st.write(claim.get("text", ""))
            cited_sources = [source_map[source_id] for source_id in claim.get("source_ids", []) if source_id in source_map]
            for source in cited_sources:
                st.markdown(
                    f"Source: [{source['title']}]({source['url']}) · "
                    f"{source['publisher']} · published {source['published_date']}"
                )

    preview_kicker("Source desk")
    for source in report.get("sources", []):
        with st.container(border=True):
            st.markdown(f"[{source['title']}]({source['url']})")
            st.caption(f"{source['publisher']} · published {source['published_date']}")

    if report.get("limitations"):
        st.subheader("Limitations")
        for limitation in report["limitations"]:
            st.warning(limitation)

    preview_kicker("Export preview")
    st.caption("The report above is the readable demo. Open the export inspector to see the exact JSON file.")
    with st.expander("Inspect JSON export"):
        st.json(report)
    st.download_button(
        "Download verified report",
        data=json.dumps(report, indent=2),
        file_name="verified_research_report.json",
        mime="application/json",
        key="download_verified",
    )


def render_reader_result(reader_result):
    markdown_text = reader_result.get("markdown") or ""
    dataframe = reader_result.get("dataframe")
    excel_bytes = reader_result.get("excel_bytes")
    st.subheader("Extracted text")
    st.markdown(markdown_text)
    if dataframe is not None:
        result_card("ocr_output.xlsx + ocr_output.csv", f"Detected a table with {len(dataframe):,} rows and {len(dataframe.columns)} columns")
        with st.expander("Table preview", expanded=True):
            st.subheader("Extracted table")
            st.dataframe(dataframe.head(50), use_container_width=True)
        if excel_bytes:
            st.download_button(
                "Download as Excel",
                data=excel_bytes,
                file_name="ocr_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="download_ocr_excel",
            )
        st.download_button(
            "Download as CSV",
            data=dataframe.to_csv(index=False).encode("utf-8"),
            file_name="ocr_output.csv",
            mime="text/csv",
            key="download_ocr_csv",
        )
    else:
        st.info("No tables were detected in the image. Check the extracted text above.")


st.markdown('<div class="app-kicker">One workspace · many finished artifacts</div>', unsafe_allow_html=True)
st.title("Multipanel")
st.markdown('<p class="app-subtitle">A focused AI workspace for asking, researching, drafting, cleaning, and exporting from one place.</p>', unsafe_allow_html=True)

chat_home, files_home, capabilities_home = st.tabs(["Chat", "Files", "Capabilities"])

with chat_home:
    st.markdown(
        '<div class="ask-hero"><div class="app-kicker">Your workspace copilot</div>'
        '<h2>What can I help you make?</h2>'
        '<p>Ask a question, build a deck, plan research, or bring a file into the workspace.</p></div>',
        unsafe_allow_html=True,
    )
    st.markdown('<div class="chat-composer">', unsafe_allow_html=True)
    chat_mode = st.radio(
        "Mode",
        ["Ask anything", "Create a presentation", "Make a research plan"],
        horizontal=True,
        label_visibility="collapsed",
        key="chat_mode",
    )
    demo_mode = st.checkbox(
        "Demo mode",
        help="Use local sample output without calling Gemini. Useful while an API quota is unavailable.",
        key="chat_demo_mode",
    )
    with st.expander("Attach context", expanded=False):
        st.file_uploader(
            "Add a document or image",
            type=["csv", "xlsx", "xls", "pdf", "txt", "md", "png", "jpg", "jpeg"],
            key="chat_attachment",
        )
    st.markdown('</div>', unsafe_allow_html=True)

    history = st.session_state.get("chat_history", [])
    if not history:
        with st.container(border=True):
            st.markdown("**Ask Multipanel**")
            st.caption("I can answer with sources, draft a presentation, or turn an idea into a research plan.")
    else:
        for message in history:
            with st.chat_message(message["role"]):
                st.write(message["content"])
                if message.get("sources"):
                    st.caption("Sources")
                    for source in message["sources"]:
                        st.markdown(f"- [{source['title']}]({source['url']})")

    chat_prompt = st.chat_input(
        "Ask Multipanel anything...",
        key="ask_multipanel_input",
    )
    if chat_prompt:
        try:
            st.session_state["chat_history"] = history + [{"role": "user", "content": chat_prompt}]
            with st.spinner("Thinking through your request..."):
                if demo_mode:
                    demo_result = demo_chat_result(chat_mode, chat_prompt)
                    if chat_mode == "Create a presentation":
                        st.session_state["presentation_result"] = demo_result["presentation"]
                    elif chat_mode == "Make a research plan":
                        st.session_state["plan_result"] = demo_result["plan"]
                    assistant_message = {"role": "assistant", "content": demo_result["content"] + " (Demo mode)"}
                elif chat_mode == "Ask anything":
                    chat_result = research_verified(chat_prompt)
                    assistant_message = {
                        "role": "assistant",
                        "content": chat_result.get("answer", "No answer returned."),
                        "sources": chat_result.get("sources", []),
                    }
                elif chat_mode == "Create a presentation":
                    chat_result = powerpoint(chat_prompt)
                    st.session_state["presentation_result"] = chat_result
                    assistant_message = {
                        "role": "assistant",
                        "content": f"I drafted a {len(chat_result.get('slides', []))}-slide presentation. Open the Presentation capability to review, refine, and export it.",
                    }
                else:
                    chat_result = generate_research_plan(chat_prompt)
                    st.session_state["plan_result"] = chat_result
                    assistant_message = {
                        "role": "assistant",
                        "content": "I drafted a research plan. Open Research Planning to review, refine, and export it.",
                    }
            st.session_state["chat_history"].append(assistant_message)
            st.rerun()
        except Exception as error:
            if is_rate_limit_error(error):
                st.session_state["chat_history"].append({
                    "role": "assistant",
                    "content": (
                        "Gemini is temporarily rate-limited. Your message is saved above. "
                        "Wait for the quota window to reset, or enable Demo mode to preview the workflow without making an AI request."
                    ),
                })
                st.rerun()
            else:
                st.error(f"Could not complete that request: {friendly_error_message(error)}")

    st.markdown('<div class="chat-actions">', unsafe_allow_html=True)
    quick_columns = st.columns(3)
    quick_prompts = [
        "Compare two ideas with sources",
        "Draft a polished presentation",
        "Plan my research project",
    ]
    for index, prompt in enumerate(quick_prompts):
        with quick_columns[index]:
            st.caption(prompt)
    st.markdown('</div>', unsafe_allow_html=True)

with files_home:
    tool_intro("Docs and files", "Keep working files together, preview supported documents, and send a file into the right Multipanel capability.")
    uploaded_docs = st.file_uploader(
        "Add files to your workspace",
        type=["csv", "xlsx", "xls", "pdf", "txt", "md", "png", "jpg", "jpeg"],
        accept_multiple_files=True,
        key="workspace_files",
    )
    if uploaded_docs:
        file_columns = st.columns(2)
        for index, uploaded_doc in enumerate(uploaded_docs):
            with file_columns[index % 2]:
                with st.container(border=True):
                    st.markdown(f"**{uploaded_doc.name}**")
                    st.caption(f"{uploaded_doc.type or 'file'} · {uploaded_doc.size:,} bytes")
                    suffix = Path(uploaded_doc.name).suffix.lower()
                    if suffix == ".csv":
                        st.dataframe(pd.read_csv(uploaded_doc).head(12), use_container_width=True, hide_index=True)
                    elif suffix in {".xlsx", ".xls"}:
                        sheets = pd.read_excel(uploaded_doc, sheet_name=None)
                        first_sheet = next(iter(sheets.values()), pd.DataFrame())
                        st.dataframe(first_sheet.head(12), use_container_width=True, hide_index=True)
                    elif suffix in {".txt", ".md"}:
                        st.code(uploaded_doc.getvalue().decode("utf-8", errors="replace")[:2_000], language="markdown")
                    elif suffix in {".png", ".jpg", ".jpeg"}:
                        st.image(uploaded_doc, use_container_width=True)
                    else:
                        st.info("Ready to send to Reader when needed.")

with capabilities_home:
    tool_intro("Capabilities", "Choose a focused workspace. Each capability keeps its own preview, refinement controls, and export path.")
    capability_columns = st.columns(3)
    capabilities = [
        ("Create", "Presentation Studio", "Draft and refine a polished PowerPoint narrative.", "Start with a topic, audience, and tone."),
        ("Research", "Verified Research", "Answer current questions with claim-level evidence.", "Use sources, citations, and visible uncertainty."),
        ("Plan", "Research Planner", "Turn a thesis into a timeline, method, and deliverables.", "Refine the plan conversationally."),
        ("Build", "Excel Research", "Create sourced tables, charts, and workbook sheets.", "Export an analysis-ready workbook."),
        ("Clean", "Data Cleaner", "Upload data and change it with plain language.", "Every operation is whitelisted and previewed."),
        ("Extract", "Document Reader", "Turn image tables into structured data.", "Preview OCR before exporting CSV or Excel."),
    ]
    for index, (tag, title, description, detail) in enumerate(capabilities):
        with capability_columns[index % 3]:
            workspace_card(tag, title, description)
            st.caption(detail)

presentation_tab, excel_tab, verified_tab, research_tab, cleaning_tab, reader_tab = st.tabs(["Presentation", "Excel research", "Verified Research", "Research Planning", "Data cleaning", "Reader"])


with presentation_tab:
    tool_intro("Presentation studio", "Build a focused slide narrative, review the story, then export the finished deck.")
    with st.form("presentation_form"):
        user_query = st.text_area(
            "Topic or thesis statement",
            placeholder="Remote work improves productivity when companies use outcome-based management.",
            height=120,
        )
        slide_count = st.slider("Number of slides", min_value=2, max_value=15, value=6)
        audience = st.text_input("Audience", value="general audience")
        tone = st.text_input("Tone", value="clear and persuasive")
        submitted = st.form_submit_button("Create presentation", type="primary")

    if submitted:
        if not user_query.strip():
            st.error("Enter a topic or thesis statement first.")
        else:
            try:
                with st.spinner("Creating your presentation plan..."):
                    presentation_info = powerpoint(
                        user_query=user_query.strip(),
                        slide_count=slide_count,
                        audience=audience.strip() or "general audience",
                        tone=tone.strip() or "clear and persuasive",
                    )

                st.session_state["presentation_result"] = presentation_info
                st.success("Presentation created.")
            except Exception as error:
                st.error(f"Could not create the presentation: {friendly_error_message(error)}")

    if st.session_state.get("presentation_result") is not None:
        render_presentation_result(st.session_state["presentation_result"])

with excel_tab:
    tool_intro("Research workbook", "Turn a research question into sourced tables, an overview sheet, and a workbook you can share.")
    with st.form("excel_form"):
        research_request = st.text_area(
            "Research question",
            placeholder="Compare renewable energy investment trends in the US and Europe from 2020 to 2025.",
            height=120,
        )
        excel_submitted = st.form_submit_button("Create Excel workbook", type="primary")

    if excel_submitted:
        if not research_request.strip():
            st.error("Enter a research question first.")
        else:
            try:
                with st.spinner("Researching and building your workbook..."):
                    response_text = research_query(research_request.strip())
                    summary, sources, dataframe = parse_research_response(response_text)
                    output_path = BASE_DIR / "ai_table.xlsx"
                    export_research(
                        dataframe,
                        excel_file=output_path,
                        summary=summary,
                        sources=sources,
                        query=research_request.strip(),
                    )

                st.session_state["excel_result"] = {
                    "dataframe": dataframe,
                    "summary": summary,
                    "sources": sources,
                    "path": str(output_path),
                    "query": research_request.strip(),
                }
                st.success("Excel workbook created.")
            except Exception as error:
                st.error(f"Could not create the Excel workbook: {friendly_error_message(error)}")

    if st.session_state.get("excel_result") is not None:
        render_excel_result(st.session_state["excel_result"])

with verified_tab:
    tool_intro("Verified research", "Build a claim-level answer with evidence, source links, and visible uncertainty.")
    with st.form("verified_research_form"):
        verified_question = st.text_area(
            "What do you need to verify?",
            placeholder="What were the latest global renewable energy capacity additions, according to primary sources?",
            height=120,
        )
        date_range = st.text_input("Date requirement", value="latest available")
        source_preference = st.text_input("Preferred sources", value="credible primary sources")
        verified_submitted = st.form_submit_button("Research and verify", type="primary")

    if verified_submitted:
        if not verified_question.strip():
            st.error("Enter a research question first.")
        else:
            try:
                with st.spinner("Searching current sources and checking claims..."):
                    report = research_verified(
                        verified_question.strip(),
                        date_range=date_range.strip() or "latest available",
                        source_preference=source_preference.strip() or "credible primary sources",
                    )

                st.session_state["verified_result"] = report
                st.success(f"Verified report generated as of {report['as_of']}.")
            except Exception as error:
                st.error(f"Could not verify this research request: {friendly_error_message(error)}")

    if st.session_state.get("verified_result") is not None:
        render_verified_result(st.session_state["verified_result"])

with research_tab:
    tool_intro("Research planner", "Shape a thesis or research question into a practical plan with milestones, methods, and deliverables.")
    with st.form("research_form"):
        thesis_statement = st.text_area(
            "Thesis statement or research question",
            placeholder="How does artificial intelligence impact job market dynamics in developing economies?",
            height=120,
        )
        field_of_study = st.text_input("Field of study", value="Research")
        research_type = st.selectbox("Research type", ["Thesis/Dissertation", "Research Paper", "Grant Proposal", "Capstone Project"])
        duration_months = st.slider("Expected duration (months)", min_value=3, max_value=48, value=12)
        research_submitted = st.form_submit_button("Generate research plan", type="primary")

    if research_submitted:
        if not thesis_statement.strip():
            st.error("Enter a thesis statement or research question first.")
        else:
            try:
                with st.spinner("Generating your research plan..."):
                    plan = generate_research_plan(
                        thesis_statement=thesis_statement.strip(),
                        field_of_study=field_of_study.strip() or "Research",
                        research_type=research_type,
                        duration_months=duration_months,
                    )

                st.session_state["plan_result"] = plan
                st.success("Research plan created!")
            except Exception as error:
                st.error(f"Could not create the research plan: {friendly_error_message(error)}")

    if st.session_state.get("plan_result") is not None:
        render_plan_result(st.session_state["plan_result"])

with cleaning_tab:
    tool_intro("Data cleaner", "Upload a CSV or workbook, inspect the cleaned result, then refine it with plain-language requests.")
    uploaded_file = st.file_uploader("Data file", type=["csv", "xlsx", "xls"])

    if uploaded_file is not None:
        if st.session_state.get("cleaning_source_name") != uploaded_file.name:
            try:
                if uploaded_file.name.lower().endswith(".csv"):
                    raw_df = pd.read_csv(uploaded_file)
                else:
                    raw_df = pd.read_excel(uploaded_file)

                st.session_state["cleaning_source_name"] = uploaded_file.name
                st.session_state["cleaning_df"] = clean_pipeline(raw_df)
                st.session_state["cleaning_chat"] = [
                    {"role": "assistant", "content": f"Loaded and auto-cleaned {len(raw_df)} rows. Ask me to make further changes, e.g. \"drop rows where amount is null\" or \"rename region to area\"."}
                ]
            except Exception as error:
                st.error(f"Could not read this file: {friendly_error_message(error)}")
                st.session_state.pop("cleaning_df", None)

    cleaned_df = st.session_state.get("cleaning_df")
    if cleaned_df is not None:
        result_card("cleaned_data.csv", f"{len(cleaned_df):,} rows · {len(cleaned_df.columns)} columns")
        st.subheader("Current data")
        st.dataframe(cleaned_df.head(50), use_container_width=True)

        with st.expander("Column summary"):
            stats = summarize(cleaned_df)
            st.write({"rows": stats["rows"], "columns": stats["columns"]})
            st.write("Missing values by column:")
            st.json(stats["missing_by_column"])
            st.write("Data types:")
            st.json(stats["dtypes"])

        st.download_button(
            "Download cleaned CSV",
            data=cleaned_df.to_csv(index=False).encode("utf-8"),
            file_name="cleaned_data.csv",
            mime="text/csv",
        )

        st.subheader("Chat with your data")
        for message in st.session_state.get("cleaning_chat", []):
            with st.chat_message(message["role"]):
                st.markdown(message["content"])

        chat_request = st.chat_input("What would you like to change?", key="cleaning_chat_input")
        if chat_request:
            st.session_state["cleaning_chat"].append({"role": "user", "content": chat_request})
            try:
                with st.spinner("Applying your request..."):
                    operations = interpret_request(cleaned_df, chat_request)
                    if not operations:
                        reply = "I couldn't turn that into a supported cleaning step. Try being more specific about a column and action."
                    else:
                        updated_df, log = apply_operations(cleaned_df, operations)
                        st.session_state["cleaning_df"] = updated_df
                        reply = "\n".join(f"- {line}" for line in log)
            except Exception as error:
                reply = f"Could not apply that request: {friendly_error_message(error)}"

            st.session_state["cleaning_chat"].append({"role": "assistant", "content": reply})
            st.rerun()

with reader_tab:
    tool_intro("Document reader", "Extract text and tables from an image, inspect the result, then export the structured data.")
    uploaded_image = st.file_uploader("Image file", type=["jpg", "jpeg", "png", "gif", "bmp"])

    if uploaded_image is not None:
        reader_key = f"{uploaded_image.name}:{uploaded_image.size}"
        if st.session_state.get("reader_source_key") != reader_key:
            try:
                with st.spinner("Extracting text from image..."):
                    output_path = BASE_DIR / "ocr_output.xlsx"
                    markdown_text, df, saved_path = process_image_ocr(uploaded_image, output_path)
                st.session_state["reader_source_key"] = reader_key
                st.session_state["reader_result"] = {
                    "markdown": markdown_text,
                    "dataframe": df,
                    "excel_bytes": Path(saved_path).read_bytes() if saved_path and Path(saved_path).exists() else None,
                }
            except Exception as error:
                st.session_state.pop("reader_result", None)
                st.session_state.pop("reader_source_key", None)
                st.error(f"Could not process the image: {friendly_error_message(error)}")

    if st.session_state.get("reader_result") is not None:
        render_reader_result(st.session_state["reader_result"])
