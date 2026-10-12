#!/usr/bin/env python3
"""Streamlit web UI for the sprint performance reporting engine.

This module is a thin presentation layer over :mod:`sprint_report`. It collects
inputs in the browser, persists any uploads to a temporary working directory, and
reuses the exact same ingestion, charting, LLM, and PowerPoint routines that the
command-line interface uses. No business logic is duplicated here.

Run with::

    streamlit run streamlit_app.py
"""

from __future__ import annotations

import logging
import os
import tempfile
from datetime import date, datetime
from pathlib import Path
from typing import Any

import pandas as pd
import streamlit as st

from sprint_report import (
    DeveloperBrief,
    ReportConfig,
    SprintTask,
    _chart_for_developer,
    _fallback_brief,
    _fallback_team_brief,
    _format_date,
    _insight_line,
    _task_in_range,
    create_presentation,
    generate_developer_brief,
    generate_team_brief,
    load_tasks,
    resolve_sprint_range,
)

BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_DIR = BASE_DIR / "templates"
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "output"
SAMPLE_WORKBOOK = TEMPLATE_DIR / "sprint_tasks.xlsx"
SAMPLE_TEMPLATE = TEMPLATE_DIR / "master_template.pptx"
LOGGER = logging.getLogger("sprint_report.ui")

# provider label -> (default model, api-key env var, default api base)
PROVIDERS: dict[str, tuple[str, str, str]] = {
    "OpenAI": ("openai/gpt-4o-mini", "OPENAI_API_KEY", ""),
    "Anthropic": ("anthropic/claude-3-5-haiku-latest", "ANTHROPIC_API_KEY", ""),
    "Ollama (local)": ("ollama/llama3.1", "", "http://localhost:11434"),
    "Custom": ("", "", ""),
}


def _get_upload_dir() -> Path:
    """Return the directory that stores every file uploaded through the UI."""
    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    return UPLOAD_DIR


def _save_upload(uploaded: Any, directory: Path) -> Path:
    """Persist an uploaded file to disk and return its path."""
    target = directory / uploaded.name
    target.write_bytes(uploaded.getbuffer())
    return target


def _chart_bytes(developer: str, tasks: list[SprintTask], start: date, end: date) -> bytes:
    """Render a developer burndown to an in-memory PNG without touching the deck."""
    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as handle:
        path = Path(handle.name)
    try:
        _chart_for_developer(developer, tasks, start, end, path)
        return path.read_bytes()
    finally:
        path.unlink(missing_ok=True)


def _build_config(**kwargs: Any) -> ReportConfig:
    """Construct a :class:`ReportConfig` straight from the sidebar widgets."""
    return ReportConfig(**kwargs)


def _status_counts(tasks: list[SprintTask]) -> dict[str, int]:
    counts: dict[str, int] = {}
    for task in tasks:
        counts[task.status] = counts.get(task.status, 0) + 1
    return counts


def _render_team_metrics(tasks: list[SprintTask], start: date, end: date) -> None:
    """Show the high-level KPI cards shared with the summary slide."""
    total = len(tasks)
    done = sum(task.status == "Done" for task in tasks)
    blocked = sum(task.status == "Blocked" for task in tasks)
    hours = sum(task.logged_hours for task in tasks)
    completion = (done / total * 100) if total else 0.0

    left, middle, right, last = st.columns(4)
    left.metric("Tasks completed", f"{done}/{total}")
    middle.metric("Completion rate", f"{completion:.0f}%")
    right.metric("Blocked tasks", str(blocked))
    last.metric("Logged effort", f"{hours:.1f} h")
    st.caption(
        f"Sprint window: {start:%b %d, %Y} - {end:%b %d, %Y}  |  "
        f"{len({task.developer for task in tasks})} developers"
    )


def _render_report(payload: dict[str, Any]) -> None:
    """Render the download button, team KPIs, and per-developer tabs."""
    st.success(
        f"Report generated at {payload['generated_at']:%Y-%m-%d %H:%M:%S} "
        f"covering {payload['task_count']} tasks."
    )
    st.download_button(
        "Download PowerPoint report",
        data=payload["bytes"],
        file_name=payload["file_name"],
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        type="primary",
    )

    team_tab, *developer_tabs = st.tabs(
        ["Team overview"] + [dev["name"] for dev in payload["developers"]]
    )
    with team_tab:
        team = payload.get("team_brief")
        if team is not None:
            st.markdown("**Executive narrative**")
            st.write(team.narrative)
            st.markdown("**Risks & attention**")
            for item in team.risks:
                st.markdown(f"- {item}")
            st.divider()
        _render_team_metrics(payload["tasks"], payload["start"], payload["end"])
        counts = _status_counts(payload["tasks"])
        status_line = "  |  ".join(f"{name}: {value}" for name, value in sorted(counts.items()))
        st.markdown(f"**Status breakdown:** {status_line}")
        st.dataframe(
            pd.DataFrame(
                {
                    "Developer": [task.developer for task in payload["tasks"]],
                    "Task": [task.task for task in payload["tasks"]],
                    "Sprint": [task.sprint for task in payload["tasks"]],
                    "Status": [task.status for task in payload["tasks"]],
                    "Logged hours": [round(task.logged_hours, 2) for task in payload["tasks"]],
                }
            ),
            width="stretch",
            hide_index=True,
        )

    for dev, tab in zip(payload["developers"], developer_tabs):
        with tab:
            chart_col, text_col = st.columns([3, 2])
            with chart_col:
                st.image(dev["chart"], width="stretch")
            with text_col:
                st.markdown("**Executive summary**")
                st.write(dev["summary"])
                st.markdown("**Risk**")
                st.write(dev["risk"])
                st.markdown("**Next actions**")
                for item in dev["actions"]:
                    st.markdown(f"- {item}")
                st.markdown("**Metrics**")
                st.write(dev["insight"])
            st.dataframe(
                pd.DataFrame(
                    {
                        "Task": [task.task for task in dev["tasks"]],
                        "Sprint": [task.sprint for task in dev["tasks"]],
                        "Status": [task.status for task in dev["tasks"]],
                        "Start": [_format_date(task.start_date) for task in dev["tasks"]],
                        "End": [_format_date(task.end_date) for task in dev["tasks"]],
                    }
                ),
                width="stretch",
                hide_index=True,
            )


def _generate(
    *,
    config: ReportConfig,
    tasks: list[SprintTask],
    use_llm: bool,
    progress: Any = None,
) -> dict[str, Any]:
    """Run the full engine and package artifacts for browser rendering.

    ``progress`` is an optional callback ``(fraction, label)`` used to drive the
    UI loader. When omitted the function behaves exactly as before.
    """

    def notify(fraction: float, label: str) -> None:
        if progress is not None:
            progress(min(max(fraction, 0.0), 1.0), label)

    notify(0.0, "Resolving sprint date range...")
    start, end = resolve_sprint_range(tasks, config.sprint_start, config.sprint_end)
    selected = [task for task in tasks if _task_in_range(task, start, end)]
    if not selected:
        raise ValueError("No tasks overlap the selected sprint date range.")

    by_developer: dict[str, list[SprintTask]] = {}
    for task in selected:
        by_developer.setdefault(task.developer, []).append(task)
    ordered = sorted(by_developer.items())
    total = max(1, len(ordered))

    # Per-developer summaries, risk notes, and action items.
    briefs: dict[str, DeveloperBrief] = {}
    for index, (developer, developer_tasks) in enumerate(ordered):
        notify(0.05 + 0.35 * index / total, f"Analysing {developer}...")
        try:
            briefs[developer] = (
                generate_developer_brief(developer_tasks, config)
                if use_llm
                else _fallback_brief(developer_tasks)
            )
        except Exception as exc:  # defensive: never let one developer halt the batch
            LOGGER.warning("Brief failed for %s: %s", developer, exc)
            briefs[developer] = _fallback_brief(developer_tasks)

    # Team-level narrative and cross-developer risk flags.
    notify(0.45, "Summarising the team...")
    try:
        team_brief = (
            generate_team_brief(briefs, selected, config)
            if use_llm
            else _fallback_team_brief(selected)
        )
    except Exception as exc:
        LOGGER.warning("Team brief failed: %s", exc)
        team_brief = _fallback_team_brief(selected)

    notify(0.6, "Rendering slides and burndown charts...")
    output = create_presentation(
        config, selected, start, end, briefs=briefs, team_brief=team_brief
    )

    notify(0.9, "Preparing charts for preview...")
    developers: list[dict[str, Any]] = []
    for developer, developer_tasks in ordered:
        brief = briefs[developer]
        developers.append(
            {
                "name": developer,
                "summary": brief.summary,
                "risk": brief.risk,
                "actions": brief.actions,
                "insight": _insight_line(developer_tasks),
                "chart": _chart_bytes(developer, developer_tasks, start, end),
                "tasks": developer_tasks,
            }
        )
    notify(0.98, "Packaging the presentation...")

    return {
        "bytes": output.read_bytes(),
        "file_name": output.name,
        "task_count": len(selected),
        "start": start,
        "end": end,
        "tasks": selected,
        "developers": developers,
        "team_brief": team_brief,
        "generated_at": datetime.now(),
    }


def main() -> None:
    """Render the Streamlit page."""
    st.set_page_config(page_title="Sprint Report Generator", page_icon="📊", layout="wide")
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(name)s: %(message)s",
    )
    logging.getLogger("LiteLLM").setLevel(logging.WARNING)
    upload_dir = _get_upload_dir()
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    st.title("Sprint Performance Report Generator")
    st.caption(
        "Ingest sprint task data, review burndown and LLM insights, then export a "
        "branded PowerPoint deck."
    )

    # ------------------------------------------------------------------ sidebar
    with st.sidebar:
        st.header("1. Data source")
        source = st.radio("Workbook", ["Sample data", "Upload files"], horizontal=True)
        excel_upload = template_upload = None
        if source == "Sample data":
            input_path: Path | None = SAMPLE_WORKBOOK
            template_path: Path | None = SAMPLE_TEMPLATE
            st.caption(
                "Using the bundled `templates/sprint_tasks.xlsx` and "
                "`templates/master_template.pptx`."
            )
        else:
            excel_upload = st.file_uploader("Task workbook (.xlsx)", type=["xlsx"])
            template_upload = st.file_uploader("PowerPoint master (.pptx)", type=["pptx"])
            if excel_upload or template_upload:
                with st.spinner("Saving uploaded files..."):
                    input_path = _save_upload(excel_upload, upload_dir) if excel_upload else None
                    template_path = _save_upload(template_upload, upload_dir) if template_upload else None
            else:
                input_path = template_path = None
            st.caption(f"Uploads are stored in `{UPLOAD_DIR.relative_to(BASE_DIR)}/`.")

        st.header("2. Sprint")
        sprint_name = st.text_input("Sprint name filter (optional)", value="")
        auto_dates = st.checkbox("Auto-detect sprint dates from data", value=True)
        if auto_dates:
            start_date = end_date = None
        else:
            picked = st.date_input(
                "Sprint date range",
                value=(date(2026, 6, 1), date(2026, 6, 7)),
            )
            start_date, end_date = (
                (picked[0], picked[1]) if isinstance(picked, (tuple, list)) and len(picked) == 2 else (None, None)
            )

        st.header("3. LLM summaries")
        use_llm = st.checkbox("Use LLM for executive summaries", value=True)
        provider = st.selectbox("Provider", list(PROVIDERS.keys()))
        default_model, default_key_env, default_base = PROVIDERS[provider]
        model = st.text_input("Model identifier", value=default_model)
        api_key = st.text_input("API key (never stored)", type="password")
        api_base = st.text_input("API base URL", value=default_base)
        temperature = st.slider("Temperature", 0.0, 1.0, 0.2, 0.05)
        max_tokens = st.number_input("Max tokens", min_value=50, max_value=2000, value=300, step=50)
        timeout = st.number_input("Timeout (seconds)", min_value=5, max_value=300, value=45, step=5)

        st.header("4. PowerPoint")
        summary_layout_index = st.number_input(
            "Summary layout index", min_value=0, max_value=20, value=6, step=1
        )
        developer_layout_index = st.number_input(
            "Developer layout index", min_value=0, max_value=20, value=6, step=1
        )
        log_level = st.selectbox("Log level", ["INFO", "DEBUG", "WARNING", "ERROR"], index=0)

    # -------------------------------------------------------------- load & preview
    if not input_path or not input_path.is_file():
        st.info("Select or upload a task workbook to begin.")
        return
    if not template_path or not template_path.is_file():
        st.warning("A PowerPoint master template is required before generating a report.")

    try:
        with st.spinner("Loading and validating the workbook..."):
            tasks = load_tasks(input_path, sprint_name.strip() or None)
    except Exception as exc:
        st.error(f"Could not load the workbook: {exc}")
        return

    st.subheader("Data preview")
    preview_range: tuple[date, date] | None = None
    try:
        with st.spinner("Computing sprint metrics..."):
            preview_range = resolve_sprint_range(tasks, start_date, end_date)
    except Exception as exc:
        LOGGER.warning("Could not infer a preview sprint range: %s", exc)
    if preview_range:
        _render_team_metrics(tasks, preview_range[0], preview_range[1])
    try:
        with st.spinner("Rendering data table..."):
            raw = pd.read_excel(input_path)
        st.dataframe(raw, width="stretch", hide_index=True)
    except Exception as exc:
        LOGGER.warning("Could not render raw preview: %s", exc)
    st.caption(f"Validated {len(tasks)} task rows against the required schema.")

    # -------------------------------------------------------------- generate
    if st.button("Generate report", type="primary"):
        if not template_path or not template_path.is_file():
            st.error("Upload a PowerPoint master template before generating.")
            return
        api_key_env = default_key_env or "OPENAI_API_KEY"
        if api_key:
            os.environ[api_key_env] = api_key
        config = _build_config(
            input_file=input_path,
            template_file=template_path,
            output_file=OUTPUT_DIR / "sprint_report.pptx",
            sprint_start=start_date,
            sprint_end=end_date,
            sprint_name=sprint_name.strip() or None,
            model=model.strip() or default_model,
            temperature=float(temperature),
            max_tokens=int(max_tokens),
            timeout=int(timeout),
            api_key_env=api_key_env,
            api_base=api_base.strip() or None,
            summary_layout_index=int(summary_layout_index),
            developer_layout_index=int(developer_layout_index),
            log_level=log_level,
        )
        status = st.status("Generating report...", expanded=True)
        try:
            with status:
                progress_bar = st.progress(0.0, text="Starting...")

                def on_progress(fraction: float, label: str) -> None:
                    progress_bar.progress(fraction, text=label)

                payload = _generate(
                    config=config, tasks=tasks, use_llm=use_llm, progress=on_progress
                )
            st.session_state.report = payload
            status.update(label="Report ready", state="complete", expanded=False)
        except Exception as exc:
            LOGGER.exception("Report generation failed")
            status.update(label="Report generation failed", state="error")
            st.error(f"Report generation failed: {exc}")
            return

    # -------------------------------------------------------------- results
    if "report" in st.session_state and st.session_state.report:
        st.subheader("Results")
        _render_report(st.session_state.report)


if __name__ == "__main__":
    main()
