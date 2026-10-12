# Sprint Performance Reporting

Generate a branded PowerPoint sprint report from an Excel task export. The application
validates and cleans the workbook, calculates team metrics, creates a task-count
burndown and daily logged-effort chart for every developer, and requests a concise
executive summary from the configured LiteLLM model.

## Requirements and setup

Use Python 3.10 or newer. Install dependencies in a virtual environment:

```bash
python -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
```

Copy `config.example.yaml` to `config.yaml` and update the paths, sprint dates/name,
and PowerPoint layout indexes. Place both the Excel workbook and the corporate `.pptx`
master inside the `templates/` directory (the defaults are
`templates/sprint_tasks.xlsx` and `templates/master_template.pptx`). Required workbook
headers (case-insensitive) are `Developer`, `Task`, `Sprint`, `Status`,
`Assigned Date`, `Start Date`, `End Date`, and `Time Spent`.

## Directory layout

```
templates/   Input workbook and PowerPoint master template
uploads/     Files uploaded through the Streamlit UI are saved here
output/      Generated PowerPoint reports
```

## LLM providers

LiteLLM uses standard model identifiers. Examples:

```yaml
llm:
  model: "openai/gpt-4o-mini"
  api_key_env: "OPENAI_API_KEY"
```

```yaml
llm:
  model: "anthropic/claude-3-5-haiku-latest"
  api_key_env: "ANTHROPIC_API_KEY"
```

For local Ollama, start the Ollama service and use:

```yaml
llm:
  model: "ollama/llama3.1"
  api_base: "http://localhost:11434"
```

Set cloud credentials in the environment (never in YAML), e.g.
`export OPENAI_API_KEY="..."`. If the LLM call fails, a warning is logged and the
report contains a factual deterministic summary instead. `LLM_MODEL`, `LLM_API_BASE`,
`LLM_API_KEY_ENV`, `SPRINT_INPUT_FILE`, `SPRINT_TEMPLATE_FILE`, `SPRINT_OUTPUT_FILE`,
`SPRINT_START`, `SPRINT_END`, `SPRINT_NAME`, and `LOG_LEVEL` can override YAML values.

## What the LLM generates

The model is asked for strict JSON and adds three things on top of the deterministic
metrics:

- **Per-developer brief** — a two-sentence `summary`, a one-line `risk`, and 1-3
  `actions`. Rendered on each developer slide and in the UI tabs.
- **Team narrative** — a 2-3 sentence executive summary for the summary slide.
- **Cross-developer risk flags** — short, name-attributed risk items for the summary
  slide.

Each call falls back to a deterministic equivalent if the provider is unreachable or
returns malformed JSON, so a report is always produced.

## Run

```bash
python sprint_report.py --config config.yaml
```

The team summary and generated developer slides appear before any existing slides in
the template. Existing template slides and the template master/layouts are retained.
Set `summary_layout_index` and `developer_layout_index` to layouts with the desired
corporate branding. Invalid indexes fall back to a blank layout if available, then
the first layout. Output folders are created automatically; chart images are temporary
and are removed after the presentation is saved.

## Streamlit UI

For a browser-based workflow, launch the Streamlit app instead of the CLI:

```bash
streamlit run streamlit_app.py
```

The UI is a thin layer over the same engine: upload an `.xlsx` workbook and `.pptx`
master (or use the bundled sample data in `templates/`), set the sprint window, choose
an LLM provider/model, then generate and download the deck. Uploaded files are saved
to the `uploads/` directory, and generated decks are written to `output/`. It also
previews the team KPIs, the per-developer burndown charts, and the executive summaries
in the browser. API keys typed into the sidebar are held only in the running process
and are never written to disk. If a provider is unreachable or summaries are disabled,
the report is produced with deterministic summaries and a warning is logged.

## Data handling

Time logs may be JSON objects such as
`{"2026-06-01": "4h", "2026-06-02": "2h"}` or text such as
`2026-06-01: 4h, 2026-06-02: 2h`. Durations support hours and minutes. Invalid
individual log fragments and dates are warned about; valid rows and other developers
continue to be processed.

The input schema does not contain estimates or remaining-effort values. Therefore,
burndown lines are measured in remaining tasks (one unit per task); actual task burn
is based on completion status and end date. Logged hours are shown separately as
daily effort and are not treated as estimates or remaining work.
