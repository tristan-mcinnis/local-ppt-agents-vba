# PowerPoint Automation System (Robust, Deterministic, Token‑Free)

A small, deterministic Python pipeline that turns an outline.json and a template_analysis.json into a ready‑to‑run VBA macro that builds your slides in the currently open PowerPoint presentation.

All code lives under `ppt_workflow/`. The system is template‑aware, cross‑platform (Windows/macOS), and designed to be robust and predictable for non‑technical users.

## Highlights

- Deterministic output: same input → same slides (no LLM calls)
- Robust VBA JSON parsing: handles whitespace, escapes, and nested arrays/objects
- Helpful error codes: grouped summary with machine‑readable codes at the end
- Template‑aware: uses your template’s layout indices and placeholder geometry
- Cross‑platform: generated VBA works on Windows and macOS
- Images optional: image placeholders are skipped; focus on text, charts, tables

## End‑to‑End Flow

```
outline.json + template_analysis.json → Python pipeline → output/generated_script.vba → run in PowerPoint
```

1) Analyze your template (one‑time per template)
- Open PowerPoint → Alt/Option+F11 → Insert Module → paste `ppt_workflow/vba/universal_template_analyzer.vba`
- Run `UniversalTemplateAnalyzer` (F5)
- It writes `template_analysis.json` next to your `.pptx`

2) Create your outline.json
- You specify slides, layouts (by name), and placeholder content
- Supported content types: Title, Subtitle, Body, Table, Chart
- Images can be provided but are ignored by design (add them manually later)

3) Generate the VBA script (Python)
- `cd ppt_workflow`
- `python workflow.py outline.json template_analysis.json`
- Output is `ppt_workflow/output/generated_script.vba`

4) Run the generated script in PowerPoint
- Open your template → Alt/Option+F11 → Insert Module
- Paste `generated_script.vba` → Run `Main`
- If issues occur, a summary dialog lists error codes and details

## Outline Format (Essentials)

Example with text, chart, and table:

```json
{
  "meta": { "title": "Q4 Results", "author": "Your Name" },
  "slides": [
    {
      "layout": "Title Slide",
      "placeholders": {
        "CenterTitle": "Q4 Financial Results",
        "Subtitle": "A strong finish"
      }
    },
    {
      "layout": "Title and Chart",
      "placeholders": {
        "Title": "North Star Metrics",
        "Chart": {
          "type": "line",
          "data": {
            "x": ["Week 1", "Week 2", "Week 3"],
            "series": [
              {"name": "WAU", "data": [200, 380, 520]}
            ]
          }
        }
      }
    },
    {
      "layout": "Title and Table",
      "placeholders": {
        "Title": "KPIs",
        "Table": {
          "headers": ["Metric", "Value"],
          "rows": [["Revenue", "25M"], ["Margin", "22%"]]
        }
      }
    }
  ]
}
```

Notes:
- Layout names must exist in `template_analysis.json` (case‑insensitive)
- When a layout has multiple placeholders of the same type, use ordinals: `Body[0]`, `Body[1]`, etc.
- Images (`Picture`, `SlideImage`) are accepted but skipped during generation

## Error Codes (VBA Summary Dialog)

The generated VBA logs all issues and shows them together at the end. These machine‑readable codes make it easy for an LLM or a human to pinpoint and fix problems.

- E1002 MissingPlaceholder — “Slide N: Missing placeholder Type=Body (type_id=2), Ordinal=1”
- E1003 ChartCreateFailed — Failed to create chart shape
- E1005 JsonKeyMissing — Missing required key in chart/table JSON (e.g., `data`, `x`, or `series`)
- E1007 ActivePresentationMissing — No active presentation is open
- E1008 LayoutResolveFailed — Layout index referenced by plan not found in the open template
- E1009 UnsupportedChartType — Chart type not recognized; defaults to column
- E1011 ChartSourceDataFailed — Could not bind worksheet range to chart

Behavior:
- Missing placeholders are logged per item; slide creation continues for other content
- Layout cache issues are logged; macro does not abort the whole run
- On completion, if any errors were logged, a summary dialog appears; otherwise a success message is shown

## Running the Pipeline

Quick setup (uses uv for fast, isolated env):

```bash
# From repo root
./setup.sh

# Run examples
uv run python ppt_workflow/workflow.py ppt_workflow/examples/simple_outline.json ppt_workflow/examples/template_analysis.json
```

```bash
# From repo root
cd ppt_workflow

# Run full workflow (validates inputs, generates plan and VBA)
python workflow.py outline.json template_analysis.json

# Outputs
# - ppt_workflow/output/slide_plan.json
# - ppt_workflow/output/generated_script.vba
```

Advanced: you can run steps individually for debugging:

```bash
# Step 1 only: Outline → Plan
python core/outline_to_plan.py outline.json template_analysis.json output/slide_plan.json

# Step 2 only: Plan → VBA
python core/plan_to_vba.py output/slide_plan.json output/generated_script.vba

# Validate artifacts (optional)
python utils/validator.py outline.json template_analysis.json output/slide_plan.json output/generated_script.vba
```

## Project Structure

```
local-ppt-agents-vba/
└── ppt_workflow/
    ├── core/
    │   ├── outline_to_plan.py      # Outline → Plan (skips images)
    │   └── plan_to_vba.py          # Plan → VBA (robust JSON parsing, error codes)
    ├── utils/
    │   └── validator.py            # Validation utilities
    ├── vba/
    │   └── universal_template_analyzer.vba  # Template analyzer macro
    ├── examples/                   # Example outline and analysis
    ├── data/                       # Template analysis samples
    ├── output/                     # Generated files (git-ignored)
    └── workflow.py                 # Orchestrator
```

## Testing

```bash
python -m unittest discover -s ppt_workflow/tests -p 'test_*.py'
```

The tests cover outline→plan basics, VBA generation presence checks, and robustness (image skipping and compact JSON emission).

## Troubleshooting

- Layout not found: verify the layout name exists in `template_analysis.json` (case‑insensitive). Consider using a different layout present in the template.
- Missing placeholder: confirm the placeholder type and ordinal exist in that layout.
- Chart/table issues: ensure JSON matches the documented shape. The error summary will indicate the missing key (E1005).

If you encounter an error, copy the final error summary (codes + details) and feed it to an LLM for a concise fix suggestion.
