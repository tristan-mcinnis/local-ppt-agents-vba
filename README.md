# PowerPoint Automation System

A deterministic Python pipeline that automatically generates PowerPoint VBA scripts from content outlines, with zero AI token usage after initial setup.

📁 **All code is in the `ppt_workflow/` directory**

The system analyzes your PowerPoint template structure and combines it with your content outline to produce a ready-to-run VBA script that adds slides to your active presentation.

## Features

-   **Deterministic:** Same input always produces the same output - no AI interpretation
-   **Zero Token Usage:** After initial outline creation, everything runs locally with Python
-   **macOS Compatible:** Generated VBA works on both Windows and Mac PowerPoint
-   **Template-Aware:** Tailored to your specific PowerPoint template structure
-   **Fast & Testable:** Pure Python functions with clear inputs and outputs

## How It Works

Simple three-step pipeline with deterministic Python processing:

```
outline.json + template_analysis.json → Python Pipeline → generated_script.vba
```

1. **Analyze Template**: Run VBA analyzer on your PowerPoint template
2. **Create Outline**: Write your content in JSON format
3. **Generate Script**: Python pipeline produces VBA script
4. **Run in PowerPoint**: Execute script to add slides to your presentation

## Usage Instructions

Follow these three steps to generate your presentation.

### Step 1: Analyze Your PowerPoint Template

This step only needs to be done once per template.

1.  Open the PowerPoint presentation or template (`.pptx` or `.potx`) you want to use.
2.  Open the VBA Editor (**Alt+F11**).
3.  Create a new Module and paste the code from the `ppt_workflow/vba/universal_template_analyzer.vba` file into it.
4.  Run the `UniversalTemplateAnalyzer` macro (**F5**).

This will create a file named `template_analysis.json` in the same directory as your PowerPoint file. This file contains all the structural information about your template.

### Step 2: Create Your Content Outline

Create an `outline.json` file that describes your slides:

**Example `outline.json`:**

```json
{
  "meta": {
    "title": "Q4 Financial Results",
    "author": "Your Name"
  },
  "slides": [
    {
      "layout": "Title Slide",
      "placeholders": {
        "Title": "Q4 Financial Results",
        "Subtitle": "A Review of a Strong Quarter"
      }
    },
    {
      "layout": "Title and Content",
      "placeholders": {
        "Title": "Key Metrics",
        "Body": "• Record revenue of $25M, up 15% YoY\n• Net profit margin increased to 22%\n• Customer acquisition grew by 18,000"
      }
    }
  ]
}
```

### Step 3: Run the Python Pipeline

```bash
# Navigate to the project directory
cd ppt_workflow

# Run the workflow
python workflow.py outline.json template_analysis.json

# Generated VBA script will be in output/generated_script.vba
```

### Step 4: Use the Generated VBA

1.  Open your PowerPoint template
2.  Press **Alt+F11** to open VBA editor
3.  Insert → Module
4.  Paste the contents of `output/generated_script.vba`
5.  Run the `Main()` subroutine (**F5**)
6.  Your slides will be added to the active presentation

## macOS Setup and Tips

- Use Homebrew `python3` (`brew install python`) and run the workflow with `python3`.
- Install dependencies with `pip3 install -r requirements-macos.txt`.
- Open the VBA editor with **Opt+F11** instead of Alt+F11.
- If macros are disabled, enable them via **PowerPoint → Preferences → Security & Privacy → Macro Security**.
- You can run macros from the terminal:

```bash
osascript -e 'tell application "Microsoft PowerPoint" to run macro "Main"'
```

- For fully automated execution on macOS, use the provided AppleScript:

```bash
osascript scripts/run_vba.applescript MyTemplate.pptx ppt_workflow/output/generated_script.vba
```
This opens your template, loads the generated VBA, and runs the `Main` macro.

- Paths on macOS can be case-sensitive; ensure your `outline.json` and `template_analysis.json` names match exactly.
- Run `scripts/mac_env_check.py` to verify your Python environment.

### macOS Compatibility

The workflow has been tested with:

| macOS Version | PowerPoint Version |
|---------------|-------------------|
| macOS 13+     | PowerPoint 16.77+ |
| macOS 14+     | PowerPoint 16.80+ |

Other versions may work but are not officially tested.

## Project Structure

```
local-ppt-agents-vba/
└── ppt_workflow/                    # Complete PowerPoint automation system
    ├── core/
    │   ├── outline_to_plan.py      # Converts outline to structured plan
    │   └── plan_to_vba.py          # Generates VBA from plan
    ├── utils/
    │   └── validator.py            # Validation utilities
    ├── vba/
    │   └── universal_template_analyzer.vba  # Template analyzer VBA script
    ├── examples/                   # Example files
    │   ├── demo_outline.json       # Full demo outline
    │   ├── simple_outline.json     # Simple example
    │   └── template_analysis.json  # Example template analysis
    ├── data/                       # Template analysis files
    │   └── ic-template-1_analysis.json
    ├── output/                     # Generated files (git-ignored)
    └── workflow.py                 # Main orchestrator
