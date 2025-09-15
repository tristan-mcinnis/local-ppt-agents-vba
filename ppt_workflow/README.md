# PowerPoint Automation Workflow (Python Pipeline)

A deterministic, token-free Python pipeline for generating PowerPoint VBA scripts from content outlines.

## Overview

This workflow replaces the complex sub-agent system with a simple, testable Python pipeline:

```
outline.json + template_analysis.json → slide_plan.json → generated_script.vba
```

### Key Benefits

- **Deterministic**: No LLM interpretation - strict mapping of layouts and placeholders
- **Zero Token Usage**: After initial outline creation, everything runs locally
- **Testable**: Pure functions with clear inputs/outputs
- **Fast**: No API calls, instant local execution
- **Debuggable**: Inspect intermediate files at each stage

## Installation

No special dependencies required - uses only Python standard library:

```bash
# Clone the repository
git clone <your-repo>
cd local-ppt-agents-vba/ppt_workflow

# Ensure Python 3.7+ is installed
python --version
# macOS users should use Homebrew python3
python3 --version
```

## Quick Start

### Step 0: Analyze Your PowerPoint Template

1. Open your PowerPoint template
2. Run the `universal_template_analyzer.vba` macro
3. This generates `template_analysis.json`

### Step 1: Create Your Outline

Create an `outline.json` file describing your slides:

```json
{
  "meta": {
    "title": "My Presentation",
    "author": "Your Name"
  },
  "slides": [
    {
      "layout": "Title Slide",
      "placeholders": {
        "Title": "Welcome to My Presentation",
        "Subtitle": "A Great Subtitle"
      }
    },
    {
      "layout": "Title and Content",
      "placeholders": {
        "Title": "Agenda",
        "Body": "• First topic\n• Second topic\n• Third topic"
      }
    }
  ]
}
```

### Step 2: Run the Workflow

```bash
# Simple one-command execution
python workflow.py outline.json template_analysis.json

# With options
python workflow.py outline.json template_analysis.json --skip-validation --quiet
```

### Step 3: Use the Generated VBA

1. Open your PowerPoint template
2. Press `Alt+F11` (Windows) or `Opt+F11` (Mac) for VBA editor
3. Insert → Module
4. Paste contents of `output/generated_script.vba`
5. Run `ValidateTemplate` to check compatibility
6. Run `Main` to create your slides

On macOS you can automate these steps with:

```bash
osascript ../scripts/run_vba.applescript MyTemplate.pptx output/generated_script.vba
```

## Detailed Usage

### Running Individual Steps

You can run each step separately for debugging:

```bash
# Step 1: Convert outline to plan
python core/outline_to_plan.py outline.json template_analysis.json output/slide_plan.json

# Step 2: Convert plan to VBA
python core/plan_to_vba.py output/slide_plan.json output/script.vba

# Validation
python utils/validator.py outline.json template_analysis.json output/slide_plan.json output/script.vba
```

### Outline Format

#### Placeholder Types

- **Text**: `"Title"`, `"Body"`, `"Subtitle"`
- **Images**: `"Picture"`, `"SlideImage"`
- **Charts**: `"Chart"` (requires chart object)
- **Tables**: `"Table"` (requires table object)

#### Using Ordinals

When a layout has multiple placeholders of the same type, use ordinals:

```json
{
  "layout": "Two Content",
  "placeholders": {
    "Title": "Comparison",
    "Body[0]": "Left side content",
    "Body[1]": "Right side content"
  }
}
```

#### Charts

```json
{
  "Chart": {
    "type": "column",
    "data": {
      "categories": ["Q1", "Q2", "Q3", "Q4"],
      "series": [
        {
          "name": "Sales",
          "values": [100, 150, 180, 220]
        }
      ]
    }
  }
}
```

#### Tables

```json
{
  "Table": {
    "headers": ["Product", "Q1", "Q2"],
    "rows": [
      ["Product A", "100", "120"],
      ["Product B", "80", "95"]
    ]
  }
}
```

## File Structure

```
ppt_workflow/
├── core/
│   ├── outline_to_plan.py    # Step 1: Outline → Plan converter
│   └── plan_to_vba.py        # Step 2: Plan → VBA generator
├── utils/
│   └── validator.py          # Validation utilities
├── output/                   # Generated files directory
│   ├── slide_plan.json       # Intermediate plan
│   └── generated_script.vba  # Final VBA script
├── workflow.py               # Main orchestrator
└── README.md                # This file
```

## Validation

The workflow includes comprehensive validation at each stage:

```bash
# Validate all stages
python utils/validator.py outline.json template_analysis.json slide_plan.json script.vba
```

Validation checks:
- Outline structure and content types
- Template analysis completeness
- Layout name matching
- Placeholder availability
- VBA script correctness

## Troubleshooting

### Common Issues

1. **"Layout not found"**
   - Check exact layout names in `template_analysis.json`
   - Layout names are case-insensitive but must match exactly

2. **"Placeholder not found"**
   - Verify placeholder type exists in the layout
   - Check ordinal if multiple placeholders of same type

3. **"Script doesn't create slides"**
   - Ensure you're running in an open presentation
   - Check that layout indices are correct
   - Verify `Main()` subroutine exists

### Debug Mode

For detailed output during processing:

```bash
# Run with verbose output
python workflow.py outline.json template_analysis.json

# Check intermediate files
cat output/slide_plan.json | python -m json.tool
```

## Advanced Features

### Custom Platform Targeting

The workflow supports platform-specific VBA generation:

```python
# In your Python code
converter = OutlineToPlanConverter(outline_path, analysis_path)
plan = converter.convert(platform_targets=["macos", "windows"])
```

### Extending the Pipeline

Add custom processing steps:

```python
from ppt_workflow.core.outline_to_plan import OutlineToPlanConverter

class CustomConverter(OutlineToPlanConverter):
    def _process_slide(self, slide, slide_no):
        # Your custom logic here
        result = super()._process_slide(slide, slide_no)
        # Additional processing
        return result
```

## Example Workflow

Complete example from start to finish:

```bash
# 1. Prepare your files
ls -la
# outline.json
# template_analysis.json

# 2. Run the workflow
python ppt_workflow/workflow.py outline.json template_analysis.json

# Output:
# [10:23:45] ℹ POWERPOINT AUTOMATION WORKFLOW
# [10:23:45] ✓ Input files validated successfully
# [10:23:45] ✓ Generated plan with 10 slides
# [10:23:45] ✓ Generated VBA for 10 slides
# [10:23:45] ✓ WORKFLOW COMPLETED SUCCESSFULLY

# 3. Check generated files
ls ppt_workflow/output/
# slide_plan.json
# generated_script.vba

# 4. Use in PowerPoint
# - Open PowerPoint template
# - Alt+F11 → Insert → Module
# - Paste generated_script.vba
# - Run Main()
```

## Contributing

To contribute to this workflow:

1. Keep functions pure and deterministic
2. Add comprehensive error messages
3. Include validation for new features
4. Test with various template types

## License

[Your License]

## Support

For issues or questions:
- Check the validation output first
- Review intermediate files (slide_plan.json)
- Ensure template_analysis.json is complete
- Verify outline.json follows the schema