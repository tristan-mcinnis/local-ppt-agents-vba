# CLAUDE.md: PowerPoint Automation System - Python Pipeline

This document describes the deterministic Python pipeline for PowerPoint automation. The system transforms content outlines into VBA scripts without any AI interpretation or token usage after initial setup.

## 1. System Overview

The PowerPoint Automation System is a **deterministic Python pipeline** that:
- Takes structured JSON input (outline + template analysis)
- Produces ready-to-run VBA scripts
- Requires zero AI tokens after initial outline creation
- Works on both Windows and macOS PowerPoint

**Key Principle: Same input → Same output, every time**

## 2. Pipeline Architecture

```
outline.json + template_analysis.json → slide_plan.json → generated_script.vba
```

### Core Components

- **`universal_template_analyzer.vba`**: VBA script that analyzes PowerPoint templates
- **`outline_to_plan.py`**: Converts user outline to structured plan
- **`plan_to_vba.py`**: Generates VBA script from plan
- **`workflow.py`**: Main orchestrator

### Data Flow

1. **Template Analysis** (One-time per template)
   - User runs `universal_template_analyzer.vba` in PowerPoint
   - Produces `template_analysis.json`

2. **Outline Creation** (Per presentation)
   - User creates `outline.json` with slide content
   - Structured JSON format with layouts and placeholders

3. **Python Processing** (Deterministic)
   - `outline_to_plan.py`: Maps content to template layouts
   - `plan_to_vba.py`: Generates VBA script
   - No AI calls, pure Python logic

4. **VBA Execution**
   - User runs generated script in PowerPoint
   - Slides added to active presentation

## 3. Key Features

### Deterministic Processing
- No LLM interpretation
- Strict layout/placeholder mapping
- Predictable, testable output

### macOS Compatibility
- Uses Collection instead of Scripting.Dictionary
- No ActiveX dependencies
- Works on Mac PowerPoint

### No Image Handling
- Images deliberately excluded from automation
- Users add images manually after slide creation
- Prevents file path and compatibility issues

## 4. VBA Script Requirements

Generated scripts MUST:

1. **Use Active Presentation**
   ```vba
   Set pres = Application.ActivePresentation
   ```

2. **Have Main() Entry Point**
   ```vba
   Sub Main()  ' User runs with F5
   ```

3. **Use Collection-based Cache (macOS)**
   ```vba
   Dim layoutCache As Collection
   Set layoutCache = New Collection
   ```

4. **Skip Image Placeholders**
   ```vba
   ' Image placeholder skipped
   ' User will add image manually
   ```

## 5. Outline Format

### Basic Structure
```json
{
  "meta": {
    "title": "Presentation Title",
    "author": "Author Name"
  },
  "slides": [
    {
      "layout": "Title Slide",
      "placeholders": {
        "Title": "Welcome",
        "Subtitle": "Subtitle text"
      }
    }
  ]
}
```

### Placeholder Types
- **Text**: Title, Body, Subtitle
- **Charts**: Column, Bar, Pie (with data object)
- **Tables**: Headers and rows arrays
- **Images**: Skipped (manual addition)

### Using Ordinals
For multiple placeholders of same type:
```json
{
  "Body[0]": "Left content",
  "Body[1]": "Right content"
}
```

## 6. Python Pipeline Details

### outline_to_plan.py
- Validates outline structure
- Matches layouts to template
- Resolves placeholder mappings
- Produces `slide_plan.json`

### plan_to_vba.py
- Generates VBA from plan
- Creates macOS-compatible code
- Skips image handling
- Includes validation checks

### Key Mappings
```python
TYPE_MAP = {
    "title": (1, "Title"),
    "body": (2, "Body"),
    "centertitle": (3, "CenterTitle"),
    "subtitle": (4, "Subtitle"),
    # ... more mappings
}
```

## 7. Usage Workflow

```bash
# 1. Analyze template (one-time)
# Run universal_template_analyzer.vba in PowerPoint

# 2. Create outline
# Write outline.json

# 3. Generate VBA
cd ppt_workflow
python workflow.py outline.json template_analysis.json

# 4. Run in PowerPoint
# Paste output/generated_script.vba and run Main()
```

## 8. Troubleshooting

### Common Issues

| Issue | Solution |
|-------|----------|
| Layout not found | Check exact names in template_analysis.json |
| Placeholder missing | Verify type exists in layout |
| macOS error 429 | Ensure using Collection, not Dictionary |
| Images not appearing | Expected - add manually |

### Validation

Run validator to check all stages:
```bash
python utils/validator.py outline.json template_analysis.json slide_plan.json script.vba
```

## 9. Design Decisions

### Why Python Pipeline?
- Deterministic output
- Zero token usage
- Fast local execution
- Easy debugging
- Testable functions

### Why No Images?
- File path complexity
- Cross-platform issues
- Security concerns
- User preference for manual control

### Why Collection vs Dictionary?
- macOS VBA doesn't support CreateObject("Scripting.Dictionary")
- Collection works on all platforms
- Minimal performance difference

## 10. Version History

- **v5.0**: Complete Python pipeline replacement
  - Removed all AI agent dependencies
  - Added macOS compatibility
  - Removed image handling
  - Deterministic processing

- **v4.0**: Last agent-based version (deprecated)