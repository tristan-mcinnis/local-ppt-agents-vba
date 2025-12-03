# Repository Guidelines

## Project Structure

```
├── skills/                    # Claude office skills (from tfriedel/claude-office-skills)
│   ├── CLAUDE.md             # Skills system documentation
│   ├── public/pptx/          # PowerPoint workflows and scripts
│   ├── public/docx/          # Word document workflows
│   ├── public/xlsx/          # Excel workflows
│   └── public/pdf/           # PDF workflows
├── templates/                 # User PowerPoint templates
├── outputs/                   # Generated files (gitignored)
├── .claude/CLAUDE.md         # Claude Code instructions
├── requirements.txt          # Python dependencies
└── package.json              # Node.js dependencies (from skills/)
```

## Development Setup

```bash
# Python virtual environment
python -m venv venv
venv/bin/pip install -r requirements.txt

# Node.js (for HTML-to-PPTX workflows)
npm install
```

## Key Workflows

### PowerPoint from Template
1. Generate thumbnails: `venv/bin/python skills/public/pptx/scripts/thumbnail.py`
2. Rearrange slides: `venv/bin/python skills/public/pptx/scripts/rearrange.py`
3. Extract inventory: `venv/bin/python skills/public/pptx/scripts/inventory.py`
4. Apply replacements: `venv/bin/python skills/public/pptx/scripts/replace.py`

### PowerPoint from Scratch
See `skills/public/pptx/html2pptx.md` for HTML-to-PPTX conversion workflow.

## Coding Style

- Python 3.10+, 4-space indentation, UTF-8
- Use type hints and concise docstrings
- **CRITICAL**: Always use `venv/bin/python` for all Python commands
- All generated files go to `outputs/<document-name>/`

## Skills System

Before writing custom code:
1. Check if a skill exists in `skills/public/*/SKILL.md`
2. Read the SKILL.md file completely (no range limits)
3. Follow the documented workflow exactly

## Commit Guidelines

- Use imperative mood with scope: `docs: update README`, `skills: add template`
- Keep changes focused and minimal
- Do not commit `outputs/` or `venv/` directories

## Security

- Do not commit templates with sensitive content
- `outputs/` directory is gitignored
- Validate all OOXML edits before packing
