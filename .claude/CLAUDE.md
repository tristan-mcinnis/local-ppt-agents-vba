# CLAUDE.md: PowerPoint Automation with Claude Code

This repository enables direct PowerPoint creation and editing through Claude Code using the `python-pptx` library and office skills workflows.

## Quick Start

```bash
# Install Python dependencies
python -m venv venv
venv/bin/pip install -r requirements.txt

# Install Node.js dependencies (for HTML-to-PPTX)
npm install
```

## How It Works

Simply ask Claude Code to create or edit PowerPoint files:

```
> Create a quarterly sales presentation using my template at templates/ic-template-1.pptx
> Add 3 slides to my presentation about product features
> Edit the title slide to say "Q4 Results"
```

Claude Code will:
1. Check the skills system for the appropriate workflow
2. Read the `skills/public/pptx/SKILL.md` workflow
3. Execute each step (thumbnails, inventory, replacements, etc.)
4. Save outputs to `outputs/<document-name>/`

## Repository Structure

```
├── skills/                    # Claude office skills (cloned repo)
│   └── public/
│       ├── pptx/             # PowerPoint workflows
│       │   ├── SKILL.md      # Main PPTX workflow
│       │   ├── html2pptx.md  # HTML-to-PPTX guide
│       │   ├── ooxml.md      # OOXML editing guide
│       │   └── scripts/      # Python utilities
│       ├── docx/             # Word workflows
│       ├── xlsx/             # Excel workflows
│       └── pdf/              # PDF workflows
├── templates/                 # Your PowerPoint templates
├── outputs/                   # Generated files (gitignored)
└── venv/                      # Python virtual environment
```

## Key Commands

### PowerPoint Operations

**Create thumbnails for template analysis:**
```bash
venv/bin/python skills/public/pptx/scripts/thumbnail.py templates/my-template.pptx outputs/project/thumbnails
```

**Extract text inventory:**
```bash
venv/bin/python skills/public/pptx/scripts/inventory.py presentation.pptx outputs/project/inventory.json
```

**Rearrange/duplicate slides:**
```bash
venv/bin/python skills/public/pptx/scripts/rearrange.py template.pptx outputs/project/working.pptx 0,5,5,12,3
```

**Replace text from JSON:**
```bash
venv/bin/python skills/public/pptx/scripts/replace.py input.pptx outputs/project/replacements.json outputs/project/final.pptx
```

**Extract text content:**
```bash
venv/bin/python -m markitdown presentation.pptx
```

## Workflow: Creating from Template

1. **Generate thumbnails** to understand template structure
2. **Create template inventory** listing all slides with indices
3. **Plan slide selection** based on content needs
4. **Rearrange slides** using `rearrange.py`
5. **Extract text inventory** from working presentation
6. **Generate replacement JSON** with new content
7. **Apply replacements** using `replace.py`

## Output Directory Convention

All generated files go to `outputs/<document-name>/`:
```
outputs/
├── quarterly-report/
│   ├── thumbnails.jpg
│   ├── inventory.json
│   ├── replacements.json
│   └── final.pptx
└── sales-deck/
    └── ...
```

## Skills System

Before writing custom code, Claude Code checks if a skill exists for the task. Skills are documented in `skills/public/*/SKILL.md` files.

**IMPORTANT**: Always read the full SKILL.md file before starting work. These files contain critical workflow steps and validation requirements.

## Dependencies

### Python (requirements.txt)
- python-pptx, openpyxl, pypdf
- defusedxml, lxml
- Pillow, pdf2image
- markitdown

### Node.js (package.json)
- pptxgenjs (HTML-to-PPTX)
- playwright (HTML rendering)
- sharp (image processing)

### System Tools
- LibreOffice (`soffice`) - PDF conversion
- Poppler (`pdftoppm`) - PDF to images
- Pandoc - document text extraction
