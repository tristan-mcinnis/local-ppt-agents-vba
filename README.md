# PowerPoint Automation with Claude Code

Create and edit PowerPoint presentations directly through Claude Code using `python-pptx` and office skills workflows.

## What's New

This repository has been updated to use the [claude-office-skills](https://github.com/tfriedel/claude-office-skills) system, which provides:

- **Direct interaction** — Ask Claude Code to create/edit presentations naturally
- **Template preservation** — Work with your existing corporate templates
- **python-pptx based** — No more VBA scripts or copy-paste workflows
- **Multi-format support** — Also supports Word, Excel, and PDF

## Quick Start

```bash
# 1. Install Python dependencies
python -m venv venv
venv/bin/pip install -r requirements.txt

# 2. Install Node.js dependencies (for HTML-to-PPTX)
npm install
```

## Usage with Claude Code

Simply ask Claude Code what you want:

```
> Create a 5-slide presentation about Q4 results using templates/ic-template-1.pptx
> Edit the title slide to say "Annual Review 2024"
> Add a comparison slide with two columns
```

Claude Code will:
1. Analyze your template (thumbnails, text inventory)
2. Select appropriate slides from the template
3. Generate replacement content
4. Apply changes and save to `outputs/`

## Supported Workflows

### From Template (Recommended)
Use an existing `.pptx` template — preserves branding, fonts, colors, and layouts.

```bash
# Generate thumbnails to understand template
venv/bin/python skills/public/pptx/scripts/thumbnail.py templates/my-template.pptx outputs/project/

# Rearrange/duplicate slides
venv/bin/python skills/public/pptx/scripts/rearrange.py template.pptx output.pptx 0,5,5,12,3

# Extract text positions
venv/bin/python skills/public/pptx/scripts/inventory.py presentation.pptx inventory.json

# Replace text from JSON
venv/bin/python skills/public/pptx/scripts/replace.py input.pptx replacements.json output.pptx
```

### From Scratch
Create presentations using HTML/CSS design with the html2pptx workflow.

See `skills/public/pptx/html2pptx.md` for details.

### Edit Existing
Modify presentations using OOXML (raw XML) editing.

See `skills/public/pptx/ooxml.md` for details.

## Project Structure

```
├── skills/                    # Claude office skills
│   └── public/
│       ├── pptx/             # PowerPoint workflows
│       │   ├── SKILL.md      # Main workflow doc
│       │   ├── scripts/      # Python utilities
│       │   ├── html2pptx.md  # Create from HTML
│       │   └── ooxml.md      # XML editing
│       ├── docx/             # Word documents
│       ├── xlsx/             # Excel spreadsheets
│       └── pdf/              # PDF manipulation
├── templates/                 # Your PowerPoint templates
├── outputs/                   # Generated files (gitignored)
├── requirements.txt          # Python dependencies
└── package.json              # Node.js dependencies
```

## Key Features

| Feature | Description |
|---------|-------------|
| Template-based | Preserves your corporate branding |
| Cross-platform | Works on Windows, macOS, Linux |
| Visual validation | Thumbnail grids catch layout issues |
| Text replacement | JSON-based content replacement |
| Slide rearrangement | Duplicate, reorder, delete slides |
| HTML-to-PPTX | Design slides in HTML/CSS |
| OOXML editing | Direct XML manipulation |

## Dependencies

### Python
- python-pptx, openpyxl, pypdf
- defusedxml, lxml, Pillow
- markitdown (text extraction)

### Node.js
- pptxgenjs, playwright, sharp

### System Tools
- LibreOffice (`soffice`) — PDF conversion
- Poppler (`pdftoppm`) — PDF to images
- Pandoc — document extraction

## Output Directory

All generated files go to `outputs/<project-name>/`:

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

## Documentation

- **Getting started**: See `.claude/CLAUDE.md`
- **Development**: See `AGENTS.md`
- **Skills system**: See `skills/CLAUDE.md`
- **PowerPoint workflows**: See `skills/public/pptx/SKILL.md`

## Attribution

Office skills from [tfriedel/claude-office-skills](https://github.com/tfriedel/claude-office-skills).
