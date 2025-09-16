#!/usr/bin/env bash
set -euo pipefail

# Quick project setup using Astral's uv (fast Python package/deps manager)
# - Installs uv if missing (Homebrew on macOS if available, else installer script)
# - Creates a local virtual environment at .venv
# - Installs dependencies if a requirements file is found

echo "==> Checking for 'uv'..."
if ! command -v uv >/dev/null 2>&1; then
  echo "'uv' not found. Attempting installation..."
  if [[ "$(uname -s)" == "Darwin" ]] && command -v brew >/dev/null 2>&1; then
    echo "Using Homebrew to install uv"
    brew install uv
  else
    echo "Using official installer script"
    # Installer places binaries under ~/.local/bin by default
    curl -LsSf https://astral.sh/uv/install.sh | sh
    export PATH="$HOME/.local/bin:$PATH"
  fi
fi

export UV_CACHE_DIR="${UV_CACHE_DIR:-$PWD/.uv-cache}"
echo "==> Creating virtual environment (.venv)"
uv venv

echo "==> Installing dependencies (if any)"
if [[ -f requirements.txt ]]; then
  uv pip install -r requirements.txt
elif [[ -f requirements-macos.txt ]]; then
  uv pip install -r requirements-macos.txt
else
  echo "No requirements file found. Skipping installs."
fi

cat <<'EOF'

Setup complete.

Common commands:
- Run tests:
    uv run python -m unittest discover -s ppt_workflow/tests -p 'test_*.py'
- Generate VBA from examples:
    uv run python ppt_workflow/workflow.py ppt_workflow/examples/simple_outline.json ppt_workflow/examples/template_analysis.json

Note: uv created an isolated env in .venv. Activate if you prefer:
    source .venv/bin/activate
EOF
