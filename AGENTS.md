# Repository Guidelines

## Project Structure & Module Organization
- `ppt_workflow/core/` — Python converters: `outline_to_plan.py`, `plan_to_vba.py`.
- `ppt_workflow/utils/` — Validators and small helpers.
- `ppt_workflow/vba/` — VBA utilities (e.g., `universal_template_analyzer.vba`).
- `ppt_workflow/examples/` — Sample `outline.json` and `template_analysis.json`.
- `ppt_workflow/data/` — Example templates/analyses for local use.
- `ppt_workflow/output/` — Generated artifacts (git-ignored).
- Tests live in `ppt_workflow/tests/` and follow `test_*.py` naming.

## Build, Test, and Development Commands
- Run end-to-end: `cd ppt_workflow && python workflow.py outline.json template_analysis.json`
  - Produces `output/slide_plan.json` and `output/generated_script.vba`.
- Step 1 only: `python core/outline_to_plan.py <outline.json> <template_analysis.json> <slide_plan.json>`
- Step 2 only: `python core/plan_to_vba.py <slide_plan.json> <script.vba>`
- Tests: `python -m unittest discover -s ppt_workflow/tests -p 'test_*.py'`

## Coding Style & Naming Conventions
- Python 3.10+, 4-space indentation, UTF-8, standard library only (no new deps).
- Use type hints and concise docstrings. Prefer pure, deterministic functions.
- Naming: modules/files `lower_snake_case.py`, functions/variables `lower_snake_case`,
  classes `PascalCase`, constants `UPPER_SNAKE_CASE`.
- Keep imports at top; avoid side effects at import time. Do not change VBA
  helper names or error codes without updating tests and README.

## Testing Guidelines
- Framework: `unittest`. Place tests under `ppt_workflow/tests/` named `test_*.py`.
- Cover: outline→plan mapping, plan→VBA emission, robustness (image skipping, compact JSON).
- Use `tempfile` for artifacts and `examples/` inputs for fixtures. Ensure tests run offline.

## Commit & Pull Request Guidelines
- Commits: imperative mood and scoped prefix, e.g., `core: fix ordinal parsing`.
- PRs must include: clear description, rationale, linked issue (if any),
  before/after behavior or sample commands, and tests for behavior changes.
- Keep changes minimal and focused; avoid unrelated refactors.

## Security & Configuration Tips
- Do not commit templates with sensitive content. `output/` is ignored; keep it that way.
- Generated VBA must remain macOS/Windows safe; use provided helpers (e.g., `SafeSetText`,
  `GetCustomLayoutByIndexSafe`) and maintain deterministic behavior (images are intentionally skipped).

