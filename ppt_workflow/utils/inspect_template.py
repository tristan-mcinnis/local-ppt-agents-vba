"""
Quick inspection tool for template_analysis.json.
Prints layouts with indices and placeholder type summaries to help choose layouts.

Usage:
  python -m ppt_workflow.utils.inspect_template <path/to/template_analysis.json>
"""

import json
import sys
from collections import Counter
from pathlib import Path


def main() -> int:
    if len(sys.argv) != 2:
        print("Usage: python -m ppt_workflow.utils.inspect_template <template_analysis.json>")
        return 1

    path = Path(sys.argv[1])
    if not path.exists():
        print(f"File not found: {path}")
        return 1

    try:
        with open(path, "r", encoding="utf-8") as f:
            analysis = json.load(f)
    except Exception as e:
        print(f"Failed to load JSON: {e}")
        return 1

    layouts = analysis.get("layouts", [])
    if not layouts:
        print("No layouts found in analysis")
        return 1

    print(f"Template: {analysis.get('template_info', {}).get('name', 'unknown')}")
    print(f"Total layouts: {len(layouts)}\n")

    for layout in layouts:
        idx = layout.get("index")
        name = layout.get("name")
        category = layout.get("category", "content")
        placeholders = layout.get("placeholders", [])
        types = Counter(ph.get("type_id") for ph in placeholders)
        summary = ", ".join(f"type_id {k}: {v}" for k, v in sorted(types.items()))
        print(f"[{idx:>2}] {name}  (category: {category})")
        print(f"     placeholders: {len(placeholders)}  |  {summary}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

