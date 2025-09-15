"""
Step 1: Convert outline.json to slide_plan.json
Deterministic mapping of content to template layouts with exact indices and placeholders
"""

import json
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Any, Optional, Tuple

from ppt_workflow.utils.path_utils import normalize_path


class OutlineToPlanConverter:
    """Converts user outline to strict slide plan using template analysis"""

    # PowerPoint placeholder type constants
    # Keys are normalized (lowercase) for matching, values are (type_id, canonical_name)
    TYPE_MAP = {
        "title": (1, "Title"),
        "body": (2, "Body"),
        "centertitle": (3, "CenterTitle"),
        "subtitle": (4, "Subtitle"),
        "object": (7, "Object"),
        "chart": (8, "Chart"),
        "table": (9, "Table"),
        "slideimage": (13, "SlideImage"),
        "picture": (18, "Picture"),
        "content": (19, "Content"),
    }

    def __init__(self, outline_path: str, analysis_path: str):
        """Initialize with paths to outline and template analysis"""
        self.outline = self._load_json(outline_path)
        self.analysis = self._load_json(analysis_path)
        self.layout_index = self._build_layout_index()
        self.errors = []
        self.warnings = []

    @staticmethod
    def _load_json(path: str) -> Dict:
        """Load and parse JSON file with macOS-friendly error reporting"""
        norm_path = normalize_path(path)
        try:
            with open(norm_path, 'r', encoding='utf-8', newline='') as f:
                return json.load(f)
        except FileNotFoundError as e:
            raise FileNotFoundError(
                f"File not found: {norm_path}. On macOS, check path spelling and "
                "permissions."
            ) from e

    @staticmethod
    def _normalize_name(s: str) -> str:
        """Normalize names for case-insensitive comparison"""
        return s.strip().lower() if s else ""

    def _build_layout_index(self) -> Dict[str, Dict]:
        """Build searchable index of layouts by normalized name"""
        index = {}
        for layout in self.analysis.get("layouts", []):
            name = self._normalize_name(layout["name"])

            # Pre-sort placeholders by type, then by position (top, left)
            ph_by_type = {}
            for ph in layout.get("placeholders", []):
                type_id = ph.get("type_id", 0)
                if type_id not in ph_by_type:
                    ph_by_type[type_id] = []
                ph_by_type[type_id].append(ph)

            # Sort each type's placeholders by position
            for type_id, phs in ph_by_type.items():
                phs.sort(key=lambda p: (
                    p.get("geometry", {}).get("top", 0),
                    p.get("geometry", {}).get("left", 0)
                ))

            index[name] = {
                "index": layout["index"],
                "name": layout["name"],
                "placeholders": layout.get("placeholders", []),
                "ph_by_type": ph_by_type,
                "category": layout.get("category", "content")
            }

        return index

    def _parse_placeholder_key(self, key: str, slide_no: int) -> Tuple[str, int]:
        """
        Parse placeholder key to extract type and ordinal.
        Examples: "Body" -> ("body", 0), "Body[1]" -> ("body", 1)

        On invalid ordinal, logs an error with slide number and
        returns a safe default ordinal of 0.
        """
        key = key.strip()
        if "[" in key and key.endswith("]"):
            base = key[:key.index("[")]
            ordinal_str = key[key.index("[")+1:-1]
            try:
                ordinal = int(ordinal_str)
            except ValueError:
                normalized = self._normalize_name(base)
                self.errors.append(
                    f"Slide {slide_no}: Placeholder '{key}' has invalid ordinal '{ordinal_str}'. Defaulting to 0."
                )
                return normalized, 0
            return self._normalize_name(base), ordinal
        return self._normalize_name(key), 0

    def _validate_placeholder(self, layout: Dict, canonical_name: str, type_id: int,
                            ordinal: int, slide_no: int, key: str) -> bool:
        """Validate that a placeholder exists in the layout"""
        placeholders = layout["ph_by_type"].get(type_id, [])

        if ordinal < 0 or ordinal >= len(placeholders):
            error = (f"Slide {slide_no}: Placeholder '{key}' (type_id={type_id}, "
                    f"ordinal={ordinal}) not found in layout '{layout['name']}'. "
                    f"Available: {len(placeholders)} of type {canonical_name}")
            self.errors.append(error)
            return False
        return True

    def _determine_content_type(self, value: Any, placeholder_type: str) -> Tuple[str, Any]:
        """Determine content type and validate/transform content data"""
        if placeholder_type in ("picture", "slideimage"):
            if not isinstance(value, str) or not value.strip():
                raise ValueError(f"Expected image path for {placeholder_type}, got: {value}")
            return "image_path", value.strip()

        elif placeholder_type == "chart":
            if not isinstance(value, dict):
                raise ValueError(f"Chart content must be an object, got: {type(value)}")
            # Validate chart structure
            if "type" not in value:
                raise ValueError("Chart must have 'type' field")
            if "data" not in value:
                raise ValueError("Chart must have 'data' field")
            return "chart", value

        elif placeholder_type == "table":
            if not isinstance(value, dict):
                raise ValueError(f"Table content must be an object, got: {type(value)}")
            if "headers" not in value or "rows" not in value:
                raise ValueError("Table must have 'headers' and 'rows' fields")
            return "table", value

        else:
            # Text content (including bullets)
            if not isinstance(value, str):
                value = str(value)  # Coerce to string
            return "text", value

    def _extract_slide_title(self, placeholders: Dict, layout_name: str,
                            slide_no: int) -> str:
        """Extract title from slide placeholders or use fallback"""
        # Look for explicit Title placeholder
        for key, value in placeholders.items():
            base, _ = self._parse_placeholder_key(key, slide_no)
            if base == "title" and isinstance(value, str) and value.strip():
                return value

        # For title slide, use meta title as fallback
        if slide_no == 1 and "title" in self._normalize_name(layout_name):
            meta_title = self.outline.get("meta", {}).get("title", "")
            if meta_title:
                return meta_title

        return f"Slide {slide_no}"

    def _find_layout_index(self, layout_name: str,
                           fallback_category: Optional[str] = None) -> int:
        """Lookup layout index by name with optional category fallback.

        Args:
            layout_name: The expected layout name to search for.
            fallback_category: Optional layout category to use if the name
                is not found. The first layout matching this category will be
                returned.

        Returns:
            The index of the matched layout.

        Raises:
            ValueError: If neither the layout name nor the fallback category
                can be found in the template analysis.
        """
        normalized = self._normalize_name(layout_name)
        if normalized in self.layout_index:
            return self.layout_index[normalized]["index"]

        if fallback_category:
            fallback_norm = self._normalize_name(fallback_category)
            for layout in self.analysis.get("layouts", []):
                if self._normalize_name(layout.get("category", "")) == fallback_norm:
                    self.warnings.append(
                        f"Layout '{layout_name}' not found. Using '{layout['name']}' "
                        f"(index {layout['index']}) as fallback."
                    )
                    return layout["index"]

        raise ValueError(
            f"Required layout '{layout_name}' not found and no fallback "
            f"available."
        )

    def _process_slide(self, slide: Dict, slide_no: int) -> Dict:
        """Process a single slide from outline to plan format"""
        layout_name = slide.get("layout", "")
        normalized_name = self._normalize_name(layout_name)

        # Find matching layout
        if normalized_name not in self.layout_index:
            available = list(self.layout_index.keys())[:10]
            raise ValueError(f"Slide {slide_no}: Layout '{layout_name}' not found. "
                           f"Available layouts: {available}...")

        layout = self.layout_index[normalized_name]
        content_map = []
        placeholders_expected = []
        referenced = set()

        # Process each placeholder in the slide
        for key, value in slide.get("placeholders", {}).items():
            base, ordinal = self._parse_placeholder_key(key, slide_no)

            # Get type ID and canonical name
            if base not in self.TYPE_MAP:
                self.warnings.append(f"Slide {slide_no}: Unknown placeholder type '{base}'")
                continue

            type_id, canonical_name = self.TYPE_MAP[base]

            # Validate placeholder exists
            if not self._validate_placeholder(layout, canonical_name, type_id, ordinal,
                                             slide_no, key):
                continue

            # Determine content type and data
            try:
                content_type, content_data = self._determine_content_type(value, base)
            except ValueError as e:
                self.errors.append(f"Slide {slide_no}, placeholder '{key}': {e}")
                continue

            # Add to content map with canonical name
            content_map.append({
                "placeholder_type": canonical_name,
                "type_id": type_id,
                "ordinal": ordinal,
                "content_type": content_type,
                "content_data": content_data
            })

            referenced.add((canonical_name, type_id, ordinal))

        # Build placeholders_expected list
        for ph_type, type_id, ordinal in sorted(referenced, key=lambda x: (x[1], x[2])):
            placeholders_expected.append({
                "type": ph_type,
                "type_id": type_id,
                "ordinal": ordinal
            })

        # Extract slide title
        slide_title = self._extract_slide_title(
            slide.get("placeholders", {}),
            layout["name"],
            slide_no
        )

        return {
            "slide_number": slide_no,
            "slide_title": slide_title,
            "selected_layout": {
                "name": layout["name"],
                "index": layout["index"],
                "reason": "exact_name_match"
            },
            "addressing": "by_type_then_ordinal",
            "fill_policy": "strict_match",
            "placeholders_expected": placeholders_expected,
            "platform_hints": {
                "mac_safe": True,
                "chart_api": "AddChart",
                "text_api": "TextFrame2_with_fallback"
            },
            "content_map": content_map
        }

    def convert(self, platform_targets: Optional[List[str]] = None) -> Dict:
        """Convert outline to slide plan"""
        if platform_targets is None:
            platform_targets = ["macos", "windows"]

        slides = []
        layout_usage = {}

        # Process each slide
        for i, slide in enumerate(self.outline.get("slides", []), start=1):
            try:
                slide_plan = self._process_slide(slide, i)
                slides.append(slide_plan)

                # Track layout usage
                layout_name = slide_plan["selected_layout"]["name"]
                layout_usage[layout_name] = layout_usage.get(layout_name, 0) + 1

            except ValueError as e:
                self.errors.append(str(e))
                # Collect errors but continue processing remaining slides

        if self.errors:
            combined = "\n".join(self.errors)
            raise ValueError(f"Conversion encountered errors:\n{combined}")

        # Build layout strategy by resolving layout names to indices
        layout_names = {
            "title_slide_index": ("Title Slide", "title"),
            "two_column_index": ("title-two-text", "content"),
            "three_column_index": ("title-three-text", "content"),
            "chart_layout_index": ("Title, Text and Chart", "content"),
            "standard_content_index": ("Title and Text", "content"),
            "contact_index": ("contact-slide-white", "content"),
        }

        layout_strategy = {}
        for key, (name, category) in layout_names.items():
            layout_strategy[key] = self._find_layout_index(name, category)

        # Build complete plan
        plan = {
            "meta": {
                "template_name": self.analysis["template_info"]["name"],
                "analysis_date": self.analysis["template_info"]["analysis_date"],
                "total_layouts": len(self.analysis.get("layouts", [])),
                "platform_targets": platform_targets,
                "planner_version": "4.0-python",
                "created_at": datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ"),
            },
            "layout_strategy": layout_strategy,
            "slides": slides,
            "validation": {
                "checks": [
                    "indices_in_range",
                    "placeholders_compatible",
                    "platform_hints_applied",
                    "content_fidelity_exact"
                ],
                "errors": self.errors,
                "warnings": self.warnings
            },
            "layout_usage_summary": layout_usage
        }

        return plan


def main():
    """CLI entry point"""
    import sys

    if len(sys.argv) != 4:
        print("Usage: python outline_to_plan.py outline.json template_analysis.json slide_plan.json")
        sys.exit(1)

    outline_path = sys.argv[1]
    analysis_path = sys.argv[2]
    output_path = sys.argv[3]

    # Convert
    converter = OutlineToPlanConverter(outline_path, analysis_path)
    plan = converter.convert()

    # Write output
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(plan, f, indent=2, ensure_ascii=False)

    # Report results
    print(f"✓ Generated slide_plan.json with {len(plan['slides'])} slides")
    if converter.errors:
        print(f"⚠ Errors: {len(converter.errors)}")
        for err in converter.errors:
            print(f"  - {err}")
    if converter.warnings:
        print(f"⚠ Warnings: {len(converter.warnings)}")
        for warn in converter.warnings:
            print(f"  - {warn}")


if __name__ == "__main__":
    main()