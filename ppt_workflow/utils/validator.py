"""
Validation utilities for PowerPoint automation workflow
Validates outline, plan, and VBA at each stage
"""

import json
from pathlib import Path
from typing import Dict, List, Tuple, Any

from ppt_workflow.utils.path_utils import normalize_path


class WorkflowValidator:
    """Validates artifacts at each stage of the workflow"""

    def __init__(self):
        self.errors = []
        self.warnings = []
        self.info = []

    def clear(self):
        """Clear all messages"""
        self.errors = []
        self.warnings = []
        self.info = []

    def validate_outline(self, outline_path: str) -> Tuple[bool, Dict]:
        """Validate outline.json structure and content"""
        self.clear()

        try:
            norm = normalize_path(outline_path)
            with open(norm, 'r', encoding='utf-8', newline='') as f:
                outline = json.load(f)
        except Exception as e:
            self.errors.append(
                f"Failed to load outline: {e}. On macOS, verify the path and permissions."
            )
            return False, self._get_report()

        # Check required fields
        if "slides" not in outline:
            self.errors.append("Missing 'slides' field in outline")

        if not isinstance(outline.get("slides", []), list):
            self.errors.append("'slides' must be an array")

        if len(outline.get("slides", [])) == 0:
            self.errors.append("No slides defined in outline")

        # Validate each slide
        for i, slide in enumerate(outline.get("slides", []), 1):
            self._validate_slide(slide, i)

        # Check optional meta
        if "meta" in outline:
            if not isinstance(outline["meta"], dict):
                self.warnings.append("'meta' should be an object")

        return len(self.errors) == 0, self._get_report()

    def _validate_slide(self, slide: Dict, slide_num: int):
        """Validate individual slide structure"""
        # Check required fields
        if "layout" not in slide:
            self.errors.append(f"Slide {slide_num}: Missing 'layout' field")
        elif not isinstance(slide["layout"], str):
            self.errors.append(f"Slide {slide_num}: 'layout' must be a string")

        if "placeholders" not in slide:
            self.errors.append(f"Slide {slide_num}: Missing 'placeholders' field")
        elif not isinstance(slide["placeholders"], dict):
            self.errors.append(f"Slide {slide_num}: 'placeholders' must be an object")

        # Validate placeholder content
        for key, value in slide.get("placeholders", {}).items():
            self._validate_placeholder_content(key, value, slide_num)

    def _validate_placeholder_content(self, key: str, value: Any, slide_num: int):
        """Validate placeholder key and content"""
        # Parse key
        if "[" in key and key.endswith("]"):
            base = key[:key.index("[")]
            try:
                ordinal = int(key[key.index("[")+1:-1])
                if ordinal < 0:
                    self.errors.append(f"Slide {slide_num}: Invalid ordinal {ordinal} in '{key}'")
            except ValueError:
                self.errors.append(f"Slide {slide_num}: Invalid ordinal in '{key}'")
        else:
            base = key

        # Check known types
        known_types = ["title", "body", "subtitle", "chart", "table", "picture", "content"]
        if base.lower() not in known_types:
            self.warnings.append(f"Slide {slide_num}: Unknown placeholder type '{base}'")

        # Validate content based on type
        if base.lower() == "chart":
            if not isinstance(value, dict):
                self.errors.append(f"Slide {slide_num}: Chart content must be an object")
            elif "type" not in value or "data" not in value:
                self.errors.append(f"Slide {slide_num}: Chart must have 'type' and 'data'")

        elif base.lower() == "table":
            if not isinstance(value, dict):
                self.errors.append(f"Slide {slide_num}: Table content must be an object")
            elif "headers" not in value or "rows" not in value:
                self.errors.append(f"Slide {slide_num}: Table must have 'headers' and 'rows'")

        elif base.lower() in ["picture", "slideimage"]:
            if not isinstance(value, str):
                self.errors.append(f"Slide {slide_num}: Image placeholder must have string path")

    def validate_template_analysis(self, analysis_path: str) -> Tuple[bool, Dict]:
        """Validate template_analysis.json structure"""
        self.clear()

        try:
            norm = normalize_path(analysis_path)
            with open(norm, 'r', encoding='utf-8', newline='') as f:
                analysis = json.load(f)
        except Exception as e:
            self.errors.append(
                f"Failed to load analysis: {e}. On macOS, verify the path and permissions."
            )
            return False, self._get_report()

        # Check required top-level fields
        required = ["template_info", "layouts"]
        for field in required:
            if field not in analysis:
                self.errors.append(f"Missing required field '{field}'")

        # Validate template_info
        if "template_info" in analysis:
            info = analysis["template_info"]
            if "name" not in info:
                self.warnings.append("template_info missing 'name'")
            if "slide_master" not in info:
                self.warnings.append("template_info missing 'slide_master'")

        # Validate layouts
        if "layouts" in analysis:
            if not isinstance(analysis["layouts"], list):
                self.errors.append("'layouts' must be an array")
            else:
                for i, layout in enumerate(analysis["layouts"]):
                    self._validate_layout(layout, i)

        return len(self.errors) == 0, self._get_report()

    def _validate_layout(self, layout: Dict, idx: int):
        """Validate individual layout structure"""
        required = ["index", "name", "placeholders"]
        for field in required:
            if field not in layout:
                self.errors.append(f"Layout {idx}: Missing '{field}'")

        # Validate placeholders
        if "placeholders" in layout:
            if not isinstance(layout["placeholders"], list):
                self.errors.append(f"Layout {idx}: 'placeholders' must be an array")
            else:
                for ph in layout["placeholders"]:
                    if "type_id" not in ph:
                        self.warnings.append(f"Layout {idx}: Placeholder missing 'type_id'")
                    if "geometry" not in ph:
                        self.warnings.append(f"Layout {idx}: Placeholder missing 'geometry'")

    def validate_slide_plan(self, plan_path: str) -> Tuple[bool, Dict]:
        """Validate slide_plan.json structure"""
        self.clear()

        try:
            norm = normalize_path(plan_path)
            with open(norm, 'r', encoding='utf-8', newline='') as f:
                plan = json.load(f)
        except Exception as e:
            self.errors.append(
                f"Failed to load plan: {e}. On macOS, verify the path and permissions."
            )
            return False, self._get_report()

        # Check required fields
        required = ["meta", "slides"]
        for field in required:
            if field not in plan:
                self.errors.append(f"Missing required field '{field}'")

        # Validate slides
        if "slides" in plan:
            for slide in plan["slides"]:
                self._validate_plan_slide(slide)

        # Check validation section
        if "validation" in plan:
            if plan["validation"].get("errors", []):
                for err in plan["validation"]["errors"]:
                    self.errors.append(f"Plan error: {err}")

        return len(self.errors) == 0, self._get_report()

    def _validate_plan_slide(self, slide: Dict):
        """Validate slide in plan"""
        required = ["slide_number", "selected_layout", "content_map"]
        for field in required:
            if field not in slide:
                self.errors.append(f"Slide missing '{field}'")

        # Validate layout selection
        if "selected_layout" in slide:
            layout = slide["selected_layout"]
            if "index" not in layout:
                self.errors.append("selected_layout missing 'index'")
            elif not isinstance(layout["index"], int):
                self.errors.append("Layout index must be integer")

        # Validate content map
        if "content_map" in slide:
            for content in slide["content_map"]:
                if "type_id" not in content:
                    self.errors.append("Content missing 'type_id'")
                if "content_type" not in content:
                    self.errors.append("Content missing 'content_type'")
                if "content_data" not in content:
                    self.errors.append("Content missing 'content_data'")

    def validate_vba_script(self, vba_path: str) -> Tuple[bool, Dict]:
        """Validate generated VBA script"""
        self.clear()

        try:
            with open(vba_path, 'r', encoding='utf-8') as f:
                vba = f.read()
        except Exception as e:
            self.errors.append(f"Failed to load VBA: {e}")
            return False, self._get_report()

        # Critical checks
        if "Sub Main()" not in vba:
            self.errors.append("Missing 'Sub Main()' entry point")

        if "Application.ActivePresentation" not in vba:
            self.errors.append("Script doesn't use ActivePresentation")

        if "Application.Presentations.Add" in vba:
            self.errors.append("Script creates new presentation (should use active)")

        # Important functions
        required_functions = [
            "GetCustomLayoutByIndexSafe",
            "GetPlaceholderByTypeAndOrdinal",
            "SafeSetText"
        ]

        for func in required_functions:
            if func not in vba:
                self.warnings.append(f"Missing helper function: {func}")

        # Platform compatibility
        if "#If Mac Then" in vba:
            self.info.append("Script includes macOS compatibility")

        # Error handling
        if "On Error" in vba:
            self.info.append("Script includes error handling")

        return len(self.errors) == 0, self._get_report()

    def _get_report(self) -> Dict:
        """Get validation report"""
        return {
            "valid": len(self.errors) == 0,
            "errors": self.errors.copy(),
            "warnings": self.warnings.copy(),
            "info": self.info.copy(),
            "summary": {
                "error_count": len(self.errors),
                "warning_count": len(self.warnings),
                "info_count": len(self.info)
            }
        }


def validate_workflow(outline_path: str, analysis_path: str,
                      plan_path: str = None, vba_path: str = None) -> Dict:
    """Validate entire workflow"""
    validator = WorkflowValidator()
    results = {}

    # Validate outline
    valid, report = validator.validate_outline(outline_path)
    results["outline"] = report

    # Validate template analysis
    valid, report = validator.validate_template_analysis(analysis_path)
    results["analysis"] = report

    # Validate plan if provided
    if plan_path and Path(plan_path).exists():
        valid, report = validator.validate_slide_plan(plan_path)
        results["plan"] = report

    # Validate VBA if provided
    if vba_path and Path(vba_path).exists():
        valid, report = validator.validate_vba_script(vba_path)
        results["vba"] = report

    # Overall status
    all_valid = all(r["valid"] for r in results.values())
    total_errors = sum(r["summary"]["error_count"] for r in results.values())
    total_warnings = sum(r["summary"]["warning_count"] for r in results.values())

    results["overall"] = {
        "valid": all_valid,
        "total_errors": total_errors,
        "total_warnings": total_warnings,
        "stages_validated": list(results.keys())
    }

    return results


if __name__ == "__main__":
    import sys

    if len(sys.argv) < 3:
        print("Usage: python validator.py outline.json template_analysis.json [slide_plan.json] [script.vba]")
        sys.exit(1)

    outline = sys.argv[1]
    analysis = sys.argv[2]
    plan = sys.argv[3] if len(sys.argv) > 3 else None
    vba = sys.argv[4] if len(sys.argv) > 4 else None

    results = validate_workflow(outline, analysis, plan, vba)

    # Print results
    print("\n" + "=" * 60)
    print("VALIDATION REPORT")
    print("=" * 60)

    for stage, report in results.items():
        if stage == "overall":
            continue

        print(f"\n{stage.upper()}:")
        if report["valid"]:
            print("  ✓ Valid")
        else:
            print("  ✗ Invalid")

        if report["errors"]:
            print("  Errors:")
            for err in report["errors"]:
                print(f"    • {err}")

        if report["warnings"]:
            print("  Warnings:")
            for warn in report["warnings"]:
                print(f"    • {warn}")

    # Overall summary
    print("\n" + "-" * 60)
    overall = results["overall"]
    if overall["valid"]:
        print("✓ ALL VALIDATIONS PASSED")
    else:
        print(f"✗ VALIDATION FAILED: {overall['total_errors']} errors, {overall['total_warnings']} warnings")