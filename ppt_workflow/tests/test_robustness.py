import json
import tempfile
import unittest
from pathlib import Path

from ppt_workflow.core.outline_to_plan import OutlineToPlanConverter
from ppt_workflow.core.plan_to_vba import PlanToVBAConverter


class TestRobustness(unittest.TestCase):
    def setUp(self):
        base = Path(__file__).resolve().parent.parent
        self.analysis = base / "examples" / "template_analysis.json"

    def test_compact_json_and_skip_images(self):
        # Build a minimal outline including chart, table, and image placeholders
        outline = {
            "meta": {"title": "Robustness Test"},
            "slides": [
                {
                    "layout": "Title and Chart",
                    "placeholders": {
                        "Title": "Metrics",
                        "Chart": {
                            "type": "line",
                            "data": {
                                "x": ["A", "B", "C"],
                                "series": [{"name": "S1", "data": [1, 2, 3]}]
                            }
                        },
                        "Picture": "ignored.png"
                    }
                },
                {
                    "layout": "Title and Table",
                    "placeholders": {
                        "Title": "Table",
                        "Table": {"headers": ["H1", "H2"], "rows": [["1", "2"]]}
                    }
                }
            ]
        }

        with tempfile.NamedTemporaryFile("w", delete=False, suffix=".json") as outline_file:
            json.dump(outline, outline_file)
            outline_path = outline_file.name

        try:
            # Convert to plan
            conv = OutlineToPlanConverter(str(outline_path), str(self.analysis))
            plan = conv.convert()

            # Ensure images were skipped in planning (no image content types)
            for slide in plan["slides"]:
                for item in slide["content_map"]:
                    self.assertNotEqual(item.get("content_type"), "image_path")

            # Generate VBA
            with tempfile.NamedTemporaryFile("w", delete=False, suffix=".json") as tmp_plan:
                json.dump(plan, tmp_plan)
                tmp_plan_path = tmp_plan.name

            try:
                vba = PlanToVBAConverter(tmp_plan_path).convert()

                # Confirm we did NOT emit any image placeholder handling comments
                self.assertNotIn("Image placeholder skipped", vba)

                # Confirm compact JSON (no spaces after colons) for chart and table specs
                # Note: within VBA string literals, quotes are doubled
                self.assertIn('""type"":', vba)
                self.assertIn('""data"":', vba)
                self.assertIn('""headers"":', vba)
                self.assertIn('""rows"":', vba)

                # Ensure no occurrences of pattern with space after colon for these keys
                self.assertNotIn('""type"": ', vba)
                self.assertNotIn('""data"": ', vba)
                self.assertNotIn('""headers"": ', vba)
                self.assertNotIn('""rows"": ', vba)
            finally:
                Path(tmp_plan_path).unlink(missing_ok=True)
        finally:
            Path(outline_path).unlink(missing_ok=True)


if __name__ == "__main__":
    unittest.main()
