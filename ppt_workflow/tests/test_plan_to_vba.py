import json
import tempfile
import unittest
from pathlib import Path

from ppt_workflow.core.outline_to_plan import OutlineToPlanConverter
from ppt_workflow.core.plan_to_vba import PlanToVBAConverter


class TestPlanToVBAConverter(unittest.TestCase):
    """Tests for generating VBA code from a slide plan."""

    def setUp(self):
        base = Path(__file__).resolve().parent.parent
        outline = base / "examples" / "simple_outline.json"
        analysis = base / "examples" / "template_analysis.json"
        converter = OutlineToPlanConverter(str(outline), str(analysis))
        self.plan = converter.convert()

    def test_vba_generation_contains_main(self):
        """Generated VBA should include entry point and active presentation reference."""
        with tempfile.NamedTemporaryFile("w", delete=False, suffix=".json") as tmp:
            json.dump(self.plan, tmp)
            tmp_path = tmp.name
        try:
            vba_converter = PlanToVBAConverter(tmp_path)
            code = vba_converter.convert()
            self.assertIn("Sub Main()", code)
            self.assertIn("Application.ActivePresentation", code)
        finally:
            Path(tmp_path).unlink(missing_ok=True)


if __name__ == "__main__":
    unittest.main()
