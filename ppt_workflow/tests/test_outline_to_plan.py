import json
import unittest
from pathlib import Path

from ppt_workflow.core.outline_to_plan import OutlineToPlanConverter


class TestOutlineToPlanConverter(unittest.TestCase):
    """Basic tests for OutlineToPlanConverter."""

    def setUp(self):
        base = Path(__file__).resolve().parent.parent
        self.outline = base / "examples" / "simple_outline.json"
        self.analysis = base / "examples" / "template_analysis.json"

    def test_convert_produces_plan(self):
        """Converting an outline should yield a slide plan with expected slides."""
        converter = OutlineToPlanConverter(str(self.outline), str(self.analysis))
        plan = converter.convert()

        self.assertEqual(len(plan["slides"]), 3)
        self.assertEqual(plan["slides"][0]["selected_layout"]["name"], "Title Slide")


if __name__ == "__main__":
    unittest.main()
