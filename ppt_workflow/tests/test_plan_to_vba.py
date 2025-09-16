import json
import unittest
from pathlib import Path

from ppt_services.core.outline_to_plan import OutlinePlanConverter
from ppt_services.core.plan_to_ppt import PlanWriter  # or PlanToPPTConverter if that’s your class

class TestPptPlanConverter(unittest.TestCase):
    """
    Unit tests for generating a .json code plan from a slide plan and
    writing a PPT, including handling of missing assets.
    """

    @classmethod
    def setUpClass(cls):
        cls.base = Path(__file__).resolve().parent.parent  # adjust if needed
        cls.fixtures = cls.base / "fixtures"
        cls.outline = cls.fixtures / "simple_outline.json"
        cls.analysis = cls.fixtures / "slide_plan_analysis.json"
        cls.tmp_dir = cls.base / ".tmp"
        cls.tmp_dir.mkdir(exist_ok=True)

    def test_outline_to_plan_writes_code(self):
        """Converts outline to plan JSON and writes presentation code."""
        with open(self.outline, "r", encoding="utf-8") as f:
            outline = json.load(f)

        converter = OutlinePlanConverter(write_paths=True, outdir=str(self.tmp_dir))
        code = converter.convert(outline)

        self.assertIsInstance(code, dict)
        self.assertIn("slides", code)
        self.assertGreater(len(code["slides"]), 0)

        code_path = self.tmp_dir / "activePresentation.json"
        self.assertTrue(code_path.exists(), "Expected activePresentation.json to be written")

    def test_write_ppt_handles_missing_images_gracefully(self):
        """
        Ensures PPT writing reports missing images but does not crash,
        and retains slide placeholders.
        """
        plan_with_missing = {
            "title": "Image Test",
            "slides": [
                {
                    "layout": "Picture with Caption",
                    "placeholders": [
                        {"type": "Image Slot", "src": "fixtures/nonexistent.png"},
                        {"type": "Title", "text": "Caption"},
                    ],
                }
            ],
        }

        out_code = self.tmp_dir / "out_code.json"
        out_ppt = self.tmp_dir / "missing_images.pptx"

        with open(out_code, "w", encoding="utf-8") as f:
            json.dump(plan_with_missing, f, ensure_ascii=False, indent=2)

        writer = PlanWriter(allow_missing=True)  # or PlanToPPTConverter(...)
        result = writer.write_from_plan(plan_with_missing, output_path=str(out_ppt))

        # Adjust these assertions to your real return shape/fields.
        self.assertTrue(out_ppt.exists(), "PPTX should still be generated")
        self.assertIn("missing_assets", result)
        self.assertGreaterEqual(len(result["missing_assets"].get("images", [])), 1)

if __name__ == "__main__":
    unittest.main()
