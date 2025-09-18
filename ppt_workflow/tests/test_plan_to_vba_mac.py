import json
import tempfile
import unittest
from pathlib import Path

from ppt_workflow.core.plan_to_vba import PlanToVBAConverter


class TestPlanToVBAMacFallbacks(unittest.TestCase):
    """Validate macOS-specific chart fallbacks and warning handling."""

    def _write_plan(self, plan: dict) -> str:
        tmp = tempfile.NamedTemporaryFile("w", delete=False, suffix=".json")
        json.dump(plan, tmp)
        tmp.close()
        self.addCleanup(lambda: Path(tmp.name).unlink(missing_ok=True))
        return tmp.name

    def _base_plan(self) -> dict:
        return {
            "meta": {
                "template_name": "mac-template.pptx"
            },
            "slides": [
                {
                    "slide_number": 1,
                    "slide_title": "Chart Slide",
                    "selected_layout": {"index": 5, "name": "Title and Content"},
                    "addressing": "by_type_then_ordinal",
                    "fill_policy": "strict_match",
                    "placeholders_expected": [
                        {"type": "Title", "type_id": 1, "ordinal": 0},
                        {"type": "Chart", "type_id": 8, "ordinal": 0}
                    ],
                    "platform_hints": {"chart_api": "AddChart"},
                    "content_map": [
                        {
                            "placeholder_type": "Title",
                            "type_id": 1,
                            "ordinal": 0,
                            "content_type": "text",
                            "content_data": "Chart slide"
                        },
                        {
                            "placeholder_type": "Chart",
                            "type_id": 8,
                            "ordinal": 0,
                            "content_type": "chart",
                            "content_data": {
                                "type": "column",
                                "data": {
                                    "categories": ["2023", "2024"],
                                    "series": [
                                        {"name": "North", "values": [10, 20]},
                                        {"name": "South", "values": [15, 25]}
                                    ]
                                }
                            }
                        }
                    ]
                }
            ]
        }

    def _render(self, plan: dict) -> str:
        path = self._write_plan(plan)
        converter = PlanToVBAConverter(path)
        return converter.convert()

    def test_chart_payload_injects_api_hint(self):
        code = self._render(self._base_plan())
        self.assertIn('""_chart_api"":""addchart""', code)
        self.assertIn('CreateChartAtPlaceholder currentSlide, shp', code)

    def test_chart_helper_omits_placeholder_insertchart(self):
        code = self._render(self._base_plan())
        self.assertNotIn('PlaceholderFormat.InsertChart', code)
        self.assertIn('EnsureSlideActive sld', code)

    def test_chart_helper_provides_table_warning_fallback(self):
        code = self._render(self._base_plan())
        self.assertIn('LogWarning "W1003"', code)
        self.assertIn('sld.Shapes.AddTable', code)

    def test_main_reports_warnings(self):
        plan = self._base_plan()
        code = self._render(plan)
        self.assertIn('ElseIf WarningsCount() > 0 Then', code)
        self.assertIn('ShowWarnings', code)

    def test_warning_infrastructure_present(self):
        code = self._render(self._base_plan())
        self.assertIn('Dim warningLog As Collection', code)
        self.assertIn('Set warningLog = New Collection', code)
        self.assertIn('Function WarningsCount()', code)

    def test_chart_helper_sets_shape_bounds(self):
        code = self._render(self._base_plan())
        self.assertIn('chartShape.Left = l', code)
        self.assertIn('chartShape.Top = t', code)
        self.assertIn('chartShape.Width = w', code)
        self.assertIn('chartShape.Height = h', code)


if __name__ == "__main__":
    unittest.main()
