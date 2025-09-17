"""
Comprehensive unit tests for plan_to_vba module
Tests VBA generation, macOS compatibility, and edge cases
"""

import json
import pytest
from pathlib import Path
import sys
import re

# Add parent directory to path for imports
sys.path.insert(0, str(Path(__file__).parent.parent))
from core.plan_to_vba import PlanToVBAConverter


class TestVBAGenerationBasics:
    """Test basic VBA generation functionality"""

    @pytest.fixture
    def minimal_plan(self, tmp_path):
        """Create minimal plan for testing"""
        plan = {
            "meta": {
                "title": "Test Presentation",
                "author": "Test Suite",
                "template_name": "test_template.pptx",
                "created_at": "2024-01-01T12:00:00Z"
            },
            "slides": [
                {
                    "slide_number": 1,
                    "layout": "Title Slide",
                    "selected_layout": {
                        "index": 1,
                        "name": "Title Slide"
                    },
                    "placeholders": [
                        {
                            "placeholder_key": "Title",
                            "type_id": 1,
                            "ordinal": 0,
                            "content": "Test Title"
                        }
                    ]
                }
            ],
            "layout_usage_summary": {"Title Slide": 1}
        }

        plan_path = tmp_path / "plan.json"
        with open(plan_path, 'w') as f:
            json.dump(plan, f)

        return str(plan_path)

    def test_converter_initialization(self, minimal_plan):
        """Test converter initializes correctly"""
        converter = PlanToVBAConverter(minimal_plan)

        assert converter.plan is not None
        assert converter.code_lines == []
        assert converter.used_layouts == set()
        assert converter.debug_slide is None

    def test_basic_vba_generation(self, minimal_plan):
        """Test basic VBA script generation"""
        converter = PlanToVBAConverter(minimal_plan)
        vba = converter.convert()

        assert vba is not None
        assert "Sub Main()" in vba
        assert "End Sub" in vba
        assert "Option Explicit" in vba

    def test_vba_escape_special_characters(self):
        """Test VBA string escaping"""
        converter = PlanToVBAConverter.__new__(PlanToVBAConverter)

        # Test quote escaping
        assert converter._vba_escape('Test "quoted"') == 'Test ""quoted""'

        # Test newline handling
        assert converter._vba_escape('Line1\nLine2') == 'Line1" & vbCrLf & "Line2'

        # Test empty string
        assert converter._vba_escape('') == ''
        assert converter._vba_escape(None) == ''


class TestMacOSCompatibility:
    """Test macOS-specific VBA generation"""

    @pytest.fixture
    def vba_converter(self, tmp_path):
        """Create converter with standard plan"""
        plan = {
            "meta": {
                "title": "Mac Test",
                "template_name": "mac_template.pptx"
            },
            "slides": []
        }

        plan_path = tmp_path / "plan.json"
        with open(plan_path, 'w') as f:
            json.dump(plan, f)

        return PlanToVBAConverter(str(plan_path))

    def test_macos_platform_detection(self, vba_converter):
        """Test that macOS platform detection code is included"""
        vba = vba_converter.convert()

        assert "#If Mac Then" in vba
        assert 'Const PLATFORM As String = "macOS"' in vba
        assert 'Const PLATFORM As String = "Windows"' in vba
        assert "#End If" in vba

    def test_collection_instead_of_dictionary(self, vba_converter):
        """Test that Collection is used instead of Scripting.Dictionary"""
        vba = vba_converter.convert()

        # Should use Collection
        assert "Dim layoutCache As Collection" in vba
        assert "Set layoutCache = New Collection" in vba

        # Should NOT use Dictionary
        assert "Scripting.Dictionary" not in vba
        assert "CreateObject" not in vba

    def test_cache_functions_for_collection(self, vba_converter):
        """Test that cache helper functions are macOS-compatible"""
        vba = vba_converter.convert()

        # Check cache functions exist
        assert "Function CacheHas" in vba
        assert "Sub CachePut" in vba
        assert "Function CacheGet" in vba

        # Check they use Collection methods
        assert "layoutCache.Add" in vba
        assert "layoutCache.Remove" in vba
        assert 'layoutCache(CStr(' in vba


class TestPlaceholderHandling:
    """Test VBA generation for different placeholder types"""

    @pytest.fixture
    def plan_with_placeholders(self, tmp_path):
        """Create plan with various placeholder types"""
        plan = {
            "meta": {"title": "Test", "template_name": "test.pptx"},
            "slides": [
                {
                    "slide_number": 1,
                    "layout": "Complex",
                    "selected_layout": {"index": 1, "name": "Complex"},
                    "placeholders": [
                        {
                            "placeholder_key": "Title",
                            "type_id": 1,
                            "ordinal": 0,
                            "content": "Title Text"
                        },
                        {
                            "placeholder_key": "Body",
                            "type_id": 2,
                            "ordinal": 0,
                            "content": "Body Text"
                        },
                        {
                            "placeholder_key": "Body",
                            "type_id": 2,
                            "ordinal": 1,
                            "content": "Second Body"
                        },
                        {
                            "placeholder_key": "Chart",
                            "type_id": 8,
                            "ordinal": 0,
                            "content": {
                                "type": "column",
                                "data": {
                                    "x": ["A", "B"],
                                    "series": [{"name": "S1", "data": [1, 2]}]
                                }
                            }
                        },
                        {
                            "placeholder_key": "Table",
                            "type_id": 9,
                            "ordinal": 0,
                            "content": {
                                "headers": ["H1", "H2"],
                                "rows": [["R1C1", "R1C2"]]
                            }
                        }
                    ]
                }
            ]
        }

        plan_path = tmp_path / "plan.json"
        with open(plan_path, 'w') as f:
            json.dump(plan, f)

        return str(plan_path)

    def test_text_placeholder_generation(self, plan_with_placeholders):
        """Test VBA generation for text placeholders"""
        converter = PlanToVBAConverter(plan_with_placeholders)
        vba = converter.convert()

        # Check text placeholder handling
        assert "GetPlaceholderByTypeAndOrdinal" in vba
        assert "SafeSetText" in vba
        assert '"Title Text"' in vba
        assert '"Body Text"' in vba
        assert '"Second Body"' in vba

    def test_chart_placeholder_generation(self, plan_with_placeholders):
        """Test VBA generation for chart placeholders"""
        converter = PlanToVBAConverter(plan_with_placeholders)
        vba = converter.convert()

        # Check chart creation function
        assert "CreateChartAtPlaceholder" in vba
        assert "xlColumnClustered" in vba
        assert "AddChart" in vba

    def test_table_placeholder_generation(self, plan_with_placeholders):
        """Test VBA generation for table placeholders"""
        converter = PlanToVBAConverter(plan_with_placeholders)
        vba = converter.convert()

        # Check table creation function
        assert "CreateTableAtPlaceholder" in vba
        assert "AddTable" in vba
        assert ".Table" in vba

    def test_multiple_ordinals_handling(self, plan_with_placeholders):
        """Test handling of multiple placeholders of same type"""
        converter = PlanToVBAConverter(plan_with_placeholders)
        vba = converter.convert()

        # Should generate calls with different ordinals
        assert "GetPlaceholderByTypeAndOrdinal(currentSlide, 2, 0)" in vba  # First Body
        assert "GetPlaceholderByTypeAndOrdinal(currentSlide, 2, 1)" in vba  # Second Body


class TestErrorHandling:
    """Test error handling and validation in VBA generation"""

    @pytest.fixture
    def error_plan(self, tmp_path):
        """Create plan that will trigger various errors"""
        plan = {
            "meta": {"title": "Error Test", "template_name": "test.pptx"},
            "slides": [
                {
                    "slide_number": 1,
                    "layout": "Test",
                    "selected_layout": {"index": 999, "name": "Nonexistent"},
                    "placeholders": [
                        {
                            "placeholder_key": "InvalidType",
                            "type_id": 999,
                            "ordinal": 0,
                            "content": "Test"
                        }
                    ]
                }
            ]
        }

        plan_path = tmp_path / "plan.json"
        with open(plan_path, 'w') as f:
            json.dump(plan, f)

        return str(plan_path)

    def test_error_logging_functions(self, error_plan):
        """Test that error logging functions are included"""
        converter = PlanToVBAConverter(error_plan)
        vba = converter.convert()

        # Check error handling functions
        assert "Sub InitErrorLog()" in vba
        assert "Sub LogError" in vba
        assert "Function ErrorsCount" in vba
        assert "Sub ShowErrors" in vba

    def test_missing_placeholder_handling(self, error_plan):
        """Test handling of missing placeholders"""
        converter = PlanToVBAConverter(error_plan)
        vba = converter.convert()

        # Should include error handling for missing placeholders
        assert "If shp Is Nothing Then" in vba
        assert 'LogError "E1002"' in vba

    def test_validation_subroutine(self, error_plan):
        """Test that validation subroutine is generated"""
        converter = PlanToVBAConverter(error_plan)
        vba = converter.convert()

        # Check validation sub exists
        assert "Sub ValidateTemplate()" in vba
        assert "Template Validation Report" in vba


class TestJSONParsingInVBA:
    """Test JSON parsing functions in generated VBA"""

    @pytest.fixture
    def vba_with_json(self, tmp_path):
        """Generate VBA with JSON parsing requirements"""
        plan = {
            "meta": {"title": "JSON Test", "template_name": "test.pptx"},
            "slides": [
                {
                    "slide_number": 1,
                    "layout": "Chart",
                    "selected_layout": {"index": 1, "name": "Chart"},
                    "placeholders": [
                        {
                            "placeholder_key": "Chart",
                            "type_id": 8,
                            "ordinal": 0,
                            "content": {
                                "type": "line",
                                "data": {
                                    "x": ["Jan", "Feb", "Mar"],
                                    "series": [
                                        {"name": "Sales", "data": [100, 120, 140]}
                                    ]
                                }
                            }
                        }
                    ]
                }
            ]
        }

        plan_path = tmp_path / "plan.json"
        with open(plan_path, 'w') as f:
            json.dump(plan, f)

        converter = PlanToVBAConverter(str(plan_path))
        return converter.convert()

    def test_json_parsing_functions_exist(self, vba_with_json):
        """Test that all JSON parsing functions are included"""
        # Core JSON functions
        assert "Function JsonValue" in vba_with_json
        assert "Function ParseJsonArray" in vba_with_json
        assert "Sub ParseSeries" in vba_with_json

        # Helper functions
        assert "Private Sub Json_SkipWs" in vba_with_json
        assert "Private Function Json_ParseString" in vba_with_json
        assert "Private Function Json_FindMatching" in vba_with_json

    def test_json_escape_handling(self, vba_with_json):
        """Test JSON string escape handling in VBA"""
        # Check for escape sequence handling
        assert r'Case "\"": res = res & "\"' in vba_with_json or \
               r'Case """": res = res & """' in vba_with_json
        assert 'Case "n": res = res & vbLf' in vba_with_json
        assert 'Case "r": res = res & vbCr' in vba_with_json
        assert 'Case "t": res = res & vbTab' in vba_with_json


class TestDebugFeatures:
    """Test debug features in VBA generation"""

    @pytest.fixture
    def plan_for_debug(self, tmp_path):
        """Create plan for debug testing"""
        plan = {
            "meta": {"title": "Debug Test", "template_name": "test.pptx"},
            "slides": [
                {
                    "slide_number": 1,
                    "layout": "Test",
                    "selected_layout": {"index": 1, "name": "Test"},
                    "placeholders": []
                }
            ]
        }

        plan_path = tmp_path / "plan.json"
        with open(plan_path, 'w') as f:
            json.dump(plan, f)

        return str(plan_path)

    def test_debug_mode_generation(self, plan_for_debug):
        """Test debug mode features"""
        converter = PlanToVBAConverter(plan_for_debug, debug_slide=1)
        vba = converter.convert()

        # Should include debug helper
        assert "Sub DebugListPlaceholders" in vba

    def test_debug_placeholder_listing(self, plan_for_debug):
        """Test debug placeholder listing function"""
        converter = PlanToVBAConverter(plan_for_debug, debug_slide=1)
        vba = converter.convert()

        # Check debug output formatting
        assert "Debug.Print" in vba
        assert "=== Placeholders on slide" in vba


class TestLayoutCaching:
    """Test layout caching optimization in VBA"""

    @pytest.fixture
    def multi_layout_plan(self, tmp_path):
        """Create plan with multiple layouts"""
        plan = {
            "meta": {"title": "Cache Test", "template_name": "test.pptx"},
            "slides": [
                {
                    "slide_number": i,
                    "layout": f"Layout{i % 3}",
                    "selected_layout": {"index": i % 3 + 1, "name": f"Layout{i % 3}"},
                    "placeholders": []
                }
                for i in range(1, 10)
            ]
        }

        plan_path = tmp_path / "plan.json"
        with open(plan_path, 'w') as f:
            json.dump(plan, f)

        return str(plan_path)

    def test_layout_cache_usage(self, multi_layout_plan):
        """Test that layout caching is implemented"""
        converter = PlanToVBAConverter(multi_layout_plan)
        vba = converter.convert()

        # Check cache pre-loading
        assert "Pre-cache layouts for performance" in vba
        assert "requiredLayouts = Array(" in vba

    def test_unique_layouts_tracked(self, multi_layout_plan):
        """Test that unique layouts are tracked"""
        converter = PlanToVBAConverter(multi_layout_plan)
        vba = converter.convert()

        assert len(converter.used_layouts) == 3  # 3 unique layouts


class TestCompleteVBAStructure:
    """Test complete VBA script structure and components"""

    @pytest.fixture
    def complete_plan(self, tmp_path):
        """Create a complete plan with all features"""
        fixtures_dir = Path(__file__).parent / "fixtures"

        # First convert outline to plan
        from core.outline_to_plan import OutlineToPlanConverter
        converter = OutlineToPlanConverter(
            str(fixtures_dir / "complex_outline.json"),
            str(fixtures_dir / "sample_template_analysis.json")
        )
        plan = converter.convert()

        plan_path = tmp_path / "complete_plan.json"
        with open(plan_path, 'w') as f:
            json.dump(plan, f)

        return str(plan_path)

    def test_complete_vba_structure(self, complete_plan):
        """Test that complete VBA has all required sections"""
        converter = PlanToVBAConverter(complete_plan)
        vba = converter.convert()

        # Check major sections
        assert "HELPER FUNCTIONS" in vba
        assert "MAIN SUBROUTINE" in vba
        assert "VALIDATION SUBROUTINE" in vba

        # Check required subs/functions
        required_elements = [
            "Sub Main()",
            "Sub ValidateTemplate()",
            "Function GetCustomLayoutByIndexSafe",
            "Function AddSlideWithLayout",
            "Function GetPlaceholderByTypeAndOrdinal",
            "Sub SafeSetText",
            "Function SortShapesByPosition"
        ]

        for element in required_elements:
            assert element in vba, f"Missing: {element}"

    def test_vba_constants_defined(self, complete_plan):
        """Test that all required constants are defined"""
        converter = PlanToVBAConverter(complete_plan)
        vba = converter.convert()

        # PowerPoint constants
        assert "Const msoPlaceholder" in vba
        assert "Const ppPlaceholderTitle" in vba
        assert "Const ppPlaceholderBody" in vba

        # Chart constants
        assert "Const xlColumnClustered" in vba
        assert "Const xlBarClustered" in vba
        assert "Const xlPie" in vba

    def test_vba_can_handle_all_placeholder_types(self, complete_plan):
        """Test that VBA can handle all placeholder types in the plan"""
        converter = PlanToVBAConverter(complete_plan)
        vba = converter.convert()

        # Load plan to check what's needed
        with open(complete_plan, 'r') as f:
            plan = json.load(f)

        # Check each slide's requirements are met
        for slide in plan["slides"]:
            for placeholder in slide["placeholders"]:
                if placeholder["placeholder_key"] == "Chart":
                    assert "CreateChartAtPlaceholder" in vba
                elif placeholder["placeholder_key"] == "Table":
                    assert "CreateTableAtPlaceholder" in vba
                else:
                    assert "SafeSetText" in vba


def test_vba_syntax_basic_validation():
    """Test basic VBA syntax validation"""
    from pathlib import Path
    fixtures_dir = Path(__file__).parent / "fixtures"

    # Create a simple converter
    plan = {
        "meta": {"title": "Syntax Test", "template_name": "test.pptx"},
        "slides": []
    }

    import tempfile
    with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False) as f:
        json.dump(plan, f)
        temp_path = f.name

    converter = PlanToVBAConverter(temp_path)
    vba = converter.convert()

    # Basic syntax checks
    assert vba.count("Sub ") == vba.count("End Sub")
    assert vba.count("Function ") == vba.count("End Function")
    assert vba.count("If ") <= vba.count("End If") + vba.count(" Then ")
    assert vba.count("For ") == vba.count("Next ")
    assert vba.count("With ") == vba.count("End With")

    # Clean up
    Path(temp_path).unlink()