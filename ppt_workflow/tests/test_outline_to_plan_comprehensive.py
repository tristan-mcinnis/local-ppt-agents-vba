"""
Comprehensive unit tests for outline_to_plan module
Tests normal cases, edge cases, and error conditions
"""

import json
import pytest
from pathlib import Path
import sys

# Add parent directory to path for imports
sys.path.insert(0, str(Path(__file__).parent.parent))
from core.outline_to_plan import OutlineToPlanConverter


class TestOutlineToPlanBasics:
    """Test basic functionality of outline to plan conversion"""

    @pytest.fixture
    def fixtures_dir(self):
        """Get fixtures directory path"""
        return Path(__file__).parent / "fixtures"

    @pytest.fixture
    def minimal_converter(self, fixtures_dir):
        """Create converter with minimal outline"""
        return OutlineToPlanConverter(
            str(fixtures_dir / "minimal_outline.json"),
            str(fixtures_dir / "sample_template_analysis.json")
        )

    def test_converter_initialization(self, minimal_converter):
        """Test that converter initializes correctly"""
        assert minimal_converter.outline is not None
        assert minimal_converter.analysis is not None
        assert minimal_converter.layout_index is not None
        assert isinstance(minimal_converter.errors, list)
        assert isinstance(minimal_converter.warnings, list)

    def test_minimal_outline_conversion(self, minimal_converter):
        """Test conversion of minimal outline"""
        plan = minimal_converter.convert()

        assert plan is not None
        assert "slides" in plan
        assert len(plan["slides"]) == 1
        assert plan["slides"][0]["layout"] == "Title Slide"

    def test_layout_index_building(self, minimal_converter):
        """Test that layout index is built correctly"""
        index = minimal_converter.layout_index

        assert "title slide" in index
        assert "title and content" in index
        assert index["title slide"]["index"] == 1
        assert index["title slide"]["placeholder_count"] == 2

    def test_normalize_name(self):
        """Test name normalization"""
        converter = OutlineToPlanConverter.__new__(OutlineToPlanConverter)

        assert converter._normalize_name("Title Slide") == "title slide"
        assert converter._normalize_name("  Title  ") == "title"
        assert converter._normalize_name("") == ""
        assert converter._normalize_name(None) == ""


class TestPlaceholderParsing:
    """Test placeholder parsing and mapping"""

    @pytest.fixture
    def converter(self, tmp_path):
        """Create converter with test data"""
        outline = {
            "meta": {"title": "Test", "author": "Test"},
            "slides": []
        }
        analysis = {
            "layouts": [{
                "index": 1,
                "name": "Test Layout",
                "placeholders": []
            }]
        }

        outline_path = tmp_path / "outline.json"
        analysis_path = tmp_path / "analysis.json"

        with open(outline_path, 'w') as f:
            json.dump(outline, f)
        with open(analysis_path, 'w') as f:
            json.dump(analysis, f)

        return OutlineToPlanConverter(str(outline_path), str(analysis_path))

    def test_parse_placeholder_key_simple(self, converter):
        """Test parsing simple placeholder keys"""
        type_name, ordinal = converter._parse_placeholder_key("Title", 1)
        assert type_name == "title"
        assert ordinal == 0

        type_name, ordinal = converter._parse_placeholder_key("Body", 1)
        assert type_name == "body"
        assert ordinal == 0

    def test_parse_placeholder_key_with_ordinal(self, converter):
        """Test parsing placeholder keys with ordinals"""
        type_name, ordinal = converter._parse_placeholder_key("Body[0]", 1)
        assert type_name == "body"
        assert ordinal == 0

        type_name, ordinal = converter._parse_placeholder_key("Body[1]", 1)
        assert type_name == "body"
        assert ordinal == 1

        type_name, ordinal = converter._parse_placeholder_key("Body[10]", 1)
        assert type_name == "body"
        assert ordinal == 10

    def test_parse_placeholder_key_invalid_ordinal(self, converter):
        """Test parsing placeholder keys with invalid ordinals"""
        type_name, ordinal = converter._parse_placeholder_key("Body[abc]", 1)
        assert type_name == "body"
        assert ordinal == 0
        assert len(converter.errors) > 0
        assert "invalid ordinal" in converter.errors[-1].lower()

    def test_parse_placeholder_key_edge_cases(self, converter):
        """Test edge cases in placeholder key parsing"""
        # Empty brackets
        type_name, ordinal = converter._parse_placeholder_key("Body[]", 1)
        assert type_name == "body"
        assert ordinal == 0

        # Missing closing bracket
        type_name, ordinal = converter._parse_placeholder_key("Body[1", 1)
        assert type_name == "body[1"
        assert ordinal == 0

        # Special characters
        type_name, ordinal = converter._parse_placeholder_key("CenterTitle", 1)
        assert type_name == "centertitle"
        assert ordinal == 0


class TestComplexOutlineConversion:
    """Test conversion of complex outlines with various content types"""

    @pytest.fixture
    def complex_converter(self):
        """Create converter with complex outline"""
        fixtures_dir = Path(__file__).parent / "fixtures"
        return OutlineToPlanConverter(
            str(fixtures_dir / "complex_outline.json"),
            str(fixtures_dir / "sample_template_analysis.json")
        )

    def test_complex_outline_structure(self, complex_converter):
        """Test that complex outline converts with all slides"""
        plan = complex_converter.convert()

        assert len(plan["slides"]) == 6
        assert plan["meta"]["title"] == "Complex Presentation"
        assert plan["meta"]["author"] == "Test Suite"

    def test_chart_placeholder_conversion(self, complex_converter):
        """Test conversion of chart placeholders"""
        plan = complex_converter.convert()

        # Find chart slide
        chart_slide = None
        for slide in plan["slides"]:
            if any("Chart" in str(p) for p in slide["placeholders"]):
                chart_slide = slide
                break

        assert chart_slide is not None

        # Check chart data is preserved
        chart_placeholder = None
        for p in chart_slide["placeholders"]:
            if p.get("placeholder_key") == "Chart":
                chart_placeholder = p
                break

        assert chart_placeholder is not None
        assert isinstance(chart_placeholder["content"], dict)
        assert chart_placeholder["content"]["type"] == "column"
        assert "data" in chart_placeholder["content"]

    def test_table_placeholder_conversion(self, complex_converter):
        """Test conversion of table placeholders"""
        plan = complex_converter.convert()

        # Find table slide
        table_slide = None
        for slide in plan["slides"]:
            if any("Table" in str(p) for p in slide["placeholders"]):
                table_slide = slide
                break

        assert table_slide is not None

        # Check table data is preserved
        table_placeholder = None
        for p in table_slide["placeholders"]:
            if p.get("placeholder_key") == "Table":
                table_placeholder = p
                break

        assert table_placeholder is not None
        assert isinstance(table_placeholder["content"], dict)
        assert "headers" in table_placeholder["content"]
        assert "rows" in table_placeholder["content"]

    def test_multiple_body_placeholders(self, complex_converter):
        """Test handling of multiple body placeholders with ordinals"""
        plan = complex_converter.convert()

        # Find two content slide
        two_content_slide = None
        for slide in plan["slides"]:
            if slide["layout"] == "Two Content":
                two_content_slide = slide
                break

        assert two_content_slide is not None

        # Check both body placeholders are mapped correctly
        body_placeholders = [
            p for p in two_content_slide["placeholders"]
            if p.get("placeholder_key", "").startswith("Body")
        ]

        assert len(body_placeholders) == 2

        # Check ordinals are correct
        ordinals = [p.get("ordinal", -1) for p in body_placeholders]
        assert 0 in ordinals
        assert 1 in ordinals


class TestEdgeCasesAndErrors:
    """Test edge cases and error handling"""

    @pytest.fixture
    def edge_converter(self):
        """Create converter with edge case outline"""
        fixtures_dir = Path(__file__).parent / "fixtures"
        return OutlineToPlanConverter(
            str(fixtures_dir / "edge_cases_outline.json"),
            str(fixtures_dir / "sample_template_analysis.json")
        )

    def test_special_characters_handling(self, edge_converter):
        """Test handling of special characters in content"""
        plan = edge_converter.convert()

        first_slide = plan["slides"][0]
        title_placeholder = None
        for p in first_slide["placeholders"]:
            if p.get("placeholder_key") == "Title":
                title_placeholder = p
                break

        assert title_placeholder is not None
        assert '"' in title_placeholder["content"]
        assert '&' in title_placeholder["content"]

    def test_empty_content_handling(self, edge_converter):
        """Test handling of empty placeholder content"""
        plan = edge_converter.convert()

        # Find slide with empty content
        empty_slide = plan["slides"][1]

        for p in empty_slide["placeholders"]:
            assert p["content"] == ""

    def test_nonexistent_layout_fallback(self, edge_converter):
        """Test fallback when layout doesn't exist"""
        plan = edge_converter.convert()

        assert len(edge_converter.warnings) > 0

        # Check that warning was generated for nonexistent layout
        layout_warning = any("NonExistentLayout" in w for w in edge_converter.warnings)
        assert layout_warning

    def test_invalid_placeholder_type(self, edge_converter):
        """Test handling of invalid placeholder types"""
        plan = edge_converter.convert()

        # Invalid placeholders should be skipped with warnings
        assert len(edge_converter.warnings) > 0

    def test_invalid_ordinal_handling(self, edge_converter):
        """Test handling of invalid ordinal numbers"""
        plan = edge_converter.convert()

        # Check for high ordinal number warning
        high_ordinal_warning = any("99" in str(w) for w in edge_converter.warnings)
        assert high_ordinal_warning

    def test_invalid_chart_type(self, edge_converter):
        """Test handling of invalid chart types"""
        plan = edge_converter.convert()

        # Find slide with invalid chart type
        for slide in plan["slides"]:
            for p in slide["placeholders"]:
                if p.get("placeholder_key") == "Chart":
                    content = p.get("content", {})
                    if isinstance(content, dict):
                        assert content.get("type") == "invalid_type"

    def test_empty_table_handling(self, edge_converter):
        """Test handling of empty tables"""
        plan = edge_converter.convert()

        # Find slide with empty table
        for slide in plan["slides"]:
            for p in slide["placeholders"]:
                if p.get("placeholder_key") == "Table":
                    content = p.get("content", {})
                    if isinstance(content, dict):
                        assert content.get("headers") == []
                        assert content.get("rows") == []


class TestLayoutMatching:
    """Test layout matching and selection logic"""

    @pytest.fixture
    def converter_with_layouts(self, tmp_path):
        """Create converter with specific layout configurations"""
        outline = {
            "meta": {"title": "Test", "author": "Test"},
            "slides": [
                {"layout": "Title Slide", "placeholders": {"Title": "Test"}},
                {"layout": "title slide", "placeholders": {"Title": "Test"}},  # Case variation
                {"layout": "Title  Slide", "placeholders": {"Title": "Test"}},  # Extra spaces
                {"layout": "Nonexistent", "placeholders": {"Title": "Test"}}  # Fallback
            ]
        }

        analysis = {
            "layouts": [
                {
                    "index": 1,
                    "name": "Title Slide",
                    "placeholders": [
                        {"type_name": "Title", "type_id": 1, "index": 0}
                    ]
                },
                {
                    "index": 2,
                    "name": "Title and Content",
                    "placeholders": [
                        {"type_name": "Title", "type_id": 1, "index": 0},
                        {"type_name": "Body", "type_id": 2, "index": 1}
                    ]
                }
            ]
        }

        outline_path = tmp_path / "outline.json"
        analysis_path = tmp_path / "analysis.json"

        with open(outline_path, 'w') as f:
            json.dump(outline, f)
        with open(analysis_path, 'w') as f:
            json.dump(analysis, f)

        return OutlineToPlanConverter(str(outline_path), str(analysis_path))

    def test_exact_layout_match(self, converter_with_layouts):
        """Test exact layout name matching"""
        plan = converter_with_layouts.convert()

        # First slide should match exactly
        assert plan["slides"][0]["selected_layout"]["name"] == "Title Slide"
        assert plan["slides"][0]["selected_layout"]["index"] == 1

    def test_case_insensitive_match(self, converter_with_layouts):
        """Test case-insensitive layout matching"""
        plan = converter_with_layouts.convert()

        # Second slide with different case should still match
        assert plan["slides"][1]["selected_layout"]["name"] == "Title Slide"
        assert plan["slides"][1]["selected_layout"]["index"] == 1

    def test_whitespace_normalized_match(self, converter_with_layouts):
        """Test whitespace normalization in matching"""
        plan = converter_with_layouts.convert()

        # Third slide with extra spaces should match
        assert plan["slides"][2]["selected_layout"]["name"] == "Title Slide"
        assert plan["slides"][2]["selected_layout"]["index"] == 1

    def test_fallback_layout_selection(self, converter_with_layouts):
        """Test fallback to default layout when no match"""
        plan = converter_with_layouts.convert()

        # Fourth slide should fallback
        assert plan["slides"][3]["selected_layout"]["name"] == "Title and Content"
        assert plan["slides"][3]["selected_layout"]["index"] == 2

        # Should have warning about fallback
        assert len(converter_with_layouts.warnings) > 0
        assert any("Nonexistent" in w for w in converter_with_layouts.warnings)


class TestValidation:
    """Test validation and error reporting"""

    def test_missing_slides_field(self, tmp_path):
        """Test handling of outline without slides field"""
        outline = {"meta": {"title": "Test"}}  # Missing slides
        analysis = {"layouts": []}

        outline_path = tmp_path / "outline.json"
        analysis_path = tmp_path / "analysis.json"

        with open(outline_path, 'w') as f:
            json.dump(outline, f)
        with open(analysis_path, 'w') as f:
            json.dump(analysis, f)

        converter = OutlineToPlanConverter(str(outline_path), str(analysis_path))
        plan = converter.convert()

        # Should still create a plan with empty slides
        assert plan is not None
        assert plan["slides"] == []

    def test_missing_meta_field(self, tmp_path):
        """Test handling of outline without meta field"""
        outline = {"slides": []}  # Missing meta
        analysis = {"layouts": []}

        outline_path = tmp_path / "outline.json"
        analysis_path = tmp_path / "analysis.json"

        with open(outline_path, 'w') as f:
            json.dump(outline, f)
        with open(analysis_path, 'w') as f:
            json.dump(analysis, f)

        converter = OutlineToPlanConverter(str(outline_path), str(analysis_path))
        plan = converter.convert()

        # Should create default meta
        assert plan is not None
        assert "meta" in plan
        assert plan["meta"]["title"] == "Untitled Presentation"

    def test_malformed_slide_structure(self, tmp_path):
        """Test handling of malformed slide objects"""
        outline = {
            "meta": {"title": "Test"},
            "slides": [
                {"layout": "Title Slide"},  # Missing placeholders
                {"placeholders": {"Title": "Test"}},  # Missing layout
                "Not a dict",  # Wrong type
                None  # Null value
            ]
        }
        analysis = {
            "layouts": [{
                "index": 1,
                "name": "Title Slide",
                "placeholders": []
            }]
        }

        outline_path = tmp_path / "outline.json"
        analysis_path = tmp_path / "analysis.json"

        with open(outline_path, 'w') as f:
            json.dump(outline, f)
        with open(analysis_path, 'w') as f:
            json.dump(analysis, f)

        converter = OutlineToPlanConverter(str(outline_path), str(analysis_path))
        plan = converter.convert()

        # Should handle gracefully
        assert plan is not None
        assert len(converter.errors) > 0 or len(converter.warnings) > 0


def test_fixtures_exist():
    """Test that all fixture files exist and are valid JSON"""
    fixtures_dir = Path(__file__).parent / "fixtures"

    fixture_files = [
        "minimal_outline.json",
        "complex_outline.json",
        "edge_cases_outline.json",
        "sample_template_analysis.json"
    ]

    for filename in fixture_files:
        filepath = fixtures_dir / filename
        assert filepath.exists(), f"Fixture {filename} does not exist"

        # Test that it's valid JSON
        with open(filepath, 'r') as f:
            data = json.load(f)
            assert data is not None