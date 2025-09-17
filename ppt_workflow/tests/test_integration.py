"""
Integration tests for the complete PowerPoint automation pipeline
Tests the full workflow from outline to VBA generation
"""

import json
import pytest
from pathlib import Path
import sys
import tempfile
import shutil

# Add parent directory to path for imports
sys.path.insert(0, str(Path(__file__).parent.parent))
from core.outline_to_plan import OutlineToPlanConverter
from core.plan_to_vba import PlanToVBAConverter
from workflow import WorkflowOrchestrator


class TestEndToEndWorkflow:
    """Test complete workflow from outline to VBA"""

    @pytest.fixture
    def temp_output_dir(self):
        """Create temporary output directory"""
        temp_dir = tempfile.mkdtemp(prefix="ppt_test_")
        yield Path(temp_dir)
        # Cleanup
        shutil.rmtree(temp_dir)

    @pytest.fixture
    def fixtures_dir(self):
        """Get fixtures directory"""
        return Path(__file__).parent / "fixtures"

    def test_minimal_workflow(self, fixtures_dir, temp_output_dir):
        """Test minimal outline through complete pipeline"""
        # Step 1: Outline to Plan
        converter1 = OutlineToPlanConverter(
            str(fixtures_dir / "minimal_outline.json"),
            str(fixtures_dir / "sample_template_analysis.json")
        )
        plan = converter1.convert()

        assert plan is not None
        assert len(plan["slides"]) == 1

        # Save plan
        plan_path = temp_output_dir / "plan.json"
        with open(plan_path, 'w') as f:
            json.dump(plan, f)

        # Step 2: Plan to VBA
        converter2 = PlanToVBAConverter(str(plan_path))
        vba = converter2.convert()

        assert vba is not None
        assert "Sub Main()" in vba
        assert len(converter1.errors) == 0

    def test_complex_workflow(self, fixtures_dir, temp_output_dir):
        """Test complex outline with charts and tables"""
        # Step 1: Outline to Plan
        converter1 = OutlineToPlanConverter(
            str(fixtures_dir / "complex_outline.json"),
            str(fixtures_dir / "sample_template_analysis.json")
        )
        plan = converter1.convert()

        assert len(plan["slides"]) == 6

        # Save plan
        plan_path = temp_output_dir / "plan.json"
        with open(plan_path, 'w') as f:
            json.dump(plan, f)

        # Step 2: Plan to VBA
        converter2 = PlanToVBAConverter(str(plan_path))
        vba = converter2.convert()

        # Check for chart and table handling
        assert "CreateChartAtPlaceholder" in vba
        assert "CreateTableAtPlaceholder" in vba

    def test_edge_cases_workflow(self, fixtures_dir, temp_output_dir):
        """Test edge cases through complete pipeline"""
        # Step 1: Outline to Plan
        converter1 = OutlineToPlanConverter(
            str(fixtures_dir / "edge_cases_outline.json"),
            str(fixtures_dir / "sample_template_analysis.json")
        )
        plan = converter1.convert()

        # Should have warnings but still generate plan
        assert len(plan["slides"]) > 0
        assert len(converter1.warnings) > 0

        # Save plan
        plan_path = temp_output_dir / "plan.json"
        with open(plan_path, 'w') as f:
            json.dump(plan, f)

        # Step 2: Plan to VBA
        converter2 = PlanToVBAConverter(str(plan_path))
        vba = converter2.convert()

        # Should still generate valid VBA
        assert "Sub Main()" in vba
        assert "End Sub" in vba


class TestWorkflowOrchestrator:
    """Test the workflow orchestrator component"""

    @pytest.fixture
    def orchestrator(self, monkeypatch, tmp_path):
        """Create orchestrator with mocked output directory"""
        orchestrator = WorkflowOrchestrator(verbose=False)
        # Override output directory
        orchestrator.output_dir = tmp_path / "output"
        orchestrator.output_dir.mkdir(exist_ok=True)
        return orchestrator

    @pytest.fixture
    def fixtures_dir(self):
        """Get fixtures directory"""
        return Path(__file__).parent / "fixtures"

    def test_orchestrator_validation(self, orchestrator, fixtures_dir):
        """Test file validation in orchestrator"""
        valid = orchestrator.validate_files(
            fixtures_dir / "minimal_outline.json",
            fixtures_dir / "sample_template_analysis.json"
        )
        assert valid is True

        # Test with non-existent file
        valid = orchestrator.validate_files(
            fixtures_dir / "nonexistent.json",
            fixtures_dir / "sample_template_analysis.json"
        )
        assert valid is False

    def test_orchestrator_complete_workflow(self, orchestrator, fixtures_dir):
        """Test complete workflow through orchestrator"""
        success = orchestrator.run_workflow(
            str(fixtures_dir / "minimal_outline.json"),
            str(fixtures_dir / "sample_template_analysis.json"),
            skip_validation=False
        )

        assert success is True

        # Check output files exist
        assert (orchestrator.output_dir / "slide_plan.json").exists()
        assert (orchestrator.output_dir / "generated_script.vba").exists()

    def test_orchestrator_error_handling(self, orchestrator, tmp_path):
        """Test orchestrator error handling"""
        # Create invalid JSON files
        invalid_outline = tmp_path / "invalid.json"
        invalid_outline.write_text("{invalid json")

        valid_analysis = tmp_path / "analysis.json"
        valid_analysis.write_text('{"layouts": []}')

        success = orchestrator.run_workflow(
            str(invalid_outline),
            str(valid_analysis)
        )

        assert success is False


class TestDataFlowIntegrity:
    """Test data integrity through the pipeline"""

    @pytest.fixture
    def test_data(self, tmp_path):
        """Create test data with known values"""
        outline = {
            "meta": {
                "title": "Data Integrity Test",
                "author": "Test Suite",
                "custom_field": "Should be preserved"
            },
            "slides": [
                {
                    "layout": "Title Slide",
                    "placeholders": {
                        "Title": "Special \" & < > Characters",
                        "Subtitle": "Multi\nLine\nText"
                    }
                },
                {
                    "layout": "Title and Content",
                    "placeholders": {
                        "Chart": {
                            "type": "column",
                            "data": {
                                "x": ["Q1", "Q2", "Q3", "Q4"],
                                "series": [
                                    {"name": "Revenue", "data": [100.5, 200.75, 150.25, 175.0]}
                                ]
                            }
                        }
                    }
                }
            ]
        }

        analysis = {
            "layouts": [
                {
                    "index": 1,
                    "name": "Title Slide",
                    "placeholders": [
                        {"type_name": "Title", "type_id": 1, "index": 0},
                        {"type_name": "Subtitle", "type_id": 4, "index": 1}
                    ]
                },
                {
                    "index": 2,
                    "name": "Title and Content",
                    "placeholders": [
                        {"type_name": "Title", "type_id": 1, "index": 0},
                        {"type_name": "Content", "type_id": 19, "index": 1}
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

        return outline_path, analysis_path, outline

    def test_metadata_preservation(self, test_data):
        """Test that metadata is preserved through pipeline"""
        outline_path, analysis_path, original_outline = test_data

        # Convert to plan
        converter = OutlineToPlanConverter(str(outline_path), str(analysis_path))
        plan = converter.convert()

        # Check metadata preservation
        assert plan["meta"]["title"] == original_outline["meta"]["title"]
        assert plan["meta"]["author"] == original_outline["meta"]["author"]

    def test_special_characters_preservation(self, test_data, tmp_path):
        """Test special characters are properly escaped"""
        outline_path, analysis_path, _ = test_data

        # Convert to plan
        converter1 = OutlineToPlanConverter(str(outline_path), str(analysis_path))
        plan = converter1.convert()

        # Save and convert to VBA
        plan_path = tmp_path / "plan.json"
        with open(plan_path, 'w') as f:
            json.dump(plan, f)

        converter2 = PlanToVBAConverter(str(plan_path))
        vba = converter2.convert()

        # Check special characters are escaped in VBA
        assert '""' in vba  # Quotes should be doubled
        assert 'vbCrLf' in vba  # Newlines should be converted

    def test_numeric_data_preservation(self, test_data, tmp_path):
        """Test numeric data in charts is preserved"""
        outline_path, analysis_path, _ = test_data

        # Full pipeline
        converter1 = OutlineToPlanConverter(str(outline_path), str(analysis_path))
        plan = converter1.convert()

        # Check chart data is preserved in plan
        chart_slide = plan["slides"][1]
        chart_placeholder = next(
            p for p in chart_slide["placeholders"]
            if p["placeholder_key"] == "Chart"
        )

        assert chart_placeholder["content"]["data"]["series"][0]["data"] == [100.5, 200.75, 150.25, 175.0]


class TestErrorRecovery:
    """Test error recovery and graceful degradation"""

    def test_missing_layout_recovery(self, tmp_path):
        """Test recovery when layout is not found"""
        outline = {
            "meta": {"title": "Test"},
            "slides": [
                {"layout": "Nonexistent Layout", "placeholders": {"Title": "Test"}}
            ]
        }

        analysis = {
            "layouts": [
                {
                    "index": 1,
                    "name": "Default",
                    "placeholders": [{"type_name": "Title", "type_id": 1, "index": 0}]
                }
            ]
        }

        outline_path = tmp_path / "outline.json"
        analysis_path = tmp_path / "analysis.json"

        with open(outline_path, 'w') as f:
            json.dump(outline, f)
        with open(analysis_path, 'w') as f:
            json.dump(analysis, f)

        # Should recover with fallback
        converter = OutlineToPlanConverter(str(outline_path), str(analysis_path))
        plan = converter.convert()

        assert len(plan["slides"]) == 1
        assert plan["slides"][0]["selected_layout"]["name"] == "Default"
        assert len(converter.warnings) > 0

    def test_invalid_placeholder_recovery(self, tmp_path):
        """Test recovery with invalid placeholder types"""
        outline = {
            "meta": {"title": "Test"},
            "slides": [
                {
                    "layout": "Test",
                    "placeholders": {
                        "InvalidType": "Content",
                        "Title": "Valid Title"
                    }
                }
            ]
        }

        analysis = {
            "layouts": [
                {
                    "index": 1,
                    "name": "Test",
                    "placeholders": [
                        {"type_name": "Title", "type_id": 1, "index": 0}
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

        converter = OutlineToPlanConverter(str(outline_path), str(analysis_path))
        plan = converter.convert()

        # Should skip invalid placeholder
        slide_placeholders = plan["slides"][0]["placeholders"]
        assert len(slide_placeholders) == 1
        assert slide_placeholders[0]["placeholder_key"] == "Title"


class TestPerformance:
    """Test performance with large inputs"""

    def test_large_presentation(self, tmp_path):
        """Test handling of large presentations"""
        # Create outline with many slides
        outline = {
            "meta": {"title": "Large Test"},
            "slides": [
                {
                    "layout": "Title and Content",
                    "placeholders": {
                        "Title": f"Slide {i}",
                        "Body": f"Content for slide {i}\n" * 10
                    }
                }
                for i in range(100)
            ]
        }

        analysis = {
            "layouts": [
                {
                    "index": 1,
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
        plan_path = tmp_path / "plan.json"

        with open(outline_path, 'w') as f:
            json.dump(outline, f)
        with open(analysis_path, 'w') as f:
            json.dump(analysis, f)

        # Test performance
        import time

        start = time.time()
        converter1 = OutlineToPlanConverter(str(outline_path), str(analysis_path))
        plan = converter1.convert()
        plan_time = time.time() - start

        with open(plan_path, 'w') as f:
            json.dump(plan, f)

        start = time.time()
        converter2 = PlanToVBAConverter(str(plan_path))
        vba = converter2.convert()
        vba_time = time.time() - start

        # Should complete in reasonable time
        assert plan_time < 1.0  # Less than 1 second for plan
        assert vba_time < 2.0   # Less than 2 seconds for VBA

        # Check output validity
        assert len(plan["slides"]) == 100
        assert "Sub Main()" in vba