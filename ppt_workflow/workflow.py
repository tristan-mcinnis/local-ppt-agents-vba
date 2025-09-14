"""
PowerPoint Automation Workflow Orchestrator
Main entry point for the VBA generation pipeline
"""

import sys
import json
import argparse
from pathlib import Path
from datetime import datetime
from typing import Optional, Dict, Any

# Add core modules to path
sys.path.insert(0, str(Path(__file__).parent))

from core.outline_to_plan import OutlineToPlanConverter
from core.plan_to_vba import PlanToVBAConverter


class WorkflowOrchestrator:
    """Orchestrates the complete PowerPoint automation workflow"""

    def __init__(self, verbose: bool = True):
        """Initialize orchestrator"""
        self.verbose = verbose
        self.workflow_dir = Path(__file__).parent
        self.output_dir = self.workflow_dir / "output"
        self.output_dir.mkdir(exist_ok=True)

    def log(self, message: str, level: str = "INFO"):
        """Log message with timestamp"""
        if self.verbose:
            timestamp = datetime.now().strftime("%H:%M:%S")
            prefix = {
                "INFO": "ℹ",
                "SUCCESS": "✓",
                "WARNING": "⚠",
                "ERROR": "✗"
            }.get(level, "•")
            print(f"[{timestamp}] {prefix} {message}")

    def validate_files(self, outline_path: Path, analysis_path: Path) -> bool:
        """Validate input files exist and are valid JSON"""
        # Check existence
        if not outline_path.exists():
            self.log(f"Outline file not found: {outline_path}", "ERROR")
            return False

        if not analysis_path.exists():
            self.log(f"Analysis file not found: {analysis_path}", "ERROR")
            return False

        # Check JSON validity
        try:
            with open(outline_path, 'r') as f:
                outline = json.load(f)
                if "slides" not in outline:
                    self.log("Outline missing 'slides' field", "ERROR")
                    return False

            with open(analysis_path, 'r') as f:
                analysis = json.load(f)
                if "layouts" not in analysis:
                    self.log("Analysis missing 'layouts' field", "ERROR")
                    return False

            self.log("Input files validated successfully", "SUCCESS")
            return True

        except json.JSONDecodeError as e:
            self.log(f"Invalid JSON: {e}", "ERROR")
            return False

    def run_step1_outline_to_plan(self, outline_path: Path, analysis_path: Path) -> Optional[Path]:
        """Step 1: Convert outline to plan"""
        self.log("Step 1: Converting outline to slide plan...")

        try:
            # Create converter
            converter = OutlineToPlanConverter(str(outline_path), str(analysis_path))

            # Convert to plan
            plan = converter.convert()

            # Save plan
            plan_path = self.output_dir / "slide_plan.json"
            with open(plan_path, 'w', encoding='utf-8') as f:
                json.dump(plan, f, indent=2, ensure_ascii=False)

            # Report results
            self.log(f"Generated plan with {len(plan['slides'])} slides", "SUCCESS")

            if converter.errors:
                for err in converter.errors:
                    self.log(err, "ERROR")
                return None

            if converter.warnings:
                for warn in converter.warnings:
                    self.log(warn, "WARNING")

            # Display layout usage
            self.log("Layout usage summary:")
            for layout, count in plan["layout_usage_summary"].items():
                self.log(f"  • {layout}: {count} slide(s)")

            return plan_path

        except Exception as e:
            self.log(f"Failed to generate plan: {e}", "ERROR")
            return None

    def run_step2_plan_to_vba(self, plan_path: Path) -> Optional[Path]:
        """Step 2: Convert plan to VBA"""
        self.log("Step 2: Generating VBA script from plan...")

        try:
            # Create converter
            converter = PlanToVBAConverter(str(plan_path))

            # Generate VBA
            vba_code = converter.convert()

            # Save VBA script
            vba_path = self.output_dir / "generated_script.vba"
            with open(vba_path, 'w', encoding='utf-8') as f:
                f.write(vba_code)

            # Report results
            slide_count = len(converter.plan["slides"])
            layout_count = len(converter.used_layouts)

            self.log(f"Generated VBA for {slide_count} slides", "SUCCESS")
            self.log(f"Using {layout_count} unique layouts", "INFO")

            return vba_path

        except Exception as e:
            self.log(f"Failed to generate VBA: {e}", "ERROR")
            return None

    def run_validation(self, plan_path: Path, vba_path: Path) -> bool:
        """Optional: Validate generated artifacts"""
        self.log("Step 3: Validating generated files...")

        try:
            # Load files
            with open(plan_path, 'r') as f:
                plan = json.load(f)

            with open(vba_path, 'r') as f:
                vba_content = f.read()

            # Basic validations
            checks = []

            # Check VBA has Main subroutine
            if "Sub Main()" in vba_content:
                checks.append(("Main subroutine present", True))
            else:
                checks.append(("Main subroutine present", False))

            # Check VBA uses ActivePresentation
            if "Application.ActivePresentation" in vba_content:
                checks.append(("Uses ActivePresentation", True))
            else:
                checks.append(("Uses ActivePresentation", False))

            # Check layout indices are referenced
            for slide in plan["slides"]:
                layout_idx = slide["selected_layout"]["index"]
                if str(layout_idx) in vba_content:
                    checks.append((f"Layout {layout_idx} referenced", True))
                else:
                    checks.append((f"Layout {layout_idx} referenced", False))

            # Report validation results
            all_passed = all(result for _, result in checks)

            for check_name, passed in checks:
                if passed:
                    self.log(f"✓ {check_name}", "SUCCESS")
                else:
                    self.log(f"✗ {check_name}", "ERROR")

            return all_passed

        except Exception as e:
            self.log(f"Validation failed: {e}", "ERROR")
            return False

    def run_workflow(self, outline_path: str, analysis_path: str,
                    skip_validation: bool = False) -> bool:
        """Run complete workflow"""
        self.log("=" * 60)
        self.log("POWERPOINT AUTOMATION WORKFLOW")
        self.log("=" * 60)

        # Convert to Path objects
        outline_path = Path(outline_path).resolve()
        analysis_path = Path(analysis_path).resolve()

        # Validate inputs
        if not self.validate_files(outline_path, analysis_path):
            return False

        # Step 1: Outline to Plan
        self.log("-" * 40)
        plan_path = self.run_step1_outline_to_plan(outline_path, analysis_path)
        if not plan_path:
            self.log("Workflow failed at Step 1", "ERROR")
            return False

        # Step 2: Plan to VBA
        self.log("-" * 40)
        vba_path = self.run_step2_plan_to_vba(plan_path)
        if not vba_path:
            self.log("Workflow failed at Step 2", "ERROR")
            return False

        # Step 3: Validation (optional)
        if not skip_validation:
            self.log("-" * 40)
            validation_passed = self.run_validation(plan_path, vba_path)
            if not validation_passed:
                self.log("Validation found issues - review before running", "WARNING")

        # Success summary
        self.log("=" * 60)
        self.log("WORKFLOW COMPLETED SUCCESSFULLY", "SUCCESS")
        self.log(f"Output directory: {self.output_dir}")
        self.log("Generated files:")
        self.log(f"  • slide_plan.json")
        self.log(f"  • generated_script.vba")
        self.log("")
        self.log("Next steps:")
        self.log("1. Open your PowerPoint template")
        self.log("2. Press Alt+F11 (Windows) or Opt+F11 (Mac)")
        self.log("3. Insert > Module")
        self.log(f"4. Copy contents of {vba_path.name}")
        self.log("5. Run 'ValidateTemplate' to check compatibility")
        self.log("6. Run 'Main' to create slides")

        return True


def main():
    """CLI entry point"""
    parser = argparse.ArgumentParser(
        description="PowerPoint Automation Workflow - Generate VBA from outline"
    )

    parser.add_argument(
        "outline",
        help="Path to outline.json file"
    )

    parser.add_argument(
        "analysis",
        help="Path to template_analysis.json file"
    )

    parser.add_argument(
        "--skip-validation",
        action="store_true",
        help="Skip validation step"
    )

    parser.add_argument(
        "--quiet",
        action="store_true",
        help="Minimal output"
    )

    args = parser.parse_args()

    # Run workflow
    orchestrator = WorkflowOrchestrator(verbose=not args.quiet)
    success = orchestrator.run_workflow(
        args.outline,
        args.analysis,
        skip_validation=args.skip_validation
    )

    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()