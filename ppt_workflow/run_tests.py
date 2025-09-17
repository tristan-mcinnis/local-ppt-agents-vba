#!/usr/bin/env python3
"""
User-friendly test runner for PowerPoint Automation System
Provides easy test execution for users with limited coding experience
"""

import sys
import subprocess
import argparse
from pathlib import Path
from datetime import datetime
import json


class TestRunner:
    """User-friendly test runner with clear output and reporting"""

    def __init__(self, verbose=False):
        self.verbose = verbose
        self.test_dir = Path(__file__).parent / "tests"
        self.results = []

    def print_header(self, text, char="="):
        """Print formatted header"""
        line = char * 60
        print(f"\n{line}")
        print(f" {text}")
        print(f"{line}")

    def print_status(self, test_name, status, details=""):
        """Print test status with color coding"""
        symbols = {
            "passed": "✓",
            "failed": "✗",
            "skipped": "○",
            "running": "⟳"
        }

        colors = {
            "passed": "\033[92m",  # Green
            "failed": "\033[91m",  # Red
            "skipped": "\033[93m",  # Yellow
            "running": "\033[94m",  # Blue
            "reset": "\033[0m"
        }

        symbol = symbols.get(status, "?")
        color = colors.get(status, "")
        reset = colors["reset"]

        if details:
            print(f"  {color}{symbol}{reset} {test_name}: {details}")
        else:
            print(f"  {color}{symbol}{reset} {test_name}")

    def run_test_module(self, module_name, description):
        """Run a specific test module"""
        self.print_status(description, "running")

        test_file = self.test_dir / module_name

        if not test_file.exists():
            self.print_status(description, "skipped", "File not found")
            return False

        try:
            cmd = [
                sys.executable, "-m", "pytest",
                str(test_file),
                "-v" if self.verbose else "-q",
                "--tb=short",
                "--no-header"
            ]

            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                cwd=self.test_dir.parent.parent  # Run from project root
            )

            if result.returncode == 0:
                # Extract test counts from output
                lines = result.stdout.split('\n')
                for line in lines:
                    if "passed" in line:
                        self.print_status(description, "passed", line.strip())
                        self.results.append((description, "passed", line.strip()))
                        return True

                self.print_status(description, "passed")
                self.results.append((description, "passed", "All tests passed"))
                return True
            else:
                # Extract failure info
                error_msg = "See verbose output for details"
                if self.verbose:
                    print(result.stdout)
                    print(result.stderr)
                else:
                    # Try to extract key error
                    for line in result.stdout.split('\n'):
                        if "FAILED" in line or "ERROR" in line:
                            error_msg = line.strip()
                            break

                self.print_status(description, "failed", error_msg)
                self.results.append((description, "failed", error_msg))
                return False

        except Exception as e:
            self.print_status(description, "failed", f"Error: {e}")
            self.results.append((description, "failed", str(e)))
            return False

    def run_quick_tests(self):
        """Run quick unit tests only"""
        self.print_header("RUNNING QUICK TESTS")

        tests = [
            ("test_outline_to_plan.py", "Basic outline conversion"),
            ("test_plan_to_vba.py", "Basic VBA generation"),
            ("test_robustness.py", "Robustness checks")
        ]

        passed = 0
        for module, desc in tests:
            if self.run_test_module(module, desc):
                passed += 1

        return passed, len(tests)

    def run_comprehensive_tests(self):
        """Run all comprehensive tests"""
        self.print_header("RUNNING COMPREHENSIVE TESTS")

        tests = [
            ("test_outline_to_plan_comprehensive.py", "Comprehensive outline tests"),
            ("test_plan_to_vba_comprehensive.py", "Comprehensive VBA tests"),
            ("test_integration.py", "Integration tests")
        ]

        passed = 0
        for module, desc in tests:
            if self.run_test_module(module, desc):
                passed += 1

        return passed, len(tests)

    def run_specific_category(self, category):
        """Run tests for a specific category"""
        categories = {
            "outline": [
                ("test_outline_to_plan.py", "Basic outline conversion"),
                ("test_outline_to_plan_comprehensive.py", "Comprehensive outline tests")
            ],
            "vba": [
                ("test_plan_to_vba.py", "Basic VBA generation"),
                ("test_plan_to_vba_comprehensive.py", "Comprehensive VBA tests")
            ],
            "integration": [
                ("test_integration.py", "Integration tests")
            ],
            "edge": [
                ("test_robustness.py", "Robustness checks")
            ]
        }

        if category not in categories:
            print(f"Unknown category: {category}")
            print(f"Available: {', '.join(categories.keys())}")
            return 0, 0

        self.print_header(f"RUNNING {category.upper()} TESTS")

        tests = categories[category]
        passed = 0
        for module, desc in tests:
            if self.run_test_module(module, desc):
                passed += 1

        return passed, len(tests)

    def check_fixtures(self):
        """Verify test fixtures are present"""
        self.print_header("CHECKING TEST FIXTURES")

        fixtures_dir = self.test_dir / "fixtures"
        required_fixtures = [
            "minimal_outline.json",
            "complex_outline.json",
            "edge_cases_outline.json",
            "sample_template_analysis.json"
        ]

        all_present = True
        for fixture in required_fixtures:
            path = fixtures_dir / fixture
            if path.exists():
                try:
                    with open(path, 'r') as f:
                        json.load(f)
                    self.print_status(fixture, "passed", "Valid JSON")
                except json.JSONDecodeError:
                    self.print_status(fixture, "failed", "Invalid JSON")
                    all_present = False
            else:
                self.print_status(fixture, "failed", "Not found")
                all_present = False

        return all_present

    def print_summary(self):
        """Print test summary"""
        self.print_header("TEST SUMMARY")

        passed = sum(1 for _, status, _ in self.results if status == "passed")
        failed = sum(1 for _, status, _ in self.results if status == "failed")
        total = len(self.results)

        print(f"\n  Total Tests Run: {total}")
        print(f"  ✓ Passed: {passed}")
        print(f"  ✗ Failed: {failed}")

        if failed > 0:
            print("\n  Failed Tests:")
            for name, status, details in self.results:
                if status == "failed":
                    print(f"    • {name}: {details}")

        # Overall status
        print()
        if failed == 0:
            print("  🎉 All tests passed successfully!")
            return True
        else:
            print(f"  ⚠️  {failed} test(s) failed. Please review the output above.")
            return False

    def generate_report(self, output_file="test_report.txt"):
        """Generate detailed test report"""
        report_path = Path(output_file)

        with open(report_path, 'w') as f:
            f.write("=" * 60 + "\n")
            f.write("POWERPOINT AUTOMATION TEST REPORT\n")
            f.write(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write("=" * 60 + "\n\n")

            f.write("TEST RESULTS:\n")
            f.write("-" * 40 + "\n")

            for name, status, details in self.results:
                symbol = "✓" if status == "passed" else "✗"
                f.write(f"{symbol} {name}\n")
                if details:
                    f.write(f"  Details: {details}\n")
                f.write("\n")

            # Summary
            passed = sum(1 for _, status, _ in self.results if status == "passed")
            failed = sum(1 for _, status, _ in self.results if status == "failed")
            total = len(self.results)

            f.write("-" * 40 + "\n")
            f.write(f"SUMMARY: {passed}/{total} tests passed\n")

            if failed == 0:
                f.write("\n✓ All tests passed successfully!\n")
            else:
                f.write(f"\n⚠ {failed} test(s) failed\n")

        print(f"\n  📄 Report saved to: {report_path.absolute()}")


def main():
    """Main entry point for test runner"""
    parser = argparse.ArgumentParser(
        description="Run tests for PowerPoint Automation System",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s              # Run quick tests
  %(prog)s --all        # Run all tests
  %(prog)s --category vba  # Run only VBA tests
  %(prog)s --verbose    # Show detailed output
        """
    )

    parser.add_argument(
        "--all", "-a",
        action="store_true",
        help="Run all comprehensive tests"
    )

    parser.add_argument(
        "--quick", "-q",
        action="store_true",
        help="Run quick tests only (default)"
    )

    parser.add_argument(
        "--category", "-c",
        choices=["outline", "vba", "integration", "edge"],
        help="Run tests for specific category"
    )

    parser.add_argument(
        "--verbose", "-v",
        action="store_true",
        help="Show detailed test output"
    )

    parser.add_argument(
        "--report", "-r",
        action="store_true",
        help="Generate test report file"
    )

    parser.add_argument(
        "--check-fixtures",
        action="store_true",
        help="Only check if test fixtures are present"
    )

    args = parser.parse_args()

    # Create test runner
    runner = TestRunner(verbose=args.verbose)

    print("=" * 60)
    print(" POWERPOINT AUTOMATION TEST SUITE")
    print("=" * 60)

    # Check fixtures first
    if args.check_fixtures or not (args.all or args.category):
        if not runner.check_fixtures():
            print("\n⚠️  Some fixtures are missing or invalid!")
            if args.check_fixtures:
                return 1

    # Run appropriate tests
    if args.category:
        passed, total = runner.run_specific_category(args.category)
    elif args.all:
        # Run all tests
        passed1, total1 = runner.run_quick_tests()
        passed2, total2 = runner.run_comprehensive_tests()
        passed = passed1 + passed2
        total = total1 + total2
    else:
        # Default to quick tests
        passed, total = runner.run_quick_tests()

    # Print summary
    success = runner.print_summary()

    # Generate report if requested
    if args.report:
        runner.generate_report()

    # Return appropriate exit code
    return 0 if success else 1


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("\n\nTest run interrupted by user")
        sys.exit(1)
    except Exception as e:
        print(f"\n✗ Unexpected error: {e}")
        sys.exit(1)