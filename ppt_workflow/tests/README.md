# PowerPoint Automation Test Suite Documentation

## Overview

This comprehensive test suite validates all aspects of the PowerPoint automation pipeline, from JSON outline parsing to VBA script generation. The tests are designed to be accessible to users of all skill levels.

## Test Structure

```
tests/
├── fixtures/                      # Test data files
│   ├── minimal_outline.json      # Simple test case
│   ├── complex_outline.json      # Advanced features
│   ├── edge_cases_outline.json   # Edge cases & errors
│   └── sample_template_analysis.json  # Template structure
├── test_outline_to_plan.py       # Basic outline conversion tests
├── test_plan_to_vba.py          # Basic VBA generation tests
├── test_robustness.py           # Robustness & edge case tests
├── test_outline_to_plan_comprehensive.py  # Detailed outline tests
├── test_plan_to_vba_comprehensive.py      # Detailed VBA tests
├── test_integration.py          # End-to-end workflow tests
└── README.md                    # This file
```

## Quick Start

### For Non-Technical Users

1. **Run basic tests** (fastest):
   ```bash
   python run_tests.py
   ```

2. **Run all tests** (most thorough):
   ```bash
   python run_tests.py --all
   ```

3. **Generate a test report**:
   ```bash
   python run_tests.py --report
   ```

### For Developers

1. **Run specific test categories**:
   ```bash
   # Test outline processing only
   python run_tests.py --category outline

   # Test VBA generation only
   python run_tests.py --category vba

   # Test integration only
   python run_tests.py --category integration

   # Test edge cases
   python run_tests.py --category edge
   ```

2. **Run with verbose output**:
   ```bash
   python run_tests.py --verbose
   ```

3. **Run with pytest directly**:
   ```bash
   # From parent directory
   python -m pytest ppt_workflow/tests/ -v

   # Specific test file
   python -m pytest ppt_workflow/tests/test_outline_to_plan.py

   # Specific test class or function
   python -m pytest ppt_workflow/tests/test_plan_to_vba.py::TestPlanToVBAConverter
   ```

## Test Categories

### 1. Unit Tests

#### Outline to Plan (`test_outline_to_plan*.py`)
- **Purpose**: Validate JSON outline parsing and plan generation
- **Coverage**:
  - Layout matching and selection
  - Placeholder mapping
  - Special character handling
  - Error recovery and fallbacks
  - Edge cases (empty content, invalid layouts)

#### Plan to VBA (`test_plan_to_vba*.py`)
- **Purpose**: Validate VBA script generation
- **Coverage**:
  - macOS compatibility (Collection vs Dictionary)
  - Placeholder handling (text, charts, tables)
  - Error logging and recovery
  - JSON parsing functions in VBA
  - Layout caching optimization

### 2. Integration Tests (`test_integration.py`)
- **Purpose**: Test complete workflow from outline to VBA
- **Coverage**:
  - End-to-end pipeline validation
  - Data integrity through transformations
  - Error propagation and handling
  - Performance with large inputs

### 3. Robustness Tests (`test_robustness.py`)
- **Purpose**: Validate system resilience
- **Coverage**:
  - Malformed input handling
  - Missing data recovery
  - Special character escaping
  - Image placeholder skipping

## Test Fixtures

The `fixtures/` directory contains test data files:

### `minimal_outline.json`
- Simplest valid outline
- Single slide with basic placeholders
- Used for smoke testing

### `complex_outline.json`
- Multiple slides with various layouts
- Charts, tables, and multiple body placeholders
- Tests advanced features

### `edge_cases_outline.json`
- Invalid and edge case data
- Special characters, empty content
- Tests error handling and recovery

### `sample_template_analysis.json`
- Mock PowerPoint template structure
- 10 different layouts with various placeholders
- Used to test layout matching

## Expected Test Results

### Passing Tests ✅
- All basic functionality tests should pass
- Original test files (`test_outline_to_plan.py`, `test_plan_to_vba.py`, `test_robustness.py`)
- Core workflow operations

### Known Limitations ⚠️
- Comprehensive tests may fail if the converter is stricter than expected
- Some edge cases intentionally trigger warnings/errors
- Performance tests have generous time limits

## Debugging Failed Tests

### Common Issues and Solutions

1. **Import Errors**
   ```
   ModuleNotFoundError: No module named 'ppt_workflow'
   ```
   **Solution**: Run tests from the parent directory or use the `run_tests.py` script

2. **Fixture Not Found**
   ```
   FileNotFoundError: fixtures/minimal_outline.json
   ```
   **Solution**: Ensure you're running from the correct directory

3. **Placeholder Not Found**
   ```
   ValueError: Placeholder 'Title' not found in layout
   ```
   **Solution**: Check that fixture files match the expected template structure

4. **Test Timeout**
   ```
   Test exceeded time limit
   ```
   **Solution**: Performance tests allow up to 2 seconds; optimize code if consistently failing

## Adding New Tests

### 1. Create a new test file:
```python
# tests/test_new_feature.py
import pytest
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).parent.parent))
from core.module_to_test import FunctionToTest

class TestNewFeature:
    def test_basic_functionality(self):
        result = FunctionToTest(input_data)
        assert result == expected_output
```

### 2. Add test fixtures if needed:
```python
@pytest.fixture
def sample_data():
    return {"key": "value"}
```

### 3. Update test runner categories:
Edit `run_tests.py` to include your new test in the appropriate category.

## Continuous Testing

### Pre-commit Testing
Run quick tests before committing:
```bash
python run_tests.py --quick
```

### Full Validation
Run comprehensive tests before releases:
```bash
python run_tests.py --all --report
```

## Test Coverage

Current test coverage includes:

- ✅ **Core Functions**: 100% of main pipeline functions
- ✅ **Error Handling**: Major error paths tested
- ✅ **Platform Compatibility**: macOS-specific features validated
- ✅ **Data Types**: Text, charts, tables, special characters
- ✅ **Edge Cases**: Empty inputs, invalid data, missing fields

## Support

If tests are failing:

1. Check you're using Python 3.7+
2. Ensure all dependencies are installed
3. Verify fixture files are present and valid JSON
4. Run with `--verbose` for detailed output
5. Check the generated `test_report.txt` for details

## Best Practices

1. **Run tests frequently** during development
2. **Start with quick tests** for rapid feedback
3. **Use comprehensive tests** before deployment
4. **Keep fixtures updated** when changing data structures
5. **Add tests for new features** before implementing them
6. **Document test failures** that are expected/acceptable