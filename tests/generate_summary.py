#!/usr/bin/env python3
"""
Generate test results summary with descriptions.

Usage: python tests/generate_summary.py
Output: tests/artifacts/TEST_RESULTS_SUMMARY.md
"""
import ast
import subprocess
import sys
from datetime import datetime
from pathlib import Path


# Test descriptions mapping
TEST_DESCRIPTIONS = {
    # Scalar replacement
    "test_scalar_replacement_body_flow": "Validates scalar placeholder replacement in BODY FLOW paragraphs. Ensures all scalar values appear in output and no placeholders remain.",
    "test_scalar_replacement_table_cells": "Validates scalar placeholder replacement in table cells. Ensures replacement works correctly in table context.",
    "test_scalar_replacement_header_footer": "Validates scalar placeholder replacement in headers and footers. Ensures replacement works across all document regions.",
    "test_scalar_replacement_determinism": "Validates Invariant 2 (Determinism): same input produces identical output. Runs export twice and compares results.",
    
    # Block markdown
    "test_block_markdown_body_flow_paragraphs": "Validates block markdown with paragraphs in BODY FLOW. Ensures paragraph text and inline formatting (bold, italic) are preserved.",
    "test_block_markdown_body_flow_headings": "Validates block markdown with headings in BODY FLOW. Ensures heading hierarchy (H1, H2, H3) is correctly rendered.",
    "test_block_markdown_table_cell": "Validates block markdown inside table cell. Ensures markdown structures (headings, lists) render correctly within table cells.",
    "test_block_markdown_mixed_content": "Validates block markdown with mixed content types (headings, paragraphs, lists, tables). Ensures all content types work together.",
    "test_block_markdown_determinism": "Validates Invariant 2 (Determinism): same markdown produces identical output. Runs export twice and compares results.",
    
    # Lists
    "test_bullet_list_flat": "Validates flat bullet list rendering. Ensures all list items appear with correct bullet markers.",
    "test_numbered_list_flat": "Validates flat numbered list rendering. Ensures all list items appear with correct numbering.",
    "test_nested_bullet_list": "Validates nested bullet list rendering (up to 3 levels). Ensures nesting depth and indentation are preserved.",
    "test_nested_numbered_list": "Validates nested numbered list rendering (up to 6 levels). Ensures hierarchical numbering (1.1., 1.1.1.) is correct.",
    "test_mixed_bullet_and_numbered_lists": "Validates that bullet and numbered lists are rendered as separate blocks. Ensures list type separation and numbering reset.",
    "test_list_with_continuation_paragraphs": "Validates list items with continuation paragraphs. Ensures continuation text is not lost (text fidelity invariant).",
    "test_list_determinism": "Validates Invariant 2 (Determinism): same list structure produces identical output. Runs export twice and compares results.",
    
    # Tables
    "test_markdown_table_simple": "Validates simple markdown table rendering. Ensures table structure, headers, and cell content are preserved.",
    "test_markdown_table_with_text_before_after": "Validates table with text before and after. Ensures block separation and ordering are correct.",
    "test_table_determinism": "Validates Invariant 2 (Determinism): same table structure produces identical output. Runs export twice and compares results.",
    
    # Textboxes
    "test_scalar_replacement_simulated_textbox": "Validates scalar replacement in textbox-like context. Tests underlying replacement logic for textbox scenarios.",
    
    # Documented limitations
    "test_table_cell_inline_markdown_renders_as_literal": "Documents limitation: inline markdown inside table cells renders as literal text. This is intentional for structure determinism.",
    "test_deep_list_nesting_preserves_all_text": "Documents limitation: deep list nesting clamps visually but preserves all text. Ensures no text loss even with extreme nesting.",
    "test_mixed_list_types_reset_deterministically": "Documents design: mixed bullet ↔ numbered lists reset deterministically. This is intentional behavior, not a bug.",
    
    # Negative cases
    "test_block_placeholder_in_header_not_expanded": "Safety test: validates that block placeholder in header does NOT expand. Block expansion is only allowed in BODY FLOW and BODY TABLE.",
    "test_block_placeholder_in_footer_not_expanded": "Safety test: validates that block placeholder in footer does NOT expand. Block expansion is only allowed in BODY FLOW and BODY TABLE.",
    "test_block_placeholder_mixed_with_text_skipped": "Safety test: validates that block placeholder mixed with other text in paragraph is skipped. Block placeholders must occupy entire paragraph.",
    "test_empty_block_content_removes_placeholder": "Validates that empty block content removes placeholder but doesn't insert content. Ensures correct handling of empty blocks.",
    
    # Golden regression
    "test_golden_regression_comprehensive": "Golden regression test with comprehensive real-world input. Exercises all features: scalar replacement, headings, paragraphs, nested lists, tables, mixed content. Serves as long-term regression anchor.",
}


def extract_test_docstring(test_file_path: Path, test_name: str) -> str:
    """Extract docstring from a test function."""
    try:
        with open(test_file_path, "r", encoding="utf-8") as f:
            tree = ast.parse(f.read(), filename=str(test_file_path))
        
        for node in ast.walk(tree):
            if isinstance(node, ast.FunctionDef) and node.name == test_name:
                if ast.get_docstring(node):
                    # Clean up docstring - take first line or first sentence
                    doc = ast.get_docstring(node)
                    # Remove "Test" prefix if present, keep it concise
                    if doc.startswith("Test "):
                        doc = doc[5:]
                    elif doc.startswith("Assert: "):
                        doc = doc[8:]
                    # Take first sentence
                    first_line = doc.split("\n")[0].strip()
                    return first_line
    except Exception:
        pass
    
    # Fallback to mapping
    return TEST_DESCRIPTIONS.get(test_name, "No description available")


def run_tests():
    """Run pytest and capture output."""
    test_files = [
        "tests/test_scalar_replacement.py",
        "tests/test_block_markdown.py",
        "tests/test_lists.py",
        "tests/test_tables.py",
        "tests/test_textboxes.py",
        "tests/test_documented_limitations.py",
        "tests/test_negative_cases.py",
        "tests/test_golden_regression.py",
    ]
    
    output_file = Path("tests/artifacts/pytest_output.txt")
    output_file.parent.mkdir(parents=True, exist_ok=True)
    
    cmd = [
        sys.executable, "-m", "pytest",
        *test_files,
        "-v", "--tb=line"
    ]
    
    print("Running production test suite...")
    with open(output_file, "w", encoding="utf-8") as f:
        result = subprocess.run(cmd, stdout=f, stderr=subprocess.STDOUT, text=True)
    
    return output_file, result.returncode == 0


def parse_pytest_output(output_file: Path):
    """Parse pytest output and extract test results."""
    results = {}
    
    with open(output_file, "r", encoding="utf-8") as f:
        for line in f:
            if "::" in line and ("PASSED" in line or "FAILED" in line or "SKIPPED" in line):
                # Format: tests/test_file.py::test_function PASSED [  X%]
                parts = line.strip().split("::")
                if len(parts) == 2:
                    test_file = Path(parts[0]).stem
                    test_func_part = parts[1].split()[0]
                    
                    if "PASSED" in line:
                        status = "? PASSED"
                    elif "FAILED" in line:
                        status = "? FAILED"
                    elif "SKIPPED" in line:
                        status = "?? SKIPPED"
                    else:
                        status = "? UNKNOWN"
                    
                    if test_file not in results:
                        results[test_file] = []
                    
                    results[test_file].append({
                        "name": test_func_part,
                        "status": status,
                    })
    
    return results


def generate_summary(results: dict, output_file: Path):
    """Generate markdown summary with descriptions."""
    summary_file = Path("tests/artifacts/TEST_RESULTS_SUMMARY.md")
    
    with open(summary_file, "w", encoding="utf-8") as f:
        f.write("# Test Results Summary\n\n")
        f.write(f"**Generated:** {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
        
        # Calculate statistics
        total = sum(len(tests) for tests in results.values())
        passed = sum(1 for tests in results.values() for t in tests if "PASSED" in t["status"])
        failed = sum(1 for tests in results.values() for t in tests if "FAILED" in t["status"])
        skipped = sum(1 for tests in results.values() for t in tests if "SKIPPED" in t["status"])
        
        f.write("## Summary Statistics\n\n")
        f.write("| Metric | Count |\n")
        f.write("|--------|-------|\n")
        f.write(f"| **Total Tests** | {total} |\n")
        f.write(f"| ? **Passed** | {passed} |\n")
        f.write(f"| ? **Failed** | {failed} |\n")
        f.write(f"| ?? **Skipped** | {skipped} |\n\n")
        
        if failed == 0:
            f.write("## ? All Tests Passed!\n\n")
        else:
            f.write("## ? Some Tests Failed\n\n")
            f.write("Check individual test artifacts in:\n")
            f.write("```\n")
            f.write("tests/artifacts/<test_name>/\n")
            f.write("```\n\n")
        
        f.write("---\n\n")
        f.write("## Test Results by File\n\n")
        
        # Group by test file
        test_file_order = [
            "test_scalar_replacement",
            "test_block_markdown",
            "test_lists",
            "test_tables",
            "test_textboxes",
            "test_documented_limitations",
            "test_negative_cases",
            "test_golden_regression",
        ]
        
        for test_file in test_file_order:
            if test_file not in results:
                continue
            
            f.write(f"### {test_file}.py\n\n")
            f.write("| Test Function | Description | Status |\n")
            f.write("|---------------|-------------|--------|\n")
            
            for test in sorted(results[test_file], key=lambda x: x["name"]):
                # Get description from mapping (preferred - more detailed)
                description = TEST_DESCRIPTIONS.get(test["name"], "No description available")
                
                # If not in mapping, try to extract from actual file
                if description == "No description available":
                    test_file_path = Path(f"tests/{test_file}.py")
                    if test_file_path.exists():
                        extracted_desc = extract_test_docstring(test_file_path, test["name"])
                        if extracted_desc and extracted_desc != "No description available":
                            description = extracted_desc
                
                f.write(f"| `{test['name']}` | {description} | {test['status']} |\n")
            
            f.write("\n")
        
        f.write("---\n\n")
        f.write("## How to Inspect Failures\n\n")
        f.write("1. Check the test artifacts:\n")
        f.write("   ```\n")
        f.write("   tests/artifacts/<test_name>/\n")
        f.write("   ```\n")
        f.write("2. Review `input.json` to see what was tested\n")
        f.write("3. Review `output.docx` to see the actual output\n")
        f.write("4. Review `rendered_markdown.md` to see the markdown\n\n")
        f.write("For detailed pytest output, see: `tests/artifacts/pytest_output.txt`\n")
    
    print(f"? Summary saved to: {summary_file}")
    return summary_file


def main():
    """Main entry point."""
    output_file, success = run_tests()
    results = parse_pytest_output(output_file)
    summary_file = generate_summary(results, output_file)
    
    # Display summary
    print("\n" + "="*60)
    with open(summary_file, "r", encoding="utf-8") as f:
        print(f.read())
    
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
