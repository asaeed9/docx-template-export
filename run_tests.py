#!/usr/bin/env python3
"""
Simple test runner script.

This script runs the test suite and provides a summary of outputs.
"""
import subprocess
import sys
from pathlib import Path

def main():
    """Run tests and display summary."""
    test_dir = Path(__file__).parent / "tests"
    output_dir = test_dir / "test_output"
    
    print("=" * 70)
    print("Running Deterministic DOCX Export Test Suite")
    print("=" * 70)
    print()
    
    # Run pytest
    result = subprocess.run(
        [sys.executable, "-m", "pytest", str(test_dir), "-v"],
        capture_output=False,
    )
    
    print()
    print("=" * 70)
    if result.returncode == 0:
        print("? All tests passed!")
    else:
        print("? Some tests failed. Check output above for details.")
    print("=" * 70)
    print()
    
    # List generated outputs
    if output_dir.exists():
        docx_files = list(output_dir.glob("*.docx"))
        json_files = list(output_dir.glob("*_input.json"))
        md_files = list(output_dir.glob("*.md"))
        
        print(f"Test outputs saved in: {output_dir}")
        print()
        print(f"Generated files:")
        print(f"  - {len(docx_files)} Word documents (.docx)")
        print(f"  - {len(json_files)} input JSON files")
        print(f"  - {len(md_files)} markdown files")
        print()
        print("You can inspect these files to verify the export functionality.")
        print()
    
    return result.returncode

if __name__ == "__main__":
    sys.exit(main())
