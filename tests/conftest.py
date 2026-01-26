"""
Pytest configuration and fixtures for tests.
"""
from pathlib import Path

import pytest
from docx import Document


@pytest.fixture
def test_output_dir():
    """Create and return test output directory."""
    output_dir = Path(__file__).parent / "test_output"
    output_dir.mkdir(exist_ok=True)
    return output_dir


@pytest.fixture
def test_fixtures_dir():
    """Return test fixtures directory."""
    fixtures_dir = Path(__file__).parent / "fixtures"
    fixtures_dir.mkdir(exist_ok=True)
    return fixtures_dir


@pytest.fixture
def artifacts_dir(test_output_dir):
    """Return artifacts directory for tests."""
    artifacts_base = test_output_dir.parent / "artifacts"
    artifacts_base.mkdir(exist_ok=True)
    return artifacts_base


@pytest.fixture
def simple_template(test_fixtures_dir):
    """Create a simple Word template for testing."""
    template_path = test_fixtures_dir / "simple_template.docx"
    
    # Create a simple template with placeholders
    doc = Document()
    doc.add_heading("Document Template", 0)
    doc.add_paragraph("Document ID: {{document_id}}")
    doc.add_paragraph("Title: {{title}}")
    doc.add_paragraph("Author: {{author}}")
    doc.add_paragraph("Date: {{date}}")
    doc.add_paragraph("")
    doc.add_paragraph("{{introduction}}")
    doc.add_paragraph("")
    doc.add_paragraph("{{body}}")
    doc.add_paragraph("")
    doc.add_paragraph("{{conclusion}}")
    
    template_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(template_path))
    
    return template_path


@pytest.fixture
def comprehensive_template(test_fixtures_dir):
    """Create a comprehensive Word template for testing."""
    template_path = test_fixtures_dir / "comprehensive_template.docx"
    
    doc = Document()
    doc.add_heading("Comprehensive Test Template", 0)
    
    # Scalar fields section
    doc.add_heading("Document Information", 1)
    doc.add_paragraph("ID: {{document_id}}")
    doc.add_paragraph("Title: {{title}}")
    doc.add_paragraph("Author: {{author}}")
    doc.add_paragraph("Version: {{version}}")
    
    # Block fields section
    doc.add_heading("Content Sections", 1)
    doc.add_paragraph("{{introduction}}")
    doc.add_paragraph("{{body}}")
    doc.add_paragraph("{{conclusion}}")
    
    # Additional test placeholders
    doc.add_paragraph("{{custom_section}}")
    
    template_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(template_path))
    
    return template_path


@pytest.fixture(autouse=True)
def validate_docx_outputs(request, test_output_dir):
    """
    Automatically validate DOCX files created during tests.
    
    This fixture runs after each test and validates that any DOCX files
    in test_output_dir that were modified during the test can be opened
    and saved without corruption. This ensures we never generate documents
    that trigger Word repair warnings.
    """
    import time
    from tests.test_artifacts_helpers import validate_docx_integrity
    
    # Record start time to identify files created during this test
    test_start_time = time.time()
    
    yield  # Run test first
    
    # After test completes, validate DOCX files modified during this test
    # We check files modified in the last 10 seconds (should cover test execution)
    current_time = time.time()
    docx_files = list(test_output_dir.glob("*.docx"))
    
    for docx_path in docx_files:
        try:
            # Check if file was modified during test execution (within last 10 seconds)
            file_mtime = docx_path.stat().st_mtime
            if current_time - file_mtime < 10:
                validate_docx_integrity(docx_path)
        except AssertionError as e:
            # Re-raise with test context
            pytest.fail(
                f"Test '{request.node.name}' generated a corrupted DOCX file: {docx_path}\n"
                f"{str(e)}\n"
                f"This document would trigger Word repair warnings. Fix the issue before proceeding."
            )
        except Exception as e:
            # Log but don't fail on file system errors (file might have been deleted)
            pass
