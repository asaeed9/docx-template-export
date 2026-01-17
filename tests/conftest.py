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
