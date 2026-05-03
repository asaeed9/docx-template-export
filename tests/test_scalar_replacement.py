"""
Test suite for scalar placeholder replacement across all locations.

Enforces:
- Invariant 1: Text Fidelity - all scalar values appear in output
- Invariant 2: Determinism - same input produces same output
- Invariant 3: Observability - all artifacts written to disk
"""
import pytest
from pathlib import Path
from docx import Document

from docx_template_export.models.export_models import WordExportRequest
from tests.test_artifacts_helpers import (
    run_export_and_collect_artifacts,
    assert_text_fidelity,
    assert_no_placeholders_remain,
    extract_all_text_from_docx,
)


@pytest.fixture
def artifacts_dir(test_output_dir):
    """Return artifacts directory for this test suite."""
    artifacts_base = test_output_dir.parent / "artifacts"
    artifacts_base.mkdir(exist_ok=True)
    return artifacts_base


@pytest.fixture
def scalar_template_body(test_fixtures_dir):
    """Create template with scalar placeholders in BODY FLOW."""
    template_path = test_fixtures_dir / "scalar_template_body.docx"
    doc = Document()
    doc.add_paragraph("Document ID: {{document_id}}")
    doc.add_paragraph("Title: {{title}}")
    doc.add_paragraph("Author: {{author}}")
    doc.add_paragraph("Version: {{version}}")
    template_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(template_path))
    return template_path


@pytest.fixture
def scalar_template_table(test_fixtures_dir):
    """Create template with scalar placeholders in table cells."""
    template_path = test_fixtures_dir / "scalar_template_table.docx"
    doc = Document()
    table = doc.add_table(rows=2, cols=2)
    table.rows[0].cells[0].paragraphs[0].text = "ID"
    table.rows[0].cells[1].paragraphs[0].text = "{{document_id}}"
    table.rows[1].cells[0].paragraphs[0].text = "Title"
    table.rows[1].cells[1].paragraphs[0].text = "{{title}}"
    template_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(template_path))
    return template_path


@pytest.fixture
def scalar_template_header_footer(test_fixtures_dir):
    """Create template with scalar placeholders in header and footer."""
    template_path = test_fixtures_dir / "scalar_template_header_footer.docx"
    doc = Document()
    
    # Header
    header = doc.sections[0].header
    header.paragraphs[0].text = "{{document_id}} - {{title}}"
    
    # Body
    doc.add_paragraph("Body content here")
    
    # Footer
    footer = doc.sections[0].footer
    footer.paragraphs[0].text = "Page {{page_number}}"
    
    template_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(template_path))
    return template_path


def test_scalar_replacement_body_flow(artifacts_dir, scalar_template_body):
    """Test scalar replacement in BODY FLOW paragraphs."""
    request = WordExportRequest(
        scalar_fields={
            "document_id": "DOC-12345",
            "title": "Test Document Title",
            "author": "John Doe",
            "version": "1.0.0",
        },
        block_fields={},
    )
    
    artifact_dir, markdown, extracted_text = run_export_and_collect_artifacts(
        test_name="scalar_replacement_body_flow",
        template_path=scalar_template_body,
        request=request,
        artifacts_dir=artifacts_dir,
        subdirectory="scalar_replacement",
    )
    
    # Assert text fidelity
    assert_text_fidelity(
        input_texts=[
            "DOC-12345",
            "Test Document Title",
            "John Doe",
            "1.0.0",
        ],
        extracted_text=extracted_text,
        test_name="scalar_replacement_body_flow",
    )
    
    # Assert no placeholders remain
    assert_no_placeholders_remain(
        extracted_text=extracted_text,
        scalar_placeholders=["document_id", "title", "author", "version"],
        block_placeholders=[],
        test_name="scalar_replacement_body_flow",
    )
    
    # Assert artifacts exist
    assert (artifact_dir / "input.json").exists()
    assert (artifact_dir / "output.docx").exists()


def test_scalar_replacement_table_cells(artifacts_dir, scalar_template_table):
    """Test scalar replacement in table cells."""
    request = WordExportRequest(
        scalar_fields={
            "document_id": "TBL-67890",
            "title": "Table Test Document",
        },
        block_fields={},
    )
    
    artifact_dir, markdown, extracted_text = run_export_and_collect_artifacts(
        test_name="scalar_replacement_table_cells",
        template_path=scalar_template_table,
        request=request,
        artifacts_dir=artifacts_dir,
        subdirectory="scalar_replacement",
    )
    
    # Assert text fidelity
    assert_text_fidelity(
        input_texts=[
            "TBL-67890",
            "Table Test Document",
        ],
        extracted_text=extracted_text,
        test_name="scalar_replacement_table_cells",
    )
    
    # Assert no placeholders remain
    assert_no_placeholders_remain(
        extracted_text=extracted_text,
        scalar_placeholders=["document_id", "title"],
        block_placeholders=[],
        test_name="scalar_replacement_table_cells",
    )


def test_scalar_replacement_header_footer(artifacts_dir, scalar_template_header_footer):
    """Test scalar replacement in headers and footers."""
    request = WordExportRequest(
        scalar_fields={
            "document_id": "HDR-11111",
            "title": "Header Footer Test",
            "page_number": "1",
        },
        block_fields={},
    )
    
    artifact_dir, markdown, extracted_text = run_export_and_collect_artifacts(
        test_name="scalar_replacement_header_footer",
        template_path=scalar_template_header_footer,
        request=request,
        artifacts_dir=artifacts_dir,
        subdirectory="scalar_replacement",
    )
    
    # Assert text fidelity
    assert_text_fidelity(
        input_texts=[
            "HDR-11111",
            "Header Footer Test",
            "1",
        ],
        extracted_text=extracted_text,
        test_name="scalar_replacement_header_footer",
    )
    
    # Assert no placeholders remain
    assert_no_placeholders_remain(
        extracted_text=extracted_text,
        scalar_placeholders=["document_id", "title", "page_number"],
        block_placeholders=[],
        test_name="scalar_replacement_header_footer",
    )


# ---------------------------------------------------------------------------
# Helpers for nested-table fixture construction
# ---------------------------------------------------------------------------

def _make_nested_table_docx(path: Path, placeholder: str, depth: int) -> None:
    """
    Create a DOCX where `placeholder` lives inside a table nested `depth` levels deep.

    depth=1 → body → table → cell → paragraph(placeholder)
    depth=2 → body → outer table → outer cell → inner table → inner cell → paragraph(placeholder)
    depth=3 → one more level of wrapping

    Each table cell always includes a trailing empty w:p so the structure satisfies
    the OOXML requirement that every w:tc must contain at least one w:p.
    """
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    doc = Document()

    # Innermost content: a paragraph with the placeholder text
    p = OxmlElement('w:p')
    r = OxmlElement('w:r')
    t = OxmlElement('w:t')
    t.text = placeholder
    r.append(t)
    p.append(r)
    content = p  # will be wrapped layer by layer

    for _i in range(depth):
        tbl = OxmlElement('w:tbl')

        tblPr = OxmlElement('w:tblPr')
        tblW = OxmlElement('w:tblW')
        tblW.set(qn('w:w'), '0')
        tblW.set(qn('w:type'), 'auto')
        tblPr.append(tblW)
        tbl.append(tblPr)

        tblGrid = OxmlElement('w:tblGrid')
        gc = OxmlElement('w:gridCol')
        gc.set(qn('w:w'), '5000')
        tblGrid.append(gc)
        tbl.append(tblGrid)

        tr = OxmlElement('w:tr')
        tc = OxmlElement('w:tc')

        tcPr = OxmlElement('w:tcPr')
        tcW = OxmlElement('w:tcW')
        tcW.set(qn('w:w'), '5000')
        tcW.set(qn('w:type'), 'dxa')
        tcPr.append(tcW)
        tc.append(tcPr)

        tc.append(content)
        tc.append(OxmlElement('w:p'))  # required trailing paragraph
        tr.append(tc)
        tbl.append(tr)
        content = tbl  # next level wraps this table

    # Insert the outermost table before sectPr so the body structure is valid
    body = doc._element.body
    sect_pr_list = body.findall(qn('w:sectPr'))
    if sect_pr_list:
        sect_pr_list[0].addprevious(content)
    else:
        body.append(content)

    path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(path))


def _make_two_col_nested_template(path: Path) -> None:
    """
    Create a DOCX with a 2-column outer table where each column contains
    an inner (nested) table:
      left cell  → inner table → {{title}}
      right cell → inner table → {{author}}
    Used to verify that each placeholder is replaced exactly once.
    """
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    doc = Document()

    def _inner_table(placeholder: str) -> "OxmlElement":
        p = OxmlElement('w:p')
        r = OxmlElement('w:r')
        t = OxmlElement('w:t')
        t.text = placeholder
        r.append(t)
        p.append(r)

        tbl = OxmlElement('w:tbl')
        tblPr = OxmlElement('w:tblPr')
        tblW = OxmlElement('w:tblW')
        tblW.set(qn('w:w'), '0')
        tblW.set(qn('w:type'), 'auto')
        tblPr.append(tblW)
        tbl.append(tblPr)
        tblGrid = OxmlElement('w:tblGrid')
        gc = OxmlElement('w:gridCol')
        gc.set(qn('w:w'), '2500')
        tblGrid.append(gc)
        tbl.append(tblGrid)
        tr = OxmlElement('w:tr')
        tc = OxmlElement('w:tc')
        tcPr = OxmlElement('w:tcPr')
        tcW = OxmlElement('w:tcW')
        tcW.set(qn('w:w'), '2500')
        tcW.set(qn('w:type'), 'dxa')
        tcPr.append(tcW)
        tc.append(tcPr)
        tc.append(p)
        tc.append(OxmlElement('w:p'))
        tr.append(tc)
        tbl.append(tr)
        return tbl

    outer_tbl = OxmlElement('w:tbl')
    outer_tblPr = OxmlElement('w:tblPr')
    outer_tblW = OxmlElement('w:tblW')
    outer_tblW.set(qn('w:w'), '0')
    outer_tblW.set(qn('w:type'), 'auto')
    outer_tblPr.append(outer_tblW)
    outer_tbl.append(outer_tblPr)
    outer_grid = OxmlElement('w:tblGrid')
    for _ in range(2):
        gc = OxmlElement('w:gridCol')
        gc.set(qn('w:w'), '2500')
        outer_grid.append(gc)
    outer_tbl.append(outer_grid)

    outer_tr = OxmlElement('w:tr')
    for ph in ['{{title}}', '{{author}}']:
        tc = OxmlElement('w:tc')
        tcPr = OxmlElement('w:tcPr')
        tcW = OxmlElement('w:tcW')
        tcW.set(qn('w:w'), '2500')
        tcW.set(qn('w:type'), 'dxa')
        tcPr.append(tcW)
        tc.append(tcPr)
        tc.append(_inner_table(ph))
        tc.append(OxmlElement('w:p'))
        outer_tr.append(tc)
    outer_tbl.append(outer_tr)

    body = doc._element.body
    sect_pr_list = body.findall(qn('w:sectPr'))
    if sect_pr_list:
        sect_pr_list[0].addprevious(outer_tbl)
    else:
        body.append(outer_tbl)

    path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(path))


# ---------------------------------------------------------------------------
# Tests: nested table scalar replacement (fix for v0.3.2)
# ---------------------------------------------------------------------------

def test_scalar_replacement_in_nested_table(artifacts_dir, test_fixtures_dir):
    """Scalar placeholder in a 2-level nested table must be replaced."""
    template_path = test_fixtures_dir / "nested_table_depth2.docx"
    _make_nested_table_docx(template_path, "{{title}}", depth=2)

    request = WordExportRequest(
        scalar_fields={"title": "Nested Title Value"},
        block_fields={},
    )
    _, _, extracted_text = run_export_and_collect_artifacts(
        test_name="scalar_nested_depth2",
        template_path=template_path,
        request=request,
        artifacts_dir=artifacts_dir,
        subdirectory="scalar_replacement",
    )

    assert_text_fidelity(["Nested Title Value"], extracted_text, "scalar_nested_depth2")
    assert_no_placeholders_remain(extracted_text, ["title"], [], "scalar_nested_depth2")


def test_scalar_replacement_in_doubly_nested_table(artifacts_dir, test_fixtures_dir):
    """Scalar placeholder in a 3-level nested table must be replaced."""
    template_path = test_fixtures_dir / "nested_table_depth3.docx"
    _make_nested_table_docx(template_path, "{{title}}", depth=3)

    request = WordExportRequest(
        scalar_fields={"title": "Triply Nested Value"},
        block_fields={},
    )
    _, _, extracted_text = run_export_and_collect_artifacts(
        test_name="scalar_nested_depth3",
        template_path=template_path,
        request=request,
        artifacts_dir=artifacts_dir,
        subdirectory="scalar_replacement",
    )

    assert_text_fidelity(["Triply Nested Value"], extracted_text, "scalar_nested_depth3")
    assert_no_placeholders_remain(extracted_text, ["title"], [], "scalar_nested_depth3")


def test_scalar_nested_replacement_no_duplication(artifacts_dir, test_fixtures_dir):
    """Each nested placeholder must be replaced exactly once — no duplication."""
    template_path = test_fixtures_dir / "nested_two_col.docx"
    _make_two_col_nested_template(template_path)

    request = WordExportRequest(
        scalar_fields={"title": "UniqueTitle999", "author": "UniqueAuthor888"},
        block_fields={},
    )
    _, _, extracted_text = run_export_and_collect_artifacts(
        test_name="scalar_nested_no_dup",
        template_path=template_path,
        request=request,
        artifacts_dir=artifacts_dir,
        subdirectory="scalar_replacement",
    )

    assert extracted_text.count("UniqueTitle999") == 1, (
        f"Expected 'UniqueTitle999' exactly once, got {extracted_text.count('UniqueTitle999')}"
    )
    assert extracted_text.count("UniqueAuthor888") == 1, (
        f"Expected 'UniqueAuthor888' exactly once, got {extracted_text.count('UniqueAuthor888')}"
    )
    assert_no_placeholders_remain(extracted_text, ["title", "author"], [], "scalar_nested_no_dup")


def test_scalar_flat_table_regression(artifacts_dir, test_fixtures_dir):
    """Flat (non-nested) table scalar replacement must continue to work after v0.3.2 change."""
    template_path = test_fixtures_dir / "flat_table_regression.docx"
    # depth=1 → single table level, same as doc.add_table() behaviour
    _make_nested_table_docx(template_path, "{{document_id}}", depth=1)

    request = WordExportRequest(
        scalar_fields={"document_id": "REG-FLAT-001"},
        block_fields={},
    )
    _, _, extracted_text = run_export_and_collect_artifacts(
        test_name="scalar_flat_regression",
        template_path=template_path,
        request=request,
        artifacts_dir=artifacts_dir,
        subdirectory="scalar_replacement",
    )

    assert_text_fidelity(["REG-FLAT-001"], extracted_text, "scalar_flat_regression")
    assert_no_placeholders_remain(extracted_text, ["document_id"], [], "scalar_flat_regression")


def test_scalar_replacement_intellectus_exec_brief_template(artifacts_dir):
    """
    Integration test using the real Intellectus executive_decision_brief_template_v4.docx.

    Verifies that the 6 header scalar fields — which live inside a 2-level nested
    table structure in that template — are all replaced after the v0.3.2 fix.

    Skipped automatically when the template is not present (e.g. CI without the
    full Intellectus workspace checked out).
    """
    template_path = (
        Path(__file__).parent.parent.parent
        / "frontend"
        / "resources"
        / "docx-templates"
        / "executive_decision_brief_template_v4.docx"
    )
    if not template_path.exists():
        pytest.skip("Intellectus template not found — skipping integration test")

    request = WordExportRequest(
        scalar_fields={
            "title": "Q3 Infrastructure Upgrade",
            "department": "Engineering",
            "presenter": "Jane Smith",
            "date": "2026-05-03",
            "priority": "High",
            "audience": "Executive Committee",
        },
        block_fields={},
    )
    _, _, extracted_text = run_export_and_collect_artifacts(
        test_name="scalar_intellectus_exec_brief",
        template_path=template_path,
        request=request,
        artifacts_dir=artifacts_dir,
        subdirectory="scalar_replacement",
    )

    assert_text_fidelity(
        [
            "Q3 Infrastructure Upgrade",
            "Engineering",
            "Jane Smith",
            "2026-05-03",
            "High",
            "Executive Committee",
        ],
        extracted_text,
        "scalar_intellectus_exec_brief",
    )
    assert_no_placeholders_remain(
        extracted_text,
        ["title", "department", "presenter", "date", "priority", "audience"],
        [],
        "scalar_intellectus_exec_brief",
    )


def test_scalar_replacement_determinism(artifacts_dir, scalar_template_body):
    """Test that same input produces identical output (Invariant 2: Determinism)."""
    request = WordExportRequest(
        scalar_fields={
            "document_id": "DET-99999",
            "title": "Determinism Test",
        },
        block_fields={},
    )
    
    # Run export twice
    artifact_dir1, markdown1, extracted_text1 = run_export_and_collect_artifacts(
        test_name="scalar_replacement_determinism_run1",
        template_path=scalar_template_body,
        request=request,
        artifacts_dir=artifacts_dir,
        subdirectory="scalar_replacement",
    )
    
    artifact_dir2, markdown2, extracted_text2 = run_export_and_collect_artifacts(
        test_name="scalar_replacement_determinism_run2",
        template_path=scalar_template_body,
        request=request,
        artifacts_dir=artifacts_dir,
        subdirectory="scalar_replacement",
    )
    
    # Assert markdown is identical
    assert markdown1 == markdown2, "Markdown output must be identical for same input"
    
    # Assert extracted text is identical (normalized)
    normalized1 = " ".join(extracted_text1.split())
    normalized2 = " ".join(extracted_text2.split())
    assert normalized1 == normalized2, "Extracted text must be identical for same input"
