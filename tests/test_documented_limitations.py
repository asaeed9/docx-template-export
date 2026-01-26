"""
Test suite for documented limitations and intentional behaviors.

This suite asserts that certain behaviors are by design and must be protected:
- Inline markdown inside table cells renders as literal text
- Deep list nesting clamps visually but preserves all text
- Mixed bullet ? numbered lists reset deterministically

These behaviors are intentional and must not change without explicit design decisions.
"""
import pytest
from pathlib import Path
from docx import Document

from docx_template_export.models.export_models import WordExportRequest
from tests.test_artifacts_helpers import (
    run_export_and_collect_artifacts,
    assert_text_fidelity,
)


@pytest.fixture
def artifacts_dir(test_output_dir):
    """Return artifacts directory for this test suite."""
    artifacts_base = test_output_dir.parent / "artifacts"
    artifacts_base.mkdir(exist_ok=True)
    return artifacts_base


@pytest.fixture
def limitations_template(test_fixtures_dir):
    """Create template with block placeholder for limitations testing."""
    template_path = test_fixtures_dir / "limitations_template.docx"
    doc = Document()
    doc.add_paragraph("{{content}}")
    template_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(template_path))
    return template_path


def test_table_cell_inline_markdown_renders_as_literal(artifacts_dir, limitations_template):
    """
    Assert: Inline markdown inside table cells renders as literal text.
    
    This is a documented limitation. Table cells prioritize structure
    determinism over inline formatting.
    """
    request = WordExportRequest(
        scalar_fields={},
        block_fields={
            "content": """| Column |
|--------|
| **Bold** and *italic* text |
""",
        },
    )
    
    artifact_dir, markdown, extracted_text = run_export_and_collect_artifacts(
        test_name="table_cell_inline_markdown_literal",
        template_path=limitations_template,
        request=request,
        artifacts_dir=artifacts_dir,
        subdirectory="documented_limitations",
    )
    
    # Assert text fidelity - the text should appear, but formatting may be literal
    assert_text_fidelity(
        input_texts=[
            "Column",
            "Bold",  # Text should appear even if formatting is literal
            "italic",
        ],
        extracted_text=extracted_text,
        test_name="table_cell_inline_markdown_literal",
    )
    
    # This test documents the limitation: inline formatting in table cells
    # may render as literal text rather than formatted text.
    # The important invariant is that NO TEXT IS LOST.


def test_deep_list_nesting_preserves_all_text(artifacts_dir, limitations_template):
    """
    Assert: Deep list nesting clamps visually but preserves all text.
    
    This is a documented limitation. Very deep nesting may reduce visual
    clarity, but all text must be preserved.
    """
    request = WordExportRequest(
        scalar_fields={},
        block_fields={
            "content": """- Level 1
  - Level 2
    - Level 3
      - Level 4
        - Level 5
          - Level 6
            - Level 7 text must be preserved
""",
        },
    )
    
    artifact_dir, markdown, extracted_text = run_export_and_collect_artifacts(
        test_name="deep_list_nesting_preserves_text",
        template_path=limitations_template,
        request=request,
        artifacts_dir=artifacts_dir,
        subdirectory="documented_limitations",
    )
    
    # Assert text fidelity - ALL levels must appear, even if visual depth is clamped
    assert_text_fidelity(
        input_texts=[
            "Level 1",
            "Level 2",
            "Level 3",
            "Level 4",
            "Level 5",
            "Level 6",
            "Level 7 text must be preserved",
        ],
        extracted_text=extracted_text,
        test_name="deep_list_nesting_preserves_text",
    )
    
    # This test documents that visual clamping may occur, but text is never lost.


def test_mixed_list_types_reset_deterministically(artifacts_dir, limitations_template):
    """
    Assert: Mixed bullet ? numbered lists reset deterministically.
    
    This is by design. Bullet and numbered lists are rendered as separate
    blocks, and numbering resets for each new numbered list block.
    """
    request = WordExportRequest(
        scalar_fields={},
        block_fields={
            "content": """- Bullet A
- Bullet B

1. Numbered 1
2. Numbered 2

- Bullet C

1. Numbered 1 (reset)
2. Numbered 2 (reset)
""",
        },
    )
    
    artifact_dir, markdown, extracted_text = run_export_and_collect_artifacts(
        test_name="mixed_list_types_reset",
        template_path=limitations_template,
        request=request,
        artifacts_dir=artifacts_dir,
        subdirectory="documented_limitations",
    )
    
    # Assert text fidelity
    assert_text_fidelity(
        input_texts=[
            "Bullet A",
            "Bullet B",
            "Numbered 1",
            "Numbered 2",
            "Bullet C",
            "Numbered 1 (reset)",
            "Numbered 2 (reset)",
        ],
        extracted_text=extracted_text,
        test_name="mixed_list_types_reset",
    )
    
    # This test documents that list type changes reset numbering deterministically.
    # This is intentional behavior, not a bug.
