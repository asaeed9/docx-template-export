"""
Pydantic models for Word export functionality.
"""
from typing import Optional, Dict
from pydantic import BaseModel, Field


class WordExportRequest(BaseModel):
    """
    Request payload for exporting markdown content into a Word template.

    This model accepts arbitrary scalar fields and block content fields,
    making it flexible for any use case. The export process is strictly
    non-generative: it does not call any LLMs and does not modify wording.
    It only maps existing content into Word document structures.

    Attributes:
        scalar_fields: Dictionary mapping placeholder names (without braces) to
            replacement values. These are used for simple text replacement in
            the template. Placeholders must appear in the template as {{key}}.
            Values are treated as plain text and inserted directly.
            
            Example: {"title": "My Document"} attempts to replace {{title}} in the template where found.
            
            Validation:
            - Keys must be non-empty strings
            - Values can be None (placeholder will be replaced with empty string)
            - Values are treated as plain text (no markdown processing)
            
        block_fields: Dictionary mapping placeholder names (without braces) to
            markdown or plain text content. These are used for structured content
            replacement. Placeholders must appear in the template as {{key}} and
            must occupy entire paragraphs (not mixed with other text).
            
            Example: {"summary": "# Introduction\\n\\nContent..."} attempts to replace
            {{summary}} with parsed markdown content where structurally safe.
            
            Validation:
            - Keys must be non-empty strings
            - Values can be None (placeholder paragraph will be removed)
            - Values are processed as markdown if markdown_mode=True in export_to_word()
            - Block placeholders must occupy entire paragraphs in the template
            - Block placeholders in headers/footers/textboxes are skipped (scalar-only)
            
        enable_export_trace: If True (default), enables observability features:
            - Export summary collection (scalar/block replacements, skipped items, etc.)
            - Structured logging aligned with summary
            - Automatic log file saving to timestamped directories
            
            If False, observability is disabled with near-zero overhead.
            This flag does NOT affect rendering behavior or output structure.

    Safety Guarantees:
        - Scalar placeholders may appear anywhere (body, headers, footers, table cells, textboxes) where found
        - Block placeholders are restricted to BODY containers only (headers/footers/textboxes excluded)
        - Block placeholders must occupy entire paragraphs (mixed with text = skipped)
        - Ambiguous placeholders (appearing multiple times) are skipped for safety
        - Content is preserved as text when structure cannot be rendered (text-preserving fallback)

    Example:
        ```python
        request = WordExportRequest(
            scalar_fields={
                "document_id": "DOC-12345",
                "title": "My Document",
                "author": "John Doe",
                "date": "2024-01-24"
            },
            block_fields={
                "summary": "# Introduction\\n\\nThis is the summary content...",
                "details": "## Details\\n\\nMore information here...",
                "conclusion": "## Conclusion\\n\\nFinal thoughts..."
            },
            enable_export_trace=True
        )
        ```

    Example JSON structure:
        {
            "scalar_fields": {
                "document_id": "DOC-12345",
                "title": "My Document",
                "author": "John Doe"
            },
            "block_fields": {
                "summary": "# Introduction\\n\\nThis is the summary...",
                "details": "## Details\\n\\nMore information...",
                "conclusion": "## Conclusion\\n\\nFinal thoughts..."
            },
            "enable_export_trace": true
        }
    """

    # Dynamic scalar fields - any key-value pairs for simple placeholder replacement
    scalar_fields: Dict[str, Optional[str]] = Field(
        default_factory=dict,
        description="Dictionary of scalar placeholder replacements. Keys should match template placeholders without braces (e.g., 'title' for {{title}})."
    )

    # Dynamic block fields - any key-value pairs for structured content blocks
    block_fields: Dict[str, Optional[str]] = Field(
        default_factory=dict,
        description="Dictionary of block content replacements. Keys should match template placeholders without braces (e.g., 'summary' for {{summary}}). Values are markdown or plain text."
    )

    # Observability control: enable/disable export trace (summary + structured logging)
    enable_export_trace: bool = Field(
        default=True,
        description="If True (default), collect export summary and emit structured logs. If False, observability is disabled with near-zero overhead. Does not affect rendering behavior."
    )
