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
        }
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
