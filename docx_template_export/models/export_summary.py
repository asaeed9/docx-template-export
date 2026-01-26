"""
Export summary model for observability.

This module provides structured tracking of export execution:
- What was successfully replaced
- What was detected but skipped
- What fallback paths were taken
- What warnings were raised

The summary is optional and gated by `enable_export_trace` in WordExportRequest.
When disabled, no summary is collected and no additional logs are emitted.
"""
from dataclasses import dataclass, field
from typing import List, Optional, Dict, Any
from enum import Enum


class Region(Enum):
    """
    Document region where content is located.
    
    This enum identifies the structural region of a Word document where
    a placeholder or content replacement occurs. Regions have different
    capabilities and restrictions:
    
    - BODY: Main document content. Supports both scalar and block replacements.
    - HEADER: Document header. Supports scalar replacements only.
    - FOOTER: Document footer. Supports scalar replacements only.
    
    Block placeholders in HEADER or FOOTER regions are automatically skipped
    for safety (block expansion is restricted to BODY only).
    """
    BODY = "BODY"
    HEADER = "HEADER"
    FOOTER = "FOOTER"


class Container(Enum):
    """
    Container type for content within a document region.
    
    This enum identifies the type of Word document container where
    a placeholder or content replacement occurs:
    
    - FLOW: Normal paragraph flow (body paragraphs, header/footer paragraphs).
        Supports both scalar and block replacements (if in BODY region).
    - TABLE: Table cell content. Supports both scalar and block replacements
        (if in BODY region). Nested tables may trigger fallback to text rendering.
    - TEXTBOX: Textbox content (w:txbxContent elements). Supports scalar
        replacements only. Block placeholders in textboxes are automatically
        skipped for safety.
    """
    FLOW = "FLOW"  # Normal paragraphs
    TABLE = "TABLE"  # Table cells
    TEXTBOX = "TEXTBOX"  # Textboxes (w:txbxContent)


@dataclass
class ScalarReplacement:
    """
    Record of a successful scalar placeholder replacement.
    
    This dataclass tracks each scalar placeholder that was successfully
    replaced during export. Scalar replacements are simple text substitutions
    that can occur in any region (BODY, HEADER, FOOTER) and any container
    (FLOW, TABLE, TEXTBOX).
    
    Attributes:
        token: The placeholder token as it appeared in the template,
            including braces (e.g., "{{document_id}}", "{{title}}").
        region: The document region where replacement occurred (Region enum value).
            One of: "BODY", "HEADER", "FOOTER".
        container: The container type where replacement occurred (Container enum value).
            One of: "FLOW", "TABLE", "TEXTBOX".
        value_preview: First 50 characters of the replacement value, truncated
            if longer. Used for audit trail without storing full content.
    """
    token: str  # e.g., "{{document_id}}"
    region: str  # Region enum value
    container: str  # Container enum value
    value_preview: str  # First 50 chars of replacement value


@dataclass
class BlockReplacement:
    """
    Record of a successful block placeholder replacement.
    
    This dataclass tracks each block placeholder that was successfully
    replaced during export. Block replacements insert structured content
    (headings, paragraphs, lists, tables) parsed from markdown.
    
    Block replacements are restricted to BODY region only. Block placeholders
    in HEADER, FOOTER, or TEXTBOX containers are automatically skipped
    and recorded as SkippedItem instead.
    
    Attributes:
        token: The placeholder token as it appeared in the template,
            including braces (e.g., "{{summary}}", "{{proposal}}").
        region: The document region where replacement occurred (Region enum value).
            Typically "BODY" for successful block replacements (block expansion restricted to BODY).
        container: The container type where replacement occurred (Container enum value).
            One of: "FLOW" (paragraph), "TABLE" (table cell).
        block_types: List of block types that were inserted, in order of appearance.
            Possible values: "heading", "paragraph", "bullet_list", "numbered_list", "table".
            Example: ["heading", "paragraph", "bullet_list", "table"] indicates
            the block content contained a heading, paragraph, bullet list, and table.
    """
    token: str  # e.g., "{{summary}}"
    region: str  # Region enum value
    container: str  # Container enum value
    block_types: List[str]  # Types of blocks inserted (e.g., ["heading", "paragraph", "bullet_list"])


@dataclass
class SkippedItem:
    """
    Record of a placeholder that was detected but skipped during export.
    
    This dataclass tracks placeholders that were found in the template but
    were intentionally skipped due to safety constraints or structural requirements.
    Skipped placeholders remain unchanged in the output document.
    
    Common skip reasons:
    - "block_in_textbox": Block placeholder found in textbox (textboxes are scalar-only)
    - "block_in_header": Block placeholder found in header (headers are scalar-only)
    - "block_in_footer": Block placeholder found in footer (footers are scalar-only)
    - "not_entire_paragraph": Block placeholder mixed with other text in same paragraph
    - "ambiguous_placeholder": Placeholder appears multiple times (safety constraint)
    - "empty_content": Block field value is None or empty string
    
    Attributes:
        token: The placeholder token as it appeared in the template,
            including braces (e.g., "{{summary}}", "{{title}}").
        region: The document region where placeholder was found (Region enum value).
        container: The container type where placeholder was found (Container enum value).
        reason: Human-readable explanation of why the placeholder was skipped.
            This helps users understand why certain replacements did not occur.
    """
    token: str  # Placeholder token
    region: str  # Region enum value
    container: str  # Container enum value
    reason: str  # Why it was skipped (e.g., "block_in_textbox", "block_in_header", "not_entire_paragraph")


@dataclass
class FallbackEvent:
    """
    Record of a fallback mechanism being triggered during export.
    
    This dataclass tracks events where the export engine fell back to
    a simpler or safer rendering path. Fallbacks preserve content fidelity
    while potentially reducing structural complexity.
    
    Common fallback types:
    - "tree_parse_failed": Tree-based renderer failed, fell back to block-based renderer
    - "markdown_parse_failed": Markdown parsing failed, fell back to plain text
    - "nested_table_detected": Nested table detected in table cell, rendered as text grid
    - "table_rendering_failed": Table rendering failed, fell back to text representation
    
    Fallbacks are intentional safety mechanisms that prioritize content preservation
    over structural complexity. They ensure the export completes successfully even
    when encountering edge cases or complex structures.
    
    Attributes:
        event_type: Type of fallback event that occurred. Describes what triggered
            the fallback (e.g., "tree_parse_failed", "nested_table_detected").
        location: Optional description of where the fallback occurred. May include
            the placeholder token or a location description (e.g., "{{summary}}",
            "table cell at row 2, column 1").
        reason: Optional explanation of why the fallback was necessary.
            Provides context for understanding the fallback decision.
    """
    event_type: str  # e.g., "tree_parse_failed", "markdown_parse_failed", "nested_table_detected"
    location: Optional[str] = None  # Where it occurred (token or location description)
    reason: Optional[str] = None  # Why fallback was needed


@dataclass
class WarningEvent:
    """
    Record of a warning raised during export execution.
    
    This dataclass tracks warnings that were emitted during export. Warnings
    indicate non-fatal issues that do not prevent export completion but may
    indicate suboptimal conditions or safety constraint violations.
    
    Warnings are informational and do not affect output correctness. They help
    users understand what happened during export and identify potential template
    or content issues.
    
    Attributes:
        message: The warning message text. Describes what condition triggered
            the warning (e.g., "Block placeholder {{summary}} found in textbox,
            skipping expansion").
        context: Optional dictionary containing additional context about the warning.
            May include keys such as:
            - "token": The placeholder token involved
            - "region": Document region where warning occurred
            - "container": Container type where warning occurred
            - "location": Additional location information
    """
    message: str  # Warning message
    context: Optional[Dict[str, Any]] = None  # Additional context (token, location, etc.)


@dataclass
class ExportSummary:
    """
    Structured summary of export execution for observability and audit trails.
    
    This summary records what happened during export without influencing
    any rendering or decision logic. It is optional and gated by
    `enable_export_trace` in WordExportRequest. When disabled, no summary
    is collected and no additional logs are emitted.
    
    The summary provides an audit trail of:
    - What placeholders were successfully replaced (scalar and block)
    - What placeholders were detected but skipped (and why)
    - What fallback mechanisms were triggered
    - What warnings were raised during export
    
    This information is useful for:
    - Debugging template or content issues
    - Understanding why certain replacements did not occur
    - Auditing export behavior for compliance or quality assurance
    - Identifying patterns in skipped items or fallbacks
    
    Attributes:
        scalar_replacements: List of scalar placeholder replacements that
            were successfully performed. Each entry records the token, location
            (region/container), and a preview of the replacement value.
        block_replacements: List of block placeholder replacements that
            were successfully performed. Each entry records the token, location,
            and the types of blocks that were inserted.
        skipped_items: List of placeholders that were detected but skipped.
            Each entry records the token, location, and reason for skipping.
        fallback_events: List of fallback mechanisms that were triggered.
            Each entry records the event type, location, and reason.
        warnings: List of warnings that were raised during export.
            Each entry records the warning message and optional context.
    
    Example:
        ```python
        summary = ExportSummary()
        # ... export execution populates summary ...
        
        # Access summary data
        print(f"Replaced {len(summary.scalar_replacements)} scalar placeholders")
        print(f"Skipped {len(summary.skipped_items)} placeholders")
        
        # Serialize to dictionary for JSON export
        summary_dict = summary.to_dict()
        ```
    """
    scalar_replacements: List[ScalarReplacement] = field(default_factory=list)
    block_replacements: List[BlockReplacement] = field(default_factory=list)
    skipped_items: List[SkippedItem] = field(default_factory=list)
    fallback_events: List[FallbackEvent] = field(default_factory=list)
    warnings: List[WarningEvent] = field(default_factory=list)
    
    def to_dict(self) -> Dict[str, Any]:
        """
        Convert summary to dictionary for serialization (e.g., JSON export).
        
        This method converts the ExportSummary and all nested dataclasses
        into plain dictionaries suitable for JSON serialization. All enum
        values are converted to their string representations.
        
        Returns:
            Dictionary containing all summary data with the following structure:
            {
                "scalar_replacements": [...],
                "block_replacements": [...],
                "skipped_items": [...],
                "fallback_events": [...],
                "warnings": [...]
            }
            
        The summary is automatically saved to a JSON file when `enable_export_trace=True`
        in the export request. The file path is returned in the export result as
        `export_summary_log_path`.
        """
        return {
            "scalar_replacements": [
                {
                    "token": r.token,
                    "region": r.region,
                    "container": r.container,
                    "value_preview": r.value_preview,
                }
                for r in self.scalar_replacements
            ],
            "block_replacements": [
                {
                    "token": r.token,
                    "region": r.region,
                    "container": r.container,
                    "block_types": r.block_types,
                }
                for r in self.block_replacements
            ],
            "skipped_items": [
                {
                    "token": s.token,
                    "region": s.region,
                    "container": s.container,
                    "reason": s.reason,
                }
                for s in self.skipped_items
            ],
            "fallback_events": [
                {
                    "event_type": f.event_type,
                    "location": f.location,
                    "reason": f.reason,
                }
                for f in self.fallback_events
            ],
            "warnings": [
                {
                    "message": w.message,
                    "context": w.context,
                }
                for w in self.warnings
            ],
        }
